VERSION 5.00
Begin VB.Form frm02010604_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "分割案件關係維護"
   ClientHeight    =   3690
   ClientLeft      =   345
   ClientTop       =   1650
   ClientWidth     =   5985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5985
   Begin VB.OptionButton optChoose 
      Caption         =   "多筆查詢條件"
      CausesValidation=   0   'False
      Height          =   252
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   2250
      Width           =   1455
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "單筆維護"
      CausesValidation=   0   'False
      Height          =   252
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   744
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4164
      TabIndex        =   11
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4992
      TabIndex        =   12
      Top             =   70
      Width           =   800
   End
   Begin VB.Frame fraChoose 
      Enabled         =   0   'False
      Height          =   975
      Index           =   1
      Left            =   180
      TabIndex        =   16
      Top             =   2490
      Width           =   5652
      Begin VB.TextBox txtCode 
         Height          =   270
         Index           =   11
         Left            =   2850
         MaxLength       =   9
         TabIndex        =   10
         Top             =   600
         Width           =   1155
      End
      Begin VB.TextBox txtCode 
         Height          =   270
         Index           =   10
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   9
         Top             =   600
         Width           =   1155
      End
      Begin VB.TextBox txtCode 
         Height          =   270
         Index           =   9
         Left            =   1440
         TabIndex        =   8
         Top             =   240
         Width           =   2565
      End
      Begin VB.Label Label3 
         Caption         =   "申請人編號："
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "系統類別："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line1 
         X1              =   2670
         X2              =   2790
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.Frame fraChoose 
      Height          =   1125
      Index           =   0
      Left            =   180
      TabIndex        =   13
      Top             =   960
      Width           =   5652
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
         Height          =   264
         Index           =   12
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   252
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
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   3
         Left            =   2820
         MaxLength       =   2
         TabIndex        =   6
         Top             =   240
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   2
         Left            =   2550
         MaxLength       =   1
         TabIndex        =   5
         Top             =   240
         Width           =   252
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   1
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtCode 
         Height          =   270
         Index           =   8
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   7
         Top             =   690
         Width           =   372
      End
      Begin VB.Label Label1 
         Caption         =   "分割案號："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "功能代號：           (1.新增  2.修改  4.刪除  5.查詢 )"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   690
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frm02010604_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/11/30 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/23 日期欄已修改
Option Explicit

Public intWhereToGo As Integer  '0從Menu來,1從frm02010604_2來
Dim m_strSystemKindForUser As String '記錄使用者可使用的系統類別
Dim m_blnFirstShow As Boolean


Private Sub cmdOK_Click(Index As Integer)
Dim i As Integer, varSaveCursor, strCode(7) As String
    Select Case Index
    Case 0 '確定
        varSaveCursor = Screen.MousePointer
        Screen.MousePointer = vbHourglass
        For i = 0 To 13
            If i <> 4 And i <> 5 And i <> 6 And i <> 7 And i <> 13 Then
                If txtCode(i).Enabled Then
                    If CheckKeyIn(i) = False Then
                        '本所案號錯誤時,將Cursor跳回系統別欄位
                        If i = 3 Or i = 7 Then i = i - 3
                        txtCode(i).SetFocus
                        txtCode_GotFocus i
                        Exit For
                    End If
                End If
            End If
        Next i
        If i = 14 Then
            If OptChoose(0).Value Then '單筆維護
                If txtCode(2) = "" Then txtCode(2) = "0"
                If txtCode(3) = "" Then txtCode(3) = "00"
                For i = 0 To 3
                    strCode(i) = txtCode(i)
                Next
                If Me.txtCode(0).Text = "TF" Then
                    strCode(1) = strCode(1) & Me.txtCode(12).Text
                End If
                If Me.txtCode(8).Text = "1" Then
                    If ChkDataRepeat(Me.txtCode(0).Text, Me.txtCode(1).Text & Me.txtCode(12).Text, Me.txtCode(2).Text, Me.txtCode(3).Text) = True Then
                        MsgBox "分割案資料重覆, 不可新增!!!", vbExclamation + vbOKOnly
                        Screen.MousePointer = varSaveCursor
                        Exit Sub
                    End If
                Else
                    If ChkDataRepeat(Me.txtCode(0).Text, Me.txtCode(1).Text & Me.txtCode(12).Text, Me.txtCode(2).Text, Me.txtCode(3).Text) = False Then
                        MsgBox "查無此分割案資料, 無法執行" & IIf(Me.txtCode(8).Text = "2", "修改", IIf(Me.txtCode(8).Text = "4", "刪除", "查詢")) & "功能!!!", vbExclamation + vbOKOnly
                        Screen.MousePointer = varSaveCursor
                        Exit Sub
                    End If
                End If
                frm02010604_2.intWhereToGo = 0
                frm02010604_2.strCode1 = txtCode(0)
                frm02010604_2.strCode2 = txtCode(1) & IIf(Me.txtCode(0).Text = "TF", Me.txtCode(12).Text, "")
                frm02010604_2.strCode3 = txtCode(2)
                frm02010604_2.strCode4 = txtCode(3)
                frm02010604_2.intChoose = Val(txtCode(8))
                frm02010604_2.Show
                Me.Hide
            Else '多筆查詢
                If Me.txtCode(9).Text = "" Then
                    MsgBox "請輸入系統類別!!!", vbExclamation + vbOKOnly
                    Me.txtCode(9).SetFocus
                ElseIf Me.txtCode(10).Text = "" Then
                    MsgBox "請輸入申請人區間起號!!!", vbExclamation + vbOKOnly
                    Me.txtCode(10).SetFocus
                ElseIf Me.txtCode(11).Text = "" Then
                    MsgBox "請輸入申請人區間迄號!!!", vbExclamation + vbOKOnly
                    Me.txtCode(11).SetFocus
                Else
                    frm02010604_3.intWhereToGo = 0
                    frm02010604_3.lblEnginer = txtCode(9).Text
                    frm02010604_3.lblDate(0) = txtCode(10).Text
                    frm02010604_3.lblDate(1) = txtCode(11).Text
                    Me.Hide
                End If
            End If
        End If
        Screen.MousePointer = varSaveCursor
    Case 1 '結束
        Unload Me
    End Select
End Sub

Private Sub Form_Activate()
If m_blnFirstShow = True Then
    If OptChoose(0) Then
       txtCode(0).SetFocus
       txtCode(8) = "1"
    Else
       txtCode(9).SetFocus
    End If
    m_blnFirstShow = False
End If
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
If intWhereToGo = 0 Then
    cmdOK(1).Caption = "結束"
    cmdOK(1).Cancel = True
End If
'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件(P)，但非此類案件時外專程序人員不可操作。
 FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
If FMP2open = True And (UCase(App.EXEName) = "PATPRO" Or UCase(App.EXEName) = "TEPATPRO") Then
    m_strSystemKindForUser = "P,PS,"
Else
    m_strSystemKindForUser = GetSystemKindByNick
 End If
m_blnFirstShow = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Select Case intWhereToGo
    Case 1
        frm02010604_2.Show
    End Select
    Set frm02010604_1 = Nothing
End Sub

Private Sub optChoose_Click(Index As Integer)
fraChoose(Index).Enabled = True
fraChoose((Index + 1) Mod 2).Enabled = False
If Index = 0 Then
    txtCode(0).SetFocus
Else
    txtCode(9).SetFocus
    If Me.txtCode(9).Text = "" Then Me.txtCode(9).Text = m_strSystemKindForUser
End If
End Sub

Private Sub txtCode_Change(Index As Integer)
Select Case Index
Case 0 '分割案系統類別
    If Me.txtCode(0).Text = "TF" Then
        Me.txtCode(1).MaxLength = 5
        Me.txtCode(12).Visible = True
        Me.txtCode(12).Enabled = True
        Me.txtCode(12).Text = ""
    Else
        Me.txtCode(1).MaxLength = 6
        Me.txtCode(12).Visible = False
        Me.txtCode(12).Enabled = False
        Me.txtCode(12).Text = ""
    End If
Case 4 '母案系統類別
    If Me.txtCode(4).Text = "TF" Then
        Me.txtCode(5).MaxLength = 5
        Me.txtCode(13).Visible = True
        Me.txtCode(13).Enabled = True
        Me.txtCode(13).Text = ""
    Else
        Me.txtCode(5).MaxLength = 6
        Me.txtCode(13).Visible = False
        Me.txtCode(13).Enabled = False
        Me.txtCode(13).Text = ""
    End If
End Select
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
txtCode(Index).SelStart = 0
txtCode(Index).SelLength = Len(txtCode(Index))
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
Select Case Index
Case 8 '功能代號
    If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 52 And KeyAscii <> 53 Then
        KeyAscii = 0
    End If
End Select
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
    Case 10
        If Me.txtCode(Index).Text <> "" Then Me.txtCode(Index).Text = Left(Me.txtCode(Index).Text & "00000000", 9)
        Me.txtCode(11).Text = Me.txtCode(10).Text
    Case 11
        If Me.txtCode(Index).Text <> "" Then Me.txtCode(Index).Text = Left(Me.txtCode(Index).Text & "00000000", 9)
    End Select
End Sub

Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 3, 7
      Case Else:
         If CheckKeyIn(Index) = False Then
            '本所案號錯誤時,讓Cursor繼續往下跳
            If Index <> 3 And Index <> 7 Then
               Cancel = True
               txtCode_GotFocus Index
            End If
         End If
   End Select
End Sub

Private Function CheckKeyIn(intIndex As Integer) As Boolean
Dim intCaseKind As Integer, intWhere As Integer, strTemp As String
Dim arrSystemKind, arrSystemKind1
Dim ii As Integer, jj As Integer
Dim blnNoRight As Boolean
Dim strNoRightSK As String
Dim strCode(7) As String

Select Case intIndex
             Case 0, 4 '系統類別
                If Me.txtCode(intIndex).Text <> "" Then
                    If m_strSystemKindForUser <> "" Then
                        blnNoRight = True
                        arrSystemKind = Split(m_strSystemKindForUser, ",")
                        For ii = LBound(arrSystemKind) To UBound(arrSystemKind)
                            If Me.txtCode(intIndex).Text = arrSystemKind(ii) Then
                                blnNoRight = False
                                Exit For
                            End If
                        Next ii
                        If blnNoRight = True Then
                            MsgBox "您無權使用 " & Me.txtCode(intIndex).Text & " 系統類別!!!", vbExclamation + vbOKOnly
                        Else
                            CheckKeyIn = True
                        End If
                    Else
                        MsgBox "您無權使用 " & Me.txtCode(intIndex).Text & " 系統類別!!!", vbExclamation + vbOKOnly
                    End If
                Else
                    CheckKeyIn = True
                End If
             Case 3
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.CheckCaseCodeIsExist(txtCode(intIndex - 3), txtCode(intIndex - 2) & Me.txtCode(12).Text, _
                             IIf(txtCode(intIndex - 1) = "", "0", txtCode(intIndex - 1)), IIf(txtCode(intIndex) = "", "00", txtCode(intIndex))) Then
                  If FMP2open = False Then
                        If ClsPDCheckCaseCodeIsExist(txtCode(intIndex - 3), txtCode(intIndex - 2) & Me.txtCode(12).Text, _
                             IIf(txtCode(intIndex - 1) = "", "0", txtCode(intIndex - 1)), IIf(txtCode(intIndex) = "", "00", txtCode(intIndex))) Then
                            CheckKeyIn = True
                        End If
                  Else
                      'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
                      If PUB_FMPtoCheck(0, 1, Pub_strUserST05, txtCode(intIndex - 3), txtCode(intIndex - 2), _
                        IIf(txtCode(intIndex - 1) = "", "0", txtCode(intIndex - 1)), IIf(txtCode(intIndex) = "", "00", txtCode(intIndex))) = True Then
                         CheckKeyIn = True
                      End If
                  End If
             Case 8
                        If Val(txtCode(intIndex)) = 1 Or Val(txtCode(intIndex)) = 2 Or Val(txtCode(intIndex)) = 4 Or Val(txtCode(intIndex)) = 5 Then
                           CheckKeyIn = True
                        Else
                           ShowMsg MsgText(9198)
                        End If
             Case 9 '系統類別
                        If txtCode(intIndex) = "" Then
                           CheckKeyIn = True
                        Else
                            If m_strSystemKindForUser <> "" Then
                                strNoRightSK = ""
                                arrSystemKind = Split(m_strSystemKindForUser, ",")
                                arrSystemKind1 = Split(Me.txtCode(intIndex).Text, ",")
                                For jj = LBound(arrSystemKind1) To UBound(arrSystemKind1)
                                    blnNoRight = True
                                    For ii = LBound(arrSystemKind) To UBound(arrSystemKind)
                                        If arrSystemKind1(jj) = arrSystemKind(ii) Then
                                            blnNoRight = False
                                            Exit For
                                        End If
                                    Next ii
                                    If blnNoRight = True Then
                                        strNoRightSK = strNoRightSK & arrSystemKind1(jj) & ","
                                    End If
                                Next jj
                                If strNoRightSK <> "" Then
                                    MsgBox "您無權使用 " & Left(strNoRightSK, Len(strNoRightSK) - 1) & " 系統類別!!!", vbExclamation + vbOKOnly
                                Else
                                    CheckKeyIn = True
                                End If
                            Else
                                MsgBox "您無權使用 " & Me.txtCode(intIndex).Text & " 系統類別!!!", vbExclamation + vbOKOnly
                            End If
                        End If
             Case 10, 11 '申請人編號
                        If txtCode(intIndex) = "" Then
                            CheckKeyIn = True
                        Else
                            CheckKeyIn = True
                            If intIndex = 11 Then
                                If Me.txtCode(10).Text <> "" And Me.txtCode(11).Text <> "" Then
                                    If Left(txtCode(10), 6) <> Left(txtCode(11), 6) Then
                                        MsgBox "申請人編號前六碼必須相同!!!", vbExclamation + vbOKOnly
                                        CheckKeyIn = False
                                    End If
                                End If
                            End If
                        End If
             Case Else
                        CheckKeyIn = True
End Select
End Function

Private Function ChkExist(ByRef strCode() As String) As Boolean
Dim strSql As String, rsRecordset As New ADODB.Recordset
    strCode(2) = IIf(strCode(2) = "", "0", strCode(2))
    strCode(3) = IIf(strCode(3) = "", "00", strCode(3))
    strCode(6) = IIf(strCode(6) = "", "0", strCode(6))
    strCode(7) = IIf(strCode(7) = "", "00", strCode(7))
    strSql = "Select Count(*) From DivisionCase Where DC01=" + CNULL(strCode(0)) + " And DC02=" + CNULL(strCode(1)) + _
                    " And DC03=" + CNULL(strCode(2)) + " And DC04=" + CNULL(strCode(3)) + " And DC05=" + CNULL(strCode(4)) + _
                    " And DC06=" + CNULL(strCode(5)) + " And DC07=" + CNULL(strCode(6)) + " And DC08=" + CNULL(strCode(7))
    rsRecordset.CursorLocation = adUseClient
    rsRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
    If Val("" & rsRecordset.Fields(0)) = 0 Then
        ShowMsg "分割案件關係資料不存在，請重新輸入 !"
        ChkExist = False
    Else
        ChkExist = True
    End If
    If rsRecordset.State <> adStateClosed Then rsRecordset.Close
    Set rsRecordset = Nothing
End Function

Private Function DeleteDivisionCase(ByRef strCode() As String) As Boolean
Dim strSql As String, i As Integer

    On Error GoTo ErrHand
    strCode(2) = IIf(strCode(2) = "", "0", strCode(2))
    strCode(3) = IIf(strCode(3) = "", "00", strCode(3))
    strCode(6) = IIf(strCode(6) = "", "0", strCode(6))
    strCode(7) = IIf(strCode(7) = "", "00", strCode(7))
    strSql = "Delete From DivisionCase Where DC01=" + CNULL(strCode(0)) + " And DC02=" + CNULL(strCode(1)) + " And DC03=" + CNULL(strCode(2)) + " And DC04=" + CNULL(strCode(3)) + " And DC05=" + CNULL(strCode(4)) + " And DC06=" + CNULL(strCode(5)) + " And DC07=" + CNULL(strCode(6)) + " And DC08=" + CNULL(strCode(7))
    cnnConnection.Execute strSql, i
    If i = 0 Then ShowMsg MsgText(1007)
    DeleteDivisionCase = True
    Exit Function
ErrHand:
    ShowMsg MsgText(9018)
End Function

Private Function ChkDataRepeat(strDC01 As String, strDC02 As String, strDC03 As String, strDC04 As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

StrSQLa = "Select Count(*) From DivisionCase Where DC01='" & strDC01 & "' And DC02='" & strDC02 & "' And DC03='" & strDC03 & "' And DC04='" & strDC04 & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If Val("" & rsA.Fields(0).Value) > 0 Then
    ChkDataRepeat = True
Else
    ChkDataRepeat = False
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function
