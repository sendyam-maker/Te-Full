VERSION 5.00
Begin VB.Form frm050106_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "國內外案件資料維護"
   ClientHeight    =   3600
   ClientLeft      =   345
   ClientTop       =   1650
   ClientWidth     =   5985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5985
   Begin VB.OptionButton optChoose 
      Caption         =   "多筆查詢條件"
      CausesValidation=   0   'False
      Height          =   252
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   600
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "單筆維護"
      CausesValidation=   0   'False
      Height          =   252
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   1950
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4164
      TabIndex        =   14
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4992
      TabIndex        =   15
      Top             =   70
      Width           =   800
   End
   Begin VB.Frame fraChoose 
      Height          =   975
      Index           =   1
      Left            =   180
      TabIndex        =   20
      Top             =   840
      Width           =   5652
      Begin VB.TextBox txtCode 
         Height          =   270
         Index           =   10
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   12
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
         Index           =   11
         Left            =   2760
         MaxLength       =   7
         TabIndex        =   13
         Top             =   600
         Width           =   972
      End
      Begin VB.Label Label3 
         Caption         =   "(民國年月日)"
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   24
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblEnginer 
         Height          =   252
         Left            =   2520
         TabIndex        =   23
         Top             =   240
         Width           =   2172
      End
      Begin VB.Label Label3 
         Caption         =   "國內案發文日："
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "國外案工程師："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line1 
         X1              =   2520
         X2              =   2640
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.Frame fraChoose 
      Enabled         =   0   'False
      Height          =   1332
      Index           =   0
      Left            =   180
      TabIndex        =   16
      Top             =   2160
      Width           =   5652
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
         Left            =   2664
         MaxLength       =   2
         TabIndex        =   5
         Top             =   240
         Width           =   372
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
         Index           =   1
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   3
         Top             =   240
         Width           =   852
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
         Index           =   6
         Left            =   2400
         MaxLength       =   1
         TabIndex        =   8
         Top             =   600
         Width           =   252
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
         Index           =   4
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   6
         Top             =   600
         Width           =   492
      End
      Begin VB.TextBox txtCode 
         Height          =   270
         Index           =   8
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   10
         Top             =   960
         Width           =   372
      End
      Begin VB.Label Label1 
         Caption         =   "國外案號："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "國內案號："
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1092
      End
      Begin VB.Label Label2 
         Caption         =   "功能代號：           (1.新增  2.修改  4.刪除  5.查詢 )"
         Height          =   252
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   3972
      End
   End
End
Attribute VB_Name = "frm050106_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/20 改成Form2.0 (無)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
'0從Menu來,1從frm050101_2來
Public intWhereToGo As Integer

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
                  'If obj003.ChkExist(strCode(), 0) Then
                  If Cls003ChkExist(strCode(), 0) Then
                     If MsgBox("是否要刪除此筆資料 ?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
                        'Modify by Morgan 2005/3/10
'                        If obj003.DeleteCaseRelation(strCode(), 0) Then
'                           For i = 0 To 8
'                              txtCode(i) = ""
'                           Next
'                           txtCode(0).SetFocus
'                        End If
                        cnnConnection.BeginTrans
                        'edit by nickc 2007/02/05 不用 dll 了
                        'If obj003.DeleteCaseRelation(strCode(), 0) Then
                        If Cls003DeleteCaseRelation(strCode(), 0) Then
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
                  
                  'Add by Morgan 2004/1/30
                  '控制國外案之申請國家不可為台灣
                  If CheckIsNotTaiwain() = False Then
                    txtCode(0).SetFocus
                    txtCode_GotFocus 0
                    Screen.MousePointer = vbDefault
                    Exit Sub
                  End If
                  'Add End ------
                  'Modify by Morgan 2004/1/30
                  '改呼叫區域函式
                  'If obj003.ChkCaseMap(strCode, 0) Then
                  If ChkCaseMap(strCode, 0) Then
                     GoTo A0
                  Else
                     txtCode(0).SetFocus
                  End If
               Else
A0:               frm050106_2.intWhereToGo = 0
                  frm050106_2.strCode1 = txtCode(0)
                  frm050106_2.strCode2 = txtCode(1)
                  frm050106_2.strCode3 = txtCode(2)
                  frm050106_2.strCode4 = txtCode(3)
                  frm050106_2.strCode5 = txtCode(4)
                  frm050106_2.strCode6 = txtCode(5)
                  frm050106_2.strCode7 = txtCode(6)
                  frm050106_2.strCode8 = txtCode(7)
                  frm050106_2.intChoose = Val(txtCode(8))
                  frm050106_2.Show
                  Me.Hide
               End If
            Else
               frm050106_3.lblEnginer = txtCode(9)
               frm050106_3.lblEnginerName = lblEnginer
               frm050106_3.lblDate(0) = ChangeTStringToTDateString(txtCode(10))
               frm050106_3.lblDate(1) = ChangeTStringToTDateString(txtCode(11))
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
    
   ChkCaseMap = False
   strCode(2) = IIf(strCode(2) = "", "0", strCode(2))
   strCode(3) = IIf(strCode(3) = "", "00", strCode(3))
   strCode(6) = IIf(strCode(6) = "", "0", strCode(6))
   strCode(7) = IIf(strCode(7) = "", "00", strCode(7))
   
   '檢查是否國內外案件關聯資料已存在
   strExc(0) = "select count(*) from casemap where cm01=" + CNULL(strCode(0)) + " and cm02=" + CNULL(strCode(1)) + _
      " and cm03=" + CNULL(strCode(2)) + " and cm04=" + CNULL(strCode(3)) + " and cm05=" + CNULL(strCode(4)) + _
      " and cm06=" + CNULL(strCode(5)) + " and cm07=" + CNULL(strCode(6)) + " and cm08=" + CNULL(strCode(7)) + _
      " and cm10='" & iSitu & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.Fields(0) > 0 Then
         ShowMsg "國內外案件關聯資料已存在，請重新輸入 !"
      Else
         '檢查是否已取消收文
         'Modify by Morgan 2006/4/14 案件性質改用常數控制
         strExc(0) = "SELECT 1 C1 FROM CASEPROGRESS WHERE CP01='" & strCode(0) & "' AND CP02='" & strCode(1) & "'" & _
            " AND CP03='" & strCode(2) & "' AND CP04='" & strCode(3) & "'" & _
            " AND CP10 in (" & CaseMapOut & ") AND CP57 IS NOT NULL"
         strExc(0) = strExc(0) & " UNION ALL" & _
            " SELECT 2 C2 FROM CASEPROGRESS WHERE CP01='" & strCode(4) & "' AND CP02='" & strCode(5) & "'" & _
            " AND CP03='" & strCode(6) & "' AND CP04='" & strCode(7) & "'" & _
            " AND CP10 in (" & CaseMapIn & ") AND CP57 IS NOT NULL"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0).Value = "1" Then
                ShowMsg strCode(0) & strCode(1) & strCode(2) & strCode(3) & " 已有取消收文日，請重新輸入 !"
            Else
                ShowMsg strCode(4) & strCode(5) & strCode(6) & strCode(7) & " 已有取消收文日，請重新輸入 !"
            End If
         Else
            'Add by Morgan 2006/7/10 控制不可輸多國案
            strExc(0) = "select * from caserelation where CR01='" & strCode(0) & "' AND CR02='" & strCode(1) & "'" & _
            " AND CR03='" & strCode(2) & "' AND CR04='" & strCode(3) & "' and rownum<2"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               ShowMsg "本國外案為多國案，請到多國案卷號關係建立輸入!"
            Else
            'end 2006/7/10
               ChkCaseMap = True
            End If
         End If
      End If
    End If
    
End Function
'Add by Morgan 2004/1/30
'檢查國外案號申請國家是否為台灣
Private Function CheckIsNotTaiwain() As Boolean
    Dim strSql As String, rstQuery As New ADODB.Recordset
On Error GoTo ErrHnd
    strSql = "Select PA09 From Patent Where PA01='" & txtCode(0) & "' AND PA02='" & txtCode(1) & "' AND PA03='" & txtCode(2) & "' AND PA04='" & txtCode(3) & "'"
    rstQuery.CursorLocation = adUseClient
    rstQuery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rstQuery.RecordCount > 0 Then
        If "" & rstQuery.Fields(0).Value <> "000" Then
            CheckIsNotTaiwain = True
        Else
            MsgBox "國外案之申請國家不可為台灣！", vbCritical, "警告"
        End If
    Else
        MsgBox "無法讀取國外案之申請國家！", vbCritical, "警告"
    End If
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
End Sub
Private Sub Form_Load()
Me.Height = 2280 'Add by Morgan 2006/9/20 隱藏維護功能
MoveFormToCenter Me
If intWhereToGo = 0 Then
   cmdOK(1).Caption = "結束"
   cmdOK(1).Cancel = True
End If
'txtCode(10) = GetTodayDate
'txtCode(11) = GetTodayDate
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
    FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Select Case intWhereToGo
      Case 1
         frm050101_2.Show
      ' 91.09.11 modify by louis
      Case 2
         frm010012_05.Show
   End Select
'Add By Cheng 2002/07/18
Set frm050106_1 = Nothing
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

Select Case intIndex
             Case 0
                'Modify By Cheng 2002/12/09
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
                'Modify By Cheng 2002/12/09
                If txtCode(intIndex).Text <> "" Then
                    'Modify by Morgan 2005/12/6 加國內案可輸FCP
                    'If txtCode(intIndex) = "P" Or txtCode(intIndex) = "CFP" Then
                    If txtCode(intIndex) = "P" Or txtCode(intIndex) = "CFP" Or txtCode(intIndex) = "FCP" Then
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
