VERSION 5.00
Begin VB.Form frm040109_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "一案兩申請案件資料維護"
   ClientHeight    =   4335
   ClientLeft      =   1020
   ClientTop       =   2430
   ClientWidth     =   4620
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4620
   Begin VB.Frame fraChoose 
      Height          =   1332
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   960
      Width           =   4335
      Begin VB.TextBox txtCode 
         Height          =   270
         Index           =   8
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   8
         Top             =   960
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   4
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   4
         Top             =   600
         Width           =   492
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   7
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   7
         Top             =   600
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   6
         Left            =   3000
         MaxLength       =   1
         TabIndex        =   6
         Top             =   600
         Width           =   252
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   5
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   5
         Top             =   600
         Width           =   852
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   1
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   1
         Top             =   240
         Width           =   852
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   2
         Left            =   3000
         MaxLength       =   1
         TabIndex        =   2
         Top             =   240
         Width           =   252
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   3
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   3
         Top             =   240
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   0
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   0
         Top             =   240
         Width           =   492
      End
      Begin VB.Label Label2 
         Caption         =   "功能代號：           (1.新增  2.修改  4.刪除  5.查詢 )"
         Height          =   252
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   3972
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請案二："
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "申請案一："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraChoose 
      Enabled         =   0   'False
      Height          =   1365
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   2760
      Width           =   4335
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   12
         Left            =   1530
         MaxLength       =   9
         TabIndex        =   9
         Top             =   270
         Width           =   975
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   13
         Left            =   2850
         MaxLength       =   9
         TabIndex        =   10
         Top             =   285
         Width           =   975
      End
      Begin VB.TextBox txtCode 
         Height          =   270
         Index           =   11
         Left            =   2850
         MaxLength       =   7
         TabIndex        =   12
         Top             =   750
         Width           =   972
      End
      Begin VB.TextBox txtCode 
         Height          =   270
         Index           =   10
         Left            =   1530
         MaxLength       =   7
         TabIndex        =   11
         Top             =   750
         Width           =   972
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   2610
         X2              =   2730
         Y1              =   390
         Y2              =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "客戶代碼："
         Height          =   180
         Index           =   0
         Left            =   225
         TabIndex        =   24
         Top             =   360
         Width           =   900
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2610
         X2              =   2730
         Y1              =   870
         Y2              =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "收文日："
         Height          =   180
         Index           =   1
         Left            =   270
         TabIndex        =   19
         Top             =   810
         Width           =   720
      End
      Begin VB.Label lblEnginer 
         Height          =   255
         Left            =   2760
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3672
      TabIndex        =   16
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2844
      TabIndex        =   15
      Top             =   70
      Width           =   800
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "單筆維護"
      CausesValidation=   0   'False
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "多筆查詢條件"
      CausesValidation=   0   'False
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   1455
   End
End
Attribute VB_Name = "frm040109_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/21 改成Form2.0 (無)
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Public m_CM10 As String 'Added by Morgan 2015/9/10
'Added by Morgan 2012/2/9
Public m_bInsert As Boolean
Public m_bUpdate As Boolean
Public m_bDelete As Boolean
Public m_bQuery As Boolean

Public frmParent As Form

Public Function ChkExist() As Boolean
            
   Dim i As Integer, strCode(7) As String
   
   For i = 0 To 7
      strCode(i) = txtCode(i)
   Next
   
On Error GoTo ErrHnd
   'Modified by Morgan 2015/9/10 +擬制喪失新穎性
   strSql = "select * from casemap where ( ( cm01=" & CNULL(strCode(0)) & " and cm02=" + CNULL(strCode(1)) & _
      " and cm03=" & CNULL(strCode(2)) & " and cm04=" & CNULL(strCode(3)) & ") or ( cm05=" + CNULL(strCode(0)) & _
      " and cm06=" & CNULL(strCode(1)) & " and cm07=" & CNULL(strCode(2)) & " and cm08=" & CNULL(strCode(3)) & ") )" & _
      " and cm10='" & IIf(m_CM10 <> "", m_CM10, "3") & "'"
      
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount > 0 Then
      For i = 0 To 7
         txtCode(i) = "" & adoRecordset.Fields(i)
      Next
      ChkExist = True
   End If
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
   CheckOC
            
End Function
Private Sub Process()

   Dim i As Integer, strCode(7) As String
   Dim oText As TextBox, bolCheckOk As Boolean
   Dim bolExists As Boolean 'Added by Morgan 2017/1/20
   
   bolCheckOk = True
   If optChoose(0).Value Then
      For Each oText In txtCode
         If oText.Enabled Then
            If CheckKeyIn(oText.Index) = False Then
               i = oText.Index
               '本所案號錯誤時,將Cursor跳回系統別欄位
               If i = 3 Or i = 7 Then i = i - 3
               txtCode(i).SetFocus
               txtCode_GotFocus i
               bolCheckOk = False
               Exit For
            End If
         End If
      Next
   End If
   
   'Added by Morgan 2012/2/9
   If bolCheckOk Then
      If optChoose(0).Value Then
         If Not ((txtCode(8) = "1" And m_bInsert) Or (txtCode(8) = "2" And m_bUpdate) Or (txtCode(8) = "4" And m_bDelete) Or (txtCode(8) = "5" And m_bQuery)) Then
            MsgBox "無權限!!", vbExclamation
            bolCheckOk = False
         End If
      Else
         If m_bQuery = False Then
            MsgBox "無權限!!", vbExclamation
            bolCheckOk = False
         End If
      End If
   End If
   'end 2012/2/9
   
   If bolCheckOk Then
   
      '單筆
      If optChoose(0).Value Then
         '刪除
         If txtCode(8) = "4" Then
            'Modified by Morgan 2018/6/8 兩案建立順序不定,正反向都要檢查
            'If PUB_ChkExist(strCode(), IIf(m_CM10 <> "", Val(m_CM10), 3)) Then
            For i = 0 To 3
               strCode(i) = txtCode(i + 4)
               strCode(i + 4) = txtCode(i)
            Next
            bolExists = PUB_ChkExist(strCode(), IIf(m_CM10 <> "", Val(m_CM10), 3), False)
            If Not bolExists Then
               'User輸入的順序放後面檢查,新增才會照此順序
               For i = 0 To 7
                  strCode(i) = txtCode(i)
               Next
               bolExists = PUB_ChkExist(strCode(), IIf(m_CM10 <> "", Val(m_CM10), 3), False)
            End If
            If Not bolExists Then
               MsgBox "關聯不存在，請重新輸入 !", vbExclamation
            Else
            'end 2017/1/20
               If MsgBox("是否要刪除此筆資料 ?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
                  'Modified by Morgan 2017/10/13 +第3參數傳False
                  If PUB_DeleteCaseRelation(strCode(), IIf(m_CM10 <> "", Val(m_CM10), 3), False) Then
                     For i = 0 To 8
                        txtCode(i) = ""
                     Next
                     txtCode(0).SetFocus
                  End If
               End If
            End If
         
         Else
            For i = 0 To 7
               strCode(i) = txtCode(i)
            Next
            '新增
            If txtCode(8) = "1" Then
               If Not PUB_ChkCaseMap(strCode, IIf(m_CM10 <> "", Val(m_CM10), 3)) Then
                  txtCode(0).SetFocus
                  bolCheckOk = False
               End If
            End If
            
            If bolCheckOk Then
               Set frm040109_2.frmParent = Me
               frm040109_2.strCode1 = strCode(0)
               frm040109_2.strCode2 = strCode(1)
               frm040109_2.strCode3 = strCode(2)
               frm040109_2.strCode4 = strCode(3)
               frm040109_2.strCode5 = strCode(4)
               frm040109_2.strCode6 = strCode(5)
               frm040109_2.strCode7 = strCode(6)
               frm040109_2.strCode8 = strCode(7)
               frm040109_2.intChoose = Val(txtCode(8))
               frm040109_2.m_CM10 = IIf(m_CM10 <> "", Val(m_CM10), 3) 'Added by Morgan 2015/9/10
               frm040109_2.Caption = Me.Caption 'Added by Morgan 2015/9/10
               frm040109_2.Show
               Me.Hide
            End If
         End If
      '多筆
      Else
         frm040109_3.lblCustNo(0) = txtCode(12)
         frm040109_3.lblCustNo(1) = txtCode(13)
         frm040109_3.lblDate(0) = ChangeTStringToTDateString(txtCode(10))
         frm040109_3.lblDate(1) = ChangeTStringToTDateString(txtCode(11))
         frm040109_3.m_CM10 = Me.m_CM10 'Added by Morgan 2015/9/10
         frm040109_3.Caption = Me.Caption 'Added by Morgan 2015/9/10
         Me.Hide
      End If
   End If
                           
End Sub

Private Function finalCheck(Optional p_Mode As String = "2") As Boolean
End Function
Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         Dim varSaveCursor
         varSaveCursor = Screen.MousePointer
         Screen.MousePointer = vbHourglass
         Process
         Screen.MousePointer = varSaveCursor
      Case 1
         Unload Me
   End Select
End Sub

Private Sub Form_Activate()
   If m_CM10 = "6" Then Me.Caption = "擬制喪失新穎性案件資料維護" 'Added by Morgan 2015/9/10
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   'Added by Morgan 2012/2/9
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If TypeName(Me.frmParent) <> "Nothing" Then
      Me.frmParent.Show
   End If
   Set frm040109_1 = Nothing
End Sub

Private Sub optChoose_Click(Index As Integer)
   fraChoose(Index).Enabled = True
   fraChoose((Index + 1) Mod 2).Enabled = False
   If Index = 0 Then
      txtCode(0).SetFocus
   Else
      txtCode(12).SetFocus
   End If
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txtCode(Index).IMEMode = 2
   CloseIme
   TextInverse txtCode(Index)
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   
   'Added by Morgan 2012/2/9
   If Index = 8 Then
      If Not (KeyAscii = 8 Or (Chr(KeyAscii) = "1" And m_bInsert) Or (Chr(KeyAscii) = "2" And m_bUpdate) Or (Chr(KeyAscii) = "4" And m_bDelete) Or (Chr(KeyAscii) = "5" And m_bQuery)) Then
         KeyAscii = 0
         Beep
      End If
   End If
End Sub

Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
   If CheckKeyIn(Index) = False Then
      '本所案號錯誤時,讓Cursor繼續往下跳
      If Index <> 3 And Index <> 7 Then
         Cancel = True
         txtCode_GotFocus Index
      End If
   End If
End Sub
Private Function CheckKeyIn(intIndex As Integer) As Boolean
   Dim intCaseKind As Integer, intWhere As Integer, strTemp As String
   
   Select Case intIndex
      'Added by Morgan 2013/7/9 開放FCP也可使用加控制可用專利系統別
      Case 0, 4
         If txtCode(intIndex) <> "" Then
         
           '  If InStr("," & Systemkind_g_P & ",", "," & txtCode(intIndex) & ",") > 0 Then
           If InStr("," & Systemkind_g_P & ",", "," & txtCode(intIndex) & ",") > 0 _
              Or (FMP2open = True And (txtCode(intIndex) = "P" Or txtCode(intIndex) = "PS")) Then
               CheckKeyIn = True
            Else
               MsgBox "無 " & txtCode(intIndex) & " 案權限！", vbCritical
            End If
         End If
      'end 2013/7/9
         
      Case 3, 7   '本所案號
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.CheckCaseCodeIsExist(txtCode(intIndex - 3), txtCode(intIndex - 2), _
              IIf(txtCode(intIndex - 1) = "", "0", txtCode(intIndex - 1)), IIf(txtCode(intIndex) = "", "00", txtCode(intIndex))) Then
        'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
        If FMP2open = False Then
            If ClsPDCheckCaseCodeIsExist(txtCode(intIndex - 3), txtCode(intIndex - 2), _
                 IIf(txtCode(intIndex - 1) = "", "0", txtCode(intIndex - 1)), IIf(txtCode(intIndex) = "", "00", txtCode(intIndex))) Then
               CheckKeyIn = True
            End If
        Else
          If PUB_FMPtoCheck(0, 1, Pub_strUserST05, txtCode(intIndex - 3), txtCode(intIndex - 2), _
             IIf(txtCode(intIndex - 1) = "", "0", txtCode(intIndex - 1)), IIf(txtCode(intIndex) = "", "00", txtCode(intIndex))) = True Then
               CheckKeyIn = True
          End If
        End If
        
      Case 8   '功能碼
         If Val(txtCode(intIndex)) = 1 Or Val(txtCode(intIndex)) = 2 Or Val(txtCode(intIndex)) = 4 Or Val(txtCode(intIndex)) = 5 Then
            CheckKeyIn = True
         Else
            ShowMsg MsgText(9198)
         End If

      Case 10, 11 '收文日
         If txtCode(intIndex) = "" Then
            CheckKeyIn = True
         ElseIf CheckIsTaiwanDate(txtCode(intIndex)) Then
            CheckKeyIn = True
         End If
         If intIndex = 11 And txtCode(10) <> "" And txtCode(11) = "" Then
            ShowMsg MsgText(9169)
            CheckKeyIn = False
         ElseIf txtCode(11) <> "" And Val(txtCode(10)) > Val(txtCode(11)) Then
            ShowMsg MsgText(9170)
            CheckKeyIn = False
         End If
         
      Case 12
         If Len(txtCode(12)) = 6 Then
            txtCode(13) = txtCode(12) & "999"
            txtCode(12) = txtCode(12) & "000"
         End If
         CheckKeyIn = True
      Case 13 '客戶代碼
         If txtCode(13) = "" Then
            If txtCode(12) <> "" Then
               If MsgBox("客戶代碼起迄都要輸入，是否要清除客戶代碼條件！", vbYesNo + vbDefaultButton1) = vbYes Then
                  txtCode(12) = ""
                  CheckKeyIn = True
               End If
            Else
               CheckKeyIn = True
            End If
         ElseIf txtCode(12) = "" Then
            MsgBox "請先輸入起始客戶代碼！", vbExclamation
            txtCode(13) = ""
         ElseIf Len(txtCode(13)) <> 9 Then
            MsgBox "客戶代碼需輸入九碼！", vbExclamation
         ElseIf Left(txtCode(12), 6) <> Left(txtCode(13), 6) Then
               MsgBox "客戶代碼前六碼需相同！", vbCritical
         ElseIf txtCode(13) < txtCode(12) Then
            MsgBox "客戶代碼區間輸入錯誤！", vbCritical
         Else
            CheckKeyIn = True
         End If
         
      Case Else
         CheckKeyIn = True
         
   End Select
   
End Function
