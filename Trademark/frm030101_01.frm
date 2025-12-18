VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm030101_01 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文"
   ClientHeight    =   5760
   ClientLeft      =   5760
   ClientTop       =   1610
   ClientWidth     =   9350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9350
   Begin VB.CommandButton cmdExtent 
      Caption         =   "延期(&D)"
      Height          =   400
      Left            =   5472
      TabIndex        =   10
      Top             =   72
      Width           =   912
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6420
      TabIndex        =   11
      Top             =   72
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   7380
      TabIndex        =   8
      Top             =   72
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8340
      TabIndex        =   9
      Top             =   72
      Width           =   888
   End
   Begin VB.TextBox textCP09 
      Height          =   264
      Left            =   1560
      MaxLength       =   9
      TabIndex        =   5
      Top             =   600
      Width           =   2892
   End
   Begin VB.TextBox textTM02_2 
      Height          =   264
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   900
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox textTM02 
      Height          =   264
      Left            =   2280
      MaxLength       =   6
      TabIndex        =   1
      Top             =   900
      Width           =   1092
   End
   Begin VB.OptionButton radio 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1332
   End
   Begin VB.OptionButton radio 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   900
      Value           =   -1  'True
      Width           =   1452
   End
   Begin VB.TextBox textTM01 
      Height          =   264
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   0
      Top             =   900
      Width           =   732
   End
   Begin VB.TextBox textTM03 
      Height          =   264
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   3
      Top             =   900
      Width           =   372
   End
   Begin VB.TextBox textTM04 
      Height          =   264
      Left            =   3720
      MaxLength       =   2
      TabIndex        =   4
      Top             =   888
      Width           =   732
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   4392
      Left            =   96
      TabIndex        =   12
      Top             =   1224
      Width           =   9132
      _ExtentX        =   16104
      _ExtentY        =   7743
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm030101_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/17 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Lydia 2021/08/10 改成Form2.0 ; grdList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/10 日期欄已修改
Option Explicit

' 使用者所選取的查詢方式是收文號還是本所案號
Dim m_KeySel As Integer
' 使用者所選取的收文號
Public m_CP09 As String
' 使用者所選取的列其位置
Dim m_CurrSel As Integer
' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
Dim m_CP14NM As String 'Add By Sindy 2013/3/22
'Add By Sindy 2024/8/14
Public bolIsEMPFlow As Boolean '是否為電子承辦簽核
Public m_EEP01 As String
Dim bolFirst As Boolean
'2024/8/14 END


Public Sub Clear()
   textCP09 = Empty
   'textTM01 = Empty
   textTM02 = Empty
   textTM02_2 = Empty
   textTM03 = Empty
   textTM04 = Empty
   InitialGrdList
   radio(0).Value = True
   radio(1).Value = False
   radio_Click 0
End Sub

'Add By Cheng 2002/01/10
Public Sub Clear1()
   textCP09 = Empty
   'textTM01 = Empty
   textTM02 = Empty
   textTM02_2 = Empty
   textTM03 = Empty
   textTM04 = Empty
   InitialGrdList
   cmdQuery.Default = True
   textTM02.SetFocus   'add by sonia 2016/9/9
End Sub

' 使用者按下延期的按紐
Private Sub cmdExtent_Click()
   
   'Add By Cheng 2002/07/15
   '所點選的案件性質不可為"延期"
   If PUB_CPKindDelay(Me.grdList.TextMatrix(Me.grdList.row, 6), "T") Then
      Exit Sub
   End If
   
   ' 檢查是否資料以完全輸入
   If CheckDataValid = True Then 'Add By Sindy 2013/3/22 +if
      'Add By Cheng 2002/07/12
      '若案件已閉卷, 不可發文
      If PUB_CaseClosedCP09(Me.grdList.TextMatrix(Me.grdList.row, 6)) = True Then
         Exit Sub
      End If
      
      'Add By Sindy 2024/8/14
      '檢查是否有承辦歷程是否有產生承辦單可以發文
      If PUB_IsEmpFlowIsSend(m_CP09) = False Then
         Exit Sub
      End If
      '2024/8/14 END
      
      frm030101_11.SetData 0, m_CP09, True
      Me.Hide
      frm030101_11.Show
      frm030101_11.QueryData
   End If
End Sub

Private Sub cmdok_Click()
'Add By Cheng 2002/07/11
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

   ' 檢查是否資料以完全輸入
   If CheckDataValid = True Then
      'Add By Cheng 2002/07/12
      '若案件已閉卷, 不可發文
      If PUB_CaseClosedCP09(Me.grdList.TextMatrix(Me.grdList.row, 6)) = True Then
         Exit Sub
      End If
      
      'Add By Sindy 2024/8/14
      '檢查是否有承辦歷程是否有產生承辦單可以發文
      If PUB_IsEmpFlowIsSend(m_CP09) = False Then
         Exit Sub
      End If
      '2024/8/14 END
      
      'Modify By Sindy 2024/1/23 改為共用函數
      If PUB_ChkCP141IsSend(m_CP09) = False Then
         Exit Sub
      End If
'      'Add By Sindy 2023/12/11
'      strExc(0) = "select cp06,nvl(cp79,0) cp79,cp141,cp142,cp164 from caseprogress where cp09='" & m_CP09 & "'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
''         If "" & RsTemp.Fields("cp141") = "2" And RsTemp.Fields("cp79") > 0 Then
''             If PUB_ChkPaidByCP09(m_CP09) = False Then   'Added by Morgan 2016/8/23 出納繳款確認後就可送件
''                If IsNull(RsTemp.Fields("cp06")) Or "" & RsTemp.Fields("cp06") > strSrvDate(1) Then
''                   MsgBox "此案智權人員欲管控收款後才可送件，暫不可發文！"
''                   Exit Sub
''                End If
''             End If
''         Else
'         If "" & RsTemp.Fields("cp141") = "3" Then
'            '1=當天
'            If "" & RsTemp.Fields("cp164") = "1" And "" & RsTemp.Fields("cp142") <> "" And "" & RsTemp.Fields("cp142") > strSrvDate(1) Then
'               MsgBox "本案需於指定日" & ChangeWStringToTDateString(RsTemp.Fields("cp142")) & "方可發文！"
'               Exit Sub
'            '3=之後
'            ElseIf "" & RsTemp.Fields("cp164") = "3" And "" & RsTemp.Fields("cp142") <> "" And "" & RsTemp.Fields("cp142") >= strSrvDate(1) Then
'               MsgBox "本案需於指定日" & ChangeWStringToTDateString(RsTemp.Fields("cp142")) & "之後方可發文！"
'               Exit Sub
'            End If
'         End If
'      End If
'      '2023/12/11 END
   
      'Add By Cheng 2002/07/11
      '檢查所點選的案件進度資料當案件性質為自請撤回"306"或自請撤銷"307"時, 其相關總收文號若為空白, 則不可進入下一畫面
      StrSQLa = "Select * From CaseProgress Where CP09='" & m_CP09 & "' AND (CP10='306' OR CP10='307') AND CP43 IS NULL "
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         MsgBox "該自撤案件未輸入相關總收文號, 請先補齊資料!!!", vbExclamation + vbOKOnly
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
         Exit Sub
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   
      '91.7.16此段錯誤, 正確的控制在DisplayNextForm的ShowMaintainForm m_CP09
      'Modify By Cheng 2002/07/11
      '檢查是否要顯示商檔基本檔資料維護的畫面暫時取消
'      ' 檢查是否要顯示商標基本檔資料維護的畫面
'      If CheckJumpFrm020501() = True Then
'         DisplayFrm020501
'      Else
         ' 檢查是否已收款
         'Modified by Morgan 2016/8/23 出納繳款確認後就可送件
         'If CheckIfFinishCP79() = False Then
'Remove by Lydia 2018/08/22  (應收帳款管控)取消預定收款日,改成付款週期=>不發email
'         If CheckIfFinishCP79() = False And PUB_ChkPaidByCP09(m_CP09) = False Then
'         'end 2016/8/23
'            ' 若未收款時則顯示 frm030101_02 的畫面
'            DisplayFrm030101_02
'         Else
            ' 若已收款時則依案件性質顯示下一個畫面
            DisplayNextForm
'         End If
'      End If
'end 2018/08/22
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Initial
   InitialGrdList
   UpdateCtrlState
End Sub

Private Sub Initial()
   ' 預設由收文號來取得資料
   'modify by sonia 2016/9/9 改本所案號,同時改form之radio(1).Value為true及欄位tabindex順序
   'm_KeySel = 0
   m_KeySel = 1
End Sub

' 按下結束離開按紐
Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload Me
End Sub
' 按下查詢按紐
Private Sub cmdQuery_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   ' 先檢查該輸入的資料是否有全部輸入
   Select Case m_KeySel
      ' 依收文號
      Case 0:
         If IsEmptyText(textCP09) = True Then
            strTit = "資料檢核"
            strMsg = "請輸入收文號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
      ' 依本所案號
      Case 1:
         If IsEmptyText(textTM01) = True Or IsEmptyText(textTM02) = True Then
            strTit = "資料檢核"
            strMsg = "請輸入本所案號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
   End Select
   ' 查詢資料
   If QueryData() = False Then
      strTit = "資料查詢"
      strMsg = "沒有符合條件的資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   Else
      cmdOK.Default = True
   End If
EXITSUB:
End Sub

' 由案件性質代碼取得國內案件性質名稱
Private Function GetCaseType(ByVal strKey1 As String, ByVal StrKey2 As String) As String
   Dim rsTmp As ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   GetCaseType = Empty
   If IsEmptyText(strKey1) = False And IsEmptyText(StrKey2) = False Then
      Set rsTmp = New ADODB.Recordset
      strSql = "SELECT * FROM CasePropertyMap " & _
               "WHERE CPM01 = '" & strKey1 & "' AND " & _
                     "CPM02 = '" & StrKey2 & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("CPM03")) = False Then
            GetCaseType = rsTmp.Fields("CPM03")
         End If
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Function

' 取得員工姓名
Private Function GetStaffName(ByVal strKey As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   GetStaffName = Empty
   strSql = "SELECT * FROM Staff " & _
            "WHERE ST01 = '" & strKey & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("ST02")) = False Then
         GetStaffName = rsTmp.Fields("ST02")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Public Sub RefreshData()
   Dim bQuery As Boolean
   bQuery = QueryData
End Sub

' 查詢資料庫
Public Function QueryData() As Boolean
   Dim strSql As String
   Dim strCP01 As String
   Dim strCP02 As String
   Dim strCP03 As String
   Dim strCP04 As String
   Dim rsTmp As New ADODB.Recordset
   
   QueryData = False
   m_CP09 = Empty
   m_CP14NM = Empty 'Add By Sindy 2013/3/22
   InitialGrdList
   
   'Add By Sindy 2024/8/14 控管多筆未發文時,後面的資料不能直接進入發文作業
   If bolFirst = True Then
      bolIsEMPFlow = False
   End If
   bolFirst = True
   '2024/8/14 END
   
   ' 組成SQL語法
   Select Case m_KeySel
      ' 依收文號
      Case 0:
         strSql = "SELECT * FROM CaseProgress " & _
                  "WHERE CP09 = '" & textCP09 & "' "
      ' 依本所案號
      Case 1:
         strCP01 = Trim(textTM01)
         strCP02 = Trim(textTM02)
         strCP03 = Trim(textTM03)
         If IsEmptyText(strCP03) = True Then: strCP03 = "0"
         strCP04 = Trim(textTM04)
         If IsEmptyText(strCP04) = True Then: strCP04 = "00"
         strSql = "SELECT * FROM CaseProgress " & _
                  "WHERE CP01 = '" & strCP01 & "' AND " & _
                        "CP02 = '" & strCP02 & "' AND " & _
                        "CP03 = '" & strCP03 & "' AND " & _
                        "CP04 = '" & strCP04 & "' "
   End Select
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 列出所有資料
   If rsTmp.RecordCount > 0 Then
      If ListData(rsTmp) = True Then
         QueryData = True
      End If
   End If
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Function

' 列出所有符合條件的資料
Private Function ListData(ByRef rsTmp As ADODB.Recordset) As Boolean
Dim nRow As Integer
Dim i As Integer, intQRow As Integer
   
   ListData = False
   If rsTmp.RecordCount <= 0 Then: GoTo EXITSUB
   rsTmp.MoveFirst
   Do While rsTmp.EOF = False
      ' 系統類別必須為"CFT", "CFC", "S"
      If IsNull(rsTmp.Fields("CP01")) = False Then
         Select Case rsTmp.Fields("CP01")
            'Modify By Sindy 2015/10/19 +CFL的901催款
            Case "CFT", "CFC", "S", "CFL":
            Case Else: GoTo EXITSUB
         End Select
      End If
      
      'Add By Sindy 2015/10/19
      If rsTmp.Fields("CP01") = "CFL" Then
         If rsTmp.Fields("CP10") <> "901" Then
            GoTo NextRecord
         End If
      End If
      '2015/10/19 END
      
      '收文號不為A,B類的不予計入
      '2008/12/11 CANCEL BY SONIA 因改為期限專業部直寄,CFT來函不先上發文日,故此處不可限制
      'Select Case Mid(rsTmp.Fields("CP09"), 1, 1)
      '   Case "A", "B":
      '   Case Else: GoTo NextRecord
      'End Select
      '2008/12/11 END
      ' 尚未輸入發文日
      If IsNull(rsTmp.Fields("CP27")) = False Then
         If IsEmptyText(rsTmp.Fields("CP27")) = False Then
            If rsTmp.Fields("CP27") <> "0" Then: GoTo NextRecord
         End If
      End If
      ' 尚未輸入取消收文日期
      If IsNull(rsTmp.Fields("CP57")) = False Then
         If IsEmptyText(rsTmp.Fields("CP57")) = False Then
            If rsTmp.Fields("CP57") <> "0" Then: GoTo NextRecord
         End If
      End If
               
      grdList.Rows = grdList.Rows + 1
      nRow = grdList.Rows - 1
      ' 收文日欄位
      If IsNull(rsTmp.Fields("CP05")) = False Then
         grdList.TextMatrix(nRow, 1) = rsTmp.Fields("CP05")
      End If
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
      
         strExc(1) = rsTmp.Fields("CP01")
         strExc(2) = rsTmp.Fields("CP02")
         strExc(3) = rsTmp.Fields("CP03")
         strExc(4) = rsTmp.Fields("CP04")
         'edit by nickc 2007/02/06 不用 dll 了
         'If objPublicData.GetSystemKind(strExc(1), intI) Then
         If ClsPDGetSystemKind(strExc(1), intI) Then
            If intI = 2 Then '商標
               strExc(0) = "SELECT TM10 FROM TRADEMARK WHERE TM01='" & strExc(1) & "' AND TM02='" & strExc(2) & "' AND TM03='" & strExc(3) & "' AND TM04='" & strExc(4) & "'"
            'Add By Sindy 2015/10/19
            ElseIf intI = 3 Then '法務
               strExc(0) = "SELECT LC15 FROM LAWCASE WHERE LC01='" & strExc(1) & "' AND LC02='" & strExc(2) & "' AND LC03='" & strExc(3) & "' AND LC04='" & strExc(4) & "'"
            '2015/10/19 END
            Else
               strExc(0) = "SELECT SP09 FROM SERVICEPRACTICE WHERE SP01='" & strExc(1) & "' AND SP02='" & strExc(2) & "' AND SP03='" & strExc(3) & "' AND SP04='" & strExc(4) & "'"
            End If
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))   'edit by nickc 2007/02/06 不用 dll 了   = objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp.Fields(0) < "010" Then
                  grdList.TextMatrix(nRow, 2) = GetCaseTypeName(rsTmp.Fields("CP01"), rsTmp.Fields("CP10"), 0)
               Else
                  grdList.TextMatrix(nRow, 2) = GetCaseTypeName(rsTmp.Fields("CP01"), rsTmp.Fields("CP10"), 1)
               End If
            End If
         End If
      End If
      ' 承辦人
      If IsNull(rsTmp.Fields("CP14")) = False Then
         grdList.TextMatrix(nRow, 3) = GetStaffName(rsTmp.Fields("CP14"))
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         grdList.TextMatrix(nRow, 4) = GetStaffName(rsTmp.Fields("CP13"))
      End If
      ' 進度備註
      If IsNull(rsTmp.Fields("CP64")) = False Then
         grdList.TextMatrix(nRow, 5) = rsTmp.Fields("CP64")
      End If
      ' 收文號
      If IsNull(rsTmp.Fields("CP09")) = False Then
         grdList.TextMatrix(nRow, 6) = rsTmp.Fields("CP09")
      End If
      'Add By Sindy 2010/12/27 判斷有相關總收文號才做
      ' 相關總收文號
      If IsNull(rsTmp.Fields("CP43")) = False Then
         '案件性質
         grdList.TextMatrix(nRow, 2) = grdList.TextMatrix(nRow, 2) & PUB_GetRelateCasePropertyName(grdList.TextMatrix(nRow, 6), "1")
      End If
      '2010/12/27 End
      ListData = True
NextRecord:
      rsTmp.MoveNext
   Loop
   
   'Added by Lydia 2023/10/17
   If grdList.Rows >= 2 Then
      grdList.FixedRows = 1
   End If
   'end 2023/10/17
   
   ' 顯示符合的所有資料
   grdList.Refresh
   
   'Modify By Sindy 2024/8/14
   ' 設定第一筆為被選取的狀態
   'grdList_SetSelection 1
   If grdList.Rows >= 2 Or (bolIsEMPFlow = True And m_EEP01 <> "") Then
      If (bolIsEMPFlow = True And m_EEP01 <> "") Then
         For i = 1 To grdList.Rows - 1
            If grdList.TextMatrix(i, 6) = m_EEP01 Then
               intQRow = i
               Exit For
            End If
         Next i
      Else
        intQRow = 1 '若有資料游標停在第一筆
      End If
   End If
   If intQRow > 0 Then
      grdList_SetSelection intQRow
      If bolIsEMPFlow = True Then Call cmdok_Click
   End If
   '2024/8/14 END
   
EXITSUB:
End Function

' 更新控制項的狀態
Private Sub UpdateCtrlState()
   Select Case m_KeySel
      Case 0:
         EnableTextBox textCP09, True
         EnableTextBox textTM01, False
         EnableTextBox textTM02, False
         EnableTextBox textTM03, False
         EnableTextBox textTM04, False
         textTM02_2.Visible = False
      Case 1:
         EnableTextBox textCP09, False
         EnableTextBox textTM01, True
         EnableTextBox textTM02, True
         EnableTextBox textTM03, True
         EnableTextBox textTM04, True
         textTM01_Validate False
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '列印接洽接案單
    PUB_PrintCaseCloseSheet strUserNum
    '刪除暫存資料
    PUB_DeleteCaseCloseSheet strUserNum
    'Add By Cheng 2002/07/19
   Set frm030101_01 = Nothing
End Sub

'Add By Cheng 2002/01/10
Private Sub grdList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If grdList.Rows > 1 Then
   If grdList.row > 0 Then
      m_CP09 = grdList.TextMatrix(grdList.row, 6)
      m_CP14NM = grdList.TextMatrix(grdList.row, 3) 'Add By Sindy 2013/3/22
   End If
End If
grdList_ShowSelection
End Sub

' 使用者按下所選取的項目
Public Sub radio_Click(Index As Integer)
   '******* 90.11.23 nick
   If frm030101_01.Visible = True Then
   m_KeySel = Index
   UpdateCtrlState
   
   ' 90.07.25 modify
   Select Case Index
      Case 0:
         textCP09.SetFocus
      Case 1:
         textTM01.SetFocus
   End Select
   End If
   '****************************
End Sub

Private Sub textCP09_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   '2005/8/29 ADD BY SONIA
   cmdQuery.Default = True
End Sub

Private Sub textTM01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 檢查系統別
Private Sub textTM01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM01) = False Then
      Select Case textTM01
         'Modify By Sindy 2015/10/19 +CFL的901催款
         Case "CFT", "CFC", "S", "CFL":
            textTM02_2.Visible = False
            textTM02_2.Locked = True
            textTM02_2.TabStop = False
            textTM02.MaxLength = 6
         Case Else
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
End Sub
' 初始化 GridList
Private Sub InitialGrdList()
   Dim nIndex As Integer
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 7
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "收文日"
   grdList.ColWidth(1) = 1200
   grdList.col = 2
   grdList.Text = "案件性質"
   grdList.ColWidth(2) = 1000
   grdList.col = 3
   grdList.Text = "承辦人"
   grdList.ColWidth(3) = 1200
   grdList.col = 4
   grdList.Text = "智權人員"
   grdList.ColWidth(4) = 1000
   grdList.col = 5
   grdList.Text = "進度備註"
   grdList.ColWidth(5) = 1200
   ' 收文號欄位
   grdList.col = 6
   grdList.Text = "收文號"
   grdList.ColWidth(6) = 0
   ' 隱藏欄位
   'nIndex = 6
   'grdList.ColIsVisible(nIndex) = False
End Sub

' 設定Grid List的一列為選取的狀態
Private Sub grdList_SetSelection(ByVal nSel As Integer)
   If nSel > 0 And nSel < grdList.Rows And grdList.Rows >= 2 Then
      grdList.row = nSel
      grdList_SelChange
   End If
End Sub

Private Sub grdList_SelChange()
   If grdList.Rows > 1 Then
      If grdList.row > 0 Then
         m_CP09 = grdList.TextMatrix(grdList.row, 6)
         m_CP14NM = grdList.TextMatrix(grdList.row, 3) 'Add By Sindy 2013/3/22
      End If
   End If
   grdList_ShowSelection
   'cmdOK.SetFocus
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = grdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = grdList.row Then
      Dim nOldCol As Integer
      nOldCol = grdList.col
      grdList.col = 1
      If grdList.CellBackColor <> &H8000000D Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H8000000D Then grdList.CellBackColor = &H8000000D
            If grdList.CellForeColor <> &H80000005 Then grdList.CellForeColor = &H80000005
         Next nCol
      End If
      grdList.col = nOldCol
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      If grdList.CellBackColor <> &H80000005 Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
            If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
         Next nCol
      End If
      grdList.col = 0
   End If
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      For nCol = 1 To grdList.Cols - 1
         grdList.col = nCol
         grdList.CellBackColor = &H8000000D
         grdList.CellForeColor = &H80000005
      Next nCol
      grdList.col = 0
      cmdOK.Default = True
   End If
EXITSUB:
End Sub

' 顯示下一個畫面
Public Sub DisplayNextForm()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strCP10 As String
   Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
   Dim frmNext As Form
   
   strCP10 = Empty
   ' 組成SQL語法
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 列出所有資料
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CP01")) = False Then
         strCP01 = rsTmp.Fields("CP01")
      End If
      'Add By Sindy 2016/8/11
      If IsNull(rsTmp.Fields("CP02")) = False Then
         strCP02 = rsTmp.Fields("CP02")
      End If
      If IsNull(rsTmp.Fields("CP03")) = False Then
         strCP03 = rsTmp.Fields("CP03")
      End If
      If IsNull(rsTmp.Fields("CP04")) = False Then
         strCP04 = rsTmp.Fields("CP04")
      End If
      '2016/8/11 END
      If IsNull(rsTmp.Fields("CP10")) = False Then
         strCP10 = rsTmp.Fields("CP10")
      End If
      'Add By Sindy 2024/12/11
      If IsNull(rsTmp.Fields("CP14")) = True Then
         MsgBox "未分案不可發文!!!", vbExclamation + vbOKOnly
         Exit Sub
      End If
      '2024/12/11 END
   End If
   rsTmp.Close
   
   Select Case strCP01
      'Add By Sindy 2015/10/21
      Case "CFL"
         frm071006.SetParent Me
         frm071006.Show
         Me.Hide
         Exit Sub
      '2015/10/21 END
      Case "CFC": Set frmNext = frm030101_06
      Case Else
         Select Case strCP10
            ' 申請, 延展, 補換發證書, 申請英文證明
            'edit by nickc 2006/08/15 304 改用新表單
            'Case "101", "102", "103", "304": Set frmNext = frm030101_03
            'Modify By Sindy 2019/10/16 + 109緩審延展
            Case "101", "102", "103", "109": Set frmNext = frm030101_03
            'add by nickc 2006/08/15 CFT B 類走新畫面
            Case "304":
               '2009/10/2 MODIFY BY SONIA 取消A,B類條件,全部用新畫面frm030101_19(CFT-012741)
               '2012/4/12 modify by sonia 改為內部收文用新畫面,接洽單則詢問是申請台灣案或非台灣案的英文證明(CFT-007279)
               'If strCP01 = "CFT" And UCase(Mid(m_CP09, 1, 1)) = "B" Then
               '   Set frmNext = frm030101_19
               'Else
               '   Set frmNext = frm030101_03
               'End If
               If strCP01 = "CFT" Then
                  If UCase(Mid(m_CP09, 1, 1)) = "B" Then
                     Set frmNext = frm030101_19
                  ElseIf MsgBox("是否申請台灣案的英文證明？", vbExclamation + vbYesNo) = vbYes Then
                     Set frmNext = frm030101_19
                  Else
                     Set frmNext = frm030101_03
                  End If
               Else
                  Set frmNext = frm030101_03
               End If
               '2012/4/12 END
               '2009/10/2 END
            ' 變更, 更正  2007/6/7加減縮商品
            Case "301", "302", "313": Set frmNext = frm030101_07
            ' 移轉
            Case "501": Set frmNext = frm030101_08
            ' 授權, 再授權, 終止授權, 終止再授權
            Case "502", "503", "504", "505": Set frmNext = frm030101_09
            'Modify By Cheng 2002/06/05
'            ' 補正, 答辯
'            Case "201", "202": Set frmNext = frm030101_10
            ' 補正, 答辯, 放棄專用權
            Case "201", "202", "206": Set frmNext = frm030101_10
            ' 延期
            Case "303": Set frmNext = frm030101_11
            ' 自請撤回, 自請撤銷
            Case "306", "307": Set frmNext = frm030101_12
            ' 設定質權, 撤銷設定質權
            Case "506", "507": Set frmNext = frm030101_13
            ' 異議, 評定, 廢止, 評定專用權, 參加評定, 自評專用權, 禁止處分
            Case "601", "603", "605", "607", "608", "609", "616": Set frmNext = frm030101_14
            ' 補充理由, 訴願, 再訴願, 行政訴訟, 參加行政訴訟, 再審之訴
            Case "612", "401", "402", "403", "404", "405": Set frmNext = frm030101_15
            ' 異議答辯, 評定答辯, 廢止答辯, 補充答辯, 參加被評定, 撤銷禁止處分, 修正, 刊登廣告, 其它
            Case "602", "604", "606", "613", "610", "617", "203", "702": Set frmNext = frm030101_16
            ' 補理由書
            Case "611": Set frmNext = frm030101_17
            ' 領證, 使用宣誓
            Case "701", "105": Set frmNext = frm030101_18
            ' 其它
            Case Else: Set frmNext = frm030101_16
         End Select
   End Select
   
   'Add By Sindy 2016/8/11
   If strCP01 = "CFT" And strCP10 = "101" Then
      'Modify by Amy 2018/07/31 ChkIsExistImg不使用
      'If ChkIsExistImg(strCP01, strCP02, strCP03, strCP04, False) = False Then '無代表圖
      If ChkImgByteFile(strCP01, strCP02, strCP03, strCP04) = False Then
         If MsgBox("是否插入商標圖？", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbYes Then
            frmPic001.oCP01 = strCP01
            frmPic001.oCP02 = strCP02
            frmPic001.oCP03 = strCP03
            frmPic001.oCP04 = strCP04
            frmPic001.StrMenu
            frmPic001.CanScan
            frmPic001.SetSeekCmdok 'Add by Amy 2018/07/19
            frmPic001.Show vbModal
         End If
      End If
   End If
   '2016/8/11 END
   
   ' 顯示下一個畫面
   If IsObject(frmNext) = True Then
      frmNext.SetData 0, m_CP09, True
      '*********** 901121     nick
      If Me.Visible = True Then
         cmdQuery.Default = True
      End If
      '****************************
      Me.Hide
      frmNext.Show
      frmNext.QueryData
      'Modify By Cheng 2002/07/11
      '檢查是否要顯示商檔基本檔資料維護的畫面暫時取消
      ' 顯示商標基本資料的畫面
      'add by nickc 2006/08/30 CFT B 類 304 不秀
      If strCP01 = "CFT" And UCase(Mid(m_CP09, 1, 1)) = "B" And strCP10 = "304" Then
      Else
         'ShowMaintainForm m_CP09
         'Modify By Sindy 2018/2/1 加 101申請,不開商標主檔,但要開申請人地址視窗
         If strCP10 = "101" Then
            'Add By Sindy 2018/2/1 增加案件申請人地址視窗彈跳
            frm020102_23.Hide
            Set frm020102_23.UpForm = frmNext
            frm020102_23.m_CP09 = m_CP09
            'Me.Hide
            frm020102_23.QueryData
            frm020102_23.Show vbModal
            '2018/2/1 End
         Else
            ShowMaintainForm m_CP09
         End If
      End If
   End If
End Sub

' 檢查是否已收款
Private Function CheckIfFinishCP79() As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   CheckIfFinishCP79 = True
   ' 查詢案件進度檔
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CP79")) = False Then
         If IsEmptyText(rsTmp.Fields("CP79")) = False Then
            If rsTmp.Fields("CP79") <> "0" Then
               CheckIfFinishCP79 = False
            End If
         End If
      End If
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Function
' 顯示向智權人員發EMail的畫面
Private Sub DisplayFrm030101_02()
   frm030101_02.SetData 0, m_CP09, True
   Me.Hide
   frm030101_02.Show
   frm030101_02.QueryData
End Sub

' 顯示商標基本檔檔案畫面要求輸入
Private Sub DisplayFrm020501()
   frm020501.SetSystem 0
   frm020501.Show
End Sub
' 若為新案件且非新申請案且卷宗性質為"申請"時, 若商標基本檔的申請案號欄位是空白, 則先切換至商標基本資料維護
Private Function CheckJumpFrm020501() As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim bShowFrm020501 As Boolean
   Dim strTM01 As String
   Dim strTM02 As String
   Dim strTM03 As String
   Dim strTM04 As String
   Dim strTM12 As String
   Dim strTM28 As String
   Dim strCP10 As String
   Dim strCP31 As String
   
   bShowFrm020501 = False
   
   GoTo EXITSUB
   
   strTM12 = Empty
   strCP10 = Empty
   strCP31 = Empty
   
   ' 查詢案件進度檔
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount <= 0 Then: GoTo EXITSUB
   
   If IsNull(rsTmp.Fields("CP01")) = False Then: strTM01 = rsTmp.Fields("CP01")
   If IsNull(rsTmp.Fields("CP02")) = False Then: strTM02 = rsTmp.Fields("CP02")
   If IsNull(rsTmp.Fields("CP03")) = False Then: strTM03 = rsTmp.Fields("CP03")
   If IsNull(rsTmp.Fields("CP04")) = False Then: strTM04 = rsTmp.Fields("CP04")
   ' 案件性質
   If IsNull(rsTmp.Fields("CP10")) = False Then
      If IsEmptyText(rsTmp.Fields("CP10")) = False Then
         strCP10 = rsTmp.Fields("CP10")
      End If
   End If
   ' 是否為新案件欄位
   If IsNull(rsTmp.Fields("CP31")) = False Then
      If IsEmptyText(rsTmp.Fields("CP31")) = False Then
         strCP31 = rsTmp.Fields("CP31")
      End If
   End If
   rsTmp.Close
   
   ' 查詢商標基本檔
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & strTM01 & "' AND " & _
                  "TM02 = '" & strTM02 & "' AND " & _
                  "TM03 = '" & strTM03 & "' AND " & _
                  "TM04 = '" & strTM04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount <= 0 Then: GoTo EXITSUB
   ' 卷宗性質
   If IsNull(rsTmp.Fields("TM28")) = False Then
      If IsEmptyText(rsTmp.Fields("TM28")) = False Then
         strTM28 = rsTmp.Fields("TM28")
      End If
   End If
   rsTmp.Close

   ' 判斷是否要顯示商標基本檔檔案維護的畫面
   If strTM28 = "1" Then
      If UCase(strCP31) = "Y" Then
         If strCP10 <> "101" Then
            bShowFrm020501 = True
         End If
      End If
   End If

   CheckJumpFrm020501 = bShowFrm020501
EXITSUB:
   Set rsTmp = Nothing
End Function
' 檢查是否已選取資料
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   If grdList.Rows <= 1 Then
      strTit = "檢核資料"
      strMsg = "請先選取資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   If IsEmptyText(m_CP09) = True Then
      strTit = "檢核資料"
      strMsg = "請先選取資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   'Add By Sindy 2013/3/22
   If IsEmptyText(m_CP14NM) = True Then
      strTit = "檢核資料"
      strMsg = "尚未分案不可發文！"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   '2013/3/22 End
   
   CheckDataValid = True
EXITSUB:
End Function

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP09 = Empty
   End If
   
   Select Case nType
      ' 本所案號
      Case 0: m_TM01 = strData
      Case 1: m_TM02 = strData
      Case 2: m_TM03 = strData
      Case 3: m_TM04 = strData
   End Select
End Sub

' 更新查詢的方式由本所案號來查詢
Public Sub SetQueryFromTM()
   textTM01 = m_TM01
   textTM02 = m_TM02
   textTM03 = m_TM03
   textTM04 = m_TM04
   radio_Click 1
End Sub

Private Sub textCP09_GotFocus()
   InverseTextBox textCP09
End Sub

Private Sub textTM01_GotFocus()
   InverseTextBox textTM01
   CloseIme
End Sub
'2005/8/29 ADD BY SONIA
Private Sub textTM02_2_KeyPress(KeyAscii As Integer)
   cmdQuery.Default = True
End Sub

Private Sub textTM02_GotFocus()
   InverseTextBox textTM02
End Sub

Private Sub textTM02_2_GotFocus()
   InverseTextBox textTM02_2
End Sub
'2005/8/29 ADD BY SONIA
Private Sub textTM02_KeyPress(KeyAscii As Integer)
   cmdQuery.Default = True
End Sub

Private Sub textTM03_GotFocus()
   InverseTextBox textTM03
   CloseIme
End Sub

Private Sub textTM03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM04_GotFocus()
   InverseTextBox textTM04
End Sub

