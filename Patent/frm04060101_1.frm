VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm04060101_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利公報資料維護"
   ClientHeight    =   5460
   ClientLeft      =   -465
   ClientTop       =   930
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9330
   Begin VB.CommandButton buttonClear 
      Caption         =   "清除查詢結果(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4920
      TabIndex        =   7
      Top             =   600
      Width           =   1500
   End
   Begin VB.CommandButton buttonQuery 
      Caption         =   "查詢(&F)"
      Height          =   400
      Left            =   7560
      TabIndex        =   5
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton buttonSearch 
      Caption         =   "同卷期多筆查詢(&S)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3012
      TabIndex        =   1
      Top             =   600
      Width           =   1860
   End
   Begin VB.CommandButton buttonAdd 
      Caption         =   "新增(&A)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5076
      TabIndex        =   2
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton buttonMod 
      Caption         =   "修改(&M)"
      Height          =   400
      Left            =   5904
      TabIndex        =   3
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton buttonDel 
      Caption         =   "刪除(&D)"
      Height          =   400
      Left            =   6732
      TabIndex        =   4
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton buttonExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8388
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox textQuery 
      Height          =   264
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   0
      Top             =   624
      Width           =   1452
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   4356
      Left            =   96
      TabIndex        =   9
      Top             =   1008
      Width           =   9108
      _ExtentX        =   16060
      _ExtentY        =   7673
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
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
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號 :"
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   648
      Width           =   816
   End
End
Attribute VB_Name = "frm04060101_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/22 改成Form2.0 (grdList)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改(尚有申請號問題)
Option Explicit

' 變數宣告區
Dim m_Recordset As New ADODB.Recordset
Dim m_QueryCommand As String
Dim m_CurrSel As Integer

' 組一個Query的
Private Sub SetQueryCommand()
   
   'm_QueryCommand = "SELECT * FROM TPBulletin, Tagent, Patent " & _
   '                 "WHERE TPB01 LIKE '" & textQuery & "%' AND " & _
   '                       "TPB01 = PA11 (+) AND " & _
   '                       "TA01 = 'P' AND " & _
   '                       "TPB07 = TA02 (+) " & _
   '                 "ORDER BY TPB01 ASC"
   
'   m_QueryCommand = "SELECT TPB01,TPB02,TPB03,TPB04,TPB05,TPB06,TPB07,TPB08,TA01,TA02,TA03," & _
      "PA01,PA02,PA03,PA04,PA09,PA11 FROM TPBULLETIN, PATENT, TAGENT " & _
                     "WHERE TPB01 LIKE '" & textQuery & "%' AND " & _
                           "TPB02 = PA11 (+) AND " & _
                           "TA01 = 'P' AND " & _
                           "TPB07 = TA02(+) AND PA01='000' UNION ALL " & _
                     "SELECT TPB01,TPB02,TPB03,TPB04,TPB05,TPB06,TPB07,TPB08,'','',''," & _
      "PA01,PA02,PA03,PA04,PA09,PA11 FROM TPBULLETIN, PATENT " & _
                     "WHERE TPB01 LIKE '" & textQuery & "%' AND " & _
                           "TPB02 = PA11 (+) AND " & _
                           "TPB07 IS NULL " & _
                           " ORDER BY TPB01 "
   
'910709 Sieg 412
   m_QueryCommand = "SELECT '',T1.TPB01,T1.TPB02," & SQLDate("T1.TPB03") & "," & _
      "T1.TPB04||'-'||T1.TPB05,NA03,TA03," & ChgPatent("", 1) & "," & _
      "T1.TPB08 AS TPB08 " & _
      "FROM TPBULLETIN T1, TPBULLETIN T2, PATENT, TAGENT,NATION " & _
      "WHERE T2.TPB01 = '" & textQuery & "' AND T1.TPB04=T2.TPB04 AND T1.TPB05=T2.TPB05 AND " & _
      "T1.TPB01 = PA11 (+) AND 'P'=TA01(+) AND " & _
      "T1.TPB07 = TA02(+) AND '000'=PA09(+) AND T1.TPB06=NA01(+) ORDER BY T1.TPB02 DESC"

End Sub

Private Sub InitialGrdList()
   FixGrid GrdList
   GrdList.row = 0
   GrdList.col = 0
   GrdList.Text = ""
   GrdList.ColWidth(0) = 300
   GrdList.col = 1
   GrdList.Text = "申請案號"
   GrdList.ColWidth(1) = 1200
   GrdList.ColAlignment(1) = flexAlignLeftCenter
   GrdList.col = 2
   GrdList.Text = "公告號"
   GrdList.ColWidth(2) = 1000
   GrdList.ColAlignment(2) = flexAlignCenterCenter
   GrdList.col = 3
   GrdList.Text = "公告日"
   GrdList.ColWidth(3) = 1000
   GrdList.ColAlignment(3) = flexAlignCenterCenter
   GrdList.col = 4
   GrdList.Text = "卷期"
   GrdList.ColWidth(4) = 600
   GrdList.ColAlignment(4) = flexAlignCenterCenter
   GrdList.col = 5
   GrdList.Text = "申請人國籍"
   GrdList.ColWidth(5) = 1000
   GrdList.ColAlignment(5) = flexAlignLeftCenter
   GrdList.col = 6
   GrdList.Text = "代理人"
   GrdList.ColWidth(6) = 1000
   GrdList.ColAlignment(6) = flexAlignLeftCenter
   GrdList.col = 7
   GrdList.Text = "本所案號"
   GrdList.ColWidth(7) = 1200
   GrdList.ColAlignment(7) = flexAlignLeftCenter
   GrdList.col = 8
   GrdList.Text = "事務所名稱"
   GrdList.ColWidth(8) = 1200
   GrdList.ColAlignment(8) = flexAlignLeftCenter
End Sub

Private Function GetNation(ByVal strNation As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   GetNation = Empty
   If strNation <> Empty Then
      strSql = "SELECT * FROM NATION WHERE NA01 = '" & strNation & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly 'edit by nickc 2007/02/06 , adOpenDynamic
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         GetNation = rsTmp.Fields("NA03")
      End If
   End If
   
   Set rsTmp = Nothing
End Function
' 取得本所案號
Private Function GetPNumber(ByVal strTPB01 As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   GetPNumber = Empty
   strSql = "SELECT * FROM Patent WHERE PA11 = '" & strTPB01 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      GetPNumber = rsTmp.Fields("PA01") & "-" & rsTmp.Fields("PA02") & "-" & rsTmp.Fields("PA03") & "-" & rsTmp.Fields("PA04")
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function
' 取得代理人名稱
Private Function GetAgentName(ByVal strAgent As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   GetAgentName = Empty
   strSql = "SELECT * FROM Tagent " & _
            "WHERE TA01 = 'P' AND " & _
                  "TA02 = '" & strAgent & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      GetAgentName = rsTmp.Fields("TA03")
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Sub ExecuteQuery()
 Dim i As Integer
   
   '910709 Sieg 412
   Screen.MousePointer = vbHourglass
   intI = 0
   'edit by nickc 2007/02/05 不用 dll 了
   'Set RsTemp = objLawDll.ReadRstMsg(intI, m_QueryCommand)
   Set RsTemp = ClsLawReadRstMsg(intI, m_QueryCommand)
   Screen.MousePointer = vbDefault
   
   ' 檢查是否有資料傳回來
   Set GrdList.Recordset = RsTemp
   
   InitialGrdList
   
   If intI = 1 Then
      For i = 0 To GrdList.Rows - 1
         If InStr(GrdList.TextMatrix(i, 1), textQuery) > 0 Then
            GrdList.TopRow = i
            GrdList.row = i
            m_CurrSel = 0
            grdList_ShowSelection
            Exit For
         End If
      Next
    'Add By Cheng 2002/11/29
    Else
      Me.textQuery.SetFocus
      TextInverse Me.textQuery
   End If
End Sub

Private Sub buttonClear_Click()
    'Add By Cheng 2002/11/19
    Me.GrdList.Rows = 1
    
   InitialGrdList
End Sub

' 按下多筆查詢按紐
Public Sub buttonQuery_Click()
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   
   If IsDataExist(textQuery) = True Then
   
      'Added by Morgan 2021/12/22
      '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
      If PUB_CheckFormExist("frm04060101_2") = False Then
         Set frm04060101_2 = Nothing
      End If
      'end 2021/12/22
      frm04060101_2.SetMode (2)
      frm04060101_2.SetData (textQuery)
      frm04060101_2.Show
      frm04060101_2.UpdateData
      Me.Hide
   Else
      strTit = "查詢資料"
      strMsg = "資料庫無此筆資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      'Add By Cheng 2002/11/29
      Me.textQuery.SetFocus
      TextInverse Me.textQuery
   End If
End Sub

' 按下新增按紐
Private Sub buttonAdd_Click()
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   
   If IsDataExist(textQuery) = True Then
        'Modify By Cheng 2003/01/16
        '修改訊息
'      strMsg = "申請案號已存在, 請輸入其它的申請案號"
      strMsg = "申請案號已存在, 是否要修改申請案號"
      strTit = "新增資料"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      nResponse = MsgBox(strMsg, vbOKCancel + vbDefaultButton1, strTit)
        '若取消作業
        If nResponse = vbCancel Then
            'Modify By Cheng 2002/11/26
    '        'Add By Cheng 2002/11/11
    '        Me.textQuery.Text = ""
          textQuery.SetFocus
            'Add By Cheng 2002/11/26
            TextInverse Me.textQuery
          GoTo EXITSUB
        '若繼續作業
        Else
            '強制按下修改按鈕
            buttonMod_Click
            frm04060101_2.Text1 = ""
            frm04060101_2.Text1.Visible = True
            frm04060101_2.Label12.Visible = True
            Exit Sub
        End If
   End If
   If IsValidData(textQuery) = False Then
      strMsg = "請輸入正確的申請案號"
      strTit = "申請案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textQuery.SetFocus
        'Add By Cheng 2002/11/26
        TextInverse Me.textQuery
      GoTo EXITSUB
   End If
    'Modify By Cheng 2002/11/22
'   strExc(0) = "SELECT PA16,PA20 FROM PATENT WHERE PA11='" & textQuery & "' AND PA09='000'"
'   intI = 1
'   Set rsTemp = clslawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      If (rsTemp.Fields("PA16") <> "1" Or IsNull(rsTemp.Fields("PA16"))) Or IsNull(rsTemp.Fields("PA20")) Then
'         MsgBox "此申請案為本所案件，但基本檔未輸入核准資料 !", vbCritical
'         textQuery.SetFocus
'         GoTo EXITSUB
'      End If
'   End If
   
   'Added by Morgan 2021/12/22
   '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
   If PUB_CheckFormExist("frm04060101_2") = False Then
      Set frm04060101_2 = Nothing
   End If
   'end 2021/12/22
   
   frm04060101_2.SetMode (0)
   frm04060101_2.SetData (textQuery)
   frm04060101_2.Show
   frm04060101_2.UpdateData
   Me.Hide
EXITSUB:
End Sub
' 按下刪除按紐
Public Sub buttonDel_Click()
   Dim strSql As String
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   ' 檢查該筆資料是否存在
   If IsDataExist(textQuery) = True Then
      'Added by Morgan 2021/12/22
      '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
      If PUB_CheckFormExist("frm04060101_2") = False Then
         Set frm04060101_2 = Nothing
      End If
      'end 2021/12/22
      frm04060101_2.SetMode (3)
      frm04060101_2.SetData (textQuery)
      frm04060101_2.Show
      frm04060101_2.UpdateData
      Me.Hide
   Else
      strTit = "刪除資料"
      strMsg = "資料庫無此筆資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      'Add By Cheng 2002/11/29
      Me.textQuery.SetFocus
      TextInverse Me.textQuery
   End If
End Sub
' 按下變更按紐
Public Sub buttonMod_Click()
   Dim strKey As String
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   
   If IsDataExist(textQuery) = True Then
      'Added by Morgan 2021/12/22
      '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
      If PUB_CheckFormExist("frm04060101_2") = False Then
         Set frm04060101_2 = Nothing
      End If
      'end 2021/12/22
      strKey = textQuery
      frm04060101_2.SetMode (1)
      frm04060101_2.SetData (strKey)
      frm04060101_2.Show
      frm04060101_2.UpdateData
      Me.Hide
   Else
      strTit = "修改資料"
      strMsg = "資料庫無此筆資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      'Add By Cheng 2002/11/29
      Me.textQuery.SetFocus
      TextInverse Me.textQuery
   End If
End Sub
' 按下離開按紐
Private Sub buttonExit_Click()
   Unload Me
End Sub

Private Sub buttonSearch_Click()
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   ' 90.06.29 modify by louis
   If IsEmptyText(textQuery) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入申請案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   Else
      SetQueryCommand
      ExecuteQuery
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   m_CurrSel = 0
   
   InitGrid 9, GrdList
   InitialGrdList
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' 關閉 rsRecordSet 物件
   If (m_Recordset.State <> adStateClosed) Then
      m_Recordset.Close
   End If
   ' 清除物件
   Set m_Recordset = Nothing
   Unload frm04060101_2
   Unload Me
   'Add By Cheng 2002/07/16
   Set frm04060101_1 = Nothing
End Sub

Private Sub grdList_SelChange()
   If GrdList.row > 0 Then
      'grdList.Col = 1
      textQuery.Text = GrdList.TextMatrix(GrdList.row, 1)
   End If
   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = GrdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = GrdList.row Then
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < GrdList.Rows Then
      GrdList.row = m_CurrSel
      GrdList.col = 1
      If GrdList.CellBackColor <> &H80000005 Then
         For nCol = 1 To GrdList.Cols - 1
            GrdList.col = nCol
            If GrdList.CellBackColor <> &H80000005 Then: GrdList.CellBackColor = &H80000005
            If GrdList.CellForeColor <> &H80000008 Then: GrdList.CellForeColor = &H80000008
         Next nCol
      End If
      GrdList.col = 0
   End If
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < GrdList.Rows Then
      GrdList.row = m_CurrSel
      GrdList.col = 1
      For nCol = 1 To GrdList.Cols - 1
         GrdList.col = nCol
         GrdList.CellBackColor = &H8000000D
         GrdList.CellForeColor = &H80000005
      Next nCol
      GrdList.col = 0
   End If
EXITSUB:
End Sub

' 檢查此筆資料是否存在
Public Function IsDataExist(ByVal strKey As String) As Boolean
   Dim rsTmp As ADODB.Recordset
   Dim strSql As String
   
   IsDataExist = False
   strSql = "SELECT * FROM TPBulletin WHERE TPB01 = '" & strKey & "'"
   
   Set rsTmp = New ADODB.Recordset
   ' 查詢
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   
   ' 檢查是否有資料傳回來
   If rsTmp.RecordCount <= 0 Then
      IsDataExist = False
   Else
      IsDataExist = True
   End If
   
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Public Function IsValidData(ByVal strData As String)
   Dim nLength As Integer
   Dim strSysDate As String
   
   'edit by nickc 2007/09/27
   'strSysDate = ChangeWDateStringToTString(Date)
   strSysDate = strSrvDate(2)
   
   nLength = Len(strData)
   IsValidData = True
   Select Case nLength
      'Modify by Morgan 2010/12/28 申請案號改碼數
      'Case 8
         'If IsNumeric(Mid(strData, 1, 8)) = False Then
      Case 9
         If IsNumeric(strData) = False Then
         
            IsValidData = False
            GoTo EXITSUB
         End If
         'Modify by Morgan 2010/12/28 申請案號改碼數
         'If Val(Mid(strData, 3, 1)) < 1 Or Val(Mid(strData, 3, 1)) > 3 Then
         If Val(Mid(strData, 4, 1)) < 1 Or Val(Mid(strData, 4, 1)) > 3 Then
            IsValidData = False
            GoTo EXITSUB
         End If
         
      'Modify by Morgan 2010/12/28 申請案號改碼數
      'Case 11
      '   If IsNumeric(Mid(strData, 1, 8)) = False Then
      Case 12
         If IsNumeric(Mid(strData, 1, 9)) = False Then
         
            IsValidData = False
            GoTo EXITSUB
         End If
         'Modify by Morgan 2010/12/28 申請案號改碼數
         'If Val(Mid(strData, 3, 1)) < 1 Or Val(Mid(strData, 3, 1)) > 3 Then
         If Val(Mid(strData, 4, 1)) < 1 Or Val(Mid(strData, 4, 1)) > 3 Then
            IsValidData = False
            GoTo EXITSUB
         End If
         'Modify by Morgan 2010/12/28 申請案號改碼數
         'If IsNumeric(Mid(strData, 10, 2)) = False Then
         If IsNumeric(Mid(strData, 11, 2)) = False Then
            IsValidData = False
            GoTo EXITSUB
         End If
         'Modify by Morgan 2010/12/28 申請案號改碼數
         'Select Case Mid(strData, 3, 1)
         Select Case Mid(strData, 4, 1)
            Case "1", "2":
               'Modify by Morgan 2010/12/28 申請案號改碼數
               'If Mid(strData, 9, 1) <> "A" Then
               If Mid(strData, 10, 1) <> "A" Then
                  IsValidData = False
                  GoTo EXITSUB
               End If
            Case "3":
               'Modify by Morgan 2010/12/28 申請案號改碼數
               'If Mid(strData, 9, 1) <> "U" Then
               If Mid(strData, 10, 1) <> "U" Then
                  IsValidData = False
                  GoTo EXITSUB
               End If
         End Select
      Case Else
         IsValidData = False
   End Select
   
   ' 前兩碼不可大於系統年
   'Modify by Morgan 2010/12/28 申請案號改碼數
   'If Val(Left(strData, 2)) > Val(Left(strSysDate, 2)) Then
   If Val(Left(strData, 3)) > Val(strSysDate) \ 10000 Then
      IsValidData = False
   End If
   'Add by Morgan 2007/12/19
   If InStr(strData, ".") > 0 Then
      IsValidData = False
   End If
   
EXITSUB:
End Function

Private Sub textQuery_GotFocus()
   InverseAll textQuery
End Sub

Private Sub textQuery_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Public Sub SetInputTPB01()
   textQuery = Empty
   textQuery.SetFocus
End Sub

' 將所有的文字反白
Private Sub InverseAll(ByRef tb As TextBox)
   tb.SelStart = 0
   tb.SelLength = Len(tb.Text)
End Sub

' 更新列表中的資料
Public Sub UpdateRecord(ByVal strKey As String)
   Dim nIndex As Integer
   Dim Str1 As String
   Dim Str2 As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   For nIndex = 0 To GrdList.Rows - 1
      If GrdList.TextMatrix(nIndex, 1) = strKey Then
         ' 組成SQL語法
         strSql = "SELECT * FROM TPBulletin WHERE TPB01 = '" & strKey & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            ' 若在資料庫中找到該筆資料則更新此筆資料的內容
            rsTmp.MoveFirst
            ' 先清除原先的資料內容
            GrdList.TextMatrix(nIndex, 2) = Empty
            GrdList.TextMatrix(nIndex, 3) = Empty
            GrdList.TextMatrix(nIndex, 4) = Empty
            GrdList.TextMatrix(nIndex, 5) = Empty
            GrdList.TextMatrix(nIndex, 6) = Empty
            GrdList.TextMatrix(nIndex, 7) = Empty
            GrdList.TextMatrix(nIndex, 8) = Empty
            
            If IsNull(rsTmp.Fields("TPB02")) = False Then
               GrdList.TextMatrix(nIndex, 2) = rsTmp.Fields("TPB02")
            End If
            If IsNull(rsTmp.Fields("TPB03")) = False Then
                'Modify By Cheng 2003/06/24
'               grdList.TextMatrix(nIndex, 3) = ChangeWStringToTString(rsTmp.Fields("TPB03"))
               GrdList.TextMatrix(nIndex, 3) = ChangeTStringToTDateString(ChangeWStringToTString(rsTmp.Fields("TPB03")))
            End If
            Str1 = Empty
            Str2 = Empty
            If IsNull(rsTmp.Fields("TPB04")) = False Then
               Str1 = rsTmp.Fields("TPB04")
            End If
            If IsNull(rsTmp.Fields("TPB05")) = False Then
               Str2 = rsTmp.Fields("TPB05")
            End If
            If Str1 = Empty Then: Str1 = "  "
            If Str2 = Empty Then: Str2 = "  "
            'Modify By Cheng 2003/06/24
'            grdList.TextMatrix(nIndex, 4) = Str1 & " - " & Str2
            GrdList.TextMatrix(nIndex, 4) = Str1 & "-" & Str2
            
            If IsNull(rsTmp.Fields("TPB06")) = False Then
               GrdList.TextMatrix(nIndex, 5) = GetNation(rsTmp.Fields("TPB06"))
            End If
            ' 代理人
            If IsNull(rsTmp.Fields("TPB07")) = False Then
               GrdList.TextMatrix(nIndex, 6) = GetAgentName(rsTmp.Fields("TPB07"))
            End If
            ' 本所案號
            If IsNull(rsTmp.Fields("TPB01")) = False Then
               GrdList.TextMatrix(nIndex, 7) = GetPNumber(rsTmp.Fields("TPB01"))
            End If
            ' 事務所名稱
            If IsNull(rsTmp.Fields("TPB08")) = False Then
               GrdList.TextMatrix(nIndex, 8) = rsTmp.Fields("TPB08")
            End If
         Else
            ' 資料庫中無該筆資料表示此筆已被刪除
            GrdList.RemoveItem (nIndex)
         End If
         Exit For
      End If
   Next nIndex
End Sub

