VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm02010403_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "審查報告輸入"
   ClientHeight    =   5736
   ClientLeft      =   180
   ClientTop       =   996
   ClientWidth     =   9336
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   9336
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7164
      TabIndex        =   1
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6336
      TabIndex        =   0
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8388
      TabIndex        =   2
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   6180
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   2172
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   2172
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   4512
      Left            =   144
      TabIndex        =   7
      Top             =   1080
      Width           =   9072
      _ExtentX        =   16002
      _ExtentY        =   7959
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
   Begin VB.Label Label2 
      Caption         =   "申請案號 :"
      Height          =   252
      Left            =   4980
      TabIndex        =   6
      Top             =   720
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "審定號數 :"
      Height          =   252
      Left            =   180
      TabIndex        =   5
      Top             =   720
      Width           =   1092
   End
End
Attribute VB_Name = "frm02010403_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/18 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Amy 2021/12/28 Form2.0已修改 grdList
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 申請案號
Dim m_TM12 As String
' 審定號
Dim m_TM15 As String
' 案件進度檔的本所案號
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
' 來函收文日
Dim m_CP05 As String
'Add By Sindy 2019/5/10
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2019/5/10 END


'Add By Sindy 2019/5/13
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   frm02010403_1.Show
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm02010403_1
   Unload Me
End Sub
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
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_CP01 & "' AND " & _
                  "TM02 = '" & m_CP02 & "' AND " & _
                  "TM03 = '" & m_CP03 & "' AND " & _
                  "TM04 = '" & m_CP04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("TM10")) = False Then
         strNation = rsTmp.Fields("TM10")
      End If
   End If
   rsTmp.Close
   
   If strNation < "010" Then
      strSql = "SELECT * FROM MailRec " & _
               "WHERE MR12 = '" & m_CP01 & "' AND " & _
                     "MR13 = '" & m_CP02 & "' AND " & _
                     "MR14 = '" & m_CP03 & "' AND " & _
                     "MR15 = '" & m_CP04 & "' AND " & _
                     "MR02 = " & ChangeTStringToWString(m_CP05) & " "
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
   
   If grdList.Rows > 0 Then
      If IsEmptyText(m_CP01) = True Or IsEmptyText(m_CP02) = True Or IsEmptyText(m_CP03) = True Or IsEmptyText(m_CP04) = True Then
         strMsg = "請先選取一筆記錄"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   Else
      strMsg = "無符合的資料"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   ' 90.07.03 modify by louis (該畫面不檢查來函記錄檔)
   'If PromptIfTaiwanNoResult = True Then
      DisplayNextForm
   'End If

EXITSUB:
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTM15.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   
   MoveFormToCenter Me
   
   'Add By Sindy 2019/5/10
   m_strIR01 = frm02010403_1.m_strIR01
   m_strIR02 = frm02010403_1.m_strIR02
   m_strIR03 = frm02010403_1.m_strIR03
   m_strIR04 = frm02010403_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2019/5/10 END
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_TM12 = Empty
      m_TM15 = Empty
      m_CP05 = Empty
   End If
   
   Select Case nType
      ' 本所案號 欄位1
      Case 0: m_TM01 = strData
      ' 本所案號 欄位2
      Case 1: m_TM02 = strData
      ' 本所案號 欄位3
      Case 2: m_TM03 = strData
      ' 本所案號 欄位4
      Case 3: m_TM04 = strData
      ' 申請案號
      Case 4: m_TM12 = strData
      ' 審定號
      Case 5: m_TM15 = strData
      ' 來函收文日
      Case 6: m_CP05 = strData
   End Select
End Sub

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 11
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "本所案號"
   grdList.ColWidth(1) = 1200
   grdList.col = 2
   grdList.Text = "商標名稱"
   grdList.ColWidth(2) = 1000
   grdList.col = 3
   grdList.Text = "商標種類"
   grdList.ColWidth(3) = 1200
   grdList.col = 4
   grdList.Text = "商品類別"
   grdList.ColWidth(4) = 1000
   grdList.col = 5
   grdList.Text = "申請人"
   grdList.ColWidth(5) = 1200
   grdList.col = 6
   grdList.Text = "申請國家"
   grdList.ColWidth(6) = 1200
   ' 本所案號 欄位一
   grdList.col = 7
   grdList.Text = Empty
   grdList.ColWidth(7) = 0
   ' 本所案號 欄位二
   grdList.col = 8
   grdList.Text = Empty
   grdList.ColWidth(8) = 0
   ' 本所案號 欄位三
   grdList.col = 9
   grdList.Text = Empty
   grdList.ColWidth(9) = 0
   ' 本所案號 欄位四
   grdList.col = 10
   grdList.Text = Empty
   grdList.ColWidth(10) = 0
End Sub

' 顯示下一個畫面
Private Sub DisplayNextForm()
   frm02010403_3.SetData 0, m_CP01, True
   frm02010403_3.SetData 1, m_CP02, False
   frm02010403_3.SetData 2, m_CP03, False
   frm02010403_3.SetData 3, m_CP04, False
   frm02010403_3.SetData 4, m_CP05, False
   If grdList.Rows > 2 Then
      frm02010403_3.SetData 5, "2", False
   Else
      frm02010403_3.SetData 5, "1", False
   End If
   'Add By Sindy 2019/5/10
   If Not m_PrevForm Is Nothing Then
      Call frm02010403_3.SetParent(m_PrevForm)
   End If
   frm02010403_3.m_strIR01 = m_strIR01
   frm02010403_3.m_strIR02 = m_strIR02
   frm02010403_3.m_strIR03 = m_strIR03
   frm02010403_3.m_strIR04 = m_strIR04
   '2019/5/10 END
   Me.Hide
   frm02010403_3.Show
   frm02010403_3.QueryData
End Sub

Private Sub ListDataFromSP(ByRef rsTmp As ADODB.Recordset)
   InitialGrdList
   rsTmp.MoveFirst
   Do While rsTmp.EOF = False
      grdList.Rows = grdList.Rows + 1
      grdList.row = grdList.Rows - 1
      ' 本所案號欄位
      grdList.TextMatrix(grdList.row, 1) = rsTmp.Fields("SP01") & rsTmp.Fields("SP02") & rsTmp.Fields("SP03") & rsTmp.Fields("SP04")
      ' 商標名稱欄位
      If IsNull(rsTmp.Fields("SP05")) = False Then
         grdList.TextMatrix(grdList.row, 2) = rsTmp.Fields("SP05")
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         grdList.TextMatrix(grdList.row, 6) = GetNationName(rsTmp.Fields("SP09"), 0)
      End If
      ' 商標種類欄位
      grdList.TextMatrix(grdList.row, 3) = Empty
      ' 商品類別
      grdList.TextMatrix(grdList.row, 4) = Empty
      ' 申請人
      If IsNull(rsTmp.Fields("SP08")) = False Then
         grdList.TextMatrix(grdList.row, 5) = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
      ' 隱藏起來的本所案號
      grdList.TextMatrix(grdList.row, 7) = rsTmp.Fields("SP01")
      grdList.TextMatrix(grdList.row, 8) = rsTmp.Fields("SP02")
      grdList.TextMatrix(grdList.row, 9) = rsTmp.Fields("SP03")
      grdList.TextMatrix(grdList.row, 10) = rsTmp.Fields("SP04")
      
      rsTmp.MoveNext
   Loop
   
   'Added by Lydia 2023/10/18
   If grdList.Rows >= 2 Then
      grdList.FixedRows = 1
   End If
   'end 2023/10/18
      
   ' 設定第一列為選取的狀態
   grdList_SetSelection 1
      
   ' 當只有一筆記錄時, 則直接跳至下一畫面
   If grdList.Rows = 2 Then
      ' 90.07.03 modify by louis (該畫面不檢查來函記錄檔)
      'If PromptIfTaiwanNoResult = True Then
         DisplayNextForm
      'Else
      '   Unload Me
      '   frm02010403_1.Show
      'End If
   Else
      Me.Show
   End If
End Sub

' 列出所有資料
Private Sub ListDataFromTM(ByRef rsTmp As ADODB.Recordset)
   Dim strNationCode As String
   InitialGrdList
   rsTmp.MoveFirst
   Do While rsTmp.EOF = False
      strNationCode = Empty
      grdList.Rows = grdList.Rows + 1
      grdList.row = grdList.Rows - 1
      ' 本所案號欄位
      grdList.TextMatrix(grdList.row, 1) = rsTmp.Fields("TM01") & rsTmp.Fields("TM02") & rsTmp.Fields("TM03") & rsTmp.Fields("TM04")
      ' 商標名稱欄位
      If IsNull(rsTmp.Fields("TM05")) = False Then
         grdList.TextMatrix(grdList.row, 2) = rsTmp.Fields("TM05")
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         strNationCode = rsTmp.Fields("TM10")
         grdList.TextMatrix(grdList.row, 6) = GetNationName(rsTmp.Fields("TM10"), 0)
      End If
      ' 商標種類欄位
      If IsNull(rsTmp.Fields("TM08")) = False Then
         If strNationCode < "010" Then
            grdList.TextMatrix(grdList.row, 3) = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
         Else
            grdList.TextMatrix(grdList.row, 3) = GetTradeMarkName(rsTmp.Fields("TM08"), 1)
         End If
      End If
      ' 商品類別
      If IsNull(rsTmp.Fields("TM09")) = False Then
         grdList.TextMatrix(grdList.row, 4) = rsTmp.Fields("TM09")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         grdList.TextMatrix(grdList.row, 5) = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      ' 隱藏起來的本所案號
      grdList.TextMatrix(grdList.row, 7) = rsTmp.Fields("TM01")
      grdList.TextMatrix(grdList.row, 8) = rsTmp.Fields("TM02")
      grdList.TextMatrix(grdList.row, 9) = rsTmp.Fields("TM03")
      grdList.TextMatrix(grdList.row, 10) = rsTmp.Fields("TM04")
               
      rsTmp.MoveNext
   Loop
      
   'Added by Lydia 2023/10/18
   If grdList.Rows >= 2 Then
      grdList.FixedRows = 1
   End If
   'end 2023/10/18
      
   textTM15 = m_TM15
   textTM12 = m_TM12
   
   ' 設定第一列為選取的狀態
   grdList_SetSelection 1
      
   ' 當只有一筆記錄時, 則直接跳至下一畫面
   If grdList.Rows = 2 Then
      'If PromptIfTaiwanNoResult = True Then
         DisplayNextForm
      'Else
      '   Unload Me
      '   frm02010403_1.Show
      'End If
   Else
      Me.Show
   End If
End Sub
' 搜尋資料
Public Sub QueryData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim nType As Integer
   Dim strTM01 As String
   Dim strTM02 As String
   Dim strTM03 As String
   Dim strTM04 As String
   
   ' 判斷所選擇的查詢方式
   nType = 0
   If IsEmptyText(m_TM01) = False Then
      nType = 3
   ElseIf IsEmptyText(m_TM12) = False Then
      nType = 1
   ElseIf IsEmptyText(m_TM15) = False Then
      nType = 2
   End If
   
   Select Case nType
      Case 1:
         ' 設定SQL語法
         strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM12 = '" & m_TM12 & "' AND " & _
                        "TM10 < '010' " & _
                  "ORDER BY TM01||TM02||TM03||TM04 "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
         If rsTmp.RecordCount > 0 Then
            ListDataFromTM rsTmp
         End If
         rsTmp.Close
      Case 2:
         ' 設定SQL語法
         strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM15 = '" & m_TM15 & "' AND " & _
                        "TM10 < '010' " & _
                  "ORDER BY TM01||TM02||TM03||TM04 "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
         If rsTmp.RecordCount > 0 Then
            ListDataFromTM rsTmp
         End If
         rsTmp.Close
      Case 3:
         strTM01 = m_TM01
         strTM02 = m_TM02
         strTM03 = m_TM03
         strTM04 = m_TM04
         
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
                              "TM04 = '" & strTM04 & "' " & _
                        "ORDER BY TM01||TM02||TM03||TM04 "
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
               If rsTmp.RecordCount > 0 Then
                  ListDataFromTM rsTmp
               End If
               rsTmp.Close
            Case "TF":
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
               strSql = strSql & "ORDER BY TM01||TM02||TM03||TM04 "
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
               If rsTmp.RecordCount > 0 Then
                  ListDataFromTM rsTmp
               End If
               rsTmp.Close
            Case Else:
               If IsEmptyText(strTM03) = True Then
                  strTM03 = "0"
               End If
               If IsEmptyText(strTM04) = True Then
                  strTM04 = "00"
               End If
               ' 設定SQL語法
               strSql = "SELECT * FROM ServicePractice " & _
                        "WHERE SP01 = '" & strTM01 & "' AND " & _
                              "SP02 = '" & strTM02 & "' AND " & _
                              "SP03 = '" & strTM03 & "' AND " & _
                              "SP04 = '" & strTM04 & "' " & _
                        "ORDER BY SP01||SP02||SP03||SP04 "
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
               If rsTmp.RecordCount > 0 Then
                  ListDataFromSP rsTmp
               End If
               rsTmp.Close
         End Select
   End Select
   Set rsTmp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
'edit by nickc 2008/04/25 改整批印
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   
   'Add By Sindy 2019/5/13
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   
   'Add By Cheng 2002/07/18
   Set frm02010403_2 = Nothing
End Sub

Private Sub grdList_SelChange()
   If grdList.row > 0 Then
      grdList.col = 7
      m_CP01 = grdList.Text
      grdList.col = 8
      m_CP02 = grdList.Text
      grdList.col = 9
      m_CP03 = grdList.Text
      grdList.col = 10
      m_CP04 = grdList.Text
   End If
   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nRow As Integer
   Dim nCol As Integer
   Dim nCurrSel As Integer
   nCurrSel = grdList.row
   For nRow = 1 To grdList.Rows - 1
      grdList.row = nRow
      If nRow = nCurrSel Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            grdList.CellBackColor = &H8000000D
            grdList.CellForeColor = &H80000005
         Next nCol
      Else
         grdList.col = 1
         If grdList.CellBackColor <> &H80000005 Then
            For nCol = 1 To grdList.Cols - 1
               grdList.col = nCol
               If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
               If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
            Next nCol
         End If
      End If
   Next nRow
   grdList.row = nCurrSel
   grdList.col = 0
End Sub

' 設定Grid List的一列為選取的狀態
Private Sub grdList_SetSelection(ByVal nSel As Integer)
   If grdList.Rows >= 2 Then
      grdList.row = nSel
      grdList_SelChange
      grdList_ShowSelection
   End If
End Sub


