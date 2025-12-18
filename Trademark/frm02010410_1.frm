VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010410_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "廣告刊出來函輸入"
   ClientHeight    =   4896
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   7896
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4896
   ScaleWidth      =   7896
   Begin VB.TextBox textSel 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1056
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1872
      Width           =   1104
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7020
      TabIndex        =   5
      Top             =   48
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   6192
      TabIndex        =   4
      Top             =   48
      Width           =   800
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "尋找(&Q)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5328
      TabIndex        =   3
      Top             =   48
      Width           =   840
   End
   Begin VB.TextBox textCP27_2 
      Height          =   264
      Left            =   2736
      TabIndex        =   2
      Top             =   1512
      Width           =   1116
   End
   Begin VB.TextBox textCP27_1 
      Height          =   264
      Left            =   1032
      TabIndex        =   1
      Top             =   1512
      Width           =   1116
   End
   Begin VB.TextBox textCU05 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   2184
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   840
      Width           =   5640
   End
   Begin VB.TextBox textTM23 
      Height          =   264
      Left            =   1032
      TabIndex        =   0
      Top             =   552
      Width           =   1116
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   2664
      Left            =   72
      TabIndex        =   13
      Top             =   2184
      Width           =   7788
      _ExtentX        =   13737
      _ExtentY        =   4699
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
   Begin MSForms.TextBox textCU06 
      Height          =   270
      Left            =   2190
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1140
      Width           =   5640
      VariousPropertyBits=   679493663
      Size            =   "9948;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCU04 
      Height          =   264
      Left            =   2184
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   552
      Width           =   5640
      VariousPropertyBits=   679493663
      Size            =   "9948;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "選取筆數 :"
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   1872
      Width           =   852
   End
   Begin VB.Line Line1 
      X1              =   2256
      X2              =   2616
      Y1              =   1632
      Y2              =   1632
   End
   Begin VB.Label Label1 
      Caption         =   "發文日期 :"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1536
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "申請人 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   552
      Width           =   852
   End
End
Attribute VB_Name = "frm02010410_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/18 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Amy 2021/12/29 Form2.0已修改 textCU04/textCU06/grdList
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bSel As Boolean
   Dim strCP01 As String
   Dim strCP02 As String
   Dim strCP03 As String
   Dim strCP04 As String
   Dim strCP09 As String
   Dim strCP12 As String
   Dim strCP13 As String
   
   bSel = False
   For nIndex = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nIndex, 0) = "V" Then
         bSel = True
         Exit For
      End If
   Next nIndex
   If Not bSel Then
      strTit = "廣告刊出來函輸入"
      strMsg = "請先選取資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   ' 顯示下一畫面
   frm02010410_2.SetData textTM23, "", "", "", "", "", "", 0, True
   frm02010410_2.SetData textCP27_1, "", "", "", "", "", "", 1, False
   frm02010410_2.SetData textCP27_2, "", "", "", "", "", "", 2, False
   frm02010410_2.SetData textSel, "", "", "", "", "", "", 3, False
   For nIndex = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nIndex, 0) = "V" Then
         strCP01 = grdList.TextMatrix(nIndex, 8)
         strCP02 = grdList.TextMatrix(nIndex, 9)
         strCP03 = grdList.TextMatrix(nIndex, 10)
         strCP04 = grdList.TextMatrix(nIndex, 11)
         strCP09 = grdList.TextMatrix(nIndex, 4)
         strCP12 = grdList.TextMatrix(nIndex, 12)
         strCP13 = grdList.TextMatrix(nIndex, 13)
         frm02010410_2.SetData strCP01, strCP02, strCP03, strCP04, strCP09, strCP12, strCP13, 4, False
      End If
   Next nIndex
   frm02010410_2.Show
   Me.Hide
         
EXITSUB:
End Sub

Private Sub cmdQuery_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   ClearGrdList
   If CheckDataValid() Then
      If QueryData() = False Then
         strTit = "查尋資料"
         strMsg = "沒有符合條件的資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
End Sub

Private Function ClearGrdList()
   grdList.Rows = 1
End Function

Private Function QueryData() As Boolean
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim nRow As Integer
   Set rsTmp = New ADODB.Recordset
   QueryData = False
   '911015 nickc 邱小姐說 cp24 要是 null 才出來
   'strSQL = "SELECT CP01,CP02,CP03,CP04,CP05,CP09,CP12,CP13,CP27,CP64,TM23,ST02 FROM CaseProgress, Trademark, STAFF " & _
            "WHERE CP27 >= " & DBDATE(textCP27_1) & " AND " & _
                  "CP27 <= " & DBDATE(textCP27_2) & " AND " & _
                  "CP13 = ST01 (+) AND " & _
                  "CP01 = TM01 (+) AND " & _
                  "CP02 = TM02 (+) AND " & _
                  "CP03 = TM03 (+) AND " & _
                  "CP04 = TM04 (+) AND " & _
                  "TM23 = '" & textTM23 & "' " & _
            "ORDER BY CP13, CP27, CP01, CP02, CP03, CP04 ASC "
   'Modify By Sindy 2011/2/16 因用SQLDate排序或取MAX或MIN,修改百年蟲問題
'   strSql = "SELECT CP01,CP02,CP03,CP04," & SQLDate("CP05") & " as cp05,CP09,CP12,CP13," & SQLDate("CP27") & " as cp27,CP64,TM23,ST02,nvl(tm05,nvl(tm06,tm07)) as tmName FROM CaseProgress, Trademark, STAFF " & _
'            "WHERE CP27 >= " & DBDATE(textCP27_1) & " AND " & _
'                  "CP27 <= " & DBDATE(textCP27_2) & " AND " & _
'                  "CP13 = ST01 (+) AND " & _
'                  "CP01 = TM01 (+) AND " & _
'                  "CP02 = TM02 (+) AND " & _
'                  "CP03 = TM03 (+) AND " & _
'                  "CP04 = TM04 (+) AND " & _
'                  "TM23 = '" & textTM23 & "' and cp24 is null " & _
'            "ORDER BY CP13, CP27, CP01, CP02, CP03, CP04 ASC "
   strSql = "SELECT CP01,CP02,CP03,CP04,CP05,CP09,CP12,CP13,CP27,CP64,TM23,ST02,nvl(tm05,nvl(tm06,tm07)) as tmName FROM CaseProgress, Trademark, STAFF " & _
            "WHERE CP27 >= " & DBDATE(textCP27_1) & " AND " & _
                  "CP27 <= " & DBDATE(textCP27_2) & " AND " & _
                  "CP13 = ST01 (+) AND " & _
                  "CP01 = TM01 (+) AND " & _
                  "CP02 = TM02 (+) AND " & _
                  "CP03 = TM03 (+) AND " & _
                  "CP04 = TM04 (+) AND " & _
                  "TM23 = '" & textTM23 & "' and cp24 is null " & _
            "ORDER BY CP13, CP27, CP01, CP02, CP03, CP04 ASC "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      QueryData = True
      Do While rsTmp.EOF = False
         grdList.Rows = grdList.Rows + 1
         nRow = grdList.Rows - 1
         ' 發文日
         If Not IsNull(rsTmp.Fields("CP27")) Then
            'Modify By Sindy 2011/2/16 因用SQLDate排序或取MAX或MIN,修改百年蟲問題
            'grdList.TextMatrix(nRow, 1) = rsTmp.Fields("CP27")
            grdList.TextMatrix(nRow, 1) = ChangeWStringToTDateString(rsTmp.Fields("CP27"))
            '911015 nick
            grdList.row = nRow
            grdList.col = 1
            grdList.CellAlignment = flexAlignLeftBottom
         End If
         ' 本所案號
         If Not IsNull(rsTmp.Fields("CP01")) Then
            grdList.TextMatrix(nRow, 2) = rsTmp.Fields("CP01") & "-" & rsTmp.Fields("CP02") & "-" & rsTmp.Fields("CP03") & "-" & rsTmp.Fields("CP04")
         End If
         '911015 nick 新增   案件名稱
         If Not IsNull(rsTmp.Fields("tmName")) Then
            grdList.TextMatrix(nRow, 3) = rsTmp.Fields("tmName")
         End If
         ' 總收文號
         If Not IsNull(rsTmp.Fields("CP09")) Then
            grdList.TextMatrix(nRow, 4) = rsTmp.Fields("CP09")
         End If
         ' 收文日
         If Not IsNull(rsTmp.Fields("CP05")) Then
            'Modify By Sindy 2011/2/16 因用SQLDate排序或取MAX或MIN,修改百年蟲問題
            'grdList.TextMatrix(nRow, 5) = rsTmp.Fields("CP05")
            grdList.TextMatrix(nRow, 5) = ChangeWStringToTDateString(rsTmp.Fields("CP05"))
            '911015 nick
            grdList.row = nRow
            grdList.col = 5
            grdList.CellAlignment = flexAlignLeftBottom
         End If
         ' 智權人員
         If Not IsNull(rsTmp.Fields("ST02")) Then
            grdList.TextMatrix(nRow, 6) = rsTmp.Fields("ST02")
         End If
         ' 進度備註
         If Not IsNull(rsTmp.Fields("CP64")) Then
            grdList.TextMatrix(nRow, 7) = rsTmp.Fields("CP64")
         End If
         ' 本所案號
         If Not IsNull(rsTmp.Fields("CP01")) Then
            grdList.TextMatrix(nRow, 8) = rsTmp.Fields("CP01")
         End If
         If Not IsNull(rsTmp.Fields("CP02")) Then
            grdList.TextMatrix(nRow, 9) = rsTmp.Fields("CP02")
         End If
         If Not IsNull(rsTmp.Fields("CP03")) Then
            grdList.TextMatrix(nRow, 10) = rsTmp.Fields("CP03")
         End If
         If Not IsNull(rsTmp.Fields("CP04")) Then
            grdList.TextMatrix(nRow, 11) = rsTmp.Fields("CP04")
         End If
         ' 業務區別
         If Not IsNull(rsTmp.Fields("CP12")) Then
            grdList.TextMatrix(nRow, 12) = rsTmp.Fields("CP12")
         End If
         ' 智權人員代號
         If Not IsNull(rsTmp.Fields("CP13")) Then
            grdList.TextMatrix(nRow, 13) = rsTmp.Fields("CP13")
         End If
         rsTmp.MoveNext
      Loop
      'Added by Lydia 2023/10/18
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/18
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Sub Form_Load()
   textCU04.BackColor = &H8000000F
   textCU05.BackColor = &H8000000F
   textCU06.BackColor = &H8000000F
   textSel.BackColor = &H8000000F
  
   textSel = "0"
   
   MoveFormToCenter Me
   
   InitialGrdList
End Sub

' 初始化 GridList
Public Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 14
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "發文日"
   grdList.ColWidth(1) = 1000
   '911015 nick 向左靠
   grdList.CellAlignment = flexAlignLeftCenter
   grdList.col = 2
   grdList.Text = "本所案號"
   grdList.ColWidth(2) = 1600
   '911015 nick 新增
   grdList.col = 3
   grdList.Text = "案件名稱"
   grdList.ColWidth(3) = 1600
   grdList.col = 4
   grdList.Text = "總收文號"
   grdList.ColWidth(4) = 1000
   grdList.col = 5
   grdList.Text = "收文日"
   grdList.ColWidth(5) = 1000
   '911015 nick 向左靠
   grdList.CellAlignment = flexAlignLeftCenter
   grdList.col = 6
   grdList.Text = "智權人員"
   grdList.ColWidth(6) = 1200
   grdList.col = 7
   grdList.Text = "進度備註"
   grdList.ColWidth(7) = 3000
   grdList.col = 8
   grdList.Text = "本所案號"
   grdList.ColWidth(8) = 0
   grdList.col = 9
   grdList.Text = "本所案號"
   grdList.ColWidth(9) = 0
   grdList.col = 10
   grdList.Text = "本所案號"
   grdList.ColWidth(10) = 0
   grdList.col = 11
   grdList.Text = "本所案號"
   grdList.ColWidth(11) = 0
   grdList.col = 12
   grdList.Text = "業務區別"
   grdList.ColWidth(12) = 0
   grdList.col = 13
   grdList.Text = "智權人員代號"
   grdList.ColWidth(13) = 0
End Sub

Public Function UpdateCustomerName(ByVal strCustomer) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   UpdateCustomerName = False
   
   If Len(strCustomer) < 9 Then: strCustomer = strCustomer & String(9 - Len(strCustomer), "0")
   
   If Len(strCustomer) > 8 Then
      strSql = "SELECT * FROM Customer " & _
               "WHERE CU01 = '" & Mid(strCustomer, 1, 8) & "' AND " & _
                     "CU02 = '" & Mid(strCustomer, 9, 1) & "'"
   Else
      strSql = "SELECT * FROM Customer " & _
               "WHERE CU01 = '" & Mid(strCustomer, 1, 8) & "' AND " & _
                     "CU02 = '0' "
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      UpdateCustomerName = True
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("CU04")) = False Then
         textCU04 = rsTmp.Fields("CU04")
      End If
      If IsNull(rsTmp.Fields("CU05")) = False Then
         textCU05 = rsTmp.Fields("CU05")
      ElseIf IsNull(rsTmp.Fields("CU88")) = False Then
         textCU05 = rsTmp.Fields("CU88")
      ElseIf IsNull(rsTmp.Fields("CU89")) = False Then
         textCU05 = rsTmp.Fields("CU89")
      ElseIf IsNull(rsTmp.Fields("CU90")) = False Then
         textCU05 = rsTmp.Fields("CU90")
      End If
      If IsNull(rsTmp.Fields("CU06")) = False Then
         textCU06 = rsTmp.Fields("CU06")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Sub CountSelect()
   Dim nIndex As Integer
   Dim nSel As Integer
   nSel = 0
   For nIndex = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nIndex, 0) = "V" Then
         nSel = nSel + 1
      End If
   Next nIndex
   textSel = CStr(nSel)
End Sub

Private Sub grdList_Click()
   If grdList.row > 0 And grdList.row < grdList.Rows Then
      If grdList.TextMatrix(grdList.row, 0) = "V" Then
         grdList.TextMatrix(grdList.row, 0) = Empty
      Else
         grdList.TextMatrix(grdList.row, 0) = "V"
         cmdOK.Default = True
      End If
   End If
   CountSelect
End Sub

Private Sub textCP27_1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP27_1) = False Then
      If CheckIsTaiwanDate(textCP27_1, False) = False Then
         Cancel = True
         strMsg = "請輸入正確的發文日起"
         strTit = "發文日起"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_1_GotFocus
         GoTo EXITSUB
      End If
      If Val(DBDATE(textCP27_1)) >= Val(DBDATE(SystemDate())) Then
         Cancel = True
         strMsg = "發文日不可超過系統日"
         strTit = "發文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_1_GotFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

Private Sub textCP27_2_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP27_2) = False Then
      If CheckIsTaiwanDate(textCP27_2, False) = False Then
         Cancel = True
         strMsg = "請輸入正確的發文日起"
         strTit = "發文日起"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_2_GotFocus
         GoTo EXITSUB
      End If
      If Val(DBDATE(textCP27_2)) >= Val(DBDATE(SystemDate())) Then
         Cancel = True
         strMsg = "發文日不可超過系統日"
         strTit = "發文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_2_GotFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

Private Sub textTM23_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCU04 = Empty
   textCU05 = Empty
   textCU06 = Empty
   If IsEmptyText(textTM23) = False Then
      textTM23 = textTM23 & String(9 - Len(textTM23), "0")
      If UpdateCustomerName(textTM23) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textTM23 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
   If Cancel = True Then textTM23_GotFocus
End Sub

Private Function CheckDataValid()
   Dim strDay As String
   Dim strTemp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   ' 申請人不可為空白
   If IsEmptyText(textTM23) = True Then
      strTit = "檢核資料"
      strMsg = "申請人不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM23.SetFocus
      GoTo EXITSUB
   End If
   ' 發文日不可為空白
   If IsEmptyText(textCP27_1) = True Then
      strTit = "檢核資料"
      strMsg = "發文日不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP27_1.SetFocus
      GoTo EXITSUB
   End If
   If IsEmptyText(textCP27_2) = True Then
      strTit = "檢核資料"
      strMsg = "發文日不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP27_2.SetFocus
      GoTo EXITSUB
   End If
   ' 發文日範圍不正確
   If Val(DBDATE(textCP27_1)) > Val(DBDATE(textCP27_2)) Then
      strTit = "檢核資料"
      strMsg = "發文日範圍不正確"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP27_1.SetFocus
      GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub grdList_SelChange()
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

Private Sub textTM23_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM23_GotFocus()
   InverseTextBox textTM23
End Sub

Private Sub textCP27_1_GotFocus()
   InverseTextBox textCP27_1
End Sub

Private Sub textCP27_2_GotFocus()
   InverseTextBox textCP27_2
End Sub

