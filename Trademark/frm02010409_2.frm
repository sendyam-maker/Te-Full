VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010409_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "服務業務結果輸入"
   ClientHeight    =   5670
   ClientLeft      =   -160
   ClientTop       =   1000
   ClientWidth     =   9340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   9340
   Begin VB.TextBox textSP06 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1320
      Width           =   7752
   End
   Begin VB.TextBox textSPKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   2532
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7152
      TabIndex        =   1
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6324
      TabIndex        =   0
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8376
      TabIndex        =   2
      Top             =   70
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   3432
      Left            =   96
      TabIndex        =   13
      Top             =   2208
      Width           =   9132
      _ExtentX        =   16104
      _ExtentY        =   6050
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
   Begin MSForms.TextBox textSP08 
      Height          =   264
      Left            =   1440
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1920
      Width           =   7752
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "13674;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textSP07 
      Height          =   264
      Left            =   1440
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1620
      Width           =   7752
      VariousPropertyBits=   679493663
      Size            =   "13674;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textSP05 
      Height          =   264
      Left            =   1440
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1020
      Width           =   7752
      VariousPropertyBits=   679493663
      Size            =   "13674;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   852
   End
   Begin VB.Label Label5 
      Caption         =   "案件日文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   11
      Top             =   1620
      Width           =   1452
   End
   Begin VB.Label Label4 
      Caption         =   "案件英文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   1212
   End
   Begin VB.Label Label3 
      Caption         =   "案件中文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   9
      Top             =   1020
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   852
   End
End
Attribute VB_Name = "frm02010409_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/18 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Amy 2021/12/29 Form2.0已修改 textSP08/textSP07/textSP05/grdList
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit

' 本所案號
Dim m_SP01 As String
Dim m_SP02 As String
Dim m_SP03 As String
Dim m_SP04 As String
' 來函收文日
Dim m_CP05 As String
' 所選取的收文號
Dim m_CP09 As String
' 申請國家
Dim m_SP09 As String
' 所選取的案件性質
Dim m_CP10 As String
'Add By Sindy 2019/5/22
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2019/5/22 END


'Add By Sindy 2019/5/22
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdCancel_Click()
   frm02010409_1.Show
   Unload Me
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm02010409_1
   Unload Me
End Sub

Private Sub cmdOK_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If CheckDataValid() = True Then
      ' 90.07.03 modify by louis (檢查來函記錄檔)
      'If PromptIfTaiwanNoResult() = True Then
         DisplayNextForm
      'End If
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textSPKey.BackColor = &H8000000F
   textSP05.BackColor = &H8000000F
   textSP06.BackColor = &H8000000F
   textSP07.BackColor = &H8000000F
   textSP08.BackColor = &H8000000F
   
   MoveFormToCenter Me
   
   'Add By Sindy 2019/5/22
   m_strIR01 = frm02010409_1.m_strIR01
   m_strIR02 = frm02010409_1.m_strIR02
   m_strIR03 = frm02010409_1.m_strIR03
   m_strIR04 = frm02010409_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2019/5/22 END
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
   strSql = "SELECT * FROM SERVICEPRACTICE " & _
            "WHERE SP01 = '" & m_SP01 & "' AND " & _
                  "SP02 = '" & m_SP02 & "' AND " & _
                  "SP03 = '" & m_SP03 & "' AND " & _
                  "SP04 = '" & m_SP04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("SP09")) = False Then
         strNation = rsTmp.Fields("SP09")
      End If
   End If
   rsTmp.Close
   
   If strNation < "010" Then
      strSql = "SELECT * FROM MailRec " & _
               "WHERE MR12 = '" & m_SP01 & "' AND " & _
                     "MR13 = '" & m_SP02 & "' AND " & _
                     "MR14 = '" & m_SP03 & "' AND " & _
                     "MR15 = '" & m_SP04 & "' AND " & _
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

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_SP01 = Empty
      m_SP02 = Empty
      m_SP03 = Empty
      m_SP04 = Empty
      m_CP05 = Empty
      m_SP09 = Empty
   End If
   
   Select Case nType
      ' 本所案號 欄位1
      Case 0: m_SP01 = strData
      ' 本所案號 欄位2
      Case 1: m_SP02 = strData
      ' 本所案號 欄位3
      Case 2: m_SP03 = strData
      ' 本所案號 欄位4
      Case 3: m_SP04 = strData
      ' 來函收文日
      Case 4: m_CP05 = strData
   End Select
End Sub

' 由客戶代碼取得客戶名稱
Private Function GetCustomer(ByVal strData As String) As String
   Dim rsTmp As ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   GetCustomer = Empty
   If IsEmptyText(strData) = False Then
      Set rsTmp = New ADODB.Recordset
      If Len(strData) > 8 Then
         strSql = "SELECT * FROM Customer " & _
                  "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                        "CU02 = '" & Mid(strData, 9, 1) & "'"
      Else
         strSql = "SELECT * FROM Customer " & _
                  "WHERE CU01 = '" & Mid(strData, 1, 8) & "'"
      End If
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("CU04")) = False Then
            GetCustomer = rsTmp.Fields("CU04")
         ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
            GetCustomer = rsTmp.Fields("CU05")
         ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
            GetCustomer = rsTmp.Fields("CU06")
         End If
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Function

' 查詢資料庫取得資料
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 申請國家
   m_SP09 = Empty
   ' Key
   textSPKey = m_SP01 & m_SP02 & m_SP03 & m_SP04
   
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_SP01 & "' AND " & _
                  "SP02 = '" & m_SP02 & "' AND " & _
                  "SP03 = '" & m_SP03 & "' AND " & _
                  "SP04 = '" & m_SP04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_SP09 = rsTmp.Fields("SP09")
      End If
      ' 案件名稱(中)
      If IsNull(rsTmp.Fields("SP05")) = False Then
         textSP05 = rsTmp.Fields("SP05")
      End If
      ' 案件名稱(英)
      If IsNull(rsTmp.Fields("SP06")) = False Then
         textSP06 = rsTmp.Fields("SP06")
      End If
      ' 案件名稱(日)
      If IsNull(rsTmp.Fields("SP07")) = False Then
         textSP07 = rsTmp.Fields("SP07")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("SP08")) = False Then
         textSP08 = GetCustomer(rsTmp.Fields("SP08"))
      End If
   End If
   rsTmp.Close
   ' 顯示符合條件的資料
   ListData
   
   ' 90.07.25 modify 若只有一筆的情況下直接進入下一個畫面
   If grdList.Rows = 2 Then
      ' 90.10.09 modify by louis
      DisplayNextForm
      ' 檢查來函記錄檔
      'If PromptIfTaiwanNoResult() = True Then
      '   DisplayNextForm
      'Else
      '   Unload Me
      '   frm02010409_1.Show
      'End If
   End If
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 列出案件進度表符合條件的資料
Private Sub ListData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim bDeal As Boolean
   Dim bSpecial As Boolean
   Dim strCP09 As String
   
   InitialGrdList
   
   m_CP09 = Empty
   
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & m_SP01 & "' AND " & _
                  "CP02 = '" & m_SP02 & "' AND " & _
                  "CP03 = '" & m_SP03 & "' AND " & _
                  "CP04 = '" & m_SP04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         bSpecial = False
         strCP09 = Empty
         ' 無發文日不予列入
         If IsNull(rsTmp.Fields("CP27")) = True Then: GoTo NextRecord
         ' 發文日是空白不予列入
         If IsEmptyText(rsTmp.Fields("CP27")) = True Then: GoTo NextRecord
         ' 無收文號不予列入
         If IsNull(rsTmp.Fields("CP09")) = True Then: GoTo NextRecord
         ' 收文號不為A,B類的不予計入
         Select Case Mid(rsTmp.Fields("CP09"), 1, 1)
            Case "A", "B":
            '2010/1/21 ADD BY SONIA
         ' C類只取 被異議, 被異議(理由), 被評定, 被評定(理由), 被廢止, 被廢止(理由), 通知參加訴願, 通知參加訴訟, 通知言詞辯論, 通知準備程序, 通知行政上訴答辯
            Case "C"
               If IsNull(rsTmp.Fields("CP10")) = False Then
                  'Modify By Sindy 2024/9/26 改為常變數
                  If InStr(商爭審查來函案件性質, rsTmp.Fields("CP10")) = 0 Then
                     GoTo NextRecord
                  End If
'                  Select Case rsTmp.Fields("CP10")
'                     Case "1601", "1602", "1603", "1604", "1605", "1606", "1404", "1405", "1203", "1204", "1406":
'                     Case "1205", "1609", "1612", "1613", "1616", "1618", "1619", "1620", "1621"
'                     Case Else: GoTo NextRecord
'                  End Select
                  '2024/9/26 END
               End If
            '2010/1/21 END
            Case Else: GoTo NextRecord
         End Select
         ' 92.11.22 已有結果者不予列入
         'modify by sonia 2024/3/20 TC案要輸著作公告通知所以不要限制已有結果TC-011051已發註冊證才輸著作公告
         'If IsNull(rsTmp.Fields("CP24")) = False Then: GoTo NextRecord
         If m_SP01 <> "TC" Then
            If IsNull(rsTmp.Fields("CP24")) = False Then: GoTo NextRecord
         End If
         '92.11.22 end
                  
         ' 列入資料
         grdList.Rows = grdList.Rows + 1
         grdList.row = grdList.Rows - 1
         ' 收文號
         If IsNull(rsTmp.Fields("CP09")) = False Then
            strCP09 = rsTmp.Fields("CP09")
            grdList.TextMatrix(grdList.row, 1) = rsTmp.Fields("CP09")
         End If
         ' 收文日
         If IsNull(rsTmp.Fields("CP05")) = False Then
            grdList.TextMatrix(grdList.row, 2) = TAIWANDATE(rsTmp.Fields("CP05"))
         End If
         ' 案件性質
         If IsNull(rsTmp.Fields("CP10")) = False Then
            If m_SP09 = "000" Then
               grdList.TextMatrix(grdList.row, 3) = GetCaseTypeName(m_SP01, rsTmp.Fields("CP10"), 0)
            Else
               grdList.TextMatrix(grdList.row, 3) = GetCaseTypeName(m_SP01, rsTmp.Fields("CP10"), 1)
            End If
         End If
         ' 發文日
         If IsNull(rsTmp.Fields("CP27")) = False Then
            grdList.TextMatrix(grdList.row, 4) = TAIWANDATE(rsTmp.Fields("CP27"))
         End If
         ' 結果
         If IsNull(rsTmp.Fields("CP24")) = False Then
            'grdList.TextMatrix(grdList.Row, 5) = rsTmp.Fields("CP24")
            Select Case rsTmp.Fields("CP24")
               Case "1":
                  grdList.TextMatrix(grdList.row, 5) = "准/勝"
               Case "2":
                  grdList.TextMatrix(grdList.row, 5) = "駁/敗"
            End Select
         End If
         '2011/6/21 ADD BY SONIA
         ' 相關人
         bDeal = False
         If bDeal = False And IsNull(rsTmp.Fields("CP40")) = False Then
            If IsEmptyText(rsTmp.Fields("CP40")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP40")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP41")) = False Then
            If IsEmptyText(rsTmp.Fields("CP41")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP41")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP42")) = False Then
            If IsEmptyText(rsTmp.Fields("CP42")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP42")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP50")) = False Then
            If IsEmptyText(rsTmp.Fields("CP50")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP50")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP51")) = False Then
            If IsEmptyText(rsTmp.Fields("CP51")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP51")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP52")) = False Then
            If IsEmptyText(rsTmp.Fields("CP52")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP52")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP56")) = False Then
            If IsEmptyText(rsTmp.Fields("CP56")) = False Then
               grdList.TextMatrix(grdList.row, 6) = GetCustomerName(rsTmp.Fields("CP56"), 0)
               bDeal = True
            End If
         End If
         If bDeal = False Then: grdList.TextMatrix(grdList.row, 6) = Empty
         ' 進度備註
         If IsNull(rsTmp.Fields("CP64")) = False Then
            grdList.TextMatrix(grdList.row, 7) = rsTmp.Fields("CP64")
         End If
         '2011/6/21 END
         ' 案件性質 (代號)
         If IsNull(rsTmp.Fields("CP10")) = False Then
            grdList.TextMatrix(grdList.row, 8) = rsTmp.Fields("CP10")
         End If
         
NextRecord:
         rsTmp.MoveNext
      Loop
      'Added by Lydia 2023/10/18
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/18
   End If
   
   ' 設定第一列為選取的狀態
   grdList_SetSelection 1
   'Add By Cheng 2002/05/11
   '若只有一筆資料, 則直接進入下一畫面
   If Me.grdList.Rows = 2 Then
      DisplayNextForm
   End If
   
End Sub

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 9
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "收文號"
   grdList.ColWidth(1) = 1200
   grdList.col = 2
   grdList.Text = "收文日"
   grdList.ColWidth(2) = 900
   grdList.col = 3
   grdList.Text = "案件性質"
   grdList.ColWidth(3) = 1200
   grdList.col = 4
   grdList.Text = "發文日"
   grdList.ColWidth(4) = 1000
   grdList.col = 5
   grdList.Text = "結果"
   grdList.ColWidth(5) = 600
   grdList.col = 6
   grdList.Text = "相關人"
   grdList.ColWidth(6) = 2600
   grdList.col = 7
   grdList.Text = "進度備註"
   grdList.ColWidth(7) = 3000
   grdList.col = 8
   grdList.Text = "案件性質"
   grdList.ColWidth(8) = 0
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

Private Sub Form_Unload(Cancel As Integer)
'edit by nickc 2008/04/25 改整批印
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   
   'Add By Sindy 2019/5/22
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   'Add By Cheng 2002/07/18
   Set frm02010409_2 = Nothing
End Sub

Private Sub grdList_SelChange()
   If grdList.row > 0 Then
      grdList.col = 1
      m_CP09 = grdList.Text
      grdList.col = 8
      m_CP10 = grdList.Text
   End If
   grdList_ShowSelection
End Sub

' 設定Grid List的一列為選取的狀態
Private Sub grdList_SetSelection(ByVal nSel As Integer)
   If grdList.Rows >= 2 Then
      grdList.row = 1
      grdList_SelChange
      grdList_ShowSelection
   End If
End Sub

Private Sub DisplayNextForm()
   If m_SP01 = "TD" And (m_CP10 = "805" Or m_CP10 = "708") Then
      frm02010409_3.SetData 0, m_SP01, True
      frm02010409_3.SetData 1, m_SP02, False
      frm02010409_3.SetData 2, m_SP03, False
      frm02010409_3.SetData 3, m_SP04, False
      frm02010409_3.SetData 4, m_CP05, False
      frm02010409_3.SetData 5, m_CP09, False
      'Add By Sindy 2019/5/22
      If Not m_PrevForm Is Nothing Then
         Call frm02010409_3.SetParent(m_PrevForm)
      End If
      frm02010409_3.m_strIR01 = m_strIR01
      frm02010409_3.m_strIR02 = m_strIR02
      frm02010409_3.m_strIR03 = m_strIR03
      frm02010409_3.m_strIR04 = m_strIR04
      '2019/5/22 END
      Me.Hide
      frm02010409_3.Show
      frm02010409_3.QueryData
   ElseIf (m_SP01 = "TD" Or m_SP01 = "TC") And m_CP10 = "301" Then
      frm02010409_4.SetData 0, m_SP01, True
      frm02010409_4.SetData 1, m_SP02, False
      frm02010409_4.SetData 2, m_SP03, False
      frm02010409_4.SetData 3, m_SP04, False
      frm02010409_4.SetData 4, m_CP05, False
      frm02010409_4.SetData 5, m_CP09, False
      'Add By Sindy 2019/5/22
      If Not m_PrevForm Is Nothing Then
         Call frm02010409_4.SetParent(m_PrevForm)
      End If
      frm02010409_4.m_strIR01 = m_strIR01
      frm02010409_4.m_strIR02 = m_strIR02
      frm02010409_4.m_strIR03 = m_strIR03
      frm02010409_4.m_strIR04 = m_strIR04
      '2019/5/22 END
      Me.Hide
      frm02010409_4.Show
      frm02010409_4.QueryData
   ElseIf m_SP01 = "TB" Then
      frm02010409_5.SetData 0, m_SP01, True
      frm02010409_5.SetData 1, m_SP02, False
      frm02010409_5.SetData 2, m_SP03, False
      frm02010409_5.SetData 3, m_SP04, False
      frm02010409_5.SetData 4, m_CP05, False
      frm02010409_5.SetData 5, m_CP09, False
      'Add By Sindy 2019/5/22
      If Not m_PrevForm Is Nothing Then
         Call frm02010409_5.SetParent(m_PrevForm)
      End If
      frm02010409_5.m_strIR01 = m_strIR01
      frm02010409_5.m_strIR02 = m_strIR02
      frm02010409_5.m_strIR03 = m_strIR03
      frm02010409_5.m_strIR04 = m_strIR04
      '2019/5/22 END
      Me.Hide
      frm02010409_5.Show
      frm02010409_5.QueryData
   ElseIf m_SP01 = "TM" Then
      frm02010409_6.SetData 0, m_SP01, True
      frm02010409_6.SetData 1, m_SP02, False
      frm02010409_6.SetData 2, m_SP03, False
      frm02010409_6.SetData 3, m_SP04, False
      frm02010409_6.SetData 4, m_CP05, False
      frm02010409_6.SetData 5, m_CP09, False
      'Add By Sindy 2019/5/22
      If Not m_PrevForm Is Nothing Then
         Call frm02010409_6.SetParent(m_PrevForm)
      End If
      frm02010409_6.m_strIR01 = m_strIR01
      frm02010409_6.m_strIR02 = m_strIR02
      frm02010409_6.m_strIR03 = m_strIR03
      frm02010409_6.m_strIR04 = m_strIR04
      '2019/5/22 END
      Me.Hide
      frm02010409_6.Show
      frm02010409_6.QueryData
   Else
      frm02010409_7.SetData 0, m_SP01, True
      frm02010409_7.SetData 1, m_SP02, False
      frm02010409_7.SetData 2, m_SP03, False
      frm02010409_7.SetData 3, m_SP04, False
      frm02010409_7.SetData 4, m_CP05, False
      frm02010409_7.SetData 5, m_CP09, False
      'Add By Sindy 2019/5/22
      If Not m_PrevForm Is Nothing Then
         Call frm02010409_7.SetParent(m_PrevForm)
      End If
      frm02010409_7.m_strIR01 = m_strIR01
      frm02010409_7.m_strIR02 = m_strIR02
      frm02010409_7.m_strIR03 = m_strIR03
      frm02010409_7.m_strIR04 = m_strIR04
      '2019/5/22 END
      Me.Hide
      frm02010409_7.Show
      frm02010409_7.QueryData
   End If
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As ADODB.Recordset
   Dim strSql As String
   
   CheckDataValid = False
   ' 使用者需先選取一筆記錄
   If IsEmptyText(m_CP09) = True Then
      strTit = "檢核資料"
      strMsg = "請先選取一筆記錄"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   ' 檢查該筆資料是否存在
   Select Case m_SP01
      Case "TD", "TC":
         Select Case m_CP10
            Case "301":
               Set rsTmp = New ADODB.Recordset
               strSql = "SELECT * FROM ChangeEvent " & _
                        "WHERE CE01 = '" & m_CP09 & "' "
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
               If rsTmp.RecordCount <= 0 Then
                  rsTmp.Close
                  Set rsTmp = Nothing
                  strTit = "檢核資料"
                  strMsg = "該筆資料無變更事項的記錄"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  GoTo EXITSUB
               End If
               rsTmp.Close
               Set rsTmp = Nothing
            Case Else:
         End Select
   End Select
   CheckDataValid = True
EXITSUB:
End Function
