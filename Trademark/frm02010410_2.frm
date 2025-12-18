VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010410_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "廣告刊出來函輸入"
   ClientHeight    =   3900
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   7920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   7920
   Begin VB.TextBox textPrint 
      Height          =   264
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   3
      Top             =   2544
      Width           =   372
   End
   Begin VB.TextBox textWord 
      Height          =   264
      Left            =   5976
      MaxLength       =   1
      TabIndex        =   4
      Top             =   2544
      Width           =   372
   End
   Begin VB.TextBox textCP05 
      Height          =   264
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2232
      Width           =   1140
   End
   Begin VB.TextBox textSel 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1896
      Width           =   1116
   End
   Begin VB.TextBox textCP27_2 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   2136
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1536
      Width           =   828
   End
   Begin VB.TextBox textCP27_1 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   984
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1536
      Width           =   780
   End
   Begin VB.TextBox textCU05 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   2136
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   888
      Width           =   5640
   End
   Begin VB.TextBox textTM23 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   576
      Width           =   1116
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7044
      TabIndex        =   8
      Top             =   72
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6216
      TabIndex        =   7
      Top             =   72
      Width           =   800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&P)"
      Height          =   400
      Left            =   5064
      TabIndex        =   6
      Top             =   72
      Width           =   1128
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   345
      Left            =   5220
      TabIndex        =   26
      Top             =   2160
      Width           =   2565
      Begin VB.OptionButton Option1 
         Caption         =   "雜誌社"
         Height          =   285
         Index           =   1
         Left            =   1020
         TabIndex        =   2
         Top             =   30
         Width           =   945
      End
      Begin VB.OptionButton Option1 
         Caption         =   "報社"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   30
         Width           =   735
      End
   End
   Begin MSForms.TextBox textCP64 
      Height          =   888
      Left            =   1440
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   6360
      VariousPropertyBits=   -1467989989
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "11218;1566"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCU06 
      Height          =   264
      Left            =   2136
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1200
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
      Left            =   2136
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   576
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
      Caption         =   "刊登媒體 :"
      Height          =   255
      Index           =   4
      Left            =   4290
      TabIndex        =   25
      Top             =   2235
      Width           =   1035
   End
   Begin VB.Label Label21 
      Caption         =   "刊登廣告備註 :"
      Enabled         =   0   'False
      Height          =   252
      Left            =   48
      TabIndex        =   24
      Top             =   2904
      Visible         =   0   'False
      Width           =   1308
   End
   Begin VB.Label Label33 
      Caption         =   "(N:不印)"
      Height          =   252
      Left            =   1920
      TabIndex        =   23
      Top             =   2568
      Width           =   1332
   End
   Begin VB.Label Label34 
      Caption         =   "列印定稿 :"
      Height          =   252
      Left            =   72
      TabIndex        =   22
      Top             =   2568
      Width           =   972
   End
   Begin VB.Label Label35 
      Caption         =   "(Y:修改)"
      Height          =   252
      Left            =   6456
      TabIndex        =   21
      Top             =   2568
      Width           =   1332
   End
   Begin VB.Label Label36 
      Caption         =   "是否修改定稿內容 :"
      Height          =   252
      Left            =   4296
      TabIndex        =   20
      Top             =   2568
      Width           =   1572
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   252
      Index           =   3
      Left            =   72
      TabIndex        =   19
      Top             =   2232
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "廣告刊出筆數 :"
      Height          =   252
      Index           =   2
      Left            =   72
      TabIndex        =   17
      Top             =   1896
      Width           =   1236
   End
   Begin VB.Line Line1 
      X1              =   1848
      X2              =   2016
      Y1              =   1656
      Y2              =   1656
   End
   Begin VB.Label Label1 
      Caption         =   "發文日期 :"
      Height          =   252
      Index           =   1
      Left            =   72
      TabIndex        =   10
      Top             =   1560
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "申請人 :"
      Height          =   252
      Index           =   0
      Left            =   72
      TabIndex        =   9
      Top             =   576
      Width           =   852
   End
End
Attribute VB_Name = "frm02010410_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/29 Form2.0已修改 textCU04/textCU06/textCP64
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/6 日期欄已修改
Option Explicit

Dim m_TM23 As String
Dim m_TM10 As String 'Add By Sindy 2009/09/24
Dim m_CP27_1 As String
Dim m_CP27_2 As String
Dim m_CP09 As String

'Add By Sindy 2012/1/13
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
'2012/1/13 End

' 宣告欄位內容結構
Private Type CPFIELD
   cfCP01 As String
   cfCP02 As String
   cfCP03 As String
   cfCP04 As String
   cfcp09 As String
   cfcp12 As String
   cfcp13 As String
End Type

Dim m_CPList() As CPFIELD
Dim m_CPListCount As Integer
Dim m_SelNo As Integer


Private Sub ClearCPList()
   If m_CPListCount > 0 Then
      Erase m_CPList
   End If
   m_CPListCount = 0
End Sub

Public Sub SetData(ByVal strCP01 As String, _
                   ByVal strCP02 As String, _
                   ByVal strCP03 As String, _
                   ByVal strCP04 As String, _
                   ByVal strCP09 As String, _
                   ByVal strCP12 As String, _
                   ByVal strCP13 As String, _
                   ByVal nType As Integer, _
                   ByVal bClear As Boolean)
   Dim strSql As String
   If bClear Then
      ClearCPList
   End If
   Select Case nType
      Case 0: m_TM23 = strCP01
      Case 1: m_CP27_1 = strCP01
      Case 2: m_CP27_2 = strCP01
      Case 3: m_SelNo = Val(strCP01)
      Case 4:
         If Not IsEmptyText(strCP01) Then
            ReDim Preserve m_CPList(m_CPListCount + 1)
            m_CPList(m_CPListCount).cfCP01 = strCP01
            m_CPList(m_CPListCount).cfCP02 = strCP02
            m_CPList(m_CPListCount).cfCP03 = strCP03
            m_CPList(m_CPListCount).cfCP04 = strCP04
            m_CPList(m_CPListCount).cfcp09 = strCP09
            m_CPList(m_CPListCount).cfcp12 = strCP12
            m_CPList(m_CPListCount).cfcp13 = strCP13
            m_CPListCount = m_CPListCount + 1
            
            'Add By Sindy 2009/09/24
            strSql = "SELECT * FROM TradeMark WHERE TM01='" & strCP01 & "' AND TM02='" & strCP02 & "' AND TM03='" & strCP03 & "' AND TM04='" & strCP04 & "' "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               m_TM10 = "" & RsTemp("TM10")
            End If
            '2009/09/24 End
         End If
   End Select
End Sub

Private Sub cmdCancel_Click()
   frm02010410_1.Show
   '911015 nick
   frm02010410_1.cmdQuery.Default = True
   frm02010410_1.textTM23.SetFocus
   Unload Me
End Sub

Private Sub cmdExit_Click()
   Unload frm02010410_1
   Unload Me
End Sub

Private Sub cmdOK_Click()
   If CheckDataValid() Then
    'Modify By Cheng 2002/11/07
'      'OnSaveData
    If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
    'Add By Cheng 2002/11/08
   ' 列印定稿
   If textPrint <> "N" Then
      PrintLetter
   End If
      
      frm02010410_1.Show
      '911015 nick 回前畫面清畫面
      '***** start
      frm02010410_1.textTM23.Text = ""
      frm02010410_1.textCU04.Text = ""
      frm02010410_1.textCU05.Text = ""
      frm02010410_1.textCU06.Text = ""
      frm02010410_1.textCP27_1.Text = ""
      frm02010410_1.textCP27_2.Text = ""
      frm02010410_1.textSel.Text = ""
      frm02010410_1.InitialGrdList
      frm02010410_1.cmdQuery.Default = True
      frm02010410_1.textTM23.SetFocus
      '***** end
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   textTM23.BackColor = &H8000000F
   textCU04.BackColor = &H8000000F
   textCU05.BackColor = &H8000000F
   textCU06.BackColor = &H8000000F
   textCP27_1.BackColor = &H8000000F
   textCP27_2.BackColor = &H8000000F
   textSel.BackColor = &H8000000F
   '911015 nick  收文日預設為系統日
   textCP05 = ChangeWStringToTString(ServerDate)
   textTM23 = m_TM23
   textCP27_1 = m_CP27_1
   textCP27_2 = m_CP27_2
   textSel = CStr(m_SelNo)
   ' 顯示申請人名稱
   UpdateCustomerName textTM23
  
   MoveFormToCenter Me
   
   '91.10.29 MODIFY BY SONIA
   textWord = "Y"
   '91.10.29 END
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

'Modify By Cheng 2002/11/07
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
   Dim strSql As String
   Dim nIndex As Integer
   Dim strCP09 As String

'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans

   For nIndex = 0 To m_CPListCount - 1
      strCP09 = AutoNo("C", 6)
      If nIndex = 0 Then
         m_CP09 = strCP09
         'Add By Sindy 2012/1/13
         m_CP01 = m_CPList(nIndex).cfCP01
         m_CP02 = m_CPList(nIndex).cfCP02
         m_CP03 = m_CPList(nIndex).cfCP03
         m_CP04 = m_CPList(nIndex).cfCP04
         '2012/1/13 End
      End If
    '承辦人為使用者, 發文日為系統日
        'Modify By Cheng 2003/04/03
        '智權人員存最近收文A類接洽記錄單的智權人員
        'Modify By Cheng 2004/02/04
        '業務區為最近收文A類接洽記錄單智權人員的業務區
'      strSQL = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP64) " & _
'               "VALUES ('" & m_CPList(nIndex).cfCP01 & "','" & m_CPList(nIndex).cfCP02 & "','" & m_CPList(nIndex).cfCP03 & "','" & m_CPList(nIndex).cfCP04 & "'," & _
'                       DBDATE(textCP05) & ",'" & strCP09 & "','" & "1710" & "','" & m_CPList(nIndex).cfcp12 & "','" & PUB_GetAKindSalesNo(m_CPList(nIndex).cfCP01, m_CPList(nIndex).cfCP02, m_CPList(nIndex).cfCP03, m_CPList(nIndex).cfCP04) & "','" & strUserNum & "','" & "N" & "','" & "N" & "'," & _
'                       DBDATE(SystemDate()) & ",'" & "N" & "','" & m_CPList(nIndex).cfcp09 & "','" & textCP64 & "') "
      strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP64) " & _
               "VALUES ('" & m_CPList(nIndex).cfCP01 & "','" & m_CPList(nIndex).cfCP02 & "','" & m_CPList(nIndex).cfCP03 & "','" & m_CPList(nIndex).cfCP04 & "'," & _
                       DBDATE(textCP05) & ",'" & strCP09 & "','" & "1710" & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_CPList(nIndex).cfCP01, m_CPList(nIndex).cfCP02, m_CPList(nIndex).cfCP03, m_CPList(nIndex).cfCP04)) & "','" & PUB_GetAKindSalesNo(m_CPList(nIndex).cfCP01, m_CPList(nIndex).cfCP02, m_CPList(nIndex).cfCP03, m_CPList(nIndex).cfCP04) & "','" & strUserNum & "','" & "N" & "','" & "N" & "'," & _
                       DBDATE(SystemDate()) & ",'" & "N" & "','" & m_CPList(nIndex).cfcp09 & "','" & textCP64 & "') "
        'End
      cnnConnection.Execute strSql
      '911015 nick  update cp24='1' and cp25= textCP05
      strSql = "update caseprogress set cp24='1',cp25=" & ChangeTStringToWString(textCP05) & " where cp09='" & m_CPList(nIndex).cfcp09 & "' "
      cnnConnection.Execute strSql
          
      'Add By Sindy 2009/09/24
      '因為有些來函由內商輸入，內商有自行控管之承辦期限及發文日。改為內商輸入所有C類來函，
      '若業務區為F字頭者，除爭議受理外，自動產生B類收文，案件性質為外商發文722，不上發文日，不向客戶請款
      Dim strCP48 As String, strCP09B As String
      If Left(GetSalesArea(PUB_GetAKindSalesNo(m_CPList(nIndex).cfCP01, m_CPList(nIndex).cfCP02, m_CPList(nIndex).cfCP03, m_CPList(nIndex).cfCP04)), 1) = "F" And _
         ((m_CPList(nIndex).cfCP01 = "T" And m_TM10 = "020") Or (m_CPList(nIndex).cfCP01 = "FCT" And m_TM10 = "000")) Then
         strCP09B = AutoNo("B", 6)
         '承辦期限為系統日加4個工作天
         strCP48 = DBDATE(Pub_GetHandleDay(m_CPList(nIndex).cfCP01, m_TM10, "722", strSrvDate(1), , m_CPList(nIndex).cfcp09))
         '2011/4/28 modify by sonia 智權人員原抓點選收文號之智權人員,改抓該案最後收文在職智權人員
         strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13,cp14,cp48,cp20,cp26,cp32,cp43) " & _
                        "values (" & CNULL(m_CPList(nIndex).cfCP01) & "," & CNULL(m_CPList(nIndex).cfCP02) & "," & CNULL(m_CPList(nIndex).cfCP03) & _
                        "," & CNULL(m_CPList(nIndex).cfCP04) & "," & CNULL(strSrvDate(1)) & "," & CNULL(strCP09B) & ",722," & _
                        CNULL(GetSalesArea(PUB_GetFCTSalesNo(CNULL(m_CPList(nIndex).cfCP01), CNULL(m_CPList(nIndex).cfCP02), CNULL(m_CPList(nIndex).cfCP03), CNULL(m_CPList(nIndex).cfCP04)))) & "," & CNULL(PUB_GetFCTSalesNo(CNULL(m_CPList(nIndex).cfCP01), CNULL(m_CPList(nIndex).cfCP02), CNULL(m_CPList(nIndex).cfCP03), CNULL(m_CPList(nIndex).cfCP04))) & "," & CNULL(PUB_GetFCTSalesNo(CNULL(m_CPList(nIndex).cfCP01), CNULL(m_CPList(nIndex).cfCP02), CNULL(m_CPList(nIndex).cfCP03), CNULL(m_CPList(nIndex).cfCP04))) & "," & CNULL(strCP48) & ",'N','N','N'," & CNULL(strCP09) & ")"
         cnnConnection.Execute strSql
      End If
   Next nIndex
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Modify By Cheng 2002/11/08
'   ' 列印定稿
'   If textPrint <> "N" Then
'      PrintLetter
'   End If
'Add By Cheng 2002/11/07
cnnConnection.CommitTrans
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    OnSaveData = False
End Function

Private Sub textCP05_GotFocus()
   InverseTextBox textCP05
End Sub

Private Sub textCP05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP05) = False Then
      If CheckIsTaiwanDate(textCP05, False) = False Then
         Cancel = True
         strMsg = "請輸入正確的來函收文日起"
         strTit = "來函收文"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP05_GotFocus
         GoTo EXITSUB
      End If
      If Val(DBDATE(textCP05)) > Val(DBDATE(SystemDate())) Then
         Cancel = True
         strMsg = "來函收文不可超過系統日"
         strTit = "來函收文"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP05_GotFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If Not IsEmptyText(textPrint) Then
      If textPrint <> "N" Then
         Cancel = True
         strMsg = "列印定稿只可輸入N或空白"
         strTit = "列印定稿"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPrint_GotFocus
      End If
   End If
End Sub

Private Sub textWord_GotFocus()
   InverseTextBox textWord
End Sub

Private Sub textWord_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textWord_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If Not IsEmptyText(textWord) Then
      If textWord <> "Y" Then
         Cancel = True
         strMsg = "是否修改定稿只可輸入Y或空白"
         strTit = "是否修改定稿"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textWord_GotFocus
      End If
   End If
End Sub

Private Function CheckDataValid()
   Dim strDay As String
   Dim strTemp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   'Add by Amy 2021/12/29檢查畫面的 TextBox是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        GoTo EXITSUB
   End If

   ' 來函收文日不可為空白
   If IsEmptyText(textCP05) = True Then
      strTit = "檢核資料"
      strMsg = "來函收文日不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP05.SetFocus
      GoTo EXITSUB
   End If
    'Add By Cheng 2003/01/01
    If Me.Option1(0).Value = False And Me.Option1(1).Value = False Then
        MsgBox "請選擇廣告刊登媒體!!!", vbExclamation + vbOKOnly
        GoTo EXITSUB
    End If
   CheckDataValid = True
EXITSUB:
End Function

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
   Dim strSql As String
   
    'Modify By Cheng 2003/01/02
    '廣告刊登於報社
    If Me.Option1(0).Value Then
        ' 清除定稿例外欄位檔原有資料
        EndLetter "13", m_CP09, "00", strUserNum
        ' 列印備註
        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                 "VALUES ('" & "13" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
                 "'" & "列印備註" & "','" & textCP64 & "')"
        cnnConnection.Execute strSql
    '廣告刊登於雜誌處
    Else
        ' 清除定稿例外欄位檔原有資料
        EndLetter "13", m_CP09, "01", strUserNum
        ' 列印備註
        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                 "VALUES ('" & "13" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & "'," & _
                 "'" & "列印備註" & "','" & textCP64 & "')"
        cnnConnection.Execute strSql
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
'Add By Sindy 2012/1/13
Dim ET01 As String, ET02 As String, ET03 As String, bolEdit As Boolean
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
'2012/1/13 End
   
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   'Add By Sindy 2012/1/13
   ET01 = "13"
   ET02 = m_CP09
   bolEdit = IIf(Me.textWord.Text = "Y", True, False)
   '2012/1/13 End
   
   ' 列印定稿
    'Modify By Cheng 2003/01/02
'   If textWord = "Y" Then
'      NowPrint m_CP09, "13", "00", True, strUserNum, 0
'   Else
'      NowPrint m_CP09, "13", "00", False, strUserNum, 0
'   End If
    If Me.Option1(0).Value = True Then
'        NowPrint m_CP09, "13", "00", IIf(Me.textWord.Text = "Y", True, False), strUserNum, 0
      ET03 = "00" 'Modify By Sindy 2012/1/13
    Else
'        NowPrint m_CP09, "13", "01", IIf(Me.textWord.Text = "Y", True, False), strUserNum, 0
      ET03 = "01" 'Modify By Sindy 2012/1/13
    End If
    
   'Add By Sindy 2012/1/13
   If ET03 <> "" Then
      bolEmail = PUB_GetEMailFlag(m_CP01 & m_CP02 & m_CP03 & m_CP04, , , bolPlusPaper)
      If bolEmail Then
         '判斷是否EMail同時寄紙本
         If Not bolPlusPaper Then
            iCopy = 1
         End If
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , True, True
         MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_CP01) & " ]！"
      Else
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0
      End If
   End If
   '2012/1/13 End
End Sub
