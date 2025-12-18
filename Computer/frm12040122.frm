VERSION 5.00
Begin VB.Form frm12040122 
   BorderStyle     =   1  '單線固定
   Caption         =   "規費資料稽核表"
   ClientHeight    =   2325
   ClientLeft      =   2715
   ClientTop       =   2385
   ClientWidth     =   4410
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4410
   Begin VB.TextBox textAX205 
      Height          =   270
      Left            =   1272
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1320
      Width           =   972
   End
   Begin VB.TextBox textDate_1 
      Height          =   270
      Left            =   1272
      MaxLength       =   7
      TabIndex        =   0
      Top             =   960
      Width           =   972
   End
   Begin VB.TextBox textDate_2 
      Height          =   270
      Left            =   2592
      MaxLength       =   7
      TabIndex        =   1
      Top             =   960
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2496
      TabIndex        =   3
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   3324
      TabIndex        =   4
      Top             =   120
      Width           =   800
   End
   Begin VB.Label Label2 
      Caption         =   "會計科目："
      Height          =   252
      Left            =   312
      TabIndex        =   7
      Top             =   1344
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "發文日："
      Height          =   252
      Left            =   312
      TabIndex        =   6
      Top             =   960
      Width           =   732
   End
   Begin VB.Label Label3 
      Caption         =   "－"
      Height          =   132
      Left            =   2352
      TabIndex        =   5
      Top             =   1008
      Width           =   252
   End
End
Attribute VB_Name = "frm12040122"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit

Const m_CharWidth = 120
Const m_CharHeight = 240
Const m_PaperSize = "A4"
' 宣告報表表頭的欄位其資料型態
Private Type REPORTFIELD
   Name As String
   Left As Long
   Width As Long
End Type
' 表頭欄位的內容
Dim m_Field(8) As REPORTFIELD
' 報表左方留白的寬度
Dim m_LeftMargin As Integer
' 報表上方留白的高度
Dim m_TopMargin As Integer
' 報表頁首的高度
Dim m_HeaderHeight As Integer
' 報表文件的寬度
Dim m_ReportWidth As Integer
' 報表文件中可容納的資料列數
Dim m_ReportDataRows As Integer
' 開始日期
Dim m_DateFrom As String
' 結束日期
Dim m_DateTo As String
' 預設印表機
'Dim m_DefaultPrinter As String
'
Dim m_nRow As Integer
Dim m_nPage As Integer
'Add By Cheng 2002/09/10
Dim blnClkSure As Boolean

Private Sub Form_Load()
   'Dim Prn As Printer
   'Dim nIndex As Integer
   'Dim nSel As Integer
   
   'm_DefaultPrinter = Printer.DeviceName
   MoveFormToCenter Me
   
   'nSel = 0
   'nIndex = 0
   'For Each Prn In Printers
   '   cmbPrinter.AddItem Prn.DeviceName
   '   If Prn.DeviceName = m_DefaultPrinter Then
   '      nSel = nIndex
   '   End If
   '   nIndex = nIndex + 1
   'Next
   'cmbPrinter.ListIndex = nSel
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Dim Prn As Printer
   ''搜尋 Printer
   'For Each Prn In Printers
   '   If Prn.DeviceName = m_DefaultPrinter Then
   '      Set Printer = Prn
   '      Exit For
   '   End If
   'Next
   'Add By Cheng 2002/07/18
   Set frm12040122 = Nothing
End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim bData As Boolean
   'Dim Prn As Printer
   
   If CheckDataValid() = True Then
      ''搜尋 Printer
      'For Each Prn In Printers
      '   If Prn.DeviceName = cmbPrinter.Text Then
      '      Set Printer = Prn
      '      Exit For
      '   End If
      'Next
      
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      
      If IsEmptyText(textDate_1) = False Then
         m_DateFrom = DBDATE(textDate_1)
      Else
         m_DateFrom = textDate_1
      End If
      If IsEmptyText(textDate_2) = False Then
         m_DateTo = DBDATE(textDate_2)
      Else
         m_DateTo = textDate_2
      End If
      ' 設定欄位規格
      BuildField
      ' 產生報表
      bData = GenerateReport
      
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      
      If bData = True Then
         strTit = "輸出報表"
         strMsg = "列印結束"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

' 搜尋資料
Private Function GenerateReport() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strSql As String
   Dim strSubSQL As String
   Dim rsTmp As New ADODB.Recordset
   Dim strAX205 As String
   Dim bData As Boolean
   
   GenerateReport = False
   bData = False
   
   If IsEmptyText(textAX205) = True Then
      strAX205 = "(AX205 = '2201' OR AX205 = '220101' OR AX205 = '220102' OR AX205 = '220103' OR AX205 = '220104' OR AX205 = '220107') "
   Else
      strAX205 = "AX205 = '" & textAX205 & "' "
   End If
   
   ' 案件進度檔有而傳票資料檔沒有的資料
   '2010/9/17 modify by sonia cp17改抓cp84,AX207應抓AX206
   strSql = "SELECT 1 AS KEY1,NVL(CP5-19110000,AC5) AS KEY2,NVL(CP1,AC1) AS KEY3,NVL(CP2,AC2) AS KEY4,NVL(CP3,AC3) AS KEY5,NVL(CP4,AC4) AS KEY6,CP1,CP2,CP3,CP4,CP5,CP6,AC1,AC2,AC3,AC4,AC5,AC6 FROM " & _
               "(SELECT CP01 AS CP1,CP02 AS CP2,CP03 AS CP3,CP04 AS CP4,CP27 AS CP5,SUM(cp84) AS CP6 FROM CASEPROGRESS "
   strSql = strSql & "WHERE CP84 IS NOT NULL AND " & _
                           "CP84 <> 0 AND " & _
                           "(CP09 LIKE 'A%' OR CP09 LIKE 'B%') "
   If IsEmptyText(m_DateFrom) = False Then: strSql = strSql & "AND CP27 >= " & DBDATE(m_DateFrom) & " "
   If IsEmptyText(m_DateTo) = False Then: strSql = strSql & "AND CP27 <= " & DBDATE(m_DateTo) & " "
   strSql = strSql & "GROUP BY CP01,CP02,CP03,CP04,CP27 "
   strSql = strSql & "),(SELECT SUBSTR(AX214,1,LENGTH(AX214) - 9) AS AC1,SUBSTR(AX214,LENGTH(AX214) - 9 + 1, 6) AS AC2,SUBSTR(AX214,LENGTH(AX214) - 3 + 1, 1) AS AC3,SUBSTR(AX214,LENGTH(AX214) - 2 + 1, 2) AS AC4,A0205 AS AC5,SUM(AX206) AS AC6 FROM ACC020, ACC021 "
   'strSQL = strSQL & "WHERE (AX205 = '2201' OR AX205 = '220101' OR AX205 = '220102' OR AX205 = '220103' OR AX205 = '220104' OR AX205 = '220107') AND AX202 = A0202 AND AX201 = A0201 "
   strSql = strSql & "WHERE " & strAX205 & "AND AX202 = A0202 AND AX201 = A0201 "
   If IsEmptyText(m_DateFrom) = False Then: strSql = strSql & "AND A0205 >= " & TAIWANDATE(m_DateFrom) & " "
   If IsEmptyText(m_DateTo) = False Then: strSql = strSql & "AND A0205 <= " & TAIWANDATE(m_DateTo) & " "
   strSql = strSql & "GROUP BY AX214, A0205) " & _
            "WHERE CP1 = AC1(+) AND " & _
                  "CP2 = AC2(+) AND " & _
                  "CP3 = AC3(+) AND " & _
                  "CP4 = AC4(+) AND " & _
                  "(CP5 - 19110000) = AC5(+) AND " & _
                  "AC1 IS NULL "

   strSql = strSql & "UNION ALL "

   ' 傳票資料檔有而案件進度檔沒有的資料
   strSql = strSql & _
            "SELECT 2 AS KEY1,NVL(AC5,CP5-19110000) AS KEY2,NVL(AC1,CP1) AS KEY3,NVL(AC2,CP2) AS KEY4,NVL(AC3,CP3) AS KEY5,NVL(AC4,CP4) AS KEY6,CP1,CP2,CP3,CP4,CP5,CP6,AC1,AC2,AC3,AC4,AC5,AC6 FROM " & _
               "(SELECT CP01 AS CP1,CP02 AS CP2,CP03 AS CP3,CP04 AS CP4,CP27 AS CP5,SUM(cp84) AS CP6 FROM CASEPROGRESS "
   strSql = strSql & "WHERE cp84 IS NOT NULL AND " & _
                           "cp84 <> 0 AND " & _
                           "(CP09 LIKE 'A%' OR CP09 LIKE 'B%') "
   If IsEmptyText(m_DateFrom) = False Then: strSql = strSql & "AND CP27 >= " & DBDATE(m_DateFrom) & " "
   If IsEmptyText(m_DateTo) = False Then: strSql = strSql & "AND CP27 <= " & DBDATE(m_DateTo) & " "
   strSql = strSql & "GROUP BY CP01,CP02,CP03,CP04,CP27 "
   strSql = strSql & "),(SELECT SUBSTR(AX214,1,LENGTH(AX214) - 9) AS AC1,SUBSTR(AX214,LENGTH(AX214) - 9 + 1, 6) AS AC2,SUBSTR(AX214,LENGTH(AX214) - 3 + 1, 1) AS AC3,SUBSTR(AX214,LENGTH(AX214) - 2 + 1, 2) AS AC4,A0205 AS AC5,SUM(AX206) As AC6 FROM ACC020, ACC021 "
   'strSQL = strSQL & "WHERE (AX205 = '2201' OR AX205 = '220101' OR AX205 = '220102' OR AX205 = '220103' OR AX205 = '220104' OR AX205 = '220107') AND AX202 = A0202 AND AX201 = A0201 "
   strSql = strSql & "WHERE " & strAX205 & "AND AX202 = A0202 AND AX201 = A0201 "
   If IsEmptyText(m_DateFrom) = False Then: strSql = strSql & "AND A0205 >= " & TAIWANDATE(m_DateFrom) & " "
   If IsEmptyText(m_DateTo) = False Then: strSql = strSql & "AND A0205 <= " & TAIWANDATE(m_DateTo) & " "
   strSql = strSql & "GROUP BY AX214, A0205) " & _
            "WHERE AC1 = CP1(+) AND " & _
                  "AC2 = CP2(+) AND " & _
                  "AC3 = CP3(+) AND " & _
                  "AC4 = CP4(+) AND " & _
                  "(AC5 + 19110000) = CP5(+) AND " & _
                  "CP1 IS NULL "
   
   strSql = strSql & "UNION ALL "
   
   ' 傳票資料檔有而案件進度檔也有但規費不相等的資料
   strSql = strSql & _
            "SELECT 3 AS KEY1,NVL(CP5-19110000,AC5) AS KEY2,NVL(CP1,AC1) AS KEY3,NVL(CP2,AC2) AS KEY4,NVL(CP3,AC3) AS KEY5,NVL(CP4,AC4) AS KEY6,CP1,CP2,CP3,CP4,CP5,CP6,AC1,AC2,AC3,AC4,AC5,AC6 FROM " & _
               "(SELECT CP01 AS CP1,CP02 AS CP2,CP03 AS CP3,CP04 AS CP4,CP27 AS CP5,SUM(cp84) AS CP6 FROM CASEPROGRESS "
   strSql = strSql & "WHERE cp84 IS NOT NULL AND " & _
                           "cp84 <> 0 AND " & _
                           "(CP09 LIKE 'A%' OR CP09 LIKE 'B%') "
   If IsEmptyText(m_DateFrom) = False Then: strSql = strSql & "AND CP27 >= " & DBDATE(m_DateFrom) & " "
   If IsEmptyText(m_DateTo) = False Then: strSql = strSql & "AND CP27 <= " & DBDATE(m_DateTo) & " "
   strSql = strSql & "GROUP BY CP01,CP02,CP03,CP04,CP27 "
   strSql = strSql & "),(SELECT SUBSTR(AX214,1,LENGTH(AX214) - 9) AS AC1,SUBSTR(AX214,LENGTH(AX214) - 9 + 1, 6) AS AC2,SUBSTR(AX214,LENGTH(AX214) - 3 + 1, 1) AS AC3,SUBSTR(AX214,LENGTH(AX214) - 2 + 1, 2) AS AC4,A0205 AS AC5,SUM(AX206) As AC6 FROM ACC020, ACC021 "
   'strSQL = strSQL & "WHERE (AX205 = '2201' OR AX205 = '220101' OR AX205 = '220102' OR AX205 = '220103' OR AX205 = '220104' OR AX205 = '220107') AND AX202 = A0202 AND AX201 = A0201 "
   strSql = strSql & "WHERE " & strAX205 & "AND AX202 = A0202 AND AX201 = A0201 "
   If IsEmptyText(m_DateFrom) = False Then: strSql = strSql & "AND A0205 >= " & TAIWANDATE(m_DateFrom) & " "
   If IsEmptyText(m_DateTo) = False Then: strSql = strSql & "AND A0205 <= " & TAIWANDATE(m_DateTo) & " "
   strSql = strSql & "GROUP BY AX214, A0205) " & _
            "WHERE AC1 = CP1 AND " & _
                  "AC2 = CP2 AND " & _
                  "AC3 = CP3 AND " & _
                  "AC4 = CP4 AND " & _
                  "(AC5 + 19110000) = CP5 AND " & _
                  "AC6 <> CP6 "
   strSql = strSql & "ORDER BY KEY1,KEY2,KEY3,KEY4,KEY5,KEY6 "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      bData = PrintData(rsTmp)
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   If bData = False Then
      strTit = "搜尋資料"
      strMsg = "資料庫中沒有符合的資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
   
   GenerateReport = bData

End Function

' 檢查輸入的資料是否完整
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
      
   'Add By Cheng 2002/09/11
   blnClkSure = False
   
   CheckDataValid = False
   ' 發文日(起)
   If IsEmptyText(textDate_1) = True Then
      strTit = "查詢資料"
      strMsg = "請輸入發文日起日"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textDate_1.SetFocus
      GoTo EXITSUB
   End If
   
   ' 發文日(迄)
   If IsEmptyText(textDate_2) = True Then
      strTit = "查詢資料"
      strMsg = "請輸入發文日迄日"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textDate_2.SetFocus
      GoTo EXITSUB
   End If

   'Add By Cheng 2002/03/21
   If PUB_CheckKeyInDate(Me.textDate_1) = -1 Then
      Me.textDate_1.SetFocus
      textDate_1_GotFocus
      GoTo EXITSUB
   End If
   If PUB_CheckKeyInDate(Me.textDate_2) = -1 Then
      Me.textDate_2.SetFocus
      textDate_2_GotFocus
      GoTo EXITSUB
   End If


   ' 起日不可超過迄日
   If IsEmptyText(textDate_1) = False And IsEmptyText(textDate_2) = False Then
      If Val(DBDATE(textDate_1)) > Val(DBDATE(textDate_2)) Then
         strTit = "查詢資料"
         strMsg = "來函收文日起日不可超過迄日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         'Add By Cheng 2002/09/11
         blnClkSure = True
         
         textDate_1.SetFocus
         textDate_1_GotFocus
         GoTo EXITSUB
      End If
   End If
   'Add By Cheng 2002/09/10
   If Me.textAX205.Text <> "" Then
      Select Case textAX205.Text
         Case "2201", "220101", "220102", "220103", "220104", "220107"
         Case Else
            MsgBox "會計科目代號輸入錯誤!!!", vbExclamation + vbOKOnly
            Me.textAX205.SetFocus
            textAX205_GotFocus
            GoTo EXITSUB
      End Select
   End If
   CheckDataValid = True
EXITSUB:
End Function

' 會計科目
Private Sub textAX205_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textAX205) = False Then
      Select Case textAX205
         Case "2201", "220101", "220102", "220103", "220104", "220107":
         Case Else
            Cancel = True
            strTit = "檢核資料"
            strMsg = "會計科目代號不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textAX205_GotFocus
      End Select
   End If
End Sub

' 發文日起日
Private Sub textDate_1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textDate_1) = False Then
      If CheckIsTaiwanDate(textDate_1, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "發文日起日格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDate_1_GotFocus
      End If
   End If
End Sub

Private Sub textDate_2_LostFocus()
   'Add By Cheng 2002/09/10
   If blnClkSure = False Then
      If Me.textDate_1.Text <> "" And Me.textDate_2.Text <> "" Then
         If Val(Me.textDate_1.Text) > Val(Me.textDate_2.Text) Then
            MsgBox "來函收文日範圍輸入錯誤!!!", vbExclamation + vbOKOnly
            Me.textDate_1.SetFocus
            textDate_1_GotFocus
            Exit Sub
         End If
      End If
   Else
      blnClkSure = False
   End If
End Sub

' 發文日迄日
Private Sub textDate_2_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textDate_2) = False Then
      If CheckIsTaiwanDate(textDate_2, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "發文日迄日格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDate_2_GotFocus
      End If
   End If
End Sub

Private Sub textDate_1_GotFocus()
   InverseTextBox textDate_1
End Sub

Private Sub textDate_2_GotFocus()
   InverseTextBox textDate_2
End Sub

Private Sub textAX205_GotFocus()
   InverseTextBox textAX205
End Sub

' 列印分隔線
Public Sub PrintSplitLine(ByVal nRow As Integer)
   Dim nCount As Integer
   For nCount = 0 To m_ReportWidth - 1
      Printer.CurrentX = (m_LeftMargin + nCount) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "-"
   Next nCount
End Sub

' 列印分隔線
Public Sub PrintTerminateLine(ByVal nRow As Integer)
   Dim nCount As Integer
   For nCount = 0 To m_ReportWidth - 1
      Printer.CurrentX = (m_LeftMargin + nCount) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "="
   Next nCount
End Sub

' 設定報表欄位的左方位置及其名稱
Public Sub BuildField()
   Dim nIndex As Integer
   Dim nFieldWidth As Integer
   Dim nLeft As Integer
   
   Select Case m_PaperSize
      Case "REPORT"
         m_LeftMargin = 1
         m_TopMargin = 3
         m_ReportWidth = 154
         m_ReportDataRows = 45
         nFieldWidth = 9
      Case Else
         m_LeftMargin = 5
         m_TopMargin = 2
         m_ReportWidth = 130
         m_ReportDataRows = 30
         nFieldWidth = 7
   End Select
   
   nLeft = m_LeftMargin
   For nIndex = 0 To 7
      m_Field(nIndex).Left = nLeft
      Select Case nIndex
         Case 0:
            m_Field(nIndex).Name = "發文日"
            m_Field(nIndex).Width = 10
         Case 1:
            m_Field(nIndex).Name = "本所案號"
            m_Field(nIndex).Width = 16
         Case 2:
            m_Field(nIndex).Name = "案件性質"
            m_Field(nIndex).Width = 20
         Case 3:
            m_Field(nIndex).Name = "傳票號碼"
            m_Field(nIndex).Width = 16
         Case 4:
            m_Field(nIndex).Name = "規費金額"
            m_Field(nIndex).Width = 14
         Case 5:
            m_Field(nIndex).Name = "備註/摘要"
            m_Field(nIndex).Width = 40
      End Select
      nLeft = nLeft + m_Field(nIndex).Width
   Next nIndex
End Sub

' 列印表頭
Private Sub PrintPageHeader(ByVal nPage As Integer)
   Dim nCount As Integer
   Dim nIndex As Integer
   Dim nRow As Integer
   Dim nX As Long
   Dim nY As Long
   Dim nCenter As Long
   Dim strTemp As String

   ' 表頭
   nRow = 1
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.FontSize = 24
   Printer.Font.Underline = True
   nX = m_LeftMargin + m_ReportWidth / 2 - 14
   Printer.CurrentX = nX * m_CharWidth
   Printer.Print "規費資料稽核表"

   Printer.Font.Underline = False
   ' 下二列
   nRow = nRow + 2
   Printer.FontSize = 12
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "列印人 : " & strUserName

   nX = m_LeftMargin + m_ReportWidth - 20
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "製表日期 : " & Format(Date, "EE/MM/DD")
   ' 下一列
   nRow = nRow + 1
   ' 發文日
   nX = m_LeftMargin + m_ReportWidth / 2 - 16
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "發文日 : "
   ' 印日期的起迄
   nX = nX + 9
   If IsEmptyText(m_DateFrom) = False Then
      Printer.CurrentX = nX * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print Format(TAIWANDATE(m_DateFrom), "###/##/##")
   End If
   nX = nX + 9
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "-"
   nX = nX + 2
   If IsEmptyText(m_DateTo) = False Then
      Printer.CurrentX = nX * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print Format(TAIWANDATE(m_DateTo), "###/##/##")
   End If
   
   nX = m_LeftMargin + m_ReportWidth - 20
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "頁"

   nX = m_LeftMargin + m_ReportWidth - 14
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "次 : " & nPage

   ' 列印分隔線
   nRow = nRow + 1
   PrintSplitLine nRow

   nRow = nRow + 1
   For nIndex = 0 To 7
      Select Case nIndex
         Case 0, 1, 3, 4
            nCenter = ((m_Field(nIndex).Left * m_CharWidth) + (m_Field(nIndex).Left + m_Field(nIndex).Width) * m_CharWidth) / 2
            strTemp = LeftStr(m_Field(nIndex).Name, m_Field(nIndex).Width)
            Printer.CurrentX = nCenter - Printer.TextWidth(strTemp) / 2
            Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
            Printer.Print strTemp
         Case Else
            strTemp = LeftStr(m_Field(nIndex).Name, m_Field(nIndex).Width)
            Printer.CurrentX = m_Field(nIndex).Left * m_CharWidth
            Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
            Printer.Print strTemp
      End Select
   Next nIndex

   ' 列印分隔線
   nRow = nRow + 1
   PrintSplitLine nRow

   m_HeaderHeight = nRow
End Sub

' 列印該筆資料
Private Sub PrintRow(ByRef strField() As String)
   Dim nIndex As Integer
   Dim nLeft As Integer

   ' 若列數超過頁面的高度限制時則換頁
   If m_nRow > m_ReportDataRows Then
      Printer.NewPage
      m_nPage = m_nPage + 1
      PrintPageHeader m_nPage
      m_nRow = 1
   End If
   
   ' 輸出
   For nIndex = 0 To UBound(strField)
      Select Case nIndex
         Case 1, 2, 3, 5:
            Printer.CurrentX = m_Field(nIndex).Left * m_CharWidth
            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + m_nRow) * m_CharHeight
            Printer.Print LeftStr(strField(nIndex), m_Field(nIndex).Width)
         Case 4:
            Printer.CurrentX = (m_Field(nIndex).Left + m_Field(nIndex).Width - 3) * m_CharWidth - Printer.TextWidth(LeftStr(strField(nIndex), m_Field(nIndex).Width))
            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + m_nRow) * m_CharHeight
            Printer.Print LeftStr(strField(nIndex), m_Field(nIndex).Width)
         Case Else:
            nLeft = m_Field(nIndex).Left + (m_Field(nIndex).Width / 2) - (StrLength(strField(nIndex)) / 2)
            Printer.CurrentX = nLeft * m_CharWidth
            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + m_nRow) * m_CharHeight
            Printer.Print LeftStr(strField(nIndex), m_Field(nIndex).Width)
      End Select
   Next nIndex
   
   m_nRow = m_nRow + 1
End Sub

' 取得國家的代碼
Private Function GetNationCode(ByVal strKey1 As String, ByVal strKey2 As String, ByVal strKey3 As String, ByVal strKey4 As String) As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   GetNationCode = "999"
   Select Case strKey1
      ' 讀取商標基本檔
      Case "T", "TF", "FCT", "CFT":
         strSql = "SELECT TM10 FROM TRADEMARK " & _
                  "WHERE TM01 = '" & strKey1 & "' AND " & _
                        "TM02 = '" & strKey2 & "' AND " & _
                        "TM03 = '" & strKey3 & "' AND " & _
                        "TM04 = '" & strKey4 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If IsNull(rsTmp.Fields("TM10")) = False Then: GetNationCode = rsTmp.Fields("TM10")
         End If
         rsTmp.Close
      ' 讀取專利基本檔
      Case "P", "CFP", "FCP":
         strSql = "SELECT PA09 FROM PATENT " & _
                  "WHERE PA01 = '" & strKey1 & "' AND " & _
                        "PA02 = '" & strKey2 & "' AND " & _
                        "PA03 = '" & strKey3 & "' AND " & _
                        "PA04 = '" & strKey4 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If IsNull(rsTmp.Fields("PA09")) = False Then: GetNationCode = rsTmp.Fields("PA09")
         End If
         rsTmp.Close
      ' 讀取法務基本檔
      Case "L", "CFL", "FCL":
         strSql = "SELECT LC15 FROM LAWCASE " & _
                  "WHERE LC01 = '" & strKey1 & "' AND " & _
                        "LC02 = '" & strKey2 & "' AND " & _
                        "LC03 = '" & strKey3 & "' AND " & _
                        "LC04 = '" & strKey4 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If IsNull(rsTmp.Fields("LC15")) = False Then: GetNationCode = rsTmp.Fields("LC15")
         End If
         rsTmp.Close
      ' 讀取顧問案件基本檔
      Case "LA":
         GetNationCode = "000"
      ' 讀取服務業務基本檔
      Case Else:
         strSql = "SELECT SP09 FROM SERVICEPRACTICE " & _
                  "WHERE SP01 = '" & strKey1 & "' AND " & _
                        "SP02 = '" & strKey2 & "' AND " & _
                        "SP03 = '" & strKey3 & "' AND " & _
                        "SP04 = '" & strKey4 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If IsNull(rsTmp.Fields("SP09")) = False Then: GetNationCode = rsTmp.Fields("SP09")
         End If
         rsTmp.Close
   End Select
   Set rsTmp = Nothing
End Function

' 列印資料
Private Function PrintData(ByRef rsSrcTmp As ADODB.Recordset) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTemp As String
   Dim fld(5) As String
   Dim nType As Integer
   Dim nIndex As Integer
   Dim strNation As String
   Dim strKey1 As String
   Dim strKey2 As String
   Dim strKey3 As String
   Dim strKey4 As String
   Dim strAX205 As String
   Dim bPrintHeader As Boolean
   
   bPrintHeader = False
   PrintData = False
   
   If IsEmptyText(textAX205) = True Then
      strAX205 = "(AX205 = '2201' OR AX205 = '220101' OR AX205 = '220102' OR AX205 = '220103' OR AX205 = '220104' OR AX205 = '220107') "
   Else
      strAX205 = "AX205 = '" & textAX205 & "' "
   End If
      
   ' 紙張大小, 方向
   Select Case m_PaperSize
      Case "A4":
         Printer.PaperSize = vbPRPSA4
         Printer.Orientation = vbPRORLandscape
      Case "REPORT":
         Printer.PaperSize = vbPRPSFanfoldUS
      Case Else:
         Printer.PaperSize = vbPRPSA4
         Printer.Orientation = vbPRORLandscape
   End Select

   Printer.Font.Name = "新細明體"

   ' 印表頭
   m_nPage = 1
   'PrintPageHeader m_nPage

   m_nRow = 1
   Do While rsSrcTmp.EOF = False
      ' 清除欄位
      For nIndex = 0 To 5: fld(nIndex) = Empty: Next nIndex
      
      ' 取得本所案號
      strKey1 = Empty
      strKey2 = Empty
      strKey3 = Empty
      strKey4 = Empty
      If IsNull(rsSrcTmp.Fields("AC1")) = False Then
         strKey1 = rsSrcTmp.Fields("AC1")
         If IsNull(rsSrcTmp.Fields("AC2")) = False Then: strKey2 = rsSrcTmp.Fields("AC2")
         If IsNull(rsSrcTmp.Fields("AC3")) = False Then: strKey3 = rsSrcTmp.Fields("AC3")
         If IsNull(rsSrcTmp.Fields("AC4")) = False Then: strKey4 = rsSrcTmp.Fields("AC4")
      ElseIf IsNull(rsSrcTmp.Fields("CP1")) = False Then
         strKey1 = rsSrcTmp.Fields("CP1")
         If IsNull(rsSrcTmp.Fields("CP2")) = False Then: strKey2 = rsSrcTmp.Fields("CP2")
         If IsNull(rsSrcTmp.Fields("CP3")) = False Then: strKey3 = rsSrcTmp.Fields("CP3")
         If IsNull(rsSrcTmp.Fields("CP4")) = False Then: strKey4 = rsSrcTmp.Fields("CP4")
      End If
      ' 本所案號資料不全不予計入
      If IsEmptyText(strKey1) = True Or IsEmptyText(strKey2) = True Or IsEmptyText(strKey3) = True Or IsEmptyText(strKey4) = True Then
         GoTo NextRecord
      End If
      ' 取得國家代碼 (非台灣的不予計入)
      If GetNationCode(strKey1, strKey2, strKey3, strKey4) >= "010" Then
         GoTo NextRecord
      End If
      
      If IsNull(rsSrcTmp.Fields("AC1")) = False Then
         If IsNull(rsSrcTmp.Fields("CP1")) = False Then
            ' 傳票資料檔與案件進度檔均有
            nType = 0
         Else
            ' 傳票資料檔有而案件進度檔沒有
            nType = 1
         End If
      Else
         If IsNull(rsSrcTmp.Fields("CP1")) = False Then
            ' 案件進度檔有而傳票資料檔沒有
            nType = 2
         Else
            nType = 3
         End If
      End If

      Select Case nType
         Case 0:
            strSql = "SELECT DISTINCT A0205,AX202,AX206,AX212,AX214 FROM ACC020, ACC021 " & _
                     "WHERE A0205 = " & rsSrcTmp.Fields("AC5") & " AND " & strAX205 & " AND " & _
                           "AX202 = A0202 AND " & _
                           "AX214 = '" & rsSrcTmp.Fields("AC1") & rsSrcTmp.Fields("AC2") & rsSrcTmp.Fields("AC3") & rsSrcTmp.Fields("AC4") & "' AND " & _
                           "AX206 IS NOT NULL AND " & _
                           "AX206 <> 0 "
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
               rsTmp.MoveFirst
               
               ' 列印表頭
               If bPrintHeader = False Then
                  PrintPageHeader m_nPage
                  bPrintHeader = True
               End If
               
               Do While rsTmp.EOF = False
                  ' 清除欄位
                  For nIndex = 0 To 5: fld(nIndex) = Empty: Next nIndex
                  ' 傳票日期
                  If IsNull(rsTmp.Fields("A0205")) = False Then
                     'fld(0) = rsTmp.Fields("A0205")
                     fld(0) = TAIWANYEAR(rsTmp.Fields("A0205")) & "/" & TAIWANMONTH(rsTmp.Fields("A0205")) & "/" & TAIWANDAY(rsTmp.Fields("A0205"))
                  End If
                  ' 本所案號
                  If IsNull(rsTmp.Fields("AX214")) = False Then
                     strTemp = rsTmp.Fields("AX214")
                     Select Case Mid(strTemp, 1, Len(strTemp) - 9)
                        Case "TF":
                           fld(1) = Mid(strTemp, 1, Len(strTemp) - 9) & "-" & Mid(strTemp, Len(strTemp) - 9 + 1, 5) & "-" & Mid(strTemp, Len(strTemp) - 4 + 1, 1) & "-" & Mid(strTemp, Len(strTemp) - 3 + 1, 1) & "-" & Mid(strTemp, Len(strTemp) - 2 + 1, 1)
                        Case Else:
                           fld(1) = Mid(strTemp, 1, Len(strTemp) - 9) & "-" & Mid(strTemp, Len(strTemp) - 9 + 1, 6) & "-" & Mid(strTemp, Len(strTemp) - 3 + 1, 1) & "-" & Mid(strTemp, Len(strTemp) - 2 + 1, 1)
                     End Select
                  End If
                  ' 傳票號碼
                  If IsNull(rsTmp.Fields("AX202")) = False Then
                     fld(3) = rsTmp.Fields("AX202")
                  End If
                  ' 規費金額
                  If IsNull(rsTmp.Fields("AX206")) = False Then
                     fld(4) = rsTmp.Fields("AX206")
                  End If
                  ' 摘要
                  If IsNull(rsTmp.Fields("AX212")) = False Then
                     fld(5) = rsTmp.Fields("AX212")
                  End If
                  ' 列印該列
                  PrintRow fld
                  
                  rsTmp.MoveNext
               Loop
            End If
            rsTmp.Close

            Select Case rsSrcTmp.Fields("CP1")
               ' 讀取商標基本檔
               Case "T", "TF", "FCT", "CFT":
                  strSql = "SELECT CP01,CP02,CP03,CP04,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) AS CP10,cp84,NVL(CP27-19110000,NULL) AS CP27,CP64 FROM TRADEMARK, CASEPROPERTYMAP, CASEPROGRESS " & _
                           "WHERE CP01 = '" & rsSrcTmp.Fields("CP1") & "' AND " & _
                                 "CP02 = '" & rsSrcTmp.Fields("CP2") & "' AND " & _
                                 "CP03 = '" & rsSrcTmp.Fields("CP3") & "' AND " & _
                                 "CP04 = '" & rsSrcTmp.Fields("CP4") & "' AND " & _
                                 "CP27 = " & rsSrcTmp.Fields("CP5") & " AND " & _
                                 "CP01 = TM01(+) AND " & _
                                 "CP02 = TM02(+) AND " & _
                                 "CP03 = TM03(+) AND " & _
                                 "CP04 = TM04(+) AND " & _
                                 "cp84 IS NOT NULL AND " & _
                                 "cp84 <> 0 AND " & _
                                 "CP01 = CPM01(+) AND " & _
                                 "CP10 = CPM02(+)"
               ' 讀取專利基本檔
               Case "P", "CFP", "FCP":
                  strSql = "SELECT  CP01,CP02,CP03,CP04,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) AS CP10,cp84,NVL(CP27-19110000,NULL) AS CP27,CP64 FROM PATENT, CASEPROPERTYMAP, CASEPROGRESS " & _
                           "WHERE CP01 = '" & rsSrcTmp.Fields("CP1") & "' AND " & _
                                 "CP02 = '" & rsSrcTmp.Fields("CP2") & "' AND " & _
                                 "CP03 = '" & rsSrcTmp.Fields("CP3") & "' AND " & _
                                 "CP04 = '" & rsSrcTmp.Fields("CP4") & "' AND " & _
                                 "CP27 = " & rsSrcTmp.Fields("CP5") & " AND " & _
                                 "CP01 = PA01(+) AND " & _
                                 "CP02 = PA02(+) AND " & _
                                 "CP03 = PA03(+) AND " & _
                                 "CP04 = PA04(+) AND " & _
                                 "cp84 IS NOT NULL AND " & _
                                 "cp84 <> 0 AND " & _
                                 "CP01 = CPM01(+) AND " & _
                                 "CP10 = CPM02(+)"
               ' 讀取法務基本檔
               Case "L", "CFL", "FCL":
                  strSql = "SELECT  CP01,CP02,CP03,CP04,NVL(DECODE(LC15,'000',CPM03,CPM04),CP10) AS CP10,cp84,NVL(CP27-19110000,NULL) AS CP27,CP64 FROM LAWCASE, CASEPROPERTYMAP, CASEPROGRESS " & _
                           "WHERE CP01 = '" & rsSrcTmp.Fields("CP1") & "' AND " & _
                                 "CP02 = '" & rsSrcTmp.Fields("CP2") & "' AND " & _
                                 "CP03 = '" & rsSrcTmp.Fields("CP3") & "' AND " & _
                                 "CP04 = '" & rsSrcTmp.Fields("CP4") & "' AND " & _
                                 "CP27 = " & rsSrcTmp.Fields("CP5") & " AND " & _
                                 "CP01 = LC01(+) AND " & _
                                 "CP02 = LC02(+) AND " & _
                                 "CP03 = LC03(+) AND " & _
                                 "CP04 = LC04(+) AND " & _
                                 "cp84 IS NOT NULL AND " & _
                                 "cp84 <> 0 AND " & _
                                 "CP01 = CPM01(+) AND " & _
                                 "CP10 = CPM02(+)"
               ' 讀取顧問案件基本檔
               Case "LA":
                  strSql = "SELECT  CP01,CP02,CP03,CP04,NVL(CPM03,CP10) AS CP10,cp84,NVL(CP27-19110000,NULL) AS CP27,CP64 FROM HIRECASE, CASEPROPERTYMAP, CASEPROGRESS " & _
                           "WHERE CP01 = '" & rsSrcTmp.Fields("CP1") & "' AND " & _
                                 "CP02 = '" & rsSrcTmp.Fields("CP2") & "' AND " & _
                                 "CP03 = '" & rsSrcTmp.Fields("CP3") & "' AND " & _
                                 "CP04 = '" & rsSrcTmp.Fields("CP4") & "' AND " & _
                                 "CP27 = " & rsSrcTmp.Fields("CP5") & " AND " & _
                                 "CP01 = HC01(+) AND " & _
                                 "CP02 = HC02(+) AND " & _
                                 "CP03 = HC03(+) AND " & _
                                 "CP04 = HC04(+) AND " & _
                                 "cp84 IS NOT NULL AND " & _
                                 "cp84 <> 0 AND " & _
                                 "CP01 = CPM01(+) AND " & _
                                 "CP10 = CPM02(+)"
               ' 讀取服務業務基本檔
               Case Else:
                  strSql = "SELECT  CP01,CP02,CP03,CP04,NVL(DECODE(SP08,'000',CPM03,CPM04),CP10) AS CP10,cp84,NVL(CP27-19110000,NULL) AS CP27,CP64 FROM SERVICEPRACTICE, CASEPROPERTYMAP, CASEPROGRESS " & _
                           "WHERE CP01 = '" & rsSrcTmp.Fields("CP1") & "' AND " & _
                                 "CP02 = '" & rsSrcTmp.Fields("CP2") & "' AND " & _
                                 "CP03 = '" & rsSrcTmp.Fields("CP3") & "' AND " & _
                                 "CP04 = '" & rsSrcTmp.Fields("CP4") & "' AND " & _
                                 "CP27 = " & rsSrcTmp.Fields("CP5") & " AND " & _
                                 "CP01 = SP01(+) AND " & _
                                 "CP02 = SP02(+) AND " & _
                                 "CP03 = SP03(+) AND " & _
                                 "CP04 = SP04(+) AND " & _
                                 "cp84 IS NOT NULL AND " & _
                                 "cp84 <> 0 AND " & _
                                 "CP01 = CPM01(+) AND " & _
                                 "CP10 = CPM02(+)"
            End Select
            'strSQL = "SELECT CP01,CP02,CP03,CP04,CP10,cp84,CP27,CP64 FROM CASEPROGRESS " & _
            '         "WHERE CP01 = '" & rsSrcTmp.Fields("CP1") & "' AND " & _
            '               "CP02 = '" & rsSrcTmp.Fields("CP2") & "' AND " & _
            '               "CP03 = '" & rsSrcTmp.Fields("CP3") & "' AND " & _
            '               "CP04 = '" & rsSrcTmp.Fields("CP4") & "' AND " & _
            '               "CP27 = " & rsSrcTmp.Fields("CP5") & " "
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
               rsTmp.MoveFirst
               Do While rsTmp.EOF = False
                  ' 清除欄位
                  For nIndex = 0 To 5: fld(nIndex) = Empty: Next nIndex
                  ' 發文日期
                  If IsNull(rsTmp.Fields("CP27")) = False Then
                     'fld(0) = rsTmp.Fields("CP27")
                     fld(0) = TAIWANYEAR(rsTmp.Fields("CP27")) & "/" & TAIWANMONTH(rsTmp.Fields("CP27")) & "/" & TAIWANDAY(rsTmp.Fields("CP27"))
                  End If
                  ' 本所案號
                  If IsNull(rsTmp.Fields("CP01")) = False Then
                     Select Case rsTmp.Fields("CP01")
                        Case "TF":
                           fld(1) = rsTmp.Fields("CP01") & "-" & Mid(rsTmp.Fields("CP02"), 1, 5) & "-" & Mid(rsTmp.Fields("CP02"), 6, 1) & "-" & rsTmp.Fields("CP03") & "-" & rsTmp.Fields("CP04")
                        Case Else:
                           fld(1) = rsTmp.Fields("CP01") & "-" & rsTmp.Fields("CP02") & "-" & rsTmp.Fields("CP03") & "-" & rsTmp.Fields("CP04")
                     End Select
                  End If
                  ' 案件性質
                  If IsNull(rsTmp.Fields("CP10")) = False Then
                     fld(2) = rsTmp.Fields("CP10")
                  End If
                  ' 規費金額
                  If IsNull(rsTmp.Fields("cp84")) = False Then
                     fld(4) = rsTmp.Fields("cp84")
                  End If
                  ' 備註
                  If IsNull(rsTmp.Fields("CP64")) = False Then
                     fld(5) = rsTmp.Fields("CP64")
                  End If
                  ' 列印該列
                  PrintRow fld

                  rsTmp.MoveNext
               Loop
            End If
            rsTmp.Close
            
         Case 1:
            strSql = "SELECT DISTINCT A0205,AX202,AX206,AX212,AX214 FROM ACC020, ACC021 " & _
                     "WHERE A0205 = " & rsSrcTmp.Fields("AC5") & " AND " & strAX205 & " AND " & _
                           "AX202 = A0202 AND " & _
                           "AX214 = '" & rsSrcTmp.Fields("AC1") & rsSrcTmp.Fields("AC2") & rsSrcTmp.Fields("AC3") & rsSrcTmp.Fields("AC4") & "' AND " & _
                           "AX206 IS NOT NULL AND " & _
                           "AX206 <> 0 "
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
               rsTmp.MoveFirst
               
               ' 列印表頭
               If bPrintHeader = False Then
                  PrintPageHeader m_nPage
                  bPrintHeader = True
               End If
               
               Do While rsTmp.EOF = False
                  ' 清除欄位
                  For nIndex = 0 To 5: fld(nIndex) = Empty: Next nIndex
                  ' 傳票日期
                  If IsNull(rsTmp.Fields("A0205")) = False Then
                     'fld(0) = rsTmp.Fields("A0205")
                     fld(0) = TAIWANYEAR(rsTmp.Fields("A0205")) & "/" & TAIWANMONTH(rsTmp.Fields("A0205")) & "/" & TAIWANDAY(rsTmp.Fields("A0205"))
                  End If
                  ' 本所案號
                  If IsNull(rsTmp.Fields("AX214")) = False Then
                     strTemp = rsTmp.Fields("AX214")
                     Select Case Mid(strTemp, 1, Len(strTemp) - 9)
                        Case "TF":
                           fld(1) = Mid(strTemp, 1, Len(strTemp) - 9) & "-" & Mid(strTemp, Len(strTemp) - 9 + 1, 5) & "-" & Mid(strTemp, Len(strTemp) - 4 + 1, 1) & "-" & Mid(strTemp, Len(strTemp) - 3 + 1, 1) & "-" & Mid(strTemp, Len(strTemp) - 2 + 1, 1)
                        Case Else:
                           fld(1) = Mid(strTemp, 1, Len(strTemp) - 9) & "-" & Mid(strTemp, Len(strTemp) - 9 + 1, 6) & "-" & Mid(strTemp, Len(strTemp) - 3 + 1, 1) & "-" & Mid(strTemp, Len(strTemp) - 2 + 1, 1)
                     End Select
                  End If
                  ' 傳票號碼
                  If IsNull(rsTmp.Fields("AX202")) = False Then
                     fld(3) = rsTmp.Fields("AX202")
                  End If
                  ' 規費金額
                  If IsNull(rsTmp.Fields("AX206")) = False Then
                     fld(4) = rsTmp.Fields("AX206")
                  End If
                  ' 摘要
                  If IsNull(rsTmp.Fields("AX212")) = False Then
                     fld(5) = rsTmp.Fields("AX212")
                  End If
                  ' 列印該列
                  PrintRow fld
                  
                  rsTmp.MoveNext
               Loop
            End If
            rsTmp.Close
         Case 2:
            Select Case rsSrcTmp.Fields("CP1")
               ' 讀取商標基本檔
               Case "T", "TF", "FCT", "CFT":
                  strSql = "SELECT  CP01,CP02,CP03,CP04,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) AS CP10,cp84,NVL(CP27-19110000,NULL) AS CP27,CP64 FROM TRADEMARK, CASEPROPERTYMAP, CASEPROGRESS " & _
                           "WHERE CP01 = '" & rsSrcTmp.Fields("CP1") & "' AND " & _
                                 "CP02 = '" & rsSrcTmp.Fields("CP2") & "' AND " & _
                                 "CP03 = '" & rsSrcTmp.Fields("CP3") & "' AND " & _
                                 "CP04 = '" & rsSrcTmp.Fields("CP4") & "' AND " & _
                                 "CP27 = " & rsSrcTmp.Fields("CP5") & " AND " & _
                                 "CP01 = TM01(+) AND " & _
                                 "CP02 = TM02(+) AND " & _
                                 "CP03 = TM03(+) AND " & _
                                 "CP04 = TM04(+) AND " & _
                                 "cp84 IS NOT NULL AND " & _
                                 "cp84 <> 0 AND " & _
                                 "CP01 = CPM01(+) AND " & _
                                 "CP10 = CPM02(+)"
               ' 讀取專利基本檔
               Case "P", "CFP", "FCP":
                  strSql = "SELECT  CP01,CP02,CP03,CP04,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) AS CP10,cp84,NVL(CP27-19110000,NULL) AS CP27,CP64 FROM PATENT, CASEPROPERTYMAP, CASEPROGRESS " & _
                           "WHERE CP01 = '" & rsSrcTmp.Fields("CP1") & "' AND " & _
                                 "CP02 = '" & rsSrcTmp.Fields("CP2") & "' AND " & _
                                 "CP03 = '" & rsSrcTmp.Fields("CP3") & "' AND " & _
                                 "CP04 = '" & rsSrcTmp.Fields("CP4") & "' AND " & _
                                 "CP27 = " & rsSrcTmp.Fields("CP5") & " AND " & _
                                 "CP01 = PA01(+) AND " & _
                                 "CP02 = PA02(+) AND " & _
                                 "CP03 = PA03(+) AND " & _
                                 "CP04 = PA04(+) AND " & _
                                 "cp84 IS NOT NULL AND " & _
                                 "cp84 <> 0 AND " & _
                                 "CP01 = CPM01(+) AND " & _
                                 "CP10 = CPM02(+)"
               ' 讀取法務基本檔
               Case "L", "CFL", "FCL":
                  strSql = "SELECT  CP01,CP02,CP03,CP04,NVL(DECODE(LC15,'000',CPM03,CPM04),CP10) AS CP10,cp84,NVL(CP27-19110000,NULL) AS CP27,CP64 FROM LAWCASE, CASEPROPERTYMAP, CASEPROGRESS " & _
                           "WHERE CP01 = '" & rsSrcTmp.Fields("CP1") & "' AND " & _
                                 "CP02 = '" & rsSrcTmp.Fields("CP2") & "' AND " & _
                                 "CP03 = '" & rsSrcTmp.Fields("CP3") & "' AND " & _
                                 "CP04 = '" & rsSrcTmp.Fields("CP4") & "' AND " & _
                                 "CP27 = " & rsSrcTmp.Fields("CP5") & " AND " & _
                                 "CP01 = LC01(+) AND " & _
                                 "CP02 = LC02(+) AND " & _
                                 "CP03 = LC03(+) AND " & _
                                 "CP04 = LC04(+) AND " & _
                                 "cp84 IS NOT NULL AND " & _
                                 "cp84 <> 0 AND " & _
                                 "CP01 = CPM01(+) AND " & _
                                 "CP10 = CPM02(+)"
               ' 讀取顧問案件基本檔
               Case "LA":
                  strSql = "SELECT  CP01,CP02,CP03,CP04,NVL(CPM03,CP10) AS CP10,cp84,NVL(CP27-19110000,NULL) AS CP27,CP64 FROM HIRECASE, CASEPROPERTYMAP, CASEPROGRESS " & _
                           "WHERE CP01 = '" & rsSrcTmp.Fields("CP1") & "' AND " & _
                                 "CP02 = '" & rsSrcTmp.Fields("CP2") & "' AND " & _
                                 "CP03 = '" & rsSrcTmp.Fields("CP3") & "' AND " & _
                                 "CP04 = '" & rsSrcTmp.Fields("CP4") & "' AND " & _
                                 "CP27 = " & rsSrcTmp.Fields("CP5") & " AND " & _
                                 "CP01 = HC01(+) AND " & _
                                 "CP02 = HC02(+) AND " & _
                                 "CP03 = HC03(+) AND " & _
                                 "CP04 = HC04(+) AND " & _
                                 "cp84 IS NOT NULL AND " & _
                                 "cp84 <> 0 AND " & _
                                 "CP01 = CPM01(+) AND " & _
                                 "CP01 = CPM01(+) AND " & _
                                 "CP10 = CPM02(+)"
               ' 讀取服務業務基本檔
               Case Else:
                  strSql = "SELECT  CP01,CP02,CP03,CP04,NVL(DECODE(SP08,'000',CPM03,CPM04),CP10) AS CP10,cp84,NVL(CP27-19110000,NULL) AS CP27,CP64 FROM SERVICEPRACTICE, CASEPROPERTYMAP, CASEPROGRESS " & _
                           "WHERE CP01 = '" & rsSrcTmp.Fields("CP1") & "' AND " & _
                                 "CP02 = '" & rsSrcTmp.Fields("CP2") & "' AND " & _
                                 "CP03 = '" & rsSrcTmp.Fields("CP3") & "' AND " & _
                                 "CP04 = '" & rsSrcTmp.Fields("CP4") & "' AND " & _
                                 "CP27 = " & rsSrcTmp.Fields("CP5") & " AND " & _
                                 "CP01 = SP01(+) AND " & _
                                 "CP02 = SP02(+) AND " & _
                                 "CP03 = SP03(+) AND " & _
                                 "CP04 = SP04(+) AND " & _
                                 "cp84 IS NOT NULL AND " & _
                                 "cp84 <> 0 AND " & _
                                 "CP01 = CPM01(+) AND " & _
                                 "CP10 = CPM02(+)"
            End Select
            'strSQL = "SELECT CP01,CP02,CP03,CP04,CP10,cp84,CP27,CP64 FROM CASEPROGRESS " & _
            '         "WHERE CP01 = '" & rsSrcTmp.Fields("CP1") & "' AND " & _
            '               "CP02 = '" & rsSrcTmp.Fields("CP2") & "' AND " & _
            '               "CP03 = '" & rsSrcTmp.Fields("CP3") & "' AND " & _
            '               "CP04 = '" & rsSrcTmp.Fields("CP4") & "' AND " & _
            '               "CP27 = " & rsSrcTmp.Fields("CP5") & " "
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
               rsTmp.MoveFirst
               
               ' 列印表頭
               If bPrintHeader = False Then
                  PrintPageHeader m_nPage
                  bPrintHeader = True
               End If
               
               Do While rsTmp.EOF = False
                  ' 清除欄位
                  For nIndex = 0 To 5: fld(nIndex) = Empty: Next nIndex
                  ' 發文日期
                  If IsNull(rsTmp.Fields("CP27")) = False Then
                     'fld(0) = rsTmp.Fields("CP27")
                     fld(0) = TAIWANYEAR(rsTmp.Fields("CP27")) & "/" & TAIWANMONTH(rsTmp.Fields("CP27")) & "/" & TAIWANDAY(rsTmp.Fields("CP27"))
                  End If
                  ' 本所案號
                  If IsNull(rsTmp.Fields("CP01")) = False Then
                     Select Case rsTmp.Fields("CP01")
                        Case "TF":
                           fld(1) = rsTmp.Fields("CP01") & "-" & Mid(rsTmp.Fields("CP02"), 1, 5) & "-" & Mid(rsTmp.Fields("CP02"), 6, 1) & "-" & rsTmp.Fields("CP03") & "-" & rsTmp.Fields("CP04")
                        Case Else:
                           fld(1) = rsTmp.Fields("CP01") & "-" & rsTmp.Fields("CP02") & "-" & rsTmp.Fields("CP03") & "-" & rsTmp.Fields("CP04")
                     End Select
                  End If
                  ' 案件性質
                  If IsNull(rsTmp.Fields("CP10")) = False Then
                     fld(2) = rsTmp.Fields("CP10")
                  End If
                  ' 規費金額
                  If IsNull(rsTmp.Fields("cp84")) = False Then
                     fld(4) = rsTmp.Fields("cp84")
                  End If
                  ' 備註
                  If IsNull(rsTmp.Fields("CP64")) = False Then
                     fld(5) = rsTmp.Fields("CP64")
                  End If
                  ' 列印該列
                  PrintRow fld
                  
                  rsTmp.MoveNext
               Loop
            End If
            rsTmp.Close
         Case Else:
      End Select
NextRecord:
      rsSrcTmp.MoveNext
   Loop
   
   ' 列印表頭
   If bPrintHeader = True Then
      PrintTerminateLine m_HeaderHeight + m_nRow
      m_nRow = m_nRow + 1
      
      Printer.EndDoc
   End If
   
   PrintData = bPrintHeader
End Function

Public Function LeftStr(ByVal strData As String, ByVal nLen As Integer) As String
   LeftStr = StrConv(MidB(StrConv(strData, vbFromUnicode), 1, nLen), vbUnicode)
End Function

