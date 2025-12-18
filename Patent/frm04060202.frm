VERSION 5.00
Begin VB.Form frm04060202 
   BorderStyle     =   1  '單線固定
   Caption         =   "大陸專利市場佔有率統計表"
   ClientHeight    =   1905
   ClientLeft      =   630
   ClientTop       =   4530
   ClientWidth     =   4770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4770
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   1260
      TabIndex        =   3
      Top             =   1440
      Width           =   3252
   End
   Begin VB.CommandButton bottonOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3012
      TabIndex        =   4
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton bottonExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3840
      TabIndex        =   5
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox text02 
      Height          =   264
      Left            =   1260
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1080
      Width           =   372
   End
   Begin VB.TextBox text01_02 
      Height          =   264
      Left            =   3030
      MaxLength       =   7
      TabIndex        =   1
      Top             =   720
      Width           =   1452
   End
   Begin VB.TextBox text01_01 
      Height          =   264
      Left            =   1260
      MaxLength       =   7
      TabIndex        =   0
      Top             =   720
      Width           =   1452
   End
   Begin VB.Label Label4 
      Caption         =   "印表機："
      Height          =   252
      Left            =   180
      TabIndex        =   9
      Top             =   1440
      Width           =   732
   End
   Begin VB.Label Label3 
      Caption         =   "( 1.市場統計   2.本所案件統計 )"
      Height          =   255
      Left            =   1860
      TabIndex        =   8
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "出表選擇："
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   180
      TabIndex        =   7
      Top             =   1080
      Width           =   972
   End
   Begin VB.Line Line1 
      X1              =   2820
      X2              =   2940
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "公告日："
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   180
      TabIndex        =   6
      Top             =   720
      Width           =   852
   End
End
Attribute VB_Name = "frm04060202"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
' 大陸公報市場佔有率統計表
Option Explicit
Const m_CharWidth = 120
Const m_CharHeight = 240
Const m_PaperSize = "REPORT"

' 宣告報表表頭的欄位其資料型態
Private Type REPORTFIELD
   Name As String
   DataCode As String
   DataName As String
   Left As Long
   Width As Long
End Type
' 表頭欄位的內容
Dim m_Field(17) As REPORTFIELD
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

' 宣告代理人項目的資料型態
Private Type AGENTITEM
   AgentCode As String
   AgentName As String
   Count As Integer
   Type1 As Integer
   Type2 As Integer
   Type3 As Integer
End Type
' 代理事務所串列
Dim m_AgentList() As AGENTITEM
' 代理人串列中的資料個數
Dim m_AgentCount As Integer
' 預設印表機
Dim m_DefaultPrinter As String

Private Sub ClearFields()
   text01_01 = Empty
   text01_02 = Empty
   text02 = Empty
End Sub

Private Sub bottonOK_Click()
   Dim Prn As Printer
   
   '搜尋 Printer
   For Each Prn In Printers
      If Prn.DeviceName = cmbPrinter.Text Then
         Set Printer = Prn
         Exit For
      End If
   Next
   
   If CheckDataValid = False Then
      GoTo EXITSUB
   End If
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/2 清除查詢印表記錄檔欄位
   If Len(text01_01) <> 0 Or Len(text01_02) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1 & text01_01 & "-" & text01_02 'Add By Sindy 2010/12/2
   End If
   Select Case text02
      Case "1":
         pub_QL05 = pub_QL05 & ";" & Label2 & "1.市場統計" 'Add By Sindy 2010/12/2
         Print_RP1
      Case "2"
         pub_QL05 = pub_QL05 & ";" & Label2 & "2.本所案件統計" 'Add By Sindy 2010/12/2
         Print_RP2
   End Select
   ' 清除畫面
   ClearFields
   ' 將Focus移至第一個欄位
   text01_01.SetFocus
EXITSUB:
End Sub

Private Sub Form_Load()
   Dim Prn As Printer
   Dim nIndex As Integer
   Dim nSel As Integer
      
   m_DefaultPrinter = Printer.DeviceName
   MoveFormToCenter Me
   
   nSel = 0
   nIndex = 0
   'For Each Prn In Printers
   '   cmbPrinter.AddItem Prn.DeviceName
   '   If Prn.DeviceName = strDeviceName Then
   '      nSel = nIndex
   '   End If
   '   nIndex = nIndex + 1
   'Next
   'cmbPrinter.ListIndex = nSel
   For Each Prn In Printers
      If Prn.DeviceName <> m_DefaultPrinter Then
         cmbPrinter.AddItem Prn.DeviceName
      End If
   Next
   If cmbPrinter.ListCount > 0 Then: cmbPrinter.ListIndex = 0
   
End Sub

Private Sub bottonExit_Click()
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim Prn As Printer
   '搜尋 Printer
   For Each Prn In Printers
      If Prn.DeviceName = m_DefaultPrinter Then
         Set Printer = Prn
         Exit For
      End If
   Next
   'Add By Cheng 2002/07/18
   Set frm04060202 = Nothing
End Sub

Private Sub text01_01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmpty(text01_01) = False Then
      If CheckIsTaiwanDate(text01_01, False) = False Then
         Cancel = True
         strMsg = "請輸入正確的公告日"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         InverseAll text01_01
      End If
   Else
      Cancel = True
      strMsg = "公告日必須輸入"
      strTit = "檢核輸入"
      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
   End If
   If Cancel Then TextInverse text01_01
End Sub

Private Sub text01_02_LostFocus()
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   If IsEmpty(text01_02) = False Then
      If CheckIsTaiwanDate(text01_02, False) = False Then
         strMsg = "請輸入正確的公告日 !"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         text01_02.SetFocus
         TextInverse text01_02
      Else
         If Not ChkRange(text01_01, text01_02, "公告日") Then
         
         End If
      End If
   Else
      strMsg = "公告日必須輸入"
      strTit = "檢核輸入"
      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      text01_02.SetFocus
   End If
End Sub

Private Sub text02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Select Case text02.Text
      Case 1:
         Cancel = False
      Case 2:
         Cancel = False
      Case Else
         Cancel = True
         strMsg = "請選擇出表的方式 1 或是 2"
         strTit = "出表選擇"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         InverseAll text02
   End Select
End Sub

' 由事務所代碼取得事務所的名稱
Public Function GetAgentCompany(ByVal strAgent As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   strSql = "SELECT * FROM CAGENT WHERE FNM01 = '" & strAgent & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      GetAgentCompany = rsTmp.Fields("FNM02")
   Else
      GetAgentCompany = Empty
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function
' 清除資料
Public Sub Clear()
   Dim nCount As Integer
   
   For nCount = 0 To 16
      m_Field(nCount).Name = Empty
      m_Field(nCount).DataCode = Empty
      m_Field(nCount).DataName = Empty
      m_Field(nCount).Left = 0
      m_Field(nCount).Width = 0
   Next nCount
   
   If m_AgentCount > 0 Then
      Erase m_AgentList
   End If
   m_AgentCount = 0
End Sub

' 取得較小的值
Public Function Min(ByVal nValue1 As Integer, ByVal nValue2 As Integer) As Integer
   If nValue2 < nValue1 Then
      Min = nValue2
   Else
      Min = nValue1
   End If
End Function

' 判斷資料是否為空的
Public Function IsEmpty(ByVal strData As String) As Boolean
   Dim nIndex As Integer
   IsEmpty = False
   
   If Len(strData) <= 0 Then
      IsEmpty = True
   Else
      IsEmpty = True
      For nIndex = 1 To Len(strData)
         If Mid(strData, nIndex, 1) <> " " Then
            IsEmpty = False
            Exit For
         End If
      Next nIndex
   End If
End Function

' 取得資料庫中的資料
Private Function GetDBData_RP(ByVal nReport As Integer) As Boolean
   Dim rsMain As New ADODB.Recordset
   Dim rsTmp As ADODB.Recordset
   Dim strSql As String
   Dim strSubSQL As String
   Dim strTmp As String
   Dim strAgent As String
   Dim nAgentIndex, nCount As Integer
   Dim bFindAgent As Boolean
   Dim nType As Integer
   Dim nSortX, nSortY As Integer
   Dim agentTemp As AGENTITEM
   Dim bProcess As Boolean
   
   m_AgentCount = 0
   GetDBData_RP = True
   
   ' 產生SQL查詢語法
   strSql = "SELECT * FROM CPBulletin "
   strSubSQL = Empty
   If IsEmpty(text01_01) = False Then
      If strSubSQL <> Empty Then: strSubSQL = strSubSQL & "AND "
      strSubSQL = strSubSQL & "CPB03 >= " & ChangeTStringToWString(text01_01) & " "
   End If
   If IsEmpty(text01_02) = False Then
      If strSubSQL <> Empty Then: strSubSQL = strSubSQL & "AND "
      strSubSQL = strSubSQL & "CPB03 <= " & ChangeTStringToWString(text01_02) & " "
   End If
   If strSubSQL <> Empty Then
      strSql = strSql & " WHERE " & strSubSQL
   End If
                           
   ' 取得資料庫的資料
   rsMain.CursorLocation = adUseClient
   rsMain.Open strSql, cnnConnection, adOpenDynamic
   ' 無資料則離開
   If rsMain.RecordCount <= 0 Then
      GetDBData_RP = False
      GoTo EXITSUB
   End If
   
   rsMain.MoveFirst
   ' 依序從資料記錄中取出欄位的內容
   While Not rsMain.EOF
      ' 事務所代號
      strAgent = rsMain.Fields("CPB06")
      ' 申請案號
      strTmp = rsMain.Fields("CPB01")
      
      bProcess = True
      ' 當產生的報表為表二時則需判斷該筆是否計入
      If nReport = 2 Then
         Set rsTmp = New ADODB.Recordset
         strSql = "SELECT * FROM Patent " & _
                  "WHERE PA11 = '" & strTmp & "'"
         ' 查詢專利檔
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenDynamic
         ' 無資料則表示此筆不為台一的案件
         If rsTmp.RecordCount <= 0 Then
            bProcess = False
         Else
            rsTmp.MoveFirst
            ' 若國籍不為大錄表示資料有誤, 不予計入
            If rsTmp.Fields("PA09") <> "020" Then
               bProcess = False
            End If
         End If
         rsTmp.Close
         Set rsTmp = Nothing
      End If
      
      ' 檢查申請案號的種類是屬於 發明, 新型還是設計
      nType = 0
      Select Case Mid(rsMain.Fields("CPB01"), 3, 1)
         Case "1": nType = 1
         Case "2": nType = 2
         Case "3": nType = 3
      End Select
      
      If bProcess = True Then
         ' 搜尋事務所代號串列
         If strAgent <> Empty Then
            bFindAgent = False
            If m_AgentCount > 0 Then
               For nAgentIndex = 0 To m_AgentCount - 1
                  If m_AgentList(nAgentIndex).AgentCode = strAgent Then
                     bFindAgent = True
                     ' 總數量加一
                     m_AgentList(nAgentIndex).Count = m_AgentList(nAgentIndex).Count + 1
                     ' 依照發明或新型或設計來累計
                     Select Case nType
                        Case 1: m_AgentList(nAgentIndex).Type1 = m_AgentList(nAgentIndex).Type1 + 1
                        Case 2: m_AgentList(nAgentIndex).Type2 = m_AgentList(nAgentIndex).Type2 + 1
                        Case 3: m_AgentList(nAgentIndex).Type3 = m_AgentList(nAgentIndex).Type3 + 1
                     End Select
                     Exit For
                  End If
               Next nAgentIndex
            End If
            ' 無此代理事務所的資料時則新產生一個
            If bFindAgent = False Then
               If m_AgentCount = 0 Then
                  nAgentIndex = 0
               Else
                  nAgentIndex = UBound(m_AgentList)
               End If
               ReDim Preserve m_AgentList(nAgentIndex + 1)
               m_AgentCount = m_AgentCount + 1
               m_AgentList(nAgentIndex).AgentCode = strAgent
               m_AgentList(nAgentIndex).AgentName = GetAgentCompany(strAgent)
               m_AgentList(nAgentIndex).Count = 1
               m_AgentList(nAgentIndex).Type1 = 0
               m_AgentList(nAgentIndex).Type2 = 0
               m_AgentList(nAgentIndex).Type3 = 0
               ' 發明或新型或設計
               Select Case nType
                  Case 1: m_AgentList(nAgentIndex).Type1 = m_AgentList(nAgentIndex).Type1 + 1
                  Case 2: m_AgentList(nAgentIndex).Type2 = m_AgentList(nAgentIndex).Type2 + 1
                  Case 3: m_AgentList(nAgentIndex).Type3 = m_AgentList(nAgentIndex).Type3 + 1
               End Select
            End If
         End If
         
      End If
      
      rsMain.MoveNext
   Wend
      
   ' 對事務所串列依數量的多寡由大到小排序
   If m_AgentCount > 0 Then
      For nSortX = 0 To m_AgentCount - 1
         For nSortY = nSortX To m_AgentCount - 1
            If m_AgentList(nSortX).Count < m_AgentList(nSortY).Count Then
               agentTemp = m_AgentList(nSortX)
               m_AgentList(nSortX) = m_AgentList(nSortY)
               m_AgentList(nSortY) = agentTemp
            ' 若資料數相同時, 依照代號來排序
            ElseIf m_AgentList(nSortX).Count = m_AgentList(nSortY).Count Then
               If m_AgentList(nSortX).AgentCode < m_AgentList(nSortY).AgentCode Then
                  agentTemp = m_AgentList(nSortX)
                  m_AgentList(nSortX) = m_AgentList(nSortY)
                  m_AgentList(nSortY) = agentTemp
               End If
            End If
         Next nSortY
      Next nSortX
   End If
                           
EXITSUB:
   rsMain.Close
   Set rsMain = Nothing
End Function

' 取得所事務所內的資料數量
' Input : nType
'         0 : 表全部
'         1 : 表發明
'         2 : 表新型
'         3 : 表設計
Public Function GetAllAgentAmount(ByVal nType As Integer) As Double
   Dim nCount As Integer
   Dim nAmount As Integer
   
   nAmount = 0
   If m_AgentCount > 0 Then
      For nCount = 0 To UBound(m_AgentList) - 1
         Select Case nType
            Case 0:
               nAmount = nAmount + m_AgentList(nCount).Count
            Case 1:
               nAmount = nAmount + m_AgentList(nCount).Type1
            Case 2:
               nAmount = nAmount + m_AgentList(nCount).Type2
            Case 3:
               nAmount = nAmount + m_AgentList(nCount).Type3
         End Select
      Next nCount
   End If
   
   GetAllAgentAmount = nAmount
End Function

' 設定報表欄位的左方位置及其名稱
Public Sub BuildField_RP(ByVal nReport As Integer)
   Dim nIndex As Integer
   Dim nFieldWidth As Integer
   
   Select Case m_PaperSize
      Case "REPORT"
         m_LeftMargin = 1
         m_TopMargin = 5
         m_ReportWidth = 154
         m_ReportDataRows = 45
         nFieldWidth = 9
      Case Else
         m_LeftMargin = 10
         m_TopMargin = 5
         m_ReportWidth = 120
         m_ReportDataRows = 27
         nFieldWidth = 7
   End Select
   
   For nIndex = 0 To 15
      m_Field(nIndex).Width = nFieldWidth - 1
      m_Field(nIndex).Left = m_LeftMargin + (nIndex * nFieldWidth)
      Select Case nIndex
         Case 0:
            m_Field(nIndex).Name = "排名"
            m_Field(nIndex).DataName = "事務所"
         Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14:
            m_Field(nIndex).Name = CStr(nIndex)
            If nIndex <= m_AgentCount Then
               m_Field(nIndex).DataCode = m_AgentList(nIndex - 1).AgentCode
               m_Field(nIndex).DataName = m_AgentList(nIndex - 1).AgentName
            End If
         Case 15:
            m_Field(nIndex).Name = "總計"
      End Select
   Next nIndex
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

' 列印表頭
Public Sub PrintPageHeader_RP(ByVal nReport As Integer, ByVal nPage As Integer)
   Dim nCount As Integer
   Dim strDate1 As String
   Dim StrDate2 As String
   Dim nIndex As Integer
   Dim nRow As Integer
   Dim nX As Long
   Dim ny As Long
   Dim nCenter As Long
   Dim strTemp As String
   
   ' 公告日 (起)
   strDate1 = text01_01
   If IsEmpty(strDate1) = True Then
      strDate1 = "        "
   Else
      strDate1 = ChangeTStringToTDateString(strDate1)
   End If
   ' 公告日 (迄)
   StrDate2 = text01_02
   If IsEmpty(StrDate2) = True Then
      StrDate2 = "        "
   Else
      StrDate2 = ChangeTStringToTDateString(StrDate2)
   End If
   
   ' 表頭
   nRow = 0
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.FontSize = 24
   Printer.Font.Underline = True
   Select Case nReport
      Case 1:
         nX = m_LeftMargin + m_ReportWidth / 2 - 18
         Printer.CurrentX = nX * m_CharWidth
         Printer.Print "大陸專利市場統計表"
      Case 2:
         nX = m_LeftMargin + m_ReportWidth / 2 - 26
         Printer.CurrentX = nX * m_CharWidth
         Printer.Print "大陸專利市場本所案件統計表"
   End Select
   
   nRow = 3
   nX = m_LeftMargin + m_ReportWidth / 2 - 12
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.FontSize = 12
   Printer.Font.Underline = False
   Printer.Print "公告日 : " & strDate1 & " - " & StrDate2
   
   nRow = nRow + 1
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "列印人 : " & strUserName
   
   nX = m_LeftMargin + m_ReportWidth - 20
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "製表日期 : " & Format(ChangeWStringToWDateString(GetTodayDate), "EE/MM/DD")
   
   nRow = nRow + 1
   ' 頁
   nX = m_LeftMargin + m_ReportWidth - 20
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "頁"
   ' 次
   nX = m_LeftMargin + m_ReportWidth - 14
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "次 : " & nPage
      
   nRow = nRow + 1
   ' 列印分隔線
   PrintSplitLine nRow
      
   nRow = nRow + 1
   For nIndex = 0 To 15
      nCenter = ((m_Field(nIndex).Left * m_CharWidth) + (m_Field(nIndex).Left + m_Field(nIndex).Width) * m_CharWidth) / 2
      strTemp = LeftStr(m_Field(nIndex).Name, m_Field(nIndex).Width)
      Printer.CurrentX = nCenter - Printer.TextWidth(strTemp) / 2
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print strTemp
   Next nIndex
   nRow = nRow + 1
   Printer.FontSize = 8
   For nIndex = 0 To 15
      nCenter = ((m_Field(nIndex).Left * m_CharWidth) + (m_Field(nIndex).Left + m_Field(nIndex).Width) * m_CharWidth) / 2
      strTemp = LeftStr(m_Field(nIndex).DataName, 12)
      Printer.CurrentX = nCenter - Printer.TextWidth(strTemp) / 2
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print strTemp
   Next nIndex
   Printer.FontSize = 12
   ' 列印分隔線
   nRow = nRow + 1
   For nX = 0 To 15
      For ny = m_Field(nX).Left To m_Field(nX).Left + m_Field(nX).Width - 1
         Printer.CurrentX = ny * m_CharWidth
         Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
         Printer.Print "-"
      Next ny
   Next nX
   
   m_HeaderHeight = nRow
End Sub

' 列印表一的內容
Public Sub Generate_RP1()
   Dim nRow As Integer
   Dim nPage As Integer
   Dim fld(16) As String
   Dim nAgentCount As Integer
   Dim bFindAgent As Boolean
   Dim nType As Integer
   Dim nAmount As Integer
   Dim nNoAgentAmount As Integer
   Dim nTotalAmount As Integer
   Dim nCount As Integer
   Dim fValue As Single
   Dim nX, ny As Integer
   Dim nRight As Long
   Dim strTemp As String
   
   ' 紙張大小
   Select Case m_PaperSize
      Case "A4":
         Printer.PaperSize = vbPRPSA4
         Printer.Orientation = vbPRORLandscape
      Case "REPORT":
         Printer.PaperSize = vbPRPSFanfoldUS
      Case Else
         Printer.PaperSize = vbPRPSA4
         Printer.Orientation = vbPRORLandscape
   End Select
   ' 紙張方向
   
   nTotalAmount = GetAllAgentAmount(0)
   
   ' 印第一頁的表頭
   nPage = 1
   PrintPageHeader_RP 1, nPage
   
   nRow = 1
   For nType = 1 To 5
      ' 清除所有欄位
      For nAgentCount = 0 To 15
         fld(nAgentCount) = Empty
      Next nAgentCount
      
      ' 第一個欄位
      Select Case nType
         Case 1: fld(0) = "發明"
         Case 2: fld(0) = "新型"
         Case 3: fld(0) = "設計"
         Case 4: fld(0) = "  大陸"
         Case 5: fld(0) = "佔專利%"
      End Select
      ' 第2 ~ 15 的欄位
      If m_AgentCount > 0 Then
         For nAgentCount = 0 To Min(13, m_AgentCount - 1)
            Select Case nType
               Case 1:
                  nAmount = m_AgentList(nAgentCount).Type1
                  fld(nAgentCount + 1) = CStr(nAmount)
               Case 2:
                  nAmount = m_AgentList(nAgentCount).Type2
                  fld(nAgentCount + 1) = CStr(nAmount)
               Case 3:
                  nAmount = m_AgentList(nAgentCount).Type3
                  fld(nAgentCount + 1) = CStr(nAmount)
               Case 4:
                  nAmount = m_AgentList(nAgentCount).Type1 + m_AgentList(nAgentCount).Type2 + m_AgentList(nAgentCount).Type3
                  fld(nAgentCount + 1) = CStr(nAmount)
               Case 5:
                  fValue = (m_AgentList(nAgentCount).Count * 100) / nTotalAmount
                  fld(nAgentCount + 1) = Format(fValue, "##0.00")
            End Select
         Next nAgentCount
      End If
      ' 第16個欄位 總計
      Select Case nType
         Case 1:
            nAmount = GetAllAgentAmount(1)
            fld(15) = CStr(nAmount)
         Case 2:
            nAmount = GetAllAgentAmount(2)
            fld(15) = CStr(nAmount)
         Case 3:
            nAmount = GetAllAgentAmount(3)
            fld(15) = CStr(nAmount)
         Case 4:
            nAmount = GetAllAgentAmount(0)
            fld(15) = CStr(nAmount)
      End Select
      ' 列印欄位內容
      For nAgentCount = 0 To 15
         Printer.CurrentX = m_Field(nAgentCount).Left * m_CharWidth
         strTemp = fld(nAgentCount)
         If nAgentCount > 0 Then
            nRight = (m_Field(nAgentCount).Left + m_Field(nAgentCount).Width - 2) * m_CharWidth
            Printer.CurrentX = nRight - Printer.TextWidth(strTemp)
         End If
         Select Case nType
            Case 1, 2, 3
               ny = m_TopMargin + m_HeaderHeight + nType
            Case 4
               ny = m_TopMargin + m_HeaderHeight + 5
            Case 5
               ny = m_TopMargin + m_HeaderHeight + 7
         End Select
         Printer.CurrentY = ny * m_CharHeight
         'Printer.Print fld(nAgentCount)
         Printer.Print strTemp
      Next nAgentCount
      ' 列印分隔線
      Select Case nType
         Case 3
            ny = m_HeaderHeight + nType + 1
            PrintSplitLine (ny)
         Case 4
            ny = m_HeaderHeight + 6
            PrintSplitLine (ny)
         Case 5
            ny = m_HeaderHeight + 8
            PrintSplitLine (ny)
      End Select
   Next nType
   Printer.EndDoc
End Sub

' 列印表二的內容
Public Sub Generate_RP2()
   Dim nRow As Integer
   Dim nPage As Integer
   Dim fld(16) As String
   Dim nAgentCount As Integer
   Dim bFindAgent As Boolean
   Dim nType As Integer
   Dim nAmount As Integer
   Dim nNoAgentAmount As Integer
   Dim nTotalAmount As Integer
   Dim nCount As Integer
   Dim fValue As Single
   Dim nX, ny As Integer
   Dim nRight As Long
   Dim strTemp As String
   
   ' 紙張大小
   Select Case m_PaperSize
      Case "A4":
         Printer.PaperSize = vbPRPSA4
         Printer.Orientation = vbPRORLandscape
      Case "REPORT":
         Printer.PaperSize = vbPRPSFanfoldUS
      Case Else
         Printer.PaperSize = vbPRPSA4
         Printer.Orientation = vbPRORLandscape
   End Select
   ' 紙張方向
   
   nTotalAmount = GetAllAgentAmount(0)
   
   ' 印第一頁的表頭
   nPage = 1
   PrintPageHeader_RP 2, nPage
   
   nRow = 1
   For nType = 1 To 5
      ' 清除所有欄位
      For nAgentCount = 0 To 15
         fld(nAgentCount) = Empty
      Next nAgentCount
      
      ' 第一個欄位
      Select Case nType
         Case 1: fld(0) = "發明"
         Case 2: fld(0) = "新型"
         Case 3: fld(0) = "設計"
         Case 4: fld(0) = "  大陸"
         Case 5: fld(0) = "佔專利%"
      End Select
      ' 第2 ~ 15 的欄位
      If m_AgentCount > 0 Then
         For nAgentCount = 0 To Min(13, m_AgentCount - 1)
            Select Case nType
               Case 1:
                  nAmount = m_AgentList(nAgentCount).Type1
                  fld(nAgentCount + 1) = CStr(nAmount)
               Case 2:
                  nAmount = m_AgentList(nAgentCount).Type2
                  fld(nAgentCount + 1) = CStr(nAmount)
               Case 3:
                  nAmount = m_AgentList(nAgentCount).Type3
                  fld(nAgentCount + 1) = CStr(nAmount)
               Case 4:
                  nAmount = m_AgentList(nAgentCount).Type1 + m_AgentList(nAgentCount).Type2 + m_AgentList(nAgentCount).Type3
                  fld(nAgentCount + 1) = CStr(nAmount)
               Case 5:
                  fValue = (m_AgentList(nAgentCount).Count * 100) / nTotalAmount
                  fld(nAgentCount + 1) = Format(fValue, "##0.00")
            End Select
         Next nAgentCount
      End If
      ' 第16個欄位 總計
      Select Case nType
         Case 1:
            nAmount = GetAllAgentAmount(1)
            fld(15) = CStr(nAmount)
         Case 2:
            nAmount = GetAllAgentAmount(2)
            fld(15) = CStr(nAmount)
         Case 3:
            nAmount = GetAllAgentAmount(3)
            fld(15) = CStr(nAmount)
         Case 4:
            nAmount = GetAllAgentAmount(0)
            fld(15) = CStr(nAmount)
      End Select
      ' 列印欄位內容
      For nAgentCount = 0 To 15
         Printer.CurrentX = m_Field(nAgentCount).Left * m_CharWidth
         strTemp = fld(nAgentCount)
         If nAgentCount > 0 Then
            nRight = (m_Field(nAgentCount).Left + m_Field(nAgentCount).Width - 2) * m_CharWidth
            Printer.CurrentX = nRight - Printer.TextWidth(strTemp)
         End If
         Select Case nType
            Case 1, 2, 3
               ny = m_TopMargin + m_HeaderHeight + nType
            Case 4
               ny = m_TopMargin + m_HeaderHeight + 5
            Case 5
               ny = m_TopMargin + m_HeaderHeight + 7
         End Select
         Printer.CurrentY = ny * m_CharHeight
         Printer.Print strTemp
      Next nAgentCount
      ' 列印分隔線
      Select Case nType
         Case 3
            ny = m_HeaderHeight + nType + 1
            PrintSplitLine (ny)
         Case 4
            ny = m_HeaderHeight + 6
            PrintSplitLine (ny)
         Case 5
            ny = m_HeaderHeight + 8
            PrintSplitLine (ny)
      End Select
   Next nType
   
   Printer.EndDoc
End Sub

Private Sub Print_RP1()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If GetDBData_RP(1) = True Then
      InsertQueryLog ("") 'Add By Sindy 2010/12/2
      BuildField_RP (1)
      Generate_RP1
      Clear
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/2
      strMsg = "無資料"
      strTit = "錯誤"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
End Sub

Private Sub Print_RP2()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If GetDBData_RP(2) = True Then
      InsertQueryLog ("") 'Add By Sindy 2010/12/2
      BuildField_RP (2)
      Generate_RP2
      Clear
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/2
      strMsg = "無資料"
      strTit = "錯誤"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
End Sub

Public Function Length(ByVal strData As String) As Integer
   Length = LenB(StrConv(strData, vbFromUnicode))
End Function

Public Function LeftStr(ByVal strData As String, ByVal nLen As Integer) As String
   LeftStr = StrConv(MidB(StrConv(strData, vbFromUnicode), 1, nLen), vbUnicode)
End Function

Public Function CheckDataValid() As Boolean
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   CheckDataValid = True
   
   If IsEmpty(text01_02) = True Then
      CheckDataValid = False
      strMsg = "公告日必須輸入"
      strTit = "檢核輸入"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   ElseIf CheckIsTaiwanDate(text01_02, False) = False Then
      CheckDataValid = False
      strMsg = "請輸入正確的公告日"
      strTit = "檢核輸入"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   If IsEmpty(text01_01) = False And IsEmpty(text01_02) = False Then
      If Val(text01_01) > Val(text01_02) Then
         CheckDataValid = False
         strMsg = "公告日起日必須小於止日"
         strTit = "檢核輸入"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   End If
   
   Select Case text02
      Case "1", "2":
      Case Else
         CheckDataValid = False
         strMsg = "請輸入出表選擇"
         strTit = "檢核輸入"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
   End Select
   
EXITSUB:
End Function

' 將所有的文字反白
Private Sub InverseAll(ByRef tb As TextBox)
   tb.SelStart = 0
   tb.SelLength = Len(tb.Text)
End Sub

Private Sub text01_01_GotFocus()
   InverseAll text01_01
End Sub

Private Sub text01_02_GotFocus()
   InverseAll text01_02
End Sub

Private Sub text02_GotFocus()
   InverseAll text02
End Sub


