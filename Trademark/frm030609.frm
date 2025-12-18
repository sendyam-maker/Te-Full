VERSION 5.00
Begin VB.Form frm030609 
   BorderStyle     =   1  '單線固定
   Caption         =   "表四.各類別市場佔有統計表"
   ClientHeight    =   2040
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   5520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5520
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      ItemData        =   "frm030609.frx":0000
      Left            =   1230
      List            =   "frm030609.frx":0002
      TabIndex        =   3
      Top             =   2100
      Visible         =   0   'False
      Width           =   3972
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   4440
      TabIndex        =   5
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3480
      TabIndex        =   4
      Top             =   60
      Width           =   912
   End
   Begin VB.TextBox textTMBM07_2 
      Height          =   264
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   1
      Top             =   720
      Width           =   1092
   End
   Begin VB.TextBox textTMBM07_1 
      Height          =   264
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   0
      Top             =   720
      Width           =   1092
   End
   Begin VB.TextBox textNA02 
      Height          =   264
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1080
      Width           =   1092
   End
   Begin VB.Label Label10 
      Caption         =   "印表機 :"
      Height          =   255
      Left            =   150
      TabIndex        =   9
      Top             =   2100
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   2520
      X2              =   2760
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "公報卷期："
      Height          =   252
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "列印區域："
      Height          =   252
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "(A:國內 B:大陸 C:國外 空白:全部)"
      Height          =   252
      Left            =   2520
      TabIndex        =   6
      Top             =   1080
      Width           =   2772
   End
End
Attribute VB_Name = "frm030609"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

Const m_CharWidth = 120
Const m_CharHeight = 240
'edit by nick 2004/12/14
'Const m_PaperSize = "REPORT"
Const m_PaperSize = "A4"

' 宣告報表表頭的欄位其資料型態
Private Type REPORTFIELD
   Name As String
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

' 宣告地區項目的資料型態
Private Type COUNTRYITEM
   ' 地區代碼
   CountryCode As String
   CountryName As String
   Count As Integer
End Type

' 商品類別資料結構
Private Type PRODUCTITEM
   Product As String
   Count As Integer
   CountryList() As COUNTRYITEM
   CountryCount As Integer
End Type
' 定義地區串列
Dim m_ProductList() As PRODUCTITEM
Dim m_ProductCount As Integer

' 定義地區串列
Dim m_ZoneList() As COUNTRYITEM
Dim m_ZoneCount As Integer

Dim m_DefaultPrinter As String

Private Sub Form_Load()
'edit by nick 2004/12/14
'   Dim Prn As Printer
'
'   m_DefaultPrinter = Printer.DeviceName
   MoveFormToCenter Me
   
'edit by nick 2004/12/14
'   For Each Prn In Printers
'      If Prn.DeviceName <> m_DefaultPrinter Then
'         cmbPrinter.AddItem Prn.DeviceName
'      End If
'   Next
'   cmbPrinter.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
'edit by nick 2004/12/14
'   Dim Prn As Printer
'   For Each Prn In Printers
'      If Prn.DeviceName = m_DefaultPrinter Then
'         Set Printer = Prn
'         Exit For
'      End If
'   Next
   'Add By Cheng 2002/07/19
   Set frm030609 = Nothing
End Sub

Private Sub Clear()
   Dim nX As Integer
   
   If m_ProductCount > 0 Then
      For nX = 0 To m_ProductCount - 1
         If m_ProductList(nX).CountryCount > 0 Then
            Erase m_ProductList(nX).CountryList
         End If
         m_ProductList(nX).CountryCount = 0
      Next nX
      Erase m_ProductList
   End If
   m_ProductCount = 0
   
   If m_ZoneCount > 0 Then
      Erase m_ZoneList
   End If
   m_ZoneCount = 0

End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdok_Click()
   Dim Prn As Printer
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If CheckDataValid() = True Then
      '搜尋 Printer
'edit by nick 2004/12/14
'      For Each Prn In Printers
'         If Prn.DeviceName = cmbPrinter.Text Then
'            Set Printer = Prn
'            Exit For
'         End If
'      Next
'      Printer.PaperSize = 39
      DoEvents
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 建立欄位資訊
      BuildField_RP
      
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/22 清除查詢印表記錄檔欄位
      If Len(textTMBM07_1) <> 0 Or Len(textTMBM07_2) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & Label1 & textTMBM07_1 & "-" & textTMBM07_2 'Add By Sindy 2010/10/22
      End If
      If Len(textNA02) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & Label2 & textNA02 & Label3  'Add By Sindy 2010/10/22
      End If
      
      ' 取得資料庫中的資料
      If GetDBData_RP = False Then
         GoTo EXITSUB
      End If
      InsertQueryLog ("") 'Add By Sindy 2010/10/22
      
      ' 列印
      Generate_RP
      'Generate_RP_SCREEN
      ' 清除
      Clear
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      
      strTit = "輸出報表"
      strMsg = "列印結束"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
    End If
EXITSUB:
    Screen.MousePointer = vbDefault
End Sub

' 取得該商品種類的總數量
Private Function GetProductAmount(ByRef ProductInfo As PRODUCTITEM) As Integer
   Dim nAmount As Variant
   Dim nX As Integer
   nAmount = 0
   For nX = 0 To ProductInfo.CountryCount - 1
      nAmount = nAmount + ProductInfo.CountryList(nX).Count
   Next nX
   GetProductAmount = nAmount
End Function

' 新增一個商品件數
Private Function InsertProductItem(ByVal strProduct As String, ByVal strZoneCode As String, ByVal strZoneName As String)
   Dim bFindProduct As Boolean
   Dim bFindCountry As Boolean
   Dim nSortX, nSortY As Integer
   Dim CountryTemp As COUNTRYITEM
   Dim ProductTemp As PRODUCTITEM
   Dim bFromSec As Boolean
   Dim bToSec As Boolean
   Dim nX, ny, nZ As Integer
   
   ' 找尋商品種類串列
   bFindProduct = False
   ' 搜尋商品種類串列
   For nX = 0 To m_ProductCount - 1
      If strProduct = m_ProductList(nX).Product Then
         bFindProduct = True
         ' 商品種類結構中的總數
         m_ProductList(nX).Count = m_ProductList(nX).Count + 1
         bFindCountry = False
         ' 搜尋商品種類結構中的地區串列
         For ny = 0 To m_ProductList(nX).CountryCount - 1
            If strZoneCode = m_ProductList(nX).CountryList(ny).CountryCode Then
               bFindCountry = True
               m_ProductList(nX).CountryList(ny).Count = m_ProductList(nX).CountryList(ny).Count + 1
               Exit For
            End If
         Next ny
         ' 找不到地區則產生一個新的地區結構
         If bFindCountry = False Then
            ny = m_ProductList(nX).CountryCount
            ReDim Preserve m_ProductList(nX).CountryList(ny + 1)
            m_ProductList(nX).CountryList(ny).CountryCode = strZoneCode
            m_ProductList(nX).CountryList(ny).CountryName = strZoneName
            m_ProductList(nX).CountryList(ny).Count = 1
            m_ProductList(nX).CountryCount = m_ProductList(nX).CountryCount + 1
         End If
         Exit For
      End If
   Next nX
   
   ' 找不到除存該商品種類的結構則新增一個
   If bFindProduct = False Then
      nX = m_ProductCount
      ReDim Preserve m_ProductList(nX + 1)
      m_ProductList(nX).Product = strProduct
      m_ProductList(nX).CountryCount = 0
      m_ProductList(nX).Count = 1
      ny = m_ProductList(nX).CountryCount
      m_ProductCount = m_ProductCount + 1
      ReDim Preserve m_ProductList(nX).CountryList(ny + 1)
      m_ProductList(nX).CountryList(ny).CountryCode = strZoneCode
      m_ProductList(nX).CountryList(ny).CountryName = strZoneName
      m_ProductList(nX).CountryList(ny).Count = 1
      m_ProductList(nX).CountryCount = m_ProductList(nX).CountryCount + 1
   End If
End Function

' 新增一個商品件數
Private Function InsertZoneItem(ByVal strZoneCode As String, ByVal strZoneName As String)
   Dim bFindCountry As Boolean
   Dim nSortX, nSortY As Integer
   Dim CountryTemp As COUNTRYITEM
   Dim nX, ny, nZ As Integer
   
   bFindCountry = False
   For nX = 0 To m_ZoneCount - 1
      If strZoneCode = m_ZoneList(nX).CountryCode Then
         bFindCountry = True
         m_ZoneList(nX).Count = m_ZoneList(nX).Count + 1
         Exit For
      End If
   Next nX
   If bFindCountry = False Then
      nX = m_ZoneCount
      ReDim Preserve m_ZoneList(nX + 1)
      m_ZoneList(nX).CountryCode = strZoneCode
      m_ZoneList(nX).CountryName = strZoneName
      m_ZoneList(nX).Count = 1
      m_ZoneCount = m_ZoneCount + 1
   End If
End Function

' 從資料庫中取得所有的資料
Private Function GetDBData_RP() As Boolean
   Dim rsMain As New ADODB.Recordset
   Dim strSql As String
   Dim strSubSQL As String
   Dim strTemp As String
   Dim strZoneKind, strZoneName, strZoneCode, strProductName, strProductCode As String
   Dim bFindProduct As Boolean
   Dim bFindCountry As Boolean
   Dim nSortX, nSortY As Integer
   Dim CountryTemp As COUNTRYITEM
   Dim ProductTemp As PRODUCTITEM
   Dim bFromSec As Boolean
   Dim bToSec As Boolean
   Dim nCount As Integer
   Dim nStart As Integer
   Dim nX, ny, nZ As Integer
   
   GetDBData_RP = True
   
   strSubSQL = Empty
   Select Case textNA02
      Case "a", "A":
         strSubSQL = "NA02 LIKE '" & "A%" & "'"
      Case "b", "B":
         strSubSQL = "NA02 LIKE '" & "B%" & "'"
      Case "c", "C":
         strSubSQL = "NA02 LIKE '" & "C%" & "'"
      Case Else:
         strSubSQL = Empty
   End Select
   
   ' 產生SQL查詢語法
   bFromSec = Not IsEmptyText(textTMBM07_1.Text)
   bToSec = Not IsEmptyText(textTMBM07_2.Text)
   'Modify By Sindy 2013/8/19 + length(na01)=3 AND
   If bFromSec = True And bToSec = True Then
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,TA01,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01 (+) AND " & _
                     "TMBM07 >= '" & textTMBM07_1 & "' AND " & _
                     "TMBM07 <= '" & textTMBM07_2 & "'"
      If strSubSQL <> Empty Then
         strSql = strSql & " " & "AND " & strSubSQL
      End If
   ElseIf bFromSec = True And bToSec = False Then
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,TA01,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01 (+) AND " & _
                     "TMBM07 >= '" & textTMBM07_1 & "'"
      If strSubSQL <> Empty Then
         strSql = strSql & " " & "AND " & strSubSQL
      End If
   ElseIf bFromSec = False And bToSec = True Then
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,TA01,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01 (+) AND " & _
                     "TMBM07 <= '" & textTMBM07_2 & "'"
      If strSubSQL <> Empty Then
         strSql = strSql & " " & "AND " & strSubSQL
      End If
   Else
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,TA01,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01 (+) "
      If strSubSQL <> Empty Then
         strSql = strSql & " " & "AND " & strSubSQL
      End If
   End If
   
   ' 取得資料庫的資料
   rsMain.CursorLocation = adUseClient
   rsMain.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 無資料則離開
   If rsMain.RecordCount <= 0 Then
      GetDBData_RP = False
      GoTo EXITSUB
   End If
   
   ' 商品類別的種類
   m_ProductCount = 0
   
   rsMain.MoveFirst
   ' 依序從資料記錄中取出欄位的內容
   Do While Not rsMain.EOF
      ' 地區名稱
      strZoneName = Empty
      If IsNull(rsMain.Fields("TMBM05")) = False Then
         strZoneName = rsMain.Fields("TMBM05")
      End If
      ' 地區別
      strZoneKind = Empty
      If IsNull(rsMain.Fields("NA02")) = False Then
         strZoneKind = rsMain.Fields("NA02")
      End If
      ' 地區代碼
      strZoneCode = Empty
      If IsNull(rsMain.Fields("NA01")) = False Then
         strZoneCode = rsMain.Fields("NA01")
      End If
      
      strTemp = Empty
      nStart = 0
      
      If IsNull(rsMain.Fields("TMBM08")) = False Then
         nCount = GetSubStringCount(rsMain.Fields("TMBM08"))
         For nX = 1 To nCount
            strTemp = GetSubString(rsMain.Fields("TMBM08"), nX)
            If strTemp <> Empty Then
               InsertProductItem strTemp, strZoneCode, strZoneName
               InsertZoneItem strZoneCode, strZoneName
            End If
         Next nX
      End If

      rsMain.MoveNext
   Loop
   
   ' 對商品類別串列依商品種類代碼小到大排序
   For nSortX = 0 To m_ProductCount - 1
      For nSortY = nSortX To m_ProductCount - 1
         If m_ProductList(nSortX).Product > m_ProductList(nSortY).Product Then
            ProductTemp = m_ProductList(nSortX)
            m_ProductList(nSortX) = m_ProductList(nSortY)
            m_ProductList(nSortY) = ProductTemp
         End If
      Next nSortY
   Next nSortX
   ' 對商品種類結構中的地區串列依該地區的數量多寡由大到小排序
   For nX = 0 To m_ProductCount - 1
      For nSortX = 0 To m_ProductList(nX).CountryCount - 1
         For nSortY = nSortX To m_ProductList(nX).CountryCount - 1
            If m_ProductList(nX).CountryList(nSortX).Count < m_ProductList(nX).CountryList(nSortY).Count Then
               CountryTemp = m_ProductList(nX).CountryList(nSortX)
               m_ProductList(nX).CountryList(nSortX) = m_ProductList(nX).CountryList(nSortY)
               m_ProductList(nX).CountryList(nSortY) = CountryTemp
            ElseIf m_ProductList(nX).CountryList(nSortX).Count = m_ProductList(nX).CountryList(nSortY).Count Then
               If m_ProductList(nX).CountryList(nSortX).CountryCode > m_ProductList(nX).CountryList(nSortY).CountryCode Then
                  CountryTemp = m_ProductList(nX).CountryList(nSortX)
                  m_ProductList(nX).CountryList(nSortX) = m_ProductList(nX).CountryList(nSortY)
                  m_ProductList(nX).CountryList(nSortY) = CountryTemp
               End If
            End If
         Next nSortY
      Next nSortX
   Next nX
   ' 對地區串列由大到小排名
   For nX = 0 To m_ZoneCount - 1
      For ny = nX To m_ZoneCount - 1
         If m_ZoneList(nX).Count < m_ZoneList(ny).Count Then
            CountryTemp = m_ZoneList(nX)
            m_ZoneList(nX) = m_ZoneList(ny)
            m_ZoneList(ny) = CountryTemp
         ElseIf m_ZoneList(nX).Count = m_ZoneList(ny).Count Then
            If m_ZoneList(nX).CountryCode > m_ZoneList(ny).CountryCode Then
               CountryTemp = m_ZoneList(nX)
               m_ZoneList(nX) = m_ZoneList(ny)
               m_ZoneList(ny) = CountryTemp
            End If
         End If
      Next ny
   Next nX
   
   ' 測試
   'For nX = 0 To m_ProductCount - 1
   '   Debug.Print "商品類別為 " & m_ProductList(nX).Product
   '   For nY = 0 To m_ProductList(nX).CountryCount - 1
   '      Debug.Print "   地區為 " & m_ProductList(nX).CountryList(nY).CountryName & "的數量為 " & m_ProductList(nX).CountryList(nY).Count
   '   Next nY
   'Next nX
   
EXITSUB:
   rsMain.Close
   Set rsMain = Nothing
End Function

' 設定報表欄位的左方位置及其名稱
Public Sub BuildField_RP()
   Dim nIndex As Integer
   Dim nFieldWidth As Integer
   
   Select Case m_PaperSize
      Case "A4"
         m_LeftMargin = 1
         m_TopMargin = 3
         m_ReportWidth = 154
         m_ReportDataRows = 27
         nFieldWidth = 7
      Case "REPORT"
         m_LeftMargin = 1
         m_TopMargin = 3
         m_ReportWidth = 154
         m_ReportDataRows = 45
         nFieldWidth = 9
      Case Else
         m_LeftMargin = 10
         m_TopMargin = 3
         m_ReportWidth = 120
         m_ReportDataRows = 27
         nFieldWidth = 7
   End Select
   
   For nIndex = 0 To 16
      m_Field(nIndex).Width = nFieldWidth - 1
      m_Field(nIndex).Left = m_LeftMargin + (nIndex * nFieldWidth) + 5
      Select Case nIndex
         Case 0:
            m_Field(nIndex).Width = 8
            m_Field(nIndex).Left = m_LeftMargin + (nIndex * nFieldWidth)
            m_Field(nIndex).Name = "商品種類"
         Case 1:
            m_Field(nIndex).Left = m_LeftMargin + (nIndex * nFieldWidth) + 2
            m_Field(nIndex).Name = "排名"
         Case 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15:
            m_Field(nIndex).Name = CStr(nIndex - 1)
         Case 16:
            m_Field(nIndex).Name = "本期"
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

' 列印分隔線
Public Sub PrintTerminateLine(ByVal nRow As Integer)
   Dim nCount As Integer
   For nCount = 0 To m_ReportWidth - 1
      Printer.CurrentX = (m_LeftMargin + nCount) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "="
   Next nCount
End Sub

' 列印表頭
Public Sub PrintPageHeader_RP(ByVal nPage As Integer)
   Dim nCount As Integer
   Dim strData1 As String
   Dim strData2 As String
   Dim nIndex As Integer
   Dim nRow As Integer
   Dim nX As Long
   Dim ny As Long
   Dim nCenter As Long
   Dim strTemp As String
   
   strData1 = textTMBM07_1
   strData2 = textTMBM07_2
      
   ' 表頭
   nRow = 1
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.FontSize = 24
   Printer.Font.Underline = True
   nX = m_LeftMargin + m_ReportWidth / 2 - 26
   Printer.CurrentX = nX * m_CharWidth
   Printer.Print "表四：各類別市場佔有統計表"
   
   Printer.Font.Underline = False
   
   nRow = nRow + 2
   Printer.FontSize = 12
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "列印人:" & strUserName
   
   nX = m_LeftMargin + m_ReportWidth / 2 - 16
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.FontSize = 12
   Printer.Print "公報卷期:" & strData1 & " - " & strData2
   
   nX = m_LeftMargin + m_ReportWidth - 38
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "製表日期:" & Format(ChangeWStringToWDateString(GetTodayDate), "EE/MM/DD")
   
   nRow = nRow + 1
   nX = m_LeftMargin + m_ReportWidth - 38
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "頁"
   
   nX = m_LeftMargin + m_ReportWidth - 32
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "次:" & nPage
   
   ' 列印分隔線
   nRow = nRow + 1
   PrintSplitLine nRow
   
   nRow = nRow + 1
   For nIndex = 0 To 16
      nCenter = ((m_Field(nIndex).Left * m_CharWidth) + (m_Field(nIndex).Left + m_Field(nIndex).Width) * m_CharWidth) / 2
      strTemp = LeftStr(m_Field(nIndex).Name, m_Field(nIndex).Width)
      Printer.CurrentX = nCenter - Printer.TextWidth(strTemp) / 2
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print strTemp
   Next nIndex
     
   ' 列印分隔線
   nRow = nRow + 1
   PrintSplitLine nRow
   
   m_HeaderHeight = nRow
End Sub

Public Function Min(ByVal nValue1 As Integer, ByVal nValue2 As Integer) As Integer
   If nValue2 < nValue1 Then
      Min = nValue2
   Else
      Min = nValue1
   End If
End Function

Public Function Length(ByVal strData As String) As Integer
   Length = LenB(StrConv(strData, vbFromUnicode))
End Function

Public Function LeftStr(ByVal strData As String, ByVal nLen As Integer) As String
   LeftStr = StrConv(MidB(StrConv(strData, vbFromUnicode), 1, nLen), vbUnicode)
End Function

' 列印表一的內容
Public Sub Generate_RP()
   Dim nRow As Integer
   Dim nPage As Integer
   Dim fld(17) As String
   Dim nType As Integer
   Dim nAmount As Variant
   Dim nTotalAmount As Variant
   Dim nZoneCount As Variant
   Dim fValue As Variant
   Dim nX As Integer
   Dim ny As Integer
   Dim nZ As Integer
   Dim nCenter As Long
   Dim nRight As Long
      
   ' 紙張大小
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
   ' 紙張方向
   
   ' 印表頭
   nPage = 1
   PrintPageHeader_RP nPage
   
   nRow = 0
   For nX = 0 To m_ProductCount - 1
      ' 若列數超過頁面的高度限制時則換頁
      If nRow > m_ReportDataRows Then
         Printer.NewPage
         nPage = nPage + 1
         PrintPageHeader_RP nPage
         nRow = 0
      End If
      
      ' 清除欄位
      For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
      ' 地區
      nRow = nRow + 1
      fld(0) = m_ProductList(nX).Product
      fld(1) = "地區"
      For ny = 0 To Min(13, m_ProductList(nX).CountryCount - 1)
         fld(ny + 2) = m_ProductList(nX).CountryList(ny).CountryName
      Next ny
      For nZ = 0 To 16
         Select Case nZ
            Case 0, 1
               Printer.FontSize = 12
               Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
               Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
               Printer.Print fld(nZ)
            Case Else
               Printer.FontSize = 8
               nCenter = ((m_Field(nZ).Left * m_CharWidth) + (m_Field(nZ).Left + m_Field(nZ).Width) * m_CharWidth) / 2
               Printer.CurrentX = nCenter - Printer.TextWidth(fld(nZ)) / 2
               Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
               Printer.Print fld(nZ)
               Printer.FontSize = 12
         End Select
      Next nZ
      ' 清除欄位
      For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
      ' 數量
      nRow = nRow + 1
      fld(0) = m_ProductList(nX).Product
      fld(1) = "數量"
      For ny = 0 To Min(13, m_ProductList(nX).CountryCount - 1)
         fld(ny + 2) = m_ProductList(nX).CountryList(ny).Count
      Next ny
      fld(16) = m_ProductList(nX).Count
       ' 輸出數量列
      For nZ = 0 To 16
         Select Case nZ
         Case 0, 1
            Printer.FontSize = 12
            Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
            Printer.Print fld(nZ)
         Case Else
            Printer.FontSize = 8
            nRight = (m_Field(nZ).Left + m_Field(nZ).Width - 2) * m_CharWidth
            Printer.CurrentX = nRight - Printer.TextWidth(fld(nZ))
            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
            Printer.Print fld(nZ)
         End Select
      Next nZ
      ' 清除欄位
      For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
      ' 百分比
      nRow = nRow + 1
      fld(0) = m_ProductList(nX).Product
      fld(1) = "百分比"
      For ny = 0 To Min(13, m_ProductList(nX).CountryCount - 1)
         nAmount = m_ProductList(nX).CountryList(ny).Count
         nTotalAmount = m_ProductList(nX).Count
         If nTotalAmount > 0 Then
            fValue = nAmount / nTotalAmount * 100
         Else
            fValue = 0#
         End If
         fld(ny + 2) = Format(fValue, "##0.00") & " %"
      Next ny
      'fld(16) = "100.00 %"
      fld(16) = Empty
      ' 輸出百分比列
      For nZ = 0 To 16
         Select Case nZ
            Case 0, 1
               Printer.FontSize = 12
               Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
               Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
               Printer.Print fld(nZ)
            Case Else
               Printer.FontSize = 8
               nRight = (m_Field(nZ).Left + m_Field(nZ).Width - 2) * m_CharWidth
               Printer.CurrentX = nRight - Printer.TextWidth(fld(nZ))
               Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
               Printer.Print fld(nZ)
         End Select
      Next nZ
      ' 列印分隔線
      nRow = nRow + 1
      PrintSplitLine m_HeaderHeight + nRow
   Next nX
   
   ' 計算總數
   nTotalAmount = 0
   For nX = 0 To m_ZoneCount - 1
      nTotalAmount = nTotalAmount + m_ZoneList(nX).Count
   Next nX
   
   ' 列印合計
   nRow = nRow + 1
   
   ' 清除欄位
   For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
   
   fld(0) = "合計"
   fld(1) = "地區"
   For nX = 0 To Min(13, m_ZoneCount - 1)
      fld(nX + 2) = m_ZoneList(nX).CountryName
   Next nX
   For nZ = 0 To 16
      Select Case nZ
         Case 0, 1
            Printer.FontSize = 12
            Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
            Printer.Print fld(nZ)
         Case Else
            Printer.FontSize = 8
            nCenter = ((m_Field(nZ).Left * m_CharWidth) + (m_Field(nZ).Left + m_Field(nZ).Width) * m_CharWidth) / 2
            Printer.CurrentX = nCenter - Printer.TextWidth(fld(nZ)) / 2
            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
            Printer.Print fld(nZ)
            Printer.FontSize = 12
      End Select
   Next nZ
   
   nRow = nRow + 1
   ' 清除欄位
   For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
   fld(1) = "數量"
   For nX = 0 To Min(13, m_ZoneCount - 1)
      fld(nX + 2) = m_ZoneList(nX).Count
   Next nX
   fld(16) = nTotalAmount
   For nZ = 0 To 16
      Select Case nZ
      Case 0, 1
         Printer.FontSize = 12
         Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
         Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
         Printer.Print fld(nZ)
      Case Else
         Printer.FontSize = 8
         nRight = (m_Field(nZ).Left + m_Field(nZ).Width - 2) * m_CharWidth
         Printer.CurrentX = nRight - Printer.TextWidth(fld(nZ))
         Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
         Printer.Print fld(nZ)
      End Select
   Next nZ
   
   nRow = nRow + 1
   
   ' 清除欄位
   For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
   
   fld(0) = Empty
   fld(1) = "百分比"
   For nX = 0 To Min(13, m_ZoneCount - 1)
      nAmount = m_ZoneList(nX).Count
      If nTotalAmount > 0 Then
         fValue = nAmount / nTotalAmount * 100
      Else
         fValue = 0#
      End If
      fld(nX + 2) = Format(fValue, "##0.00") & " %"
   Next nX
   fld(16) = Empty
   ' 輸出百分比列
   For nZ = 0 To 16
      Select Case nZ
         Case 0, 1
            Printer.FontSize = 12
            Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
            Printer.Print fld(nZ)
         Case Else
            Printer.FontSize = 8
            nRight = (m_Field(nZ).Left + m_Field(nZ).Width - 2) * m_CharWidth
            Printer.CurrentX = nRight - Printer.TextWidth(fld(nZ))
            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
            Printer.Print fld(nZ)
      End Select
   Next nZ
   ' 列印分隔線
   nRow = nRow + 1
   PrintTerminateLine m_HeaderHeight + nRow
      
   Printer.EndDoc
End Sub

' 背景列印
Public Sub PrintReportBK(ByVal strPrinter As String, ByVal TMBM07_1 As String, ByVal TMBM07_2 As String)
   Dim Prn As Printer
   
   Me.Hide
   textTMBM07_1 = TMBM07_1
   textTMBM07_2 = TMBM07_2
   textNA02 = " "
   
   '搜尋 Printer
'edit by nick 2004/12/15
'   For Each Prn In Printers
'      If Prn.DeviceName = strPrinter Then
'         Set Printer = Prn
'         Exit For
'      End If
'   Next
   
   ' 建立欄位資訊
   BuildField_RP
   ' 取得資料庫中的資料
   If GetDBData_RP = False Then: GoTo EXITSUB
   ' 列印
   Generate_RP
   ' 清除所佔用的空間
   Clear
EXITSUB:
   Set frm030609 = Nothing
End Sub

Private Sub textNA02_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 列印區域
Private Sub textNA02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textNA02) = False Then
      Select Case textNA02
         Case "A", "B", "C":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "列印區域只可輸入空白或A,B,C"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textNA02_GotFocus
      End Select
   End If
End Sub

' 公報卷期(起)
Private Sub textTMBM07_1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTMBM07_1) = False Then
      If IsNumeric(textTMBM07_1) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "公報卷期(起)只可輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTMBM07_1_GotFocus
      End If
   End If
End Sub

' 公報卷期(迄)
Private Sub textTMBM07_2_LostFocus()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   If IsEmptyText(textTMBM07_2) = False Then
      If IsNumeric(textTMBM07_2) = False Then
         strTit = "資料檢核"
         strMsg = "公報卷期(迄)只可輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTMBM07_2_GotFocus
      Else
         If Not ChkRange(textTMBM07_1, textTMBM07_2, "公報卷期") Then
         
         End If
      End If
   End If
End Sub

Private Sub textNA02_GotFocus()
   InverseTextBox textNA02
End Sub

Private Sub textTMBM07_1_GotFocus()
   InverseTextBox textTMBM07_1
End Sub

Private Sub textTMBM07_2_GotFocus()
   InverseTextBox textTMBM07_2
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   If IsEmptyText(textTMBM07_1) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入公報卷期(起)"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   If IsEmptyText(textTMBM07_1) = False And IsEmptyText(textTMBM07_2) = False Then
      If Val(textTMBM07_1) > Val(textTMBM07_2) Then
         strTit = "資料檢核"
         strMsg = "公報卷期範圍不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   End If
   
   'Added by Lydia 2017/02/06 檢查公報檔的地區
   If Pub_ChkTMBMValidate(textTMBM07_1, textTMBM07_2) = False Then
      GoTo EXITSUB
   End If
   'end 2017/02/06
   
   CheckDataValid = True
EXITSUB:
End Function

' 列印表一的內容
Public Sub Generate_RP_SCREEN()
   Dim nRow As Integer
   Dim nPage As Integer
   Dim fld(17) As String
   Dim nType As Integer
   Dim nAmount As Variant
   Dim nTotalAmount As Variant
   Dim nZoneCount As Variant
   Dim fValue As Variant
   Dim nX As Integer
   Dim ny As Integer
   Dim nZ As Integer
   Dim nCenter As Long
   Dim nRight As Long
   
   nRow = 0
   For nX = 0 To m_ProductCount - 1
      ' 清除欄位
      For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
      ' 地區
      nRow = nRow + 1
      fld(0) = m_ProductList(nX).Product
      fld(1) = "地區"
      For ny = 0 To Min(13, m_ProductList(nX).CountryCount - 1)
         fld(ny + 2) = m_ProductList(nX).CountryList(ny).CountryName
      Next ny
      
      ' 清除欄位
      For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
      ' 數量
      nRow = nRow + 1
      fld(0) = m_ProductList(nX).Product
      fld(1) = "數量"
      For ny = 0 To Min(13, m_ProductList(nX).CountryCount - 1)
         fld(ny + 2) = m_ProductList(nX).CountryList(ny).Count
      Next ny
      fld(16) = m_ProductList(nX).Count
       
      ' 清除欄位
      For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
      ' 百分比
      nRow = nRow + 1
      fld(0) = m_ProductList(nX).Product
      fld(1) = "百分比"
      For ny = 0 To Min(13, m_ProductList(nX).CountryCount - 1)
         nAmount = m_ProductList(nX).CountryList(ny).Count
         nTotalAmount = m_ProductList(nX).Count
         If nTotalAmount > 0 Then
            fValue = nAmount / nTotalAmount * 100
         Else
            fValue = 0#
         End If
         fld(ny + 2) = Format(fValue, "##0.00") & " %"
      Next ny
      fld(16) = "100.00 %"
      
      ' 列印分隔線
      nRow = nRow + 1
      
   Next nX
   
End Sub

