VERSION 5.00
Begin VB.Form frm030612 
   BorderStyle     =   1  '單線固定
   Caption         =   "國外前十大申請國及其商品類別排名"
   ClientHeight    =   1785
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4185
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4185
   Begin VB.ComboBox cmbPrinter 
      Height          =   276
      ItemData        =   "frm030612.frx":0000
      Left            =   1320
      List            =   "frm030612.frx":0002
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   2652
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   3180
      TabIndex        =   4
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2220
      TabIndex        =   3
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
   Begin VB.Label Label10 
      Caption         =   "印表機："
      Height          =   252
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   972
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
      TabIndex        =   5
      Top             =   720
      Width           =   972
   End
End
Attribute VB_Name = "frm030612"
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
'edit by nickc 2006/06/23
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

' 商品類別資料結構
Private Type PRODUCTITEM
   Product As String
   Count As Integer
End Type

Dim m_ProductTmpList() As PRODUCTITEM
Dim m_ProductTmpCount As Integer

' 宣告地區項目的資料型態
Private Type ZONEITEM
   ' 地區代碼
   ZoneCode As String
   ZoneName As String
   Count As Integer
   ProductList() As PRODUCTITEM
   ProductCount As Integer
End Type

' 定義地區串列
Dim m_ZoneList() As ZONEITEM
Dim m_ZoneCount As Integer

Dim m_DefaultPrinter As String

Private Sub Form_Load()
   Dim Prn As Printer
   
'edit by nickc 2006/06/23 改 A4
'   m_DefaultPrinter = Printer.DeviceName
   MoveFormToCenter Me
   
'edit by nickc 2006/06/23 改 A4
'   For Each Prn In Printers
'      If Prn.DeviceName <> m_DefaultPrinter Then
'         cmbPrinter.AddItem Prn.DeviceName
'      End If
'   Next
'   cmbPrinter.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
'edit by nickc 2006/06/23 改 A4
'   Dim Prn As Printer
'   For Each Prn In Printers
'      If Prn.DeviceName = m_DefaultPrinter Then
'         Set Printer = Prn
'         Exit For
'      End If
'   Next
   'Add By Cheng 2002/07/19
   Set frm030612 = Nothing
End Sub

' 清除系統所佔用的記憶體
Private Sub Clear()
   Dim nX As Integer
   Dim ny As Integer
   
   If m_ZoneCount > 0 Then
      For nX = 0 To m_ZoneCount - 1
         If m_ZoneList(nX).ProductCount > 0 Then
            Erase m_ZoneList(nX).ProductList
         End If
         m_ZoneList(nX).ProductCount = 0
      Next nX
      Erase m_ZoneList
   End If
   m_ZoneCount = 0
   
End Sub

' 離開
Private Sub cmdExit_Click()
   Unload Me
End Sub

' 確定
Private Sub cmdOK_Click()
   Dim Prn As Printer
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If CheckDataValid() = True Then
      '搜尋 Printer
'edit by nickc 2006/06/23 改 A4
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
      
      ' 取得資料庫中的資料
      If GetDBData_RP = False Then
         GoTo EXITSUB
      End If
      InsertQueryLog ("") 'Add By Sindy 2010/10/22
      
      ' 列印
      'edit by nickc 2006/06/23 改 A4
      Generate_RP
      'Generate_RP_950623
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

' 設定報表欄位的左方位置及其名稱
Public Sub BuildField_RP()
   Dim nIndex As Integer
   Dim nFieldWidth As Integer
   
   Select Case m_PaperSize
      'add by nickc 2006/06/23
      Case "A4"
         m_LeftMargin = 1
         m_TopMargin = 3
         m_ReportWidth = 154
         m_ReportDataRows = 27
         nFieldWidth = 7.5
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
      m_Field(nIndex).Left = m_LeftMargin + (nIndex * nFieldWidth) - 0.5
      Select Case nIndex
         Case 0:
            m_Field(nIndex).Left = m_LeftMargin + (nIndex * nFieldWidth)
            m_Field(nIndex).Name = "地區"
         Case 1:
            m_Field(nIndex).Left = m_LeftMargin + (nIndex * nFieldWidth)
            m_Field(nIndex).Name = "總數"
         Case 2:
            m_Field(nIndex).Width = 9
            m_Field(nIndex).Left = m_LeftMargin + (nIndex * nFieldWidth)
            m_Field(nIndex).Name = "排名"
         Case 3:
            m_Field(nIndex).Left = m_LeftMargin + (nIndex * nFieldWidth)
            m_Field(nIndex).Name = CStr(nIndex - 2)
         Case 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16:
            'm_Field(nIndex).Name = CStr(nIndex - 1)
            m_Field(nIndex).Name = CStr(nIndex - 2)
         'Case 16:
         '   m_Field(nIndex).Name = "地區總數"
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
   nX = m_LeftMargin + m_ReportWidth / 2 - 39 '34
   Printer.CurrentX = nX * m_CharWidth
   Printer.Print "國外前十大申請國及其商品類別排名表"
   Printer.Font.Underline = False
   
   nRow = nRow + 2
   Printer.FontSize = 12
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "列印人 : " & strUserName
   
   nX = m_LeftMargin + m_ReportWidth / 2 - 16 '10
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.FontSize = 12
   Printer.Print "公報卷期 : " & strData1 & " - " & strData2
   
   nX = m_LeftMargin + m_ReportWidth - 38 ' 20
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "製表日期 : " & Format(ChangeWStringToWDateString(GetTodayDate), "EE/MM/DD")
   
   nRow = nRow + 1
   nX = m_LeftMargin + m_ReportWidth - 38 '20
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "頁"
   
   nX = m_LeftMargin + m_ReportWidth - 32 '14
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "次 : " & nPage
   
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

' 新增一個商品件數
Private Function InsertProductItem(ByVal strProduct As String, ByVal strZoneCode As String, ByVal strZoneName As String)
   Dim bFindProduct As Boolean
   Dim bFindZone As Boolean
   Dim nSortX, nSortY As Integer
   Dim bFromSec As Boolean
   Dim bToSec As Boolean
   Dim nX, ny, nZ As Integer
      
   ' 搜尋地區串列
   bFindZone = False
   For nX = 0 To m_ZoneCount - 1
      If strZoneCode = m_ZoneList(nX).ZoneCode Then
         bFindZone = True
         ' 該地區的產品總數累加
         m_ZoneList(nX).Count = m_ZoneList(nX).Count + 1
         ' 搜尋產品串列
         bFindProduct = False
         For ny = 0 To m_ZoneList(nX).ProductCount - 1
            If strProduct = m_ZoneList(nX).ProductList(ny).Product Then
               bFindProduct = True
               m_ZoneList(nX).ProductList(ny).Count = m_ZoneList(nX).ProductList(ny).Count + 1
               Exit For
            End If
         Next ny
         ' 找不到該產品類別
         If bFindProduct = False Then
            ny = m_ZoneList(nX).ProductCount
            ReDim Preserve m_ZoneList(nX).ProductList(ny + 1)
            m_ZoneList(nX).ProductList(ny).Product = strProduct
            m_ZoneList(nX).ProductList(ny).Count = 1
            m_ZoneList(nX).ProductCount = m_ZoneList(nX).ProductCount + 1
         End If
         Exit For
      End If
   Next nX
   ' 找不到地區
   If bFindZone = False Then
      nX = m_ZoneCount
      ReDim Preserve m_ZoneList(nX + 1)
      m_ZoneList(nX).ZoneCode = strZoneCode
      m_ZoneList(nX).ZoneName = strZoneName
      m_ZoneList(nX).ProductCount = 0
      m_ZoneList(nX).Count = 1
      m_ZoneCount = m_ZoneCount + 1
      ny = m_ZoneList(nX).ProductCount
      ReDim Preserve m_ZoneList(nX).ProductList(ny + 1)
      m_ZoneList(nX).ProductList(ny).Product = strProduct
      m_ZoneList(nX).ProductList(ny).Count = 1
      m_ZoneList(nX).ProductCount = m_ZoneList(nX).ProductCount + 1
   End If
End Function

' 清除商品結構串列
Private Function ClearTmpProductList()
   If m_ProductTmpCount > 0 Then
      Erase m_ProductTmpList
   End If
   m_ProductTmpCount = 0
End Function

' 建立商品結構串列
Private Function BuileTmpProductList()
   Dim nProductCount As Integer
   Dim nX As Integer
   Dim ny As Integer
   Dim nSortX As Integer
   Dim nSortY As Integer
   Dim bFindProduct As Boolean
   Dim ProductTemp As PRODUCTITEM
   
   ClearTmpProductList
   
   'For nX = 0 To Min(9, m_ZoneCount - 1)
   For nX = 0 To m_ZoneCount - 1
      For ny = 0 To m_ZoneList(nX).ProductCount - 1
         bFindProduct = False
         For nProductCount = 0 To m_ProductTmpCount - 1
            If m_ZoneList(nX).ProductList(ny).Product = m_ProductTmpList(nProductCount).Product Then
               bFindProduct = True
               m_ProductTmpList(nProductCount).Count = m_ProductTmpList(nProductCount).Count + m_ZoneList(nX).ProductList(ny).Count
               Exit For
            End If
         Next nProductCount
         If bFindProduct = False Then
            nProductCount = m_ProductTmpCount
            ReDim Preserve m_ProductTmpList(nProductCount + 1)
            m_ProductTmpList(nProductCount).Product = m_ZoneList(nX).ProductList(ny).Product
            m_ProductTmpList(nProductCount).Count = m_ZoneList(nX).ProductList(ny).Count
            m_ProductTmpCount = m_ProductTmpCount + 1
         End If
      Next ny
   Next nX
   
   ' 對商品串列依其結構中的數量由大到小排序
   For nSortX = 0 To m_ProductTmpCount - 1
      For nSortY = nSortX To m_ProductTmpCount - 1
         If m_ProductTmpList(nSortX).Count < m_ProductTmpList(nSortY).Count Then
            ProductTemp = m_ProductTmpList(nSortX)
            m_ProductTmpList(nSortX) = m_ProductTmpList(nSortY)
            m_ProductTmpList(nSortY) = ProductTemp
         ElseIf m_ProductTmpList(nSortX).Count = m_ProductTmpList(nSortY).Count Then
            If m_ProductTmpList(nSortX).Product > m_ProductTmpList(nSortY).Product Then
               ProductTemp = m_ProductTmpList(nSortX)
               m_ProductTmpList(nSortX) = m_ProductTmpList(nSortY)
               m_ProductTmpList(nSortY) = ProductTemp
            End If
         End If
      Next nSortY
   Next nSortX

End Function

Private Function GetTotalAmount() As Long
   Dim nX As Long
   Dim nAmount As Long
   nAmount = 0
   For nX = 0 To m_ZoneCount - 1
      nAmount = nAmount + m_ZoneList(nX).Count
   Next nX
   GetTotalAmount = nAmount
End Function

' 從資料庫中取得所有的資料
Private Function GetDBData_RP() As Boolean
   Dim rsMain As New ADODB.Recordset
   Dim strSql As String
   Dim strProduct As String
   Dim strZoneKind, strZoneName, strZoneCode As String
   Dim bFindZone, bFindProduct As Boolean
   Dim nSortX, nSortY As Integer
   Dim ProductTemp As PRODUCTITEM
   Dim ZoneTemp As ZONEITEM
   Dim bFromSec As Boolean
   Dim bToSec As Boolean
   Dim nCount As Integer
   Dim nX, ny, nZ As Integer
   
   GetDBData_RP = True
   
   ' 產生SQL查詢語法
   bFromSec = Not IsEmptyText(textTMBM07_1.Text)
   bToSec = Not IsEmptyText(textTMBM07_2.Text)
   'Modify By Sindy 2013/8/19 + length(na01)=3 AND
   If bFromSec = True And bToSec = True Then
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03(+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01(+) AND " & _
                     "TMBM07 >= '" & textTMBM07_1 & "' AND " & _
                     "TMBM07 <= '" & textTMBM07_2 & "' AND " & _
                     "SUBSTR(NA02, 1, 1) > 'B'"
   ElseIf bFromSec = True And bToSec = False Then
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03(+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01(+) AND " & _
                     "TMBM07 >= '" & textTMBM07_1 & "' AND " & _
                     "SUBSTR(NA02, 1, 1) > 'B'"
   ElseIf bFromSec = False And bToSec = True Then
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03(+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01(+) AND " & _
                     "TMBM07 <= '" & textTMBM07_2 & "' AND " & _
                     "SUBSTR(NA02, 1, 1) > 'B'"
   Else
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03(+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01(+) AND " & _
                     "SUBSTR(NA02, 1, 1) > 'B'"
   End If
   
   ' 取得資料庫的資料
   rsMain.CursorLocation = adUseClient
   rsMain.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 無資料則離開
   If rsMain.RecordCount <= 0 Then
      GetDBData_RP = False
      GoTo EXITSUB
   End If
   
   ' 設定初始值
   m_ZoneCount = 0
   
   rsMain.MoveFirst
   ' 依序從資料記錄中取出欄位的內容
   While Not rsMain.EOF
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
            
      strProduct = Empty
      
      If IsNull(rsMain.Fields("TMBM08")) = False Then
         nCount = GetSubStringCount(rsMain.Fields("TMBM08"))
         For nX = 1 To nCount
            strProduct = GetSubString(rsMain.Fields("TMBM08"), nX)
            If strProduct <> Empty Then
               InsertProductItem strProduct, strZoneCode, strZoneName
            End If
         Next nX
      End If
      
      rsMain.MoveNext
   Wend

   ' 排序
   ' 對地區串列依其結構中的數量由大到小排序
   For nSortX = 0 To m_ZoneCount - 1
      For nSortY = nSortX To m_ZoneCount - 1
         If m_ZoneList(nSortX).Count < m_ZoneList(nSortY).Count Then
            ZoneTemp = m_ZoneList(nSortX)
            m_ZoneList(nSortX) = m_ZoneList(nSortY)
            m_ZoneList(nSortY) = ZoneTemp
         End If
      Next nSortY
   Next nSortX
   ' 對地區結構中的商品類別串列依其商品類別中的數量由大到小排序
   For nX = 0 To m_ZoneCount - 1
      For nSortX = 0 To m_ZoneList(nX).ProductCount - 1
         For nSortY = nSortX To m_ZoneList(nX).ProductCount - 1
            If m_ZoneList(nX).ProductList(nSortX).Count < m_ZoneList(nX).ProductList(nSortY).Count Then
               ProductTemp = m_ZoneList(nX).ProductList(nSortX)
               m_ZoneList(nX).ProductList(nSortX) = m_ZoneList(nX).ProductList(nSortY)
               m_ZoneList(nX).ProductList(nSortY) = ProductTemp
            ElseIf m_ZoneList(nX).ProductList(nSortX).Count = m_ZoneList(nX).ProductList(nSortY).Count Then
               If m_ZoneList(nX).ProductList(nSortX).Product > m_ZoneList(nX).ProductList(nSortY).Product Then
                  ProductTemp = m_ZoneList(nX).ProductList(nSortX)
                  m_ZoneList(nX).ProductList(nSortX) = m_ZoneList(nX).ProductList(nSortY)
                  m_ZoneList(nX).ProductList(nSortY) = ProductTemp
               End If
            End If
         Next nSortY
      Next nSortX
   Next nX

EXITSUB:
   rsMain.Close
   Set rsMain = Nothing
End Function

' 列印表一的內容
Public Sub Generate_RP()
   Dim nRow As Integer
   Dim nPage As Integer
   Dim fld(17) As String
   Dim nAmount As Long
   Dim nTotalAmount As Long
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
   
   For nX = 0 To Min(9, m_ZoneCount - 1)
      ' 換頁
      If nRow > m_ReportDataRows Then
         Printer.NewPage
         nPage = nPage + 1
         PrintPageHeader_RP nPage
         nRow = 0
      End If
      
      ' 清除欄位
      For ny = 0 To 16: fld(ny) = Empty: Next ny
      ' 列數加一
      nRow = nRow + 1
      fld(0) = m_ZoneList(nX).ZoneName
      'fld(1) = "類別"
      fld(2) = "類別"
      For ny = 0 To Min(13, m_ZoneList(nX).ProductCount - 1)
         'fld(nY + 2) = m_ZoneList(nX).ProductList(nY).Product
         fld(ny + 3) = m_ZoneList(nX).ProductList(ny).Product
      Next ny
      ' 輸出商品類別列
      For nZ = 0 To 16
         Select Case nZ
            Case 0, 2
               Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
               Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
               Printer.Print fld(nZ)
            Case Else
               nCenter = ((m_Field(nZ).Left * m_CharWidth) + (m_Field(nZ).Left + m_Field(nZ).Width) * m_CharWidth) / 2
               Printer.CurrentX = nCenter - Printer.TextWidth(fld(nZ)) / 2
               Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
               Printer.Print fld(nZ)
         End Select
      Next nZ
      
      ' 清除欄位
      For ny = 0 To 16: fld(ny) = Empty: Next ny
      ' 列數加一
      nRow = nRow + 1
      fld(0) = Empty
      'fld(1) = "數量"
      fld(2) = "數量"
      For ny = 0 To Min(13, m_ZoneList(nX).ProductCount - 1)
         'fld(nY + 2) = m_ZoneList(nX).ProductList(nY).Count
         fld(ny + 3) = m_ZoneList(nX).ProductList(ny).Count
      Next ny
      'fld(16) = m_ZoneList(nX).Count
      fld(1) = m_ZoneList(nX).Count
      ' 輸出數量列
      For nZ = 0 To 16
         Select Case nZ
            Case 0, 2
               Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
               Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
               Printer.Print fld(nZ)
            Case Else
               nRight = (m_Field(nZ).Left + m_Field(nZ).Width - 2) * m_CharWidth
               Printer.CurrentX = nRight - Printer.TextWidth(fld(nZ))
               Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
               Printer.Print fld(nZ)
         End Select
      Next nZ
      
      ' 清除欄位
      For ny = 0 To 16: fld(ny) = Empty: Next ny
      ' 列數加一
      nRow = nRow + 1
      fld(0) = Empty
      'fld(1) = "百分比"
      fld(2) = "百分比"
      For ny = 0 To Min(13, m_ZoneList(nX).ProductCount - 1)
         nAmount = m_ZoneList(nX).ProductList(ny).Count
         nTotalAmount = m_ZoneList(nX).Count
         fValue = nAmount / nTotalAmount * 100
         'fld(nY + 2) = Format(fValue, "##0.00") & " %"
         fld(ny + 3) = Format(fValue, "##0.00") & " %"
      Next ny
      ' 輸出百分比列
      For nZ = 0 To 16
         Select Case nZ
            Case 0, 2
               Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
               Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
               Printer.Print fld(nZ)
            Case Else
               nRight = (m_Field(nZ).Left + m_Field(nZ).Width - 2) * m_CharWidth
               Printer.CurrentX = nRight - Printer.TextWidth(fld(nZ))
               Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
               Printer.Print fld(nZ)
         End Select
      Next nZ
      
      ' 列印區域別的分隔線
      nRow = nRow + 1
      PrintSplitLine m_HeaderHeight + nRow
   Next nX
   
   ' 列印前十大申請國的合計資料
   BuileTmpProductList
   
   ' 清除欄位
   For ny = 0 To 16: fld(ny) = Empty: Next ny
   ' 列數加一
   nRow = nRow + 1
   fld(0) = "合計"
   fld(2) = "類別"
   For ny = 0 To Min(13, m_ProductTmpCount - 1)
      fld(ny + 3) = m_ProductTmpList(ny).Product
   Next ny
   ' 輸出商品類別列
   For nZ = 0 To 16
      Select Case nZ
         Case 0, 2
            Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
            Printer.Print fld(nZ)
         Case Else
            nCenter = ((m_Field(nZ).Left * m_CharWidth) + (m_Field(nZ).Left + m_Field(nZ).Width) * m_CharWidth) / 2
            Printer.CurrentX = nCenter - Printer.TextWidth(fld(nZ)) / 2
            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
            Printer.Print fld(nZ)
         End Select
   Next nZ
      
   ' 清除欄位
   For ny = 0 To 16: fld(ny) = Empty: Next ny
   ' 列數加一
   nRow = nRow + 1
   fld(0) = Empty
   fld(2) = "數量"
   For ny = 0 To Min(13, m_ProductTmpCount - 1)
      fld(ny + 3) = m_ProductTmpList(ny).Count
   Next ny
   ' 計算總數量
   nTotalAmount = 0
   'For nY = 0 To m_ProductTmpCount - 1
   '   nTotalAmount = nTotalAmount + m_ProductTmpList(nY).Count
   'Next nY
   nTotalAmount = GetTotalAmount
   'fld(16) = nTotalAmount
   fld(1) = nTotalAmount
   ' 輸出數量列
   For nZ = 0 To 16
      Select Case nZ
         Case 0, 2
            Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
            Printer.Print fld(nZ)
         Case Else
            nRight = (m_Field(nZ).Left + m_Field(nZ).Width - 2) * m_CharWidth
            Printer.CurrentX = nRight - Printer.TextWidth(fld(nZ))
            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
            Printer.Print fld(nZ)
      End Select
   Next nZ
      
   ' 清除欄位
   For ny = 0 To 16: fld(ny) = Empty: Next ny
   ' 列數加一
   nRow = nRow + 1
   fld(0) = Empty
   'fld(1) = "百分比"
   fld(2) = "百分比"
   For ny = 0 To Min(13, m_ProductTmpCount - 1)
      nAmount = m_ProductTmpList(ny).Count
      fValue = nAmount / nTotalAmount * 100
      'fld(nY + 2) = Format(fValue, "##0.00") & " %"
      fld(ny + 3) = Format(fValue, "##0.00") & " %"
   Next ny
   ' 輸出百分比列
   For nZ = 0 To 16
      Select Case nZ
         Case 0, 2
            Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
            Printer.Print fld(nZ)
         Case Else
            nRight = (m_Field(nZ).Left + m_Field(nZ).Width - 2) * m_CharWidth
            Printer.CurrentX = nRight - Printer.TextWidth(fld(nZ))
            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
            Printer.Print fld(nZ)
      End Select
   Next nZ
   
   ' 列印區域別的分隔線
   nRow = nRow + 1
   PrintTerminateLine m_HeaderHeight + nRow
   
   ClearTmpProductList
      
   Printer.EndDoc

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
   CheckDataValid = True
EXITSUB:
End Function

' 列印表一的內容
Public Sub Generate_RP_SCREEN()
   Dim nRow As Integer
   Dim nPage As Integer
   Dim fld(17) As String
   Dim nAmount As Integer
   Dim nTotalAmount As Integer
   Dim fValue As Variant
   Dim nX As Integer
   Dim ny As Integer
   Dim nZ As Integer
   Dim nCenter As Long
   Dim nRight As Long
   
   nRow = 0
   For nX = 0 To Min(9, m_ZoneCount - 1)
      ' 清除欄位
      For ny = 0 To 16: fld(ny) = Empty: Next ny
      ' 列數加一
      nRow = nRow + 1
      fld(0) = m_ZoneList(nX).ZoneName
      fld(1) = "類別"
      For ny = 0 To Min(13, m_ZoneList(nX).ProductCount - 1)
         fld(ny + 2) = m_ZoneList(nX).ProductList(ny).Product
      Next ny
      
      ' 清除欄位
      For ny = 0 To 16: fld(ny) = Empty: Next ny
      ' 列數加一
      nRow = nRow + 1
      fld(0) = Empty
      fld(1) = "數量"
      For ny = 0 To Min(13, m_ZoneList(nX).ProductCount - 1)
         fld(ny + 2) = m_ZoneList(nX).ProductList(ny).Count
      Next ny
      fld(16) = m_ZoneList(nX).Count
      
      ' 清除欄位
      For ny = 0 To 16: fld(ny) = Empty: Next ny
      ' 列數加一
      nRow = nRow + 1
      fld(0) = Empty
      fld(1) = "百分比"
      For ny = 0 To Min(13, m_ZoneList(nX).ProductCount - 1)
         nAmount = m_ZoneList(nX).ProductList(ny).Count
         nTotalAmount = m_ZoneList(nX).Count
         fValue = nAmount / nTotalAmount * 100
         fld(ny + 2) = Format(fValue, "##0.00") & " %"
      Next ny
      
      ' 列印區域別的分隔線
      nRow = nRow + 1
   Next nX
   
   ' 列印前十大申請國的合計資料
   BuileTmpProductList
   
   ' 清除欄位
   For ny = 0 To 16: fld(ny) = Empty: Next ny
   ' 列數加一
   nRow = nRow + 1
   fld(0) = "合計"
   fld(1) = "類別"
   For ny = 0 To Min(13, m_ProductTmpCount - 1)
      fld(ny + 2) = m_ProductTmpList(ny).Product
   Next ny
      
   ' 清除欄位
   For ny = 0 To 16: fld(ny) = Empty: Next ny
   ' 列數加一
   nRow = nRow + 1
   fld(0) = Empty
   fld(1) = "數量"
   For ny = 0 To Min(13, m_ProductTmpCount - 1)
      fld(ny + 2) = m_ProductTmpList(ny).Count
   Next ny
   ' 計算總數量
   nTotalAmount = 0
   For ny = 0 To m_ProductTmpCount - 1
      nTotalAmount = nTotalAmount + m_ProductTmpList(ny).Count
   Next ny
   fld(16) = nTotalAmount
      
   ' 清除欄位
   For ny = 0 To 16: fld(ny) = Empty: Next ny
   ' 列數加一
   nRow = nRow + 1
   fld(0) = Empty
   fld(1) = "百分比"
   For ny = 0 To Min(13, m_ProductTmpCount - 1)
      nAmount = m_ProductTmpList(ny).Count
      fValue = nAmount / nTotalAmount * 100
      fld(ny + 2) = Format(fValue, "##0.00") & " %"
   Next ny
   
   ' 列印區域別的分隔線
   nRow = nRow + 1
   
   ClearTmpProductList
   
End Sub

