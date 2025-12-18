VERSION 5.00
Begin VB.Form frm030610 
   BorderStyle     =   1  '單線固定
   Caption         =   "表五.國外市場排名"
   ClientHeight    =   1665
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4545
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      ItemData        =   "frm030610.frx":0000
      Left            =   1410
      List            =   "frm030610.frx":0002
      TabIndex        =   2
      Top             =   1740
      Visible         =   0   'False
      Width           =   2952
   End
   Begin VB.TextBox textTMBM07_1 
      Height          =   264
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   0
      Top             =   780
      Width           =   1092
   End
   Begin VB.TextBox textTMBM07_2 
      Height          =   264
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   1
      Top             =   780
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2520
      TabIndex        =   3
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   3480
      TabIndex        =   4
      Top             =   60
      Width           =   912
   End
   Begin VB.Label Label10 
      Caption         =   "印表機 :"
      Height          =   255
      Left            =   330
      TabIndex        =   6
      Top             =   1740
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   2520
      X2              =   2760
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Label Label1 
      Caption         =   "公報卷期："
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   780
      Width           =   972
   End
End
Attribute VB_Name = "frm030610"
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
'edit by nick 2004/12/15
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

Private Type ZONEITEM
   ZoneCode As String
   ZoneName As String
   Count As Integer
   'add by nick 2004/12/15
   Count08 As Integer
End Type
' 定義地區串列
Dim m_ZoneList() As ZONEITEM
Dim m_ZoneCount As Integer
'edit by nick 2004/12/15
'Dim m_DefaultPrinter As String

Private Sub Form_Load()
'edit by nick 2004/12/15
'   Dim Prn As Printer
'edit by nick 2004/12/15
'   m_DefaultPrinter = Printer.DeviceName
   MoveFormToCenter Me
'edit by nick 2004/12/15
'   For Each Prn In Printers
'      If Prn.DeviceName <> m_DefaultPrinter Then
'         cmbPrinter.AddItem Prn.DeviceName
'      End If
'   Next
'   cmbPrinter.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
'edit by nick 2004/12/15
'   Dim Prn As Printer
'   For Each Prn In Printers
'      If Prn.DeviceName = m_DefaultPrinter Then
'         Set Printer = Prn
'         Exit For
'      End If
'   Next
   'Add By Cheng 2002/07/19
   Set frm030610 = Nothing
End Sub

' 清除所有佔用的空間
Private Sub Clear()
   Dim nX As Integer
   Dim ny As Integer
   If m_ZoneCount > 0 Then
      Erase m_ZoneList
   End If
   m_ZoneCount = 0
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   Dim Prn As Printer
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If CheckDataValid() = True Then
      '搜尋 Printer
'edit by nick 2004/12/15
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
      ' 搜尋資料
      BuildField_RP
      
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/22 清除查詢印表記錄檔欄位
      If Len(textTMBM07_1) <> 0 Or Len(textTMBM07_2) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & Label1 & textTMBM07_1 & "-" & textTMBM07_2 'Add By Sindy 2010/10/22
      End If
      
      'edit by nick 2004/12/15
      'If GetDBData_RP = False Then: GoTo EXITSUB
      If GetDBData_RP_931215 = False Then: GoTo EXITSUB
      'edit by nick 2004/12/15
      'Generate_RP
      Generate_RP_931215
      'Generate_RP_SCREEN
      InsertQueryLog ("") 'Add By Sindy 2010/10/22
      
      ' 清除暫存區
      Clear
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      
      strTit = "輸出報表"
      strMsg = "列印結束"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
    End If
    
EXITSUB:
   Screen.MousePointer = vbDefault
   Clear
End Sub

' 從資料庫中取得所有的資料
Private Function GetDBData_RP() As Boolean
   Dim rsMain As New ADODB.Recordset
   Dim strSql As String
   Dim strSubSQL As String
   Dim strZoneKind, strZoneName, strZoneCode, strAgentName, strAgentCode As String
   Dim bFindZone, bFindCountry, bFindAgent As Boolean
   Dim nSortX, nSortY As Integer
   Dim ZoneTemp As ZONEITEM
   Dim bFromSec As Boolean
   Dim bToSec As Boolean
   Dim nX, ny, nZ As Integer
   
   GetDBData_RP = True
   
   ' 產生SQL查詢語法
   bFromSec = Not IsEmptyText(textTMBM07_1.Text)
   bToSec = Not IsEmptyText(textTMBM07_2.Text)
   'Modify By Sindy 2013/8/19 + length(na01)=3 AND
   If bFromSec = True And bToSec = True Then
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01 (+) AND " & _
                     "TMBM07 >= '" & textTMBM07_1 & "' AND " & _
                     "TMBM07 <= '" & textTMBM07_2 & "' AND " & _
                     "SUBSTR(NA02,1,1) > 'B'"
   ElseIf bFromSec = True And bToSec = False Then
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01 (+) AND " & _
                     "TMBM07 >= '" & textTMBM07_1 & "' AND " & _
                     "SUBSTR(NA02,1,1) > 'B'"
   ElseIf bFromSec = False And bToSec = True Then
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01 (+) AND " & _
                     "TMBM07 <= '" & textTMBM07_2 & "'" & _
                     "SUBSTR(NA02,1,1) > 'B'"
   Else
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01 (+) AND " & _
                     "SUBSTR(NA02,1,1) > 'B'"
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
      ' 地區
      'strZoneName = ""
      'strZoneCode = ""
      'strZoneKind = ""
      'If IsNull(rsMain.Fields("TMBM05")) = False Then
      '   ' 地區名稱(國名)
      '   strZoneName = rsMain.Fields("TMBM05")
      '   ' 區域別
      '   strZoneKind = GetNationZone(rsMain.Fields("TMBM05"))
      '   ' 地區代碼(國家代碼)
      '   strZoneCode = GetNationCode(rsMain.Fields("TMBM05"))
      'End If
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
      
      ' 地區串列
      bFindZone = False
      For nX = 0 To m_ZoneCount - 1
         ' 找到地區別的結構
         If m_ZoneList(nX).ZoneCode = strZoneCode Then
            bFindZone = True
            m_ZoneList(nX).Count = m_ZoneList(nX).Count + 1
            Exit For
         End If
      Next nX
      
      ' 找不到地區別則新增地區別結構
      If bFindZone = False Then
         nX = m_ZoneCount
         ReDim Preserve m_ZoneList(nX + 1)
         m_ZoneList(nX).ZoneCode = strZoneCode
         m_ZoneList(nX).ZoneName = strZoneName
         m_ZoneList(nX).Count = 1
         m_ZoneCount = m_ZoneCount + 1
      End If
      
      ' 移到下一筆記錄
      rsMain.MoveNext
   Wend
   
   ' 對地區別串列依地區別代碼小到大排序
   For nSortX = 0 To m_ZoneCount - 1
      For nSortY = nSortX To m_ZoneCount - 1
         If m_ZoneList(nSortX).Count < m_ZoneList(nSortY).Count Then
            ZoneTemp = m_ZoneList(nSortX)
            m_ZoneList(nSortX) = m_ZoneList(nSortY)
            m_ZoneList(nSortY) = ZoneTemp
         ElseIf m_ZoneList(nSortX).Count = m_ZoneList(nSortY).Count Then
            If m_ZoneList(nSortX).ZoneCode > m_ZoneList(nSortY).ZoneCode Then
               ZoneTemp = m_ZoneList(nSortX)
               m_ZoneList(nSortX) = m_ZoneList(nSortY)
               m_ZoneList(nSortY) = ZoneTemp
            End If
         End If
      Next nSortY
   Next nSortX
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
            m_Field(nIndex).Width = 8
            m_Field(nIndex).Left = m_LeftMargin + (nIndex * nFieldWidth)
            m_Field(nIndex).Name = "排名"
         Case 1
            m_Field(nIndex).Left = m_LeftMargin + (nIndex * nFieldWidth)
            m_Field(nIndex).Name = CStr(nIndex)
         Case 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15:
            m_Field(nIndex).Name = CStr(nIndex)
         Case 16:
            m_Field(nIndex).Width = 8
            m_Field(nIndex).Name = "本期總數"
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
   Printer.Print "表五：國外市場排名表"
   
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
' 取得總數
Private Function GetTotalAmount() As Variant
   Dim nAmount As Variant
   Dim nX As Integer
   
   nAmount = 0
   For nX = 0 To m_ZoneCount - 1
      nAmount = nAmount + m_ZoneList(nX).Count
   Next nX
   GetTotalAmount = nAmount
End Function

' 列印表五的內容
Public Sub Generate_RP()
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
   
   ' 取得總數量
   nTotalAmount = GetTotalAmount
      
   nRow = 0
   
   ' 清除欄位
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   ' 列印地區
   nRow = nRow + 1
   fld(0) = "地區"
   For nX = 0 To Min(14, m_ZoneCount - 1)
      fld(nX + 1) = m_ZoneList(nX).ZoneName
   Next nX
   ' 輸出代理人列
   For nZ = 0 To 16
      Select Case nZ
         Case 0
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
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   ' 數量
   nRow = nRow + 2
   fld(0) = "數量"
   For nX = 0 To Min(14, m_ZoneCount - 1)
      fld(nX + 1) = m_ZoneList(nX).Count
   Next nX
   fld(16) = nTotalAmount
   For nZ = 0 To 16
      Select Case nZ
         Case 0
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
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   ' 比例
   nRow = nRow + 2
   fld(0) = "百分比"
   For nX = 0 To Min(15, m_ZoneCount - 1)
      nAmount = m_ZoneList(nX).Count
      fValue = nAmount / nTotalAmount * 100
      fld(nX + 1) = Format(fValue, "##0.00") & " %"
   Next nX
   'fld(16) = "100.00 %"
   fld(16) = Empty
   For nZ = 0 To 16
      Select Case nZ
         Case 0
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
   
   Printer.EndDoc

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

' 背景列印
Public Sub PrintReportBK(ByVal strPrinter As String, ByVal TMBM07_1 As String, ByVal TMBM07_2 As String)
   Dim Prn As Printer
   
   Me.Hide
   textTMBM07_1 = TMBM07_1
   textTMBM07_2 = TMBM07_2
   
   '搜尋 Printer
'edit by nick 2004/12/15
'   For Each Prn In Printers
'      If Prn.DeviceName = strPrinter Then
'         Set Printer = Prn
'         Exit For
'      End If
'   Next
   
   BuildField_RP
   'edit by nick 2004/12/15
'   If GetDBData_RP = False Then: GoTo EXITSUB
'   Generate_RP
   If GetDBData_RP_931215 = False Then: GoTo EXITSUB
   Generate_RP_931215
   ' 清除所佔用的空間
   Clear
EXITSUB:
   Set frm030610 = Nothing
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
   
   'Added by Lydia 2017/02/06 檢查公報檔的地區
   If Pub_ChkTMBMValidate(textTMBM07_1, textTMBM07_2) = False Then
      GoTo EXITSUB
   End If
   'end 2017/02/06
   
   CheckDataValid = True
EXITSUB:
End Function

' 列印表五的內容
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
   
   ' 取得總數量
   nTotalAmount = GetTotalAmount
      
   nRow = 0
   
   ' 清除欄位
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   ' 列印地區
   nRow = nRow + 1
   fld(0) = "地區"
   For nX = 0 To Min(14, m_ZoneCount - 1)
      fld(nX + 1) = m_ZoneList(nX).ZoneName
   Next nX
      
   ' 清除欄位
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   ' 數量
   nRow = nRow + 2
   fld(0) = "數量"
   For nX = 0 To Min(14, m_ZoneCount - 1)
      fld(nX + 1) = m_ZoneList(nX).Count
   Next nX
   fld(16) = nTotalAmount
      
   ' 清除欄位
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   ' 比例
   nRow = nRow + 2
   fld(0) = "百分比"
   For nX = 0 To Min(15, m_ZoneCount - 1)
      nAmount = m_ZoneList(nX).Count
      fValue = nAmount / nTotalAmount * 100
      fld(nX + 1) = Format(fValue, "##0.00") & " %"
   Next nX
   'fld(16) = "100.00 %"
   
   ' 列印區域別的分隔線
   nRow = nRow + 1
   
End Sub

Public Sub Generate_RP_931215()
   Dim nRow As Integer
   Dim nPage As Integer
   Dim fld(17) As String
   Dim nAmount As Variant
   Dim nTotalAmount As Variant
   Dim nTotalAmount08 As Variant
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
   
   ' 取得總數量
   nTotalAmount = GetTotalAmount
   nTotalAmount08 = GetTotalAmount08
      
   nRow = 0
   
   ' 清除欄位
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   ' 列印地區
   nRow = nRow + 1
   fld(0) = "地區"
   For nX = 0 To Min(14, m_ZoneCount - 1)
      fld(nX + 1) = m_ZoneList(nX).ZoneName
   Next nX
   ' 輸出代理人列
   For nZ = 0 To 16
      Select Case nZ
         Case 0
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
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   ' 數量
   nRow = nRow + 2
   fld(0) = "數量(類)"
   For nX = 0 To Min(14, m_ZoneCount - 1)
      fld(nX + 1) = m_ZoneList(nX).Count08
   Next nX
   fld(16) = nTotalAmount08
   For nZ = 0 To 16
      Select Case nZ
         Case 0
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
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   ' 比例
   nRow = nRow + 2
   fld(0) = "百分比"
   For nX = 0 To Min(15, m_ZoneCount - 1)
      nAmount = m_ZoneList(nX).Count08
      fValue = nAmount / nTotalAmount08 * 100
      fld(nX + 1) = Format(fValue, "##0.00") & " %"
   Next nX
   'fld(16) = "100.00 %"
   fld(16) = Empty
   For nZ = 0 To 16
      Select Case nZ
         Case 0
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
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   ' 數量
   nRow = nRow + 2
   fld(0) = "數量(件)"
   For nX = 0 To Min(14, m_ZoneCount - 1)
      fld(nX + 1) = m_ZoneList(nX).Count
   Next nX
   fld(16) = nTotalAmount
   For nZ = 0 To 16
      Select Case nZ
         Case 0
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
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   ' 比例
   nRow = nRow + 2
   fld(0) = "百分比"
   For nX = 0 To Min(15, m_ZoneCount - 1)
      nAmount = m_ZoneList(nX).Count
      fValue = nAmount / nTotalAmount * 100
      fld(nX + 1) = Format(fValue, "##0.00") & " %"
   Next nX
   'fld(16) = "100.00 %"
   fld(16) = Empty
   For nZ = 0 To 16
      Select Case nZ
         Case 0
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
   
   
   ' 列印區域別的分隔線
   nRow = nRow + 1
   PrintTerminateLine m_HeaderHeight + nRow
   
   Printer.EndDoc
End Sub
Private Function GetDBData_RP_931215() As Boolean
   Dim rsMain As New ADODB.Recordset
   Dim strSql As String
   Dim strSubSQL As String
   Dim strZoneKind, strZoneName, strZoneCode, strAgentName, strAgentCode As String
   Dim bFindZone, bFindCountry, bFindAgent As Boolean
   Dim nSortX, nSortY As Integer
   Dim ZoneTemp As ZONEITEM
   Dim bFromSec As Boolean
   Dim bToSec As Boolean
   Dim nX, ny, nZ As Integer
   Dim TmpArr As Variant
   Dim oStrTMBM08 As String
   
   GetDBData_RP_931215 = True
   
   ' 產生SQL查詢語法
   bFromSec = Not IsEmptyText(textTMBM07_1.Text)
   bToSec = Not IsEmptyText(textTMBM07_2.Text)
   'Modify By Sindy 2013/8/19 + length(na01)=3 AND
   If bFromSec = True And bToSec = True Then
      'edit by nickc 2005/02/14 加入大陸
      'StrSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01 (+) AND " & _
                     "TMBM07 >= '" & textTMBM07_1 & "' AND " & _
                     "TMBM07 <= '" & textTMBM07_2 & "' AND " & _
                     "SUBSTR(NA02,1,1) > 'B'"
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,decode(substr(na02,1,1),'B','大陸',TMBM05) TMBM05,TMBM06,TMBM07,TMBM08,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,decode(substr(NA02,1,1),'B','020',na01) na01,decode(substr(na02,1,1),'B','B00',NA02) na02 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01 (+) AND " & _
                     "TMBM07 >= '" & textTMBM07_1 & "' AND " & _
                     "TMBM07 <= '" & textTMBM07_2 & "' AND " & _
                     "SUBSTR(NA02,1,1) > 'A'"
   ElseIf bFromSec = True And bToSec = False Then
      'edit by nickc 2005/02/14 加入大陸
      'StrSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01 (+) AND " & _
                     "TMBM07 >= '" & textTMBM07_1 & "' AND " & _
                     "SUBSTR(NA02,1,1) > 'B'"
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,decode(substr(na02,1,1),'B','大陸',TMBM05) TMBM05,TMBM06,TMBM07,TMBM08,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,decode(substr(NA02,1,1),'B','020',na01) na01,decode(substr(na02,1,1),'B','B00',NA02) na02 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01 (+) AND " & _
                     "TMBM07 >= '" & textTMBM07_1 & "' AND " & _
                     "SUBSTR(NA02,1,1) > 'A'"
   ElseIf bFromSec = False And bToSec = True Then
      'edit by nickc 2005/02/14 加入大陸
      'StrSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01 (+) AND " & _
                     "TMBM07 <= '" & textTMBM07_2 & "'" & _
                     "SUBSTR(NA02,1,1) > 'B'"
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,decode(substr(na02,1,1),'B','大陸',TMBM05) TMBM05,TMBM06,TMBM07,TMBM08,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,decode(substr(NA02,1,1),'B','020',na01) na01,decode(substr(na02,1,1),'B','B00',NA02) na02 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01 (+) AND " & _
                     "TMBM07 <= '" & textTMBM07_2 & "'" & _
                     "SUBSTR(NA02,1,1) > 'A'"
   Else
      'edit by nickc 2005/02/14 加入大陸
      'StrSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01 (+) AND " & _
                     "SUBSTR(NA02,1,1) > 'B'"
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,decode(substr(na02,1,1),'B','大陸',TMBM05) TMBM05,TMBM06,TMBM07,TMBM08,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,decode(substr(NA02,1,1),'B','020',na01) na01,decode(substr(na02,1,1),'B','B00',NA02) na02 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01 (+) AND " & _
                     "SUBSTR(NA02,1,1) > 'A'"
   End If
   
   ' 取得資料庫的資料
   rsMain.CursorLocation = adUseClient
   rsMain.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 無資料則離開
   If rsMain.RecordCount <= 0 Then
      ShowNoData
      Screen.MousePointer = vbDefault
      GetDBData_RP_931215 = False
      GoTo EXITSUB
   End If
   
   ' 設定初始值
   m_ZoneCount = 0
   
   rsMain.MoveFirst
   ' 依序從資料記錄中取出欄位的內容
   While Not rsMain.EOF
      ' 地區
      'strZoneName = ""
      'strZoneCode = ""
      'strZoneKind = ""
      'If IsNull(rsMain.Fields("TMBM05")) = False Then
      '   ' 地區名稱(國名)
      '   strZoneName = rsMain.Fields("TMBM05")
      '   ' 區域別
      '   strZoneKind = GetNationZone(rsMain.Fields("TMBM05"))
      '   ' 地區代碼(國家代碼)
      '   strZoneCode = GetNationCode(rsMain.Fields("TMBM05"))
      'End If
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
      oStrTMBM08 = "" & rsMain.Fields("TMBM08")
      TmpArr = Split(oStrTMBM08, ",")
      ' 地區串列
      bFindZone = False
      For nX = 0 To m_ZoneCount - 1
         ' 找到地區別的結構
         If m_ZoneList(nX).ZoneCode = strZoneCode Then
            bFindZone = True
            m_ZoneList(nX).Count = m_ZoneList(nX).Count + 1
            m_ZoneList(nX).Count08 = m_ZoneList(nX).Count08 + IIf(UBound(TmpArr) < 1, 1, UBound(TmpArr) + 1)
            Exit For
         End If
      Next nX
      
      ' 找不到地區別則新增地區別結構
      If bFindZone = False Then
         nX = m_ZoneCount
         ReDim Preserve m_ZoneList(nX + 1)
         m_ZoneList(nX).ZoneCode = strZoneCode
         m_ZoneList(nX).ZoneName = strZoneName
         m_ZoneList(nX).Count = 1
         m_ZoneList(nX).Count08 = IIf(UBound(TmpArr) < 1, 1, UBound(TmpArr) + 1)
         m_ZoneCount = m_ZoneCount + 1
      End If
      
      ' 移到下一筆記錄
      rsMain.MoveNext
   Wend
   
   ' 對地區別串列依地區別代碼小到大排序
   For nSortX = 0 To m_ZoneCount - 1
      For nSortY = nSortX To m_ZoneCount - 1
         If m_ZoneList(nSortX).Count08 < m_ZoneList(nSortY).Count08 Then
            ZoneTemp = m_ZoneList(nSortX)
            m_ZoneList(nSortX) = m_ZoneList(nSortY)
            m_ZoneList(nSortY) = ZoneTemp
         ElseIf m_ZoneList(nSortX).Count08 = m_ZoneList(nSortY).Count08 Then
            If m_ZoneList(nSortX).ZoneCode > m_ZoneList(nSortY).ZoneCode Then
               ZoneTemp = m_ZoneList(nSortX)
               m_ZoneList(nSortX) = m_ZoneList(nSortY)
               m_ZoneList(nSortY) = ZoneTemp
            End If
         End If
      Next nSortY
   Next nSortX
EXITSUB:
   rsMain.Close
   Set rsMain = Nothing
End Function

Private Function GetTotalAmount08() As Variant
   Dim nAmount As Variant
   Dim nX As Integer
   
   nAmount = 0
   For nX = 0 To m_ZoneCount - 1
      nAmount = nAmount + m_ZoneList(nX).Count08
   Next nX
   GetTotalAmount08 = nAmount
End Function
