VERSION 5.00
Begin VB.Form frm030613 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人國外案件排名分析表"
   ClientHeight    =   5130
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   4305
   Begin VB.ComboBox cmbPrinter 
      Height          =   276
      ItemData        =   "frm030613.frx":0000
      Left            =   1440
      List            =   "frm030613.frx":0002
      TabIndex        =   12
      Top             =   4620
      Visible         =   0   'False
      Width           =   2652
   End
   Begin VB.TextBox textAgent 
      Height          =   264
      Index           =   9
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   11
      Top             =   4260
      Width           =   2652
   End
   Begin VB.TextBox textAgent 
      Height          =   264
      Index           =   8
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   10
      Top             =   3900
      Width           =   2652
   End
   Begin VB.TextBox textAgent 
      Height          =   264
      Index           =   7
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   9
      Top             =   3540
      Width           =   2652
   End
   Begin VB.TextBox textAgent 
      Height          =   264
      Index           =   6
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   8
      Top             =   3180
      Width           =   2652
   End
   Begin VB.TextBox textAgent 
      Height          =   264
      Index           =   5
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   7
      Top             =   2820
      Width           =   2652
   End
   Begin VB.TextBox textAgent 
      Height          =   264
      Index           =   4
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   6
      Top             =   2460
      Width           =   2652
   End
   Begin VB.TextBox textAgent 
      Height          =   264
      Index           =   3
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   5
      Top             =   2100
      Width           =   2652
   End
   Begin VB.TextBox textAgent 
      Height          =   264
      Index           =   2
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1740
      Width           =   2652
   End
   Begin VB.TextBox textAgent 
      Height          =   264
      Index           =   1
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1380
      Width           =   2652
   End
   Begin VB.TextBox textAgent 
      Height          =   264
      Index           =   0
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1020
      Width           =   2652
   End
   Begin VB.TextBox textTMBM07_1 
      Height          =   264
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   0
      Top             =   660
      Width           =   1092
   End
   Begin VB.TextBox textTMBM07_2 
      Height          =   264
      Left            =   3000
      MaxLength       =   5
      TabIndex        =   1
      Top             =   660
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2340
      TabIndex        =   13
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   3300
      TabIndex        =   14
      Top             =   60
      Width           =   912
   End
   Begin VB.Label Label10 
      Caption         =   "印表機："
      Height          =   252
      Left            =   240
      TabIndex        =   17
      Top             =   4620
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "事務所名稱："
      Height          =   252
      Left            =   240
      TabIndex        =   16
      Top             =   1020
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "公報卷期："
      Height          =   252
      Left            =   240
      TabIndex        =   15
      Top             =   660
      Width           =   972
   End
   Begin VB.Line Line1 
      X1              =   2640
      X2              =   2880
      Y1              =   780
      Y2              =   780
   End
End
Attribute VB_Name = "frm030613"
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

' 地區資料結構
Private Type ZONETITEM
   ZoneCode As String
   ZoneName As String
   Count As Integer
   Count08 As Integer 'Add By Sindy 2011/3/9
End Type

' 宣告代理人項目的資料型態
Private Type AGENTITEM
   ' 代理人代碼
   AgentCode As String
   AgentName As String
   AgentCompany As String
   Count As Integer
   Count08 As Integer 'Add By Sindy 2011/3/9
   ZoneList() As ZONETITEM
   ZoneCount As Integer
End Type

' 定義地區串列
Dim m_AgentList() As AGENTITEM
Dim m_AgentCount As Integer
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
   Set frm030613 = Nothing
End Sub

' 清除系統所佔用的記憶體
Private Sub Clear()
   Dim nX As Integer
   If m_AgentCount > 0 Then
      For nX = 0 To m_AgentCount - 1
         If m_AgentList(nX).ZoneCount > 0 Then
            Erase m_AgentList(nX).ZoneList
         End If
         m_AgentList(nX).ZoneCount = 0
      Next nX
      Erase m_AgentList
   End If
   m_AgentCount = 0
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
      ' 取得資料庫中的資料
      If GetDBData_RP = False Then
         GoTo EXITSUB
      End If
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
            m_Field(nIndex).Name = "事務所"
         Case 1:
            m_Field(nIndex).Left = m_LeftMargin + (nIndex * nFieldWidth)
            m_Field(nIndex).Name = "總數"
         Case 2:
            m_Field(nIndex).Width = 8
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
   nX = m_LeftMargin + m_ReportWidth / 2 - 29 '24
   Printer.CurrentX = nX * m_CharWidth
   Printer.Print "代理人國外案件排行分析表"
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
   nX = m_LeftMargin + m_ReportWidth - 38 ' 20
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "頁"

   nX = m_LeftMargin + m_ReportWidth - 32 ' 14
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

'件總計
Private Function GetForeignAmount() As Long
   Dim rsMain As New ADODB.Recordset
   Dim strSql As String
   Dim bFromSec As Boolean
   Dim bToSec As Boolean
   Dim nCount As Integer
   
   GetForeignAmount = 0
   ' 產生SQL查詢語法
   bFromSec = Not IsEmptyText(textTMBM07_1.Text)
   bToSec = Not IsEmptyText(textTMBM07_2.Text)
   'Modify By Sindy 2013/8/19 + length(na01)=3 AND
   If bFromSec = True And bToSec = True Then
      strSql = "SELECT COUNT(*) FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03(+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03 AND " & _
                     "'T' = TA01 AND " & _
                     "TMBM07 >= '" & textTMBM07_1 & "' AND " & _
                     "TMBM07 <= '" & textTMBM07_2 & "' AND " & _
                     "SUBSTR(NA02, 1, 1) > 'B' "
   ElseIf bFromSec = True And bToSec = False Then
      strSql = "SELECT COUNT(*) FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03 AND " & _
                     "'T' = TA01 AND " & _
                     "TMBM07 >= '" & textTMBM07_1 & "' AND " & _
                     "SUBSTR(NA02, 1, 1) > 'B' "
   ElseIf bFromSec = False And bToSec = True Then
      strSql = "SELECT COUNT(*) FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03 AND " & _
                     "'T' = TA01 AND " & _
                     "TMBM07 <= '" & textTMBM07_2 & "' AND " & _
                     "SUBSTR(NA02, 1, 1) > 'B' "
   Else
      strSql = "SELECT COUNT(*) FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03 AND " & _
                     "'T' = TA01 AND " & _
                     "SUBSTR(NA02, 1, 1) > 'B' "
   End If
   
   ' 取得資料庫的資料
   rsMain.CursorLocation = adUseClient
   rsMain.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 無資料則離開
   If rsMain.RecordCount > 0 Then
      If IsNull(rsMain.Fields(0)) = False Then
         GetForeignAmount = rsMain.Fields(0)
      End If
   End If
   rsMain.Close
   Set rsMain = Nothing
End Function

'Add By Sindy 2011/3/9
'類總計
Private Function GetForeignAmount08() As Long
   Dim rsMain As New ADODB.Recordset
   Dim strSql As String
   Dim bFromSec As Boolean
   Dim bToSec As Boolean
   Dim nCount As Integer
   Dim TmpArr As Variant
   Dim oStrTMBM08 As String
   
   GetForeignAmount08 = 0
   ' 產生SQL查詢語法
   bFromSec = Not IsEmptyText(textTMBM07_1.Text)
   bToSec = Not IsEmptyText(textTMBM07_2.Text)
   'Modify By Sindy 2013/8/19 + length(na01)=3 AND
   If bFromSec = True And bToSec = True Then
      strSql = "SELECT TMBM08 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03(+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03 AND " & _
                     "'T' = TA01 AND " & _
                     "TMBM07 >= '" & textTMBM07_1 & "' AND " & _
                     "TMBM07 <= '" & textTMBM07_2 & "' AND " & _
                     "SUBSTR(NA02, 1, 1) > 'B' "
   ElseIf bFromSec = True And bToSec = False Then
      strSql = "SELECT TMBM08 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03 AND " & _
                     "'T' = TA01 AND " & _
                     "TMBM07 >= '" & textTMBM07_1 & "' AND " & _
                     "SUBSTR(NA02, 1, 1) > 'B' "
   ElseIf bFromSec = False And bToSec = True Then
      strSql = "SELECT TMBM08 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03 AND " & _
                     "'T' = TA01 AND " & _
                     "TMBM07 <= '" & textTMBM07_2 & "' AND " & _
                     "SUBSTR(NA02, 1, 1) > 'B' "
   Else
      strSql = "SELECT TMBM08 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03 AND " & _
                     "'T' = TA01 AND " & _
                     "SUBSTR(NA02, 1, 1) > 'B' "
   End If
   
   ' 取得資料庫的資料
   rsMain.CursorLocation = adUseClient
   rsMain.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 無資料則離開
   If rsMain.RecordCount > 0 Then
      rsMain.MoveFirst
      Do While Not rsMain.EOF
         oStrTMBM08 = "" & rsMain.Fields("TMBM08")
         TmpArr = Split(oStrTMBM08, ",")
         GetForeignAmount08 = GetForeignAmount08 + IIf(UBound(TmpArr) < 1, 1, UBound(TmpArr) + 1)
         rsMain.MoveNext
      Loop
   End If
   rsMain.Close
   Set rsMain = Nothing
End Function

' 從資料庫中取得所有的資料
Private Function GetDBData_RP() As Boolean
   Dim rsMain As New ADODB.Recordset
   Dim strSql As String
   Dim strSubSQL As String
   Dim strZoneName, strZoneCode As String
   Dim strAgentName, strAgentCode, strAgentCompany As String
   Dim bFindAgent As Boolean
   Dim bFindZone As Boolean
   Dim nSortX, nSortY As Integer
   Dim AgentTemp As AGENTITEM
   Dim ZoneTemp As ZONETITEM
   Dim bFromSec As Boolean
   Dim bToSec As Boolean
   Dim nCount As Integer
   Dim nX, ny, nZ As Integer
   Dim bInclude As Boolean
   Dim nIndex As Integer
   'Add By Sindy 2011/3/9
   Dim TmpArr As Variant
   Dim oStrTMBM08 As String
   '2011/3/9 End
   
   GetDBData_RP = True
   
   If Len(textTMBM07_1) <> 0 Or Len(textTMBM07_2) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1 & textTMBM07_1 & "-" & textTMBM07_2 'Add By Sindy 2010/10/22
   End If
   
   ' 產生SQL查詢語法
   bFromSec = Not IsEmptyText(textTMBM07_1.Text)
   bToSec = Not IsEmptyText(textTMBM07_2.Text)
   'Modify By Sindy 2011/3/9 +TMBM08
   'Modify By Sindy 2013/8/19 + length(na01)=3 AND
   If bFromSec = True And bToSec = True Then
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02,TMBM08 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03(+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03 AND " & _
                     "'T' = TA01 AND " & _
                     "TMBM07 >= '" & textTMBM07_1 & "' AND " & _
                     "TMBM07 <= '" & textTMBM07_2 & "' AND " & _
                     "SUBSTR(NA02, 1, 1) > 'B' "
   ElseIf bFromSec = True And bToSec = False Then
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02,TMBM08 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03 AND " & _
                     "'T' = TA01 AND " & _
                     "TMBM07 >= '" & textTMBM07_1 & "' AND " & _
                     "SUBSTR(NA02, 1, 1) > 'B' "
   ElseIf bFromSec = False And bToSec = True Then
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02,TMBM08 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03 AND " & _
                     "'T' = TA01 AND " & _
                     "TMBM07 <= '" & textTMBM07_2 & "' AND " & _
                     "SUBSTR(NA02, 1, 1) > 'B' "
   Else
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04,NA01,NA02,TMBM08 FROM TMBULLETIN, TAGENT, NATION " & _
               "WHERE TMBM05 = NA03 (+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03 AND " & _
                     "'T' = TA01 AND " & _
                     "SUBSTR(NA02, 1, 1) > 'B' "
   End If
   
   ' 依序加入比對代理人名稱的條件
   strSubSQL = Empty
   'Add By Sindy 2010/10/22
   Dim strText As String
   strText = ""
   '2010/10/22 End
   For nIndex = 0 To 9
      If IsEmptyText(textAgent(nIndex)) = False Then
         If strSubSQL <> Empty Then: strSubSQL = strSubSQL & "OR "
         'Modify By Sindy 2010/12/28 原本以代理人名稱抓資料, 改成以事務所名稱抓資料
         'strSubSQL = strSubSQL & "TMBM06 = '" & textAgent(nIndex) & "' "
         'Modify By Sindy 2011/3/7
         'strSubSQL = strSubSQL & "substr(TA04,1,4) = '" & textAgent(nIndex) & "' "
         strSubSQL = strSubSQL & " TA04 like '%" & textAgent(nIndex) & "%' "
         '2010/12/28 End
         'Add By Sindy 2010/10/22
         If Len(strText) <> 0 Then strText = strText & "、"
         strText = strText & textAgent(nIndex)
         '2010/10/22 End
      End If
   Next nIndex
   If IsEmptyText(strSubSQL) = False Then
      strSql = strSql & "AND (" & strSubSQL & ")"
   End If
   If Len(strText) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label2 & strText 'Add By Sindy 2010/10/22
   End If
   
   ' 取得資料庫的資料
   rsMain.CursorLocation = adUseClient
   rsMain.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 無資料則離開
   If rsMain.RecordCount <= 0 Then
      InsertQueryLog (0) 'Add By Sindy 2010/10/22
      GetDBData_RP = False
      GoTo EXITSUB
   End If
   
   ' 設定初始值
   m_AgentCount = 0
   
   rsMain.MoveFirst
   InsertQueryLog (rsMain.RecordCount) 'Add By Sindy 2010/10/22
   ' 依序從資料記錄中取出欄位的內容
   While Not rsMain.EOF
      ' 代理人姓名
      strAgentName = Empty
      If IsNull(rsMain.Fields("TMBM06")) = False Then
         strAgentName = rsMain.Fields("TMBM06")
      End If
      ' 代理人代碼
      strAgentCode = Empty
      If IsNull(rsMain.Fields("TA02")) = False Then
         strAgentCode = rsMain.Fields("TA02")
      End If
      ' 代理人事務所名稱
      strAgentCompany = Empty
      If IsNull(rsMain.Fields("TA04")) = False Then
         strAgentCompany = rsMain.Fields("TA04")
      End If
      ' 國籍代碼及名稱
      'strZoneCode = Empty
      'strZoneName = Empty
      'If IsNull(rsMain.Fields("TMBM05")) = False Then
      '   strZoneName = rsMain.Fields("TMBM05")
      '   strZoneCode = GetNationCode(strZoneName)
      'End If
      ' 地區名稱
      strZoneName = Empty
      If IsNull(rsMain.Fields("TMBM05")) = False Then
         strZoneName = rsMain.Fields("TMBM05")
      End If
      ' 地區代碼
      strZoneCode = Empty
      If IsNull(rsMain.Fields("NA01")) = False Then
         strZoneCode = rsMain.Fields("NA01")
      End If
      'Add By Sindy 2011/3/9
      oStrTMBM08 = "" & rsMain.Fields("TMBM08")
      TmpArr = Split(oStrTMBM08, ",")
      '2011/3/9 End
      
      ' 檢查相關資訊判斷是否要計入該筆資料
      bInclude = True
      If bInclude = True Then
         bFindAgent = False
         For nX = 0 To m_AgentCount - 1
            'If m_AgentList(nX).AgentCode = strAgentCode Then
            'Modify By Sindy 2010/12/28 原本以代理人名稱抓資料, 改成以事務所名稱抓資料
            'If m_AgentList(nX).AgentName = strAgentName Then
            If m_AgentList(nX).AgentCompany = strAgentCompany Then
            '2010/12/28 End
               bFindAgent = True
               bFindZone = False
               For ny = 0 To m_AgentList(nX).ZoneCount - 1
                  If m_AgentList(nX).ZoneList(ny).ZoneCode = strZoneCode Then
                     bFindZone = True
                     m_AgentList(nX).Count = m_AgentList(nX).Count + 1
                     'Add By Sindy 2011/3/9
                     m_AgentList(nX).Count08 = m_AgentList(nX).Count08 + IIf(UBound(TmpArr) < 1, 1, UBound(TmpArr) + 1)
                     '2011/3/9 End
                     m_AgentList(nX).ZoneList(ny).Count = m_AgentList(nX).ZoneList(ny).Count + 1
                     'Add By Sindy 2011/3/9
                     m_AgentList(nX).ZoneList(ny).Count08 = m_AgentList(nX).ZoneList(ny).Count08 + IIf(UBound(TmpArr) < 1, 1, UBound(TmpArr) + 1)
                     '2011/3/9 End
                     Exit For
                  End If
               Next ny
               If bFindZone = False Then
                  ny = m_AgentList(nX).ZoneCount
                  ReDim Preserve m_AgentList(nX).ZoneList(ny + 1)
                  m_AgentList(nX).ZoneList(ny).ZoneCode = strZoneCode
                  m_AgentList(nX).ZoneList(ny).ZoneName = strZoneName
                  m_AgentList(nX).ZoneList(ny).Count = 1
                  'Add By Sindy 2011/3/9
                  m_AgentList(nX).ZoneList(ny).Count08 = IIf(UBound(TmpArr) < 1, 1, UBound(TmpArr) + 1)
                  '2011/3/9 End
                  m_AgentList(nX).ZoneCount = m_AgentList(nX).ZoneCount + 1
                  m_AgentList(nX).Count = m_AgentList(nX).Count + 1
                  'Add By Sindy 2011/3/9
                  m_AgentList(nX).Count08 = m_AgentList(nX).Count08 + IIf(UBound(TmpArr) < 1, 1, UBound(TmpArr) + 1)
                  '2011/3/9 End
               End If
               Exit For
            End If
         Next nX
         If bFindAgent = False Then
            nX = m_AgentCount
            ReDim Preserve m_AgentList(nX + 1)
            m_AgentList(nX).AgentCode = strAgentCode
            m_AgentList(nX).AgentName = strAgentName
            m_AgentList(nX).AgentCompany = strAgentCompany
            m_AgentList(nX).Count = 1
            'Add By Sindy 2011/3/9
            m_AgentList(nX).Count08 = IIf(UBound(TmpArr) < 1, 1, UBound(TmpArr) + 1)
            '2011/3/9 End
            m_AgentList(nX).ZoneCount = 0
            m_AgentCount = m_AgentCount + 1
            ny = m_AgentList(nX).ZoneCount
            ReDim Preserve m_AgentList(nX).ZoneList(ny + 1)
            m_AgentList(nX).ZoneList(ny).ZoneCode = strZoneCode
            m_AgentList(nX).ZoneList(ny).ZoneName = strZoneName
            m_AgentList(nX).ZoneList(ny).Count = 1
            'Add By Sindy 2011/3/9
            m_AgentList(nX).ZoneList(ny).Count08 = IIf(UBound(TmpArr) < 1, 1, UBound(TmpArr) + 1)
            '2011/3/9 End
            m_AgentList(nX).ZoneCount = m_AgentList(nX).ZoneCount + 1
         End If
      End If
      
      rsMain.MoveNext
   Wend
   
   ' 對代理人排順序
   For nX = 0 To m_AgentCount - 1
      For ny = nX To m_AgentCount - 1
         If m_AgentList(nX).Count < m_AgentList(ny).Count Then
            AgentTemp = m_AgentList(nX)
            m_AgentList(nX) = m_AgentList(ny)
            m_AgentList(ny) = AgentTemp
         End If
      Next ny
   Next nX
   
   ' 對代理人中的國家排順序
   For nZ = 0 To m_AgentCount - 1
      For nX = 0 To m_AgentList(nZ).ZoneCount - 1
         For ny = nX To m_AgentList(nZ).ZoneCount - 1
            If m_AgentList(nZ).ZoneList(nX).Count < m_AgentList(nZ).ZoneList(ny).Count Then
               ZoneTemp = m_AgentList(nZ).ZoneList(nX)
               m_AgentList(nZ).ZoneList(nX) = m_AgentList(nZ).ZoneList(ny)
               m_AgentList(nZ).ZoneList(ny) = ZoneTemp
            ElseIf m_AgentList(nZ).ZoneList(nX).Count = m_AgentList(nZ).ZoneList(ny).Count Then
               If m_AgentList(nZ).ZoneList(nX).ZoneCode > m_AgentList(nZ).ZoneList(ny).ZoneCode Then
                  ZoneTemp = m_AgentList(nZ).ZoneList(nX)
                  m_AgentList(nZ).ZoneList(nX) = m_AgentList(nZ).ZoneList(ny)
                  m_AgentList(nZ).ZoneList(ny) = ZoneTemp
               End If
            End If
         Next ny
      Next nX
   Next nZ
   
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
   Dim nTotalCount As Long
   Dim nTotalCount08 As Long 'Add By Sindy 2011/3/9
   Dim fValue As Double
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
   
   nTotalCount = GetForeignAmount()
   nTotalCount08 = GetForeignAmount08() 'Add By Sindy 2011/3/9
   
   For nX = 0 To m_AgentCount - 1
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
      fld(0) = m_AgentList(nX).AgentCompany
      fld(2) = "地區"
      For ny = 0 To Min(13, m_AgentList(nX).ZoneCount - 1)
         fld(ny + 3) = m_AgentList(nX).ZoneList(ny).ZoneName
      Next ny
      ' 地區
      For nZ = 0 To 16
         Select Case nZ
            Case 0
               Printer.FontSize = 8
               Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
               Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
               Printer.Print fld(nZ)
               Printer.FontSize = 12
            Case 2
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
      
      'Add By Sindy 2011/3/9
      ' 清除欄位
      For ny = 0 To 16: fld(ny) = Empty: Next ny
      ' 列數加一
      nRow = nRow + 1
      fld(0) = "　　(類)"
      fld(1) = m_AgentList(nX).Count08
      fld(2) = "數量"
      For ny = 0 To Min(13, m_AgentList(nX).ZoneCount - 1)
         fld(ny + 3) = m_AgentList(nX).ZoneList(ny).Count08
      Next ny
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
      
      'Add By Sindy 2011/3/9
      ' 清除欄位
      For ny = 0 To 16: fld(ny) = Empty: Next ny
      ' 列數加一
      nRow = nRow + 1
      fld(0) = Empty '(類)
      For ny = 0 To Min(13, m_AgentList(nX).ZoneCount - 1)
         nAmount = m_AgentList(nX).ZoneList(ny).Count08
         nTotalAmount = m_AgentList(nX).Count08
         fValue = nAmount / nTotalAmount * 100
         fld(ny + 3) = Format(fValue, "##0.00") & " %"
      Next ny
      fld(1) = Format(m_AgentList(nX).Count08 / nTotalCount08 * 100, "##0.00") & " %"
      fld(2) = "百分比"
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
      
      ' 清除欄位
      For ny = 0 To 16: fld(ny) = Empty: Next ny
      ' 列數加一
      nRow = nRow + 1
      fld(0) = "　　(件)"
      fld(1) = m_AgentList(nX).Count
      fld(2) = "數量"
      For ny = 0 To Min(13, m_AgentList(nX).ZoneCount - 1)
         fld(ny + 3) = m_AgentList(nX).ZoneList(ny).Count
      Next ny
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
      fld(0) = Empty '(件)
      For ny = 0 To Min(13, m_AgentList(nX).ZoneCount - 1)
         nAmount = m_AgentList(nX).ZoneList(ny).Count
         nTotalAmount = m_AgentList(nX).Count
         fValue = nAmount / nTotalAmount * 100
         fld(ny + 3) = Format(fValue, "##0.00") & " %"
      Next ny
      fld(1) = Format(m_AgentList(nX).Count / nTotalCount * 100, "##0.00") & " %"
      fld(2) = "百分比"
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
   
   ' 列印國外總數
   ' 清除欄位
   For ny = 0 To 16: fld(ny) = Empty: Next ny
   ' 列數加一
   nRow = nRow + 1
   fld(0) = "國外總數"
   fld(1) = "(類)"
   fld(2) = nTotalCount08
   fld(3) = "(件)"
   fld(4) = nTotalCount
   fld(5) = "（含無代理人件數）" 'Modify By Sindy 2010/7/5 增加註解
   ' 輸出百分比列
   For nZ = 0 To 5
      Select Case nZ
         Case 0, 5
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
   
   ' 列數加一
   nRow = nRow + 1
   PrintTerminateLine m_HeaderHeight + nRow
   Printer.EndDoc
End Sub

' 事務所名稱
Private Sub textAgent_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textAgent(Index)) = False Then
      If CheckLengthIsOK(textAgent(Index), 8) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "事務所名稱內容太長"
         'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textAgent_GotFocus Index
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textAgent(Index).IMEMode = 2
   If Cancel = False Then CloseIme
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

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim bFind As Boolean
   Dim nIndex As Integer
   CheckDataValid = False
   
   ' 事務所名稱不可全為空白
   bFind = False
   For nIndex = 0 To 9
      If IsEmptyText(textAgent(nIndex)) = False Then
         bFind = True
         Exit For
      End If
   Next nIndex
   If bFind = False Then
      strTit = "檢核資料"
      strMsg = "請輸入事務所名稱"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textAgent(0).SetFocus
      GoTo EXITSUB
   End If
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textAgent_GotFocus(Index As Integer)
   InverseTextBox textAgent(Index)
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textAgent(Index).IMEMode = 1
   OpenIme
End Sub

Private Sub textTMBM07_1_GotFocus()
   InverseTextBox textTMBM07_1
End Sub

Private Sub textTMBM07_2_GotFocus()
   InverseTextBox textTMBM07_2
End Sub

' 列印表一的內容
Public Sub Generate_RP_SCREEN()
   Dim nRow As Integer
   Dim nPage As Integer
   Dim fld(17) As String
   Dim nAmount As Long
   Dim nTotalAmount As Long
   Dim fValue As Double
   Dim nX As Integer
   Dim ny As Integer
   Dim nZ As Integer
   Dim nCenter As Long
   Dim nRight As Long
   
   nRow = 0
   
   For nX = 0 To m_AgentCount - 1
      
      ' 清除欄位
      For ny = 0 To 16: fld(ny) = Empty: Next ny
      ' 列數加一
      nRow = nRow + 1
      fld(0) = m_AgentList(nX).AgentCompany
      fld(1) = "地區"
      For ny = 0 To Min(13, m_AgentList(nX).ZoneCount - 1)
         fld(ny + 2) = m_AgentList(nX).ZoneList(ny).ZoneName
      Next ny
      
      ' 清除欄位
      For ny = 0 To 16: fld(ny) = Empty: Next ny
      ' 列數加一
      nRow = nRow + 1
      fld(0) = Empty
      fld(1) = "數量"
      For ny = 0 To Min(13, m_AgentList(nX).ZoneCount - 1)
         fld(ny + 2) = m_AgentList(nX).ZoneList(ny).Count
      Next ny
      fld(16) = m_AgentList(nX).Count
      
      ' 清除欄位
      For ny = 0 To 16: fld(ny) = Empty: Next ny
      ' 列數加一
      nRow = nRow + 1
      fld(0) = Empty
      fld(1) = "百分比"
      For ny = 0 To Min(13, m_AgentList(nX).ZoneCount - 1)
         nAmount = m_AgentList(nX).ZoneList(ny).Count
         nTotalAmount = m_AgentList(nX).Count
         fValue = nAmount / nTotalAmount * 100
         fld(ny + 2) = Format(fValue, "##0.00") & " %"
      Next ny
      
      ' 列印區域別的分隔線
      nRow = nRow + 1
   Next nX
   
End Sub

