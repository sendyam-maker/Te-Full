VERSION 5.00
Begin VB.Form frm030608 
   BorderStyle     =   1  '單線固定
   Caption         =   "表三.各區市場佔有率統計表"
   ClientHeight    =   2330
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   5510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2330
   ScaleWidth      =   5510
   Begin VB.TextBox textNA03 
      Height          =   264
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1440
      Width           =   255
   End
   Begin VB.TextBox textNA02 
      Height          =   264
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1080
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
   Begin VB.TextBox textTMBM07_2 
      Height          =   264
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   1
      Top             =   720
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3480
      TabIndex        =   5
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   4440
      TabIndex        =   6
      Top             =   60
      Width           =   912
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      ItemData        =   "frm030608.frx":0000
      Left            =   1320
      List            =   "frm030608.frx":0002
      TabIndex        =   4
      Top             =   2370
      Visible         =   0   'False
      Width           =   3972
   End
   Begin VB.Label Label5 
      Caption         =   "(1:北區 2:中區 3: 南區 4:高區 5:東區)"
      Height          =   255
      Left            =   1740
      TabIndex        =   12
      Top             =   1470
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "台灣各區："
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1470
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "(A:國內 B:大陸 C:國外 空白:全部)"
      Height          =   252
      Left            =   2520
      TabIndex        =   10
      Top             =   1080
      Width           =   2772
   End
   Begin VB.Label Label2 
      Caption         =   "列印區域："
      Height          =   252
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "公報卷期："
      Height          =   252
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   972
   End
   Begin VB.Line Line1 
      X1              =   2520
      X2              =   2760
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label10 
      Caption         =   "印表機 :"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2370
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frm030608"
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

' 宣告代理人項目的資料型態
Private Type AGENTITEM
   AgentCode As String
   AgentName As String
   AgentCompany As String
   Count As Integer
   'add by nick 2004/12/14
   Count08 As Integer
End Type
' 當區域別需計算小計時才使用此串列
Dim m_AgentTmpList() As AGENTITEM
Dim m_AgentTmpListCount As Integer

' 宣告地區項目的資料型態
Private Type COUNTRYITEM
   ' 地區代碼
   CountryCode As String
   CountryName As String
   Count As Integer
   'add by nick 2004/12/14
   Count08 As Integer
   NoAgentCount As Integer
   'add by nick 2004/12/14
   NoAgentCount08 As Integer
   AgentList() As AGENTITEM
   AgentCount As Integer
   'add by nick 2004/12/14
   AgentCount08 As Integer
End Type

Private Type ZONEITEM
   ' 地區別
   ZoneKind As String
   Count As Integer
   'add by nick 2004/12/14
   Count08 As Integer
   CountryList() As COUNTRYITEM
   CountryCount As Integer
   'add by nick 2004/12/14
   CountryCount08 As Integer
End Type
' 定義地區串列
Dim m_ZoneList() As ZONEITEM
Dim m_ZoneCount As Integer
'edit by nick 2004/12/14
'Dim m_DefaultPrinter As String

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
   Set frm030608 = Nothing
End Sub

' 清除所有佔用的空間
Private Sub Clear()
   Dim nX As Integer
   Dim nY As Integer
   If m_ZoneCount > 0 Then
      For nX = 0 To m_ZoneCount - 1
         If m_ZoneList(nX).CountryCount > 0 Then
            For nY = 0 To m_ZoneList(nX).CountryCount - 1
               If m_ZoneList(nX).CountryList(nY).AgentCount > 0 Then
                  Erase m_ZoneList(nX).CountryList(nY).AgentList
               End If
               m_ZoneList(nX).CountryList(nY).AgentCount = 0
            Next nY
            Erase m_ZoneList(nX).CountryList
            m_ZoneList(nX).CountryCount = 0
         End If
      Next nX
      Erase m_ZoneList
   End If
   m_ZoneCount = 0
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
'edit by nick 2004/12/14
'   Dim Prn As Printer
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
      ' 搜尋資料
      BuildField_RP
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/22 清除查詢印表記錄檔欄位
      If Len(textTMBM07_1) <> 0 Or Len(textTMBM07_2) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & Label1 & textTMBM07_1 & "-" & textTMBM07_2 'Add By Sindy 2010/10/22
      End If
      If Len(textNA02) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & Label2 & textNA02 & Label3  'Add By Sindy 2010/10/22
      End If
      If Len(textNA03) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & Label4 & textNA03 & Label5  'Add By Sindy 2010/10/22
      End If
      Select Case textNA02
         Case "A", "a":
            'edit by nick 204/12/14
            'If GetDBData_RP(1) = False Then: GoTo EXITSUB
            If GetDBData_RP_931214(1) = False Then: GoTo EXITSUB
            'edit by nick 204/12/14
            'Generate_RP 1
            Generate_RP_931214 1
            ' 測試
            'Generate_RP_SCREEN 1
         Case "B", "b":
            'edit by nick 204/12/14
            'If GetDBData_RP(2) = False Then: GoTo EXITSUB
            If GetDBData_RP_931214(2) = False Then: GoTo EXITSUB
            'edit by nick 204/12/14
            'Generate_RP 2
            Generate_RP_931214 2
            ' 測試
            'Generate_RP_SCREEN 2
         Case "C", "c":
            'edit by nick 204/12/14
            'If GetDBData_RP(3) = False Then: GoTo EXITSUB
            If GetDBData_RP_931214(3) = False Then: GoTo EXITSUB
            'edit by nick 204/12/14
            'Generate_RP 3
            Generate_RP_931214 3
            ' 測試
            'Generate_RP_SCREEN 3
         Case " ", "":
            'edit by nick 204/12/14
            'If GetDBData_RP(4) = False Then: GoTo EXITSUB
            If GetDBData_RP_931214(4) = False Then: GoTo EXITSUB
            'edit by nick 204/12/14
            'Generate_RP 4
            Generate_RP_931214 4
            '若未有輸台灣區別
             If Me.textNA03.Text = "" Then
                'edit by nick 204/12/14
                'Generate_RPAll
                Generate_RPAll_931214
            End If
            ' 測試
            'Generate_RP_SCREEN 4
      End Select
      InsertQueryLog ("") 'Add By Sindy 2010/10/22
      
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      Clear
      
      strTit = "輸出報表"
      strMsg = "列印結束"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
   
EXITSUB:
   Screen.MousePointer = vbDefault
   Clear
End Sub

' 取得地區(國籍)的數量總計
Private Function GetCountryAmount(ByRef CountryInfo As COUNTRYITEM) As Integer
   Dim nX As Integer
   Dim nAmount As Variant
   
   nAmount = 0
   For nX = 0 To CountryInfo.AgentCount - 1
      nAmount = nAmount + CountryInfo.AgentList(nX).Count
   Next nX
   GetCountryAmount = nAmount
End Function

' 取得區域別的數量總計
Private Function GetZoneAmount(ByRef ZoneInfo As ZONEITEM) As Integer
   Dim nX As Integer
   Dim nY As Integer
   Dim nAmount As Variant
   
   nAmount = 0
   For nX = 0 To ZoneInfo.CountryCount - 1
      For nY = 0 To ZoneInfo.CountryList(nX).AgentCount - 1
         nAmount = nAmount + ZoneInfo.CountryList(nX).AgentList(nY).Count
      Next nY
   Next nX
   GetZoneAmount = nAmount
End Function

Private Function GetNoAgentAmount(ByRef CountryInfo As COUNTRYITEM) As Integer
   Dim nX As Integer
   Dim nY As Integer
   Dim nAmount As Variant
   
   nAmount = 0
   For nX = 0 To CountryInfo.AgentCount - 1
      If IsEmptyText(CountryInfo.AgentList(nX).AgentName) = True Then
         nAmount = nAmount + CountryInfo.AgentList(nX).Count
      End If
   Next nX
   GetNoAgentAmount = nAmount
End Function

' 取得暫存區的總數
Private Function GetAgentTmpListAmount() As Variant
   Dim nAmount As Variant
   Dim nIndex As Integer
   nAmount = 0
   For nIndex = 0 To m_AgentTmpListCount - 1
      nAmount = nAmount + m_AgentTmpList(nIndex).Count
   Next nIndex
   GetAgentTmpListAmount = nAmount
End Function

' 取得暫存區中無代理人的總數
Private Function GetAgentTmpListNoAgentCount() As Integer
   Dim nAmount As Variant
   Dim nIndex As Integer
   nAmount = 0
   For nIndex = 0 To m_AgentTmpListCount - 1
      If IsEmptyText(m_AgentTmpList(nIndex).AgentName) = True Then
         nAmount = nAmount + m_AgentTmpList(nIndex).Count
      End If
   Next nIndex
End Function

' 取得國內,大陸或國外的無代理人總數
' Input : nCountry = 0 ==> 國內
'         nCountry = 1 ==> 大陸
'         nCountry = 2 ==> 國外
'         nCountry = 3 ==> 大陸及國外
Private Function GetNoAgentAmountByCountry(ByVal nCountry As Integer) As Long
   Dim nAmount As Long
   Dim nX As Long
   Dim nY As Long
   nAmount = 0
   For nX = 0 To m_ZoneCount - 1
      Select Case nCountry
         Case 0:
            If Mid(m_ZoneList(nX).ZoneKind, 1, 1) <> "A" Then
               GoTo NextRecord
            End If
         Case 1:
            If Mid(m_ZoneList(nX).ZoneKind, 1, 1) <> "B" Then
               GoTo NextRecord
            End If
         Case 2:
            If Mid(m_ZoneList(nX).ZoneKind, 1, 1) <> "C" Then
               GoTo NextRecord
            End If
         Case 3:
            If Mid(m_ZoneList(nX).ZoneKind, 1, 1) = "A" Then
               GoTo NextRecord
            End If
         Case Else
      End Select
      For nY = 0 To m_ZoneList(nX).CountryCount - 1
         nAmount = nAmount + m_ZoneList(nX).CountryList(nY).NoAgentCount
      Next nY
NextRecord:
   Next nX
   GetNoAgentAmountByCountry = nAmount
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得國內,大陸或國外的總數
' Input : nCountry = 0 ==> 國內
'         nCountry = 1 ==> 大陸
'         nCountry = 2 ==> 國外
'         nCountry = 3 ==> 大陸及國外
Private Function GetAgentTotalAmountByCountry(ByVal nCountry As Integer) As Long
   Dim nAmount As Long
   Dim nX As Long
   Dim nY As Long
   nAmount = 0
   For nX = 0 To m_ZoneCount - 1
      Select Case nCountry
         Case 0:
            If Mid(m_ZoneList(nX).ZoneKind, 1, 1) <> "A" Then
               GoTo NextRecord
            End If
         Case 1:
            If Mid(m_ZoneList(nX).ZoneKind, 1, 1) <> "B" Then
               GoTo NextRecord
            End If
         Case 2:
            If Mid(m_ZoneList(nX).ZoneKind, 1, 1) <> "C" Then
               GoTo NextRecord
            End If
         Case 3:
            If Mid(m_ZoneList(nX).ZoneKind, 1, 1) = "A" Then
               GoTo NextRecord
            End If
      End Select
      For nY = 0 To m_ZoneList(nX).CountryCount - 1
         nAmount = nAmount + m_ZoneList(nX).CountryList(nY).Count - m_ZoneList(nX).CountryList(nY).NoAgentCount
      Next nY
NextRecord:
   Next nX
   GetAgentTotalAmountByCountry = nAmount
End Function

' 清除計算區域別代理人數量的串列
Private Sub ClearAgentTmpList()
   ' 清除計算區域別代理人數量的串列
   If m_AgentTmpListCount > 0 Then
      Erase m_AgentTmpList
   End If
   m_AgentTmpListCount = 0
End Sub

'' 從資料庫中取得所有的資料
'Private Function GetDBData_RP(ByVal nReport As Integer) As Boolean
'   Dim rsMain As New ADODB.Recordset
'   Dim strSql As String
'   Dim strSubSQL As String
'   Dim strZoneKind, strZoneName, strZoneCode, strAgentName, strAgentCode, strAgentCompany As String
'   Dim bFindZone, bFindCountry, bFindAgent As Boolean
'   Dim nSortX, nSortY As Integer
'   Dim AgentTemp As AGENTITEM
'   Dim CountryTemp As COUNTRYITEM
'   Dim ZoneTemp As ZONEITEM
'   Dim bFromSec As Boolean
'   Dim bToSec As Boolean
'   Dim nX, ny, nZ As Integer
'   Dim c1X, c1Y, c2X, c2Y As String
'
'   GetDBData_RP = True
'
'   strSubSQL = Empty
'   Select Case textNA02
'      Case "a", "A":
'         strSubSQL = Empty
'      Case "b", "B":
'         strSubSQL = "NA02 LIKE '" & "B%" & "' "
'      Case "c", "C":
'         strSubSQL = "NA02 LIKE '" & "C%" & "' "
'      Case Else:
'         strSubSQL = Empty
'   End Select
'
'   ' 產生SQL查詢語法
'   bFromSec = Not IsEmptyText(textTMBM07_1.Text)
'   bToSec = Not IsEmptyText(textTMBM07_2.Text)
'   If bFromSec = True And bToSec = True Then
'      strSql = "SELECT DISTINCT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,NA01,NA02,TA01,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04 FROM TMBULLETIN, NATION, TAGENT " & _
'               "WHERE TMBM05 = NA03(+) AND " & _
'                     "TMBM06 = TA03(+) AND " & _
'                     "'T' = TA01(+) AND " & _
'                     "TMBM07 >= '" & textTMBM07_1 & "' AND " & _
'                     "TMBM07 <= '" & textTMBM07_2 & "' "
'      If strSubSQL <> Empty Then
'         strSql = strSql & " " & "AND " & strSubSQL
'      End If
'   ElseIf bFromSec = True And bToSec = False Then
'      strSql = "SELECT DISTINCT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,NA01,NA02,TA01,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04 FROM TMBULLETIN, NATION, TAGENT " & _
'               "WHERE TMBM05 = NA03(+) AND " & _
'                     "TMBM06 = TA03(+) AND " & _
'                     "'T' = TA01(+) AND " & _
'                     "TMBM07 >= '" & textTMBM07_1 & "' "
'      If strSubSQL <> Empty Then
'         strSql = strSql & " " & "AND " & strSubSQL
'      End If
'   ElseIf bFromSec = False And bToSec = True Then
'      strSql = "SELECT DISTINCT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,NA01,NA02,TA01,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04 FROM TMBULLETIN, NATION, TAGENT " & _
'               "WHERE TMBM05 = NA03(+) AND " & _
'                     "TMBM06 = TA03(+) AND " & _
'                     "'T' = TA01(+) AND " & _
'                     "TMBM07 <= '" & textTMBM07_2 & "' "
'      If strSubSQL <> Empty Then
'         strSql = strSql & " " & "AND " & strSubSQL
'      End If
'   Else
'      strSql = "SELECT DISTINCT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,NA01,NA02,TA01,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04 FROM TMBULLETIN, NATION, TAGENT " & _
'               "WHERE TMBM05 = NA03(+) AND " & _
'                     "TMBM06 = TA03(+) AND " & _
'                     "'T' = TA01(+) "
'      If strSubSQL <> Empty Then
'         strSql = strSql & " " & "AND " & strSubSQL
'      End If
'   End If
'
'   ' 取得資料庫的資料
'   rsMain.CursorLocation = adUseClient
'   rsMain.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   ' 無資料則離開
'   If rsMain.RecordCount <= 0 Then
'      GetDBData_RP = False
'      GoTo EXITSUB
'   End If
'
'   ' 設定初始值
'   m_ZoneCount = 0
'
'   rsMain.MoveFirst
'   ' 依序從資料記錄中取出欄位的內容
'   While Not rsMain.EOF
'      strAgentName = Empty
'      If IsNull(rsMain.Fields("TMBM06")) = False Then
'         strAgentName = rsMain.Fields("TMBM06")
'      End If
'      strAgentCode = Empty
'      If IsNull(rsMain.Fields("TA02")) = False Then
'         strAgentCode = rsMain.Fields("TA02")
'      End If
'      ' 代理人事務所名稱
'      strAgentCompany = Empty
'      If IsNull(rsMain.Fields("TA04")) = False Then
'         strAgentCompany = rsMain.Fields("TA04")
'      End If
'
'      ' 地區名稱
'      strZoneName = Empty
'      If IsNull(rsMain.Fields("TMBM05")) = False Then
'         strZoneName = rsMain.Fields("TMBM05")
'      End If
'      ' 地區別
'      strZoneKind = Empty
'      If IsNull(rsMain.Fields("NA02")) = False Then
'         strZoneKind = rsMain.Fields("NA02")
'      End If
'      ' 地區代碼
'      strZoneCode = Empty
'      If IsNull(rsMain.Fields("NA01")) = False Then
'         strZoneCode = rsMain.Fields("NA01")
'      End If
'
'      ' 依區域別將國內及國外及大陸分類
'      Select Case textNA02
'         ' 國內
'         Case "a", "A":
'            Select Case Mid(strZoneKind, 1, 1)
'               Case "A":
'               Case "B":
'                  strZoneKind = "B00"
'                  strZoneCode = "998"
'                  strZoneName = "大陸地區"
'               Case "C":
'                  strZoneKind = "C00"
'                  strZoneCode = "999"
'                  strZoneName = "國外地區"
'               Case Else:
'                  strZoneKind = "C00"
'                  strZoneCode = "999"
'                  strZoneName = "國外地區"
'            End Select
'         ' 大陸
'         Case "b", "B":
'            Select Case Mid(strZoneKind, 1, 1)
'               Case "A":
'                  strZoneKind = "A00"
'                  strZoneCode = "997"
'                  strZoneName = "國內地區"
'               Case "B":
'               Case "C":
'                  strZoneKind = "C00"
'                  strZoneCode = "999"
'                  strZoneName = "國外地區"
'               Case Else:
'                  strZoneKind = "C00"
'                  strZoneCode = "999"
'                  strZoneName = "國外地區"
'            End Select
'         ' 國外
'         Case "c", "C":
'            Select Case Mid(strZoneKind, 1, 1)
'               Case "A":
'                  strZoneKind = "A00"
'                  strZoneCode = "997"
'                  strZoneName = "國內地區"
'               Case "B":
'                  strZoneKind = "B00"
'                  strZoneCode = "998"
'                  strZoneName = "大陸地區"
'               Case "C":
'               Case Else:
'                  strZoneKind = "C99"
'                  strZoneCode = "996"
'                  strZoneName = "國外地區"
'            End Select
'      End Select
'
'      ' 地區串列
'      bFindZone = False
'      For nX = 0 To m_ZoneCount - 1
'         ' 找到地區別的結構
'         If m_ZoneList(nX).ZoneKind = strZoneKind Then
'            bFindZone = True
'
'            bFindCountry = False
'            For ny = 0 To m_ZoneList(nX).CountryCount - 1
'               ' 找到地區別結構中的地區(國家)列表
'               If m_ZoneList(nX).CountryList(ny).CountryCode = strZoneCode Then
'                  bFindCountry = True
'                  ' 計數加一
'                  m_ZoneList(nX).CountryList(ny).Count = m_ZoneList(nX).CountryList(ny).Count + 1
'                  ' 搜尋代理人串列
'                  bFindAgent = False
'                  'If strAgentCode <> Empty Then
'                  If IsEmptyText(strAgentName) = False Then
'                     For nZ = 0 To m_ZoneList(nX).CountryList(ny).AgentCount - 1
'                        'If m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCode = strAgentCode Then
'                        'Modify By Sindy 2010/02/26
'                        'If m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentName = strAgentName Then
'                        If m_ZoneList(nX).CountryList(ny).AgentList(nZ).AgentCompany = strAgentCompany Then
'                        '2010/02/26 End
'                           bFindAgent = True
'                           m_ZoneList(nX).CountryList(ny).AgentList(nZ).Count = m_ZoneList(nX).CountryList(ny).AgentList(nZ).Count + 1
'                           Exit For
'                        End If
'                     Next nZ
'                     ' 找不到此代理人的資料則新建一個代理人的結構
'                     If bFindAgent = False Then
'                        nZ = m_ZoneList(nX).CountryList(ny).AgentCount
'                        ReDim Preserve m_ZoneList(nX).CountryList(ny).AgentList(nZ + 1)
'                        m_ZoneList(nX).CountryList(ny).AgentList(nZ).AgentCode = strAgentCode
'                        'm_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentName = strAgentCompany
'                        m_ZoneList(nX).CountryList(ny).AgentList(nZ).AgentName = strAgentName
'                        m_ZoneList(nX).CountryList(ny).AgentList(nZ).AgentCompany = strAgentCompany
'                        m_ZoneList(nX).CountryList(ny).AgentCount = m_ZoneList(nX).CountryList(ny).AgentCount + 1
'                        m_ZoneList(nX).CountryList(ny).AgentList(nZ).Count = m_ZoneList(nX).CountryList(ny).AgentList(nZ).Count + 1
'                     End If
'                  Else
'                     m_ZoneList(nX).CountryList(ny).NoAgentCount = m_ZoneList(nX).CountryList(ny).NoAgentCount + 1
'                  End If
'               End If
'            Next ny
'            ' 找不到地區則新增地區
'            If bFindCountry = False Then
'               ny = m_ZoneList(nX).CountryCount
'               ReDim Preserve m_ZoneList(nX).CountryList(ny + 1)
'               m_ZoneList(nX).CountryList(ny).CountryCode = strZoneCode
'               m_ZoneList(nX).CountryList(ny).CountryName = strZoneName
'               m_ZoneList(nX).CountryList(ny).Count = 1
'               m_ZoneList(nX).CountryList(ny).AgentCount = 0
'               m_ZoneList(nX).CountryList(ny).NoAgentCount = 0
'               m_ZoneList(nX).CountryCount = m_ZoneList(nX).CountryCount + 1
'               'If strAgentCode <> Empty Then
'               If IsEmptyText(strAgentName) = False Then
'                  nZ = 0
'                  ReDim Preserve m_ZoneList(nX).CountryList(ny).AgentList(nZ + 1)
'                  m_ZoneList(nX).CountryList(ny).AgentList(nZ).AgentCode = strAgentCode
'                  m_ZoneList(nX).CountryList(ny).AgentList(nZ).AgentName = strAgentName
'                  m_ZoneList(nX).CountryList(ny).AgentList(nZ).AgentCompany = strAgentCompany
'                  m_ZoneList(nX).CountryList(ny).AgentList(nZ).Count = m_ZoneList(nX).CountryList(ny).AgentList(nZ).Count + 1
'                  m_ZoneList(nX).CountryList(ny).AgentCount = 1
'               Else
'                  m_ZoneList(nX).CountryList(ny).NoAgentCount = 1
'               End If
'            End If
'            Exit For
'         End If
'      Next nX
'
'      ' 找不到地區別則新增地區別結構
'      If bFindZone = False Then
'         nX = m_ZoneCount
'         ReDim Preserve m_ZoneList(nX + 1)
'         m_ZoneList(nX).ZoneKind = strZoneKind
'         m_ZoneList(nX).Count = 1
'         m_ZoneList(nX).CountryCount = 0
'         m_ZoneCount = m_ZoneCount + 1
'         ny = m_ZoneList(nX).CountryCount
'         ReDim Preserve m_ZoneList(nX).CountryList(ny + 1)
'         m_ZoneList(nX).CountryCount = 1
'         m_ZoneList(nX).CountryList(ny).CountryCode = strZoneCode
'         m_ZoneList(nX).CountryList(ny).CountryName = strZoneName
'         m_ZoneList(nX).CountryList(ny).Count = 1
'         m_ZoneList(nX).CountryList(ny).AgentCount = 0
'         m_ZoneList(nX).CountryList(ny).NoAgentCount = 0
'         'If strAgentCode <> Empty Then
'         If IsEmptyText(strAgentName) = False Then
'            nZ = 0
'            ReDim Preserve m_ZoneList(nX).CountryList(ny).AgentList(nZ + 1)
'            m_ZoneList(nX).CountryList(ny).AgentList(nZ).AgentCode = strAgentCode
'            m_ZoneList(nX).CountryList(ny).AgentList(nZ).AgentName = strAgentName
'            m_ZoneList(nX).CountryList(ny).AgentList(nZ).AgentCompany = strAgentCompany
'            m_ZoneList(nX).CountryList(ny).AgentList(nZ).Count = m_ZoneList(nX).CountryList(ny).AgentList(nZ).Count + 1
'            m_ZoneList(nX).CountryList(ny).AgentCount = 1
'         Else
'            m_ZoneList(nX).CountryList(ny).NoAgentCount = 1
'         End If
'      End If
'
'      ' 移到下一筆記錄
'      rsMain.MoveNext
'   Wend
'
'   ' 對地區別串列依地區別代碼小到大排序
'   For nSortX = 0 To m_ZoneCount - 1
'      For nSortY = nSortX To m_ZoneCount - 1
'         If m_ZoneList(nSortX).ZoneKind > m_ZoneList(nSortY).ZoneKind Then
'            ZoneTemp = m_ZoneList(nSortX)
'            m_ZoneList(nSortX) = m_ZoneList(nSortY)
'            m_ZoneList(nSortY) = ZoneTemp
'         End If
'      Next nSortY
'   Next nSortX
'   ' 對地區別中的地區(國籍)串列依國籍的代碼小到大排序
'   For nX = 0 To m_ZoneCount - 1
'      For nSortX = 0 To m_ZoneList(nX).CountryCount - 1
'         For nSortY = nSortX To m_ZoneList(nX).CountryCount - 1
'            'If m_ZoneList(nX).CountryList(nSortX).Count < m_ZoneList(nX).CountryList(nSortY).Count Then
'            '   CountryTemp = m_ZoneList(nX).CountryList(nSortX)
'            '   m_ZoneList(nX).CountryList(nSortX) = m_ZoneList(nX).CountryList(nSortY)
'            '   m_ZoneList(nX).CountryList(nSortY) = CountryTemp
'            'ElseIf m_ZoneList(nX).CountryList(nSortX).Count = m_ZoneList(nX).CountryList(nSortY).Count Then
'            '   If m_ZoneList(nX).CountryList(nSortX).CountryCode > m_ZoneList(nX).CountryList(nSortY).CountryCode Then
'            '      CountryTemp = m_ZoneList(nX).CountryList(nSortX)
'            '      m_ZoneList(nX).CountryList(nSortX) = m_ZoneList(nX).CountryList(nSortY)
'            '      m_ZoneList(nX).CountryList(nSortY) = CountryTemp
'            '   End If
'            'End If
'            If m_ZoneList(nX).CountryList(nSortX).CountryCode > m_ZoneList(nX).CountryList(nSortY).CountryCode Then
'               CountryTemp = m_ZoneList(nX).CountryList(nSortX)
'               m_ZoneList(nX).CountryList(nSortX) = m_ZoneList(nX).CountryList(nSortY)
'               m_ZoneList(nX).CountryList(nSortY) = CountryTemp
'            End If
'         Next nSortY
'      Next nSortX
'   Next nX
'   ' 對地區別中的地區(國籍)串列項目中的代理人串列依數量的多寡由大到小排序
'   For nX = 0 To m_ZoneCount - 1
'      For ny = 0 To m_ZoneList(nX).CountryCount - 1
'         For nSortX = 0 To m_ZoneList(nX).CountryList(ny).AgentCount - 1
'            For nSortY = nSortX To m_ZoneList(nX).CountryList(ny).AgentCount - 1
'               If m_ZoneList(nX).CountryList(ny).AgentList(nSortX).Count < m_ZoneList(nX).CountryList(ny).AgentList(nSortY).Count Then
'                  AgentTemp = m_ZoneList(nX).CountryList(ny).AgentList(nSortX)
'                  m_ZoneList(nX).CountryList(ny).AgentList(nSortX) = m_ZoneList(nX).CountryList(ny).AgentList(nSortY)
'                  m_ZoneList(nX).CountryList(ny).AgentList(nSortY) = AgentTemp
'               End If
'            Next nSortY
'         Next nSortX
'      Next ny
'   Next nX
'
'EXITSUB:
'   rsMain.Close
'   Set rsMain = Nothing
'End Function

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
      m_Field(nIndex).Left = m_LeftMargin + (nIndex * nFieldWidth) + 9
      Select Case nIndex
         Case 0:
            m_Field(nIndex).Left = m_LeftMargin + (nIndex * nFieldWidth)
            m_Field(nIndex).Name = "地區別"
         Case 1:
            m_Field(nIndex).Name = "排名"
            m_Field(nIndex).Left = m_LeftMargin + (nIndex * nFieldWidth) + 6
         Case 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14:
            m_Field(nIndex).Name = CStr(nIndex - 1)
         Case 15:
            m_Field(nIndex).Name = "無代理人 "
            'add by nick 2004/12/14
            m_Field(nIndex).Width = 8
         Case 16:
            m_Field(nIndex).Left = m_LeftMargin + (nIndex * nFieldWidth) + 11
            'edit by nick 2004/12/22
            'm_Field(nIndex).Name = "總數"
            m_Field(nIndex).Name = "代理人總數"
            'add by nick 2004/12/22
            m_Field(nIndex).Width = 10
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
   Dim nY As Long
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
   Printer.Print "表三：各區市場佔有統計表"
   
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

'
Private Function GetNoAgentAmountByZone(ByRef ZoneInfo As ZONEITEM) As Long
   Dim nX As Integer
   Dim nAmount As Long
   nAmount = 0
   For nX = 0 To ZoneInfo.CountryCount - 1
      nAmount = nAmount + ZoneInfo.CountryList(nX).NoAgentCount
   Next nX
   GetNoAgentAmountByZone = nAmount
End Function

Private Function GetTotalAmountByZone(ByRef ZoneInfo As ZONEITEM) As Long
   Dim nX As Integer
   Dim nAmount As Long
   nAmount = 0
   For nX = 0 To ZoneInfo.CountryCount - 1
      nAmount = nAmount + ZoneInfo.CountryList(nX).Count - ZoneInfo.CountryList(nX).NoAgentCount
   Next nX
   GetTotalAmountByZone = nAmount
End Function

' 將區域別的資料依代理人的數量多寡組成暫存串列並排序
Private Sub BuildAgentTmpList(ByRef ZoneInfo As ZONEITEM)
   Dim nX As Integer
   Dim nY As Integer
   Dim nZ As Integer
   Dim nIndex As Integer
   Dim bFindAgent As Boolean
   Dim AgentTmp As AGENTITEM
   ' 清除計算區域別代理人數量的串列
   'edit by nickc 2005/09/09
   'ClearAgentTmpList
      
   For nX = 0 To ZoneInfo.CountryCount - 1
      For nY = 0 To ZoneInfo.CountryList(nX).AgentCount - 1
         bFindAgent = False
         ' 搜尋原有的暫存串列
         For nZ = 0 To m_AgentTmpListCount - 1
            'If ZoneInfo.CountryList(nX).AgentList(nY).AgentCode = m_AgentTmpList(nZ).AgentCode Then
            'Modify By Sindy 2010/02/26
            'If ZoneInfo.CountryList(nX).AgentList(nY).AgentName = m_AgentTmpList(nZ).AgentName Then
            If ZoneInfo.CountryList(nX).AgentList(nY).AgentCompany = m_AgentTmpList(nZ).AgentCompany Then
            '2010/02/26 End
               bFindAgent = True
               m_AgentTmpList(nZ).Count = m_AgentTmpList(nZ).Count + ZoneInfo.CountryList(nX).AgentList(nY).Count
               Exit For
            End If
         Next nZ
         If bFindAgent = False Then
            nIndex = m_AgentTmpListCount
            ReDim Preserve m_AgentTmpList(nIndex + 1)
            m_AgentTmpList(nIndex).AgentCode = ZoneInfo.CountryList(nX).AgentList(nY).AgentCode
            m_AgentTmpList(nIndex).AgentName = ZoneInfo.CountryList(nX).AgentList(nY).AgentName
            m_AgentTmpList(nIndex).AgentCompany = ZoneInfo.CountryList(nX).AgentList(nY).AgentCompany
            m_AgentTmpList(nIndex).Count = ZoneInfo.CountryList(nX).AgentList(nY).Count
            m_AgentTmpListCount = m_AgentTmpListCount + 1
         End If
      Next nY
   Next nX
   ' 排序
   For nX = 0 To m_AgentTmpListCount - 1
      For nY = nX To m_AgentTmpListCount - 1
         If m_AgentTmpList(nX).Count < m_AgentTmpList(nY).Count Then
            AgentTmp = m_AgentTmpList(nX)
            m_AgentTmpList(nX) = m_AgentTmpList(nY)
            m_AgentTmpList(nY) = AgentTmp
         End If
      Next nY
   Next nX
End Sub

' 將台灣的資料組成一個Temp List
Private Sub BuildTaiwanTmpList()
   Dim nX As Integer
   Dim nY As Integer
   Dim nZ As Integer
   Dim nIndex As Integer
   Dim bFindAgent As Boolean
   Dim AgentTmp As AGENTITEM
   ' 清除計算區域別代理人數量的串列
   'edit by nickc 2005/09/09
   'ClearAgentTmpList
      
   For nX = 0 To m_ZoneCount - 1
      If Mid(m_ZoneList(nX).ZoneKind, 1, 1) = "A" Then
         For nY = 0 To m_ZoneList(nX).CountryCount - 1
            For nZ = 0 To m_ZoneList(nX).CountryList(nY).AgentCount - 1
               bFindAgent = False
               ' 搜尋原有的暫存串列
               For nIndex = 0 To m_AgentTmpListCount - 1
                  'If m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCode = m_AgentTmpList(nIndex).AgentCode Then
                  'Modify By Sindy 2010/02/26
                  'If m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentName = m_AgentTmpList(nIndex).AgentName Then
                  If m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCompany = m_AgentTmpList(nIndex).AgentCompany Then
                  '2010/02/26 End
                     bFindAgent = True
                     m_AgentTmpList(nIndex).Count = m_AgentTmpList(nIndex).Count + m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count
                     Exit For
                  End If
               Next nIndex
               If bFindAgent = False Then
                  nIndex = m_AgentTmpListCount
                  ReDim Preserve m_AgentTmpList(nIndex + 1)
                  m_AgentTmpList(nIndex).AgentCode = m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCode
                  m_AgentTmpList(nIndex).AgentName = m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentName
                  m_AgentTmpList(nIndex).AgentCompany = m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCompany
                  m_AgentTmpList(nIndex).Count = m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count
                  m_AgentTmpListCount = m_AgentTmpListCount + 1
               End If
            Next nZ
         Next nY
      End If
   Next nX
   ' 排序
   For nX = 0 To m_AgentTmpListCount - 1
      For nY = nX To m_AgentTmpListCount - 1
         If m_AgentTmpList(nX).Count < m_AgentTmpList(nY).Count Then
            AgentTmp = m_AgentTmpList(nX)
            m_AgentTmpList(nX) = m_AgentTmpList(nY)
            m_AgentTmpList(nY) = AgentTmp
         End If
      Next nY
   Next nX
End Sub

' 將大陸的資料組成一個Temp List
Private Sub BuildChinaTmpList()
   Dim nX As Integer
   Dim nY As Integer
   Dim nZ As Integer
   Dim nIndex As Integer
   Dim bFindAgent As Boolean
   Dim AgentTmp As AGENTITEM
   ' 清除計算區域別代理人數量的串列
   'edit by nickc 2005/09/09
   'ClearAgentTmpList
      
   For nX = 0 To m_ZoneCount - 1
      If Mid(m_ZoneList(nX).ZoneKind, 1, 1) = "B" Then
         For nY = 0 To m_ZoneList(nX).CountryCount - 1
            For nZ = 0 To m_ZoneList(nX).CountryList(nY).AgentCount - 1
               bFindAgent = False
               ' 搜尋原有的暫存串列
               For nIndex = 0 To m_AgentTmpListCount - 1
                  'Modify By Sindy 2010/4/23
                  'If m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentName = m_AgentTmpList(nIndex).AgentName Then
                  If m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCompany = m_AgentTmpList(nIndex).AgentCompany Then
                  '2010/4/23 End
                     bFindAgent = True
                     m_AgentTmpList(nIndex).Count = m_AgentTmpList(nIndex).Count + m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count
                     Exit For
                  End If
               Next nIndex
               If bFindAgent = False Then
                  nIndex = m_AgentTmpListCount
                  ReDim Preserve m_AgentTmpList(nIndex + 1)
                  m_AgentTmpList(nIndex).AgentCode = m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCode
                  m_AgentTmpList(nIndex).AgentName = m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentName
                  m_AgentTmpList(nIndex).AgentCompany = m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCompany
                  m_AgentTmpList(nIndex).Count = m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count
                  m_AgentTmpListCount = m_AgentTmpListCount + 1
               End If
            Next nZ
         Next nY
      End If
   Next nX
   ' 排序
   For nX = 0 To m_AgentTmpListCount - 1
      For nY = nX To m_AgentTmpListCount - 1
         If m_AgentTmpList(nX).Count < m_AgentTmpList(nY).Count Then
            AgentTmp = m_AgentTmpList(nX)
            m_AgentTmpList(nX) = m_AgentTmpList(nY)
            m_AgentTmpList(nY) = AgentTmp
         End If
      Next nY
   Next nX
End Sub

' 將國外的資料組成一個Temp List
Private Sub BuildForeignTmpList()
   Dim nX As Integer
   Dim nY As Integer
   Dim nZ As Integer
   Dim nIndex As Integer
   Dim bFindAgent As Boolean
   Dim AgentTmp As AGENTITEM
   ' 清除計算區域別代理人數量的串列
   'edit by nickc 2005/09/09
   'ClearAgentTmpList
      
   For nX = 0 To m_ZoneCount - 1
      If Mid(m_ZoneList(nX).ZoneKind, 1, 1) = "C" Then
         For nY = 0 To m_ZoneList(nX).CountryCount - 1
            For nZ = 0 To m_ZoneList(nX).CountryList(nY).AgentCount - 1
               bFindAgent = False
               ' 搜尋原有的暫存串列
               For nIndex = 0 To m_AgentTmpListCount - 1
                  'Modify By Sindy 2010/4/23
                  'If m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentName = m_AgentTmpList(nIndex).AgentName Then
                  If m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCompany = m_AgentTmpList(nIndex).AgentCompany Then
                  '2010/4/23 End
                     bFindAgent = True
                     m_AgentTmpList(nIndex).Count = m_AgentTmpList(nIndex).Count + m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count
                     Exit For
                  End If
               Next nIndex
               If bFindAgent = False Then
                  nIndex = m_AgentTmpListCount
                  ReDim Preserve m_AgentTmpList(nIndex + 1)
                  m_AgentTmpList(nIndex).AgentCode = m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCode
                  m_AgentTmpList(nIndex).AgentName = m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentName
                  m_AgentTmpList(nIndex).AgentCompany = m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCompany
                  m_AgentTmpList(nIndex).Count = m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count
                  m_AgentTmpListCount = m_AgentTmpListCount + 1
               End If
            Next nZ
         Next nY
      End If
   Next nX
   ' 排序
   For nX = 0 To m_AgentTmpListCount - 1
      For nY = nX To m_AgentTmpListCount - 1
         If m_AgentTmpList(nX).Count < m_AgentTmpList(nY).Count Then
            AgentTmp = m_AgentTmpList(nX)
            m_AgentTmpList(nX) = m_AgentTmpList(nY)
            m_AgentTmpList(nY) = AgentTmp
         End If
      Next nY
   Next nX
End Sub

' 將國外的資料組成一個Temp List
Private Sub BuildTmpList(ByVal strKey As String)
   Dim nX As Integer
   Dim nY As Integer
   Dim nZ As Integer
   Dim nIndex As Integer
   Dim bFindAgent As Boolean
   Dim AgentTmp As AGENTITEM
   Dim nLen As Integer
   ' 清除計算區域別代理人數量的串列
   'edit by nickc 2005/09/09
   'ClearAgentTmpList
   nLen = Len(strKey)
   For nX = 0 To m_ZoneCount - 1
      If Mid(m_ZoneList(nX).ZoneKind, 1, nLen) = strKey Then
         For nY = 0 To m_ZoneList(nX).CountryCount - 1
            For nZ = 0 To m_ZoneList(nX).CountryList(nY).AgentCount - 1
               bFindAgent = False
               ' 搜尋原有的暫存串列
               For nIndex = 0 To m_AgentTmpListCount - 1
                  'Modify By Sindy 2010/4/23
                  'If m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentName = m_AgentTmpList(nIndex).AgentName Then
                  If m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCompany = m_AgentTmpList(nIndex).AgentCompany Then
                  '2010/4/23 End
                     bFindAgent = True
                     m_AgentTmpList(nIndex).Count = m_AgentTmpList(nIndex).Count + m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count
                     Exit For
                  End If
               Next nIndex
               If bFindAgent = False Then
                  nIndex = m_AgentTmpListCount
                  ReDim Preserve m_AgentTmpList(nIndex + 1)
                  m_AgentTmpList(nIndex).AgentCode = m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCode
                  m_AgentTmpList(nIndex).AgentName = m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentName
                  m_AgentTmpList(nIndex).AgentCompany = m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCompany
                  m_AgentTmpList(nIndex).Count = m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count
                  m_AgentTmpListCount = m_AgentTmpListCount + 1
               End If
            Next nZ
         Next nY
      End If
   Next nX
   ' 排序
   For nX = 0 To m_AgentTmpListCount - 1
      For nY = nX To m_AgentTmpListCount - 1
         If m_AgentTmpList(nX).Count < m_AgentTmpList(nY).Count Then
            AgentTmp = m_AgentTmpList(nX)
            m_AgentTmpList(nX) = m_AgentTmpList(nY)
            m_AgentTmpList(nY) = AgentTmp
         End If
      Next nY
   Next nX
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' 列印合計
'' nCountry ==> 0 : 表國內
''              1 : 表大陸
''              2 : 表國外
''              3 : 表大陸及國外
'' nRow ==> 印到第幾列
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Modify By Cheng 2003/01/27
''多加參數--頁數
''Private Sub Generate_GrandTotal(ByVal nCountry As Integer, ByRef nRow As Integer)
'Private Sub Generate_GrandTotal(ByVal nCountry As Integer, ByRef nRow As Integer, ByRef nPage As Integer)
'   Dim fld(16) As String
'   Dim nAmount As Variant
'   Dim nNoAgentAmount As Variant
'   Dim nTotalAmount As Variant
'   Dim fValue As Double
'   Dim nX As Long
'   Dim ny As Long
'   Dim nZ As Long
'   Dim nCenter As Long
'   Dim nRight As Long
'
'    'Add By Cheng 2003/01/27
'    ' 若列數超過頁面的高度限制時則換頁
'    If nRow > m_ReportDataRows Then
'       Printer.NewPage
'       nPage = nPage + 1
'       PrintPageHeader_RP nPage
'       nRow = 0
'    End If
'   ' 輸出事務所
'   For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'   nRow = nRow + 1
'   Select Case nCountry
'      Case 0:
'         'edit by nickc 2005/09/09
'         'BuildTaiwanTmpList
'         fld(0) = "國內合計"
'         nNoAgentAmount = GetNoAgentAmountByCountry(0)
'         nTotalAmount = GetAgentTotalAmountByCountry(0)
'      Case 1:
'         'edit by nickc 2005/09/09
'         'BuildChinaTmpList
'         fld(0) = "大陸合計"
'         nNoAgentAmount = GetNoAgentAmountByCountry(1)
'         nTotalAmount = GetAgentTotalAmountByCountry(1)
'      Case 2:
'         'edit by nickc 2005/09/09
'         'BuildForeignTmpList
'         fld(0) = "國外合計"
'         nNoAgentAmount = GetNoAgentAmountByCountry(2)
'         nTotalAmount = GetAgentTotalAmountByCountry(2)
'      Case 3:
'         'edit by nickc 2005/09/09
'         'BuildForeignTmpList
'         fld(0) = "國外合計"
'         nNoAgentAmount = GetNoAgentAmountByCountry(2)
'         nTotalAmount = GetAgentTotalAmountByCountry(2)
'   End Select
'   fld(1) = "事務所"
'   For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
'   fld(nZ + 2) = m_AgentTmpList(nZ).AgentCompany
'   Next nZ
'   fld(15) = Empty
'   fld(16) = Empty
'   ' 輸出代理人列
'   For nZ = 0 To 16
'      Select Case nZ
'         Case 0, 1
'            Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
'            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'            Printer.Print fld(nZ)
'         Case Else
'            Printer.FontSize = 8
'            nCenter = ((m_Field(nZ).Left * m_CharWidth) + (m_Field(nZ).Left + m_Field(nZ).Width) * m_CharWidth) / 2
'            Printer.CurrentX = nCenter - Printer.TextWidth(fld(nZ)) / 2
'            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'            Printer.Print fld(nZ)
'            Printer.FontSize = 12
'      End Select
'   Next nZ
'
'   ' 輸出數量
'   For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'   nRow = nRow + 1
'   fld(0) = Empty
'   fld(1) = "數量"
'   For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
'      fld(nZ + 2) = m_AgentTmpList(nZ).Count
'   Next nZ
'   fld(15) = nNoAgentAmount
'   fld(16) = nTotalAmount
'   ' 輸出數量列
'   For nZ = 0 To 16
'      Select Case nZ
'         Case 0, 1
'            Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
'            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'            Printer.Print fld(nZ)
'         Case Else
'            nRight = (m_Field(nZ).Left + m_Field(nZ).Width - 2) * m_CharWidth
'            Printer.CurrentX = nRight - Printer.TextWidth(fld(nZ))
'            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'            Printer.Print fld(nZ)
'      End Select
'   Next nZ
'
'   ' 輸出百分比
'   For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'   nRow = nRow + 1
'   fld(0) = Empty
'   fld(1) = "百分比"
'   For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
'      nAmount = m_AgentTmpList(nZ).Count
'      fValue = nAmount / nTotalAmount * 100
'      fld(nZ + 2) = Format(fValue, "##0.00") & " %"
'   Next nZ
'   'fValue = nNoAgentAmount / nTotalAmount * 100
'   'fld(15) = Format(fValue, "##0.00") & " %"
'   fld(15) = Empty
'   fld(16) = Empty
'   ' 輸出百分比列
'   For nZ = 0 To 16
'      Select Case nZ
'         Case 0, 1
'            Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
'            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'            Printer.Print fld(nZ)
'         Case Else
'            nRight = (m_Field(nZ).Left + m_Field(nZ).Width - 2) * m_CharWidth
'            Printer.CurrentX = nRight - Printer.TextWidth(fld(nZ))
'            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'            Printer.Print fld(nZ)
'      End Select
'   Next nZ
'   ' 列印分隔線
'   nRow = nRow + 1
'   PrintTerminateLine m_HeaderHeight + nRow
'   ' 清除暫存串列
'   ClearAgentTmpList
'End Sub

'' 列印表一的內容
'Public Sub Generate_RP(ByVal nReport As Integer)
'   Dim nRow As Integer
'   Dim nPage As Integer
'   Dim fld(17) As String
'   Dim nType As Integer
'   Dim nAmount As Variant
'   Dim nNoAgentAmount As Variant
'   Dim nTotalAmount As Variant
'   Dim nZoneCount As Integer
'   Dim fValue As Double
'   Dim nX As Integer
'   Dim ny As Integer
'   Dim nZ As Integer
'   Dim nCenter As Long
'   Dim nRight As Long
'   Dim ZonePrev As String
'   Dim ZoneCurr As String
'   Dim bChangePage As Boolean
'   Dim bPrintCF As Boolean
'
'   ' 紙張大小
'   Select Case m_PaperSize
'      Case "A4":
'         Printer.PaperSize = vbPRPSA4
'         Printer.Orientation = vbPRORLandscape
'      Case "REPORT":
'         Printer.PaperSize = vbPRPSFanfoldUS
'      Case Else:
'         Printer.PaperSize = vbPRPSA4
'         Printer.Orientation = vbPRORLandscape
'   End Select
'   ' 紙張方向
'
'   ' 當全部列印時是否已列印大陸及國外的合計資料
'   bPrintCF = False
'
'   ' 印表頭
'   nPage = 1
'   PrintPageHeader_RP nPage
'
'   ZonePrev = Empty
'   nRow = 0
'   ' 依地區別
'   For nZoneCount = 0 To m_ZoneCount - 1
'      ' 地區代碼
'      ZoneCurr = m_ZoneList(nZoneCount).ZoneKind
'      ' 地區別不同時需列印合計
'      'If ZonePrev <> Empty And Mid(ZoneCurr, 1, 1) <> Mid(ZonePrev, 1, 1) Then
'      '   Select Case Mid(ZonePrev, 1, 1)
'      '      'Case "A": Generate_GrandTotal 0, nRow
'      '      Case "B": Generate_GrandTotal 1, nRow
'      '      Case "C": Generate_GrandTotal 2, nRow
'      '   End Select
'      'End If
'
'        If Me.textNA03.Text = "" Or Left(Me.textNA03.Text, 2) = "A1" Then
'          ' 當列印全部報表時, 檢查是否已列印大陸及國外的合計資料
'          If (nReport = 1 Or nReport = 4) And bPrintCF = False Then
'             ' 當表四印台灣北區資料結束後(即其它區列印前)需列印大陸及國外的合計資料
'             If Mid(ZoneCurr, 1, 1) >= "A" And Mid(ZoneCurr, 2, 1) > "1" Then
'                ' 若列數超過頁面的高度限制時則換頁
'                If nRow > m_ReportDataRows Then
'                   Printer.NewPage
'                   nPage = nPage + 1
'                   PrintPageHeader_RP nPage
'                   nRow = 0
'                End If
'                'Modify By Cheng 2003/01/27
'                '多傳入參數--頁數
'    '            Generate_GrandTotal 0, nRow
'                Generate_GrandTotal 0, nRow, nPage
'                ' 若列數超過頁面的高度限制時則換頁
'                If nRow > m_ReportDataRows Then
'                   Printer.NewPage
'                   nPage = nPage + 1
'                   PrintPageHeader_RP nPage
'                   nRow = 0
'                End If
'                'Modify By Cheng 2003/01/27
'                '多傳入參數--頁數
'    '            Generate_GrandTotal 1, nRow
'                Generate_GrandTotal 1, nRow, nPage
'                ' 若列數超過頁面的高度限制時則換頁
'                If nRow > m_ReportDataRows Then
'                   Printer.NewPage
'                   nPage = nPage + 1
'                   PrintPageHeader_RP nPage
'                   nRow = 0
'                End If
'                'Modify By Cheng 2003/01/27
'                '多傳入參數--頁數
'    '            Generate_GrandTotal 2, nRow
'                Generate_GrandTotal 2, nRow, nPage
'
'                bPrintCF = True
'             End If
'          End If
'        End If
'
'      ' 表一只列印國內區域的資料
'      If nReport = 1 Then
'         If Mid(m_ZoneList(nZoneCount).ZoneKind, 1, 1) <> "A" Then
'            Exit For
'         End If
'      End If
'        '若有輸台灣區別
'        If Me.textNA03.Text <> "" Then
'            Select Case Me.textNA03.Text
'            Case "1"
'                If Left(ZoneCurr, 2) <> "A1" Then GoTo NextRec
'            Case "2"
'                If Left(ZoneCurr, 2) <> "A2" Then GoTo NextRec
'            Case "3"
'                If Left(ZoneCurr, 2) <> "A3" Then GoTo NextRec
'            Case "4"
'                If Left(ZoneCurr, 2) <> "A4" Then GoTo NextRec
'            Case "5"
'                If Left(ZoneCurr, 2) <> "A5" Then GoTo NextRec
'            End Select
'        End If
'
'      ' 檢查是否換頁的旗標
'      bChangePage = False
'      If ZonePrev <> Empty Then
'         If Mid(ZoneCurr, 1, 1) <> Mid(ZonePrev, 1, 1) Then
'            If Mid(ZonePrev, 1, 1) = "A" Then
'               bChangePage = True
'            End If
'         Else
'            If Mid(ZoneCurr, 1, 1) = "A" Then
'               ' 第二碼不同時需換頁
'               If Mid(ZoneCurr, 2, 1) <> Mid(ZonePrev, 2, 1) Then
'                  bChangePage = True
'               End If
'            End If
'         End If
'      End If
'      ' 換頁
'      If bChangePage Then
'         Printer.NewPage
'         nPage = nPage + 1
'         PrintPageHeader_RP nPage
'         nRow = 0
'      End If
'
'      ' 地區別中的國籍串列
'      For ny = 0 To m_ZoneList(nZoneCount).CountryCount - 1
'         ' 若列數超過頁面的高度限制時則換頁
'         If nRow > m_ReportDataRows Then
'            Printer.NewPage
'            nPage = nPage + 1
'            PrintPageHeader_RP nPage
'            nRow = 0
'         End If
'
'         ' 清除欄位
'         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'         ' 代理人
'         nRow = nRow + 1
'         fld(0) = m_ZoneList(nZoneCount).CountryList(ny).CountryName
'         fld(1) = "事務所"
'         For nZ = 0 To Min(13, m_ZoneList(nZoneCount).CountryList(ny).AgentCount - 1)
'            fld(nZ + 2) = m_ZoneList(nZoneCount).CountryList(ny).AgentList(nZ).AgentCompany
'         Next nZ
'         fld(15) = Empty
'         fld(16) = Empty
'         ' 輸出代理人列
'         For nZ = 0 To 16
'            Select Case nZ
'               Case 0, 1
'                  Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'               Case Else
'                  Printer.FontSize = 8
'                  nCenter = ((m_Field(nZ).Left * m_CharWidth) + (m_Field(nZ).Left + m_Field(nZ).Width) * m_CharWidth) / 2
'                  Printer.CurrentX = nCenter - Printer.TextWidth(fld(nZ)) / 2
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'                  Printer.FontSize = 12
'            End Select
'         Next nZ
'         ' 清除欄位
'         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'         ' 數量
'         nRow = nRow + 1
'         fld(1) = "數量"
'         For nZ = 0 To Min(13, m_ZoneList(nZoneCount).CountryList(ny).AgentCount - 1)
'            fld(nZ + 2) = m_ZoneList(nZoneCount).CountryList(ny).AgentList(nZ).Count
'         Next nZ
'         nNoAgentAmount = m_ZoneList(nZoneCount).CountryList(ny).NoAgentCount
'         fld(15) = nNoAgentAmount
'         nTotalAmount = m_ZoneList(nZoneCount).CountryList(ny).Count - nNoAgentAmount
'         fld(16) = nTotalAmount
'         ' 輸出數量列
'         For nZ = 0 To 16
'            Select Case nZ
'               Case 0, 1
'                  Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'               Case Else
'                  nRight = (m_Field(nZ).Left + m_Field(nZ).Width - 2) * m_CharWidth
'                  Printer.CurrentX = nRight - Printer.TextWidth(fld(nZ))
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'            End Select
'         Next nZ
'         ' 清除欄位
'         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'         ' 百分比列
'         nRow = nRow + 1
'         fld(1) = "百分比"
'         For nZ = 0 To Min(13, m_ZoneList(nZoneCount).CountryList(ny).AgentCount - 1)
'            nAmount = m_ZoneList(nZoneCount).CountryList(ny).AgentList(nZ).Count
'            'nTotalAmount = GetCountryAmount(m_ZoneList(nZoneCount).CountryList(nY))
'            fValue = nAmount / nTotalAmount * 100
'            fld(nZ + 2) = Format(fValue, "##0.00") & " %"
'         Next nZ
'         'fValue = nNoAgentAmount / nTotalAmount * 100
'         'fld(15) = Format(fValue, "##0.00") & " %"
'         fld(15) = Empty
'         fld(16) = Empty
'         ' 輸出百分比列
'         For nZ = 0 To 16
'            Select Case nZ
'               Case 0, 1
'                  Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'               Case Else
'                  nRight = (m_Field(nZ).Left + m_Field(nZ).Width - 2) * m_CharWidth
'                  Printer.CurrentX = nRight - Printer.TextWidth(fld(nZ))
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'            End Select
'         Next nZ
'         ' 列印分隔線
'         If ny < m_ZoneList(nZoneCount).CountryCount - 1 Then
'            nRow = nRow + 1
'            PrintSplitLine m_HeaderHeight + nRow
'         End If
'      Next ny
'
'      ' 若該區域中的地區多於一個地區則需列印小計
'      'Modify By Sindy 2011/1/25
'      'If m_ZoneList(nZoneCount).CountryCount > 1 Then
'      If m_ZoneList(nZoneCount).CountryCount >= 1 Then
'         ' 列印分隔線
'         nRow = nRow + 1
'         PrintSplitLine m_HeaderHeight + nRow
'
'         ' 若列數超過頁面的高度限制時則換頁
'         If nRow > m_ReportDataRows Then
'            Printer.NewPage
'            nPage = nPage + 1
'            PrintPageHeader_RP nPage
'            nRow = 0
'         End If
'
'         ' 計算區域別的小計
'         BuildAgentTmpList m_ZoneList(nZoneCount)
'         ' 清除欄位
'         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'         ' 代理人
'         nRow = nRow + 1
'         fld(0) = "小計"
'         fld(1) = "事務所"
'         For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
'            fld(nZ + 2) = m_AgentTmpList(nZ).AgentCompany
'         Next nZ
'         fld(15) = Empty
'         fld(16) = Empty
'         ' 輸出代理人列
'         For nZ = 0 To 16
'            Select Case nZ
'               Case 0, 1
'                  Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'               Case Else
'                  Printer.FontSize = 8
'                  nCenter = ((m_Field(nZ).Left * m_CharWidth) + (m_Field(nZ).Left + m_Field(nZ).Width) * m_CharWidth) / 2
'                  Printer.CurrentX = nCenter - Printer.TextWidth(fld(nZ)) / 2
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'                  Printer.FontSize = 12
'            End Select
'         Next nZ
'         ' 清除欄位
'         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'         ' 數量
'         nRow = nRow + 1
'         fld(1) = "數量"
'         For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
'            fld(nZ + 2) = m_AgentTmpList(nZ).Count
'         Next nZ
'         nNoAgentAmount = GetNoAgentAmountByZone(m_ZoneList(nZoneCount))
'         fld(15) = nNoAgentAmount
'         nTotalAmount = GetTotalAmountByZone(m_ZoneList(nZoneCount))
'         fld(16) = nTotalAmount
'         ' 輸出數量列
'         For nZ = 0 To 16
'            Select Case nZ
'               Case 0, 1
'                  Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'               Case Else
'                  nRight = (m_Field(nZ).Left + m_Field(nZ).Width - 2) * m_CharWidth
'                  Printer.CurrentX = nRight - Printer.TextWidth(fld(nZ))
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'            End Select
'         Next nZ
'         ' 清除欄位
'         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'         ' 百分比列
'         nRow = nRow + 1
'         fld(1) = "百分比"
'         For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
'            nAmount = m_AgentTmpList(nZ).Count
'            nTotalAmount = GetZoneAmount(m_ZoneList(nZoneCount))
'            fValue = nAmount / nTotalAmount * 100
'            fld(nZ + 2) = Format(fValue, "##0.00") & " %"
'         Next nZ
'         'fValue = nNoAgentAmount / nTotalAmount * 100
'         'fld(15) = Format(fValue, "##0.00") & " %"
'         fld(15) = Empty
'         fld(16) = Empty
'         ' 輸出百分比列
'         For nZ = 0 To 16
'            Select Case nZ
'               Case 0, 1
'                  Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'               Case Else
'                  nRight = (m_Field(nZ).Left + m_Field(nZ).Width - 2) * m_CharWidth
'                  Printer.CurrentX = nRight - Printer.TextWidth(fld(nZ))
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'            End Select
'         Next nZ
'         ' 清除暫存串列
'         ClearAgentTmpList
'      End If
'
'      ZonePrev = ZoneCurr
'
'      ' 列印區域別的分隔線
'      nRow = nRow + 1
'      PrintTerminateLine m_HeaderHeight + nRow
'
'NextRec:
'   Next nZoneCount
'
'   Select Case nReport
'      Case 2:
'         ' 若列數超過頁面的高度限制時則換頁
'         If nRow > m_ReportDataRows Then
'            Printer.NewPage
'            nPage = nPage + 1
'            PrintPageHeader_RP nPage
'            nRow = 0
'         End If
'        'Modify By Cheng 2003/01/27
'        '多傳入參數--頁數
''         Generate_GrandTotal 1, nRow
'         Generate_GrandTotal 1, nRow, nPage
'      Case 3:
'         ' 若列數超過頁面的高度限制時則換頁
'         If nRow > m_ReportDataRows Then
'            Printer.NewPage
'            nPage = nPage + 1
'            PrintPageHeader_RP nPage
'            nRow = 0
'         End If
'        'Modify By Cheng 2003/01/27
'        '多傳入參數--頁數
''         Generate_GrandTotal 2, nRow
'         Generate_GrandTotal 2, nRow, nPage
'      Case 4:
'         ' 若列數超過頁面的高度限制時則換頁
'         If nRow > m_ReportDataRows Then
'            Printer.NewPage
'            nPage = nPage + 1
'            PrintPageHeader_RP nPage
'            nRow = 0
'         End If
'        'Modify By Cheng 2003/01/27
'        '多傳入參數--頁數
''         Generate_GrandTotal 3, nRow
'         Generate_GrandTotal 3, nRow, nPage
'   End Select
'
'   Printer.EndDoc
'
'End Sub

'' 列印表一的內容
'Public Sub Generate_RPAll()
'   Dim nRow As Integer
'   Dim nPage As Integer
'   Dim fld(17) As String
'   Dim nType As Integer
'   Dim nAmount As Variant
'   Dim nTotalAmount As Variant
'   Dim nNoAgentAmount As Variant
'   Dim nZoneCount As Integer
'   Dim fValue As Variant
'   Dim nX As Integer
'   Dim ny As Integer
'   Dim nZ As Integer
'   Dim nCenter As Long
'   Dim nRight As Long
'   Dim ZonePrev As String
'   Dim ZoneCurr As String
'   Dim bPrintTotal As Boolean
'
'   ' 紙張大小
'   Select Case m_PaperSize
'      Case "A4":
'         Printer.PaperSize = vbPRPSA4
'         Printer.Orientation = vbPRORLandscape
'      Case "REPORT":
'         Printer.PaperSize = vbPRPSFanfoldUS
'      Case Else:
'         Printer.PaperSize = vbPRPSA4
'         Printer.Orientation = vbPRORLandscape
'   End Select
'   ' 紙張方向
'
'   ' 印表頭
'   nPage = 1
'   PrintPageHeader_RP nPage
'
'   ZonePrev = Empty
'   nRow = 0
'   ' 依地區別
'   For nZoneCount = 0 To m_ZoneCount - 1
'      ' 地區代碼
'      ZoneCurr = m_ZoneList(nZoneCount).ZoneKind
'      ' 地區別中的國籍串列
'      For ny = 0 To m_ZoneList(nZoneCount).CountryCount - 1
'         ' 若列數超過頁面的高度限制時則換頁
'         If nRow > m_ReportDataRows Then
'            Printer.NewPage
'            nPage = nPage + 1
'            PrintPageHeader_RP nPage
'            nRow = 0
'         End If
'
'         ' 清除欄位
'         For nZ = 0 To 13: fld(nZ) = Empty: Next nZ
'         ' 代理人
'         nRow = nRow + 1
'         fld(0) = m_ZoneList(nZoneCount).CountryList(ny).CountryName
'         fld(1) = "事務所"
'         For nZ = 0 To Min(13, m_ZoneList(nZoneCount).CountryList(ny).AgentCount - 1)
'            fld(nZ + 2) = m_ZoneList(nZoneCount).CountryList(ny).AgentList(nZ).AgentCompany
'         Next nZ
'         fld(15) = Empty
'         fld(16) = Empty
'         ' 輸出代理人列
'         For nZ = 0 To 16
'            Select Case nZ
'               Case 0, 1
'                  Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'               Case Else
'                  Printer.FontSize = 8
'                  nCenter = ((m_Field(nZ).Left * m_CharWidth) + (m_Field(nZ).Left + m_Field(nZ).Width) * m_CharWidth) / 2
'                  Printer.CurrentX = nCenter - Printer.TextWidth(fld(nZ)) / 2
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'                  Printer.FontSize = 12
'            End Select
'         Next nZ
'         ' 清除欄位
'         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'         ' 數量
'         nRow = nRow + 1
'         fld(1) = "數量"
'         For nZ = 0 To Min(13, m_ZoneList(nZoneCount).CountryList(ny).AgentCount - 1)
'            fld(nZ + 2) = m_ZoneList(nZoneCount).CountryList(ny).AgentList(nZ).Count
'         Next nZ
'         nNoAgentAmount = m_ZoneList(nZoneCount).CountryList(ny).NoAgentCount
'         nTotalAmount = m_ZoneList(nZoneCount).CountryList(ny).Count - nNoAgentAmount
'         fld(15) = nNoAgentAmount
'         fld(16) = nTotalAmount
'         ' 輸出數量列
'         For nZ = 0 To 16
'            Select Case nZ
'               Case 0, 1
'                  Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'               Case Else
'                  nRight = (m_Field(nZ).Left + m_Field(nZ).Width - 2) * m_CharWidth
'                  Printer.CurrentX = nRight - Printer.TextWidth(fld(nZ))
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'            End Select
'         Next nZ
'         ' 清除欄位
'         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'         ' 百分比列
'         nRow = nRow + 1
'         fld(1) = "百分比"
'         For nZ = 0 To Min(13, m_ZoneList(nZoneCount).CountryList(ny).AgentCount - 1)
'            nAmount = m_ZoneList(nZoneCount).CountryList(ny).AgentList(nZ).Count
'            fValue = nAmount / nTotalAmount * 100
'            fld(nZ + 2) = Format(fValue, "##0.00") & " %"
'         Next nZ
'         fld(15) = Empty
'         fld(16) = Empty
'         ' 輸出百分比列
'         For nZ = 0 To 16
'            Select Case nZ
'               Case 0, 1
'                  Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'               Case Else
'                  nRight = (m_Field(nZ).Left + m_Field(nZ).Width - 2) * m_CharWidth
'                  Printer.CurrentX = nRight - Printer.TextWidth(fld(nZ))
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'            End Select
'         Next nZ
'         ' 列印分隔線
'         If ny < m_ZoneList(nZoneCount).CountryCount - 1 Then
'            nRow = nRow + 1
'            PrintSplitLine m_HeaderHeight + nRow
'         End If
'      Next ny
'
'      ' 若該區域中的地區多於一個地區則需列印小計
'      'Modify By Sindy 2011/1/25
'      'If m_ZoneList(nZoneCount).CountryCount > 1 Then
'      If m_ZoneList(nZoneCount).CountryCount >= 1 Then
'         ' 列印分隔線
'         nRow = nRow + 1
'         PrintSplitLine m_HeaderHeight + nRow
'
'         ' 計算區域別的小計
'         BuildAgentTmpList m_ZoneList(nZoneCount)
'         ' 清除欄位
'         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'         ' 代理人
'         nRow = nRow + 1
'         fld(0) = "小計"
'         fld(1) = "事務所"
'         For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
'            fld(nZ + 2) = m_AgentTmpList(nZ).AgentCompany
'         Next nZ
'         fld(15) = Empty
'         fld(16) = Empty
'         ' 輸出代理人列
'         For nZ = 0 To 16
'            Select Case nZ
'               Case 0, 1
'                  Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'               Case Else
'                  Printer.FontSize = 8
'                  nCenter = ((m_Field(nZ).Left * m_CharWidth) + (m_Field(nZ).Left + m_Field(nZ).Width) * m_CharWidth) / 2
'                  Printer.CurrentX = nCenter - Printer.TextWidth(fld(nZ)) / 2
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'                  Printer.FontSize = 12
'            End Select
'         Next nZ
'         ' 清除欄位
'         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'         ' 數量
'         nRow = nRow + 1
'         fld(1) = "數量"
'         For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
'            fld(nZ + 2) = m_AgentTmpList(nZ).Count
'         Next nZ
'         nNoAgentAmount = GetNoAgentAmountByZone(m_ZoneList(nZoneCount))
'         nTotalAmount = GetTotalAmountByZone(m_ZoneList(nZoneCount))
'         fld(15) = nNoAgentAmount
'         fld(16) = nTotalAmount
'         ' 輸出數量列
'         For nZ = 0 To 16
'            Select Case nZ
'               Case 0, 1
'                  Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'               Case Else
'                  nRight = (m_Field(nZ).Left + m_Field(nZ).Width - 2) * m_CharWidth
'                  Printer.CurrentX = nRight - Printer.TextWidth(fld(nZ))
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'            End Select
'         Next nZ
'         ' 清除欄位
'         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'         ' 百分比列
'         nRow = nRow + 1
'         fld(1) = "百分比"
'         For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
'            nAmount = m_AgentTmpList(nZ).Count
'            nTotalAmount = GetZoneAmount(m_ZoneList(nZoneCount))
'            fValue = nAmount / nTotalAmount * 100
'            fld(nZ + 2) = Format(fValue, "##0.00") & " %"
'         Next nZ
'         fld(15) = Empty
'         fld(16) = Empty
'         ' 輸出百分比列
'         For nZ = 0 To 16
'            Select Case nZ
'               Case 0, 1
'                  Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'               Case Else
'                  nRight = (m_Field(nZ).Left + m_Field(nZ).Width - 2) * m_CharWidth
'                  Printer.CurrentX = nRight - Printer.TextWidth(fld(nZ))
'                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
'                  Printer.Print fld(nZ)
'            End Select
'         Next nZ
'         ' 清除暫存串列
'         ClearAgentTmpList
'      End If
'
'      ZonePrev = ZoneCurr
'
'      ' 列印區域別的分隔線
'      nRow = nRow + 1
'      PrintTerminateLine m_HeaderHeight + nRow
'
'      ' 若大區域不同時則需列印總計
'      bPrintTotal = False
'      If nZoneCount = m_ZoneCount - 1 Then
'         bPrintTotal = True
'      Else
'         If Mid(m_ZoneList(nZoneCount).ZoneKind, 1, 1) <> Mid(m_ZoneList(nZoneCount + 1).ZoneKind, 1, 1) Then
'            bPrintTotal = True
'         End If
'      End If
'      If bPrintTotal = True Then
'         Select Case Mid(m_ZoneList(nZoneCount).ZoneKind, 1, 1)
'            'Modify By Cheng 2003/01/27
'            '多傳入參數--頁數
''            Case "A": Generate_GrandTotal 0, nRow
''            Case "B": Generate_GrandTotal 1, nRow
''            Case "C": Generate_GrandTotal 2, nRow
'            Case "A": Generate_GrandTotal 0, nRow, nPage
'            Case "B": Generate_GrandTotal 1, nRow, nPage
'            Case "C": Generate_GrandTotal 2, nRow, nPage
'         End Select
'      End If
'   Next nZoneCount
'
'   Printer.EndDoc
'End Sub

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
   textNA02 = " "
   
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
'   If GetDBData_RP(4) = False Then: GoTo EXITSUB
'   Generate_RP 4
'   Generate_RPAll
   If GetDBData_RP_931214(4) = False Then: GoTo EXITSUB
   Generate_RP_931214 4 'Modify By Sindy 2022/3/23 桂英說表三印1份PDF(以第1份為準)
   'Generate_RPAll_931214
   ' 清除所佔用的空間
   Clear
EXITSUB:
   Set frm030608 = Nothing
End Sub

Private Sub textNA03_GotFocus()
    TextInverse Me.textNA03
End Sub

Private Sub textNA03_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 Then
        KeyAscii = 0
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

Private Sub textTMBM07_1_GotFocus()
   InverseTextBox textTMBM07_1
End Sub

Private Sub textTMBM07_2_GotFocus()
   InverseTextBox textTMBM07_2
End Sub

Private Sub textNA02_GotFocus()
   InverseTextBox textNA02
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
      textTMBM07_1.SetFocus
      GoTo EXITSUB
   End If
   If IsEmptyText(textTMBM07_2) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入公報卷期(迄)"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTMBM07_2.SetFocus
      GoTo EXITSUB
   End If
   If IsEmptyText(textTMBM07_1) = False And IsEmptyText(textTMBM07_2) = False Then
      If Val(textTMBM07_1) > Val(textTMBM07_2) Then
         strTit = "資料檢核"
         strMsg = "公報卷期範圍不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTMBM07_1.SetFocus
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

'' 列印表一的內容
'Public Sub Generate_RP_SCREEN(ByVal nReport As Integer)
'   Dim nRow As Integer
'   Dim nPage As Integer
'   Dim fld(17) As String
'   Dim nType As Integer
'   Dim nAmount As Variant
'   Dim nTotalAmount As Variant
'   Dim nZoneCount As Integer
'   Dim fValue As Double
'   Dim nX As Integer
'   Dim ny As Integer
'   Dim nZ As Integer
'   Dim nCenter As Long
'   Dim nRight As Long
'   Dim ZonePrev As String
'   Dim ZoneCurr As String
'   Dim bChangePage As Boolean
'   Dim bPrintCF As Boolean
'
'   ' 當全部列印時是否已列印大陸及國外的合計資料
'   bPrintCF = False
'
'   ZonePrev = Empty
'   nRow = 0
'   ' 依地區別
'   For nZoneCount = 0 To m_ZoneCount - 1
'      ' 地區代碼
'      ZoneCurr = m_ZoneList(nZoneCount).ZoneKind
'
'      ' 當列印全部報表時, 檢查是否已列印大陸及國外的合計資料
'      If (nReport = 1 Or nReport = 4) And bPrintCF = False Then
'         ' 當表四印台灣北區資料結束後(即其它區列印前)需列印大陸及國外的合計資料
'         If Mid(ZoneCurr, 1, 1) >= "A" And Mid(ZoneCurr, 2, 1) > "1" Then
'            ' 取得國內的暫存資料
'            BuildTaiwanTmpList
'            ' 清除欄位
'            For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'
'            ' 代理人
'            fld(0) = "國內"
'            fld(1) = "事務所"
'            'For nZ = 0 To Min(15, m_AgentTmpListCount - 1)
'            For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
'               fld(nZ + 2) = m_AgentTmpList(nZ).AgentCompany
'            Next nZ
'            fld(14) = "無代理人"
'            fld(15) = "總數"
'            'Debug.Print fld(0) & " " & fld(1) & " " & fld(2) & " " & fld(3) & " " & fld(4) & " " & fld(5) & " " & fld(6) & " " & fld(7) & " " & fld(8) & " " & fld(9) & " " & fld(10) & " " & fld(11) & " " & fld(12) & " " & fld(13) & " " & fld(14) & " " & fld(15)
'
'            ' 清除欄位
'            For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'            ' 數量
'            nRow = nRow + 1
'            fld(1) = "數量"
'            'For nZ = 0 To Min(15, m_AgentTmpListCount - 1)
'            For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
'               fld(nZ + 2) = m_AgentTmpList(nZ).Count
'            Next nZ
'            fld(14) = GetAgentTmpListNoAgentCount()
'            fld(15) = GetAgentTmpListAmount()
'            'Debug.Print fld(0) & " " & fld(1) & " " & fld(2) & " " & fld(3) & " " & fld(4) & " " & fld(5) & " " & fld(6) & " " & fld(7) & " " & fld(8) & " " & fld(9) & " " & fld(10) & " " & fld(11) & " " & fld(12) & " " & fld(13) & " " & fld(14) & " " & fld(15) & " " & fld(16)
'
'            ' 清除欄位
'            For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'            ' 百分比列
'            nRow = nRow + 1
'            fld(1) = "百分比"
'            For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
'               nAmount = m_AgentTmpList(nZ).Count
'               'nTotalAmount = GetZoneAmount(m_ZoneList(nZoneCount))
'               nTotalAmount = GetAgentTmpListAmount()
'               fValue = nAmount / nTotalAmount * 100
'               fld(nZ + 2) = Format(fValue, "##0.00") & " %"
'            Next nZ
'            'Debug.Print fld(0) & " " & fld(1) & " " & fld(2) & " " & fld(3) & " " & fld(4) & " " & fld(5) & " " & fld(6) & " " & fld(7) & " " & fld(8) & " " & fld(9) & " " & fld(10) & " " & fld(11) & " " & fld(12) & " " & fld(13) & " " & fld(14) & " " & fld(15) & " " & fld(16)            ' 列印分隔線
'
'            ' 清除暫存串列
'            'edit by nickc 2005/09/09
'            'ClearAgentTmpList
'
'            ' 取得大陸的暫存資料
'            'edit by nickc 2005/09/09
'            'BuildChinaTmpList
'            ' 清除欄位
'            For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'            ' 代理人
'            nRow = nRow + 1
'            fld(0) = "大陸"
'            fld(1) = "事務所"
'            For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
'               fld(nZ + 2) = m_AgentTmpList(nZ).AgentCompany
'            Next nZ
'            fld(14) = "無代理人"
'            fld(15) = "總數"
'            'Debug.Print fld(0) & " " & fld(1) & " " & fld(2) & " " & fld(3) & " " & fld(4) & " " & fld(5) & " " & fld(6) & " " & fld(7) & " " & fld(8) & " " & fld(9) & " " & fld(10) & " " & fld(11) & " " & fld(12) & " " & fld(13) & " " & fld(14) & " " & fld(15) & " " & fld(16)
'            ' 清除欄位
'            For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'            ' 數量
'            nRow = nRow + 1
'            fld(1) = "數量"
'            For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
'               fld(nZ + 2) = m_AgentTmpList(nZ).Count
'            Next nZ
'            fld(14) = GetAgentTmpListNoAgentCount()
'            fld(15) = GetAgentTmpListAmount()
'            'Debug.Print fld(0) & " " & fld(1) & " " & fld(2) & " " & fld(3) & " " & fld(4) & " " & fld(5) & " " & fld(6) & " " & fld(7) & " " & fld(8) & " " & fld(9) & " " & fld(10) & " " & fld(11) & " " & fld(12) & " " & fld(13) & " " & fld(14) & " " & fld(15) & " " & fld(16)
'            ' 清除欄位
'            For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'
'            ' 百分比列
'            fld(1) = "百分比"
'            For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
'               nAmount = m_AgentTmpList(nZ).Count
'               'nTotalAmount = GetZoneAmount(m_ZoneList(nZoneCount))
'               nTotalAmount = GetAgentTmpListAmount()
'               fValue = nAmount / nTotalAmount * 100
'               fld(nZ + 2) = Format(fValue, "##0.00") & " %"
'            Next nZ
'            'Debug.Print fld(0) & " " & fld(1) & " " & fld(2) & " " & fld(3) & " " & fld(4) & " " & fld(5) & " " & fld(6) & " " & fld(7) & " " & fld(8) & " " & fld(9) & " " & fld(10) & " " & fld(11) & " " & fld(12) & " " & fld(13) & " " & fld(14) & " " & fld(15) & " " & fld(16)
'
'            ' 清除暫存串列
'            'edit by nickc 2005/09/09
'            'ClearAgentTmpList
'
'            ' 取得國外的暫存資料
'            BuildForeignTmpList
'            ' 清除欄位
'            For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'            ' 代理人
'            nRow = nRow + 1
'            fld(0) = "國外"
'            fld(1) = "事務所"
'            For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
'               fld(nZ + 2) = m_AgentTmpList(nZ).AgentCompany
'            Next nZ
'            fld(14) = "無代理人"
'            fld(15) = "總數"
'            'debug.Print fld(0) & " " & fld(1) & " " & fld(2) & " " & fld(3) & " " & fld(4) & " " & fld(5) & " " & fld(6) & " " & fld(7) & " " & fld(8) & " " & fld(9) & " " & fld(10) & " " & fld(11) & " " & fld(12) & " " & fld(13) & " " & fld(14) & " " & fld(15) & " " & fld(16)
'
'            ' 清除欄位
'            For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'            ' 數量
'            nRow = nRow + 1
'            fld(1) = "數量"
'            For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
'               fld(nZ + 2) = m_AgentTmpList(nZ).Count
'            Next nZ
'            fld(14) = GetAgentTmpListNoAgentCount()
'            fld(15) = GetAgentTmpListAmount()
'            'debug.Print fld(0) & " " & fld(1) & " " & fld(2) & " " & fld(3) & " " & fld(4) & " " & fld(5) & " " & fld(6) & " " & fld(7) & " " & fld(8) & " " & fld(9) & " " & fld(10) & " " & fld(11) & " " & fld(12) & " " & fld(13) & " " & fld(14) & " " & fld(15) & " " & fld(16)
'
'            ' 清除欄位
'            For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'            ' 百分比列
'            nRow = nRow + 1
'            fld(1) = "百分比"
'            For nZ = 0 To Min(15, m_AgentTmpListCount - 1)
'               nAmount = m_AgentTmpList(nZ).Count
'               'nTotalAmount = GetZoneAmount(m_ZoneList(nZoneCount))
'               nTotalAmount = GetAgentTmpListAmount()
'               fValue = nAmount / nTotalAmount * 100
'               fld(nZ + 2) = Format(fValue, "##0.00") & " %"
'            Next nZ
'            'debug.Print fld(0) & " " & fld(1) & " " & fld(2) & " " & fld(3) & " " & fld(4) & " " & fld(5) & " " & fld(6) & " " & fld(7) & " " & fld(8) & " " & fld(9) & " " & fld(10) & " " & fld(11) & " " & fld(12) & " " & fld(13) & " " & fld(14) & " " & fld(15) & " " & fld(16)
'            ' 清除暫存串列
'            ClearAgentTmpList
'            ' 設定列印大陸及國外合計資料的旗標為 已列印過
'            bPrintCF = True
'         End If
'      End If
'
'      ' 表一只列印國內區域的資料
'      If nReport = 1 Then
'         If Mid(m_ZoneList(nZoneCount).ZoneKind, 1, 1) <> "A" Then
'            Exit For
'         End If
'      End If
'
'      ' 地區別中的國籍串列
'      For ny = 0 To m_ZoneList(nZoneCount).CountryCount - 1
'         ' 清除欄位
'         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'         ' 代理人
'         nRow = nRow + 1
'         fld(0) = m_ZoneList(nZoneCount).CountryList(ny).CountryName
'         fld(1) = "事務所"
'         For nZ = 0 To Min(13, m_ZoneList(nZoneCount).CountryList(ny).AgentCount - 1)
'            fld(nZ + 2) = m_ZoneList(nZoneCount).CountryList(ny).AgentList(nZ).AgentCompany
'         Next nZ
'         fld(14) = "無代理人"
'         fld(15) = "總數"
'         'debug.Print fld(0) & " " & fld(1) & " " & fld(2) & " " & fld(3) & " " & fld(4) & " " & fld(5) & " " & fld(6) & " " & fld(7) & " " & fld(8) & " " & fld(9) & " " & fld(10) & " " & fld(11) & " " & fld(12) & " " & fld(13) & " " & fld(14) & " " & fld(15) & " " & fld(16)
'         ' 清除欄位
'         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'
'         ' 數量
'         fld(1) = "數量"
'         For nZ = 0 To Min(13, m_ZoneList(nZoneCount).CountryList(ny).AgentCount - 1)
'            fld(nZ + 2) = m_ZoneList(nZoneCount).CountryList(ny).AgentList(nZ).Count
'         Next nZ
'         fld(14) = GetNoAgentAmount(m_ZoneList(nZoneCount).CountryList(ny))
'         fld(15) = GetCountryAmount(m_ZoneList(nZoneCount).CountryList(ny))
'         'debug.Print fld(0) & " " & fld(1) & " " & fld(2) & " " & fld(3) & " " & fld(4) & " " & fld(5) & " " & fld(6) & " " & fld(7) & " " & fld(8) & " " & fld(9) & " " & fld(10) & " " & fld(11) & " " & fld(12) & " " & fld(13) & " " & fld(14) & " " & fld(15) & " " & fld(16)
'         ' 清除欄位
'         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'
'         ' 百分比列
'         fld(1) = "百分比"
'         For nZ = 0 To Min(15, m_ZoneList(nZoneCount).CountryList(ny).AgentCount - 1)
'            nAmount = m_ZoneList(nZoneCount).CountryList(ny).AgentList(nZ).Count
'            nTotalAmount = GetCountryAmount(m_ZoneList(nZoneCount).CountryList(ny))
'            fValue = nAmount / nTotalAmount * 100
'            fld(nZ + 2) = Format(fValue, "##0.00") & " %"
'         Next nZ
'         'debug.Print fld(0) & " " & fld(1) & " " & fld(2) & " " & fld(3) & " " & fld(4) & " " & fld(5) & " " & fld(6) & " " & fld(7) & " " & fld(8) & " " & fld(9) & " " & fld(10) & " " & fld(11) & " " & fld(12) & " " & fld(13) & " " & fld(14) & " " & fld(15) & " " & fld(16)
'      Next ny
'
'      ' 若該區域中的地區多於一個地區則需列印小計
'      'Modify By Sindy 2011/1/25
'      'If m_ZoneList(nZoneCount).CountryCount > 1 Then
'      If m_ZoneList(nZoneCount).CountryCount >= 1 Then
'         ' 計算區域別的小計
'         BuildAgentTmpList m_ZoneList(nZoneCount)
'         ' 清除欄位
'         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'         ' 代理人
'         nRow = nRow + 1
'         fld(0) = "小計"
'         fld(1) = "事務所"
'         For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
'            fld(nZ + 2) = m_AgentTmpList(nZ).AgentCompany
'         Next nZ
'         fld(14) = "無代理人"
'         fld(15) = "總數"
'         'debug.Print fld(0) & " " & fld(1) & " " & fld(2) & " " & fld(3) & " " & fld(4) & " " & fld(5) & " " & fld(6) & " " & fld(7) & " " & fld(8) & " " & fld(9) & " " & fld(10) & " " & fld(11) & " " & fld(12) & " " & fld(13) & " " & fld(14) & " " & fld(15) & " " & fld(16)
'         ' 清除欄位
'         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'         ' 數量
'         fld(1) = "數量"
'         For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
'            fld(nZ + 2) = m_AgentTmpList(nZ).Count
'         Next nZ
'         fld(14) = GetAgentTmpListNoAgentCount()
'         fld(15) = GetAgentTmpListAmount()
'         'debug.Print fld(0) & " " & fld(1) & " " & fld(2) & " " & fld(3) & " " & fld(4) & " " & fld(5) & " " & fld(6) & " " & fld(7) & " " & fld(8) & " " & fld(9) & " " & fld(10) & " " & fld(11) & " " & fld(12) & " " & fld(13) & " " & fld(14) & " " & fld(15) & " " & fld(16)
'         ' 清除欄位
'         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
'         ' 百分比列
'         fld(1) = "百分比"
'         For nZ = 0 To Min(15, m_AgentTmpListCount - 1)
'            nAmount = m_AgentTmpList(nZ).Count
'            'nTotalAmount = GetZoneAmount(m_ZoneList(nZoneCount))
'            nTotalAmount = GetAgentTmpListAmount()
'            fValue = nAmount / nTotalAmount * 100
'            fld(nZ + 2) = Format(fValue, "##0.00") & " %"
'         Next nZ
'         'debug.Print fld(0) & " " & fld(1) & " " & fld(2) & " " & fld(3) & " " & fld(4) & " " & fld(5) & " " & fld(6) & " " & fld(7) & " " & fld(8) & " " & fld(9) & " " & fld(10) & " " & fld(11) & " " & fld(12) & " " & fld(13) & " " & fld(14) & " " & fld(15) & " " & fld(16)
'         ' 清除暫存串列
'         ClearAgentTmpList
'      End If
'
'      ZonePrev = ZoneCurr
'   Next nZoneCount
'   'debug.Print "CC"
'End Sub

' 列印表一的內容
Public Sub Generate_RP_931214(ByVal nReport As Integer)
   Dim nRow As Integer
   Dim nPage As Integer
   Dim fld(17) As String
   Dim nType As Integer
   Dim nAmount As Variant
   Dim nNoAgentAmount As Variant
   Dim nTotalAmount As Variant
   Dim nZoneCount As Integer
   Dim fValue As Double
   Dim nX As Integer
   Dim nY As Integer
   Dim nZ As Integer
   Dim nCenter As Long
   Dim nRight As Long
   Dim ZonePrev As String
   Dim ZoneCurr As String
   Dim bChangePage As Boolean
   Dim bPrintCF As Boolean
      
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
   
   ' 當全部列印時是否已列印大陸及國外的合計資料
   bPrintCF = False

   ' 印表頭
   nPage = 1
   PrintPageHeader_RP nPage
   ClearAgentTmpList
   
   ZonePrev = Empty
   nRow = 0
   ' 依地區別
   For nZoneCount = 0 To m_ZoneCount - 1
      ' 地區代碼
      ZoneCurr = m_ZoneList(nZoneCount).ZoneKind
      ' 地區別不同時需列印合計
      'If ZonePrev <> Empty And Mid(ZoneCurr, 1, 1) <> Mid(ZonePrev, 1, 1) Then
      '   Select Case Mid(ZonePrev, 1, 1)
      '      'Case "A": Generate_GrandTotal 0, nRow
      '      Case "B": Generate_GrandTotal 1, nRow
      '      Case "C": Generate_GrandTotal 2, nRow
      '   End Select
      'End If
        
        If Me.textNA03.Text = "" Or Left(Me.textNA03.Text, 2) = "A1" Then
          ' 當列印全部報表時, 檢查是否已列印大陸及國外的合計資料
          If (nReport = 1 Or nReport = 4) And bPrintCF = False Then
             ' 當表四印台灣北區資料結束後(即其它區列印前)需列印大陸及國外的合計資料
             If Mid(ZoneCurr, 1, 1) >= "A" And Mid(ZoneCurr, 2, 1) > "1" Then
                ' 若列數超過頁面的高度限制時則換頁
                If nRow > m_ReportDataRows Then
                   Printer.NewPage
                   nPage = nPage + 1
                   PrintPageHeader_RP nPage
                   nRow = 0
                End If
                'Modify By Cheng 2003/01/27
                '多傳入參數--頁數
    '            Generate_GrandTotal 0, nRow
                Generate_GrandTotal_931214 0, nRow, nPage
                ' 若列數超過頁面的高度限制時則換頁
                If nRow > m_ReportDataRows Then
                   Printer.NewPage
                   nPage = nPage + 1
                   PrintPageHeader_RP nPage
                   nRow = 0
                End If
                'Modify By Cheng 2003/01/27
                '多傳入參數--頁數
    '            Generate_GrandTotal 1, nRow
                Generate_GrandTotal_931214 1, nRow, nPage
                ' 若列數超過頁面的高度限制時則換頁
                If nRow > m_ReportDataRows Then
                   Printer.NewPage
                   nPage = nPage + 1
                   PrintPageHeader_RP nPage
                   nRow = 0
                End If
                'Modify By Cheng 2003/01/27
                '多傳入參數--頁數
    '            Generate_GrandTotal 2, nRow
                Generate_GrandTotal_931214 2, nRow, nPage
                
                bPrintCF = True
             End If
          End If
        End If

      ' 表一只列印國內區域的資料
      If nReport = 1 Then
         If Mid(m_ZoneList(nZoneCount).ZoneKind, 1, 1) <> "A" Then
            Exit For
         End If
      End If
        '若有輸台灣區別
        If Me.textNA03.Text <> "" Then
            Select Case Me.textNA03.Text
            Case "1"
                If Left(ZoneCurr, 2) <> "A1" Then GoTo NextRec
            Case "2"
                If Left(ZoneCurr, 2) <> "A2" Then GoTo NextRec
            Case "3"
                If Left(ZoneCurr, 2) <> "A3" Then GoTo NextRec
            Case "4"
                If Left(ZoneCurr, 2) <> "A4" Then GoTo NextRec
            Case "5"
                If Left(ZoneCurr, 2) <> "A5" Then GoTo NextRec
            End Select
        End If

      ' 檢查是否換頁的旗標
      bChangePage = False
      If ZonePrev <> Empty Then
         If Mid(ZoneCurr, 1, 1) <> Mid(ZonePrev, 1, 1) Then
            If Mid(ZonePrev, 1, 1) = "A" Then
               bChangePage = True
            End If
         Else
            If Mid(ZoneCurr, 1, 1) = "A" Then
               ' 第二碼不同時需換頁
               If Mid(ZoneCurr, 2, 1) <> Mid(ZonePrev, 2, 1) Then
                  bChangePage = True
               End If
            End If
         End If
      End If
      ' 換頁
      If bChangePage Then
         Printer.NewPage
         nPage = nPage + 1
         PrintPageHeader_RP nPage
         nRow = 0
      End If
         
      ' 地區別中的國籍串列
      For nY = 0 To m_ZoneList(nZoneCount).CountryCount - 1
         ' 若列數超過頁面的高度限制時則換頁
         If nRow > m_ReportDataRows Then
            Printer.NewPage
            nPage = nPage + 1
            PrintPageHeader_RP nPage
            nRow = 0
         End If

         ' 清除欄位
         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
         ' 代理人
         nRow = nRow + 1
         fld(0) = m_ZoneList(nZoneCount).CountryList(nY).CountryName
         fld(1) = "事務所"
         For nZ = 0 To Min(13, m_ZoneList(nZoneCount).CountryList(nY).AgentCount - 1)
            fld(nZ + 2) = m_ZoneList(nZoneCount).CountryList(nY).AgentList(nZ).AgentCompany
         Next nZ
         fld(15) = Empty
         fld(16) = Empty
         ' 輸出代理人列
         For nZ = 0 To 16
            Select Case nZ
               Case 0 '地區別
                  Printer.FontSize = 10
                  Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                  Printer.Print fld(nZ)
               Case 1 '排名
                  Printer.FontSize = 12
                  Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                  Printer.Print fld(nZ)
               Case Else '事務所
                  Printer.FontSize = 8
                  nCenter = ((m_Field(nZ).Left * m_CharWidth) + (m_Field(nZ).Left + m_Field(nZ).Width) * m_CharWidth) / 2
                  Printer.CurrentX = nCenter - Printer.TextWidth(fld(nZ)) / 2
                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                  'Modify By Sindy 2024/4/24 解決列印出來是?
                  'Printer.Print fld(nZ)
                  PUB_PrintUnicodeText fld(nZ), Printer.CurrentX, Printer.CurrentY, 0
                  '2024/4/24 END
                  Printer.FontSize = 12
            End Select
         Next nZ
         ' 清除欄位
         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
         ' 數量
         nRow = nRow + 1
         'Add By Sindy 2010/7/6
'         If nReport = 3 Then
            fld(0) = "　　(類)"
'         End If
         '2010/7/6 End
         fld(1) = "數量"
         For nZ = 0 To Min(13, m_ZoneList(nZoneCount).CountryList(nY).AgentCount - 1)
            fld(nZ + 2) = m_ZoneList(nZoneCount).CountryList(nY).AgentList(nZ).Count08
         Next nZ
         nNoAgentAmount = m_ZoneList(nZoneCount).CountryList(nY).NoAgentCount08
         nTotalAmount = m_ZoneList(nZoneCount).CountryList(nY).Count08 - nNoAgentAmount
         fld(15) = nNoAgentAmount
         fld(16) = nTotalAmount
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
         ' 百分比列
         nRow = nRow + 1
         fld(1) = "百分比"
         For nZ = 0 To Min(13, m_ZoneList(nZoneCount).CountryList(nY).AgentCount - 1)
            nAmount = m_ZoneList(nZoneCount).CountryList(nY).AgentList(nZ).Count08
            fValue = nAmount / nTotalAmount * 100
            fld(nZ + 2) = Format(fValue, "##0.00") & " %"
         Next nZ
         fld(15) = Empty
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
         'Add By Sindy 2010/7/6
         '列印區域為國外時才統計類及件的數量
'         If nReport = 3 Then
            ' 清除欄位
            For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
            ' 數量
            nRow = nRow + 1
            fld(0) = "　　(件)"
            fld(1) = "數量"
            For nZ = 0 To Min(13, m_ZoneList(nZoneCount).CountryList(nY).AgentCount - 1)
               fld(nZ + 2) = m_ZoneList(nZoneCount).CountryList(nY).AgentList(nZ).Count
            Next nZ
            nNoAgentAmount = m_ZoneList(nZoneCount).CountryList(nY).NoAgentCount
            nTotalAmount = m_ZoneList(nZoneCount).CountryList(nY).Count - nNoAgentAmount
            fld(15) = nNoAgentAmount
            fld(16) = nTotalAmount
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
            ' 百分比列
            nRow = nRow + 1
            fld(1) = "百分比"
            For nZ = 0 To Min(13, m_ZoneList(nZoneCount).CountryList(nY).AgentCount - 1)
               nAmount = m_ZoneList(nZoneCount).CountryList(nY).AgentList(nZ).Count
               fValue = nAmount / nTotalAmount * 100
               fld(nZ + 2) = Format(fValue, "##0.00") & " %"
            Next nZ
            fld(15) = Empty
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
'         End If
         '2010/7/6 End
         ' 列印分隔線
         If nY < m_ZoneList(nZoneCount).CountryCount - 1 Then
            nRow = nRow + 1
            PrintSplitLine m_HeaderHeight + nRow
         End If
      Next nY
      
      ' 若該區域中的地區多於一個地區則需列印小計
      'Modify By Sindy 2011/1/25
      'If m_ZoneList(nZoneCount).CountryCount > 1 Then
      If m_ZoneList(nZoneCount).CountryCount >= 1 Then
         ' 列印分隔線
         nRow = nRow + 1
         PrintSplitLine m_HeaderHeight + nRow
         
         ' 若列數超過頁面的高度限制時則換頁
         If nRow > m_ReportDataRows Then
            Printer.NewPage
            nPage = nPage + 1
            PrintPageHeader_RP nPage
            nRow = 0
         End If
         
         ' 計算區域別的小計
         BuildAgentTmpList08 m_ZoneList(nZoneCount)
         ' 清除欄位
         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
         ' 代理人
         nRow = nRow + 1
         fld(0) = "小計"
         fld(1) = "事務所"
         For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
            fld(nZ + 2) = m_AgentTmpList(nZ).AgentCompany
         Next nZ
         fld(15) = Empty
         fld(16) = Empty
         ' 輸出代理人列
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
                  'Modify By Sindy 2024/4/24 解決列印出來是?
                  'Printer.Print fld(nZ)
                  PUB_PrintUnicodeText fld(nZ), Printer.CurrentX, Printer.CurrentY, 0
                  '2024/4/24 END
                  Printer.FontSize = 12
            End Select
         Next nZ
         ' 清除欄位
         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
         ' 數量
         nRow = nRow + 1
         fld(0) = "　　(類)"
         fld(1) = "數量"
         For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
            fld(nZ + 2) = m_AgentTmpList(nZ).Count08
         Next nZ
         nNoAgentAmount = GetNoAgentAmountByZone08(m_ZoneList(nZoneCount))
         nTotalAmount = GetTotalAmountByZone08(m_ZoneList(nZoneCount))
         fld(15) = nNoAgentAmount
         fld(16) = nTotalAmount
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
         ' 百分比列
         nRow = nRow + 1
         fld(1) = "百分比"
         For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
            nAmount = m_AgentTmpList(nZ).Count08
            nTotalAmount = GetZoneAmount08(m_ZoneList(nZoneCount))
            fValue = nAmount / nTotalAmount * 100
            fld(nZ + 2) = Format(fValue, "##0.00") & " %"
         Next nZ
         'fValue = nNoAgentAmount / nTotalAmount * 100
         'fld(15) = Format(fValue, "##0.00") & " %"
         fld(15) = Empty
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
         ' 清除暫存串列
         'edit by nickc 2005/09/09
         'ClearAgentTmpList
         '****
         
         ' 若列數超過頁面的高度限制時則換頁
         If nRow > m_ReportDataRows Then
            Printer.NewPage
            nPage = nPage + 1
            PrintPageHeader_RP nPage
            nRow = 0
         End If
         
         ' 清除欄位
         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
         ' 數量
         nRow = nRow + 1
         fld(0) = "　　(件)"
         fld(1) = "數量"
         For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
            fld(nZ + 2) = m_AgentTmpList(nZ).Count
         Next nZ
         nNoAgentAmount = GetNoAgentAmountByZone(m_ZoneList(nZoneCount))
         nTotalAmount = GetTotalAmountByZone(m_ZoneList(nZoneCount))
         fld(15) = nNoAgentAmount
         fld(16) = nTotalAmount
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
         ' 百分比列
         nRow = nRow + 1
         fld(1) = "百分比"
         For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
            nAmount = m_AgentTmpList(nZ).Count
            nTotalAmount = GetZoneAmount(m_ZoneList(nZoneCount))
            fValue = nAmount / nTotalAmount * 100
            fld(nZ + 2) = Format(fValue, "##0.00") & " %"
         Next nZ
         'fValue = nNoAgentAmount / nTotalAmount * 100
         'fld(15) = Format(fValue, "##0.00") & " %"
         fld(15) = Empty
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
         ' 清除暫存串列
         ClearAgentTmpList
      End If
      
      ZonePrev = ZoneCurr
    
      ' 列印區域別的分隔線
      nRow = nRow + 1
      PrintTerminateLine m_HeaderHeight + nRow
   
NextRec:
   Next nZoneCount
   
   Select Case nReport
      Case 2:
         ' 若列數超過頁面的高度限制時則換頁
         If nRow > m_ReportDataRows Then
            Printer.NewPage
            nPage = nPage + 1
            PrintPageHeader_RP nPage
            nRow = 0
         End If
        'Modify By Cheng 2003/01/27
        '多傳入參數--頁數
'         Generate_GrandTotal 1, nRow
         Generate_GrandTotal_931214 1, nRow, nPage
      Case 3:
         ' 若列數超過頁面的高度限制時則換頁
         If nRow > m_ReportDataRows Then
            Printer.NewPage
            nPage = nPage + 1
            PrintPageHeader_RP nPage
            nRow = 0
         End If
        'Modify By Cheng 2003/01/27
        '多傳入參數--頁數
'         Generate_GrandTotal 2, nRow
         Generate_GrandTotal_931214 2, nRow, nPage
      Case 4:
         ' 若列數超過頁面的高度限制時則換頁
         If nRow > m_ReportDataRows Then
            Printer.NewPage
            nPage = nPage + 1
            PrintPageHeader_RP nPage
            nRow = 0
         End If
        'Modify By Cheng 2003/01/27
        '多傳入參數--頁數
'         Generate_GrandTotal 3, nRow
         Generate_GrandTotal_931214 3, nRow, nPage
   End Select
   
   Printer.EndDoc
End Sub

' 列印表一的內容
Public Sub Generate_RPAll_931214()
   Dim nRow As Integer
   Dim nPage As Integer
   Dim fld(17) As String
   Dim nType As Integer
   Dim nAmount As Variant
   Dim nTotalAmount As Variant
   Dim nNoAgentAmount As Variant
   Dim nZoneCount As Integer
   Dim fValue As Variant
   Dim nX As Integer
   Dim nY As Integer
   Dim nZ As Integer
   Dim nCenter As Long
   Dim nRight As Long
   Dim ZonePrev As String
   Dim ZoneCurr As String
   Dim bPrintTotal As Boolean
      
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
   'add by nickc 2005/09/09
   ClearAgentTmpList
   
   ZonePrev = Empty
   nRow = 0
   ' 依地區別
   For nZoneCount = 0 To m_ZoneCount - 1
      ' 地區代碼
      ZoneCurr = m_ZoneList(nZoneCount).ZoneKind
      ' 地區別中的國籍串列
      For nY = 0 To m_ZoneList(nZoneCount).CountryCount - 1
         ' 若列數超過頁面的高度限制時則換頁
         If nRow > m_ReportDataRows Then
            Printer.NewPage
            nPage = nPage + 1
            PrintPageHeader_RP nPage
            nRow = 0
         End If
      
         ' 清除欄位
         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
         ' 代理人
         nRow = nRow + 1
         fld(0) = m_ZoneList(nZoneCount).CountryList(nY).CountryName
         fld(1) = "事務所"
         For nZ = 0 To Min(13, m_ZoneList(nZoneCount).CountryList(nY).AgentCount - 1)
            fld(nZ + 2) = m_ZoneList(nZoneCount).CountryList(nY).AgentList(nZ).AgentCompany
         Next nZ
         fld(15) = Empty
         fld(16) = Empty
         ' 輸出代理人列
         For nZ = 0 To 16
            Select Case nZ
               Case 0
                  Printer.FontSize = 10
                  Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                  Printer.Print fld(nZ)
               Case 1
                  Printer.FontSize = 12
                  Printer.CurrentX = m_Field(nZ).Left * m_CharWidth
                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                  Printer.Print fld(nZ)
               Case Else
                  Printer.FontSize = 8
                  nCenter = ((m_Field(nZ).Left * m_CharWidth) + (m_Field(nZ).Left + m_Field(nZ).Width) * m_CharWidth) / 2
                  Printer.CurrentX = nCenter - Printer.TextWidth(fld(nZ)) / 2
                  Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                  'Modify By Sindy 2024/4/24 解決列印出來是?
                  'Printer.Print fld(nZ)
                  PUB_PrintUnicodeText fld(nZ), Printer.CurrentX, Printer.CurrentY, 0
                  '2024/4/24 END
                  Printer.FontSize = 12
            End Select
         Next nZ
'=== 類
         'Add By Sindy 2012/3/2
         ' 清除欄位
         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
         ' 數量
         nRow = nRow + 1
         'Add By Sindy 2012/3/2
         fld(0) = "　　(類)"
         '2012/3/2 End
         fld(1) = "數量"
         For nZ = 0 To Min(13, m_ZoneList(nZoneCount).CountryList(nY).AgentCount - 1)
            fld(nZ + 2) = m_ZoneList(nZoneCount).CountryList(nY).AgentList(nZ).Count08
         Next nZ
         nNoAgentAmount = m_ZoneList(nZoneCount).CountryList(nY).NoAgentCount08
         nTotalAmount = m_ZoneList(nZoneCount).CountryList(nY).Count08 - nNoAgentAmount
         fld(15) = nNoAgentAmount
         fld(16) = nTotalAmount
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
         ' 百分比列
         nRow = nRow + 1
         fld(1) = "百分比"
         For nZ = 0 To Min(13, m_ZoneList(nZoneCount).CountryList(nY).AgentCount - 1)
            nAmount = m_ZoneList(nZoneCount).CountryList(nY).AgentList(nZ).Count08
            fValue = nAmount / nTotalAmount * 100
            fld(nZ + 2) = Format(fValue, "##0.00") & " %"
         Next nZ
         fld(15) = Empty
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
         '2012/3/2 End
'=== 件
         ' 清除欄位
         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
         ' 數量
         nRow = nRow + 1
         'Add By Sindy 2012/3/2
         fld(0) = "　　(件)"
         '2012/3/2 End
         fld(1) = "數量"
         For nZ = 0 To Min(13, m_ZoneList(nZoneCount).CountryList(nY).AgentCount - 1)
            fld(nZ + 2) = m_ZoneList(nZoneCount).CountryList(nY).AgentList(nZ).Count
         Next nZ
         nNoAgentAmount = m_ZoneList(nZoneCount).CountryList(nY).NoAgentCount
         nTotalAmount = m_ZoneList(nZoneCount).CountryList(nY).Count - nNoAgentAmount
         fld(15) = nNoAgentAmount
         fld(16) = nTotalAmount
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
         ' 百分比列
         nRow = nRow + 1
         fld(1) = "百分比"
         For nZ = 0 To Min(13, m_ZoneList(nZoneCount).CountryList(nY).AgentCount - 1)
            nAmount = m_ZoneList(nZoneCount).CountryList(nY).AgentList(nZ).Count
            fValue = nAmount / nTotalAmount * 100
            fld(nZ + 2) = Format(fValue, "##0.00") & " %"
         Next nZ
         fld(15) = Empty
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
         If nY < m_ZoneList(nZoneCount).CountryCount - 1 Then
            nRow = nRow + 1
            PrintSplitLine m_HeaderHeight + nRow
         End If
      Next nY
      
      ' 若該區域中的地區多於一個地區則需列印小計
      'Modify By Sindy 2011/1/25
      'If m_ZoneList(nZoneCount).CountryCount > 1 Then
      If m_ZoneList(nZoneCount).CountryCount >= 1 Then
         ' 列印分隔線
         nRow = nRow + 1
         PrintSplitLine m_HeaderHeight + nRow
         
         ' 計算區域別的小計
         BuildAgentTmpList08 m_ZoneList(nZoneCount)
         ' 清除欄位
         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
         ' 代理人
         nRow = nRow + 1
         fld(0) = "小計"
         fld(1) = "事務所"
         For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
            fld(nZ + 2) = m_AgentTmpList(nZ).AgentCompany
         Next nZ
         fld(15) = Empty
         fld(16) = Empty
         ' 輸出代理人列
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
                  'Modify By Sindy 2024/4/24 解決列印出來是?
                  'Printer.Print fld(nZ)
                  PUB_PrintUnicodeText fld(nZ), Printer.CurrentX, Printer.CurrentY, 0
                  '2024/4/24 END
                  Printer.FontSize = 12
            End Select
         Next nZ
         ' 清除欄位
         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
         ' 數量
         nRow = nRow + 1
         fld(0) = "　　(類)"
         fld(1) = "數量"
         For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
            fld(nZ + 2) = m_AgentTmpList(nZ).Count08
         Next nZ
         nNoAgentAmount = GetNoAgentAmountByZone08(m_ZoneList(nZoneCount))
         nTotalAmount = GetTotalAmountByZone08(m_ZoneList(nZoneCount))
         fld(15) = nNoAgentAmount
         fld(16) = nTotalAmount
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
         ' 百分比列
         nRow = nRow + 1
         fld(1) = "百分比"
         For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
            nAmount = m_AgentTmpList(nZ).Count08
            nTotalAmount = GetZoneAmount08(m_ZoneList(nZoneCount))
            fValue = nAmount / nTotalAmount * 100
            fld(nZ + 2) = Format(fValue, "##0.00") & " %"
         Next nZ
         fld(15) = Empty
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
         ' 清除暫存串列
         'edit by nickc 2005/09/09
         'ClearAgentTmpList
         
         'edit by nickc 2005/09/09
         'BuildAgentTmpList m_ZoneList(nZoneCount)
         For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
         ' 數量
         nRow = nRow + 1
         fld(0) = "　　(件)"
         fld(1) = "數量"
         For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
            fld(nZ + 2) = m_AgentTmpList(nZ).Count
         Next nZ
         nNoAgentAmount = GetNoAgentAmountByZone(m_ZoneList(nZoneCount))
         nTotalAmount = GetTotalAmountByZone(m_ZoneList(nZoneCount))
         fld(15) = nNoAgentAmount
         fld(16) = nTotalAmount
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
         ' 百分比列
         nRow = nRow + 1
         fld(1) = "百分比"
         For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
            nAmount = m_AgentTmpList(nZ).Count
            nTotalAmount = GetZoneAmount(m_ZoneList(nZoneCount))
            fValue = nAmount / nTotalAmount * 100
            fld(nZ + 2) = Format(fValue, "##0.00") & " %"
         Next nZ
         fld(15) = Empty
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
         ' 清除暫存串列
         ClearAgentTmpList
      End If
      
      ZonePrev = ZoneCurr
    
      ' 列印區域別的分隔線
      nRow = nRow + 1
      PrintTerminateLine m_HeaderHeight + nRow
      
      ' 若大區域不同時則需列印總計
      bPrintTotal = False
      If nZoneCount = m_ZoneCount - 1 Then
         bPrintTotal = True
      Else
         If Mid(m_ZoneList(nZoneCount).ZoneKind, 1, 1) <> Mid(m_ZoneList(nZoneCount + 1).ZoneKind, 1, 1) Then
            bPrintTotal = True
         End If
      End If
      If bPrintTotal = True Then
         Select Case Mid(m_ZoneList(nZoneCount).ZoneKind, 1, 1)
            'Modify By Cheng 2003/01/27
            '多傳入參數--頁數
'            Case "A": Generate_GrandTotal 0, nRow
'            Case "B": Generate_GrandTotal 1, nRow
'            Case "C": Generate_GrandTotal 2, nRow
            Case "A": Generate_GrandTotal_931214 0, nRow, nPage
            Case "B": Generate_GrandTotal_931214 1, nRow, nPage
            Case "C": Generate_GrandTotal_931214 2, nRow, nPage
         End Select
      End If
   Next nZoneCount
   
   Printer.EndDoc
End Sub


Private Function GetDBData_RP_931214(ByVal nReport As Integer) As Boolean
   Dim rsMain As New ADODB.Recordset
   Dim strSql As String
   Dim strSubSQL As String
   Dim strZoneKind, strZoneName, strZoneCode, strAgentName, strAgentCode, strAgentCompany As String
   Dim bFindZone, bFindCountry, bFindAgent As Boolean
   Dim nSortX, nSortY As Integer
   Dim AgentTemp As AGENTITEM
   Dim CountryTemp As COUNTRYITEM
   Dim ZoneTemp As ZONEITEM
   Dim bFromSec As Boolean
   Dim bToSec As Boolean
   Dim nX, nY, nZ As Integer
   Dim c1X, c1Y, c2X, c2Y As String
   Dim tmpArr As Variant
   Dim oStrTMBM08 As String
   
   GetDBData_RP_931214 = True
   
   strSubSQL = Empty
   Select Case textNA02
      Case "a", "A":
         'edit by nick 2004/12/14
         'strSubSQL = Empty
         strSubSQL = "NA02 LIKE '" & "A%" & "' "
      Case "b", "B":
         strSubSQL = "NA02 LIKE '" & "B%" & "' "
      Case "c", "C":
         strSubSQL = "NA02 LIKE '" & "C%" & "' "
      Case Else:
         strSubSQL = Empty
   End Select
   
   ' 產生SQL查詢語法
   bFromSec = Not IsEmptyText(textTMBM07_1.Text)
   bToSec = Not IsEmptyText(textTMBM07_2.Text)
   'Modify By Sindy 2013/8/19 + length(na01)=3 AND
   If bFromSec = True And bToSec = True Then
      strSql = "SELECT DISTINCT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,NA01,NA02,TA01,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04 FROM TMBULLETIN, NATION, TAGENT " & _
               "WHERE TMBM05 = NA03(+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01(+) AND " & _
                     "TMBM07 >= '" & textTMBM07_1 & "' AND " & _
                     "TMBM07 <= '" & textTMBM07_2 & "' "
      If strSubSQL <> Empty Then
         strSql = strSql & " " & "AND " & strSubSQL
      End If
   ElseIf bFromSec = True And bToSec = False Then
      strSql = "SELECT DISTINCT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,NA01,NA02,TA01,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04 FROM TMBULLETIN, NATION, TAGENT " & _
               "WHERE TMBM05 = NA03(+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01(+) AND " & _
                     "TMBM07 >= '" & textTMBM07_1 & "' "
      If strSubSQL <> Empty Then
         strSql = strSql & " " & "AND " & strSubSQL
      End If
   ElseIf bFromSec = False And bToSec = True Then
      strSql = "SELECT DISTINCT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,NA01,NA02,TA01,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04 FROM TMBULLETIN, NATION, TAGENT " & _
               "WHERE TMBM05 = NA03(+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01(+) AND " & _
                     "TMBM07 <= '" & textTMBM07_2 & "' "
      If strSubSQL <> Empty Then
         strSql = strSql & " " & "AND " & strSubSQL
      End If
   Else
      strSql = "SELECT DISTINCT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,NA01,NA02,TA01,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04 FROM TMBULLETIN, NATION, TAGENT " & _
               "WHERE TMBM05 = NA03(+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01(+) "
      If strSubSQL <> Empty Then
         strSql = strSql & " " & "AND " & strSubSQL
      End If
   End If

   ' 取得資料庫的資料
   rsMain.CursorLocation = adUseClient
   rsMain.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 無資料則離開
   If rsMain.RecordCount <= 0 Then
      GetDBData_RP_931214 = False
      GoTo EXITSUB
   End If

   ' 設定初始值
   m_ZoneCount = 0
   
   rsMain.MoveFirst
   ' 依序從資料記錄中取出欄位的內容
   While Not rsMain.EOF
      strAgentName = Empty
      If IsNull(rsMain.Fields("TMBM06")) = False Then
         strAgentName = rsMain.Fields("TMBM06")
      End If
      strAgentCode = Empty
      If IsNull(rsMain.Fields("TA02")) = False Then
         strAgentCode = rsMain.Fields("TA02")
      End If
      ' 代理人事務所名稱
      strAgentCompany = Empty
      If IsNull(rsMain.Fields("TA04")) = False Then
         strAgentCompany = rsMain.Fields("TA04")
      End If

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
      tmpArr = Split(oStrTMBM08, ",")
      
      ' 依區域別將國內及國外及大陸分類
      Select Case textNA02
         ' 國內
         Case "a", "A":
            Select Case Mid(strZoneKind, 1, 1)
               Case "A":
               Case "B":
                  strZoneKind = "B00"
                  strZoneCode = "998"
                  strZoneName = "大陸地區"
               Case "C":
                  strZoneKind = "C00"
                  strZoneCode = "999"
                  strZoneName = "國外地區"
               Case Else:
                  strZoneKind = "C00"
                  strZoneCode = "999"
                  strZoneName = "國外地區"
            End Select
         ' 大陸
         Case "b", "B":
            Select Case Mid(strZoneKind, 1, 1)
               Case "A":
                  strZoneKind = "A00"
                  strZoneCode = "997"
                  strZoneName = "國內地區"
               Case "B":
               Case "C":
                  strZoneKind = "C00"
                  strZoneCode = "999"
                  strZoneName = "國外地區"
               Case Else:
                  strZoneKind = "C00"
                  strZoneCode = "999"
                  strZoneName = "國外地區"
            End Select
         ' 國外
         Case "c", "C":
            Select Case Mid(strZoneKind, 1, 1)
               Case "A":
                  strZoneKind = "A00"
                  strZoneCode = "997"
                  strZoneName = "國內地區"
               Case "B":
                  strZoneKind = "B00"
                  strZoneCode = "998"
                  strZoneName = "大陸地區"
               Case "C":
               Case Else:
                  strZoneKind = "C99"
                  strZoneCode = "996"
                  strZoneName = "國外地區"
            End Select
      End Select
      
      ' 地區串列
      bFindZone = False
      For nX = 0 To m_ZoneCount - 1
         ' 找到地區別的結構
         If m_ZoneList(nX).ZoneKind = strZoneKind Then
            bFindZone = True
            
            bFindCountry = False
            For nY = 0 To m_ZoneList(nX).CountryCount - 1
               ' 找到地區別結構中的地區(國家)列表
               If m_ZoneList(nX).CountryList(nY).CountryCode = strZoneCode Then
                  bFindCountry = True
                  ' 計數加一
                  m_ZoneList(nX).CountryList(nY).Count = m_ZoneList(nX).CountryList(nY).Count + 1
                  m_ZoneList(nX).CountryList(nY).Count08 = m_ZoneList(nX).CountryList(nY).Count08 + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
                  ' 搜尋代理人串列
                  bFindAgent = False
                  'If strAgentCode <> Empty Then
                  If IsEmptyText(strAgentName) = False Then
                     For nZ = 0 To m_ZoneList(nX).CountryList(nY).AgentCount - 1
                        'If m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCode = strAgentCode Then
                        'Modify By Sindy 2010/02/26
                        'If m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentName = strAgentName Then
                        If m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCompany = strAgentCompany Then
                        '2010/02/26 End
                           bFindAgent = True
                           m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count = m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count + 1
                           m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count08 = m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count08 + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
                           Exit For
                        End If
                     Next nZ
                     ' 找不到此代理人的資料則新建一個代理人的結構
                     If bFindAgent = False Then
                        nZ = m_ZoneList(nX).CountryList(nY).AgentCount
                        ReDim Preserve m_ZoneList(nX).CountryList(nY).AgentList(nZ + 1)
                        m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCode = strAgentCode
                        'm_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentName = strAgentCompany
                        m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentName = strAgentName
                        m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCompany = strAgentCompany
                        m_ZoneList(nX).CountryList(nY).AgentCount = m_ZoneList(nX).CountryList(nY).AgentCount + 1
                        m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count = m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count + 1
                        m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count08 = m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count08 + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
                     End If
                  Else
                     m_ZoneList(nX).CountryList(nY).NoAgentCount = m_ZoneList(nX).CountryList(nY).NoAgentCount + 1
                     m_ZoneList(nX).CountryList(nY).NoAgentCount08 = m_ZoneList(nX).CountryList(nY).NoAgentCount08 + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
                  End If
               End If
            Next nY
            ' 找不到地區則新增地區
            If bFindCountry = False Then
               nY = m_ZoneList(nX).CountryCount
               ReDim Preserve m_ZoneList(nX).CountryList(nY + 1)
               m_ZoneList(nX).CountryList(nY).CountryCode = strZoneCode
               m_ZoneList(nX).CountryList(nY).CountryName = strZoneName
               m_ZoneList(nX).CountryList(nY).Count = 1
               m_ZoneList(nX).CountryList(nY).Count08 = IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
               m_ZoneList(nX).CountryList(nY).AgentCount = 0
               m_ZoneList(nX).CountryList(nY).AgentCount08 = 0
               m_ZoneList(nX).CountryList(nY).NoAgentCount = 0
               m_ZoneList(nX).CountryList(nY).NoAgentCount08 = 0
               m_ZoneList(nX).CountryCount = m_ZoneList(nX).CountryCount + 1
               m_ZoneList(nX).CountryCount08 = m_ZoneList(nX).CountryCount08 + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
               'If strAgentCode <> Empty Then
               If IsEmptyText(strAgentName) = False Then
                  nZ = 0
                  ReDim Preserve m_ZoneList(nX).CountryList(nY).AgentList(nZ + 1)
                  m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCode = strAgentCode
                  m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentName = strAgentName
                  m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCompany = strAgentCompany
                  m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count = m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count + 1
                  m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count08 = m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count08 + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
                  m_ZoneList(nX).CountryList(nY).AgentCount = 1
                  m_ZoneList(nX).CountryList(nY).AgentCount08 = IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
               Else
                  m_ZoneList(nX).CountryList(nY).NoAgentCount = 1
                  m_ZoneList(nX).CountryList(nY).NoAgentCount08 = IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
               End If
            End If
            Exit For
         End If
      Next nX
      
      ' 找不到地區別則新增地區別結構
      If bFindZone = False Then
         nX = m_ZoneCount
         ReDim Preserve m_ZoneList(nX + 1)
         m_ZoneList(nX).ZoneKind = strZoneKind
         m_ZoneList(nX).Count = 1
         m_ZoneList(nX).Count08 = IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
         m_ZoneList(nX).CountryCount = 0
         m_ZoneList(nX).CountryCount08 = 0
         m_ZoneCount = m_ZoneCount + 1
         nY = m_ZoneList(nX).CountryCount
         ReDim Preserve m_ZoneList(nX).CountryList(nY + 1)
         m_ZoneList(nX).CountryCount = 1
         m_ZoneList(nX).CountryCount08 = IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
         m_ZoneList(nX).CountryList(nY).CountryCode = strZoneCode
         m_ZoneList(nX).CountryList(nY).CountryName = strZoneName
         m_ZoneList(nX).CountryList(nY).Count = 1
         m_ZoneList(nX).CountryList(nY).Count08 = IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
         m_ZoneList(nX).CountryList(nY).AgentCount = 0
         m_ZoneList(nX).CountryList(nY).AgentCount08 = 0
         m_ZoneList(nX).CountryList(nY).NoAgentCount = 0
         m_ZoneList(nX).CountryList(nY).NoAgentCount08 = 0
         'If strAgentCode <> Empty Then
         If IsEmptyText(strAgentName) = False Then
            nZ = 0
            ReDim Preserve m_ZoneList(nX).CountryList(nY).AgentList(nZ + 1)
            m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCode = strAgentCode
            m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentName = strAgentName
            m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCompany = strAgentCompany
            m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count = m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count + 1
            m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count08 = m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count08 + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
            m_ZoneList(nX).CountryList(nY).AgentCount = 1
            m_ZoneList(nX).CountryList(nY).AgentCount08 = IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
         Else
            m_ZoneList(nX).CountryList(nY).NoAgentCount = 1
            m_ZoneList(nX).CountryList(nY).NoAgentCount08 = IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
         End If
      End If
      
      ' 移到下一筆記錄
      rsMain.MoveNext
   Wend
   
   ' 對地區別串列依地區別代碼小到大排序
   For nSortX = 0 To m_ZoneCount - 1
      For nSortY = nSortX To m_ZoneCount - 1
         If m_ZoneList(nSortX).ZoneKind > m_ZoneList(nSortY).ZoneKind Then
            ZoneTemp = m_ZoneList(nSortX)
            m_ZoneList(nSortX) = m_ZoneList(nSortY)
            m_ZoneList(nSortY) = ZoneTemp
         End If
      Next nSortY
   Next nSortX
   ' 對地區別中的地區(國籍)串列依國籍的代碼小到大排序
   For nX = 0 To m_ZoneCount - 1
      For nSortX = 0 To m_ZoneList(nX).CountryCount - 1
         For nSortY = nSortX To m_ZoneList(nX).CountryCount - 1
            'If m_ZoneList(nX).CountryList(nSortX).Count < m_ZoneList(nX).CountryList(nSortY).Count Then
            '   CountryTemp = m_ZoneList(nX).CountryList(nSortX)
            '   m_ZoneList(nX).CountryList(nSortX) = m_ZoneList(nX).CountryList(nSortY)
            '   m_ZoneList(nX).CountryList(nSortY) = CountryTemp
            'ElseIf m_ZoneList(nX).CountryList(nSortX).Count = m_ZoneList(nX).CountryList(nSortY).Count Then
            '   If m_ZoneList(nX).CountryList(nSortX).CountryCode > m_ZoneList(nX).CountryList(nSortY).CountryCode Then
            '      CountryTemp = m_ZoneList(nX).CountryList(nSortX)
            '      m_ZoneList(nX).CountryList(nSortX) = m_ZoneList(nX).CountryList(nSortY)
            '      m_ZoneList(nX).CountryList(nSortY) = CountryTemp
            '   End If
            'End If
            If m_ZoneList(nX).CountryList(nSortX).CountryCode > m_ZoneList(nX).CountryList(nSortY).CountryCode Then
               CountryTemp = m_ZoneList(nX).CountryList(nSortX)
               m_ZoneList(nX).CountryList(nSortX) = m_ZoneList(nX).CountryList(nSortY)
               m_ZoneList(nX).CountryList(nSortY) = CountryTemp
            End If
         Next nSortY
      Next nSortX
   Next nX
   ' 對地區別中的地區(國籍)串列項目中的代理人串列依數量的多寡由大到小排序
   For nX = 0 To m_ZoneCount - 1 '地區別
      For nY = 0 To m_ZoneList(nX).CountryCount - 1 '國家
         For nSortX = 0 To m_ZoneList(nX).CountryList(nY).AgentCount - 1 '代理人
            For nSortY = nSortX To m_ZoneList(nX).CountryList(nY).AgentCount - 1
               If m_ZoneList(nX).CountryList(nY).AgentList(nSortX).Count08 < m_ZoneList(nX).CountryList(nY).AgentList(nSortY).Count08 Then
                  AgentTemp = m_ZoneList(nX).CountryList(nY).AgentList(nSortX)
                  m_ZoneList(nX).CountryList(nY).AgentList(nSortX) = m_ZoneList(nX).CountryList(nY).AgentList(nSortY)
                  m_ZoneList(nX).CountryList(nY).AgentList(nSortY) = AgentTemp
               'Modify By Sindy 2014/10/28 若類別數相同者,再依案件數排序
               ElseIf m_ZoneList(nX).CountryList(nY).AgentList(nSortX).Count08 = m_ZoneList(nX).CountryList(nY).AgentList(nSortY).Count08 Then
                  If m_ZoneList(nX).CountryList(nY).AgentList(nSortX).Count < m_ZoneList(nX).CountryList(nY).AgentList(nSortY).Count Then
                     AgentTemp = m_ZoneList(nX).CountryList(nY).AgentList(nSortX)
                     m_ZoneList(nX).CountryList(nY).AgentList(nSortX) = m_ZoneList(nX).CountryList(nY).AgentList(nSortY)
                     m_ZoneList(nX).CountryList(nY).AgentList(nSortY) = AgentTemp
                  End If
               '2014/10/28 END
               End If
            Next nSortY
         Next nSortX
      Next nY
   Next nX
      
EXITSUB:
   rsMain.Close
   Set rsMain = Nothing
End Function



Private Sub BuildAgentTmpList08(ByRef ZoneInfo As ZONEITEM)
   Dim nX As Integer
   Dim nY As Integer
   Dim nZ As Integer
   Dim nIndex As Integer
   Dim bFindAgent As Boolean
   Dim AgentTmp As AGENTITEM
   ' 清除計算區域別代理人數量的串列
   'edit by nickc 2005/09/09
   'ClearAgentTmpList
      
   For nX = 0 To ZoneInfo.CountryCount - 1
      For nY = 0 To ZoneInfo.CountryList(nX).AgentCount - 1
         bFindAgent = False
         ' 搜尋原有的暫存串列
         For nZ = 0 To m_AgentTmpListCount - 1
            'If ZoneInfo.CountryList(nX).AgentList(nY).AgentCode = m_AgentTmpList(nZ).AgentCode Then
            'Modify By Sindy 2010/02/26
            'If ZoneInfo.CountryList(nX).AgentList(nY).AgentName = m_AgentTmpList(nZ).AgentName Then
            If ZoneInfo.CountryList(nX).AgentList(nY).AgentCompany = m_AgentTmpList(nZ).AgentCompany Then
            '2010/02/26 End
               bFindAgent = True
               m_AgentTmpList(nZ).Count08 = m_AgentTmpList(nZ).Count08 + ZoneInfo.CountryList(nX).AgentList(nY).Count08
               'add by nickc 2005/09/09
               m_AgentTmpList(nZ).Count = m_AgentTmpList(nZ).Count + ZoneInfo.CountryList(nX).AgentList(nY).Count
               Exit For
            End If
         Next nZ
         If bFindAgent = False Then
            nIndex = m_AgentTmpListCount
            ReDim Preserve m_AgentTmpList(nIndex + 1)
            m_AgentTmpList(nIndex).AgentCode = ZoneInfo.CountryList(nX).AgentList(nY).AgentCode
            m_AgentTmpList(nIndex).AgentName = ZoneInfo.CountryList(nX).AgentList(nY).AgentName
            m_AgentTmpList(nIndex).AgentCompany = ZoneInfo.CountryList(nX).AgentList(nY).AgentCompany
            m_AgentTmpList(nIndex).Count08 = ZoneInfo.CountryList(nX).AgentList(nY).Count08
            'add by nickc 2005/09/09
            m_AgentTmpList(nIndex).Count = ZoneInfo.CountryList(nX).AgentList(nY).Count
            m_AgentTmpListCount = m_AgentTmpListCount + 1
         End If
      Next nY
   Next nX
   ' 排序
   For nX = 0 To m_AgentTmpListCount - 1
      For nY = nX To m_AgentTmpListCount - 1
         If m_AgentTmpList(nX).Count08 < m_AgentTmpList(nY).Count08 Then
            AgentTmp = m_AgentTmpList(nX)
            m_AgentTmpList(nX) = m_AgentTmpList(nY)
            m_AgentTmpList(nY) = AgentTmp
         'Modify By Sindy 2014/10/28 若類別數相同者,再依案件數排序
         ElseIf m_AgentTmpList(nX).Count08 = m_AgentTmpList(nY).Count08 Then
            If m_AgentTmpList(nX).Count < m_AgentTmpList(nY).Count Then
               AgentTmp = m_AgentTmpList(nX)
               m_AgentTmpList(nX) = m_AgentTmpList(nY)
               m_AgentTmpList(nY) = AgentTmp
            End If
         '2014/10/28 END
         End If
      Next nY
   Next nX
End Sub

Private Function GetNoAgentAmountByZone08(ByRef ZoneInfo As ZONEITEM) As Long
   Dim nX As Integer
   Dim nAmount As Long
   nAmount = 0
   For nX = 0 To ZoneInfo.CountryCount - 1
      nAmount = nAmount + ZoneInfo.CountryList(nX).NoAgentCount08
   Next nX
   GetNoAgentAmountByZone08 = nAmount
End Function
Private Function GetTotalAmountByZone08(ByRef ZoneInfo As ZONEITEM) As Long
   Dim nX As Integer
   Dim nAmount As Long
   nAmount = 0
   For nX = 0 To ZoneInfo.CountryCount - 1
      nAmount = nAmount + ZoneInfo.CountryList(nX).Count08 - ZoneInfo.CountryList(nX).NoAgentCount08
   Next nX
   GetTotalAmountByZone08 = nAmount
End Function
' 取得區域別的數量總計
Private Function GetZoneAmount08(ByRef ZoneInfo As ZONEITEM) As Integer
   Dim nX As Integer
   Dim nY As Integer
   Dim nAmount As Variant
   
   nAmount = 0
   For nX = 0 To ZoneInfo.CountryCount - 1
      For nY = 0 To ZoneInfo.CountryList(nX).AgentCount - 1
         nAmount = nAmount + ZoneInfo.CountryList(nX).AgentList(nY).Count08
      Next nY
   Next nX
   GetZoneAmount08 = nAmount
End Function
Private Sub Generate_GrandTotal_931214(ByVal nCountry As Integer, ByRef nRow As Integer, ByRef nPage As Integer)
   Dim fld(16) As String
   Dim nAmount As Variant
   Dim nNoAgentAmount As Variant
   Dim nTotalAmount As Variant
   Dim nNoAgentAmount08 As Variant
   Dim nTotalAmount08 As Variant
   Dim fValue As Double
   Dim nX As Long
   Dim nY As Long
   Dim nZ As Long
   Dim nCenter As Long
   Dim nRight As Long
   
    'Add By Cheng 2003/01/27
    ' 若列數超過頁面的高度限制時則換頁
    If nRow > m_ReportDataRows Then
       Printer.NewPage
       nPage = nPage + 1
       PrintPageHeader_RP nPage
       nRow = 0
    End If
   ' 輸出事務所
   For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
   nRow = nRow + 1
   Select Case nCountry
      Case 0:
         fld(0) = "國內合計"
         BuildTaiwanTmpList08
         nNoAgentAmount08 = GetNoAgentAmountByCountry08(0)
         nTotalAmount08 = GetAgentTotalAmountByCountry08(0)
     Case 1:
         fld(0) = "大陸合計"
         BuildChinaTmpList08
         nNoAgentAmount08 = GetNoAgentAmountByCountry08(1)
         nTotalAmount08 = GetAgentTotalAmountByCountry08(1)
      Case 2:
         fld(0) = "國外合計"
         BuildForeignTmpList08
         nNoAgentAmount08 = GetNoAgentAmountByCountry08(2)
         nTotalAmount08 = GetAgentTotalAmountByCountry08(2)
      Case 3:
         fld(0) = "國外合計"
         BuildForeignTmpList08
         nNoAgentAmount08 = GetNoAgentAmountByCountry08(2)
         nTotalAmount08 = GetAgentTotalAmountByCountry08(2)
   End Select
   fld(1) = "事務所"
   For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
   fld(nZ + 2) = m_AgentTmpList(nZ).AgentCompany
   Next nZ
   fld(15) = Empty
   fld(16) = Empty
   ' 輸出代理人列
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
   
    If nRow > m_ReportDataRows Then
       Printer.NewPage
       nPage = nPage + 1
       PrintPageHeader_RP nPage
       nRow = 0
    End If
    
   ' 輸出數量
   For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
   nRow = nRow + 1
   fld(0) = "　　(類)"
   fld(1) = "數量"
   For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
      fld(nZ + 2) = m_AgentTmpList(nZ).Count08
   Next nZ
   fld(15) = nNoAgentAmount08
   fld(16) = nTotalAmount08
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
   
    If nRow > m_ReportDataRows Then
       Printer.NewPage
       nPage = nPage + 1
       PrintPageHeader_RP nPage
       nRow = 0
    End If
    
   ' 輸出百分比
   For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
   nRow = nRow + 1
   fld(0) = Empty
   fld(1) = "百分比"
   For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
      nAmount = m_AgentTmpList(nZ).Count08
      fValue = nAmount / nTotalAmount08 * 100
      fld(nZ + 2) = Format(fValue, "##0.00") & " %"
   Next nZ
   'fValue = nNoAgentAmount / nTotalAmount * 100
   'fld(15) = Format(fValue, "##0.00") & " %"
   fld(15) = Empty
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
    If nRow > m_ReportDataRows Then
       Printer.NewPage
       nPage = nPage + 1
       PrintPageHeader_RP nPage
       nRow = 0
    End If
    
   For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
   nRow = nRow + 1
   
    Select Case nCountry
      Case 0:
         'edit by nickc 2005/09/09
         'BuildTaiwanTmpList
         nNoAgentAmount = GetNoAgentAmountByCountry(0)
         nTotalAmount = GetAgentTotalAmountByCountry(0)
     Case 1:
         'edit by nickc  2005/09/09
         'BuildChinaTmpList
         nNoAgentAmount = GetNoAgentAmountByCountry(1)
         nTotalAmount = GetAgentTotalAmountByCountry(1)
      Case 2:
         'edit by nickc 2005/09/09
         'BuildForeignTmpList
         nNoAgentAmount = GetNoAgentAmountByCountry(2)
         nTotalAmount = GetAgentTotalAmountByCountry(2)
      Case 3:
         'edit by nickc 2005/09/09
         'BuildForeignTmpList
         nNoAgentAmount = GetNoAgentAmountByCountry(2)
         nTotalAmount = GetAgentTotalAmountByCountry(2)
   End Select
   fld(0) = "　　(件)"
   fld(1) = "數量"
   For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
      fld(nZ + 2) = m_AgentTmpList(nZ).Count
   Next nZ
   fld(15) = nNoAgentAmount
   fld(16) = nTotalAmount
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
    If nRow > m_ReportDataRows Then
       Printer.NewPage
       nPage = nPage + 1
       PrintPageHeader_RP nPage
       nRow = 0
    End If
    
   ' 輸出百分比
   For nZ = 0 To 16: fld(nZ) = Empty: Next nZ
   nRow = nRow + 1
   fld(0) = Empty
   fld(1) = "百分比"
   For nZ = 0 To Min(13, m_AgentTmpListCount - 1)
      nAmount = m_AgentTmpList(nZ).Count
      fValue = nAmount / nTotalAmount * 100
      fld(nZ + 2) = Format(fValue, "##0.00") & " %"
   Next nZ
   'fValue = nNoAgentAmount / nTotalAmount * 100
   'fld(15) = Format(fValue, "##0.00") & " %"
   fld(15) = Empty
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
    If nRow > m_ReportDataRows Then
       Printer.NewPage
       nPage = nPage + 1
       PrintPageHeader_RP nPage
       nRow = 0
    End If
   ' 清除暫存串列
   ClearAgentTmpList
End Sub

Private Function GetNoAgentAmountByCountry08(ByVal nCountry As Integer) As Long
   Dim nAmount As Long
   Dim nX As Long
   Dim nY As Long
   nAmount = 0
   For nX = 0 To m_ZoneCount - 1
      Select Case nCountry
         Case 0:
            If Mid(m_ZoneList(nX).ZoneKind, 1, 1) <> "A" Then
               GoTo NextRecord
            End If
         Case 1:
            If Mid(m_ZoneList(nX).ZoneKind, 1, 1) <> "B" Then
               GoTo NextRecord
            End If
         Case 2:
            If Mid(m_ZoneList(nX).ZoneKind, 1, 1) <> "C" Then
               GoTo NextRecord
            End If
         Case 3:
            If Mid(m_ZoneList(nX).ZoneKind, 1, 1) = "A" Then
               GoTo NextRecord
            End If
         Case Else
      End Select
      For nY = 0 To m_ZoneList(nX).CountryCount - 1
         nAmount = nAmount + m_ZoneList(nX).CountryList(nY).NoAgentCount08
      Next nY
NextRecord:
   Next nX
   GetNoAgentAmountByCountry08 = nAmount
End Function

Private Function GetAgentTotalAmountByCountry08(ByVal nCountry As Integer) As Long
   Dim nAmount As Long
   Dim nX As Long
   Dim nY As Long
   nAmount = 0
   For nX = 0 To m_ZoneCount - 1
      Select Case nCountry
         Case 0:
            If Mid(m_ZoneList(nX).ZoneKind, 1, 1) <> "A" Then
               GoTo NextRecord
            End If
         Case 1:
            If Mid(m_ZoneList(nX).ZoneKind, 1, 1) <> "B" Then
               GoTo NextRecord
            End If
         Case 2:
            If Mid(m_ZoneList(nX).ZoneKind, 1, 1) <> "C" Then
               GoTo NextRecord
            End If
         Case 3:
            If Mid(m_ZoneList(nX).ZoneKind, 1, 1) = "A" Then
               GoTo NextRecord
            End If
      End Select
      For nY = 0 To m_ZoneList(nX).CountryCount - 1
         nAmount = nAmount + m_ZoneList(nX).CountryList(nY).Count08 - m_ZoneList(nX).CountryList(nY).NoAgentCount08
      Next nY
NextRecord:
   Next nX
   GetAgentTotalAmountByCountry08 = nAmount
End Function

Private Sub BuildTaiwanTmpList08()
   Dim nX As Integer
   Dim nY As Integer
   Dim nZ As Integer
   Dim nIndex As Integer
   Dim bFindAgent As Boolean
   Dim AgentTmp As AGENTITEM
   ' 清除計算區域別代理人數量的串列
   'edit by nickc 2005/09/09
   'ClearAgentTmpList
      
   For nX = 0 To m_ZoneCount - 1
      If Mid(m_ZoneList(nX).ZoneKind, 1, 1) = "A" Then
         For nY = 0 To m_ZoneList(nX).CountryCount - 1
            For nZ = 0 To m_ZoneList(nX).CountryList(nY).AgentCount - 1
               bFindAgent = False
               ' 搜尋原有的暫存串列
               For nIndex = 0 To m_AgentTmpListCount - 1
                  'If m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCode = m_AgentTmpList(nIndex).AgentCode Then
                  'Modify By Sindy 2010/02/26
                  'If m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentName = m_AgentTmpList(nIndex).AgentName Then
                  If m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCompany = m_AgentTmpList(nIndex).AgentCompany Then
                  '2010/02/26 End
                     bFindAgent = True
                     m_AgentTmpList(nIndex).Count08 = m_AgentTmpList(nIndex).Count08 + m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count08
                     'add by nickc 2005/09/09
                     m_AgentTmpList(nIndex).Count = m_AgentTmpList(nIndex).Count + m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count
                     Exit For
                  End If
               Next nIndex
               If bFindAgent = False Then
                  nIndex = m_AgentTmpListCount
                  ReDim Preserve m_AgentTmpList(nIndex + 1)
                  m_AgentTmpList(nIndex).AgentCode = m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCode
                  m_AgentTmpList(nIndex).AgentName = m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentName
                  m_AgentTmpList(nIndex).AgentCompany = m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCompany
                  m_AgentTmpList(nIndex).Count08 = m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count08
                  'add by nickc 2005/09/09
                  m_AgentTmpList(nIndex).Count = m_AgentTmpList(nIndex).Count + m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count
                  m_AgentTmpListCount = m_AgentTmpListCount + 1
               End If
            Next nZ
         Next nY
      End If
   Next nX
   ' 排序
   For nX = 0 To m_AgentTmpListCount - 1
      For nY = nX To m_AgentTmpListCount - 1
         If m_AgentTmpList(nX).Count08 < m_AgentTmpList(nY).Count08 Then
            AgentTmp = m_AgentTmpList(nX)
            m_AgentTmpList(nX) = m_AgentTmpList(nY)
            m_AgentTmpList(nY) = AgentTmp
         End If
      Next nY
   Next nX
End Sub

Private Sub BuildChinaTmpList08()
   Dim nX As Integer
   Dim nY As Integer
   Dim nZ As Integer
   Dim nIndex As Integer
   Dim bFindAgent As Boolean
   Dim AgentTmp As AGENTITEM
   ' 清除計算區域別代理人數量的串列
   'edit by nickc 2005/09/09
   'ClearAgentTmpList
      
   For nX = 0 To m_ZoneCount - 1
      If Mid(m_ZoneList(nX).ZoneKind, 1, 1) = "B" Then
         For nY = 0 To m_ZoneList(nX).CountryCount - 1
            For nZ = 0 To m_ZoneList(nX).CountryList(nY).AgentCount - 1
               bFindAgent = False
               ' 搜尋原有的暫存串列
               For nIndex = 0 To m_AgentTmpListCount - 1
                  'Modify By Sindy 2010/4/23
                  'If m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentName = m_AgentTmpList(nIndex).AgentName Then
                  If m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCompany = m_AgentTmpList(nIndex).AgentCompany Then
                  '2010/4/23 End
                     bFindAgent = True
                     m_AgentTmpList(nIndex).Count08 = m_AgentTmpList(nIndex).Count08 + m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count08
                     'add by nickc 2005/09/09
                     m_AgentTmpList(nIndex).Count = m_AgentTmpList(nIndex).Count + m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count
                     Exit For
                  End If
               Next nIndex
               If bFindAgent = False Then
                  nIndex = m_AgentTmpListCount
                  ReDim Preserve m_AgentTmpList(nIndex + 1)
                  m_AgentTmpList(nIndex).AgentCode = m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCode
                  m_AgentTmpList(nIndex).AgentName = m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentName
                  m_AgentTmpList(nIndex).AgentCompany = m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCompany
                  m_AgentTmpList(nIndex).Count08 = m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count08
                  'add by nickc 2005/09/09
                  m_AgentTmpList(nIndex).Count = m_AgentTmpList(nIndex).Count + m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count
                  m_AgentTmpListCount = m_AgentTmpListCount + 1
               End If
            Next nZ
         Next nY
      End If
   Next nX
   ' 排序
   For nX = 0 To m_AgentTmpListCount - 1
      For nY = nX To m_AgentTmpListCount - 1
         If m_AgentTmpList(nX).Count08 < m_AgentTmpList(nY).Count08 Then
            AgentTmp = m_AgentTmpList(nX)
            m_AgentTmpList(nX) = m_AgentTmpList(nY)
            m_AgentTmpList(nY) = AgentTmp
         End If
      Next nY
   Next nX
End Sub

' 將國外的資料組成一個Temp List
Private Sub BuildForeignTmpList08()
   Dim nX As Integer
   Dim nY As Integer
   Dim nZ As Integer
   Dim nIndex As Integer
   Dim bFindAgent As Boolean
   Dim AgentTmp As AGENTITEM
   ' 清除計算區域別代理人數量的串列
   'edit by nickc 2005/09/09
   'ClearAgentTmpList
      
   For nX = 0 To m_ZoneCount - 1
      If Mid(m_ZoneList(nX).ZoneKind, 1, 1) = "C" Then
         For nY = 0 To m_ZoneList(nX).CountryCount - 1
            For nZ = 0 To m_ZoneList(nX).CountryList(nY).AgentCount - 1
               bFindAgent = False
               ' 搜尋原有的暫存串列
               For nIndex = 0 To m_AgentTmpListCount - 1
                  'Modify By Sindy 2010/4/23
                  'If m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentName = m_AgentTmpList(nIndex).AgentName Then
                  If m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCompany = m_AgentTmpList(nIndex).AgentCompany Then
                  '2010/4/23 End
                     bFindAgent = True
                     m_AgentTmpList(nIndex).Count08 = m_AgentTmpList(nIndex).Count08 + m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count08
                     'add by nickc 2005/09/09
                     m_AgentTmpList(nIndex).Count = m_AgentTmpList(nIndex).Count + m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count
                     Exit For
                  End If
               Next nIndex
               If bFindAgent = False Then
                  nIndex = m_AgentTmpListCount
                  ReDim Preserve m_AgentTmpList(nIndex + 1)
                  m_AgentTmpList(nIndex).AgentCode = m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCode
                  m_AgentTmpList(nIndex).AgentName = m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentName
                  m_AgentTmpList(nIndex).AgentCompany = m_ZoneList(nX).CountryList(nY).AgentList(nZ).AgentCompany
                  m_AgentTmpList(nIndex).Count08 = m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count08
                  'add by nickc 2005/09/09
                  m_AgentTmpList(nIndex).Count = m_AgentTmpList(nIndex).Count + m_ZoneList(nX).CountryList(nY).AgentList(nZ).Count
                  m_AgentTmpListCount = m_AgentTmpListCount + 1
               End If
            Next nZ
         Next nY
      End If
   Next nX
   ' 排序
   For nX = 0 To m_AgentTmpListCount - 1
      For nY = nX To m_AgentTmpListCount - 1
         If m_AgentTmpList(nX).Count08 < m_AgentTmpList(nY).Count08 Then
            AgentTmp = m_AgentTmpList(nX)
            m_AgentTmpList(nX) = m_AgentTmpList(nY)
            m_AgentTmpList(nY) = AgentTmp
         End If
      Next nY
   Next nX
End Sub
