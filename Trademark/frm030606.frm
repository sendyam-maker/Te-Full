VERSION 5.00
Begin VB.Form frm030606 
   BorderStyle     =   1  '單線固定
   Caption         =   "表一＆表二.商標全國市場統計表"
   ClientHeight    =   1700
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1700
   ScaleWidth      =   4560
   Begin VB.TextBox txt1 
      Height          =   264
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "1"
      Top             =   1110
      Width           =   465
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      ItemData        =   "frm030606.frx":0000
      Left            =   1380
      List            =   "frm030606.frx":0002
      TabIndex        =   5
      Top             =   1740
      Visible         =   0   'False
      Width           =   3072
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   3540
      TabIndex        =   4
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2580
      TabIndex        =   3
      Top             =   60
      Width           =   912
   End
   Begin VB.TextBox textTMBM07_2 
      Height          =   264
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   1
      Top             =   780
      Width           =   1092
   End
   Begin VB.TextBox textTMBM07_1 
      Height          =   264
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   0
      Top             =   780
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "列印份數："
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1110
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "印表機 :"
      Height          =   255
      Left            =   300
      TabIndex        =   7
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
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   780
      Width           =   975
   End
End
Attribute VB_Name = "frm030606"
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
   Company As String
   Count As Variant
   'add by nick  加入計算類別
   Count08 As Variant
   KindAmount(8) As Variant
   'add by nick  加入計算類別
   KindAmount08(8) As Variant
   ZoneAmount(3) As Variant
   'add by nick  加入計算類別
   ZoneAmount08(3) As Variant
End Type
' 定義代理人陣列
Dim m_AgentList() As AGENTITEM
Dim m_AgentCount As Variant
' 定義無代理人的儲存結構變數
Dim m_NoAgentItem As AGENTITEM

' 宣告地區項目的資料型態
Private Type ZONEITEM
   ' 代理人代號
   ZoneCode As String
   ZoneName As String
   Count As Variant
   'add by nick  加入計算類別
   Count08 As Variant
   TaieCount As Variant
   'add by nick  加入計算類別
   TaieCount08 As Variant
End Type
' 定義地區串列
Dim m_ZoneList() As ZONEITEM
Dim m_ZoneCount As Variant
'edit by nick 2004/12/14
'Dim m_DefaultPrinter As String
Dim nCopys As Integer


Private Sub Form_Load()
'edit by nick 2004/12/14
'   Dim Prn As Printer
   Dim nIndex As Integer
'edit by nick 2004/12/14
'   m_DefaultPrinter = Printer.DeviceName
   MoveFormToCenter Me
'edit by nick 2004/12/14
'   For Each Prn In Printers
'      If Prn.DeviceName <> m_DefaultPrinter Then
'         cmbPrinter.AddItem Prn.DeviceName
'      End If
'   Next
'   cmbPrinter.ListIndex = 0
   
   ' 初始化
   m_NoAgentItem.Count = 0
   For nIndex = 0 To 8
      m_NoAgentItem.KindAmount(nIndex) = 0
   Next nIndex
   For nIndex = 0 To 3
      m_NoAgentItem.ZoneAmount(nIndex) = 0
   Next nIndex
   m_NoAgentItem.AgentCode = Empty
   m_NoAgentItem.AgentName = Empty
   m_NoAgentItem.Company = Empty
   m_AgentCount = 0
   m_ZoneCount = 0
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
   Set frm030606 = Nothing
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
'edit by nick 2004/12/13
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
      BuildField_RP 1
      ' 取得資料庫中的資料
      'edit by nick 2004/12/13 加入計算商品類別
      'If GetDBData_RP = False Then
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/21 清除查詢印表記錄檔欄位
      If Len(textTMBM07_1) <> 0 Or Len(textTMBM07_2) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & Label1 & textTMBM07_1 & "-" & textTMBM07_1 'Add By Sindy 2010/10/21
      End If
      If GetDBData_RP_931213 = False Then
         GoTo EXITSUB
      End If
      If Trim(txt1.Text) = "" Then txt1.Text = "1" 'Add By Sindy 2010/01/12
      pub_QL05 = pub_QL05 & ";" & Label2 & txt1 'Add By Sindy 2010/10/21
      InsertQueryLog ("") 'Add By Sindy 2010/10/21
      For nCopys = 1 To Trim(txt1.Text) '5
         ' 列印
         'edit by nick 2004/12/13
         'Generate_RP
         Generate_RP_931213
         ' 測試(輸出到畫面上)
         'Generate_RP_SCREEN
      Next nCopys
      ' 清除所佔用的空間
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
' 檢查是否為台一
Public Function IsTaie(ByVal strCode As String) As Boolean
   If strCode = "001" Or strCode = "林晉章" Then
      IsTaie = True
   Else
      IsTaie = False
   End If
End Function

' 取得該代理人的資料
' Input : nType 取得資料的種類
'         0 表取得商標類的資料
'           nIndex = 如下說明
'              0 表全部合計, 1 表正商標, 2 表聯合商標, 3 防護商標
'              4 服務標章, 5 聯合服務標章, 6 防護服務標章, 7 證明標章
'              8 團體標章, 9團體商標
'         1 表取得國籍性的資料
'           nIndex = 如下說明
'              0 表全國合計, 1 表國內合計, 2 表大陸合計, 3 表國外合計
Private Function GetAgentAmount(ByRef agent As AGENTITEM, ByVal nType As Integer, ByVal nIndex As Integer) As Variant
   Dim nAmount As Variant
   
   nAmount = 0
   Select Case nType
      Case 0
         Select Case nIndex
            Case 0: nAmount = agent.KindAmount(0) + agent.KindAmount(1) + agent.KindAmount(2) + agent.KindAmount(3) + agent.KindAmount(4) + agent.KindAmount(5) + agent.KindAmount(6) + agent.KindAmount(7)
            'Modify By Sindy 2010/4/27
            'Case 1, 2, 3, 4, 5, 6, 7, 8:
            Case 1, 2, 3, 4, 5, 6, 7, 8, 9:
               nAmount = agent.KindAmount(nIndex - 1)
         End Select
      Case 1
         Select Case nIndex
            Case 0: nAmount = agent.ZoneAmount(0) + agent.ZoneAmount(1) + agent.ZoneAmount(2)
            Case 1, 2, 3:
               nAmount = agent.ZoneAmount(nIndex - 1)
         End Select
   End Select
   
   GetAgentAmount = nAmount
End Function

Private Function GetAgentAmount08(ByRef agent As AGENTITEM, ByVal nType As Integer, ByVal nIndex As Integer) As Variant
   Dim nAmount As Variant
   
   nAmount = 0
   Select Case nType
      Case 0
         Select Case nIndex
            Case 0: nAmount = agent.KindAmount08(0) + agent.KindAmount08(1) + agent.KindAmount08(2) + agent.KindAmount08(3) + agent.KindAmount08(4) + agent.KindAmount08(5) + agent.KindAmount08(6) + agent.KindAmount08(7)
            'Modify By Sindy 2010/4/27
            'Case 1, 2, 3, 4, 5, 6, 7, 8:
            Case 1, 2, 3, 4, 5, 6, 7, 8, 9:
               nAmount = agent.KindAmount08(nIndex - 1)
         End Select
      Case 1
         Select Case nIndex
            Case 0: nAmount = agent.ZoneAmount08(0) + agent.ZoneAmount08(1) + agent.ZoneAmount08(2)
            Case 1, 2, 3:
               nAmount = agent.ZoneAmount08(nIndex - 1)
         End Select
   End Select
   
   GetAgentAmount08 = nAmount
End Function
' 取得所有的資料
' Input : nType 取得資料的種類
'         0 表取得商標類的資料
'           nIndex = 如下說明
'              0 表全部合計, 1 表正商標, 2 表聯合商標, 3 防護商標
'              4 服務標章, 5 聯合服務標章, 6 防護服務標章, 7 證明標章
'              8 團體標章, 9團體商標
'         1 表取得國籍性的資料
'           nIndex = 如下說明
'              0 表全國合計, 1 表國內合計, 2 表大陸合計, 3 表國外合計
Private Function GetTotalAmount(ByVal nType As Integer, ByVal nIndex As Integer, Optional ByVal bIncludeNoAgent As Boolean = True) As Variant
   Dim nAgentCount As Variant
   Dim nAmount As Variant
   nAmount = 0
   
   Select Case nType
      Case 0
         Select Case nIndex
            Case 0:
               For nAgentCount = 0 To m_AgentCount - 1
                  nAmount = nAmount + m_AgentList(nAgentCount).KindAmount(0) + m_AgentList(nAgentCount).KindAmount(1) + m_AgentList(nAgentCount).KindAmount(2) + m_AgentList(nAgentCount).KindAmount(3) + m_AgentList(nAgentCount).KindAmount(4) + m_AgentList(nAgentCount).KindAmount(5) + m_AgentList(nAgentCount).KindAmount(6) + m_AgentList(nAgentCount).KindAmount(7)
               Next nAgentCount
               If bIncludeNoAgent = True Then
                  nAmount = nAmount + m_NoAgentItem.KindAmount(0) + m_NoAgentItem.KindAmount(1) + m_NoAgentItem.KindAmount(2) + m_NoAgentItem.KindAmount(3) + m_NoAgentItem.KindAmount(4) + m_NoAgentItem.KindAmount(5) + m_NoAgentItem.KindAmount(6) + m_NoAgentItem.KindAmount(7)
               End If
            'Modify By Sindy 2010/4/27
            'Case 1, 2, 3, 4, 5, 6, 7, 8:
            Case 1, 2, 3, 4, 5, 6, 7, 8, 9:
               For nAgentCount = 0 To m_AgentCount - 1
                  nAmount = nAmount + m_AgentList(nAgentCount).KindAmount(nIndex - 1)
               Next nAgentCount
               If bIncludeNoAgent = True Then
                  nAmount = nAmount + m_NoAgentItem.KindAmount(nIndex - 1)
               End If
         End Select
      Case 1
         Select Case nIndex
            Case 0:
               For nAgentCount = 0 To m_AgentCount - 1
                  nAmount = nAmount + m_AgentList(nAgentCount).ZoneAmount(0) + m_AgentList(nAgentCount).ZoneAmount(1) + m_AgentList(nAgentCount).ZoneAmount(2)
               Next nAgentCount
               If bIncludeNoAgent = True Then
                  nAmount = nAmount + m_NoAgentItem.ZoneAmount(0) + m_NoAgentItem.ZoneAmount(1) + m_NoAgentItem.ZoneAmount(2)
               End If
            Case 1, 2, 3:
               For nAgentCount = 0 To m_AgentCount - 1
                  nAmount = nAmount + m_AgentList(nAgentCount).ZoneAmount(nIndex - 1)
               Next nAgentCount
               If bIncludeNoAgent = True Then
                  nAmount = nAmount + m_NoAgentItem.ZoneAmount(nIndex - 1)
               End If
         End Select
   End Select
   GetTotalAmount = nAmount
End Function

' 取得台一的件數
' Input : nType 取得資料的種類
'         0 表取得商標類的資料
'           nIndex = 如下說明
'              0 表全部合計, 1 表正商標, 2 表聯合商標, 3 防護商標
'              4 服務標章, 5 聯合服務標章, 6 防護服務標章, 7 證明標章
'              8 團體標章, 9 團體商標
'         1 表取得國籍性的資料
'           nIndex = 如下說明
'              0 表全國合計, 1 表國內合計, 2 表大陸合計, 3 表國外合計
Private Function GetTaieAmount(ByVal nType As Integer, ByVal nIndex As Integer) As Variant
   Dim nAgentCount As Variant
   Dim nAmount As Variant
   Dim bFind As Boolean
   
   nAmount = 0
   bFind = False
   For nAgentCount = 0 To m_AgentCount - 1
      'If m_AgentList(nAgentCount).AgentCode = "001" Then
      If m_AgentList(nAgentCount).AgentName = "林晉章" Then
         bFind = True
         Exit For
      End If
   Next nAgentCount
   
   If bFind = False Then
      GetTaieAmount = 0
   Else
      Select Case nType
         Case 0
            Select Case nIndex
               Case 0: nAmount = m_AgentList(nAgentCount).KindAmount(0) + m_AgentList(nAgentCount).KindAmount(1) + m_AgentList(nAgentCount).KindAmount(2) + m_AgentList(nAgentCount).KindAmount(3) + m_AgentList(nAgentCount).KindAmount(4) + m_AgentList(nAgentCount).KindAmount(5) + m_AgentList(nAgentCount).KindAmount(6) + m_AgentList(nAgentCount).KindAmount(7)
               'Modify By Sindy 2010/4/27
               'Case 1, 2, 3, 4, 5, 6, 7, 8:
               Case 1, 2, 3, 4, 5, 6, 7, 8, 9:
                  nAmount = m_AgentList(nAgentCount).KindAmount(nIndex - 1)
            End Select
         Case 1
            Select Case nIndex
               Case 0: nAmount = m_AgentList(nAgentCount).ZoneAmount(0) + m_AgentList(nAgentCount).ZoneAmount(1) + m_AgentList(nAgentCount).ZoneAmount(2)
               Case 1, 2, 3:
                  nAmount = m_AgentList(nAgentCount).ZoneAmount(nIndex - 1)
            End Select
      End Select
   End If
   GetTaieAmount = nAmount
End Function

' 設定報表欄位的左方位置及其名稱
Public Sub BuildField_RP(ByVal nReport As Integer)
   Dim nIndex As Integer
   Dim nFieldWidth As Integer
   
   Select Case m_PaperSize
      Case "A4"
         m_LeftMargin = 1
         m_TopMargin = 2
         m_ReportWidth = 154
         m_ReportDataRows = 27
         nFieldWidth = 7
      Case "REPORT"
         m_LeftMargin = 1
         m_TopMargin = 3
         m_ReportWidth = 154
         m_ReportDataRows = 45
         nFieldWidth = 8
      Case Else
         m_LeftMargin = 10
         m_TopMargin = 3
         m_ReportWidth = 120
         m_ReportDataRows = 27
         nFieldWidth = 7
   End Select
   
   For nIndex = 0 To 16
      m_Field(nIndex).Width = nFieldWidth - 1
      'edit by nick 2004/12/13
      'm_Field(nIndex).Left = m_LeftMargin + (nIndex * nFieldWidth)
      m_Field(nIndex).Left = m_LeftMargin + (nIndex * nFieldWidth) + 8
      Select Case nIndex
         Case 0:
            'add by nick 2004/12/13
            m_Field(nIndex).Width = nFieldWidth + 5
            m_Field(nIndex).Left = m_LeftMargin + (nIndex * nFieldWidth)
            m_Field(nIndex).Name = "排名"
         Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14:
            m_Field(nIndex).Name = CStr(nIndex)
         Case 15:
            Select Case nReport
               Case 1:
                  m_Field(nIndex).Name = "全國"
               Case 2:
                  m_Field(nIndex).Name = CStr(nIndex)
            End Select
         Case 16:
            Select Case nReport
               Case 1:
                  m_Field(nIndex).Name = "台一%"
               Case 2:
                  m_Field(nIndex).Name = CStr(nIndex)
            End Select
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
Public Sub PrintPageHeader_RP(ByVal nReport As Integer, ByVal nPage As Integer, ByVal nStartRow As Integer)
   Dim nCount As Variant
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
    'Add By Cheng 2003/09/10
    '若列印表一, 表頭加受文者
    'Begin
    If nReport = 1 Then
        Printer.FontSize = 12
        Printer.CurrentX = m_LeftMargin * m_CharWidth
        Printer.CurrentY = (m_TopMargin + nStartRow + nRow) * m_CharHeight
        Printer.Print "受文者：北所、中所、南所、高所"
    End If
    'End
   Printer.CurrentY = (m_TopMargin + nStartRow + nRow) * m_CharHeight
   Printer.FontSize = 24
   Printer.Font.Underline = True
   Select Case nReport
      Case 1:
         'edit by nick 2004/12/13
         'nX = m_LeftMargin + m_ReportWidth / 2 - 24
         nX = m_LeftMargin + m_ReportWidth / 2 - 28
         Printer.CurrentX = nX * m_CharWidth
         Printer.Print "表一：商標全國市場統計表"
      Case 2:
         'edit by nick 2004/12/13
         'nX = m_LeftMargin + m_ReportWidth / 2 - 24
         nX = m_LeftMargin + m_ReportWidth / 2 - 28
         Printer.CurrentX = nX * m_CharWidth
         Printer.Print "表二：台一各區市場統計表"
      'add by nickc 2007/04/19
      Case 3:
         nX = m_LeftMargin + m_ReportWidth / 2 - 28
         Printer.CurrentX = nX * m_CharWidth
         Printer.Print "本所申請但非本所繳費案件"
   End Select
   Printer.Font.Underline = False
   
   nRow = nRow + 2
   Printer.FontSize = 12
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nStartRow + nRow) * m_CharHeight
   Printer.Print "列印人:" & strUserName
   
   nX = m_LeftMargin + m_ReportWidth / 2 - 16
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nStartRow + nRow) * m_CharHeight
   Printer.FontSize = 12
   Printer.Print "公報卷期:" & strData1 & " - " & strData2
   
   'edit by nick 2004/12/13
   'nX = m_LeftMargin + m_ReportWidth - 20
   nX = m_LeftMargin + m_ReportWidth - 38
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nStartRow + nRow) * m_CharHeight
   Printer.Print "製表日期:" & Format(ChangeWStringToWDateString(GetTodayDate), "EE/MM/DD")
   
   nRow = nRow + 1
   'edit by nick 2004/12/13
   'nX = m_LeftMargin + m_ReportWidth - 20
   nX = m_LeftMargin + m_ReportWidth - 38
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nStartRow + nRow) * m_CharHeight
   Printer.Print "頁"
   
   'edit by nick 2004/12/13
   'nX = m_LeftMargin + m_ReportWidth - 14
   nX = m_LeftMargin + m_ReportWidth - 32
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nStartRow + nRow) * m_CharHeight
   Printer.Print "次:" & nPage
   
   ' 列印分隔線
   nRow = nRow + 1
   PrintSplitLine nStartRow + nRow
   
   'add by nickc 2007/04/19
   If nReport = 3 Then
        nRow = nRow + 1
           Printer.CurrentX = 120
           Printer.CurrentY = (m_TopMargin + nStartRow + nRow) * m_CharHeight
           Printer.Print "業務區"
           Printer.CurrentX = 1500
           Printer.CurrentY = (m_TopMargin + nStartRow + nRow) * m_CharHeight
           Printer.Print "智權人員"
           Printer.CurrentX = 3000
           Printer.CurrentY = (m_TopMargin + nStartRow + nRow) * m_CharHeight
           Printer.Print "本所案號"
           Printer.CurrentX = 5500
           Printer.CurrentY = (m_TopMargin + nStartRow + nRow) * m_CharHeight
           Printer.Print "商標名稱"
           Printer.CurrentX = 9000
           Printer.CurrentY = (m_TopMargin + nStartRow + nRow) * m_CharHeight
           Printer.Print "客戶"
   Else
        nRow = nRow + 1
        For nIndex = 0 To 16
           nCenter = ((m_Field(nIndex).Left * m_CharWidth) + (m_Field(nIndex).Left + m_Field(nIndex).Width) * m_CharWidth) / 2
           strTemp = LeftStr(m_Field(nIndex).Name, m_Field(nIndex).Width)
           Printer.CurrentX = nCenter - Printer.TextWidth(strTemp) / 2
           Printer.CurrentY = (m_TopMargin + nStartRow + nRow) * m_CharHeight
           Printer.Print strTemp
        Next nIndex
    End If
    
   ' 列印分隔線
   nRow = nRow + 1
   PrintSplitLine nStartRow + nRow
   
   m_HeaderHeight = nRow
End Sub

' 清除資料
Public Sub Clear()
   Dim nFieldCount As Variant
   Dim nAgentCount As Variant
   Dim nZoneCount As Variant
   Dim nIndex As Integer
   ' 清除欄位內容
   For nFieldCount = 0 To 16
      m_Field(nFieldCount).Name = Empty
      m_Field(nFieldCount).Left = 0
      m_Field(nFieldCount).Width = 0
   Next nFieldCount
   ' 清除地區串列
   If m_ZoneCount > 0 Then
      Erase m_ZoneList
   End If
   m_ZoneCount = 0
   ' 清除代理人串列
   If m_AgentCount > 0 Then
      For nAgentCount = 0 To m_AgentCount - 1
         Erase m_AgentList(nAgentCount).KindAmount
         Erase m_AgentList(nAgentCount).ZoneAmount
         'add by nick 2004/12/13
         Erase m_AgentList(nAgentCount).KindAmount08
         Erase m_AgentList(nAgentCount).ZoneAmount08
      Next nAgentCount
      Erase m_AgentList
   End If
   m_AgentCount = 0
   
   m_NoAgentItem.Count = 0
   'add by nick 2004/12/13
   m_NoAgentItem.Count08 = 0
   For nIndex = 0 To 8
      m_NoAgentItem.KindAmount(nIndex) = 0
      'add by nick 2004/12/13
      m_NoAgentItem.KindAmount08(nIndex) = 0
   Next nIndex
   For nIndex = 0 To 3
      m_NoAgentItem.ZoneAmount(nIndex) = 0
      'add by nick 2004/12/13
      m_NoAgentItem.ZoneAmount08(nIndex) = 0
   Next nIndex
   m_NoAgentItem.AgentCode = Empty
   m_NoAgentItem.AgentName = Empty
   m_NoAgentItem.Company = Empty
   
End Sub

' 從資料庫中取得所有的資料
Private Function GetDBData_RP() As Boolean
   Dim rsMain As New ADODB.Recordset
   Dim strSql As String
   Dim strZone, strZoneCode, strAgent, strAgentCode, strAgentCompany As String
   Dim nZoneIndex, nAgentIndex, nCount As Variant
   Dim bFindZone, bFindAgent As Boolean
   Dim nType As Integer
   Dim nZone As Integer
   Dim nSortX, nSortY As Integer
   Dim AgentTemp As AGENTITEM
   Dim ZoneTemp As ZONEITEM
   Dim bFromSec As Boolean
   Dim bToSec As Boolean
   
   GetDBData_RP = True
   
   ' 產生SQL查詢語法
   bFromSec = Not IsEmptyText(textTMBM07_1.Text)
   bToSec = Not IsEmptyText(textTMBM07_2.Text)
   'Modify By Sindy 2013/8/19 + length(na01)=3 AND
   If bFromSec = True And bToSec = True Then
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,NA01,NA02,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04 FROM TMBULLETIN, NATION, TAGENT " & _
               "WHERE TMBM07 >= '" & textTMBM07_1 & "' AND " & _
                     "TMBM07 <= '" & textTMBM07_2 & "' AND " & _
                     "TMBM05 = NA03(+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01(+) "
   ElseIf bFromSec = True And bToSec = False Then
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,NA01,NA02,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04 FROM TMBULLETIN, NATION, TAGENT " & _
               "WHERE TMBM07 >= '" & textTMBM07_1 & "' AND " & _
                     "TMBM05 = NA03(+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01(+) "
   ElseIf bFromSec = False And bToSec = True Then
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,NA01,NA02,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04 FROM TMBULLETIN, NATION, TAGENT " & _
               "WHERE TMBM07 <= '" & textTMBM07_2 & "' AND " & _
                     "TMBM05 = NA03(+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01(+) "
   Else
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,NA01,NA02,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04 FROM TMBULLETIN, NATION, TAGENT " & _
               "WHERE TMBM05 = NA03(+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01(+) "
   End If
   
   ' 取得資料庫的資料
   rsMain.CursorLocation = adUseClient
   rsMain.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   ' 無資料則離開
   If rsMain.RecordCount <= 0 Then
      GetDBData_RP = False
      GoTo EXITSUB
   End If

   ' 設定初始值
   m_ZoneCount = 0
   m_AgentCount = 0

   rsMain.MoveFirst
   ' 依序從資料記錄中取出欄位的內容
   While Not rsMain.EOF
      ' 代理人姓名
      strAgent = Empty
      If IsNull(rsMain.Fields("TMBM06")) = False Then
         strAgent = rsMain.Fields("TMBM06")
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
      ' 地區名稱
      strZone = Empty
      If IsNull(rsMain.Fields("TMBM05")) = False Then
         strZone = rsMain.Fields("TMBM05")
      End If
      ' 地區別
      nZone = 0
      ' 地區代碼
      strZoneCode = Empty
      'If IsNull(rsMain.Fields("NA01")) = False Then
      '   strZoneCode = rsMain.Fields("NA01")
      'End If
      'If IsNull(rsMain.Fields("NA02")) = False Then
      '   strZoneCode = rsMain.Fields("NA02")
      'End If
      
      If IsNull(rsMain.Fields("NA02")) = False Then
         Select Case Mid(rsMain.Fields("NA02"), 1, 1)
            Case "A":
               nZone = 1
               Select Case Mid(rsMain.Fields("NA02"), 1, 2)
                  Case "A1":
                     strZone = "北區"
                     strZoneCode = "A10"
                  Case "A2":
                     strZone = "中區"
                     strZoneCode = "A20"
                  Case "A3":
                     strZone = "南區"
                     strZoneCode = "A30"
                  Case "A4":
                     strZone = "高區"
                     strZoneCode = "A40"
                  Case "A5":
                     strZone = "東區"
                     strZoneCode = "A50"
               End Select
            Case "B":
               nZone = 2
               strZoneCode = "B00"
               strZone = "大陸"
            Case Else:
               nZone = 3
               strZoneCode = "C00"
               strZone = "國外"
         End Select
      Else
         nZone = 3
         strZoneCode = "C00"
         strZone = "國外"
      End If
      
      ' 檢查商標種類
      nType = 1
      If IsNull(rsMain.Fields("TMBM02")) = False Then
         Select Case rsMain.Fields("TMBM02")
            ' 正商標
            Case "1": nType = 1
            ' 聯合商標
            Case "2": nType = 2
            ' 防護商標
            Case "3": nType = 3
            ' 服務標章
            Case "4": nType = 4
            ' 聯合服務標章
            Case "5": nType = 5
            ' 防護服務標章
            Case "6": nType = 6
            ' 證明標章
            Case "7": nType = 7
            ' 團體標章
            Case "8": nType = 8
            'Add By Sindy 2010/4/27
            ' 團體商標
            Case "9": nType = 9
         End Select
      End If
      
      ' 地區串列
      bFindZone = False
      For nZoneIndex = 0 To m_ZoneCount - 1
         If m_ZoneList(nZoneIndex).ZoneCode = strZoneCode Then
            bFindZone = True
            m_ZoneList(nZoneIndex).Count = m_ZoneList(nZoneIndex).Count + 1
            'If IsTaie(strAgentCode) = True Then
            If IsTaie(strAgent) = True Then
               m_ZoneList(nZoneIndex).TaieCount = m_ZoneList(nZoneIndex).TaieCount + 1
            End If
            Exit For
         End If
      Next nZoneIndex
      If bFindZone = False Then
         If m_ZoneCount = 0 Then
            nZoneIndex = 0
         Else
            nZoneIndex = UBound(m_ZoneList)
         End If
         ReDim Preserve m_ZoneList(nZoneIndex + 1)
         m_ZoneCount = m_ZoneCount + 1
         m_ZoneList(nZoneIndex).ZoneCode = strZoneCode
         m_ZoneList(nZoneIndex).ZoneName = strZone
         m_ZoneList(nZoneIndex).Count = 1
         m_ZoneList(nZoneIndex).TaieCount = 0
         'If IsTaie(strAgentCode) = True Then
         If IsTaie(strAgent) = True Then
            m_ZoneList(nZoneIndex).TaieCount = 1
         End If
      End If
      
      ' 代理人串列
      If IsEmptyText(strAgent) = True Then
         m_NoAgentItem.Count = m_NoAgentItem.Count + 1
         ' 商標種類
         Select Case nType
            Case 1: m_NoAgentItem.KindAmount(0) = m_NoAgentItem.KindAmount(0) + 1
            Case 2: m_NoAgentItem.KindAmount(1) = m_NoAgentItem.KindAmount(1) + 1
            Case 3: m_NoAgentItem.KindAmount(2) = m_NoAgentItem.KindAmount(2) + 1
            Case 4: m_NoAgentItem.KindAmount(3) = m_NoAgentItem.KindAmount(3) + 1
            Case 5: m_NoAgentItem.KindAmount(4) = m_NoAgentItem.KindAmount(4) + 1
            Case 6: m_NoAgentItem.KindAmount(5) = m_NoAgentItem.KindAmount(5) + 1
            Case 7: m_NoAgentItem.KindAmount(6) = m_NoAgentItem.KindAmount(6) + 1
            Case 8: m_NoAgentItem.KindAmount(7) = m_NoAgentItem.KindAmount(7) + 1
            'Add By Sindy 2010/4/27
            Case 9: m_NoAgentItem.KindAmount(8) = m_NoAgentItem.KindAmount(8) + 1
         End Select
         ' 地區
         Select Case nZone
            Case 1: m_NoAgentItem.ZoneAmount(0) = m_NoAgentItem.ZoneAmount(0) + 1
            Case 2: m_NoAgentItem.ZoneAmount(1) = m_NoAgentItem.ZoneAmount(1) + 1
            Case 3: m_NoAgentItem.ZoneAmount(2) = m_NoAgentItem.ZoneAmount(2) + 1
         End Select
      Else
         ' 搜尋代理人串列
         bFindAgent = False
         For nAgentIndex = 0 To m_AgentCount - 1
            'Modify By Sindy 2010/02/26
            'If m_AgentList(nAgentIndex).AgentName = strAgent Then
            If m_AgentList(nAgentIndex).Company = strAgentCompany Then
            '2010/02/26 End
               bFindAgent = True
               m_AgentList(nAgentIndex).Count = m_AgentList(nAgentIndex).Count + 1
               ' 商標種類
               Select Case nType
                  Case 1: m_AgentList(nAgentIndex).KindAmount(0) = m_AgentList(nAgentIndex).KindAmount(0) + 1
                  Case 2: m_AgentList(nAgentIndex).KindAmount(1) = m_AgentList(nAgentIndex).KindAmount(1) + 1
                  Case 3: m_AgentList(nAgentIndex).KindAmount(2) = m_AgentList(nAgentIndex).KindAmount(2) + 1
                  Case 4: m_AgentList(nAgentIndex).KindAmount(3) = m_AgentList(nAgentIndex).KindAmount(3) + 1
                  Case 5: m_AgentList(nAgentIndex).KindAmount(4) = m_AgentList(nAgentIndex).KindAmount(4) + 1
                  Case 6: m_AgentList(nAgentIndex).KindAmount(5) = m_AgentList(nAgentIndex).KindAmount(5) + 1
                  Case 7: m_AgentList(nAgentIndex).KindAmount(6) = m_AgentList(nAgentIndex).KindAmount(6) + 1
                  Case 8: m_AgentList(nAgentIndex).KindAmount(7) = m_AgentList(nAgentIndex).KindAmount(7) + 1
                  'Add By Sindy 2010/4/27
                  Case 9: m_AgentList(nAgentIndex).KindAmount(8) = m_AgentList(nAgentIndex).KindAmount(8) + 1
               End Select
               ' 地區
               Select Case nZone
                  Case 1: m_AgentList(nAgentIndex).ZoneAmount(0) = m_AgentList(nAgentIndex).ZoneAmount(0) + 1
                  Case 2: m_AgentList(nAgentIndex).ZoneAmount(1) = m_AgentList(nAgentIndex).ZoneAmount(1) + 1
                  Case 3: m_AgentList(nAgentIndex).ZoneAmount(2) = m_AgentList(nAgentIndex).ZoneAmount(2) + 1
               End Select
               Exit For
            End If
         Next nAgentIndex
         If bFindAgent = False Then
            If m_AgentCount = 0 Then
               nAgentIndex = 0
            Else
               nAgentIndex = UBound(m_AgentList)
            End If
            ReDim Preserve m_AgentList(nAgentIndex + 1)
            m_AgentCount = m_AgentCount + 1
            m_AgentList(nAgentIndex).AgentName = strAgent
            m_AgentList(nAgentIndex).Company = strAgentCompany
            m_AgentList(nAgentIndex).AgentCode = strAgentCode
            m_AgentList(nAgentIndex).Count = 1
            ' 商標種類
            Select Case nType
               Case 1: m_AgentList(nAgentIndex).KindAmount(0) = m_AgentList(nAgentIndex).KindAmount(0) + 1
               Case 2: m_AgentList(nAgentIndex).KindAmount(1) = m_AgentList(nAgentIndex).KindAmount(1) + 1
               Case 3: m_AgentList(nAgentIndex).KindAmount(2) = m_AgentList(nAgentIndex).KindAmount(2) + 1
               Case 4: m_AgentList(nAgentIndex).KindAmount(3) = m_AgentList(nAgentIndex).KindAmount(3) + 1
               Case 5: m_AgentList(nAgentIndex).KindAmount(4) = m_AgentList(nAgentIndex).KindAmount(4) + 1
               Case 6: m_AgentList(nAgentIndex).KindAmount(5) = m_AgentList(nAgentIndex).KindAmount(5) + 1
               Case 7: m_AgentList(nAgentIndex).KindAmount(6) = m_AgentList(nAgentIndex).KindAmount(6) + 1
               Case 8: m_AgentList(nAgentIndex).KindAmount(7) = m_AgentList(nAgentIndex).KindAmount(7) + 1
               'Add By Sindy 2010/4/27
               Case 9: m_AgentList(nAgentIndex).KindAmount(8) = m_AgentList(nAgentIndex).KindAmount(8) + 1
            End Select
            ' 地區
            Select Case nZone
               Case 1: m_AgentList(nAgentIndex).ZoneAmount(0) = m_AgentList(nAgentIndex).ZoneAmount(0) + 1
               Case 2: m_AgentList(nAgentIndex).ZoneAmount(1) = m_AgentList(nAgentIndex).ZoneAmount(1) + 1
               Case 3: m_AgentList(nAgentIndex).ZoneAmount(2) = m_AgentList(nAgentIndex).ZoneAmount(2) + 1
            End Select
         End If
      End If
      ' 移到下一筆記錄
      rsMain.MoveNext
   Wend
   
   ' 對代理人串列依數量的多寡由大到小排序
   For nSortX = 0 To m_AgentCount - 1
      For nSortY = nSortX To m_AgentCount - 1
         If m_AgentList(nSortX).Count < m_AgentList(nSortY).Count Then
            AgentTemp = m_AgentList(nSortX)
            m_AgentList(nSortX) = m_AgentList(nSortY)
            m_AgentList(nSortY) = AgentTemp
         ElseIf m_AgentList(nSortX).Count = m_AgentList(nSortY).Count Then
            If m_AgentList(nSortX).Company > m_AgentList(nSortY).Company Then
               AgentTemp = m_AgentList(nSortX)
               m_AgentList(nSortX) = m_AgentList(nSortY)
               m_AgentList(nSortY) = AgentTemp
            End If
         End If
      Next nSortY
   Next nSortX
   
   ' 地區串列依台一在該地區的數量由大到小排序
   For nSortX = 0 To m_ZoneCount - 1
      For nSortY = nSortX To m_ZoneCount - 1
         If (m_ZoneList(nSortX).TaieCount / m_ZoneList(nSortX).Count) < (m_ZoneList(nSortY).TaieCount / m_ZoneList(nSortY).Count) Then
            ZoneTemp = m_ZoneList(nSortX)
            m_ZoneList(nSortX) = m_ZoneList(nSortY)
            m_ZoneList(nSortY) = ZoneTemp
         End If
      Next nSortY
   Next nSortX
   
EXITSUB:
   rsMain.Close
   Set rsMain = Nothing
End Function

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
   Dim strZone As String
   Dim nRow As Integer
   Dim nPage As Integer
   Dim fld(17) As String
   Dim nAgentCount As Variant
   Dim bFindAgent As Boolean
   Dim nType As Integer
   Dim nAmount As Variant
   Dim nTaieAmount As Variant
   Dim nNoAgentAmount As Variant
   Dim nTotalAmount As Variant
   Dim nZoneCount As Variant
   Dim nCount As Variant
   Dim fValue As Variant
   Dim nFinalAmount As Variant
   Dim nNoAgent(3) As Variant
   Dim nX As Integer
   Dim nRight As Long
   Dim nCenter As Long
   Dim strTemp As String
   Dim nSrcHeaderHeight As Integer
   
   'Printer.Copies = 5
   nAmount = 0
   nTaieAmount = 0
   nNoAgentAmount = 0
   nTotalAmount = 0
   nAgentCount = 0
   nZoneCount = 0
   nCount = 0
   fValue = 0#
   nNoAgent(0) = 0
   nNoAgent(1) = 0
   nNoAgent(2) = 0
   
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
   PrintPageHeader_RP 1, 1, 0
   
   ' 依序列印表一的第一部份
   nRow = 1
   'Modify By Sindy 2010/4/27
   'For nCount = 0 To 8
   For nCount = 0 To 9
      ' 清除內容
      For nX = 0 To 16: fld(nX) = Empty: Next nX
      ' 第 0 欄位
      Select Case nCount
         Case 0: fld(0) = "事務所"
         Case 1: fld(0) = "正商標"
         Case 2: fld(0) = "聯合商標"
         Case 3: fld(0) = "防護商標"
         Case 4: fld(0) = "服務標章"
         Case 5: fld(0) = "聯合服務標章"
         Case 6: fld(0) = "防護服務標章"
         Case 7: fld(0) = "證明標章"
         Case 8: fld(0) = "團體標章"
         'Add By Sindy 2010/4/27
         Case 9: fld(0) = "團體商標"
      End Select
      ' 第 1 ~ 14 欄位
      For nAgentCount = 0 To Min(14, m_AgentCount - 1)
         Select Case nCount
            Case 0:
               'fld(nAgentCount + 1) = m_AgentList(nAgentCount).AgentName
               fld(nAgentCount + 1) = m_AgentList(nAgentCount).Company
            'Modify By Sindy 2010/4/27
            'Case 1, 2, 3, 4, 5, 6, 7, 8:
            Case 1, 2, 3, 4, 5, 6, 7, 8, 9:
               fld(nAgentCount + 1) = GetAgentAmount(m_AgentList(nAgentCount), 0, nCount)
         End Select
      Next nAgentCount
      ' 第 15, 16 欄位
      Select Case nCount
         Case 0:
            fld(15) = Empty
            fld(16) = Empty
         'Modify By Sindy 2010/4/27
         'Case 1, 2, 3, 4, 5, 6, 7, 8:
         Case 1, 2, 3, 4, 5, 6, 7, 8, 9:
            nAmount = GetTaieAmount(0, nCount)
            nTotalAmount = GetTotalAmount(0, nCount)
            fld(15) = nTotalAmount
            If nTotalAmount > 0 Then
               fValue = nAmount / nTotalAmount * 100
            Else
               fValue = 0#
            End If
            fld(16) = Format(fValue, "##0.00")
            'fld(16) = fld(16) & " %"
      End Select
      ' 列印
      For nAgentCount = 0 To 16
         If nAgentCount = 0 Then
            Printer.CurrentX = m_Field(nAgentCount).Left * m_CharWidth
            Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
            Printer.Print fld(nAgentCount)
         Else
            Select Case nCount
               Case 0
                  Printer.FontSize = 8
                  nCenter = ((m_Field(nAgentCount).Left * m_CharWidth) + (m_Field(nAgentCount).Left + m_Field(nAgentCount).Width) * m_CharWidth) / 2
                  Printer.CurrentX = nCenter - Printer.TextWidth(fld(nAgentCount)) / 2
                  Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
                  Printer.Print fld(nAgentCount)
                  Printer.FontSize = 12
               Case Else
                  nRight = (m_Field(nAgentCount).Left + m_Field(nAgentCount).Width - 2) * m_CharWidth
                  Printer.CurrentX = nRight - Printer.TextWidth(fld(nAgentCount))
                  Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
                  Printer.Print fld(nAgentCount)
            End Select
         End If
      Next nAgentCount
      nRow = nRow + 1
   Next nCount
   ' 列印分隔線
   PrintSplitLine m_HeaderHeight + nRow
   ' 國別部份
   nRow = nRow + 1
   For nCount = 0 To 3
      ' 清除內容
      For nX = 0 To 16: fld(nX) = Empty: Next nX
      ' 第 0 欄位
      Select Case nCount
         Case 0: fld(0) = "全國合計"
         Case 1: fld(0) = "國內合計"
         Case 2: fld(0) = "大陸合計"
         Case 3: fld(0) = "國外合計"
      End Select
      ' 第 1 ~ 14 欄位
      For nAgentCount = 0 To Min(14, m_AgentCount - 1)
         fld(nAgentCount + 1) = GetAgentAmount(m_AgentList(nAgentCount), 1, nCount)
      Next nAgentCount
      ' 第 15, 16 欄位
      nAmount = 0
      nAmount = GetTaieAmount(1, nCount)
      nTotalAmount = 0
      nTotalAmount = GetTotalAmount(1, nCount)
      fld(15) = nTotalAmount
      If nTotalAmount > 0 Then
         fValue = nAmount / nTotalAmount * 100
      Else
         fValue = 0#
      End If
      fld(16) = Format(fValue, "##0.00")
      
      ' 列印
      For nAgentCount = 0 To 16
         If nAgentCount = 0 Then
            Printer.CurrentX = m_Field(nAgentCount).Left * m_CharWidth
            Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
            Printer.Print fld(nAgentCount)
         Else
            nRight = (m_Field(nAgentCount).Left + m_Field(nAgentCount).Width - 2) * m_CharWidth
            Printer.CurrentX = nRight - Printer.TextWidth(fld(nAgentCount))
            Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
            Printer.Print fld(nAgentCount)
         End If
      Next nAgentCount
      nRow = nRow + 2
   Next nCount
   ' 列印分隔線
   PrintSplitLine m_HeaderHeight + nRow
   
   ' 代理人比例
   ' 清除內容
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   nRow = nRow + 2
   nTotalAmount = GetTotalAmount(0, 0, False)
   fld(0) = "代理人比例"
   For nAgentCount = 0 To Min(14, m_AgentCount - 1)
      nAmount = GetAgentAmount(m_AgentList(nAgentCount), 0, 0)
      fValue = 0
      If nTotalAmount > 0 Then
         fValue = nAmount / nTotalAmount * 100
      End If
      fld(nAgentCount + 1) = Format(fValue, "##0.00")
   Next nAgentCount
   fld(15) = GetTotalAmount(1, 0, False)
   fValue = GetTotalAmount(1, 0, False) / GetTotalAmount(1, 0, True) * 100
   fld(16) = Format(fValue, "##0.00")
   ' 列印
   For nAgentCount = 0 To 16
      If nAgentCount = 0 Then
         Printer.CurrentX = m_Field(nAgentCount).Left * m_CharWidth
         Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nAgentCount)
      Else
         nRight = (m_Field(nAgentCount).Left + m_Field(nAgentCount).Width - 2) * m_CharWidth
         Printer.CurrentX = nRight - Printer.TextWidth(fld(nAgentCount))
         Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nAgentCount)
      End If
   Next nAgentCount
   
   ' 全國比例
   ' 清除內容
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   nRow = nRow + 2
   nTotalAmount = GetTotalAmount(0, 0, True)
   fld(0) = "全國比例"
   For nAgentCount = 0 To Min(14, m_AgentCount - 1)
      nAmount = GetAgentAmount(m_AgentList(nAgentCount), 0, 0)
      fValue = 0
      If nTotalAmount > 0 Then
         fValue = nAmount / nTotalAmount * 100
      End If
      fld(nAgentCount + 1) = Format(fValue, "##0.00")
   Next nAgentCount
   fld(15) = GetTotalAmount(1, 0, True)
   ' 列印
   For nAgentCount = 0 To 16
      If nAgentCount = 0 Then
         Printer.CurrentX = m_Field(nAgentCount).Left * m_CharWidth
         Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nAgentCount)
      Else
         nRight = (m_Field(nAgentCount).Left + m_Field(nAgentCount).Width - 2) * m_CharWidth
         Printer.CurrentX = nRight - Printer.TextWidth(fld(nAgentCount))
         Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nAgentCount)
      End If
   Next nAgentCount
   
   ' 列印雙分隔線
   nRow = nRow + 1
   PrintTerminateLine m_HeaderHeight + nRow
   ' 記錄原Header的高度
   nSrcHeaderHeight = m_HeaderHeight
   
   ' 表一與表二的間距
   nRow = nRow + 5
   
   ' 列印表二的表頭
   BuildField_RP 2
   PrintPageHeader_RP 2, 1, nSrcHeaderHeight + nRow
   nRow = nRow + m_HeaderHeight
   
   ' 地區
   ' 清除內容
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   nRow = nRow + 2
   fld(0) = "地區名稱"
   For nZoneCount = 0 To Min(16, m_ZoneCount - 1)
      fld(nZoneCount + 1) = m_ZoneList(nZoneCount).ZoneName
   Next nZoneCount
   For nCount = 0 To 16
      If nCount = 0 Then
         Printer.CurrentX = m_Field(nCount).Left * m_CharWidth
         Printer.CurrentY = (nSrcHeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nCount)
      Else
         Printer.FontSize = 8
         nCenter = ((m_Field(nCount).Left * m_CharWidth) + (m_Field(nCount).Left + m_Field(nCount).Width) * m_CharWidth) / 2
         Printer.CurrentX = nCenter - Printer.TextWidth(fld(nCount)) / 2
         Printer.CurrentY = (nSrcHeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nCount)
         Printer.FontSize = 12
      End If
   Next nCount
   
   ' 台一合計
   ' 清除內容
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   nRow = nRow + 2
   fld(0) = "台一合計"
   For nZoneCount = 0 To Min(16, m_ZoneCount - 1)
      nAmount = m_ZoneList(nZoneCount).TaieCount
      fld(nZoneCount + 1) = nAmount
   Next nZoneCount
   For nCount = 0 To 16
      If nCount = 0 Then
         Printer.CurrentX = m_Field(nCount).Left * m_CharWidth
         Printer.CurrentY = (nSrcHeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nCount)
      Else
         nRight = (m_Field(nCount).Left + m_Field(nCount).Width - 2) * m_CharWidth
         Printer.CurrentX = nRight - Printer.TextWidth(fld(nCount))
         Printer.CurrentY = (nSrcHeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nCount)
      End If
   Next nCount
   
   ' 區域合計
   ' 清除內容
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   nRow = nRow + 2
   fld(0) = "區域合計"
   For nZoneCount = 0 To Min(16, m_ZoneCount - 1)
      nAmount = m_ZoneList(nZoneCount).Count
      fld(nZoneCount + 1) = nAmount
   Next nZoneCount
   For nCount = 0 To 14
      If nCount = 0 Then
         Printer.CurrentX = m_Field(nCount).Left * m_CharWidth
         Printer.CurrentY = (nSrcHeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nCount)
      Else
         nRight = (m_Field(nCount).Left + m_Field(nCount).Width - 2) * m_CharWidth
         Printer.CurrentX = nRight - Printer.TextWidth(fld(nCount))
         Printer.CurrentY = (nSrcHeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nCount)
      End If
   Next nCount
   
   ' 百分比
   ' 清除內容
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   nRow = nRow + 2
   fld(0) = "百分比"
   For nZoneCount = 0 To Min(16, m_ZoneCount - 1)
      nAmount = m_ZoneList(nZoneCount).TaieCount
      nTotalAmount = m_ZoneList(nZoneCount).Count
      If nTotalAmount > 0 Then
         fValue = nAmount / nTotalAmount * 100
      Else
         fValue = 0#
      End If
      fld(nZoneCount + 1) = Format(fValue, "##0.00")
      'fld(nZoneCount + 1) = fld(nZoneCount + 1) & " %"
   Next nZoneCount
   For nCount = 0 To 16
      If nCount = 0 Then
         Printer.CurrentX = m_Field(nCount).Left * m_CharWidth
         Printer.CurrentY = (nSrcHeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nCount)
      Else
         nRight = (m_Field(nCount).Left + m_Field(nCount).Width - 2) * m_CharWidth
         Printer.CurrentX = nRight - Printer.TextWidth(fld(nCount))
         Printer.CurrentY = (nSrcHeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nCount)
      End If
   Next nCount
   
   ' 列印雙分隔線
   nRow = nRow + 2
   PrintTerminateLine m_HeaderHeight + nRow
   
   Printer.EndDoc
   
End Sub
' 背景列印
Public Sub PrintReportBK(ByVal strPrinter As String, ByVal TMBM07_1 As String, ByVal TMBM07_2 As String)
   Dim Prn As Printer
    'Add By Cheng 2003/01/27
    'Dim nCopys As Integer '列印份數
   
   Me.Hide
   textTMBM07_1 = TMBM07_1
   textTMBM07_2 = TMBM07_2
   
   '搜尋 Printer
   'edit by nick 2004/12/14
'   For Each Prn In Printers
'      If Prn.DeviceName = strPrinter Then
'         Set Printer = Prn
'         Exit For
'      End If
'   Next
   
   ' 建立欄位資訊
   BuildField_RP 1
   ' 取得資料庫中的資料
   'edit by nick 2004/12/147
   'If GetDBData_RP = False Then
   If GetDBData_RP_931213 = False Then
      GoTo EXITSUB
   End If
   ' 列印
    'Modify By Cheng 2003/01/27
    '列印五份
'   Generate_RP
    txt1.Text = 1 '5
    For nCopys = 1 To Trim(txt1.Text) 'Modify By Sindy 2022/3/23 桂英取消列印5份,1份PDF即可
       ' 列印
       'edit by nick 2004/12/14
       'Generate_RP
       Generate_RP_931213
       ' 測試(輸出到畫面上)
       'Generate_RP_SCREEN
    Next nCopys
   ' 清除所佔用的空間
   Clear
EXITSUB:
   Set frm030606 = Nothing
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

'Add By Sindy 2010/01/12
Private Sub txt1_GotFocus()
   InverseTextBox txt1
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

' 列印表一的內容
Public Sub Generate_RP_SCREEN()
   Dim strZone As String
   Dim nRow As Integer
   Dim nPage As Integer
   Dim fld(17) As String
   Dim nAgentCount As Variant
   Dim bFindAgent As Boolean
   Dim nType As Integer
   Dim nAmount As Variant
   Dim nTaieAmount As Variant
   Dim nNoAgentAmount As Variant
   Dim nTotalAmount As Variant
   Dim nZoneCount As Variant
   Dim nCount As Variant
   Dim fValue As Variant
   Dim nFinalAmount As Variant
   Dim nNoAgent(3) As Variant
   Dim nX As Integer
   Dim nRight As Long
   Dim nCenter As Long
   Dim strTemp As String
   Dim nSrcHeaderHeight As Integer
   
   ' 依序列印表一的第一部份
   nRow = 1
   'Modify By Sindy 2010/4/27
   'For nCount = 0 To 8
   For nCount = 0 To 9
      ' 清除內容
      For nX = 0 To 16: fld(nX) = Empty: Next nX
      ' 第 0 欄位
      Select Case nCount
         Case 0: fld(0) = "事務所"
         Case 1: fld(0) = "正商標"
         Case 2: fld(0) = "聯合商標"
         Case 3: fld(0) = "防護商標"
         Case 4: fld(0) = "服務標章"
         Case 5: fld(0) = "聯合服務標章"
         Case 6: fld(0) = "防護服務標章"
         Case 7: fld(0) = "證明標章"
         Case 8: fld(0) = "團體標章"
         'Add By Sindy 2010/4/27
         Case 9: fld(0) = "團體商標"
      End Select
      ' 第 1 ~ 14 欄位
      For nAgentCount = 0 To Min(14, m_AgentCount - 1)
         Select Case nCount
            Case 0:
               'fld(nAgentCount + 1) = m_AgentList(nAgentCount).AgentName
               fld(nAgentCount + 1) = m_AgentList(nAgentCount).Company
            'Modify By Sindy 2010/4/27
            'Case 1, 2, 3, 4, 5, 6, 7, 8:
            Case 1, 2, 3, 4, 5, 6, 7, 8, 9:
               fld(nAgentCount + 1) = GetAgentAmount(m_AgentList(nAgentCount), 0, nCount)
         End Select
      Next nAgentCount
      ' 第 15, 16 欄位
      Select Case nCount
         Case 0:
            fld(15) = Empty
            fld(16) = Empty
         'Modify By Sindy 2010/4/27
         'Case 1, 2, 3, 4, 5, 6, 7, 8:
         Case 1, 2, 3, 4, 5, 6, 7, 8, 9:
            nAmount = GetTaieAmount(0, nCount)
            nTotalAmount = GetTotalAmount(0, nCount)
            fld(15) = nTotalAmount
            If nTotalAmount > 0 Then
               fValue = nAmount / nTotalAmount * 100
            Else
               fValue = 0#
            End If
            fld(16) = Format(fValue, "##0.00")
            fld(16) = fld(16) & " %"
      End Select
   Next nCount
   
   ' 國別部份
   For nCount = 0 To 3
      ' 清除內容
      For nX = 0 To 16: fld(nX) = Empty: Next nX
      ' 第 0 欄位
      Select Case nCount
         Case 0: fld(0) = "全國合計"
         Case 1: fld(0) = "國內合計"
         Case 2: fld(0) = "大陸合計"
         Case 3: fld(0) = "國外合計"
      End Select
      ' 第 1 ~ 14 欄位
      For nAgentCount = 0 To Min(14, m_AgentCount - 1)
         fld(nAgentCount + 1) = GetAgentAmount(m_AgentList(nAgentCount), 1, nCount)
      Next nAgentCount
      ' 第 15, 16 欄位
      nAmount = GetTaieAmount(1, nCount)
      nTotalAmount = GetTotalAmount(1, nCount)
      fld(15) = nTotalAmount
      If nTotalAmount > 0 Then
         fValue = nAmount / nTotalAmount * 100
      Else
         fValue = 0#
      End If
      fld(16) = Format(fValue, "##0.00")
      fld(16) = fld(16) & " %"
      
   Next nCount
   
   ' 代理人比例
   ' 清除內容
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   nRow = nRow + 1
   nTotalAmount = GetTotalAmount(0, 0, False)
   fld(0) = "代理人比例"
   For nAgentCount = 0 To Min(14, m_AgentCount - 1)
      nAmount = GetAgentAmount(m_AgentList(nAgentCount), 0, 0)
      fValue = 0
      If nTotalAmount > 0 Then
         fValue = nAmount / nTotalAmount * 100
      End If
      fld(nAgentCount + 1) = Format(fValue, "##0.00")
   Next nAgentCount
   fld(15) = GetTotalAmount(1, 0, False)
   fValue = GetTotalAmount(1, 0, False) / GetTotalAmount(1, 0, True) * 100
   fld(16) = Format(fValue, "##0.00")
   
   ' 全國比例
   ' 清除內容
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   
   nTotalAmount = GetTotalAmount(0, 0, True)
   fld(0) = "全國比例"
   For nAgentCount = 0 To Min(14, m_AgentCount - 1)
      nAmount = GetAgentAmount(m_AgentList(nAgentCount), 0, 0)
      fValue = 0
      If nTotalAmount > 0 Then
         fValue = nAmount / nTotalAmount * 100
      End If
      fld(nAgentCount + 1) = Format(fValue, "##0.00")
   Next nAgentCount
   fld(15) = GetTotalAmount(1, 0, True)
   
   ' 地區
   ' 清除內容
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   nRow = nRow + 1
   fld(0) = "地區名稱"
   For nZoneCount = 0 To Min(16, m_ZoneCount - 1)
      fld(nZoneCount + 1) = m_ZoneList(nZoneCount).ZoneName
   Next nZoneCount
   
   ' 台一合計
   ' 清除內容
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   nRow = nRow + 2
   fld(0) = "台一合計"
   For nZoneCount = 0 To Min(16, m_ZoneCount - 1)
      nAmount = m_ZoneList(nZoneCount).TaieCount
      fld(nZoneCount + 1) = nAmount
   Next nZoneCount
   
   ' 區域合計
   ' 清除內容
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   nRow = nRow + 2
   fld(0) = "區域合計"
   For nZoneCount = 0 To Min(16, m_ZoneCount - 1)
      nAmount = m_ZoneList(nZoneCount).Count
      fld(nZoneCount + 1) = nAmount
   Next nZoneCount
   
   ' 百分比
   ' 清除內容
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   nRow = nRow + 2
   fld(0) = "百分比"
   For nZoneCount = 0 To Min(16, m_ZoneCount - 1)
      nAmount = m_ZoneList(nZoneCount).TaieCount
      nTotalAmount = m_ZoneList(nZoneCount).Count
      If nTotalAmount > 0 Then
         fValue = nAmount / nTotalAmount * 100
      Else
         fValue = 0#
      End If
      fld(nZoneCount + 1) = Format(fValue, "##0.00")
      fld(nZoneCount + 1) = fld(nZoneCount + 1) & " %"
   Next nZoneCount
   
End Sub

' 從資料庫中取得所有的資料
Private Function GetDBData_RP_931213() As Boolean
   Dim rsMain As New ADODB.Recordset
   Dim strSql As String
   Dim strZone, strZoneCode, strAgent, strAgentCode, strAgentCompany As String
   Dim nZoneIndex, nAgentIndex, nCount As Variant
   Dim bFindZone, bFindAgent As Boolean
   Dim nType As Integer
   Dim nZone As Integer
   Dim nSortX, nSortY As Integer
   Dim AgentTemp As AGENTITEM
   Dim ZoneTemp As ZONEITEM
   Dim bFromSec As Boolean
   Dim bToSec As Boolean
   Dim tmpArr As Variant
   Dim oStrTMBM08 As String
   
   GetDBData_RP_931213 = True
   
   ' 產生SQL查詢語法
   bFromSec = Not IsEmptyText(textTMBM07_1.Text)
   bToSec = Not IsEmptyText(textTMBM07_2.Text)
   'Modify By Sindy 2013/8/19 + length(na01)=3 AND
   If bFromSec = True And bToSec = True Then
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,NA01,NA02,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04 FROM TMBULLETIN, NATION, TAGENT " & _
               "WHERE TMBM07 >= '" & textTMBM07_1 & "' AND " & _
                     "TMBM07 <= '" & textTMBM07_2 & "' AND " & _
                     "TMBM05 = NA03(+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01(+) "
   ElseIf bFromSec = True And bToSec = False Then
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,NA01,NA02,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04 FROM TMBULLETIN, NATION, TAGENT " & _
               "WHERE TMBM07 >= '" & textTMBM07_1 & "' AND " & _
                     "TMBM05 = NA03(+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01(+) "
   ElseIf bFromSec = False And bToSec = True Then
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,NA01,NA02,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04 FROM TMBULLETIN, NATION, TAGENT " & _
               "WHERE TMBM07 <= '" & textTMBM07_2 & "' AND " & _
                     "TMBM05 = NA03(+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01(+) "
   Else
      strSql = "SELECT TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08,NA01,NA02,TA02,DECODE(TA04,NULL,TMBM06,TA04) AS TA04 FROM TMBULLETIN, NATION, TAGENT " & _
               "WHERE TMBM05 = NA03(+) AND " & _
                     "length(na01)=3 AND " & _
                     "TMBM06 = TA03(+) AND " & _
                     "'T' = TA01(+) "
   End If
   
   ' 取得資料庫的資料
   rsMain.CursorLocation = adUseClient
   rsMain.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 無資料則離開
   If rsMain.RecordCount <= 0 Then
      GetDBData_RP_931213 = False
      GoTo EXITSUB
   End If

   ' 設定初始值
   m_ZoneCount = 0
   m_AgentCount = 0

   rsMain.MoveFirst
   ' 依序從資料記錄中取出欄位的內容
   While Not rsMain.EOF
      ' 代理人姓名
      strAgent = Empty
      If IsNull(rsMain.Fields("TMBM06")) = False Then
         strAgent = rsMain.Fields("TMBM06")
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
      ' 地區名稱
      strZone = Empty
      If IsNull(rsMain.Fields("TMBM05")) = False Then
         strZone = rsMain.Fields("TMBM05")
      End If
      oStrTMBM08 = "" & rsMain.Fields("TMBM08")
      tmpArr = Split(oStrTMBM08, ",")
      ' 地區別
      nZone = 0
      ' 地區代碼
      strZoneCode = Empty
      'If IsNull(rsMain.Fields("NA01")) = False Then
      '   strZoneCode = rsMain.Fields("NA01")
      'End If
      'If IsNull(rsMain.Fields("NA02")) = False Then
      '   strZoneCode = rsMain.Fields("NA02")
      'End If
      
      If IsNull(rsMain.Fields("NA02")) = False Then
         Select Case Mid(rsMain.Fields("NA02"), 1, 1)
            Case "A":
               nZone = 1
               Select Case Mid(rsMain.Fields("NA02"), 1, 2)
                  Case "A1":
                     strZone = "北區"
                     strZoneCode = "A10"
                  Case "A2":
                     strZone = "中區"
                     strZoneCode = "A20"
                  Case "A3":
                     strZone = "南區"
                     strZoneCode = "A30"
                  Case "A4":
                     strZone = "高區"
                     strZoneCode = "A40"
                  Case "A5":
                     strZone = "東區"
                     strZoneCode = "A50"
               End Select
            Case "B":
               nZone = 2
               strZoneCode = "B00"
               strZone = "大陸"
            Case Else:
               nZone = 3
               strZoneCode = "C00"
               strZone = "國外"
         End Select
      Else
         nZone = 3
         strZoneCode = "C00"
         strZone = "國外"
      End If
      
      ' 檢查商標種類
      nType = 1
      If IsNull(rsMain.Fields("TMBM02")) = False Then
         Select Case rsMain.Fields("TMBM02")
            ' 正商標
            Case "1": nType = 1
            ' 證明標章
            Case "7": nType = 2
            ' 團體標章
            Case "8": nType = 3
            'Add By Sindy 2010/4/27
            ' 團體商標
            Case "9": nType = 4
         End Select
      End If
      
      ' 地區串列
      bFindZone = False
      For nZoneIndex = 0 To m_ZoneCount - 1
         If m_ZoneList(nZoneIndex).ZoneCode = strZoneCode Then
            bFindZone = True
            m_ZoneList(nZoneIndex).Count = m_ZoneList(nZoneIndex).Count + 1
            m_ZoneList(nZoneIndex).Count08 = m_ZoneList(nZoneIndex).Count08 + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
            'If IsTaie(strAgentCode) = True Then
            If IsTaie(strAgent) = True Then
               m_ZoneList(nZoneIndex).TaieCount = m_ZoneList(nZoneIndex).TaieCount + 1
               m_ZoneList(nZoneIndex).TaieCount08 = m_ZoneList(nZoneIndex).TaieCount08 + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
            End If
            Exit For
         End If
      Next nZoneIndex
      If bFindZone = False Then
         If m_ZoneCount = 0 Then
            nZoneIndex = 0
         Else
            nZoneIndex = UBound(m_ZoneList)
         End If
         ReDim Preserve m_ZoneList(nZoneIndex + 1)
         m_ZoneCount = m_ZoneCount + 1
         m_ZoneList(nZoneIndex).ZoneCode = strZoneCode
         m_ZoneList(nZoneIndex).ZoneName = strZone
         m_ZoneList(nZoneIndex).Count = 1
         m_ZoneList(nZoneIndex).Count08 = IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
         m_ZoneList(nZoneIndex).TaieCount = 0
         m_ZoneList(nZoneIndex).TaieCount08 = 0
         'If IsTaie(strAgentCode) = True Then
         If IsTaie(strAgent) = True Then
            m_ZoneList(nZoneIndex).TaieCount = 1
            m_ZoneList(nZoneIndex).TaieCount08 = IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
         End If
      End If
      
      ' 代理人串列
      If IsEmptyText(strAgent) = True Then
         m_NoAgentItem.Count = m_NoAgentItem.Count + 1
         m_NoAgentItem.Count08 = m_NoAgentItem.Count08 + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
         ' 商標種類
         Select Case nType
            Case 1: m_NoAgentItem.KindAmount(0) = m_NoAgentItem.KindAmount(0) + 1: m_NoAgentItem.KindAmount08(0) = m_NoAgentItem.KindAmount08(0) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
            Case 2: m_NoAgentItem.KindAmount(1) = m_NoAgentItem.KindAmount(1) + 1: m_NoAgentItem.KindAmount08(1) = m_NoAgentItem.KindAmount08(1) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
            Case 3: m_NoAgentItem.KindAmount(2) = m_NoAgentItem.KindAmount(2) + 1: m_NoAgentItem.KindAmount08(2) = m_NoAgentItem.KindAmount08(2) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
            Case 4: m_NoAgentItem.KindAmount(3) = m_NoAgentItem.KindAmount(3) + 1: m_NoAgentItem.KindAmount08(3) = m_NoAgentItem.KindAmount08(3) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
            Case 5: m_NoAgentItem.KindAmount(4) = m_NoAgentItem.KindAmount(4) + 1: m_NoAgentItem.KindAmount08(4) = m_NoAgentItem.KindAmount08(4) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
            Case 6: m_NoAgentItem.KindAmount(5) = m_NoAgentItem.KindAmount(5) + 1: m_NoAgentItem.KindAmount08(5) = m_NoAgentItem.KindAmount08(5) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
            Case 7: m_NoAgentItem.KindAmount(6) = m_NoAgentItem.KindAmount(6) + 1: m_NoAgentItem.KindAmount08(6) = m_NoAgentItem.KindAmount08(6) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
            Case 8: m_NoAgentItem.KindAmount(7) = m_NoAgentItem.KindAmount(7) + 1: m_NoAgentItem.KindAmount08(7) = m_NoAgentItem.KindAmount08(7) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
            'Add By Sindy 2010/4/27
            Case 9: m_NoAgentItem.KindAmount(8) = m_NoAgentItem.KindAmount(8) + 1: m_NoAgentItem.KindAmount08(8) = m_NoAgentItem.KindAmount08(8) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
         End Select
         ' 地區
         Select Case nZone
            Case 1: m_NoAgentItem.ZoneAmount(0) = m_NoAgentItem.ZoneAmount(0) + 1: m_NoAgentItem.ZoneAmount08(0) = m_NoAgentItem.ZoneAmount08(0) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
            Case 2: m_NoAgentItem.ZoneAmount(1) = m_NoAgentItem.ZoneAmount(1) + 1: m_NoAgentItem.ZoneAmount08(1) = m_NoAgentItem.ZoneAmount08(1) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
            Case 3: m_NoAgentItem.ZoneAmount(2) = m_NoAgentItem.ZoneAmount(2) + 1: m_NoAgentItem.ZoneAmount08(2) = m_NoAgentItem.ZoneAmount08(2) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
         End Select
      Else
         ' 搜尋代理人串列
         bFindAgent = False
         For nAgentIndex = 0 To m_AgentCount - 1
            'Modify By Sindy 2010/02/26
            'If m_AgentList(nAgentIndex).AgentName = strAgent Then
            If m_AgentList(nAgentIndex).Company = strAgentCompany Then
            '2010/02/26 End
               bFindAgent = True
               m_AgentList(nAgentIndex).Count = m_AgentList(nAgentIndex).Count + 1
               m_AgentList(nAgentIndex).Count08 = m_AgentList(nAgentIndex).Count08 + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
               ' 商標種類
               Select Case nType
                  Case 1: m_AgentList(nAgentIndex).KindAmount(0) = m_AgentList(nAgentIndex).KindAmount(0) + 1: m_AgentList(nAgentIndex).KindAmount08(0) = m_AgentList(nAgentIndex).KindAmount08(0) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
                  Case 2: m_AgentList(nAgentIndex).KindAmount(1) = m_AgentList(nAgentIndex).KindAmount(1) + 1: m_AgentList(nAgentIndex).KindAmount08(1) = m_AgentList(nAgentIndex).KindAmount08(1) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
                  Case 3: m_AgentList(nAgentIndex).KindAmount(2) = m_AgentList(nAgentIndex).KindAmount(2) + 1: m_AgentList(nAgentIndex).KindAmount08(2) = m_AgentList(nAgentIndex).KindAmount08(2) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
                  Case 4: m_AgentList(nAgentIndex).KindAmount(3) = m_AgentList(nAgentIndex).KindAmount(3) + 1: m_AgentList(nAgentIndex).KindAmount08(3) = m_AgentList(nAgentIndex).KindAmount08(3) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
                  Case 5: m_AgentList(nAgentIndex).KindAmount(4) = m_AgentList(nAgentIndex).KindAmount(4) + 1: m_AgentList(nAgentIndex).KindAmount08(4) = m_AgentList(nAgentIndex).KindAmount08(4) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
                  Case 6: m_AgentList(nAgentIndex).KindAmount(5) = m_AgentList(nAgentIndex).KindAmount(5) + 1: m_AgentList(nAgentIndex).KindAmount08(5) = m_AgentList(nAgentIndex).KindAmount08(5) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
                  Case 7: m_AgentList(nAgentIndex).KindAmount(6) = m_AgentList(nAgentIndex).KindAmount(6) + 1: m_AgentList(nAgentIndex).KindAmount08(6) = m_AgentList(nAgentIndex).KindAmount08(6) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
                  Case 8: m_AgentList(nAgentIndex).KindAmount(7) = m_AgentList(nAgentIndex).KindAmount(7) + 1: m_AgentList(nAgentIndex).KindAmount08(7) = m_AgentList(nAgentIndex).KindAmount08(7) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
                  'Add By Sindy 2010/4/27
                  Case 9: m_AgentList(nAgentIndex).KindAmount(8) = m_AgentList(nAgentIndex).KindAmount(8) + 1: m_AgentList(nAgentIndex).KindAmount08(8) = m_AgentList(nAgentIndex).KindAmount08(8) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
               End Select
               ' 地區
               Select Case nZone
                  Case 1: m_AgentList(nAgentIndex).ZoneAmount(0) = m_AgentList(nAgentIndex).ZoneAmount(0) + 1: m_AgentList(nAgentIndex).ZoneAmount08(0) = m_AgentList(nAgentIndex).ZoneAmount08(0) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
                  Case 2: m_AgentList(nAgentIndex).ZoneAmount(1) = m_AgentList(nAgentIndex).ZoneAmount(1) + 1: m_AgentList(nAgentIndex).ZoneAmount08(1) = m_AgentList(nAgentIndex).ZoneAmount08(1) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
                  Case 3: m_AgentList(nAgentIndex).ZoneAmount(2) = m_AgentList(nAgentIndex).ZoneAmount(2) + 1: m_AgentList(nAgentIndex).ZoneAmount08(2) = m_AgentList(nAgentIndex).ZoneAmount08(2) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
               End Select
               Exit For
            End If
         Next nAgentIndex
         If bFindAgent = False Then
            If m_AgentCount = 0 Then
               nAgentIndex = 0
            Else
               nAgentIndex = UBound(m_AgentList)
            End If
            ReDim Preserve m_AgentList(nAgentIndex + 1)
            m_AgentCount = m_AgentCount + 1
            m_AgentList(nAgentIndex).AgentName = strAgent
            m_AgentList(nAgentIndex).Company = strAgentCompany
            m_AgentList(nAgentIndex).AgentCode = strAgentCode
            m_AgentList(nAgentIndex).Count = 1 'IIf(UBound(TmpArr) < 1, 1, UBound(TmpArr) + 1)
            m_AgentList(nAgentIndex).Count08 = IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
            ' 商標種類
            Select Case nType
               Case 1: m_AgentList(nAgentIndex).KindAmount(0) = m_AgentList(nAgentIndex).KindAmount(0) + 1: m_AgentList(nAgentIndex).KindAmount08(0) = m_AgentList(nAgentIndex).KindAmount08(0) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
               Case 2: m_AgentList(nAgentIndex).KindAmount(1) = m_AgentList(nAgentIndex).KindAmount(1) + 1: m_AgentList(nAgentIndex).KindAmount08(1) = m_AgentList(nAgentIndex).KindAmount08(1) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
               Case 3: m_AgentList(nAgentIndex).KindAmount(2) = m_AgentList(nAgentIndex).KindAmount(2) + 1: m_AgentList(nAgentIndex).KindAmount08(2) = m_AgentList(nAgentIndex).KindAmount08(2) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
               Case 4: m_AgentList(nAgentIndex).KindAmount(3) = m_AgentList(nAgentIndex).KindAmount(3) + 1: m_AgentList(nAgentIndex).KindAmount08(3) = m_AgentList(nAgentIndex).KindAmount08(3) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
               Case 5: m_AgentList(nAgentIndex).KindAmount(4) = m_AgentList(nAgentIndex).KindAmount(4) + 1: m_AgentList(nAgentIndex).KindAmount08(4) = m_AgentList(nAgentIndex).KindAmount08(4) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
               Case 6: m_AgentList(nAgentIndex).KindAmount(5) = m_AgentList(nAgentIndex).KindAmount(5) + 1: m_AgentList(nAgentIndex).KindAmount08(5) = m_AgentList(nAgentIndex).KindAmount08(5) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
               Case 7: m_AgentList(nAgentIndex).KindAmount(6) = m_AgentList(nAgentIndex).KindAmount(6) + 1: m_AgentList(nAgentIndex).KindAmount08(6) = m_AgentList(nAgentIndex).KindAmount08(6) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
               Case 8: m_AgentList(nAgentIndex).KindAmount(7) = m_AgentList(nAgentIndex).KindAmount(7) + 1: m_AgentList(nAgentIndex).KindAmount08(7) = m_AgentList(nAgentIndex).KindAmount08(7) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
               'Add By Sindy 2010/4/27
               Case 9: m_AgentList(nAgentIndex).KindAmount(8) = m_AgentList(nAgentIndex).KindAmount(8) + 1: m_AgentList(nAgentIndex).KindAmount08(8) = m_AgentList(nAgentIndex).KindAmount08(8) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
            End Select
            ' 地區
            Select Case nZone
               Case 1: m_AgentList(nAgentIndex).ZoneAmount(0) = m_AgentList(nAgentIndex).ZoneAmount(0) + 1: m_AgentList(nAgentIndex).ZoneAmount08(0) = m_AgentList(nAgentIndex).ZoneAmount08(0) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
               Case 2: m_AgentList(nAgentIndex).ZoneAmount(1) = m_AgentList(nAgentIndex).ZoneAmount(1) + 1: m_AgentList(nAgentIndex).ZoneAmount08(1) = m_AgentList(nAgentIndex).ZoneAmount08(1) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
               Case 3: m_AgentList(nAgentIndex).ZoneAmount(2) = m_AgentList(nAgentIndex).ZoneAmount(2) + 1: m_AgentList(nAgentIndex).ZoneAmount08(2) = m_AgentList(nAgentIndex).ZoneAmount08(2) + IIf(UBound(tmpArr) < 1, 1, UBound(tmpArr) + 1)
            End Select
         End If
      End If
      ' 移到下一筆記錄
      rsMain.MoveNext
   Wend
   
   ' 對代理人串列依數量的多寡由大到小排序
   For nSortX = 0 To m_AgentCount - 1
      For nSortY = nSortX To m_AgentCount - 1
         If m_AgentList(nSortX).Count08 < m_AgentList(nSortY).Count08 Then
            AgentTemp = m_AgentList(nSortX)
            m_AgentList(nSortX) = m_AgentList(nSortY)
            m_AgentList(nSortY) = AgentTemp
         ElseIf m_AgentList(nSortX).Count08 = m_AgentList(nSortY).Count08 Then
            If m_AgentList(nSortX).Company > m_AgentList(nSortY).Company Then
               AgentTemp = m_AgentList(nSortX)
               m_AgentList(nSortX) = m_AgentList(nSortY)
               m_AgentList(nSortY) = AgentTemp
            End If
         End If
      Next nSortY
   Next nSortX
   
   ' 地區串列依台一在該地區的數量由大到小排序
   For nSortX = 0 To m_ZoneCount - 1
      For nSortY = nSortX To m_ZoneCount - 1
         If (m_ZoneList(nSortX).TaieCount08 / m_ZoneList(nSortX).Count08) < (m_ZoneList(nSortY).TaieCount08 / m_ZoneList(nSortY).Count08) Then
            ZoneTemp = m_ZoneList(nSortX)
            m_ZoneList(nSortX) = m_ZoneList(nSortY)
            m_ZoneList(nSortY) = ZoneTemp
         End If
      Next nSortY
   Next nSortX
   
EXITSUB:
   rsMain.Close
   Set rsMain = Nothing
End Function

' 列印表一的內容
Public Sub Generate_RP_931213()
   Dim strZone As String
   Dim nRow As Integer
   Dim nPage As Integer
   Dim fld(17) As String
   Dim nAgentCount As Variant
   Dim bFindAgent As Boolean
   Dim nType As Integer
   Dim nAmount As Variant
   Dim nTaieAmount As Variant
   Dim nNoAgentAmount As Variant
   Dim nTotalAmount As Variant
   Dim nZoneCount As Variant
   Dim nCount As Variant
   Dim fValue As Variant
   Dim nFinalAmount As Variant
   Dim nNoAgent(3) As Variant
   Dim nX As Integer
   Dim nRight As Long
   Dim nCenter As Long
   Dim strTemp As String
   Dim nSrcHeaderHeight As Integer
   
   BuildField_RP 1
   'Printer.Copies = 5
   nAmount = 0
   nTaieAmount = 0
   nNoAgentAmount = 0
   nTotalAmount = 0
   nAgentCount = 0
   nZoneCount = 0
   nCount = 0
   fValue = 0#
   nNoAgent(0) = 0
   nNoAgent(1) = 0
   nNoAgent(2) = 0
   
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
   PrintPageHeader_RP 1, 1, 0
   
   ' 依序列印表一的第一部份
   nRow = 1
   'Modify By Sindy 2010/4/27
   'For nCount = 0 To 3
   For nCount = 0 To 4
      ' 清除內容
      For nX = 0 To 16: fld(nX) = Empty: Next nX
      ' 第 0 欄位
      Select Case nCount
         Case 0: fld(0) = "事務所"
         Case 1: fld(0) = "正商標"
         Case 2: fld(0) = "證明標章"
         Case 3: fld(0) = "團體標章"
         'Add By Sindy 2010/4/27
         Case 4: fld(0) = "團體商標"
      End Select
      ' 第 1 ~ 14 欄位
      For nAgentCount = 0 To Min(14, m_AgentCount - 1)
         Select Case nCount
            Case 0:
               'fld(nAgentCount + 1) = m_AgentList(nAgentCount).AgentName
               fld(nAgentCount + 1) = m_AgentList(nAgentCount).Company
            'Modify By Sindy 2010/4/27
            'Case 1, 2, 3, 4, 5, 6, 7, 8:
            Case 1, 2, 3, 4, 5, 6, 7, 8, 9:
               fld(nAgentCount + 1) = GetAgentAmount08(m_AgentList(nAgentCount), 0, nCount)
         End Select
      Next nAgentCount
      ' 第 15, 16 欄位
      Select Case nCount
         Case 0:
            fld(15) = Empty
            fld(16) = Empty
         'Modify By Sindy 2010/4/27
         'Case 1, 2, 3, 4, 5, 6, 7, 8:
         Case 1, 2, 3, 4, 5, 6, 7, 8, 9:
            nAmount = GetTaieAmount08(0, nCount)
            nTotalAmount = GetTotalAmount08(0, nCount)
            fld(15) = nTotalAmount
            If nTotalAmount > 0 Then
               fValue = nAmount / nTotalAmount * 100
            Else
               fValue = 0#
            End If
            fld(16) = Format(fValue, "##0.00")
            'fld(16) = fld(16) & " %"
      End Select
      ' 列印
      For nAgentCount = 0 To 16
         If nAgentCount = 0 Then
            Printer.FontSize = 12
            Printer.CurrentX = m_Field(nAgentCount).Left * m_CharWidth
            Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
            Printer.Print fld(nAgentCount)
         Else
            Select Case nCount
               Case 0
                  Printer.FontSize = 8
                  nCenter = ((m_Field(nAgentCount).Left * m_CharWidth) + (m_Field(nAgentCount).Left + m_Field(nAgentCount).Width) * m_CharWidth) / 2
                  Printer.CurrentX = nCenter - Printer.TextWidth(fld(nAgentCount)) / 2
                  Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
                  'Modify By Sindy 2024/4/24 解決列印出來是?
                  'Printer.Print fld(nAgentCount)
                  PUB_PrintUnicodeText fld(nAgentCount), Printer.CurrentX, Printer.CurrentY, 0
                  '2024/4/24 END
               Case Else
                  Printer.FontSize = 8
                  nRight = (m_Field(nAgentCount).Left + m_Field(nAgentCount).Width - 2) * m_CharWidth
                  Printer.CurrentX = nRight - Printer.TextWidth(fld(nAgentCount))
                  Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
                  Printer.Print fld(nAgentCount)
            End Select
         End If
      Next nAgentCount
      nRow = nRow + 1
   Next nCount
   ' 列印分隔線
   PrintSplitLine m_HeaderHeight + nRow
   ' 國別部份
   nRow = nRow + 1
   For nCount = 0 To 3
      ' 清除內容
      '印類
      For nX = 0 To 16: fld(nX) = Empty: Next nX
      ' 第 0 欄位
      Select Case nCount
         Case 0: fld(0) = "全國合計(類)"
         Case 1: fld(0) = "國內合計(類)"
         Case 2: fld(0) = "大陸合計(類)"
         Case 3: fld(0) = "國外合計(類)"
      End Select
      ' 第 1 ~ 14 欄位
      For nAgentCount = 0 To Min(14, m_AgentCount - 1)
         fld(nAgentCount + 1) = GetAgentAmount08(m_AgentList(nAgentCount), 1, nCount)
      Next nAgentCount
      ' 第 15, 16 欄位
      nAmount = 0
      nAmount = GetTaieAmount08(1, nCount)
      nTotalAmount = 0
      nTotalAmount = GetTotalAmount08(1, nCount)
      fld(15) = nTotalAmount
      If nTotalAmount > 0 Then
         fValue = nAmount / nTotalAmount * 100
      Else
         fValue = 0#
      End If
      fld(16) = Format(fValue, "##0.00")
      
      ' 列印
      For nAgentCount = 0 To 16
         If nAgentCount = 0 Then
            Printer.FontSize = 12
            Printer.CurrentX = m_Field(nAgentCount).Left * m_CharWidth
            Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
            Printer.Print fld(nAgentCount)
         Else
            Printer.FontSize = 8
            nRight = (m_Field(nAgentCount).Left + m_Field(nAgentCount).Width - 2) * m_CharWidth
            Printer.CurrentX = nRight - Printer.TextWidth(fld(nAgentCount))
            Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
            Printer.Print fld(nAgentCount)
         End If
      Next nAgentCount
      nRow = nRow + 1
      '印件
      ' 第 0 欄位
      Select Case nCount
         Case 0: fld(0) = "全國合計(件)"
         Case 1: fld(0) = "國內合計(件)"
         Case 2: fld(0) = "大陸合計(件)"
         Case 3: fld(0) = "國外合計(件)"
      End Select
      ' 第 1 ~ 14 欄位
      For nAgentCount = 0 To Min(14, m_AgentCount - 1)
         fld(nAgentCount + 1) = GetAgentAmount(m_AgentList(nAgentCount), 1, nCount)
      Next nAgentCount
      ' 第 15, 16 欄位
      nAmount = 0
      nAmount = GetTaieAmount(1, nCount)
      nTotalAmount = 0
      nTotalAmount = GetTotalAmount(1, nCount)
      fld(15) = nTotalAmount
      If nTotalAmount > 0 Then
         fValue = nAmount / nTotalAmount * 100
      Else
         fValue = 0#
      End If
      fld(16) = Format(fValue, "##0.00")
      
      ' 列印
      For nAgentCount = 0 To 16
         If nAgentCount = 0 Then
            Printer.FontSize = 12
            Printer.CurrentX = m_Field(nAgentCount).Left * m_CharWidth
            Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
            Printer.Print fld(nAgentCount)
         Else
            Printer.FontSize = 8
            nRight = (m_Field(nAgentCount).Left + m_Field(nAgentCount).Width - 2) * m_CharWidth
            Printer.CurrentX = nRight - Printer.TextWidth(fld(nAgentCount))
            Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
            Printer.Print fld(nAgentCount)
         End If
      Next nAgentCount
      nRow = nRow + 2
   Next nCount
   ' 列印分隔線
   nRow = nRow - 1
   PrintSplitLine m_HeaderHeight + nRow
   
   ' 代理人比例
   ' 清除內容
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   nRow = nRow + 1
   nTotalAmount = GetTotalAmount08(0, 0, False)
   fld(0) = "代理人比例(類)"
   For nAgentCount = 0 To Min(14, m_AgentCount - 1)
      nAmount = GetAgentAmount08(m_AgentList(nAgentCount), 0, 0)
      fValue = 0
      If nTotalAmount > 0 Then
         fValue = nAmount / nTotalAmount * 100
      End If
      fld(nAgentCount + 1) = Format(fValue, "##0.00")
   Next nAgentCount
   fld(15) = GetTotalAmount08(1, 0, False)
   fValue = GetTotalAmount08(1, 0, False) / GetTotalAmount08(1, 0, True) * 100
   fld(16) = Format(fValue, "##0.00")
   ' 列印
   For nAgentCount = 0 To 16
      If nAgentCount = 0 Then
         Printer.FontSize = 12
         Printer.CurrentX = m_Field(nAgentCount).Left * m_CharWidth
         Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nAgentCount)
      Else
         Printer.FontSize = 8
         nRight = (m_Field(nAgentCount).Left + m_Field(nAgentCount).Width - 2) * m_CharWidth
         Printer.CurrentX = nRight - Printer.TextWidth(fld(nAgentCount))
         Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nAgentCount)
      End If
   Next nAgentCount
   
For nX = 0 To 16: fld(nX) = Empty: Next nX
   nRow = nRow + 1
   nTotalAmount = GetTotalAmount(0, 0, False)
   fld(0) = "代理人比例(件)"
   For nAgentCount = 0 To Min(14, m_AgentCount - 1)
      nAmount = GetAgentAmount(m_AgentList(nAgentCount), 0, 0)
      fValue = 0
      If nTotalAmount > 0 Then
         fValue = nAmount / nTotalAmount * 100
      End If
      fld(nAgentCount + 1) = Format(fValue, "##0.00")
   Next nAgentCount
   fld(15) = GetTotalAmount(1, 0, False)
   fValue = GetTotalAmount(1, 0, False) / GetTotalAmount(1, 0, True) * 100
   fld(16) = Format(fValue, "##0.00")
   ' 列印
   For nAgentCount = 0 To 16
      If nAgentCount = 0 Then
         Printer.FontSize = 12
         Printer.CurrentX = m_Field(nAgentCount).Left * m_CharWidth
         Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nAgentCount)
      Else
         Printer.FontSize = 8
         nRight = (m_Field(nAgentCount).Left + m_Field(nAgentCount).Width - 2) * m_CharWidth
         Printer.CurrentX = nRight - Printer.TextWidth(fld(nAgentCount))
         Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nAgentCount)
      End If
   Next nAgentCount
   
   ' 全國比例
   ' 清除內容
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   nRow = nRow + 1
   nTotalAmount = GetTotalAmount08(0, 0, True)
   fld(0) = "全國比例(類)"
   For nAgentCount = 0 To Min(14, m_AgentCount - 1)
      nAmount = GetAgentAmount08(m_AgentList(nAgentCount), 0, 0)
      fValue = 0
      If nTotalAmount > 0 Then
         fValue = nAmount / nTotalAmount * 100
      End If
      fld(nAgentCount + 1) = Format(fValue, "##0.00")
   Next nAgentCount
   fld(15) = GetTotalAmount08(1, 0, True)
   ' 列印
   For nAgentCount = 0 To 16
      If nAgentCount = 0 Then
         Printer.FontSize = 12
         Printer.CurrentX = m_Field(nAgentCount).Left * m_CharWidth
         Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nAgentCount)
      Else
         Printer.FontSize = 8
         nRight = (m_Field(nAgentCount).Left + m_Field(nAgentCount).Width - 2) * m_CharWidth
         Printer.CurrentX = nRight - Printer.TextWidth(fld(nAgentCount))
         Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nAgentCount)
      End If
   Next nAgentCount
   
For nX = 0 To 16: fld(nX) = Empty: Next nX
   nRow = nRow + 1
   nTotalAmount = GetTotalAmount(0, 0, True)
   fld(0) = "全國比例(件)"
   For nAgentCount = 0 To Min(14, m_AgentCount - 1)
      nAmount = GetAgentAmount(m_AgentList(nAgentCount), 0, 0)
      fValue = 0
      If nTotalAmount > 0 Then
         fValue = nAmount / nTotalAmount * 100
      End If
      fld(nAgentCount + 1) = Format(fValue, "##0.00")
   Next nAgentCount
   fld(15) = GetTotalAmount(1, 0, True)
   ' 列印
   For nAgentCount = 0 To 16
      If nAgentCount = 0 Then
         Printer.FontSize = 12
         Printer.CurrentX = m_Field(nAgentCount).Left * m_CharWidth
         Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nAgentCount)
      Else
         Printer.FontSize = 8
         nRight = (m_Field(nAgentCount).Left + m_Field(nAgentCount).Width - 2) * m_CharWidth
         Printer.CurrentX = nRight - Printer.TextWidth(fld(nAgentCount))
         Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nAgentCount)
      End If
   Next nAgentCount
   ' 列印雙分隔線
   nRow = nRow + 1
   PrintTerminateLine m_HeaderHeight + nRow
   
   'add by nickc 2007/04/19 加入 本所所有申請案含非本所繳註冊費件數
   Dim bFromSec As Boolean
   Dim bToSec As Boolean
   CheckOC3
   bFromSec = Not IsEmptyText(textTMBM07_1.Text)
   bToSec = Not IsEmptyText(textTMBM07_2.Text)
   If bFromSec = True And bToSec = True Then
      strSql = "SELECT count(*) A,0 B from TMBULLETIN,trademark " & _
               "WHERE TMBM07 >= '" & textTMBM07_1 & "' AND " & _
                     "TMBM07 <= '" & textTMBM07_2 & "' AND " & _
                     "tmbm04 = tm12 and tm10='000' "
      strSql = strSql & " union select 0 A,count(*) B from tmbulletin " & _
                  "where tmbm07>='" & textTMBM07_1 & "' and " & _
                  " tmbm07<='" & textTMBM07_2 & "' "
   ElseIf bFromSec = True And bToSec = False Then
      strSql = "SELECT count(*) A,0 B FROM TMBULLETIN,trademark " & _
               "WHERE TMBM07 >= '" & textTMBM07_1 & "' AND " & _
                     "TMBM04 = tm12 and tm10='000' "
      strSql = strSql & "union select 0 A,count(*) B from tmbulletin " & _
                  "where tmbm07>='" & textTMBM07_1 & "' "
   ElseIf bFromSec = False And bToSec = True Then
      strSql = "SELECT count(*) A,0 B FROM TMBULLETIN,trademark " & _
               "WHERE TMBM07 <= '" & textTMBM07_2 & "' AND " & _
                     "TMBM04 = tm12  and tm10='000' "
      strSql = strSql & " union select 0 A,count(*) B from tmbulletin " & _
                  "where tmbm07<='" & textTMBM07_2 & "' "
   Else
      strSql = "SELECT count(*) A,0 B FROM TMBULLETIN, trademark" & _
               "WHERE TMBM04 = tm12 and tm10='000' "
      strSql = strSql & " union select 0 A,count(*) B from tmbulletin "
   End If
   strSql = "select sum(A),sum(B) from (" & strSql & ") C "
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If AdoRecordSet3.RecordCount <> 0 Then
        nRow = nRow + 1
        Printer.CurrentX = (m_LeftMargin + nCount) * m_CharWidth
        Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
        Printer.Print "本所所有申請案含非本所繳註冊費件數：" & CheckStr(AdoRecordSet3.Fields(0)) & "                     全國比例：" & Format((CheckStr(AdoRecordSet3.Fields(0)) / CheckStr(AdoRecordSet3.Fields(1))) * 100, "##0.00")
   End If
   CheckOC3
   'add by nickc 2007/04/19 end ************************************
   
   ' 記錄原Header的高度
   nSrcHeaderHeight = m_HeaderHeight
   
   ' 表一與表二的間距
   nRow = nRow + 1
   
   ' 列印表二的表頭
   BuildField_RP 2
   PrintPageHeader_RP 2, 1, nSrcHeaderHeight + nRow
   nRow = nRow + m_HeaderHeight
   
   ' 地區
   ' 清除內容
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   nRow = nRow + 1
   fld(0) = "地區名稱"
   For nZoneCount = 0 To Min(16, m_ZoneCount - 1)
      fld(nZoneCount + 1) = m_ZoneList(nZoneCount).ZoneName
   Next nZoneCount
   For nCount = 0 To 16
      If nCount = 0 Then
         Printer.FontSize = 12
         Printer.CurrentX = m_Field(nCount).Left * m_CharWidth
         Printer.CurrentY = (nSrcHeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nCount)
      Else
         Printer.FontSize = 8
         nCenter = ((m_Field(nCount).Left * m_CharWidth) + (m_Field(nCount).Left + m_Field(nCount).Width) * m_CharWidth) / 2
         Printer.CurrentX = nCenter - Printer.TextWidth(fld(nCount)) / 2
         Printer.CurrentY = (nSrcHeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nCount)
         Printer.FontSize = 12
      End If
   Next nCount
   
   ' 台一合計
   ' 清除內容
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   nRow = nRow + 1
   fld(0) = "台一合計"
   For nZoneCount = 0 To Min(16, m_ZoneCount - 1)
      nAmount = m_ZoneList(nZoneCount).TaieCount08
      fld(nZoneCount + 1) = nAmount
   Next nZoneCount
   For nCount = 0 To 16
      If nCount = 0 Then
         Printer.FontSize = 12
         Printer.CurrentX = m_Field(nCount).Left * m_CharWidth
         Printer.CurrentY = (nSrcHeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nCount)
      Else
         Printer.FontSize = 8
         nRight = (m_Field(nCount).Left + m_Field(nCount).Width - 2) * m_CharWidth
         Printer.CurrentX = nRight - Printer.TextWidth(fld(nCount))
         Printer.CurrentY = (nSrcHeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nCount)
      End If
   Next nCount
   
   ' 區域合計
   ' 清除內容
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   nRow = nRow + 1
   fld(0) = "區域合計"
   For nZoneCount = 0 To Min(16, m_ZoneCount - 1)
      nAmount = m_ZoneList(nZoneCount).Count08
      fld(nZoneCount + 1) = nAmount
   Next nZoneCount
   For nCount = 0 To 14
      If nCount = 0 Then
         Printer.FontSize = 12
         Printer.CurrentX = m_Field(nCount).Left * m_CharWidth
         Printer.CurrentY = (nSrcHeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nCount)
      Else
         Printer.FontSize = 8
         nRight = (m_Field(nCount).Left + m_Field(nCount).Width - 2) * m_CharWidth
         Printer.CurrentX = nRight - Printer.TextWidth(fld(nCount))
         Printer.CurrentY = (nSrcHeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nCount)
      End If
   Next nCount
   
   ' 百分比
   ' 清除內容
   For nX = 0 To 16: fld(nX) = Empty: Next nX
   nRow = nRow + 1
   fld(0) = "百分比"
   For nZoneCount = 0 To Min(16, m_ZoneCount - 1)
      nAmount = m_ZoneList(nZoneCount).TaieCount08
      nTotalAmount = m_ZoneList(nZoneCount).Count08
      If nTotalAmount > 0 Then
         fValue = nAmount / nTotalAmount * 100
      Else
         fValue = 0#
      End If
      fld(nZoneCount + 1) = Format(fValue, "##0.00")
      'fld(nZoneCount + 1) = fld(nZoneCount + 1) & " %"
   Next nZoneCount
   For nCount = 0 To 16
      If nCount = 0 Then
         Printer.FontSize = 12
         Printer.CurrentX = m_Field(nCount).Left * m_CharWidth
         Printer.CurrentY = (nSrcHeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nCount)
      Else
         Printer.FontSize = 8
         nRight = (m_Field(nCount).Left + m_Field(nCount).Width - 2) * m_CharWidth
         Printer.CurrentX = nRight - Printer.TextWidth(fld(nCount))
         Printer.CurrentY = (nSrcHeaderHeight + m_TopMargin + nRow) * m_CharHeight
         Printer.Print fld(nCount)
      End If
   Next nCount
   
   ' 列印雙分隔線
   nRow = nRow + 1
   PrintTerminateLine m_HeaderHeight + nRow

   'add by nickc 2007/05/22 加入只印一張
   'Modify By Sindy 2022/3/23 桂英取消列印5份,1份PDF即可,此報表都要出來
   'If nCopys = 5 Then
   If nCopys = Trim(txt1.Text) Then
        'add by nickc 2007/04/19 加入 本所申請但非本所繳註冊費的案件明細
        'Modify By Sindy 2014/7/28 +and cp10='101' and cp09<'B'
        If bFromSec = True And bToSec = True Then
           strSql = "select cp12,a0902,cp13,st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04,tm05,cu01||cu02,cu04 from caseprogress,staff,trademark,customer,acc090  where (cp01,cp02,cp03,cp04,cp09) in (" & _
                               "SELECT cp01,cp02,cp03,cp04,min(cp09) from TMBULLETIN,trademark ,tagent,caseprogress " & _
                               " WHERE TMBM07 >= '" & textTMBM07_1 & "' AND tmbm06=ta03(+) and 'T'=ta01(+) " & _
                               " and TMBM07 <= '" & textTMBM07_2 & "' AND tmbm04 = tm12 and decode(ta04,'台一國際',1,0) = 0 " & _
                              " and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and tm10='000' and cp10='101' and cp09<'B' " & _
                              " group by cp01,cp02,cp03,cp04) " & _
                              " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp13=st01(+) " & _
                              " and cp12=a0901(+) and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) "
        ElseIf bFromSec = True And bToSec = False Then
           strSql = "select cp12,a0902,cp13,st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04,tm05,cu01||cu02,cu04 from caseprogress,staff,trademark,customer,acc090  where (cp01,cp02,cp03,cp04,cp09) in (" & _
                               "SELECT cp01,cp02,cp03,cp04,min(cp09) from TMBULLETIN,trademark ,tagent,caseprogress " & _
                               " WHERE TMBM07 >= '" & textTMBM07_1 & "' AND tmbm06=ta03(+) and 'T'=ta01(+) " & _
                               " AND tmbm04 = tm12 and decode(ta04,'台一國際',1,0) = 0 " & _
                              " and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and tm10='000' and cp10='101' and cp09<'B' " & _
                              " group by cp01,cp02,cp03,cp04) " & _
                              " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp13=st01(+) " & _
                              " and cp12=a0901(+) and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) "
        ElseIf bFromSec = False And bToSec = True Then
           strSql = "select cp12,a0902,cp13,st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04,tm05,cu01||cu02,cu04 from caseprogress,staff,trademark,customer,acc090  where (cp01,cp02,cp03,cp04,cp09) in (" & _
                               "SELECT cp01,cp02,cp03,cp04,min(cp09) from TMBULLETIN,trademark ,tagent,caseprogress " & _
                               " WHERE  tmbm06=ta03(+) and 'T'=ta01(+) " & _
                               " and TMBM07 <= '" & textTMBM07_2 & "' AND tmbm04 = tm12 and decode(ta04,'台一國際',1,0) = 0 " & _
                              " and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and tm10='000' and cp10='101' and cp09<'B' " & _
                              " group by cp01,cp02,cp03,cp04) " & _
                              " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp13=st01(+) " & _
                              " and cp12=a0901(+) and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) "
        Else
           strSql = "select cp12,a0902,cp13,st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04,tm05,cu01||cu02,cu04 from caseprogress,staff,trademark,customer,acc090  where (cp01,cp02,cp03,cp04,cp09) in (" & _
                               "SELECT cp01,cp02,cp03,cp04,min(cp09) from TMBULLETIN,trademark ,tagent,caseprogress " & _
                               " WHERE  tmbm06=ta03(+) and 'T'=ta01(+) " & _
                               " AND tmbm04 = tm12 and decode(ta04,'台一國際',1,0) = 0 " & _
                              " and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and tm10='000' and cp10='101' and cp09<'B' " & _
                              " group by cp01,cp02,cp03,cp04) " & _
                              " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp13=st01(+) " & _
                              " and cp12=a0901(+) and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) "
        End If
        AdoRecordSet3.CursorLocation = adUseClient
        AdoRecordSet3.Open strSql & "order by cp12,cp13 ", cnnConnection, adOpenStatic, adLockReadOnly
        If AdoRecordSet3.RecordCount <> 0 Then
             Dim SeekNowArea As String
             Dim SeekNowAreaCount As String
             Printer.NewPage
             nPage = nPage + 1
             PrintPageHeader_RP 3, 1, 0
             nRow = 1
             SeekNowArea = ""
             SeekNowAreaCount = 0
             With AdoRecordSet3
                 .MoveFirst
                 Do While Not .EOF
                     If SeekNowArea <> CheckStr(.Fields(0)) Then
                         If SeekNowArea <> "" Then
                             ' 列印雙分隔線
                             nRow = nRow + 1
                             PrintTerminateLine m_HeaderHeight + nRow
                             If nRow >= 36 Then
                                 Printer.NewPage
                                 nPage = nPage + 1
                                 PrintPageHeader_RP 3, 1, 0
                                 nRow = 1
                             End If
                             nRow = nRow + 1
                             Printer.CurrentX = 1000
                             Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                             Printer.Print "總計：" & Trim(SeekNowAreaCount) & "件"
                             If nRow >= 36 Then
                                 Printer.NewPage
                                 nPage = nPage + 1
                                 PrintPageHeader_RP 3, 1, 0
                                 nRow = 1
                             End If
                             ' 列印雙分隔線
                             nRow = nRow + 1
                             PrintTerminateLine m_HeaderHeight + nRow
                             If nRow >= 36 Then
                                 Printer.NewPage
                                 nPage = nPage + 1
                                 PrintPageHeader_RP 3, 1, 0
                                 nRow = 1
                             End If
                             nRow = nRow + 1
                         End If
                         SeekNowAreaCount = 0
                         SeekNowArea = CheckStr(.Fields(0))
                         Printer.CurrentX = 120
                         Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                         Printer.Print CheckStr(.Fields(1))
                     Else
                         nRow = nRow + 1
                     End If
                     Printer.CurrentX = 1500
                     Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                     Printer.Print CheckStr(.Fields(3))
                     Printer.CurrentX = 3000
                     Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                     Printer.Print CheckStr(.Fields(4))
                     Printer.CurrentX = 5500
                     Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                     Printer.Print CheckStr(.Fields(5))
                     Printer.CurrentX = 9000
                     Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                     Printer.Print CheckStr(.Fields(6)) & " " & CheckStr(.Fields(7))
                     SeekNowAreaCount = SeekNowAreaCount + 1
                     If nRow >= 36 Then
                         Printer.NewPage
                         nPage = nPage + 1
                         PrintPageHeader_RP 3, 1, 0
                         nRow = 1
                     End If
                     .MoveNext
                 Loop
                 nRow = nRow + 1
                 PrintTerminateLine m_HeaderHeight + nRow
                 If nRow >= 36 Then
                     Printer.NewPage
                     nPage = nPage + 1
                     PrintPageHeader_RP 3, 1, 0
                     nRow = 1
                 End If
                 nRow = nRow + 1
                 Printer.CurrentX = 1000
                 Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                 Printer.Print "總計：" & Trim(SeekNowAreaCount) & "件"
                 If nRow >= 36 Then
                     Printer.NewPage
                     nPage = nPage + 1
                     PrintPageHeader_RP 3, 1, 0
                     nRow = 1
                 End If
                 ' 列印雙分隔線
                 nRow = nRow + 1
                 PrintTerminateLine m_HeaderHeight + nRow
             End With
        End If
        CheckOC3
        'add by nickc 2007/04/19 end ************************************
   End If
   
   Printer.EndDoc
End Sub

Private Function GetTaieAmount08(ByVal nType As Integer, ByVal nIndex As Integer) As Variant
   Dim nAgentCount As Variant
   Dim nAmount As Variant
   Dim bFind As Boolean
   
   nAmount = 0
   bFind = False
   For nAgentCount = 0 To m_AgentCount - 1
      'If m_AgentList(nAgentCount).AgentCode = "001" Then
      If m_AgentList(nAgentCount).AgentName = "林晉章" Then
         bFind = True
         Exit For
      End If
   Next nAgentCount
   
   If bFind = False Then
      GetTaieAmount08 = 0
   Else
      Select Case nType
         Case 0
            Select Case nIndex
               Case 0: nAmount = m_AgentList(nAgentCount).KindAmount08(0) + m_AgentList(nAgentCount).KindAmount08(1) + m_AgentList(nAgentCount).KindAmount08(2) + m_AgentList(nAgentCount).KindAmount08(3) + m_AgentList(nAgentCount).KindAmount08(4) + m_AgentList(nAgentCount).KindAmount08(5) + m_AgentList(nAgentCount).KindAmount08(6) + m_AgentList(nAgentCount).KindAmount08(7)
               'Modify By Sindy 2010/4/27
               'Case 1, 2, 3, 4, 5, 6, 7, 8:
               Case 1, 2, 3, 4, 5, 6, 7, 8, 9:
                  nAmount = m_AgentList(nAgentCount).KindAmount08(nIndex - 1)
            End Select
         Case 1
            Select Case nIndex
               Case 0: nAmount = m_AgentList(nAgentCount).ZoneAmount08(0) + m_AgentList(nAgentCount).ZoneAmount08(1) + m_AgentList(nAgentCount).ZoneAmount08(2)
               Case 1, 2, 3:
                  nAmount = m_AgentList(nAgentCount).ZoneAmount08(nIndex - 1)
            End Select
      End Select
   End If
   GetTaieAmount08 = nAmount
End Function

Private Function GetTotalAmount08(ByVal nType As Integer, ByVal nIndex As Integer, Optional ByVal bIncludeNoAgent As Boolean = True) As Variant
   Dim nAgentCount As Variant
   Dim nAmount As Variant
   nAmount = 0
   
   Select Case nType
      Case 0
         Select Case nIndex
            Case 0:
               For nAgentCount = 0 To m_AgentCount - 1
                  nAmount = nAmount + m_AgentList(nAgentCount).KindAmount08(0) + m_AgentList(nAgentCount).KindAmount08(1) + m_AgentList(nAgentCount).KindAmount08(2) + m_AgentList(nAgentCount).KindAmount08(3) + m_AgentList(nAgentCount).KindAmount08(4) + m_AgentList(nAgentCount).KindAmount08(5) + m_AgentList(nAgentCount).KindAmount08(6) + m_AgentList(nAgentCount).KindAmount08(7)
               Next nAgentCount
               If bIncludeNoAgent = True Then
                  nAmount = nAmount + m_NoAgentItem.KindAmount08(0) + m_NoAgentItem.KindAmount08(1) + m_NoAgentItem.KindAmount08(2) + m_NoAgentItem.KindAmount08(3) + m_NoAgentItem.KindAmount08(4) + m_NoAgentItem.KindAmount08(5) + m_NoAgentItem.KindAmount08(6) + m_NoAgentItem.KindAmount08(7)
               End If
            'Modify By Sindy 2010/4/27
            'Case 1, 2, 3, 4, 5, 6, 7, 8:
            Case 1, 2, 3, 4, 5, 6, 7, 8, 9:
               For nAgentCount = 0 To m_AgentCount - 1
                  nAmount = nAmount + m_AgentList(nAgentCount).KindAmount08(nIndex - 1)
               Next nAgentCount
               If bIncludeNoAgent = True Then
                  nAmount = nAmount + m_NoAgentItem.KindAmount08(nIndex - 1)
               End If
         End Select
      Case 1
         Select Case nIndex
            Case 0:
               For nAgentCount = 0 To m_AgentCount - 1
                  nAmount = nAmount + m_AgentList(nAgentCount).ZoneAmount08(0) + m_AgentList(nAgentCount).ZoneAmount08(1) + m_AgentList(nAgentCount).ZoneAmount08(2)
               Next nAgentCount
               If bIncludeNoAgent = True Then
                  nAmount = nAmount + m_NoAgentItem.ZoneAmount08(0) + m_NoAgentItem.ZoneAmount08(1) + m_NoAgentItem.ZoneAmount08(2)
               End If
            Case 1, 2, 3:
               For nAgentCount = 0 To m_AgentCount - 1
                  nAmount = nAmount + m_AgentList(nAgentCount).ZoneAmount08(nIndex - 1)
               Next nAgentCount
               If bIncludeNoAgent = True Then
                  nAmount = nAmount + m_NoAgentItem.ZoneAmount08(nIndex - 1)
               End If
         End Select
   End Select
   GetTotalAmount08 = nAmount
End Function

'Add By Sindy 2010/01/12
Private Sub txt1_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub
