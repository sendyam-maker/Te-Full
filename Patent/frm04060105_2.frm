VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm04060105_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "國內市場佔有率查詢"
   ClientHeight    =   5760
   ClientLeft      =   -2856
   ClientTop       =   2436
   ClientWidth     =   9348
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9348
   Begin VB.TextBox text03 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   840
      TabIndex        =   7
      Top             =   5400
      Width           =   1572
   End
   Begin VB.TextBox text02 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   4080
      TabIndex        =   5
      Top             =   660
      Width           =   2652
   End
   Begin VB.TextBox text01 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1320
      TabIndex        =   4
      Top             =   660
      Width           =   1572
   End
   Begin VB.CommandButton bottonNext 
      Caption         =   "下一國籍(&N)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6828
      TabIndex        =   0
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton buttonExit 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8052
      TabIndex        =   1
      Top             =   70
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   4275
      Left            =   90
      TabIndex        =   8
      Top             =   990
      Width           =   9075
      _ExtentX        =   16002
      _ExtentY        =   7535
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
   Begin VB.Label Label3 
      Caption         =   "合計 :"
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   5400
      Width           =   732
   End
   Begin VB.Label Label2 
      Caption         =   "公告日 :"
      Height          =   252
      Left            =   3240
      TabIndex        =   3
      Top             =   660
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "申請人國籍 :"
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   660
      Width           =   1092
   End
End
Attribute VB_Name = "frm04060105_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/18 改成Form2.0 ; grdList改字型=新細明體-ExtB(MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

' 宣告報表表頭的欄位其資料型態
Private Type AGENTITEM
   AgentCode As String
   AgentName As String
   Type1 As Variant
   Type2 As Variant
   Type3 As Variant
   Count As Variant
End Type

Dim m_AgentList() As AGENTITEM
Dim m_AgentCount As Variant

Dim m_CountryBegin As String
Dim m_CountryCurr As String
Dim m_CountryEnd As String
Dim m_DateFrom As String
Dim m_DateTo As String
Dim m_Recordset As New ADODB.Recordset
'
Dim m_CurrSel As Integer

Private Sub Form_Load()
   text01.BackColor = &H8000000F
   text02.BackColor = &H8000000F
   text03.BackColor = &H8000000F
   bottonNext.Enabled = True

   MoveFormToCenter Me
   ExecuteQuery
   ListData
End Sub

Private Sub ClearAgentList()
   If m_AgentCount > 0 Then
      Erase m_AgentList
   End If
   m_AgentCount = 0
End Sub

Private Sub bottonNext_Click()
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   ListData
End Sub

Private Sub buttonExit_Click()
   Unload Me
   frm04060105_1.Show
End Sub

Public Sub SetData(ByVal dateFrom As String, ByVal dateTo As String, ByVal countryBegin As String, ByVal countryEnd As String)
   m_DateFrom = dateFrom
   m_DateTo = dateTo
   m_CountryBegin = countryBegin
   m_CountryEnd = countryEnd
   m_CountryCurr = countryBegin
End Sub

Private Sub ExecuteQuery()
Dim rsTPB As ADODB.Recordset
Dim strSql As String
Dim strSubSQL As String
Dim bContinue As Boolean
Dim nIndex
   
   'Modify by Morgan 2007/10/17 改本所辦理之不出名案件且公報也沒有代理人的也列入本所統計
   'strSQL = "SELECT * from TPBulletin where 1=1"
   strSql = "SELECT A.*,PA01,PA02,PA03,PA04 FROM TPBulletin A,patent B where pa11(+)=tpb01 and pa23(+)='1'"
   'end 2007/10/17
   
   strSubSQL = Empty
   If IsEmpty(m_DateFrom) = False Then
      strSubSQL = strSubSQL & " and TPB03 >= " & ChangeTStringToWString(m_DateFrom) & " "
   End If
   If IsEmpty(m_DateTo) = False Then
      strSubSQL = strSubSQL & " and TPB03 <= " & ChangeTStringToWString(m_DateTo) & " "
   End If
   If IsEmpty(m_DateFrom) = False Or IsEmpty(m_DateTo) = False Then
      pub_QL05 = pub_QL05 & ";" & frm04060105_1.Label1 & frm04060105_1.text01_01 & "-" & frm04060105_1.text01_02 'Add By Sindy 2010/12/2
   End If
   If IsEmpty(m_CountryBegin) = False Then
      strSubSQL = strSubSQL & " and TPB06 >= '" & m_CountryBegin & "' "
   End If
   If IsEmpty(m_CountryEnd) = False Then
      strSubSQL = strSubSQL & " and TPB06 <= '" & m_CountryEnd & "' "
   End If
   If IsEmpty(m_CountryBegin) = False Or IsEmpty(m_CountryEnd) = False Then
      pub_QL05 = pub_QL05 & ";" & frm04060105_1.Label2 & frm04060105_1.text02_01 & "-" & frm04060105_1.text02_02 'Add By Sindy 2010/12/2
   End If
   
   strSql = strSql & strSubSQL & " ORDER BY TPB06"
   
   If (m_Recordset.State <> adStateClosed) Then
      m_Recordset.Close
   End If
      
   m_Recordset.CursorLocation = adUseClient
   m_Recordset.Open strSql, cnnConnection, adOpenDynamic
   
   If m_Recordset.RecordCount > 0 Then
      InsertQueryLog (m_Recordset.RecordCount) 'Add By Sindy 2010/12/2
      m_Recordset.MoveFirst
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/2
   End If
End Sub

' 由國籍代碼取得國籍的名稱
Public Function GetNationName(ByVal strNation As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   GetNationName = Empty
   strSql = "SELECT * FROM NATION WHERE NA01 = '" & strNation & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("NA03")) = False Then
         GetNationName = rsTmp.Fields("NA03")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 由代理人代碼取得事務所的名稱
Public Function GetAgentCompany(ByVal strAgent As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   GetAgentCompany = Empty
   strSql = "SELECT * FROM TAGENT WHERE TA01 = 'P' AND TA02 = '" & strAgent & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("TA04")) = False Then
         GetAgentCompany = rsTmp.Fields("TA04")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Sub ListData()
   Dim strCountry As String
   Dim strCmp As String
   Dim nType As Integer
   Dim bFindAgent As Boolean
   Dim nAgentCount As Variant
   Dim strAgentCode As String
   Dim strAgentName As String
   Dim nTotalAmount As Double
   Dim nCount As Integer
   Dim agentTemp As AGENTITEM
   Dim nSortX As Integer
   Dim nSortY As Integer
   Dim nIndex As Variant
   
   If m_Recordset.RecordCount <= 0 Then
      GoTo EXITSUB
   End If
   
   If m_Recordset.EOF = True Then
      GoTo EXITSUB
   End If
   
   strCountry = Empty
   If IsNull(m_Recordset.Fields("TPB06")) = False Then
      strCountry = m_Recordset.Fields("TPB06")
   End If
   ' 清除代理人事務所串列
   ClearAgentList
   
   text01 = GetNationName(strCountry)
   
   nTotalAmount = 0
   Do While m_Recordset.EOF <> True
      strCmp = Empty
      If IsNull(m_Recordset.Fields("TPB06")) = False Then
         strCmp = m_Recordset.Fields("TPB06")
      End If
      
      If strCmp <> strCountry Then
         Exit Do
      End If
      
      ' 總數累計
      nTotalAmount = nTotalAmount + 1
      
      strAgentCode = Empty
      
      'Add by Morgan 2007/10/18 檢查無代理人且為本所案件
      If IsNull(m_Recordset.Fields("TPB07")) And Not IsNull(m_Recordset.Fields("PA01")) Then
         strSql = "select substr(max(cp27||cp22),9) from caseprogress where cp01='" & m_Recordset("pa01") & "' and cp02='" & m_Recordset("pa02") & "' and cp03='" & m_Recordset("pa03") & "' and cp04='" & m_Recordset("pa04") & "' and cp09<'C' and cp27<" & m_Recordset("tpb03")
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            '若最後發文的AB類程序為不出名則算台一案件
            If "" & RsTemp(0) = "N" Then
               strAgentCode = "01"
            End If
         End If
      'end 2007/10/18
      ElseIf IsNull(m_Recordset.Fields("TPB07")) = False Then
         strAgentCode = m_Recordset.Fields("TPB07")
      End If
      strAgentName = GetAgentCompany(strAgentCode)
      
      nType = 0
      'Modify by Morgan 2010/12/28 申請案號改碼數
      'Select Case Mid(m_Recordset.Fields("TPB01"), 3, 1)
      Select Case Mid(m_Recordset.Fields("TPB01"), 4, 1)
         Case 1: nType = 1
         Case 2: nType = 2
         Case 3: nType = 3
      End Select
      
      bFindAgent = False
      For nAgentCount = 0 To m_AgentCount - 1
         If m_AgentList(nAgentCount).AgentCode = strAgentCode Then
            bFindAgent = True
            m_AgentList(nAgentCount).Count = m_AgentList(nAgentCount).Count + 1
            Select Case nType
               Case 1: m_AgentList(nAgentCount).Type1 = m_AgentList(nAgentCount).Type1 + 1
               Case 2: m_AgentList(nAgentCount).Type2 = m_AgentList(nAgentCount).Type2 + 1
               Case 3: m_AgentList(nAgentCount).Type3 = m_AgentList(nAgentCount).Type3 + 1
            End Select
            Exit For
         End If
      Next nAgentCount
      If bFindAgent = False Then
         nAgentCount = m_AgentCount
         ReDim Preserve m_AgentList(nAgentCount + 1)
         m_AgentCount = m_AgentCount + 1
         m_AgentList(nAgentCount).AgentCode = strAgentCode
         m_AgentList(nAgentCount).AgentName = strAgentName
         m_AgentList(nAgentCount).Count = 1
         Select Case nType
            Case 1: m_AgentList(nAgentCount).Type1 = m_AgentList(nAgentCount).Type1 + 1
            Case 2: m_AgentList(nAgentCount).Type2 = m_AgentList(nAgentCount).Type2 + 1
            Case 3: m_AgentList(nAgentCount).Type3 = m_AgentList(nAgentCount).Type3 + 1
         End Select
      End If
      m_Recordset.MoveNext
   Loop
   
   ' 排序
   For nSortX = 0 To m_AgentCount - 1
      For nSortY = nSortX To m_AgentCount - 1
         If m_AgentList(nSortX).Count < m_AgentList(nSortY).Count Then
            agentTemp = m_AgentList(nSortX)
            m_AgentList(nSortX) = m_AgentList(nSortY)
            m_AgentList(nSortY) = agentTemp
         End If
      Next nSortY
   Next nSortX
      
   text02 = m_DateFrom & " - " & m_DateTo
   text03 = nTotalAmount
   
   InitialGridList
   For nAgentCount = 0 To m_AgentCount - 1
      grdList.Rows = grdList.Rows + 1
      nIndex = grdList.Rows - 1
         
      grdList.TextMatrix(nIndex, 1) = m_AgentList(nAgentCount).AgentName
      grdList.TextMatrix(nIndex, 2) = m_AgentList(nAgentCount).Type1
      grdList.TextMatrix(nIndex, 3) = Format(m_AgentList(nAgentCount).Type1 * 100 / m_AgentList(nAgentCount).Count, "##0.00") & " %"
      grdList.TextMatrix(nIndex, 4) = m_AgentList(nAgentCount).Type2
      grdList.TextMatrix(nIndex, 5) = Format(m_AgentList(nAgentCount).Type2 * 100 / m_AgentList(nAgentCount).Count, "##0.00") & " %"
      grdList.TextMatrix(nIndex, 6) = m_AgentList(nAgentCount).Type3
      grdList.TextMatrix(nIndex, 7) = Format(m_AgentList(nAgentCount).Type3 * 100 / m_AgentList(nAgentCount).Count, "##0.00") & " %"
      grdList.TextMatrix(nIndex, 8) = m_AgentList(nAgentCount).Count
      grdList.TextMatrix(nIndex, 9) = Format(m_AgentList(nAgentCount).Count * 100 / nTotalAmount, "##0.00") & " %"
   Next nAgentCount
   
   'Added by Lydia 2022/02/18 MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉
   If grdList.Rows > 1 Then
      grdList.FixedRows = 1
   End If
   
   ClearAgentList
EXITSUB:
   If m_Recordset.RecordCount > 0 Then
      If m_Recordset.EOF = True Then: bottonNext.Enabled = False:
   Else
      bottonNext.Enabled = False
   End If
End Sub

Private Sub InitialGridList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 10
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "事務所"
   grdList.ColWidth(1) = 1000
   grdList.col = 2
   grdList.Text = "發明"
   grdList.ColWidth(2) = 800
   grdList.col = 3
   grdList.Text = "百分比"
   grdList.ColWidth(3) = 800
   grdList.col = 4
   grdList.Text = "新型"
   grdList.ColWidth(4) = 800
   grdList.col = 5
   grdList.Text = "百分比"
   grdList.ColWidth(5) = 800
   grdList.col = 6
   grdList.Text = "設計"
   grdList.ColWidth(6) = 800
   grdList.col = 7
   grdList.Text = "百分比"
   grdList.ColWidth(7) = 800
   grdList.col = 8
   grdList.Text = "小計"
   grdList.ColWidth(8) = 800
   grdList.col = 9
   grdList.Text = "百分比"
   grdList.ColWidth(8) = 800
End Sub

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

Private Sub Form_Unload(Cancel As Integer)
'Add By Cheng 2002/07/18
Set frm04060105_2 = Nothing
End Sub

Private Sub grdList_SelChange()
   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = grdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = grdList.row Then
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      If grdList.CellBackColor <> &H80000005 Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
            If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
         Next nCol
      End If
      grdList.col = 0
   End If
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      For nCol = 1 To grdList.Cols - 1
         grdList.col = nCol
         grdList.CellBackColor = &H8000000D
         grdList.CellForeColor = &H80000005
      Next nCol
      grdList.col = 0
   End If
EXITSUB:
End Sub

