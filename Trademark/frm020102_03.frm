VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm020102_03 
   BorderStyle     =   1  '單線固定
   Caption         =   "內商-發文(馬德里案)"
   ClientHeight    =   5748
   ClientLeft      =   5136
   ClientTop       =   3900
   ClientWidth     =   9336
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9336
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   5052
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   9132
      _ExtentX        =   16108
      _ExtentY        =   8911
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
   Begin VB.CommandButton ButtonOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6384
      TabIndex        =   2
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton ButtonPrev 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7212
      TabIndex        =   1
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton ButtonExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   8436
      TabIndex        =   0
      Top             =   70
      Width           =   800
   End
End
Attribute VB_Name = "frm020102_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/13 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Amy 2021/12/23 Form2.0已修改 grdList
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 來函收文日
Dim m_CP05 As String
' 所選取的收文號
Dim m_CP09 As String
' 申請國家
Dim m_TM10 As String
' 不列出的收文號
Dim m_NOCP09 As String
'
Dim m_CurrSel As Integer

Private Sub ButtonExit_Click()
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm020102_01
   Unload Me
End Sub

Private Sub ButtonOK_Click()
   frm020102_01.SetData 0, m_TM01, True
   frm020102_01.SetData 1, m_TM02, False
   frm020102_01.SetData 2, m_TM03, False
   frm020102_01.SetData 3, m_TM04, False
   frm020102_01.SetQueryFromTM
   Unload Me
   frm020102_01.Show
   frm020102_01.radio(1).Value = True
   frm020102_01.radio_Click 1
   frm020102_01.QueryData
End Sub

Private Sub ButtonPrev_Click()
   Unload Me
   frm020102_01.Show
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
   End If
   
   Select Case nType
      ' 本所案號 欄位1
      Case 0: m_TM01 = strData
      ' 本所案號 欄位2
      Case 1: m_TM02 = strData
      ' 本所案號 欄位3
      Case 2: m_TM03 = strData
      ' 本所案號 欄位4
      Case 3: m_TM04 = strData
   End Select
End Sub

' 查詢資料庫取得資料
Public Sub QueryTradeMark()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim nIndex As Integer
   
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         grdList.Rows = grdList.Rows + 1
         nIndex = grdList.Rows - 1
         
         ' 本所案號
         grdList.TextMatrix(nIndex, 1) = rsTmp.Fields("TM01") & "-" & rsTmp.Fields("TM02") & "-" & rsTmp.Fields("TM03") & "-" & rsTmp.Fields("TM04")
         
         ' 案件名稱
         If IsNull(rsTmp.Fields("TM05")) = False Then
            grdList.TextMatrix(nIndex, 2) = rsTmp.Fields("TM05")
         ElseIf IsNull(rsTmp.Fields("TM06")) = False Then
            grdList.TextMatrix(nIndex, 2) = rsTmp.Fields("TM06")
         ElseIf IsNull(rsTmp.Fields("TM07")) = False Then
            grdList.TextMatrix(nIndex, 2) = rsTmp.Fields("TM07")
         End If
         
         ' 申請國家
         If IsNull(rsTmp.Fields("TM10")) = False Then
            grdList.TextMatrix(nIndex, 3) = GetNationName(rsTmp.Fields("TM10"), 0)
         End If
         
         ' 商品類別
         If IsNull(rsTmp.Fields("TM09")) = False Then
            grdList.TextMatrix(nIndex, 4) = GetTradeMarkName(rsTmp.Fields("TM09"), 0)
         End If
         
         If IsNull(rsTmp.Fields("TM01")) = False Then
            grdList.TextMatrix(nIndex, 5) = rsTmp.Fields("TM01")
         End If
         If IsNull(rsTmp.Fields("TM02")) = False Then
            grdList.TextMatrix(nIndex, 6) = rsTmp.Fields("TM02")
         End If
         If IsNull(rsTmp.Fields("TM03")) = False Then
            grdList.TextMatrix(nIndex, 7) = rsTmp.Fields("TM03")
         End If
         If IsNull(rsTmp.Fields("TM04")) = False Then
            grdList.TextMatrix(nIndex, 8) = rsTmp.Fields("TM04")
         End If
         
         ' 下一筆
         rsTmp.MoveNext
      Loop
      grdList.FixedRows = 1 'Added by Lydia 2023/10/13
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 查詢資料庫取得資料
Public Sub QueryServicePractice()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim nIndex As Integer
   
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         grdList.Rows = grdList.Rows + 1
         nIndex = grdList.Rows - 1
         
         ' 本所案號
         grdList.TextMatrix(nIndex, 1) = rsTmp.Fields("SP01") & "-" & rsTmp.Fields("SP02") & "-" & rsTmp.Fields("SP03") & "-" & rsTmp.Fields("SP04")
         
         ' 案件名稱
         If IsNull(rsTmp.Fields("SP05")) = False Then
            grdList.TextMatrix(nIndex, 2) = rsTmp.Fields("SP05")
         ElseIf IsNull(rsTmp.Fields("SP06")) = False Then
            grdList.TextMatrix(nIndex, 2) = rsTmp.Fields("SP06")
         ElseIf IsNull(rsTmp.Fields("SP07")) = False Then
            grdList.TextMatrix(nIndex, 2) = rsTmp.Fields("SP07")
         End If
         
         ' 申請國家
         If IsNull(rsTmp.Fields("SP09")) = False Then
            grdList.TextMatrix(nIndex, 3) = GetNationName(rsTmp.Fields("SP09"), 0)
         End If
         
         If IsNull(rsTmp.Fields("SP01")) = False Then
            grdList.TextMatrix(nIndex, 5) = rsTmp.Fields("SP01")
         End If
         If IsNull(rsTmp.Fields("SP02")) = False Then
            grdList.TextMatrix(nIndex, 6) = rsTmp.Fields("SP02")
         End If
         If IsNull(rsTmp.Fields("SP03")) = False Then
            grdList.TextMatrix(nIndex, 7) = rsTmp.Fields("SP03")
         End If
         If IsNull(rsTmp.Fields("SP04")) = False Then
            grdList.TextMatrix(nIndex, 8) = rsTmp.Fields("SP04")
         End If
         
         ' 下一筆
         rsTmp.MoveNext
      Loop
      grdList.FixedRows = 1 'Added by Lydia 2023/10/13
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub


' 查詢資料庫取得資料
Public Sub QueryData()
   InitialGrdList
   
   Select Case m_TM01
      Case "T", "TF", "FCT":
         QueryTradeMark
      Case Else:
         QueryServicePractice
   End Select
   
End Sub

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 9
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "本所案號"
   grdList.ColWidth(1) = 1200
   grdList.col = 2
   grdList.Text = "案件名稱"
   grdList.ColWidth(2) = 2000
   grdList.col = 3
   grdList.Text = "申請國家"
   grdList.ColWidth(3) = 1200
   grdList.col = 4
   grdList.Text = "商品類別"
   grdList.ColWidth(4) = 1000
   grdList.col = 5
   grdList.Text = "本所案號第一欄"
   grdList.ColWidth(5) = 0
   grdList.col = 6
   grdList.Text = "本所案號第二欄"
   grdList.ColWidth(6) = 0
   grdList.col = 7
   grdList.Text = "本所案號第三欄"
   grdList.ColWidth(7) = 0
   grdList.col = 8
   grdList.Text = "本所案號第四欄"
   grdList.ColWidth(8) = 0
End Sub

' 設定Grid List的一列為選取的狀態
Private Sub grdList_SetSelection(ByVal nSel As Integer)
   If nSel > 0 And nSel < grdList.Rows And grdList.Rows >= 2 Then
      grdList.row = nSel
      grdList_SelChange
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'edit by nickc 2008/04/25 改整批印
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
    'Add By Cheng 2002/07/18
   Set frm020102_03 = Nothing
End Sub

Private Sub grdList_SelChange()
   If grdList.Rows > 1 Then
      If grdList.row > 0 Then
         m_TM01 = grdList.TextMatrix(grdList.row, 5)
         m_TM02 = grdList.TextMatrix(grdList.row, 6)
         m_TM03 = grdList.TextMatrix(grdList.row, 7)
         m_TM04 = grdList.TextMatrix(grdList.row, 8)
      End If
   End If
   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = grdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = grdList.row Then
      Dim nOldCol As Integer
      nOldCol = grdList.col
      grdList.col = 1
      If grdList.CellBackColor <> &H8000000D Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H8000000D Then grdList.CellBackColor = &H8000000D
            If grdList.CellForeColor <> &H80000005 Then grdList.CellForeColor = &H80000005
         Next nCol
      End If
      grdList.col = nOldCol
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      If grdList.CellBackColor <> &H80000005 Then
         For nCol = 1 To grdList.Cols - 1
            If grdList.TextMatrix(grdList.row, 8) = "1" Then
               If grdList.CellBackColor <> &HFF& Then: grdList.CellBackColor = &HFF&
               If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
            Else
               grdList.col = nCol
               If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
               If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
            End If
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



