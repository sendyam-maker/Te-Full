VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010403_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "審查報告輸入"
   ClientHeight    =   5736
   ClientLeft      =   96
   ClientTop       =   996
   ClientWidth     =   9324
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   9324
   Begin VB.TextBox textResult 
      Height          =   264
      Left            =   720
      MaxLength       =   1
      TabIndex        =   0
      Top             =   5400
      Width           =   252
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8460
      TabIndex        =   3
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6408
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7236
      TabIndex        =   2
      Top             =   70
      Width           =   1200
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   960
      Width           =   2652
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   660
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   960
      Width           =   2532
   End
   Begin VB.TextBox textTM06 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1560
      Width           =   7632
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   2772
      Left            =   72
      TabIndex        =   20
      Top             =   2496
      Width           =   9072
      _ExtentX        =   16002
      _ExtentY        =   4890
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
   Begin MSForms.TextBox textTM05 
      Height          =   264
      Left            =   1560
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1260
      Width           =   7632
      VariousPropertyBits=   679493663
      Size            =   "13462;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM07 
      Height          =   264
      Left            =   1560
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1860
      Width           =   7632
      VariousPropertyBits=   679493663
      Size            =   "13462;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   264
      Left            =   1440
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2160
      Width           =   7752
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "13674;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      Caption         =   "(1:審查報告 2:核駁前先行通知 3.部分核駁)"
      Height          =   255
      Left            =   1080
      TabIndex        =   19
      Top             =   5400
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "結果 :"
      Height          =   252
      Left            =   120
      TabIndex        =   18
      Top             =   5400
      Width           =   492
   End
   Begin VB.Label Label1 
      Caption         =   "審定號 :"
      Height          =   252
      Index           =   1
      Left            =   4800
      TabIndex        =   17
      Top             =   960
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   660
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "申請案號 :"
      Height          =   252
      Index           =   13
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   1212
   End
   Begin VB.Label Label3 
      Caption         =   "案件中文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   14
      Top             =   1260
      Width           =   1452
   End
   Begin VB.Label Label4 
      Caption         =   "案件英文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   1452
   End
   Begin VB.Label Label5 
      Caption         =   "案件日文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   12
      Top             =   1860
      Width           =   1452
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   1332
   End
End
Attribute VB_Name = "frm02010403_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/18 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Amy 2021/12/28 Form2.0已修改 textTM05/textTm07/textTm23/grdList
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
' 來源畫面
Dim strPrevForm As String
' 申請國家
Dim m_TM10 As String
'Add By Sindy 2019/5/10
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2019/5/10 END


'Add By Sindy 2019/5/13
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdCancel_Click()
   Select Case strPrevForm
      Case "2"
         frm02010403_2.Show
         Unload Me
      Case Else
         frm02010403_1.Show
         Unload Me
         Unload frm02010403_2
   End Select
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move to unload by nick     2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm02010403_2
   Unload frm02010403_1
   Unload Me
End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If IsEmptyText(m_CP09) = False Then
      DisplayNextForm
   Else
      strMsg = "請先選取一筆記錄"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM05.BackColor = &H8000000F
   textTM06.BackColor = &H8000000F
   textTM07.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   ' 初始化
   Initial
   
   MoveFormToCenter Me
   
   'Add By Sindy 2019/5/10
   m_strIR01 = frm02010403_1.m_strIR01
   m_strIR02 = frm02010403_1.m_strIR02
   m_strIR03 = frm02010403_1.m_strIR03
   m_strIR04 = frm02010403_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2019/5/10 END
End Sub

Private Sub Initial()
   textResult = "1"
   'Added by Morgan 2017/4/17 電子公文
   Select Case frm02010403_1.m_NewCP10
      Case "1201"
         textResult = "1"
      Case "1202"
         textResult = "2"
      Case "1205"
         textResult = "3"
   End Select
   'end 2017/4/17
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP05 = Empty
      strPrevForm = Empty
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
      ' 來函收文日
      Case 4: m_CP05 = strData
      ' 來源畫面
      Case 5: strPrevForm = strData
   End Select
End Sub

Private Sub QueryTradeMark()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      ' 審定號
      If IsNull(rsTmp.Fields("TM15")) = False Then
         textTM15 = rsTmp.Fields("TM15")
      End If
      ' 商標名稱(中)
      If IsNull(rsTmp.Fields("TM05")) = False Then
         textTM05 = rsTmp.Fields("TM05")
      End If
      ' 商標名稱(英)
      If IsNull(rsTmp.Fields("TM06")) = False Then
         textTM06 = rsTmp.Fields("TM06")
      End If
      ' 商標名稱(日)
      If IsNull(rsTmp.Fields("TM07")) = False Then
         textTM07 = rsTmp.Fields("TM07")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Private Sub QueryServicePractice()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' AND " & _
                  "SP03 = '" & m_TM03 & "' AND " & _
                  "SP04 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("SP11")) = False Then
         textTM12 = rsTmp.Fields("SP11")
      End If
      ' 審定號
      textTM15 = Empty
      ' 商標名稱(中)
      If IsNull(rsTmp.Fields("SP05")) = False Then
         textTM05 = rsTmp.Fields("SP05")
      End If
      ' 商標名稱(英)
      If IsNull(rsTmp.Fields("SP06")) = False Then
         textTM06 = rsTmp.Fields("SP06")
      End If
      ' 商標名稱(日)
      If IsNull(rsTmp.Fields("SP07")) = False Then
         textTM07 = rsTmp.Fields("SP07")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("SP08")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 查詢資料庫取得資料
Public Sub QueryData()
   
   m_TM10 = Empty
   
   Select Case m_TM01
      Case "T", "TF", "FCT", "CFT":
         QueryTradeMark
      Case Else
         QueryServicePractice
   End Select
   
   ' 顯示符合條件的資料
   ListData
   
End Sub

' 列出案件進度表符合條件的資料
Private Sub ListData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim bDeal As Boolean
   Dim bSpecial As Boolean
   
   InitialGrdList
   
   m_CP09 = Empty
   
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & m_TM01 & "' AND " & _
                  "CP02 = '" & m_TM02 & "' AND " & _
                  "CP03 = '" & m_TM03 & "' AND " & _
                  "CP04 = '" & m_TM04 & "' " & _
            "ORDER BY CP05 DESC "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         bSpecial = False
         ' 無發文日不予列入
         If IsNull(rsTmp.Fields("CP27")) = True Then: GoTo NextRecord
         ' 發文日是空白不予列入
         If IsEmptyText(rsTmp.Fields("CP27")) = True Then: GoTo NextRecord
         ' 無收文號不予列入
         If IsNull(rsTmp.Fields("CP09")) = True Then: GoTo NextRecord
         ' 收文號不為A,B類的不予計入
         Select Case Mid(rsTmp.Fields("CP09"), 1, 1)
            Case "A", "B":
            Case Else: GoTo NextRecord
         End Select
         ' 表列出無實際結果的資料
         If IsNull(rsTmp.Fields("CP24")) = False Then
            If IsEmptyText(rsTmp.Fields("CP24")) = False Then
               GoTo NextRecord
            End If
         End If
         
         ' 列入資料
         grdList.Rows = grdList.Rows + 1
         grdList.row = grdList.Rows - 1
         ' 收文號
         If IsNull(rsTmp.Fields("CP09")) = False Then
            grdList.TextMatrix(grdList.row, 1) = rsTmp.Fields("CP09")
         End If
         ' 收文日
         If IsNull(rsTmp.Fields("CP05")) = False Then
            grdList.TextMatrix(grdList.row, 2) = TAIWANDATE(rsTmp.Fields("CP05"))
         End If
         ' 案件性質
         If IsNull(rsTmp.Fields("CP10")) = False Then
            If m_TM10 = "000" Then
               grdList.TextMatrix(grdList.row, 3) = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
            Else
               grdList.TextMatrix(grdList.row, 3) = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
            End If
         End If
         ' 發文日
         If IsNull(rsTmp.Fields("CP27")) = False Then
            grdList.TextMatrix(grdList.row, 4) = TAIWANDATE(rsTmp.Fields("CP27"))
         End If
         ' 結果
         If IsNull(rsTmp.Fields("CP24")) = False Then
            Select Case rsTmp.Fields("CP24")
               Case "1":
                  grdList.TextMatrix(grdList.row, 5) = "准/勝"
               Case "2":
                  grdList.TextMatrix(grdList.row, 5) = "駁/敗"
               Case Else:
            End Select
         End If
         ' 相關人
         bDeal = False
         If bDeal = False And IsNull(rsTmp.Fields("CP40")) = False Then
            If IsEmptyText(rsTmp.Fields("CP40")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP40")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP41")) = False Then
            If IsEmptyText(rsTmp.Fields("CP41")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP41")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP42")) = False Then
            If IsEmptyText(rsTmp.Fields("CP42")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP42")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP50")) = False Then
            If IsEmptyText(rsTmp.Fields("CP50")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP50")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP51")) = False Then
            If IsEmptyText(rsTmp.Fields("CP51")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP51")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP52")) = False Then
            If IsEmptyText(rsTmp.Fields("CP52")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP52")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP56")) = False Then
            If IsEmptyText(rsTmp.Fields("CP56")) = False Then
               grdList.TextMatrix(grdList.row, 6) = GetCustomerName(rsTmp.Fields("CP56"), 0)
               bDeal = True
            End If
         End If
         If bDeal = False Then: grdList.Text = Empty
         ' 特殊欄位
         If bSpecial = True Then
            grdList.TextMatrix(grdList.row, 7) = "1"
         End If
NextRecord:
         rsTmp.MoveNext
      Loop
      'Added by Lydia 2023/10/18
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/18
   End If
   
   ' 設定第一列為所選取的記錄
   grdList_SetSelection 1

End Sub

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 8
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "收文號"
   grdList.ColWidth(1) = 1000
   grdList.col = 2
   grdList.Text = "收文日"
   grdList.ColWidth(2) = 800
   grdList.col = 3
   grdList.Text = "案件性質"
   grdList.ColWidth(3) = 1200
   grdList.col = 4
   grdList.Text = "發文日"
   grdList.ColWidth(4) = 800
   grdList.col = 5
   grdList.Text = "結果"
   grdList.ColWidth(5) = 800
   grdList.col = 6
   grdList.Text = "相關人"
   grdList.ColWidth(6) = 1200
   grdList.col = 7
   grdList.Text = "特殊"
   grdList.ColWidth(7) = 0
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nRow As Integer
   Dim nCol As Integer
   Dim nCurrSel As Integer
   nCurrSel = grdList.row
   For nRow = 1 To grdList.Rows - 1
      grdList.row = nRow
      If nRow = nCurrSel Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            grdList.CellBackColor = &H8000000D
            grdList.CellForeColor = &H80000005
         Next nCol
      Else
         If grdList.Cols = 8 Then
            grdList.col = 7
            Select Case grdList.Text
               Case 1:
                  grdList.col = 1
                  If grdList.CellBackColor <> &HFF& Then
                     For nCol = 1 To grdList.Cols - 1
                        grdList.col = nCol
                        If grdList.CellBackColor <> &HFF& Then: grdList.CellBackColor = &HFF&
                        If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
                     Next nCol
                  End If
               Case Else:
                  grdList.col = 1
                  If grdList.CellBackColor <> &H80000005 Then
                     For nCol = 1 To grdList.Cols - 1
                        grdList.col = nCol
                        If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
                        If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
                     Next nCol
                  End If
            End Select
         Else
            grdList.col = 1
            If grdList.CellBackColor <> &H80000005 Then
               For nCol = 1 To grdList.Cols - 1
                  grdList.col = nCol
                  If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
                  If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
               Next nCol
            End If
         End If
      End If
   Next nRow
   grdList.row = nCurrSel
   grdList.col = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
'edit by nickc 2008/04/25 改整批印
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   
   'Add By Sindy 2019/5/13
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   
   'Add By Cheng 2002/07/18
   Set frm02010403_3 = Nothing
End Sub

Private Sub grdList_SelChange()
   If grdList.row > 0 Then
      grdList.col = 1
      m_CP09 = grdList.Text
   End If
   grdList_ShowSelection
End Sub

' 設定Grid List的一列為選取的狀態
Private Sub grdList_SetSelection(ByVal nSel As Integer)
   If grdList.Rows >= 2 Then
      grdList.row = nSel
      grdList_SelChange
      grdList_ShowSelection
   End If
End Sub

Private Sub textResult_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textResult) = False Then
      Select Case textResult
         Case "1", "2", "3":
            '92.8.1 ADD BY SONIA
            If m_TM10 = "000" And textResult = "3" Then
               Cancel = True
               strTit = "資料檢核"
               strMsg = "國內案結果不可以輸入３"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textResult_GotFocus
               GoTo EXITSUB
            End If
            '92.8.1 END
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "請輸入 1 或 2 或 ３"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textResult_GotFocus
            GoTo EXITSUB
      End Select
   Else
      textResult = "1"
   End If
   'ListData    '2009/10/15 CANCEL BY SONIA 不必重新讀資料,否則先點選好收文號才改此欄且直接按確定時都會帶到第一筆而非點選資料
EXITSUB:
End Sub

Private Sub DisplayNextForm()
   frm02010403_4.SetData 0, m_TM01, True
   frm02010403_4.SetData 1, m_TM02, False
   frm02010403_4.SetData 2, m_TM03, False
   frm02010403_4.SetData 3, m_TM04, False
   frm02010403_4.SetData 4, m_CP05, False
   frm02010403_4.SetData 5, m_CP09, False
   'Added by Morgan 2017/4/17 電子公文
   frm02010403_4.m_DocWord = frm02010403_1.m_DocWord
   frm02010403_4.m_DocNo = frm02010403_1.m_DocNo
   frm02010403_4.m_AppNo = frm02010403_1.m_AppNo
   frm02010403_4.m_DeadLine = frm02010403_1.m_DeadLine
   'end 2017/4/17
   'Add By Sindy 2019/5/10
   If Not m_PrevForm Is Nothing Then
      Call frm02010403_4.SetParent(m_PrevForm)
   End If
   frm02010403_4.m_strIR01 = m_strIR01
   frm02010403_4.m_strIR02 = m_strIR02
   frm02010403_4.m_strIR03 = m_strIR03
   frm02010403_4.m_strIR04 = m_strIR04
   '2019/5/10 END
   Me.Hide
   frm02010403_4.Show
   frm02010403_4.QueryData
End Sub

Public Function GetSelectResult() As String
   GetSelectResult = textResult
End Function

Private Sub textResult_GotFocus()
   InverseTextBox textResult
End Sub

