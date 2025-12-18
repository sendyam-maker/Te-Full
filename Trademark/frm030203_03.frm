VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm030203_03 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標申請案號輸入"
   ClientHeight    =   5712
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9312
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5712
   ScaleWidth      =   9312
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7185
      TabIndex        =   2
      Top             =   60
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6360
      TabIndex        =   1
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8400
      TabIndex        =   0
      Top             =   60
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   4632
      Left            =   72
      TabIndex        =   4
      Top             =   1008
      Width           =   9132
      _ExtentX        =   16108
      _ExtentY        =   8170
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   150
      TabIndex        =   3
      Top             =   645
      Width           =   75
   End
End
Attribute VB_Name = "frm030203_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/17 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Lydia 2021/09/11 改成Form2.0 ; grdList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
' 申請案號
Dim m_TM12 As String
' 審定號
Dim m_TM15 As String
' 來函收文日
Dim m_CP05 As String

Private Sub cmdCancel_Click()
   Unload Me
   frm030203_01.Show
End Sub

Private Sub cmdExit_Click()
   Unload frm030203_01
   Unload Me
End Sub


Private Sub cmdOK_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If grdList.Rows > 0 Then
      If IsEmptyText(m_CP01) = True Or IsEmptyText(m_CP02) = True Or IsEmptyText(m_CP03) = True Or IsEmptyText(m_CP04) = True Then
         strMsg = "請先選取一筆記錄"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   Else
      strMsg = "無符合的資料"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   frm030203_01.textTM01 = m_CP01
   frm030203_01.textTM02 = m_CP02
   frm030203_01.textTM03 = m_CP03
   frm030203_01.textTM04 = m_CP04
   frm030203_01.Show
   frm030203_01.cmdOK_Click
   Unload Me
EXITSUB:
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   
   MoveFormToCenter Me
End Sub

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 11
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "本所案號"
   grdList.ColWidth(1) = 1200
   grdList.col = 2
   grdList.Text = "商標名稱"
   grdList.ColWidth(2) = 1000
   grdList.col = 3
   grdList.Text = "商標種類"
   grdList.ColWidth(3) = 1200
   grdList.col = 4
   grdList.Text = "商品類別"
   grdList.ColWidth(4) = 1000
   grdList.col = 5
   grdList.Text = "申請人"
   grdList.ColWidth(5) = 1200
   grdList.col = 6
   grdList.Text = "申請國家"
   grdList.ColWidth(6) = 1200
   ' 本所案號 欄位一
   grdList.col = 7
   grdList.Text = Empty
   grdList.ColWidth(7) = 0
   ' 本所案號 欄位二
   grdList.col = 8
   grdList.Text = Empty
   grdList.ColWidth(8) = 0
   ' 本所案號 欄位三
   grdList.col = 9
   grdList.Text = Empty
   grdList.ColWidth(9) = 0
   ' 本所案號 欄位四
   grdList.col = 10
   grdList.Text = Empty
   grdList.ColWidth(10) = 0
End Sub

' 列出所有資料
Private Sub ListData(ByRef rsTmp As ADODB.Recordset)
   Dim strNationCode As String
   
   InitialGrdList
   rsTmp.MoveFirst
   Do While rsTmp.EOF = False
      strNationCode = Empty
      grdList.Rows = grdList.Rows + 1
      grdList.row = grdList.Rows - 1
      ' 本所案號欄位
      grdList.TextMatrix(grdList.row, 1) = rsTmp.Fields("TM01") & rsTmp.Fields("TM02") & rsTmp.Fields("TM03") & rsTmp.Fields("TM04")
      ' 商標名稱欄位
      If IsNull(rsTmp.Fields("TM05")) = False Then
         grdList.TextMatrix(grdList.row, 2) = rsTmp.Fields("TM05")
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         strNationCode = rsTmp.Fields("TM10")
         grdList.TextMatrix(grdList.row, 6) = GetNationName(rsTmp.Fields("TM10"), 0)
      End If
      ' 商標種類欄位
      If IsNull(rsTmp.Fields("TM08")) = False Then
         If strNationCode < "010" Then
            grdList.TextMatrix(grdList.row, 3) = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
         Else
            grdList.TextMatrix(grdList.row, 3) = GetTradeMarkName(rsTmp.Fields("TM08"), 1)
         End If
      End If
      ' 商品類別
      If IsNull(rsTmp.Fields("TM09")) = False Then
         grdList.TextMatrix(grdList.row, 4) = rsTmp.Fields("TM09")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         grdList.TextMatrix(grdList.row, 5) = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      ' 隱藏起來的本所案號
      grdList.TextMatrix(grdList.row, 7) = rsTmp.Fields("TM01")
      grdList.TextMatrix(grdList.row, 8) = rsTmp.Fields("TM02")
      grdList.TextMatrix(grdList.row, 9) = rsTmp.Fields("TM03")
      grdList.TextMatrix(grdList.row, 10) = rsTmp.Fields("TM04")
               
      rsTmp.MoveNext
   Loop
   
   'Added by Lydia 2023/10/17
   If grdList.Rows >= 2 Then
      grdList.FixedRows = 1
   End If
   'end 2023/10/17
      
   ' 設定第一列為選取的狀態
   grdList_SetSelection 1
   
   ' 當只有一筆記錄時, 則直接跳至下一畫面
   If grdList.Rows = 2 Then
        Unload Me
        frm02010301_1.Show
   Else
      Me.Show
   End If
End Sub

' 搜尋資料
Public Sub QueryData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim nType As Integer
   Dim strTM01 As String
   Dim strTM02 As String
   Dim strTM03 As String
   Dim strTM04 As String
     strTM01 = m_TM01
     strTM02 = m_TM02
     strTM03 = m_TM03
     strTM04 = m_TM04
    'Modify By Cheng 2004/02/04
    Select Case strTM01
    Case "T", "FCT", "TF"
        Select Case strTM01
           Case "T", "FCT":
              If IsEmptyText(strTM03) = True Then
                 strTM03 = "0"
              End If
              If IsEmptyText(strTM04) = True Then
                 strTM04 = "00"
              End If
              ' 設定SQL語法
              strSql = "SELECT * FROM TradeMark,divisioncase " & _
                 "WHERE dc05 = '" & strTM01 & "' AND " & _
                       "dc06 = '" & strTM02 & "' AND " & _
                       "dc07 = '" & strTM03 & "' AND " & _
                       "dc08 = '" & strTM04 & "' and dc01=tm01(+) and dc02=tm02(+) and dc03=tm03(+) and dc04=tm04(+) " & _
                 "ORDER BY TM01||TM02||TM03||TM04 "
           Case "TF":
              ' 設定SQL語法
              strSql = "SELECT * FROM TradeMark,divisioncase " & _
                 "WHERE dc05 = '" & strTM01 & "' AND " & _
                       "dc06 = '" & strTM02 & "' "
              If IsEmptyText(strTM03) = False Then
                 strSql = strSql & "AND "
                 strSql = strSql & "dc07 = '" & strTM03 & "' "
              End If
              If IsEmptyText(strTM04) = False Then
                 strSql = strSql & "AND "
                 strSql = strSql & "dc08 = '" & strTM04 & "' "
              End If
              strSql = strSql & " and dc01=tm01(+) and dc02=tm02(+) and dc03=tm03(+) and dc04=tm04(+)   ORDER BY TM01||TM02||TM03||TM04 "
        End Select
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If rsTmp.RecordCount > 0 Then
           ListData rsTmp
        End If
        rsTmp.Close
    Case "TC"
        If IsEmptyText(strTM03) = True Then
           strTM03 = "0"
        End If
        If IsEmptyText(strTM04) = True Then
           strTM04 = "00"
        End If
        ' 設定SQL語法
        strSql = "SELECT * FROM Servicepractice,divisioncase " & _
           "WHERE dc05 = '" & strTM01 & "' AND " & _
                 "dc06 = '" & strTM02 & "' AND " & _
                 "dc07 = '" & strTM03 & "' AND " & _
                 "dc08 = '" & strTM04 & "' and and dc01=sp01(+) and dc02=sp02(+) and dc03=sp03(+) and dc04=sp04(+) " & _
           "ORDER BY SP01||SP02||SP03||SP04 "
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If rsTmp.RecordCount > 0 Then
           ListData_1 rsTmp
        End If
        rsTmp.Close
    
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm030203_03 = Nothing
End Sub

Private Sub grdList_SelChange()
   If grdList.row > 0 Then
      grdList.col = 7
      m_CP01 = grdList.Text
      grdList.col = 8
      m_CP02 = grdList.Text
      grdList.col = 9
      m_CP03 = grdList.Text
      grdList.col = 10
      m_CP04 = grdList.Text
   End If
   grdList_ShowSelection
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
         grdList.col = 1
         If grdList.CellBackColor <> &H80000005 Then
            For nCol = 1 To grdList.Cols - 1
               grdList.col = nCol
               If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
               If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
            Next nCol
         End If
      End If
   Next nRow
   grdList.row = nCurrSel
   grdList.col = 0
End Sub
' 設定Grid List的一列為選取的狀態
Private Sub grdList_SetSelection(ByVal nSel As Integer)
   If grdList.Rows >= 2 Then
      grdList.row = 1
      grdList_SelChange
      grdList_ShowSelection
   End If
End Sub

'Add By Cheng 2004/02/04
' 列出所有資料
Private Sub ListData_1(ByRef rsTmp As ADODB.Recordset)
   Dim strNationCode As String
   
   InitialGrdList
   rsTmp.MoveFirst
   Do While rsTmp.EOF = False
      strNationCode = Empty
      grdList.Rows = grdList.Rows + 1
      grdList.row = grdList.Rows - 1
      ' 本所案號欄位
      grdList.TextMatrix(grdList.row, 1) = rsTmp.Fields("SP01") & rsTmp.Fields("SP02") & rsTmp.Fields("SP03") & rsTmp.Fields("SP04")
      ' 商標名稱欄位
      If IsNull(rsTmp.Fields("SP05")) = False Then
         grdList.TextMatrix(grdList.row, 2) = rsTmp.Fields("SP05")
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         strNationCode = rsTmp.Fields("SP09")
         grdList.TextMatrix(grdList.row, 6) = GetNationName(rsTmp.Fields("SP09"), 0)
      End If
      ' 商標種類欄位
        grdList.ColWidth(3) = 0
      ' 商品類別
        grdList.ColWidth(4) = 0
      ' 申請人
      If IsNull(rsTmp.Fields("SP08")) = False Then
         grdList.TextMatrix(grdList.row, 5) = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
      ' 隱藏起來的本所案號
      grdList.TextMatrix(grdList.row, 7) = rsTmp.Fields("SP01")
      grdList.TextMatrix(grdList.row, 8) = rsTmp.Fields("SP02")
      grdList.TextMatrix(grdList.row, 9) = rsTmp.Fields("SP03")
      grdList.TextMatrix(grdList.row, 10) = rsTmp.Fields("SP04")
               
      rsTmp.MoveNext
   Loop
   
   'Added by Lydia 2023/10/17
   If grdList.Rows >= 2 Then
      grdList.FixedRows = 1
   End If
   'end 2023/10/17
   
   ' 設定第一列為選取的狀態
   grdList_SetSelection 1
   
   Me.Show
End Sub



Public Sub SetData(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String)
   m_TM01 = strTM01
   m_TM02 = strTM02
   m_TM03 = strTM03
   m_TM04 = strTM04
   Label1.Caption = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & "   有分割案"
End Sub


