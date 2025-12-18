VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm030614 
   BorderStyle     =   1  '單線固定
   Caption         =   "國內商標公報查詢"
   ClientHeight    =   5748
   ClientLeft      =   180
   ClientTop       =   972
   ClientWidth     =   9336
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9336
   Begin VB.CheckBox Check1 
      Caption         =   "只顯示本所案件"
      Height          =   315
      Left            =   3660
      TabIndex        =   3
      Top             =   810
      Width           =   1995
   End
   Begin VB.TextBox textTA03 
      Height          =   264
      Index           =   2
      Left            =   1830
      MaxLength       =   1
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox textTA03 
      Height          =   264
      Index           =   1
      Left            =   3000
      MaxLength       =   5
      TabIndex        =   1
      Top             =   510
      Width           =   885
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7320
      TabIndex        =   4
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   8280
      TabIndex        =   5
      Top             =   60
      Width           =   912
   End
   Begin VB.TextBox textTA03 
      Height          =   264
      Index           =   0
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   0
      Top             =   510
      Width           =   885
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   3888
      Left            =   72
      TabIndex        =   10
      Top             =   1248
      Width           =   9132
      _ExtentX        =   16108
      _ExtentY        =   6858
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
      Caption         =   "類別筆數：  筆"
      Height          =   180
      Index           =   3
      Left            =   180
      TabIndex        =   9
      Top             =   5460
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "公報筆數：  筆"
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   8
      Top             =   5220
      Width           =   1170
   End
   Begin VB.Label Label1 
      Caption         =   "是否顯示明細資料：             (Y : 顯示)"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   870
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "公報卷期：起                        －  迄"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   510
      Width           =   2745
   End
End
Attribute VB_Name = "frm030614"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/16 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Amy 2022/01/10 Form2.0已修改 grdList
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

Dim m_CurrSel As Integer
Dim m_dblKindCnt As Double '類別筆數


' 使用者按下離開按紐
Private Sub cmdExit_Click()
   Unload Me
End Sub

' 使用者按下查詢按紐
Private Sub cmdQuery_Click()
   If Len(textTA03(0).Text) <= 0 Then
      MsgBox "請輸入欲查詢的公報起始卷期!!!", vbExclamation
      textTA03(0).SetFocus
      textTA03_GotFocus 0
      Exit Sub
   End If
   If Len(textTA03(1).Text) <= 0 Then
      MsgBox "請輸入欲查詢的公報終止卷期!!!", vbExclamation
      textTA03(1).SetFocus
      textTA03_GotFocus 1
      Exit Sub
   End If
   If CheckKeyIn(1) = -1 Then
      textTA03(1).SetFocus
      textTA03_GotFocus 1
      Exit Sub
   End If
   If CheckKeyIn(2) = -1 Then
      textTA03(2).SetFocus
      textTA03_GotFocus 2
      Exit Sub
   End If
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/21 清除查詢印表記錄檔欄位
   QueryDB
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   InitialGrdList
End Sub

' 查詢資料庫
Private Sub QueryDB()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strCon As String
   
    Screen.MousePointer = vbHourglass
    grdList.Visible = False: DoEvents
    InitialGrdList
    If Len(textTA03(0)) <> 0 Or Len(textTA03(1)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Left(Label1(0), 5) & textTA03(0) & "-" & textTA03(1) 'Add By Sindy 2010/10/21
    End If
    
    'Add By Sindy 2011/12/6
    If Check1.Value = 0 Then '全部顯示
      strCon = " and TMBM01=tm15(+) "
    Else '只顯示本所案件
      strCon = " and TMBM01=tm15 and tm28='1' "
      pub_QL05 = pub_QL05 & ";只顯示本所案件" 'Add By Sindy 2010/10/21
    End If
    
    '顯示明細
    'Modify By Sindy 2011/12/6 +本所案號
    If textTA03(2).Text = "Y" Then
        pub_QL05 = pub_QL05 & ";" & Left(Label1(1), 9) & "Y" 'Add By Sindy 2010/10/21
        ' 設定查詢的與法
        strSql = "SELECT TMBM01 AS 審定號數,TMBM02 AS 商標種類,TMBM03 AS 正商標號數,TMBM04 AS 申請案號,TMBM05 AS 地區,TMBM06 AS 代理人,TMBM07 AS 卷期,TMBM08 AS 商品類別,decode(TM01,null,' ',decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04) AS 本所案號 " & _
                        " FROM  TMBULLETIN,trademark " & _
                        " WHERE TMBM07 >= '" & textTA03(0).Text & "'" & _
                        " And   TMBM07 <= '" & textTA03(1).Text & "'" & strCon & _
                        " Order By TMBM01,TMBM02 "
    '不顯示明細
    Else
        pub_QL05 = pub_QL05 & ";" & Left(Label1(1), 9) & "N" 'Add By Sindy 2010/10/21
        ' 設定查詢的與法
'        strSQL = "SELECT Nvl(COUNT(*), 0) AS 筆數 " & _
'                        " FROM  TMBULLETIN " & _
'                        " WHERE TMBM07 >= '" & textTA03(0).Text & "'" & _
'                        " And   TMBM07 <= '" & textTA03(1).Text & "'"
        strSql = "SELECT TMBM01 AS 審定號數,TMBM02 AS 商標種類,TMBM03 AS 正商標號數,TMBM04 AS 申請案號,TMBM05 AS 地區,TMBM06 AS 代理人,TMBM07 AS 卷期,TMBM08 AS 商品類別,decode(TM01,null,' ',decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04) AS 本所案號 " & _
                        " FROM  TMBULLETIN,trademark " & _
                        " WHERE TMBM07 >= '" & textTA03(0).Text & "'" & _
                        " And   TMBM07 <= '" & textTA03(1).Text & "'" & strCon & _
                        " Order By TMBM01,TMBM02 "
    End If

    ' 查詢資料庫
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    '顯示明細
    If textTA03(2).Text = "Y" Then
        If rsTmp.RecordCount <= 0 Then
            InsertQueryLog (0) 'Add By Sindy 2010/10/21
            Label1(2).Caption = "公報筆數： 共 " & "0" & " 筆"
            Label1(3).Caption = "類別筆數： 共 " & "0" & " 筆"
        Else
            InsertQueryLog (rsTmp.RecordCount) 'Add By Sindy 2010/10/21
            ListData rsTmp
        End If
    '不顯示明細
    Else
        '若無資料
        If rsTmp.RecordCount <= 0 Then
            InsertQueryLog (0) 'Add By Sindy 2010/10/21
            Label1(2).Caption = "公報筆數： 共 " & "0" & " 筆"
            Label1(3).Caption = "類別筆數： 共 " & "0" & " 筆"
        '若有資料
        Else
            InsertQueryLog (rsTmp.RecordCount) 'Add By Sindy 2010/10/21
            ListData rsTmp
        End If
    End If

    If rsTmp.State <> adStateClosed Then rsTmp.Close
EXITSUB:
    Set rsTmp = Nothing
    grdList.Visible = True: DoEvents
    Screen.MousePointer = vbDefault
End Sub

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 9 '8
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 0
   grdList.Text = "審定號數"
   grdList.ColWidth(0) = 1000
   grdList.ColAlignment(0) = flexAlignLeftCenter
   
   grdList.col = 1
   grdList.Text = "商標種類"
   grdList.ColWidth(1) = 800
   grdList.ColAlignment(1) = flexAlignLeftCenter
   
   grdList.col = 2
   grdList.Text = "正商標號數"
   grdList.ColWidth(2) = 1000
   grdList.ColAlignment(2) = flexAlignLeftCenter
   
   grdList.col = 3
   grdList.Text = "申請案號"
   grdList.ColWidth(3) = 1000
   grdList.ColAlignment(3) = flexAlignLeftCenter
   
   grdList.col = 4
   grdList.Text = "地區"
   grdList.ColWidth(4) = 1000
   grdList.ColAlignment(4) = flexAlignLeftCenter
   
   grdList.col = 5
   grdList.Text = "代理人"
   grdList.ColWidth(5) = 800
   grdList.ColAlignment(5) = flexAlignLeftCenter
   
   
   grdList.col = 6
   grdList.Text = "卷期"
   grdList.ColWidth(6) = 600
   grdList.ColAlignment(6) = flexAlignLeftCenter
   
   grdList.col = 7
   grdList.Text = "商品類別"
   grdList.ColWidth(7) = 1000
   grdList.ColAlignment(7) = flexAlignLeftCenter
   
   'Add By Sindy 2011/12/6
   grdList.col = 8
   grdList.Text = "本所案號"
   grdList.ColWidth(8) = 1500
   grdList.ColAlignment(8) = flexAlignLeftCenter
End Sub

Private Sub ListData(ByRef rsTmp As ADODB.Recordset)
'   Dim nRow As Integer
Dim nRow As Double
Dim ii As Integer
Dim arrKind '商標類別
    
m_dblKindCnt = 0
If rsTmp.RecordCount > 0 Then
    rsTmp.MoveFirst
    m_dblKindCnt = 0
    Label1(2).Caption = "公報筆數： 共 " & Format(rsTmp.RecordCount, "#,##0") & " 筆"
    Do While rsTmp.EOF = False
        grdList.Rows = grdList.Rows + 1
        nRow = grdList.Rows - 1
        For ii = 0 To grdList.Cols - 1
            If IsNull(rsTmp.Fields(ii)) = False Then
                grdList.TextMatrix(nRow, ii) = rsTmp.Fields(ii)
            End If
        Next ii
        If Me.grdList.TextMatrix(nRow, Me.grdList.Cols - 2) <> "" Then
            arrKind = Split(Me.grdList.TextMatrix(nRow, Me.grdList.Cols - 2), ",")
            m_dblKindCnt = m_dblKindCnt + Val(UBound(arrKind) + 1)
        Else
            m_dblKindCnt = m_dblKindCnt + 1
        End If
        If textTA03(2).Text <> "Y" Then Me.grdList.RowHeight(nRow) = 0
        rsTmp.MoveNext
    Loop
    grdList.FixedRows = 1 'Added by Lydia 2023/10/16
    Label1(3).Caption = "類別筆數： 共 " & Format(m_dblKindCnt, "#,##0") & " 筆"
Else
    Label1(2).Caption = "公報筆數： 筆"
    Label1(3).Caption = "類別筆數： 筆"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/19
   Set frm030614 = Nothing
End Sub

Private Sub grdList_SelChange()
'   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
'   nCurrSel = grdList.Row
'
'   ' 與前一選擇的列位置相同則不處理
'   If m_CurrSel = grdList.Row Then
'      GoTo EXITSUB
'   End If
'
'   ' 將原先選取的列回復到正常的顏色
'   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
'      grdList.Row = m_CurrSel
'      grdList.Col = 1
'      If grdList.CellBackColor <> &H80000005 Then
'         For nCol = 1 To grdList.Cols - 1
'            grdList.Col = nCol
'            If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
'            If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
'         Next nCol
'      End If
'      grdList.Col = 0
'   End If
'   ' 設定成所選取的列
'   m_CurrSel = nCurrSel
'   ' 將所選取的列反白
'   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
'      grdList.Row = m_CurrSel
'      grdList.Col = 1
'      For nCol = 1 To grdList.Cols - 1
'         grdList.Col = nCol
'         grdList.CellBackColor = &H8000000D
'         grdList.CellForeColor = &H80000005
'      Next nCol
'      grdList.Col = 0
'   End If
'EXITSUB:
End Sub

Private Sub textTA03_Change(Index As Integer)
If Index = 2 Then Me.textTA03(Index).Text = UCase(Me.textTA03(Index).Text)
End Sub

Private Sub textTA03_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    Case 2
        KeyAscii = UpperCase(KeyAscii)
        If KeyAscii <> 8 And KeyAscii <> 89 Then
            KeyAscii = 0
        End If
    End Select
End Sub

Private Sub textTA03_Validate(Index As Integer, Cancel As Boolean)
Cancel = False
If CheckKeyIn(Index) = -1 Then
   Cancel = True
End If
End Sub

Private Sub textTA03_GotFocus(Index As Integer)
   InverseTextBox textTA03(Index)
'   textTA03.IMEMode = 1
End Sub

Private Function CheckKeyIn(Index As Integer) As Integer
CheckKeyIn = -1
Select Case Index
Case 0 '公報卷期起
'   If Len(textTA03(Index).Text) <= 0 Then
'      MsgBox "請輸入欲查詢的公報起始卷期!!!", vbExclamation
'      textTA03_GotFocus Index
'      Exit Function
'   End If
Case 1 '公報卷期迄
   If Len(textTA03(Index)) <= 0 Then
'      MsgBox "請輸入欲查詢的公報終止卷期!!!", vbExclamation
'      textTA03_GotFocus Index
'      Exit Function
   ElseIf Val(textTA03(Index).Text) < Val(textTA03(0).Text) Then
      MsgBox "公報終止卷期不可小於公報起始卷期!!!", vbExclamation
      textTA03_GotFocus Index
      Exit Function
   End If
Case 2 '是否顯示明細資料
   If textTA03(Index).Text <> "Y" And textTA03(Index).Text <> "" Then
      MsgBox "是否顯示明細資料欄位必須輸入 Y 或是不輸入!!!", vbExclamation
      textTA03_GotFocus Index
      Exit Function
   End If
End Select
CheckKeyIn = 0
End Function
