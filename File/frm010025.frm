VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm010025 
   BorderStyle     =   1  '單線固定
   Caption         =   "取消發文資料查詢"
   ClientHeight    =   5745
   ClientLeft      =   3780
   ClientTop       =   3690
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8955
   Begin VB.TextBox txtCP01 
      Height          =   264
      Left            =   1830
      MaxLength       =   3
      TabIndex        =   2
      Top             =   360
      Width           =   732
   End
   Begin VB.TextBox txtCP04 
      Height          =   264
      Left            =   4290
      MaxLength       =   2
      TabIndex        =   5
      Top             =   360
      Width           =   492
   End
   Begin VB.TextBox txtCP03 
      Height          =   264
      Left            =   3870
      MaxLength       =   1
      TabIndex        =   4
      Top             =   360
      Width           =   372
   End
   Begin VB.TextBox txtCP02 
      Height          =   264
      Left            =   2610
      MaxLength       =   6
      TabIndex        =   3
      Top             =   360
      Width           =   1212
   End
   Begin VB.TextBox txtCP27 
      Height          =   264
      Index           =   1
      Left            =   3270
      MaxLength       =   7
      TabIndex        =   1
      Top             =   60
      Width           =   1185
   End
   Begin VB.TextBox txtCP27 
      Height          =   264
      Index           =   0
      Left            =   1830
      MaxLength       =   7
      TabIndex        =   0
      Top             =   60
      Width           =   1185
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   4890
      Left            =   30
      TabIndex        =   9
      Top             =   840
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   8625
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
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
      _Band(0).Cols   =   5
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "尋找(&O)"
      Default         =   -1  'True
      Height          =   345
      Index           =   1
      Left            =   6690
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   60
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   0
      Left            =   7650
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   60
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "~"
      Height          =   255
      Index           =   1
      Left            =   3060
      TabIndex        =   12
      Top             =   60
      Width           =   165
   End
   Begin VB.Label Label1 
      Caption         =   "共　0　件"
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   5
      Left            =   7290
      TabIndex        =   8
      Top             =   570
      Width           =   1605
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "取消發文日期："
      Height          =   255
      Index           =   0
      Left            =   450
      TabIndex        =   11
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "本所案號："
      Height          =   255
      Index           =   2
      Left            =   450
      TabIndex        =   10
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frm010025"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/01 Form2.0已修改 GrdDataList
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
Option Explicit

Dim i As Integer, j As Integer
Dim lngCounterI As Long
Dim m_bolPrintRight As Boolean
Dim m_DBTime As Long '系統時間
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_intRow As Integer, m_intCol As Integer
Dim bolV As Boolean
Public cmdState As Integer


Private Sub Timer1_Timer()
   m_DBTime = Format(Now, "HHMMSS")
End Sub

Private Sub SetDataListWidth()
grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "取消發文日"
grdDataList.ColWidth(0) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "本所案號"
grdDataList.ColWidth(1) = 1500
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 2: grdDataList.Text = "案件性質"
grdDataList.ColWidth(2) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 3: grdDataList.Text = "總收文號"
grdDataList.ColWidth(3) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 4: grdDataList.Text = "發文室取消備註"
grdDataList.ColWidth(4) = 3500
grdDataList.CellAlignment = flexAlignCenterCenter
End Sub

Private Sub cmdOK_Click(Index As Integer)
cmdState = Index
PubShowNextData
Exit Sub
End Sub

Public Sub PubShowNextData()
Select Case cmdState
   Case 0 '結束
      'fnCloseAllFrm100
      Unload Me
      Set frm010025 = Nothing
   Case 1 '尋找
      Call SearchData
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   '更新系統時間
   m_DBTime = ServerTime
   time = Format(m_DBTime, "##:##:##")
   
   SetDataListWidth
   
   '發文日期
   txtCP27(0).Text = strSrvDate(2)
   txtCP27(1).Text = strSrvDate(2)
   
   m_bolPrintRight = IsUserHasRightOfFunction("frm010025", strPrint, False)
   
   cmdState = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm010025 = Nothing
End Sub

'Private Sub GrdDataList_Click()
'Dim strItem As String
'
'   GrdDataList.Visible = False
'
'   '依點選的欄位做排序
'   If GrdDataList.MouseRow = 0 Then
'      If GrdDataList.MouseCol <> 0 Then
'         m_intRow = GrdDataList.MouseRow
'         m_intCol = GrdDataList.MouseCol
'         GrdDataList.row = m_intRow
'         GrdDataList.col = m_intCol
'         Select Case m_intCol
'            Case 5, 13
'               '數字
'               If m_blnColOrderAsc = True Then
'                   Me.GrdDataList.Sort = 3 '昇冪
'                   m_blnColOrderAsc = False
'               Else
'                   Me.GrdDataList.Sort = 4 '降冪
'                   m_blnColOrderAsc = True
'               End If
'           Case Else
'               '字串
'               If m_blnColOrderAsc = True Then
'                   Me.GrdDataList.Sort = 5 '昇冪
'                   m_blnColOrderAsc = False
'               Else
'                   Me.GrdDataList.Sort = 6 '降冪
'                   m_blnColOrderAsc = True
'               End If
'         End Select
'      End If
'   End If
'
'   '勾選
'   GrdDataList.row = GrdDataList.MouseRow
'   GrdDataList.col = 0
'   If GrdDataList.row <> 0 Then
'      'If grdDataList.TextMatrix(GrdDataList.MouseRow, 13) = "" Then
'         If GrdDataList.Text = "V" Then
'              GrdDataList.Text = ""
'              For i = 0 To GrdDataList.Cols - 1
'                  GrdDataList.col = i
'                  GrdDataList.CellBackColor = QBColor(15)
'                  If i <= 2 Then
'                     GrdDataList.CellBackColor = &H8000000F
'                  End If
'             Next i
'         Else
'              GrdDataList.Text = "V"
'              For i = 0 To GrdDataList.Cols - 1
'                  GrdDataList.col = i
'                  GrdDataList.CellBackColor = &HFFC0C0
'                  If i <= 2 Then
'                     GrdDataList.CellBackColor = &H8000000F
'                  End If
'              Next i
'         End If
'      'End If
'   End If
'
'   GrdDataList.Visible = True
'End Sub

'Private Sub GrdDataList_Sort()
'   GrdDataList.Visible = False
'   '依點選的欄位做排序
'   If m_intRow = 0 Then
'      If m_intCol <> 0 Then
'         GrdDataList.row = m_intRow
'         GrdDataList.col = m_intCol
'         Select Case m_intCol
'            Case 5, 13
'               '數字
'               If m_blnColOrderAsc = False Then
'                   Me.GrdDataList.Sort = 3 '昇冪
'               Else
'                   Me.GrdDataList.Sort = 4 '降冪
'               End If
'           Case Else
'               '字串
'               If m_blnColOrderAsc = False Then
'                   Me.GrdDataList.Sort = 5 '昇冪
'               Else
'                   Me.GrdDataList.Sort = 6 '降冪
'               End If
'         End Select
'      End If
'   End If
'   GrdDataList.Visible = True
'End Sub

Private Sub SearchData()
Dim strComWhere As String, strPSel As String, strTSel As String, strSSel As String
Dim dblAL05 As Double
Dim strSql As String, strWhere As String

'重新檢查欄位有效性
If TxtValidate = False Then Exit Sub
strComWhere = ""

'發文日期
If Len(Trim(txtCP27(0).Text)) <> 0 Then
   strComWhere = strComWhere & " and cp132>=" & ChangeTStringToWString(txtCP27(0))
End If
If Len(Trim(txtCP27(1).Text)) <> 0 Then
   strComWhere = strComWhere & " and cp132<=" & ChangeTStringToWString(txtCP27(1))
End If
'本所案號
If Len(Trim(txtcp01.Text)) <> 0 Then
   strComWhere = strComWhere & " and cp01='" & Trim(txtcp01.Text) & "' "
End If
If Len(Trim(txtcp02.Text)) <> 0 Then
   strComWhere = strComWhere & " and cp02='" & Trim(txtcp02.Text) & "' "
End If
If Len(Trim(txtcp03.Text)) <> 0 Then
   strComWhere = strComWhere & " and cp03='" & Trim(txtcp03.Text) & "' "
End If
If Len(Trim(txtcp04.Text)) <> 0 Then
   strComWhere = strComWhere & " and cp04='" & Trim(txtcp04.Text) & "' "
End If

strSql = "select sqldateT(CP132),CP01||'-'||CP02||'-'||CP03||'-'||CP04,CPM03,CP09,CP131 " & _
                " from caseprogress,casepropertymap " & _
                " where cp01=cpm01(+) " & _
                    " and cp10=cpm02(+) " & _
                    " and cp132 >0 and not cp132 is null " & strComWhere & _
                " order by cp132,CP01,CP02 "

Screen.MousePointer = vbHourglass
grdDataList.Clear
grdDataList.Rows = 2
SetDataListWidth
'GrdDataList.FixedCols = 0

CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 Then
    Label1(5).Caption = "共　" & adoRecordset.RecordCount & "　件"
    Set grdDataList.Recordset = adoRecordset
Else
    Label1(5).Caption = "共　0　件"
    ShowNoData
    grdDataList.Clear
End If
SetDataListWidth
'GrdDataList.FixedCols = 3
CheckOC
'Call GrdDataList_Sort

'With Me.GrdDataList
'   For i = 1 To .Rows - 1
'      .row = i
'      .col = 9
'      If .Text <> "" Then
'         For j = 0 To .Cols - 1
'            .row = i
'            .col = j
'            .CellBackColor = &HC0FFC0   '&HFF&
'         Next j
'      End If
'   Next i
'End With

''若只有一筆資料, 則直接設定為點選此筆資料
'With Me.GrdDataList
'   If .Rows = 2 Then
'      .row = 1
'      .col = 1
'      If .Text <> "" Then
'        .Visible = False
'        .row = 1
'        .col = 0
'        .Text = "V"
'        For i = 0 To .Cols - 1
'            .col = i
'            .CellBackColor = &HFFC0C0
'            If i <= 2 Then
'              GrdDataList.CellBackColor = &H8000000F
'            End If
'        Next i
'        .Visible = True
'      End If
'   End If
'End With

Screen.MousePointer = vbDefault
End Sub

Private Sub txtcp01_GotFocus()
InverseTextBox txtcp01
End Sub

Private Sub txtcp01_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtcp01_Validate(Cancel As Boolean)
   If Trim(txtcp01.Text) > "" Then
      strExc(0) = "SELECT * FROM SystemKind WHERE SK01='" & Trim(txtcp01.Text) & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
      Else
         MsgBox "系統別錯誤！", vbInformation, "輸入條件錯誤"
         Call txtcp01_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub txtcp02_GotFocus()
InverseTextBox txtcp02
End Sub

Private Sub txtCP02_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtcp02_Validate(Cancel As Boolean)
   If Trim(txtcp02.Text) > "" Then
      txtcp02.Text = Right("000000" & Trim(txtcp02.Text), 6)
   End If
End Sub

Private Sub txtCP27_GotFocus(Index As Integer)
InverseTextBox txtCP27(Index)
End Sub

Private Sub txtCP27_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1
         KeyAscii = Pub_NumAscii(KeyAscii)
         'KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txtCP27_Validate(Index As Integer, Cancel As Boolean)
'   If CheckIsTaiwanDate(txtCP27(index), False) = False Then
'        Cancel = True
'        MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
'        Call txtCP27_GotFocus(index)
'        Exit Sub
'    End If
End Sub

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
Dim s As Integer

TxtValidate = False

If Len(Trim(txtCP27(0).Text)) = 0 And _
   Len(Trim(txtCP27(1).Text)) = 0 And _
   Len(Trim(txtcp01.Text)) = 0 And _
   Len(Trim(txtcp02.Text)) = 0 Then
   s = MsgBox("至少輸入一項查詢條件！", , "輸入條件錯誤")
   txtCP27(0).SetFocus
   Exit Function
End If

'發文日期
If Len(Trim(txtCP27(0).Text)) <> 0 Or _
   Len(Trim(txtCP27(1).Text)) <> 0 Then
   If Len(Trim(txtCP27(0).Text)) = 0 Then
      s = MsgBox("取消發文起始日期不可空白！", , "輸入條件錯誤")
      txtCP27(0).SetFocus
      Exit Function
   End If
   If Len(Trim(txtCP27(1).Text)) = 0 Then
      s = MsgBox("取消發文迄止日期不可空白！", , "輸入條件錯誤")
      txtCP27(1).SetFocus
      Exit Function
   End If
End If

'本所案號
If Len(Trim(txtcp01.Text)) <> 0 Or _
   Len(Trim(txtcp02.Text)) <> 0 Then
   If Len(Trim(txtcp01.Text)) = 0 Then
      s = MsgBox("請輸入完整本所案號！", , "輸入條件錯誤")
      txtcp01.SetFocus
      Exit Function
   End If
   If Len(Trim(txtcp02.Text)) = 0 Then
      s = MsgBox("請輸入完整本所案號！", , "輸入條件錯誤")
      txtcp02.SetFocus
      Exit Function
   End If
End If

If Me.txtCP27(0).Enabled = True Then
   Cancel = False
   txtCP27_Validate 0, Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.txtCP27(1).Enabled = True Then
   Cancel = False
   txtCP27_Validate 1, Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.txtcp01.Enabled = True Then
   Cancel = False
   txtcp01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function
