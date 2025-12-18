VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090613_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件處理時間統計查詢(案件明細)"
   ClientHeight    =   5232
   ClientLeft      =   168
   ClientTop       =   960
   ClientWidth     =   9300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5232
   ScaleWidth      =   9300
   Begin VB.CommandButton cmdOK 
      Caption         =   "會稿超時/短時案件(&Word)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   48
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   2304
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7920
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   60
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4500
      Left            =   0
      TabIndex        =   1
      Top             =   510
      Width           =   9285
      _ExtentX        =   16383
      _ExtentY        =   7938
      _Version        =   393216
      Cols            =   19
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
      _Band(0).Cols   =   19
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "游標指到會變黑色的欄位表示可點選顯示明細資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   450
      TabIndex        =   4
      Top             =   5040
      Width           =   4440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PS:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "符號說明：＊閉卷●銷卷"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   10
      Left            =   3480
      TabIndex        =   2
      Top             =   270
      Width           =   1875
   End
End
Attribute VB_Name = "frm090613_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/08 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、Label1(10)
'Create By Sindy 2015/9/30
Option Explicit

Dim m_iRow As Integer
Public m_bolNoDetail As Boolean
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序


Public Sub SetDataListWidth()
Dim iCol As Integer

m_blnColOrderAsc = True
grdDataList.Cols = 22

iCol = 0
grdDataList.row = 0
grdDataList.col = iCol: grdDataList.Text = "V"
grdDataList.ColWidth(iCol) = 0 '200
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "收文日"
grdDataList.ColWidth(iCol) = 810
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "本所案號"
grdDataList.ColWidth(iCol) = 1550
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
Dim iDep As String
iDep = PUB_GetST06(strUserNum)
grdDataList.col = iCol: grdDataList.Text = "分所號"
'電腦中心，跟分所才秀
If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
    grdDataList.ColWidth(iCol) = 0
Else
    grdDataList.ColWidth(iCol) = 620
End If
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "案件名稱"
grdDataList.ColWidth(iCol) = 1300
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "申請國家"
grdDataList.ColWidth(iCol) = 800
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "種類"
grdDataList.ColWidth(iCol) = 800
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "案件性質"
grdDataList.ColWidth(iCol) = 800
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "本所期限"
grdDataList.ColWidth(iCol) = 810
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "承辦期限"
grdDataList.ColWidth(iCol) = 810
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "齊備日"
grdDataList.ColWidth(iCol) = 810
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "完稿日"
grdDataList.ColWidth(iCol) = 810
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "預會日"
grdDataList.ColWidth(iCol) = 810
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "會稿日"
grdDataList.ColWidth(iCol) = 810
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "核稿人"
grdDataList.ColWidth(iCol) = 700
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "會稿完成日"
grdDataList.ColWidth(iCol) = 810
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "發文日"
grdDataList.ColWidth(iCol) = 810
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "承辦天數"
grdDataList.ColWidth(iCol) = 810
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "承辦備註"
grdDataList.ColWidth(iCol) = 810
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "承辦人"
grdDataList.ColWidth(iCol) = 700
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = "智權人員"
grdDataList.ColWidth(iCol) = 700
grdDataList.CellAlignment = flexAlignCenterCenter

iCol = iCol + 1
grdDataList.col = iCol: grdDataList.Text = ""
grdDataList.ColWidth(iCol) = 0
grdDataList.CellAlignment = flexAlignCenterCenter
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         Unload Me
         
      'Added by Morgan 2024/3/15
      Case 1 '會稿超時/短時案件(&Word)
         Screen.MousePointer = vbHourglass
         Me.Enabled = False
         ExportWord
         Me.Enabled = True
         Screen.MousePointer = vbDefault
   End Select
End Sub

Private Sub Command1_Click()
Dim strCaseNo As String
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
Dim ii As Integer
   
On Error GoTo ErrHnd
   
   For ii = 1 To grdDataList.Rows - 1
      grdDataList.row = ii
      grdDataList.col = 1
      If grdDataList.CellBackColor = grdDataList.BackColor And _
         Trim(grdDataList.TextMatrix(ii, 2)) <> "" Then
         strCaseNo = Pub_RplStr(Trim(grdDataList.TextMatrix(ii, 2)))
         strCP01 = SystemNumber(strCaseNo, 1)
         strCP02 = SystemNumber(strCaseNo, 2)
         strCP03 = SystemNumber(strCaseNo, 3)
         strCP04 = SystemNumber(strCaseNo, 4)
         frm100101_1.txtSystem = strCP01
         frm100101_1.txtCode(0) = strCP02
         frm100101_1.txtCode(1) = strCP03
         frm100101_1.txtCode(2) = strCP04
         frm100101_1.cmdSearch_Click
         frm100101_1.Show
      End If
   Next ii
   
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frm090613_1.Show
   Set frm090613_3 = Nothing
End Sub

Private Sub GrdDataList_Click()
Dim strCaseNo As String
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
   
   grdDataList.row = grdDataList.MouseRow
   grdDataList.col = grdDataList.MouseCol
   If grdDataList.col > 0 And grdDataList.row > 0 Then
      If Trim(grdDataList.TextMatrix(grdDataList.row, 2)) <> "" Then
         strCaseNo = Pub_RplStr(Trim(grdDataList.TextMatrix(grdDataList.row, 2)))
         strCP01 = SystemNumber(strCaseNo, 1)
         strCP02 = SystemNumber(strCaseNo, 2)
         strCP03 = SystemNumber(strCaseNo, 3)
         strCP04 = SystemNumber(strCaseNo, 4)
         frm100101_1.txtSystem = strCP01
         frm100101_1.txtCode(0) = strCP02
         frm100101_1.txtCode(1) = strCP03
         frm100101_1.txtCode(2) = strCP04
         frm100101_1.cmdSearch_Click
         frm100101_1.Show
      End If
   End If
End Sub

Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow grdDataList, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   grdDataList.col = nCol
   grdDataList.row = nRow
   If Me.grdDataList.row < 1 And Me.grdDataList.Text <> "V" Then
      If Me.grdDataList.Text = "承辦天數" Then
         If m_blnColOrderAsc = True Then
            Me.grdDataList.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.grdDataList.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.grdDataList.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.grdDataList.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

Private Sub grdDataList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If m_bolNoDetail = True Then Exit Sub
   
   Dim iRow As Integer, iCol As Integer, lBackColor As Long
   Dim ii As Integer
   
   With grdDataList
      iRow = .MouseRow
      iCol = .MouseCol
      
      If iRow = m_iRow Then
         Exit Sub
      End If
      If m_iRow <> 0 Then
         .row = m_iRow
         For iCol = 0 To .Cols - 1
            .col = iCol
            .CellForeColor = .ForeColor
            .CellBackColor = .BackColor
         Next
         m_iRow = 0
      End If
      
      If iRow > 0 Then
         .row = iRow
         For iCol = 0 To .Cols - 1
            .col = iCol
            .CellForeColor = .BackColor
            .CellBackColor = .ForeColor
         Next
         m_iRow = .row
      End If
   End With
End Sub
'Added by Morgan 2024/3/15
Private Sub ExportWord()
   Dim iResumeCnt As Integer
   Dim stTmp As String
   
On Error GoTo ErrHnd
   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
   g_WordAp.Documents.add
   With g_WordAp.Application
      .WindowState = wdWindowStateMaximize
      .Visible = True
      
      '邊框設單線
      With .Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
      End With
      '橫印
      .Selection.PageSetup.Orientation = wdOrientLandscape
      '邊界
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(1)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(1)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(1)
      
      .Selection.Font.Name = "標楷體"
      .Selection.Font.Size = 12
      If Left(frm090613.Txt1(26), 3) = Left(frm090613.Txt1(27), 3) Then
         stTmp = frm090613.Txt1(26) & "~" & Mid(frm090613.Txt1(27), 4)
      Else
         stTmp = frm090613.Txt1(26) & "~" & frm090613.Txt1(27)
      End If
      
      stTmp = "承辦天數統計：" & stTmp & String(2, vbTab) & "會稿-超時" & String(9, vbTab) & "P案超過11天(含11天)  CFP案超過22天(含22天)"
      .Selection.TypeText Text:=stTmp
      .Selection.TypeParagraph
      AddNewTable
      .Selection.EndKey Unit:=wdStory
      stTmp = String(8, vbTab) & "會稿-短時" & String(9, vbTab) & "P案少於5天(含5天)  CFP案少於10天(含10天)"
      .Selection.TypeText Text:=stTmp
      .Selection.TypeParagraph
      AddNewTable 1
      .Activate
   End With
   
ErrHnd:
   If Err.Number <> 0 Then
      If iResumeCnt > 3 Then
         MsgBox "錯誤 : " & Err.Description, vbCritical
      Else
         iResumeCnt = iResumeCnt + 1
         Select Case Err.Number
            Case 91:
               g_WordAp.Documents.add
               Resume Next
            Case 462:
               Set g_WordAp = New Word.Application
               Resume
            Case Else:
               MsgBox "錯誤" & Err.Number & " : " & Err.Description, vbCritical
         End Select
      End If
   End If
End Sub
'Added by Morgan 2024/3/15
Private Sub AddNewTable(Optional iType As Integer = 0)
   Dim oTable As Word.Table
   Dim iCol As Integer, iRow As Integer, ii As Integer
   Dim bAdd As Boolean
   
   With g_WordAp.Application
      '新增表格
      Set oTable = .Selection.Tables.add(Range:=.Selection.Range, NumRows:=1, NumColumns:=11)
      'oTable.AllowAutoFit = True
      .Selection.SelectRow
      With .Selection.Borders(wdBorderTop)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
      End With
      With .Selection.Borders(wdBorderLeft)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
      End With
      With .Selection.Borders(wdBorderBottom)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
      End With
      With .Selection.Borders(wdBorderRight)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
      End With
      With .Selection.Borders(wdBorderHorizontal)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
      End With
      With .Selection.Borders(wdBorderVertical)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
      End With
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
      
      '超時
      ii = 1
      oTable.Columns(1).SetWidth ColumnWidth:=.CentimetersToPoints(1.7), RulerStyle:=wdAdjustProportional
      oTable.Rows(ii).Cells(1).Select
      .Selection.TypeText "工程師"
      oTable.Columns(2).SetWidth ColumnWidth:=.CentimetersToPoints(1.3), RulerStyle:=wdAdjustProportional
      oTable.Rows(ii).Cells(2).Select
      .Selection.TypeText "承辦天數"
      oTable.Columns(3).SetWidth ColumnWidth:=.CentimetersToPoints(3), RulerStyle:=wdAdjustProportional
      oTable.Rows(ii).Cells(3).Select
      .Selection.TypeText "案號"
      oTable.Columns(4).SetWidth ColumnWidth:=.CentimetersToPoints(4.5), RulerStyle:=wdAdjustProportional
      oTable.Rows(ii).Cells(4).Select
      .Selection.TypeText "專利名稱"
      oTable.Columns(5).SetWidth ColumnWidth:=.CentimetersToPoints(1.3), RulerStyle:=wdAdjustProportional
      oTable.Rows(ii).Cells(5).Select
      .Selection.TypeText "其他時間"
      oTable.Columns(6).SetWidth ColumnWidth:=.CentimetersToPoints(1.3), RulerStyle:=wdAdjustProportional
      oTable.Rows(ii).Cells(6).Select
      .Selection.TypeText "撰稿天數"
      oTable.Columns(7).SetWidth ColumnWidth:=.CentimetersToPoints(1.3), RulerStyle:=wdAdjustProportional
      oTable.Rows(ii).Cells(7).Select
      .Selection.TypeText "核稿天數"
      oTable.Columns(8).SetWidth ColumnWidth:=.CentimetersToPoints(1.3), RulerStyle:=wdAdjustProportional
      oTable.Rows(ii).Cells(8).Select
      .Selection.TypeText "原因"
      oTable.Columns(9).SetWidth ColumnWidth:=.CentimetersToPoints(7), RulerStyle:=wdAdjustProportional
      oTable.Rows(ii).Cells(9).Select
      .Selection.TypeText "備註"
      oTable.Columns(10).SetWidth ColumnWidth:=.CentimetersToPoints(1.7), RulerStyle:=wdAdjustProportional
      oTable.Rows(ii).Cells(10).Select
      .Selection.TypeText "業務"
      'oTable.Columns(11).SetWidth ColumnWidth:=.CentimetersToPoints(4.5), RulerStyle:=wdAdjustProportional
      oTable.Rows(ii).Cells(11).Select
      .Selection.TypeText "客戶"
            
      For iRow = 1 To grdDataList.Rows - 1
         bAdd = False
         If iType = 0 Then
            If (Left(grdDataList.TextMatrix(iRow, 2), 2) = "P-" And Val(grdDataList.TextMatrix(iRow, 17)) >= 11) Or (Left(grdDataList.TextMatrix(iRow, 2), 4) = "CFP-" And Val(grdDataList.TextMatrix(iRow, 17)) >= 22) Then
               bAdd = True
            End If
         Else
            If (Left(grdDataList.TextMatrix(iRow, 2), 2) = "P-" And Val(grdDataList.TextMatrix(iRow, 17)) <= 5) Or (Left(grdDataList.TextMatrix(iRow, 2), 4) = "CFP-" And Val(grdDataList.TextMatrix(iRow, 17)) <= 10) Then
               bAdd = True
            End If
         End If
         If bAdd Then
            ii = ii + 1
            oTable.Rows.add
            oTable.Rows(ii).Cells(1).Select
            .Selection.TypeText grdDataList.TextMatrix(iRow, 19) '"工程師"
            oTable.Rows(ii).Cells(2).Select
            .Selection.TypeText grdDataList.TextMatrix(iRow, 17) '"承辦天數"
            oTable.Rows(ii).Cells(3).Select
            .Selection.TypeText Replace(grdDataList.TextMatrix(iRow, 2), "-0-00", "") '"案號"
            oTable.Rows(ii).Cells(4).Select
            .Selection.TypeText grdDataList.TextMatrix(iRow, 4) '"專利名稱"
            oTable.Rows(ii).Cells(6).Select
            .Selection.TypeText GetWorkDay(DBDATE(grdDataList.TextMatrix(iRow, 11)), DBDATE(grdDataList.TextMatrix(iRow, 10))) '"撰稿天數" =齊備-完稿(工作天數)
            oTable.Rows(ii).Cells(7).Select
            .Selection.TypeText GetWorkDay(DBDATE(grdDataList.TextMatrix(iRow, 13)), DBDATE(grdDataList.TextMatrix(iRow, 11))) '"核稿天數" =完稿-會稿(工作天數)
            oTable.Rows(ii).Cells(10).Select
            .Selection.TypeText grdDataList.TextMatrix(iRow, 20) '"業務"
            oTable.Rows(ii).Cells(11).Select
            .Selection.TypeText PUB_GetCustName(Replace(Pub_RplStr(Trim(grdDataList.TextMatrix(iRow, 2))), "-", "")) '"客戶"
         End If
      Next
      
   End With
End Sub
