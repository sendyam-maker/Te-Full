VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm170243 
   BorderStyle     =   1  '單線固定
   Caption         =   "福利金查詢及列印"
   ClientHeight    =   5520
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7308
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7308
   Begin VB.TextBox txtComp 
      Alignment       =   2  '置中對齊
      Height          =   270
      Left            =   972
      MaxLength       =   1
      TabIndex        =   2
      Top             =   888
      Width           =   324
   End
   Begin VB.TextBox txtYEAR 
      Height          =   270
      Left            =   972
      MaxLength       =   3
      TabIndex        =   0
      Top             =   192
      Width           =   735
   End
   Begin VB.TextBox txtZone 
      Height          =   285
      Left            =   972
      MaxLength       =   1
      TabIndex        =   1
      Top             =   540
      Width           =   324
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5412
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6312
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Enabled         =   0   'False
      Height          =   400
      Index           =   2
      Left            =   3936
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   120
      Width           =   800
   End
   Begin VB.Frame Frame1 
      Caption         =   "印表機"
      Height          =   570
      Left            =   2052
      TabIndex        =   8
      Top             =   4884
      Width           =   5070
      Begin VB.ComboBox cmbPrinter 
         Height          =   300
         Left            =   135
         Style           =   2  '單純下拉式
         TabIndex        =   4
         Top             =   210
         Width           =   4815
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3576
      Left            =   96
      TabIndex        =   7
      Top             =   1284
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   6287
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   5
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "轉Word列印中，請耐心等候..."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   16.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   324
      Left            =   3072
      TabIndex        =   13
      Top             =   576
      Visible         =   0   'False
      Width           =   4152
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "公司別："
      Height          =   180
      Index           =   1
      Left            =   228
      TabIndex        =   12
      Top             =   948
      Width           =   720
   End
   Begin VB.Label lblComp 
      AutoSize        =   -1  'True
      Caption         =   "台一國際智慧財產事務所"
      Height          =   180
      Left            =   1356
      TabIndex        =   11
      Top             =   948
      Width           =   1980
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "所別：           (1:北 2:中 3:南 4:高)"
      Height          =   180
      Left            =   360
      TabIndex        =   10
      Top             =   588
      Width           =   2580
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "年度：                   (ex:112)"
      Height          =   180
      Index           =   1
      Left            =   384
      TabIndex        =   9
      Top             =   228
      Width           =   2028
   End
End
Attribute VB_Name = "frm170243"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Morgan 2023/11/23
Option Explicit

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0 '查詢
         If txtYEAR = "" Then
            MsgBox "請輸入年度！", vbExclamation
            txtYEAR.SetFocus
            Exit Sub
         ElseIf Val(txtYEAR) < 100 Or Val(txtYEAR) > 200 Then
            MsgBox "年度輸入錯誤！", vbCritical
            txtYEAR.SetFocus
            Exit Sub
         End If
         
         If TxtValidate = True Then
            SetDataListWidth
            doQuery
         End If
         
      Case 2 '列印
         Screen.MousePointer = vbHourglass
         lblNote.Visible = True
         pub_OsPrinter = PUB_GetOsDefaultPrinter
         DoPrint
         PUB_SetOsDefaultPrinter pub_OsPrinter
         '若印表機變動, 則更新列印設定
         If cmbPrinter.Tag <> cmbPrinter Then
             PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, 0, 0, Me.cmbPrinter.Text
         End If
         lblNote.Visible = False
         Screen.MousePointer = vbDefault
      
      Case 1 '結束
         Unload Me
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, cmbPrinter
   lblComp = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170243 = Nothing
End Sub

Private Sub txtComp_Change()
   lblComp = ""
   If txtComp <> "" Then
      lblComp = CompNameQuery(txtComp)
   End If
End Sub

Private Sub txtComp_GotFocus()
   TextInverse txtComp
End Sub

Private Sub txtComp_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtYEAR_GotFocus()
   TextInverse txtYEAR
End Sub

Private Sub txtYEAR_KeyPress(KeyAscii As Integer)
   If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtZone_GotFocus()
   TextInverse txtZone
End Sub

Private Sub txtZone_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not (KeyAscii >= Asc("1") And KeyAscii <= Asc("4")) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Function DoPrint() As Boolean
   Dim iRows As Integer, iCols As Integer, iRow As Integer, iCol As Integer
   Dim stText As String
   Dim oTable As Word.Table
   Dim iPageNo As Integer, iRowNo As Integer
   
On Error GoTo ErrHnd
   If Pub_NewWordDoc(g_WordAp) = True Then
      PUB_SetOsDefaultPrinter cmbPrinter
      PUB_SetWordActivePrinter
            
      iRows = grdDataList.Rows - 1
      iCols = 5
      
      With g_WordAp
      .Move 0, 0
      '.WindowState = wdWindowStateMaximize
            
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4.5)
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.Font.Name = "標楷體"
      .Selection.Font.Size = 14
      
      If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
         .ActiveWindow.ActivePane.View.Type = wdPageView
      Else
         .ActiveWindow.View.Type = wdPageView
      End If
      
      '頁首
      .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
      .Selection.Font.Name = "標楷體"
      .Selection.Font.Size = 12
      Set oTable = .Selection.Tables.add(Range:=.Selection.Range, NumColumns:=3, NumRows:=2)
      With oTable
      .Columns(1).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(4.5), RulerStyle:=wdAdjustProportional
      .Columns(2).SetWidth ColumnWidth:=.Columns(2).Width + .Columns(3).Width - .Columns(1).Width, RulerStyle:=wdAdjustProportional
      .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
      .Borders(wdBorderRight).LineStyle = wdLineStyleNone
      .Borders(wdBorderTop).LineStyle = wdLineStyleNone
      .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
      .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
      .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
      .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
      .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
      .Borders.Shadow = False
      .Rows(1).Select
      End With
      .Selection.Cells.Merge
      .Selection.Font.Size = 22
      .Selection.TypeText Text:="福利金列印"
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.ParagraphFormat.SpaceAfter = 10
      
      oTable.Rows(2).Cells(1).Select
      .Selection.TypeText Text:="列印人：" & strUserName
      
      strExc(0) = "年度：" & txtYEAR
      If txtZone <> "" Then
         strExc(1) = ""
         If txtZone = "1" Then
            strExc(1) = "北所"
         ElseIf txtZone = "2" Then
            strExc(1) = "中所"
         ElseIf txtZone = "3" Then
            strExc(1) = "南所"
         ElseIf txtZone = "4" Then
            strExc(1) = "高所"
         End If
        strExc(0) = strExc(0) & vbCrLf & "所別：" & txtZone & strExc(1)
      End If
      
      oTable.Rows(2).Cells(2).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.TypeText Text:=strExc(0)
      
      
      oTable.Rows(2).Cells(3).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      strExc(0) = "列印日期：" & Format(strSrvDate(2), "###/##/##")
      .Selection.TypeText Text:=strExc(0)
      .Selection.TypeParagraph
      .Selection.TypeText Text:="頁　　次："
      .Selection.Fields.add Range:=.Selection.Range, Type:=wdFieldEmpty, Text:="Page", PreserveFormatting:=True
      
      '公司別
      If txtComp <> "" Then
         oTable.Rows.add
         oTable.Rows(oTable.Rows.Count).Select
         .Selection.Cells.Merge
         .Selection.TypeText Text:="公司別：" & txtComp & Me.lblComp
      End If
      
     
'      '表頭
'      oTable.Rows.add
'      oTable.Rows(oTable.Rows.Count).Select
'      .Selection.ParagraphFormat.SpaceBefore = 10
'      .Selection.Cells.Split NumRows:=1, NumColumns:=5, MergeBeforeSplit:=True
'      .Selection.Cells(1).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(3), RulerStyle:=wdAdjustProportional
'      .Selection.Cells(2).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(3), RulerStyle:=wdAdjustProportional
'      .Selection.Cells(3).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(3.5), RulerStyle:=wdAdjustProportional
'      .Selection.Cells(4).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(3.5), RulerStyle:=wdAdjustProportional
'      .Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle '下框線
'      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter '置中
'      .Selection.Font.Bold = True '粗體
'
'      For iCol = 1 To 5
'         oTable.Rows(oTable.Rows.Count).Cells(iCol).Select
'         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
'         .Selection.TypeText Text:=GrdDataList.TextMatrix(0, iCol - 1)
'      Next
            
      '本文
      .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
      .Selection.Orientation = wdTextOrientationHorizontal
      
      '行距
      With .Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceSingle
        .DisableLineHeightGrid = True
      End With
      
      '新增表格
      Set oTable = .Selection.Tables.add(Range:=.Selection.Range, NumColumns:=iCols, NumRows:=1)
      With oTable
      .Columns(1).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(3), RulerStyle:=wdAdjustProportional
      .Columns(2).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(3), RulerStyle:=wdAdjustProportional
      .Columns(3).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(3.5), RulerStyle:=wdAdjustProportional
      .Columns(4).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(3.5), RulerStyle:=wdAdjustProportional
      
      .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
      .Borders(wdBorderRight).LineStyle = wdLineStyleNone
      .Borders(wdBorderTop).LineStyle = wdLineStyleNone
      .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
      .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
      .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
      .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
      .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
      .Borders.Shadow = False
      End With
      
      oTable.Columns(1).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      oTable.Columns(2).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      oTable.Columns(3).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      oTable.Columns(4).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      oTable.Columns(5).Select
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      iPageNo = 0
      iRowNo = 0
      For iRow = 1 To iRows
         iRowNo = iRowNo + 1
         
         If iRow = 1 Then
            oTable.Rows(1).Select
         ElseIf iRowNo > 5 Then
            oTable.Rows.add
            oTable.Rows(oTable.Rows.Count).Select
         Else
            .Selection.MoveDown Unit:=wdLine, Count:=1
         End If
         
         '跳頁印表頭
         If .Selection.Information(wdActiveEndPageNumber) > iPageNo Then
            iPageNo = .Selection.Information(wdActiveEndPageNumber)
            .Selection.InsertRows 5
            .Selection.Collapse Direction:=wdCollapseStart
            .Selection.SelectRow
            .Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle '下框線
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter '置中
            .Selection.Font.Bold = True '粗體
            For iCol = 1 To 5
               .Selection.SelectRow
               .Selection.Cells(iCol).Select
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
               .Selection.TypeText Text:=grdDataList.TextMatrix(0, iCol - 1)
            Next
            .Selection.MoveDown Unit:=wdLine, Count:=1
            iRowNo = 1
         End If
         
         If iRow = iRows Then
            .Selection.SelectRow
            .Selection.Borders(wdBorderTop).LineStyle = wdLineStyleSingle '上框線
            .Selection.Font.Bold = True '粗體
         End If
         For iCol = 1 To 5
            .Selection.SelectRow
            .Selection.Cells(iCol).Select
            .Selection.TypeText Text:=grdDataList.TextMatrix(iRow, iCol - 1)
         Next
      Next
      
      .PrintOut Background:=False, Copies:=1, Collate:=True
      .ActiveDocument.Close wdDoNotSaveChanges
      .Quit wdDoNotSaveChanges
      End With
      Set g_WordAp = Nothing
      DoPrint = True
   End If
   Exit Function
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   
   If txtYEAR = "" Then
      MsgBox "請輸入年度!"
      Exit Function
   End If
   TxtValidate = True
End Function

Private Sub SetDataListWidth(Optional ByVal p_bolHeaderOnly As Boolean = False)
Dim ii As Integer
   With grdDataList
   .Visible = False
   If p_bolHeaderOnly = False Then
      .Clear
      .Rows = 2: .Cols = 5: .FixedRows = 1: .FixedCols = 0
   End If
   
   .row = 0
   .col = 0: .ColWidth(.col) = 1000: .Text = "員工編號"
   .ColAlignment(.col) = flexAlignCenterCenter
   .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
   .col = 1: .ColWidth(.col) = 1200: .Text = "姓名"
   .ColAlignment(.col) = flexAlignCenterCenter
   .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
   .col = 2: .ColWidth(.col) = 1400: .Text = "尾牙摸彩"
   .ColAlignment(.col) = flexAlignRightCenter
   .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
   .col = 3: .ColWidth(.col) = 1400: .Text = "年資"
   .ColAlignment(.col) = flexAlignRightCenter
   .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
   .col = 4: .ColWidth(.col) = 1400: .Text = "全勤"
   .ColAlignment(.col) = flexAlignRightCenter
   .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
   .Visible = True
   End With
End Sub
   
Private Sub doQuery()
   Dim stCon As String
   '紀錄最後條件
   txtYEAR.Tag = txtYEAR
   txtZone.Tag = txtZone
   txtComp.Tag = txtComp
   lblComp.Tag = lblComp
   
   stCon = ""
   If txtYEAR <> "" Then stCon = stCon & " and mb01=" & (Val(txtYEAR) + 1911)
   'Modified by Morgan 2025/1/9 尾牙摸彩以發放所統計
   'If txtZone <> "" Then stCon = stCon & " and st06='" & txtZone & "'"
   If txtZone <> "" Then stCon = stCon & " and nvl(mb10,st06)='" & txtZone & "'"
   If txtComp <> "" Then stCon = stCon & " and sd19='" & txtComp & "'"
   
   strExc(0) = "select mb04,st02,trim(to_char(nvl(sum(decode(mb02,'01',mb05)),0),'999,990')) S1" & _
      ",trim(to_char(nvl(sum(decode(mb02,'02',mb05)),0),'999,990')) S2" & _
      ",trim(to_char(nvl(sum(decode(mb02,'03',mb05)),0),'999,990')) S3" & _
      " From miscbonus, staff,salarydata where st01(+)=mb04 and sd01(+)=st01" & stCon & _
      " group by mb04,st02 order by 1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   Set grdDataList.Recordset = RsTemp.Clone
   SetDataListWidth True
   If intI = 1 Then
      AddSubTotal
      cmdok(2).Enabled = True
   Else
      cmdok(2).Enabled = False
   End If
   
End Sub

Private Sub AddSubTotal()
   Dim ii As Integer, jj As Integer, stComp As String, stGroup As String, stGroupName As String, strAddItem As String, lngColor As Long, lngColor1 As Long
   Dim dblSub1(9) As Double, dblTot(9) As Double
   
   lngColor = &H90EE90
   lngColor1 = &H90EE40
      
   With grdDataList
      .Visible = False
      ii = 1
      Do While ii < .Rows
         For jj = 2 To 4
            dblSub1(jj) = dblSub1(jj) + Val(Format(.TextMatrix(ii, jj)))
         Next
         ii = ii + 1
      Loop
      strAddItem = stGroup & vbTab & "合計:"
      For jj = 2 To 4
         strAddItem = strAddItem & vbTab & Format(dblSub1(jj), "#,##0")
         dblSub1(jj) = 0
      Next
      .AddItem strAddItem, ii
      
      .row = ii
      For jj = 0 To .Cols - 1
         .col = jj: .CellBackColor = lngColor
         .CellFontBold = True
      Next
      .Visible = True
   End With
End Sub

