VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm210117 
   BorderStyle     =   1  '單線固定
   Caption         =   "各所業務工作報告統計"
   ClientHeight    =   5364
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9492
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5364
   ScaleWidth      =   9492
   Begin VB.TextBox txtZone 
      Height          =   285
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   0
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8475
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6570
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdWord 
      Caption         =   "Word(&W)"
      Height          =   400
      Left            =   7470
      TabIndex        =   3
      Top             =   120
      Width           =   930
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   285
      Index           =   0
      Left            =   1170
      MaxLength       =   5
      TabIndex        =   1
      Top             =   390
      Width           =   915
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   285
      Index           =   1
      Left            =   2430
      MaxLength       =   5
      TabIndex        =   2
      Top             =   390
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4635
      Left            =   90
      TabIndex        =   6
      Top             =   720
      Width           =   9270
      _ExtentX        =   16341
      _ExtentY        =   8170
      _Version        =   393216
      BackColor       =   -2147483624
      Rows            =   20
      Cols            =   8
      FixedCols       =   3
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      MergeCells      =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label1 
      Caption         =   "所別"
      Height          =   180
      Left            =   135
      TabIndex        =   10
      Top             =   105
      Width           =   900
   End
   Begin VB.Label lblZone 
      Caption         =   "（1北所 2中所 3南所 4高所）"
      Height          =   180
      Left            =   2205
      TabIndex        =   9
      Top             =   105
      Width           =   4590
   End
   Begin VB.Label Label3 
      Caption         =   "(年月)"
      Height          =   180
      Left            =   3375
      TabIndex        =   8
      Top             =   450
      Width           =   585
   End
   Begin VB.Label Label2 
      Caption         =   "點數結算期間"
      Height          =   180
      Left            =   135
      TabIndex        =   5
      Top             =   435
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   2160
      X2              =   2395
      Y1              =   525
      Y2              =   525
   End
End
Attribute VB_Name = "frm210117"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/14 改成Form2.0 ; grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'Create by Morgan 2007/9/26
Option Explicit
'Add by Amy 2015/04/22
Dim strRowN()
Dim dblTot(1 To 19) As Double

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   Screen.MousePointer = vbHourglass
   If ConstrainCheck = True Then
      Call SetDataListWidth
      Call doQuery
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Function ConstrainCheck() As Boolean
   Dim bolCancel As Boolean
   Dim strA0b01_1 As String, strA0b01_J As String, strA0b05 As String 'Add by Amy 2016/03/01
   
   If txtZone = "" Then
      MsgBox "請輸入所別！", vbExclamation
      txtZone.SetFocus
      Exit Function
   End If
   
   If txtCloseDate(0) = "" Then
      MsgBox "請輸入點數結算期間起月！", vbExclamation
      txtCloseDate(0).SetFocus
      Exit Function
   Else
      bolCancel = False
      Call txtCloseDate_Validate(0, bolCancel)
      If bolCancel = True Then
         txtCloseDate(0).SetFocus
         txtCloseDate_GotFocus 0
         Exit Function
      End If
   End If
   
   strA0b01_1 = Left(GetA0b01(strA0b05, "1"), 5)
   strA0b01_J = Left(GetA0b01(strA0b05, "J"), 5)
   If txtCloseDate(1) = "" Then
      MsgBox "請輸入點數結算期間迄月！", vbExclamation
      txtCloseDate(1).SetFocus
      Exit Function
   'Add by Amy 2016/03/01 原寫於Validate 因輸完年月按查詢else被觸發不focus會設於txtCloseDate(1)無法離開
   ElseIf Val(txtCloseDate(1)) > Val(strA0b01_1) Or Val(txtCloseDate(1)) > Val(strA0b01_J) Then
        If Val(txtCloseDate(1)) > Val(strA0b01_J) Then strA0b01_1 = strA0b01_J
        'Modify by Amy 2018/02/06 +if
        If Val(Right(strA0b01_1, 2)) = 12 Then
            MsgBox Label2 & "點數結算期" & Val(Left(strA0b01_1, IIf(Len(strA0b01_1) = 5, 3, 2))) + 1 & "年1月" & _
                                     "財務處尚未確認，故不可操作"
        Else
            MsgBox Label2 & "點數結算期" & Left(strA0b01_1, IIf(Len(strA0b01_1) = 5, 3, 2)) & "年" & Val(Right(strA0b01_1, 2)) + 1 & "月" & _
                                     "財務處尚未確認，故不可操作"
        End If
        txtCloseDate(0).SetFocus
        txtCloseDate_GotFocus 0
        Exit Function
   Else
      bolCancel = False
      Call txtCloseDate_Validate(1, bolCancel)
      If bolCancel = True Then
         txtCloseDate(1).SetFocus
         txtCloseDate_GotFocus 1
         Exit Function
      End If
   End If
   ConstrainCheck = True
End Function

'Mark by Amy 2015/04/22
'Private Sub runWord()
'
'   Dim stYear(0 To 1) As String, stMonth(0 To 1) As String, stTmp As String
'   Dim iRow As Integer, iCol As Integer, lngCellWidth As Long
'   Dim oTable As Word.Table
'
'On Error GoTo ErrHnd
'
'   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
'
'   g_WordAp.Documents.Add
'   g_WordAp.Visible = True
'
'   With g_WordAp.Application
'      .WindowState = wdWindowStateMaximize
'      .Selection.Font.Name = "標楷體"
'      .Selection.PageSetup.Orientation = wdOrientPortrait
'      .Selection.Orientation = wdTextOrientationHorizontal
'      .Selection.Font.Size = 14
'      '邊界
'      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(1.6)
'      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1.4)
'      .Selection.PageSetup.TopMargin = .CentimetersToPoints(2)
'      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
'
'      .Selection.ParagraphFormat.DisableLineHeightGrid = True
'
'      stYear(0) = Val(txtCloseDate(0) \ 100)
'      stMonth(0) = Val(txtCloseDate(0) Mod 100)
'      stYear(1) = Val(txtCloseDate(1) \ 100)
'      stMonth(1) = Val(txtCloseDate(1) Mod 100)
'
'      If stYear(0) = stYear(1) And stMonth(0) = stMonth(1) Then
'         stTmp = stYear(0) & "年度" & stMonth(0) & "月份"
'      ElseIf stYear(0) = stYear(1) Then
'         stTmp = stYear(0) & "年度" & stMonth(0) & "～" & stMonth(1) & "月份"
'      Else
'         stTmp = stYear(0) & "年度" & stMonth(0) & "月～" & stYear(1) & "年度" & stMonth(1) & "月份"
'      End If
'
'
'      .Selection.Font.Size = 18
'      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
'      .Selection.TypeText Text:=stTmp
'
'      .Selection.TypeParagraph
'      .Selection.Font.Size = 14
'      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
'      Select Case txtZone
'         Case "1"
'            strExc(0) = "台北所"
'         Case "2"
'            strExc(0) = "台中所"
'         Case "3"
'            strExc(0) = "台南所"
'         Case "4"
'            strExc(0) = "高雄所"
'      End Select
'      .Selection.TypeText Text:=strExc(0) & "智權人員工作報告"
'
'      .Selection.TypeParagraph
'      .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
'      .Selection.TypeText Text:="一、業績達成情形："
'
'      .Selection.TypeParagraph
'      If grdDataList.Cols - 2 > 7 Then
'         .Selection.Font.Size = 12
'      End If
'      '列數,欄數-2(0,2不印)
'      Set oTable = .Selection.Tables.Add(Range:=.Selection.Range, NumRows:=grdDataList.Rows, NumColumns:=grdDataList.Cols - 2)
'
'      'Added by Morgan 2014/1/28 Word 2007 預設沒有框線需另外指定
'      With oTable
'        .Borders(wdBorderLeft).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
'        .Borders(wdBorderRight).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
'        .Borders(wdBorderTop).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
'        .Borders(wdBorderBottom).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
'        .Borders(wdBorderVertical).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
'        .Borders(wdBorderHorizontal).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
'      End With
'      'end 2014/1/28
'
'      '設定表格高度
'      .Selection.Cells.SetHeight RowHeight:=26, HeightRule:=wdRowHeightExactly
'      For iRow = 0 To grdDataList.Rows - 1
'         .Selection.SelectRow
'         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
'         .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
'         For iCol = 1 To grdDataList.Cols - 1
'            If iCol = 1 Then
'               '收文案源分析
'               If grdDataList.TextMatrix(iRow, iCol) <> grdDataList.TextMatrix(iRow, iCol + 1) Then
'                  '分割欄位
'                  .Selection.MoveLeft Unit:=wdCharacter, Count:=1
'                  lngCellWidth = .Selection.Cells(1).Width
'                  .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
'                  .Selection.Cells(1).SetWidth ColumnWidth:=lngCellWidth * 0.8, RulerStyle:=wdAdjustFirstColumn
'                  .Selection.MoveLeft Unit:=wdCharacter, Count:=1
'                  '只印一次
'                  If iRow = 11 Then
'                     .Selection.TypeText Text:=grdDataList.TextMatrix(iRow, iCol)
'                  End If
'                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
'                  .Selection.TypeText Text:=grdDataList.TextMatrix(iRow, iCol + 1)
'               Else
'                  .Selection.TypeText Text:=grdDataList.TextMatrix(iRow, iCol)
'               End If
'               .Selection.MoveRight Unit:=wdCharacter, Count:=1
'            ElseIf iCol > 2 Then
'               .Selection.TypeText Text:=grdDataList.TextMatrix(iRow, iCol)
'               .Selection.MoveRight Unit:=wdCharacter, Count:=1
'            End If
'         Next
'         .Selection.MoveRight Unit:=wdCharacter, Count:=1
'      Next
'      '合併收文案源分析
'      .Selection.MoveUp Unit:=wdLine, Count:=1
'      .Selection.MoveUp Unit:=wdLine, Count:=3, Extend:=wdExtend
'      .Selection.Cells.Merge
'      .Selection.MoveDown Unit:=wdLine, Count:=1
'
'      .Selection.Font.Size = 14
'      .Selection.TypeText Text:="二、工作報告："
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeText Text:=String(20, "　") & "報告人：" & strUserName & ChangeTStringToTDateString(strSrvDate(2))
'      .Selection.WholeStory
'      .Selection.Font.Name = "Times New Roman"
'      '預設游標停留位置
'      .Selection.Find.ClearFormatting
'      If .Selection.Find.Execute("工作報告：") = True Then
'         .Selection.MoveRight Unit:=wdCharacter, Count:=2 '右移一格
'         .Selection.TypeParagraph
'      End If
'      .Activate
'   End With
'
'ErrHnd:
'
'   If Err.Number <> 0 Then
'      Select Case Err.Number
'         Case 91:
'            g_WordAp.Documents.Add
'            Resume Next
'         Case 462:
'            Set g_WordAp = New Word.Application
'            g_WordAp.Documents.Add
'            Resume Next
'         Case Else:
'            MsgBox "錯誤 : " & Err.Description, vbCritical
'      End Select
'   End If
'End Sub
Private Sub cmdWord_Click()
   Screen.MousePointer = vbHourglass
   runWord
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   
   Dim stST05 As String
   'Added by Lydia 2022/05/03 簡協理69005改為抓系統特殊設定「全所智權部主管」
   Dim strSpecTmp1 As String
   strSpecTmp1 = Pub_GetSpecMan("全所智權部主管")
   If strSpecTmp1 = "" Then strSpecTmp1 = "ABC"
   'end 2022/05/03
   
   stST05 = PUB_GetST05(strUserNum)
   
   MoveFormToCenter Me
   Call SetDataListWidth
   txtZone = pub_strUserOffice
   Select Case strUserNum
      '小真,杜副總可看全部
      '2012/2/14 modify by sonia 加林特助94007權限
      'modify by sonia 2014/6/9 +美珍77027,並取消94007(因己改個人等級)
      'modify by sonia 2019/12/27 杜主秘說加開簡協理69005可看全所
      'Modified by Lydia 2022/05/03
      'Case "65001", "68006", "77027", "69005"
      Case "65001", "68006", "77027", strSpecTmp1
         txtZone.Enabled = True
      Case Else
         Select Case stST05
            '電腦中心,財務,總經理看全部
            '2015/7/28 MODIFY BY SONIA +主任秘書(等級08)可看全部
            Case "00", "01", "08"
               txtZone.Enabled = True
            Case Else
               txtZone.Enabled = False
         End Select
   End Select
   txtCloseDate(0) = (CompDate(1, -1, strSrvDate(2)) - 19110000) \ 100
   txtCloseDate(1) = txtCloseDate(0)
    
End Sub

'p_bolHeaderOnly:是否只設定表頭 true=是 false=資料一併清除
Private Sub SetDataListWidth(Optional ByVal p_bolHeaderOnly As Boolean = False)
   Dim iCol As Integer, iRow As Integer
   'Add by Amy 201/04/22
   'Modify by Amy 2016/03/01 +轉撥及轉撥備註並調顯示位置與「智權點數實績與結餘輸入」全區資料相同
   'Memo by Lydia 2019/08/08 frm210152智權點數實績與結餘輸入在承辦人系統更名為「每月點數查詢／輸入」，與財務系統不同
'   ReDim strRowN(0 To 16)
'   strRowN = Array("區　　別", "目　　標", "達成點數", "達 成 率", "人 員 數", _
'                            "期初實績保留", "期初結餘保留", "當月　　實績", "平均達成", "當月　　結餘", _
'                            "實績保留動用", "結餘保留動用", "期末實績保留", "期末結餘保留", "新增客戶數", _
'                            "收款家數", "收文案源分析")
   ReDim strRowN(0 To 20)
   strRowN = Array("區　　別", "目　　標", "達成點數", "達 成 率", "人 員 數", _
                            "期初實績保留", "當月　　實績", "實績保留動用", "期末實績保留", "轉撥實績增減", "報出實績點數", _
                            "期初結餘保留", "當月　　結餘", "結餘保留動用", "期末結餘保留", "轉撥結餘增減", "報出結餘點數", _
                            "平均達成", "新增客戶數", "收款家數", "收文案源分析")
                            
   With grdDataList
      .Visible = False
      .Clear
      .Rows = 0
      If p_bolHeaderOnly = False Then
'Modify by Amy 2016/03/01 增加rows原20
         .Rows = 24: .Cols = 8: .FixedRows = 1: .FixedCols = 3
         .ColAlignmentFixed = flexAlignCenterCenter
      End If
      .MergeCells = flexMergeRestrictColumns
      .MergeCol(1) = True
      .ColWidth(0) = 0
      .ColWidth(1) = 1000
      .ColWidth(2) = 300
'Modify by Amy 2015/04/22
'      .TextMatrix(0, 1) = "區　　別"
'      .TextMatrix(1, 1) = "目　　標"
'      .TextMatrix(2, 1) = "達成點數"
'      .TextMatrix(3, 1) = "達 成 率"
'      .TextMatrix(4, 1) = "人 員 數"
'      .TextMatrix(5, 1) = "當月達成"
'      .TextMatrix(6, 1) = "平均達成"
'      .TextMatrix(7, 1) = "結餘點數"
'      .TextMatrix(8, 1) = "保留點數"
'      .TextMatrix(9, 1) = "新增客戶數"
'      .TextMatrix(10, 1) = "收款家數"
'      .TextMatrix(11, 1) = "收文" & vbCrLf & "案源分析"
'      For iRow = 12 To 14
'         .TextMatrix(iRow, 1) = .TextMatrix(11, 1)
'      Next
'      For iRow = 0 To 10
'         .MergeRow(iRow) = True
'         .TextMatrix(iRow, 2) = .TextMatrix(iRow, 1)
'      Next
'      .TextMatrix(11, 2) = "P"
'      .TextMatrix(12, 2) = "T"
'      .TextMatrix(13, 2) = "L"
'      .TextMatrix(14, 2) = "C"
      For iRow = 0 To UBound(strRowN)
          .TextMatrix(iRow, 1) = strRowN(iRow)
      Next
      '合併儲存格(收文案源分析)
      For iRow = UBound(strRowN) To UBound(strRowN) + 3
          .TextMatrix(iRow, 1) = .TextMatrix(UBound(strRowN), 1)
      Next
      '合併儲存格(非收文案源分析)
      For iRow = 0 To UBound(strRowN) - 1
         .MergeRow(iRow) = True
         .TextMatrix(iRow, 2) = .TextMatrix(iRow, 1)
      Next
      .TextMatrix(UBound(strRowN), 2) = "專"           'modify by sonia 2019/7/31 P改專,與美珍討論
      .TextMatrix(UBound(strRowN) + 1, 2) = "商"       'modify by sonia 2019/7/31 T改商
      .TextMatrix(UBound(strRowN) + 2, 2) = "法"       'modify by sonia 2019/7/31 L改法
      .TextMatrix(UBound(strRowN) + 3, 2) = "創"       'modify by sonia 2019/7/31 C改創(ACS)
'end 2015/04/22
      .Refresh
      .Visible = True
   End With
End Sub
      
'Modify by Amy 2015/04/22 +欄位
Private Sub runWord()
   Dim stYear(0 To 1) As String, stMonth(0 To 1) As String, stTmp As String
   Dim iRow As Integer, iCol As Integer, lngCellWidth As Long
   Dim oTable As Word.Table
   Dim strFormula As String, bolSetFormula As Boolean '公式/是否設定公式
   Dim strFormat As String 'Add by Amy 2015/05/04
   
On Error GoTo ErrHnd
   
   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
   
   g_WordAp.Documents.add
   g_WordAp.Visible = True
      
   'Modify  by Amy 2016/03/01 點數結算期起迄相同才顯示轉撥備註/標題文字改14/ 內文文字改為12 /提醒文字改10
   With g_WordAp.Application
      .WindowState = wdWindowStateMaximize
      .Selection.Font.Name = "標楷體"
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.Font.Size = 12
      '邊界
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(1.6)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1.4)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
      
      .Selection.ParagraphFormat.DisableLineHeightGrid = True
      
      stYear(0) = Val(txtCloseDate(0) \ 100)
      stMonth(0) = Val(txtCloseDate(0) Mod 100)
      stYear(1) = Val(txtCloseDate(1) \ 100)
      stMonth(1) = Val(txtCloseDate(1) Mod 100)
      
      If stYear(0) = stYear(1) And stMonth(0) = stMonth(1) Then
         stTmp = stYear(0) & "年度" & stMonth(0) & "月份"
      ElseIf stYear(0) = stYear(1) Then
         stTmp = stYear(0) & "年度" & stMonth(0) & "～" & stMonth(1) & "月份"
      Else
         stTmp = stYear(0) & "年度" & stMonth(0) & "月～" & stYear(1) & "年度" & stMonth(1) & "月份"
      End If
      
      .Selection.Font.Size = 14
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.TypeText Text:=stTmp
      
      .Selection.TypeParagraph
      .Selection.Font.Size = 12
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      Select Case txtZone
         Case "1"
            strExc(0) = "台北所"
         Case "2"
            strExc(0) = "台中所"
         Case "3"
            strExc(0) = "台南所"
         Case "4"
            strExc(0) = "高雄所"
      End Select
      .Selection.TypeText Text:=strExc(0) & "智權人員工作報告"
      
      .Selection.TypeParagraph
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
      .Selection.TypeText Text:="一、業績達成情形："
      
      'Modify by Amy 2016/03/01 105年01月抓業績輸入不需顯示此文字
      If Val(txtCloseDate(0)) < Val(業績輸入啟用年月) Then
      .Selection.TypeParagraph
      .Selection.Font.Size = 10
      .Selection.Font.ColorIndex = wdBlue
      .Selection.TypeText Text:="*粗體標示欄位有公式計算,若有調整非標示欄位,則需重新計算 [註]"
      End If
      
      .Selection.TypeParagraph
      .Selection.Font.ColorIndex = wdAuto
      If grdDataList.Cols - 2 > 7 Then
         .Selection.Font.Size = 10
      End If
      '列數,欄數-2(0,2不印)
      Set oTable = .Selection.Tables.add(Range:=.Selection.Range, NumRows:=grdDataList.Rows, NumColumns:=grdDataList.Cols - 2)
      
      'Word 2007 預設沒有框線需另外指定
      With oTable
        .Borders(wdBorderLeft).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
        .Borders(wdBorderRight).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
        .Borders(wdBorderTop).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
        .Borders(wdBorderBottom).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
        .Borders(wdBorderVertical).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
        .Borders(wdBorderHorizontal).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
      End With
      
      '設定表格高度
      .Selection.Cells.SetHeight RowHeight:=26, HeightRule:=wdRowHeightExactly
      For iRow = 0 To grdDataList.Rows - 1
         .Selection.SelectRow
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
         'Modify by Amy 2015/05/04 改合計公式,拿掉轉出保留/結餘、保留/期末結餘保留公式
         For iCol = 1 To grdDataList.Cols - 1
            If iCol = 1 Then
               '收文案源分析
               If grdDataList.TextMatrix(iRow, iCol) <> grdDataList.TextMatrix(iRow, iCol + 1) Then
                  '分割欄位
                  .Selection.MoveLeft Unit:=wdCharacter, Count:=1
                  lngCellWidth = .Selection.Cells(1).Width
                  .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
                  .Selection.Cells(1).SetWidth ColumnWidth:=lngCellWidth * 0.8, RulerStyle:=wdAdjustFirstColumn
                  .Selection.MoveLeft Unit:=wdCharacter, Count:=1
                  '只印一次
                  If iRow = GetValue("收文案源分析") Then
                     .Selection.TypeText Text:=grdDataList.TextMatrix(iRow, iCol)
                  End If
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
                  .Selection.TypeText Text:=grdDataList.TextMatrix(iRow, iCol + 1)
               Else
                  '公式欄位設粗體
                  .Selection.Font.Bold = False
                  'Modify by Amy 2015/04/30 改公式 原公式設「實績保留動用」及「結餘保留動用」
                  If iRow = GetValue("達成點數") Or iRow = GetValue("達 成 率") Then
                      .Selection.Font.Bold = True
'                  ElseIf iRow = GetValue("實績保留動用") Or iRow = GetValue("結餘保留動用") Then
'                      If Val(dblTot(GetValue("期末實績保留"))) <> 0 Or Val(dblTot(GetValue("期末結餘保留"))) <> 0 Then
'                          '財務尚已做傳票,則公式設於「實績保留動用」及「結餘保留動用」
'                          .Selection.Font.Bold = True
'                      End If
'                  ElseIf iRow = GetValue("期末實績保留") Or iRow = GetValue("期末結餘保留") Then
'                      If Val(dblTot(GetValue("期末實績保留"))) = 0 And Val(dblTot(GetValue("期末結餘保留"))) = 0 Then
'                         '財務尚未做傳票,則公式設於「期末實績保留」及「期末結餘保留」
'                         .Selection.Font.Bold = True
'                       End If
                  End If
                  'end 2015/04/30
                  .Selection.TypeText Text:=grdDataList.TextMatrix(iRow, iCol)
                  .Selection.Font.Bold = False
               End If
               .Selection.MoveRight Unit:=wdCharacter, Count:=1
            ElseIf iCol > 2 Then
               strFormula = "": bolSetFormula = False: strFormat = ""
               If grdDataList.TextMatrix(0, iCol) = "合　　計" And iRow > 0 Then
                    Select Case iRow
                        'Modify by Amy 2016/03/01
                        'Case GetValue("達成點數"), GetValue("期初實績保留") To GetValue("期初結餘保留"), GetValue("當月　　實績") To GetValue("期末結餘保留")
                        Case GetValue("達成點數"), GetValue("期初實績保留") To GetValue("轉撥實績增減"), GetValue("期初結餘保留") To GetValue("轉撥結餘增減")
                            strFormula = "SUM(LEFT)"
                            'strFormat = "0.000"
                        'Modify by Amy 2024/10/23  除數 Or 被除數=0,計算達成率會無法跳下一欄
                        Case GetValue("達 成 率")
                            If Val(grdDataList.TextMatrix(GetValue("目　　標"), iCol)) > 0 And Val(grdDataList.TextMatrix(GetValue("達成點數"), iCol)) > 0 Then
                              strFormula = "if(" & Chr(iCol + 63) & GetValue("目　　標") + 1 & "=0,0," & Chr(iCol + 63) & GetValue("達成點數") + 1 & "/" & Chr(iCol + 63) & GetValue("目　　標") + 1 & "*100)"
                              strFormat = "0.000%"
                            End If
                            bolSetFormula = True
                        Case GetValue("平均達成")
                            If Val(grdDataList.TextMatrix(GetValue("人 員 數"), iCol)) > 0 And Val(grdDataList.TextMatrix(GetValue("當月　　實績"), iCol)) > 0 Then
                              strFormula = "if(" & Chr(iCol + 63) & GetValue("人 員 數") + 1 & "=0,0," & Chr(iCol + 63) & GetValue("當月　　實績") + 1 & "/" & Chr(iCol + 63) & GetValue("人 員 數") + 1 & ")"
                              strFormat = "0.000"
                            End If
                        'end 2024/10/23
                        Case Else
                            strFormula = "SUM(LEFT)"
                            'strFormat = "0"
                    End Select
                    bolSetFormula = True
               Else
                    Select Case iRow
                        Case GetValue("達成點數")
                            'Modify by Amy 2016/03/01 +轉撥
                            strFormula = Chr(iCol + 63) & GetValue("期初實績保留") + 1 & "+" & Chr(iCol + 63) & GetValue("期初結餘保留") + 1 & _
                                "+" & Chr(iCol + 63) & GetValue("當月　　實績") + 1 & "+" & Chr(iCol + 63) & GetValue("當月　　結餘") + 1 & _
                                "-" & Chr(iCol + 63) & GetValue("期末實績保留") + 1 & "-" & Chr(iCol + 63) & GetValue("期末結餘保留") + 1 & _
                                "+" & Chr(iCol + 63) & GetValue("轉撥實績增減") + 1 & "+" & Chr(iCol + 63) & GetValue("轉撥結餘增減") + 1
                            strFormat = "0.000"
                            bolSetFormula = True
                        'Modify by Amy 2024/10/23  除數 Or 被除數=0,計算達成率會無法跳下一欄
                        Case GetValue("達 成 率")
                            If Val(grdDataList.TextMatrix(GetValue("目　　標"), iCol)) > 0 And Val(grdDataList.TextMatrix(GetValue("達成點數"), iCol)) > 0 Then
                              strFormula = "if(" & Chr(iCol + 63) & GetValue("目　　標") + 1 & "=0,0," & Chr(iCol + 63) & GetValue("達成點數") + 1 & "/" & Chr(iCol + 63) & GetValue("目　　標") + 1 & "*100)"
                              strFormat = "0.000%"
                            End If
                            bolSetFormula = True
                        Case GetValue("平均達成")
                            If Val(grdDataList.TextMatrix(GetValue("人 員 數"), iCol)) > 0 And Val(grdDataList.TextMatrix(GetValue("當月　　實績"), iCol)) > 0 Then
                              strFormula = "if(" & Chr(iCol + 63) & GetValue("人 員 數") + 1 & "=0,0," & Chr(iCol + 63) & GetValue("當月　　實績") + 1 & "/" & Chr(iCol + 63) & GetValue("人 員 數") + 1 & ")"
                              strFormat = "0.000"
                            End If
                            bolSetFormula = True
                        'end 2024/10/23
                        'Modify by Amy 2015/04/30
    '                    Case GetValue("實績保留動用")
    '                        If Val(dblTot(GetValue("期末實績保留"))) > 0 Or Val(dblTot(GetValue("期末結餘保留"))) > 0 Then
    '                            strFormula = Chr(iCol + 63) & GetValue("期初實績保留") + 1 & "-" & Chr(iCol + 63) & GetValue("期末實績保留") + 1
    '                            strfromat="0.000"
    '                            bolSetFormula = True
    '                        End If
    '                    Case GetValue("結餘保留動用")
    '                        If Val(dblTot(GetValue("期末實績保留"))) > 0 Or Val(dblTot(GetValue("期末結餘保留"))) > 0 Then
    '                            strFormula = Chr(iCol + 63) & GetValue("期初結餘保留") + 1 & "+" & Chr(iCol + 63) & GetValue("當月　　結餘") + 1 & "-" & Chr(iCol + 63) & GetValue("期末結餘保留") + 1
    '                            strformat=& "0.000"
    '                            bolSetFormula = True
    '                        End If
    '                    Case GetValue("期末實績保留")
    '                        If Val(dblTot(GetValue("期末實績保留"))) = 0 And Val(dblTot(GetValue("期末結餘保留"))) = 0 Then
    '                            strFormula = Chr(iCol + 63) & GetValue("期初實績保留") + 1 & "-" & Chr(iCol + 63) & GetValue("實績保留動用") + 1
    '                            strformat=& "0.000"
    '                            bolSetFormula = True
    '                        End If
    '                    Case GetValue("期末結餘保留")
    '                        If Val(dblTot(GetValue("期末實績保留"))) = 0 And Val(dblTot(GetValue("期末結餘保留"))) = 0 Then
    '                            strFormula = Chr(iCol + 63) & GetValue("期初結餘保留") + 1 & "+" & Chr(iCol + 63) & GetValue("當月　　結餘") + 1 & "-" & Chr(iCol + 63) & GetValue("結餘保留動用") + 1
    '                            strformat=& "0.000"
    '                            bolSetFormula = True
    '                        End If
    '                    'end 2015/04/30
                        Case Else
                    End Select
                End If
                If bolSetFormula = True Then
                    .Selection.Font.Bold = True
                    'Modify by Amy 2024/10/23 +if,除數 or 被除數=0會出現發預期之公式
                    If strFormula <> MsgText(601) Then
                        .Selection.InsertFormula Formula:="=" & strFormula, NumberFormat:=strFormat
                    ElseIf Val(strFormula) = 0 Then
                        If iRow = GetValue("達 成 率") Then
                           .Selection.TypeText Text:="0%"
                        Else
                           .Selection.TypeText Text:="0"
                        End If
                    End If
                    'end 2024/10/23
                Else
                    .Selection.Font.Bold = False
                    .Selection.TypeText Text:=grdDataList.TextMatrix(iRow, iCol)
                End If
               .Selection.MoveRight Unit:=wdCharacter, Count:=1
            End If
         Next iCol
         'end 2015/05/04
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
      Next iRow
      '合併收文案源分析
      .Selection.MoveUp Unit:=wdLine, Count:=1
      .Selection.MoveUp Unit:=wdLine, Count:=3, Extend:=wdExtend
      .Selection.Cells.Merge
      .Selection.MoveDown Unit:=wdLine, Count:=1
      
      .Selection.Font.Size = 12
      .Selection.TypeText Text:="二、工作報告："
      .Selection.TypeParagraph
      .Selection.TypeText Text:=String(30, "　") & "報告人：" & strUserName & ChangeTStringToTDateString(strSrvDate(2))
      'Modify by Amy 2016/03/01 105年01月抓業績輸入不需顯示此文字
      If Val(txtCloseDate(0)) < Val(業績輸入啟用年月) Then
      Selection.TypeParagraph
      .Selection.Font.Size = 10
      .Selection.Font.ColorIndex = wdBlue
      .Selection.TypeText Text:="[註] 重新計算方式:按鍵盤「Ctrl」+「A」(全選)後,再按「F9」。"
      End If
      .Selection.Font.Size = 12
      .Selection.Font.ColorIndex = wdAuto
      .Selection.WholeStory
      '功能變數更新
      .Selection.Fields.UPDATE
      .Selection.Font.Name = "Times New Roman"
      '預設游標停留位置
      .Selection.Find.ClearFormatting
      If .Selection.Find.Execute("工作報告：") = True Then
         .Selection.MoveRight Unit:=wdCharacter, Count:=2 '右移一格
         .Selection.TypeParagraph
      End If
      .Activate
   End With
   
ErrHnd:

   If Err.Number <> 0 Then
      Select Case Err.Number
         Case 91:
            g_WordAp.Documents.add
            Resume Next
         Case 462:
            Set g_WordAp = New Word.Application
            g_WordAp.Documents.add
            Resume Next
         Case Else:
            MsgBox "錯誤 : " & Err.Description, vbCritical
      End Select
   End If
End Sub

'Add by Amy 2018/04/24
Private Function doQuery() As Boolean
   Dim strQ As String
   Dim iRow As Single, iCol As Single
   Dim dblTotCase(3) As Double, dblSumR As Double
   Dim stConCu As String, stConCP As String
   Dim strSalesArea1 As String, strSalesArea2 As String
   Dim i As Integer, j As Integer
   Dim strSqlSP As String, intQ As Integer
   Dim dblSP15 As Double, dblSP36 As Double
   Erase dblTot
   Dim dblSP19 As Double, dblSP40 As Double

   '所別
   If txtZone <> "" Then
        strSalesArea1 = "S" & txtZone
   End If
    
On Error GoTo ErrHnd
    '刪除暫存檔資料
    strQ = "Delete From R210113 Where ID='" & strUserNum & "' And R04='2'"
    adoTaie.Execute strQ
    
    '將SalesPoint人員寫入暫存檔(有控制只可以查SalesPoint上線後的資料)
    strQ = "Insert Into R210113 (ID,R01,R02,R03,R04) " & _
              "Select Distinct '" & strUserNum & "',SP02,SP48,SP01,'2' From SalesPoint " & _
              "Where SP01>=191100+" & Val(txtCloseDate(0)) & " And SP01<=191100+" & Val(txtCloseDate(1)) & _
              " And SubStr(SP48,1,2)='" & strSalesArea1 & "' "
    adoTaie.Execute strQ
 
    strSql = GetPoint(0, Val(txtCloseDate(0)), Val(txtCloseDate(1)), strSalesArea1, , , , Me.Name)
    strSqlSP = GetPoint_SP(Val(txtCloseDate(0)), Val(txtCloseDate(1)), strSalesArea1, , , , Me.Name, True)

    intI = 1: intQ = 1
    Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strSql)
    Set adoRecordset1 = ClsLawReadRstMsg(intQ, strSqlSP)
    If intI = 1 Then
       ' 數值欄位以整數顯示,有小數才顯示小數/增加欄位(同frm210152智權點數實績與結餘輸入的全區資料)
       'Memo by Lydia 2019/08/08 frm210152智權點數實績與結餘輸入在承辦人系統更名為「每月點數查詢／輸入」，與財務系統不同
       grdDataList.Visible = False
       With AdoRecordSet3
        iCol = 3
       Erase dblTot
       Do While Not .EOF
          dblSumR = 0: dblSP15 = 0: dblSP36 = 0
          dblSP19 = 0: dblSP40 = 0
          grdDataList.Cols = iCol + 1
          '區別
          iRow = GetValue("區　　別")
          grdDataList.TextMatrix(iRow, iCol) = "" & .Fields("NAME")
          grdDataList.CellAlignment = flexAlignCenterCenter
          '目標
          iRow = GetValue("目　　標")
          grdDataList.TextMatrix(iRow, iCol) = "" & .Fields("PE04")
          grdDataList.CellAlignment = flexAlignRightCenter
          dblTot(iRow) = dblTot(iRow) + Val("" & .Fields("PE04"))
          '人員數
          iRow = GetValue("人 員 數")
          grdDataList.TextMatrix(iRow, iCol) = "" & .Fields("C13")
          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
          If Val(txtCloseDate(0)) >= Val(業績輸入啟用年月) Then
            If intQ = 1 Then
                If Not adoRecordset1.EOF Then
                    If "" & .Fields("ST15") = "" & adoRecordset1.Fields("R02") Then
                        dblSP15 = Val("" & adoRecordset1.Fields("SP15"))
                        dblSP36 = Val("" & adoRecordset1.Fields("SP36"))
                        dblSP19 = Val("" & adoRecordset1.Fields("SP19"))
                        dblSP40 = Val("" & adoRecordset1.Fields("SP40"))
                        adoRecordset1.MoveNext
                    End If
                End If
            Else
                dblSP15 = Val("" & .Fields("C5"))
                dblSP36 = Val("" & .Fields("C6"))
            End If
         End If

          '期初實績保留
          iRow = GetValue("期初實績保留")
          grdDataList.TextMatrix(iRow, iCol) = Round(Val("" & .Fields("C1")), 3)
          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
          dblSumR = dblSumR + Val(grdDataList.TextMatrix(iRow, iCol))
           '當月　　實績
          iRow = GetValue("當月　　實績")
           grdDataList.TextMatrix(iRow, iCol) = Round(Val("" & .Fields("C3")), 3)
          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
          dblSumR = dblSumR + Val(grdDataList.TextMatrix(iRow, iCol))
           '期末實績保留
          iRow = GetValue("期末實績保留")
          grdDataList.TextMatrix(iRow, iCol) = Round(Val(dblSP15), 3)
          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
          dblSumR = dblSumR - Val(grdDataList.TextMatrix(iRow, iCol))
          
          If Val(txtCloseDate(0)) >= Val(業績輸入啟用年月) Then
            iRow = GetValue("轉撥實績增減")
            grdDataList.TextMatrix(iRow, iCol) = Round(Val(dblSP19), 3)
            dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
          End If
          '報出實績點數
          iRow = GetValue("報出實績點數")
          grdDataList.TextMatrix(iRow, iCol) = Round(Val("" & .Fields("C1")) + Val("" & .Fields("C3")) - Val(dblSP15) + Val(dblSP19), 3)
          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
          
          '期初結餘保留
          iRow = GetValue("期初結餘保留")
          grdDataList.TextMatrix(iRow, iCol) = Round(Val("" & .Fields("C2")), 3)
          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
          dblSumR = dblSumR + Val(grdDataList.TextMatrix(iRow, iCol))
          '當月　　結餘
          iRow = GetValue("當月　　結餘")
          grdDataList.TextMatrix(iRow, iCol) = Round(Val("" & .Fields("C4")), 3)
          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
          dblSumR = dblSumR + Val(grdDataList.TextMatrix(iRow, iCol))
          '期末結餘保留
          iRow = GetValue("期末結餘保留")
          grdDataList.TextMatrix(iRow, iCol) = Round(Val(dblSP36), 3)
          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
          dblSumR = dblSumR - Val(grdDataList.TextMatrix(iRow, iCol))
          
          If Val(txtCloseDate(0)) >= Val(業績輸入啟用年月) Then
            iRow = GetValue("轉撥結餘增減")
            grdDataList.TextMatrix(iRow, iCol) = Round(Val(dblSP40), 3)
            dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
          End If
          '報出結餘點數
          iRow = GetValue("報出結餘點數")
          grdDataList.TextMatrix(iRow, iCol) = Round(Val("" & .Fields("C2")) + Val("" & .Fields("C4")) - Val(dblSP36) + Val(dblSP40), 3)
          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
         
          '新增客戶數
          iRow = GetValue("新增客戶數")
          grdDataList.TextMatrix(iRow, iCol) = Val("" & .Fields("C7"))
          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
          '收款家數
          iRow = GetValue("收款家數")
          grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C8")))
          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
          '收文案件來源分析
          'P
           iRow = GetValue("收文案源分析")
          grdDataList.TextMatrix(iRow, iCol) = Val("" & .Fields("C9"))
          dblTotCase(0) = dblTotCase(0) + Val(grdDataList.TextMatrix(iRow, iCol))
          'T
          iRow = GetValue("收文案源分析") + 1
          grdDataList.TextMatrix(iRow, iCol) = Val("" & .Fields("C10"))
          dblTotCase(1) = dblTotCase(1) + Val(grdDataList.TextMatrix(iRow, iCol))
          'L
          iRow = GetValue("收文案源分析") + 2
          grdDataList.TextMatrix(iRow, iCol) = Val("" & .Fields("C11"))
          dblTotCase(2) = dblTotCase(2) + Val(grdDataList.TextMatrix(iRow, iCol))
          'C
          iRow = GetValue("收文案源分析") + 3
          grdDataList.TextMatrix(iRow, iCol) = Val("" & .Fields("C12"))
          dblTotCase(3) = dblTotCase(3) + Val(grdDataList.TextMatrix(iRow, iCol))

          '達成點數
          iRow = GetValue("達成點數")
          strExc(0) = Val(grdDataList.TextMatrix(GetValue("期初實績保留"), iCol)) + Val(grdDataList.TextMatrix(GetValue("期初結餘保留"), iCol))
          strExc(0) = Val(strExc(0)) + Val(grdDataList.TextMatrix(GetValue("當月　　實績"), iCol)) + Val(grdDataList.TextMatrix(GetValue("當月　　結餘"), iCol))
          strExc(0) = Val(strExc(0)) - Val(grdDataList.TextMatrix(GetValue("期末實績保留"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末結餘保留"), iCol))
          '轉撥
          strExc(0) = Val(strExc(0)) + Val(grdDataList.TextMatrix(GetValue("轉撥實績增減"), iCol)) + Val(grdDataList.TextMatrix(GetValue("轉撥結餘增減"), iCol))
          grdDataList.TextMatrix(iRow, iCol) = Round(Val(strExc(0)), 3)
          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))

          '達成率
          iRow = GetValue("達 成 率")
          If Val(grdDataList.TextMatrix(GetValue("目　　標"), iCol)) > 0 And Val(grdDataList.TextMatrix(GetValue("達成點數"), iCol)) > 0 Then
             grdDataList.TextMatrix(iRow, iCol) = Format(100 * Val(grdDataList.TextMatrix(GetValue("達成點數"), iCol)) / Val(grdDataList.TextMatrix(GetValue("目　　標"), iCol)), "0.000") & "%"
          Else
             grdDataList.TextMatrix(iRow, iCol) = "0%"
          End If
          '平均達成
          iRow = GetValue("平均達成")
          If Val(grdDataList.TextMatrix(GetValue("人 員 數"), iCol)) > 0 Then
            grdDataList.TextMatrix(iRow, iCol) = Format(Val(grdDataList.TextMatrix(GetValue("當月　　實績"), iCol)) / Val(grdDataList.TextMatrix(GetValue("人 員 數"), iCol)), "0.000")
          Else
            grdDataList.TextMatrix(iRow, iCol) = "0"
         End If
         '實績保留動用
         iRow = GetValue("實績保留動用")
        grdDataList.TextMatrix(iRow, iCol) = Round(Val(grdDataList.TextMatrix(GetValue("期初實績保留"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末實績保留"), iCol)), 3)
        dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))

         '結餘保留動用
         iRow = GetValue("結餘保留動用")
        grdDataList.TextMatrix(iRow, iCol) = Round(Val(grdDataList.TextMatrix(GetValue("期初結餘保留"), iCol)) + Val(grdDataList.TextMatrix(GetValue("當月　　結餘"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末結餘保留"), iCol)), 3)
        dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))

         .MoveNext
         iCol = iCol + 1
      Loop
      End With

      grdDataList.Cols = iCol + 1
      grdDataList.TextMatrix(0, iCol) = "合　　計"
      For i = 1 To UBound(strRowN)
         Select Case i
            Case GetValue("目　　標")
                grdDataList.TextMatrix(i, iCol) = dblTot(i)
            Case GetValue("達成點數")
                grdDataList.TextMatrix(i, iCol) = dblTot(i)
                For j = 1 To grdDataList.Cols - 1
                    grdDataList.col = j
                    grdDataList.row = i
                    grdDataList.CellBackColor = &HC000&
                Next j
            Case GetValue("達 成 率")
                If Val(grdDataList.TextMatrix(GetValue("目　　標"), iCol)) > 0 And Val(grdDataList.TextMatrix(GetValue("達成點數"), iCol)) Then
                    grdDataList.TextMatrix(i, iCol) = Format(100 * Val(grdDataList.TextMatrix(GetValue("達成點數"), iCol)) / Val(grdDataList.TextMatrix(GetValue("目　　標"), iCol)), "0.000") & "%"
                Else
                    grdDataList.TextMatrix(i, iCol) = "0%"
                End If
            Case GetValue("人 員 數")
                grdDataList.TextMatrix(i, iCol) = dblTot(i)
            Case GetValue("平均達成")
                If Val(grdDataList.TextMatrix(GetValue("人 員 數"), iCol)) > 0 And Val(grdDataList.TextMatrix(GetValue("當月　　實績"), iCol)) > 0 Then
                    grdDataList.TextMatrix(i, iCol) = Format(Val(grdDataList.TextMatrix(GetValue("當月　　實績"), iCol)) / Val(grdDataList.TextMatrix(GetValue("人 員 數"), iCol)), "0.000")
                Else
                    grdDataList.TextMatrix(i, iCol) = "0"
                End If
            Case GetValue("期末實績保留"), GetValue("期末結餘保留")
                If Val(dblTot(GetValue("期末實績保留"))) = 0 And Val(dblTot(GetValue("期末結餘保留"))) = 0 Then
                    '財務尚未做傳票,則實績保留動用及餘額需設 0 讓智權人員填
                    If i = GetValue("期末實績保留") Then
                        '實績保留動用及餘額需設 0 (設一次就好)
                        For j = 3 To grdDataList.Cols - 1
                            grdDataList.TextMatrix(GetValue("實績保留動用"), j) = 0
                            grdDataList.TextMatrix(GetValue("結餘保留動用"), j) = 0
                        Next j
                    End If
                End If
                grdDataList.TextMatrix(i, iCol) = dblTot(i)
             'Add by Amy 2015/03/01
            Case GetValue("報出實績點數"), GetValue("報出結餘點數")
                grdDataList.TextMatrix(i, iCol) = dblTot(i)
                For j = 1 To grdDataList.Cols - 1
                    grdDataList.col = j
                    grdDataList.row = i
                    grdDataList.CellBackColor = &HC000&
                Next j
            Case GetValue("新增客戶數") To GetValue("收款家數")
                grdDataList.TextMatrix(i, iCol) = dblTot(i)
            Case GetValue("收文案源分析")
                grdDataList.TextMatrix(i, iCol) = dblTotCase(0)
                grdDataList.TextMatrix(i + 1, iCol) = dblTotCase(1)
                grdDataList.TextMatrix(i + 2, iCol) = dblTotCase(2)
                grdDataList.TextMatrix(i + 3, iCol) = dblTotCase(3)
            Case Else
                grdDataList.TextMatrix(i, iCol) = dblTot(i)
         End Select
      Next i
      grdDataList.Visible = True
   Else
      MsgBox "無符合資料！", vbInformation
   End If

   doQuery = True

ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

Private Function doQuery_Old3() As Boolean
'Memo 2016/02/05 抓目標時業務區不知如何抓舊資料,且日期跨上線日期也不易抓資料,故不可查舊資料
'   Dim stConST As String, stConPE As String
'   Dim stVTB(13) As String
'   Dim iRow As Single, iCol As Single
'   Dim dblTotCase(3) As Double, dblSumR As Double
'   Dim stConCu As String, stConCP As String
'   Dim strSalesArea1 As String, strSalesArea2 As String
'   Dim i As Integer, j As Integer
'   Dim strSqlSP As String, intQ As Integer
'   Dim dblSP15 As Double, dblSP36 As Double
'   Erase dblTot
'   'Add by Amy 2016/03/01
'    Dim dblSP19 As Double, dblSP40 As Double
'
'   'Modify by Amy 2016/03/11
'   '所別
'   If txtZone <> "" Then
'        'stConST = " And ST06='" & txtZone & "' And SubStr(ST15,1,2)>='S" & txtZone & "' And SubStr(ST15,1,2)<'S" & txtZone + 1 & "'"
'        strSalesArea1 = "S" & txtZone
'        strSalesArea2 = "S" & (Val(txtZone) + 1) * 10 - 1
'   End If
'
''   '點數結算日
''   If txtCloseDate(0) <> "" Then
''      stConPE = stConPE & " And PE03(+) >= 191100+" & txtCloseDate(0)
''   End If
''   If txtCloseDate(1) <> "" Then
''      stConPE = stConPE & " And PE03(+) <= 191100+" & txtCloseDate(1)
''   End If
'    'end 2016/03/11
'
'On Error GoTo ErrHnd
'    '人員數計算排除 S開頭的員工編號
'    'Modify by Amy 2017/10/17 因中一高國碩/陳頌恩轉中二調整程式,否則抓10609年中所資料會有問題
'    'strSql = "Select ST15,min(a0902) NAME,Sum(PE04) PE04,Sum(C1) C1,Sum(C2) C2,Sum(C3) C3,Sum(C4) C4,Sum(C5) C5,Sum(C6) C6" & _
'        ", Sum(C7) C7, Sum(C8) C8, Sum(C9) C9, Sum(C10) C10, Sum(C11) C11, Sum(C12) C12, Count(Distinct Decode(SubStr(st01,1,1),'S',null,st01)) C13" & _
'        " From (" & GetPoint(0, Val(txtCloseDate(0)), Val(txtCloseDate(1)), strSalesArea1, strSalesArea2, , , Me.Name) & "),Staff,Acc090" & _
'        " Where st01(+)=ID And a0901(+)=st15 Group by st15 Order by st15"
'    strSql = "Select SP48 as ST15,min(a0902) NAME,Sum(PE04) PE04,Sum(C1) C1,Sum(C2) C2,Sum(C3) C3,Sum(C4) C4,Sum(C5) C5,Sum(C6) C6" & _
'        ", Sum(C7) C7, Sum(C8) C8, Sum(C9) C9, Sum(C10) C10, Sum(C11) C11, Sum(C12) C12, Count(Distinct Decode(SubStr(ID,1,1),'S',null,ID)) C13" & _
'        " From (" & GetPoint(0, Val(txtCloseDate(0)), Val(txtCloseDate(1)), strSalesArea1, strSalesArea2, , , Me.Name) & "),Acc090" & _
'        " Where a0901(+)=SP48 Group by SP48 Order by SP48"
'
'    strSqlSP = GetPoint_SP(Val(txtCloseDate(0)), Val(txtCloseDate(1)), "S" & txtZone & "0", "S" & txtZone & "9", , , Me.Name, True)
'
'    intI = 1: intQ = 1
'    Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strSql)
'    Set adoRecordset1 = ClsLawReadRstMsg(intQ, strSqlSP)
'    If intI = 1 Then
'       'Modify by Amy 2016/03/01 數值欄位以整數顯示,有小數才顯示小數/增加欄位(同frm210152智權點數實績與結餘輸入的全區資料)
'       grdDataList.Visible = False
'       With AdoRecordSet3
'        iCol = 3
'       Erase dblTot
'       Do While Not .EOF
'          dblSumR = 0: dblSP15 = 0: dblSP36 = 0
'          dblSP19 = 0: dblSP40 = 0 'Add  by Amy 2016/12/14
'          grdDataList.Cols = iCol + 1
'          '區別
'          iRow = GetValue("區　　別")
'          grdDataList.TextMatrix(iRow, iCol) = "" & .Fields("NAME")
'          grdDataList.CellAlignment = flexAlignCenterCenter
'          '目標
'          iRow = GetValue("目　　標")
'          grdDataList.TextMatrix(iRow, iCol) = "" & .Fields("PE04")
'          grdDataList.CellAlignment = flexAlignRightCenter
'          '人員數
'          iRow = GetValue("人 員 數")
'          grdDataList.TextMatrix(iRow, iCol) = "" & .Fields("C13")
'          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'          'Modify by Amy 2016/03/01 調位置
'          If Val(txtCloseDate(0)) >= Val(業績輸入啟用年月) Then
'            If intQ = 1 Then
'                'Add by Amy 2016/12/14 +if 中所10511月無20091時會error
'                If Not adoRecordset1.EOF Then
'                    If "" & .Fields("ST15") = "" & adoRecordset1.Fields("SP48") Then
'                        dblSP15 = Val("" & adoRecordset1.Fields("SP15"))
'                        dblSP36 = Val("" & adoRecordset1.Fields("SP36"))
'                        'Add by Amy 2016/03/01
'                        dblSP19 = Val("" & adoRecordset1.Fields("SP19"))
'                        dblSP40 = Val("" & adoRecordset1.Fields("SP40"))
'                        adoRecordset1.MoveNext
'                    End If
'                End If
'            Else
'                dblSP15 = Val("" & .Fields("C5"))
'                dblSP36 = Val("" & .Fields("C6"))
'            End If
'         End If
'
'          '期初實績保留
'          iRow = GetValue("期初實績保留")
'          grdDataList.TextMatrix(iRow, iCol) = Round(Val("" & .Fields("C1")), 3) 'Format(Val("" & .Fields("C1")), "0.000")
'          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'          dblSumR = dblSumR + Val(grdDataList.TextMatrix(iRow, iCol))
'           '當月　　實績
'          iRow = GetValue("當月　　實績")
'           grdDataList.TextMatrix(iRow, iCol) = Round(Val("" & .Fields("C3")), 3) 'Format(Val("" & .Fields("C3")), "0.000")
'          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'          dblSumR = dblSumR + Val(grdDataList.TextMatrix(iRow, iCol))
'           '期末實績保留
'          iRow = GetValue("期末實績保留")
'          grdDataList.TextMatrix(iRow, iCol) = Round(dblSP15, 3) 'Format(dblSP15, "0.000")
'          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'          dblSumR = dblSumR - Val(grdDataList.TextMatrix(iRow, iCol))
'          'Add by Amy 2015/02/15
'          If Val(txtCloseDate(0)) >= Val(業績輸入啟用年月) Then
'            iRow = GetValue("轉撥實績增減")
'            grdDataList.TextMatrix(iRow, iCol) = Round(dblSP19, 3) 'Format(dblSP19, "0.000")
'            dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'          End If
'          '報出實績點數
'          iRow = GetValue("報出實績點數")
'          grdDataList.TextMatrix(iRow, iCol) = Round(Val("" & .Fields("C1")) + Val("" & .Fields("C3")) - Val(dblSP15) + Val(dblSP19), 3)
'          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'
'          '期初結餘保留
'          iRow = GetValue("期初結餘保留")
'          grdDataList.TextMatrix(iRow, iCol) = Round(Val("" & .Fields("C2")), 3) 'Format(Val("" & .Fields("C2")), "0.000")
'          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'          dblSumR = dblSumR + Val(grdDataList.TextMatrix(iRow, iCol))
'          '當月　　結餘
'          iRow = GetValue("當月　　結餘")
'          grdDataList.TextMatrix(iRow, iCol) = Round(Val("" & .Fields("C4")), 3) 'Format(Val("" & .Fields("C4")), "0.000")
'          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'          dblSumR = dblSumR + Val(grdDataList.TextMatrix(iRow, iCol))
'          '期末結餘保留
'          iRow = GetValue("期末結餘保留")
'          grdDataList.TextMatrix(iRow, iCol) = Round(dblSP36, 3) 'Format(dblSP36, "0.000")
'          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'          dblSumR = dblSumR - Val(grdDataList.TextMatrix(iRow, iCol))
'          'Add by Amy 2015/02/15
'          If Val(txtCloseDate(0)) >= Val(業績輸入啟用年月) Then
'            iRow = GetValue("轉撥結餘增減")
'            grdDataList.TextMatrix(iRow, iCol) = Round(dblSP40, 3) 'Format(dblSP40, "0.000")
'            dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'          End If
'          '報出結餘點數
'          iRow = GetValue("報出結餘點數")
'          grdDataList.TextMatrix(iRow, iCol) = Round(Val("" & .Fields("C2")) + Val("" & .Fields("C4")) - Val(dblSP36) + Val(dblSP40), 3)
'          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'
'          '新增客戶數
'          iRow = GetValue("新增客戶數")
'          grdDataList.TextMatrix(iRow, iCol) = Val("" & .Fields("C7")) 'Format(Val("" & .Fields("C7")))
'          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'          '收款家數
'          iRow = GetValue("收款家數")
'          grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C8")))
'          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'          '收文案件來源分析
'          'P
'           iRow = GetValue("收文案源分析")
'          grdDataList.TextMatrix(iRow, iCol) = Val("" & .Fields("C9")) 'Format(Val("" & .Fields("C9")))
'          dblTotCase(0) = dblTotCase(0) + Val(grdDataList.TextMatrix(iRow, iCol))
'          'T
'          iRow = GetValue("收文案源分析") + 1
'          grdDataList.TextMatrix(iRow, iCol) = Val("" & .Fields("C10"))
'          dblTotCase(1) = dblTotCase(1) + Val(grdDataList.TextMatrix(iRow, iCol))
'          'L
'          iRow = GetValue("收文案源分析") + 2
'          grdDataList.TextMatrix(iRow, iCol) = Val("" & .Fields("C11"))
'          dblTotCase(2) = dblTotCase(2) + Val(grdDataList.TextMatrix(iRow, iCol))
'          'C
'          iRow = GetValue("收文案源分析") + 3
'          grdDataList.TextMatrix(iRow, iCol) = Val("" & .Fields("C12"))
'          dblTotCase(3) = dblTotCase(3) + Val(grdDataList.TextMatrix(iRow, iCol))
'
'          '達成點數
'          iRow = GetValue("達成點數")
'          strExc(0) = Val(grdDataList.TextMatrix(GetValue("期初實績保留"), iCol)) + Val(grdDataList.TextMatrix(GetValue("期初結餘保留"), iCol))
'          strExc(0) = Val(strExc(0)) + Val(grdDataList.TextMatrix(GetValue("當月　　實績"), iCol)) + Val(grdDataList.TextMatrix(GetValue("當月　　結餘"), iCol))
'          strExc(0) = Val(strExc(0)) - Val(grdDataList.TextMatrix(GetValue("期末實績保留"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末結餘保留"), iCol))
'          'Add by Amy 2016/03/01 +轉撥
'          strExc(0) = Val(strExc(0)) + Val(grdDataList.TextMatrix(GetValue("轉撥實績增減"), iCol)) + Val(grdDataList.TextMatrix(GetValue("轉撥結餘增減"), iCol))
'          grdDataList.TextMatrix(iRow, iCol) = Round(strExc(0), 3) 'Format(strExc(0), "0.000")
'          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'
'          '達成率
'          iRow = GetValue("達 成 率")
'          If Val(grdDataList.TextMatrix(GetValue("目　　標"), iCol)) > 0 And Val(grdDataList.TextMatrix(GetValue("達成點數"), iCol)) > 0 Then
'             grdDataList.TextMatrix(iRow, iCol) = Format(100 * Val(grdDataList.TextMatrix(GetValue("達成點數"), iCol)) / Val(grdDataList.TextMatrix(GetValue("目　　標"), iCol)), "0.000") & "%"
'          Else
'             grdDataList.TextMatrix(iRow, iCol) = "0%"
'          End If
'          '平均達成
'          iRow = GetValue("平均達成")
'          If Val(grdDataList.TextMatrix(GetValue("人 員 數"), iCol)) > 0 Then
'            grdDataList.TextMatrix(iRow, iCol) = Format(Val(grdDataList.TextMatrix(GetValue("當月　　實績"), iCol)) / Val(grdDataList.TextMatrix(GetValue("人 員 數"), iCol)), "0.000")
'          Else
'            grdDataList.TextMatrix(iRow, iCol) = "0"
'         End If
'         '實績保留動用
'         iRow = GetValue("實績保留動用")
'        grdDataList.TextMatrix(iRow, iCol) = Round(Val(grdDataList.TextMatrix(GetValue("期初實績保留"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末實績保留"), iCol)), 3)
'        dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'
'         '結餘保留動用
'         iRow = GetValue("結餘保留動用")
'        grdDataList.TextMatrix(iRow, iCol) = Round(Val(grdDataList.TextMatrix(GetValue("期初結餘保留"), iCol)) + Val(grdDataList.TextMatrix(GetValue("當月　　結餘"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末結餘保留"), iCol)), 3)
'        dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'
'         .MoveNext
'         iCol = iCol + 1
'      Loop
'      End With
'
'      grdDataList.Cols = iCol + 1
'      grdDataList.TextMatrix(0, iCol) = "合　　計"
'      '總目標需抓該區所有(因為有可能點數掛該區代碼如20021)
'      'Modify by Amy 2016/03/11 改抓函數
''      strExc(0) = "Select nvl(sum(PE04),0) PE04 From Staff,PerFormance " & _
''         "Where PE01(+)=ST01 And PE02(+)='TOT'" & stConST & stConPE
'      strExc(0) = GetPoint(3, Val(txtCloseDate(0)), Val(txtCloseDate(1)), "S" & txtZone & "0", "S" & txtZone & "9", , , Me.Name)
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         dblTot(1) = RsTemp.Fields(0)
'      End If
'      For i = 1 To UBound(strRowN)
'         Select Case i
'            Case GetValue("目　　標")
'                grdDataList.TextMatrix(i, iCol) = dblTot(i)
'            Case GetValue("達成點數")
'                grdDataList.TextMatrix(i, iCol) = dblTot(i) 'Format(dblTot(i), "0.000")
'                For j = 1 To grdDataList.Cols - 1
'                    grdDataList.col = j
'                    grdDataList.row = i
'                    grdDataList.CellBackColor = &HC000&
'                Next j
'            Case GetValue("達 成 率")
'                If Val(grdDataList.TextMatrix(GetValue("目　　標"), iCol)) > 0 And Val(grdDataList.TextMatrix(GetValue("達成點數"), iCol)) Then
'                    grdDataList.TextMatrix(i, iCol) = Format(100 * Val(grdDataList.TextMatrix(GetValue("達成點數"), iCol)) / Val(grdDataList.TextMatrix(GetValue("目　　標"), iCol)), "0.000") & "%"
'                Else
'                    grdDataList.TextMatrix(i, iCol) = "0%"
'                End If
'            Case GetValue("人 員 數")
'                grdDataList.TextMatrix(i, iCol) = dblTot(i)
'            Case GetValue("平均達成")
'                If Val(grdDataList.TextMatrix(GetValue("人 員 數"), iCol)) > 0 And Val(grdDataList.TextMatrix(GetValue("當月　　實績"), iCol)) > 0 Then
'                    grdDataList.TextMatrix(i, iCol) = Format(Val(grdDataList.TextMatrix(GetValue("當月　　實績"), iCol)) / Val(grdDataList.TextMatrix(GetValue("人 員 數"), iCol)), "0.000")
'                Else
'                    grdDataList.TextMatrix(i, iCol) = "0"
'                End If
'            'Modify by Amy 2015/02/15 拆位置 原GetValue("期末實績保留") To GetValue("期末結餘保留")
'            Case GetValue("期末實績保留"), GetValue("期末結餘保留")
'                If Val(dblTot(GetValue("期末實績保留"))) = 0 And Val(dblTot(GetValue("期末結餘保留"))) = 0 Then
'                    '財務尚未做傳票,則實績保留動用及餘額需設 0 讓智權人員填
'                    If i = GetValue("期末實績保留") Then
'                        '實績保留動用及餘額需設 0 (設一次就好)
'                        For j = 3 To grdDataList.Cols - 1
'                            grdDataList.TextMatrix(GetValue("實績保留動用"), j) = 0 'Format("0", "0.000")
'                            grdDataList.TextMatrix(GetValue("結餘保留動用"), j) = 0 'Format("0", "0.000")
'                        Next j
'                    End If
'                End If
'                grdDataList.TextMatrix(i, iCol) = dblTot(i) 'Format(dblTot(i), "0.000")
'             'Add by Amy 2015/03/01
'            Case GetValue("報出實績點數"), GetValue("報出結餘點數")
'                grdDataList.TextMatrix(i, iCol) = dblTot(i)
'                For j = 1 To grdDataList.Cols - 1
'                    grdDataList.col = j
'                    grdDataList.row = i
'                    grdDataList.CellBackColor = &HC000&
'                Next j
'            Case GetValue("新增客戶數") To GetValue("收款家數")
'                grdDataList.TextMatrix(i, iCol) = dblTot(i)
'            Case GetValue("收文案源分析")
'                grdDataList.TextMatrix(i, iCol) = dblTotCase(0)
'                grdDataList.TextMatrix(i + 1, iCol) = dblTotCase(1)
'                grdDataList.TextMatrix(i + 2, iCol) = dblTotCase(2)
'                grdDataList.TextMatrix(i + 3, iCol) = dblTotCase(3)
'            Case Else
'                grdDataList.TextMatrix(i, iCol) = dblTot(i) 'Format(dblTot(i), "0.000")
'         End Select
'      Next i
'      grdDataList.Visible = True
'      'end 2016/03/01
'   Else
'      MsgBox "無符合資料！", vbInformation
'   End If
'
'   doQuery = True
'
'ErrHnd:
'   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

Private Function doQuery_Old() As Boolean
'Mark by Amy 2016/02/05 語法改抓共用函數
'   Dim stCon As String, stConST As String, stConR1 As String, stConR2 As String, stConPE As String
'   Dim stVTB(13) As String
'   Dim iRow As Single, iCol As Single
'   Dim dblTotCase(3) As Double, dblSumR As Double
'   Dim stConCu As String, stConCP As String
'   Dim i As Integer, j As Integer
'   'Add by Amy 2016/01/30
'   Dim strSqlSP As String, intQ As Integer
'   Dim dblSP15 As Double, dblSP36 As Double
'
'   Erase dblTot
'   stCon = "": stConST = "": stConR1 = "": stConR2 = "": stConPE = "": stConCu = ""
'
'   '所別
'   If txtZone <> "" Then
'      stConST = " And ST06='" & txtZone & "' And SubStr(ST15,1,2)>='S" & txtZone & "' And SubStr(ST15,1,2)<'S" & txtZone + 1 & "'"
'      stConCu = stConCu & " And SubStr(CU12,1,2)='S" & txtZone & "'"
'      stConCP = stConCP & " And SubStr(CP12,1,2)='S" & txtZone & "'"
'   End If
'
'   '點數結算日
'   If txtCloseDate(0) <> "" Then
'      stCon = stCon & " And A0205 >= " & txtCloseDate(0) & "01"
'      '期初實績保留及期初結餘保留 抓畫面條件起日當月
'      stConR1 = "  And A0205 >= " & txtCloseDate(0) & "01 And A0205 <= " & txtCloseDate(0) & "31"
'      stConPE = stConPE & " And PE03(+) >= 191100+" & txtCloseDate(0)
'      stConCu = stConCu & " And CU14>=" & TransDate(txtCloseDate(0) & "01", 2)
'      stConCP = stConCP & " And CP05>=" & TransDate(txtCloseDate(0) & "01", 2)
'   End If
'   If txtCloseDate(1) <> "" Then
'      stCon = stCon & " And A0205 <= " & txtCloseDate(1) & "31"
'      '期末實績保留及期末結餘保留 抓畫面條件止日當月
'      stConR2 = "  And A0205 >= " & txtCloseDate(1) & "01 And A0205 <= " & txtCloseDate(1) & "31"
'      stConPE = stConPE & " And PE03(+) <= 191100+" & txtCloseDate(1)
'      stConCu = stConCu & " And CU14<=" & TransDate(txtCloseDate(1) & "31", 2)
'      stConCP = stConCP & " And CP05<=" & TransDate(txtCloseDate(1) & "31", 2)
'   End If
'
'On Error GoTo ErrHnd
'   '目標
'   stVTB(0) = "Select st01 ID,Sum(PE04) PE04" & _
'      " From Staff,PerFormance" & _
'      " Where  SubStr(st15,1,1)='S' And PE01(+)=ST01 And PE02(+)='TOT'" & stConST & stConPE & _
'      " Group by st01"
'   'Modify by Amy 2016/01/30 +InStr(ax212,'轉撥')=0
'   '期初實績保留:點數結算「起始」當月4191+4192貸方(期初實績保留)  'modify by sonia 2015/10/13 取消ST01>'60'條件,D104093822做20011
'   stVTB(1) = "Select ax209 V10, Sum(ax207) V11" & _
'      " From acc020, acc021,Staff" & _
'      " Where ax201(+) = a0201  And ax202(+) = a0202" & stConR1 & _
'      " And st01(+)=ax209 And ST01<'F' " & stConST & _
'      " And (ax205= '4191' Or ax205='4192') And InStr(ax212,'轉撥')=0 Group by ax209"
'   '期初結餘保留:點數結算「起始」當月4194貸方(期初結餘保留)  'modify by sonia 2015/10/13 取消ST01>'60'條件,D104093822做20011
'   stVTB(2) = "Select ax209 V20, Sum(ax207) V21" & _
'      " From acc020, acc021,Staff" & _
'      " Where ax201(+) = a0201  And ax202(+) = a0202" & stConR1 & _
'      " And st01(+)=ax209 And ST01<'F' " & stConST & _
'      " And ax205= '4194' And InStr(ax212,'轉撥')=0 Group by ax209"
'
'  '當月　　實績  'modify by sonia 2015/10/13 取消ST01>'60'條件,D104093822做20011
'  stVTB(3) = "Select ax209 V30, Sum(ax207-ax206) V31" & _
'      " From acc020, acc021,Staff" & _
'      " Where ax201(+) = a0201  And ax202(+) = a0202" & stCon & _
'      " And ST01(+)=ax209 And ST01<'F' " & stConST & _
'      " And (SubStr(ax205, 1, 2) = '41' Or (ax205='7121' And ax209 is not null)) And Not( ax205='4191' or ax205='4192' or ax205='4194')" & _
'      " And (ax213 Is Null or InStr(ax213||' ','結餘')=0) And InStr(ax212,'轉撥')=0 Group by ax209"
'  '當月　　結餘  'modify by sonia 2015/10/13 取消ST01>'60'條件,D104093822做20011
'   stVTB(4) = "Select ax209 V40, Sum(ax207-ax206) V41" & _
'      " From acc020, acc021,staff" & _
'      " Where ax201(+) = a0201  And ax202(+) = a0202" & stCon & _
'      " And ST01(+)=ax209 And ST01<'F' " & stConST & _
'      " And (SubStr(ax205, 1, 2) = '41' Or (ax205='7121' And ax209 is not null)) And Not( ax205='4191' or ax205='4192' or ax205='4194')" & _
'      " And InStr(ax213||' ','結餘')>0 And InStr(ax212,'轉撥')=0 Group by ax209"
'
'  '期末實績保留:點數結算「迄月」當月4194+4192借方(期末結餘保留)  'modify by sonia 2015/10/13 取消ST01>'60'條件,D104093822做20011
'   stVTB(5) = "Select ax209 V50, Sum(ax206) V51" & _
'      " From acc020, acc021,Staff" & _
'      " Where ax201(+) = a0201  And ax202(+) = a0202" & stConR2 & _
'      " And ST01(+)=ax209 And ST01<'F' " & stConST & _
'      " And ( ax205='4191' or ax205='4192') And InStr(ax212,'轉撥')=0 Group by ax209"
'   '期末結餘保留:點數結算「迄月」當月4194借方(期末結餘保留)  'modify by sonia 2015/10/13 取消ST01>'60'條件,D104093822做20011
'   stVTB(6) = "Select ax209 V60, Sum(ax206) V61" & _
'      " From acc020, acc021,Staff" & _
'      " Where ax201(+) = a0201  And ax202(+) = a0202" & stConR2 & _
'      " And ST01(+)=ax209 And ST01<'F' " & stConST & _
'      " And ax205='4194' And InStr(ax212,'轉撥')=0 Group by ax209"
'
'   '每月新增客戶數
'   stVTB(7) = "Select cu13 V70,count(*) V71" & _
'      " From Customer Where cu02='0'" & stConCu & " Group by cu13"
'   '收款家數  'modify by sonia 2015/10/13 取消ST01>'60'條件,D104093822做20011
'   'modify by sonia 2016/1/22 +412101,413101
'   stVTB(8) = "Select ax209 V80, Count(Distinct SubStr(ax208,1,6) ) V81" & _
'      " From acc020, acc021,Staff" & _
'      " Where ax201(+) = a0201  And ax202(+) = a0202" & stCon & _
'      " And ST01(+)=ax209 And ST01<'F' " & stConST & _
'      " And SubStr(ax205, 1, 2) = '41' And ax207>0 And ax208 is not null" & _
'      " And not( ax205='4191' or ax205='4192' or ((ax205='410103' or ax205='411103' or ax205='4121' or ax205='412101' or ax205='4131' or ax205='413101')" & _
'      " And InStr(ax213||' ','結餘')>0)) And InStr(ax212,'轉撥')=0 Group by ax209"
'
'   '收文案件來源分析
'   stVTB(9) = "Select cp13 V90" & _
'      ", Sum(Decode(CP01,'P',1,'PS',1,'CFP',1,'CPS',1,0)) V91" & _
'      ", Sum(Decode(CP01,'T',1,'TF',1,'CFT',1,'TC',1,0)) V92" & _
'      ", Sum(Decode(CP01,'L',1,'LA',1,0)) V93" & _
'      ", Sum(Decode(CP01,'CFC',1,'CFL',1,0)) V94" & _
'      " From caseprogress Where cp09<'B' And cp13 is not null" & stConCP & _
'      " And (  (cp01 in ('P','PS','CFP','CPS') )  or (cp01 in ('T','TF','CFT','TC') )" & _
'      " or (cp01 in ('CFC','CFL')) or (cp01 in ('L','LA'))) Group by cp13"
'
'   '人員數計算排除 S開頭的員工編號
'   strSql = "Select ST15,min(a0902) NAME,Sum(PE04) PE04,(NVL(Sum(V11),0))/1000 C1,NVL(Sum(V21),0)/1000 C2,NVL(Sum(V31),0)/1000 C3,NVL(Sum(V41),0)/1000 C4,NVL(Sum(V51),0)/1000 C5,NVL(Sum(V61),0)/1000 C6" & _
'      ", Sum(V71) C7, Sum(V81) C8, NVL(Sum(V91),0) C9, NVL(Sum(V92),0) C10, NVL(Sum(V93),0) C11, NVL(Sum(V94),0) C12, Count(Distinct Decode(SubStr(st01,1,1),'S',null,st01)) C13" & _
'      " From (" & stVTB(0) & ") V0,(" & stVTB(1) & ") V1,(" & stVTB(2) & ") V2,(" & stVTB(3) & ") V3" & _
'      ",(" & stVTB(4) & ") V4,(" & stVTB(5) & ") V5,(" & stVTB(6) & ") V6,(" & stVTB(7) & ") V7" & _
'      ",(" & stVTB(8) & ") V8,(" & stVTB(9) & ") V9,staff,acc090" & _
'      " Where V10(+)=ID And V20(+)=ID And V30(+)=ID And V40(+)=ID And V50(+)=ID And V60(+)=ID" & _
'      " And V70(+)=ID And V80(+)=ID And V90(+)=ID And st01(+)=ID And a0901(+)=st15 " & _
'      " And (PE04>0 or V11>0 or V21>0 or V31>0 or V41>0) Group by st15 Order by st15"
'    'Add by Amy 2016/01/30 抓取SalesPoint資料
'     strSqlSP = GetPoint_SP(Val(txtCloseDate(0)), Val(txtCloseDate(1)), "S" & txtZone & "0", "S" & txtZone & "9", , , Me.Name)
'
'    intI = 1: intQ = 1
'    Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strSql)
'    Set adoRecordset1 = ClsLawReadRstMsg(intQ, strSqlSP)
'    If intI = 1 Then
'       With AdoRecordSet3
'        iCol = 3
'       Erase dblTot
'       Do While Not .EOF
'          dblSumR = 0: dblSP15 = 0: dblSP36 = 0
'          grdDataList.Cols = iCol + 1
'          '區別
'          iRow = GetValue("區　　別")
'          grdDataList.TextMatrix(iRow, iCol) = "" & .Fields("NAME")
'          grdDataList.CellAlignment = flexAlignCenterCenter
'          '目標
'          iRow = GetValue("目　　標")
'          grdDataList.TextMatrix(iRow, iCol) = "" & .Fields("PE04")
'          grdDataList.CellAlignment = flexAlignRightCenter
'          '人員數
'          iRow = GetValue("人 員 數")
'          grdDataList.TextMatrix(iRow, iCol) = "" & .Fields("C13")
'          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'          '期初實績保留
'          iRow = GetValue("期初實績保留")
'          grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C1")), "0.000")
'          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'          dblSumR = dblSumR + Val(grdDataList.TextMatrix(iRow, iCol))
'          '期初結餘保留
'          iRow = GetValue("期初結餘保留")
'          grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C2")), "0.000")
'          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'          dblSumR = dblSumR + Val(grdDataList.TextMatrix(iRow, iCol))
'          '當月　　實績
'          iRow = GetValue("當月　　實績")
'           grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C3")), "0.000")
'          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'          dblSumR = dblSumR + Val(grdDataList.TextMatrix(iRow, iCol))
'          '當月　　結餘
'          iRow = GetValue("當月　　結餘")
'          grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C4")), "0.000")
'          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'          dblSumR = dblSumR + Val(grdDataList.TextMatrix(iRow, iCol))
'         'Add by Amy 2016/01/30業績輸入上線後改抓SalesPoint資料
'         If Val(txtCloseDate(0)) >= Val(業績輸入啟用年月) Then
'            If intQ = 1 Then
'                If "" & .Fields("ST15") = "" & adoRecordset1.Fields("SP48") Then
'                    dblSP15 = Val("" & adoRecordset1.Fields("SP15"))
'                    dblSP36 = Val("" & adoRecordset1.Fields("SP36"))
'                    adoRecordset1.MoveNext
'                End If
'            Else
'                dblSP15 = Val("" & .Fields("C5"))
'                dblSP36 = Val("" & .Fields("C6"))
'            End If
'         End If
'         'end 2016/01/30
'         '期末實績保留
'          iRow = GetValue("期末實績保留")
'          grdDataList.TextMatrix(iRow, iCol) = Format(dblSP15, "0.000")
'          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'          dblSumR = dblSumR - Val(grdDataList.TextMatrix(iRow, iCol))
'          '期末結餘保留
'          iRow = GetValue("期末結餘保留")
'          grdDataList.TextMatrix(iRow, iCol) = Format(dblSP36, "0.000")
'          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'          dblSumR = dblSumR - Val(grdDataList.TextMatrix(iRow, iCol))
'          '新增客戶數
'          iRow = GetValue("新增客戶數")
'          grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C7")))
'          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'          '收款家數
'          iRow = GetValue("收款家數")
'          grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C8")))
'          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'          '收文案件來源分析
'          'P
'           iRow = GetValue("收文案源分析")
'          grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C9")))
'          dblTotCase(0) = dblTotCase(0) + Val(grdDataList.TextMatrix(iRow, iCol))
'          'T
'          iRow = GetValue("收文案源分析") + 1
'          grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C10")))
'          dblTotCase(1) = dblTotCase(1) + Val(grdDataList.TextMatrix(iRow, iCol))
'          'L
'          iRow = GetValue("收文案源分析") + 2
'          grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C11")))
'          dblTotCase(2) = dblTotCase(2) + Val(grdDataList.TextMatrix(iRow, iCol))
'          'C
'          iRow = GetValue("收文案源分析") + 3
'          grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C12")))
'          dblTotCase(3) = dblTotCase(3) + Val(grdDataList.TextMatrix(iRow, iCol))
'
'          '達成點數
'          iRow = GetValue("達成點數")
'          strExc(0) = Val(grdDataList.TextMatrix(GetValue("期初實績保留"), iCol)) + Val(grdDataList.TextMatrix(GetValue("期初結餘保留"), iCol))
'          strExc(0) = Val(strExc(0)) + Val(grdDataList.TextMatrix(GetValue("當月　　實績"), iCol)) + Val(grdDataList.TextMatrix(GetValue("當月　　結餘"), iCol))
'          strExc(0) = Val(strExc(0)) - Val(grdDataList.TextMatrix(GetValue("期末實績保留"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末結餘保留"), iCol))
'          grdDataList.TextMatrix(iRow, iCol) = Format(strExc(0), "0.000")
'          dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'
'          '達成率
'          iRow = GetValue("達 成 率")
'          If Val(grdDataList.TextMatrix(GetValue("目　　標"), iCol)) > 0 Then
'             grdDataList.TextMatrix(iRow, iCol) = Format(100 * Val(grdDataList.TextMatrix(GetValue("達成點數"), iCol)) / Val(grdDataList.TextMatrix(GetValue("目　　標"), iCol)), "0.0") & "%"
'          Else
'             grdDataList.TextMatrix(iRow, iCol) = "0.0%"
'          End If
'          '平均達成
'          iRow = GetValue("平均達成")
'          If Val(grdDataList.TextMatrix(GetValue("人 員 數"), iCol)) > 0 Then
'            grdDataList.TextMatrix(iRow, iCol) = Format(Val(grdDataList.TextMatrix(GetValue("當月　　實績"), iCol)) / Val(grdDataList.TextMatrix(GetValue("人 員 數"), iCol)), "0.00")
'          Else
'            grdDataList.TextMatrix(iRow, iCol) = "0.00%"
'         End If
'         '實績保留動用
'         iRow = GetValue("實績保留動用")
'         'Modify by Amy 2015/11/09 +if 小於等於0顯示0(公式)
'         'Mark by Amy 2016/01/30 取消小於等於0顯示0(公式),因業績點數上線需與其資料一致
''         If Val(grdDataList.TextMatrix(GetValue("期初實績保留"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末實績保留"), iCol)) > 0 Then
'            grdDataList.TextMatrix(iRow, iCol) = Format(Val(grdDataList.TextMatrix(GetValue("期初實績保留"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末實績保留"), iCol)), "0.000")
'            dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
''         Else
''            grdDataList.TextMatrix(iRow, iCol) = Format("0", "0.000")
''         End If
'
'         '結餘保留動用
'         iRow = GetValue("結餘保留動用")
'         'Modify by Amy 2015/11/09 +if 小於等於0顯示0(公式)
'         'Mark by Amy 2016/01/30 取消小於等於0顯示0(公式),因業績點數上線需與其資料一致
''         If Val(grdDataList.TextMatrix(GetValue("期初結餘保留"), iCol)) + Val(grdDataList.TextMatrix(GetValue("當月　　結餘"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末結餘保留"), iCol)) > 0 Then
'            grdDataList.TextMatrix(iRow, iCol) = Format(Val(grdDataList.TextMatrix(GetValue("期初結餘保留"), iCol)) + Val(grdDataList.TextMatrix(GetValue("當月　　結餘"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末結餘保留"), iCol)), "0.000")
'            dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
''         Else
''            grdDataList.TextMatrix(iRow, iCol) = Format("0", "0.000")
''         End If
'
'         .MoveNext
'         iCol = iCol + 1
'      Loop
'      End With
'
'      grdDataList.Cols = iCol + 1
'      grdDataList.TextMatrix(0, iCol) = "合　　計"
'      '總目標需抓該區所有(因為有可能點數掛該區代碼如20021)
'      strExc(0) = "Select nvl(sum(PE04),0) PE04 From Staff,PerFormance " & _
'         "Where PE01(+)=ST01 And PE02(+)='TOT'" & stConST & stConPE
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         dblTot(1) = RsTemp.Fields(0)
'      End If
'      For i = 1 To UBound(strRowN)
'         Select Case i
'            Case GetValue("目　　標")
'                grdDataList.TextMatrix(i, iCol) = dblTot(i)
'            Case GetValue("達成點數")
'                grdDataList.TextMatrix(i, iCol) = Format(dblTot(i), "0.000")
'            Case GetValue("達 成 率")
'                If Val(grdDataList.TextMatrix(GetValue("目　　標"), iCol)) > 0 Then
'                    grdDataList.TextMatrix(i, iCol) = Format(100 * Val(grdDataList.TextMatrix(GetValue("達成點數"), iCol)) / Val(grdDataList.TextMatrix(GetValue("目　　標"), iCol)), "0.0") & "%"
'                Else
'                    grdDataList.TextMatrix(i, iCol) = "0.0%"
'                End If
'            Case GetValue("人 員 數")
'                grdDataList.TextMatrix(i, iCol) = Format(dblTot(i))
'            Case GetValue("平均達成")
'                If Val(grdDataList.TextMatrix(GetValue("人 員 數"), iCol)) > 0 Then
'                    grdDataList.TextMatrix(i, iCol) = Format(Val(grdDataList.TextMatrix(GetValue("當月　　實績"), iCol)) / Val(grdDataList.TextMatrix(GetValue("人 員 數"), iCol)), "0.00") & "%"
'                Else
'                    grdDataList.TextMatrix(i, iCol) = "0.00%"
'                End If
'            'Modify by Amy 2015/04/30
'            Case GetValue("期末實績保留") To GetValue("期末結餘保留")
'                If Val(dblTot(GetValue("期末實績保留"))) = 0 And Val(dblTot(GetValue("期末結餘保留"))) = 0 Then
'                    '財務尚未做傳票,則實績保留動用及餘額需設 0 讓智權人員填
'                    If i = GetValue("期末實績保留") Then
'                        '實績保留動用及餘額需設 0 (設一次就好)
'                        For j = 3 To grdDataList.Cols - 1
'                            grdDataList.TextMatrix(GetValue("實績保留動用"), j) = Format("0", "0.000")
'                            grdDataList.TextMatrix(GetValue("結餘保留動用"), j) = Format("0", "0.000")
'                        Next j
'                    End If
'                End If
'                grdDataList.TextMatrix(i, iCol) = Format(dblTot(i), "0.000")
'            Case GetValue("新增客戶數") To GetValue("收款家數")
'                grdDataList.TextMatrix(i, iCol) = Format(dblTot(i))
'            Case GetValue("收文案源分析")
'                grdDataList.TextMatrix(i, iCol) = Format(dblTotCase(0))
'                grdDataList.TextMatrix(i + 1, iCol) = Format(dblTotCase(1))
'                grdDataList.TextMatrix(i + 2, iCol) = Format(dblTotCase(2))
'                grdDataList.TextMatrix(i + 3, iCol) = Format(dblTotCase(3))
'            Case Else
'                grdDataList.TextMatrix(i, iCol) = Format(dblTot(i), "0.000")
'         End Select
'      Next i
'
'   Else
'      MsgBox "無符合資料！", vbInformation
'   End If
'
'
'   doQuery = True
'
'ErrHnd:
'   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

Private Function doQuery_Old2() As Boolean

'   Dim stCon As String, stConST As String, stConResv As String, stConPE As String
'   Dim stVTB0 As String, stVTB1 As String, stVTB2 As String, stVTB3 As String, stVTB4 As String
'   Dim iRow As Single, iCol As Single
'   Dim dblTot(1 To 14) As Double
'   Dim stVTB5 As String, stVTB6 As String, stVTB7 As String, stVTB8 As String
'   Dim stConCu As String, stConCP As String
'
'   stCon = "": stConST = "": stConResv = "": stConPE = "": stConCu = ""
'
'   '所別
'   If txtZone <> "" Then
'      stConST = " AND ST06='" & txtZone & "'"
'      stConCu = stConCu & " AND SUBSTR(CU12,1,2)='S" & txtZone & "'"
'      stConCP = stConCP & " AND SUBSTR(CP12,1,2)='S" & txtZone & "'"
'   End If
'
'   '點數結算日
'   If txtCloseDate(0) <> "" Then
'      stCon = stCon & " AND A0205 >= " & txtCloseDate(0) & "01"
'      stConPE = stConPE & " AND PE03(+) >= 191100+" & txtCloseDate(0)
'      stConCu = stConCu & " AND CU14>=" & TransDate(txtCloseDate(0) & "01", 2)
'      stConCP = stConCP & " AND CP05>=" & TransDate(txtCloseDate(0) & "01", 2)
'   End If
'   If txtCloseDate(1) <> "" Then
'      stCon = stCon & " AND A0205 <= " & txtCloseDate(1) & "31"
'      '保留只抓迄月
'      stConResv = "  AND A0205 >= " & txtCloseDate(1) & "01 AND A0205 <= " & txtCloseDate(1) & "31"
'      stConPE = stConPE & " AND PE03(+) <= 191100+" & txtCloseDate(1)
'      stConCu = stConCu & " AND CU14<=" & TransDate(txtCloseDate(1) & "31", 2)
'      stConCP = stConCP & " AND CP05<=" & TransDate(txtCloseDate(1) & "31", 2)
'   End If
'
'On Error GoTo ErrHnd
'   '目標
'   stVTB0 = "select st01 ID,sum(PE04) PE04" & _
'      " from staff,PERFORMANCE" & _
'      " where  SUBSTR(st15,1,1)='S' AND PE01(+)=ST01 AND PE02(+)='TOT'" & stConST & stConPE & _
'      " GROUP BY st01"
'
'   '達成點數
'   '2014/1/21 modify by sonia 取消a0201='1'條件
'   'stVTB1 = "select ax209 V10, sum(ax207-ax206) V11" & _
'      " From acc020, acc021,staff" & _
'      " where a0201='1'" & stCon & _
'      " and ax201(+) = a0201  and ax202(+) = a0202" & _
'      " and st01(+)=ax209 and ST01>'60' AND ST01<'F' AND SUBSTR(ST15,1,1)='S'" & stConST & _
'      " and (substr(ax205, 1, 2) = '41' Or substr(ax205,1,2)='71')" & _
'      " group by ax209"
'   stVTB1 = "select ax209 V10, sum(ax207-ax206) V11" & _
'      " From acc020, acc021,staff" & _
'      " where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
'      " and st01(+)=ax209 and ST01>'60' AND ST01<'F' AND SUBSTR(ST15,1,1)='S'" & stConST & _
'      " and (substr(ax205, 1, 2) = '41' Or substr(ax205,1,2)='71')" & _
'      " group by ax209"
'   '2014/1/21 end
'
'   '當月達成
'   '2014/1/21 modify by sonia 取消a0201='1'條件
'   'stVTB2 = "select ax209 V20, sum(ax207-ax206) V21" & _
'      " From acc020, acc021,staff" & _
'      " where a0201='1'" & stCon & _
'      " and ax201(+) = a0201  and ax202(+) = a0202" & _
'      " and ST01(+)=ax209 and ST01>'60' AND ST01<'F' AND SUBSTR(st15,1,1)='S'" & stConST & _
'      " and (substr(ax205, 1, 2) = '41' Or substr(ax205,1,2)='71')" & _
'      " and not( ax205='4191' or ax205='4192' or ((ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131')" & _
'      " and instr(ax213||' ','結餘')>0)) group by ax209"
'   stVTB2 = "select ax209 V20, sum(ax207-ax206) V21" & _
'      " From acc020, acc021,staff" & _
'      " where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
'      " and ST01(+)=ax209 and ST01>'60' AND ST01<'F' AND SUBSTR(st15,1,1)='S'" & stConST & _
'      " and (substr(ax205, 1, 2) = '41' Or substr(ax205,1,2)='71')" & _
'      " and not( ax205='4191' or ax205='4192' or ((ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131')" & _
'      " and instr(ax213||' ','結餘')>0)) group by ax209"
'   '2014/1/21 end
'
'   '結餘
'   '2014/1/21 modify by sonia 取消a0201='1'條件
'   'stVTB3 = "select ax209 V30, sum(ax207-ax206) V31" & _
'      " From acc020, acc021,staff" & _
'      " where a0201='1'" & stCon & _
'      " and ax201(+) = a0201  and ax202(+) = a0202" & _
'      " and ST01(+)=ax209 and ST01>'60' AND ST01<'F' AND SUBSTR(st15,1,1)='S'" & stConST & _
'      " and (ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131')" & _
'      " and instr(ax213||' ','結餘')>0 group by ax209"
'   stVTB3 = "select ax209 V30, sum(ax207-ax206) V31" & _
'      " From acc020, acc021,staff" & _
'      " where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
'      " and ST01(+)=ax209 and ST01>'60' AND ST01<'F' AND SUBSTR(st15,1,1)='S'" & stConST & _
'      " and (ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131')" & _
'      " and instr(ax213||' ','結餘')>0 group by ax209"
'   '2014/1/21 end
'
'   '保留
'   '2014/1/21 modify by sonia 取消a0201='1'條件
'   'stVTB4 = "select ax209 V40, sum(ax206) V41" & _
'      " From acc020, acc021,staff" & _
'      " where a0201='1'" & stConResv & _
'      " and ax201(+) = a0201  and ax202(+) = a0202" & _
'      " and ST01(+)=ax209 and ST01>'60' AND ST01<'F' AND SUBSTR(st15,1,1)='S'" & stConST & _
'      " and (ax205='4191' or ax205='4192') group by ax209"
'   stVTB4 = "select ax209 V40, sum(ax206) V41" & _
'      " From acc020, acc021,staff" & _
'      " where ax201(+) = a0201  and ax202(+) = a0202" & stConResv & _
'      " and ST01(+)=ax209 and ST01>'60' AND ST01<'F' AND SUBSTR(st15,1,1)='S'" & stConST & _
'      " and (ax205='4191' or ax205='4192') group by ax209"
'   '2014/1/21 end
'
'   '每月新增客戶數
'   stVTB5 = "select cu13 V50,count(*) V51" & _
'      " from customer where cu02='0'" & stConCu & " group by cu13"
'   '收款家數
'   '2014/1/21 modify by sonia 取消a0201='1'條件
'   'stVTB6 = "select ax209 V60, count( distinct substr(ax208,1,6) ) V61" & _
'      " From acc020, acc021,staff" & _
'      " where a0201='1'" & stCon & _
'      " and ax201(+) = a0201  and ax202(+) = a0202" & _
'      " and ST01(+)=ax209 and ST01>'60' AND ST01<'F' AND SUBSTR(st15,1,1)='S'" & stConST & _
'      " and substr(ax205, 1, 2) = '41' and ax207>0 and ax208 is not null" & _
'      " and not( ax205='4191' or ax205='4192' or ((ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131')" & _
'      " and instr(ax213||' ','結餘')>0)) group by ax209"
'   stVTB6 = "select ax209 V60, count( distinct substr(ax208,1,6) ) V61" & _
'      " From acc020, acc021,staff" & _
'      " where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
'      " and ST01(+)=ax209 and ST01>'60' AND ST01<'F' AND SUBSTR(st15,1,1)='S'" & stConST & _
'      " and substr(ax205, 1, 2) = '41' and ax207>0 and ax208 is not null" & _
'      " and not( ax205='4191' or ax205='4192' or ((ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131')" & _
'      " and instr(ax213||' ','結餘')>0)) group by ax209"
'   '2014/1/21 end
'
'   '收文案件來源分析
'   stVTB7 = "SELECT cp13 V70" & _
'      ", SUM(DECODE(CP01,'P',1,'PS',1,'CFP',1,'CPS',1,0)) V71" & _
'      ", SUM(DECODE(CP01,'T',1,'TF',1,'CFT',1,'TC',1,0)) V72" & _
'      ", SUM(DECODE(CP01,'L',1,'LA',1,0)) V73" & _
'      ", SUM(DECODE(CP01,'CFC',1,'CFL',1,0)) V74" & _
'      " From caseprogress where cp09<'B'" & stConCP & _
'      " and (  (cp01 in ('P','PS','CFP','CPS') )  or (cp01 in ('T','TF','CFT','TC') )" & _
'      " or (cp01 in ('CFC','CFL')) or (cp01 in ('L','LA'))) GROUP BY cp13"
'
'    '各區業務數
'   stVTB8 = "select st15,count(*) V81" & _
'      " from staff where " & stConCu & " group by st15"
'   'Modify by Morgan 2010/3/25 人員數計算排除 S開頭的編號
'   strSql = "select ST15,min(a0902) NAME,sum(PE04) PE04,(NVL(SUM(V11),0))/1000 C1,NVL(SUM(V21),0)/1000 C2,NVL(SUM(V31),0)/1000 C3,NVL(SUM(V41),0)/1000 C4" & _
'      ", SUM(V51) C5, SUM(V61) C6, NVL(SUM(V71),0) C7, NVL(SUM(V72),0) C8, NVL(SUM(V73),0) C9, NVL(SUM(V74),0) C10, count(distinct decode(substr(st01,1,1),'S',null,st01)) C11" & _
'      " from (" & stVTB0 & ") V0,(" & stVTB1 & ") V1,(" & stVTB2 & ") V2,(" & stVTB3 & ") V3" & _
'      ",(" & stVTB4 & ") V4,(" & stVTB5 & ") V5,(" & stVTB6 & ") V6,(" & stVTB7 & ") V7,staff,acc090" & _
'      " where V10(+)=ID AND V20(+)=ID AND V30(+)=ID AND V40(+)=ID AND V50(+)=ID AND V60(+)=ID" & _
'      " AND V70(+)=ID AND st01(+)=ID and a0901(+)=st15 AND (PE04>0 OR V11>0) group by st15"
'
'   intI = 1
'   Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      With AdoRecordSet3
'      iRow = 0: iCol = 3
'      Erase dblTot
'      Do While Not .EOF
'         grdDataList.Cols = iCol + 1
'         '區別
'         grdDataList.TextMatrix(0, iCol) = "" & .Fields("NAME")
'         grdDataList.row = 0: grdDataList.col = iCol
'         grdDataList.CellAlignment = flexAlignCenterCenter
'         '目標
'         grdDataList.TextMatrix(1, iCol) = "" & .Fields("PE04")
'         dblTot(1) = dblTot(1) + Val(grdDataList.TextMatrix(1, iCol))
'         '達成點數
'         grdDataList.TextMatrix(2, iCol) = Format(Val("" & .Fields("C1")), "0.000")
'         dblTot(2) = dblTot(2) + Val(grdDataList.TextMatrix(2, iCol))
'         '達成率
'         If Val(grdDataList.TextMatrix(1, iCol)) > 0 Then
'            grdDataList.TextMatrix(3, iCol) = Format(100 * Val(grdDataList.TextMatrix(2, iCol)) / Val(grdDataList.TextMatrix(1, iCol)), "0.0") & "%"
'         End If
'         'Modify by Morgan 2010/3/25 改用語法控制
'         '人員數
'         'If .Fields("ST15") = "S29" Then
'         '   grdDataList.TextMatrix(4, iCol) = "0"
'         'Else
'            grdDataList.TextMatrix(4, iCol) = "" & .Fields("C11")
'         'End If
'         dblTot(4) = dblTot(4) + Val(grdDataList.TextMatrix(4, iCol))
'         '當月達成
'         grdDataList.TextMatrix(5, iCol) = Format(Val("" & .Fields("C2")), "0.0")
'         dblTot(5) = dblTot(5) + Val(grdDataList.TextMatrix(5, iCol))
'         '平均達成
'         If Val(grdDataList.TextMatrix(4, iCol)) = 0 Then
'            grdDataList.TextMatrix(6, iCol) = "0"
'         Else
'            grdDataList.TextMatrix(6, iCol) = Format(Val(grdDataList.TextMatrix(5, iCol)) / Val(grdDataList.TextMatrix(4, iCol)), "0.00")
'         End If
'         '結餘點數
'         grdDataList.TextMatrix(7, iCol) = Format(Val("" & .Fields("C3")), "0.00")
'         dblTot(7) = dblTot(7) + Val(grdDataList.TextMatrix(7, iCol))
'         '保留點數
'         grdDataList.TextMatrix(8, iCol) = Format(Val("" & .Fields("C4")))
'         dblTot(8) = dblTot(8) + Val(grdDataList.TextMatrix(8, iCol))
'         '新增客戶數
'         grdDataList.TextMatrix(9, iCol) = Format(Val("" & .Fields("C5")))
'         dblTot(9) = dblTot(9) + Val(grdDataList.TextMatrix(9, iCol))
'         '收款家數
'         grdDataList.TextMatrix(10, iCol) = Format(Val("" & .Fields("C6")))
'         dblTot(10) = dblTot(10) + Val(grdDataList.TextMatrix(10, iCol))
'         '收文案件來源分析
'         'P
'         grdDataList.TextMatrix(11, iCol) = Format(Val("" & .Fields("C7")))
'         dblTot(11) = dblTot(11) + Val(grdDataList.TextMatrix(11, iCol))
'         'T
'         grdDataList.TextMatrix(12, iCol) = Format(Val("" & .Fields("C8")))
'         dblTot(12) = dblTot(12) + Val(grdDataList.TextMatrix(12, iCol))
'         'L
'         grdDataList.TextMatrix(13, iCol) = Format(Val("" & .Fields("C9")))
'         dblTot(13) = dblTot(13) + Val(grdDataList.TextMatrix(13, iCol))
'         'C
'         grdDataList.TextMatrix(14, iCol) = Format(Val("" & .Fields("C10")))
'         dblTot(14) = dblTot(14) + Val(grdDataList.TextMatrix(14, iCol))
'
'         .MoveNext
'         iCol = iCol + 1
'      Loop
'      End With
'
'      grdDataList.Cols = iCol + 1
'      grdDataList.TextMatrix(0, iCol) = "合　　計"
'      grdDataList.TextMatrix(1, iCol) = dblTot(1)
'      grdDataList.TextMatrix(2, iCol) = Format(dblTot(2), "0.000")
'      '達成率
'      If Val(grdDataList.TextMatrix(1, iCol)) > 0 Then
'         grdDataList.TextMatrix(3, iCol) = Format(100 * Val(grdDataList.TextMatrix(2, iCol)) / Val(grdDataList.TextMatrix(1, iCol)), "0.0") & "%"
'      End If
'      grdDataList.TextMatrix(4, iCol) = dblTot(4)
'      grdDataList.TextMatrix(5, iCol) = Format(dblTot(5), "0.00")
'      '平均達成
'      grdDataList.TextMatrix(6, iCol) = Format(Val(grdDataList.TextMatrix(5, iCol)) / Val(grdDataList.TextMatrix(4, iCol)), "0.00")
'      grdDataList.TextMatrix(7, iCol) = Format(dblTot(7))
'      grdDataList.TextMatrix(8, iCol) = Format(dblTot(8))
'      grdDataList.TextMatrix(9, iCol) = Format(dblTot(9))
'      grdDataList.TextMatrix(10, iCol) = Format(dblTot(10))
'      grdDataList.TextMatrix(11, iCol) = Format(dblTot(11))
'      grdDataList.TextMatrix(12, iCol) = Format(dblTot(12))
'      grdDataList.TextMatrix(13, iCol) = Format(dblTot(13))
'      grdDataList.TextMatrix(14, iCol) = Format(dblTot(14))
'   Else
'      MsgBox "無符合資料！", vbInformation
'   End If
'
'
'   doQuery = True
'
'ErrHnd:
'   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

Private Function GetValue(pRowN As String) As Integer
   Dim jj As Integer
 
    For jj = 1 To UBound(strRowN)
       If UCase(strRowN(jj)) = UCase(pRowN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function
'end 2015/04/22

Private Sub Form_Unload(Cancel As Integer)
   Set frm210117 = Nothing
End Sub

Private Sub txtCloseDate_GotFocus(Index As Integer)
   TextInverse txtCloseDate(Index)
   CloseIme
End Sub

Private Sub txtCloseDate_Validate(Index As Integer, Cancel As Boolean)
   
   If txtCloseDate(Index) <> "" Then
      If ChkDate(txtCloseDate(Index) & "01") = False Then
         txtCloseDate_GotFocus Index
         Cancel = True
      End If
       'Add by Amy 2016/01/30 不可查舊資料
      If Index = 0 And Val(txtCloseDate(0)) < Val(業績輸入啟用年月) Then
        MsgBox Label2 & "不可查詢105年1月前的資料"
        txtCloseDate_GotFocus Index
        Cancel = True
      End If
      'Add by Amy 2015/04/27
      If Index = 1 And Val(txtCloseDate(0)) > Val(txtCloseDate(1)) Then
         MsgBox Label2 & "起始年月不可大於截止年月"
         txtCloseDate_GotFocus Index
         Cancel = True
      End If
   End If
End Sub

Private Sub txtZone_GotFocus()
   TextInverse txtZone
   CloseIme
End Sub
