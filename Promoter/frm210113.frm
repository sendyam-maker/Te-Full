VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm210113 
   BorderStyle     =   1  '單線固定
   Caption         =   "各區業務工作報告統計"
   ClientHeight    =   5050
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   9490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5050
   ScaleWidth      =   9490
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
      TabIndex        =   9
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
      Top             =   420
      Width           =   915
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   285
      Index           =   1
      Left            =   2430
      MaxLength       =   5
      TabIndex        =   2
      Top             =   420
      Width           =   915
   End
   Begin VB.TextBox txtSalesArea 
      Height          =   285
      Left            =   1170
      MaxLength       =   3
      TabIndex        =   0
      Top             =   60
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4185
      Left            =   90
      TabIndex        =   8
      Top             =   780
      Width           =   9270
      _ExtentX        =   16334
      _ExtentY        =   7391
      _Version        =   393216
      BackColor       =   -2147483624
      Rows            =   27
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
         Size            =   9.5
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
   Begin VB.Label Label4 
      Caption         =   "實績結餘點數從105年1月啟用"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   4080
      TabIndex        =   11
      Top             =   480
      Width           =   2505
   End
   Begin VB.Label Label3 
      Caption         =   "(年月)"
      Height          =   180
      Left            =   3375
      TabIndex        =   10
      Top             =   480
      Width           =   585
   End
   Begin VB.Label Label2 
      Caption         =   "點數結算期間"
      Height          =   180
      Left            =   135
      TabIndex        =   7
      Top             =   465
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   2160
      X2              =   2395
      Y1              =   555
      Y2              =   555
   End
   Begin VB.Label Label1 
      Caption         =   "業務區"
      Height          =   180
      Left            =   180
      TabIndex        =   6
      Top             =   105
      Width           =   540
   End
   Begin VB.Label lblSalesArea 
      Height          =   180
      Left            =   2160
      TabIndex        =   5
      Top             =   105
      Width           =   1530
   End
End
Attribute VB_Name = "frm210113"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/14 改成Form2.0 ; grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit
'Add by Amy 2015/04/20
Dim strRowN()
Dim dblTot() As Double 'Modify by Amy 2016/03/01
Dim bolNoMemo As Boolean 'Add by Amy 2016/03/01
'Added by Lydia 2022/05/03 簡協理69005改為抓系統特殊設定「全所智權部主管」;林協理82026改為抓系統特殊設定「中所智權部主管」
Dim strSpecTmp1 As String, strSpecTmp2 As String
Dim rsA As New ADODB.Recordset 'Add by Amy 2024/04/19

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   'add by sonia 2016/12/21 柄佑82026可看中所全部但不能看P11
   'Modified by Lydia 2022/05/03
   'If strUserNum = "82026" Then
   If strUserNum = strSpecTmp2 Then
      If txtSalesArea <> "P11" Then
         If txtSalesArea < "S2" Or txtSalesArea > "S29" Then
            MsgBox "業務區條件錯誤！只可查中所業務區", vbExclamation
            txtSalesArea.SetFocus
            txtSalesArea_GotFocus
            Exit Sub
         End If
      Else
         MsgBox "您不可以查詢 P11 的資料", vbExclamation
         txtSalesArea.SetFocus
         txtSalesArea_GotFocus
         Exit Sub
      End If
   End If
   'end 2016/12/21
   
   Screen.MousePointer = vbHourglass
   If ConstrainCheck = True Then
      'Add byAmy 2016/03/01 起迄年月相同才顯示轉撥備註
      bolNoMemo = False
      If Val(Left(txtCloseDate(0), 5)) <> Val(Left(txtCloseDate(1), 5)) Then
         bolNoMemo = True
      End If
      Call SetDataListWidth
      Call doQuery
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Function ConstrainCheck() As Boolean
   Dim bolCancel As Boolean
   Dim strA0b01_1 As String, strA0b01_J As String, strA0b05 As String 'Add by Amy 2016/03/01
   Dim strMsg As String 'Add by Amy 2022/02/11
   
   If txtSalesArea = "" Then
      MsgBox "請輸入業務區！", vbExclamation
      txtSalesArea.SetFocus
      Exit Function
   End If
   
'cancel by sonia 2014/6/9
'   '蔣律師只可看中所全部  2006/2/7加71003
'   If strUserNum = "79037" Or strUserNum = "71003" Then
'      If Left(txtSalesArea, 2) <> "S2" Then
'         MsgBox "請輸入中所業務區！", vbExclamation
'         txtSalesArea.SetFocus
'         Exit Function
'      End If
'   End If
'end 2014/6/9
      
   '2006/11/29 ADD BY SONIA 簡金泉只可看北所全部
   'Removed by Morgan 2019/12/30 杜主秘說加開簡協理69005可看全所
   'If strUserNum = "69005" Then
   '   If Left(txtSalesArea, 2) <> "S1" Then
   '      MsgBox "請輸入北所業務區！", vbExclamation
   '      txtSalesArea.SetFocus
   '      Exit Function
   '   End If
   'End If
   'end 2019/12/30
   '2006/11/29 END
   
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
        If Right(strA0b01_1, 2) = "12" Then
            strMsg = Left(strA0b01_1, 3) + 1 & "年1月"
        Else
           strMsg = Left(strA0b01_1, 3) & "年" & Val(Right(strA0b01_1, 2)) + 1 & "月"
        End If
        MsgBox Label2 & "點數結算期" & strMsg & vbCrLf & _
                       "財務處尚未過帳，故不可操作"
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

'Mark by Amy 2015/04/20
'Private Sub runWord_Old()
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
'      .Selection.TypeText Text:=lblSalesArea & "智權人員工作報告"
'
'      .Selection.TypeParagraph
'      .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
'      .Selection.TypeText Text:="一、業績達成情形："
'
'      .Selection.TypeParagraph
'      If grdDataList.Cols - 2 > 7 Then
'         .Selection.Font.Size = 12
'      End If
'      '列數+1("評語"另外印),欄數-2(0,2不印)
'      Set oTable = .Selection.Tables.Add(Range:=.Selection.Range, NumRows:=grdDataList.Rows + 1, NumColumns:=grdDataList.Cols - 2)
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
'                  If iRow = 9 Then
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
'      '評　　語
'      .Selection.SelectRow
'      .Selection.Cells.SetHeight RowHeight:=80, HeightRule:=wdRowHeightExactly
'      .Selection.Cells.VerticalAlignment = wdCellAlignVerticalTop
'      .Selection.TypeText Text:="評　　語"
'      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
'      .Selection.MoveRight Unit:=wdCharacter, Count:=grdDataList.Cols - 1
'
'      .Selection.TypeParagraph
'      .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
'      .Selection.Font.Size = 14
'      .Selection.TypeText Text:="二、工作情形（當月發生之重要案件、市場反應…等處理）："
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeText Text:=String(20, "　") & "報告人：" & strUserName
'      .Selection.TypeParagraph
'      .Selection.TypeText Text:=String(20, "　") & "（各分區主管每月8日前填報）"
'      .Selection.WholeStory
'      .Selection.Font.Name = "Times New Roman"
'      '預設游標停留位置
'      .Selection.Find.ClearFormatting
'      If .Selection.Find.Execute("評　　語") = True Then
'         .Selection.MoveRight Unit:=wdCharacter, Count:=2 '右移一格
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
'end 2015/04/20

Private Sub cmdWord_Click()
   Screen.MousePointer = vbHourglass
   runWord
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

   Dim stST05 As String, stST15 As String
   
   MoveFormToCenter Me
   
   stST15 = PUB_GetStaffST15(strUserNum, 1)
   stST05 = PUB_GetST05(strUserNum)
   
   'Added by Lydia 2022/05/03 簡協理69005改為抓系統特殊設定「全所智權部主管」;林協理82026改為抓系統特殊設定「中所智權部主管」
   strSpecTmp1 = Pub_GetSpecMan("全所智權部主管")
   If strSpecTmp1 = "" Then strSpecTmp1 = "ABC"
   strSpecTmp2 = Pub_GetSpecMan("中所智權部主管")
   If strSpecTmp2 = "" Then strSpecTmp2 = "DEF"
   'end 2022/05/03
   
   Select Case strUserNum
      '小真,杜副總可看全部
      '2012/2/14 modify by sonia 加林特助94007權限
      'modify by sonia 2014/6/9 +美珍77027,並取消94007(因己改個人等級)
      Case "65001", "68006", "77027"
         txtSalesArea.Enabled = True
'cancel by sonia 2014/6/9
'      '蔣律師可看中所全部
'      Case "79037"
'         txtSalesArea.Enabled = True
'end 2014/6/9
      '2006/2/7 ADD BY SONIA 加71003可看中所全部
      Case "71003"
         txtSalesArea.Enabled = True
         txtSalesArea = stST15
      '2006/2/7 END
      '2006/11/29 ADD BY SONIA 加69005可看北所全部
      'Memo by Morgan 2019/12/30 杜主秘說加開簡協理69005可看全所(預設S15)
      'Modified by Lydia 2022/05/03
      'Case "69005"
      Case strSpecTmp1
         txtSalesArea.Enabled = True
         txtSalesArea = stST15
      '2006/2/7 END
      'add by sonia 2016/12/21 柄佑82026可看中所全部但不能看P11
      'Modified by Lydia 2022/05/03
      'Case "82026"
      Case strSpecTmp2
         txtSalesArea.Enabled = True
         txtSalesArea = "S21"
      'end 2016/12/21
      Case Else
         Select Case stST05
            '電腦中心,財務,總經理看全部
            '2015/7/28 MODIFY BY SONIA +主任秘書(等級08)可看全部
            Case "00", "01", "08"
               txtSalesArea.Enabled = True
            Case Else
               txtSalesArea.Enabled = False
               txtSalesArea = stST15
         End Select
   End Select
   
   txtCloseDate(0) = (CompDate(1, -1, strSrvDate(2)) - 19110000) \ 100
   txtCloseDate(1) = txtCloseDate(0)
   bolNoMemo = False
   Call SetDataListWidth
End Sub

'p_bolHeaderOnly:是否只設定表頭 true=是 false=資料一併清除
Private Sub SetDataListWidth(Optional ByVal p_bolHeaderOnly As Boolean = False)
   Dim iCol As Integer, iRow As Integer
   
   'Add by Amy 2015/04/20
   'Modify by Amy 2016/03/01 +轉撥及轉撥備註並調顯示位置與「智權點數實績與結餘輸入」全區資料相同
   'Memo by Lydia 2019/08/08 frm210152智權點數實績與結餘輸入在承辦人系統更名為「每月點數查詢／輸入」，與財務系統不同
'   ReDim strRowN(0 To 14)
'   strRowN = Array("智權人員", "目　　標", "達成點數", "達 成 率", "期初實績保留", _
'                            "期初結餘保留", "當月　　實績", "當月　　結餘", "實績保留動用", "結餘保留動用", _
'                            "期末實績保留", "期末結餘保留", "新增客戶數", "收款家數", "收文案源分析")
    'Modify by Amy 2024/04/19 增加及修改名稱 達成點數->報出點數 /達 成 率->報出點數達成率;grdDataList列屬性22->28
'    ReDim strRowN(0 To 20)
'    strRowN = Array("智權人員", "目　　標", "達成點數", "達 成 率", _
'                              "期初實績保留", "當月　　實績", "實績保留動用", "期末實績保留", "轉撥實績增減", "轉撥實績備註", "報出實績點數", _
'                              "期初結餘保留", "當月　　結餘", "結餘保留動用", "期末結餘保留", "轉撥結餘增減", "轉撥結餘備註", "報出結餘點數", _
'                              "新增客戶數", "收款家數", "收文案源分析")
   ReDim strRowN(0 To 24)
    strRowN = Array("智權人員", "目　　標", "報出點數", "報出點數達成率", _
                              "期初實績保留", "當月　　實績", "當月實績達成率", "實績保留動用", "期末實績保留", "轉撥實績增減", "轉撥實績備註", "報出實績點數", _
                              "期初結餘保留", "當月　　結餘", "結餘保留動用", "期末結餘保留", "轉撥結餘增減", "轉撥結餘備註", "報出結餘點數", _
                              "收文點數", "收文點數達成率", "出缺勤狀況", _
                              "新增客戶數", "收款家數", "收文案源分析")
                            
   With grdDataList
      .Visible = False
      .Clear
      .Rows = 0
      '.Font.Size = 12
      If p_bolHeaderOnly = False Then
         '.Rows = 14
         'Modify by Amy 2016/03/01 增加rows原18
         'Modify by Amy 2024/04/19 增加rows原24
         .Rows = UBound(strRowN) + 4: .Cols = 8: .FixedRows = 1: .FixedCols = 3
         .ColAlignmentFixed = flexAlignCenterCenter
      End If
      .MergeCells = flexMergeRestrictColumns
      .MergeCol(1) = True
      .ColWidth(0) = 0
      .ColWidth(1) = 1300 'Modify by Amy 2024/04/19 原:1000
      .ColWidth(2) = 300
'Modify by Amy 2015/04/20
'      .TextMatrix(0, 1) = "智權人員"
'      .TextMatrix(1, 1) = "目　　標"
'      .TextMatrix(2, 1) = "達成點數"
'      .TextMatrix(3, 1) = "達 成 率"
'      .TextMatrix(4, 1) = "當月達成"
'      .TextMatrix(5, 1) = "結餘點數"
'      .TextMatrix(6, 1) = "保留點數"
'      .TextMatrix(7, 1) = "新增客戶數"
'      .TextMatrix(8, 1) = "收款家數"
'      .TextMatrix(9, 1) = "收文" & vbCrLf & "案源分析"
'      For iRow = 10 To 12
'         .TextMatrix(iRow, 1) = .TextMatrix(9, 1)
'      Next
'      '.TextMatrix(13, 1) = "評　　語"
'      For iRow = 0 To 8
'         .MergeRow(iRow) = True
'         .TextMatrix(iRow, 2) = .TextMatrix(iRow, 1)
'      Next
'      .TextMatrix(9, 2) = "P"
'      .TextMatrix(10, 2) = "T"
'      .TextMatrix(11, 2) = "L"
'      .TextMatrix(12, 2) = "C"
      For iRow = 0 To UBound(strRowN)
        .TextMatrix(iRow, 1) = strRowN(iRow)
        'Add by Amy 2016/03/01
        If bolNoMemo = True And (iRow = GetValue("轉撥實績備註") Or iRow = GetValue("轉撥結餘備註")) Then
            .RowHeight(iRow) = 0
        Else
            .RowHeight(iRow) = 285
        End If
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
      .TextMatrix(UBound(strRowN), 2) = "專"            'modify by sonia 2019/7/31 P改專,與美珍討論
      .TextMatrix(UBound(strRowN) + 1, 2) = "商"        'modify by sonia 2019/7/31 T改商
      .TextMatrix(UBound(strRowN) + 2, 2) = "法"        'modify by sonia 2019/7/31 L改法
      .TextMatrix(UBound(strRowN) + 3, 2) = "創"        'modify by sonia 2019/7/31 C改創(ACS)
'end 2015/04/20
      '.MergeRow(13) = True
      '.TextMatrix(13, 2) = .TextMatrix(13, 1)
      .Refresh
      .Visible = True
   End With
End Sub

'Modify by Amy 2015/04/20 +欄位
Private Sub runWord()
   Dim stYear(0 To 1) As String, stMonth(0 To 1) As String, stTmp As String
   Dim iRow As Integer, iCol As Integer, lngCellWidth As Long
   Dim oTable As Word.Table
   Dim strFormula As String, bolSetFormula As Boolean '公式/是否設定公式
   Dim strFormat As String 'Add by Amy 2015/05/04
   Dim intRun As Integer 'Add by Amy 2016/03/01
   Dim oCell As Cell, intCell As Integer, oColWidth As Single, intFontSize As Integer 'Add by Amy 2024/04/19
   
On Error GoTo ErrHnd
   
   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
   
   g_WordAp.Documents.add
   g_WordAp.Visible = True
      
   'Modify  by Amy 2016/03/01 點數結算期起迄相同才顯示轉撥備註/標題文字改14/ 內文文字改為12 /提醒文字改10
   intRun = grdDataList.Rows
   If bolNoMemo = True Then intRun = intRun - 2
   With g_WordAp.Application
      .WindowState = wdWindowStateMaximize
      .ActiveWindow.ActivePane.View.Zoom.Percentage = 100 'Add by Amy 2024/04/19 避免表格跳錯位置
      .Selection.Font.Name = "標楷體"
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.Font.Size = 12
      '邊界
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(1.6)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1.4)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(1)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(1)
      
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
      
      .Selection.Font.Size = 12
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.TypeText Text:=stTmp
      
      .Selection.TypeParagraph
      .Selection.Font.Size = 12
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.TypeText Text:=lblSalesArea & "智權人員工作報告"
      
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
      intCell = 0: oColWidth = 0: intFontSize = 12 'Add by Amy 2024/04/19
      If grdDataList.Cols - 2 > 7 Then
         intFontSize = 10
      End If
      '列數+1("評語"另外印),欄數-2(0,2不印)
      'Modify by Amy 2024/04/10 原:NumRows:=intRun + 1,拿掉評語
      Set oTable = .Selection.Tables.add(Range:=.Selection.Range, NumRows:=intRun, NumColumns:=grdDataList.Cols - 2)
      If bolNoMemo = True Then intRun = intRun + 2 '若為跨月Word 不列轉撥實績/結餘備註,故後面合併儲存格的判斷要加回
      
      'Word 2007 預設沒有框線需另外指定
      With oTable
        .Borders(wdBorderLeft).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
        .Borders(wdBorderRight).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
        .Borders(wdBorderTop).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
        .Borders(wdBorderBottom).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
        .Borders(wdBorderVertical).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
        .Borders(wdBorderHorizontal).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
      End With
      '設定表格第一欄寬
      .Selection.Cells.SetWidth ColumnWidth:=85.4, RulerStyle:=wdAdjustFirstColumn
      '設定表格高度
      .Selection.Cells.SetHeight RowHeight:=26, HeightRule:=wdRowHeightExactly
      
      'Add by Amy 2024/04/19 S14共10欄,會導致第一個智權人員寬欄會非常小
      If grdDataList.Cols - 2 > 7 Then
         For Each oCell In oTable.Range.Cells
            '計錄第2欄的寬,再除平均
            If oCell.RowIndex = 1 And oCell.ColumnIndex >= 2 Then
               oColWidth = oColWidth + oCell.Width
               intCell = intCell + 1
            End If
         Next
         If intCell - 1 > 1 Then
            oColWidth = Round(oColWidth / (intCell), 2)
         End If
      End If
      'end 2024/04/19
     
      For iRow = 0 To grdDataList.Rows - 1
        If grdDataList.RowHeight(iRow) <> 0 Then
         .Selection.SelectRow
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
         For iCol = 1 To grdDataList.Cols - 1
            If iCol <> 1 And oColWidth > 0 Then
               .Selection.Cells.SetWidth ColumnWidth:=oColWidth, RulerStyle:=wdAdjustFirstColumn
            End If
            .Selection.Font.Size = intFontSize 'Add by Amy 2024/04/19 字太大會跳頁
            If iCol = 1 Then
               .Selection.Font.Size = 11 'Add by Amy 2024/04/19 字太大會跳頁
               '收文案源分析
               If grdDataList.TextMatrix(iRow, iCol) <> grdDataList.TextMatrix(iRow, iCol + 1) Then
                  '分割欄位
                  .Selection.MoveLeft Unit:=wdCharacter, Count:=1
                  lngCellWidth = .Selection.Cells(1).Width
                  .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
                  .Selection.Cells(1).SetWidth ColumnWidth:=lngCellWidth * 0.8, RulerStyle:=wdAdjustFirstColumn
                  .Selection.MoveLeft Unit:=wdCell, Count:=1
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
                  'Modify by Amy 2024/04/19 改名稱 原:達成點數->報出點,加[當月實績達成率][收文點數達成率]
                  If iRow = GetValue("報出點數") Or iRow = GetValue("報出點數達成率") _
                    Or iRow = GetValue("當月實績達成率") Or iRow = GetValue("收文點數達成率") Then
                      .Selection.Font.Bold = True
                   'Add by Amy 2016/03/01 +轉撥備註
                   ElseIf iRow = GetValue("轉撥實績備註") Or iRow = GetValue("轉撥結餘備註") Then
                      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
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
               .Selection.MoveRight Unit:=wdCell, Count:=1
               
            '資料內容
            ElseIf iCol > 2 Then
                strFormula = "": bolSetFormula = False: strFormat = ""
                If grdDataList.TextMatrix(0, iCol) = "合　　計" And iRow > 0 Then
                    Select Case iRow
                        'Modify by Amy 2024/04/19 改名稱原:達 成 率->報出點數達成率,加[當月實績達成率][收文點數達成率]
                        Case GetValue("報出點數達成率"), GetValue("當月實績達成率"), GetValue("收文點數達成率")
                            strExc(9) = GetValue(Replace(strRowN(iRow), "達成率", ""))
                            If Val(grdDataList.TextMatrix(GetValue("目標"), iCol)) > 0 And Val(grdDataList.TextMatrix(Val(strExc(9)), iCol)) > 0 Then
                              strFormat = "0.000%"
                              strFormula = "if(" & Chr(iCol + 63) & GetValue("目標") + 1 & "=0,0," & Chr(iCol + 63) & Val(strExc(9)) + 1 & "/" & Chr(iCol + 63) & GetValue("目標") + 1 & "*100)"
                            End If
                            'end 2024/04/19
                            bolSetFormula = True
                        'Modify by Amy 2016/03/01
                        'Case GetValue("達成點數"), GetValue("期初實績保留") To GetValue("期初結餘保留"), GetValue("當月　　實績") To GetValue("期末結餘保留")
                        'Modify by Amy 2024/04/19 達成點數->報出點數
                        Case GetValue("報出點數"), GetValue("期初實績保留") To GetValue("轉撥實績增減"), GetValue("期初結餘保留") To GetValue("轉撥結餘增減")
                            strFormula = "SUM(LEFT)"
                            'strFormat = "0.000"
                        Case Else
                            strFormula = "SUM(LEFT)"
                            'strFormat = "0.000"
                    End Select
                    bolSetFormula = True
                Else
                    Select Case iRow
                        'Modify by Amy 2024/04/19 改名稱原:達成點數
                        Case GetValue("報出點數")
                            'Modify by Amy 2016/03/01 +轉撥
                            If txtCloseDate(0) = txtCloseDate(1) Then
                                strFormula = Chr(iCol + 63) & GetValue("期初實績保留") + 1 & "+" & Chr(iCol + 63) & GetValue("當月實績") + 1 & _
                                                "-" & Chr(iCol + 63) & GetValue("期末實績保留") + 1 & "+" & Chr(iCol + 63) & GetValue("轉撥實績增減") + 1 & _
                                                "+" & Chr(iCol + 63) & GetValue("期初結餘保留") + 1 & "+" & Chr(iCol + 63) & GetValue("當月結餘") + 1 & _
                                                "-" & Chr(iCol + 63) & GetValue("期末結餘保留") + 1 & "+" & Chr(iCol + 63) & GetValue("轉撥結餘增減") + 1
                            Else
                                '跨月不show 備註欄
                                strFormula = Chr(iCol + 63) & GetValue("期初實績保留") + 1 & "+" & Chr(iCol + 63) & GetValue("當月實績") + 1 & _
                                                "-" & Chr(iCol + 63) & GetValue("期末實績保留") + 1 & "+" & Chr(iCol + 63) & GetValue("轉撥實績增減") + 1 & _
                                                "+" & Chr(iCol + 63) & GetValue("期初結餘保留") & "+" & Chr(iCol + 63) & GetValue("當月結餘") & _
                                                "-" & Chr(iCol + 63) & GetValue("期末結餘保留") & "+" & Chr(iCol + 63) & GetValue("轉撥結餘增減")
                            End If
                            bolSetFormula = True
                        'Modify by Amy 2024/04/19 改名稱原:達成率->報出點數達成率,加[當月實績達成率][收文點數達成率]
                        Case GetValue("報出點數達成率"), GetValue("當月實績達成率"), GetValue("收文點數達成率")
                            strExc(9) = GetValue(Replace(strRowN(iRow), "達成率", ""))
                            If Val(grdDataList.TextMatrix(GetValue("目標"), iCol)) > 0 And Val(grdDataList.TextMatrix(Val(strExc(9)), iCol)) > 0 Then
                              strFormat = "0.000%"
                              strFormula = "if(" & Chr(iCol + 63) & GetValue("目標") + 1 & "=0,0," & Chr(iCol + 63) & Val(strExc(9)) + 1 & "/" & Chr(iCol + 63) & GetValue("目標") + 1 & "*100)"
                            End If
                            'end 2024/04/19
                            bolSetFormula = True
                        'Modify by Amy 2015/04/30
'                        Case GetValue("實績保留動用")
'                            If Val(dblTot(GetValue("期末實績保留"))) > 0 Or Val(dblTot(GetValue("期末結餘保留"))) > 0 Then
'                                strFormula = Chr(iCol + 63) & GetValue("期初實績保留") + 1 & "-" & Chr(iCol + 63) & GetValue("期末實績保留") + 1
'                                strFormat = "0.000"
'                                bolSetFormula = True
'                            End If
'                        Case GetValue("結餘保留動用")
'                            If Val(dblTot(GetValue("期末實績保留"))) > 0 Or Val(dblTot(GetValue("期末結餘保留"))) > 0 Then
'                                strFormula = Chr(iCol + 63) & GetValue("期初結餘保留") + 1 & "+" & Chr(iCol + 63) & GetValue("當月　　結餘") + 1 & "-" & Chr(iCol + 63) & GetValue("期末結餘保留") + 1
'                                strFormat = "0.000"
'                                bolSetFormula = True
'                            End If
'                        Case GetValue("期末實績保留")
'                            If Val(dblTot(GetValue("期末實績保留"))) = 0 And Val(dblTot(GetValue("期末結餘保留"))) = 0 Then
'                                strFormula = Chr(iCol + 63) & GetValue("期初實績保留") + 1 & "-" & Chr(iCol + 63) & GetValue("實績保留動用") + 1
'                                strFormat = "0.000"
'                                bolSetFormula = True
'                            End If
'                        Case GetValue("期末結餘保留")
'                            If Val(dblTot(GetValue("期末實績保留"))) = 0 And Val(dblTot(GetValue("期末結餘保留"))) = 0 Then
'                                strFormula = Chr(iCol + 63) & GetValue("期初結餘保留") + 1 & "+" & Chr(iCol + 63) & GetValue("當月　　結餘") + 1 & "-" & Chr(iCol + 63) & GetValue("結餘保留動用") + 1
'                                strFormat = "0.000"
'                                bolSetFormula = True
'                            End If
                        'end 2015/04/30
                        Case Else
                    End Select
                End If
                'Add by Amy 2016/03/01
                If grdDataList.TextMatrix(0, iCol) = "合　　計" And (iRow = GetValue("轉撥實績備註") Or iRow = GetValue("轉撥結餘備註") Or iRow = GetValue("出缺勤狀況")) Then
                    '空白
                ElseIf bolSetFormula = True Then
                    .Selection.Font.Bold = True
                    'Modify by Amy 2024/04/19 0會出現發預期之公式
                    If strFormula <> MsgText(601) Then
                        .Selection.InsertFormula Formula:="=" & strFormula, NumberFormat:=strFormat
                    ElseIf Val(strFormula) = 0 Then
                        If iRow = GetValue("報出點數達成率") Or iRow = GetValue("當月實績達成率") Or iRow = GetValue("收文點數達成率") Then
                           .Selection.TypeText Text:="0%"
                        Else
                           .Selection.TypeText Text:="0"
                        End If
                    Else
                        .Selection.InsertFormula Formula:="=" & strFormula
                    End If
                Else
                    'Add by Amy 2024/04/19 +改字大小/靠左及有全型空白拿掉,避免欄位不夠寬
                    If iRow = GetValue("轉撥實績備註") Or iRow = GetValue("轉撥結餘備註") Or iRow = GetValue("出缺勤狀況") Then
                        .Selection.ParagraphFormat.Alignment = wdAlignRowLeft
                        .Selection.Font.Name = "標楷體"
                        .Selection.Font.Size = 11
                    End If
                    .Selection.Font.Bold = False
                    strExc(2) = grdDataList.TextMatrix(iRow, iCol)
                    If InStr(strExc(2), "　　") > 0 Then strExc(2) = Replace(strExc(2), "　　", "")
                    .Selection.TypeText Text:=strExc(2)
                    'end 2024/04/19
                End If
                If iRow = intRun - 1 And iCol = grdDataList.Cols - 1 Then
                Else
                  .Selection.MoveRight Unit:=wdCell, Count:=1
                End If
            End If
         Next iCol
            'Mark by Amy 2024/04/19 拿掉評語列,不需再往右,否則會自動加一列
            '.Selection.MoveRight Unit:=wdCharacter, Count:=1
        End If '
      Next iRow
      '合併收文案源分析
      'Modify by Amy 2024/04/19 拿掉評語,[收文案源分析]不需再往右,否則會自動加一列
'      .Selection.MoveUp Unit:=wdLine, Count:=1
'      .Selection.MoveUp Unit:=wdLine, Count:=3, Extend:=wdExtend
'      .Selection.Cells.Merge
'      .Selection.MoveDown Unit:=wdLine, Count:=1
      .Selection.Find.ClearFormatting
      .Selection.Find.Text = "收文案源分析"
      .Selection.Find.Replacement.Text = ""
      .Selection.Find.Forward = True
      .Selection.Find.Wrap = wdFindContinue
      .Selection.Find.Format = False
      .Selection.Find.MatchCase = False
      .Selection.Find.MatchWholeWord = False
      .Selection.Find.MatchWildcards = False
      .Selection.Find.MatchSoundsLike = False
      .Selection.Find.MatchAllWordForms = False
      .Selection.Find.MatchByte = True
      If .Selection.Find.Execute = True Then
         .Selection.MoveDown Unit:=wdLine, Count:=3, Extend:=wdExtend
         .Selection.Cells.Merge
         .Selection.MoveDown Unit:=wdLine, Count:=1
      End If
      
'      '評　　語
'      .Selection.SelectRow
'      'Modify by Amy 2016/03/01 原:RowHeight:=80
'      .Selection.Cells.SetHeight RowHeight:=30, HeightRule:=wdRowHeightExactly
'      .Selection.Cells.VerticalAlignment = wdCellAlignVerticalTop
'      .Selection.TypeText Text:="評　　語"
'      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
'      .Selection.MoveRight unit:=wdCharacter, Count:=grdDataList.Cols - 1
      'end 2024/04/19
      
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
      .Selection.Font.Size = 12
      'Modify by Amy 2024/04/10 改顯示文字
      '.Selection.TypeText Text:="二、工作情形（當月發生之重要案件、市場反應…等處理）："
      .Selection.TypeText Text:="二、智權人員之個人工作考核說明："
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      
      'Add by Amy 2024/04/10 +三及四項畫表格顯示資料
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
      .Selection.TypeText Text:="三、我國商標新申請案收文、發文狀況說明："
      Call NewCaseSituation("T", Val(txtCloseDate(0) & "01") + 19110000, Val(txtCloseDate(1) & "31") + 19110000, txtSalesArea)
      .Selection.TypeParagraph
      
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
      .Selection.TypeText Text:="四、我國專利新申請案收文、發文狀況說明："
      Call NewCaseSituation("P", Val(txtCloseDate(0) & "01") + 19110000, Val(txtCloseDate(1) & "31") + 19110000, txtSalesArea)
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      'end 2024/04/10
      .Selection.TypeText Text:=String(20, "　") & "報告人：" & strUserName
      .Selection.TypeParagraph
      .Selection.TypeText Text:=String(20, "　") & "（各分區主管每月8日前填報）"
      'Modify by Amy 2016/03/01 105年01月抓業績輸入不需顯示此文字
      If Val(txtCloseDate(0)) < Val(業績輸入啟用年月) Then
      .Selection.TypeParagraph
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
      'Modify by Amy 2024/04/19 拿掉評語後改抓最後一句話
'      If .Selection.Find.Execute("評　　語") = True Then
'         .Selection.MoveRight Unit:=wdCharacter, Count:=2 '右移一格
'      End If
      If .Selection.Find.Execute("各分區主管每月8日前填報") = True Then
         .Selection.MoveRight Unit:=wdCharacter, Count:=1 '右移 一個字
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

'Add by Amy 2018/04/18 寫入暫存檔,因高國碩、陳頌恩10610月由中一轉中二,若下跨月期未保留及結餘值顯示不出來
Private Function doQuery() As Boolean
    Dim stConST As String, stConPE As String, strSqlSP As String, strTotCase(3) As String, strSumR As String, strSumPE As String
    Dim iRow As Single, iCol As Single, i As Integer, j As Integer, intQ As Integer
    Dim strSP15 As String, strSP36 As String, strSP19 As String, strSP40 As String, stSP20 As String, stSP41 As String
    Dim stConCu As String, stConCP As String, strQ As String
    Dim stField(1) As String, stDate(1) As String, bolSetData As Boolean 'Add by Amy 2024/04/19
    
    'Modify by Amy 2024/04/19
    ReDim dblTot(1 To UBound(strRowN))
On Error GoTo ErrHnd
    '刪除暫存檔資料
    strQ = "Delete From R210113 Where ID='" & strUserNum & "' And SubStr(R04,1,1)='1'"
    'end 2024/04/19
    adoTaie.Execute strQ
    
    '將SalesPoint人員寫入暫存檔(有控制只可以查SalesPoint上線後的資料)
    'Modify by Amy 2024/04/19 +收文點數(R05) ex:S21 11303月 20012(中一區備用)表一人員也要出現-杜協理
'    strQ = "Insert Into R210113 (ID,R01,R02,R03,R04) " & _
'              "Select Distinct '" & strUserNum & "',SP02,'" & txtSalesArea & "',SP01,'1' From SalesPoint " & _
'              "Where SP01>=191100+" & Val(txtCloseDate(0)) & " And SP01<=191100+" & Val(txtCloseDate(1)) & _
'              " And SP48='" & txtSalesArea & "'"
    stDate(0) = Val(txtCloseDate(0)) + 191100 & "01"
    stDate(1) = Val(txtCloseDate(1)) + 191100 & "31"
    '收文點數
    strExc(2) = PUB_CountCP18("0", stDate(0), stDate(1), txtSalesArea, txtSalesArea, , strQ, , , 0, Me.Name, , True, 3)
    strQ = Replace(UCase(strQ), UCase("Select CP13,Round(sum(cp18a),3) 收文點數"), " Union Select CP13,Round(sum(cp18a),3) Point")
    
    strQ = "Insert Into R210113 (ID,R01,R02,R03,R04,R05) " & _
              "Select '" & strUserNum & "',CP13,'" & txtSalesArea & "'," & Left(stDate(0), 6) & ",'1',Sum(Point) From (" & _
              "Select SP02 as CP13,0 as Point From SalesPoint Where SP48='" & txtSalesArea & "' " & _
                            "And SP01>=191100+" & Val(txtCloseDate(0)) & " And SP01<=191100+" & Val(txtCloseDate(1)) & _
             strQ & ") Group by CP13 "
    adoTaie.Execute strQ
    'end 2024/04/19
    
    'Sql語法改寫至共用
    strSql = Replace(UCase(GetPoint(0, Val(txtCloseDate(0)), Val(txtCloseDate(1)), txtSalesArea, , , , Me.Name)), "AND (PE04>0 OR V11>0 OR V21>0 OR V31>0 OR V41>0)", "AND (R05>0 OR PE04>0 OR V11>0 OR V21>0 OR V31>0 OR V41>0)")
    strSqlSP = GetPoint_SP(Val(txtCloseDate(0)), Val(txtCloseDate(1)), txtSalesArea, , , , Me.Name, bolNoMemo)
    
    'Modify by Amy 2024/04/19 寫成一句並加,原使用AdoRecordSet3/adoRecordset1,使用舊部門因以st15為主
    stField(0) = ",SP20,SP41" '轉撥備註
    If bolNoMemo = True Then stField(0) = ""
    stField(1) = Replace(Replace(stField(0), ",SP20", ",'' as SP20"), ",SP41", ",'' as SP41")
    strQ = "Select ST02,R05,PE04,C1,C2,C3,C4,C5,C6,C7,C8,C9,C10,C11,C12,SP15,SP19,SP36,SP40" & stField(0) & ",st01,Decode(ST05,'SM',1,'76',1,2) as Sort " & _
                "From (" & strSql & "),(" & strSqlSP & ") Where ST01=mST01(+) " & _
   "Union Select '合　　計','0',0 as PE04,0,0,0,0,0 as C5,0,0,0,0,0 as C10,0,0,0 as SP15,0 as SP19,0,0" & stField(1) & ",'ZZZ' as st01,9 as Sort From Dual " & _
               "Order by Sort,st01 "
   intQ = 1
   Set rsA = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      '數值欄位以整數顯示,有小數才顯示小數/增加欄位(同frm210152智權點數實績與結餘輸入的全區資料)
      'Memo by Lydia 2019/08/08 frm210152智權點數實績與結餘輸入在承辦人系統更名為「每月點數查詢／輸入」，與財務系統不同
      grdDataList.Visible = False
      With rsA
   'end 2024/04/19
      iCol = 3
      Do While Not .EOF
         strSumR = "": strSP15 = "": strSP36 = ""
         strSP19 = "": stSP20 = "": strSP40 = "": stSP41 = ""
         
         '業績輸入上線後改抓SalesPoint資料
         If Val(txtCloseDate(0)) >= Val(業績輸入啟用年月) Then
             'Modify by Amy 2024/04/19 語法改成一句
             'Modify by Amy 2018/05/11 因瑞婷產生智權實績結餘輸入之傳票後,將當月實績吳金榮之傳票改成王俊剴
             '                                         造成AdoRecordSet3無吳金榮的資料但adoRecordset1有吳金榮的資料
'aa:         If intQ = 1 And adoRecordset1.EOF = False Then
'                If "" & .Fields("ST02") = "" & adoRecordset1.Fields("ST02") Then
                    strSP15 = Val("" & .Fields("SP15"))
                    strSP36 = Val("" & .Fields("SP36"))
                    strSP19 = Val("" & .Fields("SP19"))
                    If bolNoMemo = False Then stSP20 = "" & .Fields("SP20")
                    strSP40 = Val("" & .Fields("SP40"))
                    If bolNoMemo = False Then stSP41 = "" & .Fields("SP41")
'                    adoRecordset1.MoveNext
'                'Modify by Amy 2018/05/11
'                Else
'                    adoRecordset1.MoveNext
'                    GoTo aa
'                'end 2018/05/11
'                End If
'            End If
             'end 2024/04/19
         Else
            strSP15 = Val("" & .Fields("C5"))
            strSP36 = Val("" & .Fields("C6"))
         End If
         
         grdDataList.Cols = iCol + 1
         For iRow = GetValue("智權人員") To UBound(strRowN)
            grdDataList.col = iCol
            grdDataList.row = iRow
            If InStr(strRowN(iRow), "報出點數") = 0 And InStr(strRowN(iRow), "保留動用") = 0 Then
               strExc(9) = "": bolSetData = True
               Select Case Replace(strRowN(iRow), "　　", "")
                  Case "智權人員"
                     strExc(9) = "" & .Fields("ST02")
                  Case "目標"
                     If "" & rsA.Fields("st01") = "ZZZ" Then
                        strExc(9) = strSumPE
                     Else
                        strExc(9) = "" & .Fields("PE04")
                        strSumPE = Val(strSumPE) + Val("" & .Fields("PE04"))
                     End If
                  Case "期初實績保留", "期初結餘保留"
                     If "" & rsA.Fields("st01") = "ZZZ" Then
                        strExc(9) = dblTot(iRow)
                     Else
                        If InStr(strRowN(iRow), "實績") > 0 Then
                           strExc(9) = Round(Val("" & .Fields("C1")), 3)
                        Else
                           strExc(9) = Round(Val("" & .Fields("C2")), 3)
                        End If
                        dblTot(iRow) = dblTot(iRow) + Val(strExc(9))
                        strSumR = Val(strSumR) + Val(strExc(9))
                     End If
                  Case "當月實績", "當月結餘"
                     If "" & rsA.Fields("st01") = "ZZZ" Then
                        strExc(9) = dblTot(iRow)
                     Else
                        If InStr(strRowN(iRow), "實績") > 0 Then
                           strExc(9) = Round(Val("" & .Fields("C3")), 3)
                        Else
                           strExc(9) = Round(Val("" & .Fields("C4")), 3)
                        End If
                        dblTot(iRow) = dblTot(iRow) + Val(strExc(9))
                        strSumR = Val(strSumR) + Val(strExc(9))
                     End If
                  'Add by Amy2024/04/19
                  Case "當月實績達成率", "收文點數達成率"
                     bolSetData = False
                     j = GetValue(Replace("" & strRowN(iRow), "達成率", ""))
                     strExc(9) = grdDataList.TextMatrix(j, iCol)
                     If Val(grdDataList.TextMatrix(GetValue("目標"), iCol)) > 0 And Val(strExc(9)) > 0 Then
                        grdDataList.TextMatrix(iRow, iCol) = Format(100 * Val(strExc(9)) / Val(grdDataList.TextMatrix(GetValue("目標"), iCol)), "0.000") & "%"
                     Else
                        grdDataList.TextMatrix(iRow, iCol) = "0%"
                     End If
                  Case "期末實績保留", "期末結餘保留"
                     If "" & rsA.Fields("st01") = "ZZZ" Then
                        strExc(9) = dblTot(iRow)
                     Else
                        If InStr(strRowN(iRow), "實績") > 0 Then
                           strExc(9) = Round(Val(strSP15), 3)
                        Else
                           strExc(9) = Round(Val(strSP36), 3)
                        End If
                        dblTot(iRow) = dblTot(iRow) + Val(strExc(9))
                        strSumR = Val(strSumR) - Val(strExc(9))
                     End If
                  Case "轉撥實績增減", "轉撥結餘增減"
                     If "" & rsA.Fields("st01") = "ZZZ" Then
                        strExc(9) = dblTot(iRow)
                     Else
                        If InStr(strRowN(iRow), "實績") > 0 Then
                           strExc(9) = Round(Val(strSP19), 3)
                        Else
                           strExc(9) = Round(Val(strSP40), 3)
                        End If
                        dblTot(iRow) = dblTot(iRow) + Val(strExc(9))
                     End If
                  Case "轉撥實績備註", "轉撥結餘備註"
                     If "" & rsA.Fields("st01") = "ZZZ" Then
                        strExc(9) = ""
                     Else
                        '起迄年月相同才顯示轉撥備註
                        If bolNoMemo = False Then
                           If InStr(strRowN(iRow), "實績") > 0 Then
                              strExc(9) = stSP20
                           Else
                              strExc(9) = stSP41
                           End If
                        End If
                     End If
                  Case "報出實績點數", "報出結餘點數"
                     If "" & rsA.Fields("st01") = "ZZZ" Then
                        strExc(9) = dblTot(iRow)
                        grdDataList.row = iRow
                        For j = 1 To grdDataList.Cols - 1
                           grdDataList.col = j
                           grdDataList.CellBackColor = &HC000&
                        Next j
                     Else
                        If InStr(strRowN(iRow), "實績") > 0 Then
                           strExc(9) = Round(Val("" & .Fields("C1")) + Val("" & .Fields("C3")) - Val(strSP15) + Val(strSP19), 3)
                        Else
                           strExc(9) = Round(Val("" & .Fields("C2")) + Val("" & .Fields("C4")) - Val(strSP36) + Val(strSP40), 3)
                        End If
                        dblTot(iRow) = dblTot(iRow) + Val(strExc(9))
                     End If
                  'Add by Amy2024/04/19
                  Case "收文點數"
                     If "" & rsA.Fields("st01") = "ZZZ" Then
                        strExc(9) = dblTot(iRow)
                     Else
                        strExc(9) = PUB_ChkExcelZero(1, "" & .Fields("R05"))
                        dblTot(iRow) = dblTot(iRow) + Val(strExc(9))
                     End If
                  Case "出缺勤狀況"
                     If "" & rsA.Fields("st01") = "ZZZ" Then
                        strExc(9) = ""
                     Else
                        strExc(2) = "And ST01='" & rsA.Fields("st01") & "' "
                        strExc(9) = GetAbsenceData(strExc(2), stDate(0), stDate(1))
                     End If
                  Case "新增客戶數"
                     If "" & rsA.Fields("st01") = "ZZZ" Then
                        strExc(9) = dblTot(iRow)
                     Else
                        strExc(9) = Val("" & .Fields("C7"))
                        dblTot(iRow) = dblTot(iRow) + Val(strExc(9))
                     End If
                  Case "收款家數"
                     If "" & rsA.Fields("st01") = "ZZZ" Then
                        strExc(9) = dblTot(iRow)
                     Else
                        strExc(9) = Val("" & .Fields("C8"))
                        dblTot(iRow) = dblTot(iRow) + Val(strExc(9))
                     End If
                  Case "收文案源分析"
                     bolSetData = False
                     If "" & rsA.Fields("st01") = "ZZZ" Then
                        For j = LBound(strTotCase) To UBound(strTotCase)
                           grdDataList.TextMatrix(iRow + j, iCol) = strTotCase(j)
                        Next j
                        j = GetValue("報出點數")
                        grdDataList.TextMatrix(j, iCol) = dblTot(j)
                        grdDataList.row = j
                        For j = 1 To grdDataList.Cols - 1
                           grdDataList.col = j
                           grdDataList.CellBackColor = &HC000&
                        Next j
                        '實績保留動用
                        j = GetValue("實績保留動用")
                        grdDataList.TextMatrix(j, iCol) = dblTot(j)
                        '結餘保留動用
                         j = GetValue("結餘保留動用")
                         grdDataList.TextMatrix(j, iCol) = dblTot(j)
                     Else
                        'P
                        grdDataList.TextMatrix(iRow, iCol) = Val("" & .Fields("C9"))
                        strTotCase(0) = Val(strTotCase(0)) + Val(grdDataList.TextMatrix(iRow, iCol))
                        'T
                        grdDataList.TextMatrix(iRow + 1, iCol) = Val("" & .Fields("C10"))
                        strTotCase(1) = Val(strTotCase(1)) + Val(grdDataList.TextMatrix(iRow + 1, iCol))
                        'L
                        grdDataList.TextMatrix(iRow + 2, iCol) = Val("" & .Fields("C11"))
                        strTotCase(2) = Val(strTotCase(2)) + Val(grdDataList.TextMatrix(iRow + 2, iCol))
                        'C
                        grdDataList.TextMatrix(iRow + 3, iCol) = Val("" & .Fields("C12"))
                        strTotCase(3) = Val(strTotCase(3)) + Val(grdDataList.TextMatrix(iRow + 3, iCol))
                        '*** 資料寫入後,最後計算產生的值 ***
                        '達成點數 'Modify by Amy 2024/04/19 原:達成點數
                        j = GetValue("報出點數")
                        strExc(0) = Val(grdDataList.TextMatrix(GetValue("期初實績保留"), iCol)) + Val(grdDataList.TextMatrix(GetValue("期初結餘保留"), iCol))
                        strExc(0) = Val(strExc(0)) + Val(grdDataList.TextMatrix(GetValue("當月實績"), iCol)) + Val(grdDataList.TextMatrix(GetValue("當月結餘"), iCol))
                        strExc(0) = Val(strExc(0)) - Val(grdDataList.TextMatrix(GetValue("期末實績保留"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末結餘保留"), iCol))
                        '轉撥
                        strExc(0) = Val(strExc(0)) + Val(grdDataList.TextMatrix(GetValue("轉撥實績增減"), iCol)) + Val(grdDataList.TextMatrix(GetValue("轉撥結餘增減"), iCol))
                        grdDataList.TextMatrix(j, iCol) = Round(Val(strExc(0)), 3)
                        dblTot(j) = dblTot(j) + Val(grdDataList.TextMatrix(j, iCol))
                        
                        '實績保留動用
                         j = GetValue("實績保留動用")
                        grdDataList.TextMatrix(j, iCol) = Round(Val(grdDataList.TextMatrix(GetValue("期初實績保留"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末實績保留"), iCol)), 3)
                        dblTot(j) = dblTot(j) + Val(grdDataList.TextMatrix(j, iCol))
                         '結餘保留動用
                         j = GetValue("結餘保留動用")
                        grdDataList.TextMatrix(j, iCol) = Round(Val(grdDataList.TextMatrix(GetValue("期初結餘保留"), iCol)) + Val(grdDataList.TextMatrix(GetValue("當月結餘"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末結餘保留"), iCol)), 3)
                        dblTot(j) = dblTot(j) + Val(grdDataList.TextMatrix(j, iCol))
                     '*** End 資料寫入後,最後計算產生的值 ***
                     End If
                     j = GetValue("報出點數達成率") 'Modify by Amy 2024/04/19 改名稱原:達 成 率->報出點數達成率
                     strExc(9) = grdDataList.TextMatrix(GetValue("報出點數"), iCol)
                     If Val(grdDataList.TextMatrix(GetValue("目標"), iCol)) > 0 And Val(strExc(9)) > 0 Then
                        grdDataList.TextMatrix(j, iCol) = Format(100 * Val(strExc(9)) / Val(grdDataList.TextMatrix(GetValue("目標"), iCol)), "0.000") & "%"
                     Else
                        grdDataList.TextMatrix(j, iCol) = "0%"
                     End If
                     
                  Case Else
               End Select
               If bolSetData = True Then
                  grdDataList.TextMatrix(iRow, iCol) = strExc(9)
               End If
               
               If strRowN(iRow) = "智權人員" Then
                  grdDataList.CellAlignment = flexAlignCenterCenter
               ElseIf strRowN(iRow) = "出缺勤狀況" Then
                  grdDataList.CellAlignment = flexAlignLeftCenter
               Else
                  grdDataList.CellAlignment = flexAlignRightCenter
               End If
            End If
         Next iRow
         
         .MoveNext
         iCol = iCol + 1
      Loop
      End With
     
      grdDataList.Visible = True
   Else
      MsgBox "無符合資料！", vbInformation
   End If
   
   doQuery = True
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Function doQuery_Old3() As Boolean
'Mark by Amy 2018/04/18 寫入暫存檔,因高國碩、陳頌恩10610月由中一轉中二,若下跨月期未保留及結餘值顯示不出來
''Memo 2016/01/30 抓目標時業務區不知如何抓舊資料,且日期跨上線日期也不易抓資料,故不可查舊資料
'    Dim stConST As String, stConPE As String, strSqlSP As String
'    Dim iRow As Single, iCol As Single
'    Dim dblTotCase(3) As Double, dblSumR As Double
'    Dim dblSP15 As Double, dblSP36 As Double
'    Dim stConCu As String, stConCP As String
'    Dim i As Integer, j As Integer, intQ As Integer
'    Erase dblTot
'    'Add by Amy 2016/03/01
'    Dim dblSP19 As Double, dblSP40 As Double
'    Dim stSP20 As String, stSP41 As String
'
'    'Mark by Amy 2016/03/11
''   '業務區
''   If txtSalesArea <> "" Then
''      stConST = " AND ST15='" & txtSalesArea & "' And ST01<'F'"
''   End If
''
''   '點數結算日
''   If txtCloseDate(0) <> "" Then
''      stConPE = stConPE & " AND PE03(+) >= 191100+" & txtCloseDate(0)
''   End If
''
''   If txtCloseDate(1) <> "" Then
''      stConPE = stConPE & " AND PE03(+) <= 191100+" & txtCloseDate(1)
''   End If
'
'On Error GoTo ErrHnd
'   'Sql語法改寫至Query
'    strSql = GetPoint(0, Val(txtCloseDate(0)), Val(txtCloseDate(1)), txtSalesArea, txtSalesArea, , , Me.Name)
'    strSqlSP = GetPoint_SP(Val(txtCloseDate(0)), Val(txtCloseDate(1)), txtSalesArea, txtSalesArea, , , Me.Name, bolNoMemo)
'    'end 2016/03/01
'
'   intI = 1: intQ = 1
'   Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strSql)
'   Set adoRecordset1 = ClsLawReadRstMsg(intQ, strSqlSP)
'
'   If intI = 1 Then
'      'Modify by Amy 2016/03/01 數值欄位以整數顯示,有小數才顯示小數/增加欄位(同frm210152智權點數實績與結餘輸入的全區資料)
'      grdDataList.Visible = False
'      With AdoRecordSet3
'      iCol = 3
'
'      Do While Not .EOF
'         dblSumR = 0: dblSP15 = 0: dblSP36 = 0
'         'Add by Amy 2016/03/01
'         dblSP19 = 0: stSP20 = "": dblSP40 = 0: stSP41 = ""
'         grdDataList.Cols = iCol + 1
'
'         '智權人員
'         iRow = GetValue("智權人員")
'         grdDataList.TextMatrix(iRow, iCol) = "" & .Fields("ST02")
'         grdDataList.CellAlignment = flexAlignCenterCenter
'         '目　　標
'         iRow = GetValue("目　　標")
'         grdDataList.TextMatrix(iRow, iCol) = "" & .Fields("PE04")
'         grdDataList.CellAlignment = flexAlignRightCenter
'         'Modify by Amy 2016/03/01 調位置
'         '業績輸入上線後改抓SalesPoint資料
'         If Val(txtCloseDate(0)) >= Val(業績輸入啟用年月) Then
'            If intQ = 1 And adoRecordset1.EOF = False Then
'                If "" & .Fields("ST02") = "" & adoRecordset1.Fields("ST02") Then
'                    dblSP15 = Val("" & adoRecordset1.Fields("SP15"))
'                    dblSP36 = Val("" & adoRecordset1.Fields("SP36"))
'                    'Add by Amy 2016/03/01
'                    dblSP19 = Val("" & adoRecordset1.Fields("SP19"))
'                    If bolNoMemo = False Then stSP20 = "" & adoRecordset1.Fields("SP20")
'                    dblSP40 = Val("" & adoRecordset1.Fields("SP40"))
'                    If bolNoMemo = False Then stSP41 = "" & adoRecordset1.Fields("SP41")
'                    adoRecordset1.MoveNext
'                End If
'            End If
'         Else
'            dblSP15 = Val("" & .Fields("C5"))
'            dblSP36 = Val("" & .Fields("C6"))
'         End If
'         '期初實績保留
'         iRow = GetValue("期初實績保留")
'         grdDataList.TextMatrix(iRow, iCol) = Round(Val("" & .Fields("C1")), 3) 'Format(Val("" & .Fields("C1")), "0.000")
'         dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'         dblSumR = dblSumR + Val(grdDataList.TextMatrix(iRow, iCol))
'         '當月　　實績
'         iRow = GetValue("當月　　實績")
'          grdDataList.TextMatrix(iRow, iCol) = Round(Val("" & .Fields("C3")), 3) 'Format(Val("" & .Fields("C3")), "0.000")
'         dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'         dblSumR = dblSumR + Val(grdDataList.TextMatrix(iRow, iCol))
'         '期末實績保留
'         iRow = GetValue("期末實績保留")
'         grdDataList.TextMatrix(iRow, iCol) = Round(dblSP15, 3) 'Format(dblSP15, "0.000")
'         dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'         dblSumR = dblSumR - Val(grdDataList.TextMatrix(iRow, iCol))
'         '轉撥實績增減
'         iRow = GetValue("轉撥實績增減")
'         grdDataList.TextMatrix(iRow, iCol) = Round(dblSP19, 3) 'Format(dblSP19, "0.000")
'         dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'         'Add byAmy 2016/03/01 起迄年月相同才顯示轉撥備註
'         If bolNoMemo = False Then
'            '轉撥實績備註
'            iRow = GetValue("轉撥實績備註")
'            grdDataList.TextMatrix(iRow, iCol) = stSP20
'         End If
'         '報出實績點數
'         iRow = GetValue("報出實績點數")
'         grdDataList.TextMatrix(iRow, iCol) = Round(Val("" & .Fields("C1")) + Val("" & .Fields("C3")) - Val(dblSP15) + Val(dblSP19), 3)
'         dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'
'         '期初結餘保留
'         iRow = GetValue("期初結餘保留")
'         grdDataList.TextMatrix(iRow, iCol) = Round(Val("" & .Fields("C2")), 3) 'Format(Val("" & .Fields("C2")), "0.000")
'         dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'         dblSumR = dblSumR + Val(grdDataList.TextMatrix(iRow, iCol))
'         '當月　　結餘
'         iRow = GetValue("當月　　結餘")
'         grdDataList.TextMatrix(iRow, iCol) = Round(Val("" & .Fields("C4")), 3) 'Format(Val("" & .Fields("C4")), "0.000")
'         dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'         dblSumR = dblSumR + Val(grdDataList.TextMatrix(iRow, iCol))
'         '期末結餘保留
'         iRow = GetValue("期末結餘保留")
'         grdDataList.TextMatrix(iRow, iCol) = Round(dblSP36, 3) 'Format(dblSP36, "0.000")
'         dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'         dblSumR = dblSumR - Val(grdDataList.TextMatrix(iRow, iCol))
'         '轉撥結餘增減
'         iRow = GetValue("轉撥結餘增減")
'         grdDataList.TextMatrix(iRow, iCol) = Round(dblSP40, 3) 'Format(dblSP40, "0.000")
'         dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'         'Add byAmy 2016/03/01 起迄年月相同才顯示轉撥備註
'         If bolNoMemo = False Then
'            '轉撥結餘備註
'            iRow = GetValue("轉撥結餘備註")
'            grdDataList.TextMatrix(iRow, iCol) = stSP41
'         End If
'         '報出結餘點數
'         iRow = GetValue("報出結餘點數")
'         grdDataList.TextMatrix(iRow, iCol) = Round(Val("" & .Fields("C2")) + Val("" & .Fields("C4")) - Val(dblSP36) + Val(dblSP40), 3)
'         dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'
'         '新增客戶數
'         iRow = GetValue("新增客戶數")
'         grdDataList.TextMatrix(iRow, iCol) = Val("" & .Fields("C7"))
'         'Modify by Amy 2016/03/01
'         'dblTot(iRow - 2) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'         dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'         '收款家數
'         iRow = GetValue("收款家數")
'         grdDataList.TextMatrix(iRow, iCol) = Val("" & .Fields("C8"))
'         'Modify by Amy 2016/03/01
'         'dblTot(iRow - 2) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'         dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'         '收文案件來源分析
'         'P
'         iRow = GetValue("收文案源分析")
'         grdDataList.TextMatrix(iRow, iCol) = Val("" & .Fields("C9"))
'         dblTotCase(0) = dblTotCase(0) + Val(grdDataList.TextMatrix(iRow, iCol))
'         'T
'         iRow = GetValue("收文案源分析") + 1
'         grdDataList.TextMatrix(iRow, iCol) = Val("" & .Fields("C10"))
'         dblTotCase(1) = dblTotCase(1) + Val(grdDataList.TextMatrix(iRow, iCol))
'         'L
'         iRow = GetValue("收文案源分析") + 2
'         grdDataList.TextMatrix(iRow, iCol) = Val("" & .Fields("C11"))
'         dblTotCase(2) = dblTotCase(2) + Val(grdDataList.TextMatrix(iRow, iCol))
'         'C
'         iRow = GetValue("收文案源分析") + 3
'         grdDataList.TextMatrix(iRow, iCol) = Val("" & .Fields("C12"))
'         dblTotCase(3) = dblTotCase(3) + Val(grdDataList.TextMatrix(iRow, iCol))
'
'         '達成點數
'         iRow = GetValue("達成點數")
'         strExc(0) = Val(grdDataList.TextMatrix(GetValue("期初實績保留"), iCol)) + Val(grdDataList.TextMatrix(GetValue("期初結餘保留"), iCol))
'         strExc(0) = Val(strExc(0)) + Val(grdDataList.TextMatrix(GetValue("當月　　實績"), iCol)) + Val(grdDataList.TextMatrix(GetValue("當月　　結餘"), iCol))
'         strExc(0) = Val(strExc(0)) - Val(grdDataList.TextMatrix(GetValue("期末實績保留"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末結餘保留"), iCol))
'         'Add by Amy 2016/03/01 +轉撥
'         strExc(0) = Val(strExc(0)) + Val(grdDataList.TextMatrix(GetValue("轉撥實績增減"), iCol)) + Val(grdDataList.TextMatrix(GetValue("轉撥結餘增減"), iCol))
'         grdDataList.TextMatrix(iRow, iCol) = Round(strExc(0), 3) 'Format(strExc(0), "0.000")
'         dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'
'         '達成率
'         iRow = GetValue("達 成 率")
'         If Val(grdDataList.TextMatrix(GetValue("目　　標"), iCol)) > 0 And Val(grdDataList.TextMatrix(GetValue("達成點數"), iCol)) > 0 Then
'            grdDataList.TextMatrix(iRow, iCol) = Format(100 * Val(grdDataList.TextMatrix(GetValue("達成點數"), iCol)) / Val(grdDataList.TextMatrix(GetValue("目　　標"), iCol)), "0.000") & "%"
'         Else
'            grdDataList.TextMatrix(iRow, iCol) = "0%"
'         End If
'
'         '實績保留動用
'         iRow = GetValue("實績保留動用")
'         'Mark by Amy 2016/01/30 取消小於等於0顯示0(公式),因業績點數上線需與其資料一致
''         If Val(grdDataList.TextMatrix(GetValue("期初實績保留"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末實績保留"), iCol)) > 0 Then
'            'grdDataList.TextMatrix(iRow, iCol) = Format(Val(grdDataList.TextMatrix(GetValue("期初實績保留"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末實績保留"), iCol)), "0.000")
'            grdDataList.TextMatrix(iRow, iCol) = Round(Val(grdDataList.TextMatrix(GetValue("期初實績保留"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末實績保留"), iCol)), 3)
'            dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
''         Else
''            grdDataList.TextMatrix(iRow, iCol) = Format("0", "0.000")
''         End If
'
'         '結餘保留動用
'         iRow = GetValue("結餘保留動用")
'         'Mark by Amy 2016/01/30 取消小於等於0顯示0(公式),因業績點數上線需與其資料一致
''         If Val(grdDataList.TextMatrix(GetValue("期初結餘保留"), iCol)) + Val(grdDataList.TextMatrix(GetValue("當月　　結餘"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末結餘保留"), iCol)) > 0 Then
'            'grdDataList.TextMatrix(iRow, iCol) = Format(Val(grdDataList.TextMatrix(GetValue("期初結餘保留"), iCol)) + Val(grdDataList.TextMatrix(GetValue("當月　　結餘"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末結餘保留"), iCol)), "0.000")
'            grdDataList.TextMatrix(iRow, iCol) = Round(Val(grdDataList.TextMatrix(GetValue("期初結餘保留"), iCol)) + Val(grdDataList.TextMatrix(GetValue("當月　　結餘"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末結餘保留"), iCol)), 3)
'            dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
''        Else
''            grdDataList.TextMatrix(iRow, iCol) = Format("0", "0.000")
''        End If
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
'      strExc(0) = GetPoint(3, Val(txtCloseDate(0)), Val(txtCloseDate(1)), txtSalesArea, txtSalesArea, , , Me.Name)
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
'                If Val(grdDataList.TextMatrix(GetValue("目　　標"), iCol)) > 0 And Val(grdDataList.TextMatrix(GetValue("達成點數"), iCol)) > 0 Then
'                    grdDataList.TextMatrix(i, iCol) = Format(100 * Val(grdDataList.TextMatrix(GetValue("達成點數"), iCol)) / Val(grdDataList.TextMatrix(GetValue("目　　標"), iCol)), "0.000") & "%"
'                Else
'                    grdDataList.TextMatrix(i, iCol) = "0%"
'                End If
'            'Modify by Amy 2016/03/01 拆位置 原GetValue("期末實績保留") To GetValue("期末結餘保留")
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
'            'Add by Amy 2016/03/01
'            Case GetValue("報出實績點數"), GetValue("報出結餘點數")
'                grdDataList.TextMatrix(i, iCol) = dblTot(i)
'                For j = 1 To grdDataList.Cols - 1
'                    grdDataList.col = j
'                    grdDataList.row = i
'                    grdDataList.CellBackColor = &HC000&
'                Next j
'            Case GetValue("轉撥實績備註"), GetValue("轉撥結餘備註")
'                grdDataList.TextMatrix(i, iCol) = ""
'            Case GetValue("新增客戶數") To GetValue("收款家數")
'                grdDataList.TextMatrix(i, iCol) = dblTot(i)
'            Case GetValue("收文案源分析")
'                grdDataList.TextMatrix(i, iCol) = Format(dblTotCase(0))
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
'Mark by Amy 2016/01/30
'    Dim stCon As String, stConST As String, stConR1 As String, stConR2 As String, stConPE As String
'    Dim stVTB0 As String, stVTB1 As String, stVTB2 As String, stVTB3 As String, stVTB4 As String
'    Dim stVTB5 As String, stVTB6 As String, stVTB7 As String, stVTB8 As String, stVTB9 As String
'    Dim iRow As Single, iCol As Single
'    Dim dblTotCase(3) As Double, dblSumR As Double
'    Dim stConCu As String, stConCP As String
'    Dim i As Integer, j As Integer
'
'   Erase dblTot
'   stCon = "": stConST = "": stConR1 = "": stConR2 = "": stConPE = "": stConCu = ""
'
'   '業務區
'   If txtSalesArea <> "" Then
'      stConST = " AND ST15='" & txtSalesArea & "'"
'      stConCu = stConCu & " AND CU12='" & txtSalesArea & "'"
'      stConCP = stConCP & " AND CP12='" & txtSalesArea & "'"
'   End If
'
'   '點數結算日
'   If txtCloseDate(0) <> "" Then
'      stCon = stCon & " AND A0205 >= " & txtCloseDate(0) & "01"
'      '期初實績保留及期初結餘保留 抓畫面條件起日當月
'      stConR1 = "  AND A0205 >= " & txtCloseDate(0) & "01 AND A0205 <= " & txtCloseDate(0) & "31"
'      stConPE = stConPE & " AND PE03(+) >= 191100+" & txtCloseDate(0)
'      stConCu = stConCu & " AND CU14>=" & TransDate(txtCloseDate(0) & "01", 2)
'      stConCP = stConCP & " AND CP05>=" & TransDate(txtCloseDate(0) & "01", 2)
'   End If
'
'   If txtCloseDate(1) <> "" Then
'      stCon = stCon & " AND A0205 <= " & txtCloseDate(1) & "31"
'      '期末實績保留及期末結餘保留 抓畫面條件止日當月
'      stConR2 = "  AND A0205 >= " & txtCloseDate(1) & "01 AND A0205 <= " & txtCloseDate(1) & "31"
'      stConPE = stConPE & " AND PE03(+) <= 191100+" & txtCloseDate(1)
'      stConCu = stConCu & " AND CU14<=" & TransDate(txtCloseDate(1) & "31", 2)
'      stConCP = stConCP & " AND CP05<=" & TransDate(txtCloseDate(1) & "31", 2)
'   End If
'
'On Error GoTo ErrHnd
'   '目標  'modify by sonia 2015/10/13 取消ST01>'60'條件,D104093822做20011  '2015/11/6 再取消ST01<'F'條件,因為2015/10的S212員工編號有掛目標
'   stVTB0 = "Select ST01,ST02,ST04,ST05,sum(PE04) PE04" & _
'      " From Staff,PerFormance" & _
'      " Where PE01(+)=ST01 And PE02(+)='TOT'" & stConST & stConPE & _
'      " Group by ST01,ST02,ST04,ST05 "
'
'   '期初實績保留:點數結算「起始」當月4191+4192貸方(期初實績保留)  'modify by sonia 2015/10/13 取消ST01>'60'條件,D104093822做20011
'   stVTB1 = "Select ax209 V10, sum(ax207) V11" & _
'      " From acc020, acc021" & _
'      " Where ax201(+) = a0201  And ax202(+) = a0202" & stConR1 & _
'      " And Exists (Select ST01 From Staff Where ST01<'F' And ST01=ax209" & stConST & ")" & _
'      " And (ax205='4191' or ax205='4192') Group by ax209 "
'   '期初結餘保留:點數結算「起始」當月4194貸方(期初結餘保留)  'modify by sonia 2015/10/13 取消ST01>'60'條件,D104093822做20011
'    stVTB2 = "Select ax209 V20, sum(ax207) V21" & _
'      " From acc020, acc021" & _
'      " Where ax201(+) = a0201  And ax202(+) = a0202" & stConR1 & _
'      " And Exists (Select ST01 From Staff Where ST01<'F' And ST01=ax209" & stConST & ")" & _
'      " And ax205='4194' Group by ax209 "
'
'   '當月　　實績  'modify by sonia 2015/10/13 取消ST01>'60'條件,D104093822做20011
'    stVTB3 = "Select ax209 V30, sum(ax207-ax206) V31" & _
'      " From acc020, acc021" & _
'      " Where ax201(+) = a0201  And ax202(+) = a0202" & stCon & _
'      " And Exists (Select ST01 From Staff Where ST01<'F' And ST01=ax209" & stConST & ")" & _
'      " And (SubStr(ax205, 1, 2) = '41' Or (ax205='7121' And ax209 is not null)) And Not( ax205='4191' or ax205='4192' or ax205='4194')" & _
'      " And (ax213 Is Null or InStr(ax213||' ','結餘')=0) Group by ax209 "
'   '當月　　結餘  'modify by sonia 2015/10/13 取消ST01>'60'條件,D104093822做20011
'    stVTB4 = "Select ax209 V40, sum(ax207-ax206) V41" & _
'      " From acc020, acc021" & _
'      " Where ax201(+) = a0201  And ax202(+) = a0202" & stCon & _
'      " And Exists (Select ST01 From Staff Where ST01<'F' And ST01=ax209" & stConST & ")" & _
'      " And (SubStr(ax205, 1, 2) = '41' Or (ax205='7121'And ax209 is not null)) And Not( ax205='4191' or ax205='4192' or ax205='4194')" & _
'      " And InStr(ax213||' ','結餘')>0 Group by ax209 "
'
'   '期末實績保留:點數結算「迄月」當月4191+4192借方(期末實績保留)  'modify by sonia 2015/10/13 取消ST01>'60'條件,D104093822做20011
'   stVTB5 = "Select ax209 V50, Sum(ax206) V51" & _
'      " From acc020, acc021" & _
'      " Where ax201(+) = a0201  And ax202(+) = a0202" & stConR2 & _
'      " And exists (Select ST01 From Staff Where ST01<'F' And ST01=ax209" & stConST & ")" & _
'      " And (ax205='4191' or ax205='4192') Group by ax209 "
'   '期末結餘保留:點數結算「迄月」當月4194借方(期末結餘保留)  'modify by sonia 2015/10/13 取消ST01>'60'條件,D104093822做20011
'   stVTB6 = "Select ax209 V60, Sum(ax206) V61" & _
'      " From acc020, acc021" & _
'      " Where ax201(+) = a0201  And ax202(+) = a0202" & stConR2 & _
'      " And exists (Select ST01 From Staff Where ST01<'F' And ST01=ax209" & stConST & ")" & _
'      " And ax205='4194' Group by ax209 "
'
'   '每月新增客戶數
'   stVTB7 = "Select cu13 V70,count(*) V71 " & _
'      "From customer Where cu02='0'" & stConCu & " group by cu13"
'   '收款家數  'modify by sonia 2015/10/13 取消ST01>'60'條件,D104093822做20011
'   'modify by sonia 2016/1/22 +412101,413101
'   stVTB8 = "Select ax209 V80, count( distinct substr(ax208,1,6) ) V81" & _
'      " From acc020, acc021" & _
'      " Where ax201(+) = a0201  And ax202(+) = a0202" & stCon & _
'      " And exists (Select ST01 From Staff Where ST01<'F' And ST01=ax209" & stConST & ")" & _
'      " And substr(ax205, 1, 2) = '41' And ax207>0 And ax208 is not null" & _
'      " And not( ax205='4191' or ax205='4192' or ((ax205='410103' or ax205='411103' or ax205='4121' or ax205='412101' or ax205='4131' or ax205='413101')" & _
'      " And instr(ax213||' ','結餘')>0)) Group by ax209"
'
'   '收文案件來源分析
'   stVTB9 = "Select CP13 V90" & _
'      ", SUM(DECODE(CP01,'P',1,'PS',1,'CFP',1,'CPS',1,0)) V91" & _
'      ", SUM(DECODE(CP01,'T',1,'TF',1,'CFT',1,'TC',1,0)) V92" & _
'      ", SUM(DECODE(CP01,'L',1,'LA',1,0)) V93" & _
'      ", SUM(DECODE(CP01,'CFC',1,'CFL',1,0)) V94" & _
'      " From CaseProgress Where cp09<'B' And cp13 is not null" & stConCP & _
'      " And (  (cp01 in ('P','PS','CFP','CPS') )  or (cp01 in ('T','TF','CFT','TC') )" & _
'      " or (cp01 in ('CFC','CFL')) or (cp01 in ('L','LA'))) Group by CP13"
'
'   strSql = "Select ST02,Nvl(PE04,0) PE04,0,0,(NVL(V11,0))/1000 C1,NVL(V21,0)/1000 C2,NVL(V31,0)/1000 C3,NVL(V41,0)/1000 C4,0,0,NVL(V51,0)/1000 C5,NVL(V61,0)/1000 C6" & _
'      ", V71 C7, V81 C8, NVL(V91,0) C9, NVL(V92,0) C10, NVL(V93,0) C11, NVL(V94,0) C12" & _
'      " From (" & stVTB0 & ") V0,(" & stVTB1 & ") V1,(" & stVTB2 & ") V2,(" & stVTB3 & ") V3,(" & stVTB4 & ") V4,(" & stVTB5 & ") V5,(" & stVTB6 & ") V6" & _
'      ",(" & stVTB7 & ") V7,(" & stVTB8 & ") V8,(" & stVTB9 & ") V9" & _
'      " Where V10(+)=ST01 And V20(+)=ST01 And V30(+)=ST01 And V40(+)=ST01 And V50(+)=ST01 And V60(+)=ST01" & _
'      " And V70(+)=ST01 And V80(+)=ST01 And V90(+)=ST01 And (PE04>0 or V11>0 or V21>0 or V31>0 or V41>0) Order by Decode(ST05,'SM',1,'76',1,2),st01"
'
'   intI = 1
'   Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      With AdoRecordSet3
'      iCol = 3
'      Erase dblTot
'      Do While Not .EOF
'         dblSumR = 0
'         grdDataList.Cols = iCol + 1
'         '智權人員
'         iRow = GetValue("智權人員")
'         grdDataList.TextMatrix(iRow, iCol) = "" & .Fields("ST02")
'         grdDataList.CellAlignment = flexAlignCenterCenter
'         '目　　標
'         iRow = GetValue("目　　標")
'         grdDataList.TextMatrix(iRow, iCol) = "" & .Fields("PE04")
'         grdDataList.CellAlignment = flexAlignRightCenter
'         '期初實績保留
'         iRow = GetValue("期初實績保留")
'         grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C1")), "0.000")
'         dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'         dblSumR = dblSumR + Val(grdDataList.TextMatrix(iRow, iCol))
'         '期初結餘保留
'         iRow = GetValue("期初結餘保留")
'         grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C2")), "0.000")
'         dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'         dblSumR = dblSumR + Val(grdDataList.TextMatrix(iRow, iCol))
'         '當月　　實績
'         iRow = GetValue("當月　　實績")
'          grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C3")), "0.000")
'         dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'         dblSumR = dblSumR + Val(grdDataList.TextMatrix(iRow, iCol))
'         '當月　　結餘
'         iRow = GetValue("當月　　結餘")
'         grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C4")), "0.000")
'         dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'         dblSumR = dblSumR + Val(grdDataList.TextMatrix(iRow, iCol))
'         '期末實績保留
'         iRow = GetValue("期末實績保留")
'         grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C5")), "0.000")
'         dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'         dblSumR = dblSumR - Val(grdDataList.TextMatrix(iRow, iCol))
'         '期末結餘保留
'         iRow = GetValue("期末結餘保留")
'         grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C6")), "0.000")
'         dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'         dblSumR = dblSumR - Val(grdDataList.TextMatrix(iRow, iCol))
'         '新增客戶數
'         iRow = GetValue("新增客戶數")
'         grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C7")))
'         dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'         '收款家數
'         iRow = GetValue("收款家數")
'         grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C8")))
'         dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'         '收文案件來源分析
'         'P
'         iRow = GetValue("收文案源分析")
'         grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C9")))
'         dblTotCase(0) = dblTotCase(0) + Val(grdDataList.TextMatrix(iRow, iCol))
'         'T
'         iRow = GetValue("收文案源分析") + 1
'         grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C10")))
'         dblTotCase(1) = dblTotCase(1) + Val(grdDataList.TextMatrix(iRow, iCol))
'         'L
'         iRow = GetValue("收文案源分析") + 2
'         grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C11")))
'         dblTotCase(2) = dblTotCase(2) + Val(grdDataList.TextMatrix(iRow, iCol))
'         'C
'         iRow = GetValue("收文案源分析") + 3
'         grdDataList.TextMatrix(iRow, iCol) = Format(Val("" & .Fields("C12")))
'         dblTotCase(3) = dblTotCase(3) + Val(grdDataList.TextMatrix(iRow, iCol))
'
'         '達成點數
'         iRow = GetValue("達成點數")
'         strExc(0) = Val(grdDataList.TextMatrix(GetValue("期初實績保留"), iCol)) + Val(grdDataList.TextMatrix(GetValue("期初結餘保留"), iCol))
'         strExc(0) = Val(strExc(0)) + Val(grdDataList.TextMatrix(GetValue("當月　　實績"), iCol)) + Val(grdDataList.TextMatrix(GetValue("當月　　結餘"), iCol))
'         strExc(0) = Val(strExc(0)) - Val(grdDataList.TextMatrix(GetValue("期末實績保留"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末結餘保留"), iCol))
'         grdDataList.TextMatrix(iRow, iCol) = Format(strExc(0), "0.000")
'         dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'
'         '達成率
'         iRow = GetValue("達 成 率")
'         If Val(grdDataList.TextMatrix(GetValue("目　　標"), iCol)) > 0 Then
'            grdDataList.TextMatrix(iRow, iCol) = Format(100 * Val(grdDataList.TextMatrix(GetValue("達成點數"), iCol)) / Val(grdDataList.TextMatrix(GetValue("目　　標"), iCol)), "0.0") & "%"
'         Else
'            grdDataList.TextMatrix(iRow, iCol) = "0.0%"
'         End If
'
'         '實績保留動用
'         iRow = GetValue("實績保留動用")
'         'Modify by Amy 2015/11/09 +if 小於等於0顯示0(公式)
'         If Val(grdDataList.TextMatrix(GetValue("期初實績保留"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末實績保留"), iCol)) > 0 Then
'            grdDataList.TextMatrix(iRow, iCol) = Format(Val(grdDataList.TextMatrix(GetValue("期初實績保留"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末實績保留"), iCol)), "0.000")
'            dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'         Else
'            grdDataList.TextMatrix(iRow, iCol) = Format("0", "0.000")
'         End If
'
'         '結餘保留動用
'         iRow = GetValue("結餘保留動用")
'         'Modify by Amy 2015/11/09 +if 小於等於0顯示0(公式)
'         If Val(grdDataList.TextMatrix(GetValue("期初結餘保留"), iCol)) + Val(grdDataList.TextMatrix(GetValue("當月　　結餘"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末結餘保留"), iCol)) > 0 Then
'            grdDataList.TextMatrix(iRow, iCol) = Format(Val(grdDataList.TextMatrix(GetValue("期初結餘保留"), iCol)) + Val(grdDataList.TextMatrix(GetValue("當月　　結餘"), iCol)) - Val(grdDataList.TextMatrix(GetValue("期末結餘保留"), iCol)), "0.000")
'            dblTot(iRow) = dblTot(iRow) + Val(grdDataList.TextMatrix(iRow, iCol))
'        Else
'            grdDataList.TextMatrix(iRow, iCol) = Format("0", "0.000")
'        End If
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
'end 2015/04/20

Private Function doQuery_Old2() As Boolean
'Mark by Amy 2015/04/20
'   Dim stCon As String, stConST As String, stConResv As String, stConPE As String
'   Dim stVTB0 As String, stVTB1 As String, stVTB2 As String, stVTB3 As String, stVTB4 As String
'   Dim iRow As Single, iCol As Single
'   Dim dblTot(1 To 11) As Double
'   'Add by Morgan 2006/1/13
'   Dim stVTB5 As String, stVTB6 As String, stVTB7 As String
'   Dim stConCu As String, stConCP As String
'
'   stCon = "": stConST = "": stConResv = "": stConPE = "": stConCu = ""
'
'   '業務區
'   If txtSalesArea <> "" Then
'      stConST = " AND ST15='" & txtSalesArea & "'"
'      'Add by Morgan 2006/1/13
'      stConCu = stConCu & " AND CU12='" & txtSalesArea & "'"
'      stConCP = stConCP & " AND CP12='" & txtSalesArea & "'"
'   End If
'
'   '點數結算日
'   If txtCloseDate(0) <> "" Then
'      stCon = stCon & " AND A0205 >= " & txtCloseDate(0) & "01"
'      stConPE = stConPE & " AND PE03(+) >= 191100+" & txtCloseDate(0)
'      'Add by Morgan 2006/1/13
'      stConCu = stConCu & " AND CU14>=" & TransDate(txtCloseDate(0) & "01", 2)
'      stConCP = stConCP & " AND CP05>=" & TransDate(txtCloseDate(0) & "01", 2)
'   End If
'   If txtCloseDate(1) <> "" Then
'      stCon = stCon & " AND A0205 <= " & txtCloseDate(1) & "31"
'      '保留只抓迄月
'      stConResv = "  AND A0205 >= " & txtCloseDate(1) & "01 AND A0205 <= " & txtCloseDate(1) & "31"
'      stConPE = stConPE & " AND PE03(+) <= 191100+" & txtCloseDate(1)
'      'Add by Morgan 2006/1/13
'      stConCu = stConCu & " AND CU14<=" & TransDate(txtCloseDate(1) & "31", 2)
'      stConCP = stConCP & " AND CP05<=" & TransDate(txtCloseDate(1) & "31", 2)
'   End If
'
'On Error GoTo ErrHnd
'   '目標
'   stVTB0 = "select ST01,ST02,ST04,ST05,sum(PE04) PE04" & _
'      " from staff,PERFORMANCE" & _
'      " where ST01>'60' AND ST01<'F' AND PE01(+)=ST01 AND PE02(+)='TOT'" & stConST & stConPE & _
'      " GROUP BY ST01,ST02,ST04,ST05"
'
'   '達成點數
'   '2014/1/21 modify by sonia 取消a0201='1'條件
'   'stVTB1 = "select ax209 V10, sum(ax207-ax206) V11" & _
'      " From acc020, acc021" & _
'      " where a0201='1'" & stCon & _
'      " and ax201(+) = a0201  and ax202(+) = a0202" & _
'      " and exists (SELECT ST01 FROM STAFF WHERE ST01>'60' AND ST01<'F' and ST01=ax209" & stConST & ")" & _
'      " and (substr(ax205, 1, 2) = '41' Or substr(ax205,1,2)='71')" & _
'      " group by ax209"
'   stVTB1 = "select ax209 V10, sum(ax207-ax206) V11" & _
'      " From acc020, acc021" & _
'      " where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
'      " and exists (SELECT ST01 FROM STAFF WHERE ST01>'60' AND ST01<'F' and ST01=ax209" & stConST & ")" & _
'      " and (substr(ax205, 1, 2) = '41' Or substr(ax205,1,2)='71')" & _
'      " group by ax209"
'   '2014/1/21 end
'
'   '當月達成
'   '2014/1/21 modify by sonia 取消a0201='1'條件
'   'stVTB2 = "select ax209 V20, sum(ax207-ax206) V21" & _
'      " From acc020, acc021" & _
'      " where a0201='1'" & stCon & _
'      " and ax201(+) = a0201  and ax202(+) = a0202" & _
'      " and exists (SELECT ST01 FROM STAFF WHERE ST01>'60' AND ST01<'F' and ST01=ax209" & stConST & ")" & _
'      " and (substr(ax205, 1, 2) = '41' Or substr(ax205,1,2)='71')" & _
'      " and not( ax205='4191' or ax205='4192' or ((ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131')" & _
'      " and instr(ax213||' ','結餘')>0)) group by ax209"
'   stVTB2 = "select ax209 V20, sum(ax207-ax206) V21" & _
'      " From acc020, acc021" & _
'      " where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
'      " and exists (SELECT ST01 FROM STAFF WHERE ST01>'60' AND ST01<'F' and ST01=ax209" & stConST & ")" & _
'      " and (substr(ax205, 1, 2) = '41' Or substr(ax205,1,2)='71')" & _
'      " and not( ax205='4191' or ax205='4192' or ((ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131')" & _
'      " and instr(ax213||' ','結餘')>0)) group by ax209"
'   '2014/1/21 end
'
'   '結餘
'   '2014/1/21 modify by sonia 取消a0201='1'條件
'   'stVTB3 = "select ax209 V30, sum(ax207-ax206) V31" & _
'      " From acc020, acc021" & _
'      " where a0201='1'" & stCon & _
'      " and ax201(+) = a0201  and ax202(+) = a0202" & _
'      " and exists (SELECT ST01 FROM STAFF WHERE ST01>'60' AND ST01<'F' and ST01=ax209" & stConST & ")" & _
'      " and (ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131')" & _
'      " and instr(ax213||' ','結餘')>0 group by ax209 "
'   stVTB3 = "select ax209 V30, sum(ax207-ax206) V31" & _
'      " From acc020, acc021" & _
'      " where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
'      " and exists (SELECT ST01 FROM STAFF WHERE ST01>'60' AND ST01<'F' and ST01=ax209" & stConST & ")" & _
'      " and (ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131')" & _
'      " and instr(ax213||' ','結餘')>0 group by ax209 "
'   '2014/1/21 end
'
'   '保留
'   '2014/1/21 modify by sonia 取消a0201='1'條件
'   'stVTB4 = "select ax209 V40, sum(ax206) V41" & _
'      " From acc020, acc021" & _
'      " where a0201='1'" & stConResv & _
'      " and ax201(+) = a0201  and ax202(+) = a0202" & _
'      " and exists (SELECT ST01 FROM STAFF WHERE ST01>'60' AND ST01<'F' and ST01=ax209" & stConST & ")" & _
'      " and (ax205='4191' or ax205='4192') group by ax209"
'   stVTB4 = "select ax209 V40, sum(ax206) V41" & _
'      " From acc020, acc021" & _
'      " where ax201(+) = a0201  and ax202(+) = a0202" & stConResv & _
'      " and exists (SELECT ST01 FROM STAFF WHERE ST01>'60' AND ST01<'F' and ST01=ax209" & stConST & ")" & _
'      " and (ax205='4191' or ax205='4192') group by ax209"
'   '2014/1/21 end
'
'   'Add by Morgan 2006/1/13
'   '每月新增客戶數
'   stVTB5 = "select cu13 V50,count(*) V51" & _
'      " from customer where cu02='0'" & stConCu & " group by cu13"
'   '收款家數
'   '2014/1/21 modify by sonia 取消a0201='1'條件
'   'stVTB6 = "select ax209 V60, count( distinct substr(ax208,1,6) ) V61" & _
'      " From acc020, acc021" & _
'      " where a0201='1'" & stCon & _
'      " and ax201(+) = a0201  and ax202(+) = a0202" & _
'      " and exists (SELECT ST01 FROM STAFF WHERE ST01>'60' AND ST01<'F' and ST01=ax209" & stConST & ")" & _
'      " and substr(ax205, 1, 2) = '41' and ax207>0 and ax208 is not null" & _
'      " and not( ax205='4191' or ax205='4192' or ((ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131')" & _
'      " and instr(ax213||' ','結餘')>0)) group by ax209"
'   stVTB6 = "select ax209 V60, count( distinct substr(ax208,1,6) ) V61" & _
'      " From acc020, acc021" & _
'      " where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
'      " and exists (SELECT ST01 FROM STAFF WHERE ST01>'60' AND ST01<'F' and ST01=ax209" & stConST & ")" & _
'      " and substr(ax205, 1, 2) = '41' and ax207>0 and ax208 is not null" & _
'      " and not( ax205='4191' or ax205='4192' or ((ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131')" & _
'      " and instr(ax213||' ','結餘')>0)) group by ax209"
'   '2014/1/21 end
'
'   '收文案件來源分析
'   stVTB7 = "SELECT CP13 V70" & _
'      ", SUM(DECODE(CP01,'P',1,'PS',1,'CFP',1,'CPS',1,0)) V71" & _
'      ", SUM(DECODE(CP01,'T',1,'TF',1,'CFT',1,'TC',1,0)) V72" & _
'      ", SUM(DECODE(CP01,'L',1,'LA',1,0)) V73" & _
'      ", SUM(DECODE(CP01,'CFC',1,'CFL',1,0)) V74" & _
'      " From caseprogress where cp09<'B' AND cp13 is not null" & stConCP & _
'      " and (  (cp01 in ('P','PS','CFP','CPS') )  or (cp01 in ('T','TF','CFT','TC') )" & _
'      " or (cp01 in ('CFC','CFL')) or (cp01 in ('L','LA'))) GROUP BY CP13"
'
'   strSql = "select ST02,PE04,(NVL(V11,0))/1000 C1,NVL(V21,0)/1000 C2,NVL(V31,0)/1000 C3,NVL(V41,0)/1000 C4" & _
'      ", V51 C5, V61 C6, NVL(V71,0) C7, NVL(V72,0) C8, NVL(V73,0) C9, NVL(V74,0) C10" & _
'      " from (" & stVTB0 & ") V0,(" & stVTB1 & ") V1,(" & stVTB2 & ") V2,(" & stVTB3 & ") V3" & _
'      ",(" & stVTB4 & ") V4,(" & stVTB5 & ") V5,(" & stVTB6 & ") V6,(" & stVTB7 & ") V7" & _
'      " where V10(+)=ST01 AND V20(+)=ST01 AND V30(+)=ST01 AND V40(+)=ST01 AND V50(+)=ST01 AND V60(+)=ST01" & _
'      " AND V70(+)=ST01 AND (ST04='1' OR V11>0) ORDER BY DECODE(ST05,'SM',1,'76',1,2),st01"
'
'   intI = 1
'   Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      With AdoRecordSet3
'      iRow = 0: iCol = 3
'      Erase dblTot
'      Do While Not .EOF
'         grdDataList.Cols = iCol + 1
'         '智權人員
'         grdDataList.TextMatrix(0, iCol) = "" & .Fields("ST02")
'         grdDataList.row = 0: grdDataList.col = iCol
'         grdDataList.CellAlignment = flexAlignCenterCenter
'         '目　　標
'         grdDataList.TextMatrix(1, iCol) = "" & .Fields("PE04")
'         dblTot(1) = dblTot(1) + Val(grdDataList.TextMatrix(1, iCol))
'         '達成點數
'         grdDataList.TextMatrix(2, iCol) = Format(Val("" & .Fields("C1")), "0.000")
'         dblTot(2) = dblTot(2) + Val(grdDataList.TextMatrix(2, iCol))
'         '當月達成
'         grdDataList.TextMatrix(4, iCol) = Format(Val("" & .Fields("C2")), "0.0")
'         dblTot(3) = dblTot(3) + Val(grdDataList.TextMatrix(4, iCol))
'         '結餘點數
'         grdDataList.TextMatrix(5, iCol) = Format(Val("" & .Fields("C3")), "0.00")
'         dblTot(4) = dblTot(4) + Val(grdDataList.TextMatrix(5, iCol))
'         '保留點數
'         grdDataList.TextMatrix(6, iCol) = Format(Val("" & .Fields("C4")))
'         dblTot(5) = dblTot(5) + Val(grdDataList.TextMatrix(6, iCol))
'         '新增客戶數
'         grdDataList.TextMatrix(7, iCol) = Format(Val("" & .Fields("C5")))
'         dblTot(6) = dblTot(6) + Val(grdDataList.TextMatrix(7, iCol))
'         '收款家數
'         grdDataList.TextMatrix(8, iCol) = Format(Val("" & .Fields("C6")))
'         dblTot(7) = dblTot(7) + Val(grdDataList.TextMatrix(8, iCol))
'         '收文案件來源分析
'         'P
'         grdDataList.TextMatrix(9, iCol) = Format(Val("" & .Fields("C7")))
'         dblTot(8) = dblTot(8) + Val(grdDataList.TextMatrix(9, iCol))
'         'T
'         grdDataList.TextMatrix(10, iCol) = Format(Val("" & .Fields("C8")))
'         dblTot(9) = dblTot(9) + Val(grdDataList.TextMatrix(10, iCol))
'         'L
'         grdDataList.TextMatrix(11, iCol) = Format(Val("" & .Fields("C9")))
'         dblTot(10) = dblTot(10) + Val(grdDataList.TextMatrix(11, iCol))
'         'C
'         grdDataList.TextMatrix(12, iCol) = Format(Val("" & .Fields("C10")))
'         dblTot(11) = dblTot(11) + Val(grdDataList.TextMatrix(12, iCol))
'
'         '達成率
'         If Val(grdDataList.TextMatrix(1, iCol)) > 0 Then
'            grdDataList.TextMatrix(3, iCol) = Format(100 * Val(grdDataList.TextMatrix(2, iCol)) / Val(grdDataList.TextMatrix(1, iCol)), "0.0") & "%"
'         End If
'         .MoveNext
'         iCol = iCol + 1
'      Loop
'      End With
'
'      grdDataList.Cols = iCol + 1
'      grdDataList.TextMatrix(0, iCol) = "合　　計"
'      'Add by Morgan 2007/9/10 總目標需抓該區所有(因為有可能點數掛該區代碼如20021)
'      strExc(0) = "select nvl(sum(PE04),0) PE04 from staff,PERFORMANCE" & _
'         " where PE01(+)=ST01 AND PE02(+)='TOT'" & stConST & stConPE
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         dblTot(1) = RsTemp.Fields(0)
'      End If
'      'END 2007/9/10
'      grdDataList.TextMatrix(1, iCol) = dblTot(1)
'      grdDataList.TextMatrix(2, iCol) = Format(dblTot(2), "0.000")
'      grdDataList.TextMatrix(4, iCol) = Format(dblTot(3), "0.0")
'      grdDataList.TextMatrix(5, iCol) = Format(dblTot(4), "0.00")
'      grdDataList.TextMatrix(6, iCol) = Format(dblTot(5))
'      grdDataList.TextMatrix(7, iCol) = Format(dblTot(6))
'      grdDataList.TextMatrix(8, iCol) = Format(dblTot(7))
'      grdDataList.TextMatrix(9, iCol) = Format(dblTot(8))
'      grdDataList.TextMatrix(10, iCol) = Format(dblTot(9))
'      grdDataList.TextMatrix(11, iCol) = Format(dblTot(10))
'      grdDataList.TextMatrix(12, iCol) = Format(dblTot(11))
'
'      '達成率
'      If Val(grdDataList.TextMatrix(1, iCol)) > 0 Then
'         grdDataList.TextMatrix(3, iCol) = Format(100 * Val(grdDataList.TextMatrix(2, iCol)) / Val(grdDataList.TextMatrix(1, iCol)), "0.0") & "%"
'      End If
'   Else
'      MsgBox "無符合資料！", vbInformation
'   End If
'
'
'   doQuery = True
'
'ErrHnd:
'   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
'end 2015/04/20
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm210113 = Nothing
End Sub

Private Sub txtCloseDate_GotFocus(Index As Integer)
   TextInverse txtCloseDate(Index)
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtCloseDate(Index).IMEMode = 2
   CloseIme
End Sub

Private Sub txtCloseDate_Validate(Index As Integer, Cancel As Boolean)
   'Add by Amy 2016/01/30
   If txtCloseDate(Index) = MsgText(601) Then Exit Sub
   
   If txtCloseDate(Index) <> "" Then
      If ChkDate(txtCloseDate(Index) & "01") = False Then
         txtCloseDate_GotFocus Index
         Cancel = True
         Exit Sub
      End If
      'Add by Amy 2016/01/30 不可查舊資料
      If Index = 0 And Val(txtCloseDate(0)) < Val(業績輸入啟用年月) Then
        MsgBox Label2 & "不可查詢105年1月前的資料"
        txtCloseDate_GotFocus Index
        Cancel = True
        Exit Sub
      End If
      'Add by Amy 2015/04/27
      If Index = 1 And Val(txtCloseDate(0)) > Val(txtCloseDate(1)) Then
         MsgBox Label2 & "起始年月不可大於截止年月"
         txtCloseDate_GotFocus Index
         Cancel = True
         Exit Sub
      End If
      'Mark by Amy 2016/01/30 原要判斷此,起迄年月期間不可跨「業績輸入啟用年月」,但抓目標時業務區不知如何抓舊資料,故先檔不可查舊資料
'      If Index = 1 And (Val(txtCloseDate(0)) < Val(業績輸入啟用年月) And Val(業績輸入啟用年月) < Val(txtCloseDate(1))) _
'        Or (Val(txtCloseDate(0)) < Val(業績輸入啟用年月) And Val(業績輸入啟用年月) = Val(txtCloseDate(1))) Then
'        MsgBox Label2 & "不可跨業績輸入啟用年月"
'        txtCloseDate_GotFocus Index
'        Cancel = True
'      End If
   End If
End Sub

Private Sub txtSalesArea_Change()
   If txtSalesArea = "" Then
      lblSalesArea = ""
   Else
      lblSalesArea = A0902Query(txtSalesArea)
   End If
End Sub

Private Sub txtSalesArea_GotFocus()
   TextInverse txtSalesArea
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtSalesArea.IMEMode = 2
   CloseIme
End Sub

Private Sub txtSalesArea_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add by Amy 2015/04/20
Private Function GetValue(pRowN As String) As Integer
   Dim jj As Integer
 
    For jj = 1 To UBound(strRowN)
       If Replace(UCase(strRowN(jj)), "　　", "") = UCase(pRowN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function

'Add by Amy 2024/04/19 取得各假別時數
Private Function GetAbsenceData(ByVal stCon As String, stStartDate As String, stEndDate As String) As String
   Dim stTmp(1) As String, i As Integer
   
   GetAbsenceData = ""
   If PUB_GetAbsenceHour(stCon, stStartDate, stEndDate, dblHour(), dblCnt()) = True Then
      For i = LBound(dblHour) To UBound(dblHour)
         If dblHour(i) <> 0 Then
            Select Case i
               Case 1
                  stTmp(1) = "忘打卡"
               Case 2
                  stTmp(1) = "遲到"
               Case 3
                  stTmp(1) = "曠職"
               Case 4
                  stTmp(1) = "出差"
               Case 5
                  stTmp(1) = "事假"
               Case 6
                  stTmp(1) = "病假"
               Case 7
                  stTmp(1) = "公假"
               Case 8
                  stTmp(1) = "特別假"
               Case 9
                  stTmp(1) = "婚假"
               Case 10
                  stTmp(1) = "產假"
               Case 11
                  stTmp(1) = "流產假"
               Case 12
                  stTmp(1) = "喪假"
               Case 13
                  stTmp(1) = "公傷假"
               Case 14
                  stTmp(1) = "補休"
               Case 15
                  stTmp(1) = "其他"
               Case 16
                  stTmp(1) = "加班"
               Case 17
                  stTmp(1) = "扣年終產假"
               Case 18
                  stTmp(1) = "扣年終流產假"
               Case 19
                  stTmp(1) = "陪產假"
               Case 20
                  stTmp(1) = "生理假"
               Case 21
                  stTmp(1) = "產檢假"
               Case 22
                  stTmp(1) = "家庭照顧假"
               Case 23
                  stTmp(1) = "健檢假"
               Case 24
                  stTmp(1) = "防疫照顧假"
            End Select
            'Modify By Sindy 2025/11/17 + & IIf(i = 1 Or i = 2, "次", IIf(i = 3, "分", ""))
            stTmp(0) = stTmp(0) & ";" & stTmp(1) & ":" & dblHour(i) & IIf(i = 1 Or i = 2, "次", IIf(i = 3, "分", ""))
         End If
      Next i
      If stTmp(0) <> MsgText(601) Then
         GetAbsenceData = Mid(stTmp(0), 2)
      End If
   End If
End Function

'收文/發文件數(參考收發文量語法frm100105_2.StrMenu及frm100105_2.DoTemp修改)
'stSys:系統別 P or T / stStartD:起始日 / stStartD:截止日 / stDept:部門
Private Sub NewCaseSituation(ByVal stSys As String, ByVal stStartD As String, ByVal stEndD As String, stDept As String)
   Dim RsQ As New ADODB.Recordset, intQ As Integer, ii As Integer, strQ As String, strWhere As String
   Dim strTB As String, strBase As String, strField As String, strCP31 As String, strTemp As String
   Dim mTable As Word.Table, intRow As Integer, strState As String
   
On Error GoTo ErrHand
   If stSys = "P" Then
      strTB = ",Patent "
      strField = ",pa26,pa75"
      strWhere = "And cp01=pa01(+) And cp02=pa02(+) And cp03=pa03(+) And cp04=pa04(+) " & _
                            "And pa09='000' And cp10>='101' And cp10<='103' "
      strState = "1.3"
   Else
      strTB = ",TradeMark "
      strField = ",tm23,tm44"
      strWhere = "And cp01=tm01(+) And cp02=tm02(+) And cp03=tm03(+) And cp04=tm04(+) " & _
                            "And tm10='000' And cp10='101' "
      strState = "1.1"
   End If
   
   strWhere = strWhere & "And CP01 IN ('" & stSys & "') "
   strCP31 = ",Decode(CP31,'Y',Decode(InStr('CFT,FCT,T,TF',CP01),0,Decode(InStr('CFP,FCP,P',CP01),0,'',Decode(InStr('" & NewCasePtyList & ",801,803',CP10),0,'Y','')),Decode(InStr('101,308,601,603,605,618',CP10),0,'Y','')),'') mCP10 "
   strBase = "Select cp05,cp01,cp26,cp21,cp12,cp13,cp10,CP02||CP03||CP04,cp09,ST15" & strField & strCP31 & _
                     "From CaseProgress,Staff" & strTB & _
                     "Where CP05>=" & stStartD & " And CP05<=" & stEndD & " And Nvl(st15,cp12)='" & stDept & "' And cp26 IS NULL  And cp21 IS NULL " & _
                     "And CP09< 'B'  And CP01||CP02<>'TT999999' And CP13=ST01(+) " & strWhere
   '收文量
   strQ = "Insert Into R210113 (ID,R01,R04,R05) " & " Select '" & strUserNum & "',CP13,'" & strState & "',Nvl(Count(*),0) From (" & strBase & "And CP159=0 ) Group by CP13 "
   adoTaie.Execute strQ
   '發文量
   strQ = "Insert Into R210113 (ID,R01,R04,R05) " & " Select '" & strUserNum & "',CP13,'" & Val(strState) + 0.1 & "',Nvl(Count(*),0) From (" & Replace(UCase(strBase), "CP05", "CP27") & ") Group by CP13 "
   adoTaie.Execute strQ
   
   'CNT 收文量/PCNT 發文量 ex:11312月 劉峻綱(B2029)有SalesPoint資料,但無T及P新申請案資料
   strQ = "Select Distinct ST02,Nvl(CNT,0) as CNT,Nvl(PCNT,0) as PCNT,R01,Decode(ST05,'SM',1,'76',1,2) as Sort " & _
               "From R210113,Staff " & _
                ",(Select R01 as CNT_P,R05 as CNT From R210113 Where ID='" & strUserNum & "'  And R04='" & strState & "' ) " & _
                ",(Select R01 as PCNT_P,R05 as PCNT From R210113 Where ID='" & strUserNum & "'  And R04='" & Val(strState) + 0.1 & "' ) " & _
                "Where ID='" & strUserNum & "' And SubStr(R04,1,1)='1' And Substr(R01,1,1)<>'S' " & _
                "And R01=ST01(+) And R01=CNT_P(+) And R01=PCNT_P(+) " & _
                "Order by Sort,R01"
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, strQ)
   With g_WordAp.Application
      intRow = 2
      If intQ = 1 Then intRow = RsQ.RecordCount + 1
      Set mTable = .Selection.Tables.add(Range:=.Selection.Range, NumRows:=intRow, NumColumns:=4)
      'Word 2007 預設沒有框線需另外指定
      With mTable
        .Borders(wdBorderLeft).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
        .Borders(wdBorderRight).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
        .Borders(wdBorderTop).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
        .Borders(wdBorderBottom).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
        .Borders(wdBorderVertical).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
        .Borders(wdBorderHorizontal).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
      End With
      .Selection.SelectRow
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
      .Selection.Font.Size = 12
      '設定欄位
      For ii = 1 To 4
         strTemp = ""
         Select Case ii
            Case 2
               strTemp = "收文件數"
            Case 3
               strTemp = "發文件數"
            Case 4
               strTemp = "說明"
         End Select
         .Selection.TypeText Text:=strTemp
         .Selection.Font.Bold = True
         .Selection.ParagraphFormat.Alignment = wdAlignRowCenter
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
      Next ii
      If intQ = 1 Then
         RsQ.MoveFirst
         Do While RsQ.EOF = False
            For ii = 0 To RsQ.Fields.Count - 2
               strTemp = ""
               strExc(2) = UCase(RsQ.Fields(ii).Name)
               If strExc(2) <> "R01" Then
                  strTemp = RsQ.Fields(strExc(2))
               End If
               .Selection.TypeText Text:=strTemp
               .Selection.ParagraphFormat.Alignment = wdAlignRowCenter
               .Selection.MoveRight Unit:=wdCharacter, Count:=1
            Next ii
            RsQ.MoveNext
         Loop
      End If
      mTable.Select
      .Selection.EndOf Unit:=wdColumn, Extend:=wdExtend
      .Selection.MoveDown Unit:=wdLine, Count:=1
   End With
   Exit Sub
   
ErrHand:
   If Err.Number <> 0 Then
      MsgBox Err.Description & vbCrLf & "請洽電腦中心", vbCritical
   End If
End Sub





