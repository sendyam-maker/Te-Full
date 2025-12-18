VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm210135 
   BorderStyle     =   1  '單線固定
   Caption         =   "業績年度統計表"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8955
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList2 
      Height          =   0
      Left            =   2820
      TabIndex        =   9
      Top             =   60
      Width           =   0
      _ExtentX        =   0
      _ExtentY        =   0
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   3
      FixedCols       =   0
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   2190
      MaxLength       =   2
      TabIndex        =   2
      Top             =   360
      Width           =   510
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1530
      MaxLength       =   2
      TabIndex        =   1
      Top             =   360
      Width           =   510
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Excel(&E)"
      Height          =   400
      Left            =   7110
      TabIndex        =   5
      Top             =   120
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1530
      MaxLength       =   3
      TabIndex        =   0
      Top             =   30
      Width           =   510
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6240
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7980
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   120
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4965
      Left            =   60
      TabIndex        =   3
      Top             =   720
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   8758
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   34
      FixedCols       =   2
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   34
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Line Line1 
      X1              =   1860
      X2              =   2370
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Label Label1 
      Caption         =   "統計月份："
      Height          =   180
      Left            =   570
      TabIndex        =   8
      Top             =   420
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "統計年度："
      Height          =   180
      Left            =   570
      TabIndex        =   7
      Top             =   90
      Width           =   900
   End
End
Attribute VB_Name = "frm210135"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo by Sindy 2010/11/9 日期欄已修改
Option Explicit

Dim ii As Integer, jj As Integer
Dim strYear As String, strMonth_S As String, strMonth_E As String
Dim xlsAnnuity As New Excel.Application
Dim wksAnnuity As New Worksheet
Dim intCounter As Integer
Dim dbl_cnt As Double

Private Sub SetDataListWidth()
Dim ii As Integer
   
   With GrdDataList
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(.col) = 1000: .Text = "智權人員"
      .ColAlignment(.col) = flexAlignCenterCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 1: .ColWidth(.col) = 1000: .Text = "　　　"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 2: .ColWidth(.col) = 1000: .Text = "一月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 3: .ColWidth(.col) = 1000: .Text = "二月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 4: .ColWidth(.col) = 1000: .Text = "三月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 5: .ColWidth(.col) = 1000: .Text = "四月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 6: .ColWidth(.col) = 1000: .Text = "五月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 7: .ColWidth(.col) = 1000: .Text = "六月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 8: .ColWidth(.col) = 1000: .Text = "七月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 9: .ColWidth(.col) = 1000: .Text = "八月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 10: .ColWidth(.col) = 1000: .Text = "九月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 11: .ColWidth(.col) = 1000: .Text = "十月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 12: .ColWidth(.col) = 1000: .Text = "十一月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 13: .ColWidth(.col) = 1000: .Text = "十二月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 14: .ColWidth(.col) = 1000: .Text = "合計"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 15: .ColWidth(.col) = 1000: .Text = "年平均"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      For ii = 16 To .Cols - 1
         .ColWidth(ii) = 0
      Next
      .Visible = True
   End With
End Sub

Private Sub SetDataListWidth2()
Dim ii As Integer
   
   With grdDataList2
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(.col) = 1000: .Text = "智權人員"
      .ColAlignment(.col) = flexAlignCenterCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 1: .ColWidth(.col) = 1000: .Text = "　　　"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 2: .ColWidth(.col) = 1000: .Text = "一月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 3: .ColWidth(.col) = 1000: .Text = "二月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 4: .ColWidth(.col) = 1000: .Text = "三月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 5: .ColWidth(.col) = 1000: .Text = "四月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 6: .ColWidth(.col) = 1000: .Text = "五月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 7: .ColWidth(.col) = 1000: .Text = "六月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 8: .ColWidth(.col) = 1000: .Text = "七月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 9: .ColWidth(.col) = 1000: .Text = "八月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 10: .ColWidth(.col) = 1000: .Text = "九月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 11: .ColWidth(.col) = 1000: .Text = "十月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 12: .ColWidth(.col) = 1000: .Text = "十一月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 13: .ColWidth(.col) = 1000: .Text = "十二月"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 14: .ColWidth(.col) = 1000: .Text = "合計"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .col = 15: .ColWidth(.col) = 1000: .Text = "年平均"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      For ii = 16 To .Cols - 1
         .ColWidth(ii) = 0
      Next
      .Visible = True
   End With
End Sub

Private Sub cmdExcel_Click()
   Screen.MousePointer = vbHourglass
   PrintExcel
   MsgBox "產生Excel檔，完成！", vbInformation
   Screen.MousePointer = vbDefault
End Sub

Public Sub PrintExcel()
Dim strTemp As String
Dim ii As Integer
   
On Error GoTo ErrHnd
   
   intCounter = 0
   Set xlsAnnuity = New Excel.Application
   xlsAnnuity.Workbooks.add
   Set wksAnnuity = xlsAnnuity.Worksheets(1)
   With wksAnnuity
      .PageSetup.Orientation = xlLandscape '橫印
      .PageSetup.PrintTitleRows = "$1:$1" '標題列
      .PageSetup.PrintTitleColumns = "$A:$B" '標題欄
      .PageSetup.LeftHeader = "日期：&""新細明體,標準""&D"
      .PageSetup.CenterHeader = "&""新細明體,標準""" & strYear & "年度智權部業績明細表"
      .PageSetup.RightHeader = "第　&""新細明體,標準""&P&""新細明體,標準""　&""新細明體,標準""頁"
      .PageSetup.LeftMargin = 20
      .PageSetup.RightMargin = 20
      .PageSetup.TopMargin = 30
      .PageSetup.BottomMargin = 30
      .PageSetup.HeaderMargin = 15
      .PageSetup.FooterMargin = 15
      '設定各欄位長度
      .Columns("A:A").ColumnWidth = 10
      .Columns("B:B").ColumnWidth = 8
      .Columns("C:C").ColumnWidth = 10
      .Columns("D:D").ColumnWidth = 10
      .Columns("E:E").ColumnWidth = 10
      .Columns("F:F").ColumnWidth = 10
      .Columns("G:G").ColumnWidth = 10
      .Columns("H:H").ColumnWidth = 10
      .Columns("I:I").ColumnWidth = 10
      .Columns("J:J").ColumnWidth = 10
      .Columns("K:K").ColumnWidth = 10
      .Columns("L:L").ColumnWidth = 10
      .Columns("M:M").ColumnWidth = 10
      .Columns("N:N").ColumnWidth = 10
      .Columns("O:O").ColumnWidth = 10
      .Columns("P:P").ColumnWidth = 10
      intCounter = intCounter + 1
      .Range("A" & intCounter).Value = "智權人員"
      .Range("B" & intCounter).Value = ""
      .Range("C" & intCounter).Value = "一月"
      .Range("D" & intCounter).Value = "二月"
      .Range("E" & intCounter).Value = "三月"
      .Range("F" & intCounter).Value = "四月"
      .Range("G" & intCounter).Value = "五月"
      .Range("H" & intCounter).Value = "六月"
      .Range("I" & intCounter).Value = "七月"
      .Range("J" & intCounter).Value = "八月"
      .Range("K" & intCounter).Value = "九月"
      .Range("L" & intCounter).Value = "十月"
      .Range("M" & intCounter).Value = "十一月"
      .Range("N" & intCounter).Value = "十二月"
      .Range("O" & intCounter).Value = "合計"
      .Range("P" & intCounter).Value = "年平均"
      strTemp = "C" & intCounter & ":P" & intCounter
      .Range(strTemp).Select
      With .Application.Selection
         .HorizontalAlignment = xlCenter
         .Font.Bold = True
      End With
      strTemp = "A" & intCounter & ":P" & intCounter
      .Range(strTemp).Select
      With .Application.Selection.Borders(xlEdgeTop)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
      End With
      With .Application.Selection.Borders(xlEdgeBottom)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
      End With
      '逐筆填值
      ii = 1
      Do While ii < GrdDataList.Rows
         intCounter = intCounter + 1
         .Range("A" & intCounter).Value = GrdDataList.TextMatrix(ii, 0)
         .Range("B" & intCounter).Value = GrdDataList.TextMatrix(ii, 1)
         .Range("C" & intCounter).Value = GrdDataList.TextMatrix(ii, 2)
         .Range("D" & intCounter).Value = GrdDataList.TextMatrix(ii, 3)
         .Range("E" & intCounter).Value = GrdDataList.TextMatrix(ii, 4)
         .Range("F" & intCounter).Value = GrdDataList.TextMatrix(ii, 5)
         .Range("G" & intCounter).Value = GrdDataList.TextMatrix(ii, 6)
         .Range("H" & intCounter).Value = GrdDataList.TextMatrix(ii, 7)
         .Range("I" & intCounter).Value = GrdDataList.TextMatrix(ii, 8)
         .Range("J" & intCounter).Value = GrdDataList.TextMatrix(ii, 9)
         .Range("K" & intCounter).Value = GrdDataList.TextMatrix(ii, 10)
         .Range("L" & intCounter).Value = GrdDataList.TextMatrix(ii, 11)
         .Range("M" & intCounter).Value = GrdDataList.TextMatrix(ii, 12)
         .Range("N" & intCounter).Value = GrdDataList.TextMatrix(ii, 13)
         .Range("O" & intCounter).Value = GrdDataList.TextMatrix(ii, 14)
         .Range("P" & intCounter).Value = GrdDataList.TextMatrix(ii, 15)
         If GrdDataList.TextMatrix(ii, 1) = "目　標" Or _
            GrdDataList.TextMatrix(ii, 1) = "點　數" Or _
            GrdDataList.TextMatrix(ii, 1) = "達成率" Then
            strTemp = "A" & intCounter & ":P" & intCounter
            .Range(strTemp).Select
            If GrdDataList.TextMatrix(ii, 1) = "目　標" Or GrdDataList.TextMatrix(ii, 1) = "點　數" Then
               With .Application.Selection.Borders(xlEdgeTop)
                  .LineStyle = xlContinuous
                  .Weight = xlThin
                  .ColorIndex = xlAutomatic
               End With
            End If
            If GrdDataList.TextMatrix(ii, 1) = "達成率" Or GrdDataList.TextMatrix(ii, 1) = "點　數" Then
               With .Application.Selection.Borders(xlEdgeBottom)
                  .LineStyle = xlContinuous
                  .Weight = xlThin
                  .ColorIndex = xlAutomatic
               End With
            End If
         End If
         ii = ii + 1
      Loop
   End With
   strTemp = "A1:B" & intCounter
   wksAnnuity.Range(strTemp).Select
   With wksAnnuity.Application.Selection
       .Font.Bold = True
   End With
   strTemp = "A1:P" & intCounter
   wksAnnuity.Range(strTemp).Select
   With wksAnnuity.Application.Selection.Borders(xlEdgeLeft)
      .LineStyle = xlContinuous
      .Weight = xlThin
      .ColorIndex = xlAutomatic
   End With
   With wksAnnuity.Application.Selection.Borders(xlEdgeRight)
      .LineStyle = xlContinuous
      .Weight = xlThin
      .ColorIndex = xlAutomatic
   End With
   With wksAnnuity.Application.Selection.Borders(xlInsideVertical)
      .LineStyle = xlContinuous
      .Weight = xlThin
      .ColorIndex = xlAutomatic
   End With
   
   xlsAnnuity.Visible = True
   xlsAnnuity.WindowState = wdWindowStateMaximize
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing
   Exit Sub
   
ErrHnd:
   xlsAnnuity.Visible = True
   xlsAnnuity.WindowState = wdWindowStateMaximize
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdSearch_Click()
Dim Cancel As Boolean
   Screen.MousePointer = vbHourglass
   If Text1(0) = "" Then
      MsgBox "統計年度不可空白!!!", vbExclamation + vbOKOnly
      Call Text1_GotFocus(0)
      Exit Sub
   End If
   If Text1(1) = "" Then
      MsgBox "統計起始月份不可空白!!!", vbExclamation + vbOKOnly
      Call Text1_GotFocus(1)
      Exit Sub
   End If
   If Text1(2) = "" Then
      MsgBox "統計截止月份不可空白!!!", vbExclamation + vbOKOnly
      Call Text1_GotFocus(2)
      Exit Sub
   End If
   Cancel = False
   Call Text1_Validate(1, Cancel)
   If Cancel = True Then Exit Sub
   Cancel = False
   Call Text1_Validate(2, Cancel)
   If Cancel = True Then Exit Sub
   Call doQuery
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Text1(0) = Val(Left(Trim(strSrvDate(1)), 4)) - 1911
   Text1(1) = 1
   Text1(2) = Mid(Trim(strSrvDate(1)), 5, 2)
   CmdExcel.Enabled = False
   Call SetDataListWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm210135 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   CloseIme
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 1
         If Val(Text1(Index)) <= 0 Or Val(Text1(Index)) > 12 Then
            MsgBox "月份輸入錯誤!!!", vbExclamation + vbOKOnly
            Call Text1_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      Case 2
         If Val(Text1(Index)) <= 0 Or Val(Text1(Index)) > 12 Then
            MsgBox "月份輸入錯誤!!!", vbExclamation + vbOKOnly
            Call Text1_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
         If Val(Text1(1)) > Val(Text1(2)) Then
            MsgBox "截止月份不可小於起始月份!!!", vbExclamation + vbOKOnly
            Call Text1_GotFocus(2)
            Cancel = True
            Exit Sub
         End If
   End Select
End Sub

Private Function doQuery() As Boolean
Dim strSql_S As String, strSql As String, strSql_E As String
   
On Error GoTo ErrHnd
   
   strYear = Val(Text1(0))
   strMonth_S = IIf(Len(Text1(1)) = 1, "0" & Trim(Text1(1)), Text1(1))
   strMonth_E = IIf(Len(Text1(2)) = 1, "0" & Trim(Text1(2)), Text1(2))
   dbl_cnt = Val(strMonth_E) - Val(strMonth_S) + 1
   
   strSql_S = "": strSql = "": strSql_E = ""
   '點數合計 = 財務累計 + 結餘 + 保留
   strSql_S = "select 智權人員,'目　標',sum(Obj1) 目標1,sum(Obj2) 目標2,sum(Obj3) 目標3,sum(Obj4) 目標4,sum(Obj5) 目標5,sum(Obj6) 目標6,sum(Obj7) 目標7,sum(Obj8) 目標8,sum(Obj9) 目標9,sum(Obj10) 目標10,sum(Obj11) 目標11,sum(Obj12) 目標12,sum(Obj1)+sum(Obj2)+sum(Obj3)+sum(Obj4)+sum(Obj5)+sum(Obj6)+sum(Obj7)+sum(Obj8)+sum(Obj9)+sum(Obj10)+sum(Obj11)+sum(Obj12) 目標合計,to_char((sum(Obj1)+sum(Obj2)+sum(Obj3)+sum(Obj4)+sum(Obj5)+sum(Obj6)+sum(Obj7)+sum(Obj8)+sum(Obj9)+sum(Obj10)+sum(Obj11)+sum(Obj12))/" & dbl_cnt & ",'999999.00') 目標年平均,Zone,Area,AreaName,st01,sum(P1) 點數1,sum(P2) 點數2,sum(P3) 點數3,sum(P4) 點數4,sum(P5) 點數5,sum(P6) 點數6,sum(P7) 點數7,sum(P8) 點數8,sum(P9) 點數9,sum(P10) 點數10,sum(P11) 點數11,sum(P12) 點數12,sum(P1)+sum(P2)+sum(P3)+sum(P4)+sum(P5)+sum(P6)+sum(P7)+sum(P8)+sum(P9)+sum(P10)+sum(P11)+sum(P12) 點數合計,to_char((sum(P1)+sum(P2)+sum(P3)+sum(P4)+sum(P5)+sum(P6)+sum(P7)+sum(P8)+sum(P9)+sum(P10)+sum(P11)+sum(P12))/" & dbl_cnt & ",'999999.00') 點數年平均 " & _
                "from ( "
   '2014/1/21 modify by sonia 下列語法取消a0201='1'條件
   '1月
   If Val(Text1(1)) <= 1 And Val(Text1(2)) >= 1 Then
      If strSql <> "" Then strSql = strSql & "Union "
      '2015/4/28 MODIFY BY SONIA 1.保留科目加'4194' 2.過濾結餘傳票取消科目限制但必須為收入科目 3.有些人沒目標有點數,有些人有目標沒點數
      '2015/4/28 以下1~12月語法都改
      'strSql = strSql & "select st02 智權人員,(to_char(ROUND(nvl(W1_1,0)/1000,3)+ROUND(nvl(Z1_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(Y1_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(X1_1,0)/1000,3),'999999.00')) P1,0 P2,0 P3,0 P4,0 P5,0 P6,0 P7,0 P8,0 P9,0 P10,0 P11,0 P12,decode(substr(st15,1,1),'S',st06,'F','6','5') Zone,st15 Area,a0902 AreaName,st01,pe04 Obj1,0 Obj2,0 Obj3,0 Obj4,0 Obj5,0 Obj6,0 Obj7,0 Obj8,0 Obj9,0 Obj10,0 Obj11,0 Obj12 " & _
                                 "from staff,performance, " & _
                                 "(select ax209 W1_0,sum(ax207) W1_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0101 and a0205 <= " & Val(Text1(0)) & "0131 and ax209 is not null and (substr(ax205, 1, 2) = '41' Or ax205 = '7121') and ax207>0 and not( ax205='4191' or ax205='4192' or ((ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131') and instr(ax213||' ','結餘')>0)) group by ax209) W1, " & _
                                 "(select ax209 X1_0,sum(ax207-ax206) X1_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0101 and a0205 <= " & Val(Text1(0)) & "0131 and ax209 is not null and (ax205='4191' or ax205='4192') group by ax209) X1, " & _
                                 "(select ax209 Y1_0,sum(ax207-ax206) Y1_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0101 and a0205 <= " & Val(Text1(0)) & "0131 and ax209 is not null  and (ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131') and instr(ax213||' ','結餘')>0 group by ax209) Y1, " & _
                                 "(select ax209 Z1_0,sum(-1*ax206) Z1_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0101 and a0205 <= " & Val(Text1(0)) & "0131 and ax209 is not null and ax206>0 and (substr(ax205, 1, 2) = '41' Or ax205 = '7121') and not ( (ax205='4191' or ax205='4192') or ( (ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131') and instr(ax213||' ','結餘')>0 ) ) group by ax209) Z1,acc090 " & _
                                 "where W1_0(+)=st01 and X1_0(+)=st01 and Y1_0(+)=st01 and Z1_0(+)=st01 and a0901(+)=st15 and (nvl(W1_1,0)<>0 or nvl(Y1_1,0)<>0 or nvl(X1_1,0)<>0 or nvl(Z1_1,0)<>0) and a0901(+)=st15 " & _
                                 "and pe01(+)=st01 and pe02(+)='TOT' and pe03(+)=" & Val(Text1(0)) + 1911 & "01 "
      'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
      strSql = strSql & "select st02 智權人員,(to_char(ROUND(nvl(W1_1,0)/1000,3)+ROUND(nvl(Z1_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(Y1_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(X1_1,0)/1000,3),'999999.00')) P1,0 P2,0 P3,0 P4,0 P5,0 P6,0 P7,0 P8,0 P9,0 P10,0 P11,0 P12,decode(substr(st15,1,1),'S',st06,'F','6','5') Zone,st15 Area,a0902 AreaName,st01,pe04 Obj1,0 Obj2,0 Obj3,0 Obj4,0 Obj5,0 Obj6,0 Obj7,0 Obj8,0 Obj9,0 Obj10,0 Obj11,0 Obj12 " & _
                                 "from staff,performance, " & _
                                 "(select pe01 saleno from performance where pe03=" & Val(Text1(0)) + 1911 & "01 and pe02='TOT' and pe04>0 union select distinct ax209 from acc020,acc021 where a0205 >= " & Val(Text1(0)) & "0101 and a0205 <= " & Val(Text1(0)) & "0131 and a0201=ax201(+) and a0202=ax202(+) and (substr(ax205,1,1)='4' or ax205='7121') and ax209 is not null) P1, " & _
                                 "(select ax209 W1_0,sum(ax207) W1_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0101 and a0205 <= " & Val(Text1(0)) & "0131 and ax209 is not null and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and ax207>0 and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) W1, " & _
                                 "(select ax209 X1_0,sum(ax207-ax206) X1_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0101 and a0205 <= " & Val(Text1(0)) & "0131 and ax209 is not null and (ax205='4191' or ax205='4192' or ax205='4194') group by ax209) X1, " & _
                                 "(select ax209 Y1_0,sum(ax207-ax206) Y1_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0101 and a0205 <= " & Val(Text1(0)) & "0131 and ax209 is not null and substr(ax205,1,1)='4' and instr(ax213||' ','結餘')>0 group by ax209) Y1, " & _
                                 "(select ax209 Z1_0,sum(-1*ax206) Z1_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0101 and a0205 <= " & Val(Text1(0)) & "0131 and ax209 is not null and ax206>0 and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) Z1,acc090 " & _
                                 "where P1.saleno=st01 and W1_0(+)=st01 and X1_0(+)=st01 and Y1_0(+)=st01 and Z1_0(+)=st01 and a0901(+)=st15 and (nvl(W1_1,0)<>0 or nvl(Y1_1,0)<>0 or nvl(X1_1,0)<>0 or nvl(Z1_1,0)<>0 or nvl(pe04,0)<>0) and a0901(+)=st15 " & _
                                 "and pe01(+)=st01 and pe02(+)='TOT' and pe03(+)=" & Val(Text1(0)) + 1911 & "01 "
   End If
   '2月
   If Val(Text1(1)) <= 2 And Val(Text1(2)) >= 2 Then
      If strSql <> "" Then strSql = strSql & "Union "
      'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
      strSql = strSql & "select st02 智權人員,0 P1,(to_char(ROUND(nvl(W2_1,0)/1000,3)+ROUND(nvl(Z2_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(Y2_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(X2_1,0)/1000,3),'999999.00')) P2,0 P3,0 P4,0 P5,0 P6,0 P7,0 P8,0 P9,0 P10,0 P11,0 P12,decode(substr(st15,1,1),'S',st06,'F','6','5') Zone,st15 Area,a0902 AreaName,st01,0 Obj1,pe04 Obj2,0 Obj3,0 Obj4,0 Obj5,0 Obj6,0 Obj7,0 Obj8,0 Obj9,0 Obj10,0 Obj11,0 Obj12 " & _
                                 "from staff,performance, " & _
                                 "(select pe01 saleno from performance where pe03=" & Val(Text1(0)) + 1911 & "02 and pe02='TOT' and pe04>0 union select distinct ax209 from acc020,acc021 where a0205 >= " & Val(Text1(0)) & "0201 and a0205 <= " & Val(Text1(0)) & "0231 and a0201=ax201(+) and a0202=ax202(+) and (substr(ax205,1,1)='4' or ax205='7121') and ax209 is not null) P2, " & _
                                 "(select ax209 W2_0,sum(ax207) W2_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0201 and a0205 <= " & Val(Text1(0)) & "0231 and ax209 is not null and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and ax207>0 and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) W2, " & _
                                 "(select ax209 X2_0,sum(ax207-ax206) X2_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0201 and a0205 <= " & Val(Text1(0)) & "0231 and ax209 is not null and (ax205='4191' or ax205='4192' or ax205='4194') group by ax209) X2, " & _
                                 "(select ax209 Y2_0,sum(ax207-ax206) Y2_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0201 and a0205 <= " & Val(Text1(0)) & "0231 and ax209 is not null and substr(ax205,1,1)='4' and instr(ax213||' ','結餘')>0 group by ax209) Y2, " & _
                                 "(select ax209 Z2_0,sum(-1*ax206) Z2_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0201 and a0205 <= " & Val(Text1(0)) & "0231 and ax209 is not null and ax206>0 and (substr(ax205, 1,1) = '4' Or ax205 = '7121') and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) Z2,acc090 " & _
                                 "where P2.saleno=st01 and W2_0(+)=st01 and X2_0(+)=st01 and Y2_0(+)=st01 and Z2_0(+)=st01 and a0901(+)=st15 and (nvl(W2_1,0)<>0 or nvl(Y2_1,0)<>0 or nvl(X2_1,0)<>0 or nvl(Z2_1,0)<>0 or nvl(pe04,0)<>0) and a0901(+)=st15 " & _
                                 "and pe01(+)=st01 and pe02(+)='TOT' and pe03(+)=" & Val(Text1(0)) + 1911 & "02 "
   End If
   '3月
   If Val(Text1(1)) <= 3 And Val(Text1(2)) >= 3 Then
      If strSql <> "" Then strSql = strSql & "Union "
      'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
      strSql = strSql & "select st02 智權人員,0 P1,0 P2,(to_char(ROUND(nvl(W3_1,0)/1000,3)+ROUND(nvl(Z3_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(Y3_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(X3_1,0)/1000,3),'999999.00')) P3,0 P4,0 P5,0 P6,0 P7,0 P8,0 P9,0 P10,0 P11,0 P12,decode(substr(st15,1,1),'S',st06,'F','6','5') Zone,st15 Area,a0902 AreaName,st01,0 Obj1,0 Obj2,pe04 Obj3,0 Obj4,0 Obj5,0 Obj6,0 Obj7,0 Obj8,0 Obj9,0 Obj10,0 Obj11,0 Obj12 " & _
                                 "from staff,performance, " & _
                                 "(select pe01 saleno from performance where pe03=" & Val(Text1(0)) + 1911 & "03 and pe02='TOT' and pe04>0 union select distinct ax209 from acc020,acc021 where a0205 >= " & Val(Text1(0)) & "0301 and a0205 <= " & Val(Text1(0)) & "0331 and a0201=ax201(+) and a0202=ax202(+) and (substr(ax205,1,1)='4' or ax205='7121') and ax209 is not null) P3, " & _
                                 "(select ax209 W3_0,sum(ax207) W3_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0301 and a0205 <= " & Val(Text1(0)) & "0331 and ax209 is not null and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and ax207>0 and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) W3, " & _
                                 "(select ax209 X3_0,sum(ax207-ax206) X3_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0301 and a0205 <= " & Val(Text1(0)) & "0331 and ax209 is not null and (ax205='4191' or ax205='4192' or ax205='4194') group by ax209) X3, " & _
                                 "(select ax209 Y3_0,sum(ax207-ax206) Y3_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0301 and a0205 <= " & Val(Text1(0)) & "0331 and ax209 is not null and substr(ax205,1,1)='4' and instr(ax213||' ','結餘')>0 group by ax209) Y3, " & _
                                 "(select ax209 Z3_0,sum(-1*ax206) Z3_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0301 and a0205 <= " & Val(Text1(0)) & "0331 and ax209 is not null and ax206>0 and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) Z3,acc090 " & _
                                 "where P3.saleno=st01 and W3_0(+)=st01 and X3_0(+)=st01 and Y3_0(+)=st01 and Z3_0(+)=st01 and a0901(+)=st15 and (nvl(W3_1,0)<>0 or nvl(Y3_1,0)<>0 or nvl(X3_1,0)<>0 or nvl(Z3_1,0)<>0 or nvl(pe04,0)<>0) and a0901(+)=st15 " & _
                                 "and pe01(+)=st01 and pe02(+)='TOT' and pe03(+)=" & Val(Text1(0)) + 1911 & "03 "
   End If
   '4月
   If Val(Text1(1)) <= 4 And Val(Text1(2)) >= 4 Then
      If strSql <> "" Then strSql = strSql & "Union "
      'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
      strSql = strSql & "select st02 智權人員,0 P1,0 P2,0 P3,(to_char(ROUND(nvl(W4_1,0)/1000,3)+ROUND(nvl(Z4_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(Y4_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(X4_1,0)/1000,3),'999999.00')) P4,0 P5,0 P6,0 P7,0 P8,0 P9,0 P10,0 P11,0 P12,decode(substr(st15,1,1),'S',st06,'F','6','5') Zone,st15 Area,a0902 AreaName,st01,0 Obj1,0 Obj2,0 Obj3,pe04 Obj4,0 Obj5,0 Obj6,0 Obj7,0 Obj8,0 Obj9,0 Obj10,0 Obj11,0 Obj12 " & _
                                 "from staff,performance, " & _
                                 "(select pe01 saleno from performance where pe03=" & Val(Text1(0)) + 1911 & "04 and pe02='TOT' and pe04>0 union select distinct ax209 from acc020,acc021 where a0205 >= " & Val(Text1(0)) & "0401 and a0205 <= " & Val(Text1(0)) & "0431 and a0201=ax201(+) and a0202=ax202(+) and (substr(ax205,1,1)='4' or ax205='7121') and ax209 is not null) P4, " & _
                                 "(select ax209 W4_0,sum(ax207) W4_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0401 and a0205 <= " & Val(Text1(0)) & "0431 and ax209 is not null and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and ax207>0 and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) W4, " & _
                                 "(select ax209 X4_0,sum(ax207-ax206) X4_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0401 and a0205 <= " & Val(Text1(0)) & "0431 and ax209 is not null and (ax205='4191' or ax205='4192' or ax205='4194') group by ax209) X4, " & _
                                 "(select ax209 Y4_0,sum(ax207-ax206) Y4_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0401 and a0205 <= " & Val(Text1(0)) & "0431 and ax209 is not null and substr(ax205,1,1)='4' and instr(ax213||' ','結餘')>0 group by ax209) Y4, " & _
                                 "(select ax209 Z4_0,sum(-1*ax206) Z4_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0401 and a0205 <= " & Val(Text1(0)) & "0431 and ax209 is not null and ax206>0 and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) Z4,acc090 " & _
                                 "where P4.saleno=st01 and W4_0(+)=st01 and X4_0(+)=st01 and Y4_0(+)=st01 and Z4_0(+)=st01 and a0901(+)=st15 and (nvl(W4_1,0)<>0 or nvl(Y4_1,0)<>0 or nvl(X4_1,0)<>0 or nvl(Z4_1,0)<>0 or nvl(pe04,0)<>0) and a0901(+)=st15 " & _
                                 "and pe01(+)=st01 and pe02(+)='TOT' and pe03(+)=" & Val(Text1(0)) + 1911 & "04 "
   End If
   '5月
   If Val(Text1(1)) <= 5 And Val(Text1(2)) >= 5 Then
      If strSql <> "" Then strSql = strSql & "Union "
      'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
      strSql = strSql & "select st02 智權人員,0 P1,0 P2,0 P3,0 P4,(to_char(ROUND(nvl(W5_1,0)/1000,3)+ROUND(nvl(Z5_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(Y5_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(X5_1,0)/1000,3),'999999.00')) P5,0 P6,0 P7,0 P8,0 P9,0 P10,0 P11,0 P12,decode(substr(st15,1,1),'S',st06,'F','6','5') Zone,st15 Area,a0902 AreaName,st01,0 Obj1,0 Obj2,0 Obj3,0 Obj4,pe04 Obj5,0 Obj6,0 Obj7,0 Obj8,0 Obj9,0 Obj10,0 Obj11,0 Obj12 " & _
                                 "from staff,performance, " & _
                                 "(select pe01 saleno from performance where pe03=" & Val(Text1(0)) + 1911 & "05 and pe02='TOT' and pe04>0 union select distinct ax209 from acc020,acc021 where a0205 >= " & Val(Text1(0)) & "0501 and a0205 <= " & Val(Text1(0)) & "0531 and a0201=ax201(+) and a0202=ax202(+) and (substr(ax205,1,1)='4' or ax205='7121') and ax209 is not null) P5, " & _
                                 "(select ax209 W5_0,sum(ax207) W5_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0501 and a0205 <= " & Val(Text1(0)) & "0531 and ax209 is not null and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and ax207>0 and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) W5, " & _
                                 "(select ax209 X5_0,sum(ax207-ax206) X5_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0501 and a0205 <= " & Val(Text1(0)) & "0531 and ax209 is not null and (ax205='4191' or ax205='4192' or ax205='4194') group by ax209) X5, " & _
                                 "(select ax209 Y5_0,sum(ax207-ax206) Y5_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0501 and a0205 <= " & Val(Text1(0)) & "0531 and ax209 is not null and substr(ax205,1,1)='4' and instr(ax213||' ','結餘')>0 group by ax209) Y5, " & _
                                 "(select ax209 Z5_0,sum(-1*ax206) Z5_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0501 and a0205 <= " & Val(Text1(0)) & "0531 and ax209 is not null and ax206>0 and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) Z5,acc090 " & _
                                 "where P5.saleno=st01 and W5_0(+)=st01 and X5_0(+)=st01 and Y5_0(+)=st01 and Z5_0(+)=st01 and a0901(+)=st15 and (nvl(W5_1,0)<>0 or nvl(Y5_1,0)<>0 or nvl(X5_1,0)<>0 or nvl(Z5_1,0)<>0 or nvl(pe04,0)<>0) and a0901(+)=st15 " & _
                                 "and pe01(+)=st01 and pe02(+)='TOT' and pe03(+)=" & Val(Text1(0)) + 1911 & "05 "
   End If
   '6月
   If Val(Text1(1)) <= 6 And Val(Text1(2)) >= 6 Then
      If strSql <> "" Then strSql = strSql & "Union "
      'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
      strSql = strSql & "select st02 智權人員,0 P1,0 P2,0 P3,0 P4,0 P5,(to_char(ROUND(nvl(W6_1,0)/1000,3)+ROUND(nvl(Z6_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(Y6_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(X6_1,0)/1000,3),'999999.00')) P6,0 P7,0 P8,0 P9,0 P10,0 P11,0 P12,decode(substr(st15,1,1),'S',st06,'F','6','5') Zone,st15 Area,a0902 AreaName,st01,0 Obj1,0 Obj2,0 Obj3,0 Obj4,0 Obj5,pe04 Obj6,0 Obj7,0 Obj8,0 Obj9,0 Obj10,0 Obj11,0 Obj12 " & _
                                 "from staff,performance, " & _
                                 "(select pe01 saleno from performance where pe03=" & Val(Text1(0)) + 1911 & "06 and pe02='TOT' and pe04>0 union select distinct ax209 from acc020,acc021 where a0205 >= " & Val(Text1(0)) & "0601 and a0205 <= " & Val(Text1(0)) & "0631 and a0201=ax201(+) and a0202=ax202(+) and (substr(ax205,1,1)='4' or ax205='7121') and ax209 is not null) P6, " & _
                                 "(select ax209 W6_0,sum(ax207) W6_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0601 and a0205 <= " & Val(Text1(0)) & "0631 and ax209 is not null and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and ax207>0 and not( ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) W6, " & _
                                 "(select ax209 X6_0,sum(ax207-ax206) X6_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0601 and a0205 <= " & Val(Text1(0)) & "0631 and ax209 is not null and (ax205='4191' or ax205='4192' or ax205='4194') group by ax209) X6, " & _
                                 "(select ax209 Y6_0,sum(ax207-ax206) Y6_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0601 and a0205 <= " & Val(Text1(0)) & "0631 and ax209 is not null and substr(ax205,1,1)='4' and instr(ax213||' ','結餘')>0 group by ax209) Y6, " & _
                                 "(select ax209 Z6_0,sum(-1*ax206) Z6_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0601 and a0205 <= " & Val(Text1(0)) & "0631 and ax209 is not null and ax206>0 and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) Z6,acc090 " & _
                                 "where P6.saleno=st01 and W6_0(+)=st01 and X6_0(+)=st01 and Y6_0(+)=st01 and Z6_0(+)=st01 and a0901(+)=st15 and (nvl(W6_1,0)<>0 or nvl(Y6_1,0)<>0 or nvl(X6_1,0)<>0 or nvl(Z6_1,0)<>0 or nvl(pe04,0)<>0) and a0901(+)=st15 " & _
                                 "and pe01(+)=st01 and pe02(+)='TOT' and pe03(+)=" & Val(Text1(0)) + 1911 & "06 "
   End If
   '7月
   If Val(Text1(1)) <= 7 And Val(Text1(2)) >= 7 Then
      If strSql <> "" Then strSql = strSql & "Union "
      'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
      strSql = strSql & "select st02 智權人員,0 P1,0 P2,0 P3,0 P4,0 P5,0 P6,(to_char(ROUND(nvl(W7_1,0)/1000,3)+ROUND(nvl(Z7_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(Y7_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(X7_1,0)/1000,3),'999999.00')) P7,0 P8,0 P9,0 P10,0 P11,0 P12,decode(substr(st15,1,1),'S',st06,'F','6','5') Zone,st15 Area,a0902 AreaName,st01,0 Obj1,0 Obj2,0 Obj3,0 Obj4,0 Obj5,0 Obj6,pe04 Obj7,0 Obj8,0 Obj9,0 Obj10,0 Obj11,0 Obj12 " & _
                                 "from staff,performance, " & _
                                 "(select pe01 saleno from performance where pe03=" & Val(Text1(0)) + 1911 & "07 and pe02='TOT' and pe04>0 union select distinct ax209 from acc020,acc021 where a0205 >= " & Val(Text1(0)) & "0701 and a0205 <= " & Val(Text1(0)) & "0731 and a0201=ax201(+) and a0202=ax202(+) and (substr(ax205,1,1)='4' or ax205='7121') and ax209 is not null) P7, " & _
                                 "(select ax209 W7_0,sum(ax207) W7_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0701 and a0205 <= " & Val(Text1(0)) & "0731 and ax209 is not null and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and ax207>0 and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) W7, " & _
                                 "(select ax209 X7_0,sum(ax207-ax206) X7_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0701 and a0205 <= " & Val(Text1(0)) & "0731 and ax209 is not null and (ax205='4191' or ax205='4192' or ax205='4194') group by ax209) X7, " & _
                                 "(select ax209 Y7_0,sum(ax207-ax206) Y7_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0701 and a0205 <= " & Val(Text1(0)) & "0731 and ax209 is not null and substr(ax205,1,1)='4' and instr(ax213||' ','結餘')>0 group by ax209) Y7, " & _
                                 "(select ax209 Z7_0,sum(-1*ax206) Z7_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0701 and a0205 <= " & Val(Text1(0)) & "0731 and ax209 is not null and ax206>0 and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) Z7,acc090 " & _
                                 "where P7.saleno=st01 and W7_0(+)=st01 and X7_0(+)=st01 and Y7_0(+)=st01 and Z7_0(+)=st01 and a0901(+)=st15 and (nvl(W7_1,0)<>0 or nvl(Y7_1,0)<>0 or nvl(X7_1,0)<>0 or nvl(Z7_1,0)<>0 or nvl(pe04,0)<>0) and a0901(+)=st15 " & _
                                 "and pe01(+)=st01 and pe02(+)='TOT' and pe03(+)=" & Val(Text1(0)) + 1911 & "07 "
   End If
   '8月
   If Val(Text1(1)) <= 8 And Val(Text1(2)) >= 8 Then
      If strSql <> "" Then strSql = strSql & "Union "
      'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
      strSql = strSql & "select st02 智權人員,0 P1,0 P2,0 P3,0 P4,0 P5,0 P6,0 P7,(to_char(ROUND(nvl(W8_1,0)/1000,3)+ROUND(nvl(Z8_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(Y8_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(X8_1,0)/1000,3),'999999.00')) P8,0 P9,0 P10,0 P11,0 P12,decode(substr(st15,1,1),'S',st06,'F','6','5') Zone,st15 Area,a0902 AreaName,st01,0 Obj1,0 Obj2,0 Obj3,0 Obj4,0 Obj5,0 Obj6,0 Obj7,pe04 Obj8,0 Obj9,0 Obj10,0 Obj11,0 Obj12 " & _
                                 "from staff,performance, " & _
                                 "(select pe01 saleno from performance where pe03=" & Val(Text1(0)) + 1911 & "08 and pe02='TOT' and pe04>0 union select distinct ax209 from acc020,acc021 where a0205 >= " & Val(Text1(0)) & "0801 and a0205 <= " & Val(Text1(0)) & "0831 and a0201=ax201(+) and a0202=ax202(+) and (substr(ax205,1,1)='4' or ax205='7121') and ax209 is not null) P8, " & _
                                 "(select ax209 W8_0,sum(ax207) W8_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0801 and a0205 <= " & Val(Text1(0)) & "0831 and ax209 is not null and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and ax207>0 and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) W8, " & _
                                 "(select ax209 X8_0,sum(ax207-ax206) X8_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0801 and a0205 <= " & Val(Text1(0)) & "0831 and ax209 is not null and (ax205='4191' or ax205='4192' or ax205='4194') group by ax209) X8, " & _
                                 "(select ax209 Y8_0,sum(ax207-ax206) Y8_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0801 and a0205 <= " & Val(Text1(0)) & "0831 and ax209 is not null and substr(ax205,1,1)='4' and instr(ax213||' ','結餘')>0 group by ax209) Y8, " & _
                                 "(select ax209 Z8_0,sum(-1*ax206) Z8_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0801 and a0205 <= " & Val(Text1(0)) & "0831 and ax209 is not null and ax206>0 and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) Z8,acc090 " & _
                                 "where P8.saleno=st01 and W8_0(+)=st01 and X8_0(+)=st01 and Y8_0(+)=st01 and Z8_0(+)=st01 and a0901(+)=st15 and (nvl(W8_1,0)<>0 or nvl(Y8_1,0)<>0 or nvl(X8_1,0)<>0 or nvl(Z8_1,0)<>0 or nvl(pe04,0)<>0) and a0901(+)=st15 " & _
                                 "and pe01(+)=st01 and pe02(+)='TOT' and pe03(+)=" & Val(Text1(0)) + 1911 & "08 "
   End If
   '9月
   If Val(Text1(1)) <= 9 And Val(Text1(2)) >= 9 Then
      If strSql <> "" Then strSql = strSql & "Union "
      'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
      strSql = strSql & "select st02 智權人員,0 P1,0 P2,0 P3,0 P4,0 P5,0 P6,0 P7,0 P8,(to_char(ROUND(nvl(W9_1,0)/1000,3)+ROUND(nvl(Z9_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(Y9_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(X9_1,0)/1000,3),'999999.00')) P9,0 P10,0 P11,0 P12,decode(substr(st15,1,1),'S',st06,'F','6','5') Zone,st15 Area,a0902 AreaName,st01,0 Obj1,0 Obj2,0 Obj3,0 Obj4,0 Obj5,0 Obj6,0 Obj7,0 Obj8,pe04 Obj9,0 Obj10,0 Obj11,0 Obj12 " & _
                                 "from staff,performance, " & _
                                 "(select pe01 saleno from performance where pe03=" & Val(Text1(0)) + 1911 & "09 and pe02='TOT' and pe04>0 union select distinct ax209 from acc020,acc021 where a0205 >= " & Val(Text1(0)) & "0801 and a0205 <= " & Val(Text1(0)) & "0931 and a0201=ax201(+) and a0202=ax202(+) and (substr(ax205,1,1)='4' or ax205='7121') and ax209 is not null) P9, " & _
                                 "(select ax209 W9_0,sum(ax207) W9_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0901 and a0205 <= " & Val(Text1(0)) & "0931 and ax209 is not null and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and ax207>0 and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) W9, " & _
                                 "(select ax209 X9_0,sum(ax207-ax206) X9_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0901 and a0205 <= " & Val(Text1(0)) & "0931 and ax209 is not null and (ax205='4191' or ax205='4192' or ax205='4194') group by ax209) X9, " & _
                                 "(select ax209 Y9_0,sum(ax207-ax206) Y9_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0901 and a0205 <= " & Val(Text1(0)) & "0931 and ax209 is not null and substr(ax205,1,1)='4' and instr(ax213||' ','結餘')>0 group by ax209) Y9, " & _
                                 "(select ax209 Z9_0,sum(-1*ax206) Z9_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "0901 and a0205 <= " & Val(Text1(0)) & "0931 and ax209 is not null and ax206>0 and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) Z9,acc090 " & _
                                 "where P9.saleno=st01 and W9_0(+)=st01 and X9_0(+)=st01 and Y9_0(+)=st01 and Z9_0(+)=st01 and a0901(+)=st15 and (nvl(W9_1,0)<>0 or nvl(Y9_1,0)<>0 or nvl(X9_1,0)<>0 or nvl(Z9_1,0)<>0 or nvl(pe04,0)<>0) and a0901(+)=st15 " & _
                                 "and pe01(+)=st01 and pe02(+)='TOT' and pe03(+)=" & Val(Text1(0)) + 1911 & "09 "
   End If
   '10月
   If Val(Text1(1)) <= 10 And Val(Text1(2)) >= 10 Then
      If strSql <> "" Then strSql = strSql & "Union "
      'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
      strSql = strSql & "select st02 智權人員,0 P1,0 P2,0 P3,0 P4,0 P5,0 P6,0 P7,0 P8,0 P9,(to_char(ROUND(nvl(W10_1,0)/1000,3)+ROUND(nvl(Z10_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(Y10_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(X10_1,0)/1000,3),'999999.00')) P10,0 P11,0 P12,decode(substr(st15,1,1),'S',st06,'F','6','5') Zone,st15 Area,a0902 AreaName,st01,0 Obj1,0 Obj2,0 Obj3,0 Obj4,0 Obj5,0 Obj6,0 Obj7,0 Obj8,0 Obj9,pe04 Obj10,0 Obj11,0 Obj12 " & _
                                 "from staff,performance, " & _
                                 "(select pe01 saleno from performance where pe03=" & Val(Text1(0)) + 1911 & "10 and pe02='TOT' and pe04>0 union select distinct ax209 from acc020,acc021 where a0205 >= " & Val(Text1(0)) & "1001 and a0205 <= " & Val(Text1(0)) & "1031 and a0201=ax201(+) and a0202=ax202(+) and (substr(ax205,1,1)='4' or ax205='7121') and ax209 is not null) P10, " & _
                                 "(select ax209 W10_0,sum(ax207) W10_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "1001 and a0205 <= " & Val(Text1(0)) & "1031 and ax209 is not null and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and ax207>0 and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) W10, " & _
                                 "(select ax209 X10_0,sum(ax207-ax206) X10_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "1001 and a0205 <= " & Val(Text1(0)) & "1031 and ax209 is not null and (ax205='4191' or ax205='4192' or ax205='4194') group by ax209) X10, " & _
                                 "(select ax209 Y10_0,sum(ax207-ax206) Y10_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "1001 and a0205 <= " & Val(Text1(0)) & "1031 and ax209 is not null and substr(ax205,1,1)='4' and instr(ax213||' ','結餘')>0 group by ax209) Y10, " & _
                                 "(select ax209 Z10_0,sum(-1*ax206) Z10_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "1001 and a0205 <= " & Val(Text1(0)) & "1031 and ax209 is not null and ax206>0 and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) Z10,acc090 " & _
                                 "where P10.saleno=st01 and W10_0(+)=st01 and X10_0(+)=st01 and Y10_0(+)=st01 and Z10_0(+)=st01 and a0901(+)=st15 and (nvl(W10_1,0)<>0 or nvl(Y10_1,0)<>0 or nvl(X10_1,0)<>0 or nvl(Z10_1,0)<>0 or nvl(pe04,0)<>0) and a0901(+)=st15 " & _
                                 "and pe01(+)=st01 and pe02(+)='TOT' and pe03(+)=" & Val(Text1(0)) + 1911 & "10 "
   End If
   '11月
   If Val(Text1(1)) <= 11 And Val(Text1(2)) >= 11 Then
      If strSql <> "" Then strSql = strSql & "Union "
      'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
      strSql = strSql & "select st02 智權人員,0 P1,0 P2,0 P3,0 P4,0 P5,0 P6,0 P7,0 P8,0 P9,0 P10,(to_char(ROUND(nvl(W11_1,0)/1000,3)+ROUND(nvl(Z11_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(Y11_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(X11_1,0)/1000,3),'999999.00')) P11,0 P12,decode(substr(st15,1,1),'S',st06,'F','6','5') Zone,st15 Area,a0902 AreaName,st01,0 Obj1,0 Obj2,0 Obj3,0 Obj4,0 Obj5,0 Obj6,0 Obj7,0 Obj8,0 Obj9,0 Obj10,pe04 Obj11,0 Obj12 " & _
                                 "from staff,performance, " & _
                                 "(select pe01 saleno from performance where pe03=" & Val(Text1(0)) + 1911 & "11 and pe02='TOT' and pe04>0 union select distinct ax209 from acc020,acc021 where a0205 >= " & Val(Text1(0)) & "1101 and a0205 <= " & Val(Text1(0)) & "1131 and a0201=ax201(+) and a0202=ax202(+) and (substr(ax205,1,1)='4' or ax205='7121') and ax209 is not null) P11, " & _
                                 "(select ax209 W11_0,sum(ax207) W11_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "1101 and a0205 <= " & Val(Text1(0)) & "1131 and ax209 is not null and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and ax207>0 and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) W11, " & _
                                 "(select ax209 X11_0,sum(ax207-ax206) X11_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "1101 and a0205 <= " & Val(Text1(0)) & "1131 and ax209 is not null and (ax205='4191' or ax205='4192' or ax205='4194') group by ax209) X11, " & _
                                 "(select ax209 Y11_0,sum(ax207-ax206) Y11_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "1101 and a0205 <= " & Val(Text1(0)) & "1131 and ax209 is not null and substr(ax205,1,1)='4' and instr(ax213||' ','結餘')>0 group by ax209) Y11, " & _
                                 "(select ax209 Z11_0,sum(-1*ax206) Z11_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "1101 and a0205 <= " & Val(Text1(0)) & "1131 and ax209 is not null and ax206>0 and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) Z11,acc090 " & _
                                 "where P11.saleno=st01 and W11_0(+)=st01 and X11_0(+)=st01 and Y11_0(+)=st01 and Z11_0(+)=st01 and a0901(+)=st15 and (nvl(W11_1,0)<>0 or nvl(Y11_1,0)<>0 or nvl(X11_1,0)<>0 or nvl(Z11_1,0)<>0 or nvl(pe04,0)<>0) and a0901(+)=st15 " & _
                                 "and pe01(+)=st01 and pe02(+)='TOT' and pe03(+)=" & Val(Text1(0)) + 1911 & "11 "
   End If
   '12月
   If Val(Text1(1)) <= 12 And Val(Text1(2)) >= 12 Then
      If strSql <> "" Then strSql = strSql & "Union "
      'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
      strSql = strSql & "select st02 智權人員,0 P1,0 P2,0 P3,0 P4,0 P5,0 P6,0 P7,0 P8,0 P9,0 P10,0 P11,(to_char(ROUND(nvl(W12_1,0)/1000,3)+ROUND(nvl(Z12_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(Y12_1,0)/1000,3),'999999.00') + to_char(ROUND(nvl(X12_1,0)/1000,3),'999999.00')) P12,decode(substr(st15,1,1),'S',st06,'F','6','5') Zone,st15 Area,a0902 AreaName,st01,0 Obj1,0 Obj2,0 Obj3,0 Obj4,0 Obj5,0 Obj6,0 Obj7,0 Obj8,0 Obj9,0 Obj10,0 Obj11,pe04 Obj12 " & _
                                 "from staff,performance, " & _
                                 "(select pe01 saleno from performance where pe03=" & Val(Text1(0)) + 1911 & "12 and pe02='TOT' and pe04>0 union select distinct ax209 from acc020,acc021 where a0205 >= " & Val(Text1(0)) & "1201 and a0205 <= " & Val(Text1(0)) & "1231 and a0201=ax201(+) and a0202=ax202(+) and (substr(ax205,1,1)='4' or ax205='7121') and ax209 is not null) P12, " & _
                                 "(select ax209 W12_0,sum(ax207) W12_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "1201 and a0205 <= " & Val(Text1(0)) & "1231 and ax209 is not null and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and ax207>0 and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) W12, " & _
                                 "(select ax209 X12_0,sum(ax207-ax206) X12_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "1201 and a0205 <= " & Val(Text1(0)) & "1231 and ax209 is not null and (ax205='4191' or ax205='4192' or ax205='4194') group by ax209) X12, " & _
                                 "(select ax209 Y12_0,sum(ax207-ax206) Y12_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "1201 and a0205 <= " & Val(Text1(0)) & "1231 and ax209 is not null and substr(ax205,1,1)='4' and instr(ax213||' ','結餘')>0 group by ax209) Y12, " & _
                                 "(select ax209 Z12_0,sum(-1*ax206) Z12_1 from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202 and a0205 >= " & Val(Text1(0)) & "1201 and a0205 <= " & Val(Text1(0)) & "1231 and ax209 is not null and ax206>0 and (substr(ax205, 1, 1) = '4' Or ax205 = '7121') and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) group by ax209) Z12,acc090 " & _
                                 "where P12.saleno=st01 and W12_0(+)=st01 and X12_0(+)=st01 and Y12_0(+)=st01 and Z12_0(+)=st01 and a0901(+)=st15 and (nvl(W12_1,0)<>0 or nvl(Y12_1,0)<>0 or nvl(X12_1,0)<>0 or nvl(Z12_1,0)<>0 or nvl(pe04,0)<>0) and a0901(+)=st15 " & _
                                 "and pe01(+)=st01 and pe02(+)='TOT' and pe03(+)=" & Val(Text1(0)) + 1911 & "12 "
   End If
   strSql_E = ") " & _
                    "group by 智權人員,Zone,Area,AreaName,st01 " & _
                    "order by Zone,Area,st01 "
   strSql = strSql_S & strSql & strSql_E
   CheckOC3
   GrdDataList.Visible = False
   For ii = GrdDataList.Rows - 1 To 2 Step -1
      GrdDataList.RemoveItem ii
   Next ii
   GrdDataList.Visible = True
   GrdDataList.Clear
   grdDataList2.Clear
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         Set grdDataList2.Recordset = AdoRecordSet3.Clone
         Call SetDataListWidth2
         Calculate
         CmdExcel.Enabled = True
      Else
         MsgBox "無符合資料！", vbInformation
         CmdExcel.Enabled = False
      End If
   End With
   
   doQuery = True
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub Calculate()
Dim stZone As String, stArea As String, stAreaName As String, stZoneName As String
Dim dblSubO1(1 To 14) As Double, dblSubO2(1 To 14) As Double, dblSubO3(1 To 14) As Double
Dim dblSubP1(1 To 14) As Double, dblSubP2(1 To 14) As Double, dblSubP3(1 To 14) As Double
Dim lngColar As Long, cc As Integer, bolRun As Boolean
   
   Erase dblSubO1
   Erase dblSubO2
   Erase dblSubO3
   Erase dblSubP1
   Erase dblSubP2
   Erase dblSubP3
   GrdDataList.FixedCols = 0
   GrdDataList.Visible = False
   
   With grdDataList2
      ii = 1: jj = 0
      Do While ii < .Rows
         'If .TextMatrix(ii, 19) <> "68006" Then   '2015/4/28 cancel by sonia 否則68006不會計入國內合計
            If stZone = "" Then stZone = .TextMatrix(ii, 16)
            If stArea = "" Then stArea = .TextMatrix(ii, 17)
            If stAreaName = "" Then stAreaName = .TextMatrix(ii, 18)
            '區小計
            If .TextMatrix(ii, 17) <> stArea Then
               '智權人員才要
               If Left(stArea, 1) = "S" Then
                  jj = jj + 1
                  GrdDataList.AddItem "", jj
                  GrdDataList.TextMatrix(jj, 0) = stAreaName
                  GrdDataList.TextMatrix(jj, 1) = "目　標"
                  For cc = 1 To 14
                     'Modify By Sindy 2013/12/19
                     If cc <= 12 Then
                        '各月目標重新讀取
                        If Format(cc, "00") >= strMonth_S And Format(cc, "00") <= strMonth_E Then
                           dblSubO1(cc) = GetObjPE04(strYear & Format(cc, "00"), stAreaName) '區月目標
                        Else
                           dblSubO1(cc) = 0
                        End If
                        '各月合計
                        dblSubO2(cc) = dblSubO2(cc) + dblSubO1(cc) '所
                        dblSubO3(cc) = dblSubO3(cc) + dblSubO1(cc) '全所
                        '目標合計
                        dblSubO1(13) = dblSubO1(13) + dblSubO1(cc) '區
                        dblSubO2(13) = dblSubO2(13) + dblSubO1(cc) '所
                        dblSubO3(13) = dblSubO3(13) + dblSubO1(cc) '全所
                     ElseIf cc = 14 Then
                        '目標年平均
                        dblSubO1(14) = Round(dblSubO1(13) / dbl_cnt, 2) '區
                        dblSubO2(14) = dblSubO2(14) + dblSubO1(14) '所
                        dblSubO3(14) = dblSubO3(14) + dblSubO1(14) '全所
                     End If
                     '2013/12/19 END
                     GrdDataList.TextMatrix(jj, cc + 1) = IIf(dblSubO1(cc) = 0, "", dblSubO1(cc))
                  Next
                  If stArea = "S31" Or stArea = "S41" Then '台南所,高雄所
                     lngColar = &H90EE90
                  Else
                     lngColar = &H7FFFD4
                  End If
                  SetColor jj, lngColar
                  jj = jj + 1
                  GrdDataList.AddItem "", jj
                  GrdDataList.TextMatrix(jj, 1) = "達成數"
                  For cc = 1 To 14
                     GrdDataList.TextMatrix(jj, cc + 1) = IIf(dblSubP1(cc) = 0, "", Format(dblSubP1(cc), "##,##0.00"))
                  Next
                  If stArea = "S31" Or stArea = "S41" Then '台南所,高雄所
                     lngColar = &H90EE90
                  Else
                     lngColar = &H7FFFD4
                  End If
                  SetColor jj, lngColar
                  jj = jj + 1
                  GrdDataList.AddItem "", jj
                  GrdDataList.TextMatrix(jj, 0) = "合　計"
                  GrdDataList.TextMatrix(jj, 1) = "達成率"
                  For cc = 1 To 14
                     If dblSubO1(cc) <> 0 And dblSubP1(cc) <> 0 Then GrdDataList.TextMatrix(jj, cc + 1) = Format(dblSubP1(cc) / dblSubO1(cc) * 100, "##,##0.00") & "%"
                  Next
                  If stArea = "S31" Or stArea = "S41" Then '台南所,高雄所
                     lngColar = &H90EE90
                  Else
                     lngColar = &H7FFFD4
                  End If
                  SetColor jj, lngColar
               End If
               Erase dblSubO1
               Erase dblSubP1
               stArea = .TextMatrix(ii, 17)
               stAreaName = .TextMatrix(ii, 18)
               '所小計
               If .TextMatrix(ii, 16) <> stZone Then
                  stZoneName = GetZoneName(stZone)
                  If stZoneName <> "" Then
                     jj = jj + 1
                     GrdDataList.AddItem "", jj
                     GrdDataList.TextMatrix(jj, 0) = stZoneName
                     GrdDataList.TextMatrix(jj, 1) = "目　標"
                     For cc = 1 To 14
                        GrdDataList.TextMatrix(jj, cc + 1) = IIf(dblSubO2(cc) = 0, "", dblSubO2(cc))
                     Next
                     If stZone = "5" Then '其他
                        lngColar = &H90EE40
                     Else
                        lngColar = &H90EE90
                     End If
                     SetColor jj, lngColar
                     jj = jj + 1
                     GrdDataList.AddItem "", jj
                     GrdDataList.TextMatrix(jj, 1) = "達成數"
                     For cc = 1 To 14
                        GrdDataList.TextMatrix(jj, cc + 1) = IIf(dblSubP2(cc) = 0, "", Format(dblSubP2(cc), "##,##0.00"))
                     Next
                     If stZone = "5" Then '其他
                        lngColar = &H90EE40
                     Else
                        lngColar = &H90EE90
                     End If
                     SetColor jj, lngColar
                     jj = jj + 1
                     GrdDataList.AddItem "", jj
                     GrdDataList.TextMatrix(jj, 0) = "總　計"
                     GrdDataList.TextMatrix(jj, 1) = "達成率"
                     For cc = 1 To 14
                        If dblSubO2(cc) <> 0 And dblSubP2(cc) <> 0 Then GrdDataList.TextMatrix(jj, cc + 1) = Format(dblSubP2(cc) / dblSubO2(cc) * 100, "##,##0.00") & "%"
                     Next
                     If stZone = "5" Then '其他
                        lngColar = &H90EE40
                     Else
                        lngColar = &H90EE90
                     End If
                     SetColor jj, lngColar
                  End If
                  '全所時,其他後面加印國內部合計
                  If stZone = "5" Then
                     jj = jj + 1
                     GrdDataList.AddItem "", jj
                     GrdDataList.TextMatrix(jj, 1) = "目　標"
                     For cc = 1 To 14
                        GrdDataList.TextMatrix(jj, cc + 1) = IIf(dblSubO3(cc) = 0, "", dblSubO3(cc))
                     Next
                     lngColar = &HFFFF90
                     SetColor jj, lngColar
                     jj = jj + 1
                     GrdDataList.AddItem "", jj
                     GrdDataList.TextMatrix(jj, 0) = "國內部"
                     GrdDataList.TextMatrix(jj, 1) = "達成數"
                     For cc = 1 To 14
                        GrdDataList.TextMatrix(jj, cc + 1) = IIf(dblSubP3(cc) = 0, "", Format(dblSubP3(cc), "##,##0.00"))
                     Next
                     lngColar = &HFFFF90
                     SetColor jj, lngColar
                     jj = jj + 1
                     GrdDataList.AddItem "", jj
                     GrdDataList.TextMatrix(jj, 1) = "達成率"
                     For cc = 1 To 14
                        If dblSubO3(cc) <> 0 And dblSubP3(cc) <> 0 Then GrdDataList.TextMatrix(jj, cc + 1) = Format(dblSubP3(cc) / dblSubO3(cc) * 100, "##,##0.00") & "%"
                     Next
                     lngColar = &HFFFF90
                     SetColor jj, lngColar
                  End If
                  Erase dblSubO2
                  Erase dblSubP2
                  stZone = .TextMatrix(ii, 16)
               End If
            End If
            
            If Val(.TextMatrix(ii, 14)) = 0 And Val(.TextMatrix(ii, 32)) = 0 Then
            Else
               If .TextMatrix(ii, 16) <> "5" Then
                  '目　標
                  jj = jj + 1
                  If jj <> 1 Then GrdDataList.AddItem "", jj
                  GrdDataList.TextMatrix(jj, 0) = .TextMatrix(ii, 0)
                  GrdDataList.TextMatrix(jj, 1) = .TextMatrix(ii, 1)
                  For cc = 2 To 15
                     GrdDataList.TextMatrix(jj, cc) = IIf(Val(.TextMatrix(ii, cc)) = 0, "", .TextMatrix(ii, cc))
                  Next
                  '達成數
                  jj = jj + 1
                  GrdDataList.AddItem "", jj
                  GrdDataList.TextMatrix(jj, 1) = "達成數"
                  For cc = 2 To 15
                     GrdDataList.TextMatrix(jj, cc) = IIf(Val(.TextMatrix(ii, cc + 18)) = 0, "", Format(.TextMatrix(ii, cc + 18), "##,##0.00"))
                  Next
                  '達成率
                  jj = jj + 1
                  GrdDataList.AddItem "", jj
                  GrdDataList.TextMatrix(jj, 1) = "達成率"
                  For cc = 2 To 15
                     If Val(.TextMatrix(ii, cc + 18)) <> 0 And Val(.TextMatrix(ii, cc)) <> 0 Then GrdDataList.TextMatrix(jj, cc) = Format(Val(.TextMatrix(ii, cc + 18)) / Val(.TextMatrix(ii, cc)) * 100, "##,##0.00") & "%"
                  Next
               End If
               '累計欄位值
               For cc = 2 To 15
                  If .TextMatrix(ii, 16) = "5" Or .TextMatrix(ii, 16) = "6" Then 'Modify By Sindy 2013/12/19 +if
                     dblSubO1(cc - 1) = dblSubO1(cc - 1) + Val(.TextMatrix(ii, cc)) '區
                     dblSubO2(cc - 1) = dblSubO2(cc - 1) + Val(.TextMatrix(ii, cc)) '所
                     dblSubO3(cc - 1) = dblSubO3(cc - 1) + Val(.TextMatrix(ii, cc)) '全所
                  End If
                  dblSubP1(cc - 1) = dblSubP1(cc - 1) + Val(.TextMatrix(ii, cc + 18)) '區
                  dblSubP2(cc - 1) = dblSubP2(cc - 1) + Val(.TextMatrix(ii, cc + 18)) '所
                  dblSubP3(cc - 1) = dblSubP3(cc - 1) + Val(.TextMatrix(ii, cc + 18)) '全所
               Next
            End If
         'End If   '2015/4/28 cancel by sonia 否則68006不會計入國內合計
         ii = ii + 1
      Loop
      '國外部
      If stZone = "6" Then
         stZoneName = GetZoneName(stZone)
         If stZoneName <> "" Then
            jj = jj + 1
            GrdDataList.AddItem "", jj
            GrdDataList.TextMatrix(jj, 1) = "目　標"
            For cc = 1 To 14
               GrdDataList.TextMatrix(jj, cc + 1) = IIf(dblSubO2(cc) = 0, "", dblSubO2(cc))
            Next
            lngColar = &HFFFF90
            SetColor jj, lngColar
            jj = jj + 1
            GrdDataList.AddItem "", jj
            GrdDataList.TextMatrix(jj, 0) = stZoneName
            GrdDataList.TextMatrix(jj, 1) = "達成數"
            For cc = 1 To 14
               GrdDataList.TextMatrix(jj, cc + 1) = IIf(dblSubP2(cc) = 0, "", Format(dblSubP2(cc), "##,##0.00"))
            Next
            lngColar = &HFFFF90
            SetColor jj, lngColar
            jj = jj + 1
            GrdDataList.AddItem "", jj
            GrdDataList.TextMatrix(jj, 1) = "達成率"
            For cc = 1 To 14
               If dblSubO2(cc) <> 0 And dblSubP2(cc) <> 0 Then GrdDataList.TextMatrix(jj, cc + 1) = Format(dblSubP2(cc) / dblSubO2(cc) * 100, "##,##0.00") & "%"
            Next
            lngColar = &HFFFF90
            SetColor jj, lngColar
         End If
         '全所
         jj = jj + 1
         GrdDataList.AddItem "", jj
         GrdDataList.TextMatrix(jj, 1) = "目　標"
         For cc = 1 To 14
            GrdDataList.TextMatrix(jj, cc + 1) = IIf(dblSubO3(cc) = 0, "", dblSubO3(cc))
         Next
         lngColar = &HFFFF00
         SetColor jj, lngColar
         jj = jj + 1
         GrdDataList.AddItem "", jj
         GrdDataList.TextMatrix(jj, 0) = "全　所"
         GrdDataList.TextMatrix(jj, 1) = "達成數"
         For cc = 1 To 14
            GrdDataList.TextMatrix(jj, cc + 1) = IIf(dblSubP3(cc) = 0, "", Format(dblSubP3(cc), "##,##0.00"))
         Next
         lngColar = &HFFFF00
         SetColor jj, lngColar
         jj = jj + 1
         GrdDataList.AddItem "", jj
         GrdDataList.TextMatrix(jj, 1) = "達成率"
         For cc = 1 To 14
            If dblSubO3(cc) <> 0 And dblSubP3(cc) <> 0 Then GrdDataList.TextMatrix(jj, cc + 1) = Format(dblSubP3(cc) / dblSubO3(cc) * 100, "##,##0.00") & "%"
         Next
         lngColar = &HFFFF00
         SetColor jj, lngColar
      End If
   End With
   '其他明細資料
   bolRun = False
   Erase dblSubP1
   With grdDataList2
      ii = 1
      Do While ii < .Rows
         '2015/4/28 modify by sonia 否則68006不會計入國內合計
         'If .TextMatrix(ii, 19) = "68006" Or .TextMatrix(ii, 16) = "5" Then
         If .TextMatrix(ii, 16) = "5" Then
            bolRun = True
            '達成數
            jj = jj + 1
            GrdDataList.AddItem "", jj
            GrdDataList.TextMatrix(jj, 0) = .TextMatrix(ii, 0)
            GrdDataList.TextMatrix(jj, 1) = "點　數"
            For cc = 2 To 15
               GrdDataList.TextMatrix(jj, cc) = IIf(Val(.TextMatrix(ii, cc + 18)) = 0, "", Format(.TextMatrix(ii, cc + 18), "##,##0.00"))
            Next
            '累計欄位值
            For cc = 2 To 15
               dblSubP1(cc - 1) = dblSubP1(cc - 1) + Val(.TextMatrix(ii, cc + 18))
            Next
         End If
         ii = ii + 1
      Loop
   End With
   If bolRun = True Then
      jj = jj + 1
      GrdDataList.AddItem "", jj
      GrdDataList.TextMatrix(jj, 0) = "合　計"
      GrdDataList.TextMatrix(jj, 1) = "點　數"
      For cc = 1 To 14
         GrdDataList.TextMatrix(jj, cc + 1) = IIf(dblSubP1(cc) = 0, "", Format(dblSubP1(cc), "##,##0.00"))
      Next
      lngColar = &H90EE40
      SetColor jj, lngColar
   End If
   
   GrdDataList.Visible = True
   Call SetDataListWidth
   GrdDataList.FixedCols = 2
End Sub

'Add By Sindy 2013/12/19 目標小計
Private Function GetObjPE04(strMonth As String, strKey As String) As String
   GetObjPE04 = 0
   strMonth = Val(strMonth) + 191100
   strExc(0) = "SELECT MAX(DECODE(ST15,'F41','國外','F11','國外','F21','國外','M01','其他','P29','巨京',A0902)) C1," & _
               " NVL(SUM(PE04),0) C2," & _
               " DECODE( ST15,'F41',5,'F11',5,'F21',5,'M01',2,'P29',4,1) C3," & _
               " DECODE(ST15,'F11','F41','F21','F41',ST15) C4,'X' C5" & _
               " From PERFORMANCE, STAFF, ACC090" & _
               " WHERE PE02='TOT'" & _
               " AND PE03=" & strMonth & " AND ST01(+)=PE01 AND A0901(+)=ST15" & _
               " GROUP BY DECODE(ST15,'F11','F41','F21','F41',ST15),DECODE( ST15,'F41',5,'F11',5,'F21',5,'M01',2,'P29',4,1)" & _
               " Having NVL(Sum(PE04), 0) > 0" & _
               " and MAX(DECODE(ST15,'F41','國外','F11','國外','F21','國外','M01','其他','P29','巨京',A0902))='" & strKey & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      GetObjPE04 = RsTemp.Fields("C2")
   End If
End Function

Private Sub SetColor(iRow As Integer, lngColar As Long)
   Dim ii As Integer, jj As Integer
   With GrdDataList
      .row = iRow
      For jj = 0 To .Cols - 1
         .col = jj: .CellBackColor = lngColar
      Next
   End With
End Sub

Private Function GetZoneName(p_Zone As String) As String
   Select Case p_Zone
      Case "1"
         GetZoneName = "北　區"
      Case "2"
         GetZoneName = "中　區"
      Case "5"
         GetZoneName = "其　他"
      Case "6"
         GetZoneName = "國外部"
   End Select
End Function

Private Sub cmdExit_Click()
   Unload Me
End Sub
