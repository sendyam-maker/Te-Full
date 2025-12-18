VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm060208 
   BorderStyle     =   1  '單線固定
   Caption         =   "翻譯費用及請款明細查詢/列印"
   ClientHeight    =   5352
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5352
   ScaleWidth      =   8952
   Begin VB.Frame Frame1 
      Caption         =   "註記"
      Height          =   1395
      Left            =   3780
      TabIndex        =   18
      Top             =   2460
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton Command2 
         Caption         =   "取消"
         Height          =   315
         Index           =   1
         Left            =   3840
         TabIndex        =   21
         Top             =   1020
         Width           =   885
      End
      Begin VB.CommandButton Command2 
         Caption         =   "確定"
         Height          =   315
         Index           =   0
         Left            =   2910
         TabIndex        =   20
         Top             =   1020
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Height          =   795
         Left            =   90
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   210
         Width           =   4635
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "註記(&T)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   7050
      TabIndex        =   17
      Top             =   90
      Width           =   930
   End
   Begin VB.TextBox txtLang 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1350
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1170
      Width           =   345
   End
   Begin VB.TextBox txtPercent 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1350
      MaxLength       =   2
      TabIndex        =   3
      Top             =   840
      Width           =   1065
   End
   Begin VB.TextBox txtCP14 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1350
      MaxLength       =   6
      TabIndex        =   2
      Top             =   510
      Width           =   1065
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5010
      TabIndex        =   10
      Top             =   1125
      Width           =   3810
   End
   Begin VB.TextBox txtNo 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2610
      MaxLength       =   7
      TabIndex        =   1
      Top             =   180
      Width           =   1065
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   6000
      TabIndex        =   6
      Top             =   90
      Width           =   1020
   End
   Begin VB.TextBox txtNo 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1350
      MaxLength       =   7
      TabIndex        =   0
      Top             =   180
      Width           =   1065
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4920
      TabIndex        =   5
      Top             =   90
      Width           =   990
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   8010
      TabIndex        =   7
      Top             =   90
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3795
      Left            =   135
      TabIndex        =   9
      Top             =   1500
      Width           =   8685
      _ExtentX        =   15325
      _ExtentY        =   6689
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "本所案號|日文字|中文字|翻譯人員|翻譯費[A]|請款翻譯費[B]|打字費[C]|折扣|費用比[A/(B+C)]|註記|備註|收文號"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   12
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "異常比率: 日文35%以上, 英文及德文40%以上"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   5010
      TabIndex        =   16
      Top             =   690
      Width           =   3765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "語系：                       (1.日文 2.英文及德文等)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   15
      Top             =   1215
      Width           =   3615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "異常百分比：                          %以上"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   14
      Top             =   885
      Width           =   2880
   End
   Begin VB.Label lblName 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2475
      TabIndex        =   13
      Top             =   555
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "翻譯人員："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   12
      Top             =   555
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "印表機"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   4365
      TabIndex        =   11
      Top             =   1185
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "請款日期：                            －"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   225
      Width           =   2430
   End
End
Attribute VB_Name = "frm060208"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/04/08 Form2.0已修改 (Printer改以Excel印)
'Created by Morgan 2015/8/6
Option Explicit
Dim PLeft() As Integer, iPrint As Integer, iPage As Integer
Private Const ciTitleFontSize = 22, ciFontSize = 12
Private Const ciStartX = 500, ciStartY = 500, ciColGap = 250
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim bolBarShow As Boolean
Dim strPrinter As String
Dim m_iCols As Integer
Dim iPrevRow2 As Integer 'Added by Morgan 2019/5/30
'Add by Amy 2022/04/06
Dim strField, intWidth
Dim i As Integer, intField As Integer, intR As Integer, intTitleRow As Integer

Private Sub Command1_Click()
   If iPrevRow2 > 0 Then
      grdDataList.Enabled = False
      Frame1.Top = cmdExit.Top
      Frame1.Left = cmdExit.Left + cmdExit.Width - Frame1.Width
      Frame1.Visible = True
      Text1.Text = grdDataList.TextMatrix(iPrevRow2, 9)
      Text1.SetFocus
      TextInverse Text1
   End If
End Sub

Private Sub Command2_Click(Index As Integer)
   '確定
   If Index = 0 Then
      cnnConnection.Execute "update transfee set tf35='" & ChgSQL(Text1.Text) & "' where tf01='" & grdDataList.TextMatrix(iPrevRow2, 11) & "'", intI
      grdDataList.TextMatrix(iPrevRow2, 9) = Text1
   End If
   Frame1.Visible = False
   grdDataList.Enabled = True
End Sub

Private Sub Form_Activate()
   bolBarShow = Forms(0).StatusBar1.Visible
   Forms(0).StatusBar1.Visible = True
End Sub

Private Sub Form_Deactivate()
   Forms(0).StatusBar1.Visible = bolBarShow
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Mark by Amy 2022/04/08 不使用
Private Sub DoPrint()
'   Dim iOrientation As Integer, iRow As Integer, iCol As Integer
'   Dim strTemp() As String
'
'   iOrientation = Printer.Orientation
'   Printer.Orientation = 2
'   lngPageHeight = Printer.ScaleHeight
'   lngPageWidth = Printer.ScaleWidth
'   lngLineHeight = 300
'   With grdDataList
'      GetPleft
'      ReDim strTemp(1 To m_iCols - 1)
'      iPage = 1
'      PrintPageHeader
'      PrintPageHeader1
'      For iRow = 1 To .Rows - 1
'         For iCol = LBound(strTemp) To UBound(strTemp)
'            strTemp(iCol) = .TextMatrix(iRow, iCol - 1)
'         Next
'         PrintDetail strTemp
'      Next
'      Call PrintReportFooter(.Rows - 1)
'      Printer.EndDoc
'      MsgBox "列印完成！"
'   End With
'   Printer.Orientation = iOrientation
End Sub

Sub GetPleft()
'   Printer.Font.Size = ciFontSize
'   Printer.Font.Bold = False
'   Printer.Font.Underline = False
'   intI = m_iCols + 1
'   ReDim PLeft(1 To intI)
'   'Modified by Morgan 2019/5/29 +打字費,調整寬度
'   PLeft(1) = ciStartX
'   PLeft(2) = PLeft(1) + Printer.TextWidth(String(5, "　")) + ciColGap
'   PLeft(3) = PLeft(2) + Printer.TextWidth(String(3, "　")) + ciColGap
'   PLeft(4) = PLeft(3) + Printer.TextWidth(String(3, "　")) + ciColGap
'   PLeft(5) = PLeft(4) + Printer.TextWidth(String(4, "　")) + ciColGap
'   PLeft(6) = PLeft(5) + Printer.TextWidth(String(4, "　")) + ciColGap
'   PLeft(7) = PLeft(6) + Printer.TextWidth(String(6, "　")) + ciColGap
'   PLeft(8) = PLeft(7) + Printer.TextWidth(String(4, "　")) + ciColGap
'   PLeft(9) = PLeft(8) + Printer.TextWidth(String(3, "　")) + ciColGap
'   PLeft(10) = PLeft(9) + Printer.TextWidth(String(7, "　")) + ciColGap
'   PLeft(11) = PLeft(10) + Printer.TextWidth(String(4, "　")) + ciColGap
End Sub

Private Sub PrintNewLine(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)

'   iPrint = iPrint + lngLineHeight
'   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
'      Printer.CurrentX = ciStartX
'      Printer.CurrentY = iPrint
'      PrintLine
'      iPage = iPage + 1
'      Printer.NewPage
'      PrintPageHeader
'      If bolSubtotal Then
'         PrintPageHeader1
'         iPrint = iPrint + lngLineHeight
'      End If
'   End If
    
End Sub

Private Sub PrintLine()
'   Dim iNo As Integer
'   iNo = (Printer.ScaleWidth - Printer.CurrentX - 500) \ Printer.TextWidth("-")
'   Printer.Print String(iNo, "-")
End Sub

Sub PrintDetail(strData() As String)
'    Dim iCol As Integer
'    PrintNewLine
'    For iCol = LBound(strData) To UBound(strData)
'      Select Case iCol
'         '靠右
'         Case 2, 3, 5, 6, 7
'            Printer.CurrentX = PLeft(iCol + 1) - Printer.TextWidth(strData(iCol)) - ciColGap
'            Printer.CurrentY = iPrint
'            Printer.Print strData(iCol)
'         '置中
'         Case 8, 9
'            Printer.CurrentX = PLeft(iCol) + (PLeft(iCol + 1) - PLeft(iCol) - ciColGap - Printer.TextWidth(strData(iCol))) / 2
'            Printer.CurrentY = iPrint
'            Printer.Print strData(iCol)
'         '靠左
'         Case Else
'            Printer.CurrentX = PLeft(iCol)
'            Printer.CurrentY = iPrint
'            Printer.Print strData(iCol)
'      End Select
'    Next
End Sub

Sub PrintPageHeader()
'   Dim strPTmp As String
'   iPrint = ciStartY
'   Printer.FontName = "細明體"
'   Printer.Font.Size = ciTitleFontSize
'   Printer.Font.Bold = True
'   Printer.Font.Underline = True
'   strPTmp = "翻譯費用及請款明細表"
'   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
'   Printer.CurrentY = iPrint
'   Printer.Print strPTmp
'   iPrint = iPrint + 500
'   Printer.Font.Size = ciFontSize
'   Printer.Font.Bold = False
'   Printer.Font.Underline = False
'
'   PrintNewLine
'   strPTmp = "請款日期："
'   Printer.CurrentX = (lngPageWidth) / 2 - Printer.TextWidth(String(8, "　"))
'   Printer.CurrentY = iPrint
'   Printer.Print strPTmp & CFDate(txtNO(0).Tag) & " － " & IIf(txtNO(1) <> "", CFDate(txtNO(1).Tag), "")
'
'   PrintNewLine
'
'   If txtCP14.Tag <> "" Then
'      strPTmp = "翻譯人員："
'      Printer.CurrentX = (lngPageWidth) / 2 - Printer.TextWidth(String(8, "　"))
'      Printer.CurrentY = iPrint
'      Printer.Print strPTmp & txtCP14.Tag & " " & lblName.Tag
'      PrintNewLine
'   End If
'
'   If txtPercent.Tag <> "" Then
'      strPTmp = "異常百分比："
'      Printer.CurrentX = (lngPageWidth) / 2 - Printer.TextWidth(String(8, "　"))
'      Printer.CurrentY = iPrint
'      Printer.Print strPTmp & txtPercent.Tag & "%以上"
'      PrintNewLine
'   End If
'
'   Printer.CurrentX = ciStartX
'   Printer.CurrentY = iPrint
'   Printer.Print "列印人：" & strUserName
'
'   If txtLang.Tag <> "" Then
'      strPTmp = "語系："
'      Printer.CurrentX = (lngPageWidth) / 2 - Printer.TextWidth(String(8, "　"))
'      Printer.CurrentY = iPrint
'      Printer.Print strPTmp & IIf(txtCP14.Tag = "1", "日文", "英文及德文等")
'   End If
'
'
'   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
'   Printer.CurrentY = iPrint
'   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
'
'   PrintNewLine
'   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
'   Printer.CurrentY = iPrint
'   Printer.Print "頁    次：" & str(iPage)
'
'   PrintNewLine
'   Printer.CurrentX = ciStartX
'   Printer.CurrentY = iPrint
'   PrintLine
End Sub

Sub PrintPageHeader1()

'    Call PrintNewLine(False, 1)
'    For intI = 1 To m_iCols - 1
'      Select Case intI
'         Case 2, 3, 5, 6, 7
'            Printer.CurrentX = PLeft(intI + 1) - ciColGap - Printer.TextWidth(grdDataList.TextMatrix(0, intI - 1))
'            Printer.CurrentY = iPrint
'            Printer.Print grdDataList.TextMatrix(0, intI - 1)
'         Case Else
'            Printer.CurrentX = PLeft(intI)
'            Printer.CurrentY = iPrint
'            Printer.Print grdDataList.TextMatrix(0, intI - 1)
'      End Select
'    Next
'    PrintNewLine
'    Printer.CurrentX = ciStartX
'    Printer.CurrentY = iPrint
'    PrintLine
End Sub

'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)

'    Call PrintNewLine(True, 1)
'    Printer.CurrentX = ciStartX
'    Printer.CurrentY = iPrint
'    PrintLine
'    PrintNewLine
'    Printer.CurrentX = ciStartX
'    Printer.CurrentY = iPrint
'    Printer.Print "合計： " & iRecCount & " 筆"
'    Printer.EndDoc
End Sub
'end 2022/04/08

Private Sub cmdPrint_Click()
   If Not grdDataList.Recordset Is Nothing Then
      If grdDataList.Recordset.RecordCount > 0 Then
         'Modify by Amy 2022/04/08 改以Excel印
         'PUB_RestorePrinter Combo1
         'DoPrint
         'PUB_RestorePrinter strPrinter
         PUB_SetOsDefaultPrinter Combo1
         ExcelSave
         PUB_SetOsDefaultPrinter strPrinter
         'end 2022/04/08
      End If
   End If
End Sub

'Add by Amy 2022/04/06 以Excel印
Private Function ExcelSave() As Boolean
    Dim Xls As New Excel.Application, Wks As New Worksheet
    Dim strWkName As String, strFileName As String, strFormat As String '工作表名稱為中文/檔案名稱/儲存格格式
    Dim iCol As Integer, intAlign As Integer, iMaxLen As Single, strTmp As String
On Error GoTo ErrHnd

    intField = 65:  intR = 1
    strFileName = "翻譯費用及請款明細表" & MsgText(43)
    If Dir(strExcelPath & strFileName) = MsgText(601) Then
        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
        End If
    Else
        Kill strExcelPath & strFileName
    End If
    
    Xls.SheetsInNewWorkbook = 3
    Xls.Workbooks.add
    'Xls.Visible = True
    '工作表名稱改為中文
    If strWkName = MsgText(601) Then strWkName = Left(Xls.Worksheets(1).Name, Len(Xls.Worksheets(1).Name) - 1)
    Set Wks = Xls.Worksheets(strWkName & "1")
    Call SetTitle(Wks)
    
    With Wks
        For i = 1 To grdDataList.Rows - 1
            For iCol = LBound(strField) To UBound(strField)
                strFormat = "": strTmp = "": iMaxLen = 0
                intAlign = 1 '對齊方式,預設靠左
                Select Case strField(iCol)
                    Case "日文字", "翻譯費[A]", "請款翻譯費[B]", "打字費[C]"
                        intAlign = 2 '靠右
                        strFormat = "#,##0"
                    Case "翻譯人員"
                        iMaxLen = 8
                    Case "折扣"
                        strFormat = "0%"
                    Case "費用比[A/(B+C)]"
                        intAlign = 3 '置中
                        strFormat = "0%"
                    Case "註記"
                        intAlign = 3 '置中
                        iMaxLen = 8
                    Case "備註"
                        iMaxLen = 30
                End Select
                strTmp = grdDataList.TextMatrix(i, iCol)
                If iMaxLen > 0 Then
                    strTmp = PUB_StrToStr(strTmp, iMaxLen)
                End If
                If strFormat <> MsgText(601) Then
                    .Range(Chr(intField + iCol) & intR).NumberFormatLocal = strFormat
                End If
                If intAlign = 2 Then
                    .Range(Chr(intField + iCol) & intR).HorizontalAlignment = xlRight
                End If
                If intAlign = 3 Then
                    .Range(Chr(intField + iCol) & intR).HorizontalAlignment = xlCenter
                End If
                .Range(Chr(intField + iCol) & intR).Value = strTmp
            Next iCol
            intR = intR + 1
        Next i
        .Range(Chr(intField) & intR).Value = "合計：" & grdDataList.Rows - 1 & "筆"
        .Range(Chr(intField) & intR & ":" & Chr(intField + UBound(strField)) & intR).Borders(xlEdgeTop).LineStyle = xlDash '虛線
        .Range(Chr(intField) & intR & ":" & Chr(intField + UBound(strField)) & intR).Borders(xlEdgeTop).Weight = xlThin '細線
    End With
    Wks.PageSetup.PaperSize = 9 'A4
    Wks.PageSetup.CenterFooter = "第 &P 頁，共 &N 頁"
    Wks.PageSetup.Orientation = xlLandscape '直印
    Wks.PageSetup.PrintTitleRows = "$1:$" & intTitleRow '標題列
    Wks.PageSetup.LeftMargin = Xls.InchesToPoints(0.5) '邊界
    Wks.PageSetup.RightMargin = Xls.InchesToPoints(0.5)
    Wks.PageSetup.TopMargin = Xls.InchesToPoints(0.5)
    Wks.PageSetup.BottomMargin = Xls.InchesToPoints(0.5)
    Wks.PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
      
    '判斷版本
    If Val(Xls.Version) < 12 Then
        Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=-4143
    Else
        Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=56
    End If
    Wks.PrintOut Copies:=1, Collate:=True
    Xls.Workbooks.Close
    Xls.Quit
    Set Wks = Nothing
    Set Xls = Nothing
    ExcelSave = True
    Exit Function
  
ErrHnd:
    If Val(Xls.Version) < 12 Then
        Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=-4143
    Else
        Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=56
    End If
    Xls.Workbooks.Close
    Xls.Quit
    Set Wks = Nothing
    Set Xls = Nothing
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub SetTitle(ByRef Wks As Worksheet)
    Dim strTp As String
    ReDim strField(grdDataList.Cols - 2)
    ReDim intWidth(grdDataList.Cols - 2)
  
    For i = 0 To grdDataList.Cols - 2
        strField(i) = grdDataList.TextMatrix(0, i)
        Select Case strField(i)
            Case "本所案號"
                intWidth(i) = 11
            Case "日文字", "中文字"
                intWidth(i) = 7
            Case "翻譯人員", "翻譯費[A]", "打字費[C]"
                intWidth(i) = 9.5
            Case "請款翻譯費[B]"
                intWidth(i) = 13.75
            Case "折扣"
                intWidth(i) = 5
            Case "費用比[A/(B+C)]"
                intWidth(i) = 15.38
            Case "註記"
                intWidth(i) = 10
            Case "備註"
                intWidth(i) = 30
        End Select
    Next i
    
    strTp = GetValue("請款翻譯費[B]") '條件顯示位置
    With Wks
        .Range(Chr(intField) & intR).Font.Size = 22
        .Range(Chr(intField) & intR).Font.Bold = True
        .Range(Chr(intField) & intR).Font.Underline = True
        .Range(Chr(intField) & intR).Value = "翻譯費用及請款明細表"
        .Range(Chr(intField) & intR & ":" & Chr(UBound(strField) + intField) & intR).HorizontalAlignment = xlCenter
        .Range(Chr(intField) & intR & ":" & Chr(UBound(strField) + intField) & intR).MergeCells = True
        intR = intR + 1
        .Range(Chr(intField + Val(strTp)) & intR).Value = "請款日期："
        .Range(Chr(intField + Val(strTp)) & intR).HorizontalAlignment = xlRight
        .Range(Chr(intField + Val(strTp) + 1) & intR).Value = CFDate(txtNo(0).Tag) & " － " & IIf(txtNo(1) <> "", CFDate(txtNo(1).Tag), "")
        .Range(Chr(intField + Val(strTp) + 1) & intR).HorizontalAlignment = xlLeft
        
        If txtCP14.Tag <> MsgText(601) Then
            intR = intR + 1
            .Range(Chr(intField + Val(strTp)) & intR).Value = "翻譯人員："
            .Range(Chr(intField + Val(strTp)) & intR).HorizontalAlignment = xlRight
            .Range(Chr(intField + Val(strTp) + 1) & intR).Value = txtCP14.Tag & " " & lblName.Tag
            .Range(Chr(intField + Val(strTp) + 1) & intR).HorizontalAlignment = xlLeft
        End If
        If txtPercent.Tag <> "" Then
            intR = intR + 1
            .Range(Chr(intField + Val(strTp)) & intR).Value = "異常百分比："
            .Range(Chr(intField + Val(strTp)) & intR).HorizontalAlignment = xlRight
            .Range(Chr(intField + Val(strTp) + 1) & intR).Value = txtPercent.Tag & "%以上"
            .Range(Chr(intField + Val(strTp) + 1) & intR).HorizontalAlignment = xlLeft
        End If
        If txtLang.Tag <> "" Then
            intR = intR + 1
            .Range(Chr(intField + Val(strTp)) & intR).Value = "語系："
            .Range(Chr(intField + Val(strTp)) & intR).HorizontalAlignment = xlRight
            .Range(Chr(intField + Val(strTp) + 1) & intR).Value = IIf(txtCP14.Tag = "1", "日文", "英文及德文等")
            .Range(Chr(intField + Val(strTp) + 1) & intR).HorizontalAlignment = xlLeft
        End If
        
        intR = intR + 1
        .Range(Chr(intField) & intR).Value = "列印人員：" & StaffQuery(strUserNum)
        .Range(Chr(intField + UBound(strField) - 1) & intR).Value = "列印日期：" & CFDate(strSrvDate(2))
        
        intR = intR + 2
        For i = LBound(strField) To UBound(strField)
            .Range(Chr(intField + i) & intR).Value = strField(i)
            .Columns(Chr(intField + i) & ":" & Chr(intField + i)).ColumnWidth = intWidth(i)
        Next i
        '設定格式
        .Range(Chr(intField) & intR).RowHeight = 22 '調整列高
        .Range(Chr(intField) & intR & ":" & Chr(UBound(strField) + intField) & intR).Font.Size = 12
        .Range(Chr(intField) & intR & ":" & Chr(intField + UBound(strField)) & intR).HorizontalAlignment = xlCenter
        .Range(Chr(intField) & intR & ":" & Chr(intField + UBound(strField)) & intR).Borders(xlEdgeTop).LineStyle = xlDash '虛線
        .Range(Chr(intField) & intR & ":" & Chr(intField + UBound(strField)) & intR).Borders(xlEdgeTop).Weight = xlThin
        .Range(Chr(intField) & intR & ":" & Chr(intField + UBound(strField)) & intR).Borders(xlEdgeBottom).LineStyle = xlDash '虛線
        .Range(Chr(intField) & intR & ":" & Chr(intField + UBound(strField)) & intR).Borders(xlEdgeBottom).Weight = xlThin
        intTitleRow = intR
        intR = intR + 1
    End With
End Sub

Private Function GetValue(pFieldN As String) As Integer
    Dim jj As Integer
    
    For jj = 1 To UBound(strField)
        If UCase(strField(jj)) = UCase(pFieldN) Then
            GetValue = jj
            Exit For
        End If
    Next jj
End Function
'end 2022/04/08

Private Sub doQuery()
Dim stCon As String, stCon1 As String
    
On Error GoTo flgErr
    
    txtNo(0).Tag = txtNo(0).Text
    txtNo(1).Tag = txtNo(1).Text
    txtCP14.Tag = txtCP14.Text
    lblName.Tag = lblName.Caption
    txtPercent.Tag = txtPercent
    txtLang.Tag = txtLang
    
    ClearQueryLog Me.Name
   
    '請款日
    If txtNo(0) <> "" Then
      stCon = stCon & " and a1k02>=" & txtNo(0)
    End If
    If txtNo(1) <> "" Then
      stCon = stCon & " and a1k02<=" & txtNo(1)
    End If
    If txtNo(0) <> "" Or txtNo(1) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label1, 5) & txtNo(0) & "-" & txtNo(1)
    End If
    
    '翻譯人員
    If txtCP14 <> "" Then
      stCon = stCon & " and cp14='" & txtCP14 & "'"
      pub_QL05 = pub_QL05 & ";" & Label2(2) & txtCP14 & lblName
    End If
    
    '異常百分比
    If Val(txtPercent) > 0 Then
      'Modified by Morgan 2019/5/29 +請款打字費
      'stCon1 = stCon1 & " and PayAmt>=DebAmt*" & Val(txtPercent) / 100
      stCon1 = stCon1 & " and PayAmt>=(DebAmt+(nvl(a1l05,0)-nvl(a1l07,0)))*" & Val(txtPercent) / 100
      pub_QL05 = pub_QL05 & ";" & Left(Label2(1), 6) & txtPercent & "%以上"
    End If
    
    'Added by Morgan 2019/5/29
    If txtLang = "1" Then
      stCon = stCon & " and tf27='2'"
      pub_QL05 = pub_QL05 & ";日文"
    ElseIf txtLang = "2" Then
      stCon = stCon & " and tf27<>'2'"
      pub_QL05 = pub_QL05 & ";英文及德文等"
    End If
    
   'Modified by Morgan 2019/5/28 +請款打字費,判斷有翻譯費用及請款翻譯費的才要顯示--婧瑄
   strExc(0) = "select CaseNo 本所案號,tf03 日文字數" & _
   ",tf02 中文字數,st02 翻譯人員,to_char(PayAmt,'999,999,999') 翻譯費用,to_char(DebAmt,'999,999,999') 請款翻譯費" & _
   ",to_char(nvl(a1l05,0)-nvl(a1l07,0),'999,999,999') 請款打字費,DRate 請款折扣,decode(PayAmt,0,null,round(100*PayAmt/(DebAmt+(nvl(a1l05,0)-nvl(a1l07,0))))||'%') 費用占百分比" & _
   ",tf35 註記,decode(sign(pa49),1,'個案全部折扣:'||pa49||'%;')||decode(sign(pa50),1,'個案申請/翻譯折扣:'||pa50||'%;')" & _
   "||decode(sign(nvl(fa25,cu36)),1,'請款對象全部折扣:'||nvl(fa25,cu36)||'%;')" & _
   "||decode(sign(nvl(fa26,cu37)),1,'請款對象申請/翻譯折扣:'||nvl(fa26,cu37)||'%;') 備註,cp09 收文號" & _
   " from (select a1k13||'-'||a1k14||decode(a1k15||a1k16,'000','','-'||a1k15||'-'||a1k16) CaseNo" & _
   ",tf03,tf02,st02,round(axf04*nvl(a1906,a2103)) PayAmt,a1l05-nvl(a1l07,0) DebAmt" & _
   ",decode(a1l07,0,null,round(100*(1-a1l07/a1l05))||'%') DRate,cp01,cp02,cp03,cp04,a1k28,a1k01,tf35,cp09" & _
   " From acc1k0, acc1l0, caseprogress, transfee, staff, acc151, acc150, acc190" & _
   ",(select a2102,mod(max(a2101*1000+a2103),1000) a2103 from acc210 group by a2102) x" & _
   " where a1k13 in ('FCP','FG','P','CFP') " & stCon & " and a1l01(+)=a1k01 and a1l04='201'" & _
   " and cp60(+)=a1l01 and cp10(+)=a1l04 and cp12='F23' and cp14 like 'F%' and tf01(+)=cp09 and st01(+)=cp14" & _
   " and cp61 is not null and axf01(+)=cp61 and axf02(+)=cp09 and a1501(+)=axf01 and a1902(+)=a1501 and a2102(+)=a1505" & _
   " Union All select a1k13||'-'||a1k14||decode(a1k15||a1k16,'000','','-'||a1k15||'-'||a1k16) CaseNo,tf03,tf02,st02" & _
   ",a1p07 PayAmt,a1l05-nvl(a1l07,0) DebAmt,decode(a1l07,0,null,round(100*(1-a1l07/a1l05))||'%') DRate,cp01,cp02,cp03,cp04,a1k28,a1k01,tf35,cp09" & _
   " From acc1k0, acc1l0, caseprogress, transfee, staff, Acc1p0" & _
   " where a1k13 in ('FCP','FG','P','CFP') " & stCon & " and a1l01(+)=a1k01 and a1l04='201'" & _
   " and cp60(+)=a1l01 and cp10(+)=a1l04 and cp12='F23' and cp14 like 'F%' and cp61 is null" & _
   " and tf01(+)=cp09 and st01(+)=cp14" & _
   " and a1p04(+)=tf07 and a1p05(+)='6130' and (a1p17 is null or  a1p17=cp01||cp02||cp03||cp04)" & _
   "),acc1l0,patent,fagent,customer where PayAmt>0 and DebAmt>0 and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & stCon1 & _
   " and fa01(+)=substr(a1k28,1,8) and fa02(+)=substr(a1k28,9)" & _
   " and cu01(+)=substr(a1k28,1,8) and cu02(+)=substr(a1k28,9)" & _
   " and a1l01(+)=a1k01 and a1l04(+)='03' order by 1"
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   m_iCols = 12
   SetGrid RsTemp
   RecordShow
   If RsTemp.RecordCount = 0 Then
      InsertQueryLog (0)
      MsgBox "查無資料！", vbInformation
   Else
      InsertQueryLog (RsTemp.RecordCount)
   End If
   
flgErr:
    If Err.Number <> 0 Then
        MsgBox Err.Description
    End If
    
End Sub

Private Sub SetGrid(p_Rst As ADODB.Recordset)
   Dim iCol As Integer
   With grdDataList
      .Visible = False
      iPrevRow2 = 0
      Set .Recordset = p_Rst.Clone
      .FormatString = .FormatString
      .ColAlignment(1) = flexAlignRightCenter
      .ColAlignment(2) = flexAlignRightCenter
      .ColAlignment(4) = flexAlignRightCenter
      .ColAlignment(5) = flexAlignRightCenter
      .ColAlignment(6) = flexAlignRightCenter
      .ColAlignment(7) = flexAlignCenterCenter
      .ColAlignment(8) = flexAlignCenterCenter
      .ColWidth(0) = 1050
      .ColWidth(9) = 1000
      .ColWidth(10) = 3000
      .ColWidth(11) = 0
      .Visible = True
   End With
End Sub

Private Sub cmdQuery_Click()
    Screen.MousePointer = vbHourglass
    If TxtValidate Then doQuery
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   CmdPrint.Enabled = IsUserHasRightOfFunction(Me.Name, strPrint, False)
   PUB_SetPrinter Me.Name, Combo1, strPrinter
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   MenuEnabled
   Set frm060208 = Nothing
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Private Sub RecordShow()
   Forms(0).StatusBar1.Panels(2).Text = grdDataList.Recordset.RecordCount
End Sub

Private Sub GrdDataList_Click()
   Dim nRow As Integer
   nRow = grdDataList.MouseRow
   If nRow > 0 Then
      SelectRow nRow, grdDataList, iPrevRow2
      iPrevRow2 = nRow
   End If
End Sub

Private Sub txtCP14_Change()
   Dim strTempName As String
   lblName = ""
   If Len(txtCP14) >= 5 And txtCP14 < "G" Then
      If ClsPDGetStaff(txtCP14, strTempName) Then
         lblName = strTempName
      End If
   End If
End Sub

Private Sub txtCP14_GotFocus()
   TextInverse txtCP14
End Sub

Private Sub txtCP14_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Function GetIdFromName(ByRef pName As String, ByRef pID As String) As Boolean
   strExc(0) = "select st01,st02 from staff where st02 like '" & ChgSQL(pName) & "%' and st01>'F'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.RecordCount = 1 Then
         pID = RsTemp.Fields("st01")
         pName = RsTemp.Fields("st02")
         GetIdFromName = True
      Else
         MsgBox "員工名稱重複，請直接輸入員工編號！"
      End If
   Else
      MsgBox "該員工名稱不存在！"
   End If
End Function

Private Sub txtCP14_Validate(Cancel As Boolean)
   If txtCP14 >= "G" Then
      strExc(0) = txtCP14
      If GetIdFromName(strExc(0), strExc(1)) Then
         txtCP14 = strExc(1)
         lblName = strExc(0)
      End If
   End If
End Sub

Private Sub txtLang_GotFocus()
   TextInverse txtLang
End Sub

Private Sub txtLang_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtNo_GotFocus(Index As Integer)
   TextInverse txtNo(Index)
End Sub

Private Sub txtNo_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtNo_Validate(Index As Integer, Cancel As Boolean)
   If txtNo(Index) <> "" Then
      If Not ChkDate(txtNo(Index)) Then
        Cancel = True
      End If
   End If
End Sub

Private Function TxtValidate() As Boolean
   
   Dim bolCancel As Boolean
   
   If txtNo(0) = "" Then
      MsgBox "請款日期條件不可空白！", vbExclamation
      txtNo(0).SetFocus
      Exit Function
   End If
   
   bolCancel = False
   Call txtNo_Validate(0, bolCancel)
   If bolCancel Then
      Exit Function
   End If
   
   Call txtNo_Validate(1, bolCancel)
   If bolCancel Then
      Exit Function
   End If
   TxtValidate = True

End Function

Private Sub txtPercent_GotFocus()
   TextInverse txtPercent
End Sub

Private Sub txtPercent_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
   End If
End Sub


Private Sub SelectRow(ByRef pRow As Integer, ByRef FlexGrid As MSHFlexGrid, ByRef pPrevRow As Integer)
   Dim nCol As Integer, iCol As Integer
   With FlexGrid
   nCol = .col
   If pPrevRow > 0 Then
      If pPrevRow <> pRow Then
         .row = pPrevRow
         If .FixedCols > 0 Then
            .col = .FixedCols - 1
            .CellBackColor = .BackColorFixed
            .CellForeColor = .ForeColor
         End If
         For iCol = .FixedCols To .Cols - 1
            .col = iCol
            .CellBackColor = .BackColor
         Next
      End If
   End If

   If pRow > 0 Then
      .row = pRow
      If .FixedCols > 0 Then
         .col = .FixedCols - 1
         .CellBackColor = .BackColorSel
         .CellForeColor = .ForeColorSel
      End If
      For iCol = .FixedCols To .Cols - 1
        .col = iCol
        .CellBackColor = &HFFC0C0
      Next
   End If
   .col = nCol
   .Refresh
   pPrevRow = pRow
   End With
End Sub
