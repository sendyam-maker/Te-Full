VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm050206_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "互惠代理人目標給案未輸入明細表"
   ClientHeight    =   5730
   ClientLeft      =   165
   ClientTop       =   960
   ClientWidth     =   7530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   7530
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印(P)"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   4950
      TabIndex        =   6
      Top             =   60
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   6210
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   60
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4845
      Left            =   90
      TabIndex        =   1
      Top             =   780
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   8546
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   7
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
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
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblPeriod 
      BackColor       =   &H8000000E&
      Height          =   180
      Left            =   855
      TabIndex        =   5
      Top             =   510
      Width           =   435
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "期間：             ( 1:上半年 2:下半年 )"
      Height          =   180
      Left            =   225
      TabIndex        =   4
      Top             =   510
      Width           =   2730
   End
   Begin VB.Label lblYear 
      BackColor       =   &H8000000E&
      Height          =   180
      Left            =   855
      TabIndex        =   3
      Top             =   240
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "年度：             (民國年)"
      Height          =   180
      Left            =   225
      TabIndex        =   2
      Top             =   240
      Width           =   1785
   End
End
Attribute VB_Name = "frm050206_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/14 改成Form2.0 ; grdDataList改字型=新細明體-ExtB ; Printer列印未改
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
'Create by Morgan 2008/4/15
Option Explicit
Dim m_iCol As Integer, m_iRow As Integer
Dim PLeft() As Integer, iPrint As Integer, iPage As Integer
Private Const ciTitleFontSize = 22, ciFontSize = 12
Private Const ciStartX = 500, ciStartY = 500, ciColGap = 300
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long

'p_bolHeaderOnly:是否只設定表頭 true=是 false=資料一併清除
Public Sub SetDataListWidth(Optional ByVal p_bolHeaderOnly As Boolean = False)
   Dim ii As Integer
   With grdDataList
      .Visible = False
      .Cols = 5
      If p_bolHeaderOnly = False Then
         .Clear
         .Rows = 2: .FixedRows = 1: .FixedCols = 0
      End If
      .row = 0
      .RowHeightMin = 450
      ii = 0
      .col = ii: .ColWidth(.col) = 650: .Text = "國籍"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignLeftCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 1200: .Text = "代理人編號"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignLeftCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 2400: .Text = "名稱"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignLeftCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 1500: .Text = "聯絡人"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignLeftCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 1300: .Text = "國外部建議量"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      .Refresh
      .Visible = True
   End With
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         frm050206.Show
         Unload Me
      Case 1
         DoPrint
   End Select
   
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   lblYear.BackColor = &H8000000F
   lblPeriod.BackColor = &H8000000F
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm050206_1 = Nothing
End Sub

Private Sub DoPrint()
   Dim iOrientation As Integer, iRow As Integer, iCol As Integer
   Dim strTemp(1 To 5) As String
   
   iOrientation = Printer.Orientation
   Printer.PaperSize = 9
   Printer.Orientation = 2
   lngPageHeight = Printer.ScaleHeight
   lngPageWidth = Printer.ScaleWidth
   lngLineHeight = 300
   GetPleft
   With grdDataList
      iPage = 1
      PrintPageHeader
      For iRow = 1 To .Rows - 1
         For iCol = 1 To 5
            Select Case iCol
               Case 1
                  strTemp(iCol) = Left(.TextMatrix(iRow, iCol - 1), 4)
               Case 3
                  strTemp(iCol) = Left(.TextMatrix(iRow, iCol - 1), 50)
               Case 4
                  strTemp(iCol) = Left(.TextMatrix(iRow, iCol - 1), 30)
               Case Else
                  strTemp(iCol) = .TextMatrix(iRow, iCol - 1)
            End Select
         Next
         PrintDetail strTemp
      Next
      Call PrintReportFooter(.Rows - 1)
      Printer.EndDoc
      MsgBox "列印完成！"
   End With
   Printer.Orientation = iOrientation
End Sub

Sub GetPleft()
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   ReDim PLeft(1 To 6)
   PLeft(1) = ciStartX
   PLeft(2) = PLeft(1) + ciColGap + Printer.TextWidth(String(4, "　"))
   PLeft(3) = PLeft(2) + ciColGap + Printer.TextWidth(String(5, "　"))
   PLeft(4) = PLeft(3) + ciColGap + Printer.TextWidth(String(25, "　"))
   PLeft(5) = PLeft(4) + ciColGap + Printer.TextWidth(String(15, "　"))
   PLeft(6) = PLeft(5) + ciColGap + Printer.TextWidth(String(6, "　"))
End Sub

Sub PrintPageHeader()
   Dim strPTmp As String
   iPrint = ciStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = ciTitleFontSize
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   strPTmp = Me.Caption
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   iPrint = iPrint + 500
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   iPrint = iPrint + lngLineHeight
   strPTmp = "年度：" & lblYear & "   " & "期間：" & IIf(lblPeriod = "1", "上半年", "下半年")
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
       
   iPrint = iPrint + lngLineHeight
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   iPrint = iPrint + lngLineHeight
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
    
   iPrint = iPrint + lngLineHeight
   PLine
   iPrint = iPrint + lngLineHeight
   For intI = 1 To 5
      Printer.CurrentX = PLeft(intI)
      Printer.CurrentY = iPrint
      strPTmp = grdDataList.TextMatrix(0, intI - 1)
      Printer.Print strPTmp
   Next
   iPrint = iPrint + lngLineHeight
   PLine
End Sub

Private Sub PrintNewLine(Optional ByVal iExtraLines As Integer = 3)
   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      PLine
      iPage = iPage + 1
      Printer.NewPage
      PrintPageHeader
      iPrint = iPrint + lngLineHeight
   End If
End Sub

Private Sub PLine()
   Printer.Line (ciStartX, iPrint + lngLineHeight / 2)-(lngPageWidth - ciStartX, iPrint + lngLineHeight / 2)
End Sub

Sub PrintDetail(strData() As String)
    Dim iCol As Integer
    PrintNewLine
    For iCol = 1 To 5
      If iCol < 5 Then
         Printer.CurrentX = PLeft(iCol)
         Printer.CurrentY = iPrint
         Printer.Print strData(iCol)
      Else
        Printer.CurrentX = PLeft(iCol + 1) - Printer.TextWidth(strData(iCol)) - ciColGap
        Printer.CurrentY = iPrint
        Printer.Print strData(iCol)
      End If
    Next
End Sub
'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)
    PrintNewLine
    PLine
    iPrint = iPrint + lngLineHeight
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print "合計： " & iRecCount & " 筆"
    Printer.EndDoc
End Sub
