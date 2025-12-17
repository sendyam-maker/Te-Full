VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm010030_1 
   Caption         =   "分所寄件統計-明細"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5235
   ScaleWidth      =   8010
   StartUpPosition =   3  '系統預設值
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   900
      TabIndex        =   7
      Top             =   4890
      Width           =   7005
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6705
      TabIndex        =   1
      Top             =   90
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印非直寄(&P)"
      Height          =   400
      Left            =   5300
      TabIndex        =   0
      Top             =   90
      Width           =   1300
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4260
      Left            =   90
      TabIndex        =   2
      Top             =   570
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   7514
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "本所案號|案件名稱|案件性質|發文室發文時間"
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
      _Band(0).Cols   =   4
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "印表機："
      Height          =   180
      Left            =   135
      TabIndex        =   8
      Top             =   4950
      Width           =   720
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      Caption         =   "件數"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   1890
      TabIndex        =   6
      Top             =   330
      Width           =   390
   End
   Begin VB.Label lblZone 
      AutoSize        =   -1  'True
      Caption         =   "所別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   765
      TabIndex        =   5
      Top             =   330
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "所別："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   330
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "件數："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   5
      Left            =   1305
      TabIndex        =   3
      Top             =   330
      Width           =   585
   End
End
Attribute VB_Name = "frm010030_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Morgan 2014/5/13
Option Explicit
'列印用
Dim PLeft() As Integer, iPrint As Integer, iPage As Integer
Private Const ciTitleFontSize = 22, ciFontSize = 12
Private Const ciStartX = 500, ciStartY = 500, ciColGap = 250
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim strPrinter As String
Dim iCount As Integer    'add by sonia 2014/7/3

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SetGrid True
   PUB_SetPrinter Me.Name, cmbPrinter, strPrinter
   
End Sub

Private Sub SetGrid(Optional pReset As Boolean = False)
   Dim iCol As Integer
   Dim arrGridHeadWidth
   Dim iUbound As Integer

   '2014/7/9 modify by sonia 加發函方式600
   arrGridHeadWidth = Array(1200, 2600, 1350, 1700, 600)
   iUbound = UBound(arrGridHeadWidth)
   
   With MSHFlexGrid1
   If pReset = True Then
      .Clear
      .Rows = 2
   End If
   .FixedCols = 0
   '2014/7/9 modify by sonia 加發函方式
   .FormatString = "本所案號|案件名稱|案件性質|發文室發文時間|方式"
   For iCol = 0 To .Cols - 1
      If iCol <= iUbound Then
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         .ColAlignment(iCol) = flexAlignLeftCenter
      Else
         .ColWidth(iCol) = 0
      End If
   Next
   End With
End Sub

Public Sub SetData(ByRef rsQuery As ADODB.Recordset)
   Set MSHFlexGrid1.Recordset = rsQuery
   lblCount = MSHFlexGrid1.Rows - 1
   SetGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm010030_1 = Nothing
End Sub

Private Sub cmdPrint_Click()
   PUB_RestorePrinter cmbPrinter
   DoPrint
   PUB_RestorePrinter strPrinter
End Sub

Private Sub DoPrint()

   Dim iOrientation As Integer, iRow As Integer
   
   iOrientation = Printer.Orientation
   Printer.Orientation = 1
   lngPageHeight = Printer.ScaleHeight
   lngPageWidth = Printer.ScaleWidth
   lngLineHeight = 300
   GetPleft
   iPage = 1
   'PrintPageHeader
   'PrintPageHeader1
   iCount = 0       'add by sonia 2014/7/3
   For iRow = 1 To MSHFlexGrid1.Rows - 1
      'modify by sonia 2014/7/3 只印非直寄
      'PrintDetail iRow
      If MSHFlexGrid1.TextMatrix(iRow, 4) <> "直寄" Then PrintDetail iRow
      'end 2014/7/3
   Next
   If iCount > 0 Then          '2014/7/9 add by sonia 有非直寄資料才印
      MsgBox "列印完成！"
      PrintReportFooter
      Printer.EndDoc
   '2014/7/9 add by sonia 有非直寄資料才印
   Else
      MsgBox "無非直寄資料可列印！"
   End If
   'end 2014/7/9
   Printer.Orientation = iOrientation
End Sub

Private Sub GetPleft()
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   ReDim PLeft(MSHFlexGrid1.Cols)
   PLeft(1) = ciStartX
   PLeft(2) = PLeft(1) + Printer.TextWidth(String(7, "　")) + ciColGap
   PLeft(3) = PLeft(2) + Printer.TextWidth(String(16, "　")) + ciColGap
   PLeft(4) = PLeft(3) + Printer.TextWidth(String(8, "　")) + ciColGap
End Sub

Private Sub PrintNewLine(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)

   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint
      PrintLine
      iPage = iPage + 1
      Printer.NewPage
      PrintPageHeader
      If bolSubtotal Then
         PrintPageHeader1
         iPrint = iPrint + lngLineHeight
      End If
   End If
    
End Sub

Private Sub PrintLine()
   Dim iNo As Integer
   iNo = (Printer.ScaleWidth - Printer.CurrentX - 500) \ Printer.TextWidth("-")
   Printer.Print String(iNo, "-")
End Sub

Sub PrintPageHeader()
   Dim strPTmp As String
   iPrint = ciStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = ciTitleFontSize
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   strPTmp = "分所寄件明細 - " & lblZone & "所(非直寄)"
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   
   iPrint = iPrint + 500
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False

   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   PrintNewLine
   
   strPTmp = "發文室發文日：" & frm010030.txt1(0).Tag & " - " & frm010030.txt1(1).Tag
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
    
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   PrintLine
End Sub

Sub PrintPageHeader1()
   Dim iCol As Single
   
   Call PrintNewLine(False, 1)
   With MSHFlexGrid1
   'modify by sonia 2014/7/3 lp11不印
   'For iCol = 0 To .Cols - 1
   For iCol = 0 To .Cols - 2
      If .ColWidth(iCol) > 0 Then
         Printer.CurrentX = PLeft(iCol + 1)
         Printer.CurrentY = iPrint
         Printer.Print .TextMatrix(0, iCol)
      End If
   Next
   End With
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   PrintLine
End Sub

Sub PrintDetail(pRow As Integer)
   Dim iCol As Integer
   iCount = iCount + 1       'add by sonia 2014/7/3
   '2014/7/9 add by sonia 有非直寄資料才印
   If iCount = 1 Then
      PrintPageHeader
      PrintPageHeader1
   End If
   'end 2014/7/9
   PrintNewLine
   With MSHFlexGrid1
   'modify by sonia 2014/7/3 lp11不印
   'For iCol = 0 To .Cols - 1
   For iCol = 0 To .Cols - 2
      Printer.CurrentX = PLeft(iCol + 1)
      Printer.CurrentY = iPrint
      '本所案號
      If iCol = 1 Then
         Printer.Print convForm(.TextMatrix(pRow, iCol), 32)
      ElseIf iCol = 2 Then
         Printer.Print convForm(.TextMatrix(pRow, iCol), 16)
      Else
         Printer.Print .TextMatrix(pRow, iCol)
      End If
   Next
   End With
End Sub

'列印表尾
Private Sub PrintReportFooter()

    Call PrintNewLine(True, 1)
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    PrintLine
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    'modify by sonia 2014/7/3 只印非直寄
    'Printer.Print "共計： " & lblCount & " 件"
    Printer.Print "共計： " & iCount & " 件"
    'end 2014/7/3
    'Printer.EndDoc
End Sub
