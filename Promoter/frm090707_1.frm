VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090707_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "繪圖超時案件查詢"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   8925
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   1
      Left            =   2235
      MaxLength       =   7
      TabIndex        =   8
      Top             =   210
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   0
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   7
      Top             =   210
      Width           =   900
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   2
      Left            =   7920
      TabIndex        =   5
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&R)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6570
      TabIndex        =   4
      Top             =   60
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Height          =   400
      Index           =   1
      Left            =   5760
      TabIndex        =   2
      Top             =   60
      Width           =   800
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   4680
      Style           =   2  '單純下拉式
      TabIndex        =   1
      Top             =   5250
      Width           =   4080
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   4545
      Left            =   135
      TabIndex        =   3
      Top             =   570
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   8017
      _Version        =   393216
      Cols            =   13
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
      _Band(0).Cols   =   13
   End
   Begin VB.Line Line2 
      X1              =   1770
      X2              =   2925
      Y1              =   345
      Y2              =   345
   End
   Begin VB.Label Label1 
      Caption         =   "收文日期："
      Height          =   180
      Index           =   2
      Left            =   225
      TabIndex        =   6
      Top             =   255
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "印表機："
      Height          =   180
      Index           =   0
      Left            =   3915
      TabIndex        =   0
      Top             =   5310
      Width           =   720
   End
End
Attribute VB_Name = "frm090707_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/28 改成Form2.0 ; GrdDataList改字型=新細明體-ExtB ; Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Created by Morgan 2012/9/25
Option Explicit

Dim PLeft() As Integer, iPrint As Integer, iPage As Integer
Private Const ciTitleFontSize = 22, ciFontSize = 12
Private Const ciStartX = 500, ciStartY = 500, ciColGap = 250
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim strPrinter As String
Dim m_iCols As Integer

Private Sub cmdOK_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Select Case Index
   Case 0
      frm090707.Show
      Unload frm090707_1
   Case 1
      PUB_RestorePrinter Combo1
      DoPrint
      PUB_RestorePrinter strPrinter
   Case 2
      Unload frm090707
      Unload frm090707_1
   End Select
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, Combo1, strPrinter, , , , , True 'Modified by Morgan 2020/10/30 +只顯示有效的印表機參數
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   Set frm090707_1 = Nothing
End Sub

Public Sub SetDataListWidth()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iCol As Integer
   arrGridHeadText = Array("收文日", "本所案號", "案件名稱", "案件性質", "承辦人", "理由", "超時時數", "換算件數", "智權人員", "申請人")
   arrGridHeadWidth = Array(850, 1200, 960, 840, 620, 480, 840, 820, 620, 960)
   GrdDataList.Cols = UBound(arrGridHeadText) + 1
   With GrdDataList
   .Visible = False
   For iCol = 0 To .Cols - 1
      .row = 0
      .col = iCol
      .Text = arrGridHeadText(iCol)
      .ColWidth(iCol) = arrGridHeadWidth(iCol)
      .CellAlignment = flexAlignCenterCenter
      Select Case iCol
         Case 6, 7
            .ColAlignment(iCol) = flexAlignRightCenter
         Case Else
            .ColAlignment(iCol) = flexAlignLeftCenter
      End Select
   Next iCol
   .Visible = True
   End With
End Sub

Private Sub DoPrint()

   Dim iOrientation As Integer, iRow As Integer, iCol As Integer
   Dim strTemp() As String
   
   iOrientation = Printer.Orientation
   Printer.Orientation = 2
   lngPageHeight = Printer.ScaleHeight
   lngPageWidth = Printer.ScaleWidth
   lngLineHeight = 300
   With GrdDataList
      GetPleft
      ReDim strTemp(1 To m_iCols)
      iPage = 1
      PrintPageHeader
      PrintPageHeader1
      For iRow = 1 To .Rows - 1
         For iCol = LBound(strTemp) To UBound(strTemp)
            '案件名稱
            If iCol = 3 Then
               strTemp(iCol) = Left(.TextMatrix(iRow, iCol - 1), 12)
            ElseIf iCol = 10 Then
               strTemp(iCol) = Left(.TextMatrix(iRow, iCol - 1), 10)
            Else
               strTemp(iCol) = .TextMatrix(iRow, iCol - 1)
            End If
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
   m_iCols = 10
   ReDim PLeft(1 To m_iCols + 1)
   PLeft(1) = ciStartX
   PLeft(2) = PLeft(1) + Printer.TextWidth(String(4, "　")) + ciColGap
   PLeft(3) = PLeft(2) + Printer.TextWidth(String(8, "　")) + ciColGap
   PLeft(4) = PLeft(3) + Printer.TextWidth(String(12, "　")) + ciColGap
   PLeft(5) = PLeft(4) + Printer.TextWidth(String(4, "　")) + ciColGap
   PLeft(6) = PLeft(5) + Printer.TextWidth(String(3, "　")) + ciColGap
   PLeft(7) = PLeft(6) + Printer.TextWidth(String(2, "　")) + ciColGap
   PLeft(8) = PLeft(7) + Printer.TextWidth(String(4, "　")) + ciColGap
   PLeft(9) = PLeft(8) + Printer.TextWidth(String(4, "　")) + ciColGap
   PLeft(10) = PLeft(9) + Printer.TextWidth(String(3, "　")) + ciColGap
   PLeft(11) = PLeft(10) + Printer.TextWidth(String(10, "　")) + ciColGap
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

Sub PrintDetail(strData() As String)
    Dim iCol As Integer
    PrintNewLine
    For iCol = LBound(strData) To UBound(strData)
      Select Case iCol
         Case 7, 8
            Printer.CurrentX = PLeft(iCol + 1) - Printer.TextWidth(strData(iCol)) - ciColGap
            Printer.CurrentY = iPrint
            Printer.Print strData(iCol)
         Case Else
            Printer.CurrentX = PLeft(iCol)
            Printer.CurrentY = iPrint
            Printer.Print strData(iCol)
      End Select
    Next
End Sub

Sub PrintPageHeader()
   Dim strPTmp As String
   iPrint = ciStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = ciTitleFontSize
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   strPTmp = "繪圖超時案件明細表"
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   iPrint = iPrint + 500
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   PrintNewLine
   strPTmp = "收文日：" & txt1(0) & " - " & txt1(1)
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   PrintNewLine
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
    
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   PrintLine
End Sub

Sub PrintPageHeader1()

   Call PrintNewLine(False, 1)
    For intI = 1 To m_iCols
      Select Case intI
         Case 6, 7
            Printer.CurrentX = PLeft(intI + 1) - ciColGap - Printer.TextWidth(GrdDataList.TextMatrix(0, intI - 1))
            Printer.CurrentY = iPrint
            Printer.Print GrdDataList.TextMatrix(0, intI - 1)
         Case Else
            Printer.CurrentX = PLeft(intI)
            Printer.CurrentY = iPrint
            Printer.Print GrdDataList.TextMatrix(0, intI - 1)
      End Select
    Next
   
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   PrintLine
    
End Sub
'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)

    Call PrintNewLine(True, 1)
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    PrintLine
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print "合計： " & iRecCount & " 筆"
    Printer.EndDoc
End Sub
