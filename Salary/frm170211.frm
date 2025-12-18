VERSION 5.00
Begin VB.Form frm170211 
   BorderStyle     =   1  '單線固定
   Caption         =   "薪資基本資料檢核表"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   3945
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   975
      Width           =   780
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   1170
      MaxLength       =   6
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   975
      Width           =   780
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   1170
      Style           =   2  '單純下拉式
      TabIndex        =   8
      Top             =   1350
      Width           =   2460
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   2160
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   780
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1170
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   780
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1260
      TabIndex        =   5
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   2340
      TabIndex        =   4
      Top             =   90
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "印表機："
      Height          =   180
      Left            =   360
      TabIndex        =   9
      Top             =   1380
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   7
      Top             =   1020
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部　　門："
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   6
      Top             =   660
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1890
      X2              =   2205
      Y1              =   1110
      Y2              =   1110
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1890
      X2              =   2205
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frm170211"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Create by Morgan 2009/2/4
Option Explicit

Dim PLeft() As Integer, PColName() As String, strTemp() As String
Dim iPrint As Integer, iPage As Integer
Dim m_iStartX As Integer, m_iStartY As Integer, m_iColGap As Integer
Dim m_iPageHeight As Long, m_iLineHeight As Long, m_iMargin As Long
Dim m_DefaultPrinter As String


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
   Case 0
      Screen.MousePointer = vbHourglass
      If FormCheck Then
         If cmbPrinter <> Printer.DeviceName Then PUB_RestorePrinter cmbPrinter
         PrintReport
         If cmbPrinter.Tag <> cmbPrinter Then
            PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, 0, 0, Me.cmbPrinter.Text
         End If
         If Printer.DeviceName <> m_DefaultPrinter Then PUB_RestorePrinter m_DefaultPrinter
      End If
      Screen.MousePointer = vbDefault
   Case 1
      Unload Me
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, cmbPrinter, m_DefaultPrinter
   FormReset
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170211 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   If Index = 1 Then
      If Text1(0) <> "" Then Text1(1) = Text1(0)
   ElseIf Index = 3 Then
      If Text1(2) <> "" Then Text1(3) = Text1(2)
   End If
   
   TextInverse Text1(Index)
   CloseIme
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If Index = 4 Then
      If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
         KeyAscii = 0
         Beep
      End If
   End If
End Sub
'*************************************************
'  畫面輸入檢查
'
'*************************************************
Private Function FormCheck() As Boolean
   If Text1(0) & Text1(1) & Text1(2) & Text1(3) = "" Then
      If MsgBox("是否確定要列印所有員工資料!!", vbYesNo + vbDefaultButton2) = vbNo Then
         Text1(0).SetFocus
         Exit Function
      End If
   End If
   FormCheck = True
End Function

Private Sub FormReset()
   Dim oText As TextBox
   For Each oText In Text1
      oText.Text = Empty
   Next
End Sub

Private Sub PrintReport()
   Dim adoRst As ADODB.Recordset
   Dim stCon As String
   
   If Text1(0) <> "" Then
      stCon = stCon & " and st03>='" & Text1(0) & "'"
   End If
   
   If Text1(1) <> "" Then
      stCon = stCon & " and st03<='" & Text1(1) & "'"
   End If
   
   If Text1(2) <> "" Then
      stCon = stCon & " and sd01>='" & Text1(2) & "'"
   End If
   
   If Text1(3) <> "" Then
      stCon = stCon & " and sd01<='" & Text1(3) & "'"
   End If
   
   strExc(0) = "select x.*,st02,st03,a1.a0820 comp1,a2.a0820 comp2" & _
      " from salarydata x,staff s,acc080 a1,acc080 a2" & _
      " where st01(+)=sd01 and st04='1' and a1.a0801(+)=sd19 and a2.a0801(+)=sd28" & stCon & " order by st03,sd01"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If doPrint(adoRst) = True Then
         MsgBox "列印完畢 !"
      End If
   Else
      MsgBox "無待列印資料 !"
   End If
   Set adoRst = Nothing
End Sub

Private Function doPrint(ByRef p_Rst As ADODB.Recordset) As Boolean
   Dim strFontName As String, strFontSize As String
   
   strFontName = Printer.FontName
   strFontSize = Printer.FontSize
   
   GetPleft
On Error GoTo ErrHnd
   With p_Rst
      .MoveFirst
      iPage = 0
      PrintPageHeader
      PrintPageHeader1
      Do While Not .EOF
         strTemp(1) = "" & .Fields("st03")
         strTemp(2) = "" & .Fields("sd01")
         strTemp(3) = "" & .Fields("st02")
         strTemp(4) = "" & .Fields("comp1")
         strTemp(5) = Format(Val("" & .Fields("sd20")), "#,##0")
         strTemp(6) = Format(Val("" & .Fields("sd21")), "#,##0")
         strTemp(7) = Format(Val("" & .Fields("sd52")), "#,##0") 'Add By Sindy 2020/6/23 證照津貼
         strTemp(8) = Format(Val("" & .Fields("sd22")), "#,##0")
         strTemp(9) = Format(Val("" & .Fields("sd23")), "#,##0")
         strTemp(10) = Format(Val("" & .Fields("sd24")), "#,##0")
         strExc(1) = 0
         For intI = 5 To 9 '8
            strExc(1) = Val(strExc(1)) + Val(Format(strTemp(intI)))
         Next
         strTemp(11) = Format(strExc(1), "#,##0")
         PrintDetail strTemp
         
         If IsNull(.Fields("comp2")) = False Then
            strTemp(1) = ""
            strTemp(2) = ""
            strTemp(3) = ""
            strTemp(4) = "" & .Fields("comp2")
            strTemp(5) = Format(Val("" & .Fields("sd29")), "#,##0")
            strTemp(6) = Format(Val("" & .Fields("sd30")), "#,##0")
            strTemp(7) = Format(Val(""), "#,##0") 'Add By Sindy 2020/6/23 證照津貼
            strTemp(8) = Format(Val("" & .Fields("sd31")), "#,##0")
            strTemp(9) = Format(Val("" & .Fields("sd32")), "#,##0")
            strTemp(10) = Format(Val("" & .Fields("sd33")), "#,##0")
            strExc(1) = 0
            For intI = 5 To 9 '8
               strExc(1) = Val(strExc(1)) + Val(Format(strTemp(intI)))
            Next
            strTemp(11) = Format(strExc(1), "#,##0")
            PrintDetail strTemp
         Else
            iPrint = iPrint + m_iLineHeight
         End If
         
         .MoveNext
      Loop
      iPrint = iPrint + m_iLineHeight
      DrawLine
      Printer.EndDoc
   End With
   doPrint = True
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
      Printer.KillDoc
   End If
   
   Printer.FontName = strFontName
   Printer.FontSize = strFontSize
End Function

Private Sub GetPleft()
   Dim ii As Integer
   
   Printer.PaperSize = 9
   'Modify By Sindy 2020/6/23
   'Printer.Orientation = 1 '直印
   Printer.Orientation = 2 '橫印
   '2020/6/23 END
   Printer.FontSize = 11
   'm_iStartX = 100
   m_iStartX = 500
   m_iStartY = 500
   m_iPageHeight = Printer.ScaleHeight
   m_iLineHeight = 300
   m_iMargin = (Printer.Height - Printer.ScaleHeight) / 2
   m_iColGap = 400 '200 '欄位間隔
   
   ReDim PLeft(12) '11
   ReDim PColName(11) '10
   ReDim strTemp(11) '10
   
   ii = 1
   PColName(ii) = "部門"
   PLeft(ii) = m_iStartX
   
   ii = ii + 1
   PColName(ii) = "員工號"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(2, "　")) + m_iColGap
   
   ii = ii + 1
   PColName(ii) = "姓　　名"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(3, "　")) + m_iColGap
   
   ii = ii + 1
   PColName(ii) = "公司"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(4, "　")) + m_iColGap
   
   ii = ii + 1
   PColName(ii) = "基本薪資"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(2, "　")) + m_iColGap
   
   ii = ii + 1
   PColName(ii) = "職務津貼"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(5, "　")) + m_iColGap
   
   'Modify By Sindy 2020/6/23
   ii = ii + 1
   PColName(ii) = "證照津貼"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(5, "　")) + m_iColGap
   
   ii = ii + 1
   PColName(ii) = "技術津貼"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(5, "　")) + m_iColGap
   '2020/6/23 END
   
   ii = ii + 1
   PColName(ii) = "午餐津貼"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(5, "　")) + m_iColGap
   
   ii = ii + 1
   PColName(ii) = "差旅津貼"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(5, "　")) + m_iColGap
   
   ii = ii + 1
   PColName(ii) = "小　　計"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(5, "　")) + m_iColGap
   
   ii = ii + 1
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(5, "　")) + m_iColGap
End Sub

Private Sub PrintPageHeader()
   Dim strTmp As String
   
   strTmp = "員工薪資基本資料檢核表"
   
   iPrint = m_iStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = 20
   Printer.Font.Bold = True
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strTmp
   
   iPrint = iPrint + 600
   Printer.Font.Size = 11
   Printer.Font.Bold = False
      
   Printer.CurrentX = Printer.ScaleWidth - m_iMargin - 2500
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   iPrint = iPrint + m_iLineHeight
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   iPage = iPage + 1
   Printer.CurrentX = Printer.ScaleWidth - m_iMargin - 2500
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
   PrintNewLine
   
End Sub

Private Sub PrintPageHeader1()
   Printer.Font.Size = 11
   iPrint = iPrint + m_iLineHeight
   For intI = 1 To UBound(PColName)
      Printer.CurrentX = PLeft(intI)
      Printer.CurrentY = iPrint
      Printer.Print PColName(intI)
   Next
   
   iPrint = iPrint + m_iLineHeight
   DrawLine
End Sub

Private Sub PrintDetail(strData() As String)
    Dim iCol As Integer
    Printer.Font.Size = 11
    PrintNewLine
    For iCol = 1 To UBound(strData)
      If iCol > 4 Then
        'Modify By Sindy 2020/6/23 - 200
        Printer.CurrentX = PLeft(iCol + 1) - Printer.TextWidth(strData(iCol)) - m_iColGap - 200
        Printer.CurrentY = iPrint
        Printer.Print strData(iCol)
      Else
        Printer.CurrentX = PLeft(iCol)
        Printer.CurrentY = iPrint
        Printer.Print strData(iCol)
      End If
    Next
End Sub

Private Sub DrawLine()
   Printer.DrawWidth = 5
   Printer.Line (PLeft(1), iPrint)-(PLeft(UBound(PLeft)), iPrint)
   iPrint = iPrint - m_iLineHeight / 2
End Sub

Private Sub PrintNewLine(Optional ByVal p_iExtraLines As Integer = 2)
   iPrint = iPrint + m_iLineHeight
   If iPrint >= (m_iPageHeight - m_iMargin - p_iExtraLines * m_iLineHeight) Then
      DrawLine
      Printer.NewPage
      PrintPageHeader
      PrintPageHeader1
      iPrint = iPrint + m_iLineHeight
   End If
End Sub
