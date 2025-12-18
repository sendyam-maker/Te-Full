VERSION 5.00
Begin VB.Form frm170231 
   BorderStyle     =   1  '單線固定
   Caption         =   "人事薪資異動檢核表"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   3945
   Begin VB.TextBox Text1 
      Height          =   255
      Index           =   5
      Left            =   2250
      MaxLength       =   2
      TabIndex        =   5
      Text            =   "01"
      Top             =   1260
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Index           =   4
      Left            =   1350
      MaxLength       =   2
      TabIndex        =   4
      Text            =   "01"
      Top             =   1260
      Width           =   765
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Index           =   0
      Left            =   1365
      MaxLength       =   7
      TabIndex        =   0
      Text            =   "960101"
      Top             =   600
      Width           =   765
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Index           =   1
      Left            =   2265
      MaxLength       =   7
      TabIndex        =   1
      Text            =   "960131"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Index           =   2
      Left            =   1365
      MaxLength       =   6
      TabIndex        =   2
      Top             =   930
      Width           =   765
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Index           =   3
      Left            =   2265
      MaxLength       =   6
      TabIndex        =   3
      Top             =   930
      Width           =   735
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   1155
      Style           =   2  '單純下拉式
      TabIndex        =   6
      Top             =   1590
      Width           =   2460
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1260
      TabIndex        =   8
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   2340
      TabIndex        =   7
      Top             =   90
      Width           =   975
   End
   Begin VB.Line Line3 
      X1              =   1860
      X2              =   2520
      Y1              =   1350
      Y2              =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "異動原因："
      Height          =   180
      Left            =   405
      TabIndex        =   12
      Top             =   1260
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "異動日期："
      Height          =   180
      Left            =   405
      TabIndex        =   11
      Top             =   630
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   1875
      X2              =   2535
      Y1              =   690
      Y2              =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Left            =   405
      TabIndex        =   10
      Top             =   960
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   1875
      X2              =   2535
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Label Label15 
      Caption         =   "印表機："
      Height          =   180
      Left            =   405
      TabIndex        =   9
      Top             =   1620
      Width           =   720
   End
End
Attribute VB_Name = "frm170231"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Create by Morgan 2009/2/9
Option Explicit

Dim PLeft() As Integer, PColName() As String, strTemp() As String
Dim iPrint As Integer, iPage As Integer
Dim m_iStartX As Integer, m_iStartY As Integer, m_iColGap As Integer, m_iCols As Single
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
   Set frm170231 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   If Index = 1 Then
      If Text1(0) <> "" And Text1(1) = "" Then Text1(1) = Text1(0)
   ElseIf Index = 3 Then
      If Text1(2) <> "" Then Text1(3) = Text1(2)
   End If
   
   TextInverse Text1(Index)
   CloseIme
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'*************************************************
'  畫面輸入檢查
'
'*************************************************
Private Function FormCheck() As Boolean
   If Text1(0) = "" Then
      MsgBox "異動日期起日不可空白!!"
      Text1(0).SetFocus
      Exit Function
   ElseIf ChkDate(Text1(0)) = False Then
      Text1(0).SetFocus
      Exit Function
   End If
   If Text1(1) = "" Then
      MsgBox "異動日期迄日不可空白!!"
      Text1(1).SetFocus
      Exit Function
   ElseIf ChkDate(Text1(1)) = False Then
      Text1(1).SetFocus
      Exit Function
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
   Dim stCon As String, stCon1 As String
   
   stCon = "": stCon1 = ""
   If Text1(0) <> "" Then
      stCon = stCon & " and sl02>=" & DBDATE(Text1(0))
      stCon1 = stCon1 & " and sc02>=" & DBDATE(Text1(0))
   End If
   
   If Text1(1) <> "" Then
      stCon = stCon & " and sl02<=" & DBDATE(Text1(1))
      stCon1 = stCon1 & " and sc02<=" & DBDATE(Text1(1))
   End If
   
   If Text1(2) <> "" Then
      stCon = stCon & " and sl01>='" & Text1(2) & "'"
      stCon1 = stCon1 & " and sc01>='" & Text1(2) & "'"
   End If
   
   If Text1(3) <> "" Then
      stCon = stCon & " and sl01<='" & Text1(3) & "'"
      stCon1 = stCon1 & " and sc01<='" & Text1(3) & "'"
   End If
   
   If Text1(4) <> "" Then
      stCon1 = stCon1 & " and sc03>='" & Text1(4) & "'"
   End If
   
   If Text1(5) <> "" Then
      stCon1 = stCon1 & " and sc03<='" & Text1(5) & "'"
   End If
   
   'Modify By Sindy 2020/6/23 + ,sl39
   strExc(0) = "select sl02,sl01,st02,a0902 dept,'' reason,'' tit,'' pos" & _
      ",a1.a0820 comp1,sl04,sl05,sl06,sl09,sl10,sl11,sl12,sl13,sl14,sl15,sl16,sl17" & _
      ",a2.a0820 comp2,sl19,sl20,sl21,sl22,sl23,sl24,sl25,sl39" & _
      " from salarylog,staff,acc090,acc080 a1,acc080 a2" & _
      " where st01(+)=sl01 and a0901(+)=st03" & stCon & _
      " and a1.a0801(+)=sl33 and a2.a0801(+)=sl34"
      
   'Modify By Sindy 2020/6/23 + ,0 sl39
   strExc(0) = strExc(0) & " union select sc02,sc01,st02,a0902 dept,c3.ac03 reason,c1.ac03 tit,c2.ac03 pos" & _
      ",null comp1,0 sl04,0 sl05,0 sl06,0 sl09,0 sl10,0 sl11,0 sl12,0 sl13,0 sl14,0 sl15,0 sl16,0 sl17" & _
      ",null comp2,0 sl19,0 sl20,0 sl21,0 sl22,0 sl23,0 sl24,0 sl25,0 sl39" & _
      " from staff_change,staff,acc090,allcode c1,allcode c2,allcode c3" & _
      " where st01(+)=sc01 and a0901(+)=sc04" & stCon1 & _
      " and c1.ac02(+)=sc05 and c1.ac01(+)='01'" & _
      " and c2.ac02(+)=sc06 and c2.ac01(+)='02'" & _
      " and c3.ac02(+)=sc03 and c3.ac01(+)='05'"
      
   strExc(0) = strExc(0) & " order by 4,2,1,5"

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
   Dim strKey As String
   
   strFontName = Printer.FontName
   strFontSize = Printer.FontSize
   
   GetPleft
On Error GoTo ErrHnd
   With p_Rst
      .MoveFirst
      iPage = 0
      PrintPageHeader
      PrintPageHeader1
      strKey = ""
      Do While Not .EOF
         Erase strTemp
         ReDim strTemp(m_iCols)
         strTemp(1) = ChangeWStringToTDateString("" & .Fields("sl02"))
         'strTemp(2) = "" & .Fields("sl01")
         strTemp(2) = "" & .Fields("st02")
         'strTemp(4) = "" & .Fields("dept")
         
         'If strKey = strTemp(1) & strTemp(2) & strTemp(3) & strTemp(4) Then
         If strKey = strTemp(1) & strTemp(2) Then
            Erase strTemp
            ReDim strTemp(m_iCols)
         Else
            'strKey = strTemp(1) & strTemp(2) & strTemp(3) & strTemp(4)
            strKey = strTemp(1) & strTemp(2)
         End If
         strTemp(3) = "" & .Fields("reason")
         strTemp(4) = "" & .Fields("Tit")
         strTemp(5) = "" & .Fields("Pos")
         If IsNull(.Fields("Comp1")) = False Then
            strTemp(6) = "" & .Fields("Comp1") & "所得"
            strTemp(7) = Format(Val("" & .Fields("sl11")), "#,###")
            strTemp(8) = Format(Val("" & .Fields("sl12")), "#,###")
            strTemp(9) = Format(Val("" & .Fields("sl39")), "#,###") 'Add By Sindy 2020/6/23 證照津貼
            strTemp(10) = Format(Val("" & .Fields("sl13")), "#,###")
            strTemp(11) = Format(Val("" & .Fields("sl14")), "#,###")
            strTemp(12) = Format(Val("" & .Fields("sl15")), "#,###")
            strTemp(13) = Format(Val("" & .Fields("sl16")), "#,###")
            strTemp(14) = Format(Val("" & .Fields("sl17")), "#,###")
         End If
         PrintDetail strTemp
         
         Erase strTemp
         ReDim strTemp(m_iCols)
         
         If IsNull(.Fields("Comp1")) = False Then
            strTemp(7) = Format(Val("" & .Fields("sl09")), "#,###") '勞保費
            strTemp(8) = Format(Val("" & .Fields("sl10")), "#,###") '健保費
            strTemp(9) = Format(Val("" & .Fields("sl05")) + Val("" & .Fields("sl06")), "#,###") '婚喪戶助
            If Val("" & .Fields("sl04")) > 0 Then
               strTemp(10) = .Fields("sl04") & " %" '12
            End If
            PrintDetail strTemp
         End If
         
         If IsNull(.Fields("Comp2")) = False Then
            Erase strTemp
            ReDim strTemp(m_iCols)
            strTemp(6) = .Fields("Comp2") & "所得"
            strTemp(7) = Format(Val("" & .Fields("sl19")), "#,###")
            strTemp(8) = Format(Val("" & .Fields("sl20")), "#,###")
            strTemp(9) = Format(Val(""), "#,###") 'Add By Sindy 2020/6/23 證照津貼
            strTemp(10) = Format(Val("" & .Fields("sl21")), "#,###")
            strTemp(11) = Format(Val("" & .Fields("sl22")), "#,###")
            strTemp(12) = Format(Val("" & .Fields("sl23")), "#,###")
            strTemp(13) = Format(Val("" & .Fields("sl24")), "#,###")
            strTemp(14) = Format(Val("" & .Fields("sl25")), "#,###")
            PrintDetail strTemp
            
            Erase strTemp
            ReDim strTemp(m_iCols)
            strTemp(6) = "合計所得"
            strTemp(7) = Format(Val("" & .Fields("sl11")) + Val("" & .Fields("sl19")), "#,###")
            strTemp(8) = Format(Val("" & .Fields("sl12")) + Val("" & .Fields("sl20")), "#,###")
            strTemp(9) = Format(Val("" & .Fields("sl39")), "#,###") 'Add By Sindy 2020/6/23 證照津貼
            strTemp(10) = Format(Val("" & .Fields("sl13")) + Val("" & .Fields("sl21")), "#,###")
            strTemp(11) = Format(Val("" & .Fields("sl14")) + Val("" & .Fields("sl22")), "#,###")
            strTemp(12) = Format(Val("" & .Fields("sl15")) + Val("" & .Fields("sl23")), "#,###")
            strTemp(13) = Format(Val("" & .Fields("sl16")) + Val("" & .Fields("sl24")), "#,###")
            strTemp(14) = Format(Val("" & .Fields("sl17")) + Val("" & .Fields("sl25")), "#,###")
            PrintDetail strTemp
            
            Erase strTemp
            ReDim strTemp(m_iCols)
            strTemp(7) = Format(Val("" & .Fields("sl09")), "#,###")
            strTemp(8) = Format(Val("" & .Fields("sl10")), "#,###")
            strTemp(9) = Format(Val("" & .Fields("sl05")) + Val("" & .Fields("sl06")), "#,###")
            If Val("" & .Fields("sl04")) > 0 Then
               strTemp(10) = .Fields("sl04") & " %"
            End If
            PrintDetail strTemp
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
   Printer.Orientation = 2
   Printer.FontSize = 12
   m_iStartX = 400 '100
   m_iStartY = 500
   m_iPageHeight = Printer.ScaleHeight
   m_iLineHeight = 300
   m_iMargin = (Printer.Height - Printer.ScaleHeight) / 2
   m_iColGap = 120 '欄位間隔
   
   m_iCols = 14 '15
   ReDim PLeft(m_iCols + 1)
   ReDim PColName(m_iCols, 2)
   
   ii = 1
   PColName(ii, 2) = "異動日期"
   PLeft(ii) = m_iStartX
   
   'Modify By Sindy 2020/6/23 Mark
'   ii = ii + 1
'   PColName(ii, 2) = "員工號"
'   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(9, "9")) + m_iColGap
   
   ii = ii + 1
   PColName(ii, 2) = "姓名"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(9, "9")) + m_iColGap
   
   'Modify By Sindy 2020/6/23 Mark
'   ii = ii + 1
'   PColName(ii, 2) = "部門"
'   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(4, "字")) + m_iColGap
   
   ii = ii + 1
   PColName(ii, 2) = "異動原因"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(5, "字")) + m_iColGap
   
   ii = ii + 1
   PColName(ii, 2) = "異動職稱"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(4, "字")) + m_iColGap
   
   ii = ii + 1
   PColName(ii, 2) = "異動職位"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(5, "字")) + m_iColGap
   
   ii = ii + 1
   PColName(ii, 1) = "所得項目"
   PColName(ii, 2) = "扣款項目"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(5, "字")) + m_iColGap
   
   ii = ii + 1
   PColName(ii, 1) = "基本薪資"
   PColName(ii, 2) = "勞保費"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(4, "字")) + m_iColGap
   
   ii = ii + 1
   PColName(ii, 1) = "職務津貼"
   PColName(ii, 2) = "健保費"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(9, "9")) + m_iColGap
   
   'Modify By Sindy 2020/6/23
   ii = ii + 1
   PColName(ii, 1) = "證照津貼"
   PColName(ii, 2) = "婚喪戶助"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(9, "9")) + m_iColGap
   '2020/6/23 END
   
   ii = ii + 1
   PColName(ii, 1) = "技術津貼"
   PColName(ii, 2) = "所得稅率"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(9, "9")) + m_iColGap
   
   ii = ii + 1
   PColName(ii, 1) = "午餐津貼"
   'PColName(ii, 2) = "所得稅率"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(8, "9")) + m_iColGap

   ii = ii + 1
   PColName(ii, 1) = "差旅津貼"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(8, "9")) + m_iColGap
   
   ii = ii + 1
   PColName(ii, 1) = "房租津貼"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(8, "9")) + m_iColGap

   ii = ii + 1
   PColName(ii, 1) = "特支費"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(8, "9")) + m_iColGap
   
   ii = ii + 1
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(8, "9")) + m_iColGap
End Sub

Private Sub PrintPageHeader()
   Dim strTmp As String
   
   strTmp = "人事薪資異動檢核表"
   
   iPrint = m_iStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = 20
   Printer.Font.Bold = True
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strTmp
   
   iPrint = iPrint + 600
   Printer.Font.Size = 12
   Printer.Font.Bold = False
      
   strExc(1) = "異動日期："
   Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(strExc(1))
   Printer.CurrentY = iPrint
   Printer.Print strExc(1) & Text1(0) & " - " & Text1(1)
   
   iPrint = iPrint + m_iLineHeight
   
   strExc(1) = "員工編號："
   Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(strExc(1))
   Printer.CurrentY = iPrint
   Printer.Print strExc(1) & Text1(2) & " - " & Text1(3)
   
   iPrint = iPrint + m_iLineHeight
   
   strExc(1) = "異動原因："
   Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(strExc(1))
   Printer.CurrentY = iPrint
   Printer.Print strExc(1) & Text1(4) & " - " & Text1(5)
   
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
   Printer.Font.Size = 12
   iPrint = iPrint + m_iLineHeight
   For intI = 1 To UBound(PColName)
      Printer.CurrentX = PLeft(intI)
      Printer.CurrentY = iPrint
      Printer.Print PColName(intI, 1)
   Next
   
   iPrint = iPrint + m_iLineHeight
   For intI = 1 To UBound(PColName)
      Printer.CurrentX = PLeft(intI)
      Printer.CurrentY = iPrint
      Printer.Print PColName(intI, 2)
   Next
   
   iPrint = iPrint + m_iLineHeight
   DrawLine
End Sub

Private Sub PrintDetail(strData() As String)
    Dim iCol As Integer
    Printer.Font.Size = 12
    PrintNewLine
    For iCol = 1 To UBound(strData)
      If strData(iCol) <> "" Then
         Select Case iCol
         Case 1 To 6 '8
            Printer.CurrentX = PLeft(iCol)
            Printer.CurrentY = iPrint
            
         Case 7 To m_iCols '9
            Printer.CurrentX = PLeft(iCol + 1) - Printer.TextWidth(strData(iCol)) - m_iColGap
            Printer.CurrentY = iPrint
         End Select
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
