VERSION 5.00
Begin VB.Form frm170214 
   BorderStyle     =   1  '單線固定
   Caption         =   "互助會名單"
   ClientHeight    =   1500
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3684
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3684
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   990
      Style           =   2  '單純下拉式
      TabIndex        =   1
      Top             =   990
      Width           =   2460
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Left            =   1185
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "000"
      Top             =   660
      Width           =   435
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   2535
      TabIndex        =   3
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1455
      TabIndex        =   2
      Top             =   90
      Width           =   975
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
      Height          =   255
      Left            =   225
      TabIndex        =   5
      Top             =   1020
      Width           =   750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "互助會號："
      Height          =   180
      Left            =   225
      TabIndex        =   4
      Top             =   690
      Width           =   900
   End
End
Attribute VB_Name = "frm170214"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Create by Morgan 2009/1/20
Option Explicit

Dim co(1 To 7) As String
Dim PLeft() As Integer, PColName() As String, strTemp() As String
Dim iPrint As Integer, iPage As Integer, iFontSize As Integer
Dim m_iStartX As Integer, m_iStartY As Integer, m_iColGap As Integer
Dim m_iPageHeight As Long, m_iLineHeight As Long, m_iMargin As Long
Dim m_Tot(15, 2) As String, m_SubTot(15, 2) As String
Dim m_DefaultPrinter As String

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
   Case 0
      Screen.MousePointer = vbHourglass
      PrintReport
      Screen.MousePointer = vbDefault
   Case 1
      Unload Me
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, cmbPrinter, m_DefaultPrinter
   Text1 = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170214 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub PrintReport()
   Dim adoRst As ADODB.Recordset
   
On Error GoTo EXITSUB

   strExc(0) = "select * from Cooperation where co01='" & Text1 & "'"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With adoRst
      For intI = LBound(co) To UBound(co)
         co(intI) = "" & .Fields("co" & Format(intI, "00"))
      Next
      End With
   ElseIf intI = 0 Then
      MsgBox "查無該會號基本資料！"
      Text1.SetFocus
      TextInverse Text1
      GoTo EXITSUB
   End If
   
   strExc(0) = "select st02,x.* from CooperationMember x,staff where st01(+)=cm03 and cm01='" & Text1 & "'"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   
   If intI = 1 Then
      Me.Enabled = False
      If cmbPrinter <> Printer.DeviceName Then
         PUB_RestorePrinter cmbPrinter
      End If
      
      If doPrint(adoRst) = True Then
         MsgBox "列印完畢 !"
      End If
      
      If cmbPrinter.Tag <> cmbPrinter Then
         PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, 0, 0, Me.cmbPrinter.Text
      End If
      If Printer.DeviceName <> m_DefaultPrinter Then
         PUB_RestorePrinter m_DefaultPrinter
      End If
      Me.Enabled = True
   Else
      MsgBox "查無該會號名單資料！"
      Text1.SetFocus
      TextInverse Text1
   End If
   
EXITSUB:
   If Err.Number <> 0 Then MsgBox Err.Description
   Set adoRst = Nothing
End Sub

Private Function doPrint(ByRef p_Rst As ADODB.Recordset) As Boolean
   Dim ii As Integer
   
   GetPleft
On Error GoTo ErrHnd
   With p_Rst
      .MoveFirst
      iPage = 0
      PrintPageHeader
      PrintPageHeader1
      Do While Not .EOF
         strTemp(1) = "" & .Fields("cm02")
         strTemp(2) = "" & .Fields("cm03")
         strTemp(3) = "" & .Fields("st02")
         strTemp(4) = Format(TransDate("" & .Fields("cm04"), 1), "###/##/##")
         strTemp(5) = Format("" & .Fields("cm05"), "$#,###")
         strTemp(6) = Format("" & .Fields("cm06"), "$#,###")
         PrintDetail strTemp
         .MoveNext
      Loop
      iPrint = iPrint + m_iLineHeight
      DrawLine
      
      Printer.EndDoc
   End With
   doPrint = True
   Exit Function

ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub PrintPageHeader()
   Dim strTmp As String
   
   strTmp = "互助會名單"
   iPrint = m_iStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = 20
   Printer.Font.Bold = True
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strTmp
   iPrint = iPrint + 600
   
   Printer.Font.Size = iFontSize
   Printer.Font.Bold = False
   
   strExc(1) = "列印日期：" & Format(strSrvDate(2), "##/##/##")
   Printer.CurrentX = Printer.ScaleWidth - m_iMargin - Printer.TextWidth(strExc(1)) - 500
   Printer.CurrentY = iPrint
   Printer.Print strExc(1)
   
   iPrint = iPrint + m_iLineHeight
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   iPage = iPage + 1
   Printer.CurrentX = Printer.ScaleWidth - m_iMargin - Printer.TextWidth(strExc(1)) - 500
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
   
   iPrint = iPrint + 2 * m_iLineHeight
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print "一、會　　號：" & co(1)
   
   iPrint = iPrint + m_iLineHeight
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print "二、會　　首：台　一"
   
   iPrint = iPrint + m_iLineHeight
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print "三、金　　額：每會新台幣 " & Format(co(2), "$#,###") & " 元。"
   
   iPrint = iPrint + m_iLineHeight
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   'modify by sonia 2019/11/26 incCNV_CHINESE_MINKO改用incCNV_CHINESE_MINKO1
   Printer.Print "四、日　　期：自民國 " & TranslateKeyWord(incCNV_CHINESE_MINKO1, co(4), "") & " 起至民國 " & TranslateKeyWord(incCNV_CHINESE_MINKO1, co(5), "") & " 止 ( 最後一會 )。"
   
   iPrint = iPrint + m_iLineHeight
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print "五、底　　標： " & Format(co(7), "$#,###") & " 元 ( 外標 ) 標金一律至十位整數，無人投標時抽籤決定得標者。"
   
   iPrint = iPrint + m_iLineHeight
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print "六、標會日期：每月 " & co(6) & " 日中午 ( 遇假日順延一天 ) ，逾時以棄權論。"
   
   iPrint = iPrint + m_iLineHeight
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print "七、會員連會首共 " & co(3) & " 會"
   
   iPrint = iPrint + m_iLineHeight
End Sub

Private Sub PrintPageHeader1()
   Printer.Font.Size = iFontSize
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
    Printer.Font.Size = iFontSize
    PrintNewLine 3
    For iCol = 1 To UBound(strData)
      If iCol > 4 Then
        Printer.CurrentX = PLeft(iCol + 1) - Printer.TextWidth(strData(iCol)) - m_iColGap
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

Private Sub GetPleft()
   Dim ii As Integer
   Dim iOneCharWidth As Integer '一個字元寬(半個中文)
   Dim arrColWidth() As Integer
   
   iFontSize = 13
   Printer.PaperSize = 9 'A4
   Printer.Orientation = 1 '直印
   Printer.Font.Size = iFontSize
   m_iStartX = 500
   m_iStartY = 500
   m_iPageHeight = Printer.ScaleHeight
   m_iLineHeight = 350
   m_iMargin = (Printer.Height - Printer.ScaleHeight) / 2
   m_iColGap = 200 '欄位間隔
   
   ReDim PLeft(7)
   ReDim arrColWidth(6)
   ReDim PColName(6)
   ReDim strTemp(6)
   
   iOneCharWidth = Printer.TextWidth("全") / 2
   ii = 1
   PColName(ii) = "編號"
   arrColWidth(ii) = 4
   PLeft(ii) = m_iStartX + 1000
   
   ii = ii + 1
   PColName(ii) = "員工號"
   arrColWidth(ii) = 6
   PLeft(ii) = PLeft(ii - 1) + iOneCharWidth * arrColWidth(ii - 1) + m_iColGap
   
   ii = ii + 1
   PColName(ii) = "姓名"
   arrColWidth(ii) = 8
   PLeft(ii) = PLeft(ii - 1) + iOneCharWidth * arrColWidth(ii - 1) + m_iColGap
   
   ii = ii + 1
   PColName(ii) = "得標日期"
   arrColWidth(ii) = 9
   PLeft(ii) = PLeft(ii - 1) + iOneCharWidth * arrColWidth(ii - 1) + m_iColGap
   
   ii = ii + 1
   PColName(ii) = "得標金額"
   arrColWidth(ii) = 10
   PLeft(ii) = PLeft(ii - 1) + iOneCharWidth * arrColWidth(ii - 1) + m_iColGap
   
   ii = ii + 1
   PColName(ii) = "會款金額"
   arrColWidth(ii) = 10
   PLeft(ii) = PLeft(ii - 1) + iOneCharWidth * arrColWidth(ii - 1) + m_iColGap
   
   ii = ii + 1
   PLeft(ii) = PLeft(ii - 1) + iOneCharWidth * arrColWidth(ii - 1) + m_iColGap
End Sub

