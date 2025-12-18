VERSION 5.00
Begin VB.Form frm170233 
   BorderStyle     =   1  '單線固定
   Caption         =   "加班費明細表"
   ClientHeight    =   2724
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2724
   ScaleWidth      =   4860
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1935
      MaxLength       =   1
      TabIndex        =   0
      Top             =   750
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   2610
      MaxLength       =   1
      TabIndex        =   1
      Top             =   750
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   5
      Left            =   2925
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1410
      Width           =   780
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   3750
      TabIndex        =   7
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2670
      TabIndex        =   6
      Top             =   90
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   90
      TabIndex        =   9
      Top             =   1950
      Width           =   4665
      Begin VB.ComboBox cmbPrinter 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   8
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   10
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   4
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1425
      Width           =   780
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   1920
      MaxLength       =   5
      TabIndex        =   2
      Top             =   1080
      Width           =   780
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   2910
      MaxLength       =   5
      TabIndex        =   3
      Top             =   1080
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PS：不含超時加班費"
      Height          =   180
      Index           =   3
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "公  司  別："
      Height          =   180
      Index           =   1
      Left            =   945
      TabIndex        =   13
      Top             =   810
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2340
      X2              =   2655
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2655
      X2              =   2970
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   0
      Left            =   930
      TabIndex        =   12
      Top             =   1470
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "年　　月："
      Height          =   180
      Index           =   2
      Left            =   930
      TabIndex        =   11
      Top             =   1140
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2640
      X2              =   2955
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "frm170233"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2024/1/31 新部門已修改
'Created by Morgan 2012/6/19
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
         If TxtValidate = True Then
            Me.Enabled = False
            If cmbPrinter <> Printer.DeviceName Then
               PUB_RestorePrinter cmbPrinter
            End If
            PrintSheet
            '若印表機變動, 則更新列印設定
            If cmbPrinter.Tag <> cmbPrinter Then
                PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, 0, 0, Me.cmbPrinter.Text
            End If
            If Printer.DeviceName <> m_DefaultPrinter Then
               PUB_RestorePrinter m_DefaultPrinter
            End If
            Me.Enabled = True
         End If
         Screen.MousePointer = vbDefault
      Case 1
         Unload Me
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, cmbPrinter, m_DefaultPrinter
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170233 = Nothing
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   Dim oText As TextBox
   
   If txt1(2) = "" Then
      MsgBox "年月起不可空白 !"
      txt1(2).SetFocus
      Exit Function
   End If
   If txt1(3) = "" Then
      MsgBox "年月迄不可空白 !"
      txt1(3).SetFocus
      Exit Function
   End If
   
   TxtValidate = True
End Function

Private Sub PrintSheet()
   Dim stCon As String
   Dim stLstComp As String, stLstDept As String, stLstName As String
   Dim lngSubTot As Long, lngTotal As Long, iLstPage As Integer
   Dim lngHrSubTot As Double, lngHrTotal As Double 'Added by Morgan 2017/1/9
   
   stCon = ""
   '公司別
   If txt1(0) <> "" Then
      stCon = stCon & " and sm37>='" & txt1(0) & "'"
   End If
   If txt1(1) <> "" Then
      stCon = stCon & " and sm37<='" & txt1(1) & "'"
   End If
   '年月
   If txt1(2) <> "" Then
      stCon = stCon & " and sm02>=" & (Val(txt1(2)) + 191100)
   End If
   If txt1(3) <> "" Then
      stCon = stCon & " and sm02<=" & (Val(txt1(3)) + 191100)
   End If
   '員工編號
   If txt1(4) <> "" Then
      stCon = stCon & " and st01>='" & txt1(4) & "'"
   End If
   If txt1(5) <> "" Then
      stCon = stCon & " and st01<='" & txt1(5) & "'"
   End If
   
   'modify by sonia 2016/4/28 辜說不含超時加班費,故改sm12為sm12-nvl(sm28,0)
   'Modified by Morgan 2024/1/31 +新部門ACC090NEW
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   strExc(0) = "select sm37,decode(sign(sm02-" & Left(新部門啟用日, 6) & "),-1,a0902,a0922) a0902,st02,sm02,sm12-nvl(sm28,0) sm12,sm11" & _
      " From SalaryMonth, staff, acc090, acc090NEW" & _
      " Where st01(+)=substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4) and sm12>0" & _
      " and a0901(+)=sm03 and a0921(+)=sm03" & stCon & _
      " order by 1,2,3,4"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      GetPleft
      
      With RsTemp
      
      iPage = 0
      PrintPageHeader
      PrintPageHeader1
               
      Do While Not .EOF
         If stLstComp <> .Fields("sm37") Then
            If stLstComp <> "" Then
               PrintSubTotal stLstComp, lngSubTot, lngHrSubTot
               lngTotal = lngTotal + lngSubTot
               lngHrTotal = lngHrTotal + lngHrSubTot
               lngSubTot = 0
               lngHrSubTot = 0
               
               Printer.NewPage
               PrintPageHeader
               PrintPageHeader1
            End If
         End If
         
         iLstPage = iPage
         PrintNewLine
         
         If iLstPage = iPage And stLstComp = .Fields("sm37") Then
            strTemp(1) = ""
         Else
            strTemp(1) = .Fields("sm37")
            stLstComp = .Fields("sm37")
         End If
         
         If iLstPage = iPage And stLstDept = .Fields("a0902") Then
            strTemp(2) = ""
         Else
            strTemp(2) = "" & .Fields("a0902")
            stLstDept = "" & .Fields("a0902")
         End If
         
         If iLstPage = iPage And stLstName = .Fields("st02") Then
            strTemp(3) = ""
         Else
            strTemp(3) = .Fields("st02")
            stLstName = .Fields("st02")
         End If
         
         strTemp(4) = .Fields("sm02") - 191100
         strTemp(5) = Format(.Fields("sm12"), "#,###")
         strTemp(6) = Format(.Fields("sm11"), "#.0")
         PrintDetail strTemp
         
         lngSubTot = lngSubTot + .Fields("sm12")
         lngHrSubTot = lngHrSubTot + .Fields("sm11")
         .MoveNext
      Loop
      PrintSubTotal stLstComp, lngSubTot, lngHrSubTot
      lngTotal = lngTotal + lngSubTot
      lngHrTotal = lngHrTotal + lngHrSubTot
      PrintTotal lngTotal, lngHrTotal
      Printer.EndDoc
      End With
   End If
End Sub

Private Sub PrintSubTotal(pComp As String, pSubTot As Long, pHrSubTot As Double)
   Dim strTmp As String
   
   PrintNewLine
   DrawLine
   Printer.Font.Size = 11
   PrintNewLine
   Printer.CurrentX = PLeft(4, 1)
   Printer.CurrentY = iPrint
   Printer.Print pComp & " 公司小計："
   
   strTmp = Format(pSubTot, "#,###")
   Printer.CurrentX = PLeft(6, 1) - Printer.TextWidth(strTmp) - m_iColGap
   Printer.CurrentY = iPrint
   Printer.Print strTmp
   
   strTmp = Format(pHrSubTot, "#.0")
   Printer.CurrentX = PLeft(7, 1) - Printer.TextWidth(strTmp) - m_iColGap
   Printer.CurrentY = iPrint
   Printer.Print strTmp
End Sub

Private Sub PrintTotal(pTotal As Long, pHrTotal As Double)
   Dim strTmp As String
   
   PrintNewLine
   DrawLine
   Printer.Font.Size = 11
   PrintNewLine
   Printer.CurrentX = PLeft(4, 1)
   Printer.CurrentY = iPrint
   Printer.Print "總計："
   
   strTmp = Format(pTotal, "#,###")
   Printer.CurrentX = PLeft(6, 1) - Printer.TextWidth(strTmp) - m_iColGap
   Printer.CurrentY = iPrint
   Printer.Print strTmp
   
   strTmp = Format(pHrTotal, "#.0")
   Printer.CurrentX = PLeft(7, 1) - Printer.TextWidth(strTmp) - m_iColGap
   Printer.CurrentY = iPrint
   Printer.Print strTmp
   
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
         If Index = 0 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         Else
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Cancel = True
            End If
         End If
         
      Case 2, 3
         If txt1(Index) <> "" Then
            If ChkDate(txt1(Index) & "01") = False Then
               Cancel = True
            End If
         End If
         If Index = 2 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         Else
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Cancel = True
            End If
         End If
   End Select
End Sub

Private Function CutWords(pData As String, pColWith As Integer) As String
   Dim iLen As Integer, stNew As String, stNew1 As String
   iLen = Len(pData)
   For intI = 1 To iLen
      stNew1 = stNew & Mid(pData, intI, 1)
      If Printer.TextWidth(stNew1) > pColWith Then
         Exit For
      Else
         stNew = stNew1
      End If
   Next
   CutWords = stNew
End Function

Private Sub GetPleft()
   Dim ii As Integer
   
   Printer.PaperSize = 9
   Printer.Orientation = 1
   Printer.FontSize = 12
   m_iStartX = 300
   m_iStartY = 500
   m_iPageHeight = Printer.ScaleHeight
   m_iLineHeight = 300
   m_iMargin = (Printer.Height - Printer.ScaleHeight) / 2
   m_iColGap = 200 '欄位間隔
   
   
   'Modified by Morgan 2017/1/7 +加班時數
   ReDim PLeft(7, 2)
   ReDim PColName(6)
   ReDim strTemp(6)
   
   ii = 1
   PColName(ii) = "公司別"
   PLeft(ii, 1) = m_iStartX
   PLeft(ii, 2) = 3
   
   ii = ii + 1
   PColName(ii) = "部門"
   PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
   'Modified by Morgan 2024/1/31
   'PLeft(ii, 2) = 5
   PLeft(ii, 2) = 7
   
   ii = ii + 1
   PColName(ii) = "員工姓名"
   PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
   PLeft(ii, 2) = 4
   
   ii = ii + 1
   PColName(ii) = "年月"
   PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
   PLeft(ii, 2) = 4
   
   ii = ii + 1
   PColName(ii) = "加班費"
   PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
   PLeft(ii, 2) = 5
   
   ii = ii + 1
   PColName(ii) = "加班時數"
   PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
   PLeft(ii, 2) = 4
   
   ii = ii + 1
   PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
   
End Sub

Private Sub PrintPageHeader()
   Dim strTmp As String
   
   strTmp = "台一關係企業 加班費明細表"
   
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
   
   strTmp = "公  司  別：" & txt1(0) & " - " & txt1(1)
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(String(10, "　"))) / 2
   Printer.CurrentY = iPrint
   Printer.Print strTmp
   
   iPrint = iPrint + m_iLineHeight
   strTmp = "年　　月：" & txt1(2) & " - " & txt1(3)
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(String(10, "　"))) / 2
   Printer.CurrentY = iPrint
   Printer.Print strTmp
   
   iPrint = iPrint + m_iLineHeight
   strTmp = "員工編號：" & txt1(4) & " - " & txt1(5)
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(String(10, "　"))) / 2
   Printer.CurrentY = iPrint
   Printer.Print strTmp
   
   
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
   iPrint = iPrint + m_iLineHeight
   Printer.Font.Size = 11
   Printer.FontBold = True
   For intI = 1 To UBound(PColName)
      Printer.CurrentX = PLeft(intI, 1)
      Printer.CurrentY = iPrint
      Printer.Print PColName(intI)
   Next
   Printer.FontBold = False
   iPrint = iPrint + m_iLineHeight
   DrawLine
End Sub

Private Sub PrintDetail(strData() As String)
    Dim iCol As Integer
    Printer.Font.Size = 11
    For iCol = 1 To UBound(strData)
      If iCol = 5 Or iCol = 6 Then
        Printer.CurrentX = PLeft(iCol + 1, 1) - Printer.TextWidth(strData(iCol)) - m_iColGap
        Printer.CurrentY = iPrint
        Printer.Print strData(iCol)
      Else
        Printer.CurrentX = PLeft(iCol, 1)
        Printer.CurrentY = iPrint
        'Modified by Morgan 2024/1/31
        'Printer.Print strData(iCol)
        PUB_PrintUnicodeText strData(iCol), Printer.CurrentX, Printer.CurrentY, 0
        'end 2024/1/31
      End If
    Next
End Sub

Private Sub DrawLine()
   Printer.DrawWidth = 5
   Printer.Line (PLeft(1, 1), iPrint)-(PLeft(UBound(PLeft, 1), 1), iPrint)
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
