VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm083016 
   BorderStyle     =   1  '單線固定
   Caption         =   "請款點數明細表"
   ClientHeight    =   2985
   ClientLeft      =   4080
   ClientTop       =   3900
   ClientWidth     =   4125
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4125
   Begin VB.OptionButton opt1 
      Caption         =   "承辦人"
      Height          =   225
      Index           =   1
      Left            =   2310
      TabIndex        =   9
      Top             =   2520
      Width           =   1185
   End
   Begin VB.OptionButton opt1 
      Caption         =   "日期"
      Height          =   225
      Index           =   0
      Left            =   1320
      TabIndex        =   8
      Top             =   2520
      Value           =   -1  'True
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   2175
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1740
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   1290
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1740
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1290
      MaxLength       =   6
      TabIndex        =   7
      Top             =   2085
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1290
      TabIndex        =   0
      Top             =   690
      Width           =   1740
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1290
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1050
      Width           =   960
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2400
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1050
      Width           =   960
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1290
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1395
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2175
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1395
      Width           =   705
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2430
      TabIndex        =   10
      Top             =   165
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3225
      TabIndex        =   11
      Top             =   165
      Width           =   756
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   2505
      Left            =   150
      TabIndex        =   18
      Top             =   4170
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4419
      _Version        =   393216
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Caption         =   "列  印  別："
      Height          =   180
      Index           =   1
      Left            =   270
      TabIndex        =   19
      Top             =   2520
      Width           =   915
   End
   Begin VB.Line Line3 
      X1              =   1800
      X2              =   2550
      Y1              =   1830
      Y2              =   1830
   End
   Begin VB.Label Label1 
      Caption         =   "國　　家："
      Height          =   180
      Index           =   10
      Left            =   270
      TabIndex        =   17
      Top             =   1770
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "承  辦  人："
      Height          =   180
      Index           =   6
      Left            =   270
      TabIndex        =   16
      Top             =   2145
      Width           =   990
   End
   Begin VB.Label lbl1 
      BackColor       =   &H80000004&
      Height          =   180
      Index           =   1
      Left            =   2100
      TabIndex        =   15
      Top             =   2130
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   270
      TabIndex        =   14
      Top             =   720
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "日　　期："
      Height          =   180
      Index           =   2
      Left            =   270
      TabIndex        =   13
      Top             =   1080
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "業  務  區："
      Height          =   180
      Index           =   3
      Left            =   270
      TabIndex        =   12
      Top             =   1440
      Width           =   915
   End
   Begin VB.Line Line1 
      X1              =   1680
      X2              =   2940
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line2 
      X1              =   1815
      X2              =   2565
      Y1              =   1560
      Y2              =   1560
   End
End
Attribute VB_Name = "frm083016"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2016/12/12 請款點數明細表
Option Explicit
Dim m_SubFclLin As Double, m_SubCfl As Double '承辦人小計
Dim m_TotFclLin As Double, m_TotCfl As Double '合計
Dim mPrtOrt As Integer  '原本預設印表機的列印方向
Dim mPageSize As Integer '原本預設印表機的紙張
Private Const ciTitleFontSize = 16, ciFontSize = 11
Private Const cInX As Integer = 10
Dim strTemp(0 To cInX + 1) As String
Dim PTitle(0 To cInX + 1) As String
Dim PLeft(0 To cInX) As Integer
Private Const ciStartX = 400, ciStartY = 500, ciColGap = 150
Dim iPrint As Integer, iPage As Integer
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Private Const FatDot As String = "0.0" '統一小數點(原本0.0)
Dim mRptStr1 As String '列印條件
Dim s As Integer

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     If Len(txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         Exit Sub
     Else
         If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
            Me.txt1(1).SetFocus
            txt1_GotFocus 1
            Exit Sub
         End If
         If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
            Me.txt1(2).SetFocus
            txt1_GotFocus 2
            Exit Sub
         End If
         
         If Len(txt1(1)) = 0 Or Len(txt1(2)) = 0 Then
             s = MsgBox("日期區間不可空白!!", , "USER 輸入錯誤")
             txt1(1).SetFocus
             txt1_GotFocus (1)
             Exit Sub
         End If
            
            txt1(0) = Trim(txt1(0))
            mPageSize = Printer.PaperSize
            mPrtOrt = Printer.Orientation
            Screen.MousePointer = vbHourglass
            Me.Enabled = False
            PrintCase
            Me.Enabled = True
            Screen.MousePointer = vbDefault
            Printer.PaperSize = mPageSize
            Printer.Orientation = mPrtOrt
     End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Private Sub PrintCase()
Dim stStr As String, stStr1 As String
Dim inJ As Integer
Dim RsQ As New ADODB.Recordset
Dim strGrp As String
Dim strDiv As String
Dim tmpArr As Variant 'Added by Lydia 2017/01/03

    'Modified by Lydia 2017/01/03 判斷系統類別
    'stStr = " AND CP01 IN (" & GetAddStr(txt1(0).Text) & ") "
    tmpArr = Split(UCase(txt1(0).Text), ",")
    For intI = 0 To UBound(tmpArr)
        If tmpArr(intI) = "CFL" Or tmpArr(intI) = "FCL" Or tmpArr(intI) = "LIN" Then
           stStr = stStr & IIf(stStr <> "", ",", "") & tmpArr(intI)
        End If
    Next intI
    If Len(stStr) = 0 Then
       MsgBox "系統類別不屬於外法!"
       Exit Sub
    End If
    
    'Added by Lydia 2023/04/20
    ClearQueryLog (Me.Name) '清除查詢印表記錄檔欄位
    pub_QL05 = pub_QL05 & ";" & Me.Caption
    If Trim(txt1(0)) <> "" Then pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0)
    If Trim(txt1(1) & txt1(2)) <> "" Then pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(1) & "-" & txt1(2)
    If Trim(txt1(3) & txt1(4)) <> "" Then pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(3) & "-" & txt1(4)
    If Trim(txt1(9) & txt1(10)) <> "" Then pub_QL05 = pub_QL05 & ";" & Label1(10) & txt1(9) & "-" & txt1(10)
    If Trim(txt1(6)) <> "" Then pub_QL05 = pub_QL05 & ";" & Label1(6) & txt1(6)
    If opt1(0).Value = True Then pub_QL05 = pub_QL05 & ";" & Label1(1) & "日期"
    If opt1(1).Value = True Then pub_QL05 = pub_QL05 & ";" & Label1(1) & "承辦人"
    'end 2023/04/20
    
    stStr = " AND CP01 IN (" & GetAddStr(stStr) & ") "
    
    mRptStr1 = "日期：" & ChangeTStringToTDateString(txt1(1)) & " - " & ChangeTStringToTDateString(txt1(2))
    If txt1(3) <> "" Then
       stStr = stStr & " AND CP12>=" & CNULL(txt1(3))
       mRptStr1 = mRptStr1 & "　業務區：" & txt1(3) & " - "
    End If
    If txt1(4) <> "" Then
       stStr = stStr & " AND CP12<=" & CNULL(txt1(4))
       mRptStr1 = mRptStr1 & IIf(InStr(mRptStr1, "業務區") = 0, "　業務區：" & " - " & txt1(4), txt1(4))
    End If
    If txt1(9) <> "" Then
       stStr1 = stStr1 & " AND NA01>=" & CNULL(txt1(9))
       mRptStr1 = mRptStr1 & "　國籍：" & txt1(9) & " - "
    End If
    If txt1(10) <> "" Then
       stStr1 = stStr1 & " AND NA01<=" & CNULL(txt1(10))
       mRptStr1 = mRptStr1 & IIf(InStr(mRptStr1, "國籍") = 0, "　國籍：" & " - " & txt1(10), txt1(10))
    End If
    If txt1(6) <> "" Then
       stStr = stStr & " AND CP14 =" & CNULL(txt1(6))
       mRptStr1 = mRptStr1 & "　承辦人：" & txt1(6) & " " & Lbl1(1)
    End If
    
    '抓收據
    strSql = "SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ CP05-19110000 A00,CP09 A01,CP01||'-'||CP02||DECODE(CP03,'0','','-'||CP03)||DECODE(CP04,'00','','-'||CP04) A02" & _
             ",CPM03 A03,LC13 A04,NVL(CU04,NVL(CU05,CU06)) A05,CP40||' '||CP41||' '||CP42 A05_1,CP18 A06,S1.ST02 A07,S2.ST02 A08,DECODE(CP01,'CFL',A0902,NVL(N1.NA03,N2.NA03)) A09 " & _
             ",NVL(NVL(FA05,FA04),FA06) A10,CP60,CP01,CP02,CP03,CP04,CP12,CP14,NVL(N1.NA01,N2.NA01) NA01 " & _
             "FROM CASEPROGRESS,CASEPROPERTYMAP,LAWCASE,FAGENT,CUSTOMER,ACC090,NATION N1,NATION N2,STAFF S1,STAFF S2 " & _
             "WHERE CP05>=19110000+" & txt1(1).Text & " AND CP05<=19110000+" & txt1(2).Text & " AND CP159=0 AND SUBSTR(CP60,1,1)='E' " & stStr & _
             "AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) " & _
             "AND SUBSTR(LC22,1,8)=FA01(+) AND SUBSTR(LC22,9,1)=FA02(+) AND FA10=N1.NA01(+) " & _
             "AND SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) AND CU10=N2.NA01(+) " & _
             "AND CP14=S1.ST01(+) AND CP29=S2.ST01(+) AND CP12=A0901(+) "
    '抓請款單
    strSql = strSql & "Union All SELECT /*+ INDEX(ACC1K0 IDXA1K02) */ A1K02 A00,CP09 A01,CP01||'-'||CP02||DECODE(CP03,'0','','-'||CP03)||DECODE(CP04,'00','','-'||CP04) A02 " & _
             ",CPM03 A03,LC13 A04,NVL(CU04,NVL(CU05,CU06)) A05,CP40||' '||CP41||' '||CP42 A05_1,((A1K11-NVL(A1K06,0)-NVL(A1K09,0))/1000) A06,S1.ST02 A07,S2.ST02 A08,DECODE(CP01,'CFL',A0902,NVL(N1.NA03,N2.NA03)) A09 " & _
             ",NVL(NVL(FA05,FA04),FA06) A10,CP60,CP01,CP02,CP03,CP04,CP12,CP14,NVL(N1.NA01,N2.NA01) NA01 " & _
             "FROM ACC1K0 ,CASEPROGRESS,CASEPROPERTYMAP,LAWCASE,FAGENT,CUSTOMER,ACC090,NATION N1,NATION N2,STAFF S1,STAFF S2 " & _
             "WHERE A1K02>=" & txt1(1).Text & " AND A1K02<=" & txt1(2).Text & "  AND NVL(A1K12||A1K25,'N')='N' AND A1K01=CP60 AND CP159=0 " & stStr & _
             "AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) " & _
             "AND SUBSTR(LC22,1,8)=FA01(+) AND SUBSTR(LC22,9,1)=FA02(+) AND FA10=N1.NA01(+) " & _
             "AND SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) AND CU10=N2.NA01(+) " & _
             "AND CP14=S1.ST01(+) AND CP29=S2.ST01(+) AND CP12=A0901(+) "
    
    strSql = "SELECT * FROM (" & strSql & ") WHERE 1=1" & stStr1
    If opt1(0).Value = True Then
       strSql = strSql & "ORDER BY A00,A02,A01 "
    Else
       strSql = strSql & "ORDER BY CP14,A00,A02,A01 "
    End If

    iPage = 0: strGrp = ""
    Erase strTemp
    
    m_SubFclLin = 0: m_SubCfl = 0
    m_TotFclLin = 0: m_TotCfl = 0
    inJ = 1
    Set RsQ = ClsLawReadRstMsg(inJ, strSql)
    If inJ = 1 Then
        InsertQueryLog (RsQ.RecordCount) 'Added by Lydia 2023/04/20
        Printer.PaperSize = 9 'A4
        Printer.Orientation = 2 '橫印
        Printer.Font = "細明體"
        Printer.FontSize = ciFontSize
        lngPageHeight = Printer.ScaleHeight
        lngPageWidth = Printer.ScaleWidth
        lngLineHeight = 300
        iPage = iPage + 1
        If PLeft(1) = 0 Then Call GetPleft
        PrintHeader
        
        With RsQ
            .MoveFirst
            Do While Not .EOF
               '承辦人小計
               If opt1(1).Value = True And strDiv <> "" & .Fields("CP14") And strDiv <> "" Then
                  Call PrintTotal("1")
                  m_SubFclLin = 0: m_SubCfl = 0
               End If
               '請款日期
               strTemp(0) = ChangeTStringToTDateString("" & .Fields("A00"))
               '收文號
               strTemp(1) = "" & .Fields("A01")
               '本所案號
               strTemp(2) = "" & .Fields("A02")
               '案件性質
               strTemp(3) = PUB_StrToStr("" & .Fields("A03"), 16)
               'IP案
               strTemp(4) = "" & .Fields("A04")
               '當事人
               strTemp(5) = PUB_StrToStr("" & .Fields("A05"), 16)
               '對造當事人
               strTemp(cInX + 1) = "" & .Fields("A05_1")
               '點數
               strTemp(6) = Format("" & .Fields("A06"), FatDot)
               '同一請款單只在日期+收文號最小者,顯示點數
               If Mid("" & .Fields("CP60"), 1, 1) = "X" And strGrp = "" & .Fields("CP60") Then strTemp(6) = ""
               '承辦人
               strTemp(7) = "" & .Fields("A07")
               '協辦人員
               strTemp(8) = "" & .Fields("A08")
               '國籍/業務區
               strTemp(9) = PUB_StrToStr("" & .Fields("A09"), 10)
               'FC代理人
               strTemp(10) = PUB_StrToStr("" & .Fields("A10"), 30)
                              
               '列印明細
                For inJ = 0 To cInX + 1
                   If inJ > cInX Then  '對造當事人
                      If Trim(strTemp(inJ)) <> "" Then
                         PrintNewLine
                         Printer.CurrentX = PLeft(5)
                         Printer.CurrentY = iPrint
                         Printer.Print strTemp(inJ)
                      End If
                   Else
                      Printer.CurrentX = PLeft(inJ)
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp(inJ)
                   End If
                Next inJ
                
               PrintNewLine
               
               Select Case "" & .Fields("CP01")
                   Case "FCL", "LIN"
                       m_SubFclLin = m_SubFclLin + Val(strTemp(6))
                       m_TotFclLin = m_TotFclLin + Val(strTemp(6))
                   Case "CFL"
                       m_SubCfl = m_SubCfl + Val(strTemp(6))
                       m_TotCfl = m_TotCfl + Val(strTemp(6))
               End Select
               strGrp = "" & .Fields("CP60")
               strDiv = "" & .Fields("CP14")
               .MoveNext
            Loop
            
            If opt1(1).Value = True And strGrp <> "" Then
               Call PrintTotal("1")
               m_SubFclLin = 0: m_SubCfl = 0
            End If
            '合計
            Call PrintTotal("2")
        End With
        Printer.EndDoc
        ShowPrintOk
    Else
        InsertQueryLog (0) 'Added by Lydia 2023/04/20
        MsgBox "查無資料!!"
    End If

    Set RsQ = Nothing
End Sub

Private Sub PrintTotal(ByVal aKind As String)
    Call ShowLine
    If aKind = "1" Then
       Printer.CurrentX = PLeft(3) + 400
       Printer.CurrentY = iPrint
       Printer.Print "小　計:"
       Printer.CurrentX = PLeft(6)
       Printer.CurrentY = iPrint
       Printer.Print Format(m_SubFclLin + m_SubCfl, FatDot) & " 點"
       PrintNewLine
    End If

    If aKind = "2" Then
       Printer.CurrentX = PLeft(5)
       Printer.CurrentY = iPrint
       Printer.Print "FCL+LIN"
       Printer.CurrentX = PLeft(6)
       Printer.CurrentY = iPrint
       Printer.Print Format(m_TotFclLin, FatDot) & " 點"
       PrintNewLine
       Printer.CurrentX = PLeft(5)
       Printer.CurrentY = iPrint
       Printer.Print "CFL"
       Printer.CurrentX = PLeft(6)
       Printer.CurrentY = iPrint
       Printer.Print Format(m_TotCfl, FatDot) & " 點"
       PrintNewLine
       Call ShowLine
       Printer.CurrentX = PLeft(3) + 400
       Printer.CurrentY = iPrint
       Printer.Print "合　計:"
       Printer.CurrentX = PLeft(6)
       Printer.CurrentY = iPrint
       Printer.Print Format(m_TotFclLin + m_TotCfl, FatDot) & " 點"
    End If
End Sub

'換行判斷
Private Sub PrintNewLine(Optional ByVal mRate As Single = 1, Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)
   iPrint = iPrint + lngLineHeight * mRate
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint

      iPage = iPage + 1
      Printer.NewPage
      PrintHeader
   End If
End Sub

'列印表頭
Private Sub PrintHeader()
Dim iPos As Integer

iPrint = ciStartY

Printer.Font.Size = ciTitleFontSize
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth("請款點數明細表")) / 2
Printer.CurrentY = iPrint
Printer.Print "請款點數明細表"

Printer.Font.Size = ciFontSize
Printer.Font.Bold = False
Printer.Font.Underline = False
PrintNewLine

iPrint = iPrint + 150
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "系統類別：" & IIf(Right(txt1(0), 1) = ",", Mid(txt1(0), 1, Len(txt1(0)) - 1), txt1(0))
Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(mRptStr1)) / 2
Printer.CurrentY = iPrint
Printer.Print mRptStr1

PrintNewLine
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "列印人員：" & strUserName
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

PrintNewLine
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "列印順序：" & IIf(opt1(0).Value = True, "日期", "承辦人")
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & Printer.Page

PrintNewLine
Call ShowLine

'列印欄位抬頭
For iPos = 0 To cInX + 1
   If iPos > cInX Then '對造當事人
      PrintNewLine
      Printer.CurrentX = PLeft(5)
   Else
      Printer.CurrentX = PLeft(iPos)
   End If
   
   Printer.CurrentY = iPrint
   Printer.Print PTitle(iPos)
Next iPos

PrintNewLine
Call ShowLine

End Sub

Private Sub ShowLine()
    iPrint = iPrint - 100
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(145, "-")
    iPrint = iPrint + 200
End Sub

Private Sub GetPleft()

   Erase PTitle
   PTitle(0) = "請款日期"
   PTitle(1) = "收文號"
   PTitle(2) = "本所案號"
   PTitle(3) = "案件性質"
   PTitle(4) = "IP案"
   PTitle(5) = "當事人"
   PTitle(cInX + 1) = "對造當事人"
   PTitle(6) = "點數"
   PTitle(7) = "承辦人"
   PTitle(8) = "協辦人員"
   PTitle(9) = "國籍/業務區"
   PTitle(10) = "FC代理人"
      
   Erase PLeft
   PLeft(0) = ciStartX
   PLeft(1) = PLeft(0) + Printer.TextWidth(String(5, "　")) + ciColGap
   PLeft(2) = PLeft(1) + Printer.TextWidth(String(5, "　")) + ciColGap
   PLeft(3) = PLeft(2) + Printer.TextWidth(String(6, "　")) + ciColGap
   PLeft(4) = PLeft(3) + Printer.TextWidth(String(8, "　")) + ciColGap
   PLeft(5) = PLeft(4) + Printer.TextWidth(String(2, "　")) + ciColGap
   PLeft(6) = PLeft(5) + Printer.TextWidth(String(8, "　")) + ciColGap
   PLeft(7) = PLeft(6) + Printer.TextWidth(String(4, "　")) + ciColGap
   PLeft(8) = PLeft(7) + Printer.TextWidth(String(4, "　")) + ciColGap
   PLeft(9) = PLeft(8) + Printer.TextWidth(String(4, "　")) + ciColGap
   PLeft(10) = PLeft(9) + Printer.TextWidth(String(6, "　")) + ciColGap
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    txt1(0).Text = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm083016 = Nothing
End Sub


Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Dim strTemp1 As Variant, strTemp2 As Variant
Dim i As Integer, j As Integer

    Select Case Index
    Case 0
         strTemp1 = Split(Replace(GetSystemKindByNick, ",,", ""), ",")
         strTemp2 = Split(Replace(UCase(txt1(0).Text), ",,", ""), ",")
         For i = 0 To UBound(strTemp2)
            s = 0
            For j = 0 To UBound(strTemp1)
                If strTemp2(i) = strTemp1(j) Then
                    s = 1
                    Exit For
                End If
            Next j
            If s = 0 Then
                s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限!! ", , "USER 權限問題")
                txt1(0).SetFocus
                txt1(0).SelStart = 0
                txt1(0).SelLength = Len(txt1(0))
                Exit Sub
            End If
         Next i
    Case 1, 2
         If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
            Me.txt1(Index).SetFocus
            txt1_GotFocus Index
            Exit Sub
         End If
         If Index = 2 Then
            GoTo JumpCheck
         End If
         
    Case 9
         If txt1(Index) <> "" Then txt1(Index + 1) = txt1(Index) & "Z"
    Case 4, 10
JumpCheck:
         If txt1(Index) <> "" And txt1(Index - 1) <> "" Then
            If txt1(Index) < txt1(Index - 1) Then
               s = MsgBox("範圍 起始必須小於終止!!", , "USER 輸入錯誤")
               txt1(Index - 1).SetFocus
               txt1_GotFocus (Index - 1)
               Exit Sub
            End If
         End If
         
    Case 6
         Lbl1(1) = GetPrjSalesNM(txt1(6))
         If Trim(txt1(Index)) <> "" Then
            If Trim(Lbl1(1).Caption) = "" Then
                s = MsgBox("承辦人輸入錯誤！", , "錯誤！")
                txt1(Index).SetFocus
                txt1_GotFocus (Index)
                Exit Sub
            End If
         End If
    
    Case Else
    End Select
End Sub


