VERSION 5.00
Begin VB.Form frm170207 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工薪資明細表"
   ClientHeight    =   2196
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4668
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2196
   ScaleWidth      =   4668
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   945
      Style           =   2  '單純下拉式
      TabIndex        =   7
      Top             =   1590
      Width           =   2460
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1860
      TabIndex        =   2
      Text            =   "5"
      Top             =   1110
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1110
      TabIndex        =   1
      Text            =   "96"
      Top             =   1110
      Width           =   435
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   1110
      TabIndex        =   0
      Text            =   "1"
      Top             =   600
      Width           =   285
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   3075
      TabIndex        =   4
      Top             =   30
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1995
      TabIndex        =   3
      Top             =   30
      Width           =   975
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
      Height          =   255
      Left            =   180
      TabIndex        =   8
      Top             =   1620
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "薪資月份：            年         月"
      Height          =   180
      Left            =   180
      TabIndex        =   6
      Top             =   1170
      Width           =   2205
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "報表格式：        (1.依部門  2.依公司  3.依公司+職稱)"
      Height          =   180
      Left            =   180
      TabIndex        =   5
      Top             =   660
      Width           =   4035
   End
End
Attribute VB_Name = "frm170207"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Create by Morgan 2009/1/19
Option Explicit

Dim PLeft() As Integer, PColName() As String, strTemp() As String
Dim iPrint As Integer, iPage As Integer
Dim m_iStartX As Integer, m_iStartY As Integer, m_iColGap As Integer
Dim m_iPageHeight As Long, m_iLineHeight As Long, m_iMargin As Long
Dim m_RptType As String, m_Group As String
Dim m_Tot(15, 2) As String, m_SubTot(15, 2) As String
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
         PrintReport
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
   FormReset
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170207 = Nothing
End Sub

Private Sub PrintReport()
   Dim YM As String
   Dim adoRst As ADODB.Recordset
   
   YM = 100 * Val(Text1(1)) + Val(Text1(2)) + 191100
   
   If Text1(0) = "1" Then
      'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
      '2012/1/17 MODIFY BY SONIA 外專工程師加依組別排序
      'strExc(0) = "select min(st06) st06,sm03,st01,max(st02) st01N,max(a0902) GP" & _
         ",sum(sm04) sm04,sum(sm05) sm05,sum(sm06) sm06,sum(sm07) sm07,sum(sm08) sm08" & _
         ",sum(sm09) sm09,sum(sm10) sm10,sum(sm11) sm11,sum(sm12) sm12,sum(sm13) sm13" & _
         ",sum(sm19) sm19,sum(sm20) sm20,sum(sm21) sm21,sum(sm22) sm22,sum(sm23) sm23,sum(sm24) sm24" & _
         ",sum(sm14) sm14,sum(sm15) sm15,sum(sm16) sm16,sum(sm17) sm17,sum(sm18) sm18" & _
         ",sum(nvl(sm04,0)+nvl(sm05,0)+nvl(sm06,0)+nvl(sm07,0)+nvl(sm08,0)" & _
         "+nvl(sm09,0)+nvl(sm10,0)+nvl(sm12,0)+nvl(sm13,0)) s1" & _
         ",sum(nvl(sm14,0)+nvl(sm15,0)+nvl(sm16,0)+nvl(sm17,0)+nvl(sm18,0)+nvl(sm19,0)" & _
         "+nvl(sm20,0)+nvl(sm21,0)+nvl(sm22,0)+nvl(sm23,0)+nvl(sm24,0)) s2" & _
         ",sum(sm30) sm30" & _
         " from salarymonth,staff,acc090" & _
         " where st01(+)=substr(sm01,1,1)||replace(substr(sm01,2),'A','0') and a0901(+)=sm03" & _
         " and sm02=" & YM & " group by sm03,st01"
      'Modified by Morgan 2013/2/1 +sm43 補充保費
      'MODIFY BY SONIA 2014/4/3 不含F51國外部外翻
      'Modify By Sindy 2020/6/22 + 證照津貼
      'Modified by Morgan 2023/12/26 +新部門
      If YM >= Left(新部門啟用日, 6) Then
         strExc(1) = "max(a0922)"
         strExc(2) = ""
      Else
         strExc(1) = "max(a0902)||decode(sm03,'F21',cst16(st16),null)"
         strExc(2) = ",decode(sm03,'F21',cst16(st16),null)"
      End If
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      strExc(0) = "select min(st06) st06,sm03,st01,max(st02) st01N,max(decode(st22,'M','男','F','女',st22)) st22,max(st20||' '||substr(ac03,1,5)) st20," & strExc(1) & " GP" & _
         ",sum(sm04) sm04,sum(sm05) sm05,sum(sm06) sm06,sum(sm07) sm07,sum(sm08) sm08" & _
         ",sum(sm09) sm09,sum(sm10) sm10,sum(sm11) sm11,sum(sm12) sm12,sum(sm13) sm13" & _
         ",sum(sm19) sm19,sum(sm20) sm20,sum(sm21) sm21,sum(sm22) sm22,sum(sm23) sm23,sum(sm24) sm24" & _
         ",sum(sm14) sm14,sum(sm15) sm15,sum(sm16) sm16,sum(sm17) sm17,sum(sm18) sm18,sum(sm45) sm45" & _
         ",sum(nvl(sm04,0)+nvl(sm05,0)+nvl(sm06,0)+nvl(sm07,0)+nvl(sm08,0)+nvl(sm45,0)" & _
         "+nvl(sm09,0)+nvl(sm10,0)+nvl(sm12,0)+nvl(sm13,0)) s1" & _
         ",sum(nvl(sm14,0)+nvl(sm15,0)+nvl(sm16,0)+nvl(sm17,0)+nvl(sm18,0)+nvl(sm19,0)" & _
         "+nvl(sm20,0)+nvl(sm21,0)+nvl(sm22,0)+nvl(sm23,0)+nvl(sm24,0)+nvl(sm43,0)) s2" & _
         ",sum(sm30) sm30,sum(sm43) sm43" & _
         " from salarymonth,staff,acc090,acc090new,allcode" & _
         " where st01(+)=substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4) and a0901(+)=sm03 and a0921(+)=sm03 and sm03<>'F51' and '01'=ac01(+) and st20=ac02(+)" & _
         " and sm02=" & YM & " group by sm03" & strExc(2) & ",st01"
   'end 2023/12/26
   'add by sonia 2016/8/26 依公司+職稱
   ElseIf Text1(0) = "3" Then
      'Modify By Sindy 2020/6/22 + 證照津貼
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      strExc(0) = "select sm37,min(st06) st06,sm03,st01,max(st02) st01N,max(decode(st22,'M','男','F','女',st22)) st22,st20||' '||max(substr(ac03,1,5)) GP" & _
         ",sum(sm04) sm04,sum(sm05) sm05,sum(sm06) sm06,sum(sm07) sm07,sum(sm08) sm08" & _
         ",sum(sm09) sm09,sum(sm10) sm10,sum(sm11) sm11,sum(sm12) sm12,sum(sm13) sm13" & _
         ",sum(sm19) sm19,sum(sm20) sm20,sum(sm21) sm21,sum(sm22) sm22,sum(sm23) sm23,sum(sm24) sm24" & _
         ",sum(sm14) sm14,sum(sm15) sm15,sum(sm16) sm16,sum(sm17) sm17,sum(sm18) sm18,sum(sm45) sm45" & _
         ",sum(nvl(sm04,0)+nvl(sm05,0)+nvl(sm06,0)+nvl(sm07,0)+nvl(sm08,0)+nvl(sm45,0)" & _
         "+nvl(sm09,0)+nvl(sm10,0)+nvl(sm12,0)+nvl(sm13,0)) s1" & _
         ",sum(nvl(sm14,0)+nvl(sm15,0)+nvl(sm16,0)+nvl(sm17,0)+nvl(sm18,0)+nvl(sm19,0)" & _
         "+nvl(sm20,0)+nvl(sm21,0)+nvl(sm22,0)+nvl(sm23,0)+nvl(sm24,0)+nvl(sm43,0)) s2" & _
         ",sum(sm30) sm30,sum(sm43) sm43" & _
         " from salarymonth,staff,acc080,allcode" & _
         " where st01(+)=substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4) and a0801(+)=sm37 and sm03<>'F51' and '01'=ac01(+) and st20=ac02(+)" & _
         " and sm02=" & YM & " group by sm37,st20,st22,sm03,st01"
   'end 2016/8/26
   Else
      'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
      'Modified by Morgan 2013/2/1 +sm43 補充保費
      'MODIFY BY SONIA 2014/4/3 不含F51國外部外翻
      'Modify By Sindy 2020/6/22 + 證照津貼
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      strExc(0) = "select sm37,min(st06) st06,sm03,st01,max(st02) st01N,max(decode(st22,'M','男','F','女',st22)) st22,max(st20||' '||substr(ac03,1,5)) st20,max(a0802) GP" & _
         ",sum(sm04) sm04,sum(sm05) sm05,sum(sm06) sm06,sum(sm07) sm07,sum(sm08) sm08" & _
         ",sum(sm09) sm09,sum(sm10) sm10,sum(sm11) sm11,sum(sm12) sm12,sum(sm13) sm13" & _
         ",sum(sm19) sm19,sum(sm20) sm20,sum(sm21) sm21,sum(sm22) sm22,sum(sm23) sm23,sum(sm24) sm24" & _
         ",sum(sm14) sm14,sum(sm15) sm15,sum(sm16) sm16,sum(sm17) sm17,sum(sm18) sm18,sum(sm45) sm45" & _
         ",sum(nvl(sm04,0)+nvl(sm05,0)+nvl(sm06,0)+nvl(sm07,0)+nvl(sm08,0)+nvl(sm45,0)" & _
         "+nvl(sm09,0)+nvl(sm10,0)+nvl(sm12,0)+nvl(sm13,0)) s1" & _
         ",sum(nvl(sm14,0)+nvl(sm15,0)+nvl(sm16,0)+nvl(sm17,0)+nvl(sm18,0)+nvl(sm19,0)" & _
         "+nvl(sm20,0)+nvl(sm21,0)+nvl(sm22,0)+nvl(sm23,0)+nvl(sm24,0)+nvl(sm43,0)) s2" & _
         ",sum(sm30) sm30,sum(sm43) sm43" & _
         " from salarymonth,staff,acc080,allcode" & _
         " where st01(+)=substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4) and a0801(+)=sm37 and sm03<>'F51' and '01'=ac01(+) and st20=ac02(+)" & _
         " and sm02=" & YM & " group by sm37,sm03,st01"
   End If
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
   Dim iRecs As Integer, strDept As String, ii As Integer
   Dim strFontName As String, strFontSize As String
   
   strFontName = Printer.FontName
   strFontSize = Printer.FontSize
   
   GetPleft
On Error GoTo ErrHnd
   With p_Rst
      .MoveFirst
      m_Group = .Fields("GP")
      iPage = 0
      iRecs = 0
      Erase m_Tot
      Erase m_SubTot
      PrintPageHeader
      PrintPageHeader1
      Do While Not .EOF
         If m_Group <> .Fields("GP") Then
            PrintSubTotal
            m_Group = .Fields("GP")
            iRecs = 0
            Erase m_SubTot
            Printer.NewPage
            PrintPageHeader
            PrintPageHeader1
         End If
         iRecs = iRecs + 1
         strTemp(1, 1) = "" & .Fields("st01")
         strTemp(2, 1) = "" & .Fields("st01N")
         'Modified by Morgan 2013/2/1 +補充保費調整位置
         strTemp(2, 2) = Format("" & .Fields("sm20"), "#,###") '借支扣款
         strTemp(3, 1) = Format("" & .Fields("sm04"), "#,###") '基本薪資
         strTemp(3, 2) = Format("" & .Fields("sm15"), "#,###") '健 保 費
         strTemp(4, 1) = Format("" & .Fields("sm07"), "#,###") '午餐津貼
         strTemp(4, 2) = Format("" & .Fields("sm43"), "#,###") '補充保費
         strTemp(5, 1) = Format("" & .Fields("sm08"), "#,###") '差旅津貼
         strTemp(5, 2) = Format("" & .Fields("sm14"), "#,###") '勞 保 費
         strTemp(6, 1) = Format("" & .Fields("sm05"), "#,###") '職務津貼
         strTemp(6, 2) = Format("" & .Fields("sm24"), "#,###") '所 得 稅
         strTemp(7, 1) = Format("" & .Fields("sm45"), "#,###") '證照津貼 Modify By Sindy 2020/6/22
         strTemp(7, 2) = Format("" & .Fields("sm16"), "#,###") '退休自提
         strTemp(8, 1) = Format("" & .Fields("sm09"), "#,###") '房租津貼
         strTemp(8, 2) = Format("" & .Fields("sm19"), "#,###") '員工借款
         strTemp(9, 1) = Format("" & .Fields("sm06"), "#,###") '技術津貼 Modify By Sindy 2020/6/22
         strTemp(9, 2) = Format("" & .Fields("sm18"), "#,###") '互 助 會
         strTemp(10, 1) = Format("" & .Fields("sm10"), "#,###") '特 支 費 Modify By Sindy 2020/6/22
         strTemp(10, 2) = Format("" & .Fields("sm17"), "#,###") '婚喪扣款
         strTemp(11, 1) = Format("" & .Fields("sm11"))          '加班時數
         strTemp(11, 2) = Format("" & .Fields("sm21"), "#,###") '缺勤扣款
         strTemp(12, 1) = Format("" & .Fields("sm12"), "#,###") '加 班 費
         strTemp(12, 2) = Format("" & .Fields("sm22"), "#,###") '未 打 卡
         strTemp(13, 1) = Format("" & .Fields("sm13"), "#,###") '其他所得
         strTemp(13, 2) = Format("" & .Fields("sm23"), "#,###") '其他扣款
         strTemp(14, 1) = Format("" & .Fields("s1"), "#,###") '應發金額
         strTemp(14, 2) = Format("" & .Fields("s2"), "#,###") '應扣金額
         strTemp(15, 1) = Format(Val("" & .Fields("s1")) - Val("" & .Fields("s2")), "#,###")
         For ii = 2 To 15
            If ii > 2 Then
               m_Tot(ii, 1) = Val(m_Tot(ii, 1)) + Val(Format(strTemp(ii, 1)))
               m_SubTot(ii, 1) = Val(m_SubTot(ii, 1)) + Val(Format(strTemp(ii, 1)))
            End If
            m_Tot(ii, 2) = Val(m_Tot(ii, 2)) + Val(Format(strTemp(ii, 2)))
            m_SubTot(ii, 2) = Val(m_SubTot(ii, 2)) + Val(Format(strTemp(ii, 2)))
         Next
         PrintDetail strTemp
         .MoveNext
      Loop
      PrintSubTotal
      PrintTotal
      Printer.EndDoc
   End With
   doPrint = True

ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   Printer.FontName = strFontName
   Printer.FontSize = strFontSize
End Function

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
   Printer.Orientation = 2
   Printer.FontSize = 11
   m_iStartX = 0
   m_iStartY = 500
   m_iPageHeight = Printer.ScaleHeight
   m_iLineHeight = 300
   m_iMargin = (Printer.Height - Printer.ScaleHeight) / 2
   m_iColGap = 200 '欄位間隔
   
   ReDim PLeft(16)
   ReDim PColName(15, 2)
   ReDim strTemp(15, 2)
   
   ii = 1
   PColName(ii, 1) = "員工號"
   PColName(ii, 2) = "扣除項目"
   PLeft(ii) = m_iStartX
   
   ii = ii + 1
   PColName(ii, 1) = "姓　　名"
   PColName(ii, 2) = "借支扣款"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(3, "　")) + m_iColGap
   
   ii = ii + 1
   PColName(ii, 1) = "基本薪資"
   PColName(ii, 2) = "健 保 費"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(4, "　")) + m_iColGap
   
   ii = ii + 1
   PColName(ii, 1) = "午餐津貼"
   PColName(ii, 2) = "補充保費"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(4, "　")) + m_iColGap
   
   ii = ii + 1
   PColName(ii, 1) = "差旅津貼"
   PColName(ii, 2) = "勞 保 費"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(4, "　")) + m_iColGap
   
   ii = ii + 1
   PColName(ii, 1) = "職務津貼"
   PColName(ii, 2) = "所 得 稅"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(4, "　")) + m_iColGap
   
   ii = ii + 1
   PColName(ii, 1) = "證照津貼" '技證津貼 Modify By Sindy 2020/6/22
   PColName(ii, 2) = "退休自提"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(4, "　")) + m_iColGap
   
   ii = ii + 1
   PColName(ii, 1) = "房租津貼"
   PColName(ii, 2) = "員工借款"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(4, "　")) + m_iColGap
   
   ii = ii + 1
   PColName(ii, 1) = "技術津貼" 'Modify By Sindy 2020/6/22
   PColName(ii, 2) = "互 助 會"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(4, "　")) + m_iColGap
   
   ii = ii + 1
   PColName(ii, 1) = "特 支 費" 'Modify By Sindy 2020/6/22
   PColName(ii, 2) = "婚喪扣款"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(4, "　")) + m_iColGap
   
   ii = ii + 1
   PColName(ii, 1) = "加班時數"
   PColName(ii, 2) = "缺勤扣款"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(4, "　")) + m_iColGap
   
   ii = ii + 1
   PColName(ii, 1) = "加 班 費"
   PColName(ii, 2) = "未 打 卡"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(4, "　")) + m_iColGap
   
   ii = ii + 1
   PColName(ii, 1) = "其他所得"
   PColName(ii, 2) = "其他扣款"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(4, "　")) + m_iColGap
   
   ii = ii + 1
   PColName(ii, 1) = "應發金額"
   PColName(ii, 2) = "應扣金額"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(4, "　")) + m_iColGap
   
   ii = ii + 1
   PColName(ii, 2) = "實發金額"
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(5, "　")) + m_iColGap
   
   ii = ii + 1
   PLeft(ii) = PLeft(ii - 1) + Printer.TextWidth(String(5, "　")) + m_iColGap
End Sub

Private Sub PrintPageHeader()
   Dim strTmp As String
   
   strTmp = "台一關係企業 員工薪資明細表"
   If Text1(0) = "1" Then
      strTmp = strTmp & " ( 部門別 )"
   'add by sonia 2016/8/26
   ElseIf Text1(0) = "3" Then
      strTmp = strTmp & " ( 公司別+職稱 )"
   Else
      strTmp = strTmp & " ( 公司別 )"
   End If
   iPrint = m_iStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = 20
   Printer.Font.Bold = True
   'Printer.Font.Underline = True
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strTmp
   iPrint = iPrint + 600
   
   Printer.Font.Size = 11
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   If Text1(0) = "1" Then
      Printer.Print "部　門：" & m_Group
   Else
      Printer.Print "公　司：" & m_Group
   End If
   
   Printer.CurrentX = m_iStartX
   strTmp = "薪資月份：" & Text1(1) & " 年 " & Text1(2) & " 月"
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strTmp)) / 2
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
   Printer.Font.Size = 11
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
    Printer.Font.Size = 11
    PrintNewLine 3
    For iCol = 1 To UBound(strData)
      If iCol > 2 Then
        Printer.CurrentX = PLeft(iCol + 1) - Printer.TextWidth(strData(iCol, 1)) - m_iColGap
        Printer.CurrentY = iPrint
        Printer.Print strData(iCol, 1)
        
        Printer.CurrentX = PLeft(iCol + 1) - Printer.TextWidth(strData(iCol, 2)) - m_iColGap
        Printer.CurrentY = iPrint + m_iLineHeight
        Printer.Print strData(iCol, 2)
      Else
        If strData(iCol, 1) <> "" Then
            Printer.CurrentX = PLeft(iCol)
            Printer.CurrentY = iPrint
            'Modified by Morgan 2023/12/26
            'Printer.Print strData(iCol, 1)
            PUB_PrintUnicodeText strData(iCol, 1), Printer.CurrentX, Printer.CurrentY, 0
            'end 2023/12/26
        End If
        If iCol = 2 And strData(iCol, 2) <> "" Then
            Printer.CurrentX = PLeft(iCol + 1) - Printer.TextWidth(strData(iCol, 2)) - m_iColGap
            Printer.CurrentY = iPrint + m_iLineHeight
            Printer.Print strData(iCol, 2)
        End If
      End If
    Next
    iPrint = iPrint + m_iLineHeight
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
      'PrintMemo
      Printer.NewPage
      PrintPageHeader
      PrintPageHeader1
      iPrint = iPrint + m_iLineHeight
   End If
   
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   CloseIme
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
   Case 0
      If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") Then
         KeyAscii = 0
         Beep
      End If
      
   Case Else
      If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
         KeyAscii = 0
         Beep
      End If
   End Select
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean, oText As TextBox
   
   If Text1(0) = "" Then
      MsgBox "報表格式不可空白 !"
      Text1(0).SetFocus
      Exit Function
   End If
   If Text1(1) = "" Then
      MsgBox "年度不可空白 !"
      Text1(1).SetFocus
      Exit Function
   End If
   If Text1(2) = "" Then
      MsgBox "月份不可空白 !"
      Text1(2).SetFocus
      Exit Function
   End If
   
   For Each oText In Text1
      Text1_Validate oText.Index, bCancel
      If bCancel = True Then
         Text1(oText.Index).SetFocus
         Exit Function
      End If
   Next
   TxtValidate = True
End Function

Private Sub FormReset()
   Dim stDate As String
   If Val(Right(strSrvDate(2), 2)) < 11 Then
      stDate = CompDate("1", -1, strSrvDate(1)) - 19110000
   Else
      stDate = strSrvDate(2)
   End If
   Text1(0).Text = ""
   Text1(1).Text = stDate \ 10000
   Text1(2).Text = Val(Right(stDate \ 100, 2))
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 2
      If Val(Text1(Index)) < 1 Or Val(Text1(Index)) > 12 Then
         Cancel = True
         Text1_GotFocus Index
         MsgBox "月份輸入錯誤 !"
      End If
   End Select
End Sub

'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)
   PrintNewLine 3
   DrawLine
   PrintNewLine
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print "合計： " & iRecCount & " 筆"
End Sub

'列印小計
Private Sub PrintSubTotal()
   Dim ii As Integer
   PrintNewLine
   DrawLine
   If Text1(0) = "1" Then
      m_SubTot(1, 1) = m_Group & " 小計："
      '2012/1/17 add by sonia 因外專工程師加印組別但小計字數有限故只印組別
      If Left(m_Group, 5) = "外專工程師" Then
         m_SubTot(1, 1) = Mid(m_Group, 6) & " 小計："
      End If
      '2012/1/17 end
   Else
      m_SubTot(1, 1) = "　　小計："
   End If
   For ii = 2 To 15
      If ii > 2 Then
         m_SubTot(ii, 1) = Format(m_SubTot(ii, 1), "#,###")
      End If
      m_SubTot(ii, 2) = Format(m_SubTot(ii, 2), "#,###")
   Next
   PrintDetail m_SubTot
End Sub

'列印合計
Private Sub PrintTotal()
   Dim ii As Integer
   PrintNewLine
   DrawLine
   m_Tot(1, 1) = "　　合計："
   For ii = 2 To 15
      If ii > 2 Then
         m_Tot(ii, 1) = Format(m_Tot(ii, 1), "#,###")
      End If
      m_Tot(ii, 2) = Format(m_Tot(ii, 2), "#,###")
   Next
   PrintDetail m_Tot
End Sub
