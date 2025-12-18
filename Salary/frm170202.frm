VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm170202 
   BorderStyle     =   1  '單線固定
   Caption         =   "薪資調薪表"
   ClientHeight    =   3156
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6264
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3156
   ScaleWidth      =   6264
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   13
      Text            =   "1"
      Top             =   2370
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   1170
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   1
      Top             =   1200
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   645
      Index           =   1
      Left            =   1170
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1620
      Width           =   3975
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   1170
      Style           =   2  '單純下拉式
      TabIndex        =   3
      Top             =   2760
      Width           =   2460
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   0
      Text            =   "1"
      Top             =   630
      Width           =   255
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   4230
      TabIndex        =   5
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "產生(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3150
      TabIndex        =   4
      Top             =   90
      Width           =   975
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "產生內容：          (1.報表 2.調薪比率XLS 3.核對比率XLS)"
      Height          =   180
      Left            =   210
      TabIndex        =   12
      Top             =   2430
      Width           =   4395
   End
   Begin VB.Label Label5 
      Caption         =   "( 1.高級專員以下  2.高級專員(含)以上 3.關係企業      4.個人-不限制年資)"
      Height          =   380
      Left            =   1560
      TabIndex        =   11
      Top             =   690
      Width           =   4000
   End
   Begin MSForms.Label Label4 
      Height          =   285
      Left            =   2010
      TabIndex        =   10
      Top             =   1230
      Width           =   1215
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Left            =   210
      TabIndex        =   9
      Top             =   1260
      Width           =   900
   End
   Begin VB.Label Label15 
      Caption         =   "印表機："
      Height          =   180
      Left            =   210
      TabIndex        =   8
      Top             =   2790
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "呈閱日期："
      Height          =   180
      Left            =   210
      TabIndex        =   7
      Top             =   1650
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "調薪對象：        "
      Height          =   180
      Left            =   210
      TabIndex        =   6
      Top             =   690
      Width           =   1260
   End
End
Attribute VB_Name = "frm170202"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/4/18 Form2.0已修改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Create by Morgan 2009/2/12
Option Explicit

Dim m_DefaultPrinter As String
Dim iLineHeight As Integer, TwPerCm As Integer
Dim Px(16) As Long, Py(24) As Long
Dim iPage As Integer


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
   Case 0
      Screen.MousePointer = vbHourglass
      If TxtValidate = True Then
         Me.Enabled = False
         If cmbPrinter <> Printer.DeviceName Then PUB_RestorePrinter cmbPrinter
         
         PrintSheet
         
         '若印表機變動, 則更新列印設定
         If cmbPrinter.Tag <> cmbPrinter Then
            PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, 0, 0, Me.cmbPrinter.Text
         End If
         If Printer.DeviceName <> m_DefaultPrinter Then PUB_RestorePrinter m_DefaultPrinter
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
   Set frm170202 = Nothing
End Sub

Private Function TxtValidate() As Boolean
   'add by sonia 2016/4/11
   If Text1(0) = "4" And Text1(2) = "" Then
      ShowMsg "請輸入員工編號 !"
      Text1(2).SetFocus
      Text1_GotFocus 2
      TxtValidate = False
      Exit Function
   End If
   'end 2016/4/11
   
   'Added by Morgan 2021/4/28
   If Text1(3) = "" Then
      ShowMsg "請輸入產生內容 !"
      Text1(3).SetFocus
      Exit Function
   End If
   'end 2021/4/28
   
   TxtValidate = True
End Function

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   Select Case Index
   Case 0, 2, 3
      CloseIme
   Case 1
      OpenIme
   End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
   Case 0
      If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") And KeyAscii <> Asc("4") And KeyAscii <> Asc("5") Then
         KeyAscii = 0
         Beep
      'add by sonia 2016/4/11
      ElseIf KeyAscii = Asc("4") Then
         Text1(2).Enabled = True
         Text1(2).Locked = False
         
         'Addded by Morgan 2021/4/28
         Text1(3) = "1"
         Text1(3).Enabled = False
         'end 2021/4/28
      Else
         Text1(2) = "": Label4 = ""
         Text1(2).Enabled = False
         Text1(2).Locked = True
         
         'Added by Morgan 2021/4/28
         Text1(3).Enabled = True
         'end 2021/4/28
      End If
   Case 2
      KeyAscii = UpperCase(KeyAscii)
      
   'Added by Morgan 2021/4/28
   Case 3
      If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") Then
         KeyAscii = 0
         Beep
      End If
      
   End Select
End Sub

Private Sub PrintSheet()
Dim stCon As String
Dim strFontSize As String, strFontName As String
Dim Xo As Long, Yo As Long, xi As Long, yi As Long
Dim iMargin As Integer, iMarginY As Integer
Dim adoRst As ADODB.Recordset
Dim PosOld As String, TitOld As String
Dim stAssessList As String, iYear1 As Single, iYear2 As Single, iYear3 As Single
Dim m_Year As String, m_EndDate As String
Dim SL(11 To 39) As Long '26
Dim SO(11 To 26) As Long      'add by sonia 2017/4/7
Dim m_sl02_New As String      'add by sonia 2017/4/7 最近調薪日
Dim m_sl02_1 As String        'add by sonia 2017/4/7 上次調薪日
Dim m_sl02_2 As String        'add by sonia 2017/4/7 前次調薪日
Dim m_sl03_1 As String 'Added by Morgan 2023/5/11 上次調薪編制
Dim m_rpt2Row As Integer 'Added by Morgan 2023/5/11

   Dim xlsReport As New Excel.Application 'Added by Morgan 2020/10/15
   Dim wksReport As Excel.Worksheet 'Added by Morgan 2020/10/15
   Dim wksReport2  As Excel.Worksheet 'Added by Morgan 2023/5/11
   Dim ii As Integer, bolSaveXLS As Boolean, strFileName As String 'Added by Morgan 2020/10/15
   Dim stLsdDeptNo As String, stLsdDeptName As String, iRow1 As Integer, iRow2 As Integer 'Added by Morgan 2020/10/16
   
   '2013/1/24 add by sonia 剔除年資滿二十五年不印(年資算至系統日當月月底)-黃經理
   m_EndDate = strSrvDate(1)
   '2013/10/3 modify by sonia 總經理說改至當年年底, 否則同一年度同仁會有的有調有的沒調,102年時77年進入者不調
   'm_EndDate = DateAdd("M", 1, ChangeWStringToWDateString(m_EndDate))
   'm_EndDate = Format(DateAdd("D", -1 * Day(m_EndDate), m_EndDate), "YYYYMMDD")
   m_EndDate = Left(m_EndDate, 4) & "1231"
   '2013/10/3 end
   '2013/1/24 end
   
   stCon = ""
   Select Case Text1(0)
      Case "1"
         '2009/10/8 MODIFY BY SONIA
         '1. 控制74028編號(含)以前的都不印調薪表
         '2. 沒有職位的應在11月,但外籍顧問在5月
         'stCon = stCon & " and st03<>'R04' and (st21>'22' or st21 is null)"
         '2009/11/2 modify by sonia 辜說76012桂齊恆,79037蔣文正也不印
         '2014/1/24 modify by sonia 77027黃美珍也不印
         '2020/4/8 還原 cancel by sonia 2019/3/26
         stCon = stCon & " AND ST01>'74028' and st01<>'76012' and st01<>'79037' and st01<>'77027' "
         '2020/4/8 還原 end 2019/3/26
         stCon = stCon & " and st03<>'R04' and (st21>'22' or st20='59')"
         stCon = stCon & " and st01<>'99005' "     '2012/5/2 add by sonia
         '2009/10/8 end
      Case "2"
         '2009/10/8 MODIFY BY SONIA
         '1. 控制74028編號(含)以前的都不印調薪表
         '2. 沒有職位的應在11月
         'stCon = stCon & " and st03<>'R04' and st21<='22'"
         '2009/11/2 modify by sonia 辜說76012桂齊恆,79037蔣文正也不印
         '2010/10/6 modify by sonia 11月的林信昌68007還是要印 2010/10/29又不印了
         '2012/10/30 modify by sonia 辜說77015顏裕洋,75007魏天立不印
         '2020/4/8 還原 modify by sonia 2019/3/26
         'Modified by Morgan 2023/11/27 取消 ST01>'74028'條件(25年以上但當年有晉升/任命還是要列)
         stCon = stCon & " AND ((st01<>'76012' and st01<>'79037' and st01<>'77015' and st01<>'75007' "
         stCon = stCon & " and st03<>'R04' and (st21<='22' OR (ST21 IS NULL AND st20<>'59')))"
         'stCon = stCon & " and st03<>'R04' and ((st21<='22' OR (ST21 IS NULL AND st20<>'59'))"
         '2020/4/8 還原 end 2019/3/26
         'stCon = stCon & " AND (ST01>'74028' and st01<>'76012' and st01<>'79037' "
         'stCon = stCon & " and st03<>'R04' and (st21<='22' OR (ST21 IS NULL AND st20<>'59')) or ST01='68007') "
         '2010/10/6 end
         stCon = stCon & " or st01='99005') "     '2012/5/2 add by sonia
      Case "3"
         stCon = stCon & " and st03='R04'"
      Case "4"
         stCon = stCon & " and st01='" & Text1(2) & "'"
   End Select
   
   'Added by Morgan 2023/11/27
   '核對調薪比率只列出該年度有調薪者
   If Text1(3) = "3" Then
      If Text1(0) = "1" Then
         stCon = stCon & " and exists(select * from salarylog where sl01=st01 and sl02=" & Left(strSrvDate(1), 4) & "0501" & ")"
      ElseIf Text1(0) = "2" Then
         stCon = stCon & " and exists(select * from salarylog where sl01=st01 and sl02=" & Left(strSrvDate(1), 4) & "1101" & ")"
      End If
   End If
   'end 2023/11/27
   
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'Modified by Morgan 2024/5/16 剔除第4碼為9的
   'Modified by Morgan 2025/3/18 +新部門
   strExc(0) = "select a0902 dept,a1.ac03 tit,a2.ac03 pos,a3.a0820 comp1,a4.a0820 comp2,a3.a0802 comp1A,a0922 deptnew" & _
      ",x.*,y.* from staff x,salarydata y,acc090,allcode a1,allcode a2,acc080 a3,acc080 a4,acc090new" & _
      " where st01>'6' and st01<'F' and st04='1' and substr(st01,4,1)<>'9'" & _
      " and sd01(+)=st01 and sd01 is not null" & _
      " and a0901(+)=st03 and a0921(+)=st93" & stCon & _
      " and a1.ac02(+)=st20 and a1.ac01(+)='01'" & _
      " and a2.ac02(+)=st21 and a2.ac01(+)='02'" & _
      " and a3.a0801(+)=sd19 and a4.a0801(+)=sd28 order by st03,st01"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      
      'Added by Morgan 2020/10/15
      'Modified by Morgan 2021/4/28
      'If Text1(0) = "1" Or Text1(0) = "2" Then
      '   strFileName = strExcelPath & "調薪比率(" & Text1(0) & ")" & strSrvDate(2) & ".xlsx"
      If Text1(3) = "2" Or Text1(3) = "3" Then
         If Text1(3) = "3" Then
            strFileName = strExcelPath & "核對調薪比率(" & Text1(0) & ")" & strSrvDate(2) & ".xlsx"
         Else
            strFileName = strExcelPath & "調薪比率(" & Text1(0) & ")" & strSrvDate(2) & ".xlsx"
         End If
      'end 2021/4/28
      
         bolSaveXLS = True
         If Dir(strFileName) = "" Then
            If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = "" Then
               MkDir strExcelPath
            End If
         Else
            Kill strFileName
         End If
   
         Set xlsReport = CreateObject("Excel.Application")
         xlsReport.Visible = True
         
         If Text1(3) = "2" Then
            xlsReport.SheetsInNewWorkbook = 2
         Else
            xlsReport.SheetsInNewWorkbook = 1
         End If
         
         xlsReport.Workbooks.add
         Set wksReport = xlsReport.Worksheets(1)
         'Modified by Morgan 2023/5/11 +頁簽名稱,+"額外調薪"欄位
         If Text1(0) = "1" Then
            wksReport.Name = "5月調薪"
            
         ElseIf Text1(0) = "2" Then
            wksReport.Name = "高專調薪"
         End If
         
         wksReport.Range("A1") = "部門"
         wksReport.Range("B1") = "勞退新制"
         wksReport.Range("C1") = "編號"
         wksReport.Range("D1") = "姓名"
         'Added by Morgan 2025/3/18
         wksReport.Range("E1") = "新部門"
         wksReport.Range("F1") = "職稱"
         'end 2025/3/18
         If Text1(3) = "2" Then
            wksReport.Range("G1") = "額外調薪"
            wksReport.Range("H1") = "本薪"
            wksReport.Range("I1") = "津貼"
            wksReport.Range("J1") = "本薪調升金額"
            wksReport.Range("K1") = "本薪調升率"
            wksReport.Range("L1") = "津貼調升金額"
            wksReport.Range("M1") = "含津貼調升金額"
            wksReport.Range("N1") = "含津貼調升率"
            wksReport.Range("O1") = "調升後"
            wksReport.Range("P1") = "證照津貼(執業津貼)" 'Added by Morgan 2024/5/16
            
            wksReport.Range("A1:P1").Font.Bold = True
            
            Set wksReport2 = xlsReport.Worksheets(2)
            wksReport2.Name = "額外調薪明細"
            wksReport2.Range("A1") = "編號"
            wksReport2.Range("B1") = "姓名"
            wksReport2.Range("C1") = "調薪日期"
            wksReport2.Range("D1") = "適用期滿"
            wksReport2.Range("E1") = "本薪"
            wksReport2.Range("F1") = "職務津貼"
            wksReport2.Range("G1") = "技術津貼"
            wksReport2.Range("H1") = "差旅津貼"
            wksReport2.Range("I1") = "房租津貼"
            wksReport2.Range("J1") = "特支費"
            wksReport2.Range("K1") = "證照津貼"
            wksReport2.Range("L1") = "備註"
            wksReport2.Range("A1:L1").Font.Bold = True
            m_rpt2Row = 2
            
         Else
            wksReport.Range("G1") = "本薪"
            wksReport.Range("H1") = "津貼"
            wksReport.Range("I1") = "本薪調升金額"
            wksReport.Range("J1") = "本薪調升率"
            wksReport.Range("K1") = "津貼調升金額"
            wksReport.Range("L1") = "含津貼調升金額"
            wksReport.Range("M1") = "含津貼調升率"
            wksReport.Range("N1") = "調升後"
            wksReport.Range("O1") = "證照津貼(執業津貼)" 'Added by Morgan 2024/5/16
            
            wksReport.Range("A1:O1").Font.Bold = True
         End If
         
         stLsdDeptNo = "" & adoRst.Fields("st03")
         stLsdDeptName = "" & adoRst.Fields("dept")
         ii = 1
         iRow1 = ii
         iRow2 = ii
      End If
      'end 2020/10/15
   
   If Text1(3) = "1" Then 'Added by Morgan 2021/4/28
         
      strFontSize = Printer.FontSize
      strFontName = Printer.FontName
      Printer.EndDoc
      Printer.PaperSize = 9
      Printer.Orientation = 1
      Printer.Font = "細明體"
      Printer.FontSize = 12
      
      TwPerCm = 567
      iMargin = 150
      iMarginY = 150 '100
      iLineHeight = Printer.TextHeight("字") + 150 '80
      Xo = 300
      Yo = 300    '2017/4/10 modify by sonia 因加'註四'多一行故修改,原為500
      
      Erase Px
      Px(0) = Xo
      SetPx
      
      iYear2 = Val(Left(strSrvDate(1), 4))
      
   End If 'Added by Morgan 2021/4/28
   
      iPage = 0 '2013/4/9 add by sonia
      
      With adoRst
      Do While Not .EOF
      
         '2020/4/8 還原  cancel by sonia 2019/3/26
         If Text1(0) <> "4" Then   'add by sonia 2016/4/11 加個人不限制年資
            '2013/1/24 add by sonia 剔除年資滿二十五年不印(年資算至系統日當月月底)-黃經理
            '2013/10/3 modify by sonia 總經理說改至當年年底, 否則同一年度同仁會有的有調有的沒調,102年時77年進入者不調
            m_Year = Trim(CalYear(CheckStr(.Fields("st01")), m_EndDate))
            If Val(m_Year) >= 25 Then
               If ChkThisYearPromote(.Fields("st01")) = False Then 'Added by Morgan 2023/11/27
                  GoTo Nextstep
               End If
            End If
         End If
         '2020/4/8 還原  end 2019/3/26
         'add by sonia 2014/4/3 剔除試用期未滿者
         'Modified by Morgan 2014/4/15 改前月試用期滿也要印--辜 Ex.A2042
         'If Mid(CompDate(1, 1, Val("" & .Fields("st29"))), 1, 6) >= Mid(Val(strSrvDate(1)), 1, 6) Then
         If Mid(CompDate(1, 1, Val("" & .Fields("st29"))), 1, 6) > Mid(Val(strSrvDate(1)), 1, 6) Then
            GoTo Nextstep
         End If
         '2014/4/3 end
         
         iPage = iPage + 1 '2013/4/9 add by sonia
         '2013/1/24 end
         
         m_sl02_New = 0: m_sl02_1 = 0: m_sl02_2 = 0 'add by sonia 2017/4/7
         
         '原薪資資料
         Erase SL
         '抓前前次薪資異動內容
         'Modify by Morgan 2009/5/27 不抓異動類別為非調薪的(任用的例外TN)
         'modify by sonia 2017/4/7 +a.sl02最近調薪日
         strExc(0) = "select b.*,a.sl02 slnew from salarylog a,salarylog b" & _
            " where a.sl01='" & .Fields("st01") & "'" & _
            " and a.sl02=(select max(c.sl02) from salarylog c where c.sl01=a.sl01 and (c.sl35 is null or c.sl03||c.sl35='TN') )" & _
            " and b.sl01(+)=a.sl01" & _
            " and b.sl02=(select max(c.sl02) from salarylog c where c.sl01=a.sl01 and c.sl02<a.sl02 and (c.sl35 is null or c.sl03||c.sl35='TN') )"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            With RsTemp
               For intI = 11 To 26
                  SL(intI) = Val("" & .Fields("sl" & Format(intI, "00")))
               Next
               SL(39) = Val("" & .Fields("sl" & Format(39, "00"))) 'Add By Sindy 2020/6/23 證照津貼
            End With
            m_sl02_1 = Val(RsTemp.Fields("sl02"))     'add by sonia 2017/4/7
            m_sl02_New = Val(RsTemp.Fields("slnew"))   'add by sonia 2017/4/7
            m_sl03_1 = "" & RsTemp.Fields("sl03") 'Added by Morgan 2023/5/11
         End If
         
      If Text1(3) = "1" Then 'Added by Morgan 2021/4/28
      
         '抓前4年的考績
         'Modify by Morgan 2009/6/22 考慮復職員工改甲等也存檔
         iYear1 = iYear2 - 4
         If iYear1 < Val(Left("" & .Fields("st13"), 4)) Then
            iYear1 = Val(Left("" & .Fields("st13"), 4))
         End If
         stAssessList = ""
         strExc(0) = "select ym01,ym02 from YearMerit where YM01>=" & iYear1 & _
            " and ym03='" & .Fields("st01") & "' order by ym01 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            iYear3 = iYear2
            Do While Not RsTemp.EOF
               For intI = iYear3 - 1 To RsTemp.Fields("ym01") + 1 Step -1
                  stAssessList = stAssessList & (intI - 1911) & " 年：甲等   "
               Next
               iYear3 = RsTemp.Fields("ym01")
               Select Case RsTemp("ym02")
                  Case "1"
                     stAssessList = stAssessList & (intI - 1911) & " 年：優等   "
                  Case "2" 'Added by Morgan 2013/4/17 修正甲等沒有bug
                     stAssessList = stAssessList & (intI - 1911) & " 年：甲等   "
                  Case "3"
                     stAssessList = stAssessList & (intI - 1911) & " 年：乙等   "
                  Case "4"
                     stAssessList = stAssessList & (intI - 1911) & " 年：丁等   "
                  'add by sonia 2016/1/6
                  Case "*"
                     stAssessList = stAssessList & (intI - 1911) & " 年：不得參加  "
               End Select
               RsTemp.MoveNext
            Loop
            For intI = iYear3 - 1 To iYear1 Step -1
               stAssessList = stAssessList & (intI - 1911) & " 年：甲等   "
            Next
         Else
            For intI = iYear2 - 1 To iYear1 Step -1
               stAssessList = stAssessList & (intI - 1911) & " 年：甲等   "
            Next
         End If
         
         '原職稱,職位(抓職位&職稱不同的最大一筆)
         strExc(0) = "select c1.ac03 Tit,c2.ac03 Pos" & _
            " from staff_change a,allcode c1,allcode c2" & _
            " where sc01='" & .Fields("st01") & "' and sc02" & _
            "=(select max(b.sc02) from staff_change b where b.sc01=a.sc01 and b.sc05||b.sc06<>'" & .Fields("st20") & .Fields("st21") & "')" & _
            " and c1.ac02(+)=sc05 and c1.ac01(+)='01'" & _
            " and c2.ac02(+)=sc06 and c2.ac01(+)='02'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            PosOld = "" & RsTemp.Fields("Pos")
            TitOld = "" & RsTemp.Fields("Tit")
         Else
            PosOld = ""
            TitOld = ""
         End If

         'add by sonia 2017/4/7 再抓前一次
         '原薪資資料
         Erase SO
         If Left(strSrvDate(1), 4) = 2017 And m_sl02_1 > 0 Then
            '抓前前次薪資異動內容
            strExc(0) = "select b.* from salarylog a,salarylog b" & _
               " where a.sl01='" & .Fields("st01") & "'" & _
               " and a.sl02=(select max(c.sl02) from salarylog c where c.sl01=a.sl01 and (c.sl35 is null or c.sl03||c.sl35='TN') )" & _
               " and b.sl01(+)=a.sl01" & _
               " and b.sl02=(select max(c.sl02) from salarylog c where c.sl01=a.sl01 and c.sl02<" & m_sl02_1 & " and (c.sl35 is null or c.sl03||c.sl35='TN') )"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               With RsTemp
                  For intI = 11 To 26
                     SO(intI) = Val("" & .Fields("sl" & Format(intI, "00")))
                  Next
               End With
               m_sl02_2 = Val(RsTemp.Fields("sl02"))
            End If
         End If
         'end 2017/4/7
   
         '2013/4/9 modify by sonia 因為剔除年資滿二十五年不印,所以不能要此變數
         'If .AbsolutePosition > 1 Then Printer.NewPage
         If iPage > 1 Then Printer.NewPage
         
         Printer.FontSize = 20
         Printer.FontBold = True
         yi = Yo + 750
         
         strExc(1) = "" & .Fields("comp1A")
         xi = Px(0) + (Px(16) - Px(0)) / 2 - Printer.TextWidth(strExc(1)) / 2
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         Printer.FontSize = 16 '14
         strExc(1) = iYear2 - 1911 & " 年度薪資調薪表"
         yi = yi + 2 * iLineHeight
         xi = Px(0) + (Px(16) - Px(0)) / 2 - Printer.TextWidth(strExc(1)) / 2
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         Printer.FontSize = 12
         Printer.FontBold = False
         
         yi = yi + 2 * iLineHeight
         xi = Px(0)
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print "部　　門：" & .Fields("dept")
                  
         xi = Px(14)
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
         
         yi = yi + iLineHeight
         xi = Px(0)
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print "員工編號：" & .Fields("st01")
         
         xi = Px(14)
         Printer.CurrentX = xi: Printer.CurrentY = yi
         
         'Printer.Print "頁　　次：" & Format(.AbsolutePosition, "#000")
         Printer.Print "頁　　次：" & Format(iPage, "#000")
         
         yi = yi + iLineHeight
         xi = Px(0)
         Printer.CurrentX = xi: Printer.CurrentY = yi
         'Modify By Sindy 2022/4/18
         Printer.Print "姓　　名：" '& .Fields("st02")
         xi = 1500
         Printer.CurrentX = xi: Printer.CurrentY = yi
         HPrint .Fields("st02")
         '2022/4/18 END
         
         'Modify By Sindy 2020/7/2 拿掉不顯示
'         xi = Px(6) - 400
'         Printer.CurrentX = xi: Printer.CurrentY = yi
'         Printer.Print "性　　別：" & IIf(.Fields("st22") = "F", "女", "男")
'
'         xi = Px(9) + 300
'         Printer.CurrentX = xi: Printer.CurrentY = yi
'         Printer.Print "出生日期：" & ChangeWStringToTDateString("" & .Fields("st23"))
         
         xi = Px(14)
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print "入所日期：" & ChangeWStringToTDateString("" & .Fields("st13"))
         
         yi = yi + 2 * iLineHeight
         xi = Px(0)
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print "歷年考績：" & stAssessList
         
         '選新制加印勞退投保薪資
         If .Fields("sd16") = "Y" Then
            yi = yi + iLineHeight
            xi = Px(0)
            If Val("" & .Fields("sd27")) > 0 Then
               strExc(2) = Val("" & .Fields("sd27"))
            Else
               'Modified by Morgan 2020/10/27 改抓sd43(含證照)
               'strExc(2) = Val("" & .Fields("sd20")) + Val("" & .Fields("sd21")) + Val("" & .Fields("sd23"))
               strExc(2) = Val("" & .Fields("sd43"))
            End If
            Printer.CurrentX = xi: Printer.CurrentY = yi
            strExc(1) = "目前退休金投保薪資　" & .Fields("comp1") & "：" & Format(Val(strExc(2)), "#,###")
            Printer.Print strExc(1)
            If IsNull(.Fields("comp2")) = False Then
               If Val("" & .Fields("sd36")) > 0 Then
                  strExc(2) = Val("" & .Fields("sd36"))
               Else
                  strExc(2) = Val("" & .Fields("sd29")) + Val("" & .Fields("sd30")) + Val("" & .Fields("sd32"))
               End If
            
               xi = Px(10)
               Printer.CurrentX = xi: Printer.CurrentY = yi
               strExc(1) = .Fields("comp2") & "：" & Format(Val(strExc(2)), "#,###")
               Printer.Print strExc(1)
            End If
         End If
         
         yi = yi + iLineHeight
         
         Erase Py
         Py(0) = yi
         SetPy
         
         '畫表格
         PrintTable
         
         yi = Py(0) + iMarginY
         strExc(1) = "職　　　　　　　　位"
         xi = (Px(0) + Px(9)) / 2 - (Printer.TextWidth(strExc(1))) / 2
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
                  
         strExc(1) = "職　　　　　　　　稱"
         xi = (Px(9) + Px(16)) / 2 - (Printer.TextWidth(strExc(1))) / 2
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         yi = Py(1) + iMarginY
         strExc(1) = "上　　次"
         xi = (Px(0) + Px(4)) / 2 - (Printer.TextWidth(strExc(1))) / 2
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         strExc(1) = "目　　前"
         xi = (Px(4) + Px(6)) / 2 - (Printer.TextWidth(strExc(1))) / 2
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)

         strExc(1) = "今　　年"
         xi = (Px(6) + Px(9)) / 2 - (Printer.TextWidth(strExc(1))) / 2
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         strExc(1) = "上　　次"
         xi = (Px(9) + Px(11)) / 2 - (Printer.TextWidth(strExc(1))) / 2
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         strExc(1) = "目　　前"
         xi = (Px(11) + Px(13)) / 2 - (Printer.TextWidth(strExc(1))) / 2
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)

         strExc(1) = "今　　年"
         xi = (Px(13) + Px(16)) / 2 - (Printer.TextWidth(strExc(1))) / 2
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         yi = Py(2) + iMarginY
         
         strExc(1) = PosOld
         xi = (Px(0) + Px(4)) / 2 - (Printer.TextWidth(strExc(1))) / 2
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         strExc(1) = "" & .Fields("Pos")
         xi = (Px(4) + Px(6)) / 2 - (Printer.TextWidth(strExc(1))) / 2
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         xi = Px(6) + iMargin
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print "☆"
                 
         strExc(1) = TitOld
         xi = (Px(9) + Px(11)) / 2 - (Printer.TextWidth(strExc(1))) / 2
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         strExc(1) = "" & .Fields("tit")
         xi = (Px(11) + Px(13)) / 2 - (Printer.TextWidth(strExc(1))) / 2
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
                  
         xi = Px(13) + iMargin
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print "☆"
         
'         'add by sonia 2017/4/7 因2017/1全體員工調整午餐津貼所以要多印再前一次的調薪
'         If Left(strSrvDate(1), 4) = 2017 Then
'            yi = Py(4) + iMarginY
'            strExc(1) = "薪　　　資　　　項　　　別"
'            xi = (Px(0) + Px(7)) / 2 - (Printer.TextWidth(strExc(1))) / 2
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(1)
'
'            strExc(1) = "今　年　目　標　、　薪　資　、　津　貼"
'            xi = (Px(7) + Px(16)) / 2 - (Printer.TextWidth(strExc(1))) / 2
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(1)
'
'            yi = Py(5) + iMarginY
'            strExc(1) = "項　　目"
'            Printer.CurrentX = 600: Printer.CurrentY = yi
'            Printer.Print strExc(1)
'
'            strExc(1) = Left(ChangeTStringToTDateString(ChangeWStringToTString(m_sl02_2)), 6)   '上次
'            Printer.CurrentX = 1850: Printer.CurrentY = yi
'            Printer.Print strExc(1)
'
'            strExc(1) = Left(ChangeTStringToTDateString(ChangeWStringToTString(m_sl02_1)), 6)  '加伙食
'            Printer.CurrentX = 2900: Printer.CurrentY = yi
'            Printer.Print strExc(1)
'
'            If m_sl02_New = 0 Then
'               strExc(1) = "目　前"
'            Else
'               strExc(1) = Left(ChangeTStringToTDateString(ChangeWStringToTString(m_sl02_New)), 6)  '目　前
'            End If
'            Printer.CurrentX = 3900: Printer.CurrentY = yi
'            Printer.Print strExc(1)
'
'            strExc(1) = "主　任"
'            Printer.CurrentX = 4900: Printer.CurrentY = yi
'            Printer.Print strExc(1)
'
'            strExc(1) = "副　理"
'            Printer.CurrentX = 6000: Printer.CurrentY = yi
'            Printer.Print strExc(1)
'
'            strExc(1) = "經　理"
'            Printer.CurrentX = 7000: Printer.CurrentY = yi
'            Printer.Print strExc(1)
'
'            strExc(1) = "協　理"
'            Printer.CurrentX = 8100: Printer.CurrentY = yi
'            Printer.Print strExc(1)
'
'            strExc(1) = "副　總"
'            Printer.CurrentX = 9150: Printer.CurrentY = yi
'            Printer.Print strExc(1)
'
'            strExc(1) = "總經理"
'            Printer.CurrentX = 10259: Printer.CurrentY = yi
'            Printer.Print strExc(1)
'
'            yi = Py(6) + iMarginY
'            xi = Px(0) + iMargin
'            Printer.CurrentX = 400: Printer.CurrentY = yi
'            VPrint "" & .Fields("comp1")
'
'            strExc(1) = "本　薪"
'            Printer.CurrentX = 900: Printer.CurrentY = yi
'            Printer.Print strExc(1)
'            '上次
'            strExc(2) = Format(SO(11) + SO(14), "#,###")
'            xi = 2600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '調伙食
'            strExc(2) = Format(SL(11) + SL(14), "#,###")
'            xi = 3600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD20")) + Val("" & .Fields("SD23")), "#,###")
'            xi = 4600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            yi = Py(7) + iMarginY
'            Printer.CurrentX = 1200: Printer.CurrentY = yi
'            Printer.Print "職務"
'            '上次
'            strExc(2) = Format(SO(12), "#,###")
'            xi = 2600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '調伙食
'            strExc(2) = Format(SL(12), "#,###")
'            xi = 3600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD21")), "#,###")
'            xi = 4600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            yi = Py(8) + iMarginY
'            Printer.CurrentX = 800: Printer.CurrentY = yi
'            VPrint "津　　　貼"
'
'            Printer.CurrentX = 1200: Printer.CurrentY = yi
'            Printer.Print "技術"
'            '上次
'            strExc(2) = Format(SO(13), "#,###")
'            xi = 2600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '調伙食
'            strExc(2) = Format(SL(13), "#,###")
'            xi = 3600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD22")), "#,###")
'            xi = 4600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            yi = Py(9) + iMarginY
'            Printer.CurrentX = 1200: Printer.CurrentY = yi
'            Printer.Print "差旅"
'            '上次
'            strExc(2) = Format(SO(15), "#,###")
'            If Left("" & .Fields("st03"), 1) = "S" And "" & .Fields("st03") > "S10" Then
'               If Val(SO(15)) = 0 Then strExc(2) = "實報實銷"
'            End If
'            xi = 2700 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '調伙食
'            strExc(2) = Format(SL(15), "#,###")
'            If Left("" & .Fields("st03"), 1) = "S" And "" & .Fields("st03") > "S10" Then
'               If Val(SL(15)) = 0 Then strExc(2) = "實報實銷"
'            End If
'            xi = 3700 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD24")), "#,###")
'            If Left("" & .Fields("st03"), 1) = "S" And "" & .Fields("st03") > "S10" Then
'               If Val("" & .Fields("SD24")) = 0 Then strExc(2) = "實報實銷"
'            End If
'            xi = 4700 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            yi = Py(10) + iMarginY
'            Printer.CurrentX = 1200: Printer.CurrentY = yi
'            Printer.Print "房租"
'            '上次
'            strExc(2) = Format(SO(16), "#,###")
'            xi = 2600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '調伙食
'            strExc(2) = Format(SL(16), "#,###")
'            xi = 3600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD25")), "#,###")
'            xi = 4600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            yi = Py(11) + iMarginY
'            Printer.CurrentX = 1200: Printer.CurrentY = yi
'            Printer.Print "特支"
'            '上次
'            strExc(2) = Format(SO(17), "#,###")
'            xi = 2600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '調伙食
'            strExc(2) = Format(SL(17), "#,###")
'            xi = 3600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD26")), "#,###")
'            xi = 4600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'
'            yi = Py(12) + iMarginY
'            Printer.CurrentX = 400: Printer.CurrentY = yi
'            VPrint "" & .Fields("comp2")
'
'            strExc(1) = "本　薪"
'            Printer.CurrentX = 900: Printer.CurrentY = yi
'            Printer.Print strExc(1)
'            '上次
'            strExc(2) = Format(SO(19) + SO(22), "#,###")
'            xi = 2600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '調伙食
'            strExc(2) = Format(SL(19) + SL(22), "#,###")
'            xi = 3600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD29")) + Val("" & .Fields("SD32")), "#,###")
'            xi = 4600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            yi = Py(13) + iMarginY
'            Printer.CurrentX = 1200: Printer.CurrentY = yi
'            Printer.Print "職務"
'            '上次
'            strExc(2) = Format(SO(20), "#,###")
'            xi = 2600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '調伙食
'            strExc(2) = Format(SL(20), "#,###")
'            xi = 3600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD30")), "#,###")
'            xi = 4600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            yi = Py(14) + iMarginY
'            Printer.CurrentX = 800: Printer.CurrentY = yi
'            VPrint "津　　　貼"
'
'            Printer.CurrentX = 1200: Printer.CurrentY = yi
'            Printer.Print "技術"
'            '上次
'            strExc(2) = Format(SO(21), "#,###")
'            xi = 2600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '調伙食
'            strExc(2) = Format(SL(21), "#,###")
'            xi = 3600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD31")), "#,###")
'            xi = 4600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            yi = Py(15) + iMarginY
'            Printer.CurrentX = 1200: Printer.CurrentY = yi
'            Printer.Print "差旅"
'            '上次
'            strExc(2) = Format(SO(23), "#,###")
'            xi = 2600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '調伙食
'            strExc(2) = Format(SL(23), "#,###")
'            xi = 3600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD33")), "#,###")
'            xi = 4600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            yi = Py(16) + iMarginY
'            Printer.CurrentX = 1200: Printer.CurrentY = yi
'            Printer.Print "房租"
'            '上次
'            strExc(2) = Format(SO(24), "#,###")
'            xi = 2600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '調伙食
'            strExc(2) = Format(SL(24), "#,###")
'            xi = 3600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD34")), "#,###")
'            xi = 4600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            yi = Py(17) + iMarginY
'            Printer.CurrentX = 1200: Printer.CurrentY = yi
'            Printer.Print "特支"
'            '上次
'            strExc(2) = Format(SO(25), "#,###")
'            xi = 2600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '調伙食
'            strExc(2) = Format(SL(25), "#,###")
'            xi = 3600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD35")), "#,###")
'            xi = 4600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'
'            yi = Py(18) + iMarginY
'            strExc(1) = "本　薪"
'            Printer.CurrentX = 900: Printer.CurrentY = yi
'            Printer.Print strExc(1)
'            '上次
'            strExc(2) = Format(SO(11) + SO(14) + SO(19) + SO(22), "#,###")
'            xi = 2600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '調伙食
'            strExc(2) = Format(SL(11) + SL(14) + SL(19) + SL(22), "#,###")
'            xi = 3600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD20")) + Val("" & .Fields("SD23")) + Val("" & .Fields("SD29")) + Val("" & .Fields("SD32")), "#,###")
'            xi = 4600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            yi = Py(19) + iMarginY
'            Printer.CurrentX = 400: Printer.CurrentY = yi
'            VPrint "合　　　　計"
'
'            Printer.CurrentX = 1200: Printer.CurrentY = yi
'            Printer.Print "職務"
'            '上次
'            strExc(2) = Format(SO(12) + SO(20), "#,###")
'            xi = 2600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '調伙食
'            strExc(2) = Format(SL(12) + SL(20), "#,###")
'            xi = 3600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD21")) + Val("" & .Fields("SD30")), "#,###")
'            xi = 4600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            yi = Py(20) + iMarginY
'            Printer.CurrentX = 800: Printer.CurrentY = yi
'            VPrint "津　　　貼"
'
'            Printer.CurrentX = 1200: Printer.CurrentY = yi
'            Printer.Print "技術"
'            '上次
'            strExc(2) = Format(SO(13) + SO(21), "#,###")
'            xi = 2600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '調伙食
'            strExc(2) = Format(SL(13) + SL(21), "#,###")
'            xi = 3600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD22")) + Val("" & .Fields("SD31")), "#,###")
'            xi = 4600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            yi = Py(21) + iMarginY
'            Printer.CurrentX = 1200: Printer.CurrentY = yi
'            Printer.Print "差旅"
'            '上次
'            strExc(2) = Format(SO(15) + SO(23), "#,###")
'            xi = 2600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '調伙食
'            strExc(2) = Format(SL(15) + SL(23), "#,###")
'            xi = 3600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD24")) + Val("" & .Fields("SD33")), "#,###")
'            xi = 4600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            yi = Py(22) + iMarginY
'            Printer.CurrentX = 1200: Printer.CurrentY = yi
'            Printer.Print "房租"
'            '上次
'            strExc(2) = Format(SO(16) + SO(24), "#,###")
'            xi = 2600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '調伙食
'            strExc(2) = Format(SL(16) + SL(24), "#,###")
'            xi = 3600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD25")) + Val("" & .Fields("SD34")), "#,###")
'            xi = 4600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            yi = Py(23) + iMarginY
'            Printer.CurrentX = 1200: Printer.CurrentY = yi
'            Printer.Print "特支"
'            '上次
'            strExc(2) = Format(SO(17) + SO(25), "#,###")
'            xi = 2600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '調伙食
'            strExc(2) = Format(SL(17) + SL(25), "#,###")
'            xi = 3600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD26")) + Val("" & .Fields("SD35")), "#,###")
'            xi = 4600 - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'         Else
'         'end 2017/4/7
            yi = Py(4) + iMarginY
            strExc(1) = "薪　　資　　項　　別"
            xi = (Px(0) + Px(7)) / 2 - (Printer.TextWidth(strExc(1))) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
   
            strExc(1) = "今　年　目　標　、　薪　資　、　津　貼"
            xi = (Px(7) + Px(16)) / 2 - (Printer.TextWidth(strExc(1))) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            yi = Py(5) + iMarginY
            strExc(1) = "項　　目"
            xi = (Px(0) + Px(3)) / 2 - (Printer.TextWidth(strExc(1))) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
   
            strExc(1) = "上　次"
            xi = (Px(3) + Px(5)) / 2 - (Printer.TextWidth(strExc(1))) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            strExc(1) = "目　前"
            xi = (Px(5) + Px(7)) / 2 - (Printer.TextWidth(strExc(1))) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            strExc(1) = "主　任"
            xi = (Px(7) + Px(8)) / 2 - (Printer.TextWidth(strExc(1))) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            strExc(1) = "副　理"
            xi = (Px(8) + Px(10)) / 2 - (Printer.TextWidth(strExc(1))) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            strExc(1) = "經　理"
            xi = (Px(10) + Px(12)) / 2 - (Printer.TextWidth(strExc(1))) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            strExc(1) = "協　理"
            xi = (Px(12) + Px(14)) / 2 - (Printer.TextWidth(strExc(1))) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            strExc(1) = "副　總"
            xi = (Px(14) + Px(15)) / 2 - (Printer.TextWidth(strExc(1))) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            strExc(1) = "總經理"
            xi = (Px(15) + Px(16)) / 2 - (Printer.TextWidth(strExc(1))) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            '公司別
            yi = Py(6) + iMarginY
'            xi = Px(0) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            VPrint "" & .Fields("comp1")
            
            strExc(1) = "本　　薪"
            xi = (Px(1) + Px(3)) / 2 - (Printer.TextWidth(strExc(1))) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            '上次
            strExc(2) = Format(SL(11) + SL(14), "#,###")
            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            '目前
            strExc(2) = Format(Val("" & .Fields("SD20")) + Val("" & .Fields("SD23")), "#,###")
            xi = Px(7) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
            yi = Py(7) + iMarginY
            xi = Px(1) + iMargin + 50
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print "職　務"
            '上次
            strExc(2) = Format(SL(12), "#,###")
            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            '目前
            strExc(2) = Format(Val("" & .Fields("SD21")), "#,###")
            xi = Px(7) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
            yi = Py(8) + iMarginY
            xi = Px(0) + iMargin - 50
            Printer.CurrentX = xi: Printer.CurrentY = yi
            VPrint "津　　　　貼"
            
            xi = Px(1) + iMargin + 50
            'modify by sonia 2020/5/6 技術改技術/證照,分二行,字縮小
            Printer.CurrentX = xi: Printer.CurrentY = yi '- 100
            'Printer.FontSize = 10
            'Printer.Print "技術/"
            Printer.Print "證　照"
            'Printer.FontSize = 12
            '上次
            strExc(2) = Format(SL(39), "#,###")
            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            '目前
            strExc(2) = Format(Val("" & .Fields("SD52")), "#,###")
            xi = Px(7) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)

''add by sonia 2020/5/7 技術改技術/證照,分二行
'            xi = Px(2) + iMargin - 50
'            Printer.CurrentX = xi: Printer.CurrentY = yi + 100
'            Printer.FontSize = 10
'            Printer.Print "證照"
'            Printer.FontSize = 12
''end 2020/5/7
            
            yi = Py(9) + iMarginY
            xi = Px(1) + iMargin + 50
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print "差　旅"
            '上次
            strExc(2) = Format(SL(15), "#,###")
            '2013/10/9 ADD BY SONIA 智權部同仁若差旅費無金額差旅費欄註明 "實報實銷"
            'modify by sonia 2024/4/24 應包含S10
            If Left("" & .Fields("st03"), 1) = "S" And "" & .Fields("st03") >= "S10" Then
               If Val(SL(15)) = 0 Then strExc(2) = "實報實銷"
            End If
            '2013/10/9 END
            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            '目前
            strExc(2) = Format(Val("" & .Fields("SD24")), "#,###")
            '2013/10/9 ADD BY SONIA 智權部同仁若差旅費無金額差旅費欄註明 "實報實銷"
            'modify by sonia 2024/4/24 應包含S10
            If Left("" & .Fields("st03"), 1) = "S" And "" & .Fields("st03") >= "S10" Then
               If Val("" & .Fields("SD24")) = 0 Then strExc(2) = "實報實銷"
            End If
            '2013/10/9 END
            xi = Px(7) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
            yi = Py(10) + iMarginY
            xi = Px(1) + iMargin + 50
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print "房　租"
            '上次
            strExc(2) = Format(SL(16), "#,###")
            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            '目前
            strExc(2) = Format(Val("" & .Fields("SD25")), "#,###")
            xi = Px(7) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
            'Add By Sindy 2020/6/23
            yi = Py(11) + iMarginY
            xi = Px(1) + iMargin + 50
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print "技　術"
            '上次
            strExc(2) = Format(SL(13), "#,###")
            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            '目前
            strExc(2) = Format(Val("" & .Fields("SD22")), "#,###")
            xi = Px(7) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            '2020/6/23 END
            
            yi = Py(12) + iMarginY
            xi = Px(1) + iMargin + 50
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print "特　支"
            '上次
            strExc(2) = Format(SL(17), "#,###")
            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            '目前
            strExc(2) = Format(Val("" & .Fields("SD26")), "#,###")
            xi = Px(7) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
            'Add By Sindy 2020/7/1 + 合計
            yi = Py(13) + iMarginY
            strExc(1) = "合　　計"
            xi = (Px(1) + Px(3)) / 2 - (Printer.TextWidth(strExc(1))) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            '上次
            strExc(2) = Format(SL(11) + SL(12) + SL(13) + SL(39) + SL(14) + SL(15) + SL(16) + SL(17), "#,###")
            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            '目前
            strExc(2) = Format(Val("" & .Fields("SD20")) + Val("" & .Fields("SD23")) + Val("" & .Fields("SD21")) + Val("" & .Fields("SD52")) + Val("" & .Fields("SD24")) + Val("" & .Fields("SD25")) + Val("" & .Fields("SD22")) + Val("" & .Fields("SD26")), "#,###")
            xi = Px(7) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            '2020/7/1 END
            
'            yi = Py(12) + iMarginY
'            xi = Px(0) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            VPrint "" & .Fields("comp2")
'
'            strExc(1) = "本　　薪"
'            xi = (Px(1) + Px(3)) / 2 - (Printer.TextWidth(strExc(1))) / 2
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(1)
'            '上次
'            strExc(2) = Format(SL(19) + SL(22), "#,###")
'            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD29")) + Val("" & .Fields("SD32")), "#,###")
'            xi = Px(7) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            yi = Py(13) + iMarginY
'            xi = Px(2) + iMargin - 50
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print "職務"
'            '上次
'            strExc(2) = Format(SL(20), "#,###")
'            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD30")), "#,###")
'            xi = Px(7) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            yi = Py(14) + iMarginY
'            xi = Px(1) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            VPrint "津　　　貼"
'
'            xi = Px(2) + iMargin - 50
'            'modify by sonia 2020/5/6 技術改技術/證照,分二行,字縮小
'            Printer.CurrentX = xi: Printer.CurrentY = yi - 100
'            Printer.FontSize = 10
'            Printer.Print "技術/"
'            Printer.FontSize = 12
'            '上次
'            strExc(2) = Format(SL(21), "#,###")
'            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD31")), "#,###")
'            xi = Px(7) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
''add by sonia 2020/5/7 技術改技術/證照,分二行
'            xi = Px(2) + iMargin - 50
'            Printer.CurrentX = xi: Printer.CurrentY = yi + 100
'            Printer.FontSize = 10
'            Printer.Print "證照"
'            Printer.FontSize = 12
''end 2020/5/7
'
'            yi = Py(15) + iMarginY
'            xi = Px(2) + iMargin - 50
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print "差旅"
'            '上次
'            strExc(2) = Format(SL(23), "#,###")
'            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD33")), "#,###")
'            xi = Px(7) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            yi = Py(16) + iMarginY
'            xi = Px(2) + iMargin - 50
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print "房租"
'            '上次
'            strExc(2) = Format(SL(24), "#,###")
'            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD34")), "#,###")
'            xi = Px(7) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            yi = Py(17) + iMarginY
'            xi = Px(2) + iMargin - 50
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print "特支"
'            '上次
'            strExc(2) = Format(SL(25), "#,###")
'            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD35")), "#,###")
'            xi = Px(7) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
            
            
'            yi = Py(18) + iMarginY
'            strExc(1) = "本　　薪"
'            xi = (Px(1) + Px(3)) / 2 - (Printer.TextWidth(strExc(1))) / 2
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(1)
'            '上次
'            strExc(2) = Format(SL(11) + SL(14) + SL(19) + SL(22), "#,###")
'            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD20")) + Val("" & .Fields("SD23")) + Val("" & .Fields("SD29")) + Val("" & .Fields("SD32")), "#,###")
'            xi = Px(7) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            yi = Py(19) + iMarginY
'            xi = Px(0) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            VPrint "合　　　　計"
'
'            xi = Px(2) + iMargin - 50
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print "職務"
'            '上次
'            strExc(2) = Format(SL(12) + SL(20), "#,###")
'            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD21")) + Val("" & .Fields("SD30")), "#,###")
'            xi = Px(7) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            yi = Py(20) + iMarginY
'            xi = Px(1) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            VPrint "津　　　貼"
'
'            xi = Px(2) + iMargin - 50
'            'modify by sonia 2020/5/6 技術改技術/證照,分二行,字縮小
'            Printer.CurrentX = xi: Printer.CurrentY = yi - 100
'            Printer.FontSize = 10
'            Printer.Print "技術/"
'            Printer.FontSize = 12
'            '上次
'            strExc(2) = Format(SL(13) + SL(21), "#,###")
'            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD22")) + Val("" & .Fields("SD31")), "#,###")
'            xi = Px(7) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
''add by sonia 2020/5/7 技術改技術/證照,分二行
'            xi = Px(2) + iMargin - 50
'            Printer.CurrentX = xi: Printer.CurrentY = yi + 100
'            Printer.FontSize = 10
'            Printer.Print "證照"
'            Printer.FontSize = 12
''end 2020/5/7
'
'            yi = Py(21) + iMarginY
'            xi = Px(2) + iMargin - 50
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print "差旅"
'            '上次
'            strExc(2) = Format(SL(15) + SL(23), "#,###")
'            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD24")) + Val("" & .Fields("SD33")), "#,###")
'            xi = Px(7) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            yi = Py(22) + iMarginY
'            xi = Px(2) + iMargin - 50
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print "房租"
'            '上次
'            strExc(2) = Format(SL(16) + SL(24), "#,###")
'            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD25")) + Val("" & .Fields("SD34")), "#,###")
'            xi = Px(7) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            yi = Py(23) + iMarginY
'            xi = Px(2) + iMargin - 50
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print "特支"
'            '上次
'            strExc(2) = Format(SL(17) + SL(25), "#,###")
'            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'            '目前
'            strExc(2) = Format(Val("" & .Fields("SD26")) + Val("" & .Fields("SD35")), "#,###")
'            xi = Px(7) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'         End If
         
         yi = Py(16) + iLineHeight
         xi = Px(0) + iMargin
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print "註一：請依規定日期前呈閱( " & Text1(1) & " )"
         
'cancel by sonia 2020/3/27 二事務所合併故取消
'         yi = yi + iLineHeight
'         xi = Px(0) + iMargin
'         Printer.CurrentX = xi: Printer.CurrentY = yi
'         Printer.Print "註二：台一事務所同仁同時處理台一國際專利法律事務所（簡稱專利）及台一專利商標事務所（簡稱"
'
'         yi = yi + iLineHeight
'         xi = Px(0) + iMargin
'         Printer.CurrentX = xi: Printer.CurrentY = yi
'         Printer.Print "　　　商標）二家事務所業務，應按同仁處理工作之比例分別予以核薪，請分別填入薪資"
'end 2020/3/27

         yi = yi + iLineHeight
         xi = Px(0) + iMargin
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print "註二：本年度有職位或職稱晉升建議者，請在☆處填寫"
         
         'add by sonia 2017/4/7 因2017/1全體員工調整午餐津貼所以要多印一行
         If Left(strSrvDate(1), 4) = 2017 Then
            yi = yi + iLineHeight
            xi = Px(0) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print "註三：106年1月全體同仁伙食津貼由1,800元調為2,400元"
         End If
         'end 2017/4/7
         
      End If 'Added by Morgan 2021/4/28
         
         'Added by Morgan 2020/10/15
         If bolSaveXLS Then
            ii = ii + 1
            
            If ii > 2 And Left("" & .Fields("st03"), 2) <> Left(stLsdDeptNo, 2) Then
               '小部門合計
               strExc(1) = GetSubDeptName(stLsdDeptNo, stLsdDeptName)
               If strExc(1) <> "" Then
                  If ii > iRow1 + 1 Then
                     strExc(1) = strExc(1) & "合計"
                     SetSubTotRow wksReport, strExc(1), ii, iRow1
                     ii = ii + 1
                  Else
                     strExc(1) = ""
                  End If
                  ii = ii + 1
                  iRow1 = ii
               End If
               
               '大部門合計
               If Left("" & .Fields("st03"), 1) <> Left(stLsdDeptNo, 1) Then
                  strExc(2) = GetSubDeptName(stLsdDeptNo, stLsdDeptName, 1)
                  If strExc(2) <> "" Then
                     SetSubTotRow wksReport, strExc(2), ii, iRow2
                     ii = ii + 2
                  End If
                  iRow2 = ii
                  iRow1 = iRow2
               End If
            End If
            
            stLsdDeptNo = "" & adoRst.Fields("st03")
            stLsdDeptName = "" & adoRst.Fields("dept")
            
            wksReport.Range("A" & ii) = "" & .Fields("dept") '部門
            wksReport.Range("B" & ii) = "" & .Fields("sd16") '勞退新制
            wksReport.Range("C" & ii).NumberFormatLocal = "@"
            wksReport.Range("C" & ii) = "" & .Fields("st01") '編號
            wksReport.Range("D" & ii) = "" & .Fields("st02") '姓名
            'Added by Morgan 2025/3/18
            wksReport.Range("E" & ii) = "" & .Fields("deptnew") '新部門
            wksReport.Range("F" & ii) = "" & .Fields("tit") '職稱
            'end 2025/3/18
            'Added by Morgan 2021/4/28
            If Text1(3) = "3" Then
               wksReport.Range("G" & ii) = Val(SL(11)) + Val(SL(14)) '本薪
               wksReport.Range("G" & ii).NumberFormatLocal = "#,##0"
               wksReport.Range("H" & ii) = Val(SL(12)) + Val(SL(13)) + Val(SL(15)) + Val(SL(16)) + Val(SL(17)) + Val(SL(39)) '津貼
               wksReport.Range("H" & ii).NumberFormatLocal = "#,##0"
               wksReport.Range("I" & ii) = (Val("" & .Fields("SD20")) + Val("" & .Fields("SD23"))) - Val(wksReport.Range("G" & ii)) '本薪調升金額
               wksReport.Range("I" & ii).NumberFormatLocal = "#,##0"
               wksReport.Range("J" & ii).Formula = "=I" & ii & "/G" & ii '本薪調升率
               wksReport.Range("J" & ii).NumberFormatLocal = "0.00%"
               wksReport.Range("K" & ii) = Val("" & .Fields("SD21")) + Val("" & .Fields("SD22")) + Val("" & .Fields("SD24")) + Val("" & .Fields("SD25")) + Val("" & .Fields("SD26")) + Val("" & .Fields("SD52")) - Val(wksReport.Range("H" & ii)) '津貼調升金額"
               wksReport.Range("K" & ii).NumberFormatLocal = "#,##0"
               wksReport.Range("L" & ii).Formula = "=K" & ii & "+I" & ii '含津貼調升金額
               wksReport.Range("L" & ii).NumberFormatLocal = "#,##0"
               wksReport.Range("M" & ii).Formula = "=L" & ii & "/G" & ii  '含津貼調升率
               wksReport.Range("M" & ii).NumberFormatLocal = "0.00%"
               wksReport.Range("N" & ii).Formula = "=G" & ii & "+H" & ii & "+L" & ii '調升後
               wksReport.Range("N" & ii).NumberFormatLocal = "#,##0"
               'Added by Morgan 2024/5/16
               wksReport.Range("O" & ii) = Val(SL(39)) '證照津貼
               wksReport.Range("O" & ii).NumberFormatLocal = "#,##0"
               'end 2024/5/16
            Else
            'end 2021/4/28
               If Text1(0) = "1" Then
                  strExc(1) = Left(strSrvDate(1) - 10000, 4) & "0501"
               Else
                  strExc(1) = Left(strSrvDate(1) - 10000, 4) & "1101"
               End If
               '額外調薪:最後調薪為適用期滿或非例行調薪
               If m_sl03_1 = "T" Or m_sl02_New > strExc(1) Then
                  wksReport.Range("G" & ii) = "V" '額外調薪
                  wksReport.Range("G" & ii).Select
                  strExc(2) = "額外調薪明細!B" & m_rpt2Row
                  wksReport.Range("G" & ii).Select
                  wksReport.Hyperlinks.add Anchor:=xlsReport.Selection, address:="", SubAddress:=strExc(2)
                  
                  wksReport2.Range("A" & m_rpt2Row) = wksReport.Range("C" & ii)
                  wksReport2.Range("B" & m_rpt2Row) = wksReport.Range("D" & ii)
                  strExc(0) = "select * from salarylog" & _
                     " where sl01='" & .Fields("st01") & "' and (sl35 is null or sl03||sl35='TN') order by sl02 desc"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     Set adoRecordset = RsTemp.Clone
                     adoRecordset.MoveNext
                     With RsTemp
                     intI = 0
                     Do While Not .EOF
                        If Not adoRecordset.EOF Then
                           'Modified by Morgan 2023/10/6 至少抓1筆(才昇高專的要列最近的調薪)
                           If intI > 0 And adoRecordset.Fields("sl02") < strExc(1) And adoRecordset.Fields("sl03") <> "T" Then Exit Do
                           intI = intI + 1
                           wksReport2.Range("C" & m_rpt2Row) = ChangeWStringToTDateString(.Fields("sl02"))
                           '適用期滿
                           If .Fields("sl03") = "R" And adoRecordset.Fields("sl03") = "T" Then
                              wksReport2.Range("D" & m_rpt2Row) = "Y"
                           End If
                           
                           '本薪:基本薪資+午餐津貼
                           If Val("" & .Fields("sl11")) + Val("" & .Fields("sl14")) <> Val("" & adoRecordset.Fields("sl11")) + Val("" & adoRecordset.Fields("sl14")) Then
                              wksReport2.Range("E" & m_rpt2Row) = (Val("" & .Fields("sl11")) + Val("" & .Fields("sl14"))) - (Val("" & adoRecordset.Fields("sl11")) + Val("" & adoRecordset.Fields("sl14")))
                              wksReport2.Range("E" & m_rpt2Row).NumberFormatLocal = "+#,##0;-#,##0"
                              If wksReport2.Range("E" & m_rpt2Row) < 0 Then
                                 wksReport2.Range("E" & m_rpt2Row).Font.Color = vbRed
                              End If
                           End If
                           '職務津貼
                           If Val("" & .Fields("sl12")) <> Val("" & adoRecordset.Fields("sl12")) Then
                              wksReport2.Range("F" & m_rpt2Row) = Val("" & .Fields("sl12")) - Val("" & adoRecordset.Fields("sl12"))
                              wksReport2.Range("F" & m_rpt2Row).NumberFormatLocal = "+#,##0;-#,##0"
                              If wksReport2.Range("F" & m_rpt2Row) < 0 Then
                                 wksReport2.Range("F" & m_rpt2Row).Font.Color = vbRed
                              End If
                           End If
                           '技術津貼
                           If Val("" & .Fields("sl13")) <> Val("" & adoRecordset.Fields("sl13")) Then
                              wksReport2.Range("G" & m_rpt2Row) = Val("" & .Fields("sl13")) - Val("" & adoRecordset.Fields("sl13"))
                              wksReport2.Range("G" & m_rpt2Row).NumberFormatLocal = "+#,##0;-#,##0"
                              If wksReport2.Range("G" & m_rpt2Row) < 0 Then
                                 wksReport2.Range("G" & m_rpt2Row).Font.Color = vbRed
                              End If
                           End If
                           '差旅津貼
                           If Val("" & .Fields("sl15")) <> Val("" & adoRecordset.Fields("sl15")) Then
                              wksReport2.Range("H" & m_rpt2Row) = Val("" & .Fields("sl15")) - Val("" & adoRecordset.Fields("sl15"))
                              wksReport2.Range("H" & m_rpt2Row).NumberFormatLocal = "+#,##0;-#,##0"
                              If wksReport2.Range("H" & m_rpt2Row) < 0 Then
                                 wksReport2.Range("H" & m_rpt2Row).Font.Color = vbRed
                              End If
                           End If
                           '房租津貼
                           If Val("" & .Fields("sl16")) <> Val("" & adoRecordset.Fields("sl16")) Then
                              wksReport2.Range("I" & m_rpt2Row) = Val("" & .Fields("sl16")) - Val("" & adoRecordset.Fields("sl16"))
                              wksReport2.Range("I" & m_rpt2Row).NumberFormatLocal = "+#,##0;-#,##0"
                              If wksReport2.Range("I" & m_rpt2Row) < 0 Then
                                 wksReport2.Range("I" & m_rpt2Row).Font.Color = vbRed
                              End If
                           End If
                           '特支費
                           If Val("" & .Fields("sl17")) <> Val("" & adoRecordset.Fields("sl17")) Then
                              wksReport2.Range("J" & m_rpt2Row) = Val("" & .Fields("sl17")) - Val("" & adoRecordset.Fields("sl17"))
                              wksReport2.Range("J" & m_rpt2Row).NumberFormatLocal = "+#,##0;-#,##0"
                              If wksReport2.Range("J" & m_rpt2Row) < 0 Then
                                 wksReport2.Range("J" & m_rpt2Row).Font.Color = vbRed
                              End If
                           End If
                           '證照津貼
                           If Val("" & .Fields("sl39")) <> Val("" & adoRecordset.Fields("sl39")) Then
                              wksReport2.Range("K" & m_rpt2Row) = Val("" & .Fields("sl39")) - Val("" & adoRecordset.Fields("sl39"))
                              wksReport2.Range("K" & m_rpt2Row).NumberFormatLocal = "+#,##0;-#,##0"
                              If wksReport2.Range("K" & m_rpt2Row) < 0 Then
                                 wksReport2.Range("K" & m_rpt2Row).Font.Color = vbRed
                              End If
                           End If
                           '備註
                           wksReport2.Range("L" & m_rpt2Row) = "" & .Fields("sl37")
                           m_rpt2Row = m_rpt2Row + 1
                           adoRecordset.MoveNext
                        End If
                        .MoveNext
                     Loop
                     End With
                  End If
               End If
               wksReport.Range("H" & ii) = Val("" & .Fields("SD20")) + Val("" & .Fields("SD23")) '本薪
               wksReport.Range("H" & ii).NumberFormatLocal = "#,##0"
               wksReport.Range("I" & ii) = Val("" & .Fields("SD21")) + Val("" & .Fields("SD22")) + Val("" & .Fields("SD24")) + Val("" & .Fields("SD25")) + Val("" & .Fields("SD26")) + Val("" & .Fields("SD52")) '津貼
               wksReport.Range("I" & ii).NumberFormatLocal = "#,##0"
               wksReport.Range("J" & ii).NumberFormatLocal = "#,##0"
               wksReport.Range("K" & ii).Formula = "=J" & ii & "/H" & ii '本薪調升率
               wksReport.Range("K" & ii).NumberFormatLocal = "0.00%"
               wksReport.Range("L" & ii).NumberFormatLocal = "#,##0"
               wksReport.Range("M" & ii).Formula = "=L" & ii & "+J" & ii '含津貼調升金額
               wksReport.Range("M" & ii).NumberFormatLocal = "#,##0"
               wksReport.Range("N" & ii).Formula = "=M" & ii & "/H" & ii  '含津貼調升率
               wksReport.Range("N" & ii).NumberFormatLocal = "0.00%"
               wksReport.Range("O" & ii).Formula = "=H" & ii & "+I" & ii & "+M" & ii '調升後
               wksReport.Range("O" & ii).NumberFormatLocal = "#,##0"
               'Added by Morgan 2024/5/16
               wksReport.Range("P" & ii) = Val("" & .Fields("SD52")) '證照津貼
               wksReport.Range("P" & ii).NumberFormatLocal = "#,##0"
               'end 2024/5/16
            End If
            
         End If
         'end 2020/10/15
         
Nextstep:
         .MoveNext
      Loop
   
   If Text1(3) = "1" Then 'Added by Morgan 2021/4/28
   
      Printer.EndDoc
      Printer.FontSize = strFontSize
      Printer.FontName = strFontName
      
   End If 'Added by Morgan 2021/4/28
   
      'Added by Morgan 2020/10/15
      If bolSaveXLS Then
         If ii > iRow1 Then
            ii = ii + 1
            strExc(1) = GetSubDeptName(stLsdDeptNo, stLsdDeptName)
            strExc(1) = strExc(1) & "合計"
            
            SetSubTotRow wksReport, strExc(1), ii, iRow1
         End If
         
         ii = ii + 2
         SetSubTotRow wksReport, "全所合計", ii, 1
         
         '自動設定欄寬
         For intI = Asc("A") To Asc("P")
            wksReport.Columns(Chr(intI)).EntireColumn.AutoFit
         Next
         wksReport.Columns("B:B").HorizontalAlignment = xlCenter
         If Text1(3) = "2" Then
            wksReport.Columns("G:G").HorizontalAlignment = xlCenter
            wksReport.Range("A1").Select
            
            wksReport2.Columns("D:D").HorizontalAlignment = xlCenter
            For intI = Asc("A") To Asc("O")
               wksReport2.Columns(Chr(intI)).EntireColumn.AutoFit
            Next
         End If
         xlsReport.Workbooks(1).SaveAs strFileName
         xlsReport.Workbooks.Close
         xlsReport.Quit
      End If
      'end 2020/10/15
      
      End With
      
      'modify by sonia 2016/5/5
      'MsgBox "列印結束 !"
      If iPage > 0 Then
         'Added by Morgan 2020/10/16
         If bolSaveXLS Then
            'Modified by Morgan 2021/4/28
            'MsgBox "列印結束 !" & vbCrLf & "Excel檔已產生！（" & strFileName & "）", vbInformation
            'Modify by Amy 2021/06/22 路徑改中文字顯示
            MsgBox "Excel檔已產生！" & vbCrLf & "檔案存於" & strExcelPathN & Replace(strFileName, strExcelPath, ""), vbInformation
            'end 2021/4/28
         Else
         'end 2020/10/16
            MsgBox "列印結束 !"
         End If
      Else
         MsgBox "無資料 !"
      End If
      'end 2016/5/5
   Else
      MsgBox "無資料 !"
   End If
   Set adoRst = Nothing
End Sub

Private Sub FormReset()
   Dim oText  As TextBox
   
   For Each oText In Text1
      oText.Text = Empty
   Next
   'add by sonia 2016/4/11
   Text1(2).Enabled = False
   Text1(2).Locked = True
   'end 2016/4/11
   Label4 = "" 'Add By Sindy 2022/4/18
End Sub
'垂直列印
Private Sub VPrint(sWords As String, Optional iGap As Integer = 50)
   Dim Xo As Integer, Yo As Integer, ii As Integer
   Xo = Printer.CurrentX
   Yo = Printer.CurrentY
   For ii = 1 To Len(sWords)
      Printer.CurrentX = Xo
      Printer.CurrentY = Yo + (ii - 1) * (Printer.TextHeight(Space(1)) + iGap)
      Printer.Print Mid(sWords, ii, 1)
   Next
End Sub

'水平列印
Private Sub HPrint(sWords As String, Optional iGap As Integer = 50)
   Dim Xo As Long, Yo As Long, ii As Integer
   Xo = Printer.CurrentX
   Yo = Printer.CurrentY
   For ii = 1 To Len(sWords)
      Printer.CurrentX = Xo
      Printer.CurrentY = Yo
      'Add by Sindy 2022/4/18
      If Mid(sWords, ii, 1) <> StrConv(StrConv(Mid(sWords, ii, 1), vbFromUnicode), vbUnicode) Then
         PUB_PrintUnicodeText Mid(sWords, ii, 1), Xo, Yo
         Xo = Xo + Printer.TextWidth("　") + iGap
      Else
      'end 2022/4/18
         Printer.Print Mid(sWords, ii, 1)
         Xo = Xo + Printer.TextWidth(Mid(sWords, ii, 1)) + iGap
      End If 'Added by Morgan 2022/4/18
   Next
End Sub

Private Sub SetPx()
   Px(1) = Px(0) + 0.8 * TwPerCm
   Px(2) = Px(1) + 0.8 * TwPerCm
   Px(3) = Px(2) + 1.2 * TwPerCm
   Px(4) = Px(0) + 3.2 * TwPerCm
   Px(5) = Px(3) + 2.2 * TwPerCm
   Px(6) = Px(4) + 3.2 * TwPerCm
   Px(7) = Px(5) + 2.2 * TwPerCm
   Px(8) = Px(7) + 2 * TwPerCm
   Px(9) = Px(6) + 3.2 * TwPerCm
   Px(10) = Px(8) + 2 * TwPerCm
   Px(11) = Px(9) + 3.1 * TwPerCm
   Px(12) = Px(10) + 2 * TwPerCm
   Px(13) = Px(11) + 3.1 * TwPerCm
   Px(14) = Px(12) + 2 * TwPerCm
   Px(15) = Px(14) + 2 * TwPerCm
   Px(16) = Px(15) + 2 * TwPerCm
End Sub

Private Sub SetPy()
   For intI = 1 To 24
      'Py(intI) = Py(intI - 1) + 0.8 * TwPerCm
      Py(intI) = Py(intI - 1) + 1 * TwPerCm
   Next
End Sub

'表格
Private Sub PrintTable()
      
   Printer.DrawWidth = 5
   '框
   Printer.Line (Px(0), Py(0))-(Px(16), Py(3)), , B
   '縱線
   Printer.Line (Px(4), Py(1))-(Px(4), Py(3))
   Printer.Line (Px(6), Py(1))-(Px(6), Py(3))
   Printer.Line (Px(9), Py(0))-(Px(9), Py(3))
   Printer.Line (Px(11), Py(1))-(Px(11), Py(3))
   Printer.Line (Px(13), Py(1))-(Px(13), Py(3))
   
   '橫線
   Printer.Line (Px(0), Py(1))-(Px(16), Py(1))
   Printer.Line (Px(0), Py(2))-(Px(16), Py(2))
   
   '框
   Printer.Line (Px(0), Py(4))-(Px(16), Py(14)), , B
   '縱線
   'add by sonia 2017/4/7 因2017/1全體員工調整午餐津貼所以要多印再前一次的調薪
   
'   If Left(strSrvDate(1), 4) = 2017 Then
'      Printer.Line (700, Py(6))-(700, Py(24))     '1
'      Printer.Line (1100, Py(7))-(1100, Py(12))     '2 專利
'      Printer.Line (1100, Py(13))-(1100, Py(18))    '2 商標
'      Printer.Line (1100, Py(19))-(1100, Py(24))    '2 合計
'      Printer.Line (1730, Py(5))-(1730, Py(24))     '3
'      Printer.Line (2730, Py(5))-(2730, Py(24))     'n
'      Printer.Line (3730, Py(5))-(3730, Py(24))     '4
'      Printer.Line (4730, Py(4))-(4730, Py(24))     '5
'      Printer.Line (5795, Py(5))-(5795, Py(24))     '7
'      Printer.Line (6860, Py(5))-(6860, Py(24))     '8
'      Printer.Line (7925, Py(5))-(7925, Py(24))     '10
'      Printer.Line (8990, Py(5))-(8990, Py(24))     '12
'      Printer.Line (10055, Py(5))-(10055, Py(24))   '14
'
'      '橫線
'      For intI = 5 To 6
'         Printer.Line (Px(0), Py(intI))-(Px(16), Py(intI))
'      Next
'      Printer.Line (700, Py(7))-(Px(16), Py(7))
'      For intI = 8 To 11
'         Printer.Line (1100, Py(intI))-(Px(16), Py(intI))
'      Next
'      Printer.Line (Px(0), Py(12))-(Px(16), Py(12))
'      Printer.Line (700, Py(13))-(Px(16), Py(13))
'      For intI = 14 To 17
'         Printer.Line (1100, Py(intI))-(Px(16), Py(intI))
'      Next
'      Printer.Line (Px(0), Py(18))-(Px(16), Py(18))
'      Printer.Line (700, Py(19))-(Px(16), Py(19))
'      For intI = 20 To 23
'         Printer.Line (1100, Py(intI))-(Px(16), Py(intI))
'      Next
'   Else
'   'end 2017/4/7
      'Printer.Line (Px(1), Py(6))-(Px(1), Py(13))
      
      'Printer.Line (Px(2), Py(7))-(Px(2), Py(13))
      Printer.Line (Px(1), Py(7))-(Px(1), Py(13))
      
      'Printer.Line (Px(2), Py(13))-(Px(2), Py(18))
      'Printer.Line (Px(2), Py(19))-(Px(2), Py(13))
      Printer.Line (Px(3), Py(5))-(Px(3), Py(14))
      Printer.Line (Px(5), Py(5))-(Px(5), Py(14))
      Printer.Line (Px(7), Py(4))-(Px(7), Py(14))
      Printer.Line (Px(8), Py(5))-(Px(8), Py(14))
      Printer.Line (Px(10), Py(5))-(Px(10), Py(14))
      Printer.Line (Px(12), Py(5))-(Px(12), Py(14))
      Printer.Line (Px(14), Py(5))-(Px(14), Py(14))
      Printer.Line (Px(15), Py(5))-(Px(15), Py(14))
      
      '橫線
      For intI = 5 To 6
         Printer.Line (Px(0), Py(intI))-(Px(16), Py(intI))
      Next
      'Printer.Line (Px(1), Py(7))-(Px(16), Py(7))
      Printer.Line (Px(0), Py(7))-(Px(16), Py(7))
      For intI = 8 To 12 '11
         'Printer.Line (Px(2), Py(intI))-(Px(16), Py(intI))
         Printer.Line (Px(1), Py(intI))-(Px(16), Py(intI))
      Next
      Printer.Line (Px(0), Py(13))-(Px(16), Py(13))
      
      Printer.Line (Px(0), Py(14))-(Px(16), Py(14)) 'Add By Sindy 2020/6/23 表的最後橫線
      
'      Printer.Line (Px(0), Py(12))-(Px(16), Py(12))
'      Printer.Line (Px(1), Py(13))-(Px(16), Py(13))
'      For intI = 14 To 17
'         Printer.Line (Px(2), Py(intI))-(Px(16), Py(intI))
'      Next
'      Printer.Line (Px(0), Py(18))-(Px(16), Py(18))
'      Printer.Line (Px(1), Py(19))-(Px(16), Py(19))
'      For intI = 20 To 23
'         Printer.Line (Px(2), Py(intI))-(Px(16), Py(intI))
'      Next
'   End If
      
   Printer.DrawWidth = 1
End Sub

'add by sonia 2016/4/11
Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 2
         Label4 = ""
         If Text1(Index) <> "" Then
            If ChkStaffID(Text1(Index)) = True Then
               Cancel = True
            End If
            If ClsPDGetStaffN(Text1(Index), strExc(1), , True) = False Then
               Cancel = True
            Else
               Label4 = strExc(1)
            End If
         End If
   End Select
End Sub
'end 2016/4/11

'Added by Morgan 2020/10/16
Private Sub SetSubTotRow(wksReport, pSubName As String, pii As Integer, pRow1 As Integer)
   
   wksReport.Range("A" & pii) = pSubName '部門
   
   If Text1(3) = "3" Then
      wksReport.Range("G" & pii).Formula = "=SUMIF(C" & pRow1 & ":C" & pii - 1 & ",""*"",G" & pRow1 & ":G" & pii - 1 & ")" '本薪
      wksReport.Range("G" & pii).NumberFormatLocal = "#,##0"
      wksReport.Range("H" & pii).Formula = "=SUMIF(C" & pRow1 & ":C" & pii - 1 & ",""*"",H" & pRow1 & ":H" & pii - 1 & ")" '津貼
      wksReport.Range("H" & pii).NumberFormatLocal = "#,##0"
      wksReport.Range("I" & pii).Formula = "=SUMIF(C" & pRow1 & ":C" & pii - 1 & ",""*"",I" & pRow1 & ":I" & pii - 1 & ")" '本薪調升金額
      wksReport.Range("I" & pii).NumberFormatLocal = "#,##0"
      wksReport.Range("J" & pii).Formula = "=I" & pii & "/G" & pii '本薪調升率
      wksReport.Range("J" & pii).NumberFormatLocal = "0.00%"
      wksReport.Range("K" & pii).Formula = "=SUMIF(C" & pRow1 & ":C" & pii - 1 & ",""*"",K" & pRow1 & ":K" & pii - 1 & ")" '津貼調升金額 'Added by Morgan 2021/4/28
      wksReport.Range("L" & pii).NumberFormatLocal = "#,##0"
      wksReport.Range("L" & pii).Formula = "=K" & pii & "+I" & pii '含津貼調升金額
      wksReport.Range("M" & pii).Formula = "=L" & pii & "/G" & pii  '含津貼調升率
      wksReport.Range("M" & pii).NumberFormatLocal = "0.00%"
      wksReport.Range("N" & pii).Formula = "=G" & pii & "+H" & pii & "+L" & pii '調升後
      wksReport.Range("N" & pii).NumberFormatLocal = "#,##0"
      'Added by Morgan 2024/5/16
      wksReport.Range("O" & pii).Formula = "=SUMIF(C" & pRow1 & ":C" & pii - 1 & ",""*"",O" & pRow1 & ":O" & pii - 1 & ")" '證照津貼
      wksReport.Range("O" & pii).NumberFormatLocal = "#,##0"
      'end 2024/5/16
      '底色
      wksReport.Range("A" & pii & ":O" & pii).Interior.ThemeColor = 10
      wksReport.Range("A" & pii & ":O" & pii).Interior.TintAndShade = 0.8
   Else
      wksReport.Range("H" & pii).Formula = "=SUMIF(C" & pRow1 & ":C" & pii - 1 & ",""*"",H" & pRow1 & ":H" & pii - 1 & ")" '本薪
      wksReport.Range("H" & pii).NumberFormatLocal = "#,##0"
      wksReport.Range("I" & pii).Formula = "=SUMIF(C" & pRow1 & ":C" & pii - 1 & ",""*"",I" & pRow1 & ":I" & pii - 1 & ")" '津貼
      wksReport.Range("I" & pii).NumberFormatLocal = "#,##0"
      wksReport.Range("J" & pii).Formula = "=SUMIF(C" & pRow1 & ":C" & pii - 1 & ",""*"",J" & pRow1 & ":J" & pii - 1 & ")" '本薪調升金額
      wksReport.Range("J" & pii).NumberFormatLocal = "#,##0"
      wksReport.Range("K" & pii).Formula = "=J" & pii & "/H" & pii '本薪調升率
      wksReport.Range("K" & pii).NumberFormatLocal = "0.00%"
      wksReport.Range("L" & pii).Formula = "=SUMIF(C" & pRow1 & ":C" & pii - 1 & ",""*"",L" & pRow1 & ":L" & pii - 1 & ")" '津貼調升金額 'Added by Morgan 2021/4/28
      wksReport.Range("M" & pii).NumberFormatLocal = "#,##0"
      wksReport.Range("M" & pii).Formula = "=L" & pii & "+J" & pii '含津貼調升金額
      wksReport.Range("N" & pii).Formula = "=M" & pii & "/H" & pii  '含津貼調升率
      wksReport.Range("N" & pii).NumberFormatLocal = "0.00%"
      wksReport.Range("O" & pii).Formula = "=H" & pii & "+I" & pii & "+M" & pii '調升後
      wksReport.Range("O" & pii).NumberFormatLocal = "#,##0"
      'Added by Morgan 2024/5/16
      wksReport.Range("P" & pii).Formula = "=SUMIF(C" & pRow1 & ":C" & pii - 1 & ",""*"",P" & pRow1 & ":P" & pii - 1 & ")" '證照津貼
      wksReport.Range("P" & pii).NumberFormatLocal = "#,##0"
      'end 2024/5/16
      '底色
      wksReport.Range("A" & pii & ":P" & pii).Interior.ThemeColor = 10
      wksReport.Range("A" & pii & ":P" & pii).Interior.TintAndShade = 0.8
   
   End If
End Sub

'Added by Morgan 2020/10/20
'合計用部門名稱
Private Function GetSubDeptName(pDeptNo As String, pDeptName As String, Optional pType As Integer = 0) As String
   If pType = 1 Then
      If Left(pDeptNo, 1) = "F" Or Left(pDeptNo, 1) = "S" Or Left(pDeptNo, 1) = "M" Then
         If Left(pDeptNo, 1) = "F" Then
            GetSubDeptName = "國外部"
         ElseIf Left(pDeptNo, 1) = "S" Then
            GetSubDeptName = "智權部"
         ElseIf Left(pDeptNo, 1) = "M" Then
            GetSubDeptName = "管理部"
         End If
      End If
   Else
      If Left(pDeptNo, 1) <> "M" And Left(pDeptNo, 1) <> "S" And Left(pDeptNo, 2) <> "F6" Then
         If Left(pDeptNo, 2) = "F7" Then
            GetSubDeptName = "顧問"
         ElseIf Left(pDeptNo, 2) = "W2" Then
            GetSubDeptName = "顧服組"
         ElseIf Left(pDeptNo, 2) = "F1" Or Left(pDeptNo, 2) = "F2" Then
            GetSubDeptName = Left(pDeptName, 2)
         Else
            GetSubDeptName = Left(pDeptName, 3)
         End If
      End If
   End If
End Function

'Added by Morgan 2023/11/27
'當年度人事異動作業有「05晉升」、「12任命」、「13晉升、任命」者，且年資在25年以上者，當年11月份調薪作業時，系統仍會列印出調薪單--劉柏翰
Private Function ChkThisYearPromote(pNo As String) As Boolean
   Dim intQ As Integer, stSQL As String
   Dim rsQuery As ADODB.Recordset
   
   stSQL = "select * from staff_change where sc01='" & pNo & "' and sc02>" & Left(strSrvDate(1), 4) & "0000 and sc03 in ('05','12','13')"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      ChkThisYearPromote = True
   End If
   Set rsQuery = Nothing
End Function
