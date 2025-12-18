VERSION 5.00
Begin VB.Form frm170209 
   BorderStyle     =   1  '單線固定
   Caption         =   "薪資入帳明細"
   ClientHeight    =   3012
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4764
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3012
   ScaleWidth      =   4764
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   1860
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "1"
      Top             =   1230
      Width           =   230
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   30
      TabIndex        =   3
      Top             =   2400
      Width           =   4665
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   4
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   5
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2670
      TabIndex        =   8
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   3690
      TabIndex        =   9
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   0
      Left            =   1860
      MaxLength       =   5
      TabIndex        =   1
      Top             =   810
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PS：入帳類別資料排除關係企業"
      Height          =   180
      Left            =   936
      TabIndex        =   10
      Top             =   1800
      Width           =   2532
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "排序條件："
      Height          =   180
      Left            =   936
      TabIndex        =   7
      Top             =   1260
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(1.公司別 2.入帳類別)"
      Height          =   180
      Left            =   2316
      TabIndex        =   6
      Top             =   1260
      Width           =   1692
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "薪資月份："
      Height          =   180
      Index           =   0
      Left            =   936
      TabIndex        =   0
      Top             =   876
      Width           =   900
   End
End
Attribute VB_Name = "frm170209"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Memo by Morgan 2024/1/31 新部門已修改(不必改)
'Create by SINDY 2008/12/30
Option Explicit
Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 10) As Integer
Dim strTemp(1 To 10) As String
'Dim PaperX As Double
'Dim paperY As Double
Dim iPgae As Integer, iLine As Integer
Dim strType As String

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
        If txt1(0) = "" Then
            MsgBox "薪資月份不可以空白！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
        End If
        If txt1(0) <> "" Then
            If Len(txt1(0)) <= 3 Then
                MsgBox "薪資月份輸入錯誤！", vbInformation, "操作錯誤！"
                txt1(0).SetFocus
                Exit Sub
            End If
            If ChkDate(txt1(0) & "01") = False Then
                txt1(0).SetFocus
                Exit Sub
            End If
        End If
        If txt1(1) = "" Then
            MsgBox "排序條件不可以空白！", vbInformation, "操作錯誤！"
            txt1(1).SetFocus
            Exit Sub
        End If
        If txt1(1) <> "" Then
            If txt1(1) <> "1" And txt1(1) <> "2" Then
               MsgBox "排序條件輸入錯誤！", vbInformation, "操作錯誤！"
               txt1(1).SetFocus
               Exit Sub
            End If
        End If
        
        Screen.MousePointer = vbHourglass
        m_StrSQL = ""
        If txt1(1) = "1" Then
            m_StrSQL = m_StrSQL & " Order by T3,SD01 "
        ElseIf txt1(1) = "2" Then
            m_StrSQL = m_StrSQL & " Order by SD05,SD01 "
        End If
        StrMenu1 '1.薪資入帳明細(非F編號)
        StrMenu2 '2.外翻人員翻譯費明細表(F編號)
        Screen.MousePointer = vbDefault
      Case 1
        Unload Me
   End Select
End Sub

'1.薪資入帳明細(非F編號)
Sub StrMenu1()
Dim strYM As String
Dim dblAmt As Double      '小計
Dim dblTotAmt As Double '合計
Dim dblCnt As Double
Dim dblTotCnt As Double
Dim m_strSM37 As String      'add by sonia 2018/11/29 入帳類別資料排除關係企業 台一投資 及 台一智權

   strYM = Left(ChangeTStringToWString(txt1(0) & "01"), 6)
   
   'add by sonia 2018/11/29 入帳類別資料排除關係企業 台一投資 及 台一智權
   m_strSM37 = ""
   If txt1(1) = "2" Then
      'modify by sonia 2020/5/4 取消1公司加入L公司中所
      'm_strSM37 = " AND t3 NOT IN ('A','J') "
      m_strSM37 = " AND ((t3='2' or t3='L' and sd05='4'))"
   End If
   'end 2018/11/29
   
   Set Printer = Printers(Combo1.ListIndex)
   Printer.EndDoc
   Printer.Orientation = 1 '1.直印 2.橫印
   'Printer.PaperSize = 9  'PDF
   
   '2009/5/26 modify by sonia 5月調薪且合併者勞退自6月才合併,故5月第二家薪資無薪資但有勞退,此類資料合併至第一家入帳
   'm_str = "SELECT SD01,ST02,SD05,SD06,T.*,SUBSTR(a0802,1,12) A0802 " & _
                   "FROM Staff,SalaryData,( " & _
                  "SELECT replace(SM01,'A','0') as T1,nvl(SM04,0)+nvl(SM05,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0) as T2,SM37 as T3,SM01 as T4,SM03 " & _
                  "From SalaryMonth " & _
                  "WHERE SM02='" & strYM & "') T,acc080 " & _
                  "Where ST01 = t1 AND SD01=T1 AND T3=a0801(+) " & _
                  "AND (T1 NOT Like 'F%' Or (T1 Like 'F%' AND SM03<>'F51')) " & m_StrSQL
   '先抓第二家若為<0者合併至第一家計算,第二家只抓>0的資料
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'Modified by Morgan 2013/2/1 +sm43
   'modify by sonia 2018/11/29 +m_strSM37條件,入帳類別資料排除關係企業 台一投資 及 台一智權
   'Modify By Sindy 2020/8/4 + sm45
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   m_str = "SELECT SD01,ST02,SD05,SD06,T.*,SUBSTR(a0802,1,12) A0802 FROM Staff,SalaryData,acc080,( " & _
            "select t1,sum(t2) t2,s.SM37 as t3,s.SM01 as t4,s.SM03 from salarymonth s,( " & _
            "SELECT SM01 as T1,SM02,nvl(SM04,0)+nvl(SM05,0)+nvl(SM06,0)+nvl(SM45,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0) as T2 From SalaryMonth " & _
            "WHERE SUBSTR(SM01,3,1)<>'A' and SM02=" & strYM & " UNION " & _
            "SELECT substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4) as T1,SM02,nvl(SM04,0)+nvl(SM05,0)+nvl(SM06,0)+nvl(SM45,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0) as T2 From SalaryMonth " & _
            "WHERE SUBSTR(SM01,3,1)='A' AND SM02=" & strYM & " " & _
            "AND nvl(SM04,0)+nvl(SM05,0)+nvl(SM06,0)+nvl(SM45,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0)<=0 " & _
            ") X where x.t1=sm01 and x.sm02=s.sm02 group by T1,s.SM37 ,s.SM01,s.SM03 " & _
            "UNION SELECT substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4) as T1,nvl(SM04,0)+nvl(SM05,0)+nvl(SM06,0)+nvl(SM45,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0) as T2,SM37 as T3,SM01 as T4,SM03 From SalaryMonth " & _
            "WHERE SUBSTR(SM01,3,1)='A' AND nvl(SM04,0)+nvl(SM05,0)+nvl(SM06,0)+nvl(SM45,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0)>0 AND SM02=" & strYM & " " & _
            ") T Where ST01 = t1 AND SD01=T1 AND T3=a0801(+) " & _
            "AND (T1 NOT Like 'F%' Or (T1 Like 'F%' AND SM03<>'F51')) " & m_strSM37 & m_StrSQL
   '2009/5/26 END
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
       With m_rs
           m_rs.MoveFirst
           
           iLine = 1
           PrintTitle '列印表頭
           strType = "" '切頁條件
           dblAmt = 0
           dblTotAmt = 0
           dblCnt = 0
           dblTotCnt = 0
           Do While Not m_rs.EOF
               
               For m_i = 1 To 10
                   strTemp(m_i) = ""
               Next m_i
               
               strTemp(1) = CheckStr(m_rs.Fields("SD01"))
               strTemp(2) = CheckStr(m_rs.Fields("ST02"))
               strTemp(3) = CheckStr(m_rs.Fields("SD05")) '入帳類別
               strTemp(4) = CheckStr(m_rs.Fields("SD06"))
               strTemp(5) = CheckStr(m_rs.Fields("T4"))
               strTemp(6) = CheckStr(m_rs.Fields("T2"))
               strTemp(7) = CheckStr(m_rs.Fields("a0802")) '公司別
               
               If strType <> "" Then
                  If iLine > 50 Or _
                        (strType <> strTemp(7) And txt1(1) = "1") Or _
                        (strType <> strTemp(3) And txt1(1) = "2") Then
                        
                      If (strType <> strTemp(7) And txt1(1) = "1") Or _
                         (strType <> strTemp(3) And txt1(1) = "2") Then
                         
                         Printer.CurrentX = 500
                         Printer.CurrentY = iLine * 300
                         Printer.Print String(140, "-")
                         
                         iLine = iLine + 1
                         Printer.CurrentX = PLeft(4)
                         Printer.CurrentY = iLine * 300
                         Printer.Print "小計：(" & dblCnt & "人)"
                         Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblAmt, "##,##0"))
                         Printer.CurrentY = iLine * 300
                         Printer.Print Format(dblAmt, "##,##0")
                         
                         dblAmt = 0 '小計
                         dblCnt = 0
                      End If
                      
                      'If .AbsolutePosition <> .RecordCount Then
                          Printer.NewPage
                          iLine = 1
                          PrintTitle '列印表頭
                      'End If
                  End If
               End If
               
               PrintDetail '列印表中
               
               If txt1(1) = "1" Then
                  '公司別
                  strType = strTemp(7)
               ElseIf txt1(1) = "2" Then
                  '入帳類別
                  strType = strTemp(3)
               End If
               
               dblAmt = dblAmt + strTemp(6)  '小計
               dblTotAmt = dblTotAmt + strTemp(6)  '合計
               dblCnt = dblCnt + 1
               dblTotCnt = dblTotCnt + 1
               m_rs.MoveNext
           Loop
            
            '列印表尾
            Printer.CurrentX = 500
            Printer.CurrentY = iLine * 300
            Printer.Print String(140, "-")
            
            iLine = iLine + 1
            Printer.CurrentX = PLeft(4)
            Printer.CurrentY = iLine * 300
            Printer.Print "小計：(" & dblCnt & "人)"
            Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblAmt, "##,##0"))
            Printer.CurrentY = iLine * 300
            Printer.Print Format(dblAmt, "##,##0")
            
            iLine = iLine + 1
            Printer.CurrentX = 500
            Printer.CurrentY = iLine * 300
            Printer.Print String(140, "-")
            
            iLine = iLine + 1
            Printer.CurrentX = PLeft(4)
            Printer.CurrentY = iLine * 300
            Printer.Print "合計：(" & dblTotCnt & "人)"
            Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblTotAmt, "##,##0"))
            Printer.CurrentY = iLine * 300
            Printer.Print Format(dblTotAmt, "##,##0")
       End With
   Else
       MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
       Exit Sub
   End If
   Printer.EndDoc
   ShowPrintOk
End Sub

Sub PrintTitle()
   GetPleft
   
   'PaperX = 12000
   'paperY = 7500
   
   Printer.Font.Size = 12
   Printer.Font.Underline = False
   Printer.FontBold = False
   
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("薪　資　入　帳　明　細") / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print "薪　資　入　帳　明　細"
   
   iLine = iLine + 1
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("000 年 00 月") / 2)
   Printer.CurrentY = iLine * 300
   If Len(txt1(0)) = 5 Then
      Printer.Print Left(Trim(txt1(0)), 3) & "  年  " & Right(Trim(txt1(0)), 2) & "  月"
   Else
      Printer.Print Left(Trim(txt1(0)), 2) & "  年  " & Right(Trim(txt1(0)), 2) & "  月"
   End If
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "頁　　次：" & Printer.Page
   
   iLine = iLine + 2
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   If txt1(1) = "1" Then
      Printer.Print "公司別"
   ElseIf txt1(1) = "2" Then
      Printer.Print "入帳類別"
   End If
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print "帳　　號"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iLine * 300
   Printer.Print "員工代號"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iLine * 300
   Printer.Print "姓　名"
   Printer.CurrentX = PLeft(5) - Printer.TextWidth("金　額")
   Printer.CurrentY = iLine * 300
   Printer.Print "金　額"
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(140, "-")
   
   iLine = iLine + 1
End Sub

Sub GetPleft()
   PLeft(1) = 500
   PLeft(2) = 3500
   PLeft(3) = 6000
   PLeft(4) = 7500
   PLeft(5) = 10500
End Sub

Sub PrintDetail()
   '1.公司別/2.入帳類別
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   If txt1(1) = "1" Then
      'If strType <> strTemp(7) Then
      If iLine = 7 Then
         Printer.Print strTemp(7)
      End If
   ElseIf txt1(1) = "2" Then
      'If strType <> strTemp(3) Then
      If iLine = 7 Then
         If strTemp(3) = "1" Then
            Printer.Print "現金"
         ElseIf strTemp(3) = "2" Then
            Printer.Print "北所"
         ElseIf strTemp(3) = "3" Then
            Printer.Print "匯款"
         ElseIf strTemp(3) = "4" Then
            Printer.Print "中所"
         ElseIf strTemp(3) = "5" Then
            Printer.Print "南所"
         ElseIf strTemp(3) = "6" Then
            Printer.Print "高所"
         ElseIf strTemp(3) = "7" Then
            Printer.Print "其他"
         Else
            Printer.Print strTemp(3)
         End If
      End If
   End If
   '帳　　號
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(4)
   '員工代號
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(5)
   '姓　　名
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(2)
   '金　額
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(strTemp(6), "##,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(6), "##,##0")
   
   iLine = iLine + 1
End Sub

'2.外翻人員翻譯費明細表(F編號)
Sub StrMenu2()
Dim strYM As String
Dim dblTotAmt As Double '合計
Dim dblCnt As Double

   strYM = Left(ChangeTStringToWString(txt1(0) & "01"), 6)
   
   Set Printer = Printers(Combo1.ListIndex)
   Printer.EndDoc
   Printer.Orientation = 1 '1.直印 2.橫印
   'Printer.PaperSize = 9  'PDF
   'Modified by Morgan 2013/2/1 +sm43
   'Modify By Sindy 2020/8/4 + sm45
   m_str = "SELECT SD01,ST02,SD05,SD06,T.*,SUBSTR(a0802,1,12) A0802 " & _
                   "FROM Staff,SalaryData,( " & _
                  "SELECT SM01 as T1,nvl(SM04,0)+nvl(SM05,0)+nvl(SM06,0)+nvl(SM45,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0) as T2,SM37 as T3,SM01 as T4,SM03 " & _
                  "From SalaryMonth " & _
                  "WHERE SM02='" & strYM & "') T,acc080 " & _
                  "Where ST01 = t1 AND SD01=T1 AND T3=a0801(+) " & _
                  "AND (T1 Like 'F%' AND SM03='F51') " & m_StrSQL
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
       With m_rs
           m_rs.MoveFirst
           
           iLine = 1
           PrintTitle2 '列印表頭
           'strType = "" '切頁條件
           dblTotAmt = 0
           dblCnt = 0
           Do While Not m_rs.EOF
               
               For m_i = 1 To 10
                   strTemp(m_i) = ""
               Next m_i
               
               strTemp(1) = CheckStr(m_rs.Fields("SD01"))
               strTemp(2) = CheckStr(m_rs.Fields("ST02"))
               strTemp(3) = CheckStr(m_rs.Fields("SD05")) '入帳類別
               strTemp(4) = CheckStr(m_rs.Fields("SD06"))
               strTemp(5) = CheckStr(m_rs.Fields("T4"))
               strTemp(6) = CheckStr(m_rs.Fields("T2"))
               strTemp(7) = CheckStr(m_rs.Fields("a0802")) '公司別
               
               'If strType <> "" Then
                  If iLine > 50 Then
                      'If .AbsolutePosition <> .RecordCount Then
                          Printer.NewPage
                          iLine = 1
                          PrintTitle2 '列印表頭
                      'End If
                  End If
               'End If
               
               PrintDetail2 '列印表中
               
               dblTotAmt = dblTotAmt + strTemp(6)  '合計
               dblCnt = dblCnt + 1
               m_rs.MoveNext
           Loop
            
            '列印表尾
            Printer.CurrentX = 500
            Printer.CurrentY = iLine * 300
            Printer.Print String(140, "-")
            
            iLine = iLine + 1
            Printer.CurrentX = PLeft(2)
            Printer.CurrentY = iLine * 300
            Printer.Print "合計：(" & dblCnt & "人)"
            Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblTotAmt, "##,##0"))
            Printer.CurrentY = iLine * 300
            Printer.Print Format(dblTotAmt, "##,##0")
       End With
   Else
       MsgBox "無符合列印的翻譯費資料!!!", vbExclamation + vbOKOnly
       Exit Sub
   End If
   Printer.EndDoc
   ShowPrintOk
End Sub

Sub PrintTitle2()
   GetPleft2
   
   'PaperX = 12000
   'paperY = 7500
   
   Printer.Font.Size = 12
   Printer.Font.Underline = False
   Printer.FontBold = False
   
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("外翻人員翻譯費明細表") / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print "外翻人員翻譯費明細表"
   
   iLine = iLine + 1
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("000 年 00 月") / 2)
   Printer.CurrentY = iLine * 300
   If Len(txt1(0)) = 5 Then
      Printer.Print Left(Trim(txt1(0)), 3) & "  年  " & Right(Trim(txt1(0)), 2) & "  月"
   Else
      Printer.Print Left(Trim(txt1(0)), 2) & "  年  " & Right(Trim(txt1(0)), 2) & "  月"
   End If
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "頁　　次：" & Printer.Page
   
   iLine = iLine + 2
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "員工代號"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print "姓　名"
   Printer.CurrentX = PLeft(3) - Printer.TextWidth("金　額")
   Printer.CurrentY = iLine * 300
   Printer.Print "金　額"
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(140, "-")
   
   iLine = iLine + 1
End Sub

Sub GetPleft2()
   PLeft(1) = 3500
   PLeft(2) = 5500
   PLeft(3) = 8000
End Sub

Sub PrintDetail2()
   '員工代號
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(5)
   '姓　　名
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(2)
   '金　額
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(strTemp(6), "##,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(6), "##,##0")
   
   iLine = iLine + 1
End Sub

Private Sub Form_Load()
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim strSql As String, i As Integer, j As Integer
Dim strSystemKind As String
   
   MoveFormToCenter Me
   
   strSystemKind = GetSystemKindByNick
   strSql = Printer.DeviceName
   SeekPrintL = Printer.Orientation
   For i = 0 To Printers.Count - 1
      Set Printer = Printers(i)
      Combo1.AddItem Printer.DeviceName, j
      j = j + 1
      If Printer.DeviceName = strSql Then
         SeekPrint = i
      End If
   Next i
   
   Set Printer = Printers(SeekPrint)
   Combo1.Text = Combo1.List(SeekPrint)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170209 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 2, 3
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         If txt1(Index) <> "" Then
            If ChkDate(txt1(Index) & "01") = False Then
                Call txt1_GotFocus(Index)
                Cancel = True
                Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub
