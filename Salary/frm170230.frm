VERSION 5.00
Begin VB.Form frm170230 
   BorderStyle     =   1  '單線固定
   Caption         =   "薪資媒體轉帳遞送單及員工異動清冊"
   ClientHeight    =   3072
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4776
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3072
   ScaleWidth      =   4776
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   2
      Left            =   1380
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1485
      Width           =   855
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   0
      Left            =   1380
      MaxLength       =   7
      TabIndex        =   0
      Top             =   585
      Width           =   855
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   1
      Left            =   1380
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "1"
      Top             =   1020
      Width           =   300
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   30
      TabIndex        =   7
      Top             =   2280
      Width           =   4665
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   3
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   8
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
      TabIndex        =   4
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   3690
      TabIndex        =   5
      Top             =   60
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "(1：每月薪資  2：年終獎金    3：端午代金  4：中秋代金)"
      Height          =   360
      Left            =   1920
      TabIndex        =   11
      Top             =   1020
      Width           =   2235
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "委託日期："
      Height          =   180
      Index           =   3
      Left            =   480
      TabIndex        =   10
      Top             =   1530
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "資料類別："
      Height          =   180
      Index           =   2
      Left            =   450
      TabIndex        =   9
      Top             =   1020
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "入帳日期："
      Height          =   180
      Index           =   0
      Left            =   450
      TabIndex        =   6
      Top             =   630
      Width           =   900
   End
End
Attribute VB_Name = "frm170230"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'2009/1/23 add by sonia
Option Explicit

Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL1 As String
Dim m_StrSQL2 As String
Dim m_i As Integer
Dim PLeft(1 To 10) As Integer, TPLeft(1 To 10) As Integer
Dim strTemp(1 To 10) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String, strType1 As String
Dim dblAmtChange As Double, dblCntChange As Double
Dim dblAmtTotal As Double, dblCntTotal As Double
Dim m_Date1 As String, m_Date2 As String
Dim dblA0802 As String
Dim strSM37 As String    'add by sonia 2018/9/25


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         If txt1(0) = "" Then
            MsgBox "入帳日期不可以空白！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
         End If
         If txt1(0) <> "" Then
            If ChkDate(txt1(0)) = False Then
               txt1(0).SetFocus
               Exit Sub
            End If
            If ChkWork(ChangeTStringToWString(txt1(0))) = False Then
               txt1(0).SetFocus
               Exit Sub
            End If
         End If
         If txt1(1) = "" Then
            MsgBox "資料類別不可以空白！", vbInformation, "操作錯誤！"
            txt1(1).SetFocus
            Exit Sub
         End If
         If txt1(1) <> "" Then
            If txt1(1) <> "1" And txt1(1) <> "2" And txt1(1) <> "3" And txt1(1) <> "4" Then
               MsgBox "資料類別輸入錯誤！", vbInformation, "操作錯誤！"
               txt1(1).SetFocus
               Exit Sub
            End If
         End If
         If txt1(2) = "" Then
            MsgBox "委託日期不可以空白！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
         End If
         If txt1(2) <> "" Then
            If ChkDate(txt1(2)) = False Then
               txt1(2).SetFocus
               Exit Sub
            End If
            If ChkWork(ChangeTStringToWString(txt1(2))) = False Then
               txt1(2).SetFocus
               Exit Sub
            End If
         End If
        
         Screen.MousePointer = vbHourglass
         StrMenu
         Screen.MousePointer = vbDefault
      
      Case 1
         Unload Me
   End Select
End Sub

Sub StrMenu()
Dim strYM As String
Dim stDate As String  '2010/11/11 add by sonia
Dim stDate2 As String 'Added by Morgan 2024/2/21
   
   Set Printer = Printers(Combo1.ListIndex)
   Printer.EndDoc
   Printer.Orientation = 1 '1.直印 2.橫印
   
   m_StrSQL1 = "": m_StrSQL2 = ""
   '轉帳日期
   m_Date1 = Left(ChangeTStringToWString(txt1(0)), 4) - 1911 & " 年 " & Mid(ChangeTStringToWString(txt1(0)), 5, 2) & " 月 " & Right(ChangeTStringToWString(txt1(0)), 2) & " 日"
   '委託日期
   m_Date2 = Left(ChangeTStringToWString(txt1(2)), 4) - 1911 & " 年 " & Mid(ChangeTStringToWString(txt1(2)), 5, 2) & " 月 " & Right(ChangeTStringToWString(txt1(2)), 2) & " 日"
   
   If txt1(1) = "1" Then   '薪資
      strYM = Left(ChangeTStringToWString(txt1(0)), 6)
      'Modified by Morgan 2022/1/27
      '若日期小於15號時才抓上月薪資 Ex:1110128提早發月薪
      If Right(txt1(0), 2) < "15" Then
         If Right(strYM, 2) = "01" Then
            strYM = Left(strYM, 4) - 1 & "12"
         Else
            strYM = strYM - 1
         End If
      End If
      
      m_StrSQL1 = m_StrSQL1 & " and SM02='" & strYM & "' "
            
      'Modified by Morgan 2022/7/18 離職人員隔月出清冊(1號離職還是當月出),111/6沒給所以111/7可正常出 --辜
      'm_StrSQL2 = m_StrSQL2 & " and SC02>='" & strYM & "'||'01' AND SC02<='" & strYM & "'||'31' "
      m_StrSQL2 = m_StrSQL2 & " and SC02>'" & CompDate(1, -1, strYM & "01") & "' AND SC02<='" & strYM & "01' "
      'end 2022/7/18
      
      '2009/5/26 modify by sonia 5月調薪且合併者勞退自6月才合併,故5月第二家薪資無薪資但有勞退,此表只抓金額合計<>0
      'm_str = "SELECT DECODE(SC03,NULL,'0','01','1','02','1','2') 異動原因,ST02 姓名,SUBSTR(SD06,1,4)||'-'||SUBSTR(SD06,5,2)||'-'||SUBSTR(SD06,7,6)||'-'||SUBSTR(SD06,13,1) 帳號, " & _
              "nvl(SM04,0)+nvl(SM05,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0) 轉帳金額 " & _
              "FROM SALARYMONTH,STAFF,SALARYDATA," & _
              "(SELECT SC01,SC03 FROM STAFF_CHANGE WHERE SC03 IN ('01','02','03','04','08','09','10') " & m_StrSQL2 & ") C " & _
              "WHERE REPLACE(SM01,'A','0')=ST01(+) AND REPLACE(SM01,'A','0')=SD01(+) AND '2'=SD05 AND REPLACE(SM01,'A','0')=C.SC01(+) " & m_StrSQL1 & _
              " ORDER BY 異動原因,帳號,SM01"
      '2010/11/11 MODIFY BY SONIA 新進及復職者改抓SD46帳號通知銀行薪資年月
      'm_str = "SELECT DECODE(SC03,NULL,'0','01','1','02','1','2') 異動原因,ST02 姓名,SUBSTR(SD06,1,4)||'-'||SUBSTR(SD06,5,2)||'-'||SUBSTR(SD06,7,6)||'-'||SUBSTR(SD06,13,1) 帳號, " & _
              "nvl(SM04,0)+nvl(SM05,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0) 轉帳金額 " & _
              "FROM SALARYMONTH,STAFF,SALARYDATA," & _
              "(SELECT SC01,SC03 FROM STAFF_CHANGE WHERE SC03 IN ('01','02','03','04','08','09','10') " & m_StrSQL2 & ") C " & _
              "WHERE nvl(SM04,0)+nvl(SM05,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)<>0 AND REPLACE(SM01,'A','0')=ST01(+) AND REPLACE(SM01,'A','0')=SD01(+) AND '2'=SD05 AND REPLACE(SM01,'A','0')=C.SC01(+) " & m_StrSQL1 & _
              " ORDER BY 異動原因,帳號,SM01"
      'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
      'Modified by Morgan 2013/2/1 +sm43
      'Modified by Morgan 2022/1/27
      'stDate = Val(Left(CompDate("1", -1, DBDATE(txt1(0))), 6))
      stDate = strYM
      'end 2022/1/27
      'modify by sonia 2018/9/25 +SM37 北所台一投資及台一智權產生不同媒體檔案
      'modify by sonia 2020/5/4 +L公司故DECODE(SM37,'A',SM37,'J',SM37,'1') SM37直接抓SM37
      'Modify By Sindy 2020/6/25 + 證照津貼
      'Modified by Morgan 2022/7/18 異動排除留職停薪04--辜
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      m_str = "SELECT DECODE(SC03,NULL,'0','2') 異動原因,ST02 姓名,SUBSTR(SD06,1,4)||'-'||SUBSTR(SD06,5,2)||'-'||SUBSTR(SD06,7,6)||'-'||SUBSTR(SD06,13,1) 帳號, " & _
              "nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0) 轉帳金額,SM01,SM37 " & _
              "FROM SALARYMONTH,STAFF,SALARYDATA," & _
              "(SELECT SC01,SC03 FROM STAFF_CHANGE WHERE SC03 IN ('03','08','09','10') " & m_StrSQL2 & ") C " & _
              "WHERE nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0)<>0 AND substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)=ST01(+) AND substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)=SD01(+) AND '2'=SD05 AND substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)=C.SC01(+) " & m_StrSQL1 & _
              " UNION SELECT '1' 異動原因,ST02 姓名,SUBSTR(SD06,1,4)||'-'||SUBSTR(SD06,5,2)||'-'||SUBSTR(SD06,7,6)||'-'||SUBSTR(SD06,13,1) 帳號, " & _
              "nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0) 轉帳金額,SM01,SM37 " & _
              "FROM SALARYMONTH,STAFF,SALARYDATA " & _
              "WHERE nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0)<>0 AND substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)=ST01(+) AND substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)=SD01(+) AND '2'=SD05 AND SD46=" & stDate & m_StrSQL1
      '2011/6/2 add by sonia 每月1日離職的未帶出來 69006
      'Modify By Sindy 2020/6/25 + 證照津貼
      'Modified by Morgan 2022/7/18 異動排除留職停薪04--辜
       m_str = m_str & " UNION SELECT DECODE(SC03,NULL,'0','2') 異動原因,ST02 姓名,SUBSTR(SD06,1,4)||'-'||SUBSTR(SD06,5,2)||'-'||SUBSTR(SD06,7,6)||'-'||SUBSTR(SD06,13,1) 帳號, " & _
               "nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0) 轉帳金額,SM01,SM37 " & _
               "FROM SALARYMONTH,STAFF,SALARYDATA," & _
               "(SELECT SC01,SC03 FROM STAFF_CHANGE WHERE SC03 IN ('03','08','09','10') " & m_StrSQL2 & ") C " & _
               "WHERE C.SC01=ST01(+) AND C.SC01=SD01(+) AND '2'=SD05 AND C.SC01=sm01(+) and " & strYM & "=SM02(+) and sm01 is null and sd46 is not null"
       '2011/6/2 end
      '2010/11/11 END
      '2009/5/26 END
      'modify by sonia 2018/9/26 +SM37 北所台一投資及台一智權產生不同媒體檔案, 改只1抓異動原因非'0'的
       m_str = m_str & " ORDER BY 異動原因,帳號,SM01"
       'modify by sonia 2020/5/4 +L公司故DECODE(SM37,'A',SM37,'J',SM37,'1') SM37直接抓SM37
       m_str = "SELECT * FROM (" & m_str & ") WHERE 異動原因<>'0' ORDER BY SM37,異動原因,帳號,SM01"
       
   ElseIf txt1(1) = "2" Then  '年終
      strYM = Left(ChangeTStringToWString(txt1(0)), 4)
      strYM = strYM - 1

      m_StrSQL1 = m_StrSQL1 & " and YB01='" & strYM & "' "
      
      'Add by Sindy 2023/12/1 離職人員隔月出清冊(1號離職還是當月出)
      m_StrSQL2 = m_StrSQL2 & " and SC02>'" & CompDate(1, -1, Left(txt1(0), Len(txt1(0)) - 2) & "01") & "' AND SC02<='" & CStr(Left(txt1(0), Len(txt1(0)) - 2) + 191100) & "01' "
      
      'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
      '2013/1/23 modify by sonia 加扣補充保費yb25
      'modify by sonia 2016/1/7 留職停薪人員改發現金,不入媒體,故加SD02<>'S'
      'modify by sonia 2018/1/11 +YB26
      'modify by sonia 2018/9/27 +'1' SM37
      'modify by sonia 2020/5/12 '1' SM37->'2' SM37
      'Modify By Sindy 2023/12/1 資料類別 2年終獎金 3端午代金 4中秋代金 都要產出瑞興的"員工異動清冊"
      '原程式='0' 異動原因
      'Modified by Morgan 2024/2/1 +新進也要印 Ex:112年終比113/1月薪資早轉帳(但1月薪資的清冊要手動劃掉)
      'Modifidd by Morgan 2024/2/21 新進有設當月或前月開始轉帳且薪資未入帳時也要印
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      stDate = Left(ChangeTStringToWString(txt1(0)), 6) - 1
      stDate2 = Left(ChangeTStringToWString(txt1(0)), 6)
      m_str = "SELECT DECODE(SC03,NULL,'0','XX','1','2') 異動原因,ST02 姓名,SUBSTR(SD06,1,4)||'-'||SUBSTR(SD06,5,2)||'-'||SUBSTR(SD06,7,6)||'-'||SUBSTR(SD06,13,1) 帳號, " & _
              "nvl(YB05,0)+nvl(YB06,0)+nvl(YB26,0)+nvl(YB08,0)-nvl(YB15,0)-nvl(YB16,0)-nvl(YB17,0)-nvl(YB25,0) 轉帳金額,'2' SM37 " & _
              "FROM YEARBONUS,STAFF,SALARYDATA," & _
              "(SELECT SC01,SC03 FROM STAFF_CHANGE WHERE SC03 IN ('03','08','09','10') " & m_StrSQL2 & _
              " union select sd01 sc01,'XX' sc03 from salarydata,bookrecord where sd46>=" & stDate & " and sd46<=" & stDate2 & " and br01(+)=sd46 and br02 is null) C " & _
              "WHERE substr(YB02,1,2)||replace(substr(YB02,3,1),'A','0')||substr(YB02,4)=ST01(+)" & _
              " AND substr(YB02,1,2)||replace(substr(YB02,3,1),'A','0')||substr(YB02,4)=SD01(+)" & _
              " AND '2'=SD05 AND SD02<>'S' AND C.SC01=YB02(+)" & m_StrSQL1 & _
              " ORDER BY SM37,異動原因,帳號,YB02"
              
   '2013/3/21 ADD BY SONIA
   ElseIf txt1(1) = "3" Or txt1(1) = "4" Then  '3端午,4中秋
      strYM = Left(ChangeTStringToWString(txt1(0)), 4)

      If txt1(1) = "3" Then
         m_StrSQL1 = " and ob02='1' AND substr(ob01,1,4)='" & strYM & "'"     '端午
      Else
         m_StrSQL1 = " and ob02='2' AND substr(ob01,1,4)='" & strYM & "'"     '中秋
      End If
      
      'Add by Sindy 2023/12/1 離職人員隔月出清冊(1號離職還是當月出)
      m_StrSQL2 = m_StrSQL2 & " and SC02>'" & CompDate(1, -1, Left(txt1(0), Len(txt1(0)) - 2) & "01") & "' AND SC02<='" & CStr(Left(txt1(0), Len(txt1(0)) - 2) + 191100) & "01' "
      
      'modify by sonia 2018/9/27 +'1' SM37
      'modify by sonia 2020/5/12 '1' SM37->'2' SM37
      'Modify By Sindy 2023/12/1 資料類別 2年終獎金 3端午代金 4中秋代金 都要產出瑞興的"員工異動清冊"
      '原程式='0' 異動原因
      'Modified by Morgan 2024/2/21 新進有設當月或前月開始轉帳且薪資未入帳時也要印--婉莘
      stDate = Left(ChangeTStringToWString(txt1(0)), 6) - 1
      stDate2 = Left(ChangeTStringToWString(txt1(0)), 6)
      m_str = "SELECT DECODE(SC03,NULL,'0','XX','1','2') 異動原因,ST02 姓名,SUBSTR(SD06,1,4)||'-'||SUBSTR(SD06,5,2)||'-'||SUBSTR(SD06,7,6)||'-'||SUBSTR(SD06,13,1) 帳號, " & _
              "nvl(ob05,0) 轉帳金額,'2' SM37 " & _
              "FROM ohbonus,Staff,SalaryData," & _
              "(SELECT SC01,SC03 FROM STAFF_CHANGE WHERE SC03 IN ('03','08','09','10') " & m_StrSQL2 & _
              " union select sd01 sc01,'XX' sc03 from salarydata,bookrecord where sd46>=" & stDate & " and sd46<=" & stDate2 & " and br01(+)=sd46 and br02 is null) C " & _
              "WHERE ob05>0 AND ob03=st01(+) AND ob03=SD01(+) AND SD05='2'" & _
              " AND C.SC01=ob03(+)" & m_StrSQL1 & _
              " ORDER BY SM37,異動原因,帳號,OB03"
   '2013/3/21 END
   End If
   
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      With m_rs
         m_rs.MoveFirst
         
         '預設值
         iLine = 1
         strType = "0" '切頁條件
         strType1 = "" '2009/3/5 add by sonia
         'modify by sonia 2020/5/4
         'strSM37 = "1"  'add by sonia 2018/9/25
         strSM37 = "2"
         'end 2020/5/4
         dblAmtChange = 0: dblCntChange = 0
         
         Do While Not m_rs.EOF
             
            For m_i = 1 To 5
                strTemp(m_i) = ""
            Next m_i
             
            strTemp(1) = CheckStr(m_rs.Fields(0))   '異動原因
            strTemp(2) = CheckStr(m_rs.Fields(1))   '姓名
            strTemp(3) = CheckStr(m_rs.Fields(2))   '帳號
            strTemp(4) = CheckStr(m_rs.Fields(3))   '轉帳金額
            strTemp(5) = CheckStr(m_rs.Fields("SM37"))   '公司    add by sonia 2018/9/25
            
            'modify by sonia 2018/9/25 +判斷Or strSM37 <> strTemp(5)
            If iLine > 50 Or iLine = 1 Or strType <> strTemp(1) Or strSM37 <> strTemp(5) Then
               'modify by sonia 2018/9/25 +判斷Or strSM37 <> strTemp(5)
               If strType <> "0" And (strType <> strTemp(1) Or strSM37 <> strTemp(5)) Then
                  PrintEnd '小計
               End If

               If iLine <> 1 Then
                  PrintTail  '列印表尾
                  Printer.NewPage
               End If
               
               iLine = 1
               'modify by sonia 2018/9/25 +判斷Or strSM37 <> strTemp(5)
               If (strType <> strTemp(1) Or strSM37 <> strTemp(5)) Then
                  PrintTitle '列印表頭
               End If
            End If
            
            If strTemp(1) <> "0" Then   '0表無異動,不印異動清冊
               '2009/3/5 modify by sonia 辜說同一人只要印一筆即可
               'modify by sonia 2018/9/25 +判斷Or strSM37 <> strTemp(5)
               If strType <> strTemp(1) Or strType1 <> strTemp(3) Or strSM37 <> strTemp(5) Then
                  PrintDetail '列印表中
                  '異動小計
                  dblCntChange = dblCntChange + 1
                  dblAmtChange = dblAmtChange + strTemp(4)
                  strType1 = strTemp(3)
               End If
               '2009/3/5 end
            End If
            
            strType = strTemp(1) '依異動原因跳頁
            strSM37 = strTemp(5) 'add by sonia 2018/9/25
            m_rs.MoveNext
         Loop
          
         '列印表尾
         If strType <> "0" Then
            PrintEnd    '小計
            PrintTail   '列印表尾
            Printer.NewPage
         End If
         
'Removed by Morgan 2022/3/3 移到下面(沒有異動也要印遞送單)
'         '轉帳遞送單(北所)
'         PrintTotal
'
'         'modify by sonia 2017/3/28 高所由彰化銀行改改合作金庫(沒有固定格式)
'         'PrintKaohsiung   '2013/3/15 add by sonia 加高所媒體遞送單
'         PrintKaohsiungNew
'end 2022/3/3
           
      End With
   Else
      'Modified by Morgan 2022/3/3
      'MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
      'Exit Sub
      MsgBox "無異動資料!!!", vbExclamation + vbOKOnly
      'end 2022/3/3
   End If
   
   'Added by Morgan 2022/3/3
   '轉帳遞送單(北所)
   PrintTotal
   
   'modify by sonia 2017/3/28 高所由彰化銀行改改合作金庫(沒有固定格式)
   'PrintKaohsiung   '2013/3/15 add by sonia 加高所媒體遞送單
   'Modify By Sindy 202/9/4 mark,不用產出此報表
   'PrintKaohsiungNew
   'end 2022/3/3
   
   Printer.EndDoc
   ShowPrintOk

End Sub

Sub PrintTail()
   iLine = iLine + 5
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "委託單位簽章："
'cancel by sonia 2018/7/4
'   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("轉帳日期：" & m_Date1) - 1500
'   Printer.CurrentY = iLine * 300
'   Printer.Print "轉帳日期：" & m_Date1
'end 2018/7/4

End Sub

Sub PrintEnd()
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(130, "-")

   iLine = iLine + 1
   Printer.CurrentX = PLeft(1) - Printer.TextWidth(strTemp(2))
   Printer.CurrentY = iLine * 300
   Printer.Print "合　計："
   Printer.CurrentX = 6000
   Printer.CurrentY = iLine * 300
   Printer.Print dblCntChange & "人"
'2009/3/5 CANCEL BY SONIA 辜說不必印金額
'   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblAmtChange, "#,###,###,###"))
'   Printer.CurrentY = iLine * 300
'   Printer.Print Format(dblAmtChange, "#,###,###,###")
'2009/3/5 end

   dblAmtChange = 0: dblCntChange = 0

End Sub

'轉帳遞送單
Sub PrintTotal()
Dim adoacc080 As New ADODB.Recordset
   
   '抓1公司
   adoacc080.CursorLocation = adUseClient
   '2009/4/9 MODIFY BY SONIA 重新計算,以免當月到職當月離職會重覆計算 98006
   'adoacc080.Open "select * from acc080 where a0801 = '1'", adoTaie, adOpenStatic, adLockReadOnly
   If txt1(1) = "1" Then   '每月薪資
      '2009/5/26 modify by sonia 5月調薪且合併者勞退自6月才合併,故5月第二家薪資無薪資但有勞退,此類資料合併至第一家入帳
      'adoacc080.Open "select A0802,A0807,COUNT(*),SUM(nvl(SM04,0)+nvl(SM05,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)) " & _
                     "from acc080,SALARYMONTH,STAFF,SALARYDATA where a0801 = '1' AND REPLACE(SM01,'A','0')=ST01(+) AND REPLACE(SM01,'A','0')=SD01(+) AND '2'=SD05 " & m_StrSQL1 & _
                     "GROUP BY A0802,A0807 ", adoTaie, adOpenStatic, adLockReadOnly
      '先抓第二家若為<0者合併至第一家計算,第二家只抓>0的資料
      'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
      'Modified by Morgan 2013/2/1 +sm43
      'modify by sonia 2018/9/26 +SM37,A0825,A0826 北所台一投資及台一智權產生不同媒體檔案
      'modify by sonia 2020/5/4 取消1公司加法律所L,故直接用SM37
      'Modify By Sindy 2020/6/25 + 證照津貼
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      adoacc080.Open "SELECT A0802,A0825,A0826,COUNT(*) CNT,SUM(T2) AMT,SM37 FROM acc080,STAFF,SALARYDATA,( " & _
                     "SELECT SM01,SUM(T2) T2,SM37 FROM ( " & _
                     "select SM01,nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0) T2,SM37 from SALARYMONTH " & _
                     "where SUBSTR(SM01,3,1)<>'A' " & m_StrSQL1 & " UNION " & _
                     "select substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4),nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0) T2,SM37 from SALARYMONTH " & _
                     "where SUBSTR(SM01,3,1)='A' AND nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0)<=0 " & m_StrSQL1 & ") GROUP BY SM01,SM37 UNION " & _
                     "select SM01,nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0) T2,SM37 from SALARYMONTH " & _
                     "where SUBSTR(SM01,3,1)='A' and nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0)>0 " & m_StrSQL1 & _
                     ") X WHERE a0801 = SM37 AND substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)=ST01(+) AND substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)=SD01(+) AND '2'=SD05 " & _
                     "GROUP BY A0802,A0825,A0826,SM37 ORDER BY SM37,A0802,A0825,A0826 ", adoTaie, adOpenStatic, adLockReadOnly
      'end 2020/5/4
      '2009/5/26 END
   ElseIf txt1(1) = "2" Then  '年終
      'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
      '2013/1/23 modify by sonia 加扣補充保費yb25
      'modify by sonia 2016/1/7 留職停薪人員改發現金,不入媒體,故加SD02<>'S'
      'modify by sonia 2018/1/11 +YB26
      'modify by sonia 2018/9/26 +YB24,A0825,A0826 北所台一投資及台一智權產生不同媒體檔案
      'modify by sonia 2020/5/12 取消1公司加法律所L,故直接用YB24
      'adoacc080.Open "select A0802,A0825,A0826,COUNT(*) CNT,SUM(nvl(YB05,0)+nvl(YB06,0)+nvl(YB26,0)+nvl(YB08,0)-nvl(YB15,0)-nvl(YB16,0)-nvl(YB17,0)-nvl(YB25,0)) AMT,DECODE(YB24,'A',YB24,'J',YB24,'1') YB24 " & _
                     "from acc080,YEARBONUS,STAFF,SALARYDATA where a0801 = DECODE(SD19,'A',SD19,'J',SD19,'1') AND substr(YB02,1,1)||replace(substr(YB02,2),'A','0')=ST01(+) AND substr(YB02,1,1)||replace(substr(YB02,2),'A','0')=ST01(+) AND substr(YB02,1,1)||replace(substr(YB02,2),'A','0')=SD01(+) AND '2'=SD05 AND SD02<>'S' " & m_StrSQL1 & _
                     "GROUP BY A0802,A0825,A0826,DECODE(YB24,'A',YB24,'J',YB24,'1') ORDER BY DECODE(YB24,'A',YB24,'J',YB24,'1'),A0802,A0825,A0826 ", adoTaie, adOpenStatic, adLockReadOnly
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      adoacc080.Open "select A0802,A0825,A0826,COUNT(*) CNT,SUM(nvl(YB05,0)+nvl(YB06,0)+nvl(YB26,0)+nvl(YB08,0)-nvl(YB15,0)-nvl(YB16,0)-nvl(YB17,0)-nvl(YB25,0)) AMT,YB24 " & _
                     "from acc080,YEARBONUS,STAFF,SALARYDATA where a0801 = SD19 AND substr(YB02,1,2)||replace(substr(YB02,3,1),'A','0')||substr(YB02,4)=ST01(+) AND substr(YB02,1,2)||replace(substr(YB02,3,1),'A','0')||substr(YB02,4)=ST01(+) AND substr(YB02,1,2)||replace(substr(YB02,3,1),'A','0')||substr(YB02,4)=SD01(+) AND '2'=SD05 AND SD02<>'S' " & m_StrSQL1 & _
                     "GROUP BY A0802,A0825,A0826,YB24 ORDER BY YB24,A0802,A0825,A0826 ", adoTaie, adOpenStatic, adLockReadOnly
   '2013/3/21 ADD BY SONIA
   ElseIf txt1(1) = "3" Or txt1(1) = "4" Then  '3端午,4中秋
      'modify by sonia 2018/9/26 +OB12,A0825,A0826 北所台一投資及台一智權產生不同媒體檔案
      'modify by sonia 2020/5/12 取消1公司加法律所L,故直接用OB12
      'adoacc080.Open "SELECT A0802,A0825,A0826,COUNT(*) CNT,SUM(nvl(ob05,0)) AMT,DECODE(OB12,'A',OB12,'J',OB12,'1') SD19 " & _
                     "FROM ohbonus,acc080,Staff,SalaryData WHERE a0801 = DECODE(OB12,'A',OB12,'J',OB12,'1') AND ob05>0 AND ob03=st01(+) AND ob03=SD01(+) AND SD05='2' " & m_StrSQL1 & _
                     "GROUP BY A0802,A0825,A0826,DECODE(OB12,'A',OB12,'J',OB12,'1') ORDER BY DECODE(OB12,'A',OB12,'J',OB12,'1'),A0802,A0825,A0826 ", adoTaie, adOpenStatic, adLockReadOnly
      adoacc080.Open "SELECT A0802,A0825,A0826,COUNT(*) CNT,SUM(nvl(ob05,0)) AMT,OB12 " & _
                     "FROM ohbonus,acc080,Staff,SalaryData WHERE a0801 = OB12 AND ob05>0 AND ob03=st01(+) AND ob03=SD01(+) AND SD05='2' " & m_StrSQL1 & _
                     "GROUP BY A0802,A0825,A0826,OB12 ORDER BY OB12,A0802,A0825,A0826 ", adoTaie, adOpenStatic, adLockReadOnly
   '2013/3/21 END
   End If
   
   If Not adoacc080.EOF And Not adoacc080.BOF Then
      With adoacc080
         adoacc080.MoveFirst
         
         Do While Not adoacc080.EOF
            
            dblCntTotal = adoacc080.Fields("CNT").Value
            dblAmtTotal = adoacc080.Fields("AMT").Value
            
            Printer.Font.Size = 14
            Printer.Font.Underline = False
            Printer.FontBold = True
            
            iLine = 2
            Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("薪資集中轉帳遞送單") / 2)
            Printer.CurrentY = iLine * 300
            Printer.Print "薪資集中轉帳遞送單"
            
            Printer.Font.Size = 12
            Printer.FontBold = False
            iLine = iLine + 3
            Printer.CurrentX = 1000
            Printer.CurrentY = iLine * 300
            Printer.Print "委託　單位：　" & adoacc080.Fields("a0802").Value
            
            iLine = iLine + 2
            Printer.CurrentX = 1000
            Printer.CurrentY = iLine * 300
            Printer.Print "企業　編號：　" & adoacc080.Fields("a0825").Value
            
            iLine = iLine + 2
            Printer.CurrentX = 1000
            Printer.CurrentY = iLine * 300
            Printer.Print "委託單帳號：　" & adoacc080.Fields("a0826").Value
            
            iLine = iLine + 2
            Printer.CurrentX = 1000
            Printer.CurrentY = iLine * 300
            Printer.Print "轉帳　摘要：　" & IIf(txt1(1) = "1", "■薪資　　　　　　　　　　□獎金", "□薪資　　　　　　　　　　■獎金")
            
            iLine = iLine + 2
            Printer.CurrentX = 1000
            Printer.CurrentY = iLine * 300
            Printer.Print "轉帳　日期：　" & m_Date1
            
            iLine = iLine + 2
            Printer.CurrentX = 1000
            Printer.CurrentY = iLine * 300
            Printer.Print "轉帳筆金額：　" & Format(dblCntTotal, "#,###") & " 筆 " & Format(dblAmtTotal, "#,###,###,###") & " 元"
            
            iLine = iLine + 2
            Printer.CurrentX = 1000
            Printer.CurrentY = iLine * 300
            Printer.Print "支付　憑證：　■取款條　　　　　　　　　□支票號碼"
            
            iLine = iLine + 2
            Printer.CurrentX = 1000
            Printer.CurrentY = iLine * 300
            Printer.Print "資料　型式：　■磁片１片　　　　　　　　□清冊　　頁"
            
            iLine = iLine + 2
            Printer.CurrentX = 1000
            Printer.CurrentY = iLine * 300
            Printer.Print "委託　日期：　" & m_Date2
            
            iLine = iLine + 2
            Printer.CurrentX = 1000
            Printer.CurrentY = iLine * 300
            Printer.Print "委託單位印鑑："
            
            iLine = iLine + 3
            Printer.CurrentX = 500
            Printer.CurrentY = iLine * 300
            Printer.Print String(130, "-")
         
            iLine = iLine + 2
            Printer.Font.Size = 14
            Printer.FontBold = True
            Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("薪資集中轉帳遞送單") / 2)
            Printer.CurrentY = iLine * 300
            Printer.Print "收　妥　回　條"
            
            Printer.Font.Size = 12
            Printer.FontBold = False
            iLine = iLine + 2
            Printer.CurrentX = 1000
            Printer.CurrentY = iLine * 300
            Printer.Print "收妥　日期：　　　　年　　　月　　　日"
            
            iLine = iLine + 2
            Printer.CurrentX = 1000
            Printer.CurrentY = iLine * 300
            Printer.Print "委託　單位：　" & adoacc080.Fields("a0802").Value
          
            iLine = iLine + 2
            Printer.CurrentX = 1000
            Printer.CurrentY = iLine * 300
            Printer.Print "轉帳　摘要：　□薪資　　　　　　　　　　□獎金"
            
            iLine = iLine + 2
            Printer.CurrentX = 1000
            Printer.CurrentY = iLine * 300
            Printer.Print "轉帳　日期：　　　　年　　　月　　　日"
            
            iLine = iLine + 2
            Printer.CurrentX = 1000
            Printer.CurrentY = iLine * 300
            Printer.Print "轉帳筆金額：　　　　筆　　　　　　　　　　元　□取款條□支票號碼"
            
            iLine = iLine + 2
            Printer.CurrentX = 1000
            Printer.CurrentY = iLine * 300
            Printer.Print "資料　型式：　□磁片　　片　　　　　　　□清冊　　頁"
            
            iLine = iLine + 4
            Printer.Font.Underline = True
            Printer.CurrentX = 1000
            Printer.CurrentY = iLine * 300
            Printer.Print "瑞興商業銀行　　　　　　　　分行"
            Printer.Font.Underline = False
         
            iLine = iLine + 3
            Printer.CurrentX = 5000
            Printer.CurrentY = iLine * 300
            Printer.Print "主管：　　　　　　　　　　　經辦："
            
            Printer.NewPage
            
            adoacc080.MoveNext
         Loop
      End With
   End If
   adoacc080.Close
   
End Sub

Sub PrintTitle()

   GetPleft
   
   Printer.Font.Size = 12
   Printer.Font.Underline = False
   Printer.FontBold = False
   
   'modify by sonia 2018/9/26 +加公司名稱
   'modify by sonia 2020/5/4 +L公司
   'modify by sonia 2021/2/25 台一投資改名稱
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(IIf(strTemp(5) = "A", "臺一投資", IIf(strTemp(5) = "J", "台一智權", IIf(strTemp(5) = "L", "台一法律", "台一國際"))) & "委託薪資集中轉帳員工清冊") / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print IIf(strTemp(5) = "A", "臺一投資", IIf(strTemp(5) = "J", "台一智權", IIf(strTemp(5) = "L", "台一法律", "台一國際"))) & "委託薪資集中轉帳員工清冊"
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "頁　　次：" & Printer.Page
   
   iLine = iLine + 2
   Printer.CurrentX = PLeft(1) - Printer.TextWidth("姓　名")
   Printer.CurrentY = iLine * 300
   Printer.Print "姓　名"
   Printer.CurrentX = PLeft(2) - Printer.TextWidth("薪　資　轉　入　帳　號　　")
   Printer.CurrentY = iLine * 300
   Printer.Print "薪　資　轉　入　帳　號　　"
   Printer.CurrentX = PLeft(3) - Printer.TextWidth("金　　額")
   Printer.CurrentY = iLine * 300
   Printer.Print "金　　額"
   Printer.CurrentX = PLeft(4) - Printer.TextWidth("備註")
   Printer.CurrentY = iLine * 300
   Printer.Print "備註"
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(130, "-")
   
   iLine = iLine + 1
End Sub

Sub GetPleft()
   PLeft(1) = 2500
   PLeft(2) = 6500
   PLeft(3) = 8500
   PLeft(4) = 9500
End Sub

Sub PrintDetail()
   Printer.CurrentX = PLeft(1) - Printer.TextWidth(strTemp(2))
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(2)
   Printer.CurrentX = 3500
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(3)
'2009/3/5 CANCEL BY SONIA 辜說不必印金額
'   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(strTemp(4), "#,###,###,###"))
'   Printer.CurrentY = iLine * 300
'   Printer.Print Format(strTemp(4), "#,###,###,###")
'2009/3/5 END
   Printer.CurrentX = PLeft(4) - Printer.TextWidth("新進")
   Printer.CurrentY = iLine * 300
   If strTemp(1) = "1" Then
      Printer.Print "新進"
   Else
      Printer.Print "離職"
   End If
   
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
   Set frm170230 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         If txt1(Index) <> "" Then
             If ChkDate(txt1(Index)) = False Then
                Call txt1_GotFocus(Index)
                Cancel = False
                Exit Sub
             End If
             If ChkWork(ChangeTStringToWString(txt1(Index))) = False Then
                Call txt1_GotFocus(Index)
                Cancel = False
                Exit Sub
             End If
         End If
      Case 1
         If txt1(Index) <> "" Then
             If txt1(Index) <> "1" And txt1(Index) <> "2" And txt1(Index) <> "3" And txt1(Index) <> "4" Then
                MsgBox "資料類別輸入錯誤！", vbInformation, "操作錯誤！"
                Call txt1_GotFocus(Index)
                Cancel = False
                Exit Sub
             End If
         End If
      Case Else
   End Select
   
   If Cancel = True Then TextInverse txt1(Index)
      
End Sub

'add by sonia 2017/3/28 由彰銀改合作金庫(沒有固定格式)
'高所媒體遞送單-合作金庫
Sub PrintKaohsiungNew()
Dim m_Date3 As String

   '先抓明細資料計算 資料核證總數
   '抓1公司(專利商標),'6'=SD05(高所薪資)
   'modify by sonia 2020/5/4 a0801 = '1'改用a0801 = '2'
   If txt1(1) = "1" Then      '每月薪資
      'Modify By Sindy 2020/6/25 + 證照津貼
      m_str = "SELECT A0802,ST02,ST26,SM01,SD06,nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0) T2 " & _
               "FROM acc080,STAFF,SALARYDATA,SALARYMONTH " & _
               "WHERE a0801 = '2' AND sm01=ST01(+) AND sm01=SD01(+) AND '6'=SD05 " & m_StrSQL1
   ElseIf txt1(1) = "2" Then  '年終
      'modify by sonia 2018/1/11 +YB26
      m_str = "select A0802,ST02,ST26,YB02,SD06,nvl(YB05,0)+nvl(YB06,0)+nvl(YB26,0)+nvl(YB08,0)-nvl(YB15,0)-nvl(YB16,0)-nvl(YB17,0)-nvl(YB25,0) T2 " & _
               "from acc080,STAFF,SALARYDATA,YEARBONUS where a0801 = '2' AND YB02=ST01(+) AND YB02=SD01(+) AND '6'=SD05 " & m_StrSQL1
   ElseIf txt1(1) = "3" Or txt1(1) = "4" Then  '3端午,4中秋
      m_str = "SELECT A0802,ST02,ST26,OB03,SD06,nvl(OB05,0) T2 " & _
              "FROM acc080,STAFF,SALARYDATA,OHBONUS WHERE a0801 = '2' AND OB05>0 AND OB03=st01(+) AND OB03=SD01(+) AND SD05='6' " & m_StrSQL1
   End If
   
   m_str = m_str & " ORDER BY SD06"
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      
      '1公司(專利商標)名稱
      dblA0802 = "" & CheckStr(m_rs.Fields("A0802"))
      dblAmtTotal = 0: dblCntTotal = 0:
   
      iLine = 1
      PrintTitleKaohsiung '列印表頭
      
      With m_rs
         m_rs.MoveFirst
   
         Do While Not m_rs.EOF
            '帳號SD06,戶名ST02,身份證字號ST26,金額T2
            Printer.CurrentX = TPLeft(1)
            Printer.CurrentY = iLine * 300
            Printer.Print CheckStr(m_rs.Fields("SD06"))
            '姓　　名
            Printer.CurrentX = TPLeft(2)
            Printer.CurrentY = iLine * 300
            Printer.Print CheckStr(m_rs.Fields("ST02"))
            '身份證字號
            Printer.CurrentX = TPLeft(3)
            Printer.CurrentY = iLine * 300
            Printer.Print CheckStr(m_rs.Fields("ST26"))
            '金　額
            Printer.CurrentX = TPLeft(4) - Printer.TextWidth(Format(CheckStr(m_rs.Fields("T2")), "##,##0"))
            Printer.CurrentY = iLine * 300
            Printer.Print Format(CheckStr(m_rs.Fields("T2")), "##,##0")
            iLine = iLine + 1
            
            '累計筆數
            dblCntTotal = dblCntTotal + 1
            '累計金額
            dblAmtTotal = dblAmtTotal + Val(CheckStr(m_rs.Fields("T2")))
            
            m_rs.MoveNext
         Loop
         
         PrintTailKaohsiung  '列印表尾
            
      End With
   Else
      MsgBox "高所無符合列印的資料!!!", vbExclamation + vbOKOnly
      Exit Sub
   End If
   
End Sub

'add by sonia 2017/3/31 高所合庫加印明細
Sub PrintTitleKaohsiung()
   GetPleftKaohsiung
   
   Printer.Font.Size = 12
   Printer.Font.Underline = False
   Printer.FontBold = True
   
   iLine = iLine + 2
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(dblA0802 & "(高所)薪資轉帳明細") / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print dblA0802 & "(高所)薪資轉帳明細"
   
   Printer.FontBold = False
   iLine = iLine + 2
   Printer.CurrentX = TPLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "預定撥帳日期：" & m_Date1
   
   iLine = iLine + 2
   Printer.CurrentX = TPLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "帳　　號"
   Printer.CurrentX = TPLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print "姓　名"
   Printer.CurrentX = TPLeft(3)
   Printer.CurrentY = iLine * 300
   Printer.Print "身份證字號"
   Printer.CurrentX = TPLeft(4) - Printer.TextWidth("金　額")
   Printer.CurrentY = iLine * 300
   Printer.Print "金　額"
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(140, "-")
   
   iLine = iLine + 1
End Sub

Sub GetPleftKaohsiung()
   TPLeft(1) = 1000
   TPLeft(2) = 3500
   TPLeft(3) = 5500
   TPLeft(4) = 8500
End Sub

Sub PrintTailKaohsiung()

   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(140, "-")
   
   iLine = iLine + 1
   Printer.CurrentX = TPLeft(3)
   Printer.CurrentY = iLine * 300
   Printer.Print "合計：(" & dblCntTotal & "人)"
   Printer.CurrentX = TPLeft(4) - Printer.TextWidth(Format(dblAmtTotal, "##,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtTotal, "##,##0")

   '回條
   Printer.Font.Size = 12
   iLine = 30
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(140, "-")

   iLine = iLine + 2
   Printer.Font.Size = 14
   Printer.FontBold = True
   Printer.CurrentX = 5000
   Printer.CurrentY = iLine * 300
   Printer.Print "收　妥　回　條"
   
   Printer.Font.Size = 12
   Printer.FontBold = False
   iLine = iLine + 2
   Printer.CurrentX = 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "收　妥　日　期：　　　　年　　　月　　　日"
   
   iLine = iLine + 2
   Printer.CurrentX = 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "送　件　企　業：" & dblA0802
 
   iLine = iLine + 2
   Printer.CurrentX = 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "資　料　類　別：" & IIf(txt1(1) = "1", "■薪資　　　　　　　　　　□獎金", "□薪資　　　　　　　　　　■獎金")
   
   iLine = iLine + 2
   Printer.CurrentX = 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "預定撥帳日期　：" & m_Date1
   
   iLine = iLine + 2
   Printer.CurrentX = 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "資料筆數及金額：" & Format(dblCntTotal, "###") & " 筆，" & Format(dblAmtTotal, "$#,###,###,###") & " 元 "
   
   iLine = iLine + 2
   Printer.CurrentX = 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "資　料　型　式：隨身碟　1　個"
   
   iLine = iLine + 3
   Printer.CurrentX = 6000
   Printer.CurrentY = iLine * 300
   Printer.Print "收件人："

   Printer.EndDoc
End Sub
'end 2017/3/31

'cancel by sonia 2017/3/28
''2013/3/15 add by sonia
''高所媒體遞送單-彰化銀行
'Sub PrintKaohsiung()
'Dim dblMaxAmt As Double, dblMaxAmtAccount As Double, dblAccountTot As Double
'Dim AccountTotal As String, dblA0802 As String, dblA0807 As String
'
'   '先抓明細資料計算 資料核證總數
'   '抓1公司(專利商標),'6'=SD05(高所薪資)
'   If txt1(1) = "1" Then      '每月薪資
'      m_str = "SELECT A0802,A0807,SM01,SD06,nvl(SM04,0)+nvl(SM05,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0) T2 " & _
'               "FROM acc080,STAFF,SALARYDATA,SALARYMONTH " & _
'               "WHERE a0801 = '1' AND sm01=ST01(+) AND sm01=SD01(+) AND '6'=SD05 " & m_StrSQL1
'   ElseIf txt1(1) = "2" Then  '年終
'      m_str = "select A0802,A0807,YB02,SD06,nvl(YB05,0)+nvl(YB06,0)+nvl(YB08,0)-nvl(YB15,0)-nvl(YB16,0)-nvl(YB17,0)-nvl(YB25,0) T2 " & _
'               "from acc080,STAFF,SALARYDATA,YEARBONUS where a0801 = '1' AND YB02=ST01(+) AND YB02=SD01(+) AND '6'=SD05 " & m_StrSQL1
'   '2013/3/21 ADD BY SONIA
'   ElseIf txt1(1) = "3" Or txt1(1) = "4" Then  '3端午,4中秋
'      m_str = "SELECT A0802,A0807,OB03,SD06,nvl(OB05,0) T2 " & _
'              "FROM acc080,STAFF,SALARYDATA,OHBONUS WHERE a0801 = '1' AND OB05>0 AND OB03=st01(+) AND OB03=SD01(+) AND SD05='6' " & m_StrSQL1
'   '2013/3/21 END
'   End If
'
'   m_str = m_str & " ORDER BY T2 DESC,SD01"
'   If m_rs.State = 1 Then m_rs.Close
'   m_rs.CursorLocation = adUseClient
'   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
'   If Not m_rs.EOF And Not m_rs.BOF Then
'      With m_rs
'         m_rs.MoveFirst
'
'         dblAmtTotal = 0: dblCntTotal = 0: dblAccountTot = 0
'         '最大金額
'         dblMaxAmt = Val(CheckStr(m_rs.Fields("T2")))
'         '最大金額之帳號(取第7碼開始之6碼)
'         dblMaxAmtAccount = Val(Mid("" & CheckStr(m_rs.Fields("SD06")), 7, 6))
'         '1公司(專利商標)名稱
'         dblA0802 = "" & CheckStr(m_rs.Fields("A0802"))
'         '1公司(專利商標)統一編號
'         dblA0807 = "" & CheckStr(m_rs.Fields("A0807"))
'
'         Do While Not m_rs.EOF
'            '累計筆數
'            dblCntTotal = dblCntTotal + 1
'            '累計金額
'            dblAmtTotal = dblAmtTotal + Val(CheckStr(m_rs.Fields("T2")))
'            '累計帳號(每一帳號第7碼開始之6碼)
'            dblAccountTot = dblAccountTot + Val(Mid("" & CheckStr(m_rs.Fields("SD06")), 7, 6))
'            m_rs.MoveNext
'         Loop
'      End With
'   Else
'      MsgBox "高所無符合列印的資料!!!", vbExclamation + vbOKOnly
'      Exit Sub
'   End If
'
'   '計算資料核證總數=(預定轉撥日期(西元年月日)+最大金額之帳號(取第7碼開始之6碼)+最大金額+累計帳號(每一帳號第7碼開始之6碼))取右邊10碼
'   dblAccountTot = Val(DBDATE(txt1(0))) + Val(dblMaxAmtAccount) + Val(dblMaxAmt) + dblAccountTot
'   AccountTotal = Right(dblAccountTot, 10)
'
'   Printer.NewPage
'
'   '開始列印
'   Printer.Font.Size = 14
'   Printer.Font.Underline = False
'   Printer.FontBold = False
'
'   iLine = 2
'   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("彰化商業銀行　受託代理撥帳資料媒體遞送單") / 2)
'   Printer.CurrentY = iLine * 300
'   Printer.Print "彰化商業銀行　受託代理撥帳資料媒體遞送單"
'
'   iLine = iLine + 2
'   Printer.CurrentX = 1000
'   Printer.CurrentY = iLine * 300
'   Printer.Print "送件企業　" & "057F" & "　" & dblA0802
'
'   Printer.Font.Size = 12
'   Printer.CurrentX = 8000
'   Printer.CurrentY = iLine * 300
'   Printer.Print "送件日期：" & ChangeTStringToTDateString(txt1(0))
'
'   iLine = iLine + 2
'   '外框
'   Printer.Line (1000, iLine * 300)-(10000, iLine * 300 + 7000), , B
'   '直線
'   Printer.Line (1700, iLine * 300)-(1700, iLine * 300 + 7000)                '1
'   Printer.Line (2400, (iLine + 5) * 300)-(2400, (iLine + 5) * 300 + 3000)    '2
'   Printer.Line (5300, (iLine + 5) * 300)-(5300, (iLine + 5) * 300 + 2400)    '3上段
'   Printer.Line (5300, (iLine + 18) * 300)-(5300, (iLine + 18) * 300 + 1600)  '3下段
'   Printer.Line (6000, (iLine + 18) * 300)-(6000, (iLine + 18) * 300 + 1600)  '4(下)
'   Printer.Line (6300, (iLine + 5) * 300)-(6300, (iLine + 5) * 300 + 2400)    '5
'   '橫線
'   Printer.Line (1000, (iLine + 3) * 300)-(10000, (iLine + 3) * 300)          '1
'   Printer.Line (1700, (iLine + 5) * 300)-(10000, (iLine + 5) * 300)          '2
'   Printer.Line (1700, (iLine + 7) * 300)-(10000, (iLine + 7) * 300)          '3
'   Printer.Line (1700, (iLine + 9) * 300)-(10000, (iLine + 9) * 300)          '4
'   Printer.Line (1700, (iLine + 11) * 300)-(10000, (iLine + 11) * 300)        '5
'   Printer.Line (1700, (iLine + 13) * 300)-(10000, (iLine + 13) * 300)        '6
'   Printer.Line (1000, (iLine + 15) * 300)-(10000, (iLine + 15) * 300)        '7
'   Printer.Line (1000, (iLine + 18) * 300)-(10000, (iLine + 18) * 300)        '8
'
'   '資料類別
'   Printer.Font.Size = 8
'   Printer.CurrentX = 1250
'   Printer.CurrentY = iLine * 310 + (200 * 0)
'   Printer.Print "資"
'   Printer.CurrentX = 1250
'   Printer.CurrentY = iLine * 310 + (200 * 1)
'   Printer.Print "料"
'   Printer.CurrentX = 1250
'   Printer.CurrentY = iLine * 310 + (200 * 2)
'   Printer.Print "類"
'   Printer.CurrentX = 1250
'   Printer.CurrentY = iLine * 310 + (200 * 3)
'   Printer.Print "別"
'
'   iLine = iLine + 1
'   Printer.Font.Size = 14
'   Printer.CurrentX = 1900
'   Printer.CurrentY = iLine * 300
'   Printer.Print IIf(txt1(1) = "1", "( ) 51　　薪金入帳", "( ) 97 　獎金入帳")
'
'
'   '預定撥帳日期"   m_Date1
'   iLine = iLine + 3
'   Printer.Font.Size = 12
'   Printer.CurrentX = 1900
'   Printer.CurrentY = iLine * 300 - 100
'   Printer.Print "預定撥帳日期 " & m_Date1
'
'   iLine = iLine + 1
'   Printer.Font.Size = 8
'   Printer.CurrentX = 1250
'   Printer.CurrentY = iLine * 310 + (200 * 0)
'   Printer.Print "資　　　媒 體"
'   Printer.CurrentX = 5500
'   Printer.CurrentY = iLine * 310 + (200 * 0)
'   Printer.Print "存  款"
'   Printer.CurrentX = 1250
'   Printer.CurrentY = iLine * 310 + (200 * 1)
'   Printer.Print "　　　　種 類"
'   Printer.CurrentX = 5500
'   Printer.CurrentY = iLine * 310 + (200 * 1)
'   Printer.Print "總金額"
'
'   '存款總金額
'   iLine = iLine + 1
'   Printer.Font.Size = 12
'   Printer.CurrentX = 8500
'   Printer.CurrentY = iLine * 300 - 100
'   Printer.Print Format(dblAmtTotal, "$#,###,###,###") & " 元"
'
'   iLine = iLine + 1
'   Printer.Font.Size = 8
'   Printer.CurrentX = 1250
'   Printer.CurrentY = iLine * 310 + (200 * 0)
'   Printer.Print "料　　　媒 體"
'   Printer.CurrentX = 5500
'   Printer.CurrentY = iLine * 310 + (200 * 0)
'   Printer.Print "提  款"
'   Printer.CurrentX = 1250
'   Printer.CurrentY = iLine * 310 + (200 * 1)
'   Printer.Print "　　　　數 量"
'   Printer.CurrentX = 5500
'   Printer.CurrentY = iLine * 310 + (200 * 1)
'   Printer.Print "總金額"
'
'   '媒體數量
'   iLine = iLine + 1
'   Printer.Font.Size = 14
'   Printer.CurrentX = 3500
'   Printer.CurrentY = iLine * 300 - 100
'   Printer.Print " １"
'   Printer.Font.Size = 12
'   Printer.CurrentX = 9350
'   Printer.CurrentY = iLine * 300 - 100
'   Printer.Print " 元"
'
'   iLine = iLine + 1
'   Printer.Font.Size = 8
'   Printer.CurrentX = 1250
'   Printer.CurrentY = iLine * 310 + (200 * 0)
'   Printer.Print "　　　　資 料"
'   Printer.CurrentX = 5500
'   Printer.CurrentY = iLine * 310 + (200 * 0)
'   Printer.Print "資  　料"
'   Printer.CurrentX = 1250
'   Printer.CurrentY = iLine * 310 + (200 * 1)
'   Printer.Print "內　　　格 式"
'   Printer.CurrentX = 5500
'   Printer.CurrentY = iLine * 310 + (200 * 1)
'   Printer.Print "核證總數"
'
'   '資料核證總數"
'   iLine = iLine + 1
'   Printer.Font.Size = 12
'   Printer.CurrentX = 3000
'   Printer.CurrentY = iLine * 300 - 100
'   Printer.Print "(   ) A  (   ) B"
'   Printer.CurrentX = 8450
'   Printer.CurrentY = iLine * 300 - 100
'   Printer.Print AccountTotal
'
'   iLine = iLine + 1
'   Printer.Font.Size = 8
'   Printer.CurrentX = 1250
'   Printer.CurrentY = iLine * 310 + (200 * 0)
'   Printer.Print "　　　　資 料"
'   Printer.CurrentX = 5500
'   Printer.CurrentY = iLine * 310 + (200 * 0)
'   Printer.Print "企  　業"
'   Printer.CurrentX = 1250
'   Printer.CurrentY = iLine * 310 + (200 * 1)
'   Printer.Print "容　　　筆 數"
'   Printer.CurrentX = 5500
'   Printer.CurrentY = iLine * 310 + (200 * 1)
'   Printer.Print "統一編號"
'
'   '資料筆數
'   iLine = iLine + 1
'   Printer.Font.Size = 12
'   Printer.CurrentX = 4000
'   Printer.CurrentY = iLine * 300 - 100
'   Printer.Print dblCntTotal & " 件"
'
'   '企業統編
'   Printer.CurrentX = 7000
'   Printer.CurrentY = iLine * 300 - 100
'   Printer.Print dblA0807
'
'   '磁片規格
'   iLine = iLine + 1
'   Printer.Font.Size = 8
'   Printer.CurrentX = 1250
'   Printer.CurrentY = iLine * 305 + (200 * 0)
'   Printer.Print "　　　　磁 片　　1. MS-DOS磁片；5.25 INCH；1.44MB 或 1.2MB"
'   Printer.CurrentX = 1250
'   Printer.CurrentY = iLine * 305 + (200 * 1)
'   Printer.Print "　　　　規 格　　2. ASCII CODE"
'
'   '備註
'   iLine = iLine + 2
'   Printer.Font.Size = 8
'   Printer.CurrentX = 1250
'   Printer.CurrentY = iLine * 305 + (200 * 0)
'   Printer.Print "備　　　1. 本單由委託企業填妥連同轉帳資料媒體，付款憑證送交受託分行"
'   Printer.CurrentX = 1250
'   Printer.CurrentY = iLine * 305 + (200 * 1)
'   Printer.Print "　　　　2. 採郵遞方式處理時，受理分行章後，將付款憑證留存，遞送單與資料媒"
'   Printer.CurrentX = 1250
'   Printer.CurrentY = iLine * 305 + (200 * 2)
'   Printer.Print "註　　 　   體轉送資料訊室"
'
'   '簽章
'   iLine = iLine + 3
'   Printer.Font.Size = 8
'   Printer.CurrentX = 1200
'   Printer.CurrentY = iLine * 310 + (200 * 0)
'   Printer.Print "受責"
'   Printer.CurrentX = 5500
'   Printer.CurrentY = iLine * 310 + (200 * 0)
'   Printer.Print "委責"
'   Printer.CurrentX = 1200
'   Printer.CurrentY = iLine * 310 + (200 * 1)
'   Printer.Print "託人"
'   Printer.CurrentX = 5500
'   Printer.CurrentY = iLine * 310 + (200 * 1)
'   Printer.Print "託人"
'   Printer.CurrentX = 1200
'   Printer.CurrentY = iLine * 310 + (200 * 2)
'   Printer.Print "分收"
'   Printer.CurrentX = 5500
'   Printer.CurrentY = iLine * 310 + (200 * 2)
'   Printer.Print "企送"
'   Printer.CurrentX = 1200
'   Printer.CurrentY = iLine * 310 + (200 * 3)
'   Printer.Print "行件"
'   Printer.CurrentX = 5500
'   Printer.CurrentY = iLine * 310 + (200 * 3)
'   Printer.Print "業件"
'   Printer.CurrentX = 1200
'   Printer.CurrentY = iLine * 310 + (200 * 4)
'   Printer.Print "及簽"
'   Printer.CurrentX = 5500
'   Printer.CurrentY = iLine * 310 + (200 * 4)
'   Printer.Print "及簽"
'   Printer.CurrentX = 1200
'   Printer.CurrentY = iLine * 310 + (200 * 5)
'   Printer.Print "負章"
'   Printer.CurrentX = 5500
'   Printer.CurrentY = iLine * 310 + (200 * 5)
'   Printer.Print "負章"
'
'   '回條
'   Printer.Font.Size = 12
'   iLine = 30
'   Printer.CurrentX = 500
'   Printer.CurrentY = iLine * 300
'   Printer.Print String(140, "-")
'
'   iLine = iLine + 2
'   Printer.Font.Size = 14
'   Printer.FontBold = True
'   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("彰化商業銀行 受託代理撥帳資料媒體遞送單") / 2)
'   Printer.CurrentY = iLine * 300
'   Printer.Print "收　妥　回　條"
'
'   Printer.Font.Size = 12
'   Printer.FontBold = False
'   iLine = iLine + 2
'   Printer.CurrentX = 1000
'   Printer.CurrentY = iLine * 300
'   Printer.Print "收妥　日期：　　　　年　　　月　　　日"
'
'   iLine = iLine + 2
'   Printer.CurrentX = 1000
'   Printer.CurrentY = iLine * 300
'   Printer.Print "送件　企業：" & dblA0802
'
'   iLine = iLine + 2
'   Printer.CurrentX = 1000
'   Printer.CurrentY = iLine * 300
'   Printer.Print "資料　類別：" & IIf(txt1(1) = "1", "■薪資　　　　　　　　　　□獎金", "□薪資　　　　　　　　　　■獎金")
'
'   iLine = iLine + 2
'   Printer.CurrentX = 1000
'   Printer.CurrentY = iLine * 300
'   Printer.Print "送件　日期：" & m_Date1
'
'   iLine = iLine + 2
'   Printer.CurrentX = 1000
'   Printer.CurrentY = iLine * 300
'   Printer.Print "轉帳筆數及金額：" & Format(dblCntTotal, "###") & " 筆，" & Format(dblAmtTotal, "$#,###,###,###") & " 元 "
'
'   iLine = iLine + 2
'   Printer.CurrentX = 1000
'   Printer.CurrentY = iLine * 300
'   Printer.Print "資料　型式：隨身碟　1　個"
'
'   iLine = iLine + 4
'   Printer.Font.Underline = True
'   Printer.CurrentX = 1000
'   Printer.CurrentY = iLine * 300
'   Printer.Print "彰化商業銀行 　　　　　　　　分行"
'   Printer.Font.Underline = False
'
'   iLine = iLine + 3
'   Printer.CurrentX = 5000
'   Printer.CurrentY = iLine * 300
'   Printer.Print "收件人："
'
'End Sub
'end 2017/3/28
