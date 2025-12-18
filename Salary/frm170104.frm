VERSION 5.00
Begin VB.Form frm170104 
   BorderStyle     =   1  '單線固定
   Caption         =   "銀行入帳媒體作業"
   ClientHeight    =   3444
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5052
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3444
   ScaleWidth      =   5052
   Begin VB.CheckBox Check1 
      Caption         =   "南所"
      Height          =   225
      Index           =   1
      Left            =   4200
      TabIndex        =   3
      Top             =   1770
      Value           =   1  '核取
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   4896
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "Y"
      Top             =   1470
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   1
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "1"
      Top             =   1020
      Width           =   300
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   3750
      TabIndex        =   5
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "轉檔(&O)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2730
      TabIndex        =   4
      Top             =   90
      Width           =   975
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   0
      Left            =   1320
      MaxLength       =   7
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "　　中所檔名為salary中所中國信託.txt。"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   5
      Left            =   210
      TabIndex        =   16
      Top             =   2700
      Width           =   3160
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "　　南所檔名為salary南所台北富邦.txt。"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   4
      Left            =   190
      TabIndex        =   15
      Top             =   2910
      Width           =   3160
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(email通知同仁日期)"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   1
      Left            =   2400
      TabIndex        =   14
      Top             =   660
      Width           =   1590
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "　　高所檔名為salary高所合作金庫.txt。"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   3
      Left            =   190
      TabIndex        =   13
      Top             =   3120
      Width           =   3160
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "       　　salary、salary投資、salary智權。"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   2
      Left            =   190
      TabIndex        =   12
      Top             =   2490
      Width           =   3230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "　　北所依公司別產生三個媒體檔案，檔名分別為："
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   1
      Left            =   190
      TabIndex        =   11
      Top             =   2280
      Width           =   4140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否發通知給南所：         ( Y: 是 )"
      Height          =   180
      Index           =   4
      Left            =   3240
      TabIndex        =   10
      Top             =   1520
      Visible         =   0   'False
      Width           =   2630
   End
   Begin VB.Label Label4 
      Caption         =   "(1: 每月薪資 2 : 年終獎金   3: 端午代金  4: 中秋代金)"
      Height          =   360
      Left            =   1920
      TabIndex        =   9
      Top             =   1080
      Width           =   1995
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PS：在桌面產生資料夾名稱為YYY年MM月薪資轉檔"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   190
      TabIndex        =   8
      Top             =   2060
      Width           =   4080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "資料類別："
      Height          =   180
      Left            =   390
      TabIndex        =   7
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "入帳日期："
      Height          =   180
      Index           =   0
      Left            =   390
      TabIndex        =   6
      Top             =   660
      Width           =   900
   End
End
Attribute VB_Name = "frm170104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/2/17 Form2.0已檢查 (無需修改的物件)
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Create by SINDY 2008/12/30
'2009/2/4 modify by sonia 加入端午,中秋代金
Option Explicit

Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 10) As Integer
Dim strTemp(1 To 18) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim strYM As String  '2013/3/6 自strmenu1移過來
Dim strFolder As String 'Added by Morgan 2013/7/30


Private Sub cmdok_Click(Index As Integer)
'2013/3/6 add by sonia
Dim strReceiver As String
Dim strType As String
'2013/3/6 end

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
         
         Screen.MousePointer = vbHourglass
         
         strType = ""    '2013/3/6 add by sonia
         Select Case txt1(1)
            Case "1"       '每月薪資
               StrMenu1
               strType = Mid(strYM, 1, 4) & "/" & Mid(strYM, 5, 2) & "月薪資" '2013/3/6 add by sonia
            Case "2"       '年終獎金
               StrMenu2
               strType = "年終獎金"   '2013/3/6 add by sonia
            '2009/2/4 add by sonia
            Case "3", "4"  '端午,中秋代金
               StrMenu3
               '2013/3/6 add by sonia
               If txt1(1) = "3" Then
                  strType = "端午節代金"
               Else
                  strType = "中秋節代金"
               End If
               '2013/3/6 end
            '2009/2/4 end
         End Select
         
         'Modify By Sindy 2023/8/9 最後連南所也由北所統一出轉帳媒體檔
'         '2013/3/5 add by sonia e-mail給分所
'         If txt1(2) = "Y" Then
'            strReceiver = ""
'            '2013/7/29 CANCEL BY SONIA 中所由台北直接產生寄中所交中國信託
'            'If Check1(0).Value = "1" Then   '中所
'            '   strReceiver = strReceiver & "85003;"
'            'End If
'            '2013/7/29 END
'            If Check1(1).Value = "1" Then   '南所
'               strReceiver = strReceiver & "71002;"
'            End If
'            '2022/6/21 CANCEL BY Sindy 高所由台北直接產生寄高所交合作金庫
''            If Check1(2).Value = "1" Then   '高所
''               strReceiver = strReceiver & "68008;"
''            End If
'            '若收件人請假,不發職代
'            Call PUB_SendMail(strUserNum, strReceiver, "", "薪資媒體資料已產生, 請至分所財務系統處理 !", "這次是 " & strType & " 資料, 請至分所財務系統->一般作業->薪資媒體作業 處理 !", "謝謝 ! 日安 !", , , , , , , , , True)
'         End If
'         '2013/3/5 end
         
         Screen.MousePointer = vbDefault
      Case 1
         Unload Me
   End Select
End Sub

Sub StrMenu1()
Dim ff As Integer
Dim strText As String
Dim TempFileName As String
Dim dblTotCnt As Double
Dim dblTotAmt As Double
Dim intLen As Integer
Dim strSM37 As String  'add by sonia 2018/9/21
Dim strA0825 As String 'add by sonia 2018/9/27
   
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
   'end 2022/1/27
   
   'Add By Sindy 2023/9/25
   strFolder = PUB_Getdesktop & "\" & Left(strYM, 4) - 1911 & "年" & Right(strYM, 2) & "月薪資轉檔"
   If Dir(strFolder, vbDirectory) = "" Then
      MkDir strFolder
   End If
   '2023/9/25 END
   
   '2009/5/26 modify by sonia 5月調薪且合併者勞退自6月才合併,故5月第二家薪資無薪資但有勞退,此類資料合併至第一家入帳
   'm_str = "SELECT SD01,ST02,SD05,SD06,T.*,a0802 " & _
            "FROM Staff,SalaryData,( " & _
            "SELECT replace(SM01,'A','0') as T1,nvl(SM04,0)+nvl(SM05,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0) as T2,SM37 as T3,SM01 as T4 " & _
            "From SalaryMonth " & _
            "WHERE SM02='" & strYM & "') T,acc080 " & _
            "Where ST01 = t1 AND SD01=T1 AND T3=a0801(+) " & _
            "AND SD05='2' " & m_StrSQL
   '先抓第二家若為<0者合併至第一家計算,第二家只抓>0的資料
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'modify by sonia 2018/9/21 +DECODE(T3,'A',T3,'J',T3,'1') SM37公司別,A0825企業編號 北所台一投資及台一智權產生不同媒體檔案
   'modify by sonia 2020/5/4 +A0820,DECODE(T3,'A',T3,'J',T3,'1') SM37改直接抓T3
   'Modify By Sindy 2020/6/25 + 證照津貼
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   m_str = "SELECT SD01,ST02,SD05,SD06,T.*,A0802,A0825,A0820,T3 SM37 " & _
            "FROM Staff,SalaryData,( " & _
            "select t1,sum(t2) T2,s.SM37 as T3,s.SM01 as t4 from salarymonth s,(" & _
            "SELECT SM01 as T1,sm02,nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0) as T2 " & _
            "From SalaryMonth WHERE SUBSTR(SM01,3,1)<>'A' AND SM02=" & strYM & " UNION " & _
            "SELECT substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4) as T1,sm02,nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0) as T2 " & _
            "From SalaryMonth WHERE SUBSTR(SM01,3,1)='A' AND SM02=" & strYM & " " & _
            "and nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0)<=0 " & _
            ") X where x.t1=sm01 and x.sm02=s.sm02 group by T1,s.SM37 ,s.SM01 UNION " & _
            "SELECT substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4) as T1,nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0) as T2,SM37 as T3,SM01 as T4 " & _
            "From SalaryMonth WHERE SUBSTR(SM01,3,1)='A' AND SM02=" & strYM & " " & _
            "and nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0)>0 " & _
            ") T,acc080 Where ST01 = t1 AND SD01=T1 AND T3=a0801(+) AND SD05='2' Order by T3,SD01"
   '2009/5/26 end
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      With m_rs
         m_rs.MoveFirst
         
         If ff > 0 Then Close #ff
         
'以下程式modify by sonia 2018/9/21 北所台一投資及台一智權產生不同媒體檔案
         Do While Not m_rs.EOF
             
            If strSM37 <> m_rs.Fields("SM37") Then
            
               If strSM37 <> "" Then
                  '總計：最後一筆
                   For m_i = 1 To 13
                       strTemp(m_i) = ""
                   Next m_i
                   strTemp(1) = "3"
                   strTemp(2) = "1010075"
                   strTemp(3) = Right("0000000" & txt1(0), 7) '轉帳日期 9(07)
                   strTemp(4) = strA0825                      '企業編號  原固定50015, 2018/9/27改抓ACC080之A0825
                   strTemp(5) = "071" '071.薪水
                   strTemp(6) = Right("00000" & dblTotCnt, 5) '總件數 9(05)
                   strTemp(7) = Right("0000000000000" & dblTotAmt, 13) & "00" '總金額 9(13)V9(2)
                   strTemp(8) = String(57, "0") '全為0 X(57)
                   strText = ""
                   For m_i = 1 To 8
                      strText = strText & strTemp(m_i)
                   Next m_i
                   Print #ff, strText
                   Close ff
               End If
               
               strSM37 = m_rs.Fields("SM37")
               strA0825 = m_rs.Fields("A0825")
               dblTotAmt = 0
               dblTotCnt = 0
               ff = FreeFile
               '2009/1/5 modify by sonia c:\有權限問題,故改放桌面
               'TempFileName = "c:\Salary"
               'modify by sonia 2018/9/21 北所台一投資及台一智權產生不同媒體檔案
               'modify by sonia 2020/5/4 加法律所改抓A0820
               'TempFileName = PUB_Getdesktop & "\salary" & IIf(strSM37 = "A", "投資", IIf(strSM37 = "J", "智權", ""))  '2009/1/16銀行說檔名要全小寫
               'Modify By Sindy 2023/9/25
               'TempFileName = PUB_Getdesktop & "\salary" & IIf(strSM37 = "2", "", m_rs.Fields("A0820"))
               TempFileName = strFolder & "\salary" & IIf(strSM37 = "2", "", m_rs.Fields("A0820"))
               Open TempFileName For Output As ff
            End If
          
            For m_i = 1 To 13
               strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = "2"
            strTemp(2) = "1010075"
            strTemp(3) = Right("0000000" & txt1(0), 7) '轉帳日期 9(07)
            strTemp(4) = m_rs.Fields("A0825")          '企業編號  原固定50015, 2018/9/27改抓ACC080之A0825
            strTemp(5) = "071" '071.薪水
            '企業名+員工名 X(20)
            '2012/10/3 MODIFY BY SONIA 超過20碼截取20碼 (F5644  林信昌電機組)
            'intLen = GetTextLength(Left(CheckStr(m_rs.Fields("a0802")), 6) & CheckStr(m_rs.Fields("ST02")))
            'intLen = 20 - intLen
            'strTemp(6) = Left(CheckStr(m_rs.Fields("a0802")), 6) & CheckStr(m_rs.Fields("ST02")) & String(intLen, " ")
            intLen = GetTextLength(Left(Left(CheckStr(m_rs.Fields("A0802")), 6) & CheckStr(m_rs.Fields("ST02")), 10))
            intLen = 20 - intLen
            strTemp(6) = Left(Left(CheckStr(m_rs.Fields("A0802")), 6) & CheckStr(m_rs.Fields("ST02")), 10) & String(intLen, " ")
            '2012/10/3 END
            'modify by sonia 2018/10/5 因A7019之SD06為0181220942280，故改格式但最後一筆不改
            'strTemp(7) = "1010075"
            'strTemp(8) = "22"
            'strTemp(9) = Left(Mid(CheckStr(m_rs.Fields("SD06")), 7, 7) & "       ", 7) '轉帳帳號 9(07)
            strTemp(7) = "101"
            strTemp(8) = CheckStr(m_rs.Fields("SD06"))                                  '轉帳帳號 9(13)
            'end 2018/10/5
            strTemp(10) = Right("00000000000" & CheckStr(m_rs.Fields("T2")), 11) & "00" '轉帳金額 9(11)V9(2)
            strTemp(11) = "1"
            strTemp(12) = Left(CheckStr(m_rs.Fields("T4")) & "          ", 10) '員工代號 X(10)
            strTemp(13) = String(17, " ") '空白 X(17)
            
            '總件數
            dblTotCnt = dblTotCnt + 1
            '總金額
            dblTotAmt = dblTotAmt + CheckStr(m_rs.Fields("T2"))
            
            strText = ""
            For m_i = 1 To 13
               strText = strText & strTemp(m_i)
            Next m_i
            Print #ff, strText
            m_rs.MoveNext
         Loop
          
         '總計：最後一筆
         For m_i = 1 To 13
             strTemp(m_i) = ""
         Next m_i
         strTemp(1) = "3"
         strTemp(2) = "1010075"
         strTemp(3) = Right("0000000" & txt1(0), 7) '轉帳日期 9(07)
         strTemp(4) = strA0825                      '企業編號  原固定50015, 2018/9/27改抓ACC080之A0825
         strTemp(5) = "071" '071.薪水
         strTemp(6) = Right("00000" & dblTotCnt, 5) '總件數 9(05)
         strTemp(7) = Right("0000000000000" & dblTotAmt, 13) & "00" '總金額 9(13)V9(2)
         strTemp(8) = String(57, "0") '全為0 X(57)
         strText = ""
         For m_i = 1 To 8
            strText = strText & strTemp(m_i)
         Next m_i
         Print #ff, strText
         Close ff
         
      End With
      
      'Add by Morgan 2009/6/15
      If AddBookRecord(strYM, DBDATE(txt1(0))) = False Then
         MsgBox "新增薪資入帳記錄失敗，請通知電腦中心人員！", vbExclamation
      End If
      
      Call ProMiddleOfficeEFile("1", strYM) 'Modify By Sindy 2022/6/21 改成共用函數-產生中所媒體檔案
      Call ProSouthOfficeEFile("1", strYM) 'Add By Sindy 2023/8/9 產生南所媒體檔案
      Call ProHighOfficeEFile("1", strYM) 'Add By Sindy 2022/6/21 產生高所媒體檔案
'      'add by sonia 2013/7/30 產生中所媒體檔案
'      'Modify By Sindy 2020/6/25 + 證照津貼
'      m_str = "SELECT SD01,ST02,SD06,ST26,nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0) as T2 " & _
'              "FROM STAFF,SALARYDATA,SalaryMonth WHERE SM02=" & strYM & " AND SM01=ST01(+) AND SM01=SD01(+) AND SD05='4' Order by SD01"
'      If m_rs.State = 1 Then m_rs.Close
'      m_rs.CursorLocation = adUseClient
'      m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
'      If Not m_rs.EOF And Not m_rs.BOF Then
'         With m_rs
'            m_rs.MoveFirst
'            dblTotAmt = 0
'
'            If ff > 0 Then Close #ff
'            ff = FreeFile
'
'            TempFileName = PUB_Getdesktop & "\中國信託薪資轉帳\salary中所中國信託.txt"
'
'            'Added by Morgan 2013/7/30
'            strFolder = PUB_Getdesktop & "\中國信託薪資轉帳"
'            If Dir(strFolder, vbDirectory) = "" Then
'               MkDir strFolder
'            End If
'            'endif
'
'            Open TempFileName For Output As ff
'
'            'Modify By Sindy 2022/1/12 因中所薪資改網銀轉帳，系統產生的TXT檔格式要改
''            '首筆(第一筆)
''            strTemp(1) = "159540272896"                                                     '公司帳號      X(12)
''            strTemp(2) = Right("0000000000" & TAIWANDATE(txt1(0)), 10)                      '入帳日期      9(10) 民國年月日YYYMMDD(不足10碼前補0)
''            strTemp(3) = "A"                                                                '收付業務別    X(01) 固定為 'A'
''
''            strText = ""
''            For m_i = 1 To 3
''               strText = strText & strTemp(m_i)
''            Next m_i
''            Print #ff, strText
''
''            '明細(第二筆以後)
''            Do While Not m_rs.EOF
''
''               For m_i = 1 To 6
''                  strTemp(m_i) = ""
''               Next m_i
''
''               strTemp(1) = Left(CheckStr(m_rs.Fields("SD06")) & "            ", 12)       '入帳帳號       X(12)
''               strTemp(2) = Right("0000000000" & CheckStr(m_rs.Fields("T2")), 10)          '入帳金額       9(10)
''               strTemp(3) = "          "                                                   '收款人姓名     X(10) 中國信託建議放空白
''               strTemp(4) = Left(CheckStr(m_rs.Fields("ST26")) & "          ", 10)         '身份證字號     X(10)
''               strTemp(5) = " "                                                            '入帳種類代碼   X(01) 空白：薪資
''               strTemp(6) = "Y"                                                            '檢查功能碼     X(03) 檢查入帳戶之身份證字號
''
''               strText = ""
''               For m_i = 1 To 6
''                  strText = strText & strTemp(m_i)
''               Next m_i
''               Print #ff, strText
''               m_rs.MoveNext
''            Loop
'            '首筆(第一筆)
'            strTemp(1) = "H"
'            strTemp(2) = String(80, " ")
'            strTemp(3) = Left(CheckStr("應 發 項 目") & "                    ", 16)
'            strTemp(4) = Left(CheckStr("代 扣 項 目") & "                    ", 16)
'            strText = ""
'            For m_i = 1 To 4
'               strText = strText & strTemp(m_i)
'            Next m_i
'            strText = strText & String(879, " ")
'            Print #ff, strText
'
'            '明細(第二筆以後)
'            Do While Not m_rs.EOF
'
'               For m_i = 1 To 5
'                  strTemp(m_i) = ""
'               Next m_i
'
'               strTemp(1) = "2"                                                              '資料識別=2      X(1)  請放 “2”
'               strTemp(2) = Right("0000000000000000" & CheckStr(m_rs.Fields("SD06")), 16)     '帳號            X(16) 不足位數，右靠左補0
'               strTemp(3) = Right("00000000000000" & CheckStr(m_rs.Fields("T2")) & "00", 16) '金額            9(16) 右靠左補0,後面兩位為小數
'                                                                                             '                      EX: 金額為5000元，轉成文字檔為0000000000500000，如果金額為0，請放16個0
'               strTemp(4) = "822"                                                            '行庫代號=822    X(3)
'               strTemp(5) = Left(CheckStr(m_rs.Fields("ST26")) & "           ", 11)          '身份證號        X(11) 不足位數，左靠右補半形空白
'
'               strText = ""
'               For m_i = 1 To 5
'                  strText = strText & strTemp(m_i)
'               Next m_i
'               strText = strText & String(953, " ")
'               Print #ff, strText
'
'               m_rs.MoveNext
'            Loop
'
'         End With
'         Close ff
'
'         ZipFolder strFolder 'Added by Morgan 2013/7/30
'      Else
'         MsgBox "無符合條件的中所資料!!!", vbExclamation + vbOKOnly
'         Exit Sub
'      End If
'      'end 2013/7/30
      
      MsgBox "每月薪資媒體已完成!!!", vbExclamation + vbOKOnly
      Exit Sub
   Else
      MsgBox "無符合條件的資料!!!", vbExclamation + vbOKOnly
      Exit Sub
   End If
End Sub
'Added by Morgan 2013/7/30
Private Function ZipFolder(pFolder As String) As Boolean
   Dim program_name As String, program_path As String
   Dim process_id As Long
   Dim process_handle As Long

   program_name = "C:\Program Files\7-Zip\7z.exe"
   '檢查執行檔
   If Dir(program_name) = "" Then
      MsgBox "未安裝 7-Zip 程式，壓縮檔產生失敗！。"
      Exit Function
   End If
    
On Error GoTo ShellError
        
   '刪除舊檔
   If Dir(pFolder & ".zip") <> "" Then
      Kill pFolder & ".zip"
   End If
   
   process_id = Shell("""" & program_name & """ a -pCTCB """ & pFolder & ".zip"" """ & pFolder & "\*""", vbNormalNoFocus)
    
    On Error GoTo 0

    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If
    Exit Function

ShellError:
    MsgBox " " & _
        program_name & vbCrLf & _
        Err.Description, vbOKOnly Or vbExclamation, _
        "Error"
End Function

Sub StrMenu2()
Dim strYM As String
Dim ff As Integer
Dim strText As String
Dim TempFileName As String
Dim dblTotCnt As Double
Dim dblTotAmt As Double
Dim intLen As Integer
Dim strYB24 As String  'add by sonia 2018/9/25
Dim strA0825 As String 'add by sonia 2018/9/27

   strYM = Left(ChangeTStringToWString(txt1(0)), 4)
   strYM = strYM - 1
   
   'Add By Sindy 2023/9/25
   strFolder = PUB_Getdesktop & "\" & Val(strYM) - 1911 & "年年終獎金轉檔"
   If Dir(strFolder, vbDirectory) = "" Then
      MkDir strFolder
   End If
   '2023/9/25 END
   
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   '2013/1/23 modify by sonia 加扣補充保費yb25
   'modify by sonia 2016/1/7 留職停薪人員改發現金,不入媒體,故加SD02<>'S'
   'modify by sonia 2018/1/11 +YB26
   'modify by sonia 2018/9/25 +T3公司別,A0825企業編號 北所台一投資及台一智權產生不同媒體檔案
   'modify by sonia 2020/6/20 +A0820,DECODE(T3,'A',T3,'J',T3,'1') T3改直接抓YB24
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   m_str = "SELECT SD01,ST02,SD05,SD06,T.*,A0802,A0825,YB24,A0820 " & _
            "FROM Staff,SalaryData,( " & _
            "SELECT substr(YB02,1,2)||replace(substr(YB02,3,1),'A','0')||substr(YB02,4) as T1,nvl(YB05,0)+nvl(YB06,0)+nvl(YB26,0)+nvl(YB08,0)-nvl(YB15,0)-nvl(YB16,0)-nvl(YB17,0)-nvl(YB25,0) as T2,YB24,YB02 as T4 " & _
            "From YearBonus " & _
            "WHERE YB01='" & strYM & "') T,acc080 " & _
            "Where ST01 = t1 AND SD01=T1 AND YB24=a0801(+) " & _
            "AND SD05='2' AND SD02<>'S' Order by 10,SD01"
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      With m_rs
         m_rs.MoveFirst
         
         If ff > 0 Then Close #ff

'以下程式modify by sonia 2018/9/25 北所台一投資及台一智權產生不同媒體檔案
         Do While Not m_rs.EOF
             
            If strYB24 <> m_rs.Fields("YB24") Then  'T3改直接抓YB24
            
               If strYB24 <> "" Then
                  '總計：最後一筆
                  For m_i = 1 To 13
                      strTemp(m_i) = ""
                  Next m_i
                  strTemp(1) = "3"
                  strTemp(2) = "1010075"
                  strTemp(3) = Right("0000000" & txt1(0), 7) '轉帳日期 9(07)
                  strTemp(4) = strA0825                      '企業編號  原固定50015, 2018/9/27改抓ACC080之A0825
                  strTemp(5) = "074" '074.獎金
                  strTemp(6) = Right("00000" & dblTotCnt, 5) '總件數 9(05)
                  strTemp(7) = Right("0000000000000" & dblTotAmt, 13) & "00" '總金額 9(13)V9(2)
                  strTemp(8) = String(57, "0") '全為0 X(57)
                  strText = ""
                  For m_i = 1 To 8
                      strText = strText & strTemp(m_i)
                  Next m_i
                  Print #ff, strText
                  Close ff
               End If
               
               strYB24 = m_rs.Fields("YB24")  'T3改直接抓YB24
               strA0825 = m_rs.Fields("A0825")
               dblTotAmt = 0
               dblTotCnt = 0
               ff = FreeFile
               '2009/1/5 modify by sonia c:\有權限問題,故改放桌面
               'TempFileName = "c:\Salary"
               'modify by sonia 2018/9/25 北所台一投資及台一智權產生不同媒體檔案
               'modify by sonia 2020/5/4 加法律所改抓A0820
               'TempFileName = PUB_Getdesktop & "\salary" & IIf(strYB24 = "A", "投資", IIf(strYB24 = "J", "智權", ""))  '2009/1/16銀行說檔名要全小寫
               'Modify By Sindy 2023/9/25
               'TempFileName = PUB_Getdesktop & "\salary" & IIf(strYB24 = "2", "", m_rs.Fields("A0820"))
               TempFileName = strFolder & "\salary" & IIf(strYB24 = "2", "", m_rs.Fields("A0820"))
               Open TempFileName For Output As ff
            End If
            
            For m_i = 1 To 13
               strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = "2"
            strTemp(2) = "1010075"
            strTemp(3) = Right("0000000" & txt1(0), 7) '轉帳日期 9(07)
            strTemp(4) = m_rs.Fields("A0825")          '企業編號  原固定50015, 2018/9/27改抓ACC080之A0825
            strTemp(5) = "074" '074.獎金
            '企業名+員工名 X(20)
            intLen = GetTextLength(Left(CheckStr(m_rs.Fields("A0802")), 6) & CheckStr(m_rs.Fields("ST02")))
            intLen = 20 - intLen
            strTemp(6) = Left(CheckStr(m_rs.Fields("A0802")), 6) & CheckStr(m_rs.Fields("ST02")) & String(intLen, " ")
            'modify by sonia 2018/10/5 因A7019之SD06為0181220942280，故改格式但最後一筆不改
            'strTemp(7) = "1010075"
            'strTemp(8) = "22"
            'strTemp(9) = Left(Mid(CheckStr(m_rs.Fields("SD06")), 7, 7) & "       ", 7) '轉帳帳號 9(07)
            strTemp(7) = "101"
            strTemp(8) = CheckStr(m_rs.Fields("SD06"))                                  '轉帳帳號 9(13)
            'end 2018/10/5
            strTemp(10) = Right("00000000000" & CheckStr(m_rs.Fields("T2")), 11) & "00" '轉帳金額 9(11)V9(2)
            strTemp(11) = "1"
            strTemp(12) = Left(CheckStr(m_rs.Fields("T4")) & "          ", 10) '員工代號 X(10)
            strTemp(13) = String(17, " ") '空白 X(17)
            
            '總件數
            dblTotCnt = dblTotCnt + 1
            '總金額
            dblTotAmt = dblTotAmt + CheckStr(m_rs.Fields("T2"))
            
            strText = ""
            For m_i = 1 To 13
                strText = strText & strTemp(m_i)
            Next m_i
            Print #ff, strText
            m_rs.MoveNext
         Loop
          
         '總計：最後一筆
         For m_i = 1 To 13
             strTemp(m_i) = ""
         Next m_i
         strTemp(1) = "3"
         strTemp(2) = "1010075"
         strTemp(3) = Right("0000000" & txt1(0), 7) '轉帳日期 9(07)
         strTemp(4) = strA0825                      '企業編號  原固定50015, 2018/9/27改抓ACC080之A0825
         strTemp(5) = "071" '071.薪水
         strTemp(5) = "074" '074.獎金
         strTemp(6) = Right("00000" & dblTotCnt, 5) '總件數 9(05)
         strTemp(7) = Right("0000000000000" & dblTotAmt, 13) & "00" '總金額 9(13)V9(2)
         strTemp(8) = String(57, "0") '全為0 X(57)
         strText = ""
         For m_i = 1 To 8
             strText = strText & strTemp(m_i)
         Next m_i
         Print #ff, strText
         Close ff
          
      End With
      
      'add by sonia 2015/12/25
      If AddBookRecord(Val(strYM) + 1 & "13", DBDATE(txt1(0))) = False Then
         MsgBox "新增薪資入帳記錄失敗，請通知電腦中心人員！", vbExclamation
      End If
      'end 2015/12/25
      
      Call ProMiddleOfficeEFile("2", strYM) 'Modify By Sindy 2022/6/21 改成共用函數-產生中所媒體檔案
      Call ProSouthOfficeEFile("2", strYM) 'Add By Sindy 2023/8/9 產生南所媒體檔案
      Call ProHighOfficeEFile("2", strYM) 'Add By Sindy 2022/6/21 產生高所媒體檔案
'      'add by sonia 2013/7/30 產生中所媒體檔案
'      'modify by sonia 2016/1/7 留職停薪人員改發現金,不入媒體,故加SD02<>'S'
'      'modify by sonia 2018/1/11 +YB26
'      m_str = "SELECT SD01,ST02,SD06,ST26,nvl(YB05,0)+nvl(YB06,0)+nvl(YB26,0)+nvl(YB08,0)-nvl(YB15,0)-nvl(YB16,0)-nvl(YB17,0)-nvl(YB25,0) as T2 " & _
'              "FROM STAFF,SALARYDATA,YearBonus WHERE YB01='" & strYM & "' AND YB02=ST01(+) AND YB02=SD01(+) AND SD05='4' AND SD02<>'S' Order by SD01"
'      If m_rs.State = 1 Then m_rs.Close
'      m_rs.CursorLocation = adUseClient
'      m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
'      If Not m_rs.EOF And Not m_rs.BOF Then
'         With m_rs
'            m_rs.MoveFirst
'            dblTotAmt = 0
'
'            If ff > 0 Then Close #ff
'            ff = FreeFile
'            TempFileName = PUB_Getdesktop & "\中國信託薪資轉帳\salary中所中國信託.txt"
'
'            'Added by Morgan 2013/7/30
'            strFolder = PUB_Getdesktop & "\中國信託薪資轉帳"
'            If Dir(strFolder, vbDirectory) = "" Then
'               MkDir strFolder
'            End If
'            'endif
'
'            Open TempFileName For Output As ff
'
'            'Modify By Sindy 2022/1/12 因中所薪資改網銀轉帳，系統產生的TXT檔格式要改
''            '首筆(第一筆)
''            strTemp(1) = "159540272896"                                                     '公司帳號      X(12)
''            strTemp(2) = Right("0000000000" & TAIWANDATE(txt1(0)), 10)                      '入帳日期      9(10) 民國年月日YYYMMDD(不足10碼前補0)
''            strTemp(3) = "A"                                                                '收付業務別    X(01) 固定為 'A'
''
''            strText = ""
''            For m_i = 1 To 3
''               strText = strText & strTemp(m_i)
''            Next m_i
''            Print #ff, strText
''
''            '明細(第二筆以後)
''            Do While Not m_rs.EOF
''
''               For m_i = 1 To 6
''                  strTemp(m_i) = ""
''               Next m_i
''
''               strTemp(1) = Left(CheckStr(m_rs.Fields("SD06")) & "            ", 12)       '入帳帳號       X(12)
''               strTemp(2) = Right("0000000000" & CheckStr(m_rs.Fields("T2")), 10)          '入帳金額       9(10)
''               strTemp(3) = "          "                                                   '收款人姓名     X(10) 中國信託建議放空白
''               strTemp(4) = Left(CheckStr(m_rs.Fields("ST26")) & "          ", 10)         '身份證字號     X(10)
''               strTemp(5) = " "                                                            '入帳種類代碼   X(01) 空白：薪資
''               strTemp(6) = "Y"                                                            '檢查功能碼     X(03) 檢查入帳戶之身份證字號
''
''               strText = ""
''               For m_i = 1 To 6
''                  strText = strText & strTemp(m_i)
''               Next m_i
''               Print #ff, strText
''               m_rs.MoveNext
''            Loop
'            '首筆(第一筆)
'            strTemp(1) = "H"
'            strTemp(2) = String(80, " ")
'            strTemp(3) = Left(CheckStr("應 發 項 目") & "                    ", 16)
'            strTemp(4) = Left(CheckStr("代 扣 項 目") & "                    ", 16)
'            strText = ""
'            For m_i = 1 To 4
'               strText = strText & strTemp(m_i)
'            Next m_i
'            strText = strText & String(879, " ")
'            Print #ff, strText
'
'            '明細(第二筆以後)
'            Do While Not m_rs.EOF
'
'               For m_i = 1 To 5
'                  strTemp(m_i) = ""
'               Next m_i
'
'               strTemp(1) = "2"                                                              '資料識別=2      X(1)  請放 “2”
'               strTemp(2) = Right("0000000000000000" & CheckStr(m_rs.Fields("SD06")), 16)     '帳號            X(16) 不足位數，右靠左補0
'               strTemp(3) = Right("00000000000000" & CheckStr(m_rs.Fields("T2")) & "00", 16) '金額            9(16) 右靠左補0,後面兩位為小數
'                                                                                             '                      EX: 金額為5000元，轉成文字檔為0000000000500000，如果金額為0，請放16個0
'               strTemp(4) = "822"                                                            '行庫代號=822    X(3)
'               strTemp(5) = Left(CheckStr(m_rs.Fields("ST26")) & "           ", 11)          '身份證號        X(11) 不足位數，左靠右補半形空白
'
'               strText = ""
'               For m_i = 1 To 5
'                  strText = strText & strTemp(m_i)
'               Next m_i
'               strText = strText & String(953, " ")
'               Print #ff, strText
'
'               m_rs.MoveNext
'            Loop
'
'         End With
'         Close ff
'
'         ZipFolder strFolder 'Added by Morgan 2013/7/30
'      Else
'         MsgBox "無符合條件的中所資料!!!", vbExclamation + vbOKOnly
'         Exit Sub
'      End If
'      'end 2013/7/30
      
      MsgBox "年終獎金媒體已完成!!!", vbExclamation + vbOKOnly
      Exit Sub
   Else
      MsgBox "無符合條件的資料!!!", vbExclamation + vbOKOnly
      Exit Sub
   End If
End Sub

'2009/2/4 add by sonia 端午,中秋代金
Sub StrMenu3()
Dim strYM As String
Dim ff As Integer
Dim strText As String
Dim TempFileName As String
Dim dblTotCnt As Double
Dim dblTotAmt As Double
Dim intLen As Integer
Dim strOB12 As String  'add by sonia 2018/9/25
Dim strA0825 As String 'add by sonia 2018/9/27

   If txt1(1) = "3" Then
      m_StrSQL = " and ob02='1'"     '端午
   Else
      m_StrSQL = " and ob02='2'"     '中秋
   End If
   'modify by sonia 2018/9/25 +OB12 北所台一投資及台一智權產生不同媒體檔案
   'modify by sonia 2020/6/22 DECODE(OB12,'A',OB12,'J',OB12,'1')改直接抓OB12
   'm_StrSQL = m_StrSQL & " Order by DECODE(OB12,'A',OB12,'J',OB12,'1'),OB03 "
   m_StrSQL = m_StrSQL & " Order by OB12,OB03 "
   
   strYM = Left(ChangeTStringToWString(txt1(0)), 4)
   
   'Add By Sindy 2023/9/25
   strFolder = PUB_Getdesktop & "\" & Val(strYM) - 1911 & "年" & IIf(txt1(1) = "3", "端午", "中秋") & "代金轉檔"
   If Dir(strFolder, vbDirectory) = "" Then
      MkDir strFolder
   End If
   '2023/9/25 END
   
   '2009/9/30 MODIFY BY SONIA 加OB05>0  因為2009年96011於九月底離職
   'modify by sonia 2018/9/25 +OB12公司別,A0825企業編號 北所台一投資及台一智權產生不同媒體檔案
   'modify by sonia 2020/6/22 DECODE(OB12,'A',OB12,'J',OB12,'1')改直接抓OB12,+A0820
   m_str = "SELECT SD01,ST02,SD05,SD06,ob03 T1,nvl(ob05,0) T2,OB12 T3,ob03 T4,A0802,A0825,OB12,A0820 " & _
            "FROM Staff,SalaryData,ohbonus,acc080 " & _
            "WHERE substr(ob01,1,4)='" & strYM & "' and ob05>0 " & _
            "AND ob03=st01(+) AND ob03=SD01(+) AND OB12=a0801(+) " & _
            "AND SD05='2' " & m_StrSQL
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      With m_rs
         m_rs.MoveFirst
         
         If ff > 0 Then Close #ff
         
'以下程式modify by sonia 2018/9/25 北所台一投資及台一智權產生不同媒體檔案
         Do While Not m_rs.EOF
               
            If strOB12 <> m_rs.Fields("OB12") Then
            
               If strOB12 <> "" Then
                  '總計：最後一筆
                  For m_i = 1 To 13
                      strTemp(m_i) = ""
                  Next m_i
                  strTemp(1) = "3"
                  strTemp(2) = "1010075"
                  strTemp(3) = Right("0000000" & txt1(0), 7) '轉帳日期 9(07)
                  strTemp(4) = strA0825                      '企業編號  原固定50015, 2018/9/27改抓ACC080之A0825
                  strTemp(5) = "074" '074.獎金
                  strTemp(6) = Right("00000" & dblTotCnt, 5) '總件數 9(05)
                  strTemp(7) = Right("0000000000000" & dblTotAmt, 13) & "00" '總金額 9(13)V9(2)
                  strTemp(8) = String(57, "0") '全為0 X(57)
                  strText = ""
                  For m_i = 1 To 8
                      strText = strText & strTemp(m_i)
                  Next m_i
                  Print #ff, strText
                  Close ff
               End If
               
               strOB12 = m_rs.Fields("OB12")
               strA0825 = m_rs.Fields("A0825")
               dblTotAmt = 0
               dblTotCnt = 0
               ff = FreeFile
               'modify by sonia 2018/9/25 北所台一投資及台一智權產生不同媒體檔案
               'modify by sonia 2020/6/20 加法律所改抓A0820
               'TempFileName = PUB_Getdesktop & "\salary" & IIf(strOB12 = "A", "投資", IIf(strOB12 = "J", "智權", ""))
               'Modify By Sindy 2023/9/25
               'TempFileName = PUB_Getdesktop & "\salary" & IIf(strOB12 = "2", "", m_rs.Fields("A0820"))
               TempFileName = strFolder & "\salary" & IIf(strOB12 = "2", "", m_rs.Fields("A0820"))
               Open TempFileName For Output As ff
            End If
             
            For m_i = 1 To 13
               strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = "2"
            strTemp(2) = "1010075"
            strTemp(3) = Right("0000000" & txt1(0), 7) '轉帳日期 9(07)
            strTemp(4) = m_rs.Fields("A0825")          '企業編號  原固定50015, 2018/9/27改抓ACC080之A0825
            strTemp(5) = "074" '074.獎金
            '企業名+員工名 X(20)
            intLen = GetTextLength(Left(CheckStr(m_rs.Fields("a0802")), 6) & CheckStr(m_rs.Fields("ST02")))
            intLen = 20 - intLen
            strTemp(6) = Left(CheckStr(m_rs.Fields("a0802")), 6) & CheckStr(m_rs.Fields("ST02")) & String(intLen, " ")
            'modify by sonia 2018/10/5 因A7019之SD06為0181220942280，故改格式但最後一筆不改
            'strTemp(7) = "1010075"
            'strTemp(8) = "22"
            'strTemp(9) = Left(Mid(CheckStr(m_rs.Fields("SD06")), 7, 7) & "       ", 7) '轉帳帳號 9(07)
            strTemp(7) = "101"
            strTemp(8) = CheckStr(m_rs.Fields("SD06"))                                  '轉帳帳號 9(13)
            'end 2018/10/5
            strTemp(10) = Right("00000000000" & CheckStr(m_rs.Fields("T2")), 11) & "00" '轉帳金額 9(11)V9(2)
            strTemp(11) = "1"
            strTemp(12) = Left(CheckStr(m_rs.Fields("T4")) & "          ", 10) '員工代號 X(10)
            strTemp(13) = String(17, " ") '空白 X(17)
            
            '總件數
            dblTotCnt = dblTotCnt + 1
            '總金額
            dblTotAmt = dblTotAmt + CheckStr(m_rs.Fields("T2"))
            
            strText = ""
            For m_i = 1 To 13
                strText = strText & strTemp(m_i)
            Next m_i
            Print #ff, strText
            m_rs.MoveNext
         Loop
          
         '總計：最後一筆
         For m_i = 1 To 13
             strTemp(m_i) = ""
         Next m_i
         strTemp(1) = "3"
         strTemp(2) = "1010075"
         strTemp(3) = Right("0000000" & txt1(0), 7) '轉帳日期 9(07)
         strTemp(4) = strA0825                      '企業編號  原固定50015, 2018/9/27改抓ACC080之A0825
         strTemp(5) = "074" '074.獎金
         strTemp(6) = Right("00000" & dblTotCnt, 5) '總件數 9(05)
         strTemp(7) = Right("0000000000000" & dblTotAmt, 13) & "00" '總金額 9(13)V9(2)
         strTemp(8) = String(57, "0") '全為0 X(57)
         strText = ""
         For m_i = 1 To 8
             strText = strText & strTemp(m_i)
         Next m_i
         Print #ff, strText
         Close ff
           
      End With
      
      'add by sonia 2015/12/25
      If txt1(1) = "3" Then
         If AddBookRecord(strYM & "14", DBDATE(txt1(0))) = False Then
            MsgBox "新增薪資入帳記錄失敗，請通知電腦中心人員！", vbExclamation
         End If
      Else
         If AddBookRecord(strYM & "15", DBDATE(txt1(0))) = False Then
            MsgBox "新增薪資入帳記錄失敗，請通知電腦中心人員！", vbExclamation
         End If
      End If
      'end 2015/12/25
      
      Call ProMiddleOfficeEFile("3", strYM) 'Modify By Sindy 2022/6/21 改成共用函數-產生中所媒體檔案
      Call ProSouthOfficeEFile("3", strYM) 'Add By Sindy 2023/8/9 產生南所媒體檔案
      Call ProHighOfficeEFile("3", strYM) 'Add By Sindy 2022/6/21 產生高所媒體檔案
'      'add by sonia 2013/7/30 產生中所媒體檔案
'      m_str = "SELECT SD01,ST02,SD06,ST26,nvl(ob05,0) T2 FROM Staff,SalaryData,ohbonus " & _
'               "WHERE substr(ob01,1,4)='" & strYM & "' and ob05>0 AND ob03=st01(+) AND ob03=SD01(+) AND SD05='4' " & m_StrSQL
'      If m_rs.State = 1 Then m_rs.Close
'      m_rs.CursorLocation = adUseClient
'      m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
'      If Not m_rs.EOF And Not m_rs.BOF Then
'         With m_rs
'            m_rs.MoveFirst
'            dblTotAmt = 0
'
'            If ff > 0 Then Close #ff
'            ff = FreeFile
'            TempFileName = PUB_Getdesktop & "\中國信託薪資轉帳\salary中所中國信託.txt"
'
'            'Added by Morgan 2013/7/30
'            strFolder = PUB_Getdesktop & "\中國信託薪資轉帳"
'            If Dir(strFolder, vbDirectory) = "" Then
'               MkDir strFolder
'            End If
'            'endif
'
'            Open TempFileName For Output As ff
'
'            'Modify By Sindy 2022/1/12 因中所薪資改網銀轉帳，系統產生的TXT檔格式要改
''            '首筆(第一筆)
''            strTemp(1) = "159540272896"                                                     '公司帳號      X(12)
''            strTemp(2) = Right("0000000000" & TAIWANDATE(txt1(0)), 10)                      '入帳日期      9(10) 民國年月日YYYMMDD(不足10碼前補0)
''            strTemp(3) = "A"                                                                '收付業務別    X(01) 固定為 'A'
''
''            strText = ""
''            For m_i = 1 To 3
''               strText = strText & strTemp(m_i)
''            Next m_i
''            Print #ff, strText
''
''            '明細(第二筆以後)
''            Do While Not m_rs.EOF
''
''               For m_i = 1 To 6
''                  strTemp(m_i) = ""
''               Next m_i
''
''               strTemp(1) = Left(CheckStr(m_rs.Fields("SD06")) & "            ", 12)       '入帳帳號       X(12)
''               strTemp(2) = Right("0000000000" & CheckStr(m_rs.Fields("T2")), 10)          '入帳金額       9(10)
''               strTemp(3) = "          "                                                   '收款人姓名     X(10) 中國信託建議放空白
''               strTemp(4) = Left(CheckStr(m_rs.Fields("ST26")) & "          ", 10)         '身份證字號     X(10)
''               strTemp(5) = " "                                                            '入帳種類代碼   X(01) 空白：薪資
''               strTemp(6) = "Y"                                                            '檢查功能碼     X(03) 檢查入帳戶之身份證字號
''
''               strText = ""
''               For m_i = 1 To 6
''                  strText = strText & strTemp(m_i)
''               Next m_i
''               Print #ff, strText
''               m_rs.MoveNext
''            Loop
'            '首筆(第一筆)
'            strTemp(1) = "H"
'            strTemp(2) = String(80, " ")
'            strTemp(3) = Left(CheckStr("應 發 項 目") & "                    ", 16)
'            strTemp(4) = Left(CheckStr("代 扣 項 目") & "                    ", 16)
'            strText = ""
'            For m_i = 1 To 4
'               strText = strText & strTemp(m_i)
'            Next m_i
'            strText = strText & String(879, " ")
'            Print #ff, strText
'
'            '明細(第二筆以後)
'            Do While Not m_rs.EOF
'
'               For m_i = 1 To 5
'                  strTemp(m_i) = ""
'               Next m_i
'
'               strTemp(1) = "2"                                                              '資料識別=2      X(1)  請放 “2”
'               strTemp(2) = Right("0000000000000000" & CheckStr(m_rs.Fields("SD06")), 16)     '帳號            X(16) 不足位數，右靠左補0
'               strTemp(3) = Right("00000000000000" & CheckStr(m_rs.Fields("T2")) & "00", 16) '金額            9(16) 右靠左補0,後面兩位為小數
'                                                                                             '                      EX: 金額為5000元，轉成文字檔為0000000000500000，如果金額為0，請放16個0
'               strTemp(4) = "822"                                                            '行庫代號=822    X(3)
'               strTemp(5) = Left(CheckStr(m_rs.Fields("ST26")) & "           ", 11)          '身份證號        X(11) 不足位數，左靠右補半形空白
'
'               strText = ""
'               For m_i = 1 To 5
'                  strText = strText & strTemp(m_i)
'               Next m_i
'               strText = strText & String(953, " ")
'               Print #ff, strText
'
'               m_rs.MoveNext
'            Loop
'
'         End With
'         Close ff
'
'         ZipFolder strFolder 'Added by Morgan 2013/7/30
'      Else
'         MsgBox "無符合條件的中所資料!!!", vbExclamation + vbOKOnly
'         Exit Sub
'      End If
'      'end 2013/7/30
      
      If txt1(1) = "3" Then
         MsgBox "端午代金媒體已完成!!!", vbExclamation + vbOKOnly
      Else
         MsgBox "中秋代金媒體已完成!!!", vbExclamation + vbOKOnly
      End If
      Exit Sub
   Else
      MsgBox "無符合條件的資料!!!", vbExclamation + vbOKOnly
      Exit Sub
   End If
End Sub
'2009/2/4 END

'Modify By Sindy 2022/6/21 改成共用函數-產生中所媒體檔案 - 中國信託
Private Sub ProMiddleOfficeEFile(strType As String, strYM As String)
Dim ff As Integer
Dim strText As String
Dim TempFileName As String
Dim dblTotAmt As Double

   'add by sonia 2013/7/30 產生中所媒體檔案
   If strType = "1" Then
      'Modify By Sindy 2020/6/25 + 證照津貼
      m_str = "SELECT SD01,ST02,SD06,ST26,nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0) as T2 " & _
              "FROM STAFF,SALARYDATA,SalaryMonth WHERE SM02=" & strYM & " AND SM01=ST01(+) AND SM01=SD01(+) AND SD05='4' Order by SD01"
   ElseIf strType = "2" Then
      'modify by sonia 2016/1/7 留職停薪人員改發現金,不入媒體,故加SD02<>'S'
      'modify by sonia 2018/1/11 +YB26
      m_str = "SELECT SD01,ST02,SD06,ST26,nvl(YB05,0)+nvl(YB06,0)+nvl(YB26,0)+nvl(YB08,0)-nvl(YB15,0)-nvl(YB16,0)-nvl(YB17,0)-nvl(YB25,0) as T2 " & _
              "FROM STAFF,SALARYDATA,YearBonus WHERE YB01='" & strYM & "' AND YB02=ST01(+) AND YB02=SD01(+) AND SD05='4' AND SD02<>'S' Order by SD01"
   Else
      m_str = "SELECT SD01,ST02,SD06,ST26,nvl(ob05,0) T2 FROM Staff,SalaryData,ohbonus " & _
               "WHERE substr(ob01,1,4)='" & strYM & "' and ob05>0 AND ob03=st01(+) AND ob03=SD01(+) AND SD05='4' " & m_StrSQL
   End If
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      With m_rs
         m_rs.MoveFirst
         dblTotAmt = 0

         If ff > 0 Then Close #ff
         ff = FreeFile
         
         'Modify By Sindy 2023/9/25
         'TempFileName = PUB_Getdesktop & "\中國信託薪資轉帳\salary中所中國信託.txt"
         TempFileName = strFolder & "\salary中所中國信託.txt"
'         'Added by Morgan 2013/7/30
'         strFolder = PUB_Getdesktop & "\中國信託薪資轉帳"
'         If Dir(strFolder, vbDirectory) = "" Then
'            MkDir strFolder
'         End If
'         'endif
         '2023/9/25 END
         
         Open TempFileName For Output As ff
         
         'Modify By Sindy 2022/1/12 因中所薪資改網銀轉帳，系統產生的TXT檔格式要改
'            '首筆(第一筆)
'            strTemp(1) = "159540272896"                                                     '公司帳號      X(12)
'            strTemp(2) = Right("0000000000" & TAIWANDATE(txt1(0)), 10)                      '入帳日期      9(10) 民國年月日YYYMMDD(不足10碼前補0)
'            strTemp(3) = "A"                                                                '收付業務別    X(01) 固定為 'A'
'
'            strText = ""
'            For m_i = 1 To 3
'               strText = strText & strTemp(m_i)
'            Next m_i
'            Print #ff, strText
'
'            '明細(第二筆以後)
'            Do While Not m_rs.EOF
'
'               For m_i = 1 To 6
'                  strTemp(m_i) = ""
'               Next m_i
'
'               strTemp(1) = Left(CheckStr(m_rs.Fields("SD06")) & "            ", 12)       '入帳帳號       X(12)
'               strTemp(2) = Right("0000000000" & CheckStr(m_rs.Fields("T2")), 10)          '入帳金額       9(10)
'               strTemp(3) = "          "                                                   '收款人姓名     X(10) 中國信託建議放空白
'               strTemp(4) = Left(CheckStr(m_rs.Fields("ST26")) & "          ", 10)         '身份證字號     X(10)
'               strTemp(5) = " "                                                            '入帳種類代碼   X(01) 空白：薪資
'               strTemp(6) = "Y"                                                            '檢查功能碼     X(03) 檢查入帳戶之身份證字號
'
'               strText = ""
'               For m_i = 1 To 6
'                  strText = strText & strTemp(m_i)
'               Next m_i
'               Print #ff, strText
'               m_rs.MoveNext
'            Loop
         '首筆(第一筆)
         strTemp(1) = "H"
         strTemp(2) = String(80, " ")
         strTemp(3) = Left(CheckStr("應 發 項 目") & "                    ", 16)
         strTemp(4) = Left(CheckStr("代 扣 項 目") & "                    ", 16)
         strText = ""
         For m_i = 1 To 4
            strText = strText & strTemp(m_i)
         Next m_i
         strText = strText & String(879, " ")
         Print #ff, strText

         '明細(第二筆以後)
         Do While Not m_rs.EOF

            For m_i = 1 To 5
               strTemp(m_i) = ""
            Next m_i

            strTemp(1) = "2"                                                              '資料識別=2      X(1)  請放 “2”
            strTemp(2) = Right("0000000000000000" & CheckStr(m_rs.Fields("SD06")), 16)    '帳號            X(16) 不足位數，右靠左補0
            strTemp(3) = Right("00000000000000" & CheckStr(m_rs.Fields("T2")) & "00", 16) '金額            9(16) 右靠左補0,後面兩位為小數
                                                                                          '                      EX: 金額為5000元，轉成文字檔為0000000000500000，如果金額為0，請放16個0
            strTemp(4) = "822"                                                            '行庫代號=822    X(3)
            strTemp(5) = Left(CheckStr(m_rs.Fields("ST26")) & "           ", 11)          '身份證號        X(11) 不足位數，左靠右補半形空白
            
            strText = ""
            For m_i = 1 To 5
               strText = strText & strTemp(m_i)
            Next m_i
            strText = strText & String(953, " ")
            Print #ff, strText
            
            m_rs.MoveNext
         Loop
         
      End With
      Close ff
      
      'Modify By Sindy 2023/9/25 mark,婉莘說取消
      'ZipFolder strFolder 'Added by Morgan 2013/7/30
   Else
      MsgBox "無符合條件的中所資料!!!", vbExclamation + vbOKOnly
      Exit Sub
   End If
   'end 2013/7/30
End Sub

'Add By Sindy 2022/6/21 產生高所(SD05='6')媒體檔案 - 合作金庫
Private Sub ProHighOfficeEFile(strType As String, strYM As String)
Dim ff As Integer
Dim strText As String
Dim TempFileName As String
Dim dblTotAmt As Double

   If strType = "1" Then
      m_str = "SELECT SD01,ST02,SD06,ST26,nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0) as T2 " & _
              "FROM STAFF,SALARYDATA,SalaryMonth WHERE SM02=" & strYM & " AND SM01=ST01(+) AND SM01=SD01(+) AND SD05='6' Order by SD01"
   ElseIf strType = "2" Then
      m_str = "SELECT SD01,ST02,SD06,ST26,nvl(YB05,0)+nvl(YB06,0)+nvl(YB26,0)+nvl(YB08,0)-nvl(YB15,0)-nvl(YB16,0)-nvl(YB17,0)-nvl(YB25,0) as T2 " & _
              "FROM STAFF,SALARYDATA,YearBonus WHERE YB01='" & strYM & "' AND YB02=ST01(+) AND YB02=SD01(+) AND SD05='6' AND SD02<>'S' Order by SD01"
   Else
      m_str = "SELECT SD01,ST02,SD06,ST26,nvl(ob05,0) T2 FROM Staff,SalaryData,ohbonus " & _
               "WHERE substr(ob01,1,4)='" & strYM & "' and ob05>0 AND ob03=st01(+) AND ob03=SD01(+) AND SD05='6' " & m_StrSQL
   End If
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      With m_rs
         m_rs.MoveFirst
         dblTotAmt = 0

         If ff > 0 Then Close #ff
         ff = FreeFile
         
         'Modify By Sindy 2023/9/25
         'TempFileName = PUB_Getdesktop & "\合作金庫薪資轉帳\salary高所合作金庫.txt"
         TempFileName = strFolder & "\salary高所合作金庫.txt"
'         strFolder = PUB_Getdesktop & "\合作金庫薪資轉帳"
'         If Dir(strFolder, vbDirectory) = "" Then
'            MkDir strFolder
'         End If
         '2023/9/25 END
         
         Open TempFileName For Output As ff
         
         '明細(無需首筆,直接列明細即可)
' 1.EDI用戶代碼/1-12  X(12)   M
' 2.手續費負擔別/13-15   X(3) M  13: 收款人負擔 , 15: 付款人負擔
' 3.付款日期/16-23   X(8) M  西元 YYYYMMDD
' 4.企業參考號碼/24-58  X(35)   C
' 5.付款金額/59-72   X(14)   M  右靠左補0
' 6.付款人帳號/73-88 X(16)   M
' 7.付款銀行代號/89-95  X(7) M
' 8.付款人身分證 (統一編號) / 96 - 105 X(10)   M
' 9.付款人戶名/106-185  X(80)   M
'10.收款人帳號/186-201  X(16)   M
'11.收款銀行代號/202-208   X(7) M
'12.收款人身分證 (統一編號) / 209 - 218   X(10)   C
'13.收款人戶名/219-298  X(80)   M
'14.收款人聯絡姓名/299-318 X(20)   C
'15.收款人傳真號碼/319-333 X(15)   C  EX：0223707640、033526874
'16.入帳通知處理方式/334-336  X(3) C  AF: 不發送入帳通知
'                                  AD: 預約及扣帳傳真 (酌收5元手續費)
'                                  AC: 扣帳傳真 (酌收3元手續費)
'                                  ML: email通知 (免費)
'                                  *未指定系統自動代入AF
'17.付款說明337-406 X70)   C 可帶35個中文字或70個英數字
'18.業務類別407-409 X(3) C   SAL:薪資，一般轉帳可不帶。
         Do While Not m_rs.EOF
            For m_i = 1 To 16
               strTemp(m_i) = ""
            Next m_i

            strTemp(1) = "041464570000"
            strTemp(2) = "15 "
            strTemp(3) = DBDATE(txt1(0)) '轉帳日期
            strTemp(4) = Left("                                   ", 35)
            strTemp(5) = Right("00000000000000" & CheckStr(m_rs.Fields("T2")), 14) '金額
            strTemp(6) = Left("5252717321283    ", 16) '付款人帳號
            strTemp(7) = Left("0065252", 7) '付款銀行代號
            strTemp(8) = Left("04146457  ", 10) '付款人身分證 (統一編號)
            strTemp(9) = Left("台一國際智慧財產事務所                                                                                ", 80 - Len("台一國際智慧財產事務所")) '付款人戶名
            strTemp(10) = Left(CheckStr(Trim(m_rs.Fields("SD06"))) & "    ", 16) '收款人帳號
            strTemp(11) = Left("0065252", 7) '收款銀行代號
            strTemp(12) = Left(CheckStr(m_rs.Fields("ST26")) & "           ", 10) '收款人身分證 (統一編號)
            strTemp(13) = Left(CheckStr(Trim(m_rs.Fields("ST02"))) & "                                                                                  ", 80 - Len(Trim(m_rs.Fields("ST02")))) '收款人戶名
            strTemp(14) = Left("                    ", 20)
            strTemp(15) = Left("               ", 15)
            strTemp(16) = Left("AF ", 3)
            'strTemp(17) = Left("                                                                      ", 70)
            'strTemp(18) = Left("   ", 3)
            
            strText = ""
            For m_i = 1 To 16
               strText = strText & strTemp(m_i)
            Next m_i
            strText = strText & String(268, " ") '總長度要到604,換行/605
            Print #ff, strText
            
            m_rs.MoveNext
         Loop
         
      End With
      Close ff
   Else
      MsgBox "無符合條件的高所資料!!!", vbExclamation + vbOKOnly
      Exit Sub
   End If
End Sub

'Add By Sindy 2023/8/9 產生南所(SD05='5')媒體檔案 - 台北富邦
Private Sub ProSouthOfficeEFile(strType As String, strYM As String)
Dim ff As Integer
Dim strText As String
Dim TempFileName As String
Dim dblTotAmt As Double

   If strType = "1" Then
      m_str = "SELECT SD01,ST02,SD06,ST26,nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0) as T2 " & _
              "FROM STAFF,SALARYDATA,SalaryMonth WHERE SM02=" & strYM & " AND SM01=ST01(+) AND SM01=SD01(+) AND SD05='5' Order by SD01"
   ElseIf strType = "2" Then
      m_str = "SELECT SD01,ST02,SD06,ST26,nvl(YB05,0)+nvl(YB06,0)+nvl(YB26,0)+nvl(YB08,0)-nvl(YB15,0)-nvl(YB16,0)-nvl(YB17,0)-nvl(YB25,0) as T2 " & _
              "FROM STAFF,SALARYDATA,YearBonus WHERE YB01='" & strYM & "' AND YB02=ST01(+) AND YB02=SD01(+) AND SD05='5' AND SD02<>'S' Order by SD01"
   Else
      m_str = "SELECT SD01,ST02,SD06,ST26,nvl(ob05,0) T2 FROM Staff,SalaryData,ohbonus " & _
               "WHERE substr(ob01,1,4)='" & strYM & "' and ob05>0 AND ob03=st01(+) AND ob03=SD01(+) AND SD05='5' " & m_StrSQL
   End If
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      With m_rs
         dblTotAmt = 0
         m_rs.MoveFirst
         Do While Not m_rs.EOF
            dblTotAmt = dblTotAmt + CDbl(m_rs.Fields("T2"))
            m_rs.MoveNext
         Loop
         
         If ff > 0 Then Close #ff
         ff = FreeFile
         'Modify By Sindy 2023/9/25
         'TempFileName = PUB_Getdesktop & "\台北富邦-薪轉上傳檔案.txt"
         TempFileName = strFolder & "\salary南所台北富邦.txt"
         '2023/9/25 END
         Open TempFileName For Output As ff
         
'首筆:
'區別碼  委託單位代號   存摺摘要 檢查身分證字號          轉帳日      轉帳發薪總額
'K       7碼數字        00000    (Y<12個*>/N<12個空白>)  民國日期    13碼(11碼+點數2位)
         'strText = "K807000100000************" & txt1(0) & Format(dblTotAmt, "00000000000") & "00" '有身分證字號的首筆
         strText = "K807000100000            " & txt1(0) & Format(dblTotAmt, "00000000000") & "00"  '無身分證字號的首筆
         Print #ff, strText
         
'明細:
'區別碼  委託單位代號   補零  員工北富銀之帳號  補入英文或數字    員工入帳薪資         員工身份證字號
'C       7碼數字        00000 14碼              000000            13碼(11碼+點數2位)   Y要檢查才顯示
'結果:
'K200001300000************11208040000234567800
'C200001300000007100022172000000000000234567800F222444333
         m_rs.MoveFirst
         Do While Not m_rs.EOF
            For m_i = 1 To 6
               strTemp(m_i) = ""
            Next m_i

            strTemp(1) = "8070001"
            strTemp(2) = "00000"
            strTemp(3) = Format(CheckStr(Trim(m_rs.Fields("SD06"))), "00000000000000")  '收款人帳號
            strTemp(4) = "000000"
            strTemp(5) = Format(CDbl(m_rs.Fields("T2")), "00000000000") & "00" '金額
            'strTemp(6) = Left(CheckStr(m_rs.Fields("ST26")) & "          ", 10) '收款人身分證 (統一編號)
            
            strText = "C"
            For m_i = 1 To 6
               strText = strText & strTemp(m_i)
            Next m_i
            Print #ff, strText
            
            m_rs.MoveNext
         Loop
         
      End With
      Close ff
      
   Else
      MsgBox "無符合條件的南所資料!!!", vbExclamation + vbOKOnly
      Exit Sub
   End If
End Sub

Private Sub Form_Load()
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim strSql As String, i As Integer, j As Integer
Dim strSystemKind As String
   
   MoveFormToCenter Me
   
   strSystemKind = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170104 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 1
         If KeyAscii < Asc("1") Or KeyAscii > Asc("4") Then
            KeyAscii = 0
            Beep
         End If
      '2013/3/6 add by sonia
      Case 2
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 89 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      '2013/3/6 end
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
End Sub
'Add by Morgan 2009/6/15
'新增薪資入帳記錄
'SDate:薪資月份,BDate:入帳日期
Private Function AddBookRecord(SalaryYM As String, BookDate As String) As Boolean

On Error GoTo ErrHnd
   cnnConnection.BeginTrans
   strSql = "delete BookRecord where BR01=" & SalaryYM
   cnnConnection.Execute strSql, intI
   strSql = "insert into BookRecord(BR01,BR02) values (" & SalaryYM & "," & BookDate & ")"
   cnnConnection.Execute strSql, intI
   cnnConnection.CommitTrans
   AddBookRecord = True
   
ErrHnd:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
End Function
