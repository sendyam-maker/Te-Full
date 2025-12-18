VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm12040138 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶/代理人名冊、地址條列印"
   ClientHeight    =   5100
   ClientLeft      =   2950
   ClientTop       =   1620
   ClientWidth     =   6390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6390
   Begin VB.CommandButton cmdOK 
      Caption         =   "離職通知函"
      Height          =   400
      Index           =   2
      Left            =   216
      TabIndex        =   32
      Top             =   48
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   13
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   11
      Text            =   "N"
      Top             =   3322
      Width           =   345
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   10
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   10
      Text            =   "N"
      Top             =   2992
      Width           =   345
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   12
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   9
      Text            =   "N"
      Top             =   2662
      Width           =   345
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   11
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   1
      Top             =   797
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   9
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   6
      Top             =   2010
      Width           =   1215
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   1080
      Style           =   2  '單純下拉式
      TabIndex        =   14
      Top             =   4440
      Width           =   2775
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5310
      TabIndex        =   16
      Top             =   48
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   4485
      TabIndex        =   15
      Top             =   48
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   8
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   13
      Top             =   4125
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   7
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   12
      Top             =   3645
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   6
      Left            =   2640
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1704
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   5
      Left            =   1080
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1704
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   4
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   8
      Top             =   2310
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   7
      Top             =   2310
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   2640
      MaxLength       =   8
      TabIndex        =   3
      Top             =   1395
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   1080
      MaxLength       =   8
      TabIndex        =   2
      Top             =   1395
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   0
      Top             =   483
      Width           =   495
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   216
      Left            =   240
      TabIndex        =   33
      Top             =   4824
      Visible         =   0   'False
      Width           =   4692
      _ExtentX        =   8273
      _ExtentY        =   388
      _Version        =   393216
      Appearance      =   1
      Max             =   20
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      Caption         =   "0/1"
      Height          =   180
      Left            =   5064
      TabIndex        =   34
      Top             =   4848
      Visible         =   0   'False
      Width           =   216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "5. E-mail核對)"
      Height          =   180
      Index           =   8
      Left            =   1740
      TabIndex        =   31
      Top             =   3900
      Width           =   1065
   End
   Begin VB.Label Label5 
      Caption         =   "是否含待活化客戶:                 (N: 不含)"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   3330
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "是否含有客戶狀態者:             (N: 不含)"
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "是否含不寄雜誌對象:             (N: 不含)"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   2670
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(A:代理人律師事務所 B:公司直接委辦 C:其他)"
      Height          =   180
      Index           =   7
      Left            =   240
      TabIndex        =   27
      Top             =   1100
      Width           =   3585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "性質：                   (可複選，請以 , 區隔) 對象為代理人才適用"
      Height          =   180
      Index           =   20
      Left            =   240
      TabIndex        =   26
      Top             =   842
      Width           =   4755
   End
   Begin VB.Label lblName 
      Height          =   180
      Left            =   2370
      TabIndex        =   25
      Top             =   2055
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員"
      Height          =   180
      Index           =   6
      Left            =   240
      TabIndex        =   24
      Top             =   2055
      Width           =   720
   End
   Begin VB.Label Label2 
      Caption         =   "印表機"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   4470
      Width           =   615
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2400
      X2              =   2520
      Y1              =   1832
      Y2              =   1832
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2400
      X2              =   2520
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2400
      X2              =   2520
      Y1              =   1530
      Y2              =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "列印語文                (1. 中 2. 英 )"
      Height          =   180
      Index           =   5
      Left            =   240
      TabIndex        =   22
      Top             =   4170
      Width           =   2370
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "列印方式                (1. 地址條 2. 名冊 3. 離職定稿 4.核對用客戶名冊 "
      Height          =   180
      Index           =   4
      Left            =   240
      TabIndex        =   21
      Top             =   3690
      Width           =   5235
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "開發日期"
      Height          =   180
      Index           =   3
      Left            =   240
      TabIndex        =   20
      Top             =   1749
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "國籍"
      Height          =   180
      Index           =   2
      Left            =   240
      TabIndex        =   19
      Top             =   2355
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "編號"
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   18
      Top             =   1440
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "列印對象                (1. 客戶 2. 代理人 )"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   17
      Top             =   528
      Width           =   2916
   End
End
Attribute VB_Name = "frm12040138"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit

Dim intWhere As Integer, strReceiveNo As String, PLeft(0 To 8) As Integer
' 預設印表機
Dim m_DefaultPrinter As String
'Add By Cheng 2002/09/10
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub PrintLetter(ByVal stCustName As String, ByVal stContact As String, ByVal stCustType As String, ByVal stSalesName As String, ByVal lngPage As Long)

   Const cLMargin = 1600 '左留白
   Const cUMargin = 3300 '上留白
   Const cVPad = 400 '列距
   
   Dim stDoc As String '文章內容
   Dim stSentence As String '一列文字
   Dim lngWMax As Long '可印行寬
   Dim lngWRest As Long '剩餘可印行寬
   Dim iAWord As Integer '一個字寬
   Dim sChar As String
   Dim lPos As Long
   Dim lngX As Long, lngY As Long '列印位置
   Dim sBoldi As String '粗體開
   Dim sBoldo As String '粗體關
   
   sBoldi = Chr(30) '粗體開
   sBoldo = Chr(31) '粗體關
   
   iAWord = Printer.TextWidth("　")
   lngWMax = iAWord * ((Printer.ScaleWidth - 2 * cLMargin) \ iAWord) + 50
   lngY = cUMargin
   
   '印頁次
   stSentence = Format(lngPage, "000000")
   lngX = cLMargin + lngWMax - Printer.TextWidth(stSentence)
   Printer.CurrentX = lngX: Printer.CurrentY = lngY
   Printer.Print stSentence
      
   lngX = cLMargin
   lngY = lngY + cVPad
   lngWRest = lngWMax
   stSentence = ""
   
   stDoc = "致：" & stCustName & vbLf & _
            "　　" & stContact & "　君台鑒" & vbLf & _
            "　　　　　　　　　　　　　　　　　　　　" & vbLf & _
            "　　感謝　" & stCustType & "多年來將智慧財產權之保護與申請案件委由本所處理。" & vbLf & _
            "　　台一在提供專利、商標、著作權及法律的各項服務時，一直秉持著專業與敬業的精神，為業界在智慧財產權保護上提供最完善的服務。在本所過去近三十年的成長過程中，承蒙　" & stCustType & "持續對本所支持與肯定，讓我們能夠不斷的邁入更多的十年，並服務更多的客戶。" & vbLf & _
            "　　前任職於本所之" & stSalesName & "同仁已於九十四年五月五日離職，為順利銜接　" & stCustType & "的案件，本所特指派" & sBoldi & "葉招進先生及林永生協理" & sBoldo & "共同接手　" & stCustType & "各項業務。日後若　" & stCustType & "有任何需要服務之案件或工作，請直接與他們兩位聯繫。尤其若有已委託本所尚未完成之案件，或付款尚未接獲收據之情形，皆請主動知會葉先生及林協理，本所將立即處理，以維　" & stCustType & "的權益。" & vbLf & _
            "　　本所備有齊全的管理與整合措施，任何智權同仁的案件皆由主管做雙重核對，且葉先生已在本所任職多年，林協理擔任主管職務數十年，經驗豐富定可繼續提供　" & stCustType & "更完整的服務，謹請　" & stCustType & "能一秉初衷繼續給予本所支持與肯定。" & vbLf & vbLf & _
            "　　耑此　　順頌" & vbLf & vbLf & _
            "商祺"
   
   For lPos = 1 To Len(stDoc)
      sChar = Mid(stDoc, lPos, 1)
      
      If sChar = vbLf Then '跳行
         Printer.CurrentX = lngX: Printer.CurrentY = lngY
         Printer.Print stSentence
         lngX = cLMargin
         lngY = lngY + cVPad
         lngWRest = lngWMax
         stSentence = ""
         
      ElseIf sChar = sBoldi Then  '粗體開
         Printer.CurrentX = lngX: Printer.CurrentY = lngY
         Printer.Print stSentence
         lngX = lngX + Printer.TextWidth(stSentence)
         lngWRest = lngWMax - lngX + cLMargin
         Printer.FontBold = True
         stSentence = ""
         
      ElseIf sChar = sBoldo Then  '粗體關
         Printer.CurrentX = lngX: Printer.CurrentY = lngY
         Printer.Print stSentence
         lngX = lngX + Printer.TextWidth(stSentence)
         lngWRest = lngWMax - lngX + cLMargin
         Printer.FontBold = False
         stSentence = ""
         
      ElseIf Printer.TextWidth(stSentence & sChar) > lngWRest Then '字數超過一列可印寬
         Printer.CurrentX = lngX: Printer.CurrentY = lngY
         Printer.Print stSentence
         
         If sChar = "　" Or sChar = " " Then
            stSentence = ""
         Else
            stSentence = sChar
         End If
         lngX = cLMargin
         lngY = lngY + cVPad
         lngWRest = lngWMax
         
      Else
         stSentence = stSentence & sChar
      End If
   Next
   If stSentence <> "" Then
      Printer.CurrentX = lngX: Printer.CurrentY = lngY
      Printer.Print stSentence
   End If
   
   lngY = lngY + cVPad
   'Modified by Morgan 2020/3/30 事務所名稱改用函數抓
   'stSentence = "台一國際專利商標事務所　敬上　　" & Year(Now) - 1911 & "." & Format(Now, "MM.DD")
   stSentence = CompNameQuery("2") & "　敬上　　" & Year(Now) - 1911 & "." & Format(Now, "MM.DD")
   'end 2020/3/30
   lngX = cLMargin + lngWMax - Printer.TextWidth(stSentence)
   Printer.CurrentX = lngX: Printer.CurrentY = lngY
   Printer.Print stSentence
   
      
End Sub
Private Sub cmdok_Click(Index As Integer)

   Screen.MousePointer = vbHourglass
   
   Dim rsTemp1 As New ADODB.Recordset
   Dim nPageNo As Integer
   Dim Prn As Printer
   Dim i As Integer
   Select Case Index
      Case 0 '確定
         '檢查輸入的資料是否齊全完整
         If CheckDataValid() = False Then
            GoTo EXITSUB
         End If
         '設定使用者所選擇的印表機成預設印表機
         For Each Prn In Printers
            If Prn.DeviceName = cmbPrinter.Text Then
               Set Printer = Prn
               Exit For
            End If
         Next
         
         '客戶
         If Text1(0).Text = "1" Then
            strSql = ""
            strExc(1) = ChangeCustomerL(Text1(1).Text)
            strExc(2) = ChangeCustomerL(Text1(2).Text)
            '客戶編號區間
            If Text1(1).Text <> "" And Text1(2).Text <> "" Then
               strSql = " WHERE (CU01 BETWEEN '" & Left(strExc(1), 8) & "' AND '" & Left(strExc(2), 8) & "')"
            ElseIf Text1(1).Text = "" And Text1(2).Text <> "" Then
               strSql = " WHERE CU01 <='" & Left(strExc(2), 8) & "'"
            ElseIf Text1(1).Text <> "" And Text1(2).Text = "" Then
               strSql = " WHERE CU01 >='" & Left(strExc(1), 8) & "'"
            End If
            
            '國籍區間
            If Text1(3).Text <> "" And Text1(4).Text <> "" Then
               strSql = strSql & " AND SUBSTR(CU10,1,3) BETWEEN '" & Text1(3).Text & "' AND '" & Text1(4).Text & "'"
            ElseIf Text1(3).Text = "" And Text1(4).Text <> "" Then
               strSql = strSql & " AND SUBSTR(CU10,1,3) <='" & Text1(4).Text & "'"
            ElseIf Text1(3).Text <> "" And Text1(4).Text = "" Then
               strSql = strSql & " AND SUBSTR(CU10,1,3) >='" & Text1(3).Text & "'"
            End If
            
            '開發日期區間
            If Text1(5).Text <> "" And Text1(6).Text <> "" Then
               strSql = strSql & " AND CU14 BETWEEN " & TransDate(Text1(5).Text, 2) & " AND " & TransDate(Text1(6).Text, 2)
            ElseIf Text1(5).Text = "" And Text1(6).Text <> "" Then
               'Modify By Cheng 2002/03/21
'               strSQL = strSQL & " AND CU14 <=" & TransDate(Text1(4).Text, 2)
               strSql = strSql & " AND CU14 <=" & TransDate(Text1(6).Text, 2)
            ElseIf Text1(5).Text <> "" And Text1(6).Text = "" Then
               'MOdify By Cheng 2002/03/21
'               strSQL = strSQL & " AND CU14 >=" & TransDate(Text1(3).Text, 2)
               strSql = strSql & " AND CU14 >=" & TransDate(Text1(5).Text, 2) & _
                                 " AND CU14 <=" & ServerDate & " "
            End If
            
'edit by nickc 2005/12/02 改成在代理人
'            '910625 Sieg
'            '性質
'            Dim varTmp As Variant
'
'            If Text1(11).Text <> "" Then
'               varTmp = Split(Text1(11).Text, ",")
'               strExc(0) = ""
'               For i = 0 To UBound(varTmp)
'                  strExc(0) = strExc(0) & "'" & Format(varTmp(i)) & "',"
'               Next
'               If Right(strExc(0), 1) = "," Then strExc(0) = Left(strExc(0), Len(strExc(0)) - 1)
'               strSQL = strSQL & " AND CU101 IN (" & strExc(0) & ")"
'            End If
            
            'Add By Cheng 2002/05/13
            If Len(Me.Text1(9).Text) > 0 Then
               strSql = strSql + " AND CU13='" & Me.Text1(9).Text & "' "
            End If
            '92.6.15 CANCEL BY SONIA
            'If Me.Text1(10).Text = "1" Then
            '  'Modify By Cheng 2003/02/19
        '   '   strSQL = strSQL + " AND SUBSTR(CU12,1,1)='S' "
            '   strSQL = strSQL + " AND (ST15<'F' OR ST15>'F99') "
            'ElseIf Me.Text1(10).Text = "2" Then
            '  'Modify By Cheng 2003/02/19
        '   '   strSQL = strSQL + " AND SUBSTR(CU12,1,1)<>'S' "
            '   strSQL = strSQL + " AND ST15>='F' AND ST15<='F99' "
            'Else
            '   '無動作
            'End If
            '92.6.15 END
            '是否含不寄雜誌對象 92.6.15 ADD BY SONIA
            If Text1(12) = "N" Then
               strSql = strSql & " AND CU32 IS NULL "
            End If
            '92.6.15 END
            '2007/8/30 ADD BY SONIA 是否含有客戶狀態者
            '2013/5/6 modify by sonia 再加其他
            If Text1(10) = "N" Then
               'modify by sonia 2021/10/15 客戶狀態已多元，改寫法
               'strSql = strSql & " AND (CU80 IS NULL OR CU80='業務自行處理' OR CU80='其他' ) "
               'modify by sonia 2023/2/23 再剔除設為對造
               strSql = strSql & " And (Cu80 Is Null Or Cu80 Not In ('不得代理','不再使用','倒閉','停業','刪址','宣告破產','廢止','撤銷','歇業','死亡','解散','遷移不明','設為對造'))"
            End If
            '92.6.15 END
            
            'Added by Lydia 2020/05/04 是否含待活化客戶
            If Text1(13) = "N" Then
               strSql = strSql & " AND CU01 NOT IN (SELECT OCU01 FROM OLDCUSTOMER WHERE OCU03 IS NULL) "
            End If
            'end 2020/05/04
            
'modify by sonia 2015/4/22楊挺客戶分挑選及非挑選
'            If Me.Text1(9).Text = "77010" Then
'               strSql = strSql & " AND CU01 not IN " & _
'                              "('X0020200','X0049710','X0049712','X0073800','X0096400','X0169600','X0460900','X0464500','X0522100','X0708300'," & _
'                              " 'X0720400','X1030500','X1172600','X1216500','X1360404','X1467700','X1560500','X1576000','X1630300','X1674300'," & _
'                              " 'X1677000','X1688600','X1692900','X1693200','X1748100','X1885400','X1904000','X1963400','X2233505','X2233506'," & _
'                              " 'X2285700','X2287600','X2395000','X2451200','X2458100','X2458500','X2472900','X2487300','X2633300','X2662300'," & _
'                              " 'X2663700','X2690900','X2871000','X2900800','X3072200','X3138800','X3155100','X3175800','X3175808','X3176400'," & _
'                              " 'X3181200','X3229700','X3322000','X3338300','X3346900','X3564800','X3587500','X3589000','X3655400','X3659600'," & _
'                              " 'X3673600','X3674600','X3728100','X3728800','X3882500','X3905605','X4031400','X4031500','X4042300','X4059500'," & _
'                              " 'X4068700','X4106400','X4114700','X4253600','X4259800','X4269400','X4269700','X4271900','X4278500','X4278600'," & _
'                              " 'X4331800','X4331801','X4351600','X4364600','X4369400','X4411500','X4488300','X4495100','X4495800','X5014200'," & _
'                              " 'X5020600','X5023600','X5034400','X5041700','X5068300','X5109700','X5109702','X5124800','X5125900','X5146700'," & _
'                              " 'X5149700','X5156700','X5171100','X5215100','X5230700','X5248500','X5308600','X5324600','X5328800','X5373700'," & _
'                              " 'X5401600','X5466900','X5466904','X5490000','X5497400','X5531500','X5554700','X5701000','X5718700','X5781400'," & _
'                              " 'X5781402','X5799400','X5823200','X5837100','X5861600','X5875900','X5876600','X5946300','X5948700','X5986700'," & _
'                              " 'X5999100','X6004400','X6027000','X6035500','X6051100','X6064400','X6079400','X6087000','X6101400','X6102700'," & _
'                              " 'X6144100','X6150000','X6151700','X6151701','X6188300','X6191600','X6201100','X6270000','X6295600'," & _
'                              " 'X6302000','X6304500','X6322800','X6323000','X6369200','X6371000','X6371001','X6388500','X6395300','X6403300'," & _
'                              " 'X6413700','X6428400','X6431200','X6468200','X6476200','X6514600','X6520600','X6541500','X6549600','X6555100'," & _
'                              " 'X6560400','X6565800','X6572600','X6591800','X6593900','X6602200','X6603500','X6605400','X6615300','X6633200'," & _
'                              " 'X6634000','X6635500','X6646600','X6673000','X6676600','X6686000','X6698800','X6707400','X6720300'," & _
'                              " 'X6724000','X6758600','X6776300','X6853300','X6861000','X6879600','X6883400','X6890600','X6919800','X6919900'," & _
'                              " 'X6921500','X6922500','X6923400','X6942200','X6949300','X6967400','X6989100','X6992800','X7028900'," & _
'                              " 'X7309700','X7058700','X7089700','X7092000','X7092600','X7101300','X7109700','X7136100','X7136200','X7136300'," & _
'                              " 'X7136400','X7145400','X7155300','X7162900','X7165000','X7168000','X7184000','X7194600','X7197400','X7200300'," & _
'                              " 'X7222600','X7235300','X7236100','X7254500','X7256500','X7257300','X7282000','X7316000') "
'            End If
'2015/4/22 end
            
            '地址條
            If Text1(7) = "1" Then
               '92.5.30 MODIFY BY SONIA
               'strExc(0) = "SELECT CU01 FROM CUSTOMER" & strSQL & " AND CU32 IS NULL AND CU02='0' AND CU24 IS NOT NULL ORDER BY CU10,CU05"
               'Modify by  Morgan 2004/11/23
               'strExc(0) = "SELECT CU01 FROM CUSTOMER" & strSQL & " AND CU02='0' ORDER BY CU10,CU05"
               '92.5.30 END
               strExc(0) = "SELECT CU01 FROM CUSTOMER" & strSql & " AND CU02='0' ORDER BY CU12,CU13,CU01,CU02"
               '2004/11/23 end
               
               intI = 1
               'edit by nickc 2007/02/09 不用 dll 了
               'Set rsTemp1 = objLawDll.ReadRstMsg(intI, strExc(0))
               Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  With rsTemp1
                     ' 90.07.12 modify by louis (流水號)
                     nPageNo = 1
                     Do While Not .EOF
                        If .Fields(0) <> "" Then
                           Load frm083014
                           frm083014.Hide
                           frm083014.opt1(0).Value = True
                           frm083014.Text1(0).Text = .Fields(0)
                           frm083014.Text1(3).Text = "1"
                           frm083014.Text1(4).Text = Text1(8).Text
                           '91.5.27 modify by sonia
                           frm083014.Text1(5).Text = "Y"
                           '91.5.27 end
                           ' 90.07.12 modify by louis (地址條流水號)
                           frm083014.SetPageNo nPageNo
                           ' 90.07.12 modify by louis (設定印表機)
                           frm083014.SetPrinter cmbPrinter.List(cmbPrinter.ListIndex)
                           frm083014.cmdPrint_Click
                           frm083014.cmdBack_Click
                        End If
                        .MoveNext
                        nPageNo = nPageNo + 1
                     Loop
                  End With
                  MsgBox "列印結束 !", vbInformation
               Else
                  MsgBox "無符合條件之資料可列印 !", vbInformation
               End If
               
            '名冊
            ElseIf Text1(7) = "2" Then
               PrintCase
               
            'Add by Morgan 2004/11/24 加定稿列印
            ElseIf Text1(7) = "3" Then
               'Modified by Morgan 2012/4/27
               'PrintLetterBatch strSql
               PrintLetterBatch2 strSql
            'Add by Morgan 2008/6/18 核對用客戶名冊
            ElseIf Text1(7) = "4" Then
               PrintCaseX
               PrintCase1 1
            '2009/2/13 add by sonia e-mail核對
            ElseIf Text1(7) = "5" Then
               PrintCase2
            End If
         '代理人
         Else
            strSql = ""
            strExc(1) = ChangeCustomerL(Text1(1).Text)
            strExc(2) = ChangeCustomerL(Text1(2).Text)
            If Text1(1).Text <> "" And Text1(2).Text <> "" Then
               strSql = " WHERE (FA01 BETWEEN '" & Left(strExc(1), 8) & "' AND '" & Left(strExc(2), 8) & "')"
            ElseIf Text1(1).Text = "" And Text1(2).Text <> "" Then
               strSql = " WHERE FA01 <='" & Left(strExc(2), 8) & "'"
            ElseIf Text1(1).Text <> "" And Text1(2).Text = "" Then
               strSql = " WHERE FA01 >='" & Left(strExc(1), 8) & "'"
            End If
            
            If Text1(3).Text <> "" And Text1(4).Text <> "" Then
               strSql = strSql & " AND SUBSTR(FA10,1,3) BETWEEN '" & Text1(3).Text & "' AND '" & Text1(4).Text & "'"
            ElseIf Text1(3).Text = "" And Text1(4).Text <> "" Then
               strSql = strSql & " AND SUBSTR(FA10,1,3) <='" & Text1(4).Text & "'"
            ElseIf Text1(3).Text <> "" And Text1(4).Text = "" Then
               strSql = strSql & " AND SUBSTR(FA10,1,3) >='" & Text1(3).Text & "'"
            End If
            
            If Text1(5).Text <> "" And Text1(6).Text <> "" Then
               strSql = strSql & " AND FA11 BETWEEN " & TransDate(Text1(5).Text, 2) & " AND " & TransDate(Text1(6).Text, 2)
            ElseIf Text1(5).Text = "" And Text1(6).Text <> "" Then
               'Modify By Cheng 2002/03/21
'               strSQL = strSQL & " AND FA11 <=" & TransDate(Text1(4).Text, 2)
               strSql = strSql & " AND FA11 <=" & TransDate(Text1(6).Text, 2)
            ElseIf Text1(5).Text <> "" And Text1(6).Text = "" Then
               'MOdify By Cheng 2002/03/21
'               strSQL = strSQL & " AND FA11 >=" & TransDate(Text1(3).Text, 2)
               strSql = strSql & " AND FA11 >=" & TransDate(Text1(5).Text, 2) & _
                                 " AND FA11 <=" & ServerDate & " "
            End If
            'add by nickc 2005/11/02 增加性質
            Dim varTmp As Variant

            If Text1(11).Text <> "" Then
               varTmp = Split(Text1(11).Text, ",")
               strExc(0) = ""
               For i = 0 To UBound(varTmp)
                  strExc(0) = strExc(0) & "'" & Format(varTmp(i)) & "',"
               Next
               If Right(strExc(0), 1) = "," Then strExc(0) = Left(strExc(0), Len(strExc(0)) - 1)
               strSql = strSql & " AND fa76 IN (" & strExc(0) & ")"
            End If

            '是否含不寄雜誌對象 92.6.15 ADD BY SONIA
            If Text1(12) = "N" Then
               strSql = strSql & " AND FA24 IS NULL "
            End If
            '92.6.15 END
            
            '地址條
            If Text1(7) = "1" Then
               '92.5.30 MODIFY BY SONIA
               'strExc(0) = "SELECT FA01 FROM FAGENT" & strSQL & " AND FA24 IS NULL AND FA02='0' AND FA18 IS NOT NULL ORDER BY FA10,FA05"
               strExc(0) = "SELECT FA01 FROM FAGENT" & strSql & " AND FA02='0' ORDER BY FA10,FA05"
               '92.5.30 END
               intI = 1
               'edit by nickc 2007/02/09 不用 dll 了
               'Set rsTemp1 = objLawDll.ReadRstMsg(intI, strExc(0))
               Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  With rsTemp1
                     Do While Not .EOF
                        If .Fields(0) <> "" Then
                           Load frm083014
                           frm083014.Hide
                           'Modify By Cheng 2002/05/14
'                           frm083014.Text1(0).Text = .Fields(0)
                           frm083014.Text1(1).Text = .Fields(0)
                           frm083014.opt1(1).Value = True
                           frm083014.Text1(3).Text = "1"
                           frm083014.Text1(4).Text = Text1(8).Text
                           '91.5.27 modify by sonia
                           frm083014.Text1(5).Text = "Y"
                           ' 91.09.03 nick (地址條流水號)
                           frm083014.SetPageNo nPageNo
                           ' 91.09.03 nick (設定印表機)
                           frm083014.SetPrinter cmbPrinter.List(cmbPrinter.ListIndex)
                           
                           '91.5.27 end
                           frm083014.cmdPrint_Click
                           frm083014.cmdBack_Click
                        End If
                        .MoveNext
                     Loop
                  End With
                  MsgBox "列印結束 !", vbInformation
               Else
                  MsgBox "無符合條件之資料可列印 !", vbInformation
               End If
            '名冊
            Else
               PrintCase
            End If
         End If
      Case 1 '結束
         Unload Me
      
      Case 2 'Added by Morgan 2024/3/6
         PrintLetterBatch3
   End Select
EXITSUB:
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   Dim Prn As Printer
   MoveFormToCenter Me
   ' 暫存預設印表機
   m_DefaultPrinter = Printer.DeviceName
   
   For Each Prn In Printers
      'Modify by Morgan 2008/8/11 預設也可選以便測試
      'If Prn.DeviceName <> m_DefaultPrinter Then
         cmbPrinter.AddItem Prn.DeviceName
      'End If
   Next
   If cmbPrinter.ListCount > 0 Then
      cmbPrinter.ListIndex = 0
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim Prn As Printer
   ' 將印表機設為原先的預設印表機
   For Each Prn In Printers
      If Prn.DeviceName = m_DefaultPrinter Then
         Set Printer = Prn
         Exit For
      End If
   Next
   
   Set frm12040138 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      'Modify by Morgan 2004/11/24 列印方式加 3.定稿
      'Modify by Morgan 2008/6/16 列印方式加 4.核對用客戶名冊
      '2009/2/13 MODIFY BY SONIA 列印方式加 5.E-mail核對
      Case 7
         If (KeyAscii < 49 Or KeyAscii > 53) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 0, 8
         If (KeyAscii > 50 Or KeyAscii < 49) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 11
         'edit by nickc 2005/12/02
         'If (KeyAscii > 68 Or KeyAscii < 65) And KeyAscii <> 8 And KeyAscii <> 44 Then
         If (KeyAscii > 66 Or KeyAscii < 65) And KeyAscii <> 8 And KeyAscii <> 44 Then
            KeyAscii = 0
            Beep
         End If
      Case 12, 10
         If KeyAscii <> 78 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   Select Case Index
      Case 2, 4, 6
         'Modify By Cheng 2002/09/10
         If blnClkSure = False Then
            If Text1(Index - 1) <> "" Then
               If RunNick(Text1(Index - 1), Text1(Index)) Then
                 Text1(Index - 1).SetFocus
               End If
            End If
         Else
            blnClkSure = False
         End If
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
 Dim strTmp As String
   Select Case Index
      Case 5, 6
         If Text1(Index) <> "" Then
            Cancel = Not ChkDate(Text1(Index).Text)
         End If
      'Modify By Cheng 2002/05/13
'      Case 0, 7, 8
      Case 0, 7, 8
         If Text1(Index) = "" Then
            MsgBox "請輸入資料 !", vbCritical
            Cancel = True
         Else
            ' 90.07.12 modify by louis
            If Index = 7 Then
               RefreshPrinterList
            End If
         End If
      Case 9
         lblName.Caption = ""
         If Text1(Index) <> "" Then
            '92.5.16 MODIFY BY SONIA
            lblName.Caption = GetPrjSalesNM(Text1(Index))
            If Len(Text1(Index)) <> 0 Then
               If Len(lblName.Caption) = 0 Then
                  Cancel = True
                  MsgBox "智權人員輸入錯誤！", vbCritical
                  Text1(Index).SetFocus
                  Text1_GotFocus (Index)
                  Exit Sub
               End If
            End If
            'If Not objPublicData.GetStaff(Text1(Index), strExc(0)) Then
            '   Cancel = True
            'Else
            '   lblName.Caption = strExc(0)
            'End If
            '92.5.16 END
         End If
   End Select
   If Cancel Then TextInverse Text1(Index)
End Sub

' 更新可供選擇的印表機列表
Private Sub RefreshPrinterList()
   Dim Prn As Printer
   
   cmbPrinter.Clear
   Select Case Text1(7)
      ' 地址條
      Case 1:
         For Each Prn In Printers
            'Modify by Morgan 2008/8/11 預設也可選以便測試
            'If Prn.DeviceName <> m_DefaultPrinter Then
               cmbPrinter.AddItem Prn.DeviceName
            'End If
         Next
         If cmbPrinter.ListCount > 0 Then
            cmbPrinter.ListIndex = 0
         End If
         cmbPrinter.Enabled = True
      ' 名冊
      Case 2, 5:
         cmbPrinter.AddItem m_DefaultPrinter
         If cmbPrinter.ListCount > 0 Then
            cmbPrinter.ListIndex = 0
         End If
         cmbPrinter.Enabled = False
      ' 其它
      Case Else:
         For Each Prn In Printers
            cmbPrinter.AddItem Prn.DeviceName
         Next
         cmbPrinter.Enabled = True
   End Select
End Sub

Private Function GetNation(ByVal strTmp As String) As Boolean
   GetNation = False
   strExc(0) = "SELECT COUNT(*) FROM NATION WHERE SUBSTR(NA01,1,3)='" & strTmp & "'"
   intI = 1
   'edit by nickc 2007/02/09 不用 dll 了
   'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.Fields(0) = 1 Then GetNation = True
   End If
End Function

Private Sub PrintCase()
 Dim i As Integer, j As Integer, Page As Integer, iPrint As Integer
 Dim strTmp As String, strTmp1 As String
 'Add By Cheng 2002/05/13
 Dim strNo As String '員工代號
 
On Error GoTo ErrHand

   '客戶
   If Text1(0).Text = "1" Then
      '列印中文資料
      If Text1(8) = "1" Then
         '2008/2/15 modify by sonia 加接洽人及手機
         'Modify by Morgan 2008/8/11 接洽人改抓聯絡人檔
         'strExc(0) = "SELECT CU01||CU02,SUBSTR(CU04,1,30),SUBSTR(CU07,1,10),SUBSTR(NA03,1,14),CU16,CU30,SUBSTR(NVL(CU31,CU23),1,65),CU18,CU32,CU80,CU13,ST02,CU12,CU23,CU11,CU08,CU22 FROM CUSTOMER,NATION,STAFF " & strSQL & " AND CU02='0' AND CU10=NA01(+) AND CU13=ST01(+) "
         strExc(0) = "SELECT CU01||CU02,SUBSTR(CU04,1,30),SUBSTR(CU07,1,10),SUBSTR(NA03,1,14),CU16,CU30,SUBSTR(NVL(CU31,CU23),1,65),CU18,CU32,CU80,CU13,ST02,CU12,CU23,CU11,PCC05,CU22 FROM CUSTOMER,POTCUSTCONT,NATION,STAFF " & strSql & " AND CU02='0' AND CU10=NA01(+) AND CU13=ST01(+) AND PCC01(+)=CU01 AND PCC02(+)=CU127 "
         strExc(0) = strExc(0) + " ORDER BY CU12,CU13,CU01,CU02 "
      '列印英文資料
      Else
        '加英文地址
         'Modify by Morgan 2008/8/11 接洽人改抓聯絡人檔
         'strExc(0) = "SELECT CU01||CU02,SUBSTR(CU05||CU88||CU89||CU90,1,30),SUBSTR(CU07,1,10),SUBSTR(NA04,1,14),CU16,CU30,SUBSTR(DECODE(CU65,NULL,CU24||CU25||CU26||CU27||CU28||CU102,CU65||CU66||CU67||CU68||CU69),1,65),CU18,CU32,CU80,CU13,ST02,CU12,CU24||CU25||CU26||CU27||CU28||CU102,CU11,CU08,CU22 FROM CUSTOMER,NATION,STAFF " & strSQL & " AND CU02='0' AND CU10=NA01(+) AND CU13=ST01(+) "
         strExc(0) = "SELECT CU01||CU02,SUBSTR(CU05||CU88||CU89||CU90,1,30),SUBSTR(CU07,1,10),SUBSTR(NA04,1,14),CU16,CU30,SUBSTR(DECODE(CU65,NULL,CU24||CU25||CU26||CU27||CU28||CU102,CU65||CU66||CU67||CU68||CU69),1,65),CU18,CU32,CU80,CU13,ST02,CU12,CU24||CU25||CU26||CU27||CU28||CU102,CU11,PCC05,CU22 FROM CUSTOMER,POTCUSTCONT,NATION,STAFF " & strSql & " AND CU02='0' AND CU10=NA01(+) AND CU13=ST01(+) AND PCC01(+)=CU01 AND PCC02(+)=CU127 "
         strExc(0) = strExc(0) + " ORDER BY CU12,CU13,CU01,CU02 "
      End If
   '代理人
   Else
      '列印中文資料
      If Text1(8) = "1" Then
         strExc(0) = "SELECT FA01||FA02,SUBSTR(FA04,1,30),SUBSTR(FA07,1,10),SUBSTR(NA03,1,14),FA12,' ',SUBSTR(FA17,1,65),FA14,FA24,FA69,' ',' ',' ',' ',' ',' ',' ' FROM FAGENT,NATION " & strSql & " AND FA02='0' AND FA10=NA01(+) "
         strExc(0) = strExc(0) + " ORDER BY FA01,FA02 "
      
      '列印英文資料
      Else
         strExc(0) = "SELECT FA01||FA02,SUBSTR(FA05||FA63||FA64||FA65,1,30),SUBSTR(FA08,1,10),SUBSTR(NA04,1,14),FA12,' ',SUBSTR(DECODE(FA32,NULL,FA18||' '||FA19||' '||FA20||' '||FA21||' '||FA22||' '||FA70,FA32||FA33||FA34||FA35||FA36),1,65),FA14,FA24,FA69,' ',' ',' ' FROM FAGENT,NATION " & strSql & " AND FA02='0' AND FA10=NA01(+) "
         strExc(0) = strExc(0) + " ORDER BY FA01,FA02 "
      
      End If
   End If
   
   intI = 0
   'edit by nickc 2007/02/09 不用 dll 了
   'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Screen.MousePointer = vbHourglass
      GetPrintLeft
      Page = 1
      CaseTitle Page, "" & RsTemp.Fields(10).Value, "" & RsTemp.Fields(11).Value
        '若列印對象為客戶
        If Me.Text1(0).Text = "1" Then
            iPrint = 2700 + 300 + 300
        '若列印對象為代理人
        Else
            iPrint = 2700 + 300
        End If
      If Not IsNull(RsTemp.Fields(10).Value) Then strNo = RsTemp.Fields(10).Value
      
      i = 0
      With RsTemp
         i = 0
         Do While Not .EOF
            For j = 0 To 4
               Printer.CurrentX = PLeft(j)
               Printer.CurrentY = iPrint
               If j = 0 Then
                  '92.6.15 MODIFY BY SONIA
                  'If Not IsNull(rsTemp.Fields(8).Value) Or Not IsNull(rsTemp.Fields(9).Value) Then
                  If Not IsNull(RsTemp.Fields(9).Value) Then
                  '92.6.15 END
                     Printer.CurrentX = PLeft(j) - Printer.TextWidth("＊")
                     '2007/9/3 MODIFY BY SONIA
                     'Printer.Print "＊" & Left(.Fields(j) & "000", 9)
                     If Not IsNull(RsTemp.Fields(8).Value) Then
                        Printer.Print "＊" & Left(.Fields(j) & "000", 9) & " N"
                     Else
                        Printer.Print "＊" & Left(.Fields(j) & "000", 9)
                     End If
                     '2007/9/3 END
                  Else
                  '92.6.15 ADD BY SONIA
                     If Not IsNull(RsTemp.Fields(8).Value) Then
                        Printer.Print " " & Left(.Fields(j) & "000", 9) & " N"
                     Else
                        Printer.Print "" & Left(.Fields(j) & "000", 9)
                     End If
                  End If
               Else
                  Printer.Print "" & .Fields(j)
               End If
            Next j
            
            iPrint = iPrint + 300
            
            For j = 5 To 7
               '2008/2/15 ADD BY SONIA 加接洽人
               If Me.Text1(0).Text = "1" And j = 7 Then
                  Printer.CurrentX = PLeft(8)
                  Printer.CurrentY = iPrint
                  Printer.Print "" & .Fields(15)
               End If
               '2008/2/15 END
               Printer.CurrentX = PLeft(j)
               Printer.CurrentY = iPrint
               Printer.Print "" & .Fields(j)
            Next j
            
            iPrint = iPrint + 300
                        
            'Add By Cheng 2003/02/19
            '若列印對象為客戶
            If Me.Text1(0).Text = "1" Then
               '2007/11/8 add by sonia
               Printer.CurrentX = PLeft(5)
               Printer.CurrentY = iPrint
               Printer.Print "" & .Fields(9)
               '2007/11/8 end
               Printer.CurrentX = PLeft(6)
               Printer.CurrentY = iPrint
               Printer.Print "" & .Fields(13)
               '2008/2/15 ADD BY SONIA 加手機
               Printer.CurrentX = PLeft(8)
               Printer.CurrentY = iPrint
               Printer.Print "" & .Fields(16)
               '2008/2/15 END
               Printer.CurrentX = PLeft(7)
               Printer.CurrentY = iPrint
               Printer.Print "" & .Fields(14)
               iPrint = iPrint + 300
            End If
                        
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print String(250, "-")
            
            iPrint = iPrint + 300
            
            i = i + 1
            .MoveNext
            If RsTemp.EOF Then Exit Do
            'Modify By Cheng 2003/02/19
            '判斷列印別為客戶或代理人
'            If i > 13 Or "" & rsTemp.Fields(10).Value <> strNo Then
            If i > IIf(Me.Text1(0).Text = "1", 9, 13) Or "" & RsTemp.Fields(10).Value <> strNo Then
               strNo = "" & RsTemp.Fields(10).Value
               '92.6.17 ADD BY SONIA
               Printer.CurrentX = PLeft(0)
               Printer.CurrentY = iPrint
               Printer.Print "PS : 編號與公司名稱之間, 若有 N 表示不寄台一雜誌"
            
               Printer.NewPage
               Page = Page + 1
               CaseTitle Page, "" & RsTemp.Fields(10).Value, "" & RsTemp.Fields(11).Value
                'Modify By Cheng 2003/02/19
                '若列印對象為客戶
                If Me.Text1(0).Text = "1" Then
                    iPrint = 2700 + 300 + 300
                '若列印對象為代理人
                Else
                    iPrint = 2700 + 300
                End If
               i = 0
            End If
         Loop
      End With
      '92.6.17 ADD BY SONIA
      Printer.CurrentX = PLeft(0)
      Printer.CurrentY = iPrint
      Printer.Print "PS : 編號與公司名稱之間, 若有 N 表示不寄台一雜誌"
            
      Printer.EndDoc
      'Add By Cheng 2002/05/14
      ShowPrintOk
      Screen.MousePointer = vbDefault
   End If
   Exit Sub
ErrHand:
   MsgBox Err.Description
End Sub

Private Sub GetPrintLeft()
'Modify By Cheng 2002/05/13
'   PLeft(0) = 200:     PLeft(1) = 1500
'   PLeft(2) = 3000:    PLeft(3) = 6400
'   PLeft(4) = 8100:    PLeft(5) = 9800
   '第一列
   PLeft(0) = 200
   PLeft(1) = 1500
   PLeft(2) = 4200 + 3000 - 500 - 500
   '2008/2/15 modify by sonia 因加接洽人及手機故往右移
   'PLeft(3) = 5200 + 3000 - 500
   'PLeft(4) = 6200 + 3000 + 500 - 500
   PLeft(3) = 5200 + 3000 - 300
   PLeft(4) = 6200 + 3000 + 500 - 300
   '2008/2/15 end
   '第二列
   PLeft(5) = 200
   PLeft(6) = 1500
   '2008/2/15 modify by sonia 因加接洽人及手機故往右移
   'PLeft(7) = 6200 + 3000 + 500 - 500
   PLeft(7) = 6200 + 3000 + 500 - 300
   PLeft(8) = 5200 + 3000 - 300          '接洽人及手機
   '2008/2/15 end
End Sub

Private Sub CaseTitle(ByVal Page As String, ByVal strSNo As String, ByVal strSName As String, Optional iReportID As Integer, Optional strDept As String)
'Page : 頁數
'strSNo : 員工編號
'strSName : 員工姓名
 
 Dim i As Integer
   i = 500
   'Modify By Cheng 2002/05/13
   '改成直印
'   Printer.Orientation = vbPRORLandscape
   If Page = 1 Then Printer.Orientation = vbPRORPortrait
   Printer.FontName = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 4000
   Printer.CurrentY = i
   '2009/2/16 add by sonia
   If iReportID = 7 Then
      Printer.Print "核對E-mail之客戶名冊"
   '2009/2/16 end
   ElseIf iReportID > 0 Then
      Printer.Print "智權人員核對用之客戶名冊"
   Else
      Printer.Print "客戶/代理人名冊"
   End If
   Printer.Font.Underline = False
   
   'Modify by Morgan 2008/6/17
   strExc(2) = ""
   Select Case iReportID
      Case 1
         strExc(2) = "(業務自行處理，但三年內有收文之客戶)"
      Case 2
         strExc(2) = "(業務自行處理，但三年內無收文之客戶)"
      Case 3
         strExc(2) = "(有客戶狀態，但三年內有收文之客戶)"
      Case 4
         strExc(2) = "(有客戶狀態，但三年內無收文之客戶)"
      Case 5
         strExc(2) = "(一般客戶，但三年內有收文之客戶)"
      Case 6
         strExc(2) = "(一般客戶，但三年內無收文之客戶)"
      Case 7
         strExc(2) = "(三年內有收文之客戶)"
   End Select
   If strExc(2) <> "" Then
      Printer.Font.Size = 14
      Printer.CurrentY = i + 500
      Printer.CurrentX = 6100 - Printer.TextWidth(strExc(2)) / 2
      Printer.Print strExc(2)
   End If
   'end 2008/6/17
   
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = i + 800 - 300
'   Printer.Print "列印人 : " & strUserName
   Printer.Print "列印人　 : " & strUserName
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = i + 800
'   Printer.Print "智權人員 : " & strSNo & " " & strSName
   Printer.Print "智權人員　 : " & strSNo & " " & strSName & " " & strDept
   Printer.CurrentX = 7000 + 1500
   Printer.CurrentY = i + 800
'   Printer.Print "列印日期 : " & Format(Date, "yy/MM/dd")
   Printer.Print "列印日期 : " & ChangeTStringToTDateString("" & (Val(ServerDate) - 19110000))
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = i + 1100
   Printer.Print "開發日期 : " & ChangeTStringToTDateString(Me.Text1(5).Text) & " - " & ChangeTStringToTDateString(Me.Text1(6).Text)
   Printer.CurrentX = 7000 + 1500
   Printer.CurrentY = i + 1100
   Printer.Print "頁　　次 : " & Page
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = i + 1400
   Printer.Print String(250, "-")
   
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = i + 1700
   Printer.Print "編號"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = i + 1700
   Printer.Print "公司名稱"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = i + 1700
   Printer.Print "負責人"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = i + 1700
   Printer.Print "國籍"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = i + 1700
   Printer.Print "電話"
   
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = i + 1700 + 300
   Printer.Print "郵遞區號"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = i + 1700 + 300
'   Printer.Print "聯絡電話"
   Printer.Print "客戶聯絡地址"
   '2008/2/15 add by sonia
   If Me.Text1(0).Text = "1" Then
      Printer.CurrentX = PLeft(8)
      Printer.CurrentY = i + 1700 + 300
      Printer.Print "接洽人"
   End If
   '2008/2/15 end
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = i + 1700 + 300
   Printer.Print "傳真"
   
   'Add By Cheng 2003/02/19
   '若列印對象為客戶
   If Me.Text1(0).Text = "1" Then
      '2007/11/8 add by sonia
      Printer.CurrentX = PLeft(5)
      Printer.CurrentY = i + 1700 + 300 + 300
      Printer.Print "客戶狀態"
      '2007/11/8 end
      Printer.CurrentX = PLeft(6)
      Printer.CurrentY = i + 1700 + 300 + 300
      Printer.Print "中文地址"
      '2008/2/15 add by sonia
      Printer.CurrentX = PLeft(8)
      Printer.CurrentY = i + 1700 + 300 + 300
      '2009/2/13 MODIFY BY SONIA E-MAIL核對時手機+統一編號改為E-mail
      If Text1(7) = "5" Then
         Printer.Print "E-MAIL"
      Else
         Printer.Print "手機"
      End If
      '2008/2/15 end
      If Text1(7) <> "5" Then
         Printer.CurrentX = PLeft(7)
         Printer.CurrentY = i + 1700 + 300 + 300
         Printer.Print "統一編號"
      End If
   End If
   Printer.CurrentX = PLeft(0)
    'Modify By Cheng 2003/02/19
    '若列印對象為客戶
    If Me.Text1(0).Text = "1" Then
        Printer.CurrentY = i + 2000 + 300 + 300
    '若列印對象為代理人
    Else
        Printer.CurrentY = i + 2000 + 300
    End If
   Printer.Print String(250, "-")
   Printer.Font.Size = 10
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   'Add By Cheng 2002/09/10
   blnClkSure = False
   
   '列印對象
   If Len(Me.Text1(0).Text) <= 0 Then
      MsgBox "請輸入列印對象!!!", vbExclamation + vbOKOnly
      Text1(0).SetFocus
      GoTo EXITSUB
   End If
   
   ' 編號範圍
   'Add By Cheng 2002/05/14
   If Len(Me.Text1(1).Text) <= 0 Then
      MsgBox "請輸入起始編號!!!", vbExclamation + vbOKOnly
      Text1(1).SetFocus
      GoTo EXITSUB
   End If
   If Len(Me.Text1(2).Text) <= 0 Then
      MsgBox "請輸入起始編號!!!", vbExclamation + vbOKOnly
      Text1(2).SetFocus
      GoTo EXITSUB
   End If
   
   If IsEmptyText(Text1(1)) = False And IsEmptyText(Text1(2)) = False Then
      If Text1(1) > Text1(2) Then
         strTit = "檢核資料"
         strMsg = "編號範圍不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         'Add By Cheng 2002/09/10
         blnClkSure = True
         Text1(1).SetFocus
         Text1_GotFocus 1
         GoTo EXITSUB
      End If
   End If
   
   '開發日期
   'Modify By Cheng 2002/05/14
'   'Add By Cheng 2002/05/13
'   '若為列印名冊
'   If Me.Text1(7).Text = "2" Then
'      If Len(Me.Text1(5).Text) <= 0 Then
'         MsgBox "請輸入開發起日!!!", vbExclamation + vbOKOnly
'         Me.Text1(5).SetFocus
'         Text1_GotFocus 5
'         GoTo EXITSUB
'      End If
'      If Len(Me.Text1(6).Text) <= 0 Then
'         MsgBox "請輸入開發迄日!!!", vbExclamation + vbOKOnly
'         Me.Text1(6).SetFocus
'         Text1_GotFocus 6
'         GoTo EXITSUB
'      End If
'   End If
   'Add By Cheng 2002/03/21
   If PUB_CheckKeyInDate(Me.Text1(5)) = -1 Then
      Me.Text1(5).SetFocus
      Text1_GotFocus 5
      GoTo EXITSUB
   End If
   If PUB_CheckKeyInDate(Me.Text1(6)) = -1 Then
      Me.Text1(6).SetFocus
      Text1_GotFocus 6
      GoTo EXITSUB
   End If
   If Val("0" & Me.Text1(5).Text) > Val("0" & Me.Text1(6).Text) Then
      MsgBox "開發日期輸入範圍錯誤!!!", vbExclamation + vbOKOnly
      'Add By Cheng 2002/09/10
      blnClkSure = True
      Me.Text1(5).SetFocus
      Text1_GotFocus 5
      GoTo EXITSUB
   End If
      
   lblName.Caption = ""
   If Text1(9) <> "" Then
      '92.5.16 MODIFY BY SONIA
      lblName.Caption = GetPrjSalesNM(Text1(9))
      If Len(Text1(9)) <> 0 Then
         If Len(lblName.Caption) = 0 Then
            MsgBox "智權人員輸入錯誤！", vbCritical
            Text1(9).SetFocus
            Text1_GotFocus (9)
            GoTo EXITSUB
         End If
      End If
      'If Not objPublicData.GetStaff(Text1(9), strExc(0)) Then
      '   Me.Text1(9).SetFocus
      '   Text1_GotFocus 9
      '   GoTo EXITSUB
      'Else
      '   lblName.Caption = strExc(0)
      'End If
      '92.5.16 END
   End If
      
   '國籍範圍
   If Me.Text1(3).Text > Me.Text1(4).Text Then
      MsgBox "國籍輸入範圍錯誤!!!", vbExclamation + vbOKOnly
      'Add By Cheng 2002/09/10
      blnClkSure = True
      Me.Text1(3).SetFocus
      Text1_GotFocus 3
      GoTo EXITSUB
   End If
   '列印方式
   If Len(Me.Text1(7).Text) <= 0 Then
      MsgBox "請輸入列印方式!!!", vbExclamation + vbOKOnly
      Me.Text1(7).SetFocus
      Text1_GotFocus 7
      GoTo EXITSUB
   End If
   '列印語文
   If Len(Me.Text1(8).Text) <= 0 Then
      MsgBox "請輸入列印語文!!!", vbExclamation + vbOKOnly
      Me.Text1(8).SetFocus
      Text1_GotFocus 8
      GoTo EXITSUB
   End If
   
   CheckDataValid = True

EXITSUB:
End Function

'Add by Morgan 2008/6/17
Private Function PrintCase1(Optional p_iReportID As Integer, Optional p_stST01 As String, Optional p_Rst As ADODB.Recordset) As Integer
 Dim i As Integer, j As Integer, Page As Integer, iPrint As Integer
 Dim strTmp As String, strTmp1 As String
 Dim strNo As String '員工代號
 Dim strName As String
 Dim strDept As String

On Error GoTo ErrHand
      
   If p_iReportID = 0 Then
      strExc(0) = "select distinct ST03,CU13,ST02 FROM CUSTOMER,NATION,STAFF " & strSql & " AND CU10=NA01(+) AND CU13=ST01(+) "
   Else
      'Modify by Morgan 2008/7/25 接洽人改先用聯絡人編號抓聯絡人檔,若有聯絡人編號但該聯絡人無地址時抓原客戶地址
      'strExc(0) = "SELECT CU01||CU02,SUBSTR(CU04,1,30),SUBSTR(CU07,1,10),SUBSTR(NA03,1,14),CU16,CU30,SUBSTR(NVL(CU31,CU23),1,65),CU18,CU32,CU80,CU13,ST02,CU12,CU23,CU11,CU08,CU22,ST03 FROM CUSTOMER,NATION,STAFF " & strSQL & " AND CU10=NA01(+) AND CU13=ST01(+) "
      'Modify by Morgan 2009/11/30 不用再考慮舊的聯絡人欄位
      'strExc(0) = "SELECT CU01||CU02,SUBSTR(CU04,1,30),SUBSTR(CU07,1,10),SUBSTR(NA03,1,14),CU16,NVL(PCC21,CU30) CU30,SUBSTR(NVL(NVL(PCC22,CU31),CU23),1,65),CU18,CU32,CU80,CU13,ST02,CU12,CU23,CU11,NVL(PCC05,CU08) CU08,CU22,ST03 FROM CUSTOMER,POTCUSTCONT,NATION,STAFF " & strSQL & " AND CU10=NA01(+) AND CU13=ST01(+) AND PCC01(+)=CU01 AND PCC02(+)=CU127 "
      strExc(0) = "SELECT CU01||CU02,SUBSTR(CU04,1,30),SUBSTR(CU07,1,10),SUBSTR(NA03,1,14),CU16,NVL(PCC21,CU30) CU30,SUBSTR(NVL(NVL(PCC22,CU31),CU23),1,65),CU18,CU32,CU80,CU13,ST02,CU12,CU23,CU11,PCC05 CU08,CU22,ST03 FROM CUSTOMER,POTCUSTCONT,NATION,STAFF " & strSql & " AND CU10=NA01(+) AND CU13=ST01(+) AND PCC01(+)=CU01 AND PCC02(+)=CU127 "
      If p_stST01 <> "" Then
         strExc(0) = strExc(0) & " AND CU13='" & p_stST01 & "'"
      Else
         strExc(0) = strExc(0) & " AND CU13 is null"
      End If
   End If
   
   strExc(0) = strExc(0) & " and (cu80 is null or cu80<>'不再使用')"
   '國外部客戶且國籍非台灣的不要
   strExc(0) = strExc(0) & " AND CU13=ST01(+) and NOT (SUBSTR(ST15,1,1)='F' AND CU10>='010')"
   strExc(0) = strExc(0) & " AND CU10=NA01(+)"
   Select Case p_iReportID
      Case 1, 2 '客戶狀態=業務自行處理
         '2013/5/6 modify by sonia 再加其他
         strExc(0) = strExc(0) & " and (cu80='業務自行處理' or cu80='其他')"
      Case 3, 4 '客戶狀態<>業務自行處理
         '2013/5/6 modify by sonia 再加其他
         'modify by sonia 2023/2/23 再加解除對造
         strExc(0) = strExc(0) & " and cu80<>'業務自行處理' and cu80<>'其他' and cu80<>'解除對造'"
      Case 5, 6 '無客戶狀態
         'modify by sonia 2023/2/23 再加解除對造
         strExc(0) = strExc(0) & " and (cu80 is null or cu80='解除對造')"
   End Select
   
   strExc(0) = strExc(0) & " AND (" & _
      " EXISTS(SELECT * FROM TRADEMARK WHERE TM23=CU01||CU02 AND TM29 IS NULL AND TM57 IS NULL AND SUBSTR(TM01,1,1)<>'F'" & _
      " AND (TM22 IS NULL OR TM22>TO_CHAR(SYSDATE,'YYYYMMDD')))" & _
      " OR EXISTS(SELECT * FROM TRADEMARK WHERE TM78=CU01||CU02 AND TM29 IS NULL AND TM57 IS NULL AND SUBSTR(TM01,1,1)<>'F'" & _
      " AND (TM22 IS NULL OR TM22>TO_CHAR(SYSDATE,'YYYYMMDD')))" & _
      " OR EXISTS(SELECT * FROM TRADEMARK WHERE TM79=CU01||CU02 AND TM29 IS NULL AND TM57 IS NULL AND SUBSTR(TM01,1,1)<>'F'" & _
      " AND (TM22 IS NULL OR TM22>TO_CHAR(SYSDATE,'YYYYMMDD')))" & _
      " OR EXISTS(SELECT * FROM TRADEMARK WHERE TM80=CU01||CU02 AND TM29 IS NULL AND TM57 IS NULL AND SUBSTR(TM01,1,1)<>'F'" & _
      " AND (TM22 IS NULL OR TM22>TO_CHAR(SYSDATE,'YYYYMMDD')))" & _
      " OR EXISTS(SELECT * FROM TRADEMARK WHERE TM81=CU01||CU02 AND TM29 IS NULL AND TM57 IS NULL AND SUBSTR(TM01,1,1)<>'F'" & _
      " AND (TM22 IS NULL OR TM22>TO_CHAR(SYSDATE,'YYYYMMDD')))"
      
   strExc(0) = strExc(0) & _
      " OR EXISTS(SELECT * FROM PATENT WHERE PA26=CU01||CU02 AND PA57 IS NULL AND PA108 IS NULL AND SUBSTR(PA01,1,1)<>'F'" & _
      " AND (PA25 IS NULL OR PA25>TO_CHAR(SYSDATE,'YYYYMMDD')))" & _
      " OR EXISTS(SELECT * FROM PATENT WHERE PA27=CU01||CU02 AND PA57 IS NULL AND PA108 IS NULL AND SUBSTR(PA01,1,1)<>'F'" & _
      " AND (PA25 IS NULL OR PA25>TO_CHAR(SYSDATE,'YYYYMMDD')))" & _
      " OR EXISTS(SELECT * FROM PATENT WHERE PA28=CU01||CU02 AND PA57 IS NULL AND PA108 IS NULL AND SUBSTR(PA01,1,1)<>'F'" & _
      " AND (PA25 IS NULL OR PA25>TO_CHAR(SYSDATE,'YYYYMMDD')))" & _
      " OR EXISTS(SELECT * FROM PATENT WHERE PA29=CU01||CU02 AND PA57 IS NULL AND PA108 IS NULL AND SUBSTR(PA01,1,1)<>'F'" & _
      " AND (PA25 IS NULL OR PA25>TO_CHAR(SYSDATE,'YYYYMMDD')))" & _
      " OR EXISTS(SELECT * FROM PATENT WHERE PA30=CU01||CU02 AND PA57 IS NULL AND PA108 IS NULL AND SUBSTR(PA01,1,1)<>'F'" & _
      " AND (PA25 IS NULL OR PA25>TO_CHAR(SYSDATE,'YYYYMMDD')))"
   
   'Modify By Sindy 2011/2/24 增加SP65,SP66,LC43,LC44,LC45,LC46,HC24,HC25,HC26,HC27
   strExc(0) = strExc(0) & _
      " OR EXISTS(SELECT * FROM SERVICEPRACTICE,CASEPROGRESS WHERE SP08=CU01||CU02 AND SP15 IS NULL AND SP61 IS NULL AND SUBSTR(SP01,1,1)<>'F'" & _
      " AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 AND CP05>20050000 AND CP09<'B')" & _
      " OR EXISTS(SELECT * FROM SERVICEPRACTICE,CASEPROGRESS WHERE SP58=CU01||CU02 AND SP15 IS NULL AND SP61 IS NULL AND SUBSTR(SP01,1,1)<>'F'" & _
      " AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 AND CP05>20050000 AND CP09<'B')" & _
      " OR EXISTS(SELECT * FROM SERVICEPRACTICE,CASEPROGRESS WHERE SP59=CU01||CU02 AND SP15 IS NULL AND SP61 IS NULL AND SUBSTR(SP01,1,1)<>'F'" & _
      " AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 AND CP05>20050000 AND CP09<'B')" & _
      " OR EXISTS(SELECT * FROM SERVICEPRACTICE,CASEPROGRESS WHERE SP65=CU01||CU02 AND SP15 IS NULL AND SP61 IS NULL AND SUBSTR(SP01,1,1)<>'F'" & _
      " AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 AND CP05>20050000 AND CP09<'B')" & _
      " OR EXISTS(SELECT * FROM SERVICEPRACTICE,CASEPROGRESS WHERE SP66=CU01||CU02 AND SP15 IS NULL AND SP61 IS NULL AND SUBSTR(SP01,1,1)<>'F'" & _
      " AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 AND CP05>20050000 AND CP09<'B')"
   strExc(0) = strExc(0) & _
      " OR EXISTS(SELECT * FROM LAWCASE,CASEPROGRESS WHERE LC11=CU01||CU02 AND LC08 IS NULL AND LC34 IS NULL AND SUBSTR(LC01,1,1)<>'F'" & _
      " AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 AND CP05>20030000 AND CP09<'B')" & _
      " OR EXISTS(SELECT * FROM LAWCASE,CASEPROGRESS WHERE LC43=CU01||CU02 AND LC08 IS NULL AND LC34 IS NULL AND SUBSTR(LC01,1,1)<>'F'" & _
      " AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 AND CP05>20030000 AND CP09<'B')" & _
      " OR EXISTS(SELECT * FROM LAWCASE,CASEPROGRESS WHERE LC44=CU01||CU02 AND LC08 IS NULL AND LC34 IS NULL AND SUBSTR(LC01,1,1)<>'F'" & _
      " AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 AND CP05>20030000 AND CP09<'B')" & _
      " OR EXISTS(SELECT * FROM LAWCASE,CASEPROGRESS WHERE LC45=CU01||CU02 AND LC08 IS NULL AND LC34 IS NULL AND SUBSTR(LC01,1,1)<>'F'" & _
      " AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 AND CP05>20030000 AND CP09<'B')" & _
      " OR EXISTS(SELECT * FROM LAWCASE,CASEPROGRESS WHERE LC46=CU01||CU02 AND LC08 IS NULL AND LC34 IS NULL AND SUBSTR(LC01,1,1)<>'F'" & _
      " AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 AND CP05>20030000 AND CP09<'B')"
   strExc(0) = strExc(0) & _
      " OR EXISTS(SELECT * FROM HIRECASE,CASEPROGRESS WHERE HC05=CU01||CU02 AND HC09 IS NULL AND HC19 IS NULL AND SUBSTR(HC01,1,1)<>'F'" & _
      " AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 AND CP04(+)=HC04 AND CP05>20050000 AND CP09<'B')" & _
      " OR EXISTS(SELECT * FROM HIRECASE,CASEPROGRESS WHERE HC24=CU01||CU02 AND HC09 IS NULL AND HC19 IS NULL AND SUBSTR(HC01,1,1)<>'F'" & _
      " AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 AND CP04(+)=HC04 AND CP05>20050000 AND CP09<'B')" & _
      " OR EXISTS(SELECT * FROM HIRECASE,CASEPROGRESS WHERE HC25=CU01||CU02 AND HC09 IS NULL AND HC19 IS NULL AND SUBSTR(HC01,1,1)<>'F'" & _
      " AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 AND CP04(+)=HC04 AND CP05>20050000 AND CP09<'B')" & _
      " OR EXISTS(SELECT * FROM HIRECASE,CASEPROGRESS WHERE HC26=CU01||CU02 AND HC09 IS NULL AND HC19 IS NULL AND SUBSTR(HC01,1,1)<>'F'" & _
      " AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 AND CP04(+)=HC04 AND CP05>20050000 AND CP09<'B')" & _
      " OR EXISTS(SELECT * FROM HIRECASE,CASEPROGRESS WHERE HC27=CU01||CU02 AND HC09 IS NULL AND HC19 IS NULL AND SUBSTR(HC01,1,1)<>'F'" & _
      " AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 AND CP04(+)=HC04 AND CP05>20050000 AND CP09<'B'))"

   If p_iReportID > 0 Then

   strExc(1) = CompDate(0, -3, strSrvDate(1))
   Select Case p_iReportID
      Case 1, 3, 5 '3年內有A類收文
         strExc(0) = strExc(0) & " AND (" & _
            " EXISTS (SELECT * FROM TRADEMARK,CASEPROGRESS WHERE TM23=CU01||CU02 AND TM29 IS NULL AND TM57 IS NULL" & _
            " AND CP01(+)=TM01 AND CP02(+)=TM02 AND CP03(+)=TM03 AND CP04(+)=TM04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM TRADEMARK,CASEPROGRESS WHERE TM78=CU01||CU02 AND TM29 IS NULL AND TM57 IS NULL" & _
            " AND CP01(+)=TM01 AND CP02(+)=TM02 AND CP03(+)=TM03 AND CP04(+)=TM04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM TRADEMARK,CASEPROGRESS WHERE TM79=CU01||CU02 AND TM29 IS NULL AND TM57 IS NULL" & _
            " AND CP01(+)=TM01 AND CP02(+)=TM02 AND CP03(+)=TM03 AND CP04(+)=TM04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM TRADEMARK,CASEPROGRESS WHERE TM80=CU01||CU02 AND TM29 IS NULL AND TM57 IS NULL" & _
            " AND CP01(+)=TM01 AND CP02(+)=TM02 AND CP03(+)=TM03 AND CP04(+)=TM04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM TRADEMARK,CASEPROGRESS WHERE TM81=CU01||CU02 AND TM29 IS NULL AND TM57 IS NULL" & _
            " AND CP01(+)=TM01 AND CP02(+)=TM02 AND CP03(+)=TM03 AND CP04(+)=TM04 AND CP05>" & strExc(1) & " AND CP09<'B')"
         strExc(0) = strExc(0) & _
            " OR EXISTS(SELECT * FROM PATENT,CASEPROGRESS WHERE PA26=CU01||CU02 AND PA57 IS NULL AND PA108 IS NULL" & _
            " AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM PATENT,CASEPROGRESS WHERE PA27=CU01||CU02 AND PA57 IS NULL AND PA108 IS NULL" & _
            " AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM PATENT,CASEPROGRESS WHERE PA28=CU01||CU02 AND PA57 IS NULL AND PA108 IS NULL" & _
            " AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM PATENT,CASEPROGRESS WHERE PA29=CU01||CU02 AND PA57 IS NULL AND PA108 IS NULL" & _
            " AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM PATENT,CASEPROGRESS WHERE PA30=CU01||CU02 AND PA57 IS NULL AND PA108 IS NULL" & _
            " AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 AND CP05>" & strExc(1) & " AND CP09<'B')"
         'Modify By Sindy 2011/2/24 增加SP65,SP66,LC43,LC44,LC45,LC46,HC24,HC25,HC26,HC27
         strExc(0) = strExc(0) & _
            " OR EXISTS(SELECT * FROM SERVICEPRACTICE,CASEPROGRESS WHERE SP08=CU01||CU02 AND SP15 IS NULL AND SP61 IS NULL" & _
            " AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM SERVICEPRACTICE,CASEPROGRESS WHERE SP58=CU01||CU02 AND SP15 IS NULL AND SP61 IS NULL" & _
            " AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM SERVICEPRACTICE,CASEPROGRESS WHERE SP59=CU01||CU02 AND SP15 IS NULL AND SP61 IS NULL" & _
            " AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM SERVICEPRACTICE,CASEPROGRESS WHERE SP65=CU01||CU02 AND SP15 IS NULL AND SP61 IS NULL" & _
            " AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM SERVICEPRACTICE,CASEPROGRESS WHERE SP66=CU01||CU02 AND SP15 IS NULL AND SP61 IS NULL" & _
            " AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 AND CP05>" & strExc(1) & " AND CP09<'B')"
         strExc(0) = strExc(0) & _
            " OR EXISTS(SELECT * FROM LAWCASE,CASEPROGRESS WHERE LC11=CU01||CU02 AND LC08 IS NULL AND LC34 IS NULL" & _
            " AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM LAWCASE,CASEPROGRESS WHERE LC43=CU01||CU02 AND LC08 IS NULL AND LC34 IS NULL" & _
            " AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM LAWCASE,CASEPROGRESS WHERE LC44=CU01||CU02 AND LC08 IS NULL AND LC34 IS NULL" & _
            " AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM LAWCASE,CASEPROGRESS WHERE LC45=CU01||CU02 AND LC08 IS NULL AND LC34 IS NULL" & _
            " AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM LAWCASE,CASEPROGRESS WHERE LC46=CU01||CU02 AND LC08 IS NULL AND LC34 IS NULL" & _
            " AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 AND CP05>" & strExc(1) & " AND CP09<'B')"
         strExc(0) = strExc(0) & _
            " OR EXISTS(SELECT * FROM HIRECASE,CASEPROGRESS WHERE HC05=CU01||CU02 AND HC09 IS NULL AND HC19 IS NULL" & _
            " AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 AND CP04(+)=HC04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM HIRECASE,CASEPROGRESS WHERE HC24=CU01||CU02 AND HC09 IS NULL AND HC19 IS NULL" & _
            " AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 AND CP04(+)=HC04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM HIRECASE,CASEPROGRESS WHERE HC25=CU01||CU02 AND HC09 IS NULL AND HC19 IS NULL" & _
            " AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 AND CP04(+)=HC04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM HIRECASE,CASEPROGRESS WHERE HC26=CU01||CU02 AND HC09 IS NULL AND HC19 IS NULL" & _
            " AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 AND CP04(+)=HC04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM HIRECASE,CASEPROGRESS WHERE HC27=CU01||CU02 AND HC09 IS NULL AND HC19 IS NULL" & _
            " AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 AND CP04(+)=HC04 AND CP05>" & strExc(1) & " AND CP09<'B'))"

      Case 2, 4, 6 '3年內無A類收文
         strExc(0) = strExc(0) & " AND NOT (" & _
            " EXISTS (SELECT * FROM TRADEMARK,CASEPROGRESS WHERE TM23=CU01||CU02 AND TM29 IS NULL AND TM57 IS NULL" & _
            " AND CP01(+)=TM01 AND CP02(+)=TM02 AND CP03(+)=TM03 AND CP04(+)=TM04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM TRADEMARK,CASEPROGRESS WHERE TM78=CU01||CU02 AND TM29 IS NULL AND TM57 IS NULL" & _
            " AND CP01(+)=TM01 AND CP02(+)=TM02 AND CP03(+)=TM03 AND CP04(+)=TM04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM TRADEMARK,CASEPROGRESS WHERE TM79=CU01||CU02 AND TM29 IS NULL AND TM57 IS NULL" & _
            " AND CP01(+)=TM01 AND CP02(+)=TM02 AND CP03(+)=TM03 AND CP04(+)=TM04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM TRADEMARK,CASEPROGRESS WHERE TM80=CU01||CU02 AND TM29 IS NULL AND TM57 IS NULL" & _
            " AND CP01(+)=TM01 AND CP02(+)=TM02 AND CP03(+)=TM03 AND CP04(+)=TM04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM TRADEMARK,CASEPROGRESS WHERE TM81=CU01||CU02 AND TM29 IS NULL AND TM57 IS NULL" & _
            " AND CP01(+)=TM01 AND CP02(+)=TM02 AND CP03(+)=TM03 AND CP04(+)=TM04 AND CP05>" & strExc(1) & " AND CP09<'B')"
         strExc(0) = strExc(0) & _
            " OR EXISTS(SELECT * FROM PATENT,CASEPROGRESS WHERE PA26=CU01||CU02 AND PA57 IS NULL AND PA108 IS NULL" & _
            " AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM PATENT,CASEPROGRESS WHERE PA27=CU01||CU02 AND PA57 IS NULL AND PA108 IS NULL" & _
            " AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM PATENT,CASEPROGRESS WHERE PA28=CU01||CU02 AND PA57 IS NULL AND PA108 IS NULL" & _
            " AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM PATENT,CASEPROGRESS WHERE PA29=CU01||CU02 AND PA57 IS NULL AND PA108 IS NULL" & _
            " AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM PATENT,CASEPROGRESS WHERE PA30=CU01||CU02 AND PA57 IS NULL AND PA108 IS NULL" & _
            " AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 AND CP05>" & strExc(1) & " AND CP09<'B')"
         'Modify By Sindy 2011/2/24 增加SP65,SP66,LC43,LC44,LC45,LC46,HC24,HC25,HC26,HC27
         strExc(0) = strExc(0) & _
            " OR EXISTS(SELECT * FROM SERVICEPRACTICE,CASEPROGRESS WHERE SP08=CU01||CU02 AND SP15 IS NULL AND SP61 IS NULL" & _
            " AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM SERVICEPRACTICE,CASEPROGRESS WHERE SP58=CU01||CU02 AND SP15 IS NULL AND SP61 IS NULL" & _
            " AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM SERVICEPRACTICE,CASEPROGRESS WHERE SP59=CU01||CU02 AND SP15 IS NULL AND SP61 IS NULL" & _
            " AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM SERVICEPRACTICE,CASEPROGRESS WHERE SP65=CU01||CU02 AND SP15 IS NULL AND SP61 IS NULL" & _
            " AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM SERVICEPRACTICE,CASEPROGRESS WHERE SP66=CU01||CU02 AND SP15 IS NULL AND SP61 IS NULL" & _
            " AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 AND CP05>" & strExc(1) & " AND CP09<'B')"
         strExc(0) = strExc(0) & _
            " OR EXISTS(SELECT * FROM LAWCASE,CASEPROGRESS WHERE LC11=CU01||CU02 AND LC08 IS NULL AND LC34 IS NULL" & _
            " AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM LAWCASE,CASEPROGRESS WHERE LC43=CU01||CU02 AND LC08 IS NULL AND LC34 IS NULL" & _
            " AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM LAWCASE,CASEPROGRESS WHERE LC44=CU01||CU02 AND LC08 IS NULL AND LC34 IS NULL" & _
            " AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM LAWCASE,CASEPROGRESS WHERE LC45=CU01||CU02 AND LC08 IS NULL AND LC34 IS NULL" & _
            " AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM LAWCASE,CASEPROGRESS WHERE LC46=CU01||CU02 AND LC08 IS NULL AND LC34 IS NULL" & _
            " AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 AND CP05>" & strExc(1) & " AND CP09<'B')"
         strExc(0) = strExc(0) & _
            " OR EXISTS(SELECT * FROM HIRECASE,CASEPROGRESS WHERE HC05=CU01||CU02 AND HC09 IS NULL AND HC19 IS NULL" & _
            " AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 AND CP04(+)=HC04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM HIRECASE,CASEPROGRESS WHERE HC24=CU01||CU02 AND HC09 IS NULL AND HC19 IS NULL" & _
            " AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 AND CP04(+)=HC04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM HIRECASE,CASEPROGRESS WHERE HC25=CU01||CU02 AND HC09 IS NULL AND HC19 IS NULL" & _
            " AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 AND CP04(+)=HC04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM HIRECASE,CASEPROGRESS WHERE HC26=CU01||CU02 AND HC09 IS NULL AND HC19 IS NULL" & _
            " AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 AND CP04(+)=HC04 AND CP05>" & strExc(1) & " AND CP09<'B')" & _
            " OR EXISTS(SELECT * FROM HIRECASE,CASEPROGRESS WHERE HC27=CU01||CU02 AND HC09 IS NULL AND HC19 IS NULL" & _
            " AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 AND CP04(+)=HC04 AND CP05>" & strExc(1) & " AND CP09<'B'))"
   End Select
   
   End If
   If p_iReportID = 0 Then
         strExc(0) = strExc(0) & " ORDER BY 1,2"
   Else
      strExc(0) = strExc(0) & " ORDER BY ST03,CU13,CU01,CU02"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If p_iReportID = 0 Then
      Set p_Rst = RsTemp.Clone
      Exit Function
   End If
   If intI = 1 Then
      If Not IsNull(RsTemp.Fields(10).Value) Then
         strNo = RsTemp.Fields(10).Value
         strName = "" & RsTemp.Fields("ST02").Value
         strDept = "" & RsTemp.Fields("ST03").Value
      End If
      
      Screen.MousePointer = vbHourglass
      GetPrintLeft
      Page = 1
      CaseTitle Page, "" & RsTemp.Fields(10).Value, "" & RsTemp.Fields(11).Value, p_iReportID, strDept
      iPrint = 2700 + 300 + 300
      i = 0
      With RsTemp
         i = 0
         Do While Not .EOF
            For j = 0 To 4
               Printer.CurrentX = PLeft(j)
               Printer.CurrentY = iPrint
               If j = 0 Then
                  If Not IsNull(RsTemp.Fields(9).Value) Then
                     Printer.CurrentX = PLeft(j) - Printer.TextWidth("＊")
                     If Not IsNull(RsTemp.Fields(8).Value) Then
                        Printer.Print "＊" & Left(.Fields(j) & "000", 9) & " N"
                     Else
                        Printer.Print "＊" & Left(.Fields(j) & "000", 9)
                     End If
                  Else
                     If Not IsNull(RsTemp.Fields(8).Value) Then
                        Printer.Print " " & Left(.Fields(j) & "000", 9) & " N"
                     Else
                        Printer.Print "" & Left(.Fields(j) & "000", 9)
                     End If
                  End If
               Else
                  Printer.Print "" & .Fields(j)
               End If
            Next j
            
            iPrint = iPrint + 300
            
            For j = 5 To 7
               If Me.Text1(0).Text = "1" And j = 7 Then
                  Printer.CurrentX = PLeft(8)
                  Printer.CurrentY = iPrint
                  Printer.Print "" & .Fields(15)
               End If
               Printer.CurrentX = PLeft(j)
               Printer.CurrentY = iPrint
               Printer.Print "" & .Fields(j)
            Next j
            
            iPrint = iPrint + 300
            Printer.CurrentX = PLeft(5)
            Printer.CurrentY = iPrint
            Printer.Print "" & .Fields(9)
            Printer.CurrentX = PLeft(6)
            Printer.CurrentY = iPrint
            Printer.Print "" & .Fields(13)
            Printer.CurrentX = PLeft(8)
            Printer.CurrentY = iPrint
            Printer.Print "" & .Fields(16)
            Printer.CurrentX = PLeft(7)
            Printer.CurrentY = iPrint
            Printer.Print "" & .Fields(14)
            iPrint = iPrint + 300
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print String(250, "-")
            iPrint = iPrint + 300
            
            i = i + 1
            .MoveNext
            If RsTemp.EOF Then Exit Do
            If i > 9 Or "" & RsTemp.Fields(10).Value <> strNo Then
               Printer.CurrentX = PLeft(0)
               Printer.CurrentY = iPrint
               Printer.Print "PS：1.編號與公司名稱之間, 若有 N 表示不寄台一雜誌"
               Printer.CurrentX = PLeft(0) + Printer.TextWidth("PS：")
               Printer.CurrentY = iPrint + 300
               Printer.FontSize = 14
               Printer.FontBold = True
               Printer.Print "2.核對後請於97/7/2交各區主管。主管於97/7/4收齊交各所管理部更改資料。"
               Printer.FontBold = False
               Printer.FontSize = 10
               Printer.NewPage
               Page = Page + 1
               If "" & RsTemp.Fields(10).Value <> strNo Then
                  strNo = RsTemp.Fields(10).Value
                  strName = "" & RsTemp.Fields("ST02").Value
                  strDept = "" & RsTemp.Fields("ST03").Value
               End If
               CaseTitle Page, "" & RsTemp.Fields(10).Value, "" & RsTemp.Fields(11).Value, p_iReportID, strDept
               iPrint = 2700 + 300 + 300
               i = 0
            End If
         Loop
      End With
      
      Printer.CurrentX = PLeft(0)
      Printer.CurrentY = iPrint
      Printer.Print "PS：1.編號與公司名稱之間, 若有 N 表示不寄台一雜誌"
      Printer.CurrentX = PLeft(0) + Printer.TextWidth("PS：")
      Printer.CurrentY = iPrint + 300
      Printer.FontSize = 14
      Printer.FontBold = True
      Printer.Print "2.核對後請於97/7/2交各區主管。主管於97/7/4收齊交各所管理部更改資料。"
      Printer.FontBold = False
      Printer.FontSize = 10
      Printer.EndDoc
      PrintCase1 = Page
      Screen.MousePointer = vbDefault
   End If
   Exit Function
ErrHand:
   MsgBox Err.Description
End Function
'Add by Morgan 2008/6/17
Private Sub PrintCaseX()
   Dim iNo As Integer, adoSales As ADODB.Recordset
   Dim ffa As Integer, iPage As Integer
   Dim strDept As String
   Dim strNo As String '員工代號
   Dim strName As String
   Dim strDesc As String
   
   PrintCase1 , , adoSales
   If adoSales.RecordCount > 0 Then
      ffa = FreeFile
      strExc(1) = PUB_Getdesktop
      Open strExc(1) & "\業務清單(" & strSrvDate(2) & ").TXT" For Output As ffa
      With adoSales
      Do While Not .EOF
         strDept = "" & .Fields(0)
         strNo = "" & .Fields(1)
         strName = "" & .Fields(2)
         strDesc = ""
         For iNo = 1 To 6
            iPage = PrintCase1(iNo, strNo)
            If iPage > 0 Then strDesc = strDesc & "," & iNo
         Next
         Print #ffa, strDept & " " & strNo & " " & strName & " " & Mid(strDesc, 2)
         .MoveNext
      Loop
      End With
      Close ffa
      MsgBox "列印完成，清單已存於桌面!!", , "列印成功"
   End If
   Set adoSales = Nothing
End Sub

'2009/2/13 add by sonia e-mail核對名冊,抓95年起有收文之客戶
Private Function PrintCase2()
 Dim i As Integer, j As Integer, Page As Integer, iPrint As Integer
 Dim strTmp As String, strTmp1 As String
 Dim strNo As String '員工代號
 Dim strName As String
 Dim strDept As String

On Error GoTo ErrHand
      
   strExc(1) = 20060000
   'Modify by Morgan 2009/11/30 不用再考慮舊的聯絡人欄位
   'strExc(0) = "SELECT CU01||CU02,SUBSTR(CU04,1,30),SUBSTR(CU07,1,10),SUBSTR(NA03,1,14),CU16,NVL(PCC21,CU30) CU30,SUBSTR(NVL(NVL(PCC22,CU31),CU23),1,65),CU18,CU32,CU80,CU13,ST02,CU12,CU23,CU11,NVL(PCC05,CU08) CU08,CU20,ST03 FROM CUSTOMER,POTCUSTCONT,NATION,STAFF,( "
   strExc(0) = "SELECT CU01||CU02,SUBSTR(CU04,1,30),SUBSTR(CU07,1,10),SUBSTR(NA03,1,14),CU16,NVL(PCC21,CU30) CU30,SUBSTR(NVL(NVL(PCC22,CU31),CU23),1,65),CU18,CU32,CU80,CU13,ST02,CU12,CU23,CU11,PCC05 CU08,CU20,ST03 FROM CUSTOMER,POTCUSTCONT,NATION,STAFF,( "
   'Modify By Sindy 2011/2/24 增加SP65,SP66,LC43,LC44,LC45,LC46,HC24,HC25,HC26,HC27
   strExc(0) = strExc(0) & "select distinct pa26 CUNO from patent,caseprogress where cp01 in ('P','CFP','FCP') AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=PA01 AND CP02=PA02 AND CP03=PA03 AND CP04=PA04 AND PA26 IS NOT NULL UNION " & _
              "select distinct pa27 CUNO from patent,caseprogress where cp01 in ('P','CFP','FCP') AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=PA01 AND CP02=PA02 AND CP03=PA03 AND CP04=PA04 AND PA27 IS NOT NULL UNION " & _
              "select distinct pa28 CUNO from patent,caseprogress where cp01 in ('P','CFP','FCP') AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=PA01 AND CP02=PA02 AND CP03=PA03 AND CP04=PA04 AND PA28 IS NOT NULL UNION " & _
              "select distinct pa29 CUNO from patent,caseprogress where cp01 in ('P','CFP','FCP') AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=PA01 AND CP02=PA02 AND CP03=PA03 AND CP04=PA04 AND PA29 IS NOT NULL UNION " & _
              "select distinct pa30 CUNO from patent,caseprogress where cp01 in ('P','CFP','FCP') AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=PA01 AND CP02=PA02 AND CP03=PA03 AND CP04=PA04 AND PA30 IS NOT NULL UNION " & _
              "select distinct TM23 CUNO from TRADEMARK,caseprogress where cp01 in ('T','CFT','FCT','TF') AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=TM01 AND CP02=TM02 AND CP03=TM03 AND CP04=TM04 AND TM23 IS NOT NULL UNION " & _
              "select distinct TM78 CUNO from TRADEMARK,caseprogress where cp01 in ('T','CFT','FCT','TF') AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=TM01 AND CP02=TM02 AND CP03=TM03 AND CP04=TM04 AND TM78 IS NOT NULL UNION " & _
              "select distinct TM79 CUNO from TRADEMARK,caseprogress where cp01 in ('T','CFT','FCT','TF') AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=TM01 AND CP02=TM02 AND CP03=TM03 AND CP04=TM04 AND TM79 IS NOT NULL UNION " & _
              "select distinct TM80 CUNO from TRADEMARK,caseprogress where cp01 in ('T','CFT','FCT','TF') AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=TM01 AND CP02=TM02 AND CP03=TM03 AND CP04=TM04 AND TM80 IS NOT NULL UNION " & _
              "select distinct TM81 CUNO from TRADEMARK,caseprogress where cp01 in ('T','CFT','FCT','TF') AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=TM01 AND CP02=TM02 AND CP03=TM03 AND CP04=TM04 AND TM81 IS NOT NULL UNION " & _
              "select distinct SP08 CUNO from SERVICEPRACTICE,caseprogress where cp01 NOT in ('T','CFT','FCT','TF','P','CFP','FCP','LA','L','FCL','CFL') AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=SP01 AND CP02=SP02 AND CP03=SP03 AND CP04=SP04 AND SP08 IS NOT NULL UNION " & _
              "select distinct SP58 CUNO from SERVICEPRACTICE,caseprogress where cp01 NOT in ('T','CFT','FCT','TF','P','CFP','FCP','LA','L','FCL','CFL') AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=SP01 AND CP02=SP02 AND CP03=SP03 AND CP04=SP04 AND SP58 IS NOT NULL UNION " & _
              "select distinct SP59 CUNO from SERVICEPRACTICE,caseprogress where cp01 NOT in ('T','CFT','FCT','TF','P','CFP','FCP','LA','L','FCL','CFL') AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=SP01 AND CP02=SP02 AND CP03=SP03 AND CP04=SP04 AND SP59 IS NOT NULL UNION " & _
              "select distinct SP65 CUNO from SERVICEPRACTICE,caseprogress where cp01 NOT in ('T','CFT','FCT','TF','P','CFP','FCP','LA','L','FCL','CFL') AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=SP01 AND CP02=SP02 AND CP03=SP03 AND CP04=SP04 AND SP65 IS NOT NULL UNION " & _
              "select distinct SP66 CUNO from SERVICEPRACTICE,caseprogress where cp01 NOT in ('T','CFT','FCT','TF','P','CFP','FCP','LA','L','FCL','CFL') AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=SP01 AND CP02=SP02 AND CP03=SP03 AND CP04=SP04 AND SP66 IS NOT NULL UNION " & _
              "select distinct LC11 CUNO from LAWCASE,caseprogress where cp01 in ('L','CFL','FCL') AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 AND LC11 IS NOT NULL UNION " & _
              "select distinct LC43 CUNO from LAWCASE,caseprogress where cp01 in ('L','CFL','FCL') AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 AND LC43 IS NOT NULL UNION " & _
              "select distinct LC44 CUNO from LAWCASE,caseprogress where cp01 in ('L','CFL','FCL') AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 AND LC44 IS NOT NULL UNION " & _
              "select distinct LC45 CUNO from LAWCASE,caseprogress where cp01 in ('L','CFL','FCL') AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 AND LC45 IS NOT NULL UNION " & _
              "select distinct LC46 CUNO from LAWCASE,caseprogress where cp01 in ('L','CFL','FCL') AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 AND LC46 IS NOT NULL UNION " & _
              "select distinct HC05 CUNO from HIRECASE,caseprogress where cp01='LA' AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 AND HC05 IS NOT NULL UNION " & _
              "select distinct HC24 CUNO from HIRECASE,caseprogress where cp01='LA' AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 AND HC24 IS NOT NULL UNION " & _
              "select distinct HC25 CUNO from HIRECASE,caseprogress where cp01='LA' AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 AND HC25 IS NOT NULL UNION " & _
              "select distinct HC26 CUNO from HIRECASE,caseprogress where cp01='LA' AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 AND HC26 IS NOT NULL UNION " & _
              "select distinct HC27 CUNO from HIRECASE,caseprogress where cp01='LA' AND CP09<'B' AND CP05>=" & strExc(1) & " AND CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 AND HC27 IS NOT NULL ) C "
   
   strExc(0) = strExc(0) & strSql & " AND SUBSTR(C.CUNO,1,8)=CU01(+) AND SUBSTR(C.CUNO,9,1)=CU02(+) AND CU10=NA01(+) AND CU13=ST01(+) AND SUBSTR(CU12,1,1)<>'F' AND PCC01(+)=CU01 AND PCC02(+)=CU127 "
   '2013/5/6 modify by sonia 再加其他
   'modify by sonia 2023/2/23 再加不得代理專利,不得代理商標,解除對造,國內同業
   'Modify By Sindy 2025/6/27 +or cu80='業務自行處理' or cu80='其他' or cu80='不得代理專利' or cu80='不得代理商標' or cu80='解除對造' or cu80='國內同業'
   '                          改抓常變數
   strExc(0) = strExc(0) & " and (cu80 is null or instr('" & 客戶及代理人可讀取的狀態 & "',cu80)>0)"
   strExc(0) = strExc(0) & " ORDER BY ST03,CU13,CU01,CU02"
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(RsTemp.Fields(10).Value) Then
         strNo = RsTemp.Fields(10).Value
         strName = "" & RsTemp.Fields("ST02").Value
         strDept = "" & RsTemp.Fields("ST03").Value
      End If
      
      Screen.MousePointer = vbHourglass
      GetPrintLeft
      Page = 1
      CaseTitle Page, "" & RsTemp.Fields(10).Value, "" & RsTemp.Fields(11).Value, 7, strDept
      iPrint = 2700 + 300 + 300
      i = 0
      With RsTemp
         i = 0
         Do While Not .EOF
            For j = 0 To 4
               Printer.CurrentX = PLeft(j)
               Printer.CurrentY = iPrint
               If j = 0 Then
                  If Not IsNull(RsTemp.Fields(9).Value) Then
                     Printer.CurrentX = PLeft(j) - Printer.TextWidth("＊")
                     If Not IsNull(RsTemp.Fields(8).Value) Then
                        Printer.Print "＊" & Left(.Fields(j) & "000", 9) & " N"
                     Else
                        Printer.Print "＊" & Left(.Fields(j) & "000", 9)
                     End If
                  Else
                     If Not IsNull(RsTemp.Fields(8).Value) Then
                        Printer.Print " " & Left(.Fields(j) & "000", 9) & " N"
                     Else
                        Printer.Print "" & Left(.Fields(j) & "000", 9)
                     End If
                  End If
               Else
                  Printer.Print "" & .Fields(j)
               End If
            Next j
            
            iPrint = iPrint + 300
            
            For j = 5 To 7
               If Me.Text1(0).Text = "1" And j = 7 Then
                  Printer.CurrentX = PLeft(8)
                  Printer.CurrentY = iPrint
                  Printer.Print "" & .Fields(15)
               End If
               Printer.CurrentX = PLeft(j)
               Printer.CurrentY = iPrint
               Printer.Print "" & .Fields(j)
            Next j
            
            iPrint = iPrint + 300
            Printer.CurrentX = PLeft(5)
            Printer.CurrentY = iPrint
            Printer.Print "" & .Fields(9)
            Printer.CurrentX = PLeft(6)
            Printer.CurrentY = iPrint
            Printer.Print "" & .Fields(13)
            Printer.CurrentX = PLeft(8)
            Printer.CurrentY = iPrint
            Printer.Print "" & .Fields(16)
            iPrint = iPrint + 300
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print String(250, "-")
            iPrint = iPrint + 300
            
            i = i + 1
            .MoveNext
            If RsTemp.EOF Then Exit Do
            If i > 9 Or "" & RsTemp.Fields(10).Value <> strNo Then
               Printer.CurrentX = PLeft(0)
               Printer.CurrentY = iPrint
               Printer.Print "PS：編號與公司名稱之間, 若有 N 表示不寄台一雜誌"
               Printer.NewPage
               Page = Page + 1
               If "" & RsTemp.Fields(10).Value <> strNo Then
                  strNo = RsTemp.Fields(10).Value
                  strName = "" & RsTemp.Fields("ST02").Value
                  strDept = "" & RsTemp.Fields("ST03").Value
               End If
               CaseTitle Page, "" & RsTemp.Fields(10).Value, "" & RsTemp.Fields(11).Value, 7, strDept
               iPrint = 2700 + 300 + 300
               i = 0
            End If
         Loop
      End With
      
      Printer.CurrentX = PLeft(0)
      Printer.CurrentY = iPrint
      Printer.Print "PS：編號與公司名稱之間, 若有 N 表示不寄台一雜誌"
      Printer.EndDoc
      Screen.MousePointer = vbDefault
   End If
   Exit Function
ErrHand:
   MsgBox Err.Description
End Function

Private Sub PrintLetterBatch(pstrSql)
   Dim rsTemp1 As ADODB.Recordset
   Dim nPageNo As Integer
   
   'Modify by Morgan 2008/7/25 接洽人改先用聯絡人編號抓聯絡人檔
   'strExc(0) = "SELECT NVL(NVL(CU104,CU04),CU05||CU88||CU89) C00, NVL(NVL(NVL(CU08,CU104),CU04),CU05||CU88||CU89) C01, DECODE(CU15,'0','台端','貴公司') C02, ST02 C03 FROM CUSTOMER,STAFF " & strSQL & " AND CU02='0' AND ST01(+)=CU13 ORDER BY CU12,CU13,CU01,CU02"
   'Modify by Morgan 2009/11/30 不用再考慮舊的聯絡人欄位
   'strExc(0) = "SELECT NVL(NVL(CU104,CU04),CU05||CU88||CU89) C00, NVL(NVL(NVL(NVL(PCC05,CU08),CU104),CU04),CU05||CU88||CU89) C01, DECODE(CU15,'0','台端','貴公司') C02, ST02 C03 FROM CUSTOMER,POTCUSTCONT,STAFF " & strSQL & " AND CU02='0' AND ST01(+)=CU13 AND PCC01(+)=CU01 AND PCC02(+)=CU127 ORDER BY CU12,CU13,CU01,CU02"
   'Modify By Sindy 2012/5/24 DECODE(CU15,'0','台端','貴公司')==>DECODE(CU15,'0','台端','1','貴公司','貴單位')
   strExc(0) = "SELECT NVL(NVL(CU104,CU04),CU05||CU88||CU89) C00, NVL(NVL(NVL(PCC05,CU104),CU04),CU05||CU88||CU89) C01, DECODE(CU15,'0','台端','1','貴公司','貴單位') C02, ST02 C03 FROM CUSTOMER,POTCUSTCONT,STAFF " & pstrSql & " AND CU02='0' AND ST01(+)=CU13 AND PCC01(+)=CU01 AND PCC02(+)=CU127 ORDER BY CU12,CU13,CU01,CU02"
   intI = 1
   'edit by nickc 2007/02/09 不用 dll 了
   'Set rsTemp1 = objLawDll.ReadRstMsg(intI, strExc(0))
   Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With rsTemp1
         nPageNo = 0
         Printer.Orientation = vbPRORPortrait '橫印
         Printer.Font = "標楷體"
         Printer.FontSize = 12
         .MoveFirst
         Do While Not .EOF
            nPageNo = nPageNo + 1
            If nPageNo > 1 Then Printer.NewPage
            PrintLetter "" & .Fields(0), "" & .Fields(1), "" & .Fields(2), "" & .Fields(3), nPageNo
            .MoveNext
         Loop
         Printer.EndDoc
      End With
      MsgBox "列印結束 !", vbInformation
   Else
      MsgBox "無符合條件之資料可列印 !", vbInformation
   End If
   Set rsTemp1 = Nothing
End Sub

'Added by Morgan 2012/4/27
'改用 Word 列印,內容也改抓公用定稿以便於修改
Private Sub PrintLetterBatch2(pstrSql As String)
   Dim rsTemp1 As ADODB.Recordset
   Dim strContent As String, strContent2 As String
   Dim bVisible As Boolean
   Dim iPicNo As Integer, stFileName As String
   Dim iPicNo2 As Integer
   Dim oShape
   Dim bolRetry As Boolean
   
   Dim strCustomer As String
   Dim stReceiver As String '收件人
   Dim stAddr As String '地址
   Dim stZip As String '郵遞區號
   Dim stContact As String '接洽人
   Dim stApp1Title As String '稱謂
   Dim iLineCount As Integer '行數
   Dim strPath As String
   Dim iSNo As Integer
   
   strPath = PUB_Getdesktop
   'Add by Amy 2021/03/08 建資料夾
   If Dir(PUB_Getdesktop & "\tmp", vbDirectory) = MsgText(601) Then
        MkDir strPath & "\tmp"
   End If
   strPath = strPath & "\tmp"
   
   'Modify By Sindy 2012/5/24 DECODE(CU15,'0','台端','貴公司')==>DECODE(CU15,'0','台端','1','貴公司','貴單位')
   strExc(0) = "SELECT NVL(NVL(CU104,CU04),CU05||CU88||CU89) C00, NVL(NVL(NVL(PCC05,CU104),CU04),CU05||CU88||CU89) C01, DECODE(CU15,'0','台端','1','貴公司','貴單位') C02, ST02 C03,CU01||CU02 CuNo FROM CUSTOMER,POTCUSTCONT,STAFF " & pstrSql & " AND CU02='0' AND ST01(+)=CU13 AND PCC01(+)=CU01 AND PCC02(+)=CU127 ORDER BY CU12,CU13,CU01,CU02"
   intI = 1
   Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strExc(0) = "SELECT FTM05,FTM08 FROM FINALTEXTMAP WHERE FTM01='P' AND FTM02='22' AND FTM03='000' AND FTM04='11'"
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strContent = RsTemp(0) & RsTemp(1)
      Else
         MsgBox "無法取得定稿內容!!", vbCritical
         Exit Sub
      End If
      
      'Added by Morgan 2020/3/31
      If strSrvDate(1) >= 智慧所更名日 Then
         PUB_GetLetterPicID "2", "P", iPicNo, iPicNo2, 1, False, Pub_StrUserSt03
      Else
      'end 2020/3/31
         iPicNo = 10
         iPicNo2 = 11
      End If 'Added by Morgan 2020/3/31
      
      bolRetry = False
      
      'pub_OsPrinter = PUB_GetOsDefaultPrinter
      'PUB_SetOsDefaultPrinter cmbPrinter
      'PUB_SetWordActivePrinter
      
   
On Error GoTo ERRORSECTION1

      If TypeName(g_WordAp) <> "Application" Then
         Set g_WordAp = New Word.Application
      ''如果原先的Word的正在使用中則另外開新的跑
      ElseIf g_WordAp.Visible = True Then
         bVisible = g_WordAp.Visible
      End If
   
      'g_WordAp.Visible = False
      With g_WordAp.Application
      .Documents.add
      If PUB_ReadDB2File(stFileName, iPicNo) = True Then
         '切換為整頁模式
         If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
            .ActiveWindow.ActivePane.View.Type = wdPageView
         Else
            .ActiveWindow.View.Type = wdPageView
         End If
         
         .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
         Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
         oShape.ZOrder 4
         oShape.LockAnchor = True
         oShape.LockAspectRatio = -1
         oShape.Width = .CentimetersToPoints(21)
         oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
         oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
         oShape.Left = .CentimetersToPoints(0)
         oShape.Top = .CentimetersToPoints(0.5)
         If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
            .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
            Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
            oShape.ZOrder 4
            oShape.LockAnchor = True
            oShape.LockAspectRatio = -1
            oShape.Width = .CentimetersToPoints(21)
            oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
            oShape.Left = .CentimetersToPoints(0)
            'Added by Morgan 2020/3/31
            If strSrvDate(1) >= 智慧所更名日 Then
               oShape.Top = .CentimetersToPoints(27.2)
            Else
            'end 2020/3/31
               oShape.Top = .CentimetersToPoints(27)
            End If 'Added by Morgan 2020/3/31
         End If
         .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
         .Selection.EndKey Unit:=wdStory
      End If
      
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4.2)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(3)
      .Selection.Orientation = wdTextOrientationHorizontal
      
      
      iSNo = 0
      rsTemp1.MoveFirst
      Do While Not rsTemp1.EOF
      
         .Selection.Font.Name = "標楷體"
         .Selection.Font.Size = 14
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
         .Selection.ParagraphFormat.DisableLineHeightGrid = True
         .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
         .Selection.ParagraphFormat.LineSpacing = 15
            
         If strCustomer <> "" Then
            '跳頁
            .Selection.EndKey Unit:=wdStory
            .Selection.InsertBreak Type:=wdPageBreak
         End If
               
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         iLineCount = 0
            
         strCustomer = rsTemp1("CuNo")
         Call PUB_GetAddrRef(strCustomer, , , , , stReceiver, stContact, stZip, stAddr)
         stApp1Title = PUB_GetAppTitle(strCustomer)
            
         '郵遞區號
         .Selection.TypeText stZip
         '掛號
         '.Selection.TypeText String(14 - Len(stZip), "　")
         '.Selection.Font.Borders(1).LineStyle = .Options.DefaultBorderLineStyle
         '.Selection.TypeText "掛號"
         '.Selection.Font.Borders(1).LineStyle = wdLineStyleNone
         .Selection.TypeParagraph
         iLineCount = iLineCount + 1
            
         '地址
         If stAddr <> "" Then
            strExc(0) = stAddr
            Do While Len(strExc(0)) > 17
               .Selection.TypeText Left(strExc(0), 17)
               .Selection.TypeParagraph
               strExc(0) = Mid(strExc(0), 18)
               iLineCount = iLineCount + 1
            Loop
            If strExc(0) <> "" Then
               .Selection.TypeText strExc(0)
               .Selection.TypeParagraph
               iLineCount = iLineCount + 1
            End If
          End If
          
         '收件人
         If stReceiver <> "" Then
            strExc(0) = stReceiver
            Do While Len(strExc(0)) > 17
               .Selection.TypeText Left(strExc(0), 17)
               .Selection.TypeParagraph
               strExc(0) = Mid(strExc(0), 18)
               iLineCount = iLineCount + 1
            Loop
            .Selection.TypeText strExc(0)
            .Selection.TypeParagraph
            iLineCount = iLineCount + 1
         End If
             
         If GetTextLength(stContact) < 12 Then
            .Selection.TypeText stContact & String(12 - GetTextLength(stContact), " ") & "鈞啟"
         Else
            .Selection.TypeText stContact & " 鈞啟"
         End If
         .Selection.TypeParagraph
         iLineCount = iLineCount + 1
             
          '補滿5行
          For intI = iLineCount + 1 To 5
             .Selection.TypeParagraph
          Next
         
          .Selection.TypeParagraph
          
          .Selection.TypeText "致：" & stReceiver & stApp1Title
          .Selection.TypeParagraph
          
         .Selection.TypeParagraph
            
         '帶出系統日期
         '.Selection.TypeText "　　　　　　　　　　　　　　　　　　　　中華民國 " & Val(Mid(strSrvDate(1), 1, 4)) - 1911 & "年" & Mid(strSrvDate(1), 5, 2) & "月    日"
         '.Selection.TypeParagraph
         'Add by Amy 2021/03/08 Replace<信函下款>並設字大小
         .Selection.Font.Size = 12
         strContent2 = Replace(strContent, "<信函下款>", CompNameQuery("2"))
         strContent2 = Replace(strContent2, "<敬稱>", "" & rsTemp1("C02"))
         'Modify by Amy 2021/03/08 原:wdLineSpace1pt5 1.5倍行高
         .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceAtLeast
         .Selection.ParagraphFormat.LineSpacing = 0.7 '最小行高
         'end 2021/03/08
         .Selection.TypeText strContent2
         .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
         
         'Modify by Amy 2021/03/08 原100 換一個檔,但因文字太多造成「溢位」之錯誤
         If rsTemp1.AbsolutePosition Mod 30 = 0 Then
            .Selection.WholeStory
            ChgWordFormat g_WordAp, .Selection.Text
            '.PrintOut Background:=False, Copies:=1, Collate:=True
            iSNo = iSNo + 1
            .ActiveDocument.SaveAs strPath & "\離職通知函" & Format(iSNo, "000") & ".doc"
            
            .Selection.WholeStory
            .Selection.Delete
            strCustomer = ""
         End If
         rsTemp1.MoveNext
      Loop
      .Selection.WholeStory
      ChgWordFormat g_WordAp, .Selection.Text
      
      '.PrintOut Background:=False, Copies:=1, Collate:=True
      iSNo = iSNo + 1
      .ActiveDocument.SaveAs strPath & "\離職通知函" & Format(iSNo, "000") & ".doc"
      .ActiveDocument.Close wdDoNotSaveChanges
      
      End With
      
      'PUB_SetOsDefaultPrinter pub_OsPrinter
      
      MsgBox "列印結束 !", vbInformation
   Else
      MsgBox "無符合條件之資料可列印 !", vbInformation
   End If
   
ERRORSECTION1:
   If Err.Number <> 0 Then
      Select Case Err.Number
         Case 91, 462:
            Set g_WordAp = New Word.Application
            g_WordAp.Documents.add
            If bolRetry = False Then
               bolRetry = True
               Resume
            End If
         Case Else:
            MsgBox "錯誤 : " & Err.Description, vbCritical
      End Select
   End If
   
   Set rsTemp1 = Nothing
   Set g_WordAp = Nothing
End Sub

'Added by Morgan 2024/3/6
'離職通知函
'1.將智權部提供的定稿新增到P-22-000-XX
'2.將智權部提供的Excel檔匯入暫存表Morgan(M00,M01,M02): M00=流水號(定稿別*1000+列號),M01=客戶編號,M02=定稿處理狀況代碼
'3.逐一執行後產生Word檔
'4.手動轉存Pdf並依智權人員合併
Private Sub PrintLetterBatch3()
   Dim rsTemp1 As ADODB.Recordset
   Dim strContent As String, strContent2 As String
   Dim bVisible As Boolean
   Dim iPicNo As Integer, stFileName As String
   Dim iPicNo2 As Integer
   Dim oShape
   Dim bolRetry As Boolean
   
   Dim strCustomer As String
   Dim stReceiver As String '收件人
   Dim stAddr As String '地址
   Dim stZip As String '郵遞區號
   Dim stContact As String '接洽人
   Dim stApp1Title As String '稱謂
   Dim iLineCount As Integer '行數
   Dim strPath As String
   Dim iSNo As Integer
   Dim stFTM04 As String, stSales As String, stSaveName As String
   
   strPath = PUB_Getdesktop
   'Add by Amy 2021/03/08 建資料夾
   If Dir(PUB_Getdesktop & "\tmp", vbDirectory) = MsgText(601) Then
        MkDir strPath & "\tmp"
   End If
   strPath = strPath & "\tmp"
   
   
'   stFTM04 = "13": stSales = "劉峻綱"
'   stFTM04 = "14": stSales = "宋子明"
'   stFTM04 = "15": stSales = "張肆明"
'   stFTM04 = "16": stSales = "北五備用"
'   stFTM04 = "17": stSales = "北四備用"
'   stFTM04 = "18": stSales = "北三備用"
'   stFTM04 = "19": stSales = "宋子明" 'B2031
   stFTM04 = "20": stSales = "劉峻綱" 'B2029
   
    If MsgBox("將開始產生【" & stSales & "】的客戶函，是否確認要繼續？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Exit Sub
    End If
   
   ProgressBar1.Value = 0
   ProgressBar1.Visible = True
   lblProgress.Visible = True
   
   'Modify By Sindy 2012/5/24 DECODE(CU15,'0','台端','貴公司')==>DECODE(CU15,'0','台端','1','貴公司','貴單位')
   strExc(0) = "SELECT NVL(NVL(CU104,CU04),CU05||CU88||CU89) C00, NVL(NVL(NVL(PCC05,CU104),CU04),CU05||CU88||CU89) C01, DECODE(CU15,'0','台端','1','貴公司','貴單位') C02, ST02 C03,CU01||CU02 CuNo" & _
      " FROM MORGAN,CUSTOMER,POTCUSTCONT,STAFF WHERE M02='" & stFTM04 & "' AND CU01(+)=SUBSTR(M01,1,8) AND CU02='0' AND ST01(+)=CU13 AND PCC01(+)=CU01 AND PCC02(+)=CU127 ORDER BY M00"
   intI = 1
   Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ProgressBar1.max = rsTemp1.RecordCount
      lblProgress = ProgressBar1.Value & "/" & ProgressBar1.max
      
      strExc(0) = "SELECT FTM05,FTM08 FROM FINALTEXTMAP WHERE FTM01='P' AND FTM02='22' AND FTM03='000' AND FTM04='" & stFTM04 & "'"
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strContent = RsTemp(0) & RsTemp(1)
      Else
         MsgBox "無法取得定稿內容!!", vbCritical
         Exit Sub
      End If
      
      
      PUB_GetLetterPicID "2", "P", iPicNo, iPicNo2, 1, False, Pub_StrUserSt03
      
      bolRetry = False
      
      'pub_OsPrinter = PUB_GetOsDefaultPrinter
      'PUB_SetOsDefaultPrinter cmbPrinter
      'PUB_SetWordActivePrinter
      
   
On Error GoTo ERRORSECTION1

      If TypeName(g_WordAp) <> "Application" Then
         Set g_WordAp = New Word.Application
      ''如果原先的Word的正在使用中則另外開新的跑
      ElseIf g_WordAp.Visible = True Then
         bVisible = g_WordAp.Visible
      End If
      
      With g_WordAp.Application
      '.Visible = True
      .Visible = False
      .Documents.add
      If PUB_ReadDB2File(stFileName, iPicNo) = True Then
         '切換為整頁模式
         If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
            .ActiveWindow.ActivePane.View.Type = wdPageView
         Else
            .ActiveWindow.View.Type = wdPageView
         End If
         
         .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
         Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
         oShape.ZOrder 4
         oShape.LockAnchor = True
         oShape.LockAspectRatio = -1
         oShape.Width = .CentimetersToPoints(21)
         oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
         oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
         oShape.Left = .CentimetersToPoints(0)
         oShape.Top = .CentimetersToPoints(0.5)
         If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
            .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
            Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
            oShape.ZOrder 4
            oShape.LockAnchor = True
            oShape.LockAspectRatio = -1
            oShape.Width = .CentimetersToPoints(21)
            oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
            oShape.Left = .CentimetersToPoints(0)
            'Added by Morgan 2020/3/31
            If strSrvDate(1) >= 智慧所更名日 Then
               oShape.Top = .CentimetersToPoints(27.2)
            Else
            'end 2020/3/31
               oShape.Top = .CentimetersToPoints(27)
            End If 'Added by Morgan 2020/3/31
         End If
         .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
         .Selection.EndKey Unit:=wdStory
      End If
      
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4.2)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(3)
      .Selection.Orientation = wdTextOrientationHorizontal
      
      
      iSNo = 0
      rsTemp1.MoveFirst
      .Visible = True
      Do While Not rsTemp1.EOF
      
         .Selection.Font.Name = "標楷體"
         .Selection.Font.Size = 14
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
         .Selection.ParagraphFormat.DisableLineHeightGrid = True
         .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
         .Selection.ParagraphFormat.LineSpacing = 15
            
         If strCustomer <> "" Then
            '跳頁
            .Selection.EndKey Unit:=wdStory
            .Selection.InsertBreak Type:=wdPageBreak
         End If
               
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　中華民國" & Mid(strSrvDate(2), 1, 3) & "(" & Mid(strSrvDate(1), 1, 4) & ")年" & Mid(strSrvDate(2), 4, 2) & "月" & Right(strSrvDate(2), 2) & "日"
         iLineCount = 0
            
         strCustomer = rsTemp1("CuNo")
         Call PUB_GetAddrRef(strCustomer, , , , , stReceiver, stContact, stZip, stAddr)
         stApp1Title = PUB_GetAppTitle(strCustomer)
            
         '郵遞區號
         .Selection.TypeText stZip
         '掛號
         '.Selection.TypeText String(14 - Len(stZip), "　")
         '.Selection.Font.Borders(1).LineStyle = .Options.DefaultBorderLineStyle
         '.Selection.TypeText "掛號"
         '.Selection.Font.Borders(1).LineStyle = wdLineStyleNone
         .Selection.TypeParagraph
         iLineCount = iLineCount + 1
            
         '地址
         If stAddr <> "" Then
            strExc(0) = stAddr
            Do While Len(strExc(0)) > 17
               .Selection.TypeText Left(strExc(0), 17)
               .Selection.TypeParagraph
               strExc(0) = Mid(strExc(0), 18)
               iLineCount = iLineCount + 1
            Loop
            If strExc(0) <> "" Then
               .Selection.TypeText strExc(0)
               .Selection.TypeParagraph
               iLineCount = iLineCount + 1
            End If
          End If
          
         '收件人
         If stReceiver <> "" Then
            strExc(0) = stReceiver
            Do While Len(strExc(0)) > 17
               .Selection.TypeText Left(strExc(0), 17)
               .Selection.TypeParagraph
               strExc(0) = Mid(strExc(0), 18)
               iLineCount = iLineCount + 1
            Loop
            .Selection.TypeText strExc(0)
            .Selection.TypeParagraph
            iLineCount = iLineCount + 1
         End If
             
         If GetTextLength(stContact) < 12 Then
            .Selection.TypeText stContact & String(12 - GetTextLength(stContact), " ") & "鈞啟"
         Else
            .Selection.TypeText stContact & " 鈞啟"
         End If
         .Selection.TypeParagraph
         iLineCount = iLineCount + 1
             
          '補滿5行
          For intI = iLineCount + 1 To 5
             .Selection.TypeParagraph
          Next
         
          .Selection.TypeParagraph
          
          .Selection.TypeText "致：" & stReceiver & stApp1Title
          .Selection.TypeParagraph
          
         .Selection.TypeParagraph
         
         strContent2 = Replace(strContent, "<敬稱>", "" & rsTemp1("C02"))
         .Selection.TypeText strContent2
         .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
         
         'Modify by Amy 2021/03/08 原100 換一個檔,但因文字太多造成「溢位」之錯誤
         If rsTemp1.AbsolutePosition Mod 30 = 0 Then
            .Selection.WholeStory
            ChgWordFormat g_WordAp, .Selection.Text
            '.PrintOut Background:=False, Copies:=1, Collate:=True
            iSNo = iSNo + 1
            stSaveName = strPath & "\敬致客戶函-" & stSales & Format(iSNo, "000")
            .ActiveDocument.SaveAs stSaveName & ".doc"
            .ActiveDocument.ExportAsFixedFormat OutputFileName:=stSaveName & ".pdf", ExportFormat:=17, OpenAfterExport:=False
            .Selection.WholeStory
            .Selection.Delete
            strCustomer = ""
         End If
         ProgressBar1.Value = ProgressBar1.Value + 1
         lblProgress = ProgressBar1.Value & "/" & ProgressBar1.max
         rsTemp1.MoveNext
      Loop
      .Selection.WholeStory
      ChgWordFormat g_WordAp, .Selection.Text
      
      '.PrintOut Background:=False, Copies:=1, Collate:=True
      iSNo = iSNo + 1
      
      stSaveName = strPath & "\敬致客戶函-" & stSales & Format(iSNo, "000")
      .ActiveDocument.SaveAs stSaveName & ".doc"
      .ActiveDocument.ExportAsFixedFormat OutputFileName:=stSaveName & ".pdf", ExportFormat:=17, OpenAfterExport:=False
      .ActiveDocument.Close wdDoNotSaveChanges
      End With
      g_WordAp.Quit wdDoNotSaveChanges
      
      'PUB_SetOsDefaultPrinter pub_OsPrinter
      
      MsgBox "列印結束 !", vbInformation
   Else
      MsgBox "無符合條件之資料可列印 !", vbInformation
   End If
   
   ProgressBar1.Visible = False
   lblProgress.Visible = False
   
ERRORSECTION1:
   If Err.Number <> 0 Then
      Select Case Err.Number
         Case 91, 462:
            Set g_WordAp = New Word.Application
            g_WordAp.Documents.add
            If bolRetry = False Then
               bolRetry = True
               Resume
            End If
         Case Else:
            MsgBox "錯誤 : " & Err.Description, vbCritical
      End Select
   End If
   
   Set rsTemp1 = Nothing
   Set g_WordAp = Nothing
End Sub
