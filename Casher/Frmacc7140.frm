VERSION 5.00
Begin VB.Form Frmacc7140 
   BorderStyle     =   1  '單線固定
   Caption         =   "分所銀行入帳媒體作業"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5280
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   1
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "1"
      Top             =   1200
      Width           =   300
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   4110
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "轉檔(&O)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3090
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   0
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   0
      Top             =   780
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "請輸入媒體遞送單上之預定撥帳日期"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   2160
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Label Label4 
      Caption         =   "(1: 每月薪資 2 : 年終獎金   3: 端午代金  4: 中秋代金)"
      Height          =   360
      Left            =   1800
      TabIndex        =   7
      Top             =   1260
      Width           =   2000
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PS：媒體檔案放在桌面, 檔名為 Salary.txt"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   270
      TabIndex        =   6
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "資料類別："
      Height          =   180
      Left            =   270
      TabIndex        =   5
      Top             =   1260
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "入帳日期："
      Height          =   180
      Index           =   0
      Left            =   270
      TabIndex        =   4
      Top             =   840
      Width           =   900
   End
End
Attribute VB_Name = "Frmacc7140"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by sonia 2013/3/6
Option Explicit
Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 10) As Integer
Dim strTemp(1 To 20) As String
Dim iPgae As Integer, iLine As Integer
Dim strFileName As String
Dim strSD05 As String
Dim dblTotAmt As Double
Dim dblTotCnt As Double

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
         
         Screen.MousePointer = vbHourglass
         
         Select Case txt1(1)
            Case "1"       '每月薪資
               StrMenu1
            Case "2"       '年終獎金
               StrMenu2
            Case "3", "4"  '端午,中秋代金
               StrMenu3
         End Select
         Screen.MousePointer = vbDefault
      Case 1
         Unload Me
   End Select
End Sub

Sub StrMenu1()
Dim ff As Integer
Dim strText As String
Dim strYM As String
Dim TempFileName As String
Dim dblTotSaveCode As Double

   strYM = Left(ChangeTStringToWString(txt1(0)), 6)
   'Modified by Morgan 2022/1/28
   '若日期小於15號時才抓上月薪資 Ex:1110128提早發月薪
   If Right(txt1(0), 2) < "15" Then
      If Right(strYM, 2) = "01" Then
         strYM = Left(strYM, 4) - 1 & "12"
      Else
         strYM = strYM - 1
      End If
   End If
   'end 2022/1/28
   
   Select Case pub_strUserOffice
'2013/7/29 CANCEL BY SONIA 中所由台北直接產生寄中所交中國信託
'      '2013/7/29 add by sonia
'      Case "2"   '中所
'         m_str = "SELECT SD01,ST02,SD06,ST26,nvl(SM04,0)+nvl(SM05,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0) as T2 " & _
'                 "FROM STAFF,SALARYDATA,SalaryMonth WHERE SM02=" & strYM & " AND SM01=ST01(+) AND SM01=SD01(+) AND SD05='" & strSD05 & "' Order by SD01"
'         If m_rs.State = 1 Then m_rs.Close
'         m_rs.CursorLocation = adUseClient
'         m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
'         If Not m_rs.EOF And Not m_rs.BOF Then
'            With m_rs
'               m_rs.MoveFirst
'               dblTotAmt = 0
'               dblTotSaveCode = 0
'
'               If ff > 0 Then Close #ff
'               ff = FreeFile
'               TempFileName = PUB_Getdesktop & "\" & strFileName
'               Open TempFileName For Output As ff
'
'               '首筆(第一筆)
'               strTemp(1) = "159540272896"                                                     '公司帳號      X(12)
'               strTemp(2) = Right("0000000000" & TAIWANDATE(txt1(0)), 10)                      '入帳日期      9(10) 民國年月日YYYMMDD(不足10碼前補0)
'               strTemp(3) = "A"                                                                '收付業務別    X(01) 固定為 'A'
'
'               strText = ""
'               For m_i = 1 To 3
'                  strText = strText & strTemp(m_i)
'               Next m_i
'               Print #ff, strText
'
'               '明細(第二筆以後)
'               Do While Not m_rs.EOF
'
'                  For m_i = 1 To 6
'                     strTemp(m_i) = ""
'                  Next m_i
'
'                  strTemp(1) = Left(CheckStr(m_rs.Fields("SD06")) & "            ", 12)       '入帳帳號       X(12)
'                  strTemp(2) = Right("0000000000" & CheckStr(m_rs.Fields("T2")), 10)          '入帳金額       9(10)
'                  strTemp(3) = "          "                                                   '收款人姓名     X(10) 中國信託建議放空白
'                  strTemp(4) = Left(CheckStr(m_rs.Fields("ST26")) & "          ", 10)         '身份證字號     X(10)
'                  strTemp(5) = " "                                                            '入帳種類代碼   X(01) 空白：薪資
'                  strTemp(6) = "Y"                                                            '檢查功能碼     X(03) 檢查入帳戶之身份證字號
'
'                  strText = ""
'                  For m_i = 1 To 6
'                     strText = strText & strTemp(m_i)
'                  Next m_i
'                  Print #ff, strText
'                  m_rs.MoveNext
'               Loop
'
'            End With
'            Close ff
'
'            MsgBox "每月薪資媒體已完成! 收到北所寄下來之媒體遞送單後即可送交銀行!! 辛苦了 !!", vbExclamation + vbOKOnly
'            Exit Sub
'         Else
'            MsgBox "無符合條件的資料!!!", vbExclamation + vbOKOnly
'            Exit Sub
'         End If
'      '2013/7/29 end
'
      Case "3"   '南所
         'Modify By Sindy 2020/6/25 + 證照津貼
         m_str = "SELECT SD01,ST02,SD06,nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0) as T2 " & _
                 "FROM STAFF,SALARYDATA,SalaryMonth WHERE SM02=" & strYM & " AND SM01=ST01(+) AND SM01=SD01(+) AND SD05='" & strSD05 & "' Order by SD01"
         If m_rs.State = 1 Then m_rs.Close
         m_rs.CursorLocation = adUseClient
         m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
         If Not m_rs.EOF And Not m_rs.BOF Then
            With m_rs
               m_rs.MoveFirst
               dblTotAmt = 0
               dblTotSaveCode = 0
               
               If ff > 0 Then Close #ff
               ff = FreeFile
               TempFileName = PUB_Getdesktop & "\" & strFileName
               Open TempFileName For Output As ff
                
               Do While Not m_rs.EOF
                   
                  For m_i = 1 To 9
                     strTemp(m_i) = ""
                  Next m_i
                  
                  strTemp(1) = Right("00000000" & txt1(0), 8)                                 'Date       交易生效日 9(08) 民國年月日YYYYMMDD
                  'modify by sonia 2020/4/9
                  'strTemp(2) = "AALZ"                                                         'Code       機關代號   X(04)
                  strTemp(2) = "TAIE"                                                         'Code       機關代號   X(04)
                  'end 2020/4/9
                  strTemp(3) = Left(CheckStr(m_rs.Fields("SD06")) & "              ", 14)     'ID         帳號       X(14)
                  strTemp(4) = "C"                                                            'PayCode    存入或支出 X(01) 固定為 'C'
                  strTemp(5) = Right("00000000000" & CheckStr(m_rs.Fields("T2")), 11) & "00"  'PayAmtData 金額       9(11)V9(2)
                  strTemp(6) = Left(CheckStr(m_rs.Fields("ST02")) & "            ", 12)       'EmpID      員工編號   X(15) 南所放員工姓名,故碼數改12
                  strTemp(7) = "900"                                                          'CostCode   費用代碼   X(03) 固定為 '900'
                  strTemp(9) = String(18, " ")                                                'Filler     空白       X(18)
                  
                  '安全碼=帳號欄第1碼*金額欄第1碼+帳號欄第2碼*金額欄第2碼+帳號欄第3碼*金額欄第3碼+......帳號欄第13碼*金額欄第13碼+帳號欄第14碼
                  For m_i = 1 To 13
                     strTemp(8) = Val(strTemp(8)) + Val(Mid(strTemp(3), m_i, 1)) * Val(Mid(strTemp(5), m_i, 1))
                  Next m_i
                  strTemp(8) = Val(strTemp(8)) + Val(Mid(strTemp(3), 14, 1))
                  strTemp(8) = Right("0000" & CheckStr(strTemp(8)), 4)                        'SaftyCode  安全碼     X(04)
                  
                  '合計欄之安全碼
                  dblTotSaveCode = dblTotSaveCode + Val(strTemp(8))
                  '總金額
                  dblTotAmt = dblTotAmt + CheckStr(m_rs.Fields("T2"))
                  
                  strText = ""
                  For m_i = 1 To 9
                     strText = strText & strTemp(m_i)
                  Next m_i
                  Print #ff, strText
                  m_rs.MoveNext
               Loop
                
               '總計：最後一筆
               For m_i = 1 To 9
                   strTemp(m_i) = ""
               Next m_i
               strTemp(1) = Right("00000000" & txt1(0), 8)                                    'Date       交易生效日 9(08) 民國年月日YYYYMMDD
               'modify by sonia 2020/4/9
               'strTemp(2) = "AALZ"                                                            'Code       機關代號   X(04)
               strTemp(2) = "TAIE"                                                            'Code       機關代號   X(04)
               'end 2020/4/9
               strTemp(3) = "99999999999999"                                                  'ID         帳號       X(14)
               strTemp(4) = "C"                                                               'PayCode    存入或支出 X(01) 固定為 'C'
               strTemp(5) = Right("00000000000" & CheckStr(dblTotAmt), 11) & "00"             'PayAmtData 金額       9(11)V9(2)
               strTemp(6) = "999999999999999"                                                 'EmpID      員工編號   X(15)
               strTemp(7) = "900"                                                             'CostCode   費用代碼   X(03) 固定為 '900'
               strTemp(9) = String(18, " ")                                                   'Filler     空白       X(18)
               
               '合計欄之安全碼=SUM(明細安全碼)/11
               dblTotSaveCode = dblTotSaveCode / 11
               strTemp(8) = Right("0000" & CheckStr(Int(dblTotSaveCode)), 4)                  'SaftyCode  安全碼     X(04)
               
               strText = ""
               For m_i = 1 To 9
                  strText = strText & strTemp(m_i)
               Next m_i
               Print #ff, strText
                 
            End With
            Close ff
            
            MsgBox "每月薪資媒體已完成! 收到北所寄下來之明細表後即可送交銀行!! 辛苦了 !!", vbExclamation + vbOKOnly
            Exit Sub
         Else
            MsgBox "無符合條件的資料!!!", vbExclamation + vbOKOnly
            Exit Sub
         End If
      
'modify by sonia 2017/3/28 高所由彰銀改合庫
'      Case "4"   '高所
'         m_str = "SELECT SD01,ST02,SD06,ST26,nvl(SM04,0)+nvl(SM05,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0) as T2 " & _
'                 "FROM STAFF,SALARYDATA,SalaryMonth WHERE SM02=" & strYM & " AND SM01=ST01(+) AND SM01=SD01(+) AND SD05='" & strSD05 & "' Order by T2 DESC,SD01"
'         If m_rs.State = 1 Then m_rs.Close
'         m_rs.CursorLocation = adUseClient
'         m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
'         If Not m_rs.EOF And Not m_rs.BOF Then
'            With m_rs
'               m_rs.MoveFirst
'               dblTotAmt = 0
'               dblTotCnt = 0
'
'               If ff > 0 Then Close #ff
'               ff = FreeFile
'               TempFileName = PUB_Getdesktop & "\" & strFileName
'               Open TempFileName For Output As ff
'
'               '首筆(第一筆)
'               For m_i = 1 To 12
'                  strTemp(m_i) = ""
'               Next m_i
'
'               strTemp(1) = "1"                                                            '區別碼        9(01) 固定為 1
'               strTemp(2) = "057"                                                          '企業編號1     9(03)
'               strTemp(3) = "F"                                                            '企業編號2     X(01) 固定為 'F'
'               strTemp(4) = "9634"                                                         '分行代號      9(04)
'               strTemp(5) = DBDATE(txt1(0))                                                '日期          9(08) 西元年月日YYYYMMDD
'               strTemp(6) = "2"                                                            '存提代號      9(01) 2為存款
'               strTemp(7) = "051"                                                          '摘要          9(03)
'               strTemp(8) = "CUST "                                                        '磁片來源      X(05) 自行轉出固定為 'CUST '
'               strTemp(9) = "1"                                                            '性質別        9(01) 固定為 1 (IN)
'               strTemp(10) = "04150022  "                                                  '公司統一編號  X(10)
'               strTemp(11) = "96340100346000"                                              '公司帳號      9(14)
'               strTemp(12) = String(79, " ")                                               '空白          X(79)
'
'               strText = ""
'               For m_i = 1 To 12
'                  strText = strText & strTemp(m_i)
'               Next m_i
'               Print #ff, strText
'
'               '明細(第二筆至倒數第二筆)
'               Do While Not m_rs.EOF
'
'                  For m_i = 1 To 18
'                     strTemp(m_i) = ""
'                  Next m_i
'
'                  strTemp(1) = "2"                                                               '區別碼        9(01) 固定為 2
'                  strTemp(2) = "057"                                                             '企業編號1     9(03)
'                  strTemp(3) = "F"                                                               '企業編號2     X(01) 固定為 'F'
'                  strTemp(4) = "9634"                                                            '分行代號      9(04)
'                  strTemp(5) = DBDATE(txt1(0))                                                   '日期          9(08) 西元年月日YYYYMMDD
'                  strTemp(6) = "2"                                                               '存提代號      9(01) 2為存款
'                  strTemp(7) = "051"                                                             '摘要          9(03)
'                  strTemp(8) = String(5, " ")                                                    '空白          X(05)
'                  strTemp(9) = Left(CheckStr(m_rs.Fields("SD06")) & "              ", 14)        '銀行帳號      9(14)
'                  strTemp(10) = Right("000000000000" & CheckStr(m_rs.Fields("T2")), 12) & "00"   '金額          9(12)V9(2)
'                  strTemp(11) = "99"                                                             '狀況代號      9(02) 轉出固定為 99
'                  strTemp(12) = String(10, " ")                                                  '交易註記1     X(10)
'                  strTemp(13) = Left(CheckStr(m_rs.Fields("SD01")) & "          ", 10)           '交易註記2     X(10) 高所放員工編號
'                  strTemp(14) = Left(CheckStr(m_rs.Fields("ST26")) & "          ", 10)           '身份證字號    X(10)
'                  strTemp(15) = String(20, " ")                                                  '專用資料區    X(10)
'                  strTemp(16) = "Y"                                                              '身份證檢核註記X(01) 要檢查為 'Y'
'                  strTemp(17) = "  "                                                             '幣別          9(02) 台幣為空白
'                  strTemp(18) = String(21, " ")                                                  '空白欄        X(21)
'
'                  '合計欄之安全碼
'                  dblTotCnt = dblTotCnt + 1
'                  '總金額
'                  dblTotAmt = dblTotAmt + CheckStr(m_rs.Fields("T2"))
'
'                  strText = ""
'                  For m_i = 1 To 18
'                     strText = strText & strTemp(m_i)
'                  Next m_i
'                  Print #ff, strText
'                  m_rs.MoveNext
'               Loop
'
'               '總計：尾筆(最後一筆)
'               For m_i = 1 To 13
'                   strTemp(m_i) = ""
'               Next m_i
'
'               strTemp(1) = "3"                                                            '區別碼        9(01) 固定為 3
'               strTemp(2) = "057"                                                          '企業編號1     9(03)
'               strTemp(3) = "F"                                                            '企業編號2     X(01) 固定為 'F'
'               strTemp(4) = "9634"                                                         '分行代號      9(04)
'               strTemp(5) = DBDATE(txt1(0))                                                '日期          9(08) 西元年月日YYYYMMDD
'               strTemp(6) = "2"                                                            '存提代號      9(01) 2為存款
'               strTemp(7) = "051"                                                          '摘要          9(03)
'               strTemp(8) = String(5, " ")                                                 '空白欄        X(05)
'               strTemp(9) = Right("00000000000000" & CheckStr(dblTotAmt), 14) & "00"       '總金額        9(14)V9(2)
'               strTemp(10) = Right("0000000000" & CheckStr(dblTotCnt), 10)                 '總筆數        9(10)
'               strTemp(11) = String(16, " ")                                               '未成交總金額  9(16) 空白
'               strTemp(12) = String(10, " ")                                               '未成交總筆數  9(10) 空白
'               strTemp(13) = String(52, " ")                                               '空白欄        X(52)
'
'               strText = ""
'               For m_i = 1 To 13
'                  strText = strText & strTemp(m_i)
'               Next m_i
'               Print #ff, strText
'
'            End With
'            Close ff
'
'            MsgBox "每月薪資媒體已完成! 收到北所寄下來之媒體遞送單後即可送交銀行!! 辛苦了 !!", vbExclamation + vbOKOnly
'
'            Exit Sub
'         Else
'            MsgBox "無符合條件的資料!!!", vbExclamation + vbOKOnly
'            Exit Sub
'         End If
      Case "4"   '高所
         'Modify By Sindy 2020/6/25 + 證照津貼
         m_str = "SELECT SD01,ST02,SD06,ST26,nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)+nvl(SM06,0)+nvl(SM07,0)+nvl(SM08,0)+nvl(SM09,0)+nvl(SM10,0)+nvl(SM12,0)+nvl(SM13,0)-nvl(SM14,0)-nvl(SM15,0)-nvl(SM16,0)-nvl(SM17,0)-nvl(SM18,0)-nvl(SM19,0)-nvl(SM20,0)-nvl(SM21,0)-nvl(SM22,0)-nvl(SM23,0)-nvl(SM24,0)-nvl(SM43,0) as T2 " & _
                 "FROM STAFF,SALARYDATA,SalaryMonth WHERE SM02=" & strYM & " AND SM01=ST01(+) AND SM01=SD01(+) AND SD05='" & strSD05 & "' Order by SD06"
         If m_rs.State = 1 Then m_rs.Close
         m_rs.CursorLocation = adUseClient
         m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
         If Not m_rs.EOF And Not m_rs.BOF Then
            
            With m_rs
               m_rs.MoveFirst
               dblTotAmt = 0
               dblTotCnt = 0
               
               If ff > 0 Then Close #ff
               ff = FreeFile
               TempFileName = PUB_Getdesktop & "\" & strFileName
               Open TempFileName For Output As ff
                
               '明細
               Do While Not m_rs.EOF
                   
                  For m_i = 1 To 10
                     strTemp(m_i) = ""
                  Next m_i
                  
                  strTemp(1) = "0"                                                              '轉帳類別      X(01) 固定為 0 薪資
                  strTemp(2) = "2"                                                              '檢查身分證號  X(01) 固定為 2 有就檢查
                  strTemp(3) = Left(CheckStr(m_rs.Fields("SD06")) & "              ", 13)       '銀行帳號      9(13)
                  strTemp(4) = Right("00000000000" & CheckStr(m_rs.Fields("T2")), 11) & "00"    '轉帳金額      9(11)V9(2)
                  strTemp(5) = Left(CheckStr(m_rs.Fields("ST26")) & "          ", 10)           '身份證字號    X(10)
                  strTemp(6) = String(23, " ")                                                  '存戶使用區    X(23) 固定為空白
                  strTemp(7) = String(3, " ")                                                   '客戶編號      X(03) 固定為空白
                  strTemp(8) = Right("00000000" & CheckStr(txt1(0)), 8)                         '入帳日期      X(08) 民國年月日YYYYMMDD前補0
                  strTemp(9) = String(4, " ")                                                   'Z類摘要       X(04) 固定為空白
                  strTemp(10) = "5252"                                                          '委託單位      X(04)  固定為 5252 七賢分行
                  
                  strText = ""
                  For m_i = 1 To 10
                     strText = strText & strTemp(m_i)
                  Next m_i
                  Print #ff, strText
                  
                  m_rs.MoveNext
               Loop
                
            End With
            Close ff
            
            MsgBox "每月薪資媒體已完成! 收到北所寄下來之媒體遞送資料表後即可送交合作金庫 !! 辛苦了 !!", vbExclamation + vbOKOnly
            
            Exit Sub
         Else
            MsgBox "無符合條件的資料!!!", vbExclamation + vbOKOnly
            Exit Sub
         End If
         'end 2017/3/28
   
   End Select
End Sub

Sub StrMenu2()
Dim ff As Integer
Dim strText As String
Dim strYM As String
Dim TempFileName As String
Dim dblTotSaveCode As Double
  
   strYM = Left(ChangeTStringToWString(txt1(0)), 4)
   strYM = strYM - 1
   
   Select Case pub_strUserOffice
'2013/7/29 CANCEL BY SONIA 中所由台北直接產生寄中所交中國信託
'      '2013/7/29 add by sonia
'      Case "2"   '中所
'         'modify by sonia 2016/1/6 留職停薪人員之年終資料改以現金發放故加入SD02<>'S'條件
'         m_str = "SELECT SD01,ST02,SD06,ST26,nvl(YB05,0)+nvl(YB06,0)+nvl(YB08,0)-nvl(YB15,0)-nvl(YB16,0)-nvl(YB17,0)-nvl(YB25,0) as T2 " & _
'                 "FROM STAFF,SALARYDATA,YearBonus WHERE YB01='" & strYM & "' AND YB02=ST01(+) AND YB02=SD01(+) AND SD05='" & strSD05 & "' AND SD02<>'S' Order by SD01"
'         If m_rs.State = 1 Then m_rs.Close
'         m_rs.CursorLocation = adUseClient
'         m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
'         If Not m_rs.EOF And Not m_rs.BOF Then
'            With m_rs
'               m_rs.MoveFirst
'               dblTotAmt = 0
'               dblTotSaveCode = 0
'
'               If ff > 0 Then Close #ff
'               ff = FreeFile
'               TempFileName = PUB_Getdesktop & "\" & strFileName
'               Open TempFileName For Output As ff
'
'               '首筆(第一筆)
'               strTemp(1) = "159540272896"                                                     '公司帳號      X(12)
'               strTemp(2) = Right("0000000000" & TAIWANDATE(txt1(0)), 10)                      '入帳日期      9(10) 民國年月日YYYMMDD(不足10碼前補0)
'               strTemp(3) = "A"                                                                '收付業務別    X(01) 固定為 'A'
'
'               strText = ""
'               For m_i = 1 To 3
'                  strText = strText & strTemp(m_i)
'               Next m_i
'               Print #ff, strText
'
'               '明細(第二筆以後)
'               Do While Not m_rs.EOF
'
'                  For m_i = 1 To 6
'                     strTemp(m_i) = ""
'                  Next m_i
'
'                  strTemp(1) = Left(CheckStr(m_rs.Fields("SD06")) & "            ", 12)       '入帳帳號       X(12)
'                  strTemp(2) = Right("0000000000" & CheckStr(m_rs.Fields("T2")), 10)          '入帳金額       9(10)
'                  strTemp(3) = "          "                                                   '收款人姓名     X(10) 中國信託建議放空白
'                  strTemp(4) = Left(CheckStr(m_rs.Fields("ST26")) & "          ", 10)         '身份證字號     X(10)
'                  strTemp(5) = " "                                                            '入帳種類代碼   X(01) 空白：薪資
'                  strTemp(6) = "Y"                                                            '檢查功能碼     X(03) 檢查入帳戶之身份證字號
'
'                  strText = ""
'                  For m_i = 1 To 6
'                     strText = strText & strTemp(m_i)
'                  Next m_i
'                  Print #ff, strText
'                  m_rs.MoveNext
'               Loop
'
'            End With
'            Close ff
'
'            MsgBox "年終獎金媒體已完成! 收到北所寄下來之媒體遞送後即可送交銀行!! 辛苦了 !!", vbExclamation + vbOKOnly
'            Exit Sub
'         Else
'            MsgBox "無符合條件的資料!!!", vbExclamation + vbOKOnly
'            Exit Sub
'         End If
'      '2013/7/29 end
      
      Case "3"   '南所
         'modify by sonia 2016/1/6 留職停薪人員之年終資料改以現金發放故加入SD02<>'S'條件
         'modify by sonia 2018/1/11 +YB26
         m_str = "SELECT SD01,ST02,SD06,nvl(YB05,0)+nvl(YB06,0)+nvl(YB26,0)+nvl(YB08,0)-nvl(YB15,0)-nvl(YB16,0)-nvl(YB17,0)-nvl(YB25,0) as T2 " & _
                 "FROM STAFF,SALARYDATA,YearBonus WHERE YB01='" & strYM & "' AND YB02=ST01(+) AND YB02=SD01(+) AND SD05='" & strSD05 & "' AND SD02<>'S' Order by SD01"
         If m_rs.State = 1 Then m_rs.Close
         m_rs.CursorLocation = adUseClient
         m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
         If Not m_rs.EOF And Not m_rs.BOF Then
            With m_rs
               m_rs.MoveFirst
               dblTotAmt = 0
               dblTotSaveCode = 0
               
               If ff > 0 Then Close #ff
               ff = FreeFile
               TempFileName = PUB_Getdesktop & "\" & strFileName
               Open TempFileName For Output As ff
                
               Do While Not m_rs.EOF
                   
                  For m_i = 1 To 9
                     strTemp(m_i) = ""
                  Next m_i
                  
                  strTemp(1) = Right("00000000" & txt1(0), 8)                                 'Date       交易生效日 9(08) 民國年月日YYYYMMDD
                  'modify by sonia 2020/4/9
                  'strTemp(2) = "AALZ"                                                         'Code       機關代號   X(04)
                  strTemp(2) = "TAIE"                                                         'Code       機關代號   X(04)
                  'end 2020/4/9
                  strTemp(3) = Left(CheckStr(m_rs.Fields("SD06")) & "              ", 14)     'ID         帳號       X(14)
                  strTemp(4) = "C"                                                            'PayCode    存入或支出 X(01) 固定為 'C'
                  strTemp(5) = Right("00000000000" & CheckStr(m_rs.Fields("T2")), 11) & "00"  'PayAmtData 金額       9(11)V9(2)
                  strTemp(6) = Left(CheckStr(m_rs.Fields("ST02")) & "            ", 12)       'EmpID      員工編號   X(15) 南所放員工姓名,故碼數改12
                  strTemp(7) = "900"                                                          'CostCode   費用代碼   X(03) 固定為 '900'
                  strTemp(9) = String(18, " ")                                                'Filler     空白       X(18)
                  
                  '安全碼=帳號欄第1碼*金額欄第1碼+帳號欄第2碼*金額欄第2碼+帳號欄第3碼*金額欄第3碼+......帳號欄第13碼*金額欄第13碼+帳號欄第14碼
                  For m_i = 1 To 13
                     strTemp(8) = Val(strTemp(8)) + Val(Mid(strTemp(3), m_i, 1)) * Val(Mid(strTemp(5), m_i, 1))
                  Next m_i
                  strTemp(8) = Val(strTemp(8)) + Val(Mid(strTemp(3), 14, 1))
                  strTemp(8) = Right("0000" & CheckStr(strTemp(8)), 4)                        'SaftyCode  安全碼     X(04)
                  
                  '合計欄之安全碼
                  dblTotSaveCode = dblTotSaveCode + Val(strTemp(8))
                  '總金額
                  dblTotAmt = dblTotAmt + CheckStr(m_rs.Fields("T2"))
                  
                  strText = ""
                  For m_i = 1 To 9
                     strText = strText & strTemp(m_i)
                  Next m_i
                  Print #ff, strText
                  m_rs.MoveNext
               Loop
                
               '總計：最後一筆
               For m_i = 1 To 9
                   strTemp(m_i) = ""
               Next m_i
               strTemp(1) = Right("00000000" & txt1(0), 8)                                    'Date       交易生效日 9(08) 民國年月日YYYYMMDD
               'modify by sonia 2020/4/9
               'strTemp(2) = "AALZ"                                                            'Code       機關代號   X(04)
               strTemp(2) = "TAIE"                                                            'Code       機關代號   X(04)
               'end 2020/4/9
               strTemp(3) = "99999999999999"                                                  'ID         帳號       X(14)
               strTemp(4) = "C"                                                               'PayCode    存入或支出 X(01) 固定為 'C'
               strTemp(5) = Right("00000000000" & CheckStr(dblTotAmt), 11) & "00"             'PayAmtData 金額       9(11)V9(2)
               strTemp(6) = "999999999999999"                                                 'EmpID      員工編號   X(15)
               strTemp(7) = "900"                                                             'CostCode   費用代碼   X(03) 固定為 '900'
               strTemp(9) = String(18, " ")                                                   'Filler     空白       X(18)
               
               '合計欄之安全碼=SUM(明細安全碼)/11
               dblTotSaveCode = dblTotSaveCode / 11
               strTemp(8) = Right("0000" & CheckStr(Int(dblTotSaveCode)), 4)                  'SaftyCode  安全碼     X(04)
               
               strText = ""
               For m_i = 1 To 9
                  strText = strText & strTemp(m_i)
               Next m_i
               Print #ff, strText
                 
            End With
            Close ff
            
            MsgBox "年終獎金媒體已完成! 收到北所寄下來之明細表後即可送交銀行!! 辛苦了 !!", vbExclamation + vbOKOnly
            Exit Sub
         Else
            MsgBox "無符合條件的資料!!!", vbExclamation + vbOKOnly
            Exit Sub
         End If
      
'modify by sonia 2017/3/28 高所由彰銀改合庫
'      Case "4"   '高所
'         'modify by sonia 2016/1/6 留職停薪人員之年終資料改以現金發放故加入SD02<>'S'條件
'         m_str = "SELECT SD01,ST02,SD06,ST26,nvl(YB05,0)+nvl(YB06,0)+nvl(YB08,0)-nvl(YB15,0)-nvl(YB16,0)-nvl(YB17,0)-nvl(YB25,0) as T2 " & _
'                 "FROM STAFF,SALARYDATA,YearBonus WHERE YB01='" & strYM & "' AND YB02=ST01(+) AND YB02=SD01(+) AND SD05='" & strSD05 & "' AND SD02<>'S' Order by T2 DESC,SD01"
'         If m_rs.State = 1 Then m_rs.Close
'         m_rs.CursorLocation = adUseClient
'         m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
'         If Not m_rs.EOF And Not m_rs.BOF Then
'            With m_rs
'               m_rs.MoveFirst
'               dblTotAmt = 0
'               dblTotCnt = 0
'
'               If ff > 0 Then Close #ff
'               ff = FreeFile
'               TempFileName = PUB_Getdesktop & "\" & strFileName
'               Open TempFileName For Output As ff
'
'               '首筆(第一筆)
'               For m_i = 1 To 12
'                  strTemp(m_i) = ""
'               Next m_i
'
'               strTemp(1) = "1"                                                            '區別碼        9(01) 固定為 1
'               strTemp(2) = "057"                                                          '企業編號1     9(03)
'               strTemp(3) = "F"                                                            '企業編號2     X(01) 固定為 'F'
'               strTemp(4) = "9634"                                                         '分行代號      9(04)
'               strTemp(5) = DBDATE(txt1(0))                                                '日期          9(08) 西元年月日YYYYMMDD
'               strTemp(6) = "2"                                                            '存提代號      9(01) 2為存款
'               strTemp(7) = "097"                                                          '摘要          9(03)
'               strTemp(8) = "CUST "                                                        '磁片來源      X(05) 自行轉出固定為 'CUST '
'               strTemp(9) = "1"                                                            '性質別        9(01) 固定為 1 (IN)
'               strTemp(10) = "04150022  "                                                  '公司統一編號  X(10)
'               strTemp(11) = "96340100346000"                                              '公司帳號      9(14)
'               strTemp(12) = String(79, " ")                                               '空白          X(79)
'
'               strText = ""
'               For m_i = 1 To 12
'                  strText = strText & strTemp(m_i)
'               Next m_i
'               Print #ff, strText
'
'               '明細(第二筆至倒數第二筆)
'               Do While Not m_rs.EOF
'
'                  For m_i = 1 To 18
'                     strTemp(m_i) = ""
'                  Next m_i
'
'                  strTemp(1) = "2"                                                               '區別碼        9(01) 固定為 2
'                  strTemp(2) = "057"                                                             '企業編號1     9(03)
'                  strTemp(3) = "F"                                                               '企業編號2     X(01) 固定為 'F'
'                  strTemp(4) = "9634"                                                            '分行代號      9(04)
'                  strTemp(5) = DBDATE(txt1(0))                                                   '日期          9(08) 西元年月日YYYYMMDD
'                  strTemp(6) = "2"                                                               '存提代號      9(01) 2為存款
'                  strTemp(7) = "097"                                                             '摘要          9(03)
'                  strTemp(8) = String(5, " ")                                                    '空白          X(05)
'                  strTemp(9) = Left(CheckStr(m_rs.Fields("SD06")) & "              ", 14)        '銀行帳號      9(14)
'                  strTemp(10) = Right("000000000000" & CheckStr(m_rs.Fields("T2")), 12) & "00"   '金額          9(12)V9(2)
'                  strTemp(11) = "99"                                                             '狀況代號      9(02) 轉出固定為 99
'                  strTemp(12) = "獎金      "                                                     '交易註記1     X(10)
'                  strTemp(13) = Left(CheckStr(m_rs.Fields("SD01")) & "          ", 10)           '交易註記2     X(10) 高所放員工編號
'                  strTemp(14) = Left(CheckStr(m_rs.Fields("ST26")) & "          ", 10)           '身份證字號    X(10)
'                  strTemp(15) = String(20, " ")                                                  '專用資料區    X(10)
'                  strTemp(16) = "Y"                                                              '身份證檢核註記X(01) 要檢查為 'Y'
'                  strTemp(17) = "  "                                                             '幣別          9(02) 台幣為空白
'                  strTemp(18) = String(21, " ")                                                  '空白欄        X(21)
'
'                  '合計欄之安全碼
'                  dblTotCnt = dblTotCnt + 1
'                  '總金額
'                  dblTotAmt = dblTotAmt + CheckStr(m_rs.Fields("T2"))
'
'                  strText = ""
'                  For m_i = 1 To 18
'                     strText = strText & strTemp(m_i)
'                  Next m_i
'                  Print #ff, strText
'                  m_rs.MoveNext
'               Loop
'
'               '總計：尾筆(最後一筆)
'               For m_i = 1 To 13
'                   strTemp(m_i) = ""
'               Next m_i
'
'               strTemp(1) = "3"                                                            '區別碼        9(01) 固定為 3
'               strTemp(2) = "057"                                                          '企業編號1     9(03)
'               strTemp(3) = "F"                                                            '企業編號2     X(01) 固定為 'F'
'               strTemp(4) = "9634"                                                         '分行代號      9(04)
'               strTemp(5) = DBDATE(txt1(0))                                                '日期          9(08) 西元年月日YYYYMMDD
'               strTemp(6) = "2"                                                            '存提代號      9(01) 2為存款
'               strTemp(7) = "097"                                                          '摘要          9(03)
'               strTemp(8) = String(5, " ")                                                 '空白欄        X(05)
'               strTemp(9) = Right("00000000000000" & CheckStr(dblTotAmt), 14) & "00"       '總金額        9(14)V9(2)
'               strTemp(10) = Right("0000000000" & CheckStr(dblTotCnt), 10)                 '總筆數        9(10)
'               strTemp(11) = String(16, " ")                                               '未成交總金額  9(16) 空白
'               strTemp(12) = String(10, " ")                                               '未成交總筆數  9(10) 空白
'               strTemp(13) = String(52, " ")                                               '空白欄        X(52)
'
'               strText = ""
'               For m_i = 1 To 13
'                  strText = strText & strTemp(m_i)
'               Next m_i
'               Print #ff, strText
'
'            End With
'            Close ff
'
'            MsgBox "年終獎金媒體已完成! 收到北所寄下來之媒體遞送單後即可送交銀行!! 辛苦了 !!", vbExclamation + vbOKOnly
'
'            Exit Sub
'         Else
'            MsgBox "無符合條件的資料!!!", vbExclamation + vbOKOnly
'            Exit Sub
'         End If
'
      Case "4"   '高所
         '留職停薪人員之年終資料改以現金發放故加入SD02<>'S'條件
         'modify by sonia 2018/1/11 +YB26
         m_str = "SELECT SD01,ST02,SD06,ST26,nvl(YB05,0)+nvl(YB06,0)+nvl(YB26,0)+nvl(YB08,0)-nvl(YB15,0)-nvl(YB16,0)-nvl(YB17,0)-nvl(YB25,0) as T2 " & _
                 "FROM STAFF,SALARYDATA,YearBonus WHERE YB01='" & strYM & "' AND YB02=ST01(+) AND YB02=SD01(+) AND SD05='" & strSD05 & "' AND SD02<>'S' Order by SD06"
         If m_rs.State = 1 Then m_rs.Close
         m_rs.CursorLocation = adUseClient
         m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
         If Not m_rs.EOF And Not m_rs.BOF Then
            With m_rs
               m_rs.MoveFirst
               dblTotAmt = 0
               dblTotCnt = 0
               
               If ff > 0 Then Close #ff
               ff = FreeFile
               TempFileName = PUB_Getdesktop & "\" & strFileName
               Open TempFileName For Output As ff
                
               '明細
               Do While Not m_rs.EOF
                   
                  For m_i = 1 To 10
                     strTemp(m_i) = ""
                  Next m_i
                  
                  strTemp(1) = "D"                                                              '轉帳類別      X(01) 固定為 D 年終獎金
                  strTemp(2) = "2"                                                              '檢查身分證號  X(01) 固定為 2 有就檢查
                  strTemp(3) = Left(CheckStr(m_rs.Fields("SD06")) & "              ", 13)       '銀行帳號      9(13)
                  strTemp(4) = Right("00000000000" & CheckStr(m_rs.Fields("T2")), 11) & "00"    '轉帳金額      9(11)V9(2)
                  strTemp(5) = Left(CheckStr(m_rs.Fields("ST26")) & "          ", 10)           '身份證字號    X(10)
                  strTemp(6) = String(23, " ")                                                  '存戶使用區    X(23) 固定為空白
                  strTemp(7) = String(3, " ")                                                   '客戶編號      X(03) 固定為空白
                  strTemp(8) = Right("00000000" & CheckStr(txt1(0)), 8)                         '入帳日期      X(08) 民國年月日YYYYMMDD前補0
                  strTemp(9) = String(4, " ")                                                   'Z類摘要       X(04) 固定為空白
                  strTemp(10) = "5252"                                                          '委託單位      X(04)  固定為 5252 七賢分行
                  
                  strText = ""
                  For m_i = 1 To 10
                     strText = strText & strTemp(m_i)
                  Next m_i
                  Print #ff, strText
                  m_rs.MoveNext
               Loop
                 
            End With
            Close ff
            
            MsgBox "年終獎金媒體已完成! 收到北所寄下來之媒體遞送單後即可送交銀行!! 辛苦了 !!", vbExclamation + vbOKOnly
            
            Exit Sub
         Else
            MsgBox "無符合條件的資料!!!", vbExclamation + vbOKOnly
            Exit Sub
         End If
'end 2017/3/28
   
   End Select
End Sub

Sub StrMenu3()
Dim ff As Integer
Dim strText As String
Dim strYM As String
Dim TempFileName As String
Dim dblTotSaveCode As Double

   If txt1(1) = "3" Then
      m_StrSQL = " and ob02='1'"     '端午
   Else
      m_StrSQL = " and ob02='2'"     '中秋
   End If
   
   strYM = Left(ChangeTStringToWString(txt1(0)), 4)
   
   Select Case pub_strUserOffice
'2013/7/29 CANCEL BY SONIA 中所由台北直接產生寄中所交中國信託
'      '2013/7/29 add by sonia
'      Case "2"   '中所
'         m_str = "SELECT SD01,ST02,SD06,ST26,nvl(ob05,0) T2 FROM Staff,SalaryData,ohbonus " & _
'                  "WHERE substr(ob01,1,4)='" & strYM & "' and ob05>0 AND ob03=st01(+) AND ob03=SD01(+) AND SD05='" & strSD05 & "' " & m_StrSQL & " Order by OB03 "
'         If m_rs.State = 1 Then m_rs.Close
'         m_rs.CursorLocation = adUseClient
'         m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
'         If Not m_rs.EOF And Not m_rs.BOF Then
'            With m_rs
'               m_rs.MoveFirst
'               dblTotAmt = 0
'               dblTotSaveCode = 0
'
'               If ff > 0 Then Close #ff
'               ff = FreeFile
'               TempFileName = PUB_Getdesktop & "\" & strFileName
'               Open TempFileName For Output As ff
'
'               '首筆(第一筆)
'               strTemp(1) = "159540272896"                                                     '公司帳號      X(12)
'               strTemp(2) = Right("0000000000" & TAIWANDATE(txt1(0)), 10)                      '入帳日期      9(10) 民國年月日YYYMMDD(不足10碼前補0)
'               strTemp(3) = "A"                                                                '收付業務別    X(01) 固定為 'A'
'
'               strText = ""
'               For m_i = 1 To 3
'                  strText = strText & strTemp(m_i)
'               Next m_i
'               Print #ff, strText
'
'               '明細(第二筆以後)
'               Do While Not m_rs.EOF
'
'                  For m_i = 1 To 6
'                     strTemp(m_i) = ""
'                  Next m_i
'
'                  strTemp(1) = Left(CheckStr(m_rs.Fields("SD06")) & "            ", 12)       '入帳帳號       X(12)
'                  strTemp(2) = Right("0000000000" & CheckStr(m_rs.Fields("T2")), 10)          '入帳金額       9(10)
'                  strTemp(3) = "          "                                                   '收款人姓名     X(10) 中國信託建議放空白
'                  strTemp(4) = Left(CheckStr(m_rs.Fields("ST26")) & "          ", 10)         '身份證字號     X(10)
'                  strTemp(5) = " "                                                            '入帳種類代碼   X(01) 空白：薪資
'                  strTemp(6) = "Y"                                                            '檢查功能碼     X(03) 檢查入帳戶之身份證字號
'
'                  strText = ""
'                  For m_i = 1 To 6
'                     strText = strText & strTemp(m_i)
'                  Next m_i
'                  Print #ff, strText
'                  m_rs.MoveNext
'               Loop
'
'            End With
'            Close ff
'
'            If txt1(1) = "3" Then
'               MsgBox "端午代金媒體已完成! 收到北所寄下來之媒體遞送單後即可送交銀行!! 辛苦了 !!", vbExclamation + vbOKOnly
'            Else
'               MsgBox "中秋代金媒體已完成! 收到北所寄下來之媒體遞送單後即可送交銀行!! 辛苦了 !!", vbExclamation + vbOKOnly
'            End If
'            Exit Sub
'         Else
'            MsgBox "無符合條件的資料!!!", vbExclamation + vbOKOnly
'            Exit Sub
'         End If
'      '2013/7/29 end
      
      Case "3"   '南所
         m_str = "SELECT SD01,ST02,SD06,nvl(ob05,0) T2 FROM Staff,SalaryData,ohbonus " & _
                  "WHERE substr(ob01,1,4)='" & strYM & "' and ob05>0 AND ob03=st01(+) AND ob03=SD01(+) AND SD05='" & strSD05 & "' " & m_StrSQL & " Order by OB03 "
         If m_rs.State = 1 Then m_rs.Close
         m_rs.CursorLocation = adUseClient
         m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
         If Not m_rs.EOF And Not m_rs.BOF Then
            With m_rs
               m_rs.MoveFirst
               dblTotAmt = 0
               dblTotSaveCode = 0
               
               If ff > 0 Then Close #ff
               ff = FreeFile
               TempFileName = PUB_Getdesktop & "\" & strFileName
               Open TempFileName For Output As ff
                
               Do While Not m_rs.EOF
                   
                  For m_i = 1 To 9
                     strTemp(m_i) = ""
                  Next m_i
                  
                  strTemp(1) = Right("00000000" & txt1(0), 8)                                 'Date       交易生效日 9(08) 民國年月日YYYYMMDD
                  'modify by sonia 2020/4/9
                  'strTemp(2) = "AALZ"                                                         'Code       機關代號   X(04)
                  strTemp(2) = "TAIE"                                                         'Code       機關代號   X(04)
                  'end 2020/4/9
                  strTemp(3) = Left(CheckStr(m_rs.Fields("SD06")) & "              ", 14)     'ID         帳號       X(14)
                  strTemp(4) = "C"                                                            'PayCode    存入或支出 X(01) 固定為 'C'
                  strTemp(5) = Right("00000000000" & CheckStr(m_rs.Fields("T2")), 11) & "00"  'PayAmtData 金額       9(11)V9(2)
                  strTemp(6) = Left(CheckStr(m_rs.Fields("ST02")) & "            ", 12)       'EmpID      員工編號   X(15) 南所放員工姓名,故碼數改12
                  strTemp(7) = "900"                                                          'CostCode   費用代碼   X(03) 固定為 '900'
                  strTemp(9) = String(18, " ")                                                'Filler     空白       X(18)
                  
                  '安全碼=帳號欄第1碼*金額欄第1碼+帳號欄第2碼*金額欄第2碼+帳號欄第3碼*金額欄第3碼+......帳號欄第13碼*金額欄第13碼+帳號欄第14碼
                  For m_i = 1 To 13
                     strTemp(8) = Val(strTemp(8)) + Val(Mid(strTemp(3), m_i, 1)) * Val(Mid(strTemp(5), m_i, 1))
                  Next m_i
                  strTemp(8) = Val(strTemp(8)) + Val(Mid(strTemp(3), 14, 1))
                  strTemp(8) = Right("0000" & CheckStr(strTemp(8)), 4)                        'SaftyCode  安全碼     X(04)
                  
                  '合計欄之安全碼
                  dblTotSaveCode = dblTotSaveCode + Val(strTemp(8))
                  '總金額
                  dblTotAmt = dblTotAmt + CheckStr(m_rs.Fields("T2"))
                  
                  strText = ""
                  For m_i = 1 To 9
                     strText = strText & strTemp(m_i)
                  Next m_i
                  Print #ff, strText
                  m_rs.MoveNext
               Loop
                
               '總計：最後一筆
               For m_i = 1 To 9
                   strTemp(m_i) = ""
               Next m_i
               strTemp(1) = Right("00000000" & txt1(0), 8)                                    'Date       交易生效日 9(08) 民國年月日YYYYMMDD
               'modify by sonia 2020/4/9
               'strTemp(2) = "AALZ"                                                            'Code       機關代號   X(04)
               strTemp(2) = "TAIE"                                                            'Code       機關代號   X(04)
               'end 2020/4/9
               strTemp(3) = "99999999999999"                                                  'ID         帳號       X(14)
               strTemp(4) = "C"                                                               'PayCode    存入或支出 X(01) 固定為 'C'
               strTemp(5) = Right("00000000000" & CheckStr(dblTotAmt), 11) & "00"             'PayAmtData 金額       9(11)V9(2)
               strTemp(6) = "999999999999999"                                                 'EmpID      員工編號   X(15)
               strTemp(7) = "900"                                                             'CostCode   費用代碼   X(03) 固定為 '900'
               strTemp(9) = String(18, " ")                                                   'Filler     空白       X(18)
               
               '合計欄之安全碼=SUM(明細安全碼)/11
               dblTotSaveCode = dblTotSaveCode / 11
               strTemp(8) = Right("0000" & CheckStr(Int(dblTotSaveCode)), 4)                  'SaftyCode  安全碼     X(04)
               
               strText = ""
               For m_i = 1 To 9
                  strText = strText & strTemp(m_i)
               Next m_i
               Print #ff, strText
                 
            End With
            Close ff
            
            If txt1(1) = "3" Then
               MsgBox "端午代金媒體已完成! 收到北所寄下來之明細表後即可送交銀行!! 辛苦了 !!", vbExclamation + vbOKOnly
            Else
               MsgBox "中秋代金媒體已完成! 收到北所寄下來之明細表後即可送交銀行!! 辛苦了 !!", vbExclamation + vbOKOnly
            End If
            Exit Sub
         Else
            MsgBox "無符合條件的資料!!!", vbExclamation + vbOKOnly
            Exit Sub
         End If
      
'modify by sonia 2017/3/28 高所由彰銀改合庫
'      Case "4"   '高所
'         m_str = "SELECT SD01,ST02,SD06,ST26,nvl(ob05,0) T2 FROM Staff,SalaryData,ohbonus " & _
'                  "WHERE substr(ob01,1,4)='" & strYM & "' and ob05>0 AND ob03=st01(+) AND ob03=SD01(+) AND SD05='" & strSD05 & "' " & m_StrSQL & " Order by T2 DESC,OB03 "
'         If m_rs.State = 1 Then m_rs.Close
'         m_rs.CursorLocation = adUseClient
'         m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
'         If Not m_rs.EOF And Not m_rs.BOF Then
'            With m_rs
'               m_rs.MoveFirst
'               dblTotAmt = 0
'               dblTotCnt = 0
'
'               If ff > 0 Then Close #ff
'               ff = FreeFile
'               TempFileName = PUB_Getdesktop & "\" & strFileName
'               Open TempFileName For Output As ff
'
'               '首筆(第一筆)
'               For m_i = 1 To 12
'                  strTemp(m_i) = ""
'               Next m_i
'
'               strTemp(1) = "1"                                                            '區別碼        9(01) 固定為 1
'               strTemp(2) = "057"                                                          '企業編號1     9(03)
'               strTemp(3) = "F"                                                            '企業編號2     X(01) 固定為 'F'
'               strTemp(4) = "9634"                                                         '分行代號      9(04)
'               strTemp(5) = DBDATE(txt1(0))                                                '日期          9(08) 西元年月日YYYYMMDD
'               strTemp(6) = "2"                                                            '存提代號      9(01) 2為存款
'               strTemp(7) = "097"                                                          '摘要          9(03)
'               strTemp(8) = "CUST "                                                        '磁片來源      X(05) 自行轉出固定為 'CUST '
'               strTemp(9) = "1"                                                            '性質別        9(01) 固定為 1 (IN)
'               strTemp(10) = "04150022  "                                                  '公司統一編號  X(10)
'               strTemp(11) = "96340100346000"                                              '公司帳號      9(14)
'               strTemp(12) = String(79, " ")                                               '空白          X(79)
'
'               strText = ""
'               For m_i = 1 To 12
'                  strText = strText & strTemp(m_i)
'               Next m_i
'               Print #ff, strText
'
'               '明細(第二筆至倒數第二筆)
'               Do While Not m_rs.EOF
'
'                  For m_i = 1 To 18
'                     strTemp(m_i) = ""
'                  Next m_i
'
'                  strTemp(1) = "2"                                                               '區別碼        9(01) 固定為 2
'                  strTemp(2) = "057"                                                             '企業編號1     9(03)
'                  strTemp(3) = "F"                                                               '企業編號2     X(01) 固定為 'F'
'                  strTemp(4) = "9634"                                                            '分行代號      9(04)
'                  strTemp(5) = DBDATE(txt1(0))                                                   '日期          9(08) 西元年月日YYYYMMDD
'                  strTemp(6) = "2"                                                               '存提代號      9(01) 2為存款
'                  strTemp(7) = "097"                                                             '摘要          9(03)
'                  strTemp(8) = String(5, " ")                                                    '空白          X(05)
'                  strTemp(9) = Left(CheckStr(m_rs.Fields("SD06")) & "              ", 14)        '銀行帳號      9(14)
'                  strTemp(10) = Right("000000000000" & CheckStr(m_rs.Fields("T2")), 12) & "00"   '金額          9(12)V9(2)
'                  strTemp(11) = "99"                                                             '狀況代號      9(02) 轉出固定為 99
'                  strTemp(12) = "獎金      "                                                     '交易註記1     X(10)
'                  strTemp(13) = Left(CheckStr(m_rs.Fields("SD01")) & "          ", 10)           '交易註記2     X(10) 高所放員工編號
'                  strTemp(14) = Left(CheckStr(m_rs.Fields("ST26")) & "          ", 10)           '身份證字號    X(10)
'                  strTemp(15) = String(20, " ")                                                  '專用資料區    X(10)
'                  strTemp(16) = "Y"                                                              '身份證檢核註記X(01) 要檢查為 'Y'
'                  strTemp(17) = "  "                                                             '幣別          9(02) 台幣為空白
'                  strTemp(18) = String(21, " ")                                                  '空白欄        X(21)
'
'                  '合計欄之安全碼
'                  dblTotCnt = dblTotCnt + 1
'                  '總金額
'                  dblTotAmt = dblTotAmt + CheckStr(m_rs.Fields("T2"))
'
'                  strText = ""
'                  For m_i = 1 To 18
'                     strText = strText & strTemp(m_i)
'                  Next m_i
'                  Print #ff, strText
'                  m_rs.MoveNext
'               Loop
'
'               '總計：尾筆(最後一筆)
'               For m_i = 1 To 13
'                   strTemp(m_i) = ""
'               Next m_i
'
'               strTemp(1) = "3"                                                            '區別碼        9(01) 固定為 3
'               strTemp(2) = "057"                                                          '企業編號1     9(03)
'               strTemp(3) = "F"                                                            '企業編號2     X(01) 固定為 'F'
'               strTemp(4) = "9634"                                                         '分行代號      9(04)
'               strTemp(5) = DBDATE(txt1(0))                                                '日期          9(08) 西元年月日YYYYMMDD
'               strTemp(6) = "2"                                                            '存提代號      9(01) 2為存款
'               strTemp(7) = "097"                                                          '摘要          9(03)
'               strTemp(8) = String(5, " ")                                                 '空白欄        X(05)
'               strTemp(9) = Right("00000000000000" & CheckStr(dblTotAmt), 14) & "00"       '總金額        9(14)V9(2)
'               strTemp(10) = Right("0000000000" & CheckStr(dblTotCnt), 10)                 '總筆數        9(10)
'               strTemp(11) = String(16, " ")                                               '未成交總金額  9(16) 空白
'               strTemp(12) = String(10, " ")                                               '未成交總筆數  9(10) 空白
'               strTemp(13) = String(52, " ")                                               '空白欄        X(52)
'
'               strText = ""
'               For m_i = 1 To 13
'                  strText = strText & strTemp(m_i)
'               Next m_i
'               Print #ff, strText
'
'            End With
'            Close ff
'
'            If txt1(1) = "3" Then
'               MsgBox "端午代金媒體已完成! 收到北所寄下來之媒體遞送單後即可送交銀行!! 辛苦了 !!", vbExclamation + vbOKOnly
'            Else
'               MsgBox "中秋代金媒體已完成! 收到北所寄下來之媒體遞送單後即可送交銀行!! 辛苦了 !!", vbExclamation + vbOKOnly
'            End If
'            Exit Sub
'         Else
'            MsgBox "無符合條件的資料!!!", vbExclamation + vbOKOnly
'            Exit Sub
'         End If
      Case "4"   '高所
         m_str = "SELECT SD01,ST02,SD06,ST26,nvl(ob05,0) T2 FROM Staff,SalaryData,ohbonus " & _
                  "WHERE substr(ob01,1,4)='" & strYM & "' and ob05>0 AND ob03=st01(+) AND ob03=SD01(+) AND SD05='" & strSD05 & "' " & m_StrSQL & " Order by SD06 "
         If m_rs.State = 1 Then m_rs.Close
         m_rs.CursorLocation = adUseClient
         m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
         If Not m_rs.EOF And Not m_rs.BOF Then
            With m_rs
               m_rs.MoveFirst
               dblTotAmt = 0
               dblTotCnt = 0
               
               If ff > 0 Then Close #ff
               ff = FreeFile
               TempFileName = PUB_Getdesktop & "\" & strFileName
               Open TempFileName For Output As ff
                
               '明細
               Do While Not m_rs.EOF
                   
                  For m_i = 1 To 10
                     strTemp(m_i) = ""
                  Next m_i
                  
                  strTemp(1) = "T"                                                              '轉帳類別      X(01) 固定為 T 福利金
                  strTemp(2) = "2"                                                              '檢查身分證號  X(01) 固定為 2 有就檢查
                  strTemp(3) = Left(CheckStr(m_rs.Fields("SD06")) & "              ", 13)       '銀行帳號      9(13)
                  strTemp(4) = Right("00000000000" & CheckStr(m_rs.Fields("T2")), 11) & "00"    '轉帳金額      9(11)V9(2)
                  strTemp(5) = Left(CheckStr(m_rs.Fields("ST26")) & "          ", 10)           '身份證字號    X(10)
                  strTemp(6) = String(23, " ")                                                  '存戶使用區    X(23) 固定為空白
                  strTemp(7) = String(3, " ")                                                   '客戶編號      X(03) 固定為空白
                  strTemp(8) = Right("00000000" & CheckStr(txt1(0)), 8)                         '入帳日期      X(08) 民國年月日YYYYMMDD前補0
                  strTemp(9) = String(4, " ")                                                   'Z類摘要       X(04) 固定為空白
                  strTemp(10) = "5252"                                                          '委託單位      X(04)  固定為 5252 七賢分行
                  
                  strText = ""
                  For m_i = 1 To 10
                     strText = strText & strTemp(m_i)
                  Next m_i
                  Print #ff, strText
                  m_rs.MoveNext
               Loop
                
            End With
            Close ff
            
            If txt1(1) = "3" Then
               MsgBox "端午代金媒體已完成! 收到北所寄下來之媒體遞送單後即可送交銀行!! 辛苦了 !!", vbExclamation + vbOKOnly
            Else
               MsgBox "中秋代金媒體已完成! 收到北所寄下來之媒體遞送單後即可送交銀行!! 辛苦了 !!", vbExclamation + vbOKOnly
            End If
            Exit Sub
         Else
            MsgBox "無符合條件的資料!!!", vbExclamation + vbOKOnly
            Exit Sub
         End If
'end 2017/3/28

   End Select
End Sub

Private Sub Form_Load()
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim strSql As String, i As Integer, j As Integer
Dim strSystemKind As String
   
   MoveFormToCenter Me
   strSystemKind = GetSystemKindByNick
   
   strFileName = "": strSD05 = ""
   Select Case pub_strUserOffice
      '2013/7/29 add by sonia
      Case "2"  '中所
         Label3.Caption = "PS：媒體檔案放在桌面, 檔名為 Salary.txt"
         strFileName = "Salary.txt"
         Label5.Caption = "請輸入媒體遞送單上之撥薪日期"
         Label5.Visible = True
         strSD05 = "4"
      '2013/7/29 end
      Case "3"  '南所
         Label3.Caption = "PS：媒體檔案放在桌面, 檔名為 Salary.txt"
         Label5.Visible = False
         strFileName = "Salary.txt"
         strSD05 = "5"
      Case "4"  '高所
         'modify by sonia 2017/3/28 彰銀改合庫(不要副檔名)
         'Label3.Caption = "PS：媒體檔案放在桌面, 檔名為 PCCUT.TXT"
         'strFileName = "PCCUT.TXT"
         Label3.Caption = "PS：媒體檔案放在桌面, 檔名為 Salary"
         strFileName = "Salary"
         'end 2017/3/28
         Label5.Caption = "請輸入媒體遞送單上之預定撥帳日期"
         Label5.Visible = True
         strSD05 = "6"
      Case Else
         MsgBox "所別錯誤 !!!", vbExclamation + vbOKOnly
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Frmacc7140 = Nothing
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

