VERSION 5.00
Begin VB.Form frm160101 
   Appearance      =   0  '平面
   BorderStyle     =   1  '單線固定
   Caption         =   "員工名冊列印"
   ClientHeight    =   3590
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3590
   ScaleWidth      =   5060
   Begin VB.ComboBox Combo2 
      Height          =   260
      ItemData        =   "frm160101.frx":0000
      Left            =   1890
      List            =   "frm160101.frx":0013
      Style           =   2  '單純下拉式
      TabIndex        =   5
      Top             =   1800
      Width           =   1910
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   5
      Left            =   1890
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1440
      Width           =   320
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   150
      TabIndex        =   15
      Top             =   2940
      Width           =   4730
      Begin VB.ComboBox Combo1 
         Height          =   260
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
         TabIndex        =   16
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "含離職人員"
      Height          =   225
      Left            =   900
      TabIndex        =   7
      Top             =   2520
      Width           =   1845
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   4
      Left            =   1890
      MaxLength       =   1
      TabIndex        =   6
      Text            =   "1"
      Top             =   2160
      Width           =   260
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   2670
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1080
      Width           =   645
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   1890
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1080
      Width           =   645
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   2490
      MaxLength       =   3
      TabIndex        =   1
      Top             =   720
      Width           =   465
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   0
      Left            =   1890
      MaxLength       =   3
      TabIndex        =   0
      Top             =   720
      Width           =   465
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   3830
      TabIndex        =   10
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   2820
      TabIndex        =   9
      Top             =   90
      Width           =   915
   End
   Begin VB.Label Label7 
      Caption         =   "(1.北 2.中 3.南 4.高 5.其他)"
      Height          =   250
      Left            =   2280
      TabIndex        =   19
      Top             =   1470
      Width           =   2200
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "公司別："
      Height          =   180
      Left            =   1110
      TabIndex        =   18
      Top             =   1830
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "所　別："
      Height          =   180
      Left            =   1110
      TabIndex        =   17
      Top             =   1470
      Width           =   720
   End
   Begin VB.Line Line2 
      X1              =   2190
      X2              =   2550
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      X1              =   2130
      X2              =   3120
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(1.明細表 2.名冊)"
      Height          =   180
      Left            =   2190
      TabIndex        =   14
      Top             =   2190
      Width           =   1340
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "報表類別："
      Height          =   180
      Left            =   930
      TabIndex        =   13
      Top             =   2190
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   0
      Left            =   930
      TabIndex        =   12
      Top             =   1110
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部門代號："
      Height          =   180
      Left            =   930
      TabIndex        =   11
      Top             =   750
      Width           =   900
   End
End
Attribute VB_Name = "frm160101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/20 日期欄已修改
'Create by nickc 2008/03/18
Option Explicit

Dim m_rs As New ADODB.Recordset
Dim m_rs2 As New ADODB.Recordset
Dim m_str As String
Dim m_str2 As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 53) As Integer
Dim strTemp(1 To 53) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim LongPrintCurCnt As Long


Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
'        If txt1(0) = "" And txt1(1) = "" And txt1(2) = "" And txt1(3) = "" Then
'            MsgBox "請至少輸入一項列印條件！", vbInformation, "操作錯誤！"
'            txt1(0).SetFocus
'            Exit Sub
'        End If
        If txt1(4) = "" Then
            MsgBox "報表類別不可以空白！", vbInformation, "操作錯誤！"
            txt1(4).SetFocus
            Exit Sub
        End If
        
        Set Printer = Printers(Combo1.ListIndex)
        
        Screen.MousePointer = vbHourglass
        m_StrSQL = ""
        If txt1(0) <> "" Then
            'Modify By Sindy 2023/12/28 部門調整改抓ST93
            'm_StrSQL = m_StrSQL & " and st03>='" & txt1(0) & "'"
            m_StrSQL = m_StrSQL & " and st93>='" & txt1(0) & "'"
        End If
        If txt1(1) <> "" Then
            'Modify By Sindy 2023/12/28 部門調整改抓ST93
            'm_StrSQL = m_StrSQL & " and st03<='" & txt1(1) & "'"
            m_StrSQL = m_StrSQL & " and st93<='" & txt1(1) & "'"
        End If
        If txt1(2) <> "" Then
            m_StrSQL = m_StrSQL & " and st01>='" & txt1(2) & "'"
        End If
        If txt1(3) <> "" Then
            m_StrSQL = m_StrSQL & " and st01<='" & txt1(3) & "'"
        End If
        If Check1.Value = vbUnchecked Then
            m_StrSQL = m_StrSQL & " and st04='1'"
        End If
        'Add By Sindy 2025/9/15
        '分所
        If txt1(5) <> "" Then
            m_StrSQL = m_StrSQL & " and st06='" & txt1(5) & "'"
        End If
        '公司別
        If Trim(Combo2.Text) <> "" Then
            m_StrSQL = m_StrSQL & " and (sd19='" & Left(Trim(Combo2.Text), 1) & "' or sd19 is null)"
        End If
        '2025/9/15 END
        Select Case txt1(4)
        Case "1"
                StrMenu1
        Case "2"
                StrMenu2
        End Select
        Screen.MousePointer = vbDefault
Case 1
        Unload Me
End Select
End Sub

'明細表
Sub StrMenu1()
Printer.Orientation = 2
'Printer.FontName = "標楷體"
'Modify By Sindy 2023/12/28 部門調整改抓ST93
m_str = "select st.*,nvl(A0922,'(舊)'||A0902) A0902,a1.ac03 a1ac03,a2.ac03 a2ac03,a3.ac03 a3ac03,a4.ac03 a4ac03" & _
         " from staff st,acc090 a09,acc090NEW,allcode A1,allcode A2,allcode A3,allcode A4,SalaryData" & _
         " where ST01=SD01 and ((SD02 not in('P','F') or SD02 is null) or ST01='68007')" & _
         " and st03=a0901(+) and st93=a0921(+) and '01'=a1.ac01(+) and st20=a1.ac02(+) and '02'=a2.ac01(+)" & _
         " and st21=a2.ac02(+) and '06'=a3.ac01(+) and st27=a3.ac02(+) and '03'=a4.ac01(+)" & _
         " and st37=a4.ac02(+)" & m_StrSQL & " order by nvl(st.st93,st.st03),st01"
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
LongPrintCurCnt = 0
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        .MoveFirst
        PrintTitle
        Do While Not .EOF
            LongPrintCurCnt = LongPrintCurCnt + 1
            
            For m_i = 1 To 53
                strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = CheckStr(.Fields("st01"))
            strTemp(2) = CheckStr(.Fields("st02"))
            strTemp(3) = IIf(CheckStr(.Fields("st06")) = "1", "北所", IIf(CheckStr(.Fields("st06")) = "2", "中所", IIf(CheckStr(.Fields("st06")) = "3", "南所", IIf(CheckStr(.Fields("st06")) = "4", "高所", "其他"))))
            strTemp(4) = CheckStr(.Fields("a0902"))
            strTemp(5) = CheckStr(.Fields("a1ac03"))
            strTemp(6) = CheckStr(.Fields("a2ac03"))
            strTemp(7) = IIf(CheckStr(.Fields("st22")) = "M", "男", IIf(CheckStr(.Fields("st22")) = "F", "女", ""))
            strTemp(8) = ChangeTStringToTDateString(TAIWANDATE(CheckStr(.Fields("st23"))))
            strTemp(9) = CheckStr(.Fields("a3ac03"))
            strTemp(10) = CheckStr(.Fields("st25"))
            strTemp(11) = ChangeTStringToTDateString(TAIWANDATE(CheckStr(.Fields("st13"))))
            strTemp(12) = CheckStr(.Fields("st26"))
            strTemp(13) = ChangeTStringToTDateString(TAIWANDATE(CheckStr(.Fields("st28")))) & "-" & ChangeTStringToTDateString(TAIWANDATE(CheckStr(.Fields("st29"))))
            strTemp(14) = IIf(CheckStr(.Fields("st24")) = "L", "本國", IIf(CheckStr(.Fields("st24")) = "F", "外國", ""))
            '第二行
            strTemp(15) = ""
            strTemp(16) = CheckStr(.Fields("st09"))
            strTemp(17) = CheckStr(.Fields("st10"))
            strTemp(18) = CheckStr(.Fields("st33"))
            strTemp(19) = CheckStr(.Fields("st08"))
            strTemp(20) = CheckStr(.Fields("st19"))
            strTemp(21) = CheckStr(.Fields("st40"))
            strTemp(22) = ChangeTStringToTDateString(TAIWANDATE(CheckStr(.Fields("st31"))))
            strTemp(23) = ChangeTStringToTDateString(TAIWANDATE(CheckStr(.Fields("st32"))))
            '第三行
            strTemp(24) = CheckStr(.Fields("st35"))
            strTemp(25) = CheckStr(.Fields("st36"))
            strTemp(26) = CheckStr(.Fields("st34"))
            strTemp(27) = CheckStr(.Fields("a4ac03"))
            strTemp(28) = CheckStr(.Fields("st38"))
            '第四行
            strTemp(33) = CheckStr(.Fields("st39"))
            strTemp(34) = GetStaffName(CheckStr(.Fields("st30")), True)
            '第六行
            strTemp(41) = ChangeTStringToTDateString(TAIWANDATE(CheckStr(.Fields("st41"))))
            '眷屬資料
            'Modify By Sindy 2010/5/6 已有刪除日期的資料不顯示
            'm_str2 = "select * from staff_relation where sr01='" & CheckStr(.Fields("st01")) & "' order by sr03 "
            m_str2 = "select * from staff_relation where sr01='" & CheckStr(.Fields("st01")) & "' and (sr12 is null or sr12=0) order by sr03 "
            If m_rs2.State = 1 Then m_rs2.Close
            m_rs2.CursorLocation = adUseClient
            m_rs2.Open m_str2, cnnConnection, adOpenStatic, adLockReadOnly
            If Not m_rs2.EOF And Not m_rs2.BOF Then
                m_rs2.MoveFirst
                Do While Not m_rs2.EOF
                    Select Case CheckStr(m_rs2.Fields("sr03"))
                    Case "1"
                            strTemp(29) = CheckStr(m_rs2.Fields("sr04"))
                            strTemp(30) = CheckStr(m_rs2.Fields("sr09"))
                            strTemp(31) = CheckStr(m_rs2.Fields("sr10"))
                            strTemp(32) = CheckStr(m_rs2.Fields("sr11"))
                    Case "2"
                            strTemp(35) = CheckStr(m_rs2.Fields("sr04"))
                            strTemp(36) = CheckStr(m_rs2.Fields("sr09"))
                            strTemp(37) = CheckStr(m_rs2.Fields("sr10"))
                            strTemp(38) = CheckStr(m_rs2.Fields("sr11"))
                    Case "3"
                            strTemp(40) = CheckStr(m_rs2.Fields("sr04"))
                    Case "4"
                            If strTemp(42) = "" And strTemp(43) = "" And strTemp(44) = "" Then
                                strTemp(42) = CheckStr(m_rs2.Fields("sr04"))
                                strTemp(43) = IIf(CheckStr(m_rs2.Fields("sr05")) = "M", "男", IIf(CheckStr(m_rs2.Fields("sr05")) = "F", "女", ""))
                                strTemp(44) = ChangeTStringToTDateString(TAIWANDATE(CheckStr(m_rs2.Fields("sr06"))))
                            ElseIf strTemp(45) = "" And strTemp(46) = "" And strTemp(47) = "" Then
                                strTemp(45) = CheckStr(m_rs2.Fields("sr04"))
                                strTemp(46) = IIf(CheckStr(m_rs2.Fields("sr05")) = "M", "男", IIf(CheckStr(m_rs2.Fields("sr05")) = "F", "女", ""))
                                strTemp(47) = ChangeTStringToTDateString(TAIWANDATE(CheckStr(m_rs2.Fields("sr06"))))
                            ElseIf strTemp(48) = "" And strTemp(49) = "" And strTemp(50) = "" Then
                                strTemp(48) = CheckStr(m_rs2.Fields("sr04"))
                                strTemp(49) = IIf(CheckStr(m_rs2.Fields("sr05")) = "M", "男", IIf(CheckStr(m_rs2.Fields("sr05")) = "F", "女", ""))
                                strTemp(50) = ChangeTStringToTDateString(TAIWANDATE(CheckStr(m_rs2.Fields("sr06"))))
                            Else
                                strTemp(51) = CheckStr(m_rs2.Fields("sr04"))
                                strTemp(52) = IIf(CheckStr(m_rs2.Fields("sr05")) = "M", "男", IIf(CheckStr(m_rs2.Fields("sr05")) = "F", "女", ""))
                                strTemp(53) = ChangeTStringToTDateString(TAIWANDATE(CheckStr(m_rs2.Fields("sr06"))))
                            End If
                     End Select
                    m_rs2.MoveNext
                Loop
            End If
            '歷年考績
            'Modify by Morgan 2009/6/19 原來復職的資料會有問題，改甲等也存檔可直接抓考績檔
            'm_str2 = "SELECT myyear,decode(ym02,'1','優','3','乙','4','丙','甲')  FROM yearmerit,(select distinct to_number(substr(st01,1,2))+1911 as MyYear from staff where st01>='" & CheckStr(.Fields("st01")) & "' and st01<rtrim(ltrim(to_char(to_number(to_char(sysdate,'YYYY'))-1911)))) AA where AA.MyYear=ym01(+) and '" & CheckStr(.Fields("st01")) & "'=ym02(+)  order by myyear "
            'MODIFY BY SONIA 2015/12/25 考績檔ym02加入*不參加考核
            m_str2 = "SELECT ym01 myyear,decode(ym02,'1','優','2','甲','3','乙','4','丙','*','不參加考核',ym02) FROM yearmerit where ym03='" & CheckStr(.Fields("st01")) & "' order by myyear "
            
            If m_rs2.State = 1 Then m_rs2.Close
            m_rs2.CursorLocation = adUseClient
            m_rs2.Open m_str2, cnnConnection, adOpenStatic, adLockReadOnly
            If Not m_rs2.EOF And Not m_rs2.BOF Then
                m_rs2.MoveFirst
                Do While Not m_rs2.EOF
                    strTemp(39) = strTemp(39) & CheckStr(m_rs2.Fields(1)) & " "
                    m_rs2.MoveNext
                Loop
            End If
            PrintDetail
            '每四筆才換新頁
            If LongPrintCurCnt Mod 4 = 0 Then
               If .AbsolutePosition <> .RecordCount Then
                  Printer.NewPage
                  PrintTitle
               End If
            End If
            .MoveNext
        Loop
    End With
Else
   MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
   Exit Sub
End If
Printer.EndDoc
ShowPrintOk
End Sub

Sub GetPleft()
PLeft(1) = 200
PLeft(2) = 900
PLeft(3) = 2100
PLeft(4) = 2700
PLeft(5) = 4700
PLeft(6) = 6200
PLeft(7) = 8200
PLeft(8) = 8750
PLeft(9) = 9750
PLeft(10) = 11400
PLeft(11) = 11900
PLeft(12) = 12900
PLeft(13) = 14000 '14200
PLeft(14) = 15700
PLeft(15) = 200
PLeft(16) = 900
PLeft(17) = 2100
PLeft(18) = 3300
PLeft(19) = 4300
PLeft(20) = 9300
PLeft(21) = 10800
PLeft(22) = 12300
PLeft(23) = 13800
PLeft(24) = 900
PLeft(25) = 3300
PLeft(26) = 4300
PLeft(27) = 9300
PLeft(28) = 12300
PLeft(29) = 900
PLeft(30) = 2100
PLeft(31) = 3300
PLeft(32) = 4300
PLeft(33) = 9300
PLeft(34) = 12300
PLeft(35) = 900
PLeft(36) = 2100
PLeft(37) = 3300
PLeft(38) = 4300
PLeft(39) = 9300
PLeft(40) = 900
PLeft(41) = 2000
PLeft(42) = 3000
PLeft(43) = 4000
PLeft(44) = 5000
PLeft(45) = 6000
PLeft(46) = 7000
PLeft(47) = 8000
PLeft(48) = 9000
PLeft(49) = 10000
PLeft(50) = 11000
PLeft(51) = 12000
PLeft(52) = 13000
PLeft(53) = 14000
End Sub

Sub PrintTitle()
GetPleft

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("員工基本資料明細表") / 2)
Printer.CurrentY = 300
Printer.Print "員工基本資料明細表"
Printer.Font.Size = 10
Printer.Font.Underline = False
Printer.FontBold = False
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 600
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "頁　　次：" & Printer.Page
iLine = 4
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "所在"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iLine * 300
Printer.Print "　試用期間"
iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "員工"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "姓　名"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "所別"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "部　　門"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iLine * 300
Printer.Print "職　稱"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iLine * 300
Printer.Print "職　位"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iLine * 300
Printer.Print "性別"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iLine * 300
Printer.Print "出生日期"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iLine * 300
Printer.Print "出生地"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iLine * 300
Printer.Print "血型"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iLine * 300
Printer.Print "入所日期"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iLine * 300
Printer.Print "身分證字號"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iLine * 300
Printer.Print "起　　　迄"
Printer.CurrentX = PLeft(14)
Printer.CurrentY = iLine * 300
Printer.Print "國籍"
iLine = iLine + 1
Printer.CurrentX = PLeft(15)
Printer.CurrentY = iLine * 300
Printer.Print "編號"
Printer.CurrentX = PLeft(16)
Printer.CurrentY = iLine * 300
Printer.Print "電　話"
Printer.CurrentX = PLeft(17)
Printer.CurrentY = iLine * 300
Printer.Print "傳　真"
Printer.CurrentX = PLeft(18)
Printer.CurrentY = iLine * 300
Printer.Print "郵遞區號"
Printer.CurrentX = PLeft(19)
Printer.CurrentY = iLine * 300
Printer.Print "通訊地址"
Printer.CurrentX = PLeft(20)
Printer.CurrentY = iLine * 300
Printer.Print "行動電話"
Printer.CurrentX = PLeft(21)
Printer.CurrentY = iLine * 300
Printer.Print "可休特別假"
Printer.CurrentX = PLeft(22)
Printer.CurrentY = iLine * 300
Printer.Print "加保日期"
Printer.CurrentX = PLeft(23)
Printer.CurrentY = iLine * 300
Printer.Print "上次簽保日期"
iLine = iLine + 1
Printer.CurrentX = PLeft(24)
Printer.CurrentY = iLine * 300
Printer.Print "電　話"
Printer.CurrentX = PLeft(25)
Printer.CurrentY = iLine * 300
Printer.Print "郵遞區號"
Printer.CurrentX = PLeft(26)
Printer.CurrentY = iLine * 300
Printer.Print "戶籍地址"
Printer.CurrentX = PLeft(27)
Printer.CurrentY = iLine * 300
Printer.Print "最高學歷"
Printer.CurrentX = PLeft(28)
Printer.CurrentY = iLine * 300
Printer.Print "畢業學校"
iLine = iLine + 1
Printer.CurrentX = PLeft(29)
Printer.CurrentY = iLine * 300
Printer.Print "父　親"
Printer.CurrentX = PLeft(30)
Printer.CurrentY = iLine * 300
Printer.Print "電　話"
Printer.CurrentX = PLeft(31)
Printer.CurrentY = iLine * 300
Printer.Print "郵遞區號"
Printer.CurrentX = PLeft(32)
Printer.CurrentY = iLine * 300
Printer.Print "地　　址"
Printer.CurrentX = PLeft(33)
Printer.CurrentY = iLine * 300
Printer.Print "科　系"
Printer.CurrentX = PLeft(34)
Printer.CurrentY = iLine * 300
Printer.Print "職務代理人"
iLine = iLine + 1
Printer.CurrentX = PLeft(35)
Printer.CurrentY = iLine * 300
Printer.Print "母　親"
Printer.CurrentX = PLeft(36)
Printer.CurrentY = iLine * 300
Printer.Print "電　話"
Printer.CurrentX = PLeft(37)
Printer.CurrentY = iLine * 300
Printer.Print "郵遞區號"
Printer.CurrentX = PLeft(38)
Printer.CurrentY = iLine * 300
Printer.Print "地　　址"
Printer.CurrentX = PLeft(39)
Printer.CurrentY = iLine * 300
Printer.Print "歷年考績 ( ----> ) "
iLine = iLine + 1
Printer.CurrentX = PLeft(40)
Printer.CurrentY = iLine * 300
Printer.Print "配偶姓名"
Printer.CurrentX = PLeft(41)
Printer.CurrentY = iLine * 300
Printer.Print "結婚日期"
Printer.CurrentX = PLeft(42)
Printer.CurrentY = iLine * 300
Printer.Print "子女姓名"
Printer.CurrentX = PLeft(43)
Printer.CurrentY = iLine * 300
Printer.Print "性別"
Printer.CurrentX = PLeft(44)
Printer.CurrentY = iLine * 300
Printer.Print "出生日期"
Printer.CurrentX = PLeft(45)
Printer.CurrentY = iLine * 300
Printer.Print "子女姓名"
Printer.CurrentX = PLeft(46)
Printer.CurrentY = iLine * 300
Printer.Print "性別"
Printer.CurrentX = PLeft(47)
Printer.CurrentY = iLine * 300
Printer.Print "出生日期"
Printer.CurrentX = PLeft(48)
Printer.CurrentY = iLine * 300
Printer.Print "子女姓名"
Printer.CurrentX = PLeft(49)
Printer.CurrentY = iLine * 300
Printer.Print "性別"
Printer.CurrentX = PLeft(50)
Printer.CurrentY = iLine * 300
Printer.Print "出生日期"
Printer.CurrentX = PLeft(51)
Printer.CurrentY = iLine * 300
Printer.Print "子女姓名"
Printer.CurrentX = PLeft(52)
Printer.CurrentY = iLine * 300
Printer.Print "性別"
Printer.CurrentX = PLeft(53)
Printer.CurrentY = iLine * 300
Printer.Print "出生日期"
iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print String(300, "-")
iLine = iLine + 1
End Sub

Sub PrintDetail()
Dim m_j As Integer
For m_j = 1 To 14
    Printer.CurrentX = PLeft(m_j)
    Printer.CurrentY = iLine * 300
    Printer.Print strTemp(m_j)
Next m_j
iLine = iLine + 1
For m_j = 15 To 23
    Printer.CurrentX = PLeft(m_j)
    Printer.CurrentY = iLine * 300
    Printer.Print strTemp(m_j)
Next m_j
iLine = iLine + 1
For m_j = 24 To 28
    Printer.CurrentX = PLeft(m_j)
    Printer.CurrentY = iLine * 300
    Printer.Print strTemp(m_j)
Next m_j
iLine = iLine + 1
For m_j = 29 To 34
    Printer.CurrentX = PLeft(m_j)
    Printer.CurrentY = iLine * 300
    Printer.Print strTemp(m_j)
Next m_j
iLine = iLine + 1
For m_j = 35 To 39
    Printer.CurrentX = PLeft(m_j)
    Printer.CurrentY = iLine * 300
    Printer.Print strTemp(m_j)
Next m_j
iLine = iLine + 1
For m_j = 40 To 53
    Printer.CurrentX = PLeft(m_j)
    Printer.CurrentY = iLine * 300
    Printer.Print strTemp(m_j)
Next m_j
iLine = iLine + 1
End Sub

'名冊
Sub StrMenu2()
Printer.Orientation = 2
'Printer.FontName = "標楷體"
m_str = "select st.*,nvl(A0922,'(舊)'||A0902) A0902,a1.ac03 a1ac03,a2.ac03 a2ac03,a3.ac03 a3ac03,a4.ac03 a4ac03" & _
         " from staff st,acc090 a09,acc090NEW,allcode A1,allcode A2,allcode A3,allcode A4,SalaryData" & _
         " where ST01=SD01 and (SD02 not in('P','F') or SD02 is null) and st03=a0901(+) and st93=a0921(+)" & _
         " and '01'=a1.ac01(+) and st20=a1.ac02(+) and '02'=a2.ac01(+) and st21=a2.ac02(+)" & _
         " and '06'=a3.ac01(+) and st27=a3.ac02(+) and '03'=a4.ac01(+) and st37=a4.ac02(+)" & _
         m_StrSQL & "order by nvl(st.st93,st.st03),st01 "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        .MoveFirst
        
        '預設值
        iLine = 1
        strType = "" '切頁條件
        
        Do While Not .EOF
            For m_i = 1 To 15
                strTemp(m_i) = ""
            Next m_i
            'strTemp(16) = CheckStr(.Fields("a0902"))
            strTemp(1) = CheckStr(.Fields("st01"))
            strTemp(2) = CheckStr(.Fields("st02"))
            strTemp(3) = CheckStr(.Fields("a1ac03"))
            strTemp(4) = IIf(CheckStr(.Fields("st06")) = "1", "北所", IIf(CheckStr(.Fields("st06")) = "2", "中所", IIf(CheckStr(.Fields("st06")) = "3", "南所", IIf(CheckStr(.Fields("st06")) = "4", "高所", "其他"))))
            strTemp(5) = IIf(CheckStr(.Fields("st22")) = "M", "男", IIf(CheckStr(.Fields("st22")) = "F", "女", ""))
            strTemp(6) = CheckStr(.Fields("a3ac03"))
            strTemp(7) = CheckStr(.Fields("st26"))
            strTemp(8) = ChangeTStringToTDateString(TAIWANDATE(CheckStr(.Fields("st23"))))
            strTemp(9) = CheckStr(.Fields("st25"))
            strTemp(10) = ChangeTStringToTDateString(TAIWANDATE(CheckStr(.Fields("st13"))))
            strTemp(11) = CheckStr(.Fields("st09"))
            strTemp(12) = CheckStr(.Fields("st19"))
            '第二行
            strTemp(13) = CheckStr(.Fields("st08"))
            strTemp(14) = CheckStr(.Fields("st38"))
            strTemp(15) = CheckStr(.Fields("st39"))
                        
            If iLine > 37 Or iLine = 1 Or _
               strType <> CheckStr(.Fields("a0902")) Then
               If (strType <> "" And strType <> CheckStr(.Fields("a0902"))) Then
                  '小計
               End If
               'If .AbsolutePosition <> .RecordCount Then
                  If strType <> "" Then Printer.NewPage
                  iLine = 1
                  PrintTitle2 CheckStr(.Fields("a0902")) '列印表頭
               'End If
            End If
            
            PrintDetail2 '列印表中
            
            strType = CheckStr(.Fields("a0902"))
            .MoveNext
        Loop
    End With
Else
   MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
   Exit Sub
End If
Printer.EndDoc
ShowPrintOk
End Sub

Sub GetPleft2()
PLeft(1) = 200
PLeft(2) = 900
PLeft(3) = 2100
PLeft(4) = 4000 - 300
PLeft(5) = 5000 - 300
PLeft(6) = 6000 - 300
PLeft(7) = 7500 - 300
PLeft(8) = 9050 - 300
PLeft(9) = 10050 - 300
PLeft(10) = 11050 - 300
PLeft(11) = 12300 - 400
PLeft(12) = 13300 - 100
PLeft(13) = 2100
PLeft(14) = 10050 - 300
PLeft(15) = 13300 - 100
End Sub

Sub PrintDetail2()
Dim m_j As Integer
For m_j = 1 To 12
    Printer.CurrentX = PLeft(m_j)
    Printer.CurrentY = iLine * 300
    Printer.Print strTemp(m_j)
Next m_j
iLine = iLine + 1
For m_j = 13 To 15
    Printer.CurrentX = PLeft(m_j)
    Printer.CurrentY = iLine * 300
    Printer.Print strTemp(m_j)
Next m_j
iLine = iLine + 1
End Sub

Sub PrintTitle2(oStr As String)
GetPleft2

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False
'Modify By Sindy 2025/9/15
strExc(10) = IIf(Combo2.Text = "", "台一關係企業", Trim(Mid(Combo2.Text, 3))) & "　員工名冊"
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strExc(10)) / 2)
Printer.CurrentY = 300
Printer.Print strExc(10)
'2025/9/15 END
Printer.Font.Size = 10
Printer.Font.Underline = False
Printer.FontBold = False
'Modify By Sindy 2025/9/15
Printer.CurrentX = PLeft(1)
Printer.CurrentY = 600
Printer.Print "部　　門：" & oStr
'2025/9/15 END
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 600
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
'Modify By Sindy 2025/9/15
If Trim(Me.txt1(5).Text) <> "" Then
   strExc(10) = IIf(txt1(5) = "1", "北所", IIf(txt1(5) = "2", "中所", IIf(txt1(5) = "3", "南所", IIf(txt1(5) = "4", "高所", "其他"))))
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = 900
   Printer.Print "所　　別：" & strExc(10)
End If
'2025/9/15 END
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "頁　　次：" & Printer.Page
iLine = 5
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "員工"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "姓　名"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "職稱"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "所在"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iLine * 300
Printer.Print "性別"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iLine * 300
Printer.Print "出生地"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iLine * 300
Printer.Print "身分證字號"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iLine * 300
Printer.Print "出生日期"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iLine * 300
Printer.Print "血型"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iLine * 300
Printer.Print "入所日期"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iLine * 300
Printer.Print "聯絡電話"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iLine * 300
Printer.Print "行動電話"
iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "編號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "通訊地址"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "所別"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iLine * 300
Printer.Print "畢業學校"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iLine * 300
Printer.Print "科　　系"
iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print String(300, "-")
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
   Set frm160101 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 4, 5 'Modify By Sindy 2025/9/15 +,5
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 0, 1, 2, 3
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
         If Index = 0 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 1 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case 2, 3
         ' 檢查員工編號規則
         If txt1(Index).Text <> "" Then
            If ChkStaffID(txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         If Index = 2 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 3 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case 4
        If txt1(4) <> "" Then
            Select Case txt1(4)
            Case "1", "2"
            Case Else
                MsgBox "報表類別只可以輸入 1 或 2！", vbInformation, "輸入錯誤！"
                Cancel = True
            End Select
        End If
      Case Else
   End Select
End Sub
