VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040327 
   BorderStyle     =   1  '單線固定
   Caption         =   "作業失誤明細表"
   ClientHeight    =   2820
   ClientLeft      =   3165
   ClientTop       =   1200
   ClientWidth     =   3570
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   3570
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   17
      Left            =   990
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1920
      Width           =   345
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2565
      TabIndex        =   7
      Top             =   12
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1785
      TabIndex        =   6
      Top             =   12
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   11
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "Y"
      Top             =   1545
      Width           =   330
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   960
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1188
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2280
      MaxLength       =   7
      TabIndex        =   2
      Top             =   852
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   960
      MaxLength       =   7
      TabIndex        =   1
      Top             =   852
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   516
      Width           =   2430
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "PS：需專業部輸入作業失誤資料"
      Height          =   180
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   2520
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "列印順序："
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   1950
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(1.失誤人員 2.失誤日期)"
      Height          =   180
      Left            =   1470
      TabIndex        =   14
      Top             =   1935
      Width           =   1875
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "(Y:印)"
      Height          =   180
      Left            =   1890
      TabIndex        =   13
      Top             =   1545
      Width           =   465
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "是否列印明細："
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   1260
   End
   Begin MSForms.Label lblStaffName 
      Height          =   255
      Left            =   1905
      TabIndex        =   11
      Top             =   1215
      Width           =   1335
      VariousPropertyBits=   27
      Size            =   "2355;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "失誤人員："
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   1230
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   1920
      X2              =   2160
      Y1              =   972
      Y2              =   972
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "失誤日期："
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   900
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   570
      Width           =   900
   End
End
Attribute VB_Name = "frm040327"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/12 改成Form2.0 (lblStaffName,Printer列印未改)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim strSql As String, strTemp1 As Variant, strTemp2 As Variant, StrTest1 As String, StrTest2 As String, i As Integer, j As Integer, s As Integer
Dim PLeft(0 To 11) As Integer, k As Integer, TmpArea As String, iLine As Integer, Page As Integer
Dim strTemp3(0 To 11) As String, iPrint As Integer
Dim StrTest3 As String, Day1 As String, Day2 As String, StrTemp4 As String
Dim St As String, iK As Integer, dblKAmt As Double, iTotal As Integer, dblTotalAmt As Double
Dim iKK As Integer '業務區合計
Dim StrTest4 As String, StrTest5 As String
'Added by Lydia 2015/10/26 +處置金額
Dim dblMD06 As Double, dblTotalAmt2 As Double
Dim intAdd As Integer
Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0 '確定
    If Len(txt1(0)) = 0 Then
        s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
        txt1(0).SetFocus
        Exit Sub
    End If
    If Len(txt1(1)) = 0 Then
        s = MsgBox("失誤起日不可空白!!", , "USER 輸入錯誤")
        txt1(1).SetFocus
        txt1_GotFocus (1)
        Exit Sub
    Else
        If CheckIsTaiwanDate(Me.txt1(1).Text, True) = False Then
            txt1(1).SetFocus
            txt1_GotFocus (1)
            Exit Sub
        End If
    End If
    If Len(txt1(2)) = 0 Then
        s = MsgBox("失誤迄日不可空白!!", , "USER 輸入錯誤")
        txt1(2).SetFocus
        txt1_GotFocus (2)
        Exit Sub
    Else
        If CheckIsTaiwanDate(Me.txt1(2).Text, True) = False Then
            txt1(2).SetFocus
            txt1_GotFocus (2)
            Exit Sub
        End If
    End If
    If Val(Me.txt1(1).Text) > Val(Me.txt1(2).Text) Then
        s = MsgBox("失誤日期區間輸入錯誤!!", , "USER 輸入錯誤")
        txt1(1).SetFocus
        txt1_GotFocus (1)
        Exit Sub
    End If
    '若未輸入列印順序
    If Me.txt1(17).Text = "" Then
        s = MsgBox("請輸入列印順序", , "USER 輸入錯誤")
        txt1(17).SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/21 清除查詢印表記錄檔欄位
    StrMenu
    Me.Enabled = True
    Screen.MousePointer = vbDefault
Case 1 '結束
     Unload Me
Case Else
End Select
End Sub

Sub StrMenu()
Screen.MousePointer = vbHourglass
StrTest1 = "": StrTest2 = "": StrTest3 = "": StrTest4 = "": StrTest5 = ""
If Len(txt1(1)) <> 0 Then
    StrTest1 = StrTest1 + " AND MD02>=" & ChangeTStringToWString(txt1(1)) & " "
End If
If Len(txt1(2)) <> 0 Then
    StrTest1 = StrTest1 + " AND MD02<=" & ChangeTStringToWString(txt1(2)) & " "
End If
If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label2 & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/21
End If
If Len(txt1(3)) <> 0 Then
    StrTest1 = StrTest1 + " AND MD03='" & txt1(3) & "' "
    pub_QL05 = pub_QL05 & ";" & Label3 & txt1(3) & lblStaffName 'Add By Sindy 2010/12/21
End If
StrTest2 = StrTest1: StrTest3 = StrTest1: StrTest4 = StrTest1: StrTest5 = StrTest1
If Len(txt1(0)) <> 0 Then
   StrTest1 = StrTest1 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") "
   StrTest2 = StrTest2 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
   StrTest3 = StrTest3 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 3) & ") "
   StrTest4 = StrTest4 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 4) & ") "
   StrTest5 = StrTest5 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) 'Add By Sindy 2010/12/21
End If
'列印明細
'Modified by Lydia 2015/10/26 +處置金額MD06
If Me.txt1(11).Text = "Y" Then
    pub_QL05 = pub_QL05 & ";" & Label10 & txt1(11) 'Add By Sindy 2010/12/21
    '列印順序--失誤人員
    If Me.txt1(17).Text = "1" Then
        pub_QL05 = pub_QL05 & ";" & Label6 & "1.失誤人員" 'Add By Sindy 2010/12/21
        'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
        strSql = "Select MD03, MD02, Decode(PA57,'Y','＊','')||CP01||'-'||CP02||'-'||CP03||'-'||CP04, CP05, Decode(PA09,'000',CPM03,CPM04), MD04,  MD06, MD01, MD05, ST06, CP01, CP02, CP03, CP04 From MissData, CaseProgress, Patent, CasePropertyMap, Staff Where MD01=CP09 And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 And CP01=CPM01 And CP10=CPM02 And MD03=ST01 " & StrTest1
        strSql = strSql & " Union Select MD03, MD02, Decode(TM29,'Y','＊','')||CP01||'-'||CP02||'-'||CP03||'-'||CP04, CP05, Decode(TM10,'000',CPM03,CPM04), MD04,  MD06, MD01, MD05, ST06, CP01, CP02, CP03, CP04 From MissData, CaseProgress, Trademark, CasePropertyMap, Staff Where MD01=CP09 And CP01=TM01 And CP02=TM02 And CP03=TM03 And CP04=TM04 And CP01=CPM01 And CP10=CPM02 And MD03=ST01 " & StrTest2
        strSql = strSql & " Union Select MD03, MD02, Decode(LC08,'Y','＊','')||CP01||'-'||CP02||'-'||CP03||'-'||CP04, CP05, Decode(LC15,'000',CPM03,CPM04), MD04,  MD06, MD01, MD05, ST06, CP01, CP02, CP03, CP04 From MissData, CaseProgress, Lawcase, CasePropertyMap, Staff Where MD01=CP09 And CP01=LC01 And CP02=LC02 And CP03=LC03 And CP04=LC04 And CP01=CPM01 And CP10=CPM02 And MD03=ST01 " & StrTest3
        strSql = strSql & " Union Select MD03, MD02, Decode(HC09,'Y','＊','')||CP01||'-'||CP02||'-'||CP03||'-'||CP04, CP05, Decode('000','000',CPM03,CPM04), MD04,  MD06, MD01, MD05, ST06, CP01, CP02, CP03, CP04 From MissData, CaseProgress, Hirecase, CasePropertyMap, Staff Where MD01=CP09 And CP01=HC01 And CP02=HC02 And CP03=HC03 And CP04=HC04 And CP01=CPM01 And CP10=CPM02 And MD03=ST01 " & StrTest4
        strSql = strSql & " Union Select MD03, MD02, Decode(SP15,'Y','＊','')||CP01||'-'||CP02||'-'||CP03||'-'||CP04, CP05, Decode(SP09,'000',CPM03,CPM04), MD04,  MD06, MD01, MD05, ST06, CP01, CP02, CP03, CP04 From MissData, CaseProgress, Servicepractice, CasePropertyMap, Staff Where MD01=CP09 And CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04 And CP01=CPM01 And CP10=CPM02 And MD03=ST01 " & StrTest5
        'end 2018/06/05
        strSql = strSql & " Order By ST06, MD03, MD02, CP01, CP02, CP03, CP04, CP05, MD01 "
    '列印順序--失誤日期
    Else
        pub_QL05 = pub_QL05 & ";" & Label6 & "2.失誤日期" 'Add By Sindy 2010/12/21
        'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
        strSql = "Select MD02, MD03, Decode(PA57,'Y','＊','')||CP01||'-'||CP02||'-'||CP03||'-'||CP04, CP05, Decode(PA09,'000',CPM03,CPM04), MD04,  MD06, MD01, MD05, ST06, CP01, CP02, CP03, CP04 From MissData, CaseProgress, Patent, CasePropertyMap, Staff Where MD01=CP09 And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 And CP01=CPM01 And CP10=CPM02 And MD03=ST01 " & StrTest1
        strSql = strSql & " Union Select MD02, MD03, Decode(TM29,'Y','＊','')||CP01||'-'||CP02||'-'||CP03||'-'||CP04, CP05, Decode(TM10,'000',CPM03,CPM04), MD04,  MD06, MD01, MD05, ST06, CP01, CP02, CP03, CP04 From MissData, CaseProgress, Trademark, CasePropertyMap, Staff Where MD01=CP09 And CP01=TM01 And CP02=TM02 And CP03=TM03 And CP04=TM04 And CP01=CPM01 And CP10=CPM02 And MD03=ST01 " & StrTest2
        strSql = strSql & " Union Select MD02, MD03, Decode(LC08,'Y','＊','')||CP01||'-'||CP02||'-'||CP03||'-'||CP04, CP05, Decode(LC15,'000',CPM03,CPM04), MD04,  MD06, MD01, MD05, ST06, CP01, CP02, CP03, CP04 From MissData, CaseProgress, Lawcase, CasePropertyMap, Staff Where MD01=CP09 And CP01=LC01 And CP02=LC02 And CP03=LC03 And CP04=LC04 And CP01=CPM01 And CP10=CPM02 And MD03=ST01 " & StrTest3
        strSql = strSql & " Union Select MD02, MD03, Decode(HC09,'Y','＊','')||CP01||'-'||CP02||'-'||CP03||'-'||CP04, CP05, Decode('000','000',CPM03,CPM04), MD04,  MD06, MD01, MD05, ST06, CP01, CP02, CP03, CP04 From MissData, CaseProgress, Hirecase, CasePropertyMap, Staff Where MD01=CP09 And CP01=HC01 And CP02=HC02 And CP03=HC03 And CP04=HC04 And CP01=CPM01 And CP10=CPM02 And MD03=ST01 " & StrTest4
        strSql = strSql & " Union Select MD02, MD03, Decode(SP15,'Y','＊','')||CP01||'-'||CP02||'-'||CP03||'-'||CP04, CP05, Decode(SP09,'000',CPM03,CPM04), MD04,  MD06, MD01, MD05, ST06, CP01, CP02, CP03, CP04 From MissData, CaseProgress, Servicepractice, CasePropertyMap, Staff Where MD01=CP09 And CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04 And CP01=CPM01 And CP10=CPM02 And MD03=ST01 " & StrTest5
        'end 2018/06/05
        strSql = strSql & " Order By MD02, ST06, MD03, CP01, CP02, CP03, CP04, CP05, MD01 "
    End If
'不列印明細
Else
    '列印順序--失誤人員
    If Me.txt1(17).Text = "1" Then
        pub_QL05 = pub_QL05 & ";" & Label6 & "1.失誤人員" 'Add By Sindy 2010/12/21
        strSql = "Select MD03, Count(MD01), Sum(Nvl(MD04,0)), Sum(Nvl(MD06,0)), ST06 From MissData, CaseProgress, Patent, CasePropertyMap, Staff Where MD01=CP09 And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 And CP01=CPM01 And CP10=CPM02 And MD03=ST01 " & StrTest1 & " Group By MD03, ST06 "
        strSql = strSql & " Union Select MD03, Count(MD01), Sum(Nvl(MD04,0)), Sum(Nvl(MD06,0)), ST06 From MissData, CaseProgress, Trademark, CasePropertyMap, Staff Where MD01=CP09 And CP01=TM01 And CP02=TM02 And CP03=TM03 And CP04=TM04 And CP01=CPM01 And CP10=CPM02 And MD03=ST01 " & StrTest2 & " Group By MD03, ST06 "
        strSql = strSql & " Union Select MD03, Count(MD01), Sum(Nvl(MD04,0)), Sum(Nvl(MD06,0)), ST06 From MissData, CaseProgress, Lawcase, CasePropertyMap, Staff Where MD01=CP09 And CP01=LC01 And CP02=LC02 And CP03=LC03 And CP04=LC04 And CP01=CPM01 And CP10=CPM02 And MD03=ST01 " & StrTest3 & " Group By MD03, ST06 "
        strSql = strSql & " Union Select MD03, Count(MD01), Sum(Nvl(MD04,0)), Sum(Nvl(MD06,0)), ST06 From MissData, CaseProgress, Hirecase, CasePropertyMap, Staff Where MD01=CP09 And CP01=HC01 And CP02=HC02 And CP03=HC03 And CP04=HC04 And CP01=CPM01 And CP10=CPM02 And MD03=ST01 " & StrTest4 & " Group By MD03, ST06 "
        strSql = strSql & " Union Select MD03, Count(MD01), Sum(Nvl(MD04,0)), Sum(Nvl(MD06,0)), ST06 From MissData, CaseProgress, Servicepractice, CasePropertyMap, Staff Where MD01=CP09 And CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04 And CP01=CPM01 And CP10=CPM02 And MD03=ST01 " & StrTest5 & " Group By MD03, ST06 "
        strSql = strSql & " Order By ST06, MD03 "
    '列印順序--失誤日期
    Else
        pub_QL05 = pub_QL05 & ";" & Label6 & "2.失誤日期" 'Add By Sindy 2010/12/21
        strSql = "Select MD02, Count(MD01), Sum(Nvl(MD04,0)), Sum(Nvl(MD06,0)) From MissData, CaseProgress, Patent, CasePropertyMap, Staff Where MD01=CP09 And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 And CP01=CPM01 And CP10=CPM02 And MD03=ST01 " & StrTest1 & " Group By MD02 "
        strSql = strSql & " Union Select MD02, Count(MD01), Sum(Nvl(MD04,0)), Sum(Nvl(MD06,0)) From MissData, CaseProgress, Trademark, CasePropertyMap, Staff Where MD01=CP09 And CP01=TM01 And CP02=TM02 And CP03=TM03 And CP04=TM04 And CP01=CPM01 And CP10=CPM02 And MD03=ST01 " & StrTest2 & " Group By MD02 "
        strSql = strSql & " Union Select MD02, Count(MD01), Sum(Nvl(MD04,0)), Sum(Nvl(MD06,0)) From MissData, CaseProgress, Lawcase, CasePropertyMap, Staff Where MD01=CP09 And CP01=LC01 And CP02=LC02 And CP03=LC03 And CP04=LC04 And CP01=CPM01 And CP10=CPM02 And MD03=ST01 " & StrTest3 & " Group By MD02 "
        strSql = strSql & " Union Select MD02, Count(MD01), Sum(Nvl(MD04,0)), Sum(Nvl(MD06,0)) From MissData, CaseProgress, Hirecase, CasePropertyMap, Staff Where MD01=CP09 And CP01=HC01 And CP02=HC02 And CP03=HC03 And CP04=HC04 And CP01=CPM01 And CP10=CPM02 And MD03=ST01 " & StrTest4 & " Group By MD02 "
        strSql = strSql & " Union Select MD02, Count(MD01), Sum(Nvl(MD04,0)), Sum(Nvl(MD06,0)) From MissData, CaseProgress, Servicepractice, CasePropertyMap, Staff Where MD01=CP09 And CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04 And CP01=CPM01 And CP10=CPM02 And MD03=ST01 " & StrTest5 & " Group By MD02 "
        strSql = strSql & " Order By MD02 "
    End If
End If
'end 2015/10/26
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/12/21
    adoRecordset.MoveFirst
    Do While adoRecordset.EOF = False
        adoRecordset.MoveNext
    Loop
    
    'Added by Lydia 2015/10/26
    GetPrintLeft
    '設定紙張-橫印
    Printer.PaperSize = 9
    Printer.Orientation = 2
    
    '若選擇列印明細
    If Me.txt1(11).Text = "Y" Then
        StrPrintDoc
    '若選擇不列印明細
    Else
        StrPrintDocTotal
    End If
    CheckOC
Else
    InsertQueryLog (0)  'Add By Sindy 2010/12/21
    ShowNoData
    CheckOC
    Screen.MousePointer = vbDefault
    Exit Sub
End If
Screen.MousePointer = vbDefault
End Sub

'明細
Sub StrPrintDoc()

'Remove by Lydia 2015/10/26
'GetPrintLeft
iLine = 1
Page = 1
StrPrintTitle str(Page)
iPrint = 2700
iTotal = 0: dblTotalAmt = 0 '總計
iK = 0: dblKAmt = 0 '小計
'Added by Lydia 2015/10/26
dblMD06 = 0: dblTotalAmt2 = 0
With adoRecordset
    .MoveFirst
    Do While .EOF = False
        'Modified by Lydia 2015/10/26
        'For j = 0 To 8
        For j = 0 To 9
            If Not IsNull(.Fields(j)) Then
                strTemp3(j) = .Fields(j)
            Else
                strTemp3(j) = ""
            End If
        Next j
        St = strTemp3(0)
        iK = iK + 1
        dblKAmt = dblKAmt + Val(strTemp3(5))
        iTotal = iTotal + 1
        dblTotalAmt = dblTotalAmt + Val(strTemp3(5))
        'Added by Lydia 2015/10/26
        dblMD06 = dblMD06 + Val(strTemp3(6))
        dblTotalAmt2 = dblTotalAmt2 + Val(strTemp3(6))
        '失誤人員或失誤日期
        Printer.CurrentX = PLeft(0)
        Printer.CurrentY = iPrint
        If iK = 1 Then
            If Me.txt1(17).Text = "1" Then
                Printer.Print StrToStr(GetStaffName(strTemp3(0), True), 4)
            Else
                Printer.Print ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(0)))
            End If
        Else
           Printer.Print ""
        End If
        Printer.CurrentX = PLeft(1)
        Printer.CurrentY = iPrint
        If Me.txt1(17).Text = "1" Then
            Printer.Print ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(1)))
        Else
            Printer.Print StrToStr(GetStaffName(strTemp3(1), True), 4)
        End If
        '本所案號
        Printer.CurrentX = PLeft(2)
        Printer.CurrentY = iPrint
        Printer.Print strTemp3(2)
        '收文日
        Printer.CurrentX = PLeft(3)
        Printer.CurrentY = iPrint
        Printer.Print ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(3)))
        ''案件性質
        Printer.CurrentX = PLeft(4)
        Printer.CurrentY = iPrint
        Printer.Print StrToStr(strTemp3(4), 6)
        '失誤金額
        Printer.CurrentX = PLeft(5) + Printer.TextWidth("失誤金額") - Printer.TextWidth(Format(strTemp3(5), "#,##0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(strTemp3(5), "#,##0")
        'Modified by Lydia 2015/10/26
        intAdd = 6
        '處置金額 (Add by Lydia 2015/10/26)
        Printer.CurrentX = PLeft(intAdd) + Printer.TextWidth("處置金額") - Printer.TextWidth(Format(strTemp3(intAdd), "#,##0"))
        Printer.CurrentY = iPrint
        Printer.Print strTemp3(intAdd)
        intAdd = intAdd + 1
        '總收文號
        'Printer.CurrentX = PLeft(6)
        'Printer.CurrentY = iPrint
        'Printer.Print strTemp3(6)
        Printer.CurrentX = PLeft(intAdd)
        Printer.CurrentY = iPrint
        Printer.Print strTemp3(intAdd)
        intAdd = intAdd + 1
        '備註
'        Printer.CurrentX = PLeft(7)
'        Printer.CurrentY = iPrint
'        Printer.Print strTemp3(7)
        Printer.CurrentX = PLeft(intAdd)
        Printer.CurrentY = iPrint
        Printer.Print PUB_StrToStr(strTemp3(intAdd), 46)
        'end 2015/10/26
        .MoveNext
        If .EOF = False Then
            If Not IsNull(.Fields(0)) Then
                StrTest1 = .Fields(0)
            Else
                StrTest1 = ""
            End If
        End If
        If .EOF = False Then
            If StrTest1 <> St Then
                iPrint = iPrint + 300
                Printer.CurrentX = 0
                Printer.CurrentY = iPrint
                'Modified by Lydia 2015/10/26
                'Printer.Print String(200, "-")
                SubPrintLine
                iPrint = iPrint + 300
                Printer.CurrentX = 500
                Printer.CurrentY = iPrint
                'Modified by Lydia 2015/10/26
                'Printer.Print "小計==>筆數：" & Format(Trim(str(iK)), "#,##0") & "，失誤金額：" & Format(dblKAmt, "#,##0")
                Printer.Print "小計==>筆數：" & Format(Trim(str(iK)), "#,##0") & "，失誤金額：" & Format(dblKAmt, "#,##0") & "，處置金額：" & Format(dblMD06, "#,##0")
                dblMD06 = 0
                'end 2015/10/26
                iK = 0
                dblKAmt = 0
                iPrint = iPrint + 300
                Printer.CurrentX = 0
                Printer.CurrentY = iPrint
                'Modified by Lydia 2015/10/26
                'Printer.Print String(200, "-")
                SubPrintLine
                iLine = iLine + 3
                St = StrTest1
            End If
            'Modified by Lydia 2015/10/26
            'If iPrint >= 14000 Then
            If iPrint >= 11000 Then
                iPrint = iPrint + 300
                Printer.NewPage
                Page = Page + 1
                StrPrintTitle str(Page)
                iPrint = 2400
                iLine = 0
            End If
            iLine = iLine + 1
            iPrint = iPrint + 300
        End If
    Loop
End With
'小計
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
'Modified by Lydia 2015/10/26
'Printer.Print String(200, "-")
SubPrintLine
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
'Modified by Lydia 2015/10/26
'Printer.Print "小計==>筆數：" & Format(Trim(str(iK)), "#,##0") & "，失誤金額：" & Format(dblKAmt, "#,##0")
Printer.Print "小計==>筆數：" & Format(Trim(str(iK)), "#,##0") & "，失誤金額：" & Format(dblKAmt, "#,##0") & "，處置金額：" & Format(dblMD06, "#,##0")
'合計
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
'Modified by Lydia 2015/10/26
'Printer.Print String(200, "-")
SubPrintLine
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
'Modified by Lydia 2015/10/26
'Printer.Print "合計==>筆數：" & Format(Trim(str(iTotal)), "#,##0") & "，失誤金額：" & Format(dblTotalAmt, "#,##0")
Printer.Print "合計==>筆數：" & Format(Trim(str(iTotal)), "#,##0") & "，失誤金額：" & Format(dblTotalAmt, "#,##0") & "，處置金額：" & Format(dblTotalAmt2, "#,##0")
Printer.EndDoc
ShowPrintOk
CheckOC

End Sub

'Add By Cheng 2003/01/09
Sub StrPrintDocTotal()
Dim strTempName As String '代理人名稱
'Remove by Lydia 2015/10/26
'GetPrintLeft
iLine = 1
Page = 1
StrPrintTitle str(Page)
iPrint = 2700
iTotal = 0       ' 總數
dblTotalAmt = 0
dblTotalAmt2 = 0 'Added by Lydia 2015/10/26

With adoRecordset
    .MoveFirst
    Do While .EOF = False
        'Modified by Lydia 2015/10/26
        'For j = 0 To 2
        For j = 0 To 3
            If Not IsNull(.Fields(j)) Then
                strTemp3(j) = .Fields(j)
            Else
                strTemp3(j) = ""
            End If
        Next j
        '失誤人員或日期
        Printer.CurrentX = PLeft(0)
        Printer.CurrentY = iPrint
        If Me.txt1(17).Text = "1" Then
            Printer.Print GetStaffName(strTemp3(0), True)
        Else
            Printer.Print ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(0)))
        End If
        '失誤筆數
        Printer.CurrentX = PLeft(2) + Printer.TextWidth("本所案號") - Printer.TextWidth(Format(strTemp3(1), "#,##0") & " 筆")
        Printer.CurrentY = iPrint
        Printer.Print Format(strTemp3(1), "#,##0") & " 筆"
        iTotal = iTotal + Val(strTemp3(1))
        '失誤金額
        Printer.CurrentX = PLeft(5) + Printer.TextWidth("失誤金額") - Printer.TextWidth(Format(strTemp3(2), "#,##0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(strTemp3(2), "#,##0")
        
        'Added by Lydia 2015/10/26
        Printer.CurrentX = PLeft(6) + Printer.TextWidth("處置金額") - Printer.TextWidth(Format(strTemp3(3), "#,##0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(strTemp3(3), "#,##0")
        dblTotalAmt = dblTotalAmt + Val(strTemp3(2))
        dblTotalAmt2 = dblTotalAmt2 + Val(strTemp3(3)) 'Added by Lydia 2015/10/26
        
        .MoveNext
        If iPrint >= 14000 Then
            iPrint = iPrint + 300
            Printer.NewPage
            Page = Page + 1
            StrPrintTitle str(Page)
            iPrint = 2400
            iLine = 0
        End If
        iLine = iLine + 1
        iPrint = iPrint + 300
    Loop
End With
'合計
Printer.CurrentX = 0
Printer.CurrentY = iPrint
'Modified by Lydia 2015/10/26
'Printer.Print String(200, "-")
SubPrintLine
iPrint = iPrint + 300
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
'Modified by Lydia 2015/10/26
'Printer.Print "合計：共 " & Format(iTotal, "#,##0") & " 筆，失誤金額：" & Format(dblTotalAmt, "#,##0")
Printer.Print "合計==>筆數：" & Format(iTotal, "#,##0") & " 筆，失誤金額：" & Format(dblTotalAmt, "#,##0") & "，處置金額：" & Format(dblTotalAmt2, "#,##0")

Printer.EndDoc
ShowPrintOk
CheckOC
End Sub

Sub StrPrintTitle(ByRef Page As String)
Dim intMid As Long 'Added by Lydia 2015/10/26
'Remove by Lydia 2015/10/26
'GetPrintLeft
k = 500
'Remove by Lydia 2015/10/26
'Printer.Orientation = 1

Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
'Modified by Lydia 2015/10/26
'Printer.CurrentX = 4000
intMid = Printer.ScaleWidth / 2 - (Printer.TextWidth("作業失誤明細表") / 2)
Printer.CurrentX = intMid
Printer.CurrentY = k
Printer.Print "作業失誤明細表"
Printer.Font.Underline = False
Printer.Font.Bold = False
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.CurrentY = k + 800
Printer.Print "列印人：" & strUserName
'Modified by Lydia 2015/10/26
'Printer.CurrentX = 4000
Printer.CurrentX = intMid
Printer.CurrentY = k + 800
Printer.Print "失誤日期：" & Me.txt1(1).Text & "－" & Me.txt1(2).Text
'Modified by Lydia 2015/10/26
'Printer.CurrentX = 8875
intMid = Printer.ScaleWidth - Printer.TextWidth(String(20, "字"))
Printer.CurrentX = intMid
Printer.CurrentY = k + 800
Printer.Print "列印日期：" & Format(strSrvDate(2), "###/##/##")
Printer.CurrentX = 0
Printer.CurrentY = k + 1100
Printer.Print "系統類別：" & Me.txt1(0).Text
'Modified by Lydia 2015/10/26
'Printer.CurrentX = 8875
Printer.CurrentX = intMid
Printer.CurrentY = k + 1100
Printer.Print "頁　　次：" & Page
Printer.CurrentX = 0
Printer.CurrentY = k + 1400
'Modified by Lydia 2015/10/26
'Printer.Print String(200, "-")
SubPrintLine
Printer.CurrentX = PLeft(0)
Printer.CurrentY = k + 1700
Printer.Print IIf(Me.txt1(17).Text = "1", "失誤人員", "失誤日期")
Printer.CurrentX = PLeft(1)
Printer.CurrentY = k + 1700
Printer.Print IIf(Me.txt1(17).Text = "1", "失誤日期", "失誤人員")
Printer.CurrentX = PLeft(2)
Printer.CurrentY = k + 1700
Printer.Print "本所案號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = k + 1700
Printer.Print "收文日"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = k + 1700
Printer.Print "案件性質"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = k + 1700
Printer.Print "失誤金額"
'Modified by Lydia 2015/10/26
'Printer.CurrentX = PLeft(6)
'Printer.CurrentY = k + 1700
'Printer.Print "總收文號"
'Printer.CurrentX = PLeft(7)
'Printer.CurrentY = k + 1700
'Printer.Print "備註"
intAdd = 6
Printer.CurrentX = PLeft(intAdd)
Printer.CurrentY = k + 1700
Printer.Print "處置金額"
intAdd = intAdd + 1
Printer.CurrentX = PLeft(intAdd)
Printer.CurrentY = k + 1700
Printer.Print "總收文號"
intAdd = intAdd + 1
Printer.CurrentX = PLeft(intAdd)
Printer.CurrentY = k + 1700
Printer.Print "備註"
'end 2015/10/26
Printer.CurrentX = 0
Printer.CurrentY = k + 2000
'Modified by Lydia 2015/10/26
'Printer.Print String(200, "-")
SubPrintLine
End Sub

Sub GetPrintLeft()
Erase PLeft
PLeft(0) = 0 '失誤日期或失誤人員
PLeft(1) = 1125 '失誤人員或失誤日期
PLeft(2) = 2250 '本所案號
PLeft(3) = 4500 '收文日
PLeft(4) = 5750 '案件性質
PLeft(5) = 7500 '失誤金額
'Modified by Lydia 2015/10/26
'PLeft(6) = 8625 '總收文號
'PLeft(7) = 9875 '備註
intAdd = 6
PLeft(intAdd) = 8625 '處置金額
intAdd = intAdd + 1
PLeft(intAdd) = 9750 '總收文號
intAdd = intAdd + 1
PLeft(intAdd) = 11000 '備註
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   strTemp1 = Split(UCase(GetSystemKindByNick), ",")
   For i = 0 To UBound(strTemp1)
       txt1(0) = txt1(0) + strTemp1(i) + ","
   Next i
   'add by sonia 2016/11/15
   If PUB_GetST05(strUserNum) = "00" Then
      Label5.Visible = True
   Else
      Label5.Visible = False
   End If
   'end 2016/11/15
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm040327 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
Select Case Index
Case 11 '是否列印明細
    If KeyAscii <> 89 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
Case 17 '列印順序
    If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0 '系統類別
     strTemp1 = Split(UCase(GetSystemKindByNick), ",")
     strTemp2 = Split(UCase(txt1(0)), ",")
     For i = 0 To UBound(strTemp2)
        s = 0
        For j = 0 To UBound(strTemp1)
            If strTemp1(j) = strTemp2(i) Then
                s = 1
                Exit For
            End If
        Next j
        If s = 0 Then
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限 ", , "權限問題")
            txt1(0).SetFocus
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            Exit Sub
        End If
    Next i
Case 11 '是否列印明細
     Select Case Trim(txt1(11))
     Case "y", "Y", ""
     Case Else
         s = MsgBox("是否列印明細只能輸入 Y 或空白 !!", , "USER 輸入錯誤")
         txt1(11).SetFocus
         txt1(11).SelStart = 0
         txt1(11).SelLength = Len(txt1(11))
         Exit Sub
     End Select
Case 3 '失誤人員
     lblStaffName = GetPrjSalesNM(txt1(Index))
     If Len(txt1(Index)) <> 0 Then
        If Len(lblStaffName.Caption) = 0 Then
            s = MsgBox("失誤人員輸入錯誤！", , "錯誤！")
            txt1(Index).SetFocus
            txt1_GotFocus (Index)
            Exit Sub
        End If
     End If
Case 2 '失誤日期
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
End Select

End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
    Case 1, 2 '失誤日期
        If Me.txt1(Index).Text <> "" Then
            If CheckIsTaiwanDate(Me.txt1(Index).Text, True) = False Then
                Cancel = True
            End If
        End If
    End Select
    If Cancel = True Then TextInverse Me.txt1(Index)
End Sub
'Added by Lydia 2015/10/26
Private Sub SubPrintLine()
  Printer.Print String(138, "-")
End Sub
