VERSION 5.00
Begin VB.Form frm050321 
   BorderStyle     =   1  '單線固定
   Caption         =   "新案件承辦人明細表"
   ClientHeight    =   2190
   ClientLeft      =   3120
   ClientTop       =   1530
   ClientWidth     =   3135
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3135
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   930
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1380
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   924
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1710
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2244
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1080
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   924
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1080
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   924
      MaxLength       =   7
      TabIndex        =   1
      Top             =   780
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   924
      TabIndex        =   0
      Top             =   480
      Width           =   2136
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2235
      TabIndex        =   7
      Top             =   24
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1440
      TabIndex        =   6
      Top             =   24
      Width           =   756
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "業務區："
      Height          =   180
      Left            =   90
      TabIndex        =   14
      Top             =   1410
      Width           =   720
   End
   Begin VB.Label lblSaleZone 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1770
      TabIndex        =   13
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Line Line2 
      X1              =   1884
      X2              =   2124
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1770
      TabIndex        =   12
      Top             =   1770
      Width           =   1080
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   180
      Left            =   90
      TabIndex        =   11
      Top             =   1740
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Left            =   90
      TabIndex        =   10
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "收文日："
      Height          =   180
      Left            =   90
      TabIndex        =   9
      Top             =   780
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Left            =   90
      TabIndex        =   8
      Top             =   480
      Width           =   900
   End
End
Attribute VB_Name = "frm050321"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strSql As String, i As Integer, j As Integer, s As Integer, k As Integer
Dim strTemp1 As Variant, strTemp2 As Variant, StrTest As String, StrTest2 As String
Dim strSQL1 As String, strSQL2 As String, strTemp(0 To 11) As String
Dim PLeft(0 To 9) As Integer, Page As Integer, iPrint As Integer
Dim StrTemp6 As String, StrTemp4(0 To 4) As String, StrTemp5(0 To 4) As String
Dim blnClkSure As Boolean '判斷是否按下確定按鈕
Dim m_strSaleZoneCode '業務區代碼
Dim m_strSaleZone '業務區名稱
Dim m_strReceiveDate '收文日
 
Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0 '確定
    'Add By Cheng 2002/09/16
    blnClkSure = False
    If Len(txt1(0)) = 0 Then
       s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
       txt1(0).SetFocus
       txt1(0).SelStart = 0
       txt1(0).SelLength = Len(txt1(0))
       Exit Sub
    Else
       If Len(txt1(1)) = 0 Then
           s = MsgBox("收文日不可空白!!", , "USER 輸入錯誤")
           txt1(1).SetFocus: txt1(1).SelStart = 0: txt1(1).SelLength = Len(txt1(1))
           Exit Sub
       Else
           If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
              Me.txt1(1).SetFocus
              txt1_GotFocus 1
              Exit Sub
           End If
           If Me.txt1(3).Text <> "" And Me.txt1(4).Text <> "" Then
              If Me.txt1(3).Text > Me.txt1(4).Text Then
                 MsgBox "申請國家範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                 blnClkSure = True
                 Me.txt1(3).SetFocus
                 txt1_GotFocus 3
                 Exit Sub
              End If
           End If
            Me.lblSaleZone.Caption = GetDepartmentName(txt1(5).Text)
            '若有輸入業務區
            If Me.txt1(5).Text <> "" Then
                '若未帶出業務區名稱
                If Me.lblSaleZone.Caption = "" Then
                    MsgBox "業務區輸入錯誤!!!", vbExclamation + vbOKOnly
                    Me.txt1(5).SetFocus
                    txt1_GotFocus 5
                    Exit Sub
                End If
            End If
            lbl1 = GetPrjSales(txt1(8))
            '若有輸入承辦人
            If Me.txt1(8).Text <> "" Then
               If Me.txt1(8).Text = Me.lbl1.Caption Then
                  Me.lbl1.Caption = ""
                  Me.txt1(8).SetFocus
                  txt1_GotFocus 8
                  Exit Sub
               End If
            End If
            Me.Enabled = False
            Screen.MousePointer = vbHourglass
            Process
            Screen.MousePointer = vbDefault
            Me.Enabled = True
       End If
    End If
Case 1 '結束
    Unload Me
Case Else
End Select
End Sub

Sub Process()          '處理主程式
ClearQueryLog (Me.Name) 'Add By Sindy 2010/01/22 清除查詢印表記錄檔欄位
strSQL1 = ""
strSQL2 = ""
'組字串
'系統類別
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 & " and cp01 in (" & SQLGrpStr(txt1(0), 1) & ") "
   strSQL2 = strSQL2 & " and cp01 in (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) 'Add By Sindy 2010/01/22
End If
'收文日
If Len(Trim(txt1(1))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP05=" & Val(ChangeTStringToWString(txt1(1))) & " "
    strSQL2 = strSQL2 + " AND CP05=" & Val(ChangeTStringToWString(txt1(1))) & " "
    pub_QL05 = pub_QL05 & ";" & Label2 & txt1(1) 'Add By Sindy 2010/01/22
End If
'申請國家
If Len(Trim(txt1(3))) <> 0 Then
    strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)>='" & txt1(3) & "' "
    strSQL2 = strSQL2 + " AND SUBSTR(SP09,1,3)>='" & txt1(3) & "' "
End If
If Len(Trim(txt1(4))) <> 0 Then
    strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)<='" & txt1(4) & "' "
    strSQL2 = strSQL2 + " AND SUBSTR(SP09,1,3)<='" & txt1(4) & "' "
End If
If Len(Trim(txt1(3))) <> 0 Or Len(Trim(txt1(4))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label3 & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/01/22
End If
'業務區
If Len(txt1(5)) <> 0 Then
    strSQL1 = strSQL1 + " AND S1.ST03='" & txt1(5) & "' "
    strSQL2 = strSQL2 + " AND S1.ST03='" & txt1(5) & "' "
    pub_QL05 = pub_QL05 & ";" & Label8 & txt1(5) 'Add By Sindy 2010/01/22
End If
'承辦人
If Len(txt1(8)) <> 0 Then
    strSQL1 = strSQL1 + " AND CP14='" & txt1(8) & "' "
    strSQL2 = strSQL2 + " AND CP14='" & txt1(8) & "' "
    pub_QL05 = pub_QL05 & ";" & Label7 & txt1(8) 'Add By Sindy 2010/01/22
End If
strSQL1 = strSQL1 + " And CP04 ='00' "
strSQL2 = strSQL2 + " And CP04 ='00' "
'組合
'Modify By Cheng 2003/04/15
'                                     strSQL = "SELECT S1.ST03,A0902," & SQLDate("CP05") & ",CP13,S1.ST02,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(CU04,DECODE(CU05,NULL,CU06,CU05|| ' '||CU88||' '||CU89||' '||CU90)),NVL(PA05,NVL(PA06,PA07)),CP18,CP14,S2.ST02 FROM CASEPROGRESS,STAFF S1,STAFF S2,NATION,PATENT,CUSTOMER,ACC090 WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA09=NA01(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP10>='101' AND CP10<='105' AND CP09<'B' AND (CP14<>'72006' AND S2.ST03<>'P12') And SUBSTR(PA26,1,8)=CU01(+) And SUBSTR(PA26,9,1)=CU02(+) AND S1.ST03=A0901(+) " & strSQL1
'strSQL = strSQL + " UNION ALL SELECT S1.ST03,A0902," & SQLDate("CP05") & ",CP13,S1.ST02,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(CU04,DECODE(CU05,NULL,CU06,CU05|| ' '||CU88||' '||CU89||' '||CU90)),NVL(SP05,NVL(SP06,SP07)),CP18,CP14,S2.ST02 FROM CASEPROGRESS,STAFF S1,STAFF S2,NATION,SERVICEPRACTICE,CUSTOMER,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP10>='101' AND CP10<='105' AND CP09<'B' AND (CP14<>'72006' AND S2.ST03<>'P12') AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND S1.ST03=A0901(+) " & strSQL2
'Modify by Morgan 2010/8/12 百年蟲 " & SQLDate("CP05") & "-->substrb(' '||sqldatet(cp05),-9)
'modify by sonia 2016/3/3 +CP14<>'87025'
strSql = "SELECT S1.ST03,A0902,substrb(' '||sqldatet(cp05),-9),CP13,S1.ST02,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(CU04,DECODE(CU05,NULL,CU06,CU05|| ' '||CU88||' '||CU89||' '||CU90)),NVL(PA05,NVL(PA06,PA07)),CP18,CP14,S2.ST02, CP21 FROM CASEPROGRESS,STAFF S1,STAFF S2,NATION,PATENT,CUSTOMER,ACC090 WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA09=NA01(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND ((CP10>='101' AND CP10<='105') or cp10='125') AND CP09<'B' AND (CP14<>'72006' AND CP14<>'87025' AND S2.ST03<>'P12') And SUBSTR(PA26,1,8)=CU01(+) And SUBSTR(PA26,9,1)=CU02(+) AND S1.ST03=A0901(+) " & strSQL1
strSql = strSql + " UNION ALL SELECT S1.ST03,A0902,substrb(' '||sqldatet(cp05),-9),CP13,S1.ST02,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(CU04,DECODE(CU05,NULL,CU06,CU05|| ' '||CU88||' '||CU89||' '||CU90)),NVL(SP05,NVL(SP06,SP07)),CP18,CP14,S2.ST02, CP21 FROM CASEPROGRESS,STAFF S1,STAFF S2,NATION,SERVICEPRACTICE,CUSTOMER,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND ((CP10>='101' AND CP10<='105') or cp10='125') AND CP09<'B' AND (CP14<>'72006' AND CP14<>'87025' AND S2.ST03<>'P12') AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND S1.ST03=A0901(+) " & strSQL2

strSql = strSql & " ORDER BY 1, 3, 4, 6"
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/01/22
Else
   InsertQueryLog (0) 'Add By Sindy 2010/01/22
   ShowNoData
   CheckOC
   Screen.MousePointer = vbDefault
   Exit Sub
End If
'列印資料
PrintData
End Sub

Sub PrintData()
Page = 1
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    adoRecordset.MoveFirst
    m_strSaleZoneCode = "" & adoRecordset.Fields(0).Value
    m_strSaleZone = "" & adoRecordset.Fields(1).Value
    m_strReceiveDate = "" & adoRecordset.Fields(2).Value
    PrintTitle
    Do While adoRecordset.EOF = False
'        For i = 0 To 10
        For i = 0 To 11
            strTemp(i) = CheckStr(adoRecordset.Fields(i))
        Next i
        strTemp(6) = StrConv(MidB(StrConv(strTemp(6), vbFromUnicode), 1, 28), vbUnicode)
        strTemp(7) = StrConv(MidB(StrConv(strTemp(7), vbFromUnicode), 1, 28), vbUnicode)
        '若業務區不同
        If m_strSaleZoneCode <> strTemp(0) Then
            Printer.NewPage
            Page = Page + 1
            m_strSaleZoneCode = "" & adoRecordset.Fields(0).Value
            m_strSaleZone = "" & adoRecordset.Fields(1).Value
            m_strReceiveDate = "" & adoRecordset.Fields(2).Value
            PrintTitle
        End If
        If iPrint > 10000 Then
            Printer.NewPage
            Page = Page + 1
            m_strSaleZoneCode = "" & adoRecordset.Fields(0).Value
            m_strSaleZone = "" & adoRecordset.Fields(1).Value
            m_strReceiveDate = "" & adoRecordset.Fields(2).Value
            PrintTitle
        End If
        PrintDatil
        adoRecordset.MoveNext
    Loop
End If
CheckOC
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintDatil()           '印內容
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(4)
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(5)
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print strTemp(6)
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print strTemp(7)
Printer.CurrentX = PLeft(4) + 500 - Printer.TextWidth(Format(strTemp(8), "###.00"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(8), "###.00")
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print strTemp(10)
'Add By Cheng 2003/04/15
'是否多國案
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print strTemp(11)
iPrint = iPrint + 300
End Sub

Sub PrintTitle()     '印抬頭
    GetPleft
    iPrint = 500
    Printer.Orientation = 2
    Printer.Font.Name = "細明體"
    Printer.Font.Size = 22
    Printer.Font.Bold = True
    Printer.Font.Underline = True
    Printer.CurrentX = 6300
    Printer.CurrentY = iPrint
    Printer.Print "新案承辦人明細表"
    iPrint = iPrint + 500
    Printer.Font.Size = 12
    Printer.Font.Bold = False
    Printer.Font.Underline = False
    Printer.CurrentX = 500
    Printer.CurrentY = iPrint
    Printer.Print "列印人：" & strUserName
    Printer.CurrentX = 6500
    Printer.CurrentY = iPrint
    Printer.Print "收文日：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@")
    Printer.CurrentX = 13000
    Printer.CurrentY = iPrint
    Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
    iPrint = iPrint + 300
    Printer.CurrentX = 500
    Printer.CurrentY = iPrint
    Printer.Print "業務區：" & m_strSaleZoneCode & "　　" & m_strSaleZone
    Printer.CurrentX = 13000
    Printer.CurrentY = iPrint
    Printer.Print "頁　　次：" & str(Page)
    iPrint = iPrint + 300
    Printer.CurrentX = 500
    Printer.CurrentY = iPrint
    Printer.Print String(200, "-")
    iPrint = iPrint + 300
    Printer.CurrentX = PLeft(0)
    Printer.CurrentY = iPrint
    Printer.Print "智權人員"
    Printer.CurrentX = PLeft(1)
    Printer.CurrentY = iPrint
    Printer.Print "本所案號"
    Printer.CurrentX = PLeft(2)
    Printer.CurrentY = iPrint
    Printer.Print "申請人"
    Printer.CurrentX = PLeft(3)
    Printer.CurrentY = iPrint
    Printer.Print "案件名稱"
    Printer.CurrentX = PLeft(4)
    Printer.CurrentY = iPrint
    Printer.Print "點數"
    Printer.CurrentX = PLeft(5)
    Printer.CurrentY = iPrint
    Printer.Print "承辦人"
    'Add By Cheng 2003/03/2
    Printer.CurrentX = PLeft(6)
    Printer.CurrentY = iPrint
    Printer.Print "多國案"
    iPrint = iPrint + 300
    Printer.CurrentX = 500
    Printer.CurrentY = iPrint
    Printer.Print String(200, "-")
    iPrint = iPrint + 300
End Sub

Sub GetPleft()
    Erase PLeft
    PLeft(0) = 500
    PLeft(1) = 2000
    PLeft(2) = 4500
    PLeft(3) = 8500
    PLeft(4) = 13000
    PLeft(5) = 14000
    'Add By Cheng 2003/03/28
    PLeft(6) = 15250 '多國
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    txt1(0) = GetSystemKindByNick
    'Add By Cheng 2003/02/25
    '收文日預設系統日
    Me.txt1(1).Text = ServerDate - 19110000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm050321 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
    txt1(Index).SelStart = 0
    txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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
Case 4  '申請國家
   If blnClkSure = False Then
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
   Else
      blnClkSure = False
   End If
Case 5 '業務區
    Me.lblSaleZone.Caption = GetDepartmentName(txt1(5).Text)
    '若有輸入業務區
    If Me.txt1(5).Text <> "" Then
        '若未帶出業務區名稱
        If Me.lblSaleZone.Caption = "" Then
            MsgBox "業務區輸入錯誤!!!", vbExclamation + vbOKOnly
            Me.txt1(5).SetFocus
            txt1_GotFocus 5
            Exit Sub
        End If
    End If
Case 8 '承辦人
    lbl1 = GetPrjSales(txt1(Index))
    If Me.txt1(8).Text <> "" Then
        If Me.txt1(8).Text = Me.lbl1.Caption Then
            Me.lbl1.Caption = ""
            Me.txt1(8).SetFocus
            txt1_GotFocus 8
            Exit Sub
        End If
    End If
End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 1 '收文日
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
   End If
End Select
End Sub
