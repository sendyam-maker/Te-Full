VERSION 5.00
Begin VB.Form frm050314 
   BorderStyle     =   1  '單線固定
   Caption         =   "後金案件表"
   ClientHeight    =   1515
   ClientLeft      =   4110
   ClientTop       =   2715
   ClientWidth     =   3510
   ControlBox      =   0   'False
   LinkTopic       =   "Form13"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   3510
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2652
      TabIndex        =   6
      Top             =   20
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1824
      TabIndex        =   5
      Top             =   20
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2160
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1188
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   840
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1188
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2160
      MaxLength       =   7
      TabIndex        =   2
      Top             =   852
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   840
      MaxLength       =   7
      TabIndex        =   1
      Top             =   852
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   840
      MaxLength       =   1
      TabIndex        =   0
      Top             =   516
      Width           =   255
   End
   Begin VB.Line Line2 
      X1              =   1800
      X2              =   2040
      Y1              =   1308
      Y2              =   1308
   End
   Begin VB.Line Line1 
      X1              =   1800
      X2              =   2040
      Y1              =   972
      Y2              =   972
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "結果日："
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   1170
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收文日："
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   855
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(1. 收文  2. 無結果  3.有結果)"
      Height          =   180
      Left            =   1200
      TabIndex        =   8
      Top             =   516
      Width           =   2232
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "查詢別："
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   510
      Width           =   720
   End
End
Attribute VB_Name = "frm050314"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, k As Integer, s As Integer
Dim strTemp(0 To 8) As String, PLeft(0 To 8), iPrint As Integer, Page As Integer
Dim strTemp1 As Variant, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
     Select Case Val(txt1(0))
     Case 1, 2
         'Add By Cheng 2002/03/19
         If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
            Me.txt1(1).SetFocus
            txt1_GotFocus 1
            Exit Sub
         End If
         If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
            Me.txt1(2).SetFocus
            txt1_GotFocus 2
            Exit Sub
         End If
         
          If Len(txt1(2)) = 0 Then
              s = MsgBox("收文日不可空白!!", , "USER 輸入錯誤")
              txt1(1).SetFocus
              txt1_GotFocus (1)
              Exit Sub
          Else
              ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/4 清除查詢印表記錄檔欄位
              Screen.MousePointer = vbHourglass
              Me.Enabled = False
              Process
              Me.Enabled = True
              Screen.MousePointer = vbDefault
          End If
     Case 3
         'Add By Cheng 2002/03/19
         If PUB_CheckKeyInDate(Me.txt1(3)) = -1 Then
            Me.txt1(3).SetFocus
            txt1_GotFocus 3
            Exit Sub
         End If
         If PUB_CheckKeyInDate(Me.txt1(4)) = -1 Then
            Me.txt1(4).SetFocus
            txt1_GotFocus 4
            Exit Sub
         End If
         
          If Len(txt1(4)) = 0 Then
              s = MsgBox("結果日不可空白!!", , "USER 輸入錯誤")
              txt1(3).SetFocus
              txt1_GotFocus (3)
              Exit Sub
          Else
              ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/4 清除查詢印表記錄檔欄位
              Screen.MousePointer = vbHourglass
              Me.Enabled = False
              Process
              Me.Enabled = True
              Screen.MousePointer = vbDefault
          End If
     Case Else
          s = MsgBox("查詢別只能輸入 1 或 2 或 3 !!", , "USER 輸入錯誤")
          txt1(0).SetFocus
          txt1(0).SelStart = 0
          txt1(0).SelLength = Len(txt1(0))
          Exit Sub
     End Select
Case 1
     Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
End Sub

Private Sub Process()
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R050314 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
strSQL5 = ""
If Val(txt1(0)) = 3 Then '有結果
   pub_QL05 = pub_QL05 & ";" & Label1 & "有結果" 'Add By Sindy 2010/10/4
   If Len(Trim(txt1(3))) <> 0 Then
      strSQL1 = strSQL1 + " AND CP25>=" & Val(ChangeTStringToWString(txt1(3))) & " "
   End If
   If Len(Trim(txt1(4))) <> 0 Then
      strSQL1 = strSQL1 & " AND CP25<=" & Val(ChangeTStringToWString(txt1(4))) & " "
   End If
   If Len(Trim(txt1(3))) <> 0 Or Len(Trim(txt1(4))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label4 & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/10/4
    End If
   strSQL1 = strSQL1 & " AND CP19 IS NOT NULL AND CP57 IS NULL and cp24 is not null "
Else
    If Val(txt1(0)) = 2 Then '無結果
        pub_QL05 = pub_QL05 & ";" & Label1 & "無結果" 'Add By Sindy 2010/10/4
        strSQL1 = strSQL1 + " AND CP24 IS NULL "
    Else
        pub_QL05 = pub_QL05 & ";" & Label1 & "收文" 'Add By Sindy 2010/10/4
    End If
    If Len(Trim(txt1(1))) <> 0 Then
      strSQL1 = strSQL1 & " AND CP05>=" & Val(ChangeTStringToWString(txt1(1))) & " "
    End If
    If Len(Trim(txt1(2))) <> 0 Then
      strSQL1 = strSQL1 & " AND CP05<=" & Val(ChangeTStringToWString(txt1(2))) & " "
    End If
    If Len(Trim(txt1(1))) <> 0 Or Len(Trim(txt1(2))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label3 & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/10/4
    End If
    strSQL1 = strSQL1 & " AND CP19 IS NOT NULL AND CP57 IS NULL "
End If
strSQL2 = strSQL1
StrSQL3 = strSQL1
StrSQL4 = strSQL1
strSQL5 = strSQL1
strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr(GetSystemKindByNick, 1) & ") "
strSQL2 = strSQL2 & " AND CP01 IN (" & SQLGrpStr(GetSystemKindByNick, 2) & ") "
StrSQL3 = StrSQL3 & " AND CP01 IN (" & SQLGrpStr(GetSystemKindByNick, 3) & ") "
StrSQL4 = StrSQL4 & " AND CP01 IN (" & SQLGrpStr(GetSystemKindByNick, 4) & ") "
strSQL5 = strSQL5 & " AND CP01 IN (" & SQLGrpStr(GetSystemKindByNick, 5) & ") "

'Modify By Sindy 2011/2/23 增加TM78,TM79,TM80,TM81
'Modify By Sindy 2011/2/23 增加SP58,SP59,SP65,SP66
'Modify By Sindy 2011/2/23 增加LC43,LC44,LC45,LC46
'Modify By Sindy 2011/2/23 增加HC24,HC25,HC26,HC27
'商標
strSql = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),NVL(TM23,NVL(TM78,NVL(TM79,NVL(TM80,TM81)))),nvl(decode(tm10,'000',CPM03,CPM04),cp10),nvl(S1.ST02,cp14),nvl(S2.ST02,cp13),CP19,CP09,DECODE(CP24,'1','准/勝','2','駁/敗','') FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP WHERE cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND cp01=cpm01(+) AND cp10=cpm02(+) and cp14=s1.st01(+)  and cp13=s2.st01(+)  " & strSQL2
'專利
strSql = strSql + " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),NVL(PA26,NVL(PA27,NVL(PA28,NVL(PA29,PA30)))),nvl(decode(pa09,'000',CPM03,CPM04),cp10),nvl(S1.ST02,cp14),nvl(S2.ST02,cp13),CP19,CP09,DECODE(CP24,'1','准/勝','2','駁/敗','') FROM CASEPROGRESS,PATENT,STAFF S1,STAFF S2,CASEPROPERTYMAP WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND cp01=cpm01(+) AND cp10=cpm02(+) and cp14=s1.st01(+)  and cp13=s2.st01(+)  " & strSQL1
'服務
strSql = strSql + " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),NVL(SP08,NVL(SP58,NVL(SP59,NVL(SP65,SP66)))),nvl(decode(sp09,'000',CPM03,CPM04),cp10),nvl(S1.ST02,cp14),nvl(S2.ST02,cp13),CP19,CP09,DECODE(CP24,'1','准/勝','2','駁/敗','') FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cp01=cpm01(+) AND cp10=cpm02(+) and cp14=s1.st01(+)  and cp13=s2.st01(+)  " & strSQL5
'法務
strSql = strSql + " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(LC05,NVL(LC06,LC07)),NVL(LC11,NVL(LC43,NVL(LC44,NVL(LC45,LC46)))),nvl(decode(lc15,'000',CPM03,CPM04),cp10),nvl(S1.ST02,cp14),nvl(S2.ST02,cp13),CP19,CP09,DECODE(CP24,'1','准/勝','2','駁/敗','') FROM CASEPROGRESS,LAWCASE,STAFF S1,STAFF S2,CASEPROPERTYMAP WHERE cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND cp01=cpm01(+) AND cp10=cpm02(+) and cp14=s1.st01(+)  and cp13=s2.st01(+)  " & StrSQL3
'顧問
strSql = strSql + " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04,HC06,NVL(HC05,NVL(HC24,NVL(HC25,NVL(HC26,HC27)))),nvl(CPM03,cp10),nvl(S1.ST02,cp14),nvl(S2.ST02,cp13),CP19,CP09,DECODE(CP24,'1','准/勝','2','駁/敗','') FROM CASEPROGRESS,HIRECASE,STAFF S1,STAFF S2,CASEPROPERTYMAP WHERE cp01=HC01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND cp04=HC04(+) AND cp01=cpm01(+) AND cp10=cpm02(+) and cp14=s1.st01(+)  and cp13=s2.st01(+)  " & StrSQL4
CheckOC
k = 0
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/4
        .MoveFirst
        DoEvents
        Do While .EOF = False
'            For i = 0 To 7
            For i = 0 To 8
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(2) = GetPrjPeople1(strTemp(2))
            strTemp(7) = GetPrjGoBackDate(GetPrjCaseNumber(strTemp(7)))
'            strSQL = "INSERT INTO R050314 VALUES('" & chgsql(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & chgsql(strTemp(3)) & "','" & chgsql(strTemp(4)) & "','" & chgsql(strTemp(5)) & "'," & chgsql(strTemp(6)) & ",'" & chgsql(strTemp(7)) & "','" & strUserNum & "') "
            strSql = "INSERT INTO R050314 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "'," & ChgSQL(strTemp(6)) & ",'" & ChgSQL(strTemp(7)) & "','" & strUserNum & "','" & ChgSQL(strTemp(8)) & "') "
            cnnConnection.Execute strSql
            k = k + 1
            DoEvents
            .MoveNext
        Loop
    End With
Else
   InsertQueryLog (0)  'Add By Sindy 2010/10/4
   ShowNoData
   Screen.MousePointer = vbDefault
   Exit Sub
End If
CheckOC
PrintData
Screen.MousePointer = vbDefault
End Sub

Private Sub PrintData()
strSql = "SELECT * FROM R050314 WHERE ID='" & strUserNum & "' ORDER BY R013001"
CheckOC
Page = 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        PrintTitle
        Do While .EOF = False
            'Modify By Cheng 2002/02/19
'            For i = 0 To 7
            For i = 0 To 6
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(7) = CheckStr(.Fields("r013009").Value)
            strTemp(8) = CheckStr(.Fields("r013008").Value)
            strTemp(1) = StrConv(MidB(StrConv(strTemp(1), vbFromUnicode), 1, 34), vbUnicode)
            strTemp(2) = StrConv(MidB(StrConv(strTemp(2), vbFromUnicode), 1, 8), vbUnicode)
            strTemp(3) = StrConv(MidB(StrConv(strTemp(3), vbFromUnicode), 1, 8), vbUnicode)
            strTemp(4) = StrConv(MidB(StrConv(strTemp(4), vbFromUnicode), 1, 8), vbUnicode)
            strTemp(5) = StrConv(MidB(StrConv(strTemp(5), vbFromUnicode), 1, 8), vbUnicode)
            If iPrint > 10000 Then
                Printer.CurrentX = 0
                Printer.CurrentY = iPrint
                Printer.Print String(200, "-")
                iPrint = iPrint + 300
                Printer.NewPage
                Page = Page + 1
                PrintTitle
            End If
            PrintDatil
            .MoveNext
        Loop
    End With
    Printer.CurrentX = 0
    Printer.CurrentY = iPrint
    Printer.Print String(200, "-")
    Printer.EndDoc
    ShowPrintOk
End If
CheckOC
End Sub

Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
Printer.Print "後金案件表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 6300
Printer.CurrentY = iPrint
If Val(txt1(0)) = 3 Then
    Printer.Print "結果日：" & Format(ChangeTStringToTDateString(txt1(3)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(4))
Else
    Printer.Print "收文日：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
End If
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
If Val(txt1(0)) = 1 Then
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Print "(收文)"
End If
If Val(txt1(0)) = 2 Then
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Print "(無結果)"
End If
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "申請人"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "後金"
'Add By Cheng 2002/02/19
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "准駁/勝敗"
'Printer.CurrentX = PLeft(7)
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "收回日"
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
End Sub

Sub PrintDatil()
For i = 0 To 5
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
Printer.CurrentX = 13100 - Printer.TextWidth(strTemp(6))
Printer.CurrentY = iPrint
Printer.Print strTemp(6)
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print strTemp(7)
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print strTemp(8)
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 0
PLeft(1) = 2200
PLeft(2) = 6500
PLeft(3) = 8000
PLeft(4) = 9500
PLeft(5) = 11000
PLeft(6) = 12500
'Add By Cheng 2002/02/19
PLeft(7) = 13600
'PLeft(7) = 13500
PLeft(8) = 15000
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm050314 = Nothing
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
Case 0
     Select Case Trim(txt1(Index))
     Case "1", "2", "3", ""
     Case Else
          s = MsgBox("查詢別只能輸入 1 或 2 或 3 !!", , "USER 輸入錯誤")
          txt1(0).SetFocus
          txt1(0).SelStart = 0
          txt1(0).SelLength = Len(txt1(0))
          Exit Sub
     End Select
Case 2, 4, 1, 3
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 2 Or Index = 4 Then
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
    End If
Case Else
End Select
End Sub

