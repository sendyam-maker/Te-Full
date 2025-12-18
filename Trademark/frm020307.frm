VERSION 5.00
Begin VB.Form frm020307 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人案件明細表"
   ClientHeight    =   2160
   ClientLeft      =   3360
   ClientTop       =   2910
   ClientWidth     =   3825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   3825
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2844
      TabIndex        =   8
      Top             =   12
      Width           =   972
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2052
      TabIndex        =   7
      Top             =   0
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   2172
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1848
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1008
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1848
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1008
      MaxLength       =   1
      TabIndex        =   1
      Top             =   852
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   2208
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1164
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1008
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1164
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1008
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1500
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1008
      TabIndex        =   0
      Top             =   504
      Width           =   2730
   End
   Begin VB.Line Line4 
      X1              =   2088
      X2              =   2193
      Y1              =   1992
      Y2              =   1992
   End
   Begin VB.Line Line2 
      X1              =   1884
      X2              =   1974
      Y1              =   1344
      Y2              =   1344
   End
   Begin VB.Line Line1 
      X1              =   2112
      X2              =   2187
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label lbl1 
      Height          =   180
      Left            =   2160
      TabIndex        =   15
      Top             =   1524
      Width           =   1620
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   5
      Left            =   108
      TabIndex        =   14
      Top             =   1860
      Width           =   996
   End
   Begin VB.Label Label1 
      Caption         =   "列印別："
      Height          =   180
      Index           =   4
      Left            =   108
      TabIndex        =   13
      Top             =   900
      Width           =   996
   End
   Begin VB.Label Label1 
      Caption         =   "日期："
      Height          =   180
      Index           =   3
      Left            =   108
      TabIndex        =   12
      Top             =   1212
      Width           =   996
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      Height          =   180
      Index           =   2
      Left            =   108
      TabIndex        =   11
      Top             =   1548
      Width           =   996
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   108
      TabIndex        =   10
      Top             =   564
      Width           =   996
   End
   Begin VB.Label Label1 
      Caption         =   "(1.收文  2.發文)"
      Height          =   180
      Index           =   6
      Left            =   1500
      TabIndex        =   9
      Top             =   900
      Width           =   1680
   End
End
Attribute VB_Name = "frm020307"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/9 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay(0 To 2) As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 15) As String, strTemp3 As String, TestOk As Boolean
Dim PLeft(0 To 13) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
      Printer.Orientation = 2
      DoEvents
     If Len(txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         Exit Sub
     Else
         If Len(txt1(1)) = 0 Then
             s = MsgBox("列印別不可空白!!", , "USER 輸入錯誤")
             txt1(1).SetFocus
             Exit Sub
         Else
            'Add By Cheng 2002/03/20
            If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
               Me.txt1(2).SetFocus
               txt1_GotFocus 2
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.txt1(3)) = -1 Then
               Me.txt1(3).SetFocus
               txt1_GotFocus 3
               Exit Sub
            End If
            
             If Len(txt1(3)) = 0 Then
                 s = MsgBox("日期區間不可空白!!", , "USER 輸入錯誤")
                 txt1(2).SetFocus
                 txt1_GotFocus (2)
                 Exit Sub
             Else
                 ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/4 清除查詢印表記錄檔欄位
                 Screen.MousePointer = vbHourglass
                 Me.Enabled = False
                 Process
                 Me.Enabled = True
                 Screen.MousePointer = vbDefault
             End If
         End If
     End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub Process()
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R020307 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL6 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/10/4
End If
StrSQL6 = ""
Select Case Val(txt1(1))
Case 1
     pub_QL05 = pub_QL05 & ";" & Label1(4) & "收文" 'Add By Sindy 2010/10/4
     If Len(txt1(2)) <> 0 Then
         StrSQL6 = StrSQL6 + "  AND CP05>=" & Val(ChangeTStringToWString(txt1(2))) & ""
     End If
     If Len(Trim(txt1(3))) <> 0 Then
         StrSQL6 = StrSQL6 + " AND CP05<=" & Val(ChangeTStringToWString(txt1(3))) & " "
     End If
     If Len(txt1(2)) <> 0 Or Len(Trim(txt1(3))) <> 0 Then
         pub_QL05 = pub_QL05 & ";收文" & Label1(3) & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/10/4
     End If
Case 2
     pub_QL05 = pub_QL05 & ";" & Label1(4) & "發文" 'Add By Sindy 2010/10/4
     If Len(txt1(2)) <> 0 Then
         StrSQL6 = StrSQL6 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(2))) & ""
     End If
     If Len(Trim(txt1(3))) <> 0 Then
         StrSQL6 = StrSQL6 + " AND CP27<=" & Val(ChangeTStringToWString(txt1(3))) & " "
     End If
     If Len(txt1(2)) <> 0 Or Len(Trim(txt1(3))) <> 0 Then
         pub_QL05 = pub_QL05 & ";發文" & Label1(3) & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/10/4
     End If
Case Else
End Select
If Len(txt1(4)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP14='" & txt1(4) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(4) & lbl1 'Add By Sindy 2010/10/4
End If
If Len(txt1(5)) <> 0 Then
    strSQL1 = strSQL1 + " AND TM10>='" & txt1(5) & "' "
    strSQL2 = strSQL2 + " AND SP09>='" & txt1(5) & "' "
End If
If Len(txt1(6)) <> 0 Then
    strSQL1 = strSQL1 + " AND TM10<='" & txt1(6) & "' "
    strSQL2 = strSQL2 + " AND SP09<='" & txt1(6) & "' "
End If
If Len(txt1(5)) <> 0 Or Len(Trim(txt1(6))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(5) & "-" & txt1(6) 'Add By Sindy 2010/10/4
End If
CheckOC
Select Case Val(txt1(1))
Case 1
     strSql = "SELECT cp14,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),nvl(cu04,nvl(cu05||cu88||cu89||cu90,cu06)),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),S2.ST02,NVL(NA03,NA04)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",CP57 FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,NATION,PATENTTRADEMARKMAP,CASEPROPERTYMAP,customer WHERE substr(tm23,1,8)=cu01(+) and decode(substr(tm23,9,1),'','0',substr(tm23,9,1))=cu02(+) and  CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) AND TM10=NA01(+) " & strSQL1 + StrSQL6
     strSql = strSql + " union all select cp14,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),nvl(cu04,nvl(cu05||cu88||cu89||cu90,cu06)),NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),S2.ST02,NVL(NA03,NA04)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",CP57 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,NATION,CASEPROPERTYMAP,customer WHERE substr(sp08,1,8)=cu01(+) and decode(substr(sp08,9,1),'','0',substr(sp08,9,1))=cu02(+) and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SP09=NA01(+) " & strSQL2 + StrSQL6
Case 2
     strSql = "SELECT cp14,CP27,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),nvl(cu04,nvl(cu05||cu88||cu89||cu90,cu06)),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),S2.ST02,NVL(NA03,NA04)," & SQLDate("CP06") & "," & SQLDate("CP05") & ",CP57 FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,NATION,PATENTTRADEMARKMAP,CASEPROPERTYMAP,customer WHERE substr(tm23,1,8)=cu01(+) and decode(substr(tm23,9,1),'','0',substr(tm23,9,1))=cu02(+) and CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND '2'=PTM01(+)  AND TM08=PTM02(+) AND TM10=NA01(+) " & strSQL1 + StrSQL6
     strSql = strSql + " union all select cp14,CP27,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),nvl(cu04,nvl(cu05||cu88||cu89||cu90,cu06)),NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),S2.ST02,NVL(NA03,NA04)," & SQLDate("CP06") & "," & SQLDate("CP05") & ",CP57 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,NATION,CASEPROPERTYMAP,customer WHERE substr(sp08,1,8)=cu01(+) and decode(substr(sp08,9,1),'','0',substr(sp08,9,1))=cu02(+) and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SP09=NA01(+) " & strSQL2 + StrSQL6
Case Else
End Select
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/4
        .MoveFirst
        DoEvents
        Do While .EOF = False
            For i = 0 To 9
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(10) = ""
            strTemp(11) = ""
                 If strTemp(1) >= ChangeTStringToWString(txt1(2)) And strTemp(1) <= ChangeTStringToWString(txt1(3)) Then
                    strTemp(10) = "*"
                 End If
            If CheckStr(.Fields(10)) >= ChangeTStringToWString(txt1(2)) And CheckStr(.Fields(10)) <= ChangeTStringToWString(txt1(3)) Then
                strTemp(11) = "*"
            End If
            strTemp(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
            strSql = "INSERT INTO R020307 VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
            .MoveNext
            DoEvents
        Loop
    Else
      InsertQueryLog (0) 'Add By Sindy 2010/10/4
      ShowNoData
      Screen.MousePointer = vbDefault
      Exit Sub
    End If
End With
CheckOC
PrintData
ShowPrintOk
Screen.MousePointer = vbDefault
End Sub

Sub PrintData()
strSql = "select NVL(st02,R058001),r058002,r058003,r058004,r058005,r058006,r058007,r058008,r058009,r058010,r058011,r058012,r058001 from r020307,staff WHERE r058001=ST01(+) and ID='" & strUserNum & "' ORDER BY R058001,R058002,R058003 "
CheckOC
Page = 1
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        PrintTitle
        SavDay1 = CheckStr(.Fields(0))
        Do While .EOF = False
            For i = 0 To 9
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If SavDay1 <> strTemp(0) Then
                Page = Page + 1
                PrintEnd
                SavDay(0) = "0"
                SavDay(1) = "0"
                SavDay(2) = "0"
                Printer.NewPage
                SavDay1 = strTemp(0)
                PrintTitle
            End If
            If Len(CheckStr(.Fields(11))) <> 0 Then
                SavDay(2) = Trim(str(Val(SavDay(2)) + 1))
            Else
                SavDay(1) = Trim(str(Val(SavDay(1)) + 1))
            End If
            SavDay(0) = Trim(str(Val(SavDay(0)) + 1))
            strTemp(3) = StrToStr(strTemp(3), 25)
            strTemp(4) = StrToStr(strTemp(4), 10)
            strTemp(5) = StrToStr(strTemp(5), 4)
            strTemp(6) = StrToStr(strTemp(6), 4)
            strTemp(7) = StrToStr(strTemp(7), 4)
            PrintDatil
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    End If
End With
PrintEnd
Printer.EndDoc
CheckOC
End Sub

Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print "內商承辦人案件明細表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6200
Printer.CurrentY = iPrint
If txt1(1) = "1" Then
    Printer.Print "收文日：" & Format(ChangeTStringToTDateString(txt1(2)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(3))
Else
    Printer.Print "發文日：" & Format(ChangeTStringToTDateString(txt1(2)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(3))
End If
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "承辦人：" & SavDay1
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.Font.Size = 10
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Select Case Val(txt1(1))
Case 1
     Printer.Print "收文日"
Case 2
     Printer.Print "發文日"
Case Else
End Select
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "申請人"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "申請國家"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "本所期限"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Select Case Val(txt1(1))
Case 1
     Printer.Print "發文日"
Case 2
     Printer.Print "收文日"
Case Else
End Select
iPrint = iPrint + 300
Printer.Font.Size = 12
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.Font.Size = 10
End Sub

Sub PrintDatil()
For i = 1 To 9
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 500
PLeft(2) = 1500
PLeft(3) = 3300
PLeft(4) = 8500
PLeft(5) = 11000
PLeft(6) = 12000
PLeft(7) = 13000
PLeft(8) = 14000
PLeft(9) = 15000
End Sub

Sub PrintEnd()
Printer.Font.Size = 12
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.Font.Size = 10
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "總件數：" & SavDay(0)
Printer.CurrentX = 3000
Printer.CurrentY = iPrint
Select Case Val(txt1(1))
Case 1
     Printer.Print "收文件數：" & SavDay(1)
Case 2
     Printer.Print "發文件數：" & SavDay(1)
Case Else
End Select
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
Printer.Print "取消收文件數：" & SavDay(2)
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm020307 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdok(0).SetFocus
End If
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0
     strTemp1 = Split(UCase(GetSystemKindByNick), ",")
     strTemp2 = Split(UCase(txt1(0)), ",")
     For i = 0 To UBound(strTemp2)
        s = 0
        For j = 0 To UBound(strTemp1)
            If strTemp2(i) = strTemp1(j) Then
                s = 1
                Exit For
            End If
        Next j
        If s = 0 Then
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限!! ", , "USER 權限問題")
            txt1(0).SetFocus
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            Exit Sub
        End If
     Next i
Case 1
     Select Case Trim(txt1(1))
     Case "1", "2", ""
     Case Else
          s = MsgBox("列印別只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          txt1(1).SetFocus
          txt1(1).SelStart = 0
          txt1(1).SelLength = Len(txt1(1))
          Exit Sub
     End Select
Case 4
     lbl1 = GetPrjSalesNM(txt1(4))
     If Trim(txt1(4)) <> "" Then
        If Trim(lbl1.Caption) = "" Then
            s = MsgBox("承辦人輸入錯誤！", , "錯誤！")
            txt1(4).SetFocus
            txt1_GotFocus (4)
            Exit Sub
        End If
     End If
Case 2, 3
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 3 Then
     If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
   End If
Case 6
     If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If

Case Else
End Select
End Sub

