VERSION 5.00
Begin VB.Form frm020312 
   BorderStyle     =   1  '單線固定
   Caption         =   "不出名案件明細表"
   ClientHeight    =   1800
   ClientLeft      =   2400
   ClientTop       =   2895
   ClientWidth     =   3405
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   3405
   Begin VB.CommandButton CMDOK 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2640
      TabIndex        =   10
      Top             =   36
      Width           =   756
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1848
      TabIndex        =   9
      Top             =   36
      Width           =   756
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   2
      Left            =   2364
      MaxLength       =   4
      TabIndex        =   2
      Top             =   792
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   1
      Left            =   1248
      MaxLength       =   4
      TabIndex        =   1
      Top             =   792
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   6
      Left            =   2364
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1500
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   5
      Left            =   1248
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1488
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   4
      Left            =   2364
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1152
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   3
      Left            =   1248
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1152
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   0
      Left            =   1248
      TabIndex        =   0
      Top             =   456
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "發文日期："
      Height          =   180
      Index           =   1
      Left            =   108
      TabIndex        =   8
      Top             =   1536
      Width           =   1200
   End
   Begin VB.OptionButton Option1 
      Caption         =   "催審期限："
      Height          =   180
      Index           =   0
      Left            =   108
      TabIndex        =   7
      Top             =   1236
      Value           =   -1  'True
      Width           =   1200
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1800
      X2              =   3000
      Y1              =   924
      Y2              =   924
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1776
      X2              =   2976
      Y1              =   1644
      Y2              =   1644
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1752
      X2              =   2952
      Y1              =   1296
      Y2              =   1296
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   1
      Left            =   108
      TabIndex        =   12
      Top             =   852
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   108
      TabIndex        =   11
      Top             =   516
      Width           =   948
   End
End
Attribute VB_Name = "frm020312"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/9 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay(0 To 1) As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 15) As String, strTemp3 As String, TestOk As Boolean
Dim PLeft(0 To 13) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
     Printer.Orientation = 2
     DoEvents
     If Len(TXT1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         TXT1(0).SetFocus
         Exit Sub
     Else
         If Option1(0).Value = True Then
            'Add By Cheng 2002/03/21
            If PUB_CheckKeyInDate(Me.TXT1(3)) = -1 Then
               Me.TXT1(3).SetFocus
               txt1_GotFocus 3
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.TXT1(4)) = -1 Then
               Me.TXT1(4).SetFocus
               txt1_GotFocus 4
               Exit Sub
            End If
             
             If Len(TXT1(4)) = 0 Then
                 s = MsgBox("催審期限區間不可空白!!", , "USER 輸入錯誤")
                 TXT1(3).SetFocus
                 txt1_GotFocus (3)
                 Exit Sub
             End If
         Else
            'Add By Cheng 2002/03/21
            If PUB_CheckKeyInDate(Me.TXT1(5)) = -1 Then
               Me.TXT1(5).SetFocus
               txt1_GotFocus 5
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.TXT1(6)) = -1 Then
               Me.TXT1(6).SetFocus
               txt1_GotFocus 6
               Exit Sub
            End If
             If Len(TXT1(6)) = 0 Then
                 s = MsgBox("發文日期區間不可空白!!", , "USER 輸入錯誤")
                 TXT1(5).SetFocus
                 txt1_GotFocus (5)
                 Exit Sub
             End If
         End If
         ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/4 清除查詢印表記錄檔欄位
         Screen.MousePointer = vbHourglass
         Me.Enabled = False
         Process
         Me.Enabled = True
         Screen.MousePointer = vbDefault
     End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub Process()
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R020312 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL6 = ""
StrSQL6 = ""
If Len(TXT1(1)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP10>='" & TXT1(1) & "' "
End If
If Len(TXT1(2)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP10<='" & TXT1(2) & "' "
End If
If Len(TXT1(1)) <> 0 Or Len(TXT1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & TXT1(1) & "-" & TXT1(2) 'Add By Sindy 2010/10/4
End If
If Option1(0).Value = True Then
   If Len(TXT1(0)) <> 0 Then
      strSQL1 = strSQL1 + " AND nP02 IN (" & SQLGrpStr(TXT1(0), 2) & ") "
      strSQL2 = strSQL2 + " AND nP02 IN (" & SQLGrpStr(TXT1(0), 5) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1(0) & TXT1(0)  'Add By Sindy 2010/10/4
   End If
   If Len(TXT1(3)) <> 0 Then
      StrSQL6 = StrSQL6 + "AND NP08>=" & Val(ChangeTStringToWString(TXT1(3))) & " "
   End If
   If Len(Trim(TXT1(4))) <> 0 Then
      StrSQL6 = StrSQL6 + "AND NP08<=" & Val(ChangeTStringToWString(TXT1(4))) & " "
   End If
   If Len(TXT1(3)) <> 0 Or Len(Trim(TXT1(4))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & TXT1(3) & "-" & TXT1(4) 'Add By Sindy 2010/10/4
   End If
    strSql = "SELECT " & SQLDate("CP27") & "," & SQLDate("NP08") & ",NVL(TM15,TM12),NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(TM05,NVL(TM06,TM07)),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(S1.ST02,CP14),NVL(S2.ST02,CP13),'" & strUserNum & "' FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CUSTOMER,CASEPROPERTYMAP WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP22='N' AND NP07=305 AND (NP06='' OR NP06 IS NULL) AND (TM29='N' OR TM29='' OR TM29 IS NULL) " & strSQL1 + StrSQL6
    strSql = strSql + " union all select " & SQLDate("CP27") & "," & SQLDate("NP08") & ",NVL(SP14,SP11),NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(S1.ST02,CP14),NVL(S2.ST02,CP13),'" & strUserNum & "' FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CUSTOMER,CASEPROPERTYMAP WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP22='N' AND NP07=305 AND (NP06='' OR NP06 IS NULL) AND (SP15='N' OR SP15='' OR SP15 IS NULL) " & strSQL2 + StrSQL6
Else
   If Len(TXT1(0)) <> 0 Then
      strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(TXT1(0), 2) & ") "
      strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(TXT1(0), 5) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1(0) & TXT1(0)  'Add By Sindy 2010/10/4
   End If
   If Len(TXT1(5)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP27>=" & Val(ChangeTStringToWString(TXT1(5))) & " "
   End If
   If Len(Trim(TXT1(6))) <> 0 Then
      StrSQL6 = StrSQL6 + " AND CP27<=" & Val(ChangeTStringToWString(TXT1(6))) & " "
   End If
   If Len(TXT1(5)) <> 0 Or Len(Trim(TXT1(6))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & TXT1(5) & "-" & TXT1(6) 'Add By Sindy 2010/10/4
   End If
    strSql = "SELECT " & SQLDate("CP27") & ",' ',NVL(TM15,TM12),CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(S1.ST02,CP14),NVL(S2.ST02,CP13),'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CUSTOMER,CASEPROPERTYMAP WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP22='N' AND (TM29='N' OR TM29='' OR TM29 IS NULL) " & strSQL1 + StrSQL6
    strSql = strSql + " union all select " & SQLDate("CP27") & ",' ',NVL(SP14,SP11),CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(S1.ST02,CP14),NVL(S2.ST02,CP13),'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CUSTOMER,CASEPROPERTYMAP WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP22='N' AND (SP15='N' OR SP15='' OR SP15 IS NULL) " & strSQL2 + StrSQL6
End If
cnnConnection.Execute "INSERT INTO R020312 " & strSql
strSql = "SELECT * FROM R020312 WHERE ID='" & strUserNum & "' "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/4
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
strSql = "select * from r020312 where ID='" & strUserNum & "' "
If Option1(0).Value = True Then
    strSql = strSql + " ORDER BY R062002,R062004 "
Else
    strSql = strSql + " ORDER BY R062001,R062004 "
End If
CheckOC
Page = 1
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        PrintTitle
        Do While .EOF = False
            For i = 0 To 8
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(2) = StrToStr(strTemp(2), 8)
            strTemp(4) = StrToStr(strTemp(4), 10)
            strTemp(5) = StrToStr(strTemp(5), 5)
            strTemp(6) = StrToStr(strTemp(6), 11)
            strTemp(7) = StrToStr(strTemp(7), 5)
            strTemp(8) = StrToStr(strTemp(8), 4)
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
CheckOC
Printer.EndDoc
End Sub

Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
Printer.Print "內商不出名案件明細表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6200
Printer.CurrentY = iPrint
If Option1(0).Value = True Then
    Printer.Print "催審期限：" & Format(ChangeTStringToTDateString(TXT1(3)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(TXT1(4))
Else
    Printer.Print "發文日期：" & Format(ChangeTStringToTDateString(TXT1(5)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(TXT1(6))
End If
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "發文日"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "催審期限"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "申請案號/審定號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "申請人"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
End Sub

Sub PrintDatil()
For i = 0 To 8
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 1700
PLeft(2) = 3000
PLeft(3) = 5000
PLeft(4) = 6800
PLeft(5) = 9500
PLeft(6) = 10800
PLeft(7) = 13500
PLeft(8) = 15000
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
TXT1(0) = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm020312 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
   TXT1(3).SetFocus
   txt1_GotFocus (3)
Case 1
   TXT1(5).SetFocus
   txt1_GotFocus (5)
End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
TXT1(Index).SelStart = 0
TXT1(Index).SelLength = Len(TXT1(Index))
End Sub

Private Sub txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CMDOK(0).SetFocus
End If
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0
     strTemp1 = Split(UCase(GetSystemKindByNick), ",")
     strTemp2 = Split(UCase(TXT1(0)), ",")
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
            TXT1(0).SetFocus
            TXT1(0).SelStart = 0
            TXT1(0).SelLength = Len(TXT1(0))
            Exit Sub
        End If
     Next i
Case 2
     If RunNick(TXT1(Index - 1), TXT1(Index)) Then
         TXT1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
Case 3, 4, 5, 6
   If PUB_CheckKeyInDate(Me.TXT1(Index)) = -1 Then

      Me.TXT1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 4 Or Index = 6 Then
     If RunNick(TXT1(Index - 1), TXT1(Index)) Then
         TXT1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
    End If
Case Else
End Select
End Sub

