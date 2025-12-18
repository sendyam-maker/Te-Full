VERSION 5.00
Begin VB.Form frm040316 
   BorderStyle     =   1  '單線固定
   Caption         =   "不出名案件明細表"
   ClientHeight    =   1680
   ClientLeft      =   2310
   ClientTop       =   1425
   ClientWidth     =   3150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   3150
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2256
      TabIndex        =   8
      Top             =   20
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1440
      TabIndex        =   7
      Top             =   20
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   2184
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1320
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1116
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1320
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2184
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1032
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1116
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1032
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2124
      MaxLength       =   4
      TabIndex        =   2
      Top             =   732
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1008
      MaxLength       =   4
      TabIndex        =   1
      Top             =   732
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   996
      TabIndex        =   0
      Top             =   432
      Width           =   2070
   End
   Begin VB.OptionButton Option1 
      Caption         =   "發文日："
      Height          =   180
      Index           =   1
      Left            =   36
      TabIndex        =   12
      Top             =   1368
      Width           =   1260
   End
   Begin VB.OptionButton Option1 
      Caption         =   "催審期限："
      Height          =   180
      Index           =   0
      Left            =   36
      TabIndex        =   11
      Top             =   1080
      Value           =   -1  'True
      Width           =   1260
   End
   Begin VB.Line Line3 
      X1              =   2052
      X2              =   2157
      Y1              =   1512
      Y2              =   1512
   End
   Begin VB.Line Line2 
      X1              =   2052
      X2              =   2172
      Y1              =   1236
      Y2              =   1236
   End
   Begin VB.Line Line1 
      X1              =   1944
      X2              =   2154
      Y1              =   864
      Y2              =   864
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   1
      Left            =   36
      TabIndex        =   10
      Top             =   792
      Width           =   1128
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   36
      TabIndex        =   9
      Top             =   504
      Width           =   1128
   End
End
Attribute VB_Name = "frm040316"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, strSQL2 As String, strSQL5 As String, StrSQL6 As String, i As Integer, j As Integer, s As Integer
Dim strTemp(0 To 8) As String, strTemp1 As Variant, strTemp2 As Variant, k As Integer, StrSQL3 As String
Dim PLeft(0 To 8) As Integer, iPrint As Integer, Page As Integer, strTemp3(0 To 2) As String
'Add By Cheng 2002/09/11
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
     If Len(txt1(0)) = 0 Then
        s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
        txt1(0).SetFocus
        Exit Sub
     Else
         'Add By Cheng 2002/09/11
         If Me.txt1(1).Text <> "" And Me.txt1(2).Text <> "" Then
            If Me.txt1(1).Text > Me.txt1(2).Text Then
               MsgBox "案件性質範圍輸入錯誤!!!", vbExclamation + vbOKOnly
               blnClkSure = True
               Me.txt1(1).SetFocus
               txt1_GotFocus 1
               Exit Sub
            End If
         End If
         
         If Option1(0).Value = True Then
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
            'Add By Cheng 2002/09/11
            If Me.txt1(3).Text <> "" And Me.txt1(4).Text <> "" Then
               If Val(Me.txt1(3).Text) > Val(Me.txt1(4).Text) Then
                  MsgBox "催審期限範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(3).SetFocus
                  txt1_GotFocus 3
                  Exit Sub
               End If
            End If
             
             If Len(txt1(4)) = 0 Then
                s = MsgBox("催審期限區間不可空白!!", , "USER輸入錯誤")
                txt1(3).SetFocus
                txt1_GotFocus (3)
                Exit Sub
             Else
                Screen.MousePointer = vbHourglass
                Me.Enabled = False
                ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/2 清除查詢印表記錄檔欄位
                Process
                Me.Enabled = True
                Screen.MousePointer = vbDefault
             End If
         Else
             'Add By Cheng 2002/03/19
            If PUB_CheckKeyInDate(Me.txt1(5)) = -1 Then
               Me.txt1(5).SetFocus
               txt1_GotFocus 5
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.txt1(6)) = -1 Then
               Me.txt1(6).SetFocus
               txt1_GotFocus 6
               Exit Sub
            End If
            'Add By Cheng 2002/09/11
            If Me.txt1(5).Text <> "" And Me.txt1(6).Text <> "" Then
               If Val(Me.txt1(5).Text) > Val(Me.txt1(6).Text) Then
                  MsgBox "發文日範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(5).SetFocus
                  txt1_GotFocus 5
                  Exit Sub
               End If
            End If
             If Len(txt1(6)) = 0 Then
                s = MsgBox("發文日區間不可空白!!", , "USER輸入錯誤")
                txt1(5).SetFocus
                txt1_GotFocus (5)
                Exit Sub
             Else
                Screen.MousePointer = vbHourglass
                Me.Enabled = False
                ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/2 清除查詢印表記錄檔欄位
                Process1
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
cnnConnection.Execute "DELETE FROM R040316 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
strSQL5 = ""
StrSQL6 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " and NP02 in (" & SQLGrpStr(txt1(0), 1) & ") "
   strSQL2 = strSQL2 + " and NP02 in (" & SQLGrpStr(txt1(0), 2) & ") "
   strSQL5 = strSQL5 + " and NP02 in (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/2
End If
StrSQL6 = " AND NP06 IS NULL AND NP07=411 AND CP22='N' "
If Len(Trim(txt1(3))) <> 0 Then
   StrSQL6 = StrSQL6 + " AND NP08>=" & Val(ChangeTStringToWString(txt1(3))) & " "
End If
If Len(Trim(txt1(4))) <> 0 Then
   StrSQL6 = StrSQL6 & " AND NP08<=" & Val(ChangeTStringToWString(txt1(4))) & " "
End If
If Len(Trim(txt1(3))) <> 0 Or Len(Trim(txt1(4))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/12/2
End If
If Len(txt1(1)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP10>='" & txt1(1) & "' "
End If
If Len(txt1(2)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP10<='" & txt1(2) & "' "
End If
If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/2
End If

strSql = "SELECT " & SQLDate("CP27") & "," & SQLDate("NP08") & ",NVL(PA15,PA11),NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(PA05,NVL(PA06,PA07)),decode(pa09,'000',cpm03,cpm04),NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),S1.ST02,S2.ST02,'" & strUserNum & "' FROM NEXTPROGRESS,CASEPROGRESS,PATENT,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE CP09=NP01(+) AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND cp01=cpm01(+) AND CP10=CPM02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) " & strSQL1 & StrSQL6
strSql = strSql + " union ALL SELECT " & SQLDate("CP27") & "," & SQLDate("NP08") & ",NVL(TM15,TM12),NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(TM05,NVL(TM06,TM07)),decode(tm10,'000',cpm03,cpm04),NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),S1.ST02,S2.ST02,'" & strUserNum & "' FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE CP09=NP01(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND cp01=cpm01(+) AND CP10=CPM02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+) " & strSQL2 & StrSQL6
strSql = strSql + " union ALL SELECT " & SQLDate("CP27") & "," & SQLDate("NP08") & ",NVL(SP14,SP11),NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),decode(sp09,'000',cpm03,cpm04),NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),S1.ST02,S2.ST02,'" & strUserNum & "' FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE CP09=NP01(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND cp01=cpm01(+) AND CP10=CPM02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) " & StrSQL3 & StrSQL6
cnnConnection.Execute "insert into r040316 " & strSql
strSql = "sleect * from r040316 where id='" & strUserNum & "'"
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
k = 0
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/12/2
Else
   InsertQueryLog (0) 'Add By Sindy 2010/12/2
   ShowNoData
   CheckOC
   Screen.MousePointer = vbDefault
   Exit Sub
End If
CheckOC
PrintData
ShowPrintOk
Screen.MousePointer = vbDefault
End Sub

Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print "不出名案件明細表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6300
Printer.CurrentY = iPrint
If Option1(0).Value = True Then
    Printer.Print "催審期限：" & Format(ChangeTStringToTDateString(txt1(3)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(4))
Else
    Printer.Print "發文日：" & Format(ChangeTStringToTDateString(txt1(5)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(6))
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
Printer.Print "申請案號/專利號"
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

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 1500
PLeft(2) = 2500
PLeft(3) = 4500
PLeft(4) = 6300
PLeft(5) = 9500
PLeft(6) = 10800
PLeft(7) = 13200
PLeft(8) = 14500
End Sub

Sub PrintDatil()
For i = 0 To 8
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub


Sub PrintData()
If Option1(0).Value = True Then
    strSql = "SELECT * FROM R040316 WHERE ID='" & strUserNum & "' ORDER BY R029002,R029004 "
Else
    strSql = "SELECT * FROM R040316 WHERE ID='" & strUserNum & "' ORDER BY R029001,R029004 "
End If
CheckOC
Page = 1
strTemp3(0) = " "
strTemp3(1) = " "
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        PrintTitle
        Do While .EOF = False
            For i = 0 To 8
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(2) = StrToStr(strTemp(2), 5)
            strTemp(4) = StrToStr(strTemp(4), 13)
            strTemp(5) = StrToStr(strTemp(5), 4)
            strTemp(6) = StrToStr(strTemp(6), 9)
            strTemp(7) = StrToStr(strTemp(7), 4)
            strTemp(8) = StrToStr(strTemp(8), 4)
            If Option1(0).Value = True Then
                If strTemp3(0) <> strTemp(1) Then
                    strTemp3(0) = strTemp(1)
                    strTemp3(1) = strTemp(3)
                Else
                    strTemp(1) = ""
                    If strTemp3(1) <> strTemp(3) Then
                        strTemp3(1) = strTemp(3)
                    Else
                        strTemp(3) = ""
                    End If
                End If
            Else
                If strTemp3(0) <> strTemp(0) Then
                    strTemp3(0) = strTemp(0)
                    strTemp3(1) = strTemp(3)
                Else
                    strTemp(0) = ""
                    If strTemp3(1) <> strTemp(3) Then
                        strTemp3(1) = strTemp(3)
                    Else
                        strTemp(3) = ""
                    End If
                End If
            End If
            PrintDatil
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    End With
End If
CheckOC
Printer.EndDoc
End Sub

Sub Process1()
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R040316 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
strSQL5 = ""
StrSQL6 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " and CP01 in (" & SQLGrpStr(txt1(0), 1) & ") "
   strSQL2 = strSQL2 + " and CP01 in (" & SQLGrpStr(txt1(0), 2) & ") "
   strSQL5 = strSQL5 + " and CP01 in (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/2
End If
StrSQL6 = " AND CP22='N' AND CP24 IS NULL "
If Len(Trim(txt1(5))) <> 0 Then
   StrSQL6 = StrSQL6 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(5))) & " "
End If
If Len(Trim(txt1(6))) <> 0 Then
   StrSQL6 = StrSQL6 & " and CP05<=" & Val(ChangeTStringToWString(txt1(6))) & " "
End If
If Len(Trim(txt1(5))) <> 0 Or Len(Trim(txt1(6))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & txt1(5) & "-" & txt1(6) 'Add By Sindy 2010/12/2
End If
If Len(txt1(1)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP10>='" & txt1(1) & "' "
End If
If Len(txt1(2)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP10<='" & txt1(2) & "' "
End If
If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/2
End If

strSql = "SELECT " & SQLDate("CP27") & ",' ',NVL(PA15,PA11),CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),decode(pa09,'000',cpm03,cpm04),NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),S1.ST02,S2.ST02,'" & strUserNum & "' FROM CASEPROGRESS,PATENT,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp01=cpm01(+) AND CP10=CPM02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) " & strSQL1 & StrSQL6
strSql = strSql + " union ALL SELECT " & SQLDate("CP27") & ",' ',NVL(TM15,TM12),CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),decode(tm10,'000',cpm03,cpm04),NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),S1.ST02,S2.ST02,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND cp01=cpm01(+) AND CP10=CPM02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+)  " & strSQL2 & StrSQL6
strSql = strSql + " union ALL SELECT " & SQLDate("CP27") & ",' ',NVL(SP14,SP11),CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),decode(sp09,'000',cpm03,cpm04),NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),S1.ST02,S2.ST02,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND cp01=cpm01(+) AND CP10=CPM02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+)  " & strSQL5 & StrSQL6
cnnConnection.Execute "insert into r040316 " & strSql
strSql = "select * from r040316 where id='" & strUserNum & "' "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
k = 0
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/12/2
Else
   InsertQueryLog (0) 'Add By Sindy 2010/12/2
   ShowNoData
   CheckOC
   Screen.MousePointer = vbDefault
   Exit Sub
End If
CheckOC
PrintData
ShowPrintOk
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm040316 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
On Error Resume Next
If Option1(0).Value = True Then
   txt1(3).Enabled = True
   txt1(4).Enabled = True
   txt1(5).Enabled = False
   txt1(6).Enabled = False
   txt1(3).SetFocus
   txt1_GotFocus (3)
Else
   txt1(3).Enabled = False
   txt1(4).Enabled = False
   txt1(5).Enabled = True
   txt1(6).Enabled = True
   txt1(5).SetFocus
   txt1_GotFocus (5)
End If
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
Case 2, 4, 6
   'Modify By Cheng 2002/09/11
   If blnClkSure = False Then
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
   Else
      blnClkSure = False
   End If
Case Else
End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 3, 4, 5, 6 '催審期限起, 迄, 發文日起, 迄
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
   End If
End Select
If Cancel Then TextInverse txt1(Index)
End Sub
