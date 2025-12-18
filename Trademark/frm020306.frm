VERSION 5.00
Begin VB.Form frm020306 
   BorderStyle     =   1  '單線固定
   Caption         =   "智權人員收文明細表"
   ClientHeight    =   3015
   ClientLeft      =   3645
   ClientTop       =   1965
   ClientWidth     =   4425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4425
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1170
      TabIndex        =   0
      Top             =   735
      Width           =   2736
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1170
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1065
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2370
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1065
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1170
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1410
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1170
      MaxLength       =   7
      TabIndex        =   5
      Top             =   2130
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   2370
      MaxLength       =   7
      TabIndex        =   6
      Top             =   2130
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1770
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   1170
      MaxLength       =   4
      TabIndex        =   7
      Top             =   2475
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   2370
      MaxLength       =   4
      TabIndex        =   8
      Top             =   2475
      Width           =   1110
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2565
      TabIndex        =   9
      Top             =   150
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3360
      TabIndex        =   10
      Top             =   150
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "(1.收文  2.取消收文)"
      Height          =   180
      Index           =   6
      Left            =   1665
      TabIndex        =   18
      Top             =   1815
      Width           =   1680
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   270
      TabIndex        =   17
      Top             =   795
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "業務區："
      Height          =   180
      Index           =   1
      Left            =   270
      TabIndex        =   16
      Top             =   1125
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員："
      Height          =   180
      Index           =   2
      Left            =   270
      TabIndex        =   15
      Top             =   1470
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "日期："
      Height          =   180
      Index           =   3
      Left            =   255
      TabIndex        =   14
      Top             =   2175
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "列印別："
      Height          =   180
      Index           =   4
      Left            =   285
      TabIndex        =   13
      Top             =   1815
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   5
      Left            =   270
      TabIndex        =   12
      Top             =   2490
      Width           =   990
   End
   Begin VB.Label lbl1 
      Height          =   180
      Left            =   2325
      TabIndex        =   11
      Top             =   1470
      Width           =   1620
   End
   Begin VB.Line Line1 
      X1              =   2280
      X2              =   2355
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Line Line2 
      X1              =   2295
      X2              =   2385
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line4 
      X1              =   2280
      X2              =   2385
      Y1              =   2610
      Y2              =   2610
   End
End
Attribute VB_Name = "frm020306"
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
         If Len(txt1(4)) = 0 Then
             s = MsgBox("列印別不可空白!!", , "USER 輸入錯誤")
             txt1(4).SetFocus
             Exit Sub
         Else
             'Add By Cheng 2002/03/21
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
             
             If Len(txt1(6)) = 0 Then
                 s = MsgBox("日期區間不可空白!!", , "USER 輸入錯誤")
                 txt1(5).SetFocus
                 txt1_GotFocus (5)
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
cnnConnection.Execute "DELETE FROM R020306 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL6 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/10/4
End If
StrSQL6 = ""
Select Case Val(txt1(4))
Case 1 '列印收文
     pub_QL05 = pub_QL05 & ";" & Label1(4) & "收文" 'Add By Sindy 2010/10/4
     If Len(txt1(5)) <> 0 Then
         StrSQL6 = StrSQL6 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(5))) & ""
     End If
     If Len(Trim(txt1(6))) <> 0 Then
         StrSQL6 = StrSQL6 + " AND CP05<=" & Val(ChangeTStringToWString(txt1(6))) & " "
     End If
     If Len(txt1(5)) <> 0 Or Len(Trim(txt1(6))) <> 0 Then
         pub_QL05 = pub_QL05 & ";收文" & Label1(3) & txt1(5) & "-" & txt1(6)  'Add By Sindy 2010/10/4
     End If
Case 2 '列印取消收文
      pub_QL05 = pub_QL05 & ";" & Label1(4) & "取消收文" 'Add By Sindy 2010/10/4
      'Add By Cheng 2002/04/30
      StrSQL6 = StrSQL6 + " AND CP27 IS NULL "
     If Len(txt1(5)) <> 0 Then
         StrSQL6 = StrSQL6 + " AND CP57>=" & Val(ChangeTStringToWString(txt1(5))) & ""
     End If
     If Len(Trim(txt1(6))) <> 0 Then
         StrSQL6 = StrSQL6 + " AND CP57<=" & Val(ChangeTStringToWString(txt1(6))) & " "
     End If
     If Len(txt1(5)) <> 0 Or Len(Trim(txt1(6))) <> 0 Then
         pub_QL05 = pub_QL05 & ";取消收文" & Label1(3) & txt1(5) & "-" & txt1(6)  'Add By Sindy 2010/10/4
     End If
Case Else
End Select
If Len(txt1(1)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP12>='" & txt1(1) & "' "
End If
If Len(txt1(2)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP12<='" & txt1(2) & "' "
End If
If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(1) & "-" & txt1(2)  'Add By Sindy 2010/10/4
End If
If Len(txt1(3)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP13='" & txt1(3) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(3) & lbl1  'Add By Sindy 2010/10/4
End If
If Len(txt1(7)) <> 0 Then
    strSQL1 = strSQL1 + " AND TM10>='" & txt1(7) & "' "
    strSQL2 = strSQL2 + " AND SP09>='" & txt1(7) & "' "
End If
If Len(txt1(8)) <> 0 Then
    strSQL1 = strSQL1 + " AND TM10<='" & txt1(8) & "' "
    strSQL2 = strSQL2 + " AND SP09<='" & txt1(8) & "' "
End If
If Len(txt1(7)) <> 0 Or Len(txt1(8)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(7) & "-" & txt1(8)  'Add By Sindy 2010/10/4
End If
CheckOC
'**************** 將業務區改成抓案件進度檔   91.08.15  nick
'Modify By Cheng 2001/12/26
'總收文號小於 "B" 開頭號數
'strSQL = "SELECT s1.st03,cp13," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),TM09,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(S2.ST02,CP13)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",CP57 FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,CUSTOMER,STAFF S1,STAFF S2 WHERE CP09<'C' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) " & strSQL1 + StrSQL6
'Modify By Cheng 2002/02/19
'strSQL = "SELECT s1.st03,cp13," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),TM09,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(S2.ST02,CP13)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",CP57 FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,CUSTOMER,STAFF S1,STAFF S2 WHERE CP09<'B' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) " & strSQL1 + StrSQL6

'Modify By Cheng 2002/12/16
'只抓承辦人CP14之ST03為"P2"字頭 or CP14 IS NULL
'strSQL = "SELECT cp12,cp13," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),TM09,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(S2.ST02,CP13)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",CP57,NVL(NA03,NA04) AS NATIONNAME FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,CUSTOMER,STAFF S1,STAFF S2,Nation WHERE CP09<'B' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) And TM10=NA01(+) " & strSQL1 + StrSQL6
strSql = "SELECT cp12,cp13," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),TM09,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(S2.ST02,CP14)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",CP57,NVL(NA03,NA04) AS NATIONNAME FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,CUSTOMER,STAFF S1,STAFF S2,Nation WHERE CP09<'B' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) And TM10=NA01(+) AND (SUBSTR(S2.ST03,1,2) = 'P2' OR CP14 IS NULL ) " & strSQL1 + StrSQL6
'strSQL = strSQL + " UNION all SELECT s1.st03,cp13," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),' ',NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(S2.ST02,CP13)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",CP57 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,CUSTOMER,STAFF S1,STAFF S2 WHERE CP09<'C' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) " & strSQL2 + StrSQL6
'Modify By Cheng 2002/12/16
'只抓承辦人CP14之ST03為"P2"字頭 or CP14 IS NULL
'strSQL = strSQL + " UNION all SELECT cp12,cp13," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),' ',NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(S2.ST02,CP13)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",CP57,NVL(NA03,NA04) AS NATIONNAME FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,CUSTOMER,STAFF S1,STAFF S2,NATION WHERE CP09<'B' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND SP09=NA01(+) " & strSQL2 + StrSQL6
strSql = strSql + " UNION all SELECT cp12,cp13," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),' ',NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(S2.ST02,CP14)," & SQLDate("CP06") & "," & SQLDate("CP27") & ",CP57,NVL(NA03,NA04) AS NATIONNAME FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,CUSTOMER,STAFF S1,STAFF S2,NATION WHERE CP09<'B' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND SP09=NA01(+) AND (SUBSTR(S2.ST03,1,2) = 'P2' OR CP14 IS NULL ) " & strSQL2 + StrSQL6

With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/4
        .MoveFirst

        DoEvents
        Do While .EOF = False
            For i = 0 To 11
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(12) = ""
            If Val(ChangeWStringToTString(strTemp(11))) >= Val(txt1(5)) And Val(ChangeWStringToTString(strTemp(11))) <= Val(txt1(6)) Then
                strTemp(12) = "*"
            End If
            strTemp(11) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(11)))
            strTemp(13) = Left("" & .Fields("NATIONNAME").Value, 8)
            'If txt1(4) = "1" Then
               strSql = "INSERT INTO R020306 VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & strUserNum & "','" & ChgSQL(strTemp(13)) & "') "
            'Else
            '   StrSQL = "INSERT INTO R020306 VALUES ('" & chgsql(strTemp(0)) & "','" & chgsql(strTemp(1)) & "','" & chgsql(strTemp(11)) & "','" & chgsql(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & chgsql(strTemp(5)) & "','" & chgsql(strTemp(6)) & "','" & chgsql(strTemp(7)) & "','" & chgsql(strTemp(8)) & "','" & chgsql(strTemp(9)) & "','" & chgsql(strTemp(10)) & "','" & chgsql(strTemp(2)) & "','" & chgsql(strTemp(12)) & "','" & strUserNum & "') "
            'End If
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
Screen.MousePointer = vbDefault
End Sub

Sub PrintData()
'Modify By Cheng 2002/02/19
'strSQL = "SELECT nvl(a0902,a0903),nvl(st02,r057002),r057003,r057004,r057005,r057006,r057007,r057008,r057009,r057010,r057011,r057012,r057013,r057001,r057002 FROM R020306,acc090,staff WHERE r057002=st01(+) and r057001=a0901(+) and ID='" & strUserNum & "' ORDER BY R057001,R057002,R057003"
'Modify By Cheng 2003/01/30
'再依本所案號排序
'strSQL = "SELECT nvl(a0902,a0903),nvl(st02,r057002),r057003,r057004,r057005,r057006,r057007,r057008,r057009,r057010,r057011,r057012,r057014,r057013,r057001,r057002 FROM R020306,acc090,staff WHERE r057002=st01(+) and r057001=a0901(+) and ID='" & strUserNum & "' ORDER BY R057001,R057002,R057003"
strSql = "SELECT nvl(a0902,a0903),nvl(st02,r057002),r057003,r057004,r057005,r057006,r057007,r057008,r057009,r057010,r057011,r057012,r057014,r057013,r057001,r057002 FROM R020306,acc090,staff WHERE r057002=st01(+) and r057001=a0901(+) and ID='" & strUserNum & "' ORDER BY R057001,R057002,R057003,R057004"
CheckOC
Page = 1
SavDay1 = ""
SavDay2 = ""
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        SavDay1 = CheckStr(.Fields(0))
        SavDay2 = CheckStr(.Fields(1))
        SavDay(2) = "0"
        SavDay(0) = "0"
        PrintTitle
        Do While .EOF = False
'            For i = 0 To 11
            For i = 0 To 12
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If SavDay1 <> strTemp(0) Then
                PrintEnd
                Page = Page + 1
                Printer.NewPage
                SavDay1 = strTemp(0)
                SavDay2 = strTemp(1)
                PrintTitle
                SavDay(2) = "0"
                SavDay(0) = "0"
            Else
                If SavDay2 <> strTemp(1) Then
                    PrintEnd
                    Page = Page + 1
                    Printer.NewPage
                    SavDay2 = strTemp(1)
                    PrintTitle
                    SavDay(2) = "0"
                    SavDay(0) = "0"
                End If
            End If
            strTemp(4) = StrToStr(strTemp(4), 8)
            strTemp(5) = StrToStr(strTemp(5), 7)
            strTemp(6) = StrToStr(strTemp(6), 4)
            strTemp(7) = StrToStr(strTemp(7), 10)
            strTemp(8) = StrToStr(strTemp(8), 4)
            PrintDatil
            SavDay(0) = Trim(str(Val(SavDay(0)) + 1))
            'edit by nickc 2005/04/21
            'If Len(CheckStr(.Fields(12))) <> 0 Then
            If Len(CheckStr(.Fields(11))) <> 0 Then
                SavDay(2) = Trim(str(Val(SavDay(2)) + 1))
            End If
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
ShowPrintOk
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
Printer.Print "內商智權人員收文明細表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6200
Printer.CurrentY = iPrint
If txt1(4) = "1" Then
    Printer.Print "收文日：" & Format(ChangeTStringToTDateString(txt1(5)) & " ", "@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(6))
Else
    Printer.Print "取消收文日：" & Format(ChangeTStringToTDateString(txt1(5)) & " ", "@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(6))
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
Printer.Print "業務區：" & SavDay1
Printer.CurrentX = 4000
Printer.CurrentY = iPrint
Printer.Print "智權人員：" & SavDay2
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.Font.Size = 10
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
'If txt1(4) = "1" Then
   Printer.Print "收文日"
'Else
'   Printer.Print "取消收文日"
'End If
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
'Modify By Cheng 2002/02/19
'Printer.Print "商品類別"
Printer.Print "商品別"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "申請人"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
'Modify By Cheng 2002/02/19
'Printer.Print "承辦人員"
Printer.Print "承辦人"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "本所期限"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "發文日"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
'If txt1(4) = "1" Then
   Printer.Print "取消收文日"
'Else
'   Printer.Print "收文日"
'End If
'Add By Cheng 2002/02/19
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "申請國家"

'Printer.CurrentX = PLeft(12)
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "說明"
iPrint = iPrint + 300
Printer.Font.Size = 12
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.Font.Size = 10
End Sub

Sub PrintDatil()
For i = 2 To 12
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 1500
PLeft(2) = 500
PLeft(3) = 1500
PLeft(4) = 3300
PLeft(5) = 5000
'Modify By Cheng 2002/02/19
'PLeft(6) = 6500
PLeft(6) = 5800
'PLeft(7) = 7500
PLeft(7) = 6800
'PLeft(8) = 9700
PLeft(8) = 9000
'PLeft(9) = 10700
PLeft(9) = 9800
'PLeft(10) = 11700
PLeft(10) = 10880
'PLeft(11) = 12700
PLeft(11) = 11880
PLeft(12) = 13090
'PLeft(12) = 14000
PLeft(13) = 14200
End Sub

Sub PrintEnd()
SavDay(1) = Trim(str(Val(SavDay(0)) - Val(SavDay(2))))
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
'Modify By Cheng 2002/12/16
'Printer.Print "收文件數：" & SavDay(1)
'edit by nickc 2005/04/21
'Printer.Print "收文件數：" & IIf(Me.txt1(4).Text = "1", SavDay(2), SavDay(1))
Printer.Print "收文件數：" & IIf(Me.txt1(4).Text = "2", SavDay(2), SavDay(1))
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
'Modify By Cheng 2002/12/16
'Printer.Print "取消收文件數：" & SavDay(2)
'edit by nickc 2005/04/21
'Printer.Print "取消收文件數：" & IIf(Me.txt1(4).Text = "1", SavDay(1), SavDay(2))
Printer.Print "取消收文件數：" & IIf(Me.txt1(4).Text = "2", SavDay(1), SavDay(2))
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = GetSystemKindByNick
txt1(4) = "1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm020306 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CmdOK(0).SetFocus
End If
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
'Add By Cheng 2002/12/16
Select Case Index
Case 4 '列印別
    If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Select
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
Case 3
     lbl1 = GetPrjSalesNM(txt1(3))
     If Trim(txt1(3)) <> "" Then
        If Trim(lbl1.Caption) = "" Then
            s = MsgBox("智權人員輸入錯誤！", , "錯誤！")
            txt1(3).SetFocus
            txt1_GotFocus (3)
            Exit Sub
        End If
     End If
Case 4
     Select Case Trim(txt1(Index))
     Case "1", "2", ""
     Case Else
          s = MsgBox("列印別只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          txt1(4).SetFocus
          txt1(4).SelStart = 0
          txt1(4).SelLength = Len(txt1(4))
          Exit Sub
     End Select
Case 5, 6
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 6 Then
     If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
   End If
Case 2, 6, 8
     If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If

Case Else
End Select
End Sub


