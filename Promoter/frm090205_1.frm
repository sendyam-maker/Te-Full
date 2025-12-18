VERSION 5.00
Begin VB.Form frm090205_1 
   BorderStyle     =   1  '單線固定
   Caption         =   " 工作進度資料列印"
   ClientHeight    =   855
   ClientLeft      =   1950
   ClientTop       =   2985
   ClientWidth     =   2685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   2685
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   350
      Index           =   1
      Left            =   1548
      TabIndex        =   2
      Top             =   12
      Width           =   1092
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   756
      TabIndex        =   1
      Top             =   12
      Width           =   756
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Left            =   1740
      MaxLength       =   1
      TabIndex        =   0
      Text            =   "N"
      Top             =   468
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "是否含己發文資料：          (Y/N)"
      Height          =   180
      Left            =   36
      TabIndex        =   3
      Top             =   480
      Width           =   2580
   End
End
Attribute VB_Name = "frm090205_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim pemain As New ADODB.Recordset, k As Integer
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, StrSQL7 As String, StrSQL4 As String, strSQL5 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 21) As String, strTemp3 As String, TestOk As Boolean, StrTemp99(0 To 21) As String
Dim PLeft(0 To 21) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, Seekok As Integer, SeekTemp As Integer, Print1Ok As Boolean

Private Sub Form_Activate()
If ChkStaff(strUserNum) Then
   Unload Me
   Exit Sub
End If
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090205_1 = Nothing
End Sub

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
    If Len(Trim(Txt1)) = 0 Then
        s = MsgBox("是否含已發文資料不可空白!!", , "USER 輸入錯誤")
        Txt1.SetFocus
        txt1_GotFocus
        Exit Sub
    End If
    Printer.Orientation = 2
    DoEvents
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/14 清除查詢印表記錄檔欄位
    Process
    Process1
    Me.Enabled = True
    Screen.MousePointer = vbDefault
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub Process()
'Modify By Cheng 2003/05/08
'cnnConnection.Execute "DELETE FROM R090614 WHERE ID='" & strUserNum & "' "
adoEng.Execute "DELETE FROM R090614 WHERE ID='" & strUserNum & "' "
    'add by nickc 2007/12/17
    adoEng.Execute "drop table R090614 "
    adoEng.Execute "create table R090614 (R110001 text,R110002 double,R110003 text,R110004 text,R110005 text,R110006 text,R110007 text,R110008 text,R110009 text,R110010 text,R110011 text,R110012 text,R110013 text,R110014 text,R110015 text,R110016 text,R110017 text,R110018 text,R110019 double,R110020 memo,R110021 text,R110022 text,ID text,R110023 text, R110024 text,R110025 text)"
    'adoEng.Execute "create table R090614 (R110001 text,R110002 double,R110003 text,R110004 text,R110005 text,R110006 text,R110007 text,R110008 text,R110009 text,R110010 text,R110011 text,R110012 text,R110013 text,R110014 text,R110015 text,R110016 text,R110017 text,R110018 text,R110019 double,R110020 memo,R110021 text,R110022 text,ID text,R110023 text, R110024 text,R110025 text,R110026 double,R110027 double,R110028 double,R110029 text,R110030 text)"

StrSQL6 = ""
StrSQL6 = StrSQL6 + " AND ep05='" & strUserNum & "' "
If Txt1 = "Y" Then
   pub_QL05 = pub_QL05 & ";" & Left(Label1, 9) & Txt1 'Add By Sindy 2010/12/14
   StrSQL6 = StrSQL6 + " AND CP27>=" & Mid(GetTodayDate, 1, 6) & "01 and cp27<=" & Mid(GetTodayDate, 1, 6) & "31 "
End If
CheckOC
'未發文，未取消收文
                strSql = "SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND CP01 IN (" & SQLGrpStr("", 1) & ")  and cp27 is null  and cp57 is null  AND ep05='" & strUserNum & "' "
strSql = strSql + " UNION all  SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),decode(tm10,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+)  AND CP01 IN (" & SQLGrpStr("", 2) & ")   and cp27 is null  and cp57 is null  AND ep05='" & strUserNum & "' "
strSql = strSql + " UNION all  SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(LC05,NVL(LC06,LC07)),CP26,'',decode(lc15,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+)  AND CP01 IN (" & SQLGrpStr("", 3) & ")   and cp27 is null and cp57 is null  AND ep05='" & strUserNum & "' "
strSql = strSql + " UNION all  SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,HC06,CP26,'',CPM03," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND CP01 IN (" & SQLGrpStr("", 4) & ")   and cp27 is null and cp57 is null  AND ep05='" & strUserNum & "' "
strSql = strSql + " UNION all  SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),CP26,'',decode(sp09,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+)  AND CP01 IN (" & SQLGrpStr("", 5) & ")  and cp27 is null and cp57 is null  AND ep05='" & strUserNum & "' "
'未發文，當月取消收文
strSql = strSql + " UNION all  SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND CP01 IN (" & SQLGrpStr("", 1) & ")  and cp27 is null  and cp57>=" & Mid(GetTodayDate, 1, 6) & "01 and cp57<=" & Mid(GetTodayDate, 1, 6) & "31 AND ep05='" & strUserNum & "' "
strSql = strSql + " UNION all  SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),decode(tm10,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+)  AND CP01 IN (" & SQLGrpStr("", 2) & ")   and cp27 is null  and cp57>=" & Mid(GetTodayDate, 1, 6) & "01 and cp57<=" & Mid(GetTodayDate, 1, 6) & "31  AND ep05='" & strUserNum & "' "
strSql = strSql + " UNION all  SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(LC05,NVL(LC06,LC07)),CP26,'',decode(lc15,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+)  AND CP01 IN (" & SQLGrpStr("", 3) & ")   and cp27 is null and cp57>=" & Mid(GetTodayDate, 1, 6) & "01 and cp57<=" & Mid(GetTodayDate, 1, 6) & "31  AND ep05='" & strUserNum & "' "
strSql = strSql + " UNION all  SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,HC06,CP26,'',CPM03," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND CP01 IN (" & SQLGrpStr("", 4) & ")   and cp27 is null and cp57>=" & Mid(GetTodayDate, 1, 6) & "01 and cp57<=" & Mid(GetTodayDate, 1, 6) & "31  AND ep05='" & strUserNum & "' "
strSql = strSql + " UNION all  SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),CP26,'',decode(sp09,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+)  AND CP01 IN (" & SQLGrpStr("", 5) & ")  and cp27 is null and cp57>=" & Mid(GetTodayDate, 1, 6) & "01 and cp57<=" & Mid(GetTodayDate, 1, 6) & "31  AND ep05='" & strUserNum & "' "
'當月發文，未取消收文
If Txt1 = "Y" Then
'edit by nickc 2005/05/13
'   strSQL = strSQL & " UNION all  SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr("", 1) & ") and cp57 is null "
'   strSQL = strSQL + " UNION all  SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),decode(tm10,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr("", 2) & ")   and cp57 is null "
'   strSQL = strSQL + " UNION all  SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(LC05,NVL(LC06,LC07)),CP26,'',decode(lc15,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr("", 3) & ")   and cp57 is null "
'   strSQL = strSQL + " UNION all  SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,HC06,CP26,'',CPM03," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr("", 4) & ")   and cp57 is null "
'   strSQL = strSQL + " UNION all  SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),CP26,'',decode(sp09,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr("", 5) & ")  and cp57 is null "
   strSql = strSql & " UNION all  SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr("", 1) & ") "
   strSql = strSql + " UNION all  SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),decode(tm10,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr("", 2) & ")  "
   strSql = strSql + " UNION all  SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(LC05,NVL(LC06,LC07)),CP26,'',decode(lc15,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr("", 3) & ") "
   strSql = strSql + " UNION all  SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,HC06,CP26,'',CPM03," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr("", 4) & ") "
   strSql = strSql + " UNION all  SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),CP26,'',decode(sp09,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",NVL(S3.ST02,EP04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,NVL(S2.ST02,CP13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr("", 5) & ") "
End If
strSql = strSql + " ORDER BY 1 "
CheckOC
Print1Ok = False
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        DoEvents
        '判斷等級是否屬於專利
        'modify by sonia 2014/4/29 加EP05=94007
        If (Val(CheckStr(.Fields(22))) >= 31 And Val(CheckStr(.Fields(22))) <= 39) Or (Val(CheckStr(.Fields(22))) >= 71 And Val(CheckStr(.Fields(22))) <= 89) Or CheckStr(.Fields("EP05")) = "94007" Then
            Print1Ok = True
        End If
        Do While .EOF = False
            For i = 0 To 21
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            '計算承辦天數
            If Len(strTemp(14)) <> 0 And Len(strTemp(12)) <> 0 And Val(strTemp(14)) <> 0 And Val(strTemp(12)) <> 0 Then
                 strTemp(18) = Trim(str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(14))), ChangeTStringToWString(ChangeTDateStringToTString(strTemp(12))))))
            Else
                If Len(strTemp(13)) <> 0 And Len(strTemp(12)) <> 0 And Val(strTemp(13)) <> 0 And Val(strTemp(12)) <> 0 Then
                    strTemp(18) = Trim(str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(13))), ChangeTStringToWString(ChangeTDateStringToTString(strTemp(12))))))
                End If
            End If
            'Modify By Cheng 2003/05/08
'            strSQL = "INSERT INTO R090614 VALUES ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & strTemp(5) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "','" & strTemp(14) & "','" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "'," & Val(strTemp(18)) & ",'" & strTemp(19) & "','" & strTemp(20) & "','" & strTemp(21) & "','" & strUserNum & "','') "
            strSql = "INSERT INTO R090614 VALUES ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & strTemp(5) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "','" & strTemp(14) & "','" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "'," & Val(strTemp(18)) & ",'" & strTemp(19) & "','" & strTemp(20) & "','" & strTemp(21) & "','" & strUserNum & "','','','') "
'            cnnConnection.Execute strSQL
            adoEng.Execute strSql
            .MoveNext
            DoEvents
        Loop
    Else
        ShowNoData
        Exit Sub
    End If
    CheckOC
End With
'Modify By Cheng 2003/07/07
'CALCUTE_090201 strUserNum, Mid(GetTodayDate, 1, 6)
CALCUTE_090201 strUserNum, Mid(strSrvDate(1), 1, 6)
End Sub

Sub Process1()
    PrintData
End Sub

Sub PrintData()
If Print1Ok = True Then
    PrintData2   '專利
Else
    PrintData1   '一般
End If
End Sub

Sub PrintData1()
strSql = "SELECT DISTINCT R110001 FROM R090614 WHERE ID='" & strUserNum & "' "
CheckOC2
Page = 1
adoRecordset1.CursorLocation = adUseClient
'Modify By Cheng 2003/05/08
'adoRecordset1.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
adoRecordset1.Open strSql, adoEng, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
    InsertQueryLog (adoRecordset1.RecordCount) 'Add By Sindy 2010/12/14
    adoRecordset1.MoveFirst
    Do While adoRecordset1.EOF = False
        strTemp3 = CheckStr(adoRecordset1.Fields(0))
        PrintData1_1 (CheckStr(adoRecordset1.Fields(0)))
        PrintEnd1_1 (CheckStr(adoRecordset1.Fields(0)))
        adoRecordset1.MoveNext
        If adoRecordset1.EOF = False Then
            Page = Page + 1
            Printer.NewPage
        End If
    Loop
Else
    InsertQueryLog (0) 'Add By Sindy 2010/12/14
End If
CheckOC2
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintData1_1(Strindex As String)
If Len(Strindex) = 0 Then
    strSql = "SELECT * FROM R090614 WHERE ID='" & strUserNum & "' AND (R110001 IS NULL OR R110001='') order by r110002 "
Else
    strSql = "SELECT * FROM R090614 WHERE ID='" & strUserNum & "' AND R110001='" & Strindex & "'  order by r110002 "
End If
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    'Modify By Cheng 2003/05/08
'    .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
    .Open strSql, adoEng, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        PrintTitle_1
        Do While .EOF = False
            For i = 0 To 20
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(5) = StrToStr(strTemp(5), 10)
            strTemp(7) = StrToStr(strTemp(7), 3)
            strTemp(8) = StrToStr(strTemp(8), 4)
            strTemp(15) = StrToStr(strTemp(15), 3)
            strTemp(19) = StrToStr(strTemp(19), 5)
            strTemp(20) = StrToStr(strTemp(20), 3)
            If iPrint >= 9000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle_1
            End If
            PrintDatil_1
            .MoveNext
        Loop
    End If
End With
CheckOC
End Sub

Sub PrintData2_1(Strindex As String)
If Len(Strindex) = 0 Then
    strSql = "SELECT * FROM R090614 WHERE ID='" & strUserNum & "' AND (R110001 IS NULL OR R110001='') order by r110002 "
Else
    strSql = "SELECT * FROM R090614 WHERE ID='" & strUserNum & "' AND R110001='" & Strindex & "' order by r110002 "
End If
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    'Modify By Cheng 2003/05/08
'    .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
    .Open strSql, adoEng, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        PrintTitle_2
        Do While .EOF = False
            For i = 0 To 20
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(5) = StrToStr(strTemp(5), 10)
            strTemp(7) = StrToStr(strTemp(7), 3)
            strTemp(8) = StrToStr(strTemp(8), 4)
            strTemp(15) = StrToStr(strTemp(15), 3)
            strTemp(19) = StrToStr(strTemp(19), 5)
            strTemp(20) = StrToStr(strTemp(20), 3)
            If iPrint >= 9000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle_2
            End If
            PrintDatil_2
            .MoveNext
        Loop
    End If
End With
CheckOC
End Sub

Sub PrintData2()
strSql = "SELECT DISTINCT R110001 FROM R090614 WHERE ID='" & strUserNum & "' "
CheckOC2
Page = 1
adoRecordset1.CursorLocation = adUseClient
'Modify By Cheng 2003/05/08
'adoRecordset1.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
adoRecordset1.Open strSql, adoEng, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
    InsertQueryLog (adoRecordset1.RecordCount) 'Add By Sindy 2010/12/14
    adoRecordset1.MoveFirst
    Do While adoRecordset1.EOF = False
        strTemp3 = CheckStr(adoRecordset1.Fields(0))
        PrintData2_1 (CheckStr(adoRecordset1.Fields(0)))
        PrintEnd2_1 (CheckStr(adoRecordset1.Fields(0)))
        adoRecordset1.MoveNext
        If adoRecordset1.EOF = False Then
            Page = Page + 1
            Printer.NewPage
        End If
        Page = Page + 1
        Printer.NewPage
    Loop
Else
    InsertQueryLog (0) 'Add By Sindy 2010/12/14
End If
CheckOC2
Printer.EndDoc
End Sub

Sub PrintEnd1_1(Strindex As String)
'列印結尾
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_1
End If
If Len(Strindex) = 0 Then
    'Modify By Cheng 2003/05/08
'    strSQL = "SELECT SUM(DECODE(R111003,0,0,R111003)),SUM(DECODE(R111004,0,0,R111004)),SUM(DECODE(R111005,0,0,R111005)),SUM(DECODE(R111006,0,0,R111006)),SUM(DECODE(R111007,0,0,R111007)),SUM(DECODE(R111008,0,0,R111008)),SUM(DECODE(R111009,0,0,R111009)) FROM R090614_1 WHERE ID='" & strUserNum & "' AND (R111001 IS NULL OR R111001='') AND R111002='1' "
    strSql = "SELECT SUM(R111003),SUM(R111004),SUM(R111005),SUM(R111006),SUM(R111007),SUM(R111008),SUM(R111009) FROM R090614_1 WHERE ID='" & strUserNum & "' AND (R111001 IS NULL OR R111001='') AND R111002='1' "
Else
    'Modify By Cheng 2003/05/08
'    strSQL = "SELECT SUM(DECODE(R111003,0,0,R111003)),SUM(DECODE(R111004,0,0,R111004)),SUM(DECODE(R111005,0,0,R111005)),SUM(DECODE(R111006,0,0,R111006)),SUM(DECODE(R111007,0,0,R111007)),SUM(DECODE(R111008,0,0,R111008)),SUM(DECODE(R111009,0,0,R111009)) FROM R090614_1 WHERE ID='" & strUserNum & "' AND R111001='" & Strindex & "' AND R111002='1' "
    strSql = "SELECT SUM(R111003),SUM(R111004),SUM(R111005),SUM(R111006),SUM(R111007),SUM(R111008),SUM(R111009) FROM R090614_1 WHERE ID='" & strUserNum & "' AND R111001='" & Strindex & "' AND R111002='1' "
End If
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    'Modify By Cheng 2003/05/08
'    .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
    .Open strSql, adoEng, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        Printer.Print "本月收文件數：" & Format(CheckStr(.Fields(0)), "###,###,###,###,##0") & " 件"
        iPrint = iPrint + 300
        If iPrint >= 9000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle_1
        End If
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        Printer.Print "本月發文件數：" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0") & " 件, "
        Printer.CurrentX = 7000
        Printer.CurrentY = iPrint
        Printer.Print "點數：" & Format(CheckStr(.Fields(2)), "###,###,###,###,##0.00") & " 點"
        iPrint = iPrint + 300
        If iPrint >= 9000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle_1
        End If
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        Printer.Print "目前未完稿的件數：" & Format(CheckStr(.Fields(3)), "###,###,###,###,##0") & " 件"
        Printer.CurrentX = 7000
        Printer.CurrentY = iPrint
        Printer.Print "會稿中的件數：" & Format(CheckStr(.Fields(4)), "###,###,###,###,##0") & " 件"
        iPrint = iPrint + 300
        If iPrint >= 9000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle_1
        End If
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        Printer.Print "超過承辦期限之件數：" & Format(CheckStr(.Fields(5)), "###,###,###,###,##0") & " 件"
        Printer.CurrentX = 7000
        Printer.CurrentY = iPrint
        Printer.Print "當日法定期限之件數：" & Format(CheckStr(.Fields(6)), "###,###,###,###,##0") & " 件"
        iPrint = iPrint + 300
        If iPrint >= 9000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle_1
        End If
        ShowLine
    End If
End With
CheckOC
End Sub

Sub PrintEnd2_1(Strindex As String)
'列印結尾
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_2
End If
If Len(Strindex) = 0 Then
    'Modify By Cheng 2003/05/08
'    strSQL = "SELECT SUM(DECODE(R111010,0,0,R111010)),SUM(DECODE(R111011,0,0,R111011)),SUM(DECODE(R111012,0,0,R111012)),SUM(DECODE(R111013,0,0,R111013)),SUM(DECODE(R111009,0,0,R111009)),SUM(DECODE(R111004,0,0,R111004)),SUM(DECODE(R111005,0,0,R111005)),SUM(DECODE(R111008,0,0,R111008)) FROM R090614_1 WHERE ID='" & strUserNum & "' AND (R111001 IS NULL OR R111001='') AND R111002='2' "
    strSql = "SELECT SUM(R111010),SUM(R111011),SUM(R111012),SUM(R111013),SUM(R111009),SUM(R111004),SUM(R111005),SUM(R111008) FROM R090614_1 WHERE ID='" & strUserNum & "' AND (R111001 IS NULL OR R111001='') AND R111002='2' "
Else
    'Modify By Cheng 2003/05/08
'    strSQL = "SELECT SUM(DECODE(R111010,0,0,R111010)),SUM(DECODE(R111011,0,0,R111011)),SUM(DECODE(R111012,0,0,R111012)),SUM(DECODE(R111013,0,0,R111013)),SUM(DECODE(R111009,0,0,R111009)),SUM(DECODE(R111004,0,0,R111004)),SUM(DECODE(R111005,0,0,R111005)),SUM(DECODE(R111008,0,0,R111008)) FROM R090614_1 WHERE ID='" & strUserNum & "' AND R111001='" & Strindex & "' AND R111002='2' "
    strSql = "SELECT SUM(R111010),SUM(R111011),SUM(R111012),SUM(R111013),SUM(R111009),SUM(R111004),SUM(R111005),SUM(R111008) FROM R090614_1 WHERE ID='" & strUserNum & "' AND R111001='" & Strindex & "' AND R111002='2' "
End If
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    'Modify By Cheng 2003/05/08
'    .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
    .Open strSql, adoEng, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        Printer.Print "可辦非設計案件：" & Format(CheckStr(.Fields(0)), "###,###,###,###,##0") & " 件"
        Printer.CurrentX = 7000
        Printer.CurrentY = iPrint
        Printer.Print "可辦設計案件：" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0") & " 件"
        iPrint = iPrint + 300
        If iPrint >= 9000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle_2
        End If
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        Printer.Print "本月已完稿非設計件數：" & Format(CheckStr(.Fields(2)), "###,###,###,###,##0") & " 件"
        Printer.CurrentX = 7000
        Printer.CurrentY = iPrint
        Printer.Print "本月已完稿設計件數：" & Format(CheckStr(.Fields(3)), "###,###,###,###,##0") & " 件"
        iPrint = iPrint + 300
        If iPrint >= 9000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle_2
        End If
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        Printer.Print "當日法定期限之件數：" & Format(CheckStr(.Fields(4)), "###,###,###,###,##0") & " 件"
        Printer.CurrentX = 7000
        Printer.CurrentY = iPrint
        Printer.Print "本月發文件數：" & Format(CheckStr(.Fields(5)), "###,###,###,###,##0") & " 件, "
        iPrint = iPrint + 300
        If iPrint >= 9000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle_2
        End If
        Printer.CurrentX = 7000
        Printer.CurrentY = iPrint
        Printer.Print "本月發文點數：" & Format(CheckStr(.Fields(6)), "###,###,###,###,##0.00") & " 點"
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        Printer.Print "超過承辦期限之件數：" & Format(CheckStr(.Fields(7)), "###,###,###,###,##0") & " 件"
        iPrint = iPrint + 300
        If iPrint >= 9000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle_2
        End If
        ShowLine
    End If
End With
CheckOC
End Sub

Sub PrintTitle_1() '列印抬頭

iPrint = 0
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
Printer.Print "工作進度資料表(一般)"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
Printer.Print "發文年月：" & Format(ChangeTStringToTDateString(GetTaiwanTodayDate), "YY") & "/" & Format(ChangeTStringToTDateString(GetTaiwanTodayDate), "MM")
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "承辦人：" & GetPrjSalesNM(strTemp3)
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_1
    Exit Sub
End If
GetPleft_1
Printer.Font.Size = 9
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "目次"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "收文"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "收文日"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "Y/N"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "種類"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "承辦"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "本所"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "法定"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "齊備日"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "完稿日"
Printer.CurrentX = PLeft(14)
Printer.CurrentY = iPrint
Printer.Print "會稿日"
Printer.CurrentX = PLeft(15)
Printer.CurrentY = iPrint
Printer.Print "核稿人"
Printer.CurrentX = PLeft(16)
Printer.CurrentY = iPrint
Printer.Print "會稿"
Printer.CurrentX = PLeft(17)
Printer.CurrentY = iPrint
Printer.Print "發文日"
Printer.CurrentX = PLeft(18)
Printer.CurrentY = iPrint
Printer.Print "承辦"
Printer.CurrentX = PLeft(19)
Printer.CurrentY = iPrint
Printer.Print "備註"
Printer.CurrentX = PLeft(20)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_1
    Exit Sub
End If
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "類別"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "期限"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "期限"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "期限"
Printer.CurrentX = PLeft(16)
Printer.CurrentY = iPrint
Printer.Print "完成日"
Printer.CurrentX = PLeft(18)
Printer.CurrentY = iPrint
Printer.Print "天數"
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_1
    Exit Sub
End If
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_1
    Exit Sub
End If
Printer.Font.Size = 9
End Sub

Sub PrintTitle_2() '列印抬頭

iPrint = 0
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
Printer.Print "工作進度資料表(專利)"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
Printer.Print "發文年月：" & Format(ChangeTStringToTDateString(GetTaiwanTodayDate), "YY") & "/" & Format(ChangeTStringToTDateString(GetTaiwanTodayDate), "MM")
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "承辦人：" & GetPrjSalesNM(strTemp3)
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_2
    Exit Sub
End If
GetPleft_2
Printer.Font.Size = 9
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "目次"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "收文"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "收文日"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "Y/N"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "種類"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "承辦"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "本所"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "法定"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "齊備日"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "完稿日"
Printer.CurrentX = PLeft(14)
Printer.CurrentY = iPrint
Printer.Print "會稿日"
Printer.CurrentX = PLeft(15)
Printer.CurrentY = iPrint
Printer.Print "核稿人"
Printer.CurrentX = PLeft(16)
Printer.CurrentY = iPrint
Printer.Print "會稿"
Printer.CurrentX = PLeft(17)
Printer.CurrentY = iPrint
Printer.Print "發文日"
Printer.CurrentX = PLeft(18)
Printer.CurrentY = iPrint
Printer.Print "承辦"
Printer.CurrentX = PLeft(19)
Printer.CurrentY = iPrint
Printer.Print "備註"
Printer.CurrentX = PLeft(20)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_2
    Exit Sub
End If
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "類別"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "期限"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "期限"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "期限"
Printer.CurrentX = PLeft(16)
Printer.CurrentY = iPrint
Printer.Print "完成日"
Printer.CurrentX = PLeft(18)
Printer.CurrentY = iPrint
Printer.Print "天數"
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_2
    Exit Sub
End If
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_2
    Exit Sub
End If
Printer.Font.Size = 9
End Sub

Sub PrintDatil_1() '列印資料

For i = 1 To 20
    If i = 1 Or i = 18 Then
        Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(strTemp(i), "####0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(strTemp(i), "####0")
    Else
        Printer.CurrentX = PLeft(i)
        Printer.CurrentY = iPrint
        Printer.Print strTemp(i)
    End If
Next i
iPrint = iPrint + 300
End Sub

Sub PrintDatil_2() '列印資料

For i = 1 To 20
    If i = 1 Or i = 18 Then
        Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(strTemp(i), "####0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(strTemp(i), "####0")
    Else
        Printer.CurrentX = PLeft(i)
        Printer.CurrentY = iPrint
        Printer.Print strTemp(i)
    End If
Next i
iPrint = iPrint + 300
End Sub


Sub GetPleft_1()
'定陣列
'字 SIZE = 9
'1 WORD = 180 PIX
'0.5 WORD = 90 PIX
'SPACE = 90 PIX
Erase PLeft
PLeft(0) = 0
PLeft(1) = 0
PLeft(2) = PLeft(1) + (2.5 * 180)
PLeft(3) = PLeft(2) + (2.5 * 180)
PLeft(4) = PLeft(3) + (4.5 * 180)
PLeft(5) = PLeft(4) + (8 * 180)
PLeft(6) = PLeft(5) + (10.5 * 180)
PLeft(7) = PLeft(6) + (2 * 180)
PLeft(8) = PLeft(7) + (3.5 * 180)
PLeft(9) = PLeft(8) + (4.5 * 180)
PLeft(10) = PLeft(9) + (4.5 * 180)
PLeft(11) = PLeft(10) + (4.5 * 180)
PLeft(12) = PLeft(11) + (4.5 * 180)
PLeft(13) = PLeft(12) + (4.5 * 180)
PLeft(14) = PLeft(13) + (4.5 * 180)
PLeft(15) = PLeft(14) + (4.5 * 180)
PLeft(16) = PLeft(15) + (4.5 * 180)
PLeft(17) = PLeft(16) + (4.5 * 180)
PLeft(18) = PLeft(17) + (4.5 * 180)
PLeft(19) = PLeft(18) + (2.5 * 180)
PLeft(20) = PLeft(19) + (5.5 * 180)
End Sub

Sub GetPleft_2()
'定陣列
'字 SIZE = 9
'1 WORD = 180 PIX
'0.5 WORD = 90 PIX
'SPACE = 90 PIX
Erase PLeft
PLeft(0) = 0
PLeft(1) = 0
PLeft(2) = PLeft(1) + (2.5 * 180)
PLeft(3) = PLeft(2) + (2.5 * 180)
PLeft(4) = PLeft(3) + (4.5 * 180)
PLeft(5) = PLeft(4) + (8 * 180)
PLeft(6) = PLeft(5) + (10.5 * 180)
PLeft(7) = PLeft(6) + (2 * 180)
PLeft(8) = PLeft(7) + (3.5 * 180)
PLeft(9) = PLeft(8) + (4.5 * 180)
PLeft(10) = PLeft(9) + (4.5 * 180)
PLeft(11) = PLeft(10) + (4.5 * 180)
PLeft(12) = PLeft(11) + (4.5 * 180)
PLeft(13) = PLeft(12) + (4.5 * 180)
PLeft(14) = PLeft(13) + (4.5 * 180)
PLeft(15) = PLeft(14) + (4.5 * 180)
PLeft(16) = PLeft(15) + (4.5 * 180)
PLeft(17) = PLeft(16) + (4.5 * 180)
PLeft(18) = PLeft(17) + (4.5 * 180)
PLeft(19) = PLeft(18) + (2.5 * 180)
PLeft(20) = PLeft(19) + (5.5 * 180)
End Sub

Sub ShowLine()
Printer.Line (0, iPrint + 150)-(16500, iPrint + 150)
iPrint = iPrint + 300
End Sub

Private Sub txt1_GotFocus()
Txt1.SelStart = 0
Txt1.SelLength = Len(Txt1)

End Sub


Private Sub Txt1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdOK(0).SetFocus
End If
End Sub

Private Sub txt1_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus()
Select Case Trim(Txt1)
Case "Y", "N", ""
Case Else
     s = MsgBox("所別只能輸入 Y OR N !!", , "USER 輸入錯誤")
     Txt1.SetFocus
     Txt1.SelStart = 0
     Txt1.SelLength = Len(Txt1)
     Exit Sub
End Select
End Sub

