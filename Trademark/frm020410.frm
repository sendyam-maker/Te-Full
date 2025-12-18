VERSION 5.00
Begin VB.Form frm020410 
   BorderStyle     =   1  '單線固定
   Caption         =   "申請案收/發件數月統計表"
   ClientHeight    =   2625
   ClientLeft      =   4740
   ClientTop       =   705
   ClientWidth     =   3870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3870
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   630
      Left            =   12
      TabIndex        =   11
      Top             =   1620
      Width           =   3825
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   5
         Top             =   180
         Width           =   2880
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   180
         Left            =   105
         TabIndex        =   12
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3084
      TabIndex        =   7
      Top             =   24
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2292
      TabIndex        =   6
      Top             =   24
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2250
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1290
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1230
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1290
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2250
      MaxLength       =   5
      TabIndex        =   2
      Top             =   930
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1230
      MaxLength       =   5
      TabIndex        =   1
      Top             =   930
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1230
      TabIndex        =   0
      Top             =   570
      Width           =   1956
   End
   Begin VB.Line Line2 
      X1              =   1770
      X2              =   2520
      Y1              =   1395
      Y2              =   1395
   End
   Begin VB.Line Line1 
      X1              =   1680
      X2              =   2940
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   3
      Left            =   270
      TabIndex        =   10
      Top             =   1320
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "年月範圍："
      Height          =   180
      Index           =   2
      Left            =   270
      TabIndex        =   9
      Top             =   975
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   270
      TabIndex        =   8
      Top             =   630
      Width           =   915
   End
End
Attribute VB_Name = "frm020410"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/10 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, SavDay4 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 6) As String, strTemp3 As String, TestOk As Boolean, StrTemp7(0 To 6) As String, StrSQL7 As String, strSQL8 As String
Dim PLeft(0 To 6) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, SeekPrint As Integer, SeekPrintL As Integer, StrSQL3 As String, k As Integer

Private Sub cmdok_Click(Index As Integer)
'Add By Cheng 2002/11/15
On Error GoTo ErrorHandler

Select Case Index
Case 0
     PUB_RestorePrinter Combo1 'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除
     Printer.EndDoc
     'Modified by Moran 2015/6/1
     'Printer.PaperSize = 39
     Printer.PaperSize = PUB_GetPaperSize(15, 2)
     'end 2015/6/1
     DoEvents
     If Len(txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         Exit Sub
     Else
         'Add By Cheng 2002/03/21
         If PUB_CheckKeyInYYMM(Me.txt1(1)) = -1 Then
            Me.txt1(1).SetFocus
            txt1_GotFocus 1
            Exit Sub
         End If
         If PUB_CheckKeyInYYMM(Me.txt1(2)) = -1 Then
            Me.txt1(2).SetFocus
            txt1_GotFocus 2
            Exit Sub
         End If

         If Len(txt1(2)) = 0 Then
             s = MsgBox("年月範圍區間不可空白!!", , "USER 輸入錯誤")
             txt1(1).SetFocus
             txt1_GotFocus (1)
             Exit Sub
         Else
             Screen.MousePointer = vbHourglass
             Me.Enabled = False
             TestOk = False
             Page = 1
             ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/19 清除查詢印表記錄檔欄位
             Printer.EndDoc 'Add By Sindy 2011/11/1
             Process
             'Process1
             Me.Enabled = True
             Screen.MousePointer = vbDefault
         End If
     End If
Case 1
    'Add By Cheng 2004/04/30
    '若印表機變動, 則更新列印設定
    If Me.Combo1.Text <> Me.Combo1.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
    End If
     Unload Me
Case Else
End Select
'Add By Cheng 2002/11/15
Exit Sub
ErrorHandler:
    Select Case Err.Number
    Case 380
        MsgBox "印表機選擇錯誤!!!"
    Case Else
        MsgBox "(" & Err.Number & ")" & Err.Description
    End Select
End Sub

Sub Process1()
cnnConnection.Execute "DELETE FROM R020410 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL6 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   StrSQL3 = StrSQL3 + " AND PE02 IN (" & GetAddStr(txt1(0)) & ") "
End If
StrSQL6 = ""
If Len(txt1(3)) <> 0 Then
    strSQL1 = strSQL1 + " AND TM10>='" & txt1(3) & "' "
    strSQL2 = strSQL2 + " AND SP09>='" & txt1(3) & "' "
End If
If Len(txt1(4)) <> 0 Then
    strSQL1 = strSQL1 + " AND TM10<='" & txt1(4) & "' "
    strSQL2 = strSQL2 + " AND SP09<='" & txt1(4) & "' "
End If
StrSQL6 = StrSQL6 + " AND CP26 IS NULL AND CP57 IS NULL "
CheckOC
StrSQL7 = ""
strSQL8 = ""
If Len(txt1(1)) <> 0 Then
StrSQL7 = StrSQL7 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(1) & "01")) - 10000 & " "
strSQL8 = strSQL8 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(1) & "01")) - 10000 & " "
End If
If Len(Trim(txt1(2))) <> 0 Then
   StrSQL7 = StrSQL7 + " AND CP05<=" & Val(ChangeTStringToWString(txt1(2) & "31")) - 10000 & " "
   strSQL8 = strSQL8 + " AND CP27<=" & Val(ChangeTStringToWString(txt1(2) & "31")) - 10000 & " "
End If
Dim tmpcp13 As String
Dim tmpPerformance As Integer
'Modify By Cheng 2002/04/10
'業務區應抓CP12
'strSQL = "SELECT NVL(A0902,A0903),CP13,'1','','','',ST03 FROM CASEPROGRESS,TRADEMARK,NATION,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP10='101'  AND TM10=NA01(+) AND ST03=A0901(+) AND CP13=ST01(+) " & strSQL1 & StrSQL6 & StrSQL7
'strSQL = strSQL + " UNION ALL SELECT NVL(A0902,A0903),CP13,'','','1','',ST03 FROM CASEPROGRESS,TRADEMARK,NATION,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP10='101' AND TM10=NA01(+) AND ST03=A0901(+) AND CP13=ST01(+) " & strSQL1 & StrSQL6 & StrSQL8
'strSQL = strSQL + " UNION ALL SELECT NVL(A0902,A0903),CP13,'1','','','',ST03 FROM CASEPROGRESS,SERVICEPRACTICE,NATION,ACC090,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP10='806' AND SP09=NA01(+) AND ST03=A0901(+) AND CP13=ST01(+) " & strSQL2 & StrSQL6 & StrSQL7
'strSQL = strSQL + " UNION ALL SELECT NVL(A0902,A0903),CP13,'','','1','',ST03 FROM CASEPROGRESS,SERVICEPRACTICE,NATION,ACC090,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP10='806' AND SP09=NA01(+) AND ST03=A0901(+) AND CP13=ST01(+) " & strSQL2 & StrSQL6 & StrSQL8
strSql = "SELECT NVL(A0902,A0903),CP13,'1','','','',CP12 FROM CASEPROGRESS,TRADEMARK,NATION,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP10='101'  AND TM10=NA01(+) AND CP12=A0901(+) AND CP13=ST01(+) " & strSQL1 & StrSQL6 & StrSQL7
strSql = strSql + " UNION ALL SELECT NVL(A0902,A0903),CP13,'','','1','',CP12 FROM CASEPROGRESS,TRADEMARK,NATION,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP10='101' AND TM10=NA01(+) AND CP12=A0901(+) AND CP13=ST01(+) " & strSQL1 & StrSQL6 & strSQL8
strSql = strSql + " UNION ALL SELECT NVL(A0902,A0903),CP13,'1','','','',CP12 FROM CASEPROGRESS,SERVICEPRACTICE,NATION,ACC090,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP10='806' AND SP09=NA01(+) AND CP12=A0901(+) AND CP13=ST01(+) " & strSQL2 & StrSQL6 & StrSQL7
strSql = strSql + " UNION ALL SELECT NVL(A0902,A0903),CP13,'','','1','',CP12 FROM CASEPROGRESS,SERVICEPRACTICE,NATION,ACC090,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP10='806' AND SP09=NA01(+) AND CP12=A0901(+) AND CP13=ST01(+) " & strSQL2 & StrSQL6 & strSQL8
strSql = strSql + " order by cp13 "
tmpcp13 = ""
tmpPerformance = 0
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        k = 0
        DoEvents
        Do While .EOF = False
            For i = 0 To 6
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If tmpcp13 <> strTemp(1) Then
               tmpcp13 = strTemp(1)
               'Modify By Chneg 2002/04/10
'               strTemp(1) = GetPerformanceByNick_1(Val("0000" & Format(DateAdd("Y", -1, ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & "01"))), "MM")), Val("0000" & Format(DateAdd("Y", -1, ChangeWStringToWDateString(ChangeTStringToWString(txt1(2) & "01"))), "MM")), txt1(0), strTemp(1))
               'Modify By Sindy 2010/10/15
               'strTemp(1) = GetPerformanceByNick_1(Val("0000" & Format(DateAdd("Y", -1, ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & "01"))), "MM")), Val("0000" & Format(DateAdd("Y", -1, ChangeWStringToWDateString(ChangeTStringToWString(txt1(2) & "01"))), "MM")), txt1(0), strTemp(6))
               strTemp(1) = GetPerformanceByNick_1(Val(Format(DateAdd("yyyy", -1, ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & "01"))), "YYYYMM")), Val(Format(DateAdd("yyyy", -1, ChangeWStringToWDateString(ChangeTStringToWString(txt1(2) & "01"))), "YYYYMM")), txt1(0), strTemp(6))
               '2010/10/15 End
               If Val(strTemp(1)) = 0 Then
                  tmpPerformance = 0
               Else
                  tmpPerformance = strTemp(1)
               End If
            Else
               strTemp(1) = tmpPerformance
            End If
            strSql = "INSERT INTO R020410 VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & ",''," & Val(strTemp(4)) & ",'','" & ChgSQL(strTemp(6)) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
            .MoveNext
            k = k + 1
            DoEvents
        Loop
    End If
End With
CheckOC
'strSQL = "select distinct r083001 from r020410 where id='" & strUserNum & "' "
'With adoRecordset
'   .CursorLocation = adUseClient
'   .Open strSQL, Connection, adOpenStatic, adLockReadOnly
'   If .EOF = False And .BOF = False Then
'
'      strTemp(1) = GetPerformanceByNick_1(Val("0000" & Format(DateAdd("Y", -1, ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & "01"))), "MM")), Val("0000" & Format(DateAdd("Y", -1, ChangeWStringToWDateString(ChangeTStringToWString(txt1(2) & "01"))), "MM")), txt1(0), strTemp(1))
'
'   End If
'End With
End Sub

Sub Process()
cnnConnection.Execute "DELETE FROM R020410 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL6 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   StrSQL3 = StrSQL3 + " AND PE02 IN (" & GetAddStr(txt1(0)) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/10/19
End If
StrSQL6 = ""
If Len(txt1(3)) <> 0 Then
    strSQL1 = strSQL1 + " AND TM10>='" & txt1(3) & "' "
    strSQL2 = strSQL2 + " AND SP09>='" & txt1(3) & "' "
End If
If Len(txt1(4)) <> 0 Then
    strSQL1 = strSQL1 + " AND TM10<='" & txt1(4) & "' "
    strSQL2 = strSQL2 + " AND SP09<='" & txt1(4) & "' "
End If
If Len(txt1(3)) <> 0 Or Len(txt1(4)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/10/19
End If
StrSQL6 = StrSQL6 + " AND CP26 IS NULL AND CP57 IS NULL "

StrSQL7 = ""
strSQL8 = ""
If Len(txt1(1)) <> 0 Then
StrSQL7 = StrSQL7 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(1) & "01")) & " "
strSQL8 = strSQL8 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(1) & "01")) & " "
End If
If Len(Trim(txt1(2))) <> 0 Then
   StrSQL7 = StrSQL7 + " AND CP05<=" & Val(ChangeTStringToWString(txt1(2) & "31")) & " "
   strSQL8 = strSQL8 + " AND CP27<=" & Val(ChangeTStringToWString(txt1(2) & "31")) & " "
End If
If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/10/19
End If
CheckOC
Dim tmpcp13 As String
Dim tmpPerformance As Integer
'Modify By Cheng 2002/04/08
'業務區應抓CP12
'strSQL = "SELECT NVL(A0902,A0903),CP13,'1','','','',ST03,cp09 FROM CASEPROGRESS,TRADEMARK,NATION,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP10='101'  AND TM10=NA01(+) AND ST03=A0901(+) AND CP13=ST01(+) " & strSQL1 & StrSQL6 & StrSQL7
'strSQL = strSQL + " UNION ALL SELECT NVL(A0902,A0903),CP13,'','','1','',ST03,cp09 FROM CASEPROGRESS,TRADEMARK,NATION,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP10='101'  AND TM10=NA01(+) AND ST03=A0901(+) AND CP13=ST01(+) " & strSQL1 & StrSQL6 & StrSQL8
'strSQL = strSQL + " UNION ALL SELECT NVL(A0902,A0903),CP13,'1','','','',ST03,cp09 FROM CASEPROGRESS,SERVICEPRACTICE,NATION,ACC090,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP10='806' AND SP09=NA01(+) AND ST03=A0901(+) AND CP13=ST01(+) " & strSQL2 & StrSQL6 & StrSQL7
'strSQL = strSQL + " UNION ALL SELECT NVL(A0902,A0903),CP13,'','','1','',ST03,cp09 FROM CASEPROGRESS,SERVICEPRACTICE,NATION,ACC090,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP10='806' AND SP09=NA01(+) AND ST03=A0901(+) AND CP13=ST01(+) " & strSQL2 & StrSQL6 & StrSQL8
'Modify By Cheng 2003/12/30
'加申請國家, 商品類別
'strSQL = "SELECT NVL(A0902,A0903),CP13,'1','','','',CP12,cp09 FROM CASEPROGRESS,TRADEMARK,NATION,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP10='101'  AND TM10=NA01(+) AND CP12=A0901(+) AND CP13=ST01(+) " & strSQL1 & StrSQL6 & strSQL7
'strSQL = strSQL + " UNION ALL SELECT NVL(A0902,A0903),CP13,'','','1','',CP12,cp09 FROM CASEPROGRESS,TRADEMARK,NATION,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP10='101'  AND TM10=NA01(+) AND CP12=A0901(+) AND CP13=ST01(+) " & strSQL1 & StrSQL6 & strSQL8
'strSQL = strSQL + " UNION ALL SELECT NVL(A0902,A0903),CP13,'1','','','',CP12,cp09 FROM CASEPROGRESS,SERVICEPRACTICE,NATION,ACC090,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP10='806' AND SP09=NA01(+) AND CP12=A0901(+) AND CP13=ST01(+) " & strSQL2 & StrSQL6 & strSQL7
'strSQL = strSQL + " UNION ALL SELECT NVL(A0902,A0903),CP13,'','','1','',CP12,cp09 FROM CASEPROGRESS,SERVICEPRACTICE,NATION,ACC090,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP10='806' AND SP09=NA01(+) AND CP12=A0901(+) AND CP13=ST01(+) " & strSQL2 & StrSQL6 & strSQL8
strSql = "SELECT NVL(A0902,A0903),CP13,'1','','','',CP12,cp09, TM10, TM09 FROM CASEPROGRESS,TRADEMARK,NATION,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP10='101'  AND TM10=NA01(+) AND CP12=A0901(+) AND CP13=ST01(+) " & strSQL1 & StrSQL6 & StrSQL7
strSql = strSql + " UNION ALL SELECT NVL(A0902,A0903),CP13,'','','1','',CP12,cp09, TM10, TM09 FROM CASEPROGRESS,TRADEMARK,NATION,ACC090,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP10='101'  AND TM10=NA01(+) AND CP12=A0901(+) AND CP13=ST01(+) " & strSQL1 & StrSQL6 & strSQL8
strSql = strSql + " UNION ALL SELECT NVL(A0902,A0903),CP13,'1','','','',CP12,cp09, '' As TM10, '' As TM09 FROM CASEPROGRESS,SERVICEPRACTICE,NATION,ACC090,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP10='806' AND SP09=NA01(+) AND CP12=A0901(+) AND CP13=ST01(+) " & strSQL2 & StrSQL6 & StrSQL7
strSql = strSql + " UNION ALL SELECT NVL(A0902,A0903),CP13,'','','1','',CP12,cp09, '' As TM10, '' As TM09 FROM CASEPROGRESS,SERVICEPRACTICE,NATION,ACC090,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP10='806' AND SP09=NA01(+) AND CP12=A0901(+) AND CP13=ST01(+) " & strSQL2 & StrSQL6 & strSQL8
'End
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    tmpcp13 = ""
    tmpPerformance = 0
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/19
        .MoveFirst
        k = 0
        DoEvents
        Do While .EOF = False
            For i = 0 To 6
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If tmpcp13 <> strTemp(1) Then
               tmpcp13 = strTemp(1)
               'Modify By Cheng 2002/04/10
'               strTemp(1) = GetPerformanceByNick_1(Val("0000" & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & "01")), "MM")), Val("0000" & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2) & "01")), "MM")), txt1(0), strTemp(1))
               'Modify By Sindy 2010/10/15
               'strTemp(1) = GetPerformanceByNick_1(Val("0000" & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & "01")), "MM")), Val("0000" & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2) & "01")), "MM")), txt1(0), strTemp(6))
               strTemp(1) = GetPerformanceByNick_1(Val(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & "01")), "YYYYMM")), Val(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2) & "01")), "YYYYMM")), txt1(0), strTemp(6))
               '2010/10/15 End
               If Val(strTemp(1)) = 0 Then
                  tmpPerformance = 0
               Else
                  tmpPerformance = strTemp(1)
               End If
            Else
               strTemp(1) = tmpPerformance
            End If
            If UCase(Mid(strTemp(6), 1, 1)) <> "S" Then
               strTemp(6) = "000"
               strTemp(0) = ""
            End If
            'Add By Cheng 2003/12/30
            '若申請國家為台灣
            If "" & .Fields("TM10").Value = "000" Then
                '收文件數
                If strTemp(2) = "1" Then
                    If "" & .Fields("TM09").Value <> "" Then
                        strTemp(2) = UBound(Split("" & .Fields("TM09").Value, ",")) + 1
                    End If
                '發文件數
                ElseIf strTemp(4) = "1" Then
                    If "" & .Fields("TM09").Value <> "" Then
                        strTemp(4) = UBound(Split("" & .Fields("TM09").Value, ",")) + 1
                    End If
                End If
            End If
            'End
            strSql = "INSERT INTO R020410 VALUES ('" & ChgSQL(strTemp(0)) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & ",''," & Val(strTemp(4)) & ",'','" & ChgSQL(strTemp(6)) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
            .MoveNext
            k = k + 1
            ''frm100.Tag = Trim(Str(.RecordCount)) & "=" & Trim(Str(k))
            ''frm100.StrMenu
            DoEvents
        Loop
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/10/19
        ShowNoData
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
End With
PrintData
Screen.MousePointer = vbDefault
End Sub

Sub PrintData()
strSql = "select r083001,max(r083002),sum(r083003),'',SUM(R083005),'',R083007 FROM R020410 WHERE ID='" & strUserNum & "' GROUP BY R083007,R083001 "
CheckOC
strTemp3 = ""
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        strTemp3 = CheckStr(.Fields(6))
        PrintTitle
        Do While .EOF = False
            For i = 0 To 6
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If Val(strTemp(2)) >= Val(strTemp(1)) Then
                strTemp(3) = "是"
            Else
                strTemp(3) = "否"
            End If
            If Val(strTemp(4)) >= Val(strTemp(1)) Then
                strTemp(5) = "是"
            Else
                strTemp(5) = "否"
            End If
            If StrToStr(strTemp3, 1) <> StrToStr(strTemp(6), 1) Then
                ShowLine
                If strTemp3 <> "000" Then
                  PrintEnd (0)
                   ShowLine
                End If
                'PrintTitle
                strTemp3 = strTemp(6)
            End If
            strTemp(0) = StrToStr(strTemp(0), 10)
            PrintDatil
            If iPrint >= 14000 Then
                ShowLine
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    Else
         'Add By Sindy 2011/3/1
         ShowNoData
         Exit Sub
         '2011/3/1 End
    End If
End With
ShowLine
PrintEnd (0)
ShowLine
PrintEnd (1)
ShowLine
Process1
PrintEnd (2)
ShowLine
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintEnd(Strindex As Integer)
Select Case Strindex
Case 0
     strSql = "select '各所小計',max(r083002),sum(r083003),'',sum(r083005),'' from r020410 where id='" & strUserNum & "' AND substr(r083007,1,2)='" & StrToStr(strTemp3, 1) & "' "
Case 1
     strSql = "select '全所總計',max(r083002),sum(r083003),'',sum(r083005),'' from r020410 where id='" & strUserNum & "' "
Case 2
     strSql = "select '去年全所總計',max(r083002),sum(r083003),'',sum(r083005),'' from r020410 where id='" & strUserNum & "' "
Case Else
     Exit Sub
End Select
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 5
                StrTemp7(i) = CheckStr(.Fields(i))
            Next i
            'Add By Cheng 2002/04/10
            '重新取得目標件數
            Select Case Strindex
            Case 0 '小計
               'Modify By Sindy 2010/10/15
               'StrTemp7(1) = GetPerformanceByNick_2(Val("0000" & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & "01")), "MM")), Val("0000" & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2) & "01")), "MM")), txt1(0), Left(strTemp3, 2))
               StrTemp7(1) = GetPerformanceByNick_2(Val(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & "01")), "YYYYMM")), Val(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2) & "01")), "YYYYMM")), txt1(0), Left(strTemp3, 2))
               '2010/10/15 End
            Case 1 '全所總計
               'Modify By Sindy 2010/10/15
               'StrTemp7(1) = GetPerformanceByNick_2(Val("0000" & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & "01")), "MM")), Val("0000" & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2) & "01")), "MM")), txt1(0), "")
               StrTemp7(1) = GetPerformanceByNick_2(Val(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & "01")), "YYYYMM")), Val(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2) & "01")), "YYYYMM")), txt1(0), "")
               '2010/10/15 End
            Case 2 '去年全所總計
               'Modify By Sindy 2010/10/15
               'StrTemp7(1) = GetPerformanceByNick_2(1, 12, txt1(0), "")
               StrTemp7(1) = GetPerformanceByNick_2(Val(Format(DateAdd("yyyy", -1, ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & "01"))), "YYYYMM")), Val(Format(DateAdd("yyyy", -1, ChangeWStringToWDateString(ChangeTStringToWString(txt1(2) & "01"))), "YYYYMM")), txt1(0), "")
               '2010/10/15 End
            End Select
            
            If Val(StrTemp7(2)) > Val(StrTemp7(1)) Then
                StrTemp7(3) = "是"
            Else
                StrTemp7(3) = "否"
            End If
            If Val(StrTemp7(4)) >= Val(StrTemp7(1)) Then
                StrTemp7(5) = "是"
            Else
                StrTemp7(5) = "否"
            End If
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
'            Printer.CurrentX = PLeft(1) + 300 - Printer.TextWidth(Format(StrTemp7(1), "####0"))
            Printer.CurrentX = PLeft(1) + 750 - Printer.TextWidth(Format(StrTemp7(1), "####0"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(1), "####0")
'            Printer.CurrentX = PLeft(2) + 300 - Printer.TextWidth(Format(StrTemp7(2), "####0"))
            Printer.CurrentX = PLeft(2) + 750 - Printer.TextWidth(Format(StrTemp7(2), "####0"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(2), "####0")
'            Printer.CurrentX = PLeft(3)
            Printer.CurrentX = PLeft(3) + 250
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(3)
'            Printer.CurrentX = PLeft(4) + 300 - Printer.TextWidth(Format(StrTemp7(4), "####0"))
            Printer.CurrentX = PLeft(4) + 750 - Printer.TextWidth(Format(StrTemp7(4), "####0"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(4), "####0")
'            Printer.CurrentX = PLeft(5)
            Printer.CurrentX = PLeft(5) + 250
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(5)
            iPrint = iPrint + 300
            If iPrint >= 14000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC2
End Sub


Sub PrintTitle()
GetPleft
iPrint = 0
'Printer.Orientation = 1 'Removed by Morgan 2015/6/3
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "申請案收/發件數月統計表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 7000
Printer.CurrentY = iPrint
If TestOk = False Then
    Printer.Print "日期：" & Format(ChangeTStringToTDateString(txt1(1) & "01"), "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2) & Format(DateAdd("D", -1, DateAdd("M", 1, ChangeWStringToWDateString(ChangeTStringToWString(txt1(2) & "01")))), "DD"))
Else
    Printer.Print "日期：" & Format(DateAdd("YYYY", -1, ChangeWStringToWDateString(ChangeTStringToWString(txt1(1) & "01"))), "@@@@@@@@@@") & "－" & DateAdd("YYYY", -1, ChangeWStringToWDateString(ChangeTStringToWString(txt1(2) & Format(DateAdd("D", -1, DateAdd("M", 1, ChangeWStringToWDateString(ChangeTStringToWString(txt1(2) & "01")))), "DD"))))
End If
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 16800
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "系統類別：" & txt1(0) & "    申請國家：" & Format(txt1(3), "@@@@@@@@@@") & "－" & txt1(4)
Printer.CurrentX = 16800
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "業務區別"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "目標件數"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "收文件數"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "達成否"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "發文件數"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "達成否"
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
End Sub

Sub PrintDatil()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(0)
'Printer.CurrentX = PLeft(1) + 300 - Printer.TextWidth(Format(strTemp(1), "####0"))
Printer.CurrentX = PLeft(1) + 750 - Printer.TextWidth(Format(strTemp(1), "####0"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(1), "####0")
'Printer.CurrentX = PLeft(2) + 300 - Printer.TextWidth(Format(strTemp(2), "####0"))
Printer.CurrentX = PLeft(2) + 750 - Printer.TextWidth(Format(strTemp(2), "####0"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(2), "####0")
'Printer.CurrentX = PLeft(3)
Printer.CurrentX = PLeft(3) + 250
Printer.CurrentY = iPrint
Printer.Print strTemp(3)
'Printer.CurrentX = PLeft(4) + 300 - Printer.TextWidth(Format(strTemp(4), "####0"))
Printer.CurrentX = PLeft(4) + 750 - Printer.TextWidth(Format(strTemp(4), "####0"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(4), "####0")
'Printer.CurrentX = PLeft(5)
Printer.CurrentX = PLeft(5) + 250
Printer.CurrentY = iPrint
Printer.Print strTemp(5)
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 0
PLeft(1) = 2000 + 3000
PLeft(2) = 4500 + 3000
PLeft(3) = 7000 + 3000
PLeft(4) = 9500 + 3000
PLeft(5) = 12000 + 3000
End Sub

Sub ShowLine()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
End If
End Sub

Private Sub Form_Load()

MoveFormToCenter Me
txt1(0) = GetSystemKindByNickT

SeekPrintL = Printer.Orientation
PUB_SetPrinter Me.Name, Combo1, , , SeekPrint     'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Printer = Printers(SeekPrint)
Printer.Orientation = SeekPrintL
Set frm020410 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub Txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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
     strTemp1 = Split(Replace(UCase(GetSystemKindByNickT), ",,", ""), ",")
     strTemp2 = Split(Replace(UCase(txt1(0)), ",,", ""), ",")
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
Case 4
     If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
Case 1, 2
   If PUB_CheckKeyInYYMM(Me.txt1(Index)) = -1 Then
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 2 Then
        If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
    End If
Case Else
End Select
End Sub


