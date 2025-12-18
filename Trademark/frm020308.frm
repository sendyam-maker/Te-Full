VERSION 5.00
Begin VB.Form frm020308 
   BorderStyle     =   1  '單線固定
   Caption         =   "申請意見書案件明細表"
   ClientHeight    =   2010
   ClientLeft      =   3510
   ClientTop       =   2430
   ClientWidth     =   3390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   3390
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2592
      TabIndex        =   5
      Top             =   12
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1800
      TabIndex        =   4
      Top             =   12
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1020
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1470
      Width           =   255
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1020
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1125
      Width           =   255
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1860
      MaxLength       =   7
      TabIndex        =   1
      Top             =   690
      Width           =   975
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   744
      MaxLength       =   7
      TabIndex        =   0
      Top             =   690
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "(1.承辦人   2.業務區)"
      Height          =   180
      Index           =   4
      Left            =   1335
      TabIndex        =   10
      Top             =   1515
      Width           =   1755
   End
   Begin VB.Label Label1 
      Caption         =   "(1.發文明細   2.准/駁明細)"
      Height          =   180
      Index           =   3
      Left            =   1290
      TabIndex        =   9
      Top             =   1155
      Width           =   2040
   End
   Begin VB.Line Line1 
      X1              =   1530
      X2              =   2175
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "列印順序："
      Height          =   180
      Index           =   2
      Left            =   75
      TabIndex        =   8
      Top             =   1515
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "列印內容："
      Height          =   180
      Index           =   1
      Left            =   75
      TabIndex        =   7
      Top             =   1185
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "日期："
      Height          =   180
      Index           =   0
      Left            =   75
      TabIndex        =   6
      Top             =   750
      Width           =   825
   End
End
Attribute VB_Name = "frm020308"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/9 日期欄已修改
'2007/9/10 整理
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay(0 To 5) As String, SavDayT(0 To 5) As String, SavDay22(0 To 5) As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 15) As String, strTemp3 As String, TestOk As Boolean
Dim PLeft(0 To 13) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String
'Add By Cheng 2002/02/01
Dim m_strSaleZone As String '業務區
Dim m_strSales As String '智權人員
'Add By Cheng 2002/02/25
Dim m_strPromoter As String '承辦人

'            SavDay(I)        '智權人員(承辦人)計算
'            SavDayT(I)       '總計算
'            SavDay22(I)      '業務區計算
'0      案件數
'1      結果數
'2      核准數
'3      准      勝訴率
'4      駁      勝訴率
'5      點數
Private Sub cmdOK_Click(index As Integer)
   Select Case index
   Case 0
        Printer.Orientation = 2
        DoEvents
         'Add By Cheng 2002/03/21
         If PUB_CheckKeyInDate(Me.txt1(0)) = -1 Then
            Me.txt1(0).SetFocus
            txt1_GotFocus 0
            Exit Sub
         End If
         If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
            Me.txt1(1).SetFocus
            txt1_GotFocus 1
            Exit Sub
         End If
        
        If Len(txt1(1)) = 0 Then
            s = MsgBox("日期區間不可空白!!", , "USER 輸入錯誤")
            txt1(0).SetFocus
            txt1_GotFocus (0)
            Exit Sub
        Else
            If Len(txt1(2)) = 0 Then
                s = MsgBox("列印內容不可空白!!", , "USER 輸入錯誤")
                txt1(2).SetFocus
                Exit Sub
            Else
                If Len(txt1(3)) = 0 Then
                    s = MsgBox("列印順序不可空白!!", , "USER 輸入錯誤")
                    txt1(3).SetFocus
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
Dim StrSQLa As String
   
   Screen.MousePointer = vbHourglass
   cnnConnection.Execute "DELETE FROM R020308_1 WHERE ID='" & strUserNum & "' "
   cnnConnection.Execute "DELETE FROM R020308_2 WHERE ID='" & strUserNum & "' "
   strSQL1 = ""
   strSQL2 = ""
   StrSQL6 = ""
   If Len(txt1(0)) <> 0 Then
      strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr("", 2) & ") "
      strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr("", 5) & ") "
   End If
   StrSQL6 = ""
   Select Case Val(txt1(2))
   Case 1 '發文明細
        pub_QL05 = pub_QL05 & ";" & Label1(1) & "發文明細"  'Add By Sindy 2010/10/4
        If Len(txt1(0)) <> 0 Then
            StrSQL6 = StrSQL6 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(0))) & ""
         End If
         If Len(Trim(txt1(1))) <> 0 Then
            StrSQL6 = StrSQL6 + " AND CP27<=" & Val(ChangeTStringToWString(txt1(1))) & " "
         End If
         If Len(txt1(0)) <> 0 Or Len(Trim(txt1(1))) <> 0 Then
            pub_QL05 = pub_QL05 & ";發文" & Label1(0) & txt1(0) & "-" & txt1(1)  'Add By Sindy 2010/10/4
         End If
   Case 2 '准駁明細
        pub_QL05 = pub_QL05 & ";" & Label1(1) & "准/駁明細"  'Add By Sindy 2010/10/4
        If Len(txt1(0)) <> 0 Then
            'Modify By Cheng 2002/04/11 改用商標基本檔的公告日TM14
            '2008/11/20 MODIFY BY SONIA 林副理說應改回審定來函日TM13
            StrSQL6 = StrSQL6 + " AND TM13>=" & Val(ChangeTStringToWString(txt1(0))) & ""
            'StrSQL6 = StrSQL6 + " AND TM14>=" & Val(ChangeTStringToWString(txt1(0))) & ""
        End If
        If Len(Trim(txt1(1))) <> 0 Then
            'Modify By Cheng 2002/04/11 改用商標基本檔的公告日TM14
            '2008/11/20 MODIFY BY SONIA 林副理說應改回審定來函日TM13
            StrSQL6 = StrSQL6 + " AND TM13<=" & Val(ChangeTStringToWString(txt1(1))) & " "
            'StrSQL6 = StrSQL6 + " AND TM14<=" & Val(ChangeTStringToWString(txt1(1))) & " "
        End If
        If Len(txt1(0)) <> 0 Or Len(Trim(txt1(1))) <> 0 Then
            pub_QL05 = pub_QL05 & ";審定來函" & Label1(0) & txt1(0) & "-" & txt1(1)  'Add By Sindy 2010/10/4
        End If
   Case Else
   End Select
   Select Case Val(txt1(3)) '列印順序
   Case 1 '承辦人
        pub_QL05 = pub_QL05 & ";" & Label1(2) & "承辦人"  'Add By Sindy 2010/10/4
        Select Case Val(txt1(2))
        '2011/12/8 modify by sonia 加210陳述意見,案號前加'N'
        Case 1 '發文明細
            'Modify By Cheng 2002/04/08
   '          strSQL = "SELECT NVL(S1.ST01,CP14)," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),DECODE(CP24,'1','准','2','駁',NULL,' '),NVL(A0902,A0903),NVL(S2.ST02,CP13),CP18,DECODE(CP25,NULL,' ','*'),'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND S2.ST03=A0901(+) AND CP10='202' " & IIf(intPWhere = 國內, " AND (SUBSTR(S1.ST03,1,2)='P2' OR CP14 IS NULL) ", " AND (SUBSTR(S1.ST03,1,2)='F1' OR CP14 IS NULL) ") & strSQL1 & StrSQL6
   '          strSQL = strSQL + " union all select NVL(S1.ST01,CP14)," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),DECODE(CP24,'1','准','2','駁',NULL,' '),NVL(A0902,A0903),NVL(S2.ST02,CP13),CP18,DECODE(CP25,NULL,' ','*'),'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND S2.ST03=A0901(+) AND CP10='202' " & IIf(intPWhere = 國內, " AND (SUBSTR(S1.ST03,1,2)='P2' OR CP14 IS NULL) ", " AND (SUBSTR(S1.ST03,1,2)='F1' OR CP14 IS NULL) ") & strSQL2 & StrSQL6
             strSql = "SELECT NVL(S1.ST01,CP14)," & SQLDate("CP27") & ",DECODE(TM28,'1',NULL,'N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),DECODE(TM16,'1','准','2','駁',NULL,' '),NVL(A0902,A0903),NVL(S2.ST02,CP13),CP18,DECODE(CP25,NULL,' ','*'),'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND cp12=A0901(+) AND CP10 in ('202','210') " & IIf(intPWhere = 國內, " AND (SUBSTR(S1.ST03,1,2)='P2' OR CP14 IS NULL) ", " AND (SUBSTR(S1.ST03,1,2)='F1' OR CP14 IS NULL) ") & strSQL1 & StrSQL6
        Case 2 '准駁明細
            'Add By Sindy 2013/1/9 若非商申收發文時, 若為TF案則不抓後三碼為"000"的資料
            StrSQLa = " SELECT CP43 FROM CASEPROGRESS,TRADEMARK " & _
                      "WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP27 IS NOT NULL AND (CP57 IS NULL OR CP57='') " & _
                      "AND CP03 <> Decode(CP01,'TF','0','z') " & _
                      "AND CP04 <> Decode(CP01,'TF','00','zz') " & _
                      "And (CP10='1003' OR CP10='1004') " & _
                      "AND CP05>=" & DBDATE(txt1(0)) & " AND CP05<=" & DBDATE(txt1(1)) & " "
            strSql = "SELECT NVL(S1.ST01,CP14)," & SQLDate("cp25") & ",DECODE(TM28,'1',NULL,'N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),DECODE(CP24,'1','准','2','駁',NULL,' '),NVL(A0902,A0903),NVL(S2.ST02,CP13),0,DECODE(CP25,NULL,' ','*'),'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND cp12=A0901(+) AND CP10 in ('202','210') " & IIf(intPWhere = 國內, " AND (SUBSTR(S1.ST03,1,2)='P2' OR CP14 IS NULL) ", " AND (SUBSTR(S1.ST03,1,2)='F1' OR CP14 IS NULL) ") & strSQL1 & " AND CP09 IN ( " & StrSQLa & " ) "
            '2013/1/9 End
            'Add By Cheng 2002/04/09 先取得符合條件的收文號且必須要有發文日
            'strSQL = "SELECT C2.CP09 FROM (SELECT CP01,CP02,CP03,CP04,MAX(CP27) FROM CASEPROGRESS,TRADEMARK WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND '202'=CP10 AND CP27 IS NOT NULL " & StrSQL6 & " GROUP BY CP01,CP02,CP03,CP04 ) C1,CASEPROGRESS C2 WHERE C1.CP01=C2.CP01(+) AND C1.CP02=C2.CP02(+) AND C1.CP03=C2.CP03(+) AND C1.CP04=C2.CP04(+) AND '202'=C2.CP10 "
            '2007/9/11 MODIFY BY SONIA 同一案號若有二筆202則會都抓出來
            'strSQL = "SELECT C2.CP09 FROM (SELECT CP01,CP02,CP03,CP04,MAX(CP27) MAXCP27 FROM CASEPROGRESS,TRADEMARK WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND '202'=CP10 AND CP27 IS NOT NULL " & StrSQL6 & " GROUP BY CP01,CP02,CP03,CP04 ) C1,CASEPROGRESS C2 WHERE C1.CP01=C2.CP01(+) AND C1.CP02=C2.CP02(+) AND C1.CP03=C2.CP03(+) AND C1.CP04=C2.CP04(+) AND '202'=C2.CP10(+) "
            StrSQLa = "SELECT C2.CP09 FROM (SELECT CP01,CP02,CP03,CP04,MAX(CP27) MAXCP27 FROM CASEPROGRESS,TRADEMARK WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP10 in ('202','210') AND CP27 IS NOT NULL AND (CP57 IS NULL OR CP57='') " & StrSQL6 & " GROUP BY CP01,CP02,CP03,CP04 ) C1,CASEPROGRESS C2 WHERE C1.CP01=C2.CP01(+) AND C1.CP02=C2.CP02(+) AND C1.CP03=C2.CP03(+) AND C1.CP04=C2.CP04(+) AND C2.CP10 in('202','210') AND C1.MAXCP27=C2.CP27(+) "
            '2008/11/20 MODIFY BY SONIA 林副理說應改回審定來函日TM13
            'strSQL = "SELECT NVL(S1.ST01,CP14)," & SQLDate("TM14") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),DECODE(TM16,'1','准','2','駁',NULL,' '),NVL(A0902,A0903),NVL(S2.ST02,CP13),CP18,DECODE(CP25,NULL,' ','*'),'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND cp12=A0901(+) AND CP10='202' " & IIf(intPWhere = 國內, " AND (SUBSTR(S1.ST03,1,2)='P2' OR CP14 IS NULL) ", " AND (SUBSTR(S1.ST03,1,2)='F1' OR CP14 IS NULL) ") & strSQL1 & StrSQL6 & " AND CP09 IN ( " & strSQL & " ) "
            strSql = strSql & "union SELECT NVL(S1.ST01,CP14)," & SQLDate("TM13") & ",DECODE(TM28,'1',NULL,'N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),DECODE(TM16,'1','准','2','駁',NULL,' '),NVL(A0902,A0903),NVL(S2.ST02,CP13),CP18,DECODE(CP25,NULL,' ','*'),'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND cp12=A0901(+) AND CP10 in ('202','210') " & IIf(intPWhere = 國內, " AND (SUBSTR(S1.ST03,1,2)='P2' OR CP14 IS NULL) ", " AND (SUBSTR(S1.ST03,1,2)='F1' OR CP14 IS NULL) ") & strSQL1 & StrSQL6 & " AND CP09 IN ( " & StrSQLa & " ) "
        Case Else
        End Select
        cnnConnection.Execute "INSERT INTO R020308_2 " & strSql
        strSql = "SELECT * FROM R020308_2 WHERE ID='" & strUserNum & "' "
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
        PrintData2
   Case 2 '業務區
        pub_QL05 = pub_QL05 & ";" & Label1(2) & "業務區"  'Add By Sindy 2010/10/4
        Select Case Val(txt1(2))
        '2011/12/8 modify by sonia 加210陳述意見,案號前加'N'
        Case 1 '發文明細
            'Modify By Cheng 2002/04/08
            'strSQL = "SELECT NVL(s2.st03,CP13),CP13," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),DECODE(CP24,'1','准','2','駁',NULL,' '),NVL(S1.ST02,CP14),CP18,DECODE(CP25,NULL,' ','*'),'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)  AND CP10='202' " & IIf(intPWhere = 國內, " AND (SUBSTR(S1.ST03,1,2)='P2' OR CP14 IS NULL) ", " AND (SUBSTR(S1.ST03,1,2)='F1' OR CP14 IS NULL) ") & strSQL1 & StrSQL6
            'strSQL = strSQL + " union all select NVL(s2.st03,CP13),CP13," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),DECODE(CP24,'1','准','2','駁',NULL,' '),NVL(S1.ST02,CP14),CP18,DECODE(CP25,NULL,' ','*'),'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP10='202' " & IIf(intPWhere = 國內, " AND (SUBSTR(S1.ST03,1,2)='P2' OR CP14 IS NULL) ", " AND (SUBSTR(S1.ST03,1,2)='F1' OR CP14 IS NULL) ") & strSQL2 & StrSQL6
            strSql = "SELECT cp12,CP13," & SQLDate("CP27") & ",DECODE(TM28,'1',NULL,'N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),DECODE(TM16,'1','准','2','駁',NULL,' '),NVL(S1.ST02,CP14),CP18,DECODE(CP25,NULL,' ','*'),'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)  AND CP10 in ('202','210') " & IIf(intPWhere = 國內, " AND (SUBSTR(S1.ST03,1,2)='P2' OR CP14 IS NULL) ", " AND (SUBSTR(S1.ST03,1,2)='F1' OR CP14 IS NULL) ") & strSQL1 & StrSQL6
        Case 2 '准駁明細
            'Add By Sindy 2013/1/9 若非商申收發文時, 若為TF案則不抓後三碼為"000"的資料
            StrSQLa = " SELECT CP43 FROM CASEPROGRESS,TRADEMARK " & _
                      "WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP27 IS NOT NULL AND (CP57 IS NULL OR CP57='') " & _
                      "AND CP03 <> Decode(CP01,'TF','0','z') " & _
                      "AND CP04 <> Decode(CP01,'TF','00','zz') " & _
                      "And (CP10='1003' OR CP10='1004') " & _
                      "AND CP05>=" & DBDATE(txt1(0)) & " AND CP05<=" & DBDATE(txt1(1)) & " "
            strSql = "SELECT cp12,CP13," & SQLDate("cp25") & ",DECODE(TM28,'1',NULL,'N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),DECODE(CP24,'1','准','2','駁',NULL,' '),NVL(S1.ST02,CP14),0,DECODE(CP25,NULL,' ','*'),'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP10 in ('202','210') " & IIf(intPWhere = 國內, " AND (SUBSTR(S1.ST03,1,2)='P2' OR CP14 IS NULL) ", " AND (SUBSTR(S1.ST03,1,2)='F1' OR CP14 IS NULL) ") & strSQL1 & " AND CP09 IN ( " & StrSQLa & " ) "
            '2013/1/9 End
            'Add By Cheng 2002/04/09 先取得符合條件的收文號且必須要有發文日
            'strSQL = "SELECT C2.CP09 FROM (SELECT CP01,CP02,CP03,CP04,MAX(CP27) FROM CASEPROGRESS,TRADEMARK WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND '202'=CP10 " & StrSQL6 & " GROUP BY CP01,CP02,CP03,CP04 ) C1,CASEPROGRESS C2 WHERE C1.CP01=C2.CP01(+) AND C1.CP02=C2.CP02(+) AND C1.CP03=C2.CP03(+) AND C1.CP04=C2.CP04(+) AND '202'=C2.CP10 "
            '2007/9/11 MODIFY BY SONIA 同一案號若有二筆202則會都抓出來
            'strSQL = "SELECT C2.CP09 FROM (SELECT CP01,CP02,CP03,CP04,MAX(CP27) FROM CASEPROGRESS,TRADEMARK WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND '202'=CP10 AND CP27 IS NOT NULL " & StrSQL6 & " GROUP BY CP01,CP02,CP03,CP04 ) C1,CASEPROGRESS C2 WHERE C1.CP01=C2.CP01(+) AND C1.CP02=C2.CP02(+) AND C1.CP03=C2.CP03(+) AND C1.CP04=C2.CP04(+) AND '202'=C2.CP10(+) "
            StrSQLa = "SELECT C2.CP09 FROM (SELECT CP01,CP02,CP03,CP04,MAX(CP27) MAXCP27 FROM CASEPROGRESS,TRADEMARK WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP10 in ('202','210') AND CP27 IS NOT NULL AND (CP57 IS NULL OR CP57='') " & StrSQL6 & " GROUP BY CP01,CP02,CP03,CP04 ) C1,CASEPROGRESS C2 WHERE C1.CP01=C2.CP01(+) AND C1.CP02=C2.CP02(+) AND C1.CP03=C2.CP03(+) AND C1.CP04=C2.CP04(+) AND C2.CP10 in('202','210') AND C1.MAXCP27=C2.CP27(+) "
            '2008/11/20 MODIFY BY SONIA 林副理說應改回審定來函日TM13
            'strSQL = "SELECT cp12,CP13," & SQLDate("TM14") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),DECODE(TM16,'1','准','2','駁',NULL,' '),NVL(S1.ST02,CP14),CP18,DECODE(CP25,NULL,' ','*'),'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP10='202' " & IIf(intPWhere = 國內, " AND (SUBSTR(S1.ST03,1,2)='P2' OR CP14 IS NULL) ", " AND (SUBSTR(S1.ST03,1,2)='F1' OR CP14 IS NULL) ") & strSQL1 & StrSQL6 & " AND CP09 IN ( " & strSQL & " ) "
            strSql = strSql & "union SELECT cp12,CP13," & SQLDate("TM13") & ",DECODE(TM28,'1',NULL,'N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),DECODE(TM16,'1','准','2','駁',NULL,' '),NVL(S1.ST02,CP14),CP18,DECODE(CP25,NULL,' ','*'),'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP10 in ('202','210') " & IIf(intPWhere = 國內, " AND (SUBSTR(S1.ST03,1,2)='P2' OR CP14 IS NULL) ", " AND (SUBSTR(S1.ST03,1,2)='F1' OR CP14 IS NULL) ") & strSQL1 & StrSQL6 & " AND CP09 IN ( " & StrSQLa & " ) "
        Case Else
        End Select
        cnnConnection.Execute "INSERT INTO R020308_1 " & strSql
        CheckOC
        strSql = "SELECT * FROM R020308_1 WHERE ID='" & strUserNum & "'  "
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
        PrintData1
   Case Else
   End Select
   
   Screen.MousePointer = vbDefault
End Sub

'依業務區列印
Sub PrintData1()
   'Add By Cheng 2002/02/01
   m_strSaleZone = ""
   m_strSales = ""
   
   strSql = "SELECT nvl(a0902,a0903),st02,r059003,r059004,r059005,r059006,r059007,r059008,r059009,r059001,r059002 FROM R020308_1,acc090,staff WHERE r059001=A0901(+) and r059002=ST01(+) and  ID='" & strUserNum & "' order by r059001,r059002,r059003,r059004"
   CheckOC
   Page = 1
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
           .MoveFirst
           PrintTitle
           PrintTitle1
           '記錄業務區名稱
           SavDay1 = CheckStr(.Fields(0))
           '記錄智權人員名稱
           SavDay2 = CheckStr(.Fields(1))
           For i = 0 To 5
               SavDay(i) = "0"       '智權人員計算
               SavDayT(i) = "0"      '總計算
               SavDay22(i) = "0"     '業務區計算
           Next i
           Do While .EOF = False
               For i = 0 To 8
                   strTemp(i) = CheckStr(.Fields(i))
               Next i
               'Add By Sindy 2013/1/9
               If strTemp(7) = "0" Then
                  strTemp(7) = ""
               End If
               '2013/1/9 End
               If SavDay1 <> strTemp(0) Then
                  'Modify By Cheng 2002/04/11
   '                PrintEnd1 (0)
                   PrintEnd1 (1)
                   SavDay1 = strTemp(0)
                   SavDay2 = strTemp(1)
                   For i = 0 To 5
                       SavDay(i) = "0"
                       SavDay22(i) = "0"
                   Next i
                   iPrint = iPrint + 300
                   If iPrint >= 10000 Then
                       Page = Page + 1
                       Printer.NewPage
                       PrintTitle
                        'Modify By Cheng 2002/04/11
   '                    PrintTitle1
                   End If
                   PrintTitle1
               Else
                   If SavDay2 <> strTemp(1) Then
                     'Modify By Cheng 2002/04/11
   '                    PrintEnd1 (0)
                       SavDay2 = strTemp(1)
                       For i = 0 To 5
                           SavDay(i) = "0"
                       Next i
                       'Modify By Cheng 2002/04/11
   '                    iPrint = iPrint + 300
                       If iPrint >= 10000 Then
                           Page = Page + 1
                           Printer.NewPage
                           PrintTitle
                           PrintTitle1
                       End If
                       'Modify By Cheng 2002/04/11
   '                    PrintTitle1
                   End If
               End If
               SavDay(0) = Trim(str(Val(SavDay(0)) + 1))
               SavDayT(0) = Trim(str(Val(SavDayT(0)) + 1))
               SavDay22(0) = Trim(str(Val(SavDay22(0)) + 1))
               'Modify By Cheng 2002/04/11
   '            If Len(CheckStr(.Fields(9))) <> 0 Then
               If Len(CheckStr(strTemp(5))) <> 0 Then
                   SavDay(1) = Trim(str(Val(SavDay(1)) + 1))
                   SavDayT(1) = Trim(str(Val(SavDayT(1)) + 1))
                   SavDay22(1) = Trim(str(Val(SavDay22(1)) + 1))
               End If
               Select Case strTemp(5)
               Case "准"
                   SavDay(2) = Trim(str(Val(SavDay(2)) + 1))
                   SavDayT(2) = Trim(str(Val(SavDayT(2)) + 1))
                   SavDay22(2) = Trim(str(Val(SavDay22(2)) + 1))
               Case Else
               End Select
               SavDay(5) = Trim(str(Val(SavDay(5)) + Val(strTemp(7))))
               SavDayT(5) = Trim(str(Val(SavDayT(5)) + Val(strTemp(7))))
               SavDay22(5) = Trim(str(Val(SavDay22(5)) + Val(strTemp(7))))
               strTemp(0) = StrToStr(strTemp(0), 4)
               strTemp(1) = StrToStr(strTemp(1), 4)
               'Modify By Cheng 2002/04/08
   '            strTemp(4) = StrToStr(strTemp(4), 4)
               strTemp(4) = StrToStr(strTemp(4), 20)
               strTemp(6) = StrToStr(strTemp(6), 4)
               PrintDatil1
               If iPrint >= 10000 Then
                   Page = Page + 1
                   Printer.NewPage
                   PrintTitle
                   PrintTitle1
               End If
               .MoveNext
           Loop
       End If
   End With
   CheckOC
   'Modify By Cheng 2002/04/11
   'PrintEnd1 (0)
   PrintEnd1 (1)
   PrintEnd1 (2)
   Printer.EndDoc
   ShowPrintOk
End Sub

'依承辦人順序
Sub PrintData2()
   'Add By Cheng 2002/02/01
   m_strSaleZone = ""
   m_strSales = ""
   'Add By Cheng 2002/02/25
   m_strPromoter = ""
   
   strSql = "SELECT st02,r060002,r060003,r060004,r060005,r060006,r060007,r060008,r060009,r060001 FROM R020308_2,staff WHERE r060001=ST01(+) and ID='" & strUserNum & "' order by r060001,r060002,r060003"
   CheckOC
   Page = 1
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
           .MoveFirst
           PrintTitle
           PrintTitle2
           SavDay1 = CheckStr(.Fields(0))
           For i = 0 To 5
               SavDay(i) = "0"
               SavDayT(i) = "0"
           Next i
           Do While .EOF = False
               For i = 0 To 8
                   strTemp(i) = CheckStr(.Fields(i))
               Next i
               'Add By Sindy 2013/1/9
               If strTemp(7) = "0" Then
                  strTemp(7) = ""
               End If
               '2013/1/9 End
               If SavDay1 <> strTemp(0) Then
                   PrintEnd2 (0)
                   SavDay1 = strTemp(0)
                   For i = 0 To 5
                       SavDay(i) = "0"
                   Next i
                   iPrint = iPrint + 300
                   If iPrint >= 10000 Then
                       Page = Page + 1
                       Printer.NewPage
                       PrintTitle
                       PrintTitle2
                   End If
                   PrintTitle2
               End If
               '案件數
               SavDay(0) = Trim(str(Val(SavDay(0)) + 1))
               SavDayT(0) = Trim(str(Val(SavDayT(0)) + 1))
               '結果數
               'Modify By Cheng 2002/04/11
   '            If Len(CheckStr(.Fields(9))) <> 0 Then
               If Len(CheckStr(strTemp(4))) <> 0 Then
                   SavDay(1) = Trim(str(Val(SavDay(1)) + 1))
                   SavDayT(1) = Trim(str(Val(SavDayT(1)) + 1))
               End If
               '核准數
               Select Case strTemp(4)
               Case "准"
                   SavDay(2) = Trim(str(Val(SavDay(2)) + 1))
                   SavDayT(2) = Trim(str(Val(SavDayT(2)) + 1))
               Case Else
               End Select
               SavDay(5) = Trim(str(Val(SavDay(5)) + Val(strTemp(7))))
               SavDayT(5) = Trim(str(Val(SavDayT(5)) + Val(strTemp(7))))
               strTemp(0) = StrToStr(strTemp(0), 4)
               'Modify By Cheng 2002/04/08
   '            strTemp(3) = StrToStr(strTemp(3), 4)
               strTemp(3) = StrToStr(strTemp(3), 20)
               strTemp(5) = StrToStr(strTemp(5), 4)
               strTemp(6) = StrToStr(strTemp(6), 4)
               PrintDatil2
               If iPrint >= 10000 Then
                   Page = Page + 1
                   Printer.NewPage
                   PrintTitle
                   PrintTitle2
               End If
               .MoveNext
           Loop
       End If
   End With
   CheckOC
   PrintEnd2 (0)
   PrintEnd2 (1)
   Printer.EndDoc
   ShowPrintOk
End Sub

Sub PrintTitle2()
   GetPleft2
   Printer.Font.Size = 12
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print "承辦人"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   If txt1(2) = "1" Then
       Printer.Print "發文日期"
   Else
       Printer.Print "准駁日期"
   End If
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "案件名稱"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "准/駁"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "業務區"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "智權人員"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "點數"
   iPrint = iPrint + 300
   If iPrint >= 10000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle
       PrintTitle2
       Exit Sub
   End If
   Printer.Font.Size = 12
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   If iPrint >= 10000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle
       PrintTitle2
       Exit Sub
   End If
End Sub

Sub PrintDatil2()
   'Modify By Cheng 2002/02/01
   'For i = 0 To 6
   '    Printer.CurrentX = PLeft(i)
   '    Printer.CurrentY = iPrint
   '    Printer.Print strTemp(i)
   'Next i
   i = 0 '承辦人
   'Modify By Cheng 2002/02/25
   If strTemp(i) <> m_strPromoter Then
      Printer.CurrentX = PLeft(i)
      Printer.CurrentY = iPrint
      Printer.Print strTemp(i)
      m_strPromoter = strTemp(i)
   End If
   i = 1 '發文日期
   Printer.CurrentX = PLeft(i)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(i)
   i = 2 '本所案號
   Printer.CurrentX = PLeft(i)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(i)
   i = 3 '案件名稱
   Printer.CurrentX = PLeft(i)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(i)
   i = 4 '准/駁
   Printer.CurrentX = PLeft(i)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(i)
   i = 5 '業務區
   'Modify By Cheng 2002/04/11
   '不論業務區是否相同皆要列印出來
   
   ''若業務區不同
   'If m_strSaleZone <> strTemp(i) Then
      Printer.CurrentX = PLeft(i)
      Printer.CurrentY = iPrint
      Printer.Print strTemp(i)
      m_strSaleZone = strTemp(i)
   '   i = 6 '智權人員
   '   Printer.CurrentX = PLeft(i)
   '   Printer.CurrentY = iPrint
   '   Printer.Print strTemp(i)
   '   m_strSales = strTemp(i)
   ''若業務區相同
   'Else
      i = 6 '智權人員
      'Modify By Cheng 2002/04/11
      '不論智權人員是否相同皆要列印出來
   '   '若智權人員不相同
   '   If m_strSales <> strTemp(i) Then
         Printer.CurrentX = PLeft(i)
         Printer.CurrentY = iPrint
         Printer.Print strTemp(i)
         m_strSales = strTemp(i)
   '   End If
   'End If
   
   Printer.CurrentX = PLeft(7) + 300 - Printer.TextWidth(strTemp(7))
   Printer.CurrentY = iPrint
    Printer.Print strTemp(7)
   iPrint = iPrint + 300
End Sub

Sub GetPleft2()
   Erase PLeft
   PLeft(0) = 500
   PLeft(1) = 1500
   PLeft(2) = 2700
   PLeft(3) = 4700
   PLeft(4) = 11000
   PLeft(5) = 12000
   PLeft(6) = 13500
   PLeft(7) = 15000
End Sub

Sub PrintEnd2(x As Integer)
   '   0   小計
   '   1   總計
   'Add By Cheng 2002/02/01
   m_strSaleZone = ""
   m_strSales = ""
   
   Select Case x
   Case 0 '小計
        If Val(SavDay(0)) = 0 Then
            SavDay(4) = "0"
        Else
            'Modify By Cheng 2002/04/11
   '         SavDay(4) = Trim(str(Val(SavDay(2)) / Val(SavDay(0)) * 100))
            If Val(SavDay(1)) = 0 Then
               SavDay(4) = "0"
            Else
               SavDay(4) = Trim(str(Val(SavDay(2)) / Val(SavDay(1)) * 100))
            End If
        End If
        Printer.CurrentX = 500
        Printer.CurrentY = iPrint
        Printer.Print String(200, "-")
        iPrint = iPrint + 300
        If iPrint >= 10000 Then
           Page = Page + 1
           Printer.NewPage
           PrintTitle
           PrintTitle2
        End If
        Printer.CurrentX = 500
        Printer.CurrentY = iPrint
        Printer.Print "小計 ==＞   "
        Printer.CurrentX = 500 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "案件數：" & SavDay(0)
        Printer.CurrentX = 3000 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "結果數：" & SavDay(1)
        Printer.CurrentX = 5500 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "核准數：" & SavDay(2)
        Printer.CurrentX = 8000 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "勝訴率：" & Format(SavDay(4), "###.00") & "%"
        Printer.CurrentX = 10500 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "點數：" & SavDay(5)
        iPrint = iPrint + 300
   Case 1 '總計
        If Val(SavDayT(0)) = 0 Then
            SavDayT(4) = "0"
        Else
            'Modify By Cheng 2002/04/11
   '         SavDayT(4) = Trim(str(Val(SavDayT(2)) / Val(SavDayT(0)) * 100))
            If Val(SavDayT(1)) = 0 Then
               SavDayT(4) = "0"
            Else
               SavDayT(4) = Trim(str(Val(SavDayT(2)) / Val(SavDayT(1)) * 100))
            End If
        End If
        Printer.CurrentX = 500
        Printer.CurrentY = iPrint
        Printer.Print String(200, "-")
        iPrint = iPrint + 300
        If iPrint >= 10000 Then
           Page = Page + 1
           Printer.NewPage
           PrintTitle
           PrintTitle2
        End If
        Printer.CurrentX = 500
        Printer.CurrentY = iPrint
        Printer.Print "總計 ==＞   "
        Printer.CurrentX = 500 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "案件數：" & SavDayT(0)
        Printer.CurrentX = 3000 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "結果數：" & SavDayT(1)
        Printer.CurrentX = 5500 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "核准數：" & SavDayT(2)
        Printer.CurrentX = 8000 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "勝訴率：" & Format(SavDayT(4), "###.00") & "%"
        Printer.CurrentX = 10500 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "點數：" & SavDayT(5)
        iPrint = iPrint + 300
   Case Else
   End Select
End Sub

Sub PrintTitle()
   iPrint = 500
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 5800
   Printer.CurrentY = iPrint
   Printer.Print "申請意見書案件明細表"
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   iPrint = iPrint + 500
   Printer.CurrentX = 6200
   Printer.CurrentY = iPrint
   If txt1(2) = "1" Then
       Printer.Print "發文日期：" & Format(ChangeTStringToTDateString(txt1(0)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(1))
   Else
       Printer.Print "准駁日期：" & Format(ChangeTStringToTDateString(txt1(0)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(1))
   End If
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
   iPrint = iPrint + 300
   
   'Add By Cheng 2002/02/01
   'Printer.CurrentX = 6200
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   If Me.txt1(3).Text = "1" Then
       Printer.Print "列印順序： 承 辦 人"
   Else
       Printer.Print "列印順序： 業 務 區"
   End If
   
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(Page)
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
End Sub

Sub PrintTitle1()
   GetPleft1
   Printer.Font.Size = 12
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print "業務區"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "智權人員"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   If txt1(2) = "1" Then
       Printer.Print "發文日期"
   Else
       Printer.Print "准駁日期"
   End If
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "案件名稱"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "准/駁"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "承辦人"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "點數"
   iPrint = iPrint + 300
   If iPrint >= 10000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle
       PrintTitle1
       Exit Sub
   End If
   Printer.Font.Size = 12
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   If iPrint >= 10000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle
       PrintTitle1
       Exit Sub
   End If
End Sub

Sub PrintDatil1()
   'Modify By Cheng 2002/02/01
   'For i = 0 To 6
   '    Printer.CurrentX = PLeft(i)
   '    Printer.CurrentY = iPrint
   '    Printer.Print strTemp(i)
   'Next i
   i = 0 '業務區
   '若業務區不同
   If m_strSaleZone <> strTemp(i) Then
      Printer.CurrentX = PLeft(i)
      Printer.CurrentY = iPrint
      Printer.Print strTemp(i)
      m_strSaleZone = strTemp(i)
      i = 1 '智權人員
      Printer.CurrentX = PLeft(i)
      Printer.CurrentY = iPrint
      Printer.Print strTemp(i)
      m_strSales = strTemp(i)
   '若業務區相同
   Else
      i = 1 '智權人員
      '若智權人員不相同
      If m_strSales <> strTemp(i) Then
         Printer.CurrentX = PLeft(i)
         Printer.CurrentY = iPrint
         Printer.Print strTemp(i)
         m_strSales = strTemp(i)
      End If
   End If
   i = 2 '日期
   Printer.CurrentX = PLeft(i)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(i)
   i = 3 '本所案號
   Printer.CurrentX = PLeft(i)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(i)
   i = 4 '案件名稱
   Printer.CurrentX = PLeft(i)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(i)
   i = 5 '准/駁
   Printer.CurrentX = PLeft(i)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(i)
   i = 6 '承辦人
   Printer.CurrentX = PLeft(i)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(i)
   
   Printer.CurrentX = PLeft(7) + 300 - Printer.TextWidth(strTemp(7))
   Printer.CurrentY = iPrint
   Printer.Print strTemp(7)
   iPrint = iPrint + 300
End Sub

Sub GetPleft1()
   Erase PLeft
   PLeft(0) = 500
   PLeft(1) = 1500
   PLeft(2) = 2500
   PLeft(3) = 3700
   PLeft(4) = 6000
   PLeft(5) = 12000
   PLeft(6) = 13500
   PLeft(7) = 15000
End Sub

Sub PrintEnd1(x As Integer)
   '    0          小計
   '    1          合計
   '    2          總計
   'Add By Cheng 2002/02/01
   m_strSaleZone = ""
   m_strSales = ""
   
   Select Case x
   Case 0 '個人小計
        If Val(SavDay(0)) = 0 Then
            SavDay(4) = "0"
        Else
            'Modify By Cheng 2002/04/11
   '         SavDay(4) = Trim(str(Val(SavDay(2)) / Val(SavDay(0)) * 100))
            If Val(SavDay(1)) = 0 Then
               SavDay(4) = "0"
            Else
               SavDay(4) = Trim(str(Val(SavDay(2)) / Val(SavDay(1)) * 100))
            End If
        End If
        Printer.CurrentX = 500
        Printer.CurrentY = iPrint
        Printer.Print String(200, "-")
        iPrint = iPrint + 300
        If iPrint >= 10000 Then
           Page = Page + 1
           Printer.NewPage
           PrintTitle
           PrintTitle1
        End If
        Printer.CurrentX = 500
        Printer.CurrentY = iPrint
        Printer.Print "個人小計 ==＞   "
        Printer.CurrentX = 500 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "案件數：" & SavDay(0)
        Printer.CurrentX = 3000 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "結果數：" & SavDay(1)
        Printer.CurrentX = 5500 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "核准數：" & SavDay(2)
        Printer.CurrentX = 8000 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "勝訴率：" & Format(SavDay(4), "###.00") & "%"
        Printer.CurrentX = 10500 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "點數：" & SavDay(5)
        iPrint = iPrint + 300
   Case 1 '區小計
        If Val(SavDay22(0)) = 0 Then
            SavDay22(4) = "0"
        Else
            'Modify By Cheng 2002/04/11
   '         SavDay22(4) = Trim(str(Val(SavDay22(2)) / Val(SavDay22(0)) * 100))
            If Val(SavDay22(1)) = 0 Then
               SavDay22(4) = "0"
            Else
               SavDay22(4) = Trim(str(Val(SavDay22(2)) / Val(SavDay22(1)) * 100))
            End If
        End If
        Printer.CurrentX = 500
        Printer.CurrentY = iPrint
        Printer.Print String(200, "-")
        iPrint = iPrint + 300
        If iPrint >= 10000 Then
           Page = Page + 1
           Printer.NewPage
           PrintTitle
           PrintTitle1
        End If
        Printer.CurrentX = 500
        Printer.CurrentY = iPrint
        Printer.Print "區小計 ==＞   "
        Printer.CurrentX = 500 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "案件數：" & SavDay22(0)
        Printer.CurrentX = 3000 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "結果數：" & SavDay22(1)
        Printer.CurrentX = 5500 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "核准數：" & SavDay22(2)
        Printer.CurrentX = 8000 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "勝訴率：" & Format(SavDay22(4), "###.00") & "%"
        Printer.CurrentX = 10500 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "點數：" & SavDay22(5)
        iPrint = iPrint + 300
   Case 2 '總計
        If Val(SavDayT(0)) = 0 Then
            SavDayT(4) = "0"
        Else
            'Modify By Cheng 2002/04/11
   '         SavDayT(4) = Trim(str(Val(SavDayT(2)) / Val(SavDayT(0)) * 100))
            If Val(SavDayT(1)) = 0 Then
               SavDayT(4) = "0"
            Else
               SavDayT(4) = Trim(str(Val(SavDayT(2)) / Val(SavDayT(1)) * 100))
            End If
        End If
        Printer.CurrentX = 500
        Printer.CurrentY = iPrint
        Printer.Print String(200, "-")
        iPrint = iPrint + 300
        If iPrint >= 10000 Then
           Page = Page + 1
           Printer.NewPage
           PrintTitle
           PrintTitle1
        End If
        Printer.CurrentX = 500
        Printer.CurrentY = iPrint
        Printer.Print "總計 ==＞   "
        Printer.CurrentX = 500 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "案件數：" & SavDayT(0)
        Printer.CurrentX = 3000 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "結果數：" & SavDayT(1)
        Printer.CurrentX = 5500 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "核准數：" & SavDayT(2)
        Printer.CurrentX = 8000 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "勝訴率：" & Format(SavDayT(4), "###.00") & "%"
        Printer.CurrentX = 10500 + 2500
        Printer.CurrentY = iPrint
        Printer.Print "點數：" & SavDayT(5)
   Case Else
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(cancel As Integer)
   Set frm020308 = Nothing
End Sub

Private Sub txt1_GotFocus(index As Integer)
   txt1(index).SelStart = 0
   txt1(index).SelLength = Len(txt1(index))
End Sub

Private Sub txt1_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
       cmdOK(0).SetFocus
   End If
End Sub

Private Sub txt1_KeyPress(index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(index As Integer)
   Select Case index
   Case 0, 1
      If PUB_CheckKeyInDate(Me.txt1(index)) = -1 Then
         Me.txt1(index).SetFocus
         txt1_GotFocus index
         Exit Sub
      End If
      If index = 1 Then
        If RunNick(txt1(index - 1), txt1(index)) Then
            txt1(index - 1).SetFocus
            txt1_GotFocus (index - 1)
            Exit Sub
         End If
      End If
   Case 2
        Select Case Trim(txt1(2))
        Case "1", "2", ""
        Case Else
             s = MsgBox("列印內容只能輸入 1 或 2 !!", , "USER 輸入錯誤")
             txt1(2).SetFocus
             Exit Sub
        End Select
   Case 3
        Select Case Trim(txt1(3))
        Case "1", "2", ""
        Case Else
             s = MsgBox("列印順序只能輸入 1 或 2 !!", , "USER 輸入錯誤")
             txt1(3).SetFocus
             Exit Sub
        End Select
   Case Else
   End Select
End Sub
