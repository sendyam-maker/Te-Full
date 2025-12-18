VERSION 5.00
Begin VB.Form frm020408 
   BorderStyle     =   1  '單線固定
   Caption         =   "商爭案承辦人勝敗統計表"
   ClientHeight    =   2700
   ClientLeft      =   3030
   ClientTop       =   1530
   ClientWidth     =   3900
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   3900
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   0
      TabIndex        =   11
      Top             =   1635
      Width           =   3825
      Begin VB.ComboBox Combo1 
         Height          =   276
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
      Left            =   3144
      TabIndex        =   7
      Top             =   12
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2352
      TabIndex        =   6
      Top             =   0
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2112
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1290
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1092
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1290
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2112
      MaxLength       =   7
      TabIndex        =   2
      Top             =   930
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1092
      MaxLength       =   7
      TabIndex        =   1
      Top             =   930
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1092
      TabIndex        =   0
      Top             =   570
      Width           =   1740
   End
   Begin VB.Line Line2 
      X1              =   1635
      X2              =   2385
      Y1              =   1395
      Y2              =   1395
   End
   Begin VB.Line Line1 
      X1              =   1530
      X2              =   2790
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   3
      Left            =   150
      TabIndex        =   10
      Top             =   1320
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "勝敗日期："
      Height          =   180
      Index           =   2
      Left            =   150
      TabIndex        =   9
      Top             =   975
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   8
      Top             =   630
      Width           =   915
   End
End
Attribute VB_Name = "frm020408"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/9 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, SavDay4 As String, SavDay5 As String, SavDay6 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 32) As String, strTemp3 As String, TestOk As Boolean, StrTemp7(0 To 31) As String
Dim PLeft(0 To 31) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, PLeft1(1 To 7) As Integer, SeekPrint As Integer, SeekPrintL As Integer, k As Integer
Dim BolEndThisPage As Boolean
'Add By Cheng 2002/05/06
Dim m_strPromoter As String '承辦人
Dim Adorecordset99 As New ADODB.Recordset
'Add By Cheng 2003/03/12
Dim strSQL6_1 As String
Dim strSQL6_2 As String     '2008/11/18 add by sonia
Dim bolNotData As Boolean 'Add By Sindy 2011/3/1
'Modify By Sindy 2015/2/4
Dim System_ID As String
Dim bolIsChina As Boolean
Dim Title_601 As String
Dim Title_603 As String
Dim Title_605 As String
Dim Title_401 As String
Dim Title_403 As String
Dim Title_404 As String
Dim Title_408 As String
Dim Title_602 As String
Dim Title_604 As String
Dim Title_606 As String
Dim Title_406 As String
Dim Title_407 As String
'2015/2/4 END


Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     PUB_RestorePrinter Combo1 'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除
     Printer.EndDoc 'Add By Sindy 2011/11/1
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
             s = MsgBox("勝敗日期區間不可空白!!", , "USER 輸入錯誤")
             txt1(1).SetFocus
             txt1_GotFocus (1)
             Exit Sub
         Else
            'Modify By Cheng 2004/03/11
'             If Format(ChangeTStringToTDateString(txt1(1)), "yy") = Format(ChangeTStringToTDateString(txt1(2)), "yy") Then
'                 If Format(ChangeTStringToTDateString(txt1(1)), "mm") = Format(ChangeTStringToTDateString(txt1(2)), "mm") Then
'                    If Val(Format(ChangeTStringToTDateString(txt1(1)), "dd")) = 1 Then
             If Format(ChangeTStringToWDateString(txt1(1)), "yyyy") = Format(ChangeTStringToWDateString(txt1(2)), "yyyy") Then
                 If Format(ChangeTStringToWDateString(txt1(1)), "mm") = Format(ChangeTStringToWDateString(txt1(2)), "mm") Then
                    If Val(Format(ChangeTStringToWDateString(txt1(1)), "dd")) = 1 Then
            'End
                        'Modify By Cheng 2003/03/26
'                        If Format(ChangeTStringToTDateString(txt1(1)), "mm") <> Format(DateAdd("d", 1, ChangeTStringToTDateString(txt1(2))), "mm") Then
                        If Format(ChangeTStringToWDateString(txt1(1)), "mm") <> Format(DateAdd("d", 1, ChangeTStringToWDateString(txt1(2))), "mm") Then
                            TestOk = True
                        Else
                            TestOk = False
                        End If
                    Else
                        TestOk = False
                    End If
                 Else
                    TestOk = False
                 End If
             Else
                 TestOk = False
             End If
             Screen.MousePointer = vbHourglass
             Me.Enabled = False
             ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/19 清除查詢印表記錄檔欄位
             Process
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
End Sub

Sub Process()
'Add By Cheng 2003/03/11
Dim StrSQLa As String

Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM r020408_1 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "DELETE FROM r020408_2 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL6 = ""
strSQL6_1 = "": strSQL6_2 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
'   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/10/19
End If
'Add By Cheng 2003/07/11
'若非商申收發文時, 若為TF案則不抓後三碼為"000"的資料
strSQL1 = strSQL1 + " AND CP03 <> Decode(CP01,'TF','0','z') AND CP04 <> Decode(CP01,'TF','00','zz') "
StrSQL6 = ""
If Len(txt1(3)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND TM10>='" & txt1(3) & "' "
    strSQL6_1 = strSQL6_1 + " AND TM10>='" & txt1(3) & "' "
End If
If Len(txt1(4)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND TM10<='" & txt1(4) & "' "
    strSQL6_1 = strSQL6_1 + " AND TM10<='" & txt1(4) & "' "
End If
If Len(txt1(3)) <> 0 Or Len(txt1(4)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/10/19
End If
'2008/11/18 modify by sonia 申請意見書不抓公告日改抓審定來函日,大陸案已無此案件性質
strSQL6_2 = strSQL6_1
If Len(txt1(1)) <> 0 Then
   StrSQL6 = StrSQL6 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(1))) & ""
   strSQL6_1 = strSQL6_1 + " AND TM14>=" & Val(ChangeTStringToWString(txt1(1))) & ""
   strSQL6_2 = strSQL6_2 + " AND TM13>=" & Val(ChangeTStringToWString(txt1(1))) & ""
End If
If Len(Trim(txt1(2))) <> 0 Then
   StrSQL6 = StrSQL6 + " AND CP05<=" & Val(ChangeTStringToWString(txt1(2))) & " "
   strSQL6_1 = strSQL6_1 + " AND TM14<=" & Val(ChangeTStringToWString(txt1(2))) & " "
   strSQL6_2 = strSQL6_2 + " AND TM13<=" & Val(ChangeTStringToWString(txt1(2))) & " "
End If
If Len(txt1(1)) <> 0 Or Len(Trim(txt1(2))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/10/19
End If
'Add By Cheng 2002/02/07
strSQL1 = strSQL1 + " And (CP10='1003' OR CP10='1004')"
'Add By Cheng 2003/03/11
'2011/12/7 modify by sonia
'strSQL2 = strSQL2 + " And CP10='202' "
strSQL2 = strSQL2 + " And CP10 in ('202','210') "
'2011/12/7 end
'93.6.14 CANCEL BY SONIA 僅收/發文統計表才控制不算案件數
''Add By Cheng 2004/04/29
''抓計件的資料
'StrSQL6 = StrSQL6 & " And CP26 Is Null "
'strSQL6_1 = strSQL6_1 & " And CP26 Is Null "
''End
CheckOC
'有承辦人申爭之分
'Modify By Cheng 2002/02/07
'strSQL = "SELECT s2.st02,NVL(A0902,A0903),CP24,S2.ST03,CP10,decode(tm10,'000','*','') FROM CASEPROGRESS,STAFF S1,STAFF S2,TRADEMARK,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND (CP26 IS NULL OR CP26='') AND (CP57 IS NULL OR CP57='') AND (S2.ST05='97' OR S2.ST05='17') AND CP09>'C' AND S1.ST03=A0901(+) " & strSQL1 + StrSQL6
'取消抓承辦人之ST05為"97"或"17"及CP09>"C"的條件
'先設定相關總收文號
strSql = " SELECT CP43 " & _
         " FROM CASEPROGRESS,TRADEMARK " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP27 IS NOT NULL AND (CP57 IS NULL OR CP57='') " & strSQL1 + StrSQL6
'Add By Cheng 2003/03/11
'抓案件性質為申請意見書的本所案號
'strSQLA = "SELECT DISTINCT CP09 FROM CASEPROGRESS,( SELECT CP01 C1, CP02 C2,CP03 C3, CP04 C4 " & _
'         " FROM CASEPROGRESS,TRADEMARK " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND (CP57 IS NULL OR CP57='') And (CP22 IS NULL OR CP22 <>'N')  " & strSQL2 + StrSQL6 & " GROUP BY CP01,CP02, CP03, CP04 ) C WHERE CP01=C.C1 AND CP02=C.C2 AND CP03=C.C3 AND CP04=C.C4 "
'2008/11/18 modify by sonia 申請意見書不抓公告日改抓審定來函日
'StrSQLa = " SELECT DISTINCT CP09 " & _
         " FROM CASEPROGRESS,TRADEMARK " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND (CP57 IS NULL OR CP57='') " & strSQL2 + strSQL6_1
'Modify By Sindy 2013/1/9
'StrSQLa = " SELECT DISTINCT CP09 " & _
'         " FROM CASEPROGRESS,TRADEMARK " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND (CP57 IS NULL OR CP57='') " & strSQL2 + strSQL6_2
StrSQLa = " SELECT C2.CP09 FROM (" & _
"SELECT CP01,CP02,CP03,CP04,MAX(CP27) MAXCP27 " & _
"From CASEPROGRESS, Trademark " & _
"WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) " & _
"AND CP27 IS NOT NULL AND (CP57 IS NULL OR CP57='') " & strSQL2 + strSQL6_2 & _
"GROUP BY CP01,CP02,CP03,CP04) C1,CASEPROGRESS C2 " & _
"WHERE C1.CP01=C2.CP01(+) AND C1.CP02=C2.CP02(+) AND C1.CP03=C2.CP03(+) AND C1.CP04=C2.CP04(+) " & _
"AND C2.CP10 in('202','210') AND C1.MAXCP27=C2.CP27(+) "
'2013/1/9 End
'2008/11/18 end
'910521 nick
'strSQL = "SELECT st02,NVL(A0902,A0903),CP24,ST03,CP10,decode(tm10,'000','*','') " & _
         " FROM CASEPROGRESS,STAFF,TRADEMARK,ACC090 " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND (CP26 IS NULL OR CP26='') AND (CP57 IS NULL OR CP57='') AND CP12=A0901(+) AND CP09 IN ( " & strSQL & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(ST03,1,2)='P2' OR CP14 IS NULL ) ", "  AND (SUBSTR(ST03,1,2)='F1' OR CP14 IS NULL )  ")
'**************** 將業務區改成抓案件進度檔   91.08.15  nick
'Modify By Cheng 2003/03/07
'不限制是否要算案件數
'strSQL = "SELECT cp14,cp12,CP24,ST03,CP10,decode(tm10,'000','*','') " & _
'         " FROM CASEPROGRESS,STAFF,TRADEMARK,ACC090 " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND (CP26 IS NULL OR CP26='') AND (CP57 IS NULL OR CP57='') AND CP12=A0901(+) AND CP09 IN ( " & strSQL & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(ST03,1,2)='P2' OR CP14 IS NULL ) ", "  AND (SUBSTR(ST03,1,2)='F1' OR CP14 IS NULL )  ")
'edit by nickc 2005/05/13
'StrSql = "SELECT cp14,cp12,CP24,ST03,CP10,decode(tm10,'000','*','') " & _
         " FROM CASEPROGRESS,STAFF,TRADEMARK,ACC090 " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND (CP57 IS NULL OR CP57='') AND CP12=A0901(+) AND CP09 IN ( " & StrSql & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(ST03,1,2)='P2' OR CP14 IS NULL ) ", "  AND (SUBSTR(ST03,1,2)='F1' OR CP14 IS NULL )  ")
strSql = "SELECT cp14,cp12,CP24,s1.ST03,CP10,decode(substr(s2.st15,1,1),'F',' ','*') " & _
         " FROM CASEPROGRESS,STAFF s1,TRADEMARK,ACC090,staff s2 " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND cp13=s2.st01(+) and CP14=s1.ST01(+) AND CP27 IS NOT NULL AND (CP57 IS NULL OR CP57='') AND CP12=A0901(+) AND CP09 IN ( " & strSql & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(s1.ST03,1,2)='P2' OR CP14 IS NULL Or CP14='A6015') ", "  AND (SUBSTR(s1.ST03,1,2)='F1' OR CP14 IS NULL )  ")

'沒有承辦人申爭之分
'StrSQL = "SELECT NVL(A0902,A0903),S1.ST02,CP24,S1.ST03,CP10 FROM CASEPROGRESS,STAFF S1,STAFF S2,TRADEMARK,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND (CP26 IS NULL OR CP26='') AND (CP57 IS NULL OR CP57='') AND CP09>'C'  AND S1.ST03=A0901(+) " & StrSQL1 + StrSQL6
'Add By Cheng 2003/03/11
'與申請意見書相同本所案號且案件性質為申請的資料
'strSQL = strSQL & " Union All SELECT cp14,cp12,CP24,ST03,CP10,decode(tm10,'000','*','') " & _
'         " FROM CASEPROGRESS,STAFF,TRADEMARK,ACC090 " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND (CP57 IS NULL OR CP57='') AND CP12=A0901(+) AND CP10='101' AND CP24 IS NOT NULL AND CP09 IN ( " & strSQLA & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(ST03,1,2)='P2' OR CP14 IS NULL ) ", "  AND (SUBSTR(ST03,1,2)='F1' OR CP14 IS NULL )  ")
'Modify By Cheng 2003/04/10
'不考慮是否出名
'strSQL = strSQL & " UNION ALL SELECT cp14,cp12,TM16,ST03,CP10,decode(tm10,'000','*','') " & _
'         " FROM CASEPROGRESS,STAFF,TRADEMARK,ACC090 " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND (CP57 IS NULL OR CP57='') AND CP12=A0901(+) AND (CP22 IS NULL OR CP22 <>'N') AND TM16 IS NOT NULL AND CP09 IN ( " & strSQLA & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(ST03,1,2)='P2' OR CP14 IS NULL ) ", "  AND (SUBSTR(ST03,1,2)='F1' OR CP14 IS NULL )  ")
'edit by nickc 2005/05/13
'StrSql = StrSql & " UNION ALL SELECT cp14,cp12,TM16,ST03,CP10,decode(tm10,'000','*','') " & _
         " FROM CASEPROGRESS,STAFF,TRADEMARK,ACC090 " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND (CP57 IS NULL OR CP57='') AND CP12=A0901(+) AND TM16 IS NOT NULL AND CP09 IN ( " & strSQLA & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(ST03,1,2)='P2' OR CP14 IS NULL ) ", "  AND (SUBSTR(ST03,1,2)='F1' OR CP14 IS NULL )  ")
strSql = strSql & " UNION ALL SELECT cp14,cp12,TM16,s1.ST03,CP10,decode(substr(s2.st15,1,1),'F',' ','*') " & _
         " FROM CASEPROGRESS,STAFF s1,TRADEMARK,ACC090,staff s2 " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) and cp13=s2.st01(+) AND CP14=s1.ST01(+) AND CP27 IS NOT NULL AND (CP57 IS NULL OR CP57='') AND CP12=A0901(+) AND TM16 IS NOT NULL AND CP09 IN ( " & StrSQLa & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(s1.ST03,1,2)='P2' OR CP14 IS NULL Or CP14='A6015') ", "  AND (SUBSTR(s1.ST03,1,2)='F1' OR CP14 IS NULL )  ")

With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/19
        .MoveFirst
        DoEvents
        Do While .EOF = False
            For i = 0 To 5
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If Val(strTemp(2)) = 1 Then
                SavDay1 = "1"
                SavDay2 = "0"
            Else
                SavDay1 = "0"
                SavDay2 = "1"
            End If
            Select Case Val(strTemp(4))
            '表(1)格式1
            Case 601, 627 '異議,Add by Sindy 2019/8/15 +部分異議
                 cnnConnection.Execute "INSERT INTO r020408_1 (r079001,r079002,r079003,r079004,r079005,r079024,r079025,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & ",0,'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
            Case 603, 629 '評定,Add by Sindy 2019/8/15 +部分評定
                 cnnConnection.Execute "INSERT INTO r020408_1 (r079001,r079002,r079006,r079007,r079008,r079024,r079025,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & ",0,'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
            Case 605, 623 '廢止,Add by Sindy 2019/8/15 +部分廢止
                 cnnConnection.Execute "INSERT INTO r020408_1 (r079001,r079002,r079009,r079010,r079011,r079024,r079025,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & ",0,'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
            'Modify By Cheng 2002/05/06
'            Case 401 '訴願
            'Modify By Cheng 2003/03/07
            '參加訴願移至表(2)
'            Case 401, 406 '訴願,參加訴願
'            'add by nickc 2007/07/27 商標處改格式
'            Case 618
'                If txt1(3) = "020" And txt1(4) = "020" Then
'                     cnnConnection.Execute "INSERT INTO r020408_1 (r079001,r079002,r079012,r079013,r079014,r079024,r079025,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & ",0,'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
'                     cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
'                End If
            Case 401 '訴願
                'add by nickc 2007/07/27 商標處改格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                    cnnConnection.Execute "INSERT INTO r020408_1 (r079001,r079002,r079015,r079016,r079017,r079024,r079025,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & ",0,'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                Else
                    cnnConnection.Execute "INSERT INTO r020408_1 (r079001,r079002,r079012,r079013,r079014,r079024,r079025,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & ",0,'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                End If
            'Modify By Cheng 2002/05/06
            '取消再訴願欄, 以行政訴訟取代之
'            Case 402 '再訴願
            'Modify By Cheng 2003/03/07
            '參加訴訟移至表(2)
'            Case 403, 407 '行政訴訟, 參加訴訟
            Case 403 '行政訴訟
                'add by nickc 2007/07/27 商標處改格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                    cnnConnection.Execute "INSERT INTO r020408_1 (r079001,r079002,r079018,r079019,r079020,r079024,r079025,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & ",0,'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                Else
                    cnnConnection.Execute "INSERT INTO r020408_1 (r079001,r079002,r079015,r079016,r079017,r079024,r079025,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & ",0,'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                End If
            'Modify By Cheng 2002/05/06
            '本欄以行政訴訟上訴取代
'            Case 403 '行政訴訟
            Case 408 '行政訴訟上訴
                'add by nickc 2007/07/27 商標處改格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                Else
                    cnnConnection.Execute "INSERT INTO r020408_1 (r079001,r079002,r079018,r079019,r079020,r079024,r079025,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & ",0,'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                End If
            'Modify By Cheng 2002/05/06
'            Case 405 '再審之訴
            Case 404 '再審之訴
                'add by nickc 2007/07/27 商標處改格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                Else
                    cnnConnection.Execute "INSERT INTO r020408_1 (r079001,r079002,r079021,r079022,r079023,r079024,r079025,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & ",0,'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                End If
            '表(1)格式2
            'Modify By Cheng 2003/03/11
'            'Add By Cheng 2003/03/07
'            Case "202" '申請意見書
'            Case "101" '申請意見書(<--申請)
            '2011/12/7 MODIFY BY SONIA 加210陳述意見書
            Case "202", "210" '申請意見書,陳述意見書
                'add by nickc 2007/07/27 商標處改格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                Else
                    cnnConnection.Execute "INSERT INTO r020408_1 (r079001,r079002,r079026,r079027,r079028,r079024,r079025,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & ",0,'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                End If
'***********************************
            '表(2)格式1
            Case 602, 628 '異議答辯,Add by Sindy 2019/8/15 +部分異議答辯
                 cnnConnection.Execute "insert into r020408_1 (r079001,r079002,r079024,r079025,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080003,r080004,r080005,r080006,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
            Case 1602, 1601 '被異議(理由), 被異議
                 cnnConnection.Execute "insert into r020408_1 (r079001,r079002,r079024,r079025,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080003,r080004,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",0,0,'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
            Case 604, 630 '評定答辯,Add by Sindy 2019/8/15 +部分評定答辯
                 cnnConnection.Execute "insert into r020408_1 (r079001,r079002,r079024,r079025,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080009,r080010,r080011,r080012,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
            Case 1603, 1604 '被評定, 被評定(理由)
                 cnnConnection.Execute "insert into r020408_1 (r079001,r079002,r079024,r079025,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080009,r080010,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",0,0,'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
            Case 606, 624 '廢止答辯,Add by Sindy 2019/8/15 +部分廢止答辯
                 cnnConnection.Execute "insert into r020408_1 (r079001,r079002,r079024,r079025,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080015,r080016,r080017,r080018,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
            Case 1605, 1606 '被廢止, 被廢止(理由)
                 cnnConnection.Execute "insert into r020408_1 (r079001,r079002,r079024,r079025,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080015,r080016,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",0,0,'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
            'Modify By Cheng 2003/03/07
            '移至表(1)
'            Case 202 '申請意見書
'                 cnnConnection.Execute "insert into r020408_1 (r079001,r079002,r079024,r079025,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
'                 cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080021,r080022,r080023,r080024,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
'            Case 1202 '核駁前先行通知
'                 cnnConnection.Execute "insert into r020408_1 (r079001,r079002,r079024,r079025,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
'                 cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080021,r080022,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",0,0,'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
'            'add by nickc 2007/07/27 商標處改格式
'            Case 619
'                If txt1(3) = "020" And txt1(4) = "020" Then
'                     cnnConnection.Execute "insert into r020408_1 (r079001,r079002,r079024,r079025,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
'                     cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080021,r080022,r080023,r080024,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
'                End If
            'Add By Cheng 2003/03/07
            Case 406 '參加訴願
                'add by nickc 2007/07/27 商標處改格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                    cnnConnection.Execute "insert into r020408_1 (r079001,r079002,r079024,r079025,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080035,r080036,r080037,r080038,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                Else
                    cnnConnection.Execute "insert into r020408_1 (r079001,r079002,r079024,r079025,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080021,r080022,r080023,r080024,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                End If
            Case 1404 '通知參加訴願
                'add by nickc 2007/07/27 商標處改格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                    cnnConnection.Execute "insert into r020408_1 (r079001,r079002,r079024,r079025,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080035,r080036,r080037,r080038,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                Else
                    cnnConnection.Execute "insert into r020408_1 (r079001,r079002,r079024,r079025,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080021,r080022,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",0,0,'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                End If
            Case 407 '參加訴訟
                'add by nickc 2007/07/27 商標處改格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                Else
                    cnnConnection.Execute "insert into r020408_1 (r079001,r079002,r079024,r079025,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080035,r080036,r080037,r080038,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                End If
            Case 1405 '通知參加訴訟
                'add by nickc 2007/07/27 商標處改格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                Else
                    cnnConnection.Execute "insert into r020408_1 (r079001,r079002,r079024,r079025,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080035,r080036,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",0,0,'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                End If
            Case 410 '上訴答辯
                'add by nickc 2007/07/27 商標處改格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                Else
                    cnnConnection.Execute "insert into r020408_1 (r079001,r079002,r079024,r079025,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080041,r080042,r080043,r080044,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                End If
            Case 1406 '通知上訴答辯
                'add by nickc 2007/07/27 商標處改格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                Else
                    cnnConnection.Execute "insert into r020408_1 (r079001,r079002,r079024,r079025,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020408_2 (r080001,r080002,r080041,r080042,r080027,r080028,r080029,r080030,r080033,r080034,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",0,0,'" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
                End If
            Case Else
            End Select
            DoEvents
            .MoveNext
        Loop
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/10/19
        ShowNoData
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
End With
CheckOC
'Modify By Sindy 2015/2/4
If Trim(txt1(0)) = "" Then
   strTemp1 = Split(GetSystemKindByNickT, ",")
Else
   strTemp1 = Split(txt1(0), ",")
End If
System_ID = strTemp1(0)
bolIsChina = False
If txt1(3) = "020" And txt1(4) = "020" Then
   bolIsChina = True
End If
Call ClsPDGetCaseProperty("T", "601", Title_601, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "603", Title_603, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "605", Title_605, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "401", Title_401, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "403", Title_403, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "404", Title_404, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "408", Title_408, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "602", Title_602, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "604", Title_604, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "606", Title_606, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "406", Title_406, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "407", Title_407, bolIsChina, False)
'2015/2/4 END
bolNotData = True 'Add By Sindy 2011/3/1
PrintData
'Add By Sindy 2011/3/1
If bolNotData = True Then
   ShowNoData
'2011/3/1 End
Else
   ShowPrintOk
End If
Screen.MousePointer = vbDefault
End Sub

Sub PrintData()
BolEndThisPage = False
'表(1)格式1
'nick 910521
'strSQL = "select r079001,r079002,sum(r079003),sum(r079004),sum(r079005),sum(r079006),sum(r079007),sum(r079008),sum(r079009),sum(r079010),sum(r079011),sum(r079012),sum(r079013),sum(r079014),sum(r079015),sum(r079016),sum(r079017),sum(r079018),sum(r079019),sum(r079020),sum(r079021),sum(r079022),sum(r079023),r079024 from r020408_1 where id='" & strUserNum & "' group by r079001,r079002,r079024 "
'Modify By Cheng 2003/03/07
'strSQL = "select st02,NVL(A0902,A0903),sum(r079003),sum(r079004),sum(r079005),sum(r079006),sum(r079007),sum(r079008),sum(r079009),sum(r079010),sum(r079011),sum(r079012),sum(r079013),sum(r079014),sum(r079015),sum(r079016),sum(r079017),sum(r079018),sum(r079019),sum(r079020),sum(r079021),sum(r079022),sum(r079023),r079024,r079001,r079002 from r020408_1,staff,acc090 where R079001=st01(+) and r079002=a0901(+) and id='" & strUserNum & "' group by r079001,r079002,r079024,st02,NVL(A0902,A0903) order by r079001,r079002 "
strSql = "select st02,NVL(A0902,A0903),sum(r079003),sum(r079004),sum(r079005),sum(r079006),sum(r079007),sum(r079008),sum(r079009),sum(r079010),sum(r079011),sum(r079012),sum(r079013),sum(r079014),sum(r079015),sum(r079016),sum(r079017),sum(r079018),sum(r079019),sum(r079020),sum(r079021),sum(r079022),sum(r079023),r079024,r079001,r079002,sum(r079026),sum(r079027),sum(r079028) from r020408_1,staff,acc090 where R079001=st01(+) and r079002=a0901(+) and id='" & strUserNum & "' group by r079001,r079002,r079024,st02,NVL(A0902,A0903) order by r079001,r079002 "
CheckOC
'Add By Cheng 2002/05/06
m_strPromoter = ""
SavDay1 = ""
SavDay2 = ""
SavDay3 = ""
SavDay5 = ""
SavDay6 = ""
Page = 1
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        bolNotData = False 'Add By Sindy 2011/3/1
        .MoveFirst
        SavDay1 = CheckStr(.Fields(0))
        SavDay2 = CheckStr(.Fields(1))
        SavDay3 = CheckStr(.Fields(23))
        SavDay5 = CheckStr(.Fields(24))
        SavDay6 = CheckStr(.Fields(25))
        PrintTitle (1)
        PrintTitle1
        Do While .EOF = False
            For i = 0 To 23
                strTemp(i) = CheckStr(.Fields(i))
                If Val(strTemp(i)) = 0 And i <> 0 And i <> 1 And i <> 23 Then
                    strTemp(i) = "0"
                End If
            Next i
'            strTemp(0) = StrToStr(strTemp(0), 3)
'            strTemp(1) = StrToStr(strTemp(1), 4)
            strTemp(0) = StrToStr(strTemp(0), 5)
            strTemp(1) = StrToStr(strTemp(1), 5)
            If Val(strTemp(2)) + Val(strTemp(3)) = 0 Then
                strTemp(4) = "0"
            Else
                strTemp(4) = Trim(str(Val(strTemp(2)) / (Val(strTemp(2)) + Val(strTemp(3))) * 100))
            End If
            If Val(strTemp(5)) + Val(strTemp(6)) = 0 Then
                strTemp(7) = "0"
            Else
                strTemp(7) = Trim(str(Val(strTemp(5)) / (Val(strTemp(5)) + Val(strTemp(6))) * 100))
            End If
            If Val(strTemp(8)) + Val(strTemp(9)) = 0 Then
                strTemp(10) = "0"
            Else
                strTemp(10) = Trim(str(Val(strTemp(8)) / (Val(strTemp(8)) + Val(strTemp(9))) * 100))
            End If
            If Val(strTemp(11)) + Val(strTemp(12)) = 0 Then
                strTemp(13) = "0"
            Else
                strTemp(13) = Trim(str(Val(strTemp(11)) / (Val(strTemp(11)) + Val(strTemp(12))) * 100))
            End If
            If Val(strTemp(14)) + Val(strTemp(15)) = 0 Then
                strTemp(16) = "0"
            Else
                strTemp(16) = Trim(str(Val(strTemp(14)) / (Val(strTemp(14)) + Val(strTemp(15))) * 100))
            End If
            If Val(strTemp(17)) + Val(strTemp(18)) = 0 Then
                strTemp(19) = "0"
            Else
                strTemp(19) = Trim(str(Val(strTemp(17)) / (Val(strTemp(17)) + Val(strTemp(18))) * 100))
            End If
            If Val(strTemp(20)) + Val(strTemp(21)) = 0 Then
                strTemp(22) = "0"
            Else
                strTemp(22) = Trim(str(Val(strTemp(20)) / (Val(strTemp(20)) + Val(strTemp(21))) * 100))
            End If
            If SavDay1 <> strTemp(0) Then
                'Add By Cheng 2002/05/06
                m_strPromoter = ""
                ShowLine1
                PrintEnd1 (2)
                ShowLine1
                PrintEnd1 (3)
                ShowLine1
                PrintEnd1 (0)
                ShowLine1
                SavDay1 = strTemp(0)
                SavDay5 = CheckStr(.Fields(24))
                SavDay6 = CheckStr(.Fields(25))
            End If
            PrintDatil1
            If iPrint >= 13900 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle (1)
                PrintTitle1
            End If
            .MoveNext
        Loop
    Else
         GoTo ReadNext1: 'Add By Sindy 2011/3/1
    End If
End With
CheckOC
ShowLine1
PrintEnd1 (2)
ShowLine1
PrintEnd1 (3)
ShowLine1
PrintEnd1 (0)
ShowLine1
PrintEnd1 (1)
ShowLine1
Page = Page + 1
Printer.NewPage
ReadNext1: 'Add By Sindy 2011/3/1

'表(1)格式2
'nick 910521
'strSQL = "select r079001,r079002,sum(r079003),sum(r079004),sum(r079005),sum(r079006),sum(r079007),sum(r079008),sum(r079009),sum(r079010),sum(r079011),sum(r079012),sum(r079013),sum(r079014),sum(r079015),sum(r079016),sum(r079017),sum(r079018),sum(r079019),sum(r079020),sum(r079021),sum(r079022),sum(r079023),r079024 from r020408_1 where id='" & strUserNum & "' group by r079001,r079002,r079024 "
'Modify By Cheng 2003/03/07
'strSQL = "select st02,NVL(A0902,A0903),sum(r079003),sum(r079004),sum(r079005),sum(r079006),sum(r079007),sum(r079008),sum(r079009),sum(r079010),sum(r079011),sum(r079012),sum(r079013),sum(r079014),sum(r079015),sum(r079016),sum(r079017),sum(r079018),sum(r079019),sum(r079020),sum(r079021),sum(r079022),sum(r079023),r079024,r079001,r079002 from r020408_1,staff,acc090 where R079001=st01(+) and r079002=a0901(+) and id='" & strUserNum & "' group by r079001,r079002,r079024,st02,NVL(A0902,A0903) order by r079001,r079002 "
strSql = "select st02,NVL(A0902,A0903),sum(r079003),sum(r079004),sum(r079005),sum(r079006),sum(r079007),sum(r079008),sum(r079009),sum(r079010),sum(r079011),sum(r079012),sum(r079013),sum(r079014),sum(r079015),sum(r079016),sum(r079017),sum(r079018),sum(r079019),sum(r079020),sum(r079021),sum(r079022),sum(r079023),r079024,r079001,r079002,sum(r079026),sum(r079027),sum(r079028) from r020408_1,staff,acc090 where R079001=st01(+) and r079002=a0901(+) and id='" & strUserNum & "' group by r079001,r079002,r079024,st02,NVL(A0902,A0903) order by r079001,r079002 "
CheckOC
'Add By Cheng 2002/05/06
m_strPromoter = ""
SavDay1 = ""
SavDay2 = ""
SavDay3 = ""
SavDay5 = ""
SavDay6 = ""
'Page = 1
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        bolNotData = False 'Add By Sindy 2011/3/1
        .MoveFirst
        SavDay1 = CheckStr(.Fields(0))
        SavDay2 = CheckStr(.Fields(1))
        SavDay3 = CheckStr(.Fields(23))
        SavDay5 = CheckStr(.Fields(24))
        SavDay6 = CheckStr(.Fields(25))
        PrintTitle (1)
        PrintTitle1_1
        Do While .EOF = False
            For i = 0 To 23
                strTemp(i) = CheckStr(.Fields(i))
                If Val(strTemp(i)) = 0 And i <> 0 And i <> 1 And i <> 23 Then
                    strTemp(i) = "0"
                End If
            Next i
            strTemp(2) = Val("0" & .Fields(26).Value)
            strTemp(3) = Val("0" & .Fields(27).Value)
            strTemp(4) = Val("0" & .Fields(28).Value)
'            strTemp(0) = StrToStr(strTemp(0), 3)
'            strTemp(1) = StrToStr(strTemp(1), 4)
            strTemp(0) = StrToStr(strTemp(0), 5)
            strTemp(1) = StrToStr(strTemp(1), 5)
            If Val(strTemp(2)) + Val(strTemp(3)) = 0 Then
                strTemp(4) = "0"
            Else
                strTemp(4) = Trim(str(Val(strTemp(2)) / (Val(strTemp(2)) + Val(strTemp(3))) * 100))
            End If
            If SavDay1 <> strTemp(0) Then
                'Add By Cheng 2002/05/06
                m_strPromoter = ""
                
                ShowLine1_1
                PrintEnd1_1 (2)
                ShowLine1_1
                PrintEnd1_1 (3)
                ShowLine1_1
                PrintEnd1_1 (0)
                ShowLine1_1
                SavDay1 = strTemp(0)
                SavDay5 = CheckStr(.Fields(24))
                SavDay6 = CheckStr(.Fields(25))
            End If
            PrintDatil1_1
            If iPrint >= 13900 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle (1)
                PrintTitle1_1
            End If
            .MoveNext
        Loop
    Else
         GoTo ReadNext2: 'Add By Sindy 2011/3/1
    End If
End With
CheckOC
ShowLine1_1
PrintEnd1_1 (2)
ShowLine1_1
PrintEnd1_1 (3)
ShowLine1_1
PrintEnd1_1 (0)
ShowLine1_1
PrintEnd1_1 (1)
ShowLine1_1
Page = Page + 1
'Modified by Morgan 2015/7/2
'Printer.EndDoc
Printer.NewPage
'end 2015/7/2
ReadNext2: 'Add By Sindy 2011/3/1

'表(2)格式1
'nick 910521
'strSQL = "select r080001,r080002,sum(r080003),sum(r080004),sum(r080005),sum(r080006),sum(r080007),sum(r080008),sum(r080009),sum(r080010),sum(r080011),sum(r080012),sum(r080013),sum(r080014),sum(r080015),sum(r080016),sum(r080017),sum(r080018),sum(r080019),sum(r080020),sum(r080021),sum(r080022),sum(r080023),sum(r080024),sum(r080025),sum(r080026),sum(r080027),sum(r080028),sum(r080029),sum(r080030),sum(r080031),sum(r080032),r080033 from r020408_2 group by r080001,r080002,r080033"
'Modify By Cheng 2003/03/07
'strSQL = "select st02,NVL(A0902,A0903),sum(r080003),sum(r080004),sum(r080005),sum(r080006),sum(r080007),sum(r080008),sum(r080009),sum(r080010),sum(r080011),sum(r080012),sum(r080013),sum(r080014),sum(r080015),sum(r080016),sum(r080017),sum(r080018),sum(r080019),sum(r080020),sum(r080021),sum(r080022),sum(r080023),sum(r080024),sum(r080025),sum(r080026),sum(r080027),sum(r080028),sum(r080029),sum(r080030),sum(r080031),sum(r080032),r080033,r080001,r080002 from r020408_2,staff,acc090 where r080001=st01(+) and r080002=a0901(+) and id='" & strUserNum & "' group by r080001,r080002,r080033,st02,NVL(A0902,A0903) order by r080001,r080002 "
strSql = "select st02,NVL(A0902,A0903),sum(r080003),sum(r080004),sum(r080005),sum(r080006),sum(r080007),sum(r080008),sum(r080009),sum(r080010),sum(r080011),sum(r080012),sum(r080013),sum(r080014),sum(r080015),sum(r080016),sum(r080017),sum(r080018),sum(r080019),sum(r080020),sum(r080021),sum(r080022),sum(r080023),sum(r080024),sum(r080025),sum(r080026),sum(r080027),sum(r080028),sum(r080029),sum(r080030),sum(r080031),sum(r080032),r080033,r080001,r080002,sum(r080035),sum(r080036),sum(r080037),sum(r080038),sum(r080039),sum(r080040),sum(r080041),sum(r080042),sum(r080043),sum(r080044),sum(r080045),sum(r080046) from r020408_2,staff,acc090 where r080001=st01(+) and r080002=a0901(+) and id='" & strUserNum & "' group by r080001,r080002,r080033,st02,NVL(A0902,A0903) order by r080001,r080002 "
CheckOC
'Add By Cheng 2002/05/06
m_strPromoter = ""
SavDay1 = ""
SavDay2 = ""
SavDay3 = ""
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        bolNotData = False 'Add By Sindy 2011/3/1
        .MoveFirst
        SavDay1 = CheckStr(.Fields(0))
        SavDay2 = CheckStr(.Fields(1))
        SavDay3 = CheckStr(.Fields(32))
        SavDay5 = CheckStr(.Fields(33))
        SavDay6 = CheckStr(.Fields(34))
        PrintTitle (2)
        PrintTitle2
        Do While .EOF = False
            For i = 0 To 32
                strTemp(i) = CheckStr(.Fields(i))
                If Val(strTemp(i)) = 0 And i <> 0 And i <> 1 And i <> 32 Then
                    strTemp(i) = "0"
                End If
            Next i
            'Add By Cheng 2003/03/07
            strTemp(26) = Val("0" & .Fields(35))
            strTemp(27) = Val("0" & .Fields(36))
            strTemp(28) = Val("0" & .Fields(37))
            strTemp(29) = Val("0" & .Fields(38))
            strTemp(30) = Val("0" & .Fields(39))
            strTemp(31) = Val("0" & .Fields(40))
'            strTemp(0) = StrToStr(strTemp(0), 3)
'            strTemp(1) = StrToStr(strTemp(1), 4)
            strTemp(0) = StrToStr(strTemp(0), 5)
            strTemp(1) = StrToStr(strTemp(1), 5)
            If Val(strTemp(2)) + Val(strTemp(3)) = 0 Then
                strTemp(6) = "0"
            Else
                strTemp(6) = Trim(str(Val(strTemp(2)) / (Val(strTemp(2)) + Val(strTemp(3))) * 100))
            End If
            If Val(strTemp(4)) + Val(strTemp(5)) = 0 Then
                strTemp(7) = "0"
            Else
                strTemp(7) = Trim(str(Val(strTemp(4)) / (Val(strTemp(4)) + Val(strTemp(5))) * 100))
            End If
            If Val(strTemp(8)) + Val(strTemp(9)) = 0 Then
                strTemp(12) = "0"
            Else
                strTemp(12) = Trim(str(Val(strTemp(8)) / (Val(strTemp(8)) + Val(strTemp(9))) * 100))
            End If
            If Val(strTemp(10)) + Val(strTemp(11)) = 0 Then
                strTemp(13) = "0"
            Else
                strTemp(13) = Trim(str(Val(strTemp(10)) / (Val(strTemp(10)) + Val(strTemp(11))) * 100))
            End If
            If Val(strTemp(14)) + Val(strTemp(15)) = 0 Then
                strTemp(18) = "0"
            Else
                strTemp(18) = Trim(str(Val(strTemp(14)) / (Val(strTemp(14)) + Val(strTemp(15))) * 100))
            End If
            If Val(strTemp(16)) + Val(strTemp(17)) = 0 Then
                strTemp(19) = "0"
            Else
                strTemp(19) = Trim(str(Val(strTemp(16)) / (Val(strTemp(16)) + Val(strTemp(17))) * 100))
            End If
            If Val(strTemp(20)) + Val(strTemp(21)) = 0 Then
                strTemp(24) = "0"
            Else
                strTemp(24) = Trim(str(Val(strTemp(20)) / (Val(strTemp(20)) + Val(strTemp(21))) * 100))
            End If
            If Val(strTemp(22)) + Val(strTemp(23)) = 0 Then
                strTemp(25) = "0"
            Else
                strTemp(25) = Trim(str(Val(strTemp(22)) / (Val(strTemp(22)) + Val(strTemp(23))) * 100))
            End If
            If Val(strTemp(26)) + Val(strTemp(27)) = 0 Then
                strTemp(30) = "0"
            Else
                strTemp(30) = Trim(str(Val(strTemp(26)) / (Val(strTemp(26)) + Val(strTemp(27))) * 100))
            End If
            If Val(strTemp(28)) + Val(strTemp(29)) = 0 Then
                strTemp(31) = "0"
            Else
                strTemp(31) = Trim(str(Val(strTemp(28)) / (Val(strTemp(28)) + Val(strTemp(29))) * 100))
            End If
            If SavDay1 <> strTemp(0) Then
                'Add By Cheng 2002/05/06
                m_strPromoter = ""
                
                ShowLine2
                PrintEnd2 (2)
                ShowLine2
                PrintEnd2 (3)
                ShowLine2
                PrintEnd2 (0)
                ShowLine2
                SavDay1 = strTemp(0)
               SavDay5 = CheckStr(.Fields(33))
               SavDay6 = CheckStr(.Fields(34))
            End If
            PrintDatil2
            If iPrint >= 13900 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle (2)
                PrintTitle2
            End If
            .MoveNext
        Loop
    Else
         GoTo ReadNext3: 'Add By Sindy 2011/3/1
    End If
End With
ShowLine2
PrintEnd2 (2)
ShowLine2
PrintEnd2 (3)
ShowLine2
PrintEnd2 (0)
ShowLine2
PrintEnd2 (1)
ShowLine2
Page = Page + 1
Printer.NewPage
ReadNext3: 'Add By Sindy 2011/3/1

'表(2)格式2
'nick 910521
'strSQL = "select r080001,r080002,sum(r080003),sum(r080004),sum(r080005),sum(r080006),sum(r080007),sum(r080008),sum(r080009),sum(r080010),sum(r080011),sum(r080012),sum(r080013),sum(r080014),sum(r080015),sum(r080016),sum(r080017),sum(r080018),sum(r080019),sum(r080020),sum(r080021),sum(r080022),sum(r080023),sum(r080024),sum(r080025),sum(r080026),sum(r080027),sum(r080028),sum(r080029),sum(r080030),sum(r080031),sum(r080032),r080033 from r020408_2 group by r080001,r080002,r080033"
'Modify By Cheng 2003/03/07
'strSQL = "select st02,NVL(A0902,A0903),sum(r080003),sum(r080004),sum(r080005),sum(r080006),sum(r080007),sum(r080008),sum(r080009),sum(r080010),sum(r080011),sum(r080012),sum(r080013),sum(r080014),sum(r080015),sum(r080016),sum(r080017),sum(r080018),sum(r080019),sum(r080020),sum(r080021),sum(r080022),sum(r080023),sum(r080024),sum(r080025),sum(r080026),sum(r080027),sum(r080028),sum(r080029),sum(r080030),sum(r080031),sum(r080032),r080033,r080001,r080002 from r020408_2,staff,acc090 where r080001=st01(+) and r080002=a0901(+) and id='" & strUserNum & "' group by r080001,r080002,r080033,st02,NVL(A0902,A0903) order by r080001,r080002 "
strSql = "select st02,NVL(A0902,A0903),sum(r080003),sum(r080004),sum(r080005),sum(r080006),sum(r080007),sum(r080008),sum(r080009),sum(r080010),sum(r080011),sum(r080012),sum(r080013),sum(r080014),sum(r080015),sum(r080016),sum(r080017),sum(r080018),sum(r080019),sum(r080020),sum(r080021),sum(r080022),sum(r080023),sum(r080024),sum(r080025),sum(r080026),sum(r080027),sum(r080028),sum(r080029),sum(r080030),sum(r080031),sum(r080032),r080033,r080001,r080002,sum(r080035),sum(r080036),sum(r080037),sum(r080038),sum(r080039),sum(r080040),sum(r080041),sum(r080042),sum(r080043),sum(r080044),sum(r080045),sum(r080046) from r020408_2,staff,acc090 where r080001=st01(+) and r080002=a0901(+) and id='" & strUserNum & "' group by r080001,r080002,r080033,st02,NVL(A0902,A0903) order by r080001,r080002 "
CheckOC
'Add By Cheng 2002/05/06
m_strPromoter = ""
SavDay1 = ""
SavDay2 = ""
SavDay3 = ""
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        bolNotData = False 'Add By Sindy 2011/3/1
        .MoveFirst
        SavDay1 = CheckStr(.Fields(0))
        SavDay2 = CheckStr(.Fields(1))
        SavDay3 = CheckStr(.Fields(32))
        SavDay5 = CheckStr(.Fields(33))
        SavDay6 = CheckStr(.Fields(34))
        PrintTitle (2)
        PrintTitle2_1
        Do While .EOF = False
            For i = 0 To 32
                strTemp(i) = CheckStr(.Fields(i))
                If Val(strTemp(i)) = 0 And i <> 0 And i <> 1 And i <> 32 Then
                    strTemp(i) = "0"
                End If
            Next i
            'Add By Cheng 2003/03/07
            strTemp(2) = CheckStr(.Fields(41))
            strTemp(3) = CheckStr(.Fields(42))
            strTemp(4) = CheckStr(.Fields(43))
            strTemp(5) = CheckStr(.Fields(44))
            strTemp(6) = CheckStr(.Fields(45))
            strTemp(7) = CheckStr(.Fields(46))
            strTemp(8) = CheckStr(.Fields(26))
            strTemp(9) = CheckStr(.Fields(27))
            strTemp(10) = CheckStr(.Fields(28))
            strTemp(11) = CheckStr(.Fields(29))
            strTemp(12) = CheckStr(.Fields(30))
            strTemp(13) = CheckStr(.Fields(31))
            
'            strTemp(0) = StrToStr(strTemp(0), 3)
'            strTemp(1) = StrToStr(strTemp(1), 4)
            strTemp(0) = StrToStr(strTemp(0), 5)
            strTemp(1) = StrToStr(strTemp(1), 5)
            If Val(strTemp(2)) + Val(strTemp(3)) = 0 Then
                strTemp(6) = "0"
            Else
                strTemp(6) = Trim(str(Val(strTemp(2)) / (Val(strTemp(2)) + Val(strTemp(3))) * 100))
            End If
            If Val(strTemp(4)) + Val(strTemp(5)) = 0 Then
                strTemp(7) = "0"
            Else
                strTemp(7) = Trim(str(Val(strTemp(4)) / (Val(strTemp(4)) + Val(strTemp(5))) * 100))
            End If
            If Val(strTemp(8)) + Val(strTemp(9)) = 0 Then
                strTemp(12) = "0"
            Else
                strTemp(12) = Trim(str(Val(strTemp(8)) / (Val(strTemp(8)) + Val(strTemp(9))) * 100))
            End If
            If Val(strTemp(10)) + Val(strTemp(11)) = 0 Then
                strTemp(13) = "0"
            Else
                strTemp(13) = Trim(str(Val(strTemp(10)) / (Val(strTemp(10)) + Val(strTemp(11))) * 100))
            End If
            If SavDay1 <> strTemp(0) Then
                'Add By Cheng 2002/05/06
                m_strPromoter = ""
                
                ShowLine2_1
                PrintEnd2_1 (2)
                ShowLine2_1
                PrintEnd2_1 (3)
                ShowLine2_1
                PrintEnd2_1 (0)
                ShowLine2_1
                SavDay1 = strTemp(0)
               SavDay5 = CheckStr(.Fields(33))
               SavDay6 = CheckStr(.Fields(34))
            End If
            PrintDatil2_1
            If iPrint >= 13900 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle (2)
                PrintTitle2_1
            End If
            .MoveNext
        Loop
    Else
         GoTo ReadNext4: 'Add By Sindy 2011/3/1
    End If
End With
ShowLine2_1
PrintEnd2_1 (2)
ShowLine2_1
PrintEnd2_1 (3)
ShowLine2_1
PrintEnd2_1 (0)
ShowLine2_1
PrintEnd2_1 (1)
ShowLine2_1
Page = Page + 1
ReadNext4: 'Add By Sindy 2011/3/1
If bolNotData = False Then Printer.EndDoc
End Sub

Sub PrintEnd1(Strindex As Integer)
Select Case Strindex
Case 0
     strSql = "select '個人小計','',SUM(r079003),sum(r079004),sum(r079005),sum(r079006),sum(r079007),sum(r079008),sum(r079009),sum(r079010),sum(r079011),sum(r079012),sum(r079013),sum(r079014),sum(r079015),sum(r079016),sum(r079017),sum(r079018),sum(r079019),sum(r079020),sum(r079021),sum(r079022),sum(r079023) from r020408_1 where id='" & strUserNum & "' AND r079001='" & SavDay5 & "' "
Case 1
     strSql = "select '全所總計','',SUM(r079003),sum(r079004),sum(r079005),sum(r079006),sum(r079007),sum(r079008),sum(r079009),sum(r079010),sum(r079011),sum(r079012),sum(r079013),sum(r079014),sum(r079015),sum(r079016),sum(r079017),sum(r079018),sum(r079019),sum(r079020),sum(r079021),sum(r079022),sum(r079023) from r020408_1 where id='" & strUserNum & "' "
Case 2
     'edit by nickc 2005/05/13
     'StrSql = "select '國內小計','',SUM(r079003),sum(r079004),sum(r079005),sum(r079006),sum(r079007),sum(r079008),sum(r079009),sum(r079010),sum(r079011),sum(r079012),sum(r079013),sum(r079014),sum(r079015),sum(r079016),sum(r079017),sum(r079018),sum(r079019),sum(r079020),sum(r079021),sum(r079022),sum(r079023) from r020408_1 where id='" & strUserNum & "' AND r079001='" & SavDay5 & "' and r079025='*' "
     strSql = "select '國內業務小計','',SUM(r079003),sum(r079004),sum(r079005),sum(r079006),sum(r079007),sum(r079008),sum(r079009),sum(r079010),sum(r079011),sum(r079012),sum(r079013),sum(r079014),sum(r079015),sum(r079016),sum(r079017),sum(r079018),sum(r079019),sum(r079020),sum(r079021),sum(r079022),sum(r079023) from r020408_1 where id='" & strUserNum & "' AND r079001='" & SavDay5 & "' and r079025='*' "
Case 3
     'edit by nickc 2005/05/13
     'StrSql = "select '國外小計','',SUM(r079003),sum(r079004),sum(r079005),sum(r079006),sum(r079007),sum(r079008),sum(r079009),sum(r079010),sum(r079011),sum(r079012),sum(r079013),sum(r079014),sum(r079015),sum(r079016),sum(r079017),sum(r079018),sum(r079019),sum(r079020),sum(r079021),sum(r079022),sum(r079023) from r020408_1 where id='" & strUserNum & "' AND r079001='" & SavDay5 & "' and (r079025='' or r079025 is null ) "
     strSql = "select '國外業務小計','',SUM(r079003),sum(r079004),sum(r079005),sum(r079006),sum(r079007),sum(r079008),sum(r079009),sum(r079010),sum(r079011),sum(r079012),sum(r079013),sum(r079014),sum(r079015),sum(r079016),sum(r079017),sum(r079018),sum(r079019),sum(r079020),sum(r079021),sum(r079022),sum(r079023) from r020408_1 where id='" & strUserNum & "' AND r079001='" & SavDay5 & "' and (r079025='' or r079025 is null ) "
     BolEndThisPage = True
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
            For i = 2 To 20 Step 3
                Printer.Line (PLeft(i) - 50, iPrint - 150)-(PLeft(i) - 50, iPrint + 450)
            Next i
            For i = 0 To 22
                StrTemp7(i) = CheckStr(.Fields(i))
                If Val(StrTemp7(i)) = 0 And i <> 0 And i <> 1 Then
                    StrTemp7(i) = "0"
                End If
            Next i
            If Val(StrTemp7(2)) + Val(StrTemp7(3)) = 0 Then
                StrTemp7(4) = "0"
            Else
                StrTemp7(4) = Trim(str(Val(StrTemp7(2)) / (Val(StrTemp7(2)) + Val(StrTemp7(3))) * 100))
            End If
            If Val(StrTemp7(5)) + Val(StrTemp7(6)) = 0 Then
                StrTemp7(7) = "0"
            Else
                StrTemp7(7) = Trim(str(Val(StrTemp7(5)) / (Val(StrTemp7(5)) + Val(StrTemp7(6))) * 100))
            End If
            If Val(StrTemp7(8)) + Val(StrTemp7(9)) = 0 Then
                StrTemp7(10) = "0"
            Else
                StrTemp7(10) = Trim(str(Val(StrTemp7(8)) / (Val(StrTemp7(8)) + Val(StrTemp7(9))) * 100))
            End If
            If Val(StrTemp7(11)) + Val(StrTemp7(12)) = 0 Then
                StrTemp7(13) = "0"
            Else
                StrTemp7(13) = Trim(str(Val(StrTemp7(11)) / (Val(StrTemp7(11)) + Val(StrTemp7(12))) * 100))
            End If
            If Val(StrTemp7(14)) + Val(StrTemp7(15)) = 0 Then
                StrTemp7(16) = "0"
            Else
                StrTemp7(16) = Trim(str(Val(StrTemp7(14)) / (Val(StrTemp7(14)) + Val(StrTemp7(15))) * 100))
            End If
            If Val(StrTemp7(17)) + Val(StrTemp7(18)) = 0 Then
                StrTemp7(19) = "0"
            Else
                StrTemp7(19) = Trim(str(Val(StrTemp7(17)) / (Val(StrTemp7(17)) + Val(StrTemp7(18))) * 100))
            End If
            If Val(StrTemp7(20)) + Val(StrTemp7(21)) = 0 Then
                StrTemp7(22) = "0"
            Else
                StrTemp7(22) = Trim(str(Val(StrTemp7(20)) / (Val(StrTemp7(20)) + Val(StrTemp7(21))) * 100))
            End If
            
            
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(1)
            For i = 2 To 22
                'add by nickc 2007/07/27 商標處改格式
                If txt1(3) = "020" And txt1(4) = "020" And i > 19 Then
                Else
                    Select Case i
                    Case 4, 7, 10, 13, 16, 19, 22
                        Printer.CurrentX = PLeft(i) + 600 - Printer.TextWidth(Format(StrTemp7(i), "##0.00"))
                        Printer.CurrentY = iPrint
                        Printer.Print Format(StrTemp7(i), "##0.00") & "%"
                    Case Else
                        Printer.CurrentX = PLeft(i) + 400 - Printer.TextWidth(Format(StrTemp7(i), "####0"))
                        Printer.CurrentY = iPrint
                        Printer.Print Format(StrTemp7(i), "####0")
                    End Select
                End If
            Next i
            iPrint = iPrint + 300
            If iPrint >= 14000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle (1)
                PrintTitle1
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC2
End Sub

'Add By Cheng 2003/03/07
Sub PrintEnd1_1(Strindex As Integer)
Select Case Strindex
Case 0
     strSql = "select '個人小計','',SUM(r079026),sum(r079027),sum(r079028) from r020408_1 where id='" & strUserNum & "' AND r079001='" & SavDay5 & "' "
Case 1
     strSql = "select '全所總計','',SUM(r079026),sum(r079027),sum(r079028) from r020408_1 where id='" & strUserNum & "' "
Case 2
     'edit by nickc 2005/05/13
     'StrSql = "select '國內小計','',SUM(r079026),sum(r079027),sum(r079028) from r020408_1 where id='" & strUserNum & "' AND r079001='" & SavDay5 & "' and r079025='*' "
     strSql = "select '國內業務小計','',SUM(r079026),sum(r079027),sum(r079028) from r020408_1 where id='" & strUserNum & "' AND r079001='" & SavDay5 & "' and r079025='*' "
Case 3
     'edit by nickc 2005/05/13
     'StrSql = "select '國外小計','',SUM(r079026),sum(r079027),sum(r079028) from r020408_1 where id='" & strUserNum & "' AND r079001='" & SavDay5 & "' and (r079025='' or r079025 is null ) "
     strSql = "select '國外業務小計','',SUM(r079026),sum(r079027),sum(r079028) from r020408_1 where id='" & strUserNum & "' AND r079001='" & SavDay5 & "' and (r079025='' or r079025 is null ) "
     BolEndThisPage = True
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
            For i = 2 To 2 Step 3
                Printer.Line (PLeft(i) - 50, iPrint - 150)-(PLeft(i) - 50, iPrint + 450)
            Next i
            For i = 0 To 4
                StrTemp7(i) = CheckStr(.Fields(i))
                If Val(StrTemp7(i)) = 0 And i <> 0 And i <> 1 Then
                    StrTemp7(i) = "0"
                End If
            Next i
            If Val(StrTemp7(2)) + Val(StrTemp7(3)) = 0 Then
                StrTemp7(4) = "0"
            Else
                StrTemp7(4) = Trim(str(Val(StrTemp7(2)) / (Val(StrTemp7(2)) + Val(StrTemp7(3))) * 100))
            End If
            
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(1)
            For i = 2 To 4
                Select Case i
                Case 4, 7, 10, 13, 16, 19, 22
                    Printer.CurrentX = PLeft(i) + 600 - Printer.TextWidth(Format(StrTemp7(i), "##0.00"))
                    Printer.CurrentY = iPrint
                    Printer.Print Format(StrTemp7(i), "##0.00") & "%"
                Case Else
                    Printer.CurrentX = PLeft(i) + 400 - Printer.TextWidth(Format(StrTemp7(i), "####0"))
                    Printer.CurrentY = iPrint
                    Printer.Print Format(StrTemp7(i), "####0")
                End Select
            Next i
            iPrint = iPrint + 300
            If iPrint >= 14000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle (1)
                PrintTitle1_1
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC2
End Sub

Sub PrintEnd2(Strindex As Integer)
Select Case Strindex
Case 0
    'Modify By Cheng 2003/03/07
'     strSQL = "select '個人小計','',sum(r080003),sum(r080004),sum(r080005),sum(r080006),sum(r080007),sum(r080008),sum(r080009),sum(r080010),sum(r080011),sum(r080012),sum(r080013),sum(r080014),sum(r080015),sum(r080016),sum(r080017),sum(r080018),sum(r080019),sum(r080020),sum(r080021),sum(r080022),sum(r080023),sum(r080024),sum(r080025),sum(r080026),sum(r080027),sum(r080028),sum(r080029),sum(r080030),sum(r080031),sum(r080032),r080033,r080001 from r020408_2 WHERE ID='" & strUserNum & "' AND r080001='" & SavDay5 & "' group by r080033,r080001"
     strSql = "select '個人小計','',sum(r080003),sum(r080004),sum(r080005),sum(r080006),sum(r080007),sum(r080008),sum(r080009),sum(r080010),sum(r080011),sum(r080012),sum(r080013),sum(r080014),sum(r080015),sum(r080016),sum(r080017),sum(r080018),sum(r080019),sum(r080020),sum(r080021),sum(r080022),sum(r080023),sum(r080024),sum(r080025),sum(r080026),sum(r080035),sum(r080036),sum(r080037),sum(r080038),sum(r080039),sum(r080040),r080033,r080001 from r020408_2 WHERE ID='" & strUserNum & "' AND r080001='" & SavDay5 & "' group by r080033,r080001"
Case 1
    'Modify By Cheng 2003/03/07
'     strSQL = "select '全所總計','',sum(r080003),sum(r080004),sum(r080005),sum(r080006),sum(r080007),sum(r080008),sum(r080009),sum(r080010),sum(r080011),sum(r080012),sum(r080013),sum(r080014),sum(r080015),sum(r080016),sum(r080017),sum(r080018),sum(r080019),sum(r080020),sum(r080021),sum(r080022),sum(r080023),sum(r080024),sum(r080025),sum(r080026),sum(r080027),sum(r080028),sum(r080029),sum(r080030),sum(r080031),sum(r080032) from r020408_2 WHERE ID='" & strUserNum & "' "
     strSql = "select '全所總計','',sum(r080003),sum(r080004),sum(r080005),sum(r080006),sum(r080007),sum(r080008),sum(r080009),sum(r080010),sum(r080011),sum(r080012),sum(r080013),sum(r080014),sum(r080015),sum(r080016),sum(r080017),sum(r080018),sum(r080019),sum(r080020),sum(r080021),sum(r080022),sum(r080023),sum(r080024),sum(r080025),sum(r080026),sum(r080035),sum(r080036),sum(r080037),sum(r080038),sum(r080039),sum(r080040) from r020408_2 WHERE ID='" & strUserNum & "' "
Case 2
    'Modify By Cheng 2003/03/07
'     strSQL = "select '國內小計','',sum(r080003),sum(r080004),sum(r080005),sum(r080006),sum(r080007),sum(r080008),sum(r080009),sum(r080010),sum(r080011),sum(r080012),sum(r080013),sum(r080014),sum(r080015),sum(r080016),sum(r080017),sum(r080018),sum(r080019),sum(r080020),sum(r080021),sum(r080022),sum(r080023),sum(r080024),sum(r080025),sum(r080026),sum(r080027),sum(r080028),sum(r080029),sum(r080030),sum(r080031),sum(r080032) from r020408_2 WHERE ID='" & strUserNum & "' AND r080001='" & SavDay5 & "' and r080034='*' "
     'edit by nickc 2005/05/13
     'StrSql = "select '國內小計','',sum(r080003),sum(r080004),sum(r080005),sum(r080006),sum(r080007),sum(r080008),sum(r080009),sum(r080010),sum(r080011),sum(r080012),sum(r080013),sum(r080014),sum(r080015),sum(r080016),sum(r080017),sum(r080018),sum(r080019),sum(r080020),sum(r080021),sum(r080022),sum(r080023),sum(r080024),sum(r080025),sum(r080026),sum(r080035),sum(r080036),sum(r080037),sum(r080038),sum(r080039),sum(r080040) from r020408_2 WHERE ID='" & strUserNum & "' AND r080001='" & SavDay5 & "' and r080034='*' "
     strSql = "select '國內業務小計','',sum(r080003),sum(r080004),sum(r080005),sum(r080006),sum(r080007),sum(r080008),sum(r080009),sum(r080010),sum(r080011),sum(r080012),sum(r080013),sum(r080014),sum(r080015),sum(r080016),sum(r080017),sum(r080018),sum(r080019),sum(r080020),sum(r080021),sum(r080022),sum(r080023),sum(r080024),sum(r080025),sum(r080026),sum(r080035),sum(r080036),sum(r080037),sum(r080038),sum(r080039),sum(r080040) from r020408_2 WHERE ID='" & strUserNum & "' AND r080001='" & SavDay5 & "' and r080034='*' "
Case 3
    'Modify By Cheng 2003/03/07
'     strSQL = "select '國外小計','',sum(r080003),sum(r080004),sum(r080005),sum(r080006),sum(r080007),sum(r080008),sum(r080009),sum(r080010),sum(r080011),sum(r080012),sum(r080013),sum(r080014),sum(r080015),sum(r080016),sum(r080017),sum(r080018),sum(r080019),sum(r080020),sum(r080021),sum(r080022),sum(r080023),sum(r080024),sum(r080025),sum(r080026),sum(r080027),sum(r080028),sum(r080029),sum(r080030),sum(r080031),sum(r080032) from r020408_2 WHERE ID='" & strUserNum & "' AND r080001='" & SavDay5 & "' and (r080034='' or r080034 is null) "
     'edit by nickc 2005/05/13
     'StrSql = "select '國外小計','',sum(r080003),sum(r080004),sum(r080005),sum(r080006),sum(r080007),sum(r080008),sum(r080009),sum(r080010),sum(r080011),sum(r080012),sum(r080013),sum(r080014),sum(r080015),sum(r080016),sum(r080017),sum(r080018),sum(r080019),sum(r080020),sum(r080021),sum(r080022),sum(r080023),sum(r080024),sum(r080025),sum(r080026),sum(r080035),sum(r080036),sum(r080037),sum(r080038),sum(r080039),sum(r080040) from r020408_2 WHERE ID='" & strUserNum & "' AND r080001='" & SavDay5 & "' and (r080034='' or r080034 is null) "
     strSql = "select '國外業務小計','',sum(r080003),sum(r080004),sum(r080005),sum(r080006),sum(r080007),sum(r080008),sum(r080009),sum(r080010),sum(r080011),sum(r080012),sum(r080013),sum(r080014),sum(r080015),sum(r080016),sum(r080017),sum(r080018),sum(r080019),sum(r080020),sum(r080021),sum(r080022),sum(r080023),sum(r080024),sum(r080025),sum(r080026),sum(r080035),sum(r080036),sum(r080037),sum(r080038),sum(r080039),sum(r080040) from r020408_2 WHERE ID='" & strUserNum & "' AND r080001='" & SavDay5 & "' and (r080034='' or r080034 is null) "
     BolEndThisPage = True
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
            For i = 2 To 26 Step 6
                Printer.Line (PLeft(i) - 50, iPrint - 150)-(PLeft(i) - 50, iPrint + 450)
            Next i
            For i = 0 To 31
                StrTemp7(i) = CheckStr(.Fields(i))
                If Val(StrTemp7(i)) = 0 And i <> 0 And i <> 1 Then
                    StrTemp7(i) = "0"
                End If
            Next i
            If Val(StrTemp7(2)) + Val(StrTemp7(3)) = 0 Then
                StrTemp7(6) = "0"
            Else
                StrTemp7(6) = Trim(str(Val(StrTemp7(2)) / (Val(StrTemp7(2)) + Val(StrTemp7(3))) * 100))
            End If
            If Val(StrTemp7(4)) + Val(StrTemp7(5)) = 0 Then
                StrTemp7(7) = "0"
            Else
                StrTemp7(7) = Trim(str(Val(StrTemp7(4)) / (Val(StrTemp7(4)) + Val(StrTemp7(5))) * 100))
            End If
            If Val(StrTemp7(8)) + Val(StrTemp7(9)) = 0 Then
                StrTemp7(12) = "0"
            Else
                StrTemp7(12) = Trim(str(Val(StrTemp7(8)) / (Val(StrTemp7(8)) + Val(StrTemp7(9))) * 100))
            End If
            If Val(StrTemp7(10)) + Val(StrTemp7(11)) = 0 Then
                StrTemp7(13) = "0"
            Else
                StrTemp7(13) = Trim(str(Val(StrTemp7(10)) / (Val(StrTemp7(10)) + Val(StrTemp7(11))) * 100))
            End If
            If Val(StrTemp7(14)) + Val(StrTemp7(15)) = 0 Then
                StrTemp7(18) = "0"
            Else
                StrTemp7(18) = Trim(str(Val(StrTemp7(14)) / (Val(StrTemp7(14)) + Val(StrTemp7(15))) * 100))
            End If
            If Val(StrTemp7(16)) + Val(StrTemp7(17)) = 0 Then
                StrTemp7(19) = "0"
            Else
                StrTemp7(19) = Trim(str(Val(StrTemp7(16)) / (Val(StrTemp7(16)) + Val(StrTemp7(17))) * 100))
            End If
            If Val(StrTemp7(20)) + Val(StrTemp7(21)) = 0 Then
                StrTemp7(24) = "0"
            Else
                StrTemp7(24) = Trim(str(Val(StrTemp7(20)) / (Val(StrTemp7(20)) + Val(StrTemp7(21))) * 100))
            End If
            If Val(StrTemp7(22)) + Val(StrTemp7(23)) = 0 Then
                StrTemp7(25) = "0"
            Else
                StrTemp7(25) = Trim(str(Val(StrTemp7(22)) / (Val(StrTemp7(22)) + Val(StrTemp7(23))) * 100))
            End If
            If Val(StrTemp7(26)) + Val(StrTemp7(27)) = 0 Then
                StrTemp7(30) = "0"
            Else
                StrTemp7(30) = Trim(str(Val(StrTemp7(26)) / (Val(StrTemp7(26)) + Val(StrTemp7(27))) * 100))
            End If
            If Val(StrTemp7(28)) + Val(StrTemp7(29)) = 0 Then
                StrTemp7(31) = "0"
            Else
                StrTemp7(31) = Trim(str(Val(StrTemp7(28)) / (Val(StrTemp7(28)) + Val(StrTemp7(29))) * 100))
            End If
            'Modify By Cheng 2003/03/07
'            'update 個人目標資料檔
'            If TestOk = True And Strindex = 0 Then
'                '檢查有沒有存在
'                Set adoRecordset99 = New ADODB.Recordset
'                strSQL = "select * from performance where pe01='" & CheckStr(.Fields(33)) & "' and pe02='T' and pe03=" & Mid(txt1(1) + 19110000, 1, 6) & " "
'                adoRecordset99.CursorLocation = adUseClient
'                adoRecordset99.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'                If Not adoRecordset99.EOF And Not adoRecordset99.BOF Then
'                    cnnConnection.Execute "update performance set pe18=" & Val(StrTemp7(30)) & ",pe19=" & Val(StrTemp7(31)) & " where pe01='" & CheckStr(.Fields(33)) & "' and pe02='T' and pe03=" & Mid(txt1(1) + 19110000, 1, 6) & " "
'                Else
'                    cnnConnection.Execute "insert into performance (pe01,pe02,pe03,pe18,pe19) values ('" & CheckStr(.Fields(33)) & "','T'," & Mid(txt1(1) + 19110000, 1, 6) & "," & Val(StrTemp7(30)) & "," & Val(StrTemp7(31)) & ") "
'                End If
'            End If
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(1)
            For i = 2 To 31
                Select Case i
                Case 6, 12, 18, 24, 30
                     Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(StrTemp7(i), "##0.00"))
                     Printer.CurrentY = iPrint
                     Printer.Print Format(StrTemp7(i), "##0.00") & "%"
                Case 7, 13, 19, 25, 31
                     Printer.CurrentX = PLeft(i) + 400 - Printer.TextWidth(Format(StrTemp7(i), "##0.00"))
                     Printer.CurrentY = iPrint
                     Printer.Print Format(StrTemp7(i), "##0.00") & "%"
                Case Else
                     Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(StrTemp7(i), "####0"))
                     Printer.CurrentY = iPrint
                     Printer.Print Format(StrTemp7(i), "####0")
                End Select
            Next i
            iPrint = iPrint + 300
            If iPrint >= 14000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle (2)
                PrintTitle2
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC2
End Sub

'Add By Cheng 2003/03/07
Sub PrintEnd2_1(Strindex As Integer)
Select Case Strindex
Case 0
     strSql = "select '個人小計','',sum(r080041),sum(r080042),sum(r080043),sum(r080044),sum(r080045),sum(r080046),sum(r080027),sum(r080028),sum(r080029),sum(r080030),sum(r080031),sum(r080032),r080033,r080001 from r020408_2 WHERE ID='" & strUserNum & "' AND r080001='" & SavDay5 & "' group by r080033,r080001"
Case 1
     strSql = "select '全所總計','',sum(r080041),sum(r080042),sum(r080043),sum(r080044),sum(r080045),sum(r080046),sum(r080027),sum(r080028),sum(r080029),sum(r080030),sum(r080031),sum(r080032) from r020408_2 WHERE ID='" & strUserNum & "' "
Case 2
     'edit by nickc 2005/05/13
     'StrSql = "select '國內小計','',sum(r080041),sum(r080042),sum(r080043),sum(r080044),sum(r080045),sum(r080046),sum(r080027),sum(r080028),sum(r080029),sum(r080030),sum(r080031),sum(r080032) from r020408_2 WHERE ID='" & strUserNum & "' AND r080001='" & SavDay5 & "' and r080034='*' "
     strSql = "select '國內業務小計','',sum(r080041),sum(r080042),sum(r080043),sum(r080044),sum(r080045),sum(r080046),sum(r080027),sum(r080028),sum(r080029),sum(r080030),sum(r080031),sum(r080032) from r020408_2 WHERE ID='" & strUserNum & "' AND r080001='" & SavDay5 & "' and r080034='*' "
Case 3
     'edit by nickc 2005/05/13
     'StrSql = "select '國外小計','',sum(r080041),sum(r080042),sum(r080043),sum(r080044),sum(r080045),sum(r080046),sum(r080027),sum(r080028),sum(r080029),sum(r080030),sum(r080031),sum(r080032) from r020408_2 WHERE ID='" & strUserNum & "' AND r080001='" & SavDay5 & "' and (r080034='' or r080034 is null) "
     strSql = "select '國外業務小計','',sum(r080041),sum(r080042),sum(r080043),sum(r080044),sum(r080045),sum(r080046),sum(r080027),sum(r080028),sum(r080029),sum(r080030),sum(r080031),sum(r080032) from r020408_2 WHERE ID='" & strUserNum & "' AND r080001='" & SavDay5 & "' and (r080034='' or r080034 is null) "
     BolEndThisPage = True
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
            For i = 2 To 13 Step 6
                Printer.Line (PLeft(i) - 50, iPrint - 150)-(PLeft(i) - 50, iPrint + 450)
            Next i
            For i = 0 To 13
                StrTemp7(i) = CheckStr(.Fields(i))
                If Val(StrTemp7(i)) = 0 And i <> 0 And i <> 1 Then
                    StrTemp7(i) = "0"
                End If
            Next i
            If Val(StrTemp7(2)) + Val(StrTemp7(3)) = 0 Then
                StrTemp7(6) = "0"
            Else
                StrTemp7(6) = Trim(str(Val(StrTemp7(2)) / (Val(StrTemp7(2)) + Val(StrTemp7(3))) * 100))
            End If
            If Val(StrTemp7(4)) + Val(StrTemp7(5)) = 0 Then
                StrTemp7(7) = "0"
            Else
                StrTemp7(7) = Trim(str(Val(StrTemp7(4)) / (Val(StrTemp7(4)) + Val(StrTemp7(5))) * 100))
            End If
            If Val(StrTemp7(8)) + Val(StrTemp7(9)) = 0 Then
                StrTemp7(12) = "0"
            Else
                StrTemp7(12) = Trim(str(Val(StrTemp7(8)) / (Val(StrTemp7(8)) + Val(StrTemp7(9))) * 100))
            End If
            If Val(StrTemp7(10)) + Val(StrTemp7(11)) = 0 Then
                StrTemp7(13) = "0"
            Else
                StrTemp7(13) = Trim(str(Val(StrTemp7(10)) / (Val(StrTemp7(10)) + Val(StrTemp7(11))) * 100))
            End If
            'update 個人目標資料檔
            '2012/4/6 MODIFY BY SONIA 只做台灣案的預估準確率
            'If TestOk = True And Strindex = 0 Then
            If TestOk = True And Strindex = 0 And txt1(3) = "000" And txt1(4) = "000" Then
                '檢查有沒有存在
                Set Adorecordset99 = New ADODB.Recordset
                strSql = "select * from performance where pe01='" & CheckStr(.Fields(15)) & "' and pe02='T' and pe03=" & Mid(txt1(1) + 19110000, 1, 6) & " "
                Adorecordset99.CursorLocation = adUseClient
                Adorecordset99.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If Not Adorecordset99.EOF And Not Adorecordset99.BOF Then
                    cnnConnection.Execute "update performance set pe18=" & Val(StrTemp7(12)) & ",pe19=" & Val(StrTemp7(13)) & " where pe01='" & CheckStr(.Fields(15)) & "' and pe02='T' and pe03=" & Mid(txt1(1) + 19110000, 1, 6) & " "
                Else
                    cnnConnection.Execute "insert into performance (pe01,pe02,pe03,pe18,pe19) values ('" & CheckStr(.Fields(15)) & "','T'," & Mid(txt1(1) + 19110000, 1, 6) & "," & Val(StrTemp7(12)) & "," & Val(StrTemp7(13)) & ") "
                End If
            End If
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(1)
            For i = 2 To 13
                Select Case i
                Case 6, 12, 18, 24, 30
                     Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(StrTemp7(i), "##0.00"))
                     Printer.CurrentY = iPrint
                     Printer.Print Format(StrTemp7(i), "##0.00") & "%"
                Case 7, 13, 19, 25, 31
                     Printer.CurrentX = PLeft(i) + 400 - Printer.TextWidth(Format(StrTemp7(i), "##0.00"))
                     Printer.CurrentY = iPrint
                     Printer.Print Format(StrTemp7(i), "##0.00") & "%"
                Case Else
                     Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(StrTemp7(i), "####0"))
                     Printer.CurrentY = iPrint
                     Printer.Print Format(StrTemp7(i), "####0")
                End Select
            Next i
            iPrint = iPrint + 300
            If iPrint >= 14000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle (2)
                PrintTitle2_1
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC2
End Sub

Sub PrintTitle(Strindex As String)
iPrint = 0
'Printer.Orientation = 1 'Removed by Morgan 2015/6/3
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "商爭案承辦人勝敗統計表(" & Trim(str(Strindex)) & ") "
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 7000
Printer.CurrentY = iPrint
Printer.Print "勝敗日期：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 16800
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300

'Add By Cheng 2002/02/07
Printer.CurrentX = 250
Printer.CurrentY = iPrint
Printer.Print "申請國家：" & Me.txt1(3).Text & " － " & Me.txt1(4).Text
Printer.CurrentX = 6750
Printer.CurrentY = iPrint
Printer.Print "系統類別：" & Me.txt1(0).Text

Printer.CurrentX = 16800
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
End Sub

Sub PrintTitle2()
GetPleft2
Printer.Font.Size = 8
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
For i = 2 To 26 Step 6
    Printer.Line (PLeft(i) - 50, iPrint + 150)-(PLeft(i) - 50, iPrint + 1750)
Next i
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle (2)
    PrintTitle2
    Exit Sub
End If
'add by nickc 2007/07/27 商標處改格式
If txt1(3) = "020" And txt1(4) = "020" Then
    Printer.CurrentX = PLeft1(1) - (Printer.TextWidth(Title_602) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_602 'Modify By Sindy 2015/2/4"異議答辯"
    Printer.CurrentX = PLeft1(2) - (Printer.TextWidth(Title_604) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_604 'Modify By Sindy 2015/2/4"裁定答辯"
    Printer.CurrentX = PLeft1(3) - (Printer.TextWidth(Title_606) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_606 'Modify By Sindy 2015/2/4"撤銷答辯"
    Printer.CurrentX = PLeft1(4) - (Printer.TextWidth("") / 2)
    Printer.CurrentY = iPrint
    Printer.Print "" 'Modify By Sindy 2015/2/4"註冊不當撤銷答辯"
    Printer.CurrentX = PLeft1(5) - (Printer.TextWidth(Title_406) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_406 'Modify By Sindy 2015/2/4"復審答辯"
Else
    Printer.CurrentX = PLeft1(1) - (Printer.TextWidth(Title_602) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_602 'Modify By Sindy 2015/2/4"異議答辯"
    Printer.CurrentX = PLeft1(2) - (Printer.TextWidth(Title_604) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_604 'Modify By Sindy 2015/2/4"評定答辯"
    Printer.CurrentX = PLeft1(3) - (Printer.TextWidth(Title_606) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_606 'Modify By Sindy 2015/2/4"廢止答辯"
    'Modify By Cheng 2003/03/11
    '改為參加訴願
    'Printer.CurrentX = PLeft1(4) - (Printer.TextWidth("申請意見書") / 2)
    Printer.CurrentX = PLeft1(4) - (Printer.TextWidth(Title_406) / 2)
    Printer.CurrentY = iPrint
    'Printer.Print "申請意見書"
    Printer.Print Title_406 'Modify By Sindy 2015/2/4"參加訴願"
    '改為參加訴訟
    'Printer.CurrentX = PLeft1(5) - (Printer.TextWidth("總計") / 2)
    Printer.CurrentX = PLeft1(5) - (Printer.TextWidth(Title_407) / 2)
    Printer.CurrentY = iPrint
    'Printer.Print "總計"
    Printer.Print Title_407 'Modify By Sindy 2015/2/4"參加訴訟"
End If
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle (2)
    PrintTitle2
    Exit Sub
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "業務區別"
For k = 2 To 26 Step 6
    Printer.CurrentX = PLeft(k)
    Printer.CurrentY = iPrint
    Printer.Print "勝"
    Printer.CurrentX = PLeft(k + 1)
    Printer.CurrentY = iPrint
    Printer.Print "敗"
    Printer.CurrentX = PLeft(k + 2)
    Printer.CurrentY = iPrint
    Printer.Print "答勝"
    Printer.CurrentX = PLeft(k + 3)
    Printer.CurrentY = iPrint
    Printer.Print "答敗"
    Printer.CurrentX = PLeft(k + 4)
    Printer.CurrentY = iPrint
    Printer.Print "勝訴率"
    Printer.CurrentX = PLeft(k + 5)
    Printer.CurrentY = iPrint
    Printer.Print "勝訴率"
Next k
iPrint = iPrint + 300
For k = 6 To 30 Step 6
    Printer.CurrentX = PLeft(k) + (Printer.TextWidth("勝訴率") / 2)
    Printer.CurrentY = iPrint
    Printer.Print "1"
    Printer.CurrentX = PLeft(k + 1) + (Printer.TextWidth("勝訴率") / 2)
    Printer.CurrentY = iPrint
    Printer.Print "2"
Next k
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle (2)
    PrintTitle2
    Exit Sub
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle (2)
    PrintTitle2
    Exit Sub
End If
End Sub

'Add By Cheng 2003/03/07
Sub PrintTitle2_1()
GetPleft2
Printer.Font.Size = 8
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000 / 11 * 5, iPrint + 150)
For i = 2 To 13 Step 6
    Printer.Line (PLeft(i) - 50, iPrint + 150)-(PLeft(i) - 50, iPrint + 1750)
Next i
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle (2)
    PrintTitle2_1
    Exit Sub
End If
Printer.CurrentX = PLeft1(1) - (Printer.TextWidth("上訴答辯") / 2)
Printer.CurrentY = iPrint
Printer.Print "上訴答辯"
Printer.CurrentX = PLeft1(2) - (Printer.TextWidth("總計") / 2)
Printer.CurrentY = iPrint
Printer.Print "總計"
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle (2)
    PrintTitle2_1
    Exit Sub
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "業務區別"
For k = 2 To 13 Step 6
    Printer.CurrentX = PLeft(k)
    Printer.CurrentY = iPrint
    Printer.Print "勝"
    Printer.CurrentX = PLeft(k + 1)
    Printer.CurrentY = iPrint
    Printer.Print "敗"
    Printer.CurrentX = PLeft(k + 2)
    Printer.CurrentY = iPrint
    Printer.Print "答勝"
    Printer.CurrentX = PLeft(k + 3)
    Printer.CurrentY = iPrint
    Printer.Print "答敗"
    Printer.CurrentX = PLeft(k + 4)
    Printer.CurrentY = iPrint
    Printer.Print "勝訴率"
    Printer.CurrentX = PLeft(k + 5)
    Printer.CurrentY = iPrint
    Printer.Print "勝訴率"
Next k
iPrint = iPrint + 300
For k = 6 To 13 Step 6
    Printer.CurrentX = PLeft(k) + (Printer.TextWidth("勝訴率") / 2)
    Printer.CurrentY = iPrint
    Printer.Print "1"
    Printer.CurrentX = PLeft(k + 1) + (Printer.TextWidth("勝訴率") / 2)
    Printer.CurrentY = iPrint
    Printer.Print "2"
Next k
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle (2)
    PrintTitle2_1
    Exit Sub
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000 / 11 * 5, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle (2)
    PrintTitle2_1
    Exit Sub
End If
End Sub

Sub PrintDatil2()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
'Modify By Cheng 2002/05/06
'若承辦人相同則不印
If m_strPromoter <> strTemp(0) Then
   Printer.Print strTemp(0)
   m_strPromoter = strTemp(0)
Else
   Printer.Print ""
End If

Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(1)
For i = 2 To 26 Step 6
    Printer.Line (PLeft(i) - 50, iPrint - 150)-(PLeft(i) - 50, iPrint + 450)
    Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(strTemp(i), "####0"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i), "####0")
    Printer.CurrentX = PLeft(i + 1) + 300 - Printer.TextWidth(Format(strTemp(i + 1), "####0"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i + 1), "####0")
    Printer.CurrentX = PLeft(i + 2) + 300 - Printer.TextWidth(Format(strTemp(i + 2), "####0"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i + 2), "####0")
    Printer.CurrentX = PLeft(i + 3) + 300 - Printer.TextWidth(Format(strTemp(i + 3), "####0"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i + 3), "####0")
Next i
For i = 6 To 30 Step 6
    Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(strTemp(i), "##0.00"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i), "##0.00") & "%"
    Printer.CurrentX = PLeft(i + 1) + 400 - Printer.TextWidth(Format(strTemp(i + 1), "##0.00"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i + 1), "##0.00") & "%"
Next i
iPrint = iPrint + 300
End Sub

'Add By Cheng 2003/03/07
Sub PrintDatil2_1()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
'Modify By Cheng 2002/05/06
'若承辦人相同則不印
If m_strPromoter <> strTemp(0) Then
   Printer.Print strTemp(0)
   m_strPromoter = strTemp(0)
Else
   Printer.Print ""
End If

Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(1)
For i = 2 To 13 Step 6
    Printer.Line (PLeft(i) - 50, iPrint - 150)-(PLeft(i) - 50, iPrint + 450)
    Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(strTemp(i), "####0"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i), "####0")
    Printer.CurrentX = PLeft(i + 1) + 300 - Printer.TextWidth(Format(strTemp(i + 1), "####0"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i + 1), "####0")
    Printer.CurrentX = PLeft(i + 2) + 300 - Printer.TextWidth(Format(strTemp(i + 2), "####0"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i + 2), "####0")
    Printer.CurrentX = PLeft(i + 3) + 300 - Printer.TextWidth(Format(strTemp(i + 3), "####0"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i + 3), "####0")
Next i
For i = 6 To 12 Step 6
    Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(strTemp(i), "##0.00"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i), "##0.00") & "%"
    Printer.CurrentX = PLeft(i + 1) + 400 - Printer.TextWidth(Format(strTemp(i + 1), "##0.00"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i + 1), "##0.00") & "%"
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft2()
Erase PLeft
Erase PLeft1
PLeft(0) = 0
PLeft(1) = 700
PLeft(2) = 1600
For i = 3 To 31
    PLeft(i) = 1600 + ((i - 2) * 580)
Next i
PLeft1(1) = PLeft(5)
PLeft1(2) = PLeft(11)
PLeft1(3) = PLeft(17)
PLeft1(4) = PLeft(23)
PLeft1(5) = PLeft(29)
End Sub

Sub PrintTitle1()
GetPleft1
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
For i = 2 To 20 Step 3
    Printer.Line (PLeft(i) - 50, iPrint + 150)-(PLeft(i) - 50, iPrint + 1350)
Next i
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle (1)
    PrintTitle1
    Exit Sub
End If
'add by nickc 2007/07/27 商標處改格式
If txt1(3) = "020" And txt1(4) = "020" Then
    Printer.CurrentX = PLeft1(1) - (Printer.TextWidth(Title_601) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_601 'Modify By Sindy 2015/2/4"異議"
    Printer.CurrentX = PLeft1(2) - (Printer.TextWidth(Title_603) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_603 'Modify By Sindy 2015/2/4"裁定"
    Printer.CurrentX = PLeft1(3) - (Printer.TextWidth(Title_605) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_605 'Modify By Sindy 2015/2/4"撤銷"
    Printer.CurrentX = PLeft1(4) - (Printer.TextWidth("") / 2)
    Printer.CurrentY = iPrint
    Printer.Print "" 'Modify By Sindy 2015/2/4"註冊不當撤銷"
    Printer.CurrentX = PLeft1(5) - (Printer.TextWidth(Title_401) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_401 'Modify By Sindy 2015/2/4"復審"
    Printer.CurrentX = PLeft1(6) - (Printer.TextWidth(Title_408) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_408 'Modify By Sindy 2015/2/4"大陸上訴"
    Printer.CurrentX = PLeft1(7) - (Printer.TextWidth("") / 2)
    Printer.CurrentY = iPrint
    Printer.Print ""
Else
    Printer.CurrentX = PLeft1(1) - (Printer.TextWidth(Title_601) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_601 'Modify By Sindy 2015/2/4"異議"
    Printer.CurrentX = PLeft1(2) - (Printer.TextWidth(Title_603) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_603 'Modify By Sindy 2015/2/4"評定"
    Printer.CurrentX = PLeft1(3) - (Printer.TextWidth(Title_605) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_605 'Modify By Sindy 2015/2/4"廢止"
    Printer.CurrentX = PLeft1(4) - (Printer.TextWidth(Title_401) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_401 'Modify By Sindy 2015/2/4"訴願"
    'Printer.CurrentX = PLeft1(5) - (Printer.TextWidth("再訴願") / 2)
    Printer.CurrentX = PLeft1(5) - (Printer.TextWidth(Title_403) / 2)
    Printer.CurrentY = iPrint
    'Printer.Print "再訴願"
    Printer.Print Title_403 'Modify By Sindy 2015/2/4"行政訴訟"
    'Printer.CurrentX = PLeft1(6) - (Printer.TextWidth("行政訴訟") / 2)
    Printer.CurrentX = PLeft1(6) - (Printer.TextWidth(Title_408) / 2)
    Printer.CurrentY = iPrint
    'Printer.Print "行政訴訟"
    Printer.Print Title_408 'Modify By Sindy 2015/2/4"行政訴訟上訴"
    Printer.CurrentX = PLeft1(7) - (Printer.TextWidth(Title_404) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_404 'Modify By Sindy 2015/2/4"再審之訴"
End If
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle (1)
    PrintTitle1
    Exit Sub
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "業務區別"
For k = 2 To 20 Step 3
    'add by nickc 2007/07/27 商標處改格式
    If txt1(3) = "020" And txt1(4) = "020" And i > 20 Then
    Else
        Printer.CurrentX = PLeft(k)
        Printer.CurrentY = iPrint
        Printer.Print "勝"
        Printer.CurrentX = PLeft(k + 1)
        Printer.CurrentY = iPrint
        Printer.Print "敗"
        Printer.CurrentX = PLeft(k + 2)
        Printer.CurrentY = iPrint
        Printer.Print "勝訴率"
    End If
Next k
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle (1)
    PrintTitle1
    Exit Sub
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle (1)
    PrintTitle1
    Exit Sub
End If
End Sub

'Add By Cheng 2003/03/07
Sub PrintTitle1_1()
GetPleft1
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000 / 4, iPrint + 150)
For i = 2 To 2 Step 3
    Printer.Line (PLeft(i) - 50, iPrint + 150)-(PLeft(i) - 50, iPrint + 1350)
Next i
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle (1)
    PrintTitle1_1
    Exit Sub
End If
Printer.CurrentX = PLeft1(1) - (Printer.TextWidth("申請意見書") / 2)
Printer.CurrentY = iPrint
Printer.Print "申請意見書"
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle (1)
    PrintTitle1_1
    Exit Sub
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "業務區別"
For k = 2 To 2 Step 3
    Printer.CurrentX = PLeft(k)
    Printer.CurrentY = iPrint
    Printer.Print "勝"
    Printer.CurrentX = PLeft(k + 1)
    Printer.CurrentY = iPrint
    Printer.Print "敗"
    Printer.CurrentX = PLeft(k + 2)
    Printer.CurrentY = iPrint
    Printer.Print "勝訴率"
Next k
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle (1)
    PrintTitle1_1
    Exit Sub
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000 / 4, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle (1)
    PrintTitle1_1
    Exit Sub
End If
End Sub

Sub PrintDatil1()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
'Modify By Cheng 2002/05/06
'若承辦人相同則不印
If m_strPromoter <> strTemp(0) Then
   Printer.Print strTemp(0)
   m_strPromoter = strTemp(0)
Else
   Printer.Print ""
End If
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(1)
For i = 2 To 22
    'add by nickc 2007/07/27 商標處改格式
    If txt1(3) = "020" And txt1(4) = "020" And i > 19 Then
    Else
        Select Case i
        Case 4, 7, 10, 13, 16, 19, 22
             Printer.CurrentX = PLeft(i) + 600 - Printer.TextWidth(Format(strTemp(i), "##0.00"))
             Printer.CurrentY = iPrint
             Printer.Print Format(strTemp(i), "##0.00") & "%"
        Case Else
             Printer.CurrentX = PLeft(i) + 400 - Printer.TextWidth(Format(strTemp(i), "####0"))
             Printer.CurrentY = iPrint
             Printer.Print Format(strTemp(i), "####0")
        End Select
    End If
Next i
For i = 2 To 20 Step 3
    'add by nickc 2007/07/27 商標處改格式
    If txt1(3) = "020" And txt1(4) = "020" And i > 20 Then
    Else
        Printer.Line (PLeft(i) - 50, iPrint - 150)-(PLeft(i) - 50, iPrint + 450)
    End If
Next i
iPrint = iPrint + 300
End Sub

'Add By Cheng 2003/03/07
Sub PrintDatil1_1()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
'Modify By Cheng 2002/05/06
'若承辦人相同則不印
If m_strPromoter <> strTemp(0) Then
   Printer.Print strTemp(0)
   m_strPromoter = strTemp(0)
Else
   Printer.Print ""
End If
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(1)
For i = 2 To 4
    Select Case i
    Case 4, 7, 10, 13, 16, 19, 22
         Printer.CurrentX = PLeft(i) + 600 - Printer.TextWidth(Format(strTemp(i), "##0.00"))
         Printer.CurrentY = iPrint
         Printer.Print Format(strTemp(i), "##0.00") & "%"
    Case Else
         Printer.CurrentX = PLeft(i) + 400 - Printer.TextWidth(Format(strTemp(i), "####0"))
         Printer.CurrentY = iPrint
         Printer.Print Format(strTemp(i), "####0")
    End Select
Next i
For i = 2 To 2 Step 3
    Printer.Line (PLeft(i) - 50, iPrint - 150)-(PLeft(i) - 50, iPrint + 450)
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft1()
Erase PLeft
Erase PLeft1
PLeft(0) = 0
PLeft(1) = 1000
PLeft(2) = 2200
For i = 3 To 22
    PLeft(i) = 2200 + ((i - 2) * 776)
Next i
PLeft1(1) = PLeft(3) + 776 / 2
PLeft1(2) = PLeft(6) + 776 / 2
PLeft1(3) = PLeft(9) + 776 / 2
PLeft1(4) = PLeft(12) + 776 / 2
PLeft1(5) = PLeft(15) + 776 / 2
PLeft1(6) = PLeft(18) + 776 / 2
PLeft1(7) = PLeft(21) + 776 / 2
End Sub

Private Sub Form_Load()

MoveFormToCenter Me
txt1(0) = GetSystemKindByNickTnoS

SeekPrintL = Printer.Orientation
PUB_SetPrinter Me.Name, Combo1, , , SeekPrint     'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Printer = Printers(SeekPrint)
Printer.Orientation = SeekPrintL
Set frm020408 = Nothing
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
     strTemp1 = Split(Replace(UCase(GetSystemKindByNickTnoS), ",,", ""), ",")
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
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
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

Sub ShowLine1()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
   If BolEndThisPage = False Then
      Page = Page + 1
      Printer.NewPage
      PrintTitle (1)
      PrintTitle1
   Else
      BolEndThisPage = False
   End If
End If
End Sub

'Add By Cheng 2003/03/07
Sub ShowLine1_1()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000 / 4, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
   If BolEndThisPage = False Then
      Page = Page + 1
      Printer.NewPage
      PrintTitle (1)
      PrintTitle1_1
   Else
      BolEndThisPage = False
   End If
End If
End Sub

Sub ShowLine2()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
   If BolEndThisPage = False Then
      Page = Page + 1
      Printer.NewPage
      PrintTitle (2)
      PrintTitle2
   Else
      BolEndThisPage = False
   End If
End If
End Sub

'Add By Cheng 2003/03/07
Sub ShowLine2_1()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000 / 11 * 5, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
   If BolEndThisPage = False Then
      Page = Page + 1
      Printer.NewPage
      PrintTitle (2)
      PrintTitle2_1
   Else
      BolEndThisPage = False
   End If
End If
End Sub

