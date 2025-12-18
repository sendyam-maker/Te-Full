VERSION 5.00
Begin VB.Form frm020404 
   BorderStyle     =   1  '單線固定
   Caption         =   "商爭案智權人員勝敗統計表"
   ClientHeight    =   2490
   ClientLeft      =   1425
   ClientTop       =   3030
   ClientWidth     =   3900
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   3900
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   528
      Left            =   30
      TabIndex        =   11
      Top             =   1530
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
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1056
      TabIndex        =   0
      Top             =   570
      Width           =   1956
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1056
      MaxLength       =   7
      TabIndex        =   1
      Top             =   885
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2076
      MaxLength       =   7
      TabIndex        =   2
      Top             =   885
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1056
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1215
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2076
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1215
      Width           =   930
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2352
      TabIndex        =   7
      Top             =   12
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3144
      TabIndex        =   6
      Top             =   12
      Width           =   756
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   105
      TabIndex        =   10
      Top             =   570
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "勝敗日期："
      Height          =   180
      Index           =   2
      Left            =   105
      TabIndex        =   9
      Top             =   930
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   3
      Left            =   105
      TabIndex        =   8
      Top             =   1230
      Width           =   915
   End
   Begin VB.Line Line1 
      X1              =   1515
      X2              =   2775
      Y1              =   1035
      Y2              =   1035
   End
   Begin VB.Line Line2 
      X1              =   1590
      X2              =   2340
      Y1              =   1305
      Y2              =   1305
   End
End
Attribute VB_Name = "frm020404"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/9 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, SavDay4 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 50) As String, strTemp3 As String, TestOk As Boolean, StrTemp7(0 To 31) As String, k As Integer
Dim PLeft(0 To 31) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, PLeft1(1 To 7) As Integer, SeekPrint As Integer, SeekPrintL As Integer
Dim BolEndThisPage As Boolean
'Add By Cheng 2002/05/03
Dim m_strSaleZone As String '業務區
'Add By Cheng 2003/03/12
Dim strSQL6_1 As String
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


Private Sub cmdok_Click(Index As Integer)
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
            Txt1_GotFocus 1
            Exit Sub
         End If
         If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
            Me.txt1(2).SetFocus
            Txt1_GotFocus 2
            Exit Sub
         End If

         If Len(txt1(2)) = 0 Then
             s = MsgBox("勝敗日期區間不可空白!!", , "USER 輸入錯誤")
             txt1(1).SetFocus
             Txt1_GotFocus (1)
             Exit Sub
         Else
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
cnnConnection.Execute "DELETE FROM R020404_1 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "DELETE FROM R020404_2 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL6 = ""
'Add By Cheng 2003/03/12
strSQL6_1 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
    'Modify By Cheng 2003/03/11
'   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/10/19
End If
'Add By Cheng 2003/07/11
'若非商申收發文時, 若為TF案則不抓後三碼為"000"的資料
strSQL1 = strSQL1 + " AND CP03 <> Decode(CP01,'TF','0','z') AND CP04 <> Decode(CP01,'TF','00','zz') "
If Len(txt1(1)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(1))) & ""
    'Add By Cheng 2003/03/12
    strSQL6_1 = strSQL6_1 + " AND TM14>=" & Val(ChangeTStringToWString(txt1(1))) & ""
End If
If Len(Trim(txt1(2))) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP05<=" & Val(ChangeTStringToWString(txt1(2))) & " "
    'Add By Cheng 2003/03/12
    strSQL6_1 = strSQL6_1 + " AND TM14<=" & Val(ChangeTStringToWString(txt1(2))) & " "
End If
If Len(txt1(1)) <> 0 Or Len(Trim(txt1(2))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/10/19
End If
If Len(txt1(3)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND TM10>='" & txt1(3) & "' "
    'Add By Cheng 2003/03/12
    strSQL6_1 = strSQL6_1 + " AND TM10>='" & txt1(3) & "' "
End If
If Len(txt1(4)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND TM10<='" & txt1(4) & "' "
    'Add By Cheng 2003/03/12
    strSQL6_1 = strSQL6_1 + " AND TM10<='" & txt1(4) & "' "
End If
If Len(txt1(3)) <> 0 Or Len(Trim(txt1(4))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/10/19
End If
'Modify By Cheng 2002/02/06
'StrSQL6 = StrSQL6 + " AND (CP10='601' OR CP10='603' OR CP10='605' OR CP10='401' OR CP10='402' OR CP10='403' OR CP10='405' OR CP10='602' OR CP10='1601' OR CP10='1602' OR CP10='1603' OR CP10='1604' OR CP10='604' OR CP10='606' OR CP10='1605' OR CP10='1606' OR CP10='202' OR CP10='1202') "
'Modify By Cheng 2003/03/11
'StrSQL6 = StrSQL6 + " AND (CP10='1003' OR CP10='1004') "
strSQL1 = strSQL1 + " AND (CP10='1003' OR CP10='1004') "
'Add By Cheng 2003/03/11
'2011/12/7 modify by sonia
'strSQL2 = strSQL2 + " AND CP10='202' "
strSQL2 = strSQL2 + " AND CP10 in ('202','210') "
'2011/12/7 end
'93.6.14 CANCEL BY SONIA 僅收/發文統計表才控制不算案件數
''Add By Cheng 2004/04/29
''抓計件的資料
'StrSQL6 = StrSQL6 & " And CP26 Is Null "
'strSQL6_1 = strSQL6_1 & " And CP26 Is Null "
''End
CheckOC
'有承辦人申爭之分
'Modify By Cheng 2002/02/06
'取消抓承辦人之ST05為"97"或"17"及CP09>"C"的條件
'strSQL = "SELECT NVL(A0902,A0903),S1.ST02,CP24,S1.ST03,CP10 FROM CASEPROGRESS,STAFF S1,STAFF S2,TRADEMARK,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND (CP26 IS NULL OR CP26='') AND (CP57 IS NULL OR CP57='') AND (S2.ST05='97' OR S2.ST05='17') AND CP09>'C' AND S1.ST03=A0901(+) " & strSQL1 + StrSQL6
'先取得相關總收文號
strSql = " SELECT CP43 " & _
         " FROM CASEPROGRESS,TRADEMARK " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND (CP57 IS NULL OR CP57='') " & strSQL1 + StrSQL6
'Add By Cheng 2003/03/11
'抓案件性質為申請意見書的本所案號
'strSQLA = "SELECT DISTINCT CP09 FROM CASEPROGRESS,( SELECT CP01 C1, CP02 C2,CP03 C3, CP04 C4 " & _
'         " FROM CASEPROGRESS,TRADEMARK " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND (CP57 IS NULL OR CP57='') And (CP22 IS NULL OR CP22 <>'N')  " & strSQL2 + StrSQL6 & " GROUP BY CP01,CP02, CP03, CP04 ) C WHERE CP01=C.C1 AND CP02=C.C2 AND CP03=C.C3 AND CP04=C.C4 "
StrSQLa = " SELECT DISTINCT CP09 " & _
         " FROM CASEPROGRESS,TRADEMARK " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND (CP57 IS NULL OR CP57='') " & strSQL2 + strSQL6_1

'strSQL = "SELECT NVL(A0902,A0903),S1.ST02,CP24,CP12,CP10 " & _
         " FROM CASEPROGRESS,STAFF S1,STAFF S2,ACC090 " & _
         " WHERE CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND (CP26 IS NULL OR CP26='') AND (CP57 IS NULL OR CP57='') AND CP12=A0901(+) AND CP09 IN ( " & strSQL & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(S2.ST03,1,2)='P2' OR CP14 IS NULL ) ", " AND (SUBSTR(S2.ST03,1,2)='F1' OR CP14 IS NULL ) ")
'92.3.6 MODIFY BY SONIA 從外商系統進入, 智權人員統計表不限制承辦人部門別
'strSQL = "SELECT cp12,cp13,CP24,CP12,CP10 " & _
'         " FROM CASEPROGRESS,STAFF S1,STAFF S2,ACC090 " & _
'         " WHERE CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND (CP26 IS NULL OR CP26='') AND (CP57 IS NULL OR CP57='') AND CP12=A0901(+) AND CP09 IN ( " & strSQL & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(S2.ST03,1,2)='P2' OR CP14 IS NULL ) ", " AND (SUBSTR(S2.ST03,1,2)='F1' OR CP14 IS NULL ) ")
'Modify By Cheng 2003/03/06
'不管是否算案件數
'strSQL = "SELECT cp12,cp13,CP24,CP12,CP10 " & _
'         " FROM CASEPROGRESS,STAFF S1,STAFF S2,ACC090 " & _
'         " WHERE CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND (CP26 IS NULL OR CP26='') AND (CP57 IS NULL OR CP57='') AND CP12=A0901(+) AND CP09 IN ( " & strSQL & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(S2.ST03,1,2)='P2' OR CP14 IS NULL ) ", "")
'92.3.6 END
strSql = "SELECT CP12,CP13,CP24,CP12,CP10 " & _
         " FROM CASEPROGRESS,STAFF S1,STAFF S2,ACC090 " & _
         " WHERE CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND (CP57 IS NULL OR CP57='') AND CP12=A0901(+) AND CP09 IN ( " & strSql & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(S2.ST03,1,2)='P2' OR CP14 IS NULL ) ", "")
'沒有承辦人申爭之分
'StrSQL = "SELECT NVL(A0902,A0903),S1.ST02,CP24,S1.ST03,CP10 FROM CASEPROGRESS,STAFF S1,STAFF S2,TRADEMARK,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND (CP26 IS NULL OR CP26='') AND (CP57 IS NULL OR CP57='') AND CP09>'C'  AND S1.ST03=A0901(+) " & StrSQL1 + StrSQL6
'Add By Cheng 2003/03/11
'與申請意見書相同本所案號且案件性質為申請的資料
'strSQL = strSQL & " Union All SELECT CP12,CP13,CP24,CP12,CP10 " & _
'         " FROM CASEPROGRESS,STAFF S1,STAFF S2,ACC090 " & _
'         " WHERE CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND (CP57 IS NULL OR CP57='') AND CP12=A0901(+) AND CP10='101' AND CP24 IS NOT NULL AND CP09 IN ( " & strSQLA & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(S2.ST03,1,2)='P2' OR CP14 IS NULL ) ", "")
'Modify By Cheng 2003/04/10
'不考慮是否出名
'strSQL = strSQL & " union all SELECT CP12,CP13,TM16,CP12,CP10 " & _
'         " FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,ACC090 " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND (CP57 IS NULL OR CP57='') AND CP12=A0901(+) AND (CP22 IS NULL OR CP22 <> 'N') AND TM16 IS NOT NULL AND CP09 IN ( " & strSQLA & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(S2.ST03,1,2)='P2' OR CP14 IS NULL ) ", "")
strSql = strSql & " union all SELECT CP12,CP13,TM16,CP12,CP10 " & _
         " FROM CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,ACC090 " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND (CP57 IS NULL OR CP57='') AND CP12=A0901(+) AND TM16 IS NOT NULL AND CP09 IN ( " & StrSQLa & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(S2.ST03,1,2)='P2' OR CP14 IS NULL ) ", "")

With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/19
        .MoveFirst
        DoEvents
        Do While .EOF = False
            For i = 0 To 4
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
            '****表(1)****
            Case 601, 627 '異議,Add by Sindy 2019/8/15 +部分異議
                 cnnConnection.Execute "INSERT INTO R020404_1 (R070001,R070002,R070003,R070004,R070005,R070024,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & ",0,'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
            Case 603, 629 '評定,Add by Sindy 2019/8/15 +部分評定
                 cnnConnection.Execute "INSERT INTO R020404_1 (R070001,R070002,R070006,R070007,R070008,R070024,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & ",0,'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
            Case 605, 623 '廢止,Add by Sindy 2019/8/15 +部分廢止
                 cnnConnection.Execute "INSERT INTO R020404_1 (R070001,R070002,R070009,R070010,R070011,R070024,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & ",0,'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
            'Modify By Cheng 2002/05/03
'            Case 401 '訴願
            'Modify By Cheng 2003/03/06
            '將參加訴願(406)移至表(2)
'            Case 401, 406 '訴願, 參加訴願
            Case 401 '訴願
                 'edit by nickc 2007/07/27
                 If txt1(3) = "020" And txt1(4) = "020" Then
                 Else
                    cnnConnection.Execute "INSERT INTO R020404_1 (R070001,R070002,R070012,R070013,R070014,R070024,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & ",0,'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                End If
'            'add by nickc 2007/07/27  商標處改大陸格式
'            Case 618
'                If txt1(3) = "020" And txt1(4) = "020" Then
'                    cnnConnection.Execute "INSERT INTO R020404_1 (R070001,R070002,R070012,R070013,R070014,R070024,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & ",0,'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
'                    cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
'                End If
            'Modify By Cheng 2002/05/03
            '取消再訴願, 位置由行政訴訟取代之
'            Case 402 '再訴願
            'Modify By Cheng 2003/03/06
            '將參加訴訟(407)移至表(2)
'            Case 403, 407 '行政訴訟, 參加訴訟
            'Modify By Cheng 2003/07/25
            '通知言詞辯論(1204), 通知準備程序(1203)計入行政訴訟
'            Case 403 '行政訴訟
            Case 403, 1204, 1203 '行政訴訟
                'edit by nickc 2007/07/27  商標處改大陸格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                    If Val(strTemp(4)) = 403 Then
                        cnnConnection.Execute "INSERT INTO R020404_1 (R070001,R070002,R070015,R070016,R070017,R070024,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & ",0,'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                        cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                    End If
                Else
                    cnnConnection.Execute "INSERT INTO R020404_1 (R070001,R070002,R070015,R070016,R070017,R070024,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & ",0,'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                End If
            'Modify By Cheng 2002/05/03
            '行政訴訟欄改為行政訴訟上訴
'            Case 403 '行政訴訟
            Case 408 '行政訴訟上訴
                'edit by nickc 2007/07/27  商標處改大陸格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                Else
                    cnnConnection.Execute "INSERT INTO R020404_1 (R070001,R070002,R070018,R070019,R070020,R070024,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & ",0,'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                End If
            'Modify By Cheng 2002/05/03
'            Case 405
            Case 404 '再審之訴
                'edit by nickc 2007/07/27  商標處改大陸格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                Else
                    cnnConnection.Execute "INSERT INTO R020404_1 (R070001,R070002,R070021,R070022,R070023,R070024,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & ",0,'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                End If
            'Modify By Cheng 2003/03/11
'            'Add By Cheng 2003/03/06
'            '加申請意見書(202)
'            Case 202
'                 cnnConnection.Execute "INSERT INTO R020404_1 (R070001,R070002,R070025,R070026,R070027,R070024,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & ",0,'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
'                 cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
'            Case "101" '申請意見書(<--申請)
            '2011/12/7 MODIFY BY SONIA 加210陳述意見書
            Case "202", "210" '申請意見書,陳述意見書
                'edit by nickc 2007/07/27  商標處改大陸格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                Else
                    cnnConnection.Execute "INSERT INTO R020404_1 (R070001,R070002,R070025,R070026,R070027,R070024,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & ",0,'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                End If
'*******************************
'*******************************
'*******************************
            '****表(2)****
            '異議答辯與被異議(理由)及被異議一組
            Case 602, 628 '異議答辯,Add by Sindy 2019/8/15 +部分異議答辯
                 cnnConnection.Execute "insert into r020404_1 (r070001,r070002,r070024,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071003,R071004,R071005,R071006,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
            Case 1602, 1601 '被異議(理由), 被異議
                 cnnConnection.Execute "insert into r020404_1 (r070001,r070002,r070024,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071003,R071004,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",0,0,'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
            '評定答辯與被評定(理由)及被評定一組
            Case 604, 630 '評定答辯,Add by Sindy 2019/8/15 +部分評定答辯
                 cnnConnection.Execute "insert into r020404_1 (r070001,r070002,r070024,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071009,R071010,R071011,R071012,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
            Case 1603, 1604 '被評定, 被評定(理由)
                 cnnConnection.Execute "insert into r020404_1 (r070001,r070002,r070024,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071009,R071010,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",0,0,'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
            '廢止答辯與被廢止及被廢止(理由)一組
            Case 606, 624 '廢止答辯,Add by Sindy 2019/8/15 +部分廢止答辯
                 cnnConnection.Execute "insert into r020404_1 (r070001,r070002,r070024,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071015,R071016,R071017,R071018,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
            Case 1605, 1606 '被廢止, 被廢止(理由)
                 cnnConnection.Execute "insert into r020404_1 (r070001,r070002,r070024,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071015,R071016,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",0,0,'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
            'Modify By Cheng 2003/03/06
            '申請意見書移至表(1)
'            Case 202 '申請意見書
'                 cnnConnection.Execute "insert into r020404_1 (r070001,r070002,r070024,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
'                 cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071021,R071022,R071023,R071024,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
'            Case 1202 '核駁前先行通知
'                 cnnConnection.Execute "insert into r020404_1 (r070001,r070002,r070024,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
'                 cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071021,R071022,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",0,0,'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
'            'add by nickc 2007/07/27  商標處改大陸格式
'            Case 619
'                If txt1(3) = "020" And txt1(4) = "020" Then
'                    cnnConnection.Execute "insert into r020404_1 (r070001,r070002,r070024,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
'                    cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071021,R071022,R071023,R071024,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
'                End If
            '參加訴願與通知參加訴願一組
            Case 406 '參加訴願
                'edit by nickc 2007/07/27  商標處改大陸格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                    cnnConnection.Execute "insert into r020404_1 (r070001,r070002,r070024,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071034,R071035,R071036,R071037,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                Else
                    cnnConnection.Execute "insert into r020404_1 (r070001,r070002,r070024,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071021,R071022,R071023,R071024,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                End If
            Case 1404 '通知參加訴願
                'edit by nickc 2007/07/27  商標處改大陸格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                    cnnConnection.Execute "insert into r020404_1 (r070001,r070002,r070024,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071034,R071035,R071036,R071037,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                Else
                    cnnConnection.Execute "insert into r020404_1 (r070001,r070002,r070024,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071021,R071022,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",0,0,'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                End If
            '參加訴訟與通知參加訴訟一組
            Case 407 '參加訴訟
                'edit by nickc 2007/07/27  商標處改大陸格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                Else
                    cnnConnection.Execute "insert into r020404_1 (r070001,r070002,r070024,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071034,R071035,R071036,R071037,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                End If
            Case 1405 '通知參加訴訟
                'edit by nickc 2007/07/27  商標處改大陸格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                Else
                    cnnConnection.Execute "insert into r020404_1 (r070001,r070002,r070024,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071034,R071035,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",0,0,'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                End If
            '上訴答辯與通知上訴答辯一組
            Case 410 '上訴答辯
                'edit by nickc 2007/07/27  商標處改大陸格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                Else
                    cnnConnection.Execute "insert into r020404_1 (r070001,r070002,r070024,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071040,R071041,R071042,R071043,R071027,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                End If
            Case 1406 '通知上訴答辯
                'edit by nickc 2007/07/27  商標處改大陸格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                Else
                    cnnConnection.Execute "insert into r020404_1 (r070001,r070002,r070024,id) values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO R020404_2 (R071001,R071002,R071040,R071022,R071041,R071028,R071029,R071030,R071033,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay1) & "," & Val(SavDay2) & ",0,0,'" & ChgSQL(strTemp(3)) & "','" & strUserNum & "') "
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
'*****表(1)格式1**********
BolEndThisPage = False
'Modify By Cheng 2003/03/06
'strSQL = "select NVL(A0902,A0903),ST02,sum(r070003),sum(r070004),sum(r070005),sum(r070006),sum(r070007),sum(r070008),sum(r070009),sum(r070010),sum(r070011),sum(r070012),sum(r070013),sum(r070014),sum(r070015),sum(r070016),sum(r070017),sum(r070018),sum(r070019),sum(r070020),sum(r070021),sum(r070022),sum(r070023),r070024,r070001,r070002 from r020404_1,acc090,staff where R070001=a0901(+) and R070002=st01(+) and id='" & strUserNum & "' group by r070001,r070002,r070024,NVL(A0902,A0903),ST02 order by R070001,R070002 "
strSql = "select NVL(A0902,A0903),ST02,sum(r070003),sum(r070004),sum(r070005),sum(r070006),sum(r070007),sum(r070008),sum(r070009),sum(r070010),sum(r070011),sum(r070012),sum(r070013),sum(r070014),sum(r070015),sum(r070016),sum(r070017),sum(r070018),sum(r070019),sum(r070020),sum(r070021),sum(r070022),sum(r070023),r070024,r070001,r070002,sum(r070025),sum(r070026),sum(r070027) from r020404_1,acc090,staff where R070001=a0901(+) and R070002=st01(+) and id='" & strUserNum & "' group by r070001,r070002,r070024,NVL(A0902,A0903),ST02 order by R070001,R070002 "
CheckOC
'Add By Cheng 2002/05/03
m_strSaleZone = ""

SavDay1 = ""
SavDay2 = ""
SavDay3 = ""
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
        PrintTitle (1)
        PrintTitle1
        Do While .EOF = False
            For i = 0 To 23
                strTemp(i) = CheckStr(.Fields(i))
                If Val(strTemp(i)) = 0 And i <> 0 And i <> 1 And i <> 23 Then
                    strTemp(i) = "0"
                End If
            Next i
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
            'edit by nick 2004/10/06
            'If SavDay1 <> strTemp(0) Then
            If SavDay1 <> strTemp(0) Or StrToStr(SavDay3, 1) <> StrToStr(strTemp(23), 1) Then
               'Add By Cheng 2002/05/03
               m_strSaleZone = ""
               
               ShowLine1
               PrintEnd1 (0)
               If StrToStr(SavDay3, 1) <> StrToStr(strTemp(23), 1) Then
                   ShowLine1
                   PrintEnd1 (1)
                   
               End If
               SavDay3 = strTemp(23)
               ShowLine1
               SavDay1 = strTemp(0)
               SavDay2 = strTemp(1)
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
        'MsgBox "資料庫搜尋不到<<表一>>資料！", vbInformation, "沒有資料"
        GoTo Lett1
    End If
End With
CheckOC
ShowLine1
PrintEnd1 (0)
ShowLine1
PrintEnd1 (1)
ShowLine1
PrintEnd1 (2)
ShowLine1
Page = Page + 1
Printer.NewPage
Lett1:
'add by nickc 2007/07/27 商標處改大陸格式
If txt1(3) = "020" And txt1(4) = "020" Then
Else
    '*****表(1)格式2**********
    BolEndThisPage = False
    'Modify By Cheng 2003/03/06
    'strSQL = "select NVL(A0902,A0903),ST02,sum(r070003),sum(r070004),sum(r070005),sum(r070006),sum(r070007),sum(r070008),sum(r070009),sum(r070010),sum(r070011),sum(r070012),sum(r070013),sum(r070014),sum(r070015),sum(r070016),sum(r070017),sum(r070018),sum(r070019),sum(r070020),sum(r070021),sum(r070022),sum(r070023),r070024,r070001,r070002 from r020404_1,acc090,staff where R070001=a0901(+) and R070002=st01(+) and id='" & strUserNum & "' group by r070001,r070002,r070024,NVL(A0902,A0903),ST02 order by R070001,R070002 "
    strSql = "select NVL(A0902,A0903),ST02,sum(r070003),sum(r070004),sum(r070005),sum(r070006),sum(r070007),sum(r070008),sum(r070009),sum(r070010),sum(r070011),sum(r070012),sum(r070013),sum(r070014),sum(r070015),sum(r070016),sum(r070017),sum(r070018),sum(r070019),sum(r070020),sum(r070021),sum(r070022),sum(r070023),r070024,r070001,r070002,sum(r070025),sum(r070026),sum(r070027) from r020404_1,acc090,staff where R070001=a0901(+) and R070002=st01(+) and id='" & strUserNum & "' group by r070001,r070002,r070024,NVL(A0902,A0903),ST02 order by R070001,R070002 "
    CheckOC
    'Add By Cheng 2002/05/03
    m_strSaleZone = ""
    SavDay1 = ""
    SavDay2 = ""
    SavDay3 = ""
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
            PrintTitle (1)
            PrintTitle1_1
            Do While .EOF = False
                For i = 0 To 23
                    strTemp(i) = CheckStr(.Fields(i))
                    If Val(strTemp(i)) = 0 And i <> 0 And i <> 1 And i <> 23 Then
                        strTemp(i) = "0"
                    End If
                Next i
                strTemp(2) = Val("0" & .Fields(26))
                strTemp(3) = Val("0" & .Fields(27))
                strTemp(4) = Val("0" & .Fields(28))
                If Val(strTemp(2)) + Val(strTemp(3)) = 0 Then
                    strTemp(4) = "0"
                Else
                    strTemp(4) = Trim(str(Val(strTemp(2)) / (Val(strTemp(2)) + Val(strTemp(3))) * 100))
                End If
                'edit by nick 2004/10/06
                'If SavDay1 <> strTemp(0) Then
                If SavDay1 <> strTemp(0) Or StrToStr(SavDay3, 1) <> StrToStr(strTemp(23), 1) Then
                   'Add By Cheng 2002/05/03
                   m_strSaleZone = ""
                   ShowLine1_1
                   PrintEnd1_1 (0)
                   If StrToStr(SavDay3, 1) <> StrToStr(strTemp(23), 1) Then
                       ShowLine1_1
                       PrintEnd1_1 (1)
                   End If
                   SavDay3 = strTemp(23)
                   ShowLine1_1
                   SavDay1 = strTemp(0)
                   SavDay2 = strTemp(1)
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
            'MsgBox "資料庫搜尋不到<<表一>>資料！", vbInformation, "沒有資料"
            GoTo Lett2
        End If
    End With
    CheckOC
    ShowLine1_1
    PrintEnd1_1 (0)
    ShowLine1_1
    PrintEnd1_1 (1)
    ShowLine1_1
    PrintEnd1_1 (2)
    ShowLine1_1
    Page = Page + 1
    Printer.EndDoc
End If
Lett2:
'*****表(2)格式1*********
'Modify By Cheng 2003/03/06
'strSQL = "select NVL(A0902,A0903),ST02,sum(r071003),sum(r071004),sum(r071005),sum(r071006),sum(r071007),sum(r071008),sum(r071009),sum(r071010),sum(r071011),sum(r071012),sum(r071013),sum(r071014),sum(r071015),sum(r071016),sum(r071017),sum(r071018),sum(r071019),sum(r071020),sum(r071021),sum(r071022),sum(r071023),sum(r071024),sum(r071025),sum(r071026),sum(r071027),sum(r071028),sum(r071029),sum(r071030),sum(r071031),sum(r071032),R071033,R071001,r071002 from r020404_2,acc090,staff where R071001=a0901(+) and R071002=st01(+) and id='" & strUserNum & "' group by r071001,r071002,r071033,NVL(A0902,A0903),ST02 order by R071001,R071002 "
strSql = "select NVL(A0902,A0903),ST02,sum(r071003),sum(r071004),sum(r071005),sum(r071006),sum(r071007),sum(r071008),sum(r071009),sum(r071010),sum(r071011),sum(r071012),sum(r071013),sum(r071014),sum(r071015),sum(r071016),sum(r071017),sum(r071018),sum(r071019),sum(r071020),sum(r071021),sum(r071022),sum(r071023),sum(r071024),sum(r071025),sum(r071026),sum(r071027),sum(r071028),sum(r071029),sum(r071030),sum(r071031),sum(r071032),R071033,R071001,r071002,sum(r071034),sum(r071035),sum(r071036),sum(r071037),sum(r071038),sum(r071039),sum(r071040),sum(r071041),sum(r071042),sum(r071043),sum(r071044),sum(r071045) " & _
            " from r020404_2,acc090,staff where R071001=a0901(+) and R071002=st01(+) and id='" & strUserNum & "' group by r071001,r071002,r071033,NVL(A0902,A0903),ST02 order by R071001,R071002 "
'strSQL = "select NVL(A0902,A0903),ST02,sum(r070003),sum(r070004),sum(r070005),sum(r070006),sum(r070007),sum(r070008),sum(r070009),sum(r070010),sum(r070011),sum(r070012),sum(r070013),sum(r070014),sum(r070015),sum(r070016),sum(r070017),sum(r070018),sum(r070019),sum(r070020),sum(r070021),sum(r070022),sum(r070023),r070024,r070001,r070002 from r020404_1,acc090,staff where R070001=a0901(+) and R070002=st01(+) and id='" & strUserNum & "' group by r070001,r070002,r070024,NVL(A0902,A0903),ST02 order by R070001,R070002 "
CheckOC
'Add By Cheng 2002/05/03
m_strSaleZone = ""
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
        PrintTitle (2)
        PrintTitle2
        Do While .EOF = False
            For i = 0 To 32 + 14
                strTemp(i) = CheckStr(.Fields(i))
                If Val(strTemp(i)) = 0 And i <> 0 And i <> 1 And i <> 32 Then
                    strTemp(i) = "0"
                End If
            Next i
            strTemp(26) = Val("0" & .Fields(35))
            strTemp(27) = Val("0" & .Fields(36))
            strTemp(28) = Val("0" & .Fields(37))
            strTemp(29) = Val("0" & .Fields(38))
            strTemp(30) = Val("0" & .Fields(39))
            strTemp(31) = Val("0" & .Fields(40))
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
            'edit by nick 2004/10/06
            'If SavDay1 <> strTemp(0) Then
            If SavDay1 <> strTemp(0) Or StrToStr(SavDay3, 1) <> StrToStr(strTemp(32), 1) Then
                m_strSaleZone = ""
                ShowLine2
                PrintEnd2 (0)
                If StrToStr(SavDay3, 1) <> StrToStr(strTemp(32), 1) Then
                    ShowLine2
                    PrintEnd2 (1)
               End If
               SavDay3 = strTemp(32)
                ShowLine2
            SavDay1 = strTemp(0)
            SavDay2 = strTemp(1)
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
        'MsgBox "資料庫搜尋不到<<表二>>資料！", vbInformation, "沒有資料"
        'Exit Sub
        GoTo Lett3
    End If
End With
ShowLine2
PrintEnd2 (0)
ShowLine2
PrintEnd2 (1)
ShowLine2
PrintEnd2 (2)
ShowLine2
Page = Page + 1
Printer.NewPage
Lett3:
'add by nickc 2007/07/27 商標處改大陸格式
If txt1(3) = "020" And txt1(4) = "020" Then
Else
    '*****表(2)格式2*********
    'Modify By Cheng 2003/03/06
    'strSQL = "select NVL(A0902,A0903),ST02,sum(r071003),sum(r071004),sum(r071005),sum(r071006),sum(r071007),sum(r071008),sum(r071009),sum(r071010),sum(r071011),sum(r071012),sum(r071013),sum(r071014),sum(r071015),sum(r071016),sum(r071017),sum(r071018),sum(r071019),sum(r071020),sum(r071021),sum(r071022),sum(r071023),sum(r071024),sum(r071025),sum(r071026),sum(r071027),sum(r071028),sum(r071029),sum(r071030),sum(r071031),sum(r071032),R071033,R071001,r071002 from r020404_2,acc090,staff where R071001=a0901(+) and R071002=st01(+) and id='" & strUserNum & "' group by r071001,r071002,r071033,NVL(A0902,A0903),ST02 order by R071001,R071002 "
    strSql = "select NVL(A0902,A0903),ST02,sum(r071003),sum(r071004),sum(r071005),sum(r071006),sum(r071007),sum(r071008),sum(r071009),sum(r071010),sum(r071011),sum(r071012),sum(r071013),sum(r071014),sum(r071015),sum(r071016),sum(r071017),sum(r071018),sum(r071019),sum(r071020),sum(r071021),sum(r071022),sum(r071023),sum(r071024),sum(r071025),sum(r071026),sum(r071027),sum(r071028),sum(r071029),sum(r071030),sum(r071031),sum(r071032),R071033,R071001,r071002,sum(r071034),sum(r071035),sum(r071036),sum(r071037),sum(r071038),sum(r071039),sum(r071040),sum(r071041),sum(r071042),sum(r071043),sum(r071044),sum(r071045) " & _
                " from r020404_2,acc090,staff where R071001=a0901(+) and R071002=st01(+) and id='" & strUserNum & "' group by r071001,r071002,r071033,NVL(A0902,A0903),ST02 order by R071001,R071002 "
    'strSQL = "select NVL(A0902,A0903),ST02,sum(r070003),sum(r070004),sum(r070005),sum(r070006),sum(r070007),sum(r070008),sum(r070009),sum(r070010),sum(r070011),sum(r070012),sum(r070013),sum(r070014),sum(r070015),sum(r070016),sum(r070017),sum(r070018),sum(r070019),sum(r070020),sum(r070021),sum(r070022),sum(r070023),r070024,r070001,r070002 from r020404_1,acc090,staff where R070001=a0901(+) and R070002=st01(+) and id='" & strUserNum & "' group by r070001,r070002,r070024,NVL(A0902,A0903),ST02 order by R070001,R070002 "
    CheckOC
    'Add By Cheng 2002/05/03
    m_strSaleZone = ""
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
            PrintTitle (2)
            PrintTitle2_1
            Do While .EOF = False
                For i = 0 To 32 + 14
                    strTemp(i) = CheckStr(.Fields(i))
                    If Val(strTemp(i)) = 0 And i <> 0 And i <> 1 And i <> 32 Then
                        strTemp(i) = "0"
                    End If
                Next i
                strTemp(2) = Val("0" & .Fields(41))
                strTemp(3) = Val("0" & .Fields(42))
                strTemp(4) = Val("0" & .Fields(43))
                strTemp(5) = Val("0" & .Fields(44))
                strTemp(6) = Val("0" & .Fields(45))
                strTemp(7) = Val("0" & .Fields(46))
                
                strTemp(8) = Val("0" & .Fields(26))
                strTemp(9) = Val("0" & .Fields(27))
                strTemp(10) = Val("0" & .Fields(28))
                strTemp(11) = Val("0" & .Fields(29))
                strTemp(12) = Val("0" & .Fields(30))
                strTemp(13) = Val("0" & .Fields(31))
                
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
                'edit by nick 2004/10/06
                'If SavDay1 <> strTemp(0) Then
                If SavDay1 <> strTemp(0) Or StrToStr(SavDay3, 1) <> StrToStr(strTemp(32), 1) Then
                    m_strSaleZone = ""
                    ShowLine2_1
                    PrintEnd2_1 (0)
                    If StrToStr(SavDay3, 1) <> StrToStr(strTemp(32), 1) Then
                        ShowLine2_1
                        PrintEnd2_1 (1)
                   End If
                   SavDay3 = strTemp(32)
                    ShowLine2_1
                SavDay1 = strTemp(0)
                SavDay2 = strTemp(1)
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
            'MsgBox "資料庫搜尋不到<<表二>>資料！", vbInformation, "沒有資料"
            'Exit Sub
            GoTo Lett4
        End If
    End With
    ShowLine2_1
    PrintEnd2_1 (0)
    ShowLine2_1
    PrintEnd2_1 (1)
    ShowLine2_1
    PrintEnd2_1 (2)
    ShowLine2_1
    Page = Page + 1
Lett4:
End If
If bolNotData = False Then Printer.EndDoc
End Sub

Sub PrintEnd1(Strindex As Integer)
Select Case Strindex
Case 0
      'Modify By Cheng 2002/05/03
'     strSQL = "select '各區小計','',SUM(R070003),sum(r070004),sum(r070005),sum(r070006),sum(r070007),sum(r070008),sum(r070009),sum(r070010),sum(r070011),sum(r070012),sum(r070013),sum(r070014),sum(r070015),sum(r070016),sum(r070017),sum(r070018),sum(r070019),sum(r070020),sum(r070021),sum(r070022),sum(r070023) from r020404_1 where id='" & strUserNum & "' AND R070001='" & SavDay1 & "' AND R070002='" & SavDay2 & "' AND R070024='" & SavDay3 & "'  group by r070001,r070002,r070024 "
     strSql = "select '各區小計','',SUM(R070003),sum(r070004),sum(r070005),sum(r070006),sum(r070007),sum(r070008),sum(r070009),sum(r070010),sum(r070011),sum(r070012),sum(r070013),sum(r070014),sum(r070015),sum(r070016),sum(r070017),sum(r070018),sum(r070019),sum(r070020),sum(r070021),sum(r070022),sum(r070023) from r020404_1 where id='" & strUserNum & "' AND R070024='" & SavDay3 & "'  group by r070001,r070024 "
Case 1
     strSql = "select '各所小計','',SUM(R070003),sum(r070004),sum(r070005),sum(r070006),sum(r070007),sum(r070008),sum(r070009),sum(r070010),sum(r070011),sum(r070012),sum(r070013),sum(r070014),sum(r070015),sum(r070016),sum(r070017),sum(r070018),sum(r070019),sum(r070020),sum(r070021),sum(r070022),sum(r070023) from r020404_1 where id='" & strUserNum & "' AND SUBSTR(R070024,1,2)='" & StrToStr(SavDay3, 1) & "'  "
Case 2
     strSql = "select '全所總計','',SUM(R070003),sum(r070004),sum(r070005),sum(r070006),sum(r070007),sum(r070008),sum(r070009),sum(r070010),sum(r070011),sum(r070012),sum(r070013),sum(r070014),sum(r070015),sum(r070016),sum(r070017),sum(r070018),sum(r070019),sum(r070020),sum(r070021),sum(r070022),sum(r070023) from r020404_1 where id='" & strUserNum & "' "
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
                'add by nickc 2007/07/27  商標處改大陸格式
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

Sub PrintEnd2(Strindex As Integer)
Select Case Strindex
Case 0
     'strSQL = "select '各區小計','',sum(r071003),sum(r071004),sum(r071005),sum(r071006),sum(r071007),sum(r071008),sum(r071009),sum(r071010),sum(r071011),sum(r071012),sum(r071013),sum(r071014),sum(r071015),sum(r071016),sum(r071017),sum(r071018),sum(r071019),sum(r071020),sum(r071021),sum(r071022),sum(r071023),sum(r071024),sum(r071025),sum(r071026),sum(r071027),sum(r071028),sum(r071029),sum(r071030),sum(r071031),sum(r071032) from r020404_2 WHERE ID='" & strUserNum & "' AND R071001='" & SavDay1 & "' AND R071002='" & SavDay2 & "' AND R071033='" & SavDay3 & "' group by r071001,r071002,r071033"
     'nick 91/05/21
    'Modify By Cheng 2003/03/06
'     strSQL = "select '各區小計','',sum(r071003),sum(r071004),sum(r071005),sum(r071006),sum(r071007),sum(r071008),sum(r071009),sum(r071010),sum(r071011),sum(r071012),sum(r071013),sum(r071014),sum(r071015),sum(r071016),sum(r071017),sum(r071018),sum(r071019),sum(r071020),sum(r071021),sum(r071022),sum(r071023),sum(r071024),sum(r071025),sum(r071026),sum(r071027),sum(r071028),sum(r071029),sum(r071030),sum(r071031),sum(r071032) from r020404_2 WHERE ID='" & strUserNum & "' AND R071033='" & SavDay3 & "' group by r071033"
     strSql = "select '各區小計','',sum(r071003),sum(r071004),sum(r071005),sum(r071006),sum(r071007),sum(r071008),sum(r071009),sum(r071010),sum(r071011),sum(r071012),sum(r071013),sum(r071014),sum(r071015),sum(r071016),sum(r071017),sum(r071018),sum(r071019),sum(r071020),sum(r071021),sum(r071022),sum(r071023),sum(r071024),sum(r071025),sum(r071026),sum(r071034),sum(r071035),sum(r071036),sum(r071037),sum(r071038),sum(r071039) from r020404_2 WHERE ID='" & strUserNum & "' AND R071033='" & SavDay3 & "' group by r071033"
Case 1
        'Modify By Cheng 2003/03/06
'     strSQL = "select '各所小計','',sum(r071003),sum(r071004),sum(r071005),sum(r071006),sum(r071007),sum(r071008),sum(r071009),sum(r071010),sum(r071011),sum(r071012),sum(r071013),sum(r071014),sum(r071015),sum(r071016),sum(r071017),sum(r071018),sum(r071019),sum(r071020),sum(r071021),sum(r071022),sum(r071023),sum(r071024),sum(r071025),sum(r071026),sum(r071027),sum(r071028),sum(r071029),sum(r071030),sum(r071031),sum(r071032) from r020404_2 WHERE ID='" & strUserNum & "' AND SUBSTR(R071033,1,2)='" & StrToStr(SavDay3, 1) & "' "
     strSql = "select '各所小計','',sum(r071003),sum(r071004),sum(r071005),sum(r071006),sum(r071007),sum(r071008),sum(r071009),sum(r071010),sum(r071011),sum(r071012),sum(r071013),sum(r071014),sum(r071015),sum(r071016),sum(r071017),sum(r071018),sum(r071019),sum(r071020),sum(r071021),sum(r071022),sum(r071023),sum(r071024),sum(r071025),sum(r071026),sum(r071034),sum(r071035),sum(r071036),sum(r071037),sum(r071038),sum(r071039) from r020404_2 WHERE ID='" & strUserNum & "' AND SUBSTR(R071033,1,2)='" & StrToStr(SavDay3, 1) & "' "
Case 2
        'Modify By Cheng 2003/03/06
'     strSQL = "select '全所總計','',sum(r071003),sum(r071004),sum(r071005),sum(r071006),sum(r071007),sum(r071008),sum(r071009),sum(r071010),sum(r071011),sum(r071012),sum(r071013),sum(r071014),sum(r071015),sum(r071016),sum(r071017),sum(r071018),sum(r071019),sum(r071020),sum(r071021),sum(r071022),sum(r071023),sum(r071024),sum(r071025),sum(r071026),sum(r071027),sum(r071028),sum(r071029),sum(r071030),sum(r071031),sum(r071032) from r020404_2 WHERE ID='" & strUserNum & "' "
     strSql = "select '全所總計','',sum(r071003),sum(r071004),sum(r071005),sum(r071006),sum(r071007),sum(r071008),sum(r071009),sum(r071010),sum(r071011),sum(r071012),sum(r071013),sum(r071014),sum(r071015),sum(r071016),sum(r071017),sum(r071018),sum(r071019),sum(r071020),sum(r071021),sum(r071022),sum(r071023),sum(r071024),sum(r071025),sum(r071026),sum(r071034),sum(r071035),sum(r071036),sum(r071037),sum(r071038),sum(r071039) from r020404_2 WHERE ID='" & strUserNum & "' "
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

Sub PrintTitle(Strindex As String)
iPrint = 0
'Printer.Orientation = 1 'Removed by Morgan 2015/6/3
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "商爭案智權人員勝敗統計表(" & Trim(str(Strindex)) & ") "
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

'Add By Cheng 2002/02/06
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
'add by nickc 2007/07/27  商標處改大陸格式
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
    'Modify By Cheng 2003/03/06
    'Printer.CurrentX = PLeft1(4) - (Printer.TextWidth("申請意見書") / 2)
    Printer.CurrentX = PLeft1(4) - (Printer.TextWidth(Title_406) / 2)
    Printer.CurrentY = iPrint
    'Printer.Print "申請意見書"
    Printer.Print Title_406 'Modify By Sindy 2015/2/4"參加訴願"
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
Printer.Print "業務區別"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
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

Sub PrintDatil2()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
'Modify By Cheng 2002/05/03
If m_strSaleZone <> strTemp(0) Then
   Printer.Print strTemp(0)
   m_strSaleZone = strTemp(0)
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

Sub GetPleft2()
Erase PLeft
Erase PLeft1
PLeft(0) = 0
PLeft(1) = 900
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
'add by nickc 2007/07/27  商標處改大陸格式
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
    Printer.Print "" 'Modify By Sindy 2015/2/4 "註冊不當撤銷"
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
Printer.Print "業務區別"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
For k = 2 To 20 Step 3
    'add by nickc 2007/07/27  商標處改大陸格式
    If txt1(3) = "020" And txt1(4) = "020" And k > 17 Then
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

Sub PrintDatil1()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
'Modify By Cheng 2002/05/03
'若業務區相同, 則不印出業務區名稱
If m_strSaleZone <> strTemp(0) Then
   Printer.Print strTemp(0)
   m_strSaleZone = strTemp(0)
Else
   Printer.Print ""
End If
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(1)
For i = 2 To 22
    'add by nickc 2007/07/27  商標處改大陸格式
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
    'add by nickc 2007/07/27  商標處改大陸格式
    If txt1(3) = "020" And txt1(4) = "020" And i > 20 Then
    Else
        Printer.Line (PLeft(i) - 50, iPrint - 150)-(PLeft(i) - 50, iPrint + 450)
    End If
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft1()
Erase PLeft
Erase PLeft1
PLeft(0) = 0
PLeft(1) = 1200
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
Set frm020404 = Nothing
End Sub

Private Sub Txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub Txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdok(0).SetFocus
End If
End Sub

Private Sub Txt1_KeyPress(Index As Integer, KeyAscii As Integer)
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
Case 2, 1
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Me.txt1(Index).SetFocus
      Txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 2 Then
     If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         Txt1_GotFocus (Index - 1)
         Exit Sub
      End If
    End If
Case 4
     If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         Txt1_GotFocus (Index - 1)
         Exit Sub
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

'Add By Cheng 2003/03/06
Sub PrintDatil1_1()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
'Modify By Cheng 2002/05/03
'若業務區相同, 則不印出業務區名稱
If m_strSaleZone <> strTemp(0) Then
   Printer.Print strTemp(0)
   m_strSaleZone = strTemp(0)
Else
   Printer.Print ""
End If
'智權人員
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

'Add By Cheng 2003/03/06
Sub PrintEnd1_1(Strindex As Integer)
Select Case Strindex
Case 0
     strSql = "select '各區小計','',SUM(R070025),sum(r070026),sum(r070027) from r020404_1 where id='" & strUserNum & "' AND R070024='" & SavDay3 & "'  group by r070001,r070024 "
Case 1
     strSql = "select '各所小計','',SUM(R070025),sum(r070026),sum(r070027) from r020404_1 where id='" & strUserNum & "' AND SUBSTR(R070024,1,2)='" & StrToStr(SavDay3, 1) & "'  "
Case 2
     strSql = "select '全所總計','',SUM(R070025),sum(r070026),sum(r070027) from r020404_1 where id='" & strUserNum & "' "
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

'Add By Cheng 2003/03/06
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

'Add By Cheng 2003/03/06
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
    Page = Page
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
    PrintTitle1
    Exit Sub
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "業務區別"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
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
    Page = Page
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
    Page = Page
    Printer.NewPage
    PrintTitle (1)
    PrintTitle1_1
    Exit Sub
End If
End Sub

'Add By Cheng 2003/03/06
'表(2)格式2
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
Printer.Print "業務區別"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
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

'Add By Cheng 2003/03/06
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

'Add By Cheng 2003/03/06
Sub PrintEnd2_1(Strindex As Integer)
Select Case Strindex
Case 0
     'strSQL = "select '各區小計','',sum(r071003),sum(r071004),sum(r071005),sum(r071006),sum(r071007),sum(r071008),sum(r071009),sum(r071010),sum(r071011),sum(r071012),sum(r071013),sum(r071014),sum(r071015),sum(r071016),sum(r071017),sum(r071018),sum(r071019),sum(r071020),sum(r071021),sum(r071022),sum(r071023),sum(r071024),sum(r071025),sum(r071026),sum(r071027),sum(r071028),sum(r071029),sum(r071030),sum(r071031),sum(r071032) from r020404_2 WHERE ID='" & strUserNum & "' AND R071001='" & SavDay1 & "' AND R071002='" & SavDay2 & "' AND R071033='" & SavDay3 & "' group by r071001,r071002,r071033"
     'nick 91/05/21
    'Modify By Cheng 2003/03/06
'     strSQL = "select '各區小計','',sum(r071003),sum(r071004),sum(r071005),sum(r071006),sum(r071007),sum(r071008),sum(r071009),sum(r071010),sum(r071011),sum(r071012),sum(r071013),sum(r071014),sum(r071015),sum(r071016),sum(r071017),sum(r071018),sum(r071019),sum(r071020),sum(r071021),sum(r071022),sum(r071023),sum(r071024),sum(r071025),sum(r071026),sum(r071027),sum(r071028),sum(r071029),sum(r071030),sum(r071031),sum(r071032) from r020404_2 WHERE ID='" & strUserNum & "' AND R071033='" & SavDay3 & "' group by r071033"
     strSql = "select '各區小計','',sum(r071040),sum(r071041),sum(r071042),sum(r071043),sum(r071044),sum(r071045),sum(r071027),sum(r071028),sum(r071029),sum(r071030),sum(r071031),sum(r071032) from r020404_2 WHERE ID='" & strUserNum & "' AND R071033='" & SavDay3 & "' group by r071033"
Case 1
        'Modify By Cheng 2003/03/06
'     strSQL = "select '各所小計','',sum(r071003),sum(r071004),sum(r071005),sum(r071006),sum(r071007),sum(r071008),sum(r071009),sum(r071010),sum(r071011),sum(r071012),sum(r071013),sum(r071014),sum(r071015),sum(r071016),sum(r071017),sum(r071018),sum(r071019),sum(r071020),sum(r071021),sum(r071022),sum(r071023),sum(r071024),sum(r071025),sum(r071026),sum(r071027),sum(r071028),sum(r071029),sum(r071030),sum(r071031),sum(r071032) from r020404_2 WHERE ID='" & strUserNum & "' AND SUBSTR(R071033,1,2)='" & StrToStr(SavDay3, 1) & "' "
     strSql = "select '各所小計','',sum(r071040),sum(r071041),sum(r071042),sum(r071043),sum(r071044),sum(r071045),sum(r071027),sum(r071028),sum(r071029),sum(r071030),sum(r071031),sum(r071032) from r020404_2 WHERE ID='" & strUserNum & "' AND SUBSTR(R071033,1,2)='" & StrToStr(SavDay3, 1) & "' "
Case 2
        'Modify By Cheng 2003/03/06
'     strSQL = "select '全所總計','',sum(r071003),sum(r071004),sum(r071005),sum(r071006),sum(r071007),sum(r071008),sum(r071009),sum(r071010),sum(r071011),sum(r071012),sum(r071013),sum(r071014),sum(r071015),sum(r071016),sum(r071017),sum(r071018),sum(r071019),sum(r071020),sum(r071021),sum(r071022),sum(r071023),sum(r071024),sum(r071025),sum(r071026),sum(r071027),sum(r071028),sum(r071029),sum(r071030),sum(r071031),sum(r071032) from r020404_2 WHERE ID='" & strUserNum & "' "
     strSql = "select '全所總計','',sum(r071040),sum(r071041),sum(r071042),sum(r071043),sum(r071044),sum(r071045),sum(r071027),sum(r071028),sum(r071029),sum(r071030),sum(r071031),sum(r071032) from r020404_2 WHERE ID='" & strUserNum & "' "
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

Sub PrintDatil2_1()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
'Modify By Cheng 2002/05/03
If m_strSaleZone <> strTemp(0) Then
   Printer.Print strTemp(0)
   m_strSaleZone = strTemp(0)
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
For i = 6 To 13 Step 6
    Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(strTemp(i), "##0.00"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i), "##0.00") & "%"
    Printer.CurrentX = PLeft(i + 1) + 400 - Printer.TextWidth(Format(strTemp(i + 1), "##0.00"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i + 1), "##0.00") & "%"
Next i
iPrint = iPrint + 300
End Sub

