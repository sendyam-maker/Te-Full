VERSION 5.00
Begin VB.Form frm020412 
   BorderStyle     =   1  '單線固定
   Caption         =   "各區收/發文件數明細表"
   ClientHeight    =   2010
   ClientLeft      =   3600
   ClientTop       =   3360
   ClientWidth     =   4320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4320
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   540
      Left            =   240
      TabIndex        =   5
      Top             =   1380
      Width           =   3825
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   6
         Top             =   180
         Width           =   2880
      End
      Begin VB.Label Label4 
         Caption         =   "印表機"
         Height          =   180
         Left            =   105
         TabIndex        =   7
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3375
      TabIndex        =   3
      Top             =   135
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2580
      TabIndex        =   2
      Top             =   135
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   2700
      MaxLength       =   7
      TabIndex        =   1
      Top             =   840
      Width           =   1065
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1500
      MaxLength       =   7
      TabIndex        =   0
      Top             =   840
      Width           =   1065
   End
   Begin VB.Line Line1 
      X1              =   2220
      X2              =   3240
      Y1              =   975
      Y2              =   975
   End
   Begin VB.Label Label1 
      Caption         =   "收/發文期間："
      Height          =   180
      Left            =   345
      TabIndex        =   4
      Top             =   870
      Width           =   1170
   End
End
Attribute VB_Name = "frm020412"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/5/4 改成Form2.0 (Unicode文字以圖片方式列印)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/10 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, SavDay4 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 17) As String, strTemp3 As String, TestOk As Boolean, StrTemp7(0 To 10) As String
Dim PLeft(0 To 10) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, SeekPrint As Integer, SeekPrintL As Integer, StrSQL7 As String, strSQL8 As String
'Add By Cheng 2002/02/25
Dim strTemp3_1 As String
Dim StrTemp4 As String
Dim m_strSaleZone1 As String
Dim m_strSaleZone2 As String
Const m_intLastLn  As Integer = 63 '65
Dim strPrinter As String 'Add By Sindy 2015/7/3


Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     If Len(txt1(0)) = 0 Then
         s = MsgBox("收/發文區間不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         txt1_GotFocus (0)
         Exit Sub
      End If
     If Len(txt1(1)) = 0 Then
         s = MsgBox("收/發文區間不可空白!!", , "USER 輸入錯誤")
         txt1(1).SetFocus
         txt1_GotFocus (1)
         Exit Sub
      End If
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
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/19 清除查詢印表記錄檔欄位
      Printer.EndDoc 'Add By Sindy 2011/11/1
      PUB_RestorePrinter Combo1.Text 'Add By Sindy 2015/7/3
      Process
      PUB_RestorePrinter strPrinter 'Add By Sindy 2015/7/3
      Me.Enabled = True
      Screen.MousePointer = vbDefault
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub Process()
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R020412 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
'Modify By Cheng 2002/02/22
'只抓"T"類資料
'StrSQL6 = GetSystemKindByNickT
StrSQL6 = "T"
'For i = 0 To UBound(strTemp1)
   strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(StrSQL6, 2) & ") "
   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(StrSQL6, 5) & ") "
'Next i
StrSQL6 = ""
StrSQL7 = ""
strSQL8 = ""
If Len(txt1(0)) <> 0 Then
   StrSQL7 = StrSQL7 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(0))) & " "
   strSQL8 = strSQL8 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(0))) & " "
End If
If Len(Trim(txt1(1))) <> 0 Then
   StrSQL7 = StrSQL7 + " AND CP05<=" & Val(ChangeTStringToWString(txt1(1)))
   strSQL8 = strSQL8 + " AND CP27<=" & Val(ChangeTStringToWString(txt1(1)))
End If
If Len(txt1(0)) <> 0 Or Len(Trim(txt1(1))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) & "-" & txt1(1) 'Add By Sindy 2010/10/19
End If
'StrSQL = "SELECT NVL(A0902,A0903),ST02,1,0,0,0,0,0,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='101' AND TM10='000' " & strSQL1 & StrSQL7
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,1,0,0,0,0,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='101' AND TM10>'000' " & strSQL1 & StrSQL7
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,1,0,0,0,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='806' " & strSQL1 & StrSQL7
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,0,1,0,0,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND (CP10='801' OR CP10='802' OR CP10='805') " & strSQL1 & StrSQL7
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,0,0,1,0,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='101' AND TM10='000' " & strSQL1 & StrSQL8
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,0,0,0,1,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='101' AND TM10>'000' " & strSQL1 & StrSQL8
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,0,0,0,0,1,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='806' " & strSQL1 & StrSQL8
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,0,0,0,0,0,1,ST03,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND (CP10='801' OR CP10='802' OR CP10='805') " & strSQL1 & StrSQL8
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,1,0,0,0,0,0,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='101' AND SP09='000' " & strSQL2 & StrSQL7
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,1,0,0,0,0,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='101' AND SP09>'000' " & strSQL2 & StrSQL7
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,1,0,0,0,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='806' " & strSQL2 & StrSQL7
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,0,1,0,0,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND (CP10='801' OR CP10='802' OR CP10='805') " & strSQL2 & StrSQL7
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,0,0,1,0,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='101' AND SP09='000' " & strSQL2 & StrSQL8
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,0,0,0,1,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='101' AND SP09>'000' " & strSQL2 & StrSQL8
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,0,0,0,0,1,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='806' " & strSQL2 & StrSQL8
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,0,0,0,0,0,1,ST03,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND (CP10='801' OR CP10='802' OR CP10='805') " & strSQL2 & StrSQL8
TestProcess
'StrSQL = "INSERT INTO R020412 " & StrSQL

'cnnConnection.Execute StrSQL

strSql = "SELECT * FROM R020412 WHERE ID='" & strUserNum & "' "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 Then
   InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/10/19
Else
   InsertQueryLog (0) 'Add By Sindy 2010/10/19
   ShowNoData
   Screen.MousePointer = vbDefault
   Exit Sub
End If
'cnnConnection.Execute "UPDATE R020412 SET R085001='',r085002='',R085011='000' WHERE (R085011<'S' OR R085011>'T' )AND ID='" & strUserNum & "' "
'Modify By Cheng 2003/07/01
'廣東所歸為非業務區
'cnnConnection.Execute "UPDATE R020412 SET R085001='非業務區',r085002='非智權人員',R085011='000',R085012='00000' WHERE (R085011<'S' OR R085011>'T' )AND ID='" & strUserNum & "' "
cnnConnection.Execute "UPDATE R020412 SET R085001='非業務區',r085002='非智權人員',R085011='000',R085012='00000' WHERE ((R085011<'S' OR R085011>'T') Or R085011='S91' ) AND ID='" & strUserNum & "' "
PrintData
'Add By Cheng 2002/03/05
ShowPrintOk
Screen.MousePointer = vbDefault
End Sub

Sub TestProcess()
'Add By Cheng 2003/02/05
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
Dim rsB As New ADODB.Recordset
Dim StrSqlB As String
Dim dblTMKindCnt As Double '商品類別數

StrSQL7 = "": strSQL8 = ""
If Len(txt1(0)) <> 0 And Len(txt1(1)) <> 0 Then
'   strSQL7 = strSQL7 + " AND ((CP05>=" & Val(ChangeTStringToWString(txt1(0))) & " AND CP05<=" & Val(ChangeTStringToWString(txt1(1))) & ") OR (CP27>=" & Val(ChangeTStringToWString(txt1(0))) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(1))) & ")) "
   StrSQL7 = StrSQL7 + " AND (CP05>=" & Val(ChangeTStringToWString(txt1(0))) & " AND CP05<=" & Val(ChangeTStringToWString(txt1(1))) & ") "
   strSQL8 = strSQL8 + " AND (CP27>=" & Val(ChangeTStringToWString(txt1(0))) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(1))) & ") "
Else
   If Len(Trim(txt1(0))) <> 0 And Len(txt1(1)) = 0 Then
'      strSQL7 = strSQL7 + " AND (CP05>=" & Val(ChangeTStringToWString(txt1(0))) & " Or CP27 >= " & Val(ChangeTStringToWString(txt1(0))) & ") "
      StrSQL7 = StrSQL7 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(0))) & " "
      strSQL8 = strSQL8 + " AND CP27>= " & Val(ChangeTStringToWString(txt1(0))) & " "
   Else
      If Len(Trim(txt1(0))) = 0 And Len(txt1(1)) <> 0 Then
'         strSQL7 = strSQL7 + " AND (CP05<=" & Val(ChangeTStringToWString(txt1(1))) & " Or CP27 <= " & Val(ChangeTStringToWString(txt1(1))) & ") "
         StrSQL7 = StrSQL7 + " AND CP05<=" & Val(ChangeTStringToWString(txt1(1))) & " "
         strSQL8 = strSQL8 + " AND CP27<= " & Val(ChangeTStringToWString(txt1(1))) & " "
      End If
   End If
End If
'Modify By Cheng 2002/02/22
'                strSQL = "SELECT NVL(A0902,A0903),ST02,0,0,0,0,0,0,0,0,ST03,CP10,TM10,CP09,cp05,cp27 FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND ST03=A0901(+) " & strSQL1 & StrSQL7
'strSQL = strSQL & " union all select NVL(A0902,A0903),ST02,0,0,0,0,0,0,0,0,ST03,CP10,SP09,CP09,cp05,cp27 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=ST01(+) AND ST03=A0901(+) " & strSQL2 & StrSQL7
'Modify By Cheng 2002/04/08
'業務區抓CP12
'strSQL = "SELECT NVL(A0902,A0903),DECODE(ST04,'1',ST02,'離職'),0,0,0,0,0,0,0,0,ST03,CP10,TM10,CP09,cp05,cp27,DECODE(ST03,'2','ZZZZZ',ST01),ST04 " & _
'         " FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND '101'=CP10 AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "' ) " & _
'         strSQL1 & StrSQL7
'Modify By Cheng 2003/12/30
'strSQL = "SELECT NVL(A0902,A0903),DECODE(ST04,'1',ST02,'離職'),0,0,0,0,0,0,0,0,CP12,CP10,TM10,CP09,cp05,cp27,DECODE(ST04,'2','ZZZZZ',ST01),ST04 " & _
'         " FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND CP12=A0901(+) AND '101'=CP10 AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "' ) " & _
'         " AND (CP26 IS NULL OR CP26='') AND (CP57 IS NULL OR CP57='') " & strSQL1 & strSQL7
'Modify By Cheng 2004/04/13
'中四區的資料合併至中二區
'收文
'strSQL = "SELECT NVL(A0902,A0903),DECODE(ST04,'1',ST02,'離職'),0,0,0,0,0,0,0,0,CP12,CP10,TM10,CP09,cp05, 0,DECODE(ST04,'2','ZZZZZ',ST01),ST04, TM09 " & _
'         " FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND CP12=A0901(+) AND '101'=CP10 AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "' ) " & _
'         " AND (CP26 IS NULL OR CP26='') AND (CP57 IS NULL OR CP57='') " & strSQL1 & strSQL7
'Modified by Morgan 2016/2/2 105年1月起又有中四區
'strSql = "SELECT NVL(A0902,A0903),DECODE(ST04,'1',ST02,'離職'),0,0,0,0,0,0,0,0,Decode(CP12,'S24','S22',CP12) As CP12,CP10,TM10,CP09,cp05, 0,DECODE(ST04,'2','ZZZZZ',ST01),ST04, TM09 " & _
         " FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND Decode(CP12,'S24','S22',CP12)=A0901(+) AND '101'=CP10 AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "' ) " & _
         " AND (CP26 IS NULL OR CP26='') AND (CP57 IS NULL OR CP57='') " & strSQL1 & StrSQL7
strSql = "SELECT NVL(A0902,A0903),DECODE(ST04,'1',ST02,'離職'),0,0,0,0,0,0,0,0,CP12,CP10,TM10,CP09,cp05, 0,DECODE(ST04,'2','ZZZZZ',ST01),ST04, TM09 " & _
         " FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND CP12=A0901(+) AND '101'=CP10 AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "' ) " & _
         " AND (CP26 IS NULL OR CP26='') AND (CP57 IS NULL OR CP57='') " & strSQL1 & StrSQL7
'發文
'strSQL = strSQL & " Union All SELECT NVL(A0902,A0903),DECODE(ST04,'1',ST02,'離職'),0,0,0,0,0,0,0,0,CP12,CP10,TM10,CP09, 0,cp27,DECODE(ST04,'2','ZZZZZ',ST01),ST04, TM09 " & _
'         " FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND CP12=A0901(+) AND '101'=CP10 AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "' ) " & _
'         " AND (CP26 IS NULL OR CP26='') AND (CP57 IS NULL OR CP57='') " & strSQL1 & strSQL8
'Modified by Morgan 2016/2/2 105年1月起又有中四區
'strSql = strSql & " Union All SELECT NVL(A0902,A0903),DECODE(ST04,'1',ST02,'離職'),0,0,0,0,0,0,0,0,Decode(CP12,'S24','S22',CP12) As CP12,CP10,TM10,CP09, 0,cp27,DECODE(ST04,'2','ZZZZZ',ST01),ST04, TM09 " & _
         " FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND Decode(CP12,'S24','S22',CP12)=A0901(+) AND '101'=CP10 AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "' ) " & _
         " AND (CP26 IS NULL OR CP26='') AND (CP57 IS NULL OR CP57='') " & strSQL1 & strSQL8
strSql = strSql & " Union All SELECT NVL(A0902,A0903),DECODE(ST04,'1',ST02,'離職'),0,0,0,0,0,0,0,0,CP12,CP10,TM10,CP09, 0,cp27,DECODE(ST04,'2','ZZZZZ',ST01),ST04, TM09 " & _
         " FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND CP12=A0901(+) AND '101'=CP10 AND (TM10='" & 台灣國家代號 & "' OR TM10='" & 大陸國家代號 & "' ) " & _
         " AND (CP26 IS NULL OR CP26='') AND (CP57 IS NULL OR CP57='') " & strSQL1 & strSQL8
'End
CheckOC
With adoRecordset
   .CursorLocation = adUseClient
   .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 Then
      .MoveFirst
      Do While .EOF = False
         'Modify By Cheng 2002/02/25
'         For i = 0 To 15
         For i = 0 To 17
            strTemp(i) = CheckStr(.Fields(i))
         Next i
         Select Case strTemp(11) '判斷案件性質
         Case "101" '申請
            If strTemp(12) = 台灣國家代號 Then '申請國家為台灣
               '符合收文
               If IIf(Len(txt1(0)) <> 0, IIf(strTemp(14) >= ChangeTStringToWString(txt1(0)), True, False), True) And IIf(Len(txt1(1)) <> 0, IIf(strTemp(14) <= ChangeTStringToWString(txt1(1)), True, False), True) And .Fields("CP09").Value < "B" Then
                    '若無商品類別資料
                    If "" & .Fields("TM09").Value = "" Then
                        dblTMKindCnt = 1
                    '若有商品類別資料
                    Else
                        dblTMKindCnt = UBound(Split("" & .Fields("TM09").Value, ",")) + 1
                    End If
'                  strSQL = "('" & chgsql(strTemp(0)) & "','" & chgsql(strTemp(1)) & "',1,0,0,0,0,0,0,0,'" & chgsql(strTemp(10)) & "','" & strUserNum & "')"
'                  strSQL = "('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "',1,0,0,0,0,0,0,0,'" & ChgSQL(strTemp(10)) & "','" & strUserNum & "','" & ChgSQL(strTemp(16)) & "','" & ChgSQL(strTemp(17)) & "')"
                  strSql = "('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & dblTMKindCnt & ",0,0,0,0,0,0,0,'" & ChgSQL(strTemp(10)) & "','" & strUserNum & "','" & ChgSQL(strTemp(16)) & "','" & ChgSQL(strTemp(17)) & "')"
                  strSql = "INSERT INTO R020412 VALUES " & strSql
                  cnnConnection.Execute strSql
               End If
               '符合發文
               'Modify By Cheng 2002/04/08
'               If IIf(Len(txt1(0)) <> 0, IIf(strTemp(15) >= ChangeTStringToWString(txt1(0)), True, False), True) And IIf(Len(txt1(1)) <> 0, IIf(strTemp(15) <= ChangeTStringToWString(txt1(1)), True, False), True) And .Fields("CP09").Value < "C" Then
               If IIf(Len(txt1(0)) <> 0, IIf(strTemp(15) >= ChangeTStringToWString(txt1(0)), True, False), True) And IIf(Len(txt1(1)) <> 0, IIf(strTemp(15) <= ChangeTStringToWString(txt1(1)), True, False), True) And .Fields("CP09").Value < "B" Then
                    '若無商品類別資料
                    If "" & .Fields("TM09").Value = "" Then
                        dblTMKindCnt = 1
                    '若有商品類別資料
                    Else
                        dblTMKindCnt = UBound(Split("" & .Fields("TM09").Value, ",")) + 1
                    End If
'                  strSQL = "('" & chgsql(strTemp(0)) & "','" & chgsql(strTemp(1)) & "',0,0,0,0,1,0,0,0,'" & chgsql(strTemp(10)) & "','" & strUserNum & "')"
'                  strSQL = "('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "',0,0,0,0,1,0,0,0,'" & ChgSQL(strTemp(10)) & "','" & strUserNum & "','" & ChgSQL(strTemp(16)) & "','" & ChgSQL(strTemp(17)) & "')"
                  strSql = "('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "',0,0,0,0," & dblTMKindCnt & ",0,0,0,'" & ChgSQL(strTemp(10)) & "','" & strUserNum & "','" & ChgSQL(strTemp(16)) & "','" & ChgSQL(strTemp(17)) & "')"
                  strSql = "INSERT INTO R020412 VALUES " & strSql
                  cnnConnection.Execute strSql
               End If
            ElseIf strTemp(12) = 大陸國家代號 Then
               '符合收文
               If IIf(Len(txt1(0)) <> 0, IIf(strTemp(14) >= ChangeTStringToWString(txt1(0)), True, False), True) And IIf(Len(txt1(1)) <> 0, IIf(strTemp(14) <= ChangeTStringToWString(txt1(1)), True, False), True) And .Fields("CP09") < "B" Then
'                  strSQL = "('" & chgsql(strTemp(0)) & "','" & chgsql(strTemp(1)) & "',0,1,0,0,0,0,0,0,'" & chgsql(strTemp(10)) & "','" & strUserNum & "')"
                  strSql = "('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "',0,1,0,0,0,0,0,0,'" & ChgSQL(strTemp(10)) & "','" & strUserNum & "','" & ChgSQL(strTemp(16)) & "','" & ChgSQL(strTemp(17)) & "')"
                  strSql = "INSERT INTO R020412  VALUES " & strSql
                  cnnConnection.Execute strSql
               End If
               '符合發文
               'Modify By Cheng 2002/04/08
'               If IIf(Len(txt1(0)) <> 0, IIf(strTemp(15) >= ChangeTStringToWString(txt1(0)), True, False), True) And IIf(Len(txt1(1)) <> 0, IIf(strTemp(15) <= ChangeTStringToWString(txt1(1)), True, False), True) And .Fields("CP09") < "C" Then
               If IIf(Len(txt1(0)) <> 0, IIf(strTemp(15) >= ChangeTStringToWString(txt1(0)), True, False), True) And IIf(Len(txt1(1)) <> 0, IIf(strTemp(15) <= ChangeTStringToWString(txt1(1)), True, False), True) And .Fields("CP09") < "B" Then
'                  strSQL = "('" & chgsql(strTemp(0)) & "','" & chgsql(strTemp(1)) & "',0,0,0,0,0,1,0,0,'" & chgsql(strTemp(10)) & "','" & strUserNum & "')"
                  strSql = "('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "',0,0,0,0,0,1,0,0,'" & ChgSQL(strTemp(10)) & "','" & strUserNum & "','" & ChgSQL(strTemp(16)) & "','" & ChgSQL(strTemp(17)) & "')"
                  strSql = "INSERT INTO R020412  VALUES " & strSql
                  cnnConnection.Execute strSql
               End If
            End If
'         Case "806"
'               '符合收文
'               If IIf(Len(txt1(0)) <> 0, IIf(strTemp(14) >= ChangeTStringToWString(txt1(0)), True, False), True) And IIf(Len(txt1(1)) <> 0, IIf(strTemp(14) <= ChangeTStringToWString(txt1(1)), True, False), True) Then
'                  strSQL = "('" & chgsql(strTemp(0)) & "','" & chgsql(strTemp(1)) & "',0,0,1,0,0,0,0,0,'" & chgsql(strTemp(10)) & "','" & strUserNum & "')"
'                  strSQL = "INSERT INTO R020412  VALUES " & strSQL
'                  cnnConnection.Execute strSQL
'               End If
'               '符合發文
'               If IIf(Len(txt1(0)) <> 0, IIf(strTemp(15) >= ChangeTStringToWString(txt1(0)), True, False), True) And IIf(Len(txt1(1)) <> 0, IIf(strTemp(15) <= ChangeTStringToWString(txt1(1)), True, False), True) Then
'                  strSQL = "('" & chgsql(strTemp(0)) & "','" & chgsql(strTemp(1)) & "',0,0,0,0,0,0,1,0,'" & chgsql(strTemp(10)) & "','" & strUserNum & "')"
'                  strSQL = "INSERT INTO R020412  VALUES " & strSQL
'                  cnnConnection.Execute strSQL
'               End If
'         Case "801", "802", "805"
'               '符合收文
'               If IIf(Len(txt1(0)) <> 0, IIf(strTemp(14) >= ChangeTStringToWString(txt1(0)), True, False), True) And IIf(Len(txt1(1)) <> 0, IIf(strTemp(14) <= ChangeTStringToWString(txt1(1)), True, False), True) Then
'                  strSQL = "('" & chgsql(strTemp(0)) & "','" & chgsql(strTemp(1)) & "',0,0,0,1,0,0,0,0,'" & chgsql(strTemp(10)) & "','" & strUserNum & "')"
'                  strSQL = "INSERT INTO R020412  VALUES " & strSQL
'                  cnnConnection.Execute strSQL
'               End If
'               '符合發文
'               If IIf(Len(txt1(0)) <> 0, IIf(strTemp(15) >= ChangeTStringToWString(txt1(0)), True, False), True) And IIf(Len(txt1(1)) <> 0, IIf(strTemp(15) <= ChangeTStringToWString(txt1(1)), True, False), True) Then
'                  strSQL = "('" & chgsql(strTemp(0)) & "','" & chgsql(strTemp(1)) & "',0,0,0,0,0,0,0,1,'" & chgsql(strTemp(10)) & "','" & strUserNum & "')"
'                  strSQL = "INSERT INTO R020412  VALUES " & strSQL
'                  cnnConnection.Execute strSQL
'               End If
         Case Else
         End Select
         .MoveNext
      Loop
   End If
End With
'StrSQL = "SELECT NVL(A0902,A0903),ST02,1,0,0,0,0,0,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='101' AND TM10='000' " & strSQL1 & StrSQL7
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,1,0,0,0,0,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='101' AND TM10>'000' " & strSQL1 & StrSQL7
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,1,0,0,0,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='806' " & strSQL1 & StrSQL7
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,0,1,0,0,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND (CP10='801' OR CP10='802' OR CP10='805') " & strSQL1 & StrSQL7
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,0,0,1,0,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='101' AND TM10='000' " & strSQL1 & StrSQL8
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,0,0,0,1,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='101' AND TM10>'000' " & strSQL1 & StrSQL8
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,0,0,0,0,1,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='806' " & strSQL1 & StrSQL8
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,0,0,0,0,0,1,ST03,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,STAFF,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND (CP10='801' OR CP10='802' OR CP10='805') " & strSQL1 & StrSQL8
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,1,0,0,0,0,0,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='101' AND SP09='000' " & strSQL2 & StrSQL7
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,1,0,0,0,0,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='101' AND SP09>'000' " & strSQL2 & StrSQL7
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,1,0,0,0,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='806' " & strSQL2 & StrSQL7
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,0,1,0,0,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND (CP10='801' OR CP10='802' OR CP10='805') " & strSQL2 & StrSQL7
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,0,0,1,0,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='101' AND SP09='000' " & strSQL2 & StrSQL8
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,0,0,0,1,0,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='101' AND SP09>'000' " & strSQL2 & StrSQL8
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,0,0,0,0,1,0,ST03,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND CP10='806' " & strSQL2 & StrSQL8
'StrSQL = StrSQL + " UNION ALL SELECT NVL(A0902,A0903),ST02,0,0,0,0,0,0,0,1,ST03,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,STAFF,ACC090 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=ST01(+) AND ST03=A0901(+) AND (CP10='801' OR CP10='802' OR CP10='805') " & strSQL2 & StrSQL8

'StrSQL = "INSERT INTO R020412 " & StrSQL
'cnnConnection.Execute StrSQL
'Add By Cheng 2003/02/05
'建立無收發文的智權人員資料
'Modify By Cheng 2003/02/17
'ST03為人事歸屬的部門, ST15為收文歸屬的部門
'strSQLA = "Select * From Staff,Acc090 Where ST03=A0901(+) And ST01>='60000' And SUBSTR(ST03,1,2)>='S1' And SUBSTR(ST03,1,2)<='S4' And ST04='1' "
'Modify By Cheng 2003/03/03
'剔除USER
'strSQLA = "Select * From Staff,Acc090 Where ST15=A0901(+) And ST01>='60000' And SUBSTR(ST03,1,2)>='S1' And SUBSTR(ST03,1,2)<='S4' And ST04='1' "
'Modify By Cheng 2003/08/08
'strSQLA = "Select * From Staff,Acc090 Where ST15=A0901(+) And ST01>='60000' And ST01<>'USER' And SUBSTR(ST03,1,2)>='S1' And SUBSTR(ST03,1,2)<='S4' And ST04='1' "
'Modify By Sindy 2010/11/26
'StrSQLa = "Select * From Staff,Acc090 Where ST15=A0901(+) And ST01>='60000' And ST01<='999999' And SUBSTR(ST15,1,2)>='S1' And SUBSTR(ST15,1,2)<='S4' And ST04='1' "
StrSQLa = "Select * From Staff,Acc090 Where ST15=A0901(+) And ST01>='60000' And ST01<'F' And SUBSTR(ST15,1,2)>='S1' And SUBSTR(ST15,1,2)<='S4' And ST04='1' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    While Not rsA.EOF
        StrSqlB = "Select Count(*) From R020412 Where R085012='" & rsA("ST01").Value & "' And ID='" & strUserNum & "' "
        rsB.CursorLocation = adUseClient
        rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
        If rsB.Fields(0).Value = 0 Then
            'Modify By Cheng 2003/02/17
'            strSQL = "Insert Into R020412 Values('" & rsA("A0902").Value & "','" & rsA("ST02").Value & "',0,0,0,0,0,0,0,0,'" & rsA("ST03").Value & "','" & strUserNum & "','" & rsA("ST01").Value & "','" & rsA("ST04").Value & "') "
            strSql = "Insert Into R020412 Values('" & rsA("A0902").Value & "','" & rsA("ST02").Value & "',0,0,0,0,0,0,0,0,'" & rsA("ST15").Value & "','" & strUserNum & "','" & rsA("ST01").Value & "','" & rsA("ST04").Value & "') "
            cnnConnection.Execute strSql
        End If
        If rsB.State <> adStateClosed Then rsB.Close
        Set rsB = Nothing
        rsA.MoveNext
    Wend
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Sub

Sub PrintData()
'Add By Cheng 2002/11/20
Dim intLN As Integer

'Modify By Cheng 2002/02/25
'strSQL = "SELECT R085001,R085002,SUM(R085003),SUM(R085004),SUM(R085005),SUM(R085006),SUM(R085007),SUM(R085008),SUM(R085009),SUM(R085010),R085011 FROM R020412 WHERE ID='" & strUserNum & "' GROUP BY R085011,R085001,R085002 ORDER BY R085011,R085001,R085002 "
strSql = "SELECT R085001,R085002,SUM(R085003),SUM(R085004),R085001,R085002,SUM(R085007),SUM(R085008),R085011,R085012,R085013 FROM R020412 WHERE ID='" & strUserNum & "' GROUP BY R085011,R085001,R085002,R085012,R085013 ORDER BY R085011,R085013,R085012 "
CheckOC
Page = 1
strTemp3 = "": strTemp3_1 = "": StrTemp4 = ""
m_strSaleZone1 = "": m_strSaleZone2 = ""
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        'Modify By Cheng 2002/11/20
'        PrintTitle
        PrintTitle Page
        'Add By Cheng 2002/11/20
        intLN = 0
'        strTemp3 = CheckStr(.Fields(10))
        strTemp3 = CheckStr(.Fields(8))
      'Modify By Cheng 2002/04/08
'        strTemp3_1 = Left("" & .Fields(0), 1)
        'Modify By Sindy 2011/7/15
        'strTemp3_1 = IIf(Left("" & strTemp(0), 2) = "台中" Or Left("" & strTemp(0), 2) = "台南", Mid("" & strTemp(0), 2, 1), Left("" & strTemp(0), 1))
        strTemp3_1 = IIf(Left("" & strTemp(0), 2) = "台北" Or Left("" & strTemp(0), 2) = "台中" Or Left("" & strTemp(0), 2) = "台南", Mid("" & strTemp(0), 2, 1), Left("" & strTemp(0), 1))
        StrTemp4 = .Fields(8)
         m_strSaleZone1 = .Fields(0)
         m_strSaleZone2 = .Fields(4)
        Do While .EOF = False
            For i = 0 To 10
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If StrTemp4 <> strTemp(8) Then
                If StrTemp4 <> "000" Then
                  ShowLine
                  PrintEnd (2)
                'Add By Cheng 2002/11/20
                intLN = intLN + 1
                  ShowLine
                End If
                StrTemp4 = strTemp(8)
            End If
            'Add By Cheng 2002/11/20
            If intLN >= m_intLastLn Then
                Printer.NewPage
                Page = Page + 1
                PrintTitle Page
                intLN = 0
            End If
'            If StrToStr(strTemp3, 1) <> StrToStr(strTemp(10), 1) Then
            If StrToStr(strTemp3, 1) <> StrToStr(strTemp(8), 1) Then
                ShowLine
                If strTemp3 <> "000" Then
                  PrintEnd (0)
                'Add By Cheng 2002/11/20
                intLN = intLN + 1
                  ShowLine
                End If
            'Add By Cheng 2002/11/20
            If intLN >= m_intLastLn Then
                Printer.NewPage
                Page = Page + 1
                PrintTitle Page
                intLN = 0
            End If
'                strTemp3 = strTemp(10)
                strTemp3 = strTemp(8)
               'Modify By Cheng 2002/04/08
'                strTemp3_1 = Left("" & strTemp(0), 1)
                'Modify By Sindy 2011/7/15
                'strTemp3_1 = IIf(Left("" & strTemp(0), 2) = "台中" Or Left("" & strTemp(0), 2) = "台南", Mid("" & strTemp(0), 2, 1), Left("" & strTemp(0), 1))
                strTemp3_1 = IIf(Left("" & strTemp(0), 2) = "台北" Or Left("" & strTemp(0), 2) = "台中" Or Left("" & strTemp(0), 2) = "台南", Mid("" & strTemp(0), 2, 1), Left("" & strTemp(0), 1))
            End If
            strTemp(0) = StrToStr(strTemp(0), 4)
            'strTemp(1) = StrToStr(strTemp(1), 4)   'CANCEL BY SONIA 2015/11/2 收文明細的人名不必截取, 因為發文明細未截取
            PrintDatil
            'Add By Cheng 2002/11/20
            intLN = intLN + 1
            If intLN >= m_intLastLn Then
                Printer.NewPage
                Page = Page + 1
                PrintTitle Page
                intLN = 0
            End If
'            If iPrint > 14000 Then
'                Page = Page + 1
'                Printer.NewPage
'                PrintTitle
'            End If
            .MoveNext
        Loop
    End If
End With
ShowLine
PrintEnd (2) '小計
ShowLine
'Add By Cheng 2002/11/20
intLN = intLN + 1
If intLN >= m_intLastLn Then
    Printer.NewPage
    Page = Page + 1
    PrintTitle Page
    intLN = 0
End If

PrintEnd (0) '各區合計
ShowLine
'Add By Cheng 2002/11/20
intLN = intLN + 1
If intLN >= m_intLastLn Then
    Printer.NewPage
    Page = Page + 1
    PrintTitle Page
    intLN = 0
End If

PrintEnd (1) '全所總計
ShowLine
Printer.EndDoc
CheckOC
End Sub

Sub PrintEnd(Strindex As Integer)
Select Case Strindex
Case 0 '各區合計
     strSql = "SELECT '" & strTemp3_1 & "區合計','',SUM(R085003),SUM(R085004),'" & strTemp3_1 & "區合計','',SUM(R085007),SUM(R085008) FROM R020412 WHERE ID='" & strUserNum & "' AND substr(R085011,1,2)='" & Mid(strTemp3, 1, 2) & "' "
Case 1 '全所總計
'     strSQL = "SELECT '全所總計','',SUM(R085003),SUM(R085004),SUM(R085005),SUM(R085006),SUM(R085007),SUM(R085008),SUM(R085009),SUM(R085010),'' FROM R020412 WHERE ID='" & strUserNum & "' "
     strSql = "SELECT '全所總計','',SUM(R085003),SUM(R085004),'全所總計','',SUM(R085007),SUM(R085008) FROM R020412 WHERE ID='" & strUserNum & "' "
Case 2 '小計
     strSql = "SELECT '小計','',SUM(R085003),SUM(R085004),'小計','',SUM(R085007),SUM(R085008) FROM R020412 WHERE ID='" & strUserNum & "' AND substr(R085011,1,3)='" & Mid(StrTemp4, 1, 3) & "' "
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
'            For i = 0 To 10
            For i = 0 To 7
                StrTemp7(i) = CheckStr(.Fields(i))
            Next i
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(1)
'            For i = 2 To 9
            For i = 2 To 3
                Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(StrTemp7(i), "####0"))
                Printer.CurrentY = iPrint
                Printer.Print Format(StrTemp7(i), "####0")
            Next i
            
            Printer.CurrentX = PLeft(4)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(4)
            Printer.CurrentX = PLeft(5)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(5)
            
            For i = 6 To 7
                Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(StrTemp7(i), "####0"))
                Printer.CurrentY = iPrint
                Printer.Print Format(StrTemp7(i), "####0")
            Next i
            
            Printer.Line (PLeft(0), iPrint)-(PLeft(0), iPrint + 230)
            Printer.Line (PLeft(4) - 50, iPrint)-(PLeft(4) - 50, iPrint + 230)
            Printer.Line (PLeft(4), iPrint)-(PLeft(4), iPrint + 230)
            'Modify By Sindy 98/04/16
            Printer.Line (11300, iPrint)-(11300, iPrint + 230)
            
            iPrint = iPrint + 230
'            If iPrint >= 14000 Then
'                Page = Page + 1
'                Printer.NewPage
'                PrintTitle
'            End If
            .MoveNext
        Loop
    End If
End With
CheckOC2
End Sub

'Modify By Cheng 2002/11/20
'Sub PrintTitle()
Sub PrintTitle(intPage As Integer)
GetPleft
iPrint = 300 'Modify By Sindy 98/04/16
Printer.Orientation = 1
Printer.Font.Name = "細明體"
Printer.Font.Size = 12
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 4500
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "各區收/發文件數明細表"
Printer.Font.Size = 10
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 300
Printer.CurrentX = 500 'Modify By Sindy 98/04/16
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
'Add By Cheng 2002/11/20
Printer.CurrentX = 9500
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")

iPrint = iPrint + 230
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "日期：" & Format(ChangeTStringToTDateString(txt1(0)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(1))
Printer.CurrentX = 9500
Printer.CurrentY = iPrint
'Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
Printer.Print "頁　　數：" & intPage
iPrint = iPrint + 230
Printer.CurrentX = 8400
Printer.CurrentY = iPrint
'Printer.Print "頁    次：" & str(Page)
'iPrint = iPrint + 230
Printer.CurrentX = 500
Printer.CurrentY = iPrint
'Modify By Sindy 98/04/16
Printer.Line (500, iPrint - 20)-(11300, iPrint - 20)
'iPrint = iPrint + 300
'If iPrint >= 14000 Then
'    Page = Page + 1
'    Printer.NewPage
'    PrintTitle
'    Exit Sub
'End If
Printer.CurrentX = ((PLeft(2) - PLeft(1)) / 2) + PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "收文明細"
Printer.CurrentX = ((PLeft(6) - PLeft(5)) / 2) + PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "發文明細"

Printer.Line (PLeft(0), iPrint)-(PLeft(0), iPrint + 230)
Printer.Line (PLeft(4) - 50, iPrint)-(PLeft(4) - 50, iPrint + 230)
Printer.Line (PLeft(4), iPrint)-(PLeft(4), iPrint + 230)
'Modify By Sindy 98/04/16
Printer.Line (11300, iPrint)-(11300, iPrint + 230)

iPrint = iPrint + 230
'Printer.Line (0, iPrint)-(PLeft(4) - 300, iPrint)
'Printer.Line (PLeft(4), iPrint)-(10800 - 300, iPrint)
'Modify By Sindy 98/04/16
Printer.Line (500, iPrint - 20)-(11300, iPrint - 20)
'iPrint = iPrint + 220
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "業務區別"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "國內"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "大陸"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "業務區別"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "國內"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "大陸"
'Printer.CurrentX = PLeft(8)
'Printer.CurrentY = iPrint
'Printer.Print "著作權"
'Printer.CurrentX = PLeft(9)
'Printer.CurrentY = iPrint
'Printer.Print "其他"

Printer.Line (PLeft(0), iPrint)-(PLeft(0), iPrint + 230)
Printer.Line (PLeft(4) - 50, iPrint)-(PLeft(4) - 50, iPrint + 230)
Printer.Line (PLeft(4), iPrint)-(PLeft(4), iPrint + 230)
'Modify By Sindy 98/04/16
Printer.Line (11300, iPrint)-(11300, iPrint + 230)

iPrint = iPrint + 230
'If iPrint >= 14000 Then
'    Page = Page + 1
'    Printer.NewPage
'    PrintTitle
'    Exit Sub
'End If
Printer.CurrentX = 500
Printer.CurrentY = iPrint
'Modify By Sindy 98/04/16
Printer.Line (500, iPrint - 20)-(11300, iPrint - 20)
'iPrint = iPrint + 300
'If iPrint >= 14000 Then
'    Page = Page + 1
'    Printer.NewPage
'    PrintTitle
'    Exit Sub
'End If
End Sub

Sub PrintDatil()
'收文明細
'業務區別
If m_strSaleZone1 <> strTemp(0) Then
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(0)
   m_strSaleZone1 = strTemp(0)
End If

'智權人員
'Modified by Morgan 2022/5/3 修正Unidcode問題
'Printer.CurrentX = PLeft(1)
'Printer.CurrentY = iPrint
'Printer.Print strTemp(1)
PUB_PrintUnicodeText strTemp(1), 1& * PLeft(1), 1& * iPrint, 0
'end 2022/5/3

'收文件數
'For i = 2 To 9
For i = 2 To 3
    Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(strTemp(i), "####0"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i), "####0")
Next i

'發文明細
'業務區別
If m_strSaleZone2 <> strTemp(4) Then
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(4)
   m_strSaleZone2 = strTemp(4)
End If
'智權人員
'Modified by Morgan 2022/5/3 修正Unidcode問題
'Printer.CurrentX = PLeft(5)
'Printer.CurrentY = iPrint
'Printer.Print strTemp(5)
PUB_PrintUnicodeText strTemp(5), 1& * PLeft(5), 1& * iPrint, 0
'end 2022/5/3

'發文件數
For i = 6 To 7
    Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(strTemp(i), "####0"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i), "####0")
Next i

Printer.Line (PLeft(0), iPrint)-(PLeft(0), iPrint + 230)
Printer.Line (PLeft(4) - 50, iPrint)-(PLeft(4) - 50, iPrint + 230)
Printer.Line (PLeft(4), iPrint)-(PLeft(4), iPrint + 230)
'Modify By Sindy 98/04/16
Printer.Line (11300, iPrint)-(11300, iPrint + 230)

iPrint = iPrint + 230
End Sub

Sub GetPleft()
Erase PLeft
'PLeft(0) = 0
''PLeft(1) = 2000
''PLeft(2) = 4000
''PLeft(3) = 6000
''PLeft(4) = 8000
''PLeft(5) = 10000
''PLeft(6) = 12000
''PLeft(7) = 14000
''PLeft(8) = 16000
''PLeft(9) = 18000
'PLeft(1) = 1400
'PLeft(2) = 2800
'PLeft(3) = 4200
'PLeft(4) = 5600
'PLeft(5) = 7000
'PLeft(6) = 8400
'PLeft(7) = 9800
'Modify By Sindy 98/04/16
PLeft(0) = 500
PLeft(1) = 1900
PLeft(2) = 3300
PLeft(3) = 4700
PLeft(4) = 6100
PLeft(5) = 7500
PLeft(6) = 8900
PLeft(7) = 10300
End Sub

Sub ShowLine()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
'Modify By Sindy 98/04/16
Printer.Line (500, iPrint - 20)-(11300, iPrint - 20)
'iPrint = iPrint + 300
'If iPrint >= 14000 Then
'    Page = Page + 1
'    Printer.NewPage
'    PrintTitle
'End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, Combo1, strPrinter 'Add By Sindy 2015/7/3
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2015/7/3
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   '2015/7/3 END
   
   Set frm020412 = Nothing
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
Case 0, 1
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 1 Then
     If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
    End If

Case Else
End Select
End Sub

