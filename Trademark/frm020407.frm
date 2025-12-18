VERSION 5.00
Begin VB.Form frm020407 
   BorderStyle     =   1  '單線固定
   Caption         =   "商爭案承辦人收/發文統計表"
   ClientHeight    =   3060
   ClientLeft      =   990
   ClientTop       =   3410
   ClientWidth     =   3960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   3960
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   528
      Left            =   75
      TabIndex        =   14
      Top             =   1980
      Width           =   3825
      Begin VB.ComboBox Combo1 
         Height          =   276
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   6
         Top             =   180
         Width           =   2880
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   180
         Left            =   105
         TabIndex        =   15
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3144
      TabIndex        =   8
      Top             =   36
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2352
      TabIndex        =   7
      Top             =   36
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   2010
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1635
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1020
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1635
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   2010
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1290
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1020
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1290
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1020
      MaxLength       =   1
      TabIndex        =   1
      Top             =   960
      Width           =   270
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1020
      TabIndex        =   0
      Top             =   600
      Width           =   1920
   End
   Begin VB.Line Line2 
      X1              =   1485
      X2              =   2235
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Line Line1 
      X1              =   1440
      X2              =   2700
      Y1              =   1455
      Y2              =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "(1.收文  2.發文)"
      Height          =   180
      Index           =   4
      Left            =   1395
      TabIndex        =   13
      Top             =   990
      Width           =   1410
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   3
      Left            =   105
      TabIndex        =   12
      Top             =   1665
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "日期："
      Height          =   180
      Index           =   2
      Left            =   105
      TabIndex        =   11
      Top             =   1350
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "列印別："
      Height          =   180
      Index           =   1
      Left            =   105
      TabIndex        =   10
      Top             =   990
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   105
      TabIndex        =   9
      Top             =   615
      Width           =   915
   End
End
Attribute VB_Name = "frm020407"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/9 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, intS As Integer
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 40) As String, strTemp3 As String, TestOk As Boolean, StrTemp7(1 To 9) As String, k As Integer
Dim PLeft(0 To 28) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, PLeft1(1 To 9) As Integer, SeekPrint As Integer, SeekPrintL As Integer
Dim BolEndThisPage As Boolean


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
         If Len(txt1(1)) = 0 Then
            s = MsgBox("列印別不可空白!!", , "USER 輸入錯誤")
            txt1(1).SetFocus
            Exit Sub
         Else
            'Add By Cheng 2002/03/21
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
               Screen.MousePointer = vbHourglass
               Me.Enabled = False
               ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/19 清除查詢印表記錄檔欄位
               Process
               Me.Enabled = True
               Screen.MousePointer = vbDefault
            End If
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
'Add By Cheng 2003/03/05
Dim blnFix1 As Boolean '判斷是否有固定第一頁的資料
Dim blnFix2 As Boolean '判斷是否有固定第二頁的資料
'Add By Cheng 2003/09/03
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim ii As Integer
'Modify By Sindy 2015/1/22
Dim System_ID As String
Dim bolIsChina As Boolean
Dim Title_601 As String
Dim Title_603 As String
Dim Title_605 As String
Dim Title_401 As String
Dim Title_410 As String
Dim Title_403 As String
Dim Title_408 As String
Dim Title_602 As String
Dim Title_604 As String
Dim Title_606 As String
Dim Title_202 As String
Dim Title_406 As String
Dim Title_407 As String
'2015/1/22

'Add By Cheng 2003/03/05
'初始化是否有固定頁資料
blnFix1 = False: blnFix2 = False
Screen.MousePointer = vbHourglass
'刪除暫存檔資料
cnnConnection.Execute "DELETE FROM r020407 WHERE ID='" & strUserNum & "' " '明細
cnnConnection.Execute "DELETE FROM r020407_1 WHERE ID='" & strUserNum & "' " '標題
cnnConnection.Execute "DELETE FROM r020407_2 WHERE ID='" & strUserNum & "' " '合計
strSQL1 = ""
strSQL2 = ""
StrSQL6 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/10/19
End If
'Add By Cheng 2003/07/11
'若非商申收發文時, 若為TF案則不抓後三碼為"000"的資料
strSQL1 = strSQL1 + " AND CP03 <> Decode(CP01,'TF','0','z') AND CP04 <> Decode(CP01,'TF','00','zz') "
StrSQL6 = ""
Select Case Val(txt1(1))
Case 1 '列印收文
   If Len(txt1(2)) <> 0 Then
      StrSQL6 = StrSQL6 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(2))) & ""
   End If
   If Len(Trim(txt1(3))) <> 0 Then
      StrSQL6 = StrSQL6 + " AND CP05<=" & Val(ChangeTStringToWString(txt1(3))) & " "
   End If
   If Len(txt1(2)) <> 0 Or Len(Trim(txt1(3))) <> 0 Then
      pub_QL05 = pub_QL05 & ";收文" & Label1(2) & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/10/19
   End If
    'Add By Cheng 2003/05/06
    '不統計無發文日但有取消收文日的資料
    StrSQL6 = StrSQL6 & " And ((CP27 IS NULL And CP57 IS NULL) Or (CP27 IS NOT NULL And CP57 IS NULL) Or (CP27 IS NOT NULL And CP57 IS NOT NULL)) "
Case 2 '列印發文
   If Len(txt1(2)) <> 0 Then
      StrSQL6 = StrSQL6 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(2))) & ""
   End If
   If Len(Trim(txt1(3))) <> 0 Then
      StrSQL6 = StrSQL6 + " AND CP27<=" & Val(ChangeTStringToWString(txt1(3))) & " "
   End If
   If Len(txt1(2)) <> 0 Or Len(Trim(txt1(3))) <> 0 Then
      pub_QL05 = pub_QL05 & ";發文" & Label1(2) & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/10/19
   End If
Case Else
End Select
If Len(txt1(4)) <> 0 Then
    strSQL1 = strSQL1 + " AND TM10>='" & txt1(4) & "' "
    strSQL2 = strSQL2 + " and SP09>='" & txt1(4) & "' "
End If
If Len(txt1(5)) <> 0 Then
    strSQL1 = strSQL1 + " AND TM10<='" & txt1(5) & "' "
    strSQL2 = strSQL2 + " AND SP09<='" & txt1(5) & "' "
End If
If Len(txt1(4)) <> 0 Or Len(Trim(txt1(5))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(4) & "-" & txt1(5) 'Add By Sindy 2010/10/19
End If
'Add By Cheng 2004/04/29
'抓計件的資料
StrSQL6 = StrSQL6 & " And CP26 Is Null "
'End
CheckOC
'Modify By Cheng 2002/02/06
'取消抓承辦人之ST05為"97"或"17"的資料, 改抓案件性質為"4"或"6"開頭且CP09<"C"的資料
'注意：業務區是CASEPROGRESS的CP12
'strSQL = "SELECT cp14,S1.st03,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,decode(tm10,'000','*',' '),CP09 FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF S1,STAFF S2 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and (S2.ST05='97' OR S2.ST05='17') " & strSQL1 + StrSQL6
'strSQL = strSQL + " union all select cp14,S1.st03,NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,decode(sp09,'000','*',' '),CP09 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1,STAFF S2 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and (S2.ST05='97' OR S2.ST05='17') " & strSQL2 + StrSQL6
'Modify By Cheng 2002/02/08 案件名稱一筆以台灣名為主
'strSQL = "SELECT cp14,CP12,NVL(CPM03,CP10),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,decode(tm10,'000','*',' '),CP09 " & _
'         " FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and (SubStr(CP10,1,1)='4' OR SubStr(CP10,1,1)='6' ) And CP09<'C' " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' Or CP14 IS NULL ) ", " And (SubStr(ST03,1,2)='P2' Or CP14 IS NULL ) ") & strSQL1 + StrSQL6
'strSQL = strSQL + " union all select cp14,CP12,NVL(CPM03,CP10),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,decode(sp09,'000','*',' '),CP09 " & _
'         " FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF " & _
'         " WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and (SubStr(CP10,1,1)='4' OR SubStr(CP10,1,1)='6' ) And CP09<'C' " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' Or CP14 IS NULL ) ", " And (SubStr(ST03,1,2)='P2' Or CP14 IS NULL ) ") & strSQL2 + StrSQL6
'Modify By Cheng 2002/04/04
'加抓案件性質為"202","204","205","207"的條件
'**************** 將業務區改成抓案件進度檔   91.08.15  nick
'edit by nickc 2005/05/12
'StrSql = "SELECT cp14,CP12,Decode(NVL(CPM03,CP10),'（無）',CPM04,NVL(CPM03,CP10)),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,decode(tm10,'000','*',' '),CP09 " & _
'         " FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and (SubStr(CP10,1,1)='4' OR SubStr(CP10,1,1)='6' OR CP10='202' OR CP10='204' OR CP10='205' OR CP10='207' ) And CP09<'C' " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' Or CP14 IS NULL ) ", " And (SubStr(ST03,1,2)='F1' Or CP14 IS NULL ) ") & strSQL1 + StrSQL6
'StrSql = StrSql + " union all select cp14,CP12,Decode(NVL(CPM03,CP10),'（無）',CPM04,NVL(CPM03,CP10)),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,decode(sp09,'000','*',' '),CP09 " & _
'         " FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF " & _
         " WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and (SubStr(CP10,1,1)='4' OR SubStr(CP10,1,1)='6' OR CP10='202' OR CP10='204' OR CP10='205' OR CP10='207' ) And CP09<'C' " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' Or CP14 IS NULL ) ", " And (SubStr(ST03,1,2)='F1' Or CP14 IS NULL ) ") & strSQL2 + StrSQL6
'add by nickc 2007/08/01 商標處改大陸格式
Dim MyCPMStr1 As String
Dim MyCPMStr2 As String
If txt1(4) = "020" And txt1(5) = "020" Then
    MyCPMStr1 = "cpm04"
    MyCPMStr2 = "cpm04"
Else
    MyCPMStr1 = "DECODE(NVL(CPM03,CP10),'（無）',CPM04,NVL(CPM03,CP10))"
    MyCPMStr2 = "DECODE(NVL(CPM03,CP10),'（無）',CPM04,NVL(CPM03,CP10))"
End If
'edit by nickc 2007/08/01
'strSQL = "SELECT cp14,CP12,Decode(NVL(CPM03,CP10),'（無）',CPM04,NVL(CPM03,CP10)),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,decode(substr(s2.st15,1,1),'F',' ','*'),CP09 " & _
'         " FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF s1,staff s2 " & _
'         " WHERE  CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND cp13=s2.st01(+) and CP14=s1.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and (SubStr(CP10,1,1)='4' OR SubStr(CP10,1,1)='6' OR CP10='202' OR CP10='204' OR CP10='205' OR CP10='207' ) And CP09<'C' " & IIf(intPWhere = 國內, " And (SubStr(s1.ST03,1,2)='P2' Or CP14 IS NULL ) ", " And (SubStr(s1.ST03,1,2)='F1' Or CP14 IS NULL ) ") & strSQL1 + StrSQL6
'strSQL = strSQL + " union all select cp14,CP12,Decode(NVL(CPM03,CP10),'（無）',CPM04,NVL(CPM03,CP10)),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,decode(substr(s2.st15,1,1),'F',' ','*'),CP09 " & _
'         " FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF s1,staff s2 " & _
'         " WHERE  CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND cp13=s2.st01(+) and CP14=s1.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and (SubStr(CP10,1,1)='4' OR SubStr(CP10,1,1)='6' OR CP10='202' OR CP10='204' OR CP10='205' OR CP10='207' ) And CP09<'C' " & IIf(intPWhere = 國內, " And (SubStr(s1.ST03,1,2)='P2' Or CP14 IS NULL ) ", " And (SubStr(s1.ST03,1,2)='F1' Or CP14 IS NULL ) ") & strSQL2 + StrSQL6
'edit by nickc 2007/11/06 閱卷一率放商爭
'strSQL = "SELECT cp14,CP12," & MyCPMStr1 & ",1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,decode(substr(s2.st15,1,1),'F',' ','*'),CP09 " & _
'         " FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF s1,staff s2 " & _
'         " WHERE  CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND cp13=s2.st01(+) and CP14=s1.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and (SubStr(CP10,1,1)='4' OR SubStr(CP10,1,1)='6' OR CP10='202' OR CP10='204' OR CP10='205' OR CP10='207' ) And CP09<'C' " & IIf(intPWhere = 國內, " And (SubStr(s1.ST03,1,2)='P2' Or CP14 IS NULL ) ", " And (SubStr(s1.ST03,1,2)='F1' Or CP14 IS NULL ) ") & strSQL1 + StrSQL6
'strSQL = strSQL + " union all select cp14,CP12," & MyCPMStr2 & ",1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,decode(substr(s2.st15,1,1),'F',' ','*'),CP09 " & _
'         " FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF s1,staff s2 " & _
'         " WHERE  CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND cp13=s2.st01(+) and CP14=s1.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and (SubStr(CP10,1,1)='4' OR SubStr(CP10,1,1)='6' OR CP10='202' OR CP10='204' OR CP10='205' OR CP10='207' ) And CP09<'C' " & IIf(intPWhere = 國內, " And (SubStr(s1.ST03,1,2)='P2' Or CP14 IS NULL ) ", " And (SubStr(s1.ST03,1,2)='F1' Or CP14 IS NULL ) ") & strSQL2 + StrSQL6
'2011/12/7 modify by sonia 加入 OR CP10='210'
'Modify By Sindy 2021/2/23 +,CP162案源單號
'Modify By Sindy 2021/2/23 改抓案件性質的語法加入TMdebate
'  " and (SubStr(CP10,1,1)='4' OR SubStr(CP10,1,1)='6' OR CP10='202' OR CP10='210' OR CP10='204' OR CP10='205' OR CP10='207' or cp10='712')"
' ==> " and instr('" & TMdebate & ",204,205,207,712',cp10)>0"
'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
strSql = "SELECT cp14,CP12," & MyCPMStr1 & ",1,CP18,DECODE(CP27,NULL,1,'',1,0),DECODE(CP10,'210','202',CP10) CP10,decode(substr(s2.st15,1,1),'F',' ','*'),CP09,CP162 " & _
         " FROM CASEPROGRESS ,TRADEMARK,CASEPROPERTYMAP,STAFF s1,staff s2 " & _
         " WHERE  CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND cp13=s2.st01(+) and CP14=s1.ST01(+) AND CP01=CPM01(+) AND DECODE(CP10,'210','202',CP10)=CPM02(+)" & _
         " and instr('" & TMdebate & ",204,205,207,712',cp10)>0 and not(cp01='FCT' And InStr('" & FCT_NotTMdebate & "', cp10) > 0)" & _
         " And CP09<'C' " & IIf(intPWhere = 國內, " And (SubStr(s1.ST03,1,2)='P2' Or CP14 IS NULL Or CP14='A6015') ", " And (SubStr(s1.ST03,1,2)='F1' Or CP14 IS NULL ) ") & strSQL1 + StrSQL6
strSql = strSql + " union all select cp14,CP12," & MyCPMStr2 & ",1,CP18,DECODE(CP27,NULL,1,'',1,0),DECODE(CP10,'210','202',CP10) CP10,decode(substr(s2.st15,1,1),'F',' ','*'),CP09,CP162 " & _
         " FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF s1,staff s2 " & _
         " WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND cp13=s2.st01(+) and CP14=s1.ST01(+) AND CP01=CPM01(+) AND DECODE(CP10,'210','202',CP10)=CPM02(+)" & _
         " and instr('" & TMdebate & ",204,205,207,712',cp10)>0 and not(cp01='FCT' And InStr('" & FCT_NotTMdebate & "', cp10) > 0)" & _
         " And CP09<'C' " & IIf(intPWhere = 國內, " And (SubStr(s1.ST03,1,2)='P2' Or CP14 IS NULL Or CP14='A6015') ", " And (SubStr(s1.ST03,1,2)='F1' Or CP14 IS NULL ) ") & strSQL2 + StrSQL6

'Add By Cheng 2003/07/01
'自請撤回(306)須依其相關總收文號的案件性質判斷是商爭案
'edit by nickc 2005/05/12
'StrSql = StrSql & " Union All SELECT cp14,CP12,Decode(NVL(CPM03,CP10),'（無）',CPM04,NVL(CPM03,CP10)),1,CP18,DECODE(CP27,NULL,1,'',1,0), '306',decode(tm10,'000','*',' '),CP09 " & _
'         " FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF " & _
'         " WHERE CP09 In (Select CP43 From CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and CP10='306' And CP09<'C' " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' Or CP14 IS NULL ) ", " And (SubStr(ST03,1,2)='F1' Or CP14 IS NULL ) ") & strSQL1 + StrSQL6 & " ) " & _
'         " And CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND '306'=CPM02(+) and (SubStr(CP10,1,1)='4' OR SubStr(CP10,1,1)='6' OR CP10='202' OR CP10='204' OR CP10='205' OR CP10='207' ) "
'StrSql = StrSql & " Union All SELECT cp14,CP12,Decode(NVL(CPM03,CP10),'（無）',CPM04,NVL(CPM03,CP10)),1,CP18,DECODE(CP27,NULL,1,'',1,0), '306',decode(SP09,'000','*',' '),CP09 " & _
'         " FROM CASEPROGRESS, Servicepractice, CASEPROPERTYMAP,STAFF " & _
'         " WHERE CP09 In (Select CP43 From CASEPROGRESS, Servicepractice, CASEPROPERTYMAP,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and CP10='306' And CP09<'C' " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' Or CP14 IS NULL ) ", " And (SubStr(ST03,1,2)='F1' Or CP14 IS NULL ) ") & strSQL2 + StrSQL6 & " ) " & _
'         " And CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND '306'=CPM02(+) and (SubStr(CP10,1,1)='4' OR SubStr(CP10,1,1)='6' OR CP10='202' OR CP10='204' OR CP10='205' OR CP10='207' ) "
'edit by nickc 2007/08/01
'strSQL = strSQL & " Union All SELECT cp14,CP12,Decode(NVL(CPM03,CP10),'（無）',CPM04,NVL(CPM03,CP10)),1,CP18,DECODE(CP27,NULL,1,'',1,0), '306',decode(substr(s2.st15,1,1),'F',' ','*'),CP09 " & _
'         " FROM CASEPROGRESS,CASEPROPERTYMAP,STAFF s1,staff s2 " & _
'         " WHERE CP09 In (Select CP43 From CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and CP10='306' And CP09<'C' " & IIf(intPWhere = 國內, " And (SubStr(s1.ST03,1,2)='P2' Or CP14 IS NULL ) ", " And (SubStr(s1.ST03,1,2)='F1' Or CP14 IS NULL ) ") & strSQL1 + StrSQL6 & " ) " & _
'         " And cp13=s2.st01(+) AND CP14=s1.ST01(+) AND CP01=CPM01(+) AND '306'=CPM02(+) "
'strSQL = strSQL & " Union All SELECT cp14,CP12,Decode(NVL(CPM03,CP10),'（無）',CPM04,NVL(CPM03,CP10)),1,CP18,DECODE(CP27,NULL,1,'',1,0), '306',decode(substr(s2.st15,1,1),'F',' ','*'),CP09 " & _
'         " FROM CASEPROGRESS,  CASEPROPERTYMAP,STAFF s1,staff s2" & _
'         " WHERE CP09 In (Select CP43 From CASEPROGRESS, Servicepractice, CASEPROPERTYMAP,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and CP10='306' And CP09<'C' " & IIf(intPWhere = 國內, " And (SubStr(s1.ST03,1,2)='P2' Or CP14 IS NULL ) ", " And (SubStr(s1.ST03,1,2)='F1' Or CP14 IS NULL ) ") & strSQL2 + StrSQL6 & " ) " & _
'         " And cp13=s2.st01(+) AND CP14=s1.ST01(+) AND CP01=CPM01(+) AND '306'=CPM02(+) "
'Modify By Sindy 2021/2/23 改抓案件性質的語法加入TMdebate
'  " and (SubStr(c1.Cp10,1,1)='4' Or SubStr(c1.Cp10,1,1)='6' OR c1.CP10='202' OR c1.CP10='210' OR c1.CP10='204' OR c1.CP10='205' OR c1.CP10='207')"
' ==> " and instr('" & TMdebate & ",204,205,207,712',c1.Cp10)>0"
'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
strSql = strSql & " Union All SELECT c1.cp14,c1.CP12," & Replace(Replace(UCase(MyCPMStr1), "CP", "C1.CP"), "C1.CPM", "CPM") & ",1,c1.CP18,DECODE(c1.CP27,NULL,1,'',1,0), '306',decode(substr(s2.st15,1,1),'F',' ','*'),c1.CP09,c1.CP162 " & _
         " FROM CASEPROGRESS c1,CASEPROPERTYMAP,STAFF s1,staff s2,(Select CP43,cp14 From CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and CP10='306' And CP09<'C' " & strSQL1 + StrSQL6 & " ) ab " & _
         " WHERE ab.cp43=c1.CP09(+) " & IIf(intPWhere = 國內, " And (SubStr(s1.ST03,1,2)='P2' Or ab.CP14 IS NULL Or ab.CP14='A6015') ", " And (SubStr(s1.ST03,1,2)='F1' Or ab.CP14 IS NULL ) ") & _
         " And c1.cp13=s2.st01(+) AND c1.CP14=s1.ST01(+) AND CP01=CPM01(+) AND '306'=CPM02(+)" & _
         " and instr('" & TMdebate & ",204,205,207,712',c1.Cp10)>0 and not(c1.cp01='FCT' And InStr('" & FCT_NotTMdebate & "', c1.cp10) > 0)"
strSql = strSql & " Union All SELECT c1.cp14,c1.CP12," & Replace(Replace(UCase(MyCPMStr2), "CP", "C1.CP"), "C1.CPM", "CPM") & ",1,c1.CP18,DECODE(c1.CP27,NULL,1,'',1,0), '306',decode(substr(s2.st15,1,1),'F',' ','*'),c1.CP09,c1.CP162 " & _
         " FROM CASEPROGRESS c1,CASEPROPERTYMAP,STAFF s1,staff s2,(Select CP43,cp14 From CASEPROGRESS, Servicepractice, CASEPROPERTYMAP,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and CP10='306' And CP09<'C' " & strSQL2 + StrSQL6 & " ) ab " & _
         " WHERE ab.cp43=c1.CP09(+) " & IIf(intPWhere = 國內, " And (SubStr(s1.ST03,1,2)='P2' Or ab.CP14 IS NULL Or ab.CP14='A6015') ", " And (SubStr(s1.ST03,1,2)='F1' Or ab.CP14 IS NULL ) ") & _
         " And c1.cp13=s2.st01(+) AND c1.CP14=s1.ST01(+) AND c1.CP01=CPM01(+) AND '306'=CPM02(+)" & _
         " and instr('" & TMdebate & ",204,205,207,712',c1.Cp10)>0 and not(c1.cp01='FCT' And InStr('" & FCT_NotTMdebate & "', c1.cp10) > 0)"

'add by nickc 2007/07/24 加入  705 補收款 712 閱卷 但必須是商爭案
'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
strSql = strSql & " Union All SELECT c1.cp14,c1.CP12," & Replace(Replace(UCase(MyCPMStr1), "CP", "C1.CP"), "C1.CPM", "CPM") & ",1,c1.CP18,DECODE(c1.CP27,NULL,1,'',1,0), '705',decode(substr(s2.st15,1,1),'F',' ','*'),c1.CP09,c1.CP162 " & _
         " FROM CASEPROGRESS c1,CASEPROPERTYMAP,STAFF s1,staff s2,(Select CP43,cp14 From CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and CP10='705' And CP09<'C' " & strSQL1 + StrSQL6 & " ) ab " & _
         " WHERE ab.cp43=c1.CP09(+) " & IIf(intPWhere = 國內, " And (SubStr(s1.ST03,1,2)='P2' Or ab.CP14 IS NULL Or ab.CP14='A6015') ", " And (SubStr(s1.ST03,1,2)='F1' Or ab.CP14 IS NULL ) ") & _
         " And c1.cp13=s2.st01(+) AND c1.CP14=s1.ST01(+) AND c1.CP01=CPM01(+) AND '705'=CPM02(+)" & _
         " and instr('" & TMdebate & ",204,205,207,712',c1.Cp10)>0 and not(c1.cp01='FCT' And InStr('" & FCT_NotTMdebate & "', c1.cp10) > 0)"
strSql = strSql & " Union All SELECT c1.cp14,c1.CP12," & Replace(Replace(UCase(MyCPMStr2), "CP", "C1.CP"), "C1.CPM", "CPM") & ",1,c1.CP18,DECODE(c1.CP27,NULL,1,'',1,0), '705',decode(substr(s2.st15,1,1),'F',' ','*'),c1.CP09,c1.CP162 " & _
         " FROM CASEPROGRESS c1,CASEPROPERTYMAP,STAFF s1,staff s2,(Select CP43,cp14 From CASEPROGRESS, Servicepractice, CASEPROPERTYMAP,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and CP10='705' And CP09<'C' " & strSQL2 + StrSQL6 & " ) ab  " & _
         " WHERE ab.cp43=c1.CP09(+)  " & IIf(intPWhere = 國內, " And (SubStr(s1.ST03,1,2)='P2' Or ab.CP14 IS NULL Or ab.CP14='A6015') ", " And (SubStr(s1.ST03,1,2)='F1' Or ab.CP14 IS NULL ) ") & _
         " And c1.cp13=s2.st01(+) AND c1.CP14=s1.ST01(+) AND c1.CP01=CPM01(+) AND '705'=CPM02(+)" & _
         " and instr('" & TMdebate & ",204,205,207,712',c1.Cp10)>0 and not(c1.cp01='FCT' And InStr('" & FCT_NotTMdebate & "', c1.cp10) > 0)"
'edit by nickc 2007/11/06 閱卷 712 一率放商爭
'strSQL = strSQL & " Union All SELECT cp14,CP12," & MyCPMStr1 & ",1,CP18,DECODE(CP27,NULL,1,'',1,0), '712',decode(substr(s2.st15,1,1),'F',' ','*'),CP09 " & _
'         " FROM CASEPROGRESS,CASEPROPERTYMAP,STAFF s1,staff s2 " & _
'         " WHERE CP09 In (Select CP43 From CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and CP10='712' And CP09<'C' " & IIf(intPWhere = 國內, " And (SubStr(s1.ST03,1,2)='P2' Or CP14 IS NULL ) ", " And (SubStr(s1.ST03,1,2)='F1' Or CP14 IS NULL ) ") & strSQL1 + StrSQL6 & " ) " & _
'         " And cp13=s2.st01(+) AND CP14=s1.ST01(+) AND CP01=CPM01(+) AND '712'=CPM02(+) "
'strSQL = strSQL & " Union All SELECT cp14,CP12," & MyCPMStr2 & ",1,CP18,DECODE(CP27,NULL,1,'',1,0), '712',decode(substr(s2.st15,1,1),'F',' ','*'),CP09 " & _
'         " FROM CASEPROGRESS,  CASEPROPERTYMAP,STAFF s1,staff s2" & _
'         " WHERE CP09 In (Select CP43 From CASEPROGRESS, Servicepractice, CASEPROPERTYMAP,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and CP10='712' And CP09<'C' " & IIf(intPWhere = 國內, " And (SubStr(s1.ST03,1,2)='P2' Or CP14 IS NULL ) ", " And (SubStr(s1.ST03,1,2)='F1' Or CP14 IS NULL ) ") & strSQL2 + StrSQL6 & " ) " & _
'         " And cp13=s2.st01(+) AND CP14=s1.ST01(+) AND CP01=CPM01(+) AND '712'=CPM02(+) "
'Modify By Sindy 2015/1/22
If Trim(txt1(0)) = "" Then
   strTemp1 = Split(GetSystemKindByNickT, ",")
Else
   strTemp1 = Split(txt1(0), ",")
End If
System_ID = strTemp1(0)
bolIsChina = False
If txt1(4) = "020" And txt1(5) = "020" Then
   bolIsChina = True
End If
Call ClsPDGetCaseProperty("T", "601", Title_601, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "603", Title_603, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "605", Title_605, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "401", Title_401, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "410", Title_410, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "403", Title_403, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "408", Title_408, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "602", Title_602, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "604", Title_604, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "606", Title_606, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "202", Title_202, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "406", Title_406, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "407", Title_407, bolIsChina, False)
'2015/1/22 END
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
            Select Case Val(CheckStr(.Fields(6)))
            '固定之第一頁
            Case 601, 627 '異議,Add by Sindy 2019/8/15 +部分異議
                blnFix1 = True
                 strTemp(6) = "*"
                 strTemp(7) = "1"
                 strTemp(2) = Title_601 'Modify By Sindy 2015/1/22
'                 'add by nickc 2007/11/02
'                 If txt1(4) = "020" And txt1(5) = "020" Then
'                    strTemp(2) = "異議"
'                 Else
'                    strTemp(2) = "異議"
'                 End If
            Case 603, 629 '評定,Add by Sindy 2019/8/15 +部分評定
                blnFix1 = True
                 strTemp(6) = "*"
                 strTemp(7) = "2"
                 strTemp(2) = Title_603 'Modify By Sindy 2015/1/22
'                 'add by nickc 2007/11/02
'                 If txt1(4) = "020" And txt1(5) = "020" Then
'                    strTemp(2) = "裁定"
'                 Else
'                    strTemp(2) = "評定"
'                 End If
            Case 605, 623 '廢止,Add by Sindy 2019/8/15 +部分廢止
                blnFix1 = True
                 strTemp(6) = "*"
                 strTemp(7) = "3"
                 strTemp(2) = Title_605 'Modify By Sindy 2015/1/22
'                 'add by nickc 2007/11/02
'                 If txt1(4) = "020" And txt1(5) = "020" Then
'                    strTemp(2) = "撤銷"
'                 Else
'                    strTemp(2) = "廢止"
'                 End If
            Case 401 '訴願
                'edit by nickc 2007/07/27 商標處改格式
                If txt1(4) = "020" And txt1(5) = "020" Then
                     blnFix1 = True
                     strTemp(6) = "*"
                     strTemp(7) = "5"
                Else
                    blnFix1 = True
                    strTemp(6) = "*"
                    strTemp(7) = "4"
                End If
                strTemp(2) = Title_401 'Modify By Sindy 2015/1/22
'                 'add by nickc 2007/11/02
'                 If txt1(4) = "020" And txt1(5) = "020" Then
'                    strTemp(2) = "復審"
'                 Else
'                    strTemp(2) = "訴願"
'                 End If
             'add by nickc 2007/07/27 商標處改格式
'             Case 618
'                If txt1(4) = "020" And txt1(5) = "020" Then
'                     blnFix1 = True
'                     strTemp(6) = "*"
'                     strTemp(7) = "4"
'                End If
'                strTemp(2) = Title_618 'Modify By Sindy 2015/1/22
''                 'add by nickc 2007/11/02
''                 If txt1(4) = "020" And txt1(5) = "020" Then
''                    strTemp(2) = "註冊不當撤銷"
''                 End If
            'edit by nickc 2007/07/24 葉大說，再訴願早就沒了，改放  410 行政上訴答辯
            'Case 402 '再訴願
            Case 410  '行政上訴答辯
                'edit by nickc 2007/07/27 商標處改格式
                If txt1(4) = "020" And txt1(5) = "020" Then
                Else
                    blnFix1 = True
                    strTemp(6) = "*"
                    'edit by nickc 2007/11/06 林純貞改位置
                    'strTemp(7) = "5"
                    strTemp(7) = "7"
                End If
                strTemp(2) = Title_410 'Modify By Sindy 2015/1/22
'                 'add by nickc 2007/11/02
'                 If txt1(4) = "020" And txt1(5) = "020" Then
'
'                 Else
'                    strTemp(2) = "行政上訴答辯"
'                 End If
            Case 403 '行政訴訟
                blnFix1 = True
                 strTemp(6) = "*"
                 'edit by nickc 2007/11/06 林純貞改位置
                 'strTemp(7) = "6"
                 'add by nickc 2007/11/02
                 strTemp(2) = Title_403 'Modify By Sindy 2015/1/22
                 If txt1(4) = "020" And txt1(5) = "020" Then
'                    strTemp(2) = "大陸上訴"
                    'add by nickc 2007/11/06 林純貞改位置
                    strTemp(7) = "6"
                 Else
'                    strTemp(2) = "行政訴訟"
                    'add by nickc 2007/11/06 林純貞改位置
                    strTemp(7) = "5"
                 End If
            '將再審之訴(404)移至非固定區, 並以行政訴訟上訴(408)取代之
'            Case 405 '再審之訴
'            Case 409 '行政訴訟上訴
            Case 408 '行政訴訟上訴
                'edit by nickc 2007/07/27 商標處改格式
                If txt1(4) = "020" And txt1(5) = "020" Then
                Else
                    blnFix1 = True
                    strTemp(6) = "*"
                    'edit by nickc 2007/11/06 林純貞改位置
                    'strTemp(7) = "7"
                    strTemp(7) = "6"
                End If
                strTemp(2) = Title_408 'Modify By Sindy 2015/1/22
'                 'add by nickc 2007/11/02
'                 If txt1(4) = "020" And txt1(5) = "020" Then
'
'                 Else
'                    strTemp(2) = "行政訴訟上訴"
'                 End If
            '固定之第二頁
            Case 602, 628 '異議答辯,Add by Sindy 2019/8/15 +部分異議答辯
                blnFix2 = True
                 strTemp(6) = "*"
                 strTemp(7) = "8"
                 strTemp(2) = Title_602 'Modify By Sindy 2015/1/22
'                 'add by nickc 2007/11/02
'                 If txt1(4) = "020" And txt1(5) = "020" Then
'                    strTemp(2) = "異議答辯"
'                 Else
'                    strTemp(2) = "異議答辯"
'                 End If
            Case 604, 630 '評定答辯,Add by Sindy 2019/8/15 +部分評定答辯
                blnFix2 = True
                 strTemp(6) = "*"
                 strTemp(7) = "9"
                 strTemp(2) = Title_604 'Modify By Sindy 2015/1/22
'                 'add by nickc 2007/11/02
'                 If txt1(4) = "020" And txt1(5) = "020" Then
'                    strTemp(2) = "裁定答辯"
'                 Else
'                    strTemp(2) = "評定答辯"
'                 End If
            Case 606, 624 '廢止答辯,Add by Sindy 2019/8/15 +部分廢止答辯
                blnFix2 = True
                 strTemp(6) = "*"
                 strTemp(7) = "10"
                 strTemp(2) = Title_606 'Modify By Sindy 2015/1/22
'                 'add by nickc 2007/11/02
'                 If txt1(4) = "020" And txt1(5) = "020" Then
'                    strTemp(2) = "撤銷答辯"
'                 Else
'                    strTemp(2) = "廢止答辯"
'                 End If
            Case 202 '申請意見書
                'edit by nickc 2007/07/27 商標處改格式
                If txt1(4) = "020" And txt1(5) = "020" Then
                Else
                    blnFix2 = True
                    strTemp(6) = "*"
                    strTemp(7) = "11"
                End If
                strTemp(2) = Title_202 'Modify By Sindy 2015/1/22
'                 'add by nickc 2007/11/02
'                 If txt1(4) = "020" And txt1(5) = "020" Then
'
'                 Else
'                    strTemp(2) = "申請意見書"
'                 End If
'            'add by nickc 2007/07/27 商標處改格式
'            Case 619
'                If txt1(4) = "020" And txt1(5) = "020" Then
'                    blnFix2 = True
'                    strTemp(6) = "*"
'                    strTemp(7) = "11"
'                End If
'                strTemp(2) = Title_619 'Modify By Sindy 2015/1/22
''                 'add by nickc 2007/11/02
''                 If txt1(4) = "020" And txt1(5) = "020" Then
''                    strTemp(2) = "註冊不當撤銷答辯"
''                 End If
            '以參加訴願(406)取代
'            Case 612 '補充理由
'            Case 407 '參加訴願
            Case 406 '參加訴願
                blnFix2 = True
                 strTemp(6) = "*"
                 strTemp(7) = "12"
                 strTemp(2) = Title_406 'Modify By Sindy 2015/1/22
'                 'add by nickc 2007/11/02
'                 If txt1(4) = "020" And txt1(5) = "020" Then
'                    strTemp(2) = "復審答辯"
'                 Else
'                    strTemp(2) = "參加訴願"
'                 End If
            '以參加訴訟(407)取代
'            Case 613 '補充答辯
'            Case 408 '參加訴訟
            Case 407 '參加訴訟
                'edit by nickc 2007/07/27 商標處改格式
                If txt1(4) = "020" And txt1(5) = "020" Then
                Else
                    blnFix2 = True
                    strTemp(6) = "*"
                    strTemp(7) = "13"
                End If
                strTemp(2) = Title_407 'Modify By Sindy 2015/1/22
'                 'add by nickc 2007/11/02
'                 If txt1(4) = "020" And txt1(5) = "020" Then
'
'                 Else
'                    strTemp(2) = "參加訴訟"
'                 End If
            Case Else '其他(屬非固定)
                'Modify By Cheng 2003/12/04
'                blnFix2 = True
                'End
                 strTemp(6) = ""
                 strTemp(7) = "14"
            End Select
            'Add By Cheng 2003/09/03
            '若案件性質為自請撤回(306)
            'edit by nickc 2007/07/24 補收款、閱卷，作法同  自請撤回
            'If "" & .Fields(6).Value = "306" Then
            If "" & .Fields(6).Value = "306" Or "" & .Fields(6).Value = "705" Or "" & .Fields(6).Value = "712" Then
                'Modify By Cheng 2003/12/04
                '排除C類資料
'                strSQLA = "Select * From CaseProgress Where CP43='" & .Fields(8).Value & "' "
                '93.6.4 MODIFY BY SONIA
                'strSQLA = "Select * From CaseProgress Where CP43='" & .Fields(8).Value & "' And CP09<'C' "
                'edit by nickc 2007/07/24 補收款、閱卷，作法同  自請撤回
                'StrSQLa = "Select * From CaseProgress Where CP43='" & .Fields(8).Value & "' AND CP10='306'"
                StrSQLa = "Select * From CaseProgress Where CP43='" & .Fields(8).Value & "' AND CP10='" & "" & .Fields(6).Value & "'"
                '93.6.4 END
                'End
                rsA.CursorLocation = adUseClient
                rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                If rsA.RecordCount > 0 Then
                    '93.6.4 CANCEL BY SONIA 宋若蘭說自請撤回只抓相關總收文號之承辦人及業務區
                    'Modify By Sindy 2012/1/17 仍以原自請撤回智權人員或承辦人來統計
                    strTemp(0) = "" & rsA("CP14").Value
                    strTemp(1) = "" & rsA("CP12").Value
                    '93.6.4 END
                    strTemp(4) = "" & rsA("CP18").Value
                    strTemp(5) = IIf("" & rsA("CP27").Value = "", "1", "0")
                End If
                If rsA.State <> adStateClosed Then rsA.Close
                Set rsA = Nothing
            End If
            
            'Add By Sindy 2021/2/23 有案源單號無點數,並且總收文號為LOS01 P/T案總收文號者
            strTemp(9) = ""
            If "" & .Fields("cp162") <> "" And Val("" & .Fields("cp18")) = 0 Then
               StrSQLa = "SELECT los01,los10,los15,cp09,cp18 FROM lawofficesource,caseprogress" & _
                         " WHERE los15='" & .Fields("cp162").Value & "' AND los01='" & .Fields("cp09").Value & "'" & _
                         " AND los10=cp09(+) AND cp18>0"
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  strTemp(4) = Val("" & rsA("CP18").Value)
                  strTemp(9) = "*"
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            'Add By Sindy 2021/2/23 有案源單號,要扣掉介紹規費cp17的點數
            ElseIf "" & .Fields("cp162") <> "" Then
               If Val("" & .Fields("cp18")) > 0 Then
                  StrSQLa = "SELECT los01,los10,los15,cp09,cp18,cp17 FROM lawofficesource,caseprogress" & _
                            " WHERE los15='" & .Fields("cp162").Value & "' AND los01='" & .Fields("cp09").Value & "'" & _
                            " AND los10=cp09(+) AND cp18>0"
                  rsA.CursorLocation = adUseClient
                  rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsA.RecordCount > 0 Then
                     strTemp(4) = strTemp(4) - Val(Format(Val("" & rsA("CP17").Value) / 1000, "0.0"))
                  End If
                  If rsA.State <> adStateClosed Then rsA.Close
                  Set rsA = Nothing
               End If
            End If
            '2021/2/23 END
            
            'Modify By Sindy 2021/2/26 +R076010:案源點數為*
            strSql = "INSERT INTO r020407 VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "'," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & ",'" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & CheckStr(.Fields(7)) & "','" & strUserNum & "','" & strTemp(9) & "') "
            cnnConnection.Execute strSql
            .MoveNext
            DoEvents
        Loop
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/10/19
        ShowNoData
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
End With
CheckOC
'固定標題:
'edit by nickc 2007/07/27 商標處改格式
If txt1(4) = "020" And txt1(5) = "020" Then
'    cnnConnection.Execute "INSERT INTO r020407_1 VALUES (1,'','','異議','裁定','撤銷','註冊不當撤銷','復審','大陸上訴',' ','','" & strUserNum & "') "
'    cnnConnection.Execute "INSERT INTO r020407_1 VALUES (2,'','','異議答辯','裁定答辯','撤銷答辯','註冊不當撤銷答辯','復審答辯','  ','合計','','" & strUserNum & "') "
    'Modify By Sindy 2015/1/22
                                                                 '601    603    605    401    403
    'cnnConnection.Execute "INSERT INTO r020407_1 VALUES (1,'','','異議','裁定','撤銷','復審','大陸上訴','',' ','','" & strUserNum & "') "
    cnnConnection.Execute "INSERT INTO r020407_1 VALUES (1,'','','" & Title_601 & "','" & Title_603 & "','" & Title_605 & "','" & Title_401 & "','" & Title_403 & "','',' ','','" & strUserNum & "') "
                                                                 '602        604        606        406
    'cnnConnection.Execute "INSERT INTO r020407_1 VALUES (2,'','','異議答辯','裁定答辯','撤銷答辯','復審答辯','','  ','合計','','" & strUserNum & "') "
    cnnConnection.Execute "INSERT INTO r020407_1 VALUES (2,'','','" & Title_602 & "','" & Title_604 & "','" & Title_606 & "','" & Title_406 & "','','  ','合計','','" & strUserNum & "') "
Else
    'cnnConnection.Execute "INSERT INTO r020407_1 VALUES (1,'','','異議','評定','廢止','訴願','再訴願','行政訴訟','再審之訴','','" & strUserNum & "') "
    '不論固定的是否有資料都先加入標題
    ''若有固定第一頁的資料
    'If blnFix1 = True Then
        'edit by nickc 2007/07/24 再訴願 改 行政上訴答辯
        'cnnConnection.Execute "INSERT INTO r020407_1 VALUES (1,'','','異議','評定','廢止','訴願','再訴願','行政訴訟','行政訴訟上訴','','" & strUserNum & "') "
        'edit by nickc 2007/11/06 林純貞改位置
        'cnnConnection.Execute "INSERT INTO r020407_1 VALUES (1,'','','異議','評定','廢止','訴願','行政上訴答辯','行政訴訟','行政訴訟上訴','','" & strUserNum & "') "
                                                                     '601    603    605    401    403        408            410
        'cnnConnection.Execute "INSERT INTO r020407_1 VALUES (1,'','','異議','評定','廢止','訴願','行政訴訟','行政訴訟上訴','行政上訴答辯','','" & strUserNum & "') "
        cnnConnection.Execute "INSERT INTO r020407_1 VALUES (1,'','','" & Title_601 & "','" & Title_603 & "','" & Title_605 & "','" & Title_401 & "','" & Title_403 & "','" & Title_408 & "','" & Title_410 & "','','" & strUserNum & "') "
    'End If
    'cnnConnection.Execute "INSERT INTO r020407_1 VALUES (2,'','','異議答辯','評定答辯','廢止答辯','申請意見書','補充理由','補充答辯','合計','','" & strUserNum & "') "
    ''若有固定第二頁的資料
    'If blnFix2 = True Then
                                                                     '602        604        606        202          406        407
        'cnnConnection.Execute "INSERT INTO r020407_1 VALUES (2,'','','異議答辯','評定答辯','廢止答辯','申請意見書','參加訴願','參加訴訟','合計','','" & strUserNum & "') "
        cnnConnection.Execute "INSERT INTO r020407_1 VALUES (2,'','','" & Title_602 & "','" & Title_604 & "','" & Title_606 & "','" & Title_202 & "','" & Title_406 & "','" & Title_407 & "','合計','','" & strUserNum & "') "
    'End If
End If
'非固定標題:
strSql = "select distinct r076003 from r020407 where id='" & strUserNum & "' and (r076007='' or r076007 is null)  "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        DoEvents
        strTemp3 = "3"
        For i = 1 To 10
            strTemp(i) = ""
        Next i
        Do While .EOF = False
            For i = 3 To 10
                If .EOF = True Then
                    For j = i To 10
                        strTemp(j) = " "
                    Next j
                    Exit For
                Else
                    strTemp(i) = CheckStr(.Fields(0))
                End If
                .MoveNext
            Next i
            strSql = "insert into r020407_1 values (" & Val(strTemp3) & ",'" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','','" & strUserNum & "') "
            cnnConnection.Execute strSql
            strTemp3 = Trim(str(Val(strTemp3) + 1))
            If .EOF = False Then
               'Modify By Cheng 2002/04/12
'                .MoveNext
            End If
            DoEvents
        Loop
    End If
End With
CheckOC

'寫入統計資料:
strSql = "SELECT * FROM r020407_1 WHERE ID='" & strUserNum & "' ORDER BY r077001 "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            'For I = 0 To 10
            '    StrTemp(I) = CheckStr(.Fields(I))
            'Next I
            For i = 3 To 10
                '依案件性質名稱抓資料
                strSql = "SELECT * FROM r020407 WHERE ID='" & strUserNum & "' AND r076003='" & CheckStr(.Fields(i)) & "' "
                CheckOC2
                adoRecordset1.CursorLocation = adUseClient
                adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                    adoRecordset1.MoveFirst
                    Do While adoRecordset1.EOF = False
                        For j = 3 To 33 '26
                            strTemp(j) = ""
                        Next j
                        strTemp(0) = CheckStr(.Fields(0))
                        strTemp(1) = CheckStr(adoRecordset1.Fields(0)) 'R076001:員工編號
                        strTemp(2) = CheckStr(adoRecordset1.Fields(1)) 'R076002:部門
                        Select Case i
                        Case 3
                             strTemp(3) = CheckStr(adoRecordset1.Fields(3))
                             strTemp(4) = CheckStr(adoRecordset1.Fields(4))
                             strTemp(5) = CheckStr(adoRecordset1.Fields(5))
                             strTemp(27) = CheckStr(adoRecordset1.Fields(10)) 'Add By Sindy 2021/2/26
                        Case 4
                             strTemp(6) = CheckStr(adoRecordset1.Fields(3))
                             strTemp(7) = CheckStr(adoRecordset1.Fields(4))
                             strTemp(8) = CheckStr(adoRecordset1.Fields(5))
                             strTemp(28) = CheckStr(adoRecordset1.Fields(10)) 'Add By Sindy 2021/2/26
                        Case 5
                             strTemp(9) = CheckStr(adoRecordset1.Fields(3))
                             strTemp(10) = CheckStr(adoRecordset1.Fields(4))
                             strTemp(11) = CheckStr(adoRecordset1.Fields(5))
                             strTemp(29) = CheckStr(adoRecordset1.Fields(10)) 'Add By Sindy 2021/2/26
                        Case 6
                             strTemp(12) = CheckStr(adoRecordset1.Fields(3))
                             strTemp(13) = CheckStr(adoRecordset1.Fields(4))
                             strTemp(14) = CheckStr(adoRecordset1.Fields(5))
                             strTemp(30) = CheckStr(adoRecordset1.Fields(10)) 'Add By Sindy 2021/2/26
                        Case 7
                             strTemp(15) = CheckStr(adoRecordset1.Fields(3))
                             strTemp(16) = CheckStr(adoRecordset1.Fields(4))
                             strTemp(17) = CheckStr(adoRecordset1.Fields(5))
                             strTemp(31) = CheckStr(adoRecordset1.Fields(10)) 'Add By Sindy 2021/2/26
                        Case 8
                             strTemp(18) = CheckStr(adoRecordset1.Fields(3))
                             strTemp(19) = CheckStr(adoRecordset1.Fields(4))
                             strTemp(20) = CheckStr(adoRecordset1.Fields(5))
                             strTemp(32) = CheckStr(adoRecordset1.Fields(10)) 'Add By Sindy 2021/2/26
                        Case 9
                             strTemp(21) = CheckStr(adoRecordset1.Fields(3))
                             strTemp(22) = CheckStr(adoRecordset1.Fields(4))
                             strTemp(23) = CheckStr(adoRecordset1.Fields(5))
                             strTemp(33) = CheckStr(adoRecordset1.Fields(10)) 'Add By Sindy 2021/2/26
                        Case Else
                        End Select
                        strTemp(24) = CheckStr(adoRecordset1.Fields(3))
                        strTemp(25) = CheckStr(adoRecordset1.Fields(4))
                        strTemp(26) = CheckStr(adoRecordset1.Fields(5))
                        strSql = "INSERT INTO r020407_2 VALUES (" & Val(strTemp(0)) & ",'" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "'," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & "," & Val(strTemp(15)) & "," & Val(strTemp(16)) & "," & Val(strTemp(17)) & "," & Val(strTemp(18)) & "," & Val(strTemp(19)) & "," & Val(strTemp(20)) & "," & Val(strTemp(21)) & "," & Val(strTemp(22)) & "," & Val(strTemp(23)) & "," & Val(strTemp(24)) & "," & Val(strTemp(25)) & "," & Val(strTemp(26)) & ",'" & ChgSQL(CheckStr(adoRecordset1.Fields(8))) & "','" & strUserNum & "'" & _
                                 ",'" & ChgSQL(strTemp(27)) & "','" & ChgSQL(strTemp(28)) & "','" & ChgSQL(strTemp(29)) & "','" & ChgSQL(strTemp(30)) & "','" & ChgSQL(strTemp(31)) & "','" & ChgSQL(strTemp(32)) & "','" & ChgSQL(strTemp(33)) & "') "
                        cnnConnection.Execute strSql
                        adoRecordset1.MoveNext
                    Loop
                End If
            Next i
            .MoveNext
        Loop
    End If
End With
CheckOC
'將該承辦人的業務區補齊
strSql = "select max(r077001) from r020407_1 where id='" & strUserNum & "' "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    s = Val(CheckStr(adoRecordset.Fields(0)))
End If
CheckOC
strSql = "select distinct r076001,r076002,r076009 from r020407 where id='" & strUserNum & "' "
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 1 To s
               'Add By Sindy 2024/2/19 檢查資料若已存在不可再新增
               strSql = "SELECT * FROM r020407_2" & _
                        " WHERE R078001=" & i & _
                        " and R078002='" & ChgSQL(CheckStr(.Fields(0))) & "'" & _
                        " and R078003='" & ChgSQL(CheckStr(.Fields(1))) & "'" & _
                        " and R078028='" & ChgSQL(CheckStr(.Fields(2))) & "'" & _
                        " and Id='" & strUserNum & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 0 Then
               '2024/2/19 END
                  strSql = "insert into r020407_2 values (" & i & ",'" & ChgSQL(CheckStr(.Fields(0))) & "','" & ChgSQL(CheckStr(.Fields(1))) & "',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,'" & ChgSQL(CheckStr(.Fields(2))) & "','" & strUserNum & "',null,null,null,null,null,null,null) "
                  cnnConnection.Execute strSql
               End If
            Next i
            .MoveNext
        Loop
    End If
End With
CheckOC
''Add By Cheng 2003/12/04
'If blnFix1 = False Then
'    strSQL = "Delete From R020407_2 Where R078001='1' And ID='" & strUserNum & "' "
'    cnnConnection.Execute strSQL
'End If
'If blnFix2 = False Then
'    strSQL = "Delete From R020407_2 Where R078001='2' And ID='" & strUserNum & "' "
'    cnnConnection.Execute strSQL
'End If
''End
'重整
'Modify By Cheng 2002/04/12
'strSQL = "select r078001,r078002,r078003,sum(r078004),sum(r078005),sum(r078006),sum(r078007),sum(r078008),sum(r078009),sum(r078010),sum(r078011),sum(r078012),sum(r078013),sum(r078014),sum(r078015),sum(r078016),sum(r078017),sum(r078018),sum(r078019),sum(r078020),sum(r078021),sum(r078022),sum(r078023),sum(r078024),sum(r078025),sum(r078026),sum(r078027),'' from r020407_2 where id='" & strUserNum & "' group by r078001,r078002,r078003 order by r078001,r078002,r078003 "
strSql = "select r078001,r078002,r078003,sum(r078004),sum(r078005),sum(r078006),sum(r078007),sum(r078008),sum(r078009),sum(r078010),sum(r078011),sum(r078012),sum(r078013),sum(r078014),sum(r078015),sum(r078016),sum(r078017),sum(r078018),sum(r078019),sum(r078020),sum(r078021),sum(r078022),sum(r078023),sum(r078024),sum(r078025),sum(r078026),sum(r078027),r078028" & _
         ",r078029,r078030,r078031,r078032,r078033,r078034,r078035" & _
         " from r020407_2 where id='" & strUserNum & "' group by r078001,r078002,r078003,r078028,r078029,r078030,r078031,r078032,r078033,r078034,r078035 order by r078001,r078002,r078003 "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        cnnConnection.Execute "DELETE FROM r020407_2 WHERE ID='" & strUserNum & "' "
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 34 '26
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strSql = "INSERT INTO r020407_2 VALUES (" & Val(strTemp(0)) & ",'" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "'," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & "," & Val(strTemp(15)) & "," & Val(strTemp(16)) & "," & Val(strTemp(17)) & "," & Val(strTemp(18)) & "," & Val(strTemp(19)) & "," & Val(strTemp(20)) & "," & Val(strTemp(21)) & "," & Val(strTemp(22)) & "," & Val(strTemp(23)) & "," & Val(strTemp(24)) & "," & Val(strTemp(25)) & "," & Val(strTemp(26)) & ",'" & ChgSQL(CheckStr(.Fields(27))) & "','" & strUserNum & "'" & _
                     ",'" & ChgSQL(strTemp(28)) & "','" & ChgSQL(strTemp(29)) & "','" & ChgSQL(strTemp(30)) & "','" & ChgSQL(strTemp(31)) & "','" & ChgSQL(strTemp(32)) & "','" & ChgSQL(strTemp(33)) & "','" & ChgSQL(strTemp(34)) & "') "
            cnnConnection.Execute strSql
            .MoveNext
        Loop
    End If
End With
CheckOC
'update 固定的合計
strSql = "select r078002,r078003,sum(r078025),sum(r078026),sum(r078027),r078028 from r020407_2 where id='" & strUserNum & "' and r078001<=2 group by r078002,r078003,r078028"
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   adoRecordset.MoveFirst
   Do While adoRecordset.EOF = False
      strSql = "UPDATE r020407_2 SET r078022=" & Val(CheckStr(adoRecordset.Fields(2))) & ",r078023=" & Val(CheckStr(adoRecordset.Fields(3))) & ",r078024=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE r078002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND r078003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' AND ID='" & strUserNum & "' AND r078001=2 and " & ChangeNewSQLStr(adoRecordset.Fields(5), "r078028")
      cnnConnection.Execute strSql
      adoRecordset.MoveNext
   Loop
End If
CheckOC

'update 非固定的合計
s = 8
strSql = "select r077001,r077004,r077005,r077006,r077007,r077008,r077009,r077010 from r020407_1 where ID='" & strUserNum & "' AND r077001 in (select max(r077001) from r020407_1 where id='" & strUserNum & "') "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   For i = 1 To 7
       If Len(CheckStr(adoRecordset.Fields(i))) = 0 Then
           s = i
           Exit For
       End If
   Next i
   Dim strMax As String
   strMax = CheckStr(adoRecordset.Fields(0))
   If s <> 8 Then
      Select Case s
      Case 1
            strSql = "update r020407_1 set r077004='合計',r077005='總計' where id='" & strUserNum & "' and r077001='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' "
            cnnConnection.Execute strSql
            '統計合計
            strSql = "select r078002,r078003,sum(r078025),sum(r078026),sum(r078027),r078028 from r020407_2 where id='" & strUserNum & "' and r078001>2 group by r078002,r078003,r078028"
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
               adoRecordset.MoveFirst
               Do While adoRecordset.EOF = False
                  strSql = "UPDATE r020407_2 SET r078004=" & Val(CheckStr(adoRecordset.Fields(2))) & ",r078004=" & Val(CheckStr(adoRecordset.Fields(3))) & ",r078006=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r078002 is null ", " r078002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r078003 is null ", " AND r078003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND r078001=" & strMax & " and " & ChangeNewSQLStr(adoRecordset.Fields(5), "r078028")
                  cnnConnection.Execute strSql
                  adoRecordset.MoveNext
               Loop
            End If
            CheckOC
            '統計總計
            strSql = "select r078002,r078003,sum(r078025),sum(r078026),sum(r078027),r078028 from r020407_2 where id='" & strUserNum & "' group by r078002,r078003,r078028"
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
               adoRecordset.MoveFirst
               Do While adoRecordset.EOF = False
                  strSql = "UPDATE r020407_2 SET r078007=" & Val(CheckStr(adoRecordset.Fields(2))) & ",r078008=" & Val(CheckStr(adoRecordset.Fields(3))) & ",r078009=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r078002 is null ", " r078002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r078003 is null ", " AND r078003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND r078001=" & strMax & "  and " & ChangeNewSQLStr(adoRecordset.Fields(5), "r078028")
                  cnnConnection.Execute strSql
                  adoRecordset.MoveNext
               Loop
            End If
            CheckOC
      Case 2
            strSql = "update r020407_1 set r077005='合計',r077006='總計' where id='" & strUserNum & "' and r077001='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' "
            cnnConnection.Execute strSql
            '統計合計
            strSql = "select r078002,r078003,sum(r078025),sum(r078026),sum(r078027),r078028 from r020407_2 where id='" & strUserNum & "' and r078001>2 group by r078002,r078003,r078028"
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
               adoRecordset.MoveFirst
               Do While adoRecordset.EOF = False
                  strSql = "UPDATE r020407_2 SET r078007=" & Val(CheckStr(adoRecordset.Fields(2))) & ",r078008=" & Val(CheckStr(adoRecordset.Fields(3))) & ",r078009=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r078002 is null ", " r078002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r078003 is null ", " AND r078003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND r078001=" & strMax & "  and " & ChangeNewSQLStr(adoRecordset.Fields(5), "r078028")
                  cnnConnection.Execute strSql
                  adoRecordset.MoveNext
               Loop
            End If
            CheckOC
            '統計總計
            strSql = "select r078002,r078003,sum(r078025),sum(r078026),sum(r078027),r078028 from r020407_2 where id='" & strUserNum & "' group by r078002,r078003,r078028"
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
               adoRecordset.MoveFirst
               Do While adoRecordset.EOF = False
                  strSql = "UPDATE r020407_2 SET r078010=" & Val(CheckStr(adoRecordset.Fields(2))) & ",r078011=" & Val(CheckStr(adoRecordset.Fields(3))) & ",r078012=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r078002 is null ", " r078002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r078003 is null ", " AND r078003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND r078001=" & strMax & "  and " & ChangeNewSQLStr(adoRecordset.Fields(5), "r078028")
                  cnnConnection.Execute strSql
                  adoRecordset.MoveNext
               Loop
            End If
            CheckOC
      Case 3
            strSql = "update r020407_1 set r077006='合計',r077007='總計' where id='" & strUserNum & "' and r077001='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' "
            cnnConnection.Execute strSql
            '統計合計
            strSql = "select r078002,r078003,sum(r078025),sum(r078026),sum(r078027),r078028 from r020407_2 where id='" & strUserNum & "' and r078001>2 group by r078002,r078003,r078028"
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
               adoRecordset.MoveFirst
               Do While adoRecordset.EOF = False
                  strSql = "UPDATE r020407_2 SET r078010=" & Val(CheckStr(adoRecordset.Fields(2))) & ",r078011=" & Val(CheckStr(adoRecordset.Fields(3))) & ",r078012=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r078002 is null ", " r078002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r078003 is null ", " AND r078003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND r078001=" & strMax & "  and " & ChangeNewSQLStr(adoRecordset.Fields(5), "r078028")
                  cnnConnection.Execute strSql
                  adoRecordset.MoveNext
               Loop
            End If
            CheckOC
            '統計總計
            strSql = "select r078002,r078003,sum(r078025),sum(r078026),sum(r078027),r078028 from r020407_2 where id='" & strUserNum & "' group by r078002,r078003,r078028"
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
               adoRecordset.MoveFirst
               Do While adoRecordset.EOF = False
                  strSql = "UPDATE r020407_2 SET r078013=" & Val(CheckStr(adoRecordset.Fields(2))) & ",r078014=" & Val(CheckStr(adoRecordset.Fields(3))) & ",r078015=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r078002 is null ", " r078002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r078003 is null ", " AND r078003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND r078001=" & strMax & "  and " & ChangeNewSQLStr(adoRecordset.Fields(5), "r078028")
                  cnnConnection.Execute strSql
                  adoRecordset.MoveNext
               Loop
            End If
            CheckOC
      Case 4
            strSql = "update r020407_1 set r077007='合計',r077008='總計' where id='" & strUserNum & "' and r077001='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' "
            cnnConnection.Execute strSql
            '統計合計
            strSql = "select r078002,r078003,sum(r078025),sum(r078026),sum(r078027),r078028 from r020407_2 where id='" & strUserNum & "' and r078001>2 group by r078002,r078003,r078028"
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
               adoRecordset.MoveFirst
               Do While adoRecordset.EOF = False
                  strSql = "UPDATE r020407_2 SET r078013=" & Val(CheckStr(adoRecordset.Fields(2))) & ",r078014=" & Val(CheckStr(adoRecordset.Fields(3))) & ",r078015=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r078002 is null ", " r078002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r078003 is null ", " AND r078003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND r078001=" & strMax & "  and " & ChangeNewSQLStr(adoRecordset.Fields(5), "r078028")
                  cnnConnection.Execute strSql
                  adoRecordset.MoveNext
               Loop
            End If
            CheckOC
            '統計總計
            strSql = "select r078002,r078003,sum(r078025),sum(r078026),sum(r078027),r078028 from r020407_2 where id='" & strUserNum & "' group by r078002,r078003,r078028"
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
               adoRecordset.MoveFirst
               Do While adoRecordset.EOF = False
                  strSql = "UPDATE r020407_2 SET r078016=" & Val(CheckStr(adoRecordset.Fields(2))) & ",r078017=" & Val(CheckStr(adoRecordset.Fields(3))) & ",r078018=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r078002 is null ", " r078002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r078003 is null ", " AND r078003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND r078001=" & strMax & "  and " & ChangeNewSQLStr(adoRecordset.Fields(5), "r078028")
                  cnnConnection.Execute strSql
                  adoRecordset.MoveNext
               Loop
            End If
            CheckOC
      Case 5
            strSql = "update r020407_1 set r077008='合計',r077009='總計' where id='" & strUserNum & "' and r077001='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' "
            cnnConnection.Execute strSql
            '統計合計
            strSql = "select r078002,r078003,sum(r078025),sum(r078026),sum(r078027),r078028 from r020407_2 where id='" & strUserNum & "' and r078001>2 group by r078002,r078003,r078028"
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
               adoRecordset.MoveFirst
               Do While adoRecordset.EOF = False
                  strSql = "UPDATE r020407_2 SET r078016=" & Val(CheckStr(adoRecordset.Fields(2))) & ",r078017=" & Val(CheckStr(adoRecordset.Fields(3))) & ",r078018=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r078002 is null ", " r078002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r078003 is null ", " AND r078003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND r078001=" & strMax & "  and " & ChangeNewSQLStr(adoRecordset.Fields(5), "r078028")
                  cnnConnection.Execute strSql
                  adoRecordset.MoveNext
               Loop
            End If
            CheckOC
            '統計總計
            strSql = "select r078002,r078003,sum(r078025),sum(r078026),sum(r078027),r078028 from r020407_2 where id='" & strUserNum & "' group by r078002,r078003,r078028"
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
               adoRecordset.MoveFirst
               Do While adoRecordset.EOF = False
                  strSql = "UPDATE r020407_2 SET r078019=" & Val(CheckStr(adoRecordset.Fields(2))) & ",r078020=" & Val(CheckStr(adoRecordset.Fields(3))) & ",r078021=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r078002 is null ", " r078002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r078003 is null ", " AND r078003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND r078001=" & strMax & "  and " & ChangeNewSQLStr(adoRecordset.Fields(5), "r078028")
                  cnnConnection.Execute strSql, intI
'                  If adoRecordset.Fields(0) = "A4022" Then
'                     MsgBox "ttt"
'                  End If
                  If intI = 0 Then
                     MsgBox "ttt"
                  End If
                  adoRecordset.MoveNext
               Loop
            End If
            CheckOC
      Case 6
            strSql = "update r020407_1 set r077009='合計',r077010='總計' where id='" & strUserNum & "' and r077001='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' "
            cnnConnection.Execute strSql
            '統計合計
            strSql = "select r078002,r078003,sum(r078025),sum(r078026),sum(r078027),r078028 from r020407_2 where id='" & strUserNum & "' and r078001>2 group by r078002,r078003,r078028"
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
               adoRecordset.MoveFirst
               Do While adoRecordset.EOF = False
                  strSql = "UPDATE r020407_2 SET r078019=" & Val(CheckStr(adoRecordset.Fields(2))) & ",r078020=" & Val(CheckStr(adoRecordset.Fields(3))) & ",r078021=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r078002 is null ", " r078002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r078003 is null ", " AND r078003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND r078001=" & strMax & "  and " & ChangeNewSQLStr(adoRecordset.Fields(5), "r078028")
                  cnnConnection.Execute strSql
                  adoRecordset.MoveNext
               Loop
            End If
            CheckOC
            '統計總計
            strSql = "select r078002,r078003,sum(r078025),sum(r078026),sum(r078027),r078028 from r020407_2 where id='" & strUserNum & "' group by r078002,r078003,r078028"
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
               adoRecordset.MoveFirst
               Do While adoRecordset.EOF = False
                  strSql = "UPDATE r020407_2 SET r078022=" & Val(CheckStr(adoRecordset.Fields(2))) & ",r078023=" & Val(CheckStr(adoRecordset.Fields(3))) & ",r078024=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r078002 is null ", " r078002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r078003 is null ", " AND r078003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND r078001=" & strMax & "  and " & ChangeNewSQLStr(adoRecordset.Fields(5), "r078028")
                  cnnConnection.Execute strSql
                  adoRecordset.MoveNext
               Loop
            End If
            CheckOC
      Case 7
            strSql = "update r020407_1 set r077010='合計' where id='" & strUserNum & "' and r077001=" & strMax & " "
            cnnConnection.Execute strSql
            strSql = "insert into r020407_1 values (" & Val(strMax) + 1 & ",'','','總計','','','','','','','','" & strUserNum & "' )"
            cnnConnection.Execute strSql
            '統計合計
            strSql = "select r078002,r078003,sum(r078025),sum(r078026),sum(r078027),r078028 from r020407_2 where id='" & strUserNum & "' and r078001>2 group by r078002,r078003,r078028"
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
               adoRecordset.MoveFirst
               Do While adoRecordset.EOF = False
                  strSql = "UPDATE r020407_2 SET r078022=" & Val(CheckStr(adoRecordset.Fields(2))) & ",r078023=" & Val(CheckStr(adoRecordset.Fields(3))) & ",r078024=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r078002 is null ", " r078002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r078003 is null ", " AND r078003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND r078001=" & strMax & "  and " & ChangeNewSQLStr(adoRecordset.Fields(5), "r078028")
                  cnnConnection.Execute strSql
                  adoRecordset.MoveNext
               Loop
            End If
            CheckOC
            '統計總計
            strSql = "select r078002,r078003,sum(r078025),sum(r078026),sum(r078027),r078028 from r020407_2 where id='" & strUserNum & "' group by r078002,r078003,r078028"
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
               adoRecordset.MoveFirst
               Do While adoRecordset.EOF = False
                  strSql = "insert into r020407_2 (r078001,r078002,r078003,r078004,r078005,r078006,id) values (" & Val(strMax) + 1 & ",'" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "','" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "'," & Val(CheckStr(adoRecordset.Fields(2))) & "," & Val(CheckStr(adoRecordset.Fields(3))) & "," & Val(CheckStr(adoRecordset.Fields(4))) & ",'" & strUserNum & "') "
                  cnnConnection.Execute strSql
                  adoRecordset.MoveNext
               Loop
            End If
            CheckOC
      Case Else
      End Select
      
   End If
End If
CheckOC
If s = 8 Then
   strSql = "insert into r020407_1 (select max(r077001)+1,'','','合計','總計','','','','','','','" & strUserNum & "' from r020407_1 where id='" & strUserNum & "')"
   cnnConnection.Execute strSql
   '統計合計
   strSql = "select r078002,r078003,sum(r078025),sum(r078026),sum(r078027),r078028 from r020407_2 where id='" & strUserNum & "' and r078001>2 group by r078002,r078003,r078028"
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
        adoRecordset.MoveFirst
        Do While adoRecordset.EOF = False
            strSql = "insert into r020407_2 (r078001,r078002,r078003,r078004,r078005,r078006,id)  (select max(r077001) ,'" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "','" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "'," & Val(CheckStr(adoRecordset.Fields(2))) & "," & Val(CheckStr(adoRecordset.Fields(3))) & "," & Val(CheckStr(adoRecordset.Fields(4))) & ",'" & strUserNum & "' from r020407_1 where id='" & strUserNum & "') "
            cnnConnection.Execute strSql
            adoRecordset.MoveNext
        Loop
    End If
   CheckOC
   '統計總計
   'edit by nick 2005/02/04 總計應該不能分 固定非固定
   'StrSql = "select r078002,r078003,sum(r078025),sum(r078026),sum(r078027),r078028 from r020407_2 where id='" & strUserNum & "' group by r078002,r078003,r078028"
   strSql = "select r078002,r078003,sum(r078025),sum(r078026),sum(r078027) from r020407_2 where id='" & strUserNum & "' group by r078002,r078003"
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
      adoRecordset.MoveFirst
      Do While adoRecordset.EOF = False
         'StrSQL = "update r020407_2 set (r078001,r078002,r078003,r078004,r078005,r078006,id) (select " & max(r077001) & ",'" & chgsql(CheckStr(adoRecordset.Fields(0))) & "','" & chgsql(CheckStr(adoRecordset.Fields(1))) & "'," & Val(CheckStr(adoRecordset.Fields(2))) & "," & Val(CheckStr(adoRecordset.Fields(3))) & "," & Val(CheckStr(adoRecordset.Fields(4))) & ",'" & strUserNum & "') "
            'Modify By Cheng 2004/04/07
'         strSQL = "UPDATE r020407_2 SET r078007=" & Val(CheckStr(adoRecordset.Fields(2))) & ",r078008=" & Val(CheckStr(adoRecordset.Fields(3))) & ",r078009=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r078002 is null ", " r078002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r078003 is null ", " AND r078003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND r078001 in (select  max(r077001)  from r020407_1 where id='" & strUserNum & "')  and " & ChangeNewSQLStr(adoRecordset.Fields(5), "r078028")
         'edit by nick 2005/02/04 總計應該不能分 固定非固定
         'StrSql = "UPDATE r020407_2 SET r078007=" & Val(CheckStr(adoRecordset.Fields(2))) & ",r078008=" & Val(CheckStr(adoRecordset.Fields(3))) & ",r078009=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r078002 is null ", " r078002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r078003 is null ", " AND r078003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND r078001 in (select  max(r077001)  from r020407_1 where id='" & strUserNum & "') And R078001>2 and " & ChangeNewSQLStr(adoRecordset.Fields(5), "r078028")
         strSql = "UPDATE r020407_2 SET r078007=" & Val(CheckStr(adoRecordset.Fields(2))) & ",r078008=" & Val(CheckStr(adoRecordset.Fields(3))) & ",r078009=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r078002 is null ", " r078002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r078003 is null ", " AND r078003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND r078001 in (select  max(r077001)  from r020407_1 where id='" & strUserNum & "') And R078001>2 "
            'End
         cnnConnection.Execute strSql, ii
         If ii <= 0 Then
            strSql = "insert into r020407_2 (r078001,r078002,r078003,r078004,r078005,r078006,r078007,r078008,r078009,id)  (select Max(3),'" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "','" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "',0,0,0," & Val(CheckStr(adoRecordset.Fields(2))) & "," & Val(CheckStr(adoRecordset.Fields(3))) & "," & Val(CheckStr(adoRecordset.Fields(4))) & ",'" & strUserNum & "' from r020407_1 where id='" & strUserNum & "') "
            cnnConnection.Execute strSql
         End If
         adoRecordset.MoveNext
      Loop
   End If
   CheckOC
End If
PrintData
Screen.MousePointer = vbDefault
End Sub

Sub PrintData()
Select Case Val(txt1(1))
Case 1 '收文
     PrintData1
Case 2 '發文
     PrintData2
Case Else
End Select
End Sub

'選擇發文
Sub PrintData2()
'Add By Cheng 2002/02/06
Dim strSQL_1 As String
Dim strSQL_2 As String
Dim strSQL_3 As String
Dim rstmp_1 As New ADODB.Recordset
Dim rstmp_2 As New ADODB.Recordset
Dim rstmp_3 As New ADODB.Recordset
Dim dblTotal As Double '總計

BolEndThisPage = False

'Add By Cheng 2002/02/06
'刪除固定區的合計為 0 的資料, 及非固定區的總計為 0 的資料
'搜尋固定區的第一頁
'Modify By Sindy 2021/6/22 , ==> +
strSQL_1 = "SELECT r078001,nvl(st02,r078002),NVL(A0902,A0903),decode(r078004,null,0,r078004)+decode(r078005,null,0,r078005)+decode(r078006,null,0,r078006)" & _
            "+decode(r078007,null,0,r078007)+decode(r078008,null,0,r078008)+decode(r078009,null,0,r078009)" & _
            "+decode(r078010,null,0,r078010)+decode(r078011,null,0,r078011)+decode(r078012,null,0,r078012)" & _
            "+decode(r078013,null,0,r078013)+decode(r078014,null,0,r078014)+decode(r078015,null,0,r078015)" & _
            "+decode(r078016,NULL,0,r078016)+DECODE(r078017,NULL,0,r078017)+DECODE(r078018,NULL,0,r078018)" & _
            "+DECODE(R078019,NULL,0,R078019)+DECODE(r078020,NULL,0,r078020)+DECODE(r078021,NULL,0,r078021)" & _
            "+DECODE(r078022,NULL,0,r078022)+DECODE(r078023,NULL,0,r078023)+DECODE(r078024,NULL,0,r078024)" & _
            "+DECODE(r078025,NULL,0,r078025)+DECODE(r078026,NULL,0,r078026)+DECODE(r078027,NULL,0,r078027) As Total,'',r078002,r078003 " & _
            " FROM r020407_2,ACC090,staff " & _
            " WHERE r078002=st01(+) and ID='" & strUserNum & "' AND r078003=A0901(+) And R078001='1' ORDER BY r078001,r078002,r078003 "
If rstmp_1.State <> adStateClosed Then rstmp_1.Close
rstmp_1.CursorLocation = adUseClient
rstmp_1.Open strSQL_1, cnnConnection, adOpenStatic, adLockReadOnly
If rstmp_1.RecordCount > 0 Then
   rstmp_1.MoveFirst
   While Not rstmp_1.EOF
      '搜尋固定區的第二頁
      'Modify By Sindy 2021/6/22 , ==> +
      strSQL_2 = "SELECT r078001,nvl(st02,r078002),NVL(A0902,A0903),decode(r078004,null,0,r078004)+decode(r078005,null,0,r078005)+decode(r078006,null,0,r078006)" & _
                  "+decode(r078007,null,0,r078007)+decode(r078008,null,0,r078008)+decode(r078009,null,0,r078009)" & _
                  "+decode(r078010,null,0,r078010)+decode(r078011,null,0,r078011)+decode(r078012,null,0,r078012)" & _
                  "+decode(r078013,null,0,r078013)+decode(r078014,null,0,r078014)+decode(r078015,null,0,r078015)" & _
                  "+decode(r078016,NULL,0,r078016)+DECODE(r078017,NULL,0,r078017)+DECODE(r078018,NULL,0,r078018)" & _
                  "+DECODE(R078019,NULL,0,R078019)+DECODE(r078020,NULL,0,r078020)+DECODE(r078021,NULL,0,r078021)" & _
                  "+DECODE(r078022,NULL,0,r078022)+DECODE(r078023,NULL,0,r078023)+DECODE(r078024,NULL,0,r078024)" & _
                  "+DECODE(r078025,NULL,0,r078025)+DECODE(r078026,NULL,0,r078026)+DECODE(r078027,NULL,0,r078027) As Total,'',r078002,r078003 " & _
                  " FROM r020407_2,ACC090,staff " & _
                  " WHERE r078002=st01(+) and ID='" & strUserNum & "' AND r078003=A0901(+) And R078001='2' And R078002='" & rstmp_1("R078002").Value & "' And R078003='" & rstmp_1("R078003").Value & "' ORDER BY r078001,r078002,r078003 "
      If rstmp_2.State <> adStateClosed Then rstmp_2.Close
      rstmp_2.CursorLocation = adUseClient
      rstmp_2.Open strSQL_2, cnnConnection, adOpenStatic, adLockReadOnly
      If rstmp_2.RecordCount > 0 Then
         rstmp_2.MoveFirst
         '若固定區合計為0
         If rstmp_1("Total") + rstmp_2("Total") = 0 Then
            'Modify By Sindy 2015/7/6 不要刪除固定區資料
            'cnnConnection.Execute "Delete From R020407_2 Where ID ='" & strUserNum & "' And R078002='" & rstmp_1("R078002").Value & "' And R078003='" & rstmp_1("R078003").Value & "' And (R078001='1' Or R078001='2') "
            '2015/7/6 END
            '搜尋固定區的第三頁以上
            'Modify By Sindy 2021/6/22 , ==> +
            strSQL_3 = "SELECT r078001,nvl(st02,r078002),NVL(A0902,A0903),decode(r078004,null,0,r078004)+decode(r078005,null,0,r078005)+decode(r078006,null,0,r078006)" & _
                        "+decode(r078007,null,0,r078007)+decode(r078008,null,0,r078008)+decode(r078009,null,0,r078009)" & _
                        "+decode(r078010,null,0,r078010)+decode(r078011,null,0,r078011)+decode(r078012,null,0,r078012)" & _
                        "+decode(r078013,null,0,r078013)+decode(r078014,null,0,r078014)+decode(r078015,null,0,r078015)" & _
                        "+decode(r078016,NULL,0,r078016)+DECODE(r078017,NULL,0,r078017)+DECODE(r078018,NULL,0,r078018)" & _
                        "+DECODE(R078019,NULL,0,R078019)+DECODE(r078020,NULL,0,r078020)+DECODE(r078021,NULL,0,r078021)" & _
                        "+DECODE(r078022,NULL,0,r078022)+DECODE(r078023,NULL,0,r078023)+DECODE(r078024,NULL,0,r078024)" & _
                        "+DECODE(r078025,NULL,0,r078025)+DECODE(r078026,NULL,0,r078026)+DECODE(r078027,NULL,0,r078027) As Total,'',r078002,r078003 " & _
                        " FROM r020407_2,ACC090,staff " & _
                        " WHERE r078002=st01(+) and ID='" & strUserNum & "' AND r078003=A0901(+) And (R078001<>'1' And R078001<>'2') And R078002='" & rstmp_1("R078002").Value & "' And R078003='" & rstmp_1("R078003").Value & "' ORDER BY r078001,r078002,r078003 "
            If rstmp_3.State <> adStateClosed Then rstmp_3.Close
            rstmp_3.CursorLocation = adUseClient
            rstmp_3.Open strSQL_3, cnnConnection, adOpenStatic, adLockReadOnly
            dblTotal = 0
            If rstmp_3.RecordCount > 0 Then
               rstmp_3.MoveFirst
               While Not rstmp_3.EOF
                  dblTotal = dblTotal + rstmp_3("Total")
                  rstmp_3.MoveNext
               Wend
               If dblTotal = 0 Then
                  cnnConnection.Execute "Delete From R020407_2 Where ID ='" & strUserNum & "' And R078002='" & rstmp_1("R078002").Value & "' And R078003='" & rstmp_1("R078003").Value & "' And (R078001<>'1' And R078001<>'2') "
               End If
            End If
         End If
      End If
      rstmp_1.MoveNext
   Wend
Else
   ShowNoData
   If rstmp_1.State <> adStateClosed Then rstmp_1.Close
   Set rstmp_1 = Nothing
   Screen.MousePointer = vbDefault
   Exit Sub
End If
If rstmp_1.State <> adStateClosed Then rstmp_1.Close
Set rstmp_1 = Nothing
If rstmp_2.State <> adStateClosed Then rstmp_2.Close
Set rstmp_2 = Nothing
If rstmp_3.State <> adStateClosed Then rstmp_3.Close
Set rstmp_3 = Nothing

'Add By Sindy 2016/9/8
cnnConnection.Execute "Update R020407_2 set R078028=decode(substr(R078003,1,1),'F',' ','*') Where ID ='" & strUserNum & "'"
'2016/9/8 END
'Add By Sindy 2021/3/2
cnnConnection.Execute "UPDATE r020407_2 SET r078029='*' where r078001||r078002||r078003 in(SELECT r078001||r078002||r078003 FROM r020407_2 WHERE ID='" & strUserNum & "' AND r078029='*' GROUP BY r078001,r078002,r078003)"
cnnConnection.Execute "UPDATE r020407_2 SET r078030='*' where r078001||r078002||r078003 in(SELECT r078001||r078002||r078003 FROM r020407_2 WHERE ID='" & strUserNum & "' AND r078030='*' GROUP BY r078001,r078002,r078003)"
cnnConnection.Execute "UPDATE r020407_2 SET r078031='*' where r078001||r078002||r078003 in(SELECT r078001||r078002||r078003 FROM r020407_2 WHERE ID='" & strUserNum & "' AND r078031='*' GROUP BY r078001,r078002,r078003)"
cnnConnection.Execute "UPDATE r020407_2 SET r078032='*' where r078001||r078002||r078003 in(SELECT r078001||r078002||r078003 FROM r020407_2 WHERE ID='" & strUserNum & "' AND r078032='*' GROUP BY r078001,r078002,r078003)"
cnnConnection.Execute "UPDATE r020407_2 SET r078033='*' where r078001||r078002||r078003 in(SELECT r078001||r078002||r078003 FROM r020407_2 WHERE ID='" & strUserNum & "' AND r078033='*' GROUP BY r078001,r078002,r078003)"
cnnConnection.Execute "UPDATE r020407_2 SET r078034='*' where r078001||r078002||r078003 in(SELECT r078001||r078002||r078003 FROM r020407_2 WHERE ID='" & strUserNum & "' AND r078034='*' GROUP BY r078001,r078002,r078003)"
cnnConnection.Execute "UPDATE r020407_2 SET r078035='*' where r078001||r078002||r078003 in(SELECT r078001||r078002||r078003 FROM r020407_2 WHERE ID='" & strUserNum & "' AND r078035='*' GROUP BY r078001,r078002,r078003)"
'2021/3/2 END

'Modify By Cheng 2002/04/12
'strSQL = "SELECT r078001,nvl(st02,r078002),NVL(A0902,A0903),DECODE(r078004,NULL,0,R078004),DECODE(r078005,NULL,0,R078005),DECODE(r078007,null,0,r078007),decode(r078008,null,0,r078008),decode(r078010,null,0,r078010),decode(r078011,null,0,r078011),decode(r078013,null,0,r078013),decode(r078014,null,0,r078014),decode(r078016,null,0,r078016),decode(r078017,null,0,r078017),decode(r078019,null,0,r078019),decode(r078020,null,0,r078020),decode(r078022,null,0,r078022),decode(r078023,null,0,r078023),decode(r078025,null,0,r078025),decode(r078026,null,0,r078026),'',r078002,r078003 FROM r020407_2,ACC090,staff WHERE r078002=st01(+) and ID='" & strUserNum & "' AND r078003=A0901(+) ORDER BY r078001,r078002,r078003 "
'911128 nick 修正都是同一 user 的情形
'strSQL = "SELECT r078001,nvl(st02,r078002),NVL(A0902,A0903)," & _
         "sum(DECODE(r078004,NULL,0,R078004)),sum(DECODE(r078005,NULL,0,R078005)),sum(DECODE(r078007,null,0,r078007))," & _
         "sum(decode(r078008,null,0,r078008)),sum(decode(r078010,null,0,r078010)),sum(decode(r078011,null,0,r078011))," & _
         "sum(decode(r078013,null,0,r078013)),sum(decode(r078014,null,0,r078014)),sum(decode(r078016,null,0,r078016))," & _
         "sum(decode(r078017,null,0,r078017)),sum(decode(r078019,null,0,r078019)),sum(decode(r078020,null,0,r078020))," & _
         "sum(decode(r078022,null,0,r078022)),sum(decode(r078023,null,0,r078023)),sum(decode(r078025,null,0,r078025))," & _
         "sum(decode(r078026,null,0,r078026)),'',r078002,r078003 FROM r020407_2,ACC090,staff WHERE r078002=st01(+) and ID='90019' AND r078003=A0901(+)" & _
         "GROUP BY r078001,nvl(st02,r078002),NVL(A0902,A0903),'',r078002,r078003 ORDER BY r078001,r078002,r078003"
'Modify By Sindy 2021/3/2 +||decode(r078029,null,'',r078029) ~ ||decode(r078035,null,'',r078035)
strSql = "SELECT r078001,nvl(st02,r078002),NVL(A0902,A0903)," & _
         "sum(DECODE(r078004,NULL,0,R078004)),sum(DECODE(r078005,NULL,0,R078005))||decode(r078029,null,'',r078029),sum(DECODE(r078007,null,0,r078007))," & _
         "sum(decode(r078008,null,0,r078008))||decode(r078030,null,'',r078030),sum(decode(r078010,null,0,r078010)),sum(decode(r078011,null,0,r078011))||decode(r078031,null,'',r078031)," & _
         "sum(decode(r078013,null,0,r078013)),sum(decode(r078014,null,0,r078014))||decode(r078032,null,'',r078032),sum(decode(r078016,null,0,r078016))," & _
         "sum(decode(r078017,null,0,r078017))||decode(r078033,null,'',r078033),sum(decode(r078019,null,0,r078019)),sum(decode(r078020,null,0,r078020))||decode(r078034,null,'',r078034)," & _
         "sum(decode(r078022,null,0,r078022)),sum(decode(r078023,null,0,r078023))||decode(r078035,null,'',r078035),sum(decode(r078025,null,0,r078025))," & _
         "sum(decode(r078026,null,0,r078026)),'',r078002,r078003 FROM r020407_2,ACC090,staff WHERE r078002=st01(+) and ID='" & strUserNum & "' AND r078003=A0901(+)" & _
         "GROUP BY r078001,nvl(st02,r078002),NVL(A0902,A0903),'',r078002,r078003,r078029,r078030,r078031,r078032,r078033,r078034,r078035 ORDER BY r078001,r078002,r078003"

CheckOC
Page = 1
TestOk = True
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        SavDay1 = CheckStr(.Fields(0))
        SavDay2 = CheckStr(.Fields(1))
        'SavDay3 = CheckStr(.Fields(19))
        strSql = "SELECT r077001,r077004,r077005,r077006,r077007,r077008,r077009,r077010 FROM r020407_1 WHERE ID='" & strUserNum & "' AND r077001=" & Val(SavDay1) & " "
        CheckOC2
        adoRecordset1.CursorLocation = adUseClient
        adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
            For i = 1 To 7
                StrTemp7(i) = StrToStr(CheckStr(adoRecordset1.Fields(i)), 7)
            Next i
            StrTemp7(8) = ""
'            For i = 1 To 7
'                If Len(StrTemp7(i)) = 0 Then
'                    If i <> 1 Then
'                        StrTemp7(i) = "合計"
'                    End If
'                    If i <> 7 Then
'                        If i = 1 Then
'                           StrTemp7(i) = "總計"
'                        Else
'                           StrTemp7(i + 1) = "總計"
'                        End If
'                    End If
'                    Exit For
'                End If
'            Next i
        End If
        CheckOC2
        PrintTitle
        PrintTitle2
        SavDay1 = " "
        SavDay2 = " "
        Do While .EOF = False
            For i = 0 To 19
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(1) = StrToStr(strTemp(1), 5)
            'strTemp(2) = StrToStr(strTemp(2), 3)
            strTemp(2) = StrToStr(strTemp(2), 5)
            If Len(Trim(SavDay1)) > 0 Or Len(Trim(SavDay2)) > 0 Then
                'Modify By Cheng 2004/04/07
'                If strTemp(1) <> SavDay2 Then
'                    ShowLine2
'                    PrintEnd2 (2)
'                    ShowLine2
'                    PrintEnd2 (3)
'                    ShowLine2
'                    PrintEnd2 (0)
'                    ShowLine2
'                    If Val(SavDay1) <> Val(strTemp(0)) Then
'                        PrintEnd2 (1)
'                        ShowLine2
'                        strSQL = "SELECT r077001,r077004,r077005,r077006,r077007,r077008,r077009,r077010 FROM r020407_1 WHERE ID='" & strUserNum & "' AND r077001=" & Val(CheckStr(.Fields(0))) & " "
'                        CheckOC2
'                        adoRecordset1.CursorLocation = adUseClient
'                        adoRecordset1.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'                        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'                            For i = 1 To 7
'                                StrTemp7(i) = StrToStr(CheckStr(adoRecordset1.Fields(i)), 7)
'                            Next i
'                            StrTemp7(8) = ""
'    '                        For i = 1 To 7
'    '                            If Len(StrTemp7(i)) = 0 Then
'    '                                If i <> 1 Then
'    '                                    StrTemp7(i) = "合計"
'    '                                End If
'    '                                If i <> 7 Then
'    '                                    If i = 1 Then
'    '                                       StrTemp7(i) = "總計"
'    '                                    Else
'    '                                       StrTemp7(i + 1) = "總計"
'    '                                    End If
'    '                                End If
'    '                                Exit For
'    '                            End If
'    '                        Next i
'                        End If
'                        CheckOC2
'                        Page = Page + 1
'                        SavDay1 = strTemp(0)
'                        Printer.NewPage
'                        PrintTitle
'                        PrintTitle2
'                    End If
'                    SavDay2 = CheckStr(.Fields(1))
                If Val(SavDay1) <> Val(strTemp(0)) Then
                    ShowLine2
                    PrintEnd2 (2)
                    ShowLine2
                    PrintEnd2 (3)
                    ShowLine2
                    PrintEnd2 (0)
                    ShowLine2
                    'Add By Sindy 2015/6/29
                    PrintEnd2 (4)
                    ShowLine2
                    PrintEnd2 (5)
                    ShowLine2
                    '2015/6/29 END
                    PrintEnd2 (1)
                    ShowLine2
                    strSql = "SELECT r077001,r077004,r077005,r077006,r077007,r077008,r077009,r077010 FROM r020407_1 WHERE ID='" & strUserNum & "' AND r077001=" & Val(CheckStr(.Fields(0))) & " "
                    CheckOC2
                    adoRecordset1.CursorLocation = adUseClient
                    adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                    If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                        For i = 1 To 7
                            StrTemp7(i) = StrToStr(CheckStr(adoRecordset1.Fields(i)), 7)
                        Next i
                        StrTemp7(8) = ""
                    End If
                    CheckOC2
                    Page = Page + 1
                    SavDay1 = strTemp(0)
                    Printer.NewPage
                    PrintTitle
                    PrintTitle2
                    SavDay2 = CheckStr(.Fields(1))
                ElseIf strTemp(1) <> SavDay2 Then
                    ShowLine2
                    PrintEnd2 (2)
                    ShowLine2
                    PrintEnd2 (3)
                    ShowLine2
                    PrintEnd2 (0)
                    ShowLine2
                    SavDay2 = CheckStr(.Fields(1))
                    'End
                Else
                   strTemp(1) = ""
                End If
            Else
               SavDay1 = strTemp(0)
               SavDay2 = CheckStr(.Fields(1))
            End If
            strTemp(2) = StrToStr(strTemp(2), 5)
            PrintDatil2
            If iPrint >= 14000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
                PrintTitle2
            End If
            .MoveNext
        Loop
    Else
      ShowNoData
      Exit Sub
    End If

End With
ShowLine2
PrintEnd2 (2)
ShowLine2
PrintEnd2 (3)
ShowLine2
PrintEnd2 (0)
ShowLine2
'Add By Sindy 2015/6/29
PrintEnd2 (4)
ShowLine2
PrintEnd2 (5)
ShowLine2
'2015/6/29 END
PrintEnd2 (1)
ShowLine2
Printer.EndDoc
ShowPrintOk
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
      PrintTitle
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
      PrintTitle
      PrintTitle2
   Else
      BolEndThisPage = False
   End If
End If
End Sub

'選擇收文
Sub PrintData1()
'Add By Cheng 2002/02/06
Dim strSQL_1 As String
Dim strSQL_2 As String
Dim strSQL_3 As String
Dim rstmp_1 As New ADODB.Recordset
Dim rstmp_2 As New ADODB.Recordset
Dim rstmp_3 As New ADODB.Recordset
Dim dblTotal As Double '總計

BolEndThisPage = False

'搜尋固定區的第一頁
'Modify By Sindy 2021/6/22 , ==> +
strSQL_1 = "SELECT r078001,nvl(st02,r078002),NVL(A0902,A0903),decode(r078004,null,0,r078004)+decode(r078005,null,0,r078005)+decode(r078006,null,0,r078006)" & _
            "+decode(r078007,null,0,r078007)+decode(r078008,null,0,r078008)+decode(r078009,null,0,r078009)" & _
            "+decode(r078010,null,0,r078010)+decode(r078011,null,0,r078011)+decode(r078012,null,0,r078012)" & _
            "+decode(r078013,null,0,r078013)+decode(r078014,null,0,r078014)+decode(r078015,null,0,r078015)" & _
            "+decode(r078016,NULL,0,r078016)+DECODE(r078017,NULL,0,r078017)+DECODE(r078018,NULL,0,r078018)" & _
            "+DECODE(R078019,NULL,0,R078019)+DECODE(r078020,NULL,0,r078020)+DECODE(r078021,NULL,0,r078021)" & _
            "+DECODE(r078022,NULL,0,r078022)+DECODE(r078023,NULL,0,r078023)+DECODE(r078024,NULL,0,r078024)" & _
            "+DECODE(r078025,NULL,0,r078025)+DECODE(r078026,NULL,0,r078026)+DECODE(r078027,NULL,0,r078027) As Total,'',r078002,r078003 " & _
            " FROM r020407_2,ACC090,staff " & _
            " WHERE r078002=st01(+) and ID='" & strUserNum & "' AND r078003=A0901(+) And R078001='1' ORDER BY r078001,r078002,r078003 "
If rstmp_1.State <> adStateClosed Then rstmp_1.Close
rstmp_1.CursorLocation = adUseClient
rstmp_1.Open strSQL_1, cnnConnection, adOpenStatic, adLockReadOnly
If rstmp_1.RecordCount > 0 Then
   rstmp_1.MoveFirst
   While Not rstmp_1.EOF
      '搜尋固定區的第二頁
      'Modify By Sindy 2021/6/22 , ==> +
      strSQL_2 = "SELECT r078001,nvl(st02,r078002),NVL(A0902,A0903),decode(r078004,null,0,r078004)+decode(r078005,null,0,r078005)+decode(r078006,null,0,r078006)" & _
                  "+decode(r078007,null,0,r078007)+decode(r078008,null,0,r078008)+decode(r078009,null,0,r078009)" & _
                  "+decode(r078010,null,0,r078010)+decode(r078011,null,0,r078011)+decode(r078012,null,0,r078012)" & _
                  "+decode(r078013,null,0,r078013)+decode(r078014,null,0,r078014)+decode(r078015,null,0,r078015)" & _
                  "+decode(r078016,NULL,0,r078016)+DECODE(r078017,NULL,0,r078017)+DECODE(r078018,NULL,0,r078018)" & _
                  "+DECODE(R078019,NULL,0,R078019)+DECODE(r078020,NULL,0,r078020)+DECODE(r078021,NULL,0,r078021)" & _
                  "+DECODE(r078022,NULL,0,r078022)+DECODE(r078023,NULL,0,r078023)+DECODE(r078024,NULL,0,r078024)" & _
                  "+DECODE(r078025,NULL,0,r078025)+DECODE(r078026,NULL,0,r078026)+DECODE(r078027,NULL,0,r078027) As Total,'',r078002,r078003 " & _
                  " FROM r020407_2,ACC090,staff " & _
                  " WHERE r078002=st01(+) and ID='" & strUserNum & "' AND r078003=A0901(+) And R078001='2' And R078002='" & rstmp_1("R078002").Value & "' And R078003='" & rstmp_1("R078003").Value & "' ORDER BY r078001,r078002,r078003 "
      If rstmp_2.State <> adStateClosed Then rstmp_2.Close
      rstmp_2.CursorLocation = adUseClient
      rstmp_2.Open strSQL_2, cnnConnection, adOpenStatic, adLockReadOnly
      If rstmp_2.RecordCount > 0 Then
         rstmp_2.MoveFirst
         '若固定區合計為0
         If rstmp_1("Total") + rstmp_2("Total") = 0 Then
            'Modify By Cheng 2003/12/29
            '不要刪除固定區資料
'            cnnConnection.Execute "Delete From R020407_2 Where ID ='" & strUserNum & "' And R078002='" & rstmp_1("R078002").Value & "' And R078003='" & rstmp_1("R078003").Value & "' And (R078001='1' Or R078001='2') "
            'End
            '搜尋固定區的第三頁以上
            'Modify By Sindy 2021/6/22 , ==> +
            strSQL_3 = "SELECT r078001,nvl(st02,r078002),NVL(A0902,A0903),decode(r078004,null,0,r078004)+decode(r078005,null,0,r078005)+decode(r078006,null,0,r078006)" & _
                        "+decode(r078007,null,0,r078007)+decode(r078008,null,0,r078008)+decode(r078009,null,0,r078009)" & _
                        "+decode(r078010,null,0,r078010)+decode(r078011,null,0,r078011)+decode(r078012,null,0,r078012)" & _
                        "+decode(r078013,null,0,r078013)+decode(r078014,null,0,r078014)+decode(r078015,null,0,r078015)" & _
                        "+decode(r078016,NULL,0,r078016)+DECODE(r078017,NULL,0,r078017)+DECODE(r078018,NULL,0,r078018)" & _
                        "+DECODE(R078019,NULL,0,R078019)+DECODE(r078020,NULL,0,r078020)+DECODE(r078021,NULL,0,r078021)" & _
                        "+DECODE(r078022,NULL,0,r078022)+DECODE(r078023,NULL,0,r078023)+DECODE(r078024,NULL,0,r078024)" & _
                        "+DECODE(r078025,NULL,0,r078025)+DECODE(r078026,NULL,0,r078026)+DECODE(r078027,NULL,0,r078027) As Total,'',r078002,r078003 " & _
                        " FROM r020407_2,ACC090,staff " & _
                        " WHERE r078002=st01(+) and ID='" & strUserNum & "' AND r078003=A0901(+) And (R078001<>'1' And R078001<>'2') And R078002='" & rstmp_1("R078002").Value & "' And R078003='" & rstmp_1("R078003").Value & "' ORDER BY r078001,r078002,r078003 "
            If rstmp_3.State <> adStateClosed Then rstmp_3.Close
            rstmp_3.CursorLocation = adUseClient
            rstmp_3.Open strSQL_3, cnnConnection, adOpenStatic, adLockReadOnly
            dblTotal = 0
            If rstmp_3.RecordCount > 0 Then
               rstmp_3.MoveFirst
               While Not rstmp_3.EOF
                  dblTotal = dblTotal + rstmp_3("Total")
                  rstmp_3.MoveNext
               Wend
               If dblTotal = 0 Then
                  cnnConnection.Execute "Delete From R020407_2 Where ID ='" & strUserNum & "' And R078002='" & rstmp_1("R078002").Value & "' And R078003='" & rstmp_1("R078003").Value & "' And (R078001<>'1' And R078001<>'2') "
               End If
            End If
         End If
      End If
      rstmp_1.MoveNext
   Wend
Else
   ShowNoData
   If rstmp_1.State <> adStateClosed Then rstmp_1.Close
   Set rstmp_1 = Nothing
   Screen.MousePointer = vbDefault
   Exit Sub
End If
If rstmp_1.State <> adStateClosed Then rstmp_1.Close
Set rstmp_1 = Nothing
If rstmp_2.State <> adStateClosed Then rstmp_2.Close
Set rstmp_2 = Nothing
If rstmp_3.State <> adStateClosed Then rstmp_3.Close
Set rstmp_3 = Nothing

'Add By Sindy 2016/9/8
cnnConnection.Execute "Update R020407_2 set R078028=decode(substr(R078003,1,1),'F',' ','*') Where ID ='" & strUserNum & "'"
'2016/9/8 END
'Add By Sindy 2021/3/2
cnnConnection.Execute "UPDATE r020407_2 SET r078029='*' where r078001||r078002||r078003 in(SELECT r078001||r078002||r078003 FROM r020407_2 WHERE ID='" & strUserNum & "' AND r078029='*' GROUP BY r078001,r078002,r078003)"
cnnConnection.Execute "UPDATE r020407_2 SET r078030='*' where r078001||r078002||r078003 in(SELECT r078001||r078002||r078003 FROM r020407_2 WHERE ID='" & strUserNum & "' AND r078030='*' GROUP BY r078001,r078002,r078003)"
cnnConnection.Execute "UPDATE r020407_2 SET r078031='*' where r078001||r078002||r078003 in(SELECT r078001||r078002||r078003 FROM r020407_2 WHERE ID='" & strUserNum & "' AND r078031='*' GROUP BY r078001,r078002,r078003)"
cnnConnection.Execute "UPDATE r020407_2 SET r078032='*' where r078001||r078002||r078003 in(SELECT r078001||r078002||r078003 FROM r020407_2 WHERE ID='" & strUserNum & "' AND r078032='*' GROUP BY r078001,r078002,r078003)"
cnnConnection.Execute "UPDATE r020407_2 SET r078033='*' where r078001||r078002||r078003 in(SELECT r078001||r078002||r078003 FROM r020407_2 WHERE ID='" & strUserNum & "' AND r078033='*' GROUP BY r078001,r078002,r078003)"
cnnConnection.Execute "UPDATE r020407_2 SET r078034='*' where r078001||r078002||r078003 in(SELECT r078001||r078002||r078003 FROM r020407_2 WHERE ID='" & strUserNum & "' AND r078034='*' GROUP BY r078001,r078002,r078003)"
cnnConnection.Execute "UPDATE r020407_2 SET r078035='*' where r078001||r078002||r078003 in(SELECT r078001||r078002||r078003 FROM r020407_2 WHERE ID='" & strUserNum & "' AND r078035='*' GROUP BY r078001,r078002,r078003)"
'2021/3/2 END

'Modify By Cheng 2002/04/12
'strSQL = "SELECT r078001,nvl(st02,r078002),NVL(A0902,A0903),decode(r078004,null,0,r078004),decode(r078005,null,0,r078005),decode(r078006,null,0,r078006),decode(r078007,null,0,r078007),decode(r078008,null,0,r078008),decode(r078009,null,0,r078009),decode(r078010,null,0,r078010),decode(r078011,null,0,r078011),decode(r078012,null,0,r078012),decode(r078013,null,0,r078013),decode(r078014,null,0,r078014),decode(r078015,null,0,r078015),decode(r078016,NULL,0,r078016),DECODE(r078017,NULL,0,r078017),DECODE(r078018,NULL,0,r078018),DECODE(R078019,NULL,0,R078019),DECODE(r078020,NULL,0,r078020),DECODE(r078021,NULL,0,r078021),DECODE(r078022,NULL,0,r078022),DECODE(r078023,NULL,0,r078023),DECODE(r078024,NULL,0,r078024),DECODE(r078025,NULL,0,r078025),DECODE(r078026,NULL,0,r078026),DECODE(r078027,NULL,0,r078027),'',r078002,r078003 FROM r020407_2,ACC090,staff WHERE r078002=st01(+) and ID='" & strUserNum & "' AND r078003=A0901(+) ORDER BY r078001,r078002,r078003 "
'911128 nick 修正都是 90019 的 user
'strSQL = "SELECT r078001,nvl(st02,r078002),NVL(A0902,A0903)," & _
         "sum(decode(r078004,null,0,r078004)),sum(decode(r078005,null,0,r078005)),sum(decode(r078006,null,0,r078006))," & _
         "sum(decode(r078007,null,0,r078007)),sum(decode(r078008,null,0,r078008)),sum(decode(r078009,null,0,r078009))," & _
         "sum(decode(r078010,null,0,r078010)),sum(decode(r078011,null,0,r078011)),sum(decode(r078012,null,0,r078012))," & _
         "sum(decode(r078013,null,0,r078013)),sum(decode(r078014,null,0,r078014)),sum(decode(r078015,null,0,r078015))," & _
         "sum(decode(r078016,NULL,0,r078016)),sum(DECODE(r078017,NULL,0,r078017)),sum(DECODE(r078018,NULL,0,r078018))," & _
         "sum(DECODE(R078019,NULL,0,R078019)),sum(DECODE(r078020,NULL,0,r078020)),sum(DECODE(r078021,NULL,0,r078021))," & _
         "sum(DECODE(r078022,NULL,0,r078022)),sum(DECODE(r078023,NULL,0,r078023)),sum(DECODE(r078024,NULL,0,r078024))," & _
         "sum(DECODE(r078025,NULL,0,r078025)),sum(DECODE(r078026,NULL,0,r078026)),sum(DECODE(r078027,NULL,0,r078027))," & _
         "'',r078002,r078003 FROM r020407_2,ACC090,staff WHERE r078002=st01(+) and ID='90019' AND r078003=A0901(+) group by r078001,nvl(st02,r078002),NVL(A0902,A0903),'',r078002,r078003 ORDER BY r078001,r078002,r078003"
'Modify By Sindy 2021/3/2 +||decode(r078029,null,'',r078029) ~ ||decode(r078035,null,'',r078035)
strSql = "SELECT r078001,nvl(st02,r078002),NVL(A0902,A0903)," & _
         "sum(decode(r078004,null,0,r078004)),sum(decode(r078005,null,0,r078005))||decode(r078029,null,'',r078029),sum(decode(r078006,null,0,r078006))," & _
         "sum(decode(r078007,null,0,r078007)),sum(decode(r078008,null,0,r078008))||decode(r078030,null,'',r078030),sum(decode(r078009,null,0,r078009))," & _
         "sum(decode(r078010,null,0,r078010)),sum(decode(r078011,null,0,r078011))||decode(r078031,null,'',r078031),sum(decode(r078012,null,0,r078012))," & _
         "sum(decode(r078013,null,0,r078013)),sum(decode(r078014,null,0,r078014))||decode(r078032,null,'',r078032),sum(decode(r078015,null,0,r078015))," & _
         "sum(decode(r078016,NULL,0,r078016)),sum(DECODE(r078017,NULL,0,r078017))||decode(r078033,null,'',r078033),sum(DECODE(r078018,NULL,0,r078018))," & _
         "sum(DECODE(R078019,NULL,0,R078019)),sum(DECODE(r078020,NULL,0,r078020))||decode(r078034,null,'',r078034),sum(DECODE(r078021,NULL,0,r078021))," & _
         "sum(DECODE(r078022,NULL,0,r078022)),sum(DECODE(r078023,NULL,0,r078023))||decode(r078035,null,'',r078035),sum(DECODE(r078024,NULL,0,r078024))," & _
         "sum(DECODE(r078025,NULL,0,r078025)),sum(DECODE(r078026,NULL,0,r078026)),sum(DECODE(r078027,NULL,0,r078027))," & _
         "'',r078002,r078003 FROM r020407_2,ACC090,staff WHERE r078002=st01(+) and ID='" & strUserNum & "' AND r078003=A0901(+) group by r078001,nvl(st02,r078002),NVL(A0902,A0903),'',r078002,r078003,r078029,r078030,r078031,r078032,r078033,r078034,r078035 ORDER BY r078001,r078002,r078003"

CheckOC
Page = 1
TestOk = True
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        SavDay1 = CheckStr(.Fields(0))
        SavDay2 = CheckStr(.Fields(1))
        strSql = "SELECT distinct r077001,r077004,r077005,r077006,r077007,r077008,r077009,r077010 FROM r020407_1 WHERE ID='" & strUserNum & "' AND r077001=" & Val(SavDay1) & " "
        CheckOC2
        adoRecordset1.CursorLocation = adUseClient
        adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
            For i = 1 To 7
                StrTemp7(i) = StrToStr(CheckStr(adoRecordset1.Fields(i)), 7)
            Next i
            StrTemp7(8) = ""
            'StrTemp7(9) = ""
'            For i = 1 To 7
'                If Len(StrTemp7(i)) = 0 Then
'                    If i <> 1 Then
'                        StrTemp7(i) = "合計"
'                    End If
'                    If i <> 7 Then
'                        If i = 1 Then
'                           StrTemp7(i) = "總計"
'                        Else
'                           StrTemp7(i + 1) = "總計"
'                        End If
'                    End If
'                    Exit For
'                End If
'            Next i
        End If
        CheckOC2
        PrintTitle
        PrintTitle1
        SavDay1 = " "
        SavDay2 = " "
        Do While .EOF = False
            For i = 0 To 27
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(1) = StrToStr(strTemp(1), 5)
            'strTemp(2) = StrToStr(strTemp(2), 3)
            strTemp(2) = StrToStr(strTemp(2), 5)
            If Len(Trim(SavDay1)) > 0 Or Len(Trim(SavDay2)) > 0 Then
                'Modify By Cheng 2004/04/07
'                If strTemp(1) <> SavDay2 Then
'                    ShowLine1
'                    PrintEnd1 (2)
'                    ShowLine1
'                    PrintEnd1 (3)
'                    ShowLine1
'                    PrintEnd1 (0)
'                    ShowLine1
'                    If Val(SavDay1) <> Val(strTemp(0)) Then
'                       PrintEnd1 (1)
'                        ShowLine1
'                        strSQL = "SELECT r077001,r077004,r077005,r077006,r077007,r077008,r077009,r077010 FROM r020407_1 WHERE ID='" & strUserNum & "' AND r077001=" & Val(CheckStr(.Fields(0))) & " "
'                        CheckOC2
'                        adoRecordset1.CursorLocation = adUseClient
'                        adoRecordset1.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'                        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'                            For i = 1 To 7
'                                StrTemp7(i) = StrToStr(CheckStr(adoRecordset1.Fields(i)), 7)
'                            Next i
'                            StrTemp7(8) = ""
'    '                        For i = 1 To 7
'    '                            If Len(StrTemp7(i)) = 0 Then
'    '                                If i <> 1 Then
'    '                                    StrTemp7(i) = "合計"
'    '                                End If
'    '                                If i <> 7 Then
'    '                                    If i = 1 Then
'    '                                       StrTemp7(i) = "總計"
'    '                                    Else
'    '                                       StrTemp7(i + 1) = "總計"
'    '                                    End If
'    '                                End If
'    '                                Exit For
'    '                            End If
'    '                        Next i
'                        End If
'                        CheckOC2
'                        Page = Page + 1
'                        SavDay1 = strTemp(0)
'                        Printer.NewPage
'                        PrintTitle
'                        PrintTitle1
'                    End If
'                    SavDay2 = CheckStr(.Fields(1))
'                Else
'                   strTemp(1) = ""
'                End If
                If Val(SavDay1) <> Val(strTemp(0)) Then
                    ShowLine1
                    PrintEnd1 (2)
                    ShowLine1
                    PrintEnd1 (3)
                    ShowLine1
                    PrintEnd1 (0)
                    ShowLine1
                    'Add By Sindy 2015/6/29
                    PrintEnd1 (4)
                    ShowLine1
                    PrintEnd1 (5)
                    ShowLine1
                    '2015/6/29 END
                    PrintEnd1 (1)
                     ShowLine1
                     strSql = "SELECT r077001,r077004,r077005,r077006,r077007,r077008,r077009,r077010 FROM r020407_1 WHERE ID='" & strUserNum & "' AND r077001=" & Val(CheckStr(.Fields(0))) & " "
                     CheckOC2
                     adoRecordset1.CursorLocation = adUseClient
                     adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                         For i = 1 To 7
                             StrTemp7(i) = StrToStr(CheckStr(adoRecordset1.Fields(i)), 7)
                         Next i
                         StrTemp7(8) = ""
                     End If
                     CheckOC2
                     Page = Page + 1
                     SavDay1 = strTemp(0)
                     Printer.NewPage
                     PrintTitle
                     PrintTitle1
                    SavDay2 = CheckStr(.Fields(1))
                ElseIf strTemp(1) <> SavDay2 Then
                    ShowLine1
                    PrintEnd1 (2)
                    ShowLine1
                    PrintEnd1 (3)
                    ShowLine1
                    PrintEnd1 (0)
                    ShowLine1
                    SavDay2 = CheckStr(.Fields(1))
                Else
                   strTemp(1) = ""
                End If
                'End
            Else
               SavDay1 = strTemp(0)
               SavDay2 = CheckStr(.Fields(1))
            End If
            strTemp(2) = StrToStr(strTemp(2), 5)
            PrintDatil1
            If iPrint >= 14000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
                PrintTitle1
            End If
            .MoveNext
        Loop
   Else
      Exit Sub
    End If
End With
'adoRecordset.MoveLast
ShowLine1
PrintEnd1 (2)
ShowLine1
PrintEnd1 (3)
ShowLine1
PrintEnd1 (0)
ShowLine1
'Add By Sindy 2015/6/29
PrintEnd1 (4)
ShowLine1
PrintEnd1 (5)
ShowLine1
'2015/6/29 END
PrintEnd1 (1)
ShowLine1
Printer.EndDoc
ShowPrintOk
End Sub

'收文
Sub PrintEnd1(Strindex As Integer)
adoRecordset.MovePrevious
'Add By Cheng 2003/07/01
If adoRecordset.BOF Then adoRecordset.MoveFirst
Select Case Strindex
Case 0
     strSql = "select '個人小計','','',sum(r078004),sum(r078005),sum(r078006),sum(r078007),sum(r078008),sum(r078009),sum(r078010),sum(r078011),sum(r078012),sum(r078013),sum(r078014),sum(r078015),sum(r078016),sum(r078017),sum(r078018),sum(r078019),sum(r078020),sum(r078021),sum(r078022),sum(r078023),sum(r078024),sum(r078025),sum(r078026),sum(r078027) from r020407_2 where id='" & strUserNum & "' " & IIf(Len(CheckStr(adoRecordset.Fields(28))) = 0, " and r078002 is null ", " AND r078002='" & CheckStr(adoRecordset.Fields(28)) & "' ") & " AND r078001='" & SavDay1 & "' "
Case 1
     strSql = "select '全所總計','','',sum(r078004),sum(r078005),sum(r078006),sum(r078007),sum(r078008),sum(r078009),sum(r078010),sum(r078011),sum(r078012),sum(r078013),sum(r078014),sum(r078015),sum(r078016),sum(r078017),sum(r078018),sum(r078019),sum(r078020),sum(r078021),sum(r078022),sum(r078023),sum(r078024),sum(r078025),sum(r078026),sum(r078027) from r020407_2 where id='" & strUserNum & "' AND r078001=" & Val(SavDay1) & " "
Case 2
     'edit by nickc 2005/05/12
     'StrSql = "select '國內小計','','',sum(r078004),sum(r078005),sum(r078006),sum(r078007),sum(r078008),sum(r078009),sum(r078010),sum(r078011),sum(r078012),sum(r078013),sum(r078014),sum(r078015),sum(r078016),sum(r078017),sum(r078018),sum(r078019),sum(r078020),sum(r078021),sum(r078022),sum(r078023),sum(r078024),sum(r078025),sum(r078026),sum(r078027) from r020407_2 where id='" & strUserNum & "' " & IIf(Len(CheckStr(adoRecordset.Fields(28))) = 0, " and r078002 is null ", " AND r078002='" & CheckStr(adoRecordset.Fields(28)) & "' ") & " AND r078001='" & SavDay1 & "' and r078028='*' "
     strSql = "select '國內業務小計','','',sum(r078004),sum(r078005),sum(r078006),sum(r078007),sum(r078008),sum(r078009),sum(r078010),sum(r078011),sum(r078012),sum(r078013),sum(r078014),sum(r078015),sum(r078016),sum(r078017),sum(r078018),sum(r078019),sum(r078020),sum(r078021),sum(r078022),sum(r078023),sum(r078024),sum(r078025),sum(r078026),sum(r078027) from r020407_2 where id='" & strUserNum & "' " & IIf(Len(CheckStr(adoRecordset.Fields(28))) = 0, " and r078002 is null ", " AND r078002='" & CheckStr(adoRecordset.Fields(28)) & "' ") & " AND r078001='" & SavDay1 & "' and r078028='*' "
Case 3
     'edit by nickc 2005/05/12
     'StrSql = "select '國外小計','','',sum(r078004),sum(r078005),sum(r078006),sum(r078007),sum(r078008),sum(r078009),sum(r078010),sum(r078011),sum(r078012),sum(r078013),sum(r078014),sum(r078015),sum(r078016),sum(r078017),sum(r078018),sum(r078019),sum(r078020),sum(r078021),sum(r078022),sum(r078023),sum(r078024),sum(r078025),sum(r078026),sum(r078027) from r020407_2 where id='" & strUserNum & "' " & IIf(Len(CheckStr(adoRecordset.Fields(28))) = 0, " and r078002 is null ", " AND r078002='" & CheckStr(adoRecordset.Fields(28)) & "' ") & " AND r078001='" & SavDay1 & "' and (r078028=' ' or r078028='' or r078028 is null) "
     strSql = "select '國外業務小計','','',sum(r078004),sum(r078005),sum(r078006),sum(r078007),sum(r078008),sum(r078009),sum(r078010),sum(r078011),sum(r078012),sum(r078013),sum(r078014),sum(r078015),sum(r078016),sum(r078017),sum(r078018),sum(r078019),sum(r078020),sum(r078021),sum(r078022),sum(r078023),sum(r078024),sum(r078025),sum(r078026),sum(r078027) from r020407_2 where id='" & strUserNum & "' " & IIf(Len(CheckStr(adoRecordset.Fields(28))) = 0, " and r078002 is null ", " AND r078002='" & CheckStr(adoRecordset.Fields(28)) & "' ") & " AND r078001='" & SavDay1 & "' and (r078028=' ' or r078028='' or r078028 is null) "
     BolEndThisPage = True
'Add By Sindy 2015/6/29
Case 4
     strSql = "select '國內業務總計','','',sum(r078004),sum(r078005),sum(r078006),sum(r078007),sum(r078008),sum(r078009),sum(r078010),sum(r078011),sum(r078012),sum(r078013),sum(r078014),sum(r078015),sum(r078016),sum(r078017),sum(r078018),sum(r078019),sum(r078020),sum(r078021),sum(r078022),sum(r078023),sum(r078024),sum(r078025),sum(r078026),sum(r078027) from r020407_2 where id='" & strUserNum & "' AND r078001='" & SavDay1 & "' and r078028='*' "
Case 5
     strSql = "select '國外業務總計','','',sum(r078004),sum(r078005),sum(r078006),sum(r078007),sum(r078008),sum(r078009),sum(r078010),sum(r078011),sum(r078012),sum(r078013),sum(r078014),sum(r078015),sum(r078016),sum(r078017),sum(r078018),sum(r078019),sum(r078020),sum(r078021),sum(r078022),sum(r078023),sum(r078024),sum(r078025),sum(r078026),sum(r078027) from r020407_2 where id='" & strUserNum & "' AND r078001='" & SavDay1 & "' and (r078028=' ' or r078028='' or r078028 is null) "
'2015/6/29 END
Case Else
     Exit Sub
End Select
adoRecordset.MoveNext
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
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print CheckStr(.Fields(0))
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print CheckStr(.Fields(1))
            For i = 2 To 22
               If Len(Trim(StrTemp7(Int((i + 1) / 3)))) <> 0 Then
                   Select Case i
                  Case 3, 6, 9, 12, 15, 18, 21
                       Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(CheckNum(.Fields(i + 1)) & " ", "###,##0.0"))
                       Printer.CurrentY = iPrint
                       Printer.Print Format(CheckNum(.Fields(i + 1)) & " ", "###,##0.0")
                  Case Else
                       Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(CheckNum(.Fields(i + 1)))
                       Printer.CurrentY = iPrint
                       Printer.Print CheckNum(.Fields(i + 1))
                  End Select
               End If
            Next i
            iPrint = iPrint + 300
            If iPrint >= 14000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
                PrintTitle1
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC2
End Sub

'發文
Sub PrintEnd2(Strindex As Integer)
Dim strTemp As String

adoRecordset.MovePrevious
'Add By Cheng 2003/07/01
If adoRecordset.BOF Then adoRecordset.MoveFirst
Select Case Strindex
Case 0
     strSql = "select '個人小計','','',sum(r078004),sum(r078005),sum(r078007),sum(r078008),sum(r078010),sum(r078011),sum(r078013),sum(r078014),sum(r078016),sum(r078017),sum(r078019),sum(r078020),sum(r078022),sum(r078023),sum(r078025),sum(r078026) from r020407_2 where id='" & strUserNum & "' " & IIf(Len(CheckStr(adoRecordset.Fields(20))) = 0, " and r078002 is null ", " AND r078002='" & CheckStr(adoRecordset.Fields(20)) & "' ") & " AND r078001=" & Val(SavDay1) & " "
Case 1
     strSql = "select '全所總計','','',sum(r078004),sum(r078005),sum(r078007),sum(r078008),sum(r078010),sum(r078011),sum(r078013),sum(r078014),sum(r078016),sum(r078017),sum(r078019),sum(r078020),sum(r078022),sum(r078023),sum(r078025),sum(r078026) from r020407_2 where id='" & strUserNum & "' AND r078001=" & Val(SavDay1) & " "
Case 2
     'edit by nickc 2005/05/12
     'StrSql = "select '國內小計','','',sum(r078004),sum(r078005),sum(r078007),sum(r078008),sum(r078010),sum(r078011),sum(r078013),sum(r078014),sum(r078016),sum(r078017),sum(r078019),sum(r078020),sum(r078022),sum(r078023),sum(r078025),sum(r078026) from r020407_2 where id='" & strUserNum & "' " & IIf(Len(CheckStr(adoRecordset.Fields(20))) = 0, " and r078002 is null ", " AND r078002='" & CheckStr(adoRecordset.Fields(20)) & "' ") & " AND r078001=" & Val(SavDay1) & " and r078028='*' "
     strSql = "select '國內業務小計','','',sum(r078004),sum(r078005),sum(r078007),sum(r078008),sum(r078010),sum(r078011),sum(r078013),sum(r078014),sum(r078016),sum(r078017),sum(r078019),sum(r078020),sum(r078022),sum(r078023),sum(r078025),sum(r078026) from r020407_2 where id='" & strUserNum & "' " & IIf(Len(CheckStr(adoRecordset.Fields(20))) = 0, " and r078002 is null ", " AND r078002='" & CheckStr(adoRecordset.Fields(20)) & "' ") & " AND r078001=" & Val(SavDay1) & " and r078028='*' "
Case 3
     'edit by nickc 2005/05/12
     'StrSql = "select '國外小計','','',sum(r078004),sum(r078005),sum(r078007),sum(r078008),sum(r078010),sum(r078011),sum(r078013),sum(r078014),sum(r078016),sum(r078017),sum(r078019),sum(r078020),sum(r078022),sum(r078023),sum(r078025),sum(r078026) from r020407_2 where id='" & strUserNum & "' " & IIf(Len(CheckStr(adoRecordset.Fields(20))) = 0, " and r078002 is null ", " AND r078002='" & CheckStr(adoRecordset.Fields(20)) & "' ") & " AND r078001=" & Val(SavDay1) & " and (r078028 is null or r078028='' or r078028=' ') "
     strSql = "select '國外業務小計','','',sum(r078004),sum(r078005),sum(r078007),sum(r078008),sum(r078010),sum(r078011),sum(r078013),sum(r078014),sum(r078016),sum(r078017),sum(r078019),sum(r078020),sum(r078022),sum(r078023),sum(r078025),sum(r078026) from r020407_2 where id='" & strUserNum & "' " & IIf(Len(CheckStr(adoRecordset.Fields(20))) = 0, " and r078002 is null ", " AND r078002='" & CheckStr(adoRecordset.Fields(20)) & "' ") & " AND r078001=" & Val(SavDay1) & " and (r078028 is null or r078028='' or r078028=' ') "
     BolEndThisPage = True
'Add By Sindy 2015/6/29
Case 4
     strSql = "select '國內業務總計','','',sum(r078004),sum(r078005),sum(r078007),sum(r078008),sum(r078010),sum(r078011),sum(r078013),sum(r078014),sum(r078016),sum(r078017),sum(r078019),sum(r078020),sum(r078022),sum(r078023),sum(r078025),sum(r078026) from r020407_2 where id='" & strUserNum & "' AND r078001=" & Val(SavDay1) & " and r078028='*' "
Case 5
     strSql = "select '國外業務總計','','',sum(r078004),sum(r078005),sum(r078007),sum(r078008),sum(r078010),sum(r078011),sum(r078013),sum(r078014),sum(r078016),sum(r078017),sum(r078019),sum(r078020),sum(r078022),sum(r078023),sum(r078025),sum(r078026) from r020407_2 where id='" & strUserNum & "' AND r078001=" & Val(SavDay1) & " and (r078028 is null or r078028='' or r078028=' ') "
'2015/6/29 END
Case Else
     Exit Sub
End Select
adoRecordset.MoveNext
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 2 To 14 Step 2
                Printer.Line (PLeft(i) - 50, iPrint - 150)-(PLeft(i) - 50, iPrint + 450)
            Next i
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print CheckStr(.Fields(0))
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print CheckStr(.Fields(1))
            For i = 2 To 15
'               If Trim(StrTemp7(Int((i) / 2))) = "總計" Then
'                  strSql = "select sum(r078025),sum(r078026) from r020407_2 where id='" & strUserNum & "' AND r078001=" & Val(SavDay1) & " "
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                  If intI = 1 Then
'                     If i = 13 Then
'                        strTemp = RsTemp.Fields(0)
'                     Else
'                        strTemp = RsTemp.Fields(1)
'                     End If
'                     Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(CheckNum(strTemp) & " ", "###,##0.0"))
'                     Printer.CurrentY = iPrint
'                     Printer.Print Format(CheckNum(strTemp) & " ", "###,##0.0")
'                  End If
'               Else
               If Len(Trim(StrTemp7(Int((i) / 2)))) <> 0 Then
                  Select Case i
                  Case 3, 5, 7, 9, 11, 13, 15
                     Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(CheckNum(.Fields(i + 1)) & " ", "###,##0.0"))
                     Printer.CurrentY = iPrint
                     Printer.Print Format(CheckNum(.Fields(i + 1)) & " ", "###,##0.0")
                  Case Else
                     Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(CheckNum(.Fields(i + 1)))
                     Printer.CurrentY = iPrint
                     Printer.Print CheckNum(.Fields(i + 1))
                  End Select
               End If
            Next i
            iPrint = iPrint + 300
            If iPrint >= 14000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
                PrintTitle2
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC2
End Sub


Sub PrintTitle()
iPrint = 0
Printer.Font.Name = "細明體"
'Printer.Orientation = 1 'Removed by Morgan 2015/6/3
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
If Val(txt1(1)) = 1 Then
    If Val(SavDay1) <= 2 Then
        Printer.Print GetTitleNick & "商爭案承辦人收文統計表(固定)"
    Else
        Printer.Print GetTitleNick & "商爭案承辦人收文統計表(非固定)"
    End If
Else
    If Val(SavDay1) <= 2 Then
        Printer.Print GetTitleNick & "商爭案承辦人發文統計表(固定)"
    Else
        Printer.Print GetTitleNick & "商爭案承辦人發文統計表(非固定)"
    End If
End If
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 7000
Printer.CurrentY = iPrint
If Val(txt1(1)) = 1 Then
    Printer.Print "收文日：" & Format(ChangeTStringToTDateString(txt1(2)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(3))
Else
    Printer.Print "發文日：" & Format(ChangeTStringToTDateString(txt1(2)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(3))
End If
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
Printer.Print "申請國家：" & Me.txt1(4).Text & " － " & Me.txt1(5).Text
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
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
For i = 2 To 14 Step 2
    Printer.Line (PLeft(i) - 50, iPrint + 150)-(PLeft(i) - 50, iPrint + 1350)
Next i
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle2
    Exit Sub
End If
Printer.CurrentX = PLeft1(1) - (Printer.TextWidth(CheckStr(StrTemp7(1))) / 2)
Printer.CurrentY = iPrint
Printer.Print CheckStr(StrTemp7(1))
Printer.CurrentX = PLeft1(2) - (Printer.TextWidth(CheckStr(StrTemp7(2))) / 2)
Printer.CurrentY = iPrint
Printer.Print CheckStr(StrTemp7(2))
Printer.CurrentX = PLeft1(3) - (Printer.TextWidth(CheckStr(StrTemp7(3))) / 2)
Printer.CurrentY = iPrint
Printer.Print CheckStr(StrTemp7(3))
Printer.CurrentX = PLeft1(4) - (Printer.TextWidth(CheckStr(StrTemp7(4))) / 2)
Printer.CurrentY = iPrint
Printer.Print CheckStr(StrTemp7(4))
Printer.CurrentX = PLeft1(5) - (Printer.TextWidth(CheckStr(StrTemp7(5))) / 2)
Printer.CurrentY = iPrint
Printer.Print CheckStr(StrTemp7(5))
Printer.CurrentX = PLeft1(6) - (Printer.TextWidth(CheckStr(StrTemp7(6))) / 2)
Printer.CurrentY = iPrint
Printer.Print CheckStr(StrTemp7(6))
Printer.CurrentX = PLeft1(7) - (Printer.TextWidth(CheckStr(StrTemp7(7))) / 2)
Printer.CurrentY = iPrint
Printer.Print CheckStr(StrTemp7(7))
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle2
    Exit Sub
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "業務區"
For k = 2 To 15 Step 2
   If Len(Trim(StrTemp7((k) / 2))) <> 0 Then
      Printer.CurrentX = PLeft(k)
      Printer.CurrentY = iPrint
      Printer.Print "案數"
      Printer.CurrentX = PLeft(k + 1)
      Printer.CurrentY = iPrint
      Printer.Print "點數"
   End If
Next k
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
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
    PrintTitle
    PrintTitle2
    Exit Sub
End If
Printer.Font.Size = 10
End Sub

Sub PrintDatil2()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(1)
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(2)
For i = 2 To 15
   If Len(Trim(StrTemp7(Int((i) / 2)))) <> 0 Then
      Select Case i
      Case 3, 5, 7, 9, 11, 13, 15
            Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(strTemp(i + 1) & " ", "###,##0.0"))
            Printer.CurrentY = iPrint
            Printer.Print Format(strTemp(i + 1) & " ", "###,##0.0")
      Case Else
            Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(strTemp(i + 1))
            Printer.CurrentY = iPrint
            Printer.Print strTemp(i + 1)
      End Select
   End If
Next i
For i = 2 To 14 Step 2
    Printer.Line (PLeft(i) - 50, iPrint - 150)-(PLeft(i) - 50, iPrint + 450)
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft2()
Erase PLeft
Erase PLeft1
PLeft(0) = 0
PLeft(1) = 1000
PLeft(2) = 2200
For i = 3 To 15
    PLeft(i) = 2200 + ((i - 2) * 1200)
Next i
PLeft1(1) = PLeft(3)
PLeft1(2) = PLeft(5)
PLeft1(3) = PLeft(7)
PLeft1(4) = PLeft(9)
PLeft1(5) = PLeft(11)
PLeft1(6) = PLeft(13)
PLeft1(7) = PLeft(15)
PLeft1(8) = PLeft(17)
PLeft1(9) = PLeft(19)
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
    PrintTitle
    PrintTitle1
    Exit Sub
End If
Printer.CurrentX = PLeft1(1) - (Printer.TextWidth(CheckStr(StrTemp7(1))) / 2)
Printer.CurrentY = iPrint
Printer.Print CheckStr(StrTemp7(1))
Printer.CurrentX = PLeft1(2) - (Printer.TextWidth(CheckStr(StrTemp7(2))) / 2)
Printer.CurrentY = iPrint
Printer.Print CheckStr(StrTemp7(2))
Printer.CurrentX = PLeft1(3) - (Printer.TextWidth(CheckStr(StrTemp7(3))) / 2)
Printer.CurrentY = iPrint
Printer.Print CheckStr(StrTemp7(3))
Printer.CurrentX = PLeft1(4) - (Printer.TextWidth(CheckStr(StrTemp7(4))) / 2)
Printer.CurrentY = iPrint
Printer.Print CheckStr(StrTemp7(4))
Printer.CurrentX = PLeft1(5) - (Printer.TextWidth(CheckStr(StrTemp7(5))) / 2)
Printer.CurrentY = iPrint
Printer.Print CheckStr(StrTemp7(5))
Printer.CurrentX = PLeft1(6) - (Printer.TextWidth(CheckStr(StrTemp7(6))) / 2)
Printer.CurrentY = iPrint
Printer.Print CheckStr(StrTemp7(6))
Printer.CurrentX = PLeft1(7) - (Printer.TextWidth(CheckStr(StrTemp7(7))) / 2)
Printer.CurrentY = iPrint
Printer.Print CheckStr(StrTemp7(7))
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle1
    Exit Sub
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "業務區"
For k = 2 To 20 Step 3
   If Len(Trim(StrTemp7(Int((k + 1) / 3)))) <> 0 Then
      Printer.CurrentX = PLeft(k)
      Printer.CurrentY = iPrint
      Printer.Print "案數"
      Printer.CurrentX = PLeft(k + 1)
      Printer.CurrentY = iPrint
      Printer.Print "點數"
      Printer.CurrentX = PLeft(k + 2)
      Printer.CurrentY = iPrint
      Printer.Print "待辦"
   End If
Next k
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
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
    PrintTitle
    PrintTitle1
    Exit Sub
End If
Printer.Font.Size = 10
End Sub

Sub PrintDatil1()
'For I = 2 To 26 Step 3
'    Printer.Line (PLeft(I) - 50, iPrint - 150)-(PLeft(I) - 50, iPrint + 450)
'Next I
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(1)
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(2)
For i = 2 To 22
   If Len(Trim(StrTemp7(Int((i + 1) / 3)))) <> 0 Then
      Select Case i
      Case 3, 6, 9, 12, 15, 18, 21
           Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(strTemp(i + 1) & " ", "#####0.0"))
           Printer.CurrentY = iPrint
           Printer.Print Format(strTemp(i + 1) & " ", "#####0.0")
      Case Else
           Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(strTemp(i + 1))
           Printer.CurrentY = iPrint
           Printer.Print strTemp(i + 1)
      End Select
   End If
Next i
For i = 2 To 20 Step 3
    Printer.Line (PLeft(i) - 50, iPrint - 150)-(PLeft(i) - 50, iPrint + 450)
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft1()
Erase PLeft
Erase PLeft1
PLeft(0) = 0
PLeft(1) = 1000
'PLeft(2) = 2200
For i = 0 To 6
    PLeft(2 + (i * 3)) = 2200 + (i * 2200)
    PLeft(3 + (i * 3)) = 3000 + (i * 2200)
    PLeft(4 + (i * 3)) = 3700 + (i * 2200)
Next i
PLeft1(1) = PLeft(3) + 700 / 2
PLeft1(2) = PLeft(6) + 700 / 2
PLeft1(3) = PLeft(9) + 700 / 2
PLeft1(4) = PLeft(12) + 700 / 2
PLeft1(5) = PLeft(15) + 700 / 2
PLeft1(6) = PLeft(18) + 700 / 2
PLeft1(7) = PLeft(21) + 700 / 2

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
Set frm020407 = Nothing
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
Case 5
     If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
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
Case Else
End Select
End Sub
