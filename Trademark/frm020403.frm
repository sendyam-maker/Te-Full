VERSION 5.00
Begin VB.Form frm020403 
   BorderStyle     =   1  '單線固定
   Caption         =   "商爭案智權人員收/發文統計表"
   ClientHeight    =   2830
   ClientLeft      =   3990
   ClientTop       =   2000
   ClientWidth     =   3870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2830
   ScaleWidth      =   3870
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   12
      TabIndex        =   14
      Top             =   1905
      Width           =   3825
      Begin VB.ComboBox Combo1 
         Height          =   300
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
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1128
      TabIndex        =   0
      Top             =   570
      Width           =   1740
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1128
      MaxLength       =   1
      TabIndex        =   1
      Top             =   870
      Width           =   270
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1128
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1200
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   2148
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1200
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1128
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1530
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   2148
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1530
      Width           =   930
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2292
      TabIndex        =   7
      Top             =   12
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3084
      TabIndex        =   8
      Top             =   12
      Width           =   756
   End
   Begin VB.Line Line3 
      X1              =   1590
      X2              =   2850
      Y1              =   1665
      Y2              =   1665
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   195
      TabIndex        =   13
      Top             =   585
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "列印別："
      Height          =   180
      Index           =   1
      Left            =   195
      TabIndex        =   12
      Top             =   930
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "日期："
      Height          =   180
      Index           =   2
      Left            =   195
      TabIndex        =   11
      Top             =   1245
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   3
      Left            =   195
      TabIndex        =   10
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "(1.收文  2.發文)"
      Height          =   180
      Index           =   4
      Left            =   1440
      TabIndex        =   9
      Top             =   915
      Width           =   1410
   End
   Begin VB.Line Line1 
      X1              =   1125
      X2              =   2385
      Y1              =   1350
      Y2              =   1350
   End
   Begin VB.Line Line2 
      X1              =   1125
      X2              =   1875
      Y1              =   1635
      Y2              =   1635
   End
End
Attribute VB_Name = "frm020403"
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
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 30) As String, strTemp3 As String, TestOk As Boolean, StrTemp7(1 To 9) As String, k As Integer
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
'Add By Cheng 2003/09/03
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'Modify By Sindy 2015/1/22
Dim System_ID As String
Dim bolIsChina As Boolean
Dim Title_601 As String
Dim Title_603 As String
Dim Title_605 As String
Dim Title_401 As String
Dim Title_402 As String
Dim Title_403 As String
Dim Title_408 As String
Dim Title_602 As String
Dim Title_604 As String
Dim Title_606 As String
Dim Title_202 As String
Dim Title_406 As String
Dim Title_407 As String
'2015/1/22

Screen.MousePointer = vbHourglass
cnnConnection.Execute "delete from R020403 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "DELETE FROM R020403_1 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "DELETE FROM R020403_2 WHERE ID='" & strUserNum & "' "
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
Case 1 '收文
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
Case 2 '發文
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
If Len(txt1(4)) <> 0 Or Len(txt1(5)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(4) & "-" & txt1(5) 'Add By Sindy 2010/10/19
End If
'Add By Cheng 2004/04/29
'抓計件的資料
StrSQL6 = StrSQL6 & " And CP26 Is Null "
'End
CheckOC
'Modify By Cheng 2002/02/06
'取消抓承辦人ST05為"97"或"17"條件, 改為抓案件性質為"4"或"6"開頭CP09<"B"的資料
'注意：業務區是CASEPROGRESS的CP12
'strSQL = "SELECT S1.st03,cp13,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF S1,STAFF S2 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and (S2.ST05='97' OR S2.ST05='17') AND CP09<'C' " & strSQL1 + StrSQL6
'strSQL = strSQL + " union all select S1.st03,cp13,NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1,STAFF S2 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and (S2.ST05='97' OR S2.ST05='17') AND CP09<'C' " & strSQL2 + StrSQL6
'Modify By Cheng 2002/02/08 案件名稱一律顯示台灣名稱
'strSQL = "SELECT CP12,cp13,NVL(CPM03,CP10),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 " & _
'         " FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and (SubStr(CP10,1,1)='4' Or SubStr(CP10,1,1)='6') AND CP09<'B' " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' OR CP14 IS NULL ) ", " And (SubStr(ST03,1,2)='F1' OR CP14 IS NULL )") & strSQL1 + StrSQL6
'strSQL = strSQL + " union all select CP12,cp13,NVL(DECODE(CPM03,CP10),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 " & _
'         " FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF " & _
'         " WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and (SubStr(Cp10,1,1)='4' Or SubStr(Cp10,1,1)='6') AND CP09<'B' " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' OR CP14 IS NULL ) ", " And (SubStr(ST03,1,2)='F1' OR CP14 IS NULL )") & strSQL2 + StrSQL6
'Modify By Cheng 2002/04/04
'加抓案件性質為"202"或"204"或"205"或"207"

'Modify By Cheng 2002/04/11
'若台灣案件性質名稱為"(無)"則改抓大陸案件性質名稱
'strSQL = "SELECT CP12,cp13,NVL(CPM03,CP10),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 " & _
'         " FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and (SubStr(CP10,1,1)='4' Or SubStr(CP10,1,1)='6' OR CP10='202' OR CP10='204' OR CP10='205' ) AND CP09<'B' " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' OR CP14 IS NULL ) ", " And (SubStr(ST03,1,2)='F1' OR CP14 IS NULL )") & strSQL1 + StrSQL6
'strSQL = strSQL + " union all select CP12,cp13,NVL(CPM03,CP10),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 " & _
'         " FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF " & _
'         " WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and (SubStr(Cp10,1,1)='4' Or SubStr(Cp10,1,1)='6' OR CP10='202' OR CP10='204' OR CP10='205' ) AND CP09<'B' " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' OR CP14 IS NULL ) ", " And (SubStr(ST03,1,2)='F1' OR CP14 IS NULL )") & strSQL2 + StrSQL6
'**************** 將業務區改成抓案件進度檔   91.08.15  nick
'92.3.6 MODIFY BY SONIA 若從外商的統計報表進入, 不管承辦人部門別
'strSQL = "SELECT CP12,cp13,DECODE(NVL(CPM03,CP10),'（無）',CPM04,NVL(CPM03,CP10)),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 " & _
'         " FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and (SubStr(CP10,1,1)='4' Or SubStr(CP10,1,1)='6' OR CP10='202' OR CP10='204' OR CP10='205' OR CP10='207' ) AND CP09<'B' " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' OR CP14 IS NULL ) ", " And (SubStr(ST03,1,2)='F1' OR CP14 IS NULL )") & strSQL1 + StrSQL6
'strSQL = strSQL + " union all select CP12,cp13,DECODE(NVL(CPM03,CP10),'（無）',CPM04,NVL(CPM03,CP10)),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 " & _
'         " FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF " & _
'         " WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and (SubStr(Cp10,1,1)='4' Or SubStr(Cp10,1,1)='6' OR CP10='202' OR CP10='204' OR CP10='205' OR CP10='207' ) AND CP09<'B' " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' OR CP14 IS NULL ) ", " And (SubStr(ST03,1,2)='F1' OR CP14 IS NULL )") & strSQL2 + StrSQL6


'add by nickc 2007/07/26 商標處改大陸格式
Dim MyCPMStr1 As String
Dim MyCPMStr2 As String
If txt1(4) = "020" And txt1(5) = "020" Then
    MyCPMStr1 = "cpm04"
    MyCPMStr2 = "cpm04"
Else
    MyCPMStr1 = "DECODE(NVL(CPM03,CP10),'（無）',CPM04,NVL(CPM03,CP10))"
    MyCPMStr2 = "DECODE(NVL(CPM03,CP10),'（無）',CPM04,NVL(CPM03,CP10))"
End If
'strSQL = "SELECT CP12,cp13,DECODE(NVL(CPM03,CP10),'（無）',CPM04,NVL(CPM03,CP10)),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 " & _
'         " FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and (SubStr(CP10,1,1)='4' Or SubStr(CP10,1,1)='6' OR CP10='202' OR CP10='204' OR CP10='205' OR CP10='207' ) AND CP09<'B' " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' OR CP14 IS NULL ) ", "") & strSQL1 + StrSQL6
'strSQL = strSQL + " union all select CP12,cp13,DECODE(NVL(CPM03,CP10),'（無）',CPM04,NVL(CPM03,CP10)),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 " & _
'         " FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF " & _
'         " WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and (SubStr(Cp10,1,1)='4' Or SubStr(Cp10,1,1)='6' OR CP10='202' OR CP10='204' OR CP10='205' OR CP10='207' ) AND CP09<'B' " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' OR CP14 IS NULL ) ", "") & strSQL2 + StrSQL6
'2011/12/7 modify by sonia 加入OR CP10='210'
'Modify By Sindy 2021/2/23 改抓案件性質的語法加入TMdebate
'  " and (SubStr(CP10,1,1)='4' Or SubStr(CP10,1,1)='6' OR CP10='202' OR CP10='210' OR CP10='204' OR CP10='205' OR CP10='207')"
' ==> " and instr('" & TMdebate & ",204,205,207',cp10)>0"
'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
strSql = "SELECT CP12,cp13," & MyCPMStr1 & ",1,CP18,DECODE(CP27,NULL,1,'',1,0),DECODE(CP10,'210','202',CP10) CP10,cp09 " & _
         " FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND DECODE(CP10,'210','202',CP10)=CPM02(+)" & _
         " and instr('" & TMdebate & ",204,205,207',cp10)>0 and not(cp01='FCT' And InStr('" & FCT_NotTMdebate & "', cp10) > 0)" & _
         " AND CP09<'B' " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' OR CP14 IS NULL ) ", "") & strSQL1 + StrSQL6
strSql = strSql + " union all select CP12,cp13," & MyCPMStr2 & ",1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 " & _
         " FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF " & _
         " WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND DECODE(CP10,'210','202',CP10)=CPM02(+)" & _
         " and instr('" & TMdebate & ",204,205,207',cp10)>0 and not(cp01='FCT' And InStr('" & FCT_NotTMdebate & "', cp10) > 0)" & _
         " AND CP09<'B' " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' OR CP14 IS NULL ) ", "") & strSQL2 + StrSQL6

'92.3.6 END
'Add By Cheng 2003/07/01
'自請撤回(306)須依其相關總收文號的案件性質判斷是商爭案
'add by nickc 2007/07/26 商標處改大陸格式
'strSQL = strSQL & " Union All SELECT CP12,cp13,DECODE(NVL(CPM03,CP10),'（無）',CPM04,NVL(CPM03,CP10)),1,CP18,DECODE(CP27,NULL,1,'',1,0), '306', cp09 " & _
'         " FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF " & _
'         " WHERE CP09 In (Select CP43 From CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) And CP10='306' AND CP09<'B' " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' OR CP14 IS NULL ) ", "") & strSQL1 + StrSQL6 & " ) " & _
'         " And CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND '306'=CPM02(+) and (SubStr(Cp10,1,1)='4' Or SubStr(Cp10,1,1)='6' OR CP10='202' OR CP10='204' OR CP10='205' OR CP10='207' ) "
'strSQL = strSQL & " Union All SELECT CP12,cp13,DECODE(NVL(CPM03,CP10),'（無）',CPM04,NVL(CPM03,CP10)),1,CP18,DECODE(CP27,NULL,1,'',1,0), '306', cp09 " & _
'         " FROM CASEPROGRESS, Servicepractice, CASEPROPERTYMAP,STAFF " & _
'         " WHERE CP09 In (Select CP43 From CASEPROGRESS, Servicepractice, CASEPROPERTYMAP,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) And CP10='306' AND CP09<'B' " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' OR CP14 IS NULL ) ", "") & strSQL2 + StrSQL6 & " ) " & _
'         " And CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND '306'=CPM02(+) and (SubStr(Cp10,1,1)='4' Or SubStr(Cp10,1,1)='6' OR CP10='202' OR CP10='204' OR CP10='205' OR CP10='207' ) "
'Modify By Sindy 2021/2/23 改抓案件性質的語法加入TMdebate
'  " and (SubStr(Cp10,1,1)='4' Or SubStr(Cp10,1,1)='6' OR CP10='202' OR CP10='210' OR CP10='204' OR CP10='205' OR CP10='207') "
' ==> " and instr('" & TMdebate & ",204,205,207',cp10)>0"
'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
strSql = strSql & " Union All SELECT CP12,cp13," & MyCPMStr1 & ",1,CP18,DECODE(CP27,NULL,1,'',1,0), '306', cp09 " & _
         " FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF " & _
         " WHERE CP09 In (Select CP43 From CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) And CP10='306' AND CP09<'B' " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' OR CP14 IS NULL ) ", "") & strSQL1 + StrSQL6 & " ) " & _
         " And CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND '306'=CPM02(+)" & _
         " and instr('" & TMdebate & ",204,205,207',cp10)>0 and not(cp01='FCT' And InStr('" & FCT_NotTMdebate & "', cp10) > 0)"
strSql = strSql & " Union All SELECT CP12,cp13," & MyCPMStr2 & ",1,CP18,DECODE(CP27,NULL,1,'',1,0), '306', cp09 " & _
         " FROM CASEPROGRESS, Servicepractice, CASEPROPERTYMAP,STAFF " & _
         " WHERE CP09 In (Select CP43 From CASEPROGRESS, Servicepractice, CASEPROPERTYMAP,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) And CP10='306' AND CP09<'B' " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' OR CP14 IS NULL ) ", "") & strSQL2 + StrSQL6 & " ) " & _
         " And CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND '306'=CPM02(+)" & _
         " and instr('" & TMdebate & ",204,205,207',cp10)>0 and not(cp01='FCT' And InStr('" & FCT_NotTMdebate & "', cp10) > 0)"
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
Call ClsPDGetCaseProperty("T", "402", Title_402, bolIsChina, False)
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
        Do While .EOF = False
            For i = 0 To 5
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            Select Case Val(CheckStr(.Fields(6)))
            '固定的第一頁
            Case 601, 627 '異議,Add by Sindy 2019/8/15 +部分異議
                 strTemp(6) = "*"
                 strTemp(7) = "1"
                 strTemp(2) = Title_601 'Modify By Sindy 2015/1/22
'                 'edit by nickc 2007/07/26 商標處改大陸格式
'                 If txt1(4) = "020" And txt1(5) = "020" Then
'                    strTemp(2) = "異議"
'                 Else
'                    strTemp(2) = "異議"
'                 End If
            Case 603, 629 '評定,Add by Sindy 2019/8/15 +部分評定
                 strTemp(6) = "*"
                 strTemp(7) = "2"
                 strTemp(2) = Title_603 'Modify By Sindy 2015/1/22
'                 'edit by nickc 2007/07/26 商標處改大陸格式
'                 If txt1(4) = "020" And txt1(5) = "020" Then
'                    strTemp(2) = "裁定"
'                 Else
'                    strTemp(2) = "評定"
'                 End If
            Case 605, 623 '廢止,Add by Sindy 2019/8/15 +部分廢止
                 strTemp(6) = "*"
                 strTemp(7) = "3"
                 strTemp(2) = Title_605 'Modify By Sindy 2015/1/22
'                 'edit by nickc 2007/07/26 商標處改大陸格式
'                 If txt1(4) = "020" And txt1(5) = "020" Then
'                    strTemp(2) = "撤銷"
'                 Else
'                    strTemp(2) = "廢止"
'                 End If
            Case 401 '訴願
                 strTemp(6) = "*"
                 strTemp(2) = Title_401 'Modify By Sindy 2015/1/22
                 'edit by nickc 2007/07/26 商標處改大陸格式
                 If txt1(4) = "020" And txt1(5) = "020" Then
                    strTemp(7) = "5"
'                    strTemp(2) = "復審"
                 Else
                    strTemp(7) = "4"
'                    strTemp(2) = "訴願"
                 End If
            Case 402 '再訴願
                 strTemp(2) = Title_402 'Modify By Sindy 2015/1/22
                 'edit by nickc 2007/07/26 商標處改大陸格式
                 If txt1(4) = "020" And txt1(5) = "020" Then
                    strTemp(6) = ""
                    strTemp(7) = "14"
                 Else
                    strTemp(6) = "*"
                    strTemp(7) = "5"
'                    strTemp(2) = "再訴願"
                 End If
            Case 403 '行政訴訟
                 strTemp(6) = "*"
                 strTemp(7) = "6"
                 strTemp(2) = Title_403 'Modify By Sindy 2015/1/22
'                 'edit by nickc 2007/07/26 商標處改大陸格式
'                 If txt1(4) = "020" And txt1(5) = "020" Then
'                    strTemp(2) = "大陸上訴"
'                 Else
'                    strTemp(2) = "行政訴訟"
'                 End If
            'Modify By Cheng 2002/02/06
            '將再審之訴移至非固定, 並以行政訴訟上訴取代之(408)
'            Case 405 '再審之訴
'            Case 409 '行政訴訟上訴
            Case 408 '行政訴訟上訴
                 strTemp(2) = Title_408 'Modify By Sindy 2015/1/22
                 'edit by nickc 2007/07/26 商標處改大陸格式
                 If txt1(4) = "020" And txt1(5) = "020" Then
                    strTemp(6) = ""
                    strTemp(7) = "14"
                 Else
                    strTemp(6) = "*"
                    strTemp(7) = "7"
'                    strTemp(2) = "行政訴訟上訴"
                 End If
            '固定的第二頁
            Case 602, 628 '異議答辯,Add by Sindy 2019/8/15 +部分異議答辯
                 strTemp(6) = "*"
                 strTemp(7) = "8"
                 strTemp(2) = Title_602 'Modify By Sindy 2015/1/22
'                 'edit by nickc 2007/07/26 商標處改大陸格式
'                 If txt1(4) = "020" And txt1(5) = "020" Then
'                    strTemp(2) = "異議答辯"
'                 Else
'                    strTemp(2) = "異議答辯"
'                 End If
            Case 604, 630 '評定答辯,Add by Sindy 2019/8/15 +部分評定答辯
                 strTemp(6) = "*"
                 strTemp(7) = "9"
                 strTemp(2) = Title_604 'Modify By Sindy 2015/1/22
'                 'edit by nickc 2007/07/26 商標處改大陸格式
'                 If txt1(4) = "020" And txt1(5) = "020" Then
'                    strTemp(2) = "裁定答辯"
'                 Else
'                    strTemp(2) = "評定答辯"
'                 End If
            Case 606, 624 '廢止答辯,Add by Sindy 2019/8/15 +部分廢止答辯
                 strTemp(6) = "*"
                 strTemp(7) = "10"
                 strTemp(2) = Title_606 'Modify By Sindy 2015/1/22
'                 'edit by nickc 2007/07/26 商標處改大陸格式
'                 If txt1(4) = "020" And txt1(5) = "020" Then
'                    strTemp(2) = "撤銷答辯"
'                 Else
'                    strTemp(2) = "廢止答辯"
'                 End If
            Case 202 '申請意見書
                 strTemp(2) = Title_202 'Modify By Sindy 2015/1/22
                 'edit by nickc 2007/07/26 商標處改大陸格式
                 If txt1(4) = "020" And txt1(5) = "020" Then
                    strTemp(6) = ""
                    strTemp(7) = "14"
                 Else
                    strTemp(6) = "*"
                    strTemp(7) = "11"
'                    strTemp(2) = "申請意見書"
                 End If
            '補充理由(612)移至非固定, 改為參加訴願(406)
'            Case 612 '補充理由
'            Case 407 '參加訴願
            Case 406 '參加訴願                     '大陸405秀玲在 96/06/08 併入 406 了
                strTemp(6) = "*"
                strTemp(7) = "12"
                strTemp(2) = Title_406 'Modify By Sindy 2015/1/22
'                 'edit by nickc 2007/07/26 商標處改大陸格式
'                 If txt1(4) = "020" And txt1(5) = "020" Then
'                    strTemp(2) = "復審答辯"
'                 Else
'                    strTemp(2) = "參加訴願"
'                 End If
            '補充答辯(613)移至非固定, 改為參加訴訟(407)
'            Case 613 '補充答辯
'            Case 408 '參加訴訟
            Case 407 '參加訴訟
                 strTemp(2) = Title_407 'Modify By Sindy 2015/1/22
                 'edit by nickc 2007/07/26 商標處改大陸格式
                 If txt1(4) = "020" And txt1(5) = "020" Then
                    strTemp(6) = ""
                    strTemp(7) = "14"
                 Else
                    strTemp(6) = "*"
                    strTemp(7) = "13"
'                    strTemp(2) = "參加訴訟"
                 End If
'            'add by nickc 2007/07/26 商標處改大陸格式
'            Case 618
'                 strTemp(2) = Title_618 'Modify By Sindy 2015/1/22
'                 If txt1(4) = "020" And txt1(5) = "020" Then
'                    strTemp(6) = "*"
'                    strTemp(7) = "4"
''                    strTemp(2) = "註冊不當撤銷"
'                 Else
'                    strTemp(6) = ""
'                    strTemp(7) = "14"
'                 End If
'            'add by nickc 2007/07/26 商標處改大陸格式
'            Case 619
'                 strTemp(2) = Title_619 'Modify By Sindy 2015/1/22
'                 If txt1(4) = "020" And txt1(5) = "020" Then
'                    strTemp(6) = "*"
'                    strTemp(7) = "11"
''                    strTemp(2) = "註冊不當撤銷答辯"
'                 Else
'                    strTemp(6) = ""
'                    strTemp(7) = "14"
'                 End If
            Case Else '其他
                 strTemp(6) = ""
                 strTemp(7) = "14"
            End Select
            'Add By Cheng 2003/09/03
            '若案件性質為自請撤回(306)
            If "" & .Fields(6).Value = "306" Then
                '93.6.4 MODIFY BY SONIA
                'strSQLA = "Select * From CaseProgress Where CP43='" & .Fields(7).Value & "' "
                StrSQLa = "Select * From CaseProgress Where CP43='" & .Fields(7).Value & "' AND CP10='306'"
                '93.6.4 END
                rsA.CursorLocation = adUseClient
                rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                If rsA.RecordCount > 0 Then
                    '93.6.4 CANCEL BY SONIA 宋若蘭說自請撤回只抓相關總收文號之智權人員及業務區
                    'Modify By Sindy 2012/1/17 仍以原自請撤回智權人員或承辦人來統計
                    strTemp(0) = "" & rsA("CP12").Value
                    strTemp(1) = "" & rsA("CP13").Value
                    '93.6.4 END
                    strTemp(4) = "" & rsA("CP18").Value
                    strTemp(5) = IIf("" & rsA("CP27").Value = "", "1", "0")
                End If
                If rsA.State <> adStateClosed Then rsA.Close
                Set rsA = Nothing
            End If
            strSql = "INSERT INTO R020403 VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "'," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & ",'" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & strUserNum & "') "
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
'Modify By Cheng 2002/02/06
'cnnConnection.Execute "INSERT INTO R020403_1 VALUES (1,'','','異議','評定','廢止','訴願','再訴願','行政訴訟','再審之訴','" & strUserNum & "') "
'add by nick 2005/01/04 系統類別為 "S" 時，不印固定
If txt1(0) <> "S" Then
    'add by nickc 2007/07/26 商標處改大陸格式
    If txt1(4) = "020" And txt1(5) = "020" Then
'                                                                     '601    603    605    618            401    403
'        'cnnConnection.Execute "INSERT INTO R020403_1 VALUES (1,'','','異議','裁定','撤銷','註冊不當撤銷','復審','大陸上訴','  ','" & strUserNum & "') "
'        cnnConnection.Execute "INSERT INTO R020403_1 VALUES (1,'','','" & Title_601 & "','" & Title_603 & "','" & Title_605 & "','" & Title_618 & "','" & Title_401 & "','" & Title_403 & "','  ','" & strUserNum & "') "
'                                                                     '602        604        606        619                406
'        'cnnConnection.Execute "INSERT INTO R020403_1 VALUES (2,'','','異議答辯','裁定答辯','撤銷答辯','註冊不當撤銷答辯','復審答辯','  ','','" & strUserNum & "') "
'        cnnConnection.Execute "INSERT INTO R020403_1 VALUES (2,'','','" & Title_602 & "','" & Title_604 & "','" & Title_606 & "','" & Title_619 & "','" & Title_406 & "','  ','','" & strUserNum & "') "
        
        cnnConnection.Execute "INSERT INTO R020403_1 VALUES (1,'','','" & Title_601 & "','" & Title_603 & "','" & Title_605 & "','" & Title_401 & "','" & Title_403 & "','','','" & strUserNum & "') "
        cnnConnection.Execute "INSERT INTO R020403_1 VALUES (2,'','','" & Title_602 & "','" & Title_604 & "','" & Title_606 & "','" & Title_406 & "','','','','" & strUserNum & "') "
    Else
'                                                                     '601    603    605    401    402      403        408
'        'cnnConnection.Execute "INSERT INTO R020403_1 VALUES (1,'','','異議','評定','廢止','訴願','再訴願','行政訴訟','行政訴訟上訴','" & strUserNum & "') "
'        cnnConnection.Execute "INSERT INTO R020403_1 VALUES (1,'','','" & Title_601 & "','" & Title_603 & "','" & Title_605 & "','" & Title_401 & "','" & Title_402 & "','" & Title_403 & "','" & Title_408 & "','" & strUserNum & "') "
'                                                                     '602        604        606        202          406        407
'        'cnnConnection.Execute "INSERT INTO R020403_1 VALUES (2,'','','異議答辯','評定答辯','廢止答辯','申請意見書','參加訴願','參加訴訟','','" & strUserNum & "') "
'        cnnConnection.Execute "INSERT INTO R020403_1 VALUES (2,'','','" & Title_602 & "','" & Title_604 & "','" & Title_606 & "','" & Title_202 & "','" & Title_406 & "','" & Title_407 & "','','" & strUserNum & "') "
        
        cnnConnection.Execute "INSERT INTO R020403_1 VALUES (1,'','','" & Title_601 & "','" & Title_603 & "','" & Title_605 & "','" & Title_401 & "','" & Title_402 & "','" & Title_403 & "','" & Title_408 & "','" & strUserNum & "') "
        cnnConnection.Execute "INSERT INTO R020403_1 VALUES (2,'','','" & Title_602 & "','" & Title_604 & "','" & Title_606 & "','" & Title_202 & "','" & Title_406 & "','" & Title_407 & "','','" & strUserNum & "') "
    End If
End If
strSql = "select distinct r066003 from r020403 where id='" & strUserNum & "' and (r066007='' or r066007 is null)  "
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
            strSql = "insert into r020403_1 values (" & Val(strTemp3) & ",'" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & strUserNum & "') "
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
strSql = "SELECT * FROM R020403_1 WHERE ID='" & strUserNum & "' ORDER BY R067001 "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 3 To 10
                strSql = "SELECT * FROM R020403 WHERE ID='" & strUserNum & "' AND R066003='" & CheckStr(.Fields(i)) & "' "
                CheckOC2
                adoRecordset1.CursorLocation = adUseClient
                adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                    adoRecordset1.MoveFirst
                    Do While adoRecordset1.EOF = False
                        For j = 3 To 26
                            strTemp(j) = ""
                        Next j
                        strTemp(0) = CheckStr(.Fields(0))
                        strTemp(1) = CheckStr(adoRecordset1.Fields(0))
                        strTemp(2) = CheckStr(adoRecordset1.Fields(1))
                        Select Case i
                        Case 3
                             strTemp(3) = CheckStr(adoRecordset1.Fields(3))
                             strTemp(4) = CheckStr(adoRecordset1.Fields(4))
                             strTemp(5) = CheckStr(adoRecordset1.Fields(5))
                        Case 4
                             strTemp(6) = CheckStr(adoRecordset1.Fields(3))
                             strTemp(7) = CheckStr(adoRecordset1.Fields(4))
                             strTemp(8) = CheckStr(adoRecordset1.Fields(5))
                        Case 5
                             strTemp(9) = CheckStr(adoRecordset1.Fields(3))
                             strTemp(10) = CheckStr(adoRecordset1.Fields(4))
                             strTemp(11) = CheckStr(adoRecordset1.Fields(5))
                        Case 6
                             strTemp(12) = CheckStr(adoRecordset1.Fields(3))
                             strTemp(13) = CheckStr(adoRecordset1.Fields(4))
                             strTemp(14) = CheckStr(adoRecordset1.Fields(5))
                        Case 7
                             strTemp(15) = CheckStr(adoRecordset1.Fields(3))
                             strTemp(16) = CheckStr(adoRecordset1.Fields(4))
                             strTemp(17) = CheckStr(adoRecordset1.Fields(5))
                        Case 8
                             strTemp(18) = CheckStr(adoRecordset1.Fields(3))
                             strTemp(19) = CheckStr(adoRecordset1.Fields(4))
                             strTemp(20) = CheckStr(adoRecordset1.Fields(5))
                        Case 9
                             strTemp(21) = CheckStr(adoRecordset1.Fields(3))
                             strTemp(22) = CheckStr(adoRecordset1.Fields(4))
                             strTemp(23) = CheckStr(adoRecordset1.Fields(5))
                        Case Else
                        End Select
                        strTemp(24) = CheckStr(adoRecordset1.Fields(3))
                        strTemp(25) = CheckStr(adoRecordset1.Fields(4))
                        strTemp(26) = CheckStr(adoRecordset1.Fields(5))
                        strSql = "INSERT INTO R020403_2 VALUES (" & Val(strTemp(0)) & ",'" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "'," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & "," & Val(strTemp(15)) & "," & Val(strTemp(16)) & "," & Val(strTemp(17)) & "," & Val(strTemp(18)) & "," & Val(strTemp(19)) & "," & Val(strTemp(20)) & "," & Val(strTemp(21)) & "," & Val(strTemp(22)) & "," & Val(strTemp(23)) & "," & Val(strTemp(24)) & "," & Val(strTemp(25)) & "," & Val(strTemp(26)) & ",'" & strUserNum & "') "
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
strSql = "select max(r067001) from r020403_1 where id='" & strUserNum & "' "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    s = Val(CheckStr(adoRecordset.Fields(0)))
End If
CheckOC
strSql = "select distinct r066001,r066002 from r020403 where id='" & strUserNum & "' "
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 1 To s
                'add by nick 2005/01/04 系統類別為 "S" 時，不印固定
                'If txt1(0) <> "S" And i <> 1 And i <> 2 Then
                If txt1(0) <> "S" Then
                    strSql = "insert into r020403_2 values (" & i & ",'" & ChgSQL(CheckStr(.Fields(0))) & "','" & ChgSQL(CheckStr(.Fields(1))) & "',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "') "
                    cnnConnection.Execute strSql
                End If
            Next i
            .MoveNext
        Loop
    End If
End With
CheckOC
'重整
strSql = "select r068001,R068002,r068003,sum(r068004),sum(r068005),sum(r068006),sum(r068007),sum(r068008),sum(r068009),sum(r068010),sum(r068011),sum(r068012),sum(r068013),sum(r068014),sum(r068015),sum(r068016),sum(r068017),sum(r068018),sum(r068019),sum(r068020),sum(r068021),sum(r068022),sum(r068023),sum(r068024),sum(r068025),sum(r068026),sum(r068027) from r020403_2 where id='" & strUserNum & "' group by r068001,r068002,r068003 order by r068001,r068002,r068003 "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        cnnConnection.Execute "DELETE FROM R020403_2 WHERE ID='" & strUserNum & "' "
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 26
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strSql = "INSERT INTO R020403_2 VALUES (" & Val(strTemp(0)) & ",'" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "'," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & "," & Val(strTemp(15)) & "," & Val(strTemp(16)) & "," & Val(strTemp(17)) & "," & Val(strTemp(18)) & "," & Val(strTemp(19)) & "," & Val(strTemp(20)) & "," & Val(strTemp(21)) & "," & Val(strTemp(22)) & "," & Val(strTemp(23)) & "," & Val(strTemp(24)) & "," & Val(strTemp(25)) & "," & Val(strTemp(26)) & ",'" & strUserNum & "') "
            cnnConnection.Execute strSql
            .MoveNext
        Loop
    End If
End With
CheckOC
'抓固定合計欄位:
s = 7
'Modify By Sindy 2018/12/6 原Mark,放開
'edit by nickc 2007/07/26 固定第二頁的合計一律放在最右邊
strSql = "select r067001,r067004,r067005,r067006,r067007,r067008,r067009,r067010 from r020403_1 where ID='" & strUserNum & "' AND r067001=2 "
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
End If
CheckOC
'固定合計:
strSql = "select r068002,r068003,sum(r068025),sum(r068026),sum(r068027) from r020403_2 where id='" & strUserNum & "' and r068001<=2 group by r068002,r068003"
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    adoRecordset.MoveFirst
    Do While adoRecordset.EOF = False
        'strSql = "update r020403_2 set r068025=0,r068026=0,r068027=0 WHERE R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' AND ID='" & strUserNum & "' AND R068001=1 "
        'cnnConnection.Execute strSql
        Select Case s
        Case 1
             strSql = "UPDATE R020403_2 SET R068004=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068005=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068006=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' AND ID='" & strUserNum & "' AND R068001=2 "
             cnnConnection.Execute strSql
        Case 2
             strSql = "UPDATE R020403_2 SET R068007=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068008=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068009=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' AND ID='" & strUserNum & "' AND R068001=2 "
             cnnConnection.Execute strSql
        Case 3
             strSql = "UPDATE r020403_2 SET R068010=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068011=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068012=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' AND ID='" & strUserNum & "' AND R068001=2 "
             cnnConnection.Execute strSql
        Case 4
             strSql = "UPDATE r020403_2 SET R068013=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068014=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068015=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' AND ID='" & strUserNum & "' AND R068001=2 "
             cnnConnection.Execute strSql
        Case 5
             strSql = "UPDATE r020403_2 SET R068016=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068017=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068018=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' AND ID='" & strUserNum & "' AND R068001=2 "
             cnnConnection.Execute strSql
        Case 6
             strSql = "UPDATE r020403_2 SET R068019=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068020=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068021=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' AND ID='" & strUserNum & "' AND R068001=2 "
             cnnConnection.Execute strSql
        Case 7
             strSql = "UPDATE r020403_2 SET R068022=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068023=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068024=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' AND ID='" & strUserNum & "' AND R068001=2 "
             cnnConnection.Execute strSql
        Case Else
             strSql = "select r067001 from r020403_1 where r067001>2 order by r067001"
             CheckOC2
             With adoRecordset1
                  .CursorLocation = adUseClient
                  .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                  If .RecordCount <> 0 And .RecordCount > 0 Then
                       .MoveFirst
                       Do While .EOF = False
                            strSql = "update r020403_1 set r067001=" & Val(CheckStr(.Fields(0))) + 1 & " where r067001=" & Val(CheckStr(.Fields(0)))
                            cnnConnection.Execute strSql
                            .MoveNext
                       Loop
                  End If
             End With
             CheckOC2
             strSql = "select r068001 from r020403_2 where r068001>2 order by r068001"
             CheckOC2
             With adoRecordset1
                  .CursorLocation = adUseClient
                  .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                  If .RecordCount <> 0 And .RecordCount > 0 Then
                       .MoveFirst
                       Do While .EOF = False
                            strSql = "update r020403_2 set r068025=0,r068026=0,r068027=0,r068001=" & Val(CheckStr(.Fields(0))) + 1 & " where r068001=" & Val(CheckStr(.Fields(0)))
                            cnnConnection.Execute strSql
                            .MoveNext
                       Loop
                  End If
             End With
             CheckOC2
             strSql = "insert into r020403_2 values (3,'" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "','" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "'," & Val(CheckStr(adoRecordset.Fields(2))) & "," & Val(CheckStr(adoRecordset.Fields(3))) & "," & Val(CheckStr(adoRecordset.Fields(4))) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "') "
             cnnConnection.Execute strSql
             strSql = "INSERT INTO r020403_1 VALUES (3,'','','','','','','','','','" & strUserNum & "') "
             cnnConnection.Execute strSql
        End Select
        adoRecordset.MoveNext
    Loop
End If
CheckOC
'2018/12/6 END
'strSql = "select r068002,r068003,sum(r068025),sum(r068026),sum(r068027) from r020403_2 where id='" & strUserNum & "' and r068001<=2 group by r068002,r068003"
'CheckOC
'adoRecordset.CursorLocation = adUseClient
'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'    adoRecordset.MoveFirst
'    Do While adoRecordset.EOF = False
'        strSql = "update r020403_2 set r068025=0,r068026=0,r068027=0 WHERE R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' AND ID='" & strUserNum & "' AND R068001=1 "
'        cnnConnection.Execute strSql
'        Select Case s
'        Case 1
'             strSql = "UPDATE R020403_2 SET r068025=0,r068026=0,r068027=0,R068004=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068005=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068006=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' AND ID='" & strUserNum & "' AND R068001=2 "
'             cnnConnection.Execute strSql
'        Case 2
'             strSql = "UPDATE R020403_2 SET r068025=0,r068026=0,r068027=0,R068007=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068008=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068009=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' AND ID='" & strUserNum & "' AND R068001=2 "
'             cnnConnection.Execute strSql
'        Case 3
'             strSql = "UPDATE r020403_2 SET r068025=0,r068026=0,r068027=0,R068010=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068011=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068012=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' AND ID='" & strUserNum & "' AND R068001=2 "
'             cnnConnection.Execute strSql
'        Case 4
'             strSql = "UPDATE r020403_2 SET r068025=0,r068026=0,r068027=0,R068013=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068014=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068015=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' AND ID='" & strUserNum & "' AND R068001=2 "
'             cnnConnection.Execute strSql
'        Case 5
'             strSql = "UPDATE r020403_2 SET r068025=0,r068026=0,r068027=0,R068016=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068017=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068018=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' AND ID='" & strUserNum & "' AND R068001=2 "
'             cnnConnection.Execute strSql
'        Case 6
'             strSql = "UPDATE r020403_2 SET r068025=0,r068026=0,r068027=0,R068019=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068020=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068021=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' AND ID='" & strUserNum & "' AND R068001=2 "
'             cnnConnection.Execute strSql
'        Case 7
'             strSql = "UPDATE r020403_2 SET r068025=0,r068026=0,r068027=0,R068022=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068023=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068024=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' AND ID='" & strUserNum & "' AND R068001=2 "
'             cnnConnection.Execute strSql
'        Case Else
'             strSql = "select r067001 from r020403_1 where r067001>2 order by r067001"
'             CheckOC2
'             With adoRecordset1
'                  .CursorLocation = adUseClient
'                  .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                  If .RecordCount <> 0 And .RecordCount > 0 Then
'                       .MoveFirst
'                       Do While .EOF = False
'                            strSql = "update r020403_1 set r067001=" & Val(CheckStr(.Fields(0))) + 1 & " where r067001=" & Val(CheckStr(.Fields(0)))
'                            cnnConnection.Execute strSql
'                            .MoveNext
'                       Loop
'                  End If
'             End With
'             CheckOC2
'             strSql = "select r068001 from r020403_2 where r068001>2 order by r068001"
'             CheckOC2
'             With adoRecordset1
'                  .CursorLocation = adUseClient
'                  .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                  If .RecordCount <> 0 And .RecordCount > 0 Then
'                       .MoveFirst
'                       Do While .EOF = False
'                            strSql = "update r020403_2 set r068025=0,r068026=0,r068027=0,r068001=" & Val(CheckStr(.Fields(0))) + 1 & " where r068001=" & Val(CheckStr(.Fields(0)))
'                            cnnConnection.Execute strSql
'                            .MoveNext
'                       Loop
'                  End If
'             End With
'             CheckOC2
'             strSql = "insert into r020403_2 values (3,'" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "','" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "'," & Val(CheckStr(adoRecordset.Fields(2))) & "," & Val(CheckStr(adoRecordset.Fields(3))) & "," & Val(CheckStr(adoRecordset.Fields(4))) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "') "
'             cnnConnection.Execute strSql
'             strSql = "INSERT INTO r020403_1 VALUES (3,'','','','','','','','','','" & strUserNum & "') "
'             cnnConnection.Execute strSql
'        End Select
'        adoRecordset.MoveNext
'    Loop
'End If
'CheckOC

'911128 nick 下面區塊修正，以免合計計算錯誤
'***** start
'抓非固定合計欄位:
s = 8
strSql = "select r067001,r067004,r067005,r067006,r067007,r067008,r067009,r067010 from r020403_1 where ID='" & strUserNum & "' AND r067001 in (select max(r067001) from r020403_1 where id='" & strUserNum & "') and r067001 >2"
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
Else
    s = 0
End If
CheckOC
If s <> 0 Then
    strSql = "select max(r067001) from r020403_1 where id='" & strUserNum & "' "
    CheckOC2
    adoRecordset1.CursorLocation = adUseClient
    adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
         intS = Val(CheckStr(adoRecordset1.Fields(0)))
    End If
    '固定與非固定的總計:
    'strSql = "select r068002,r068003,sum(r068004)+sum(r068007)+sum(r068010)+sum(r068013)+sum(r068016)+sum(r068019)+sum(decode(R068001,2,0,r068022)),sum(r068005)+sum(r068008)+sum(r068011)+sum(r068014)+sum(r068017)+sum(r068020)+sum(decode(r068001,2,0,r068023)),sum(r068006)+sum(r068009)+sum(r068012)+sum(r068015)+sum(r068018)+sum(r068021)+sum(decode(r068001,2,0,r068024)) from r020403_2 where id='" & strUserNum & "' group by r068002,r068003"
    strSql = "select r068002,r068003,sum(r068025),sum(r068026),sum(r068027) from r020403_2 where id='" & strUserNum & "' group by r068002,r068003"
    CheckOC
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
        adoRecordset.MoveFirst
        '非固定合計:
        If s > 7 Then
            intS = intS + 1
        End If
        Do While adoRecordset.EOF = False
            strSql = "SELECT R068025,R068026,R068027 FROM R020403_2 WHERE ID='" & strUserNum & "' AND R068001 IN (SELECT MAX(R067001) FROM R020403_1 WHERE ID='" & strUserNum & "') " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " and r068002 is null ", " AND R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r068003 is null ", " AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ")
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
               Select Case s
               Case 1
                     strSql = "UPDATE R020403_2 SET R068004=" & Val(CheckStr(adoRecordset1.Fields(0))) & ",R068005=" & Val(CheckStr(adoRecordset1.Fields(1))) & ",R068006=" & Val(CheckStr(adoRecordset1.Fields(2))) & ",R068007=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068008=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068009=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r068002 is null ", " R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r068003 is null ", " AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND R068001 IN (SELECT MAX(R068001) FROM R020403_2 WHERE ID='" & strUserNum & "') "
                     cnnConnection.Execute strSql
               Case 2
                     strSql = "UPDATE r020403_2 SET R068007=" & Val(CheckStr(adoRecordset1.Fields(0))) & ",R068008=" & Val(CheckStr(adoRecordset1.Fields(1))) & ",R068009=" & Val(CheckStr(adoRecordset1.Fields(2))) & ",R068010=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068011=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068012=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r068002 is null ", " R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r068003 is null ", " AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND R068001 IN (SELECT MAX(R068001) FROM r020403_2 WHERE ID='" & strUserNum & "') "
                     cnnConnection.Execute strSql
               Case 3
                     strSql = "UPDATE r020403_2 SET R068010=" & Val(CheckStr(adoRecordset1.Fields(0))) & ",R068011=" & Val(CheckStr(adoRecordset1.Fields(1))) & ",R068012=" & Val(CheckStr(adoRecordset1.Fields(2))) & ",R068013=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068014=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068015=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r068002 is null ", " R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r068003 is null ", " AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND R068001 IN (SELECT MAX(R068001) FROM r020403_2 WHERE ID='" & strUserNum & "') "
                     cnnConnection.Execute strSql
               Case 4
                     strSql = "UPDATE r020403_2 SET R068013=" & Val(CheckStr(adoRecordset1.Fields(0))) & ",R068014=" & Val(CheckStr(adoRecordset1.Fields(1))) & ",R068015=" & Val(CheckStr(adoRecordset1.Fields(2))) & ",R068016=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068017=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068018=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r068002 is null ", " R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r068003 is null ", " AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND R068001 IN (SELECT MAX(R068001) FROM r020403_2 WHERE ID='" & strUserNum & "') "
                     cnnConnection.Execute strSql
               Case 5
                     strSql = "UPDATE r020403_2 SET R068016=" & Val(CheckStr(adoRecordset1.Fields(0))) & ",R068017=" & Val(CheckStr(adoRecordset1.Fields(1))) & ",R068018=" & Val(CheckStr(adoRecordset1.Fields(2))) & ",R068019=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068020=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068021=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r068002 is null ", " R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r068003 is null ", " AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND R068001 IN (SELECT MAX(R068001) FROM r020403_2 WHERE ID='" & strUserNum & "') "
                     cnnConnection.Execute strSql
               Case 6
                     strSql = "UPDATE r020403_2 SET R068019=" & Val(CheckStr(adoRecordset1.Fields(0))) & ",R068020=" & Val(CheckStr(adoRecordset1.Fields(1))) & ",R068021=" & Val(CheckStr(adoRecordset1.Fields(2))) & ",R068022=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068023=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068024=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r068002 is null ", " R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r068003 is null ", " AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND R068001 IN (SELECT MAX(R068001) FROM r020403_2 WHERE ID='" & strUserNum & "') "
                     cnnConnection.Execute strSql
               Case 7
                     strSql = "UPDATE r020403_2 SET R068022=" & Val(CheckStr(adoRecordset1.Fields(0))) & ",R068023=" & Val(CheckStr(adoRecordset1.Fields(1))) & ",R068024=" & Val(CheckStr(adoRecordset1.Fields(2))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r068002 is null ", " R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r068003 is null ", " AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND R068001 IN (SELECT MAX(R068001) FROM r020403_2 WHERE ID='" & strUserNum & "') "
                     cnnConnection.Execute strSql
               Case Else
                     strSql = "insert into r020403_2 values (" & intS & ",'" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "','" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "'," & Val(CheckStr(adoRecordset1.Fields(0))) & "," & Val(CheckStr(adoRecordset1.Fields(1))) & "," & Val(CheckStr(adoRecordset1.Fields(2))) & "," & Val(CheckStr(adoRecordset.Fields(2))) & "," & Val(CheckStr(adoRecordset.Fields(3))) & "," & Val(CheckStr(adoRecordset.Fields(4))) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "') "
                     cnnConnection.Execute strSql
               End Select
            End If
            CheckOC2
            adoRecordset.MoveNext
        Loop
        If s > 7 Then
            strSql = "INSERT INTO r020403_1 VALUES (" & intS & ",'','','合計','','','','','','','" & strUserNum & "') "
            cnnConnection.Execute strSql
        End If
    End If
    CheckOC
End If
'If s <> 0 Then
'    strSql = "select max(r067001) from r020403_1 where id='" & strUserNum & "' "
'    CheckOC2
'    adoRecordset1.CursorLocation = adUseClient
'    adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'    If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'         intS = Val(CheckStr(adoRecordset1.Fields(0)))
'    End If
'    '固定與非固定的總計
'    'strSql = "select r068002,r068003,sum(r068004)+sum(r068007)+sum(r068010)+sum(r068013)+sum(r068016)+sum(r068019)+sum(decode(R068001,2,0,r068022)),sum(r068005)+sum(r068008)+sum(r068011)+sum(r068014)+sum(r068017)+sum(r068020)+sum(decode(r068001,2,0,r068023)),sum(r068006)+sum(r068009)+sum(r068012)+sum(r068015)+sum(r068018)+sum(r068021)+sum(decode(r068001,2,0,r068024)) from r020403_2 where id='" & strUserNum & "' group by r068002,r068003"
'    strSql = "select r068002,r068003,sum(r068025),sum(r068026),sum(r068027) from r020403_2 where id='" & strUserNum & "' group by r068002,r068003"
'    CheckOC
'    adoRecordset.CursorLocation = adUseClient
'    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'    If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'        adoRecordset.MoveFirst
'        '非固定合計及(固定與非固定)的總計:
'        Do While adoRecordset.EOF = False
'            Select Case s
'            Case 1
'                 strSql = "SELECT R068025,R068026,R068027 FROM R020403_2 WHERE ID='" & strUserNum & "' AND R068001 IN (SELECT MAX(R068001) FROM R020403_2 WHERE ID='" & strUserNum & "') " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " and r068002 is null ", " AND R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r068003 is null ", " AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ")
'                 CheckOC2
'                 adoRecordset1.CursorLocation = adUseClient
'                 adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                 If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'                    strSql = "UPDATE R020403_2 SET r068025=0,r068026=0,r068027=0,R068004=" & Val(CheckStr(adoRecordset1.Fields(0))) & ",R068005=" & Val(CheckStr(adoRecordset1.Fields(1))) & ",R068006=" & Val(CheckStr(adoRecordset1.Fields(2))) & ",R068007=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068008=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068009=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r068002 is null ", " R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r068003 is null ", " AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND R068001 IN (SELECT MAX(R068001) FROM R020403_2 WHERE ID='" & strUserNum & "') "
'                    cnnConnection.Execute strSql
'                 End If
'                 CheckOC2
'            Case 2
'                 strSql = "SELECT r068025,r068026,r068027 FROM r020403_2 WHERE ID='" & strUserNum & "' AND R068001 IN (SELECT MAX(R068001) FROM r020403_2 WHERE ID='" & strUserNum & "') " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " and r068002 is null ", " AND R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r068003 is null ", " AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ")
'                 CheckOC2
'                 adoRecordset1.CursorLocation = adUseClient
'                 adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                 If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'                    strSql = "UPDATE r020403_2 SET r068025=0,r068026=0,r068027=0,R068007=" & Val(CheckStr(adoRecordset1.Fields(0))) & ",R068008=" & Val(CheckStr(adoRecordset1.Fields(1))) & ",R068009=" & Val(CheckStr(adoRecordset1.Fields(2))) & ",R068010=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068011=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068012=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r068002 is null ", " R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r068003 is null ", " AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND R068001 IN (SELECT MAX(R068001) FROM r020403_2 WHERE ID='" & strUserNum & "') "
'                    cnnConnection.Execute strSql
'                 End If
'                 CheckOC2
'            Case 3
'                 strSql = "SELECT r068025,r068026,r068027 FROM r020403_2 WHERE ID='" & strUserNum & "' AND R068001 IN (SELECT MAX(R068001) FROM r020403_2 WHERE ID='" & strUserNum & "') " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " and r068002 is null ", " AND R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r068003 is null ", " AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ")
'                 CheckOC2
'                 adoRecordset1.CursorLocation = adUseClient
'                 adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                 If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'                    strSql = "UPDATE r020403_2 SET r068025=0,r068026=0,r068027=0,R068010=" & Val(CheckStr(adoRecordset1.Fields(0))) & ",R068011=" & Val(CheckStr(adoRecordset1.Fields(1))) & ",R068012=" & Val(CheckStr(adoRecordset1.Fields(2))) & ",R068013=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068014=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068015=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r068002 is null ", " R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r068003 is null ", " AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND R068001 IN (SELECT MAX(R068001) FROM r020403_2 WHERE ID='" & strUserNum & "') "
'                    cnnConnection.Execute strSql
'                 End If
'                 CheckOC2
'            Case 4
'                 strSql = "SELECT r068025,r068026,r068027 FROM r020403_2 WHERE ID='" & strUserNum & "' AND R068001 IN (SELECT MAX(R068001) FROM r020403_2 WHERE ID='" & strUserNum & "') " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " and r068002 is null ", " AND R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r068003 is null ", " AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ")
'                 CheckOC2
'                 adoRecordset1.CursorLocation = adUseClient
'                 adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                 If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'                    strSql = "UPDATE r020403_2 SET r068025=0,r068026=0,r068027=0,R068013=" & Val(CheckStr(adoRecordset1.Fields(0))) & ",R068014=" & Val(CheckStr(adoRecordset1.Fields(1))) & ",R068015=" & Val(CheckStr(adoRecordset1.Fields(2))) & ",R068016=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068017=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068018=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r068002 is null ", " R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r068003 is null ", " AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND R068001 IN (SELECT MAX(R068001) FROM r020403_2 WHERE ID='" & strUserNum & "') "
'                   cnnConnection.Execute strSql
'                 End If
'                 CheckOC2
'            Case 5
'                 strSql = "SELECT r068025,r068026,r068027 FROM r020403_2 WHERE ID='" & strUserNum & "' AND R068001 IN (SELECT MAX(R068001) FROM r020403_2 WHERE ID='" & strUserNum & "') " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " and r068002 is null ", " AND R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r068003 is null ", " AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ")
'                 CheckOC2
'                 adoRecordset1.CursorLocation = adUseClient
'                 adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                 If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'                    'strSql = "UPDATE r020403_2 SET r068025=0,r068026=0,r068027=0,R068016=" & Val(CheckStr(adoRecordset1.Fields(0))) & ",R068017=" & Val(CheckStr(adoRecordset1.Fields(1))) & ",R068018=" & Val(CheckStr(adoRecordset1.Fields(2))) & ",R068019=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068020=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068021=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r068002 is null ", " R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r068003 is null ", " AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND R068001 IN (SELECT MAX(R068001) FROM r020403_2 WHERE ID='" & strUserNum & "') "
'                    strSql = "UPDATE r020403_2 SET R068016=" & Val(CheckStr(adoRecordset1.Fields(0))) & ",R068017=" & Val(CheckStr(adoRecordset1.Fields(1))) & ",R068018=" & Val(CheckStr(adoRecordset1.Fields(2))) & ",R068019=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068020=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068021=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r068002 is null ", " R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r068003 is null ", " AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND R068001 IN (SELECT MAX(R068001) FROM r020403_2 WHERE ID='" & strUserNum & "') "
'                    cnnConnection.Execute strSql
'                 End If
'                 CheckOC2
'            Case 6
'                 strSql = "SELECT r068025,r068026,r068027 FROM r020403_2 WHERE ID='" & strUserNum & "' AND R068001 IN (SELECT MAX(R068001) FROM r020403_2 WHERE ID='" & strUserNum & "') " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " and r068002 is null ", " AND R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r068003 is null ", " AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ")
'                 CheckOC2
'                 adoRecordset1.CursorLocation = adUseClient
'                adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                 If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'                    strSql = "UPDATE r020403_2 SET r068025=0,r068026=0,r068027=0,R068019=" & Val(CheckStr(adoRecordset1.Fields(0))) & ",R068020=" & Val(CheckStr(adoRecordset1.Fields(1))) & ",R068021=" & Val(CheckStr(adoRecordset1.Fields(2))) & ",R068022=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R068023=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R068024=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r068002 is null ", " R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r068003 is null ", " AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND R068001 IN (SELECT MAX(R068001) FROM r020403_2 WHERE ID='" & strUserNum & "') "
'                    cnnConnection.Execute strSql
'                 End If
'                 CheckOC2
'            Case 7
'                 strSql = "SELECT r068025,r068026,r068027 FROM r020403_2 WHERE ID='" & strUserNum & "' AND R068001 IN (SELECT MAX(R068001) FROM r020403_2 WHERE ID='" & strUserNum & "') " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " and r068002 is null ", " AND R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r068003 is null ", " AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ")
'                 CheckOC2
'                 adoRecordset1.CursorLocation = adUseClient
'                 adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                 If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'                    strSql = "UPDATE r020403_2 SET r068025=0,r068026=0,r068027=0,R068022=" & Val(CheckStr(adoRecordset1.Fields(0))) & ",R068023=" & Val(CheckStr(adoRecordset1.Fields(1))) & ",R068024=" & Val(CheckStr(adoRecordset1.Fields(2))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r068002 is null ", " R068002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r068003 is null ", " AND R068003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND R068001 IN (SELECT MAX(R068001) FROM r020403_2 WHERE ID='" & strUserNum & "') "
'                    cnnConnection.Execute strSql
'                 End If
'                 CheckOC2
'            Case Else
'                   strSql = "insert into r020403_2 values (" & intS + 1 & ",'" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "','" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "'," & Val(CheckStr(adoRecordset.Fields(2))) & "," & Val(CheckStr(adoRecordset.Fields(3))) & "," & Val(CheckStr(adoRecordset.Fields(4))) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "') "
'                   cnnConnection.Execute strSql
'                   strSql = "INSERT INTO r020403_1 VALUES (" & intS + 1 & ",'','','','','','','','','','" & strUserNum & "') "
'                   cnnConnection.Execute strSql
'            End Select
'            adoRecordset.MoveNext
'        Loop
'    End If
'    CheckOC
'End If
strSql = "update r020403_2 set r068025=0,r068026=0,r068027=0 WHERE ID='" & strUserNum & "'"
cnnConnection.Execute strSql
'***** end
'add by nick 2005/01/04 系統類別為 "S" 時，不印固定
If txt1(0) <> "S" Then
   'add by nickc 2007/07/26 商標處改大陸格式
   If txt1(4) = "020" And txt1(5) = "020" Then
'      'cnnConnection.Execute "UPDATE R020403_1 SET R067004='異議',R067005='裁定',R067006='撤銷',R067007='註冊不當撤銷',R067008='復審',R067009='大陸上訴',R067010='   ' WHERE ID='" & strUserNum & "' AND R067001=1 "
'      cnnConnection.Execute "UPDATE R020403_1 SET R067004='" & Title_601 & "',R067005='" & Title_603 & "',R067006='" & Title_605 & "',R067007='" & Title_618 & "',R067008='" & Title_401 & "',R067009='" & Title_403 & "',R067010='   ' WHERE ID='" & strUserNum & "' AND R067001=1 "
'      'cnnConnection.Execute "UPDATE R020403_1 SET R067004='異議答辯',R067005='裁定答辯',R067006='撤銷答辯',R067007='註冊不當撤銷答辯',R067008='復審答辯',R067009='   ',R067010='' WHERE ID='" & strUserNum & "' AND R067001=2 "
'      cnnConnection.Execute "UPDATE R020403_1 SET R067004='" & Title_602 & "',R067005='" & Title_604 & "',R067006='" & Title_606 & "',R067007='" & Title_619 & "',R067008='" & Title_406 & "',R067009='   ',R067010='' WHERE ID='" & strUserNum & "' AND R067001=2 "
      cnnConnection.Execute "UPDATE R020403_1 SET R067004='" & Title_601 & "',R067005='" & Title_603 & "',R067006='" & Title_605 & "',R067007='" & Title_401 & "',R067008='" & Title_403 & "',R067009='',R067010='' WHERE ID='" & strUserNum & "' AND R067001=1 "
      cnnConnection.Execute "UPDATE R020403_1 SET R067004='" & Title_602 & "',R067005='" & Title_604 & "',R067006='" & Title_606 & "',R067007='" & Title_406 & "',R067008='',R067009='',R067010='' WHERE ID='" & strUserNum & "' AND R067001=2 "
   Else
      'cnnConnection.Execute "UPDATE R020403_1 SET R067004='異議',R067005='評定',R067006='廢止',R067007='訴願',R067008='再訴願',R067009='行政訴訟',R067010='行政訴訟上訴' WHERE ID='" & strUserNum & "' AND R067001=1 "
      cnnConnection.Execute "UPDATE R020403_1 SET R067004='" & Title_601 & "',R067005='" & Title_603 & "',R067006='" & Title_605 & "',R067007='" & Title_401 & "',R067008='" & Title_402 & "',R067009='" & Title_403 & "',R067010='" & Title_408 & "' WHERE ID='" & strUserNum & "' AND R067001=1 "
      'cnnConnection.Execute "UPDATE R020403_1 SET R067004='異議答辯',R067005='評定答辯',R067006='廢止答辯',R067007='申請意見書',R067008='參加訴願',R067009='參加訴訟',R067010='' WHERE ID='" & strUserNum & "' AND R067001=2 "
      cnnConnection.Execute "UPDATE R020403_1 SET R067004='" & Title_602 & "',R067005='" & Title_604 & "',R067006='" & Title_606 & "',R067007='" & Title_202 & "',R067008='" & Title_406 & "',R067009='" & Title_407 & "',R067010='' WHERE ID='" & strUserNum & "' AND R067001=2 "
   End If
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
strSQL_1 = "SELECT R068001,NVL(A0902,A0903),nvl(st02,R068003),R068004+R068005+R068006+R068007+R068008+R068009+R068010+R068011+R068012+R068013+R068014+R068015+R068016+R068017+R068018+R068019+R068020+R068021+R068022+R068023+R068024+R068025+R068026+R068027 As Total ,R068002,r068003 " & _
            " FROM R020403_2,ACC090,staff " & _
            " WHERE ID='" & strUserNum & "' and r068003=st01(+) AND r068002=a0901(+) And R068001 = '1' ORDER BY R068001,R068002,R068003 "
If rstmp_1.State <> adStateClosed Then rstmp_1.Close
rstmp_1.CursorLocation = adUseClient
rstmp_1.Open strSQL_1, cnnConnection, adOpenStatic, adLockReadOnly
If rstmp_1.RecordCount > 0 Then
   rstmp_1.MoveFirst
   While Not rstmp_1.EOF
      '搜尋固定區的第二頁
      strSQL_2 = "SELECT R068001,NVL(A0902,A0903),nvl(st02,R068003),R068004+R068005+R068006+R068007+R068008+R068009+R068010+R068011+R068012+R068013+R068014+R068015+R068016+R068017+R068018+R068019+R068020+R068021+R068022+R068023+R068024+R068025+R068026+R068027 As Total ,R068002,r068003 " & _
                  " FROM R020403_2,ACC090,staff " & _
                  " WHERE ID='" & strUserNum & "' and r068003=st01(+) AND r068002=a0901(+) And R068001 = '2' And R068002='" & rstmp_1("R068002").Value & "' And R068003='" & rstmp_1("R068003").Value & "' ORDER BY R068001,R068002,R068003 "
      If rstmp_2.State <> adStateClosed Then rstmp_2.Close
      rstmp_2.CursorLocation = adUseClient
      rstmp_2.Open strSQL_2, cnnConnection, adOpenStatic, adLockReadOnly
      If rstmp_2.RecordCount > 0 Then
         rstmp_2.MoveFirst
         '若固定區合計為0
         If rstmp_1("Total") + rstmp_2("Total") = 0 Then
            cnnConnection.Execute "Delete From R020403_2 Where ID ='" & strUserNum & "' And R068002='" & rstmp_1("R068002").Value & "' And R068003='" & rstmp_1("R068003").Value & "' And (R068001='1' Or R068001='2') "
            '搜尋固定區的第三頁以上
            strSQL_3 = "SELECT R068001,NVL(A0902,A0903),nvl(st02,R068003),R068004+R068005+R068006+R068007+R068008+R068009+R068010+R068011+R068012+R068013+R068014+R068015+R068016+R068017+R068018+R068019+R068020+R068021+R068022+R068023+R068024+R068025+R068026+R068027 As Total ,R068002,r068003 " & _
                        " FROM R020403_2,ACC090,staff " & _
                        " WHERE ID='" & strUserNum & "' and r068003=st01(+) AND r068002=a0901(+) And (R068001 <> '1' AND R068001 <> '2') And R068002='" & rstmp_1("R068002").Value & "' And R068003='" & rstmp_1("R068003").Value & "' ORDER BY R068001,R068002,R068003 "
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
                  cnnConnection.Execute "Delete From R020403_2 Where ID ='" & strUserNum & "' And R068002='" & rstmp_1("R068002").Value & "' And R068003='" & rstmp_1("R068003").Value & "' And (R068001<>'1' And R068001<>'2') "
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

strSql = "SELECT R068001,NVL(A0902,A0903),nvl(st02,R068003),R068004,R068005,R068007,R068008,R068010,R068011,R068013,R068014,R068016,R068017,R068019,R068020,R068022,R068023,R068025,R068026,R068002,r068003 FROM R020403_2,staff,ACC090 WHERE ID='" & strUserNum & "' and r068003=st01(+) AND r068002=A0901(+) ORDER BY R068001,R068002,R068003 "
CheckOC
Page = 1
TestOk = True
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        SavDay1 = CheckStr(.Fields(0)) '1,2,3 頁
        SavDay2 = CheckStr(.Fields(1)) '部門名稱
        SavDay3 = CheckStr(.Fields(19)) '部門代碼
        strSql = "SELECT R067001,R067004,R067005,R067006,R067007,R067008,R067009,R067010 FROM R020403_1 WHERE ID='" & strUserNum & "' AND R067001=" & Val(SavDay1) & " "
        CheckOC2
        adoRecordset1.CursorLocation = adUseClient
        adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
            For i = 1 To 7
                StrTemp7(i) = StrToStr(CheckStr(adoRecordset1.Fields(i)), 7)
            Next i
            'Modify By Cheng 2002/02/06
'            StrTemp7(7) = ""
            For i = 1 To 7
                If Len("" & StrTemp7(i)) = 0 And Val(SavDay1) >= 2 Then
                     If i <> 1 Then
                            If Page > 1 Then
                               StrTemp7(i) = "合計"
                            Else
                               StrTemp7(7) = "合計"
                            End If
                     End If
                     If i <> 7 Then
                         If i = 1 Then
                            If Page > 1 Then
                                 StrTemp7(i) = "總計"
                            End If
                         Else
                            If Page > 1 Then
                                 StrTemp7(i + 1) = "總計"
                            End If
                         End If
                     End If
                     Exit For
                End If
            Next i
        End If
        CheckOC2
        PrintTitle
        PrintTitle2
        SavDay1 = " "
        SavDay2 = " "
        SavDay3 = " "
        Do While .EOF = False
            For i = 0 To 19
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(1) = StrToStr(strTemp(1), 5)
            strTemp(2) = StrToStr(strTemp(2), 3)
            If Len(Trim(SavDay1)) > 0 Or Len(Trim(SavDay2)) > 0 Or Len(Trim(SavDay3)) > 0 Then
            If Val(strTemp(0)) <> Val(SavDay1) Then
                ShowLine2
                PrintEnd2 (0)
                ShowLine2
                PrintEnd2 (1)
                If Val(SavDay1) = 1 Then
                    ShowLine2
                    PrintEnd2 (2)
                    ShowLine2
                Else
                    ShowLine2
                    PrintEnd2 (3)
                    ShowLine2
                End If
                strSql = "SELECT R067001,R067004,R067005,R067006,R067007,R067008,R067009,R067010 FROM R020403_1 WHERE ID='" & strUserNum & "' AND R067001=" & Val(CheckStr(.Fields(0))) & " "
                CheckOC2
                adoRecordset1.CursorLocation = adUseClient
                adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                    For i = 1 To 7
                        StrTemp7(i) = StrToStr(CheckStr(adoRecordset1.Fields(i)), 7)
                    Next i
                    'Modify By Sindy 2018/12/7
'                     StrTemp7(7) = ""
                     For i = 1 To 7
                         If Len(StrTemp7(i)) = 0 Then
                              If i <> 1 Then
                                 'Add By Sindy 2018/12/7
                                 If StrTemp7(i - 1) <> "合計" Then
                                 '2018/12/7 END
                                    If Page > 1 Then
                                       StrTemp7(i) = "合計"
                                    Else
                                       StrTemp7(7) = "合計"
                                    End If
                                 End If
                              End If
                             If i <> 7 Then
                                 If i = 1 Then
                                    'Modify By Sindy 2018/12/6 + And Val(CheckStr(.Fields(0))) > 2
                                    If Page > 1 And Val(CheckStr(.Fields(0))) > 2 Then
                                        StrTemp7(i) = "總計"
                                    End If
                                 Else
                                    'Modify By Sindy 2018/12/6 + And Val(CheckStr(.Fields(0))) > 2
                                    If Page > 1 And Val(CheckStr(.Fields(0))) > 2 Then
                                       'Add By Sindy 2018/12/7
                                       If StrTemp7(i - 1) = "合計" Then
                                          StrTemp7(i) = "總計"
                                       Else
                                       '2018/12/7 END
                                          StrTemp7(i + 1) = "總計"
                                       End If
                                    End If
                                 End If
                             End If
                             Exit For
                         End If
                     Next i
                End If
                CheckOC2
                Page = Page + 1
                SavDay1 = strTemp(0)
                Printer.NewPage
                PrintTitle
                PrintTitle2
                SavDay2 = " "
                SavDay3 = " "
            End If
            'edit by nick 2004/10/06
            'If SavDay2 <> strTemp(1) Then
            If SavDay2 <> strTemp(1) Or StrToStr(SavDay3, 1) <> StrToStr(strTemp(19), 1) Then
                If Len(Trim(SavDay2)) <> 0 Then
                ShowLine2
                PrintEnd2 (0)
                End If
                SavDay2 = strTemp(1)
                If Len(Trim(SavDay3)) <> 0 Then
                If StrToStr(SavDay3, 1) <> StrToStr(strTemp(19), 1) Then
                    ShowLine2
                    PrintEnd2 (1)
                    
                End If
                ShowLine2
                End If
                SavDay3 = strTemp(19)
            Else
               strTemp(1) = ""
            End If
            Else
            SavDay1 = strTemp(0)
            SavDay2 = strTemp(1)
            SavDay3 = strTemp(19)
            End If
            PrintDatil2
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
ShowLine2
PrintEnd2 (0)
ShowLine2
PrintEnd2 (1)
ShowLine2
PrintEnd2 (3)
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

'Add By Cheng 2002/02/06
'刪除固定區的合計為 0 的資料, 及非固定區的總計為 0 的資料
'搜尋固定區的第一頁
strSQL_1 = "SELECT R068001,NVL(A0902,A0903),nvl(st02,R068003),R068004+R068005+R068006+R068007+R068008+R068009+R068010+R068011+R068012+R068013+R068014+R068015+R068016+R068017+R068018+R068019+R068020+R068021+R068022+R068023+R068024+R068025+R068026+R068027 As Total ,R068002,r068003 " & _
            " FROM R020403_2,ACC090,staff " & _
            " WHERE ID='" & strUserNum & "' and r068003=st01(+) AND r068002=a0901(+) And R068001 = '1' ORDER BY R068001,R068002,R068003 "
If rstmp_1.State <> adStateClosed Then rstmp_1.Close
rstmp_1.CursorLocation = adUseClient
rstmp_1.Open strSQL_1, cnnConnection, adOpenStatic, adLockReadOnly
If rstmp_1.RecordCount > 0 Then
   rstmp_1.MoveFirst
   While Not rstmp_1.EOF
      '搜尋固定區的第二頁
      strSQL_2 = "SELECT R068001,NVL(A0902,A0903),nvl(st02,R068003),R068004+R068005+R068006+R068007+R068008+R068009+R068010+R068011+R068012+R068013+R068014+R068015+R068016+R068017+R068018+R068019+R068020+R068021+R068022+R068023+R068024+R068025+R068026+R068027 As Total ,R068002,r068003 " & _
                  " FROM R020403_2,ACC090,staff " & _
                  " WHERE ID='" & strUserNum & "' and r068003=st01(+) AND r068002=a0901(+) And R068001 = '2' And R068002='" & rstmp_1("R068002").Value & "' And R068003='" & rstmp_1("R068003").Value & "' ORDER BY R068001,R068002,R068003 "
      If rstmp_2.State <> adStateClosed Then rstmp_2.Close
      rstmp_2.CursorLocation = adUseClient
      rstmp_2.Open strSQL_2, cnnConnection, adOpenStatic, adLockReadOnly
      If rstmp_2.RecordCount > 0 Then
         rstmp_2.MoveFirst
         '若固定區合計為0
         If rstmp_1("Total") + rstmp_2("Total") = 0 Then
            cnnConnection.Execute "Delete From R020403_2 Where ID ='" & strUserNum & "' And R068002='" & rstmp_1("R068002").Value & "' And R068003='" & rstmp_1("R068003").Value & "' And (R068001='1' Or R068001='2') "
            '搜尋固定區的第三頁以上
            strSQL_3 = "SELECT R068001,NVL(A0902,A0903),nvl(st02,R068003),R068004+R068005+R068006+R068007+R068008+R068009+R068010+R068011+R068012+R068013+R068014+R068015+R068016+R068017+R068018+R068019+R068020+R068021+R068022+R068023+R068024+R068025+R068026+R068027 As Total ,R068002,r068003 " & _
                        " FROM R020403_2,ACC090,staff " & _
                        " WHERE ID='" & strUserNum & "' and r068003=st01(+) AND r068002=a0901(+) And (R068001 <> '1' AND R068001 <> '2') And R068002='" & rstmp_1("R068002").Value & "' And R068003='" & rstmp_1("R068003").Value & "' ORDER BY R068001,R068002,R068003 "
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
                  cnnConnection.Execute "Delete From R020403_2 Where ID ='" & strUserNum & "' And R068002='" & rstmp_1("R068002").Value & "' And R068003='" & rstmp_1("R068003").Value & "' And (R068001<>'1' And R068001<>'2') "
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

strSql = "SELECT R068001,NVL(A0902,A0903),nvl(st02,R068003),R068004,R068005,R068006,R068007,R068008,R068009,R068010,R068011,R068012,R068013,R068014,R068015,R068016,R068017,R068018,R068019,R068020,R068021,R068022,R068023,R068024,R068025,R068026,R068027,R068002,r068003 FROM R020403_2,staff,ACC090 WHERE ID='" & strUserNum & "' and r068003=st01(+) AND r068002=A0901(+) ORDER BY R068001,R068002,R068003 "
CheckOC
Page = 1
TestOk = True
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        SavDay1 = CheckStr(.Fields(0)) '1,2,3 頁
        SavDay2 = CheckStr(.Fields(1)) '部門名稱
        SavDay3 = CheckStr(.Fields(27)) '部門代碼
        strSql = "SELECT R067001,R067004,R067005,R067006,R067007,R067008,R067009,R067010 FROM R020403_1 WHERE ID='" & strUserNum & "' AND R067001=" & Val(SavDay1) & " "
        CheckOC2
        adoRecordset1.CursorLocation = adUseClient
        adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
            For i = 1 To 7
                StrTemp7(i) = StrToStr(CheckStr(adoRecordset1.Fields(i)), 7)
            Next i
            'Modify By Cheng 2002/02/06
'            StrTemp7(7) = ""
            For i = 1 To 7
                If Len("" & StrTemp7(i)) = 0 And Val(SavDay1) >= 2 Then
                     If i <> 1 Then
                            If Page > 1 Then
                               StrTemp7(i) = "合計"
                            Else
                               StrTemp7(7) = "合計"
                            End If
                     End If
                     If i <> 7 Then
                         If i = 1 Then
                            If Page > 1 Then
                                 StrTemp7(i) = "總計"
                            End If
                         Else
                            If Page > 1 Then
                                 StrTemp7(i + 1) = "總計"
                            End If
                         End If
                     End If
                     Exit For
                End If
            Next i
        End If
        CheckOC2
        PrintTitle
        PrintTitle1
        SavDay1 = " "
        SavDay2 = " "
        SavDay3 = " "
        Do While .EOF = False
            For i = 0 To 27
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(1) = StrToStr(strTemp(1), 5)
            strTemp(2) = StrToStr(strTemp(2), 3)
            If Len(Trim(SavDay1)) > 0 Or Len(Trim(SavDay2)) > 0 Or Len(Trim(SavDay3)) > 0 Then
            If Val(strTemp(0)) <> Val(SavDay1) Then
                ShowLine1
                PrintEnd1 (0)
                ShowLine1
                PrintEnd1 (1)
                If Val(SavDay1) = 1 Then
                    ShowLine1
                    PrintEnd1 (2)
                    ShowLine1
                Else
                    ShowLine1
                    PrintEnd1 (3)
                    ShowLine1
                End If
                strSql = "SELECT R067001,R067004,R067005,R067006,R067007,R067008,R067009,R067010 FROM R020403_1 WHERE ID='" & strUserNum & "' AND R067001=" & Val(CheckStr(.Fields(0))) & " "
                CheckOC2
                adoRecordset1.CursorLocation = adUseClient
                adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                    For i = 1 To 7
                        StrTemp7(i) = StrToStr(CheckStr(adoRecordset1.Fields(i)), 7)
                    Next i
                    'Modify By Sindy 2018/12/7
'                     StrTemp7(7) = ""
                     For i = 1 To 7
                         If Len(StrTemp7(i)) = 0 Then
                              If i <> 1 Then
                                 'Add By Sindy 2018/12/7
                                 If StrTemp7(i - 1) <> "合計" Then
                                 '2018/12/7 END
                                    If Page > 1 Then
                                       StrTemp7(i) = "合計"
                                    Else
                                       StrTemp7(7) = "合計"
                                    End If
                                 End If
                              End If
                             If i <> 7 Then
                                 If i = 1 Then
                                    'Modify By Sindy 2018/12/6 + And Val(CheckStr(.Fields(0))) > 2
                                    If Page > 1 And Val(CheckStr(.Fields(0))) > 2 Then
                                        StrTemp7(i) = "總計"
                                    End If
                                 Else
                                    'Modify By Sindy 2018/12/6 + And Val(CheckStr(.Fields(0))) > 2
                                    If Page > 1 And Val(CheckStr(.Fields(0))) > 2 Then
                                       'Add By Sindy 2018/12/7
                                       If StrTemp7(i - 1) = "合計" Then
                                          StrTemp7(i) = "總計"
                                       Else
                                       '2018/12/7 END
                                          StrTemp7(i + 1) = "總計"
                                       End If
                                    End If
                                 End If
                             End If
                             Exit For
                         End If
                     Next i
                End If
                CheckOC2
                Page = Page + 1
                SavDay1 = strTemp(0)
                Printer.NewPage
                PrintTitle
                PrintTitle1
                SavDay2 = " "
                SavDay3 = " "
            End If
            'edit by nick 2004/10/06
            'If SavDay2 <> strTemp(1) Then
            If SavDay2 <> strTemp(1) Or StrToStr(SavDay3, 1) <> StrToStr(strTemp(27), 1) Then
                If Len(Trim(SavDay2)) <> 0 Then
                  ShowLine1
                  PrintEnd1 (0)
                End If
                SavDay2 = strTemp(1)
                If Len(Trim(SavDay3)) <> 0 Then
                If StrToStr(SavDay3, 1) <> StrToStr(strTemp(27), 1) Then
                    ShowLine1
                    PrintEnd1 (1)
                    
                End If
                ShowLine1
                End If
                SavDay3 = strTemp(27)
            Else
            strTemp(1) = ""
            End If
            Else
            SavDay1 = strTemp(0)
            SavDay2 = strTemp(1)
            SavDay3 = strTemp(27)
            End If
            PrintDatil1
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
ShowLine1
PrintEnd1 (0)
ShowLine1
PrintEnd1 (1)
ShowLine1
PrintEnd1 (3)
ShowLine1
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintEnd1(Strindex As Integer)
Select Case Strindex
Case 0
     strSql = "select '','各區小計','',sum(r068004),sum(r068005),sum(r068006),sum(r068007),sum(r068008),sum(r068009),sum(r068010),sum(r068011),sum(r068012),sum(r068013),sum(r068014),sum(r068015),sum(r068016),sum(r068017),sum(r068018),sum(r068019),sum(r068020),sum(r068021),sum(r068022),sum(r068023),sum(r068024),sum(r068025),sum(r068026),sum(r068027) from r020403_2 where id='" & strUserNum & "' AND R068002='" & SavDay3 & "' AND R068001=" & Val(SavDay1) & " "
Case 1
     strSql = "select '','各所小計','',sum(r068004),sum(r068005),sum(r068006),sum(r068007),sum(r068008),sum(r068009),sum(r068010),sum(r068011),sum(r068012),sum(r068013),sum(r068014),sum(r068015),sum(r068016),sum(r068017),sum(r068018),sum(r068019),sum(r068020),sum(r068021),sum(r068022),sum(r068023),sum(r068024),sum(r068025),sum(r068026),sum(r068027) from r020403_2 where id='" & strUserNum & "' AND SUBSTR(R068002,1,2)='" & StrToStr(SavDay3, 1) & "' AND R068001=" & Val(SavDay1) & " "
Case 2
     strSql = "select '','全所總計','',sum(r068004),sum(r068005),sum(r068006),sum(r068007),sum(r068008),sum(r068009),sum(r068010),sum(r068011),sum(r068012),sum(r068013),sum(r068014),sum(r068015),sum(r068016),sum(r068017),sum(r068018),sum(r068019),sum(r068020),sum(r068021),sum(r068022),sum(r068023),sum(r068024),sum(r068025),sum(r068026),sum(r068027) from r020403_2 where id='" & strUserNum & "' AND R068001=1 "
     BolEndThisPage = True
Case 3
     strSql = "select '','全所總計','',sum(r068004),sum(r068005),sum(r068006),sum(r068007),sum(r068008),sum(r068009),sum(r068010),sum(r068011),sum(r068012),sum(r068013),sum(r068014),sum(r068015),sum(r068016),sum(r068017),sum(r068018),sum(r068019),sum(r068020),sum(r068021),sum(r068022),sum(r068023),sum(r068024),sum(r068025),sum(r068026),sum(r068027) from r020403_2 where id='" & strUserNum & "' AND R068001=" & Val(SavDay1) & " "
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
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print CheckStr(.Fields(1))
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print CheckStr(.Fields(2))
            For i = 2 To 22
               If Len(Trim(StrTemp7(Int((i + 1) / 3)))) <> 0 Or (Page < 3 And i < 16) Then '20
                   Select Case i
                  Case 3, 6, 9, 12, 15, 18, 21
                       Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(CheckNum(.Fields(i + 1)) & " ", "###,##0.00"))
                       Printer.CurrentY = iPrint
                       Printer.Print Format(CheckNum(.Fields(i + 1)) & " ", "###,##0.00")
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

Sub PrintEnd2(Strindex As Integer)
Select Case Strindex
Case 0
     strSql = "select '','各區小計','',sum(r068004),sum(r068005),sum(r068007),sum(r068008),sum(r068010),sum(r068011),sum(r068013),sum(r068014),sum(r068016),sum(r068017),sum(r068019),sum(r068020),sum(r068022),sum(r068023),sum(r068025),sum(r068026) from r020403_2 where id='" & strUserNum & "' AND R068002='" & SavDay3 & "' AND R068001=" & Val(SavDay1) & " "
Case 1
     strSql = "select '','各所小計','',sum(r068004),sum(r068005),sum(r068007),sum(r068008),sum(r068010),sum(r068011),sum(r068013),sum(r068014),sum(r068016),sum(r068017),sum(r068019),sum(r068020),sum(r068022),sum(r068023),sum(r068025),sum(r068026) from r020403_2 where id='" & strUserNum & "' AND SUBSTR(R068002,1,2)='" & StrToStr(SavDay3, 1) & "' AND R068001=" & Val(SavDay1) & " "
Case 2
     strSql = "select '','全所總計','',sum(r068004),sum(r068005),sum(r068007),sum(r068008),sum(r068010),sum(r068011),sum(r068013),sum(r068014),sum(r068016),sum(r068017),sum(r068019),sum(r068020),sum(r068022),sum(r068023),sum(r068025),sum(r068026) from r020403_2 where id='" & strUserNum & "' AND R068001=1 "
     BolEndThisPage = True
Case 3
     strSql = "select '','全所總計','',sum(r068004),sum(r068005),sum(r068007),sum(r068008),sum(r068010),sum(r068011),sum(r068013),sum(r068014),sum(r068016),sum(r068017),sum(r068019),sum(r068020),sum(r068022),sum(r068023),sum(r068025),sum(r068026) from r020403_2 where id='" & strUserNum & "' AND R068001=" & Val(SavDay1) & " "
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
            For i = 2 To 14 Step 2
                Printer.Line (PLeft(i) - 50, iPrint - 150)-(PLeft(i) - 50, iPrint + 450)
            Next i
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print CheckStr(.Fields(1))
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print CheckStr(.Fields(2))
            For i = 2 To 15
               If Len(Trim(StrTemp7(Int((i) / 2)))) <> 0 Or (Page < 3 And i < 12) Then
                      Select Case i
                     Case 3, 5, 7, 9, 11, 13, 15
                        Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(CheckNum(.Fields(i + 1)) & " ", "###,##0.00"))
                        Printer.CurrentY = iPrint
                        Printer.Print Format(CheckNum(.Fields(i + 1)) & " ", "###,##0.00")
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
'Printer.Orientation = 1 'Removed by Morgan 2015/6/3
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
If Val(txt1(1)) = 1 Then
    If Val(SavDay1) <= 2 Then
        Printer.Print GetTitleNick & "商爭案智權人員收文統計表(固定)"
    Else
        Printer.Print GetTitleNick & "商爭案智權人員收文統計表(非固定)"
    End If
Else
    If Val(SavDay1) <= 2 Then
        Printer.Print GetTitleNick & "商爭案智權人員發文統計表(固定)"
    Else
        Printer.Print GetTitleNick & "商爭案智權人員發文統計表(非固定)"
    End If
End If
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 7000
Printer.CurrentY = iPrint
If Val(txt1(1)) = 1 Then
    Printer.Print "收文日：" & Format(ChangeTStringToTDateString(txt1(2)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(3))
Else
    Printer.Print "發文日：" & Format(ChangeTStringToTDateString(txt1(2)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(3))
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
Printer.Print "業務區"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
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
   If Len(Trim(StrTemp7(Int((i) / 2)))) <> 0 Or (Page < 3 And i < 12) Then
      Select Case i
      Case 3, 5, 7, 9, 11, 13, 15
            Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(strTemp(i + 1) & " ", "###,##0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(strTemp(i + 1) & " ", "###,##0.00")
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
PLeft(1) = 1200
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
Printer.Print "業務區"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
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
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(1)
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(2)
For i = 2 To 22
   If Len(Trim(StrTemp7(Int((i + 1) / 3)))) <> 0 Or (Page < 3 And i < 16) Then '20
      Select Case i
      Case 3, 6, 9, 12, 15, 18, 21
           Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(strTemp(i + 1) & " ", "###,##0.00"))
           Printer.CurrentY = iPrint
           Printer.Print Format(strTemp(i + 1) & " ", "###,##0.00")
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
PLeft(1) = 1200
PLeft(2) = 2200
For i = 3 To 22
    PLeft(i) = 2200 + ((i - 2) * 700)
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
Set frm020403 = Nothing
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
Case 3, 2
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
Case 5
     If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If

Case Else
End Select
End Sub

