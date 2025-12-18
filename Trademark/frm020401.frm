VERSION 5.00
Begin VB.Form frm020401 
   BorderStyle     =   1  '單線固定
   Caption         =   "商申案智權人員收/發文統計表"
   ClientHeight    =   2930
   ClientLeft      =   1970
   ClientTop       =   1320
   ClientWidth     =   3910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2930
   ScaleWidth      =   3910
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   585
      Left            =   60
      TabIndex        =   14
      Top             =   2070
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
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3132
      TabIndex        =   12
      Top             =   12
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2340
      TabIndex        =   11
      Top             =   12
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1725
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   972
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1725
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1968
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1395
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   972
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1395
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   972
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1050
      Width           =   270
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   972
      TabIndex        =   0
      Top             =   690
      Width           =   1740
   End
   Begin VB.Line Line2 
      X1              =   1440
      X2              =   2190
      Y1              =   1890
      Y2              =   1890
   End
   Begin VB.Line Line1 
      X1              =   1395
      X2              =   2655
      Y1              =   1530
      Y2              =   1530
   End
   Begin VB.Label Label1 
      Caption         =   "(1.收文  2.發文)"
      Height          =   180
      Index           =   4
      Left            =   1290
      TabIndex        =   13
      Top             =   1080
      Width           =   1410
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   3
      Left            =   60
      TabIndex        =   10
      Top             =   1785
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "日期："
      Height          =   180
      Index           =   2
      Left            =   60
      TabIndex        =   9
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "列印別："
      Height          =   180
      Index           =   1
      Left            =   60
      TabIndex        =   8
      Top             =   1095
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   7
      Top             =   750
      Width           =   915
   End
End
Attribute VB_Name = "frm020401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/9 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, SavDay5 As String, SavDay4 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 30) As String, strTemp3 As String, TestOk As Boolean, StrTemp7(1 To 9) As String, intS As Integer, k As Integer
Dim PLeft(0 To 28) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, PLeft1(1 To 9) As Integer, SeekPrint As Integer, SeekPrintL As Integer
Dim BolEndThisPage As Boolean


Private Sub cmdok_Click(Index As Integer)
'Add By Cheng 2003/02/12
On Error GoTo ErrorHandler
Select Case Index
Case 0 '確定
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
'Add By Cheng 2003/02/12
Exit Sub
ErrorHandler:
    Select Case Err.Number
    Case 380
        MsgBox "印表機選擇錯誤!!!", vbExclamation + vbOKOnly
    Case Else
        MsgBox "(" & Err.Number & ")" & Err.Description, vbExclamation + vbOKOnly
    End Select
    Me.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

Sub Process()
'Add By Cheng 2003/09/03
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim intCnt As Integer '商品類別數
'Modify By Sindy 2015/1/22
Dim System_ID As String
Dim bolIsChina As Boolean
Dim Title_101 As String
Dim Title_102 As String
Dim Title_501 As String
Dim Title_502 As String
Dim Title_301 As String
Dim Title_103 As String
Dim Title_304 As String
'2015/1/22
Dim strMainSql As String 'Add By Sindy 2019/1/18

Screen.MousePointer = vbHourglass
cnnConnection.Execute "delete from R020401 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "DELETE FROM R020401_1 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "DELETE FROM R020401_2 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL6 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/10/19
End If
'Add By Cheng 2003/07/11
'商申收發文若為TF案則只抓後三碼為"000"的資料
strSQL1 = strSQL1 + " AND CP03 = Decode(CP01,'TF','0',CP03) AND CP04 = Decode(CP01,'TF','00',CP04) "
StrSQL6 = ""
Select Case Val(txt1(1))
Case 1 '收文列印
    If Len(txt1(2)) <> 0 Then
       StrSQL6 = StrSQL6 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(2))) & ""
    End If
    If Len(Trim(txt1(3))) <> 0 Then
       StrSQL6 = StrSQL6 + " AND CP05<=" & Val(ChangeTStringToWString(txt1(3))) & " "
    End If
    If Len(txt1(2)) <> 0 Or Len(Trim(txt1(3))) <> 0 Then
      pub_QL05 = pub_QL05 & ";收文" & Label1(2) & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/10/19
    End If
    'Add By Cheng 2003/05/07
    '不統計無發文日但有取消收文日的資料
    StrSQL6 = StrSQL6 & " And ((CP27 IS NULL And CP57 IS NULL) Or (CP27 IS NOT NULL And CP57 IS NULL) Or (CP27 IS NOT NULL And CP57 IS NOT NULL)) "
Case 2 '發文列印
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
'modify by sonia 2020/7/27 剔除TT-999999法律所案源資料
StrSQL6 = StrSQL6 & " And CP26 Is Null and cp01||cp02<>'TT999999' "
'End
CheckOC

'Modify By Sindy 2015/1/22
Dim MyCPMStr1 As String
Dim MyCPMStr2 As String
If txt1(4) = "020" And txt1(5) = "020" Then
    MyCPMStr1 = "cpm04"
    MyCPMStr2 = "cpm04"
Else
    MyCPMStr1 = "DECODE(NVL(CPM03,CP10),'（無）',CPM04,NVL(CPM03,CP10))"
    MyCPMStr2 = "DECODE(NVL(CPM03,CP10),'（無）',CPM04,NVL(CPM03,CP10))"
End If

'Modify By Cheng 2002/02/05
'取消抓承辦人之ST05為"95"或"15"的條件, 改為抓案件性質非"4"字頭且非"6"字頭且CP09<"B"的資料
'若從內商的統計報表進入, 承辦人部門別必須為"P2"字頭或承辦人為NULL; 若從外商的統計報表進入, 承辦人部門別必須為"F1"字頭或承辦人為NULL
'注意：業務區是CASEPROGRESS的CP12
'strMainSql = "SELECT S1.st03,cp13,decode(tm10,'000',CPM03,CPM04),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF S1,STAFF S2 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and (S2.ST05='95' OR S2.ST05='15') and cp09<'C' " & strSQL1 + StrSQL6
'strMainSql = strMainSql + " union all select S1.st03,cp13,decode(sp09,'000',CPM03,CPM04),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1,STAFF S2 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and (S2.ST05='95' OR S2.ST05='15') AND CP09<'C' " & strSQL2 + StrSQL6
'Modify By Cheng 2002/02/08 案件名稱一律顯示台灣名稱
'strMainSql = "SELECT CP12,cp13,CPM03,1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 " & _
'         " FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' Or CP14 Is NULL) ", " And (SubStr(ST03,1,2)='F1' Or CP14 Is NULL) ") & " AND CP01=CPM01(+) AND CP10=CPM02(+) And ((SubStr(CP10,1,1)<>'4' And SubStr(CP10,1,1)<>'6') And cp09<'B') " & strSQL1 + StrSQL6
'strMainSql = strMainSql + " union all select CP12,cp13,CPM03,1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 " & _
'         " FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF " & _
'         " WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' Or CP14 Is NULL) ", " And (SubStr(ST03,1,2)='F1' Or CP14 Is NULL) ") & " AND CP01=CPM01(+) AND CP10=CPM02(+) And ((SubStr(CP10,1,1)<>'4' And SubStr(CP10,1,1)<>'6') AND CP09<'B') " & strSQL2 + StrSQL6
'Modify By Cheng 2002/04/02
'不論收發文再加案件性質非"202","204","205","207"的限制條件
'Modify By Cheng 2002/04/11
'若台灣案件性質名稱為"(無)",則改抓大陸案件性質名稱
'strMainSql = "SELECT CP12,cp13,CPM03,1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 " & _
'         " FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' Or CP14 Is NULL) ", " And (SubStr(ST03,1,2)='F1' Or CP14 Is NULL) ") & " AND CP01=CPM01(+) AND CP10=CPM02(+) And ((SubStr(CP10,1,1)<>'4' And SubStr(CP10,1,1)<>'6') AND CP10<>'202' AND CP10<>'204' AND CP10<>'205' And cp09<'B') " & strSQL1 + StrSQL6
'strMainSql = strMainSql + " union all select CP12,cp13,CPM03,1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 " & _
'         " FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF " & _
'         " WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' Or CP14 Is NULL) ", " And (SubStr(ST03,1,2)='F1' Or CP14 Is NULL) ") & " AND CP01=CPM01(+) AND CP10=CPM02(+) And ((SubStr(CP10,1,1)<>'4' And SubStr(CP10,1,1)<>'6') AND CP10<>'202' AND CP10<>'204' AND CP10<>'205' And CP09<'B') " & strSQL2 + StrSQL6
'**************** 將業務區改成抓案件進度檔   91.08.15  nick
'92.3.6 MODIFY BY SONIA 若從外商的統計報表進入, 不管承辦人部門別
'strMainSql = "SELECT CP12,cp13,DECODE(NVL(CPM03,CPM04),'（無）',CPM04,CPM03),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 " & _
'         " FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' Or CP14 Is NULL) ", " And (SubStr(ST03,1,2)='F1' Or CP14 Is NULL) ") & " AND CP01=CPM01(+) AND CP10=CPM02(+) And ((SubStr(CP10,1,1)<>'4' And SubStr(CP10,1,1)<>'6') AND CP10<>'202' AND CP10<>'204' AND CP10<>'205' AND CP10<>'207' And cp09<'B') " & strSQL1 + StrSQL6
'Modify By Cheng 2003/07/01
'自請撤回(306)要單獨處理
'strMainSql = "SELECT CP12,cp13,DECODE(NVL(CPM03,CPM04),'（無）',CPM04,CPM03),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 " & _
'         " FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' Or CP14 Is NULL) ", "") & " AND CP01=CPM01(+) AND CP10=CPM02(+) And ((SubStr(CP10,1,1)<>'4' And SubStr(CP10,1,1)<>'6') AND CP10<>'202' AND CP10<>'204' AND CP10<>'205' AND CP10<>'207' And cp09<'B') " & strSQL1 + StrSQL6
'2011/12/7 modify by sonia 加入AND CP10<>'210'
'Modify By Sindy 2021/2/23 改抓案件性質的語法加入TMdebate
'  " And ((SubStr(CP10,1,1)<>'4' And SubStr(CP10,1,1)<>'6') AND CP10<>'202' AND CP10<>'210' AND CP10<>'204' AND CP10<>'205' AND CP10<>'207' And CP10<>'306')"
' ==> " and instr('" & TMdebate & ",204,205,207,306',cp10)=0"
'Modify By Sindy 2021/7/9 排除跨類107,另外判斷
'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
strMainSql = "SELECT CP12,cp13," & MyCPMStr1 & ",1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 " & _
         " FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' Or CP14 Is NULL) ", "") & " AND CP01=CPM01(+) AND CP10=CPM02(+)" & _
         " and (instr('" & TMdebate & ",204,205,207,306,107',cp10)=0 or (cp01='FCT' And InStr('" & FCT_NotTMdebate & "', cp10) > 0))" & _
         " And cp09<'B' " & strSQL1 + StrSQL6
'92.3.6 END
'Modify By Sindy 2021/7/9 跨類107,若該案件申請101那一道的資料年月區間與跨類不相同時，跨類就要單獨計算，而且要出現在非固定報表內。
strSql = strSql + " union all SELECT CP12,cp13," & MyCPMStr1 & ",1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 " & _
         " FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' Or CP14 Is NULL) ", "") & " AND CP01=CPM01(+) AND CP10=CPM02(+)" & _
         " and cp10='107'" & _
         " And cp09<'B' " & strSQL1 + StrSQL6 & _
         " AND not Exists (select * from caseprogress cp2 where cp2.cp01=cp01 and cp2.cp02=cp02 and cp2.cp03=cp03 and cp2.cp04=cp04 and cp2.cp10='101'" & IIf(txt1(1) = "1", " and substr(cp2.cp05,1,6)=substr(cp05,1,6)", " and substr(cp2.cp27,1,6)=substr(cp27,1,6)") & ")"
         
'**************** 將業務區改成抓案件進度檔   91.08.15  nick
'92.3.6 MODIFY BY SONIA 若從外商的統計報表進入, 不管承辦人部門別
'strMainSql = strMainSql + " union all select CP12,cp13,DECODE(NVL(CPM03,CPM04),'（無）',CPM04,CPM03),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 " & _
'         " FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF " & _
'         " WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' Or CP14 Is NULL) ", " And (SubStr(ST03,1,2)='F1' Or CP14 Is NULL) ") & " AND CP01=CPM01(+) AND CP10=CPM02(+) And ((SubStr(CP10,1,1)<>'4' And SubStr(CP10,1,1)<>'6') AND CP10<>'202' AND CP10<>'204' AND CP10<>'205' AND CP10<>'205' And CP09<'B') " & strSQL2 + StrSQL6
'Modify By Cheng 2003/07/01
'自請撤回(306)要單獨處理
'strMainSql = strMainSql + " union all select CP12,cp13,DECODE(NVL(CPM03,CPM04),'（無）',CPM04,CPM03),1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 " & _
'         " FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF " & _
'         " WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' Or CP14 Is NULL) ", "") & " AND CP01=CPM01(+) AND CP10=CPM02(+) And ((SubStr(CP10,1,1)<>'4' And SubStr(CP10,1,1)<>'6') AND CP10<>'202' AND CP10<>'204' AND CP10<>'205' AND CP10<>'205' And CP09<'B') " & strSQL2 + StrSQL6
'Modify By Sindy 2021/2/23 改抓案件性質的語法加入TMdebate
'  " And ((SubStr(CP10,1,1)<>'4' And SubStr(CP10,1,1)<>'6') AND CP10<>'202' AND CP10<>'210' AND CP10<>'204' AND CP10<>'205' AND CP10<>'207' And CP10<>'306')"
' ==> " and instr('" & TMdebate & ",204,205,207,306',cp10)=0"
'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
strMainSql = strMainSql + " union all select CP12,cp13," & MyCPMStr2 & ",1,CP18,DECODE(CP27,NULL,1,'',1,0),CP10,cp09 " & _
         " FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF " & _
         " WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' Or CP14 Is NULL) ", "") & " AND CP01=CPM01(+) AND CP10=CPM02(+)" & _
         " and (instr('" & TMdebate & ",204,205,207,306',cp10)=0 or (cp01='FCT' And InStr('" & FCT_NotTMdebate & "', cp10) > 0))" & _
         " And CP09<'B' " & strSQL2 + StrSQL6
'92.3.6 END
'Add By Cheng 2003/07/01
'自請撤回(306)須依其相關總收文號的案件性質判斷是商申案
'Modify By Sindy 2021/2/23 改抓案件性質的語法加入TMdebate
'  " And ((SubStr(CP10,1,1)<>'4' And SubStr(CP10,1,1)<>'6') AND CP10<>'202' AND CP10<>'210' AND CP10<>'204' AND CP10<>'205' AND CP10<>'207') "
' ==> " and instr('" & TMdebate & ",204,205,207',cp10)=0"
'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
strMainSql = strMainSql & " Union SELECT CP12,cp13," & MyCPMStr1 & ",1,CP18,DECODE(CP27,NULL,1,'',1,0), '306', cp09 " & _
         " FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF " & _
         " WHERE CP09 In (Select CP43 From CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' Or CP14 Is NULL) ", "") & " AND CP01=CPM01(+) AND CP10=CPM02(+) And CP10='306' And cp09<'B' " & strSQL1 + StrSQL6 & " ) " & _
         " And CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND '306'=CPM02(+)" & _
         " and (instr('" & TMdebate & ",204,205,207',cp10)=0 or (cp01='FCT' And InStr('" & FCT_NotTMdebate & "', cp10) > 0))"
strMainSql = strMainSql & " Union SELECT CP12,cp13," & MyCPMStr2 & ",1,CP18,DECODE(CP27,NULL,1,'',1,0), '306', cp09 " & _
         " FROM CASEPROGRESS, Servicepractice, CASEPROPERTYMAP,STAFF " & _
         " WHERE CP09 In (Select CP43 From CASEPROGRESS, Servicepractice, CASEPROPERTYMAP,STAFF WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) " & IIf(intPWhere = 國內, " And (SubStr(ST03,1,2)='P2' Or CP14 Is NULL) ", "") & " AND CP01=CPM01(+) AND CP10=CPM02(+) And CP10='306' And cp09<'B' " & strSQL2 + StrSQL6 & " ) " & _
         " And CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND '306'=CPM02(+)" & _
         " and (instr('" & TMdebate & ",204,205,207',cp10)=0 or (cp01='FCT' And InStr('" & FCT_NotTMdebate & "', cp10) > 0))"
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
Call ClsPDGetCaseProperty("T", "101", Title_101, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "102", Title_102, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "501", Title_501, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "502", Title_502, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "301", Title_301, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "103", Title_103, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "304", Title_304, bolIsChina, False)
'2015/1/22 END
With adoRecordset
    .CursorLocation = adUseClient
    .Open strMainSql, cnnConnection, adOpenStatic, adLockReadOnly
    'Fields--0:業務區, 1:智權人員代號, 2:案件性質名稱, 3:???, 4:點數, 5:判斷是否有發文日, 6:案件性質, 7:收文號
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/19
        .MoveFirst
        DoEvents
        Do While .EOF = False
            For i = 0 To 5
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            Select Case Val(CheckStr(.Fields(6)))
            '2005/6/2 MODIFY BY SONIA 分割改為非固定,不分類別
            'Case 101, 308 '申請(件)   '93.12.31 MODIFY BY SONIA 加入 308
            Case 101 '申請(件)
                  'strTemp(2) = "申請（件）"
                  strTemp(2) = Title_101 & "（件）" 'Modify By Sindy 2015/1/22
                  strTemp(6) = "*"
                  strTemp(7) = "1"
                  'Add By Sindy 2021/7/9 跨類107,若該案件申請101那一道的資料年月區間與跨類相同，
                  '則跨類的點數加入申請的點數，案數不算；而且在非固定報表內不再出現跨類。
                  strExc(0) = "SELECT sum(nvl(c2.cp18,0)) c2_cp18" & _
                              " FROM CASEPROGRESS c1,CASEPROGRESS c2" & _
                              " WHERE c1.CP09='" & Trim(.Fields(7)) & "'" & _
                              " AND c2.cp01(+)=c1.cp01 AND c2.cp02(+)=c1.cp02 AND c2.cp03(+)=c1.cp03 AND c2.cp04(+)=c1.cp04" & _
                              " AND c2.cp10='107'" & _
                              IIf(txt1(1) = "1", " and substr(c2.cp05,1,6)=substr(c1.cp05,1,6)", " and substr(c2.cp27,1,6)=substr(c1.cp27,1,6)")
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If Val("" & RsTemp.Fields(0)) > 0 Then
                        strTemp(4) = strTemp(4) + Val("" & RsTemp.Fields(0))
                     End If
                  End If
                  '2021/7/9 END
            Case 102 '延展
                  'strTemp(2) = "延展" 'Add By Sindy 2013/6/3
                  strTemp(2) = Title_102 'Modify By Sindy 2015/1/22
                  strTemp(6) = "*"
                  strTemp(7) = "3"
            Case 501 '移轉
                  'strTemp(2) = "移轉" 'Add By Sindy 2013/6/3
                  strTemp(2) = Title_501 'Modify By Sindy 2015/1/22
                  strTemp(6) = "*"
                  strTemp(7) = "4"
            Case 502 '授權
                  'strTemp(2) = "授權" 'Add By Sindy 2013/6/3
                  strTemp(2) = Title_502 'Modify By Sindy 2015/1/22
                  strTemp(6) = "*"
                  strTemp(7) = "5"
            Case 301 '變更
                  'strTemp(2) = "變更" 'Add By Sindy 2013/6/3
                  strTemp(2) = Title_301 'Modify By Sindy 2015/1/22
                  strTemp(6) = "*"
                  strTemp(7) = "6"
            Case 103 '補換發證書
                  'strTemp(2) = "補換發證書" 'Add By Sindy 2013/6/3
                  strTemp(2) = Title_103 'Modify By Sindy 2015/1/22
                  strTemp(6) = "*"
                  strTemp(7) = "7"
            Case 304 '申請英文證明
                  'strTemp(2) = "申請英文證明" 'Add By Sindy 2013/6/3
                  strTemp(2) = Title_304 'Modify By Sindy 2015/1/22
                  strTemp(6) = "*"
                  strTemp(7) = "8"
            'Modify By Cheng 2003/12/29
            '取消設定質權
'            'Modify By Cheng 2002/02/05
'            '將更正(302)移至非固定, 並以設定質權(506)取代之
''            Case 302 '更正
'            Case 506 '設定質權
'                 strTemp(6) = "*"
'                 strTemp(7) = "8"
            'End
            Case Else '其他
'               If Val(CheckStr(.Fields(6))) = "717" Then
'                  MsgBox Val(CheckStr(.Fields(6)))
'               End If
                  strTemp(6) = ""
                  strTemp(7) = "9"
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
            strSql = "INSERT INTO R020401 VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "'," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & ",'" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
            'Add By Cheng 2003/12/29
            '申請(類)
            '2005/6/2 MODIFY BY SONIA 分割改為非固定,不分類別
            '93.12.31 MODIFY BY SONIA 加入 308
            If "" & .Fields(6).Value = "101" Then
            'If ("" & .Fields(6).Value = "101") Or ("" & .Fields(6).Value = "308") Then
                intCnt = 0
                'strTemp(2) = "申請（類）"
                strTemp(2) = Title_101 & "（類）" 'Modify By Sindy 2015/1/22
                strTemp(6) = "*"
                strTemp(7) = "2"
                StrSQLa = "Select TM09, TM10 From Trademark, CaseProgress Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP09='" & .Fields(7).Value & "' "
                rsA.CursorLocation = adUseClient
                rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                If rsA.RecordCount > 0 Then
                    'Modify By Sindy 2021/7/9 原控制申請(類)區分台灣案及非台灣案不同計算方式案數欄，
                    '請取消申請國家的限制，申請(類)一律都依照基本檔的類別數來計算
'                    '若申請國家不為台灣(000)
'                    If "" & rsA.Fields(1).Value <> "000" Then
'                        intCnt = 1
'                    '若申請國家為台灣(000)
'                    Else
                        If "" & rsA.Fields(0).Value = "" Then
                            intCnt = intCnt + 1
                        Else
                            intCnt = UBound(Split(rsA.Fields(0).Value, ",")) + 1
                        End If
'                    End If
                    strTemp(3) = intCnt
                End If
                If rsA.State <> adStateClosed Then rsA.Close
                Set rsA = Nothing
                strSql = "INSERT INTO R020401 VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "'," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & ",'" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & strUserNum & "') "
                cnnConnection.Execute strSql
            End If
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
'固定式標題
'Modify By Cheng 2002/02/05
'cnnConnection.Execute "INSERT INTO R020401_1 VALUES (1,'','','申請','延展','移轉','授權','變更','補換發證書','申請英文證明','更正','" & strUserNum & "') "
'Modify By Cheng 2003/12/29
'cnnConnection.Execute "INSERT INTO R020401_1 VALUES (1,'','','申請','延展','移轉','授權','變更','補換發證書','申請英文證明','設定質權','" & strUserNum & "') "
'cnnConnection.Execute "INSERT INTO R020401_1 VALUES (1,'','','申請（件）','申請（類）','延展','移轉','授權','變更','補換發證書','申請英文證明','" & strUserNum & "') "
cnnConnection.Execute "INSERT INTO R020401_1 VALUES (1,'','','" & Title_101 & "（件）','" & Title_101 & "（類）','" & Title_102 & "','" & Title_501 & "','" & Title_502 & "','" & Title_301 & "','" & Title_103 & "','" & Title_304 & "','" & strUserNum & "') "
'End
'非固定式標題
strSql = "select distinct r063003 from r020401 where id='" & strUserNum & "' and (r063007='' or r063007 is null)"
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        DoEvents
        strTemp3 = "2"
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
            strSql = "insert into r020401_1 values (" & Val(strTemp3) & ",'" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
            strTemp3 = Trim(str(Val(strTemp3) + 1))
            If .EOF = False Then
               '.MoveNext Modify By Sindy 2019/1/18 Mark,會有跳掉案件性質的問題
            End If
            DoEvents
        Loop
    End If
End With
CheckOC
strSql = "SELECT * FROM R020401_1 WHERE ID='" & strUserNum & "' ORDER BY R064001 "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 3 To 10
                strSql = "SELECT * FROM R020401 WHERE ID='" & strUserNum & "' AND R063003='" & CheckStr(.Fields(i)) & "' "
                CheckOC2
                adoRecordset1.CursorLocation = adUseClient
                adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                    adoRecordset1.MoveFirst
                    Do While adoRecordset1.EOF = False
                        For j = 3 To 29
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
                        Case 10
                             strTemp(24) = CheckStr(adoRecordset1.Fields(3))
                             strTemp(25) = CheckStr(adoRecordset1.Fields(4))
                             strTemp(26) = CheckStr(adoRecordset1.Fields(5))
                        Case Else
                        End Select
                        strTemp(27) = CheckStr(adoRecordset1.Fields(3))
                        strTemp(28) = CheckStr(adoRecordset1.Fields(4))
                        strTemp(29) = CheckStr(adoRecordset1.Fields(5))
                        strSql = "INSERT INTO R020401_2 VALUES (" & Val(strTemp(0)) & ",'" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "'," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & "," & Val(strTemp(15)) & "," & Val(strTemp(16)) & "," & Val(strTemp(17)) & "," & Val(strTemp(18)) & "," & Val(strTemp(19)) & "," & Val(strTemp(20)) & "," & Val(strTemp(21)) & "," & Val(strTemp(22)) & "," & Val(strTemp(23)) & "," & Val(strTemp(24)) & "," & Val(strTemp(25)) & "," & Val(strTemp(26)) & "," & Val(strTemp(27)) & "," & Val(strTemp(28)) & "," & Val(strTemp(29)) & ",'" & strUserNum & "') "
                        cnnConnection.Execute strSql
                        adoRecordset1.MoveNext
                        DoEvents
                    Loop
                End If
                DoEvents
            Next i
            DoEvents
            .MoveNext
        Loop
    End If
End With
CheckOC
strSql = "select max(r064001) from r020401_1 where id='" & strUserNum & "' "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    s = Val(CheckStr(adoRecordset.Fields(0)))
End If
CheckOC
strSql = "select distinct r063001,r063002 from r020401 where id='" & strUserNum & "' "
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 1 To s
                strSql = "insert into r020401_2 values (" & i & ",'" & ChgSQL(CheckStr(.Fields(0))) & "','" & ChgSQL(CheckStr(.Fields(1))) & "',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "') "
                cnnConnection.Execute strSql
                DoEvents
            Next i
            .MoveNext
            DoEvents
        Loop
    End If
End With
CheckOC
'重整
'Modify By Cheng 2003/12/03
'排除申請（類）
'strSQL = "select r065001,R065002,r065003,sum(r065004),sum(r065005),sum(r065006),sum(r065007),sum(r065008),sum(r065009),sum(r065010),sum(r065011),sum(r065012),sum(r065013),sum(r065014),sum(r065015),sum(r065016),sum(r065017),sum(r065018),sum(r065019),sum(r065020),sum(r065021),sum(r065022),sum(r065023),sum(r065024),sum(r065025),sum(r065026),sum(r065027),sum(r065028),sum(r065029),sum(r065030) from r020401_2 where id='" & strUserNum & "' group by r065001,r065002,r065003 order by r065001,r065002,r065003 "
strSql = "select r065001,R065002,r065003,sum(r065004),sum(r065005),sum(r065006),sum(r065007),sum(r065008),sum(r065009),sum(r065010),sum(r065011),sum(r065012),sum(r065013),sum(r065014),sum(r065015),sum(r065016),sum(r065017),sum(r065018),sum(r065019),sum(r065020),sum(r065021),sum(r065022),sum(r065023),sum(r065024),sum(r065025),sum(r065026),sum(r065027),sum(r065028)-Sum(Decode(R065001, 1, R065007, 0)),sum(r065029)-Sum(Decode(R065001, 1, R065008, 0)) ,sum(r065030)-Sum(Decode(R065001, 1, R065009, 0)) from r020401_2 where id='" & strUserNum & "' group by r065001,r065002,r065003 order by r065001,r065002,r065003 "
'End
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        cnnConnection.Execute "DELETE FROM R020401_2 WHERE ID='" & strUserNum & "' "
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 29
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strSql = "INSERT INTO R020401_2 VALUES (" & Val(strTemp(0)) & ",'" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "'," & Val(strTemp(3)) & "," & Val(strTemp(4)) & "," & Val(strTemp(5)) & "," & Val(strTemp(6)) & "," & Val(strTemp(7)) & "," & Val(strTemp(8)) & "," & Val(strTemp(9)) & "," & Val(strTemp(10)) & "," & Val(strTemp(11)) & "," & Val(strTemp(12)) & "," & Val(strTemp(13)) & "," & Val(strTemp(14)) & "," & Val(strTemp(15)) & "," & Val(strTemp(16)) & "," & Val(strTemp(17)) & "," & Val(strTemp(18)) & "," & Val(strTemp(19)) & "," & Val(strTemp(20)) & "," & Val(strTemp(21)) & "," & Val(strTemp(22)) & "," & Val(strTemp(23)) & "," & Val(strTemp(24)) & "," & Val(strTemp(25)) & "," & Val(strTemp(26)) & "," & Val(strTemp(27)) & "," & Val(strTemp(28)) & "," & Val(strTemp(29)) & ",'" & strUserNum & "') "
            cnnConnection.Execute strSql
            .MoveNext
            DoEvents
        Loop
    End If
End With
CheckOC
s = 9
strSql = "select r064001,r064004,r064005,r064006,r064007,r064008,r064009,r064010,r064011 from r020401_1 where ID='" & strUserNum & "' AND r064001 in (select max(r064001) from r020401_1 where id='" & strUserNum & "') "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    For i = 1 To 8
        If Len(CheckStr(adoRecordset.Fields(i))) = 0 Then
            s = i
            Exit For
        End If
    Next i
End If
CheckOC
             strSql = "select max(r064001) from r020401_1 where id='" & strUserNum & "' "
             CheckOC2
             adoRecordset1.CursorLocation = adUseClient
             adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
             If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
             intS = Val(CheckStr(adoRecordset1.Fields(0)))
             End If
'Modify By Cheng 2003/12/30
'總計排除申請(類)
'strSQL = "select r065002,r065003,sum(r065004)+sum(r065007)+sum(r065010)+sum(r065013)+sum(r065016)+sum(r065019)+sum(r065022)+sum(r065025),sum(r065005)+sum(r065008)+sum(r065011)+sum(r065014)+sum(r065017)+sum(r065020)+sum(r065023)+sum(r065026),sum(r065006)+sum(r065009)+sum(r065012)+sum(r065015)+sum(r065018)+sum(r065021)+sum(r065024)+sum(r065027) from r020401_2 where id='" & strUserNum & "' group by r065002,r065003"
strSql = "select r065002,r065003,sum(r065004)+sum(Decode(R065001, 1, 0, r065007))+sum(r065010)+sum(r065013)+sum(r065016)+sum(r065019)+sum(r065022)+sum(r065025),sum(r065005)+sum(Decode(R065001, 1, 0, r065008))+sum(r065011)+sum(r065014)+sum(r065017)+sum(r065020)+sum(r065023)+sum(r065026),sum(r065006)+sum(Decode(R065001, 1, 0, r065009))+sum(r065012)+sum(r065015)+sum(r065018)+sum(r065021)+sum(r065024)+sum(r065027) from r020401_2 where id='" & strUserNum & "' group by r065002,r065003"
'End
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    adoRecordset.MoveFirst
    Do While adoRecordset.EOF = False
        Select Case s
        Case 1
             strSql = "SELECT R065028,R065029,R065030 FROM R020401_2 WHERE ID='" & strUserNum & "' AND R065001 IN (SELECT MAX(R065001) FROM R020401_2 WHERE ID='" & strUserNum & "') " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " and r065002 is null ", " AND R065002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r065003 is null ", " AND R065003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ")
             CheckOC2
             adoRecordset1.CursorLocation = adUseClient
             adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
             If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                strSql = "UPDATE R020401_2 SET R065028=0,R065029=0,R065030=0,R065004=" & Val(CheckStr(adoRecordset1.Fields(0))) & ",R065005=" & Val(CheckStr(adoRecordset1.Fields(1))) & ",R065006=" & Val(CheckStr(adoRecordset1.Fields(2))) & ",R065007=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R065008=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R065009=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r065002 is null ", " R065002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r065003 is null ", " AND R065003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND R065001 IN (SELECT MAX(R065001) FROM R020401_2 WHERE ID='" & strUserNum & "') "
                cnnConnection.Execute strSql
             End If
             CheckOC2
        Case 2
             strSql = "SELECT R065028,R065029,R065030 FROM R020401_2 WHERE ID='" & strUserNum & "' AND R065001 IN (SELECT MAX(R065001) FROM R020401_2 WHERE ID='" & strUserNum & "') " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " and r065002 is null ", " AND R065002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r065003 is null ", " AND R065003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ")
             CheckOC2
             adoRecordset1.CursorLocation = adUseClient
             adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
             If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                strSql = "UPDATE R020401_2 SET R065028=0,R065029=0,R065030=0,R065007=" & Val(CheckStr(adoRecordset1.Fields(0))) & ",R065008=" & Val(CheckStr(adoRecordset1.Fields(1))) & ",R065009=" & Val(CheckStr(adoRecordset1.Fields(2))) & ",R065010=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R065011=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R065012=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r065002 is null ", " R065002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r065003 is null ", " AND R065003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND R065001 IN (SELECT MAX(R065001) FROM R020401_2 WHERE ID='" & strUserNum & "') "
                cnnConnection.Execute strSql
             End If
             CheckOC2
        Case 3
             strSql = "SELECT R065028,R065029,R065030 FROM R020401_2 WHERE ID='" & strUserNum & "' AND R065001 IN (SELECT MAX(R065001) FROM R020401_2 WHERE ID='" & strUserNum & "') " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " and r065002 is null ", " AND R065002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r065003 is null ", " AND R065003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ")
             CheckOC2
             adoRecordset1.CursorLocation = adUseClient
             adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
             If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                strSql = "UPDATE R020401_2 SET R065028=0,R065029=0,R065030=0,R065010=" & Val(CheckStr(adoRecordset1.Fields(0))) & ",R065011=" & Val(CheckStr(adoRecordset1.Fields(1))) & ",R065012=" & Val(CheckStr(adoRecordset1.Fields(2))) & ",R065013=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R065014=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R065015=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r065002 is null ", " R065002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r065003 is null ", " AND R065003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND R065001 IN (SELECT MAX(R065001) FROM R020401_2 WHERE ID='" & strUserNum & "') "
                cnnConnection.Execute strSql
             End If
             CheckOC2
        Case 4
             strSql = "SELECT R065028,R065029,R065030 FROM R020401_2 WHERE ID='" & strUserNum & "' AND R065001 IN (SELECT MAX(R065001) FROM R020401_2 WHERE ID='" & strUserNum & "') " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " and r065002 is null ", " AND R065002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r065003 is null ", " AND R065003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ")
             CheckOC2
             adoRecordset1.CursorLocation = adUseClient
             adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
             If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                strSql = "UPDATE R020401_2 SET R065028=0,R065029=0,R065030=0,R065013=" & Val(CheckStr(adoRecordset1.Fields(0))) & ",R065014=" & Val(CheckStr(adoRecordset1.Fields(1))) & ",R065015=" & Val(CheckStr(adoRecordset1.Fields(2))) & ",R065016=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R065017=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R065018=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r065002 is null ", " R065002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r065003 is null ", " AND R065003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND R065001 IN (SELECT MAX(R065001) FROM R020401_2 WHERE ID='" & strUserNum & "') "
                cnnConnection.Execute strSql
             End If
             CheckOC2
        Case 5
             strSql = "SELECT R065028,R065029,R065030 FROM R020401_2 WHERE ID='" & strUserNum & "' AND R065001 IN (SELECT MAX(R065001) FROM R020401_2 WHERE ID='" & strUserNum & "') " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " and r065002 is null ", " AND R065002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r065003 is null ", " AND R065003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ")
             CheckOC2
             adoRecordset1.CursorLocation = adUseClient
             adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
             If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                strSql = "UPDATE R020401_2 SET R065028=0,R065029=0,R065030=0,R065016=" & Val(CheckStr(adoRecordset1.Fields(0))) & ",R065017=" & Val(CheckStr(adoRecordset1.Fields(1))) & ",R065018=" & Val(CheckStr(adoRecordset1.Fields(2))) & ",R065019=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R065020=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R065021=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r065002 is null ", " R065002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r065003 is null ", " AND R065003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND R065001 IN (SELECT MAX(R065001) FROM R020401_2 WHERE ID='" & strUserNum & "') "
                cnnConnection.Execute strSql
             End If
             CheckOC2
        Case 6
             strSql = "SELECT R065028,R065029,R065030 FROM R020401_2 WHERE ID='" & strUserNum & "' AND R065001 IN (SELECT MAX(R065001) FROM R020401_2 WHERE ID='" & strUserNum & "') " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " and r065002 is null ", " AND R065002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r065003 is null ", " AND R065003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ")
             CheckOC2
             adoRecordset1.CursorLocation = adUseClient
             adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
             If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                strSql = "UPDATE R020401_2 SET R065028=0,R065029=0,R065030=0,R065019=" & Val(CheckStr(adoRecordset1.Fields(0))) & ",R065020=" & Val(CheckStr(adoRecordset1.Fields(1))) & ",R065021=" & Val(CheckStr(adoRecordset1.Fields(2))) & ",R065022=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R065023=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R065024=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r065002 is null ", " R065002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r065003 is null ", " AND R065003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND R065001 IN (SELECT MAX(R065001) FROM R020401_2 WHERE ID='" & strUserNum & "') "
                cnnConnection.Execute strSql
             End If
             CheckOC2
        Case 7
             strSql = "SELECT R065028,R065029,R065030 FROM R020401_2 WHERE ID='" & strUserNum & "' AND R065001 IN (SELECT MAX(R065001) FROM R020401_2 WHERE ID='" & strUserNum & "') " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " and r065002 is null ", " AND R065002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r065003 is null ", " AND R065003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ")
             CheckOC2
             adoRecordset1.CursorLocation = adUseClient
             adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
             If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                strSql = "UPDATE R020401_2 SET R065028=0,R065029=0,R065030=0,R065022=" & Val(CheckStr(adoRecordset1.Fields(0))) & ",R065023=" & Val(CheckStr(adoRecordset1.Fields(1))) & ",R065024=" & Val(CheckStr(adoRecordset1.Fields(2))) & ",R065025=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R065026=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R065027=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r065002 is null ", " R065002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r065003 is null ", " AND R065003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND R065001 IN (SELECT MAX(R065001) FROM R020401_2 WHERE ID='" & strUserNum & "') "
                cnnConnection.Execute strSql
             End If
             CheckOC2
        Case 8
             strSql = "SELECT R065028,R065029,R065030 FROM R020401_2 WHERE ID='" & strUserNum & "' AND R065001 IN (SELECT MAX(R065001) FROM R020401_2 WHERE ID='" & strUserNum & "') " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " and r065002 is null ", " AND R065002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r065003 is null ", " AND R065003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ")
             CheckOC2
             adoRecordset1.CursorLocation = adUseClient
             adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
             If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                strSql = "UPDATE R020401_2 SET R065025=" & Val(CheckStr(adoRecordset1.Fields(0))) & ",R065026=" & Val(CheckStr(adoRecordset1.Fields(1))) & ",R065027=" & Val(CheckStr(adoRecordset1.Fields(2))) & ",R065028=" & Val(CheckStr(adoRecordset.Fields(2))) & ",R065029=" & Val(CheckStr(adoRecordset.Fields(3))) & ",R065030=" & Val(CheckStr(adoRecordset.Fields(4))) & " WHERE " & IIf(Len(CheckStr(adoRecordset.Fields(0))) = 0, " r065002 is null ", " R065002='" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "' ") & IIf(Len(CheckStr(adoRecordset.Fields(1))) = 0, " and r065003 is null ", " AND R065003='" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "' ") & " AND ID='" & strUserNum & "' AND R065001 IN (SELECT MAX(R065001) FROM R020401_2 WHERE ID='" & strUserNum & "') "
                cnnConnection.Execute strSql
             End If
             CheckOC2
        Case Else
                        strSql = "insert into r020401_2 values (" & intS + 1 & ",'" & ChgSQL(CheckStr(adoRecordset.Fields(0))) & "','" & ChgSQL(CheckStr(adoRecordset.Fields(1))) & "'," & Val(CheckStr(adoRecordset.Fields(2))) & "," & Val(CheckStr(adoRecordset.Fields(3))) & "," & Val(CheckStr(adoRecordset.Fields(4))) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,'" & strUserNum & "') "
                        cnnConnection.Execute strSql
                        strSql = "INSERT INTO R020401_1 VALUES (" & intS + 1 & ",'','','','','','','','','','','" & strUserNum & "') "
                        cnnConnection.Execute strSql
        End Select
        adoRecordset.MoveNext
        DoEvents
    Loop
End If
CheckOC
'Modify By Cheng 2002/02/05
'cnnConnection.Execute "UPDATE R020401_1 SET R064004='申請',R064005='延展',R064006='移轉',R064007='授權',R064008='變更',R064009='補換發證書',R064010='申請英文證明',R064011='更正' WHERE ID='" & strUserNum & "' AND R064001=1 "
'Modify By Cheng 2003/12/29
'cnnConnection.Execute "UPDATE R020401_1 SET R064004='申請',R064005='延展',R064006='移轉',R064007='授權',R064008='變更',R064009='補換發證書',R064010='申請英文證明',R064011='設定質權' WHERE ID='" & strUserNum & "' AND R064001=1 "
'cnnConnection.Execute "UPDATE R020401_1 SET R064004='申請（件）',R064005='申請（類）',R064006='延展',R064007='移轉',R064008='授權',R064009='變更',R064010='補換發證書',R064011='申請英文證明' WHERE ID='" & strUserNum & "' AND R064001=1 "
cnnConnection.Execute "UPDATE R020401_1 SET R064004='" & Title_101 & "（件）',R064005='" & Title_101 & "（類）',R064006='" & Title_102 & "',R064007='" & Title_501 & "',R064008='" & Title_502 & "',R064009='" & Title_301 & "',R064010='" & Title_103 & "',R064011='" & Title_304 & "' WHERE ID='" & strUserNum & "' AND R064001=1 "
'End
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
BolEndThisPage = False
strSql = "SELECT R065001,NVL(A0902,A0903),nvl(st02,R065003),R065004,R065005,R065007,R065008,R065010,R065011,R065013,R065014,R065016,R065017,R065019,R065020,R065022,R065023,R065025,R065026,R065028,R065029,R065002,r065003 FROM R020401_2,ACC090,staff  WHERE ID='" & strUserNum & "' and r065003=st01(+) AND r065002=A0901(+) ORDER BY R065001,R065002,R065003 "
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
        SavDay3 = CheckStr(.Fields(21))
        strSql = "SELECT R064001,R064004,R064005,R064006,R064007,R064008,R064009,R064010,R064011 FROM R020401_1 WHERE ID='" & strUserNum & "' AND R064001=" & Val(SavDay1) & " "
        CheckOC2
        adoRecordset1.CursorLocation = adUseClient
        adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
            For i = 1 To 8
                StrTemp7(i) = StrToStr(CheckStr(adoRecordset1.Fields(i)), 7)
            Next i
            StrTemp7(9) = ""
                     For i = 1 To 9
                         If Len(StrTemp7(i)) = 0 Then
                             If i <> 1 Then
                                 StrTemp7(i) = "合計"
                             End If
                             If i <> 9 Then
                                 If i = 1 Then
                                    StrTemp7(i) = "總計"
                                 Else
                                    StrTemp7(i + 1) = "總計"
                                 End If
                             End If
                             Exit For
                         End If
                     Next i
         End If
        CheckOC2
        PrintTitle
        PrintTitle2
        SavDay1 = "  "
        SavDay2 = " "
        SavDay3 = "     "
        Do While .EOF = False
            For i = 0 To 21
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
                strSql = "SELECT R064001,R064004,R064005,R064006,R064007,R064008,R064009,R064010,R064011 FROM R020401_1 WHERE ID='" & strUserNum & "' AND R064001=" & Val(CheckStr(.Fields(0))) & " "
                CheckOC2
                adoRecordset1.CursorLocation = adUseClient
                adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                    For i = 1 To 8
                        StrTemp7(i) = StrToStr(CheckStr(adoRecordset1.Fields(i)), 7)
                    Next i
                    StrTemp7(9) = ""
                     For i = 1 To 9
                         If Len(StrTemp7(i)) = 0 Then
                             If i <> 1 Then
                                 StrTemp7(i) = "合計"
                             End If
                             If i <> 9 Then
                                 If i = 1 Then
                                    StrTemp7(i) = "總計"
                                 Else
                                    StrTemp7(i + 1) = "總計"
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
            If SavDay2 <> strTemp(1) Or StrToStr(SavDay3, 1) <> StrToStr(strTemp(21), 1) Then
                If Len(Trim(SavDay2)) <> 0 Then
                  ShowLine2
                  PrintEnd2 (0)
                End If
                SavDay2 = strTemp(1)
                If Len(Trim(SavDay3)) <> 0 Then
                  If StrToStr(SavDay3, 1) <> StrToStr(strTemp(21), 1) Then
                       ShowLine2
                     PrintEnd2 (1)
                  End If
                  ShowLine2
                End If
                SavDay3 = strTemp(21)
                
            Else
               strTemp(1) = ""
            End If
            Else
            SavDay1 = strTemp(0)
            SavDay2 = strTemp(1)
            SavDay3 = strTemp(21)
            SavDay4 = strTemp(1)
            SavDay5 = strTemp(21)
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
    Else
         'Add By Sindy 2011/3/1
         ShowNoData
         Exit Sub
         '2011/3/1 End
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
BolEndThisPage = False
strSql = "SELECT R065001,NVL(A0902,A0903),nvl(st02,R065003),R065004,R065005,R065006,R065007,R065008,R065009,R065010,R065011,R065012,R065013,R065014,R065015,R065016,R065017,R065018,R065019,R065020,R065021,R065022,R065023,R065024,R065025,R065026,R065027,R065028,R065029,R065030,R065002,r065003 FROM R020401_2,ACC090,staff WHERE ID='" & strUserNum & "' AND r065003=st01(+) and r065002=A0901(+) ORDER BY R065001,R065002,R065003 "
CheckOC
Page = 1
TestOk = True
'SavDay1 = ""
'SavDay2 = ""
'SavDay3 = ""
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        SavDay1 = CheckStr(.Fields(0))
        SavDay2 = CheckStr(.Fields(1))
        SavDay3 = CheckStr(.Fields(30))
        strSql = "SELECT R064001,R064004,R064005,R064006,R064007,R064008,R064009,R064010,R064011 FROM R020401_1 WHERE ID='" & strUserNum & "' AND R064001=" & Val(SavDay1) & " "
        CheckOC2
        adoRecordset1.CursorLocation = adUseClient
        adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
            For i = 1 To 8
                StrTemp7(i) = StrToStr(CheckStr(adoRecordset1.Fields(i)), 7)
            Next i
            StrTemp7(9) = ""
            For i = 1 To 9
                If Len(StrTemp7(i)) = 0 Then
                    If i <> 1 Then
                        StrTemp7(i) = "合計"
                    End If
                    If i <> 9 Then
                        If i = 1 Then
                           StrTemp7(i) = "總計"
                        Else
                           StrTemp7(i + 1) = "總計"
                        End If
                    End If
                    Exit For
                End If
            Next i
        End If
        CheckOC2
        PrintTitle
        PrintTitle1
        SavDay1 = ""
         SavDay2 = ""
         SavDay3 = ""

        Do While .EOF = False
            For i = 0 To 30
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
                strSql = "SELECT R064001,R064004,R064005,R064006,R064007,R064008,R064009,R064010,R064011 FROM R020401_1 WHERE ID='" & strUserNum & "' AND R064001=" & Val(CheckStr(.Fields(0))) & " "
                CheckOC2
                adoRecordset1.CursorLocation = adUseClient
                adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                    For i = 1 To 8
                        StrTemp7(i) = StrToStr(CheckStr(adoRecordset1.Fields(i)), 7)
                    Next i
                    StrTemp7(9) = ""
                     For i = 1 To 9
                         If Len(StrTemp7(i)) = 0 Then
                             If i <> 1 Then
                                 StrTemp7(i) = "合計"
                             End If
                             If i <> 9 Then
                                 If i = 1 Then
                                    StrTemp7(i) = "總計"
                                 Else
                                    StrTemp7(i + 1) = "總計"
                                 End If
                             End If
                             Exit For
                         End If
                     Next i
                 End If
                
                Page = Page + 1
                SavDay1 = strTemp(0)
                Printer.NewPage
                PrintTitle
                PrintTitle1
                CheckOC2
                SavDay2 = " "
                SavDay3 = " "
            End If
            'edit by nick 2004/10/06
            'If SavDay2 <> strTemp(1) Then
            If SavDay2 <> strTemp(1) Or StrToStr(SavDay3, 1) <> StrToStr(strTemp(30), 1) Then
                If Len(Trim(SavDay2)) <> 0 Then
                  ShowLine1
                  PrintEnd1 (0)
                End If
                SavDay2 = strTemp(1)
                If Len(Trim(SavDay3)) <> 0 Then
                  If StrToStr(SavDay3, 1) <> StrToStr(strTemp(30), 1) Then
                     ShowLine1
                     PrintEnd1 (1)
                  End If
                  ShowLine1
                End If
                SavDay3 = strTemp(30)
            Else
               strTemp(1) = ""
            End If
            Else
               SavDay1 = strTemp(0)
               SavDay2 = strTemp(1)
               SavDay3 = strTemp(30)
               SavDay4 = strTemp(1)
               SavDay5 = strTemp(30)
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
    Else
         'Add By Sindy 2011/3/1
         ShowNoData
         Exit Sub
         '2011/3/1 End
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
     strSql = "select '','各區小計','',sum(r065004),sum(r065005),sum(r065006),sum(r065007),sum(r065008),sum(r065009),sum(r065010),sum(r065011),sum(r065012),sum(r065013),sum(r065014),sum(r065015),sum(r065016),sum(r065017),sum(r065018),sum(r065019),sum(r065020),sum(r065021),sum(r065022),sum(r065023),sum(r065024),sum(r065025),sum(r065026),sum(r065027),sum(r065028),sum(r065029),sum(r065030) from r020401_2 where id='" & strUserNum & "' AND R065002='" & SavDay3 & "' AND R065001='" & SavDay1 & "' "
Case 1
     strSql = "select '','各所小計','',sum(r065004),sum(r065005),sum(r065006),sum(r065007),sum(r065008),sum(r065009),sum(r065010),sum(r065011),sum(r065012),sum(r065013),sum(r065014),sum(r065015),sum(r065016),sum(r065017),sum(r065018),sum(r065019),sum(r065020),sum(r065021),sum(r065022),sum(r065023),sum(r065024),sum(r065025),sum(r065026),sum(r065027),sum(r065028),sum(r065029),sum(r065030) from r020401_2 where id='" & strUserNum & "' AND SUBSTR(R065002,1,2)='" & StrToStr(SavDay3, 1) & "' AND R065001='" & SavDay1 & "' "
Case 2
     strSql = "select '','全所總計','',sum(r065004),sum(r065005),sum(r065006),sum(r065007),sum(r065008),sum(r065009),sum(r065010),sum(r065011),sum(r065012),sum(r065013),sum(r065014),sum(r065015),sum(r065016),sum(r065017),sum(r065018),sum(r065019),sum(r065020),sum(r065021),sum(r065022),sum(r065023),sum(r065024),sum(r065025),sum(r065026),sum(r065027),sum(r065028),sum(r065029),sum(r065030) from r020401_2 where id='" & strUserNum & "' AND R065001=1 "
     BolEndThisPage = True
Case 3
     strSql = "select '','全所總計','',sum(r065004),sum(r065005),sum(r065006),sum(r065007),sum(r065008),sum(r065009),sum(r065010),sum(r065011),sum(r065012),sum(r065013),sum(r065014),sum(r065015),sum(r065016),sum(r065017),sum(r065018),sum(r065019),sum(r065020),sum(r065021),sum(r065022),sum(r065023),sum(r065024),sum(r065025),sum(r065026),sum(r065027),sum(r065028),sum(r065029),sum(r065030) from r020401_2 where id='" & strUserNum & "' AND R065001=" & Val(SavDay1) & " "
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
            For i = 2 To 26 Step 3
                Printer.Line (PLeft(i) - 50, iPrint - 150)-(PLeft(i) - 50, iPrint + 450)
            Next i
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print CheckStr(.Fields(1))
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print CheckStr(.Fields(2))
            For i = 2 To 28
               If Len(Trim(StrTemp7(Int((i + 1) / 3)))) <> 0 Then
                   Select Case i
                   Case 3, 6, 9, 12, 15, 18, 21, 24, 27
                        Printer.CurrentX = PLeft(i) + 700 - Printer.TextWidth(Format(CheckNum(.Fields(i + 1)) & " ", "#####0.0"))
                        Printer.CurrentY = iPrint
                        Printer.Print Format(CheckNum(.Fields(i + 1)) & " ", "#####0.0")
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
     strSql = "select '','各區小計','',sum(r065004),sum(r065005),sum(r065007),sum(r065008),sum(r065010),sum(r065011),sum(r065013),sum(r065014),sum(r065016),sum(r065017),sum(r065019),sum(r065020),sum(r065022),sum(r065023),sum(r065025),sum(r065026),sum(r065028),sum(r065029) from r020401_2 where id='" & strUserNum & "' AND R065002='" & SavDay3 & "' AND R065001='" & SavDay1 & "' "
Case 1
     strSql = "select '','各所小計','',sum(r065004),sum(r065005),sum(r065007),sum(r065008),sum(r065010),sum(r065011),sum(r065013),sum(r065014),sum(r065016),sum(r065017),sum(r065019),sum(r065020),sum(r065022),sum(r065023),sum(r065025),sum(r065026),sum(r065028),sum(r065029) from r020401_2 where id='" & strUserNum & "' AND SUBSTR(R065002,1,2)='" & StrToStr(SavDay3, 1) & "' AND R065001='" & SavDay1 & "' "
Case 2
     strSql = "select '','全所總計','',sum(r065004),sum(r065005),sum(r065007),sum(r065008),sum(r065010),sum(r065011),sum(r065013),sum(r065014),sum(r065016),sum(r065017),sum(r065019),sum(r065020),sum(r065022),sum(r065023),sum(r065025),sum(r065026),sum(r065028),sum(r065029) from r020401_2 where id='" & strUserNum & "' AND R065001=1 "
     BolEndThisPage = True
Case 3
     strSql = "select '','全所總計','',sum(r065004),sum(r065005),sum(r065007),sum(r065008),sum(r065010),sum(r065011),sum(r065013),sum(r065014),sum(r065016),sum(r065017),sum(r065019),sum(r065020),sum(r065022),sum(r065023),sum(r065025),sum(r065026),sum(r065028),sum(r065029) from r020401_2 where id='" & strUserNum & "' AND R065001=" & Val(SavDay1) & " "
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
            For i = 2 To 18 Step 2
                Printer.Line (PLeft(i) - 50, iPrint - 150)-(PLeft(i) - 50, iPrint + 450)
            Next i
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print CheckStr(.Fields(1))
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print CheckStr(.Fields(2))
            For i = 2 To 19
               If Len(Trim(StrTemp7(Int((i) / 2)))) <> 0 Then
                   Select Case i
                  Case 3, 5, 7, 9, 11, 13, 15, 17, 19
                        Printer.CurrentX = PLeft(i) + 700 - Printer.TextWidth(Format(CheckNum(.Fields(i + 1)) & " ", "#####0.0"))
                        Printer.CurrentY = iPrint
                        Printer.Print Format(CheckNum(.Fields(i + 1)) & " ", "#####0.0")
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
    If Val(SavDay1) = 1 Then
        Printer.Print GetTitleNick & "商申案智權人員收文統計表(固定)"
    Else
        Printer.Print GetTitleNick & "商申案智權人員收文統計表(非固定)"
    End If
Else
    If Val(SavDay1) = 1 Then
        Printer.Print GetTitleNick & "商申案智權人員發文統計表(固定)"
    Else
        Printer.Print GetTitleNick & "商申案智權人員發文統計表(非固定)"
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
Printer.CurrentX = 16500
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300

'Add By Cheng 2002/02/05
Printer.CurrentX = 250
Printer.CurrentY = iPrint
Printer.Print "申請國家：" & "" & Me.txt1(4).Text & " － " & "" & Me.txt1(5).Text
Printer.CurrentX = 6750
Printer.CurrentY = iPrint
Printer.Print "系統類別：" & "" & Me.txt1(0).Text

Printer.CurrentX = 16500
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
End Sub

Sub PrintTitle2()
GetPleft2
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
For i = 2 To 18 Step 2
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
Printer.CurrentX = PLeft1(8) - (Printer.TextWidth(CheckStr(StrTemp7(8))) / 2)
Printer.CurrentY = iPrint
Printer.Print CheckStr(StrTemp7(8))
Printer.CurrentX = PLeft1(9) - (Printer.TextWidth(CheckStr(StrTemp7(9))) / 2)
Printer.CurrentY = iPrint
Printer.Print CheckStr(StrTemp7(9))
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
For k = 2 To 18 Step 2
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
For i = 2 To 19
   If Len(Trim(StrTemp7(Int((i) / 2)))) <> 0 Then
      Select Case i
      Case 3, 5, 7, 9, 11, 13, 15, 17, 19
            Printer.CurrentX = PLeft(i) + 700 - Printer.TextWidth(Format(strTemp(i + 1) & " ", "#####0.0"))
            Printer.CurrentY = iPrint
            Printer.Print Format(strTemp(i + 1) & " ", "#####0.0")
      Case Else
            Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(strTemp(i + 1))
            Printer.CurrentY = iPrint
            Printer.Print strTemp(i + 1)
      End Select
   End If
Next i
For i = 2 To 18 Step 2
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
For i = 3 To 19
    PLeft(i) = 2200 + ((i - 2) * 936)
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
For i = 2 To 26 Step 3
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
Printer.CurrentX = PLeft1(8) - (Printer.TextWidth(CheckStr(StrTemp7(8))) / 2)
Printer.CurrentY = iPrint
Printer.Print CheckStr(StrTemp7(8))
Printer.CurrentX = PLeft1(9) - (Printer.TextWidth(CheckStr(StrTemp7(9))) / 2)
Printer.CurrentY = iPrint
Printer.Print CheckStr(StrTemp7(9))
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
For k = 2 To 26 Step 3
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
For i = 2 To 28
    If Len(Trim(StrTemp7(Int((i + 1) / 3)))) <> 0 Then
      Select Case i
      Case 3, 6, 9, 12, 15, 18, 21, 24, 27
           Printer.CurrentX = PLeft(i) + 700 - Printer.TextWidth(Format(strTemp(i + 1) & " ", "#####0.0"))
           Printer.CurrentY = iPrint
           Printer.Print Format(strTemp(i + 1) & " ", "#####0.0")
      Case Else
           Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(strTemp(i + 1))
           Printer.CurrentY = iPrint
           Printer.Print strTemp(i + 1)
      End Select
   End If
Next i
For i = 2 To 26 Step 3
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
For i = 3 To 28
    PLeft(i) = 2200 + ((i - 2) * 624)
Next i
PLeft1(1) = PLeft(3) + 624 / 2
PLeft1(2) = PLeft(6) + 624 / 2
PLeft1(3) = PLeft(9) + 624 / 2
PLeft1(4) = PLeft(12) + 624 / 2
PLeft1(5) = PLeft(15) + 624 / 2
PLeft1(6) = PLeft(18) + 624 / 2
PLeft1(7) = PLeft(21) + 624 / 2
PLeft1(8) = PLeft(24) + 624 / 2
PLeft1(9) = PLeft(27) + 624 / 2

End Sub

Private Sub Form_Load()

MoveFormToCenter Me
txt1(0) = GetSystemKindByNickT
'txt1(0) = Str020401SysKind

SeekPrintL = Printer.Orientation
PUB_SetPrinter Me.Name, Combo1, , , SeekPrint     'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Printer = Printers(SeekPrint)
Printer.Orientation = SeekPrintL
Set frm020401 = Nothing
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
     strTemp1 = Split(UCase(GetSystemKindByNickT), ",")
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

