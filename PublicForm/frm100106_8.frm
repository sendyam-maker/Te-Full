VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100106_8 
   BorderStyle     =   1  '單線固定
   Caption         =   "台灣新型修正申復未續辦明細"
   ClientHeight    =   4040
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   7030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4040
   ScaleWidth      =   7030
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件進度(&C)"
      Height          =   400
      Index           =   1
      Left            =   4620
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   15
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料(&B)"
      Height          =   400
      Index           =   0
      Left            =   3105
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   15
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   2
      Left            =   5715
      TabIndex        =   0
      Top             =   15
      Width           =   1155
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3465
      Left            =   90
      TabIndex        =   1
      Top             =   465
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   6121
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
End
Attribute VB_Name = "frm100106_8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/05/24 Form2.0已修改: grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/14 日期欄已修改
'add by nickc 2005/08/11 台灣新型修正申復未續辦明細
Option Explicit
Dim strSql As String, i As Integer, j As Integer, s As Integer, k As Integer, strSQL1 As String
Dim StrSQLa As String, intK As Integer, bolSelData As Boolean
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
Dim StrTag As String
'Added by Lydia 2019/11/01 利益衝突案件
Dim m_AllSys As String '預設全部系統別
Dim intCufaCnt As Integer '限閱案件X件

Private Sub cmdok_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
End Sub

Private Sub Form_Load()
bolToEndByNick = False
MoveFormToCenter Me
SetDataListWidth
bolSelData = False
End Sub

Private Sub SetDataListWidth()

Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer

'Modified by Lydia 2019/11/01 +PA26~PA30, PA75
'arrGridHeadText = Array("V", "本所期限", "法定期限", "進度備註", "本所案號", "案件名稱" _
                  , "案件性質", "承辦人", "智權人員", "收文日", "本所期限" _
                  , "申請人", "是否出名", "延期日", "申請國家", "申請人國籍" _
                  , "發文日", "代理人", "彼所案號", "申請案號", "承辦人備註" _
                  , "取消收文日", "", "", "", "", "")
arrGridHeadText = Array("V", "本所期限", "法定期限", "進度備註", "本所案號", "案件名稱" _
                  , "案件性質", "承辦人", "智權人員", "收文日", "本所期限" _
                  , "申請人", "是否出名", "延期日", "申請國家", "申請人國籍" _
                  , "發文日", "代理人", "彼所案號", "申請案號", "承辦人備註" _
                  , "取消收文日", "CP13", "CP09", "CP01", "CP10", "PA26", "PA27", "PA28", "PA29", "PA30", "PA75")
                  
If bolFNation = False Then
   'Modified by Lydia 2019/11/01 +PA26~PA30, PA75
   'arrGridHeadWidth = Array(200, 810, 810, 810, 1300, 800 _
                     , 800, 850, 850, 810, 810 _
                     , 1000, 850, 810, 800, 1000 _
                     , 810, 0, 800, 800, 1000 _
                     , 1035, 0, 0, 0, 0, 0)
   arrGridHeadWidth = Array(200, 810, 810, 810, 1300, 800 _
                     , 800, 850, 850, 810, 810 _
                     , 1000, 850, 810, 800, 1000 _
                     , 810, 0, 800, 800, 1000 _
                     , 1035, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
Else
   'Modified by Lydia 2019/11/01 +PA26~PA30, PA75
   'arrGridHeadWidth = Array(200, 810, 810, 810, 1300, 800 _
                     , 800, 850, 850, 810, 810 _
                     , 1000, 850, 810, 800, 1000 _
                     , 810, 810, 800, 800, 1000 _
                     , 1035, 0, 0, 0, 0, 0)
   arrGridHeadWidth = Array(200, 810, 810, 810, 1300, 800 _
                     , 800, 850, 850, 810, 810 _
                     , 1000, 850, 810, 800, 1000 _
                     , 810, 810, 800, 800, 1000 _
                     , 1035, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
End If
                     
grdDataList.Cols = UBound(arrGridHeadText) + 1
For iRow = 0 To grdDataList.Cols - 1
   grdDataList.row = 0
   grdDataList.col = iRow
   grdDataList.Text = arrGridHeadText(iRow)
   grdDataList.ColWidth(iRow) = arrGridHeadWidth(iRow)
   grdDataList.CellAlignment = flexAlignCenterCenter
Next
End Sub

Public Function Process() As Boolean
Dim dblRow As Double 'Add By Sindy 2025/9/3
   
      bolSelData = False
      
      'Added by Lydia 2019/11/01
      m_AllSys = IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(, frm100106_1.txt5(0).Text))
      intCufaCnt = 0
      'end 2019/11/01
      
      ClearQueryLog ("frm100106_1") 'Add By Sindy 2010/11/3 清除查詢印表記錄檔欄位
      strSQL1 = ""
      If frm100106_1.opt1(0).Value = True Then
         'Modify by Morgan 2010/3/11 會有延期問題,故改抓下一程序期限(Ex.P-91870)
         'strSQL1 = strSQL1 + " and CP06>=" & Val(ChangeTStringToWString(frm100106_1.txt1(0))) & " "
         'strSQL1 = strSQL1 + " and CP06<=" & Val(ChangeTStringToWString(frm100106_1.txt1(1))) & " "
         strSQL1 = strSQL1 + " and NP08>=" & Val(ChangeTStringToWString(frm100106_1.txt1(0))) & " "
         strSQL1 = strSQL1 + " and NP08<=" & Val(ChangeTStringToWString(frm100106_1.txt1(1))) & " "
         pub_QL05 = pub_QL05 & ";" & frm100106_1.opt1(0).Caption & frm100106_1.txt1(0) & "-" & frm100106_1.txt1(1) 'Add By Sindy 2010/11/3
      ElseIf frm100106_1.opt1(1).Value = True Then
         'Modify by Morgan 2010/3/11 會有延期問題,故改抓下一程序期限(Ex.P-91870)
         'strSQL1 = strSQL1 + " and CP07>=" & Val(ChangeTStringToWString(frm100106_1.txt2(0))) & " "
         'strSQL1 = strSQL1 + " and CP07<=" & Val(ChangeTStringToWString(frm100106_1.txt2(1))) & " "
         strSQL1 = strSQL1 + " and NP09>=" & Val(ChangeTStringToWString(frm100106_1.txt2(0))) & " "
         strSQL1 = strSQL1 + " and NP09<=" & Val(ChangeTStringToWString(frm100106_1.txt2(1))) & " "
         pub_QL05 = pub_QL05 & ";" & frm100106_1.opt1(1).Caption & frm100106_1.txt2(0) & "-" & frm100106_1.txt2(1) 'Add By Sindy 2010/11/3
      ElseIf frm100106_1.opt1(2).Value = True Then
         strSQL1 = strSQL1 + " and CP01='" & frm100106_1.txt3(0) & "' and cp02='" & frm100106_1.txt3(1) & "' and cp03='" & IIf(Trim(frm100106_1.txt3(2)) = "", "0", frm100106_1.txt3(2)) & "' and cp04='" & IIf(Trim(frm100106_1.txt3(3)) = "", "00", frm100106_1.txt3(3)) & "' "
         pub_QL05 = pub_QL05 & ";" & frm100106_1.opt1(2).Caption & frm100106_1.txt3(0) & "-" & frm100106_1.txt3(1) & "-" & frm100106_1.txt3(2) & "-" & frm100106_1.txt3(3) 'Add By Sindy 2010/11/3
      End If
      
      '承辦人
      If Len(Trim(frm100106_1.txt5(1))) <> 0 Then
         strSQL1 = strSQL1 & " AND CP14||''='" & frm100106_1.txt5(1) & "' "
         pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(2) & frm100106_1.txt5(1) & frm100106_1.LBL1(0) 'Add By Sindy 2010/11/3
      End If
      '業務區
      If Len(Trim(frm100106_1.txt5(2))) <> 0 Then
         strSQL1 = strSQL1 & " AND cp12>='" & frm100106_1.txt5(2) & "' "
      End If
      If Len(Trim(frm100106_1.txt5(3))) <> 0 Then
         strSQL1 = strSQL1 & " AND cp12<='" & frm100106_1.txt5(3) & "' "
      End If
      If Len(Trim(frm100106_1.txt5(2))) <> 0 Or Len(Trim(frm100106_1.txt5(3))) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(3) & frm100106_1.txt5(2) & "-" & frm100106_1.txt5(3) 'Add By Sindy 2010/11/3
      End If
      '智權人員
      If Len(Trim(frm100106_1.txt5(4))) <> 0 Then
         strSQL1 = strSQL1 + " AND CP13||''='" & frm100106_1.txt5(4) & "' "
         pub_QL05 = pub_QL05 & ";" & frm100106_1.Label2(0) & frm100106_1.txt5(4) & frm100106_1.LBL1(1) 'Add By Sindy 2010/11/3
      End If
      '申請人國籍
      If Len(Trim(frm100106_1.txt5(9))) <> 0 Then
         strSQL1 = strSQL1 & " AND CU10 >='" & frm100106_1.txt5(9) & "' "
      End If
      If Len(Trim(frm100106_1.txt5(10))) <> 0 Then
         strSQL1 = strSQL1 + " AND CU10 <='" & frm100106_1.txt5(10) & "' "
      End If
      If Len(Trim(frm100106_1.txt5(9))) <> 0 Or Len(Trim(frm100106_1.txt5(10))) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(1) & frm100106_1.txt5(9) & "-" & frm100106_1.txt5(10) 'Add By Sindy 2010/11/3
      End If
      '系統類別
      If Len(Trim(frm100106_1.txt5(0))) <> 0 Then
         strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(frm100106_1.txt5(0))), 1) & ") "
         pub_QL05 = pub_QL05 & ";" & Left(frm100106_1.Label1(4), 5) & frm100106_1.txt5(0)  'Add By Sindy 2010/11/3
      End If
      '申請人
      If Len(Trim(frm100106_1.txt4(0))) <> 0 Then
         strSQL1 = strSQL1 & " AND PA26>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
      End If
      If Len(Trim(frm100106_1.txt4(1))) <> 0 Then
         strSQL1 = strSQL1 & " AND PA26<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
      End If
      If Len(Trim(frm100106_1.txt4(0))) <> 0 Or Len(Trim(frm100106_1.txt4(1))) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(6) & frm100106_1.txt4(0) & "-" & frm100106_1.txt4(1) 'Add By Sindy 2010/11/3
      End If
      '代理人
      If Len(Trim(frm100106_1.txt4(2))) <> 0 Then
         strSQL1 = strSQL1 & " AND PA75>='" & GetNewFagent(frm100106_1.txt4(2)) & "' "
      End If
      If Len(Trim(frm100106_1.txt4(3))) <> 0 Then
         strSQL1 = strSQL1 & " AND PA75<='" & GetNewFagent(frm100106_1.txt4(3)) & "' "
      End If
      If Len(Trim(frm100106_1.txt4(2))) <> 0 Or Len(Trim(frm100106_1.txt4(3))) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(5) & frm100106_1.txt4(2) & "-" & frm100106_1.txt4(3) 'Add By Sindy 2010/11/3
      End If
      'FCP管制人
      If Len(Trim(frm100106_1.txt5(11).Text)) <> 0 Then
         'Modified by Lydia 2017/02/13 +FMP管制人
         If strSrvDate(1) < FMP管制人啟用日 Then
            strSQL1 = strSQL1 & " AND DECODE(PA75,NULL,N2.NA16,N3.NA16) >='" & frm100106_1.txt5(11).Text & "' "
         Else
            strSQL1 = strSQL1 & " AND DECODE(PA01,'P',DECODE(PA75,NULL,NVL(N2.NA79,N2.NA16),NVL(N3.NA79,N3.NA16)),DECODE(PA75,NULL,N2.NA16,N3.NA16)) >='" & frm100106_1.txt5(11).Text & "' "
         End If
         'end 2017/02/13
      End If
      If Len(Trim(frm100106_1.txt5(12).Text)) <> 0 Then
         'Modified by Lydia 2017/02/13 +FMP管制人
         If strSrvDate(1) < FMP管制人啟用日 Then
            strSQL1 = strSQL1 + " AND DECODE(PA75,NULL,N2.NA16,N3.NA16) <='" & frm100106_1.txt5(12).Text & "' "
         Else
            strSQL1 = strSQL1 + " AND DECODE(PA01,'P',DECODE(PA75,NULL,NVL(N2.NA79,N2.NA16),NVL(N3.NA79,N3.NA16)),DECODE(PA75,NULL,N2.NA16,N3.NA16)) >='" & frm100106_1.txt5(12).Text & "' "
         End If
         'end 2017/02/13
      End If
      If Len(Trim(frm100106_1.txt5(11).Text)) <> 0 Or Len(Trim(frm100106_1.txt5(12).Text)) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(8) & frm100106_1.txt5(11) & "-" & frm100106_1.txt5(12) 'Add By Sindy 2010/11/3
      End If
      '2005/11/15 MODIFY BY SONIA 不應有PA16的限制,只抓P之台灣新型
      'strSQL1 = strSQL1 & " and cp10 in ('1201','1202') and np06 is null and pa09='000' and '2'=pa08(+) and pa10>=20040701 and pa01 in ('P','CFP','FCP') and pa16='2' "
      'Modify by Morgan 2010/3/11 會有延期問題,故改抓下一程序期限(Ex.P-91870)
      'strSQL1 = strSQL1 & " and cp10 in ('1201','1202') and np06 is null and pa09='000' and '2'=pa08(+) and pa10>=20040701 and pa01='P' "
      strSQL1 = strSQL1 & " and cp10 in ('1201','1202') and np06 is null and pa09='000' and '2'=pa08(+) and pa10>=20040701 and pa01='P' and np02||''='P' "
      'end 2010/3/11
      '2005/11/15 END
      
      CheckOC
      StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人,"
      
      'Modify by Morgan 2010/3/11 會有延期問題,故改抓下一程序期限(Ex.P-91870)
      'strSql = "SELECT '' AS V," & SQLDate("cP06") & " AS 本所期限," & SQLDate("cP07") & " AS 法定期限,CP64 AS 進度備註,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員," & SQLDate("CP05") & " AS 收文日," & SQLDate("cP06") & " AS 本所期限, SUBSTRB(NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),1,10) AS 申請人,DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍," & SQLDate("CP27") & " AS 發文日," & StrSQLa & "DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號,ep12 AS 承辦人備註," & SQLDate("CP57") & " AS 取消收文日,cp13,cp09, CP01, CP10 " & _
               " FROM CASEPROGRESS,PATENT,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND,nextprogress " & _
                " WHERE (PA57<>'Y' or pa57 is null) AND CP01=PA01(+) AND Cp02=PA02(+) AND Cp03=PA03(+) AND Cp04=PA04(+) AND cP09=eP02(+) and cp09=np01(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTr(PA75,9,1),NULL,'0',SUBSTR(PA75,9,1))=FA02(+) AND CP01=CPM01(+) AND cp10=CPM02(+) AND PA09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL1 & _
                " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND PA09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND PA09 <= '" & frm100106_1.txt5(8).Text & "'", "")
      '2010/9/14 MODIFY BY SONIA 日期欄改百年日期排序問題
      'Modified by Lydia 2019/11/01 利益衝突案件：於c004後面，增加申請人1~5,FC代理人
      'strSql = "SELECT '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,CP64 AS 進度備註,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP06),-9) AS 本所期限, SUBSTRB(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),1,10) AS 申請人,DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP01, CP10 " & _
                " FROM nextprogress,CASEPROGRESS,PATENT,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                " WHERE CP09(+)=NP01 AND (PA57<>'Y' or pa57 is null) AND CP01=PA01(+) AND Cp02=PA02(+) AND Cp03=PA03(+) AND Cp04=PA04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTr(PA75,9,1),NULL,'0',SUBSTR(PA75,9,1))=FA02(+) AND CP01=CPM01(+) AND cp10=CPM02(+) AND PA09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL1 & _
                " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND PA09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND PA09 <= '" & frm100106_1.txt5(8).Text & "'", "")
      'Modified by Lydia 2021/05/24 限制欄位長度: CP64 AS 進度備註=> substr(CP64,1,500) AS 進度備註、ep12 AS 承辦人備註 =>substr(ep12,1,500) AS 承辦人備註
      strSql = "SELECT '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,substr(CP64,1,500) AS 進度備註,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員," & _
                "SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP06),-9) AS 本所期限, SUBSTRB(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),1,10) AS 申請人,DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍," & _
                "SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號,substr(ep12,1,500) AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP01, CP10 " & _
                ", PA26 , PA27 , PA28 , PA29 , PA30 , PA75" & _
                " FROM nextprogress,CASEPROGRESS,PATENT,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                " WHERE CP09(+)=NP01 AND (PA57<>'Y' or pa57 is null) AND CP01=PA01(+) AND Cp02=PA02(+) AND Cp03=PA03(+) AND Cp04=PA04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) " & _
                "and cp14=ep05(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTr(PA75,9,1),NULL,'0',SUBSTR(PA75,9,1))=FA02(+) AND CP01=CPM01(+) AND cp10=CPM02(+) AND PA09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL1 & _
                " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND PA09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND PA09 <= '" & frm100106_1.txt5(8).Text & "'", "")
      '是否按智權人員排序
      If frm100106_1.txt5(13).Text = "Y" Then
         strSql = strSql & " ORDER BY s2.ST03, s2.ST01, " & IIf(frm100106_1.opt1(0).Value, "2,", IIf(frm100106_1.opt1(1).Value, "3,", "")) & "本所案號 "
         pub_QL05 = pub_QL05 & ";" & Left(frm100106_1.Label1(9), 11) & frm100106_1.txt5(13) 'Add By Sindy 2010/11/3
      Else
         strSql = strSql & " ORDER BY " & IIf(frm100106_1.opt1(0).Value, "2,", IIf(frm100106_1.opt1(1).Value, "3,", "")) & "本所案號 "
      End If
       
      CheckOC
      adoRecordset.CursorLocation = adUseClient
      'Modified by Lydia 2019/11/01 改變型態
      'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
      If adoRecordset.RecordCount <> 0 Then
         dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3

         'Added by Lydia 2019/11/01 逐案號判斷
         If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
            adoRecordset.MoveFirst
            Do While adoRecordset.EOF = False
                '利益衝突案件：逐案號判斷
                If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & adoRecordset.Fields(4), "" & adoRecordset.Fields("pa26") & "," & adoRecordset.Fields("pa27") & "," & adoRecordset.Fields("pa28") & "," & adoRecordset.Fields("pa29") & "," & adoRecordset.Fields("pa30"), "" & adoRecordset.Fields("pa75")) = False Then
                    intCufaCnt = intCufaCnt + 1
                    adoRecordset.Delete
                End If
                adoRecordset.MoveNext
            Loop
            '利益衝突案件：限閱案件
            If intCufaCnt > 0 Then
               pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
               MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
            End If
            InsertQueryLog (dblRow) 'Add By Sindy 2010/11/3
            If adoRecordset.RecordCount = 0 Then
                  GoTo JumpToNoData
            End If
         Else
            InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/11/3
         End If
        'end 2019/11/01
         
         Process = True
      Else
         InsertQueryLog (0) 'Add By Sindy 2010/11/3
JumpToNoData:   'Added by Lydia 2019/11/01
         Screen.MousePointer = vbDefault
         Process = False
         Me.Hide
         Exit Function
      End If
      
      Set grdDataList.Recordset = adoRecordset
      intK = grdDataList.Rows - 1
      CheckOC
      grdDataList.Visible = False
      For i = 1 To grdDataList.Rows - 1
         Me.grdDataList.TextMatrix(i, 6) = Me.grdDataList.TextMatrix(i, 6) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(i, 23), "1")
         grdDataList.row = i
         grdDataList.col = grdDataList.Cols - 3
         strSql = "SELECT " & SQLDate("DL02") & " FROM DATELIMIT WHERE DL01='" & grdDataList.Text & "' ORDER BY DL02"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         grdDataList.col = 13
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            adoRecordset.MoveLast
            If Not IsNull(adoRecordset.Fields(0)) Then
               grdDataList.Text = adoRecordset.Fields(0)
            Else
               grdDataList.Text = ""
            End If
         End If
         grdDataList.col = 16
           If Len(Trim(grdDataList.Text)) = 0 Then
               grdDataList.col = 4
               For j = 1 To grdDataList.Cols - 1
                   grdDataList.col = j
               Next j
           Else
               grdDataList.col = 1
               '2010/9/14 MODIFY BY SONIA 修改百年蟲
               'If ChangeTDateStringToTString(grdDataList.Text) < ChangeWStringToTString(ServerDate) And Trim(grdDataList.Text) <> "" Then
               If Val(ChangeTDateStringToTString(grdDataList.Text)) < Val(ChangeWStringToTString(ServerDate)) And Trim(grdDataList.Text) <> "" Then
                   grdDataList.col = 4
                   grdDataList.Text = "*" + grdDataList.Text
                   For j = 1 To grdDataList.Cols - 1
                       grdDataList.col = j
                       '紅色
                       grdDataList.CellBackColor = &HFF&
                   Next j
               Else
                   '2010/9/20 modify by sonia 因日期加空格故加val
                   If Val(ChangeTDateStringToTString(grdDataList.Text)) = Val(ChangeWStringToTString(ServerDate)) Then
                      grdDataList.col = 4
                      grdDataList.Text = "v" & grdDataList.Text
                       For j = 1 To grdDataList.Cols - 1
                           grdDataList.col = j
                           '橙色
                           grdDataList.CellBackColor = &H80FF&
                       Next j
                   End If
               End If
           End If
      Next i
      grdDataList.Visible = True
End Function

Private Sub Form_Unload(Cancel As Integer)
tmpBol = fnCancelNowFormAndShowParentForm(Me)
Set frm100106_8 = Nothing
End Sub

Private Sub grdDataList_SelChange()
bolSelData = True
grdDataList.Visible = False
grdDataList.row = grdDataList.MouseRow
grdDataList.col = 0
If grdDataList.row <> 0 Then
If grdDataList.Text = "V" Then
     grdDataList.Text = ""
     For i = 0 To grdDataList.Cols - 1
          grdDataList.col = i
          grdDataList.CellBackColor = QBColor(15)
    Next i
   'Add By Cheng 2002/03/15
    bolSelData = False
Else
     grdDataList.Text = "V"
     For i = 0 To grdDataList.Cols - 1
         grdDataList.col = i
         grdDataList.CellBackColor = &HFFC0C0
     Next i

End If
End If
grdDataList.Visible = True
End Sub

Public Sub PubShowNextData()
Dim i As Integer, j As Integer
Select Case cmdState
Case 0 '案件基本資料
      Me.Enabled = False
      For i = 1 To grdDataList.Rows - 1
      grdDataList.col = 0
      grdDataList.row = i
      If Trim(grdDataList.Text) = "V" Then
        Dim Str01 As String
        grdDataList.col = 0
        grdDataList.Text = ""
        For j = 0 To grdDataList.Cols - 1
            grdDataList.col = j
            grdDataList.CellBackColor = QBColor(15)
        Next j
        grdDataList.col = 4
        Str01 = SystemNumber(grdDataList, 1)
        If Mid(UCase(Str01), 1, 1) = "N" Then
            Str01 = Mid(Str01, 2, 3)
        End If
        If Not IsNull(grdDataList.Text) Then
'            If fnSaveParentForm(Me) = False Then
'                Me.Enabled = True
'                Exit Sub
'            End If
            If InStr(1, ArrFormByNick, Trim(Me.Name)) = 0 Then
                ArrFormByNick = ArrFormByNick & ";" & Trim(Me.Name)
                If Left(ArrFormByNick, 1) = ";" Then
                    ArrFormByNick = Mid(ArrFormByNick, 2)
                End If
            End If
            Select Case Pub_RplStr(Str01)
                Case "CFP", "FCP", "P"   '專利
                      Screen.MousePointer = vbHourglass
                      frm100101_3.Show
                      frm100101_3.Tag = Pub_RplStr(grdDataList.Text)
                      frm100101_3.StrMenu
                      Screen.MousePointer = vbDefault
                Case "CFT", "FCT", "T", "TF"   '商標
                      Screen.MousePointer = vbHourglass
                      frm100101_4.Show
                      frm100101_4.Tag = Pub_RplStr(grdDataList.Text)
                      frm100101_4.StrMenu
                      Screen.MousePointer = vbDefault
                'Modify By Sindy 2009/07/24 增加LIN系統類別
                'modify by sonia 2019/7/29 +ACS系統類別
                Case "CFL", "FCL", "L", "LIN", "ACS"  '法務
                      Screen.MousePointer = vbHourglass
                      frm100101_5.Show
                      frm100101_5.Tag = Pub_RplStr(grdDataList.Text)
                      frm100101_5.StrMenu
                      Screen.MousePointer = vbDefault
                Case "LA"            '顧問
                      Screen.MousePointer = vbHourglass
                      frm100101_6.Show
                      frm100101_6.Tag = Pub_RplStr(grdDataList.Text)
                      frm100101_6.StrMenu
                      Screen.MousePointer = vbDefault
                Case Else                  '服務
                     Select Case Pub_RplStr(Str01)
                         Case "TB"    '條碼
                            Screen.MousePointer = vbHourglass
                            frm100101_7.Show
                            frm100101_7.Tag = Pub_RplStr(grdDataList.Text)
                            frm100101_7.StrMenu
                            Screen.MousePointer = vbDefault
                         Case "TM"
                            Screen.MousePointer = vbHourglass
                            frm100101_8.Show
                            frm100101_8.Tag = Pub_RplStr(grdDataList.Text)
                            frm100101_8.StrMenu
                            Screen.MousePointer = vbDefault
                         Case "TD"
                            Screen.MousePointer = vbHourglass
                            frm100101_9.Show
                            frm100101_9.Tag = Pub_RplStr(grdDataList.Text)
                            frm100101_9.StrMenu
                            Screen.MousePointer = vbDefault
                         Case "TC", "CFC"
                            Screen.MousePointer = vbHourglass
                            frm100101_A.Show
                            frm100101_A.Tag = Pub_RplStr(grdDataList.Text)
                            frm100101_A.StrMenu
                            Screen.MousePointer = vbDefault
                         Case Else
                            Screen.MousePointer = vbHourglass
                            frm100101_B.Show
                            frm100101_B.Tag = Pub_RplStr(grdDataList.Text)
                            frm100101_B.StrMenu
                            Screen.MousePointer = vbDefault
                      End Select
            End Select
        End If
        Me.Enabled = True
        Exit Sub
     End If
     Next i
     Me.Enabled = True
Case 1 '案件進度
     Me.Enabled = False
     StrTag = ""
     For i = 1 To grdDataList.Rows - 1
     grdDataList.col = 0
     grdDataList.row = i
     If Trim(grdDataList.Text) = "V" Then
        grdDataList.col = 0
        grdDataList.Text = ""
        For j = 0 To grdDataList.Cols - 1
            grdDataList.col = j
            grdDataList.CellBackColor = QBColor(15)
        Next j
         grdDataList.col = 4
         If Not IsNull(grdDataList.Text) Then
'            If fnSaveParentForm(Me) = False Then
'                Me.Enabled = True
'                Exit Sub
'            End If
            If InStr(1, ArrFormByNick, Trim(Me.Name)) = 0 Then
                ArrFormByNick = ArrFormByNick & ";" & Trim(Me.Name)
                If Left(ArrFormByNick, 1) = ";" Then
                    ArrFormByNick = Mid(ArrFormByNick, 2)
                End If
            End If
            Screen.MousePointer = vbHourglass
            frm100101_2.Show
            frm100101_2.Tag = Pub_RplStr(grdDataList.Text)
            frm100101_2.StrMenu
            Screen.MousePointer = vbDefault

            Me.Enabled = True
            Exit Sub
         End If
     End If
     Next i
     Me.Enabled = True
Case 2
      Me.Hide
Case Else
End Select
End Sub
