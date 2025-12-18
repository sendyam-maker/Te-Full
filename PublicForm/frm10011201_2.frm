VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm10011201_2 
   Appearance      =   0  '平面
   BackColor       =   &H80000000&
   BorderStyle     =   1  '單線固定
   Caption         =   "後金案件及結果查詢"
   ClientHeight    =   5730
   ClientLeft      =   15
   ClientTop       =   945
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   7272
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   8496
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   70
      Width           =   756
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   5076
      Left            =   72
      TabIndex        =   2
      Top             =   624
      Width           =   9204
      _ExtentX        =   16245
      _ExtentY        =   8943
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      HighLight       =   0
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
      _Band(0).Cols   =   10
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   1
      Left            =   4524
      TabIndex        =   6
      Top             =   432
      Width           =   2532
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   0
      Left            =   924
      TabIndex        =   5
      Top             =   432
      Width           =   2532
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "結果日 :"
      Height          =   180
      Left            =   3690
      TabIndex        =   4
      Top             =   435
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "收文日 :"
      Height          =   180
      Left            =   90
      TabIndex        =   3
      Top             =   435
      Width           =   630
   End
End
Attribute VB_Name = "frm10011201_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/07 改成Form2.0 ; grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/14 日期欄已修改
Option Explicit
Dim strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String, strSQL1 As String
Dim strSql As String, i As Integer, j As Integer, intK As Integer, strTemp As Variant, strTemp1 As String, s As Integer, StrTest As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer

Private Sub SetDataListWidth()
grdDataList.Cols = 11
grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "本所案號"
grdDataList.ColWidth(0) = 1550
grdDataList.CellAlignment = flexAlignCenterCenter
Dim iDep As String
iDep = PUB_GetST06(strUserNum)
grdDataList.col = 1: grdDataList.Text = "分所號"
'電腦中心，跟分所才秀
If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
    grdDataList.ColWidth(1) = 0
Else
    grdDataList.ColWidth(1) = 620
End If
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 2: grdDataList.Text = "案件名稱"
grdDataList.ColWidth(2) = 1400
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 3: grdDataList.Text = "申請人"
grdDataList.ColWidth(3) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 4: grdDataList.Text = "案件性質"
grdDataList.ColWidth(4) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 5: grdDataList.Text = "承辦人"
grdDataList.ColWidth(5) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 6: grdDataList.Text = "智權人員"
grdDataList.ColWidth(6) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 7: grdDataList.Text = "後金"
grdDataList.ColWidth(7) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 8: grdDataList.Text = "收回日"
grdDataList.ColWidth(8) = 850
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 9: grdDataList.Text = "收回金額"
grdDataList.ColWidth(9) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
'Add By Cheng 2003/08/15
grdDataList.col = 10: grdDataList.Text = "CP09"
grdDataList.ColWidth(10) = 0
grdDataList.CellAlignment = flexAlignCenterCenter
End Sub

'92.04.16 nick
Public Sub PubShowNextData()
Select Case cmdState
Case 0
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 1
     fnCloseAllFrm100
Case Else
End Select
End Sub


Private Sub cmdok_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
'92.04.16 nick 以下無效
Select Case Index
Case 0
     Me.Hide
Case 1
     bolToEndByNick = True
     Unload Me
     Exit Sub
Case Else
End Select
End Sub

Private Sub Form_Load()
bolToEndByNick = False
   MoveFormToCenter Me
SetDataListWidth
'92.04.16 nick
cmdState = -1
End Sub

Sub StrMenu()
Me.Enabled = False
If frm10011201_1.txt1(0) <> "3" Then
    lbl1(0).Caption = frm10011201_1.txt1(1) + "-" + frm10011201_1.txt1(2)
    lbl1(1).Caption = ""
Else
    lbl1(0).Caption = ""
    lbl1(1).Caption = frm10011201_1.txt1(3) + "-" + frm10011201_1.txt1(4)
End If
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
strSQL5 = ""
If frm10011201_1.txt1(0) = "1" Then
    strSQL1 = strSQL1 + " AND CP05>=" & Val(ChangeTStringToWString(frm10011201_1.txt1(1))) & " AND CP05<=" & Val(ChangeTStringToWString(frm10011201_1.txt1(2))) & " "
    pub_QL05 = pub_QL05 & ";" & frm10011201_1.Label1 & "收文" 'Add By Sindy 2010/11/4
    pub_QL05 = pub_QL05 & ";" & frm10011201_1.Label3 & frm10011201_1.txt1(1) & "-" & frm10011201_1.txt1(2) 'Add By Sindy 2010/11/4
Else
    If frm10011201_1.txt1(0) = "2" Then
        strSQL1 = strSQL1 + " AND CP05>=" & Val(ChangeTStringToWString(frm10011201_1.txt1(1))) & " AND CP05<=" & Val(ChangeTStringToWString(frm10011201_1.txt1(2))) & " AND CP25 IS NULL "
        pub_QL05 = pub_QL05 & ";" & frm10011201_1.Label1 & "無結果" 'Add By Sindy 2010/11/4
        pub_QL05 = pub_QL05 & ";" & frm10011201_1.Label3 & frm10011201_1.txt1(1) & "-" & frm10011201_1.txt1(2) 'Add By Sindy 2010/11/4
    Else
        If frm10011201_1.txt1(0) = "3" Then
            strSQL1 = strSQL1 + " AND CP25>=" & Val(ChangeTStringToWString(frm10011201_1.txt1(3))) & " AND CP25<=" & Val(ChangeTStringToWString(frm10011201_1.txt1(4))) & " "
            pub_QL05 = pub_QL05 & ";" & frm10011201_1.Label1 & "有結果" 'Add By Sindy 2010/11/4
            pub_QL05 = pub_QL05 & ";" & frm10011201_1.Label4(0) & frm10011201_1.txt1(3) & "-" & frm10011201_1.txt1(4) 'Add By Sindy 2010/11/4
        End If
    End If
End If

'Modify by Morgan 2009/10/19 FCP 退費908 除外
'strSQL1 = strSQL1 & " AND CP19<>0 "
strSQL1 = strSQL1 & " AND CP19<>0 AND CP01||CP10<>'FCP908'"

strSQL2 = strSQL1
StrSQL3 = strSQL1
StrSQL4 = strSQL1
strSQL5 = strSQL1
If Len(Trim(frm10011201_1.txt1(5))) <> 0 Then
   'Modify By Cheng 2002/03/14
'   strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr(frm10011201_1.txt1(5), 1) & ") "
'   strSQL2 = strSQL2 & " AND CP01 IN (" & SQLGrpStr(frm10011201_1.txt1(5), 2) & ") "
'   StrSQL3 = StrSQL3 & " AND CP01 IN (" & SQLGrpStr(frm10011201_1.txt1(5), 3) & ") "
'   StrSQL4 = StrSQL4 & " AND CP01 IN (" & SQLGrpStr(frm10011201_1.txt1(5), 4) & ") "
'   StrSQL5 = StrSQL5 & " AND CP01 IN (" & SQLGrpStr(frm10011201_1.txt1(5), 5) & ") "
   strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr(IIf(frm10011201_1.txt1(5).Text <> "ALL", frm10011201_1.txt1(5).Text, GetAllSysKind(frm10011201_1.txt1(5))), 1) & ") "
   strSQL2 = strSQL2 & " AND CP01 IN (" & SQLGrpStr(IIf(frm10011201_1.txt1(5).Text <> "ALL", frm10011201_1.txt1(5).Text, GetAllSysKind(frm10011201_1.txt1(5))), 2) & ") "
   StrSQL3 = StrSQL3 & " AND CP01 IN (" & SQLGrpStr(IIf(frm10011201_1.txt1(5).Text <> "ALL", frm10011201_1.txt1(5).Text, GetAllSysKind(frm10011201_1.txt1(5))), 3) & ") "
   StrSQL4 = StrSQL4 & " AND CP01 IN (" & SQLGrpStr(IIf(frm10011201_1.txt1(5).Text <> "ALL", frm10011201_1.txt1(5).Text, GetAllSysKind(frm10011201_1.txt1(5))), 4) & ") "
   strSQL5 = strSQL5 & " AND CP01 IN (" & SQLGrpStr(IIf(frm10011201_1.txt1(5).Text <> "ALL", frm10011201_1.txt1(5).Text, GetAllSysKind(frm10011201_1.txt1(5))), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Left(frm10011201_1.Label4(1), 5) & frm10011201_1.txt1(5) 'Add By Sindy 2010/11/4
End If
'Modify By Cheng 2002/04/25
'edit by nick 2004/11/29
'                    StrSql = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU88||' '||CU89,CU06)),PA26) AS 申請人,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP19 AS 後金,CP09 AS 收回日,CP09 AS 收回金額, CP09 FROM CASEPROGRESS,PATENT,CUSTOMER,STAFF S1,STAFF S2,CASEPROPERTYMAP          WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND " & SQLNewFag("PA26", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1
'StrSql = StrSql & " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU88||' '||CU89,CU06)),TM23) AS 申請人,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP19 AS 後金,CP09 AS 收回日,CP09 AS 收回金額, CP09 FROM CASEPROGRESS,TRADEMARK,CUSTOMER,STAFF S1,STAFF S2,CASEPROPERTYMAP       WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND " & SQLNewFag("TM23", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2
'StrSql = StrSql & " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU88||' '||CU89,CU06)),LC11) AS 申請人,NVL(DECODE(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP19 AS 後金,CP09 AS 收回日,CP09 AS 收回金額, CP09 FROM CASEPROGRESS,LAWCASE,CUSTOMER,STAFF S1,STAFF S2,CASEPROPERTYMAP         WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND " & SQLNewFag("LC11", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3
'StrSql = StrSql & " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU88||' '||CU89,CU06)),HC05) AS 申請人,NVL(CPM03,CP10)                          AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP19 AS 後金,CP09 AS 收回日,CP09 AS 收回金額, CP09 FROM CASEPROGRESS,HIRECASE,CUSTOMER,STAFF S1,STAFF S2,CASEPROPERTYMAP        WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND " & SQLNewFag("HC05", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4
'StrSql = StrSql & " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU88||' '||CU89,CU06)),SP08) AS 申請人,NVL(DECODE(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP19 AS 後金,CP09 AS 收回日,CP09 AS 收回金額, CP09 FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,STAFF S1,STAFF S2,CASEPROPERTYMAP WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND " & SQLNewFag("SP08", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5
                    strSql = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),PA26) AS 申請人,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP19 AS 後金,decode(substr(cp09,1,1),'A',cp60,'B',cp60,CP09) AS 收回日,CP75 AS 收回金額, CP09 FROM CASEPROGRESS,PATENT,CUSTOMER,STAFF S1,STAFF S2,CASEPROPERTYMAP          WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND " & SQLNewFag("PA26", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1
strSql = strSql & " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP19 AS 後金,decode(substr(cp09,1,1),'A',cp60,'B',cp60,CP09) AS 收回日,CP75 AS 收回金額, CP09 FROM CASEPROGRESS,TRADEMARK,CUSTOMER,STAFF S1,STAFF S2,CASEPROPERTYMAP       WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND " & SQLNewFag("TM23", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2
strSql = strSql & " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),LC11) AS 申請人,NVL(DECODE(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP19 AS 後金,decode(substr(cp09,1,1),'A',cp60,'B',cp60,CP09) AS 收回日,CP75 AS 收回金額, CP09 FROM CASEPROGRESS,LAWCASE,CUSTOMER,STAFF S1,STAFF S2,CASEPROPERTYMAP         WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND " & SQLNewFag("LC11", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3
strSql = strSql & " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),HC05) AS 申請人,NVL(CPM03,CP10)                          AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP19 AS 後金,decode(substr(cp09,1,1),'A',cp60,'B',cp60,CP09) AS 收回日,CP75 AS 收回金額, CP09 FROM CASEPROGRESS,HIRECASE,CUSTOMER,STAFF S1,STAFF S2,CASEPROPERTYMAP        WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND " & SQLNewFag("HC05", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4
strSql = strSql & " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),SP08) AS 申請人,NVL(DECODE(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,CP19 AS 後金,decode(substr(cp09,1,1),'A',cp60,'B',cp60,CP09) AS 收回日,CP75 AS 收回金額, CP09 FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,STAFF S1,STAFF S2,CASEPROPERTYMAP WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND " & SQLNewFag("SP08", "CU") & " AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5

strSql = strSql + "  ORDER BY 本所案號 "
CheckOC
s = 0
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/11/4
Else
    InsertQueryLog (0) 'Add By Sindy 2010/11/4
    ShowNoData
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    '92.04.18 nick
    'Me.Hide
    tmpBol = fnCancelNowFormAndShowParentForm(Me)
    Exit Sub
End If
Me.grdDataList.Visible = False
Set grdDataList.Recordset = adoRecordset
'add by nick 2004/11/29
Dim tmpIsC As Boolean
'Add By Cheng 2002/03/14
Me.grdDataList.ColAlignment(7) = flexAlignRightCenter
For i = 1 To grdDataList.Rows - 1
    'add by nick 2004/11/29
    tmpIsC = False
    Me.grdDataList.TextMatrix(i, 4) = Me.grdDataList.TextMatrix(i, 4) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(i, 10), "1")
    grdDataList.row = i
    grdDataList.col = 8
    'add by nick 2004/11/29 C 類才要用相關收文號
    If UCase(Mid(grdDataList.Text, 1, 1)) = "C" Then
        strTemp1 = GetPrjCaseNumber(grdDataList.Text)
        tmpIsC = True
    Else
        strTemp1 = grdDataList.Text
    End If
    grdDataList.Text = GetPrjGoBackDate(strTemp1)
    'edit by nick 2004/11/29
    If tmpIsC = True Then
        grdDataList.col = 9
        grdDataList.Text = GetPrjGoBackMoney(strTemp1)
    End If
    DoEvents
Next i
Me.grdDataList.Visible = True
Me.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm10011201_2 = Nothing
End Sub
