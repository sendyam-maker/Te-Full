VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100117_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文日查詢代理人作業進度"
   ClientHeight    =   5730
   ClientLeft      =   170
   ClientTop       =   960
   ClientWidth     =   9310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9310
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件進度(&C)"
      Default         =   -1  'True
      Height          =   400
      Index           =   2
      Left            =   5970
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   30
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   8460
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   30
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   7230
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   30
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   5208
      Left            =   36
      TabIndex        =   3
      Top             =   468
      Width           =   9264
      _ExtentX        =   16351
      _ExtentY        =   9172
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      ScrollTrack     =   -1  'True
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
      _Band(0).Cols   =   11
   End
End
Attribute VB_Name = "frm100117_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Sonia 2022/1/20 改成Form2.0(grdDataList改Fonts)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit

Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
Dim s As Integer, i As Integer, j As Integer
Dim strSql As String, strTemp As Variant, intK As Integer, StrTest As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
Dim StrTag As String
'Added by Lydia 2019/11/01 利益衝突案件
Dim m_AllSys As String '預設全部系統別
Dim intCufaCnt As Integer '限閱案件X件

Private Sub SetDataListWidth()
'Modified by Lydia 2019/11/01
'GrdDataList.Cols = 13 '12
Dim intField As Integer
intField = 19
grdDataList.Cols = intField
'end 2019/11/01

grdDataList.row = 0
'Add By Sindy 2012/2/8
grdDataList.col = 0: grdDataList.Text = "V"
grdDataList.ColWidth(0) = 200
'2012/2/8 End
grdDataList.col = 1: grdDataList.Text = "發文日"
grdDataList.ColWidth(1) = 850
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 2: grdDataList.Text = "本所案號"
grdDataList.ColWidth(2) = 1400
grdDataList.CellAlignment = flexAlignCenterCenter
Dim iDep As String
iDep = PUB_GetST06(strUserNum)
grdDataList.col = 3: grdDataList.Text = "分所號"
'電腦中心，跟分所才秀
If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
    grdDataList.ColWidth(3) = 0
Else
    grdDataList.ColWidth(3) = 620
End If
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 4: grdDataList.Text = "案件名稱"
grdDataList.ColWidth(4) = 2000
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 5: grdDataList.Text = "申請國家"
grdDataList.ColWidth(5) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 6: grdDataList.Text = "案件性質"
grdDataList.ColWidth(6) = 1600
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 7: grdDataList.Text = "代理人"
grdDataList.ColWidth(7) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 8: grdDataList.Text = "收達日"
grdDataList.ColWidth(8) = 850
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 9: grdDataList.Text = "提申日"
grdDataList.ColWidth(9) = 850
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 10: grdDataList.Text = "申請人"
grdDataList.ColWidth(10) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
'Add By Cheng 2003/08/18
grdDataList.col = 11: grdDataList.Text = "CP09"
grdDataList.ColWidth(11) = 0
grdDataList.CellAlignment = flexAlignCenterCenter
'add by nickc 2005/05/13
grdDataList.col = 12: grdDataList.Text = ""
grdDataList.ColWidth(12) = 0
grdDataList.CellAlignment = flexAlignCenterCenter

'Added by Lydia 2019/11/01 隱藏欄位：申請人1~5, FC代理人
For intI = 13 To intField - 1
     grdDataList.col = intI
     grdDataList.ColWidth(intI) = 0
Next intI
'end 2019/11/01
End Sub

'92.04.16 nick
Public Sub PubShowNextData()
   Select Case cmdState
      Case 0
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Case 1
         fnCloseAllFrm100
      'Add By Sindy 2012/2/8 案件進度
      Case 2
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
             grdDataList.col = 2
             If Not IsNull(grdDataList.Text) Then
                If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
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
      Case Else
   End Select
End Sub

Private Sub cmdok_Click(Index As Integer)
   '92.04.16 nick 紀錄作用按鍵
   cmdState = Index
   PubShowNextData
   Exit Sub
'   '92.04.16 nick 以下無效
'   Select Case index
'      Case 0
'         Me.Hide
'      Case 1
'         bolToEndByNick = True
'         Unload Me
'         Exit Sub
'      Case Else
'   End Select
End Sub

Private Sub Form_Activate()
If bolFNation = False Then
    s = MsgBox("國內人員不可查詢代理人案件", , "違規.....")
    Unload Me
    Exit Sub
End If
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   SetDataListWidth
   '92.04.16 nick
   cmdState = -1
End Sub

Sub StrMenu()
'Add By Cheng 2002/07/09
Dim StrSQLa As String
Dim ii As Integer
Dim dblRow As Double 'Add By Sindy 2025/9/3

Me.Enabled = False
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
strSQL5 = ""
StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人,"
If Len(Trim(frm100117_1.txt1(0))) <> 0 Then
   strSQL1 = strSQL1 & " and cp27>=" & Val(ChangeTStringToWString(frm100117_1.txt1(0))) & " "
End If
If Len(Trim(frm100117_1.txt1(1))) <> 0 Then
   strSQL1 = strSQL1 & " and cp27<=" & Val(ChangeTStringToWString(frm100117_1.txt1(1))) & " "
End If
If Len(Trim(frm100117_1.txt1(0))) <> 0 Or Len(Trim(frm100117_1.txt1(1))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100117_1.Label1 & frm100117_1.txt1(0) & "-" & frm100117_1.txt1(1) 'Add By Sindy 2010/11/15
End If
If Len(Trim(frm100117_1.txt1(4))) <> 0 Then
   strSQL1 = strSQL1 & " and cp44='" & GetNewFagent(frm100117_1.txt1(4)) & "' "
   pub_QL05 = pub_QL05 & ";" & frm100117_1.Label3(0) & frm100117_1.txt1(4) & frm100117_1.LBL1 'Add By Sindy 2010/11/15
End If

'Add By Sindy 2012/2/8 案件性質
If Len(Trim(frm100117_1.txt1(6))) <> 0 Then
   strSQL1 = strSQL1 & " and cp10>='" & Trim(frm100117_1.txt1(6)) & "' "
End If
If Len(Trim(frm100117_1.txt1(7))) <> 0 Then
   strSQL1 = strSQL1 & " and cp10<='" & Trim(frm100117_1.txt1(7)) & "' "
End If
If Len(Trim(frm100117_1.txt1(6))) <> 0 Or Len(Trim(frm100117_1.txt1(7))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100117_1.Label5(0) & frm100117_1.txt1(6) & "-" & frm100117_1.txt1(7)
End If
'2012/2/8 End

strSQL2 = strSQL1
StrSQL3 = strSQL1
StrSQL4 = strSQL1
strSQL5 = strSQL1
'Added by Lydia 2019/11/01 利益衝突案件
m_AllSys = IIf(frm100117_1.txt1(5).Text <> "ALL", frm100117_1.txt1(5).Text, GetAllSysKind(, frm100117_1.txt1(5).Text))
intCufaCnt = 0
'end 2019/11/01

If Len(Trim(frm100117_1.txt1(2))) <> 0 Then
   strSQL1 = strSQL1 & " AND PA09>='" & frm100117_1.txt1(2) & "' "
   strSQL2 = strSQL2 & " AND TM10>='" & frm100117_1.txt1(2) & "' "
   StrSQL3 = StrSQL3 & " AND LC15>='" & frm100117_1.txt1(2) & "' "
   'StrSQL4 = StrSQL4 & " AND PA09>='" & frm100117_1.txt1(2) & "' "
   strSQL5 = strSQL5 & " AND SP09>='" & frm100117_1.txt1(2) & "' "
End If
If Len(Trim(frm100117_1.txt1(3))) <> 0 Then
   strSQL1 = strSQL1 & " AND PA09<='" & frm100117_1.txt1(3) & "' "
   strSQL2 = strSQL2 & " AND TM10<='" & frm100117_1.txt1(3) & "' "
   StrSQL3 = StrSQL3 & " AND LC15<='" & frm100117_1.txt1(3) & "' "
   'StrSQL4 = StrSQL4 & " AND PA09<='" & frm100117_1.txt1(3) & "' "
   strSQL5 = strSQL5 & " AND SP09<='" & frm100117_1.txt1(3) & "' "
End If
If Len(Trim(frm100117_1.txt1(2))) <> 0 Or Len(Trim(frm100117_1.txt1(3))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100117_1.Label2 & frm100117_1.txt1(2) & "-" & frm100117_1.txt1(3) 'Add By Sindy 2010/11/15
End If
If Len(Trim(frm100117_1.txt1(5))) <> 0 Then
   'Modify By Cheng 2002/03/15
'   strSQL1 = strSQL1 & " and cp01 in(" & SQLGrpStr(frm100117_1.txt1(5), 1) & ") "
'   strSQL2 = strSQL2 & " and cp01 in(" & SQLGrpStr(frm100117_1.txt1(5), 2) & ") "
'   StrSQL3 = StrSQL3 & " and cp01 in(" & SQLGrpStr(frm100117_1.txt1(5), 3) & ") "
'   StrSQL5 = StrSQL5 & " and cp01 in(" & SQLGrpStr(frm100117_1.txt1(5), 5) & ") "
   strSQL1 = strSQL1 & " and cp01 in(" & SQLGrpStr(IIf(frm100117_1.txt1(5).Text <> "ALL", frm100117_1.txt1(5).Text, GetAllSysKind(frm100117_1.txt1(5))), 1) & ") "
   strSQL2 = strSQL2 & " and cp01 in(" & SQLGrpStr(IIf(frm100117_1.txt1(5).Text <> "ALL", frm100117_1.txt1(5).Text, GetAllSysKind(frm100117_1.txt1(5))), 2) & ") "
   StrSQL3 = StrSQL3 & " and cp01 in(" & SQLGrpStr(IIf(frm100117_1.txt1(5).Text <> "ALL", frm100117_1.txt1(5).Text, GetAllSysKind(frm100117_1.txt1(5))), 3) & ") "
   strSQL5 = strSQL5 & " and cp01 in(" & SQLGrpStr(IIf(frm100117_1.txt1(5).Text <> "ALL", frm100117_1.txt1(5).Text, GetAllSysKind(frm100117_1.txt1(5))), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Left(frm100117_1.Label3(1), 5) & frm100117_1.txt1(5) 'Add By Sindy 2010/11/15
End If

'strSQL = "SELECT SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,nvl(na03,TM10) AS 申請國家,nvl(decode(tm10,'000',cpm03,cpm04),CP10) AS 案件性質,nvl(nvl(fa04,nvl(fa05||' '||fa63||' '||fa64||' '||fa65,fa06)),cp44) AS 代理人," & SQLDate("CP46") & " AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日 FROM CASEPROGRESS,TRADEMARK,fagent,nation,casepropertymap WHERE cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) and CP44 IS NOT NULL  AND (TM29<>'Y' or tm29 is null) and tm10=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL2
'strSQL = strSQL + " union all select SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,nvl(na03,pa09) AS 申請國家,nvl(decode(pa09,'000',cpm03,cpm04),CP10) AS 案件性質,nvl(nvl(fa04,nvl(fa05||' '||fa63||' '||fa64||' '||fa65,fa06)),cp44) AS 代理人," & SQLDate("CP46") & " AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日 FROM CASEPROGRESS,PATENT,fagent,nation,casepropertymap WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP44 IS NOT NULL AND (PA57<>'Y' or pa57 is null) and pa09=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1
'strSQL = strSQL + " union all select SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,nvl(na03,SP09) AS 申請國家,nvl(decode(sp09,'000',cpm03,cpm04),CP10) AS 案件性質,nvl(nvl(fa04,nvl(fa05||' '||fa63||' '||fa64||' '||fa65,fa06)),cp44) AS 代理人," & SQLDate("CP46") & " AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日 FROM CASEPROGRESS,SERVICEPRACTICE,fagent,nation,casepropertymap WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND CP44 IS NOT NULL AND  (SP15<>'Y' or sp15 is null) and sp09=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL5
'strSQL = strSQL + " union all select SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,nvl(na03,LC15) AS 申請國家,nvl(decode(lc15,'000',cpm03,cpm04),CP10) AS 案件性質,nvl(nvl(fa04,nvl(fa05||' '||fa63||' '||fa64||' '||fa65,fa06)),cp44) AS 代理人," & SQLDate("CP46") & " AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日 FROM CASEPROGRESS,LAWCASE,fagent,nation,casepropertymap WHERE cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND CP44 IS NOT NULL AND (LC08<>'Y' or lc08 is null) and lc15=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL3
'strSQL = strSQL + " ORDER BY 發文日,本所案號 "
'Modify By Cheng 2002/01/04
'多顯示申請人
'Modify By Cheng 2002/07/09
'strSQL = "SELECT SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,nvl(na03,TM10) AS 申請國家,nvl(decode(tm10,'000',cpm03,cpm04),CP10) AS 案件性質,nvl(nvl(fa04,nvl(fa05||' '||fa63||' '||fa64||' '||fa65,fa06)),cp44) AS 代理人," & SQLDate("CP46") & " AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人 FROM CASEPROGRESS,TRADEMARK,fagent,nation,casepropertymap,Customer WHERE cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) and CP44 IS NOT NULL  AND (TM29<>'Y' or tm29 is null) and tm10=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) And SUBSTR(TM23,1,8)=CU01 AND SUBSTR(TM23,9,1)=CU02(+) " & strSQL2
'strSQL = strSQL + " union all select SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,nvl(na03,pa09) AS 申請國家,nvl(decode(pa09,'000',cpm03,cpm04),CP10) AS 案件性質,nvl(nvl(fa04,nvl(fa05||' '||fa63||' '||fa64||' '||fa65,fa06)),cp44) AS 代理人," & SQLDate("CP46") & " AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人 FROM CASEPROGRESS,PATENT,fagent,nation,casepropertymap,Customer WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP44 IS NOT NULL AND (PA57<>'Y' or pa57 is null) and pa09=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) And SUBSTR(PA26,1,8)=CU01 AND SUBSTR(PA26,9,1)=CU02(+) " & strSQL1
'strSQL = strSQL + " union all select SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,nvl(na03,SP09) AS 申請國家,nvl(decode(sp09,'000',cpm03,cpm04),CP10) AS 案件性質,nvl(nvl(fa04,nvl(fa05||' '||fa63||' '||fa64||' '||fa65,fa06)),cp44) AS 代理人," & SQLDate("CP46") & " AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人 FROM CASEPROGRESS,SERVICEPRACTICE,fagent,nation,casepropertymap,Customer WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND CP44 IS NOT NULL AND  (SP15<>'Y' or sp15 is null) and sp09=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) And SUBSTR(SP08,1,8)=CU01 AND SUBSTR(SP08,9,1)=CU02(+) " & StrSQL5
'strSQL = strSQL + " union all select SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,nvl(na03,LC15) AS 申請國家,nvl(decode(lc15,'000',cpm03,cpm04),CP10) AS 案件性質,nvl(nvl(fa04,nvl(fa05||' '||fa63||' '||fa64||' '||fa65,fa06)),cp44) AS 代理人," & SQLDate("CP46") & " AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人 FROM CASEPROGRESS,LAWCASE,fagent,nation,casepropertymap,Customer WHERE cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND CP44 IS NOT NULL AND (LC08<>'Y' or lc08 is null) and lc15=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) And SUBSTR(LC11,1,8)=CU01 AND SUBSTR(LC11,9,1)=CU02(+) " & StrSQL3
'Modify By Cheng 2003/08/18
'strSQL = "SELECT SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,nvl(na03,TM10) AS 申請國家,nvl(decode(tm10,'000',cpm03,cpm04),CP10) AS 案件性質," & StrSQLa & " SUBSTR(' '||sqldatet(CP46),-9) AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人 FROM CASEPROGRESS,TRADEMARK,fagent,nation,casepropertymap,Customer,SYSTEMKIND WHERE cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) and CP44 IS NOT NULL  AND (TM29<>'Y' or tm29 is null) and tm10=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) And SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+) AND CP01=SK01(+) " & strSQL2
'strSQL = strSQL + " union all select SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,nvl(na03,pa09) AS 申請國家,nvl(decode(pa09,'000',cpm03,cpm04),CP10) AS 案件性質," & StrSQLa & " SUBSTR(' '||sqldatet(CP46),-9) AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人 FROM CASEPROGRESS,PATENT,fagent,nation,casepropertymap,Customer,SYSTEMKIND WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP44 IS NOT NULL AND (PA57<>'Y' or pa57 is null) and pa09=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) And SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND CP01=SK01(+) " & strSQL1
'strSQL = strSQL + " union all select SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,nvl(na03,SP09) AS 申請國家,nvl(decode(sp09,'000',cpm03,cpm04),CP10) AS 案件性質," & StrSQLa & " SUBSTR(' '||sqldatet(CP46),-9) AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人 FROM CASEPROGRESS,SERVICEPRACTICE,fagent,nation,casepropertymap,Customer,SYSTEMKIND WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND CP44 IS NOT NULL AND  (SP15<>'Y' or sp15 is null) and sp09=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) And SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND CP01=SK01(+) " & StrSQL5
'strSQL = strSQL + " union all select SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,nvl(na03,LC15) AS 申請國家,nvl(decode(lc15,'000',cpm03,cpm04),CP10) AS 案件性質," & StrSQLa & " SUBSTR(' '||sqldatet(CP46),-9) AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人 FROM CASEPROGRESS,LAWCASE,fagent,nation,casepropertymap,Customer,SYSTEMKIND WHERE cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND CP44 IS NOT NULL AND (LC08<>'Y' or lc08 is null) and lc15=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) And SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) AND CP01=SK01(+) " & StrSQL3
'strSQL = strSQL + " ORDER BY 發文日,本所案號 "
'edit by nickc 2005/05/13
'strSQL = "SELECT SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,nvl(na03,TM10) AS 申請國家,nvl(decode(tm10,'000',cpm03,cpm04),CP10) AS 案件性質," & StrSQLa & " SUBSTR(' '||sqldatet(CP46),-9) AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人, CP09 FROM CASEPROGRESS,TRADEMARK,fagent,nation,casepropertymap,Customer,SYSTEMKIND WHERE cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) and CP44 IS NOT NULL  AND (TM29<>'Y' or tm29 is null) and tm10=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) And SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+) AND CP01=SK01(+) " & strSQL2
'strSQL = strSQL + " union all select SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,nvl(na03,pa09) AS 申請國家,nvl(decode(pa09,'000',cpm03,cpm04),CP10) AS 案件性質," & StrSQLa & " SUBSTR(' '||sqldatet(CP46),-9) AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人, CP09 FROM CASEPROGRESS,PATENT,fagent,nation,casepropertymap,Customer,SYSTEMKIND WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP44 IS NOT NULL AND (PA57<>'Y' or pa57 is null) and pa09=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) And SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND CP01=SK01(+) " & strSQL1
'strSQL = strSQL + " union all select SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,nvl(na03,SP09) AS 申請國家,nvl(decode(sp09,'000',cpm03,cpm04),CP10) AS 案件性質," & StrSQLa & " SUBSTR(' '||sqldatet(CP46),-9) AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人, CP09 FROM CASEPROGRESS,SERVICEPRACTICE,fagent,nation,casepropertymap,Customer,SYSTEMKIND WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND CP44 IS NOT NULL AND  (SP15<>'Y' or sp15 is null) and sp09=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) And SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND CP01=SK01(+) " & strSQL5
'strSQL = strSQL + " union all select SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,nvl(na03,LC15) AS 申請國家,nvl(decode(lc15,'000',cpm03,cpm04),CP10) AS 案件性質," & StrSQLa & " SUBSTR(' '||sqldatet(CP46),-9) AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日,NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)) AS 申請人, CP09 FROM CASEPROGRESS,LAWCASE,fagent,nation,casepropertymap,Customer,SYSTEMKIND WHERE cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND CP44 IS NOT NULL AND (LC08<>'Y' or lc08 is null) and lc15=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) And SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) AND CP01=SK01(+) " & StrSQL3
'strSQL = strSQL + " ORDER BY 發文日,本所案號 "
'2010/9/15 MODIFY BY SONIA 日期欄改百年日期排序問題
'2010/11/8 MODIFY BY SONIA 加CP09<'C'條件
'Modified by Lydia 2019/11/01 增加欄位:申請人1~5(cust01~cust05),FC代理人
'strSql = "SELECT ' ' AS V,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,decode(tm28,'1','','N')||cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,nvl(na03,TM10) AS 申請國家,nvl(decode(tm10,'000',cpm03,cpm04),CP10) AS 案件性質," & StrSQLa & " SUBSTR(' '||sqldatet(CP46),-9) AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 申請人, CP09,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as FSort FROM CASEPROGRESS,TRADEMARK,fagent,nation,casepropertymap,Customer,SYSTEMKIND WHERE CP09<'C' AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) and CP44 IS NOT NULL  AND (TM29<>'Y' or tm29 is null) and tm10=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) And SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+) AND CP01=SK01(+) " & strSQL2
'strSql = strSql + " union all select ' ' AS V,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,decode(pa23,'1','','N')||cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,nvl(na03,pa09) AS 申請國家,nvl(decode(pa09,'000',cpm03,cpm04),CP10) AS 案件性質," & StrSQLa & " SUBSTR(' '||sqldatet(CP46),-9) AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 申請人, CP09,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as FSort FROM CASEPROGRESS,PATENT,fagent,nation,casepropertymap,Customer,SYSTEMKIND WHERE CP09<'C' AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP44 IS NOT NULL AND (PA57<>'Y' or pa57 is null) and pa09=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) And SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND CP01=SK01(+) " & strSQL1
'strSql = strSql + " union all select ' ' AS V,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,nvl(na03,SP09) AS 申請國家,nvl(decode(sp09,'000',cpm03,cpm04),CP10) AS 案件性質," & StrSQLa & " SUBSTR(' '||sqldatet(CP46),-9) AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 申請人, CP09,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as FSort FROM CASEPROGRESS,SERVICEPRACTICE,fagent,nation,casepropertymap,Customer,SYSTEMKIND WHERE CP09<'C' AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND CP44 IS NOT NULL AND  (SP15<>'Y' or sp15 is null) and sp09=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) And SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND CP01=SK01(+) " & strSQL5
'strSql = strSql + " union all select ' ' AS V,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,nvl(na03,LC15) AS 申請國家,nvl(decode(lc15,'000',cpm03,cpm04),CP10) AS 案件性質," & StrSQLa & " SUBSTR(' '||sqldatet(CP46),-9) AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 申請人, CP09,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as FSort FROM CASEPROGRESS,LAWCASE,fagent,nation,casepropertymap,Customer,SYSTEMKIND WHERE CP09<'C' AND cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND CP44 IS NOT NULL AND (LC08<>'Y' or lc08 is null) and lc15=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) And SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) AND CP01=SK01(+) " & StrSQL3
strSql = "SELECT ' ' AS V,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,decode(tm28,'1','','N')||cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,nvl(na03,TM10) AS 申請國家,nvl(decode(tm10,'000',cpm03,cpm04),CP10) AS 案件性質," & StrSQLa & " SUBSTR(' '||sqldatet(CP46),-9) AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 申請人" & _
            ", CP09,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as FSort,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno" & _
            " FROM CASEPROGRESS,TRADEMARK,fagent,nation,casepropertymap,Customer,SYSTEMKIND" & _
            " WHERE CP09<'C' AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) and CP44 IS NOT NULL  AND (TM29<>'Y' or tm29 is null) and tm10=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) And SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+) AND CP01=SK01(+) " & strSQL2
strSql = strSql & " union all select ' ' AS V,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,decode(pa23,'1','','N')||cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,nvl(na03,pa09) AS 申請國家,nvl(decode(pa09,'000',cpm03,cpm04),CP10) AS 案件性質," & StrSQLa & " SUBSTR(' '||sqldatet(CP46),-9) AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 申請人" & _
            ", CP09,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as FSort,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno" & _
            " FROM CASEPROGRESS,PATENT,fagent,nation,casepropertymap,Customer,SYSTEMKIND" & _
            " WHERE CP09<'C' AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP44 IS NOT NULL AND (PA57<>'Y' or pa57 is null) and pa09=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) And SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND CP01=SK01(+) " & strSQL1
strSql = strSql & " union all select ' ' AS V,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,nvl(na03,SP09) AS 申請國家,nvl(decode(sp09,'000',cpm03,cpm04),CP10) AS 案件性質," & StrSQLa & " SUBSTR(' '||sqldatet(CP46),-9) AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 申請人" & _
            ", CP09,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as FSort,sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno" & _
            " FROM CASEPROGRESS,SERVICEPRACTICE,fagent,nation,casepropertymap,Customer,SYSTEMKIND" & _
            " WHERE CP09<'C' AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND CP44 IS NOT NULL AND  (SP15<>'Y' or sp15 is null) and sp09=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) And SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND CP01=SK01(+) " & strSQL5
strSql = strSql & " union all select ' ' AS V,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,nvl(na03,LC15) AS 申請國家,nvl(decode(lc15,'000',cpm03,cpm04),CP10) AS 案件性質," & StrSQLa & " SUBSTR(' '||sqldatet(CP46),-9) AS 收達日,SUBSTR(' '||sqldatet(CP47),-9) AS 提申日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 申請人" & _
            ", CP09,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as FSort,lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno" & _
            " FROM CASEPROGRESS,LAWCASE,fagent,nation,casepropertymap,Customer,SYSTEMKIND" & _
            " WHERE CP09<'C' AND cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND CP44 IS NOT NULL AND (LC08<>'Y' or lc08 is null) and lc15=na01(+) and " & SQLNewFag("cp44", "fa") & " and cp01=cpm01(+) and cp10=cpm02(+) And SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) AND CP01=SK01(+) " & StrSQL3
'end 2019/11/01

strSql = strSql & " ORDER BY 發文日,FSort,本所案號 "
        
CheckOC
adoRecordset.CursorLocation = adUseClient
'Modified by Lydia 2019/11/01 改變型態
'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic

If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3

    'Added by Lydia 2019/11/01 逐案號判斷
    If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
        adoRecordset.MoveFirst
        Do While adoRecordset.EOF = False
            '利益衝突案件：逐案號判斷
            If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & adoRecordset.Fields("本所案號"), "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
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
        InsertQueryLog (dblRow) 'Add By Sindy 2010/11/15
        If adoRecordset.RecordCount = 0 Then
              GoTo JumpToNoData
        End If
    Else
        InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/11/15
    End If
    'end 2019/11/01
    
Else
    InsertQueryLog (0) 'Add By Sindy 2010/11/15
JumpToNoData:   'Added by Lydia 2019/11/01
    ShowNoData
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    'Modify By Cheng 2003/07/30
'    Me.Hide
    tmpBol = fnCancelNowFormAndShowParentForm(Me)
    Exit Sub
End If
Me.grdDataList.Visible = False
Set grdDataList.Recordset = adoRecordset
SetDataListWidth
For ii = 1 To Me.grdDataList.Rows - 1
    Me.grdDataList.TextMatrix(ii, 6) = Me.grdDataList.TextMatrix(ii, 6) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(ii, 11), "1")
Next ii
Me.grdDataList.Visible = True
CheckOC
Me.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100117_2 = Nothing
End Sub

Private Sub grdDataList_SelChange()
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
