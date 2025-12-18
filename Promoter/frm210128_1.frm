VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210128_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶名稱已存在，是否要繼續？"
   ClientHeight    =   5736
   ClientLeft      =   156
   ClientTop       =   996
   ClientWidth     =   8460
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   8460
   Begin VB.CommandButton cmdok 
      Caption         =   "否(&N)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7470
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   15
      Width           =   870
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "是(&Y)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   6540
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   15
      Width           =   870
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4800
      Left            =   30
      TabIndex        =   2
      Top             =   930
      Width           =   8385
      _ExtentX        =   14796
      _ExtentY        =   8467
      _Version        =   393216
      Cols            =   6
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
      _Band(0).Cols   =   6
   End
   Begin VB.Label Label3 
      Caption         =   "　　　X/Y/R-.客戶/代理人/潛在客戶 聯絡人　編號欄為空白.國內開拓函特定公司不列印者"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   750
      Width           =   8295
   End
   Begin VB.Label Label2 
      Caption         =   "備註：X.申請人 Y.國外代理人 R.國內外潛在客戶 XXX-.法務開拓客戶 XXX.不得代理案件之客戶或代理人"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   540
      Width           =   8295
   End
   Begin MSForms.Label Label1 
      Height          =   495
      Left            =   390
      TabIndex        =   3
      Top             =   30
      Width           =   5895
      VariousPropertyBits=   27
      Caption         =   "Label1"
      Size            =   "10398;873"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm210128_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/14 Form2.0已修改 label1/grdDataList
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

Public G_strText As String, G_intIndex As Integer
Public m_PrevForm As Form 'Add by Amy 2021/08/16 前一畫面

Private Sub SetDataListWidth()
grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "編號"
grdDataList.ColWidth(0) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "名稱"
grdDataList.ColWidth(1) = 2000
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "國籍"
grdDataList.ColWidth(2) = 1000
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "智權人員"
grdDataList.ColWidth(3) = 1000
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "地址"
grdDataList.ColWidth(4) = 3000
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "電話"
grdDataList.ColWidth(5) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "傳真"
grdDataList.ColWidth(6) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "狀態"
grdDataList.ColWidth(7) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "備註"
grdDataList.ColWidth(8) = 2000
grdDataList.CellAlignment = flexAlignCenterCenter
End Sub

Private Sub cmdOK_Click(Index As Integer)
'Modify by Amy 2021/08/16
If TypeName(m_PrevForm) = "Nothing" Then
    

    Select Case Index
       Case 0
          frm210128.txtSameCnt = "Y"
       Case 1
          frm210128.txtSameCnt = "N"
       Case Else
    End Select
End If
'Modify By Sindy 2014/2/27
'Call frm210128.txtPOC_Validate(G_intIndex, False)
'Me.Hide
'frm210128.Show
Unload frm210128_1
'2014/2/27 END
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SetDataListWidth
   'Modify by Amy 2015/01/15 隱藏「是」案鈕,改「否」按鈕顯示為「回上一頁」(名稱重覆不可存檔-王副總)
   'Memo 2017/02/23 名稱與資料庫資料相同確定為不同主體者需於名稱後面加 1,2...
   'ex:誠品股份有限公司-政興無法存檔,因與吉甲地誠品股份有限公司相似
   cmdOK(0).Visible = False
   cmdOK(1).Caption = "回上一頁"
   'add by sonia 2015/6/8
   Me.Caption = "客戶名稱已存在，不可重覆建檔！"
End Sub

'Modify By Sindy 2014/2/27
Public Function StrMenu(strCustNo As String) As Boolean
Dim m_i As Integer
Dim rsTmp As New ADODB.Recordset
Dim strSR04 As String
Dim StrSQLa As String
'Add by Amy 2021/08/13
Dim strCheckWay As String, strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
Dim strRCLSql As String 'Add by Amy 2024/05/21

StrMenu = True
'Me.Enabled = False

'若為國內智權人員或國內工程師, 不可查代理人資料
'Modify By Sindy 2011/01/04 取消
'If bolFNation = False Then
'    StrSQLa = " And FA01<'Y' "
'End If
Screen.MousePointer = vbHourglass

'Add by Amy 2021/08/13 +檢查對造
strSQL1 = " AND CP01 IN (" & SQLGrpStr(GetGroupKindByTwo, 2) & ") "
strSQL2 = " AND CP01 IN (" & SQLGrpStr("", 1) & ") "
StrSQL3 = " AND CP01 IN (" & SQLGrpStr("", 3) & ") "
StrSQL4 = " AND CP01 IN (" & SQLGrpStr("", 4) & ") "
strSQL5 = " AND CP01 IN (" & SQLGrpStr("", 5) & ") "
strCheckWay = "=" 'Modify by Amy 2021/08/16 原:=1字首比對,若為[對造後加字]要可存檔,故與下方用>0的判斷不同
Call Pub_ProcR100102_1(strUserNum & "@" & Me.Name, strSQL1, strSQL2, StrSQL3, StrSQL4, strSQL5, ChgSQL(G_strText), strCheckWay, True)
'end 2021/08/13
Call ChkRiskData(2, Me.Name, , , G_strText, strRCLSql) 'Add by Amy 2024/05/20 +風險檢查對象(判斷同對造,用=)

strSql = "SELECT * FROM ("
'客戶檔
strSql = strSql & "SELECT CU01||CU02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','')||decode(cu121,'Y','●','') AS 編號,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU31,null,CU23,CU31) AS 地址,Decode(CU16,null,CU17,CU16) AS 電話,Decode(CU18,null,CU19,CU18) AS 傳真,CU80 AS 狀態,CU79 AS 備註 FROM CUSTOMER,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(CU04,'" & ChgSQL(G_strText) & "')>0 or instr(CU05,'" & ChgSQL(G_strText) & "')>0 or instr(CU88,'" & ChgSQL(G_strText) & "')>0 or instr(CU89,'" & ChgSQL(G_strText) & "')>0 or instr(CU90,'" & ChgSQL(G_strText) & "')>0 or instr(CU06,'" & ChgSQL(G_strText) & "')>0) A WHERE CU10=NA01(+) AND CU01=A.A1 AND CU13=ST01(+)"
'國外潛在客戶
strSql = strSql & " union all SELECT PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,NVL(PCU08,DECODE(PCU03,NULL,PCU07,PCU03||' '||PCU04||' '||PCU05||' '||PCU06)) AS 名稱,NA03 AS 國籍,' ' AS 智權人員,PCU27 AS 地址,Decode(PCU13,null,PCU14,PCU13) AS 電話,Decode(PCU15,null,PCU16,PCU15) AS 傳真,PCU39 AS 狀態,PCU40 AS 備註 from potcustomer,nation, (Select Distinct pcu01 As A1 From potcustomer Where instr(pcu08,'" & ChgSQL(G_strText) & "')>0 or instr(pcu03,'" & ChgSQL(G_strText) & "')>0 or instr(pcu04,'" & ChgSQL(G_strText) & "')>0 or instr(pcu05,'" & ChgSQL(G_strText) & "')>0 or instr(pcu06,'" & ChgSQL(G_strText) & "')>0 or instr(pcu07,'" & ChgSQL(G_strText) & "')>0) A where pcu09=na01(+) and pcu01=A.A1"
'國內潛在客戶
strSql = strSql & " union all SELECT POC01||POC02||Decode(POC02,'0','','＊') AS 編號,NVL(POC03,DECODE(POC23,NULL,POC27,POC23||' '||POC24||' '||POC25||' '||POC26)) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,POC10 AS 地址,Decode(POC05,null,POC06,POC05) AS 電話,POC07 AS 傳真,POC14 AS 狀態,POC15 AS 備註 from potcustomer1,nation,STAFF, (Select Distinct poc01 As A1 From potcustomer1 Where instr(poc03,'" & ChgSQL(G_strText) & "')>0 or instr(poc23,'" & ChgSQL(G_strText) & "')>0 or instr(poc24,'" & ChgSQL(G_strText) & "')>0 or instr(poc25,'" & ChgSQL(G_strText) & "')>0 or instr(poc26,'" & ChgSQL(G_strText) & "')>0 or instr(poc27,'" & ChgSQL(G_strText) & "')>0) A where poc04=na01(+) and poc01=A.A1 AND poc13=ST01(+)" & IIf(strCustNo <> "", " and poc01||poc02<>'" & strCustNo & "'", "")
'國外代理人
strSql = strSql & " union all select FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) as 名稱,NA03 AS 國籍,' ' AS 智權人員,FA17 AS 地址,Decode(FA12,null,FA13,FA12) AS 電話,Decode(FA14,null,FA15,FA14) AS 傳真,FA69 AS 狀態, FA29 AS 備註 from fagent,nation, (Select Distinct FA01 As A1 From Fagent Where instr(fa04,'" & ChgSQL(G_strText) & "')>0 or instr(fa05,'" & ChgSQL(G_strText) & "')>0 or instr(fa63,'" & ChgSQL(G_strText) & "')>0 or instr(fa64,'" & ChgSQL(G_strText) & "')>0 or instr(fa65,'" & ChgSQL(G_strText) & "')>0 or instr(fa06,'" & ChgSQL(G_strText) & "')>0) A where fa10=na01(+) AND FA01=A.A1 " & StrSQLa
'法務開拓客戶
strSql = strSql & " union all SELECT ecd02||'-'||LPAD(ecd01,6,'0') AS 編號,NVL(ecd03,'')||NVL(ecd04,'') AS 名稱,NA03 AS 國籍,' ' AS 智權人員,''||ECD05||''||ECD06||''||ECD07||''||ECD08||''||ECD09 AS 地址,' ' AS 電話,' ' AS 傳真,ecd15 AS 狀態,ecd16 AS 備註 From expandcusdetail, expandcusattr, nation,(Select Distinct nvl(ecd01,'')||nvl(ecd02,'') as A1 From expandcusdetail Where instr(ecd03,'" & ChgSQL(G_strText) & "')>0 or instr(ecd04,'" & ChgSQL(G_strText) & "')>0) A Where ecd10=na01(+) and ecd02=eca01(+) and nvl(ecd01,'')||nvl(ecd02,'')=A.A1"
'Add By Sindy 2012/4/5 不得代理案件之客戶或代理人
strSql = strSql & " union all SELECT NT01||Decode(NT21,null,'⊕','') AS 編號,NVL(NT02,DECODE(NT03,NULL,NT07,NT03||' '||NT04||' '||NT05||' '||NT06)) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,NVL(NT09,DECODE(NT10,NULL,NT16,NT10||' '||NT11||' '||NT12||' '||NT13||' '||NT14||' '||NT15)) AS 地址,' ' AS 電話,' ' AS 傳真,Decode(NT21,null,'不得代理','') AS 狀態,Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 As 備註 FROM NotAgent,NATION,STAFF,(Select Distinct NT01 As A1 From NotAgent Where instr(NT02,'" & ChgSQL(G_strText) & "')>0 or instr(NT03,'" & ChgSQL(G_strText) & "')>0 or instr(NT04,'" & ChgSQL(G_strText) & "')>0 or instr(NT05,'" & ChgSQL(G_strText) & "')>0 or instr(NT06,'" & ChgSQL(G_strText) & "')>0 or instr(NT07,'" & ChgSQL(G_strText) & "')>0) A WHERE NT08=NA01(+) AND NT01=A.A1 AND NT18=ST01(+)"
'Add by Amy 2021/08/30 國內開拓函特定公司不列印者
strSql = strSql & " union all Select '' AS 編號,TBNP01 AS 名稱,'' AS 國籍,' ' AS 智權人員,'' AS 地址,' ' AS 電話,' ' AS 傳真,'' AS 狀態,'國內開拓函特定公司不列印者' AS 備註 From TMBulletinnp Where Instr(TBNP01,'" & ChgSQL(G_strText) & "')>0 "
'Add by Amy 2024/05/20 +風險檢查對象
strSql = strSql & " Union All Select RCL01 AS 編號,Decode(RCLFIeld,'中',RCL02,'英',rtrim(RCL03||' '||RCL04||' '||RCL05||' '||RCL06),RCL07) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(RCL09,null,Decode(RCL10,null,RCL16,rtrim(RCL10||' '||RCL11||' '||RCL12||' '||RCL13||' '||RCL14||' '||RCL15)),RCL09) AS 地址,' ' AS 電話,' ' AS 傳真,'' AS 狀態,RCL23 AS 備註 From (" & strRCLSql & "),Nation Where SubStr(RCL08,1,3)=NA01(+) "
'Add by Amy 2021/08/30 聯絡人
'中文欄
strSql = strSql & " union all Select PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,'' AS 國籍,ST02 AS 智權人員,'' AS 地址,' ' AS 電話,' ' AS 傳真,'' AS 狀態,'客戶聯絡人' AS 備註 From (Select * From PotCustCont Where instr(pcc05,'" & ChgSQL(UCase(G_strText)) & "')>0) A,Customer,Staff Where CU13=ST01(+) And CU01(+)=PCC01 AND CU02='0' "
strSql = strSql & " union all Select PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,'' AS 國籍,ST02 AS 智權人員,'' AS 地址,' ' AS 電話,' ' AS 傳真,'' AS 狀態,'客戶聯絡人' AS 備註 From (Select * From PotCustCont Where instr(pcc05,'" & ChgSQL(UCase(G_strText)) & "')>0) A,PotCustomer,Staff Where PCU01(+)=PCC01 And PCU02='0' And substr(LTrim(PCU38),1,5)=ST01(+) "
strSql = strSql & " union all Select PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,'' AS 國籍,ST02 AS 智權人員,'' AS 地址,' ' AS 電話,' ' AS 傳真,'' AS 狀態,'客戶聯絡人' AS 備註 From (Select * From PotCustCont Where instr(pcc05,'" & ChgSQL(UCase(G_strText)) & "')>0) A,PotCustomer1,Staff Where POC01(+)=PCC01 And POC02='0' And POC13=ST01(+) "
strSql = strSql & " union all Select PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,'' AS 國籍,'' AS 智權人員,'' AS 地址,' ' AS 電話,' ' AS 傳真,'' AS 狀態,'客戶聯絡人' AS 備註 From (Select * From PotCustCont Where instr(pcc05,'" & ChgSQL(UCase(G_strText)) & "')>0) A,Fagent Where FA01(+)=PCC01 And FA02='0' "
'英文欄
strSql = strSql & " union all Select PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,'' AS 國籍,ST02 AS 智權人員,'' AS 地址,' ' AS 電話,' ' AS 傳真,'' AS 狀態,'客戶聯絡人' AS 備註 From (Select * From PotCustCont Where instr(Upper(pcc03),'" & ChgSQL(UCase(G_strText)) & "')>0) A,Customer,Staff Where CU13=ST01(+) And CU01(+)=PCC01 AND CU02='0' "
strSql = strSql & " union all Select PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,'' AS 國籍,ST02 AS 智權人員,'' AS 地址,' ' AS 電話,' ' AS 傳真,'' AS 狀態,'客戶聯絡人' AS 備註 From (Select * From PotCustCont Where instr(Upper(pcc03),'" & ChgSQL(UCase(G_strText)) & "')>0) A,PotCustomer,Staff Where PCU01(+)=PCC01 And PCU02='0' And substr(LTrim(PCU38),1,5)=ST01(+) "
strSql = strSql & " union all Select PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,'' AS 國籍,ST02 AS 智權人員,'' AS 地址,' ' AS 電話,' ' AS 傳真,'' AS 狀態,'客戶聯絡人' AS 備註 From (Select * From PotCustCont Where instr(Upper(pcc03),'" & ChgSQL(UCase(G_strText)) & "')>0) A,PotCustomer1,Staff Where POC01(+)=PCC01 And POC02='0' And POC13=ST01(+) "
strSql = strSql & " union all Select PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,'' AS 國籍,'' AS 智權人員,'' AS 地址,' ' AS 電話,' ' AS 傳真,'' AS 狀態,'客戶聯絡人' AS 備註 From (Select * From PotCustCont Where instr(Upper(pcc03),'" & ChgSQL(UCase(G_strText)) & "')>0) A,Fagent Where FA01(+)=PCC01 And FA02='0' "
'日文欄
strSql = strSql & " union all Select PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,'' AS 國籍,ST02 AS 智權人員,'' AS 地址,' ' AS 電話,' ' AS 傳真,'' AS 狀態,'客戶聯絡人' AS 備註 From (Select * From PotCustCont Where instr(pcc04,'" & ChgSQL(UCase(G_strText)) & "')>0) A,Customer,Staff Where CU13=ST01(+) And CU01(+)=PCC01 AND CU02='0' "
strSql = strSql & " union all Select PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,'' AS 國籍,ST02 AS 智權人員,'' AS 地址,' ' AS 電話,' ' AS 傳真,'' AS 狀態,'客戶聯絡人' AS 備註 From (Select * From PotCustCont Where instr(pcc04,'" & ChgSQL(UCase(G_strText)) & "')>0) A,PotCustomer,Staff Where PCU01(+)=PCC01 And PCU02='0' And substr(LTrim(PCU38),1,5)=ST01(+) "
strSql = strSql & " union all Select PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,'' AS 國籍,ST02 AS 智權人員,'' AS 地址,' ' AS 電話,' ' AS 傳真,'' AS 狀態,'客戶聯絡人' AS 備註 From (Select * From PotCustCont Where instr(pcc04,'" & ChgSQL(UCase(G_strText)) & "')>0) A,PotCustomer1,Staff Where POC01(+)=PCC01 And POC02='0' And POC13=ST01(+) "
strSql = strSql & " union all Select PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,'' AS 國籍,'' AS 智權人員,'' AS 地址,' ' AS 電話,' ' AS 傳真,'' AS 狀態,'客戶聯絡人' AS 備註 From (Select * From PotCustCont Where instr(pcc04,'" & ChgSQL(UCase(G_strText)) & "')>0) A,Fagent Where FA01(+)=PCC01 And FA02='0' "
'end 2021/08/30
'Add by Amy 2021/08/13 對造
strSql = strSql & " union all Select Distinct R021001 AS 編號,R021002 AS 名稱,'' AS 國籍,'' AS 智權人員,'' AS 地址,'' AS 電話,'' AS 傳真,Decode(R021004,'1','對造','其他相關人') AS 狀態,'' AS 備註 From R100102_1 Where ID='" & strUserNum & "@" & Me.Name & "' And R021004<3 "
strSql = strSql & ") X order by 編號"

CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <= 0 Then
   Screen.MousePointer = vbDefault
   StrMenu = False
   Exit Function
End If
Set grdDataList.Recordset = adoRecordset
frm210128.txtSameCnt = adoRecordset.RecordCount
Label1.Caption = "客戶名稱：" & frm210128.txtPOC(G_intIndex)
CheckOC
'Me.Enabled = True
Screen.MousePointer = vbDefault
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set frm210128_1 = Nothing
End Sub
