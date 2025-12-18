VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm10011202_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "後金收回查詢"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   975
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   8496
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   0
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7272
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   5016
      Left            =   72
      TabIndex        =   2
      Top             =   684
      Width           =   9204
      _ExtentX        =   16245
      _ExtentY        =   8837
      _Version        =   393216
      Cols            =   20
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
      _Band(0).Cols   =   20
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   1
      Left            =   4956
      TabIndex        =   6
      Top             =   444
      Width           =   2892
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   0
      Left            =   1116
      TabIndex        =   5
      Top             =   444
      Width           =   2412
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請人編號 :"
      Height          =   180
      Left            =   3750
      TabIndex        =   4
      Top             =   450
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號 :"
      Height          =   180
      Left            =   30
      TabIndex        =   3
      Top             =   450
      Width           =   810
   End
End
Attribute VB_Name = "frm10011202_2"
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
Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
Dim s As Integer, i As Integer, j As Integer, intK As Integer
Dim strSql As String, strTemp As Variant, StrTest As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer

Private Sub SetDataListWidth()
Dim iDep As String   '2010/9/15 ADD BY SONIA

iDep = PUB_GetST06(strUserNum)  '2010/9/15 ADD BY SONIA

grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "本所案號"
grdDataList.ColWidth(0) = 1550
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "案件名稱"
grdDataList.ColWidth(1) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 2: grdDataList.Text = "分所號"
'2010/9/15 ADD BY SONIA 電腦中心，跟分所才秀
If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
   grdDataList.ColWidth(2) = 0
Else
   grdDataList.ColWidth(2) = 620
End If
'2010/9/15 END
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 3: grdDataList.Text = "申請人"
grdDataList.ColWidth(3) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 4: grdDataList.Text = "相關案件性質"
grdDataList.ColWidth(4) = 1500
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 5: grdDataList.Text = "承辦人"
grdDataList.ColWidth(5) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 6: grdDataList.Text = "智權人員"
grdDataList.ColWidth(6) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 7: grdDataList.Text = "費用"
grdDataList.ColWidth(7) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 8: grdDataList.Text = "收回日"
grdDataList.ColWidth(8) = 850
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 9: grdDataList.Text = "收回金額"
grdDataList.ColWidth(9) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 10: grdDataList.Text = ""          '
grdDataList.ColWidth(10) = 0
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 11: grdDataList.Text = ""          '
grdDataList.ColWidth(11) = 0
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 12: grdDataList.Text = ""          '
grdDataList.ColWidth(12) = 0
grdDataList.CellAlignment = flexAlignCenterCenter
'Add By Cheng 2003/08/15
grdDataList.col = 13: grdDataList.Text = "CP09"          '
grdDataList.ColWidth(13) = 0
grdDataList.CellAlignment = flexAlignCenterCenter
'add by nickc 2005/05/10
grdDataList.col = 14: grdDataList.Text = ""
grdDataList.ColWidth(14) = 0
grdDataList.CellAlignment = flexAlignCenterCenter
'2010/9/15 ADD BY SONIA
grdDataList.col = 15: grdDataList.Text = ""
grdDataList.ColWidth(15) = 0
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


Private Sub cmdOK_Click(Index As Integer)
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
Dim ii As Integer

Me.Enabled = False
If frm10011202_1.Option1(0).Value = True Then
     lbl1(0).Caption = frm10011202_1.txt1(0) + "-" + frm10011202_1.txt1(1) + "-" + frm10011202_1.txt1(2) + "-" + frm10011202_1.txt1(3)
Else
    If frm10011202_1.Option1(1).Value = True Then
        lbl1(1).Caption = frm10011202_1.txt1(4) + "-" + frm10011202_1.txt1(5)
    End If
End If
CheckOC
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
strSQL5 = ""
'本所案號
If frm10011202_1.Option1(0).Value = True Then
   If Len(Trim(frm10011202_1.txt1(0))) <> 0 Then
      strSQL1 = strSQL1 & " AND C1.CP01='" & frm10011202_1.txt1(0) & "' "
   End If
   If Len(Trim(frm10011202_1.txt1(1))) <> 0 Then
      strSQL1 = strSQL1 & " AND C1.CP02='" & frm10011202_1.txt1(1) & "' "
   End If
   If Len(Trim(frm10011202_1.txt1(2))) <> 0 Then
      strSQL1 = strSQL1 & " AND C1.CP03='" & frm10011202_1.txt1(2) & "' "
   Else
      frm10011202_1.txt1(2) = "0"
      strSQL1 = strSQL1 & " AND C1.CP03='0' "
   End If
   If Len(Trim(frm10011202_1.txt1(3))) <> 0 Then
      strSQL1 = strSQL1 & " AND C1.CP04='" & frm10011202_1.txt1(3) & "' "
   Else
      frm10011202_1.txt1(3) = "00"
      strSQL1 = strSQL1 & " AND C1.CP04='00' "
   End If
   strSQL2 = strSQL1
   StrSQL3 = strSQL1
   StrSQL4 = strSQL1
   strSQL5 = strSQL1
   pub_QL05 = pub_QL05 & ";" & frm10011202_1.Option1(0).Caption & frm10011202_1.txt1(0) & "-" & frm10011202_1.txt1(1) & "-" & frm10011202_1.txt1(2) & "-" & frm10011202_1.txt1(3) 'Add By Sindy 2010/11/4
'申請人編號
Else
   strSQL1 = strSQL1 & " AND ((PA26>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND PA26<='" & GetNewFagent(frm10011202_1.txt1(5)) & "') OR (PA27>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND PA27<='" & GetNewFagent(frm10011202_1.txt1(5)) & "') OR (PA28>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND PA28<='" & GetNewFagent(frm10011202_1.txt1(5)) & "') OR (PA29>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND PA29<='" & GetNewFagent(frm10011202_1.txt1(5)) & "') OR (PA30>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND PA30<='" & GetNewFagent(frm10011202_1.txt1(5)) & "')) "
   strSQL2 = strSQL2 & " AND ((TM23>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND TM23<='" & GetNewFagent(frm10011202_1.txt1(5)) & "') or (TM78>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND TM78<='" & GetNewFagent(frm10011202_1.txt1(5)) & "') or (TM79>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND TM79<='" & GetNewFagent(frm10011202_1.txt1(5)) & "') or (TM80>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND TM80<='" & GetNewFagent(frm10011202_1.txt1(5)) & "') or (TM81>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND TM81<='" & GetNewFagent(frm10011202_1.txt1(5)) & "')) "
   'Modify By Sindy 2011/2/18 增加LC43,LC44,LC45,LC46
   StrSQL3 = StrSQL3 & " AND ((LC11>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND LC11<='" & GetNewFagent(frm10011202_1.txt1(5)) & "') or (LC43>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND LC43<='" & GetNewFagent(frm10011202_1.txt1(5)) & "') or (LC44>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND LC44<='" & GetNewFagent(frm10011202_1.txt1(5)) & "') or (LC45>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND LC45<='" & GetNewFagent(frm10011202_1.txt1(5)) & "') or (LC46>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND LC46<='" & GetNewFagent(frm10011202_1.txt1(5)) & "')) "
   'Modify By Sindy 2011/2/18 增加HC24,HC25,HC26,HC27
   StrSQL4 = StrSQL4 & " AND ((HC05>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND HC05<='" & GetNewFagent(frm10011202_1.txt1(5)) & "') or (HC24>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND HC24<='" & GetNewFagent(frm10011202_1.txt1(5)) & "') or (HC25>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND HC25<='" & GetNewFagent(frm10011202_1.txt1(5)) & "') or (HC26>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND HC26<='" & GetNewFagent(frm10011202_1.txt1(5)) & "') or (HC27>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND HC27<='" & GetNewFagent(frm10011202_1.txt1(5)) & "')) "
   strSQL5 = strSQL5 & " AND ((SP08>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND SP08<='" & GetNewFagent(frm10011202_1.txt1(5)) & "') OR (SP58>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND SP58<='" & GetNewFagent(frm10011202_1.txt1(5)) & "') OR (SP59>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND SP59<='" & GetNewFagent(frm10011202_1.txt1(5)) & "') or (SP65>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND SP65<='" & GetNewFagent(frm10011202_1.txt1(5)) & "') or (SP66>='" & GetNewFagent(frm10011202_1.txt1(4)) & "' AND SP66<='" & GetNewFagent(frm10011202_1.txt1(5)) & "')) "
   pub_QL05 = pub_QL05 & ";" & frm10011202_1.Option1(1).Caption & frm10011202_1.txt1(4) & "-" & frm10011202_1.txt1(5) 'Add By Sindy 2010/11/4
End If
'收回日期
strSQL1 = strSQL1 & " AND A0L02>=" & Val(frm10011202_1.txt1(6)) & " AND A0L02<=" & Val(frm10011202_1.txt1(7)) & " AND C1.CP10='909' "
strSQL2 = strSQL2 & " AND A0L02>=" & Val(frm10011202_1.txt1(6)) & " AND A0L02<=" & Val(frm10011202_1.txt1(7)) & " AND C1.CP10='909' "
StrSQL3 = StrSQL3 & " AND A0L02>=" & Val(frm10011202_1.txt1(6)) & " AND A0L02<=" & Val(frm10011202_1.txt1(7)) & " AND C1.CP10='909' "
StrSQL4 = StrSQL4 & " AND A0L02>=" & Val(frm10011202_1.txt1(6)) & " AND A0L02<=" & Val(frm10011202_1.txt1(7)) & " AND C1.CP10='909' "
strSQL5 = strSQL5 & " AND A0L02>=" & Val(frm10011202_1.txt1(6)) & " AND A0L02<=" & Val(frm10011202_1.txt1(7)) & " AND C1.CP10='909' "
pub_QL05 = pub_QL05 & ";" & frm10011202_1.Label2(1) & frm10011202_1.txt1(6) & "-" & frm10011202_1.txt1(7) 'Add By Sindy 2010/11/4
If Len(Trim(frm10011202_1.txt1(8))) <> 0 Then
   'Modify By Cheng 2002/03/14
'   strSQL1 = strSQL1 & " AND C1.CP01 IN (" & SQLGrpStr(frm10011202_1.txt1(8), 1) & ") "
'   strSQL2 = strSQL2 & " AND C1.CP01 IN (" & SQLGrpStr(frm10011202_1.txt1(8), 2) & ") "
'   StrSQL3 = StrSQL3 & " AND C1.CP01 IN (" & SQLGrpStr(frm10011202_1.txt1(8), 3) & ") "
'   StrSQL4 = StrSQL4 & " AND C1.CP01 IN (" & SQLGrpStr(frm10011202_1.txt1(8), 4) & ") "
'   StrSQL5 = StrSQL5 & " AND C1.CP01 IN (" & SQLGrpStr(frm10011202_1.txt1(8), 5) & ") "
   strSQL1 = strSQL1 & " AND C1.CP01 IN (" & SQLGrpStr(IIf(frm10011202_1.txt1(8).Text <> "ALL", frm10011202_1.txt1(8).Text, GetAllSysKind(frm10011202_1.txt1(8))), 1) & ") "
   strSQL2 = strSQL2 & " AND C1.CP01 IN (" & SQLGrpStr(IIf(frm10011202_1.txt1(8).Text <> "ALL", frm10011202_1.txt1(8).Text, GetAllSysKind(frm10011202_1.txt1(8))), 2) & ") "
   StrSQL3 = StrSQL3 & " AND C1.CP01 IN (" & SQLGrpStr(IIf(frm10011202_1.txt1(8).Text <> "ALL", frm10011202_1.txt1(8).Text, GetAllSysKind(frm10011202_1.txt1(8))), 3) & ") "
   StrSQL4 = StrSQL4 & " AND C1.CP01 IN (" & SQLGrpStr(IIf(frm10011202_1.txt1(8).Text <> "ALL", frm10011202_1.txt1(8).Text, GetAllSysKind(frm10011202_1.txt1(8))), 4) & ") "
   strSQL5 = strSQL5 & " AND C1.CP01 IN (" & SQLGrpStr(IIf(frm10011202_1.txt1(8).Text <> "ALL", frm10011202_1.txt1(8).Text, GetAllSysKind(frm10011202_1.txt1(8))), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Left(frm10011202_1.Label2(2), 5) & frm10011202_1.txt1(8) 'Add By Sindy 2010/11/4
End If

'Modify By Cheng 2002/04/25
'若已閉卷, 則在本所案號後加"*"號
'edit by nickc 2006/12/11
'                    strSQL = "SELECT decode(tm28,'1','','N')||C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,NVL(DECODE(TM10,'000',CPM03,CPM04),C2.CP10) AS 相關案件性質,NVL(S1.ST02,C1.CP14) AS 承辦人,NVL(S2.ST02,C1.CP13) AS 智權人員,C1.CP16 AS 費用," & SqlDateT("GG1.A0L02") & " AS 收回日,DECODE(C1.cp75,NULL,0,C1.cp75) 收回金額,'','','','', C1.CP09,C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04 as FSort         " & _
                             " FROM (select a0l01,max(a0l02) as a0l02 from ACC0L0 group by a0l01) GG1,ACC0M0,CASEPROGRESS C1,CASEPROGRESS C2,TRADEMARK,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER       WHERE GG1.A0L01=A0M01(+) AND A0M02=C1.CP60(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C1.CP14=S1.ST01(+) AND C1.CP13=S2.ST01(+) AND C2.CP01=CPM01(+) AND C2.CP10=CPM02(+) AND C1.CP43=C2.CP09(+) AND " & SQLNewFag("TM12", "CU") & " " & strSQL2
                    strSql = "SELECT decode(tm28,'1','','N')||C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,NVL(DECODE(TM10,'000',CPM03,CPM04),C2.CP10) AS 相關案件性質,NVL(S1.ST02,C1.CP14) AS 承辦人,NVL(S2.ST02,C1.CP13) AS 智權人員,C1.CP16 AS 費用," & SqlDateT("GG1.A0L02") & " AS 收回日,DECODE(C1.cp75,NULL,0,C1.cp75) 收回金額,tm78,tm79,tm80,tm81, C1.CP09,C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04 as FSort         " & _
                             " FROM (select a0l01,max(a0l02) as a0l02 from ACC0L0 group by a0l01) GG1,ACC0M0,CASEPROGRESS C1,CASEPROGRESS C2,TRADEMARK,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER       WHERE GG1.A0L01=A0M01(+) AND A0M02=C1.CP60(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C1.CP14=S1.ST01(+) AND C1.CP13=S2.ST01(+) AND C2.CP01=CPM01(+) AND C2.CP10=CPM02(+) AND C1.CP43=C2.CP09(+) AND " & SQLNewFag("TM23", "CU") & " " & strSQL2

strSql = strSql + " union all select decode(pa23,'1','','N')||C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),PA26) AS 申請人,NVL(DECODE(PA09,'000',CPM03,CPM04),C2.CP10) AS 相關案件性質,NVL(S1.ST02,C1.CP14) AS 承辦人,NVL(S2.ST02,C1.CP13) AS 智權人員,C1.CP16 AS 費用," & SqlDateT("GG2.A0L02") & " AS 收回日,DECODE(C1.cp75,NULL,0,C1.cp75) 收回金額,PA27,PA28,PA29,PA30, C1.CP09,C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04 as FSort " & _
                             " FROM (select a0l01,max(a0l02) as a0l02 from ACC0L0 group by a0l01) GG2,ACC0M0,CASEPROGRESS C1,CASEPROGRESS C2,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER          WHERE GG2.A0L01=A0M01(+) AND A0M02=C1.CP60(+) AND C1.CP01=PA01(+) AND C1.CP02=PA02(+) AND C1.CP03=PA03(+) AND C1.CP04=PA04(+) AND C1.CP14=S1.ST01(+) AND C1.CP13=S2.ST01(+) AND C2.CP01=CPM01(+) AND C2.CP10=CPM02(+) AND C1.CP43=C2.CP09(+) AND " & SQLNewFag("PA26", "CU") & " " & strSQL1
'edit by nickc 2006/12/11
'strSQL = strSQL + " union all select C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),SP08) AS 申請人,NVL(DECODE(SP09,'000',CPM03,CPM04),C2.CP10) AS 相關案件性質,NVL(S1.ST02,C1.CP14) AS 承辦人,NVL(S2.ST02,C1.CP13) AS 智權人員,C1.CP16 AS 費用," & SqlDateT("GG3.A0L02") & " AS 收回日,DECODE(C1.cp75,NULL,0,C1.cp75) 收回金額,SP58,SP59,'','', C1.CP09,C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04 as FSort     " & _
                             " FROM (select a0l01,max(a0l02) as a0l02 from ACC0L0 group by a0l01) GG3,ACC0M0,CASEPROGRESS C1,CASEPROGRESS C2,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER WHERE GG3.A0L01=A0M01(+) AND A0M02=C1.CP60(+) AND C1.CP01=SP01(+) AND C1.CP02=SP02(+) AND C1.CP03=SP03(+) AND C1.CP04=SP04(+) AND C1.CP14=S1.ST01(+) AND C1.CP13=S2.ST01(+) AND C2.CP01=CPM01(+) AND C2.CP10=CPM02(+) AND C1.CP43=C2.CP09(+) AND " & SQLNewFag("SP08", "CU") & " " & strSQL5
strSql = strSql + " union all select C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),SP08) AS 申請人,NVL(DECODE(SP09,'000',CPM03,CPM04),C2.CP10) AS 相關案件性質,NVL(S1.ST02,C1.CP14) AS 承辦人,NVL(S2.ST02,C1.CP13) AS 智權人員,C1.CP16 AS 費用," & SqlDateT("GG3.A0L02") & " AS 收回日,DECODE(C1.cp75,NULL,0,C1.cp75) 收回金額,SP58,SP59,sp65,sp66, C1.CP09,C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04 as FSort     " & _
                             " FROM (select a0l01,max(a0l02) as a0l02 from ACC0L0 group by a0l01) GG3,ACC0M0,CASEPROGRESS C1,CASEPROGRESS C2,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER WHERE GG3.A0L01=A0M01(+) AND A0M02=C1.CP60(+) AND C1.CP01=SP01(+) AND C1.CP02=SP02(+) AND C1.CP03=SP03(+) AND C1.CP04=SP04(+) AND C1.CP14=S1.ST01(+) AND C1.CP13=S2.ST01(+) AND C2.CP01=CPM01(+) AND C2.CP10=CPM02(+) AND C1.CP43=C2.CP09(+) AND " & SQLNewFag("SP08", "CU") & " " & strSQL5
'Modify By Sindy 2011/2/18 增加LC43,LC44,LC45,LC46
strSql = strSql + " union all select C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),LC11) AS 申請人,NVL(DECODE(LC15,'000',CPM03,CPM04),C2.CP10) AS 相關案件性質,NVL(S1.ST02,C1.CP14) AS 承辦人,NVL(S2.ST02,C1.CP13) AS 智權人員,C1.CP16 AS 費用," & SqlDateT("GG4.A0L02") & " AS 收回日,DECODE(C1.cp75,NULL,0,C1.cp75) 收回金額,LC43,LC44,LC45,LC46, C1.CP09,C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04 as FSort         " & _
                             " FROM (select a0l01,max(a0l02) as a0l02 from ACC0L0 group by a0l01) GG4,ACC0M0,CASEPROGRESS C1,CASEPROGRESS C2,LAWCASE,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER         WHERE GG4.A0L01=A0M01(+) AND A0M02=C1.CP60(+) AND C1.CP01=LC01(+) AND C1.CP02=LC02(+) AND C1.CP03=LC03(+) AND C1.CP04=LC04(+) AND C1.CP14=S1.ST01(+) AND C1.CP13=S2.ST01(+) AND C2.CP01=CPM01(+) AND C2.CP10=CPM02(+) AND C1.CP43=C2.CP09(+) AND " & SQLNewFag("LC11", "CU") & " " & StrSQL3
'Modify By Sindy 2011/2/18 增加HC24,HC25,HC26,HC27
strSql = strSql + " union all select C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),HC05) AS 申請人,NVL(CPM03,C2.CP10)                          AS 相關案件性質,NVL(S1.ST02,C1.CP14) AS 承辦人,NVL(S2.ST02,C1.CP13) AS 智權人員,C1.CP16 AS 費用," & SqlDateT("GG5.A0L02") & " AS 收回日,DECODE(C1.cp75,NULL,0,C1.cp75) 收回金額,HC24,HC25,HC26,HC27, C1.CP09,C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04 as FSort         " & _
                             " FROM (select a0l01,max(a0l02) as a0l02 from ACC0L0 group by a0l01) GG5,ACC0M0,CASEPROGRESS C1,CASEPROGRESS C2,HIRECASE,CASEPROPERTYMAP,STAFF S1,STAFF S2,CUSTOMER        WHERE GG5.A0L01=A0M01(+) AND A0M02=C1.CP60(+) AND C1.CP01=HC01(+) AND C1.CP02=HC02(+) AND C1.CP03=HC03(+) AND C1.CP04=HC04(+) AND C1.CP14=S1.ST01(+) AND C1.CP13=S2.ST01(+) AND C2.CP01=CPM01(+) AND C2.CP10=CPM02(+) AND C1.CP43=C2.CP09(+) AND " & SQLNewFag("HC05", "CU") & " " & StrSQL4
strSql = strSql & " order by FSort "
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/11/4
Else
    InsertQueryLog (0)  'Add By Sindy 2010/11/4
    Me.Enabled = True
    ShowNoData
    Screen.MousePointer = vbDefault
    '92.04.18 nick
    'Me.Hide
    tmpBol = fnCancelNowFormAndShowParentForm(Me)
    Exit Sub
End If
Me.grdDataList.Visible = False
Set grdDataList.Recordset = adoRecordset
'2010/9/15 CANCEL BY SONIA
'For ii = 1 To Me.grdDataList.Rows - 1
'    Me.grdDataList.TextMatrix(ii, 4) = Me.grdDataList.TextMatrix(ii, 4) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(ii, 4), "1")
'Next ii
Me.grdDataList.Visible = True
CheckOC
Me.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm10011202_2 = Nothing
End Sub
