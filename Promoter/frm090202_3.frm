VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090202_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利／商標會稿"
   ClientHeight    =   6120
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   9930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   9930
   Begin TabDlg.SSTab SSTab1 
      Height          =   5385
      Left            =   30
      TabIndex        =   5
      Top             =   480
      Width           =   9855
      _ExtentX        =   17374
      _ExtentY        =   9507
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "未會稿區"
      TabPicture(0)   =   "frm090202_3.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Grd1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Combo2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "已會稿區"
      TabPicture(1)   =   "frm090202_3.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Grd1(1)"
      Tab(1).ControlCount=   1
      Begin VB.ComboBox Combo2 
         Height          =   260
         ItemData        =   "frm090202_3.frx":0038
         Left            =   1050
         List            =   "frm090202_3.frx":0048
         Style           =   2  '單純下拉式
         TabIndex        =   7
         Top             =   300
         Width           =   1350
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd1 
         Height          =   4695
         Index           =   0
         Left            =   60
         TabIndex        =   6
         Top             =   630
         Width           =   9720
         _ExtentX        =   17145
         _ExtentY        =   8273
         _Version        =   393216
         Cols            =   12
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V| 流程日期| 本所案號 |案件名稱| 國家| 種類|  案件性質| 本所期限|承辦人|承辦期限|目前流程狀態|不顯示"
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
         _Band(0).Cols   =   12
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd1 
         Height          =   4965
         Index           =   1
         Left            =   -74940
         TabIndex        =   9
         Top             =   360
         Width           =   9720
         _ExtentX        =   17145
         _ExtentY        =   8767
         _Version        =   393216
         Cols            =   12
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V| 流程日期| 本所案號 |案件名稱| 國家| 種類|  案件性質| 本所期限|承辦人|承辦期限|目前流程狀態|不顯示"
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
         _Band(0).Cols   =   12
      End
      Begin VB.Label Label2 
         Caption         =   "註：在”不顯示”欄位上點一下(V)，可以取消顯示聯絡歷程。"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   4740
         TabIndex        =   11
         Top             =   390
         Width           =   4935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "最近聯絡："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "畫面更新(&Q)"
      Height          =   360
      Left            =   6180
      TabIndex        =   0
      Top             =   60
      Width           =   1155
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "明細資料(&D)"
      Height          =   360
      Left            =   7380
      TabIndex        =   1
      Top             =   60
      Width           =   1305
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   8715
      TabIndex        =   2
      Top             =   60
      Width           =   800
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1110
      TabIndex        =   10
      Top             =   90
      Width           =   2430
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "4286;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label16 
      Caption         =   "註：雙擊選取時，開啟承辦歷程。　多案歷程：案件名稱欄位顯示紫紅色。   "
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   150
      TabIndex        =   4
      Top             =   5880
      Width           =   8175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "會稿人員："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   150
      Width           =   900
   End
End
Attribute VB_Name = "frm090202_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/14 Form2.0已修改
'Create by Sindy 2013/4/30
'Memo by Lydia 2019/07/01 表單名稱:待會稿區=>專利／商標會稿
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim i As Integer
Dim dblPrevRow As Double
Dim m_EMPCon As String 'Add By Sindy 2017/1/18
Const Show會稿方式 = "'1','EMail','2','紙本','3','微信','4','LINE','9','其他'"

'明細資料
Private Sub cmdDetail_Click()
Dim ii As Integer
Dim nFrm As Form
Dim Index As Integer
   
   'Modify By Sindy 2018/8/29 分未會稿及已會稿2區
   Index = SSTab1.Tab
   For i = 1 To GRD1(Index).Rows - 1
      If GRD1(Index).TextMatrix(i, 0) = "V" Then
'         'Add By Sindy 2017/9/19
'         '檢查表單是否已開啟，若是，則關閉
'         For Each nFrm In Forms
'            If StrComp(nFrm.Name, "frm090202_2", vbTextCompare) = 0 Then
'               Unload frm090202_2
'               If strSaveConfirm = True Then frm090202_2.ZOrder: Exit Sub 'Add By Sindy 2020/1/17 有資料要儲存,尚需處理...
'            End If
'         Next
'         '2017/9/19 END
         If PUB_ChkFormIsClose("frm090202_2") = False Then Exit Sub 'Add By Sindy 2020/1/17
         frm090202_2.Hide
         frm090202_2.m_EEP01 = GRD1(Index).TextMatrix(i, 11) '總收文號
         
         'Add By Sindy 2017/1/18 若點選的是聯終歷程,必須檢查是否有待回覆歷程要操作 ex:P-116465 聯絡/送會
         If GRD1(Index).TextMatrix(i, 10) = "聯絡" Then
            strSql = "select eep02,eep05 from empelectronprocess" & _
                     " where eep04 in(" & m_EMPCon & ") and eep09='Y'" & _
                     " and eep01='" & GRD1(Index).TextMatrix(i, 11) & "'" & _
                     " order by eep06 asc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               RsTemp.MoveFirst
               Do While Not RsTemp.EOF
                  For ii = 0 To Combo1.ListCount - 1
                     If Trim(Left(Combo1.List(ii), 6)) = RsTemp.Fields("eep05") Then
                        frm090202_2.m_FlowUserNum = RsTemp.Fields("eep05")
                        frm090202_2.m_CurrFlowEEP02 = RsTemp.Fields("eep02")
                        GoTo carry_on
                     End If
                  Next ii
                  RsTemp.MoveNext
               Loop
            End If
         End If
         '2017/1/18 END
         
         frm090202_2.m_FlowUserNum = Trim(Left("" & Combo1.Text, 6)) 'Add By Sindy 2013/9/12 案件流程所屬人員
         frm090202_2.m_CurrFlowEEP02 = GRD1(Index).TextMatrix(i, 13) 'Add By Sindy 2016/3/25 目前要處理的歷程序號
carry_on: 'Add By Sindy 2017/1/18
         frm090202_2.intReceiveKind = 2
         frm090202_2.SetParent Me
         If frm090202_2.QueryData = True Then
            frm090202_2.Show
            Me.Hide
         End If
         Exit For
      End If
   Next i
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Public Function QueryData(Index As Integer) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strVal As String
Dim strQuyDate As String, strConSql As String
   
   m_blnColOrderAsc = True
   QueryData = True
   
   GRD1(Index).Clear
   SetGrd

   Screen.MousePointer = vbHourglass
   'Modify By Sindy 2014/2/26 因智權人員原是96027.林佳芳改為94026.林建志; 因此調整程式,不管送會收受者
   '" and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" &
   'Modify By Sindy 2016/3/3 調整聯絡的抓法
   'Modify By Sindy 2016/3/15 +EMP_會圖
   'Modify By Sindy 2016/3/24 + 因智權人員離職的狀況,不可判斷CP13改抓EEP05
   m_EMPCon = "'" & EMP_送會 & "','" & EMP_會圖 & "'"
   'Add By Sindy 2013/9/17
   'Modify By Sindy 2018/8/29
   If Index = 0 Then '未會稿區
   '2018/8/29 END
      If Combo2.ListIndex = 0 Then
         strQuyDate = CompWorkDay(3, strSrvDate(1), 1) '不含當天,3個工作天
      ElseIf Combo2.ListIndex = 1 Then
         strQuyDate = CompWorkDay(5, strSrvDate(1), 1) '不含當天,5個工作天
      ElseIf Combo2.ListIndex = 2 Then
         strQuyDate = CompWorkDay(7, strSrvDate(1), 1) '不含當天,7個工作天
      Else
         '全部
      End If
      'Modify By Sindy 2018/4/13 加入商標檔及服務業務檔
      'Modify By Sindy 2018/10/4 加入讀取重新送會後未客戶會稿 ex:P-120779簡國靜提
      strVal = "select EEP01,max(EEP02) as EEP02 from(" & _
               "select EEP01,EEP02,EEP04 from EmpElectronProcess,engineerprogress" & _
               " where EEP09='Y'" & _
               " And EEP04 in(" & m_EMPCon & ")" & _
               " And EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "' and eep01=ep02 and ep10 is null and nvl(ep37,0)=0" & _
               " Union" & _
               " select E08.EEP01,E08.EEP02,E08.EEP04 from EmpElectronProcess E08,engineerprogress" & _
               " where E08.EEP09='Y'" & _
               " And E08.EEP04 in(" & m_EMPCon & ")" & _
               " And E08.EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "' and E08.eep01=ep02 and ep10 is null and nvl(ep37,0)>0" & _
               " And not exists(select E35.EEP01 from EmpElectronProcess E35 where E35.eep01=E08.eep01 and E35.eep02>E08.eep02 and E35.eep04='" & EMP_客戶會稿 & "')" & _
               " Union" & _
               " select EEP01,EEP02,EEP04 from EmpElectronProcess" & _
               " where EEP13='Y'" & _
               " and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
               " And EEP04='" & EMP_聯絡 & "'" & _
               IIf(strQuyDate <> "", " And EEP06>=" & strQuyDate, "") & _
               ") group by EEP01"
      'strConSql = " And e1.eep01=(select ep02 from engineerprogress where e1.eep01=ep02 and ep37 is null)"
   Else '已會稿區
      'Modify By Sindy 2018/9/5 解決查詢慢的問題 + Union select null,null,null from dual where 1=0
'      strVal = "select EEP01,max(EEP02) as EEP02 from(" & _
'               "select EEP01,EEP02,EEP04 from EmpElectronProcess,engineerprogress" & _
'               " where EEP09='Y'" & _
'               " And EEP04 in(" & m_EMPCon & ")" & _
'               " And EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "' and eep01=ep02(+) and ep37>0" & _
'               " Union select null,null,null from dual where 1=0" & _
'               ") group by EEP01"
      'Modify By Sindy 2018/10/4 排除重新送會未客戶會稿 ex:P-120779簡國靜提
      strVal = "select EEP01,max(EEP02) as EEP02 from EmpElectronProcess" & _
               " where eep01 in(select E08.EEP01 from EmpElectronProcess E08,engineerprogress" & _
               " where E08.EEP09='Y'" & _
               " And E08.EEP04 in(" & m_EMPCon & ")" & _
               " And E08.EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "' and E08.eep01=ep02 and ep10 is null and nvl(ep37,0)>0" & _
               " And exists(select E35.EEP01 from EmpElectronProcess E35 where E35.eep01=E08.eep01 and E35.eep02>E08.eep02 and E35.eep04='" & EMP_客戶會稿 & "')" & _
               ")" & _
               " And EEP04='" & EMP_客戶會稿 & "' group by EEP01"
      'strConSql = " And e1.eep01=(select ep02 from engineerprogress where e1.eep01=ep02 and ep37>0)"
   End If
   '2013/9/17 END
   
   'Add By Sindy 2019/9/5 客服組專利會稿工程師,只能操作專利案
   If InStr(Pub_GetSpecMan("客服組專利會稿工程師"), strUserNum) > 0 Then
      strConSql = strConSql & " And cp01 in('P','PS','CFP','CPS')"
   End If
   '2019/9/5 END
   
   'Modify By Sindy 2016/3/3 +不顯示,e2.EEP02
   'Modify By Sindy 2016/3/24 + 因智權人員離職的狀況,不可判斷CP13固取消「And cp13='" & Trim(Left("" & Combo1.Text, 6)) & "'"」
   'Modify By Sindy 2016/9/2 And cp27 is null And cp57 is null -> and cp158=0 and cp159=0
   'Modify By Sindy 2018/4/13 加入商標檔及服務業務檔
   'Modify By Sindy 2020/12/1 + ,e1.eep15 AS eep15,e1.eep11 AS eep11
   strSql = "Select ' ' as V,SqlDateT(e1.EEP06)||' '||sqltime(e1.EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,PA05||PA06||PA07 as 案件名稱," & _
            "NA03 as 國家,Decode(PA09,'000',PTM03,PTM04) as ""類別/種類"",Decode(PA09,'000',CPM03,CPM04) as 案件性質,SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人," & _
            "SqlDateT(cp48) as 承辦期限, ac03 as 目前流程狀態,e1.EEP01 as 總收文號,' ' as 不顯示,e1.EEP02,e1.eep06 a,e1.eep07 b,decode(eep14," & Show會稿方式 & ",eep14) as 會稿方式,e1.eep15 AS eep15,e1.eep11 AS eep11" & _
            " from EmpElectronProcess e1,CaseProgress,Patent,staff s1,nation,CasePropertyMap,PatentTradeMarkMap,allcode" & _
            " where (e1.eep01,e1.eep02) in(" & strVal & ")" & _
            " AND e1.EEP01=CP09" & _
            " And CP01=PA01(+) And CP02=PA02(+) And CP03=PA03(+) And CP04=PA04(+) AND PA01 is not null" & _
            " And CP14=s1.ST01(+) And PA09=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And '1'=PTM01(+) AND PA08=PTM02(+)" & _
            " And ac01='09' And e1.EEP04=ac02(+)" & _
            " And cp158=0 And cp159=0" & strConSql
   'Modify By Sindy 2018/10/11 Decode(TM10,'000',PTM03,PTM04)種類 ==> TM09商品類別
   strSql = strSql & " union " & _
            "Select ' ' as V,SqlDateT(e1.EEP06)||' '||sqltime(e1.EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,TM05||TM06||TM07 as 案件名稱," & _
            "NA03 as 國家,TM09 as ""類別/種類"",Decode(TM10,'000',CPM03,CPM04) as 案件性質,SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人," & _
            "SqlDateT(cp48) as 承辦期限, ac03 as 目前流程狀態,e1.EEP01 as 總收文號,' ' as 不顯示,e1.EEP02,e1.eep06 a,e1.eep07 b,decode(eep14," & Show會稿方式 & ",eep14) as 會稿方式,e1.eep15 AS eep15,e1.eep11 AS eep11" & _
            " from EmpElectronProcess e1,CaseProgress,Trademark,staff s1,nation,CasePropertyMap,PatentTradeMarkMap,allcode" & _
            " where (e1.eep01,e1.eep02) in(" & strVal & ")" & _
            " AND e1.EEP01=CP09" & _
            " And CP01=TM01(+) And CP02=TM02(+) And CP03=TM03(+) And CP04=TM04(+) AND TM01 is not null" & _
            " And CP14=s1.ST01(+) And TM10=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And '2'=PTM01(+) AND TM08=PTM02(+)" & _
            " And ac01='09' And e1.EEP04=ac02(+)" & _
            " And cp158=0 And cp159=0" & strConSql
   strSql = strSql & " union " & _
            "Select ' ' as V,SqlDateT(e1.EEP06)||' '||sqltime(e1.EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,SP05||SP06||SP07 as 案件名稱," & _
            "NA03 as 國家,'' as ""類別/種類"",Decode(SP09,'000',CPM03,CPM04) as 案件性質,SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人," & _
            "SqlDateT(cp48) as 承辦期限, ac03 as 目前流程狀態,e1.EEP01 as 總收文號,' ' as 不顯示,e1.EEP02,e1.eep06 a,e1.eep07 b,decode(eep14," & Show會稿方式 & ",eep14) as 會稿方式,e1.eep15 AS eep15,e1.eep11 AS eep11" & _
            " from EmpElectronProcess e1,CaseProgress,servicepractice,staff s1,nation,CasePropertyMap,allcode" & _
            " where (e1.eep01,e1.eep02) in(" & strVal & ")" & _
            " AND e1.EEP01=CP09" & _
            " And CP01=SP01(+) And CP02=SP02(+) And CP03=SP03(+) And CP04=SP04(+) AND SP01 is not null" & _
            " And CP14=s1.ST01(+) And SP09=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And ac01='09' And e1.EEP04=ac02(+)" & _
            " And cp158=0 And cp159=0" & strConSql
   'Add By Sindy 2021/7/14
   strSql = strSql & " union " & _
            "Select ' ' as V,SqlDateT(e1.EEP06)||' '||sqltime(e1.EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,LC05||LC06||LC07 as 案件名稱," & _
            "NA03 as 國家,'' as ""類別/種類"",Decode(LC15,'000',CPM03,CPM04) as 案件性質,SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人," & _
            "SqlDateT(cp48) as 承辦期限, ac03 as 目前流程狀態,e1.EEP01 as 總收文號,' ' as 不顯示,e1.EEP02,e1.eep06 a,e1.eep07 b,decode(eep14," & Show會稿方式 & ",eep14) as 會稿方式,e1.eep15 AS eep15,e1.eep11 AS eep11" & _
            " from EmpElectronProcess e1,CaseProgress,LawCase,staff s1,nation,CasePropertyMap,allcode" & _
            " where (e1.eep01,e1.eep02) in(" & strVal & ")" & _
            " AND e1.EEP01=CP09" & _
            " And CP01=LC01(+) And CP02=LC02(+) And CP03=LC03(+) And CP04=LC04(+) AND LC01 is not null" & _
            " And CP14=s1.ST01(+) And LC15=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And ac01='09' And e1.EEP04=ac02(+)" & _
            " And cp158=0 And cp159=0" & strConSql
   strSql = strSql & " union " & _
            "Select ' ' as V,SqlDateT(e1.EEP06)||' '||sqltime(e1.EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,HC06 as 案件名稱," & _
            "NA03 as 國家,'' as ""類別/種類"",Decode('000','000',CPM03,CPM04) as 案件性質,SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人," & _
            "SqlDateT(cp48) as 承辦期限, ac03 as 目前流程狀態,e1.EEP01 as 總收文號,' ' as 不顯示,e1.EEP02,e1.eep06 a,e1.eep07 b,decode(eep14," & Show會稿方式 & ",eep14) as 會稿方式,e1.eep15 AS eep15,e1.eep11 AS eep11" & _
            " from EmpElectronProcess e1,CaseProgress,HireCase,staff s1,nation,CasePropertyMap,allcode" & _
            " where (e1.eep01,e1.eep02) in(" & strVal & ")" & _
            " AND e1.EEP01=CP09" & _
            " And CP01=HC01(+) And CP02=HC02(+) And CP03=HC03(+) And CP04=HC04(+) AND HC01 is not null" & _
            " And CP14=s1.ST01(+) And '000'=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And ac01='09' And e1.EEP04=ac02(+)" & _
            " And cp158=0 And cp159=0" & strConSql
   '2021/7/14 END
   strSql = strSql & " order by a desc,b desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1(Index).Recordset = rsTmp
      'Add By Sindy 2020/12/1
      For i = 1 To GRD1(Index).Rows - 1
         Call SetColColor(i, Index)
      Next i
      '2020/12/1 END
   Else
      QueryData = False
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
      'Modify By Sindy 2024/7/18
      'ShowNoData
      If Index = 0 Then '未會稿區
         MsgBox "未會稿區，尚無資料 !!"
      Else
         MsgBox "已會稿區，尚無資料 !!"
      End If
      '2024/7/18 END
      Exit Function
   End If

   '若有資料游標停在第一筆
   GRD1(Index).Visible = False
   GRD1(Index).col = 0
   GRD1(Index).row = 1
   dblPrevRow = GRD1(Index).row
   If rsTmp.RecordCount > 0 Then
      GRD1(Index).Text = "V"
      For i = 0 To GRD1(Index).Cols - 1
         'Modify By Sindy 2020/12/1
         If i <> 3 Then
         '2020/12/1 END
            GRD1(Index).col = i
            GRD1(Index).CellBackColor = &HFFC0C0
         End If
      Next i
   End If
   GRD1(Index).Visible = True

   rsTmp.Close
   Screen.MousePointer = vbDefault

EXITSUB:
   Set rsTmp = Nothing
'Dim rsTmp As New ADODB.Recordset
'Dim strSql As String
'Dim strVal As String
'Dim strQuyDate As String
'
'   m_blnColOrderAsc = True
'   QueryData = True
'
'   'Add By Sindy 2013/9/17
'   'Modify By Sindy 2018/8/29
'   If index = 0 Then '未會稿區
'   '2018/8/29 END
'      If Combo2.ListIndex = 0 Then
'         strQuyDate = CompWorkDay(3, strSrvDate(1), 1) '不含當天,3個工作天
'      ElseIf Combo2.ListIndex = 1 Then
'         strQuyDate = CompWorkDay(5, strSrvDate(1), 1) '不含當天,5個工作天
'      ElseIf Combo2.ListIndex = 2 Then
'         strQuyDate = CompWorkDay(7, strSrvDate(1), 1) '不含當天,7個工作天
'      Else
'         '全部
'      End If
'   Else '已會稿區
'   End If
'   '2013/9/17 END
'
'   Grd1(index).Clear
'   SetGrd
'
'   Screen.MousePointer = vbHourglass
''   strSql = "Select ' ' as V,SqlDateT(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,PA05||PA06||PA07 as 案件名稱," & _
''            "NA03 as 國家,Decode(PA09,'000',PTM03,PTM04) as 種類,Decode(PA09,'000',CPM03,CPM04) as 案件性質," & _
''            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限," & _
''            "EEP01 as 總收文號" & _
''            " From EmpElectronProcess,CaseProgress,Patent," & _
''            "staff s1,nation,CasePropertyMap,PatentTradeMarkMap,allcode" & _
''            " Where EEP09='Y' and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
''            " And EEP01=CP09(+)" & _
''            " And CP01=PA01(+) And CP02=PA02(+) And CP03=PA03(+) And CP04=PA04(+)" & _
''            " And CP14=s1.ST01(+)" & _
''            " And PA09=NA01(+)" & _
''            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
''            " And '1'=PTM01(+) AND PA08=PTM02(+)" & _
''            " And ac01='09' And EEP04=ac02(+)" & _
''            " And EEP04 in('" & EMP_送會 & "')" & _
''            " order by cp01,cp02,cp03,cp04 asc"
'   'Add By Sindy 2013/9/17 +IIf(strQuyDate <> "", " And e1.EEP06 >=" & strQuyDate, "")
'
'   'Modify By Sindy 2014/2/26 因智權人員原是96027.林佳芳改為94026.林建志; 因此調整程式,不管送會收受者
'   '" and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" &
'   'Modify By Sindy 2016/3/3 調整聯絡的抓法
''            " Union" & _
''            " select EEP01,EEP02,EEP04 from EmpElectronProcess e1" & _
''            " where e1.EEP02 in (select max(eep02) from EmpElectronProcess where eep01=e1.eep01)" & _
''            " and e1.EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
''            " And e1.EEP04 in('" & EMP_聯絡 & "')" & _
''            IIf(strQuyDate <> "", " And e1.EEP06>=" & strQuyDate, "")
'   'Modify By Sindy 2016/3/15 +EMP_會圖
'   'Modify By Sindy 2016/3/24 + 因智權人員離職的狀況,不可判斷CP13改抓EEP05
'
'   m_EMPCon = "'" & EMP_送會 & "','" & EMP_會圖 & "'"
'   'Modify By Sindy 2018/4/13 加入商標檔及服務業務檔
''   strVal = "(select EEP01,max(EEP02) as EEP02 from(" & _
''            "select EEP01,EEP02,EEP04 from EmpElectronProcess" & _
''            " where EEP09='Y'" & _
''            " And EEP04 in(" & m_EMPCon & ")" & _
''            " And EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
''            " Union" & _
''            " select EEP01,EEP02,EEP04 from EmpElectronProcess e1" & _
''            " where e1.EEP13='Y'" & _
''            " and e1.EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
''            " And e1.EEP04 in('" & EMP_聯絡 & "')" & _
''            IIf(strQuyDate <> "", " And e1.EEP06>=" & strQuyDate, "") & _
''            ") group by EEP01) V1"
'   strVal = "select EEP01,max(EEP02) as EEP02 from(" & _
'            "select EEP01,EEP02,EEP04 from EmpElectronProcess" & _
'            " where EEP09='Y'" & _
'            " And EEP04 in(" & m_EMPCon & ")" & _
'            " And EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
'            " Union" & _
'            " select EEP01,EEP02,EEP04 from EmpElectronProcess" & _
'            " where EEP13='Y'" & _
'            " and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
'            " And EEP04 in('" & EMP_聯絡 & "')" & _
'            IIf(strQuyDate <> "", " And EEP06>=" & strQuyDate, "") & _
'            ") group by EEP01"
'   'Modify By Sindy 2016/3/3 +不顯示,e2.EEP02
'   'Modify By Sindy 2016/3/24 + 因智權人員離職的狀況,不可判斷CP13固取消「And cp13='" & Trim(Left("" & Combo1.Text, 6)) & "'"」
'   'Modify By Sindy 2016/9/2 And cp27 is null And cp57 is null -> and cp158=0 and cp159=0
'   'Modify By Sindy 2018/4/13 加入商標檔及服務業務檔
''   strSql = "Select ' ' as V,SqlDateT(e2.EEP06)||' '||sqltime(e2.EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,PA05||PA06||PA07 as 案件名稱," & _
''            "NA03 as 國家,Decode(PA09,'000',PTM03,PTM04) as 種類,Decode(PA09,'000',CPM03,CPM04) as 案件性質," & _
''            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, ac03 as 目前流程狀態," & _
''            "e2.EEP01 as 總收文號,' ' as 不顯示,e2.EEP02" & _
''            " From EmpElectronProcess e2," & strVal & ",CaseProgress,Patent," & _
''            "staff s1,nation,CasePropertyMap,PatentTradeMarkMap,allcode" & _
''            " Where V1.EEP01=e2.EEP01 AND V1.EEP02=e2.EEP02 AND e2.EEP01=CP09(+)" & _
''            " And CP01=PA01(+) And CP02=PA02(+) And CP03=PA03(+) And CP04=PA04(+)" & _
''            " And CP14=s1.ST01(+)" & _
''            " And PA09=NA01(+)" & _
''            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
''            " And '1'=PTM01(+) AND PA08=PTM02(+)" & _
''            " And ac01='09' And e2.EEP04=ac02(+)" & _
''            " and cp158=0 and cp159=0" & _
''            " order by e2.EEP06 desc,e2.EEP07 desc"
''            '" order by cp01,cp02,cp03,cp04 asc"
'   strSql = "Select ' ' as V,SqlDateT(e1.EEP06)||' '||sqltime(e1.EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,PA05||PA06||PA07 as 案件名稱," & _
'            "NA03 as 國家,Decode(PA09,'000',PTM03,PTM04) as 種類,Decode(PA09,'000',CPM03,CPM04) as 案件性質,SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人," & _
'            "SqlDateT(cp48) as 承辦期限, ac03 as 目前流程狀態,e1.EEP01 as 總收文號,' ' as 不顯示,e1.EEP02,e1.eep06 a,e1.eep07 b" & _
'            " from EmpElectronProcess e1,CaseProgress,Patent,staff s1,nation,CasePropertyMap,PatentTradeMarkMap,allcode" & _
'            " where (e1.eep01,e1.eep02) in(" & strVal & ")" & _
'            " AND e1.EEP01=CP09(+)" & _
'            " And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04" & _
'            " And CP14=s1.ST01(+) And PA09=NA01(+)" & _
'            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
'            " And '1'=PTM01(+) AND PA08=PTM02(+)" & _
'            " And ac01='09' And e1.EEP04=ac02(+)" & _
'            " and cp158=0 and cp159=0"
'   strSql = strSql & " union " & _
'            "Select ' ' as V,SqlDateT(e1.EEP06)||' '||sqltime(e1.EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,TM05||TM06||TM07 as 案件名稱," & _
'            "NA03 as 國家,Decode(TM10,'000',PTM03,PTM04) as 種類,Decode(TM10,'000',CPM03,CPM04) as 案件性質,SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人," & _
'            "SqlDateT(cp48) as 承辦期限, ac03 as 目前流程狀態,e1.EEP01 as 總收文號,' ' as 不顯示,e1.EEP02,e1.eep06 a,e1.eep07 b" & _
'            " from EmpElectronProcess e1,CaseProgress,Trademark,staff s1,nation,CasePropertyMap,PatentTradeMarkMap,allcode" & _
'            " where (e1.eep01,e1.eep02) in(" & strVal & ")" & _
'            " AND e1.EEP01=CP09(+)" & _
'            " And CP01=TM01 And CP02=TM02 And CP03=TM03 And CP04=TM04" & _
'            " And CP14=s1.ST01(+) And TM10=NA01(+)" & _
'            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
'            " And '2'=PTM01(+) AND TM08=PTM02(+)" & _
'            " And ac01='09' And e1.EEP04=ac02(+)" & _
'            " and cp158=0 and cp159=0"
'   strSql = strSql & " union " & _
'            "Select ' ' as V,SqlDateT(e1.EEP06)||' '||sqltime(e1.EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,SP05||SP06||SP07 as 案件名稱," & _
'            "NA03 as 國家,'' as 種類,Decode(SP09,'000',CPM03,CPM04) as 案件性質,SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人," & _
'            "SqlDateT(cp48) as 承辦期限, ac03 as 目前流程狀態,e1.EEP01 as 總收文號,' ' as 不顯示,e1.EEP02,e1.eep06 a,e1.eep07 b" & _
'            " from EmpElectronProcess e1,CaseProgress,servicepractice,staff s1,nation,CasePropertyMap,allcode" & _
'            " where (e1.eep01,e1.eep02) in(" & strVal & ")" & _
'            " AND e1.EEP01=CP09(+)" & _
'            " And CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04" & _
'            " And CP14=s1.ST01(+) And SP09=NA01(+)" & _
'            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
'            " And ac01='09' And e1.EEP04=ac02(+)" & _
'            " and cp158=0 and cp159=0"
'   strSql = strSql & " order by a desc,b desc"
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      Set Grd1(index).Recordset = rsTmp
'   Else
'      QueryData = False
'      Screen.MousePointer = vbDefault
'      rsTmp.Close
'      Set rsTmp = Nothing
'      ShowNoData
'      Exit Function
'   End If
'
'   '若有資料游標停在第一筆
'   Grd1(index).Visible = False
'   Grd1(index).col = 0
'   Grd1(index).row = 1
'   dblPrevRow = Grd1(index).row
'   If rsTmp.RecordCount > 0 Then
'      Grd1(index).Text = "V"
'      For i = 0 To Grd1(index).Cols - 1
'         Grd1(index).col = i
'         Grd1(index).CellBackColor = &HFFC0C0
'      Next i
'   End If
'   Grd1(index).Visible = True
'
'   rsTmp.Close
'   Screen.MousePointer = vbDefault
'
'EXITSUB:
'   Set rsTmp = Nothing
End Function

'畫面更新
Private Sub cmdQuery_Click()
   'If QueryData = False Then ShowNoData
   'Call Combo1_Click
   Call QueryData(SSTab1.Tab)
End Sub

'Add By Sindy 2023/1/17
Private Sub Combo1_Change()
   If Me.Visible = True Then Call QueryData(SSTab1.Tab)
End Sub

Private Sub Combo2_Click()
   If Me.Visible = True Then Call QueryData(SSTab1.Tab) 'Add By Sindy 2023/4/12
End Sub

Private Sub Form_Activate()
'Dim nFrm As Form
'
'   'Add By Sindy 2017/9/5
'   '檢查表單是否已開啟，若是，則關閉
'   If Me.Visible = True Then
'      For Each nFrm In Forms
'         If StrComp(nFrm.Name, "frm090202_2", vbTextCompare) = 0 Then
'            If UCase(frm090202_2.m_PrevForm.Name) <> UCase(Me.Name) Then Exit For
'            Unload frm090202_2
'         End If
'      Next
'   End If
'   '2017/9/5 END
End Sub

''Add By Sindy 2013/9/12
'Private Sub Combo1_Click()
'Dim strST01 As String
'
'   If bolCombo1Run1 = False Then
'      If Combo1.Text <> "" Then strST01 = Trim(Left(Combo1.Text, 6))
'      Combo1.Clear
'      Combo1.AddItem strUserNum & " " & strUserName
'      '檢查當時是否需要為他人職代
'      Call Pub_SetForOthersEmpCombo(strUserNum, Combo1, False)
'      For i = 0 To Combo1.ListCount - 1
'         If Trim(Left(Combo1.List(i), 6)) = strST01 Then
'            bolCombo1Run1 = True
'            Combo1.ListIndex = i
'            GoTo RunEnd
'         End If
'      Next i
'      bolCombo1Run1 = True
'      If Combo1.Text = "" Then Combo1.ListIndex = 0
'      GoTo RunEnd
'   End If
'
'   Call QueryData
'RunEnd:
'   bolCombo1Run1 = False
'End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   'Combo1.Clear
   'Call Combo1_Click
   SSTab1.Tab = 0
   Call SetCombo1
   Combo2.Text = Combo2.List(3) 'Add By Sindy 2013/9/17
   Call QueryData(SSTab1.Tab)
   
   If strSrvDate(1) < T商標電子化第2階段啟用日 Then
      Label16.Caption = "註：雙擊選取時，開啟承辦歷程。"
   End If
End Sub

Private Sub SetCombo1()
Dim strTemp As String, arrData As Variant
Dim ii As Integer
   
   'Modify By Sindy 2022/5/25 設定屬智權人員作業的下拉選單(共用模組)
   Call PUB_SetCombo1Sales(Combo1)
   
'   Combo1.Clear
'   Combo1.AddItem strUserNum & " " & strUserName
'   '檢查當時是否需要為他人職代
'   Call Pub_SetForOthersEmpCombo(strUserNum, Combo1, False)
''   'Add By Sindy 2013/12/18 暫時的智權人員ID其部門主管
''   strSql = "select st01,st02,a0901,a0908 from acc090,staff where substr(a0901,1,1)='S'" & _
''            " and substr(st01,1,1)<'6'" & _
''            " and st03=a0901" & _
''            " and st01<>'001-1'" & _
''            " and a0908='" & strUserNum & "' and st01 not in('00001','00002','00003','00004')"
''   intI = 1
''   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
''   If intI = 1 Then
''      RsTemp.MoveFirst
''      Do While Not RsTemp.EOF
''         Combo1.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
''         RsTemp.MoveNext
''      Loop
''   End If
''   '2013/12/18 END
'   'Modify By Sindy 2014/9/4 帶人的權限
'   Call Pub_SetSAManageEmpCombo(strUserNum, Combo1, False, , True)
'   '2014/9/4 END
'
'   'Add By Sindy 2017/11/9
'   'Added by Morgan 2014/5/15
'   '專利處智權同仁代處理人
'   'Modify by Amy 2015/03/13 +特殊設定(總經理業務工作代理人員)
'   If InStr(Pub_GetSpecMan("A8"), strUserNum) > 0 Or InStr(Pub_GetSpecMan("總經理業務工作代理人員"), strUserNum) > 0 Then
'      If InStr(Pub_GetSpecMan("A8"), strUserNum) > 0 Then
'        strSql = "select st01,st02 from setSpecMan,staff where ocode='A7' and instr(';'||replace(oMan,',',';')||';',';'||st01||';')>0"
'      Else
'        strSql = "select st01,st02 from setSpecMan,staff where ocode='總經理員工編號' and instr(';'||replace(oMan,',',';')||';',';'||st01||';')>0"
'      End If
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         RsTemp.MoveFirst
'         Do While Not RsTemp.EOF
'            For ii = 0 To Combo1.ListCount - 1
'               If InStr(Combo1.List(ii), RsTemp(0)) = 1 Then
'                  Exit For
'               End If
'            Next
'            If ii = Combo1.ListCount Then
'               Combo1.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
'            End If
'
'            RsTemp.MoveNext
'         Loop
'      End If
'   End If
'   'end 2014/5/15
   
   'Add By Sindy 2018/9/21 檢查是否有處理商標處MCT案件的權限
   'Add By Sindy 2019/4/26 改成共用函數
   Call PUB_AddComboMCTF(strUserNum, Combo1)
''oCode         OMAN
''---------     --------------
''MCTF01        86048;A4021
''MCTF02        98019;A4022
''MCTF03        A6023
''MCTM          67002;69008
'   If InStr(Pub_GetSpecMan("MCTM"), strUserNum) > 0 Then
'      Combo1.AddItem "MCTF01"
'      Combo1.AddItem "MCTF02"
'      Combo1.AddItem "MCTF03"
'   Else
'      If InStr(Pub_GetSpecMan("MCTF01"), strUserNum) > 0 Then
'         Combo1.AddItem "MCTF01"
'      End If
'      If InStr(Pub_GetSpecMan("MCTF02"), strUserNum) > 0 Then
'         Combo1.AddItem "MCTF02"
'      End If
'      If InStr(Pub_GetSpecMan("MCTF03"), strUserNum) > 0 Then
'         Combo1.AddItem "MCTF03"
'      End If
'   End If
   '2018/9/21 END
   
'   'Add By Sindy 2014/5/29 開放部份智權同仁的資料給彥葶操作
'   'If Pub_GetSpecMan("A8") = strUserNum Then
'   If InStr(Pub_GetSpecMan("A8"), strUserNum) > 0 Then
'      strTemp = Pub_GetSpecMan("A7")
'      arrData = Split(strTemp, ";")
'      For i = 0 To UBound(arrData)
'         For ii = 0 To Combo1.ListCount - 1
'            If InStr(Combo1.List(ii), arrData(i)) = 1 Then
'               Exit For
'            End If
'         Next
'         If ii = Combo1.ListCount Then
'            Combo1.AddItem arrData(i) & " " & GetPrjSalesNM(CStr(arrData(i)))
'         End If
'      Next
'   End If
'   '2014/5/29 END
   
   'Add By Sindy 2019/9/5 開放專利工程師可以做客服組案件的會稿
   If InStr(Pub_GetSpecMan("客服組專利會稿工程師"), strUserNum) > 0 Then
      Combo1.AddItem "W1001" & " " & GetPrjSalesNM("W1001")
   End If
   '2019/9/5 END
   
'   'Add By Sindy 2014/8/8
'   '帶人主管抓虛建編號 ex.86047.高國碩,要帶出20011.中一區
'   strSql = "select st01,st02 from staff where st01<'63001' and instr(';'||st52||';'||st53||';'||st54||';'||st55||';',';" & strUserNum & ";')>0"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      RsTemp.MoveFirst
'      Do While Not RsTemp.EOF
'         For i = 0 To Combo1.ListCount - 1
'            If InStr(Combo1.List(i), RsTemp(0)) = 1 Then
'               Exit For
'            End If
'         Next i
'         If i = Combo1.ListCount Then
'            Combo1.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
'         End If
'         RsTemp.MoveNext
'      Loop
'   End If
'   '2014/8/8 END
'   Combo1.Text = Combo1.List(0)
   
'cancel by sonia 2024/9/27
'   'Add By Sindy 2023/5/16
'   If InStr(Pub_GetSpecMan("P1004管理人員"), strUserNum) > 0 Then
'      Combo1 = "P1004 " & GetPrjSalesNM("P1004")
'   End If
'   '2023/5/16 END
'end 2024/9/27
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Dim strText As String
'
'   '一進入系統,檢查是否有須要開啟此作業
'   If pub_CallNextABSForm = True Then
'      strText = ChkIsAbsenceMustPro
'      Me.Hide
'      If InStr(1, strText, "C") > 0 Then
''         frm160201.intChoose = 1
''         frm160201.Hide
''         Call frm160201.cmdOK_Click(0)
'         frm180203_1.Show
'      Else
'         pub_CallNextABSForm = False
'      End If
'   End If
   
   Set frm090202_3 = Nothing
'   If pub_CallNextABSForm = False Then
'      Call Forms(0).SysStartCallForm 'Add By Sindy 2011/10/7
'   End If
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   Dim Index As Integer
   
   'Modify By Sindy 2016/3/3 +不顯示,EEP02
   'Modify By Sindy 2018/4/16 e1.eep06 a,e1.eep07 b
   'Modify By Sindy 2020/12/1 + eep15,eep11
   arrGridHeadText = Array("V", "流程日期", "本所案號", "案件名稱", "國家", _
                           "類別/種類", "案件性質", "本所期限", "承辦人", "承辦期限", "目前流程狀態", _
                           "總收文號", "不顯示", "EEP02", "e1.eep06", "e1.eep07", "會稿方式", "eep15", "eep11")
   'Modify By Sindy 2018/8/29
   If SSTab1.Tab = 1 Then '已會稿
      arrGridHeadWidth = Array(200, 800, 1400, 1100, 800, _
                               500, 1200, 800, 800, 800, 0, _
                               0, 0, 0, 0, 0, 1000, 0, 0)
   Else
   '2018/8/29 END
      arrGridHeadWidth = Array(200, 800, 1400, 1100, 800, _
                               500, 1200, 800, 800, 800, 600, _
                               0, 600, 0, 0, 0, 0, 0, 0)
   End If
   Index = SSTab1.Tab
   GRD1(Index).Visible = False
   GRD1(Index).Cols = UBound(arrGridHeadText) + 1
   GRD1(Index).Rows = 2
   For iRow = 0 To GRD1(Index).Cols - 1
      GRD1(Index).row = 0
      GRD1(Index).col = iRow
      GRD1(Index).Text = arrGridHeadText(iRow)
      GRD1(Index).ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1(Index).CellAlignment = flexAlignCenterCenter
   Next
   GRD1(Index).Visible = True
End Sub

'Add By Sindy 2020/12/1
Private Sub SetColColor(intRow As Integer, Index As Integer)
   If intRow < 1 Then Exit Sub
   GRD1(Index).row = intRow
   '多案時,案件名稱變亮粉色
   If GRD1(Index).TextMatrix(intRow, 17) <> "" And _
      InStr(GRD1(Index).TextMatrix(intRow, 18), "多案單筆歷程") > 0 Then
      GRD1(Index).col = 3
      GRD1(Index).CellBackColor = &HFF00FF 'QBColor(Rnd * 5)
   End If
End Sub

'Add By Sindy 2016/3/3 增加不顯示功能
Private Sub Grd1_Click(Index As Integer)
Dim intRow As Integer, intCol As Integer
   
   intRow = GRD1(Index).MouseRow
   intCol = GRD1(Index).MouseCol
   If intRow <> 0 Then
      If intCol = 12 Then '不顯示
         If GRD1(Index).TextMatrix(intRow, 11) <> "" And GRD1(Index).TextMatrix(intRow, 10) = "聯絡" Then
            GRD1(Index).TextMatrix(intRow, 12) = "V"
            If MsgBox("請再次確定不顯示 " & vbCrLf & GRD1(Index).TextMatrix(intRow, 2) & " " & GRD1(Index).TextMatrix(intRow, 10) & " 嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
               GRD1(Index).TextMatrix(intRow, 12) = ""
            Else
               strExc(0) = "update EmpElectronProcess set eep13=null" & _
                           " where eep01='" & GRD1(Index).TextMatrix(intRow, 11) & "'" & _
                             " and eep02=" & GRD1(Index).TextMatrix(intRow, 13)
               Pub_SeekTbLog strExc(0) 'Add By Sindy 2020/11/30
               cnnConnection.Execute strExc(0)
               GRD1(Index).RowHeight(intRow) = 0
            End If
         End If
      End If
   End If
End Sub

Private Sub GRD1_DblClick(Index As Integer)
   'Modify By Sindy 2016/3/3
   If GRD1(Index).MouseRow <> 0 Then
      If GRD1(Index).MouseCol <> 12 Then
   '2016/3/3 END
         cmdDetail_Click
      End If
   End If
End Sub

Private Sub grd1_SelChange(Index As Integer)
Dim j As Integer 'Add By Sindy 2016/3/4

GRD1(Index).Visible = False
'Add By Sindy 2016/3/4
If GRD1(Index).MouseRow = 0 Then
   '已選取的資料列清除反白
   For j = 1 To GRD1(Index).Rows - 1
      If GRD1(Index).TextMatrix(j, 0) = "V" Then
         GRD1(Index).col = 0
         GRD1(Index).row = j
         GRD1(Index).Text = ""
         For i = 0 To GRD1(Index).Cols - 1
            'Modify By Sindy 2020/12/1
            If i <> 3 Then
            '2020/12/1 END
               GRD1(Index).col = i
               GRD1(Index).CellBackColor = QBColor(15)
            End If
         Next i
         Exit For
      End If
   Next j
Else
'2016/3/4 END
   '上一筆資料列清除反白
   'Modify By Sindy 2016/5/9
   'If dblPrevRow > 0 Then
   If dblPrevRow > 0 And dblPrevRow <= (GRD1(Index).Rows - 1) Then
   '2016/5/9 END
      GRD1(Index).col = 0
      GRD1(Index).row = dblPrevRow
      GRD1(Index).Text = ""
      For i = 0 To GRD1(Index).Cols - 1
         'Modify By Sindy 2020/12/1
         If i <> 3 Then
         '2020/12/1 END
            GRD1(Index).col = i
            GRD1(Index).CellBackColor = QBColor(15)
         End If
      Next i
   End If
   '目前資料列反白
   GRD1(Index).col = 0
   GRD1(Index).row = GRD1(Index).MouseRow
   dblPrevRow = GRD1(Index).row
'   If Grd1(index).Text = "V" Then
'      Grd1(index).Text = ""
'      For i = 0 To Grd1(index).Cols - 1
'         Grd1(index).col = i
'         Grd1(index).CellBackColor = QBColor(15)
'      Next i
'   Else
      If GRD1(Index).TextMatrix(GRD1(Index).row, 1) <> "" Then
         GRD1(Index).Text = "V"
         For i = 0 To GRD1(Index).Cols - 1
            'Modify By Sindy 2020/12/1
            If i <> 3 Then
            '2020/12/1 END
               GRD1(Index).col = i
               GRD1(Index).CellBackColor = &HFFC0C0
            End If
         Next i
      End If
'   End If
End If
GRD1(Index).Visible = True
End Sub

Private Sub grd1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow GRD1(Index), x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   'Grd1(index).col = nCol
   GRD1(Index).row = nRow
   If Me.GRD1(Index).row < 1 And Me.GRD1(Index).Text <> "V" Then
      If Me.GRD1(Index).Text = "目次" Then
         If m_blnColOrderAsc = True Then
            Me.GRD1(Index).Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1(Index).Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.GRD1(Index).Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1(Index).Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

'Add By Sindy 2018/8/29
Private Sub SSTab1_Click(PreviousTab As Integer)
   If Me.Visible = True Then Call QueryData(SSTab1.Tab)
End Sub
