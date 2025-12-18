VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090202_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "待核判區"
   ClientHeight    =   5750
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   9930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   9930
   Begin VB.ComboBox Combo3 
      Height          =   300
      ItemData        =   "frm090202_1.frx":0000
      Left            =   5700
      List            =   "frm090202_1.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   780
      Width           =   3180
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frm090202_1.frx":0004
      Left            =   1020
      List            =   "frm090202_1.frx":0014
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   450
      Width           =   1350
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "畫面更新(&Q)"
      Height          =   360
      Left            =   5520
      TabIndex        =   1
      Top             =   60
      Width           =   1155
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "明細資料(&D)"
      Height          =   360
      Left            =   6720
      TabIndex        =   2
      Top             =   60
      Width           =   1305
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   8055
      TabIndex        =   3
      Top             =   60
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Height          =   4575
      Left            =   60
      TabIndex        =   5
      Top             =   1110
      Width           =   9795
      _ExtentX        =   17268
      _ExtentY        =   8079
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V|目次|流程日期|本所案號|案件名稱|國家|種類|案件性質|本所期限|承辦人|承辦期限|智權人員|目前流程狀態|不顯示"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   14
   End
   Begin VB.Label Label2 
      Caption         =   "註：在”不顯示”欄位上點一下(V)，可以取消顯示聯絡歷程。"
      ForeColor       =   &H000000C0&
      Height          =   200
      Left            =   4800
      TabIndex        =   11
      Top             =   570
      Width           =   4970
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1020
      TabIndex        =   0
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
   Begin VB.Label Label5 
      Caption         =   "顏色說明："
      Height          =   225
      Left            =   4800
      TabIndex        =   10
      Top             =   825
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "最近聯絡："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   1
      Left            =   90
      TabIndex        =   8
      Top             =   510
      Width           =   900
   End
   Begin VB.Label Label16 
      Caption         =   "註：雙擊選取時，開啟承辦歷程"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   150
      TabIndex        =   7
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "核判人員："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   6
      Top             =   180
      Width           =   900
   End
End
Attribute VB_Name = "frm090202_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/14 Form2.0已修改
'Create by Sindy 2013/4/30
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim i As Integer
Dim dblPrevRow As Double
Public m_ProSysState As String '設定承辦人或繪圖人員 1：承辦人 2：繪圖人員 3：中打室
Dim m_EMPCon As String 'Add By Sindy 2017/1/18
'Add By Sindy 2024/1/4
Const 待核判區使用到的歷程狀態 As String = "'" & EMP_送英核 & "','" & EMP_送核 & "','" & EMP_送判 & "'," & _
                 "'" & EMP_翻譯交稿 & "','" & EMP_排版完成 & "','" & EMP_送核稿分案 & "','" & EMP_程序送判 & "'," & _
                 "'" & EMP_墨完 & "','" & EMP_草核 & "'," & _
                 "'" & EMP_送排版 & "','" & EMP_送轉檔 & "'"
'2024/1/4 END


'明細資料
Private Sub cmdDetail_Click()
Dim ii As Integer
Dim nFrm As Form
   
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
'         'Add By Sindy 2017/9/19
'         '檢查表單是否已開啟，若是，則關閉
'         For Each nFrm In Forms
'            If StrComp(nFrm.Name, "frm090202_2", vbTextCompare) = 0 Then
'               Unload frm090202_2
'               'Add By Sindy 2020/1/17 有資料要儲存,尚需處理...
'               If strSaveConfirm = True Then
'                  frm090202_2.ZOrder
'                  Exit Sub
'               Else
'               '2020/1/17 END
'                  Exit For
'               End If
'            End If
'         Next
'         '2017/9/19 END
         If PUB_ChkFormIsClose("frm090202_2") = False Then Exit Sub 'Add By Sindy 2020/1/17
         frm090202_2.Hide
         frm090202_2.m_EEP01 = GRD1.TextMatrix(i, 13) '總收文號
         
         'Add By Sindy 2017/1/18 若點選的是聯終歷程,必須檢查是否有待回覆歷程要操作 ex:P-116465 聯絡/送會
         If GRD1.TextMatrix(i, 12) = "聯絡" Then
            strSql = "select eep02,eep05 from empelectronprocess" & _
                     " where (EEP04 in(" & m_EMPCon & ") and EEP09='Y')" & _
                     " and eep01='" & GRD1.TextMatrix(i, 13) & "'" & _
                     " order by eep06 asc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               RsTemp.MoveFirst
               Do While Not RsTemp.EOF
                  For ii = 0 To Combo1.ListCount - 1
                     If Left(Trim(Combo1.List(ii)), 5) = RsTemp.Fields("eep05") Then
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
         frm090202_2.m_CurrFlowEEP02 = GRD1.TextMatrix(i, 15) 'Add By Sindy 2016/3/25 目前要處理的歷程序號
carry_on: 'Add By Sindy 2017/1/18
         frm090202_2.intReceiveKind = 1
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

Public Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strConSql As String
Dim strVal As String
Dim strQuyDate As String
Dim ii As Integer
   
   m_blnColOrderAsc = True
   QueryData = True
   
   'Add By Sindy 2013/9/17
   If Combo2.ListIndex = 0 Then
      strQuyDate = CompWorkDay(3, strSrvDate(1), 1) '不含當天,3個工作天
   ElseIf Combo2.ListIndex = 1 Then
      strQuyDate = CompWorkDay(5, strSrvDate(1), 1) '不含當天,5個工作天
   ElseIf Combo2.ListIndex = 2 Then
      strQuyDate = CompWorkDay(7, strSrvDate(1), 1) '不含當天,7個工作天
   Else
      '全部
   End If
   '2013/9/17 END
   
   GRD1.Clear
   SetGrd
   
   '2023/11/2 END
   If m_ProSysState = "1" Then '承辦人
      'Modify By Sindy 2023/9/28 +EMP_翻譯交稿
      ''" & EMP_轉回 & "'
      m_EMPCon = "'" & EMP_送英核 & "','" & EMP_送核 & "','" & EMP_送判 & "'," & _
                 "'" & EMP_翻譯交稿 & "','" & EMP_排版完成 & "','" & EMP_送核稿分案 & "','" & EMP_程序送判 & "'"
      'Modify By Sindy 2023/4/25 Mark;發生承辦人離職改為主管的案子,但案況在送核中 P-127816
      'strConSql = " And cp14<>'" & Trim(Left("" & Combo1.Text, 6)) & "'"
      'Modify By Sindy 2023/4/27 排除自己承辦的案件聯絡
      'Modify By Sindy 2024/8/20 cp14改抓ep05
      strConSql = " And not(ep05='" & Trim(Left("" & Combo1.Text, 6)) & "' and ep05 is not null and e1.EEP04='" & EMP_聯絡 & "')"
      
   ElseIf m_ProSysState = "2" Then '繪圖人員
      'Modify By Sindy 2015/4/22 +,'" & EMP_草核 & "'"
      m_EMPCon = "'" & EMP_墨完 & "','" & EMP_草核 & "'"
      'Modify By Sindy 2023/4/25 Mark
      'strConSql = " And ep13<>'" & Trim(Left("" & Combo1.Text, 6)) & "'"
      'Modify By Sindy 2023/4/27 排除自己承辦的案件聯絡
      strConSql = " And not(ep13='" & Trim(Left("" & Combo1.Text, 6)) & "' and ep13 is not null and e1.EEP04='" & EMP_聯絡 & "')"
      
   'Add By Sindy 2023/10/2
   ElseIf m_ProSysState = "3" Then '中打室
      m_EMPCon = "'" & EMP_送排版 & "','" & EMP_送轉檔 & "'"
      strConSql = ""
      '2023/10/2 END
   End If
   
   Screen.MousePointer = vbHourglass
'   strSql = "Select ' ' as V,EP01 as 目次,SqlDateT(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,PA05||PA06||PA07 as 案件名稱," & _
'            "NA03 as 國家,Decode(PA09,'000',PTM03,PTM04) as 種類,Decode(PA09,'000',CPM03,CPM04) as 案件性質," & _
'            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員, ac03 as 目前流程狀態," & _
'            "EEP01 as 總收文號,EEP02 as 序號" & _
'            " From EmpElectronProcess,CaseProgress,EngineerProgress,Patent," & _
'            "staff s1,staff s2,nation,CasePropertyMap,PatentTradeMarkMap,allcode" & _
'            " Where EEP09='Y' and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
'            " And EEP01=CP09(+)" & _
'            " And EEP01=EP02(+)" & _
'            " And CP01=PA01(+) And CP02=PA02(+) And CP03=PA03(+) And CP04=PA04(+)" & _
'            " And CP14=s1.ST01(+) And CP13=s2.ST01(+)" & _
'            " And PA09=NA01(+)" & _
'            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
'            " And '1'=PTM01(+) AND PA08=PTM02(+)" & _
'            " And ac01='09' And EEP04=ac02(+)" & _
'            " And EEP04 in(" & strCon & ")" & _
'            " order by EP01 desc"
   'Add By Sindy 2013/9/17 +IIf(strQuyDate <> "", " And e1.EEP06 >=" & strQuyDate, "")
   'Modify By Sindy 2014/1/9 +聯絡送英核尚待核稿
   'Modify By Sindy 2015/3/16 +送日核
   'Modify By Sindy 2016/3/3 調整聯絡的抓法
'            " Union" & _
'            " select EEP01,EEP02,EEP04 from EmpElectronProcess e1" & _
'            " where e1.EEP02 in (select max(eep02) from EmpElectronProcess where eep01=e1.eep01)" & _
'            " and e1.EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
'            " And e1.EEP04 in('" & EMP_聯絡 & "')" & _
'            IIf(strQuyDate <> "", " And e1.EEP06>=" & strQuyDate, "")
   'Modify By Sindy 2018/4/16 修改SQL
'   strVal = "(select EEP01,max(EEP02) as EEP02 from(" & _
'            "select EEP01,EEP02,EEP04 from EmpElectronProcess" & _
'            " where EEP09='Y'" & _
'            " and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
'            " And EEP04 in(" & m_EMPCon & ")" & _
'            " Union" & _
'            " select EEP01,EEP02,EEP04 from EmpElectronProcess e1" & _
'            " where e1.EEP13='Y'" & _
'            " and e1.EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
'            " And e1.EEP04 in('" & EMP_聯絡 & "')" & _
'            IIf(strQuyDate <> "", " And e1.EEP06>=" & strQuyDate, "") & _
'            " Union" & _
'            " select EEP01,EEP02,EEP04 from EmpElectronProcess e2,engineerprogress" & _
'            " where e2.EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
'            " And e2.EEP04='" & EMP_聯絡 & "'" & _
'            " and e2.EEP01=EP02(+)" & _
'            " And (e2.eep08 like '%[送英核]%' or e2.eep08 like '%[送日核]%') and nvl(EP33,0)=0" & _
'            ") group by EEP01) V1"
   strVal = "select EEP01,max(EEP02) as EEP02 from (" & _
            "select EEP01,EEP02,EEP04 from EmpElectronProcess" & _
            " where EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
            " And (EEP04 in(" & m_EMPCon & ") and EEP09='Y')" & _
            " Union" & _
            " select EEP01,EEP02,EEP04 from EmpElectronProcess" & _
            " where EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
            " And (EEP04 in('" & EMP_聯絡 & "') and EEP13='Y')" & _
            IIf(strQuyDate <> "", " And EEP06>=" & strQuyDate, "") & _
            " Union" & _
            " select EEP01,EEP02,EEP04 from EmpElectronProcess,engineerprogress" & _
            " where EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
            " And EEP04='" & EMP_聯絡 & "'" & _
            " and EEP01=EP02(+)" & _
            " And (eep08 like '%[送英核]%' or eep08 like '%[送日核]%') and nvl(EP33,0)=0" & _
            ") group by EEP01"
   'Modify By Sindy 2016/3/3 +不顯示,e2.EEP02
   'Modify By Sindy 2016/9/2 And cp27 is null And cp57 is null -> and cp158=0 and cp159=0
   'Modify By Sindy 2018/4/16 修改SQL + 商標檔和服務業務檔
'   strSql = "Select ' ' as V,EP01 as 目次,SqlDateT(e2.EEP06)||' '||sqltime(e2.EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,PA05||PA06||PA07 as 案件名稱," & _
'            "NA03 as 國家,Decode(PA09,'000',PTM03,PTM04) as 種類,Decode(PA09,'000',CPM03,CPM04) as 案件性質," & _
'            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員, ac03 as 目前流程狀態," & _
'            "e2.EEP01 as 總收文號,' ' as 不顯示,e2.EEP02" & _
'            " From EmpElectronProcess e2," & strVal & ",CaseProgress,EngineerProgress,Patent," & _
'            "staff s1,staff s2,nation,CasePropertyMap,PatentTradeMarkMap,allcode" & _
'            " Where V1.EEP01=e2.EEP01 AND V1.EEP02=e2.EEP02 AND e2.EEP01=CP09(+)" & _
'            " And e2.EEP01=EP02(+)" & _
'            " And CP01=PA01(+) And CP02=PA02(+) And CP03=PA03(+) And CP04=PA04(+)" & _
'            " And CP14=s1.ST01(+) And CP13=s2.ST01(+)" & _
'            " And PA09=NA01(+)" & _
'            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
'            " And '1'=PTM01(+) AND PA08=PTM02(+)" & _
'            " And ac01='09' And e2.EEP04=ac02(+)" & _
'            " and cp158=0 and cp159=0" & strConSql & _
'            " order by e2.EEP06 desc,e2.EEP07 desc"
'            '" order by EP01 desc"
   'Modify By Sindy 2019/10/17 送核或送判在16:00以後則以下一工作日起算: ,decode(e1.EEP09,'Y',decode(sign(e1.EEP07-160000),1,to_char(to_date(e1.EEP06,'YYYYMMDD')+1,'YYYYMMDD'),e1.EEP06),'') AS ChkDate,cpm02
   'Modify By Sindy 2020/11/30 + EEP15
   'Modify By Sindy 2024/8/20 cp14改抓ep05
   'Modify By Sindy 2024/12/18 淑華提的需求+指定送件日
   strSql = "Select ' ' as V,EP01 as 目次,SqlDateT(e1.EEP06)||' '||sqltime(e1.EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,PA05||PA06||PA07 as 案件名稱," & _
            "NA03 as 國家,Decode(PA09,'000',PTM03,PTM04) as 種類,Decode(PA09,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT(cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員, ac03 as 目前流程狀態," & _
            "e1.EEP01 as 總收文號,' ' as 不顯示,e1.EEP02,e1.EEP06 a,e1.EEP07 b,'' as ChkDate,cpm02,EEP15,SqlDateT(cp142) as 指定送件日" & _
            " From EmpElectronProcess e1,CaseProgress,EngineerProgress,Patent,staff s1,staff s2,nation,CasePropertyMap,PatentTradeMarkMap,allcode" & _
            " Where (e1.eep01,e1.eep02) in(" & strVal & ")" & _
            " AND e1.EEP01=CP09(+)" & _
            " And e1.EEP01=EP02(+)" & _
            " And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04" & _
            " And ep05=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And PA09=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And '1'=PTM01(+) AND PA08=PTM02(+)" & _
            " And ac01='09' And e1.EEP04=ac02(+)" & _
            " and cp158=0 and cp159=0" & strConSql
   strSql = strSql & " union " & _
            "Select ' ' as V,EP01 as 目次,SqlDateT(e1.EEP06)||' '||sqltime(e1.EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,TM05||TM06||TM07 as 案件名稱," & _
            "NA03 as 國家,Decode(TM10,'000',PTM03,PTM04) as 種類,Decode(TM10,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT(cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員, ac03 as 目前流程狀態," & _
            "e1.EEP01 as 總收文號,' ' as 不顯示,e1.EEP02,e1.EEP06 a,e1.EEP07 b,decode(e1.EEP09,'Y',decode(sign(e1.EEP07-160000),1,to_char(to_date(e1.EEP06,'YYYYMMDD')+1,'YYYYMMDD'),e1.EEP06),'') AS ChkDate,cpm02,EEP15,SqlDateT(cp142) as 指定送件日" & _
            " From EmpElectronProcess e1,CaseProgress,EngineerProgress,Trademark,staff s1,staff s2,nation,CasePropertyMap,PatentTradeMarkMap,allcode" & _
            " Where (e1.eep01,e1.eep02) in(" & strVal & ")" & _
            " AND e1.EEP01=CP09(+)" & _
            " And e1.EEP01=EP02(+)" & _
            " And CP01=TM01 And CP02=TM02 And CP03=TM03 And CP04=TM04" & _
            " And ep05=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And TM10=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And '2'=PTM01(+) AND TM08=PTM02(+)" & _
            " And ac01='09' And e1.EEP04=ac02(+)" & _
            " and cp158=0 and cp159=0" & strConSql
   strSql = strSql & " union " & _
            "Select ' ' as V,EP01 as 目次,SqlDateT(e1.EEP06)||' '||sqltime(e1.EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,SP05||SP06||SP07 as 案件名稱," & _
            "NA03 as 國家,'' as 種類,Decode(SP09,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT(cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員, ac03 as 目前流程狀態," & _
            "e1.EEP01 as 總收文號,' ' as 不顯示,e1.EEP02,e1.EEP06 a,e1.EEP07 b,decode(e1.EEP09,'Y',decode(sign(e1.EEP07-160000),1,to_char(to_date(e1.EEP06,'YYYYMMDD')+1,'YYYYMMDD'),e1.EEP06),'') AS ChkDate,cpm02,EEP15,SqlDateT(cp142) as 指定送件日" & _
            " From EmpElectronProcess e1,CaseProgress,EngineerProgress,servicepractice,staff s1,staff s2,nation,CasePropertyMap,allcode" & _
            " Where (e1.eep01,e1.eep02) in(" & strVal & ")" & _
            " AND e1.EEP01=CP09(+)" & _
            " And e1.EEP01=EP02(+)" & _
            " And CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04" & _
            " And ep05=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And SP09=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And ac01='09' And e1.EEP04=ac02(+)" & _
            " and cp158=0 and cp159=0" & strConSql
   'Add By Sindy 2021/7/13
   strSql = strSql & " union " & _
            "Select ' ' as V,EP01 as 目次,SqlDateT(e1.EEP06)||' '||sqltime(e1.EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,LC05||LC06||LC07 as 案件名稱," & _
            "NA03 as 國家,'' as 種類,Decode(LC15,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT(cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員, ac03 as 目前流程狀態," & _
            "e1.EEP01 as 總收文號,' ' as 不顯示,e1.EEP02,e1.EEP06 a,e1.EEP07 b,decode(e1.EEP09,'Y',decode(sign(e1.EEP07-160000),1,to_char(to_date(e1.EEP06,'YYYYMMDD')+1,'YYYYMMDD'),e1.EEP06),'') AS ChkDate,cpm02,EEP15,SqlDateT(cp142) as 指定送件日" & _
            " From EmpElectronProcess e1,CaseProgress,EngineerProgress,Lawcase,staff s1,staff s2,nation,CasePropertyMap,allcode" & _
            " Where (e1.eep01,e1.eep02) in(" & strVal & ")" & _
            " AND e1.EEP01=CP09(+)" & _
            " And e1.EEP01=EP02(+)" & _
            " And CP01=LC01 And CP02=LC02 And CP03=LC03 And CP04=LC04" & _
            " And ep05=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And LC15=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And ac01='09' And e1.EEP04=ac02(+)" & _
            " and cp158=0 and cp159=0" & strConSql
   strSql = strSql & " union " & _
            "Select ' ' as V,EP01 as 目次,SqlDateT(e1.EEP06)||' '||sqltime(e1.EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,HC06 as 案件名稱," & _
            "NA03 as 國家,'' as 種類,Decode('000','000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT(cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員, ac03 as 目前流程狀態," & _
            "e1.EEP01 as 總收文號,' ' as 不顯示,e1.EEP02,e1.EEP06 a,e1.EEP07 b,decode(e1.EEP09,'Y',decode(sign(e1.EEP07-160000),1,to_char(to_date(e1.EEP06,'YYYYMMDD')+1,'YYYYMMDD'),e1.EEP06),'') AS ChkDate,cpm02,EEP15,SqlDateT(cp142) as 指定送件日" & _
            " From EmpElectronProcess e1,CaseProgress,EngineerProgress,Hirecase,staff s1,staff s2,nation,CasePropertyMap,allcode" & _
            " Where (e1.eep01,e1.eep02) in(" & strVal & ")" & _
            " AND e1.EEP01=CP09(+)" & _
            " And e1.EEP01=EP02(+)" & _
            " And CP01=HC01 And CP02=HC02 And CP03=HC03 And CP04=HC04" & _
            " And ep05=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And '000'=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And ac01='09' And e1.EEP04=ac02(+)" & _
            " and cp158=0 and cp159=0" & strConSql
   '2021/7/13 END
   strSql = strSql & " order by a desc,b desc"
   '不可控制CP13,因有可能核稿人亦也是智權人員 " And cp13<>'" & Trim(Left("" & Combo1.Text, 6)) & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
      'Add By Sindy 2016/3/4
      For ii = 1 To GRD1.Rows - 1
         Call SetColColor(ii)
      Next ii
      '2016/3/4 END
   Else
      QueryData = False
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
      'Modify By Sindy 2019/12/25
      'ShowNoData
      If Me.Visible = True Then ShowNoData
      '2019/12/25 END
      Exit Function
   End If
   
   '若有資料游標停在第一筆
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
   'Modify By Sindy 2016/4/15 因資料列會顯示其他顏色做提醒,所以不要預設查詢後停在第一筆資料列,不然顏色會顯示不出來
'   dblPrevRow = GRD1.row
'   If rsTmp.RecordCount > 0 Then
'      GRD1.Text = "V"
'      For i = 0 To GRD1.Cols - 1
'         GRD1.col = i
'         GRD1.CellBackColor = &HFFC0C0
'      Next i
'   End If
   GRD1.Visible = True
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Function

'畫面更新
Private Sub cmdQuery_Click()
   'If QueryData = False Then ShowNoData
   'Call Combo1_Click
   Call QueryData
End Sub

'Add By Sindy 2023/1/17
Private Sub Combo1_Change()
   Call QueryData
End Sub

Private Sub Combo2_Click()
   Call QueryData 'Add By Sindy 2023/4/12
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
'      If Combo1.Text <> "" Then strST01 = Left(Combo1.Text, 5)
'      Combo1.Clear
'      Combo1.AddItem strUserNum & " " & strUserName
'      '檢查當時是否需要為他人職代
'      Call Pub_SetForOthersEmpCombo(strUserNum, Combo1, False)
'      For i = 0 To Combo1.ListCount - 1
'         If Left(Combo1.List(i), 5) = strST01 Then
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
   Me.Tag = ""
   
   Combo3.Clear
   'Modify By Sindy 2024/12/18 淑華提的需求+指定送件日
   'Combo3.AddItem "淺紅色：逾(含當日)本所期限"
   Combo3.AddItem "淺紅色：逾(含當日)本所期限或指定送件日"
   '2024/12/18 END
   If Left(Pub_StrUserSt03, 2) = "P2" Then '商標處人員才需要
      Combo3.AddItem "黃色：逾(含當日)核判期限"
      If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
         Combo3.AddItem "紫紅色：案件名稱變色，多案歷程"
      End If
   End If
   
   'Combo1.Clear
   'Call Combo1_Click
   Call SetCombo1
   Combo2.Text = Combo2.List(3) 'Add By Sindy 2013/9/17
   Call QueryData
   
   If m_ProSysState = "3" Then Me.Caption = "待排版區"
End Sub

Private Sub SetCombo1()
Dim strCon As String
Dim ii As Integer
Dim strEmp As String, varTmp As Variant 'Add By Sindy 2019/3/11
   
   Combo1.Clear
   Combo1.AddItem strUserNum & " " & strUserName
   '檢查當時是否需要為他人職代
   'Modify By Sindy 2019/3/11 + strEmp
   Call Pub_SetForOthersEmpCombo(strUserNum, Combo1, False, strEmp)
   
   'Add By Sindy 2024/6/21
   If Mid(Pub_StrUserSt03, 1, 2) <> "F2" Then
   '2024/6/21 END
      'ADD BY SONIA 2018/8/22 核判人為外翻人員時,由其內部郵件收件員工編號st14人員操作 P-119644
      strSql = "select st01,st02 from staff where st03 like 'F5%' and instr(st14,'" & strUserNum & "')>0"
      'Modify By Sindy 2019/3/11 增加幫他人代理工作時
      varTmp = Split(strEmp, ";")
      For intI = 0 To UBound(varTmp) - 1
         If varTmp(intI) <> "" Then
            strSql = strSql & " union select st01,st02 from staff where st03 like 'F5%' and instr(st14,'" & varTmp(intI) & "')>0"
         End If
      Next intI
      strSql = strSql & " order by 1"
      '2019/3/11 END
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         With RsTemp
            .MoveFirst
            Do While Not .EOF
               If Not IsNull(RsTemp.Fields(0)) Then
                  For ii = 0 To Combo1.ListCount - 1
                     If InStr(Combo1.List(ii), RsTemp(0)) = 1 Then
                        Exit For
                     End If
                  Next
                  If ii = Combo1.ListCount Then
                     Combo1.AddItem RsTemp.Fields(0) & " " & RsTemp.Fields(1)
                  End If
               End If
               .MoveNext
            Loop
         End With
      End If
      'END 2018/8/22
   End If
   
   'Add By Sindy 2015/5/20 增加特殊人員可以看到各區核稿主管的待核資料
   'modify by sonia 2023/4/11 北所也改為二人且可以看到全所核稿主管的待核資料
   'If strUserNum = Pub_GetSpecMan("A2") Then '專利處台北區主管
   '   strCon = " and st06='1'"
   If InStr(Replace(Pub_GetSpecMan("A2", False), ";", ","), strUserNum) > 0 Then '專利處台北區主管
      strCon = " " '要一個空白格,下面if判斷用
   'end 2023/4/11
   'modify by sonia 2023/4/10 中所加入林柄佑之專利國內部虛帳號82099
   'ElseIf strUserNum = Pub_GetSpecMan("A3") Then '專利處台中區主管
   ElseIf InStr(Replace(Pub_GetSpecMan("A3", False), ";", ","), strUserNum) > 0 Then '專利處台中區主管
      strCon = " and st06='2'"
   ElseIf strUserNum = Pub_GetSpecMan("A4") Then '專利處台南區主管(南高所核判主管)
      strCon = " and st06 in('3','4')"
   'Modify By Sindy 2024/1/4
   '外專:
   '直屬下屬有待核判資料的人員，會列在核判人員欄的下拉選單內，主管可隨時選擇人員的資料去做代為核判的動作。
   ElseIf Mid(Pub_StrUserSt03, 1, 2) = "F2" Then
      strCon = ""
      'Pub_GetSpecMan("S") = strUserNum:日本部大主管,要看全日本部的二級主管
      'Modify By Sindy 2025/3/12 簡經理提調整系統中『待核判區』裡，林軒吉副理及沈彥伶副理的系統權限，
      '                          讓他們可以看掛在底下主任身上的待核判案件並代為核判。
      strSql = "select st01,st02 from staff" & _
               " where st01 in(select st52 from staff,acc090,acc090new" & _
                              " where st04='1' and st03 in('F21','F22') and st01<>'" & strUserNum & "'" & _
                                " and st93=a0921(+) and st03=a0901(+) and a0924='" & strUserNum & "'" & _
                              " group by st52)" & _
               " union " & _
               "select st01,st02 from staff" & _
               " where st01 in(select st52 from staff,acc090,acc090new" & _
                              " where st04='1' and st03 in('F21','F22') and st01<>'" & strUserNum & "'" & _
                                " and st93=a0921(+) and st03=a0901(+)" & _
                                " and st16='3' and (st52='" & strUserNum & "' or st53='" & strUserNum & "' or st54='" & strUserNum & "'))" & _
               " order by st01"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         With RsTemp
            .MoveFirst
            Do While Not .EOF
               If strCon <> "" Then strCon = strCon & ","
               strCon = strCon & "'" & RsTemp.Fields(0) & "'"
               .MoveNext
            Loop
         End With
      End If
      If strCon <> "" Then
         strCon = " and eep05 in(" & strCon & ")"
      End If
   '2024/1/4 END
   Else
      strCon = ""
   End If
   
   '71011.王副總能夠看到全部核判人員的待核判區資料
   'Modify By Sindy 2015/9/4 + 待判的資料也要看的到
   'Modify By Sindy 2022/4/27 + 可看全所人員資料者，原只有王副總，請再增加郭雅娟及李柏翰
   'modify by sonia 2023/4/19  2023/4/11游經理提出北所A2人員可看全所(已加入99050),4/19將71011及79075加入A2即可
   'If (strUserNum = "71011" Or strUserNum = "79075" Or strUserNum = "99050") Or strCon <> "" Then
   If strCon <> "" Then
      'Modify By Sindy 2023/5/22 +,'" & EMP_送英核 & "'
      ''" & EMP_送英核 & "','" & EMP_送核 & "','" & EMP_送判 & "','" & EMP_墨完 & "','" & EMP_草核 & "'
      strSql = "select distinct eep05,st02,st06,st93,1 sort from empelectronprocess,staff,caseprogress" & _
               " where eep04 in(" & 待核判區使用到的歷程狀態 & ") and eep09='Y' and eep05=st01(+) " & strCon & _
               " and substr(st03,1,2)='" & Left(Pub_StrUserSt03, 2) & "'" & _
               " and eep01=cp09(+) and cp158=0 and cp159=0"
      'Modify By Sindy 2025/3/18 專利處增加檢查有聯絡案件也要顯示出來
      If InStr(Replace(Pub_GetSpecMan("A2", False), ";", ","), strUserNum) > 0 Then '專利處台北區主管
         strSql = strSql & " union " & _
               "select distinct eep05,st02,st06,st93,2 sort from empelectronprocess,staff,caseprogress" & _
               " where eep04 ='" & EMP_聯絡 & "' and eep13='Y' and eep05=st01(+) " & strCon & _
               " and substr(st03,1,2)='" & Left(Pub_StrUserSt03, 2) & "'" & _
               " and eep01=cp09(+) and cp158=0 and cp159=0 and cp14<>eep05"
      End If
      'strSql = strSql & " order by eep05 asc"
      strSql = strSql & " order by st93,sort,eep05 asc"
      '2025/3/18 END
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         With RsTemp
            .MoveFirst
            Do While Not .EOF
               If Not IsNull(RsTemp.Fields(0)) Then
                  For ii = 0 To Combo1.ListCount - 1
                     If InStr(Combo1.List(ii), RsTemp(0)) = 1 Then
                        Exit For
                     End If
                  Next
                  If ii = Combo1.ListCount Then
                     Combo1.AddItem RsTemp.Fields(0) & " " & RsTemp.Fields(1)
                  End If
               End If
               .MoveNext
            Loop
         End With
      End If
   End If
   '2015/5/20 END
   
   'Add By Sindy 2022/7/14 黃教威:為因應秀娟經理業務出差或與客戶會議時而部門同仁臨時有核判ACS案件之需求，
   '                              經請示過王經理，請設定由我可代為操作判發，
   If Pub_StrUserSt03 = "W20" And strUserNum = "A9005" Then
      Combo1.AddItem "A5024 " & GetPrjSalesNM("A5024")
   End If
   '2022/7/14 END
   
   Combo1.Text = Combo1.List(0)
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
   
   Set frm090202_1 = Nothing
'   If pub_CallNextABSForm = False Then
'      Call Forms(0).SysStartCallForm 'Add By Sindy 2011/10/7
'   End If
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modify By Sindy 2016/3/3 +不顯示,EEP02
   'Modify By Sindy 2019/10/17 + , "ChkDate"
   'Modify By Sindy 2020/11/30 + EEP15
   'Modify By Sindy 2024/12/18 淑華提的需求+指定送件日
   arrGridHeadText = Array("V", "目次", "流程日期", "本所案號", "案件名稱", "國家", _
                           "種類", "案件性質", "本所期限", "承辦人", "承辦期限", _
                           "智權人員", "目前流程狀態", "總收文號", "不顯示", "EEP02", _
                           "e1.EEP06 a", "e1.EEP07 b", "ChkDate", "cpm02", "EEP15", _
                           "指定送件日")
   arrGridHeadWidth = Array(200, 400, 800, 1400, 1000, 700, _
                            450, 900, 800, 600, 800, _
                            600, 600, 0, 600, 0, _
                            0, 0, 0, 0, 0, _
                            800)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = 2
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      If iRow = 11 Or iRow = 12 Then
         GRD1.CellAlignment = flexAlignLeftCenter
      Else
         GRD1.CellAlignment = flexAlignCenterCenter
      End If
   Next
   GRD1.Visible = True
End Sub

'Add By Sindy 2016/3/3 增加不顯示功能
Private Sub Grd1_Click()
Dim intRow As Integer, intCol As Integer
Dim i As Integer, j As Integer 'Add By Sindy 2016/3/4
   
   intRow = GRD1.MouseRow
   intCol = GRD1.MouseCol
'   If intRow <> 0 Then
'      If intCol = 14 Then '不顯示
'         If GRD1.TextMatrix(intRow, 13) <> "" And GRD1.TextMatrix(intRow, 12) = "聯絡" Then
'            GRD1.TextMatrix(intRow, 14) = "V"
'            If MsgBox("請再次確定不顯示 " & vbCrLf & GRD1.TextMatrix(intRow, 3) & " " & GRD1.TextMatrix(intRow, 12) & " 嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'               GRD1.TextMatrix(intRow, 14) = ""
'            Else
'               strExc(0) = "update EmpElectronProcess set eep13=null" & _
'                           " where eep01='" & GRD1.TextMatrix(intRow, 13) & "'" & _
'                             " and eep02=" & GRD1.TextMatrix(intRow, 15)
'               cnnConnection.Execute strExc(0)
'               GRD1.RowHeight(intRow) = 0
'            End If
'         End If
'      End If
'   End If
   
   'Modify By Sindy 2018/10/3 下面程式從 grd1_SelChange 移過來==>一筆資料時,點勾勾有時會點不出來
   GRD1.Visible = False
   'Add By Sindy 2016/3/4
   If GRD1.MouseRow = 0 Then
      '已選取的資料列清除反白
      For j = 1 To GRD1.Rows - 1
         If GRD1.TextMatrix(j, 0) = "V" Then
            GRD1.col = 0
            GRD1.row = j
            GRD1.Text = ""
            For i = 0 To GRD1.Cols - 1
               'Modify By Sindy 2020/11/30
               If i <> 4 Then
               '2020/11/30 END
                  GRD1.col = i
                  GRD1.CellBackColor = QBColor(15)
               End If
            Next i
            Call SetColColor(j)
            Exit For
         End If
      Next j
   Else
   '2016/3/4 END
      '上一筆資料列清除反白
      'Modify By Sindy 2016/5/9
      'If dblPrevRow > 0 Then
      If dblPrevRow > 0 And dblPrevRow <= (GRD1.Rows - 1) Then
      '2016/5/9 END
         GRD1.col = 0
         GRD1.row = dblPrevRow
         GRD1.Text = ""
         For i = 0 To GRD1.Cols - 1
            'Modify By Sindy 2020/11/30
            If i <> 4 Then
            '2020/11/30 END
               GRD1.col = i
               GRD1.CellBackColor = QBColor(15)
            End If
         Next i
         Call SetColColor(CInt(dblPrevRow))
      End If
      '目前資料列反白
      GRD1.col = 0
      GRD1.row = GRD1.MouseRow
      dblPrevRow = GRD1.row
   '   If GRD1.Text = "V" Then
   '      GRD1.Text = ""
   '      For i = 0 To GRD1.Cols - 1
   '         GRD1.col = i
   '         GRD1.CellBackColor = QBColor(15)
   '      Next i
   '   Else
         'Modify By Sindy 2024/1/4
         'If GRD1.TextMatrix(GRD1.row, 1) <> "" Then '目次
         If GRD1.TextMatrix(GRD1.row, 3) <> "" Then '案號
         '2024/1/4 END
            GRD1.Text = "V"
            For i = 0 To GRD1.Cols - 1
               'Modify By Sindy 2020/11/30
               If i <> 4 Then
               '2020/11/30 END
                  GRD1.col = i
                  GRD1.CellBackColor = &HFFC0C0
               End If
            Next i
         End If
   '   End If
      'Modify By Sindy 2018/10/3 由上面程式改放至此區檢查
      If intRow <> 0 Then
         If intCol = 14 Then '不顯示
            If GRD1.TextMatrix(intRow, 13) <> "" And GRD1.TextMatrix(intRow, 12) = "聯絡" Then
               GRD1.Visible = True
               GRD1.TextMatrix(intRow, 14) = "V"
               If MsgBox("請再次確定不顯示 " & vbCrLf & GRD1.TextMatrix(intRow, 3) & " " & GRD1.TextMatrix(intRow, 12) & " 嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                  GRD1.TextMatrix(intRow, 14) = ""
               Else
                  strExc(0) = "update EmpElectronProcess set eep13=null" & _
                              " where eep01='" & GRD1.TextMatrix(intRow, 13) & "'" & _
                                " and eep02=" & GRD1.TextMatrix(intRow, 15)
                  Pub_SeekTbLog strExc(0) 'Add By Sindy 2020/11/30
                  cnnConnection.Execute strExc(0)
                  GRD1.RowHeight(intRow) = 0
               End If
            End If
         End If
      End If
      '2018/10/3 END
   End If
   GRD1.Visible = True
   '2018/10/3 END
End Sub

Private Sub GRD1_DblClick()
   'Modify By Sindy 2016/3/3
   If GRD1.MouseRow <> 0 Then
      If GRD1.MouseCol <> 14 Then
   '2016/3/3 END
         cmdDetail_Click
      End If
   End If
End Sub

'Private Sub grd1_SelChange()
'Dim j As Integer 'Add By Sindy 2016/3/4
'
'GRD1.Visible = False
''Add By Sindy 2016/3/4
'If GRD1.MouseRow = 0 Then
'   '已選取的資料列清除反白
'   For j = 1 To GRD1.Rows - 1
'      If GRD1.TextMatrix(j, 0) = "V" Then
'         GRD1.col = 0
'         GRD1.row = j
'         GRD1.Text = ""
'         For i = 0 To GRD1.Cols - 1
'            GRD1.col = i
'            GRD1.CellBackColor = QBColor(15)
'         Next i
'         Call SetColColor(j)
'         Exit For
'      End If
'   Next j
'Else
''2016/3/4 END
'   '上一筆資料列清除反白
'   'Modify By Sindy 2016/5/9
'   'If dblPrevRow > 0 Then
'   If dblPrevRow > 0 And dblPrevRow <= (GRD1.Rows - 1) Then
'   '2016/5/9 END
'      GRD1.col = 0
'      GRD1.row = dblPrevRow
'      GRD1.Text = ""
'      For i = 0 To GRD1.Cols - 1
'         GRD1.col = i
'         GRD1.CellBackColor = QBColor(15)
'      Next i
'      Call SetColColor(CInt(dblPrevRow))
'   End If
'   '目前資料列反白
'   GRD1.col = 0
'   GRD1.row = GRD1.MouseRow
'   dblPrevRow = GRD1.row
''   If GRD1.Text = "V" Then
''      GRD1.Text = ""
''      For i = 0 To GRD1.Cols - 1
''         GRD1.col = i
''         GRD1.CellBackColor = QBColor(15)
''      Next i
''   Else
'      If GRD1.TextMatrix(GRD1.row, 1) <> "" Then
'         GRD1.Text = "V"
'         For i = 0 To GRD1.Cols - 1
'            GRD1.col = i
'            GRD1.CellBackColor = &HFFC0C0
'         Next i
'      End If
''   End If
'End If
'GRD1.Visible = True
'End Sub

'Add By Sindy 2016/3/4
Private Sub SetColColor(intRow As Integer)
Dim strChkDate As String, strCP10 As String
Dim Str01 As String, Str02 As String, Str03 As String, Str04 As String
Dim bolChangeColor As Boolean
Dim i As Integer
   
   GRD1.row = intRow
   If GRD1.TextMatrix(intRow, 8) <> "" Then
      '逾(含當日)本所期限變淺紅色
      If DBDATE(GRD1.TextMatrix(intRow, 8)) <= strSrvDate(1) Then
         For i = 0 To GRD1.Cols - 1
            'Modify By Sindy 2020/11/30
            If i <> 4 Then
            '2020/11/30 END
               GRD1.col = i
               GRD1.CellBackColor = &H8080FF
            End If
         Next i
         bolChangeColor = True
      End If
   End If
   'Modify By Sindy 2024/12/18 淑華提的需求+指定送件日
   If GRD1.TextMatrix(intRow, 21) <> "" Then
      '逾(含當日)指定送件日變淺紅色
      If DBDATE(GRD1.TextMatrix(intRow, 21)) <= strSrvDate(1) Then
         For i = 0 To GRD1.Cols - 1
            If i <> 4 Then
               GRD1.col = i
               GRD1.CellBackColor = &H8080FF
            End If
         Next i
         bolChangeColor = True
      End If
   End If
   '2024/12/18 END
   
   'Add By Sindy 2020/11/30
   If GRD1.TextMatrix(intRow, 20) <> "" Then
      GRD1.col = 4
      GRD1.CellBackColor = &HFF00FF 'QBColor(Rnd * 5)
   End If
   '2020/11/30 END
   
   'Add By Sindy 2019/10/17
   '商標處核判逾時提醒
   '商申案：一個工作日
   '商爭案、馬德里案：二個工作日
   '核稿人或判發人應於規定時間內完成，惟送核或送判在16:00以後則以下一工作日起算。
   '在待核判區若為上述案件以黃色顯示
   If Left(Pub_StrUserSt03, 2) = "P2" And bolChangeColor = False Then
      Str01 = SystemNumber(GRD1.TextMatrix(intRow, 3), 1)
      Str02 = SystemNumber(GRD1.TextMatrix(intRow, 3), 2)
      Str03 = SystemNumber(GRD1.TextMatrix(intRow, 3), 3)
      Str04 = SystemNumber(GRD1.TextMatrix(intRow, 3), 4)
      
      strChkDate = GRD1.TextMatrix(intRow, 18)
      If strChkDate = "" Then
         '個案檢查:因有可能有送核等又有聯絡
         strSql = "select eep01,eep04,eep05,EEP06,EEP07,decode(sign(EEP07-160000),1,to_char(to_date(EEP06,'YYYYMMDD')+1,'YYYYMMDD'),EEP06) AS ChkDate" & _
                  " from empelectronprocess,caseprogress" & _
                  " where cp01='" & Str01 & "' and cp02='" & Str02 & "' and cp03='" & Str03 & "' and cp04='" & Str04 & "'" & _
                  " and cp09=eep01" & _
                  " and eep04 in('" & EMP_送核 & "','" & EMP_送判 & "') and eep09='Y'" & _
                  " and eep05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
                  " order by eep06 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If "" & RsTemp.Fields("ChkDate") <> "" Then
               strChkDate = RsTemp.Fields("ChkDate")
            End If
         End If
      End If
      If strChkDate <> "" Then
         strCP10 = GRD1.TextMatrix(intRow, 19)
         '商爭案、馬德里案：二個工作日
         'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
         If (InStr(TMdebate, strCP10) > 0 And Not (Str01 = "FCT" And InStr(FCT_NotTMdebate, strCP10) > 0)) Or Str01 = "TF" Then
            strChkDate = CompWorkDay(3, strChkDate, 0) '不含當天
         '商申案：一個工作日
         Else
            strChkDate = CompWorkDay(2, strChkDate, 0) '不含當天
         End If
         '核判期限到期或逾期均顯示黃色
         If strChkDate <= strSrvDate(1) Then
            For i = 0 To GRD1.Cols - 1
               'Modify By Sindy 2020/11/30
               If i <> 4 Then
               '2020/11/30 END
                  GRD1.col = i
                  GRD1.CellBackColor = &H80FFFF   '黃色
                  Me.Tag = "有黃色期限"
               End If
            Next i
            bolChangeColor = True
         End If
      End If
   End If
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   
   getGrdColRow GRD1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   'GRD1.col = nCol
   GRD1.row = nRow
   If Me.GRD1.row < 1 And Me.GRD1.Text <> "V" Then
      If Me.GRD1.Text = "目次" Then
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

Private Sub GRD1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static iRow As Integer, iCol As Integer
   
   'GRD1.ToolTipText = ""
   If GRD1.MouseRow <> 0 And _
      GRD1.MouseCol = 8 Then
'      If iRow <> GRD1.MouseRow Or iCol <> GRD1.MouseCol Then
         If GRD1.TextMatrix(GRD1.MouseRow, 21) <> "" Then
            CreateToolTip GetHWndForToolTip(GRD1), "指定送件日=" & GRD1.TextMatrix(GRD1.MouseRow, 21)
'            iRow = GRD1.MouseRow
'            iCol = GRD1.MouseCol
         Else
            CreateToolTip GetHWndForToolTip(GRD1), ""
         End If
'      End If
   Else
      CreateToolTip GetHWndForToolTip(GRD1), ""
   End If
End Sub
