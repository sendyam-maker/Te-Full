VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210147 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件目前表單"
   ClientHeight    =   5750
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   8960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   8960
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "frm210147.frx":0000
      Left            =   3990
      List            =   "frm210147.frx":0002
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   120
      Width           =   1425
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "畫面更新(&Q)"
      Default         =   -1  'True
      Height          =   360
      Left            =   6030
      TabIndex        =   2
      Top             =   30
      Width           =   1125
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "明細資料"
      Height          =   360
      Left            =   7200
      TabIndex        =   3
      Top             =   30
      Width           =   885
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   8115
      TabIndex        =   4
      Top             =   30
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm210147.frx":0004
      Height          =   5055
      Left            =   60
      TabIndex        =   5
      Top             =   465
      Width           =   8835
      _ExtentX        =   15593
      _ExtentY        =   8908
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V|表單編號|表單類別|本所案號|總收文號|案件性質|本所期限|法定期限|目前處理人員|目前表單狀態"
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "註：接洽單簽核的〔通知信件〕是關閉此作業時，才統一寄發；以免太多通知信件。"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   2250
      TabIndex        =   8
      Top             =   5550
      Width           =   6660
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   285
      Left            =   990
      TabIndex        =   0
      Top             =   120
      Width           =   1710
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3016;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "員工姓名："
      Height          =   240
      Left            =   0
      TabIndex        =   7
      Top             =   150
      Width           =   945
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      Caption         =   "表單類別："
      Height          =   180
      Left            =   2970
      TabIndex        =   6
      Top             =   150
      Width           =   975
   End
End
Attribute VB_Name = "frm210147"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/24 Form2.0已修改
'Create by Sindy 2015/1/9
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim i As Integer, j As Integer
Dim dblPrevRow As Double

'查詢明細資料
Private Sub cmdDetail_Click()
Dim nFrm As Form 'Added by Lydia 2023/08/17
Dim intFCState As String, strST15 As String, strSysKind As String, strNation As String 'Add by Amy 2025/04/10
Dim strCCM18 As String 'Add by Amy 2025/06/19

   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         'Modify By Sindy 2022/8/22
         If Len(GRD1.TextMatrix(i, PUB_MGridGetId("表單編號", GRD1))) = 10 Then '接洽單
            Screen.MousePointer = vbHourglass
            '呼叫接洽單
            frm090801_New.SetParent Me
            frm090801_New.m_SignFlowEmp = IIf(Left(Combo1, 4) = "MCTF", Left(Combo1, 6), Left(Combo1, 5))
            If Trim(GRD1.TextMatrix(i, PUB_MGridGetId("目前表單狀態", GRD1))) = "退回" Then '目前表單狀況=退回
               frm090801_New.m_blnCallPrint = False
            Else
               frm090801_New.m_blnCallPrint = True
            End If
            'Added by Lydia 2023/08/17 查名單輸入
            Set nFrm = Forms(0).GetForm("frm090126")
            If Not nFrm Is Nothing Then
               Set frm090801_New.Tmpfrm090126 = nFrm
            End If
            'end 2023/08/17
            frm090801_New.Text5 = Trim(GRD1.TextMatrix(i, PUB_MGridGetId("表單編號", GRD1)))
            Call frm090801_New.cmdok_Click(4)
            frm090801_New.ZOrder
            frm090801_New.Show
            Screen.MousePointer = vbDefault
            Me.Hide
            Exit For
         Else
         '2022/8/22 END
            'Add by Amy 2025/04/10 +FC結案單
            intFCState = 0 '非FC結案單
            strST15 = PUB_GetStaffST15(GRD1.TextMatrix(i, PUB_MGridGetId("F0316", GRD1)), 1)
            strSysKind = GRD1.TextMatrix(i, PUB_MGridGetId("SYSKIND", GRD1))
            strNation = GetPrjNation1(GRD1.TextMatrix(i, PUB_MGridGetId("本所案號", GRD1)))
            If strSrvDate(1) >= FCP結案單電子化啟用日 Then
               'Modify by Amy 2025/06/26 發現舊資料會頁籤判斷會有問題FCP-065275 由P案轉入,1080730閉卷-年費時,由林士堯結案
               '       ex:FCP-065275 由P案轉入,1080730閉卷-年費時,由林士堯結案 / 外商承辦使用國內結案單操作結案 ex:T-242111(結案單號11203939)
               strCCM18 = Pub_GetField("CloseCaseMain", "CCM01='" & GRD1.TextMatrix(i, PUB_MGridGetId("表單編號", GRD1)) & "'", "CCM18")
               If strCCM18 = "F" Then
                  If strSysKind = "FCP" Or strSysKind = "FG" Or strSysKind = "P" Or strSysKind = "CFP" Then
                     intFCState = 2
                  Else
                     intFCState = 1
                  End If
               End If
               'end 2025/06/26
            End If
            frm210147_1.intFCState = intFCState
            frm210147_1.m_NP07 = GRD1.TextMatrix(i, PUB_MGridGetId("CP10", GRD1))
            'end 2025/04/10
            
            Call frm210147_1.SetParent(Me)
            frm210147_1.Hide
            frm210147_1.m_SignFlowEmp = IIf(Left(Combo1, 4) = "MCTF", Left(Combo1, 6), Left(Combo1, 5)) 'Modify by Amy 2018/10/29
            frm210147_1.txtF0301 = GRD1.TextMatrix(i, PUB_MGridGetId("表單編號", GRD1)) '表單編號
            'Add by Amy 2020/05/21
            frm210147_1.m_stNP01 = GRD1.TextMatrix(i, PUB_MGridGetId("總收文號", GRD1)) '總收文號
            frm210147_1.m_stNP22 = GRD1.TextMatrix(i, PUB_MGridGetId("ti04", GRD1)) '下一程序序號
            frm210147_1.Show
            frm210147_1.QueryData
            Me.Hide
            Exit For
         End If
      End If
   Next i
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdQuery_Click()
   If QueryData = False Then ShowNoData
End Sub

Public Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strCon As String
Dim strTB As String, strWhrCP As String, strWhrNp As String, strWhrBase As String 'Add by Amy 2025/04/10
   
   m_blnColOrderAsc = True
   QueryData = True
   GRD1.Clear
   SetGrd
   
   'Modify by Amy 2018/10/29
   'Modify By Sindy 2022/10/17
   strCon = "F0316='" & IIf(Left(Combo1, 4) = "MCTF", Left(Combo1, 6), Left(Combo1, 5)) & "'" & _
            " and ((length(F0301)=10 and F0309 not in('" & Flow_已完成 & "','" & Flow_已分案 & "','" & Flow_已收文 & "','" & Flow_放棄案源 & "'))" & _
              " or (length(F0301)=8 and F0309 in('" & Flow_主管審核中 & "','" & Flow_處理中 & "','" & Flow_退回 & "')))"
            '2022/10/17 END
   'end 2018/10/29
   If InStr(Combo2.Text, "全部") = 0 Then
      strCon = strCon & " and F0302='" & Trim(Left(Combo2.Text, 2)) & "'"
   End If
   'Add by Amy 2025/04/10 +FC結案單,將Flow003中屬於結案單資料者拆至結案單主檔中
   If strSrvDate(1) >= FCP結案單電子化啟用日 Then
      strTB = ",CloseCaseMain"
      strWhrCP = " And F0301=CCM01 and ccm03 is null and length(ccm02)=9 and ccm02=cp09 "
      strWhrNp = " And F0301=CCM01 and ccm03 is not null and ccm02=NP01(+) and ccm03=NP22(+) "
      strWhrBase = " And F0301=CCM01 and length(ccm02)>=10" & _
                               " and substr(ccm02,1,length(ccm02)-9)=PA01(+) and substr(ccm02,length(ccm02)-8,6)=PA02(+)" & _
                               " and substr(ccm02,length(ccm02)-2,1)=PA03(+) and substr(ccm02,length(ccm02)-1,2)=PA04(+) "
   Else
      strWhrCP = " and F0304 is null and length(F0303)=9 and F0303=cp09 "
      strWhrNp = " and F0304 is not null and F0303=NP01(+) and F0304=NP22(+) "
      strWhrBase = " and length(F0303)>=10 and substr(F0303,1,length(F0303)-9)=PA01(+) and substr(F0303,length(F0303)-8,6)=PA02(+) " & _
                               "and substr(F0303,length(F0303)-2,1)=PA03(+) and substr(F0303,length(F0303)-1,2)=PA04(+) "
   End If
   
   Screen.MousePointer = vbHourglass
   'Modify by Amy 2018/06/19 非P 案結案電子化,加入其他基本檔
   'Modify by Amy 2020/05/20 +商標延展結案記錄檔/ti04/ti01
   'Modify By Sindy 2022/11/4 + ,'' 急件
   'Modify by Amy 2025/04/10 +FC結案單,f0316,cp01,cp10
   strSql = "select '' V,sqldatet(F0311)||' '||sqltime(F0312) 表單日期,decode(F0302," & ShowFlow表單類別中文 & ") 表單類別,CP01||'-'||CP02||'-'||CP03||'-'||CP04 本所案號,CP09 總收文號,DECODE(PA09,'000',CPM03,Decode(PA09,'1',Nvl(CPM03,CPM04),CPM04)) 案件性質,casename 案件名稱,tm09 商品類別" & _
            ",sqldatet(cp06) 本所期限,sqldatet(cp07) 法定期限,st02 目前處理人員,decode(F0309," & ShowFlow表單狀態中文 & ") 目前表單狀態,'' ti04,'' 結案單日,'' 急件,F0301 表單編號,F0316,CP01 as SYSKIND,CP10" & _
            " from (" & _
            "select flow003.*,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,pa09,pa05||pa06||pa07 as casename,'' as tm09 from flow003" & strTB & ",caseprogress,patent where " & strCon & " and length(f0301)=8 " & strWhrCP & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) And PA01 is not null" & _
            " Union select flow003.*,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,TM10 as pa09,TM05||TM06||TM07 as casename,tm09 from flow003" & strTB & ",caseprogress,TradeMark where " & strCon & " and length(f0301)=8 " & strWhrCP & " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) And TM01 is not null" & _
            " Union select flow003.*,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,LC15 as pa09,LC05||LC06||LC07 as casename,'' as tm09 from flow003" & strTB & ",caseprogress,LawCase where " & strCon & " and length(f0301)=8 " & strWhrCP & " and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) And LC01 is not null" & _
            " Union select flow003.*,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,'1'  as pa09,HC06 as casename,'' as tm09 from flow003" & strTB & ",caseprogress,HireCase where " & strCon & " and length(f0301)=8 " & strWhrCP & " and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) And HC01 is not null" & _
            " Union select flow003.*,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,SP09 as pa09,SP05||SP06||SP07 as casename,'' as tm09 from flow003" & strTB & ",caseprogress,ServicePractice where " & strCon & " and length(f0301)=8 " & strWhrCP & " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) And SP01 is not null" & _
            " Union select flow003.*,np02,np03,np04,np05,np01,np07,np08,np09,pa09,pa05||pa06||pa07 as casename,'' as tm09 from flow003" & strTB & ",nextprogress,patent where " & strCon & " and length(f0301)=8 " & strWhrNp & " and np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) And PA01 is not null" & _
            " Union select flow003.*,np02,np03,np04,np05,np01,np07,np08,np09,Tm10 as pa09,TM05||TM06||TM07 as casename,tm09 from flow003" & strTB & ",nextprogress,TradeMark where " & strCon & " and length(f0301)=8 " & strWhrNp & " and np02=tm01(+) and np03=tm02(+) and np04=tm03(+) and np05=tm04(+) And TM01 is not null" & _
            " Union select flow003.*,np02,np03,np04,np05,np01,np07,np08,np09,LC15 as pa09,LC05||LC06||LC07 as casename,'' as tm09 from flow003" & strTB & ",nextprogress,LawCase where " & strCon & " and length(f0301)=8 " & strWhrNp & " and np02=lc01(+) and np03=lc02(+) and np04=lc03(+) and np05=lc04(+) And LC01 is not null" & _
            " Union select flow003.*,np02,np03,np04,np05,np01,np07,np08,np09,'1'  as pa09,HC06 as casename,'' as tm09 from flow003" & strTB & ",nextprogress,HireCase where " & strCon & " and length(f0301)=8 " & strWhrNp & " and np02=hc01(+) and np03=hc02(+) and np04=hc03(+) and np05=hc04(+) And HC01 is not null" & _
            " Union select flow003.*,np02,np03,np04,np05,np01,np07,np08,np09,SP09 as pa09,SP05||SP06||SP07 as casename,'' as tm09 from flow003" & strTB & ",nextprogress,ServicePractice where " & strCon & " and length(f0301)=8 " & strWhrNp & " and np02=sp01(+) and np03=sp02(+) and np04=sp03(+) and np05=sp04(+) And SP01 is not null" & _
            " Union select flow003.*,PA01,PA02,PA03,PA04,'','',0,0,pa09,pa05||pa06||pa07 as casename,'' as tm09 from flow003" & strTB & ",patent where " & strCon & " and length(f0301)=8 " & strWhrBase & " And PA01 is not null" & _
            " Union select flow003.*,TM01,TM02,TM03,TM04,'','',0,0,TM10 as pa09,TM05||TM06||TM07 as casename,tm09 from flow003" & strTB & ",TradeMark where " & strCon & " and length(f0301)=8 " & Replace(strWhrBase, "PA", "TM") & " And TM01 is not null" & _
            " Union select flow003.*,LC01,LC02,LC03,LC04,'','',0,0,LC15 as pa09,LC05||LC06||LC07 as casename,'' as tm09 from flow003" & strTB & ",LawCase where " & strCon & " and length(f0301)=8 " & Replace(strWhrBase, "PA", "LC") & " And LC01 is not null" & _
            " Union select flow003.*,HC01,HC02,HC03,HC04,'','',0,0,'1' as pa09,HC06 as casename,'' as tm09 from flow003" & strTB & ",HireCase where " & strCon & " and length(f0301)=8 " & Replace(strWhrBase, "PA", "HC") & " And  HC01 is not null" & _
            " Union select flow003.*,SP01,SP02,SP03,SP04,'','',0,0,SP09 as pa09,SP05||SP06||SP07 as casename,'' as tm09 from flow003" & strTB & ",ServicePractice where " & strCon & " and length(f0301)=8 " & Replace(strWhrBase, "PA", "SP") & " And SP01 is not null" & _
            "),CASEPROPERTYMAP,staff" & _
            " where cp01=cpm01(+) and cp10=cpm02(+) and F0308=st01(+)"
   'Modify By Sindy 2024/9/20 + and np22 is not null;防止把下一程序沒有資料的也抓出來了 ex:Ti02=AA2002669 /Ti04=184600 /Ti05=FCT閉卷後轉T案收文補結案資料
   'Modify by Amy 2025/04/10 +F0316,TM01,NP07
   strSql = strSql & " Union Select '' V,'' 表單日期,'結案單' 表單類別,CP01||'-'||CP02||'-'||CP03||'-'||CP04 本所案號,CP09 總收文號,Decode(TM10,'000',CPM03,Decode(TM10,'1',Nvl(CPM03,CPM04),CPM04)) 案件性質,TM05||TM06||TM07 案件名稱,tm09 商品類別" & _
            ",sqldatet(np08) 本所期限,sqldatet(np09) 法定期限,'" & GetTS1Name & "' 目前處理人員,Decode(ti06,'Y','退回',Decode('02'," & ShowFlow表單狀態中文 & ")) 目前表單狀態,''||ti04,''||ti01 結案單日,'' 急件,'' 表單編號,'' F0316,TM01 as SYSKIND,NP07 as CP10" & _
            " From T102Inform,NextProgress,CaseProgress,TradeMark,CasePropertyMap" & _
            " Where ti02=np01(+) And ti04=np22(+) And np22 is not null And ti02=cp09(+) And cp01=cpm01(+) And np07=cpm02(+) And ti03='" & strUserNum & "' " & _
            " And cp01=tm01(+) And cp02=tm02(+) And cp03=tm03(+) And cp04=tm04(+) And cp01 in('T','TF')  And TM29 is null And TM57 is null And NP06 is null "
   'Add By Sindy 2022/8/22 + 接洽單
   'Modify By Sindy 2023/8/2 排除已取消收文 + caseprogress; and cp140(+)=crl01 and (cp159 is null or cp159=0)
   'Modify by Amy 2025/04/10 +F0316,CRL07,CRL01
   strSql = strSql & " union select '' V,sqldatet(F0311)||' '||sqltime(F0312) 表單日期,decode(F0302," & ShowFlow表單類別中文 & ") 表單類別,decode(CRL08,null,CRL07||'-',CRL07||'-'||CRL08||'-'||CRL09||'-'||CRL10) 本所案號,'' 總收文號,GetCRCaseNmFee(crl01,'2') 案件性質,CRL17 案件名稱,CRL73 商品類別" & _
            ",sqldatet(CRL12) 本所期限,sqldatet(CRL13) 法定期限,decode(F0308," & ShowFlow特殊簽核人員 & ",st02) 目前處理人員,decode(F0309," & ShowFlow表單狀態中文 & ") 目前表單狀態,'' ti04,'' 結案單日,CRL90 急件,F0301 表單編號,F0316,CRL07 as SYSKIND,CRL01 as CP10" & _
            " from flow003,staff,ConsultRecordList,caseprogress" & _
            " where " & strCon & " and length(f0301)=10 and f0301=CRL01(+) and F0308=st01(+) and CRL01 is not null" & _
            " and cp140(+)=crl01 and (cp159 is null or cp159=0)"
   '2022/8/22 END
   strSql = strSql & " Order by 表單編號 asc,結案單日 Desc"
   'end 2018/06/19
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
   Else
      QueryData = False
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
      'ShowNoData
      Exit Function
   End If
   
   '若有資料游標停在第一筆
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
   dblPrevRow = GRD1.row
   If rsTmp.RecordCount > 0 Then
      GRD1.Text = "V"
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = &HFFC0C0
      Next i
      'Add By Sindy 2022/11/4
      For i = 1 To rsTmp.RecordCount
         Call recovercolor(i)
      Next i
      '2022/11/4 END
   End If
   GRD1.Visible = True
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
End Function

'Add By Sindy 2022/11/11
Private Sub Combo1_Click()
   Call QueryData
End Sub

'Add By Sindy 2022/11/11
Private Sub Combo2_Click()
   Call QueryData
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   Call Flow_SetF0302Combo(Combo2)
    'Add by Amy 2025/08/04
   If Left(Pub_StrUserSt03, 1) = "F" Then
      Label2.Visible = False
      Combo2.Visible = False
   End If
   'Modify By Sindy 2025/8/6 +, , Me.Name
   Call SetEmpDutyCombo(Combo1, True, , Me.Name)
   Call SetSpecMan 'Add by Amy 2018/08/16 +特殊設定
   
'   QueryData
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache '發信 Add By Sindy 2023/11/7
   Call FlowBatchSendMail(strUserNum) 'Add By Sindy 2022/10/4 整批發通知信: 電子收文(或部分簽核)通知信
   
   'Me.Form=H
   '一進入系統,檢查是否有須要開啟的作業,到此作業結束
   pub_CallNextABSForm = False
   Set frm210147 = Nothing
   
   If pub_CallNextABSForm = False Then
      Call Forms(0).SysStartCallForm 'Add By Sindy 2011/10/7
   End If
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modify By Sindy 2022/8/22
   'Modify by Amy 2025/04/10 +F0316,SYSKIND,CP10
   '                        0    1           2           3           4           5           6           7           8           9           10              11              12      13          14      15
   arrGridHeadText = Array("V", "表單日期", "表單類別", "本所案號", "總收文號", "案件性質", "案件名稱", "商品類別", "本所期限", "法定期限", _
                                          "目前處理人員", "目前表單狀態", "ti04", "結案單日", "急件", "表單編號", "F0316", "SYSKIND", "CP10")
'   If Left(Combo2.Text, 1) = "3" Then '接洽單
      arrGridHeadWidth = Array(200, 800, 500, 1000, 0, 1200, 1000, 800, 800, 800, 1000, 1200, 0, 0, 0, 1000, 0, 0, 0)
'   Else
'      arrGridHeadWidth = Array(200, 800, 500, 1000, 0, 1200, 1000, 800, 800, 800, 1000, 1200, 0, 0, 0, 1000)
'   End If
   '2022/8/22
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = 2
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Private Sub GRD1_DblClick()
   cmdDetail_Click
End Sub

Private Sub grd1_SelChange()
GRD1.Visible = False
If GRD1.MouseRow <> 0 Then
   '上一筆資料列清除反白
   If dblPrevRow > 0 Then
      GRD1.col = 0
      GRD1.row = dblPrevRow
      GRD1.Text = ""
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = QBColor(15)
      Next i
      Call recovercolor(CInt(dblPrevRow)) 'Add By Sindy 2022/11/4
   End If
   '目前資料列反白
   GRD1.col = 0
   GRD1.row = GRD1.MouseRow
   dblPrevRow = GRD1.row
'   If grd1.Text = "V" Then
'      grd1.Text = ""
'      For i = 0 To grd1.Cols - 1
'         grd1.col = i
'         grd1.CellBackColor = QBColor(15)
'      Next i
'   Else
      'Modify by Amy 2020/05/21 原抓1-表單編號,改抓案號欄,因商標延展不會有表單編號
      If GRD1.TextMatrix(GRD1.row, PUB_MGridGetId("本所案號", GRD1)) <> "" Then
         GRD1.Text = "V"
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
         Call recovercolor(GRD1.row) 'Add By Sindy 2022/11/4
      End If
'   End If
End If
GRD1.Visible = True
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow GRD1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   GRD1.col = nCol
   GRD1.row = nRow
   If Me.GRD1.row < 1 And Me.GRD1.Text <> "V" Then
      If Me.GRD1.Text = "表單編號" Or Me.GRD1.Text = "天數" Or Me.GRD1.Text = "時數" Then
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

'欄位變色
Private Sub recovercolor(intRow As Integer)
Dim j As Integer
   
   If GRD1.TextMatrix(intRow, 0) = "V" Then Exit Sub
   
   GRD1.row = intRow
   'Add By Sindy 2022/11/4 急件顯示紅色
   If Trim(GRD1.TextMatrix(intRow, PUB_MGridGetId("急件", GRD1))) = "Y" Then
      For j = 1 To 11
         GRD1.col = j
         GRD1.CellBackColor = &H8080FF '紅色
      Next j
   End If
End Sub

'Add by Amy 2018/08/16
Private Sub SetSpecMan()
    Dim arrData
    Dim strTemp As String
    Dim strMCTF As String 'Add by Amy 2018/10/29
    
    'Modify by Amy 2018/10/29 智權為MCTF被退回無法操作 ex:T-214458 紹禎
    'modify by sonia 2019/3/15 改用GetMCTF0XAllCode,因為MCTF01及MCTF04都有鄭天雲A4021
    'strMCTF = GetMCTF0XCode(strUserNum)
    'If strMCTF <> MsgText(601) Then
    '    Combo1.AddItem strMCTF & " " & GetPrjSalesNM(CStr(strMCTF))
    'End If
    strMCTF = GetMCTF0XAllCode(strUserNum)
    If strMCTF <> MsgText(601) Then
        arrData = Split(strMCTF, "','")
        For i = 0 To UBound(arrData)
            'Modify By Sindy 2025/8/5 下拉選單已存在的人員,就不需再加入
            For j = 0 To Combo1.ListCount - 1
               If InStr(Combo1.List(j), CStr(arrData(i))) = 1 Then
                  Exit For
               End If
            Next j
            If j = Combo1.ListCount Then
            '2025/8/5 END
               Combo1.AddItem arrData(i) & " " & GetPrjSalesNM(CStr(arrData(i)))
            End If
        Next
    End If
    'end 2019/3/15
    If InStr(Pub_GetSpecMan("A8"), strUserNum) > 0 Then
        strTemp = Pub_GetSpecMan("A7")
        arrData = Split(strTemp, ";")
        For i = 0 To UBound(arrData)
            'Modify By Sindy 2025/8/5 下拉選單已存在的人員,就不需再加入
            For j = 0 To Combo1.ListCount - 1
               If InStr(Combo1.List(j), CStr(arrData(i))) = 1 Then
                  Exit For
               End If
            Next j
            If j = Combo1.ListCount Then
            '2025/8/5 END
               Combo1.AddItem arrData(i) & " " & GetPrjSalesNM(CStr(arrData(i)))
            End If
        Next
    End If
    'end 2018/10/29
End Sub

Private Function GetTS1Name() As String
    Dim RsQ As New ADODB.Recordset
    Dim stQ As String, intQ As Integer
    
    GetTS1Name = ""
    stQ = " And ST01 in ('" & Replace(Pub_GetSpecMan("TS1"), ";", "','") & "')"
    stQ = "Select * From Staff Where 1=1 " & stQ
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, stQ)
    If intQ = 1 Then
        Do While Not RsQ.EOF
            GetTS1Name = GetTS1Name & "," & RsQ.Fields("st02")
            RsQ.MoveNext
        Loop
        If GetTS1Name <> MsgText(601) Then
            GetTS1Name = Mid(GetTS1Name, 2)
        End If
    End If
    Set RsQ = Nothing
End Function
