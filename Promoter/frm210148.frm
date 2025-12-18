VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210148 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件簽核作業"
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
      ItemData        =   "frm210148.frx":0000
      Left            =   3900
      List            =   "frm210148.frx":0002
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   90
      Width           =   1425
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "全選(&A)"
      Height          =   360
      Index           =   0
      Left            =   5370
      TabIndex        =   2
      Top             =   30
      Width           =   765
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "畫面更新(&Q)"
      Default         =   -1  'True
      Height          =   360
      Index           =   1
      Left            =   6180
      TabIndex        =   3
      Top             =   30
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "簽核(&O)"
      Height          =   360
      Index           =   2
      Left            =   7350
      TabIndex        =   4
      Top             =   30
      Width           =   765
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Index           =   3
      Left            =   8145
      TabIndex        =   5
      Top             =   30
      Width           =   765
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm210148.frx":0004
      Height          =   4905
      Left            =   60
      TabIndex        =   6
      Top             =   630
      Width           =   8835
      _ExtentX        =   15593
      _ExtentY        =   8661
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V|表單編號|表單類別|智權人員|本所案號|總收文號|案件性質|本所期限|法定期限|急件"
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
   Begin VB.Label Lbl_A_Note 
      AutoSize        =   -1  'True
      Caption         =   "注意：建檔時請從”第一筆”依序處理，因會影響到收文的給號順序。"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   90
      TabIndex        =   10
      Top             =   450
      Width           =   5850
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "註：接洽單簽核的〔通知信件〕是關閉此作業時，才統一寄發；以免太多通知信件。"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   2220
      TabIndex        =   9
      Top             =   5550
      Width           =   6660
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   285
      Left            =   1050
      TabIndex        =   0
      Top             =   90
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
      Caption         =   "簽核人員："
      Height          =   240
      Left            =   60
      TabIndex        =   8
      Top             =   120
      Width           =   945
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      Caption         =   "表單類別："
      Height          =   240
      Left            =   2880
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frm210148"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/27 Form2.0已修改
'Create by Sindy 2015/1/12
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim i As Integer, j As Integer
Public cmdState As Integer '紀錄作用按鍵
Public m_FlowType As String 'Add By Sindy 2022/8/4 A.櫃檯接洽單
Dim strAllEmp As String 'Add By Sindy 2023/1/5

Private Sub cmdok_Click(Index As Integer)
   cmdState = Index
   
   PubShowNextData
End Sub

Public Sub PubShowNextData()
   Dim strF0302 As String 'Add by Amy 2022/09/16
   Dim intFCState As String, strST15 As String, strSysKind As String, strNation As String 'Add by Amy 2025/04/10
   Dim strCCM18 As String 'Add by Amy 2025/06/19
   
   Select Case cmdState
      Case 0 '全選
         GRD1.Visible = False
         If GRD1.Rows > 1 Then
            If GRD1.TextMatrix(1, PUB_MGridGetId("表單編號", GRD1)) <> "" Then
               For j = 1 To GRD1.Rows - 1
                  GRD1.col = 0
                  GRD1.row = j
                  GRD1.Text = "V"
                  For i = 0 To GRD1.Cols - 1
                     GRD1.col = i
                     GRD1.CellBackColor = &HFFC0C0
                  Next i
                  Call recovercolor(j) 'Add By Sindy 2022/11/4
               Next j
            End If
         End If
         GRD1.Visible = True
      Case 1 '查詢
         If QueryData = False Then ShowNoData
      Case 2 '簽核 or 明細(櫃檯用)
         Me.Enabled = False
         For i = 1 To GRD1.Rows - 1
            GRD1.col = 0
            GRD1.row = i
            If Trim(GRD1.Text) = "V" Then
               GRD1.col = 0
               GRD1.Text = ""
               For j = 0 To GRD1.Cols - 1
                  GRD1.col = j
                  GRD1.CellBackColor = QBColor(15)
               Next j
               Call recovercolor(i) 'Add By Sindy 2022/11/4
               'Modify By Sindy 2022/8/22
               If Len(GRD1.TextMatrix(i, PUB_MGridGetId("表單編號", GRD1))) = 10 Then '接洽單
                  Screen.MousePointer = vbHourglass
                  Me.Hide
                  If m_FlowType = "A" Then '櫃檯接洽單
                     'Add by Amy 2022/08/30
                     Call frm210148_2.SetParent(Me) '呼叫接洽單
                     frm210148_2.Show
                     frm210148_2.txtF0301 = GRD1.TextMatrix(i, PUB_MGridGetId("表單編號", GRD1)) '表單編號
                     frm210148_2.QueryData
                  Else
                     frm090801_New.SetParent Me
                     frm090801_New.m_SignFlowEmp = Trim(GRD1.TextMatrix(i, PUB_MGridGetId("F0308", GRD1))) 'Left(Combo1, 5) =>F0308
                     frm090801_New.m_blnCallPrint = True
                     frm090801_New.Text5 = Trim(GRD1.TextMatrix(i, PUB_MGridGetId("表單編號", GRD1)))
                     Call frm090801_New.cmdok_Click(4)
                     frm090801_New.ZOrder
                     frm090801_New.Show
                  End If
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Exit Sub
               Else
               '2022/8/22 END
                  GRD1.col = PUB_MGridGetId("總收文號", GRD1)
                  If Not IsNull(GRD1.Text) Then
                     Screen.MousePointer = vbHourglass
                     Me.Hide
                     'Add by Amy 2025/04/10
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
                     frm210148_1.intFCState = intFCState
                     frm210148_1.m_NP07 = GRD1.TextMatrix(i, PUB_MGridGetId("CP10", GRD1))
                     'end 2025/04/10
                     
                     Call frm210148_1.SetParent(Me)
                     frm210148_1.Hide
                     frm210148_1.m_SignFlowEmp = Trim(GRD1.TextMatrix(i, PUB_MGridGetId("F0308", GRD1))) 'Left(Combo1, 5) =>F0308
                     frm210148_1.txtF0301 = GRD1.TextMatrix(i, PUB_MGridGetId("表單編號", GRD1)) '表單編號
                     frm210148_1.Command1(1).Visible = False '進度
                     frm210148_1.Command1(3).Visible = False '完整卷宗
                     frm210148_1.cmdFile.Visible = True '檢覆回覆單
                     frm210148_1.PreF0302 = GetFlowNo(GRD1.TextMatrix(i, PUB_MGridGetId("表單類別", GRD1))) 'Add by Amy 2022/09/16
                     frm210148_1.Show
                     frm210148_1.QueryData
                     Screen.MousePointer = vbDefault
                     Me.Enabled = True
                     Exit Sub
                  End If
               End If
            End If
         Next i
         Me.Enabled = True
         Call QueryData
      Case 3 '結束
         Unload Me
      Case Else
   End Select
End Sub

Public Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strCon As String
Dim strTB As String, strWhrBase As String, strWhrCP As String, strWhrNp As String 'Add by Amy 2025/04/10
   
   m_blnColOrderAsc = True
   QueryData = True
   GRD1.Clear
   SetGrd
   
   'Modify By Sindy 2022/8/22
   'Modify by Amy 2022/08/30 +m_FlowType= "A" 櫃檯接洽單
   If Len(Trim(Left(Combo1, 5))) = 2 Or m_FlowType = "A" Then '行政人員處理中
      strCon = "F0308='" & Trim(Left(Combo1, 5)) & "' and F0309 in('" & Flow_處理中 & "')"
   Else
   '2022/8/22 END
      'Add By Sindy 2023/1/5
      If InStr(Combo1.Text, "全部") > 0 And strAllEmp <> "" Then
         strCon = "F0308 in(" & strAllEmp & ") and F0309 in('" & Flow_主管審核中 & "')"
      Else
      '2023/1/5 END
         strCon = "F0308='" & Trim(Left(Combo1, 5)) & "' and F0309 in('" & Flow_主管審核中 & "')"
      End If
   End If
   If InStr(Combo2.Text, "全部") = 0 Then
      strCon = strCon & " and F0302='" & Trim(Left(Combo2.Text, 2)) & "'"
   End If
   'Add by Amy 2025/04/10 +FC結案單,將Flow003中屬於結案單資料者拆至結案單主檔中
   If strSrvDate(1) >= FCP結案單電子化啟用日 Then
      strTB = ",CloseCaseMain"
      strWhrCP = " And F0301=CCM01 and ccm03 is null and length(ccm02)=9 and ccm02=cp09 "
      strWhrNp = " And F0301=CCM01 and ccm03 is not null and ccm02=NP01(+) and ccm03=NP22(+) "
      strWhrBase = " And F0301=CCM01 and length(CCM02)>=10 " & _
                              " and substr(CCM02,1,length(CCM02)-9)=pa01(+) and substr(CCM02,length(CCM02)-8,6)=pa02(+)" & _
                              " and substr(CCM02,length(CCM02)-2,1)=pa03(+) and substr(CCM02,length(CCM02)-1,2)=pa04(+) "
   Else
      strWhrCP = " and F0304 is null and length(F0303)=9 and F0303=cp09 "
      strWhrNp = " and F0304 is not null and F0303=NP01(+) and F0304=NP22(+) "
      strWhrBase = " and length(F0303)>=10 and substr(F0303,1,length(F0303)-9)=pa01(+) and substr(F0303,length(F0303)-8,6)=pa02(+) " & _
                              " and substr(F0303,length(F0303)-2,1)=pa03(+) and substr(F0303,length(F0303)-1,2)=pa04(+) "
   End If
   
   Screen.MousePointer = vbHourglass
   'Modify by Amy 2018/06/19 非P 案結案電子化,加入其他基本檔
   'Modify By Sindy 2022/11/4 + ,'' 急件
   'Modify by Amy 2025/04/10 +FC結案單,f0316,cp01,cp10
   strSql = "select '' V,sqldatet(F0311)||' '||sqltime(F0312) 表單日期,decode(F0302," & ShowFlow表單類別中文 & ") 表單類別,st02 智權人員,CP01||'-'||CP02||'-'||CP03||'-'||CP04 本所案號,'' 客戶,CP09 總收文號,DECODE(PA09,'000',CPM03,Decode(PA09,'1',Nvl(CPM03,CPM04),CPM04)) 案件性質,casename 案件名稱,tm09 商品類別" & _
            ",sqldatet(cp06) 本所期限,sqldatet(cp07) 法定期限,'' 急件,F0301 表單編號,F0308,F0316,CP01 as SYSKIND,CP10" & _
            " from (" & _
            "select flow003.*,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,pa09,pa05||pa06||pa07 as casename,'' as tm09 From Flow003" & strTB & ",caseprogress,patent where " & strCon & " and length(f0301)=8 " & strWhrCP & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) And PA01 is not null" & _
            " Union select flow003.*,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,TM10 as pa09,TM05||TM06||TM07 as casename,tm09 From Flow003" & strTB & ",caseprogress,TradeMark where " & strCon & " and length(f0301)=8 " & strWhrCP & " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) And TM01 is not null" & _
            " Union select flow003.*,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,LC15 as pa09,LC05||LC06||LC07 as casename,'' as tm09 From Flow003" & strTB & ",caseprogress,LawCase where " & strCon & " and length(f0301)=8 " & strWhrCP & " and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) And LC01 is not null" & _
            " Union select flow003.*,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,'1' as pa09,HC06 as casename,'' as tm09 From Flow003" & strTB & ",caseprogress,HireCase where " & strCon & " and length(f0301)=8 " & strWhrCP & " and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) And HC01 is not null" & _
            " Union select flow003.*,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,SP09 as pa09,SP05||SP06||SP07 as casename,'' as tm09 From Flow003" & strTB & ",caseprogress,ServicePractice where " & strCon & " and length(f0301)=8 " & strWhrCP & " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) And SP01 is not null" & _
            " Union select flow003.*,np02,np03,np04,np05,np01,np07,np08,np09,pa09,pa05||pa06||pa07 as casename,'' as tm09 From Flow003" & strTB & ",nextprogress,patent where " & strCon & " and length(f0301)=8 " & strWhrNp & " and np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) And PA01 is not null" & _
            " Union select flow003.*,np02,np03,np04,np05,np01,np07,np08,np09,Tm10 as pa09,TM05||TM06||TM07 as casename,tm09 From Flow003" & strTB & ",nextprogress,TradeMark where " & strCon & " and length(f0301)=8 " & strWhrNp & " and np02=tm01(+) and np03=tm02(+) and np04=tm03(+) and np05=tm04(+) And TM01 is not null" & _
            " Union select flow003.*,np02,np03,np04,np05,np01,np07,np08,np09,LC15 as pa09,LC05||LC06||LC07 as casename,'' as tm09 From Flow003" & strTB & ",nextprogress,LawCase where " & strCon & " and length(f0301)=8 " & strWhrNp & " and np02=lc01(+) and np03=lc02(+) and np04=lc03(+) and np05=lc04(+) And LC01 is not null" & _
            " Union select flow003.*,np02,np03,np04,np05,np01,np07,np08,np09,'1' as pa09,HC06 as casename,'' as tm09 From Flow003" & strTB & ",nextprogress,HireCase where " & strCon & " and length(f0301)=8 " & strWhrNp & " and np02=hc01(+) and np03=hc02(+) and np04=hc03(+) and np05=hc04(+) And HC01 is not null" & _
            " Union select flow003.*,np02,np03,np04,np05,np01,np07,np08,np09,SP09 as pa09,SP05||SP06||SP07 as casename,'' as tm09 From Flow003" & strTB & ",nextprogress,ServicePractice where " & strCon & " and length(f0301)=8 " & strWhrNp & " and np02=sp01(+) and np03=sp02(+) and np04=sp03(+) and np05=sp04(+) And SP01 is not null" & _
            " Union select flow003.*,PA01,PA02,PA03,PA04,'','',0,0,pa09,pa05||pa06||pa07 as casename,'' as tm09 From Flow003" & strTB & ",patent where " & strCon & " and length(f0301)=8 " & strWhrBase & " And PA01 is not null" & _
            " Union select flow003.*,TM01,TM02,TM03,TM04,'','',0,0,TM10 as pa09,TM05||TM06||TM07 as casename,tm09 From Flow003" & strTB & ",TradeMark where " & strCon & " and length(f0301)=8 " & Replace(UCase(strWhrBase), "PA", "TM") & " And TM01 is not null" & _
            " Union select flow003.*,LC01,LC02,LC03,LC04,'','',0,0,LC15 as pa09,LC05||LC06||LC07 as casename,'' as tm09 From Flow003" & strTB & ",LawCase where " & strCon & " and length(f0301)=8 " & Replace(UCase(strWhrBase), "PA", "LC") & " And LC01 is not null" & _
            " Union select flow003.*,HC01,HC02,HC03,HC04,'','',0,0,'1' as pa09,HC06 as casename,'' as tm09 From Flow003" & strTB & ",HireCase where " & strCon & " and length(f0301)=8 " & Replace(UCase(strWhrBase), "PA", "HC") & " And HC01 is not null" & _
            " Union select flow003.*,SP01,SP02,SP03,SP04,'','',0,0,SP09 as pa09,SP05||SP06||SP07 as casename,'' as tm09 From Flow003" & strTB & ",ServicePractice where " & strCon & " and length(f0301)=8 " & Replace(UCase(strWhrBase), "PA", "SP") & " And SP01 is not null" & _
            "),CASEPROPERTYMAP,staff" & _
            " where cp01=cpm01(+) and cp10=cpm02(+) and f0316=st01(+)"
   'end 2025/04/10
   'Add By Sindy 2022/8/30 + 接洽單
   'Modify by Amy 2025/04/10 +f0316,CRL07,CRL01
   strSql = strSql & " union select '' V,sqldatet(F0311)||' '||sqltime(F0312) 表單日期,decode(F0302," & ShowFlow表單類別中文 & ") 表單類別,st02 智權人員,decode(CRL08,null,CRL07||'-',CRL07||'-'||CRL08||'-'||CRL09||'-'||CRL10) 本所案號,GetCRAName(CRL01) 客戶,'' 總收文號,GetCRCaseNmFee(CRL01,2) 案件性質,CRL17 案件名稱,CRL73 商品類別" & _
            ",sqldatet(CRL12) 本所期限,sqldatet(CRL13) 法定期限,CRL90 急件,F0301 表單編號,F0308,F0316,CRL07 as SYSKIND,CRL01 as cp10" & _
            " from flow003,staff,ConsultRecordList" & _
            " where " & strCon & " and length(f0301)=10 and f0301=CRL01(+)  and f0316=st01(+) and CRL01 is not null "
   '2022/8/30 END
   strSql = strSql & " order by 表單編號 asc"
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
   If rsTmp.RecordCount > 0 Then
'      GRD1.Text = "V"
'      For i = 0 To GRD1.Cols - 1
'         GRD1.col = i
'         GRD1.CellBackColor = &HFFC0C0
'      Next i
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
Private Sub Combo1_Change()
   Call QueryData
End Sub

'Add By Sindy 2022/11/11
'Private Sub Combo2_Change()
'   Call QueryData
'End Sub
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
   
   'Add By Sindy 2022/8/4
   Lbl_A_Note.Visible = False 'Add By Sindy 2023/3/31
   If m_FlowType = "A" Then
      Label3.Visible = False 'Add By Sindy 2023/2/2
      Combo2.Clear
      Combo2.AddItem "3  接洽單"
      Combo2.ListIndex = 0
      Label2.Visible = False
      Combo2.Visible = False
      cmdOK(2).Caption = "明細"
      Lbl_A_Note.Visible = True 'Add By Sindy 2023/3/31
   Else
      GRD1.Top = Lbl_A_Note.Top 'Add By Sindy 2023/3/31
   End If
   '2022/8/4 END
   
   'Modify By Sindy 2022/12/6 + , True : 抓出帶的離職人員
   If m_FlowType = "A" Then
      Call SetEmpDutyCombo(Combo1, True)
   Else
      'Modify By Sindy 2023/1/5 + 全部
      'Modify By Sindy 2025/8/6 +, Me.Name
      Call SetEmpDutyCombo(Combo1, True, True, Me.Name)
   End If
   
'   'Add By Sindy 2023/1/4 張哲瑋專利工程師兼智權人員會收商標案件,簽核主管蔡順興不能簽核商標案件,
'   '                      須由中所智權部主管簽核~
'   If InStr(Pub_GetSpecMan("中所智權部主管"), strUserNum) > 0 Then
'      strSql = "select * from staff where st15='P11' and st06='2' and st04='1' and length(st01)=5 and substr(st01,4,1)<>'9' order by st15,st01"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         RsTemp.MoveFirst
'         Do While Not RsTemp.EOF
'            For i = 0 To Combo1.ListCount - 1
'               If InStr(Combo1.List(i), RsTemp(0)) = 1 Then
'                  Exit For
'               End If
'            Next i
'            If i = Combo1.ListCount Then
'               Combo1.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
'            End If
'            RsTemp.MoveNext
'         Loop
'      End If
'   End If
'   '2023/1/4 END
   
   'Add By Sindy 2023/1/5
   strAllEmp = ""
   For i = 0 To Combo1.ListCount - 1
      If Len(Trim(Left(Combo1.List(i), 5))) = 5 Then
         strAllEmp = strAllEmp & ",'" & Trim(Left(Combo1.List(i), 5)) & "'"
      End If
   Next i
   If strAllEmp <> "" Then strAllEmp = Mid(strAllEmp, 2)
   For i = 0 To Combo1.ListCount - 1
      If Trim(Left(Combo1.List(i), 5)) = strUserNum Then
         Combo1.ListIndex = i
         Exit For
      End If
   Next i
   '2023/1/5 END
   
'   Call QueryData
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim strText As String
   PUB_SendMailCache '發信 Add By Sindy 2023/11/7
   Call FlowBatchSendMail(strUserNum) 'Add By Sindy 2022/10/4 整批發通知信: 電子收文(或部分簽核)通知信
   
   'Me.Form=G
   '一進入系統,檢查是否有須要開啟此作業
   If pub_CallNextABSForm = True Then
      strText = ChkIsAbsenceMustPro
      Me.Hide
      If InStr(1, strText, "H") > 0 Then
         If TypeName(Tmpfrm210147) <> "Nothing" Then
            Tmpfrm210147.Show
         End If
      Else
         pub_CallNextABSForm = False
      End If
   End If
   
   Set frm210148 = Nothing
   If pub_CallNextABSForm = False Then
      Call Forms(0).SysStartCallForm
   End If
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modify By Sindy 2022/12/22
   'Modify by Amy 2025/04/10 +F0316,SYSKIND,CP10
   '                        0    1           2           3           4           5       6           7           8           9           10          11          12      13          14            15
   arrGridHeadText = Array("V", "表單日期", "表單類別", "智權人員", "本所案號", "客戶", "總收文號", "案件性質", "案件名稱", "商品類別", _
                                       "本所期限", "法定期限", "急件", "表單編號", "F0308", "F0316", "SYSKIND", "CP10")
   If m_FlowType = "A" Then '櫃檯接洽單
      arrGridHeadWidth = Array(200, 800, 500, 850, 500, 1000, 0, 1200, 1000, 500, 800, 800, 0, 1000, 0, 0, 0, 0)
   Else
      arrGridHeadWidth = Array(200, 800, 500, 850, 1000, 0, 0, 1200, 1000, 500, 800, 800, 0, 1000, 0, 0, 0, 0)
   End If
   '2022/12/22 END
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

'欄位變色
Private Sub recovercolor(intRow As Integer)
Dim j As Integer
   
   If GRD1.TextMatrix(intRow, 0) = "V" Then Exit Sub
   
   GRD1.row = intRow
   'Add By Sindy 2022/11/4 急件顯示紅色
   If Trim(GRD1.TextMatrix(intRow, PUB_MGridGetId("急件", GRD1))) = "Y" Then
      For j = 1 To 12
         GRD1.col = j
         GRD1.CellBackColor = &H8080FF '紅色
      Next j
   End If
End Sub

'Add By Sindy 2022/10/17
Private Sub GRD1_DblClick()
   Call cmdok_Click(2)
End Sub

Private Sub grd1_SelChange()
GRD1.Visible = False
If GRD1.MouseRow <> 0 Then
   GRD1.col = 0
   GRD1.row = GRD1.MouseRow
   If GRD1.TextMatrix(GRD1.MouseRow, PUB_MGridGetId("表單編號", GRD1)) <> "" Then
      If GRD1.Text = "V" Then
         GRD1.Text = ""
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = QBColor(15)
         Next i
      Else
         GRD1.Text = "V"
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
      End If
      Call recovercolor(GRD1.MouseRow) 'Add By Sindy 2022/11/4
   End If
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
      If Me.GRD1.Text = "表單編號" Then
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

'Add by Amy 2022/09/16
Private Function GetFlowNo(ByVal stName As String) As String
    Select Case stName
        Case "結案單"
            GetFlowNo = "1"
        Case "銷案銷帳單"
            GetFlowNo = "2"
        Case "接洽單"
            GetFlowNo = "3"
    End Select
End Function
