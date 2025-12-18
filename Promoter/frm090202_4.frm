VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090202_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "待送件區"
   ClientHeight    =   5750
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   9530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   9530
   Begin VB.CommandButton cmdNewFlow 
      Caption         =   "新增歷程(&A)"
      Height          =   360
      Left            =   4830
      TabIndex        =   9
      Top             =   408
      Visible         =   0   'False
      Width           =   1160
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   432
      Width           =   3975
      Begin MSForms.ComboBox Combo1 
         Height          =   300
         Left            =   1170
         TabIndex        =   7
         Top             =   0
         Width           =   2085
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3678;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "人員："
         Height          =   240
         Left            =   0
         TabIndex        =   6
         Top             =   60
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "畫面更新(&Q)"
      Height          =   360
      Left            =   6030
      TabIndex        =   1
      Top             =   408
      Width           =   1160
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "明細資料(&D)"
      Default         =   -1  'True
      Height          =   360
      Left            =   7230
      TabIndex        =   0
      Top             =   408
      Width           =   1310
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   8570
      TabIndex        =   2
      Top             =   408
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Height          =   4308
      Left            =   60
      TabIndex        =   3
      Top             =   1368
      Width           =   9408
      _ExtentX        =   16598
      _ExtentY        =   7602
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V|判發日期|收文日| 本所案號 |案件名稱| 國家| 種類|  案件性質| 本所期限|承辦人|智權人員"
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
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Left            =   1296
      TabIndex        =   13
      Top             =   72
      Visible         =   0   'False
      Width           =   2088
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3678;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      Caption         =   "程序人員："
      Height          =   240
      Left            =   312
      TabIndex        =   12
      Top             =   108
      Visible         =   0   'False
      Width           =   948
   End
   Begin VB.Label LblFCPNote 
      Caption         =   "需經主管機關，經程序二級主管判發是”送件”狀態的，承辦人呈現黃色。"
      ForeColor       =   &H00FF0000&
      Height          =   228
      Left            =   96
      TabIndex        =   11
      Top             =   1152
      Width           =   6684
   End
   Begin VB.Label Label3 
      Caption         =   "共 0 筆"
      ForeColor       =   &H00C00000&
      Height          =   228
      Left            =   7140
      TabIndex        =   10
      Top             =   1068
      Width           =   1464
   End
   Begin VB.Label Label2 
      Caption         =   "顯示顏色說明：已處理過電子檔，為紅色(Y)；電子送件，為淡綠色。"
      ForeColor       =   &H00FF0000&
      Height          =   228
      Left            =   96
      TabIndex        =   8
      Top             =   972
      Width           =   6684
   End
   Begin VB.Label Label16 
      Caption         =   "註：雙擊選取時，開啟承辦歷程。"
      ForeColor       =   &H000000C0&
      Height          =   228
      Left            =   96
      TabIndex        =   4
      Top             =   792
      Width           =   6684
   End
End
Attribute VB_Name = "frm090202_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/16 Form2.0已修改
'Create by Sindy 2013/4/30
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim dblPrevRow As Double
Public m_ProState As String 'P,CFP,T,A(其他),FCP,FCT,CFT
'Modify By Sindy 2018/11/19 + 4.待轉檔區
Dim m_NPManKind As String '程序人員種類：1.台灣案 2.非台灣案 3.非台灣案歸檔 4.待轉檔區
                          '              空白.CFP 或 其他
'Add By Sindy 2018/12/25
'Modify By Sindy 2019/5/30 + 243.補書狀 244.補中文說明書
'Modified by Morgan 2021/2/24 +417提早公開 --蕭茹曣
'Modified by Morgan 2024/11/18 +447再審查加速審查
'Modified by Morgan 2025/4/23 -(213,226,243,410,501-804);+(245,415,433,446,933) --韻丞
Const 要進入轉檔區的案件性質 = "'101','102','103','104','105','107','117','125'" & _
                               ",'202','203','204','205','206','207','219','227','239','244','245'" & _
                               ",'301','302','303','304','305','306','307','308','309'" & _
                               ",'401','402','404','407','409','414','415','417','421','422','425','431','433','434','446','447'" & _
                               ",'807','933'"


'明細資料
Private Sub cmdDetail_Click()
Dim i As Integer
'Dim rsA As New ADODB.Recordset
'Dim stFileName As String, strCP10 As String, stFileTime As String
   
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 0) = "V" Then
         cmdDetail.Enabled = False 'Add By Sindy 2020/12/1
         'Add By Sindy 2017/1/9 為防止使用者明細畫面未關閉又從Menu再進入主作業畫面操作(因此下層明細先Unload再開啟)
         If TypeName(frm090202_4_1) <> "Nothing" Then
            Unload frm090202_4_1
         End If
         '2017/1/9 END
         
         'Add By Sindy 2020/9/30
         If GRD1.TextMatrix(i, 1) = "多案" Then
'            strSql = "select cpp01,cpp10,cp10 from casepaperpdf,caseprogress" & _
'                     " where cpp01='" & GRD1.TextMatrix(i, 22) & "' and cpp12='S'" & _
'                     " and instr(upper(cpp02),upper('" & EMP_多案承辦單 & "'))>0" & _
'                     " and cpp01=cp09(+) and cp27 is not null"
'            If rsA.State = adStateOpen Then rsA.Close
'            rsA.CursorLocation = adUseClient
'            rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'            If rsA.RecordCount = 0 Then
'               MsgBox "主歷程尚未發文歸卷，不可操作此案發文！", vbExclamation
'               Exit Sub
'            Else
'               strCP10 = rsA.Fields("cp10")
'               '承辦單檔案名稱
'               Call PUB_ChkEmpFlowFNMRule(GRD1.TextMatrix(i, 3), "", "Y", strCP10, stFileName, , False)
'               stFileName = stFileName & "." & strCP10 & "." & EMP_多案承辦單
'               stFileName = stFileName & ".menu"
'               stFileTime = Right("000000" & ServerTime, 6)
'               '以防重覆歸卷
'               strSql = "delete from CasePaperPDF where cpp01='" & GRD1.TextMatrix(i, 12) & "' and cpp02='" & stFileName & "'"
'               cnnConnection.Execute strSql
'               '新增一筆承辦單.menu至卷宗區
'               strSql = "insert into casepaperpdf(cpp01,cpp02,cpp03,CPP05,CPP06,CPP07,cpp08,cpp09,cpp10,cpp12)" & _
'                        " values('" & GRD1.TextMatrix(i, 12) & "'," & _
'                                "'" & stFileName & "',0,'" & strUserNum & "'," & _
'                                strSrvDate(1) & "," & stFileTime & "," & _
'                                strSrvDate(1) & "," & stFileTime & ",'Y','S')"
'               cnnConnection.Execute strSql, intI
'
'               '直接進發文作業
'               Call mdiMain.frm090202_4CallFrm(m_ProState, GRD1.TextMatrix(i, 23), GRD1.TextMatrix(i, 18), GRD1.TextMatrix(i, 19), GRD1.TextMatrix(i, 20), GRD1.TextMatrix(i, 21), GRD1.TextMatrix(i, 12))
'            End If
'            rsA.Close
            If PUB_InsChkWrkSht(GRD1.TextMatrix(i, 12), , GRD1.TextMatrix(i, 22)) = True Then
               '直接進發文作業
               Call mdiMain.frm090202_4CallFrm(m_ProState, GRD1.TextMatrix(i, 23), GRD1.TextMatrix(i, 18), GRD1.TextMatrix(i, 19), GRD1.TextMatrix(i, 20), GRD1.TextMatrix(i, 21), GRD1.TextMatrix(i, 12))
            End If
         Else
         '2020/9/30 END
            frm090202_4_1.Hide
            frm090202_4_1.m_EEP01 = GRD1.TextMatrix(i, 12) '總收文號
            frm090202_4_1.m_AttEEP02 = GRD1.TextMatrix(i, 13) '序號 Add By Sindy 2017/8/14
            'Add By Sindy 2025/1/15
            If Frame1.Visible = True And InStr(Label1.Caption, "人員") > 0 Then
               frm090202_4_1.m_FlowUserNum = Left(Combo1, 5)
            ElseIf Combo2.Visible = True Then
               frm090202_4_1.m_FlowUserNum = Left(Combo2, 5)
            End If
            '2025/1/15 END
            frm090202_4_1.SetParent Me
            'Modify By Sindy 2018/5/2
            frm090202_4_1.m_ProState = Me.m_ProState
   '         'Modify By Sindy 2018/4/27
   '         If Left(grd1.TextMatrix(i, 3), 3) = "CFP" Then
   '            Set frm090202_4_1.m_SendRecvForm = frm050102_1
   '         ElseIf Left(grd1.TextMatrix(i, 3), 1) = "P" Then
   '            Set frm090202_4_1.m_SendRecvForm = frm040104_1
   '         End If
   '         '2018/4/27 END
            '2018/5/2 END
            frm090202_4_1.m_NPManKind = m_NPManKind
            If frm090202_4_1.QueryData = True Then
               frm090202_4_1.Show
               Me.Hide
            End If
            Exit For
         End If
      End If
   Next i
   cmdDetail.Enabled = True 'Add By Sindy 2020/12/1
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

'Add By Sindy 2015/9/3 新增歷程
'Modify By Sindy 2024/6/28 改用物件:frmTmp
Private Sub cmdNewFlow_Click()
Dim frmTmp As Form
   
'   frm090202_7.SetParent Me
'   frm090202_7.Show
   Set frmTmp = Forms(0).GetForm("frm090202_7")
   frmTmp.SetParent Me
   frmTmp.Show
   Me.Hide
   
   Set frmTmp = Nothing
End Sub

'畫面更新
Private Sub cmdQuery_Click()
   If QueryData = False Then ShowNoData
End Sub

Public Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strConSql As String
Dim strConCP01_P As String, strConSql_M As String
Dim strConCP01_PS As String, strConSql_S As String
Dim i As Integer
Dim strEEP01 As String, strEEP02 As String
   
   m_blnColOrderAsc = True
   QueryData = True
   GRD1.Clear
   SetGrd
   strConSql = "": strConSql_M = "": strConSql_S = ""
   
'   m_NPManKind = ""
   If m_ProState = "CFP" Then
      strConCP01_P = " and CP01='CFP'"
      strConCP01_PS = " and CP01='CPS'" 'Add By Sindy 2015/10/21 +服務
      If Trim(Combo1.Text) <> "" Then
         strConSql = " And (EEP05='" & Left(Combo1, 5) & "' or EEP05 is null)"
      End If
   ElseIf m_ProState = "P" Then
      strConCP01_P = " and CP01='P'"
      strConCP01_PS = " and CP01='PS'" 'Add By Sindy 2015/10/21 +服務
      'Add By Sindy 2018/11/19 + 4.待轉檔區
      If Left(Combo1, 1) = "1" Or Left(Combo1, 1) = "4" Then '台灣案
'         m_NPManKind = "1"
         'Modify By Sindy 2013/11/13 含非台灣C類來函
         'strConSql = " And PA09='000'"
         'Modify By Sindy 2013/11/14 非台灣的分析也歸台灣送件,但因它有可能是A或B類收文
         'Modify By Sindy 2013/12/27 玲玲說901.告知代理人也歸台灣送件
         'Modify by Morgan 2016/8/1 +903專利調查,927其他翻譯--玲玲 ExP-115116
         'Modify by Morgan 2019/3/25 + (956)報告客戶,(1209)檢索報告(C類原來就有)--玲玲 Ex:P-120034
         'Modified by Morgan 2021/12/15 +P935 案件轉至本所 --玲玲 Ex:P-128768
         strConSql_M = " And (PA09='000' or (PA09||''<>'000' and (cp10||'' in('941','901','903','927','935','956') or substr(cp09,1,1)||''>='C')))"
         strConSql_S = " And (SP09='000' or (SP09||''<>'000' and (cp10||'' in('941','901','903','927','956') or substr(cp09,1,1)||''>='C')))"
         'Add By Sindy 2018/11/19
         If Left(Combo1, 1) = "4" Then '待轉檔區
            strConSql = strConSql & " And cp118 is not null and nvl(cp160,0)=0 And cp10 in(" & 要進入轉檔區的案件性質 & ")"
            'Added by Morgan 2025/4/23 +判斷承辦人為程序都不進4號區
            strConSql = strConSql & " and not exists(select * from staff where st01=cp14 and st03='P12')"
         Else
            'Modified by Morgan 2025/4/23 +判斷承辦人為程序直接進1號區
            strConSql = strConSql & " And (cp118 is null or cp10 not in(" & 要進入轉檔區的案件性質 & ") or (cp118 is not null and nvl(cp160,0)>0) or exists(select * from staff where st01=cp14 and st03='P12'))"
         End If
         '2018/11/19 END
         
         
      ElseIf Left(Combo1, 1) = "2" Or Left(Combo1, 1) = "3" Then '非台灣案
'         m_NPManKind = Left(Combo1, 1)
         'Modify By Sindy 2013/11/13 不含C類來函
         'strConSql = " And PA09<>'000'"
         'Modify By Sindy 2013/11/14 排除分析
         'Modify By Sindy 2013/12/27 玲玲說排除901.告知代理人
         'Modify by Morgan 2016/8/1 +排除903專利調查,927其他翻譯--玲玲 ExP-115116
         'Modify by Morgan 2019/3/25 +排除(956)報告客戶,(1209)檢索報告(C類原來就有)--玲玲 ExP-120034
         'Modified by Morgan 2021/7/29 +927其他翻譯改台灣非台灣都要顯示--玲玲 Ex:P-109152
         'Modified by Morgan 2021/12/15 -P935 案件轉至本所 --玲玲 Ex:P-128768
         strConSql_M = " And PA09||''<>'000' and cp09||''<'C' and cp10||'' not in('941','901','903','935','956')"
         strConSql_S = " And SP09||''<>'000' and cp09||''<'C' and cp10||'' not in('941','901','903','956')"
         If Left(Combo1, 1) = "2" Then '非台灣案
            strConSql = strConSql & " And eep01=(select max(smb01) from smailbackup where smb01=eep01)"
            '原讀取歷程附件檔改判斷卷宗區
            'strConSql = strConSql & " And eep1.eep01=(select eef01 from EmpElectronFile where eef01=eep1.eep01 and eef02=eep1.eep02 and instr(upper(eef03),upper('." & EMP_承辦單 & "'))>0)"
            'Modify By Sindy 2020/9/29 + instr(upper(cpp02),upper('." & EMP_多案承辦單 & "'))>0)
            strConSql = strConSql & " And eep01=(select cpp01 from casepaperpdf where cpp01=eep01 and (instr(upper(cpp02),upper('." & EMP_承辦單 & "'))>0 or instr(upper(cpp02),upper('." & EMP_多案承辦單 & "'))>0))"
         End If
         
      End If
      
      'Added by Morgan 2025/1/14
      If Trim(Combo2.Text) <> "" Then
         strConSql = strConSql & " And (EEP05='" & Left(Combo2, 5) & "' or EEP05 is null)"
      End If
      'end 2025/1/14
      
   'Add By Sindy 2021/10/13 顧服人員選單
   ElseIf Left(PUB_GetStaffST15(strUserNum, "1"), 2) = "W2" Then
      strConSql = strConSql & " And CP01='ACS'"
      If Trim(Combo1.Text) <> "" Then
         strConSql = strConSql & " And (EEP05='" & Left(Combo1, 5) & "' or EEP05 is null)"
      End If
   'Add By Sindy 2023/11/3
   'Modify By Sindy 2024/8/14 + Or m_ProState = "FCT" Or m_ProState = "CFT"
   ElseIf m_ProState = "FCP" Or m_ProState = "FCT" Or m_ProState = "CFT" Then
      'Modify By Sindy 2025/1/24
      'If Combo1.Visible = True And Combo1.Enabled = True Then
      If Frame1.Enabled = True And Trim(Combo1.Text) <> "" Then
         strConSql = " And EEP05='" & Left(Combo1, 5) & "'"
      Else
      '2025/1/24 END
         strConSql = " And EEP05='" & strUserNum & "'"
      End If
      'Add By Sindy 2025/1/22
      If m_ProState = "FCT" Then
         strConSql_M = " And TM01='FCT'"
         strConSql_S = " And (SP01='S' And SP09='000')"
      ElseIf m_ProState = "CFT" Then
         strConSql_M = " And TM01='CFT'"
         strConSql_S = " And ((SP01='S' And SP09<>'000') or SP01='CFC')"
      End If
      '2025/1/22 END
   'Add By Sindy 2018/4/26
   'Modify By Sindy 2021/9/1
   'ElseIf m_ProState = "T" Then
   Else
      'strConSql = " And EEP05='" & Pub_GetSpecMan("內商發文人員") & "'"
      strConSql = " And st03='" & PUB_GetST03(strUserNum) & "'"
      'st01
   '2021/9/1 END
   '2018/4/26 END
   End If
   m_NPManKind = Left(Combo1, 1)
   Screen.MousePointer = vbHourglass
   
   'Add By Sindy 2018/4/26
   'Modify By Sindy 2024/8/14 + Or m_ProState = "FCT" Or m_ProState = "CFT"
   If m_ProState = "T" Or m_ProState = "FCT" Or m_ProState = "CFT" Then
      'Modify By Sindy 2020/9/30 + cp163,pa11
      'Modify By Sindy 2024/1/18 +,cp10,cp14,pa09
      strSql = "Select ' ' as V,decode(cp163,null,SqlDateT(EEP06),decode(cp163,cp09,SqlDateT(EEP06),'多案')) as 判發日期,SqlDateT(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,TM05||TM06||TM07 as 案件名稱,NA03 as 國家,TM09 as 類別,Decode(pa09,'000',CPM03,CPM04) as 案件性質,SqlDateT (cp06) as 本所期限,SqlDateT (cp48) as 承辦期限, s1.ST02 as 承辦人, s2.ST02 as 智權人員,cp09 As 總收文號, EEP02 As 序號,'' as 狀態,EEP05,CP118,EEP06,cp01,cp02,cp03,cp04,cp163,pa11,cp10,cp14,pa09" & _
               " from (select eep01,eep02,eep04,eep05,eep06,eep13,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp13,cp14,cp118,TM01,TM02,TM03,TM04,TM05,TM06,TM07,TM08,TM10 as pa09,Decode(TM10,'000',PTM03,PTM04) as PTM03,TM09,cp163,tm12 as pa11,cp48" & _
                       " from EmpElectronProcess,CaseProgress,Trademark,PatentTradeMarkMap,staff" & _
                       " where eep01=cp09(+) and cp158=0 and cp159=0" & _
                       " and eep01 is not null and eep04 in('" & EMP_送件 & "','" & EMP_退件重送 & "','" & EMP_程序退回 & "')" & _
                       " and EEP13='Y' and CP01=TM01(+) And CP02=TM02(+) And CP03=TM03(+) And CP04=TM04(+) And TM01 is not null And EEP05=st01(+)" & strConSql & _
                       " And '2'=PTM01 AND TM08=PTM02(+)" & strConSql_M & _
               " union select eep01,eep02,eep04,eep05,eep06,eep13,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp13,cp14,cp118,SP01,SP02,SP03,SP04,SP05,SP06,SP07,'',SP09 as pa09,'' as PTM03,'' as TM09,cp163,sp11 as pa11,cp48" & _
                       " from EmpElectronProcess,CaseProgress,Servicepractice,staff" & _
                       " where eep01=cp09(+) and cp158=0 and cp159=0" & _
                       " and eep01 is not null and eep04 in('" & EMP_送件 & "','" & EMP_退件重送 & "','" & EMP_程序退回 & "')" & _
                       " and EEP13='Y' and CP01=SP01(+) And CP02=SP02(+) And CP03=SP03(+) And CP04=SP04(+) And SP01 is not null And EEP05=st01(+)" & strConSql & strConSql_S
      'Modify By Sindy 2020/9/30 + 讀取待送件的多案歷程總收文號
      strSql = strSql & " union select eep01,eep02,eep04,eep05,eep06,eep13,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp13,cp14,cp118,TM01,TM02,TM03,TM04,TM05,TM06,TM07,TM08,TM10 as pa09,Decode(TM10,'000',PTM03,PTM04) as PTM03,TM09,cp163,tm12 as pa11,cp48" & _
                       " from EmpElectronProcess,CaseProgress,Trademark,PatentTradeMarkMap" & _
                       " where cp163 in(SELECT eep01 FROM EmpElectronProcess,staff WHERE eep04 IN('" & EMP_送件 & "','" & EMP_退件重送 & "','" & EMP_程序退回 & "') AND EEP13='Y' AND eep15 IS NOT NULL And EEP05=st01(+)" & strConSql & ")" & _
                       " and eep01(+)=cp163 and cp09<>cp163 and cp158=0 and cp159=0" & _
                       " and eep01 is not null and eep04 in('" & EMP_送件 & "','" & EMP_退件重送 & "','" & EMP_程序退回 & "')" & _
                       " and EEP13='Y' and CP01=TM01(+) And CP02=TM02(+) And CP03=TM03(+) And CP04=TM04(+) And TM01 is not null" & " And '2'=PTM01 AND TM08=PTM02(+)" & strConSql_M & _
               " union select eep01,eep02,eep04,eep05,eep06,eep13,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp13,cp14,cp118,SP01,SP02,SP03,SP04,SP05,SP06,SP07,'',SP09 as pa09,'' as PTM03,'' as TM09,cp163,sp11 as pa11,cp48" & _
                       " from EmpElectronProcess,CaseProgress,Servicepractice" & _
                       " where cp163 in(SELECT eep01 FROM EmpElectronProcess,staff WHERE eep04 IN('" & EMP_送件 & "','" & EMP_退件重送 & "','" & EMP_程序退回 & "') AND EEP13='Y' AND eep15 IS NOT NULL And EEP05=st01(+)" & strConSql & ")" & _
                       " and eep01(+)=cp163 and cp09<>cp163 and cp158=0 and cp159=0" & _
                       " and eep01 is not null and eep04 in('" & EMP_送件 & "','" & EMP_退件重送 & "','" & EMP_程序退回 & "')" & _
                       " and EEP13='Y' and CP01=SP01(+) And CP02=SP02(+) And CP03=SP03(+) And CP04=SP04(+) And SP01 is not null" & strConSql_S
      strSql = strSql & " union select eep01,eep02,eep04,eep05,eep06,eep13,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp13,cp14,cp118,TM01,TM02,TM03,TM04,TM05,TM06,TM07,TM08,TM10 as pa09,Decode(TM10,'000',PTM03,PTM04) as PTM03,TM09,cp163,tm12 as pa11,cp48" & _
                       " from EmpElectronProcess,CaseProgress,Trademark,PatentTradeMarkMap" & _
                       " where cp163 IS NOT NULL" & _
                       " and eep01(+)=cp163 and eep02=(SELECT max(eep02) FROM EmpElectronProcess,staff WHERE eep01=cp163 and eep04 IN('" & EMP_送件 & "','" & EMP_退件重送 & "','" & EMP_程序退回 & "') AND eep15 IS NOT NULL And EEP05=st01(+)" & strConSql & ") and cp09<>cp163 and cp158=0 and cp159=0" & _
                       " AND exists (select * from casepaperpdf where cpp01(+)=cp163 and instr(upper(cpp02),upper('" & EMP_多案承辦單 & "'))>0)" & _
                       " and CP01=TM01(+) And CP02=TM02(+) And CP03=TM03(+) And CP04=TM04(+) And TM01 is not null" & " And '2'=PTM01 AND TM08=PTM02(+)" & strConSql_M & _
               " union select eep01,eep02,eep04,eep05,eep06,eep13,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp13,cp14,cp118,SP01,SP02,SP03,SP04,SP05,SP06,SP07,'',SP09 as pa09,'' as PTM03,'' as TM09,cp163,sp11 as pa11,cp48" & _
                       " from EmpElectronProcess,CaseProgress,Servicepractice" & _
                       " where cp163 IS NOT NULL" & _
                       " and eep01(+)=cp163 and eep02=(SELECT max(eep02) FROM EmpElectronProcess,staff WHERE eep01=cp163 and eep04 IN('" & EMP_送件 & "','" & EMP_退件重送 & "','" & EMP_程序退回 & "') AND eep15 IS NOT NULL And EEP05=st01(+)" & strConSql & ") and cp09<>cp163 and cp158=0 and cp159=0" & _
                       " AND exists (select * from casepaperpdf where cpp01(+)=cp163 and instr(upper(cpp02),upper('" & EMP_多案承辦單 & "'))>0)" & _
                       " and CP01=SP01(+) And CP02=SP02(+) And CP03=SP03(+) And CP04=SP04(+) And SP01 is not null" & strConSql_S
      strSql = strSql & ") A" & _
               ",staff s1,staff s2,nation,CasePropertyMap,ENGINEERPROGRESS" & _
               " where cp09=ep02(+) And ep05=s1.ST01(+) And CP13=s2.ST01(+)" & _
               " And PA09=NA01(+)" & _
               " And CP01=CPM01(+) And CP10=CPM02(+)" '& _
               " order by eep06,cp163,cp01,cp02,cp03,cp04 asc"
      'Modify By Sindy 2025/9/10
      If m_ProState = "FCT" Then
         strSql = strSql & " order by eep06 desc,cp163,cp01,cp02,cp03,cp04"
      Else
      '2025/9/10 END
         strSql = strSql & " order by eep06,cp163,cp01,cp02,cp03,cp04 asc"
      End If
      '2025/9/10 END
   'Add By Sindy 2021/9/1 其他案件
   ElseIf m_ProState = "A" Then
      strSql = "Select ' ' as V,SqlDateT(EEP06) as 判發日期,SqlDateT(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,HC06 as 案件名稱,NA03 as 國家,'' as 種類,Decode(PA09,'000',CPM03,CPM04) as 案件性質,SqlDateT (cp06) as 本所期限,SqlDateT (cp48) as 承辦期限, s1.ST02 as 承辦人, s2.ST02 as 智權人員,EEP01 As 總收文號, EEP02 As 序號,'' as 狀態,EEP05,CP118,EEP06,cp01,cp02,cp03,cp04,cp163,pa11,cp10,cp14,pa09" & _
               " from (select eep01,eep02,eep04,eep05,eep06,eep13,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp13,cp14,cp118,LC01,LC02,LC03,LC04,LC05,LC05||LC06||LC07 as HC06,LC07,LC15 as PA09,cp163,'' as pa11,cp48" & _
                       " from EmpElectronProcess,CaseProgress,Lawcase,staff" & _
                       " where eep01=cp09(+) and cp158=0 and cp159=0" & _
                       " and eep01 is not null and eep04 in('" & EMP_判發 & "','" & EMP_退件重送 & "')" & _
                       " and EEP13='Y' and CP01=LC01(+) And CP02=LC02(+) And CP03=LC03(+) And CP04=LC04(+) And LC01 is not null And EEP05=st01(+)" & strConSql & _
               " union select eep01,eep02,eep04,eep05,eep06,eep13,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp13,cp14,cp118,HC01,HC02,HC03,HC04,HC05,HC06,HC07,'000' as PA09,cp163,'' as pa11,cp48" & _
                       " from EmpElectronProcess,CaseProgress,Hirecase,staff" & _
                       " where eep01=cp09(+) and cp158=0 and cp159=0" & _
                       " and eep01 is not null and eep04 in('" & EMP_判發 & "','" & EMP_退件重送 & "')" & _
                       " and EEP13='Y' and CP01=HC01(+) And CP02=HC02(+) And CP03=HC03(+) And CP04=HC04(+) And HC01 is not null And EEP05=st01(+)" & strConSql & _
               " union select eep01,eep02,eep04,eep05,eep06,eep13,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp13,cp14,cp118,SP01,SP02,SP03,SP04,SP05,SP05||SP06||SP07 as HC06,SP07,SP09 as PA09,cp163,sp11 as pa11,cp48" & _
                       " from EmpElectronProcess,CaseProgress,Servicepractice,staff" & _
                       " where eep01=cp09(+) and cp158=0 and cp159=0" & _
                       " and eep01 is not null and eep04 in('" & EMP_判發 & "','" & EMP_退件重送 & "')" & _
                       " and EEP13='Y' and CP01=SP01(+) And CP02=SP02(+) And CP03=SP03(+) And CP04=SP04(+) And SP01 is not null And EEP05=st01(+)" & strConSql & _
                     ") A" & _
               ",staff s1,staff s2,nation,CasePropertyMap" & _
               " where CP14=s1.ST01(+) And CP13=s2.ST01(+)" & _
               " And PA09=NA01(+)" & _
               " And CP01=CPM01(+) And CP10=CPM02(+)" & _
               " order by eep06,cp01,cp02,cp03,cp04 asc"
   
   'Add By Sindy 2023/11/3 +FCP
   ElseIf m_ProState = "FCP" Then
      strSql = "Select ' ' as V,SqlDateT(EEP06) as 判發日期,SqlDateT(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,PA05||PA06||PA07 as 案件名稱,NA03 as 國家,PTM03 as 種類,Decode(PA09,'000',CPM03,CPM04) as 案件性質,SqlDateT (cp06) as 本所期限,SqlDateT (cp48) as 承辦期限, s1.ST02 as 承辦人, s2.ST02 as 智權人員,EEP01 As 總收文號, EEP02 As 序號,decode(eep04,'" & EMP_程序退回 & "','退回','') as 狀態,EEP05,CP118,EEP06,cp01,cp02,cp03,cp04,cp163,pa11,cp10,cp14,pa09" & _
               " from (select eep01,eep02,eep04,eep05,eep06,eep13,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp13,cp14,cp118,PA01,PA02,PA03,PA04,PA05,PA06,PA07,PA08,PA09,Decode(PA09,'000',PTM03,PTM04) as PTM03,cp163,pa11,cp48" & _
                       " from EmpElectronProcess,CaseProgress,Patent,PatentTradeMarkMap,staff" & _
                       " where eep01=cp09(+) and cp158=0 and cp159=0" & _
                       " and eep01 is not null and eep04 in('" & EMP_送件 & "','" & EMP_退件重送 & "','" & EMP_程序退回 & "')" & _
                       " and EEP13='Y' and CP01=PA01(+) And CP02=PA02(+) And CP03=PA03(+) And CP04=PA04(+) And PA01 is not null And EEP05=st01(+)" & strConSql & _
                       " And '1'=PTM01 AND PA08=PTM02(+)" & _
               " union select eep01,eep02,eep04,eep05,eep06,eep13,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp13,cp14,cp118,SP01,SP02,SP03,SP04,SP05,SP06,SP07,'',SP09 as PA09,'' as PTM03,cp163,sp11 as pa11,cp48" & _
                       " from EmpElectronProcess,CaseProgress,Servicepractice,staff" & _
                       " where eep01=cp09(+) and cp158=0 and cp159=0" & _
                       " and eep01 is not null and eep04 in('" & EMP_送件 & "','" & EMP_退件重送 & "','" & EMP_程序退回 & "')" & _
                       " and EEP13='Y' and CP01=SP01(+) And CP02=SP02(+) And CP03=SP03(+) And CP04=SP04(+) And SP01 is not null And EEP05=st01(+)" & strConSql
      strSql = strSql & ") A" & _
               ",staff s1,staff s2,nation,CasePropertyMap" & _
               " where CP14=s1.ST01(+) And CP13=s2.ST01(+)" & _
               " And PA09=NA01(+)" & _
               " And CP01=CPM01(+) And CP10=CPM02(+)" & _
               " order by eep06,cp01,cp02,cp03,cp04 asc"
   
   'P,CFP
   Else
   '2018/4/26 END
      'Modify By Sindy 2020/9/30 + cp163,pa11
      'Modify By Sindy 2023/11/30 + EMP_送件,因FMP案; +and EEP05=ST01(+) and not(substr(ST03,1,1)='F' and eep04='" & EMP_送件 & "')
      'Modify By Sindy 2024/1/10 ex:P-126752判發時待送件區就出現了 SQL改為 and EEP05=ST01(+) and not(substr(ST03,1,1)='F' and (eep04='" & EMP_送件 & "' or eep04='" & EMP_判發 & "'))
      strSql = "Select ' ' as V,SqlDateT(EEP06) as 判發日期,SqlDateT(CP05) as 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,PA05||PA06||PA07 as 案件名稱,NA03 as 國家,PTM03 as 種類,Decode(PA09,'000',CPM03,CPM04) as 案件性質,SqlDateT (cp06) as 本所期限,SqlDateT (cp48) as 承辦期限, s1.ST02 as 承辦人, s2.ST02 as 智權人員,EEP01 As 總收文號, EEP02 As 序號,'' as 狀態,EEP05,CP118,EEP06,cp01,cp02,cp03,cp04,cp163,pa11,cp10,cp14,pa09" & _
               " from (select eep01,eep02,eep04,eep05,eep06,eep13,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp13,cp14,cp118,PA01,PA02,PA03,PA04,PA05,PA06,PA07,PA08,PA09,Decode(PA09,'000',PTM03,PTM04) as PTM03,cp163,pa11,cp48 from EmpElectronProcess,CaseProgress,Patent,PatentTradeMarkMap,staff" & _
                       " where eep01=cp09(+)" & strConCP01_P & " and cp158=0 and cp159=0" & _
                       " and eep01 is not null and eep04 in('" & EMP_判發 & "','" & EMP_送件 & "','" & EMP_退件重送 & "')" & _
                       " and EEP13='Y' and CP01=PA01(+) And CP02=PA02(+) And CP03=PA03(+) And CP04=PA04(+) And PA01 is not null" & strConSql & strConSql_M & " And '1'=PTM01 AND PA08=PTM02(+)" & _
                       " and EEP05=ST01(+) and not(substr(ST03,1,1)='F' and eep04 in('" & EMP_判發 & "','" & EMP_送件 & "','" & EMP_退件重送 & "'))" & _
               " union select eep01,eep02,eep04,eep05,eep06,eep13,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp13,cp14,cp118,SP01,SP02,SP03,SP04,SP05,SP06,SP07,'',SP09 as PA09,'' as PTM03,cp163,sp11 as pa11,cp48 from EmpElectronProcess,CaseProgress,Servicepractice,staff" & _
                       " where eep01=cp09(+)" & strConCP01_PS & " and cp158=0 and cp159=0" & _
                       " and eep01 is not null and eep04 in('" & EMP_判發 & "','" & EMP_送件 & "','" & EMP_退件重送 & "')" & _
                       " and EEP13='Y' and CP01=SP01(+) And CP02=SP02(+) And CP03=SP03(+) And CP04=SP04(+) And SP01 is not null" & strConSql & strConSql_S & _
                       " and EEP05=ST01(+) and not(substr(ST03,1,1)='F' and eep04 in('" & EMP_判發 & "','" & EMP_送件 & "','" & EMP_退件重送 & "'))" & _
                     ") A" & _
               ",staff s1,staff s2,nation,CasePropertyMap" & _
               " where CP14=s1.ST01(+) And CP13=s2.ST01(+)" & _
               " And PA09=NA01(+)" & _
               " And CP01=CPM01(+) And CP10=CPM02(+)" & _
               " order by eep06,cp01,cp02,cp03,cp04 asc"
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   Label3.Caption = "共 0 筆"
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
      Label3.Caption = "共 " & rsTmp.RecordCount & " 筆"
   Else
      QueryData = False
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Function
   End If
   
   'Add By Sindy 2013/9/5
   GRD1.Visible = False
   For i = 1 To GRD1.Rows - 1
      'Modify By Sindy 2014/5/20 逐筆檢查是否已有承辦單
      strEEP01 = GRD1.TextMatrix(i, 12)
      strEEP02 = GRD1.TextMatrix(i, 13)
      If m_ProState = "P" And Left(Combo1, 1) = "2" Then '非台灣案
         GRD1.TextMatrix(i, 14) = "Y"
      Else
         '原讀取歷程附件檔改判斷卷宗區
'         strExc(0) = "select eef03 from EmpElectronFile" & _
'                     " where eef01='" & strEEP01 & "' and eef02='" & strEEP02 & "'" & _
'                     " and instr(eef03,'." & EMP_承辦單 & "')>0"
         'Modify By Sindy 2020/9/29 + instr(upper(cpp02),upper('." & EMP_多案承辦單 & "'))>0
         strExc(0) = "select cpp02 from Casepaperpdf" & _
                     " where cpp01='" & strEEP01 & "'" & _
                     " and (instr(upper(cpp02),upper('." & EMP_承辦單 & "'))>0 or instr(upper(cpp02),upper('." & EMP_多案承辦單 & "'))>0)"
         intI = 1
         Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            GRD1.TextMatrix(i, 14) = "Y"
         End If
      End If
      '2014/5/20 END
      Call SetColColor(i)
   Next i
   GRD1.Visible = True
   
   dblPrevRow = 0
'   '若有資料游標停在第一筆
'   GRD1.Visible = False
'   GRD1.col = 0
'   GRD1.row = 1
'   dblPrevRow = GRD1.row
'   If GRD1.Rows - 1 = 1 And GRD1.TextMatrix(GRD1.row, 12) <> "" Then
'      GRD1.Text = "V"
'      For i = 0 To GRD1.Cols - 1
'         GRD1.col = i
'         GRD1.CellBackColor = &HFFC0C0
'      Next i
'   End If
'   GRD1.Visible = True
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Function

'Add By Sindy 2013/9/5
'已處理過電子承辦單時以紅色標註
'電子送件以淡綠色標註
Private Sub SetColColor(intRow As Integer)
Dim i As Integer
Dim intCF10Kind As Integer
Dim rsTmp As New ADODB.Recordset
Dim bolChkCF10 As Boolean
   
   GRD1.row = intRow
   If GRD1.TextMatrix(intRow, 14) = "Y" Then '已處理過
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = &H8080FF '紅色
      Next i
   'Modify By Sindy 2013/11/8 若電子送件已處理過時,至少案件性質欄顯示綠色才能辨識出是電子送件
   End If
   If Trim(GRD1.TextMatrix(intRow, 16)) <> "" Then '電子送件
      GRD1.col = 2
      If GRD1.CellBackColor = &H8080FF Then
         GRD1.col = 7 '案件性質欄位
         GRD1.CellBackColor = &HC0FFC0 '淡綠色
      Else
   '2013/11/8 END
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HC0FFC0 '淡綠色
         Next i
      End If
   End If
                         
   'Add By Sindy 2024/1/18 外專要檢查台灣案發文前是否有需要送判程序主管
   'Modify By Sindy 2024/8/14 +FCT
   bolChkCF10 = False
   If Trim(GRD1.TextMatrix(intRow, 26)) = "000" Then
      If m_ProState = "FCT" And PUB_GetST03(strUserNum) = "F12" Then
         strExc(0) = "select ep02,ep05 from engineerprogress" & _
                     " where ep02='" & GRD1.TextMatrix(intRow, 12) & "'"
         intI = 1
         Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If PUB_GetST03("" & rsTmp.Fields("ep05")) <> "F12" Then
               bolChkCF10 = True
            End If
         End If
      ElseIf m_ProState = "FCP" Then
         If PUB_GetST03(Trim(GRD1.TextMatrix(intRow, 25))) <> "F22" Then
            bolChkCF10 = True
         End If
      End If
      If bolChkCF10 = True Then
         intCF10Kind = PUB_ChkhadCF10forEMP_46(Trim(GRD1.TextMatrix(intRow, 18)), _
                       Trim(GRD1.TextMatrix(intRow, 26)), Trim(GRD1.TextMatrix(intRow, 24)), _
                       Trim(GRD1.TextMatrix(intRow, 12)), GRD1.TextMatrix(intRow, 13))
         If intCF10Kind = 2 Then '已程序判發回來
            GRD1.col = 10 '承辦人欄位
            GRD1.CellBackColor = &HFFFF& '黃色
         End If
      End If
   End If
   '2024/1/18 END
   
   Set rsTmp = Nothing
End Sub

'Add By Sindy 2015/9/3
Private Sub Combo1_Click()
   If m_ProState = "P" And Left(Combo1, 1) = "3" Then '非台灣案歸檔
      cmdNewFlow.Visible = True
   Else
      cmdNewFlow.Visible = False
   End If
   Call QueryData
End Sub

'Added by Morgan 2025/1/14
Private Sub Combo2_Click()
   Call QueryData
End Sub
'end 2025/1/14

Private Sub Form_Load()
Dim rsTmp As New ADODB.Recordset
Dim i As Integer
   
   MoveFormToCenter Me
   LblFCPNote.Visible = False 'Add By Sindy 2024/1/18
   
   cmdNewFlow.Visible = False
   Combo1.Clear
   'Modify by Sindy 2020/3/12 程序人員選單
   If m_ProState = "CFP" Then
      Call SetPatentP12Combo(Combo1, "CFP", Label1)
'      Label1.Caption = "程序人員："
'      strExc(0) = "select na73 from nation " & _
'                  "Union " & _
'                  "select na74 from nation " & _
'                  "order by 1 asc "
'      intI = 1
'      Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         rsTmp.MoveFirst
'         Combo1.AddItem ""
'         Do While Not rsTmp.EOF
'            Combo1.AddItem rsTmp.Fields(0) & " " & GetPrjSalesNM(rsTmp.Fields(0))
'            rsTmp.MoveNext
'         Loop
'      End If
'      Combo1.ListIndex = 1
'      For i = 0 To Combo1.ListCount - 1
'         If Trim(Left(Combo1.List(i), 6)) = strUserNum Then
'            Combo1.ListIndex = i
'            Exit For
'         End If
'      Next i
'      rsTmp.Close
'      Set rsTmp = Nothing
      '2020/3/12 END
   ElseIf m_ProState = "P" Then
      Label1.Caption = "性質：" '"國家："
      'Combo1.AddItem "" 'Removed by Morgan 2025/1/14 不應該有空白，否則抓出來的資料條件有少且進明細畫面的控制也會有問題
      Combo1.AddItem "1 台灣案"
      Combo1.AddItem "2 非台灣案"
      Combo1.AddItem "3 非台灣案歸檔" '品薇在操作使用
      Combo1.AddItem "4 待轉檔區" '4.待轉檔區 Add By Sindy 2018/11/19
'      Combo1.ListIndex = 1
'      'Modify By Sindy 2014/9/17
''      If Pub_GetSpecMan("PS1") = strUserNum Then '台灣案
''         Combo1.ListIndex = 1
''      Else
'      '2014/9/17 END

      'Added by Morgan 2025/1/13
      If strSrvDate(1) >= P業務區劃分啟用日 Then
         Combo2.Visible = True
         Label4.Visible = True
         Call SetPatentP12Combo(Combo2, "P", Label4)
         Combo1.ListIndex = 0
      Else
      'end 2025/1/13
      
         If Pub_GetSpecMan("PS1") = strUserNum Then '台灣案
            'Modified by Morgan 2025/1/14
            'Combo1.ListIndex = 1
            Combo1.ListIndex = 0
         ElseIf Pub_GetSpecMan("PS2") = strUserNum Then '非台灣案
            'Modified by Morgan 2025/1/14
            'Combo1.ListIndex = 2
            Combo1.ListIndex = 1
         Else
            'Modified by Morgan 2025/1/14
            'Combo1.ListIndex = 3
            Combo1.ListIndex = 2
         End If
         
      End If 'Added by Morgan 2025/1/13
   
   'Add By Sindy 2021/10/13 顧服人員選單
   'Add By Sindy 2023/11/3 +FCP程序
   'Modify By Sindy 2024/8/14 +FCT程序
   'Modify By Sindy 2025/1/22 (m_ProState = "FCT" And PUB_GetST03(strUserNum) = "F12") 改為 PUB_GetST03(strUserNum) = "F12"
   ElseIf Left(PUB_GetStaffST15(strUserNum, "1"), 2) = "W2" Or _
          m_ProState = "FCP" Or _
          PUB_GetST03(strUserNum) = "F12" Then
      'Modify By Sindy 2024/8/14
      If PUB_GetST03(strUserNum) = "F12" Then 'm_ProState = "FCT"
         LblFCPNote.Visible = True
         strExc(0) = "select * from staff where st15='F12' and st04='1'" & _
                     " order by 1 asc"
      'Add By Sindy 2023/11/3
      ElseIf m_ProState = "FCP" Then
         LblFCPNote.Visible = True 'Add By Sindy 2024/1/18
         strExc(0) = "select * from staff where st15='F22' and st04='1'" & _
                     " order by 1 asc"
      Else
      '2023/11/3 END
         strExc(0) = "select * from staff where substr(st15,1,2)='W2' and st04='1' and st01 not in('W2001')" & _
                     " order by 1 asc"
      End If
      intI = 1
      Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         rsTmp.MoveFirst
         If Left(PUB_GetStaffST15(strUserNum, "1"), 2) = "W2" Then Combo1.AddItem ""
         Do While Not rsTmp.EOF
            Combo1.AddItem rsTmp.Fields(0) & " " & GetPrjSalesNM(rsTmp.Fields(0))
            rsTmp.MoveNext
         Loop
      End If
      Combo1.ListIndex = 1
      For i = 0 To Combo1.ListCount - 1
         If Trim(Left(Combo1.List(i), 6)) = strUserNum Then
            Combo1.ListIndex = i
            Exit For
         End If
      Next i
      rsTmp.Close
      Set rsTmp = Nothing
      
   'Add By Sindy 2018/4/26
   'Modify By Sindy 2021/7/15 mark:If m_ProState = "T" Then
   Else 'If m_ProState = "T" Then
      'Add By Sindy 2024/10/28
      Dim pNoList As String, arrID As Variant
      If m_ProState = "FCT" Or m_ProState = "CFT" Then
         '檢查當時是否需要為他人職代
         Call Pub_SetForOthersEmpCombo(strUserNum, , False, pNoList)
         If Trim(pNoList) <> "" Then
            Combo1.AddItem strUserNum & " " & strUserName
            arrID = Split(pNoList, ";")
            For intI = 0 To UBound(arrID)
               If Trim(arrID(intI)) <> "" Then
                  Combo1.AddItem arrID(intI) & " " & GetPrjSalesNM(CStr(arrID(intI)))
               End If
            Next intI
            Combo1.Text = Combo1.List(0)
         Else
            Frame1.Visible = False
         End If
      Else
      '2024/10/28 END
         Frame1.Visible = False
      End If
      Call QueryData
   '2018/4/26 END
   End If
   
   'Call QueryData 'Modify By Sindy 2017/5/9 會呼叫到Combo1_Click函數, 裡面就會Run到QueryData, 所以此處Mark
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
   Set frm090202_4 = Nothing
'   If pub_CallNextABSForm = False Then
'      Call Forms(0).SysStartCallForm 'Add By Sindy 2011/10/7
'   End If
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modify By Sindy 2020/9/30 + cp163,pa11
   'Modify By Sindy 2024/12/5 FCT,CFT案件 +承辦期限
   arrGridHeadText = Array("V", "判發日期", "收文日", "本所案號", "案件名稱", _
                           "國家", "類別/種類", "案件性質", "本所期限", "承辦期限", "承辦人", _
                           "智權人員", "總收文號", "序號", "狀態", "EEP05", _
                           "CP118", "EEP06", "cp01", "cp02", "cp03", "cp04", "cp163", "pa11", _
                           "cp10", "cp14", "pa09")
   'Add By Sindy 2024/12/5 FCT,CFT案件
   If m_ProState = "FCT" Or m_ProState = "CFT" Then
      arrGridHeadWidth = Array(200, 780, 780, 1200, 1100, _
                               700, 0, 1000, 780, 780, 700, _
                               700, 0, 0, 600, 0, _
                               0, 0, 0, 0, 0, 0, 0, 0, _
                               0, 0, 0)
   'Add By Sindy 2021/9/1 其他案件
   ElseIf m_ProState = "A" Then
      arrGridHeadWidth = Array(200, 780, 780, 1200, 1100, _
                               700, 0, 1000, 780, 0, 700, _
                               700, 0, 0, 600, 0, _
                               0, 0, 0, 0, 0, 0, 0, 0, _
                               0, 0, 0)
   Else
   '2021/9/1 END
      arrGridHeadWidth = Array(200, 780, 780, 1200, 1100, _
                               700, 600, 1000, 780, 0, 650, _
                               650, 0, 0, 500, 0, _
                               0, 0, 0, 0, 0, 0, 0, 0, _
                               0, 0, 0)
   End If
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
Dim i As Integer

GRD1.Visible = False
If GRD1.MouseRow <> 0 Then
   '上一筆資料列清除反白
   'Modify By Sindy 2016/5/9
   'If dblPrevRow > 0 Then
   If dblPrevRow > 0 And dblPrevRow <= (GRD1.Rows - 1) Then
   '2016/5/9 END
      GRD1.col = 0
      GRD1.row = dblPrevRow
      GRD1.Text = ""
      For i = 0 To GRD1.Cols - 1
         'Add By Sindy 2024/8/19
         If i <> 7 And i <> 10 Then
         '2024/8/19 END
            GRD1.col = i
            GRD1.CellBackColor = QBColor(15)
         End If
      Next i
      Call SetColColor(CStr(dblPrevRow))
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
      If GRD1.TextMatrix(GRD1.row, 1) <> "" Then
         GRD1.Text = "V"
         For i = 0 To GRD1.Cols - 1
            'Add By Sindy 2024/8/19
            If i <> 7 And i <> 10 Then
            '2024/8/19 END
               GRD1.col = i
               GRD1.CellBackColor = &HFFC0C0
            End If
         Next i
      End If
'   End If
End If
GRD1.Visible = True
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
