VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm083006 
   BorderStyle     =   1  '單線固定
   Caption         =   "收 / 發文明細表"
   ClientHeight    =   2724
   ClientLeft      =   2856
   ClientTop       =   2076
   ClientWidth     =   4572
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2724
   ScaleWidth      =   4572
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   4
      Left            =   1500
      MaxLength       =   1
      TabIndex        =   6
      Text            =   "3"
      Top             =   2250
      Width           =   375
   End
   Begin VB.OptionButton Opt1 
      Caption         =   "承辦人"
      Height          =   180
      Index           =   1
      Left            =   2400
      TabIndex        =   5
      Top             =   1950
      Width           =   855
   End
   Begin VB.OptionButton Opt1 
      Caption         =   "日期"
      Height          =   180
      Index           =   0
      Left            =   1500
      TabIndex        =   4
      Top             =   1950
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   1500
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   2700
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   1500
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1500
      MaxLength       =   1
      TabIndex        =   0
      Top             =   810
      Width           =   255
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   3615
      TabIndex        =   8
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2790
      TabIndex        =   7
      Top             =   120
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   2505
      Left            =   300
      TabIndex        =   15
      Top             =   2850
      Width           =   5775
      _ExtentX        =   10181
      _ExtentY        =   4424
      _Version        =   393216
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "系  統  別：           (1.FCL,2.CFL,3.全部)"
      Height          =   180
      Index           =   2
      Left            =   540
      TabIndex        =   14
      Top             =   2310
      Width           =   3000
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   2280
      TabIndex        =   13
      Top             =   1590
      Width           =   1410
   End
   Begin VB.Line Line1 
      X1              =   2340
      X2              =   2580
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "列  印  別："
      Height          =   180
      Index           =   3
      Left            =   540
      TabIndex        =   12
      Top             =   1950
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承  辦  人："
      Height          =   180
      Index           =   2
      Left            =   540
      TabIndex        =   11
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "日        期："
      Height          =   180
      Index           =   1
      Left            =   540
      TabIndex        =   10
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "報表方式：            (1.收文 2.發文)"
      Height          =   180
      Index           =   0
      Left            =   540
      TabIndex        =   9
      Top             =   840
      Width           =   2715
   End
End
Attribute VB_Name = "frm083006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/31 法務系統的工作點數分配功能先上線(110/9/1)
'Memo by Lydia 2020/04/09  因為2019增加ACS系統，但是遺失2015-2016寫的工作點數程式，所以比較二個版本的差異性進行修改。
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim PLeft(0 To 15) As Integer
Dim m_print As Integer
Dim m_totCFLCP18 As Double, m_totFCLCP18 As Double, m_totCP18 As Double
Dim m_totCFLCP113 As Double, m_totFCLCP113 As Double, m_totCP113 As Double  '2009/12/8 add by sonia
Dim Page As Integer, iPrint As Integer
'Add By Sindy 2010/5/4
Dim intCompRow As Integer
Dim dblPointTot As Double
Dim m_strTemp1 As String, m_strTemp3 As String
Dim dblSum As Double
'2010/5/4 End
Dim dblPointTotSub As Double, dblPointRow As Integer, i As Integer, j As Integer 'Add By Sindy 2012/2/3
Dim aX(1 To 5) As Integer 'Modified by Lydia 2015/05/08 +收據資料(位置設定)
Private Const FatDot As String = "0.000" 'Added by Lydia 2015/06/08 統一小數點(原本0.0)

Private Sub cmdBack_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim strTempName As String
   
   m_print = 0
   If Text1(0) = "" Then
      Text1(0).SetFocus
      MsgBox "報表方式不得為空值 !", vbCritical
      Exit Sub
   End If
   If ChkRange(Text1(1), Text1(2), "日期") = False Then Exit Sub
   If PUB_CheckKeyInDate(Me.Text1(1)) = -1 Then
      Me.Text1(1).SetFocus
      Text1_GotFocus 1
      Exit Sub
   End If
   If PUB_CheckKeyInDate(Me.Text1(2)) = -1 Then
      Me.Text1(2).SetFocus
      Text1_GotFocus 2
      Exit Sub
   End If
   '檢查承辦人
   If Me.Text1(3).Text <> "" Then
      If ClsPDGetStaffN(Text1(3), strTempName) Then
         Label2 = strTempName
      Else
         Label2 = ""
         Me.Text1(3).SetFocus
         Text1_GotFocus 3
         Exit Sub
      End If
   End If
   'Add By Sindy 2009/08/20
   If Text1(4).Visible = True And Text1(4) = "" Then
      MsgBox "系統別不得為空值 !", vbCritical
      Text1(4).SetFocus
      Exit Sub
   End If
   '2009/08/20 End
   
   Screen.MousePointer = 11
   GetPrintLeft
   PrintCase
   Screen.MousePointer = 0
   If m_print = 0 Then
      MsgBox "列印結束!", vbInformation
   End If
End Sub

Private Sub PrintCase()
                                              
Dim strCon1 As String, strCon2 As String
Dim strExSql As String
Dim strSelCol As String 'Add By Sindy 2009/07/20
Dim strPointSql As String, strEmp As String 'Add By Sindy 2010/5/4
Dim strA1k01List As String 'Added by Lydia 2016/12/13 記錄已列印點數的收款單

'Memo by Lydia 2020/04/09 Me.tag從mdiMain傳入：0-內法、1-外法、2-ACS創新業務
'內法-列印別by日期
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'收文日　收文號　本所案號　案件性質　IP案　當事人　　　點數　智權人員　承辦人員　協辦人員　備註
'                                                                              對造當事人
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'                                                                              合計 點數: XX
'
'內法-列印別by承辦人
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'承辦人　收文日　收文號　本所案號　當事人　　　案件性質　智權人員　費用　點數　法院案號　工時　IP案件　取消日　備註
'                                                                對造當事人                                                                                                            cp57
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'                                                                                   （by收文部門）  小計　　點數：xx


On Error GoTo ErrHand
   
   m_totCP18 = 0: m_totFCLCP18 = 0: m_totCFLCP18 = 0
   m_totCP113 = 0: m_totFCLCP113 = 0: m_totCFLCP113 = 0
   
   'Add By Sindy 2009/07/20
   '外法CP35改抓取CP18
   If Me.Tag = 1 Then '外法
      'Modify By Sindy 2010/5/10 點數改抓點數分配檔
      'strSelCol = "CP18"
      'Modified by Lydia 2015/05/08 + acc0k0,設別名 a07,a09,a02
      'Modified by Lydia 2015/06/02 +工作點數分配
      strSelCol = "'' as cp35,decode(p1.a1n01,null,decode(p2.a1n01,null,decode(p3.a1n01,null,decode(p4.a1n01,null,cp18,p4.a1n05),p3.a1n05),p2.a1n05),p1.a1n05) a07"
   Else '內法, ACS
      strSelCol = "CP35,decode(p1.a1n01,null,decode(p2.a1n01,null,decode(p3.a1n01,null,decode(p4.a1n01,null,cp18,p4.a1n05),p3.a1n05),p2.a1n05),p1.a1n05) a07"
   End If
   If Opt1(0).Value = True Then
      strCon2 = "CP05 as a09"
   Else
      strCon2 = "CP14 as a09"
   End If
   If Text1(0).Text = "1" Then '收文
      strCon1 = "SQLDATET(CP05) "
   Else
      strCon1 = "SQLDATET(CP27) "
   End If
   
   strExSql = ""
   strPointSql = "" 'Add By Sindy 2010/5/4
   '系統別
   If Me.Tag = 0 Then  '內法
      strExSql = " AND CP01 in('L','LA')"
   ElseIf Me.Tag = 1 Then '外法
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'Modify By Sindy 2009/08/20
      'strExSql = " AND CP01 in('FCL','CFL','LIN')"
      If Text1(4).Text = "1" Then
         strExSql = " AND CP01 in('FCL','LIN')"
      ElseIf Text1(4).Text = "2" Then
         strExSql = " AND CP01 in('CFL')"
      Else
         strExSql = " AND CP01 in('FCL','CFL','LIN')"
      End If
      '2009/08/20 End
   'add by sonia 2019/8/8 +ACS系統類別
   ElseIf Me.Tag = 2 Then
      strExSql = " AND CP01='ACS'"
   'end 2019/8/8
   End If
   
   '日期
   If Text1(1) = "" And Text1(2) <> "" Then
      strExSql = strExSql & " AND " & IIf(Trim(Text1(0).Text) = "1", "CP05", "CP27") & "<='" & ChangeTStringToWString(Text1(2)) + "' "
   ElseIf Text1(1) <> "" And Text1(2) <> "" Then
      strExSql = strExSql & " AND " & IIf(Trim(Text1(0).Text) = "1", "CP05", "CP27") & " BETWEEN '" & _
         ChangeTStringToWString(Text1(1)) + "' AND '" + ChangeTStringToWString(Text1(2)) + "' "
   End If
   strPointSql = strExSql 'Add By Sindy 2010/5/4
   '承辦人
   If Text1(3) <> "" Then
      strExSql = strExSql & " and CP14='" + Text1(3) + "' "
      'Modified by Lydia 2016/12/13 移到QueryPointData
      'strPointSql = strPointSql & " and a1n04='" + Text1(3) + "' " 'Add By Sindy 2010/5/4
   End If

   'Add By Sindy 2010/5/4 外法-承辦人報表
   'Modified by Lydia 2015/06/08 點數：統一改在模組處理分配點數明細
   'If Opt1(1).Value = True And Me.Tag = 1 Then
   If Opt1(1).Value = True Then
      Call QueryPointData(strPointSql)
   End If
   '2010/5/4 End
   
   strExc(0) = ""
   'Modify By Sindy 2010/3/9 增加對造當事人
   'Modify By Sindy 2010/5/10 日期報表, 點數改抓請款單
   If Opt1(0).Value = True Then '列印別:日期 'Memo by Lydia 2020/04/10 只抓未取消收文CP57 is null
      '2008/9/9 modify by sonia 外法列印時cp64改印fa05
      If Me.Tag = 0 Then '內法
         '顧問基本檔
        'Modified by Lydia 2015/06/08 +工作點數
        'Modified by Lydia 2016/12/13 CP40||CP41||CP42=>CP40||' '||CP41||' '||CP42 ; +CP60
        'Modified by Lydia 2020/04/09 重新整理 ;
        'Modified by Lydia 2020/04/10 CP06的後面+CP14
        'Modified by Lydia 2024/07/08 acc1n0分配點數會有一筆以上; ex.FCL-010987, FCL-010990
        'strExc(0) = "SELECT " & strCon1 & " as a01 , CP09,CP01||'-'||CP02||DECODE(CP03,'0','','-'||CP03)||DECODE(CP04,'00','','-'||CP04) a02," & _
            "DECODE(CP01||CP10,CPM01||CPM02,CPM03) a03,DECODE(HC05,CU01||CU02,NVL(CU04,NVL(CU05,CU06))) a04," & _
            "DECODE(CP13,S2.ST01,S2.ST02) a05,DECODE(CP14,S1.ST01,S1.ST02) a06," & _
            "DECODE(CP29,S3.ST01,S3.ST02) a08,CP64," & strCon2 & ",CP01,CP02,CP03,CP04,'' fa05,N1.na03 as N1na03,a0902,'' N2na03,'' as lc13,CP40||' '||CP41||' '||CP42 as cp404142,CP60,SQLDATET(CP06) CP06,CP14 " & _
            ",CP18,nvl(sum(p1.a1n05),0) b01,nvl(sum(p2.a1n05),0) c01 FROM STAFF S1,STAFF S2,STAFF S3,CASEPROGRESS,HIRECASE,CASEPROPERTYMAP,CUSTOMER,nation N1,acc090,acc1n0 p1,acc1n0 p2 " & _
            " WHERE CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP29=S3.ST01(+) AND " & _
            "(SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+)) AND " & _
            "CU10=N1.na01(+) AND CP12=a0901(+) AND p1.a1n02(+)='3' AND p1.a1n03(+)=cp09 and p2.a1n02(+)='2' AND p2.a1n03(+)=cp09 " & _
            "and (CP01=CPM01(+) AND CP10=CPM02(+)) AND CP57 IS NULL AND CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 AND CP09<'C' " & strExSql & _
            "group by " & strCon1 & " ," & _
            "CP09,CP01||'-'||CP02||DECODE(CP03,'0','','-'||CP03)||DECODE(CP04,'00','','-'||CP04),DECODE(CP01||CP10,CPM01||CPM02,CPM03)," & _
            "DECODE(HC05,CU01||CU02,NVL(CU04,NVL(CU05,CU06))),DECODE(CP13,S2.ST01,S2.ST02),DECODE(CP14,S1.ST01,S1.ST02)," & _
            "DECODE(CP29,S3.ST01,S3.ST02),CP64," & Left(strCon2, 4) & ",CP01,CP02,CP03,CP04,N1.na03,a0902,CP40||' '||CP41||' '||CP42,CP18,CP60, SQLDATET(CP06),CP14" & _
            " UNION "
        strExc(0) = "SELECT " & strCon1 & " as a01 , CP09,CP01||'-'||CP02||DECODE(CP03,'0','','-'||CP03)||DECODE(CP04,'00','','-'||CP04) a02," & _
            "DECODE(CP01||CP10,CPM01||CPM02,CPM03) a03,DECODE(HC05,CU01||CU02,NVL(CU04,NVL(CU05,CU06))) a04," & _
            "DECODE(CP13,S2.ST01,S2.ST02) a05,DECODE(CP14,S1.ST01,S1.ST02) a06," & _
            "DECODE(CP29,S3.ST01,S3.ST02) a08,CP64," & strCon2 & ",CP01,CP02,CP03,CP04,'' fa05,N1.na03 as N1na03,a0902,'' N2na03,'' as lc13,CP40||' '||CP41||' '||CP42 as cp404142,CP60,SQLDATET(CP06) CP06,CP14 " & _
            ",CP18,nvl(v103,0) b01,nvl(v203,0) c01,nvl(v303,0) t01 FROM STAFF S1,STAFF S2,STAFF S3,CASEPROGRESS,HIRECASE,CASEPROPERTYMAP,CUSTOMER,nation N1,acc090," & _
            "(select a1n03 as v101,a1n04 as v102,sum(a1n05) as v103 from caseprogress,acc1n0 where cp09=a1n03(+) and a1n02='3' " & strExSql & " group by a1n03,a1n04) vt1," & _
            "(select a1n03 as v201,a1n04 as v202,sum(a1n05) as v203 from caseprogress,acc1n0 where cp09=a1n03(+) and a1n02='2' " & strExSql & " group by a1n03,a1n04) vt2, " & _
            "(select cp60 as v301,a1n04 as v302,sum(a1n05) as v303,min(cp09) mcp09 from caseprogress,acc1n0 where cp09=a1n03(+) and a1n02='3' " & strExSql & " group by cp60,a1n04 having count(*) > 1) vt3 " & _
            " WHERE CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP29=S3.ST01(+) AND " & _
            "(SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+)) AND " & _
            "CU10=N1.na01(+) AND CP12=a0901(+) and cp09=v101(+) and cp14=v102(+) and cp09=v201(+) and cp14=v202(+) and cp60=v301(+) and cp14=v302(+) and cp09=mcp09(+) " & _
            "and (CP01=CPM01(+) AND CP10=CPM02(+)) AND CP57 IS NULL AND CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 AND CP09<'C' " & strExSql & _
            " UNION "
      End If
      '法務基本檔
      '2009/12/8 modify by sonia 加是否為智慧財產權案LC13
      'Modify by Morgan 2011/6/1 若有建點數分配資料時點數改分配點數(目前L會有分配) cp18->nvl(a0n03/1000,cp18)
      'Modified by Lydia 2015/05/08 strCon2 =>改別名 a09
      'Modified by Lydia 2015/06/08 +工作點數
      'Modified by Lydia 2016/12/13 CP40||CP41||CP42=>CP40||' '||CP41||' '||CP42 ; +CP60
      'modify by sonia 2019/8/8 +CP06
      'Modified by Lydia 2020/04/09 重新整理
      'Modified by Lydia 2020/04/10 CP06的後面+CP14
      'Modified by Lydia 2024/07/08 acc1n0分配點數會有一筆以上; ex.FCL-010987, FCL-010990
      'strExc(0) = strExc(0) & "SELECT " & strCon1 & " as a01, CP09,CP01||'-'||CP02||DECODE(CP03,'0','','-'||CP03)||DECODE(CP04,'00','','-'||CP04) a02," & _
         "DECODE(CP01||CP10,CPM01||CPM02,CPM03) a03,DECODE(LC11,CU01||CU02,NVL(CU04,NVL(CU05,CU06))) a04," & _
         "DECODE(CP13,S2.ST01,S2.ST02) a05,DECODE(CP14,S1.ST01,S1.ST02) a06," & _
         "DECODE(CP29,S3.ST01,S3.ST02) a08,CP64," & strCon2 & ",CP01,CP02,CP03,CP04,nvl(nvl(fa05,fa04),fa06) fa05,N1.na03 as N1na03,a0902,N2.na03 as N2na03,lc13,CP40||' '||CP41||' '||CP42 as cp404142,CP60,SQLDATET(CP06) CP06,CP14 " & _
         ",CP18,nvl(sum(p1.a1n05),0) b01,nvl(sum(p2.a1n05),0) c01 FROM STAFF S1,STAFF S2,STAFF S3,CASEPROGRESS,LAWCASE,CASEPROPERTYMAP,CUSTOMER,fagent,nation N1,nation N2,acc090,acc1n0 p1,acc1n0 p2 " & _
         " where CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP29=S3.ST01(+) AND " & _
         "(SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+)) and substr(lc22,1,8)=fa01(+) and substr(lc22,9,1)=fa02(+) AND " & _
         "CU10=N1.na01(+) AND CP12=a0901(+) AND FA10=N2.na01(+) AND p1.a1n02(+)='3' AND p1.a1n03(+)=cp09 and p2.a1n02(+)='2' AND p2.a1n03(+)=cp09 " & _
         "and (CP01=CPM01(+) AND CP10=CPM02(+)) AND CP57 IS NULL AND CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 AND CP09<'C' " & strExSql & _
         " group by " & strCon1 & " ," & _
         "CP09,CP01||'-'||CP02||DECODE(CP03,'0','','-'||CP03)||DECODE(CP04,'00','','-'||CP04) ," & _
         "DECODE(CP01||CP10,CPM01||CPM02,CPM03),DECODE(LC11,CU01||CU02,NVL(CU04,NVL(CU05,CU06))) ," & _
         "DECODE(CP13,S2.ST01,S2.ST02),DECODE(CP14,S1.ST01,S1.ST02) ," & _
         "DECODE(CP29,S3.ST01,S3.ST02) ,CP64," & Left(strCon2, 4) & ",CP01,CP02,CP03,CP04,nvl(nvl(fa05,fa04),fa06) ,N1.na03 ,a0902,N2.na03 ,lc13,CP40||' '||CP41||' '||CP42,CP60,SQLDATET(CP06),CP14 " & _
         ",CP18"
      strExc(0) = strExc(0) & "SELECT " & strCon1 & " as a01, CP09,CP01||'-'||CP02||DECODE(CP03,'0','','-'||CP03)||DECODE(CP04,'00','','-'||CP04) a02," & _
         "DECODE(CP01||CP10,CPM01||CPM02,CPM03) a03,DECODE(LC11,CU01||CU02,NVL(CU04,NVL(CU05,CU06))) a04," & _
         "DECODE(CP13,S2.ST01,S2.ST02) a05,DECODE(CP14,S1.ST01,S1.ST02) a06," & _
         "DECODE(CP29,S3.ST01,S3.ST02) a08,CP64," & strCon2 & ",CP01,CP02,CP03,CP04,nvl(nvl(fa05,fa04),fa06) fa05,N1.na03 as N1na03,a0902,N2.na03 as N2na03,lc13,CP40||' '||CP41||' '||CP42 as cp404142,CP60,SQLDATET(CP06) CP06,CP14 " & _
         ",CP18,nvl(v103,0) b01,nvl(v203,0) c01,nvl(v303,0) t01 FROM STAFF S1,STAFF S2,STAFF S3,CASEPROGRESS,LAWCASE,CASEPROPERTYMAP,CUSTOMER,fagent,nation N1,nation N2,acc090," & _
         "(select a1n03 as v101,a1n04 as v102,sum(a1n05) as v103 from caseprogress,acc1n0 where cp09=a1n03(+) and a1n02='3' " & strExSql & " group by a1n03,a1n04) vt1," & _
         "(select a1n03 as v201,a1n04 as v202,sum(a1n05) as v203 from caseprogress,acc1n0 where cp09=a1n03(+) and a1n02='2' " & strExSql & " group by a1n03,a1n04) vt2, " & _
            "(select cp60 as v301,a1n04 as v302,sum(a1n05) as v303,min(cp09) mcp09 from caseprogress,acc1n0 where cp09=a1n03(+) and a1n02='3' " & strExSql & " group by cp60,a1n04 having count(*) > 1) vt3 " & _
         " where CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP29=S3.ST01(+) AND " & _
         "(SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+)) and substr(lc22,1,8)=fa01(+) and substr(lc22,9,1)=fa02(+) AND " & _
         "CU10=N1.na01(+) AND CP12=a0901(+) AND FA10=N2.na01(+) and cp09=v101(+) and cp14=v102(+) and cp09=v201(+) and cp14=v202(+) and cp60=v301(+) and cp14=v302(+) and cp09=mcp09(+) " & _
         "and (CP01=CPM01(+) AND CP10=CPM02(+)) AND CP57 IS NULL AND CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 AND CP09<'C' " & strExSql

      'Modified by Lydia 2016/12/13 + CP60
      'modify by sonia 2019/8/8 +CP06
      'Modified by Lydia 2020/04/10 CP06的後面+CP14
      'Modified by Lydia 2024/07/08 同一請款單只在日期+收文號最小者,顯示點數>>點數加總
      strExc(0) = "select x1.a01,x1.cp09,x1.a02,x1.a03,x1.a04,decode(x1.t01,0,decode(x1.b01,0,decode(x1.c01,0,cp18,x1.c01),x1.b01),x1.t01) dot,x1.a05,x1.a06,x1.a08,x1.cp64,x1.a09,x1.cp01,x1.cp02,x1.cp03,x1.cp04,x1.fa05,x1.n1na03,x1.a0902,x1.n2na03,x1.lc13,x1.cp404142,x1.cp60,x1.cp06,x1.cp14 " & _
                  "from (" & strExc(0) & ") x1 order by x1.a09,x1.cp01,x1.cp02,x1.cp03,x1.cp04  "

   'Modify By Sindy 2010/5/10 承辦人報表, 點數改抓點數分配檔
   Else     '列印別:承辦人   'Memo by Lydia 2020/04/10 含已取消收文CP57
       'Modified by Lydia 2015/05/08 + acc0k0 ,判斷承辦人是否有分配點數
       'Memo 2015/03/11 明細不排除已銷帳(a1k25,a0k10),業務分配則排除已銷帳(QueryPointData)
       'Modified by Lydia 2015/06/02 設定共同sql
        strExc(9) = ",acc1n0 p1,acc1n0 p2,acc1n0 p3,acc1n0 p4"
        strExSql = strExSql & "AND p1.a1n02(+)='3' AND p1.a1n03(+)=cp09 AND p1.a1n04(+)=cp14 " & _
                  "AND p2.a1n02(+)='3' AND p2.a1n03(+)=cp09 AND p2.a1n04(+)<>decode(cp14,null,'111',cp14) " & _
                  "AND p3.a1n02(+)='2' AND p3.a1n03(+)=cp09 AND p3.a1n01(+)=cp60 AND p3.a1n04(+)=cp14 " & _
                  "AND p4.a1n02(+)='2' AND p4.a1n03(+)=cp09 AND p4.a1n01(+)=cp60 AND p4.a1n04(+)<>decode(cp14,null,'111',cp14) "
      'end 2015/06/02
      
      If Me.Tag = 0 Then '內法
         '顧問基本檔
         'Modify By Sindy 2009/07/20 CP35改抓取strSelCol
         '2009/12/8 modify by sonia 加cp113工作時數
         'Modified by Lydia 2015/06/02 工作點數分配=> 改在模組處理分配點數明細
         'Modified by Lydia 2016/12/13 CP40||CP41||CP42=>CP40||' '||CP41||' '||CP42 ; +CP60
         'Modified by Lydia 2020/04/09 重新整理
         strExc(0) = "SELECT DECODE(CP14,S1.ST01,S1.ST02) as a01," & strCon1 & " as a02, CP09," & _
            "CP01||'-'||CP02||DECODE(CP03,'0','','-'||CP03)||DECODE(CP04,'00','','-'||CP04) a03," & _
            "DECODE(HC05,CU01||CU02,NVL(CU04,NVL(CU05,CU06))) a04," & _
            "DECODE(CP01||CP10,CPM01||CPM02,CPM03) a05,DECODE(CP13,S2.ST01,S2.ST02) a06," & _
            "CP16," & strSelCol & "," & _
            "SQLDATET(CP57) as a08,CP64," & strCon2 & ",CP01,CP02,CP03,CP04,CP113,'' AS LC13,CP40||' '||CP41||' '||CP42 as cp404142,S1.ST15,CP60 " & _
            "FROM STAFF S1,STAFF S2,CASEPROGRESS,HIRECASE,CASEPROPERTYMAP,CUSTOMER " & strExc(9) & " WHERE " & _
            "CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND (SUBSTR(HC05,1,8)=CU01(+) AND " & _
            "SUBSTR(HC05,9,1)=CU02(+)) AND (CP01=CPM01(+) AND CP10=CPM02(+)) AND " & _
            "CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 AND CP09<'C' " & strExSql & _
            " UNION "
      End If
      '法務基本檔
      'Modified by Lydia 2015/06/02 工作點數分配=> 改在模組處理分配點數明細
      'Modified by Lydia 2016/12/13 CP40||CP41||CP42=>CP40||' '||CP41||' '||CP42 ; +CP60
      'Modified by Lydia 2020/04/09 重新整理
      strExc(10) = "SELECT DECODE(CP14,S1.ST01,S1.ST02) as a01," & strCon1 & " as a02,CP09," & _
         "CP01||'-'||CP02||DECODE(CP03,'0','','-'||CP03)||DECODE(CP04,'00','','-'||CP04) a03," & _
         "DECODE(LC11,CU01||CU02,NVL(CU04,NVL(CU05,CU06))) a04," & _
         "DECODE(CP01||CP10,CPM01||CPM02,CPM03) a05,DECODE(CP13,S2.ST01,S2.ST02) a06," & _
         "CP16," & strSelCol & "," & _
         "SQLDATET(CP57) a08,CP64," & strCon2 & ",CP01,CP02,CP03,CP04,CP113,lc13,CP40||' '||CP41||' '||CP42 as cp404142,S1.ST15,CP60 " & _
         "FROM STAFF S1,STAFF S2,CASEPROGRESS,LAWCASE,CASEPROPERTYMAP,CUSTOMER " & strExc(9) & " WHERE " & _
         "CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND (SUBSTR(LC11,1,8)=CU01(+) AND " & _
         "SUBSTR(LC11,9,1)=CU02(+)) AND (CP01=CPM01(+) AND CP10=CPM02(+)) AND " & _
         "CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 AND CP09<'C' " & strExSql
                
      strExc(0) = strExc(0) & strExc(10)
      'Modified by Lydia 2016/12/13 + x1.cp60
      strExc(0) = "select x1.a01,x1.a02,x1.cp09,x1.a03,x1.a04,x1.a05,x1.a06,x1.cp16,x1.cp35,x1.a07,x1.a08,x1.cp64,x1.a09,x1.cp01,x1.CP02,x1.CP03,x1.CP04,x1.CP113,x1.lc13,x1.cp404142,x1.st15,x1.cp60 " & _
                 "from (" & strExc(0) & ") x1 group by x1.a01,x1.a02,x1.cp09,x1.a03,x1.a04,x1.a05,x1.a06,x1.cp16,x1.cp35,x1.a07,x1.a08,x1.cp64,x1.a09,x1.cp01,x1.CP02,x1.CP03,x1.CP04,x1.CP113,x1.lc13,x1.cp404142,x1.st15,x1.cp60 " & _
                 " ORDER BY ST15,a09,CP01,CP02,CP03,CP04"
   End If
   If RsTemp.State = adStateOpen Then RsTemp.Close
   RsTemp.Open strExc(0), cnnConnection
   If RsTemp.EOF And RsTemp.BOF Then
      'Modify By Sindy 2010/5/6
      If Opt1(0).Value = True Then '列印別:日期
         MsgBox "資料庫內無資料 !", vbInformation
      Else
         'If dblPointTot > 0 Then
         'Modify By Sindy 2012/2/3
         If dblPointRow > 0 Then
            Page = 1
            CaseTitle
            Call PrintEndPoint("", "")
            ShowLine
            PrintEnd1
            Printer.EndDoc
            ShowPrintOk
            Exit Sub
         Else
            MsgBox "資料庫內無資料 !", vbInformation
         End If
      End If '列印別
      
      m_print = 1
      Exit Sub
   End If
   
   '【開始產生報表資料...】
   Page = 1
   CaseTitle
   With RsTemp
   .MoveFirst
   If Opt1(0).Value = True Then '列印別:日期
      Do While Not .EOF
         If iPrint >= 11000 Then
            Page = Page + 1
            Printer.NewPage
            CaseTitle
         End If
         '第1欄=by日期：收文日
         Printer.CurrentX = PLeft(0):      Printer.CurrentY = iPrint
         Printer.Print "" & .Fields("a01")
         '第2欄=by日期：收文號CP09
         Printer.CurrentX = PLeft(1):    Printer.CurrentY = iPrint
         Printer.Print "" & .Fields("cp09")
         '第3欄=by日期：本所案號
         Printer.CurrentX = PLeft(2):    Printer.CurrentY = iPrint
         Printer.Print "" & .Fields("a02")
         '第4欄=by日期：案件性質
         Printer.CurrentX = PLeft(3):    Printer.CurrentY = iPrint
         Printer.Print StrToStr("" & .Fields("a03"), 6)
         
         '第5欄=by日期：IP案
         '2009/12/8 add by sonia IP案件lc13
         Printer.CurrentX = PLeft(10):    Printer.CurrentY = iPrint
         Printer.Print "" & .Fields("lc13")
         '2009/12/8 end
         
         '第6欄=by日期：當事人
         Printer.CurrentX = PLeft(4):    Printer.CurrentY = iPrint
         Printer.Print convForm("" & .Fields("a04"), 12)
         
         '第7欄=by日期：點數
         'Modified by Lydia 2016/12/13 同一請款單只在日期+收文號最小者,顯示點數
         If strA1k01List <> "" And Mid("" & .Fields("CP60"), 1, 1) = "X" And InStr(strA1k01List, "" & .Fields("CP60")) > 0 Then
            '相同請款單不顯示點數
            'Memo by Lydia 2024/07/08 語法已先加總同一請款單的點數
         Else
            Printer.CurrentX = PLeft(5):    Printer.CurrentY = iPrint
            If Val("" & .Fields("dot")) > 0 Then
                If IsNull(.Fields("dot")) = False Then Printer.Print StrToStr(Format("" & .Fields("dot"), "0.0"), 8)
                'Added by Lydia 2016/12/13 目前只有外法有
                If "" & .Fields("CP60") <> "" And Me.Tag = 1 Then
                   strA1k01List = strA1k01List & .Fields("CP60") & ","
                End If
            End If
             'Add By Sindy 2009/08/12 合計點數
             'Modified by Lydia 2015/06/08 內法+合計
             'If Me.Tag = 1 Then '外法
                If Left(Trim(.Fields("a02")), 3) = "FCL" Or Left(Trim(.Fields("a02")), 3) = "LIN" Then
                   If IsNull(.Fields("dot")) = False Then m_totFCLCP18 = m_totFCLCP18 + Val(Format(.Fields("dot"), FatDot))  'Modified by Lydia 2015/06/08 統一小數點
                Else
                   If IsNull(.Fields("dot")) = False Then m_totCFLCP18 = m_totCFLCP18 + Val(Format(.Fields("dot"), FatDot))  'Modified by Lydia 2015/06/08 統一小數點
                End If
                If IsNull(.Fields("dot")) = False Then m_totCP18 = m_totCP18 + Val(Format(.Fields("dot"), FatDot))   'Modified by Lydia 2015/06/08 統一小數點
            ' End If
             '2009/08/12 End
         End If
         'end 2016/12/13
         
         '第8欄=by日期：智權人員
         Printer.CurrentX = PLeft(6):    Printer.CurrentY = iPrint
         Printer.Print convForm("" & .Fields("a05"), 8)
         '第9欄=by日期：承辦人員／第5欄=by承辦人：當事人
         Printer.CurrentX = PLeft(7):    Printer.CurrentY = iPrint
         'Modified by Lydia 2020/04/10
         Printer.Print convForm("" & .Fields("a06"), 8)
          
         '第10欄=by日期
         Printer.CurrentX = PLeft(8):    Printer.CurrentY = iPrint
         'Add By Sindy 2009/08/12 國籍/業務區
         If Me.Tag = 1 Then '外法
            If Left(Trim(.Fields(2)), 3) = "FCL" Or Left(Trim(.Fields(2)), 3) = "LIN" Then
               If IsNull(.Fields("N1na03")) = False Then
                  Printer.Print StrToStr("" & .Fields("N1na03"), 4)
               Else
                  If IsNull(.Fields("N2na03")) = False Then Printer.Print StrToStr("" & .Fields("N2na03"), 4)
               End If
            Else
               If IsNull(.Fields("a0902")) = False Then Printer.Print StrToStr("" & .Fields("a0902"), 4)
            End If
         ElseIf Me.Tag = 0 Then '內法
             '第10欄=by日期：協辦人員
             Printer.Print convForm("" & .Fields("a08"), 8)
         'add by sonia 2019/8/8
         ElseIf Me.Tag = 2 Then    'ACS-本所期限
            If IsNull(.Fields("CP06")) = False Then Printer.Print .Fields("CP06")
         'end 2019/8/8
         End If
         '2009/08/12 End
         
         '第11欄=by日期
         Printer.CurrentX = PLeft(9):    Printer.CurrentY = iPrint
         '2008/9/9 modify by sonia 外法列印時cp64改印fa05
         'If IsNull(.Fields(9)) = False Then Printer.Print StrToStr("" & .Fields(9), 12)
         If Me.Tag <> 1 Then  '非外法：內法, ACS
            '備註
            'Modified by Lydia 2015/06/09
            'If IsNull(.Fields(9)) = False Then Printer.Print StrToStr("" & .Fields(9), 12)
             Printer.Print convForm("" & .Fields("cp64"), 16)
         Else  '外法
            'FA05
            If IsNull(.Fields(15)) = False Then Printer.Print StrToStr("" & .Fields(15), 12)
         End If
         '2008/9/9 END
         
         '第2行：對造當事人
         'Add By Sindy 2010/3/9 增加對造當事人
         'Modified by Lydia 2016/12/13
         'If IsNull(.Fields("cp404142")) = False Then
         If Trim("" & .Fields("cp404142")) <> "" Then
            iPrint = iPrint + 300
            Printer.CurrentX = PLeft(4):    Printer.CurrentY = iPrint
            Printer.Print Replace(.Fields("cp404142"), vbCrLf, "")
         End If
         '2010/3/9 End
         iPrint = iPrint + 300
         
         .MoveNext
         If .EOF Then 'Memo by Lydia 2020/04/09 最後一筆做合計列印
            'Add By Sindy 2009/08/12
            '外法: 合計
            If Me.Tag = 1 Then
               ShowLine
               If Text1(4).Text = "1" Or Text1(4).Text = "3" Then
                  If iPrint >= 11000 Then
                     Page = Page + 1
                     Printer.NewPage
                     CaseTitle
                  End If
                  Printer.CurrentX = PLeft(4):    Printer.CurrentY = iPrint
                  Printer.Print "FCL+LIN"
                  Printer.CurrentX = PLeft(5):    Printer.CurrentY = iPrint
                  Printer.Print StrToStr(CStr(Format(m_totFCLCP18, "0.0")), 8) & " 點"
                  iPrint = iPrint + 300
               End If
               If Text1(4).Text = "2" Or Text1(4).Text = "3" Then
                  If iPrint >= 11000 Then
                     Page = Page + 1
                     Printer.NewPage
                     CaseTitle
                  End If
                  Printer.CurrentX = PLeft(4):    Printer.CurrentY = iPrint
                  Printer.Print "CFL"
                  Printer.CurrentX = PLeft(5):    Printer.CurrentY = iPrint
                  Printer.Print StrToStr(CStr(Format(m_totCFLCP18, "0.0")), 8) & " 點"
                  iPrint = iPrint + 300
               End If
            End If
            '2009/08/12 End
            'End If 'Remove by Lydia 2015/06/08
               'Added by Lydia 2015/06/08 +合計
               If Text1(4).Text = "3" Then '
                  ShowLine
                  If iPrint >= 11000 Then
                     Page = Page + 1
                     Printer.NewPage
                     CaseTitle
                  End If
                  Printer.CurrentX = PLeft(4):    Printer.CurrentY = iPrint
                  Printer.Print "合計"
                  Printer.CurrentX = PLeft(5):    Printer.CurrentY = iPrint
                  Printer.Print StrToStr(CStr(Format(m_totCP18, FatDot)), 8) & " 點"
                  iPrint = iPrint + 300
               End If
               'end 2015/06/08
            'End If  'Remove by Lydia 2015/06/08 調整
            '2009/08/12 End
         End If
      Loop
   Else '列印別:承辦人
      strEmp = "" 'Add By Sindy 2010/5/11
      Dim dblPoint As Double
      dblPoint = 0 'Add By Sindy 2012/2/3
      Do While Not .EOF
         'Add By Sindy 2010/5/11
         'If strEmp <> ("" & .Fields(11)) Then
         'Modified by Lydia 2015/06/08 內法+點數
         'If strEmp <> ("" & .Fields(19) & "" & .Fields(11)) Then
         If strEmp <> ("" & .Fields("st15") & "" & .Fields("a09")) Then '部門: 小計
            If strEmp <> "" Then
               ShowLine
            'Add By Sindy 2012/2/3 加入小計
               If dblPoint > 0 Then
                  Printer.CurrentX = PLeft(6) - 200:    Printer.CurrentY = iPrint
                  Printer.Print "小計"
                  If dblPoint > 0 Then
                     Printer.CurrentX = PLeft(7):      Printer.CurrentY = iPrint
                     Printer.Print "點數：" & StrToStr(Format(dblPoint, "0.0"), 8)
                  End If
                  Call GetPointTotSub(Right(Trim(strEmp), 5)) '取得此人員的分配點數小計
                  If dblPointTotSub > 0 Then
                    Printer.CurrentX = PLeft(9):      Printer.CurrentY = iPrint
                    Printer.Print "分配點數：" & dblPointTotSub
                  End If
                  iPrint = iPrint + 300 'Add By Sindy 2012/9/5
               End If
            End If
            dblPoint = 0
            '2012/2/3 End
            
            'Modified by Lydia 2015/06/08 內法+點數
            'If PrintEndPoint(strEmp, ("" & .Fields(19) & "" & .Fields(11))) = True Then
            If PrintEndPoint(strEmp, ("" & .Fields("st15") & "" & .Fields("a09"))) = True Then
               ShowLine
            End If
         End If
         '2010/5/11 End
         If iPrint >= 11000 Then
            Page = Page + 1
            Printer.NewPage
            CaseTitle
         End If
         '第1欄=by承辦人：承辦人
         Printer.CurrentX = PLeft(0):      Printer.CurrentY = iPrint
         Printer.Print convForm("" & .Fields("a01"), 8)
         'Add By Sindy 2010/5/11
         'strEmp = "" & .Fields(11)
         'Modified by Lydia 2015/06/08 內法+點數
         'strEmp = "" & .Fields(19) & "" & .Fields(11)
         strEmp = "" & .Fields("st15") & "" & .Fields("a09")
         '第2欄=by承辦人：收文日
         Printer.CurrentX = PLeft(1):    Printer.CurrentY = iPrint
         Printer.Print "" & .Fields("a02")
         '第3欄=by承辦人：收文號
         Printer.CurrentX = PLeft(2):    Printer.CurrentY = iPrint
         Printer.Print "" & .Fields("cp09")
         '第4欄=by承辦人：本所案號
         Printer.CurrentX = PLeft(3):    Printer.CurrentY = iPrint
         Printer.Print "" & .Fields("a03")
         '第5欄=by承辦人：當事人
         Printer.CurrentX = PLeft(4):    Printer.CurrentY = iPrint
         Printer.Print convForm("" & .Fields("a04"), 12)
         '第6欄=by承辦人：案件性質
         Printer.CurrentX = PLeft(5):    Printer.CurrentY = iPrint
         Printer.Print convForm("" & .Fields("a05"), 12)
         '第7欄=by承辦人：智權人員
         Printer.CurrentX = PLeft(6):    Printer.CurrentY = iPrint
         Printer.Print convForm("" & .Fields("a06"), 8)
         '第8欄=by承辦人：費用
         Printer.CurrentX = PLeft(7):    Printer.CurrentY = iPrint
         If Val("" & .Fields("cp16")) > 0 Then
            Printer.Print StrToStr(Format("" & .Fields("cp16"), "#,##0"), 8)
         End If
         '點數 / 法院案號: 'Modified by Lydia 2015/? ：分別抓
         'Modified by Lydia 2015/? ：分別抓
'         Printer.CurrentX = PLeft(8):    Printer.CurrentY = iPrint
'         'Add By Sindy 2009/08/12 合計點數
'         If IsNull(.Fields(8)) = False Then
'            If Me.Tag = 1 Then '外法
'               If Val("" & .Fields(8)) > 0 Then
'                  Printer.Print StrToStr(Format("" & .Fields(8), fatdot), 8)
'                  If Left(Trim(.Fields(3)), 3) = "FCL" Or Left(Trim(.Fields(3)), 3) = "LIN" Then
'                     m_totFCLCP18 = m_totFCLCP18 + .Fields(8)
'                  Else
'                     m_totCFLCP18 = m_totCFLCP18 + .Fields(8)
'                  End If
'                  m_totCP18 = m_totCP18 + .Fields(8)
'                  dblPoint = dblPoint + .Fields(8) 'Add By Sindy 2012/2/3
'               End If
'            Else
'               Printer.Print StrToStr("" & .Fields(8), 8)
'            End If
'         End If

         '第9欄=by承辦人：點數
         If Me.Tag <> 2 Then
            Printer.CurrentX = PLeft(13):    Printer.CurrentY = iPrint
            If IsNull(.Fields("A07")) = False Then
               'Memo by Lydia 2016/12/13 因為acc1n0有判斷收文號,所在不做同一請款單只在日期+收文號最小者顯示點數
               If Val("" & .Fields("A07")) > 0 Then
                  Printer.Print StrToStr(Format("" & .Fields("A07"), FatDot), 8)
                  If Left(Trim(.Fields(3)), 3) = "FCL" Or Left(Trim(.Fields(3)), 3) = "LIN" Then
                     m_totFCLCP18 = m_totFCLCP18 + Val(Format(.Fields("A07"), FatDot))
                  Else
                     m_totCFLCP18 = m_totCFLCP18 + Val(Format(.Fields("A07"), FatDot))
                  End If
                  m_totCP18 = m_totCP18 + Val(Format(.Fields("A07"), FatDot))
                  dblPoint = dblPoint + Val(Format(.Fields("A07"), FatDot))
               End If
            End If
         End If
         '第10欄=by承辦人：法院案號
         If Me.Tag = 0 Then '內法
            Printer.CurrentX = PLeft(8):    Printer.CurrentY = iPrint
            Printer.Print convForm("" & .Fields("CP35"), 16)
         End If
'end 2015/06/08 內法+點數
         
         '第11欄=by承辦人：工時CP113
         '2009/12/8 add by sonia 加CP113及LC13
         Printer.CurrentX = PLeft(9):    Printer.CurrentY = iPrint
         If IsNull(.Fields("cp113")) = False Then
            Printer.Print .Fields("cp113")
            If Left(Trim(.Fields(3)), 3) = "FCL" Or Left(Trim(.Fields(3)), 3) = "LIN" Then
               m_totFCLCP113 = m_totFCLCP113 + .Fields("cp113")
            Else
               m_totCFLCP113 = m_totCFLCP113 + .Fields("cp113")
            End If
            m_totCP113 = m_totCP113 + .Fields("cp113")
         End If
         '第12欄=by承辦人：IP案
         Printer.CurrentX = PLeft(10) + 300:  Printer.CurrentY = iPrint
         Printer.Print "" & .Fields("LC13")
         '2009/12/8 end
         '第13欄=by承辦人：取消日CP57
         Printer.CurrentX = PLeft(11):    Printer.CurrentY = iPrint
         Printer.Print Replace("" & .Fields("A08"), "//", "")
         
         '第14欄=by承辦人：備註
         Printer.CurrentX = PLeft(12):    Printer.CurrentY = iPrint
         Printer.Print convForm("" & .Fields("CP64"), 16)
         
         '第2行：對造當事人
         'Add By Sindy 2010/3/9 增加對造當事人
         'Modified by Lydia 2016/12/13
         'If IsNull(.Fields("cp404142")) = False Then
         If Trim("" & .Fields("cp404142")) <> "" Then
            iPrint = iPrint + 300
            Printer.CurrentX = PLeft(4):    Printer.CurrentY = iPrint
            Printer.Print Replace(.Fields("cp404142"), vbCrLf, "")
         End If
         '2010/3/9 End
         iPrint = iPrint + 300
         
         .MoveNext
         If .EOF Then 'Memo by Lydia 2020/04/09 最後一筆做合計列印
            ShowLine
            'Add By Sindy 2012/2/3 加入小計
            If dblPoint > 0 Then
               Printer.CurrentX = PLeft(6) - 200:    Printer.CurrentY = iPrint
               Printer.Print "小計"
               If dblPoint > 0 Then
                  Printer.CurrentX = PLeft(7):      Printer.CurrentY = iPrint
                  Printer.Print "點數：" & StrToStr(Format(dblPoint, "0.0"), 8)
               End If
               Call GetPointTotSub(Right(Trim(strEmp), 5)) '取得此人員的分配點數小計
               If dblPointTotSub > 0 Then
                 Printer.CurrentX = PLeft(9):      Printer.CurrentY = iPrint
                 Printer.Print "分配點數：" & dblPointTotSub
               End If
            End If
            '2012/2/3 End
            If PrintEndPoint(strEmp, "") = True Then 'Add By Sindy 2010/5/11
               ShowLine
            End If
            'Add By Sindy 2009/08/12
            '外法：合計
            If Me.Tag = 1 Then
               'Add By Sindy 2012/10/29
               If dblPoint > 0 Then
                  iPrint = iPrint + 300
               End If
               '2012/10/29 End
               If Text1(4).Text = "1" Or Text1(4).Text = "3" Then
                  If iPrint >= 11000 Then
                     Page = Page + 1
                     Printer.NewPage
                     CaseTitle
                  'Added by Lydia 2015/06/08 多空一行
                  Else
                     iPrint = iPrint + 300
                  'end 2015/06/08
                  End If
                  Printer.CurrentX = PLeft(6) - 200:  Printer.CurrentY = iPrint
                  Printer.Print "FCL+LIN"
                  Printer.CurrentX = PLeft(8) - 200:  Printer.CurrentY = iPrint
                  Printer.Print StrToStr(CStr(Format(m_totFCLCP18, "0.0")), 8) & " 點"
                  
                  '2009/12/8 add by sonia
                  'Modified by Lydia 2015/06/09
                  'Printer.CurrentX = PLeft(9):  Printer.CurrentY = iPrint
                  Printer.CurrentX = PLeft(8) + 1200:  Printer.CurrentY = iPrint
                  Printer.Print StrToStr(CStr(m_totFCLCP113), 8) & " 時"
                  '2009/12/8 end

                  iPrint = iPrint + 300
               End If
               If Text1(4).Text = "2" Or Text1(4).Text = "3" Then
                  If iPrint >= 11000 Then
                     Page = Page + 1
                     Printer.NewPage
                     CaseTitle
                  'Added by Lydia 2015/06/08 多空一行
                  Else
                     If Text1(4).Text = "2" Then iPrint = iPrint + 300
                  'end 2015/06/08
                  End If
                  Printer.CurrentX = PLeft(6) - 200:  Printer.CurrentY = iPrint
                  Printer.Print "CFL"
                  Printer.CurrentX = PLeft(8) - 200:  Printer.CurrentY = iPrint
                  Printer.Print StrToStr(CStr(Format(m_totCFLCP18, "0.0")), 8) & " 點"
                  '2009/12/8 add by sonia
                  'Modified by Lydia 2015/06/09
                  'Printer.CurrentX = PLeft(9):  Printer.CurrentY = iPrint
                  Printer.CurrentX = PLeft(8) + 1200:  Printer.CurrentY = iPrint
                  Printer.Print StrToStr(CStr(m_totCFLCP113), 8) & " 時"
                  '2009/12/8 end
                  iPrint = iPrint + 300
               End If
            End If
               'Added by Lydia 2015/06/08 +合計
               If Text1(4).Text = "3" Then
                    If dblPoint > 0 And Me.Tag = 0 Then
                       iPrint = iPrint + 300
                    End If
                  ShowLine
                  If iPrint >= 11000 Then
                     Page = Page + 1
                     Printer.NewPage
                     CaseTitle
                  End If
                  Printer.CurrentX = PLeft(6) - 200:  Printer.CurrentY = iPrint
                  Printer.Print "合計"
                  Printer.CurrentX = PLeft(8) - 200:  Printer.CurrentY = iPrint
                  Printer.Print StrToStr(CStr(Format(m_totCP18, FatDot)), 8) & " 點"
                  '2009/12/8 add by sonia
                  'Modified by Lydia 2015/06/09
                  'Printer.CurrentX = PLeft(9):  Printer.CurrentY = iPrint
                  Printer.CurrentX = PLeft(8) + 1200:  Printer.CurrentY = iPrint
                  Printer.Print StrToStr(CStr(m_totCP113), 8) & " 時"
                  '2009/12/8 end
                  iPrint = iPrint + 300
               End If
               ShowLine
            'end 2015/06/08
            
            '2009/08/12 End
            PrintEnd1
         End If
      Loop
   End If
   End With
   'Add By Sindy 2010/5/11
   iPrint = iPrint + 300
   If iPrint >= 11000 Then
      Page = Page + 1
      Printer.NewPage
      CaseTitle
   End If
   Printer.CurrentX = 500:               Printer.CurrentY = iPrint
   If Opt1(0).Value = True Then '列印別:日期
      'Modify by Morgan 2011/6/1
      'Printer.Print "PS.點數為請款單點數，未扣除跨部門合作點數"
      Printer.Print "PS.點數為收據或請款單點數，未扣除跨部門合作點數，但有扣除專利處配合開庭分配點數。"
   Else
      Printer.Print "PS.不含非個人點數"
   End If
   iPrint = iPrint + 300
   '2010/5/11 End
   Printer.EndDoc
   Exit Sub
   
ErrHand:
   MsgBox Err.Description
   Resume
End Sub

Private Sub CaseTitle()
   iPrint = 500
   Printer.Orientation = 2
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 7000:         Printer.CurrentY = iPrint
   If Text1(0).Text = "1" Then
      Printer.Print "收文明細表"
   Else
      Printer.Print "發文明細表"
   End If
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   iPrint = iPrint + 500
   Printer.CurrentX = 6900:         Printer.CurrentY = iPrint
   Printer.Print "日期 : " & ChangeTStringToTDateString(Text1(1)) & _
      " - " & ChangeTStringToTDateString(Text1(2))
   Printer.Font.Bold = False
   iPrint = iPrint + 300
   Printer.CurrentX = 500:               Printer.CurrentY = iPrint
   Printer.Print "列印人 　: " & strUserName
   Printer.CurrentX = 13000:             Printer.CurrentY = iPrint
   Printer.Print "列印日期 : " & ChangeTStringToTDateString(GetTaiwanTodayDate)
   '2008/9/9 add by sonia
   iPrint = iPrint + 300
   Printer.CurrentX = 500:               Printer.CurrentY = iPrint
   If Opt1(0).Value = True Then
      Printer.Print "列印順序 : " & Opt1(0).Caption
   Else
      Printer.Print "列印順序 : " & Opt1(1).Caption
   End If
   '2008/9/9 end
   Printer.CurrentX = 13000:               Printer.CurrentY = iPrint
   Printer.Print "頁次 : " & Page
   iPrint = iPrint + 300
   ShowLine
   If Opt1(0).Value = True Then '列印別:日期
      Printer.CurrentX = PLeft(0):          Printer.CurrentY = iPrint
      If Text1(0).Text = "1" Then
         Printer.Print "收文日"
      Else
         Printer.Print "發文日"
      End If
      Printer.CurrentX = PLeft(1):          Printer.CurrentY = iPrint
      Printer.Print "收文號"
      Printer.CurrentX = PLeft(2):          Printer.CurrentY = iPrint
      Printer.Print "本所案號"
      Printer.CurrentX = PLeft(3):          Printer.CurrentY = iPrint
      Printer.Print "案件性質"
      '2009/12/8 add by sonia
      Printer.CurrentX = PLeft(10) - 300:   Printer.CurrentY = iPrint
      'modify by sonia 2019/8/8
      'Printer.Print "IP案"
      If Me.Tag <> 2 Then Printer.Print "IP案"
      '2009/12/8 end
      Printer.CurrentX = PLeft(4):          Printer.CurrentY = iPrint
      Printer.Print "當事人"
      Printer.CurrentX = PLeft(5):          Printer.CurrentY = iPrint
      Printer.Print "點數"
      Printer.CurrentX = PLeft(6):          Printer.CurrentY = iPrint
      Printer.Print "智權人員"
      Printer.CurrentX = PLeft(7):          Printer.CurrentY = iPrint
      Printer.Print "承辦人"
      Printer.CurrentX = PLeft(8):          Printer.CurrentY = iPrint
      'Modify By Sindy 2009/08/12
      If Me.Tag = 1 Then '外法
         Printer.Print "國籍/業務區"
      '2009/08/12 End
      ElseIf Me.Tag = 0 Then  '內法
         'Modified by Lydia 2015/10/05
         'Printer.Print "法務人員"
         Printer.Print "協辦人員"
      'add by sonia 2019/8/8 +ACS
      ElseIf Me.Tag = 2 Then
         Printer.Print "本所期限"
      'end 2019/8/8
      End If
      Printer.CurrentX = PLeft(9):          Printer.CurrentY = iPrint
      '2008/9/9 modify by sonia 外法列印時cp64備註改印fa05(FC代理人)
      'Printer.Print "備註"
      If Me.Tag = 1 Then
         Printer.Print "FC代理人"
      Else
         Printer.Print "備註"
      End If
      If Me.Tag <> 2 Then 'add by sonia 2019/8/8
         'Add By Sindy 2010/3/9
         iPrint = iPrint + 300
         Printer.CurrentX = PLeft(4):          Printer.CurrentY = iPrint
         Printer.Print "對造當事人"
      End If 'add by sonia 2019/8/8
      '2010/3/9 End
   Else  '列印別:承辦人
      Printer.CurrentX = PLeft(0):          Printer.CurrentY = iPrint
      Printer.Print "承辦人"
      Printer.CurrentX = PLeft(1):          Printer.CurrentY = iPrint
      If Text1(0).Text = "1" Then
         Printer.Print "收文日"
      Else
         Printer.Print "發文日"
      End If
      Printer.CurrentX = PLeft(2):          Printer.CurrentY = iPrint
      Printer.Print "收文號"
      Printer.CurrentX = PLeft(3):          Printer.CurrentY = iPrint
      Printer.Print "本所案號"
      Printer.CurrentX = PLeft(4):          Printer.CurrentY = iPrint
      Printer.Print "當事人"
      Printer.CurrentX = PLeft(5):          Printer.CurrentY = iPrint
      Printer.Print "案件性質"
      Printer.CurrentX = PLeft(6):          Printer.CurrentY = iPrint
      Printer.Print "智權人員"
      Printer.CurrentX = PLeft(7):          Printer.CurrentY = iPrint
      Printer.Print "費用"
      
      'Modified by Lydia 2020/04/09 重新整理
      If Me.Tag = 0 Then '內法
            Printer.CurrentX = PLeft(8) - 500:        Printer.CurrentY = iPrint
            Printer.Print "法院案號"
      ElseIf Me.Tag = 1 Then '外法
            Printer.CurrentX = PLeft(13) - 100:        Printer.CurrentY = iPrint
            Printer.Print "點數"
      ElseIf Me.Tag = 0 Then '創新業務 ACS
      End If

      '2009/12/8 add BY SONIA
      Printer.CurrentX = PLeft(9) - 200:        Printer.CurrentY = iPrint
      Printer.Print "工時"
      Printer.CurrentX = PLeft(10):          Printer.CurrentY = iPrint
      'modify by sonia 2019/8/8
      'Printer.Print "IP案件"
      If Me.Tag <> 2 Then Printer.Print "IP案件"
      '2009/12/8 END
      Printer.CurrentX = PLeft(11):          Printer.CurrentY = iPrint
      Printer.Print "取消日"
      Printer.CurrentX = PLeft(12):         Printer.CurrentY = iPrint
      Printer.Print "備註"
      'Add By Sindy 2010/3/9
      If Me.Tag <> 2 Then 'add by sonia 2019/8/8
         iPrint = iPrint + 300
         Printer.CurrentX = PLeft(4):          Printer.CurrentY = iPrint
         Printer.Print "對造當事人"
      End If 'add by sonia 2019/8/8
      '2010/3/9 End
   End If
   iPrint = iPrint + 300
   ShowLine
End Sub

Private Sub GetPrintLeft()
   Erase PLeft
   If Opt1(0).Value = False Then
      PLeft(0) = 500        '承辦人
      PLeft(1) = 1500      '收文日期
      'Modified by Lydia 2024/07/08
      'PLeft(2) = 2500      '收文號
      PLeft(2) = 2600      '收文號
      PLeft(3) = 3800      '本所案號
      PLeft(5) = 6700      '案件性質
      PLeft(4) = 5100      '當事人
      PLeft(6) = 8300      '智權人員
      PLeft(7) = 9500      '費用
      'Added by Lydia 2015/06/08 內法+點數
      If Me.Tag = 0 Then
            PLeft(13) = 10600    '點數
            PLeft(8) = 11700     '法院案號
            PLeft(9) = 12600     '工作時數
            PLeft(10) = 13000    '是否為智慧財產權案
            PLeft(11) = 13800    '取消收文日期
            PLeft(12) = 14700    '備註
      Else
      'end 2015/06/08
            PLeft(8) = 10600     '法院案號
            PLeft(13) = PLeft(8) 'Added by Lydia 2015/06/08
            '2009/12/8 add BY SONIA 加工作時數CP113加是否為智慧財產權案LC13
            PLeft(9) = 11500     '工作時數
            PLeft(10) = 11900    '是否為智慧財產權案
            '2009/12/8 END
            'Modified by Lydia 2024/07/08
            'PLeft(11) = 12700    '取消收文日期
            PLeft(11) = 12800    '取消收文日期
            PLeft(12) = 13600    '備註
      End If  'Added by Lydia 2015/06/08
   Else
      PLeft(0) = 500        '收文日期
      PLeft(1) = 1700      '收文號
      PLeft(2) = 2900      '本所案號
      PLeft(3) = 4300      '案件性質
      PLeft(10) = 5900      '是否為智慧財產權案LC13  2009/12/8 add BY SONIA
      PLeft(4) = 6200      '當事人
      PLeft(5) = 8300      '點數
      PLeft(6) = 9400      '智權人員
      PLeft(7) = 10400    '承辦人
      PLeft(8) = 11400    '協辦人員
      PLeft(9) = 12800    '備註
   End If
End Sub

Private Sub Form_Activate()
   Text1(0).SetFocus
   'Add By Sindy 2009/08/20
   If Me.Tag <> 1 Then
       Label5(2).Visible = False
       Text1(4).Visible = False
   Else
       Label5(2).Visible = True
       Text1(4).Visible = True
   End If
   '2009/08/20 End
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
         If (KeyAscii > 50 Or KeyAscii < 49) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 3
         KeyAscii = UpperCase(KeyAscii)
      'Add By Sindy 2009/08/20
      Case 4
         If (KeyAscii > 51 Or KeyAscii < 49) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      '2009/08/20 End
   End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   Select Case Index
   Case 2
      If Me.Text1(1).Text <> "" And Me.Text1(2).Text <> "" Then
         If Val(Me.Text1(1).Text) > Val(Me.Text1(2).Text) Then
            MsgBox "日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
            Me.Text1(1).SetFocus
            Text1_GotFocus 1
            Exit Sub
         End If
      End If
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
 Dim strTempName As String, i As Integer, t As Integer
   If Text1(Index) = "" Then Label2 = "": Exit Sub
   Select Case Index
      Case 1, 2
         If CheckIsTaiwanDate(Text1(Index)) = False Then Cancel = True
      Case 3
         If ClsPDGetStaffN(Text1(Index), strTempName) Then
            Label2 = strTempName
         Else
            Label2 = ""
            Cancel = True
         End If
      End Select
      If Cancel Then TextInverse Text1(Index)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm083006 = Nothing
End Sub

Sub ShowLine()
   Printer.CurrentX = 500:               Printer.CurrentY = iPrint
   'Modified by Lydia 2015/06/09
   'Printer.Print String(200, "-")
   Printer.Print String(215, "-")
   iPrint = iPrint + 300
   If iPrint >= 11000 Then
      Page = Page + 1
      Printer.NewPage
      CaseTitle
   End If
End Sub

'Add By Sindy 2010/5/10
'非自己辦理的案件但點數歸屬自己
Private Sub QueryPointData(strConSql As String)
Dim strEmp As String 'Add By Sindy 2012/2/3
Dim rsA1 As New ADODB.Recordset 'Added by Lydia 2016/12/13

   With grd1
      .Clear
      .Rows = 2
      .FixedRows = 1
      .FixedCols = 0
      'Modified by Lydia 2015/05/08 +CP60
      '.FormatString = "cp12|a1n04|a1n03|a1n05|cp01|cp02|cp03|cp04|cp10|a1k02"
      .FormatString = "cp12|a1n04|a1n03|a1n05|cp01|cp02|cp03|cp04|cp10|a1k02|cp60"
   End With
   intCompRow = 0
   dblPointRow = 0 'Add By Sindy 2012/2/3
   dblPointTot = 0
   '依承辦人統計 AND substr(a.st15,1,2)=substr(b.st15,1,2)
    'Modified by Lydia 2015/05/08 +收據資料
'   If opt1(1).Value = True Then
'      strSql = "SELECT NVL(A0902,A0903) tt,a1n04,a1n03,a1n05,cp01,cp02,cp03,cp04,CPM03,a1k02,cp14," & IIf(Trim(Text1(0).Text) = "1", "cp05", "cp27") & ",a.st02,a.ST15,' ' 點數小計 " & _
'                     "From CASEPROGRESS,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff a,staff b " & _
'                     "WHERE cp60>'X' AND cp14 is not null " & _
'                     "AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)<>cp14 and a1n05>0 " & _
'                     "AND a1k01(+)=cp60 AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
'                     "AND a1k25 is null " & _
'                     "AND CP12=A0901(+) AND a1n04=a.st01(+) AND cp14=b.st01(+) " & strConSql & _
'                     "Order By a.ST15,a1n04,a1k02,cp01,cp02,cp03,cp04 "
'   '依智權人員統計 AND substr(a.st15,1,2)=substr(b.st15,1,2)
'   Else
'      strSql = "SELECT NVL(A0902,A0903) tt,a1n04,a1n03,a1n05,cp01,cp02,cp03,cp04,CPM03,a1k02,cp13," & IIf(Trim(Text1(0).Text) = "1", "cp05", "cp27") & ",a.st02,a.ST15,' ' 點數小計 " & _
'                     "From CASEPROGRESS,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff a,staff b " & _
'                     "WHERE cp60>'X' AND cp13 is not null " & _
'                     "AND a1n01(+)=cp60 AND a1n02(+)='1' AND a1n03(+)=cp09 AND a1n04(+)<>cp13 and a1n05>0 " & _
'                     "AND a1k01(+)=cp60 AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
'                     "AND a1k25 is null " & _
'                     "AND CP12=A0901(+) AND a1n04=a.st01(+) AND cp13=b.st01(+) " & strConSql & _
'                     "Order By a.ST15,a1n04,a1k02,cp01,cp02,cp03,cp04 "
'   End If
                
    If Text1(3) <> "" Then
      strSql = strSql & " and CP14='" + Text1(3) + "' "
      'Modified by Lydia 2016/12/13 移到QueryPointData
      'strPointSql = strPointSql & " and a1n04='" + Text1(3) + "' " 'Add By Sindy 2010/5/4
   End If
   
   If Opt1(1).Value = True Then
        'Modified by Lydia 2015/06/02 +工作點數分配
      'strSql = "SELECT NVL(A0902,A0903) tt,a1n04,a1n03,a1n05,cp01,cp02,cp03,cp04,CPM03" & _
                " ,decode(sign(a0k02),1,a0k02,a1k02) as a1k02,cp14," & IIf(Trim(Text1(0).Text) = "1", "cp05", "cp27") & ",a.st02,a.ST15,' ' 點數小計,cp60 " & _
                     "From CASEPROGRESS,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff a,staff b,acc0k0 " & _
                     "WHERE (substr(cp60,1,1) ='X' or substr(cp60,1,1) ='E') AND cp14 is not null " & _
                     "AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)<>cp14 and a1n05>0 " & _
                     "AND a1k01(+)=cp60 AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
                     "AND a1k25 is null AND a0k01(+)=cp60 and a0k10 is null " & _
                     "AND CP12=A0901(+) AND a1n04=a.st01(+) AND cp14=b.st01(+) " & strConSql & _
                     "Order By a.ST15,a1n04,a1k02,cp01,cp02,cp03,cp04 "
      'Modified by Lydia 2016/12/13 +承辦人條件 IIf(Text1(3) <> "", " and (p1.a1n04=" & CNULL(Trim(Text1(3))) & " or p2.a1n04=" & CNULL(Trim(Text1(3))) & ") ", "")
      'Modified by Lydia 2020/04/17 拿掉判斷銷帳=> and a0k10 is null , AND a1k25 is null
      strSql = "select * from (SELECT NVL(A0902,A0903) tt,nvl(p2.a1n04,p1.a1n04) n04,nvl(p2.a1n03,p1.a1n01) n03,nvl(p2.a1n05,p1.a1n05) n05, " & _
               "cp01,cp02,cp03,cp04,substr(CPM03,1,6) CPM03,decode(sign(a0k02),1,a0k02,a1k02) as a1k02,cp14,cp05,nvl(s3.st02,a.st02) sname,nvl(s3.st15,a.ST15) sst15,' ' 點數小計,cp60 " & _
               "From CASEPROGRESS,acc1k0,acc1n0 p1,ACC090,CASEPROPERTYMAP,staff a,staff b,acc0k0,acc1n0 p2,staff s3 " & _
               "WHERE (substr(cp60,1,1) ='X' or substr(cp60,1,1) ='E' or cp18>0) AND cp14 is not null " & _
               "AND p1.a1n01(+)=cp60 AND p1.a1n02(+)='2' AND p1.a1n03(+)=cp09 AND p1.a1n04(+)<>cp14 and p1.a1n05(+)>0 " & _
               "AND p2.a1n02(+)='3' AND p2.a1n03(+)=cp09 AND p2.a1n04(+)<>cp14 and p2.a1n05(+)>0 " & _
               "AND a1k01(+)=cp60 AND CP01=CPM01(+) AND CP10=CPM02(+) AND a0k01(+)=cp60 " & _
               "AND CP12=A0901(+) AND p1.a1n04=a.st01(+) AND cp14=b.st01(+) and p2.a1n04=s3.st01(+) " & strConSql & _
               IIf(Text1(3) <> "", " and (p1.a1n04=" & CNULL(Trim(Text1(3))) & " or p2.a1n04=" & CNULL(Trim(Text1(3))) & ") ", "") & _
               ") X1 where n03>'A' order By sst15,n04,a1k02,cp01,cp02,cp03,cp04"
                     
   '依智權人員統計 AND substr(a.st15,1,2)=substr(b.st15,1,2)
   Else
      'Modified by Lydia 2016/12/13 +承辦人條件 IIf(Text1(3) <> "", " and a1n04=" & CNULL(Trim(Text1(3))), "")
      'Modified by Lydia 2020/04/17 拿掉判斷銷帳=>AND a1k25 is null ,and a0k10 is null
      strSql = "SELECT NVL(A0902,A0903) tt,a1n04,a1n03,a1n05,cp01,cp02,cp03,cp04,CPM03" & _
               ",decode(sign(a0k02),1,a0k02,a1k02) as a1k02,a1k02,cp13," & IIf(Trim(Text1(0).Text) = "1", "cp05", "cp27") & ",a.st02,a.ST15,' ' 點數小計,cp60 " & _
                     "From CASEPROGRESS,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff a,staff b,acc0k0 " & _
                     "WHERE (substr(cp60,1,1) ='X' or substr(cp60,1,1) ='E') AND cp13 is not null " & _
                     "AND a1n01(+)=cp60 AND a1n02(+)='1' AND a1n03(+)=cp09 AND a1n04(+)<>cp13 and a1n05>0 " & _
                     "AND a1k01(+)=cp60 AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
                     "AND a0k01(+)=cp60 " & _
                     "AND CP12=A0901(+) AND a1n04=a.st01(+) AND cp13=b.st01(+) " & strConSql & _
                     IIf(Text1(3) <> "", " and a1n04=" & CNULL(Trim(Text1(3))), "") & _
                     " Order By a.ST15,a1n04,a1k02,cp01,cp02,cp03,cp04 "
   End If
   intI = 1
   'Modified by Lydia 2016/12/13
   'Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
   Set rsA1 = ClsLawReadRstMsg(intI, strSql)
   
   If intI = 1 Then
      intCompRow = 1 '代表有分配點數資料,從第1筆開始讀取
      'Modified by Lydia 2016/12/13
      'dblPointRow = adoRecordset.RecordCount 'Add By Sindy 2012/2/3
      'Set grd1.Recordset = adoRecordset.Clone
      dblPointRow = rsA1.RecordCount
      Set grd1.Recordset = rsA1.Clone
      
      'Modify By Sindy 2012/2/3
'      '依承辦人統計
'      If opt1(1).Value = True Then
'         strSql = "SELECT sum(a1n05) tt " & _
'                        "From CASEPROGRESS,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff " & _
'                        "WHERE cp60>'X' AND cp14 is not null " & _
'                        "AND a1n01(+)=cp60 AND a1n02(+)='2' AND a1n03(+)=cp09 AND a1n04(+)<>cp14 and a1n05>0 " & _
'                        "AND a1k01(+)=cp60 AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
'                        "AND a1k25 is null " & _
'                        "AND CP12=A0901(+) AND a1n04=st01(+) " & strConSql
'      '依智權人員統計
'      Else
'         strSql = "SELECT sum(a1n05) tt " & _
'                        "From CASEPROGRESS,acc1k0,acc1n0,ACC090,CASEPROPERTYMAP,staff " & _
'                        "WHERE cp60>'X' AND cp13 is not null " & _
'                        "AND a1n01(+)=cp60 AND a1n02(+)='1' AND a1n03(+)=cp09 AND a1n04(+)<>cp13 and a1n05>0 " & _
'                        "AND a1k01(+)=cp60 AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
'                        "AND a1k25 is null " & _
'                        "AND CP12=A0901(+) AND a1n04=st01(+) " & strConSql
'      End If
'      intI = 1
'      Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         dblPointTot = adoRecordset.Fields(0)
'      End If
      '計算分配點數的小計及總計
      For i = 1 To grd1.Rows - 1
         If strEmp <> grd1.TextMatrix(i, 1) Then
            If strEmp <> "" Then
               For j = 1 To grd1.Rows - 1
                  If grd1.TextMatrix(j, 1) = strEmp Then
                     grd1.TextMatrix(j, 14) = dblPointTotSub
                  End If
               Next j
            End If
            strEmp = grd1.TextMatrix(i, 1)
            dblPointTotSub = 0
         End If
         dblPointTotSub = dblPointTotSub + Val(grd1.TextMatrix(i, 3))
         dblPointTot = dblPointTot + Val(grd1.TextMatrix(i, 3))
      Next i
      If strEmp <> "" Then
         For j = 1 To grd1.Rows - 1
            If grd1.TextMatrix(j, 1) = strEmp Then
               grd1.TextMatrix(j, 14) = dblPointTotSub
            End If
         Next j
      End If
      '2012/2/3 End
   End If
End Sub

'Add By Sindy 2010/5/10
'ST15+CP14
Function PrintEndPoint(strComp1 As String, strNext1 As String) As Boolean
Dim ii As Integer
Dim bNextTrue As Boolean

   aX(1) = 500: aX(2) = 1700: aX(3) = 2700: aX(4) = 4700: aX(5) = 7200

   PrintEndPoint = False
   If intCompRow = 0 Or intCompRow > (grd1.Rows - 1) Then Exit Function

   '處理傳進來比對人員的資料
   'If grd1.TextMatrix(intCompRow, 1) = strComp1 Then
   If grd1.TextMatrix(intCompRow, 13) & grd1.TextMatrix(intCompRow, 1) = strComp1 Then
      PrintEndPoint = True
      m_strTemp1 = strComp1
      m_strTemp3 = Trim(grd1.TextMatrix(intCompRow, 12))
      'Added by Lydia 2020/04/10 分配人員非法律所人員請姓名後加*
      If Trim(grd1.TextMatrix(intCompRow, 1)) <> "" Then
         If PUB_ChkLCompStaff(grd1.TextMatrix(intCompRow, 1)) = False Then
             m_strTemp3 = m_strTemp3 & "*"
         End If
      End If
      'end 2020/04/10
      iPrint = iPrint + 300
      If iPrint >= 11000 Then
          Page = Page + 1
          Printer.NewPage
          CaseTitle
      End If

      Printer.CurrentX = aX(1)
      Printer.CurrentY = iPrint
      'Modified by Lydia 2015/05/08 +收據資料,位置調整成陣列aX( )儲存
      'Modified by Lydia 2015/06/02 改抬頭
      Printer.Print "請款單/工作點數分配：" & m_strTemp3
      iPrint = iPrint + 300
      'Modified by Lydia 2015/05/08 +收據資料
      Printer.CurrentX = aX(1)
      Printer.CurrentY = iPrint
      Printer.Print "單據號碼"
      
      Printer.CurrentX = aX(2)
      Printer.CurrentY = iPrint
      Printer.Print "日　期"
      Printer.CurrentX = aX(3)
      Printer.CurrentY = iPrint
      Printer.Print "本所案號"
      Printer.CurrentX = aX(4)
      Printer.CurrentY = iPrint
      Printer.Print "案件性質"
      Printer.CurrentX = aX(5) - Printer.TextWidth("點數")
      'end 2015/05/08
      Printer.CurrentY = iPrint
      Printer.Print "點數"
      iPrint = iPrint + 300
      ShowLine1
      dblSum = 0
      For ii = intCompRow To grd1.Rows - 1
         'If grd1.TextMatrix(intCompRow, 1) = strComp1 Then
         If grd1.TextMatrix(intCompRow, 13) & grd1.TextMatrix(intCompRow, 1) = strComp1 Then
            m_strTemp1 = strComp1
            m_strTemp3 = Trim(grd1.TextMatrix(intCompRow, 12))
            If iPrint >= 11000 Then
                Page = Page + 1
                Printer.NewPage
                CaseTitle
            End If
            
            'Modified by Lydia 2015/05/08 +收據資料,位置調整成陣列aX( )儲存
            Printer.CurrentX = aX(1)
            Printer.CurrentY = iPrint
            'Modified by Lydia 2015/06/05 改收文號
            'Printer.Print GRD1.TextMatrix(intCompRow, 15)
            Printer.Print grd1.TextMatrix(intCompRow, 2)
            
            Printer.CurrentX = aX(2)
            Printer.CurrentY = iPrint
            Printer.Print ChangeTStringToTDateString(grd1.TextMatrix(intCompRow, 9))
            Printer.CurrentX = aX(3)
            Printer.CurrentY = iPrint
            Printer.Print grd1.TextMatrix(intCompRow, 4) & "-" & grd1.TextMatrix(intCompRow, 5) & "-" & grd1.TextMatrix(intCompRow, 6) & "-" & grd1.TextMatrix(intCompRow, 7)
            Printer.CurrentX = aX(4)
            Printer.CurrentY = iPrint
            Printer.Print grd1.TextMatrix(intCompRow, 8)
            
            'Printer.CurrentX = aX(5) - Printer.TextWidth(CheckStr(Val(GRD1.TextMatrix(intCompRow, 3))))
            Printer.CurrentX = aX(5) - Printer.TextWidth(Format(Val(grd1.TextMatrix(intCompRow, 3)), "##,##0.000"))
            Printer.CurrentY = iPrint
           ' Printer.Print CheckStr(Val(GRD1.TextMatrix(intCompRow, 3)))
            Printer.Print Format(Val(grd1.TextMatrix(intCompRow, 3)), "##,##0.000")
            dblSum = dblSum + Val(grd1.TextMatrix(intCompRow, 3))
            'end 2015/05/08
            
            iPrint = iPrint + 300
            intCompRow = intCompRow + 1
         Else
            Exit For
         End If
      Next ii
      PrintEnd2
   End If

   '處理傳進來下一人員(之前人員)的資料
   m_strTemp1 = "": m_strTemp3 = "": bNextTrue = False: dblSum = 0
   For ii = intCompRow To grd1.Rows - 1

      'If (Val(grd1.TextMatrix(ii, 1)) >= Val(strNext1) And Val(strNext1) <> 0) Then GoTo GoToExit
      If (grd1.TextMatrix(ii, 13) & grd1.TextMatrix(ii, 1) >= strNext1 And Len(strNext1) <> 0) Then GoTo gotoExit
      
      'If (grd1.TextMatrix(ii, 1) <> m_strTemp1) Then
      If (grd1.TextMatrix(ii, 13) & grd1.TextMatrix(ii, 1) <> m_strTemp1) Then
         PrintEndPoint = True
         bNextTrue = True
         PrintEnd2
         dblSum = 0
         'one new data start
         'm_strTemp1 = Trim(grd1.TextMatrix(ii, 1))
         m_strTemp1 = Trim(grd1.TextMatrix(ii, 13)) & Trim(grd1.TextMatrix(ii, 1))
         m_strTemp3 = Trim(grd1.TextMatrix(ii, 12))
         'Added by Lydia 2020/04/10 分配人員非法律所人員請姓名後加*
         If Trim(grd1.TextMatrix(ii, 1)) <> "" Then
            If PUB_ChkLCompStaff(grd1.TextMatrix(ii, 1)) = False Then
                m_strTemp3 = m_strTemp3 & "*"
            End If
         End If
         'end 2020/04/10
         iPrint = iPrint + 300
         If iPrint >= 11000 Then
            Page = Page + 1
            Printer.NewPage
            CaseTitle
         End If

         Printer.CurrentX = aX(1)
         Printer.CurrentY = iPrint
         'Modified by Lydia 2015/05/08 +收據資料,位置調整成陣列aX( )儲存
         'Modified by Lydia 2015/06/02 改抬頭
         Printer.Print "請款單/工作點數分配：" & m_strTemp3
         iPrint = iPrint + 300
         Printer.CurrentX = aX(1)
         Printer.CurrentY = iPrint
         Printer.Print "單據號碼"
         
         Printer.CurrentX = aX(2)
         Printer.CurrentY = iPrint
         'Printer.Print "請款日"
         Printer.Print "日　期"
         Printer.CurrentX = aX(3)
         Printer.CurrentY = iPrint
         Printer.Print "本所案號"
         Printer.CurrentX = aX(4)
         Printer.CurrentY = iPrint
         Printer.Print "案件性質"
         Printer.CurrentX = aX(5) - Printer.TextWidth("點數")
         'end 2015/05/08
         Printer.CurrentY = iPrint
         Printer.Print "點數"
         iPrint = iPrint + 300
         ShowLine1
      End If
      If iPrint >= 11000 Then
         Page = Page + 1
         Printer.NewPage
         CaseTitle
      End If
      'Modified by Lydia 2015/05/08 +收據資料,位置調整成陣列aX( )儲存
      Printer.CurrentX = aX(1)
      Printer.CurrentY = iPrint
      'Modified by Lydia 2015/06/05 改收文號
      'Printer.Print GRD1.TextMatrix(intCompRow, 15)
      Printer.Print grd1.TextMatrix(intCompRow, 2)
      
      Printer.CurrentX = aX(2)
      Printer.CurrentY = iPrint
      Printer.Print ChangeTStringToTDateString(grd1.TextMatrix(ii, 9))
      Printer.CurrentX = aX(3)
      Printer.CurrentY = iPrint
      Printer.Print grd1.TextMatrix(ii, 4) & "-" & grd1.TextMatrix(ii, 5) & "-" & grd1.TextMatrix(ii, 6) & "-" & grd1.TextMatrix(ii, 7)
      Printer.CurrentX = aX(4)
      Printer.CurrentY = iPrint
      Printer.Print grd1.TextMatrix(ii, 8)
    '  Printer.CurrentX = aX(5) - Printer.TextWidth(CheckStr(Val(GRD1.TextMatrix(ii, 3))))
      Printer.CurrentX = aX(5) - Printer.TextWidth(Format(Val(grd1.TextMatrix(ii, 3)), "##,##0.000"))
      Printer.CurrentY = iPrint
    '  Printer.Print CheckStr(Val(GRD1.TextMatrix(ii, 3)))
      Printer.Print Format(Val(grd1.TextMatrix(ii, 3)), "##,##0.000")
      dblSum = dblSum + Val(grd1.TextMatrix(ii, 3))
      'end 2015/05/08
      
      iPrint = iPrint + 300
      intCompRow = intCompRow + 1
   Next ii
gotoExit:
   If bNextTrue = True Then
      PrintEnd2
   End If
End Function

'Add By Sindy 2010/5/10
Sub ShowLine1()
   Printer.CurrentX = 500:               Printer.CurrentY = iPrint
   'Modified by Lydia 2015/05/08
   'Printer.Print String(80, "-")
   Printer.Print String(95, "-")
   iPrint = iPrint + 300
   If iPrint >= 11000 Then
      Page = Page + 1
      Printer.NewPage
      CaseTitle
   End If
End Sub

'Add By Sindy 2010/5/11
Sub PrintEnd1()
   If dblPointTot > 0 Then
      If iPrint >= 11000 Then
         Page = Page + 1
         Printer.NewPage
         CaseTitle
      End If
      Printer.CurrentX = PLeft(6) - 200:  Printer.CurrentY = iPrint
      Printer.Print "分配點數"
      Printer.CurrentX = PLeft(8) - 200:  Printer.CurrentY = iPrint
      Printer.Print dblPointTot & " 點"
      iPrint = iPrint + 300
      'Add By Sindy 2010/10/5
      ShowLine
      If iPrint >= 11000 Then
         Page = Page + 1
         Printer.NewPage
         CaseTitle
      End If
      Printer.CurrentX = PLeft(6) - 200:  Printer.CurrentY = iPrint
      Printer.Print "總計"
      Printer.CurrentX = PLeft(8) - 200:  Printer.CurrentY = iPrint
      Printer.Print StrToStr(CStr(Format(m_totCP18, FatDot) + dblPointTot), 8) & " 點"
      iPrint = iPrint + 300
      '2010/10/5 End
   End If
End Sub

'Add By Sindy 2010/5/10
Sub PrintEnd2()
   If dblSum > 0 Then
      ShowLine1
      If iPrint >= 11000 Then
         Page = Page + 1
         Printer.NewPage
         CaseTitle
      End If
'Modified by Lydia 2015/05/08 +收據資料,位置調整成陣列aX( ) 儲存
      Printer.CurrentX = aX(4)
      Printer.CurrentY = iPrint
      Printer.Print "合計"
'      Printer.CurrentX = 6000 - Printer.TextWidth(CheckStr(dblSum))
      Printer.CurrentX = aX(5) - Printer.TextWidth(Format(dblSum, "##,##0.000"))
      Printer.CurrentY = iPrint
      'Printer.Print CheckStr(dblSum)
      Printer.Print Format(dblSum, "##,##0.000")
'end 2015/05/08
      iPrint = iPrint + 300
   End If

End Sub

'Add By Sindy 2012/2/3 取得此人員的分配點數小計
Sub GetPointTotSub(strEmp As String)
   dblPointTotSub = 0
   If dblPointRow > 0 Then
      For i = 1 To grd1.Rows - 1
         If strEmp = grd1.TextMatrix(i, 1) Then
            dblPointTotSub = grd1.TextMatrix(i, 14)
            Exit Sub
         End If
      Next i
   End If
End Sub
