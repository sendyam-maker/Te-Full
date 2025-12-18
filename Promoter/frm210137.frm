VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210137 
   BorderStyle     =   1  '單線固定
   Caption         =   "業績點數統計"
   ClientHeight    =   5310
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   7070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7070
   Begin VB.CheckBox Check1 
      Caption         =   "入帳點數包含轉撥點數資料"
      Height          =   180
      Left            =   240
      TabIndex        =   14
      Top             =   1250
      Width           =   2500
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "收文點數明細(&B)"
      Height          =   400
      Left            =   5085
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   720
      Width           =   1800
   End
   Begin VB.CommandButton cmdWord 
      Caption         =   "Word(&W)"
      Height          =   400
      Left            =   5085
      TabIndex        =   10
      Top             =   180
      Width           =   930
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   4200
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   180
      Width           =   795
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   6090
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   180
      Width           =   800
   End
   Begin VB.TextBox txtSalesArea 
      Height          =   285
      Left            =   1260
      TabIndex        =   0
      Top             =   180
      Width           =   915
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   285
      Index           =   0
      Left            =   1260
      MaxLength       =   7
      TabIndex        =   3
      Top             =   840
      Width           =   915
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   285
      Index           =   1
      Left            =   2520
      MaxLength       =   7
      TabIndex        =   4
      Top             =   840
      Width           =   915
   End
   Begin VB.TextBox txtSales 
      Height          =   285
      Left            =   1260
      MaxLength       =   6
      TabIndex        =   2
      Top             =   510
      Width           =   915
   End
   Begin VB.TextBox txtSalesArea1 
      Height          =   285
      Left            =   2475
      TabIndex        =   1
      Top             =   180
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3615
      Left            =   180
      TabIndex        =   11
      Top             =   1550
      Width           =   6705
      _ExtentX        =   11836
      _ExtentY        =   6368
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   10
      FixedCols       =   0
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      MergeCells      =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.Label lblSalesName 
      Height          =   255
      Left            =   2250
      TabIndex        =   15
      Top             =   510
      Width           =   1290
      VariousPropertyBits=   27
      Size            =   "2275;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   2205
      X2              =   2475
      Y1              =   990
      Y2              =   990
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   2205
      X2              =   2475
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Label lblCaution 
      AutoSize        =   -1  'True
      Caption         =   "入帳點數已扣除財務扣點數"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   3585
      TabIndex        =   12
      Top             =   690
      Width           =   1380
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "業務區"
      Height          =   180
      Left            =   225
      TabIndex        =   7
      Top             =   225
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "點數結算日"
      Height          =   180
      Left            =   225
      TabIndex        =   6
      Top             =   885
      Width           =   900
   End
   Begin VB.Label Label4 
      Caption         =   "智權人員"
      Height          =   180
      Left            =   225
      TabIndex        =   5
      Top             =   555
      Width           =   900
   End
End
Attribute VB_Name = "frm210137"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/25 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、lblSalesName
'Memo by Lydia 2019/07/01 表單名稱:各區業績點數統計=>業績點數統計
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Created by Morgan 2012/6/11
Option Explicit
 
Dim i As Integer, j As Integer, lngColar As Long, bolSelData As Boolean  'Add by Amy 2012/04/24
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
Dim m_strListPer As String 'Add By Sindy 2020/7/1
Dim m_PrevForm As Form 'Added by Lydia 2021/07/27 前一畫面 'Memo by Lydia 2021/08/27 上線
Dim strField() As String 'Add by Amy 2024/04/03


'Added by Lydia 2021/07/27 外部呼叫使用
Public Sub SetParent(ByVal pForm As Form)
   Set m_PrevForm = pForm
End Sub

'Add By Amy 2012/04/24
Private Sub cmdDetail_Click()
    PubShowNextData
    Exit Sub
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

'Modify By Sindy 2020/8/31
'Private Sub cmdSearch_Click()
Public Sub cmdSearch_Click()
'2020/8/31 END
   Screen.MousePointer = vbHourglass
   If ConstrainCheck = True Then
      Call SetDataListWidth
      Call doQuery
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdWord_Click()
   Screen.MousePointer = vbHourglass
   runWord
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   Dim stST05 As String, stST15 As String
   Dim bolQSelf As Boolean
   
   bolSelData = False 'Add by Amy 2013/04/24
   MoveFormToCenter Me
   SetDataListWidth , True 'Modify by Amy 2024/04/03 +True
   
   stST15 = PUB_GetStaffST15(strUserNum, 1)
   stST05 = PUB_GetST05(strUserNum)
   
   'Modify By Sindy 2020/7/28 設定員編,部門,所別權限
   Call PUB_SetFormSaleDept(strUserNum, , txtSalesArea, txtSalesArea1, txtSales, bolSpecMan, strSpecCode, , , bolQSelf)
   'Modify By Sindy 2021/10/12 + And txtSales.Enabled = False
   If bolQSelf = True And txtSales.Enabled = False Then
      Me.Caption = "個人業績點數統計"
   End If
   
'   'Add by Amy 2013/05/23 開放所有人可使用並控管權限
'   txtSalesArea.Enabled = False
'   txtSalesArea1.Enabled = False
'   txtSales.Enabled = False
'
'   Select Case strUserNum
'      'Add by Amy 2013/05/23 copy frm210104 權限控管(蔣律師及分所財務登入改只能看自己的)
'      '杜燕文,劉大愛可看S31
'      'modify by sonia 2016/8/22劉大愛78007改為蘇嫄媛79053
'      Case "74018", "79053"
'         txtSalesArea = "S31"
'         txtSalesArea1 = "S31"
'         txtSales.Enabled = True
'      '王協理可看專利處
'      Case "71011"
'         txtSalesArea = "P10"
'         txtSalesArea1 = "P19"
'         txtSales.Enabled = True
'      '葉經理可看商標處
'      'modify by sonia 2016/2/24 +69008
'      Case "67002", "69008"
'         txtSalesArea = "P20"
'         txtSalesArea1 = "P29"
'         txtSales.Enabled = True
'      '何副總可看國外部
'      'modify by sonia 2015/6/30 改68009為81040
'      Case "81040"
'         txtSalesArea = "F10"
'         txtSalesArea1 = "F41"
'         txtSales.Enabled = True
'      Case "71003"
'        txtSalesArea = "S23"
'        txtSalesArea1 = "S23"
'        txtSales.Enabled = True
'      'end
'
'      '小真,杜副總,林特助可看全部
'      'modify by sonia 2014/6/9 +美珍77027,並取消94007(因己改個人等級)
'      'modify by sonia 2019/12/27 杜主秘說加開簡協理69005可看全所
'      Case "65001", "68006", "77027", "69005"
'         txtSalesArea.Enabled = True
'         txtSalesArea1.Enabled = True
'         txtSales.Enabled = True
'
''cancel by sonia 2019/12/27 杜主秘說加開簡協理69005可看全所
''      '69005可看北所全部
''      Case "69005"
''         txtSalesArea.Enabled = True
''         txtSalesArea1.Enabled = True
''         txtSalesArea = stST15
''         txtSalesArea1 = stST15
''         txtSales.Enabled = True
''end 2019/12/27
'      'add by sonia 2016/12/21 柄佑可看中所全部但業務區仍預設自已部門
'      Case "82026"
'         txtSalesArea.Enabled = True
'         txtSalesArea1.Enabled = True
'         txtSales.Enabled = True
'         txtSalesArea = stST15
'         txtSalesArea1 = stST15
'         txtSales = strUserNum
'      'end 2016/12/21
'      Case Else
'         Select Case stST05
'            '電腦中心,財務,總經理看全部
'            '2015/7/28 MODIFY BY SONIA +主任秘書(等級08)可看全部
'            Case "00", "01", "08"
'               txtSalesArea.Enabled = True
'               txtSalesArea1.Enabled = True
'               txtSales.Enabled = True
'            'Add by Amy 2013/05/23
'            '各區主管
'            Case "SM"
'               txtSalesArea.Locked = True
'               txtSalesArea1.Locked = True
'               txtSalesArea = stST15
'               txtSalesArea1 = stST15
'               txtSales.Enabled = True
'            'end 2013/05/23
'
'            '其他只能看自己
'            Case Else
'               'txtSalesArea.Enabled = False 'Modify By Amy 2013/05/23
'               'txtSalesArea1.Enabled = False 'Modify By Amy 2013/05/23
'               txtSalesArea = stST15
'               txtSalesArea1 = stST15
'               'Add by Amy 2013/05/23
'               txtSales = strUserNum
'               Me.Caption = "個人業績點數統計"
'               'end 2013/05/23
'         End Select
'   End Select
'
'   'Add by Amy 2013/05/23
'   '若操作人員的ST05=SA且在職員工的ST52有該編號存在,此類人員稱之為帶人主管,開放智權人員欄位可輸入
'   If PUB_GetST05Limits(strUserNum) = True Then
'      txtSales.Enabled = True
'   End If
'   'end 2013/05/23
'
'   'Add By Sindy 2016/5/6 記錄原操作人可以查詢的業務區及所別
'   'txtZone.Tag = txtZone
'   txtSalesArea.Tag = txtSalesArea
'   txtSalesArea1.Tag = txtSalesArea1
'   '2016/5/6 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Added by Lydia 2021/07/27 回前一畫面
   If TypeName(m_PrevForm) <> "Nothing" Then
       m_PrevForm.Show
   End If
   'end 2021/07/27
   
   Set frm210137 = Nothing
End Sub

Private Function doQuery() As Boolean
   Dim stCon As String, stConST As String, stConDF As String
   Dim stCon020 As String, stCon021 As String, stConWD As String, stConWD1 As String, stConPE As String, stConCP As String
   Dim stVTB1 As String, stVTB2 As String, stVTB3 As String, stVTB4 As String
   Dim str_ax212 As String   'add by sonia 2017/1/9
   
   '業務區
   If txtSalesArea <> "" Then
      stConST = stConST & " and st15>='" & txtSalesArea & "'"
      stConCP = stConCP & " and st15||''>='" & txtSalesArea & "'" 'Modify By Sindy 2021/8/4 CP12 => st15
   End If
   If txtSalesArea1 <> "" Then
      stConST = stConST & " and st15<='" & txtSalesArea1 & "'"
      stConCP = stConCP & " and st15||''<='" & txtSalesArea1 & "'" 'Modify By Sindy 2021/8/4 CP12 => st15
   End If
   
   '智權人員
   If txtSales <> "" Then
      stConST = stConST & " and st01='" & txtSales & "'"
      stCon021 = stCon021 & " and ax209='" & txtSales & "'"
      stConCP = stConCP & " and cp13='" & txtSales & "'"
   End If
   
   '點數結算日
   If txtCloseDate(0) <> "" Then
      stConDF = stConDF & " and df02 >=" & txtCloseDate(0)
      stCon020 = stCon020 & " and a0205 >=" & txtCloseDate(0)
      stConWD = stConWD & " and wd01 >=" & DBDATE(txtCloseDate(0))
      stConWD1 = stConWD1 & " and wd01 >=" & (DBDATE(txtCloseDate(0)) \ 100) & "01"
      stConPE = stConPE & " and pe03 >=" & (DBDATE(txtCloseDate(0)) \ 100)
      stConCP = stConCP & " and cp05>=" & DBDATE(txtCloseDate(0))
   End If
   If txtCloseDate(1) <> "" Then
      stConDF = stConDF & " and df02 <=" & txtCloseDate(1)
      stCon020 = stCon020 & " and a0205 <=" & txtCloseDate(1)
      stConWD = stConWD & " and wd01 <=" & DBDATE(txtCloseDate(1))
      stConWD1 = stConWD1 & " and wd01 <=" & (DBDATE(txtCloseDate(1)) \ 100) & "31"
      stConPE = stConPE & " and pe03 <=" & DBDATE(txtCloseDate(1)) \ 100
      stConCP = stConCP & " and cp05<=" & DBDATE(txtCloseDate(1))
   End If
   
   '簽約點數
   '2012/10/16 modify by sonia 員工資料改在目標檔抓
   'stVTB1 = "select st06,st15,st01,st02,sum(df04) tot" & _
      " From staff, DailyFeat" & _
      " where DF01(+)=ST01" & stConST & stConDF & _
      " group by st06,st15,st01,st02"
   stVTB1 = "select st01,sum(df04) tot" & _
      " From staff, DailyFeat" & _
      " where DF01(+)=ST01" & stConST & stConDF & _
      " group by st01"

   '入帳點數(扣除扣點數)
   'MODIFY BY SONIA 2014/1/21 取消a0201='1'條件
   'stVTB2 = "select ax209 ,ROUND(sum(ax207-ax206)/1000,2) V1C2" & _
      " From acc020, acc021,STAFF Where a0201='1' " & stCon020 & " and ax201(+) = a0201" & _
      " and ax202(+) = a0202 and ax209 Is Not Null" & stCon021 & _
      " and (substr(ax205, 1, 2) = '41' Or ax205 = '7121')" & _
      " and not( ax205='4191' or ax205='4192' or ((ax205='410103' or ax205='411103'" & _
      " or ax205='4121' or ax205='4131') and instr(ax213||' ','結餘')>0)) AND ST01(+)=AX209" & stConST & _
      " GROUP BY AX209"
   'modify by sonia 2015/4/23 加不含'4194'科目,不含結餘傳票不要加科目限制
   'stVTB2 = "select ax209 ,ROUND(sum(ax207-ax206)/1000,2) V1C2" & _
      " From acc020, acc021,STAFF Where ax201(+) = a0201 and ax202(+) = a0202 " & stCon020 & " and ax209 Is Not Null" & stCon021 & _
      " and (substr(ax205, 1, 2) = '41' Or ax205 = '7121')" & _
      " and not( ax205='4191' or ax205='4192' or ((ax205='410103' or ax205='411103'" & _
      " or ax205='4121' or ax205='4131') and instr(ax213||' ','結餘')>0)) AND ST01(+)=AX209" & stConST & _
      " GROUP BY AX209"
   'modify by sonia 2017/1/9 是否包含轉撥點數資料
   'stVTB2 = "select ax209 ,ROUND(sum(ax207-ax206)/1000,2) V1C2" & _
      " From acc020, acc021,STAFF Where ax201(+) = a0201 and ax202(+) = a0202 " & stCon020 & " and ax209 Is Not Null" & stCon021 & _
      " and (substr(ax205, 1, 2) = '41' Or ax205 = '7121')" & _
      " and not( ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0 or instr(ax212||' ','轉撥')>0) AND ST01(+)=AX209" & stConST & _
      " GROUP BY AX209"
   If Check1.Value = 0 Then
      str_ax212 = " or instr(ax212||' ','轉撥')>0"
   Else
      str_ax212 = ""
   End If
   'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
   stVTB2 = "select ax209 ,ROUND(sum(ax207-ax206)/1000,2) V1C2" & _
      " From acc020, acc021,STAFF Where ax201(+) = a0201 and ax202(+) = a0202 " & stCon020 & " and ax209 Is Not Null" & stCon021 & _
      " and (substr(ax205, 1, 1) = '4' Or ax205 = '7121')" & _
      " and not( ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0" & str_ax212 & ") AND ST01(+)=AX209" & stConST & _
      " GROUP BY AX209"
   'end 2017/1/9
   '2014/1/21 END
      
   '目標點數
   stVTB3 = "select pe01,st06,st15,st02,round(sum(pe04*y2/x2),2) z1" & _
      " from performance,staff,(select substr(wd01,1,6) x1,count(*) x2" & _
      " from workday where 1=1" & stConWD1 & _
      " group by substr(wd01,1,6)) x,(select substr(wd01,1,6) y1,count(*) y2" & _
      " from workday where 1=1" & stConWD & _
      " group by substr(wd01,1,6)) y" & _
      " where pe02='TOT'" & stConPE & " and st01(+)=pe01" & stConST & _
      " and x1(+)=pe03 and y1(+)=pe03 group by pe01,st06,st15,st02"

   '收文點數
   'Modify by Amy 2018/04/30 銷案不銷帳仍要計算,故拿掉cp57 is null -秀玲 P-122345 AA8013150
'   stVTB4 = "select cp13,sum(cp18-nvl(a1u07,0)/1000) RecPt" & _
'      " From (select cp13,cp09,cp18 from caseprogress where cp57 is null" & stConCP & ")" & _
'      ",(select a1u03,sum(a1u07) a1u07 from caseprogress,acc1u0 where cp57 is null and a1u03(+)=cp09 and a1u07>0" & stConCP & " group by a1u03)" & _
'      " where a1u03(+)=cp09 group by cp13"
    'Modify By Sindy 2021/8/4 + staff, cp13=st01(+)
    'Modify by Amy 2022/08/02 +And ((cp159>0 And cp60 is not null) Or cp159=0),要與 PUB_CountCP18的結果相同-秀玲
    'Modify by Amy 2022/08/03 RecPt值的算法要與 PUB_CountCP18一致
'    stVTB4 = "select cp13,sum(cp18-nvl(a1u07,0)/1000) RecPt" & _
'      " From (select cp13,cp09,cp18 from caseprogress,staff where cp13=st01(+) And ((cp159>0 And cp60 is not null) Or cp159=0) " & stConCP & ")" & _
'      ",(select a1u03,sum(a1u07) a1u07 from caseprogress,acc1u0,staff where a1u03(+)=cp09 and a1u07>0 and cp13=st01(+) And ((cp159>0 And cp60 is not null) Or cp159=0) " & stConCP & " group by a1u03)" & _
'      " where a1u03(+)=cp09 group by cp13 Having sum(cp18-nvl(a1u07,0)/1000)<>0"
      'Modify by Amy 2024/04/19 改抓共用函數
'      stVTB4 = "select cp13,Sum(((Nvl(CP16,0)-NVL(A1U07,0)-NVL(A1U09,0))-(NVL(CP17,0)-NVL(A1U09,0)))/1000) RecPt" & _
'      " From (select cp13,cp09,cp18,Nvl(cp16,0) cp16,Nvl(cp17,0) cp17 from caseprogress,staff where cp13=st01(+) And ((cp159>0 And cp60 is not null) Or cp159=0) " & stConCP & ")" & _
'      ",(select a1u03,sum(Nvl(a1u07,0)) a1u07,Sum(Nvl(a1u09,0)) a1u09 from caseprogress,acc1u0,staff where a1u03(+)=cp09 and a1u07>0 and cp13=st01(+) And ((cp159>0 And cp60 is not null) Or cp159=0) " & stConCP & " group by a1u03)" & _
'      " where a1u03(+)=cp09 group by cp13 Having Sum(((Nvl(CP16,0)-NVL(A1U07,0)-NVL(A1U09,0))-(NVL(CP17,0)-NVL(A1U09,0)))/1000)<>0"
   strExc(2) = PUB_CountCP18(0, txtCloseDate(0), txtCloseDate(1), txtSalesArea, txtSalesArea1, txtSales, stVTB4, , , 0, Me.Name, , True, 3)
      
   '2012/10/16 modify by sonia 以目標檔串其他資料,否則未輸每日業績前當日收文點數抓不到
   'strExc(0) = "select decode(st06,'1','北','2','中','3','南','4','高')||'所' 所別" & _
      ",a0902 業務區,st02 智權人員,to_char(tot,'99999.00') 簽約點數,to_char(RecPt,'99999.00') 收文點數,to_char(V1C2,'99999.00') 入帳點數,to_char(Z1,'99999.00') 目標點數" & _
      " from (" & stVTB1 & ") x, (" & stVTB2 & ") y,(" & stVTB3 & ") z,(" & stVTB4 & ") w,acc090" & _
      " where a0901(+)=st15 and ax209(+)=st01 and pe01(+)=st01 and cp13(+)=st01 order by st06,st15,st01"
   '2013/04/24 select 加'' V及智權人員編號
   'Modify by Amy 20130912 +join staff (因DailyFeat抓不到資料導致st01為空)
   'strExc(0) = "select '' V,decode(st06,'1','北','2','中','3','南','4','高')||'所' 所別" & _
      ",a0902 業務區,st02 智權人員,to_char(tot,'99999.00') 簽約點數,to_char(RecPt,'99999.00') 收文點數,to_char(V1C2,'99999.00') 入帳點數,to_char(Z1,'99999.00') 目標點數,ST01,ST15" & _
      " from (" & stVTB1 & ") x, (" & stVTB2 & ") y,(" & stVTB3 & ") z,(" & stVTB4 & ") w,acc090" & _
      " where a0901(+)=st15 and ax209(+)=pe01 and pe01=st01(+) and cp13(+)=pe01 order by st06,st15,st01"
  '2015/4/29 MODIFY BY SONIA 智權人員名單抓有輸入點數或有財務點數的,例D104012494李羽宸84045
  'strExc(0) = "select '' V,decode(S.st06,'1','北','2','中','3','南','4','高')||'所' 所別" & _
      ",a0902 業務區,S.st02 智權人員,to_char(tot,'99999.00') 簽約點數,to_char(RecPt,'99999.00') 收文點數,to_char(V1C2,'99999.00') 入帳點數,to_char(Z1,'99999.00') 目標點數,S.ST01,S.ST15" & _
      " from Staff S,(" & stVTB1 & ") x, (" & stVTB2 & ") y,(" & stVTB3 & ") z,(" & stVTB4 & ") w,acc090" & _
      " where a0901(+)=Z.st15 and ax209(+)=pe01 and pe01=X.st01(+) and cp13(+)=pe01 AND PE01=S.ST01(+) order by S.st06,S.st15,S.st01"
  'Modify by 2024/04/19 改抓共用函數 原:RecPt->收文點數
  strExc(0) = "select '' V,decode(S.st06,'1','北','2','中','3','南','4','高','其他')||'所' 所別" & _
      ",a0902 業務區,S.st02 智權人員,to_char(tot,'99999.00') 簽約點數,to_char(收文點數,'99999.00') 收文點數,to_char(V1C2,'99999.00') 入帳點數,to_char(Z1,'99999.00') 目標點數,S.ST01,S.ST15" & _
      " from Staff S,(" & stVTB1 & ") x, (" & stVTB2 & ") y,(" & stVTB3 & ") z,(" & stVTB4 & ") w,acc090" & _
      " where s.st01=X.st01(+) and s.st15=a0901(+) and s.st01=y.ax209(+) and s.st01=cp13(+) AND s.st01=z.pe01(+) and (tot<>0 or 收文點數<>0 or V1C2<>0 or z1<>0) order by S.st06,S.st15,S.st01"
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      grdDataList.Visible = False
      Set grdDataList.Recordset = RsTemp.Clone
      SetDataListWidth True
      Calculate2
      grdDataList.Visible = True
      
      '若查詢結果只有一筆資料
      If Me.grdDataList.Rows = 2 Then
            grdDataList.row = 1
            grdDataList.col = 1
            If grdDataList.Text <> "" Then
                '直接選定
                bolSelData = True
                grdDataList.Visible = False
                grdDataList.row = 1
                grdDataList.col = 0
                grdDataList.Text = "V"
                For i = 0 To grdDataList.Cols - 1
                     grdDataList.col = i
                     grdDataList.CellBackColor = &HFFC0C0
                Next i
                grdDataList.Visible = True
            End If
        End If
   Else
      If Me.Visible = True Then MsgBox "無符合資料！", vbInformation
   End If
   doQuery = True
   
End Function

'Modify by Amy 2024/04/03 +IsFormLoad
Private Sub SetDataListWidth(Optional ByVal p_bolHeaderOnly As Boolean = False, Optional ByVal IsFormLoad As Boolean = False)
   Dim ii  As Integer, iRow As Integer 'Add by Amy 2024/04/03
   
   With grdDataList
   If p_bolHeaderOnly = False Then
      .Clear
      .Rows = 2
      .ColAlignmentFixed = flexAlignCenterCenter
   End If
   'Modify By Amy 2013/04/24 增加V及智權人員編號欄位
   '.TextMatrix(0, 0) = "所別"
   '.ColWidth(0) = 800
   '.TextMatrix(0, 1) = "業務區"
   '.ColWidth(1) = 800
   '.TextMatrix(0, 2) = "智權人員"
   '.ColWidth(2) = 1000
   '.TextMatrix(0, 3) = "簽約點數"
   '.ColWidth(3) = 1000
   '.ColAlignment(3) = flexAlignRightCenter
   '.TextMatrix(0, 4) = "收文點數"
   '.ColWidth(4) = 1000
   '.ColAlignment(4) = flexAlignRightCenter
   '.TextMatrix(0, 5) = "入帳點數"
   '.ColWidth(5) = 1000
   '.ColAlignment(5) = flexAlignRightCenter
   '.TextMatrix(0, 6) = "目標點數"
   '.ColWidth(6) = 1000
   '.ColAlignment(6) = flexAlignRightCenter
   'Modify by Amy 2024/04/03
   iRow = 0
   .TextMatrix(0, iRow) = "V"
   .ColWidth(iRow) = 250
   iRow = iRow + 1
   .TextMatrix(0, iRow) = "所別"
   .ColWidth(iRow) = 500
   iRow = iRow + 1
   .TextMatrix(0, iRow) = "業務區"
   .ColWidth(iRow) = 900
   iRow = iRow + 1
   .TextMatrix(0, iRow) = "智權人員"
   .ColWidth(iRow) = 900
   iRow = iRow + 1
   .TextMatrix(0, iRow) = "簽約點數"
   .ColWidth(iRow) = 0 'Modify By Sindy 2020/8/31 1000 => 0
   .ColAlignment(iRow) = flexAlignRightCenter
   iRow = iRow + 1
   .TextMatrix(0, iRow) = "收文點數"
   .ColWidth(5) = 1000
   .ColAlignment(iRow) = flexAlignRightCenter
   iRow = iRow + 1
   .TextMatrix(0, iRow) = "入帳點數"
   .ColWidth(iRow) = 1000
   .ColAlignment(iRow) = flexAlignRightCenter
   iRow = iRow + 1
   .TextMatrix(0, iRow) = "目標點數"
   .ColWidth(iRow) = 1000
   .ColAlignment(iRow) = flexAlignRightCenter
   iRow = iRow + 1
   .TextMatrix(0, iRow) = "StaffNo" '智權人編號
   .ColWidth(iRow) = 0
   iRow = iRow + 1
   .TextMatrix(0, iRow) = "SaleArea" '業務區代碼
   .ColWidth(iRow) = 0
   If IsFormLoad = True Then
      ReDim strField(iRow)
      For ii = 0 To iRow
         strField(ii) = .TextMatrix(0, ii)
      Next ii
   End If
   End With
End Sub

Private Function ConstrainCheck() As Boolean
Dim bolCancel As Boolean
Dim intErrCol As Integer
   
   'Add By Sindy 2020/6/12
   Call txtSales_Validate(bolCancel)
   If bolCancel = True Then
      txtSales.SetFocus
      txtSales_GotFocus
      ConstrainCheck = False
      Exit Function
   End If
   '2020/6/12 END
   
   'Modify By Sindy 2020/7/29 檢查部門欄位
   If PUB_ChkFormSalesDept(strUserNum, txtSales, txtSalesArea, txtSalesArea1, intErrCol) = False Then
      If intErrCol = 0 Then
         txtSales.SetFocus
         txtSales_GotFocus
      ElseIf intErrCol = 1 Then
         txtSalesArea.SetFocus
         txtSalesArea_GotFocus
      Else
         txtSalesArea1.SetFocus
         txtSalesArea1_GotFocus
      End If
      ConstrainCheck = False
      Exit Function
   End If
      
''cancel by sonia 2019/12/27 杜主秘說加開簡協理69005可看全所
''   '簡金泉只可看北所全部
''   If strUserNum = "69005" Then
''      If Left(txtSalesArea, 2) <> "S1" Then
''         MsgBox "請輸入北所業務區！", vbExclamation
''         txtSalesArea.SetFocus
''         txtSalesArea_GotFocus
''         ConstrainCheck = False
''         Exit Function
''      End If
''      If Left(txtSalesArea1, 2) <> "S1" Then
''         MsgBox "請輸入北所業務區！", vbExclamation
''         txtSalesArea1.SetFocus
''         txtSalesArea1_GotFocus
''         ConstrainCheck = False
''         Exit Function
''      End If
''   End If
''end 2019/12/27
'
'   'add by sonia 2016/12/21 柄佑82026可看中所全部或自已
'   If strUserNum = "82026" Then
'      If txtSalesArea <> "P11" Or txtSalesArea1 <> "P11" Then
'         If txtSalesArea < "S2" Or txtSalesArea > "S29" Then
'            MsgBox "業務區起始條件錯誤！只可查中所業務區", vbExclamation
'            txtSalesArea.SetFocus
'            txtSalesArea_GotFocus
'            ConstrainCheck = False
'            Exit Function
'         End If
'         If txtSalesArea1 < "S2" Or txtSalesArea1 > "S29" Then
'            MsgBox "業務區迄止條件錯誤！只可查中所業務區", vbExclamation
'            txtSalesArea1.SetFocus
'            txtSalesArea1_GotFocus
'            ConstrainCheck = False
'            Exit Function
'         End If
'      Else
'         If Trim(txtSales) <> strUserNum Then
'            MsgBox "查專利處工程師時只可查自己的資料", vbExclamation
'            txtSales.SetFocus
'            txtSales_GotFocus
'            ConstrainCheck = False
'            Exit Function
'         End If
'      End If
'   End If
'   'end 2016/12/21
'   'Added by Lydia 2020/06/15 簡協理可看所有智權人員
'   If strUserNum = "69005" Then
'      If Left(txtSalesArea, 1) <> "S" Then
'         MsgBox "業務區起始條件錯誤！只可查智權部", vbExclamation
'         txtSalesArea.SetFocus
'         txtSalesArea_GotFocus
'         ConstrainCheck = False 'Added by Lydia 2020/07/14
'         Exit Function
'      End If
'      If Left(txtSalesArea1, 1) <> "S" Then
'         MsgBox "業務區迄止條件錯誤！只可查智權部", vbExclamation
'         txtSalesArea1.SetFocus
'         txtSalesArea1_GotFocus
'         ConstrainCheck = False 'Added by Lydia 2020/07/14
'         Exit Function
'      End If
'   End If
'   'end 2020/06/15
'   If txtSalesArea > txtSalesArea1 Then
'      MsgBox "業務區範圍條件錯誤！", vbExclamation
'      txtSalesArea.SetFocus
'      txtSalesArea_GotFocus
'      ConstrainCheck = False
'      Exit Function
'   End If
   
   If txtCloseDate(0) = "" Then
      MsgBox "請輸入點數結算起日！", vbExclamation
      txtCloseDate(0).SetFocus
      txtCloseDate_GotFocus (0)
      ConstrainCheck = False
      Exit Function
   Else
      bolCancel = False
      Call txtCloseDate_Validate(0, bolCancel)
      If bolCancel = True Then
         ConstrainCheck = False
         Exit Function
      End If
   End If
   If txtCloseDate(1) = "" Then
      MsgBox "請輸入點數結算迄日！", vbExclamation
      txtCloseDate(1).SetFocus
      txtCloseDate_GotFocus (1)
      ConstrainCheck = False
      Exit Function
   Else
      bolCancel = False
      Call txtCloseDate_Validate(1, bolCancel)
      If bolCancel = True Then
         ConstrainCheck = False
         Exit Function
      End If
   End If
   '2015/4/29 add by sonia
   If Val(txtCloseDate(1)) < Val(txtCloseDate(0)) Then
      MsgBox "點數結算起迄日範圍錯誤！", vbExclamation
      txtCloseDate(0).SetFocus
      txtCloseDate_GotFocus (0)
      ConstrainCheck = False
      Exit Function
   End If
   'end 2015/4/29
   
   'Added by Lydia 2024/01/31 另外限制不可查詢
   If strUserNum = "75007" Then
      If DBDATE(txtCloseDate(0)) >= "20240201" Then
         MsgBox "不可查詢113年2月後的資料！", vbExclamation
         txtCloseDate(0).SetFocus
         txtCloseDate_GotFocus (0)
         ConstrainCheck = False
         Exit Function
      End If
      If DBDATE(txtCloseDate(1)) >= "20240201" Then
         MsgBox "不可查詢113年2月後的資料！", vbExclamation
         txtCloseDate(1).SetFocus
         txtCloseDate_GotFocus (1)
         ConstrainCheck = False
         Exit Function
      End If
   End If
   'end 2024/01/31
   
   ConstrainCheck = True
End Function

Private Sub runWord()
   Dim iRow As Integer, iCol As Integer, lngCellWidth As Long
   Dim iResumeCnt As Integer
   Dim stKey1 As String, stKey2 As String
   Dim oTable As Word.Table
   
On Error GoTo ErrHnd

   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
   g_WordAp.Documents.add
   With g_WordAp
      .Visible = True
      '.Visible = False
      
      
      .Selection.PageSetup.PaperSize = wdPaperA4
      .Selection.PageSetup.Orientation = wdOrientPortrait
      '邊界
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(1.6)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1.4)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
      
      .Selection.Font.Name = "標楷體"
      .Selection.Font.Size = 14
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.ParagraphFormat.DisableLineHeightGrid = True
      
      With .Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        '.DefaultBorderColor = wdColorBlack 'Word97 沒有這個屬性及常數(Word2007 有)
      End With
      
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
      .Selection.TypeText Text:="點數結算日：" & txtCloseDate(0) & " - " & txtCloseDate(1)
      
      .Selection.TypeParagraph
      'Modified by Lydia 2023/06/08 拿掉簽約點數: grdDataList.Cols - 3 => 4
      Set oTable = .Selection.Tables.add(Range:=.Selection.Range, NumRows:=grdDataList.Rows, NumColumns:=grdDataList.Cols - 4)
      
      'Added by Morgan 2014/1/28 Word 2007 預設沒有框線需另外指定
      With oTable
        .Borders(wdBorderLeft).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
        .Borders(wdBorderRight).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
        .Borders(wdBorderTop).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
        .Borders(wdBorderBottom).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
        .Borders(wdBorderVertical).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
        .Borders(wdBorderHorizontal).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
      End With
      'end 2014/1/28
      
      '設定表格高度
      .Selection.Cells.SetHeight RowHeight:=26, HeightRule:=wdRowHeightExactly
      For iRow = 0 To grdDataList.Rows - 1
         .Selection.SelectRow
         If iRow = 0 Then
            .Selection.Font.Bold = True
         Else
            .Selection.Font.Bold = False
         End If
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
         .Selection.MoveLeft
         'Modify by Amy 2013/05/23 不顯示5/24增加的欄位(0,8,9)
         'For iCol = 0 To GrdDataList.Cols - 1
         For iCol = 1 To grdDataList.Cols - 3
            'Added by Lydia 2023/06/08 跳過Grid欄位不產生欄位: 簽約點數
            If iCol = 4 Then
            Else
            'end 2023/06/08
               'If iCol = 0 Then 'Modify 2013/05/23
               If iCol = 1 Then
                  If stKey1 <> grdDataList.TextMatrix(iRow, iCol) Then
                     .Selection.TypeText Text:=grdDataList.TextMatrix(iRow, iCol)
                     stKey1 = grdDataList.TextMatrix(iRow, iCol)
                  End If
               'ElseIf iCol = 1 Then 'Modify 2013/05/23
               ElseIf iCol = 2 Then
                  If stKey2 <> grdDataList.TextMatrix(iRow, iCol) Then
                     .Selection.TypeText Text:=grdDataList.TextMatrix(iRow, iCol)
                     stKey2 = grdDataList.TextMatrix(iRow, iCol)
                  End If
               Else
                  .Selection.TypeText Text:=grdDataList.TextMatrix(iRow, iCol)
               End If
               .Selection.MoveRight Unit:=wdCharacter, Count:=1
            End If 'Added by Lydia 2023/06/08
         Next
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
      Next
      
      .Selection.WholeStory
      .Selection.Font.Name = "Times New Roman"
      .Selection.HomeKey
      .Visible = True
      .Activate
   End With
   
   MsgBox "Word檔已產生完畢!", vbInformation 'Added by Lydia 2023/06/08
   Exit Sub
   
ErrHnd:
   If Err.Number <> 0 Then
      Select Case Err.Number
         Case 91:
            g_WordAp.Documents.add
            Resume Next
         Case 462:
            Set g_WordAp = New Word.Application
            g_WordAp.Documents.add
            Resume Next
         Case Else:
            MsgBox "錯誤 : " & Err.Description, vbCritical
      End Select
   End If
End Sub

'Add By Amy 2013/04/24
Private Sub grdDataList_SelChange()
   bolSelData = True
   grdDataList.Visible = False
   grdDataList.row = grdDataList.MouseRow
   grdDataList.col = 0
   If grdDataList.row <> 0 Then
   If grdDataList.Text = "V" Then
        grdDataList.Text = ""
        bolSelData = False
        If grdDataList.TextMatrix(grdDataList.row, 3) = "合計：" Then
            lngColar = &H7FFFD4
            SetColor grdDataList.row, lngColar
        Else
           For i = 0 To grdDataList.Cols - 1
             grdDataList.col = i
             grdDataList.CellBackColor = &H80000018
           Next i
        End If
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

Private Sub txtCloseDate_GotFocus(Index As Integer)
   TextInverse txtCloseDate(Index)
   CloseIme
End Sub

Private Sub txtCloseDate_Validate(Index As Integer, Cancel As Boolean)
   If txtCloseDate(Index) <> "" Then
      If ChkDate(txtCloseDate(Index)) = False Then
         Cancel = True
         txtCloseDate(Index).SetFocus
         txtCloseDate_GotFocus Index
      End If
   End If
   'Added by Lydia 2024/01/31
   If strUserNum = "75007" Then
      If DBDATE(txtCloseDate(Index)) >= "20240201" Then
         MsgBox "不可查詢113年2月後的資料！", vbExclamation
         Cancel = True
         txtCloseDate(Index).SetFocus
         txtCloseDate_GotFocus Index
      End If
   End If
   'end 2024/01/31
End Sub

Private Sub txtSales_Change()
   If Len(txtSales) > 4 Then
      lblSalesName = StaffQuery(txtSales)
   Else
      lblSalesName = ""
   End If
End Sub

Private Sub txtSales_GotFocus()
   TextInverse txtSales
   CloseIme
End Sub

Private Sub txtSales_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSales_Validate(Cancel As Boolean)
   'Modify By Sindy 2024/9/18 智權部:使用智權人員ID查詢權限; 改為共用函數
   If PUB_txtSales_Limit(txtSales, m_strListPer, , txtSalesArea, txtSalesArea1, _
                         bolSpecMan, strSpecCode, lblSalesName) = False Then
      If txtSales.Visible = True Then 'Added by Lydia 2021/05/20 排除隱藏
        txtSales.SetFocus
        txtSales_GotFocus
      End If 'Added by Lydia 2021/05/20
      Cancel = True
      Exit Sub
   End If
   '2024/9/18 END
End Sub

Private Sub txtSalesArea_GotFocus()
   TextInverse txtSalesArea
   CloseIme
End Sub

Private Sub txtSalesArea_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesArea1_GotFocus()
   TextInverse txtSalesArea1
   CloseIme
End Sub

Private Sub txtSalesArea1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Calculate2()
   Dim ii As Integer, jj As Integer
   Dim iCount As Integer, dblSub(4) As Double
   Dim strAddItem As String ', lngColar As Long Modify By Amy 2013/04/24
   Dim dblTotSub(4) As Double 'Add By Sindy 2023/9/1
   
   With grdDataList
      .Visible = False
      iCount = 0
      Erase dblSub
      Erase dblTotSub 'Add By Sindy 2023/9/1
      ii = 1
      Do While ii < .Rows
         If ii > 1 Then
            'Modify By Amy 2013/04/24 增加勾選欄
            '區小計
            'If .TextMatrix(ii, 1) <> .TextMatrix(ii - 1, 1) Then
            'Modify by Amy 2024/04/03 改不用固定,原判斷iCount > 1下1130402 S00-S41 北一區只有張肆明一人不會加總
            'If .TextMatrix(ii, 2) <> .TextMatrix(ii - 1, 2) Then
            If .TextMatrix(ii, GetColVal(strField, "業務區")) <> .TextMatrix(ii - 1, GetColVal(strField, "業務區")) Then
               If iCount >= 1 Then
                  'strAddItem = .TextMatrix(ii - 1, 0)
                  'strAddItem = strAddItem & vbTab & .TextMatrix(ii - 1, 1)
                  strAddItem = vbTab & .TextMatrix(ii - 1, GetColVal(strField, "所別"))
                  strAddItem = strAddItem & vbTab & .TextMatrix(ii - 1, GetColVal(strField, "業務區"))
                  strAddItem = strAddItem & vbTab & "合計："
                  For jj = 1 To 4
                     strAddItem = strAddItem & vbTab & Right(String(5, " ") & Format(dblSub(jj), "0.00"), 9)
                  Next
                  .AddItem strAddItem, ii
                  lngColar = &H7FFFD4
                  SetColor ii, lngColar
                  ii = ii + 1
               End If
               iCount = 0
               Erase dblSub
            End If
         End If
         
         iCount = iCount + 1
         For jj = 1 To 4 '欄
            'Modify By Amy 2013/04/24 增加勾選欄
            'dblSub(jj) = dblSub(jj) + Val(.TextMatrix(ii, 2 + jj))
            dblSub(jj) = dblSub(jj) + Val(.TextMatrix(ii, GetColVal(strField, "智權人員") + jj))
            dblTotSub(jj) = dblTotSub(jj) + Val(.TextMatrix(ii, GetColVal(strField, "智權人員") + jj)) 'Add By Sindy 2023/9/1
         Next
         ii = ii + 1
      Loop
      If iCount > 1 Then
         'Modify By Amy 2013/04/24 增加勾選欄
         'strAddItem = .TextMatrix(ii - 1, 0)
         'strAddItem = strAddItem & vbTab & .TextMatrix(ii - 1, 1)
         strAddItem = vbTab & .TextMatrix(ii - 1, GetColVal(strField, "所別"))
         strAddItem = strAddItem & vbTab & .TextMatrix(ii - 1, GetColVal(strField, "業務區"))
         strAddItem = strAddItem & vbTab & "合計："
         For jj = 1 To 4
            strAddItem = strAddItem & vbTab & Right(String(5, " ") & Format(dblSub(jj), "0.00"), 9)
         Next
         .AddItem strAddItem, ii
         lngColar = &H7FFFD4
         SetColor ii, lngColar
         'Add By Sindy 2023/9/1 增加顯示總計
         ii = ii + 1
         strAddItem = vbTab & ""
         strAddItem = strAddItem & vbTab & ""
         strAddItem = strAddItem & vbTab & "總計："
         For jj = 1 To 4
            strAddItem = strAddItem & vbTab & Right(String(5, " ") & Format(dblTotSub(jj), "0.00"), 9)
         Next
         .AddItem strAddItem, ii
         lngColar = &H7FFFD4
         SetColor ii, lngColar
         '2023/9/1 END
      End If
      .Visible = True
   End With
End Sub

Private Sub SetColor(iRow As Integer, lngColar As Long)
Dim ii As Integer, jj As Integer
   
   With grdDataList
   .row = iRow
   For jj = 0 To .Cols - 1
      .col = jj: .CellBackColor = lngColar
   Next
   End With
End Sub

'Add By Amy 2013/04/24
Public Sub PubShowNextData()
Dim Str00 As String, Str01 As String

   Me.Enabled = False
   For i = 1 To grdDataList.Rows - 1
      grdDataList.col = 0
      grdDataList.row = i
      Str00 = Trim(grdDataList.Text) '勾選欄
      grdDataList.col = 3
      grdDataList.row = i
      Str01 = Trim(grdDataList.Text) '取智權人姓名
      If Str00 = "V" And Str01 <> "合計：" Then
         Dim Str02 As String, Str03, Str04 As String
         grdDataList.col = 0
         grdDataList.Text = ""
         For j = 0 To grdDataList.Cols - 1
             grdDataList.col = j
             grdDataList.CellBackColor = &H80000018
         Next j
         
         '取智權人編號
         grdDataList.col = 8
         Str02 = grdDataList.Text
         
         '取業務區代號
         grdDataList.col = 9
         Str03 = grdDataList.Text
         
         '取點數合計
         grdDataList.col = 5
         Str04 = Trim(grdDataList.Text)
         
         If Not IsNull(grdDataList.Text) Then
             If fnSaveParentForm(Me) = False Then
                 Me.Enabled = True
                 Exit Sub
             End If
             Screen.MousePointer = vbHourglass
             frm210137_1.Show
             frm210137_1.Tag = Str01 & "," & Str02 & "," & Str03 & "," & txtCloseDate(0) & "," & txtCloseDate(1) & "," & Str04
             frm210137_1.StrMenu
             Screen.MousePointer = vbDefault
             Me.Enabled = True
             Exit Sub
           End If
      ElseIf Str01 = "合計：" Then
         grdDataList.TextMatrix(i, 0) = ""
         lngColar = &H7FFFD4
         SetColor i, lngColar
      End If
   Next i
   Me.Enabled = True

End Sub
