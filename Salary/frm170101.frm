VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm170101 
   BorderStyle     =   1  '單線固定
   Caption         =   "每月薪資計算"
   ClientHeight    =   2952
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2952
   ScaleWidth      =   5160
   Begin VB.CommandButton cmdok 
      Caption         =   "取消計算(&C)"
      Height          =   405
      Index           =   2
      Left            =   135
      TabIndex        =   8
      Top             =   180
      Width           =   1245
   End
   Begin VB.ListBox List1 
      Height          =   948
      Left            =   90
      TabIndex        =   6
      Top             =   1710
      Width           =   4920
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   405
      Index           =   1
      Left            =   3945
      TabIndex        =   5
      Top             =   180
      Width           =   1065
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "開始(&O)"
      Height          =   405
      Index           =   0
      Left            =   2775
      TabIndex        =   4
      Top             =   180
      Width           =   1065
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   30
      TabIndex        =   3
      Top             =   1110
      Width           =   5010
      _ExtentX        =   8827
      _ExtentY        =   445
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '靠右對齊
      Height          =   255
      Left            =   1920
      MaxLength       =   2
      TabIndex        =   2
      Text            =   "5"
      Top             =   750
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '靠右對齊
      Height          =   255
      Left            =   1170
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "96"
      Top             =   750
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "( 0/0 )"
      Height          =   255
      Left            =   90
      TabIndex        =   7
      Top             =   1410
      Width           =   4920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "薪資年月：           年              月"
      Height          =   180
      Left            =   210
      TabIndex        =   0
      Top             =   780
      Width           =   2385
   End
End
Attribute VB_Name = "frm170101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/1/26 Form2.0已檢查 (無需修改的物件)
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Create by Morgan 2008/12/25
Option Explicit

Dim m_Actived As Boolean
Dim YM As String
Dim WDs As String '當月工作天數

'Removed by Morgan 2017/7/10 配合薪資明細查詢改用 GetDaiyHour
''特殊每日工作時數 Iain
''Modify by Morgan 2011/3/8 Iain 改工作時數 4->5
''Modify by Morgan 2011/10/8 Iain 10月起改工作時數 5->6
'Const cDailyHr99029 As Integer = 6
''Removed by Morgan 2016/4/14 尤春彬8404 105/3/1 改回全職
''Const cDailyHr84043 As Integer = 4 'Added by Morgan 2012/7/9 101/7月起尤春彬84043 工作時數改 4 小時
'Const cDailyHr73029 As Integer = 4 'Added by Morgan 2013/8/8 102/8月起廖宗岳73029 工作時數改 4 小時
'Dim iDailyHours As Integer '每日小時數,正常為8 例外為 4 Add by Morgan 2010/7/14
'end 2017/7/10
Private Sub cmdok_Click(Index As Integer)
   Dim bolMsg As Boolean
   
   Select Case Index
      Case 0
         Screen.MousePointer = vbHourglass
         If TxtValidate = True Then
            Me.Enabled = False
            If CancelSalaryMonth = True Then 'Added by Morgan 2022/3/3
               
               If Process = True Then
                  '檢查(基本薪資+職務津貼+超時加班費-缺勤扣薪)不可小於0
                  'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
                  'strExc(0) = "select st02||'( '||sm01||' )' from salarymonth,staff" & _
                     " where sm02=" & (Val(Text1) * 100 + Val(Text2) + 191100) & _
                     " and nvl(sm04,0)+nvl(sm05,0)+nvl(sm28,0)-nvl(sm21,0)<0" & _
                     " and st01(+)=replace(sm01,'A','0')"
                  'Modify By Sindy 2020/6/24 + nvl(sm45,0) (基本薪資+職務津貼+證照津貼+超時加班費-缺勤扣薪)不可小於0
                  'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
                  strExc(0) = "select st02||'( '||sm01||' )' from salarymonth,staff" & _
                     " where sm02=" & (Val(Text1) * 100 + Val(Text2) + 191100) & _
                     " and nvl(sm04,0)+nvl(sm05,0)+nvl(sm45,0)+nvl(sm28,0)-nvl(sm21,0)<0" & _
                     " and st01(+)=substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     List1.AddItem time & " --> " & RsTemp.GetString(adClipString, , , " ") & " 薪資所得小於 0", 0
                     RsTemp.MoveFirst
                     strExc(1) = "薪資計算結束但下列員工薪資所得小於 0 !!" & vbCrLf & vbCrLf & RsTemp.GetString(adClipString, , , vbCrLf)
                     MsgBox strExc(1), vbExclamation
                     bolMsg = True
                  End If
                  
                  'Added by Morgan 2015/1/28
                  '檢查是否有當月滿65歲且有勞保費的
                  'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
                  strExc(0) = "select st02||'( '||sm01||' ) 生日:'||sqldatet(st23) from salarymonth,staff" & _
                     " where sm02=" & (Val(Text1) * 100 + Val(Text2) + 191100) & _
                     " and sm14>0" & _
                     " and st01(+)=substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4) and st23>= (sm02-6500)*100+1 and st23< (sm02-6500)*100+32"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     List1.AddItem time & " --> " & RsTemp.GetString(adClipString, , , " ") & " 滿65歲", 0
                     RsTemp.MoveFirst
                     strExc(1) = RsTemp.GetString(adClipString, , , vbCrLf) & vbCrLf & vbCrLf & _
                        "上列員工該月份年滿65歲請執行下列步驟：" & vbCrLf & vbCrLf & _
                        "1. 視需要自行調整勞保費(依照天數比例扣除就保保費)" & vbCrLf & vbCrLf & _
                        "2. 新增次月1日之薪資異動且設定為逕行調整（薪資基本檔勞保費會即時更新）" & vbCrLf & vbCrLf & _
                        "註：年滿 65 歲 , 是指年滿 65 歲生日的前一天 ( 扣費以天數計算)"
                     MsgBox strExc(1), vbExclamation
                     bolMsg = True
                  End If
                  'end 2015/1/28
                  
                  If bolMsg = False Then
                     MsgBox "成功!", vbInformation
                  End If
                  PUB_SendMailCache 'Added by Morgan 2022/11/16
               Else
                  MsgBox "失敗!", vbCritical
               End If
               
            End If 'Added by Morgan 2022/3/3
            Me.Enabled = True
         End If
         Screen.MousePointer = vbDefault
      Case 1
         Unload Me
      
      'Added by Morgan 2016/10/4
      Case 2
         Screen.MousePointer = vbHourglass
         If TxtValidate(True) = True Then
            Me.Enabled = False
            If MsgBox("是否確定要取消計算?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
               If CancelSalaryMonth = True Then
                  MsgBox "計算已取消!", vbInformation
               End If
            End If
            Me.Enabled = True
         End If
         Screen.MousePointer = vbDefault
      'end 2016/10/4
   End Select
End Sub

Private Function CancelSalaryMonth() As Boolean

   YM = 100 * Val(Text1) + Val(Text2) + 191100
   
   cnnConnection.BeginTrans
On Error GoTo ErrHnd
      
   strSql = "delete salarymonth where sm02=" & YM
   cnnConnection.Execute strSql, intI
   'Added by Morgan 2022/11/16 刪除薪資計算自動新增的每月獎金資料
   strSql = "delete MonthBonus where mb01>=" & YM & "01 and mb05='QPGMR' and MB13>0"
   cnnConnection.Execute strSql, intI
   'end 2022/11/16
   strSql = "delete from nhi2nd where trunc(nhi02/100)=" & YM & " and nhi10>=235910 and nhi03='50'"
   cnnConnection.Execute strSql, intI
   cnnConnection.CommitTrans
   CancelSalaryMonth = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbExclamation
   
End Function

Private Sub Form_Activate()
   If m_Actived = False Then
      FormReset
      Text2.SetFocus
      Text2_GotFocus
      m_Actived = True
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170101 = Nothing
End Sub

Private Sub FormReset()
   Dim stDate As String
   If Val(Right(strSrvDate(2), 2)) < 11 Then
      stDate = CompDate("1", -1, strSrvDate(1)) - 19110000
   Else
      stDate = strSrvDate(2)
   End If
   Text1.Text = stDate \ 10000
   Text2.Text = Val(Right(stDate \ 100, 2))
   '薪資不可重跑非當次的
   Text1.Locked = True
   Text2.Locked = True
   Label2 = "( 0/0 )"
   List1.Clear
End Sub

Private Function Process() As Boolean
   Dim stSQL As String
   Dim adoRst As ADODB.Recordset, intR As Integer
   Dim stLstNo As String
   'Dim stLstNo2 As String 'Removed by Morgan 2015/12/23 取消,已經沒有第2家公司
   Dim stLstDate As String
   Dim dblHourPay As Double '時薪
   Dim dblOTHour As Double '加班時數
   Dim dblOTHourH As Double '假日加班時數(8小時以內)
   Dim dblOTHourTot As Double '加班時數累計
   Dim iDailyHours As Integer '每日小時數,正常為8 例外為 4 Add by Morgan 2010/7/14
   
   'Modified by Morgan 2015/12/23
   'Dim dblOTPay1 As Double '第一家加班費
   'Dim dblOTPay2 As Double '第二家加班費
   'Dim lngOTPayTot1 As Long '第一家加班費累計
   'Dim lngOTPayTot2 As Long '第二家加班費累計
   'Dim dblRestHour2 As Double '第二家可用加班時數
   Dim dblOTPay1 As Double '應稅加班費(46小時以後)
   Dim dblOTPay2 As Double '加班費
   Dim lngOTPayTot1 As Long '應稅加班費累計
   Dim lngOTPayTot2 As Long '加班費累計
   Dim dblRestHour2 As Double '免稅加班時數
   'end 2015/12/23
   
   Dim lngSM21 As Double '缺勤扣款
   Dim lngSM22 As Long '未打卡扣款
   Dim dblHour As Double '曠職事病假扣薪時數
   Dim dblSickHour As Double '病假時數累計
   Dim dblGirlSickHour As Double '生理假時數累計
   Dim dblSickBaseHour As Double '病假扣半薪時數(超過扣全薪)
   Dim dblGirlSickBaseHour As Double '生理假半薪時數(超過且合併病假超過33天扣全薪)
   Dim dblGrilSickRemainHour As Double '未計算生理假時數
   Dim dblGrilSickMergeHour As Double
   Dim stNHI() As String
   Dim intCount As Integer 'Added by Morgan 2013/4/24
   Dim NextYM As String    'add by sonia 2016/9/2
   Dim bolRestDay As Boolean '休息日 Added by Morgan 2016/12/26
   Dim dblSickHr As Double, dblGirlSickHr As Double 'Added by Morgan 2020/10/7
   Dim stDepCol As String 'Added by Morgan 2023/12/25 員工檔部門欄位
   Dim dblTFeeNet As Double 'Added by Morgan 2024/5/13 約定薪資+上月翻譯預支金額(列本月扣款)-翻譯費, >0 時列次月扣款(本月翻譯預支), <0 時列本月翻譯薪資
   Dim stLstFNo As String 'Added by Morgan 2024/5/13
   
   ReDim stNHI(TF_NHI) As String
   
   cnnConnection.BeginTrans
   
On Error GoTo ErrHnd
   
   '計算年月(西元)
   YM = 100 * Val(Text1) + Val(Text2) + 191100
   'add by sonia 2016/9/2
   If Val(Text2) = 12 Then
      NextYM = 100 * Val(Text1) + 191201
   Else
      NextYM = 100 * Val(Text1) + Val(Text2) + 191101
   End If
   'end 2016/9/2
   
   'Added by Morgan 2023/12/25 新部門抓st93
   If YM >= Left(新部門啟用日, 6) Then
      stDepCol = "st93"
   Else
      stDepCol = "st03"
   End If
   'end 2023/12/25
   
   'Modify by Morgan 2009/3/5 str() 函數有前置符號，長度為9 碼會導致日期計算有誤
   'WDs = Right(CompDate(2, -1, CompDate(1, 1, str(100 * YM + 1))), 2)
   WDs = Right(CompDate(2, -1, CompDate(1, 1, Format(100 * YM + 1))), 2)
   List1.Clear
   
   List1.AddItem time & " --> 刪除互助會扣款資料開始...", 0
   DoEvents
   stSQL = "delete WFAmount where substr(wfa01,1,6)=" & YM & " and wfa05='2'"
   cnnConnection.Execute stSQL, intR
   
   List1.AddItem time & " --> 刪除互助會扣款資料結束,共 " & intR & " 筆", 0
   
   List1.AddItem time & " --> 新增互助會扣款資料開始...", 0
   DoEvents
   '新增互助會扣款資料
   '1.活會抓會金CO02,2.死會抓會金CO02+得標金CM05,3.當月得標者不扣
   'Modified by Morgan 2014/5/5 得標日可能不是標會日
   'stSQL = "INSERT INTO WFAmount(WFA01,WFA02,WFA03,WFA04,WFA05)" & _
      " SELECT SUBSTR(MAX(" & YM & "||CO06),1,8) X1,CO01 X2,CM03 X3,SUM(DECODE(CM04,NULL,CO02,CM05+CO02)) X4,'2'" & _
      " From Cooperation, CooperationMEMBER" & _
      " WHERE CO05>" & YM & "00 AND CM01(+)=CO01 AND CM03<>'6000' AND ( CM04 IS NULL OR substr(CM04,1,6)<>" & YM & " )" & _
      " GROUP BY CO01,CM03"
   '2015/1/5 modify by sonia 首會會抓不到得標日
   'stSQL = "INSERT INTO WFAmount(WFA01,WFA02,WFA03,WFA04,WFA05)" & _
      " SELECT MAX(b.cm04) X1,CO01 X2,a.CM03 X3,SUM(DECODE(a.CM04,NULL,CO02,a.CM05+CO02)) X4,'2'" & _
      " From Cooperation, CooperationMEMBER a, CooperationMEMBER b" & _
      " WHERE CO05>" & YM & "00 AND a.CM01(+)=CO01 AND a.CM03<>'6000' AND ( a.CM04 IS NULL OR substr(a.CM04,1,6)<>" & YM & " )" & _
      " and b.cm01(+)=a.cm01 and b.cm04(+)>=" & YM & "01 and b.cm04(+)<=" & YM & "31 GROUP BY CO01,a.CM03"
   'modify by sonia 2016/8/3 要加入互助會的期間起日CO04判斷
   stSQL = "INSERT INTO WFAmount(WFA01,WFA02,WFA03,WFA04,WFA05)" & _
      " SELECT NVL(MAX(b.cm04),SUBSTR(MAX(" & YM & "||CO06),1,8)) X1,CO01 X2,a.CM03 X3,SUM(DECODE(a.CM04,NULL,CO02,a.CM05+CO02)) X4,'2'" & _
      " From Cooperation, CooperationMEMBER a, CooperationMEMBER b" & _
      " WHERE CO04<" & NextYM & "00 AND CO05>" & YM & "00 AND a.CM01(+)=CO01 AND a.CM03<>'6000' AND ( a.CM04 IS NULL OR substr(a.CM04,1,6)<>" & YM & " )" & _
      " and b.cm01(+)=a.cm01 and b.cm04(+)>=" & YM & "01 and b.cm04(+)<=" & YM & "31 GROUP BY CO01,a.CM03"
   cnnConnection.Execute stSQL, intR
   
   List1.AddItem time & " --> 新增互助會扣款資料結束,共 " & intR & " 筆", 0
   
   List1.AddItem time & " --> 更新互助會當月得標會款金開始...", 0
   DoEvents
   '會首固定為台一且不建會員資料所以加總時把得標人當台一來加總
   stSQL = "Update CooperationMEMBER a set CM06=( select SUM(decode(sign(b.CM04-a.CM04),-1,b.CM05+CO02,CO02))" & _
      " from CooperationMEMBER b,Cooperation where co01(+)=b.cm01 and b.cm01=a.cm01 )" & _
      " where substr(CM04,1,6)=" & YM
   
   cnnConnection.Execute stSQL, intR
   List1.AddItem time & " --> 更新互助會當月得標會款金結束,共 " & intR & " 筆", 0
   
   List1.AddItem time & " --> 刪除月薪資資料開始...", 0
   DoEvents
   stSQL = "delete salarymonth where sm02=" & YM
   cnnConnection.Execute stSQL, intR
   
   List1.AddItem time & " --> 刪除月薪資資料結束,共 " & intR & " 筆", 0
   
   
   List1.AddItem time & " --> 新增第一家月薪資資料(部門別,工作天數,公司別)", 0
   DoEvents
   'Modify by Morgan 2009/3/5 +排除復職日大於當月
   '抓有建薪資基本資料的在職員工(到職日<=當月31號)及離職日>當月1號者(若1號離職當月份不用付薪水)
   '工作天數:
   '1.非當月到(復)職或離職=全部;
   '2.當月到(復)職非當月離職=全部-到職日+1;
   '3.當月離職=離職日-當月到(復)職日或1(離職日不算);
   '4.勞退投保天數=當月離職(含當月到職)時同(工作天),當月到職時(30-(到職日-1));比率固定除30天
   '  Ex.2/2日到職 *29/30,2/28離職 *27/30,3/31到職 *0/30,3/31離職 *30/30,3/2到職 3/31離職 *29/30
   'Modify by Morgan 2009/11/26 新增勞保健保投保薪資欄位SM40,SM41
   'Modified by Morgan 2013/1/21 +sm42<-sd47
   stSQL = "insert into salarymonth(sm01,sm02,sm03,sm27,sm39,sm37,sm40,sm41,sm42)" & _
      " select st01," & YM & "," & stDepCol & _
      ",decode(b2,null,decode(nvl(c2,d2),null," & WDs & "," & WDs & "-nvl(c2,d2)+1),b2-nvl(nvl(c2,d2),1)) sm27" & _
      ",decode(b2,null,decode(nvl(c2,d2),null,30,30-nvl(c2,d2)+1),b2-nvl(nvl(c2,d2),1)) sm39" & _
      ",sd19,sm40,sm41,sd47 from (SELECT " & stDepCol & ",st01,sd19,nvl(sd12,sd45) sm40,nvl(sd13,sd45) sm41,sd47 FROM salarydata,staff" & _
      " WHERE sd01<'F' and st01(+)=sd01 and st04='1' and st13<=" & YM & "31" & _
      " and not exists(select * from staff_Change where sc01=st01 and sc03='02' and sc02>" & YM & "31)" & _
      " Union SELECT " & stDepCol & ",st01,sd19,nvl(sd12,sd45) sm40,nvl(sd13,sd45) sm41,sd47 FROM staff,salarydata" & _
      " WHERE st13<=" & YM & "31 and ST51>" & YM & "01 and ST51<=" & strSrvDate(1) & " and sd01(+)=st01" & _
      " and not exists(select 1 from Staff_Change where sc01=st01 and sc02<=" & YM & "01 having substr(max(sc02||sc03),9)='04')) a" & _
      ",(select st01 b1,substr(st51,7) b2 from staff where substr(st51,1,6)=" & YM & ") b" & _
      ",(select st01 c1,substr(st13,7) c2 from staff where substr(st13,1,6)=" & YM & ") c" & _
      ",(select sc01 d1,substr(sc02,7) d2 from Staff_Change where substr(sc02,1,6)=" & YM & " and sc03='02') d" & _
      " where b1(+)=st01 and c1(+)=st01 and d1(+)=st01"
   
   cnnConnection.Execute stSQL, intR
   
   List1.AddItem time & " --> 新增第一家月薪資資料結束,共 " & intR & " 筆", 0
   
   'Modify by Morgan 2009/12/25 當月到職又離職的也要扣健保費
   'Add by Morgan 2009/6/22
   '新增健保費明細資料(當月離職不用)
   List1.AddItem time & " --> 刪除健保費明細資料開始...", 0
   DoEvents
   stSQL = "delete HIMONTH where HM03=" & YM
   cnnConnection.Execute stSQL, intR
   
   List1.AddItem time & " --> 刪除健保費明細資料結束,共 " & intR & " 筆", 0

   List1.AddItem time & " --> 新增健保費明細資料開始...", 0
   DoEvents
   strExc(1) = CompDate(2, -1, CompDate(1, 1, YM & "01"))
   '是否建保眷屬判斷當月底前的最後異動
   'Modified by Morgan 2016/11/12
   '當月最後一天轉出或當月份轉入又轉出都要扣--辜,靖蓉
   'Modified by Morgan 2017/12/6 106/12起健保改為以當月最末日所屬之投保單位計費(最末日退保者核定次月1日為退保日)
   'Modified by Morgan 2019/3/27 本所離職日原則即為退保日再改回當月離職不扣健保費 Ex:柯冠羽A0031 2019/2/28離職--辜(有跟劉經理確認)
   stSQL = "INSERT INTO HIMONTH (HM01,HM02,HM03,HM04,HM06)" & _
      " select SD01,0,SM02,round(decode(hr02,null,sd15" & _
      ",decode(hr03,null,sd15*(1-hr02/100),decode(sign(hr03-sd15*hr02/100),1,sd15*(1-hr02/100),hr03))),0),HR01" & _
      " From salarymonth, salarydata, staff, HiReduce" & _
      " where sm02=" & YM & " and sd01(+)=sm01 and sd15>0" & _
      " and st01(+)=sd01 and (st51 is null or st51>" & strExc(1) & ") and hr01(+)=st50" & _
      " Union All" & _
      " select SR01,SR02,SM02,round(sd15-decode(hr02,null,0" & _
      ",decode(hr03,null,sd15*hr02/100,decode(sign(hr03-sd15*hr02/100),1,sd15*hr02/100,hr03))),0),HR01" & _
      " from salarymonth,salarydata, staff,staff_Relation,hirelationlog a,HiReduce" & _
      " where sm02=" & YM & " and sd01(+)=sm01 and st01(+)=sd01 and (st51 is null or st51>" & strExc(1) & ")" & _
      " and SR01(+)=sd01 and hl01(+)=sr01 and hl02(+)=sr02" & _
      " and HL03=(select max(b.hl03) from hirelationlog b" & _
      " where b.hl01=a.hl01 and b.hl02=a.hl02 and b.hl03<=" & YM & "31 group by hl02)" & _
      " and (HL04<>'2' or HL03=" & strExc(1) & ")" & _
      " and HR01(+)=HL05"
   'end 2017/12/12
   cnnConnection.Execute stSQL, intR
   List1.AddItem time & " --> 新增健保費明細資料結束,共 " & intR & " 筆", 0
   
   'Added by Morgan 2017/12/26 沒有健保費明細的清除健保投保薪資(sm41)及投保金額(sm42)
   List1.AddItem time & " --> 更新健保投保金額開始...", 0
   stSQL = "update salarymonth set sm41=0,sm42=0 where sm02=" & YM & _
      " and sm42>0 and not exists(select * from himonth where hm01=sm01 and hm03=sm02)"
   cnnConnection.Execute stSQL, intR
   List1.AddItem time & " --> 更新健保投保金額結束,共 " & intR & " 筆", 0
   'end 2017/12/26
   
   List1.AddItem time & " --> 設定超過最高眷口數之健保費明細...", 0
   stSQL = "select hm01 from himonth where hm03=" & YM & " and hm02<>0 group by hm01 having count(*)>3"
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL, , True)
   If intR = 1 Then
      With adoRst
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      Do While Not .EOF
         List1.AddItem time & " --> 設定 " & .Fields(0) & " 健保費明細資料", 0
         ProgressBar1.Value = ProgressBar1.Value + 1
         Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         UpdHiMonth "" & .Fields(0), YM
         .MoveNext
      Loop
      End With
   End If
   
   List1.AddItem time & " --> 更新健保費資料開始...", 0
   DoEvents
   
   stSQL = "update salarymonth set sm15=(select nvl(sum(HM04),0) from HiMonth" & _
      " where HM01=sm01 and HM03=sm02 and HM05 is null) where SM02=" & YM
      
   cnnConnection.Execute stSQL, intR
   
   List1.AddItem time & " --> 更新健保費資料結束,共 " & intR & " 筆", 0

   'end 2009/6/22
   
   
   'Add by Morgan 2009/4/6 部門可能會有異動故要以人事異動資料更新
   List1.AddItem time & " --> 更新部門資料...", 0
   'Modified by Morgan 2015/8/4 人事異動要抓到薪資月份的最後一日
   stSQL = "update salarymonth set sm03=(select nvl(max(sc04),sm03) from staff_change a" & _
      " where sc01=sm01 and sc02=(select max(b.sc02) from staff_change b where b.sc01=sm01 and b.sc02<=" & YM & "31))" & _
      " where sm02=" & YM & " and exists(select * from staff_change where sc01=sm01 and sc02>=" & YM & "01)"
   cnnConnection.Execute stSQL, intR
   List1.AddItem time & " --> 更新部門資料結束,共 " & intR & " 筆", 0
   
   
   List1.AddItem time & " --> 更新月薪資資料...", 0
   DoEvents
   stSQL = "select sm01,sd02 from salarymonth,salarydata where sm02=" & YM & " and sd01(+)=sm01"
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL, , True)
   If intR = 1 Then
      With adoRst
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      Do While Not .EOF
         List1.AddItem time & " --> 更新 " & .Fields(0) & " 月薪資資料", 0
         ProgressBar1.Value = ProgressBar1.Value + 1
         Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         '抓薪資基本檔編制判斷是否為兼職人員
         UpdMainItems "" & .Fields(0), YM, "" & .Fields(1)
         .MoveNext
      Loop
      End With
   End If

   'Add by Morgan 2009/6/24
   List1.AddItem time & " --> 更新有補助勞保費資料開始...", 0
   DoEvents
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'stSQL = "update salarymonth set sm14=(select round(round(sm14-decode(lr03,null,sm14*lr02/100" & _
      ",decode(sign(lr03-sm14*lr02/100),1,sm14*lr02/100,lr03)))*sm39/30) from staff,LiReduce" & _
      " where st01=replace(sm01,'A','0') and LR01(+)=st56) where SM02=" & YM & _
      " and exists(select * from staff where st01=replace(sm01,'A','0') and st56 is not null)"
   'Modified by Morgan 2022/10/5 修正重複計算投保天數比例(*sm39/30)問題(批次計算勞保費時已乘過)
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   stSQL = "update salarymonth set sm14=(select round(round(sm14-decode(lr03,null,sm14*lr02/100" & _
      ",decode(sign(lr03-sm14*lr02/100),1,sm14*lr02/100,lr03)))) from staff,LiReduce" & _
      " where st01=substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4) and LR01(+)=st56) where SM02=" & YM & _
      " and exists(select * from staff where st01=substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4) and st56 is not null)"
   cnnConnection.Execute stSQL, intR
   
   List1.AddItem time & " --> 更新有補助勞保費資料結束,共 " & intR & " 筆", 0
   'end 2009/6/22
   
   List1.AddItem time & " --> 更新加班費資料開始...", 0
   
   'Modified by Morgan 2023/6/6 薪資計算後才新增的加班資料算到隔月薪資給付，改為該月前未算過的都抓
   'stSQL = "select so01,so02,sum(so05) ot1,sum(so06) ot2" & _
      " from Staff_Overtime where substr(so02,1,6)=" & YM & "" & _
      " group by so01,so02 order by 1 asc,2 asc"
   stSQL = "update Staff_Overtime set so16=" & YM & " where substr(so02,1,6)<=" & YM & " and ( so16=0 or so16=" & YM & ")"
   cnnConnection.Execute stSQL, intR
   
   stSQL = "select so01,so02,sum(so05) ot1,sum(so06) ot2" & _
      " from Staff_Overtime where substr(so02,1,6)<=" & YM & " and SO16=" & YM & _
      " group by so01,so02 order by 1 asc,2 asc"
   'end 2023/6/6
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL, , True)
   If intR = 1 Then
      With adoRst
      .MoveFirst
      stLstNo = ""
      '加班人數
      intR = 0
      Do While Not .EOF
         If .Fields("so01") <> stLstNo Then
            intR = intR + 1
            stLstNo = "" & .Fields("so01")
         End If
         .MoveNext
      Loop
      ProgressBar1.max = intR
      ProgressBar1.Value = 0
      Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      
      .MoveFirst
      stLstNo = "" & .Fields("so01")
      'stLstNo2 = Left(stLstNo, 2) & "A" & Mid(stLstNo, 4) 'Removed by Morgan 2015/12/23 取消,已經沒有第2家公司
      stLstDate = "" & .Fields("so02")
      'Modified by Morgan 2017/7/10
      'dblHourPay = GetHourPay(stLstNo, YM)
      iDailyHours = GetDaiyHour(stLstNo)
      'dblHourPay = GetHourPay(stLstNo, YM, iDailyHours) 'Removed by Morgan 2023/6/7
      'end 2017/7/10
      dblRestHour2 = 46
      dblOTHourTot = 0
      lngOTPayTot2 = 0
      lngOTPayTot1 = 0
      
      Do While Not .EOF
         
         If .Fields("so01") <> stLstNo Then
            'Modified by Morgan 2015/12/23 已經沒有第2家公司
            '---舊程式已刪除---
            If lngOTPayTot2 > 0 Then
               stSQL = "update salarymonth set sm11=" & dblOTHourTot & ",sm12=" & lngOTPayTot2 & ",sm28=" & lngOTPayTot1 & _
                  " where sm01='" & stLstNo & "' and sm02=" & YM
               cnnConnection.Execute stSQL, intI
               If intI = 0 Then
                  List1.AddItem time & " --> 更新 " & stLstNo & " 加班費資料失敗", 0
                  GoTo ErrHnd
               End If
            End If
            'end 2015/12/23
            
            List1.AddItem time & " --> 更新 " & stLstNo & " 加班費資料成功", 0
            ProgressBar1.Value = ProgressBar1.Value + 1
            Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
            DoEvents
            
            stLstNo = "" & .Fields("so01")
            'stLstNo2 = Left(stLstNo, 2) & "A" & Mid(stLstNo, 4) 'Removed by Morgan 2015/12/23 取消,已經沒有第2家公司
            stLstDate = "" & .Fields("so02")
            'Modified by Morgan 2017/7/10
            'dblHourPay = GetHourPay(stLstNo, YM)
            iDailyHours = GetDaiyHour(stLstNo)
            'dblHourPay = GetHourPay(stLstNo, YM, iDailyHours) 'Removed by Morgan 2023/6/7
            'end 2017/7/10
            dblRestHour2 = 46
            dblOTHourTot = 0
            lngOTPayTot2 = 0
            lngOTPayTot1 = 0
         End If
         
         '46小時以內的加班費放第二家,超過的放第一家
         
         dblOTHourH = 0
         dblOTHour = 0
         dblOTPay2 = 0
         dblOTPay1 = 0
         dblHourPay = GetHourPay(stLstNo, Left(.Fields("so02"), 6), iDailyHours) 'Added by Morgan 2023/6/7 可能會有非當月的加班費,改都以加班日的月份抓時薪
         
         '**假日加班**
         
         '假日加班不會與平日加班時數同時存在
         '假日加班前8小時*1
         '假日加班超過8小時部份改與平日加班算法相同
         dblOTHourH = Val("" & .Fields("ot2"))
         
         'Modified by Morgan 2016/12/26
         'If dblOTHourH > 0 Then
         
         '12/23起加班費改新法計算
         '1.休息日(周六):加班費計算:前2小時 x4/3, 第3-8小時 x5/3, 第9-12小時 x8/3 (未滿4小時以4小時計,以此類推)
         '2.例假日:未滿8小時以8小時計(x1),超過8小時以平日加班費計算(沒變)
         bolRestDay = False
         If .Fields("so02") >= 20161223 And dblOTHourH > 0 Then
            '休息日(周六)
            If Weekday(Format(.Fields("so02"), "@@@@/@@/@@")) = 7 Then
               bolRestDay = True
'Removed by Morgan 2018/1/24 加班單有控制,且又有新修法故此處不可設規則
'               If dblOTHourH <= 4 Then
'                  dblOTHourH = 4
'               ElseIf dblOTHourH <= 8 Then
'                  dblOTHourH = 8
'               ElseIf dblOTHourH <= 12 Then
'                  dblOTHourH = 12
'               End If
'
'            '例假日
'            Else
'               If dblOTHourH < 8 Then
'                  dblOTHourH = 8
'               End If
'end 2018/1/24
            End If
         End If
         
         '休息日(周六):1.要累計46小時, 2.加班費計算:前2小時 x4/3, 第3-8小時 x5/3, 第9-12小時 x8/3 (未滿4小時以4小時計,以此類推)
         If bolRestDay Then
            
            If dblOTHourH <= 2 Then
               dblOTPay2 = dblOTPay2 + (4 / 3) * dblOTHourH * dblHourPay
            ElseIf dblOTHourH <= 8 Then
               dblOTPay2 = dblOTPay2 + (4 / 3) * 2 * dblHourPay
               dblOTPay2 = dblOTPay2 + (5 / 3) * (dblOTHourH - 2) * dblHourPay
            Else
               dblOTPay2 = dblOTPay2 + (4 / 3) * 2 * dblHourPay
               dblOTPay2 = dblOTPay2 + (5 / 3) * 6 * dblHourPay
               dblOTPay2 = dblOTPay2 + (8 / 3) * (dblOTHourH - 8) * dblHourPay
            End If
            
            '應稅加班費
            '本日加班時數超過剩餘免稅加班時數才有應稅加班費
            If dblOTHourH > dblRestHour2 Then
               '剩餘超過8小時
               If dblRestHour2 > 8 Then
                  dblOTPay1 = (8 / 3) * (dblOTHourH - dblRestHour2) * dblHourPay
               '剩餘超過2小時未滿8小時
               ElseIf dblRestHour2 > 2 Then
                  dblOTPay1 = dblOTPay2 - (4 / 3) * 2 * dblHourPay - (5 / 3) * (dblRestHour2 - 2) * dblHourPay
               '有剩餘但未超過2小時
               ElseIf dblRestHour2 > 0 Then
                  dblOTPay1 = dblOTPay2 - (4 / 3) * dblRestHour2 * dblHourPay
               '沒剩餘(全部都應稅)
               Else
                  dblOTPay1 = dblOTPay2
               End If
            End If
            
            dblRestHour2 = dblRestHour2 - dblOTHourH '剩餘免稅加班時數(46小時剩餘時數)
            dblOTHourTot = dblOTHourTot + dblOTHourH '加班時數累計
         
         '例假日
         ElseIf dblOTHourH > 0 Then
         'end 2016/12/26
         
            '假日加班前8小時 x1 不算46小時累計
            If dblOTHourH > 8 Then
               dblOTHour = 8
            Else
               dblOTHour = dblOTHourH
            End If
            
            'Modified by Morgan 2015/12/23 假日加班前8小時不計入46小時計算(免所得稅)
            '---舊程式已刪除---
            dblOTPay2 = dblOTHour * dblHourPay
            'end 2015/12/23
            
            dblOTHourTot = dblOTHourTot + dblOTHour '加班時數累計
            
            '8小時以後的加班時數以平日加班費計算(併入下面的平日加班費計算)
            dblOTHour = dblOTHourH - dblOTHour
            
         End If
         
         '***********
         
         '平日(兩小時以內*1.33,超過兩小時部分*1.66)
         
         'Modified by Morgan 2014/10/30 兩小時以內*4/3,超過兩小時部分*5/3
         dblOTHour = dblOTHour + Val("" & .Fields("ot1")) '假日8小時以後的加班時數數+平日加班時數(不會同時有時數)
         If dblOTHour > 0 Then
            'Modified by Morgan 2015/12/23 假日加班超過8小時以平日加班費計算
            '---舊程式已刪除---
            '兩小時以內
            If dblOTHour <= 2 Then
               dblOTPay2 = dblOTPay2 + (4 / 3) * dblOTHour * dblHourPay
               'Modified by Morgan 2016/8/15 double 會有極小的誤差
               'If dblOTHour > dblRestHour2 Then
               If Round(dblOTHour - dblRestHour2, 3) > 0 Then
               'end 2016/8/15
                  '已超過46小時
                  If dblRestHour2 < 0 Then
                     dblOTPay1 = dblOTPay1 + (4 / 3) * dblOTHour * dblHourPay
                  '部分超過46小時
                  Else
                     dblOTPay1 = dblOTPay1 + (4 / 3) * (dblOTHour - dblRestHour2) * dblHourPay
                  End If
               End If
            Else
               dblOTPay2 = dblOTPay2 + ((4 / 3) * 2 + (5 / 3) * (dblOTHour - 2)) * dblHourPay
               'Modified by Morgan 2016/8/15 double 會有極小的誤差
               'If dblOTHour > dblRestHour2 Then
               If Round(dblOTHour - dblRestHour2, 3) > 0 Then
               'end 2016/8/15
                  '已超過46小時
                  If dblRestHour2 < 0 Then
                     dblOTPay1 = dblOTPay1 + ((4 / 3) * 2 + (5 / 3) * (dblOTHour - 2)) * dblHourPay
                  '2小時內超過46小時
                  ElseIf dblRestHour2 < 2 Then
                     dblOTPay1 = dblOTPay1 + ((4 / 3) * (2 - dblRestHour2) + (5 / 3) * (dblOTHour - 2)) * dblHourPay
                  '2小時後超過46小時
                  Else
                     dblOTPay1 = dblOTPay1 + (5 / 3) * (dblOTHour - dblRestHour2) * dblHourPay
                  End If
               End If
            End If
            'end 2015/12/23
         End If
         'end 2014/10/30
         
         dblRestHour2 = dblRestHour2 - dblOTHour '剩餘免稅加班時數(46小時剩餘時數)
         dblOTHourTot = dblOTHourTot + dblOTHour '加班時數累計
         
         '以每天為單位去四捨五入
         'Modified by Morgan 2014/10/30 每日無條件進位到整數
         'lngOTPayTot2 = lngOTPayTot2 + Round(dblOTPay2)
         'lngOTPayTot1 = lngOTPayTot1 + Round(dblOTPay1)
         lngOTPayTot2 = lngOTPayTot2 + -1 * Int(-1 * dblOTPay2)
         lngOTPayTot1 = lngOTPayTot1 + -1 * Int(-1 * dblOTPay1)
         'end 2014/10/30
         
         .MoveNext
      Loop
      
      'Modified by Morgan 2015/12/23 已經沒有第2家公司
      '---舊程式已刪除---
      If lngOTPayTot2 > 0 Then
         stSQL = "update salarymonth set sm11=" & dblOTHourTot & ",sm12=" & lngOTPayTot2 & ",sm28=" & lngOTPayTot1 & _
            " where sm01='" & stLstNo & "' and sm02=" & YM
         cnnConnection.Execute stSQL, intI
         If intI = 0 Then
            List1.AddItem time & " --> 更新 " & stLstNo & " 加班費資料失敗", 0
            GoTo ErrHnd
         End If
      End If
      'end 2015/12/23

      List1.AddItem time & " --> 更新 " & stLstNo & " 加班費資料成功", 0
      ProgressBar1.Value = ProgressBar1.Value + 1
      Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      DoEvents
         
      End With
   End If
   
   List1.AddItem time & " --> 更新加班費資料結束,共 " & ProgressBar1.max & " 筆", 0
   

   List1.AddItem time & " --> 更新缺勤未打卡扣款資料開始...", 0
   DoEvents
   'Modify by Morgan 2009/5/26 若公司合併時改扣第一家公司(因為勞退隔月才申報所以當月會有沒薪資但有勞退的情形)
   'Modify by Morgan 2009/3/9 病假超過30日則超出的部份扣全薪
   'x1:忘記打卡次數, x2:遲到次數, x3:曠職時數, y1:事假時數, y2:病假時數
   'Modified by Morgan 2014/12/12
   '增加 20生理假(半薪,與病假合計超過30天時前3天半薪), 22家庭照顧假(不給薪)
   'Modified by Morgan 2020/2/4 +24防疫照顧假(不給薪)
   'Modified by Morgan 2020/10/7 修正特殊工時(非8小時)問題
   'stSQL = "select sd01,x1,x2,x3,y1,y2,y3 from salarydata" & _
      ",(select sa01 x0,sum(sa03) x1,sum(sa04) x2,sum(" & GetDaiyHourSQL("sa01") & "*nvl(sa05,0)+nvl(sa06,0)) x3" & _
      " from Staff_Assist where substr(sa02,1,6)=" & YM & " group by sa01) x" & _
      ",(select sa01 y0,sum(decode(sign(instr('05,22,24',sa06)),1," & GetDaiyHourSQL("sa01") & "*nvl(sa07,0)+nvl(sa08,0))) y1" & _
      ",sum(decode(sa06,'06'," & GetDaiyHourSQL("sa01") & "*nvl(sa07,0)+nvl(sa08,0))) y2" & _
      ",sum(decode(sa06,'20'," & GetDaiyHourSQL("sa01") & "*nvl(sa07,0)+nvl(sa08,0))) y3" & _
      " from Staff_Absence where substr(sa02,1,6)=" & YM & " and sa06 in ('05','06','20','22','24')" & _
      " group by sa01) y where x0(+)=sd01 and y0(+)=sd01 and (x0 is not null or y0 is not null) order by 1 asc"
   'Modified by Morgan 2023/6/6 薪資計算後才新增的請假資料算到隔月薪資扣款，故改為該月前未扣過的都抓
   'stSQL = "select sd01,x1,x2,x3,x3_1,y1,y1_1,y2,y2_1,y3,y3_1 from salarydata" & _
      ",(select sa01 x0,sum(sa03) x1,sum(sa04) x2,sum(sa06) x3,sum(sa05) x3_1" & _
      " from Staff_Assist where substr(sa02,1,6)=" & YM & " group by sa01) x" & _
      ",(select sa01 y0,sum(decode(sign(instr('05,22,24',sa06)),1,sa08)) y1" & _
      ",sum(decode(sign(instr('05,22,24',sa06)),1,sa07)) y1_1" & _
      ",sum(decode(sa06,'06',sa08)) y2,sum(decode(sa06,'06',sa07)) y2_1" & _
      ",sum(decode(sa06,'20',sa08)) y3,sum(decode(sa06,'20',sa07)) y3_1" & _
      " from Staff_Absence where substr(sa02,1,6)=" & YM & " and sa06 in ('05','06','20','22','24')" & _
      " group by sa01) y where x0(+)=sd01 and y0(+)=sd01 and (x0 is not null or y0 is not null) order by 1 asc"
   stSQL = "update Staff_Absence set sa18=" & YM & " where substr(sa02,1,6)<=" & YM & " and ( sa18=0 or sa18=" & YM & ")"
   cnnConnection.Execute stSQL, intR
   
   'Modified by Morgan 2024/7/29 +天災不給薪(25)
   'Modified by Morgan 2025/8/15 修改曠職時間改以「分」計算
   If strSrvDate(1) >= 曠職以分計啟用日 Then
      stSQL = "select sa01,min(sa02) sa02,sum(sa03) x1,sum(sa04) x2,sum(sa06/60) x3,sum(sa05) x3_1,0 y1,0 y1_1,0 y2,0 y2_1,0 y3,0 y3_1" & _
         " from Staff_Assist where substr(sa02,1,6)=" & YM & " group by sa01"
   Else
      stSQL = "select sa01,min(sa02) sa02,sum(sa03) x1,sum(sa04) x2,sum(sa06) x3,sum(sa05) x3_1,0 y1,0 y1_1,0 y2,0 y2_1,0 y3,0 y3_1" & _
         " from Staff_Assist where substr(sa02,1,6)=" & YM & " group by sa01"
   End If
   'end 2025/8/15
   
   stSQL = stSQL & " union select sa01,sa02,0 x1,0 x2,0 x3,0 x3_1" & _
      ",decode(sa06,'06',0,'20',0,sa08) y1,decode(sa06,'06',0,'20',0,sa07) y1_1" & _
      ",decode(sa06,'06',sa08) y2,decode(sa06,'06',sa07) y2_1" & _
      ",decode(sa06,'20',sa08) y3,decode(sa06,'20',sa07) y3_1" & _
      " from Staff_Absence where substr(sa02,1,6)<=" & YM & " and SA18=" & YM & " and sa06 in ('05','06','20','22','24','25')" & _
      " order by 1,2"
   'end 2023/6/6
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL, , True)
   If intR = 1 Then
      With adoRst
      .MoveFirst
      stLstNo = ""
      '請假人數
      intR = 0
      Do While Not .EOF
         If .Fields("sa01") <> stLstNo Then
            intR = intR + 1
            stLstNo = "" & .Fields("sa01")
         End If
         .MoveNext
      Loop
      ProgressBar1.max = intR
      ProgressBar1.Value = 0
      Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      
      .MoveFirst
      
      lngSM21 = 0
      lngSM22 = 0
      Do While Not .EOF
         'Modified by Morgan 2023/6/7
         'stLstNo = "" & .Fields("sd01")
         stLstNo = "" & .Fields("sa01")
         'end 2023/6/7
         
         'stLstNo2 = Left(stLstNo, 2) & "A" & Mid(stLstNo, 4) 'Removed by Morgan 2015/12/23 取消,已經沒有第2家公司
         'Modified by Morgan 2017/7/10
         'dblHourPay = GetHourPay(stLstNo, YM)
         iDailyHours = GetDaiyHour(stLstNo)
         'Modified by Morgan 2023/6/7 可能會有非當月的請假單,改都以請假日的月份抓時薪
         'dblHourPay = GetHourPay(stLstNo, YM, iDailyHours)
         dblHourPay = GetHourPay(stLstNo, Left(.Fields("sa02"), 6), iDailyHours)
         'end 2023/6/7
         'end 2017/7/10
         '未打卡扣款
         'Removed by Morgan 2022/6/1 取消扣款--劉經理
         'lngSM22 = 10 * Val("" & .Fields("x1"))
         'end 2022/6/1
         
         'Modified by Morgan 2020/10/7 修正特殊工時(非8小時)問題
         'dblHour = Val("" & .Fields("x3")) + Val("" & .Fields("y1"))
         dblHour = Val("" & .Fields("x3")) + Val("" & .Fields("y1")) + iDailyHours * (Val("" & .Fields("x3_1")) + Val("" & .Fields("y1_1")))
         
         'Added by Morgan 2020/10/7
         dblSickHr = Val("" & .Fields("y2")) + iDailyHours * Val("" & .Fields("y2_1"))
         dblGirlSickHr = Val("" & .Fields("y3")) + iDailyHours * Val("" & .Fields("y3_1"))
         
         'Add by Morgan 2009/3/9
         'Modified by Morgan 2014/12/12 +生理假(半薪,與病假合計超過30天時前3天扣半薪)
         '病假扣半薪時數(超過扣全薪)
'         dblSickBaseHour = 30 * iDailyHours
'         'Add by Morgan 2009/3/9
'         If Val("" & .Fields("y2")) > 0 Then
'            dblSickHour = GetSickHour(.Fields("sd01"), (Val(YM) \ 100) * 10000, Val(YM) * 100)
'            '上月累計已超過30日,扣全薪
'            If dblSickHour > dblSickBaseHour Then
'               dblHour = dblHour + Val("" & .Fields("y2"))
'            '累計本月未超過30日,扣半薪
'            ElseIf dblSickHour + Val("" & .Fields("y2")) < dblSickBaseHour Then
'               dblHour = dblHour + 0.5 * Val("" & .Fields("y2"))
'            '累計本月超過30日,超過部份扣全薪
'            Else
'               dblHour = dblHour + 0.5 * (dblSickBaseHour - dblSickHour) + (Val("" & .Fields("y2")) + dblSickHour - dblSickBaseHour)
'            End If
'         End If
         
         'Modified by Morgan 2020/10/7 修正特殊工時(非8小時)問題
         'If Val("" & .Fields("y2")) + Val("" & .Fields("y3")) > 0 Then
         If dblSickHr + dblGirlSickHr > 0 Then
         
            '病假扣半薪時數(超過扣全薪)
            dblSickBaseHour = 30 * iDailyHours
            '生理假扣半薪時數(超過且合併病假超過33天扣全薪)
            dblGirlSickBaseHour = 3 * iDailyHours
         
            'Modified by Morgan 2023/6/7
            'dblSickHour = GetSickHour(.Fields("sd01"), (Val(YM) \ 100) * 10000, Val(YM) * 100) '已請病假
            'dblGirlSickHour = GetSickHour(.Fields("sd01"), (Val(YM) \ 100) * 10000, Val(YM) * 100, True) '已請生理假
            dblSickHour = GetSickHour(.Fields("sa01"), (Val(YM) \ 100) * 10000, .Fields("sa02") - 1) '已請病假
            dblGirlSickHour = GetSickHour(.Fields("sa01"), (Val(YM) \ 100) * 10000, .Fields("sa02") - 1, True)   '已請生理假
            'end 2023/6/7
            
            '已請超過3天之生理假天數
            If dblGirlSickHour > dblGirlSickBaseHour Then
               dblGrilSickMergeHour = dblGirlSickHour - dblGirlSickBaseHour
            Else
               dblGrilSickMergeHour = 0
            End If
            
            '未計算生理假=本月生理假
            dblGrilSickRemainHour = dblGirlSickHr
            '生理假,前3天扣半薪
            If dblGrilSickRemainHour > 0 Then
               '前月累計未超過3天
               If dblGirlSickHour < dblGirlSickBaseHour Then
                  '累計未超過3天部分扣半薪,超過部分要併入病假考慮
                  If dblGrilSickRemainHour + dblGirlSickHour <= dblGirlSickBaseHour Then
                     dblHour = dblHour + 0.5 * dblGrilSickRemainHour
                     dblGrilSickRemainHour = 0
                  '有請非整天時才會執行到
                  Else
                     dblHour = dblHour + 0.5 * (dblGirlSickBaseHour - dblGirlSickHour)
                     dblGrilSickRemainHour = dblGrilSickRemainHour - (dblGirlSickBaseHour - dblGirlSickHour)
                  End If
               End If
            
               '生理假超過3天部分
               If dblGrilSickRemainHour > 0 Then
                  '合併病假超過30天扣全薪
                  If dblSickHour + dblGrilSickMergeHour >= dblSickBaseHour Then
                     dblHour = dblHour + dblGrilSickRemainHour
                  '合併病假未超過30天扣半薪
                  ElseIf dblSickHour + dblGrilSickMergeHour + dblGrilSickRemainHour <= dblSickBaseHour Then
                     dblHour = dblHour + 0.5 * dblGrilSickRemainHour
                  Else
                     dblHour = dblHour + 0.5 * (dblSickBaseHour - (dblSickHour + dblGrilSickMergeHour))
                     dblHour = dblHour + (dblGrilSickRemainHour - (dblSickBaseHour - (dblSickHour + dblGrilSickMergeHour)))
                  End If
                  dblGrilSickMergeHour = dblGrilSickMergeHour + dblGrilSickRemainHour
                  dblGrilSickRemainHour = 0
               End If
            End If
               
            '病假
            'Modified by Morgan 2020/10/7 修正特殊工時(非8小時)問題
            'If Val("" & .Fields("y2")) > 0 Then
            If dblSickHr > 0 Then
            
               '合併生理假超過30天扣全薪
               If dblSickHour + dblGrilSickMergeHour >= dblSickBaseHour Then
                  dblHour = dblHour + dblSickHr
               Else
                  '累計未超過30天扣半薪,超過部分扣全薪
                  If dblSickHour + dblGrilSickMergeHour + dblSickHr <= dblSickBaseHour Then
                     dblHour = dblHour + 0.5 * dblSickHr
                  Else
                     dblHour = dblHour + 0.5 * (dblSickBaseHour - (dblSickHour + dblGrilSickMergeHour))
                     dblHour = dblHour + dblSickHr - (dblSickBaseHour - (dblSickHour + dblGrilSickMergeHour))
                  End If
               End If
            End If
         End If
         'end 2014/12/12
         
         If Val("" & .Fields("x2")) > 2 Then
            'Modified by Morgan 2012/12/18 改前兩次不扣薪
            'dblHour = dblHour + 0.5 * Val("" & .Fields("x2"))
            dblHour = dblHour + 0.5 * (Val("" & .Fields("x2")) - 2)
         End If
         
         'Modified by Morgan 2023/6/7
         'lngSM21 = Round(dblHour * dblHourPay)
         lngSM21 = lngSM21 + dblHour * dblHourPay
         
         .MoveNext
         intR = 0
         If .EOF Then
            intR = 1
         ElseIf .Fields("sa01") <> stLstNo Then
            intR = 1
         End If
         If intR = 1 Then
         'end 2023/6/7
            '缺勤扣第二家
            If lngSM21 > 0 Then
               'Modified by Morgan 2023/9/28 改無條件捨去
               'lngSM21 = Round(lngSM21)
               lngSM21 = Trunc(lngSM21)
               'end 2023/9/28
               'Modify by Morgan 2009/5/26 判斷無基本薪資為合併則改扣第一家
               'Modified by Morgan 2015/12/23 取消,已經沒有第2家公司
               'stSQL = "update salarymonth set sm21=" & lngSM21 & _
               '   " where sm01='" & stLstNo2 & "' and sm02=" & YM & " and sm04>0"
               'cnnConnection.Execute stSQL, intR
               ''無第二家改放第一家
               'If intR = 0 Then
                  stSQL = "update salarymonth set sm21=" & lngSM21 & _
                     " where sm01='" & stLstNo & "' and sm02=" & YM
                  cnnConnection.Execute stSQL, intR
                  If intR = 0 Then
                     List1.AddItem time & " --> 更新 " & stLstNo & " 缺勤扣款資料失敗(該月無薪資可扣)", 0
                     GoTo ErrHnd
                  End If
               'End If
               'end 2015/12/23
            End If
   
            '未打卡扣第一家
            If lngSM22 > 0 Then
               stSQL = "update salarymonth set sm22=" & lngSM22 & _
                  " where sm01='" & stLstNo & "' and sm02=" & YM
               cnnConnection.Execute stSQL, intR
               If intR = 0 Then
                  List1.AddItem time & " --> 更新 " & stLstNo & " 未打卡扣款資料失敗(該月無薪資可扣)", 0
                  GoTo ErrHnd
               End If
            End If
   
            List1.AddItem time & " --> 更新 " & stLstNo & " 缺勤未打卡扣款資料成功", 0
            ProgressBar1.Value = ProgressBar1.Value + 1
            Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
            DoEvents
               
         'Modified by Morgan 2023/6/7
         '.MoveNext
            lngSM21 = 0
            lngSM22 = 0
         End If
         'end 2023/6/7
      Loop
      End With
   End If
   
   List1.AddItem time & " --> 更新缺勤未打卡扣款資料結束,共 " & ProgressBar1.max & " 筆", 0

   'Modified by Morgan 2022/10/5 婚喪扣款改以扣款計算日的月份為扣款薪資月份
   'List1.AddItem time & " --> 更新婚喪戶助會扣款資料開始...", 0
   'DoEvents
   'stSQL = "update salarymonth set (sm17,sm18)=(select sum(decode(wfa05,'1',wfa04,0)) s1,sum(decode(wfa05,'2',wfa04,0)) s2" & _
      " from WFAmount where substr(wfa01,1,6)=sm02 and wfa03=sm01)" & _
      " where sm02=" & YM & " and sm01 in (select wfa03 from WFAmount where SUBSTR(wfa01,1,6)=" & YM & ")"
   'cnnConnection.Execute stSQL, intR
   'List1.AddItem time & " --> 更新婚喪戶助會扣款資料結束共 " & intR & " 筆", 0
   
'Removed by Morgan 2025/7/29 114/7/28起廢止婚喪互助辦法
'   List1.AddItem time & " --> 更新婚喪扣款資料開始...", 0
'   DoEvents
'   '婚喪扣款
'   stSQL = "update salarymonth set sm17=(select nvl(sum(wfa04),0) s1" & _
'      " from WeddingAndFuneral,WFAmount where substr(wf04,1,6)=sm02 and wfa01(+)=wf01 and wfa02(+)=wf02 and wfa03=sm01 and wfa05='1')" & _
'      " where sm02=" & YM
'   cnnConnection.Execute stSQL, intR
'   List1.AddItem time & " --> 更新婚喪扣款資料結束共 " & intR & " 筆", 0
'end 2025/7/29
   
   List1.AddItem time & " --> 更新戶助會扣款資料開始...", 0
   DoEvents
   stSQL = "update salarymonth set sm18=(select nvl(sum(wfa04),0) s2" & _
      " from WFAmount where substr(wfa01,1,6)=sm02 and wfa03=sm01 and wfa05='2')" & _
      " where sm02=" & YM
   cnnConnection.Execute stSQL, intR
   List1.AddItem time & " --> 更新戶助會扣款資料結束共 " & intR & " 筆", 0
   'end 2022/10/5
   
   List1.AddItem time & " --> 更新其他所得、其他所得稅金及其他扣款資料開始...", 0
   
   '排除翻譯費01
   'Modify by Morgan 2009/6/23 排除第二家勞保費12
   'Modified by Morgan 2024/5/13 排除翻譯預支39
   stSQL = "update salarymonth set (sm13,sm23,sm29)" & _
      "=(select sum(decode(oc02,'A',od05)),sum(decode(oc02,'D',od05)),sum(od06)" & _
      " From OtherSalaryData, OtherSalaryCode" & _
      " where od03=sm01 and substr(od02,1,6)=sm02 and od04<>'01' and od04<>'12' and od04<>'39' and oc01(+)=od04" & _
      ") where sm02=" & YM & " and sm01 in (select od03 from OtherSalaryData where substr(od02,1,6)=" & YM & " and od04<>'01' and od04<>'12')"
      
   cnnConnection.Execute stSQL, intR
   
   List1.AddItem time & " --> 更新其他所得扣款資料結束共 " & intR & " 筆", 0
   
   
   'Add by Morgan 2009/6/23 更新第二家勞保費
   List1.AddItem time & " --> 更新第二家勞保費資料開始...", 0
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'stSQL = "update salarymonth set sm14" & _
      "=(select sum(od05) From OtherSalaryData where od03=replace(sm01,'A','0') and substr(od02,1,6)=sm02 and od04='12')" & _
      " where sm02=" & YM & " and sm01 in (select substr(od03,1,2)||'A'||substr(od03,4,2) from OtherSalaryData where substr(od02,1,6)=" & YM & " and od04='12')"
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   stSQL = "update salarymonth set sm14" & _
      "=(select sum(od05) From OtherSalaryData where od03=substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4) and substr(od02,1,6)=sm02 and od04='12')" & _
      " where sm02=" & YM & " and sm01 in (select substr(od03,1,2)||'A'||substr(od03,4,2) from OtherSalaryData where substr(od02,1,6)=" & YM & " and od04='12')"
      
   cnnConnection.Execute stSQL, intR
   
   List1.AddItem time & " --> 更新第二家勞保費資料結束共 " & intR & " 筆", 0
   
   
   List1.AddItem time & " --> 更新借支還款金額開始...", 0
   DoEvents
   stSQL = "update salarymonth set sm20" & _
      "=(select sum(ae03) from Advance_Employee where ae01=sm01 and ae04=sm02)" & _
      " where sm02=" & YM & " and sm01 in (select ae01 from Advance_Employee where ae04=" & YM & ")"
      
   cnnConnection.Execute stSQL, intR
   
   List1.AddItem time & " --> 更新借支還款金額結束共 " & intR & " 筆", 0
   
   List1.AddItem time & " --> 更新貸款還款金額開始...", 0
   DoEvents
   'Modify by Morgan 2010/10/12 第一個月償還金額=總額-月償還金額*(期數-1)
   stSQL = "update salarymonth set sm19" & _
      "=(select sum(decode(le05,sm02,le03+nvl(le04,0)-le07*(12*(substr(le06,1,4)-substr(le05,1,4))+(substr(le06,5)-substr(le05,5)))" & _
      ",le07)) from Loan_Employee where le05<=sm02 and le06>=sm02 and le01=sm01)" & _
      " where sm02=" & YM & " and sm01 in (select le01 from Loan_Employee where le05<=" & YM & " and le06>=" & YM & ")"
      
   cnnConnection.Execute stSQL, intR
   
   List1.AddItem time & " --> 更新貸款還款金額結束,共 " & intR & " 筆", 0
   
   List1.AddItem time & " --> 更新員工貸款資料開始...", 0
   DoEvents
   '若分期後除不盡的第一月還款金額由人工修改
   'Modify by Morgan 2010/10/12 第一個月償還金額=總額-月償還金額*(期數-1)
   stSQL = "update Loan_Employee set le08=le07*(12*(substr(le06,1,4)-substr(le05,1,4))+(substr(le06,5)-substr(le05,5)))" & _
      ",le09=" & YM & " where le05=" & YM
      
   cnnConnection.Execute stSQL, intR
   
   stSQL = "update Loan_Employee set le08=nvl(le08,0)-nvl(le07,0),le09=" & YM & _
      " where le05<" & YM & " and le06>=" & YM & " and le09<" & YM
      
   cnnConnection.Execute stSQL, intR
   
   List1.AddItem time & " --> 更新員工貸款資料結束,共 " & intR & " 筆", 0
   
   '更新所得稅
   List1.AddItem time & " --> 更新所得稅資料...", 0
   
   stSQL = "select sm01 from salarymonth where sm02=" & YM & " and sm01<'F'"
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL, , True)
   If intR = 1 Then
      With adoRst
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      Do While Not .EOF
         List1.AddItem time & " --> 更新 " & .Fields(0) & " 所得稅...", 0
         '抓薪資基本檔編制判斷是否為兼職人員
         UpdTax "" & .Fields(0), YM
         ProgressBar1.Value = ProgressBar1.Value + 1
         Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         .MoveNext
      Loop
      End With
   End If
   
   List1.AddItem time & " --> 更新所得稅資料結束,共 " & ProgressBar1.max & " 筆", 0

   'Added by Morgan 2022/11/18
   '差旅/技術/房租津貼轉入每月獎金
   BatchMonthBonus
   'end 2022/11/18
   
   'Added by Morgan 2013/1/30
   '計算代扣日期為當月的補充保費
   List1.AddItem time & " --> 計算每月獎金補充保費開始...", 0
   DoEvents
      
   '已離職員工若有未計算之補充保費則代扣日期更新為當月底
   stSQL = "update MonthBonus set mb13=" & GetLastDay(YM & "01") & " where mb13>" & YM & "31" & _
      " and exists(select * from staff where st01=mb02 and st04='2')"
   cnnConnection.Execute stSQL, intR
   
   stSQL = "select * from MonthBonus where mb13>=" & YM & "01 and mb13<=" & YM & "31 order by mb02,mb01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      With RsTemp
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      Do While Not .EOF
         List1.AddItem time & " --> 計算 " & .Fields("mb02") & "-" & .Fields("mb01") & " 其他獎金補充保費", 0
         ProgressBar1.Value = ProgressBar1.Value + 1
         Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         If .Fields("mb02") = stNHI(1) Then
            intCount = intCount + 1
         Else
            intCount = 0
         End If
         'Added by Morgan 2014/7/31
         For intI = LBound(stNHI) To UBound(stNHI)
            stNHI(intI) = ""
         Next
         'end 2014/7/31
         stNHI(1) = .Fields("mb02")
         stNHI(2) = .Fields("mb13")
         stNHI(3) = "50"
         stNHI(4) = "6"
         stNHI(5) = ""
         stNHI(6) = ""
         stNHI(7) = .Fields("mb03")
         stNHI(8) = ""
         stNHI(10) = Format(235910 + intCount)
         stNHI(11) = .Fields("mb11") 'Added by Morgan 2013/2/26
         stNHI(14) = .Fields("mb01") 'Added by Morgan 2013/4/24
         'Modified by Morgan 2013/3/12 +NHI13
         PUB_NHI2nd stNHI(1), stNHI(2), stNHI(3), stNHI(4), stNHI(7), stNHI(5), stNHI(6), stNHI(8), stNHI(10), stNHI(11), stNHI(13)
         stSQL = "update MonthBonus set mb12=" & Val(stNHI(6)) & " where mb01=" & .Fields("mb01") & " and mb02='" & .Fields("mb02") & "'"
         cnnConnection.Execute stSQL, intR
         PUB_InsertNHI2nd stNHI
         .MoveNext
         
      Loop
      intR = .RecordCount
      End With
   End If
   List1.AddItem time & " --> 計算每月獎金補充保費結束,共 " & intR & " 筆", 0
   'end 2013/1/30
   
   'Added by Morgan 2013/1/30
   '不可在翻譯費之後否則會覆蓋 sm43
   List1.AddItem time & " --> 更新補充保費資料開始...", 0
   DoEvents
   
   '新增非投保公司月薪資之補充保費(目前只有68099)
   'Modified by Morgan 2017/12/26 因當月非最後1日離職者沒有健保投保金額(sm42)故加判斷薪資檔沒有投保金額(sd47)者
   'Modify By Sindy 2020/6/24 and sm05>0 => and (sm05>0 or sm45>0)
   'Modified by Morgan 2022/5/25 因A5011為外派但超過2年健保被退故再加判斷沒有勞健保投保薪資(sd45)者
   stSQL = "select * from salarymonth,salarydata where sm02=" & YM & " and (sm05>0 or sm45>0) and nvl(sm42,0)=0 and sd01(+)=sm01 and nvl(sd47,0)=0 and nvl(sd45,0)=0"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         'Added by Morgan 2014/7/31
         For intI = LBound(stNHI) To UBound(stNHI)
            stNHI(intI) = ""
         Next
         'end 2014/7/31
         stNHI(1) = .Fields("sm01")
         stNHI(2) = GetLastDay(YM & "01")
         stNHI(3) = "50"
         stNHI(4) = "7"
         stNHI(5) = ""
         stNHI(6) = ""
         'Modify By Sindy 2020/6/24 + Val("" & .Fields("sm45"))
         stNHI(7) = Val("" & .Fields("sm05")) + Val("" & .Fields("sm45"))
         stNHI(8) = ""
         stNHI(10) = "235930"
         stNHI(11) = .Fields("sm37") 'Added by Morgan 2013/2/26
         PUB_NHI2nd stNHI(1), stNHI(2), stNHI(3), stNHI(4), stNHI(7), stNHI(5), stNHI(6), stNHI(8), stNHI(10), stNHI(11), stNHI(13) 'Modified by Morgan 2013/3/12 +NHI13
         PUB_InsertNHI2nd stNHI
         .MoveNext
      Loop
      End With
   End If
   
   '更新月薪資補充保費
   stSQL = "update salarymonth set sm43=(select nvl(sum(nhi06),0) from nhi2nd" & _
      " where nhi01=sm01 and substr(nhi02,1,6)=sm02 and nhi04 in('3','5','6','7')) where SM02=" & YM
      
   cnnConnection.Execute stSQL, intR
   List1.AddItem time & " --> 更新補充保費資料結束,共 " & intR & " 筆", 0
   'end 2013/1/30
   
   List1.AddItem time & " --> 新增翻譯所得資料開始...", 0
   DoEvents
   'Modify by Morgan 2009/2/4 稅金超過2000才扣的控制改在輸入所得時做(允許人工輸入2000以下)
   'Modified by Morgan 2013/1/30 +sm43 補充保費
   'Modified by Morgan 2024/5/13 排除專職翻譯同仁(最後1碼>=A)，因規則特別會在下面另外計算
   stSQL = "insert into salarymonth(sm01,sm02,sm03,sm04,sm24,sm37,sm43)" & _
      " select sd01 sm01," & YM & " sm02," & stDepCol & " sm03,s01 sm04,trunc(s02) sm24,sd19 sm37,s03 sm43 from (" & _
      " select od03 ,sum(od05) s01,sum(od06) s02,sum(od13) s03 from OtherSalaryData" & _
      " where substr(od02,1,6)=" & YM & " and od04='01'" & _
      " group by od03) s,salarydata,staff s1" & _
      " where sd01(+)=od03 and st01(+)=od03" & _
      " and not exists(select * from salarymonth,staff s2 where sm02=" & YM & " and substr(sm01,-2)>='9A' and st01(+)=sm01" & _
      " and st26=s1.st26)"
   
   cnnConnection.Execute stSQL, intR
   
   List1.AddItem time & " --> 新增翻譯所得資料結束,共 " & intR & " 筆", 0
   
   'Added by Morgan 2024/5/13
   List1.AddItem time & " --> 新增專職翻譯同仁翻譯所得資料開始...", 0
   DoEvents
   
   '刪除已產生的次月翻譯預支39
   stSQL = "DELETE othersalarydata WHERE od02=" & NextYM & "01 AND od04='39'"
   cnnConnection.Execute stSQL, intR
               
   '專職翻譯同仁(最後1碼>=A)
   '超過約定薪資的部分轉列翻譯費 (計算四倍補充保費)
   stSQL = "select sm01,sm37,sm04+nvl(sm07,0) ps,o.* from salarymonth,staff s1,staff s2,OtherSalaryData o" & _
      " where sm02=" & YM & " and substr(sm01,-2)>='9A' and s1.st01(+)=sm01" & _
      " and s2.st26(+)=s1.st26 and od03(+)=s2.st01 and substr(od02,1,6)=sm02 and od04='01'" & _
      " order by sm01,od03,od02,od01"
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL, , True)
   If intR = 1 Then
      stLstNo = ""
      dblTFeeNet = 0
      With adoRst
      Do While Not .EOF
         '不足額新增次月翻譯預支39，不轉月薪資，不列補充保費，
         If .Fields("sm01") <> stLstNo Then
            If stLstNo <> "" Then
               '不足額新增次月翻譯預支39
               If dblTFeeNet > 0 Then
                  stSQL = "INSERT INTO othersalarydata(od01,od02,od03,od04,od05)" & _
                        " select od01," & NextYM & "01,'" & stLstNo & "','39'," & dblTFeeNet & _
                        " from (select max(OD01)+1 od01 from othersalarydata ) X"
                  cnnConnection.Execute stSQL, intR
               '超出部分轉月薪資
               ElseIf dblTFeeNet < 0 Then
                  stSQL = "insert into salarymonth(sm01,sm02,sm03,sm04,sm24,sm37,sm43)" & _
                     " select sd01 sm01," & YM & " sm02," & stDepCol & " sm03," & Abs(dblTFeeNet) & " sm04,s02 sm24,sd19 sm37,s03 sm43 from (" & _
                     " select od03 ,sum(od05) s01,sum(od06) s02,sum(od13) s03 from OtherSalaryData" & _
                     " where substr(od02,1,6)=" & YM & " and od04='01' and od03='" & stLstFNo & "'" & _
                     " group by od03) s,salarydata,staff s1" & _
                     " where sd01(+)=od03 and st01(+)=od03"
                  cnnConnection.Execute stSQL, intR
               End If
            End If

            dblTFeeNet = .Fields("ps") '約定薪資
            '翻譯預支39
            stSQL = "select nvl(sum(od05),0) pp from OtherSalaryData where substr(od02,1,6)=" & YM & " and od04='39' and od03='" & .Fields("sm01") & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
            If intI = 1 Then
               dblTFeeNet = dblTFeeNet + RsTemp(0)
            End If
            stLstNo = .Fields("sm01")
            
         End If
         
         If stLstFNo <> .Fields("od03") Then stLstFNo = .Fields("od03")
         
         '刪除已產生的補充保費紀錄
         stSQL = "DELETE NHI2ND WHERE NHI01='" & .Fields("od03") & "' AND NHI02=" & .Fields("od02") & " AND NHI03='50' AND NHI04='4'"
         cnnConnection.Execute stSQL, intR
         
         
         For intI = LBound(stNHI) To UBound(stNHI)
            stNHI(intI) = ""
         Next
         
         '若該月多筆翻譯費且前1筆已經超過約定薪資+上月翻譯預支時，本次翻譯費要全列補充保費
         If dblTFeeNet < 0 Then stNHI(7) = .Fields("od05")
         
         dblTFeeNet = dblTFeeNet - .Fields("od05")
         
         '超出部分列補充保費
         If dblTFeeNet < 0 Then
            stNHI(1) = .Fields("od03")
            stNHI(2) = .Fields("od02")
            stNHI(3) = "50"
            stNHI(4) = "4"
            stNHI(5) = ""
            stNHI(6) = ""
            If stNHI(7) = "" Then
               stNHI(7) = Abs(dblTFeeNet)
            End If
            stNHI(8) = ""
            stNHI(10) = "235910"
            stNHI(11) = .Fields("sm37")
            stNHI(11) = "2"
            stNHI(14) = .Fields("od02")
            PUB_NHI2nd stNHI(1), stNHI(2), stNHI(3), stNHI(4), stNHI(7), stNHI(5), stNHI(6), stNHI(8), stNHI(10), stNHI(11), stNHI(13)
            PUB_InsertNHI2nd stNHI
            
            '更新其他所得補充保費
            stSQL = "update OtherSalaryData set od13=" & Val(stNHI(6)) & " where od01=" & .Fields("od01")
            cnnConnection.Execute stSQL, intR
         End If
         .MoveNext
      Loop
      End With

      '不足額新增次月翻譯預支39
      If dblTFeeNet > 0 Then
         stSQL = "INSERT INTO othersalarydata(od01,od02,od03,od04,od05)" & _
               " select od01," & NextYM & "01,'" & stLstNo & "','39'," & dblTFeeNet & _
               " from (select max(OD01)+1 od01 from othersalarydata ) X"
         cnnConnection.Execute stSQL, intR
      '超出部分轉月薪資
      ElseIf dblTFeeNet < 0 Then
         stSQL = "insert into salarymonth(sm01,sm02,sm03,sm04,sm24,sm37,sm43)" & _
            " select sd01 sm01," & YM & " sm02," & stDepCol & " sm03," & Abs(dblTFeeNet) & " sm04,s02 sm24,sd19 sm37,s03 sm43 from (" & _
            " select od03 ,sum(od05) s01,sum(od06) s02,sum(od13) s03 from OtherSalaryData" & _
            " where substr(od02,1,6)=" & YM & " and od04='01' and od03='" & stLstFNo & "'" & _
            " group by od03) s,salarydata,staff s1" & _
            " where sd01(+)=od03 and st01(+)=od03"
         cnnConnection.Execute stSQL, intR
      End If
      
   End If
   
   List1.AddItem time & " --> 新增專職翻譯同仁翻譯所得資料結束,共 " & intR & " 筆", 0
   'end 2024/5/13
   
   List1.AddItem time & " --> 刪除總金額為 0 的資料開始", 0
   DoEvents
   'Modify By Sindy 2020/6/24 +nvl(sm45,0)
   stSQL = "delete from salarymonth where sm02=" & YM & _
      " and nvl(sm04,0)+nvl(sm05,0)+nvl(sm45,0)+nvl(sm06,0)+nvl(sm07,0)+nvl(sm08,0)+nvl(sm09,0)" & _
      "+nvl(sm10,0)+nvl(sm12,0)+nvl(sm13,0)+nvl(sm14,0)+nvl(sm15,0)+nvl(sm16,0)" & _
      "+nvl(sm17,0)+nvl(sm18,0)+nvl(sm19,0)+nvl(sm20,0)+nvl(sm21,0)+nvl(sm22,0)" & _
      "+nvl(sm23,0)+nvl(sm24,0)+nvl(sm43,0)+nvl(sm30,0)=0"
   
   cnnConnection.Execute stSQL, intR
   
   List1.AddItem time & " --> 刪除總金額為 0 的資料結束,共 " & intR & " 筆", 0
   
   cnnConnection.CommitTrans
   Process = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then MsgBox Err.Description
   
   'Added by Morgan 2022/11/16 若發生錯誤時需還原(因設定指定不會被Rollbak)
   strSql = "BEGIN user_data.user_num:='" & strUserNum & "'; END;"
   cnnConnection.Execute strSql, intI
End Function

'更新月薪欄位
Private Sub UpdMainItems(pUserNo As String, pYM As String, pSD02 As String)
   Dim stSQL As String, intR As Integer, adoRst As ADODB.Recordset
   Dim stWDs As String '當月工作天數
   Dim stDt1 As String '到(復)職日
   Dim stDt2 As String '離職日
   Dim stDt3 As String '異動日
   Dim iDay1 As Integer '舊薪資天數
   Dim iDay2 As Integer '新薪資天數
   Dim dblHour As Double '工作時數(非兼職為1)
   Dim stSM15 As String
   Dim iWDs As Integer '非整月分母天數
   
      
   'Modify by Morgan 2009/3/5 str() 函數有前置符號，長度為9 碼會導致日期計算有誤
   'stWDs = Right(CompDate(2, -1, CompDate(1, 1, str(100 * pYM + 1))), 2)
   stWDs = Right(CompDate(2, -1, CompDate(1, 1, Format(100 * pYM + 1))), 2)
   
   '抓兼職的工作時數
   If pSD02 = "P" Then
      stSQL = "select nvl(sum(PH03),0) from PTHour where PH02='" & pUserNo & "' and PH01=" & pYM
      intR = 1
      Set adoRst = ClsLawReadRstMsg(intR, stSQL, , True)
      If intR = 1 Then
         dblHour = adoRst.Fields(0)
      End If
   Else
      dblHour = 1
   End If
   
   
   '判斷該月份是否有異動
   'Modified by Morgan 2019/6/5 99037 108/5月有兩筆異動(1號調薪,7號調津貼)
   stSQL = "select SL02,nvl(sc02,st13) d1,st51 d2 from STAFF" & _
      ",(select SL01,max(SL02) SL02 from salarylog where SL01='" & pUserNo & "'" & _
      " and SUBSTR(SL02,1,6)=" & pYM & " group by sl01) a" & _
      ",(select sc01,sc02 from Staff_Change where sc01='" & pUserNo & "'" & _
      " and substr(sc02,1,6)=" & pYM & " and sc03='02') b" & _
      " where ST01='" & pUserNo & "' and SL01(+)=ST01 and sc01(+)=ST01" & _
      " order by SL02 asc"
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL, , True)
   If intR = 1 Then
      stDt1 = "" & adoRst.Fields("d1") '到(復)職日
      stDt2 = "" & adoRst.Fields("d2") '離職日
      stDt3 = "" & adoRst.Fields("SL02") '異動日
         
      '當月薪資改抓基本薪資資料，並限制計算薪資前不可變動
      '勞保費及勞退依照工作天比例,健保費固定整月,勞退自提費率抓基本薪資(異動沒放)
      
      '沒異動或當月到(復)職,抓基本薪資資料
      'Modify by Morgan 2009/6/22 異動日為當月1號的也抓基本薪資
      'If IsNull(adoRst("SL02")) Or Left(adoRst("d1"), 6) = pYM Then
      If stDt3 = "" Or Left(stDt1, 6) = pYM Or Val(Right(stDt3, 2)) = 1 Then
         
         'Added by Morgan 2023/10/11 非整月薪資分母不分大小月都用30，2月則維持實際天數
         If Val(stWDs) > 30 Then
            iWDs = 30
         Else
            iWDs = Val(stWDs)
         End If
         'end 2023/10/11
      
         'Modify by Morgan 2009/5/4 健保費到職扣,離職不扣
         'Modify by Morgan 2009/6/25
         '健保費改抓明細資料以批次方式更新
         '勞退投保薪資無特殊改抓SD43,SD44且新進同仁要抓全額級距才會對
         '勞保投保天數比照勞退方式
         'If Left(stDt2, 6) = pYM Then
         '   stSM15 = "0"
         'Else
         '   stSM15 = "SD15*(1+DECODE(SD11,'N',0,DECODE(SIGN(FMC-IR09),-1,FMC,IR09)))"
         'End If
         
         'stSQL = "update salarymonth set sm26=decode(round(sm27/" & stWDs & ",1),0,0.1,round(sm27/" & stWDs & ",1))" & _
            ",(sm04,sm05,sm06,sm07,sm08,sm09,sm10,sm14,sm15,sm25,sm38)" & _
            "=(select round(SD20*sm27/" & stWDs & "*" & dblHour & ") sm04" & _
            ",round(SD21*sm27/" & stWDs & ") sm05" & _
            ",round(SD22*sm27/" & stWDs & ") sm06" & _
            ",round(SD23*sm27/" & stWDs & ") sm07" & _
            ",round(SD24*sm27/" & stWDs & ") sm08" & _
            ",round(SD25*sm27/" & stWDs & ") sm09" & _
            ",round(SD26*sm27/" & stWDs & ") sm10" & _
            ",round(SD14*sm27/" & stWDs & ") sm14" & _
            "," & stSM15 & " sm15" & _
            ",round((NVL(SD20,0)+NVL(SD23,0))*" & dblHour & ") sm25" & _
            ",decode(SD16,'Y',round(DECODE(NVL(SD27,0),0,NVL(SD20,0)+NVL(SD21,0)+NVL(SD23,0),SD27) *sm27/" & stWDs & ")) sm38" & _
            " from salarydata,(select nvl(count(*),0) FMC" & _
            " from Staff_Relation where sr01='" & pUserNo & "' and SR08 is null) s,InsuranceRate" & _
            " where SD01=sm01 )" & _
            " where sm01='" & pUserNo & "' and sm02=" & pYM
         'Modify by Morgan 2011/8/4 時薪員工不必再乘工作天比例
         'Modified by Morgan 2013/3/29 合夥人投保者沒有自提
         'Modified by Morgan 2013/8/28 合夥人投保者又改可以自提--辜
         'Modified by Morgan 2015/6/8 特殊退休金投保薪資設 0 表示已退休 Ex.63001
         'Modified by Morgan 2020/1/9 +sm26改為未休假代金基準月薪
         'Modify By Sindy 2020/6/24 + ",round(SD52*sm27/" & stWDs & ") sm45" : 證照津貼
         '                          + SD52
         'stSQL = "update salarymonth set sm26=decode(round(sm27/" & stWDs & ",1),0,0.1,round(sm27/" & stWDs & ",1))"
         'Modified by Morgan 2023/10/11 非整月薪資分母不分大小月都用30，2月則維持實際天數
         stSQL = "update salarymonth set (sm04,sm05,sm06,sm45,sm07,sm08,sm09,sm10,sm14,sm25,sm26,sm38)" & _
            "=(select round(SD20" & IIf(pSD02 = "P", "", "*decode(sm27," & stWDs & ",1,sm27/" & iWDs & ")") & "*" & dblHour & ") sm04" & _
            ",round(SD21*decode(sm27," & stWDs & ",1,sm27/" & iWDs & ")) sm05" & _
            ",round(SD22*decode(sm27," & stWDs & ",1,sm27/" & iWDs & ")) sm06" & _
            ",round(SD52*decode(sm27," & stWDs & ",1,sm27/" & iWDs & ")) sm45" & _
            ",round(SD23*decode(sm27," & stWDs & ",1,sm27/" & iWDs & ")) sm07" & _
            ",round(SD24*decode(sm27," & stWDs & ",1,sm27/" & iWDs & ")) sm08" & _
            ",round(SD25*decode(sm27," & stWDs & ",1,sm27/" & iWDs & ")) sm09" & _
            ",round(SD26*decode(sm27," & stWDs & ",1,sm27/" & iWDs & ")) sm10" & _
            ",round(SD14*sm39/30) sm14" & _
            ",round((NVL(SD20,0)+NVL(SD23,0))*" & dblHour & ") sm25" & _
            ",round((NVL(SD20,0)+NVL(SD23,0)+NVL(SD21,0)+NVL(SD52,0))*" & dblHour & ") sm26" & _
            ",DECODE(SD16,'Y',round(NVL(SD27,NVL(SD43,0)))) sm38" & _
            " from salarydata where SD01=sm01 )" & _
            " where sm01='" & pUserNo & "' and sm02=" & pYM
         'end 2009/6/22
            
         cnnConnection.Execute stSQL, intR
               
         '新增第二家資料
         'Modified by Morgan 2020/1/9 +sm26 未休假代金基準月薪
         'Removed by Morgan 2023/10/11 目前已經沒有第二家
         'stSQL = "insert into salarymonth(sm01,sm02,sm03,sm04,sm05,sm06,sm07,sm08,sm09,sm10,sm25,sm26,sm27,sm37,sm38,sm39)" & _
            " select substr(sd01,1,2)||'A'||substr(sd01,4) sm01,sm02,sm03" & _
            ",round(SD29*sm27/" & stWDs & "*" & dblHour & ") sm04" & _
            ",round(SD30*sm27/" & stWDs & ") sm05" & _
            ",round(SD31*sm27/" & stWDs & ") sm06" & _
            ",round(SD32*sm27/" & stWDs & ") sm07" & _
            ",round(SD33*sm27/" & stWDs & ") sm08" & _
            ",round(SD34*sm27/" & stWDs & ") sm09" & _
            ",round(SD35*sm27/" & stWDs & ") sm10" & _
            ",round((NVL(SD29,0)+NVL(SD32,0))*" & dblHour & ") sm25" & _
            ",round((NVL(SD29,0)+NVL(SD32,0)+NVL(SD30,0))*" & dblHour & ") sm26" & _
            ",sm27,SD28 sm37,decode(SD16,'Y',round(NVL(SD36,NVL(SD44,0)))) sm38" & _
            ",sm39 From salarymonth,salarydata" & _
            " where sm01='" & pUserNo & "' and sm02=" & pYM & _
            " and SD01(+)=sm01 AND (nvl(SD29,0)+nvl(SD30,0)+nvl(SD31,0)+nvl(SD32,0)+nvl(SD33,0)+nvl(SD34,0)+nvl(SD35,0))>0"
            
         'cnnConnection.Execute stSQL, intR
         'end 2023/10/11
      
      Else
      
'Modify by Morgan 2009/6/22 健保費改抓明細資料以批次方式更新,1號異動併入上面,勞健退改抓薪資基本檔

'         'Modify by Morgan 2009/5/4 健保費到職扣,離職不扣
'         If Left(stDt2, 6) = pYM Then
'            stSM15 = "0"
'         Else
'            stSM15 = "SL10*(1+DECODE(SD11,'N',0,DECODE(SIGN(FMC-IR09),-1,FMC,IR09)))"
'         End If
'         '1號異動
'         '勞健保及勞退抓前次異動,其他欄位抓薪資基本資料
'         If Val(Right(adoRst.Fields(0), 2)) = 1 Then
'            stSQL = "update salarymonth set sm26=decode(round(sm27/" & stWDs & ",1),0,0.1,round(sm27/" & stWDs & ",1))" & _
'               ",(sm04,sm05,sm06,sm07,sm08,sm09,sm10,sm14,sm15,sm25,sm38)" & _
'               "=(select round(SD20*sm27/" & stWDs & "*" & dblHour & ") sm04" & _
'               ",round(SD21*sm27/" & stWDs & ") sm05" & _
'               ",round(SD22*sm27/" & stWDs & ") sm06" & _
'               ",round(SD23*sm27/" & stWDs & ") sm07" & _
'               ",round(SD24*sm27/" & stWDs & ") sm08" & _
'               ",round(SD25*sm27/" & stWDs & ") sm09" & _
'               ",round(SD26*sm27/" & stWDs & ") sm10" & _
'               ",round(SL09*sm27/" & stWDs & ") sm14" & _
'               "," & stSM15 & " sm15" & _
'               ",round((NVL(SD20,0)+NVL(SD23,0))*" & dblHour & ") sm25" & _
'               ",decode(SD16,'Y',round(DECODE(NVL(SL18,0),0,NVL(SL11,0)+NVL(SL12,0)+NVL(SL14,0),SL18) *sm27/" & stWDs & ")) sm38" & _
'               " from salarydata,salarylog b,(select NVL(count(*),0) FMC" & _
'               " from Staff_Relation where sr01='" & pUserNo & "' and SR08 is null) s,InsuranceRate" & _
'               " where sd01=sm01 and b.SL01(+)=sd01" & _
'               " and b.SL02=(select max(c.SL02) from salarylog c where c.SL01=sm01 and c.SL02<sm02*100))" & _
'               " where sm01='" & pUserNo & "' and sm02=" & pYM
'            cnnConnection.Execute stSQL, intR
'
'            '新增第二家資料
'            '基本薪資或前次異動的加項有金額的才做,兩家合併的第二家勞退投保薪資也要存
'            stSQL = "insert into salarymonth(sm01,sm02,sm03,sm04,sm05,sm06,sm07,sm08,sm09,sm10,sm25,sm26,sm27,sm37,sm38)" & _
'               " select substr(sd01,1,2)||'A'||substr(sd01,4) sm01,sm02,sm03" & _
'               ",round(SD29*sm27/" & stWDs & "*" & dblHour & ") sm04" & _
'               ",round(SD30*sm27/" & stWDs & ") sm05" & _
'               ",round(SD31*sm27/" & stWDs & ") sm06" & _
'               ",round(SD32*sm27/" & stWDs & ") sm07" & _
'               ",round(SD33*sm27/" & stWDs & ") sm08" & _
'               ",round(SD34*sm27/" & stWDs & ") sm09" & _
'               ",round(SD35*sm27/" & stWDs & ") sm10" & _
'               ",round((NVL(SD29,0)+NVL(SD32,0))*" & dblHour & ") sm25" & _
'               ",sm26,sm27,NVL(SL34,SD28) sm37" & _
'               ",decode(sd16,'Y',round(DECODE(NVL(SL26,0),0,NVL(SL19,0)+NVL(SL20,0)+NVL(SL22,0),SL26) *sm27/" & stWDs & ")) sm38" & _
'               " From salarymonth,salarydata,salarylog b" & _
'               " where sm01='" & pUserNo & "' and sm02=" & pYM & _
'               " and sd01(+)=sm01 and b.SL01(+)=sm01" & _
'               " and b.SL02=(select max(c.SL02) from salarylog c where c.SL01=sm01 and c.SL02<sm02*100)" & _
'               " AND (nvl(SD29,0)+nvl(SD30,0)+nvl(SD31,0)+nvl(SD32,0)+nvl(SD33,0)+nvl(SD34,0)+nvl(SD35,0)" & _
'               "+nvl(SL19,0)+nvl(SL20,0)+nvl(SL21,0)+nvl(SL22,0)+nvl(SL23,0)+nvl(SL24,0)+nvl(SL25,0))>0"
'
'            cnnConnection.Execute stSQL, intR
'         Else

            '1號以後
            '勞健保勞退抓前次異動,其他抓本次與前次的平均
            
            '異動當日用異動後薪資所以要減1
            iDay1 = Val(Right(stDt3, 2)) - 1
            '當月異動且離職
            If Left(stDt2, 6) = pYM Then
               iDay2 = Val(stDt2) - Val(stDt3)
               'Added by Morgan 2023/10/11 非整月薪資分母不分大小月都用30，2月則維持實際天數
               If Val(stWDs) > 30 Then
                  iWDs = 30
               Else
                  iWDs = Val(stWDs)
               End If
               'end 2023/10/11
            Else
               iDay2 = stWDs - iDay1
               iWDs = Val(stWDs) 'Added by Morgan 2023/10/11
            End If
         
            'stSQL = "update salarymonth set sm26=decode(round(sm27/" & stWDs & ",1),0,0.1,round(sm27/" & stWDs & ",1))" & _
               ",(sm04,sm05,sm06,sm07,sm08,sm09,sm10,sm14,sm15,sm25,sm38)=(select" & _
               " round((nvl(SD20,0)*" & iDay2 & "+nvl(SL11,0)*" & iDay1 & ")/" & stWDs & "*sm27/" & stWDs & "*" & dblHour & ") sm04" & _
               ",round((nvl(SD21,0)*" & iDay2 & "+nvl(SL12,0)*" & iDay1 & ")/" & stWDs & "*sm27/" & stWDs & ") sm05" & _
               ",round((nvl(SD22,0)*" & iDay2 & "+nvl(SL13,0)*" & iDay1 & ")/" & stWDs & "*sm27/" & stWDs & ") sm06" & _
               ",round((nvl(SD23,0)*" & iDay2 & "+nvl(SL14,0)*" & iDay1 & ")/" & stWDs & "*sm27/" & stWDs & ") sm07" & _
               ",round((nvl(SD24,0)*" & iDay2 & "+nvl(SL15,0)*" & iDay1 & ")/" & stWDs & "*sm27/" & stWDs & ") sm08" & _
               ",round((nvl(SD25,0)*" & iDay2 & "+nvl(SL16,0)*" & iDay1 & ")/" & stWDs & "*sm27/" & stWDs & ") sm09" & _
               ",round((nvl(SD26,0)*" & iDay2 & "+nvl(SL17,0)*" & iDay1 & ")/" & stWDs & "*sm27/" & stWDs & ") sm10" & _
               ",round(SL09*sm27/" & stWDs & ") sm14" & _
               "," & stSM15 & " sm15" & _
               ",round(((nvl(SD20,0)+nvl(SD23,0))*" & iDay2 & "+(nvl(SL11,0)+nvl(SL14,0))*" & iDay1 & ")/" & stWDs & "*" & dblHour & ") sm25" & _
               ",decode(SD16,'Y',round(DECODE(NVL(SL18,0),0,NVL(SL11,0)+NVL(SL12,0)+NVL(SL14,0),SL18) *sm27/" & stWDs & ")) sm38" & _
               " from salarydata,salarylog b,(select NVL(count(*),0) FMC" & _
               " from Staff_Relation where sr01='" & pUserNo & "' and SR08 is null) s,InsuranceRate" & _
               " where sd01=sm01 and b.SL01(+)=sd01" & _
               " and b.SL02=(select max(c.SL02) from salarylog c where c.SL01=sm01 and c.SL02<sm02*100))" & _
               " where sm01='" & pUserNo & "' and sm02=" & pYM
               
            'Modify by Morgan 2010/8/2 不必再乘天數比例否則當月調薪又離職的會少算
            'Modified by Morgan 2019/6/5 前次異動要考慮1號有調薪的狀況 99037 108/5月有兩筆異動(1號調薪,7號調津貼)
            'Modified by Morgan 2020/1/8 sm26改為未休假代金基準月薪
            'stSQL = "update salarymonth set sm26=decode(round(sm27/" & stWDs & ",1),0,0.1,round(sm27/" & stWDs & ",1))"
            'Modify By Sindy 2020/6/24 + ",round((nvl(SD52,0)*" & iDay2 & "+nvl(SL39,0)*" & iDay1 & ")/" & stWDs & ") sm45" : 證照津貼
            '                          + SD52
            'Modified by Morgan 2023/10/11 非整月薪資分母不分大小月都用30，2月則維持實際天數
            stSQL = "update salarymonth set (sm04,sm05,sm06,sm45,sm07,sm08,sm09,sm10,sm14,sm25,sm26,sm38)=(select" & _
               " round((nvl(SD20,0)*" & iDay2 & "+nvl(SL11,0)*" & iDay1 & ")/" & iWDs & "*" & dblHour & ") sm04" & _
               ",round((nvl(SD21,0)*" & iDay2 & "+nvl(SL12,0)*" & iDay1 & ")/" & iWDs & ") sm05" & _
               ",round((nvl(SD22,0)*" & iDay2 & "+nvl(SL13,0)*" & iDay1 & ")/" & iWDs & ") sm06" & _
               ",round((nvl(SD52,0)*" & iDay2 & "+nvl(SL39,0)*" & iDay1 & ")/" & iWDs & ") sm45" & _
               ",round((nvl(SD23,0)*" & iDay2 & "+nvl(SL14,0)*" & iDay1 & ")/" & iWDs & ") sm07" & _
               ",round((nvl(SD24,0)*" & iDay2 & "+nvl(SL15,0)*" & iDay1 & ")/" & iWDs & ") sm08" & _
               ",round((nvl(SD25,0)*" & iDay2 & "+nvl(SL16,0)*" & iDay1 & ")/" & iWDs & ") sm09" & _
               ",round((nvl(SD26,0)*" & iDay2 & "+nvl(SL17,0)*" & iDay1 & ")/" & iWDs & ") sm10" & _
               ",round(SD14*sm39/30) sm14" & _
               ",round(((nvl(SD20,0)+nvl(SD23,0))*" & iDay2 & "+(nvl(SL11,0)+nvl(SL14,0))*" & iDay1 & ")/" & iWDs & "*" & dblHour & ") sm25" & _
               ",round(((nvl(SD20,0)+nvl(SD23,0)+NVL(SD21,0)+NVL(SD52,0))*" & iDay2 & "+(nvl(SL11,0)+nvl(SL14,0)+nvl(SL12,0)+nvl(SL39,0))*" & iDay1 & ")/" & iWDs & "*" & dblHour & ") sm26" & _
               ",decode(SD16,'Y',round(DECODE(NVL(SD27,0),0,NVL(SD43,0),SD27))) sm38" & _
               " from salarydata,salarylog b" & _
               " where sd01=sm01 and b.SL01(+)=sd01" & _
               " and b.SL02=(select max(c.SL02) from salarylog c where c.SL01=sm01 and c.SL02<" & stDt3 & "))" & _
               " where sm01='" & pUserNo & "' and sm02=" & pYM
               
            cnnConnection.Execute stSQL, intR
                     
            '新增第二家資料
            'stSQL = "insert into salarymonth(sm01,sm02,sm03,sm04,sm05,sm06,sm07,sm08,sm09,sm10,sm25,sm26,sm27,sm37,sm38)" & _
               " select substr(sm01,1,2)||'A'||substr(sm01,4) sm01,sm02,sm03" & _
               ",round((nvl(SD29,0)*" & iDay2 & "+nvl(SL19,0)*" & iDay1 & ")/" & stWDs & "*sm27/" & stWDs & "*" & dblHour & ") sm04" & _
               ",round((nvl(SD30,0)*" & iDay2 & "+nvl(SL20,0)*" & iDay1 & ")/" & stWDs & "*sm27/" & stWDs & ") sm05" & _
               ",round((nvl(SD31,0)*" & iDay2 & "+nvl(SL21,0)*" & iDay1 & ")/" & stWDs & "*sm27/" & stWDs & ") sm06" & _
               ",round((nvl(SD32,0)*" & iDay2 & "+nvl(SL22,0)*" & iDay1 & ")/" & stWDs & "*sm27/" & stWDs & ") sm07" & _
               ",round((nvl(SD33,0)*" & iDay2 & "+nvl(SL23,0)*" & iDay1 & ")/" & stWDs & "*sm27/" & stWDs & ") sm08" & _
               ",round((nvl(SD34,0)*" & iDay2 & "+nvl(SL24,0)*" & iDay1 & ")/" & stWDs & "*sm27/" & stWDs & ") sm09" & _
               ",round((nvl(SD35,0)*" & iDay2 & "+nvl(SL25,0)*" & iDay1 & ")/" & stWDs & "*sm27/" & stWDs & ") sm10" & _
               ",round(((nvl(SD29,0)+nvl(SD32,0))*" & iDay2 & "+(nvl(SL19,0)+nvl(SL22,0))*" & iDay1 & ")/" & stWDs & "*" & dblHour & ") sm25" & _
               ",sm26,sm27,b.SL34 sm37" & _
               ",decode(sd16,'Y',round(DECODE(NVL(SL26,0),0,NVL(SL19,0)+NVL(SL20,0)+NVL(SL22,0),SL26) *sm27/" & stWDs & ")) sm38" & _
               " From salarymonth,salarydata,salarylog b" & _
               " where sm01='" & pUserNo & "' and sm02=" & pYM & _
               " and sd01(+)=sm01 and b.SL01(+)=sm01" & _
               " and b.SL02=(select max(c.SL02) from salarylog c where c.SL01=sm01 and c.SL02<sm02*100)" & _
               " AND (nvl(SD29,0)+nvl(SD30,0)+nvl(SD31,0)+nvl(SD32,0)+nvl(SD33,0)+nvl(SD34,0)+nvl(SD35,0)" & _
               "+nvl(SL19,0)+nvl(SL20,0)+nvl(SL21,0)+nvl(SL22,0)+nvl(SL23,0)+nvl(SL24,0)+nvl(SL25,0))>0"
               
            'Modify by Morgan 2010/8/2 不必再乘天數比例否則當月調薪又離職的會少算
            'Removed by Morgan 2023/10/11 目前已經沒有第二家
            'stSQL = "insert into salarymonth(sm01,sm02,sm03,sm04,sm05,sm06,sm07,sm08,sm09,sm10,sm25,sm26,sm27,sm37,sm38,sm39)" & _
               " select substr(sm01,1,2)||'A'||substr(sm01,4) sm01,sm02,sm03" & _
               ",round((nvl(SD29,0)*" & iDay2 & "+nvl(SL19,0)*" & iDay1 & ")/" & stWDs & "*" & dblHour & ") sm04" & _
               ",round((nvl(SD30,0)*" & iDay2 & "+nvl(SL20,0)*" & iDay1 & ")/" & stWDs & ") sm05" & _
               ",round((nvl(SD31,0)*" & iDay2 & "+nvl(SL21,0)*" & iDay1 & ")/" & stWDs & ") sm06" & _
               ",round((nvl(SD32,0)*" & iDay2 & "+nvl(SL22,0)*" & iDay1 & ")/" & stWDs & ") sm07" & _
               ",round((nvl(SD33,0)*" & iDay2 & "+nvl(SL23,0)*" & iDay1 & ")/" & stWDs & ") sm08" & _
               ",round((nvl(SD34,0)*" & iDay2 & "+nvl(SL24,0)*" & iDay1 & ")/" & stWDs & ") sm09" & _
               ",round((nvl(SD35,0)*" & iDay2 & "+nvl(SL25,0)*" & iDay1 & ")/" & stWDs & ") sm10" & _
               ",round(((nvl(SD29,0)+nvl(SD32,0))*" & iDay2 & "+(nvl(SL19,0)+nvl(SL22,0))*" & iDay1 & ")/" & stWDs & "*" & dblHour & ") sm25" & _
               ",round(((nvl(SD29,0)+nvl(SD32,0)+NVL(SD30,0))*" & iDay2 & "+(nvl(SL19,0)+nvl(SL22,0)+nvl(SL20,0))*" & iDay1 & ")/" & stWDs & "*" & dblHour & ") sm26" & _
               ",sm27,b.SL34 sm37,decode(SD16,'Y',round(DECODE(NVL(SD36,0),0,NVL(SD44,0),SD36))) sm38" & _
               ",sm39 From salarymonth,salarydata,salarylog b" & _
               " where sm01='" & pUserNo & "' and sm02=" & pYM & _
               " and sd01(+)=sm01 and b.SL01(+)=sm01" & _
               " and b.SL02=(select max(c.SL02) from salarylog c where c.SL01=sm01 and c.SL02<sm02*100)" & _
               " AND (nvl(SD29,0)+nvl(SD30,0)+nvl(SD31,0)+nvl(SD32,0)+nvl(SD33,0)+nvl(SD34,0)+nvl(SD35,0)" & _
               "+nvl(SL19,0)+nvl(SL20,0)+nvl(SL21,0)+nvl(SL22,0)+nvl(SL23,0)+nvl(SL24,0)+nvl(SL25,0))>0"
            
            'cnnConnection.Execute stSQL, intR
            '2023/10/11
'         End If
      End If
   End If
   
   Set adoRst = Nothing
End Sub
'勞退投保金額
Private Function GetInsureBase(pAmount As Long) As Long
   Dim stSQL As String, intR As Integer, adoRst As ADODB.Recordset
   stSQL = "select si02 from SalaryInsurance where si01='R' and si03<=" & pAmount & " and si04>=" & pAmount
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL, , False)
   If intR = 1 Then
      GetInsureBase = adoRst.Fields(0)
   End If
   Set adoRst = Nothing
End Function
'更新所得稅勞退欄位
Private Sub UpdTax(pUserNo As String, pYM As String)
   Dim stSQL As String, intR As Integer, adoRst As ADODB.Recordset
   Dim lngIncome As Long, dblRate As Double, dblTax As Double, lngTaxIncome As Long
   Dim lngSM38 As Long '勞退投保薪資
   Dim lngIBase As Long '勞退投保計算薪資
   Dim dblStaffRate As Double '自提費率
   Dim dblCompRate As Double '公司提撥費率
   Dim intFMCMax As Integer '最高眷口數
   Dim intFMC As Integer '眷口數
   Dim lngSM16 As Long '勞退自提金額
   Dim lngSM30 As Long '勞退公司提撥
   Dim dblWDRate As Double '工作天比例

   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'stSQL = "select * from salarymonth,salarydata,IncomeTax" & _
      ",(select IR05,IR09 from InsuranceRate) y" & _
      " where sm01='" & pUserNo & "' and sm02=" & pYM & " and sd01(+)=replace(sm01,'A','0')"
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   stSQL = "select * from salarymonth,salarydata,IncomeTax" & _
      ",(select IR05,IR09 from InsuranceRate) y" & _
      " where sm01='" & pUserNo & "' and sm02=" & pYM & " and sd01(+)=substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)"
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL, , False)
   If intR = 1 Then
      With adoRst
      lngSM38 = Val("" & .Fields("sm38"))
      dblStaffRate = Val("" & .Fields("sd17"))
      dblCompRate = Val("" & .Fields("IR05"))
      intFMC = Val("" & .Fields("sd07"))
      'Remove by Morgan 2009/6/15 不必控制眷口數(健保才有)
      'intFMCMax = Val("" & .Fields("IR09"))
      'If intFMC > intFMCMax Then
      '   intFMC = intFMCMax
      'End If
      'end 2009/6/15
      If lngSM38 > 0 Then
         lngIBase = GetInsureBase(lngSM38)
         'Modify by Morgan 2009/6/24 工作天不滿整月的改在這裡處理(先算會造成級數抓錯),固定除以30天
         'lngSM16 = Round(dblStaffRate * lngIBase / 100)
         'lngSM30 = Round(dblCompRate * lngIBase / 100)
         If Val("" & .Fields("sm27")) = WDs Then
            dblWDRate = 1
         Else
            dblWDRate = Val("" & .Fields("sm39")) / 30
         End If
         lngSM16 = Round(dblStaffRate * lngIBase / 100 * dblWDRate)
         lngSM30 = Round(dblCompRate * lngIBase / 100 * dblWDRate)
         'Added by Morgan 2013/8/28 合夥人公司沒有提撥(原控制投保薪資為0但,8月起合夥人也可自提所以改此處控制)
         If .Fields("SD11") = "Y" Then
            lngSM30 = 0
         End If
         'end 2013/8/28
         'end 2009/6/24
      End If
      
      '薪資所得=基本薪資+職務津貼+超時加班費-缺勤扣款
      'Modify By Sindy 2020/6/24 '薪資所得=基本薪資+職務津貼+證照津貼+超時加班費-缺勤扣款
      'lngIncome = Val("" & .Fields("sm04")) + Val("" & .Fields("sm05")) + Val("" & .Fields("sm28")) - Val("" & .Fields("sm21"))
      lngIncome = Val("" & .Fields("sm04")) + Val("" & .Fields("sm05")) + Val("" & .Fields("sm45")) + _
                  Val("" & .Fields("sm28")) - Val("" & .Fields("sm21"))
      '基本檔有設定所得稅率
      dblRate = Val("" & .Fields("sd08"))
      If dblRate > 0 Then
         dblTax = dblRate * lngIncome / 100 + Val("" & .Fields("sm29"))
      Else
         '計算薪資=(薪資所得x12粗估年度所得－個人免稅額IT01－扶養人數x扶養親屬寬減額IT02－夫妻標準扣除額IT03或單身標準扣除額IT04－薪資扣除額IT05）；
         '再以計算薪資抓所得稅率表的級距帶出稅率及累進差額，但若有輸入所得稅率者依輸入之稅率計算；
         lngTaxIncome = 12 * lngIncome - Val("" & .Fields("it01")) - intFMC * Val("" & .Fields("it02")) - Val("" & .Fields("it05"))
         '已婚
         If "" & .Fields("sd03") = "Y" Then
            lngTaxIncome = lngTaxIncome - Val("" & .Fields("it03"))
         '未婚
         Else
            lngTaxIncome = lngTaxIncome - Val("" & .Fields("it04"))
         End If
         '計算薪資＊稅率－累進差額＝整年所得稅稅額；
         'Modified by Morgan 2024/12/5 應該是要抓下一階的薪資判斷
         'For intR = 1 To 15
         '   If Val("" & .Fields("it" & Format(3 + 3 * intR, "00"))) >= lngTaxIncome Then
         For intR = 1 To 14
            If Val("" & .Fields("it" & Format(3 + 3 * (intR + 1), "00"))) > lngTaxIncome Then
         'end 2024/12/5
               dblRate = Val("" & .Fields("it" & Format(4 + 3 * intR, "00")))
               dblTax = (dblRate * lngTaxIncome / 100) - Val("" & .Fields("it" & Format(5 + 3 * intR, "00")))
               Exit For
            End If
         Next
         '整年所得稅稅額/12=每月所得稅。
         dblTax = dblTax / 12
         '若所得稅(含其他所得稅金)低於2000元者不扣
         dblTax = dblTax + Val("" & .Fields("sm29"))
         If dblTax < 2000 Then dblTax = 0
      End If
      '四捨五入
      'Modified by Morgan 2015/12/10 +SM44
      stSQL = "update salarymonth set sm24=round(" & dblTax & "),sm16=" & lngSM16 & ",sm30=" & lngSM30 & _
         ",sm44=" & dblStaffRate & " where sm01='" & pUserNo & "' and sm02=" & pYM
      cnnConnection.Execute stSQL, intR
      End With
   End If
      
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Function TxtValidate(Optional bolCancel As Boolean = False) As Boolean
   Dim bCancel As Boolean
   
   If Text1 = "" Then
      MsgBox "年度不可空白 !"
      Text1.SetFocus
      Exit Function
   End If
   If Text2 = "" Then
      MsgBox "月份不可空白 !"
      Text2.SetFocus
      Exit Function
   End If
   
   Text2_Validate bCancel
   If bCancel = True Then
      Text2.SetFocus
      Text2_GotFocus
      Exit Function
   End If
   
   'Add by Morgan 2009/6/15
   strSql = "select * from BookRecord where BR01=" & (100 * Val(Text1) + Val(Text2) + 191100)
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If bolCancel Then
         MsgBox "該月份已有薪資入帳紀錄，不可取消計算！", vbExclamation
      Else
         MsgBox "該月份已有薪資入帳紀錄，不可重新計算！", vbExclamation
      End If
      Exit Function
   End If
   'end 2009/6/15
   
   If bolCancel Then
      strSql = "select * from salarymonth where sm02=" & (100 * Val(Text1) + Val(Text2) + 191100) & " and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI <> 1 Then
         MsgBox "尚未計算，不必取消！", vbInformation
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function
   
Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 <> "" Then
      If Val(Text2) < 1 Or Val(Text2) > 12 Then
         MsgBox "月份輸入錯誤 !"
         Text2_GotFocus
         Cancel = True
      End If
   End If
End Sub

'Add by Morgan 2009/6/22
'設定超過最高眷口數的健保費明細(除保費最低的三個外其他設為超額眷口)
Private Sub UpdHiMonth(pUserNo As String, pYM As String)
   Dim stSQL As String, intR As Integer, adoRst As ADODB.Recordset
   
   'Modify by Morgan 2011/4/19 +第二排序改出生日大的
   stSQL = "select HM02 from HiMonth,Staff_Relation where HM01='" & pUserNo & "' and HM02<>0 and HM03=" & pYM & _
      " and sr01(+)=hm01 and sr02(+)=hm02 order by HM04 asc,sr06 desc,HM02"
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      With adoRst
      Do While Not .EOF
         If .AbsolutePosition > 3 Then
            stSQL = "update HiMonth set HM04=0,HM05='Y' where HM01='" & pUserNo & "' and HM02=" & .Fields(0) & " and HM03=" & pYM
            cnnConnection.Execute stSQL, intR
         End If
         .MoveNext
      Loop
      End With
   End If
   
End Sub

'Added by Morgan 2022/11/18
'技術/差旅/房租津貼轉入獎金系統
Private Sub BatchMonthBonus()
   Dim stYr As String, stMn As String, stYrW As String, stMB13 As String
   Dim stMPeriod As String, stSubject As String, stTO As String
   Dim stSQL As String, intR As Integer, intRecs As Integer, lngTotal As Long

   List1.AddItem time & " --> 差旅/技術/房租津貼轉入每月獎金...", 0
   DoEvents
   
   'Added by Morgan 2022/11/18 補資料用
   'YM = 100 * Val(Text1) + Val(Text2) + 191100
   'strSql = "delete MonthBonus where mb01>=" & YM & "01 and mb05='QPGMR' and MB13>0"
   'cnnConnection.Execute strSql, intI
   'end 2022/11/18
   
   stYr = Val(Text1)
   stMn = Val(Text2)
   stYrW = Val(Text1) + 1911
   'Modified by Morgan 2024/1/5 修正代扣日期都在1231問題
   If Val(stMn) <= 4 Then
      stMB13 = stYrW & "0430"
   ElseIf Val(stMn) <= 8 Then
      stMB13 = stYrW & "0831"
   Else
      stMB13 = stYrW & "1231"
   End If
   
   stSQL = "select sm01,sm37,nvl(sm06,0) S1,nvl(sm08,0) S2,nvl(sm09,0) S3 from salarymonth " & _
      " where sm02=" & YM & " and (sm06>0 or sm08>0 or sm09>0)" & _
      " order by sm01"
   intR = 1
   Set RsTemp = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      stSQL = "BEGIN user_data.user_num:='QPGMR'; END;"
      cnnConnection.Execute stSQL, intR
      With RsTemp
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      intRecs = 0
      lngTotal = 0
      Do While Not .EOF
         '技術津貼
         If .Fields("S1") > 0 Then
            List1.AddItem time & " --> 新增 " & .Fields("sm01") & " 技術津貼...", 0
            '獎金日期抓24日(含)以前的最大工作日
            stSQL = "INSERT INTO MonthBonus(MB01,MB02,MB03,MB11,MB13,MB14)" & _
               " select max(wd01) MB01,'" & .Fields("sm01") & "' MB02," & .Fields("S1") & " MB03" & _
               ",'" & .Fields("sm37") & "' MB11," & stMB13 & " MB13,'" & stYr & "年" & stMn & "月技術津貼' MB14" & _
               " from (select (" & YM & "00+rownum) wd01 from workday where rownum<=24) X" & _
               " where not exists(select * from MonthBonus where mb02='" & .Fields("sm01") & "' and mb01=wd01 )"
            cnnConnection.Execute stSQL, intR
            intRecs = intRecs + 1
            lngTotal = lngTotal + .Fields("S1")
         End If
         '差旅津貼
         If .Fields("S2") > 0 Then
            List1.AddItem time & " --> 新增 " & .Fields("sm01") & " 差旅津貼...", 0
            '獎金日期抓25日(含)以前的最大工作日
            stSQL = "INSERT INTO MonthBonus(MB01,MB02,MB03,MB11,MB13,MB14)" & _
               " select max(wd01) MB01,'" & .Fields("sm01") & "' MB02," & .Fields("S2") & " MB03" & _
               ",'" & .Fields("sm37") & "' MB11," & stMB13 & " MB13,'" & stYr & "年" & stMn & "月差旅津貼' MB14" & _
               " from (select (" & YM & "00+rownum) wd01 from workday where rownum<=25) X" & _
               " where not exists(select * from MonthBonus where mb02='" & .Fields("sm01") & "' and mb01=wd01 )"
            cnnConnection.Execute stSQL, intR
            intRecs = intRecs + 1
            lngTotal = lngTotal + .Fields("S2")
         End If
         '房租津貼
         If .Fields("S3") > 0 Then
            List1.AddItem time & " --> 新增 " & .Fields("sm01") & " 房租津貼...", 0
            stSQL = "INSERT INTO MonthBonus(MB01,MB02,MB03,MB11,MB13,MB14)" & _
               " select max(wd01) MB01,'" & .Fields("sm01") & "' MB02," & .Fields("S3") & " MB03" & _
               ",'" & .Fields("sm37") & "' MB11," & stMB13 & " MB13,'" & stYr & "年" & stMn & "月房租津貼' MB14" & _
               " from (select (" & YM & "00+rownum) wd01 from workday where rownum<=26) X" & _
               " where not exists(select * from MonthBonus where mb02='" & .Fields("sm01") & "' and mb01=wd01 )"
            cnnConnection.Execute stSQL, intR
            intRecs = intRecs + 1
            lngTotal = lngTotal + .Fields("S3")
         End If
         ProgressBar1.Value = ProgressBar1.Value + 1
         Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         .MoveNext
      Loop
      End With
      stSQL = "BEGIN user_data.user_num:='" & strUserNum & "'; END;"
      cnnConnection.Execute stSQL, intR
      
      stSubject = stYr & "年" & stMn & "月之技術/差旅/房租津貼已轉入獎金系統(合計" & Format(lngTotal, "#,###") & "元)，請紀錄各類所得！"
      stTO = Pub_GetSpecMan("試用期滿追蹤薪資人員")
      stSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
         " values('" & strUserNum & "','" & stTO & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
         ",'" & stSubject & "','如旨')"
      cnnConnection.Execute stSQL, intR
   End If
   
   List1.AddItem time & " --> 差旅/技術/房租津貼轉入每月獎金結束,共 " & intRecs & " 筆", 0
End Sub

