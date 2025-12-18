VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm170302 
   BorderStyle     =   1  '單線固定
   Caption         =   "計算年終獎金"
   ClientHeight    =   4056
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   5160
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4056
   ScaleWidth      =   5160
   Begin VB.ListBox List1 
      Height          =   2208
      Left            =   60
      TabIndex        =   4
      Top             =   1680
      Width           =   4920
   End
   Begin VB.TextBox txtYEAR 
      Height          =   270
      Left            =   1380
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "96"
      Top             =   675
      Width           =   400
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "計算(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   2820
      TabIndex        =   1
      Top             =   60
      Width           =   1065
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   405
      Index           =   1
      Left            =   3990
      TabIndex        =   2
      Top             =   60
      Width           =   1065
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   5010
      _ExtentX        =   8827
      _ExtentY        =   466
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "** 試算 **"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   432
      TabIndex        =   7
      Top             =   96
      Visible         =   0   'False
      Width           =   1584
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "( 0/0 )"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   1380
      Width           =   4920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "計算年度：             年"
      Height          =   180
      Left            =   420
      TabIndex        =   3
      Top             =   720
      Width           =   1665
   End
End
Attribute VB_Name = "frm170302"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/1/26 Form2.0已檢查 (無需修改的物件)
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'2009/1/2 add by sonia
Option Explicit

Public m_bolIsTrial As Boolean 'Added by Morgan 2023/12/11 是否為試算

Dim NHI() As String


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0 '計算
         If TxtValidate() = True Then
            If Progress() = True Then
            Else
               MsgBox "無符合條件資料可計算！", vbInformation
            End If
         End If
         txtYEAR.SetFocus
      Case 1 '結束
         Unload Me
   End Select
End Sub

Private Function Progress() As Boolean
Dim strSql As String
Dim stSQL As String
Dim adoRst As ADODB.Recordset, intR As Integer
Dim m_YearDay As Long       '年度總天數
Dim m_HourPay As Double     '時薪
Dim m_HourPayVC As Double   '未休假代金時薪  'add by sonia 2020/1/13
'Dim douHour(18) As Double   '出缺勤陣列
'Dim douCnt(18) As Double    '出缺勤人數陣列
'Modify By Sindy 2012/1/4
Dim douHour(25) As Double   '出缺勤陣列
Dim douCnt(25) As Double    '出缺勤人數陣列
Dim m_yb11 As Double        '曠職時數
Dim m_SickHour As Double    '病假不扣年終時數
Dim m_AbsenceHour As Double '事假不扣年終時數
Dim m_BornHour As Double    '產假不扣年終時數
Dim m_NoBornHour As Double  '流產假不扣年終時數
Dim m_hurtHour As Double  '公傷假不扣年終時數   'add by sonia 2016/1/6
Dim m_SubHour As Double     '年終缺勤扣款總時數
Dim m_TaxTotal As Long      '二家所得稅總額
Dim m_yb05 As Long
Dim m_yb06 As Long
Dim m_yb26 As Long        'add by sonia 2018/1/30
Dim m_taxrate As String   '2010/12/30 add by sonia 非固定之薪資所得扣繳稅率
Dim m_yvhour As Double    '2013/1/22 ad by sonia 可休假時數
Dim stWDay1 As String  'Added by Morgan 2023/12/15 第一個工作日

   Progress = False
   
   '取得計算年度之總天數
   If PUB_GetMonthDays((Val(txtYEAR) + 1911), 2) = 28 Then
      m_YearDay = 365
   Else
      m_YearDay = 366
   End If
   
   stWDay1 = PUB_GetWorkDay1((Val(txtYEAR) + 1911) & "0101", False) 'Added by Morgan 2023/12/15 第一個工作日
   
   '2010/12/30 add by sonia 非固定之薪資所得扣繳稅率改抓 翻譯所得oc01='01'的稅率
   m_taxrate = 0
   strExc(0) = "select oc04 from OtherSalaryCode where oc01='01'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      m_taxrate = "" & RsTemp.Fields(0)
   End If
   '2010/12/30 end
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   List1.Clear
   
   List1.AddItem time & " --> 刪除年終獎金資料開始...", 0
   DoEvents
   stSQL = "delete yearbonus where yb01=" & Val(txtYEAR) + 1911
   cnnConnection.Execute stSQL, intR
   
   If m_bolIsTrial = False Then 'Added by Morgan 2023/12/12 非試算才執行
      '2013/1/17 ADD BY SONIA
      List1.AddItem time & " --> 刪除年終獎金補充保費資料開始...", 0
      DoEvents
      stSQL = "delete NHI2ND where SUBSTR(NHI02,1,4)=" & Val(txtYEAR) + 1912 & " AND NHI04='1'"
      cnnConnection.Execute stSQL, intR
      '2013/1/17 END
   End If
   
   List1.AddItem time & " --> 刪除年終獎金資料結束,共 " & intR & " 筆", 0
   
   List1.AddItem time & " --> 新增第一家年終獎金資料開始...", 0
   DoEvents
   '當時在職6~9員工編號且該年度二家年終獎金基準月薪合計>0者寫入年終獎金資料
   '2010/1/18 modify by sonia 因二家合併至一家的當月第二家有sm25=0的資料,故取消m2.sm25的條件,否則第一家當月的工作天抓不到
   'stSQL = "insert into yearbonus (yb24,yb01,yb02,yb03,yb04,yb05,yb06) " & _
           "(select sd19 一公司別,substr(m1.sm02,1,4) 薪資年度,m1.sm01 員工代號,st03 部門," & _
           "round(sum(nvl(m1.sm25,0)+nvl(m2.sm25,0))/count(*),0) 平均基準月薪," & _
           "round(round(sum(nvl(m1.sm25,0)+nvl(m2.sm25,0))/count(*),0)*ybm03*decode(nvl(ym02,'2'),'1',110,'2',100,'3',85,'4',60)/100*sum(nvl(m1.sm27,0))/" & m_YearDay & ",0) 年終獎金," & _
           "sd18 特殊功績獎金 from salarymonth m1,salarymonth m2,salarydata,staff,yearbonusmonth,yearmerit " & _
           "where st04='1' and st01>='6' and st01<'A' and st01=m1.sm01 and substr(m1.sm02,1,4)='" & Val(txtYEAR) + 1911 & "' and m1.sm25>0 " & _
           "and st01=sd01(+) and '" & Val(txtYEAR) + 1911 & "'=ybm01(+) and decode(sd19,'0','1','2','1','9','1',sd19)=ybm02(+) " & _
           "and '" & Val(txtYEAR) + 1911 & "'=ym01(+) and sd01=ym03(+) " & _
           "and substr(m1.sm01,1,2)||'A'||substr(m1.sm01,4,2)=m2.sm01(+) and m1.sm02=m2.sm02(+) and (m2.sm25>0 or m2.sm25 is null) " & _
           "group by sd19,substr(m1.sm02,1,4),m1.sm01,st03,ybm03,decode(nvl(ym02,'2'),'1',110,'2',100,'3',85,'4',60),sd18)"
   'modify by sonia 2016/1/6 1.留職停薪人員也要計算 2.年終考績不得參加人員ym02='*'者以100計算
   'stSQL = "insert into yearbonus (yb24,yb01,yb02,yb03,yb04,yb05,yb06) " & _
           "(select sd19 一公司別,substr(m1.sm02,1,4) 薪資年度,m1.sm01 員工代號,st03 部門," & _
           "round(sum(nvl(m1.sm25,0)+nvl(m2.sm25,0))/count(*),0) 平均基準月薪," & _
           "round(round(sum(nvl(m1.sm25,0)+nvl(m2.sm25,0))/count(*),0)*ybm03*decode(nvl(ym02,'2'),'1',110,'2',100,'3',85,'4',60)/100*sum(nvl(m1.sm27,0))/" & m_YearDay & ",0) 年終獎金," & _
           "sd18 特殊功績獎金 from salarymonth m1,salarymonth m2,salarydata,staff,yearbonusmonth,yearmerit " & _
           "where st04='1' and st01>='6' and st01<'F' and st01=m1.sm01 and substr(m1.sm02,1,4)='" & Val(txtYEAR) + 1911 & "' and m1.sm25>0 " & _
           "and st01=sd01(+) and " & Val(txtYEAR) + 1911 & "=ybm01(+) and decode(sd19,'0','1','2','1','9','1',sd19)=ybm02(+) " & _
           "and " & Val(txtYEAR) + 1911 & "=ym01(+) and sd01=ym03(+) " & _
           "and substr(m1.sm01,1,2)||'A'||substr(m1.sm01,4,2)=m2.sm01(+) and m1.sm02=m2.sm02(+) " & _
           "group by sd19,substr(m1.sm02,1,4),m1.sm01,st03,ybm03,decode(nvl(ym02,'2'),'1',110,'2',100,'3',85,'4',60),sd18)"
   'modify by sonia 2018/1/10 +yb26(sd51)
   'modify by sonia 2019/1/19 改為未輸基準月數的公司都抓1公司的基準月數
   'stSQL = "insert into yearbonus (yb24,yb01,yb02,yb03,yb04,yb05,yb06,yb26) " & _
           "(select sd19 一公司別,substr(m1.sm02,1,4) 薪資年度,m1.sm01 員工代號,st03 部門," & _
           "round(sum(nvl(m1.sm25,0)+nvl(m2.sm25,0))/count(*),0) 平均基準月薪," & _
           "round(round(sum(nvl(m1.sm25,0)+nvl(m2.sm25,0))/count(*),0)*ybm03*decode(nvl(ym02,'2'),'1',110,'2',100,'3',85,'4',60,'*',100)/100*sum(nvl(m1.sm27,0))/" & m_YearDay & ",0) 年終獎金," & _
           "sd18 特殊功績獎金,sd51 紅利 from salarymonth m1,salarymonth m2,salarydata,staff,yearbonusmonth,yearmerit " & _
           "where (st04='1' or sd02='S') and st01>='6' and st01<'F' and st01=m1.sm01 and substr(m1.sm02,1,4)='" & Val(txtYEAR) + 1911 & "' and m1.sm25>0 " & _
           "and st01=sd01(+) and " & Val(txtYEAR) + 1911 & "=ybm01(+) and decode(sd19,'0','1','2','1','9','1',sd19)=ybm02(+) " & _
           "and " & Val(txtYEAR) + 1911 & "=ym01(+) and sd01=ym03(+) " & _
           "and substr(m1.sm01,1,2)||'A'||substr(m1.sm01,4,2)=m2.sm01(+) and m1.sm02=m2.sm02(+) " & _
           "group by sd19,substr(m1.sm02,1,4),m1.sm01,st03,ybm03,decode(nvl(ym02,'2'),'1',110,'2',100,'3',85,'4',60,'*',100),sd18,sd51)"
   'modify by sonia 2022/1/5 改為未輸基準月數的公司都抓2公司的基準月數
   'Modified by Morgan 2023/12/15 第1個工作天到職/復職也要算全年在職(100%),
   'stSQL = "insert into yearbonus (yb24,yb01,yb02,yb03,yb04,yb05,yb06,yb26) " & _
           "(select sd19 一公司別,substr(m1.sm02,1,4) 薪資年度,m1.sm01 員工代號,st03 部門," & _
           "round(sum(nvl(m1.sm25,0)+nvl(m2.sm25,0))/count(*),0) 平均基準月薪," & _
           "round(round(sum(nvl(m1.sm25,0)+nvl(m2.sm25,0))/count(*),0)*nvl(y1.ybm03,y2.ybm03)*decode(nvl(ym02,'2'),'1',110,'2',100,'3',85,'4',60,'*',100)/100*sum(nvl(m1.sm27,0))/" & m_YearDay & ",0) 年終獎金," & _
           "sd18 特殊功績獎金,sd51 紅利 from salarymonth m1,salarymonth m2,salarydata,staff,yearbonusmonth y1,yearbonusmonth y2,yearmerit " & _
           "where (st04='1' or sd02='S') and st01>='6' and st01<'F' and st01=m1.sm01 and substr(m1.sm02,1,4)='" & Val(txtYEAR) + 1911 & "' and m1.sm25>0 " & _
           "and st01=sd01(+) and " & Val(txtYEAR) + 1911 & "=y1.ybm01(+) and sd19=y1.ybm02(+) and " & Val(txtYEAR) + 1911 & "=y2.ybm01(+) and '2'=y2.ybm02(+) " & _
           "and " & Val(txtYEAR) + 1911 & "=ym01(+) and sd01=ym03(+) " & _
           "and substr(m1.sm01,1,2)||'A'||substr(m1.sm01,4,2)=m2.sm01(+) and m1.sm02=m2.sm02(+) " & _
           "group by sd19,substr(m1.sm02,1,4),m1.sm01,st03,nvl(y1.ybm03,y2.ybm03),decode(nvl(ym02,'2'),'1',110,'2',100,'3',85,'4',60,'*',100),sd18,sd51)"
   
   'Modified by Morgan 2023/12/20 新部門啟用日的前年度要開始抓新部門st93，因為發放(扣繳)是隔年--秀玲
   'Modified by Morgan 2025/1/22 排除第4碼為9的
   stSQL = "insert into yearbonus (yb24,yb01,yb02,yb03,yb04,yb05,yb06,yb26) " & _
           "select sd19 一公司別,substr(sm02,1,4) 薪資年度,sm01 員工代號,decode(sign(substr(sm02,1,4)-" & Left(新部門啟用日, 4) & "+1),-1,st03,st93) 部門,sm25 平均基準月薪," & _
           "round(sm25*ybm03*(decode(nvl(ym02,'2'),'1',110,'2',100,'3',85,'4',60,'*',100)/100)" & _
           "*nvl(decode(nvl(sc02,st13)," & stWDay1 & ",decode(sm27+" & (Right(stWDay1, 2) - 1) & "," & m_YearDay & ",1)),sm27/" & m_YearDay & "),0) 年終獎金," & _
           "sd18 特殊功績獎金,sd51 紅利 from staff,salarydata,yearbonusmonth,yearmerit " & _
           ",(select sm01,min(sm02) sm02,round(sum(sm25)/count(*),0) sm25,sum(sm27) sm27 from salarymonth where substr(sm02,1,4)='" & (Val(txtYEAR) + 1911) & "' group by sm01) m" & _
           ",(select sc01,min(sc02) sc02 from staff_change where substr(sc02,1,4)='" & (Val(txtYEAR) + 1911) & "' and sc03 in ('01','02') group by sc01 ) s" & _
           " where st01>='6' and st01<'F' and substr(st01,4,1)<>'9' and sd01(+)=st01 and (st04='1' or sd02='S') and ybm01(+)=" & (Val(txtYEAR) + 1911) & " and ybm02(+)=sd19" & _
           " and ym01(+)=" & (Val(txtYEAR) + 1911) & " and ym03(+)=sd01 and sm01(+)=sd01 and sm25>0 and sc01(+)=sd01"
   'end 2023/12/15
   cnnConnection.Execute stSQL, intR
   
   List1.AddItem time & " --> 新增第一家年終獎金資料結束,共 " & intR & " 筆", 0
   
   'add by sonia 2018/2/9 計算年度下半年屆齡65歲強制退休人員也要發年終獎金,但未休假代金已於退休時發放,此處不再算
   'modify by sonia 2021/2/22 改全年度
   'List1.AddItem time & " --> 新增下半年屆齡強制退休資料開始...", 0
   List1.AddItem time & " --> 新增當年屆齡強制退休資料開始...", 0
   DoEvents
   'modify by sonia 2019/1/19 改為未輸基準月數的公司都抓1公司的基準月數
   'stSQL = "insert into yearbonus (yb24,yb01,yb02,yb03,yb04,yb05,yb06,yb26) " & _
           "(select sd19 一公司別,substr(m1.sm02,1,4) 薪資年度,m1.sm01 員工代號,st03 部門," & _
           "round(sum(nvl(m1.sm25,0)+nvl(m2.sm25,0))/count(*),0) 平均基準月薪," & _
           "round(round(sum(nvl(m1.sm25,0)+nvl(m2.sm25,0))/count(*),0)*ybm03*decode(nvl(ym02,'2'),'1',110,'2',100,'3',85,'4',60,'*',100)/100*sum(nvl(m1.sm27,0))/" & m_YearDay & ",0) 年終獎金," & _
           "sd18 特殊功績獎金,sd51 紅利 from salarymonth m1,salarymonth m2,salarydata,staff,yearbonusmonth,yearmerit,staff_change " & _
           "where sc03='08' and sc02>=" & Val(txtYEAR) + 1911 & "0701 and sc02<=" & Val(txtYEAR) + 1911 & "1231 and sc01=st01(+) and st04='2' and substr(sc02,1,4)-substr(st23,1,4)=65 " & _
           "and st01=m1.sm01 and substr(m1.sm02,1,4)='" & Val(txtYEAR) + 1911 & "' and m1.sm25>0 " & _
           "and st01=sd01(+) and " & Val(txtYEAR) + 1911 & "=ybm01(+) and decode(sd19,'0','1','2','1','9','1',sd19)=ybm02(+) " & _
           "and " & Val(txtYEAR) + 1911 & "=ym01(+) and sd01=ym03(+) " & _
           "and substr(m1.sm01,1,2)||'A'||substr(m1.sm01,4,2)=m2.sm01(+) and m1.sm02=m2.sm02(+) " & _
           "group by sd19,substr(m1.sm02,1,4),m1.sm01,st03,ybm03,decode(nvl(ym02,'2'),'1',110,'2',100,'3',85,'4',60,'*',100),sd18,sd51)"
   'modify by sonia 2021/2/22 改為當年全年度屆齡強制退休人員0701改為0101
   'modify by sonia 2022/1/20 改為未輸基準月數的公司都抓2公司的基準月數
   'modify by sonia 2023/1/10 因為69009楊監察人在112/1/2退休但111年年終仍要發放,所以取消sc02<=" & Val(txtYEAR) + 1911 & "1231...的條件
   
   'Modified by Morgan 2023/12/20 新部門啟用日的前年度要開始抓新部門st93，因為發放(扣繳)是隔年--秀玲
   'Modified by Morgan 2023/12/25 退休年齡改>=65(76012 桂齊恆,70004 吳婧瑄 延後退休)
   'Modified by Morgan 2024/2/21 屆齡退休定義修改為退休日>=生日+65年(第66歲生日起) --秀玲,婉莘 (Ex:71011王錦寬提前退休不應發年終)
   ' and substr(sc02,1,4)-substr(st23,1,4)>=65 --> and sc02-st23>=650000
   'Modified by Morgan 2025/1/22 排除第4碼為9的
   stSQL = "insert into yearbonus (yb24,yb01,yb02,yb03,yb04,yb05,yb06,yb26) " & _
           "(select sd19 一公司別,substr(m1.sm02,1,4) 薪資年度,m1.sm01 員工代號,decode(sign(substr(m1.sm02,1,4)-" & Left(新部門啟用日, 4) & "+1),-1,st03,max(st93)) 部門," & _
           "round(sum(nvl(m1.sm25,0)+nvl(m2.sm25,0))/count(*),0) 平均基準月薪," & _
           "round(round(sum(nvl(m1.sm25,0)+nvl(m2.sm25,0))/count(*),0)*nvl(y1.ybm03,y2.ybm03)*decode(nvl(ym02,'2'),'1',110,'2',100,'3',85,'4',60,'*',100)/100*sum(nvl(m1.sm27,0))/" & m_YearDay & ",0) 年終獎金," & _
           "sd18 特殊功績獎金,sd51 紅利 from salarymonth m1,salarymonth m2,salarydata,staff,yearbonusmonth y1,yearbonusmonth y2,yearmerit,staff_change " & _
           "where sc03='08' and sc02>=" & Val(txtYEAR) + 1911 & "0101 and substr(sc01,4,1)<>'9' and sc01=st01(+) and st04='2' and sc02-st23>=650000 " & _
           "and st01=m1.sm01 and substr(m1.sm02,1,4)='" & Val(txtYEAR) + 1911 & "' and m1.sm25>0 " & _
           "and st01=sd01(+) and " & Val(txtYEAR) + 1911 & "=y1.ybm01(+) and sd19=y1.ybm02(+) and " & Val(txtYEAR) + 1911 & "=y2.ybm01(+) and '2'=y2.ybm02(+) " & _
           "and " & Val(txtYEAR) + 1911 & "=ym01(+) and sd01=ym03(+) " & _
           "and substr(m1.sm01,1,2)||'A'||substr(m1.sm01,4,2)=m2.sm01(+) and m1.sm02=m2.sm02(+) " & _
           "group by sd19,substr(m1.sm02,1,4),m1.sm01,st03,nvl(y1.ybm03,y2.ybm03),decode(nvl(ym02,'2'),'1',110,'2',100,'3',85,'4',60,'*',100),sd18,sd51)"
   
   cnnConnection.Execute stSQL, intR
   'modify by sonia 2021/2/22 改全年度
   'List1.AddItem time & " --> 新增下半年屆齡強制退休資料結束,共 " & intR & " 筆", 0
   List1.AddItem time & " --> 新增當年屆齡強制退休資料結束,共 " & intR & " 筆", 0
   'end 2018/2/9
   
   List1.AddItem time & " --> 更新每人未休假,缺勤,借支,所得稅及第二家年終獎金資料...", 0
   DoEvents
   '抓每人可休假時數,平均基準月薪以計算時薪,第二家公司別及第二家是否有基本薪資+午餐津貼,當年工作天數及二家年終獎金基準月薪以計算比例分二家資料
   '2010/1/11 MODIFY BY SONIA 未休假天數改抓特別假記錄檔YEARVACATION否則員工基本檔一旦更新為新年度可休假天數則會計算錯€
   'stSQL = "select yb01 獎金年度,yb02 員工代號,yb03 部門別,nvl(st40*8,0) 可休假時數,nvl(yb04,0) 平均基準月薪,nvl(yb05,0) 年終獎金,nvl(yb06,0) 特殊功績獎金,sd28 二公司別,nvl(sd29,0)+nvl(sd32,0) 二基午" & _
            ",sum(nvl(m1.sm27,0)) 工作天數,sum(nvl(m1.sm25,0)) 一年終獎金基準月薪,sum(nvl(m2.sm25,0)) 二年終獎金基準月薪 " & _
            "from yearbonus,salarydata,staff,salarymonth m1,salarymonth m2 " & _
            "where yb01='" & Val(txtYEAR) + 1911 & "' and yb02=st01(+) and yb02=sd01(+) " & _
            "and yb02=m1.sm01(+) and substr(m1.sm02,1,4)='" & Val(txtYEAR) + 1911 & "' " & _
            "and substr(m1.sm01,1,2)||'A'||substr(m1.sm01,4,2)=m2.sm01(+) and m1.sm02=m2.sm02(+) " & _
            "group by yb01,yb02,yb03,st40*8,yb04,yb05,yb06,sd28,nvl(sd29,0)+nvl(sd32,0)"
   '2013/1/22 modify by sonia 特殊工作時數,原固定用8,現改用PUB_intWkHour,故可休假時數改為先抓可休假天數再換算可休假時數
   'stSQL = "select yb01 獎金年度,yb02 員工代號,yb03 部門別,nvl(nvl(yv04,0)*8,0) 可休假時數,nvl(yb04,0) 平均基準月薪,nvl(yb05,0) 年終獎金,nvl(yb06,0) 特殊功績獎金,sd28 二公司別,nvl(sd29,0)+nvl(sd32,0) 二基午" & _
            ",sum(nvl(m1.sm27,0)) 工作天數,sum(nvl(m1.sm25,0)) 一年終獎金基準月薪,sum(nvl(m2.sm25,0)) 二年終獎金基準月薪 " & _
            "from yearbonus,salarydata,yearvacation,salarymonth m1,salarymonth m2 " & _
            "where yb01=" & Val(txtYEAR) + 1911 & " and yb01=yv01(+) and yb02=yv02(+) and yb02=sd01(+) " & _
            "and yb02=m1.sm01(+) and substr(m1.sm02,1,4)='" & Val(txtYEAR) + 1911 & "' " & _
            "and substr(m1.sm01,1,2)||'A'||substr(m1.sm01,4,2)=m2.sm01(+) and m1.sm02=m2.sm02(+) " & _
            "group by yb01,yb02,yb03,nvl(yv04,0)*8,yb04,yb05,yb06,sd28,nvl(sd29,0)+nvl(sd32,0)"
   'Modified by Morgan 2013/2/26 +yb24 新增補充保費公司別用
   'modify by sonia 2017/1/4 +st13到職日
   'modify by sonia 2018/1/30 +yb26紅利
   'modify by sonia 2018/2/9  +st04屆齡強制退休人員不發未休假代金
   'modify by sonia 2019/1/13 未休假代金改以計算年度12月的基本薪資+午餐津貼+職務津貼(sm26)計算
   'modify by sonia 2021/1/29 A1032留職停薪人員無12薪資無法計算未休假代金故改以個人當年在職最後一個月資料計算
   'stSQL = "select yb01 獎金年度,yb02 員工代號,yb03 部門別,nvl(yv04,0) 可休假天數,nvl(yb04,0) 平均基準月薪,nvl(yb05,0) 年終獎金,nvl(yb06,0) 特殊功績獎金,nvl(yb26,0) 紅利,sd28 二公司別,nvl(sd29,0)+nvl(sd32,0) 二基午" & _
            ",sum(nvl(m1.sm27,0)) 工作天數,sum(nvl(m1.sm25,0)) 一年終獎金基準月薪,sum(nvl(m2.sm25,0)) 二年終獎金基準月薪 ,max(yb24) yb24,nvl(st13,0) st13,st04,NVL(M3.SM26,0)+NVL(M4.SM26,0) 未休假代金計算月薪 " & _
            "from yearbonus,salarydata,yearvacation,salarymonth m1,salarymonth m2,staff,salarymonth m3,salarymonth m4 " & _
            "where yb01=" & Val(txtYEAR) + 1911 & " and yb01=yv01(+) and yb02=yv02(+) and yb02=sd01(+) and yb02=st01(+) " & _
            "and yb02=m1.sm01(+) and substr(m1.sm02,1,4)='" & Val(txtYEAR) + 1911 & "' " & _
            "and substr(m1.sm01,1,2)||'A'||substr(m1.sm01,4,2)=m2.sm01(+) and m1.sm02=m2.sm02(+) " & _
            "and m1.sm01=m3.sm01(+) and yb01||'12'=m3.sm02(+) and m2.sm01=m4.sm01(+) and yb01||'12'=m4.sm02(+) " & _
            "group by yb01,yb02,yb03,nvl(yv04,0),yb04,yb05,yb06,yb26,sd28,nvl(sd29,0)+nvl(sd32,0),st13,st04,NVL(M3.SM26,0)+NVL(M4.SM26,0)"
   stSQL = "select yb01 獎金年度,yb02 員工代號,yb03 部門別,nvl(yv04,0) 可休假天數,nvl(yb04,0) 平均基準月薪,nvl(yb05,0) 年終獎金,nvl(yb06,0) 特殊功績獎金,nvl(yb26,0) 紅利,sd28 二公司別,nvl(sd29,0)+nvl(sd32,0) 二基午" & _
            ",sum(nvl(m1.sm27,0)) 工作天數,sum(nvl(m1.sm25,0)) 一年終獎金基準月薪,sum(nvl(m2.sm25,0)) 二年終獎金基準月薪 ,max(yb24) yb24,nvl(st13,0) st13,st04,NVL(M3.SM26,0)+NVL(M4.SM26,0) 未休假代金計算月薪,st51 " & _
            "from yearbonus,salarydata,yearvacation,salarymonth m1,salarymonth m2,staff,salarymonth m3,salarymonth m4, " & _
            "(select yb02 empno,max(sm02) maxmm from yearbonus,salarymonth where yb01=" & Val(txtYEAR) + 1911 & " and yb02=sm01(+) and substr(sm02,1,4)='" & Val(txtYEAR) + 1911 & "' group by yb02) " & _
            "where yb01=" & Val(txtYEAR) + 1911 & " and yb01=yv01(+) and yb02=yv02(+) and yb02=sd01(+) and yb02=st01(+) " & _
            "and yb02=m1.sm01(+) and substr(m1.sm02,1,4)='" & Val(txtYEAR) + 1911 & "' " & _
            "and substr(m1.sm01,1,2)||'A'||substr(m1.sm01,4,2)=m2.sm01(+) and m1.sm02=m2.sm02(+) and yb02=empno(+) " & _
            "and m1.sm01=m3.sm01(+) and maxmm=m3.sm02(+) and m2.sm01=m4.sm01(+) and maxmm=m4.sm02(+) " & _
            "group by yb01,yb02,yb03,nvl(yv04,0),yb04,yb05,yb06,yb26,sd28,nvl(sd29,0)+nvl(sd32,0),st13,st04,NVL(M3.SM26,0)+NVL(M4.SM26,0),st51"
   
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL, , True)
   If intR = 1 Then
      With adoRst
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      Do While Not .EOF
         List1.AddItem time & " --> 更新 " & .Fields("員工代號") & " 年終獎金資料", 0
         ProgressBar1.Value = ProgressBar1.Value + 1
         Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         
         '抓整年出缺勤陣列
         strSql = " and ST01='" & Trim(.Fields("員工代號")) & "' "
         If PUB_GetAbsenceHour(strSql, (Val(txtYEAR) + 1911) * 10000 + "0101", (Val(txtYEAR) + 1911) * 10000 + "1231", douHour(), douCnt()) = True Then
         End If
         
         m_yb05 = Val(.Fields("年終獎金"))
         m_yb06 = Val(.Fields("特殊功績獎金"))
         m_yb26 = Val(.Fields("紅利"))    'add by sonia 2018/1/30
         
         '2013/1/22 add by sonia
         Call Pub_GetSpecWorkHour(.Fields("員工代號"), strSrvDate(1))
         '尤春彬84043和99029伊恩於101年工作時數有變動,要設定為該年度期初的工作時數
         If Val(txtYEAR) = "101" And (Trim(.Fields("員工代號")) = "84043" Or Trim(.Fields("員工代號")) = "99029") Then
            If Trim(.Fields("員工代號")) = "84043" Then
               PUB_intWkHour = 8
            Else
               PUB_intWkHour = 5
            End If
         End If
         m_yvhour = Val(.Fields("可休假天數")) * PUB_intWkHour
         '2013/1/22 end
         'add by sonia 2017/1/4 計算105年年終時,因勞基法修法,105/1/1~6/30到職人員加發3天未休假代金
         If Val(.Fields("st13")) >= 20160101 And Val(.Fields("st13")) <= 20160630 And strSrvDate(1) >= 20170101 And strSrvDate(1) <= 20170228 Then
            m_yvhour = 3 * PUB_intWkHour
         End If
         'end 2017/1/4
         
         '時薪
         m_HourPay = Val(.Fields("平均基準月薪")) / 30 / PUB_intWkHour
         '時薪  'add by sonia 2020/1/13 未休假代金改以計算年度12月的基本薪資+午餐津貼+職務津貼計算
         m_HourPayVC = Val(.Fields("未休假代金計算月薪")) / 30 / PUB_intWkHour
         
         '預設各假別不扣年終時數,依個人總工作天/年度總天數比例計算
'2013/1/28 modify by sonia 劉經理說依勞基法不管個人總工作天數,都以相同規則計算,99025蘇韋寧產假
'         m_SickHour = PUB_intWkHour * 30 * Val(.Fields("工作天數")) / m_YearDay
'         m_AbsenceHour = PUB_intWkHour * 14 * Val(.Fields("工作天數")) / m_YearDay
'         m_BornHour = PUB_intWkHour * 32 * Val(.Fields("工作天數")) / m_YearDay
'         m_NoBornHour = PUB_intWkHour * 7 * Val(.Fields("工作天數")) / m_YearDay
         m_SickHour = PUB_intWkHour * 30
         m_AbsenceHour = PUB_intWkHour * 14
         m_BornHour = PUB_intWkHour * 32
         m_NoBornHour = PUB_intWkHour * 7
         m_hurtHour = PUB_intWkHour * 7   'add by sonia 2016/1/6
         
         '計算個人年終缺勤扣款總時數
         m_SubHour = 0
         
         '累計事病假超過30日者扣除全部事病假
         If Val(douHour(5)) + Val(douHour(6)) > Val(m_SickHour) Then
            m_SubHour = m_SubHour + Val(douHour(5)) + Val(douHour(6))
         '累計事假超過14日者扣除全部事假
         ElseIf Val(douHour(5)) > Val(m_AbsenceHour) Then
            m_SubHour = m_SubHour + Val(douHour(5))
            douHour(6) = 0 '2010/2/4 ADD BY SONIA 婧瑄說改放扣年終獎金的請假時數
         '累計病假超過30日者扣除全部病假
         ElseIf Val(douHour(6)) > Val(m_SickHour) Then
            m_SubHour = m_SubHour + Val(douHour(6))
         '2010/2/4 ADD BY SONIA 婧瑄說改放扣年終獎金的請假時數
            douHour(5) = 0
         Else
            douHour(5) = 0
            douHour(6) = 0
         '2010/2/4 END
         End If
         
         '扣年終產假、扣年終流產假、公傷假扣除全部
         'modify by sonia 2016/1/6 公傷假改超過7者扣除超過部分
         'm_SubHour = m_SubHour + Val(douHour(17)) + Val(douHour(18)) + Val(douHour(13))
         m_SubHour = m_SubHour + Val(douHour(17)) + Val(douHour(18))
         
         '產假超過32日者扣除超過部分
         If Val(douHour(10)) > Val(m_BornHour) Then
            m_SubHour = m_SubHour + Val(douHour(10)) - m_BornHour
         '2010/2/4 ADD BY SONIA 婧瑄說改放扣年終獎金的請假時數
            douHour(10) = Val(douHour(10)) - m_BornHour
         Else
            douHour(10) = 0
         '2010/2/4 END
         End If
         
         '流產假超過7日者扣除超過部分
         If Val(douHour(11)) > Val(m_NoBornHour) Then
            m_SubHour = m_SubHour + Val(douHour(11)) - m_NoBornHour
         '2010/2/4 ADD BY SONIA 婧瑄說改放扣年終獎金的請假時數
            douHour(11) = Val(douHour(11)) - m_NoBornHour
         Else
            douHour(11) = 0
         '2010/2/4 END
         End If
         
         'add by sonia 2016/1/6 公傷假超過7日者扣除超過部分
         If Val(douHour(13)) > Val(m_hurtHour) Then
            m_SubHour = m_SubHour + Val(douHour(13)) - m_hurtHour
            douHour(13) = Val(douHour(13)) - m_hurtHour
         Else
            douHour(13) = 0
         End If
         'end 2016/1/6
         
         '曠職加每月遲到曠職時數,曠職時數扣三倍
         '2010/10/8 MODIFY BY SONIA 因遲到改當日逐筆輸,故改在PUB_GetAbsenceHour,此處不必再加
         'strExc(0) = "SELECT nvl(sum(decode(nvl(sa04,0),1,0,2,0,sa04*0.5)),0) 遲到曠職時數 FROM staff_assist " & _
         '            " where sa01='" & .Fields("員工代號") & "' and substr(sa02,1,4)='" & Val(txtYEAR) + 1911 & "'"
         'intI = 1
         'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         'm_yb11 = Val(douHour(3)) + Val(RsTemp.Fields("遲到曠職時數"))
         'Modified by Morgan 2025/8/15
         'm_yb11 = Val(douHour(3))
         m_yb11 = Round(Val(douHour(3)), 1)
         'end 2025/8/15
         '2010/10/8 END
         m_SubHour = m_SubHour + 3 * m_yb11
         
         '年終獎金+特殊功績獎金-缺勤扣款計算所得稅, 未休假代金及借支扣款不計入
         '所得稅以6%計算,低於2000不扣,高於2000者二家分別以6%計算
         '2010/12/30 modify by sonia 非固定之薪資所得扣繳稅率改抓 翻譯所得oc01='01'的稅率
         'm_TaxTotal = Round((m_yb05 + m_yb06 - (m_yb05 * m_SubHour / 8 / m_YearDay)) * 6 / 100, 0)
         
         'Modified by Morgan 2016/6/24
         '其它薪資 代號"50" 含年終/三節/翻譯等 所得73,001 以上扣稅
         'm_TaxTotal = Round((m_yb05 + m_yb06 - (m_yb05 * m_SubHour / PUB_intWkHour / m_YearDay)) * m_taxrate / 100, 0)
         '2010/12/30 end
         'If m_TaxTotal < 2000 Then m_TaxTotal = 0
         'modify by sonia 2018/1/30 +yb26
         'modify by sonia 2018/4/17 改84,501 以上扣稅
         'If Val(m_yb05 + m_yb06 + m_yb26 - (m_yb05 * m_SubHour / PUB_intWkHour / m_YearDay)) >= 73001 Then
         If Val(m_yb05 + m_yb06 + m_yb26 - (m_yb05 * m_SubHour / PUB_intWkHour / m_YearDay)) >= 84501 Then
            m_TaxTotal = Round((m_yb05 + m_yb06 + m_yb26 - (m_yb05 * m_SubHour / PUB_intWkHour / m_YearDay)) * m_taxrate / 100, 0)
         Else
            m_TaxTotal = 0
         End If
         'end 2016/6/24
         
         If m_bolIsTrial = False Then 'Added by Morgan 2023/12/12 非試算才執行
         
            '2013/1/17 ADD BY SONIA 補充保費
            NHI(1) = .Fields("員工代號")
            NHI(2) = strSrvDate(1)
            NHI(3) = "50"
            NHI(4) = "1"
            NHI(7) = m_yb05 + m_yb06 + m_yb26 - (m_yb05 * m_SubHour / PUB_intWkHour / m_YearDay)
            NHI(5) = 0: NHI(6) = 0: NHI(8) = 0
            NHI(10) = ServerTime
            NHI(11) = .Fields("yb24") 'Added by Morgan 2013/2/26
            PUB_NHI2nd NHI(1), NHI(2), NHI(3), NHI(4), NHI(7), NHI(5), NHI(6), NHI(8), NHI(10), NHI(11), NHI(13) 'Modified by Morgan 2013/3/12 +NHI13 2014/5/1 +NHI11
            
            '新增補充保費
            PUB_InsertNHI2nd NHI
            '2013/1/17 END
            
         End If
         
         '借支
         strExc(0) = "SELECT nvl(sum(nvl(ae03,0)),0) 借支 FROM Advance_Employee " & _
                     " where ae01='" & .Fields("員工代號") & "' and ae04=" & Val(txtYEAR) + 1911 & "13"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         
         
         '更新年終獎金資料其他欄位
         '2009/11/16 MODIFY BY SONIA 董事長,副董事長,4個律師(桂,蔣,廖,謝)因不打卡,不發未休假代金
         'cnnConnection.Execute "update YearBonus set yb07=" & CNULL(Val(.Fields("可休假時數")) - Val(douHour(8))) & ",yb08=" & CNULL(Round((Val(.Fields("可休假時數")) - Val(douHour(8))) * m_HourPay, 0)) & _
                               ",yb09=" & CNULL(Val(douHour(6))) & ",yb10=" & CNULL(Val(douHour(5))) & ",yb11=" & CNULL(Val(m_yb11)) & _
                               ",yb12=" & CNULL(Val(douHour(10)) + Val(Val(douHour(17)))) & ",yb13=" & CNULL(Val(douHour(11)) + Val(douHour(18))) & ",yb14=" & CNULL(Val(douHour(13))) & _
                               ",yb15=" & CNULL(Round(Val(m_yb05 * m_SubHour / 8 / m_YearDay)), 0) & ",yb16=" & CNULL(Val(RsTemp("借支"))) & ",yb17=" & CNULL(Val(m_TaxTotal)) & _
                               " where yb01='" & .Fields("獎金年度") & "' and yb02='" & .Fields("員工代號") & "'"
         'modify by sonia 2016/10/11 劉經理說律師僅桂,蔣,廖不必打卡,故不可用部門L01判斷
         'If .Fields("員工代號") = "63001" Or .Fields("員工代號") = "67004" Or .Fields("部門別") = "L01" Then
         'modify by sonia 2018/2/9 屆齡強制退休人員不發未休假代金
         'modify by sonia 2023/1/10 因為69009楊監察人在112/1/2退休但111年年終仍要發放,所以.Fields("st04") = "2"條件要再加離職日期判斷
         'If .Fields("員工代號") = "63001" Or .Fields("員工代號") = "67004" Or .Fields("員工代號") = "76012" Or .Fields("員工代號") = "79037" Or .Fields("員工代號") = "94015" Or .Fields("st04") = "2" Then
         If .Fields("員工代號") = "63001" Or .Fields("員工代號") = "67004" Or .Fields("員工代號") = "76012" Or .Fields("員工代號") = "79037" Or .Fields("員工代號") = "94015" Or (.Fields("st04") = "2" And .Fields("st51") <= Val(txtYEAR) + 1911 & "1231") Then
            '2013/1/17 MODIFY BY SONIA 加YB25補充保費
            'Modified by Morgan 2023/9/28 扣款由四捨五入(round)改為無條件捨去(trunc)
            cnnConnection.Execute "update YearBonus set yb07=0,yb08=0" & _
                               ",yb09=" & CNULL(Val(douHour(6))) & ",yb10=" & CNULL(Val(douHour(5))) & ",yb11=" & CNULL(Val(m_yb11)) & _
                               ",yb12=" & CNULL(Val(douHour(10)) + Val(Val(douHour(17)))) & ",yb13=" & CNULL(Val(douHour(11)) + Val(douHour(18))) & ",yb14=" & CNULL(Val(douHour(13))) & _
                               ",yb15=" & CNULL(Trunc(Val(m_yb05 * m_SubHour / PUB_intWkHour / m_YearDay)), 0) & ",yb16=" & CNULL(Val(RsTemp("借支"))) & ",yb17=" & CNULL(Val(m_TaxTotal)) & ",yb25=" & CNULL(Val(NHI(6))) & _
                               " where yb01='" & .Fields("獎金年度") & "' and yb02='" & .Fields("員工代號") & "'"
         Else
            '2013/1/17 modify BY SONIA 補充保費 2013/1/22 改可休假時數
            'cnnConnection.Execute "update YearBonus set yb07=" & CNULL(Val(.Fields("可休假時數")) - Val(douHour(8))) & ",yb08=" & CNULL(Round((Val(.Fields("可休假時數")) - Val(douHour(8))) * m_HourPay, 0)) & _
                               ",yb09=" & CNULL(Val(douHour(6))) & ",yb10=" & CNULL(Val(douHour(5))) & ",yb11=" & CNULL(Val(m_yb11)) & _
                               ",yb12=" & CNULL(Val(douHour(10)) + Val(Val(douHour(17)))) & ",yb13=" & CNULL(Val(douHour(11)) + Val(douHour(18))) & ",yb14=" & CNULL(Val(douHour(13))) & _
                               ",yb15=" & CNULL(Round(Val(m_yb05 * m_SubHour / 8 / m_YearDay)), 0) & ",yb16=" & CNULL(Val(RsTemp("借支"))) & ",yb17=" & CNULL(Val(m_TaxTotal)) & ",yb25=" & CNULL(Val(NHI(6))) & _
                               " where yb01='" & .Fields("獎金年度") & "' and yb02='" & .Fields("員工代號") & "'"
            'modify by sonia 2020/1/9 未休假代金改以計算年度12月的基本薪資+午餐津貼+職務津貼計算m_HourPayVC(原用m_HourPay)
            'Modified by Morgan 2023/9/28 扣款由四捨五入(round)改為無條件捨去(trunc)
            cnnConnection.Execute "update YearBonus set yb07=" & CNULL(Val(m_yvhour) - Val(douHour(8))) & ",yb08=" & CNULL(Round((Val(m_yvhour) - Val(douHour(8))) * m_HourPayVC, 0)) & _
                               ",yb09=" & CNULL(Val(douHour(6))) & ",yb10=" & CNULL(Val(douHour(5))) & ",yb11=" & CNULL(Val(m_yb11)) & _
                               ",yb12=" & CNULL(Val(douHour(10)) + Val(Val(douHour(17)))) & ",yb13=" & CNULL(Val(douHour(11)) + Val(douHour(18))) & ",yb14=" & CNULL(Val(douHour(13))) & _
                               ",yb15=" & CNULL(Trunc(Val(m_yb05 * m_SubHour / PUB_intWkHour / m_YearDay)), 0) & ",yb16=" & CNULL(Val(RsTemp("借支"))) & ",yb17=" & CNULL(Val(m_TaxTotal)) & ",yb25=" & CNULL(Val(NHI(6))) & _
                               " where yb01='" & .Fields("獎金年度") & "' and yb02='" & .Fields("員工代號") & "'"
         End If
         '2009/11/16 END
         
         '若有第二家則二家依比例計算,新增第二家,第一家再用減的以免有誤差
         '2009/1/13 modify by sonia 第二家0公司則不拆 79037
         'If .Fields("二基午") > 0 Then
         If .Fields("二基午") > 0 And .Fields("二公司別") <> "0" Then
            '新增第二家
            'modify by sonia 2018/1/30 +yb26紅利
            cnnConnection.Execute "insert into yearbonus (yb24,yb01,yb02,yb03,yb04,yb05,yb06,yb26) " & _
                                  "values (" & CNULL(.Fields("二公司別")) & "," & CNULL(.Fields("獎金年度")) & ",'" & Mid(.Fields("員工代號"), 1, 2) & "A" & Mid(.Fields("員工代號"), 4, 2) & "'" & _
                                  "," & CNULL(.Fields("部門別")) & "," & CNULL(.Fields("平均基準月薪")) & _
                                  ",Round(" & m_yb05 & " * " & Val(.Fields("二年終獎金基準月薪")) & " / (" & Val(.Fields("一年終獎金基準月薪")) & " + " & Val(.Fields("二年終獎金基準月薪")) & "), 0) " & _
                                  ",Round(" & m_yb06 & " * " & Val(.Fields("二年終獎金基準月薪")) & " / (" & Val(.Fields("一年終獎金基準月薪")) & " + " & Val(.Fields("二年終獎金基準月薪")) & "), 0) " & _
                                  ",Round(" & m_yb26 & " * " & Val(.Fields("二年終獎金基準月薪")) & " / (" & Val(.Fields("一年終獎金基準月薪")) & " + " & Val(.Fields("二年終獎金基準月薪")) & "), 0) ) "
            '更新第二家所得稅
            If m_TaxTotal > 0 Then
               '2010/12/30 modify by sonia 非固定之薪資所得扣繳稅率改抓 翻譯所得oc01='01'的稅率
               'cnnConnection.Execute "update YearBonus set yb17=round((yb05+yb06) * 6 / 100,0) " & _
                                     " where yb01='" & .Fields("獎金年度") & "' and yb02='" & Mid(.Fields("員工代號"), 1, 2) & "A" & Mid(.Fields("員工代號"), 4, 2) & "'"
               'modify by sonia 2018/1/30 +yb26紅利
               cnnConnection.Execute "update YearBonus set yb17=round((yb05+yb06+yb26) * " & m_taxrate & " / 100,0) " & _
                                     " where yb01='" & .Fields("獎金年度") & "' and yb02='" & Mid(.Fields("員工代號"), 1, 2) & "A" & Mid(.Fields("員工代號"), 4, 2) & "'"
               '2010/12/30 end
            End If
            '更新第一家
            cnnConnection.Execute "update YearBonus A set " & _
                                  "A.yb05=(select A.yb05 - B.yb05 from yearbonus B where A.yb01=B.yb01 and substr(A.yb02,1,2)||'A'||substr(A.yb02,4,2)=B.yb02), " & _
                                  "A.yb06=(select A.yb06 - B.yb06 from yearbonus B where A.yb01=B.yb01 and substr(A.yb02,1,2)||'A'||substr(A.yb02,4,2)=B.yb02), " & _
                                  "A.yb26=(select A.yb26 - B.yb26 from yearbonus B where A.yb01=B.yb01 and substr(A.yb02,1,2)||'A'||substr(A.yb02,4,2)=B.yb02), " & _
                                  "A.yb17=(select A.yb17 - B.yb17 from yearbonus B where A.yb01=B.yb01 and substr(A.yb02,1,2)||'A'||substr(A.yb02,4,2)=B.yb02) " & _
                                  " where yb01='" & .Fields("獎金年度") & "' and yb02='" & .Fields("員工代號") & "'"
         End If
         
         '檢查缺勤扣款第一家是否夠扣,不扣則以比例分二家扣
         'modify by sonia 2018/1/30 +yb26紅利
         strExc(0) = "SELECT nvl(yb05,0)+nvl(yb06,0)+nvl(yb26,0)+nvl(yb08,0)-nvl(yb15,0),nvl(yb15,0) FROM YearBonus " & _
                     " where yb01='" & .Fields("獎金年度") & "' and yb02='" & .Fields("員工代號") & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) < 0 Then
               cnnConnection.Execute "update YearBonus set " & _
                                     "yb15=round(" & RsTemp.Fields(1) & " * " & Val(.Fields("二年終獎金基準月薪")) & " / (" & Val(.Fields("一年終獎金基準月薪")) & " + " & Val(.Fields("二年終獎金基準月薪")) & "), 0) " & _
                                     " where yb01='" & .Fields("獎金年度") & "' and yb02='" & Mid(.Fields("員工代號"), 1, 2) & "A" & Mid(.Fields("員工代號"), 4, 2) & "'"
               cnnConnection.Execute "update YearBonus A set " & _
                                     "A.yb15=(select A.yb15 - B.yb15 from yearbonus B where A.yb01=B.yb01 and substr(A.yb02,1,2)||'A'||substr(A.yb02,4,2)=B.yb02) " & _
                                     " where yb01='" & .Fields("獎金年度") & "' and yb02='" & .Fields("員工代號") & "'"
            End If
         End If

         '檢查借支扣款第一家是否夠扣,不扣則以比例分二家扣
         'modify by sonia 2018/1/30 +yb26紅利
         strExc(0) = "SELECT nvl(yb05,0)+nvl(yb06,0)+nvl(yb26,0)+nvl(yb08,0)-nvl(yb15,0)-nvl(yb16,0),nvl(yb16,0) FROM YearBonus " & _
                     " where yb01='" & .Fields("獎金年度") & "' and yb02='" & .Fields("員工代號") & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) < 0 Then
               cnnConnection.Execute "update YearBonus set " & _
                                     "yb16=round(" & RsTemp.Fields(1) & " * " & Val(.Fields("二年終獎金基準月薪")) & " / (" & Val(.Fields("一年終獎金基準月薪")) & " + " & Val(.Fields("二年終獎金基準月薪")) & "), 0) " & _
                                     " where yb01='" & .Fields("獎金年度") & "' and yb02='" & Mid(.Fields("員工代號"), 1, 2) & "A" & Mid(.Fields("員工代號"), 4, 2) & "'"
               cnnConnection.Execute "update YearBonus A set " & _
                                     "A.yb16=(select A.yb16 - B.yb16 from yearbonus B where A.yb01=B.yb01 and substr(A.yb02,1,2)||'A'||substr(A.yb02,4,2)=B.yb02) " & _
                                     " where yb01='" & .Fields("獎金年度") & "' and yb02='" & .Fields("員工代號") & "'"
            End If
         End If

         .MoveNext
      Loop
      End With
   End If
   
   cnnConnection.CommitTrans
   Progress = True
   
   MsgBox "年終獎金計算完畢！", vbInformation
   List1.Clear
   
   List1.AddItem time & " --> 檢查是否仍有不夠扣的資料...", 0
   DoEvents
   '檢查是否仍有不夠扣的資料
   '2013/1/17 MODIF BY SONIA 加入補充保費YB25
   'modify by sonia 2018/1/30 +yb26紅利
   stSQL = "select yb02 員工代號, st02 姓名 from YearBonus,staff " & _
           "where yb01=" & Val(txtYEAR) + 1911 & " and yb02=st01(+) and nvl(yb05,0)+nvl(yb06,0)+nvl(yb26,0)+nvl(yb08,0)-nvl(yb15,0)-nvl(yb16,0)-nvl(yb17,0)-nvl(yb25,0)<0 " & _
           "order by yb02"
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL, , True)
   If intR = 1 Then
      With adoRst
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      Do While Not .EOF
         List1.AddItem time & " --> " & .Fields("員工代號") & " " & .Fields("姓名") & " 年終獎金不夠扣除減項 ! ", 0
         ProgressBar1.Value = ProgressBar1.Value + 1
         Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         .MoveNext
      Loop
      End With
      List1.AddItem time & " --> 檢查是否仍有不夠扣的資料結束...", 0
      MsgBox "注意 ! 年終獎金有不夠扣的資料 !", vbInformation
   Else
      InitialField
   End If
         
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
    
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
   InitialField
   
   If m_bolIsTrial Then lblNote.Visible = True 'Added by Morgan 2023/12/12
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170302 = Nothing
End Sub

Private Sub InitialField()
   txtYEAR = strSrvDate(2) \ 10000 - 1
   Label2 = "( 0/0 )"
   List1.Clear
   
   ReDim NHI(TF_NHI) As String  '2013/1/22 ADD BY SONIA
End Sub

Private Sub txtYEAR_GotFocus()
   CloseIme
   TextInverse txtYEAR
End Sub

Private Sub txtYEAR_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Function TxtValidate() As Boolean
   
   TxtValidate = True
   
   If Val(txtYEAR) = 0 Then
      ShowMsg "請輸入計算年度 !"
      TxtValidate = False
   End If
      
   If TxtValidate = True Then
      strExc(0) = "SELECT * FROM YearBonusMonth where ybm01=" & Val(txtYEAR) + 1911 & " order by ybm02"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI <> 1 Then
         MsgBox "尚未輸入 " & txtYEAR & " 年年終獎金基準月數資料  !", vbCritical
         TxtValidate = False
      End If
   End If
   
   '2013/1/17 ADD BY SONIA 計算年度再加1年,因為是隔年算年終,日期為計算的系統日
   If TxtValidate = True Then
      strExc(0) = "SELECT * FROM NHI2ND WHERE SUBSTR(NHI02,1,4)=" & Val(txtYEAR) + 1912 & " AND NVL(NHI09,0)>0 "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         MsgBox (" " & txtYEAR & " 年年終獎金補充保費已申報，不可重新計算  !")
         TxtValidate = False
      End If
   End If
   '2013/1/17 END
   
   '2013/1/23 ADD BY SONIA 檢查計算過年終之後是否還有其他非年終獎金資料
   If TxtValidate = True Then
      strExc(0) = "SELECT * FROM NHI2ND,(SELECT MIN(NHI02||NHI10) MIN FROM NHI2ND WHERE SUBSTR(NHI02,1,4)=" & Val(txtYEAR) + 1912 & " AND NHI04='1') A " & _
                  "WHERE SUBSTR(NHI02,1,4)=" & Val(txtYEAR) + 1912 & " AND NHI04<>'1' AND NHI02||NHI10>A.MIN"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         MsgBox (" " & txtYEAR & " 年年終獎金上次計算後已輸入其他５０格式之補充保費資料，不可重新計算  !")
         TxtValidate = False
      End If
   End If
   '2013/1/23 END
   
   If TxtValidate = True Then
      
      strExc(0) = "SELECT * FROM YearBonus where yb01=" & Val(txtYEAR) + 1911 & " order by yb02"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Added by Morgan 2023/12/12 增加補充保費判斷
         strExc(0) = "SELECT * FROM NHI2ND WHERE SUBSTR(NHI02,1,4)=" & Val(txtYEAR) + 1912 & " AND NHI04='1'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If m_bolIsTrial = True Then
               MsgBox txtYEAR & " 年年終獎金已計算，不可再試算  !", vbExclamation
               TxtValidate = False
               Exit Function
            Else
         'end 2023/12/12
            
               If MsgBox(" " & txtYEAR & " 年年終獎金已計算，是否重新計算 ?", vbYesNo + vbCritical) = vbNo Then
                  TxtValidate = False
               End If
               
            End If
         End If
      End If
   End If
   
   If TxtValidate = False Then TextInverse txtYEAR
   
   'Added by Morgan 2023/12/25
   If TxtValidate = True Then
      strExc(0) = "SELECT st02||'('||st01||')' FROM staff_change,staff" & _
         " where sc03='08' and sc02>=" & Val(txtYEAR) + 1911 & "0101 and sc01=st01(+) and st04='2' and substr(sc02,1,4)-substr(st23,1,4)>65" & _
         " and exists(select * from salarymonth where sm01=st01 and substr(sm02,1,4)=substr(sc02,1,4) and sm25>0)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         MsgBox RsTemp.GetString & "超過65歲退休！", vbInformation
      End If
   End If
   'end 2023/12/25
End Function
