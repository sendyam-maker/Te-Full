VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm170305 
   BorderStyle     =   1  '單線固定
   Caption         =   "各類所得轉入扣繳憑單－所得資料"
   ClientHeight    =   2784
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5064
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2784
   ScaleWidth      =   5064
   Begin VB.ListBox List1 
      Height          =   1128
      Left            =   45
      TabIndex        =   6
      Top             =   1530
      Width           =   4920
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Left            =   1140
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "96"
      Top             =   570
      Width           =   375
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "轉入(&T)"
      Height          =   405
      Left            =   2700
      TabIndex        =   1
      Top             =   60
      Width           =   1065
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   405
      Left            =   3870
      TabIndex        =   2
      Top             =   60
      Width           =   1065
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   45
      TabIndex        =   4
      Top             =   900
      Width           =   4950
      _ExtentX        =   8721
      _ExtentY        =   508
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "( 0/0 )"
      Height          =   255
      Left            =   90
      TabIndex        =   5
      Top             =   1230
      Width           =   4875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "資料年度：           年"
      Height          =   180
      Left            =   180
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "frm170305"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/23 Form2.0已檢查 (無需修改的物件)
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Memo by Morgan 2024/1/31 新部門已修改
'Create by Morgan 2009/1/14
Option Explicit


Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdok_Click()
   Screen.MousePointer = vbHourglass
   If Text1 <> "" Then
      Me.Enabled = False
      If Process = True Then
         MsgBox "成功!", vbInformation
      Else
         MsgBox "失敗!", vbCritical
      End If
      Me.Enabled = True
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Text1 = strSrvDate(2) \ 10000 - 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170305 = Nothing
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

Private Function Process() As Boolean
   Dim stSQL As String, intR As Integer
   Dim YR As Integer, YRW As Integer
   Dim stVTB(9) As String, stVTBs As String
   Dim adoRst As New ADODB.Recordset
   Dim iSNo As Integer, strUnitNo As String
   Dim strIDflag As String, strID As String, iIDtype As Integer
   
   YR = Val(Text1)
   YRW = YR + 1911
   
   cnnConnection.BeginTrans
   
On Error GoTo ErrHnd

   List1.Clear
   List1.AddItem time & " --> 刪除舊所得資料開始...", 0
   ProgressBar1.max = 1
   ProgressBar1.Value = 0
   Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
   DoEvents
   
   '刪除所得資料
   stSQL = "DELETE IncomeData WHERE ID14=" & YRW
   cnnConnection.Execute stSQL, intR
   
   ProgressBar1.Value = 1
   Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
   List1.AddItem time & " --> 刪除舊所得資料結束,共 " & intR & " 筆", 0
   
   List1.AddItem time & " --> 新增所得資料開始...", 0
   ProgressBar1.max = 1
   ProgressBar1.Value = 0
   Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
   DoEvents
   
   'SALARYMONTH 每月薪資(所得要減勞退自提)
   'Modify by Morgan 2009/5/25 修正給付總額未減勞退自提
   'Modify By Sindy 2020/6/25 + 證照津貼
   stVTB(1) = "SELECT '50' ID05,'' ID11,SM01 ID25,'2' ID27,SM37 Comp" & _
      ",NVL(SM04,0)+NVL(SM05,0)+NVL(SM45,0)+NVL(SM28,0)-NVL(SM21,0)-NVL(SM16,0) InCome" & _
      ",SM24 Tax,SM16 Reduce,SUBSTR(SM02,5)+0 MonA,SUBSTR(SM02,5)+0 MonB" & _
      " FROM SALARYMONTH WHERE SM02>" & YRW & "00 AND SM02<" & YRW & "99"
      
   'OTHERPAYDATA 其他給付
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'Modified by Morgan 2013/1/18 公司別改先抓od10(因若改過公司別抓薪資基本檔不對)
   'Modified by Morgan 2013/3/21 od01改放日期(原放年月)
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   stVTB(2) = "SELECT '50' ID05,'' ID11,OD02 ID25,'2' ID27,nvl(od10,DECODE(substr(OD02,3,1),'A',SD28,SD19)) Comp" & _
      " ,OD03 InCome,0 Tax,0 Reduce,SUBSTR(OD01,5,2)+0 MonA,SUBSTR(OD01,5,2)+0 MonB" & _
      " From OTHERPAYDATA,SALARYDATA" & _
      " WHERE OD01>" & YRW & "0000 AND OD01<" & YRW & "9999 AND SD01(+)=substr(od02,1,2)||replace(substr(od02,3,1),'A','0')||substr(od02,4)"
   
   'YEARBONUS 年終獎金(抓前一年),所得=年終獎金+特殊功績獎金-缺勤扣款
   'modify by sonia 2018/1/11 +YB26
   stVTB(3) = "SELECT '50' ID05,'' ID11,YB02 ID25,'2' ID27,YB24 Comp" & _
      " ,NVL(YB05,0)+NVL(YB06,0)+NVL(YB26,0)-NVL(YB15,0) InCome,YB17 Tax,0 Reduce,1 MonA,1 MonB" & _
      " From YEARBONUS WHERE YB01=" & (YRW - 1)
      
   'MonthBonus 每月獎金資料
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'Modified by Morgan 2013/3/21 MB01改放日期(原放年月)
   '2014/2/19 modify by sonia 每月獎金之公司別改抓MB11
   'stVTB(4) = "SELECT '50' ID05,'' ID11,MB02 ID25,'2' ID27,DECODE(substr(mb02,3,1),'A',SD28,SD19) Comp" & _
      " ,MB03 InCome,MB04 Tax,0 Reduce,SUBSTR(MB01,5,2)+0 MonA,SUBSTR(MB01,5,2)+0 MonB" & _
      " From MonthBonus,SALARYDATA" & _
      " WHERE MB01>" & YRW & "0000 AND MB01<" & YRW & "9999 AND SD01(+)=substr(MB02,1,1)||replace(substr(MB02,2),'A','0')"
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   stVTB(4) = "SELECT '50' ID05,'' ID11,MB02 ID25,'2' ID27,MB11 Comp" & _
      " ,MB03 InCome,MB04 Tax,0 Reduce,SUBSTR(MB01,5,2)+0 MonA,SUBSTR(MB01,5,2)+0 MonB" & _
      " From MonthBonus,SALARYDATA" & _
      " WHERE MB01>" & YRW & "0000 AND MB01<" & YRW & "9999 AND SD01(+)=substr(MB02,1,2)||replace(substr(MB02,3,1),'A','0')||substr(MB02,4)"
      
   'OHBONUS 端午、中秋獎金獎金資料(0公司不要)
   '2009/9/30 MODIFY BY SONIA 加OB05>0  因為2009年96011於九月底離職
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'Modified by Morgan 2012/8/14 台一投資(A公司)和9公司併入薪資格式
   'stVTB(5) = "SELECT '92' ID05,'8A' ID11,...
   'Modified by Morgan 2012/11/22 恢復 92 格式
   'stVTB(5) = "SELECT '50' ID05,'' ID11,OB03 ID25,'2' ID27,DECODE(substr(OB03,3,1),'A',SD28,SD19) Comp" & _
      " ,OB05 InCome,0 Tax,0 Reduce,SUBSTR(OB01,5)+0 MonA,SUBSTR(OB01,5)+0 MonB" & _
      " From OHBONUS,SALARYDATA" & _
      " WHERE OB01>" & YRW & "00 AND OB01<" & YRW & "99 and ob05>0 AND SD01(+)=substr(OB03,1,1)||replace(substr(OB03,2),'A','0') and DECODE(substr(OB03,3,1),'A',SD28,SD19) in ('A','9')" & _
      " union all SELECT '92' ID05,'8A' ID11,OB03 ID25,'2' ID27,DECODE(substr(OB03,3,1),'A',SD28,SD19) Comp" & _
      " ,OB05 InCome,0 Tax,0 Reduce,SUBSTR(OB01,5)+0 MonA,SUBSTR(OB01,5)+0 MonB" & _
      " From OHBONUS,SALARYDATA" & _
      " WHERE OB01>" & YRW & "00 AND OB01<" & YRW & "99 and ob05>0 AND SD01(+)=substr(OB03,1,1)||replace(substr(OB03,2),'A','0') and DECODE(substr(OB03,3,1),'A',SD28,SD19) not in ('A','9')"
   'modify by sonia 2016/1/18 端午、中秋獎金之公司別改抓OB12,不再抓sd19
   'stVTB(5) = "SELECT '92' ID05,'8A' ID11,OB03 ID25,'2' ID27,DECODE(substr(OB03,3,1),'A',SD28,SD19) Comp" & _
      " ,OB05 InCome,0 Tax,0 Reduce,SUBSTR(OB01,5)+0 MonA,SUBSTR(OB01,5)+0 MonB" & _
      " From OHBONUS,SALARYDATA" & _
      " WHERE OB01>" & YRW & "00 AND OB01<" & YRW & "99 and ob05>0 AND SD01(+)=substr(OB03,1,1)||replace(substr(OB03,2),'A','0')"
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   stVTB(5) = "SELECT '92' ID05,'8A' ID11,OB03 ID25,'2' ID27,NVL(OB12,DECODE(substr(OB03,3,1),'A',SD28,SD19)) Comp" & _
      " ,OB05 InCome,0 Tax,0 Reduce,SUBSTR(OB01,5)+0 MonA,SUBSTR(OB01,5)+0 MonB" & _
      " From OHBONUS,SALARYDATA" & _
      " WHERE OB01>" & YRW & "00 AND OB01<" & YRW & "99 and ob05>0 AND SD01(+)=substr(OB03,1,2)||replace(substr(OB03,3,1),'A','0')||substr(OB03,4)"
      
   'OtherIncomeData 其他各類所得資料
   stVTB(6) = "SELECT OID04 ID05,OID07 ID11,OID02 ID25" & _
      ",DECODE(OID04,'9A','1','50','2','51','5','92','9','A') ID27,OID03 Comp" & _
      " ,OID08 InCome,OID09 Tax,0 Reduce,OID05 MonA,OID06 MonB" & _
      " From OtherIncomeData WHERE OID01=" & YRW
   
   'BonusRetire 股利及退職所得資料(注意!!BR08放的是股利淨額,總額需加上可扣抵稅額,與其他的所得不同.)
   'Modify by Morgan 2010/1/25 股利的分配次數改1位除權基準年月日改7位
   stVTB(7) = "SELECT BR04 ID05,DECODE(BR04,'54',LPAD(BR10*100||'1'||LPAD(BR07-19110000,7,'0'),12,'0')) ID11" & _
      ",BR02 ID25,DECODE(BR04,'9A','1','50','2','51','5','92','9','A') ID27,BR03 Comp" & _
      ",NVL(BR08,0)+DECODE(BR04,'54',NVL(BR09,0),0)-NVL(BR13,0) InCome,BR09 Tax,BR13 Reduce,1 MonA,12 MonB" & _
      " From BonusRetire WHERE BR01=" & YRW
      
   '2014/1/14 add by sonia
   'OtherIncomeData 其他各類所得資料
   stVTB(8) = "SELECT OID04 ID05,OID07 ID11,OID02 ID25" & _
      ",DECODE(OID04,'9A','1','50','2','51','5','92','9','A') ID27,OID03 Comp" & _
      " ,OID08 InCome,OID09 Tax,0 Reduce,SUBSTR(OID05,5,2)+0 MonA,SUBSTR(OID05,5,2)+0 MonB" & _
      " From OtherIncomeDataDaily WHERE OID01=" & YRW
   '2014/1/14 end
      
   'add by sonia 2017/1/24
   'OtherSalaryData 其他所得/扣款資料 之補扣/退勞退自提
   stVTB(9) = "SELECT '50' ID05,'' ID11,SM01 ID25,'2' ID27,SM37 Comp" & _
      ",decode(od04,'33',-1*od05,'34',od05) InCome,0 Tax,decode(od04,'33',od05,'34',-1*od05) Reduce,0 MonA,0 MonB" & _
      " FROM SALARYMONTH,othersalarydata WHERE OD14>" & YRW & "00 AND OD14<" & YRW & "99 and od04 in ('33','34') AND OD14=SM02(+) AND OD03=SM01(+)"
   'end 2017/1/24
   
   stVTBs = stVTB(1)
   For intI = 2 To UBound(stVTB)
      stVTBs = stVTBs & " union all " & stVTB(intI)
   Next
   
   '新增薪資所得(0公司不要)
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'Modified by Morgan 2011/12/20 ID15,ID17 改長度
   'Modified by Morgan 2024/1/31 +以年度判斷是否抓新部門
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   stSQL = "insert into IncomeData(ID01,ID03,ID04,ID05,ID06,ID07,ID08,ID09,ID10,ID11,ID12,ID14,ID15" & _
      ",ID16,ID17,ID21,ID22,ID23,ID24,ID25,ID26,ID27,ID28)" & _
      " SELECT 'A08' ID01,A0807 ID03,DECODE(ID05,'54','C') ID04,ID05" & _
      ",DECODE(OI01,NULL,ST26,OI02) ID06" & _
      ",DECODE(LENGTH(OI02),8,'1',DECODE(OI01,NULL,DECODE(ST24,'F','3','0'),DECODE(OI03,NULL,'0','3'))) ID07" & _
      ",S1 ID08,S2 ID09,NVL(S1,0)-NVL(S2,0) ID10,ID11" & _
      ",'A' ID12," & YRW & " ID14" & _
      ",substrb(DECODE(OI01,NULL,ST02,OI04),1,40) ID15" & _
      ",substrb(DECODE(OI01,NULL,ST34,OI05),1,60) ID16" & _
      ",RPAD(DECODE(INSTR('50,93',ID05),0,CHR(32),LPAD(LTRIM(NVL(S3,0)),10,'0')),49,CHR(32)) ID17" & _
      ",TO_CHAR(SYSDATE,'MMDD')+0 ID21,S4 ID22,DECODE(SIGN(S5-12),1,12,S5) ID23" & _
      "," & IIf(YRW >= Left(新部門啟用日, 4), "ST93", "ST03") & " ID24,ID25" & _
      ",DECODE(OI01,NULL,DECODE(ST24,'F','2','1'),DECODE(OI03,NULL,'1','2')) ID26" & _
      ",DECODE(ID05,'54','1','9A','2','9B','2','50','3','51','5','93','9','A') ID27" & _
      ",DECODE(ID05,'54','C','9A','1','9B','2','51','1') ID28" & _
      " FROM (SELECT ID05,ID25,Comp,MAX(ID27) ID27,MAX(ID11) ID11" & _
      ",SUM(InCome) S1,SUM(Tax) S2,SUM(Reduce) S3,MIN(MonA) S4,MAX(MonB) S5" & _
      " FROM (" & stVTBs & ") GROUP BY ID05,ID25,Comp) S" & _
      ",ACC080,OtherIncomer,STAFF WHERE Comp<>'0'" & _
      " and A0801(+)=Comp AND OI01(+)=ID25 AND ST01(+)=substr(id25,1,2)||replace(substr(id25,3,1),'A','0')||substr(id25,4)"
      
   cnnConnection.Execute stSQL, intR
   
   ProgressBar1.Value = 1
   Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
   List1.AddItem time & " --> 新增所得資料結束,共 " & intR & " 筆", 0
   
   '2010/1/13 MODIFY BY SONIA
   'List1.AddItem time & " --> 更新流水號及錯誤註記...", 0
   List1.AddItem time & " --> 更新流水號......", 0
   
   '因為要更新錯誤註記,改跑回圈
   'stSQL = "update INCOMEDATA a set id02=(select count(*)+1" & _
      " from INCOMEDATA b where b.id03=a.id03 and" & _
      " nvl(b.id26,chr(32))||nvl(b.id07,chr(32))||nvl(b.id27,chr(32))" & _
      "||nvl(b.id28,chr(32))||rpad(b.id06,12,chr(23))||b.id05" & _
      "<nvl(a.id26,chr(32))||nvl(a.id07,chr(32))||nvl(a.id27,chr(32))" & _
      "||nvl(a.id28,chr(32))||rpad(a.id06,12,chr(23))||a.id05)"
   
   stSQL = "select id03,id26,id07,id27,id28,id06,id05,id25 from incomedata" & _
      " order by 1,2,3,4,5,6,7"
   With adoRst
   .CursorLocation = adUseClient
   .Open stSQL, cnnConnection, adOpenForwardOnly, adLockReadOnly
   If .RecordCount > 0 Then
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      DoEvents
      strUnitNo = .Fields("id03")
      iSNo = 0
      Do While Not .EOF
         If strUnitNo <> .Fields("id03") Then
            strUnitNo = .Fields("id03")
            iSNo = 0
         End If
         iSNo = iSNo + 1
         strID = "" & .Fields("id06")
         '判斷身分證號或統編
         If Len(strID) = 10 Then
            '外僑
            If "" & .Fields("id07") = "3" Then
               iIDtype = 2
            Else
               iIDtype = 0
            End If
         Else
            iIDtype = 1
         End If
         '檢查編號
         If CheckID(iIDtype, strID) = False Then
            strIDflag = "A"
         Else
            strIDflag = ""
         End If
         stSQL = "update incomedata set id02=" & iSNo & _
            ",id13='" & strIDflag & "'" & _
            " where id03='" & .Fields("id03") & "'" & _
            " and id05='" & .Fields("id05") & "'" & _
            " and id25='" & .Fields("id25") & "'"
            
         cnnConnection.Execute stSQL, intR
         
         ProgressBar1.Value = ProgressBar1.Value + 1
         Label2 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         .MoveNext
      Loop
      
   
      'Add by Morgan 2010/1/22 +股利要寫在共用欄位二
      'Modified by Morgan 2013/1/23 現金股利淨額後面+資本公積股利淨額[10位數字],後面轉出時會補空白
      'stSQL = "Update IncomeData set ID17=LPAD(ID10,10,'0')||LPAD('0',20,'0')||RPAD(CHR(32),8,CHR(32))" & _
         " where id14=" & YRW & " and id07 in ('0','1','3') and ID04='C' AND id05='54'"
      stSQL = "Update IncomeData set ID17=LPAD(ID10,10,'0')||LPAD('0',30,'0')" & _
         " where id14=" & YRW & " and id07 in ('0','1','3') and ID04='C' AND id05='54'"
         
      cnnConnection.Execute stSQL, intR
   
'      'add by sonia 2017/1/23 個人股利之共用欄位一前四碼之稅額扣抵比率要/2,否則媒體申報會錯
'      stSQL = "Update IncomeData set ID11=SUBSTR(TO_CHAR(SUBSTR(ID11,1,4)/2,'9999')||SUBSTR(ID11,5),2,12)" & _
'         " where id14=" & YRW & " and id07 in ('0','3','5','7') and ID04='C' AND id05='54'"
'      cnnConnection.Execute stSQL, intR
'      'end 2017/1/23
   
      '2010/1/13 MODIFY BY SONIA
      'List1.AddItem time & " --> 更新流水號及錯誤註記結束,共 " & .RecordCount & " 筆", 0
      List1.AddItem time & " --> 更新流水號結束,共 " & .RecordCount & " 筆", 0
   End If
   End With
   
   cnnConnection.CommitTrans
   Process = True
   GoTo XProtal
   
ErrHnd:
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical

XProtal:
   Set adoRst = Nothing
End Function

