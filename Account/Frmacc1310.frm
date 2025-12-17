VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc1310 
   AutoRedraw      =   -1  'True
   Caption         =   "翻譯費轉應付作業"
   ClientHeight    =   2076
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4764
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2076
   ScaleWidth      =   4764
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1665
      MaxLength       =   1
      TabIndex        =   0
      Top             =   300
      Width           =   465
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "執行"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1020
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   1380
      Width           =   2745
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1665
      TabIndex        =   1
      Top             =   750
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "（1.台一 2.智權）"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2190
      TabIndex        =   5
      Top             =   330
      Width           =   1785
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   330
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "入帳日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   780
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc1310"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/01 Form2.0已修改 (無需修改)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
'Add by Morgan 2007/6/1
Option Explicit


Private Sub Command1_Click()
   If ConCheck = True Then
      Screen.MousePointer = vbHourglass
      doBatch
      Screen.MousePointer = vbDefault
      FormClear
   End If
End Sub

Private Sub Form_Load()
   PUB_InitForm Me, 4890, 2490
   MaskEdBox1.Mask = DFormat
   
   'Added by Morgan 2020/3/26
   If strSrvDate(1) >= 智慧所更名日 Then
      Label3.Caption = ""
   End If
   'end 2020/3/26
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc1310 = Nothing
End Sub

Private Sub doBatch()
   Dim A0o0(1 To 19) As String, A1p0(1 To 31) As String
   Dim strLstCP14 As String, strLstST04 As String, strTName As String, strLstST03 As String
   Dim iSNo As Integer, dblTotFee As Double, dblFee As Double, dblTax As Double
   Dim stErrMsg As String
   Dim dblTaxRate As Double
   'Added by Morgan 2013/1/31
   Dim OD(14) As String
   Dim stNHI() As String
   Dim dblOdTax As Double
   Dim stCFP_TF22 As String 'Added by Morgan 2025/11/10
   
   ReDim stNHI(TF_NHI) As String
   
On Error GoTo Checking
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   
   'Add by Morgan 2011/2/8 翻譯費稅率改抓設定
   strExc(0) = "select oc04 from othersalarycode where oc01='01'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(RsTemp(0)) Then
         dblTaxRate = RsTemp(0) / 100
      End If
   End If
   'end 2011/2/8
   
   '依照員工編號排序
   'Modified by Morgan 2017/2/10 +TF21,SPR11 +判斷英文字數(or tf21>0)
   'Modified by Morgan 2019/7/3 不需再抓staff_idmap,新編號復職且關聯同一F編號時資料會重複 Ex:F5656顧家盛 ->A1025,A8008
   'strExc(0) = "select TF01,TF02,TF03,TF04,NVL(TF05,100) TF05,NVL(TF06,100) TF06,NVL(TF18,100) TF18,TF21,CP01,CP02,CP03,CP04,CP14,SPR01,SPR02,SPR03,SPR04,SPR11,S3.ST04,S4.ST02,S4.ST03" & _
      " from transfee T,caseprogress C,staff_payrate S1,staff_idmap S2,staff S3, STAFF S4" & _
      " where tf07 is null and (tf02>0 or tf21>0) and cp09(+)=tf01 and spr01(+)=cp14 and sim02(+)=cp14 and S3.st01(+)=sim01 AND S4.ST01(+)=CP14" & _
      " order by CP14,CP01,CP02,CP03,CP04,CP09"
   'Modified by Morgan 2019/8/14 +TF27,SPR12,SPR13,EP09
   'Modified by Morgan 2025/11/10 +TF28,SPR14,SPR15,SPR16
   strExc(0) = "select TF01,TF02,TF03,TF04,NVL(TF05,100) TF05,NVL(TF06,100) TF06,NVL(TF18,100) TF18,TF21,TF27,TF28,CP01,CP02,CP03,CP04,CP14,SPR01,SPR02,SPR03,SPR04,SPR11,SPR12,SPR13,SPR14,SPR15,SPR16,S4.ST02,S4.ST03,EP09" & _
      " from transfee T,caseprogress C,staff_payrate S1, STAFF S4,engineerprogress" & _
      " where tf07 is null and (tf02>0 or tf21>0) and cp09(+)=tf01 and spr01(+)=cp14 AND S4.ST01(+)=CP14 and ep02(+)=cp09" & _
      " order by CP14,CP01,CP02,CP03,CP04,CP09"
   'end 2019/7/3
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      '應付款資料
      A0o0(2) = "'1'"
      '入帳日期
      A0o0(5) = Val(FCDate(MaskEdBox1.Text))
      '欲處理日期->次月的10號
      A0o0(6) = CompDate(1, 1, A0o0(5) \ 100 & 10) - 19110000
      A0o0(13) = strSrvDate(2)
      A0o0(14) = "TO_CHAR(SYSDATE,'HH24MISS')"
      A0o0(15) = "'" & strUserNum & "'"
      A0o0(19) = "'1'"
      '分錄資料
      'A1p0(1) = "'1'" 'Modify By Sindy 2014/1/16 Mark
      
      adoTaie.BeginTrans
      
On Error GoTo Checking1

      'Added by Morgan 2013/1/31
      '其他所得
      OD(2) = DBDATE(A0o0(5))
      OD(4) = "01"
      '補充保費
      stNHI(2) = OD(2)
      stNHI(3) = "50"
      stNHI(4) = "4"
      If OD(2) = strSrvDate(1) Then
         stNHI(10) = ServerTime
      Else
         stNHI(10) = "235900"
      End If
      'end 2013/1/31
      
      stErrMsg = ""
      Do While Not .EOF
         'Add By Sindy 2014/1/16 特殊出名公司
         'Added by Morgan 2020/3/26
         If strSrvDate(1) >= 智慧所更名日 Then
            A1p0(1) = PUB_GetReceiptComp(.Fields("CP01"), .Fields("CP02"), .Fields("CP03"), .Fields("CP04"), True, True)
            If Text1 <> A1p0(1) Then
               GoTo RunNext
            End If
         Else
         'end 2020/3/26
         
            A1p0(1) = PUB_GetReceiptComp(.Fields("CP01"), .Fields("CP02"), .Fields("CP03"), .Fields("CP04"))
            
            If Text1 = "1" And A1p0(1) = "J" Then '1.台一
               GoTo RunNext
            ElseIf Text1 = "2" And A1p0(1) <> "J" Then '2.智權
               GoTo RunNext
            End If
            If A1p0(1) = "J" Then
               A1p0(1) = "'J'" '智權公司
            Else
               A1p0(1) = "'1'" '台一
            End If
            
         End If 'Added by Morgan 2020/3/26
         '2014/1/16 END
         
         If IsNull(.Fields("SPR01")) Then
            stErrMsg = "【" & .Fields("CP14") & " " & .Fields("ST02") & "】未設定翻譯費率無法計算，作業已全部取消！"
            GoTo Checking1
         End If
         If strLstCP14 <> "" & .Fields("CP14") Then
            OD(13) = "" 'Added by Morgan 2013/2/4
            'Acc1p0--貸方科目(前一筆)
            If strLstCP14 <> "" Then
               'Added by Morgan 2013/1/31
               '楊志雄 F5219,王雅萍 F5542 例外(非薪資,工作室有開發票)
               'If strLstCP14 <> "F5219" Or strLstCP14 <> "F5542" Then 'Removed by Morgan 2018/5/2 取消--婧瑄
               
                  '其他所得,內外翻都依相同規則扣稅(稅不滿2000不扣)
                  OD(3) = strLstCP14
                  OD(5) = dblTotFee
                  OD(6) = ""
                  
                  'Modified by Morgan 2016/6/24
                  '其它薪資 代號"50" 含年終/三節/翻譯等 所得73,001 以上扣稅
                  'dblOdTax = dblTotFee * dblTaxRate
                  'dblOdTax = Fix(dblOdTax)
                  'If dblOdTax >= 2000 Then
                  '   OD(6) = dblOdTax
                  'End If
                  'modify by sonia 2018/4/18 改84,501
                  'If dblTotFee >= 73001 Then
                  If dblTotFee >= 84501 Then
                     dblOdTax = dblTotFee * dblTaxRate
                     OD(6) = Fix(dblOdTax)
                  End If
                  'end 2016/6/24
                  
                  OD(13) = ""
                  '補充保費
                  stNHI(1) = OD(3)
                  stNHI(5) = ""
                  stNHI(6) = ""
                  stNHI(7) = OD(5)
                  stNHI(8) = ""
                  stNHI(11) = GetSalaryCompany(stNHI(1), stNHI(2))   'Added by Morgan 2013/3/1
                  
                  'Added by Morgan 2014/7/25
                  '一個人可能會有一筆以上的翻譯費 Ex.林信昌有多個外譯編號
                  If OD(2) = strSrvDate(1) Then
                     stNHI(10) = ServerTime
                  Else
                     stNHI(10) = "235900"
                  End If
                  SetNH10 stNHI(1), stNHI(2), stNHI(10)
                  'end 2014/7/25
                  
                  'Added by Morgan 2013/2/6
                  '檢查內翻人員不可有晚於該筆資料的補充保費
                  If PUB_ChkNHi2nd(stNHI(1), stNHI(2), stNHI(10), True) = False Then
                     GoTo EscPoint
                  End If
                  'end 2013/2/6
                  'Modified by Morgan 2013/3/12 +NHI13
                  'Modified by Morgan 2014/7/24 +NHI11
                  PUB_NHI2nd stNHI(1), stNHI(2), stNHI(3), stNHI(4), stNHI(7), stNHI(5), stNHI(6), stNHI(8), stNHI(10), stNHI(11), stNHI(13)
                  
                  OD(13) = Val(stNHI(6))
                  
                  strSql = "declare intMax number;begin select max(OD01)+1 into intMax from othersalarydata where substr(od02,1,4)=" & Left(OD(2), 4) & ";IF intMax IS NULL THEN intMax:=" & (Val(Left(OD(2), 4)) - 1911) & "00001; END IF;"
                  strSql = strSql & "INSERT INTO othersalarydata(od01,od02,od03,od04,od05,od06,od13)" & _
                        " VALUES(intMax," & OD(2) & ",'" & OD(3) & "','" & OD(4) & "'," & CNULL(OD(5), True) & "," & CNULL(OD(6), True) & "," & CNULL(OD(13)) & ");end;"
                  adoTaie.Execute strSql, intI
                                       
                  PUB_InsertNHI2nd stNHI
                  
               'End If 'Removed by Morgan 2018/5/2 取消--婧瑄
               'end 2013/1/31
               
               iSNo = iSNo + 1
               A1p0(3) = "'" & Format(iSNo, "000") & "'"
               '內翻
               'Modify by Morgan 2011/3/3 改判斷部門
               'If strLstST04 = "1" Then
               If strLstST03 = "F52" Then
                  A1p0(5) = "'2121'"
                  dblTax = 0
               '外翻
               Else
                  A1p0(5) = "'2113'"
                  '薪所稅=翻譯費*6%(超過2000才要) 例外:F5219 楊志雄 不扣稅款--辜
                  'Modified by Morgan 2013/2/4 +F5542 --辜
                  'Removed by Morgan 2018/5/2 取消--婧瑄
                  'If strLstCP14 = "F5219" Or strLstCP14 = "F5542" Then
                  '   dblTax = 0
                  'Else
                  'end 2018/5/2
                  
                     'Modify by Morgan 2009/12/14 fix 函數直接用會有問題
                     'Modify by Morgan 2011/2/8
                     'dblTax = dblTotFee * 0.06
                     
                     'Modified by Morgan 2016/6/24
                     '其它薪資 代號"50" 含年終/三節/翻譯等 所得73,001 以上扣稅
                     'dblTax = dblTotFee * dblTaxRate
                     'dblTax = Fix(dblTax)
                     'If dblTax >= 2000 Then
                     'modify by sonia 2018/4/18 改84,501
                     'If dblTotFee > 73001 Then
                     If dblTotFee > 84501 Then
                        dblTax = dblTotFee * dblTaxRate
                        dblTax = Fix(dblTax)
                     'end 2016/6/24
                        dblTotFee = dblTotFee - dblTax
                     Else
                        dblTax = 0
                     End If
                     
                  'End If 'Removed by Morgan 2018/5/2 取消--婧瑄
                  
                  'Added by Morgan 2013/2/4
                  '外翻加補充保費科目 2401
                  If Val(OD(13)) > 0 Then
                     '應付翻譯費要減補充保費
                     dblTotFee = dblTotFee - Val(OD(13))
                  End If
                  'end 2013/2/4
               End If
               A1p0(6) = "'TOT'"
               A1p0(7) = 0
               A1p0(8) = dblTotFee
               A1p0(14) = "'" & strTName & "/翻譯費'"
               A1p0(15) = "'" & strLstCP14 & "'"
               A1p0(18) = A0o0(5)
               strSql = "INSERT INTO ACC1p0(A1p01,A1p02,A1p03,A1p04,A1p05,A1p06,A1p07,A1p08,A1p14,A1p15,A1p18)" & _
                  " VALUES(" & A1p0(1) & "," & A1p0(2) & "," & A1p0(3) & "," & A1p0(4) & "," & A1p0(5) & "," & A1p0(6) & "," & A1p0(7) & "," & A1p0(8) & "," & A1p0(14) & "," & A1p0(15) & "," & A1p0(18) & ")"
               adoTaie.Execute strSql, intI
               If dblTax > 0 Then
                  iSNo = iSNo + 1
                  A1p0(3) = "'" & Format(iSNo, "000") & "'"
                  A1p0(5) = "'2401'"
                  A1p0(8) = dblTax
                  A1p0(14) = "'" & strTName & "/翻譯費稅款'"
                  A1p0(30) = "'" & A0o0(5) \ 10000 & "薪所稅'"
                  strSql = "INSERT INTO ACC1p0(A1p01,A1p02,A1p03,A1p04,A1p05,A1p06,A1p07,A1p08,A1p14,A1p15,A1p18,A1p30)" & _
                     " VALUES(" & A1p0(1) & "," & A1p0(2) & "," & A1p0(3) & "," & A1p0(4) & "," & A1p0(5) & "," & A1p0(6) & "," & A1p0(7) & "," & A1p0(8) & "," & A1p0(14) & "," & A1p0(15) & "," & A1p0(18) & "," & A1p0(30) & ")"
                  adoTaie.Execute strSql, intI
               End If
               
               'Added by Morgan 2013/2/4
               '外翻加補充保費科目 2401
               If strLstST03 <> "F52" And Val(OD(13)) > 0 Then
                  iSNo = iSNo + 1
                  A1p0(3) = "'" & Format(iSNo, "000") & "'"
                  'Modified by Morgan 2023/5/29 改2409 代收代付款--婉莘
                  'A1p0(5) = "'2401'"
                  A1p0(5) = "'2409'"
                  'end 2023/5/29
                  A1p0(8) = Val(OD(13))
                  A1p0(14) = "'" & strTName & "/補充健保費'"
                  'Modified by Morgan 2013/7/2 改對沖--辜
                  'A1p0(30) = "'" & A0o0(5) \ 10000 & "健保費'"
                  A1p0(30) = "'" & A0o0(5) \ 10000 & "補充保'"
                  strSql = "INSERT INTO ACC1p0(A1p01,A1p02,A1p03,A1p04,A1p05,A1p06,A1p07,A1p08,A1p14,A1p15,A1p18,A1p30)" & _
                     " VALUES(" & A1p0(1) & "," & A1p0(2) & "," & A1p0(3) & "," & A1p0(4) & "," & A1p0(5) & "," & A1p0(6) & "," & A1p0(7) & "," & A1p0(8) & "," & A1p0(14) & "," & A1p0(15) & "," & A1p0(18) & "," & A1p0(30) & ")"
                  adoTaie.Execute strSql, intI
               End If
               'end 2013/2/4
            End If
            dblTotFee = 0
            iSNo = 0
            '內翻
            'Modify by Morgan 2011/3/3 改判斷部門
            'If .Fields("ST04") = "1" Then
            If .Fields("ST03") = "F52" Then
               A1p0(2) = "'H'"
               '單據編號=員工號碼+入帳日期
               A1p0(4) = "'" & .Fields("CP14") & Format(A0o0(5), "000000") & "'"
            '外翻
            Else
               '新增應付資料
               A0o0(1) = "'" & AutoNo(MsgText(804), 5, 1) & "'"
               A0o0(3) = "'" & .Fields("CP14") & "'"
               'Modify By Sindy 2014/1/16 +A0O07
               strSql = "INSERT INTO ACC0O0(A0O01,A0O02,A0O03,A0O05,A0O06,A0O13,A0O14,A0O15,A0O19,A0O07)" & _
                  " VALUES(" & A0o0(1) & "," & A0o0(2) & "," & A0o0(3) & "," & A0o0(5) & "," & A0o0(6) & "," & A0o0(13) & "," & A0o0(14) & "," & A0o0(15) & "," & A0o0(19) & "," & A1p0(1) & ")"
               adoTaie.Execute strSql, intI
               
               A1p0(2) = "'B'"
               '單據編號=應付單號
               A1p0(4) = A0o0(1)
            End If
         End If
         'Acc1p0--借方科目
         iSNo = iSNo + 1
         A1p0(3) = "'" & Format(iSNo, "000") & "'"
         A1p0(5) = "'6130'"
         A1p0(6) = "'" & .Fields("CP01") & "'"
         
'Modified by Morgan 2017/2/10 改用函數
'         '日文翻譯費=日文字數*日文翻譯費率+(中文字數+數學式數)*中文打字費率
'         If Val("" & .Fields("TF03")) > 0 Then
'            '翻譯費,打字費要個別四捨五入
'            dblFee = Round(Val("" & .Fields("TF03")) * Val("" & .Fields("SPR03")) / 1000) + Round((Val("" & .Fields("TF02")) + Val("" & .Fields("TF04"))) * (Val("" & .Fields("SPR04")) / 1000))
'         '英文翻譯費=(中文字數+數學式數)*英文翻譯費率+(中文字數+數學式數)*中文打字費率
'         Else
'            '翻譯費,打字費要個別四捨五入
'            dblFee = Round((Val("" & .Fields("TF02")) + Val("" & .Fields("TF04"))) * Val("" & .Fields("SPR02")) / 1000) + Round((Val("" & .Fields("TF02")) + Val("" & .Fields("TF04"))) * (Val("" & .Fields("SPR04")) / 1000))
'         End If
'         '翻譯費=原翻譯費*相似折扣%*瑕疵折扣%*加成比率%
'         dblFee = Round(dblFee * Val("" & .Fields("TF05")) / 100 * Val("" & .Fields("TF06")) / 100 * Val("" & .Fields("TF18")) / 100)
         
         'Added by Morgan 2025/11/10
         If .Fields("cp01") = "CFP" Then
            If "" & .Fields("TF27") = "5" Then
               If .Fields("TF28") = "3" Then '中翻日
                  stCFP_TF22 = "" & .Fields("SPR15")
               ElseIf .Fields("TF28") = "4" Then '中翻德
                  stCFP_TF22 = "" & .Fields("SPR16")
               End If
            Else
               If .Fields("TF27") = "2" Then '日翻中
                  stCFP_TF22 = "" & .Fields("SPR13")
               ElseIf .Fields("TF27") = "3" Then '德翻中
                  stCFP_TF22 = "" & .Fields("SPR14")
               End If
            End If
            dblFee = PUB_GetTransFeeNew(Val("" & .Fields("TF21")), Val(stCFP_TF22), 0, 0, 0)
         'end 2025/11/10
         
         'Added by Morgan 2019/8/14
         '108.8.15 以後完稿案件改以原文字數計算翻譯費並取消中文打字費
         ElseIf Val("" & .Fields("ep09")) >= 20190815 Then
            dblFee = PUB_GetTransFeeNew(Val("" & .Fields("TF21")), IIf("" & .Fields("TF27") = "1", Val("" & .Fields("SPR12")), Val("" & .Fields("SPR13"))), Val("" & .Fields("TF05")), Val("" & .Fields("TF06")), Val("" & .Fields("TF18")))
            
         Else
         'end 2019/8/14
            dblFee = PUB_GetTransFee(Val("" & .Fields("TF02")), Val("" & .Fields("TF03")), Val("" & .Fields("TF04")), Val("" & .Fields("SPR02")), Val("" & .Fields("SPR03")), Val("" & .Fields("SPR04")), Val("" & .Fields("TF05")), Val("" & .Fields("TF06")), Val("" & .Fields("TF18")), Val("" & .Fields("TF21")), Val("" & .Fields("SPR11")))
            
         End If 'Added by Morgan 2019/8/14
         
'end 2017/2/10

         A1p0(7) = dblFee
         A1p0(8) = 0
         '2011/12/7 modify by sonia 配合結匯傳票摘要格式(收款日/收款金額 本所案號數字6碼 抬頭或申請人),但此處必定未收款故放/0
         'A1p0(14) = "'" & .Fields("CP01") & .Fields("CP02") & .Fields("CP03") & .Fields("CP04") & "/翻譯費'"
         'Modified by Morgan 2012/8/3 收據抬頭(客戶名稱)有造字時可能會造成語法錯誤
         'A1p0(14) = "'/0 " & .Fields("CP02") & " " & Left("" & GetA0K04("" & .Fields("CP01") & .Fields("CP02") & .Fields("CP03") & .Fields("CP04"), "" & .Fields("TF01")), 4) & "'"
         A1p0(14) = "rtrim('/0 " & .Fields("CP02") & " " & Left("" & GetA0K04("" & .Fields("CP01") & .Fields("CP02") & .Fields("CP03") & .Fields("CP04"), "" & .Fields("TF01")), 4) & " ')"
         'end 2012/8/3
         '2011/12/7 end
         A1p0(15) = "'" & .Fields("CP14") & "'"
         A1p0(17) = "'" & .Fields("CP01") & .Fields("CP02") & .Fields("CP03") & .Fields("CP04") & "'"
         A1p0(18) = A0o0(5)
         strSql = "INSERT INTO ACC1p0(A1p01,A1p02,A1p03,A1p04,A1p05,A1p06,A1p07,A1p08,A1p14,A1p15,A1p17,A1p18)" & _
            " VALUES(" & A1p0(1) & "," & A1p0(2) & "," & A1p0(3) & "," & A1p0(4) & "," & A1p0(5) & "," & A1p0(6) & "," & A1p0(7) & "," & A1p0(8) & "," & A1p0(14) & "," & A1p0(15) & "," & A1p0(17) & "," & A1p0(18) & ")"
         adoTaie.Execute strSql, intI
         
         dblTotFee = dblTotFee + dblFee
         strLstCP14 = "" & .Fields("CP14")
         'strLstST04 = "" & .Fields("ST04")
         strLstST03 = "" & .Fields("ST03") 'Add by Morgan 2011/3/3
         strTName = "" & .Fields("ST02")
         '更新翻譯費用檔
         'Added by Morgan 2025/11/10
         If .Fields("cp01") = "CFP" Then
            strSql = "UPDATE TRANSFEE SET TF07=" & A1p0(4) & ",TF14=" & A1p0(17) & ",TF22=" & Val(stCFP_TF22) & " WHERE TF01='" & .Fields("TF01") & "'"
         Else
         'end 2025/11/10
         
            'Modified by Morgan 2017/2/10 +TF22
            strSql = "UPDATE TRANSFEE SET TF07=" & A1p0(4) & ",TF14=" & A1p0(17) & ",TF15=" & Val("" & .Fields("SPR02")) & ",TF16=" & Val("" & .Fields("SPR03")) & ",TF17=" & Val("" & .Fields("SPR04")) & ",TF22=" & Val("" & .Fields("SPR11")) & " WHERE TF01='" & .Fields("TF01") & "'"
         End If
         adoTaie.Execute strSql, intI
         
RunNext: 'Add By Sindy 2014/1/17
         .MoveNext
      Loop
      
      OD(13) = "" 'Added by Morgan 2013/2/4
      'Added by Morgan 2013/1/31
      '楊志雄 F5219,王雅萍 F5542 例外(非薪資,工作室有開發票)
      'If strLstCP14 <> "F5219" Or strLstCP14 <> "F5542" Then 'Removed by Morgan 2018/5/2 取消--婧瑄
         '其他所得
         OD(3) = strLstCP14
         OD(5) = dblTotFee
         OD(6) = ""
         
         'Modified by Morgan 2016/6/24
         '其它薪資 代號"50" 含年終/三節/翻譯等 所得73,001 以上扣稅
         'dblOdTax = dblTotFee * dblTaxRate
         'dblOdTax = Fix(dblOdTax)
         'If dblOdTax >= 2000 Then
         '   OD(6) = dblOdTax
         'End If
         'modify by sonia 2018/4/18 改84,501
         'If dblTotFee >= 73001 Then
         If dblTotFee >= 84501 Then
            dblOdTax = dblTotFee * dblTaxRate
            OD(6) = Fix(dblOdTax)
         End If
         'end 2016/6/24
         
         OD(13) = ""
         '補充保費
         stNHI(1) = OD(3)
         stNHI(5) = ""
         stNHI(6) = ""
         stNHI(7) = OD(5)
         stNHI(8) = ""
         stNHI(11) = GetSalaryCompany(stNHI(1), stNHI(2))   'Added by Morgan 2013/3/1
         
         'Added by Morgan 2014/7/25
         '一個人可能會有一筆以上的翻譯費 Ex.林信昌有多個外譯編號
         If OD(2) = strSrvDate(1) Then
            stNHI(10) = ServerTime
         Else
            stNHI(10) = "235900"
         End If
         SetNH10 stNHI(1), stNHI(2), stNHI(10)
         'end 2014/7/25
         
         'Added by Morgan 2013/2/6
         '檢查內翻人員不可有晚於該筆資料的補充保費
         If PUB_ChkNHi2nd(stNHI(1), stNHI(2), stNHI(10), True) = False Then
            GoTo EscPoint
         End If
         'end 2013/2/6
         
         'Modified by Morgan 2013/3/12 +NHI13
         'Modified by Morgan 2014/7/24 +NHI11
         PUB_NHI2nd stNHI(1), stNHI(2), stNHI(3), stNHI(4), stNHI(7), stNHI(5), stNHI(6), stNHI(8), stNHI(10), stNHI(11), stNHI(13)
         OD(13) = Val(stNHI(6))
         
         strSql = "declare intMax number;begin select max(OD01)+1 into intMax from othersalarydata where substr(od02,1,4)=" & Left(OD(2), 4) & ";IF intMax IS NULL THEN intMax:=" & (Val(Left(OD(2), 4)) - 1911) & "00001; END IF;"
         strSql = strSql & "INSERT INTO othersalarydata(od01,od02,od03,od04,od05,od06,od13)" & _
               " VALUES(intMax," & OD(2) & ",'" & OD(3) & "','" & OD(4) & "'," & CNULL(OD(5), True) & "," & CNULL(OD(6), True) & "," & CNULL(OD(13)) & ");end;"
         adoTaie.Execute strSql, intI
         
            
         PUB_InsertNHI2nd stNHI
         
      'End If 'Removed by Morgan 2018/5/2 取消--婧瑄
      'end 2013/1/31
               
      'Acc1p0--貸方科目
      iSNo = iSNo + 1
      A1p0(3) = "'" & Format(iSNo, "000") & "'"
      '內翻
      'Modify by Morgan 2011/3/3 改判斷部門
      'If strLstST04 = "1" Then
      If strLstST03 = "F52" Then
         A1p0(5) = "'2121'"
         dblTax = 0
      '外翻
      Else
         A1p0(5) = "'2113'"
         '薪所稅=翻譯費*6%(超過2000才要) 例外:F5219 楊志雄 不扣稅款--辜
         'Modified by Morgan 2013/2/4 +F5542 --辜
         'Removed by Morgan 2018/5/2 取消--婧瑄
         'If strLstCP14 = "F5542" Then
         '   dblTax = 0
         'Else
         'end 2018/5/2
         
            'Modify by Morgan 2009/12/14 fix 函數直接用會有問題
            'Modify by Morgan 2011/2/8
            'dblTax = dblTotFee * 0.06
            
            'Modified by Morgan 2016/6/24
            '其它薪資 代號"50" 含年終/三節/翻譯等 所得73,001 以上扣稅
            'dblTax = dblTotFee * dblTaxRate
            'dblTax = Fix(dblTax)
            'If dblTax >= 2000 Then
            'modify by sonia 2018/4/18 改84,501
            'If dblTotFee > 73001 Then
            If dblTotFee > 84501 Then
               dblTax = dblTotFee * dblTaxRate
               dblTax = Fix(dblTax)
            'end 2016/6/24
               dblTotFee = dblTotFee - dblTax
            Else
               dblTax = 0
            End If
            
         'End If 'Removed by Morgan 2018/5/2 取消--婧瑄
         
         'Added by Morgan 2013/2/4
         '外翻加補充保費科目 2401
         If Val(OD(13)) > 0 Then
            '應付翻譯費要減補充保費
            dblTotFee = dblTotFee - Val(OD(13))
         End If
         'end 2013/2/4
         
      End If
      A1p0(6) = "'TOT'"
      '翻譯費=(中文字數+數學式數)*(英文翻譯費率+中文打字費率)+日文字數*(日文翻譯費率)
      A1p0(7) = 0
      A1p0(8) = dblTotFee
      A1p0(14) = "'" & strTName & "/翻譯費'"
      A1p0(15) = "'" & strLstCP14 & "'"
      A1p0(18) = A0o0(5)
      strSql = "INSERT INTO ACC1p0(A1p01,A1p02,A1p03,A1p04,A1p05,A1p06,A1p07,A1p08,A1p14,A1p15,A1p18)" & _
         " VALUES(" & A1p0(1) & "," & A1p0(2) & "," & A1p0(3) & "," & A1p0(4) & "," & A1p0(5) & "," & A1p0(6) & "," & A1p0(7) & "," & A1p0(8) & "," & A1p0(14) & "," & A1p0(15) & "," & A1p0(18) & ")"
      adoTaie.Execute strSql, intI
      If dblTax > 0 Then
         iSNo = iSNo + 1
         A1p0(3) = "'" & Format(iSNo, "000") & "'"
         A1p0(5) = "'2401'"
         A1p0(8) = dblTax
         A1p0(14) = "'" & strTName & "/翻譯費稅款'"
         A1p0(30) = "'" & A0o0(5) \ 10000 & "薪所稅'"
         strSql = "INSERT INTO ACC1p0(A1p01,A1p02,A1p03,A1p04,A1p05,A1p06,A1p07,A1p08,A1p14,A1p15,A1p18,A1p30)" & _
            " VALUES(" & A1p0(1) & "," & A1p0(2) & "," & A1p0(3) & "," & A1p0(4) & "," & A1p0(5) & "," & A1p0(6) & "," & A1p0(7) & "," & A1p0(8) & "," & A1p0(14) & "," & A1p0(15) & "," & A1p0(18) & "," & A1p0(30) & ")"
         adoTaie.Execute strSql, intI
      End If
      'Added by Morgan 2013/2/4
      '外翻加補充保費科目 2401
      If strLstST03 <> "F52" And Val(OD(13)) > 0 Then
         iSNo = iSNo + 1
         A1p0(3) = "'" & Format(iSNo, "000") & "'"
         'Modified by Morgan 2023/5/29 改2409 代收代付款--婉莘
         'A1p0(5) = "'2401'"
         A1p0(5) = "'2409'"
         'end 2023/5/29
         A1p0(8) = Val(OD(13))
         A1p0(14) = "'" & strTName & "/補充健保費'"
         'Modified by Morgan 2013/7/2 改對沖--辜
         'A1p0(30) = "'" & A0o0(5) \ 10000 & "健保費'"
         A1p0(30) = "'" & A0o0(5) \ 10000 & "補充保'"
         strSql = "INSERT INTO ACC1p0(A1p01,A1p02,A1p03,A1p04,A1p05,A1p06,A1p07,A1p08,A1p14,A1p15,A1p18,A1p30)" & _
            " VALUES(" & A1p0(1) & "," & A1p0(2) & "," & A1p0(3) & "," & A1p0(4) & "," & A1p0(5) & "," & A1p0(6) & "," & A1p0(7) & "," & A1p0(8) & "," & A1p0(14) & "," & A1p0(15) & "," & A1p0(18) & "," & A1p0(30) & ")"
         adoTaie.Execute strSql, intI
      End If
      'end 2013/2/4
               
      adoTaie.CommitTrans
      End With
   Else
      MsgBox "無待轉資料！"
   End If
   
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
   Exit Sub
Checking1:
   adoTaie.RollbackTrans
   
Checking:
   If stErrMsg <> "" Then
      MsgBox stErrMsg
   ElseIf Err.Number <> 0 Then
      MsgBox Err.Description, , MsgText(5)
   End If
   
EscPoint:

End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1.Text = "" 'Add By Sindy 2014/1/16
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox Label1 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label1 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
End Sub

Private Function ConCheck() As Boolean
   Dim bCancel As Boolean
   Dim strMsg As String 'Add by Amy 2014/10/30
   
   MaskEdBox1_Validate bCancel
   If bCancel = True Then
      Exit Function
   End If
   
   'Add By Sindy 2014/1/16
   If Trim(Text1) = "" Then
      MsgBox "公司別不可空白!!", , MsgText(5)
      Text1.SetFocus
      Exit Function
   End If
   '2014/1/16 END
   
   'Add by Amy 2014/10/30 +系統日期比較
   If ChkWorkData(IIf(Text1 = "1", Text1, "J"), DBDATE(MaskEdBox1), strMsg) = False Then
        MsgBox Label1 & strMsg, , MsgText(5)
        MaskEdBox1.SetFocus
        Exit Function
    End If
   'end 2014/10/30
   
   'Remove by Morgan 2007/10/5 開放可轉部份資料(已輸入字數的)--婧瑄
   'strExc(0) = "select * from transfee,caseprogress,engineerprogress where tf02 is null and cp09(+)=tf01 and ep02(+)=tf01 and ep09>0 and rownum<2"
   'intI = 1
   'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'If intI = 1 Then
   '   MsgBox "尚有案件未輸入中文字數【如" & RsTemp.Fields("cp01") & "-" & RsTemp.Fields("cp02") & "-" & RsTemp.Fields("cp03") & "-" & RsTemp.Fields("cp04") & "】！"
   '   Exit Function
   'End If
   'end 2007/10/5
   
   'Added by Morgan 2023/4/27
   If PUB_ExistsSalaryMonth(DBDATE(MaskEdBox1)) = True Then
      MsgBox "入帳日期的月薪資已計算，請先取消後再執行！", vbExclamation
      Exit Function
   End If
   'end 2023/4/27
         
   ConCheck = True
End Function

Private Sub Text1_Change()
   'Added by Morgan 2020/3/26
   If strSrvDate(1) >= 智慧所更名日 Then
      Label3 = A0802Query(Text1)
   End If
   'end 2020/3/26
End Sub

'Add By Sindy 2014/1/16
Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

'Add By Sindy 2014/1/16
Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   
   Select Case KeyAscii
    'Modified by Morgan 2020/3/26+74,76
    Case 49, 50, 74, 76, 8
        '無動作
        If strSrvDate(1) >= 智慧所更名日 Then
            If KeyAscii = 50 Then KeyAscii = 0
        End If
    Case Else
        KeyAscii = 0
   End Select
    
End Sub

'Added by Morgan 2014/7/25
Private Sub SetNH10(pNHI01 As String, pNHI02 As String, pNHI10 As String)
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   
   stSQL = "select nhi10 from staff s1,staff s2,nhi2nd where s1.st01='" & pNHI01 & "' and s2.st26(+)=s1.st26 and nhi01(+)=s2.st01 and nhi02=" & pNHI02 & " order by nhi10 desc"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      pNHI10 = rsQuery(0) + 1
   End If
   Set rsQuery = Nothing
End Sub
