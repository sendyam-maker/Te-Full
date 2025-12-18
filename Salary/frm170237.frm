VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm170237 
   BorderStyle     =   1  '單線固定
   Caption         =   "勞保/健保/勞退金明細"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8925
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4635
      Top             =   0
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   1845
      Max             =   200
      Min             =   150
      TabIndex        =   7
      Top             =   5490
      Value           =   200
      Width           =   4785
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   765
      MaxLength       =   3
      TabIndex        =   1
      Top             =   510
      Width           =   600
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7920
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   420
      Width           =   800
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7020
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   420
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4245
      Left            =   180
      TabIndex        =   4
      Top             =   930
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   7488
      _Version        =   393216
      BackColor       =   -2147483628
      Cols            =   5
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.ComboBox cboUser 
      Height          =   300
      Left            =   765
      TabIndex        =   0
      Top             =   120
      Width           =   2400
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "4233;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblStaffNo 
      AutoSize        =   -1  'True
      Caption         =   "員工："
      Height          =   180
      Left            =   225
      TabIndex        =   11
      Top             =   180
      Width           =   540
   End
   Begin VB.Label lblTimeOut 
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "若您未繼續移動滑鼠,將會於 59 秒後登出"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5220
      TabIndex        =   10
      Top             =   90
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "薪資資料濃淡設定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   9
      Top             =   5520
      Width           =   1560
   End
   Begin VB.Label lblTest 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      Caption         =   "這是濃淡設定預覽"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6795
      TabIndex        =   8
      Top             =   5520
      Width           =   1905
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "PS: 點選表格內數字欄位可查詢明細"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   5250
      Width           =   3030
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "年度："
      Height          =   180
      Index           =   0
      Left            =   225
      TabIndex        =   5
      Top             =   555
      Width           =   540
   End
End
Attribute VB_Name = "frm170237"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/14 Form2.0已修改
'Created by Morgan 2015/12/16
Option Explicit
Dim m_iCol As Integer, m_iRow As Integer
Dim m_StaffNoCon As String

Private Sub cboUser_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cboUser_Validate(Cancel As Boolean)
   Dim ii As Integer
   For ii = 0 To cboUser.ListCount - 1
      If InStr(cboUser.List(ii), cboUser) > 0 Then
         cboUser.ListIndex = ii
      End If
   Next
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdSearch_Click()
      
   If txtYear = "" Then
      MsgBox "請輸入年度！", vbExclamation
      txtYear.SetFocus
      Exit Sub
   End If
   
   cboUser_Validate False
   If cboUser.ListIndex < 0 Then
      MsgBox "請選擇員工！", vbExclamation
      cboUser.SetFocus
      Exit Sub
   End If
   
   QueryData

End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   PUB_AddSalaryUser cboUser, False
   
   'modify by sonia 2016/1/29
   'If Pub_MaxSMYM <> "" Then
   If Val(Pub_MaxSMYM) <> 0 Then
      txtYear = Left(Pub_MaxSMYM, 4) - 1911
   'add by sonia 2016/1/29
   Else
      txtYear = ""
   'end 2016/1/29
   End If
   
   PUB_SetForeColorScroll HScroll1
   SetGridColor
   PUB_EnableSalaryTimer
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170237 = Nothing
End Sub

Private Sub SetGridColor()
   lblTest.ForeColor = PUB_GetColor(HScroll1.Value)
   grdDataList.ForeColor = lblTest.ForeColor
End Sub

Private Sub HScroll1_Change()
   HScroll1_Scroll
End Sub

Private Sub HScroll1_Scroll()
   SetGridColor
   PUB_SaveForeColor HScroll1
End Sub

Private Sub Timer1_Timer()
PUB_ShowSalaryCountDown lblTimeOut
End Sub

Private Sub txtYear_GotFocus()
   TextInverse txtYear
   CloseIme
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub SetGrid()
   Dim iCol As Integer
   With grdDataList
      .Visible = False
      .FixedCols = 1
      .ColAlignmentFixed(0) = flexAlignCenterCenter
      .ColAlignment(0) = flexAlignCenterCenter
      For iCol = 1 To .Cols - 1
         .ColAlignmentFixed(iCol) = flexAlignCenterCenter
         .ColAlignment(iCol) = flexAlignRightCenter
         
         If iCol > 4 Then
            .ColWidth(iCol) = 0
         End If
      Next
      .Visible = True
   End With
End Sub

Private Sub QueryData()
   Dim stYrMn1 As String, stYrMn2 As String
   Dim arrTmp() As String
   
   arrTmp = Split(cboUser, " ")
   cboUser.Tag = arrTmp(1)
   
   txtYear.Tag = txtYear
   
   stYrMn1 = (Val(txtYear) + 1911) & "01"
   stYrMn2 = (Val(txtYear) + 1911) & "12"
   
   '不可早於啟用年月
   If stYrMn1 < Pub_StartYM Then
      stYrMn1 = Pub_StartYM
   End If
   
   '不可大於薪資入帳年月
   If stYrMn2 > Pub_MaxSMYM Then
      stYrMn2 = Pub_MaxSMYM
   End If
   
   m_StaffNoCon = " = '" & cboUser.Tag & "'"
   
   strExc(0) = "select s2.st01 from staff s1,staff s2 where s1.st01='" & cboUser.Tag & "' and s2.st26(+)=s1.st26"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.RecordCount = 1 Then
         m_StaffNoCon = " = '" & RsTemp(0) & "'"
      Else
         m_StaffNoCon = " in ('" & RsTemp.GetString(, , , "','") & cboUser.Tag & "')"
      End If
   End If
   
   '資料來源:1.月薪資 2.其他所得/扣款(代號:31-36) 3.非月薪資扣款之補充保費(只列出健保投保公司的資料--婧瑄)
   strExc(0) = "select substr(sm02,-2) 月份,to_char(sum(sm14),'999,999') 勞保費,to_char(sum(sm16),'999,999') 勞退自提,to_char(sum(sm43),'999,999') 補充保費,to_char(sum(sm15),'999,999') 健保費,sm02" & _
      " from (select sm02,sm14,sm16,sm43,sm15" & _
      " From salarymonth where sm01='" & cboUser.Tag & "' and sm02 between " & stYrMn1 & " and " & stYrMn2 & _
      " union all select od14,decode(od04,'31',od05,'32',-1*od05),decode(od04,'33',od05,'34',-1*od05)" & _
      ",0,decode(od04,'35',od05,'36',-1*od05)" & _
      " From othersalarydata where od03='" & cboUser.Tag & "' and od04>='31' and od04<='36' and od14 between " & stYrMn1 & " and " & stYrMn2 & _
      " union all select 1*substr(nhi02,1,6),0,0,nhi06,0" & _
      " From nhi2nd where nhi01" & m_StaffNoCon & " and nhi02 between " & stYrMn1 & "01 and " & stYrMn2 & "31 and nhi04 not in('3','6','7')" & _
      " and exists(select * from salarymonth where sm01" & m_StaffNoCon & " and sm02=substr(nhi02,1,6) and sm42>0)" & _
      " ) group by sm02  union all select '總計','0','0','0','0',0 from dual"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With grdDataList
      .FixedCols = 0
      Set .Recordset = RsTemp.Clone
      For intI = 1 To .Rows - 2
         .TextMatrix(intI, 0) = PUB_ChgNumber2Chinese(Val(.TextMatrix(intI, 0))) & "月"
         .TextMatrix(.Rows - 1, 1) = Format(Val(Format(.TextMatrix(.Rows - 1, 1))) + Val(Format(.TextMatrix(intI, 1))), DDollar2)
         .TextMatrix(.Rows - 1, 2) = Format(Val(Format(.TextMatrix(.Rows - 1, 2))) + Val(Format(.TextMatrix(intI, 2))), DDollar2)
         .TextMatrix(.Rows - 1, 3) = Format(Val(Format(.TextMatrix(.Rows - 1, 3))) + Val(Format(.TextMatrix(intI, 3))), DDollar2)
         .TextMatrix(.Rows - 1, 4) = Format(Val(Format(.TextMatrix(.Rows - 1, 4))) + Val(Format(.TextMatrix(intI, 4))), DDollar2)
      Next
      End With
      SetGrid
   End If
End Sub

Private Sub GetDeatil(pRow As Integer, pCol As Integer)
   Dim stYrMn1 As String, stYrMn2 As String, iCol As Integer, iRow As Integer, stConNhi As String, stVTB As String, stCon As String
   
   If pRow = grdDataList.Rows - 1 Then
      stYrMn2 = grdDataList.TextMatrix(pRow - 1, 5)
      stYrMn1 = Left(stYrMn2, 4) & "01"
   Else
      stYrMn1 = grdDataList.TextMatrix(pRow, 5)
      stYrMn2 = stYrMn1
   End If
   strExc(1) = (stYrMn1 \ 100 - 1911) & "年" & Val(Right(stYrMn1, 2)) & IIf(stYrMn2 = stYrMn1, "", "~" & Val(Right(stYrMn2, 2))) & "月份"
      
   Select Case pCol
   Case 1 '勞保費
      strExc(1) = strExc(1) & "勞保費明細"
      strExc(0) = "select trunc(sm02/100)-1911||'/'||mod(sm02,100) 年月" & _
         ",c1 說明,to_char(sm14,'999,999') 金額,sqldatet(br02) 代扣日期" & _
         " from (select '勞保費' c1,sm01,sm02,sm14,br02" & _
         " From salarymonth, bookrecord where sm01='" & cboUser.Tag & "' and sm14>0 and sm02 between " & stYrMn1 & " and " & stYrMn2 & " and br01(+)=sm02" & _
         " union all select oc03,od03,od14,decode(OC02,'A',-1,1)*od05,od02" & _
         " From othersalarydata, OtherSalaryCode" & _
         " where od03='" & cboUser.Tag & "' and od04 in ('31','32') and od14 between " & stYrMn1 & " and " & stYrMn2 & _
         " and oc01(+)=od04),staff" & _
         " where st01(+)=sm01 order by sm02,br02"
   
   Case 2 '勞退自提
      strExc(1) = strExc(1) & "勞退自提明細"
      strExc(0) = "select trunc(sm02/100)-1911||'/'||mod(sm02,100) 年月" & _
         ",c1 說明,to_char(sm16,'999,999') 金額,sqldatet(br02) 代扣日期" & _
         " from (select '勞退自提' c1,sm01,sm02,sm16,br02" & _
         " From salarymonth, bookrecord where sm01='" & cboUser.Tag & "' and sm16>0 and sm02 between " & stYrMn1 & " and " & stYrMn2 & " and br01(+)=sm02" & _
         " union all select oc03,od03,od14,decode(OC02,'A',-1,1)*od05,od02" & _
         " From othersalarydata, OtherSalaryCode" & _
         " where od03='" & cboUser.Tag & "' and od04 in ('33','34') and od14 between " & stYrMn1 & " and " & stYrMn2 & _
         " and oc01(+)=od04),staff" & _
         " where st01(+)=sm01 order by sm02,br02"
         
   Case 3 '補充保費
      strExc(1) = strExc(1) & "補充保費明細"
      'mhi04:1年終獎金 2,6每月獎金 3同仁其他 4翻譯費 0其他所得(目前不會有)
      stVTB = "select nhi11 x1,nhi01 x8,nhi05*4 x3,'年終獎金' x4,nhi02 x5,nhi07 x6,nhi06 x7,nhi01 x2,nhi14 x9" & _
         " From nhi2nd where nhi01='" & cboUser.Tag & "' and nhi03='50' and nhi05>0" & _
         " and nhi02 between " & stYrMn1 & "01 and " & stYrMn2 & "31 and nhi04='1'" & _
         " Union All select nhi11,nhi01, nhi05*4,mb14,nhi02,nhi07,nhi06,nhi01,nhi14" & _
         " From nhi2nd, monthbonus where nhi01='" & cboUser.Tag & "' and nhi03='50' and nhi05>0" & _
         " and nhi02 between " & stYrMn1 & "01 and " & stYrMn2 & "31 and nhi04 in ('2','6')" & _
         " and mb01(+)=nhi14 and mb02(+)=nhi01" & _
         " Union All select nhi11,nhi01, nhi05*4,'同仁其他',nhi02,nhi07,nhi06,nhi01,nhi14" & _
         " From nhi2nd where nhi01='" & cboUser.Tag & "' and nhi03='50' and nhi05>0" & _
         " and nhi02 between " & stYrMn1 & "01 and " & stYrMn2 & "31 and nhi04='3'" & _
         " Union All select nhi11,nhi01, nhi05*4,'翻譯費('||nhi01||')',nhi02,nhi07,nhi06,sm01,nhi14" & _
         " From nhi2nd,staff s1,staff s2,salarymonth where nhi01 " & m_StaffNoCon & " and nhi03='50' and nhi05>0" & _
         " and nhi02 between " & stYrMn1 & "01 and " & stYrMn2 & "31 and nhi04='4'" & _
         " and s1.st01(+)=nhi01 and s2.st26(+)=s1.st26 and sm01(+)=s2.st01 and sm02=substr(nhi02,1,6) and sm42>0" & _
         " Union All select nhi11,nhi01, nhi05*4,'其他所得',nhi02,nhi07,nhi06,nhi01,nhi14" & _
         " From nhi2nd where nhi01='" & cboUser.Tag & "' and nhi03='50' and nhi05>0" & _
         " and nhi02 between " & stYrMn1 & "01 and " & stYrMn2 & "31 and nhi04='0'"

      strExc(0) = "select trunc(x5/10000)-1911||'/'||mod(trunc(x5/100),100) 年月,sqldatet(x9) 給付日期,x4 名目,to_char(x6,'9,999,999') 給付金額,to_char(x7,'999,999') 補充保費,sqldatet(x5) 代扣日期,st02||'('||x8||')' 員工,x1 公司" & _
         " from (" & stVTB & "),staff where st01(+)=x2" & _
         " order by x5,x9"
         
   Case 4 '健保費
      strExc(1) = strExc(1) & "健保費明細"
      strExc(0) = "select trunc(hm03/100)-1911||'/'||mod(hm03,100) 年月,c1 說明" & _
         ",nvl(sr04,st02)||' ('||decode(sr03,'','自己','1','父親','2','母親','3','配偶','4','子女','其他')||')' 對象,to_char(hm04,'999,999') 金額,sqldatet(br02) 代扣日期" & _
         " from (select '健保費' c1,hm01,hm02,hm03,hm04,br02" & _
         " From himonth, bookrecord where hm01='" & cboUser.Tag & "' and hm03 between " & stYrMn1 & " and " & stYrMn2 & " and br01(+)=hm03" & _
         " union all select oc03,od03,od16,od14,decode(OC02,'A',-1,1)*od05,od02" & _
         " From othersalarydata, OtherSalaryCode" & _
         " where od03='" & cboUser.Tag & "' and od04 in ('35','36') and od14 between " & stYrMn1 & " and " & stYrMn2 & _
         " and oc01(+)=od04),staff_relation,staff" & _
         " where sr01(+)=hm01 and sr02(+)=hm02 and st01(+)=hm01 order by hm03,br02,hm02"
         
   Case Else
      Exit Sub
   End Select
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With frm880014
      .grdDataList.BackColor = grdDataList.BackColor
      Set .grdDataList.Recordset = RsTemp
      Set .fmParent = Me
      .bNewFormat = True
      With .grdDataList
      .ForeColor = lblTest.ForeColor '明細資料變淡
      .Cols = .Recordset.Fields.Count
      .ColAlignmentFixed(iCol) = flexAlignCenterCenter
      Select Case pCol
      Case 1 '勞保費
         .ColWidth(0) = 700
         .ColAlignment(0) = flexAlignCenterCenter
         .ColWidth(1) = 1350
         .ColAlignment(1) = flexAlignLeftCenter
         .ColWidth(2) = 1095
         .ColAlignment(2) = flexAlignRightCenter
         .ColWidth(3) = 900
         .ColAlignment(3) = flexAlignCenterCenter
         iCol = 3
         
      Case 2 '勞退自提
         .ColWidth(0) = 700
         .ColAlignment(0) = flexAlignCenterCenter
         .ColWidth(1) = 1350
         .ColAlignment(1) = flexAlignLeftCenter
         .ColWidth(2) = 1095
         .ColAlignment(2) = flexAlignRightCenter
         .ColWidth(3) = 900
         .ColAlignment(3) = flexAlignCenterCenter
         iCol = 3
         
      Case 3 '補充保費
         .ColWidth(0) = 700
         .ColAlignment(0) = flexAlignCenterCenter
         .ColWidth(1) = 900
         .ColAlignment(1) = flexAlignCenterCenter
         .ColWidth(2) = 2250
         .ColAlignment(2) = flexAlignLeftCenter
         .ColWidth(3) = 900
         .ColAlignment(3) = flexAlignRightCenter
         .ColWidth(4) = 900
         .ColAlignment(4) = flexAlignRightCenter
         .ColWidth(5) = 900
         .ColAlignment(5) = flexAlignCenterCenter
         iCol = 5
         
      Case 4 '健保費
         .ColWidth(0) = 700
         .ColAlignment(0) = flexAlignCenterCenter
         .ColWidth(1) = 1400
         .ColAlignment(1) = flexAlignLeftCenter
         .ColWidth(2) = 1600
         .ColAlignment(2) = flexAlignLeftCenter
         .ColWidth(3) = 1200
         .ColAlignment(3) = flexAlignRightCenter
         .ColWidth(4) = 900
         .ColAlignment(4) = flexAlignCenterCenter
         iCol = 5
      End Select
      For intI = iCol + 1 To .Cols - 1
         .ColWidth(intI) = 0
      Next
      .MergeCol(0) = True
      .MergeCells = flexMergeFree
      End With
      .Caption = strExc(1)
      .Show vbModal
      End With
   Else
      MsgBox "無明細資料！", vbInformation, strExc(1)
   End If
End Sub

Private Sub GrdDataList_Click()
   Dim iRow As Integer, iCol As Integer
   
   If grdDataList.Recordset Is Nothing Then Exit Sub
   With grdDataList
      iRow = .MouseRow
      iCol = .MouseCol
      If iRow > 0 And iCol > 0 Then
         .Enabled = False
         GetDeatil iRow, iCol
         .Enabled = True
      End If
   End With
End Sub

Private Sub grdDataList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
PUB_ResetSalaryTimer Me
'取消,反白還原再改濃度反白過的儲存格濃度不會用新的設定值
'   Dim iRow As Integer, iCol As Integer
'
'   If grdDataList.Recordset Is Nothing Then Exit Sub
'
'   With grdDataList
'      iRow = .MouseRow
'      iCol = .MouseCol
'      If iRow < 0 Or iCol < 0 Then Exit Sub
'      If iRow = m_iRow And iCol = m_iCol Then
'         Exit Sub
'      End If
'      If m_iCol <> 0 Then
'         .row = m_iRow: .col = m_iCol
'         '.CellBackColor = .BackColor
'         '.CellForeColor = .ForeColor
'         .CellTextStyle = .TextStyle
'         m_iRow = 0: m_iCol = 0
'      End If
'
'      If iRow > 0 And iRow <= .Rows - 1 Then
'         If iCol > 0 Then
'            .row = iRow: .col = iCol
'            '.CellBackColor = lblTest.ForeColor
'            '.CellForeColor = .BackColor
'            .CellTextStyle = flexTextInset
'            m_iCol = .col
'            m_iRow = .row
'         End If
'      End If
'   End With
End Sub
