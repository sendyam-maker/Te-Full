VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm170236 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工薪資明細"
   ClientHeight    =   5772
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8892
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5772
   ScaleWidth      =   8892
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6450
      Top             =   510
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   1845
      Max             =   255
      Min             =   200
      TabIndex        =   10
      Top             =   5490
      Value           =   200
      Width           =   4785
   End
   Begin VB.TextBox txtMn 
      Height          =   285
      Index           =   1
      Left            =   2790
      MaxLength       =   3
      TabIndex        =   3
      Top             =   660
      Width           =   375
   End
   Begin VB.TextBox txtMn 
      Height          =   285
      Index           =   0
      Left            =   2250
      MaxLength       =   3
      TabIndex        =   2
      Top             =   660
      Width           =   375
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7020
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   540
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7920
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   540
      Width           =   800
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   765
      MaxLength       =   3
      TabIndex        =   1
      Top             =   630
      Width           =   600
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4185
      Left            =   180
      TabIndex        =   6
      Top             =   960
      Width           =   8550
      _ExtentX        =   15071
      _ExtentY        =   7387
      _Version        =   393216
      BackColor       =   -2147483628
      Cols            =   5
      FixedCols       =   2
      ScrollTrack     =   -1  'True
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
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
      Top             =   270
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
   Begin VB.Label lblCompNo 
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
      Height          =   180
      Left            =   5160
      TabIndex        =   18
      Top             =   705
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "投保單位統一編號："
      Height          =   180
      Index           =   3
      Left            =   3510
      TabIndex        =   17
      Top             =   705
      Width           =   1620
   End
   Begin VB.Label lblCompName 
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
      Height          =   180
      Left            =   4980
      TabIndex        =   16
      Top             =   330
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "勞健保投保單位："
      Height          =   180
      Index           =   2
      Left            =   3510
      TabIndex        =   15
      Top             =   330
      Width           =   1440
   End
   Begin VB.Label lblStaffNo 
      AutoSize        =   -1  'True
      Caption         =   "員工："
      Height          =   180
      Left            =   225
      TabIndex        =   14
      Top             =   330
      Width           =   540
   End
   Begin VB.Label lblTimeOut 
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "若您未繼續移動滑鼠,將會於 59 秒後登出"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5430
      TabIndex        =   13
      Top             =   30
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PS: 點選表格內灰底反白的數字欄位可查詢明細"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   135
      TabIndex        =   12
      Top             =   5250
      Width           =   4005
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
      TabIndex        =   11
      Top             =   5520
      Width           =   1905
   End
   Begin VB.Line Line1 
      X1              =   2520
      X2              =   2835
      Y1              =   810
      Y2              =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "月份："
      Height          =   180
      Index           =   1
      Left            =   1710
      TabIndex        =   9
      Top             =   705
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "年度："
      Height          =   180
      Index           =   0
      Left            =   225
      TabIndex        =   8
      Top             =   675
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "薪資資料濃淡設定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   5520
      Width           =   1560
   End
End
Attribute VB_Name = "frm170236"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/14 Form2.0已修改
'Created by Morgan 2015/12/23
Option Explicit


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
      
   ElseIf txtMn(0) = "" Then
      MsgBox "請輸入月份(起)！", vbExclamation
      txtMn(0).SetFocus
      Exit Sub
      
   ElseIf Val(txtMn(0)) < 1 Or Val(txtMn(0)) > 12 Then
      MsgBox "月份(起)輸入錯誤！", vbExclamation
      txtMn(0).SetFocus
      Exit Sub
      
   ElseIf txtMn(1) = "" Then
      MsgBox "請輸入月份(迄)！", vbExclamation
      txtMn(1).SetFocus
      Exit Sub
      
   ElseIf Val(txtMn(1)) < 1 Or Val(txtMn(1)) > 12 Then
      MsgBox "月份(迄)輸入錯誤！", vbExclamation
      txtMn(1).SetFocus
      Exit Sub
      
   ElseIf Val(txtMn(1)) < Val(txtMn(0)) Then
      MsgBox "月份(迄)不可小於月份(起)！", vbExclamation
      txtMn(1).SetFocus
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
   
   PUB_AddSalaryUser cboUser
      
   'modify by sonia 2016/1/29
   'If Pub_MaxSMYM <> "" Then
   If Val(Pub_MaxSMYM) <> 0 Then
      txtYear = Left(Pub_MaxSMYM, 4) - 1911
      txtMn(0) = Val(Right(Pub_MaxSMYM, 2))
      txtMn(1) = txtMn(0)
   'add by sonia 2016/1/29
   Else
      txtYear = ""
      txtMn(0) = ""
      txtMn(1) = ""
   'end 2016/1/29
   End If
   
   'Added by Morgan 2019/7/1
   lblCompName = ""
   lblCompNo = ""
   'end 2019/7/1

   PUB_SetForeColorScroll HScroll1
   SetGridColor
   PUB_EnableSalaryTimer
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_ResetSalaryTimer Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170236 = Nothing
End Sub

Private Sub SetGridColor()
   lblTest.ForeColor = PUB_GetColor(HScroll1.Value)
   grdDataList.ForeColor = lblTest.ForeColor
   
   Dim iRow As Integer, iCol As Integer
   
   If grdDataList.Recordset Is Nothing Then Exit Sub
   With grdDataList
   .Visible = False
   'add by sonia 2016/8/5 F編號可查翻譯費明細
   If Left(Right(cboUser, 5), 1) = "F" Then
      For iRow = 3 To 16 '15
         Select Case iRow
         'Modify By Sindy 2020/6/24
         'Case 3, 15
         Case 3, 16
         '2020/6/24 END
            .row = iRow
            For iCol = 2 To .Cols - 1
               .col = iCol
               .CellBackColor = lblTest.ForeColor
               .CellForeColor = .BackColor
            Next
         End Select
      Next
   Else
   'end 2016/8/5
      For iRow = 6 To 25 '24
         Select Case iRow
         'Modify By Sindy 2020/6/24
         'Case 6, 11, 14, 15, 18, 19, 20, 22, 23, 24
         Case 6, 12, 15, 16, 19, 20, 21, 23, 24, 25
         '2020/6/24 END
            .row = iRow
            For iCol = 2 To .Cols - 1
               .col = iCol
               .CellBackColor = lblTest.ForeColor
               .CellForeColor = .BackColor
            Next
         End Select
      Next
   End If   'add by sonia 2016/8/5
   .Visible = True
   End With
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

Private Sub txtMn_GotFocus(Index As Integer)
   TextInverse txtMn(Index)
   CloseIme
End Sub

Private Sub txtMn_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
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

Private Sub QueryData()
   Dim stYrMn1 As String, stYrMn2 As String, stConSm02 As String
   Dim arrList() As String, arrTmp() As String
   Dim lngYM As Long
   
   'Modify By Sindy 2020/6/24 29 => 30
   'ReDim arrList(29) As String
   ReDim arrList(30) As String
   arrList(1) = ";;sm02"
   arrList(2) = ";支給日數;sm27"
   arrList(3) = "應發薪資;基本薪資;sm04"
   arrList(4) = "應發薪資;伙食津貼;sm07"
   arrList(5) = "應發薪資;職務津貼;sm05"
   arrList(6) = "應發薪資;加 班 費;sm12"
   arrList(7) = "應發薪資;差旅津貼;sm08"
   arrList(8) = "應發薪資;證照津貼;sm45" 'Add By Sindy 2020/6/24
   arrList(9) = "應發薪資;技術津貼;sm06"
   arrList(10) = "應發薪資;房租津貼;sm09"
   arrList(11) = "應發薪資;特 支 費;sm10"
   arrList(12) = "應發薪資;其他所得;sm13"
   'Modify By Sindy 2020/6/24 +nvl(sm45,0)
   arrList(13) = ";應發總額;nvl(sm04,0)+nvl(sm07,0)+nvl(sm05,0)+nvl(sm12,0)+nvl(sm08,0)+nvl(sm45,0)+nvl(sm06,0)+nvl(sm09,0)+nvl(sm10,0)+nvl(sm13,0)"
   arrList(14) = "應扣項目;勞 保 費;sm14"
   arrList(15) = "應扣項目;健 保 費;sm15"
   arrList(16) = "應扣項目;補充保費;sm43"
   arrList(17) = "應扣項目;退休金自提;sm16"
   arrList(18) = "應扣項目;所 得 稅;sm24"
   arrList(19) = "應扣項目;互助會會款;sm18"
   arrList(20) = "應扣項目;員工貸款;sm19"
   arrList(21) = "應扣項目;借    支;sm20"
   arrList(22) = "應扣項目;未打卡扣款;sm22"
   arrList(23) = "應扣項目;婚喪喜慶;sm17"
   arrList(24) = "應扣項目;其它扣款;sm23"
   arrList(25) = "應扣項目;缺勤扣薪;sm21"
   arrList(26) = ";應扣總額;nvl(sm14,0)+nvl(sm15,0)+nvl(sm43,0)+nvl(sm16,0)+nvl(sm24,0)+nvl(sm18,0)+nvl(sm19,0)+nvl(sm20,0)+nvl(sm22,0)+nvl(sm17,0)+nvl(sm23,0)+nvl(sm21,0)"
   'Modify By Sindy 2020/6/24 +nvl(sm45,0)
   arrList(27) = ";實發金額;nvl(sm04,0)+nvl(sm07,0)+nvl(sm05,0)+nvl(sm12,0)+nvl(sm08,0)+nvl(sm45,0)+nvl(sm06,0)+nvl(sm09,0)+nvl(sm10,0)+nvl(sm13,0)-(nvl(sm14,0)+nvl(sm15,0)+nvl(sm43,0)+nvl(sm16,0)+nvl(sm24,0)+nvl(sm18,0)+nvl(sm19,0)+nvl(sm20,0)+nvl(sm22,0)+nvl(sm17,0)+nvl(sm23,0)+nvl(sm21,0))"
   arrList(28) = ";;null"
   arrList(29) = ";事務所提繳退休金;sm30"
   arrList(30) = ";勞退自提比率;sm44"
  
   arrTmp = Split(cboUser, " ")
   cboUser.Tag = arrTmp(1)
   
   txtYear.Tag = txtYear
   txtMn(0).Tag = txtMn(0)
   txtMn(1).Tag = txtMn(1)
   
   stYrMn1 = (Val(txtYear) + 1911) * 100 + Val(txtMn(0))
   stYrMn2 = (Val(txtYear) + 1911) * 100 + Val(txtMn(1))
   
   '不可早於啟用年月
   If stYrMn1 < Pub_StartYM Then
      stYrMn1 = Pub_StartYM
   End If
   
   '不可大於薪資入帳年月
   If stYrMn2 > Pub_MaxSMYM Then
      stYrMn2 = Pub_MaxSMYM
   End If
   
   If stYrMn1 = stYrMn2 Then
      stConSm02 = " = " & stYrMn1
   Else
      stConSm02 = " between " & stYrMn1 & " and " & stYrMn2
   End If
      
   strExc(0) = ""
   For intI = 1 To UBound(arrList)
      arrTmp = Split(arrList(intI), ";")
      If intI > 1 Then strExc(0) = strExc(0) & " union all "
      strExc(0) = strExc(0) & "select " & intI & " idx"
      For lngYM = Val(stYrMn1) To Val(stYrMn2)
         Select Case intI
         'Modify By Sindy 2020/6/24
         'Case 1, 2, 27, 29
         Case 1, 2, 28, 30
         '2020/6/24 END
            strExc(0) = strExc(0) & ",to_char(sum(decode(sm02," & lngYM & "," & arrTmp(2) & "))) " & PUB_ChgNumber2Chinese(lngYM Mod 100) & "月"
         Case Else
            strExc(0) = strExc(0) & ",to_char(sum(decode(sm02," & lngYM & "," & arrTmp(2) & ")),'999,999') " & PUB_ChgNumber2Chinese(lngYM Mod 100) & "月"
         End Select
      Next
      strExc(0) = strExc(0) & " from salarymonth where sm01='" & cboUser.Tag & "' and sm02" & stConSm02
   Next
   strExc(0) = strExc(0) & " order by idx"
      
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With grdDataList
      .WordWrap = True
      .FixedCols = 2
      Set .Recordset = RsTemp.Clone
      For intI = 1 To UBound(arrList)
         If intI < .Rows Then
            arrTmp = Split(arrList(intI), ";")
            .TextMatrix(intI, 0) = "" & arrTmp(0)
            .TextMatrix(intI, 1) = "" & arrTmp(1)
         End If
      Next
      End With
      SetGrid
      SetCompany cboUser.Tag, stYrMn2 'Added by Morgan 2019/7/1 顯示最後月份的投保單位及統編--
   Else
      grdDataList.Clear
      MsgBox "無資料！", vbInformation
   End If
End Sub

'Added by Morgan 2019/7/1
'設定公司別(投保單位)資料
Private Sub SetCompany(pNo As String, pYM As String)
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   stSQL = "select sm37,a0807,a0802 from salarymonth,acc080 " & _
      " where sm01='" & pNo & "' and sm02=" & pYM & " and a0801(+)=sm37"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      lblCompName = rsQuery.Fields("a0802")
      lblCompNo = rsQuery.Fields("a0807")
   End If
   Set rsQuery = Nothing
End Sub

Private Sub SetGrid()
   Dim iCol As Integer
   
   With grdDataList
   .Visible = False
   
   .ColWordWrapOption(0) = True '自動折行
   .MergeCol(0) = True
   .MergeCells = flexMergeFree
   .ColAlignmentFixed = flexAlignCenterCenter
   .ColWidth(0) = 450
   .ColAlignment(0) = flexAlignCenterCenter
   .ColWidth(1) = 1300
   .ColAlignment(1) = flexAlignCenterCenter
   .ColWidth(2) = 0 'idx
   .RowHeight(1) = 0 'sm02
   .RowHeight(28) = 1.5 * .RowHeight(27) '事務所提繳退休金
   
   For iCol = 3 To .Cols - 1
      If "" & .TextMatrix(1, iCol) = "" Then
         .ColWidth(iCol) = 0
      Else
         .ColWidth(iCol) = 950
         .ColAlignment(iCol) = flexAlignRightCenter
      End If
   Next
   SetGridColor
   .Visible = True
   End With
End Sub

Private Sub GetDeatil(pRow As Integer, pCol As Integer)
   Dim stYrMn As String, iCol As Integer, iRow As Integer, stConNhi As String, stVTB As String, stCon As String
   'Added by Morgan 2017/7/11
   Dim iDailyHours As Integer '每日工時
   Dim dblHourPay As Double '時薪
   Dim dblHours As Double '時數
   Dim dblDHours As Double '計算時數
   Dim dblAmt As Double '金額
   Dim iOtType As Integer '加班費類別:1=平日, 2=休假日, 3=休息日
   Dim dblSickHour As Double '病假時數累計
   Dim dblSickBaseHour As Double '病假扣半薪時數
   Dim dblGirlSickHour As Double '生理假時數累計
   Dim dblGirlSickBaseHour As Double '生理假半薪時數
   
   stYrMn = grdDataList.TextMatrix(1, pCol)
   
   If stYrMn = "" Then Exit Sub
   strExc(1) = (stYrMn \ 100 - 1911) & "年" & Val(Right(stYrMn, 2)) & "月份"
   
   'add by sonia 2016/8/5
   '翻譯費明細
   If pRow = 3 Then
      strExc(0) = "select * From acc1p0,staff_idmap,staff,ACC250 where substr(a1p18+19110000,1,6)='" & stYrMn & "' and a1p15='" & cboUser.Tag & "' and a1p05='6130' and sim02(+)=a1p15 and st01(+)=sim01 and A2502(+)='5' AND A2503(+)=A1P15 AND A2505(+)=A1P04"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then   '有資料才進下一畫面
         With frm170236_1
            .FstYrMn = grdDataList.TextMatrix(1, pCol)
            .grdDataList.BackColor = grdDataList.BackColor
            Set .fmParent = Me
            .bNewFormat = True
            .Show vbModal
         End With
      Else
         MsgBox "無明細資料！", vbInformation, strExc(1)
      End If
      Exit Sub
   End If
   'end 2016/8/5
   
   strExc(0) = ""
   Select Case pRow
   Case 6 '加班費
      strExc(1) = strExc(1) & "加班時數明細"
      'Modified by Morgan 2017/7/11 +平日加班費, 休假日加班費, 休息日加班費
      'Modified by Morgan 2023/6/7 +可能隔月給付 substr(so02,1,6)-->so16
      'strExc(0) = "select sqldatet(so02) 日期,nvl(so05,so06) 時數, 0 平日加班費, 0 休假日加班費, 0 休息日加班費,so01,so02,so05,so06 From Staff_Overtime where substr(so02,1,6)='" & stYrMn & "' and so01='" & cboUser.Tag & "' order by 1"
      strExc(0) = "select sqldatet(so02) 日期,nvl(so05,so06) 時數, 0 平日加班費, 0 休假日加班費, 0 休息日加班費,so01,so02,so05,so06 From Staff_Overtime where so16=" & stYrMn & " and so01='" & cboUser.Tag & "' order by 1"
      'end 2023/6/7
   Case 12 '11 '其它所得
      strExc(1) = strExc(1) & "其它所得明細"
      'modify by sonia 2016/9/5 退健保費時,備註欄加眷屬姓名
      'strExc(0) = "select sqldatet(od02) 日期,oc03 項目,to_char(od05,'999,999') 金額,to_char(od06,'999,999') 稅金,od15 備註 From OtherSalaryData, OtherSalaryCode where od03='" & cboUser.Tag & "' and substr(od02,1,6)='" & stYrMn & "' and oc02='A' and od04<>'01' and od04<>'12' and oc01(+)=od04 order by 1"
      strExc(0) = "select sqldatet(od02) 日期,oc03 項目,to_char(od05,'999,999') 金額,to_char(od06,'999,999') 稅金,od15||decode(od04,'36',decode(od16,'0',null,decode(od15,null,null,'-')||sr04),null) 備註 From OtherSalaryData, OtherSalaryCode, staff_relation where od03='" & cboUser.Tag & "' and substr(od02,1,6)='" & stYrMn & "' and oc02='A' and od04<>'01' and od04<>'12' and oc01(+)=od04 and od03=sr01(+) and od16=sr02(+) order by 1"
   Case 15 '14 '健保費
      strExc(1) = strExc(1) & "健保費明細"
      strExc(0) = "select trunc(hm03/100)-1911||'/'||mod(hm03,100) 年月" & _
         ",nvl(sr04,st02)||' ('||decode(sr03,'','自己','1','父親','2','母親','3','配偶','4','子女','其他')||')' 對象,to_char(hm04,'999,999') 金額" & _
         " From himonth,staff_relation,staff where hm01='" & cboUser.Tag & "' and hm03=" & stYrMn & " and sr01(+)=hm01 and sr02(+)=hm02 and st01(+)=hm01 order by hm02"
   Case 16 '15 '補充保費
      strExc(1) = strExc(1) & "補充保費明細"
      If Left(cboUser.Tag, 1) = "F" Then
         strExc(0) = "select sqldatet(nhi14) 給付日期,'翻譯費' 名目,to_char(nhi07,'999,999') 給付金額,to_char(nhi06,'999,999') 補充保費 from nhi2nd where nhi01='" & cboUser.Tag & "' and substr(nhi02,1,6)='" & stYrMn & "' and nhi04='4' "
      Else
         strExc(0) = "select sqldatet(nhi14) 給付日期,'同仁其他給付' 名目,to_char(nhi07,'999,999') 給付金額,to_char(nhi06,'999,999') 補充保費 from nhi2nd where nhi01='" & cboUser.Tag & "' and substr(nhi02,1,6)='" & stYrMn & "' and nhi04='3' " & _
            " union all select sqldatet(nhi14) 給付日期,mb14 名目,to_char(nhi07,'9,999,999') 給付金額,to_char(nhi06,'999,999') 補充保費 from nhi2nd,monthbonus where nhi01='" & cboUser.Tag & "' and substr(nhi02,1,6)='" & stYrMn & "' and nhi04='6' and mb01(+)=nhi14 and mb02(+)=nhi01 and mb01>0" & _
            " union all select sqldatet(nhi14) 給付日期,'非投保公司月薪資' 名目,to_char(nhi07,'999,999') 給付金額,to_char(nhi06,'999,999') 補充保費 from nhi2nd where nhi01='" & cboUser.Tag & "' and substr(nhi02,1,6)='" & stYrMn & "' and nhi04='7' " & _
            " order by 1"
      End If
   Case 19 '18 '互助會會款
      strExc(1) = strExc(1) & "互助會會款扣款明細"
      strExc(0) = "select sqldatet(WFA01) 日期,wfa02 會號,to_char(wfa04,'999,999') 金額 From WFAmount where substr(wfa01,1,6)='" & stYrMn & "' and wfa03='" & cboUser.Tag & "' and wfa05='2' order by 1,2"
   Case 20 '19 '員工貸款
      strExc(1) = strExc(1) & "員工貸款還款明細"
      strExc(0) = "select sqldatet(le02) 貸款日期,to_char(le03,'999,999') 貸款本金,to_char(le04,'999,999') 利息,le05||'-'||le06 償還時間,to_char(decode(le05," & stYrMn & ",le03+nvl(le04,0)-le07*(12*(substr(le06,1,4)-substr(le05,1,4))+(substr(le06,5)-substr(le05,5))),le07),'999,999') 本月償還金額" & _
         " from Loan_Employee where le05<=" & stYrMn & " and le06>=" & stYrMn & " and le01='" & cboUser.Tag & "' order by 1"
   Case 21 '20 '借支
      strExc(1) = strExc(1) & "員工借支還款明細"
      strExc(0) = "select sqldatet(ae02) 借支日期,to_char(ae03,'999,999') 借支金額 from Advance_Employee where ae01='" & cboUser.Tag & "' and ae04=" & stYrMn & " order by 1"
   Case 23 '22 '婚喪喜慶
      strExc(1) = strExc(1) & "婚喪喜慶明細"
      'modify by sonia 2022/10/4 2022/9起薪資改為顯示發生日及扣款日20220718 A2027,20220818 81009,20221001 A8008
      'strExc(0) = "select sqldatet(WFA01) 日期,st02 當事人,to_char(wfa04,'999,999') 金額 From WFAmount,staff where substr(wfa01,1,6)='" & stYrMn & "' and wfa03='" & cboUser.Tag & "' and st01(+)=wfa02 and wfa05='1' order by 1"
      'Modified by Morgan 2022/10/7
      'If stYrMn < 202209 Then
      '   strExc(0) = "select sqldatet(WFA01) 日期,st02 當事人,to_char(wfa04,'999,999') 金額 From WFAmount,staff where substr(wfa01,1,6)='" & stYrMn & "' and wfa03='" & cboUser.Tag & "' and st01(+)=wfa02 and wfa05='1' order by 1"
      'Else
      '   strExc(0) = "select sqldatet(WFA01) 扣款日期,sqldatet(WF01) 發生日期,st02 當事人,to_char(wfa04,'999,999') 金額 From WFAmount,staff,WEDDINGANDFUNERAL where substr(wfa01,1,6)='" & stYrMn & "' and wfa03='" & cboUser.Tag & "' and st01(+)=wfa02 and wfa05='1' and wfa01=wf04(+) and wfa02=wf02(+) order by 1"
      'End If
      'Modified by Morgan 2022/10/31 +原因
      strExc(0) = "select sqldatet(WF04) 扣款日期,sqldatet(WF01) 發生日期,st02 當事人,to_char(wfa04,'999,999') 金額,decode(wf03,'1','婚','2','喪','')||decode(wf11,'1','(父親)','2','(母親)','3','(配偶)','4','(兒子)','5','(女兒)','') 原因 From WEDDINGANDFUNERAL,WFAmount,staff where substr(wf04,1,6)='" & stYrMn & "' and wfa01(+)=wf01 and wfa02(+)=wf02 and wfa03='" & cboUser.Tag & "' and wfa05='1' and st01(+)=wfa02 order by 1"
      'end 2022/10/7
      'end 2022/10/4
   Case 24 '23 '其它扣款
      strExc(1) = strExc(1) & "其它扣款明細"
      'modify by sonia 2016/9/5 補扣健保費時,備註欄加眷屬姓名
      'strExc(0) = "select sqldatet(od02) 日期,oc03 項目,to_char(od05,'999,999') 金額,od15 備註 From OtherSalaryData, OtherSalaryCode where od03='" & cboUser.Tag & "' and substr(od02,1,6)='" & stYrMn & "' and oc02='D' and od04<>'01' and od04<>'12' and oc01(+)=od04 order by 1"
      strExc(0) = "select sqldatet(od02) 日期,oc03 項目,to_char(od05,'999,999') 金額,od15||decode(od04,'35',decode(od16,'0',null,decode(od15,null,null,'-')||sr04),null) 備註 From OtherSalaryData, OtherSalaryCode, staff_relation where od03='" & cboUser.Tag & "' and substr(od02,1,6)='" & stYrMn & "' and oc02='D' and od04<>'01' and od04<>'12' and oc01(+)=od04 and od03=sr01(+) and od16=sr02(+) order by 1"
   Case 25 '24 '缺勤扣薪
      strExc(1) = strExc(1) & "缺勤明細"
      'modify by sonia 2018/9/5 +曠職扣薪 79017 20180821 1.5時
      'strExc(0) = "select sqldatet(sa02)||' '||trunc(sa03/100)||':'||substr(sa03,-2)||' - '||sqldatet(sa04)||' '||trunc(sa05/100)||':'||substr(sa05,-2) 缺勤起迄,ac03 假別,sa07 天數,sa08 時數,0 金額,sa06,sa02 from Staff_Absence,allcode where sa01='" & cboUser.Tag & "' and substr(sa02,1,6)='" & stYrMn & "' and sa06 in ('05','06','20','22') and ac01(+)='04' and ac02(+)=sa06 order by 1"
      'Modified by Morgan 2020/2/5 曠職加 and (sa05>0 or sa06>0) 條件，否則會抓到忘打卡及遲到
      'Modified by Morgan 2021/10/5 +24防疫照顧假
      'Modified by Morgan 2023/6/7 +可能隔月扣款 substr(sa02,1,6)='" & stYrMn & "'-->sa18=" & stYrMn & "
      'Modified by Morgan 2025/8/15 修改曠職時間改以「分」計算
      If strSrvDate(1) >= 曠職以分計啟用日 Then
         strExc(0) = "select sqldatet(sa02)||' '||trunc(sa03/100)||':'||substr(sa03,-2)||' - '||sqldatet(sa04)||' '||trunc(sa05/100)||':'||substr(sa05,-2) 缺勤起迄,ac03 假別,sa07 天數,sa08 時數,0 金額,sa06,sa02 from Staff_Absence,allcode where sa01='" & cboUser.Tag & "' and sa18=" & stYrMn & " and sa06 in ('05','06','20','22','24','25') and ac01(+)='04' and ac02(+)=sa06 " & _
                  "union select sqldatet(sa02),'曠職',sa05,round(sa06/60,4) sa06,0,'03',sa02 from Staff_Assist where sa01='" & cboUser.Tag & "' and substr(sa02,1,6)='" & stYrMn & "' and (sa05>0 or sa06>0) order by 1"
      Else
         strExc(0) = "select sqldatet(sa02)||' '||trunc(sa03/100)||':'||substr(sa03,-2)||' - '||sqldatet(sa04)||' '||trunc(sa05/100)||':'||substr(sa05,-2) 缺勤起迄,ac03 假別,sa07 天數,sa08 時數,0 金額,sa06,sa02 from Staff_Absence,allcode where sa01='" & cboUser.Tag & "' and sa18=" & stYrMn & " and sa06 in ('05','06','20','22','24','25') and ac01(+)='04' and ac02(+)=sa06 " & _
                  "union select sqldatet(sa02),'曠職',sa05,sa06,0,'03',sa02 from Staff_Assist where sa01='" & cboUser.Tag & "' and substr(sa02,1,6)='" & stYrMn & "' and (sa05>0 or sa06>0) order by 1"
      End If
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
      .ColAlignmentFixed = flexAlignCenterCenter
      Select Case pRow
      Case 6 '加班費
         .ColWidth(0) = 900
         .ColAlignment(0) = flexAlignCenterCenter
         .ColWidth(1) = 700
         .ColAlignment(1) = flexAlignRightCenter
         'Modified by Morgan 2017/7/11
         'iCol = 1
         .ColWidth(2) = 1050
         .ColAlignment(2) = flexAlignRightCenter
         .ColWidth(3) = 1250
         .ColAlignment(3) = flexAlignRightCenter
         .ColWidth(4) = 1250
         .ColAlignment(4) = flexAlignRightCenter
         iCol = 4
         
         '加班費明細
         .row = 1
         iDailyHours = GetDaiyHour(cboUser.Tag, stYrMn)
         'dblHourPay = GetHourPay(cboUser.Tag, stYrMn, iDailyHours) 'Removed by Morgan 2019/6/6
         For iRow = 1 To .Rows - 1
            'Modified by Morgan 2023/6/7
            'dblHourPay = GetHourPay(cboUser.Tag, .TextMatrix(iRow, 6), iDailyHours) 'Added by Morgan 2019/6/6 時薪改以日期判斷
            If stYrMn < "201905" Then
               dblHourPay = GetHourPay(cboUser.Tag, .TextMatrix(iRow, 6), iDailyHours) 'Added by Morgan 2019/6/6 時薪改以日期判斷
            Else
               dblHourPay = GetHourPay(cboUser.Tag, Left(.TextMatrix(iRow, 6), 6), iDailyHours)
            End If
            
            dblHours = Val("" & .TextMatrix(iRow, 8)) 'so06
            dblAmt = 0
            iOtType = 1
            '假日
            If dblHours > 0 Then
               '休息日(周六):1.要累計46小時, 2.加班費計算:前2小時 x4/3, 第3-8小時 x5/3, 第9-12小時 x8/3 (未滿4小時以4小時計,以此類推)
               If .TextMatrix(iRow, 6) > "20161223" And Weekday(Format(.TextMatrix(iRow, 6), "@@@@/@@/@@")) = 7 Then
                  iOtType = 3
                  dblAmt = (4 / 3) * 2 * dblHourPay
                  If dblHours > 8 Then
                     dblAmt = dblAmt + (5 / 3) * 6 * dblHourPay
                     dblAmt = dblAmt + (8 / 3) * (dblHours - 8) * dblHourPay
                  Else
                     dblAmt = dblAmt + (5 / 3) * (dblHours - 2) * dblHourPay
                  End If
                  dblHours = 0
               '休假日:前8小時 x1,8小時以後的加班時數以平日加班費計算
               Else
                  iOtType = 2
                  dblAmt = 8 * dblHourPay
                  dblHours = dblHours - 8
               End If
            End If
            
            '平日(兩小時以內*4/3,超過兩小時部分*5/3)
            dblHours = dblHours + Val("" & .TextMatrix(iRow, 7)) '休假日超過8小時的部分比照平日的算法
            If dblHours > 0 Then
               '兩小時以內
               If dblHours <= 2 Then
                  dblAmt = dblAmt + (4 / 3) * dblHours * dblHourPay
               Else
                   dblAmt = dblAmt + ((4 / 3) * 2 + (5 / 3) * (dblHours - 2)) * dblHourPay
               End If
            End If
                        
            If iOtType = 2 Then
               .TextMatrix(iRow, 3) = -1 * Int(-1 * dblAmt)
            ElseIf iOtType = 3 Then
               .TextMatrix(iRow, 4) = -1 * Int(-1 * dblAmt)
            Else
               .TextMatrix(iRow, 2) = -1 * Int(-1 * dblAmt)
            End If
         Next
         'end 2017/7/11
      Case 12 '11 '其它所得
         .ColWidth(0) = 900
         .ColAlignment(0) = flexAlignCenterCenter
         .ColWidth(1) = 1500
         .ColAlignment(1) = flexAlignLeftCenter
         .ColWidth(2) = 900
         .ColAlignment(2) = flexAlignRightCenter
         .ColWidth(3) = 800
         .ColAlignment(3) = flexAlignRightCenter
         .ColWidth(4) = 2500
         .ColAlignment(4) = flexAlignLeftCenter
         iCol = 4
      Case 15 '14 '健保費
         .ColWidth(0) = 900
         .ColAlignment(0) = flexAlignCenterCenter
         .ColWidth(1) = 2050
         .ColAlignment(1) = flexAlignLeftCenter
         .ColWidth(2) = 900
         .ColAlignment(2) = flexAlignRightCenter
         iCol = 2
      Case 16 '15 '補充保費
         .ColWidth(0) = 900
         .ColAlignment(0) = flexAlignCenterCenter
         .ColWidth(1) = 3050
         .ColAlignment(1) = flexAlignLeftCenter
         .ColWidth(2) = 900
         .ColAlignment(2) = flexAlignRightCenter
         .ColWidth(3) = 900
         .ColAlignment(3) = flexAlignRightCenter
         iCol = 3
      Case 19 '18 '互助會會款
         .ColWidth(0) = 900
         .ColAlignment(0) = flexAlignCenterCenter
         .ColWidth(1) = 900
         .ColAlignment(1) = flexAlignCenterCenter
         .ColWidth(2) = 900
         .ColAlignment(2) = flexAlignRightCenter
         iCol = 2
      Case 20 '19 '員工貸款
         .ColWidth(0) = 900
         .ColAlignment(0) = flexAlignCenterCenter
         .ColWidth(1) = 900
         .ColAlignment(1) = flexAlignRightCenter
         .ColWidth(2) = 900
         .ColAlignment(2) = flexAlignRightCenter
         .ColWidth(3) = 1600
         .ColAlignment(3) = flexAlignCenterCenter
         .ColWidth(4) = 1600
         .ColAlignment(4) = flexAlignRightCenter
         iCol = 4
      Case 21 '20 '借支
         .ColWidth(0) = 900
         .ColAlignment(0) = flexAlignCenterCenter
         .ColWidth(1) = 900
         .ColAlignment(1) = flexAlignRightCenter
         iCol = 1
      Case 23 '22 '婚喪喜慶
         .ColWidth(0) = 900
         .ColAlignment(0) = flexAlignCenterCenter
         .ColWidth(1) = 900
         .ColAlignment(1) = flexAlignCenterCenter
         .ColWidth(2) = 900
         .ColAlignment(2) = flexAlignRightCenter
         'modify by sonia 2022/10/4 2022/9起薪資加扣款日期欄
         'iCol = 2
         'Modified by Morgan 2022/10/31
         'If stYrMn < 202209 Then
         '   iCol = 2
         'Else
         '   .ColWidth(3) = 900
         '   .ColAlignment(3) = flexAlignRightCenter
         '   iCol = 3
         'End If
         .ColWidth(3) = 900
         .ColAlignment(3) = flexAlignRightCenter
         .ColWidth(4) = 900
         .ColAlignment(4) = flexAlignLeftCenter
         iCol = 4
         'end 2022/10/31
         
         'end 2022/10/4
      Case 24 '23 '其它扣款
         .ColWidth(0) = 900
         .ColAlignment(0) = flexAlignCenterCenter
         .ColWidth(1) = 1500
         .ColAlignment(1) = flexAlignLeftCenter
         .ColWidth(2) = 900
         .ColAlignment(2) = flexAlignRightCenter
         .ColWidth(3) = 2500
         .ColAlignment(3) = flexAlignLeftCenter
         iCol = 3
      Case 25 '24 '缺勤扣薪
         .ColWidth(0) = 2700
         .ColAlignment(0) = flexAlignCenterCenter
         .ColWidth(1) = 1200
         .ColAlignment(1) = flexAlignCenterCenter
         .ColWidth(2) = 900
         .ColAlignment(2) = flexAlignRightCenter
         .ColWidth(3) = 900
         .ColAlignment(3) = flexAlignRightCenter
         'Modified by Morgan 2017/7/11
         'iCol = 3
         .ColWidth(4) = 900
         .ColAlignment(4) = flexAlignRightCenter
         iCol = 4
         '金額
         .row = 1
         iDailyHours = GetDaiyHour(cboUser.Tag, stYrMn)
         
         'Modified by Morgan 2019/7/2 時薪改以日期判斷--未完成
         'dblHourPay = GetHourPay(cboUser.Tag, stYrMn, iDailyHours)
         'Removed by Morgan 2023/6/7 可能隔月扣款,改移到下面
         'If stYrMn < "10905" Then
         '   dblHourPay = GetHourPay(cboUser.Tag, stYrMn, iDailyHours)
         'Else
         '   dblHourPay = GetHourPay(cboUser.Tag, stYrMn & "31", iDailyHours)
         'End If
         'end 2023/6/7
         'end 2019/7/2
         For iRow = 1 To .Rows - 1
            'Added by Morgan 2023/6/7
            If stYrMn < "201905" Then
               dblHourPay = GetHourPay(cboUser.Tag, .TextMatrix(iRow, 6), iDailyHours)
            Else
               dblHourPay = GetHourPay(cboUser.Tag, Left(.TextMatrix(iRow, 6), 6), iDailyHours)
            End If
            'end 2023/6/7
            dblDHours = 0
            'modify by sonia 2018/9/5 +曠職03扣薪 79017 20180821 1.5時
            'Modified by Morgan 2021/10/5 +24防疫照顧假
            '事假
            If .TextMatrix(iRow, 5) = "05" Or .TextMatrix(iRow, 5) = "22" Or .TextMatrix(iRow, 5) = "24" Or .TextMatrix(iRow, 5) = "03" Or .TextMatrix(iRow, 5) = "25" Then
               dblDHours = (iDailyHours * Val(.TextMatrix(iRow, 2)) + Val(.TextMatrix(iRow, 3)))
            '病假,生理假
            Else
               '已請病假
               dblSickHour = GetSickHour(cboUser.Tag, Val(Left(stYrMn, 4) & "0101"), Val(.TextMatrix(iRow, 6)) - 1)
               '已請生理假
               dblGirlSickHour = GetSickHour(cboUser.Tag, Val(Left(stYrMn, 4) & "0101"), Val(.TextMatrix(iRow, 6)) - 1, True)
               '生理假超過3天部分併入病假
               If dblGirlSickHour > dblGirlSickBaseHour Then
                  dblSickHour = dblSickHour + (dblGirlSickHour - dblGirlSickBaseHour)
               End If
               
               '本次請假時數
               dblHours = (iDailyHours * Val(.TextMatrix(iRow, 2)) + Val(.TextMatrix(iRow, 3)))
               '病假扣半薪時數
               dblSickBaseHour = 30 * iDailyHours
               '生理假扣半薪時數
               dblGirlSickBaseHour = 3 * iDailyHours
               
               '生理假:前3天扣半薪,超過3天部分合併病假,超過30天扣全薪,未超過30天扣半薪
               If .TextMatrix(iRow, 5) = "20" Then
                  '累計未超過3天
                  If dblGirlSickHour < dblGirlSickBaseHour Then
                     '累計未超過3天部分扣半薪,超過部分要併入病假考慮
                     If dblGirlSickHour + dblHours <= dblGirlSickBaseHour Then
                        dblDHours = dblDHours + 0.5 * dblHours
                        dblHours = 0
                     Else
                        dblDHours = dblDHours + 0.5 * (dblGirlSickBaseHour - dblGirlSickHour)
                        dblHours = dblHours - (dblGirlSickBaseHour - dblGirlSickHour)
                     End If
                  End If
               End If
               
               '超過3天後的生理假合併病假計算,超過30天扣全薪,未超過30天扣半薪
               If dblHours > 0 Then
                  If dblSickHour >= dblSickBaseHour Then
                     dblDHours = dblDHours + dblHours
                  ElseIf dblSickHour + dblHours <= dblSickBaseHour Then
                     dblDHours = dblDHours + 0.5 * dblHours
                  Else
                     dblDHours = dblDHours + 0.5 * (dblSickBaseHour - dblSickHour) + (dblHours - (dblSickBaseHour - dblSickHour))
                  End If
               End If
               
            End If
            dblAmt = dblHourPay * dblDHours
            'Modified by Morgan 2024/8/2
            '.TextMatrix(iRow, 4) = Round(dblAmt)
            .TextMatrix(iRow, 4) = Trunc(dblAmt)
            'end 2024/8/2
            
            'Added by Morgan 2025/9/4
            If Val(.TextMatrix(iRow, 3)) < 1 Then
               .TextMatrix(iRow, 3) = Format(.TextMatrix(iRow, 3), "0.00")
            End If
            'end 2025/9/4
         Next
         'end 2017/7/11
      End Select
      For intI = iCol + 1 To .Cols - 1
         .ColWidth(intI) = 0
      Next
      .MergeCol(0) = True
      .MergeCells = flexMergeFree
      End With
      .Caption = strExc(1)
      .lblRemark.Caption = .lblRemark.Caption & "缺勤扣薪是以總時數計算可能會與明細金額加總有小數進位上的誤差！"
      .lblRemark.Visible = True
      .Show vbModal
      End With
   Else
      MsgBox "無明細資料！", vbInformation, strExc(1)
   End If
End Sub
