VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210111 
   BorderStyle     =   1  '單線固定
   Caption         =   "新客戶來源分析"
   ClientHeight    =   5640
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   9380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   9380
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Height          =   400
      Left            =   6585
      Style           =   1  '圖片外觀
      TabIndex        =   21
      Top             =   105
      Width           =   800
   End
   Begin VB.OptionButton opt1 
      Caption         =   "圓餅圖"
      Height          =   210
      Index           =   1
      Left            =   6240
      TabIndex        =   20
      Top             =   1455
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.OptionButton opt1 
      Caption         =   "長條圖"
      Enabled         =   0   'False
      Height          =   210
      Index           =   0
      Left            =   6225
      TabIndex        =   19
      Top             =   1215
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "統計圖表(&G)"
      Height          =   375
      Left            =   7725
      TabIndex        =   18
      Top             =   1275
      Width           =   1410
   End
   Begin VB.TextBox txtCuNa 
      Height          =   285
      Index           =   1
      Left            =   2400
      MaxLength       =   3
      TabIndex        =   7
      Top             =   1410
      Width           =   915
   End
   Begin VB.TextBox txtCuNa 
      Height          =   285
      Index           =   0
      Left            =   1185
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1410
      Width           =   915
   End
   Begin VB.TextBox txtSalesArea1 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Top             =   420
      Width           =   915
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7515
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8415
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   90
      Width           =   800
   End
   Begin VB.TextBox txtZone 
      Height          =   285
      Left            =   1185
      MaxLength       =   1
      TabIndex        =   0
      Top             =   90
      Width           =   915
   End
   Begin VB.TextBox txtSales 
      Height          =   285
      Left            =   1185
      MaxLength       =   6
      TabIndex        =   3
      Top             =   750
      Width           =   915
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   285
      Index           =   1
      Left            =   2400
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1080
      Width           =   915
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   285
      Index           =   0
      Left            =   1185
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1080
      Width           =   915
   End
   Begin VB.TextBox txtSalesArea 
      Height          =   285
      Left            =   1185
      TabIndex        =   1
      Top             =   420
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3810
      Left            =   60
      TabIndex        =   10
      Top             =   1785
      Width           =   9285
      _ExtentX        =   16387
      _ExtentY        =   6703
      _Version        =   393216
      BackColor       =   -2147483624
      Rows            =   3
      Cols            =   7
      FixedRows       =   2
      FixedCols       =   4
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
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
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.ComboBox Combo3 
      Height          =   336
      Left            =   1188
      TabIndex        =   23
      Top             =   720
      Width           =   1920
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3387;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblSalesName 
      Height          =   300
      Left            =   2160
      TabIndex        =   22
      Top             =   750
      Width           =   4590
      VariousPropertyBits=   27
      Size            =   "8096;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line3 
      X1              =   2115
      X2              =   2385
      Y1              =   1545
      Y2              =   1545
   End
   Begin VB.Label Label5 
      Caption         =   "客戶國籍："
      Height          =   180
      Left            =   135
      TabIndex        =   17
      Top             =   1455
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   2130
      X2              =   2400
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   105
      TabIndex        =   16
      Top             =   5430
      Width           =   45
   End
   Begin VB.Label lblZone 
      AutoSize        =   -1  'True
      Caption         =   "（1北所 2中所 3南所 4高所）"
      Height          =   180
      Left            =   2220
      TabIndex        =   15
      Top             =   135
      Width           =   2295
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   150
      TabIndex        =   14
      Top             =   795
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   2130
      X2              =   2400
      Y1              =   1215
      Y2              =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "所別："
      Height          =   180
      Left            =   150
      TabIndex        =   13
      Top             =   135
      Width           =   540
   End
   Begin VB.Label Label2 
      Caption         =   "開發日期："
      Height          =   180
      Left            =   150
      TabIndex        =   12
      Top             =   1125
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "業務區："
      Height          =   180
      Left            =   150
      TabIndex        =   11
      Top             =   465
      Width           =   720
   End
End
Attribute VB_Name = "frm210111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/04 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、lblSalesName ; Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'2005/7/5整理
Option Explicit

Dim i As Integer
Dim j As Integer
Dim iPrint As Integer, Page As Integer, strTemp(0 To 12) As String
Dim PLeft(0 To 12) As Integer
Public Calrs As New ADODB.Recordset
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
Dim m_strListPer As String 'Add By Sindy 2020/7/1
'Add By Sindy 2023/6/12
Dim arrID, stST05 As String ', stST15 As String
'Dim bolAreaMan As Boolean '下拉選單有區主管
'2023/6/12 END


Private Sub SetDataListWidth()
Dim arrGridHeadText, arrGridHeadWidth, arrGridHeadText2

Dim iRow As Integer

arrGridHeadText = Array("", "", "", "" _
                  , "", "個人總", "", "該區", "", "全所業務", "區佔全所" _
                  , "全所總", "區佔")
arrGridHeadText2 = Array("所別", "業務區", "智權人員", "客戶來源" _
                  , "客戶數", "客戶數", "個人比例", "客戶數", "該區比例", "客戶數", "業務比例" _
                  , "客戶數", "全所比例")
                     
arrGridHeadWidth = Array(500, 700, 700, 900 _
                     , 700, 700, 850, 700, 850, 850, 850 _
                     , 700, 850)
If Me.txtZone.Locked = True Or Me.txtZone.Enabled = False Then
   'arrGridHeadWidth(0) = 0
End If
If (Me.txtSalesArea.Locked = True And Me.txtSalesArea1.Locked = True) Or (Me.txtSalesArea.Enabled = False And Me.txtSalesArea1.Enabled = False) Then
   'arrGridHeadWidth(1) = 0
'   arrGridHeadWidth(7) = 0
'   arrGridHeadWidth(8) = 0
End If
If Me.txtSales.Locked = True Or Me.txtSales.Enabled = False Then
   'arrGridHeadWidth(2) = 0
   arrGridHeadWidth(9) = 0
   arrGridHeadWidth(10) = 0
   arrGridHeadWidth(11) = 0
   arrGridHeadWidth(12) = 0
End If
                     
grdDataList.Cols = UBound(arrGridHeadText) + 1
For iRow = 0 To grdDataList.Cols - 1
   grdDataList.row = 0
   grdDataList.col = iRow
   grdDataList.Text = arrGridHeadText(iRow)
   grdDataList.ColWidth(iRow) = arrGridHeadWidth(iRow)
   grdDataList.row = 1
   grdDataList.col = iRow
   grdDataList.Text = arrGridHeadText2(iRow)
   'grdDataList.ColWidth(iRow) = arrGridHeadWidth(iRow)
   If iRow <= 3 Then
      grdDataList.CellAlignment = flexAlignLeftCenter
      grdDataList.ColAlignment(iRow) = 1
   Else
      grdDataList.CellAlignment = flexAlignRightCenter
      grdDataList.ColAlignment(iRow) = 7
   End If
Next
'add by nickc 2005/12/28 副總說要合併
grdDataList.MergeCells = flexMergeRestrictColumns
grdDataList.MergeCol(0) = True
grdDataList.MergeCol(1) = True
grdDataList.MergeCol(2) = True
grdDataList.MergeCol(5) = True
grdDataList.MergeCol(6) = True
grdDataList.MergeCol(7) = True
grdDataList.MergeCol(8) = True
grdDataList.MergeCol(9) = True
grdDataList.MergeCol(10) = True
grdDataList.MergeCol(11) = True
grdDataList.MergeCol(12) = True
End Sub

Private Function ConstrainCheck() As Boolean
Dim bolCancel As Boolean
Dim intErrCol As Integer

   ConstrainCheck = True
   
   'Add By Sindy 2023/6/12
   If Combo3.Visible = True Then
      bolCancel = False
      Call Combo3_Validate(bolCancel)
      If bolCancel = True Then
         ConstrainCheck = False
         Exit Function
      End If
   End If
   'Add by Amy 2020/03/25 +有下拉選單
   If Combo3.Visible = True Then
      Call Combo3_LostFocus 'Add By Sindy 2020/7/15 讓人員按Enter,須再啟動此函數,txtSales欄位值才會置換
      If Combo3 = MsgText(601) Then
          Call Combo3_Validate(bolCancel)
          If bolCancel = True Then
              Combo3.SetFocus
              ConstrainCheck = False
              Exit Function
          End If
      ElseIf txtSales = MsgText(601) Then
          txtSales = Mid(Combo3, 1, Val(InStr(Combo3, " ")) - 1)
      End If
   End If
   Call txtSales_Validate(bolCancel)
   If bolCancel = True Then
      'Modify by Amy 2020/03/25 +有下拉選單
      If Combo3.Visible = True Then
         Combo3.SetFocus
      'Modified by Lydia 2021/05/20 排除隱藏
      'ElseIf txtSales.Enabled = True Then
      ElseIf txtSales.Enabled = True And txtSales.Visible = True Then
         txtSales.SetFocus
         txtSales_GotFocus
      End If
      ConstrainCheck = False
      Exit Function
   End If
   '2023/6/12 END
'   'Add By Sindy 2009/05/14
'   Call txtSales_Validate(bolCancel)
'   If bolCancel = True Then
'      txtSales.SetFocus
'      txtSales_GotFocus
'      ConstrainCheck = False
'      Exit Function
'   End If
   
   If txtCloseDate(0) = "" Then
      MsgBox "請輸入開發日期起日！", vbExclamation
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
      MsgBox "請輸入開發日期迄日！", vbExclamation
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
   
   'Modify By Sindy 2020/7/29 檢查部門欄位
   'Modify By Sindy 2025/8/11 +, txtZone
   If PUB_ChkFormSalesDept(strUserNum, txtSales, txtSalesArea, txtSalesArea1, intErrCol, txtZone) = False Then
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
   
'   '2005/7/5 ADD BY SONIA 林永生71003檢查業務區範圍
'   If strUserNum = "71003" Then
'      If txtSalesArea < "S2" Or txtSalesArea > "S29" Then
'         MsgBox "業務區起始條件錯誤！只可查中所業務區", vbExclamation
'         txtSalesArea.SetFocus
'         txtSalesArea_GotFocus
'         ConstrainCheck = False
'         Exit Function
'      End If
'      If txtSalesArea1 < "S2" Or txtSalesArea1 > "S29" Then
'         MsgBox "業務區迄止條件錯誤！只可查中所業務區", vbExclamation
'         txtSalesArea1.SetFocus
'         txtSalesArea1_GotFocus
'         ConstrainCheck = False
'         Exit Function
'      End If
'   End If
'   '2006/11/29 ADD BY SONIA 簡金泉檢查業務區範圍
''Removed by Morgan 2019/12/30 杜主秘說加開簡協理69005可看全所
''   If strUserNum = "69005" Then
''      If txtSalesArea < "S1" Or txtSalesArea > "S19" Then
''         MsgBox "業務區起始條件錯誤！只可查北所業務區", vbExclamation
''         txtSalesArea.SetFocus
''         txtSalesArea_GotFocus
''         ConstrainCheck = False
''         Exit Function
''      End If
''      If txtSalesArea1 < "S1" Or txtSalesArea1 > "S19" Then
''         MsgBox "業務區迄止條件錯誤！只可查北所業務區", vbExclamation
''         txtSalesArea1.SetFocus
''         txtSalesArea1_GotFocus
''         ConstrainCheck = False
''         Exit Function
''      End If
''   End If
''end 2019/12/30
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
'
'   If txtSalesArea > txtSalesArea1 Then
'      MsgBox "業務區範圍條件錯誤！", vbExclamation
'      txtSalesArea.SetFocus
'      txtSalesArea_GotFocus
'      ConstrainCheck = False
'      Exit Function
'   End If
End Function

Public Function doQuery() As Boolean
   Dim stCon As String, stConST As String, stConST1 As String, stConST2 As String, stConST3 As String
   Dim stConAA As String, stConBB As String, stConCC As String, stConDD As String, stConEE As String
   Dim stVTBx As String, stVTBy As String, stVTBz As String
   
   stCon = "": stConST = " and cu13 is not null ": stConAA = "": stConBB = "": stConCC = "": stConDD = "": stConEE = ""
   stConST1 = " and cu13 is not null ": stConST2 = " and cu13 is not null ": stConST3 = " and cu13 is not null "
   'add by nickc 2005/09/26 只算新客戶
   stConST = stConST & " and cu02='0' "
   stConST1 = stConST1 & " and cu02='0' "
   stConST2 = stConST2 & " and cu02='0' "
   stConST3 = stConST3 & " and cu02='0' "
   
   '所別
   If txtZone <> "" Then
      stConST = stConST & " and st06 = '" & txtZone & "' "
      pub_QL05 = pub_QL05 & ";" & Label3 & txtZone & lblZone 'Add By Sindy 2010/12/23
   End If
   
   '區別
   'Modify By Sindy 2009/05/12
   '若為帶人主管權限時,查詢之智權人員編號非本人時,不限制區別
   If (Trim(txtSales) <> "" And PUB_GetST05(strUserNum) = "SA" And txtSales.Enabled = True And txtSales <> strUserNum) Then
      '不限制區別
   Else
      If txtSalesArea <> "" Then
         stConST = stConST & " and cu12||''>='" & txtSalesArea & "' "
      End If
      If txtSalesArea1 <> "" Then
         stConST = stConST & " and cu12||''<='" & txtSalesArea1 & "' "
      End If
      If txtSalesArea <> "" Or txtSalesArea1 <> "" Then
         pub_QL05 = pub_QL05 & ";" & Label1 & txtSalesArea & "-" & txtSalesArea1 'Add By Sindy 2010/12/23
      End If
   End If
   
   '智權人員
   If txtSales <> "" Then
      'edit by nickc 2008/04/25 秀玲說要改共用控制
      'stConST = stConST & " and cu13||'' = '" & txtSales & "' "
      stConST = stConST & " and cu13||'' in (" & PUB_GetSalesList(txtSales, txtSalesArea, txtSalesArea1, txtZone) & ")"
      pub_QL05 = pub_QL05 & ";" & Label4 & txtSales & lblSalesName 'Add By Sindy 2010/12/23
   End If

   '開發日期
   If txtCloseDate(0) <> "" Then
      stConST = stConST & " and cu14>=" & ChangeTStringToWString(txtCloseDate(0))
      stConST1 = stConST1 & " and cu14>=" & ChangeTStringToWString(txtCloseDate(0))
   End If
   If txtCloseDate(1) <> "" Then
      stConST = stConST & " and cu14<=" & ChangeTStringToWString(txtCloseDate(1))
      stConST1 = stConST1 & " and cu14<=" & ChangeTStringToWString(txtCloseDate(1))
   End If
   If txtCloseDate(0) <> "" Or txtCloseDate(1) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label2 & txtCloseDate(0) & "-" & txtCloseDate(1) 'Add By Sindy 2010/12/23
   End If
   
   '客戶國籍
   If txtCuNa(0) <> "" Then
      stConST = stConST & " and cu10||''>='" & txtCuNa(0) & "' "
      stConST1 = stConST1 & " and cu10||''>='" & txtCuNa(0) & "' "
   End If
   If txtCuNa(1) <> "" Then
      stConST = stConST & " and cu10||''<='" & txtCuNa(1) & "' "
      stConST1 = stConST1 & " and cu10||''<='" & txtCuNa(1) & "' "
   End If
   If txtCuNa(0) <> "" Or txtCuNa(1) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label5 & txtCuNa(0) & "-" & txtCuNa(1) 'Add By Sindy 2010/12/23
   End If
   
On Error GoTo ErrHnd
   
   '個人單項件數
   stConAA = "select st06 as 所別,st15 as 業務區,cu13 as 智權人員,decode(cu09,null,'99',cu09) as 客戶來源,count(cu01||cu02) as 個人件數  from customer,staff " & _
                     " where cu13=st01(+) " & stConST & " group by st06,st15,cu13,cu09 "
   '各區總件數
   stConBB = "select st06 as 所別,st15 as 業務區,count(cu01||cu02) as 區總件數  from customer,staff " & _
                     " where cu13=st01(+) " & stConST1 & " group by st06,st15 "
   '各所總件數
'   stConCC = "select st06 as 所別,count(cu01||cu02) as 所總件數  from customer,staff " & _
                     " where cu13=st01(+) " & stConST1 & " group by st06 "
   '全所業務總件數
   stConCC = "select count(cu01||cu02) as 全所業務總件數  from customer,staff " & _
                     " where cu13=st01(+) and cu12||''>='S11' and cu12||''<='S99' " & stConST1 & " "
   
   '全所總件數
   stConDD = "select count(cu01||cu02) as 全所總件數  from customer,staff " & _
                      " where cu13=st01(+) " & stConST1 & "  "
   '個人總件數
   stConEE = "select st06 as 所別,st15 as 業務區,cu13 as 智權人員,count(cu01||cu02) as 個人總件數  from customer,staff " & _
                    " where cu13=st01(+) " & stConST1 & " group by st06,st15,cu13 "
                     
   stCon = "select decode(ZZ.所別,'1','北所','2','中所','3','南所','4','高所','5','其他','') as 所別,a0902 as 業務區,decode(ZZ.智權人員,'ZA','區統計','ZB','該所統計','ZC','全所統計','ZD','全所合計',st02) as 智權人員, " & _
                " decode(ZZ.客戶來源,'ALL','所有',csm02) as 客戶來源,ZZ.個人件數,ZZ.個人總件數,ZZ.個人比例, " & _
                " ZZ.區總件數,ZZ.該區總比例, " & _
                " ZZ.全所業務總件數,ZZ.全所業務比例, " & _
                " ZZ.全所總件數,ZZ.全所總比例 " & _
                " from (select AA.所別,AA.業務區,AA.智權人員, " & _
                " AA.客戶來源,AA.個人件數,EE.個人總件數,decode(to_char(round(AA.個人件數/EE.個人總件數*100,2)),null,'',to_char(round(AA.個人件數/EE.個人總件數*100,2),9999.99)||' %') as 個人比例, " & _
                " BB.區總件數,decode(to_char(round(AA.個人件數/BB.區總件數*100,2)),null,'',to_char(round(AA.個人件數/BB.區總件數*100,2),9999.99)||' %') as 該區總比例, " & _
                " CC.全所業務總件數,decode(to_char(round(BB.區總件數/CC.全所業務總件數*100,2)),null,'',to_char(round(BB.區總件數/CC.全所業務總件數*100,2),9999.99)||' %') as 全所業務比例, " & _
                " DD.全所總件數,decode(to_char(round(BB.區總件數/DD.全所總件數*100,2)),null,'',to_char(round(BB.區總件數/DD.全所總件數*100,2),9999.99)||' %') as 全所總比例 " & _
                " from (" & stConAA & ") AA,(" & stConBB & ") BB,(" & stConCC & ") CC,(" & stConDD & ") DD,(" & stConEE & ") EE " & _
                " where AA.所別=BB.所別(+) " & _
                " and AA.業務區=BB.業務區(+)   " & _
                " and AA.所別=EE.所別(+) and AA.業務區=EE.業務區(+) and AA.智權人員=EE.智權人員(+) "
     stCon = stCon & " ) ZZ,staff,acc090,casesourcemap " & _
                " where ZZ.業務區=a0901(+) and ZZ.智權人員=st01(+) and ZZ.客戶來源=csm01(+) " & _
                " order by zz.所別,zz.業務區,zz.智權人員,zz.客戶來源 "

   SetDataListWidth
   grdDataList.Rows = 3
   grdDataList.Clear
   SetDataListWidth
   Set Calrs = New ADODB.Recordset
   With Calrs
      If .State = 1 Then .Close
      .CursorLocation = adUseClient
      .Open stCon, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/23
         'edit by nickc 2005/09/23 因為秀玲說前面要固定，所以就只能跑回圈塞
         'Set grdDataList.Recordset = AdoRecordSet3.Clone
         For i = 1 To .RecordCount
            .MoveFirst
            .Move i - 1
            For j = 1 To .Fields.Count
                grdDataList.TextMatrix(i + 1, j - 1) = CheckStr(.Fields(j - 1))
            Next j
            If i <> .RecordCount Then
               grdDataList.Rows = grdDataList.Rows + 1
            End If
         Next i
         SetDataListWidth
      Else
         InsertQueryLog (0) 'Add By Sindy 2010/12/23
         MsgBox "無符合資料！", vbInformation
      End If
   End With
   
   doQuery = True
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdok_Click()
   If grdDataList.TextMatrix(1, 1) <> "" Then
      frm210111_Img.Show vbModal
   End If
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 0
PLeft(1) = 560
PLeft(2) = PLeft(1) + IIf(grdDataList.ColWidth(2) <> 0, Printer.TextWidth(Trim(grdDataList.TextMatrix(1, 1))) + 200 + 560, 0)
For i = 2 To 11
   PLeft(i + 1) = PLeft(i) + IIf(grdDataList.ColWidth(i) <> 0, 200 + Printer.TextWidth(Trim(grdDataList.TextMatrix(1, i))), 0)
Next i
End Sub

Private Sub cmdPrint_Click()
If grdDataList.Rows = 3 And grdDataList.TextMatrix(2, 0) = "" Then MsgBox "沒有資料可以列印！", vbExclamation, "錯誤！": Exit Sub
Screen.MousePointer = vbHourglass
grdDataList.MousePointer = flexHourglass
Dim k As Integer
Page = 1
PrintTitle
For j = 2 To grdDataList.Rows - 1
   Printer.Font.Size = 9
   For k = 0 To 3
      Printer.CurrentX = PLeft(k)
      Printer.CurrentY = iPrint
      Printer.Print IIf(grdDataList.ColWidth(k) <> 0, grdDataList.TextMatrix(j, k), "")
   Next k
   For k = 4 To 12
      If txtSales.Enabled = False And k > 6 Then
      ElseIf txtZone.Enabled = False And k > 8 Then
      Else
         Printer.CurrentX = PLeft(k) + Printer.TextWidth(grdDataList.TextMatrix(1, k)) - Printer.TextWidth(IIf(grdDataList.ColWidth(k) <> 0, grdDataList.TextMatrix(j, k), ""))
         Printer.CurrentY = iPrint
         Printer.Print IIf(grdDataList.ColWidth(k) <> 0, grdDataList.TextMatrix(j, k), "")
      End If
   Next k
   iPrint = iPrint + 200
   If iPrint >= Printer.ScaleHeight - 2000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle
   End If
Next j
Printer.EndDoc
ShowPrintOk
grdDataList.MousePointer = flexDefault
Screen.MousePointer = vbDefault
End Sub

Sub ShowLine()
Printer.Line (0, iPrint + 100)-(Printer.ScaleWidth, iPrint + 100)
iPrint = iPrint + 200
End Sub

Sub PrintTitle()
Printer.Orientation = 1
iPrint = 0
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("新客戶來源分析") / 2
Printer.CurrentY = iPrint
Printer.Print "新客戶來源分析"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
'Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("開發日期：" & ChangeTStringToTDateString(txtCloseDate(0)) & "-" & ChangeTStringToTDateString(txtCloseDate(1)) & "   客戶國籍：" & (txtCuNa(0)) & IIf(txtCuNa(0) & txtCuNa(1) = "", "", "-") & (txtCuNa(1))) / 2
Printer.CurrentX = 2000
Printer.CurrentY = iPrint
Printer.Print "開發日期：" & ChangeTStringToTDateString(txtCloseDate(0)) & "-" & ChangeTStringToTDateString(txtCloseDate(1))
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
Printer.Print "客戶國籍：" & (txtCuNa(0)) & IIf(txtCuNa(0) & txtCuNa(1) = "", "", "-") & (txtCuNa(1))
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & Format(GetTaiwanTodayDate, "##/##/##") & "　")
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "智權人員：" & GetPrjSalesNM(txtSales)
'Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("所別：" & txtZone & " (1.北所 2.中所 3.南所 4.高所)    業務區：" & txtSalesArea & IIf(txtSalesArea & txtSalesArea1 = "", "", "-") & txtSalesArea1) / 2
Printer.CurrentX = 2000
Printer.CurrentY = iPrint
Printer.Print "所別：" & txtZone & " (1.北所 2.中所 3.南所 4.高所)    "
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
Printer.Print "業務區：" & txtSalesArea & IIf(txtSalesArea & txtSalesArea1 = "", "", "-") & txtSalesArea1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & Format(GetTaiwanTodayDate, "##/##/##") & "　")
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
ShowLine
If iPrint >= Printer.ScaleHeight - 2000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
Printer.Font.Size = 9
GetPleft
For i = 0 To 12
   If txtSales.Enabled = False And i > 6 Then
   ElseIf txtZone.Enabled = False And i > 8 Then
   Else
      Printer.CurrentX = PLeft(i) + IIf(grdDataList.ColWidth(i) <> 0, Printer.TextWidth(grdDataList.TextMatrix(1, i)) / 2, 0) - IIf(grdDataList.ColWidth(i) <> 0, Printer.TextWidth(grdDataList.TextMatrix(0, i)) / 2, 0)
      Printer.CurrentY = iPrint
      Printer.Print IIf(grdDataList.ColWidth(i) <> 0, grdDataList.TextMatrix(0, i), "")
   End If
Next i
iPrint = iPrint + 200
If iPrint >= Printer.ScaleHeight - 2000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
For i = 0 To 12
   If txtSales.Enabled = False And i > 6 Then
   ElseIf txtZone.Enabled = False And i > 8 Then
   Else
      Printer.CurrentX = PLeft(i)
      Printer.CurrentY = iPrint
      Printer.Print IIf(grdDataList.ColWidth(i) <> 0, grdDataList.TextMatrix(1, i), "")
   End If
Next i
iPrint = iPrint + 200
If iPrint >= Printer.ScaleHeight - 2000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
ShowLine
If iPrint >= Printer.ScaleHeight - 2000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
Printer.Font.Size = 9
End Sub

Private Sub cmdSearch_Click()
   Screen.MousePointer = vbHourglass
   grdDataList.MousePointer = flexHourglass
   If ConstrainCheck = True Then
      SetDataListWidth
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/23 清除查詢印表記錄檔欄位
      Call doQuery
   End If
   grdDataList.MousePointer = flexDefault
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   txtCloseDate(0) = strSrvDate(2)
   txtCloseDate(1) = strSrvDate(2)
   
   'stST15 = PUB_GetStaffST15(strUserNum, 1)
   stST05 = PUB_GetST05(strUserNum)
   
   'bolAreaMan = False 'Add By Sindy 2023/6/12
   'Modify By Sindy 2020/7/28 設定員編,部門,所別權限
   Call PUB_SetFormSaleDept(strUserNum, txtZone, txtSalesArea, txtSalesArea1, txtSales, bolSpecMan, strSpecCode)
   'Add By Sindy 2023/6/12
   '檢查當時是否需要為他人職代
   Combo3.Clear
   If txtSales <> strUserNum And txtSales <> "" Then
      Combo3.AddItem txtSales & " " & GetPrjSalesNM(txtSales)
   End If
   Combo3.AddItem strUserNum & " " & strUserName
   Call Pub_SetForOthersEmpCombo(strUserNum, Combo3, False, m_strListPer)
   If m_strListPer = "" Then
      Combo3.Visible = False
   Else
      'Add by Amy 2020/03/25 判斷下拉選單是否有區主管
'      If InStr(m_strListPer, GetDeptMan(stST15)) > 0 Then
'         bolAreaMan = True
'      End If
      Combo3.Visible = True
      Combo3.ListIndex = 0
      'Added by Lydia 2021/05/20 Form 2.0物件無法覆蓋Form 1.0
      txtSales.Visible = False
      lblSalesName.Visible = False
   End If
   '2023/6/12 END
   
'   txtZone.Enabled = False
'   txtSalesArea.Enabled = False
'   txtSalesArea1.Enabled = False
'   txtSales.Enabled = False
'   Select Case strUserNum
''cancel by sonia 2014/6/9
''      '2005/9/8 ADD BY SONIA
''      '蔣律師可看中所全部
''      Case "79037"
''         txtZone = pub_strUserOffice
''         txtSalesArea.Enabled = True
''         txtSalesArea1.Enabled = True
''         txtSales.Enabled = True
''         txtSalesArea = "S2"
''         txtSalesArea1 = "S29"
''end 2014/6/9
'      '2005/9/8 END
'      '小真,杜副總可看全部
'      '2013/1/2 加入王副總及郭雅娟可看全部(江總同意)
'      'modify by sonia 2014/6/9 +美珍77027
'      Case "65001", "68006", "71011", "79075", "77027"
'         txtZone.Enabled = True
'         txtSalesArea.Enabled = True
'         txtSalesArea1.Enabled = True
'         txtSales.Enabled = True
'         'Add by Morgan 2005/7/4 副總預設所有業務
'         If strUserNum = "68006" Then
'            txtSalesArea = "S"
'            txtSalesArea1 = "S99"
'         End If
'         '2013/1/2 ADD BY SONIA
'         If strUserNum = "71011" Then
'            txtZone = pub_strUserOffice
'            txtSalesArea = "P10"
'            txtSalesArea1 = "P19"
'         End If
'         If strUserNum = "79075" Then
'            txtZone = pub_strUserOffice
'            txtSalesArea = "P10"
'            txtSalesArea1 = "P19"
'            txtSales = strUserNum
'         End If
'         '2013/1/2 END
'
'      '杜燕文,劉大愛可看S31
'      'modify by sonia 2016/8/22劉大愛78007改為蘇嫄媛79053
'      Case "74018", "79053"
'         txtZone = pub_strUserOffice
'         txtSalesArea = "S31"
'         txtSalesArea1 = "S31"
'         txtSales.Enabled = True
''2013/1/2 CANCEL BY SONIA 因上面已全面開放
''      '王協理可看專利處
''      Case "71011"
''         txtZone = pub_strUserOffice
''         txtSalesArea = "P10"
''         txtSalesArea1 = "P19"
''         txtSales.Enabled = True
'      '葉經理可看商標處
'      'modify by sonia 2016/2/24 +69008
'      Case "67002", "69008"
'         txtZone = pub_strUserOffice
'         txtSalesArea = "P20"
'         txtSalesArea1 = "P29"
'         txtSales.Enabled = True
'      'add by sonia 2016/12/21 柄佑可看中所全部但業務區仍預設自已部門
'      Case "82026"
'         txtZone = pub_strUserOffice
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
'               txtZone.Enabled = True
'               txtSalesArea.Enabled = True
'               txtSalesArea1.Enabled = True
'               txtSales.Enabled = True
'
'            '分所財務人員可看該所全部
''            Case "C1", "NM", "KM"
''               txtZone = pub_strUserOffice
''               txtSalesArea.Enabled = True
''               txtSalesArea1.Enabled = True
''               txtSales.Enabled = True
'
'            '各區主管
'            Case "SM"
'               txtZone = pub_strUserOffice
'               '2005/7/5 ADD BY SONIA
'               txtSalesArea.Locked = True
'               txtSalesArea1.Locked = True
'               '2005/7/5 END
'               '原羅文旭72009可兼看中一區,94/7/1只可看S22
'               '2005/7/5林永生71003可看中所全部,但預設S23
'               If strUserNum = "71003" Then
'                  txtSalesArea = "S23"
'                  txtSalesArea1 = "S23"
'                  '2005/7/5 ADD BY SONIA
'                  txtSalesArea.Locked = False
'                  txtSalesArea1.Locked = False
'                  txtSalesArea.Enabled = True
'                  txtSalesArea1.Enabled = True
'                  '2005/7/5 END
'               Else
'               '2006/11/29簡金泉69005可看北所全部,但預設S15
'               If strUserNum = "69005" Then
'                  txtZone.Enabled = True 'Added by Morgan 2019/12/30 杜主秘說加開簡協理69005可看全所(預設S15)
'                  txtSalesArea = "S15"
'                  txtSalesArea1 = "S15"
'                  txtSalesArea.Locked = False
'                  txtSalesArea1.Locked = False
'                  txtSalesArea.Enabled = True
'                  txtSalesArea1.Enabled = True
'               '2006/11/29 END
'               Else
'                  txtSalesArea = stST15
'                  txtSalesArea1 = stST15
'               End If
'               End If
'               txtSales.Enabled = True
'
'            '其他只能看自己
'            Case Else
'               txtZone = pub_strUserOffice
'               txtSalesArea = stST15
'               txtSalesArea1 = stST15
'               txtSales = strUserNum
'               'Added by Lydia 2017/07/25 多使用者權限,則增加部門範圍
'               strExc(1) = PUB_GetSalesList(strUserNum, , , , , strExc(2), strExc(3))
'               If strExc(3) <> "" And strExc(3) > txtSalesArea1 Then
'                  txtSalesArea1 = strExc(3)
'               End If
'               'end 2017/07/25
'         End Select
'   End Select
'
'   'Add By Sindy 2009/05/12
'   '若操作人員的ST05=SA,此類人員稱之為帶人主管,開放智權人員欄位可輸入
'   If PUB_GetST05Limits(strUserNum) = True Then
'      txtSales.Enabled = True
'   End If
   
   txtSales = strUserNum
   If Mid(txtSalesArea, 1, 1) = "S" And Mid(txtSalesArea1, 1, 1) = "S" Then
      txtCuNa(0).Text = "001"
      txtCuNa(1).Text = "008"
   End If
   SetDataListWidth
   
'   'Add By Sindy 2020/7/1 記錄原操作人可以查詢的業務區及所別
'   txtZone.Tag = txtZone
'   txtSalesArea.Tag = txtSalesArea
'   txtSalesArea1.Tag = txtSalesArea1
'   '2020/7/1 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm210111 = Nothing
End Sub

Private Sub grdDataList_SelChange()
Dim ClickRow As Integer
Dim i As Integer
Dim j As Integer
grdDataList.Visible = False
ClickRow = grdDataList.MouseRow
If ClickRow >= 2 Then
   For j = 2 To grdDataList.Rows - 1
     grdDataList.row = j
     grdDataList.col = 0
     If grdDataList.CellBackColor <> &H80000018 Then
          For i = 4 To grdDataList.Cols - 1
               grdDataList.col = i
               grdDataList.CellBackColor = &H80000018
         Next i
      End If
   Next j
   grdDataList.row = ClickRow
     For i = 4 To grdDataList.Cols - 1
         grdDataList.col = i
         grdDataList.CellBackColor = &HFFC0C0
     Next i
End If
grdDataList.Visible = True
End Sub

Private Sub txtCloseDate_GotFocus(Index As Integer)
   'If Index = 1 Then txtCloseDate(Index) = txtCloseDate(Index - 1)
   TextInverse txtCloseDate(Index)
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtCloseDate(Index).IMEMode = 2
   CloseIme
End Sub

Private Sub txtCloseDate_Validate(Index As Integer, Cancel As Boolean)
   If txtCloseDate(Index) <> "" Then
      If ChkDate(txtCloseDate(Index)) = False Then
         Cancel = True
         txtCloseDate(Index).SetFocus
         txtCloseDate_GotFocus Index
         'add by nickc 2005/08/15
         Exit Sub
      End If
      'add by nickc 2005/08/15
      If Index = 1 Then
         If RunNick2(txtCloseDate(0), txtCloseDate(1)) = True Then
            Cancel = True
            Exit Sub
         End If
      End If
   End If
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
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtSales.IMEMode = 2
   CloseIme
   If Combo3.Enabled = True And Combo3.Visible = True Then Combo3.SetFocus 'Add By Sindy 2023/6/12
End Sub

'Add By Sindy 2010/11/26
Private Sub txtSales_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSales_LostFocus()
   If Trim(txtSales) = "" Then
       lblSalesName = ""
   End If
End Sub

Private Sub txtSales_Validate(Cancel As Boolean)
   'Modify By Sindy 2024/9/18 智權部:使用智權人員ID查詢權限; 改為共用函數
   If PUB_txtSales_Limit(txtSales, m_strListPer, txtZone, txtSalesArea, txtSalesArea1, _
                         bolSpecMan, strSpecCode, lblSalesName) = False Then
      If txtSales.Visible = True Then 'Added by Lydia 2021/05/20 排除隱藏
        txtSales.SetFocus
        txtSales_GotFocus
      End If 'Added by Lydia 2021/05/20
      Cancel = True
      Exit Sub
   End If
   '2024/9/18 END
   
'   'add by sonia 2016/12/21 取消智權人員編號時,無跨所別權限者重新預設所別
'   If Trim(txtSales) = "" And txtZone.Enabled = False Then
'      txtZone = pub_strUserOffice
'   End If
'   'end 2016/12/21
End Sub

Private Sub txtSalesArea_GotFocus()
   TextInverse txtSalesArea
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtSalesArea.IMEMode = 2
   CloseIme
End Sub

Private Sub txtSalesArea_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesArea1_GotFocus()
   TextInverse txtSalesArea1
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtSalesArea.IMEMode = 2
   CloseIme
End Sub

Private Sub txtSalesArea1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'add by nickc 2005/08/15
Private Sub txtSalesArea1_Validate(Cancel As Boolean)
If Trim(txtSalesArea1) <> "" Then
   If RunNick(txtSalesArea, txtSalesArea1) = True Then
      Cancel = True
      Exit Sub
   End If
End If
End Sub

Private Sub txtZone_GotFocus()
   TextInverse txtZone
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtZone.IMEMode = 2
   CloseIme
End Sub

Private Sub txtZone_KeyPress(KeyAscii As Integer)
   If (KeyAscii < Asc("1") Or KeyAscii > Asc("4")) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Add By Sindy 2023/6/12
Private Sub Combo3_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Combo3_LostFocus()
   If Trim(Combo3) <> "" And Trim(Combo3) <> "全部" Then
      arrID = Split(Combo3, " ")
      txtSales = arrID(0)
      Combo3 = txtSales & " " & GetPrjSalesNM(txtSales)
   ElseIf Trim(Combo3) <> "全部" Then
      txtSales = ""
   End If
End Sub
Private Sub Combo3_Validate(Cancel As Boolean)
Dim strEmp As String
Dim stTmp As String 'Add by Amy 2020/03/25
   
   If Combo3 <> "" And Trim(Combo3) <> "全部" Then
      'Add by Amy 2020/03/25 只能輸入下拉選單中已有的人員
      stTmp = Combo3
      '直接輸員編未串名字會錯
      If InStr(stTmp, " ") > 0 Then
        stTmp = Mid(stTmp, 1, Val(InStr(stTmp, " ")) - 1)
      Else
        stTmp = Combo3
      End If
      'Modify By Sindy 2020/6/15 Mark
'      If InStr(m_strListPer, stTmp) = 0 And stTmp <> strUserNum And Pub_StrUserSt03 <> "M51" Then
'         MsgBox "不可輸入下拉選單以外的人員！"
'         Cancel = True
'         Combo3.SetFocus
'         Exit Sub
'      End If
      'end 2020/03/25
      arrID = Split(Combo3, " ")
      txtSales = arrID(0)
      lblSalesName.Caption = GetStaffName(txtSales, True)
      If lblSalesName.Caption = "" Then
         MsgBox "智權人員輸入錯誤！", vbCritical
         Combo3.SetFocus
         Cancel = True
      End If
      Combo3 = txtSales & " " & GetPrjSalesNM(txtSales)
   'Modify By Sindy 2024/8/5 mark; 因 txtSales_Validate 會檢查相關的權限
   Else
      txtSales = ""
'   'Modify by Amy 2023/05/09 +st05
'   ElseIf Combo3 = MsgText(601) And stST05 <> "00" And stST05 <> "01" And stST05 <> "08" Then
'        'Add by Amy 2020/03/25 下拉選單無區主管智權人員不可為空
'        'Modify By Sindy 2020/7/14
'        'If bolAreaMan = False And Pub_StrUserSt03 <> "M51" Then
'        'Modify By Sindy 2023/9/21 開放杜協理權限 + And Not (Mid(txtSalesArea, 1, 1) = "S" And Mid(txtSalesArea1, 1, 1) = "S" And InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) > 0)
'        If (GetDeptMan(txtSalesArea) <> strUserNum Or GetDeptMan(txtSalesArea1) <> strUserNum) _
'            And Pub_StrUserSt03 <> "M51" _
'            And Not (Mid(txtSalesArea, 1, 1) = "S" And Mid(txtSalesArea1, 1, 1) = "S" And InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) > 0) Then
'        '2020/7/14 END
'           MsgBox "非區主管職代智權人員不可空白！"
'           Cancel = True
'           Combo3.SetFocus
'           Exit Sub
'        End If
'        'end 2020/03/25
   '2024/8/5 END
   End If
   'end 2016/6/7
End Sub
'2023/6/12 END
