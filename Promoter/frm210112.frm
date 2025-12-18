VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210112 
   BorderStyle     =   1  '單線固定
   Caption         =   "智權人員收/發文量分析"
   ClientHeight    =   5640
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   9390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   9390
   Begin VB.CommandButton cmdSelCp10 
      Caption         =   "選擇"
      Height          =   255
      Index           =   3
      Left            =   8535
      TabIndex        =   15
      Top             =   1800
      Width           =   810
   End
   Begin VB.CommandButton cmdSelCp10 
      Caption         =   "選擇"
      Height          =   255
      Index           =   2
      Left            =   8535
      TabIndex        =   13
      Top             =   1545
      Width           =   810
   End
   Begin VB.CommandButton cmdSelCp10 
      Caption         =   "選擇"
      Height          =   255
      Index           =   1
      Left            =   8535
      TabIndex        =   11
      Top             =   1290
      Width           =   810
   End
   Begin VB.CommandButton cmdSelCp10 
      Caption         =   "選擇"
      Height          =   255
      Index           =   0
      Left            =   8535
      TabIndex        =   9
      Top             =   1035
      Width           =   810
   End
   Begin VB.TextBox txt_CP10 
      Height          =   270
      Index           =   3
      Left            =   1635
      TabIndex        =   14
      Text            =   "ALL"
      Top             =   1800
      Width           =   6870
   End
   Begin VB.TextBox txt_CP10 
      Height          =   270
      Index           =   2
      Left            =   1635
      TabIndex        =   12
      Text            =   "ALL"
      Top             =   1545
      Width           =   6870
   End
   Begin VB.TextBox txt_CP10 
      Height          =   270
      Index           =   1
      Left            =   1635
      TabIndex        =   10
      Text            =   "ALL"
      Top             =   1290
      Width           =   6870
   End
   Begin VB.TextBox txt_CP10 
      Height          =   270
      Index           =   0
      Left            =   1635
      TabIndex        =   8
      Text            =   "ALL"
      Top             =   1035
      Width           =   6870
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Height          =   340
      Left            =   6410
      Style           =   1  '圖片外觀
      TabIndex        =   17
      Top             =   0
      Width           =   800
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   180
      TabIndex        =   29
      Top             =   735
      Width           =   1845
      Begin VB.OptionButton opt2 
         Caption         =   "發文"
         Height          =   240
         Index           =   1
         Left            =   1050
         TabIndex        =   5
         Top             =   30
         Width           =   765
      End
      Begin VB.OptionButton opt2 
         Caption         =   "收文"
         Height          =   240
         Index           =   0
         Left            =   45
         TabIndex        =   4
         Top             =   45
         Value           =   -1  'True
         Width           =   765
      End
   End
   Begin VB.OptionButton opt1 
      Caption         =   "圓餅圖"
      Height          =   210
      Index           =   1
      Left            =   6435
      TabIndex        =   28
      Top             =   765
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.OptionButton opt1 
      Caption         =   "長條圖"
      Height          =   210
      Index           =   0
      Left            =   5190
      TabIndex        =   27
      Top             =   780
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "統計圖表(&P)"
      Height          =   340
      Left            =   7300
      TabIndex        =   18
      Top             =   0
      Width           =   1160
   End
   Begin VB.TextBox txtSalesArea1 
      Height          =   285
      Left            =   2415
      TabIndex        =   2
      Top             =   405
      Width           =   915
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   340
      Left            =   5520
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   340
      Left            =   8550
      Style           =   1  '圖片外觀
      TabIndex        =   19
      Top             =   15
      Width           =   800
   End
   Begin VB.TextBox txtZone 
      Height          =   285
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   0
      Top             =   75
      Width           =   915
   End
   Begin VB.TextBox txtSales 
      Height          =   285
      Left            =   4665
      MaxLength       =   6
      TabIndex        =   3
      Top             =   435
      Width           =   915
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   285
      Index           =   1
      Left            =   3960
      MaxLength       =   7
      TabIndex        =   7
      Top             =   735
      Width           =   915
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   285
      Index           =   0
      Left            =   2700
      MaxLength       =   7
      TabIndex        =   6
      Top             =   735
      Width           =   915
   End
   Begin VB.TextBox txtSalesArea 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   405
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3510
      Left            =   90
      TabIndex        =   20
      Top             =   2070
      Width           =   9285
      _ExtentX        =   16387
      _ExtentY        =   6191
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
         Size            =   10
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
   Begin MSForms.Label lblSalesName 
      Height          =   300
      Left            =   5640
      TabIndex        =   34
      Top             =   450
      Width           =   1470
      VariousPropertyBits=   27
      Size            =   "2593;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "著作權案件性質："
      Height          =   180
      Left            =   165
      TabIndex        =   33
      Top             =   1845
      Width           =   1440
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "法務案件性質："
      Height          =   180
      Left            =   165
      TabIndex        =   32
      Top             =   1590
      Width           =   1260
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "商標案件性質："
      Height          =   180
      Left            =   165
      TabIndex        =   31
      Top             =   1335
      Width           =   1260
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利案件性質："
      Height          =   180
      Left            =   165
      TabIndex        =   30
      Top             =   1095
      Width           =   1260
   End
   Begin VB.Line Line2 
      X1              =   2145
      X2              =   2415
      Y1              =   525
      Y2              =   525
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   120
      TabIndex        =   26
      Top             =   5415
      Width           =   45
   End
   Begin VB.Label lblZone 
      AutoSize        =   -1  'True
      Caption         =   "（1北所 2中所 3南所 4高所）"
      Height          =   180
      Left            =   2235
      TabIndex        =   25
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   3630
      TabIndex        =   24
      Top             =   480
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   3645
      X2              =   3915
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "所別："
      Height          =   180
      Left            =   165
      TabIndex        =   23
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label2 
      Caption         =   "日期："
      Height          =   180
      Left            =   2130
      TabIndex        =   22
      Top             =   795
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "業務區："
      Height          =   180
      Left            =   165
      TabIndex        =   21
      Top             =   450
      Width           =   720
   End
End
Attribute VB_Name = "frm210112"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/05 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、lblSalesName
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

Dim i As Integer
Dim j As Integer
Dim iPrint As Integer, Page As Integer, strTemp(0 To 12) As String
Dim PLeft(0 To 12) As Integer
Public Calrs As New ADODB.Recordset
Public SeekIdx As Integer
Dim bolQuery As Boolean 'Added by Lydia 2017/12/21 查詢是否有資料(By 秀玲)

Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
Dim m_strListPer As String 'Add By Sindy 2020/7/1


Private Sub SetDataListWidth()
Dim arrGridHeadText, arrGridHeadWidth, arrGridHeadText2
Dim iRow As Integer

arrGridHeadText = Array("", "", "", "" _
                  , "個人", "個人", "個人", "該區", "該區", "全所", "區佔全所" _
                  , "全所", "全所")
                     
arrGridHeadText2 = Array("所別", "業務區", "智權人員", "種類" _
                  , "件數", "總件數", "比　　例", "件數", "比　　例", "業務件數", "業務比例" _
                  , "總件數", "比　　例")
                     
                     
arrGridHeadWidth = Array(500, 700, 700, 500 _
                     , 500, 850, 850, 850, 850, 850, 850 _
                     , 850, 850)
                     
If Me.txtZone.Locked = True Or Me.txtZone.Enabled = False Then

End If
If (Me.txtSalesArea.Locked = True And Me.txtSalesArea1.Locked = True) Or (Me.txtSalesArea.Enabled = False And Me.txtSalesArea1.Enabled = False) Then
   
End If
If Me.txtSales.Locked = True Or Me.txtSales.Enabled = False Then
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
   grdDataList.CellAlignment = flexAlignCenterCenter
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

grdDataList.MergeCells = flexMergeRestrictRows
grdDataList.MergeRow(0) = True
End Sub

Private Function ConstrainCheck() As Boolean
Dim bolCancel As Boolean
Dim intErrCol As Integer
   
   ConstrainCheck = True
   
   'Add By Sindy 2009/05/14
   Call txtSales_Validate(bolCancel)
   If bolCancel = True Then
      txtSales.SetFocus
      txtSales_GotFocus
      ConstrainCheck = False
      Exit Function
   End If
   
   If txtCloseDate(0) = "" Then
      MsgBox "請輸入收/發文起日！", vbExclamation
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
      MsgBox "請輸入收/發文迄日！", vbExclamation
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
'   '2006/11/29 ADD BY SONIA 簡金泉69005檢查業務區範圍
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
   Dim stVTBx As String, stVTBy As String, stVTBz As String
   Dim stConAA As String, stConBB As String, stConCC As String, stConDD As String
   Dim stConEE As String, stConFF As String, stConGG As String, stConHH As String
   Dim stCon1 As String, stCon2 As String
   Dim mainData2 As String 'Added by Lydia 2017/12/21
   
   bolQuery = False 'Added by Lydia 2017/12/21 查詢是否有資料
   
   stCon1 = "": stCon2 = ""
   stCon = ""
   stConAA = "": stConBB = "": stConCC = "": stConDD = "": stConEE = "": stConFF = "": stConGG = "": stConHH = ""
   stConST = " and cp13 is not null "
   stConST1 = " and cp13 is not null "
   stConST2 = " and cp13 is not null "
   stConST3 = " and cp13 is not null "
   '所別
   If txtZone <> "" Then
      stConST = stConST & " and s2.st06 = '" & txtZone & "' " 'Modify By Sindy 2021/8/5 + s2.
      stConST2 = stConST2 & " and s2.st06 = '" & txtZone & "' " 'Modify By Sindy 2021/8/5 + s2.
      stConST3 = stConST3 & " and s2.st06 = '" & txtZone & "' " 'Modify By Sindy 2021/8/5 + s2.
      pub_QL05 = pub_QL05 & ";" & Label3 & txtZone & lblZone 'Add By Sindy 2010/12/23
   End If
   
   '區別
   'Modify By Sindy 2009/05/12
   '若為帶人主管權限時,查詢之智權人員編號非本人時,不限制區別
   If (Trim(txtSales) <> "" And PUB_GetST05(strUserNum) = "SA" And txtSales.Enabled = True And txtSales <> strUserNum) Then
      '不限制區別
   Else
      If txtSalesArea <> "" Then
         stConST = stConST & " and s2.st15||''>='" & txtSalesArea & "' " 'Modify By Sindy 2021/8/4 cp12 => s2.st15
         stConST2 = stConST2 & " and s2.st15||''>='" & txtSalesArea & "' " 'Modify By Sindy 2021/8/4 cp12 => s2.st15
         stConST3 = stConST3 & " and s2.st15||''>='" & txtSalesArea & "' " 'Modify By Sindy 2021/8/4 cp12 => s2.st15
      End If
      If txtSalesArea1 <> "" Then
         stConST = stConST & " and s2.st15||''<='" & txtSalesArea1 & "' " 'Modify By Sindy 2021/8/4 cp12 => s2.st15
         stConST2 = stConST2 & " and s2.st15||''<='" & txtSalesArea1 & "' " 'Modify By Sindy 2021/8/4 cp12 => s2.st15
         stConST3 = stConST3 & " and s2.st15||''<='" & txtSalesArea1 & "' " 'Modify By Sindy 2021/8/4 cp12 => s2.st15
      End If
      If txtSalesArea <> "" Or txtSalesArea1 <> "" Then
         pub_QL05 = pub_QL05 & ";" & Label1 & txtSalesArea & "-" & txtSalesArea1 'Add By Sindy 2010/12/23
      End If
   End If
   
   '智權人員
   If txtSales <> "" Then
      stConST = stConST & " and cp13||'' = '" & txtSales & "' "
      pub_QL05 = pub_QL05 & ";" & Label4 & txtSales & lblSalesName 'Add By Sindy 2010/12/23
   End If

   '本所期限
   If txtCloseDate(0) <> "" Then
      If opt2(0).Value = True Then
         stCon1 = stCon1 & " and cp05>=" & ChangeTStringToWString(txtCloseDate(0))
      Else
         stCon1 = stCon1 & " and cp27>=" & ChangeTStringToWString(txtCloseDate(0))
      End If
   End If
   If txtCloseDate(1) <> "" Then
      If opt2(0).Value = True Then
         stCon1 = stCon1 & " and cp05<=" & ChangeTStringToWString(txtCloseDate(1))
      Else
         stCon1 = stCon1 & " and cp27<=" & ChangeTStringToWString(txtCloseDate(1))
      End If
   End If
   If txtCloseDate(0) <> "" Or txtCloseDate(1) <> "" Then
      If opt2(0).Value = True Then
         pub_QL05 = pub_QL05 & ";" & opt2(0).Caption 'Add By Sindy 2010/12/23
      Else
         pub_QL05 = pub_QL05 & ";" & opt2(1).Caption 'Add By Sindy 2010/12/23
      End If
      pub_QL05 = pub_QL05 & txtCloseDate(0) & "-" & txtCloseDate(1)  'Add By Sindy 2010/12/23
   End If
   
   'modify by sonia 2022/5/2 加入cp159=0條件，否則會與收/發文量查詢結果不同
   stCon1 = stCon1 & " and cp09<'B' and cp159=0 "
   If txt_CP10(0) & txt_CP10(1) & txt_CP10(2) & txt_CP10(3) <> "" Then
         If txt_CP10(0) <> "" Then
            pub_QL05 = pub_QL05 & ";" & Label5 & txt_CP10(0) 'Add By Sindy 2010/12/23
         End If
         If txt_CP10(1) <> "" Then
            pub_QL05 = pub_QL05 & ";" & Label6 & txt_CP10(1) 'Add By Sindy 2010/12/23
         End If
         If txt_CP10(2) <> "" Then
            pub_QL05 = pub_QL05 & ";" & Label7 & txt_CP10(2) 'Add By Sindy 2010/12/23
         End If
         If txt_CP10(3) <> "" Then
            pub_QL05 = pub_QL05 & ";" & Label8 & txt_CP10(3) 'Add By Sindy 2010/12/23
         End If
         stCon2 = stCon2 & " and ( "
         '將所有的案件性質組合
         If txt_CP10(0) <> "ALL" And txt_CP10(0) <> "" Then
            stCon2 = stCon2 & " (cp01 in ('P','PS','CFP','CPS') and cp10||'' in (" & GetAddStr(txt_CP10(0)) & ") ) "
         ElseIf txt_CP10(0) <> "" Then
            stCon2 = stCon2 & " (cp01 in ('P','PS','CFP','CPS') ) "
         Else
            stCon2 = stCon2 & " (2=1) "
         End If
         If txt_CP10(1) <> "ALL" And txt_CP10(1) <> "" Then
            stCon2 = stCon2 & " or (cp01 in ('T','TF','CFT') and cp10||'' in (" & GetAddStr(txt_CP10(1)) & ") ) "
         ElseIf txt_CP10(1) <> "" Then
            stCon2 = stCon2 & " or (cp01 in ('T','TF','CFT') ) "
         Else
            stCon2 = stCon2 & " or (2=1) "
         End If
         'modify by sonia 2022/5/2 L,LA,CFL屬於法務；CFC,TC屬於著作權
         If txt_CP10(2) <> "ALL" And txt_CP10(2) <> "" Then
            stCon2 = stCon2 & " or (cp01 in ('L','LA','CFL') and cp10||'' in (" & GetAddStr(txt_CP10(2)) & ") ) "
         ElseIf txt_CP10(2) <> "" Then
            stCon2 = stCon2 & " or (cp01 in ('L','LA','CFL') ) "
         Else
            stCon2 = stCon2 & " or (2=1) "
         End If
         'modify by sonia 2022/5/2 L,LA,CFL屬於法務；CFC,TC屬於著作權
         If txt_CP10(3) <> "ALL" And txt_CP10(3) <> "" Then
            stCon2 = stCon2 & " or (cp01 in ('CFC','TC') and cp10||'' in (" & GetAddStr(txt_CP10(3)) & ") ) "
         ElseIf txt_CP10(3) <> "" Then
            stCon2 = stCon2 & " or (cp01 in ('CFC','TC') ) "
         Else
            stCon2 = stCon2 & " or (2=1) "
         End If
         stCon2 = stCon2 & " ) "
   End If
   
On Error GoTo ErrHnd

   Dim MainData As String
   'modify by sonia 2022/5/2 L,LA,CFL屬於法務；CFC,TC屬於著作權
   'MainData = "select decode(cp01,'P','1','PS','1','CFP','1','CPS','1','T','2','TF','2','CFT','2','TC','2','CFC','3','CFL','3','L','4','LA','4','5') as cp01,cp12,cp13,cp09 from caseprogress,staff where cp13=st01(+) and substr(cp01,1,2)<>'FC' and substr(cp01,1,2)<>'FG' " & stCon1 & stCon2
   MainData = "select decode(cp01,'P','1','PS','1','CFP','1','CPS','1','T','2','TF','2','CFT','2','CFL','4','L','4','LA','4','TC','3','CFC','3','5') as cp01,cp12,cp13,cp09 from caseprogress,staff where cp13=st01(+) and substr(cp01,1,2)<>'FC' and substr(cp01,1,2)<>'FG' " & stCon1 & stCon2
   'Added by Lydia 2017/12/21 O12會分析staff,在全所業務總件數(CC)和全所件數(DD)會造成SQL 錯誤: 沒有從埠座讀取資料, 進而連線中斷
   'modify by sonia 2022/5/2 L,LA,CFL屬於法務；CFC,TC屬於著作權
   'mainData2 = "select decode(cp01,'P','1','PS','1','CFP','1','CPS','1','T','2','TF','2','CFT','2','TC','2','CFC','3','CFL','3','L','4','LA','4','5') as cp01,cp12,cp13,cp09 from caseprogress where substr(cp01,1,2)<>'FC' and substr(cp01,1,2)<>'FG' " & stCon1 & stCon2
   mainData2 = "select decode(cp01,'P','1','PS','1','CFP','1','CPS','1','T','2','TF','2','CFT','2','CFL','4','L','4','LA','4','TC','3','CFC','3','5') as cp01,cp12,cp13,cp09 from caseprogress where substr(cp01,1,2)<>'FC' and substr(cp01,1,2)<>'FG' " & stCon1 & stCon2
   'Modify By Sindy 2021/8/4 + ,staff
   stConAA = "select s2.st06 as 所別,cp12 as 業務區,cp13 as 智權人員,cp01 as 種類,count(cp09) as 個人件數  from  (" & MainData & ") QA,staff s2 " & _
                     " where cp13=s2.st01(+) " & stConST & " group by s2.st06,cp12,cp13,cp01 "
   stConBB = "select s2.st06 as 所別,cp12 as 業務區,cp01 as 種類,count(cp09) as 區件數  from  (" & MainData & ") QA,staff s2 " & _
                     " where cp13=s2.st01(+) " & stConST3 & " group by s2.st06,cp12,cp01 "
   'Modified by Lydia 2017/12/21 MainData ->mainData2
   stConCC = "select cp01 as 種類,count(cp09) as 全所業務總件數  from  (" & mainData2 & ") QA,staff s2 " & _
                     " where cp13=s2.st01(+) and s2.st15||''>='S11' and s2.st15||''<='S99' " & stConST1 & " group by cp01 "
   'Modified by Lydia 2017/12/21 MainData ->mainData2
   stConDD = "select cp01 as 種類,count(cp09) as 全所件數  from  (" & mainData2 & ") QA,staff s2 " & _
                      " where cp13=s2.st01(+) " & stConST1 & " group by cp01 "
   stConEE = "select s2.st06 as 所別,cp12 as 業務區,cp13 as 智權人員,count(cp09) as 個人總件數  from  (" & MainData & ") QA,staff s2 " & _
                    " where cp13=s2.st01(+) " & stConST & " group by s2.st06,cp12,cp13 "
   
   stCon = "select decode(ZZ.所別,'1','北所','2','中所','3','南所','4','高所','5','其他','') as 所別,a0902 as 業務區,decode(ZZ.智權人員,'ZA','區統計','ZB','該所統計','ZC','全所統計','ZD','全所合計',st02) as 智權人員, " & _
                " decode(ZZ.種類,'ALL','所有','1','專利','2','商標','3','著作權','4','法務','其他') as 種類,ZZ.個人件數,ZZ.個人總件數,ZZ.個人比例, " & _
                " ZZ.區件數,ZZ.該區比例, " & _
                " ZZ.全所業務總件數,ZZ.全所業務比例, " & _
                " ZZ.全所件數,ZZ.全所比例 " & _
                " from (select AA.所別,AA.業務區,AA.智權人員, " & _
                " AA.種類,AA.個人件數,EE.個人總件數,decode(to_char(round(AA.個人件數/EE.個人總件數*100,2)),null,'',to_char(round(AA.個人件數/EE.個人總件數*100,2),9999.99)||' %') as 個人比例, " & _
                " BB.區件數,decode(to_char(round(AA.個人件數/BB.區件數*100,2)),null,'',to_char(round(AA.個人件數/BB.區件數*100,2),9999.99)||' %') as 該區比例, " & _
                " CC.全所業務總件數,decode(to_char(round(BB.區件數/CC.全所業務總件數*100,2)),null,'',to_char(round(BB.區件數/CC.全所業務總件數*100,2),9999.99)||' %') as 全所業務比例, " & _
                " DD.全所件數,decode(to_char(round(BB.區件數/DD.全所件數*100,2)),null,'',to_char(round(BB.區件數/DD.全所件數*100,2),9999.99)||' %') as 全所比例 " & _
                " from (" & stConAA & ") AA,(" & stConBB & ") BB,(" & stConCC & ") CC,(" & stConDD & ") DD,(" & stConEE & ") EE " & _
                " where AA.所別=BB.所別(+) " & _
                " and AA.業務區=BB.業務區(+) and AA.種類=BB.種類(+) and AA.種類=CC.種類(+) " & _
                " and AA.種類=DD.種類(+) and AA.所別=EE.所別(+) and AA.業務區=EE.業務區(+) and AA.智權人員=EE.智權人員(+) " '& _
                " union select distinct AA.所別,AA.業務區,'ZA' as 智權人員, " & _
                " AA.種類,nvl(BB.區件數,0) as 個人件數,0 as 個人總件數,'' as 個人比例, " & _
                " 0 as 區件數,'' as 該區比例, " & _
                " 0 as 全所業務總件數,'' as 全所業務比例, " & _
                " 0 as 全所件數,'' as 全所比例 " & _
                " from (" & stConAA & ") AA,(" & stConBB & ") BB " & _
                " where AA.種類=BB.種類(+) and AA.所別=BB.所別(+) and AA.業務區=BB.業務區(+)  "
'      stCon = stCon & " union select distinct AA.所別,'' as 業務區,'ZB' as 智權人員, " & _
'                " AA.種類,nvl(CC.所件數,0) as 個人件數,0 as 個人總件數,'' as 個人比例, " & _
'                " 0 as 區件數,'' as 該區比例, " & _
'                " 0 as 所件數,'' as 該所比例, " & _
'                " 0 as 全所件數,'' as 全所比例 " & _
'                " from (" & stConAA & ") AA,(" & stConCC & ") CC " & _
'                " where AA.種類=CC.種類(+) and AA.所別=CC.所別(+) "
'      stCon = stCon & " union select distinct '' as 所別,'' as 業務區,'ZC' as 智權人員, " & _
'                " AA.種類,nvl(DD.全所件數,0) as 個人件數,0 as 個人總件數,'' as 個人比例, " & _
'                " 0 as 區件數,'' as 該區比例, " & _
'                " 0 as 所件數,'' as 該所比例, " & _
'                " 0 as 全所件數,'' as 全所比例 " & _
'                " from (" & stConAA & ") AA,(" & stConDD & ") DD " & _
'                " where AA.種類=DD.種類(+) "
'      stCon = stCon & " union select '' as 所別,'' as 業務區,'ZD' as 智權人員, " & _
'                " 'ALL' as 種類,sum(nvl(DD.全所件數,0)) as 個人件數,0 as 個人總件數,'' as 個人比例, " & _
'                " 0 as 區件數,'' as 該區比例, " & _
'                " 0 as 所件數,'' as 該所比例, " & _
'                " 0 as 全所件數,'' as 全所比例 " & _
'                " from (" & stConDD & ") DD "

     stCon = stCon & " ) ZZ,staff,acc090 " & _
                " where ZZ.業務區=a0901(+) and ZZ.智權人員=st01(+)  " & _
                " order by ZZ.所別,ZZ.業務區,ZZ.智權人員,ZZ.種類 "

   CheckOC3
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
         'Set grdDataList.Recordset = AdoRecordSet3.Clone
         bolQuery = True 'Added by Lydia 2017/12/21 查詢是否有資料
         
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
   'Resume
End Function

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdok_Click()
   'Added by Lydia 2017/12/21 查詢是否有資料
   If bolQuery = False Then MsgBox "沒有資料可以呈現！", vbExclamation, "錯誤！": Exit Sub
   'end 2017/12/21
   frm210112_Img.Show vbModal
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

Private Sub cmdSelCp10_Click(Index As Integer)
SeekIdx = Index
frm210112_1.Show vbModal
End Sub

Private Sub Form_Load()
   Dim stST05 As String, stST15 As String
   
   MoveFormToCenter Me
   
   stST15 = PUB_GetStaffST15(strUserNum, 1)
   stST05 = PUB_GetST05(strUserNum)
   
   txtCloseDate(0) = strSrvDate(2)
   txtCloseDate(1) = strSrvDate(2)
   
   'Modify By Sindy 2020/7/28 設定員編,部門,所別權限
   Call PUB_SetFormSaleDept(strUserNum, txtZone, txtSalesArea, txtSalesArea1, txtSales, bolSpecMan, strSpecCode)
   
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
'      'modify by sonia 2014/6/9 +美珍77027
'      Case "65001", "68006", "77027"
'         txtZone.Enabled = True
'         txtSalesArea.Enabled = True
'         txtSalesArea1.Enabled = True
'         txtSales.Enabled = True
'         'Add by Morgan 2005/7/4 副總預設所有智權人員
'         If strUserNum = "68006" Then
'            txtSalesArea = "S"
'            txtSalesArea1 = "S99"
'         End If
'      '杜燕文,劉大愛可看S31
'      'modify by sonia 2016/8/22劉大愛78007改為蘇嫄媛79053
'      Case "74018", "79053"
'         txtZone = pub_strUserOffice
'         txtSalesArea = "S31"
'         txtSalesArea1 = "S31"
'         txtSales.Enabled = True
'
'      '王協理可看專利處
'      Case "71011"
'         txtZone = pub_strUserOffice
'         txtSalesArea = "P10"
'         txtSalesArea1 = "P19"
'         txtSales.Enabled = True
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
''               txtSalesArea.Locked = True
''               txtSalesArea1.Locked = True
''               txtSales.Locked = True
''               txtZone.Locked = True
'         End Select
'   End Select
'
'   'Add By Sindy 2009/05/12
'   '若操作人員的ST05=SA,此類人員稱之為帶人主管,開放智權人員欄位可輸入
'   If PUB_GetST05Limits(strUserNum) = True Then
'      txtSales.Enabled = True
'   End If
'   txtSales = strUserNum
   
   SetDataListWidth
   
'   'Add By Sindy 2020/7/1 記錄原操作人可以查詢的業務區及所別
'   txtZone.Tag = txtZone
'   txtSalesArea.Tag = txtSalesArea
'   txtSalesArea1.Tag = txtSalesArea1
'   '2020/7/1 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm210112 = Nothing
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

Private Sub txt_CP10_GotFocus(Index As Integer)
txt_CP10(Index).SelStart = 0
txt_CP10(Index).SelLength = Len(txt_CP10(Index))
End Sub

Private Sub txt_CP10_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt_CP10_Validate(Index As Integer, Cancel As Boolean)
'If Trim(txt_CP10(Index)) = "" Then
'   MsgBox "案件性質不可空白！" & vbCrLf & vbCrLf & "全部請輸入 ALL！", vbExclamation, "警告！"
'   txt_CP10(Index).SetFocus
'   Cancel = True
'End If
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
End Sub

'Add By Sindy 2010/11/26
Private Sub txtSales_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
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
   
'   'Add By Sindy 2015/6/26 若有異動智權人員,需重新查詢業務區和所別
'   'modify by sonia 2016/6/15 加入帶人主管條件
'   'If txtSalesArea.Enabled = True Then 'Modify By Sindy 2016/5/5 + if
'   If txtSalesArea.Enabled = True Or PUB_GetST05Limits(strUserNum) = True Then
'      If txtSales.Text <> "" And txtSales.Text <> txtSales.Tag Then
'         txtZone = PUB_GetST06(Trim(txtSales))
'         txtSalesArea = PUB_GetStaffST15(Trim(txtSales), "1")
'         txtSalesArea1 = PUB_GetStaffST15(Trim(txtSales), "1")
'      End If
'   Else
'      'Add By Sindy 2016/5/6 還原(原操作人)可以查詢的業務區及所別
'      txtZone = txtZone.Tag
'      txtSalesArea = txtSalesArea.Tag
'      txtSalesArea1 = txtSalesArea1.Tag
'      '2016/5/6 END
'   End If
'
''   'add by sonia 2016/6/7 S29
''   If Len(txtSales) <= 4 Then
''      txtSales.Text = Mid(txtSales.Text & "  ", 1, 5)
''   End If
''   'end 2016/6/7
'
'   'Add by Amy 2017/01/12 +MCTF人員
'   'Remove by Lydia 2017/07/21 併入PUB_GetSalesList
'   'strMCTF = GetMCTF0XCode(txtSales)
'
'   txtSales.Tag = txtSales.Text
'
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
'Modified by Lydia 2017/12/21 查詢是否有資料
'If grdDataList.Rows = 2 And grdDataList.TextMatrix(1, 0) = "" Then MsgBox "沒有資料可以列印！", vbExclamation, "錯誤！": Exit Sub
If bolQuery = False Then MsgBox "沒有資料可以列印！", vbExclamation, "錯誤！": Exit Sub
'end 2017/12/21

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
Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("智權人員收/發文量分析") / 2
Printer.CurrentY = iPrint
Printer.Print "智權人員收/發文量分析"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 2000
Printer.CurrentY = iPrint
Printer.Print IIf(opt2(0).Value = True, "收文", "發文") & "日期：" & ChangeTStringToTDateString(txtCloseDate(0)) & "-" & ChangeTStringToTDateString(txtCloseDate(1))
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
Printer.CurrentX = 2000
Printer.CurrentY = iPrint
Printer.Print "所別：" & txtZone & " (1.北所 2.中所 3.南所 4.高所)    "
Printer.CurrentX = 7000
Printer.CurrentY = iPrint
Printer.Print "業務區：" & txtSalesArea & IIf(txtSalesArea & txtSalesArea1 = "", "", "-") & txtSalesArea1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & Format(GetTaiwanTodayDate, "##/##/##") & "　")
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "專利案件性質：" & txt_CP10(0)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "商標案件性質：" & txt_CP10(1)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "法務案件性質：" & txt_CP10(2)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "著作權案件性質：" & txt_CP10(3)
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
      Printer.CurrentX = PLeft(i)
      Printer.CurrentY = iPrint
      Select Case i
      Case 6
         Printer.Print "個　　人"
      Case 8
         Printer.Print "該　　區"
      Case 9
         Printer.Print "全　　所"
      Case 12
         Printer.Print "全　　所"
      Case Else
         Printer.Print IIf(grdDataList.ColWidth(i) <> 0, grdDataList.TextMatrix(0, i), "")
      End Select
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
