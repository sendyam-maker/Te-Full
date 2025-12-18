VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210134 
   BorderStyle     =   1  '單線固定
   Caption         =   "未發文案件管制"
   ClientHeight    =   5900
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5900
   ScaleWidth      =   8950
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm210134.frx":0000
      Left            =   4050
      List            =   "frm210134.frx":001F
      TabIndex        =   10
      Top             =   690
      Width           =   750
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frm210134.frx":0041
      Left            =   4050
      List            =   "frm210134.frx":0060
      TabIndex        =   6
      Top             =   360
      Width           =   750
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "搜尋(&Q)"
      Height          =   315
      Left            =   4815
      TabIndex        =   15
      Top             =   1650
      Width           =   765
   End
   Begin VB.TextBox txtRecvDate 
      Height          =   285
      Index           =   0
      Left            =   1305
      MaxLength       =   7
      TabIndex        =   8
      Top             =   690
      Width           =   915
   End
   Begin VB.TextBox txtRecvDate 
      Height          =   285
      Index           =   1
      Left            =   2400
      MaxLength       =   7
      TabIndex        =   9
      Top             =   690
      Width           =   915
   End
   Begin VB.OptionButton Option1 
      Caption         =   "收文日期："
      Height          =   285
      Index           =   2
      Left            =   60
      TabIndex        =   7
      Top             =   690
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "所有未發文資料"
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   11
      Top             =   990
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "管制期限："
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   390
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmdControlNDate 
      Caption         =   "管制新期限(&M)"
      Height          =   375
      Left            =   7515
      Style           =   1  '圖片外觀
      TabIndex        =   30
      Top             =   870
      Width           =   1380
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "繪圖進度(&D)"
      Height          =   375
      Index           =   2
      Left            =   7785
      Style           =   1  '圖片外觀
      TabIndex        =   28
      Top             =   465
      Width           =   1110
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8100
      Style           =   1  '圖片外觀
      TabIndex        =   25
      Top             =   75
      Width           =   800
   End
   Begin VB.CommandButton cmdControlOK 
      Caption         =   "管制完成(&C)"
      Height          =   375
      Left            =   7785
      Style           =   1  '圖片外觀
      TabIndex        =   31
      Top             =   1260
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "發E-Mail(&S)"
      Height          =   375
      Index           =   3
      Left            =   6360
      Style           =   1  '圖片外觀
      TabIndex        =   29
      Top             =   870
      Width           =   1110
   End
   Begin VB.TextBox txtCP10 
      Height          =   285
      Index           =   1
      Left            =   5010
      MaxLength       =   4
      TabIndex        =   18
      Top             =   1980
      Width           =   435
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1305
      MaxLength       =   3
      TabIndex        =   19
      Top             =   2280
      Width           =   612
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   1950
      MaxLength       =   6
      TabIndex        =   20
      Top             =   2280
      Width           =   1236
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   3210
      MaxLength       =   1
      TabIndex        =   21
      Top             =   2280
      Width           =   276
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   3510
      MaxLength       =   2
      TabIndex        =   22
      Top             =   2280
      Width           =   420
   End
   Begin VB.TextBox txtCP10 
      Height          =   285
      Index           =   0
      Left            =   4410
      MaxLength       =   4
      TabIndex        =   17
      Top             =   1980
      Width           =   435
   End
   Begin VB.TextBox systemkind 
      Height          =   270
      Left            =   1305
      TabIndex        =   16
      Text            =   "ALL"
      Top             =   1980
      Width           =   2130
   End
   Begin VB.TextBox txtCU2 
      Height          =   285
      Left            =   2400
      MaxLength       =   9
      TabIndex        =   13
      Top             =   1320
      Width           =   970
   End
   Begin VB.TextBox txtControlDate 
      Height          =   285
      Index           =   1
      Left            =   2400
      MaxLength       =   7
      TabIndex        =   5
      Top             =   390
      Width           =   915
   End
   Begin VB.TextBox txtCU1 
      Height          =   285
      Left            =   1305
      MaxLength       =   9
      TabIndex        =   12
      Top             =   1320
      Width           =   970
   End
   Begin VB.TextBox txtControlDate 
      Height          =   285
      Index           =   0
      Left            =   1305
      MaxLength       =   7
      TabIndex        =   4
      Top             =   390
      Width           =   915
   End
   Begin VB.TextBox txtSales 
      Height          =   285
      Left            =   3345
      MaxLength       =   6
      TabIndex        =   2
      Top             =   30
      Width           =   915
   End
   Begin VB.TextBox txtSalesArea 
      Height          =   285
      Left            =   1305
      MaxLength       =   3
      TabIndex        =   0
      Top             =   30
      Width           =   435
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "承辦進度(&E)"
      Height          =   375
      Index           =   1
      Left            =   6630
      Style           =   1  '圖片外觀
      TabIndex        =   27
      Top             =   465
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "歷史記錄(&H)"
      Height          =   375
      Index           =   0
      Left            =   5460
      Style           =   1  '圖片外觀
      TabIndex        =   26
      Top             =   465
      Width           =   1110
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   375
      Left            =   7260
      Style           =   1  '圖片外觀
      TabIndex        =   24
      Top             =   75
      Width           =   800
   End
   Begin VB.TextBox txtSalesArea1 
      Height          =   285
      Left            =   1890
      MaxLength       =   3
      TabIndex        =   1
      Top             =   30
      Width           =   435
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3285
      Left            =   60
      TabIndex        =   23
      Top             =   2610
      Width           =   8895
      _ExtentX        =   15699
      _ExtentY        =   5786
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
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
   Begin MSForms.ComboBox Combo3 
      Height          =   330
      Left            =   3345
      TabIndex        =   52
      Top             =   30
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
   Begin MSForms.TextBox txtCuName 
      Height          =   330
      Left            =   1680
      TabIndex        =   14
      Top             =   1620
      Width           =   3075
      VariousPropertyBits=   671105051
      Size            =   "5424;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblSalesName 
      Height          =   300
      Left            =   4320
      TabIndex        =   51
      Top             =   60
      Width           =   1470
      VariousPropertyBits=   27
      Size            =   "2593;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "個月)"
      Height          =   195
      Index           =   2
      Left            =   4800
      TabIndex        =   50
      Top             =   735
      Width           =   435
   End
   Begin VB.Label Label13 
      Caption         =   "系統日("
      Height          =   255
      Left            =   3360
      TabIndex        =   49
      Top             =   735
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "個月)"
      Height          =   195
      Index           =   1
      Left            =   4800
      TabIndex        =   48
      Top             =   405
      Width           =   435
   End
   Begin VB.Label Label7 
      Caption         =   "系統日("
      Height          =   255
      Left            =   3360
      TabIndex        =   47
      Top             =   405
      Width           =   645
   End
   Begin VB.Label lblCuNam 
      Caption         =   "客戶中文名稱："
      Height          =   180
      Left            =   330
      TabIndex        =   46
      Top             =   1680
      Width           =   1275
   End
   Begin VB.Label LblCntTime 
      AutoSize        =   -1  'True
      Caption         =   "執行時間："
      ForeColor       =   &H00000080&
      Height          =   180
      Left            =   1830
      TabIndex        =   45
      Top             =   1050
      Width           =   4440
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "逾管制期限"
      Height          =   180
      Left            =   6600
      TabIndex        =   44
      Top             =   1590
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "    "
      Height          =   180
      Left            =   6420
      TabIndex        =   43
      Top             =   1590
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "99/9/1起收文"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   6120
      TabIndex        =   42
      Top             =   120
      Width           =   990
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "    "
      Height          =   180
      Left            =   6420
      TabIndex        =   41
      Top             =   1380
      Width           =   180
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "逾預定會稿日"
      Height          =   180
      Left            =   6600
      TabIndex        =   40
      Top             =   1380
      Width           =   1080
   End
   Begin VB.Line Line5 
      X1              =   2220
      X2              =   2490
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PS：請自行點選資料排序條件（點選該欄位標題）"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   3960
      TabIndex        =   39
      Top             =   2340
      Width           =   4860
   End
   Begin VB.Line Line4 
      X1              =   4800
      X2              =   5070
      Y1              =   2130
      Y2              =   2130
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   330
      TabIndex        =   38
      Top             =   2310
      Width           =   900
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   3510
      TabIndex        =   37
      Top             =   2025
      Width           =   900
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Left            =   330
      TabIndex        =   36
      Top             =   2025
      Width           =   900
   End
   Begin VB.Line Line3 
      X1              =   2220
      X2              =   2490
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Left            =   330
      TabIndex        =   35
      Top             =   1365
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "業務區："
      Height          =   180
      Index           =   0
      Left            =   330
      TabIndex        =   34
      Top             =   75
      Width           =   720
   End
   Begin VB.Line Line1 
      X1              =   2220
      X2              =   2490
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   2430
      TabIndex        =   33
      Top             =   75
      Width           =   900
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   105
      TabIndex        =   32
      Top             =   5430
      Width           =   45
   End
   Begin VB.Line Line2 
      X1              =   1710
      X2              =   1980
      Y1              =   150
      Y2              =   150
   End
End
Attribute VB_Name = "frm210134"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/03 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、txtCuName、lblSalesName
'Memo by Lydia 2019/07/01 表單名稱:智權部自行管制未發文案件作業=>未發文案件管制
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/20 日期欄已修改
Option Explicit

Dim bolShowMsgBox As Boolean, bolSelData As Boolean
'紀錄作用按鍵
Public cmdState As Integer
Dim i As Integer, j As Integer
Dim isLoad As Boolean
Dim stST05 As String, stST15 As String
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim strOrderBy As String
'Add by Amy 2014/05/20
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
'Added by Lydia 2015/06/18 +客戶中文名稱查詢
Public m_strCustCode As String
Public m_blnOneRec As Boolean
Dim m_strListPer As String 'Add By Sindy 2020/7/1
Dim arrID 'Add By Sindy 2022/2/21


Private Sub SetDataListWidth()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
'0.V
'1.所別
'2.部門
'3.收文日
'4.本所案號
'5.案件性質
'6.申請國家(4個中文)
'7.智權人員
'8.承辦人
'9.管制期限
'10.本所期限
'11.法定期限
'12.預定會稿日
'13.備註
'14.總收文號
'15.sc07
'16.sc02
'17.sc03
'18.sc08
'19.s1st06
'20.cp12
'21.cp13
'22.sc05
'23.cp09
'Modify By Sindy 2014/6/23 Grid中於"管制期限"欄後增加"自", 若SC07='3'的資料, 此欄顯示"自"
'Modify By Sindy 2024/12/31 + 最近歷程
arrGridHeadText = Array("V", "所別", "部門", "收文日", "本所案號" _
                  , "案件性質", "案件名稱", "申請國家", "智權人員", "承辦人", "管制期限", "自" _
                  , "本所期限", "法定期限", "預定會稿日", "最近歷程", "備註", "總收文號", "sc07" _
                  , "sc02", "sc03", "sc08", "s1st06", "cp12", "cp13", "sc05", "cp09")
arrGridHeadWidth = Array(200, 0, 0, 850, 1200 _
                     , 850, 1250, 850, 680, 680, 850, 250 _
                     , 850, 850, 850, 900, 2000, 950, 0 _
                     , 0, 0, 0, 0, 0, 0, 0, 0)
grdDataList.MergeCells = flexMergeRestrictColumns
grdDataList.Cols = UBound(arrGridHeadText) + 1
For iRow = 0 To grdDataList.Cols - 1
   grdDataList.row = 0
   grdDataList.col = iRow
   grdDataList.Text = arrGridHeadText(iRow)
   grdDataList.ColWidth(iRow) = arrGridHeadWidth(iRow)
   grdDataList.CellAlignment = flexAlignLeftCenter
Next
End Sub

Private Function ConstrainCheck() As Boolean
Dim bolCancel As Boolean
Dim intErrCol As Integer
   
   ConstrainCheck = True
   
   'Add By Sindy 2022/2/21 +有下拉選單
   If Combo3.Visible = True Then
      Call Combo3_LostFocus '讓人員按Enter,須再啟動此函數,txtSales欄位值才會置換
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
   '2022/2/21 END
   
   Call txtSales_Validate(bolCancel)
   If bolCancel = True Then
'      txtSales.SetFocus
'      txtSales_GotFocus
      'Add By Sindy 2022/2/21
      '有下拉選單
      If Combo3.Visible = True Then
         Combo3.SetFocus
      '排除隱藏
      'ElseIf txtSales.Enabled = True Then
      ElseIf txtSales.Enabled = True And txtSales.Visible = True Then
         txtSales.SetFocus
         txtSales_GotFocus
      End If
      '2022/2/21 END
      ConstrainCheck = False
      Exit Function
   End If
   
   If Text1(0) <> "" And Text1(1) <> "" Then
      If Text1(2) = "" Then Text1(2) = "0"
      If Text1(3) = "" Then Text1(3) = "00"
   End If
   
   If Option1(0).Value = True Then
      If txtControlDate(0) = "" Then
         MsgBox "請輸入管制期限(起)！", vbExclamation
         txtControlDate(0).SetFocus
         txtControlDate_GotFocus (0)
         ConstrainCheck = False
         Exit Function
      Else
         bolCancel = False
         Call txtControlDate_Validate(0, bolCancel)
         If bolCancel = True Then
            ConstrainCheck = False
            Exit Function
         End If
      End If
      If txtControlDate(1) = "" Then
         MsgBox "請輸入管制期限(迄)！", vbExclamation
         txtControlDate(1).SetFocus
         txtControlDate_GotFocus (1)
         ConstrainCheck = False
         Exit Function
      Else
         bolCancel = False
         Call txtControlDate_Validate(1, bolCancel)
         If bolCancel = True Then
            ConstrainCheck = False
            Exit Function
         End If
      End If
   End If
   If Option1(2).Value = True Then
      If txtRecvDate(0) = "" Then
         MsgBox "請輸入收文日期(起)！", vbExclamation
         txtRecvDate(0).SetFocus
         txtRecvDate_GotFocus (0)
         ConstrainCheck = False
         Exit Function
      Else
         bolCancel = False
         Call txtRecvDate_Validate(0, bolCancel)
         If bolCancel = True Then
            ConstrainCheck = False
            Exit Function
         End If
      End If
      If txtRecvDate(1) = "" Then
         MsgBox "請輸入收文日期(迄)！", vbExclamation
         txtRecvDate(1).SetFocus
         txtRecvDate_GotFocus (1)
         ConstrainCheck = False
         Exit Function
      Else
         bolCancel = False
         Call txtRecvDate_Validate(1, bolCancel)
         If bolCancel = True Then
            ConstrainCheck = False
            Exit Function
         End If
      End If
   End If
   
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
   
'   '林永生71003檢查業務區範圍
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
'      If txtSalesArea > txtSalesArea1 Then
'         MsgBox "業務區範圍條件錯誤！", vbExclamation
'         txtSalesArea.SetFocus
'         txtSalesArea_GotFocus
'         ConstrainCheck = False
'         Exit Function
'      End If
'   End If
'
'   '簡金泉69005檢查業務區範圍
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
''      If txtSalesArea > txtSalesArea1 Then
''         MsgBox "業務區範圍條件錯誤！", vbExclamation
''         txtSalesArea.SetFocus
''         txtSalesArea_GotFocus
''         ConstrainCheck = False
''         Exit Function
''      End If
''   End If
''end 2019/12/30
'
'   'Added by Lydia 2020/07/02 柄佑82026可看中所全部或自已
'   If strUserNum = "82026" Then
'      If txtSalesArea = "" Then
'         txtSalesArea = "S21"
'         txtSalesArea1 = "S29"
'      End If
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
'   'end 2020/07/02
'
'   '加入外商主管  可以輸入相同組別的
'   If (stST05 = "21" Or stST05 = "26" Or stST05 = "28") Then
'      If Trim(txtSales) = "" Then
'         MsgBox "智權人員不可以空白！", vbExclamation, "操作錯誤！"
'         txtSales.SetFocus
'         txtSales_GotFocus
'         ConstrainCheck = False
'         Exit Function
'      End If
'      If PUB_GetStaffST16(strUserNum) <> PUB_GetStaffST16(txtSales) Then
'         MsgBox "僅可以輸入相同組別的智權人員！", vbExclamation, "操作錯誤！"
'         txtSales.SetFocus
'         txtSales_GotFocus
'         ConstrainCheck = False
'         Exit Function
'      End If
'   End If
   
   '申請人
   If Trim(txtCU1) <> "" Or Trim(txtCU2) <> "" Then
      If Mid(txtCU1, 1, 6) <> Mid(txtCU2, 1, 6) Then
          MsgBox "申請人前6碼必須相同！", vbExclamation
          txtCU1.SetFocus
          txtCU1_GotFocus
          ConstrainCheck = False
          Exit Function
      End If
   End If
End Function

Public Function doQuery() As Boolean
Dim stCon As String, stConIns As String
Dim stConP As String, stConT As String, stConS As String, stConL As String, stConH As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strInData As String
Dim stIdList As String, stConId As String
Dim bolChgSystemkind As Boolean, strOldSystemkind As String
   
'On Error GoTo ErrHnd
   
   Screen.MousePointer = vbHourglass 'Add By Sindy 2015/3/20
   grdDataList.MousePointer = flexHourglass 'Add By Sindy 2015/3/20
   LblCntTime.Caption = "執行時間：" & Format(ServerTime, "##:##:##") 'Add By Sindy 2014/6/30
   
   stCon = "": stConIns = "": stConP = "": stConT = "": stConS = "": stConL = "": stConH = ""
   bolChgSystemkind = False
   If Text1(0) = "" Or Text1(1) = "" Then
      If systemkind = "" Then
          systemkind = "ALL"
      End If
   Else
      bolChgSystemkind = True
      strOldSystemkind = systemkind
      systemkind = Text1(0).Text
   End If
   '陳經理查詢所有智權人員要控制系統類別
   If strUserNum = "68005" And txtSales <> "68005" Then
      systemkind = "CFT,FCT,S,CFC"
   End If
   
'cancel by sonia 2014/6/9
'   '蔣律師要控制所別
'   If strUserNum = "79037" Then
'      stCon = stCon & " and s1.st06 = '" & pub_strUserOffice & "'"
'      stConIns = stConIns & " and s1.st06 = '" & pub_strUserOffice & "'"
'   End If
'end 2014/6/9
   
   '區別
   '若智權人員為80030時, 不限制區別
   '若為帶人主管權限時,查詢之智權人員編號非本人時,不限制區別
   '加巨京專利給郭雅娟79075看,所以不限制區別
   If txtSales = "80030" Or txtSales = "79075" Or _
      (Trim(txtSales) <> "" And PUB_GetST05(strUserNum) = "SA" And txtSales.Enabled = True And txtSales <> strUserNum) Then
      '不限制區別
   'Add by Amy 2014/05/20
   'Modify by Amy 2019/02/12 總經理業務工作代理人員,可處理總經理員工編號
   ElseIf bolSpecMan = True And (InStr(strSpecCode, "A8") > 0 Or InStr(strSpecCode, "總經理業務工作代理人員") > 0) And txtSales <> strUserNum Then
            '開放專利處部份智權同仁資料給彥葶代為處理,不考慮業務區(因彥葶與開放的智權同仁業務區不同)
   'end 2014/05/20
   Else
      If txtSalesArea <> "" Then
         stCon = stCon & " and s1.st15>='" & txtSalesArea & "'" 'Modify By Sindy 2021/8/4 cp12 => s1.st15
         stConIns = stConIns & " and s1.st15>='" & txtSalesArea & "'" 'Modify By Sindy 2021/8/4 cp12 => s1.st15
      End If
      If txtSalesArea1 <> "" Then
         stCon = stCon & " and s1.st15<='" & txtSalesArea1 & "'" 'Modify By Sindy 2021/8/4 cp12 => s1.st15
         stConIns = stConIns & " and s1.st15<='" & txtSalesArea1 & "'" 'Modify By Sindy 2021/8/4 cp12 => s1.st15
      End If
   End If
   
   '智權人員
   If txtSales <> "" Then
        If (strUserNum <> "80030" And txtSales <> "80030") Then
            'Modify by Amy 2014/05/20 +if
            'Modify by Amy 2019/02/12 總經理業務工作代理人員,可處理總經理員工編號
            If bolSpecMan = True And (InStr(strSpecCode, "A8") > 0 Or InStr(strSpecCode, "總經理業務工作代理人員") > 0) And txtSales <> strUserNum Then
                '開放專利處部份智權同仁資料給彥葶代為處理,不考慮業務區(因彥葶與開放的智權同仁業務區不同)
                stIdList = PUB_GetSalesList(txtSales)
            Else
                stIdList = PUB_GetSalesList(txtSales, txtSalesArea, txtSalesArea1, PUB_GetST06(Trim(txtSales)))
            End If
            'end 2014/05/20
            
            '若不是多員工編號時用 = 算符合加速查詢
            If InStr(stIdList, ",") = 0 Then
               stConId = " = " & stIdList & " "
            Else
               stConId = " in (" & stIdList & " ) "
            End If
            stCon = stCon & " and cp13 " & stConId
            stConIns = stConIns & " and cp13 " & stConId
        Else
            '查87027陳淑芳時同時查20001台中所
            '查80030洪琬姿時同時查F4103
            If txtSales = "80030" Then
               StrSQLa = "select ST01 from STAFF where ST04<>'1' and ST03 like 'F1%' "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               strInData = "'80030','F4103'"
               If rsA.RecordCount > 0 Then
                  rsA.MoveFirst
                  Do While rsA.EOF = False
                     strInData = strInData & ",'" & rsA.Fields(0).Value & "'"
                     rsA.MoveNext
                  Loop
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               stCon = stCon & " and cp13 IN (" & strInData & ") "
               stConIns = stConIns & " and cp13 IN (" & strInData & ") "
            Else
               stCon = stCon & " and cp13='" & txtSales & "'"
               stConIns = stConIns & " and cp13='" & txtSales & "'"
            End If
        End If
   'Modify by Amy 2014/05/20
   '智權人員 為空
   Else
        If bolSpecMan = True And InStr(strSpecCode, "A8") > 0 Then
            'A2023彥葶登入,未輸智權人員-設定查A7人員
            stConId = " in ('" & Replace(Pub_GetSpecMan("A7"), ";", "','") & "') "
            stCon = stCon & " and cp13 " & stConId
            stConIns = stConIns & " and cp13 " & stConId
        End If
   'end 2014/05/20
   End If
   
   '申請人
   If Trim(txtCU1) <> "" Then
       txtCU1 = Mid(txtCU1 & "000000000", 1, 9)
       txtCU2 = Mid(txtCU2 & "000000000", 1, 9)
       stConP = stConP & " and ((pa26>='" & txtCU1 & "' and pa26<='" & txtCU2 & "') or (pa27>='" & txtCU1 & "' and pa27<='" & txtCU2 & "') or (pa28>='" & txtCU1 & "' and pa28<='" & txtCU2 & "') or (pa29>='" & txtCU1 & "' and pa29<='" & txtCU2 & "') or (pa30>='" & txtCU1 & "' and pa30<='" & txtCU2 & "')) "
       stConT = stConT & " and ((tm23>='" & txtCU1 & "' and tm23<='" & txtCU2 & "') or (tm78>='" & txtCU1 & "' and tm78<='" & txtCU2 & "') or (tm79>='" & txtCU1 & "' and tm79<='" & txtCU2 & "') or (tm80>='" & txtCU1 & "' and tm80<='" & txtCU2 & "') or (tm81>='" & txtCU1 & "' and tm81<='" & txtCU2 & "')) "
       stConS = stConS & " and ((sp08>='" & txtCU1 & "' and sp08<='" & txtCU2 & "') or (sp58>='" & txtCU1 & "' and sp58<='" & txtCU2 & "') or (sp59>='" & txtCU1 & "' and sp59<='" & txtCU2 & "') or (sp65>='" & txtCU1 & "' and sp65<='" & txtCU2 & "') or (sp66>='" & txtCU1 & "' and sp66<='" & txtCU2 & "')) "
       'Modify By Sindy 2011/2/21 增加LC43,LC44,LC45,LC46
       stConL = stConL & " and ((lc11>='" & txtCU1 & "' and lc11<='" & txtCU2 & "') or (lc43>='" & txtCU1 & "' and lc43<='" & txtCU2 & "') or (lc44>='" & txtCU1 & "' and lc44<='" & txtCU2 & "') or (lc45>='" & txtCU1 & "' and lc45<='" & txtCU2 & "') or (lc46>='" & txtCU1 & "' and lc46<='" & txtCU2 & "')) "
       'Modify By Sindy 2011/2/21 增加HC24,HC25,HC26,HC27
       stConH = stConH & " and ((hc05>='" & txtCU1 & "' and hc05<='" & txtCU2 & "') or (hc24>='" & txtCU1 & "' and hc24<='" & txtCU2 & "') or (hc25>='" & txtCU1 & "' and hc25<='" & txtCU2 & "') or (hc26>='" & txtCU1 & "' and hc26<='" & txtCU2 & "') or (hc27>='" & txtCU1 & "' and hc27<='" & txtCU2 & "')) "
   End If
   
   '案件性質
   If Trim(txtCP10(0)) <> "" Then
       stCon = stCon & " and cp10>='" & txtCP10(0) & "' "
       stConIns = stConIns & " and cp10>='" & txtCP10(0) & "' "
   End If
   If Trim(txtCP10(1)) <> "" Then
       stCon = stCon & " and cp10<='" & txtCP10(1) & "' "
       stConIns = stConIns & " and cp10<='" & txtCP10(1) & "' "
   End If
   
   '本所案號
   If Text1(0) <> "" And Text1(1) <> "" Then
      stCon = stCon & " and (cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "') "
      stConIns = stConIns & " and (cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "') "
   End If
   
   '管制期限(且未完成)
   If Option1(0).Value = True Then
      stCon = stCon & " and (sc05<=" & strSrvDate(1) & " or (sc05>=" & ChangeTStringToWString(txtControlDate(0)) & " and sc05<=" & ChangeTStringToWString(txtControlDate(1)) & "))"
      stCon = stCon & " and nvl(sc08,'0')='0'"
   Else
      '所有未發文資料,帶出未完成的管制期限資料,但若已完成且未發文未取消的總收文號資料一併帶出
      stCon = stCon & " and (nvl(sc08,'0')='0' or (sc05<=" & strSrvDate(1) & " and nvl(sc08,'0')='0') or exists (select * from salescontroldate where sc01=cp09 and sc08='Y' and not exists (select * from salescontroldate where sc01=cp09 and nvl(sc08,'0')='0') and rownum=1)) "
      '收文日期(未發文資料)
      If Option1(2).Value = True Then
         If txtRecvDate(0) <> "" Then
            stCon = stCon & " and cp05>=" & ChangeTStringToWString(txtRecvDate(0))
            stConIns = stConIns & " and cp05>=" & ChangeTStringToWString(txtRecvDate(0))
         End If
         If txtRecvDate(1) <> "" Then
            stCon = stCon & " and cp05<=" & ChangeTStringToWString(txtRecvDate(1))
            stConIns = stConIns & " and cp05<=" & ChangeTStringToWString(txtRecvDate(1))
         End If
      End If
   End If
   
   '若有符合條件的總收文號資料,且不存在於salescontroldate則自動新增一筆記錄
   Call InsertSalesControlDate(stConIns, stConT, stConP, stConS, stConH, stConL)
'   LblCntTime.Caption = LblCntTime.Caption & " ~ " & Format(ServerTime, "##:##:##") 'Add By Sindy 2014/6/30
   '查詢SQL
   '2010/9/28 modify by sonia 剔除LA的顧問聘任資料
   'Modify By Sindy 2014/6/23 Grid中於"管制期限"欄後增加"自", 若SC07='3'的資料, 此欄顯示"自"
   'Modify By Sindy 2016/9/5 and cp57 is null and cp27 is null ==> and cp158=0 and cp159=0
'   strSql = "select '' as V,decode(s1.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 部門,substrb(' '||sqldatet(cp05),-9) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,NVL(DECODE(TM10,'000',CPM03,CPM04),cp10) as 案件性質,nvl(tm05,'')||nvl(tm06,'')||nvl(tm07,'') 案件名稱,na03 申請國家,s1.ST02 as 智權人員,s2.ST02 as 承辦人,decode(sc08,'Y','',substrb(' '||sqldatet(sc05),-9)) as 管制期限,decode(SC07,'3','自','') 自,substrb(' '||sqldatet(cp06),-9) as 本所期限,substrb(' '||sqldatet(cp07),-9) as 法定期限,substrb(' '||sqldatet(ep28),-9) as 預定會稿日,decode(sc08,'Y','',sc06) as 備註,cp09 as 總收文號,decode(sc08,'Y','',sc07),decode(sc08,'Y','',sc02),decode(sc08,'Y','',sc03),sc08,s1.st06 as s1st06,cp12,cp13,decode(sc08,'Y','',sc05) as sc05,cp09" & _
'                "  from (select /*+index(caseprogress IDXCP010510132757)*/ cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp12,cp13,cp14,cp27,cp57 from caseprogress where cp05>=20100901 and cp09<'B' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ") and cp158=0 and cp159=0),staff s1,staff s2,acc090,casepropertymap,trademark,salescontroldate,nation,engineerprogress" & _
'                " where cp13=s1.st01(+) and cp14=s2.st01(+) and cp12=a0901(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp09=sc01(+) and cp09=ep02(+) and TM10=na01(+) " & stCon & stConT & _
'                " union " & _
'                "select '' as V,decode(s1.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 部門,substrb(' '||sqldatet(cp05),-9) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,NVL(DECODE(PA09,'000',CPM03,CPM04),cp10) as 案件性質,nvl(pa05,'')||nvl(pa06,'')||nvl(pa07,'') 案件名稱,na03 申請國家,s1.ST02 as 智權人員,s2.ST02 as 承辦人,decode(sc08,'Y','',substrb(' '||sqldatet(sc05),-9)) as 管制期限,decode(SC07,'3','自','') 自,substrb(' '||sqldatet(cp06),-9) as 本所期限,substrb(' '||sqldatet(cp07),-9) as 法定期限,substrb(' '||sqldatet(ep28),-9) as 預定會稿日,decode(sc08,'Y','',sc06) as 備註,cp09 as 總收文號,decode(sc08,'Y','',sc07),decode(sc08,'Y','',sc02),decode(sc08,'Y','',sc03),sc08,s1.st06 as s1st06,cp12,cp13,decode(sc08,'Y','',sc05) as sc05,cp09" & _
'                "  from (select /*+index(caseprogress IDXCP010510132757)*/ cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp12,cp13,cp14,cp27,cp57 from caseprogress where cp05>=20100901 and cp09<'B' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ") and cp158=0 and cp159=0),staff s1,staff s2,acc090,casepropertymap,patent,salescontroldate,nation,engineerprogress" & _
'                " where cp13=s1.st01(+) and cp14=s2.st01(+) and cp12=a0901(+) and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp09=sc01(+) and cp09=ep02(+) and PA09=na01(+) " & stCon & stConP & _
'                " union " & _
'                "select '' as V,decode(s1.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 部門,substrb(' '||sqldatet(cp05),-9) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,NVL(DECODE(SP09,'000',CPM03,CPM04),cp10) as 案件性質,nvl(sp05,'')||nvl(sp06,'')||nvl(sp07,'') 案件名稱,na03 申請國家,s1.ST02 as 智權人員,s2.ST02 as 承辦人,decode(sc08,'Y','',substrb(' '||sqldatet(sc05),-9)) as 管制期限,decode(SC07,'3','自','') 自,substrb(' '||sqldatet(cp06),-9) as 本所期限,substrb(' '||sqldatet(cp07),-9) as 法定期限,substrb(' '||sqldatet(ep28),-9) as 預定會稿日,decode(sc08,'Y','',sc06) as 備註,cp09 as 總收文號,decode(sc08,'Y','',sc07),decode(sc08,'Y','',sc02),decode(sc08,'Y','',sc03),sc08,s1.st06 as s1st06,cp12,cp13,decode(sc08,'Y','',sc05) as sc05,cp09" & _
'                "  from (select /*+index(caseprogress IDXCP010510132757)*/ cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp12,cp13,cp14,cp27,cp57 from caseprogress where cp05>=20100901 and cp09<'B' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 5) & ") and cp158=0 and cp159=0),staff s1,staff s2,acc090,casepropertymap,servicepractice,salescontroldate,nation,engineerprogress" & _
'                " where cp13=s1.st01(+) and cp14=s2.st01(+) and cp12=a0901(+) and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp09=sc01(+) and cp09=ep02(+) and SP09=na01(+) " & stCon & stConS & _
'                " union " & _
'                "select '' as V,decode(s1.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 部門,substrb(' '||sqldatet(cp05),-9) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,NVL(CPM03,cp10) as 案件性質,hc06 案件名稱,'' 申請國家,s1.ST02 as 智權人員,s2.ST02 as 承辦人,decode(sc08,'Y','',substrb(' '||sqldatet(sc05),-9)) as 管制期限,decode(SC07,'3','自','') 自,substrb(' '||sqldatet(cp06),-9) as 本所期限,substrb(' '||sqldatet(cp07),-9) as 法定期限,substrb(' '||sqldatet(ep28),-9) as 預定會稿日,decode(sc08,'Y','',sc06) as 備註,cp09 as 總收文號,decode(sc08,'Y','',sc07),decode(sc08,'Y','',sc02),decode(sc08,'Y','',sc03),sc08,s1.st06 as s1st06,cp12,cp13,decode(sc08,'Y','',sc05) as sc05,cp09" & _
'                "  from (select /*+index(caseprogress IDXCP010510132757)*/ cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp12,cp13,cp14,cp27,cp57 from caseprogress where cp10<>'0' and cp05>=20100901 and cp09<'B' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 4) & ") and cp158=0 and cp159=0),staff s1,staff s2,acc090,casepropertymap,hirecase,salescontroldate,engineerprogress" & _
'                " where cp13=s1.st01(+) and cp14=s2.st01(+) and cp12=a0901(+) and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp09=sc01(+) and cp09=ep02(+) " & stCon & stConH & _
'                " union " & _
'                "select '' as V,decode(s1.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 部門,substrb(' '||sqldatet(cp05),-9) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,NVL(DECODE(lc15,'000',CPM03,CPM04),cp10) as 案件性質,nvl(lc05,'')||nvl(lc06,'')||nvl(lc07,'') 案件名稱,na03 申請國家,s1.ST02 as 智權人員,s2.ST02 as 承辦人,decode(sc08,'Y','',substrb(' '||sqldatet(sc05),-9)) as 管制期限,decode(SC07,'3','自','') 自,substrb(' '||sqldatet(cp06),-9) as 本所期限,substrb(' '||sqldatet(cp07),-9) as 法定期限,substrb(' '||sqldatet(ep28),-9) as 預定會稿日,decode(sc08,'Y','',sc06) as 備註,cp09 as 總收文號,decode(sc08,'Y','',sc07),decode(sc08,'Y','',sc02),decode(sc08,'Y','',sc03),sc08,s1.st06 as s1st06,cp12,cp13,decode(sc08,'Y','',sc05) as sc05,cp09" & _
'                "  from (select /*+index(caseprogress IDXCP010510132757)*/ cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp12,cp13,cp14,cp27,cp57 from caseprogress where cp05>=20100901 and cp09<'B' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 3) & ") and cp158=0 and cp159=0),staff s1,staff s2,acc090,casepropertymap,lawcase,salescontroldate,nation,engineerprogress" & _
'                " where cp13=s1.st01(+) and cp14=s2.st01(+) and cp12=a0901(+) and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp09=sc01(+) and cp09=ep02(+) and LC15=na01(+) " & stCon & stConL
   'Modify By Sindy 2017/10/26 調整SQL
   'Modify By Sindy 2024/12/31 + 最近歷程
   strSql = "select '' as V,decode(s1.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 部門,substrb(' '||sqldatet(cp05),-9) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,NVL(DECODE(TM10,'000',CPM03,CPM04),cp10) as 案件性質,nvl(tm05,'')||nvl(tm06,'')||nvl(tm07,'') 案件名稱,na03 申請國家,s1.ST02 as 智權人員,s2.ST02 as 承辦人,decode(sc08,'Y','',substrb(' '||sqldatet(sc05),-9)) as 管制期限,decode(SC07,'3','自','') 自,substrb(' '||sqldatet(cp06),-9) as 本所期限,substrb(' '||sqldatet(cp07),-9) as 法定期限,substrb(' '||sqldatet(ep28),-9) as 預定會稿日,GetEEPCurState(cp09) as 最近歷程,decode(sc08,'Y','',sc06) as 備註,cp09 as 總收文號,decode(sc08,'Y','',sc07),decode(sc08,'Y','',sc02),decode(sc08,'Y','',sc03),sc08,s1.st06 as s1st06,cp12,cp13,decode(sc08,'Y','',sc05) as sc05,cp09" & _
            " from caseprogress,staff s1,staff s2,acc090,casepropertymap,trademark,salescontroldate,nation,engineerprogress" & _
            " where cp05>=20100901 and cp09<'B' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ") and cp158=0 and cp159=0" & _
            " and cp13=s1.st01(+) and cp14=s2.st01(+) and cp12=a0901(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp09=sc01(+) and cp09=ep02(+) and TM10=na01(+) " & stCon & stConT & _
            " union " & _
            "select '' as V,decode(s1.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 部門,substrb(' '||sqldatet(cp05),-9) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,NVL(DECODE(PA09,'000',CPM03,CPM04),cp10) as 案件性質,nvl(pa05,'')||nvl(pa06,'')||nvl(pa07,'') 案件名稱,na03 申請國家,s1.ST02 as 智權人員,s2.ST02 as 承辦人,decode(sc08,'Y','',substrb(' '||sqldatet(sc05),-9)) as 管制期限,decode(SC07,'3','自','') 自,substrb(' '||sqldatet(cp06),-9) as 本所期限,substrb(' '||sqldatet(cp07),-9) as 法定期限,substrb(' '||sqldatet(ep28),-9) as 預定會稿日,GetEEPCurState(cp09) as 最近歷程,decode(sc08,'Y','',sc06) as 備註,cp09 as 總收文號,decode(sc08,'Y','',sc07),decode(sc08,'Y','',sc02),decode(sc08,'Y','',sc03),sc08,s1.st06 as s1st06,cp12,cp13,decode(sc08,'Y','',sc05) as sc05,cp09" & _
            " from caseprogress,staff s1,staff s2,acc090,casepropertymap,patent,salescontroldate,nation,engineerprogress" & _
            " where cp05>=20100901 and cp09<'B' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ") and cp158=0 and cp159=0" & _
            " and cp13=s1.st01(+) and cp14=s2.st01(+) and cp12=a0901(+) and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp09=sc01(+) and cp09=ep02(+) and PA09=na01(+) " & stCon & stConP & _
            " union " & _
            "select '' as V,decode(s1.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 部門,substrb(' '||sqldatet(cp05),-9) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,NVL(DECODE(SP09,'000',CPM03,CPM04),cp10) as 案件性質,nvl(sp05,'')||nvl(sp06,'')||nvl(sp07,'') 案件名稱,na03 申請國家,s1.ST02 as 智權人員,s2.ST02 as 承辦人,decode(sc08,'Y','',substrb(' '||sqldatet(sc05),-9)) as 管制期限,decode(SC07,'3','自','') 自,substrb(' '||sqldatet(cp06),-9) as 本所期限,substrb(' '||sqldatet(cp07),-9) as 法定期限,substrb(' '||sqldatet(ep28),-9) as 預定會稿日,GetEEPCurState(cp09) as 最近歷程,decode(sc08,'Y','',sc06) as 備註,cp09 as 總收文號,decode(sc08,'Y','',sc07),decode(sc08,'Y','',sc02),decode(sc08,'Y','',sc03),sc08,s1.st06 as s1st06,cp12,cp13,decode(sc08,'Y','',sc05) as sc05,cp09" & _
            " from caseprogress,staff s1,staff s2,acc090,casepropertymap,servicepractice,salescontroldate,nation,engineerprogress" & _
            " where cp05>=20100901 and cp09<'B' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 5) & ") and cp158=0 and cp159=0" & _
            " and cp13=s1.st01(+) and cp14=s2.st01(+) and cp12=a0901(+) and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp09=sc01(+) and cp09=ep02(+) and SP09=na01(+) " & stCon & stConS & _
            " union " & _
            "select '' as V,decode(s1.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 部門,substrb(' '||sqldatet(cp05),-9) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,NVL(CPM03,cp10) as 案件性質,hc06 案件名稱,'' 申請國家,s1.ST02 as 智權人員,s2.ST02 as 承辦人,decode(sc08,'Y','',substrb(' '||sqldatet(sc05),-9)) as 管制期限,decode(SC07,'3','自','') 自,substrb(' '||sqldatet(cp06),-9) as 本所期限,substrb(' '||sqldatet(cp07),-9) as 法定期限,substrb(' '||sqldatet(ep28),-9) as 預定會稿日,GetEEPCurState(cp09) as 最近歷程,decode(sc08,'Y','',sc06) as 備註,cp09 as 總收文號,decode(sc08,'Y','',sc07),decode(sc08,'Y','',sc02),decode(sc08,'Y','',sc03),sc08,s1.st06 as s1st06,cp12,cp13,decode(sc08,'Y','',sc05) as sc05,cp09" & _
            " from caseprogress,staff s1,staff s2,acc090,casepropertymap,hirecase,salescontroldate,engineerprogress" & _
            " where cp10<>'0' and cp05>=20100901 and cp09<'B' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 4) & ") and cp158=0 and cp159=0" & _
            " and cp13=s1.st01(+) and cp14=s2.st01(+) and cp12=a0901(+) and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp09=sc01(+) and cp09=ep02(+) " & stCon & stConH & _
            " union " & _
            "select '' as V,decode(s1.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 部門,substrb(' '||sqldatet(cp05),-9) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,NVL(DECODE(lc15,'000',CPM03,CPM04),cp10) as 案件性質,nvl(lc05,'')||nvl(lc06,'')||nvl(lc07,'') 案件名稱,na03 申請國家,s1.ST02 as 智權人員,s2.ST02 as 承辦人,decode(sc08,'Y','',substrb(' '||sqldatet(sc05),-9)) as 管制期限,decode(SC07,'3','自','') 自,substrb(' '||sqldatet(cp06),-9) as 本所期限,substrb(' '||sqldatet(cp07),-9) as 法定期限,substrb(' '||sqldatet(ep28),-9) as 預定會稿日,GetEEPCurState(cp09) as 最近歷程,decode(sc08,'Y','',sc06) as 備註,cp09 as 總收文號,decode(sc08,'Y','',sc07),decode(sc08,'Y','',sc02),decode(sc08,'Y','',sc03),sc08,s1.st06 as s1st06,cp12,cp13,decode(sc08,'Y','',sc05) as sc05,cp09" & _
            " from caseprogress,staff s1,staff s2,acc090,casepropertymap,lawcase,salescontroldate,nation,engineerprogress" & _
            " where cp05>=20100901 and cp09<'B' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 3) & ") and cp158=0 and cp159=0" & _
            " and cp13=s1.st01(+) and cp14=s2.st01(+) and cp12=a0901(+) and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp09=sc01(+) and cp09=ep02(+) and LC15=na01(+) " & stCon & stConL
   If strOrderBy <> "" Then
      strSql = strSql & " order by " & strOrderBy
   End If
   CheckOC3
   grdDataList.Rows = 2
   grdDataList.Clear
   SetDataListWidth
   grdDataList.FixedCols = 0
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      LblCntTime.Caption = LblCntTime.Caption & " ~ " & Format(ServerTime, "##:##:##") & " 共 " & .RecordCount & " 筆" 'Add By Sindy 2014/6/30
      If .RecordCount > 0 Then
         Set grdDataList.Recordset = AdoRecordSet3.Clone
         SetDataListWidth
         grdDataList.FixedCols = 7
         Call SetColor
         Label3.Caption = "PS：請自行點選資料排序條件（點選該欄位標題） 共 " & .RecordCount & " 筆"
      Else
         Label3.Caption = "PS：請自行點選資料排序條件（點選該欄位標題） 共 0 筆"
         If bolShowMsgBox = True Then
            MsgBox "無符合資料！", vbInformation
         End If
      End If
   End With
   
   grdDataList.MousePointer = flexDefault 'Add By Sindy 2015/3/20
   Screen.MousePointer = vbDefault 'Add By Sindy 2015/3/20
   doQuery = True
   If bolChgSystemkind = True Then systemkind = strOldSystemkind
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub SetColor()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   grdDataList.Visible = False
   For i = 1 To grdDataList.Rows - 1
      grdDataList.row = i
      '逾管制期限變黃色
      If grdDataList.TextMatrix(i, 10) <> "" And _
         ChangeTStringToWString(ChangeTDateStringToTString(grdDataList.TextMatrix(i, 10))) < strSrvDate(1) Then
         For j = 0 To grdDataList.Cols - 1
            grdDataList.col = j
            grdDataList.CellBackColor = &HFFFF& '黃色
         Next j
      '逾預定會稿日變淺紅
      ElseIf grdDataList.TextMatrix(i, 14) <> "" And _
         ChangeTStringToWString(ChangeTDateStringToTString(grdDataList.TextMatrix(i, 14))) < strSrvDate(1) Then
         '2010/9/28 add by sonia 有會稿日即使逾預定會稿日也不必變淺紅
         StrSQLa = "select ep07 from engineerprogress where ep02='" & grdDataList.TextMatrix(i, 17) & "' "
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            If "" & rsA.Fields(0) = "" Then
         '2010/9/28 end
               For j = 0 To grdDataList.Cols - 1
                  grdDataList.col = j
                  grdDataList.CellBackColor = &H8080FF '淺紅
               Next j
            End If
         End If
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
      End If
   Next i
   grdDataList.Visible = True
End Sub

'若有符合條件的總收文號資料,且不存在於salescontroldate則自動新增一筆記錄
Private Function InsertSalesControlDate(stConIns As String, stConT As String, stConP As String, stConS As String, stConH As String, stConL As String) As Boolean
Dim longSeqno As Long
Dim strSC05 As String
On Error GoTo ErrHnd
   InsertSalesControlDate = True
   'Modify By Sindy 2014/6/30 IDXCP0501092757 ==> IDXCP010510132757
   'Modify By Sindy 2016/9/5 and cp57 is null and cp27 is null ==> and cp158=0 and cp159=0
'   strSql = "select cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp48,ep06,ep07,sc01" & _
'                "  from (select /*+index(caseprogress IDXCP010510132757)*/ cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp12,cp13,cp14,cp27,cp48,cp57 from caseprogress where cp05>=20100901 and cp09<'B' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ") and cp158=0 and cp159=0" & stConIns & "),trademark,engineerprogress,staff s1,salescontroldate" & _
'                " where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp09=ep02(+) and ((ep06 is null and (not exists (select * from salescontroldate where cp09=sc01 and sc07='1'))) or (ep06 is not null and ep07 is null and (not exists (select * from salescontroldate where cp09=sc01 and sc07='2'))) or not exists (select * from salescontroldate where cp09=sc01)) and cp09=sc01(+) and cp13=s1.st01(+) " & stConT & _
'                " union " & _
'                "select cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp48,ep06,ep07,sc01" & _
'                "  from (select /*+index(caseprogress IDXCP010510132757)*/ cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp12,cp13,cp14,cp27,cp48,cp57 from caseprogress where cp05>=20100901 and cp09<'B' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ") and cp158=0 and cp159=0" & stConIns & "),patent,engineerprogress,staff s1,salescontroldate" & _
'                " where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp09=ep02(+) and ((ep06 is null and (not exists (select * from salescontroldate where cp09=sc01 and sc07='1'))) or (ep06 is not null and ep07 is null and (not exists (select * from salescontroldate where cp09=sc01 and sc07='2'))) or not exists (select * from salescontroldate where cp09=sc01)) and cp09=sc01(+) and cp13=s1.st01(+) " & stConP & _
'                " union " & _
'                "select cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp48,ep06,ep07,sc01" & _
'                "  from (select /*+index(caseprogress IDXCP010510132757)*/ cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp12,cp13,cp14,cp27,cp48,cp57 from caseprogress where cp05>=20100901 and cp09<'B' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 5) & ") and cp158=0 and cp159=0" & stConIns & "),servicepractice,engineerprogress,staff s1,salescontroldate" & _
'                " where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and cp09=ep02(+) and ((ep06 is null and (not exists (select * from salescontroldate where cp09=sc01 and sc07='1'))) or (ep06 is not null and ep07 is null and (not exists (select * from salescontroldate where cp09=sc01 and sc07='2'))) or not exists (select * from salescontroldate where cp09=sc01)) and cp09=sc01(+) and cp13=s1.st01(+) " & stConS & _
'                " union " & _
'                "select cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp48,ep06,ep07,sc01" & _
'                "  from (select /*+index(caseprogress IDXCP010510132757)*/ cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp12,cp13,cp14,cp27,cp48,cp57 from caseprogress where cp05>=20100901 and cp09<'B' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 4) & ") and cp158=0 and cp159=0" & stConIns & "),hirecase,engineerprogress,staff s1,salescontroldate" & _
'                " where cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and cp09=ep02(+) and ((ep06 is null and (not exists (select * from salescontroldate where cp09=sc01 and sc07='1'))) or (ep06 is not null and ep07 is null and (not exists (select * from salescontroldate where cp09=sc01 and sc07='2'))) or not exists (select * from salescontroldate where cp09=sc01)) and cp09=sc01(+) and cp13=s1.st01(+) " & stConH & _
'                " union " & _
'                "select cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp48,ep06,ep07,sc01" & _
'                "  from (select /*+index(caseprogress IDXCP010510132757)*/ cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp12,cp13,cp14,cp27,cp48,cp57 from caseprogress where cp05>=20100901 and cp09<'B' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 3) & ") and cp158=0 and cp159=0" & stConIns & "),lawcase,engineerprogress,staff s1,salescontroldate" & _
'                " where cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp09=ep02(+) and ((ep06 is null and (not exists (select * from salescontroldate where cp09=sc01 and sc07='1'))) or (ep06 is not null and ep07 is null and (not exists (select * from salescontroldate where cp09=sc01 and sc07='2'))) or not exists (select * from salescontroldate where cp09=sc01)) and cp09=sc01(+) and cp13=s1.st01(+) " & stConL & _
'                " order by cp01 asc,cp09 asc"
   'Modify By Sindy 2017/10/26 調整SQL
   strSql = "select cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp48,ep06,ep07" & _
            " from caseprogress,trademark,engineerprogress,staff s1" & _
            " where cp05>=20100901 and cp09<'B' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ") and cp158=0 and cp159=0" & stConIns & _
            " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp09=ep02(+) and ((nvl(ep06,0)=0 and not exists (select * from salescontroldate where cp09=sc01(+) and sc07='1')) or (nvl(ep06,0)>0 and nvl(ep07,0)=0 and not exists (select * from salescontroldate where cp09=sc01(+) and sc07='2')) or not exists (select * from salescontroldate where cp09=sc01(+))) and cp13=s1.st01(+) " & stConT & _
            " union " & _
            "select cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp48,ep06,ep07" & _
            " from caseprogress,patent,engineerprogress,staff s1" & _
            " where cp05>=20100901 and cp09<'B' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ") and cp158=0 and cp159=0" & stConIns & _
            " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp09=ep02(+) and ((nvl(ep06,0)=0 and not exists (select * from salescontroldate where cp09=sc01(+) and sc07='1')) or (nvl(ep06,0)>0 and nvl(ep07,0)=0 and not exists (select * from salescontroldate where cp09=sc01(+) and sc07='2')) or not exists (select * from salescontroldate where cp09=sc01(+))) and cp13=s1.st01(+) " & stConP & _
            " union " & _
            "select cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp48,ep06,ep07" & _
            " from caseprogress,servicepractice,engineerprogress,staff s1" & _
            " where cp05>=20100901 and cp09<'B' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 5) & ") and cp158=0 and cp159=0" & stConIns & _
            " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and cp09=ep02(+) and ((nvl(ep06,0)=0 and not exists (select * from salescontroldate where cp09=sc01(+) and sc07='1')) or (nvl(ep06,0)>0 and nvl(ep07,0)=0 and not exists (select * from salescontroldate where cp09=sc01(+) and sc07='2')) or not exists (select * from salescontroldate where cp09=sc01(+))) and cp13=s1.st01(+) " & stConS & _
            " union " & _
            "select cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp48,ep06,ep07" & _
            " from caseprogress,hirecase,engineerprogress,staff s1" & _
            " where cp05>=20100901 and cp09<'B' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 4) & ") and cp158=0 and cp159=0" & stConIns & _
            " and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and cp09=ep02(+) and ((nvl(ep06,0)=0 and not exists (select * from salescontroldate where cp09=sc01(+) and sc07='1')) or (nvl(ep06,0)>0 and nvl(ep07,0)=0 and not exists (select * from salescontroldate where cp09=sc01(+) and sc07='2')) or not exists (select * from salescontroldate where cp09=sc01(+))) and cp13=s1.st01(+) " & stConH & _
            " union " & _
            "select cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp48,ep06,ep07" & _
            " from caseprogress,lawcase,engineerprogress,staff s1" & _
            " where cp05>=20100901 and cp09<'B' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 3) & ") and cp158=0 and cp159=0" & stConIns & _
            " and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp09=ep02(+) and ((nvl(ep06,0)=0 and not exists (select * from salescontroldate where cp09=sc01(+) and sc07='1')) or (nvl(ep06,0)>0 and nvl(ep07,0)=0 and not exists (select * from salescontroldate where cp09=sc01(+) and sc07='2')) or not exists (select * from salescontroldate where cp09=sc01(+))) and cp13=s1.st01(+) " & stConL & _
            " order by cp01 asc,cp09 asc"
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         cnnConnection.BeginTrans
         AdoRecordSet3.MoveFirst
         Do While AdoRecordSet3.EOF = False
            '讀取該文號且為系統日的最大序號
            strSql = "SELECT * FROM salescontroldate" & _
                         " WHERE SC01='" & AdoRecordSet3.Fields("cp09") & "' " & _
                                 " and SC02=" & strSrvDate(1) & " " & _
                         " order by SC03 desc "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            longSeqno = 1
            If intI = 1 Then
               If Not IsNull(RsTemp.Fields("SC03")) Then
                  longSeqno = Val(RsTemp.Fields("SC03")) + 1
               End If
            End If
'            '專利未齊備
'            If (AdoRecordSet3.Fields("cp01") = "P" Or AdoRecordSet3.Fields("cp01") = "CFP" Or _
'               AdoRecordSet3.Fields("cp01") = "FCP" Or AdoRecordSet3.Fields("cp01") = "PS" Or _
'               AdoRecordSet3.Fields("cp01") = "CPS" Or AdoRecordSet3.Fields("cp01") = "FG") And _
'               (IsNull(AdoRecordSet3.Fields("ep06")) Or AdoRecordSet3.Fields("ep06") = 0) Then
'               '管制期限=收文日+2天
'               'strSC05 = DBDATE(DateAdd("d", 2, ChangeWStringToWDateString(DBDATE(AdoRecordSet3.Fields("cp05")))))
'               '管制期限=收文日+2天(工作天)
'               strSC05 = CompWorkDay(3, DBDATE(AdoRecordSet3.Fields("cp05")), 0)
'               strSql = "insert into salescontroldate(SC01,SC02,SC03,SC04,SC05,SC06,SC07)" & _
'                            " values('" & AdoRecordSet3.Fields("cp09") & "'," & strSrvDate(1) & "," & longSeqno & ",'M0100'," & strSC05 & ",'待齊備','1') "
'               cnnConnection.Execute strSql
'            '專利已齊備未會稿且有承辦期限
'            ElseIf (AdoRecordSet3.Fields("cp01") = "P" Or AdoRecordSet3.Fields("cp01") = "CFP" Or _
'               AdoRecordSet3.Fields("cp01") = "FCP" Or AdoRecordSet3.Fields("cp01") = "PS" Or _
'               AdoRecordSet3.Fields("cp01") = "CPS" Or AdoRecordSet3.Fields("cp01") = "FG") And _
'               (Not IsNull(AdoRecordSet3.Fields("ep06")) And AdoRecordSet3.Fields("ep06") > 0) And _
'               (IsNull(AdoRecordSet3.Fields("ep07")) Or AdoRecordSet3.Fields("ep07") = 0) And _
'               (Not IsNull(AdoRecordSet3.Fields("cp48")) And AdoRecordSet3.Fields("cp48") > 0) Then
'               '管制期限=承辦期限
'               strSql = "insert into salescontroldate(SC01,SC02,SC03,SC04,SC05,SC06,SC07)" & _
'                            " values('" & AdoRecordSet3.Fields("cp09") & "'," & strSrvDate(1) & "," & longSeqno & ",'M0100'," & AdoRecordSet3.Fields("cp48") & ",'待會稿','2') "
'               cnnConnection.Execute strSql
'            '其他
'            ElseIf "" & AdoRecordSet3.Fields("sc01") = "" Or IsNull("" & AdoRecordSet3.Fields("sc01")) Then
'               '無管制期限無備註
'               strSql = "insert into salescontroldate(SC01,SC02,SC03,SC04,SC07)" & _
'                            " values('" & AdoRecordSet3.Fields("cp09") & "'," & strSrvDate(1) & "," & longSeqno & ",'M0100','3') "
'               cnnConnection.Execute strSql
'            End If
            'Modify By Sindy 2010/10/20
            '專利改為屬於 NewCasePtyList 的案件性質才做待齊備及待會稿的控管
            If (AdoRecordSet3.Fields("cp01") = "P" Or AdoRecordSet3.Fields("cp01") = "CFP" Or _
               AdoRecordSet3.Fields("cp01") = "FCP" Or AdoRecordSet3.Fields("cp01") = "PS" Or _
               AdoRecordSet3.Fields("cp01") = "CPS" Or AdoRecordSet3.Fields("cp01") = "FG") And _
               InStr(NewCasePtyList, AdoRecordSet3.Fields("cp10")) > 0 Then
               '專利未齊備
               If (IsNull(AdoRecordSet3.Fields("ep06")) Or AdoRecordSet3.Fields("ep06") = 0) Then
                  '待齊備管制期限=收文日+2天(工作天)
                  strSC05 = CompWorkDay(3, DBDATE(AdoRecordSet3.Fields("cp05")), 0)
                  '檢查若待齊備管制期限有大於本所期限, 則待齊備管制期限=本所期限
                  If Not IsNull(AdoRecordSet3.Fields("cp06")) Then
                     If Val(strSC05) > Val(AdoRecordSet3.Fields("cp06")) Then
                        strSC05 = AdoRecordSet3.Fields("cp06")
                     End If
                  End If
                  strSql = "insert into salescontroldate(SC01,SC02,SC03,SC04,SC05,SC06,SC07)" & _
                               " values('" & AdoRecordSet3.Fields("cp09") & "'," & strSrvDate(1) & "," & longSeqno & ",'M0100'," & strSC05 & ",'待齊備','1') "
                  cnnConnection.Execute strSql
               '專利已齊備未會稿
               ElseIf (Not IsNull(AdoRecordSet3.Fields("ep06")) And AdoRecordSet3.Fields("ep06") > 0) And _
                         (IsNull(AdoRecordSet3.Fields("ep07")) Or AdoRecordSet3.Fields("ep07") = 0) Then
                  '待會稿管制期限=齊備日+oMan(工作天)
                  '以’NEW’+系統類別抓特殊設定檔 SetSpecMan 的內容(oMan)來設定管制期限
                  strSql = "SELECT OMAN FROM SetSpecMan WHERE OCODE='NEW" & AdoRecordSet3.Fields("cp01") & "' "
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     strSC05 = CompWorkDay(RsTemp.Fields("OMAN"), DBDATE(AdoRecordSet3.Fields("ep06")), 0)
                     '檢查若待會稿管制期限有大於本所期限, 則待會稿管制期限=本所期限
                     If Not IsNull(AdoRecordSet3.Fields("cp06")) Then
                        If Val(strSC05) > Val(AdoRecordSet3.Fields("cp06")) Then
                           strSC05 = AdoRecordSet3.Fields("cp06")
                        End If
                     End If
                     strSql = "insert into salescontroldate(SC01,SC02,SC03,SC04,SC05,SC06,SC07)" & _
                                  " values('" & AdoRecordSet3.Fields("cp09") & "'," & strSrvDate(1) & "," & longSeqno & ",'M0100'," & strSC05 & ",'待會稿','2') "
                     cnnConnection.Execute strSql
                  End If
'               '其他
'               ElseIf "" & AdoRecordSet3.Fields("sc01") = "" Or IsNull("" & AdoRecordSet3.Fields("sc01")) Then
'                  '無管制期限無備註
'                  strSql = "insert into salescontroldate(SC01,SC02,SC03,SC04,SC07)" & _
'                               " values('" & AdoRecordSet3.Fields("cp09") & "'," & strSrvDate(1) & "," & longSeqno & ",'M0100','3') "
'                  cnnConnection.Execute strSql
               End If
            '2010/10/20 End
'            '其他
'            ElseIf "" & AdoRecordSet3.Fields("sc01") = "" Or IsNull("" & AdoRecordSet3.Fields("sc01")) Then
'               '無管制期限無備註
'               strSql = "insert into salescontroldate(SC01,SC02,SC03,SC04,SC07)" & _
'                            " values('" & AdoRecordSet3.Fields("cp09") & "'," & strSrvDate(1) & "," & longSeqno & ",'M0100','3') "
'               cnnConnection.Execute strSql
            End If
            AdoRecordSet3.MoveNext
         Loop
         cnnConnection.CommitTrans
      End If
   End With
   
   Exit Function
   
ErrHnd:
   InsertSalesControlDate = False
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

'管制新期限
Private Sub cmdControlNDate_Click()
Dim intRow As Integer
Dim bolSelData As Boolean 'Add By Sindy 2015/3/20
   
On Error GoTo ErrHnd
   Me.Hide
   bolSelData = False 'Add By Sindy 2015/3/20
   For i = 1 To grdDataList.Rows - 1
      grdDataList.col = 0
      grdDataList.row = i
      If Trim(grdDataList.Text) = "V" Then
         bolSelData = True 'Add By Sindy 2015/3/20
         If Not IsNull(grdDataList.TextMatrix(i, 17)) Then
            frm210134_1.m_strSC01 = Pub_RplStr(grdDataList.TextMatrix(i, 17)) '總收文號
            If Pub_RplStr(grdDataList.TextMatrix(i, 10)) = "" Then
              frm210134_1.m_strSC05 = ""
            Else
              frm210134_1.m_strSC05 = DBDATE(Pub_RplStr(grdDataList.TextMatrix(i, 10))) - 19110000 '管制期限
            End If
            frm210134_1.m_strSC06 = Pub_RplStr(grdDataList.TextMatrix(i, 16)) '備註
            frm210134_1.m_strSC02 = Pub_RplStr(grdDataList.TextMatrix(i, 19))
            frm210134_1.m_strSC03 = Pub_RplStr(grdDataList.TextMatrix(i, 20))
            '開啟視窗
            If frm210134_1.Process() Then
               frm210134_1.Show vbModal
            End If
            Unload frm210134_1
            Set frm210134_1 = Nothing
         End If
      End If
   Next i
   Me.Show
   bolShowMsgBox = False
   If bolSelData = True Then Call doQuery 'Modify By Sindy 2015/3/20
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Sub

'管制完成
Private Sub cmdControlOK_Click()
Dim intRow As Integer
On Error GoTo ErrHnd
   If Val(grdDataList.Rows - 1) = 0 Or (Val(grdDataList.Rows - 1) = 1 And grdDataList.TextMatrix(1, 17) = "") Then Exit Sub
   Screen.MousePointer = vbHourglass
   grdDataList.MousePointer = flexHourglass
   '檢查資料
   bolSelData = False
   For intRow = 1 To grdDataList.Rows - 1
      If grdDataList.TextMatrix(intRow, 0) = "V" Then
         bolSelData = True
         If grdDataList.TextMatrix(intRow, 18) = "1" Or _
            grdDataList.TextMatrix(intRow, 18) = "2" Then
            MsgBox grdDataList.TextMatrix(intRow, 17) & "不可點選系統管制的期限資料，此類期限待文件齊備或會稿後會自動註記為管制完成！", vbExclamation
            grdDataList.MousePointer = flexDefault: Screen.MousePointer = vbDefault
            Exit Sub
         ElseIf grdDataList.TextMatrix(intRow, 10) = "" Then
            MsgBox grdDataList.TextMatrix(intRow, 17) & "無管制期限！", vbExclamation
            grdDataList.MousePointer = flexDefault: Screen.MousePointer = vbDefault
            Exit Sub
         End If
      End If
   Next intRow
   If bolSelData = False Then grdDataList.MousePointer = flexDefault: Screen.MousePointer = vbDefault: Exit Sub
   '開始更新資料
   cnnConnection.BeginTrans
   For intRow = 1 To grdDataList.Rows - 1
      If grdDataList.TextMatrix(intRow, 0) = "V" Then
         strSql = "update salescontroldate" & _
                      " set SC08='Y'" & _
                      " where SC01='" & grdDataList.TextMatrix(intRow, 17) & "'" & _
                          " and SC02=" & grdDataList.TextMatrix(intRow, 19) & _
                          " and SC03=" & grdDataList.TextMatrix(intRow, 20)
         cnnConnection.Execute strSql
      End If
   Next intRow
   MsgBox "管制完成！", vbExclamation
   cnnConnection.CommitTrans
   bolShowMsgBox = False
   Call doQuery
ErrHnd:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
   grdDataList.MousePointer = flexDefault
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Public Sub cmdSearch_Click()
   'Screen.MousePointer = vbHourglass
   'grdDataList.MousePointer = flexHourglass
   strOrderBy = "s1st06 asc,cp12 asc,cp13 asc,sc05 asc,cp09 asc"
   If ConstrainCheck = True Then
      SetDataListWidth
      bolShowMsgBox = True
      Call doQuery
   End If
   'grdDataList.MousePointer = flexDefault
   'Screen.MousePointer = vbDefault
   m_blnColOrderAsc = True    '2010/9/28 ADD BY SONIA
End Sub

'Add By Sindy 2022/2/21
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
Dim stTmp As String
   
   If Combo3 <> "" And Trim(Combo3) <> "全部" Then
      stTmp = Combo3
      '直接輸員編未串名字會錯
      If InStr(stTmp, " ") > 0 Then
        stTmp = Mid(stTmp, 1, Val(InStr(stTmp, " ")) - 1)
      Else
        stTmp = Combo3
      End If
      arrID = Split(Combo3, " ")
      txtSales = arrID(0)
      lblSalesName.Caption = GetStaffName(txtSales)
      If lblSalesName.Caption = "" Then
         MsgBox "智權人員輸入錯誤！", vbCritical
         Combo3.SetFocus
         Cancel = True
      End If
      Combo3 = txtSales & " " & GetPrjSalesNM(txtSales)
   'Modify By Sindy 2024/8/5 mark; 因 txtSales_Validate 會檢查相關的權限
   Else
      txtSales = ""
'   'Modify by Amy 2023/05/09 +stST05
'   ElseIf Combo3 = MsgText(601) And stST05 <> "00" And stST05 <> "01" And stST05 <> "08" Then
'        '下拉選單無區主管智權人員不可為空
'        'Modify By Sindy 2023/9/21 開放杜協理權限 + And Not (Mid(txtSalesArea, 1, 1) = "S" And Mid(txtSalesArea1, 1, 1) = "S" And InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) > 0)
'        If (GetDeptMan(txtSalesArea) <> strUserNum Or GetDeptMan(txtSalesArea1) <> strUserNum) _
'            And Pub_StrUserSt03 <> "M51" _
'            And Not (Mid(txtSalesArea, 1, 1) = "S" And Mid(txtSalesArea1, 1, 1) = "S" And InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) > 0) Then
'           MsgBox "非區主管職代智權人員不可空白！"
'           Cancel = True
'           Combo3.SetFocus
'           Exit Sub
'        End If
   '2024/8/5 END
   End If
End Sub
'2022/2/21 END

Private Sub Form_Activate()
If isLoad = False Then
   MoveFormToCenter Me
   isLoad = True
End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   bolShowMsgBox = False
   
   stST15 = PUB_GetStaffST15(strUserNum, 1)
   stST05 = PUB_GetST05(strUserNum)
   
   txtControlDate(0) = ChangeWDateStringToTString(DateAdd("d", -2, ChangeWStringToWDateString(strSrvDate(1)))) 'strSrvDate(2)
   txtControlDate(1) = ChangeWDateStringToTString(DateAdd("d", 2, ChangeWStringToWDateString(strSrvDate(1))))  'strSrvDate(2)
   Combo1.ListIndex = 4 '3 設n個月 為空白
   Combo2.ListIndex = 4 '3 設n個月 為空白
   
   'Modify By Sindy 2020/7/28 設定員編,部門,所別權限
   'Modify By Sindy 2023/5/18 +, , , , , , True, txtControlDate(0)
   Call PUB_SetFormSaleDept(strUserNum, , txtSalesArea, txtSalesArea1, txtSales, bolSpecMan, strSpecCode, , , , , , True, txtControlDate(0))
   
   'Add By Sindy 2022/2/21
   '檢查當時是否需要為他人職代
   Combo3.Clear
   'Add By Sindy 2023/5/18
   If txtSales <> strUserNum And txtSales <> "" Then
      Combo3.AddItem txtSales & " " & GetPrjSalesNM(txtSales)
   End If
   '2023/5/18 END
   Combo3.AddItem strUserNum & " " & strUserName
   Call Pub_SetForOthersEmpCombo(strUserNum, Combo3, False, m_strListPer)
   If m_strListPer = "" Then
      Combo3.Visible = False
   Else
'      'Add by Amy 2020/03/25 判斷下拉選單是否有區主管
'      If InStr(m_strListPer, GetDeptMan(stST15)) > 0 Then
'         bolAreaMan = True
'      End If
      Combo3.Visible = True
      Combo3.ListIndex = 0
      'Form 2.0物件無法覆蓋Form 1.0
      txtSales.Visible = False
      lblSalesName.Visible = False
   End If
   '2022/2/21 END
   
'   txtSalesArea.Enabled = False
'   txtSalesArea1.Enabled = False
'   txtSales.Enabled = False
'   Select Case strUserNum
'      '外商陳經理可看全所CFT,FCT,S,CFC
'      Case "68005"
'         txtSalesArea.Enabled = True
'         txtSalesArea1.Enabled = True
'         txtSales.Enabled = True
'         txtSalesArea = stST15
'         txtSalesArea1 = stST15
''cancel by sonia 2014/6/9
''      '蔣律師可看中所全部
''      Case "79037"
''         txtSalesArea.Enabled = True
''         txtSalesArea1.Enabled = True
''         txtSales.Enabled = True
''         txtSalesArea = "S2"
''         txtSalesArea1 = "S29"
''end 2014/6/9
'      '小真,杜副總可看全部
'      'modify by sonia 2014/6/9 +美珍77027
'      'Modify by Amy 2015/02/03 拿掉美珍 改寫至特殊人員(總經理業務工作代理人員)
'      Case "65001", "68006"
'         txtSalesArea.Enabled = True
'         txtSalesArea1.Enabled = True
'         txtSales.Enabled = True
'         '副總預設所有智權人員
'         If strUserNum = "68006" Then
'            txtSalesArea = "S"
'            txtSalesArea1 = "S99"
'         End If
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
'      'Added by Lydia 2020/07/02 柄佑可看中所全部但業務區仍預設自已部門
'      Case "82026"
'         txtSalesArea.Enabled = True
'         txtSalesArea1.Enabled = True
'         txtSales.Enabled = True
'         txtSalesArea = GetST15(strUserNum)
'         txtSalesArea1 = txtSalesArea
'      'end 2002/07/02
'      Case Else
'         Select Case stST05
'            '電腦中心,財務,總經理看全部
'            '2015/7/28 MODIFY BY SONIA +主任秘書(等級08)可看全部
'            Case "00", "01", "08"
'               txtSalesArea.Enabled = True
'               txtSalesArea1.Enabled = True
'               txtSales.Enabled = True
'            '分所財務人員可看該所全部
''            Case "C1", "NM", "KM"
''               txtSalesArea.Enabled = True
''               txtSalesArea1.Enabled = True
''               txtSales.Enabled = True
'            '各區主管
'            Case "SM"
'               txtSalesArea.Locked = True
'               txtSalesArea1.Locked = True
'               '原羅文旭72009可兼看中一區,94/7/1只可看S22
'               '2005/7/5林永生71003可看中所全部,但預設S23
'               If strUserNum = "71003" Then
'                  txtSalesArea = "S23"
'                  txtSalesArea1 = "S23"
'                  txtSalesArea.Locked = False
'                  txtSalesArea1.Locked = False
'                  txtSalesArea.Enabled = True
'                  txtSalesArea1.Enabled = True
'               Else
'               '簡協理可看北所全部但預設S15
'               If strUserNum = "69005" Then
'                  txtSalesArea = "S15"
'                  txtSalesArea1 = "S15"
'                  txtSalesArea.Locked = False
'                  txtSalesArea1.Locked = False
'                  txtSalesArea.Enabled = True
'                  txtSalesArea1.Enabled = True
'               Else
'                  txtSalesArea = stST15
'                  txtSalesArea1 = stST15
'               End If
'               End If
'               txtSales.Enabled = True
'            '加入外商主管  王宗珮、洪琬姿、葉易雲
'            Case "21", "26", "28"
'               txtSalesArea = stST15
'               txtSalesArea1 = stST15
'               txtSales = strUserNum
'               txtSales.Enabled = True
'            '其他只能看自己
'            Case Else
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
'   '若操作人員的ST05=SA,此類人員稱之為帶人主管,開放智權人員欄位可輸入
'   If PUB_GetST05Limits(strUserNum) = True Then
'      txtSales.Enabled = True
'   End If
'
'   'Add by Amy 2015/02/03 +總經理業務工作代理人員
'   If CheckLevel(strUserNum, "總經理業務工作代理人員") = True Then
'       bolSpecMan = True
'       strSpecCode = "總經理業務工作代理人員"
'   'Modify  by Amy 2014/05/20 開放專利處部份智權同仁資料給彥葶代為處理
'   ElseIf CheckLevel(strUserNum, "A8") = True Then
'        bolSpecMan = True
'        strSpecCode = "A8"
'   End If
'
'   If bolSpecMan = True Then
'        'Add by Amy 2015/02/03 +總經理業務工作代理人員
'        If InStr(strSpecCode, "總經理業務工作代理人員") > 0 Then
'            txtSalesArea.Enabled = True: txtSalesArea = ""
'            txtSalesArea1.Enabled = True: txtSalesArea1 = ""
'            txtSales.Enabled = True
'        End If
'        If InStr(strSpecCode, "A8") > 0 Then txtSales.Enabled = True: txtSales = ""
'   Else
'        txtSales = strUserNum
'   End If
'   'end 2014/05/20
   
   'Add By Sindy 2014/6/30
   If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Or Pub_StrUserSt03 = "M51" Then
      LblCntTime.Visible = True
   Else
      LblCntTime.Visible = False
   End If
   '2014/6/30 END
   
   SetDataListWidth
   
'   'Add By Sindy 2020/7/1 記錄原操作人可以查詢的業務區及所別
'   'txtZone.Tag = txtZone
'   txtSalesArea.Tag = txtSalesArea
'   txtSalesArea1.Tag = txtSalesArea1
'   '2020/7/1 END
   
   PUB_AddExcuteLog Me.Name 'Added by Lydia 2021/01/11 登入記錄
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Modify By Sindy 2012/5/15
   'pub_CallNextForm = False
   If pub_CallNextForm = True Then
      pub_CallNextForm = True
      frm210136.Show
      frm210136.cmdSearch_Click
   End If
   Set frm210134 = Nothing
End Sub

'取得正確的 row & col
Public Sub getGrdColRow(ByRef oObj As MSHFlexGrid, ByVal x As Single, ByVal y As Single, ByRef col As Long, ByRef row As Long)
Dim nIndex As Integer
col = 0: row = 0
For nIndex = 0 To oObj.Rows - 1
    If y > oObj.RowHeight(nIndex) Then
        row = row + 1
        y = y - oObj.RowHeight(nIndex)
    ElseIf y > 0 Then
        row = row + 1
        Exit For
    End If
Next nIndex
For nIndex = 0 To oObj.Cols - 1
    If x > oObj.ColWidth(nIndex) Then
        col = col + 1
        x = x - oObj.ColWidth(nIndex)
    ElseIf x > 0 Then
        col = col + 1
        Exit For
    End If
Next nIndex
col = col - 1 + IIf(oObj.LeftCol <> oObj.FixedCols And oObj.LeftCol <> 0, oObj.LeftCol - oObj.FixedCols, 0)
row = row - 1 + IIf(oObj.TopRow <> oObj.FixedRows And oObj.TopRow <> 0, oObj.TopRow - oObj.FixedRows, 0)

If col > oObj.Cols - 1 Then col = oObj.Cols - 1
If row > oObj.Rows - 1 Then row = oObj.Rows - 1
End Sub

Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow grdDataList, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   grdDataList.col = nCol
   grdDataList.row = nRow
   If Me.grdDataList.row < 1 Then
      'Modify By Sindy 2024/12/31 + 最近歷程
      If Me.grdDataList.Text = "V" Or Me.grdDataList.Text = "收文日" Or _
         Me.grdDataList.Text = "本所案號" Or Me.grdDataList.Text = "案件性質" Or _
         Me.grdDataList.Text = "申請國家" Or _
         Me.grdDataList.Text = "智權人員" Or Me.grdDataList.Text = "承辦人" Or _
         Me.grdDataList.Text = "管制期限" Or Me.grdDataList.Text = "本所期限" Or _
         Me.grdDataList.Text = "法定期限" Or Me.grdDataList.Text = "預定會稿日" Or _
         Me.grdDataList.Text = "最近歷程" Or Me.grdDataList.Text = "備註" Or Me.grdDataList.Text = "總收文號" Then
         If m_blnColOrderAsc = True Then
            strOrderBy = Me.grdDataList.Text & " asc,總收文號 asc" '昇冪
            m_blnColOrderAsc = False
         Else
            strOrderBy = Me.grdDataList.Text & " desc,總收文號 desc" '降冪
            m_blnColOrderAsc = True
         End If
         bolShowMsgBox = False
         Call doQuery
      End If
'      Select Case Me.grdDataList.col
'         Case 0
'            If m_blnColOrderAsc = True Then
'               Me.grdDataList.Sort = 3 '數值昇冪
'               m_blnColOrderAsc = False
'            Else
'               Me.grdDataList.Sort = 4 '數值降冪
'               m_blnColOrderAsc = True
'            End If
'         Case Else
'            If m_blnColOrderAsc = True Then
'               Me.grdDataList.Sort = 5 '字串昇冪
'               m_blnColOrderAsc = False
'            Else
'               Me.grdDataList.Sort = 6 '字串降冪
'               m_blnColOrderAsc = True
'            End If
'      End Select
   End If
End Sub

Private Sub grdDataList_SelChange()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

grdDataList.row = grdDataList.MouseRow
grdDataList.col = 0
If grdDataList.row <> 0 Then
   If grdDataList.Text = "V" Then
      grdDataList.Text = ""
      '逾管制期限變黃色
      If grdDataList.TextMatrix(grdDataList.MouseRow, 10) <> "" And _
         ChangeTStringToWString(ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.MouseRow, 10))) < strSrvDate(1) Then
         For j = 0 To grdDataList.Cols - 1
            grdDataList.col = j
            grdDataList.CellBackColor = &HFFFF& '黃色
         Next j
      '逾預定會稿日變淺紅
      ElseIf grdDataList.TextMatrix(grdDataList.MouseRow, 14) <> "" And _
         ChangeTStringToWString(ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.MouseRow, 14))) < strSrvDate(1) Then
         '2010/9/28 add by sonia 有會稿日即使逾預定會稿日也不必變淺紅
         StrSQLa = "select ep07 from engineerprogress where ep02='" & grdDataList.TextMatrix(grdDataList.MouseRow, 17) & "' "
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            If "" & rsA.Fields(0) = "" Then
         '2010/9/28 end
               For j = 0 To grdDataList.Cols - 1
                  grdDataList.col = j
                  grdDataList.CellBackColor = &H8080FF '淺紅
               Next j
            End If
         End If
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
      Else
         For i = 0 To grdDataList.Cols - 1
            grdDataList.col = i
            If i <= 6 Then
               grdDataList.CellBackColor = QBColor(7)
            Else
               grdDataList.CellBackColor = QBColor(15)
            End If
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
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
         txtControlDate(0).SetFocus
      Case 2
         txtRecvDate(0).SetFocus
   End Select
End Sub

Private Sub systemkind_GotFocus()
   TextInverse systemkind
   CloseIme
   'Modify By Sindy 2017/10/25 Mark
'   'ADD BY SONIA 2015/11/20
'   Option1(1).Value = True
'   Option1(0).Value = False
'   Option1(2).Value = False
'   'END 2015/11/20
End Sub

Private Sub systemkind_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtControlDate_Click(Index As Integer)
   Option1(0).Value = True
End Sub

Private Sub txtControlDate_GotFocus(Index As Integer)
   TextInverse txtControlDate(Index)
   CloseIme
End Sub

Private Sub txtControlDate_LostFocus(Index As Integer)
   If txtControlDate(Index).Tag <> txtControlDate(Index).Text Then
      Combo2.ListIndex = 4 '3 設n個月 為空白
   End If
End Sub

Private Sub txtControlDate_Validate(Index As Integer, Cancel As Boolean)
   If txtControlDate(Index) <> "" Then
      If ChkDate(txtControlDate(Index)) = False Then
         Cancel = True
         txtControlDate(Index).SetFocus
         txtControlDate_GotFocus Index
         Exit Sub
      End If
      If Index = 1 Then
         If RunNick2(txtControlDate(0), txtControlDate(1)) = True Then
            txtControlDate(Index).SetFocus
            txtControlDate_GotFocus Index
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

'CANCEL BY SONIA 2014/6/26
'Private Sub txtCU1_LostFocus()
'   txtCU1 = Mid(txtCU1 & "000000000", 1, 9)
'End Sub
'END 2014/6/26

Private Sub txtRecvDate_Click(Index As Integer)
   Option1(2).Value = True
End Sub

Private Sub txtRecvDate_GotFocus(Index As Integer)
   TextInverse txtRecvDate(Index)
   CloseIme
End Sub

Private Sub txtRecvDate_LostFocus(Index As Integer)
   If txtRecvDate(Index).Tag <> txtRecvDate(Index).Text Then
      Combo1.ListIndex = 4 '3 設n個月 為空白
   End If
End Sub

Private Sub txtRecvDate_Validate(Index As Integer, Cancel As Boolean)
   If txtRecvDate(Index) <> "" Then
      If ChkDate(txtRecvDate(Index)) = False Then
         Cancel = True
         txtRecvDate(Index).SetFocus
         txtRecvDate_GotFocus Index
         Exit Sub
      End If
      If Index = 1 Then
         If RunNick2(txtRecvDate(0), txtRecvDate(1)) = True Then
            txtRecvDate(Index).SetFocus
            txtRecvDate_GotFocus Index
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub txtCP10_GotFocus(Index As Integer)
   TextInverse txtCP10(Index)
   CloseIme
   'Modify By Sindy 2017/10/25 Mark
'   'ADD BY SONIA 2015/11/20
'   Option1(1).Value = True
'   Option1(0).Value = False
'   Option1(2).Value = False
'   'END 2015/11/20
End Sub

Private Sub txtCU1_GotFocus()
   TextInverse txtCU1
   CloseIme
   'Modify By Sindy 2017/10/25 Mark
'   'ADD BY SONIA 2015/11/20
'   Option1(1).Value = True
'   Option1(0).Value = False
'   Option1(2).Value = False
'   'END 2015/11/20
End Sub

Private Sub txtCU1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCU2_GotFocus()
   'MODIFY BY SONIA 2014/6/26
   'If Len(txtCU1) = 9 Then
   If txtCU1 <> "" Then
   'END 2014/6/26
      txtCU2 = Left(txtCU1, 6) & "ZZZ"
      txtCU2.SelStart = 6
      txtCU2.SelLength = 3
   End If
   TextInverse txtCU2
   CloseIme
   'Modify By Sindy 2017/10/25 Mark
'   'ADD BY SONIA 2015/11/20
'   Option1(1).Value = True
'   Option1(0).Value = False
'   Option1(2).Value = False
'   'END 2015/11/20
End Sub

Private Sub txtCU2_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSales_Change()
   If Len(txtSales) > 4 Then
      lblSalesName = GetStaffName(txtSales, True)
   Else
      lblSalesName = ""
   End If
End Sub

Private Sub txtSales_GotFocus()
   TextInverse txtSales
   CloseIme
   If Combo3.Enabled = True And Combo3.Visible = True Then Combo3.SetFocus 'Add By Sindy 2022/2/21
End Sub

'Add By Sindy 2010/11/26
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

Private Sub txtSalesArea1_Validate(Cancel As Boolean)
If Trim(txtSalesArea1) <> "" Then
   If RunNick(txtSalesArea, txtSalesArea1) = True Then
      Cancel = True
      Exit Sub
   End If
End If
End Sub

Private Sub cmdok_Click(Index As Integer)
cmdState = Index
PubShowNextData
End Sub

Public Sub PubShowNextData()
Dim i As Integer, j As Integer
Dim intRow As Integer
Dim strCP01 As String 'Add By Sindy 2012/5/21

Select Case cmdState
Case 0 '歷史記錄
     Me.Enabled = False
     For i = 1 To grdDataList.Rows - 1
     grdDataList.col = 0
     grdDataList.row = i
     If Trim(grdDataList.Text) = "V" Then
         grdDataList.col = 0
         grdDataList.Text = ""
         For j = 0 To grdDataList.Cols - 1
            grdDataList.col = j
            If j <= 6 Then
               grdDataList.CellBackColor = QBColor(7)
            Else
               grdDataList.CellBackColor = QBColor(15)
            End If
         Next j
         grdDataList.col = 17 '總收文號
         If Not IsNull(grdDataList.Text) Then
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            frm210134_3.Show
            frm210134_3.Process Pub_RplStr(grdDataList.Text)
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
         End If
     End If
     Next i
     Call SetColor
     Me.Enabled = True
Case 1 '承辦進度
     Me.Enabled = False
     For i = 1 To grdDataList.Rows - 1
     grdDataList.col = 0
     grdDataList.row = i
     If Trim(grdDataList.Text) = "V" Then
         grdDataList.col = 0
         grdDataList.Text = ""
         For j = 0 To grdDataList.Cols - 1
            grdDataList.col = j
            If j <= 6 Then
               grdDataList.CellBackColor = QBColor(7)
            Else
               grdDataList.CellBackColor = QBColor(15)
            End If
         Next j
         grdDataList.col = 17 '總收文號
         If Not IsNull(grdDataList.Text) Then
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            'Modify By Sindy 2012/5/21 +if,frm100101_K
            strCP01 = GetCaseProData(Trim(Pub_RplStr(grdDataList.Text)), "CP01")
            If strCP01 = "P" Or strCP01 = "PS" Or strCP01 = "FG" Or _
               strCP01 = "FCP" Or strCP01 = "CFP" Or strCP01 = "CPS" Or _
               Val(strSrvDate(1)) < Val(TMdebateStarDT) Then  '專利處工作進度
               frm100101_F.CmdFormName = UCase(Me.Name) 'Add By Sindy 2010/10/29
               frm100101_F.Show
               frm100101_F.Process Pub_RplStr(grdDataList.Text)
            Else
               frm100101_K.CmdFormName = UCase(Me.Name)
               frm100101_K.Show
               frm100101_K.Process Pub_RplStr(grdDataList.Text)
            End If
            '2012/5/21 End
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
         End If
     End If
     Next i
     Call SetColor
     Me.Enabled = True
Case 2 '繪圖進度
     Me.Enabled = False
     For i = 1 To grdDataList.Rows - 1
     grdDataList.col = 0
     grdDataList.row = i
     If Trim(grdDataList.Text) = "V" Then
         grdDataList.col = 0
         grdDataList.Text = ""
         For j = 0 To grdDataList.Cols - 1
            grdDataList.col = j
            If j <= 6 Then
               grdDataList.CellBackColor = QBColor(7)
            Else
               grdDataList.CellBackColor = QBColor(15)
            End If
         Next j
         grdDataList.col = 17 '總收文號
         If Not IsNull(grdDataList.Text) Then
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            frm100101_g.Show
            frm100101_g.Process Pub_RplStr(grdDataList.Text)
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
         End If
     End If
     Next i
     Call SetColor
     Me.Enabled = True
Case 3 '發E-Mail
      Me.Enabled = False
      For i = 1 To grdDataList.Rows - 1
         grdDataList.col = 0
         grdDataList.row = i
         If Trim(grdDataList.Text) = "V" Then
            grdDataList.col = 0
            grdDataList.Text = ""
            For j = 0 To grdDataList.Cols - 1
               grdDataList.col = j
               If j <= 6 Then
                  grdDataList.CellBackColor = QBColor(7)
               Else
                  grdDataList.CellBackColor = QBColor(15)
               End If
            Next j
            If Not IsNull(grdDataList.TextMatrix(i, 17)) Then
               If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
               End If
               Screen.MousePointer = vbHourglass
               frm210134_2.m_strSC01 = Pub_RplStr(grdDataList.TextMatrix(i, 17)) '總收文號
               frm210134_2.m_strSC02 = Pub_RplStr(grdDataList.TextMatrix(i, 19))
               frm210134_2.m_strSC03 = Pub_RplStr(grdDataList.TextMatrix(i, 20))
               frm210134_2.Show
               frm210134_2.Process
               Screen.MousePointer = vbDefault
               Me.Enabled = True
               Exit Sub
            End If
         End If
      Next i
      Call SetColor
      Me.Enabled = True
Case Else
End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Me.Text1(Index)
   CloseIme
   'Modify By Sindy 2017/10/25 Mark
'   'ADD BY SONIA 2015/11/20
'   Option1(1).Value = True
'   Option1(0).Value = False
'   Option1(2).Value = False
'   'END 2015/11/20
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Added by Lydia 2015/06/18 +客戶中文名稱查詢
Private Sub txtCuName_GotFocus()
   'Added by Lydia 2015/06/25 輸入客戶名稱按Enter自動執行搜尋後,執行查詢
   cmdSearch.Default = False: cmdFind.Default = True
   
   TextInverse Me.txtCuName
   OpenIme
   'Modify By Sindy 2017/10/25 Mark
'   'ADD BY SONIA 2015/11/20
'   Option1(1).Value = True
'   Option1(0).Value = False
'   Option1(2).Value = False
'   'END 2015/11/20
End Sub

'Added by Lydia 2015/06/25 輸入客戶名稱按Enter自動執行搜尋後,執行查詢
Private Sub txtCuName_LostFocus()
   cmdSearch.Default = True: cmdFind.Default = False
End Sub

'Added by Lydia 2015/06/18 +客戶中文名稱查詢
Private Sub cmdFind_Click()
   If Me.txtCuName.Text = "" Then
      MsgBox "請輸入客戶中文名稱的關鍵字!!!", vbExclamation + vbOKOnly
      Me.txtCuName.SetFocus
      txtCuName_GotFocus
      Exit Sub
   End If
   frm090801_1.m_strCustChnName = Me.txtCuName.Text
   frm090801_1.lblName.Caption = Me.txtCuName.Text
   m_blnOneRec = False
   m_strCustCode = ""
   If frm090801_1.StrMenu = True Then
      If frm090801_1.m_blnOneRec = False Then
         frm090801_1.Show vbModal
      End If
      m_blnOneRec = frm090801_1.m_blnOneRec
      m_strCustCode = frm090801_1.m_strCustCode
      Unload frm090801_1
   Else
      Unload frm090801_1
   End If
   If m_blnOneRec = True And m_strCustCode <> "" Then
      Me.txtCU1.Text = m_strCustCode
      Me.txtCU2.Text = IIf(Right(m_strCustCode, 3) = "000", Mid(m_strCustCode, 1, 6) & "ZZZ", IIf(Right(m_strCustCode, 1) = "0", Mid(m_strCustCode, 1, 8) & "Z", m_strCustCode))
      Me.txtCuName.Text = GetCustomerName(m_strCustCode)
   End If
   'Added by Lydia 2015/06/25 輸入客戶名稱按Enter自動執行搜尋後,執行查詢
   If Me.txtCU1.Text <> "" And Me.txtCU2.Text <> "" Then
      Call cmdSearch_Click
   End If
End Sub

'Add By Sindy 2017/10/25 管制期限
Private Sub CountMonthToDay2()
   If Combo2.Text <> "" Then
      If Val(Combo2.Text) > 0 Then
         txtControlDate(0) = ChangeWDateStringToTString(ChangeWStringToWDateString(strSrvDate(1)))
         txtControlDate(1).Text = Val(Format(DateAdd("M", Val(Combo2.Text), ChangeTStringToWDateString(txtControlDate(0))), "YYYYMMDD")) - 19110000
      ElseIf Val(Combo2.Text) < 0 Then
         txtControlDate(1) = ChangeWDateStringToTString(ChangeWStringToWDateString(strSrvDate(1)))
         txtControlDate(0).Text = Val(Format(DateAdd("M", Val(Combo2.Text), ChangeTStringToWDateString(txtControlDate(1))), "YYYYMMDD")) - 19110000
      End If
      txtControlDate(0).Tag = txtControlDate(0).Text
      txtControlDate(1).Tag = txtControlDate(1).Text
   End If
End Sub
Private Sub Combo2_Change()
   Call CountMonthToDay2
End Sub
Private Sub Combo2_Click()
   Call CountMonthToDay2
End Sub
'Add By Sindy 2017/10/25 收文日期
Private Sub CountMonthToDay()
   If Combo1.Text <> "" Then
      If Val(Combo1.Text) > 0 Then
         txtRecvDate(0) = ChangeWDateStringToTString(ChangeWStringToWDateString(strSrvDate(1)))
         txtRecvDate(1).Text = Val(Format(DateAdd("M", Val(Combo1.Text), ChangeTStringToWDateString(txtRecvDate(0))), "YYYYMMDD")) - 19110000
      ElseIf Val(Combo1.Text) < 0 Then
         txtRecvDate(1) = ChangeWDateStringToTString(ChangeWStringToWDateString(strSrvDate(1)))
         txtRecvDate(0).Text = Val(Format(DateAdd("M", Val(Combo1.Text), ChangeTStringToWDateString(txtRecvDate(1))), "YYYYMMDD")) - 19110000
      End If
      txtRecvDate(0).Tag = txtRecvDate(0).Text
      txtRecvDate(1).Tag = txtRecvDate(1).Text
   End If
End Sub
Private Sub Combo1_Change()
   Call CountMonthToDay
End Sub
Private Sub Combo1_Click()
   Call CountMonthToDay
End Sub
'2017/10/25 END
