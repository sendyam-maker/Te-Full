VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm210107 
   BorderStyle     =   1  '單線固定
   Caption         =   "業績達成日報表"
   ClientHeight    =   5712
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9432
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5712
   ScaleWidth      =   9432
   Begin VB.ListBox List1 
      Height          =   948
      Left            =   36
      TabIndex        =   14
      Top             =   3564
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "產生Excel"
      Height          =   400
      Left            =   7440
      TabIndex        =   5
      Top             =   90
      Width           =   1000
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   285
      Left            =   1260
      MaxLength       =   7
      TabIndex        =   4
      Top             =   36
      Width           =   1140
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6570
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8520
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   90
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4068
      Left            =   96
      TabIndex        =   2
      Top             =   660
      Width           =   9216
      _ExtentX        =   16235
      _ExtentY        =   7176
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   3
      FixedCols       =   2
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   7.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   8.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "label6"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   228
      Left            =   180
      TabIndex        =   15
      Top             =   5040
      Width           =   528
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   2520
      TabIndex        =   13
      Top             =   72
      Width           =   468
   End
   Begin VB.Label lblPS 
      AutoSize        =   -1  'True
      Caption         =   "lblPS"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   180
      TabIndex        =   12
      Top             =   4788
      Width           =   444
   End
   Begin VB.Label lblNet 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   7548
      TabIndex        =   11
      Top             =   5016
      Width           =   1152
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "本月累計之收款-應收："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   5040
      TabIndex        =   10
      Top             =   5016
      Width           =   2496
   End
   Begin VB.Label lblWorkDay 
      Caption         =   "lblWorkDay"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4500
      TabIndex        =   9
      Top             =   350
      Width           =   276
   End
   Begin VB.Label lblWorkDays 
      Caption         =   "lblWorkDays"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1260
      TabIndex        =   8
      Top             =   350
      Width           =   276
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "今天為第　　　工作天"
      Height          =   180
      Left            =   3600
      TabIndex        =   7
      Top             =   380
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本月工作天　　　天"
      Height          =   180
      Left            =   228
      TabIndex        =   6
      Top             =   380
      Width           =   1620
   End
   Begin VB.Label Label2 
      Caption         =   "統計日期"
      Height          =   180
      Left            =   228
      TabIndex        =   3
      Top             =   72
      Width           =   900
   End
End
Attribute VB_Name = "frm210107"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/25 改成Form2.0 ; grdDataList改字型=新細明體-ExtB
'Memo by Lydia 2021/07/27 更名為「業績達成日報表」 'Memo by Lydia 2021/08/27 上線
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo by Morgan2010/8/19 日期欄已修改
'Add by Morgan 2005/5/3
Option Explicit

Dim stWDate As String, stTDate As String, stWYM As String, stTYM As String
Dim iRow As Integer, iCol As Integer, iXCol As Integer
Dim iDept01 As Integer, iDept02 As Integer, iDept03 As Integer, iDeptS29 As Integer 'Added by Lydia 2017/10/12 例外欄位的行位置
'Add by Amy 2018/11/28
Dim iDept04 As Integer '國外部拆成FCP/FCT
Dim stDate0 As String, stDDate As String
'Add by Amy 2021/02/17
Dim iDept05 As Integer, iDept06 As Integer 'F4102/F4103拆成F4104~07
Dim Is17Col As Boolean '顯示17行資料(不含1及2行欄位名)
'Add by Amy 2021/02/23
Dim iDept07 As Integer '顧服組
Dim i As Integer, intCounter As Integer, intField As Integer
Dim strCol As String
Dim arrField() As String, arrWidth, arrField2() As String, arrWidth2
Dim iDept08 As Integer 'Add by Amy 2021/03/12 MCT
Dim iDeptSub As Integer 'Added by Lydia 2021/09/03 小計=智權部+客服組
Dim iDept09 As Integer 'Add by Amy 2022/04/11 MCP
'Modify by 2024/04/02 改變數(原Add by Amy 2023/06/01 strF4104/5/strF4104Acs/strF4105Acs)
Dim strF410x(4 To 5) As String, strF410xAcs(4 To 5) As String '專利國外部,日本部/專利國外部ACS累計,日本部ACS累計
Dim bolErr As Boolean '橫向加總<>全所
Dim strSPData As String, stColSQL As String, strShowDept(1) As String, stSumMonGT As String 'SalesPoint有資料/抓欄位語法/欄位需show部門/本月累計收文-全所

'Modify by Morgan 2006/2/9 改Grid字體大小及欄位寬要能顯示全部資料--杜副總
'p_bolHeaderOnly:是否只設定表頭 true=是 false=資料一併清除
Private Sub SetDataListWidth_old(Optional ByVal p_bolHeaderOnly As Boolean = False)
   Dim ii As Integer, iColWith(1 To 3) As Integer
   Dim iCol As Integer
   'Modify by Morgan 2006/2/9
'   iColWith(1) = 300
'   iColWith(2) = 990
'   iColWith(3) = 915
   iColWith(1) = 280
   iColWith(2) = 1200 '850
   iColWith(3) = 665
   'Modify by Amy 2018/11/28 國外拆成FCP/FCT
   'Modify by Amy 2021/02/17 F4102/F4103 拆成F4104~07
   'iXCol = IIf(Val(txtCloseDate) > Val(每日業務點數FCPFCT啟用日), 18, 17) 'Added by Lydia 2017/10/17
   iXCol = 18

   'Modify by Amy 2022/04/11 +11005月啟顯示MCP
   If Val(txtCloseDate) >= 1110401 Then
        iXCol = 23
   'Modify by Amy 2021/02/23 +if 11004月起顯示MCT,每日業務點數FCPFCT啟用日後都顯示顧服組
   ElseIf Val(txtCloseDate) >= 1100401 Then
        iXCol = 22
   ElseIf Val(txtCloseDate) >= 1100101 Then
        iXCol = 21
   ElseIf Val(txtCloseDate) > Val(每日業務點數FCPFCT啟用日) Then
        iXCol = 19
   End If
   'end 2021/02/23
   'end 2021/02/17
   'Added by Lydia 2021/08/10 日期條件>1100801時，中三區不要出現
   If Val(txtCloseDate) >= 1100801 Then
        'Modify by Amy 2022/04/11
        'iXCol = 21
        iXCol = iXCol - 1
        'Add by Amy 2022/07/12 11107月起不再顯示客服組W10
        If Val(txtCloseDate) >= 1110701 Then
            iXCol = iXCol - 1
        End If
   End If
   'end 2021/08/10
   
   'Added by Lydia 2021/09/03 增加小計
   iDeptSub = 1 '預設1=增加小計
   If iDeptSub > 0 Then iXCol = iXCol + 1
   'end 2021/09/03
   
   With grdDataList
      .Visible = False
      .WordWrap = True
      If p_bolHeaderOnly = False Then
         .Clear
         'Modify by Morgan 2007/10/18 加巨京
         '.Rows = 11: .Cols = 15: .FixedRows = 1: .FixedCols = 2
         'Modified by Lydia 2017/10/12 加中二S22
         '.Rows = 11: .Cols = 16: .FixedRows = 1: .FixedCols = 2
         'Modify by Amy 2022/06/27 原:.Rows = 11
         .Rows = 13: .Cols = iXCol: .FixedRows = 1: .FixedCols = 2
         'end 2007/10/18
         .MergeCells = flexMergeRestrictColumns
         .MergeCol(0) = True '目標 達成
      End If
      iCol = 0
      .row = 0
      .col = iCol: .ColWidth(.col) = iColWith(1): .Text = ""
      .CellAlignment = flexAlignRightCenter: .CellFontBold = True
      .ColAlignmentFixed(.col) = flexAlignCenterCenter
      
      iCol = iCol + 1
      .col = iCol: .ColWidth(.col) = iColWith(2): .Text = "項目"
      .CellAlignment = flexAlignLeftCenter: .CellFontBold = True
      .ColAlignmentFixed(.col) = flexAlignCenterCenter
      
      iCol = iCol + 1
      .col = iCol: .ColWidth(.col) = iColWith(3): .Text = "北一區"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignmentFixed(.col) = flexAlignCenterCenter
      .ColAlignment(.col) = flexAlignRightCenter
      
      iCol = iCol + 1
      .col = iCol: .ColWidth(.col) = iColWith(3): .Text = "北三區"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignmentFixed(.col) = flexAlignCenterCenter
      .ColAlignment(.col) = flexAlignRightCenter
      
      iCol = iCol + 1
      .col = iCol: .ColWidth(.col) = iColWith(3): .Text = "北四區"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignmentFixed(.col) = flexAlignCenterCenter
      .ColAlignment(.col) = flexAlignRightCenter
      
      iCol = iCol + 1
      .col = iCol: .ColWidth(.col) = iColWith(3): .Text = "北五區"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignmentFixed(.col) = flexAlignCenterCenter
      .ColAlignment(.col) = flexAlignRightCenter
      
      iCol = iCol + 1
      .col = iCol: .ColWidth(.col) = iColWith(3): .Text = "中一區"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignmentFixed(.col) = flexAlignCenterCenter
      .ColAlignment(.col) = flexAlignRightCenter
      
'MODIFY BY SONIA 2016/1/13 取消S22加入S24
'Modified by Lydia 2017/10/12 (10/1起) S22成立 (取消Mark)
      iCol = iCol + 1
      .col = iCol: .ColWidth(.col) = iColWith(3): .Text = "中二區"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignmentFixed(.col) = flexAlignCenterCenter
      .ColAlignment(.col) = flexAlignRightCenter
'end 2017/10/12
      
      If Val(txtCloseDate) < 1100801 Then 'Added by Lydia 2021/08/10 日期條件>1100801時，中三區不要出現
        iCol = iCol + 1
        .col = iCol: .ColWidth(.col) = iColWith(3): .Text = "中三區"
        .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
        .ColAlignmentFixed(.col) = flexAlignCenterCenter
        .ColAlignment(.col) = flexAlignRightCenter
      End If 'Added by Lydia 2021/08/10
      
      'ADD BY SONIA 2016/1/13 取消S22加入S24
      iCol = iCol + 1
      .col = iCol: .ColWidth(.col) = iColWith(3): .Text = "中四區"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignmentFixed(.col) = flexAlignCenterCenter
      .ColAlignment(.col) = flexAlignRightCenter
      
      iCol = iCol + 1
      iDeptS29 = iCol 'Added by Lydia 2017/10/12
      .col = iCol: .ColWidth(.col) = iColWith(3) + 150: .Text = "中區其他"
      'Mark by Amy 2021/02/23 中區其他也列示
      'If Val(txtCloseDate) >= 1040401 Then .ColWidth(.col) = 0   'ADD BY SONIA 2015/4/9 104/4起取消中區其他的目標
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignmentFixed(.col) = flexAlignCenterCenter
      .ColAlignment(.col) = flexAlignRightCenter
      
      iCol = iCol + 1
      .col = iCol: .ColWidth(.col) = iColWith(3): .Text = "台南所"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignmentFixed(.col) = flexAlignCenterCenter
      .ColAlignment(.col) = flexAlignRightCenter
      
      iCol = iCol + 1
      .col = iCol: .ColWidth(.col) = iColWith(3): .Text = "高雄所"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignmentFixed(.col) = flexAlignCenterCenter
      .ColAlignment(.col) = flexAlignRightCenter
      
      'Modify by Amy 2022/07/12 11107月起不再顯示客服組W10
      If Val(txtCloseDate) < 1110701 Then
        'Modify by Amy 2021/02/23 原:其他欄於客服組前,改客服組 +顧服組/MCT 再顯示其他
        'Add by Morgan 2007/10/18
        iCol = iCol + 1
        iDept02 = iCol 'Added by Lydia 2017/10/12
        .col = iCol: .ColWidth(.col) = iColWith(3)
        'Modify by Amy 2019/07/12 原:巨京
        .Text = "客服組"
        .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
        .ColAlignmentFixed(.col) = flexAlignCenterCenter
        .ColAlignment(.col) = flexAlignRightCenter
        'end 2007/10/18
      End If
      
      'Added by Lydia 2021/09/03 增加小計=智權部+客服組
      If iDeptSub > 0 Then
         iCol = iCol + 1
         iDeptSub = iCol
         'Modify by Amy 2021/12/23 原:小計
         .col = iCol: .ColWidth(.col) = 760: .Text = "智權部小計"
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .ColAlignmentFixed(.col) = flexAlignCenterCenter
         .ColAlignment(.col) = flexAlignRightCenter
      End If
      'end 2021/09/03
      
      iCol = iCol + 1
      iDept07 = iCol
      .col = iCol: .ColWidth(.col) = iColWith(3)
      .Text = "顧服組"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignmentFixed(.col) = flexAlignCenterCenter
      .ColAlignment(.col) = flexAlignRightCenter
      
      iCol = iCol + 1
      iDept01 = iCol 'Added by Lydia 2017/10/12
      .col = iCol: .ColWidth(.col) = iColWith(3): .Text = "其他"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignmentFixed(.col) = flexAlignCenterCenter
      .ColAlignment(.col) = flexAlignRightCenter
      
      'Add by Amy 2022/04/11 MCP
      If Val(txtCloseDate) >= 1110401 Then
        iCol = iCol + 1
        iDept09 = iCol
        .col = iCol: .ColWidth(.col) = iColWith(3): .Text = Replace(StaffQuery("P1005"), "專利", " 專利 ")
        .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
        .ColAlignmentFixed(.col) = flexAlignCenterCenter
        .ColAlignment(.col) = flexAlignRightCenter
      End If
      'end 2022/04/11
      
      'MCT
      If Val(txtCloseDate) >= 1100401 Then
        iCol = iCol + 1
        iDept08 = iCol 'Modify by Amy 2021/03/12
        .col = iCol: .ColWidth(.col) = iColWith(3): .Text = StaffQuery("P2005")
        .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
        .ColAlignmentFixed(.col) = flexAlignCenterCenter
        .ColAlignment(.col) = flexAlignRightCenter
      End If
      'end 2021/02/23
     
      
      'Modify by Morgan 2006/2/9
      '.col = 12: .ColWidth(.col) = 1200: .Text = "國外部"
      iCol = iCol + 1
      iDept03 = iCol 'Added by Lydia 2017/10/12
      .col = iCol: .ColWidth(.col) = 760
      'Modify by Amy 2021/02/17 F4102/F4103 拆成F4104~07
      If Val(txtCloseDate) >= 1100101 Then
        .Text = StaffQuery("F4104")
      'Modify by Amy 2018/11/28 1080101國外部改區分FCP及FCT
      ElseIf Val(txtCloseDate) > Val(每日業務點數FCPFCT啟用日) Then
        .Text = "FCP"
      Else
        .Text = "國外部"
      End If
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignmentFixed(.col) = flexAlignCenterCenter
      .ColAlignment(.col) = flexAlignRightCenter
      
      'Add by Amy 2018/11/28 1080101國外部改區分FCP及FCT
      If Val(txtCloseDate) > Val(每日業務點數FCPFCT啟用日) Then
        iCol = iCol + 1
        iDept04 = iCol
        .col = iCol: .ColWidth(.col) = 760
        If Val(txtCloseDate) >= 1100101 Then
            .Text = StaffQuery("F4105")
        Else
            .Text = "FCT"
        End If
        .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
        .ColAlignmentFixed(.col) = flexAlignCenterCenter
        .ColAlignment(.col) = flexAlignRightCenter
      End If
      If Val(txtCloseDate) >= 1100101 Then
        iCol = iCol + 1
        iDept05 = iCol
        .col = iCol: .ColWidth(.col) = 760
        .Text = StaffQuery("F4106")
        .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
        .ColAlignmentFixed(.col) = flexAlignCenterCenter
        .ColAlignment(.col) = flexAlignRightCenter
        
        iCol = iCol + 1
        iDept06 = iCol
        .col = iCol: .ColWidth(.col) = 760
        .Text = StaffQuery("F4107")
        .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
        .ColAlignmentFixed(.col) = flexAlignCenterCenter
        .ColAlignment(.col) = flexAlignRightCenter
      End If
      'end 2021/02/17
      'Add by Amy 2022/04/11
      If Val(txtCloseDate) >= 1100101 Then
        .RowHeight(0) = 755
      End If
      
      'Modify by Morgan 2006/2/9
      '.col = 13: .ColWidth(.col) = 1200: .Text = "全所"
      iCol = iCol + 1
      .col = iCol: .ColWidth(.col) = 760: .Text = "全所"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignmentFixed(.col) = flexAlignCenterCenter
      .ColAlignment(.col) = flexAlignRightCenter
      
      .col = 0
      For ii = 1 To 3
         .row = ii: .Text = "目標": .CellFontBold = True
      Next
      'Modify by Amy 2022/06/27 +本月ACS累計/ACS未收款 原:For ii = 4 To 9
      For ii = 4 To 11
         .row = ii: .Text = "達成": .CellFontBold = True
      Next
      
      .row = 12 'Modify by Amy 2022/06/27 +本月ACS累計/ACS未收款 (原:10)
      .col = 1: .Text = "備註"
      .CellAlignment = flexAlignLeftCenter: .CellFontBold = True
      
      .TextMatrix(1, 1) = "本月目標"
      .TextMatrix(2, 1) = "每日應收"
      .TextMatrix(3, 1) = "本月累計應收" '合計應收
      .TextMatrix(4, 1) = "本日入帳" '本日財務
      .TextMatrix(5, 1) = "本月累計入帳" '本月財務
      .TextMatrix(6, 1) = "本日收款" '本日實收
      .TextMatrix(7, 1) = "本月累計收款" '本月實收
      'Modify By Sindy 2020/9/1
'      .TextMatrix(8, 1) = "已簽約"
'      .TextMatrix(9, 1) = "簽收合計"
      .TextMatrix(8, 1) = "本月累計收文"
      'Modify by Amy 2022/06/27 +本月ACS累計/ACS未收款
      '.TextMatrix(9, 1) = "未收款"
      .TextMatrix(9, 1) = "本月ACS累計"
      .TextMatrix(10, 1) = "未收款"
      .TextMatrix(11, 1) = "ACS未收款"
      '2020/9/1 END
      .row = 3: .col = 1: .CellFontBold = True
      .row = 7: .col = 1: .CellFontBold = True
      .Refresh
      .Visible = True
   End With
End Sub

'Mark by Amy 2024/03/01 欄位改寫入暫存檔
Private Function doQuery_Old() As Boolean
'Dim stVTable1 As String, stVTable2 As String, stVTable3 As String, stVTable4 As String, stVTable5 As String
'Dim stWorkDays As String, stWorkDay As String
'Dim stF4100Amt As Variant   'add by sonia 2014/6/11
'Dim stCon As String 'Added by Lydia 2018/08/23
''Add by Amy 2018/11/28
'Dim i As Integer, intRun As Integer
''modify by sonia 2021/1/19
''Dim stF410X(1 To 2) As Variant, stVal As Variant
'Dim stF410X() As Variant, stVal As Variant
'Dim stF410XVal() As Variant, k As Integer
''end 2021/1/19
'Dim stSalesArea As String
'
'   If GetWorkDay(txtCloseDate, stWorkDay, stWorkDays) = True Then
'      lblWorkDays = stWorkDays
'      lblWorkDay = stWorkDay
'   Else
'      Exit Function
'   End If
'   'Memo by Lydia 2015/07/29 程式若有改變,請一併變更frmAutoBatchDay.strMenu64
'   '統計日
'   stDDate = Val(txtCloseDate)
'   '統計月份
'   stDate0 = Val(stDDate) \ 100 & "01"
'
'   stWDate = TransDate(txtCloseDate, 2)
'   stTDate = Format(stWDate - 19110000)
'   stWYM = Left(stWDate, 6)
'   stTYM = Format(stWYM - 191100)
'   'Added by Lydia 2018/08/23 比照每日業績點數輸入(frm210103),排除特定員工代號(ex.中所20011,20021,20031)
'   stCon = " And ST01>'60000' "
'   'Added by Lydia 2018/08/23 判斷離職前一天的月份要顯示(離職日為1號則當月不用)
'   stCon = stCon & " And (st04='1' OR substr( decode(ST51,null,'',to_char(to_date(st51,'yyyymmdd')-1,'yyyymmdd')) ,1,6)>= " & Left(TransDate(stDate0, 2), 6) & ") "
'
'   '已入帳資料
'   'modify by sonia 2014/1/21 取消a0201='1'條件
'   'modify by sonia 2015/4/20 加不含'4194'科目,不含結餘傳票不要加科目限制
'   'stVTable1 = "select ax209 V1C0,sum(decode(a0205," & stDDate & ", ax207)) V1C1" & _
'      ",sum(ax207) V1C2" & _
'      " From acc020, acc021" & _
'      " Where a0205 >= " & stDate0 & " And a0205 <= " & stDDate & _
'      " And ax201(+) = a0201  And ax202(+) = a0202" & _
'      " And ax209 Is Not Null And (substr(ax205, 1, 2) = '41' Or ax205 = '7121')" & _
'      " And ax207>0 And not( ax205='4191' or ax205='4192'" & _
'      " or ((ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131') And instr(ax213||' ','結餘')>0))" & _
'      " group by ax209"
'   'Modified by Lydia 2018/08/23  +判斷離職人員
'   'stVTable1 = "select ax209 V1C0,sum(decode(a0205," & stDDate & ", ax207)) V1C1" & _
'      ",sum(ax207) V1C2" & _
'      " From acc020, acc021" & _
'      " Where a0205 >= " & stDate0 & " And a0205 <= " & stDDate & _
'      " And ax201(+) = a0201  And ax202(+) = a0202" & _
'      " And ax209 Is Not Null And (substr(ax205, 1, 2) = '41' Or ax205 = '7121')" & _
'      " And ax207>0 And not( ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0)" & _
'      " group by ax209"
'   'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
'   'Memo by Lydia 2022/06/06 若抓取的會計科目有變更，請一併修改每日批次strMenu64為相同條件
'   stVTable1 = "select ax209 V1C0,sum(decode(a0205," & stDDate & ", ax207)) V1C1" & _
'      ",sum(ax207) V1C2" & _
'      " From acc020, acc021,staff" & _
'      " Where a0205 >= " & stDate0 & " And a0205 <= " & stDDate & _
'      " And ax201(+) = a0201  And ax202(+) = a0202" & _
'      " And ax209 Is Not Null And (substr(ax205, 1, 1) = '4' Or ax205 = '7121')" & _
'      " And ax207>0 And not( ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) And ST01(+)=AX209" & stCon & _
'      " group by ax209"
'
'   '已收款資料
'   'Modified by Lydia 2018/08/23 +判斷離職人員
'   'stVTable2 = "select df01 V2C0,sum(decode(df02," & stDDate & ",0,df03)) V2C1" & _
'      ",sum(decode(df02," & stDDate & ",df03,0)) V2C2" & _
'      ",decode(Max(df02)," & stDDate & ",'*',null) V2C3 From DailyFeat" & _
'      " Where DF02>=" & stDate0 & " And df02 <= " & stDDate & " group by df01"
'   stVTable2 = "select df01 V2C0,sum(decode(df02," & stDDate & ",0,df03)) V2C1" & _
'      ",sum(decode(df02," & stDDate & ",df03,0)) V2C2" & _
'      ",decode(Max(df02)," & stDDate & ",'*',null) V2C3 From STAFF, DailyFeat" & _
'      " Where DF01(+)=ST01 And DF02>=" & stDate0 & " And df02 <= " & stDDate & stCon & " group by df01"
'
'   '已簽約資料
'   'Modified by Lydia 2018/08/23 +判斷離職人員stCon
'   stVTable3 = "select df01 V3C0,sum(decode(df02," & stDDate & ",0,NVL(df04,0)-NVL(DF03,0))) V3C1" & _
'      ",sum(decode(df02," & stDDate & ",df04,0)) V3C2" & _
'      ",decode(Max(df02)," & stDDate & ",'*',null) V3C3 From STAFF,DailyFeat" & _
'      " Where DF01(+)=ST01 And df02 <= " & stDDate & stCon & " group by df01"
'
'   '先不抓銷帳及扣點數資料
'   '銷帳資料 94/6以後才要
'   'Modified by Lydia 2018/08/23 +判斷離職人員
'   'stVTable4 = "SELECT A0K20 V4C0,SUM(DECODE(A0S03," & stDDate & ",0,A1U07)) V4C1" & _
'      ",SUM(DECODE(A0S03," & stDDate & ",A1U07,0)) V4C2" & _
'      " From ACC0S0, ACC0K0, ACC1U0 WHERE A0S04='1' And A0S03>=940601 And A0S03 <= " & stDDate & _
'      " And A0K01(+)=A0S02 And A0K20 IS NOT NULL And A1U01(+)=A0S01 GROUP BY A0K20"
'   stVTable4 = "SELECT A0K20 V4C0,SUM(DECODE(A0S03," & stDDate & ",0,A1U07)) V4C1" & _
'      ",SUM(DECODE(A0S03," & stDDate & ",A1U07,0)) V4C2" & _
'      " From ACC0S0, ACC0K0, ACC1U0,STAFF WHERE A0S04='1' And A0S03>=940601 And A0S03 <= " & stDDate & _
'      " And A0K01(+)=A0S02 And A0K20 IS NOT NULL And A1U01(+)=A0S01 And ST01(+)=A0K20 " & stCon & " GROUP BY A0K20"
'
'   '扣點數資料 94/6以後才要
'   'modify by sonia 2014/1/21 取消a0201='1'條件
'   'Modified by Morgan 2015/3/27 科目條件改與每日業績點數輸入一致
'   'modify by sonia 2015/4/20 加不含'4194'科目,不含結餘傳票不要加科目限制
'   'stVTable5 = "select ax209 V5C0,sum(DECODE(A0205," & stDDate & ",0,ax206)) V5C1" & _
'      ",sum(DECODE(A0205," & stDDate & ",ax206,0)) V5C2" & _
'      " from acc020, acc021 where ax201(+) = a0201  And ax202(+) = a0202" & _
'      " And A0205>=940601 And A0205>=" & stDate0 & " And A0205<=" & stDDate & _
'      " And ax209 is not null And ax206>0 And (substr(ax205, 1, 2) = '41' Or ax205 = '7121')" & _
'      " And not (  (ax205='4191' or ax205='4192')" & _
'      " or ( (ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131') And instr(ax213||' ','結餘')>0 ) )" & _
'      " group by ax209"
'   'Modified by Lydia 2018/08/23 +判斷離職人員
'   'stVTable5 = "select ax209 V5C0,sum(DECODE(A0205," & stDDate & ",0,ax206)) V5C1" & _
'      ",sum(DECODE(A0205," & stDDate & ",ax206,0)) V5C2" & _
'      " from acc020, acc021 where ax201(+) = a0201  And ax202(+) = a0202" & _
'      " And A0205>=940601 And A0205>=" & stDate0 & " And A0205<=" & stDDate & _
'      " And ax209 is not null And ax206>0 And (substr(ax205, 1, 2) = '41' Or ax205 = '7121')" & _
'      " And not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) " & _
'      " group by ax209"
'   'Modify by Amy 2019/08/01 增加創新業務組用收入 420101  原:substr(ax205, 1, 2) = '41'
'   stVTable5 = "select ax209 V5C0,sum(DECODE(A0205," & stDDate & ",0,ax206)) V5C1" & _
'      ",sum(DECODE(A0205," & stDDate & ",ax206,0)) V5C2" & _
'      " from acc020, acc021,STAFF where ax201(+) = a0201  And ax202(+) = a0202" & _
'      " And A0205>=940601 And A0205>=" & stDate0 & " And A0205<=" & stDDate & _
'      " And ax209 is not null And ax206>0 And (substr(ax205, 1, 1) = '4' Or ax205 = '7121')" & _
'      " And not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) And ST01(+)=AX209" & stCon & _
'      " group by ax209"
'
'On Error GoTo ErrHnd
'
'   'ADD BY SONIA 2014/6/11 **** 小真說江副所長指示：國外部：實收=財務入帳,為日報,月報統一故改跑日報時先更新每日業績點數輸入,資料自2014/6/1起   *****
'   '但因仍有尾數差異,日報表之實收仍於SetExceptCell改為財務入帳數字
'   cnnConnection.BeginTrans
'
'   'Modify by Amy 2018/11/28 本日入帳原TO_CHAR(ROUND(SUM(NVL(V1C1,0))/1000,3),'9999999.000') 導致與月報不合需扣點數,1080101後國外部改區分FCP及FCT(此有修改frm210108 也要修改)
'   'modify by sonia 2021/1/19 2021/1/1起F4102改F4104及F4105,F4103改F4106及F4107,再加V1C0做小計
'   If Val(txtCloseDate) > Val(每日業務點數FCPFCT啟用日) Then
'        strExc(0) = "SELECT V1C0,TO_CHAR(ROUND(NVL(SUM(NVL(V1C1,0)),0)/1000,2),'9999999.00') 本日入帳,1 Sort from STAFF," & _
'                    "(" & stVTable1 & ") VT1 WHERE SUBSTR(ST15,1,1)='F' And V1C0(+)=ST01 And V1C0 in ('F4102','F4104','F4105') group by V1C0 " & _
'                "Union SELECT V1C0,TO_CHAR(ROUND(NVL(SUM(NVL(V1C1,0)),0)/1000,2),'9999999.00') 本日入帳,2 Sort from STAFF," & _
'                    "(" & stVTable1 & ") VT1 WHERE SUBSTR(ST15,1,1)='F' And V1C0(+)=ST01 And V1C0 in ('F4103','F4106','F4107') group by V1C0 "
'        intRun = 2
'        If Val(txtCloseDate) > 1100101 Then intRun = 4
'        ReDim stF410X(1 To intRun) As Variant
'        ReDim stF410XVal(1 To intRun) As Variant
'   Else
'        strExc(0) = "SELECT TO_CHAR(ROUND(NVL(SUM(NVL(V1C1,0)),0)/1000,2),'9999999.00') 本日入帳 from STAFF," & _
'                    "(" & stVTable1 & ") VT1 WHERE SUBSTR(ST15,1,1)='F' And V1C0(+)=ST01 "
'        intRun = 1
'   End If
'   intI = 1: stF4100Amt = 0
'   'modify by sonia 2021/1/19 2021/1/1起F4102改F4104及F4105,F4103改F4106及F4107,故先刪除當日所有F410%資料
'   'stF410X(1) = 0: stF410X(2) = 0
'   'Add by Amy 2021/02/23 +if 輸每日業務點數FCPFCT啟用日前資料會錯
'   If Val(txtCloseDate) > Val(每日業務點數FCPFCT啟用日) Then
'    For i = 1 To intRun
'       stF410X(i) = ""
'       stF410XVal(i) = 0
'    Next i
'   End If
'   strSql = "delete DailyFeat WHERE DF01 like 'F410%' And DF02=" & stDDate
'   cnnConnection.Execute strSql
'   'end 2021/1/19
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      RsTemp.MoveFirst
'      i = 1  'add by sonia 2021/1/19
'      Do While Not RsTemp.EOF
'         If Val(txtCloseDate) > Val(每日業務點數FCPFCT啟用日) Then
'            'modify by sonia 2021/1/19 2021/1/1起F4102改F4104及F4105,F4103改F4106及F4107
'            'stF410X(Val(RsTemp.Fields("Sort"))) = Val("" & RsTemp(0))
'            stF410X(i) = "" & RsTemp(0)
'            stF410XVal(i) = Val("" & RsTemp(1))
'            i = i + 1
'            'end 2021/1/19
'         Else
'            stF4100Amt = Val("" & RsTemp(0))
'         End If
'         RsTemp.MoveNext
'      Loop
'   'add by sonia 2021/1/19  當日沒收款也要寫DailyFeat
'   Else
'      If intRun = 4 Then
'         For i = 1 To intRun
'            stF410X(i) = "F410" & i + 3
'         Next i
'      ElseIf intRun = 2 Then
'         For i = 1 To intRun
'            stF410X(i) = "F410" & i + 1
'         Next i
'      End If
'   End If
'   For i = 1 To intRun
'      'stVal = stF4100Amt
'      'If intRun > 1 Then stVal = stF410X(i)
'      'strExc(0) = "SELECT * from DailyFeat WHERE DF01='F410" & IIf(Val(txtCloseDate) > Val(每日業務點數FCPFCT啟用日), i + 1, i - 1) & "' And DF02=" & stDDate
'      'intI = 1
'      'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      'If intI = 1 Then
'      '   strSql = "UPDATE DailyFeat SET DF03=" & stVal & ",DF04=" & stVal & " WHERE DF01='F410" & IIf(Val(txtCloseDate) > Val(每日業務點數FCPFCT啟用日), i + 1, i - 1) & "' And DF02=" & stDDate
'      'Else
'      '   strSql = "insert into DailyFeat (df01,df02,df03,df04) values('F410" & IIf(Val(txtCloseDate) > Val(每日業務點數FCPFCT啟用日), i + 1, i - 1) & "'," & stDDate & "," & stVal & "," & stVal & ")"
'      'End If
'      'cnnConnection.Execute strSql
'      If intRun = 1 And i = 1 Then
'         strSql = "insert into DailyFeat (df01,df02,df03,df04,df06,df07,df08) values('F4100'," & stDDate & "," & stF4100Amt & "," & stF4100Amt & ",'QPGMR'," & strSrvDate(1) & "," & Mid(Right("000000" & ServerTime, 6), 1, 4) & ")"
'         cnnConnection.Execute strSql
'      Else
'         If stF410X(i) <> "" Then
'            strSql = "insert into DailyFeat (df01,df02,df03,df04,df06,df07,df08) values('" & stF410X(i) & "'," & stDDate & "," & stF410XVal(i) & "," & stF410XVal(i) & ",'QPGMR'," & strSrvDate(1) & "," & Mid(Right("000000" & ServerTime, 6), 1, 4) & ")"
'            cnnConnection.Execute strSql
'         End If
'      End If
'   Next i
'   'end 2018/11/28
'   cnnConnection.CommitTrans
'   'END 2014/6/11
'
'   'Modify by Morgan 2009/7/31 國外部判斷部門為 F 字頭( 原來抓 F41 )
'   'Modify by Morgan 2007/4/2 不再控制員工是否在職否則過去資料將無法正確表示
'   'Modify by Morgan 2007/10/18 加巨京
'   '******** 2014/6/11 小真說江副所長指示：國外部：實收=財務入帳,為日報,月報統一故改跑日報時先更新每日業績點數輸入,資料自2014/6/1起   *****
'   '但因仍有尾數差異,日報表之實收仍於SetExceptCell改為財務入帳數字
'   'MODIFY BY SONIA 2016/1/13 取消S22加入S24
'   'Modified by Lydia 2016/02/15 因為中區組織變動(取消S22加入S24),改成公用常數與每日批次共用(frmAutoBatchDay.strMenu64)
'   'strSql = "SELECT MAX(DECODE(ST15,'W','其他','X','巨京','Z','國外部',A0902)) 項目" & _
'      ",TO_CHAR(NVL(SUM(PE04),0),'9999999.00') 本月目標" & _
'      ",TO_CHAR(ROUND(NVL(SUM(PE04),0)/" & stWorkDays & ",2),'9999999.00') 每日應收" & _
'      ",TO_CHAR(ROUND(NVL(SUM(PE04),0)*" & stWorkDay & "/" & stWorkDays & ",2),'9999999.00') 合計應收" & _
'      ",TO_CHAR(ROUND(NVL(SUM(NVL(V1C1,0)-NVL(V5C2,0)),0)/1000,2),'9999999.00') 本日入帳" & _
'      ",TO_CHAR(ROUND(NVL(SUM(NVL(V1C2,0)-NVL(V5C1,0)-NVL(V5C2,0)),0)/1000,2),'9999999.00') 本月入帳" & _
'      ",TO_CHAR(NVL(SUM(V2C2),0)-ROUND(NVL(SUM(V5C2),0)/1000,2),'9999999.00') 本日實收" & _
'      ",TO_CHAR(NVL(SUM(V2C1),0)-ROUND(NVL(SUM(V5C1),0)/1000,2)+NVL(SUM(V2C2),0)-ROUND(NVL(SUM(V5C2),0)/1000,2),'9999999.00') 本月實收" & _
'      ",TO_CHAR(NVL(SUM(V3C1),0)-ROUND(NVL(SUM(V4C1),0)/1000,2)+NVL(SUM(V3C2),0)-NVL(SUM(V2C2),0)-ROUND(NVL(SUM(V4C2),0)/1000,2),'9999999.00') 已簽約" & _
'      " FROM ( SELECT ST01,DECODE(SUBSTR(ST15,1,1),'F','Z',DECODE(ST15,'P29','X',DECODE(INSTR('S11,S13,S14,S15,S21,S23,S24,S29,S31,S41',ST15),0,'W',ST15))) ST15" & _
'      " FROM STAFF WHERE ST15 IS NOT NULL) STAFF,ACC090,PERFORMANCE" & _
'      ",(" & stVTable1 & ") VT1,(" & stVTable2 & ") VT2,(" & stVTable3 & ") VT3" & _
'      ",(" & stVTable4 & ") VT4,(" & stVTable5 & ") VT5" & _
'      " WHERE A0901(+)=ST15 And PE01(+)=ST01 And PE02(+)='TOT' And PE03(+)=" & stWYM & _
'      " And V1C0(+)=ST01 And V2C0(+)=ST01 And V3C0(+)=ST01 And V4C0(+)=ST01 And V5C0(+)=ST01" & _
'      " GROUP BY ST15"
'   'Modified by Lydia 2018/08/23 判斷離職前一天的月份要顯示(離職日為1號則當月不用) => stCon
'   'strSql = "SELECT MAX(DECODE(ST15,'W','其他','X','巨京','Z','國外部',A0902)) 項目" & _
'      ",TO_CHAR(NVL(SUM(PE04),0),'9999999.00') 本月目標" & _
'      ",TO_CHAR(ROUND(NVL(SUM(PE04),0)/" & stWorkDays & ",2),'9999999.00') 每日應收" & _
'      ",TO_CHAR(ROUND(NVL(SUM(PE04),0)*" & stWorkDay & "/" & stWorkDays & ",2),'9999999.00') 合計應收" & _
'      ",TO_CHAR(ROUND(NVL(SUM(NVL(V1C1,0)-NVL(V5C2,0)),0)/1000,2),'9999999.00') 本日入帳" & _
'      ",TO_CHAR(ROUND(NVL(SUM(NVL(V1C2,0)-NVL(V5C1,0)-NVL(V5C2,0)),0)/1000,2),'9999999.00') 本月入帳" & _
'      ",TO_CHAR(NVL(SUM(V2C2),0)-ROUND(NVL(SUM(V5C2),0)/1000,2),'9999999.00') 本日實收" & _
'      ",TO_CHAR(NVL(SUM(V2C1),0)-ROUND(NVL(SUM(V5C1),0)/1000,2)+NVL(SUM(V2C2),0)-ROUND(NVL(SUM(V5C2),0)/1000,2),'9999999.00') 本月實收" & _
'      ",TO_CHAR(NVL(SUM(V3C1),0)-ROUND(NVL(SUM(V4C1),0)/1000,2)+NVL(SUM(V3C2),0)-NVL(SUM(V2C2),0)-ROUND(NVL(SUM(V4C2),0)/1000,2),'9999999.00') 已簽約" & _
'      " FROM ( SELECT ST01," & AutoBatchSalesArea & ") ST15" & _
'      " FROM STAFF WHERE ST15 IS NOT NULL) STAFF,ACC090,PERFORMANCE" & _
'      ",(" & stVTable1 & ") VT1,(" & stVTable2 & ") VT2,(" & stVTable3 & ") VT3" & _
'      ",(" & stVTable4 & ") VT4,(" & stVTable5 & ") VT5" & _
'      " WHERE A0901(+)=ST15 And PE01(+)=ST01 And PE02(+)='TOT' And PE03(+)=" & stWYM & _
'      " And V1C0(+)=ST01 And V2C0(+)=ST01 And V3C0(+)=ST01 And V4C0(+)=ST01 And V5C0(+)=ST01" & _
'      " GROUP BY ST15"
'   'Modify by Amy 2018/11/28 1080101後巨京一定顯示及國外部改區分FCP及FCT
'   'Modify by Amy 2019/07/12 巨京值併入其他顯示,巨京欄位改顯示客服組
'   If Val(txtCloseDate) > Val(每日業務點數FCPFCT啟用日) And Val(txtCloseDate) < 1100101 Then
'        'Modify by Amy 2021/02/23 S29/客服組/顧服組 改都列示,F部門非F41XX員編需列於「其他」(ex:11001月A6043陳蒲璇 有點數需列於個人)
'        'strSql = "SELECT MAX(DECODE(ST15,'W','其他','X','客服組','Z1','FCP','Z2','FCT',A0902)) 項目"
'        'modify by sonia 2021/1/14 F4104及F4105也列入F4102
'        'stSalesArea = "DECODE(SUBSTR(ST15,1,1),'F',Decode(ST01,'F4102','Z1','F4104','Z1','F4105','Z1','Z2'),DECODE(ST15,'W10','X',DECODE(INSTR('S11,S13,S14,S15,S21,S22,S23,S24,S31,S41',ST15),0,'W',ST15)))"
'        strSql = "SELECT MAX(DECODE(ST15,'W1','客服組','W2','顧服組','X','其他','Z1','FCP','Z2','FCT',A0902)) 項目"
'        stSalesArea = "Decode(st01,'F4102','Z1','F4103','Z2','W1001','W1','W2001','W2',Decode(InStr('S11,S13,S14,S15,S21,S22,S23,S24,S29,S31,S41',ST15),0,'X',ST15))"
'   'Added by Lydia 2021/08/10 日期條件>1100801時，中三區不要出現
'   ElseIf Val(txtCloseDate) >= 1100801 Then
'        strSql = ",'W3','" & StaffQuery("P2005") & "'"
'        stSalesArea = ",'P2005','W3'"
'        'Add by Amy 2022/04/11 +MCP
'        If Val(txtCloseDate) >= 1110401 Then
'            strSql = strSql & ",'W4','" & Replace(StaffQuery("P1005"), "專利", " 專利 ") & "'"
'            stSalesArea = stSalesArea & ",'P1005','W4'"
'        End If
'        strSql = "SELECT MAX(DECODE(ST15,'W1','客服組','W2','顧服組'" & strSql & ",'X','其他','Z1','" & StaffQuery("F4104") & "','Z2','" & StaffQuery("F4105") & "','Z3','" & StaffQuery("F4106") & "','Z4','" & StaffQuery("F4107") & "',A0902)) 項目"
'        stSalesArea = "Decode(st01,'F4102','Z1','F4104','Z1','F4105','Z2','F4106','Z3','F4107','Z4','W1001','W1','W2001','W2'" & stSalesArea & ",Decode(InStr('S11,S13,S14,S15,S21,S22,S24,S29,S31,S41',ST15),0,'X',ST15))"
'        'Add by Amy 2022/07/12 11107月起不再顯示客服組W10
'        If Val(txtCloseDate) >= 1110701 Then
'            strSql = Replace(strSql, ",'W1','客服組'", "")
'            stSalesArea = Replace(stSalesArea, ",'W1001','W1'", "")
'        End If
'   'end 2021/08/10
'   'Modify by Amy 2021/02/17 F4102/F4103 拆成F4104~07
'   ElseIf Val(txtCloseDate) >= 1100101 Then
'        'Modify by Amy 2021/02/23 S29/顧服(原列於「其他」)拆開顯示,11004月後增加MCT,F部門非F41XX員編需列於「其他」(ex:11001月A6043陳蒲璇 有點數需列於個人)
'        strSql = "": stSalesArea = ""
'        If Val(txtCloseDate) >= 1100401 Then
'            strSql = ",'W3','" & StaffQuery("P2005") & "'"
'            stSalesArea = ",'P2005','W3'"
'        End If
'        strSql = "SELECT MAX(DECODE(ST15,'W1','客服組','W2','顧服組'" & strSql & ",'X','其他','Z1','" & StaffQuery("F4104") & "','Z2','" & StaffQuery("F4105") & "','Z3','" & StaffQuery("F4106") & "','Z4','" & StaffQuery("F4107") & "',A0902)) 項目"
'        'stSalesArea = "DECODE(SUBSTR(ST15,1,1),'F',Decode(ST01,'F4102','Z1','F4104','Z1','F4105','Z2','F4106','Z3','Z4'),DECODE(ST15,'W10','X',DECODE(INSTR('S11,S13,S14,S15,S21,S22,S23,S24,S29,S31,S41',ST15),0,'W',ST15)))"
'        stSalesArea = "Decode(st01,'F4102','Z1','F4104','Z1','F4105','Z2','F4106','Z3','F4107','Z4','W1001','W1','W2001','W2'" & stSalesArea & ",Decode(InStr('S11,S13,S14,S15,S21,S22,S23,S24,S29,S31,S41',ST15),0,'X',ST15))"
'        'end 2021/02/23
'   Else
'         'Modify by Amy 2021/02/23 S29/客服組/顧服組 改都列示
'         'strSql = "SELECT MAX(DECODE(ST15,'W','其他','X','客服組','Z','國外部',A0902)) 項目"
'         'stSalesArea = AutoBatchSalesArea & ")"
'         strSql = "SELECT MAX(DECODE(ST15,'W1','客服組','W2','顧服組','X','其他','Z','國外部',A0902)) 項目"
'         stSalesArea = "Decode(SUBSTR(ST15,1,1),'F','Z',Decode(st01,'W1001','W1','W2001','W2',Decode(InStr('S11,S13,S14,S15,S21,S22,S23,S24,S29,S31,S41',ST15),0,'X',st15)))"
'   End If
'   'Modify By Sindy 2020/9/1 +,ST15
'   strSql = strSql & _
'      ",TO_CHAR(NVL(SUM(PE04),0),'9999999.00') 本月目標" & _
'      ",TO_CHAR(ROUND(NVL(SUM(PE04),0)/" & stWorkDays & ",2),'9999999.00') 每日應收" & _
'      ",TO_CHAR(ROUND(NVL(SUM(PE04),0)*" & stWorkDay & "/" & stWorkDays & ",2),'9999999.00') 合計應收" & _
'      ",TO_CHAR(ROUND(NVL(SUM(NVL(V1C1,0)-NVL(V5C2,0)),0)/1000,2),'9999999.00') 本日入帳" & _
'      ",TO_CHAR(ROUND(NVL(SUM(NVL(V1C2,0)-NVL(V5C1,0)-NVL(V5C2,0)),0)/1000,2),'9999999.00') 本月入帳" & _
'      ",TO_CHAR(NVL(SUM(V2C2),0)-ROUND(NVL(SUM(V5C2),0)/1000,2),'9999999.00') 本日實收" & _
'      ",TO_CHAR(NVL(SUM(V2C1),0)-ROUND(NVL(SUM(V5C1),0)/1000,2)+NVL(SUM(V2C2),0)-ROUND(NVL(SUM(V5C2),0)/1000,2),'9999999.00') 本月實收" & _
'      ",TO_CHAR(NVL(SUM(V3C1),0)-ROUND(NVL(SUM(V4C1),0)/1000,2)+NVL(SUM(V3C2),0)-NVL(SUM(V2C2),0)-ROUND(NVL(SUM(V4C2),0)/1000,2),'9999999.00') 已簽約,ST15" & _
'      " FROM ( SELECT ST01," & stSalesArea & " ST15" & _
'      " FROM STAFF WHERE ST15 IS NOT NULL  " & stCon & " ) STAFF," & _
'      " ACC090,PERFORMANCE,(" & stVTable1 & ") VT1,(" & stVTable2 & ") VT2,(" & stVTable3 & ") VT3,(" & stVTable4 & ") VT4,(" & stVTable5 & ") VT5" & _
'      " WHERE A0901(+)=ST15 And PE01(+)=ST01 And PE02(+)='TOT' And PE03(+)=" & stWYM & _
'      " And V1C0(+)=ST01 And V2C0(+)=ST01 And V3C0(+)=ST01 And V4C0(+)=ST01 And V5C0(+)=ST01" & _
'      " GROUP BY ST15"
'    'end 2018/11/28
'   CheckOC3
'   With AdoRecordSet3
'      .CursorLocation = adUseClient
'      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'      If .RecordCount > 0 Then
'         'Modify by Morgan 2007/10/18 加巨京
'         'iXCol = 13
'         'Mark by Lydia 2017/10/12
'         'iXCol = 14
'         'end 2007/10/18
'         'Call SetDataListWidth(True)
'         grdDataList.Visible = False
'         List1.Clear 'Add By Sindy 2020/9/1
'         Do While Not .EOF
'            'Modify by Amy 2023/06/01 發現「解析業務區」"其他"會抓到"中區其他"
'            List1.AddItem .Fields(9) & " " & Replace(.Fields(0), " ", "") 'Add By Sindy 2020/9/1 代號+Grid顯示之名稱
'            For iCol = 2 To iXCol
'            'end 2007/10/18
'               If grdDataList.TextMatrix(0, iCol) = "" & .Fields(0) Then
'                  Exit For
'               End If
'            Next
'            If iCol <= iXCol Then
'               For iRow = 1 To 7 '8 Modify By Sindy 2020/9/1 取消已簽約
'                  grdDataList.TextMatrix(iRow, iCol) = Format(Trim("" & .Fields(iRow)), "#.00")
'                  If iRow = 3 Or iRow = 7 Then
'                     '2009/10/26 MODIFY BY SONIA 加顏色
'                     grdDataList.row = iRow: grdDataList.col = iCol: grdDataList.CellFontBold = True: grdDataList.CellBackColor = vbCyan
'                  End If
'               Next
'            End If
'            .MoveNext
'         Loop
'         Call SetExceptCell
'         Call Calculate
'         grdDataList.Visible = True
'         txtCloseDate.Tag = txtCloseDate
'      Else
'         MsgBox "無符合資料！", vbInformation
'      End If
'   End With
'
'   doQuery = True
'
'   Exit Function
'
'ErrHnd:
'   cnnConnection.RollbackTrans
'   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Function doQuery() As Boolean
Dim stVTable0 As String, stVTable1 As String, stVTable2 As String, stVTable3 As String, stVTable4 As String, stVTable5 As String, stVTB(3) As String
Dim stWorkDays As String, stWorkDay As String, stCon As String, strDeptS As String, strDeptE As String, intMaxState As Integer
Dim stSalesArea As String, stTB(1) As String, stField(2) As String, stWhere1 As String, stWhere2 As String, stWhere3 As String
Dim stGrpAX209 As String, stGrpDF01 As String, stGrpA0K20 As String, stGrpSP48ST15 As String, stGrpDeptEmp As String
Dim intRunCP18Type As Integer, intRunCP18Ch As Integer 'intRunCP18Type: 0-已收文點數/1-已收文未收款點數；intRunCP18Ch: 1-不包含ACS/2-只顯示ACS
Dim stFieldF4100 As String, stWhrF4100 As String '10801月前國外部資料
   
   If GetWorkDay(txtCloseDate, stWorkDay, stWorkDays) = True Then
      lblWorkDays = stWorkDays
      lblWorkDay = stWorkDay
   Else
      Exit Function
   End If
   grdDataList.Clear
   Call SetGridTitleRow(False)
   Call SetGridColData(1)
  
   '統計日
   stDDate = Val(txtCloseDate)
   '統計月份
   stDate0 = Val(stDDate) \ 100 & "01"
   
   stWDate = TransDate(txtCloseDate, 2)
   stTDate = Format(stWDate - 19110000)
   stWYM = Left(stWDate, 6)
   stTYM = Format(stWYM - 191100)
   '比照每日業績點數輸入(frm210103),排除特定員工代號(ex.中所20011,20021,20031)
   stCon = " And ST01>'60000' "
   '判斷離職前一天的月份要顯示(離職日為1號則當月不用)
   stCon = stCon & " And (st04='1' OR substr( decode(ST51,null,'',to_char(to_date(st51,'yyyymmdd')-1,'yyyymmdd')) ,1,6)>= " & Left(TransDate(stDate0, 2), 6) & ") "
   stTB(0) = ",Staff,Salespoint"
   stGrpSP48ST15 = " Group by Decode(SP48,null,ST15,SP48)"
 
   '已入帳資料 (Memo by Lydia 2022/06/06 若抓取的會計科目有變更，請一併修改每日批次strMenu64為相同條件)
   stVTable1 = "Select ax209 V1C0,Sum(Decode(a0205," & stDDate & ", ax207)) V1C1,Sum(ax207) V1C2" & _
                           " From acc020,acc021" & _
                           " Where a0205 >= " & stDate0 & " And a0205 <= " & stDDate & " And ax201(+) = a0201  And ax202(+) = a0202" & _
                           " And ax209 Is Not Null And (substr(ax205, 1, 1) = '4' Or ax205 = '7121')" & _
                           " And ax207>0 And not( ax205='4191' or ax205='4192' or ax205='4194' or InStr(ax213||' ','結餘')>0) "
   stGrpAX209 = " Group by AX209 "
   
   '已收款資料 （DailyFeat 目前有輸為S、Ｆ部門及W1001）
   stVTable2 = "Select df01 V2C0,Sum(Decode(df02," & stDDate & ",0,df03)) V2C1" & _
                                 ",Sum(Decode(df02," & stDDate & ",df03,0)) V2C2,Decode(Max(df02)," & stDDate & ",'*',null) V2C3 " & _
                           " From DailyFeat" & _
                           " Where DF02>=" & stDate0 & " And df02 <= " & stDDate
   stGrpDF01 = " Group by DF01 "
         
   '已簽約資料
   stVTable3 = "Select df01 V3C0,Sum(Decode(df02," & stDDate & ",0,Nvl(df04,0)-Nvl(DF03,0))) V3C1" & _
                                 ",Sum(Decode(df02," & stDDate & ",df04,0)) V3C2,Decode(Max(df02)," & stDDate & ",'*',null) V3C3 " & _
                           " From DailyFeat" & _
                           " Where df02 <= " & stDDate
         
   '先不抓銷帳及扣點數資料
   '銷帳資料 94/6以後才要
   stVTable4 = "Select A0K20 V4C0,Sum(Decode(A0S03," & stDDate & ",0,A1U07)) V4C1,Sum(Decode(A0S03," & stDDate & ",A1U07,0)) V4C2" & _
                           " From ACC0S0,ACC0K0,ACC1U0" & _
                           " Where A0S04='1' And A0S03>=940601 And A0S03 <= " & stDDate & _
                           " And A0K01(+)=A0S02 And A0K20 is not null And A1U01(+)=A0S01 "
   stGrpA0K20 = " Group by A0K20 "
      
   '扣點數資料 94/6以後才要
   stVTable5 = "Select ax209 V5C0,Sum(Decode(A0205," & stDDate & ",0,ax206)) V5C1,Sum(Decode(A0205," & stDDate & ",ax206,0)) V5C2" & _
                           " From acc020,acc021" & _
                           " Where ax201(+) = a0201  And ax202(+) = a0202" & _
                           " And A0205>=940601 And A0205>=" & stDate0 & " And A0205<=" & stDDate & _
                           " And ax209 is not null And ax206>0 And (substr(ax205, 1, 1) = '4' Or ax205 = '7121')" & _
                           " And not (ax205='4191' or ax205='4192' or ax205='4194' or InStr(ax213||' ','結餘')>0) "
 
   '******** 2014/6/11 小真說江副所長指示：國外部：實收=財務入帳,為日報,月報統一故改跑日報時先更新每日業績點數輸入,資料自2014/6/1起   *****
   '但因仍有尾數差異,日報表之實收仍於SetExceptCell改為財務入帳數字
   If InsertDailyFeat(stWorkDays, stWorkDay, stVTable1 & stGrpAX209, stCon) = True Then
      Load frmpic002
      frmpic002.Label1.Caption = "資料計算中...請稍候..."
      frmpic002.Show
      frmpic002.ZOrder 0: DoEvents
   '****** 寫入暫存檔 ******
      If InStr(";" & strShowDept(0), ";F4100") > 0 Then
         stFieldF4100 = "Decode(SubStr(FieldNo,1,3),'F41','F4100',FieldNo) as mField"
         If Replace(";" & strShowDept(0), ";F4100", "") = MsgText(601) Then
            stWhrF4100 = " And SubStr(FieldNo,1,3)='F41' "
         Else
            'F410X
            stWhrF4100 = " And ( SubStr(FieldNo,1,3)='F41' Or FieldNo In('" & Replace(strShowDept(0), ";", "','") & " )"
         End If
      End If
      '[有]W30-開發組資料(目前不會有目標) ,則併入W20-顧服組顯示;[不是]S/W/P/F部門 者列於[其他]
      stField(0) = "ID,StateNo,DeptNo,EmpNo,R001"
      stWhere1 = " And ID='" & strUserNum & "' And StateNo='0' "
      stWhere2 = " And PE02(+)='TOT' And PE03(+)=" & stWYM & " And Nvl(PE04,0)>0 "
      
      '目前需顯示的部門(欄)
      i = 1
      stVTB(0) = "Select '" & strUserNum & "','" & i & "',DeptNo,EmpNo From R210107 Where SubStr(DeptNo,1,1)='S' " & stWhere1
      '目標
      stVTB(1) = "Select PE01,Sum(Nvl(PE04,0)) as PE04 From PerForMance"
      
   '*** [StateNo=1] 本月目標=R001欄 ***
      '[非]S部門（點數輸入部門,以寫入暫存檔部門為主）
      If stWhrF4100 <> MsgText(601) Then
         stVTB(2) = Replace(Replace(stVTB(1), "PE01", Replace(stFieldF4100, "FieldNo", "PE01")), "mField", "PE01") & _
                             " Where " & Replace(Mid(stWhrF4100, 5), "FieldNo", "PE01") & stWhere2 & " Group by Decode(SubStr(PE01,1,3),'F41','F4100',PE01) "
      Else
         stVTB(2) = stVTB(1) & " Where PE01 in ('" & Replace(strShowDept(0), ";", "','") & "')" & stWhere2 & " Group by PE01 "
      End If
      stVTB(3) = Replace(UCase(stVTB(0)), UCase("SubStr(DeptNo,1,1)='S'"), "SubStr(DeptNo,1,1)<>'S'") & _
                        " And EmpNo In('" & Replace(strShowDept(0), ";", "','") & "') "
      strSql = "Insert Into R210107 (" & stField(0) & ") " & _
                     "Select a.*,PE04 From (" & stVTB(3) & " ) a,(" & stVTB(2) & ") " & _
                     " Where EmpNo=PE01(+) "
      adoTaie.Execute strSql
      'S部門
      stVTB(2) = Replace(UCase(stVTB(1)), "SELECT PE01,", "Select Decode(SP48,null,ST15,SP48) as Dept,") & stTB(0) & _
                        " Where SubStr(Decode(SP48,null,ST15,SP48),1,1)='S' And PE01=ST01(+) And PE01=SP02(+) And " & stWYM & "=SP01(+) " & stWhere2 & _
                        " Group by Decode(SP48,null,ST15,SP48) "
      strSql = "Insert Into R210107 (" & stField(0) & ") " & _
                     "Select a.*,PE04 From (" & stVTB(0) & " And DeptNo<>'SZZ' ) a,(" & stVTB(2) & ") " & _
                     " Where DeptNo=Dept(+) "
      adoTaie.Execute strSql
      '其他
      stWhere3 = ""
      If stWhrF4100 <> MsgText(601) Then
         stWhere3 = ",'F4101','F4102','F4103'"
      End If
      stVTB(2) = Replace(UCase(stVTB(1)), "SELECT PE01,", "Select 'OZZ' as Dept,") & stTB(0) & _
                        " Where SubStr(Decode(SP48,null,ST15,SP48),1,1)<>'S' And PE01 Not In ('" & Replace(strShowDept(0), ";", "','") & "'" & stWhere3 & ") " & _
                        " And PE01=ST01(+) And PE01=SP02(+) And " & stWYM & "=SP01(+) " & stWhere2
      strSql = "Insert Into R210107 (" & stField(0) & ") " & _
                     "Select a.*,PE04 From (" & Replace(stVTB(0), "='S'", "='O'") & " ) a,(" & stVTB(2) & ") " & _
                     " Where DeptNo=Dept(+) "
      adoTaie.Execute strSql
      
   '*** [StateNo=1] 每日應收=R002欄 / 本月累計應收=R003欄  ***
      stField(1) = "To_Char(Round(R001/" & stWorkDays & ",2),'9999999.00') 每日應收" & _
                            ",To_Char(Round(R001*" & stWorkDay & "/" & stWorkDays & ",2),'9999999.00') 本月累計應收"
      '依部門更新
      strSql = "Update R210107 o Set (R002,R003)=" & _
                      "( Select " & stField(1) & " From R210107 t  Where o.EmpNo=t.EmpNo " & Replace(UCase(stWhere1), UCase("StateNo='0'"), "StateNo='1'") & _
                      ") Where 1=1 " & Replace(Replace(UCase(stWhere1), " AND ", " AND o."), UCase("StateNo='0'"), "StateNo='1'")
      adoTaie.Execute strSql
      
      i = i + 1
   '*** [StateNo=2] 本日入帳=R001欄 / 本月入帳(本月累計入帳)=R002欄 ***
      stField(0) = stField(0) & ",R002"
      stField(1) = "'" & strUserNum & "','" & i & "',DeptNo,EmpNo," & _
                            "To_Char(Round(Nvl(Sum(Nvl(V1C1,0)-Nvl(V5C2,0)),0)/1000,2),'9999999.00') 本日入帳" & _
                            ",To_Char(Round(Nvl(Sum(Nvl(V1C2,0)-Nvl(V5C1,0)-Nvl(V5C2,0)),0)/1000,2),'9999999.00') 本月入帳"
      stGrpDeptEmp = " Group by DeptNo,EmpNo "
      
      '[非]S部門（點數輸入部門,以寫入暫存檔部門為主）
      If stWhrF4100 <> MsgText(601) Then
         stVTB(2) = Replace(Replace(UCase(stVTable1), "AX209 V1C0", Replace(stFieldF4100, "FieldNo", "AX209")), "mField", "V1C0") & _
                             Replace(stWhrF4100, "FieldNo", "AX209") & " Group by Decode(SubStr(AX209,1,3),'F41','F4100',AX209) "
         stVTB(3) = Replace(Replace(UCase(stVTable5), "AX209 V5C0", Replace(stFieldF4100, "FieldNo", "AX209")), "mField", "V5C0") & _
                             Replace(stWhrF4100, "FieldNo", "AX209") & " Group by Decode(SubStr(AX209,1,3),'F41','F4100',AX209) "
      Else
         stVTB(2) = stVTable1 & " And ax209 in ('" & Replace(strShowDept(0), ";", "','") & "')" & stGrpAX209
         stVTB(3) = stVTable5 & " And ax209 in ('" & Replace(strShowDept(0), ";", "','") & "')" & stGrpAX209
      End If
      strSql = "Insert Into R210107 (" & stField(0) & ") " & _
                     "Select " & stField(1) & " From R210107,(" & stVTB(2) & "),(" & stVTB(3) & ")" & _
                     " Where SubStr(DeptNo,1,1)<>'S' And EmpNo In('" & Replace(strShowDept(0), ";", "','") & "') " & stWhere1 & _
                     " And EmpNo=V1C0(+) And EmpNo=V5C0(+) " & stGrpDeptEmp
      adoTaie.Execute strSql
      'S部門
      stWhere3 = " And SubStr(Decode(SP48,null,ST15,SP48),1,1)='S' And ax209=st01(+) And ax209=sp02(+) And " & stWYM & "=SP01(+) "
      stVTB(2) = Replace(Replace(UCase(stVTable1), "AX209 V1C0", "Decode(SP48,null,ST15,SP48) V1C0"), "FROM ACC020,ACC021", "From ACC020,ACC021" & stTB(0)) & _
                        stWhere3 & " Group by Decode(SP48,null,ST15,SP48) "
      stVTB(3) = Replace(Replace(UCase(stVTable5), "AX209 V5C0", "Decode(SP48,null,ST15,SP48) V5C0"), "FROM ACC020,ACC021", "From ACC020,ACC021" & stTB(0)) & _
                        stWhere3 & " Group by Decode(SP48,null,ST15,SP48) "
      strExc(2) = "Select " & Replace(UCase(stField(1)), UCase(",ST01"), ",ST15||'ZZ'") & _
                           " From R210107,(" & stVTB(2) & "),(" & stVTB(3) & ")" & _
                           " Where SubStr(DeptNo,1,1)='S' And DeptNo<>'SZZ' " & stWhere1 & _
                           " And DeptNo=V1C0(+) And DeptNo=V5C0(+) " & stGrpDeptEmp
      strSql = "Insert Into R210107 (" & stField(0) & ") " & strExc(2)
      adoTaie.Execute strSql
      '其他
      strExc(2) = ""
      If stWhrF4100 <> MsgText(601) Then
         strExc(2) = ",'F4101','F4102','F4103'"
      End If
      stWhere3 = Replace(stWhere3, "='S'", "<>'S'") & " And ax209 Not In('" & Replace(strShowDept(0), ";", "','") & "'" & strExc(2) & ") "
      stVTB(2) = Replace(Replace(UCase(stVTable1), "AX209 V1C0", "'OZZ' V1C0"), "FROM ACC020,ACC021", "From ACC020,ACC021" & stTB(0)) & _
                        stWhere3
      stVTB(3) = Replace(Replace(UCase(stVTable5), "AX209 V5C0", "'OZZ' V5C0"), "FROM ACC020,ACC021", "From ACC020,ACC021" & stTB(0)) & _
                        stWhere3
      strSql = "Insert Into R210107 (" & stField(0) & ") " & _
                     "Select " & stField(1) & " From R210107,(" & stVTB(2) & "),(" & stVTB(3) & ")" & _
                     " Where SubStr(DeptNo,1,1)='O' " & stWhere1 & _
                     " And DeptNo=V1C0(+) And DeptNo=V5C0(+) " & stGrpDeptEmp
      adoTaie.Execute strSql
      
      i = i + 1
   '*** [StateNo=3] 本日實收(本日收款)=R001欄 / 本月實收(本月累計收款)=R002欄 ***
      stField(1) = "'" & strUserNum & "','" & i & "',DeptNo,EmpNo," & _
                            "To_Char(Nvl(Sum(V2C2),0)-Round(Nvl(Sum(V5C2),0)/1000,2),'9999999.00') 本日實收" & _
                            ",To_Char(Nvl(Sum(V2C1),0)-Round(Nvl(Sum(V5C1),0)/1000,2)+Nvl(Sum(V2C2),0)-Round(Nvl(Sum(V5C2),0)/1000,2),'9999999.00') 本月實收"
      
      '點數輸入[W部門]照算(若有W3001歸於W2001顯示)
      stVTB(2) = Replace(UCase(stVTable2), "DF01 V2C0", "Decode(DF01,'W3001','W2001',DF01) V2C0") & " And DF01 In('W1001','W2001','W3001') " & _
                           " Group by Decode(DF01,'W3001','W2001',DF01) " '員編W開頭
      stVTB(3) = Replace(UCase(stVTable5), "AX209 V5C0", "Decode(AX209,'W3001','W2001',AX209) V5C0") & " And AX209 In('W1001','W2001','W3001') " & _
                           " Group by Decode(AX209,'W3001','W2001',AX209)"
      strSql = "Insert Into R210107 (" & stField(0) & ") " & _
                     "Select " & stField(1) & " From R210107 t,(" & stVTB(2) & "),(" & stVTB(3) & ") " & _
                     " Where SubStr(DeptNo,1,1)='W' " & stWhere1 & _
                     " And EmpNo=V2C0(+) And EmpNo=V5C0(+) " & stGrpDeptEmp
      adoTaie.Execute strSql
      
      '點數輸入F及P部門 (FXXXX/P1005/P2005 員編):[本日實收]顯示[本日入帳] / [本月實收]顯示[本月入帳],故抓[StateNo=2]
      stVTB(2) = "Select EmpNo as Emp,R001 as V1,R002 as V2 From R210107 Where SubStr(EmpNo,1,1) In('F','P') " & Replace(UCase(stWhere1), UCase("StateNo='0'"), "StateNo='2'")
      strSql = "Insert Into R210107 (" & stField(0) & ") " & _
                     "Select '" & strUserNum & "','" & i & "',DeptNo,EmpNo,V1,V2 From R210107 t,(" & stVTB(2) & ") " & _
                     " Where SubStr(EmpNo,1,1) In('F','P') " & stWhere1 & _
                     " And EmpNo=Emp(+) "
      adoTaie.Execute strSql
      'S部門
      stWhere3 = " And SubStr(Decode(SP48,null,ST15,SP48),1,1)='S' And DF01=st01(+) And DF01=sp02(+) And " & stWYM & "=SP01(+) "
      stVTB(2) = Replace(Replace(UCase(stVTable2), "DF01 V2C0", "Decode(SP48,null,ST15,SP48) V2C0"), "FROM DAILYFEAT", "From DailyFeat" & stTB(0)) & _
                         stWhere3 & stGrpSP48ST15
      stVTB(3) = Replace(Replace(UCase(stVTable5), "AX209 V5C0", "Decode(SP48,null,ST15,SP48) V5C0"), "FROM ACC020,ACC021", "From Acc020,Acc021" & stTB(0)) & _
                         Replace(UCase(stWhere3), "DF01", "AX209") & stGrpSP48ST15
     strExc(2) = "Select " & stField(1) & " From R210107 t,(" & stVTB(2) & "),(" & stVTB(3) & ") " & _
                         " Where SubStr(DeptNo,1,1)='S' And DeptNo<>'SZZ' " & stWhere1 & _
                         " And DeptNo=V2C0(+) And DeptNo=V5C0(+) " & stGrpDeptEmp
      strSql = "Insert Into R210107 (" & stField(0) & ") " & strExc(2)
      adoTaie.Execute strSql
      '其他 (其他人員[沒]DailyFeat資料):[本日實收]顯示[本日入帳] / [本月實收]顯示[本月入帳],故抓[StateNo=2]
      stVTB(2) = "Select DeptNo as Dept,R001 as V1,R002 as V2 From R210107 Where DeptNo='OZZ'" & Replace(UCase(stWhere1), UCase("StateNo='0'"), "StateNo='2'")
      strSql = "Insert Into R210107 (" & stField(0) & ") " & _
                     "Select '" & strUserNum & "','" & i & "',DeptNo,EmpNo,V1,V2 From R210107 t,(" & stVTB(2) & ") " & _
                     " Where SubStr(DeptNo,1,1)='O' " & stWhere1 & _
                     " And DeptNo=Dept(+) "
      adoTaie.Execute strSql
      
      i = i + 1
      '*** [StateNo=4] 本月累計收文=R001欄 /本月ACS累計=R002欄 ***
      intRunCP18Type = 0: intRunCP18Ch = 1
      stField(1) = "'" & strUserNum & "','" & i & "',DeptNo,EmpNo,"
      '[非]S部門 F4104/F4105(需寫入不顯示)
      '國外部 F4100/F4102/F4103
      If stWhrF4100 <> MsgText(601) Or InStr(";" & strShowDept(0), ";F4102") > 0 Or InStr(";" & strShowDept(0), ";F4103") > 0 Then
         If stWhrF4100 <> MsgText(601) Then
            strExc(2) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "F10", "F29", , , , , intRunCP18Ch, Me.Name)
            If Val(strExc(2)) <> 0 Then strExc(2) = Format(strExc(2), "#.00")
            strExc(3) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "F10", "F29", , , , , intRunCP18Ch + 1, Me.Name)
            If Val(strExc(3)) <> 0 Then strExc(2) = Format(strExc(3), "#.00")
            strSql = "Insert Into R210107 (" & stField(0) & ") " & _
                      "Select " & stField(1) & strExc(2) & "," & strExc(3) & " From R210107 " & " Where EmpNo='F4100' " & stWhere1
            adoTaie.Execute strSql
         Else
            'F4102
            strExc(2) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "F20", "F29", , , , , intRunCP18Ch, Me.Name)
            If Val(strExc(2)) <> 0 Then strExc(2) = Format(strExc(2), "#.00")
            strExc(3) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "F20", "F29", , , , , intRunCP18Ch + 1, Me.Name)
            If Val(strExc(3)) <> 0 Then strExc(2) = Format(strExc(3), "#.00")
            strSql = "Insert Into R210107 (" & stField(0) & ") " & _
                           "Select " & stField(1) & strExc(2) & "," & strExc(3) & " From R210107 " & " Where EmpNo='F4102' " & stWhere1
            adoTaie.Execute strSql
            'F4103
            strExc(2) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "F10", "F19", , , , , intRunCP18Ch, Me.Name)
            If Val(strExc(4)) <> 0 Then strExc(2) = Format(strExc(2), "#.00")
            strExc(3) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "F10", "F19", , , , , intRunCP18Ch + 1, Me.Name)
            If Val(strExc(3)) <> 0 Then strExc(2) = Format(strExc(3), "#.00")
            strSql = "Insert Into R210107 (" & stField(0) & ") " & _
                           "Select " & stField(1) & strExc(2) & "," & strExc(3) & " From R210107 " & " Where EmpNo='F4103' " & stWhere1
            adoTaie.Execute strSql
         End If
      Else
         '本月累計收文 F4104
         strExc(2) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "F20", "F29", "F4104", , , , intRunCP18Ch, Me.Name)
         If Val(strExc(2)) <> 0 Then
            strF410x(4) = strExc(2)
            strExc(2) = Format(strExc(2), "#.00")
         End If
         '本月ACS累計 F4104
         strExc(3) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "F20", "F29", "F4104", , , , intRunCP18Ch + 1, Me.Name)
         If Val(strExc(3)) <> 0 Then
            strF410xAcs(4) = strExc(3)
            strExc(3) = Format(strExc(3), "#.00")
         End If
         strSql = "Insert Into R210107 (" & stField(0) & ") " & _
                         "Select " & stField(1) & strExc(2) & "," & strExc(3) & " From R210107 " & " Where EmpNo='F4104' " & stWhere1
         adoTaie.Execute strSql
         '本月累計收文 F4105
         strExc(2) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "F20", "F29", "F4105", , , , intRunCP18Ch, Me.Name)
         If Val(strExc(2)) <> 0 Then
            strF410x(5) = strExc(2)
            strExc(2) = Format(strExc(2), "#.00")
         End If
         '本月ACS累計 F4105
         strExc(3) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "F20", "F29", "F4105", , , , intRunCP18Ch + 1, Me.Name)
         If Val(strExc(3)) <> 0 Then
            strF410xAcs(4) = strExc(3)
            strExc(3) = Format(strExc(3), "#.00")
         End If
         strSql = "Insert Into R210107 (" & stField(0) & ") " & _
                         "Select " & stField(1) & strExc(2) & "," & strExc(3) & " From R210107 Where EmpNo='F4105' " & stWhere1
         adoTaie.Execute strSql
         '本月累計收文 F4106
         strExc(2) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "F10", "F19", "F4106", , , , intRunCP18Ch, Me.Name)
         If Val(strExc(2)) <> 0 Then strExc(2) = Format(strExc(2), "#.00")
         '本月ACS累計 F4106
         strExc(3) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "F20", "F29", "F4106", , , , intRunCP18Ch + 1, Me.Name)
         strSql = "Insert Into R210107 (" & stField(0) & ") " & _
                         "Select " & stField(1) & strExc(2) & "," & strExc(3) & " From R210107 Where EmpNo='F4106' " & stWhere1
         adoTaie.Execute strSql
         '本月累計收文 F4107
         strExc(2) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "F10", "F19", "F4107", , , , intRunCP18Ch, Me.Name)
         If Val(strExc(2)) <> 0 Then strExc(2) = Format(strExc(2), "#.00")
         '本月ACS累計 F4107
         strExc(3) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "F20", "F29", "F4107", , , , intRunCP18Ch + 1, Me.Name)
         strSql = "Insert Into R210107 (" & stField(0) & ") " & _
                         "Select " & stField(1) & strExc(2) & "," & strExc(3) & " From R210107 Where EmpNo='F4107' " & stWhere1
         adoTaie.Execute strSql
      End If
      
      '本月累計收文 P1005
      strExc(2) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "P10", "P19", "P1005", , , , intRunCP18Ch, Me.Name)
      If Val(strExc(2)) <> 0 Then strExc(2) = Format(strExc(2), "#.00")
      '本月ACS累計 P1005
      strExc(3) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "P20", "P29", "P1005", , , , intRunCP18Ch + 1, Me.Name)
      strSql = "Insert Into R210107 (" & stField(0) & ") " & _
                      "Select " & stField(1) & strExc(2) & "," & strExc(3) & " From R210107 Where EmpNo='P1005' " & stWhere1
      adoTaie.Execute strSql
      '本月累計收文 P2005
      strExc(2) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "P20", "P29", "P2005", , , , intRunCP18Ch, Me.Name)
      If Val(strExc(2)) <> 0 Then strExc(2) = Format(strExc(2), "#.00")
      '本月ACS累計 P2005
      strExc(3) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "F20", "F29", "P2005", , , , intRunCP18Ch + 1, Me.Name)
      strSql = "Insert Into R210107 (" & stField(0) & ") " & _
                      "Select " & stField(1) & strExc(2) & "," & strExc(3) & " From R210107 Where EmpNo='P2005' " & stWhere1
      adoTaie.Execute strSql
      
      '本月累計收文 W1001（11107月起不再顯示客服組）
      strExc(2) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "W10", "W19", "W1001", , , , intRunCP18Ch, Me.Name)
      If Val(strExc(2)) <> 0 Then strExc(2) = Format(strExc(2), "#.00")
      '本月ACS累計 W1001
      strExc(3) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "W10", "W19", "W1001", , , , intRunCP18Ch + 1, Me.Name)
      If Val(strExc(3)) <> 0 Then strExc(3) = Format(strExc(3), "#.00")
      '10807~11106 W1001需顯示,11107月後有才顯示 (避免總數不合)
      If (Val(Left(Val(stDate0 + 19110000), 6)) - 191100 >= 10807 And Val(Left(Val(stDate0 + 19110000), 6)) - 191100 < 11107) _
       Or (Val(Left(Val(stDate0 + 19110000), 6)) - 191100 >= 11107 And (Val(strExc(2)) + Val(strExc(3)) <> 0)) Then
         strSql = "Insert Into R210107 (" & stField(0) & ") " & _
                         "Select " & stField(1) & strExc(2) & "," & strExc(3) & " From R210107 Where EmpNo='W1001' " & stWhere1
         adoTaie.Execute strSql
      End If
    
      '本月累計收文 W2001（W3001有值歸入W2001 計算）
      strExc(2) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "W20", "W29", "W2001", , , , intRunCP18Ch, Me.Name)
      strExc(4) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "W30", "W39", "W3001", , , , intRunCP18Ch, Me.Name)
      If Val(strExc(2)) + Val(strExc(4)) <> 0 Then strExc(2) = Format(Val(strExc(2)) + Val(strExc(4)), "#.00")
      '本月ACS累計 W2001
      strExc(3) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "W20", "W29", "W2001", , , , intRunCP18Ch + 1, Me.Name)
      strExc(5) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "W20", "W29", "W3001", , , , intRunCP18Ch + 1, Me.Name)
      If Val(strExc(3)) + Val(strExc(5)) <> 0 Then strExc(3) = Format(Val(strExc(3)) + Val(strExc(5)), "#.00")
      strSql = "Insert Into R210107 (" & stField(0) & ") " & _
                     "Select " & stField(1) & strExc(2) & "," & strExc(3) & " From R210107 Where EmpNo='W2001' " & stWhere1
      adoTaie.Execute strSql
      'S部門
      '本月累計收文
      strExc(1) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "S10", "S99", , strExc(2), , , intRunCP18Ch, Me.Name, , True)
      strExc(2) = "Select " & stField(1) & "收文點數" & _
                           " From R210107,(" & strExc(2) & ")" & _
                           " Where SubStr(DeptNo,1,1)='S' And DeptNo<>'SZZ' And DeptNo=ST15(+) " & stWhere1
      strSql = "Insert Into R210107 (" & Replace(stField(0), ",R002", "") & ") " & strExc(2)
      adoTaie.Execute strSql
      '本月ACS累計
      strExc(1) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "S10", "S99", , strExc(2), , , intRunCP18Ch + 1, Me.Name, , True)
      strSql = "Update R210107 Set R002=(" & Replace(Replace(UCase(strExc(2)), "SELECT ST15,", "Select "), UCase("Group by ST15"), " And DeptNo=ST15 Group by ST15 ") & _
                      " ) Where SubStr(DeptNo,1,1)='S' " & Replace(stWhere1, "StateNo='0'", "StateNo='" & i & "'")
      adoTaie.Execute strSql
      '其他=全所-本月累計收文（其他以外的欄位相加）
      strExc(2) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, , , , , , , intRunCP18Ch, Me.Name)
      strExc(4) = Pub_GetField("R210107", Mid(Replace(stWhere1, "StateNo='0'", "StateNo='" & i & "'"), 5), "Sum(R001)")
      '本月ACS累計
      strExc(3) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, , , , , , , intRunCP18Ch + 1, Me.Name)
      strExc(5) = Pub_GetField("R210107", Mid(Replace(stWhere1, "StateNo='0'", "StateNo='" & i & "'"), 5), "Sum(R002)")
      strSql = "Insert Into R210107 (" & stField(0) & ") " & _
                     "Select " & stField(1) & Val(strExc(2)) - Val(strExc(4)) & "," & Val(strExc(3)) - Val(strExc(5)) & " From R210107 " & _
                     " Where DeptNo='OZZ' " & stWhere1
      adoTaie.Execute strSql
      stSumMonGT = strExc(2)
      
      i = i + 1
      '*** [StateNo=5] 未收款=已收文未收款點數 / ACS未收款 ***
      intRunCP18Type = 1: intRunCP18Ch = 1
      stField(1) = "'" & strUserNum & "','" & i & "',ST15,ST01,"
      'Memo by Lydia 2021/07/30 (總計)未收款點數,不要傳入日期止日;因為不管收據日期只要是未收款都列入計算,但不改模組,怕將來有其他需求
      '[非]S部門（只抓W部門,且不含ACS）
      '未收款 W1001（11107月起不再顯示客服組,但不確定之後會不會再使用,還是先寫入）
      strExc(2) = PUB_CountCP18(intRunCP18Type, "", "", "W10", "W19", "W1001", , , , intRunCP18Ch, Me.Name)
      If Val(strExc(2)) <> 0 Then strExc(2) = Format(strExc(2), "#.00")
      'ACS未收款 W1001
      strExc(3) = PUB_CountCP18(intRunCP18Type, stDate0, stDDate, "W10", "W19", "W1001", , , , intRunCP18Ch + 1, Me.Name)
      '10807~11106 W1001需顯示,11107月後有才顯示 (避免總數不合)
      If (Val(Left(Val(stDate0 + 19110000), 6)) - 191100 >= 10807 And Val(Left(Val(stDate0 + 19110000), 6)) - 191100 < 11107) _
       Or (Val(Left(Val(stDate0 + 19110000), 6)) - 191100 >= 11107 And (Val(strExc(2)) + Val(strExc(3)) <> 0)) Then
         'W1001顯示於智權部小計前
          strSql = "Insert Into R210107 (" & stField(0) & ") " & _
                         "Select " & stField(1) & strExc(2) & "," & strExc(3) & " From Staff Where ST01='W1001' "
         adoTaie.Execute strSql
      End If
      '未收款 W2001 (W3001有資料列於W2001)
      strExc(2) = PUB_CountCP18(intRunCP18Type, "", "", "W20", "W29", "W2001", , , , intRunCP18Ch, Me.Name)
      strExc(4) = PUB_CountCP18(intRunCP18Type, "", "", "W30", "W39", "W3001", , , , intRunCP18Ch, Me.Name)
      If Val(strExc(2)) + Val(strExc(4)) <> 0 Then strExc(2) = Format(Val(strExc(2)) + Val(strExc(4)), "#.00")
      'ACS未收款 W2001 (W3001有資料列於W2001)
      strExc(3) = PUB_CountCP18(intRunCP18Type, "", "", "W20", "W29", "W2001", , , , intRunCP18Ch + 1, Me.Name)
      strExc(5) = PUB_CountCP18(intRunCP18Type, "", "", "W30", "W39", "W3001", , , , intRunCP18Ch + 1, Me.Name)
      If Val(strExc(3)) + Val(strExc(5)) <> 0 Then strExc(3) = Format(Val(strExc(3)) + Val(strExc(5)), "#.00")
      strSql = "Insert Into R210107 (" & stField(0) & ") " & _
                     "Select " & stField(1) & strExc(2) & "," & strExc(3) & " From Staff Where ST01='W2001' "
      adoTaie.Execute strSql
      'S部門 未收款=R001,ACS未收款=R002
      strExc(1) = PUB_CountCP18(intRunCP18Type, "", "", "S10", "S99", , strExc(2), , , intRunCP18Ch, Me.Name, , True)
      strExc(1) = PUB_CountCP18(intRunCP18Type, "", "", "S10", "S99", , strExc(3), , , intRunCP18Ch + 1, Me.Name, , True)
      strExc(2) = Replace(UCase(strExc(2)), "SUM(RECPNT) 未收款", "Round(SUM(RECPNT)/1000,2) as a1")
      strExc(3) = Replace(UCase(strExc(3)), "ST15,SUM(RECPNT) 未收款", "ST15 as ACSST15,Round(SUM(RECPNT)/1000,2) as b1")
      strExc(1) = "Select  " & Replace(stField(1), "ST15,ST01,", "A0921,A0921||'ZZ',Nvl(a1,0),Nvl(b1,0)") & _
                          " From Acc090New,(" & strExc(2) & "),(" & strExc(3) & ") " & _
                          " Where SubStr(a0921,1,1)='S' And (a1<>0 or b1<>0) And a0921=ST15(+) And a0921=ACSST15(+) "
      strSql = "Insert Into R210107 (" & stField(0) & ") " & strExc(1)
      adoTaie.Execute strSql
      
      i = i + 1
      '*** [StateNo=6] 已  簽  約=R012欄 ***
      stField(0) = Replace(stField(0), ",R002", "")
      stField(1) = "'" & strUserNum & "','" & i & "',ST15,ST01," & _
                           "To_Char(Nvl(Sum(V3C1),0)-Round(Nvl(Sum(V4C1),0)/1000,2)+Nvl(Sum(V3C2),0)-Nvl(Sum(V2C2),0)-Round(Nvl(Sum(V4C2),0)/1000,2),'9999999.00') 已簽約"

      '[非]S部門（點數輸入部門）
      If stWhrF4100 <> MsgText(601) Then
         stVTB(1) = Replace(Replace(UCase(stVTable2), "DF01 V2C0", Replace(stFieldF4100, "FieldNo", "DF01")), "mField", "V2C0") & _
                             Replace(stWhrF4100, "FieldNo", "DF01") & " Group by Decode(SubStr(DF01,1,3),'F41','F4100',DF01) "
         stVTB(2) = Replace(Replace(UCase(stVTable3), "DF01 V3C0", Replace(stFieldF4100, "FieldNo", "DF01")), "mField", "V3C0") & _
                             Replace(stWhrF4100, "FieldNo", "DF01") & " Group by Decode(SubStr(DF01,1,3),'F41','F4100',DF01) "
         stVTB(3) = Replace(Replace(UCase(stVTable4), "A0K20 V4C0", Replace(stFieldF4100, "FieldNo", "A0K20")), "mField", "V4C0") & _
                             Replace(stWhrF4100, "FieldNo", "A0K20") & " Group by Decode(SubStr(A0K20,1,3),'F41','F4100',A0K20) "
      Else
         stVTB(1) = stVTable2 & " And DF01 in('" & Replace(strShowDept(0), ";", "','") & "') " & stGrpDF01
         stVTB(2) = stVTable3 & " And DF01 in('" & Replace(strShowDept(0), ";", "','") & "') " & stGrpDF01
         stVTB(3) = stVTable4 & " And A0K20 in('" & Replace(strShowDept(0), ";", "','") & "') " & stGrpA0K20
      End If
      strSql = "Insert Into R210107 (" & stField(0) & ") " & _
                        "Select " & Replace(UCase(stField(1)), UCase(",ST15"), ",Decode(ST01,'W3001','W20',ST15) as ST15") & _
                        " From Staff,(" & stVTB(1) & "),(" & stVTB(2) & "),(" & stVTB(3) & ") " & _
                        " Where SubStr(ST15,1,1)<>'S' And ST01 In('" & Replace(strShowDept(0), ";", "','") & "') " & _
                           " And ST01=V2C0(+) And ST01=V3C0(+) And ST01=V4C0(+) Group by Decode(ST01,'W3001','W20',ST15),ST01"
      adoTaie.Execute strSql
      'S部門
      stWhere3 = " And SubStr(Decode(SP48,null,ST15,SP48),1,1)='S' And df01=st01(+) And df01=sp02(+) And " & stWYM & "=SP01(+) "
      stVTB(1) = Replace(Replace(UCase(stVTable2), "DF01 V2C0", "Decode(SP48,null,ST15,SP48) V2C0"), "FROM DAILYFEAT", "From DailyFeat" & stTB(0)) & _
                           stWhere3 & " Group by Decode(SP48,null,ST15,SP48) "
      stVTB(2) = Replace(Replace(UCase(stVTable3), "DF01 V3C0", "Decode(SP48,null,ST15,SP48) V3C0"), "FROM DAILYFEAT", "From DailyFeat" & stTB(0)) & _
                           stWhere3 & " Group by Decode(SP48,null,ST15,SP48) "
      stVTB(3) = Replace(Replace(UCase(stVTable4), "A0K20 V4C0", "Decode(SP48,null,ST15,SP48) V4C0"), ",ACC1U0", ",ACC1U0" & stTB(0)) & _
                           Replace(UCase(stWhere3), "DF01", "A0K22") & " Group by Decode(SP48,null,ST15,SP48) "
      strExc(2) = "Select " & Replace(UCase(stField(1)), UCase(",ST01"), ",ST15||'ZZ'") & _
                           " From STaff,(" & stVTB(1) & "),(" & stVTB(2) & "),(" & stVTB(3) & ") " & _
                           " Where SubStr(ST15,1,1)='S' And ST15=V2C0(+) And ST15=V3C0(+) And ST15=V4C0(+) " & _
                           " Group by ST15,ST15||'ZZ'"
      strSql = "Insert Into R210107 (" & stField(0) & ") " & strExc(2)
      adoTaie.Execute strSql
      '其他
      stWhere3 = " And SubStr(Decode(SP48,null,ST15,SP48),1,1)<>'S' And A0K20=st01(+) And A0K20=sp02(+) And " & stWYM & "=SP01(+) "
      stVTB(1) = "Select 'OZZ' V2C0,0 V2C1,0 V2C2 From Dual "
      stVTB(2) = "Select 'OZZ' V3C0,0 V3C1,0 V3C2 From Dual "
      stVTB(3) = Replace(Replace(UCase(stVTable4), "A0K20 V4C0", "'OZZ' V4C0"), ",ACC1U0", ",ACC1U0" & stTB(0)) & _
                           stWhere3 & " And A0K20 not in('" & Replace(strShowDept(0), ";", "','") & "') "
      strExc(2) = "Select " & Replace(UCase(stField(1)), UCase(",ST15,ST01"), ",'OZZ','OZZ'") & _
                           " From Staff,(" & stVTB(1) & "),(" & stVTB(2) & "),(" & stVTB(3) & ")" & _
                           " Where SubStr(ST15,1,1)<>'S' And ST01 Not In('" & Replace(strShowDept(0), ";", "','") & "') " & _
                           " And ST01=V2C0(+) And ST01=V3C0(+) And ST01=V4C0(+) "
      strSql = "Insert Into R210107 (" & stField(0) & ") " & strExc(2)
      adoTaie.Execute strSql
      intMaxState = i
      
      '畫面當月無資料,但有 未收款(已簽約目前不顯示,故排除) 資料,於新增資料時再加入部門欄位
      'ex:1130217 當下S10部門並無目標及點數,但有未收款,故加顯示S10
      'S部門
      strExc(3) = "Select DeptNo From R210107 Where " & Mid(stWhere1, 5)
      strExc(2) = "Select '" & strUserNum & "','0',A0922,A0921,A0921||'ZZ',0 From Acc090New" & _
                        ",(Select Distinct DeptNo From R210107 Where ID='" & strUserNum & "' And StateNo<>'0' And StateNo<>'6' And SubStr(DeptNo,1,1)='S' " & _
                           " And Nvl(R001,0)+Nvl(R002,0)+Nvl(R003,0)<>0 And DeptNo Not In(" & strExc(3) & " And SubStr(DeptNo,1,1)='S') " & _
                           ") Where DeptNo=a0921(+) "
      strSql = "Insert Into R210107 (ID,StateNo,DeptName,DeptNo,EmpNo,R001) " & strExc(2)
      adoTaie.Execute strSql
      '[非]S部門 (部門名稱抓st02)
      strExc(3) = "Select EmpNo From R210107 Where " & Mid(stWhere1, 5)
      strExc(2) = "Select '" & strUserNum & "','0',ST02,ST15,ST01,0 From Staff" & _
                        ",(Select Distinct EmpNo From R210107 Where ID='" & strUserNum & "' And StateNo<>'0' And StateNo<>'6' And SubStr(DeptNo,1,1)<>'S' " & _
                              " And Nvl(R001,0)+Nvl(R002,0)+Nvl(R003,0)<>0 And EmpNo Not In(" & strExc(3) & " And SubStr(DeptNo,1,1)<>'S') " & _
                           ") Where EmpNo=ST01(+) "
      strSql = "Insert Into R210107 (ID,StateNo,DeptName,DeptNo,EmpNo,R001) " & strExc(2)
      adoTaie.Execute strSql
      
      '更新顯示順序
      Call UpdSeqNo(2)
      '智權部小計 /全所合計
      For i = 1 To intMaxState
         If i = 1 Then
            strExc(2) = ",R002,R003 "
            strExc(3) = ",Sum(Nvl(R002,0)),Sum(Nvl(R003,0))"
         Else
            strExc(2) = ",R002 "
            strExc(3) = ",Sum(Nvl(R002,0))"
         End If
         strSql = "Insert Into R210107 (" & stField(0) & strExc(2) & ") " & _
                     "Select '" & strUserNum & "','" & i & "','SZZ','S9999',Sum(Nvl(R001,0))" & strExc(3) & _
                     " From R210107 Where SubStr(DeptNo,1,1)='S' And SubStr(DeptNo,-2,2)<>'ZZ' " & Replace(UCase(stWhere1), UCase("StateNo='0'"), "StateNo='" & i & "'")
         adoTaie.Execute strSql
         strSql = "Insert Into R210107 (" & stField(0) & strExc(2) & ") " & _
                     "Select '" & strUserNum & "','" & i & "','ZZZ','Z9999',Sum(Nvl(R001,0))" & strExc(3) & _
                     " From R210107 Where (SubStr(DeptNo,-2,2)<>'ZZ' Or DeptNo='OZZ')" & Replace(UCase(stWhere1), UCase("StateNo='0'"), "StateNo='" & i & "'")
         adoTaie.Execute strSql
      Next i
      '更新顯示順序
      Call UpdSeqNo(0)
   '****** End 寫入暫存檔 ******
      Unload frmpic002
      '再重抓欄位,因畫面當月無資料,但有 未收款 資料,於新增資料時再加入部門欄位
      'ex:1130217 當下S10部門並無目標及點數,但有未收款,故加顯示S10
      Call SetDataListWidth
      
      'Memo by Amy 列名及順序同arrField2
      stVTB(1) = "": stVTB(2) = "": stWhere2 = "": stField(0) = ""
      For i = 1 To intMaxState
         Select Case i
            Case 1
               stField(0) = "R001 as PE04,R002 as DayRec,R003 as MonRec"
            Case 2
               stField(0) = "R001 as DayCreDit,R002 as MonCreDit"
            Case 3
               stField(0) = "R001 as DayReality,R002 as MonReality"
            Case 4
               stField(0) = "R001 as MonGT,R002 as MonACSGT"
            Case 5
               stField(0) = "R001 as UnCollect,R002 as ACSUnCollect"
            Case 6
               stField(0) = "R001 as Sign"
         End Select
         stVTB(2) = "Select DeptNo as Dept" & i & ",EmpNo as Emp" & i & "," & stField(0) & _
                             " From R210107 Where 1=1 " & Replace(stWhere1, "StateNo='0' ", "StateNo='" & i & "' ")
         stVTB(1) = stVTB(1) & ",(" & stVTB(2) & ")"
         stWhere2 = stWhere2 & " And DeptNo=Dept" & i & "(+) And EmpNo=Emp" & i & "(+)"
      Next i
      stVTB(2) = "Select DeptName,DeptNo,EmpNo,SeqNo From R210107 Where 1=1 " & stWhere1
      strSql = "Select '        ' ,DeptName 項目" & _
         ",To_Char(PE04,'9999999.00') 本月目標" & _
         ",To_Char(DayRec,'9999999.00') 每日應收,To_Char(MonRec,'9999999.00') 本月累計應收" & _
         ",To_Char(DayCreDit,'9999999.00') 本日入帳,To_Char(MonCreDit,'9999999.00') 本月累計入帳" & _
         ",To_Char(DayReality,'9999999.00') 本日收款,To_Char(MonReality,'9999999.00') 本月累計收款" & _
         ",To_Char(MonGT,'9999999.00') 本月累計收文,To_Char(MonACSGT,'9999999.00') 本月ACS累計" & _
         ",To_Char(UnCollect,'9999999.00') 未收款,To_Char(ACSUnCollect,'9999999.00') ACS未收款" & _
         ",'' as 備註,To_Char(Sign,'9999999.00') 已簽約,SeqNo,DeptNo,EmpNo" & _
         " From (" & stVTB(2) & ")" & stVTB(1) & _
         " Where 1=1 " & stWhere2 & " Order by SeqNo,DeptNo,EmpNo "
         
      CheckOC3
      With AdoRecordSet3
         .CursorLocation = adUseClient
         .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If .RecordCount > 0 Then
            grdDataList.Visible = False
            Call Calculate
            grdDataList.Visible = True
            txtCloseDate.Tag = txtCloseDate
         Else
            MsgBox "無符合資料！", vbInformation
         End If
      End With
   End If
   doQuery = True
   
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

Private Sub Calculate_old()
'Dim iX As Integer 'Added by Lydia 2017/10/12
'Dim dblSum As Double
''Add By Sindy 2020/9/1
'Dim ii As Integer, varTmp As Variant, strDeptS As String, strDeptE As String
'Dim dblSum8 As Double, dblSum9 As Double
'Dim intWCol As Integer
'Dim strSales As String 'Add by Amy 2021/02/17
'Dim dblSubSum As Double 'Added by Lydia 2021/09/03 小計
'Dim dblSum8_ACS, dblSum9_ACS As Double, dblSubSum_ACS As Double, dblSubSum8 As Double, dblSubSum8_ACS As Double 'Add by Amy 2022/06/27
'
'   'Add By Sindy 2020/9/1
'   Load frmpic002
'   frmpic002.Label1.Caption = "資料計算中...請稍候..."
'   frmpic002.Show
'   frmpic002.ZOrder 0: DoEvents
''   If PUB_IsFormExist("frm210137") = True Then
''      Unload frm210137
''   End If
''   If PUB_IsFormExist("frm210141") = True Then
''      Unload frm210141
''   End If
'   '2020/9/1 END
'   'grdDataList.Visible = True '測式用
'   With grdDataList
'      'Modify by Morgan 2007/10/18 加巨京
'      'iXCol = 14
'      'Modified by Lydia 2017/10/12
'      'iXCol = 15
'      iX = iXCol - 1
'      'end 2007/10/18
'
'      For iRow = 1 To 8
'         dblSum = 0
'         dblSubSum = 0 'Added by Lydia 2021/09/03
'         'Modified by Lydia 2017/10/12
'         'For iCol = 2 To iXCol - 1
'         For iCol = 2 To iX - 1
'            dblSum = dblSum + Val(.TextMatrix(iRow, iCol))
'            'Added by Lydia 2021/09/03 小計
'            If iDeptSub > 0 And iCol <= iDeptSub - 1 Then
'                dblSubSum = dblSubSum + Val(.TextMatrix(iRow, iCol))
'            End If
'            'end 2021/09/03
'         Next iCol
'         'Added by Lydia 2021/09/03 小計
'         If iDeptSub > 0 Then
'             .TextMatrix(iRow, iDeptSub) = Format(dblSubSum, "#.00")
'         End If
'         'end 2021/09/03
'         '全所
'         'Modified by Lydia 2017/10/12
'         '.TextMatrix(iRow, iCol) = Format(dblSum, "#.00")
'         .TextMatrix(iRow, iX) = Format(dblSum, "#.00")
'      Next iRow
'
'      'Modified by Lydia 2017/10/12
'      '.row = 3: .col = iXCol: .CellFontBold = True
'      '.row = 7: .col = iXCol: .CellFontBold = True
'      .row = 3: .col = iX: .CellFontBold = True
'      .row = 7: .col = iX: .CellFontBold = True
'      'Added by Lydia 2021/09/03 小計
'      If iDeptSub > 0 Then
'        .row = 3: .col = iDeptSub: .CellFontBold = True
'        .row = 7: .col = iDeptSub: .CellFontBold = True
'        dblSubSum = 0
'      End If
'      'end 2021/09/03
'
'      'Add By Sindy 2020/9/1
'      dblSum8 = 0: dblSum9 = 0
'      dblSum8_ACS = 0: dblSum9_ACS = 0 'Add by Amy 2022/06/27
'      For iCol = 2 To iX '12
'         strSales = "" 'Add by Amy 2021/02/17
'         '解析業務區
'         strDeptS = ""
'         For ii = 0 To List1.ListCount - 1
'            'Modify by Amy 2023/06/01 避免抓錯,發現"其他"會抓到"中區其他"
'            'If InStr(List1.List(ii), .TextMatrix(0, iCol)) > 0 Then
'            If Mid(List1.List(ii), InStr(List1.List(ii), " ") + 1) = Replace(.TextMatrix(0, iCol), " ", "") Then
'               varTmp = Split(List1.List(ii), " ")
'               strDeptS = varTmp(0)
'               Exit For
'            End If
'         Next ii
'
'         'If strDeptS = "X" Then intWCol = iCol '其他第幾欄位'Mark by Amy 2021/02/23 有中所其他會抓錯,故Mark
'         'Modify by Amy 2021/02/23 原W為其他,X為客服組,改為W1客服組,X為其他,增加W2顧服組/W3MCT
'         If strDeptS <> "" And strDeptS <> "X" Then '其他
'            If strDeptS = "W1" Then
'               strDeptS = "W10"
'               strDeptE = "W19"
'            ElseIf strDeptS = "W2" Then
'               strDeptS = "W20"
'               strDeptE = "W29"
'            ElseIf strDeptS = "W3" Then
'               strDeptS = "P20"
'               strDeptE = "P29"
'               strSales = "P2005"
'            'Add by Amy 2023/06/01 +MCP
'            ElseIf strDeptS = "W4" Then
'               strDeptS = "P10"
'               strDeptE = "P19"
'               strSales = "P1005"
'            ElseIf strDeptS = "Z1" Then 'FCP/專利國內部F4104
'               strDeptS = "F20"
'               strDeptE = "F29"
'               If Val(txtCloseDate) >= 1100101 Then
'                    strSales = "F4104"
'               End If
'            ElseIf strDeptS = "Z2" Then 'FCT/專利日本部F4105
'               If Val(txtCloseDate) >= 1100101 Then
'                    strDeptS = "F20"
'                    strDeptE = "F29"
'                    strSales = "F4105"
'               Else
'                    strDeptS = "F10"
'                    strDeptE = "F19"
'               End If
'            'Add by Amy 2021/02/17
'            ElseIf strDeptS = "Z3" Or strDeptS = "Z4" Then
'               If strDeptS = "Z3" Then
'                    strSales = "F4106"
'               Else
'                    strSales = "F4107"
'               End If
'               strDeptS = "F10"
'               strDeptE = "F19"
'            Else
'               strDeptE = strDeptS
'            End If
'            '已收文點數:
'            iRow = 8
'            'Modify by Amy 2023/06/01 +if 專利國外部 及 專利日本部 不顯示「本月累計收文」
'            If .TextMatrix(0, iCol) = "專利國外部" Then
'               strF4104 = Format(PUB_CountCP18(0, stDate0, stDDate, strDeptS, strDeptE, strSales, , , , 1), "#.00")
'            ElseIf .TextMatrix(0, iCol) = "專利日本部" Then
'               strF4105 = Format(PUB_CountCP18(0, stDate0, stDDate, strDeptS, strDeptE, strSales, , , , 1), "#.00")
'            Else
'               'Modify by Amy 2021/02/17 +strSales
'               'Modify by Amy 2022/06/27 +intChoose參數
'               .TextMatrix(iRow, iCol) = Format(PUB_CountCP18(0, stDate0, stDDate, strDeptS, strDeptE, strSales, , , , 1), "#.00")
'            End If
'            dblSum8 = dblSum8 + Val(.TextMatrix(iRow, iCol)) 'Add by Amy 2022/06/27 改為累計(原:小計=智權部+客服組)
''            frm210137.Hide
''            frm210137.txtSalesArea = strDeptS '業務區(起)
''            frm210137.txtSalesArea1 = strDeptE '業務區(迄)
''            frm210137.txtSales = "" '智權人員ID
''            frm210137.txtCloseDate(0) = stDate0 '點數結算日(起)
''            frm210137.txtCloseDate(1) = stDDate '點數結算日(迄)
''            frm210137.cmdSearch_Click
''            If frm210137.grdDataList.Rows > 1 Then
''               '最後一筆合計
''               If Trim(frm210137.grdDataList.TextMatrix(frm210137.grdDataList.Rows - 1, 5)) <> "" Then
''                  .TextMatrix(iRow, iCol) = Format(frm210137.grdDataList.TextMatrix(frm210137.grdDataList.Rows - 1, 5), "#.00")
''               End If
''            End If
'            'Modify by Amy 2022/06/27 dblSum8 改為累計使用(與dblSum9同)
'            'Modified by Lydia 2021/09/03 改成小計=智權部+客服組
'            'dblSum8 = dblSum8 + Val(.TextMatrix(iRow, iCol))
'            If iDeptSub > 0 And iCol <= iDeptSub - 1 Then
'                'dblSum8 = dblSum8 + Val(.TextMatrix(iRow, iCol))
'                '記錄小計值-本月累計收文
'                dblSubSum8 = dblSubSum8 + Val(.TextMatrix(iRow, iCol))
'            End If
'            '本月ACS累計
'            iRow = 9
'            'Modify by Amy 2023/06/01 +if 專利國外部 及 專利日本部 不顯示「本月ACS累計」
'            If .TextMatrix(0, iCol) = "專利國外部" Then
'               strF4104Acs = Format(PUB_CountCP18(0, stDate0, stDDate, strDeptS, strDeptE, strSales, , , , 2), "#.00")
'            ElseIf .TextMatrix(0, iCol) = "專利日本部" Then
'               strF4105Acs = Format(PUB_CountCP18(0, stDate0, stDDate, strDeptS, strDeptE, strSales, , , , 2), "#.00")
'            Else
'               .TextMatrix(iRow, iCol) = Format(PUB_CountCP18(0, stDate0, stDDate, strDeptS, strDeptE, strSales, , , , 2), "#.00")
'            End If
'            dblSum8_ACS = dblSum8_ACS + Val(.TextMatrix(iRow, iCol))
'            If iDeptSub > 0 And iCol <= iDeptSub - 1 Then
'                '記錄小計值-本月ACS累計
'                dblSubSum8_ACS = dblSubSum8_ACS + Val(.TextMatrix(iRow, iCol))
'            End If
'            'end 2022/06/27
'            'end 2021/09/03
'
'            '已收文未收款點數:
'            'Modify by Amy 2021/02/17 +left,專利/ＦＣＴ
'            'Modify by Amy 2021/03/12 +MCT(P2005)
'            'Modify by Amy 2023/06/01 +MCP(P1005)
'            If Left(.TextMatrix(0, iCol), 3) <> "FCP" And Left(.TextMatrix(0, iCol), 3) <> "FCT" _
'              And Left(.TextMatrix(0, iCol), 2) <> "專利" And Left(.TextMatrix(0, iCol), 3) <> "ＦＣＴ" _
'              And .TextMatrix(0, iCol) <> StaffQuery("P2005") And Replace(.TextMatrix(0, iCol), " ", "") <> StaffQuery("P1005") Then
'               iRow = 10 'Modiffy by Amy 2022/06/27 +本月ACS累計/ACS未收款 原:iRow = 9
'                'Modified by Lydia 2021/07/30 (總計)未收款點數，請不要傳入日期止日；因為不管收據日期只要是未收款都列入計算，但不改模組，怕將來有其他需求
'               '.TextMatrix(iRow, iCol) = Format(PUB_CountCP18(1, "", stDDate, strDeptS, strDeptE), "#.00")
'               'Modify by Amy 2022/06/27 +intChoose參數/ACS未收款
'               '未收款(不是ACS)
'               .TextMatrix(iRow, iCol) = Format(PUB_CountCP18(1, "", "", strDeptS, strDeptE, , , , , 1), "#.00")
'               dblSum9 = dblSum9 + Val(.TextMatrix(iRow, iCol))
'               '未收款(不是ACS):智權部小計前欄位加總
'               If iDeptSub > 0 And iCol <= iDeptSub - 1 Then
'                    dblSubSum = dblSubSum + Val(.TextMatrix(iRow, iCol))
'               End If
'               'ACS未收款
'               iRow = 11
'               .TextMatrix(iRow, iCol) = Format(PUB_CountCP18(1, "", "", strDeptS, strDeptE, , , , , 2), "#.00")
'                dblSum9_ACS = dblSum9_ACS + Val(.TextMatrix(iRow, iCol))
'               '未收款(ACS):智權部小計前欄位加總
'               If iDeptSub > 0 And iCol <= iDeptSub - 1 Then
'                    dblSubSum_ACS = dblSubSum_ACS + Val(.TextMatrix(iRow, iCol))
'               End If
'               'end 2022/06/27
'
'   '            frm210141.Hide
'   '            frm210141.FrameDept.Visible = True
'   '            frm210141.txtSalesArea = strDeptS '業務區(起)
'   '            frm210141.txtSalesArea1 = strDeptE '業務區(迄)
'   '            frm210141.txtSales = "" '智權人員ID
'   '            frm210141.txtDate(0) = "" '點數結算日(起)
'   '            frm210141.txtDate(1) = stDDate '點數結算日(迄)
'   '            frm210141.cmdok_Click (1)
'   '            If Trim(frm210141.txtTot(0)) <> "" Then
'   '               .TextMatrix(iRow, iCol) = Format(frm210141.txtTot(0), "#.00")
'   '            End If
'               'Modify by Amy 2022/06/27 dblSum9 / dblSubSum 往上搬
''               dblSum9 = dblSum9 + Val(.TextMatrix(iRow, iCol))
''              'Added by Lydia 2021/09/03 小計
''              If iDeptSub > 0 And iCol <= iDeptSub - 1 Then
''                 dblSubSum = dblSubSum + Val(.TextMatrix(iRow, iCol))
''              End If
''              'end 2021/09/03
'              'end 2022/06/27
'            End If
'         End If
'      Next iCol
'      '全所
''      .TextMatrix(8, iX) = Format(dblSum8, "#.00")
'      'Modiffy by Amy 2022/06/27 +本月ACS累計/ACS未收款 原:TextMatrix(9, iX)=未收款
'      '.TextMatrix(8, iX) = Format(PUB_CountCP18(0, stDate0, stDDate), "#.00") '本月累計收文-全所
'      .TextMatrix(8, iX) = Format(PUB_CountCP18(0, stDate0, stDDate, , , , , , , 1), "#.00") '本月累計收文-全所
'      .TextMatrix(9, iX) = Format(PUB_CountCP18(0, stDate0, stDDate, , , , , , , 2), "#.00") '本月ACS累計-全所
'      .TextMatrix(10, iX) = Format(dblSum9, "#.00") '未收款-全所
'      .TextMatrix(11, iX) = Format(dblSum9_ACS, "#.00") '未收款ACS-全所
'      'end 2022/06/27
''      .TextMatrix(9, iX) = Format(PUB_CountCP18(1, "", stDDate), "#.00")
'
'       'Added by Lydia 2021/09/03 智權部小計
'       If iDeptSub > 0 Then
'          'Modiffy by Amy 2022/06/27 +本月ACS累計/ACS未收款 原:TextMatrix(9, iDeptSub)=未收款
'          '.TextMatrix(8, iDeptSub) = Format(dblSum8, "#.00")
'          .TextMatrix(8, iDeptSub) = Format(dblSubSum8, "#.00")
'          .TextMatrix(9, iDeptSub) = Format(dblSubSum8_ACS, "#.00") '本月ACS累計
'          .TextMatrix(10, iDeptSub) = Format(dblSubSum, "#.00")
'          .TextMatrix(11, iDeptSub) = Format(dblSubSum_ACS, "#.00") 'ACS未收款
'          'end 2022/06/27
'          'Added by Lydia 2021/09/06 小計縱向用顏色區隔
'          'Modify by Amy 2022/06/27 +本月ACS累計/ACS未收款 原:For iCol = 1 To 9
'          For iCol = 1 To 11
'             .row = iCol: .col = iDeptSub: .CellBackColor = vbYellow
'          Next iCol
'          'end 2021/09/06
'       End If
'       'end 2021/09/03
'
'      'Add by Amy 2022/06/27 其他=全所-畫面上除「其他」欄位的總合
'      '其他=全所-智權部加總
'      'Modify by Amy 2021/02/23 避免抓錯改抓iDept01 原:intWCol
'      'Modify by Amy 2023/06/01 加 - Val(strF4104) - Val(strF4105) ,專利國外部  及 專利日本部 不顯示 「本月累計收文」及「本月ACS累計」
'      .TextMatrix(8, iDept01) = Format(.TextMatrix(8, iX) - dblSum8 - Val(strF4104) - Val(strF4105), "#.00")
'      .TextMatrix(9, iDept01) = Format(.TextMatrix(9, iX) - dblSum8_ACS - Val(strF4104Acs) - Val(strF4105Acs), "#.00")
'      'end 2022/06/27
''      .TextMatrix(9, intWCol) = Format(.TextMatrix(9, iX) - dblSum9, "#.00")
'      '2020/9/1 END
'
'      'Modify By Sindy 2020/9/1 Mark
''      iRow = 9
''      '簽收合計
''      'Modified by Lydia 2017/10/12
''      'For iCol = 2 To iXCol
''      For iCol = 2 To iX
''         .TextMatrix(iRow, iCol) = Format(Val(.TextMatrix(7, iCol)) + Val(.TextMatrix(8, iCol)), "#.00")
''      Next iCol
'
'      '備註
'      'Modify by Amy 2022/06/27 +本月ACS累計/ACS未收款  原: iRow = 10
'      iRow = 12
'      'Modified by Lydia 2017/10/12
'      'For iCol = 2 To iXCol
'      For iCol = 2 To iX
'         '其他&國外部不顯示燈號
'         'Modify by Morgan 2007/10/18 加巨京
'         'If iCol = 9 Or iCol = 12 Or iCol = 13 Then
'         'Modified by Lydia 2017/10/12
'         'If iCol = 9 Or iCol = 12 Or iCol = 13 Or iCol = 14 Then
'         'Modify by Amy 2018/11/28 國外拆成FCP/FCT
'         'Modify by Amy 2021/02/17 F4102/F4103拆成F4104~07,有目標都顯示登號
'         'Modify by Amy 2021/02/18 有目標(不為0)都顯示登號
'         'Modify by Amy 2021/02/24 +iDept07
'         'Modify by Amy 2021/03/12 +iDept08
'         'Modify by Amy 2022/04/11 +iDept09-MCP
'         If (iCol = iDept01 Or iCol = iDept02 Or iCol = iDept03 Or iCol = iDept04 Or iCol = iDept05 Or iCol = iDept06 Or iCol = iDept07 Or iCol = iDept08 Or iCol = iDept09) _
'           And Val(.TextMatrix(1, iCol)) = 0 Then
'         'end 2007/10/18
'            .col = iCol: .Text = "": .CellBackColor = .BackColor
'         Else
'            'Add by Morgan 2005/5/31
'            '達成：本月實收>=本月目標
'            '      本月累計收款>=本月目標
'            If Val(.TextMatrix(7, iCol)) >= Val(.TextMatrix(1, iCol)) Then
'               .TextMatrix(iRow, iCol) = "達成"
'               .row = iRow: .col = iCol: .CellBackColor = vbGreen
'            '綠燈：本月實收>=合計應收
'            '      本月累計收款>=本月累計應收
'            ElseIf Val(.TextMatrix(7, iCol)) >= Val(.TextMatrix(3, iCol)) Then
'               .TextMatrix(iRow, iCol) = "綠燈"
'               .row = iRow: .col = iCol: .CellBackColor = vbGreen
'            'Modify By Sindy 2020/9/15 Mark
''            '黃燈：簽收合計>=合計應收
''            ElseIf Val(.TextMatrix(9, iCol)) >= Val(.TextMatrix(3, iCol)) Then
''               .TextMatrix(iRow, iCol) = "黃燈"
''               .row = iRow: .col = iCol: .CellBackColor = vbYellow
'            '紅燈：簽收合計<合計應收
'            Else
'               .TextMatrix(iRow, iCol) = "紅燈"
'               .row = iRow: .col = iCol: .CellBackColor = vbRed
'            End If
'         End If
'      Next iCol
'      'Modified by Lydia 2017/10/12
'      'lblNet = Format(Val(.TextMatrix(7, iXCol)) - Val(.TextMatrix(3, iXCol)), "#.00")
'      lblNet = Format(Val(.TextMatrix(7, iX)) - Val(.TextMatrix(3, iX)), "#.00")
'   End With
'
'   'Add By Sindy 2020/9/1
''   If PUB_IsFormExist("frm210137") = True Then
''      Unload frm210137
''   End If
''   If PUB_IsFormExist("frm210141") = True Then
''      Unload frm210141
''   End If
'   Unload frmpic002
'   '2020/9/1 END
End Sub

'Mark by Amy 2024/03/01 欄位寫入暫在檔,故不使用
'設定例外欄位內容
Private Sub SetExceptCell()
'    Dim intEnd As Integer 'Add by Amy 2018/11/28
'
'   With grdDataList
'      '其他：實收=財務入帳
'      'Modified by Lydia 2017/10/12
'      'iCol = 12
'      iCol = iDept01
'      .TextMatrix(6, iCol) = .TextMatrix(4, iCol) '本日實收
'      .TextMatrix(7, iCol) = .TextMatrix(5, iCol) '本月實收
'
'      'Modify by Amy 2022/07/12 11107月起不再顯示客服組W10
'      If Val(txtCloseDate) >= 1110701 Then
'        'Add by Morgan 2007/10/18 加巨京
'        '巨京：實收=財務入帳
'        'Modified by Lydia 2017/10/12
'        'iCol = 13
'        iCol = iDept02
'        .TextMatrix(6, iCol) = .TextMatrix(4, iCol)
'        .TextMatrix(7, iCol) = .TextMatrix(5, iCol)
'        'end 2007/10/18
'      End If
'
'      'ADD BY SONIA 2014/6/11 加國外部
'      '國外部/FCP/F4102：實收=財務入帳
'      'Modified by Lydia 2017/10/12
'      'iCol = 14
'      iCol = iDept03
'      .TextMatrix(6, iCol) = .TextMatrix(4, iCol)
'      .TextMatrix(7, iCol) = .TextMatrix(5, iCol)
'      'end 2007/10/18
'
'      'Add by Amy 2018/11/28 國外拆成FCP/FCT
'      If Val(txtCloseDate) > Val(每日業務點數FCPFCT啟用日) Then
'        iCol = iDept04
'        .TextMatrix(6, iCol) = .TextMatrix(4, iCol)
'        .TextMatrix(7, iCol) = .TextMatrix(5, iCol)
'        'Add by Amy 2021/02/17 F4102/F4103拆成F4104~07
'        If Val(txtCloseDate) >= 1100101 Then
'            iCol = iDept05
'            .TextMatrix(6, iCol) = .TextMatrix(4, iCol)
'            .TextMatrix(7, iCol) = .TextMatrix(5, iCol)
'             iCol = iDept06
'            .TextMatrix(6, iCol) = .TextMatrix(4, iCol)
'            .TextMatrix(7, iCol) = .TextMatrix(5, iCol)
'        End If
'
'        'Add by Amy 2022/04/11 +MCP
'        If Val(txtCloseDate) >= 1110401 Then
'            iCol = iDept09
'            .TextMatrix(6, iCol) = .TextMatrix(4, iCol)
'            .TextMatrix(7, iCol) = .TextMatrix(5, iCol)
'        End If
'
'        'Add by Amy 2021/02/23 +MCT
'        If Val(txtCloseDate) >= 1100401 Then
'            iCol = iDept08 'Modify by Amy 2021/03/12
'            .TextMatrix(6, iCol) = .TextMatrix(4, iCol)
'            .TextMatrix(7, iCol) = .TextMatrix(5, iCol)
'        End If
'      End If
'
'      'Modify By Sindy 2020/9/1 Mark
''      '其他&國外部已簽約=0
''      'Add by Morgan 2007/10/18 加巨京
''      'For iCol = 12 To 13
''      'Modified by Lydia 2017/10/12
''      'For iCol = 12 To 14
''      'Modify by Amy 2018/11/28 國外拆成FCP/FCT
''      intEnd = iDept03
''      If Val(txtCloseDate) > Val(每日業務點數FCPFCT啟用日) Then intEnd = iDept04
''      For iCol = iDept01 To intEnd
''      'end 2007/10/18
''         .TextMatrix(8, iCol) = ".00" '已簽約
''      Next
'   End With
End Sub

Private Sub cmdExit_Click()
 Unload Me
End Sub

Private Sub cmdPrint_Click()
   Dim stMsg As String 'Add by Amy 2023/06/01
   
   If txtCloseDate.Tag = "" Then
      'MsgBox "請先查詢後再按列印！"
      MsgBox "請先查詢後再按產生Excel！"
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   'Modify by Amy 2021/02/23 改產生Excel
'   Is17Col = False
'   If DoPrint = True Then
'      MsgBox "列印完成", vbInformation
'   End If
   If ExcelSave = True Then
        'Modify by Amy 2021/06/22 +中文字顯示路徑
        'Modify by Amy 2023/06/01
        If bolErr = True Then stMsg = "但加總有誤,請洽電腦中心"
        MsgBox "Excel已產生" & vbCrLf & "檔案存於" & strExcelPathN & vbCrLf & stMsg, vbInformation
        'end 2023/06/01
   End If
   Screen.MousePointer = vbDefault
End Sub

'Mark by Amy 2021/02/23 改成Excel,因A4已無法印成一頁
Private Function DoPrint() As Boolean
'   Dim iRowHeight As Integer
'   Dim iTableWidth As Integer, iTableHeight As Integer
'   Dim iColWidth As Integer
'   Dim iColWidth0 As Integer, iColWidth1 As Integer
'   Dim iX As Integer, iX0 As Integer, iX1 As Integer
'   Dim iY As Integer, iY0 As Integer, iY1 As Integer
'   Dim strContent As String, iContent As Integer
'   Dim iXp As Integer, iYp As Integer
'   Dim iPos As Integer 'Added by Lydia 2017/10/12
'   Dim intFontSize As Integer, intFontSize2 As Integer, bolMultLine As Boolean 'Add by Amy 2021/02/17
'
'On Error GoTo ErrHnd
'   'Add by Amy 2021/02/17
'   intFontSize = 14
'   intFontSize2 = 11
'   If Val(txtCloseDate) >= 1100101 Then
'        Is17Col = True
'        intFontSize = 12
'        intFontSize2 = 9
'   End If
'   'end 2021/02/17
'   iRowHeight = 650
'   'Modify by Morgan 2007/10/18 加巨京
'   'iXCol = 13
'   'Modified by Lydia 2017/10/12 -3包括去掉S29
'   'iXCol = 14
'   iPos = iXCol - 3
'   'end 2007/10/18
'
'   'Added by Morgan 2015/4/9 104/4起取消中區其他的目標
'   'Mark by Lydia 2017/10/12
'   'If Val(txtCloseDate) >= 1040401 Then
'   '   iXCol = iXCol - 1
'   'End If
'   ''end 2015/4/9
'   'end 2017/10/12
'
'   iColWidth0 = 400 '450
'   iColWidth1 = 1400 '1100
'   'Modified by Lydia 2017/10/12
'   'iColWidth = (15850 - iColWidth0 - iColWidth1) \ iXCol
'   'iTableWidth = iColWidth0 + iColWidth1 + iXCol * iColWidth
'   ''end 2007/10/18
'   iColWidth = (15850 - iColWidth0 - iColWidth1) \ iPos
'   iTableWidth = iColWidth0 + iColWidth1 + iPos * iColWidth
'   'end 2017/10/12
'
'   iTableHeight = 11 * iRowHeight
'   iX0 = 260: iX1 = iX0 + iTableWidth
'   iY0 = 2700: iY1 = iY0 + iTableHeight
'
'   Printer.Orientation = 2
'   Printer.FontName = "標楷體"
'   Printer.DrawWidth = 2
''   Printer.FillColor = RGB(255, 255, 255)
''   Printer.FillStyle = 1
''   Printer.ColorMode = 1
'
'   '橫線
'   For iRow = 0 To 11
'      If iRow < 2 Or iRow = 4 Or iRow > 9 Then
'         iX = iX0
'      Else
'         iX = iX0 + iColWidth0
'      End If
'      iY = iY0 + iRow * iRowHeight
'      Printer.Line (iX, iY)-(iX1, iY)
'      If iRow = 3 Or iRow = 7 Then
'         Printer.Line (iX + 20, iY + 20)-(iX1 - 20, iY + iRowHeight - 20), RGB(200, 200, 200), BF
'      End If
'   Next
'   '豎線
'   Printer.Line (iX0, iY0)-(iX0, iY1)
'   Printer.Line (iX0 + iColWidth0, iY0 + iRowHeight)-(iX0 + iColWidth0, iY1 - iRowHeight)
'   'Modified by Lydia 2017/10/12
'   'For iCol = 0 To iXCol
'   For iCol = 0 To iPos
'      iX = iX0 + iColWidth0 + iColWidth1 + iCol * iColWidth
'      Printer.Line (iX, iY0)-(iX, iY1)
'   Next
'
'   Printer.FontSize = 26
'   Printer.FontBold = False
'   strContent = CompNameQuery(2) 'Modify by Amy 2020/03/27 原："台一國際專利商標事務所"
'   iX = Printer.ScaleWidth / 2 - Printer.TextWidth(strContent) / 2
'   iY = 800
'   Printer.CurrentX = iX: Printer.CurrentY = iY
'   Printer.Print strContent
'
'   Printer.FontSize = 18
'   Printer.FontBold = False
'   strContent = Format(Left(stWYM, 4) - 1911) & "年" & Val(Mid(stWYM, 5)) & "月業務目標及達成通知"
'   iX = Printer.ScaleWidth / 2 - Printer.TextWidth(strContent) / 2
'   iY = 1500
'   Printer.CurrentX = iX: Printer.CurrentY = iY
'   Printer.Print strContent
'
'   Printer.FontSize = 12
'   Printer.FontBold = False
'   strContent = "本月工作天 "
'   iX = iX0 + 300
'   iY = iY0 - Printer.TextHeight(strContent) - 100
'   Printer.CurrentX = iX: Printer.CurrentY = iY
'   Printer.Print strContent
'
'   '接著印
'   iX = iX + Printer.TextWidth(strContent)
'   Printer.FontSize = 18
'   Printer.FontBold = True
'   strContent = lblWorkDays
'   Printer.CurrentX = iX: Printer.CurrentY = iY - 60
'   Printer.Print strContent
'
'   iX = iX + Printer.TextWidth(strContent)
'   Printer.FontSize = 12
'   Printer.FontBold = False
'   strContent = " 天"
'   Printer.CurrentX = iX: Printer.CurrentY = iY
'   Printer.Print strContent
'
'   iContent = 0
'   Printer.FontSize = 12
'   Printer.FontBold = False
'   strContent = "今天為第 "
'   iContent = iContent + Printer.TextWidth(strContent)
'   Printer.FontSize = 18
'   Printer.FontBold = True
'   strContent = lblWorkDay
'   iContent = iContent + Printer.TextWidth(strContent)
'   strContent = " 工作天"
'   Printer.FontSize = 12
'   Printer.FontBold = False
'   iContent = iContent + Printer.TextWidth(strContent)
'
'   Printer.FontSize = 12
'   Printer.FontBold = False
'   strContent = "今天為第 "
'   iX = Printer.ScaleWidth / 2 - iContent / 2
'   iY = iY0 - Printer.TextHeight(strContent) - 100
'   Printer.CurrentX = iX: Printer.CurrentY = iY
'   Printer.Print strContent
'
'   iX = iX + Printer.TextWidth(strContent)
'   Printer.FontSize = 18
'   Printer.FontBold = True
'   strContent = lblWorkDay
'   Printer.CurrentX = iX: Printer.CurrentY = iY - 60
'   Printer.Print strContent
'
'   iX = iX + Printer.TextWidth(strContent)
'   Printer.FontSize = 12
'   Printer.FontBold = False
'   strContent = " 工作天"
'   Printer.CurrentX = iX: Printer.CurrentY = iY
'   Printer.Print strContent
'
'   Printer.FontSize = 12
'   Printer.FontBold = False
'   strContent = "日期：" & Val(Mid(stWYM, 5)) & "月" & Val(Mid(stWDate, 7)) & "日"
'   iX = 14000
'   iY = iY0 - Printer.TextHeight(strContent) - 100
'   Printer.CurrentX = iX: Printer.CurrentY = iY
'   Printer.Print strContent
'
'   With grdDataList
'      '固定行1
'      Printer.FontSize = 18
'      Printer.FontBold = True
'      '項目
'      iRow = 0
'      strContent = .TextMatrix(iRow, 1)
'      iX = iX0 + ((iColWidth0 + iColWidth1) / 2 - Printer.TextWidth(strContent) / 2)
'      iY = iY0 + iRowHeight / 2 - Printer.TextHeight(strContent) / 2
'      Printer.CurrentX = iX: Printer.CurrentY = iY
'      Printer.Print strContent
'
'      '目標
'      iRow = 1
'      strContent = "目"
'      iX = iX0 + ((iColWidth0) / 2 - Printer.TextWidth(strContent) / 2)
'      iY = iY0 + iRow * iRowHeight + iRowHeight / 2 - Printer.TextHeight(strContent) / 2
'      Printer.CurrentX = iX: Printer.CurrentY = iY
'      Printer.Print strContent
'
'      iRow = 3
'      strContent = "標"
'      iY = iY0 + iRow * iRowHeight + iRowHeight / 2 - Printer.TextHeight(strContent) / 2
'      Printer.CurrentX = iX: Printer.CurrentY = iY
'      Printer.Print strContent
'
'      iRow = 4
'      strContent = "達"
'      iY = iY0 + iRow * iRowHeight + iRowHeight / 2 - Printer.TextHeight(strContent) / 2
'      Printer.CurrentX = iX: Printer.CurrentY = iY
'      Printer.Print strContent
'
'      iRow = 9
'      strContent = "成"
'      iY = iY0 + iRow * iRowHeight + iRowHeight / 2 - Printer.TextHeight(strContent) / 2
'      Printer.CurrentX = iX: Printer.CurrentY = iY
'      Printer.Print strContent
'
'      '備註
'      iRow = 10
'      strContent = .TextMatrix(iRow, 1)
'      iX = iX0 + ((iColWidth0 + iColWidth1) / 2 - Printer.TextWidth(strContent) / 2)
'      iY = iY0 + iRow * iRowHeight + iRowHeight / 2 - Printer.TextHeight(strContent) / 2
'      Printer.CurrentX = iX: Printer.CurrentY = iY
'      Printer.Print strContent
'
'      '固定行2
'      Printer.FontSize = 12
'      Printer.FontBold = False
'      iCol = 1
'      iXp = iX0 + iColWidth0
'      For iRow = 1 To 9
'         iYp = iY0 + iRow * iRowHeight
'         If iRow = 3 Or iRow = 7 Then
'            Printer.FontBold = True
'         End If
'         strContent = .TextMatrix(iRow, iCol)
'         iX = iXp + (iColWidth1 / 2 - Printer.TextWidth(strContent) / 2)
'         iY = iYp + (iRowHeight / 2 - Printer.TextHeight(strContent) / 2)
'         Printer.CurrentX = iX: Printer.CurrentY = iY
'         Printer.Print strContent
'         If iRow = 3 Or iRow = 7 Then
'            Printer.FontBold = False
'         End If
'      Next
'
'      '列0,10
'      Printer.FontSize = intFontSize 'Modify by Amy 2021/02/17 原:14
'      For iRow = 0 To 10 Step 10
'         iXp = iX0 + iColWidth0 + iColWidth1
'         iYp = iY0 + iRow * iRowHeight
'         For iCol = 2 To .Cols - 1
'            'ADD BY SONIA 2015/4/9 104/4起取消中區其他的目標
'            'Modified by Lydia 2017/10/12
'            'If iCol = 9 And Val(txtCloseDate) >= 1040401 Then
'            If iCol = iDeptS29 Then
'            Else
'            'END 2015/4/9
'               strContent = .TextMatrix(iRow, iCol)
'               'Modify by Amy 2021/02/17
'               bolMultLine = False
'               If Left(strContent, 2) = "專利" Or Left(strContent, 3) = "FCT" Then
'                    If Left(strContent, 2) = "專利" Then
'                        strContent = Left(strContent, 2)
'                    Else
'                        strContent = Left(strContent, 3)
'                    End If
'                    iY = iYp + 50
'                    bolMultLine = True
'               Else
'                    iY = iYp + (iRowHeight / 2 - Printer.TextHeight(strContent) / 2)
'               End If
'               'end 2021/02/17
'               '置中
'               iX = iXp + (iColWidth / 2 - Printer.TextWidth(strContent) / 2)
'               Printer.CurrentX = iX: Printer.CurrentY = iY
'               Printer.Print strContent
'               'Add by Amy 2021/02/17
'               If bolMultLine = True Then
'                    strContent = Replace(.TextMatrix(iRow, iCol), strContent, "")
'                    If InStr(strContent, "部") > 0 Then
'                        iX = iXp + (iColWidth / 2 - Printer.TextWidth(strContent) / 2)
'                    Else
'                        iX = iX - 200
'                    End If
'                    iY = iYp + Printer.TextHeight(strContent) + 50
'                    Printer.CurrentX = iX: Printer.CurrentY = iY
'                    Printer.Print strContent
'               End If
'               iXp = iXp + iColWidth
'            End If  'ADD BY SONIA 2015/4/9
'         Next
'      Next
'      '列1-9
'      Printer.FontSize = intFontSize2 'Moidy by Amy 2021/02/17 原:11
'      For iRow = 1 To 9
'         '粗體
'         If iRow = 3 Or iRow = 7 Then
'            Printer.FontBold = True
'         '一般
'         Else
'            Printer.FontBold = False
'         End If
'         iXp = iX0 + iColWidth0 + iColWidth1
'         iYp = iY0 + iRow * iRowHeight
'         For iCol = 2 To .Cols - 1
'            'ADD BY SONIA 2015/4/9 104/4起取消中區其他的目標
'            'Modified by Lydia 2017/10/12
'            'If iCol = 9 And Val(txtCloseDate) >= 1040401 Then
'            If iCol = iDeptS29 Then
'            Else
'            'END 2015/4/9
'               strContent = .TextMatrix(iRow, iCol)
'               iX = iXp + (iColWidth - Printer.TextWidth(strContent) - 50)
'               iY = iYp + (iRowHeight / 2 - Printer.TextHeight(strContent) / 2)
'               Printer.CurrentX = iX: Printer.CurrentY = iY
'               Printer.Print strContent
'               iXp = iXp + iColWidth
'            End If  'ADD BY SONIA 2015/4/9
'         Next
'      Next
'      'PS.
'      Printer.FontSize = 14
'      Printer.FontBold = True
'      strContent = lblPS
'      iX = iX0 + 300
'      iY = iY0 + 11 * iRowHeight + 150
'      Printer.CurrentX = iX: Printer.CurrentY = iY
'      Printer.Print strContent
'
'      '實收-應收
'      Printer.FontSize = 18
'      Printer.FontBold = True
'      'Modify By Sindy 2020/9/11
'      'strContent = "實收-應收：" & lblNet
'      strContent = "本月累計之收款-應收：" & lblNet
'      '2020/9/11 END
'      iX = 9000 '11000
'      iY = iY0 + 11 * iRowHeight + 150
'      Printer.CurrentX = iX: Printer.CurrentY = iY
'      Printer.Print strContent
'   End With
'
'   Printer.EndDoc
'   DoPrint = True
'
'ErrHnd:
'   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub cmdSearch_Click()
Dim m_Dept As String  'add by sonia 2015/4/23
Dim strTime As String
   
   strTime = ServerTime
   Screen.MousePointer = vbHourglass
   
   If ConstrainCheck = True Then
      Call DelR210107 'Add by Amy 2024/03/01 改寫入暫存檔
      Call doQuery
      
      'ADD BY SONIA 2015/4/23 提醒是否有未輸業績的智權部部門
      'Modify by Amy 2024/03/01 原S11~S99,改為S10,部門人員有目標或有點數,未輸才顯示,SalesPoint有資料以SalesPoint為主
      If UCase(strSPData) = UCase("NoData") Or Val(strDate) < Val(業績輸入啟用年月 & "00") Then
         strExc(2) = "Select Distinct st15 From Staff Where st04='1' And st15>='S10' And st15<='S99' And st01>'6' And st01<'F'"
         strExc(3) = "Select Distinct st15 From Staff,DailyFeat Where df02=" & txtCloseDate & " And df01=st01(+) And SubStr(st15,1,1)='S' "
      Else
         strExc(2) = "Select Distinct SP48 as ST15 From SalesPoint Where SP01=" & Left(TransDate(txtCloseDate, 2), 6) & " And SubStr(SP48,1,1)='S' "
         strExc(3) = "Select Distinct SP48 as ST15 From SalesPoint,DailyFeat Where df02=" & txtCloseDate & " And df01=SP02(+) And SubStr(SP48,1,1)='S' " & _
                              "And SP01=" & Left(TransDate(txtCloseDate, 2), 6)
      End If
      Label5 = "": m_Dept = ""
      strSql = "Select A.st15 From (" & strExc(2) & ") A, (" & strExc(3) & ") B " & _
                     "Where a.st15=b.st15(+) And b.st15 is null order by 1"
      CheckOC3
      With AdoRecordSet3
         .CursorLocation = adUseClient
         .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If .RecordCount > 0 Then
            Do While Not .EOF
               m_Dept = m_Dept & .Fields(0) & ","
               .MoveNext
            Loop
         End If
         If m_Dept <> "" Then
            Label5 = "尚有" & Left(m_Dept, Len(m_Dept) - 1) & "未輸入業績資料！"
         End If
      End With
      '2015/4/23 END
   End If
   Screen.MousePointer = vbDefault
   'MsgBox strTime & " ~ " & ServerTime
End Sub

Private Sub Form_Load()
   lblWorkDays.Caption = "": lblWorkDay.Caption = "": Label5.Caption = "" 'Add by Amy 2024/03/01
   MoveFormToCenter Me
   '預設前一工作日
   txtCloseDate = TransDate(CompWorkDay(2, strSrvDate(1), 1), 1)
   txtCloseDate_LostFocus
   'Add by Amy 2024/03/01 設定 列/欄 名
   Call SetGridTitleRow(True)
   strShowDept(0) = "": strShowDept(1) = ""
   Call DelR210107
   Call SetGridColData(0)
   'end 2024/03/01
   Call SetDataListWidth
   
   'Add by Amy 2024/03/01 顯示註記文字
   Label6.Caption = "其他=非上述欄位所列(ex:總經理、巨京...)" & vbCrLf & _
                                    "若「開發組」有資料併入「顧服組」"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm210107 = Nothing
End Sub

Private Function ConstrainCheck() As Boolean
   Dim bolCancel As Boolean 'Add by Amy 2018/11/28
    
   ConstrainCheck = False
   If txtCloseDate = "" Then
      MsgBox "請輸入點數結算日！", vbExclamation
      txtCloseDate.SetFocus
   '日期格式
   ElseIf ChkDate(txtCloseDate) = False Then
      txtCloseDate.SetFocus
      txtCloseDate_GotFocus
   '工作天
   ElseIf ChkWorkDay(TransDate(txtCloseDate, 2)) = False Then
      MsgBox "請輸入工作天！", vbExclamation
      txtCloseDate.SetFocus
      txtCloseDate_GotFocus
   Else
      'Add by Amy 2018/11/28
      txtCloseDate_Validate (bolCancel)
      If bolCancel = True Then
        txtCloseDate.SetFocus
        txtCloseDate_GotFocus
        Exit Function
      End If
      'end 2018/11/28
      ConstrainCheck = True
   End If
End Function

Private Sub txtCloseDate_Change()
   txtCloseDate.Tag = ""
End Sub

Private Sub txtCloseDate_GotFocus()
   TextInverse txtCloseDate
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtCloseDate.IMEMode = 2
   CloseIme
End Sub

Private Function GetWorkDay(p_stDate As String, ByRef p_stWorkday As String, ByRef p_stWorkdays As String) As Boolean
   
   Dim stDate As String, stYM As String

   stDate = TransDate(p_stDate, 2)
   stYM = Left(stDate, 6)
   
On Error GoTo ErrHnd
   
   strSql = " SELECT COUNT(*) TDAYS,SUM(DECODE(SIGN(WD01-" & stDate & "),1,0,1)) DAYS FROM WORKDAY WHERE WD01>" & stYM & "00 And WD01<" & stYM & "99"
   
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         p_stWorkdays = "" & .Fields(0)
         p_stWorkday = "" & .Fields(1)
         GetWorkDay = True
      Else
         MsgBox "無法讀取工作天資料！"
      End If
   End With
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

Private Sub txtCloseDate_LostFocus()
   If Val(txtCloseDate) < 940601 Then
      lblPS = "注意事項：財務資料未減去扣點數部分！"
   Else
      'Modify by Amy 2022/06/30 原:lblPS =""
      lblPS = "本月累計收文及未收款不包含ACS案！"
   End If
End Sub

'Add by Amy 2018/11/28
Private Sub txtCloseDate_Validate(Cancel As Boolean)
    '判斷不可輸下下個工作日判斷不可輸下下個工作日
    If Val(txtCloseDate) >= PUB_GetWorkDayAfterSysDate(CDbl(strSrvDate(1)), 2) Then
        txtCloseDate.Tag = ""
        MsgBox "只可輸至下一個工作日！"
        Cancel = True
    End If
End Sub

'Add by Amy 2021/02/23
Private Function ExcelSave() As Boolean
    Dim xlsAgentPoint As New Excel.Application, wksrpt As New Worksheet, xlsFileName As String
    Dim j As Integer, intStartR As Integer, intTitleR As Integer, stTp As String, stOldN As String
    
On Error GoTo ErrHand
    ExcelSave = False: bolErr = False
    xlsFileName = txtCloseDate & "業績達成日報表" & Format(Now, "yyyymmddhhmmss") & MsgText(43)
    If Dir(strExcelPath & xlsFileName) = MsgText(601) Then
        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
        End If
    Else
        Kill strExcelPath & xlsFileName
    End If
    
    xlsAgentPoint.Visible = True
    xlsAgentPoint.SheetsInNewWorkbook = 3 '工作表數量
    xlsAgentPoint.Workbooks.add
    Set wksrpt = xlsAgentPoint.Worksheets(1)
    intField = 65:  intCounter = 1
    
    Call SetColName
    Call PrintTitle(wksrpt)
    intTitleR = intCounter - 1
    
    '印資料
    intStartR = intCounter
    With grdDataList
        For i = 1 To .Rows - 1
            For j = 0 To .Cols - 1
                .row = i
                .col = j
                '非隱藏欄才顯示
                If .ColWidth(j) > 0 Then
                    strCol = GetFieldStr(j, intField)
                    stTp = .TextMatrix(i, j)
                    '目標/達成
                    If j = GetValue("") Then
                        If stOldN = "" Or (stTp <> stOldN And stTp <> MsgText(601)) Then
                            wksrpt.Range(strCol & intCounter).Value = Left(stTp, 1) & Chr(10) & Right(stTp, 1)
                            wksrpt.Range(strCol & intCounter).Font.Bold = True
                        End If
                        If stOldN <> "" And stTp <> stOldN Then
                            wksrpt.Range(strCol & intStartR & ":" & strCol & intCounter - 1).HorizontalAlignment = xlCenter
                            wksrpt.Range(strCol & intStartR & ":" & strCol & intCounter - 1).MergeCells = True
                            intStartR = intCounter
                        End If
                        stOldN = stTp
                    ElseIf j = GetValue("項目") Then
                        wksrpt.Range(strCol & intCounter).Value = stTp
                        If i = .Rows - 1 Then
                            wksrpt.Range(strCol & intCounter).Value = stTp
                            If stTp = "備註" Then
                                wksrpt.Range(strCol & intCounter).Font.Bold = True
                                wksrpt.Range(Chr(Asc(strCol) - 1) & intCounter & ":" & strCol & intCounter).HorizontalAlignment = xlCenter
                                wksrpt.Range(Chr(Asc(strCol) - 1) & intCounter & ":" & strCol & intCounter).MergeCells = True
                            Else
                                wksrpt.Range(strCol & intCounter).HorizontalAlignment = xlCenter
                            End If
                        End If
                    '其他欄位
                    Else
                        'Add by Amy 2023/06/02 專利國外部 及 專利日本部 不顯示 本月累計收文 及 本月ACS累計
                        If .TextMatrix(i, 1) = "本月累計收文" And (j = GetValue("專利國外部") Or j = GetValue("專利日本部")) Then
                           If j = GetValue("專利國外部") Then
                              stTp = strF410x(4)
                           ElseIf j = GetValue("專利日本部") Then
                              stTp = strF410x(5)
                           End If
                        ElseIf .TextMatrix(i, 1) = "本月ACS累計" And (j = GetValue("專利國外部") Or j = GetValue("專利日本部")) Then
                           If j = GetValue("專利國外部") Then
                              stTp = strF410xAcs(4)
                           ElseIf j = GetValue("專利日本部") Then
                              stTp = strF410xAcs(5)
                           End If
                        End If
                        'end 2023/06/01
                        wksrpt.Range(strCol & intCounter).Value = stTp
                        '數字靠右
                        If IsNumeric(stTp) = True Then
                            wksrpt.Range(strCol & intCounter).HorizontalAlignment = xlRight
                        Else
                            wksrpt.Range(strCol & intCounter).HorizontalAlignment = xlCenter
                        End If
                        'Add by Amy 2023/06/01
                        '專利國外部 及 專利日本部 不顯示(隱藏才不會橫向加總<>全所) 本月累計收文 及 本月ACS累計
                        If (.TextMatrix(i, 1) = "本月累計收文" And (j = GetValue("專利國外部") Or j = GetValue("專利日本部"))) _
                         Or (.TextMatrix(i, 1) = "本月ACS累計" And (j = GetValue("專利國外部") Or j = GetValue("專利日本部"))) Then
                           'Memo 值要先寫入再設定才有效
                           wksrpt.Range(strCol & intCounter).NumberFormatLocal = ";;;@" '隱藏
                        End If
                        '確認 「全所」是否與橫向加總相符
                        If j = GetValue("全所") And .TextMatrix(.row, 1) <> "備註" Then
                           'Modify by Amy 2024/03/01 原:加總至ＦＣＴ日文欄位,下不同年月會顯示不同欄,故改全所-1
                           stTp = "Sum(" & GetFieldStr(GetValue("智權部小計"), intField) & intCounter & ":" & GetFieldStr(GetValue("全所") - 1, intField) & intCounter & ")"
                           stTp = "IF(" & GetFieldStr(GetValue("全所"), intField) & intCounter & "<>" & stTp & ",""加總有誤"","""")"
                           strCol = Chr(Asc(strCol) + 1)
                           wksrpt.Range(strCol & intCounter).Value = "=" & stTp
                           If wksrpt.Range(strCol & intCounter).Value = "加總有誤" Then bolErr = True
                        End If
                        'end 2023/06/01
                    End If
                End If 'end 非隱藏欄才顯示
                'Add by Amy 2021/10/08 備註中依燈號顯示顏色
                If .TextMatrix(i, j) = "紅燈" Then
                    wksrpt.Range(Chr(j + intField) & intCounter).Interior.ColorIndex = 3
                ElseIf .TextMatrix(i, j) = "綠燈" Then
                    wksrpt.Range(Chr(j + intField) & intCounter).Interior.ColorIndex = 4
                End If
                'end 2021/10/08
            Next j
            If InStr(.TextMatrix(i, GetValue("項目")), "累計應收") > 0 Or InStr(.TextMatrix(i, GetValue("項目")), "累計收款") > 0 Then
                wksrpt.Range(Chr(GetValue("項目") + intField) & intCounter & ":" & GetFieldStr(UBound(arrField), intField) & intCounter).Font.Bold = True
                wksrpt.Range(Chr(GetValue("項目") + intField) & intCounter & ":" & GetFieldStr(UBound(arrField), intField) & intCounter).Interior.ColorIndex = 40 '膚色
                wksrpt.Range(Chr(GetValue("項目") + intField) & intCounter & ":" & GetFieldStr(UBound(arrField), intField) & intCounter).Interior.tintandshade = 0.2  '設深淺
            End If
            intCounter = intCounter + 1
        Next i
    End With
    
    '設內文文字大小
    wksrpt.Range(Chr(GetValue("項目") + 1 + intField) & intTitleR + 1 & ":" & GetFieldStr(UBound(arrField), intField) & intCounter - 1).Font.Size = 12
    wksrpt.Range(Chr(GetValue("項目") + 1 + intField) & intTitleR + 1 & ":" & GetFieldStr(UBound(arrField), intField) & intCounter - 1).RowHeight = 27.75
    '框線
    wksrpt.Range(Chr(intField) & intTitleR & ":" & GetFieldStr(UBound(arrField), intField) & intCounter - 1).Borders(xlEdgeTop).LineStyle = xlContinuous
    wksrpt.Range(Chr(intField) & intTitleR & ":" & GetFieldStr(UBound(arrField), intField) & intCounter - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
    wksrpt.Range(Chr(intField) & intTitleR & ":" & GetFieldStr(UBound(arrField), intField) & intCounter - 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
    wksrpt.Range(Chr(intField) & intTitleR & ":" & GetFieldStr(UBound(arrField), intField) & intCounter - 1).Borders(xlEdgeRight).LineStyle = xlContinuous
    wksrpt.Range(Chr(intField) & intTitleR & ":" & GetFieldStr(UBound(arrField), intField) & intCounter - 1).Borders(xlInsideVertical).LineStyle = xlContinuous
    wksrpt.Range(Chr(intField) & intTitleR & ":" & GetFieldStr(UBound(arrField), intField) & intCounter - 1).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    
    'Add by Amy 2021/03/05 +本月累計之收款-應收
    wksrpt.Range(Chr(intField) & intCounter).Value = Label4 & lblNet
    wksrpt.Range(Chr(intField) & intCounter).Font.Size = 14
    wksrpt.Range(Chr(intField) & intCounter).Font.Bold = True
    
    'Add by Amy 2022/06/30 +提醒文字
    wksrpt.Range("G" & intCounter).Value = "本月累計收文及未收款不包含ACS案"
    wksrpt.Range("G" & intCounter).Font.Size = 14
    wksrpt.Range("G" & intCounter).Font.Bold = True
    
    'Addedby Lydia 2021/09/06 小計縱向用顏色區隔(智權部小計)
    'Modify by Amy 2022/06/27 原:Pub_NumberToSystem26(iDeptSub + 1) & "5:" & Pub_NumberToSystem26(iDeptSub + 1) & "13"
    wksrpt.Range(Pub_NumberToSystem26(iDeptSub + 1) & intTitleR + 1 & ":" & Pub_NumberToSystem26(iDeptSub + 1) & (grdDataList.Rows + intTitleR) - 2).Interior.ColorIndex = 45
    'end 2021/09/06
    wksrpt.PageSetup.PaperSize = 9 '設定紙張 A4
    wksrpt.PageSetup.Orientation = xlLandscape '橫印
    wksrpt.PageSetup.LeftMargin = xlsAgentPoint.InchesToPoints(0.19)
    wksrpt.PageSetup.RightMargin = xlsAgentPoint.InchesToPoints(0.19)
    wksrpt.PageSetup.TopMargin = xlsAgentPoint.InchesToPoints(0.9)
    wksrpt.PageSetup.BottomMargin = xlsAgentPoint.InchesToPoints(0)
    wksrpt.PageSetup.HeaderMargin = xlsAgentPoint.InchesToPoints(0.38)
    wksrpt.PageSetup.FooterMargin = xlsAgentPoint.InchesToPoints(0)
    wksrpt.PageSetup.Zoom = 70
    
    If Val(xlsAgentPoint.Version) < 12 Then
       xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
    Else
       xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
    End If
    Set wksrpt = Nothing
    xlsAgentPoint.Workbooks.Close
    xlsAgentPoint.Quit
    Set xlsAgentPoint = Nothing
    ExcelSave = True
    Exit Function
        
ErrHand:
    MsgBox Err.Description, , MsgText(5)
    If Val(xlsAgentPoint.Version) < 12 Then
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
    Else
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
    End If
    Set wksrpt = Nothing
    xlsAgentPoint.Workbooks.Close
    xlsAgentPoint.Quit
    Set xlsAgentPoint = Nothing
End Function

Public Function GetFieldStr(ByVal intAdd As Integer, ByVal intField As Integer) As String
    Dim intDiv As Integer, intMod As Integer
    
    GetFieldStr = ""
    If intAdd + intField > 90 Then
        intDiv = intAdd \ 26 - 1
        intMod = intAdd Mod 26
        GetFieldStr = Chr(intDiv + intField) & Chr(intMod + intField)
    Else
        GetFieldStr = Chr(intAdd + intField)
    End If
End Function

'設定欄名
Private Sub SetColName()
    Dim stTmp1 As String, stTmp2 As String
    
    With grdDataList
        For i = 0 To grdDataList.Cols - 1
            .row = 0
            .col = i
            If .Width <> 0 Then
                stTmp1 = stTmp1 & "," & .TextMatrix(.row, .col)
                If .col = 0 Then
                    stTmp2 = stTmp2 & ",2.5"
                ElseIf .col = 1 Then
                    stTmp2 = stTmp2 & ",14.5"
                Else
                    stTmp2 = stTmp2 & ",9"
                End If
            End If
        Next i
    End With
    If stTmp1 <> MsgText(601) Then
        arrField = Split(Mid(stTmp1, 2), ",")
        arrWidth = Split(Mid(stTmp2, 2), ",")
    End If
End Sub

Private Sub PrintTitle(Wks As Worksheet)
    Dim stTmp As String, stTmp2 As String, intStartF As Integer, intStartChar As Integer, intLenChar As Integer
    
    strCol = GetFieldStr(UBound(arrField), intField)
    stTmp = CompNameQuery(2)
    Wks.Range(Chr(intField) & intCounter).Value = stTmp
    Wks.Range(Chr(intField) & intCounter).Font.Size = 26
    Wks.Range(Chr(intField) & intCounter & ":" & strCol & intCounter).HorizontalAlignment = xlCenter
    Wks.Range(Chr(intField) & intCounter & ":" & strCol & intCounter).MergeCells = True
    intCounter = intCounter + 1
    
    stTmp = Val(txtCloseDate) + 19110000
    stTmp = Val(Left(Val(stTmp), 4)) - 1911 & "年" & Val(Mid(Val(stTmp), 5, 2)) & "月業務目標及達成通知"
    Wks.Range(Chr(intField) & intCounter).Value = stTmp
    Wks.Range(Chr(intField) & intCounter).Font.Size = 18
    Wks.Range(Chr(intField) & intCounter & ":" & strCol & intCounter).HorizontalAlignment = xlCenter
    Wks.Range(Chr(intField) & intCounter & ":" & strCol & intCounter).MergeCells = True
    intCounter = intCounter + 1
    
    intStartF = LBound(arrField) + intField
    'Modify by Amy 2024/04/02 +統計日
    strCol = LBound(arrField) + 3
    stTmp = Replace(Label1, "　　　", lblWorkDays)
    Wks.Range(Chr(intStartF) & intCounter).Value = stTmp & "　統計日" & CFDate(txtCloseDate)
    'end 2024/04/02
    Wks.Range(Chr(intStartF) & intCounter).Font.Size = 12
    '日-天數字體放大且設粗體
     intStartChar = InStr(stTmp, lblWorkDays)
    intLenChar = Len(lblWorkDays)
    With Wks.Range(Chr(intStartF) & intCounter).Characters(intStartChar, intLenChar).Font
        .Size = 18
        .Bold = True
    End With
    Wks.Range(Chr(intStartF) & intCounter & ":" & GetFieldStr(Val(strCol), intField) & intCounter).HorizontalAlignment = xlLeft
    Wks.Range(Chr(intStartF) & intCounter & ":" & GetFieldStr(Val(strCol), intField) & intCounter).MergeCells = True
   
    intStartF = strCol + 1 + intField
    strCol = UBound(arrField) - 3
    stTmp = Replace(Label3, "　　　", lblWorkDay)
    Wks.Range(Chr(intStartF) & intCounter).Value = stTmp
    Wks.Range(Chr(intStartF) & intCounter).Font.Size = 12
    '日-天數字體放大且設粗體
    intStartChar = InStr(stTmp, lblWorkDay)
    intLenChar = Len(lblWorkDay)
    With Wks.Range(Chr(intStartF) & intCounter).Characters(intStartChar, intLenChar).Font
        .Size = 18
        .Bold = True
    End With
    Wks.Range(Chr(intStartF) & intCounter & ":" & GetFieldStr(Val(strCol), intField) & intCounter).HorizontalAlignment = xlCenter
    Wks.Range(Chr(intStartF) & intCounter & ":" & GetFieldStr(Val(strCol), intField) & intCounter).MergeCells = True
    
    '列印日期
    intStartF = strCol + 1 + intField
    strCol = UBound(arrField)
    'Modify by Amy 2024/04/02 顯示系統日
    'stTmp = Val(txtCloseDate) + 19110000
    stTmp = strSrvDate(1)
    stTmp = "列印日期：" & Val(Mid(stTmp, 1, 4)) - 1911 & "年" & Val(Mid(stTmp, 5, 2)) & "月" & Val(Mid(stTmp, 7, 2)) & "日"
    'end 2024/04/02
    Wks.Range(Chr(intStartF) & intCounter).Value = stTmp
    Wks.Range(Chr(intStartF) & intCounter).Font.Size = 12
    Wks.Range(Chr(intStartF) & intCounter & ":" & GetFieldStr(Val(strCol), intField) & intCounter).HorizontalAlignment = xlRight
    Wks.Range(Chr(intStartF) & intCounter & ":" & GetFieldStr(Val(strCol), intField) & intCounter).MergeCells = True
    intCounter = intCounter + 1
    
    For i = LBound(arrField) To UBound(arrField)
        stTmp2 = ""
        stTmp = arrField(i)
        If Len(stTmp) > 3 Then
            If Left(stTmp, 2) = "專利" Or stTmp = "中區其他" Then
                stTmp = Left(stTmp, 2)
            Else
                stTmp = Left(stTmp, 3)
            End If
            stTmp2 = Replace(arrField(i), stTmp, "")
            stTmp = stTmp & Chr(10) & stTmp2
        End If
        Wks.Range(GetFieldStr(i, intField) & intCounter).Value = stTmp
        Wks.Range(GetFieldStr(i, intField) & intCounter).Font.Size = 12
        Wks.Range(GetFieldStr(i, intField) & intCounter).HorizontalAlignment = xlCenter
        Wks.Columns(GetFieldStr(i, intField) & ":" & GetFieldStr(i, intField)).ColumnWidth = arrWidth(i)
    Next i
    'Modify by Amy 2024/03/01 原寫於for迴圈,會彈合併儲存格訊息,故改寫
    i = GetValue("項目")
    If GetFieldStr(i, intField) = Chr(Asc("A") + 1) Then
      Wks.Range(GetFieldStr(i - 1, intField) & intCounter).Value = Wks.Range(GetFieldStr(i, intField) & intCounter).Value
      Wks.Range(GetFieldStr(i, intField) & intCounter).Value = ""
      Wks.Range(GetFieldStr(i - 1, intField) & intCounter).Font.Bold = True
      Wks.Range(GetFieldStr(i - 1, intField) & intCounter & ":" & GetFieldStr(i, intField) & intCounter).MergeCells = True
    End If
    'end 2024/03/01
    intCounter = intCounter + 1
End Sub

Private Function GetValue(pFieldN As String) As Integer
   Dim jj As Integer
 
    For jj = LBound(arrField) To UBound(arrField)
       'Modify by Amy 2024/03/01 去空白
       If Replace(UCase(arrField(jj)), " ", "") = UCase(pFieldN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function

'Add by Amy 2024/03/01 欄位設定改寫至暫存檔
Private Sub SetDataListWidth()
   Dim RsQ As New ADODB.Recordset, strQ As String, intQ As Integer, ii As Integer
   
   If stColSQL = MsgText(601) Then
      '部門欄位順序 S->W->其他->P->F
      stColSQL = "Select DeptName,DeptNo,EmpNo,SeqNo From R210107 Where ID='" & strUserNum & "' And StateNo='0' " & _
                           "Order by SeqNo,DeptNo,EmpNo"
   End If
   strQ = "Select '        ' as DeptName,'A99' as DeptNo,'A9999' as EmpNo,'01' as SeqNo From Dual " & _
   "Union Select '項目' as DeptName,'AZZ' as DeptNo,'AZZZZ' as EmpNo,'02' as SeqNo From Dual " & _
   "Union " & stColSQL
   
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      grdDataList.Visible = False
      grdDataList.Cols = RsQ.RecordCount
      ReDim arrField(RsQ.RecordCount - 1)
      
      grdDataList.WordWrap = True
      grdDataList.Cols = RsQ.RecordCount
      grdDataList.row = 0
      RsQ.MoveFirst
      For ii = LBound(arrField) To UBound(arrField)
         arrField(ii) = "" & RsQ.Fields("DeptName")
         grdDataList.col = ii
         grdDataList.Text = arrField(ii)
         grdDataList.CellFontBold = True
         grdDataList.ColWidth(ii) = 665
         If Trim(arrField(ii)) = MsgText(601) Or Trim(RsQ.Fields("DeptName")) = "項目" Then
            If Trim(RsQ.Fields("DeptName")) = MsgText(601) Then
               grdDataList.ColWidth(ii) = 280
            ElseIf Trim(RsQ.Fields("DeptName")) = "項目" Then
               grdDataList.ColWidth(ii) = 1200
            End If
            grdDataList.CellAlignment = flexAlignLeftCenter '儲存格中資料的水平和垂直對齊
         Else
            If Trim(arrField(ii)) = "中區其他" Then
               grdDataList.ColWidth(ii) = 815
            ElseIf Replace(arrField(ii), " ", "") = "專利國外部" Or Replace(arrField(ii), " ", "") = "專利日本部" Then
               grdDataList.ColWidth(ii) = 720
            ElseIf Trim(arrField(ii)) = "智權部小計" Or Trim(arrField(ii)) = "全所" Then
               grdDataList.ColWidth(ii) = 760
            End If
            grdDataList.CellAlignment = flexAlignCenterCenter '儲存格中資料的水平和垂直對齊
            grdDataList.ColAlignment(ii) = flexAlignRightCenter
         End If
         grdDataList.ColAlignmentFixed(ii) = flexAlignCenterCenter '資料行中資料的對齊模式
         RsQ.MoveNext
      Next ii
      grdDataList.RowHeight(0) = 755
      grdDataList.Visible = True
   End If
   
End Sub

'intChoose:0-全部/1-不更新W1001部門/3-只更新W1001部門
Private Sub SetGridColData(intChoose As Integer)
   Dim strA As String, strDate As String, strVTB(2) As String, strTp(1) As String, strConSP As String
'*** Memo ***
'★下列變數有 增加 或 減少 都[要確認]之舊資料是否該 固定顯示
      '1121019 智權點數實績與結餘特殊員編=F4102;F4103;W1001;W2001;W3001;20091;F4104;F4105;F4106;F4107;P2005;P1005
      '                  智權點數實績與結餘輸入部門=S,F,W,P
'★1121018 記錄(與秀玲討論後修改)
      '畫面當月[巨京] 於9610~10806(不含)前用,之後併入[其他]->不論新舊資料改都列於[其他],不必單獨列示
      '畫面當月[智權部] [有]SalesPoint資料->抓有目標 或 有SalesPoint(20091-屬S29部門無目標也無點數,有就出現,加入智權部合計,因有陣子掛於[其他])
      '                               [無]SalesPoint資料->抓st15在職且員編大於6字頭且小於字頭之人員部門,避免Sxx員編(設定目標)也出現
      '                               畫面當月無資料,但有 未收款/已簽約 資料,於新增資料時再加入欄位 ex:1130217 當下S10部門並無目標及點數,但有未收款
      '畫面當月[非智權部]
      '[國外部(F4100)]：10801(不含)前用／[FCP(F4102)/FCT(F4103)]：10801~11001(不含)前用->固定出現
      '[F41XX名稱]：11001後用->抓SalesPoint資料
      '[客服組]：10806~11107(不含)前用／[顧服組]：10806後用->抓SalesPoint資料
      '[MCP(P1005)]：11004後用／'[MCT(P2005)]：11004後用->抓SalesPoint資料
'*** End Memo ***
   
   strDate = txtCloseDate
   strSPData = Pub_GetField("SalesPoint", " sp01=" & Left(strDate + 19110000, 6) & " And sp01<>201512", "sp01", True)
   '[有]目標 或[有]點數 要出現
   '畫面當月SalesPoint [沒]資料 或 業績輸入啟用年月-10501月前(不含) 抓st15
   If UCase(strSPData) = UCase("NoData") Or Val(strDate) < Val(業績輸入啟用年月 & "00") Then
      strShowDept(0) = Replace(";" & 智權點數實績與結餘特殊員編, ";20091", "")
      '[FCP及FCT] 11001月(含)後不顯示
      strShowDept(0) = Replace(Replace(strShowDept(0), ";F4102", ""), ";F4103", "")
      '[W10-客服組] 11107月(不含)後不顯示;若有W30-開發組資料(目前不會有目標) ,則併入W20-顧服組顯示
      strShowDept(0) = Mid(Replace(Replace(strShowDept(0), ";W1001", ""), ";W3001", ""), 2)
      strShowDept(1) = Mid("," & 智權點數實績與結餘輸入部門, 2)
      
      '[非 智權部] 業績啟用日-201601月後用
      If Val(strDate) >= Val(業績輸入啟用年月 & "00") Then
         strVTB(1) = GetPoint(3, strDate, strDate, , , strShowDept(0), False, Me.Name, True) & _
         " Union " & GetPoint(1.3, strDate, strDate, , , strShowDept(0), False, Me.Name, True)
      End If
      '[智權部] 員編為Sxx為目標設定,避免被抓出故需判斷st01<'F';[中區其他]固定顯示,故排除;
      strA = GetPoint(3, strDate, strDate, "S10", "S99", , False, Me.Name) & _
      " Union " & GetPoint(1.3, strDate, strDate, "S10", "S99", , False, Me.Name)
      If strVTB(1) <> MsgText(601) Then strA = strA & "Union " & strVTB(1)
      '有點數
   '畫面當月SalesPoint[已]有資料,抓sp48
   Else
      '中區其他(20091) SalesPoint不一定有資料->[固定列示-strVTB(0) ]
      strShowDept(0) = ""
      strShowDept(1) = ",S,F"
     '*** F部門 ***
      '201901(不含)前->固定顯示[國外部] (業績啟用日 後~每日業務點數FCPFCT啟用日)'Z'
      If Val(strDate) < Val(每日業務點數FCPFCT啟用日) Then
         strShowDept(0) = strShowDept(0) & ";F4100"
         strVTB(1) = "Select '" & strUserNum & "','0' as StateNo,st02 as DepN,st15 as SP48,st01 From Staff Where st01='F4100' "
      '201901(含)後~202101(不含)前->固定顯示[FCP及FCT] (每日業務點數FCPFCT啟用日~改為F4104~F4107)
      ElseIf Val(strDate) >= Val(每日業務點數FCPFCT啟用日) And Val(strDate) < 1100100 Then
         strShowDept(0) = strShowDept(0) & ";F4102;F4103"
         strVTB(1) = "Select '" & strUserNum & "','0' as StateNo,'FCP' as DepN,st15 as SP48,st01 From Staff Where st01='F4102' " & _
                 " Union Select '" & strUserNum & "','0' as StateNo,'FCT' as DepN,st15 as SP48,st01 From Staff Where st01='F4103' "
      '202101(含)後
      Else
         strShowDept(0) = strShowDept(0) & ";F4104;F4105;F4106;F4107"
      End If
     '*** W部門 ***
      If Val(strDate) >= 1080600 Then '201906月
         strShowDept(0) = strShowDept(0) & ";W1001;W2001"
         strShowDept(1) = strShowDept(1) & ",W"
         '202207(11107月不含)後不顯示W10-客服組
         If Val(strDate) >= 1110700 Then
            strShowDept(0) = Replace(strShowDept(0), ";W1001", "")
         End If
         'Memo 若有W30-開發組資料(目前不會有目標) ,則併入W20-顧服組顯示
      End If
     '*** P部門 ***
      If Val(strDate) >= 1100400 Then '202104月
         strShowDept(0) = strShowDept(0) & ";P1005;P2005"
         strShowDept(1) = strShowDept(1) & ",P"
      End If
      
      strShowDept(0) = Mid(strShowDept(0), 2)
      strShowDept(1) = Mid(strShowDept(1), 2)
      
      '財務當月關閉點數輸入時,會將當月[不需]輸入且有點數的人員寫入SalesPoint
      strConSP = strConSP & " And SP48<>'S29' "
      '智權部
      strA = GetPoint_SP(strDate, strDate, "S10", "S99", , False, Me.Name, , False, strConSP)
      '[非]智權部
      'F部門資料:107年(含)以前,顯示於「國外部」/ 108年(含)~109年(含) 顯示於「FCP/FCT」/ 110年(含)以後顯示於F4104~F4107 名稱
      strConSP = " And SP02 In('" & Replace(strShowDept(0), ";", "','") & "')"
      If Val(strDate) < 1100100 Then
         strConSP = strConSP & " And SubStr(SP02,1,3)<>'F41' "
         strA = strA & " Union " & strVTB(1)
      End If
      strVTB(2) = GetPoint_SP(strDate, strDate, "*NS", , , False, Me.Name, , True, strConSP)
      strA = strA & " Union " & strVTB(2)
   End If
   strA = Replace(strA, ",StateNo", ",'0' as StateNo")
   
   '中區其他(20091) SalesPoint不一定有資料->使用S29部門資料[固定列示]
   strVTB(0) = "Select '" & strUserNum & "','0' as StateNo,a0902 as DepN,a0901,a0901||'ZZ'  From Acc090 Where a0901='S29' "
   
   '欄位寫入暫存檔
   strA = "Insert Into R210107 (ID,StateNo,DeptName,DeptNo,EmpNo) " & strA & " Union " & strVTB(0)
   adoTaie.Execute strA
   '新增 智權小計/其他/全所
   strA = "Insert Into R210107 (ID,StateNo,DeptName,DeptNo,EmpNo) Values('" & strUserNum & "','0','智權部小計','SZZ','S9999') "
   adoTaie.Execute strA
   strA = "Insert Into R210107 (ID,StateNo,SeqNo,DeptName,DeptNo,EmpNo) Values('" & strUserNum & "','0','O','其他','OZZ','O9999') "
   adoTaie.Execute strA
   strA = "Insert Into R210107 (ID,StateNo,SeqNo,DeptName,DeptNo,EmpNo) Values('" & strUserNum & "','0','Z','全所','ZZZ','Z9999') "
   adoTaie.Execute strA
   
   '更新顯示
   '更新顯示順序
   Call UpdSeqNo(intChoose)
   '專利xxx 換行… ex:專利 國內部 MCP
   strA = "Update R210107 set DeptName=Replace(DeptName,'專利',' 專利 ') Where ID='" & strUserNum & "' And Instr(DeptName,'專利')>0 " 'EmpNo='P1005'
   adoTaie.Execute strA
   
End Sub

Private Sub SetGridTitleRow(Optional ByVal IsFirst As Boolean = False)
   Dim ii As Integer, stRowN As String, stRowWidth As String
   
   grdDataList.Visible = False
   If IsFirst = True Then
      '依顯示順序寫入                                                                                                       本日實收     本月實績
      stRowN = "項目,本月目標,每日應收,本月累計應收,本日入帳,本月累計入帳,本日收款,本月累計收款,本月累計收文,本月ACS累計,未收款,ACS未收款,備註"
      arrField2 = Split(stRowN, ",")
   End If
   For ii = LBound(arrField2) To UBound(arrField2)
      grdDataList.Rows = UBound(arrField2) + 1: grdDataList.FixedRows = 1
      '項目 第2欄
      grdDataList.col = 1
      grdDataList.row = ii
      grdDataList.ColWidth(ii) = 1200
      grdDataList.Text = arrField2(ii)
      If ii = GetColVal(arrField2, "本月累計應收") Or ii = GetColVal(arrField2, "本月累計收款") Or ii = GetColVal(arrField2, "備註") Then
         grdDataList.CellFontBold = True
      End If
   Next ii
   '項目 第1欄
   grdDataList.col = 0
   grdDataList.row = 0
   grdDataList.ColWidth(0) = 280
   grdDataList.Text = ""
   grdDataList.CellAlignment = flexAlignRightCenter
   grdDataList.CellFontBold = True
   grdDataList.ColAlignmentFixed(grdDataList.col) = flexAlignCenterCenter
   For ii = GetColVal(arrField2, "本月目標") To GetColVal(arrField2, "本月累計應收")
      grdDataList.row = ii
      grdDataList.Text = "目標"
      grdDataList.CellFontBold = True
   Next ii
   For ii = GetColVal(arrField2, "本日入帳") To GetColVal(arrField2, "ACS未收款")
      grdDataList.row = ii
      grdDataList.Text = "達成"
      grdDataList.CellFontBold = True
   Next ii
   grdDataList.MergeCells = flexMergeRestrictColumns
   grdDataList.MergeCol(0) = True '目標 達成
   grdDataList.Refresh
   grdDataList.Visible = True
End Sub

'從doQuery搬出來
Private Function InsertDailyFeat(ByRef stWorkDays, ByRef stWorkDay, ByRef stVTable1, ByRef stCon As String) As Boolean
   Dim stF4100Amt As Variant   'add by sonia 2014/6/11
   'Add by Amy 2018/11/28
   Dim i As Integer, intRun As Integer
   'modify by sonia 2021/1/19
   'Dim stF410X(1 To 2) As Variant, stVal As Variant
   Dim stF410X() As Variant, stVal As Variant
   Dim stF410XVal() As Variant, k As Integer
   'end 2021/1/19
   'Dim stSalesArea As String
   
   'Memo by Lydia 2015/07/29 程式若有改變,請一併變更frmAutoBatchDay.strMenu64
   InsertDailyFeat = False
   
      
On Error GoTo ErrHnd

'Modify by Amy 2024/03/01 +if SalesPoint當月 或 DailyFeat當日 [沒]資料才寫,[舊資料不重寫],因為目前舊資料「部門」改抓SalesPoint,重抓資料比較麻煩,故可不抓-秀玲
If UCase(strSPData) = UCase("NoData") Or Val(Pub_GetField("DailyFeat", "DF02=" & stDDate & " And DF01 in('F4104','F4105','F4106','F4107')", "Count(DF01)")) <> 4 Then
   'ADD BY SONIA 2014/6/11 **** 小真說江副所長指示：國外部：實收=財務入帳,為日報,月報統一故改跑日報時先更新每日業績點數輸入,資料自2014/6/1起   *****
   '但因仍有尾數差異,日報表之實收仍於SetExceptCell改為財務入帳數字
   cnnConnection.BeginTrans
      
   'Modify by Amy 2018/11/28 本日入帳原TO_CHAR(ROUND(SUM(NVL(V1C1,0))/1000,3),'9999999.000') 導致與月報不合需扣點數,1080101後國外部改區分FCP及FCT(此有修改frm210108 也要修改)
   'modify by sonia 2021/1/19 2021/1/1起F4102改F4104及F4105,F4103改F4106及F4107,再加V1C0做小計
   If Val(txtCloseDate) > Val(每日業務點數FCPFCT啟用日) Then
        strExc(0) = "SELECT V1C0,TO_CHAR(ROUND(NVL(SUM(NVL(V1C1,0)),0)/1000,2),'9999999.00') 本日入帳,1 Sort from STAFF," & _
                    "(" & stVTable1 & ") VT1 WHERE SUBSTR(ST15,1,1)='F' And V1C0(+)=ST01 And V1C0 in ('F4102','F4104','F4105') group by V1C0 " & _
                "Union SELECT V1C0,TO_CHAR(ROUND(NVL(SUM(NVL(V1C1,0)),0)/1000,2),'9999999.00') 本日入帳,2 Sort from STAFF," & _
                    "(" & stVTable1 & ") VT1 WHERE SUBSTR(ST15,1,1)='F' And V1C0(+)=ST01 And V1C0 in ('F4103','F4106','F4107') group by V1C0 "
        intRun = 2
        If Val(txtCloseDate) > 1100101 Then intRun = 4
        ReDim stF410X(1 To intRun) As Variant
        ReDim stF410XVal(1 To intRun) As Variant
   Else
        strExc(0) = "SELECT TO_CHAR(ROUND(NVL(SUM(NVL(V1C1,0)),0)/1000,2),'9999999.00') 本日入帳 from STAFF," & _
                    "(" & stVTable1 & ") VT1 WHERE SUBSTR(ST15,1,1)='F' And V1C0(+)=ST01 "
        intRun = 1
   End If
   intI = 1: stF4100Amt = 0
   'modify by sonia 2021/1/19 2021/1/1起F4102改F4104及F4105,F4103改F4106及F4107,故先刪除當日所有F410%資料
   'stF410X(1) = 0: stF410X(2) = 0
   'Add by Amy 2021/02/23 +if 輸每日業務點數FCPFCT啟用日前資料會錯
   If Val(txtCloseDate) > Val(每日業務點數FCPFCT啟用日) Then
    For i = 1 To intRun
       stF410X(i) = ""
       stF410XVal(i) = 0
    Next i
   End If
   strSql = "delete DailyFeat WHERE DF01 like 'F410%' And DF02=" & stDDate
   cnnConnection.Execute strSql
   'end 2021/1/19
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      RsTemp.MoveFirst
      i = 1  'add by sonia 2021/1/19
      Do While Not RsTemp.EOF
         If Val(txtCloseDate) > Val(每日業務點數FCPFCT啟用日) Then
            'modify by sonia 2021/1/19 2021/1/1起F4102改F4104及F4105,F4103改F4106及F4107
            'stF410X(Val(RsTemp.Fields("Sort"))) = Val("" & RsTemp(0))
            stF410X(i) = "" & RsTemp(0)
            stF410XVal(i) = Val("" & RsTemp(1))
            i = i + 1
            'end 2021/1/19
         Else
            stF4100Amt = Val("" & RsTemp(0))
         End If
         RsTemp.MoveNext
      Loop
   'add by sonia 2021/1/19  當日沒收款也要寫DailyFeat
   Else
      If intRun = 4 Then
         For i = 1 To intRun
            stF410X(i) = "F410" & i + 3
         Next i
      ElseIf intRun = 2 Then
         For i = 1 To intRun
            stF410X(i) = "F410" & i + 1
         Next i
      End If
   End If
   For i = 1 To intRun
      'stVal = stF4100Amt
      'If intRun > 1 Then stVal = stF410X(i)
      'strExc(0) = "SELECT * from DailyFeat WHERE DF01='F410" & IIf(Val(txtCloseDate) > Val(每日業務點數FCPFCT啟用日), i + 1, i - 1) & "' And DF02=" & stDDate
      'intI = 1
      'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      'If intI = 1 Then
      '   strSql = "UPDATE DailyFeat SET DF03=" & stVal & ",DF04=" & stVal & " WHERE DF01='F410" & IIf(Val(txtCloseDate) > Val(每日業務點數FCPFCT啟用日), i + 1, i - 1) & "' And DF02=" & stDDate
      'Else
      '   strSql = "insert into DailyFeat (df01,df02,df03,df04) values('F410" & IIf(Val(txtCloseDate) > Val(每日業務點數FCPFCT啟用日), i + 1, i - 1) & "'," & stDDate & "," & stVal & "," & stVal & ")"
      'End If
      'cnnConnection.Execute strSql
      If intRun = 1 And i = 1 Then
         strSql = "insert into DailyFeat (df01,df02,df03,df04,df06,df07,df08) values('F4100'," & stDDate & "," & stF4100Amt & "," & stF4100Amt & ",'QPGMR'," & strSrvDate(1) & "," & Mid(Right("000000" & ServerTime, 6), 1, 4) & ")"
         cnnConnection.Execute strSql
      Else
         If stF410X(i) <> "" Then
            strSql = "insert into DailyFeat (df01,df02,df03,df04,df06,df07,df08) values('" & stF410X(i) & "'," & stDDate & "," & stF410XVal(i) & "," & stF410XVal(i) & ",'QPGMR'," & strSrvDate(1) & "," & Mid(Right("000000" & ServerTime, 6), 1, 4) & ")"
            cnnConnection.Execute strSql
         End If
      End If
   Next i
   'end 2018/11/28
   cnnConnection.CommitTrans
End If
   InsertDailyFeat = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub Calculate()
   Dim ii As Integer, jj As Integer, stDept As String, stDeptName As String, sTemp As String, stVal As String, stMsg As String
   Dim stSumAll() As String
   'Memo 1121101 前修改請參閱Calculate_old
   ReDim stSumAll(GetColVal(arrField2, "本月目標") To GetColVal(arrField2, "備註") - 1)
   
   With grdDataList
      AdoRecordSet3.MoveFirst
      Do While Not AdoRecordSet3.EOF
         stDeptName = "" & AdoRecordSet3.Fields("項目")
         stDept = "" & AdoRecordSet3.Fields("DeptNo")
         sTemp = "" & AdoRecordSet3.Fields("EmpNo")
         ii = GetColVal(arrField, stDeptName)
         .col = ii '欄
         For jj = GetColVal(arrField2, "本月目標") To GetColVal(arrField2, "備註")  '列
            .row = jj
            '[全所]欄
            If stDeptName = "全所" And jj <> GetColVal(arrField2, "備註") Then
               .TextMatrix(jj, ii) = Format(Val("" & AdoRecordSet3.Fields(jj + 1)), "#.00")
               '驗證加總是否有誤
               If Val(.TextMatrix(jj, ii)) <> Val(stSumAll(jj)) Then
                  stMsg = stMsg & ",[" & arrField2(jj) & "]" & .TextMatrix(jj, ii) & "及" & Val(stSumAll(jj))
               End If
            '[備註]列
            ElseIf jj = GetColVal(arrField2, "備註") Then
               '有目標(不為0)都顯示登號
               If Val(.TextMatrix(GetColVal(arrField2, "本月目標"), ii)) = 0 Then
                  .Text = "": .CellBackColor = .BackColor
               Else
                  strExc(3) = Val(.TextMatrix(GetColVal(arrField2, "本月累計收款"), ii))
                  If Val(strExc(3)) >= Val(.TextMatrix(GetColVal(arrField2, "本月目標"), ii)) Then
                     .TextMatrix(jj, ii) = "達成":  .CellBackColor = vbGreen
                  ElseIf Val(strExc(3)) >= Val(.TextMatrix(GetColVal(arrField2, "本月累計應收"), ii)) Then
                     '綠燈：本月實收(本月累計收款)>=合計應收(本月累計應收)
                     .TextMatrix(jj, ii) = "綠燈": .CellBackColor = vbGreen
                  Else
                     '紅燈：簽收合計<合計應收
                     .TextMatrix(jj, ii) = "紅燈":  .CellBackColor = vbRed
                  End If
               End If
            '[非 備註]列
            Else
               stVal = Format(Val("" & AdoRecordSet3.Fields(jj + 1)), "#.00")
               '列加總
               If stDeptName <> "智權部小計" Then
                  stSumAll(jj) = Val(stSumAll(jj)) + Val(stVal)
               End If
                '專利國外部 及 專利日本部 不顯示「本月累計收文」及「本月ACS累計」,但合計仍需算
               If (sTemp = "F4104" Or "" & sTemp = "F4105") _
                  And (jj = GetColVal(arrField2, "本月累計收文") Or jj = GetColVal(arrField2, "本月ACS累計")) Then
                  stVal = ""
               ElseIf Left(stDept, 1) <> "S" And Left(stDept, 1) <> "W" _
                  And (jj = GetColVal(arrField2, "未收款") Or jj = GetColVal(arrField2, "ACS未收款")) Then
                  stVal = ""
               End If
               .TextMatrix(jj, ii) = stVal
            End If
            '** 顏色設定 **
            If jj = GetColVal(arrField2, "本月累計應收") Or jj = GetColVal(arrField2, "本月累計收款") Then
               .CellFontBold = True
            If stDeptName <> "全所" Then .CellBackColor = vbCyan '藍
            End If
            If stDeptName = "智權部小計" And jj <> GetColVal(arrField2, "備註") Then
               .CellBackColor = vbYellow '黃
            End If
            '** End 顏色設定 **
         Next jj
         AdoRecordSet3.MoveNext
      Loop
      If stMsg = MsgText(601) Then
         lblNet = Format(Val(.TextMatrix(GetColVal(arrField2, "本月累計收款"), GetColVal(arrField, "全所"))) - Val(.TextMatrix(GetColVal(arrField2, "本月累計應收"), GetColVal(arrField, "全所"))), "#.00")
      Else
         stMsg = Replace(Mid(stMsg, 2), ",", vbCrLf) & vbCrLf & _
                        "資料有誤 , 請洽電腦中心"
         MsgBox stMsg
      End If
   End With
End Sub

Private Sub DelR210107()
   Dim stDel As String
   
   '刪暫存檔-欄
   stDel = "Delete From R210107 Where ID='" & strUserNum & "' "
   adoTaie.Execute stDel
End Sub

'intChoose:0-全部/1-不更新W1001部門/2-只更新W1001部門
Private Sub UpdSeqNo(intChoose As Integer)
   Dim intSeq As Integer, strA As String, stWhere As String
   
   If intChoose = 0 Or intChoose = 2 Then
      '[客戶組]顯示於 智權部小計 之前,並加於[智權部小計=SZZ]
      strA = "Update R210107 set DeptNo='SW1' Where ID='" & strUserNum & "' And EmpNo='W1001' "
      adoTaie.Execute strA
   End If
   
   '更新顯示順序
   stWhere = " And ID='" & strUserNum & "' And StateNo='0' "
   intSeq = 1
   strA = "Update R210107 set SeqNo='" & intSeq & "' Where Substr(DeptNo,1,1)='S' " & stWhere
   adoTaie.Execute strA
   intSeq = intSeq + 1
   strA = "Update R210107 set SeqNo='" & intSeq & "' Where Substr(DeptNo,1,1)='W' And SeqNo is null " & stWhere
   adoTaie.Execute strA
   intSeq = intSeq + 1
   '其他
   strA = "Update R210107 set SeqNo='" & intSeq & "' Where Substr(DeptNo,1,1)='O' " & stWhere
   adoTaie.Execute strA
   intSeq = intSeq + 1
   strA = "Update R210107 set SeqNo='" & intSeq & "' Where Substr(DeptNo,1,1)='P' " & stWhere
   adoTaie.Execute strA
   intSeq = intSeq + 1
   '國外部依員編顯示
   strA = "Update R210107 set SeqNo='" & intSeq & "'||Substr(EmpNo,-1)  Where Substr(DeptNo,1,1)='F' " & stWhere
   adoTaie.Execute strA
End Sub






