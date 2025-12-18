VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210104 
   BorderStyle     =   1  '單線固定
   Caption         =   "點數輸入作業及查詢"
   ClientHeight    =   5720
   ClientLeft      =   50
   ClientTop       =   350
   ClientWidth     =   9430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5720
   ScaleWidth      =   9430
   Begin VB.OptionButton Option1 
      Caption         =   "扣點數"
      Height          =   315
      Index           =   4
      Left            =   7770
      TabIndex        =   30
      Top             =   510
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "業績點數統計"
      Height          =   400
      Left            =   6120
      Style           =   1  '圖片外觀
      TabIndex        =   28
      Top             =   30
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "每日點數輸入"
      Height          =   400
      Left            =   4740
      Style           =   1  '圖片外觀
      TabIndex        =   27
      Top             =   30
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "顯示每日累計未收款點數(查詢費時)"
      Height          =   255
      Left            =   5190
      TabIndex        =   26
      Top             =   1530
      Width           =   3645
   End
   Begin VB.OptionButton Option1 
      Caption         =   "銷帳及扣點數"
      Height          =   315
      Index           =   5
      Left            =   7920
      TabIndex        =   23
      Top             =   1260
      Width           =   1425
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1260
      MaxLength       =   3
      TabIndex        =   5
      Top             =   850
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1740
      MaxLength       =   6
      TabIndex        =   6
      Top             =   850
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   2580
      MaxLength       =   1
      TabIndex        =   7
      Top             =   850
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   2820
      MaxLength       =   2
      TabIndex        =   8
      Top             =   850
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Caption         =   "個人總表"
      Height          =   315
      Index           =   3
      Left            =   6750
      TabIndex        =   22
      Top             =   1260
      Width           =   1110
   End
   Begin VB.TextBox txtSalesArea1 
      Height          =   300
      Left            =   2475
      TabIndex        =   2
      Top             =   470
      Width           =   915
   End
   Begin VB.OptionButton Option1 
      Caption         =   "銷帳"
      Height          =   315
      Index           =   2
      Left            =   6990
      TabIndex        =   21
      Top             =   510
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7740
      Style           =   1  '圖片外觀
      TabIndex        =   24
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8580
      Style           =   1  '圖片外觀
      TabIndex        =   25
      Top             =   30
      Width           =   800
   End
   Begin VB.TextBox txtZone 
      Height          =   300
      Left            =   1260
      MaxLength       =   1
      TabIndex        =   0
      Top             =   90
      Width           =   915
   End
   Begin VB.TextBox txtSales 
      Height          =   300
      Left            =   4500
      MaxLength       =   6
      TabIndex        =   3
      Top             =   470
      Width           =   915
   End
   Begin VB.OptionButton Option1 
      Caption         =   "點數統計"
      Height          =   315
      Index           =   1
      Left            =   5580
      TabIndex        =   20
      Top             =   1260
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "收款及扣點數"
      Height          =   315
      Index           =   0
      Left            =   4080
      TabIndex        =   19
      Top             =   1260
      Value           =   -1  'True
      Width           =   1395
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   300
      Index           =   1
      Left            =   2520
      MaxLength       =   7
      TabIndex        =   10
      Top             =   1230
      Width           =   915
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   300
      Index           =   0
      Left            =   1260
      MaxLength       =   7
      TabIndex        =   9
      Top             =   1230
      Width           =   915
   End
   Begin VB.TextBox txtSalesArea 
      Height          =   300
      Left            =   1260
      TabIndex        =   1
      Top             =   470
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3915
      Left            =   45
      TabIndex        =   11
      Top             =   1800
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   6897
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   7
      ScrollTrack     =   -1  'True
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
      Left            =   4500
      TabIndex        =   4
      Top             =   468
      Width           =   1920
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3387;593"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblMemo 
      Caption         =   "P.S.個人總表和各區業務工作報告統計有差異在轉撥增減"
      ForeColor       =   &H00FF0000&
      Height          =   165
      Left            =   0
      TabIndex        =   31
      Top             =   1590
      Visible         =   0   'False
      Width           =   5175
   End
   Begin MSForms.Label lblSalesName 
      Height          =   300
      Left            =   5490
      TabIndex        =   29
      Top             =   510
      Width           =   1470
      VariousPropertyBits=   27
      Size            =   "2593;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      Caption         =   "本所案號"
      Height          =   180
      Left            =   225
      TabIndex        =   18
      Top             =   910
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   2205
      X2              =   2475
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "lblNote"
      Height          =   180
      Left            =   4005
      TabIndex        =   17
      Top             =   855
      Width           =   510
   End
   Begin VB.Label lblZone 
      AutoSize        =   -1  'True
      Caption         =   "（1北所 2中所 3南所 4高所）"
      Height          =   180
      Left            =   2295
      TabIndex        =   16
      Top             =   150
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "智權人員："
      Height          =   180
      Left            =   3465
      TabIndex        =   15
      Top             =   530
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   2205
      X2              =   2475
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Label Label3 
      Caption         =   "所別"
      Height          =   180
      Left            =   225
      TabIndex        =   14
      Top             =   150
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "點數結算日"
      Height          =   180
      Left            =   225
      TabIndex        =   13
      Top             =   1290
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "業務區"
      Height          =   180
      Left            =   225
      TabIndex        =   12
      Top             =   530
      Width           =   900
   End
End
Attribute VB_Name = "frm210104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/03 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、lblSalesName
'Memo by Lydia 2021/07/27 智權-調整財務系統(20200909) 'Memo by Lydia 2021/08/27 上線
'原「原業績點數查詢」標題修改為「點數輸入作業及查詢」。
'原畫面只有查詢Option選項(5選項)「收款及扣點數、點數統計、銷帳、個人總表、扣點數」
'修改後：新增「每日點數輸入」、「業績點數統計」按鈕，查詢Option選項(5選項，銷帳和扣點數合併)「收款及扣點數、點數統計、個人總表、銷帳扣點數」
'「每日點數輸入」按鈕：呼叫「每日業績點數輸入frm210103」，區主管作業選單先隱藏。
'「業績點數統計」按鈕：呼叫「業績點數統計frm210137」，區主管作業選單先隱藏。
'end 2021/07/27
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'2005/7/5整理
Option Explicit

Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
Dim m_strListPer As String 'Add By Sindy 2020/7/1
Dim stST05 As String, stST15 As String
Dim strA0b01_1 As String, strA0b05_1 As String, strA0b01_J As String, strA0b05_J As String 'Added by Lydia 2021/07/27 會計過帳日/業績輸入關閉年月
'Add By Sindy 2023/6/12
Dim arrID
Dim bolAreaMan As Boolean '下拉選單有區主管
'2023/6/12 END


'p_iKind:統計方式 0=明細 1=每日統計
'p_bolHeaderOnly:是否只設定表頭 true=是 false=資料一併清除
Private Sub SetDataListWidth(Optional ByVal p_iKind As Integer = 0, Optional ByVal p_bolHeaderOnly As Boolean = False)
   Dim intP As Integer
   
   Check1.Visible = False 'Add By Sindy 2020/10/5
   With grdDataList
      .Visible = False
      
      'Added by Lydia 2021/07/27
      LblNote = ""
      If p_iKind = 1 Then '點數統計
          'Memo by Lydia 2021/08/10 改成和個人總表一樣的算法
          If Val(Mid(txtCloseDate(1), 1, Len(txtCloseDate(1)) - 2)) <= Val(strA0b05_1) And Val(Mid(txtCloseDate(1), 1, Len(txtCloseDate(1)) - 2)) <= Val(strA0b05_J) Then '輸入日期為已決算的月份
               LblNote = "注意事項：總計 = 累計 + 實績動用 + 結餘動用"
          Else
               LblNote = "注意事項：總計 = 累計 + 現有實績 + 現有結餘"
          End If
      ElseIf p_iKind = 3 Then  '個人總表
          If Val(Mid(txtCloseDate(1), 1, Len(txtCloseDate(1)) - 2)) <= Val(strA0b05_1) And Val(Mid(txtCloseDate(1), 1, Len(txtCloseDate(1)) - 2)) <= Val(strA0b05_J) Then '輸入日期為已決算的月份
             LblNote = "注意事項：點數合計 = 財務累計 + 實績動用 + 結餘動用"
          Else
             LblNote = "注意事項：點數合計 = 已入帳點數 + 現有實績 + 現有結餘"
          End If
      ElseIf p_iKind = 5 Then  '銷帳及扣點數
          LblNote = "注意事項：總計 = 銷帳點數 + 扣點數"
      End If
      'end 2021/07/27
      '明細 0:收款及扣點數、4:扣點數
      If p_iKind = 0 Or p_iKind = 4 Then 'Modify by Amy 2013/08/07 +扣點數欄位與收款及扣點數同
         If p_bolHeaderOnly = False Then
            .Clear
            .Rows = 3: .Cols = 7: .FixedRows = 2: .FixedCols = 0
            .MergeRow(0) = True
            .MergeRow(1) = True
            .MergeCol(0) = True
            .MergeCells = flexMergeFree
         End If
         .row = 0
         .col = 0: .ColWidth(.col) = 900: .Text = "智權人員"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         'Add by Morgan 2005/6/16
         .col = 1: .ColWidth(.col) = 1100: .Text = "傳票日期"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         '2005/6/16 end
         .col = 2: .ColWidth(.col) = 1200: .Text = "傳票號碼"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 3: .ColWidth(.col) = 1450: .Text = "本所案號"
         .ColAlignment(.col) = flexAlignLeftCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 4: .ColWidth(.col) = 3200: .Text = "摘要內容"
         .ColAlignment(.col) = flexAlignLeftCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 5: .ColWidth(.col) = 1000: .Text = "點數"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         For intP = 6 To .Cols - 1
            .ColWidth(intP) = 0
         Next
         .row = 1
         .col = 4: .Text = "總計："
         .CellAlignment = flexAlignRightCenter
         .col = 5: .Text = ""
         .CellAlignment = flexAlignRightCenter
         For intP = 0 To .Cols - 1
            .col = intP
            .CellBackColor = &H90EE90
         Next
         
      '1:點數統計
      ElseIf p_iKind = 1 Then
         Check1.Visible = True 'Add By Sindy 2020/10/5 顯示每日累計未收款點數
         If p_bolHeaderOnly = False Then
            .Clear
            .Rows = 3: .Cols = 12: .FixedRows = 2: .FixedCols = 0
            .MergeRow(0) = True
            .MergeCol(0) = True
            .MergeCells = flexMergeFree
         End If
         .row = 0
         .col = 0: .ColWidth(.col) = 680: .Text = ""
         .col = 1: .ColWidth(.col) = 800: .Text = ""
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 2: .ColWidth(.col) = 880: .Text = ""
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         'Modified by Lydia 2021/07/27 「財務結帳」更名為「已入帳點數」
         .col = 3: .ColWidth(.col) = 700: .Text = "已入帳點數"
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 4: .ColWidth(.col) = 880: .Text = "已入帳點數"
         'end 2021/07/27
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         'Modified by Lydia 2021/07/27 「智權人員收款」更名為「已收款點數」
         .col = 5: .ColWidth(.col) = 700: .Text = "已收款點數"
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 6: .ColWidth(.col) = 880: .Text = "已收款點數"
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 7: .ColWidth(.col) = 880: .Text = "已收款點數"
         'end 2021/07/27
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 8: .ColWidth(.col) = 1100: .Text = "" 'Modify By Sindy 2020/8/31 "智權人員簽約"
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 9: .ColWidth(.col) = 1100: .Text = "" 'Modify By Sindy 2020/8/31 "智權人員簽約"
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 10: .ColWidth(.col) = 700: .Text = "" 'Modify By Sindy 2020/8/31 "智權人員簽約"
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .ColWidth(11) = 0
         .ColWidth(12) = 0
         .row = 1
         .col = 0: .Text = "智權人員"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 1: .Text = "結算日"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 2: .Text = "點數差"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 3: .Text = "當日"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 4: .Text = "累計"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 5: .Text = "當日"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 6: .Text = "扣點數"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 7: .Text = "累計"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 8: .Text = "已收文點數" 'Modify By Sindy 2020/8/31 "當日"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 9: .Text = "未收款點數" 'Modify By Sindy 2020/8/31 "銷帳"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         'Modify By Sindy 2020/8/31 Mark
         .col = 10: .Text = ""
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      
      '2:銷帳
      ElseIf p_iKind = 2 Then
         If p_bolHeaderOnly = False Then
            .Clear
            .Rows = 3: .Cols = 6: .FixedRows = 2: .FixedCols = 0
            .MergeRow(0) = True
            .MergeRow(1) = True
            .MergeCol(0) = True
            .MergeCells = flexMergeFree
         End If
         .row = 0
         .col = 0: .ColWidth(.col) = 900: .Text = "智權人員"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 1: .ColWidth(.col) = 1200: .Text = "收據編號"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 2: .ColWidth(.col) = 1450: .Text = "本所案號"
         .ColAlignment(.col) = flexAlignLeftCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 3: .ColWidth(.col) = 4100: .Text = "摘要內容"
         .ColAlignment(.col) = flexAlignLeftCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 4: .ColWidth(.col) = 1200: .Text = "點數"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .ColWidth(5) = 0
         .row = 1
         .col = 3: .Text = "總計："
         .CellAlignment = flexAlignRightCenter
         .col = 4: .Text = ""
         .CellAlignment = flexAlignRightCenter
         For intP = 0 To .Cols - 1
            .col = intP
            .CellBackColor = &H90EE90
         Next
      'Added by Lydia 2021/07/27 銷帳及扣點數合併=5, 原本銷帳選項2和扣點數選項4隱藏保留
      ElseIf p_iKind = 5 Then
         If p_bolHeaderOnly = False Then
            .Clear
            .Rows = 3: .Cols = 10: .FixedRows = 2: .FixedCols = 0
            .MergeRow(0) = True
            .MergeRow(1) = True
            .MergeCol(0) = True
            .MergeCells = flexMergeFree
         End If
         .row = 0
         .col = 0: .ColWidth(.col) = 900: .Text = "智權人員"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 1: .ColWidth(.col) = 1350: .Text = "本所案號"
         .ColAlignment(.col) = flexAlignLeftCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 2: .ColWidth(.col) = 1100: .Text = "收據號碼"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 3: .ColWidth(.col) = 1100: .Text = "傳票日期"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 4: .ColWidth(.col) = 1200: .Text = "傳票號碼"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 5: .ColWidth(.col) = 1800: .Text = "摘要內容"
         .ColAlignment(.col) = flexAlignLeftCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 6: .ColWidth(.col) = 950: .Text = "銷帳點數"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 7: .ColWidth(.col) = 950: .Text = "扣點數"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         For intP = 8 To .Cols - 1
            .ColWidth(intP) = 0
         Next
         .row = 1
         .col = 5: .Text = "總計："
         .CellAlignment = flexAlignRightCenter
         .col = 6: .Text = ""  '銷帳總計
         .CellAlignment = flexAlignRightCenter
         .col = 7: .Text = ""  '扣點數總計
         .CellAlignment = flexAlignRightCenter
         For intP = 0 To .Cols - 1
            .col = intP
            .CellBackColor = &H90EE90
         Next
      'end 2021/07/27
      '3:個人總表
      Else
         If p_bolHeaderOnly = False Then
            .Clear
            'Modified by Lydia 2021/07/27 加「收文點數、目標點數、未收款點數」.Cols = 6 =>.Cols = 9
            .Rows = 2: .Cols = 9: .FixedRows = 1: .FixedCols = 0
            .MergeCells = flexMergeNever
         End If
         .row = 0
         .col = 0: .ColWidth(.col) = 1600: .Text = "智權人員"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 1: .ColWidth(.col) = 1450: .Text = "點數合計"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         'Modified by Lydia 2021/07/27 變更欄位
'         .col = 2: .ColWidth(.col) = 1450: .Text = "財務累計"
'         .ColAlignment(.col) = flexAlignRightCenter
'         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
'         .col = 3: .ColWidth(.col) = 1450: .Text = "動用保留"
'         .ColAlignment(.col) = flexAlignRightCenter
'         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
'         .col = 4: .ColWidth(.col) = 1450: .Text = "動用結餘"
'         .ColAlignment(.col) = flexAlignRightCenter
'         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
'         .col = 5: .ColWidth(.col) = 1400: .Text = "扣點數"
'         .ColAlignment(.col) = flexAlignRightCenter
'         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
'         For intP = 6 To .Cols - 1
'            .ColWidth(intP) = 0
'         Next
         intP = 1
         '輸入日期為已決算的月份
         If Val(Mid(txtCloseDate(1), 1, Len(txtCloseDate(1)) - 2)) <= Val(strA0b05_1) And Val(Mid(txtCloseDate(1), 1, Len(txtCloseDate(1)) - 2)) <= Val(strA0b05_J) Then
            intP = intP + 1
            .col = intP: .ColWidth(.col) = 1300: .Text = "財務累計"
            .ColAlignment(.col) = flexAlignRightCenter
            .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
            intP = intP + 1
            .col = intP: .ColWidth(.col) = 1300: .Text = "實績動用"
            .ColAlignment(.col) = flexAlignRightCenter
            .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
            intP = intP + 1
            .col = intP: .ColWidth(.col) = 1300: .Text = "結餘動用"
            .ColAlignment(.col) = flexAlignRightCenter
            .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         '輸入日期為當月的日期
         Else
            intP = intP + 1
            .col = intP: .ColWidth(.col) = 1300: .Text = "已入帳點數"
            .ColAlignment(.col) = flexAlignRightCenter
            .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
            intP = intP + 1
            .col = intP: .ColWidth(.col) = 1300: .Text = "現有實績"
            .ColAlignment(.col) = flexAlignRightCenter
            .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
            intP = intP + 1
            .col = intP: .ColWidth(.col) = 1300: .Text = "現有結餘"
            .ColAlignment(.col) = flexAlignRightCenter
            .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         End If
         intP = intP + 1
         .col = intP: .ColWidth(.col) = 1400: .Text = "扣點數"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         intP = intP + 1
         .col = intP: .ColWidth(.col) = 1300: .Text = "已收文點數"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         intP = intP + 1
         .col = intP: .ColWidth(.col) = 1300: .Text = "目標點數"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         intP = intP + 1
         .col = intP: .ColWidth(.col) = 1300: .Text = "未收款點數"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         For intP = 9 To .Cols - 1
               .ColWidth(intP) = 0
         Next
      End If
      .Refresh
      .Visible = True
   End With
End Sub

Private Function ConstrainCheck() As Boolean
Dim bolCancel As Boolean
Dim intErrCol As Integer

   ConstrainCheck = True
   
   'Modify by Amy 2013/08/07 +本所案號不為空值則點數結算日可不填
   If (Text1 <> "" And Text2 = "") Then
      MsgBox "本所案號輸入錯誤，請重新輸入 ！", vbExclamation
      Text2.SetFocus
      Text2_GotFocus
      ConstrainCheck = False
      Exit Function
   End If
   
   If (Text1 = "" And Text2 <> "") Then
      MsgBox "系統類別錯誤，請重新輸入！", vbExclamation
      Text1.SetFocus
      Text1_GotFocus
      ConstrainCheck = False
      Exit Function
   End If
   
   Call Text1_Validate(bolCancel)
   If bolCancel = True Then
      Text1.SetFocus
      Text1_GotFocus
      ConstrainCheck = False
      Exit Function
   End If
   
   Call Text2_Validate(bolCancel)
   If bolCancel = True Then
      Text2.SetFocus
      Text2_GotFocus
      ConstrainCheck = False
      Exit Function
   End If
   
   If txtCloseDate(0) = "" And Trim(Text1) = "" Then
      MsgBox "請輸入點數結算起日！", vbExclamation
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
   If txtCloseDate(1) = "" And Trim(Text1) = "" Then
      MsgBox "請輸入點數結算迄日！", vbExclamation
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
   '2015/4/30 add by sonia
   If Val(txtCloseDate(1)) < Val(txtCloseDate(0)) Then
      MsgBox "點數結算起迄日範圍錯誤！", vbExclamation
      txtCloseDate(0).SetFocus
      txtCloseDate_GotFocus (0)
      ConstrainCheck = False
      Exit Function
   End If
   'end 2015/4/29
   'end 2013/08/07
   
   'Added by Lydia 2024/01/31 另外限制不可查詢
   If strUserNum = "75007" Then
      If DBDATE(txtCloseDate(0)) >= "20240201" Then
         MsgBox "不可查詢113年2月後的資料！", vbExclamation
         txtCloseDate(0).SetFocus
         txtCloseDate_GotFocus (0)
         ConstrainCheck = False
         Exit Function
      End If
      If DBDATE(txtCloseDate(1)) >= "20240201" Then
         MsgBox "不可查詢113年2月後的資料！", vbExclamation
         txtCloseDate(1).SetFocus
         txtCloseDate_GotFocus (1)
         ConstrainCheck = False
         Exit Function
      End If
   End If
   'end 2024/01/31
   
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
'   'Added by Lydia 2020/06/15 簡協理可看所有智權人員
'   If strUserNum = "69005" Then
'      If Left(txtSalesArea, 1) <> "S" Then
'         MsgBox "業務區起始條件錯誤！只可查智權部", vbExclamation
'         txtSalesArea.SetFocus
'         txtSalesArea_GotFocus
'         ConstrainCheck = False 'Added by Lydia 2020/07/14
'         Exit Function
'      End If
'      If Left(txtSalesArea1, 1) <> "S" Then
'         MsgBox "業務區迄止條件錯誤！只可查智權部", vbExclamation
'         txtSalesArea1.SetFocus
'         txtSalesArea1_GotFocus
'         ConstrainCheck = False 'Added by Lydia 2020/07/14
'         Exit Function
'      End If
'   End If
'   'end 2020/06/15
'
'   If txtSalesArea > txtSalesArea1 Then
'      MsgBox "業務區範圍條件錯誤！", vbExclamation
'      txtSalesArea.SetFocus
'      txtSalesArea_GotFocus
'      ConstrainCheck = False
'      Exit Function
'   End If
   
End Function

Private Function doQuery(ByVal p_iMode As Integer) As Boolean
   Dim stCon As String, stConST As String, stConDF As String, stConWD As String
   Dim stConDF1 As String
   Dim stVTBw As String, stVTBx As String, stVTBy As String, stVTBz As String
   Dim stCon2 As String
   Dim strTime As String

   strTime = ServerTime
   
   stCon = "": stConST = "": stConDF = "": stConDF1 = ""
   '所別
   If txtZone <> "" Then
      stConST = stConST & " and st06 = '" & txtZone & "'"
   End If
   
   '區別
   'Modify By Sindy 2020/6/15 Mark,input有區別就要加入sql
''Modify by Morgan 2005/6/24 改起迄
'   'Modify By Sindy 2009/05/12
'   '若為帶人主管權限時,查詢之智權人員編號非本人時,不限制區別
'   If (Trim(txtSales) <> "" And PUB_GetST05(strUserNum) = "SA" And txtSales.Enabled = True And txtSales <> strUserNum) Then
'      '不限制區別
'   Else
      If txtSalesArea <> "" Then
         stConST = stConST & " and st15>='" & txtSalesArea & "'"
      End If
      If txtSalesArea1 <> "" Then
         stConST = stConST & " and st15<='" & txtSalesArea1 & "'"
      End If
'   End If
''2005/6/24 end
   
   '智權人員
   If txtSales <> "" Then
      'Modify by Morgan 2008/4/25 若區主管查自己資料時要含該區已離職及虛建編號
      'stCon = stCon & " and ax209 = '" & txtSales & "'"
      'stConDF = stConDF & " and df01='" & txtSales & "'"
      'stConST = stConST & " and st01='" & txtSales & "'"
      stCon = stCon & " and ax209 in (" & PUB_GetSalesList(txtSales, txtSalesArea, txtSalesArea1) & ")"
      stConDF = stConDF & " and df01 in (" & PUB_GetSalesList(txtSales, txtSalesArea, txtSalesArea1) & ")"
      'modify by sonia 2017/1/26 開放吳碧梧70005可操作林文雄A4023
      'stConST = stConST & " and st01 in (" & PUB_GetSalesList(txtSales, txtSalesArea, txtSalesArea1) & ")"
      If PUB_GetSalesList(txtSales, txtSalesArea, txtSalesArea1) <> "" Then
         stConST = stConST & " and st01 in (" & PUB_GetSalesList(txtSales, txtSalesArea, txtSalesArea1) & ")"
      Else
         stConST = stConST & " and st01 in (" & PUB_GetSalesList(txtSales, txtSalesArea, txtSalesArea1) & ")"
      End If
      'end 2017/1/26
   End If
   
   '點數結算日
   If txtCloseDate(0) <> "" Then
      stConDF1 = stConDF & " and df02<" & txtCloseDate(0)
      stCon = stCon & " and a0205 >= " & txtCloseDate(0)
      stConDF = stConDF & " and df02 >=" & txtCloseDate(0)
      stConWD = stConWD & " and wd01 >=19110000+" & txtCloseDate(0)
   End If
   If txtCloseDate(1) <> "" Then
      stCon = stCon & " and a0205 <= " & txtCloseDate(1)
      stConDF = stConDF & " and df02 <=" & txtCloseDate(1)
      stConWD = stConWD & " and wd01 <=19110000+" & txtCloseDate(1)
   End If
   
   'Memo by Morgan 2013/8/14 本條件須放在最後面
   'Add by Amy 2013/08/07 +本所案號查詢
   If Text1 <> "" Then
      If Text3 = "" Then Text3 = "0"
      If Text4 = "" Then Text4 = "00"
      
      'Modified by Morgan 2013/8/14 若查詢扣點數則需同時抓相同案(同收款扣點數檢查)
      If Option1(4).Value = True Then
         strExc(1) = PUB_GetSameCaseSQL2(Text1, Text2, Text3, Text4)
         stCon2 = stCon & " and ax214 in (" & strExc(1) & ") and ax214<>'" & Text1 & Text2 & Text3 & Text4 & "'"
      End If
      'end 2013/8/14
      
      stCon = stCon & " and ax214='" & Text1 & Text2 & Text3 & Text4 & "' "
   End If
   'end 2013/08/07

On Error GoTo ErrHnd

   'Add by Amy 2013/08/07 +本所案號查詢
   If Trim(Text1) <> "" Then
     strExc(0) = " and ax201 = a0201(+)  and ax202 = a0202(+) "
   Else
     strExc(0) = " and ax201(+) = a0201  and ax202(+) = a0202 "
   End If
   
   '明細 'Memo by Lydia 2021/07/27 收款及扣點數
   If p_iMode = 0 Then
      LblNote = ""
      'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
      'Moidfy by Amy 2020/03/31 +7129
      strSql = "select st02,A0205,ax202, ax214, ax212, to_char(round((ax207-ax206)/1000,3),'999999.000') Point, ax209" & _
         " from acc020, acc021, staff where st01(+)=ax209 " & strExc(0) & _
         " and ax209 Is Not Null and (substr(ax205, 1, 1) = '4' Or ax205 = '7121' Or ax205 = '7129') " & stCon & stConST
         
      '排除保留&結餘
      'StrSql = StrSql & " and ax207>0 and not ( ax205='4191' or ax205='4192' and st04='1'" & _
         " or ((ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131') and instr(ax213||' ','結餘')>0))"
         
      strSql = strSql & " order by ax209,A0205,ax202,ax203 asc"
   
   '統計 'Memo by Lydia 2021/07/27 點數統計
   ElseIf p_iMode = 1 Then
      'LblNote = "注意事項：總計 = 累計 + 用保留 + 用結餘" 'Remove by Lydia 2021/08/10
      'Add By Sindy 2020/9/15 取得收文點數的語法stVTBw
      Call PUB_CountCP18(0, txtCloseDate(0).Text, txtCloseDate(1).Text, txtSalesArea, txtSalesArea1, txtSales, stVTBw, True)
      
      '銷帳資料 94/6以後才要
      'x:94/6/1~起日前一天的加總
      stVTBx = "SELECT A0K20 VxC0,SUM(A1U07) VxC1" & _
         " From ACC0S0, ACC0K0, ACC1U0 WHERE A0S04='1' AND A0S03>=940601 AND A0S03 < " & txtCloseDate(0) & _
         " AND A0K01(+)=A0S02 AND A0K20 IS NOT NULL AND A1U01(+)=A0S01 GROUP BY A0K20"
      'y:統計區間的銷帳資料
      stVTBy = "SELECT A0K20 VyC0,A0S03 VyC1,SUM(A1U07) VyC2" & _
         " From ACC0S0, ACC0K0, ACC1U0 WHERE A0S04='1' AND A0S03>=940601 AND A0S03>=" & txtCloseDate(0) & " AND A0S03 <= " & txtCloseDate(1) & _
         " AND A0K01(+)=A0S02 AND A0K20 IS NOT NULL AND A1U01(+)=A0S01 GROUP BY A0K20,A0S03"
      
      'z:統計區間的扣點數資料
      'Modify by Morgan 2005/9/2 扣點數條件改與點數一致,只差抓借方>0
      'modfiy by sonia 2014/6/30 結餘不限制科目改為41xx  D103062357的416102
      'stVTBz = "select ax209 VzC0,a0205 VzC1,sum(ax206) VzC2" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null and ax206>0 and (substr(ax205, 1, 2) = '41' Or ax205 = '7121')" & _
         " and not (  (ax205='4191' or ax205='4192')" & _
         " or ( (ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131') and instr(ax213||' ','結餘')>0 ) )" & _
         " group by ax209,a0205"
      'modify by sonia 2015/4/23 加不含'4194'科目,不含結餘傳票不要加科目限制
      'stVTBz = "select ax209 VzC0,a0205 VzC1,sum(ax206) VzC2" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null and ax206>0 and (substr(ax205, 1, 2) = '41' Or ax205 = '7121')" & _
         " and not (  (ax205='4191' or ax205='4192')" & _
         " or (substr(ax205, 1, 2) = '41' and instr(ax213||' ','結餘')>0 ) )" & _
         " group by ax209,a0205"
      'Modify by Amy 2019/08/01 增加創新業務組用收入 420101  原:substr(ax205, 1, 2) = '41'
      'Modify by Amy 2019/03/31 +7129
      stVTBz = "select ax209 VzC0,a0205 VzC1,sum(ax206) VzC2" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null and ax206>0 and (substr(ax205, 1, 1) = '4' Or ax205 = '7121' Or ax205 = '7129')" & _
         " and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0)" & _
         " group by ax209,a0205"
      
      '點數：貸方 ax207>0,科目substr(ax205, 1, 2) = '41' Or ax205 = '7121',排除保留&結餘
      '用保留：貸-借ax207-ax206, 科目ax205='4191' or ax205='4192'
      '用結餘：貸-借ax207-ax206, 科目ax205='4194' or 對沖-其他有'結餘'
      '扣點數：借方 ax206>0,科目substr(ax205, 1, 2) = '41' Or ax205 = '7121',排除保留&結餘
      '簽約：要減銷帳 94/6以後才要
      'Modify by Morgan 2005/6/9 財務及收款點數都要扣銷帳(與點數輸入一致)
      '2005/7/6 MODIFY BY SONIA 原控制 ST01<'99999'但有P1001及P2001故改為 ST01<'P2999'
      'Modify by Morgan 2005/9/4 簽約補扣銷帳點數
      'modfiy by sonia 2014/6/30 結餘不限制科目改為41xx  D103062357的416102
      'strSql = "Select st02, to_char(wd01) wd01,null D0" & _
         ", to_char(ROUND(nvl(X1,0)/1000,3)-round(nvl(VzC2,0)/1000,3),'999999.000') X1, null SX1" & _
         ", to_char(ROUND(nvl(Y1,0),2),'999999.000') Y1, to_char(round(nvl(VzC2,0)/1000,3),'999999.000') Ym1, null SY1" & _
         ", to_char(ROUND(nvl(Y2,0),2),'999999.000') Y2, to_char(round(nvl(VyC2,0)/1000,3),'999999.000') Ym2, to_char(nvl(Z1,0)-round(nvl(VxC1,0)/1000,3),'999999.000') SY2" & _
         ", st01,1 D1 from (select st01,st02,wd01-19110000 wd01" & _
         " from staff,workday where (st04='1' or st51>" & DBDATE(txtCloseDate(0)) & ") and st01>'6' and st01<'P2999'" & stConST & stConWD & _
         " ) SW, ( select ax209 X0,a0205, sum(ax207) X1" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null and (substr(ax205, 1, 2) = '41' Or ax205 = '7121')" & _
         " and ax207>0 and not( ax205='4191' or ax205='4192'" & _
         " or ((ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131') and instr(ax213||' ','結餘')>0))" & _
         " group by ax209,a0205) X,(select df01 Y0,df02,df03 Y1,df04 Y2" & _
         " From DailyFeat Where 1=1 " & stConDF & " ) Y" & _
         ",(select df01 Z0,sum(nvl(df04,0)-nvl(df03,0)) Z1" & _
         " From DailyFeat Where 1=1 " & stConDF1 & " group by df01) Z,(" & stVTBx & ") Vx,(" & stVTBy & ") Vy,(" & stVTBz & ") Vz"
      'modify by sonia 2015/4/23 加不含'4194'科目,不含結餘傳票不要加科目限制
      'strSql = "Select st02, to_char(wd01) wd01,null D0" & _
         ", to_char(ROUND(nvl(X1,0)/1000,3)-round(nvl(VzC2,0)/1000,3),'999999.000') X1, null SX1" & _
         ", to_char(ROUND(nvl(Y1,0),2),'999999.000') Y1, to_char(round(nvl(VzC2,0)/1000,3),'999999.000') Ym1, null SY1" & _
         ", to_char(ROUND(nvl(Y2,0),2),'999999.000') Y2, to_char(round(nvl(VyC2,0)/1000,3),'999999.000') Ym2, to_char(nvl(Z1,0)-round(nvl(VxC1,0)/1000,3),'999999.000') SY2" & _
         ", st01,1 D1 from (select st01,st02,wd01-19110000 wd01" & _
         " from staff,workday where (st04='1' or st51>" & DBDATE(txtCloseDate(0)) & ") and st01>'6' and st01<'P2999'" & stConST & stConWD & _
         " ) SW, ( select ax209 X0,a0205, sum(ax207) X1" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null and (substr(ax205, 1, 2) = '41' Or ax205 = '7121')" & _
         " and ax207>0 and not( ax205='4191' or ax205='4192'" & _
         " or (substr(ax205, 1, 2) = '41' and instr(ax213||' ','結餘')>0))" & _
         " group by ax209,a0205) X,(select df01 Y0,df02,df03 Y1,df04 Y2" & _
         " From DailyFeat Where 1=1 " & stConDF & " ) Y" & _
         ",(select df01 Z0,sum(nvl(df04,0)-nvl(df03,0)) Z1" & _
         " From DailyFeat Where 1=1 " & stConDF1 & " group by df01) Z,(" & stVTBx & ") Vx,(" & stVTBy & ") Vy,(" & stVTBz & ") Vz"
      'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
      'Modify by Amy 2020/03/31 +7129
      'Modify by Sindy 2020/9/1 Mark:", to_char(ROUND(nvl(Y2,0),2),'999999.000') Y2, to_char(round(nvl(VyC2,0)/1000,3),'999999.000') Ym2, to_char(nvl(Z1,0)-round(nvl(VxC1,0)/1000,3),'999999.000') SY2" => 0
      strSql = "Select st02, to_char(wd01) wd01,null D0" & _
         ", to_char(ROUND(nvl(X1,0)/1000,3)-round(nvl(VzC2,0)/1000,3),'999999.000') X1, null SX1" & _
         ", to_char(ROUND(nvl(Y1,0),2),'999999.000') Y1, to_char(round(nvl(VzC2,0)/1000,3),'999999.000') Ym1, null SY1" & _
         ", to_char(ROUND(nvl(收文點數,0),2),'999999.000') Y2, to_char(0,'999999.000') Ym2, ' - ' SY2" & _
         ", st01,1 D1 from (select st01,st02,wd01-19110000 wd01" & _
         " from staff,workday where (st04='1' or st51>" & DBDATE(txtCloseDate(0)) & ") and st01>'6' and st01<'P2999'" & stConST & stConWD & _
         " ) SW, ( select ax209 X0,a0205, sum(ax207) X1" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null and (substr(ax205, 1,1) = '4' Or ax205 = '7121' Or ax205 = '7129')" & _
         " and ax207>0 and not( ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0)" & _
         " group by ax209,a0205) X,(select df01 Y0,df02,df03 Y1,df04 Y2" & _
         " From DailyFeat Where 1=1 " & stConDF & " ) Y" & _
         ",(select df01 Z0,sum(nvl(df04,0)-nvl(df03,0)) Z1" & _
         " From DailyFeat Where 1=1 " & stConDF1 & " group by df01) Z,(" & stVTBx & ") Vx,(" & stVTBy & ") Vy,(" & stVTBz & ") Vz,(" & stVTBw & ") Vw"
      'Modify by Morgan 2005/5/3 改帶所有工作天
      'strSQL = strSQL & " where  X0(+)=st01 and Y0(+)=st01 and df02(+)=wd01 and a0205(+)=wd01 and (X1>0 or Y1>0 or Y2>0) and Z0(+)=st01"
      strSql = strSql & " where X0(+)=st01 and Y0(+)=st01 and df02(+)=wd01 and a0205(+)=wd01  and Z0(+)=st01" & _
         " AND VxC0(+)=st01 and VyC0(+)=st01 and VyC1(+)=wd01 and VzC0(+)=st01 and VzC1(+)=wd01 and Vw.CP13(+)=st01 and Vw.CP05(+)=wd01"
      
      'Modify by Morgan 2007/4/30  原控制 ST01<'99999'但有F4101~3 & P1001及P2001故不再限制
      'Modify by Morgan 2005/9/2 扣點數條件改與點數一致,只差抓借方>0
      'Modify by Morgan 2009/8/28 改抓在職或離職日大於結算起日的員工
      'modfiy by sonia 2014/6/30 結餘不限制科目改為41xx  D103062357的416102
      'strSql = strSql & " union all" & _
         " select st02" & _
         ", '總計:',null,'累計:',null" & _
         ", '結餘:', to_char(ROUND(nvl(Y1,0)/1000,3),'999999.000')" & _
         ", '保留:', to_char(ROUND(nvl(X1,0)/1000,3),'999999.000')" & _
         ", '扣點數:', to_char(ROUND(nvl(Z1,0)/1000,3),'999999.000')" & _
         ", st01,0 D1 from staff,( select ax209 X0,sum(ax207-ax206) X1" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null and (ax205='4191' or ax205='4192')" & _
         " group by ax209 ) X,( select ax209 Y0,sum(ax207-ax206) Y1" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null  and (ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131')" & _
         " and instr(ax213||' ','結餘')>0 group by ax209 ) Y,( select ax209 Z0,sum(-1*ax206) Z1" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null and ax206>0 and (substr(ax205, 1, 2) = '41' Or ax205 = '7121')" & _
         " and not ( (ax205='4191' or ax205='4192')" & _
         " or ( (ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131') and instr(ax213||' ','結餘')>0 ) )" & _
         " group by ax209 ) Z where X0(+)=st01 and Y0(+)=st01 and Z0(+)=st01 and (st04='1' or st51>" & DBDATE(txtCloseDate(0)) & ") and st01>'6' and st01<'P2999'" & stConST
      '2015/2/4 modify by sonia 結餘欄要包含4194結餘保留(但ax213不會輸'結餘'),動用保留欄不含4194結餘保留
      'strSql = strSql & " union all" & _
         " select st02" & _
         ", '總計:',null,'累計:',null" & _
         ", '結餘:', to_char(ROUND(nvl(Y1,0)/1000,3),'999999.000')" & _
         ", '保留:', to_char(ROUND(nvl(X1,0)/1000,3),'999999.000')" & _
         ", '扣點數:', to_char(ROUND(nvl(Z1,0)/1000,3),'999999.000')" & _
         ", st01,0 D1 from staff,( select ax209 X0,sum(ax207-ax206) X1" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null and (ax205='4191' or ax205='4192' or ax205='4194')" & _
         " group by ax209 ) X,( select ax209 Y0,sum(ax207-ax206) Y1" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null  and substr(ax205, 1, 2) = '41'" & _
         " and instr(ax213||' ','結餘')>0 group by ax209 ) Y,( select ax209 Z0,sum(-1*ax206) Z1" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null and ax206>0 and (substr(ax205, 1, 2) = '41' Or ax205 = '7121')" & _
         " and not ( (ax205='4191' or ax205='4192' or ax205='4194')" & _
         " or (substr(ax205, 1, 2) = '41' and instr(ax213||' ','結餘')>0 ) )" & _
         " group by ax209 ) Z where X0(+)=st01 and Y0(+)=st01 and Z0(+)=st01 and (st04='1' or st51>" & DBDATE(txtCloseDate(0)) & ") and st01>'6' and st01<'P2999'" & stConST
      'Modify by Amy 2019/08/01 增加創新業務組用收入 420101  原:substr(ax205, 1, 2) = '41'
      'Modify by Amy 2020/03/31 +7129
      'Modified by Lydia 2021/08/10 改成和個人總表一樣的算法
      'strSql = strSql & " union all" & _
         " select st02" & _
         ", '總計:',null,'累計:',null" & _
         ", '用保留:', to_char(ROUND(nvl(X1,0)/1000,3),'999999.000')" & _
         ", '用結餘:', to_char(ROUND(nvl(Y1,0)/1000,3),'999999.000')" & _
         ", '扣點數:', to_char(ROUND(nvl(Z1,0)/1000,3),'999999.000')" & _
         ", st01,0 D1 from staff,( select ax209 X0,sum(ax207-ax206) X1" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null and (ax205='4191' or ax205='4192')" & _
         " group by ax209 ) X,( select ax209 Y0,sum(ax207-ax206) Y1" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null  and substr(ax205, 1, 1) = '4'" & _
         " and (instr(ax213||' ','結餘')>0 or ax205='4194') group by ax209 ) Y,( select ax209 Z0,sum(-1*ax206) Z1" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null and ax206>0 and (substr(ax205, 1, 1) = '4' Or ax205 = '7121' Or ax205 = '7129')" & _
         " and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0)" & _
         " group by ax209 ) Z where X0(+)=st01 and Y0(+)=st01 and Z0(+)=st01 and (st04='1' or st51>" & DBDATE(txtCloseDate(0)) & ") and st01>'6' and st01<'P2999'" & stConST
      If Val(Mid(txtCloseDate(1), 1, Len(txtCloseDate(1)) - 2)) <= Val(strA0b05_1) And Val(Mid(txtCloseDate(1), 1, Len(txtCloseDate(1)) - 2)) <= Val(strA0b05_J) Then '輸入日期為已決算的月份
          LblNote = "注意事項：總計 = 累計 + 實績動用 + 結餘動用"
          '取得實績點數統計語法
          'Modify by Amy 2022/05/17 改傳畫面上原始日期
          'ex:下智權 87052 或 83004 點選「點數統計」日期1110401-0430 會有實績點數 與1110101-0430 實績點數沒累計,因只抓1月
          'strExc(2) = UCase(GetPoint(1, Val(Mid(txtCloseDate(0), 1, Len(txtCloseDate(0)) - 2)), Val(Mid(txtCloseDate(1), 1, Len(txtCloseDate(1)) - 2)), strExc(5), strExc(6), strExc(7)))
          strExc(2) = UCase(GetPoint(1, txtCloseDate(0), txtCloseDate(1), strExc(5), strExc(6), strExc(7), , Me.Name))
          '若個人尚未輸入期末實績C5則期末實績=期初實績C1
          '若個人尚未輸入期末結餘C6則期末結餘=期初結餘C2+當月結餘C4
          strExc(2) = "SELECT ST01 AS C0, SUM(PE04) AS PE04, SUM(C1) AS C1, SUM(C2) AS C2, SUM(C3) AS C3, SUM(C4) AS C4, " & _
                           "SUM(DECODE(C5,NULL,C1,C5)) AS C5, SUM(DECODE(C6,NULL,C2+C4,C6)) AS C6 FROM (" & strExc(2) & ") GROUP BY ST01"
          strSql = strSql & " union all" & _
             " select st02" & _
             ", '總計:',null,'累計:',null" & _
             ", '實績動用:',to_char(C1-C5,'999999.000')" & _
             ", '結餘動用:',to_char(C2+C4-C6,'999999.000') " & _
             ", '扣點數:', to_char(ROUND(nvl(Z1,0)/1000,3),'999999.000')" & _
             ", st01,0 D1 from staff,(" & strExc(2) & " ) C,( select ax209 Z0,sum(-1*ax206) Z1" & _
             " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
             " and ax209 is not null and ax206>0 and (substr(ax205, 1, 1) = '4' Or ax205 = '7121' Or ax205 = '7129')" & _
             " and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0)" & _
             " group by ax209 ) Z where C0(+)=st01 and Z0(+)=st01 and (st04='1' or st51>" & DBDATE(txtCloseDate(0)) & ") and st01>'6' and st01<'P2999'" & stConST
      Else   '輸入日期為當月的日期
          LblNote = "注意事項：總計 = 累計 + 現有實績 + 現有結餘"
          '取得實績點數統計語法
          strExc(2) = UCase(GetPoint(1, Val(Mid(txtCloseDate(0), 1, Len(txtCloseDate(0)) - 2)), Val(Mid(txtCloseDate(1), 1, Len(txtCloseDate(1)) - 2)), strExc(5), strExc(6), strExc(7)))
          strExc(2) = "SELECT ST01 AS C0, SUM(PE04) AS PE04, SUM(C1) AS C1, SUM(C2) AS C2, SUM(C3) AS C3, SUM(C4) AS C4, " & _
                           "SUM(DECODE(C5,NULL,C1,C5)) AS C5, SUM(DECODE(C6,NULL,C2+C4,C6)) AS C6 FROM (" & strExc(2) & ") GROUP BY ST01"
          'Modified by Lydia 2021/09/16 當月期間的現有結餘改回舊抓法; 因為每月2,13日有可能輸入結餘,
          'strSql = strSql & " union all" & _
             " select st02" & _
             ", '總計:',null,'累計:',null" & _
             ", '現有實績:',to_char(C1,'999999.000')" & _
             ", '現有結餘:',to_char(C2,'999999.000')" & _
             ", '扣點數:', to_char(ROUND(nvl(Z1,0)/1000,3),'999999.000')" & _
             ", st01,0 D1 from staff,(" & strExc(2) & " ) C,( select ax209 Z0,sum(-1*ax206) Z1" & _
             " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
             " and ax209 is not null and ax206>0 and (substr(ax205, 1, 1) = '4' Or ax205 = '7121' Or ax205 = '7129')" & _
             " and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0)" & _
             " group by ax209 ) Z where C0(+)=st01 and Z0(+)=st01 and (st04='1' or st51>" & DBDATE(txtCloseDate(0)) & ") and st01>'6' and st01<'P2999'" & stConST
          strSql = strSql & " union all" & _
             " select st02" & _
             ", '總計:',null,'累計:',null" & _
             ", '現有實績:',to_char(C1,'999999.000')" & _
             ", '現有結餘:', to_char(ROUND(nvl(Y1,0)/1000,3),'999999.000')" & _
             ", '扣點數:', to_char(ROUND(nvl(Z1,0)/1000,3),'999999.000')" & _
             ", st01,0 D1 from staff,(" & strExc(2) & " ) C, ( select ax209 Y0,sum(ax207-ax206) Y1" & _
             " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
             " and ax209 is not null  and substr(ax205, 1, 1) = '4'" & _
             " and (instr(ax213||' ','結餘')>0 or ax205='4194') group by ax209 ) Y, ( select ax209 Z0,sum(-1*ax206) Z1" & _
             " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
             " and ax209 is not null and ax206>0 and (substr(ax205, 1, 1) = '4' Or ax205 = '7121' Or ax205 = '7129')" & _
             " and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0)" & _
             " group by ax209 ) Z where C0(+)=st01 and Z0(+)=st01 and Y0(+)=ST01 and (st04='1' or st51>" & DBDATE(txtCloseDate(0)) & ") and st01>'6' and st01<'P2999'" & stConST
      End If
      'end 2021/08/10
      strSql = strSql & " order by 12 asc,2 desc,3 asc"
      
   'Memo by Lydia 2021/07/27 銷帳：銷帳及扣點數合併=5, 原本銷帳選項2和扣點數選項4隱藏保留
   ElseIf p_iMode = 2 Then
      LblNote = ""
      '銷帳資料
      'Add by Amy 2013/08/07 +本所案號查詢
      If Trim(Text1) <> "" Then
        'Modified by Morgan 2023/1/7 銷帳服務費可能為負(扣點數) A1U07>0->A1U07<>0
        strSql = "SELECT st02,A0K01,A0J02,A0K04||'/'||getcp10desc(cp01,cp10,a0j04) C1,to_char(ROUND(nvl(A1U07,0)/1000,3),'999999.000') C2,A0K20" & _
             " From ACC0S0,ACC0K0,staff,ACC1U0,ACC0J0,caseprogress WHERE A0S04='1' AND A0S03>=940601" & _
             " AND A0J02='" & Text1 & Text2 & Text3 & Text4 & "' " & _
             " AND A0K01(+)=A0S02 AND ST01(+)=A0K20" & stConST & _
             " AND A1U01=A0S01(+) AND A0J01=A1U03(+) AND A1U07<>0 and cp09(+)=a0j01"
        strSql = strSql & " order by A0K20,A1U01,A1U02,A1U03"
      Else
        'Modified by Morgan 2011/12/27 取消 a0j20
        'Modified by Morgan 2023/1/7 銷帳服務費可能為負(扣點數) A1U07>0->A1U07<>0
        strSql = "SELECT st02,A0K01,A0J02,A0K04||'/'||getcp10desc(cp01,cp10,a0j04) C1,to_char(ROUND(nvl(A1U07,0)/1000,3),'999999.000') C2,A0K20" & _
             " From ACC0S0,ACC0K0,staff,ACC1U0,ACC0J0,caseprogress WHERE A0S04='1' AND A0S03>=940601" & _
             " AND A0S03>=" & txtCloseDate(0) & " AND A0S03 <= " & txtCloseDate(1) & _
             " AND A0K01(+)=A0S02 AND ST01(+)=A0K20" & stConST & _
             " AND A1U01(+)=A0S01 AND A0J01(+)=A1U03 AND A1U07<>0 and cp09(+)=a0j01"
        strSql = strSql & " order by A0K20,A1U01,A1U02,A1U03"
      End If
      'end 2013/08/07
      
   'Add by Morgan 2010/8/27
   '個人點數總表
   ElseIf p_iMode = 3 Then
      'Modified by Lydia 2021/07/27 變更欄位
      'lblNote = "注意事項：點數合計 = 財務累計 + 動用保留 + 動用結餘"
      strExc(5) = "": strExc(6) = "": strExc(7) = ""
      '因為GetPoint模組內有語法分成只以部門或智權人員來查詢，所以視情況傳條件變數
      '另外暫時不考慮智權人員改部門的歷史資料查詢(7/22)
      If txtSales <> "" Then
          strExc(7) = txtSales
      Else
          If txtSalesArea <> "" Then strExc(5) = txtSalesArea
          If txtSalesArea1 <> "" Then strExc(6) = txtSalesArea1
      End If
      If Val(Mid(txtCloseDate(1), 1, Len(txtCloseDate(1)) - 2)) <= Val(strA0b05_1) And Val(Mid(txtCloseDate(1), 1, Len(txtCloseDate(1)) - 2)) <= Val(strA0b05_J) Then '輸入日期為已決算的月份
          LblNote = "注意事項：點數合計 = 財務累計 + 實績動用 + 結餘動用"
          '取得實績點數統計語法
          strExc(2) = UCase(GetPoint(1, Val(Mid(txtCloseDate(0), 1, Len(txtCloseDate(0)) - 2)), Val(Mid(txtCloseDate(1), 1, Len(txtCloseDate(1)) - 2)), strExc(5), strExc(6), strExc(7)))
          '若個人尚未輸入期末實績C5則期末實績=期初實績C1
          '若個人尚未輸入期末結餘C6則期末結餘=期初結餘C2+當月結餘C4
          strExc(2) = "SELECT ST01 AS C0, SUM(PE04) AS PE04, SUM(C1) AS C1, SUM(C2) AS C2, SUM(C3) AS C3, SUM(C4) AS C4, " & _
                           "SUM(DECODE(C5,NULL,C1,C5)) AS C5, SUM(DECODE(C6,NULL,C2+C4,C6)) AS C6 FROM (" & strExc(2) & ") GROUP BY ST01"
      Else   '輸入日期為當月的日期
          LblNote = "注意事項：點數合計 = 已入帳點數 + 現有實績 + 現有結餘"
          '取得實績點數統計語法
          'strExc(2) = GetPoint_SP(Val(strA0b05_1), Val(strA0b05_1), strExc(5), strExc(6), strExc(7))
          'strExc(2) = "SELECT ST01 AS C0,NVL(PE04,0) PE04,NVL(SP15,0) C1,NVL(SP36,0) C2 " & _
                           "FROM STAFF,PERFORMANCE, (" & strExc(2) & ") WHERE PE01(+)=ST01 AND PE02(+)='TOT' " & stConST & _
                           "AND PE03(+)=" & Left(TransDate(txtCloseDate(0), 2), 6) & " AND SP02(+)=ST01 "
          strExc(2) = UCase(GetPoint(1, Val(Mid(txtCloseDate(0), 1, Len(txtCloseDate(0)) - 2)), Val(Mid(txtCloseDate(1), 1, Len(txtCloseDate(1)) - 2)), strExc(5), strExc(6), strExc(7)))
          strExc(2) = "SELECT ST01 AS C0, SUM(PE04) AS PE04, SUM(C1) AS C1, SUM(C2) AS C2, SUM(C3) AS C3, SUM(C4) AS C4, " & _
                           "SUM(DECODE(C5,NULL,C1,C5)) AS C5, SUM(DECODE(C6,NULL,C2+C4,C6)) AS C6 FROM (" & strExc(2) & ") GROUP BY ST01"
      End If
      '收文點數：結算日期間
      Call PUB_CountCP18(0, txtCloseDate(0).Text, txtCloseDate(1).Text, txtSalesArea, txtSalesArea1, txtSales, strExc(3), , True)
      If Val(Mid(txtCloseDate(1), 1, Len(txtCloseDate(1)) - 2)) <= Val(strA0b05_1) And Val(Mid(txtCloseDate(1), 1, Len(txtCloseDate(1)) - 2)) <= Val(strA0b05_J) Then 'Added by Lydia 2021/09/16 +判斷;輸入日期為已決算的月份
           strExc(3) = "select cp13 AS Y0 ,sum(收文點數) 收文點數 from (" & strExc(3) & ") group by cp13"
      'Added by Lydia 2021/09/16
      Else
           strExc(3) = "select cp13 AS E0 ,sum(收文點數) 收文點數 from (" & strExc(3) & ") group by cp13"
      End If
      'end 2021/09/16
      
      '未收款點數：到迄日為止的未收款
       'Modified by Lydia 2021/07/30 (總計)未收款點數，請不要傳入日期止日；因為不管收據日期只要是未收款都列入計算，但不改模組，怕將來有其他需求
      'Call PUB_CountCP18(1, "", txtCloseDate(1).Text, txtSalesArea, txtSalesArea1, txtSales, strExc(4), , True)
      Call PUB_CountCP18(1, "", "", txtSalesArea, txtSalesArea1, txtSales, strExc(4), , True)
      strExc(4) = "select a0k20 AS R0,sum(未收款) 未收款 from (" & strExc(4) & ") group by a0k20"
      'end 2021/07/27
      
      '點數
      'modfiy by sonia 2014/6/30 結餘不限制科目改為41xx  D103062357的416102
      'stVTBw = "select ax209 W0, sum(ax207) W1" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null and (substr(ax205, 1, 2) = '41' Or ax205 = '7121')" & _
         " and ax207>0 and not( ax205='4191' or ax205='4192'" & _
         " or ((ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131') and instr(ax213||' ','結餘')>0))" & _
         " group by ax209"
      'modify by sonia 2015/4/23 加不含'4194'科目,不含結餘傳票不要加科目限制
      'stVTBw = "select ax209 W0, sum(ax207) W1" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null and (substr(ax205, 1, 2) = '41' Or ax205 = '7121')" & _
         " and ax207>0 and not( ax205='4191' or ax205='4192'" & _
         " or (substr(ax205, 1, 2) = '41' and instr(ax213||' ','結餘')>0))" & _
         " group by ax209"
      'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
      'Modify by Amy 2020/03/31 +7129
      stVTBw = "select ax209 W0, sum(ax207) W1" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null and (substr(ax205, 1, 1) = '4' Or ax205 = '7121' Or ax205 = '7129')" & _
         " and ax207>0 and not( ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0)" & _
         " group by ax209"
      
      '動用保留
      '2015/2/4 modify by sonia 保留欄不含4194結餘保留
      'stVTBx = "select ax209 X0,sum(ax207-ax206) X1" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null and (ax205='4191' or ax205='4192' or ax205='4194')" & _
         " group by ax209"
      stVTBx = "select ax209 X0,sum(ax207-ax206) X1" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null and (ax205='4191' or ax205='4192')" & _
         " group by ax209"
      
      '動用結餘
      'modfiy by sonia 2014/6/30 結餘不限制科目改為41xx  D103062357的416102
      'stVTBy = "select ax209 Y0,sum(ax207-ax206) Y1" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null  and (ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131')" & _
         " and instr(ax213||' ','結餘')>0 group by ax209 "
      '2015/2/4 modify by sonia 動用結餘欄要包含4194結餘保留(但ax213不會輸'結餘')
      'stVTBy = "select ax209 Y0,sum(ax207-ax206) Y1" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null  and substr(ax205, 1, 2) = '41' " & _
         " and instr(ax213||' ','結餘')>0 group by ax209 "
      'Modify by Amy 2019/08/01 增加創新業務組用收入 420101  原:substr(ax205, 1, 2) = '41'
      stVTBy = "select ax209 Y0,sum(ax207-ax206) Y1" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null  and substr(ax205, 1, 1) = '4' " & _
         " and (instr(ax213||' ','結餘')>0 or ax205='4194') group by ax209 "
      
      '扣點數
      'modfiy by sonia 2014/6/30 結餘不限制科目改為41xx  D103062357的416102
      'stVTBz = "select ax209 Z0,sum(-1*ax206) Z1" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null and ax206>0 and (substr(ax205, 1, 2) = '41' Or ax205 = '7121')" & _
         " and not ( (ax205='4191' or ax205='4192')" & _
         " or ( (ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131') and instr(ax213||' ','結餘')>0 ) )" & _
         " group by ax209"
      'modify by sonia 2015/4/23 加不含'4194'科目,不含結餘傳票不要加科目限制
      'stVTBz = "select ax209 Z0,sum(-1*ax206) Z1" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null and ax206>0 and (substr(ax205, 1, 2) = '41' Or ax205 = '7121')" & _
         " and not ( (ax205='4191' or ax205='4192')" & _
         " or (substr(ax205, 1, 2) = '41' and instr(ax213||' ','結餘')>0 ) )" & _
         " group by ax209"
      'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
      'Modify by Amy 2020/03/31 +7129
      stVTBz = "select ax209 Z0,sum(-1*ax206) Z1" & _
         " from acc020, acc021 where ax201(+) = a0201  and ax202(+) = a0202" & stCon & _
         " and ax209 is not null and ax206>0 and (substr(ax205, 1, 1) = '4' Or ax205 = '7121' Or ax205 = '7129')" & _
         " and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0)" & _
         " group by ax209"
      'Modified by Lydia 2021/07/27 最後面增加「收文點數」、「未收款點數」
      'strSql = "select st02||'('||st01||')' 智權人員,null 點數總計" & _
         ",to_char(ROUND(nvl(W1,0)/1000,3)+ROUND(nvl(Z1,0)/1000,3),'999999.000') 財務累計" & _
         ",to_char(ROUND(nvl(X1,0)/1000,3),'999999.000') 動用保留" & _
         ",to_char(ROUND(nvl(Y1,0)/1000,3),'999999.000') 動用結餘" & _
         ",to_char(ROUND(nvl(Z1,0)/1000,3),'999999.000') 扣點數" & _
         ",decode(substr(st15,1,1),'S',st06,'F','6','5') Zone" & _
         ",st15 Area,a0902 AreaName,st01" & _
         " from staff,( " & stVTBw & " ) W,( " & stVTBx & " ) X,( " & stVTBy & " ) Y,( " & stVTBz & " ) Z,acc090" & _
         " where W0(+)=st01 and X0(+)=st01 and Y0(+)=st01 and Z0(+)=st01" & stConST & _
         " and a0901(+)=st15 and (nvl(W1,0)<>0 or nvl(Y1,0)<>0 or nvl(X1,0)<>0 or nvl(Z1,0)<>0)"
      If Val(Mid(txtCloseDate(1), 1, Len(txtCloseDate(1)) - 2)) <= Val(strA0b05_1) And Val(Mid(txtCloseDate(1), 1, Len(txtCloseDate(1)) - 2)) <= Val(strA0b05_J) Then '輸入日期為已決算的月份
          '「動用保留」改抓「實績動用」，「動用結餘」改為「結餘動用」
          '參考frm210152, 「實績動用」=期初實績-期末實績，若個人尚未輸入則期末實績=期初實績；因為當月實績在當月會全報，所以不用加入計算
                                    '「動用結餘」=期初結餘+當月結餘-期末結餘，若個人尚未輸入則期末結餘=期初結餘+當月結餘(在前面語法已處理未輸入期末)
          strSql = "select st02||'('||st01||')' 智權人員,null 點數總計" & _
                     ",to_char(ROUND(nvl(W1,0)/1000,3)+ROUND(nvl(Z1,0)/1000,3),'999999.000') 財務累計" & _
                     ",to_char(C1-C5,'999999.000') 實績動用" & _
                     ",to_char(C2+C4-C6,'999999.000') 結餘動用" & _
                     ",to_char(ROUND(nvl(Z1,0)/1000,3),'999999.000') 扣點數" & _
                     ",to_char(nvl(Y.收文點數,0),'999999.000') 收文點數" & _
                     ",to_char(nvl(C.PE04,0),'999999.000') 目標點數" & _
                     ",to_char(ROUND(nvl(R.未收款,0)/1000,3),'999999.000') 未收款點數" & _
                     ",decode(substr(st15,1,1),'S',st06,'F','6','5') Zone" & _
                     ",st15 Area,a0902 AreaName,st01" & _
                     " from staff ,( " & stVTBw & " ) W,( " & strExc(2) & " ) C,( " & strExc(3) & " ) Y,( " & stVTBz & " ) Z,( " & strExc(4) & " ) R ,acc090" & _
                     " where W0(+)=st01 and C0(+)=st01 and Y0(+)=st01 and Z0(+)=st01 and R0(+)=st01" & stConST & _
                     " and (nvl(W1,0)<>0 or nvl(Z1,0)<>0 or nvl(C1,0) <> 0 or nvl(C5,0) <> 0 or nvl(C2,0) <> 0 or nvl(C4,0) <> 0 or nvl(C6,0) <> 0 or nvl(C.PE04,0) <> 0 or nvl(Y.收文點數,0) <> 0 or nvl(R.未收款,0) <> 0)"
      Else
          '原「財務累計」直接改名為「已入帳點數」，「動用保留」改抓「現有實績」=期初實績，「動用結餘」改抓「現有結餘」=期初結餘
          'Modified by Lydia 2021/09/16 當月期間的現有結餘改回舊抓法; 因為每月2,13日有可能輸入結餘
          'strSql = "select st02||'('||st01||')' 智權人員,null 點數總計" & _
                     ",to_char(ROUND(nvl(W1,0)/1000,3)+ROUND(nvl(Z1,0)/1000,3),'999999.000') 已入帳點數" & _
                     ",to_char(C1,'999999.000') 現有實績" & _
                     ",to_char(C2,'999999.000') 現有結餘" & _
                     ",to_char(ROUND(nvl(Z1,0)/1000,3),'999999.000') 扣點數" & _
                     ",to_char(nvl(Y.收文點數,0),'999999.000') 收文點數" & _
                     ",to_char(nvl(C.PE04,0),'999999.000') 目標點數" & _
                     ",to_char(ROUND(nvl(R.未收款,0)/1000,3),'999999.000') 未收款點數" & _
                     ",decode(substr(st15,1,1),'S',st06,'F','6','5') Zone" & _
                     ",st15 Area,a0902 AreaName,st01" & _
                     " from staff ,( " & stVTBw & " ) W,( " & strExc(2) & " ) C,( " & strExc(3) & " ) Y,( " & stVTBz & " ) Z,( " & strExc(4) & " ) R ,acc090" & _
                     " where W0(+)=st01 and C0(+)=st01 and Y0(+)=st01 and Z0(+)=st01 and R0(+)=st01" & stConST & _
                     " and (nvl(W1,0)<>0 or nvl(Z1,0)<>0 or nvl(C1,0) <> 0 or nvl(C2,0) <> 0 or nvl(C.PE04,0) <> 0 or nvl(Y.收文點數,0) <> 0 or nvl(R.未收款,0) <> 0)"
          strSql = "select st02||'('||st01||')' 智權人員,null 點數總計" & _
                     ",to_char(ROUND(nvl(W1,0)/1000,3)+ROUND(nvl(Z1,0)/1000,3),'999999.000') 已入帳點數" & _
                     ",to_char(C1,'999999.000') 現有實績" & _
                     ",to_char(ROUND(nvl(Y1,0)/1000,3),'999999.000') 現有結餘" & _
                     ",to_char(ROUND(nvl(Z1,0)/1000,3),'999999.000') 扣點數" & _
                     ",to_char(nvl(E.收文點數,0),'999999.000') 收文點數" & _
                     ",to_char(nvl(C.PE04,0),'999999.000') 目標點數" & _
                     ",to_char(ROUND(nvl(R.未收款,0)/1000,3),'999999.000') 未收款點數" & _
                     ",decode(substr(st15,1,1),'S',st06,'F','6','5') Zone" & _
                     ",st15 Area,a0902 AreaName,st01" & _
                     " from staff ,( " & stVTBw & " ) W,( " & strExc(2) & " ) C,( " & stVTBy & " ) Y,( " & stVTBz & " ) Z,( " & strExc(4) & " ) R,( " & strExc(3) & " ) E ,acc090" & _
                     " where W0(+)=st01 and C0(+)=st01 and Y0(+)=st01 and Z0(+)=st01 and R0(+)=st01 and E0(+)=st01" & stConST & _
                     " and (nvl(W1,0)<>0 or nvl(Z1,0)<>0 or nvl(C1,0) <> 0 or nvl(C2,0) <> 0 or nvl(C.PE04,0) <> 0 or nvl(Y.Y1,0) <> 0  or nvl(E.收文點數,0) <> 0 or nvl(R.未收款,0) <> 0)"
      End If
      'end 2021/07/27
      If txtZone <> "" Then
         strSql = strSql & " and substr(st15,1,1)='S'"
      End If
      
      strSql = strSql & " and a0901(+)=st15 order by Zone,Area,st01"
      
   'Add by Amy 2013/08/07
   '扣點數 'Memo by Lydia 2021/07/27 銷帳及扣點數合併=5, 原本銷帳選項2和扣點數選項4隱藏保留
   ElseIf p_iMode = 4 Then
      LblNote = ""
      
      'modfiy by sonia 2014/6/30 結餘不限制科目改為41xx  D103062357的416102
      'strSql = "select st02,A0205,ax202, ax214, ax212, to_char(round((ax207-ax206)/1000,3),'999999.000') Point, ax209,ax203" & _
         " from acc020, acc021, staff where st01(+)=ax209 and ax209 is not null and ax206>0" & strExc(0) & _
         " and (substr(ax205, 1, 2) = '41' Or ax205 = '7121') and not ( (ax205='4191' or ax205='4192') or ( (ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131') " & _
         " and instr(ax213||' ','結餘')>0 ) ) " & stCon & stConST
      'modify by sonia 2015/4/23 加不含'4194'科目,不含結餘傳票不要加科目限制
      'strSql = "select st02,A0205,ax202, ax214, ax212, to_char(round((ax207-ax206)/1000,3),'999999.000') Point, ax209,ax203" & _
         " from acc020, acc021, staff where st01(+)=ax209 and ax209 is not null and ax206>0" & strExc(0) & _
         " and (substr(ax205, 1, 2) = '41' Or ax205 = '7121') and not ( (ax205='4191' or ax205='4192') or ( substr(ax205, 1, 2) = '41' " & _
         " and instr(ax213||' ','結餘')>0 ) ) " & stCon & stConST
      'Modify by Amy 2019/08/01 增加創新業務組用收入 420101  原:substr(ax205, 1, 2) = '41'
      'Modify by Amy 2020/03/31 +7129
      strSql = "select st02,A0205,ax202, ax214, ax212, to_char(round((ax207-ax206)/1000,3),'999999.000') Point, ax209,ax203" & _
         " from acc020, acc021, staff where st01(+)=ax209 and ax209 is not null and ax206>0" & strExc(0) & _
         " and (substr(ax205, 1, 1) = '4' Or ax205 = '7121' Or ax205 = '7129') and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) " & stCon & stConST
         
      'Added by Morgan 2013/8/14
      '以本所案號查詢時要同時抓相同案的支援扣點數
      If Text1 <> "" Then
         strSql = strSql & " union all select st02,A0205,ax202, ax214, ax212, to_char(round((ax207-ax206)/1000,3),'999999.000') Point, ax209,ax203" & _
            " from acc020, acc021 a, staff where st01(+)=ax209 and ax209 is not null and ax206>0" & strExc(0) & _
            " and substr(ax205,1,1)='4' and ax206=5000" & stCon2 & stConST & " and exists(select * from acc021 b where b.ax202=a.ax202 and b.ax205=a.ax205 and b.ax207=a.ax206 and b.ax214=a.ax214 and b.ax209='P1001')"
      End If
      strSql = strSql & " order by ax209,A0205,ax202,ax203 asc"
   'end 2013/08/07
   
   'Added by Lydia 2021/07/27 銷帳及扣點數合併=5, 原本銷帳選項2和扣點數選項4隱藏保留
   ElseIf p_iMode = 5 Then
      LblNote = ""
      '銷帳
      If Trim(Text1) <> "" Then  '本所案號查詢
        'Modified by Morgan 2023/1/7 銷帳服務費可能為負(扣點數) A1U07>0->A1U07<>0
        stVTBw = "SELECT st02 AS W0, a0s03 AS W1,A0K01 AS W2,A0J02 AS W3,A0K04||'/'||getcp10desc(cp01,cp10,a0j04) AS W4,to_char(ROUND(nvl(A1U07,0)/1000,3),'999999.000') AS W5, A0K20 AS W6,'001' AS W7, ST15 AS ZONE,'W' AS ORD1 " & _
             " From ACC0S0,ACC0K0,staff,ACC1U0,ACC0J0,caseprogress WHERE A0S04='1' AND A0S03>=940601" & _
             " AND A0J02='" & Text1 & Text2 & Text3 & Text4 & "' " & _
             " AND A0K01(+)=A0S02 AND ST01(+)=A0K20" & stConST & _
             " AND A1U01=A0S01(+) AND A0J01=A1U03(+) AND A1U07<>0 and cp09(+)=a0j01"
      Else
        'Modified by Morgan 2023/1/7 銷帳服務費可能為負(扣點數) A1U07>0->A1U07<>0
        stVTBw = "SELECT st02 AS W0, a0s03 AS W1, A0K01 AS W2, A0J02 AS W3, A0K04||'/'||getcp10desc(cp01,cp10,a0j04) W4, to_char(ROUND(nvl(A1U07,0)/1000,3),'999999.000') AS W5,A0K20 AS W6,'001' as W7, ST15 AS ZONE,'W' AS ORD1 " & _
             " From ACC0S0,ACC0K0,staff,ACC1U0,ACC0J0,caseprogress WHERE A0S04='1' AND A0S03>=940601" & _
             " AND A0S03>=" & txtCloseDate(0) & " AND A0S03 <= " & txtCloseDate(1) & _
             " AND A0K01(+)=A0S02 AND ST01(+)=A0K20" & stConST & _
             " AND A1U01(+)=A0S01 AND A0J01(+)=A1U03 AND A1U07<>0 and cp09(+)=a0j01"
      End If
      '-----2021/7/29 E11016451銷帳一筆資料變三筆
      stVTBw = "SELECT W0,W1,W2,W3,W4,W5,W6,W7,ZONE,ORD1 FROM (" & stVTBw & ") GROUP BY W0,W1,W2,W3,W4,W5,W6,W7,ZONE,ORD1"
      
      '扣點數
      stVTBx = " select st02,A0205,ax202, ax214, ax212, to_char(round((ax207-ax206)/1000,3),'999999.000') Point, ax209,ax203,ST15 AS ZONE,'X' AS ORD1 " & _
         " from acc020, acc021, staff where st01(+)=ax209 and ax209 is not null and ax206>0" & strExc(0) & _
         " and (substr(ax205, 1, 1) = '4' Or ax205 = '7121' Or ax205 = '7129') and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0) " & stCon & stConST
      
      If Text1 <> "" Then  '以本所案號查詢時要同時抓相同案的支援扣點數
         stVTBx = stVTBx & "union all select  st02,A0205,ax202, ax214, ax212, to_char(round((ax207-ax206)/1000,3),'999999.000') Point, ax209,ax203,ST15 AS ZONE,'X' AS ORD1 " & _
            " from acc020, acc021 a, staff where st01(+)=ax209 and ax209 is not null and ax206>0" & strExc(0) & _
            " and substr(ax205,1,1)='4' and ax206=5000" & stCon2 & stConST & " and exists(select * from acc021 b where b.ax202=a.ax202 and b.ax205=a.ax205 and b.ax207=a.ax206 and b.ax214=a.ax214 and b.ax209='P1001')"
      End If
      strSql = "SELECT W0 AS 智權人員,W3 AS 本所案號,DECODE(ORD1,'W',W2,NULL) AS 收據號碼, W1 AS 傳票日期, DECODE(ORD1,'W',NULL,W2) AS 傳票號碼, W4 AS 摘要, DECODE(ORD1,'W',W5,NULL) AS 銷帳點數, DECODE(ORD1,'W',NULL,W5) AS 扣點數, W6 AS ST01, W7 AS AX203,ZONE " & _
                  "from ( " & stVTBw & " union " & stVTBx & ") ORDER BY ZONE,ST01,4"  '依業務區+業務員+傳票日期排序
      
   'end 2021/07/27
   End If
   
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         Set grdDataList.Recordset = AdoRecordSet3.Clone
         Call SetDataListWidth(p_iMode, True)
         If p_iMode = 1 Then  '點數統計
            Calculate
         ElseIf p_iMode = 3 Then '個人總表
            Calculate2
         Else
            Calculate1 p_iMode
         End If
      Else
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

Private Sub cmdSearch_Click()
   Dim iOpt As Integer, oOption As OptionButton
   Screen.MousePointer = vbHourglass
   If ConstrainCheck = True Then
      For Each oOption In Option1
         If oOption.Value = True Then
            iOpt = oOption.Index
            Exit For
         End If
      Next
      Call SetDataListWidth(iOpt)
      Call doQuery(iOpt)
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Call SetDataListWidth(Abs(Option1(1).Value))
   LblNote.Caption = "" 'Added by Lydia 2021/07/27
   
   stST15 = PUB_GetStaffST15(strUserNum, 1)
   stST05 = PUB_GetST05(strUserNum)
   
   'Added by Lydia 2024/01/31 不預設日期
   If strUserNum = "75007" Then
      txtCloseDate(0) = ""
      txtCloseDate(1) = ""
   Else
   'end 2024/01/31
      'Add by Morgan 2005/7/4 預設1號到當天
      txtCloseDate(0) = strSrvDate(2) \ 100 & "01"
      txtCloseDate(1) = strSrvDate(2)
   End If
   
   bolAreaMan = False 'Add By Sindy 2023/6/12
   'Modify By Sindy 2020/7/28 設定員編,部門,所別權限
   'Modify By Sindy 2025/3/18 +Me.Name
   Call PUB_SetFormSaleDept(strUserNum, txtZone, txtSalesArea, txtSalesArea1, txtSales, bolSpecMan, strSpecCode, , , , , , , , Me.Name)
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
      If InStr(m_strListPer, GetDeptMan(stST15)) > 0 Then
         bolAreaMan = True
      End If
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
'      'modify by sonia 2014/6/9 +美珍77027
'      Case "65001", "68006", "77027"
'         txtZone.Enabled = True
'         txtSalesArea.Enabled = True
'         txtSalesArea1.Enabled = True
'         txtSales.Enabled = True
'         'Add by Morgan 2005/7/4 副總預設所有業務區
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
'      '郭雅娟也可看專利處 2014/8/15
'      Case "71011", "79075"
'         '2007/10/17 modify by sonia 不預設所別
'         'txtZone = pub_strUserOffice
'         txtZone.Enabled = True
'         txtZone = ""
'         '2007/10/17 end
'         txtSalesArea = "P10"
'         txtSalesArea1 = "P19"
'         txtSales.Enabled = True
'      '葉經理可看商標處
'      'modify by sonia 2016/2/24 +69008
'      Case "67002", "69008"
'         '2007/10/17 modify by sonia 加入96029巨京商標故不預設所別
'         'txtZone = pub_strUserOffice
'         txtZone.Enabled = True
'         txtZone = ""
'         '2007/10/17 end
'         txtSalesArea = "P20"
'         txtSalesArea1 = "P29"
'         txtSales.Enabled = True
'      'Add by Morgan 2007/3/23
'      '何副總可看國外部
'      'modify by sonia 2015/6/30 改68009為81040
'      Case "81040"
'         txtZone = pub_strUserOffice
'         txtSalesArea = "F10"
'         txtSalesArea1 = "F41"
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
'            'MODIFY BY SONIA 2015/6/1 中所加CM
'            Case "CM", "C1", "NM", "KM"
'               txtZone = pub_strUserOffice
'               txtSalesArea.Enabled = True
'               txtSalesArea1.Enabled = True
'               txtSales.Enabled = True
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
   
   'Modify By Sindy 2025/3/18 mark
'   'Modify By Sindy 權限特例
'   Select Case strUserNum
'      '王協理可看專利處
'      '郭雅娟也可看專利處 2014/8/15
'      Case "71011", "79075"
'         '2007/10/17 modify by sonia 不預設所別
'         'txtZone = pub_strUserOffice
'         txtZone.Enabled = True
'         txtZone = ""
'         '2007/10/17 end
'         txtSalesArea = "P10"
'         txtSalesArea1 = "P19"
'         txtSales.Enabled = True
'      'Add By Sindy 2020/10/5 總經理:業績點數查詢，林總同意楊監察人可查詢全所人員資料
'      Case "69009" '楊毓純
'         txtZone.Enabled = True
'         txtZone = ""
'         txtSalesArea.Enabled = True
'         txtSalesArea1.Enabled = True
''         txtSalesArea = ""
''         txtSalesArea1 = ""
'         txtSales.Enabled = True
'         txtSales = ""
'      '2020/10/5 END
'      Case Else
'         Select Case stST05
'            '分所財務人員可看該所全部
'            'MODIFY BY SONIA 2015/6/1 中所加CM
'            Case "CM", "C1", "NM", "KM"
'               txtZone = pub_strUserOffice
'               txtSalesArea.Enabled = True
'               txtSalesArea1.Enabled = True
'               txtSales.Enabled = True
'         End Select
'   End Select
   
'   'Add By Sindy 2009/05/12
'   '若操作人員的ST05=SA,此類人員稱之為帶人主管,開放智權人員欄位可輸入
'   If PUB_GetST05Limits(strUserNum) = True Then
'      txtSales.Enabled = True
'   End If
'
   'Add By Sindy 2016/5/6 記錄原操作人可以查詢的業務區及所別
   txtZone.Tag = txtZone
   txtSalesArea.Tag = txtSalesArea
   txtSalesArea1.Tag = txtSalesArea1
   '2016/5/6 END
   
   'Added by Lydia 2021/07/27
   strA0b01_1 = GetA0b01(strA0b05_1, "1")
   strA0b01_J = GetA0b01(strA0b05_J, "J")
   If Pub_StrUserSt03 = "M51" Then lblMemo.Visible = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
   MenuEnabled
   Set frm210104 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
   'Added by Lydia 2024/02/01
   If Val(txtCloseDate(0)) = 0 Then
      MsgBox "請輸入點數結算起日！", vbExclamation
      txtCloseDate(0).SetFocus
      Call txtCloseDate_GotFocus(0)
      Exit Sub
   End If
   If Val(txtCloseDate(1)) = 0 Then
      MsgBox "請輸入點數結算止日！", vbExclamation
      txtCloseDate(1).SetFocus
      Call txtCloseDate_GotFocus(1)
      Exit Sub
   End If
   'end 2024/02/01
   SetDataListWidth (Index)
   'Modify by Morgan 2010/8/30 已加判斷離職日大於統計日的也要
   'If Index = 1 Then
   '   lblCaution.Visible = True
   'Else
   '   lblCaution.Visible = False   '2013/10/8 cancel by sonia 無用取消
   'End If
End Sub

'Add by Amy 2013/08/07 +本所案號查詢(只可點選收款及扣點數、銷帳、扣點數)
Private Sub Text1_Change()
    SetQueryOption
End Sub

Private Sub Text1_GotFocus()
    TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    
    If Len(Trim(Text1)) > 0 Then
        If Text1 <> "FCP" And Text1 <> "FG" And Text1 <> "P" And Text1 <> "PS" And Text1 <> "CFP" And Text1 <> "CPS" _
            And Text1 <> "CFT" And Text1 <> "CFC" And Text1 <> "FCT" And Text1 <> "S" _
            And Text1 <> "CFL" And Text1 <> "FCL" And Text1 <> "LIN" And Text1 <> "L" And Text1 <> "LA" _
            And Text1 <> "T" And Text1 <> "TB" And Text1 <> "TC" And Text1 <> "TD" And Text1 <> "TF" _
            And Text1 <> "TM" And Text1 <> "TR" And Text1 <> "TS" And Text1 <> "TT" Then
            
            MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
            Cancel = True
        End If
   End If
End Sub

Private Sub Text2_Change()
    SetQueryOption
End Sub

Private Sub Text2_GotFocus()
    TextInverse Text2
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
    If Len(Trim(Text2)) > 0 And Len(Trim(Text2)) <> 6 Then
       MsgBox "本所案號輸入錯誤，請重新輸入 !", vbCritical
      Cancel = True
    End If
End Sub

Private Sub Text3_Change()
    SetQueryOption
End Sub

Private Sub Text3_GotFocus()
    TextInverse Text3
End Sub

Private Sub Text4_Change()
    SetQueryOption
End Sub

Private Sub Text4_GotFocus()
    TextInverse Text4
End Sub

Private Sub SetQueryOption()
    If Text1 & Text2 = "" Then
        Text3 = "": Text4 = ""
        Option1(1).Enabled = True: Option1(3).Enabled = True
    Else
        If Option1(1).Value Or Option1(3).Value Then Option1(0).Value = True
        Option1(1).Enabled = False: Option1(3).Enabled = False
    End If
End Sub
'end 2013/08/07

Private Sub txtCloseDate_GotFocus(Index As Integer)
   SetQueryOption 'Add by Amy 2013/08/07 +本所案號查詢
   'If Index = 1 Then txtCloseDate(Index) = txtCloseDate(Index - 1)
   TextInverse txtCloseDate(Index)
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtCloseDate(Index).IMEMode = 2
   CloseIme
End Sub

Private Sub txtCloseDate_Validate(Index As Integer, Cancel As Boolean)
   If txtCloseDate(Index) <> "" Then
      If ChkDate(txtCloseDate(Index)) = False Then
         Cancel = True
         txtCloseDate(Index).SetFocus
         txtCloseDate_GotFocus Index
      End If
   End If
   'Added by Lydia 2024/01/31
   If strUserNum = "75007" Then
      If DBDATE(txtCloseDate(Index)) >= "20240201" Then
         MsgBox "不可查詢113年2月後的資料！", vbExclamation
         Cancel = True
         txtCloseDate(Index).SetFocus
         txtCloseDate_GotFocus Index
      End If
   End If
   'end 2024/01/31
   
   'Added by Lydia 2021/07/27
   'Memo by Lydia 2021/08/10 +點數統計
   If Option1(1).Value = True Or Option1(3).Value = True Then
      If Option1(1).Value = True Then strExc(2) = Option1(1).Caption
      If Option1(3).Value = True Then strExc(2) = Option1(3).Caption
      If Val(txtCloseDate(0)) > 0 Then 'Added by Lydia 2024/02/01
         If Index = 0 And Val(Mid(txtCloseDate(0), 1, Len(txtCloseDate(0)) - 2)) < Val(業績輸入啟用年月) Then
           MsgBox strExc(2) & "不可查詢105年1月前的資料"
           txtCloseDate_GotFocus Index
           Cancel = True
           Exit Sub
         End If
      End If 'Added by Lydia 2024/02/01
      
      If Index = 1 Then
         If Val(txtCloseDate(0)) > 0 And Val(txtCloseDate(1)) > 0 Then     'Added by Lydia 2024/02/01
            If Val(Mid(txtCloseDate(0), 1, Len(txtCloseDate(0)) - 2)) > Val(txtCloseDate(1)) Then
               MsgBox strExc(2) & "起始年月不可大於截止年月"
               txtCloseDate_GotFocus Index
               Cancel = True
               Exit Sub
            End If
            strExc(1) = ""
            If Val(Mid(txtCloseDate(0), 1, Len(txtCloseDate(0)) - 2)) <= Val(strA0b05_1) And Val(Mid(txtCloseDate(1), 1, Len(txtCloseDate(1)) - 2)) > Val(strA0b05_1) Then
               strExc(1) = Mid(strA0b05_1, 1, 3) & "年" & Mid(strA0b05_1, 4, 2) & "月"
            End If
            If Val(Mid(txtCloseDate(0), 1, Len(txtCloseDate(0)) - 2)) <= Val(strA0b05_J) And Val(Mid(txtCloseDate(1), 1, Len(txtCloseDate(1)) - 2)) > Val(strA0b05_J) Then
                 strExc(1) = Mid(strA0b05_J, 1, 3) & "年" & Mid(strA0b05_J, 4, 2) & "月"
            End If
            If strExc(1) <> "" Then
                 MsgBox "點數結算日區間不可跨過結算期" & strExc(1) & "，請重新輸入！", vbExclamation
                 txtCloseDate(1).SetFocus
                 txtCloseDate_GotFocus (1)
                 Cancel = True
                 Exit Sub
            End If
         End If 'Added by Lydia 2024/02/01
      End If
   End If
   'end 2021/07/27
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
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtSales.IMEMode = 2
   CloseIme
   If Combo3.Enabled = True And Combo3.Visible = True Then Combo3.SetFocus 'Add By Sindy 2023/6/12
End Sub

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
   'Modify By Sindy 2025/3/17 +Me.Name
   If PUB_txtSales_Limit(txtSales, m_strListPer, txtZone, txtSalesArea, txtSalesArea1, _
                         bolSpecMan, strSpecCode, lblSalesName, Me.Name) = False Then
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
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtSalesArea.IMEMode = 2
   CloseIme
End Sub

Private Sub txtSalesArea_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesArea1_GotFocus()
   TextInverse txtSalesArea1
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtSalesArea.IMEMode = 2
   CloseIme
End Sub

Private Sub txtSalesArea1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtZone_GotFocus()
   TextInverse txtZone
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtZone.IMEMode = 2
   CloseIme
End Sub

Private Sub txtZone_KeyPress(KeyAscii As Integer)
   If (KeyAscii < Asc("1") Or KeyAscii > Asc("4")) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub SetColor(iRow As Integer, lngColor As Long)
   Dim ii As Integer, jj As Integer
   With grdDataList
   .row = iRow
   For jj = 0 To .Cols - 1
      .col = jj: .CellBackColor = lngColor
   Next
   End With
End Sub

Private Function GetZoneName(p_Zone As String) As String
   Select Case p_Zone
      Case "1"
         GetZoneName = "台北所"
      Case "2"
         GetZoneName = "台中所"
      Case "5"
         GetZoneName = "其他"
      Case "6"
         GetZoneName = "國外"
   End Select

End Function

'Add by Morgan 2010/8/27
Private Sub Calculate2()
   Dim ii As Integer, jj As Integer, dblNet As Double
   Dim stZone As String, stZoneName As String, stArea As String, stAreaName As String
   'Modified by Lydia 2021/07/27
   'Dim dblSub1(1 To 5) As Double, dblSub2(1 To 5) As Double, dblSub3(1 To 5) As Double
   Dim dblSub1() As Double, dblSub2() As Double, dblSub3() As Double
   Dim strAddItem As String, lngColor As Long
   
   'Added by Lydia 2021/07/27 記錄欄位位置
   Dim colZone As Integer, colArea As Integer, colAreaName As Integer
   colZone = PUB_MGridGetId("ZONE", grdDataList)
   colArea = PUB_MGridGetId("AREA", grdDataList)
   colAreaName = PUB_MGridGetId("AREANAME", grdDataList)
   ReDim dblSub1(1 To colZone - 1) As Double
   ReDim dblSub2(1 To colZone - 1) As Double
   ReDim dblSub3(1 To colZone - 1) As Double
   'end 2021/07/27
   
   With grdDataList
      .Visible = False
      'Modified by Lydia 2021/07/27 改用變數
      'stZone = .TextMatrix(1, 6)
      'stArea = .TextMatrix(1, 7)
      'stAreaName = .TextMatrix(1, 8)
      stZone = .TextMatrix(1, colZone)
      stArea = .TextMatrix(1, colArea)
      stAreaName = .TextMatrix(1, colAreaName)
      'end 2021/07/27
      ii = 1
      Do While ii < .Rows
         '區小計
         'Modified by Lydia 2021/07/27 改用變數
         'If .TextMatrix(ii, 7) <> stArea Then
         If .TextMatrix(ii, colArea) <> stArea Then
            '智權人員才要
            If txtSales = "" And Left(stArea, 1) = "S" Then
               strAddItem = stAreaName & "合計:"
               'Modified by Lydia 2021/07/27 改用變數
               'For jj = 1 To 5
               For jj = 1 To colZone - 1
                  strAddItem = strAddItem & vbTab & Format(dblSub1(jj), "###,###,###.000")
               Next
               .AddItem strAddItem, ii
               If stArea = "S31" Or stArea = "S41" Then '台南所,高雄所
                  lngColor = &H90EE90
               Else
                  lngColor = &H7FFFD4
               End If
               SetColor ii, lngColor
               ii = ii + 1
            End If
            Erase dblSub1
            'Modified by Lydia 2021/07/27 改用變數
            'stArea = .TextMatrix(ii, 7)
            'stAreaName = .TextMatrix(ii, 8)
            stArea = .TextMatrix(ii, colArea)
            stAreaName = .TextMatrix(ii, colAreaName)
            ReDim dblSub1(1 To colZone - 1) As Double
            'end 2021/07/27
            
            '所小計
            'Modified by Lydia 2021/07/27 改用變數
            'If .TextMatrix(ii, 6) <> stZone Then
            If .TextMatrix(ii, colZone) <> stZone Then
               stZoneName = GetZoneName(stZone)
               If txtSales & txtSalesArea & txtSalesArea1 = "" And stZoneName <> "" Then
                  strAddItem = stZoneName & "合計:"
                  'Modified by Lydia 2021/07/27 改用變數
                  'For jj = 1 To 5
                  For jj = 1 To colZone - 1
                     strAddItem = strAddItem & vbTab & Format(dblSub2(jj), "###,###,###.000")
                  Next
                  .AddItem strAddItem, ii
                  If stZone = "5" Then '其他
                     lngColor = &H90EE40
                  Else
                     lngColor = &H90EE90
                  End If
                  SetColor ii, lngColor
                  ii = ii + 1
               End If
               '全所時,高所合計後面加印智權部合計
               If txtZone = "" Then
                  If stZone = "4" Then
                     strAddItem = "智權部合計:"
                     lngColor = &H90EE40
                  ElseIf stZone = "5" Then
                     strAddItem = "國內合計:"
                     lngColor = &HFFFF90
                  Else
                     strAddItem = ""
                  End If
                  If strAddItem <> "" Then
                     'Modified by Lydia 2021/07/27 改用變數
                     'For jj = 1 To 5
                     For jj = 1 To colZone - 1
                        strAddItem = strAddItem & vbTab & Format(dblSub3(jj), "###,###,###.000")
                     Next
                     .AddItem strAddItem, ii
                     SetColor ii, lngColor
                     ii = ii + 1
                  End If
               End If
               
               Erase dblSub2
               'Modified by Lydia 2021/07/27 改用變數
               'stZone = .TextMatrix(ii, 6)
               stZone = .TextMatrix(ii, colZone)
               ReDim dblSub2(1 To colZone - 1) As Double
               'end 2021/07/27
            End If
         End If
         
         .TextMatrix(ii, 1) = Val(.TextMatrix(ii, 2)) + Val(.TextMatrix(ii, 3)) + Val(.TextMatrix(ii, 4))
         'Modified by Lydia 2021/07/27 改用變數
         'For jj = 1 To 5
         For jj = 1 To colZone - 1
            dblSub1(jj) = dblSub1(jj) + Val(.TextMatrix(ii, jj))
            dblSub2(jj) = dblSub2(jj) + Val(.TextMatrix(ii, jj))
            dblSub3(jj) = dblSub3(jj) + Val(.TextMatrix(ii, jj))
            .TextMatrix(ii, jj) = Format(.TextMatrix(ii, jj), "###,###,###.000")
         Next
         ii = ii + 1
      Loop
      
      '區小計
      If txtSales = "" And Left(stArea, 1) = "S" Then
         strAddItem = stAreaName & "合計:"
         'Modified by Lydia 2021/07/27 改用變數
         'For jj = 1 To 5
         For jj = 1 To colZone - 1
            strAddItem = strAddItem & vbTab & Format(dblSub1(jj), "###,###,###.000")
         Next
         .AddItem strAddItem, ii
         If stArea = "S31" Or stArea = "S41" Then '台南所,高雄所
            lngColor = &H90EE90
         Else
            lngColor = &H7FFFD4
         End If
         SetColor ii, lngColor
         ii = ii + 1
      End If
      
      '所小計
      stZoneName = GetZoneName(stZone)
      If txtSales & txtSalesArea & txtSalesArea1 = "" And stZoneName <> "" Then
         strAddItem = stZoneName & "合計:"
         'Modified by Lydia 2021/07/27 改用變數
         'For jj = 1 To 5
         For jj = 1 To colZone - 1
            strAddItem = strAddItem & vbTab & Format(dblSub2(jj), "###,###,###.000")
         Next
         .AddItem strAddItem, ii
         If stZone = "5" Then '其他
            lngColor = &H90EE40
         ElseIf stZone = "6" Then '國外部
            lngColor = &HFFFF90
         Else
            lngColor = &H90EE90
         End If
         SetColor ii, lngColor
         ii = ii + 1
      End If
      
      '全所
      If txtSales & txtSalesArea & txtSalesArea1 & txtZone = "" Then
         strAddItem = "全所合計:"
         'Modified by Lydia 2021/07/27 改用變數
         'For jj = 1 To 5
         For jj = 1 To colZone - 1
            strAddItem = strAddItem & vbTab & Format(dblSub3(jj), "###,###,###.000")
         Next
         .AddItem strAddItem, ii
         SetColor ii, &HFFFF00
      End If
      
      .Visible = True
   End With
End Sub

'明細累計計算
Private Sub Calculate1(Optional ByVal p_iMode As Integer)
   Dim ii As Integer, dblSum As Double
   Dim dblESum As Double 'Added by Lydia 2021/07/27
   
   With grdDataList
      .Visible = False
      'Add by Amy 2013/08/07 +if 判斷因發現2012拿掉欄位合計有誤(銷帳)
      'Modified by Lydia 2021/07/27 +銷帳及扣點數合併=5
      If p_iMode = 0 Or p_iMode = 4 Then
         For ii = 2 To .Rows - 1
            dblSum = dblSum + Val(.TextMatrix(ii, 5))
            
            'Added by Morgan 2013/8/14
            If Text1 & Text2 <> "" Then
               If .TextMatrix(ii, 3) <> Text1 & Text2 & Text3 & Text4 Then
                  .row = ii
                  .col = 3
                  .CellBackColor = vbRed
                  .CellForeColor = vbWhite
               End If
            End If
        Next ii
      'Added by Lydia 2021/07/27 銷帳及扣點數合併=5
      ElseIf p_iMode = 5 Then
         For ii = 2 To .Rows - 1
            '銷帳、扣點數分開總計
            dblSum = dblSum + Val(.TextMatrix(ii, 6))
            dblESum = dblESum + Val(.TextMatrix(ii, 7))
            If Text1 & Text2 <> "" Then
               If .TextMatrix(ii, 2) <> Text1 & Text2 & Text3 & Text4 Then
                  .row = ii
                  .col = 2
                  .CellBackColor = vbRed
                  .CellForeColor = vbWhite
               End If
            End If
        Next ii
      'end 2021/07/27
      Else
        For ii = 2 To .Rows - 1
            dblSum = dblSum + Val(.TextMatrix(ii, 4))
        Next ii
      End If
      'Added by Lydia 2021/07/27 判斷類別
      If p_iMode = 0 Or p_iMode = 4 Then
         .TextMatrix(1, 5) = Format(dblSum, "###,###,###.000")
      ElseIf p_iMode = 5 Then
         .TextMatrix(1, 6) = Format(dblSum, "###,###,###.000")
         .TextMatrix(1, 7) = Format(dblESum, "###,###,###.000")
      Else  '銷帳
      'end 2021/07/27
         .TextMatrix(1, 4) = Format(dblSum, "###,###,###.000")
      End If 'Added by Lydia 2021/07/27
      .Visible = True
   End With
End Sub

'統計計算
Private Sub Calculate()
Dim ii As Integer, stLstSales As String, dblSum(1 To 3) As Double, jj As Integer, dblNet As Double
Dim bolFirstRow As Boolean
   
On Error GoTo ErrHnd
   
   'Add By Sindy 2020/8/31
   Load frmpic002
   frmpic002.Label1.Caption = "資料計算中...請稍候..."
   frmpic002.Show
   frmpic002.ZOrder 0: DoEvents
'   If PUB_IsFormExist("frm210137") = True Then
'      Unload frm210137
'   End If
'   If PUB_IsFormExist("frm210141") = True Then
'      Unload frm210141
'   End If
   '2020/8/31 END
   
   With grdDataList
      .Visible = False
      stLstSales = "": bolFirstRow = True
      For ii = .Rows - 1 To 2 Step -1
         If .TextMatrix(ii, 12) = 0 Then
            .row = ii
            If .TextMatrix(ii, 11) = stLstSales Then
               'Modify by Morgan 點數已扣除扣點數
               'dblNet = dblSum(1) + Val(.TextMatrix(ii, 4)) + Val(.TextMatrix(ii, 6)) + Val(.TextMatrix(ii, 8))
               .TextMatrix(ii, 4) = dblSum(1)
               dblNet = Val(.TextMatrix(ii, 4)) + Val(.TextMatrix(ii, 6)) + Val(.TextMatrix(ii, 8))
            Else
               dblNet = 0
            End If
            .TextMatrix(ii, 2) = Format(dblNet, "###,###,###.000")
            .TextMatrix(ii, 4) = Format(.TextMatrix(ii, 4), "###,###,###.000")
            .TextMatrix(ii, 6) = Format(.TextMatrix(ii, 6), "###,###,###.000")
            .TextMatrix(ii, 8) = Format(.TextMatrix(ii, 8), "###,###,###.000")
            .TextMatrix(ii, 10) = Format(.TextMatrix(ii, 10), "###,###,###.000")
            
            '已收文未收款點數
            'Add By Sindy 2020/10/5 顯示每日累計未收款點數
            If Check1.Visible = True And Check1.Value = 1 Then
               bolFirstRow = True
            Else
            '2020/10/5 END
               'Modified by Lydia 2021/07/30 (總計)未收款點數，請不要傳入日期止日；因為不管收據日期只要是未收款都列入計算，但不改模組，怕將來有其他需求
               '.TextMatrix(ii + 1, 9) = Format(PUB_CountCP18(1, "", DBDATE(.TextMatrix(ii + 1, 1)) - 19110000, , , .TextMatrix(ii + 1, 11)), "#.00")
               .TextMatrix(ii + 1, 9) = Format(PUB_CountCP18(1, "", "", , , .TextMatrix(ii + 1, 11)), "#.00")
            End If
            
            For jj = 1 To .Cols - 1
               .col = jj: .CellBackColor = &H90EE90
            Next

         Else
            .TextMatrix(ii, 1) = Format(.TextMatrix(ii, 1), "@@@/@@/@@")
            If .TextMatrix(ii, 11) <> stLstSales Then
               stLstSales = .TextMatrix(ii, 11)
               dblSum(1) = Val(.TextMatrix(ii, 3))
               dblSum(2) = Val(.TextMatrix(ii, 5)) - Val(.TextMatrix(ii, 6))
'               dblSum(3) = Val(.TextMatrix(ii, 10)) + Val(.TextMatrix(ii, 8)) - Val(.TextMatrix(ii, 5)) - Val(.TextMatrix(ii, 9))
            Else
               dblSum(1) = dblSum(1) + Val(.TextMatrix(ii, 3))
               dblSum(2) = dblSum(2) + Val(.TextMatrix(ii, 5)) - Val(.TextMatrix(ii, 6))
'               dblSum(3) = dblSum(3) + Val(.TextMatrix(ii, 8)) - Val(.TextMatrix(ii, 5)) - Val(.TextMatrix(ii, 9))
            End If
            .row = ii
            '差
            .col = 2: .CellBackColor = &H7FFFD4: .Text = Format(Val(.TextMatrix(ii, 3)) - (Val(.TextMatrix(ii, 5)) - Val(.TextMatrix(ii, 6))), "###,###,###.000")
            '累計
            .col = 4: .CellBackColor = &H7FFFD4: .Text = Format(dblSum(1), "###,###,###.000")
            .col = 7: .CellBackColor = &H7FFFD4: .Text = Format(dblSum(2), "###,###,###.000")
'            .col = 10: .CellBackColor = &H7FFFD4: .Text = Format(dblSum(3), "###,###,###.000")
            '當日
            .TextMatrix(ii, 3) = Format(.TextMatrix(ii, 3), "###,###,###.000")
            .TextMatrix(ii, 5) = Format(.TextMatrix(ii, 5), "###,###,###.000")
'            .TextMatrix(ii, 8) = Format(.TextMatrix(ii, 8), "###,###,###.000")
            
            'Add By Sindy 2020/8/31
            '已收文點數
'            .TextMatrix(ii, 8) = Format(PUB_CountCP18(0, txtCloseDate(0).Text, DBDATE(.TextMatrix(ii, 1)) - 19110000, , .TextMatrix(ii, 11)), "#.00")
'            frm210137.Hide
'            frm210137.txtSalesArea = "" '業務區(起)
'            frm210137.txtSalesArea1 = "" '業務區(迄)
'            frm210137.txtSales = .TextMatrix(ii, 11) '智權人員ID
'            frm210137.txtCloseDate(0) = txtCloseDate(0).Text '點數結算日(起)
'            frm210137.txtCloseDate(1) = DBDATE(.TextMatrix(ii, 1)) - 19110000 '點數結算日(迄)
'            frm210137.cmdSearch_Click
'            If frm210137.grdDataList.Rows > 1 Then
'               If Trim(frm210137.grdDataList.TextMatrix(1, 5)) <> "" Then
'                  .TextMatrix(ii, 8) = Format(frm210137.grdDataList.TextMatrix(1, 5), "#.00")
'               End If
'            End If
             
            '已收文未收款點數
            'Add By Sindy 2020/10/5 顯示每日累計未收款點數
            If Check1.Visible = True And Check1.Value = 1 Then
               '.TextMatrix(ii, 9) = Format(PUB_CountCP18(1, "", DBDATE(.TextMatrix(ii, 1)) - 19110000, , .TextMatrix(ii, 11)), "#.00")
               If bolFirstRow = True Then
                  .TextMatrix(ii, 9) = Format(PUB_CountCP18(1, "", DBDATE(.TextMatrix(ii, 1)) - 19110000, , , .TextMatrix(ii, 11)), "#.00")
                  bolFirstRow = False
               Else
                  '每日小計
                  .TextMatrix(ii, 9) = Format(.TextMatrix(ii + 1, 9) + PUB_CountCP18(1, DBDATE(.TextMatrix(ii, 1)) - 19110000, DBDATE(.TextMatrix(ii, 1)) - 19110000, , , .TextMatrix(ii, 11)), "#.00")
               End If
            Else
            '2020/10/5 END
               .TextMatrix(ii, 9) = "-"
            End If
'            frm210141.Hide
'            frm210141.txtSales = .TextMatrix(ii, 11) '智權人員ID
'            frm210141.txtDate(0) = "" '點數結算日(起)
'            frm210141.txtDate(1) = DBDATE(.TextMatrix(ii, 1)) - 19110000 '點數結算日(迄)
'            frm210141.cmdok_Click (1)
'            If Trim(frm210141.txtTot(0)) <> "" Then
'               .TextMatrix(ii, 9) = Format(frm210141.txtTot(0), "#.00")
'            End If
            '2020/8/31 END
         End If
      Next ii
'      'Add By Sindy 2020/8/31
'      If PUB_IsFormExist("frm210137") = True Then
'         Unload frm210137
'      End If
'      If PUB_IsFormExist("frm210141") = True Then
'         Unload frm210141
'      End If
'      '2020/8/31 END
      .Visible = True
   End With
   
   Unload frmpic002 'Add By Sindy 2020/8/31
   Exit Sub
   
ErrHnd:
   'Resume
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'Added by Lydia 2021/07/27 每日點數輸入
Private Sub Command1_Click()
    If PUB_CheckFormExist("frm210103") Then
        MsgBox "請先關閉〔每日點數輸入〕畫面！"
        Exit Sub
    End If
    
    'Added by Lydia 2024/01/31
    If strUserNum = "75007" And strSrvDate(1) >= "20240202" Then
        MsgBox "無此使用權限...", , "警告!!"
    Else
    'end 2024/01/31
      Call frm210103.SetParent(Me)
      frm210103.Show
      Me.Hide
    End If
End Sub

'Added by Lydia 2021/07/27 業績點數統計
Private Sub Command2_Click()
    If PUB_CheckFormExist("frm210137") Then
        MsgBox "請先關閉〔業績點數統計〕畫面！"
        Exit Sub
    End If
    
    'Added by Lydia 2024/01/31
    If strUserNum = "75007" And strSrvDate(1) >= "20240202" Then
        MsgBox "無此使用權限...", , "警告!!"
    Else
    'end 2024/01/31
        Call frm210137.SetParent(Me)
        frm210137.Show
        Me.Hide
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
