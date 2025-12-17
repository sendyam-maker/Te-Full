VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc2170 
   AutoRedraw      =   -1  'True
   Caption         =   "結匯資料輸入 --> 帳單"
   ClientHeight    =   5230
   ClientLeft      =   50
   ClientTop       =   280
   ClientWidth     =   8810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5230
   ScaleWidth      =   8810
   Begin VB.ComboBox cboPrinters 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3810
      Style           =   2  '單純下拉式
      TabIndex        =   8
      Top             =   4740
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "其他"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2880
      TabIndex        =   6
      Top             =   4380
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Frmacc2170.frx":0000
      Left            =   3810
      List            =   "Frmacc2170.frx":0002
      TabIndex        =   7
      Top             =   4380
      Width           =   3375
   End
   Begin VB.CommandButton Command5 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8010
      Picture         =   "Frmacc2170.frx":0004
      Style           =   1  '圖片外觀
      TabIndex        =   2
      ToolTipText     =   "取消"
      Top             =   30
      Width           =   450
   End
   Begin VB.CommandButton Command4 
      Caption         =   "FC暫收款退費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1200
      TabIndex        =   5
      Top             =   4380
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "抵帳單"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   270
      TabIndex        =   4
      Top             =   4380
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2170.frx":066E
      Height          =   3825
      Left            =   240
      TabIndex        =   3
      Top             =   510
      Width           =   8295
      _ExtentX        =   14623
      _ExtentY        =   6738
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "ChkJ"
         Caption         =   "J"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "A1702"
         Caption         =   "帳單編號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "A1703"
         Caption         =   "幣別"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "A1704"
         Caption         =   "帳單金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "AXF14"
         Caption         =   "目前盈虧"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "A1707"
         Caption         =   "本所案號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "FagentName"
         Caption         =   "代理人名稱"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "A1705"
         Caption         =   "代理人"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "A1706"
         Caption         =   "代理人 D/N No."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   290.268
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1269.921
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   590.173
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1310.173
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2970.142
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1399.748
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1679.811
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "列印並產生付款單"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   7320
      TabIndex        =   9
      Top             =   4380
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   0
      Top             =   120
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2787
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   240
      Top             =   360
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "印表機"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3030
      TabIndex        =   12
      Top             =   4770
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "作業日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   -30
      Top             =   4680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "帳單編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc2170"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/06 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB; Printer列印未改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoacc150 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoacc170 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Dim adoaccrpt216 As New ADODB.Recordset
Dim adoaccrpt217 As New ADODB.Recordset
Dim adoaccrpt218 As New ADODB.Recordset
'Dim dllaccrpt As Object 'Mark by Lydia 2022/03/30
Dim m_iDefaultPrinter As String '預設印表機
'Added by Lydia 2016/08/15 列印用
Dim mPrtOrt As Integer  '原本預設印表機的列印方向
Private Const ciTitleFontSize = 14, cInX = 13
Private Const ciStartX = 400, ciStartY = 400, ciColGap = 150
Dim ciFontSize As Integer '報表內容字型大小
Dim mRptTitle As String '報表抬頭
Dim strTitle As String, strTitle2 As String '欄位抬頭/起始位置
Dim PLeft(0 To cInX) As Integer '欄位起始位置陣列
Dim PTitle(0 To cInX) As String '欄位抬頭陣列
Dim iPrint As Integer, iPage As Integer
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim iPageLine As Integer '頁面資料列
Private Const cntReport003 = "案件失誤明細表" 'Added by Lydia 2022/03/30

Private Sub Command1_Click()
Dim blnExcel As Boolean 'Add By Cheng 2004/03/09是否產生Excel檔案

   'Add by Sindy 2010/8/30
   'Modified by Lydia 2016/08/15
   'pub_OsPrinter = PUB_GetOsDefaultPrinter
   'PUB_SetOsDefaultPrinter cboPrinters
   ''2010/8/30 End
   PUB_RestorePrinter cboPrinters
   PUB_SetOsDefaultPrinter cboPrinters 'Added by Lydia 2022/04/25 變更OS印表機(Excel列印)
   mPrtOrt = Printer.Orientation
   Screen.MousePointer = vbHourglass
   'Modified by Lydia 2016/08/15 改成Printer輸出
   Select Case Combo1
      Case Mid(ReportTitle(216), 6, 8) '國內未收款明細表
           ProcessData '產生付款單
           Call PrintRpt2160
      Case Mid(ReportTitle(217), 6, 10) '代理人未收未付對照表
           Call PrintRpt2170
      'Added by Lydia 2022/03/30
      Case cntReport003
           Call PrintReport003
      'end 2022/03/30
      Case Mid(ReportTitle(218), 6, 6) '付款明細草稿
           Call PrintRpt2180
   End Select
   Screen.MousePointer = vbDefault
   Printer.Orientation = mPrtOrt
   PUB_RestorePrinter pub_OsPrinter
   PUB_SetOsDefaultPrinter pub_OsPrinter 'Added by Lydia 2022/04/25 還原OS印表機
   
   Exit Sub
   'end 2016/08/15
   
   'Remove by Lydia 2018/06/27
'   ProcessData
'   Select Case Combo1
'      Case Mid(ReportTitle(216), 6, 8) '國內未收款明細表
'         If adoaccrpt216.State = adStateOpen Then
'            adoaccrpt216.Close
'         End If
'         adoaccrpt216.CursorLocation = adUseClient
'         adoaccrpt216.Open "select * from accrpt216", adoTaie, adOpenStatic, adLockReadOnly
'         If adoaccrpt216.RecordCount <> 0 Then
'            dllaccrpt.Acc2160 ReportTitle(216), StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'         Else
'            MsgBox MsgText(28), , MsgText(5)
'         End If
'         adoaccrpt216.Close
'      Case Mid(ReportTitle(217), 6, 10) '代理人未收未付對照表
'         If adoaccrpt217.State = adStateOpen Then
'            adoaccrpt217.Close
'         End If
'         adoaccrpt217.CursorLocation = adUseClient
'         adoaccrpt217.Open "select * from accrpt217", adoTaie, adOpenStatic, adLockReadOnly
'         If adoaccrpt217.RecordCount <> 0 Then
'            dllaccrpt.Acc2170 ReportTitle(217), StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'         Else
'            MsgBox MsgText(28), , MsgText(5)
'         End If
'         adoaccrpt217.Close
'      Case Mid(ReportTitle(218), 6, 6) '付款明細草稿
'         If adoaccrpt218.State = adStateOpen Then
'            adoaccrpt218.Close
'         End If
'         adoaccrpt218.CursorLocation = adUseClient
'         'Added by Lydia 2015/03/18 +鎖定使用者 2015/9/14 cancel by soina 因為accreport無法抓使用者
'         'Added by Lydia 2015/03/31 + 台一備註
'         'modify by sonia 2015/9/14 沒有以代理人及幣別去抓acc220會很久
'         'adoaccrpt218.Open "select a.*,substrb(b.a2223,1,30) T1MEMO from accrpt218 a ,acc220 b where R21801='" & strUserNum & "' ", adoTaie, adOpenStatic, adLockReadOnly
'         adoaccrpt218.Open "select a.*,substrb(b.a2223,1,30) T1MEMO from accrpt218 a ,acc220 b where r21803=a2201(+) and r21805=a2202(+)", adoTaie, adOpenStatic, adLockReadOnly
'         If adoaccrpt218.RecordCount <> 0 Then
'            dllaccrpt.Acc2180 ReportTitle(218), StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'            blnExcel = True
'         Else
'            MsgBox MsgText(28), , MsgText(5)
'            blnExcel = False
'         End If
'         adoaccrpt218.Close
'        'Modified by Lydia 2015/03/04 改在frmacc24m0
''        If blnExcel = True Then
''           'Modified by Lydia 2015/02/25 結匯明細匯總表-依幣別,不分公司
''           ' PUB_ExcelSave '結匯明細匯總表
''            PUB_ExcelSave2
''            MsgBox "Excel檔案產生完成!!!", vbExclamation + vbOKOnly
''        End If
'   End Select
'   Screen.MousePointer = vbDefault
'   'Add by Sindy 2010/8/30
'   PUB_SetOsDefaultPrinter pub_OsPrinter
    'end 2018/06/27
   'Remove by Lydia 2016/08/03
  ' cboPrinters = pub_OsPrinter
   '2010/8/30 End
End Sub

Private Sub Command2_Click()
   strExitControl = MsgText(601)
   tool3_enabled
   Frmacc2180.Show
   Unload Me
End Sub

Private Sub Command3_Click()
   strExitControl = MsgText(601)
   tool3_enabled
   Frmacc2191.Show
   Unload Me
End Sub

Private Sub Command4_Click()
   strExitControl = MsgText(601)
   tool3_enabled
   Frmacc2190.Show
   Unload Me
End Sub

Private Sub Command5_Click()
   AdodcDelete
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   'Modified by Lydia 2021/12/07 改成模組
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   strExitControl = MsgText(602) 'Add by Morgan 2006/7/17
'   Me.Width = 8850
'   Me.Height = 5500
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath1)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   strFormName = Name
   strExitControl = MsgText(602)
   'Modify by Amy 2023/08/18 W8880 H5550
   PUB_InitForm Me, 8900, 5680, strBackPicPath1
   'end 2021/12/07
   MaskEdBox1.Mask = MsgText(601)
   MaskEdBox1.Text = CFDate(ACDate(ServerDate))
   MaskEdBox1.Mask = DFormat
   Combo1.AddItem Mid(ReportTitle(216), 6, 8)  ''國內未收款明細表
   Combo1.AddItem Mid(ReportTitle(217), 6, 10) '代理人未收未付對照表
   Combo1.AddItem cntReport003 'Added by Lydia 2022/03/30
   Combo1.AddItem Mid(ReportTitle(218), 6, 6) '付款明細草稿
   Combo1 = Mid(ReportTitle(216), 6, 8)
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133) & ", " & MsgText(134)
   'Set dllaccrpt = CreateObject("AccReport.ReportSelect") 'Mark by Lydia 2022/03/30
    'Add By Cheng 2003/05/14
    Text1.Text = "U"
   'Modified by Lydia 2016/08/03 改成共用模組
   ' AddPrinter 'Add by Sindy 2010/8/30 加印表機選擇
   PUB_SetPrinter Me.Name, cboPrinters, pub_OsPrinter
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
   
   If strExitControl = MsgText(602) Then
      strFormName = MsgText(601)
      KeyEnter vbKeyEscape
      MenuEnabled
      'Set dllaccrpt = Nothing 'Mark by Lydia 2022/03/30
      Set Frmacc2170 = Nothing
   End If
   strExitControl = MsgText(602)
   
   'Added by Lydia 2016/08/03
   '若印表機變動, 則更新列印設定
   If cboPrinters.Text <> cboPrinters.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cboPrinters.Name, "0", "0", Me.cboPrinters.Text
   End If
   'PUB_RestorePrinter pub_OsPrinter 'Remove by Lydia 2016/08/15
   
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox Label2 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label2 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133) & ", " & MsgText(134)
End Sub

Private Sub Text1_GotFocus()
   'MODIFY BY SONIA 2014/3/19
   'TextInverse Text1
   If Len(Text1) > 0 Then
      Text1.SelStart = 1
      Text1.SelLength = Len(Text1) - 1
   End If
   '2014/3/19 END
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
    'Modify By Cheng 2003/06/02
'   adoadodc1.Open "select a1702, a1703, axf04 as a1704, a1705, a1706, axf03 as a1707, nvl(fa05||fa63||fa64||fa65,nvl(fa04, fa06)) FagentName from acc151, acc170, fagent where axf01 = a1702 and substr(a1705, 1, 8) = fa01 (+) and substr(a1705, 9, 1) = fa02 (+) and a1701 = '1' and (a1709 is null or a1709 = '') order by a1702 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
    'Modify By Cheng 2003/06/05
'   adoadodc1.Open "select a1702, a1703, axf04 as a1704, a1705, a1706, axf03 as a1707, nvl(fa05||fa63||fa64||fa65,nvl(fa04, fa06)), axf14 FagentName from acc151, acc170, fagent where axf01 = a1702 and substr(a1705, 1, 8) = fa01 (+) and substr(a1705, 9, 1) = fa02 (+) and a1701 = '1' and (a1709 is null or a1709 = '') order by a1702 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modified by Lydia 2016/08/15 加註ChkJ="智"
   'adoadodc1.Open "select a1702, a1703, axf04 as a1704, a1705, a1706, axf03 as a1707, nvl(fa05||fa63||fa64||fa65,nvl(fa04, fa06)) As FagentName, axf14 from acc151, acc170, fagent where axf01 = a1702 and substr(a1705, 1, 8) = fa01 (+) and substr(a1705, 9, 1) = fa02 (+) and a1701 = '1' and (a1709 is null or a1709 = '') order by a1702 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modified by Morgan 2020/10/23 +L公司顯示"法"--婉莘(只有L案,其他還是用智慧所)
   strSql = "select decode(nvl(pa161,nvl(tm130,nvl(lc48,sp85))),'J','智',decode(lc01,'L','法')) ChkJ,a1702, a1703, axf04 as a1704, a1705, a1706, axf03 as a1707, nvl(fa05||fa63||fa64||fa65,nvl(fa04, fa06)) As FagentName, axf14" & _
            " from acc151, acc170, fagent,caseprogress,patent,trademark,lawcase,servicepractice" & _
            " where axf01 = a1702 and substr(a1705, 1, 8) = fa01 (+) and substr(a1705, 9, 1) = fa02 (+) and a1701 = '1' and (a1709 is null or a1709 = '')" & _
            " and axf02=cp09 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & _
            " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)" & _
            " and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+)" & _
            " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) order by a1702 asc"
   adoadodc1.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   'adoadodc1.Open "select * from acc170 where a1701 = '1' and a1708 = " & Val(FCDate(MaskEdBox1.Text)) & " order by a1702 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
    'Modify By Cheng 2003/06/02
'   adoadodc1.Open "select a1702, a1703, axf04 as a1704, a1705, a1706, axf03 as a1707, nvl(fa05||fa63||fa64||fa65,nvl(fa04, fa06)) FagentName from acc151, acc170, fagent where axf01 = a1702 and substr(a1705, 1, 8) = fa01 (+) and substr(a1705, 9, 1) = fa02 (+) and a1701 = '1' and (a1709 is null or a1709 = '')  order by a1702 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
    'Modify By Cheng 2003/06/05
'   adoadodc1.Open "select a1702, a1703, axf04 as a1704, a1705, a1706, axf03 as a1707, nvl(fa05||fa63||fa64||fa65,nvl(fa04, fa06)), axf14 FagentName from acc151, acc170, fagent where axf01 = a1702 and substr(a1705, 1, 8) = fa01 (+) and substr(a1705, 9, 1) = fa02 (+) and a1701 = '1' and (a1709 is null or a1709 = '')  order by a1702 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modified by Lydia 2016/08/15 加註ChkJ="智"
   'adoadodc1.Open "select a1702, a1703, axf04 as a1704, a1705, a1706, axf03 as a1707, nvl(fa05||fa63||fa64||fa65,nvl(fa04, fa06)) As FagentName, axf14  from acc151, acc170, fagent where axf01 = a1702 and substr(a1705, 1, 8) = fa01 (+) and substr(a1705, 9, 1) = fa02 (+) and a1701 = '1' and (a1709 is null or a1709 = '')  order by a1702 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modified by Morgan 2020/10/23 +L公司顯示"法"--婉莘(只有L案,其他還是用智慧所)
   strSql = "select decode(nvl(pa161,nvl(tm130,nvl(lc48,sp85))),'J','智',decode(lc01,'L','法')) ChkJ,a1702, a1703, axf04 as a1704, a1705, a1706, axf03 as a1707, nvl(fa05||fa63||fa64||fa65,nvl(fa04, fa06)) As FagentName, axf14" & _
            " from acc151, acc170, fagent,caseprogress,patent,trademark,lawcase,servicepractice" & _
            " where axf01 = a1702 and substr(a1705, 1, 8) = fa01 (+) and substr(a1705, 9, 1) = fa02 (+) and a1701 = '1' and (a1709 is null or a1709 = '')" & _
            " and axf02=cp09 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & _
            " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)" & _
            " and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+)" & _
            " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) order by a1702 asc"
   adoadodc1.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.Find "a1702 = '" & Text1 & "'", 0, adSearchForward, 1
      If Adodc1.Recordset.EOF Then
         Exit Sub
      Else
         DataGrid1.SelBookmarks.add Adodc1.Recordset.Bookmark
      End If
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Added by Morgan 2017/1/24 改 Acc170Save 為此共用函數(Frmacc2171也要用)
Public Function Acc170SaveNew(pBillNo As String, Optional pRefreshGrid As Boolean, Optional pEBill As Boolean = False) As Boolean
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset, rsQuery2 As ADODB.Recordset
   
   Dim stra1703 As String
   Dim stra1704 As String
   Dim stra1705 As String
   Dim stra1706 As String
   Dim stra1707 As String
   Dim stra1708 As String
   
   '檢查進度檔是否有此帳單號
   stSQL = "select cp61, cp62, cp63,cp87,cp88 from caseprogress where cp61 = '" & pBillNo & "' or cp62 = '" & pBillNo & "' or cp63 = '" & pBillNo & "' or cp87 = '" & pBillNo & "' or cp88 = '" & pBillNo & "'"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 0 Then
      MsgBox MsgText(208), , MsgText(5)
      GoTo ExitFunction
   End If
   
   '檢查是否已有此帳單的結匯資料
   stSQL = "select * from acc170 where a1701 = '1' and a1702 = '" & pBillNo & "'"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR <> 0 Then
      MsgBox MsgText(9), , MsgText(5)
      pRefreshGrid = True
      GoTo ExitFunction
   End If
   
   stSQL = "select a1.*,a2.axf01,a2.axf03 from acc150 a1, acc151 a2 where a1501 = '" & pBillNo & "' and a1501=axf01(+) "
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   With rsQuery
   If intR = 0 Then
      MsgBox MsgText(28), , MsgText(5)
      GoTo ExitFunction
      
   '審核
   'Modified by Morgan 2023/12/6 +R退回
   ElseIf .Fields("a1521") = "N" Or .Fields("a1521") = "W" Or .Fields("a1521") = "R" Then
      MsgBox MsgText(190), , MsgText(5)
      GoTo ExitFunction
      
   ElseIf Not IsNull(.Fields("a1512")) Then
      MsgBox "該帳單已抵帳！"
      GoTo ExitFunction
   Else
      'Added by Morgan 2024/3/21
      'Y55766德國專利局帳單防呆檢查 1.不可有電子檔 2.金額非整數提醒可繼續
      If PUB_Y55766BillCheck(pBillNo, .Fields("a1503"), .Fields("a1506")) = False Then
         GoTo ExitFunction
      End If
      'end 2024/3/21
   End If
      
   '智權年底結匯檢查
   If Val(Right(strSrvDate(1), 4)) >= 1001 And Val(Right(strSrvDate(1), 4)) <= 1231 Then
      .MoveFirst
      Do While Not .EOF
        Select Case Left(.Fields("AXF03"), Len(.Fields("AXF03")) - 9)
           Case "CFP", "P"
              stSQL = "SELECT AXF01,AXF03,AXF02,PA161 VC1,DECODE(SUBSTR(CP60,1,1),'E',DECODE(SIGN(CP79),1,'Y',''),'X',DECODE(A1K29,NULL,'Y',''),'') CHK FROM ACC151 A1, PATENT P1,CASEPROGRESS C1,ACC1K0 A2 " & _
                       "WHERE AXF01 = '" & .Fields("AXF01") & "' AND SUBSTR(AXF03, 1, LENGTH(AXF03) - 9)=P1.PA01(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 8, 6)=P1.PA02(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 2, 1)=P1.PA03(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 1, 2)=P1.PA04(+) " & _
                       "AND PA161='J' AND AXF02=CP09(+) AND CP60=A1K01(+) "
               stSQL = stSQL & " and not exists(select * from acc150 where a1501=axf01 and a1503='Y49572000')" 'Added by Morgan 2020/10/5 Y49572 USPTO的帳單是以刷卡方式支付, 是已經支付的帳單, 不適用此規則--婉莘
               
           Case "TF", "T", "CFT"
              stSQL = "SELECT AXF01,AXF03,AXF02,TM130 VC1,DECODE(SUBSTR(CP60,1,1),'E',DECODE(SIGN(CP79),1,'Y',''),'X',DECODE(A1K29,NULL,'Y',''),'') CHK FROM ACC151 A1, TRADEMARK P1,CASEPROGRESS C1,ACC1K0 A2 " & _
                       "WHERE AXF01 = '" & .Fields("AXF01") & "' AND SUBSTR(AXF03, 1, LENGTH(AXF03) - 9)=P1.TM01(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 8, 6)=P1.TM02(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 2, 1)=P1.TM03(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 1, 2)=P1.TM04(+) " & _
                       "AND TM130='J' AND AXF02=CP09(+) AND CP60=A1K01(+) "
           Case "L", "CFL"
              stSQL = "SELECT AXF01,AXF03,AXF02,LC48 VC1,DECODE(SUBSTR(CP60,1,1),'E',DECODE(SIGN(CP79),1,'Y',''),'X',DECODE(A1K29,NULL,'Y',''),'') CHK FROM ACC151 A1, LAWCASE P1,CASEPROGRESS C1,ACC1K0 A2 " & _
                       "WHERE AXF01 = '" & .Fields("AXF01") & "' AND SUBSTR(AXF03, 1, LENGTH(AXF03) - 9)=P1.LC01(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 8, 6)=P1.LC02(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 2, 1)=P1.LC03(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 1, 2)=P1.LC04(+) " & _
                       "AND LC48='J' AND AXF02=CP09(+) AND CP60=A1K01(+) "
           Case Else
              stSQL = "SELECT AXF01,AXF03,AXF02,SP85 VC1,DECODE(SUBSTR(CP60,1,1),'E',DECODE(SIGN(CP79),1,'Y',''),'X',DECODE(A1K29,NULL,'Y',''),'') CHK FROM ACC151 A1, SERVICEPRACTICE P1,CASEPROGRESS C1,ACC1K0 A2 " & _
                       "WHERE AXF01 = '" & .Fields("AXF01") & "' AND SUBSTR(AXF03, 1, LENGTH(AXF03) - 9)=P1.SP01(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 8, 6)=P1.SP02(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 2, 1)=P1.SP03(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 1, 2)=P1.SP04(+) " & _
                       "AND SP85='J' AND AXF02=CP09(+) AND CP60=A1K01(+) "
        End Select
        intR = 1
        Set rsQuery2 = ClsLawReadRstMsg(intR, stSQL)
        If intR = 1 Then
           If "" & rsQuery2.Fields("CHK") = "Y" Then
              'Modified by Morgan 2020/12/3 改用詢問方式--婉莘
              'MsgBox "此款項尚未收齊,不可結匯!"
              'GoTo ExitFunction
              If MsgBox("此款項尚未收齊,是否確定要結匯??", vbYesNo + vbDefaultButton2) = vbYes Then
               Exit Do
              Else
                  GoTo ExitFunction
              End If
              'end 2020/12/3
           End If
        End If
         .MoveNext
      Loop
   End If
   
   .MoveFirst
   'Add by Amy 2013/11/18 +帳款處理訊息
   GetDizhang "" & .Fields("a1503"), , True
   'end 2013/11/18
   
   stra1703 = "" & .Fields("a1505").Value
   stra1704 = Val("" & .Fields("a1506").Value)
   stra1705 = "" & .Fields("a1503").Value
      
   '*****注意, 此處改, acc_fun的GetTermOfPayment也要改
   '2009/9/23 add by sonia 付款對象固定改為另一家,財務說不限CFT案,CFP也要
   Select Case Mid(stra1705, 1, 6)
      '2010/2/6 modify by sonia 婧瑄說加Y20915(U09900096)
      '2012/7/9 modify by sonia 婧瑄說加Y20908
      'modify by sonia 2017/8/1  陳經理再加Y54715'2014/11/24 modify by sonia 婉莘再加Y54053
      'modify by sonia 2016/4/28 陳經理再加Y54052
      'modify by sonia 2017/8/1  陳經理再加Y54715
      'modify by sonia 2020/1/16 陳經理再加Y55351
      'modify by sonia 2024/8/12 再加Y56014
      'modify by sonia 2025/3/4  再加Y56137
      'modify by sonia 2025/8/29 再加Y56167000及Y56167B10
      Case "Y20908", "Y20915", "Y20919", "Y20929", "Y20934", "Y34282", "Y22247", "Y30249", "Y51368", "Y51523", "Y52243", "Y20076", "Y20339", "Y54053", "Y54052", "Y54715", "Y55351", "Y56014", "Y56137", "Y56167"
         stra1705 = "Y20076000"          '2009/9/23 add by sonia 外商陳經理提出中東地區Abu-Ghazaleh代理人之付款對象改為Y20076,財務說不限CFT案,CFP也要
      '2010/11/19 add by sonia 陳經理說再加一組
      '2010/12/6 modify by sonia 加Y53117--Y53119,Y53122-3
      '2011/4/11 modify by sonia 婧瑄說加Y53121
      '2012/5/4  modify by sonia 加Y51352
      Case "Y45778", "Y53120", "Y20917", "Y53117", "Y53118", "Y53119", "Y53122", "Y53123", "Y53121", "Y51352"
         stra1705 = "Y20917000"
      '2010/11/22 add by sonia 陳經理說再加一組
      Case "Y49419", "Y53188"
         stra1705 = "Y53188000"
      'add by sonia 2018/1/26
      Case Else
         If stra1705 = "Y20026020" Then stra1705 = "Y20026000"
         'If stra1705 = "Y20284000" Then stra1705 = "Y20284010"   'add by sonia 2018/5/21
         If stra1705 = "Y45878000" Then stra1705 = "Y55253000"   'add by sonia 2019/5/28 婉莘
         If stra1705 = "Y51333010" Then stra1705 = "Y51333000"   'add by sonia 2019/6/24 婉莘
         If stra1705 = "Y52754020" Then stra1705 = "Y52754010"   'add by sonia 2022/1/20 婉莘
      'end 2018/1/26
   End Select
   '2009/9/23 end
   
   stra1706 = "" & .Fields("a1504").Value
   stra1707 = "" & .Fields("axf03").Value
   stra1708 = strSrvDate(2)
   End With
   'Modified by Morgan 2018/1/12 +a1719
   adoTaie.Execute "insert into acc170 (a1701, a1702, a1703, a1704, a1705, a1706, a1707, a1708, a1710, a1711, a1712, a1719) values ('1', '" & pBillNo & "', " & CNULL(stra1703) & ", " & CNULL(stra1704) & ", " & CNULL(stra1705) & ", " & CNULL(stra1706) & ", " & CNULL(stra1707) & ", " & CNULL(stra1708) & ", " & strSrvDate(2) & ", to_char(sysdate, 'HH24MISS'), '" & strUserNum & "','" & IIf(pEBill, "Y", "") & "')"
   pRefreshGrid = True
   Acc170SaveNew = True
   
   'add by sonia 2014/7/21 J公司或獨立水單都要彈訊息提醒
   Select Case Left(stra1707, Len(stra1707) - 9)
      Case "CFP", "P"
         stSQL = "select nvl(p1.pa26,p2.pa26) pa26,nvl(p1.pa161,p2.pa161) pa161 from acc170, acc151, patent p1, acc161, patent p2 where a1702='" & pBillNo & "' and a1702=axf01(+) and a1702=axg01(+) " & _
                   "and substr(axf03, 1, length(axf03) - 9)=p1.pa01(+) and substr(axf03, length(axf03) - 8, 6)=p1.pa02(+) and substr(axf03, length(axf03) - 2, 1)=p1.pa03(+) and substr(axf03, length(axf03) - 1, 2)=p1.pa04(+) " & _
                   "and substr(axg03, 1, length(axg03) - 9)=p2.pa01(+) and substr(axg03, length(axg03) - 8, 6)=p2.pa02(+) and substr(axg03, length(axg03) - 2, 1)=p2.pa03(+) and substr(axg03, length(axg03) - 1, 2)=p2.pa04(+) "
      Case "TF", "T", "CFT"
         stSQL = "select nvl(t1.tm23,t2.tm23) pa26,nvl(t1.tm130,t2.tm130) pa161 from acc170, acc151, trademark t1, acc161, trademark t2 where a1702='" & pBillNo & "' and a1702=axf01(+) and a1702=axg01(+) " & _
                   "and substr(axf03, 1, length(axf03) - 9)=t1.tm01(+) and substr(axf03, length(axf03) - 8, 6)=t1.tm02(+) and substr(axf03, length(axf03) - 2, 1)=t1.tm03(+) and substr(axf03, length(axf03) - 1, 2)=t1.tm04(+) " & _
                   "and substr(axg03, 1, length(axg03) - 9)=t2.tm01(+) and substr(axg03, length(axg03) - 8, 6)=t2.tm02(+) and substr(axg03, length(axg03) - 2, 1)=t2.tm03(+) and substr(axg03, length(axg03) - 1, 2)=t2.tm04(+) "
      Case "L", "CFL"
         stSQL = "select nvl(L1.LC11,L2.LC11) pa26,nvl(L1.lc48,L2.lc48) pa161 from acc170, acc151, LAWCASE L1, acc161, LAWCASE L2 where a1702='" & pBillNo & "' and a1702=axf01(+) and a1702=axg01(+) " & _
                   "and substr(axf03, 1, length(axf03) - 9)=L1.LC01(+) and substr(axf03, length(axf03) - 8, 6)=L1.LC02(+) and substr(axf03, length(axf03) - 2, 1)=L1.LC03(+) and substr(axf03, length(axf03) - 1, 2)=L1.LC04(+) " & _
                   "and substr(axg03, 1, length(axg03) - 9)=L2.LC01(+) and substr(axg03, length(axg03) - 8, 6)=L2.LC02(+) and substr(axg03, length(axg03) - 2, 1)=L2.LC03(+) and substr(axg03, length(axg03) - 1, 2)=L2.LC04(+) "
      Case Else
         stSQL = "select nvl(s1.sp08,s2.sp08) pa26,nvl(s1.sp85,s2.sp85) pa161 from acc170, acc151, servicepractice s1, acc161, servicepractice s2 where a1702='" & pBillNo & "' and a1702=axf01(+) and a1702=axg01(+) " & _
                   "and substr(axf03, 1, length(axf03) - 9)=s1.sp01(+) and substr(axf03, length(axf03) - 8, 6)=s1.sp02(+) and substr(axf03, length(axf03) - 2, 1)=s1.sp03(+) and substr(axf03, length(axf03) - 1, 2)=s1.sp04(+) " & _
                   "and substr(axg03, 1, length(axg03) - 9)=s2.sp01(+) and substr(axg03, length(axg03) - 8, 6)=s2.sp02(+) and substr(axg03, length(axg03) - 2, 1)=s2.sp03(+) and substr(axg03, length(axg03) - 1, 2)=s2.sp04(+) "
   End Select
   
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      With rsQuery
      If IsNull(.Fields("pa26").Value) = False Then
      
'Modified by Morgan 2019/7/16 改用函數判斷
'         'modify by sonia 2017/4/10 婉莘要求加X69534彩豐精技   2017/5/15陳德發及郭雅娟要求取消
'         'modify by sonia 2017/4/27 高國碩要求X60149改為只要母號,關係企業不要
'         If Left(.Fields("pa26").Value, 6) = "X44551" Or Left(.Fields("pa26").Value, 6) = "X62079" _
'         Or Left(.Fields("pa26").Value, 6) = "X43988" Or Left(.Fields("pa26").Value, 6) = "X63219" Or Left(.Fields("pa26").Value, 6) = "X62319" _
'         Or Left(.Fields("pa26").Value, 6) = "X60498" Or Left(.Fields("pa26").Value, 6) = "X62702" Or Left(.Fields("pa26").Value, 6) = "X63838" _
'          Then
'            MsgBox "此為特定客戶, 水單不能合併！", , MsgText(5)
'         End If
'         'add by sonia 2017/4/14 單獨編號而關係企業不要的,例:王副總的華碩客戶(X69011010)
'         'modify by sonia 2017/4/27 高國碩要求X60149改為只要母號,關係企業不要
'         'modify by sonia 2017/10/11 張詠翔要求取消X6014900
'         'modify by sonia 2017/10/26 郭雅娟要求加X60738010國立清華大學
'         If Left(.Fields("pa26").Value, 8) = "X6901101" Or Left(.Fields("pa26").Value, 8) = "X6073801" Then
'            MsgBox "此為特定客戶, 水單不能合併！", , MsgText(5)
'         End If
'         'end 2017/4/14

         If PUB_ChkNoMergePayCust("", .Fields("pa26")) = True Then
            MsgBox "此為特定客戶, 水單不能合併！", , MsgText(5)
         End If
'end 2019/7/16

      End If
      'Modified by Morgan 2017/6/26 電子結匯畫面已有註記不用再提醒 -- 婉莘
      If "" & .Fields("pa161").Value = "J" And pEBill = False Then
         MsgBox "此為智權公司出名案件！", , MsgText(5)
      End If
      End With
   End If
   'end 2014/7/21
   
ExitFunction:
   
   If Not rsQuery Is Nothing Then
      If rsQuery.State = adStateOpen Then rsQuery.Close
      Set rsQuery = Nothing
   End If
   
   If Not rsQuery2 Is Nothing Then
      If rsQuery2.State = adStateOpen Then rsQuery2.Close
      Set rsQuery2 = Nothing
   End If
End Function

'*************************************************
'  儲存資料表(國外結匯資料)
'
'*************************************************
'Removed by Morgan 2019/7/16 沒用,上註解以免又改到
'Private Sub Acc170Save()
'Dim stra1703 As String
'Dim stra1704 As String
'Dim stra1705 As String
'Dim stra1706 As String
'Dim stra1707 As String
'Dim stra1708 As String
'Dim StrSQLa As String
'
'On Error GoTo Checking
'
'   If adoquery.State = adStateOpen Then
'      adoquery.Close
'   End If
'
'   'Modify by Morgan 2009/3/10 若已抵帳時也要提醒
'   adoquery.CursorLocation = adUseClient
'   'adoquery.Open "select a1521 from acc150 where a1521 = '" & MsgText(603) & "' and a1501 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
'   'If adoquery.RecordCount <> 0 Then
'   '   MsgBox MsgText(190), , MsgText(5)
'   '   Exit Sub
'   'End If
'   'Modified by Lydia 2015/11/27 抓ACC151         判斷案件
'   'adoquery.Open "select * from acc150 where a1501 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
'   adoquery.Open "select a1.*,a2.axf01,a2.axf03 from acc150 a1, acc151 a2 where a1501 = '" & Text1 & "' and a1501=axf01(+) ", adoTaie, adOpenStatic, adLockReadOnly
'   If adoquery.RecordCount = 0 Then
'      MsgBox MsgText(28), , MsgText(5)
'      Exit Sub
'   'Modified by Morgan 2016/7/29 增加W(待審核)
'   'ElseIf adoquery.Fields("a1521") = "N" Then
'   ElseIf adoquery.Fields("a1521") = "N" Or adoquery.Fields("a1521") = "W" Then
'   'end 2016/7/29
'      MsgBox MsgText(190), , MsgText(5)
'      Exit Sub
'   ElseIf Not IsNull(adoquery.Fields("a1512")) Then
'      MsgBox "該帳單已抵帳！"
'      Exit Sub
'   End If
'   'end 2009/3/10
'
'   'Added by Lydia 2015/11/27 智權年底結匯檢查
'   If Val(Right(strSrvDate(1), 4)) >= 1001 And Val(Right(strSrvDate(1), 4)) <= 1231 Then
'      adoquery.MoveFirst
'      Do While Not adoquery.EOF
'        Select Case Left(adoquery.Fields("AXF03"), Len(adoquery.Fields("AXF03")) - 9)
'           Case "CFP", "P"
'              strSql = "SELECT AXF01,AXF03,AXF02,PA161 VC1,DECODE(SUBSTR(CP60,1,1),'E',DECODE(SIGN(CP79),1,'Y',''),'X',DECODE(A1K29,NULL,'Y',''),'') CHK FROM ACC151 A1, PATENT P1,CASEPROGRESS C1,ACC1K0 A2 " & _
'                       "WHERE AXF01 = '" & adoquery.Fields("AXF01") & "' AND SUBSTR(AXF03, 1, LENGTH(AXF03) - 9)=P1.PA01(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 8, 6)=P1.PA02(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 2, 1)=P1.PA03(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 1, 2)=P1.PA04(+) " & _
'                       "AND PA161='J' AND AXF02=CP09(+) AND CP60=A1K01(+) "
'           Case "TF", "T", "CFT"
'              strSql = "SELECT AXF01,AXF03,AXF02,TM130 VC1,DECODE(SUBSTR(CP60,1,1),'E',DECODE(SIGN(CP79),1,'Y',''),'X',DECODE(A1K29,NULL,'Y',''),'') CHK FROM ACC151 A1, TRADEMARK P1,CASEPROGRESS C1,ACC1K0 A2 " & _
'                       "WHERE AXF01 = '" & adoquery.Fields("AXF01") & "' AND SUBSTR(AXF03, 1, LENGTH(AXF03) - 9)=P1.TM01(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 8, 6)=P1.TM02(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 2, 1)=P1.TM03(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 1, 2)=P1.TM04(+) " & _
'                       "AND TM130='J' AND AXF02=CP09(+) AND CP60=A1K01(+) "
'           Case "L", "CFL"
'              strSql = "SELECT AXF01,AXF03,AXF02,LC48 VC1,DECODE(SUBSTR(CP60,1,1),'E',DECODE(SIGN(CP79),1,'Y',''),'X',DECODE(A1K29,NULL,'Y',''),'') CHK FROM ACC151 A1, LAWCASE P1,CASEPROGRESS C1,ACC1K0 A2 " & _
'                       "WHERE AXF01 = '" & adoquery.Fields("AXF01") & "' AND SUBSTR(AXF03, 1, LENGTH(AXF03) - 9)=P1.LC01(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 8, 6)=P1.LC02(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 2, 1)=P1.LC03(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 1, 2)=P1.LC04(+) " & _
'                       "AND LC48='J' AND AXF02=CP09(+) AND CP60=A1K01(+) "
'           Case Else
'              strSql = "SELECT AXF01,AXF03,AXF02,SP85 VC1,DECODE(SUBSTR(CP60,1,1),'E',DECODE(SIGN(CP79),1,'Y',''),'X',DECODE(A1K29,NULL,'Y',''),'') CHK FROM ACC151 A1, SERVICEPRACTICE P1,CASEPROGRESS C1,ACC1K0 A2 " & _
'                       "WHERE AXF01 = '" & adoquery.Fields("AXF01") & "' AND SUBSTR(AXF03, 1, LENGTH(AXF03) - 9)=P1.SP01(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 8, 6)=P1.SP02(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 2, 1)=P1.SP03(+) AND SUBSTR(AXF03, LENGTH(AXF03) - 1, 2)=P1.SP04(+) " & _
'                       "AND SP85='J' AND AXF02=CP09(+) AND CP60=A1K01(+) "
'        End Select
'        intI = 1
'        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'        If intI = 1 Then
'           If "" & RsTemp.Fields("CHK") = "Y" Then
'              MsgBox "此款項尚未收齊,不可結匯!"
'              Exit Sub
'           End If
'        End If
'         adoquery.MoveNext
'      Loop
'   End If
'   'end 2015/11/27
'
'   adoquery.Close
'
'   adoquery.CursorLocation = adUseClient
'   '2007/3/2 modify by sonia
'   'adoquery.Open "select cp61, cp62, cp63 from caseprogress where cp61 = '" & Text1 & "' or cp62 = '" & Text1 & "' or cp63 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
'   adoquery.Open "select cp61, cp62, cp63,cp87,cp88 from caseprogress where cp61 = '" & Text1 & "' or cp62 = '" & Text1 & "' or cp63 = '" & Text1 & "' or cp87 = '" & Text1 & "' or cp88 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
'   '2007/3/2 end
'   If adoquery.RecordCount = 0 Then
'      MsgBox MsgText(208), , MsgText(5)
'      Exit Sub
'   End If
'   adoquery.Close
'
'   adoacc150.CursorLocation = adUseClient
'   adoacc150.Open "select * from acc151, acc150 where acc151.axf01 = acc150.a1501 and a1507 is null and axf01 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoacc150.RecordCount <> 0 Then
'      adoacc150.MoveFirst
'      adoacc170.CursorLocation = adUseClient
'      'Modify by Morgan 2005/4/20
'      'adoacc170.Open "select * from acc170 where a1701 = '1' order by a1702 asc", adoTaie, adOpenStatic, adLockReadOnly
'      adoacc170.Open "select * from acc170 where a1701 = '1' and a1702 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adoacc170.RecordCount <> 0 Then
'         'Modify by Morgan 2005/4/20
'         'adoacc170.Find "a1702 = '" & Text1 & "'", 0, adSearchForward, 1
'         'If adoacc170.EOF = False Then
'            adoacc150.Close
'            adoacc170.Close
'            MsgBox MsgText(9), , MsgText(5)
'            AdodcRefresh
'            Exit Sub
'         'End If
'      End If
'      adoacc170.Close
'      'Adodc1.Recordset.AddNew
'      'Adodc1.Recordset.Fields("a1701").Value = "1"
'      'Adodc1.Recordset.Fields("a1702").Value = Text1
'      If IsNull(adoacc150.Fields("a1505").Value) Then
'         stra1703 = ""
'      Else
'         stra1703 = adoacc150.Fields("a1505").Value
'      End If
'      If IsNull(adoacc150.Fields("a1506").Value) Then
'         stra1704 = "0"
'      Else
'         stra1704 = adoacc150.Fields("a1506").Value
'      End If
'      If IsNull(adoacc150.Fields("a1503").Value) Then
'         stra1705 = ""
'      Else
'         stra1705 = adoacc150.Fields("a1503").Value
'      End If
'
'      '*****注意, 此處改, acc_fun的GetTermOfPayment也要改
'      '2009/9/23 add by sonia 付款對象固定改為另一家,財務說不限CFT案,CFP也要
'      Select Case Mid(stra1705, 1, 6)
'         '2010/2/6 modify by sonia 婧瑄說加Y20915(U09900096)
'         '2012/7/9 modify by sonia 婧瑄說加Y20908
'         '2014/11/24 modify by sonia 婉莘再加Y54053
'         'modify by sonia 2016/4/28 陳經理再加Y54052
'         'modify by sonia 2017/8/1  陳經理再加Y54715
'         Case "Y20908", "Y20915", "Y20919", "Y20929", "Y20934", "Y34282", "Y22247", "Y30249", "Y51368", "Y51523", "Y52243", "Y20076", "Y20339", "Y54053", "Y54052", "Y54715"
'            stra1705 = "Y20076000"          '2009/9/23 add by sonia 外商陳經理提出中東地區Abu-Ghazaleh代理人之付款對象改為Y20076,財務說不限CFT案,CFP也要
'         '2010/11/19 add by sonia 陳經理說再加一組
'         '2010/12/6 modify by sonia 加Y53117--Y53119,Y53122-3
'         '2011/4/11 modify by sonia 婧瑄說加Y53121
'         '2012/5/4  modify by sonia 加Y51352
'         Case "Y45778", "Y53120", "Y20917", "Y53117", "Y53118", "Y53119", "Y53122", "Y53123", "Y53121", "Y51352"
'            stra1705 = "Y20917000"
'         '2010/11/22 add by sonia 陳經理說再加一組
'         Case "Y49419", "Y53188"
'            stra1705 = "Y53188000"
'         'add by sonia 2018/1/26
'         Case Else
'            If stra1705 = "Y20026020" Then stra1705 = "Y20026000"
'            If stra1705 = "Y45878000" Then stra1705 = "Y55253000"   'add by sonia 2019/5/28 婉莘
'            If stra1705 = "Y51333010" Then stra1705 = "Y51333000"   'add by sonia 2019/6/24 婉莘
'      'end 2018/1/26
'      End Select
'      '2009/9/23 end
'      'Add by Amy 2013/11/18 +帳款處理訊息
'      strExc(0) = GetDizhang("" & adoacc150.Fields("a1503"), , True)
'      'end 2013/11/18
'      If IsNull(adoacc150.Fields("a1504").Value) Then
'         stra1706 = ""
'      Else
'         stra1706 = adoacc150.Fields("a1504").Value
'      End If
'      If IsNull(adoacc150.Fields("axf03").Value) Then
'         stra1707 = ""
'      Else
'         stra1707 = adoacc150.Fields("axf03").Value
'      End If
'      'If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'      '   stra1708 = Val(FCDate(MaskEdBox1.Text))
'      'Else
'      '   stra1708 = ""
'      'End If
'      stra1708 = strSrvDate(2)
'      'Adodc1.Recordset.Fields("a1710").Value = Val(ACDate(ServerDate))
'      'Adodc1.Recordset.Fields("a1711").Value = ServerTime
'      'Adodc1.Recordset.Fields("a1712").Value = strUserNum
'      'Adodc1.Recordset.UpdateBatch
'      adoTaie.Execute "insert into acc170 (a1701, a1702, a1703, a1704, a1705, a1706, a1707, a1708, a1710, a1711, a1712) values ('1', '" & Text1 & "', " & CNULL(stra1703) & ", " & CNULL(stra1704) & ", " & CNULL(stra1705) & ", " & CNULL(stra1706) & ", " & CNULL(stra1707) & ", " & CNULL(stra1708) & ", " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "')"
'      AdodcRefresh
'      'add by sonia 2014/7/21 J公司或獨立水單都要彈訊息提醒
'      Select Case Left(stra1707, Len(stra1707) - 9)
'         Case "CFP", "P"
'            StrSQLa = "select nvl(p1.pa26,p2.pa26) pa26,nvl(p1.pa161,p2.pa161) pa161 from acc170, acc151, patent p1, acc161, patent p2 where a1702='" & Text1 & "' and a1702=axf01(+) and a1702=axg01(+) " & _
'                      "and substr(axf03, 1, length(axf03) - 9)=p1.pa01(+) and substr(axf03, length(axf03) - 8, 6)=p1.pa02(+) and substr(axf03, length(axf03) - 2, 1)=p1.pa03(+) and substr(axf03, length(axf03) - 1, 2)=p1.pa04(+) " & _
'                      "and substr(axg03, 1, length(axg03) - 9)=p2.pa01(+) and substr(axg03, length(axg03) - 8, 6)=p2.pa02(+) and substr(axg03, length(axg03) - 2, 1)=p2.pa03(+) and substr(axg03, length(axg03) - 1, 2)=p2.pa04(+) "
'         Case "TF", "T", "CFT"
'            StrSQLa = "select nvl(t1.tm23,t2.tm23) pa26,nvl(t1.tm130,t2.tm130) pa161 from acc170, acc151, trademark t1, acc161, trademark t2 where a1702='" & Text1 & "' and a1702=axf01(+) and a1702=axg01(+) " & _
'                      "and substr(axf03, 1, length(axf03) - 9)=t1.tm01(+) and substr(axf03, length(axf03) - 8, 6)=t1.tm02(+) and substr(axf03, length(axf03) - 2, 1)=t1.tm03(+) and substr(axf03, length(axf03) - 1, 2)=t1.tm04(+) " & _
'                      "and substr(axg03, 1, length(axg03) - 9)=t2.tm01(+) and substr(axg03, length(axg03) - 8, 6)=t2.tm02(+) and substr(axg03, length(axg03) - 2, 1)=t2.tm03(+) and substr(axg03, length(axg03) - 1, 2)=t2.tm04(+) "
'         Case "L", "CFL"
'            StrSQLa = "select nvl(L1.LC11,L2.LC11) pa26,nvl(L1.lc48,L2.lc48) pa161 from acc170, acc151, LAWCASE L1, acc161, LAWCASE L2 where a1702='" & Text1 & "' and a1702=axf01(+) and a1702=axg01(+) " & _
'                      "and substr(axf03, 1, length(axf03) - 9)=L1.LC01(+) and substr(axf03, length(axf03) - 8, 6)=L1.LC02(+) and substr(axf03, length(axf03) - 2, 1)=L1.LC03(+) and substr(axf03, length(axf03) - 1, 2)=L1.LC04(+) " & _
'                      "and substr(axg03, 1, length(axg03) - 9)=L2.LC01(+) and substr(axg03, length(axg03) - 8, 6)=L2.LC02(+) and substr(axg03, length(axg03) - 2, 1)=L2.LC03(+) and substr(axg03, length(axg03) - 1, 2)=L2.LC04(+) "
'         Case Else
'            StrSQLa = "select nvl(s1.sp08,s2.sp08) pa26,nvl(s1.sp85,s2.sp85) pa161 from acc170, acc151, servicepractice s1, acc161, servicepractice s2 where a1702='" & Text1 & "' and a1702=axf01(+) and a1702=axg01(+) " & _
'                      "and substr(axf03, 1, length(axf03) - 9)=s1.sp01(+) and substr(axf03, length(axf03) - 8, 6)=s1.sp02(+) and substr(axf03, length(axf03) - 2, 1)=s1.sp03(+) and substr(axf03, length(axf03) - 1, 2)=s1.sp04(+) " & _
'                      "and substr(axg03, 1, length(axg03) - 9)=s2.sp01(+) and substr(axg03, length(axg03) - 8, 6)=s2.sp02(+) and substr(axg03, length(axg03) - 2, 1)=s2.sp03(+) and substr(axg03, length(axg03) - 1, 2)=s2.sp04(+) "
'      End Select
'      adoacc170.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'      If adoacc170.RecordCount <> 0 Then
'         If IsNull(adoacc170.Fields("pa26").Value) = False Then
'            'modify by sonia 2017/4/10 婉莘要求加X69534彩豐精技   2017/5/15陳德發及郭雅娟要求取消
'            'modify by sonia 2017/4/27 高國碩要求X60149改為只要母號,關係企業不要
'            If Left(adoacc170.Fields("pa26").Value, 6) = "X44551" Or Left(adoacc170.Fields("pa26").Value, 6) = "X62079" _
'            Or Left(adoacc170.Fields("pa26").Value, 6) = "X43988" Or Left(adoacc170.Fields("pa26").Value, 6) = "X63219" Or Left(adoacc170.Fields("pa26").Value, 6) = "X62319" _
'            Or Left(adoacc170.Fields("pa26").Value, 6) = "X60498" Or Left(adoacc170.Fields("pa26").Value, 6) = "X62702" Or Left(adoacc170.Fields("pa26").Value, 6) = "X63838" _
'             Then
'               MsgBox "此為特定客戶, 水單不能合併！", , MsgText(5)
'            End If
'            'add by sonia 2017/4/14 單獨編號而關係企業不要的,例:王副總的華碩客戶(X69011010)
'            'modify by sonia 2017/4/27 高國碩要求X60149改為只要母號,關係企業不要
'            'modify by sonia 2017/10/11 張詠翔要求取消X6014900
'            'modify by sonia 2017/10/26 郭雅娟要求加X60738010國立清華大學
'            If Left(adoacc170.Fields("pa26").Value, 8) = "X6901101" Or Left(adoacc170.Fields("pa26").Value, 8) = "X6073801" Then
'               MsgBox "此為特定客戶, 水單不能合併！", , MsgText(5)
'            End If
'            'end 2017/4/14
'         End If
'         If "" & adoacc170.Fields("pa161").Value = "J" Then
'            MsgBox "此為智權公司出名案件！", , MsgText(5)
'         End If
'      End If
'      adoacc170.Close
'      'end 2014/7/21
'   Else
'      MsgBox MsgText(28), , MsgText(5)
'   End If
'   adoacc150.Close
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
'End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Dim bolRefreshGrid As Boolean 'Added by Morgan 2017/1/24
On Error GoTo Checking
   Select Case KeyCode
      Case vbKeyInsert
         'Modified by Morgan 2017/1/24
         'Acc170Save
         Acc170SaveNew Me.Text1, bolRefreshGrid
         If bolRefreshGrid Then AdodcRefresh
         'end 2017/1/24
         
        'Modify By Cheng 2003/05/14
'         Text1 = ""
         Text1.Text = "U"
         Text1.SetFocus
         Text1_GotFocus  'ADD BY SONIA 2014/3/19
      Case vbKeyF12
         AdodcRefresh
   End Select
   KeyEnter KeyCode
Checking:
   Exit Sub
End Sub

'*************************************************
'  刪除資料表
'
'*************************************************
Private Sub AdodcDelete()
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   'Adodc1.Recordset.Delete
   'Adodc1.Recordset.UpdateBatch
   adoTaie.Execute "delete from acc170 where a1701 = '1' and a1702 = '" & Adodc1.Recordset.Fields("a1702").Value & "'"
   AdodcRefresh
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  產生結匯前核對表
'
'*************************************************
'2005/9/9 MODIFY BY SONIA
'Private Sub ProcessData()
Sub ProcessData()
Dim adocaseprogress As New ADODB.Recordset
Dim adofagent As New ADODB.Recordset
Dim adoacc180 As New ADODB.Recordset
Dim adoacc190 As New ADODB.Recordset
Dim adoquery As New ADODB.Recordset
Dim adopatent As New ADODB.Recordset   '2012/10/5 add by sonia
Dim strFagent As String, strNo As String
Dim strCompanyType As String '收據個人/公司 '2009/6/2 婧瑄說取消個人或公司及金額的條件
Dim strSystemType As String
Dim strCurrency As String
Dim strSameNo As String
Dim strYes As String
'Add By Cheng 2003/05/14
Dim StrSQLa As String
Dim strCompany As String      '公司別
Dim strLastCompany As String  '前一筆公司別
Dim strCaseNo As String       '本所案號
Dim strIndependent As String  '獨立水單  2013/2/20 ADD BY SONIA
Dim bolAddLog As Boolean 'Added by Lydia 2016/05/11
Dim strA1801List As String 'Added by Lydia 2017/09/22 記錄產生的付款單號

On Error GoTo Checking
   Select Case Combo1
      Case Mid(ReportTitle(216), 6, 8) '國內未收款明細表
' 付款明細產生
         strSameNo = ""
         strIndependent = "" '2013/2/20 ADD BY SONIA
         adoquery.CursorLocation = adUseClient
         '2005/9/20 MODIFY BY SONIA 加 a1717
         'Modify by Morgan 2006 加申請國家a0j04, 排序不用依照
         '2006/11/1 MODIFY BY SONIA 加入 CP87,CP88,並修改抵帳單抓收據方式, 帳單不能改同抵帳單因為舊帳單無收文號之故
         '2009/6/2 MODIFY BY SONIA 原順序為代理人+收據公司別+系統類別+個人或公司+幣別+金額order by a1705 asc, a0k11 asc, cp01 asc, a0k05 asc, a1703 asc, a1704 asc "
         '2009/6/2 婧瑄說取消個人或公司及金額的條件(程式碼已刪除)
         '2012/10/4 MODIFY BY SONIA 婧瑄說取消CP01條件,但大陸地區專利/商標仍要分開故加fa10
         'StrSQLa = "select a1701, a1702, a1703, a1704, a1705, a0k05, cp01, a0k04, Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11, a1717, a0j04,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,a0k01 from acc170, caseprogress, acc0k0, systemkind, acc0j0 where a0j01(+)=cp09 and a1702 = cp61 and cp60 = a0k01 (+) and a1701 = '1' and (a1709 is null or a1709 = '') and cp01=sk01 union " & _
                   "select a1701, a1702, a1703, a1704, a1705, a0k05, cp01, a0k04, Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11, a1717, a0j04,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,a0k01 from acc170, caseprogress, acc0k0, systemkind, acc0j0 where a0j01(+)=cp09 and a1702 = cp62 and cp60 = a0k01 (+) and a1701 = '1' and (a1709 is null or a1709 = '') and cp01=sk01 union " & _
                   "select a1701, a1702, a1703, a1704, a1705, a0k05, cp01, a0k04, Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11, a1717, a0j04,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,a0k01 from acc170, caseprogress, acc0k0, systemkind, acc0j0 where a0j01(+)=cp09 and a1702 = cp63 and cp60 = a0k01 (+) and a1701 = '1' and (a1709 is null or a1709 = '') and cp01=sk01 union " & _
                   "select a1701, a1702, a1703, a1704, a1705, a0k05, cp01, a0k04, Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11, a1717, a0j04,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,a0k01 from acc170, caseprogress, acc0k0, systemkind, acc0j0 where a0j01(+)=cp09 and a1702 = cp87 and cp60 = a0k01 (+) and a1701 = '1' and (a1709 is null or a1709 = '') and cp01=sk01 union " & _
                   "select a1701, a1702, a1703, a1704, a1705, a0k05, cp01, a0k04, Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11, a1717, a0j04,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,a0k01 from acc170, caseprogress, acc0k0, systemkind, acc0j0 where a0j01(+)=cp09 and a1702 = cp88 and cp60 = a0k01 (+) and a1701 = '1' and (a1709 is null or a1709 = '') and cp01=sk01 union " & _
                   "select a1701, a1702, a1703, a1704, a1705, a0k05, cp01, a0k04, Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11, a1717, a0j04,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,a0k01 from acc170, caseprogress, acc0k0, systemkind, acc0j0 where a0j01(+)=cp09 and a1702 = cp61 (+) and cp60 = a0k01 (+) and a1701 = '1' and length(a1702) = 10 and (cp61 is null and cp62 is null and cp63 is null and cp87 is null and cp88 is null) and (a1709 is null or a1709 = '') and cp01=sk01 union " & _
                   "select a1701, a1702, a1703, a1704 * (-1) as a1704, a1705, a0k05, cp01, a0k04, Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11, a1717, a0j04,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,a0k01 from acc170, acc161, acc160, caseprogress, acc0k0, systemkind, acc0j0 where a0j01(+)=cp09 and a1702 = axg01 and a1702 = a1601 and AXG02=CP09 AND CP60 = a0k01 (+) and a1701 = '2' and (a1709 is null or a1709 = '') and cp01=sk01 union " & _
                   "select a1701, a1702, a1703, a1704, a1705, ' ' as a0k05, ' ' as cp01, ' ' as a0k04, '2' as a0k11, a1717,null a0j04,null caseNo,null a0k01 from acc170 where a1701 = '3' and (a1709 is null or a1709 = '') union " & _
                   "select a1701, a1702, a1703, a1704, a1705, ' ' as a0k05, ' ' as cp01, ' ' as a0k04, '2' as a0k11, a1717,null a0j04,null caseNo,null a0k01 from acc170 where a1701 = '4' and (a1709 is null or a1709 = '') " & _
                   "order by a1705 asc, a0k11 asc, cp01 asc, a1703 asc, a0k05 ASC, a1704 asc "
         '2012/12/10 modify by sonia 因要依代理人+收據公司別+幣別排序,但無收據案件會因本所案號不同造成排序錯誤,故加function(GetA0K11)
         'StrSQLa = "select a1701, a1702, a1703, a1704, a1705, a0k05, cp01, a0k04, Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11, a1717, a0j04,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,a0k01,fa10 from acc170, caseprogress, acc0k0, systemkind, acc0j0, fagent where a0j01(+)=cp09 and a1702 = cp61 and cp60 = a0k01 (+) and a1701 = '1' and (a1709 is null or a1709 = '') and cp01=sk01 and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) union " & _
                   "select a1701, a1702, a1703, a1704, a1705, a0k05, cp01, a0k04, Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11, a1717, a0j04,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,a0k01,fa10 from acc170, caseprogress, acc0k0, systemkind, acc0j0, fagent where a0j01(+)=cp09 and a1702 = cp62 and cp60 = a0k01 (+) and a1701 = '1' and (a1709 is null or a1709 = '') and cp01=sk01 and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) union " & _
                   "select a1701, a1702, a1703, a1704, a1705, a0k05, cp01, a0k04, Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11, a1717, a0j04,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,a0k01,fa10 from acc170, caseprogress, acc0k0, systemkind, acc0j0, fagent where a0j01(+)=cp09 and a1702 = cp63 and cp60 = a0k01 (+) and a1701 = '1' and (a1709 is null or a1709 = '') and cp01=sk01 and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) union " & _
                   "select a1701, a1702, a1703, a1704, a1705, a0k05, cp01, a0k04, Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11, a1717, a0j04,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,a0k01,fa10 from acc170, caseprogress, acc0k0, systemkind, acc0j0, fagent where a0j01(+)=cp09 and a1702 = cp87 and cp60 = a0k01 (+) and a1701 = '1' and (a1709 is null or a1709 = '') and cp01=sk01 and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) union " & _
                   "select a1701, a1702, a1703, a1704, a1705, a0k05, cp01, a0k04, Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11, a1717, a0j04,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,a0k01,fa10 from acc170, caseprogress, acc0k0, systemkind, acc0j0, fagent where a0j01(+)=cp09 and a1702 = cp88 and cp60 = a0k01 (+) and a1701 = '1' and (a1709 is null or a1709 = '') and cp01=sk01 and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) union " & _
                   "select a1701, a1702, a1703, a1704, a1705, a0k05, cp01, a0k04, Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11, a1717, a0j04,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,a0k01,fa10 from acc170, caseprogress, acc0k0, systemkind, acc0j0, fagent where a0j01(+)=cp09 and a1702 = cp61 (+) and cp60 = a0k01 (+) and a1701 = '1' and length(a1702) = 10 and (cp61 is null and cp62 is null and cp63 is null and cp87 is null and cp88 is null) and (a1709 is null or a1709 = '') and cp01=sk01 and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) union " & _
                   "select a1701, a1702, a1703, a1704 * (-1) as a1704, a1705, a0k05, cp01, a0k04, Decode(a0k01,Null,Decode(sk02,'1','2','5','2','1'),a0k11) as a0k11, a1717, a0j04,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,a0k01,fa10 from acc170, acc161, acc160, caseprogress, acc0k0, systemkind, acc0j0, fagent where a0j01(+)=cp09 and a1702 = axg01 and a1702 = a1601 and AXG02=CP09 AND CP60 = a0k01 (+) and a1701 = '2' and (a1709 is null or a1709 = '') and cp01=sk01 and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) union " & _
                   "select a1701, a1702, a1703, a1704, a1705, ' ' as a0k05, ' ' as cp01, ' ' as a0k04, '2' as a0k11, a1717,null a0j04,null caseNo,null a0k01,fa10 from acc170, fagent where a1701 = '3' and (a1709 is null or a1709 = '') and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) union " & _
                   "select a1701, a1702, a1703, a1704, a1705, ' ' as a0k05, ' ' as cp01, ' ' as a0k04, '2' as a0k11, a1717,null a0j04,null caseNo,null a0k01,fa10 from acc170, fagent where a1701 = '4' and (a1709 is null or a1709 = '') and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) " & _
                   "order by a1705 asc, a0k11 asc, a1703 asc, a0k05 ASC, a1704 asc "
         '2013/1/29 modify by sonia cp01改用sk02,同時取消a0k05,a1704的排序條件但加入sk02
         'StrSQLa = "select a1701, a1702, a1703, a1704, a1705, a0k05, cp01, a0k04, GetA0K11(cp09) as a0k11, a1717, a0j04,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,a0k01,fa10 from acc170, caseprogress, acc0k0, acc0j0, fagent where a0j01(+)=cp09 and a1702 = cp61 and cp60 = a0k01 (+) and a1701 = '1' and (a1709 is null or a1709 = '') and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) union " & _
                   "select a1701, a1702, a1703, a1704, a1705, a0k05, cp01, a0k04, GetA0K11(cp09) as a0k11, a1717, a0j04,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,a0k01,fa10 from acc170, caseprogress, acc0k0, acc0j0, fagent where a0j01(+)=cp09 and a1702 = cp62 and cp60 = a0k01 (+) and a1701 = '1' and (a1709 is null or a1709 = '') and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) union " & _
                   "select a1701, a1702, a1703, a1704, a1705, a0k05, cp01, a0k04, GetA0K11(cp09) as a0k11, a1717, a0j04,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,a0k01,fa10 from acc170, caseprogress, acc0k0, acc0j0, fagent where a0j01(+)=cp09 and a1702 = cp63 and cp60 = a0k01 (+) and a1701 = '1' and (a1709 is null or a1709 = '') and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) union " & _
                   "select a1701, a1702, a1703, a1704, a1705, a0k05, cp01, a0k04, GetA0K11(cp09) as a0k11, a1717, a0j04,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,a0k01,fa10 from acc170, caseprogress, acc0k0, acc0j0, fagent where a0j01(+)=cp09 and a1702 = cp87 and cp60 = a0k01 (+) and a1701 = '1' and (a1709 is null or a1709 = '') and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) union " & _
                   "select a1701, a1702, a1703, a1704, a1705, a0k05, cp01, a0k04, GetA0K11(cp09) as a0k11, a1717, a0j04,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,a0k01,fa10 from acc170, caseprogress, acc0k0, acc0j0, fagent where a0j01(+)=cp09 and a1702 = cp88 and cp60 = a0k01 (+) and a1701 = '1' and (a1709 is null or a1709 = '') and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) union " & _
                   "select a1701, a1702, a1703, a1704, a1705, a0k05, cp01, a0k04, GetA0K11(cp09) as a0k11, a1717, a0j04,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,a0k01,fa10 from acc170, caseprogress, acc0k0, acc0j0, fagent where a0j01(+)=cp09 and a1702 = cp61 (+) and cp60 = a0k01 (+) and a1701 = '1' and length(a1702) = 10 and (cp61 is null and cp62 is null and cp63 is null and cp87 is null and cp88 is null) and (a1709 is null or a1709 = '') and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) union " & _
                   "select a1701, a1702, a1703, a1704 * (-1) as a1704, a1705, a0k05, cp01, a0k04, GetA0K11(cp09) as a0k11, a1717, a0j04,cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo,a0k01,fa10 from acc170, acc161, acc160, caseprogress, acc0k0, acc0j0, fagent where a0j01(+)=cp09 and a1702 = axg01 and a1702 = a1601 and AXG02=CP09 AND CP60 = a0k01 (+) and a1701 = '2' and (a1709 is null or a1709 = '') and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) union " & _
                   "select a1701, a1702, a1703, a1704, a1705, ' ' as a0k05, ' ' as cp01, ' ' as a0k04, '2' as a0k11, a1717,null a0j04,null caseNo,null a0k01,fa10 from acc170, fagent where a1701 = '3' and (a1709 is null or a1709 = '') and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) union " & _
                   "select a1701, a1702, a1703, a1704, a1705, ' ' as a0k05, ' ' as cp01, ' ' as a0k04, '2' as a0k11, a1717,null a0j04,null caseNo,null a0k01,fa10 from acc170, fagent where a1701 = '4' and (a1709 is null or a1709 = '') and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) " & _
                   "order by a1705 asc, 9 asc, a1703 asc, a0k05 ASC, a1704 asc "
         '2013/2/21 modify by sonia 加入排序a0k04 asc, caseNo asc
         'modify by sonia 2014/7/21 a0k04改為GETA0K04(CP09) CFT-16504(U10303792),cp01||'-'||cp02||'-'||cp03||'-'||cp04 caseNo改為cp01||cp02||cp03||cp04 caseNo
         'Modified by Lydia 2015/10/06 +A1718
         'Modified by Morgan 2019/7/16 +A1720
         StrSQLa = "select a1701, a1702, a1703, a1704, a1705, a0k05, decode(sk02,'5','1','6','2','7','3','8','4',sk02) sk02, GETA0K04(CP09) as a0k04, GetA0K11(cp09) as a0k11, a1717, a0j04,cp01||cp02||cp03||cp04 caseNo,a0k01,fa10,a1718,A1720 from acc170, caseprogress, acc0k0, acc0j0, fagent, systemkind where a0j01(+)=cp09 and a1702 = cp61 and cp60 = a0k01 (+) and a1701 = '1' and (a1709 is null or a1709 = '') and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) and cp01=sk01(+) union " & _
                   "select a1701, a1702, a1703, a1704, a1705, a0k05, decode(sk02,'5','1','6','2','7','3','8','4',sk02) sk02, GETA0K04(CP09) as a0k04, GetA0K11(cp09) as a0k11, a1717, a0j04,cp01||cp02||cp03||cp04 caseNo,a0k01,fa10,a1718,A1720 from acc170, caseprogress, acc0k0, acc0j0, fagent, systemkind where a0j01(+)=cp09 and a1702 = cp62 and cp60 = a0k01 (+) and a1701 = '1' and (a1709 is null or a1709 = '') and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) and cp01=sk01(+) union " & _
                   "select a1701, a1702, a1703, a1704, a1705, a0k05, decode(sk02,'5','1','6','2','7','3','8','4',sk02) sk02, GETA0K04(CP09) as a0k04, GetA0K11(cp09) as a0k11, a1717, a0j04,cp01||cp02||cp03||cp04 caseNo,a0k01,fa10,a1718,A1720 from acc170, caseprogress, acc0k0, acc0j0, fagent, systemkind where a0j01(+)=cp09 and a1702 = cp63 and cp60 = a0k01 (+) and a1701 = '1' and (a1709 is null or a1709 = '') and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) and cp01=sk01(+) union " & _
                   "select a1701, a1702, a1703, a1704, a1705, a0k05, decode(sk02,'5','1','6','2','7','3','8','4',sk02) sk02, GETA0K04(CP09) as a0k04, GetA0K11(cp09) as a0k11, a1717, a0j04,cp01||cp02||cp03||cp04 caseNo,a0k01,fa10,a1718,A1720 from acc170, caseprogress, acc0k0, acc0j0, fagent, systemkind where a0j01(+)=cp09 and a1702 = cp87 and cp60 = a0k01 (+) and a1701 = '1' and (a1709 is null or a1709 = '') and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) and cp01=sk01(+) union " & _
                   "select a1701, a1702, a1703, a1704, a1705, a0k05, decode(sk02,'5','1','6','2','7','3','8','4',sk02) sk02, GETA0K04(CP09) as a0k04, GetA0K11(cp09) as a0k11, a1717, a0j04,cp01||cp02||cp03||cp04 caseNo,a0k01,fa10,a1718,A1720 from acc170, caseprogress, acc0k0, acc0j0, fagent, systemkind where a0j01(+)=cp09 and a1702 = cp88 and cp60 = a0k01 (+) and a1701 = '1' and (a1709 is null or a1709 = '') and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) and cp01=sk01(+) union " & _
                   "select a1701, a1702, a1703, a1704, a1705, a0k05, decode(sk02,'5','1','6','2','7','3','8','4',sk02) sk02, GETA0K04(CP09) as a0k04, GetA0K11(cp09) as a0k11, a1717, a0j04,cp01||cp02||cp03||cp04 caseNo,a0k01,fa10,a1718,A1720 from acc170, caseprogress, acc0k0, acc0j0, fagent, systemkind where a0j01(+)=cp09 and a1702 = cp61 (+) and cp60 = a0k01 (+) and a1701 = '1' and length(a1702) = 10 and (cp61 is null and cp62 is null and cp63 is null and cp87 is null and cp88 is null) and (a1709 is null or a1709 = '') and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) and cp01=sk01(+) union " & _
                   "select a1701, a1702, a1703, a1704 * (-1) as a1704, a1705, a0k05, decode(sk02,'5','1','6','2','7','3','8','4',sk02) sk02, GETA0K04(CP09) as a0k04, GetA0K11(cp09) as a0k11, a1717, a0j04,cp01||cp02||cp03||cp04 caseNo,a0k01,fa10,a1718,A1720 from acc170, acc161, acc160, caseprogress, acc0k0, acc0j0, fagent, systemkind where a0j01(+)=cp09 and a1702 = axg01 and a1702 = a1601 and AXG02=CP09 AND CP60 = a0k01 (+) and a1701 = '2' and (a1709 is null or a1709 = '') and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) and cp01=sk01(+) union " & _
                   "select a1701, a1702, a1703, a1704, a1705, ' ' as a0k05, ' ' as sk02, ' ' as a0k04, '2' as a0k11, a1717,null a0j04,null caseNo,null a0k01,fa10,a1718,A1720 from acc170, fagent where a1701 = '3' and (a1709 is null or a1709 = '') and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) union " & _
                   "select a1701, a1702, a1703, a1704, a1705, ' ' as a0k05, ' ' as sk02, ' ' as a0k04, '2' as a0k11, a1717,null a0j04,null caseNo,null a0k01,fa10,a1718,A1720 from acc170, fagent where a1701 = '4' and (a1709 is null or a1709 = '') and substr(a1705,1,8)=fa01(+) and substr(a1705,9,1)=fa02(+) "
         'Modified by Moran 2019/7/16 +獨立水單(但相同案號的可合併)
         'StrSQLa = StrSQLa & " order by a1705 asc, 9 asc, a1703 asc, 7 asc, 8 asc, caseNo asc, a1702 asc "
         StrSQLa = "select * from (" & StrSQLa & ") X order by a1705 asc, 9 asc, a1703 asc,decode(a1720,null,'2','1'||caseNo) asc, 7 asc, 8 asc, caseNo asc, a1702 asc "
         
         adoquery.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
        'Added by Lydia 2016/05/11 記錄執行時間,供台銀結匯資料比對,請參考GetTermOfPayment
        If adoquery.RecordCount > 0 And Pub_StrUserSt03 <> "M51" Then
           PUB_AddExcuteLog "Frmacc2170"
           bolAddLog = True
        End If
        
        'Added by Morgan 2019/7/16
        '更新 獨立水單註記 A1720
        strSameNo = ""
        Do While Not adoquery.EOF
            If strSameNo <> adoquery.Fields("a1701") & adoquery.Fields("a1702") Then
               strSameNo = adoquery.Fields("a1701") & adoquery.Fields("a1702")
               If Not IsNull(adoquery.Fields("caseno")) Then
                  strIndependent = ""
                  If PUB_ChkNoMergePayCust(adoquery.Fields("caseno")) = True Then
                     strIndependent = "X"
                  End If
                  If adoquery.Fields("a1701") = "1" Then
                     cnnConnection.Execute "update acc150 set a1526=a1526 where a1501='" & adoquery.Fields("a1702") & "' and a1526='Y'", intI
                     If intI = 1 Then
                        If strIndependent = "" Then
                           strIndependent = "U"
                        Else
                           strIndependent = "B"
                        End If
                     End If
                  End If
               End If
               cnnConnection.Execute "update acc170 set a1720='" & strIndependent & "' where a1701='" & adoquery.Fields("a1701") & "' and a1702='" & adoquery.Fields("a1702") & "'", intI
               
            End If
            adoquery.MoveNext
        Loop
        strSameNo = ""
        strIndependent = ""
        adoquery.Requery
        'end 2019/7/16
        
        'Add By Cheng 2003/05/14
        '記錄公司別
         Do While adoquery.EOF = False
            'Debug.Print "a1702=" & adoquery.Fields("a1702").Value & "  前一筆=" & strSameNo
            If strSameNo <> "" And strSameNo = adoquery.Fields("a1702").Value Then
               strSameNo = adoquery.Fields("a1702").Value
               GoTo NextSkip
            Else
               strSameNo = adoquery.Fields("a1702").Value
            End If
            
            '若有代理人
            If IsNull(adoquery.Fields("a1705").Value) = False Then
               '2012/12/10 MODIFY BY SONIA 移至上面抓資料語法內以function(GetAmtComp)抓
               'strCompany = GetComp("" & adoquery.Fields("a0k01").Value, "" & adoquery.Fields("caseno").Value, "" & adoquery.Fields("a0k11").Value)
               strCompany = "" & adoquery.Fields("a0k11").Value
               '2012/12/10 end
               
               '若代理人不同或公司別不同時
'Modify by Morgan 2006/7/17
               'Debug.Print "a1705=" & adoquery.Fields("a1705").Value & "  前一筆=" & strFagent & ", a0k01=" & strCompany & "  前一筆=" & strLastCompany
               If strFagent <> adoquery.Fields("a1705").Value Or strCompany <> strLastCompany Then
                  strLastCompany = strCompany
'end 2006/7/17
                  strNo = AutoNo(MsgText(814), 5)
                  strA1801List = strA1801List & strNo & ","  'Added by Lydia 2017/09/22 記錄產生的付款單號
                  
                  If IsNull(adoquery.Fields("a0k05").Value) = False Then
                     strCompanyType = adoquery.Fields("a0k05").Value
                  Else
                     strCompanyType = "2"
                  End If
                  strCurrency = adoquery.Fields("a1703").Value
                  strSystemType = adoquery.Fields("sk02").Value
                  '2005/9/20 MODIFY BY SONIA
                  'adoTaie.Execute "insert into acc180 (a1801, a1802, a1803, a1806, a1804, a1805, a1811) values ('" & strNo & "', " & Val(ACDate(ServerDate)) & ", '" & adoquery.Fields("a1705").Value & "', '" & strUserNum & "', " & Val(ACDate(ServerDate)) & ", " & ServerTime & ",'" & GetTermOfPayment(strSameNo, "" & adoquery.Fields("a1703").Value) & "' )"
                  'Modified by Lydia 2015/06/12 +台銀電匯紙本判斷
                  'adoTaie.Execute "insert into acc180 (a1801, a1802, a1803, a1806, a1804, a1805, a1811, a1810) values ('" & strNo & "', " & strSrvDate(2) & ", '" & adoquery.Fields("a1705").Value & "', '" & strUserNum & "', " & strSrvDate(2) & ", " & ServerTime & ",'" & GetTermOfPayment(strSameNo, "" & adoquery.Fields("a1703").Value) & "', '" & ChgSQL("" & adoquery.Fields("a1717").Value) & "' )"
                  'Modified by Lydia 2015/10/06 +A1718
                  adoTaie.Execute "insert into acc180 (a1801, a1802, a1803, a1806, a1804, a1805, a1811, a1810) values ('" & strNo & "', " & strSrvDate(2) & ", '" & adoquery.Fields("a1705").Value & "', '" & strUserNum & "', " & strSrvDate(2) & ", " & ServerTime & ",'" & GetTermOfPayment(strSameNo, "" & adoquery.Fields("a1703").Value, strCompany, "" & adoquery.Fields("a1718")) & "', '" & ChgSQL("" & adoquery.Fields("a1717").Value) & "' )"
                  '2005/9/20 END
                  strFagent = adoquery.Fields("a1705").Value
                  GoTo insertdata
               End If
               
               If strCompany <> "J" Then  'add by sonia 2014/7/21 非智權公司的大陸才要分開 W10301380,W10301379不必分開
                  '2012/10/4 MODIFY BY SONIA 婧瑄說取消CP01條件,但大陸地區專利/商標仍要分開
                  'If strSystemType <> adoquery.Fields("cp01").Value Then
                  If strSystemType <> adoquery.Fields("sk02").Value And adoquery.Fields("fa10").Value = "020" Then
                     strNo = AutoNo(MsgText(814), 5)
                     strA1801List = strA1801List & strNo & ","  'Added by Lydia 2017/09/22 記錄產生的付款單號
                     
                     If IsNull(adoquery.Fields("a0k05").Value) = False Then
                        strCompanyType = adoquery.Fields("a0k05").Value
                     Else
                        'modify by sonia 2014/7/21 改同上面
                        'strCompanyType = ""
                        strCompanyType = "2"
                     End If
                     strCurrency = adoquery.Fields("a1703").Value
                     strSystemType = adoquery.Fields("sk02").Value
                     '2005/9/20 MODIFY BY SONIA
                     'adoTaie.Execute "insert into acc180 (a1801, a1802, a1803, a1806, a1804, a1805, a1811) values ('" & strNo & "', " & Val(ACDate(ServerDate)) & ", '" & adoquery.Fields("a1705").Value & "', '" & strUserNum & "', " & Val(ACDate(ServerDate)) & ", " & ServerTime & ",'" & GetTermOfPayment(strSameNo, "" & adoquery.Fields("a1703").Value) & "' )"
                     'Modified by Lydia 2015/06/12 +台銀電匯紙本判斷
                     'adoTaie.Execute "insert into acc180 (a1801, a1802, a1803, a1806, a1804, a1805, a1811, a1810) values ('" & strNo & "', " & strSrvDate(2) & ", '" & adoquery.Fields("a1705").Value & "', '" & strUserNum & "', " & strSrvDate(2) & ", " & ServerTime & ",'" & GetTermOfPayment(strSameNo, "" & adoquery.Fields("a1703").Value) & "', '" & ChgSQL("" & adoquery.Fields("a1717").Value) & "' )"
                     'Modified by Lydia 2015/10/06 +A1718
                     adoTaie.Execute "insert into acc180 (a1801, a1802, a1803, a1806, a1804, a1805, a1811, a1810) values ('" & strNo & "', " & strSrvDate(2) & ", '" & adoquery.Fields("a1705").Value & "', '" & strUserNum & "', " & strSrvDate(2) & ", " & ServerTime & ",'" & GetTermOfPayment(strSameNo, "" & adoquery.Fields("a1703").Value, strCompany, "" & adoquery.Fields("a1718")) & "', '" & ChgSQL("" & adoquery.Fields("a1717").Value) & "' )"
                     
                     '2005/9/20 END
                     GoTo insertdata
                  End If
               End If  'add by sonia 2014/7/21
                        
               '若幣別不同
               'Debug.Print "a1703=" & adoquery.Fields("a1703").Value & "  前一筆=" & strCurrency
               If strCurrency <> adoquery.Fields("a1703").Value Then
                  strNo = AutoNo(MsgText(814), 5)
                  strA1801List = strA1801List & strNo & ","  'Added by Lydia 2017/09/22 記錄產生的付款單號
                  
                 'Modify By Cheng 2003/07/22
'                     adoTaie.Execute "insert into acc180 (a1801, a1802, a1803, a1806, a1804, a1805) values ('" & strNo & "', " & Val(ACDate(ServerDate)) & ", '" & adoquery.Fields("a1705").Value & "', '" & strUserNum & "', " & Val(ACDate(ServerDate)) & ", " & ServerTime & ")"
                  '2005/9/20 MODIFY BY SONIA
                  'adoTaie.Execute "insert into acc180 (a1801, a1802, a1803, a1806, a1804, a1805, a1811) values ('" & strNo & "', " & Val(ACDate(ServerDate)) & ", '" & adoquery.Fields("a1705").Value & "', '" & strUserNum & "', " & Val(ACDate(ServerDate)) & ", " & ServerTime & ",'" & GetTermOfPayment(strSameNo, "" & adoquery.Fields("a1703").Value) & "' )"
                  'Modified by Lydia 2015/06/12 +台銀電匯紙本判斷
                  'adoTaie.Execute "insert into acc180 (a1801, a1802, a1803, a1806, a1804, a1805, a1811, a1810) values ('" & strNo & "', " & strSrvDate(2) & ", '" & adoquery.Fields("a1705").Value & "', '" & strUserNum & "', " & strSrvDate(2) & ", " & ServerTime & ",'" & GetTermOfPayment(strSameNo, "" & adoquery.Fields("a1703").Value) & "', '" & ChgSQL("" & adoquery.Fields("a1717").Value) & "' )"
                  'Modified by Lydia 2015/10/06 +A1718
                  adoTaie.Execute "insert into acc180 (a1801, a1802, a1803, a1806, a1804, a1805, a1811, a1810) values ('" & strNo & "', " & strSrvDate(2) & ", '" & adoquery.Fields("a1705").Value & "', '" & strUserNum & "', " & strSrvDate(2) & ", " & ServerTime & ",'" & GetTermOfPayment(strSameNo, "" & adoquery.Fields("a1703").Value, strCompany, "" & adoquery.Fields("a1718")) & "', '" & ChgSQL("" & adoquery.Fields("a1717").Value) & "' )"
                  '2005/9/20 END
                  strCurrency = adoquery.Fields("a1703").Value
                  GoTo insertdata
               End If
               
               'Added by Morgan 2022/7/8
               '德國年費要單筆單筆繳交, 一個U單號一張水單, 同案號也不合併 --婉莘
               If strFagent = "Y55766000" Then
                  strIndependent = "Y"
                  strNo = AutoNo(MsgText(814), 5)
                  strA1801List = strA1801List & strNo & ","
                  'Modified by Morgan 2022/7/20 +a1812
                  adoTaie.Execute "insert into acc180 (a1801, a1802, a1803, a1806, a1804, a1805, a1811, a1810, a1812) values ('" & strNo & "', " & strSrvDate(2) & ", '" & adoquery.Fields("a1705").Value & "', '" & strUserNum & "', " & strSrvDate(2) & ", " & ServerTime & ",'" & GetTermOfPayment(strSameNo, "" & adoquery.Fields("a1703").Value, strCompany, "" & adoquery.Fields("a1718")) & "', '" & ChgSQL("" & adoquery.Fields("a1717").Value) & "','Y')"
                  GoTo insertdata
               End If
               'end 2022/7/8
               
               '2013/2/20 ADD BY SONIA 前一筆為獨立水單時不同案號也要分開
               'Debug.Print "caseno=" & adoquery.Fields("caseno").Value & "  前一筆=" & strCaseNo
               If strIndependent = "Y" Then
                  If IsNull(adoquery.Fields("caseno").Value) = False Then
                     If adoquery.Fields("caseno").Value <> strCaseNo Then
                        strNo = AutoNo(MsgText(814), 5)
                        strA1801List = strA1801List & strNo & ","  'Added by Lydia 2017/09/22 記錄產生的付款單號
                        
                        'Modified by Lydia 2015/06/12 +台銀電匯紙本判斷
                        'adoTaie.Execute "insert into acc180 (a1801, a1802, a1803, a1806, a1804, a1805, a1811, a1810) values ('" & strNo & "', " & strSrvDate(2) & ", '" & adoquery.Fields("a1705").Value & "', '" & strUserNum & "', " & strSrvDate(2) & ", " & ServerTime & ",'" & GetTermOfPayment(strSameNo, "" & adoquery.Fields("a1703").Value) & "', '" & ChgSQL("" & adoquery.Fields("a1717").Value) & "' )"
                        'Modified by Lydia 2015/10/06 +A1718
                        adoTaie.Execute "insert into acc180 (a1801, a1802, a1803, a1806, a1804, a1805, a1811, a1810) values ('" & strNo & "', " & strSrvDate(2) & ", '" & adoquery.Fields("a1705").Value & "', '" & strUserNum & "', " & strSrvDate(2) & ", " & ServerTime & ",'" & GetTermOfPayment(strSameNo, "" & adoquery.Fields("a1703").Value, strCompany, "" & adoquery.Fields("a1718")) & "', '" & ChgSQL("" & adoquery.Fields("a1717").Value) & "' )"
                     End If
                     GoTo insertdata
                  End If
               End If
               '2013/2/20 END
               
'Modified by Morgan 2019/7/16 是否獨立水單改判斷 A1720
'               '2012/10/4 ADD BY SONIA X63219國立中正大學 及 X43988060國立虎尾科技大學 的CFP案水單要個案單獨出,且不同案號也要分開
'               If IsNull(adoquery.Fields("caseno").Value) = False Then
'                  '2014/4/1 CANCEL BY SONIA 所有國外案都要(原只做CFP)
'                  adopatent.CursorLocation = adUseClient
'                  'modify by sonia 2014/7/21 四個基本檔都抓
'                  'If Left(adoquery.Fields("caseno").Value, 3) = "CFP" Or Left(adoquery.Fields("caseno").Value, 3) = "P" Then
'                  '   StrSQLa = "select nvl(p1.pa26,p2.pa26) pa26 from acc170, acc151, patent p1, acc161, patent p2 where a1702='" & strSameNo & "' and a1702=axf01(+) and a1702=axg01(+) " & _
'                  '             "and substr(axf03, 1, length(axf03) - 9)=p1.pa01(+) and substr(axf03, length(axf03) - 8, 6)=p1.pa02(+) and substr(axf03, length(axf03) - 2, 1)=p1.pa03(+) and substr(axf03, length(axf03) - 1, 2)=p1.pa04(+) " & _
'                  '             "and substr(axg03, 1, length(axg03) - 9)=p2.pa01(+) and substr(axg03, length(axg03) - 8, 6)=p2.pa02(+) and substr(axg03, length(axg03) - 2, 1)=p2.pa03(+) and substr(axg03, length(axg03) - 1, 2)=p2.pa04(+) "
'                  'Else
'                  '   StrSQLa = "select nvl(t1.tm23,t2.tm23) pa26 from acc170, acc151, trademark t1, acc161, trademark t2 where a1702='" & strSameNo & "' and a1702=axf01(+) and a1702=axg01(+) " & _
'                  '             "and substr(axf03, 1, length(axf03) - 9)=t1.tm01(+) and substr(axf03, length(axf03) - 8, 6)=t1.tm02(+) and substr(axf03, length(axf03) - 2, 1)=t1.tm03(+) and substr(axf03, length(axf03) - 1, 2)=t1.tm04(+) " & _
'                  '             "and substr(axg03, 1, length(axg03) - 9)=t2.tm01(+) and substr(axg03, length(axg03) - 8, 6)=t2.tm02(+) and substr(axg03, length(axg03) - 2, 1)=t2.tm03(+) and substr(axg03, length(axg03) - 1, 2)=t2.tm04(+) "
'                  'End If
'                  Select Case Left(adoquery.Fields("caseno").Value, Len(adoquery.Fields("caseno").Value) - 9)
'                     Case "CFP", "P"
'                        StrSQLa = "select nvl(p1.pa26,p2.pa26) pa26 from acc170, acc151, patent p1, acc161, patent p2 where a1702='" & strSameNo & "' and a1702=axf01(+) and a1702=axg01(+) " & _
'                                  "and substr(axf03, 1, length(axf03) - 9)=p1.pa01(+) and substr(axf03, length(axf03) - 8, 6)=p1.pa02(+) and substr(axf03, length(axf03) - 2, 1)=p1.pa03(+) and substr(axf03, length(axf03) - 1, 2)=p1.pa04(+) " & _
'                                  "and substr(axg03, 1, length(axg03) - 9)=p2.pa01(+) and substr(axg03, length(axg03) - 8, 6)=p2.pa02(+) and substr(axg03, length(axg03) - 2, 1)=p2.pa03(+) and substr(axg03, length(axg03) - 1, 2)=p2.pa04(+) "
'                     Case "TF", "T", "CFT"
'                        StrSQLa = "select nvl(t1.tm23,t2.tm23) pa26 from acc170, acc151, trademark t1, acc161, trademark t2 where a1702='" & strSameNo & "' and a1702=axf01(+) and a1702=axg01(+) " & _
'                                  "and substr(axf03, 1, length(axf03) - 9)=t1.tm01(+) and substr(axf03, length(axf03) - 8, 6)=t1.tm02(+) and substr(axf03, length(axf03) - 2, 1)=t1.tm03(+) and substr(axf03, length(axf03) - 1, 2)=t1.tm04(+) " & _
'                                  "and substr(axg03, 1, length(axg03) - 9)=t2.tm01(+) and substr(axg03, length(axg03) - 8, 6)=t2.tm02(+) and substr(axg03, length(axg03) - 2, 1)=t2.tm03(+) and substr(axg03, length(axg03) - 1, 2)=t2.tm04(+) "
'                     Case "L", "CFL"
'                        StrSQLa = "select nvl(L1.LC11,L2.LC11) pa26 from acc170, acc151, LAWCASE L1, acc161, LAWCASE L2 where a1702='" & strSameNo & "' and a1702=axf01(+) and a1702=axg01(+) " & _
'                                  "and substr(axf03, 1, length(axf03) - 9)=L1.LC01(+) and substr(axf03, length(axf03) - 8, 6)=L1.LC02(+) and substr(axf03, length(axf03) - 2, 1)=L1.LC03(+) and substr(axf03, length(axf03) - 1, 2)=L1.LC04(+) " & _
'                                  "and substr(axg03, 1, length(axg03) - 9)=L2.LC01(+) and substr(axg03, length(axg03) - 8, 6)=L2.LC02(+) and substr(axg03, length(axg03) - 2, 1)=L2.LC03(+) and substr(axg03, length(axg03) - 1, 2)=L2.LC04(+) "
'                     Case Else
'                        StrSQLa = "select nvl(s1.sp08,s2.sp08) pa26 from acc170, acc151, servicepractice s1, acc161, servicepractice s2 where a1702='" & strSameNo & "' and a1702=axf01(+) and a1702=axg01(+) " & _
'                                  "and substr(axf03, 1, length(axf03) - 9)=s1.sp01(+) and substr(axf03, length(axf03) - 8, 6)=s1.sp02(+) and substr(axf03, length(axf03) - 2, 1)=s1.sp03(+) and substr(axf03, length(axf03) - 1, 2)=s1.sp04(+) " & _
'                                  "and substr(axg03, 1, length(axg03) - 9)=s2.sp01(+) and substr(axg03, length(axg03) - 8, 6)=s2.sp02(+) and substr(axg03, length(axg03) - 2, 1)=s2.sp03(+) and substr(axg03, length(axg03) - 1, 2)=s2.sp04(+) "
'                  End Select
'                  'end 2014/7/21
'                  adopatent.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'                  If adopatent.RecordCount <> 0 Then
'                     If IsNull(adopatent.Fields("pa26").Value) = False Then
'                        '2013/3/12 MODIFY BY SONIA 加入 X6383801中國醫藥大學
'                        'MODIFY BY SONIA 2014/3/28 顏永堅3/25郵件所列大學及其關係企業都要加
'                        'If Left(adopatent.Fields("pa26").Value, 8) = "X6321900" Or Left(adopatent.Fields("pa26").Value, 8) = "X4398806" Or Left(adopatent.Fields("pa26").Value, 8) = "X6383801" Then
'                        'modify by sonia 2017/4/10 婉莘要求加X69534彩豐精技   2017/5/15陳德發及郭雅娟要求取消
'                        'modify by sonia 2017/4/27 高國碩要求X60149改為只要母號,關係企業不要
'                        If Left(adopatent.Fields("pa26").Value, 6) = "X44551" Or Left(adopatent.Fields("pa26").Value, 6) = "X62079" _
'                        Or Left(adopatent.Fields("pa26").Value, 6) = "X43988" Or Left(adopatent.Fields("pa26").Value, 6) = "X63219" Or Left(adopatent.Fields("pa26").Value, 6) = "X62319" _
'                        Or Left(adopatent.Fields("pa26").Value, 6) = "X60498" Or Left(adopatent.Fields("pa26").Value, 6) = "X62702" Or Left(adopatent.Fields("pa26").Value, 6) = "X63838" _
'                         Then
'                           strNo = AutoNo(MsgText(814), 5)
'                           strA1801List = strA1801List & strNo & ","  'Added by Lydia 2017/09/22 記錄產生的付款單號
'
'                           'Modified by Lydia 2015/06/12 +台銀電匯紙本判斷
'                           'adoTaie.Execute "insert into acc180 (a1801, a1802, a1803, a1806, a1804, a1805, a1811, a1810) values ('" & strNo & "', " & strSrvDate(2) & ", '" & adoquery.Fields("a1705").Value & "', '" & strUserNum & "', " & strSrvDate(2) & ", " & ServerTime & ",'" & GetTermOfPayment(strSameNo, "" & adoquery.Fields("a1703").Value) & "', '" & ChgSQL("" & adoquery.Fields("a1717").Value) & "' )"
'                           'Modified by Lydia 2015/10/06 +A1718
'                           adoTaie.Execute "insert into acc180 (a1801, a1802, a1803, a1806, a1804, a1805, a1811, a1810) values ('" & strNo & "', " & strSrvDate(2) & ", '" & adoquery.Fields("a1705").Value & "', '" & strUserNum & "', " & strSrvDate(2) & ", " & ServerTime & ",'" & GetTermOfPayment(strSameNo, "" & adoquery.Fields("a1703").Value, strCompany, "" & adoquery.Fields("a1718")) & "', '" & ChgSQL("" & adoquery.Fields("a1717").Value) & "' )"
'                           adopatent.Close
'                           GoTo InsertData
'                        End If
'                        'add by sonia 2017/4/14 單獨編號而關係企業不要的,例:王副總的華碩客戶(X69011010)
'                        'modify by sonia 2017/4/27 高國碩要求X60149改為只要母號,關係企業不要
'                        'modify by sonia 2017/10/11 張詠翔要求取消X6014900
'                        'modify by sonia 2017/10/26 郭雅娟要求加X60738010國立清華大學
'                        If Left(adopatent.Fields("pa26").Value, 8) = "X6901101" Or Left(adopatent.Fields("pa26").Value, 8) = "X6073801" Then
'                           strNo = AutoNo(MsgText(814), 5)
'                           strA1801List = strA1801List & strNo & ","  'Added by Lydia 2017/09/22 記錄產生的付款單號
'
'                           adoTaie.Execute "insert into acc180 (a1801, a1802, a1803, a1806, a1804, a1805, a1811, a1810) values ('" & strNo & "', " & strSrvDate(2) & ", '" & adoquery.Fields("a1705").Value & "', '" & strUserNum & "', " & strSrvDate(2) & ", " & ServerTime & ",'" & GetTermOfPayment(strSameNo, "" & adoquery.Fields("a1703").Value, strCompany, "" & adoquery.Fields("a1718")) & "', '" & ChgSQL("" & adoquery.Fields("a1717").Value) & "' )"
'                           adopatent.Close
'                           GoTo InsertData
'                        End If
'                        'end 2017/4/14
'                     End If
'                  End If
'                  adopatent.Close
'               End If
'               '2012/10/4 END

               If Not IsNull(adoquery.Fields("A1720")) Then
                  strNo = AutoNo(MsgText(814), 5)
                  strA1801List = strA1801List & strNo & ","
                  adoTaie.Execute "insert into acc180 (a1801, a1802, a1803, a1806, a1804, a1805, a1811, a1810) values ('" & strNo & "', " & strSrvDate(2) & ", '" & adoquery.Fields("a1705").Value & "', '" & strUserNum & "', " & strSrvDate(2) & ", " & ServerTime & ",'" & GetTermOfPayment(strSameNo, "" & adoquery.Fields("a1703").Value, strCompany, "" & adoquery.Fields("a1718")) & "', '" & ChgSQL("" & adoquery.Fields("a1717").Value) & "' )"
               End If
'end 2019/7/16

insertdata:
               '2011/9/13 ADD BY SONIA U10005781有二收文號但有一筆無收據,故加入此控制,否則A1907會存NULL
               '此資料若無收據但此編號已存在於ACC190則略過此筆資料
               If IsNull(adoquery.Fields("a0k01").Value) = True Then
                  adoaccrpt217.CursorLocation = adUseClient
                  adoaccrpt217.Open "select * from acc190 where a1902 = '" & adoquery.Fields("a1702").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
                  If adoaccrpt217.RecordCount <> 0 Then
                     adoaccrpt217.Close
                     GoTo NextSkip
                  End If
                  adoaccrpt217.Close
               End If
               '2011/9/13 END
               
               '2013/6/3 add by sonia
               strIndependent = ""
               'CFP案判斷是否為獨立水單
               'Modified by Morgan 2019/3/15 +加判斷a1526
'Modified by Morgan 2019/7/16 是否獨立水單改判斷 A1720
'               If IsNull(adoquery.Fields("caseno").Value) = False Then
'                  '2014/4/1 CANCEL BY SONIA 所有國外案都要(原只做CFP)
'                  adopatent.CursorLocation = adUseClient
'                  'modify by sonia 2014/7/21 四個基本檔都抓
'                  'If Left(adoquery.Fields("caseno").Value, 3) = "CFP" Or Left(adoquery.Fields("caseno").Value, 3) = "P" Then
'                  '   StrSQLa = "select nvl(p1.pa26,p2.pa26) pa26 from acc170, acc151, patent p1, acc161, patent p2 where a1702='" & strSameNo & "' and a1702=axf01(+) and a1702=axg01(+) " & _
'                  '             "and substr(axf03, 1, length(axf03) - 9)=p1.pa01(+) and substr(axf03, length(axf03) - 8, 6)=p1.pa02(+) and substr(axf03, length(axf03) - 2, 1)=p1.pa03(+) and substr(axf03, length(axf03) - 1, 2)=p1.pa04(+) " & _
'                  '             "and substr(axg03, 1, length(axg03) - 9)=p2.pa01(+) and substr(axg03, length(axg03) - 8, 6)=p2.pa02(+) and substr(axg03, length(axg03) - 2, 1)=p2.pa03(+) and substr(axg03, length(axg03) - 1, 2)=p2.pa04(+) "
'                  'Else
'                  '   StrSQLa = "select nvl(t1.tm23,t2.tm23) pa26 from acc170, acc151, trademark t1, acc161, trademark t2 where a1702='" & strSameNo & "' and a1702=axf01(+) and a1702=axg01(+) " & _
'                  '             "and substr(axf03, 1, length(axf03) - 9)=t1.tm01(+) and substr(axf03, length(axf03) - 8, 6)=t1.tm02(+) and substr(axf03, length(axf03) - 2, 1)=t1.tm03(+) and substr(axf03, length(axf03) - 1, 2)=t1.tm04(+) " & _
'                  '             "and substr(axg03, 1, length(axg03) - 9)=t2.tm01(+) and substr(axg03, length(axg03) - 8, 6)=t2.tm02(+) and substr(axg03, length(axg03) - 2, 1)=t2.tm03(+) and substr(axg03, length(axg03) - 1, 2)=t2.tm04(+) "
'                  'End If
'                  Select Case Left(adoquery.Fields("caseno").Value, Len(adoquery.Fields("caseno").Value) - 9)
'                     Case "CFP", "P"
'                        StrSQLa = "select nvl(p1.pa26,p2.pa26) pa26,a1526 from acc170, acc151,acc150, patent p1, acc161, patent p2 where a1702='" & strSameNo & "' and a1702=axf01(+) and a1501(+)=axf01 and a1702=axg01(+) " & _
'                                  "and substr(axf03, 1, length(axf03) - 9)=p1.pa01(+) and substr(axf03, length(axf03) - 8, 6)=p1.pa02(+) and substr(axf03, length(axf03) - 2, 1)=p1.pa03(+) and substr(axf03, length(axf03) - 1, 2)=p1.pa04(+) " & _
'                                  "and substr(axg03, 1, length(axg03) - 9)=p2.pa01(+) and substr(axg03, length(axg03) - 8, 6)=p2.pa02(+) and substr(axg03, length(axg03) - 2, 1)=p2.pa03(+) and substr(axg03, length(axg03) - 1, 2)=p2.pa04(+) "
'                     Case "TF", "T", "CFT"
'                        StrSQLa = "select nvl(t1.tm23,t2.tm23) pa26,a1526 from acc170, acc151,acc150, trademark t1, acc161, trademark t2 where a1702='" & strSameNo & "' and a1702=axf01(+) and a1501(+)=axf01 and a1702=axg01(+) " & _
'                                  "and substr(axf03, 1, length(axf03) - 9)=t1.tm01(+) and substr(axf03, length(axf03) - 8, 6)=t1.tm02(+) and substr(axf03, length(axf03) - 2, 1)=t1.tm03(+) and substr(axf03, length(axf03) - 1, 2)=t1.tm04(+) " & _
'                                  "and substr(axg03, 1, length(axg03) - 9)=t2.tm01(+) and substr(axg03, length(axg03) - 8, 6)=t2.tm02(+) and substr(axg03, length(axg03) - 2, 1)=t2.tm03(+) and substr(axg03, length(axg03) - 1, 2)=t2.tm04(+) "
'                     Case "L", "CFL"
'                        StrSQLa = "select nvl(L1.LC11,L2.LC11) pa26,a1526 from acc170, acc151,acc150, LAWCASE L1, acc161, LAWCASE L2 where a1702='" & strSameNo & "' and a1702=axf01(+) and a1501(+)=axf01 and a1702=axg01(+) " & _
'                                  "and substr(axf03, 1, length(axf03) - 9)=L1.LC01(+) and substr(axf03, length(axf03) - 8, 6)=L1.LC02(+) and substr(axf03, length(axf03) - 2, 1)=L1.LC03(+) and substr(axf03, length(axf03) - 1, 2)=L1.LC04(+) " & _
'                                  "and substr(axg03, 1, length(axg03) - 9)=L2.LC01(+) and substr(axg03, length(axg03) - 8, 6)=L2.LC02(+) and substr(axg03, length(axg03) - 2, 1)=L2.LC03(+) and substr(axg03, length(axg03) - 1, 2)=L2.LC04(+) "
'                     Case Else
'                        StrSQLa = "select nvl(s1.sp08,s2.sp08) pa26,a1526 from acc170, acc151,acc150, servicepractice s1, acc161, servicepractice s2 where a1702='" & strSameNo & "' and a1702=axf01(+) and a1501(+)=axf01 and a1702=axg01(+) " & _
'                                  "and substr(axf03, 1, length(axf03) - 9)=s1.sp01(+) and substr(axf03, length(axf03) - 8, 6)=s1.sp02(+) and substr(axf03, length(axf03) - 2, 1)=s1.sp03(+) and substr(axf03, length(axf03) - 1, 2)=s1.sp04(+) " & _
'                                  "and substr(axg03, 1, length(axg03) - 9)=s2.sp01(+) and substr(axg03, length(axg03) - 8, 6)=s2.sp02(+) and substr(axg03, length(axg03) - 2, 1)=s2.sp03(+) and substr(axg03, length(axg03) - 1, 2)=s2.sp04(+) "
'                  End Select
'                  'end 2014/7/21
'                  adopatent.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'                  If adopatent.RecordCount <> 0 Then
'                     'Added by Morgan 2019/3/15
'                     If adopatent.Fields("a1526").Value = "Y" Then
'                        strIndependent = "Y"
'                        adoTaie.Execute "update acc180 set a1812='Y' where a1801='" & strNo & "'"
'                     'end 2019/3/15
'                     ElseIf IsNull(adopatent.Fields("pa26").Value) = False Then
'                        '2013/3/12 MODIFY BY SONIA 加入 X6383801中國醫藥大學
'                        'MODIFY BY SONIA 2014/3/28 顏永堅3/25郵件所列大學及其關係企業都要加
'                        'modify by sonia 2017/4/10 婉莘要求加X69534彩豐精技   2017/5/15陳德發及郭雅娟要求取消
'                        'modify by sonia 2017/4/27 高國碩要求X60149改為只要母號,關係企業不要
'                        If Left(adopatent.Fields("pa26").Value, 6) = "X44551" Or Left(adopatent.Fields("pa26").Value, 6) = "X62079" _
'                        Or Left(adopatent.Fields("pa26").Value, 6) = "X43988" Or Left(adopatent.Fields("pa26").Value, 6) = "X63219" Or Left(adopatent.Fields("pa26").Value, 6) = "X62319" _
'                        Or Left(adopatent.Fields("pa26").Value, 6) = "X60498" Or Left(adopatent.Fields("pa26").Value, 6) = "X62702" Or Left(adopatent.Fields("pa26").Value, 6) = "X63838" _
'                         Then
'                           strIndependent = "Y"
'                           adoTaie.Execute "update acc180 set a1812='Y' where a1801='" & strNo & "'"  '2014/7/21 add by sonia 獨立水單,水單合計頁不加計
'                        End If
'                        'add by sonia 2017/4/14 單獨編號而關係企業不要的,例:王副總的華碩客戶(X69011010)
'                        'modify by sonia 2017/4/27 高國碩要求X60149改為只要母號,關係企業不要
'                        'modify by sonia 2017/10/11 張詠翔要求取消X6014900
'                        'modify by sonia 2017/10/26 郭雅娟要求加X60738010國立清華大學
'                        If Left(adopatent.Fields("pa26").Value, 8) = "X6901101" Or Left(adopatent.Fields("pa26").Value, 8) = "X6073801" Then
'                           strIndependent = "Y"
'                           adoTaie.Execute "update acc180 set a1812='Y' where a1801='" & strNo & "'"  '2014/7/21 add by sonia 獨立水單,水單合計頁不加計
'                        End If
'                        'end 2017/4/14
'                     End If
'                  End If
'                  adopatent.Close
'               End If
'               '2013/6/3 end

               'Modified by Morgan 2022/8/23
               'If Not IsNull(adoquery.Fields("A1720")) Then
               If Not IsNull(adoquery.Fields("A1720")) Or strFagent = "Y55766000" Then
               'end 2022/8/23
                  strIndependent = "Y"
                  adoTaie.Execute "update acc180 set a1812='Y' where a1801='" & strNo & "'"
               End If
'end 2019/7/16
               
               adoTaie.Execute "delete from acc190 where a1902 = '" & adoquery.Fields("a1702").Value & "'"
               If IsNull(adoquery.Fields("a0k04").Value) Then
                  adoTaie.Execute "insert into acc190 (a1901, a1902, a1903, a1904, a1905, a1906, a1915, a1911, a1909, a1910, a1907, a1916, a1917) values ('" & strNo & "', '" & adoquery.Fields("a1702").Value & "', '" & adoquery.Fields("a1703").Value & "', " & Val(adoquery.Fields("a1704").Value) & ", 0, 0, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strUserNum & "', " & strSrvDate(2) & ", " & ServerTime & ", '', '" & strCompanyType & "','" & strCompany & "')"
               Else
                  adoTaie.Execute "insert into acc190 (a1901, a1902, a1903, a1904, a1905, a1906, a1915, a1911, a1909, a1910, a1907, a1916, a1917) values ('" & strNo & "', '" & adoquery.Fields("a1702").Value & "', '" & adoquery.Fields("a1703").Value & "', " & Val(adoquery.Fields("a1704").Value) & ", 0, 0, " & Val(FCDate(MaskEdBox1.Text)) & ", '" & strUserNum & "', " & strSrvDate(2) & ", " & ServerTime & ", '" & ChgSQL(adoquery.Fields("a0k04").Value) & "', '" & strCompanyType & "','" & strCompany & "')"
               End If
               adoTaie.Execute "update acc170 set a1709 = '" & strNo & "', a1708 = " & strSrvDate(2) & " where a1701 = '" & adoquery.Fields("a1701").Value & "' and a1702 = '" & adoquery.Fields("a1702").Value & "'"
               
               strCaseNo = "" & adoquery.Fields("caseno").Value  '2013/2/20 ADD BY SONIA
            End If
'            End If
'            adoquery.Close
NextSkip:
            adoquery.MoveNext
         Loop
         adoquery.Close
        'Added by Lydia 2016/05/11 記錄執行時間,供台銀結匯資料比對,請參考GetTermOfPayment
        If bolAddLog Then
           PUB_AddExcuteLog "Frmacc2170"
        End If
        
        'Added by Lydia 2017/09/22  更新付款單的匯款方式=>5.台銀合併結匯
        If strA1801List <> "" Then
            PUB_UpdateA1811toType strA1801List
        End If
        'end 2017/09/22
        
        Exit Sub 'Added by Lydia 2016/08/15 改成Printer輸出
        
'Removed by Morgan 2019/7/16 沒用,上註解以免又改到
'' 國內未收款明細表
'         adoTaie.Execute "delete from accrpt216"
'         adoaccrpt216.CursorLocation = adUseClient
'         adoaccrpt216.Open "select * from accrpt216", adoTaie, adOpenDynamic, adLockBatchOptimistic
'         adocaseprogress.CursorLocation = adUseClient
'         'Ken 91/12/23 改為未結匯全印
'         '93.11.23 MODIFY BY SONIA 改抓CASEPROGRESS之收款記錄, 另部分銷帳未銷部分已收者也不印
'         '93.12.8 MODIFY BY SONIA 加入國外請款資料
'         '2005/5/13 MODIFY BY SONIA 再加未請款資料
'         'Modify by Morgan 2006/7/27 國外的改判斷未結清
'         '2007/3/2 modify by sonia 加入 cp87,cp88
'         '2011/8/17 modify by sonia 取消結匯日期欄的判斷,因為程式沒有更新此欄
'         StrSQLa = "select * from (select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505 from caseprogress, acc0k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='E' AND cp60 = a0k01 (+) and cp61 = a1501 and cp61 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0)) or (a0k01 is null)) and a1908 is null union " & _
'                                  "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505 from caseprogress, acc0k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='E' AND cp60 = a0k01 (+) and cp62 = a1501 and cp62 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0)) or (a0k01 is null)) and a1908 is null union " & _
'                                  "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505 from caseprogress, acc0k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='E' AND cp60 = a0k01 (+) and cp63 = a1501 and cp63 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0)) or (a0k01 is null)) and a1908 is null union " & _
'                                  "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505 from caseprogress, acc0k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='E' AND cp60 = a0k01 (+) and cp87 = a1501 and cp87 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0)) or (a0k01 is null)) and a1908 is null union " & _
'                                  "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505 from caseprogress, acc0k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='E' AND cp60 = a0k01 (+) and cp88 = a1501 and cp88 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0)) or (a0k01 is null)) and a1908 is null union " & _
'                                  "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505 from caseprogress, acc0k0, acc150, acc170, acc190 where CP60 IS NULL AND cp60 = a0k01 (+) and cp61 = a1501 and cp61 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0))) and a1908 is null union " & _
'                                  "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505 from caseprogress, acc0k0, acc150, acc170, acc190 where CP60 IS NULL AND cp60 = a0k01 (+) and cp62 = a1501 and cp62 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0))) and a1908 is null union " & _
'                                  "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505 from caseprogress, acc0k0, acc150, acc170, acc190 where CP60 IS NULL AND cp60 = a0k01 (+) and cp63 = a1501 and cp63 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0))) and a1908 is null union " & _
'                                  "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505 from caseprogress, acc0k0, acc150, acc170, acc190 where CP60 IS NULL AND cp60 = a0k01 (+) and cp87 = a1501 and cp87 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0))) and a1908 is null union " & _
'                                  "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505 from caseprogress, acc0k0, acc150, acc170, acc190 where CP60 IS NULL AND cp60 = a0k01 (+) and cp88 = a1501 and cp88 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0))) and a1908 is null union " & _
'                                  "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A1K01,A1K02,A1501,A1503,A1505 from caseprogress, acc1k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='X' AND cp60 = a1k01 (+) and cp61 = a1501 and cp61 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and a1k29 is null and a1908 is null union " & _
'                                  "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A1K01,A1K02,A1501,A1503,A1505 from caseprogress, acc1k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='X' AND cp60 = a1k01 (+) and cp62 = a1501 and cp62 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and a1k29 is null and a1908 is null union " & _
'                                  "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A1K01,A1K02,A1501,A1503,A1505 from caseprogress, acc1k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='X' AND cp60 = a1k01 (+) and cp63 = a1501 and cp63 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and a1k29 is null and a1908 is null union " & _
'                                  "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A1K01,A1K02,A1501,A1503,A1505 from caseprogress, acc1k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='X' AND cp60 = a1k01 (+) and cp87 = a1501 and cp87 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and a1k29 is null and a1908 is null union " & _
'                                  "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A1K01,A1K02,A1501,A1503,A1505 from caseprogress, acc1k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='X' AND cp60 = a1k01 (+) and cp88 = a1501 and cp88 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and a1k29 is null and a1908 is null) new " & _
'                                  "order by cp44 asc"
'
'         '2007/3/2 end
'         '2005/5/13 END
'         '93.12.8 END
'         '93.11.23 END
'         adocaseprogress.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'         Do While adocaseprogress.EOF = False
'            adoaccrpt216.AddNew
'            adoaccrpt216.Fields("r21601").Value = strUserNum
'            If IsNull(adocaseprogress.Fields("a1503").Value) = False Then
'               adoaccrpt216.Fields("r21602").Value = adocaseprogress.Fields("a1503").Value
'               adoaccrpt216.Fields("r21603").Value = FagentQuery(adocaseprogress.Fields("a1503").Value, 2)
'               If adoaccrpt216.Fields("r21603").Value = "" Then
'                  adoaccrpt216.Fields("r21603").Value = FagentQuery(adocaseprogress.Fields("a1503").Value, 1)
'               End If
'               If adoaccrpt216.Fields("r21603").Value = "" Then
'                  adoaccrpt216.Fields("r21603").Value = FagentQuery(adocaseprogress.Fields("a1503").Value, 3)
'               End If
'            End If
'            adoaccrpt216.Fields("r21604").Value = adocaseprogress.Fields("a1501").Value
'            If IsNull(adocaseprogress.Fields("a1505").Value) = False Then
'               adoaccrpt216.Fields("r21605").Value = adocaseprogress.Fields("a1505").Value
'            End If
'            If adoquery.State = adStateOpen Then
'               adoquery.Close
'            End If
'            adoquery.CursorLocation = adUseClient
'            adoquery.Open "select * from acc151 where axf01 = '" & adocaseprogress.Fields("a1501").Value & "' and axf02 = '" & adocaseprogress.Fields("cp09").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'            If adoquery.RecordCount <> 0 Then
'               If IsNull(adoquery.Fields("axf04").Value) = False Then
'                  adoaccrpt216.Fields("r21606").Value = adoquery.Fields("axf04").Value
'               Else
'                  adoaccrpt216.Fields("r21606").Value = 0
'               End If
'            Else
'               adoaccrpt216.Fields("r21606").Value = 0
'            End If
'            adoquery.Close
'            If IsNull(adocaseprogress.Fields("cp01").Value) = False Then
'               adoaccrpt216.Fields("r21607").Value = adocaseprogress.Fields("cp01").Value
'               If IsNull(adocaseprogress.Fields("cp02").Value) = False Then
'                  adoaccrpt216.Fields("r21607").Value = adoaccrpt216.Fields("r21607").Value & adocaseprogress.Fields("cp02").Value
'               End If
'               If IsNull(adocaseprogress.Fields("cp03").Value) = False Then
'                  adoaccrpt216.Fields("r21607").Value = adoaccrpt216.Fields("r21607").Value & adocaseprogress.Fields("cp03").Value
'               End If
'               If IsNull(adocaseprogress.Fields("cp04").Value) = False Then
'                  adoaccrpt216.Fields("r21607").Value = adoaccrpt216.Fields("r21607").Value & adocaseprogress.Fields("cp04").Value
'               End If
'            End If
'            adoquery.CursorLocation = adUseClient
'            Select Case adocaseprogress.Fields("cp01").Value
'               Case "CFP", "FCP", "P"
'                  adoquery.Open "select pa26 from patent where pa01 = '" & adocaseprogress.Fields("cp01").Value & "' and pa02 = '" & adocaseprogress.Fields("cp02").Value & "' and pa03 = '" & adocaseprogress.Fields("cp03").Value & "' and pa04 = '" & adocaseprogress.Fields("cp04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'               Case "CFT", "FCT", "T"
'                  adoquery.Open "select tm23 from trademark where tm01 = '" & adocaseprogress.Fields("cp01").Value & "' and tm02 = '" & adocaseprogress.Fields("cp02").Value & "' and tm03 = '" & adocaseprogress.Fields("cp03").Value & "' and tm04 = '" & adocaseprogress.Fields("cp04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'               Case "CFL", "FCL", "L"
'                  adoquery.Open "select lc11 from lawcase where lc01 = '" & adocaseprogress.Fields("cp01").Value & "' and lc02 = '" & adocaseprogress.Fields("cp02").Value & "' and lc03 = '" & adocaseprogress.Fields("cp03").Value & "' and lc04 = '" & adocaseprogress.Fields("cp04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'               Case Else
'                  adoquery.Open "select sp08 from servicepractice where sp01 = '" & adocaseprogress.Fields("cp01").Value & "' and sp02 = '" & adocaseprogress.Fields("cp02").Value & "' and sp03 = '" & adocaseprogress.Fields("cp03").Value & "' and sp04 = '" & adocaseprogress.Fields("cp04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'            End Select
'            If adoquery.RecordCount <> 0 Then
'               If IsNull(adocaseprogress.Fields(0).Value) = False Then
'                  adoaccrpt216.Fields("r21608").Value = adoquery.Fields(0).Value
'               End If
'            End If
'            adoquery.Close
'            If IsNull(adocaseprogress.Fields("cp13").Value) = False Then
'               adoaccrpt216.Fields("r21610").Value = adocaseprogress.Fields("cp13").Value
'            End If
'            If IsNull(adocaseprogress.Fields("cp16").Value) = False Then
'               adoaccrpt216.Fields("r21612").Value = Val(adocaseprogress.Fields("cp16").Value)
'            Else
'               adoaccrpt216.Fields("r21612").Value = 0
'            End If
'            If IsNull(adocaseprogress.Fields("cp77").Value) = False Then
'               adoaccrpt216.Fields("r21612").Value = Val(adoaccrpt216.Fields("r21612").Value) - Val(adocaseprogress.Fields("cp77").Value)
'            End If
'            If IsNull(adocaseprogress.Fields("a0k01").Value) = False Then
'               If IsNull(adocaseprogress.Fields("a0k02").Value) = False Then
'                  adoaccrpt216.Fields("r21609").Value = adocaseprogress.Fields("a0k02").Value
'               End If
'               adoaccrpt216.Fields("r21611").Value = adocaseprogress.Fields("a0k01").Value
'               If IsNull(adocaseprogress.Fields("cp75").Value) = False Then
'                  adoaccrpt216.Fields("r21613").Value = Val(adoaccrpt216.Fields("r21612").Value) - Val(adocaseprogress.Fields("cp75").Value)
'               Else
'                  adoaccrpt216.Fields("r21613").Value = Val(adoaccrpt216.Fields("r21612").Value)
'               End If
'               If IsNull(adocaseprogress.Fields("cp78").Value) = False Then
'                  adoaccrpt216.Fields("r21613").Value = Val(adoaccrpt216.Fields("r21613").Value) + Val(adocaseprogress.Fields("cp78").Value)
'               End If
'               strYes = MsgText(601)
'            Else
'               adoquery.CursorLocation = adUseClient
'               adoquery.Open "select * from acc1k0 where a1k01 = '" & adocaseprogress.Fields("cp60").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'               If adoquery.RecordCount <> 0 Then
'                  If IsNull(adoquery.Fields("a1k02").Value) = False Then
'                     adoaccrpt216.Fields("r21609").Value = adoquery.Fields("a1k02").Value
'                  Else
'                     adoaccrpt216.Fields("r21609").Value = Null
'                  End If
'                  adoaccrpt216.Fields("r21611").Value = adoquery.Fields("a1k01").Value
'                  If IsNull(adoquery.Fields("a1k30").Value) = False Then
'                     adoaccrpt216.Fields("r21613").Value = Val(adoaccrpt216.Fields("r21612").Value) - Val(adoquery.Fields("a1k30").Value)
'                  Else
'                     adoaccrpt216.Fields("r21613").Value = Val(adoaccrpt216.Fields("r21612").Value)
'                  End If
'                  If IsNull(adoquery.Fields("a1k29").Value) Then
'                     strYes = MsgText(601)
'                  Else
'                     strYes = MsgText(602)
'                  End If
'               Else
'                  adoaccrpt216.Fields("r21609").Value = Null
'                  adoaccrpt216.Fields("r21611").Value = Null
'                  adoaccrpt216.Fields("r21613").Value = 0
'                  strYes = MsgText(601)
'               End If
'               adoquery.Close
'            End If
'            If strYes = MsgText(602) Then
'               adoaccrpt216.Delete
'            End If
'            adoaccrpt216.UpdateBatch
'            adocaseprogress.MoveNext
'         Loop
'         adoaccrpt216.Close
'         adocaseprogress.Close
'
'
'      Case Mid(ReportTitle(217), 6, 10) '代理人未收未付對照表
'' 代理人未收未付對照表
'         adoTaie.Execute "delete from accrpt217"
'         adoaccrpt217.CursorLocation = adUseClient
'         adoaccrpt217.Open "select * from accrpt217", adoTaie, adOpenDynamic, adLockBatchOptimistic
'         adofagent.CursorLocation = adUseClient
'         'Ken 91/12/23 改為未結匯全印
'         'adofagent.Open "select distinct fa01, fa02, nvl(FA05||FA63||FA64||FA65, nvl(fa04, fa06)) as Name from fagent, acc150, acc170 where fa01 = substr(a1503, 1, 8) and fa02 = substr(a1503, 9, 1) and a1501 = a1702 and a1708 = " & Val(ACDate(ServerDate)) & " order by fa01 asc, fa02 asc", adoTaie, adOpenStatic, adLockReadOnly
'         StrSQLa = "select distinct fa01, fa02, nvl(FA05||FA63||FA64||FA65, nvl(fa04, fa06)) as Name from acc150, fagent, acc170, acc190 where substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and a1501 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and a1908 is null order by fa01 asc, fa02 asc"
'         adofagent.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'         Do While adofagent.EOF = False
'            adoaccrpt217.AddNew
'            adoaccrpt217.Fields("r21701").Value = strUserNum
'            adoaccrpt217.Fields("r21702").Value = adofagent.Fields("fa01").Value & adofagent.Fields("fa02").Value
'            If IsNull(adofagent.Fields("Name").Value) = False Then
'               adoaccrpt217.Fields("r21703").Value = adofagent.Fields("Name").Value
'            End If
'            adoquery.CursorLocation = adUseClient
'            '93.6.4 modify by sonia 已銷帳不抓  - nvl(a1k30, 0)
'            '93.7.19 modify by sonia 改抓外幣金額 a0k08 並扣除折讓   - nvl(a1k06, 0)
'            'Modify By Sindy 2012/12/7
'            'adoquery.Open "select min(a1k02), count(a1k01), sum(a1k08 - (nvl(a1k30, 0)) / decode(a1k10, 0, 1, a1k10) - nvl(a1k06, 0)) from acc1k0 where a1k28 = '" & adofagent.Fields("fa01").Value & adofagent.Fields("fa02").Value & "' and (a1k11 > a1k30 or a1k30 is null or a1k30 = 0) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = 0) and a1k25 is null", adoTaie, adOpenStatic, adLockReadOnly
'            adoquery.Open "select min(a1k02), count(a1k01), sum(a1k08 - (nvl(a1k30, 0)) / decode(a1k10, 0, 1, a1k10) - nvl(a1k31, 0)) from acc1k0 where a1k28 = '" & adofagent.Fields("fa01").Value & adofagent.Fields("fa02").Value & "' and (a1k11 > a1k30 or a1k30 is null or a1k30 = 0) and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = 0) and a1k25 is null", adoTaie, adOpenStatic, adLockReadOnly
'            '2012/12/7 End
'            '93.7.19 end
'            '93.6.4 end
'            If adoquery.RecordCount <> 0 Then
'               If IsNull(adoquery.Fields(0).Value) = False Then
'                  adoaccrpt217.Fields("r21704").Value = adoquery.Fields(0).Value
'               End If
'               If IsNull(adoquery.Fields(1).Value) = False Then
'                  adoaccrpt217.Fields("r21705").Value = adoquery.Fields(1).Value
'               End If
'               If IsNull(adoquery.Fields(2).Value) = False Then
'                  adoaccrpt217.Fields("r21706").Value = Val(Format(adoquery.Fields(2).Value, FAmount))
'               End If
'            End If
'            adoquery.Close
'            adoquery.CursorLocation = adUseClient
'            adoquery.Open "select min(a1502), count(a1501), sum((a1506 - nvl(a1520, 0))), min(a1505) from acc150, acc170 where a1501 = a1702 and a1503 = '" & adofagent.Fields("fa01").Value & adofagent.Fields("fa02").Value & "' and (a1506 > a1520 or a1520 is null or a1520 = 0)", adoTaie, adOpenStatic, adLockReadOnly
'            If adoquery.RecordCount <> 0 Then
'               If IsNull(adoquery.Fields(0).Value) = False Then
'                  adoaccrpt217.Fields("r21707").Value = adoquery.Fields(0).Value
'               End If
'               If IsNull(adoquery.Fields(1).Value) = False Then
'                  adoaccrpt217.Fields("r21708").Value = adoquery.Fields(1).Value
'               End If
'               If IsNull(adoquery.Fields(2).Value) = False Then
'                  adoaccrpt217.Fields("r21709").Value = adoquery.Fields(2).Value
'               End If
'               If IsNull(adoquery.Fields(3).Value) = False Then
'                  adoaccrpt217.Fields("r21710").Value = adoquery.Fields(3).Value
'               End If
'            End If
'            adoquery.Close
'            adoaccrpt217.UpdateBatch
'            adofagent.MoveNext
'         Loop
'         adoaccrpt217.Close
'         adofagent.Close
'         adoTaie.Execute "delete from accrpt217 where r21705 = 0 or r21708 = 0"
'
'
'      Case Mid(ReportTitle(218), 6, 6) '付款明細草稿
'' 付款明細草稿
'         'Added by Lydia 2015/03/18 +鎖定使用者
'         'modify by sonia 2015/9/14 因為accreport無法抓使用者
'         'adoTaie.Execute "delete from accrpt218 where R21801='" & strUserNum & "' "
'         adoTaie.Execute "delete from accrpt218 "
'         adoaccrpt218.CursorLocation = adUseClient
'         'adoaccrpt218.Open "select * from accrpt218 where R21801='" & strUserNum & "' ", adoTaie, adOpenDynamic, adLockBatchOptimistic
'         adoaccrpt218.Open "select * from accrpt218 ", adoTaie, adOpenDynamic, adLockBatchOptimistic
'         adoacc180.CursorLocation = adUseClient
''         adoacc180.Open "select * from acc190, acc180 where a1901 = a1801 and a1802 = " & Val(FCDate(MaskEdBox1.Text)) & " and substr(a1903, 1, 2) <> 'US'", adoTaie, adOpenStatic, adLockReadOnly
'         StrSQLa = "select * from acc190, acc180 where a1901 = a1801 and a1908 is null"
'         adoacc180.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'         Do While adoacc180.EOF = False
'            adoaccrpt218.AddNew
'            adoaccrpt218.Fields("r21801").Value = strUserNum
'            adoaccrpt218.Fields("r21802").Value = adoacc180.Fields("a1901").Value
'            If IsNull(adoacc180.Fields("a1803").Value) = False Then
'               adoaccrpt218.Fields("r21803").Value = adoacc180.Fields("a1803").Value
'            End If
'            '2012/11/12 add by sonia 若該代理人三個月之內未曾匯過款,於代理人編號後加◎
'            'Added by Lydia 2015/06/12 +台銀電匯紙本
'            'Modified by Lydia 2017/10/03 + 4.華銀電匯紙本,5.台銀合併結匯
'            'If "" & adoacc180.Fields("a1811").Value = "2" Or "" & adoacc180.Fields("a1811").Value = "3" Then '2013/2/7 ADD BY SONIA 加入票匯才要檢查
'            If InStr("2,3,4,5", "" & adoacc180.Fields("a1811").Value) > 0 Then
'               adopatent.CursorLocation = adUseClient
'               'Modified by Lydia 2017/10/11 判斷幣別
'               'StrSQLa = "select a1b03 from acc1b0 where a1b02 = '" & adoacc180.Fields("a1803").Value & "' and a1b03> " & ChangeWStringToTString(CompDate(1, -3, strSrvDate(2))) & _
'                         " union select a1i03 from acc1i0,acc150 where a1i03> " & ChangeWStringToTString(CompDate(1, -3, strSrvDate(2))) & " and a1i01=a1512(+) and a1503 = '" & adoacc180.Fields("a1803").Value & "'"
'                         'select c.* from acc1b0 a,acc1c0 b,acc170 c where a1b02 = 'Y20757010' and a1b03> 1060701 and a1b01=a1c01 and a1b02=a1c02 and a1c03=a1702 and a1703='JPY'
'               StrSQLa = "select a1b03 from acc1b0, acc1c0, acc170 where a1b02 = '" & adoacc180.Fields("a1803").Value & "' and a1b03> " & ChangeWStringToTString(CompDate(1, -3, strSrvDate(2))) & _
'                         " and a1b01=a1c01 and a1b02=a1c02 and a1c03=a1702 and a1703=" & CNULL(adoacc180.Fields("a1903")) & _
'                         " union select a1i03 from acc1i0,acc150 where a1i03> " & ChangeWStringToTString(CompDate(1, -3, strSrvDate(2))) & " and a1i01=a1512(+) and a1503 = '" & adoacc180.Fields("a1803").Value & "'" & _
'                         " and a1i05=" & CNULL(adoacc180.Fields("a1903"))
'               adopatent.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'               If adopatent.RecordCount = 0 And Left(adoacc180.Fields("a1803").Value, 1) = "Y" Then
'                  adoaccrpt218.Fields("r21803").Value = adoaccrpt218.Fields("r21803").Value & "◎"
'               End If
'               adopatent.Close
'            End If
'            '2012/11/12 end
'            If IsNull(adoacc180.Fields("a1902").Value) = False Then
'               adoaccrpt218.Fields("r21804").Value = adoacc180.Fields("a1902").Value
'            End If
'            If IsNull(adoacc180.Fields("a1903").Value) = False Then
'               adoaccrpt218.Fields("r21805").Value = adoacc180.Fields("a1903").Value
'            End If
'            If IsNull(adoacc180.Fields("a1904").Value) = False Then
'               adoaccrpt218.Fields("r21806").Value = adoacc180.Fields("a1904").Value
'            Else
'               adoaccrpt218.Fields("r21806").Value = 0
'            End If
'            If IsNull(adoacc180.Fields("a1907").Value) = False Then
'               adoaccrpt218.Fields("r21807").Value = adoacc180.Fields("a1907").Value
'            End If
'            '2012/10/22 add by sonia X63219國立中正大學 及 X43988060國立虎尾科技大學 的CFP案水單要個案單獨出並在此報表的 A1907國內客戶(收據抬頭) 之後加註 '不能合併'
'            '2013/3/12 MODIFY BY SONIA 加入 X6383801中國醫藥大學
'            '2014/4/1 CANCEL BY SONIA 所有國外案都要(原只做CFP),顏永堅3/25郵件所列大學及其關係企業都要加
'            adopatent.CursorLocation = adUseClient
'            StrSQLa = "select nvl(axf03,axg03) caseno,nvl(p1.pa26,p2.pa26) pa26 from acc190, acc151, patent p1, acc161, patent p2 where a1901 = '" & adoacc180.Fields("a1901").Value & "' and a1902=axf01(+) and a1902=axg01(+) " & _
'                      "and substr(axf03, 1, length(axf03) - 9)=p1.pa01(+) and substr(axf03, length(axf03) - 8, 6)=p1.pa02(+) and substr(axf03, length(axf03) - 2, 1)=p1.pa03(+) and substr(axf03, length(axf03) - 1, 2)=p1.pa04(+) " & _
'                      "and substr(axg03, 1, length(axg03) - 9)=p2.pa01(+) and substr(axg03, length(axg03) - 8, 6)=p2.pa02(+) and substr(axg03, length(axg03) - 2, 1)=p2.pa03(+) and substr(axg03, length(axg03) - 1, 2)=p2.pa04(+) " & _
'                "union select nvl(axf03,axg03) caseno,nvl(t1.tm23,t2.tm23) pa26 from acc190, acc151, trademark t1, acc161, trademark t2 where a1901 = '" & adoacc180.Fields("a1901").Value & "' and a1902=axf01(+) and a1902=axg01(+) " & _
'                      "and substr(axf03, 1, length(axf03) - 9)=t1.tm01(+) and substr(axf03, length(axf03) - 8, 6)=t1.tm02(+) and substr(axf03, length(axf03) - 2, 1)=t1.tm03(+) and substr(axf03, length(axf03) - 1, 2)=t1.tm04(+) " & _
'                      "and substr(axg03, 1, length(axg03) - 9)=t2.tm01(+) and substr(axg03, length(axg03) - 8, 6)=t2.tm02(+) and substr(axg03, length(axg03) - 2, 1)=t2.tm03(+) and substr(axg03, length(axg03) - 1, 2)=t2.tm04(+) "
'            adopatent.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'            'modify by sonia 2017/4/10 婉莘要求加X69534彩豐精技   2017/5/15陳德發及郭雅娟要求取消
'            'modify by sonia 2017/4/27 高國碩要求X60149改為只要母號,關係企業不要
'            If IsNull(adopatent.Fields("caseno").Value) = False And IsNull(adopatent.Fields("pa26").Value) = False Then
'               If Left(adopatent.Fields("pa26").Value, 6) = "X44551" Or Left(adopatent.Fields("pa26").Value, 6) = "X62079" _
'               Or Left(adopatent.Fields("pa26").Value, 6) = "X43988" Or Left(adopatent.Fields("pa26").Value, 6) = "X63219" Or Left(adopatent.Fields("pa26").Value, 6) = "X62319" _
'               Or Left(adopatent.Fields("pa26").Value, 6) = "X60498" Or Left(adopatent.Fields("pa26").Value, 6) = "X62702" Or Left(adopatent.Fields("pa26").Value, 6) = "X63838" _
'                Then
'                  adoaccrpt218.Fields("r21807").Value = Left(adoaccrpt218.Fields("r21807").Value, 20) & "   水單不能合併 !"
'               End If
'               'add by sonia 2017/4/14 單獨編號而關係企業不要的,例:王副總的華碩客戶(X69011010)
'               'modify by sonia 2017/4/27 高國碩要求X60149改為只要母號,關係企業不要
'               'modify by sonia 2017/10/11 張詠翔要求取消X6014900
'               'modify by sonia 2017/10/26 郭雅娟要求加X60738010國立清華大學
'               If Left(adopatent.Fields("pa26").Value, 8) = "X6901101" Or Left(adopatent.Fields("pa26").Value, 8) = "X6073801" Then
'                  adoaccrpt218.Fields("r21807").Value = Left(adoaccrpt218.Fields("r21807").Value, 20) & "   水單不能合併 !"
'               End If
'               'end 2017/4/14
'            End If
'            adopatent.Close
'            '2012/10/22 end
'            'Add By Cheng 2003/05/15
'            '公司別
'            adoaccrpt218.Fields("r21808").Value = "" & adoacc180.Fields("a1917").Value
'            adoaccrpt218.UpdateBatch
'            adoacc180.MoveNext
'         Loop
'         adoaccrpt218.Close
'         adoacc180.Close
'
'end 2019/7/16

   End Select


Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   Resume Next
End Sub
'Remove by Lydia 2016/08/03
''Add by Sindy 2010/8/30
'Private Sub AddPrinter()
'   Dim i As Integer
'   For i = 0 To Printers.Count - 1
'      cboPrinters.AddItem Printers(i).DeviceName, i
'      If Printers(i).DeviceName = Printer.DeviceName Then m_iDefaultPrinter = i
'   Next i
'   cboPrinters.ListIndex = m_iDefaultPrinter
'End Sub

Private Sub cboPrinters_Click()
   'Remove by Morgan 2010/10/5 此處不可設定,否則會因改變程式預設印表機而影響後續的報表列印
   'Set Printer = Printers(cboPrinters.ListIndex)
End Sub
'2010/8/30 End

'Added by Lydia 2016/08/15 設定印表機
Private Sub SettingPrtSet()
Dim inX As Integer
Dim tmpArr As Variant, tmpArr2 As Variant

    '設定印表機
     Printer.EndDoc
     Printer.PaperSize = 9  'A4
     '付款明細草稿
     If Combo1 = Mid(ReportTitle(218), 6, 6) Then
         Printer.Orientation = 1 '1.直印
     Else
         Printer.Orientation = 2 '2.橫印
     End If
     
     lngPageHeight = Printer.ScaleHeight
     lngPageWidth = Printer.ScaleWidth
     lngLineHeight = 300
     Printer.Font.Name = "新細明體"
     Printer.Font.Size = ciFontSize
     Erase PLeft
     Erase PTitle
     tmpArr = Empty: tmpArr2 = Empty
     
     '設定欄位抬頭和位置
     If strTitle <> "" And strTitle2 <> "" Then
        tmpArr = Split(strTitle, ",")
        tmpArr2 = Split(strTitle2, ",")
        For inX = 0 To UBound(tmpArr)
            If Trim(tmpArr(inX)) <> "" And Trim(tmpArr2(inX)) <> "" Then
                If Trim(tmpArr(inX)) <> "結束" Then PTitle(inX) = Trim(tmpArr(inX))
                
                If inX < 1 Then
                   PLeft(inX) = ciStartX
                Else
                   PLeft(inX) = PLeft(inX - 1) + Printer.TextWidth(String(Val(tmpArr2(inX)), "　")) + ciColGap
                End If
                
                If Trim(tmpArr(inX)) = "結束" Then Exit For
            End If
        Next
     End If
     
     iPage = 0
     
End Sub

'Added by Lydia 2016/08/15 換行判斷
Private Sub PrintNewLine(Optional ByVal mRate As Single = 1, Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 4)
   iPrint = iPrint + lngLineHeight * mRate
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint

      iPage = iPage + 1
      Printer.NewPage
      PrintHeader
   End If
End Sub

'Added by Lydia 2016/08/15 列印表頭
Private Sub PrintHeader()
Dim x1 As Integer
Dim x2 As Integer
Dim iPos As Integer

iPrint = ciStartY
iPageLine = 0

Printer.Font.Size = ciTitleFontSize
Printer.Font.Bold = True
Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(mRptTitle)) / 2
Printer.CurrentY = iPrint
Printer.Print mRptTitle

Printer.Font.Size = ciFontSize
PrintNewLine
PrintNewLine

Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "列印人員：" & strUserName
x1 = Printer.ScaleWidth - Printer.TextWidth(String(12, "　"))
Printer.CurrentX = x1
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

PrintNewLine
'付款明細草稿
If Combo1 = Mid(ReportTitle(218), 6, 6) Then
    Printer.CurrentX = PLeft(0)
    Printer.CurrentY = iPrint
    Printer.Print "有◎者請確認該代理人銀行帳號是否有更新！ 有＊者為ｅ化帳單！"
    
    'Added by Morgan 2019/3/15
    Printer.CurrentX = PLeft(0)
    Printer.CurrentY = iPrint + Printer.TextHeight("有")
    Printer.Print "獨立水單結匯註記, 1. 承辦確認""獨立水單!"" 2. 系統設定客戶""@"""
    'end 2019/3/15
End If
Printer.CurrentX = x1
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & Printer.Page
Printer.Font.Bold = False
PrintNewLine
PrintNewLine

'代理人未收未付對照表
If Combo1 = Mid(ReportTitle(217), 6, 10) Then
    Printer.CurrentX = PLeft(2) + ciColGap
    Printer.CurrentY = iPrint
    Printer.Print String(24, "-") & " 未收請款單 " & String(24, "-")
    Printer.CurrentX = PLeft(5) + ciColGap
    Printer.CurrentY = iPrint
    Printer.Print String(28, "-") & " 未收帳單 " & String(28, "-")
    PrintNewLine
End If

'列印欄位抬頭
For iPos = 0 To cInX
    If PTitle(iPos) <> "" And PTitle(iPos) <> "結束" Then
       If InStr(PTitle(iPos), "金額") > 0 Then '置中
          x2 = PLeft(iPos) + (PLeft(iPos + 1) - PLeft(iPos) - Printer.TextWidth(PTitle(iPos))) / 2
       Else
          x2 = PLeft(iPos)
       End If
       Printer.CurrentX = x2 'PLeft(iPos)
       Printer.CurrentY = iPrint
       Printer.Print PTitle(iPos)
    ElseIf iPos > 1 Then
        x1 = iPos '結束
        Exit For
    End If
Next
PrintNewLine
Printer.Line (PLeft(0), iPrint)-(PLeft(x1), iPrint)
iPrint = iPrint + 150

End Sub

'Added by Lydia 2016/08/15 列印-付款明細草稿
Private Sub PrintRpt2180()
Dim inP As Integer
Dim rsPrt As New ADODB.Recordset
Dim strGrp As String '小計組群
Dim strGrp1 As String '公司別
Dim strTmp As String
Dim strSubTotal As String '小計

    mRptTitle = ReportTitle(218)
    'Modified by Morgan 2023/3/30
    'strTitle = "付款單號,代理人,單據編號,幣別,金額,國內客戶(收據抬頭),結束"
    'strTitle2 = "0,5,6,5,3,6,18"
    strTitle = "付款單號,代理人　　急件付款日,單據編號,幣別,金額,國內客戶(收據抬頭),結束"
    strTitle2 = "0,5,10,5,3,6,14"
    'end 2023/3/30
    ciFontSize = 12
    
    'Modified by Morgan 2018/1/17 +acc170判斷是否電子結匯a1719
    strSql = "select * from acc190, acc180,acc170 where a1901 = a1801 and a1908 is null and a1702(+)=a1902"
    'Modified by Lydia 2018/06/27 取消台一備註的長度限制(ex. W10701098代理人編號Y20064)
    'strSql = "select X1.*,substrb(X2.a2223,1,30) T1MEMO from (" & strSql & ") X1,acc220 X2 " & _
             "where a1803=a2201(+) and a1903=a2202(+) order by a1917,a1901,a1902 "
    'Modified by Morgan 2023/3/30 +a1527
    strSql = "select X1.*,X2.a2223 T1MEMO,sqldatet(a1527) UDate from (" & strSql & ") X1,acc220 X2,acc150 " & _
             "where a1803=a2201(+) and a1903=a2202(+) and a1501(+)=a1702 order by a1917,a1901,a1902 "
    inP = 1
    Set rsPrt = ClsLawReadRstMsg(inP, strSql)
    If inP = 1 Then
       SettingPrtSet '設定印表機
       With rsPrt
          .MoveFirst
          iPage = iPage + 1
          PrintHeader
          Printer.Font.Size = ciFontSize
          Printer.FontBold = False
          Do While Not .EOF
             If strGrp <> "" & .Fields("a1901").Value Then
                If strGrp <> "" Then '小計
                   Printer.Line (PLeft(4), iPrint)-(PLeft(5) - ciColGap, iPrint)
                   iPrint = iPrint + 150
                   Printer.CurrentX = PLeft(3)
                   Printer.CurrentY = iPrint
                   Printer.Print "小計:"
                   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(strSubTotal, "##,##0.00")) - ciColGap
                   Printer.CurrentY = iPrint
                   Printer.Print Format(strSubTotal, "##,##0.00")
                   iPageLine = iPageLine + 1
                   PrintNewLine
                   PrintNewLine
                   '智權公司換頁
                   If strGrp1 <> "" & .Fields("a1917") And UCase("" & .Fields("a1917").Value) = "J" And iPageLine > 0 Then
                      Printer.NewPage
                      PrintHeader
                   End If
                End If
                strSubTotal = ""
                Printer.CurrentX = PLeft(0)
                Printer.CurrentY = iPrint
                Printer.Print .Fields("a1901").Value & "　公司別：" & .Fields("a1917").Value & "　" & .Fields("T1MEMO") '付款單號R21802+公司別R21808 + 台一備註a2223
                PrintNewLine
             End If
             For inP = 1 To cInX
                If PTitle(inP) = "" Then Exit For
                
                If PTitle(inP) <> "" And PTitle(inP) <> "結束" Then
                   Printer.CurrentX = PLeft(inP)
                   Printer.CurrentY = iPrint
                   Select Case inP
                       Case 1 '代理人 R21803
                            '若該代理人三個月之內未曾匯過款(票匯+台銀電匯紙本),於代理人編號後加◎
                            strTmp = "　"
                            'Modified by Lydia 2017/10/03 + 4.華銀電匯紙本,5.台銀合併結匯
                            'If "" & .Fields("a1811").Value = "2" Or "" & .Fields("a1811").Value = "3" Then
                            If InStr("2,3,4,5", "" & .Fields("a1811").Value) > 0 Then
                               'Modified by Lydia 2017/10/11 判斷幣別
                               'strSql = "select a1b03 from acc1b0 where a1b02 = '" & .Fields("a1803").Value & "' and a1b03> " & ChangeWStringToTString(CompDate(1, -3, strSrvDate(2))) & _
                                        " union select a1i03 from acc1i0,acc150 where a1i03> " & ChangeWStringToTString(CompDate(1, -3, strSrvDate(2))) & " and a1i01=a1512(+) and a1503 = '" & .Fields("a1803").Value & "'"
                                strSql = "select a1b03 from acc1b0, acc1c0, acc170 where a1b02 = '" & .Fields("a1803").Value & "' and a1b03> " & ChangeWStringToTString(CompDate(1, -3, strSrvDate(2))) & _
                                          " and a1b01=a1c01 and a1b02=a1c02 and a1c03=a1702 and a1703=" & CNULL(.Fields("a1903")) & _
                                          " union select a1i03 from acc1i0,acc150 where a1i03> " & ChangeWStringToTString(CompDate(1, -3, strSrvDate(2))) & " and a1i01=a1512(+) and a1503 = '" & .Fields("a1803").Value & "'" & _
                                          " and a1i05=" & CNULL(.Fields("a1903"))
                               intI = 1
                               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                               If intI = 0 And Left(.Fields("a1803").Value, 1) = "Y" Then
                                  strTmp = "◎"
                               End If
                            End If
                            'Modified by Morgan 2023/3/30 +UDate急件付款日
                            Printer.Print .Fields("a1803").Value & strTmp & .Fields("UDate")
                            
                       Case 2 '單據編號 R21804
                            'Modified by Morgan 2018/1/17
                            'Printer.Print "" & .Fields("a1902").Value
                            Printer.Print "" & .Fields("a1902").Value & IIf("" & .Fields("a1719") = "Y", "＊", "　")
                            
                       Case 3 '幣別 R21805
                            Printer.Print "" & .Fields("a1903").Value
                   
                       Case 4 '金額 R21806
                            strTmp = Format(IIf(Val("" & .Fields("a1904").Value) = 0, "0", .Fields("a1904").Value), "##,##0.00")
                            Printer.CurrentX = PLeft(inP + 1) - Printer.TextWidth(strTmp) - ciColGap '靠右
                            Printer.CurrentY = iPrint
                            Printer.Print strTmp
                            
                       Case 5 '國內客戶(收據抬頭) R21807
                            'X6383801中國醫藥大學、X63219國立中正大學 及 X43988060國立虎尾科技大學 的CFP案水單要個案單獨出並在此報表的 A1907國內客戶(收據抬頭) 之後加註 '不能合併'
                            '所有國外案都要(原只做CFP),顏永堅3/25郵件所列大學及其關係企業都要加
                            'Modified by Morgan 2019/3/15 +加判斷a1526
                            'Modified by Moragn 2019/7/15 +a1501
                            strTmp = ""
'Modified by Morgan 2019/7/16 改判斷 A1720
'                            strSql = "select nvl(axf03,axg03) caseno,nvl(p1.pa26,p2.pa26) pa26,a1501,a1526 from acc190, acc151,acc150, patent p1, acc161, patent p2 where a1901 = '" & .Fields("a1901").Value & "' and a1902=axf01(+) and a1501(+)=axf01 and a1902=axg01(+) " & _
'                                      "and substr(axf03, 1, length(axf03) - 9)=p1.pa01(+) and substr(axf03, length(axf03) - 8, 6)=p1.pa02(+) and substr(axf03, length(axf03) - 2, 1)=p1.pa03(+) and substr(axf03, length(axf03) - 1, 2)=p1.pa04(+) " & _
'                                      "and substr(axg03, 1, length(axg03) - 9)=p2.pa01(+) and substr(axg03, length(axg03) - 8, 6)=p2.pa02(+) and substr(axg03, length(axg03) - 2, 1)=p2.pa03(+) and substr(axg03, length(axg03) - 1, 2)=p2.pa04(+) " & _
'                                "union select nvl(axf03,axg03) caseno,nvl(t1.tm23,t2.tm23) pa26,a1501,a1526 from acc190, acc151,acc150, trademark t1, acc161, trademark t2 where a1901 = '" & .Fields("a1901").Value & "' and a1902=axf01(+)and a1501(+)=axf01 and a1902=axg01(+) " & _
'                                      "and substr(axf03, 1, length(axf03) - 9)=t1.tm01(+) and substr(axf03, length(axf03) - 8, 6)=t1.tm02(+) and substr(axf03, length(axf03) - 2, 1)=t1.tm03(+) and substr(axf03, length(axf03) - 1, 2)=t1.tm04(+) " & _
'                                      "and substr(axg03, 1, length(axg03) - 9)=t2.tm01(+) and substr(axg03, length(axg03) - 8, 6)=t2.tm02(+) and substr(axg03, length(axg03) - 2, 1)=t2.tm03(+) and substr(axg03, length(axg03) - 1, 2)=t2.tm04(+) "
'                            intI = 1
'                            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                            If intI = 1 Then
'                                If IsNull(RsTemp.Fields("caseno").Value) = False And IsNull(RsTemp.Fields("pa26").Value) = False Then
'                                   'modify by sonia 2017/4/10 婉莘要求加X69534彩豐精技   2017/5/15陳德發及郭雅娟要求取消
'                                   'modify by sonia 2017/4/27 高國碩要求X60149改為只要母號,關係企業不要
'                                   If Left(RsTemp.Fields("pa26").Value, 6) = "X44551" Or Left(RsTemp.Fields("pa26").Value, 6) = "X62079" _
'                                   Or Left(RsTemp.Fields("pa26").Value, 6) = "X43988" Or Left(RsTemp.Fields("pa26").Value, 6) = "X63219" Or Left(RsTemp.Fields("pa26").Value, 6) = "X62319" _
'                                   Or Left(RsTemp.Fields("pa26").Value, 6) = "X60498" Or Left(RsTemp.Fields("pa26").Value, 6) = "X62702" Or Left(RsTemp.Fields("pa26").Value, 6) = "X63838" _
'                                    Then
'                                      'Modified by Morgan 2019/3/15 改印標記(表頭加說明)--婉莘
'                                      'strTmp = Left("" & .Fields("a1907").Value, 20) & "   水單不合併!"
'                                      strTmp = Left("" & .Fields("a1907").Value, 6) & "   @"
'                                   End If
'                                   'add by sonia 2017/4/14 單獨編號而關係企業不要的,例:王副總的華碩客戶(X69011010)
'                                   'modify by sonia 2017/4/27 高國碩要求X60149改為只要母號,關係企業不要
'                                   'modify by sonia 2017/10/11 張詠翔要求取消X6014900
'                                   'modify by sonia 2017/10/26 郭雅娟要求加X60738010國立清華大學
'                                   If Left(RsTemp.Fields("pa26").Value, 8) = "X6901101" Or Left(RsTemp.Fields("pa26").Value, 8) = "X6073801" Then
'                                      'Modified by Morgan 2019/3/15 改印標記(表頭加說明)--婉莘
'                                      'strTmp = Left("" & .Fields("a1907").Value, 20) & "   水單不合併!"
'                                      strTmp = Left("" & .Fields("a1907").Value, 6) & "   @"
'                                   End If
'                                   'end 2017/4/14
'                                End If
'
'                                'Added by Morgan 2019/3/15
'                                'Modified by Morgan 2019/7/15 帳單可能不只一張
'                                RsTemp.Find "a1501 = '" & .Fields("a1902") & "'", 0, adSearchForward, 1
'                                If Not RsTemp.EOF Then
'                                    If RsTemp.Fields("a1526") = "Y" Then
'                                        If strTmp <> "" Then
'                                           strTmp = strTmp & "獨立水單!"
'                                        Else
'                                           strTmp = Left("" & .Fields("a1907").Value, 6) & "   獨立水單!"
'                                        End If
'                                    End If
'                                End If
'                                'end 2019/3/15
'                            End If
                           If .Fields("A1720") = "B" Then
                              strTmp = Left("" & .Fields("a1907").Value, 6) & "   @獨立水單!"
                           ElseIf .Fields("A1720") = "X" Then
                              strTmp = Left("" & .Fields("a1907").Value, 6) & "   @"
                           ElseIf .Fields("A1720") = "U" Then
                              strTmp = Left("" & .Fields("a1907").Value, 6) & "   獨立水單!"
                           End If
'end 2019/7/16
                            strTmp = PUB_StrToStr(IIf(strTmp <> "", strTmp, "" & .Fields("a1907").Value), 40)
                            Printer.Print strTmp
                   End Select
                End If 'If PTitle(inP) <> "" And PTitle(inP) <> "結束" Then
             Next 'For inP = 1 To cInX
             strSubTotal = Val(strSubTotal) + Val(.Fields("a1904").Value)
             strGrp = "" & .Fields("a1901").Value '付款單號
             strGrp1 = "" & .Fields("a1917").Value '收據公司別
             PrintNewLine '換行
             .MoveNext
             '最後一行的小計
             If .EOF = True Then
                Printer.Line (PLeft(4), iPrint)-(PLeft(5) - ciColGap, iPrint)
                iPrint = iPrint + 150
                Printer.CurrentX = PLeft(3)
                Printer.CurrentY = iPrint
                Printer.Print "小計:"
                Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(strSubTotal, "##,##0.00")) - ciColGap
                Printer.CurrentY = iPrint
                Printer.Print Format(strSubTotal, "##,##0.00")
                PrintNewLine
             End If
          Loop
       End With

       PrintNewLine
       strTmp = "*** 結束 ***"
       Printer.Font.Bold = True
       Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strTmp)) / 2
       Printer.CurrentY = iPrint
       Printer.Print strTmp
       
       Printer.EndDoc
       ShowPrintOk
       Set RsTemp = Nothing
    Else
        MsgBox MsgText(28), , MsgText(5)
    End If
    
    Set rsPrt = Nothing

End Sub

'Added by Lydia 2016/08/15 列印-代理人未收未付對照表
Private Sub PrintRpt2170()
Dim inP As Integer
Dim inX As Integer
Dim rsPrt As New ADODB.Recordset
Dim strTmp As String

    mRptTitle = ReportTitle(217)
    '改A4橫印
    strTitle = "代理人編號,代理人名稱,第一筆未收請款日期,未收筆數,未收美金總金額,第一筆未付帳單日期,未收筆數,未付總金額,幣別,結束"
    strTitle2 = "0,5,11,9,4,7,9,4,6,3"
    ciFontSize = 12
    
    '抓未結匯的代理人
    strSql = "select distinct fa01,fa02, nvl(FA05||FA63||FA64||FA65, nvl(fa04, fa06)) as Name from acc150, fagent, acc170, acc190 where substr(a1503, 1, 8)=fa01(+) and substr(a1503, 9, 1)=fa02(+) and a1501=a1702 and a1709=a1901(+) and a1702=a1902(+) and a1908 is null order by fa01 asc, fa02 asc "
    inP = 1
    inX = 0
    Set rsPrt = ClsLawReadRstMsg(inP, strSql)
    If inP = 1 Then
       SettingPrtSet '設定印表機
       With rsPrt
          .MoveFirst
          Do While Not .EOF
             'A01~A04 未收請款單R21704~R21706,B01~B04 未付帳單R21707~R21710
             strSql = "select fa01||fa02 FANO,A01,A02,A03,A04,B01,B02,B03,B04 " & _
                      "from fagent,(select a1k28 fa1,min(a1k02) A01,count(a1k01) A02,sum(a1k08-(nvl(a1k30,0))/decode(a1k10,0,1,a1k10)-nvl(a1k31,0)) A03,min(a1k18) A04 from acc1k0 where a1k28='" & .Fields("fa01") & .Fields("fa02") & "' and (a1k11 > a1k30 or a1k30 is null or a1k30=0) and (a1k29 is null or a1k29='') and (a1k12 is null or a1k12=0) and a1k25 is null group by a1k28) X1 " & _
                      ",(select a1503 fa2,min(a1502) B01,count(a1501) B02,sum((a1506-nvl(a1520,0))) B03,min(a1505) B04 from acc150,acc170 where a1501=a1702 and a1503='" & .Fields("fa01") & .Fields("fa02") & "' and (a1506 > a1520 or a1520 is null or a1520=0) group by a1503) X2 " & _
                      "where fa01='" & .Fields("fa01") & "' and fa02='" & .Fields("fa02") & "' and fa01||fa02=fa1(+) and fa01||fa02=fa2(+) "
             intI = 1
             Set RsTemp = ClsLawReadRstMsg(intI, strSql)
             If intI = 1 Then
                '列印有未收和未付的資料
                If Val("" & RsTemp.Fields("A02")) > 0 And Val("" & RsTemp.Fields("B02")) > 0 Then
                    If inX = 0 Then
                        iPage = iPage + 1
                        PrintHeader
                        Printer.Font.Size = ciFontSize
                        Printer.FontBold = False
                    End If
                    '列印內容
                    For inP = 0 To cInX
                       If PTitle(inP) = "" Then Exit For
                       
                       If PTitle(inP) <> "" And PTitle(inP) <> "結束" Then
                          Printer.CurrentX = PLeft(inP)
                          Printer.CurrentY = iPrint
                          
                          Select Case inP
                              Case 0 '代理人編號R21702
                                   Printer.Print "" & RsTemp.Fields("FANO")
                                   
                              Case 1 '代理人名稱R21703
                                   Printer.Print PUB_StrToStr("" & .Fields("Name"), 20)
                                   
                              Case 2, 5 '第一筆未收請款日期R21704/未付帳單日期R21707
                                   strTmp = IIf(inP = 2, "" & RsTemp.Fields("A01"), "" & RsTemp.Fields("B01"))
                                   Printer.CurrentX = PLeft(inP) + (PLeft(inP + 1) - PLeft(inP) - Printer.TextWidth(strTmp) - ciColGap) / 2   '置中
                                   Printer.CurrentY = iPrint
                                   Printer.Print ChangeTStringToTDateString(strTmp)
                                   
                              Case 3, 6 '未收筆數R21705/未付筆數R21708
                                   strTmp = IIf(inP = 3, "" & RsTemp.Fields("A02"), "" & RsTemp.Fields("B02"))
                                   Printer.CurrentX = PLeft(inP + 1) - Printer.TextWidth(strTmp) - ciColGap '靠右
                                   Printer.CurrentY = iPrint
                                   Printer.Print strTmp
                                   
                              Case 4, 7 '未收美金總金額R21706/未付總金額R21709
                                   strTmp = Format(IIf(inP = 4, "" & RsTemp.Fields("A03"), "" & RsTemp.Fields("B03")), "##,##0.00")
                                   Printer.CurrentX = PLeft(inP + 1) - Printer.TextWidth(strTmp) - ciColGap '靠右
                                   Printer.CurrentY = iPrint
                                   Printer.Print strTmp
                                   
                              Case 8 '幣別 R21710
                                   Printer.Print "" & RsTemp.Fields("B04")
                          
                          End Select
                       End If 'If PTitle(inP) <> "" And PTitle(inP) <> "結束" Then
                    Next 'For inP = 1 To cInX
                    inX = inX + 1
                    PrintNewLine
                End If
             End If
             .MoveNext
          Loop

       PrintNewLine
       strTmp = "*** 結束 ***"
       Printer.Font.Bold = True
       Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strTmp)) / 2
       Printer.CurrentY = iPrint
       Printer.Print strTmp
       
       End With
       
       Printer.EndDoc
       If inX > 0 Then
          ShowPrintOk
          Set RsTemp = Nothing
       End If
    Else
        MsgBox MsgText(28), , MsgText(5)
    End If
    
    Set rsPrt = Nothing
    
End Sub

'Added by Lydia 2016/08/15 列印-國內未收款明細表
Private Sub PrintRpt2160()
Dim inP As Integer
Dim rsPrt As New ADODB.Recordset
Dim strTmp As String
Dim strAcNo As String, strAcDate As String '收據編號,收據日期
Dim strAmt As String '應收金額
Dim strAcAmt As String '未收金額

    mRptTitle = ReportTitle(216)
    '改A4橫印
    strTitle = "代理人編號,代理人名稱,帳單編號,幣別,金額,本所案號,客戶編號,收據日期,智權人員,收據編號,應收金額,未收金額,結束"
    strTitle2 = "0,5,11,5,2,5,6,5,4,4,5,5,5"
    ciFontSize = 11
    
    '抓未結匯資料
    strSql = "select new.*,nvl(FA05,nvl(fa04, fa06)) as FName from (select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505 from caseprogress, acc0k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='E' AND cp60 = a0k01 (+) and cp61 = a1501 and cp61 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0)) or (a0k01 is null)) and a1908 is null union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505 from caseprogress, acc0k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='E' AND cp60 = a0k01 (+) and cp62 = a1501 and cp62 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0)) or (a0k01 is null)) and a1908 is null union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505 from caseprogress, acc0k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='E' AND cp60 = a0k01 (+) and cp63 = a1501 and cp63 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0)) or (a0k01 is null)) and a1908 is null union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505 from caseprogress, acc0k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='E' AND cp60 = a0k01 (+) and cp87 = a1501 and cp87 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0)) or (a0k01 is null)) and a1908 is null union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505 from caseprogress, acc0k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='E' AND cp60 = a0k01 (+) and cp88 = a1501 and cp88 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0)) or (a0k01 is null)) and a1908 is null union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505 from caseprogress, acc0k0, acc150, acc170, acc190 where CP60 IS NULL AND cp60 = a0k01 (+) and cp61 = a1501 and cp61 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0))) and a1908 is null union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505 from caseprogress, acc0k0, acc150, acc170, acc190 where CP60 IS NULL AND cp60 = a0k01 (+) and cp62 = a1501 and cp62 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0))) and a1908 is null union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505 from caseprogress, acc0k0, acc150, acc170, acc190 where CP60 IS NULL AND cp60 = a0k01 (+) and cp63 = a1501 and cp63 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0))) and a1908 is null union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505 from caseprogress, acc0k0, acc150, acc170, acc190 where CP60 IS NULL AND cp60 = a0k01 (+) and cp87 = a1501 and cp87 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0))) and a1908 is null union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505 from caseprogress, acc0k0, acc150, acc170, acc190 where CP60 IS NULL AND cp60 = a0k01 (+) and cp88 = a1501 and cp88 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0))) and a1908 is null union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A1K01,A1K02,A1501,A1503,A1505 from caseprogress, acc1k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='X' AND cp60 = a1k01 (+) and cp61 = a1501 and cp61 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and a1k29 is null and a1908 is null union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A1K01,A1K02,A1501,A1503,A1505 from caseprogress, acc1k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='X' AND cp60 = a1k01 (+) and cp62 = a1501 and cp62 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and a1k29 is null and a1908 is null union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A1K01,A1K02,A1501,A1503,A1505 from caseprogress, acc1k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='X' AND cp60 = a1k01 (+) and cp63 = a1501 and cp63 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and a1k29 is null and a1908 is null union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A1K01,A1K02,A1501,A1503,A1505 from caseprogress, acc1k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='X' AND cp60 = a1k01 (+) and cp87 = a1501 and cp87 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and a1k29 is null and a1908 is null union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A1K01,A1K02,A1501,A1503,A1505 from caseprogress, acc1k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='X' AND cp60 = a1k01 (+) and cp88 = a1501 and cp88 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and a1k29 is null and a1908 is null) new " & _
                             ",FAGENT WHERE SUBSTR(A1503,1,8)=FA01(+) AND SUBSTR(A1503,9,1)=FA02(+) " & _
                             "order by a1501,cp44 asc"
    inP = 1
    Set rsPrt = ClsLawReadRstMsg(inP, strSql)
    If inP = 1 Then
       SettingPrtSet '設定印表機
       With rsPrt
          .MoveFirst
          iPage = iPage + 1
          PrintHeader
          Printer.Font.Size = ciFontSize
          Printer.FontBold = False
          Do While Not .EOF
          
            strAmt = "" & .Fields("CP16") '應收金額
            If "" & .Fields("CP77") <> "" Then strAmt = Val(strAmt) - Val(.Fields("CP77"))
            '抓收據編號、日期和未收金額
            If "" & .Fields("A0K01") <> "" Then
                strAcNo = .Fields("A0K01")
                strAcDate = "" & .Fields("A0K02")
                strAcAmt = strAmt
                '減已收金額
                If "" & .Fields("CP75") <> "" Then strAcAmt = Val(strAcAmt) - Val("" & .Fields("CP75"))
                '已退費金額
                If "" & .Fields("CP78") <> "" Then strAcAmt = Val(strAcAmt) + Val("" & .Fields("CP78"))
            Else
                strSql = "select A1K01,A1K02,A1K29,A1K30 from acc1k0 where a1k01 = '" & "" & .Fields("CP60") & "'"
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                If intI = 1 Then
                    '尚未結清
                    If IsNull(RsTemp.Fields("A1K29")) Then
                        strAcNo = "" & RsTemp.Fields("A1K01")
                        strAcDate = "" & RsTemp.Fields("A1K02")
                        '減已收金額
                        strAcAmt = Val(strAmt) - Val("" & RsTemp.Fields("A1K30"))
                    Else '已結清,不列印
                        GoTo JumpPrint
                    End If
                Else
                    strAcNo = ""
                    strAcDate = ""
                    strAcAmt = "0"
                End If
            End If

            '列印內容
            For inP = 0 To cInX
               If PTitle(inP) = "" Then Exit For
               
               If PTitle(inP) <> "" And PTitle(inP) <> "結束" Then
                  Printer.CurrentX = PLeft(inP)
                  Printer.CurrentY = iPrint
                  
                  Select Case inP
                      Case 0 '代理人編號R21602
                           Printer.Print "" & .Fields("A1503")
                           
                      Case 1 '代理人名稱R21603
                           Printer.Print PUB_StrToStr("" & .Fields("FName"), 19)
                           
                      Case 2 '帳單編號R21604
                           Printer.Print "" & .Fields("A1501")
                           
                      Case 3 '幣別R21605
                           Printer.Print "" & .Fields("A1505")
                           
                      Case 4 '金額R21606
                           strTmp = "0"
                           strSql = "select axf04 from acc151 where axf01='" & .Fields("A1501") & "' and axf02='" & .Fields("cp09") & "' "
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                           If intI = 1 Then
                              strTmp = "" & RsTemp(0)
                           End If
                           strTmp = Format(Val(strTmp), "##,##0.00")
                           Printer.CurrentX = PLeft(inP + 1) - Printer.TextWidth(strTmp) - ciColGap '靠右
                           Printer.CurrentY = iPrint
                           Printer.Print strTmp
                           
                      Case 5 '本所案號R21607
                           Printer.Print "" & .Fields("CP01") & .Fields("CP02") & .Fields("CP03") & .Fields("CP04")
                           
                      Case 6 '客戶編號R21608
                           strTmp = PUB_GetCustNo("" & .Fields("CP01") & .Fields("CP02") & .Fields("CP03") & .Fields("CP04"))
                           Printer.Print strTmp
                           
                      Case 7 '收據日期R21609
                           Printer.Print ChangeTStringToTDateString(strAcDate)
                           
                      Case 8 '智權人員R21610
                           Printer.Print "" & .Fields("CP13")
                           
                      Case 9 '收據編號R21611
                           Printer.Print strAcNo
                           
                      Case 10, 11 '應收金額R21612 / 未收金額R21613
                           strTmp = Format(Val(IIf(inP = 10, strAmt, strAcAmt)), DDollar2)
                           Printer.CurrentX = PLeft(inP + 1) - Printer.TextWidth(strTmp) - ciColGap '靠右
                           Printer.CurrentY = iPrint
                           Printer.Print strTmp
                           
                  End Select
               End If 'If PTitle(inP) <> "" And PTitle(inP) <> "結束" Then
            Next 'For inP = 1 To cInX

            PrintNewLine
JumpPrint:
             .MoveNext
          Loop

       PrintNewLine
       strTmp = "*** 結束 ***"
       Printer.Font.Bold = True
       Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strTmp)) / 2
       Printer.CurrentY = iPrint
       Printer.Print strTmp
       
       End With
       
       Printer.EndDoc
       ShowPrintOk
       Set RsTemp = Nothing
    Else
        MsgBox MsgText(28), , MsgText(5)
    End If
    Set rsPrt = Nothing
End Sub

'Added by Lydia 2022/03/30 案件失誤明細表
Private Sub PrintReport003()
Dim inP As Integer
Dim rsPrt As New ADODB.Recordset
Dim strTmp As String, strFileN As String
Dim strAmt As String '應收金額
Dim strAcAmt As String '未收金額
'----------
Dim xlsReport As New Excel.Application
Dim wksrpt As New Worksheet
Dim intRow As Integer, intField As Integer
Dim intCount As Integer
Dim strField, intWidth
Dim strArr(0 To 8) As Variant
    
    '抓目前未結匯資料並且符合以下兩個條件:
    '1.U帳單日期相較收文日達2年,顯示收文日;排除年費程序(專利之605)
    '2.帳單需要主管審核者,顯示帳單備註; (from 婉莘: 帳單需要主管審核者=>主管已審核完A1521=Y,同時也會寫備註)
    strTmp = " AND ((CP05+20000<=A1502+19110000 AND CP01||CP10 NOT IN ('P605','CFP605','FCP605')) OR A1521='Y') "
    strSql = "select new.*,nvl(FA05,nvl(fa04, fa06)) as FName,DECODE(PA09||TM10||SP09||LC15,'000',CPM03,CPM04) CPM0304 from (" & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505,CP05,CP10,A1521,A1509 from caseprogress, acc0k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='E' AND cp60 = a0k01 (+) and cp61 = a1501 and cp61 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0)) or (a0k01 is null)) and a1908 is null " & strTmp & " union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505,CP05,CP10,A1521,A1509 from caseprogress, acc0k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='E' AND cp60 = a0k01 (+) and cp62 = a1501 and cp62 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0)) or (a0k01 is null)) and a1908 is null " & strTmp & " union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505,CP05,CP10,A1521,A1509 from caseprogress, acc0k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='E' AND cp60 = a0k01 (+) and cp63 = a1501 and cp63 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0)) or (a0k01 is null)) and a1908 is null " & strTmp & " union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505,CP05,CP10,A1521,A1509 from caseprogress, acc0k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='E' AND cp60 = a0k01 (+) and cp87 = a1501 and cp87 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0)) or (a0k01 is null)) and a1908 is null " & strTmp & " union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505,CP05,CP10,A1521,A1509 from caseprogress, acc0k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='E' AND cp60 = a0k01 (+) and cp88 = a1501 and cp88 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0)) or (a0k01 is null)) and a1908 is null " & strTmp & " union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505,CP05,CP10,A1521,A1509 from caseprogress, acc0k0, acc150, acc170, acc190 where CP60 IS NULL AND cp60 = a0k01 (+) and cp61 = a1501 and cp61 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0))) and a1908 is null " & strTmp & " union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505,CP05,CP10,A1521,A1509 from caseprogress, acc0k0, acc150, acc170, acc190 where CP60 IS NULL AND cp60 = a0k01 (+) and cp62 = a1501 and cp62 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0))) and a1908 is null " & strTmp & " union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505,CP05,CP10,A1521,A1509 from caseprogress, acc0k0, acc150, acc170, acc190 where CP60 IS NULL AND cp60 = a0k01 (+) and cp63 = a1501 and cp63 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0))) and a1908 is null " & strTmp & " union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505,CP05,CP10,A1521,A1509 from caseprogress, acc0k0, acc150, acc170, acc190 where CP60 IS NULL AND cp60 = a0k01 (+) and cp87 = a1501 and cp87 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0))) and a1908 is null " & strTmp & " union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A0K01,A0K02,A1501,A1503,A1505,CP05,CP10,A1521,A1509 from caseprogress, acc0k0, acc150, acc170, acc190 where CP60 IS NULL AND cp60 = a0k01 (+) and cp88 = a1501 and cp88 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and (NVL(CP16,0) > (nvl(CP75, 0) + NVL(CP77,0))) and a1908 is null " & strTmp & " union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A1K01,A1K02,A1501,A1503,A1505,CP05,CP10,A1521,A1509 from caseprogress, acc1k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='X' AND cp60 = a1k01 (+) and cp61 = a1501 and cp61 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and a1k29 is null and a1908 is null " & strTmp & " union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A1K01,A1K02,A1501,A1503,A1505,CP05,CP10,A1521,A1509 from caseprogress, acc1k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='X' AND cp60 = a1k01 (+) and cp62 = a1501 and cp62 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and a1k29 is null and a1908 is null " & strTmp & " union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A1K01,A1K02,A1501,A1503,A1505,CP05,CP10,A1521,A1509 from caseprogress, acc1k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='X' AND cp60 = a1k01 (+) and cp63 = a1501 and cp63 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and a1k29 is null and a1908 is null " & strTmp & " union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A1K01,A1K02,A1501,A1503,A1505,CP05,CP10,A1521,A1509 from caseprogress, acc1k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='X' AND cp60 = a1k01 (+) and cp87 = a1501 and cp87 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and a1k29 is null and a1908 is null " & strTmp & " union " & _
                             "select CP01,CP02,CP03,CP04,CP09,CP13,CP16,CP44,CP60,CP75,CP77,CP78,A1K01,A1K02,A1501,A1503,A1505,CP05,CP10,A1521,A1509 from caseprogress, acc1k0, acc150, acc170, acc190 where SUBSTR(CP60,1,1)='X' AND cp60 = a1k01 (+) and cp88 = a1501 and cp88 = a1702 and a1709 = a1901 (+) and a1702 = a1902 (+) and a1k29 is null and a1908 is null) new " & _
                             ",FAGENT,PATENT,TRADEMARK,SERVICEPRACTICE,LAWCASE,CASEPROPERTYMAP WHERE SUBSTR(A1503,1,8)=FA01(+) AND SUBSTR(A1503,9,1)=FA02(+) " & _
                             "AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) " & _
                             "AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) " & _
                             "AND CP01=CPM01(+) AND CP10=CPM02(+) order by a1503,a1503 asc"
    inP = 1
    Set rsPrt = ClsLawReadRstMsg(inP, strSql)
    If inP = 0 Then
        MsgBox MsgText(28), , cntReport003
    Else
        rsPrt.MoveFirst
        Do While Not rsPrt.EOF
            If intRow = 0 Then
                strFileN = cntReport003 & MsgText(43)
                If Dir(strExcelPath & strFileN) = MsgText(601) Then
                    Call Pub_ChkExcelPath(strExcelPath)
                Else
                    Kill strExcelPath & strFileN
                End If
                xlsReport.SheetsInNewWorkbook = 1
                xlsReport.Workbooks.add
                Set wksrpt = xlsReport.Worksheets(1)
                wksrpt.Activate
                strTitle = "代理人編號,代理人名稱,帳單編號,本所案號,收文日,案件性質,幣別,應收金額,帳　單　備　註"
                strTitle2 = "10,25,10,12,9,15,6,10,35"
                strField = Split(strTitle, ",")
                intWidth = Split(strTitle2, ",")
                intField = 65 '起始位置intField=65=>A
                intRow = 1
                '設定抬頭
                For inP = 0 To UBound(strField)
                    wksrpt.Range(Chr(intField + inP) & ":" & Chr(intField + inP)).Font.Size = 11
                    wksrpt.Range(Chr(intField + inP) & ":" & Chr(intField + inP)).ColumnWidth = Val(intWidth(inP))
                    If inP = 7 Then '金額
                        wksrpt.Range(Chr(intField + inP) & ":" & Chr(intField + inP)).NumberFormatLocal = "#,###,##0"
                    Else
                        wksrpt.Range(Chr(intField + inP) & ":" & Chr(intField + inP)).NumberFormatLocal = "@"
                    End If
                    If inP = 2 Or inP = 4 Or inP = 6 Then '置中：帳單編號,收文日,幣別
                       wksrpt.Range(Chr(intField + inP) & ":" & Chr(intField + inP)).HorizontalAlignment = xlCenter
                    ElseIf inP = 7 Then '靠右: 應收金額
                       wksrpt.Range(Chr(intField + inP) & ":" & Chr(intField + inP)).HorizontalAlignment = xlRight
                    Else
                       wksrpt.Range(Chr(intField + inP) & ":" & Chr(intField + inP)).HorizontalAlignment = xlLeft
                    End If
                    wksrpt.Range(Chr(intField + inP) & "4").Value = Trim(strField(inP))
                Next inP
                  wksrpt.Range(Chr(intField) & intRow).Value = "*** " & cntReport003 & " ***"
                  wksrpt.Range(Chr(intField) & intRow).Font.Size = 18
                  wksrpt.Range(Chr(intField) & intRow).Font.Bold = True
                  wksrpt.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).MergeCells = True
                  wksrpt.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).HorizontalAlignment = xlCenter
                  wksrpt.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).VerticalAlignment = xlCenter
                  intRow = intRow + 2
                  wksrpt.Range(Chr(intField) & intRow).Value = "列印人員："
                  wksrpt.Range(Chr(intField + 1) & intRow).Value = strUserName
                  wksrpt.Range(Chr(intField + UBound(strField)) & intRow).Value = "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
                  wksrpt.Range(Chr(intField + UBound(strField)) & intRow).HorizontalAlignment = xlRight
                  intRow = intRow + 1
                  '底部格線
                  wksrpt.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).Borders(xlEdgeBottom).LineStyle = xlContinuous
                  wksrpt.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).Borders(xlEdgeBottom).Weight = xlThin
                  intRow = intRow + 1
            End If
            '代理人編號
            strArr(0) = "" & rsPrt.Fields("A1503")
            '代理人名稱
            strArr(1) = "" & rsPrt.Fields("FNAME")
            '帳單編號
            strArr(2) = "" & rsPrt.Fields("A1501")
            '本所案號
            strArr(3) = "" & rsPrt.Fields("CP01") & rsPrt.Fields("CP02") & rsPrt.Fields("CP03") & rsPrt.Fields("CP04")
            '收文日
            strArr(4) = IIf("" & rsPrt.Fields("A1521") = "", ChangeWStringToTDateString("" & rsPrt.Fields("CP05")), "")
            '案件性質
            strArr(5) = "" & rsPrt.Fields("CPM0304")
            '幣別
            strArr(6) = "" & rsPrt.Fields("A1505")
            '應收金額
              strAmt = Val("" & rsPrt.Fields("CP16")) '應收金額
              If "" & rsPrt.Fields("CP77") <> "" Then strAmt = Val(strAmt) - Val(rsPrt.Fields("CP77"))
              '抓收據編號、日期和未收金額
              If "" & rsPrt.Fields("A0K01") <> "" Then
                  strAcAmt = strAmt
                  '減已收金額
                  If "" & rsPrt.Fields("CP75") <> "" Then strAcAmt = Val(strAcAmt) - Val("" & rsPrt.Fields("CP75"))
                  '已退費金額
                  If "" & rsPrt.Fields("CP78") <> "" Then strAcAmt = Val(strAcAmt) + Val("" & rsPrt.Fields("CP78"))
              Else
                  strSql = "select A1K01,A1K02,A1K29,A1K30 from acc1k0 where a1k01 = '" & "" & rsPrt.Fields("CP60") & "'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                      '尚未結清
                      If "" & RsTemp.Fields("A1K29") = "" Then
                          '減已收金額
                          strAcAmt = Val(strAmt) - Val("" & RsTemp.Fields("A1K30"))
                      Else '已結清,不列印
                          GoTo JumpPrint
                      End If
                  Else
                      strAcAmt = "0"
                  End If
              End If
            strArr(7) = strAcAmt
            
            '帳　單　備　註
            strArr(8) = IIf(strArr(4) = "", PUB_StrToStr("" & rsPrt.Fields("A1509"), 36), "")
            wksrpt.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).Value = strArr
            intCount = intCount + 1
            intRow = intRow + 1
JumpPrint:
            rsPrt.MoveNext
        Loop
          
        '列印
        If intCount > 0 Then
            intRow = intRow + 1
            wksrpt.Range(Chr(intField) & intRow).Value = "*** 結　　束 ***"
            wksrpt.Range(Chr(intField) & intRow).Font.Size = 18
            wksrpt.Range(Chr(intField) & intRow).Font.Bold = True
            wksrpt.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).MergeCells = True
            wksrpt.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).HorizontalAlignment = xlCenter
            wksrpt.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).VerticalAlignment = xlCenter
            '列印設定
            wksrpt.PageSetup.PaperSize = 9 'A4
            wksrpt.PageSetup.PrintTitleRows = "$1:$4"
            wksrpt.PageSetup.Orientation = xlLandscape '橫印
            wksrpt.PageSetup.LeftMargin = xlsReport.InchesToPoints(0.4) '邊界
            wksrpt.PageSetup.RightMargin = xlsReport.InchesToPoints(0.4)
            wksrpt.PageSetup.TopMargin = xlsReport.InchesToPoints(0.4)
            wksrpt.PageSetup.BottomMargin = xlsReport.InchesToPoints(0.4)
            wksrpt.PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
            wksrpt.PrintOut Copies:=1, Collate:=True
        End If
        '判斷若版本2007以上改變存格式
        If Val(xlsReport.Version) < 12 Then
            xlsReport.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=-4143
        Else
            xlsReport.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=56
        End If
        xlsReport.Workbooks.Close
        xlsReport.Quit
        
        Set wksrpt = Nothing
        Set xlsReport = Nothing
        Set RsTemp = Nothing
        ShowPrintOk
        
        If Dir(strExcelPath & strFileN) <> MsgText(601) Then
           Kill strExcelPath & strFileN
        End If
    End If
    Set rsPrt = Nothing
End Sub

