VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc1151 
   AutoRedraw      =   -1  'True
   Caption         =   "收款資料輸入"
   ClientHeight    =   5112
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5112
   ScaleWidth      =   8760
   Begin VB.TextBox Text5 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4332
      MaxLength       =   15
      TabIndex        =   1
      Top             =   828
      Width           =   1572
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc1151.frx":0000
      Height          =   3228
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   8292
      _ExtentX        =   14626
      _ExtentY        =   5694
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   14
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "t0202"
         Caption         =   "收據編號"
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
         DataField       =   "t0217"
         Caption         =   "銷帳否"
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
         DataField       =   "t0203"
         Caption         =   "發票號碼"
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
         DataField       =   "t0204"
         Caption         =   "本次服務費"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "t0205"
         Caption         =   "本次規費"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "t0206"
         Caption         =   "扣繳額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "t0207"
         Caption         =   "扣繳年度"
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
         DataField       =   "t0212"
         Caption         =   "智權人員"
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
         DataField       =   "t0208"
         Caption         =   "應收服務費"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "t0209"
         Caption         =   "應收規費"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "t0210"
         Caption         =   "已收服務費"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "t0211"
         Caption         =   "已收規費"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "t0219"
         Caption         =   "已銷服務費"
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
      BeginProperty Column13 
         DataField       =   "t0220"
         Caption         =   "已銷規費"
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
      BeginProperty Column14 
         DataField       =   "t0221"
         Caption         =   "已退服務費"
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
      BeginProperty Column15 
         DataField       =   "t0222"
         Caption         =   "已退規費"
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
            Locked          =   -1  'True
            ColumnWidth     =   1344.189
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   684.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1175.811
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1116.284
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   852.095
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   875.906
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1091.906
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1116.284
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column12 
            Alignment       =   1
         EndProperty
         BeginProperty Column13 
            Alignment       =   1
         EndProperty
         BeginProperty Column14 
         EndProperty
         BeginProperty Column15 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   8064
      Picture         =   "Frmacc1151.frx":0015
      Style           =   1  '圖片外觀
      TabIndex        =   4
      ToolTipText     =   "取消"
      Top             =   744
      Width           =   350
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6840
      TabIndex        =   13
      Top             =   240
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1308
      TabIndex        =   10
      Top             =   4572
      Width           =   1572
   End
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
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   0
      Top             =   780
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      Caption         =   "轉暫收款"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3840
      TabIndex        =   5
      Top             =   240
      Width           =   1332
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      TabIndex        =   7
      Top             =   240
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   6816
      TabIndex        =   3
      Top             =   4584
      Width           =   1572
      _ExtentX        =   2773
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   228
      Top             =   1080
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   2117
      _ExtentY        =   572
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
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "扣繳金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3372
      TabIndex        =   14
      Top             =   828
      Width           =   972
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "轉暫收款單號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5388
      TabIndex        =   12
      Top             =   252
      Width           =   1452
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   -24
      Top             =   4944
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "欲處理日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5616
      TabIndex        =   11
      Top             =   4584
      Width           =   1212
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "溢收金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   336
      TabIndex        =   9
      Top             =   4584
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   540
      Left            =   240
      Top             =   108
      Width           =   8292
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收據編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   8
      Top             =   816
      Width           =   972
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "收款單號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc1151"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit

Public adoacc0k0 As New ADODB.Recordset
Public adoacc0l0 As New ADODB.Recordset
Public adoacc0m0 As New ADODB.Recordset
Public adoacctmp02 As New ADODB.Recordset
Public adosubsum As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoacc0e0 As New ADODB.Recordset
Public adocheck As New ADODB.Recordset
Public adoacc1u0 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Dim lng0 As Long
Dim lng1 As Long
Dim lng2 As Long
Dim lng3 As Long
Dim lng4 As Long
Dim douAmount As Double
Dim douFAmount As Double
Dim douRAmount1 As Double
Dim douRAmount2 As Double
Dim douRAmount3 As Double
Dim douTAmount(8) As Double
Dim strCustomer As String
Dim strSalesMan As String
Dim strMan As String
Dim strManNo As String
Dim strCompany As String
Dim strCompanyNo As String
Dim strProperty As String
Dim strCaseNo As String
Dim strDept As String
Dim stra1p22 As String
Dim stra1p27 As String
Dim strNation As String
Dim strNationNo As String
Dim dou0 As Double
Dim dou1 As Double
Dim dou2 As Double
Dim dou3 As Double
Dim dou4 As Double
Dim bolShow As Boolean
Dim strSupportCaseList As String '要扣支援點數的案件清單 Add by Morgan 2010/6/2
Dim strXFeeCaseList As String '要扣出庭費的案件清單 Add by Morgan 2011/5/25
Dim strAssignCaseList As String '要分配點數的案件清單 Add by Morgan 2011/5/25
Dim strProFeeCaseList As String '要扣智慧所專業部點數的案件清單 Add by Morgan 2021/1/20
'Add by Morgan 2006/8/21 新增專業點數傳票分錄語法
Dim strSQLD As String, strSQLc As String
'Added by Morgan 2013/12/19
Dim strA1P01 As String '公司別
'Dim strA1P22_1 As String 'J公司傳票號 'Removed by Morgan 2020/4/15
'Dim strA1P22_J As String '1公司傳票號 'Removed by Morgan 2020/4/15

Dim strA0L05 As String '主要公司別
Dim strA0L05A1P22 As String '主要公司別傳票號
Dim strSerialNo As String '分錄序次
Dim bolDetailChangeMail As Boolean '繳款明細與收款不同通知
Dim F5639NO As String, strF5639NO As String 'add by sonia 2016/10/18 記錄寰華介紹案件之收據編號
Dim strLOS02 As String, bolB2NeeCourt As Boolean, strTTMan As String, strTTManSN As String, strLCaseNo As String, strLRcpTitle As String, lngTTAmt As Long 'Added by Morgan 2021/1/18
Dim stNoLawyerAlert As String 'Added by Morgan 2021/2/1
Dim strA1P22_TT As String 'Added by Morgan 2021/4/12
Dim m_bolRcptClear As Boolean 'Added by Morgan 2023/9/26 收據是否結清
Dim strA1P22_L As String 'Added by Morgan 2023/11/20
Dim bolACSNoTaxItem As Boolean 'Added by Morgan 2025/8/19 是否ACS未收款已沖帳

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   strCustNo = MsgText(601)
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 8850
   Me.Height = 5500
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   Text2 = strItemNo '收款單號
   Text1 = MsgText(802) '收據編號
   MaskEdBox1.Mask = "" '欲處理日期--預設系統日
   MaskEdBox1.Text = CFDate(ACDate(ServerDate))
   MaskEdBox1.Mask = DFormat
   strA0L05 = Frmacc1150.Text21 'Added by Morgan 2014/1/2
   OpenTable
   ZeroTaxCustAlert Text2 'Added by Morgan 2024/11/11
   SumShow
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133) & "/" & MsgText(107)
End Sub

Private Sub Form_Activate()
   strFormName = Name
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call PUB_SaveTrackMode(1, KeyCode)  'Form2.0 記錄鍵盤傳入順序
   KeyDefine KeyCode
End Sub

Private Sub Text1_GotFocus()
   'Modify by Morgan 2005/11/2
   'TextInverse Text1
   If Len(Text1) > 0 Then
      Text1.SelStart = 1
      Text1.SelLength = Len(Text1) - 1
   End If
   
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133) & "/" & MsgText(107)
End Sub

Private Sub Command1_Click()
   Dim bAuto As Boolean
   Dim strPrinter As String
   
   If Text4 <> MsgText(601) Then
      Exit Sub
   End If
   'Added by Morgan 2013/12/26
   strExc(0) = "select * from acc440 where a4416='" & Text2 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      bAuto = True
      If Val(Format(Text3)) <> Val("" & RsTemp("A4410")) Then
         If MsgBox("溢收金額與繳款記錄不符是否要繼續？", vbYesNo + vbQuestion, "收款檢查") = vbNo Then
            Exit Sub
         End If
      End If
   End If
   'end 2013/12/26
   Acc0t0Save
   FormEnabled
   
'Removed by Morgan 2014/1/9 取消,另有功能做整批列印
'   'Added by Morgan 2013/12/26
'   If bAuto = True Then
'      If MsgBox("是否要列印付款憑證？" & vbCrLf & vbCrLf & "備註：" & vbCrLf & Frmacc1150.Text1, vbYesNo + vbQuestion, "收款檢查") = vbYes Then
'         strPrinter = Printer.DeviceName
'         frm880011.bolAppOnly = True
'         frm880011.Show 1
'         PrintProof
'         PUB_RestorePrinter strPrinter
'      End If
'   End If
'   'end 2013/12/26
'end 2014/1/9

End Sub
'Added by Morgan 2013/12/26
Private Sub PrintProof()
   Const RowHeight As Integer = 500
   Dim iRow As Integer
   
   strExc(0) = "select distinct a4401||' '||st02 C1,sqldatet(a4402)||' '||sqltime(a4403) C2,x.*,a0k04" & _
      " from acc440 x,staff,acc441,acc0k0 where a4416='" & Text2 & "' and st01(+)=a4401" & _
      " and axd01(+)=a4401 and axd02(+)=a4402 and axd03(+)=a4403 and a0k01(+)=axd04 order by a0k04"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      
      Printer.FontSize = 18
      Printer.FontBold = True
      Printer.CurrentX = 1500
      Printer.CurrentY = 300
      Printer.Print "付款憑證"
      
      Printer.FontSize = 16
      Printer.FontBold = False
      
      iRow = 0
      strExc(1) = "智權人員：" & RsTemp.Fields("C1")
      iRow = iRow + 1
      Printer.CurrentX = 500
      Printer.CurrentY = 500 + iRow * RowHeight
      Printer.Print strExc(1)
      
      strExc(1) = "繳款日期時間：" & RsTemp.Fields("C2")
      iRow = iRow + 1
      Printer.CurrentX = 500
      Printer.CurrentY = 500 + iRow * RowHeight
      Printer.Print strExc(1)
      
      strExc(1) = "收據抬頭："
      Do While Not .EOF
         strExc(1) = strExc(1) & RsTemp("a0k04") & " "
         .MoveNext
      Loop
      .MoveFirst
      iRow = iRow + 1
      Printer.CurrentX = 500
      Printer.CurrentY = 500 + iRow * RowHeight
      Printer.Print strExc(1)
      
      strExc(1) = "暫收款單號：" & Text4
      iRow = iRow + 1
      Printer.CurrentX = 500
      Printer.CurrentY = 500 + iRow * RowHeight
      Printer.Print strExc(1)
      
      strExc(1) = "收款情形："
      
      'Removed by Morgan 2014/11/6 繳款作業已取消票號欄位
      'If Not IsNull(.Fields("a4404")) Then
      '   strExc(1) = strExc(1) & " 票據號碼 " & .Fields("a4404")
      'End If
      
      If .Fields("a4405") > 0 Then
         strExc(1) = strExc(1) & " 票據金額 " & Format(.Fields("a4405"), "#,##0")
      End If
      If .Fields("a4406") > 0 Then
         strExc(1) = strExc(1) & " 北所電匯金額 " & Format(.Fields("a4406"), "#,##0")
      End If
      If .Fields("a4407") > 0 Then
         strExc(1) = strExc(1) & " 分所電匯金額 " & Format(.Fields("a4407"), "#,##0")
      End If
      If .Fields("a4408") > 0 Then
         strExc(1) = strExc(1) & " 現金 " & Format(.Fields("a4408"), "#,##0")
      End If
      If .Fields("a4409") > 0 Then
         strExc(1) = strExc(1) & " 抵暫收款 " & Format(.Fields("a4409"), "#,##0")
      End If
      If .Fields("a4410") > 0 Then
         strExc(1) = strExc(1) & " 溢收款 " & Format(.Fields("a4410"), "#,##0")
      End If
      If .Fields("a4411") > 0 Then
         strExc(1) = strExc(1) & " 手續費 " & Format(.Fields("a4411"), "#,##0")
      End If
      If .Fields("a4422") > 0 Then
         strExc(1) = strExc(1) & " 補扣繳/外幣 " & Format(.Fields("a4411"), "#,##0")
      End If
      iRow = iRow + 1
      Printer.CurrentX = 500
      Printer.CurrentY = 500 + iRow * RowHeight
      Printer.Print strExc(1)
      
      If Text4 <> "" Then
         strExc(1) = "溢收金額：" & Format(Text3, "#,##0")
         iRow = iRow + 1
         Printer.CurrentX = 500
         Printer.CurrentY = 500 + iRow * RowHeight
         Printer.Print strExc(1)
      End If
      
      If Frmacc1150.Text1 <> "" Then
         strExc(1) = "備註：" & Frmacc1150.Text1
         iRow = iRow + 1
         Printer.CurrentX = 500
         Printer.CurrentY = 500 + iRow * RowHeight
         Printer.Print strExc(1)
      End If
      
      End With
      
      strExc(0) = "select a1p12 from acc1p0 where a1p04='" & Text2 & "' and a1p05='113001'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strExc(1) = "票據到期日："
         With RsTemp
         Do While Not .EOF
            strExc(1) = strExc(1) & ChangeTStringToTDateString(.Fields("a1p12")) & " "
            .MoveNext
         Loop
         End With
         iRow = iRow + 1
         Printer.CurrentX = 500
         Printer.CurrentY = 500 + iRow * RowHeight
         Printer.Print strExc(1)
      End If
      Printer.EndDoc
   End If
End Sub

'*************************************************
'  儲存資料表(國內暫收款資料)
'
'*************************************************
Private Sub Acc0t0Save()
Dim strNo As String

   If Val(Text3) <= 0 Then
      Exit Sub
   End If
   If MaskEdBox1.Text = MsgText(29) Then
      MsgBox MsgText(85), , MsgText(5)
      Exit Sub
   End If
   strNo = AutoNo(MsgText(806), 5)
   If Adodc1.Recordset.RecordCount <> 0 Then
      'Modify by Morgan 2008/1/16
      '因為大陸王俊貴會代收多個客戶的款做一筆收款所以要抓最後一張收據的資料這樣ACC0T0才會和ACC1P0的資料同步
      'Adodc1.Recordset.MoveFirst
      Adodc1.Recordset.MoveLast
      adocheck.CursorLocation = adUseClient
      'adocheck.Open "select a0k20, a0k03 from acc0m0, acc0k0 where a0m02 = a0k01 (+) and a0m02 = '" & Adodc1.Recordset.Fields("t0202").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      'modify by sonia 2020/7/28 法律所案源收據收款產生之暫收款,智權人員a0t05改掛介紹人
      'adocheck.Open "select a0k20, a0k03 from acctmp02, acc0k0 where t0202 = a0k01 (+) and t0218 = '" & strUserNum & "' and t0202 = '" & Adodc1.Recordset.Fields("t0202").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      adocheck.Open "select nvl(los04,a0k20) a0k20, a0k03 from acctmp02,acc0k0,acc0j0,lawofficesource where t0202 = a0k01 (+) and t0218 = '" & strUserNum & "' and t0202 = '" & Adodc1.Recordset.Fields("t0202").Value & "' and a0k01=a0j13(+) and a0j01=los06(+) ", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
         adoTaie.Execute "delete from acc0t0 where a0t07 = '" & Text2 & "'"
         'Modify by Morgan 2005/5/13 輸入日期改抓收款日期
         'adoTaie.Execute "insert into acc0t0 (a0t01, a0t02, a0t03, a0t04, a0t07, a0t08, a0t13, a0t11, a0t12, a0t05, a0t06) values ('" & strNo & "', '2', " & FCDate(MaskEdBox1.Text) & ", " & FCDate(MaskEdBox1.Text) & ", '" & Text2 & "', " & Val(Text3) & ", '" & strUserNum & "', " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", null, null)"
         'Modified by Morgan 2014/1/3 +a0t18
         adoTaie.Execute "insert into acc0t0 (a0t01, a0t02, a0t03, a0t04, a0t07, a0t08, a0t13, a0t11, a0t12, a0t05, a0t06,a0t18) values ('" & strNo & "', '2', " & FCDate(Frmacc1150.MaskEdBox1.Text) & ", " & FCDate(MaskEdBox1.Text) & ", '" & Text2 & "', " & Val(Text3) & ", '" & strUserNum & "', " & Val(strSrvDate(2)) & ", " & ServerTime & ", '" & IIf(IsNull(adocheck.Fields("a0k20").Value), "", adocheck.Fields("a0k20").Value) & "', '" & IIf(IsNull(adocheck.Fields("a0k03").Value), "", adocheck.Fields("a0k03").Value) & "','" & strA0L05 & "')"
      Else
         'Modify by Morgan 2005/5/13 輸入日期改抓收款日期
         'adoTaie.Execute "insert into acc0t0 (a0t01, a0t02, a0t03, a0t04, a0t07, a0t08, a0t13, a0t11, a0t12, a0t05, a0t06) values ('" & strNo & "', '2', " & FCDate(MaskEdBox1.Text) & ", " & FCDate(MaskEdBox1.Text) & ", '" & Text2 & "', " & Val(Text3) & ", '" & strUserNum & "', " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", null, null)"
         'Modified by Morgan 2014/1/3 +a0t18
         adoTaie.Execute "insert into acc0t0 (a0t01, a0t02, a0t03, a0t04, a0t07, a0t08, a0t13, a0t11, a0t12, a0t05, a0t06,a0t18) values ('" & strNo & "', '2', " & FCDate(Frmacc1150.MaskEdBox1.Text) & ", " & FCDate(MaskEdBox1.Text) & ", '" & Text2 & "', " & Val(Text3) & ", '" & strUserNum & "', " & Val(strSrvDate(2)) & ", " & ServerTime & ", null, null,'" & strA0L05 & "')"
      End If
      adocheck.Close
   Else
      'Modified by Morgan 2014/1/3 +a0t18
      adoTaie.Execute "insert into acc0t0 (a0t01, a0t02, a0t03, a0t04, a0t07, a0t08, a0t13, a0t11, a0t12, a0t05, a0t06,a0t18) values ('" & strNo & "', '2', " & FCDate(MaskEdBox1.Text) & ", " & FCDate(MaskEdBox1.Text) & ", '" & Text2 & "', " & Val(Text3) & ", '" & strUserNum & "', " & Val(strSrvDate(2)) & ", " & ServerTime & ", null, null,'" & strA0L05 & "')"
   End If
   Text4 = strNo
End Sub


Private Sub Command2_Click()
   If DeleteCheck = True Then
      Acc0k0Delete
   End If
End Sub
'Added by Morgan 2013/12/26
Private Function DeleteCheck() As Boolean
   'Modified by Morgan 2015/5/28 改判斷是否為本次收款開立的發票
   'strExc(0) = "select * from acc0k0,acc431 where a0k01='" & Text1 & "' and nvl(a0k19,0)=0 and axc02(+)=a0k01 and axc01 is not null"
   strExc(0) = "select axc01 from acc0k0,acc431 where a0k01='" & Text1 & "' and axc02(+)=a0k01 and axc03='" & Text2 & "' and axc01 is not null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Modified by Morgan 2015/5/28
      'MsgBox "已開發票未列之請款單不可刪除!" & vbCrLf & vbCrLf & "(若要取消收款請先作廢該發票)", vbExclamation, "收款檢查"
      MsgBox "請款單已開發票 " & RsTemp(0) & "，不可刪除!" & vbCrLf & vbCrLf & "(若要取消收款請先作廢該發票)", vbExclamation, "收款檢查"
      'end 2015/5/28
      Exit Function
   End If
   DeleteCheck = True
End Function

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
   
   Dim douSAmount As Double
   Dim douTAmount As Double
   Dim bolRollback As Boolean

On Error GoTo Checking
'   Select Case ColIndex
'      Case 2
'         DataGrid1.Columns(8) = Val(DataGrid1.Columns(6)) - Val(DataGrid1.Columns(2))
'   End Select
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   bolShow = False
   adocheck.CursorLocation = adUseClient
   adocheck.Open "select sum(nvl(a0m04, 0)+nvl(a0m05, 0)) from acc0m0 where a0m02 = '" & Adodc1.Recordset.Fields("t0202").Value & "' and a0m01 <> '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount <> 0 Then
      If IsNull(adocheck.Fields(0).Value) = False Then
         If adocheck.Fields(0).Value = (Adodc1.Recordset.Fields("t0208").Value + Adodc1.Recordset.Fields("t0209").Value) Then
            DataGrid1.Columns(3).Value = 0
            DataGrid1.Columns(4).Value = 0
            bolShow = False
         End If
      End If
   End If
   adocheck.Close
   
   'Add by Morgan 2007/10/17 讀取銷退資料
   strSql = "select sum(a1u07), sum(a1u09) from acc1u0 where a1u02 ='" & Adodc1.Recordset.Fields("t0202").Value & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      douSAmount = Val("" & RsTemp.Fields(0))
      douTAmount = Val("" & RsTemp.Fields(1))
   End If
   'end 2007/10/17
   
   With DataGrid1
      Select Case ColIndex
      Case 3, 4
         'Modify by Morgan 2006/3/17
         'If Val(.Columns(3).Value) > Val(.Columns(7).Value) Then
         'Modify by Morgan 2007/10/17 要減已收金額及銷退金額
         'If Val(.Columns(3).Value) > Val(.Columns(8).Value) Then
         '   MsgBox MsgText(81), , MsgText(5)
         If Val(.Columns(3).Value) > Val(.Columns(8).Value) - Val(.Columns(10).Value) - douSAmount Then
            MsgBox "本次收款服務費不可大於未收服務費"
            bolRollback = True
         'end 2007/10/17
         End If
         'Modify by Morgan 2006/3/17
         'If Val(.Columns(4).Value) > Val(.Columns(8).Value) Then
         'Modify by Morgan 2007/10/17 要減已收金額及銷退金額
         'If Val(.Columns(4).Value) > Val(.Columns(9).Value) Then
         '   MsgBox MsgText(81), , MsgText(5)
         If Val(.Columns(4).Value) > Val(.Columns(9).Value) - Val(.Columns(11).Value) - douTAmount Then
            MsgBox "本次收款規費不可大於未收規費"
            bolRollback = True
         'end 2007/10/17
         End If
         If bolShow Then
            CalAmount
         End If
      'Added by Morgan 2012/9/19
      Case 5 '個人不可扣繳
         If Val(Format(.Columns(5))) > 0 Then
            'Modify By Sindy 2015/8/6
            'If ChkIsPerson(.Columns(0)) = True Then
            If PUB_ChkIsPerson(.Columns(0)) = True Then
            '2015/8/6 END
               'MsgBox "本收據屬個人不能扣繳!!" 'Revmoed by Morgan 2013/12/26 +公司別檢查,訊息移到函數內
               bolRollback = True
            End If
         End If
      'Add by Morgan 2007/12/21
      Case 6
         If Val(.Columns(6)) < 80 Then
            MsgBox "扣繳年度輸入錯誤！"
            bolRollback = True
         End If
      End Select
   End With
   
   'Modify by Morgan 2007/10/17 若有錯誤時需回復
   'Adodc1.Recordset.UpdateBatch
   If bolRollback = True Then
      Adodc1.Recordset.CancelBatch
   Else
      Adodc1.Recordset.UpdateBatch
   End If
   'end 2007/10/17
   
   Select Case ColIndex
      Case 3
         SumShow
      Case 4
         SumShow
      Case 5
         SumShow
   End Select
   If Val(Text3) < 0 Then
      MsgBox MsgText(82), , MsgText(5)
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub DataGrid1_GotFocus()
'   DataGrid1.Col = 0
'   SendKeys "{RIGHT}"
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim intCounter As Integer

   Select Case KeyCode
      Case vbKeyReturn
         Select Case DataGrid1.col
            Case 1
               SendKeys "{RIGHT}"
            Case 2
               SendKeys "{RIGHT}"
            Case 3
               SendKeys "{RIGHT}"
            Case 4
               SendKeys "{RIGHT}"
            Case 5
               SendKeys "{RIGHT}"
            Case 6
               SendKeys "{RIGHT}"
            Case 7
               For intCounter = 1 To 5
                  SendKeys "{RIGHT}"
               Next intCounter
            Case 12
               SendKeys "{RIGHT}"
            Case 13
               SendKeys "{RIGHT}"
            Case 14
               SendKeys "{RIGHT}"
            Case 15
               SendKeys "{DOWN}"
               For intCounter = 1 To 13
                  SendKeys "{LEFT}"
               Next intCounter
         End Select
         Exit Sub
      Case vbKeyF12
'         If DataGrid1.Col = 3 Or DataGrid1.Col = 4 Then
            CalAmount
'         End If
         Exit Sub
   End Select
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount <> 0 Then
      Text1 = Adodc1.Recordset.Fields("t0202").Value
   End If
End Sub

Private Sub DataGrid1_Validate(Cancel As Boolean)
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133) & "/" & MsgText(107)
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label4 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoTaie.Execute "delete from acctmp02 where t0201 = '" & Text2 & "' and t0218 = '" & strUserNum & "'"
   'Modify by Morgan 2011/9/30 a0m24 不再使用
   'Modified by Morgan 2013/12/26 a0m11,a0m12,a0m13,a0m14 不再使用,t0214 改放 a0k11
   'adoTaie.Execute "insert into acctmp02 (t0201, t0202, t0203, t0204, t0205, t0206, t0207, t0208, t0209, t0210, t0211, t0212, t0213, t0214, t0215, t0216, t0218) " & _
                   "select a0m01, a0m02, a0m03, a0m04, a0m05, a0m06, a0m07, a0m15, a0m16, a0m08, a0m09, a0m10, a0m11, a0m12, a0m13, a0m14, '" & strUserNum & "' from acc0m0 where a0m01 = '" & Text2 & "'"
   adoTaie.Execute "insert into acctmp02 (t0201, t0202, t0203, t0204, t0205, t0206, t0207, t0208, t0209, t0210, t0211, t0212,t0214, t0218) " & _
                   "select a0m01, a0m02, a0m03, a0m04, a0m05, a0m06, a0m07, a0m15, a0m16, a0m08, a0m09, a0m10,a0k11,'" & strUserNum & "' from acc0m0,acc0k0 where a0m01 = '" & Text2 & "' and a0k01(+)=a0m02"
                   
   'Add by Morgan 2011/8/25
   'Modify by Morgan 2011/9/30 +t0217 改即時更新,因不再存檔
   adoTaie.Execute "update acctmp02 set (t0217,t0219,t0220,t0221,t0222)=(select decode(sign(nvl(sum(a1u07),0)+nvl(sum(a1u09),0)),1,'Y') ,nvl(sum(a1u07),0),nvl(sum(a1u09),0),nvl(sum(a1u08),0), nvl(sum(a1u10),0) from acc1u0 where a1u02=t0202) where t0218 = '" & strUserNum & "'"
   'end 2011/8/25
   
   'Added by Morgan 2017/2/16 +全額應扣額 t0223
   adoTaie.Execute "update acctmp02 set t0223=(select sum(decode(a0j07,'Y',nvl(a0j09,0)+nvl(a0j10,0),nvl(a0j09,0)))*0.1 from acc0j0 where a0j13=t0202) where t0218 = '" & strUserNum & "'"
   adoTaie.Execute "update acctmp02 set t0223=(select t0223-nvl(sum(decode(a0j07,'Y',nvl(a1u07,0)+nvl(a1u09,0),nvl(a1u07,0))),0)*0.1 from acc1u0,acc0j0 where a1u02=t0202 and a1u01<>t0201 and a0j01(+)=a1u03 and a0j13(+)=a1u02 ) where t0218 = '" & strUserNum & "'"
   'end 2017/2/16
   
   adoacctmp02.CursorLocation = adUseClient
   adoacctmp02.Open "select * from acctmp02 where t0202 = '" & Text2 & "' and t0218 = '" & strUserNum & "' order by t0201 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0k0.CursorLocation = adUseClient
   adoacc0k0.Open "select * from acc0k0 where a0k01 = '" & Text1 & "' and a0k09 = 0 order by a0k01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0l0.CursorLocation = adUseClient
   adoacc0l0.Open "select * from acc0l0 where a0l01 = '" & Text2 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0m0.CursorLocation = adUseClient
   adoacc0m0.Open "select * from acc0m0 where a0m01 = '" & Text2 & "' and a0m02 = '" & Text1 & "' order by a0m02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acctmp02 where t0218 = '" & strUserNum & "' order by t0202 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
   '轉暫收款單號,欲處理日期
   adocheck.CursorLocation = adUseClient
   adocheck.Open "select a0t01, a0t04 from acc0t0 where a0t07 = '" & Text2 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adocheck.RecordCount <> 0 Then
      If IsNull(adocheck.Fields(0).Value) = False Then
         Text4 = adocheck.Fields(0).Value
      End If
      If IsNull(adocheck.Fields(1).Value) = False Then
         MaskEdBox1.Mask = ""
         MaskEdBox1.Text = CFDate(adocheck.Fields(1).Value)
         MaskEdBox1.Mask = DFormat
      End If
   End If
   FormEnabled
   adocheck.Close
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Add by Morgan 2011/4/15
'更正收據與收文不相符的資料
Private Function FixAccData(pA0j01 As String, pSys As String, pCountry As String, pProperty As String) As Boolean
   Dim stSQL As String, intR As Integer
On Error GoTo ErrHnd
   'Modified by Morgan 2011/12/26 取消 a0j03,a0j20,a0j21
   stSQL = "Update acc0j0 set a0j04='" & pCountry & "' where a0j01='" & pA0j01 & "'"
   cnnConnection.Execute stSQL, intR
   FixAccData = True
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function

'*************************************************
'  儲存資料表(收款資料暫存檔)
'
'*************************************************
Private Sub Acctmp02Save()
Dim douSAmount As Double
Dim douTAmount As Double
Dim douBSAmount As Double   '2007/11/9 add by sonia
Dim douBTAmount As Double   '2007/11/9 add by sonia

On Error GoTo Checking
   douSAmount = 0
   douTAmount = 0
   douBSAmount = 0
   douBTAmount = 0
   If Text1 = MsgText(601) Or Text2 = MsgText(601) Then
      MsgBox MsgText(10), , MsgText(5)
      Exit Sub
   End If
   adocheck.CursorLocation = adUseClient
   adocheck.Open "select * from acc0m0 where a0m01 = '" & Text2 & "' and a0m02 = '" & Text1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adocheck.RecordCount <> 0 Then
      MsgBox MsgText(9), , MsgText(5)
      adocheck.Close
      Exit Sub
   End If
   adocheck.Close
   
   'Added by Morgan 2014/9/18
   strExc(0) = "select * from acc0k0,acc431 where a0k01='" & Text1 & "' and a0k11='J' and axc02(+)=a0k01 and axc01 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strExc(1) = GetInvDate
      If Val(ChangeTDateStringToTString(Frmacc1150.MaskEdBox1)) < Val(strExc(1)) Then
         MsgBox "公司未開發票請款單，收款日期不可早於最後發票日【" & ChangeTStringToTDateString(strExc(1)) & "】!!", vbExclamation, "收款檢查"
         Text1.SetFocus
         Exit Sub
      Else
         strExc(1) = PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1)))
         If Val(ChangeTDateStringToTString(Frmacc1150.MaskEdBox1)) > Val(strExc(1)) Then
            MsgBox "J公司未開發票請款單，收款日期不可晚於下一工作日【" & ChangeTStringToTDateString(strExc(1)) & "】!!", vbExclamation, "收款檢查"
            Text1.SetFocus
            Exit Sub
         End If
      End If
      
      strExc(1) = Val(ChangeTDateStringToTString(Frmacc1150.MaskEdBox1)) \ 100
      strExc(0) = "select * from acc410 where a4101<=" & strExc(1) & " and a4102>=" & strExc(1)
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI <> 1 Then
         MsgBox Left(Frmacc1150.MaskEdBox1, 6) & "月份發票資料尚未建立，無法開立發票!!", vbExclamation, "收款檢查"
         Text1.SetFocus
         Exit Sub
      End If
      
      ZeroTaxCustAlert Text1 'Added by Morgan 2024/11/11 智權公司收款零稅率客戶提醒
   End If
   'end 2014/9/18
   
   'Added by Morgan 2012/7/11--瑞婷
   If Val(Text5) > 0 Then
      'Modify By Sindy 2015/8/6
      'If ChkIsPerson(Text1.Text) = True Then
      If PUB_ChkIsPerson(Text1.Text) = True Then
      '2015/8/6 END
         'MsgBox "本收據屬個人不能扣繳!!" 'Revmoed by Morgan 2013/12/26 +公司別檢查,訊息移到函數內
         Exit Sub
      End If
   End If
   'end 2012/7/11
   
   '2010/3/3 ADD BY SONIA 檢查收據國家或案件性質或本所案號與案件不符時,先改資料再收款
   'Modify by Morgan 2011/4/15 +
   adocheck.CursorLocation = adUseClient
   'Modify by Morgan 2011/10/13 考慮拆收據情形
   'adocheck.Open "select A0J04,A0J03,A0J02,PA09 NATION,CP10,CP01||CP02||CP03||CP04 CASENO,a0j01,cp01 from caseprogress,PATENT, acc0j0 where cp09 = a0j01 (+) and cp60 = '" & Text1 & "' AND CP01 IN ('P','CFP','FCP') AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) UNION " & _
                 "select A0J04,A0J03,A0J02,TM10 NATION,CP10,CP01||CP02||CP03||CP04 CASENO,a0j01,cp01 from caseprogress,TRADEMARK, acc0j0 where cp09 = a0j01 (+) and cp60 = '" & Text1 & "' AND CP01 IN ('T','CFT','FCT','TF') AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) UNION " & _
                 "select A0J04,A0J03,A0J02,LC15 NATION,CP10,CP01||CP02||CP03||CP04 CASENO,a0j01,cp01 from caseprogress,LAWCASE, acc0j0 where cp09 = a0j01 (+) and cp60 = '" & Text1 & "' AND CP01 IN ('L','CFL','FCL','LIN') AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) UNION " & _
                 "select A0J04,A0J03,A0J02,SP09 NATION,CP10,CP01||CP02||CP03||CP04 CASENO,a0j01,cp01 from caseprogress,SERVICEPRACTICE, acc0j0 where cp09 = a0j01 (+) and cp60 = '" & Text1 & "' AND CP01 NOT IN ('P','CFP','FCP','T','CFT','FCT','TF','L','CFL','FCL','LIN','LA') AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) UNION " & _
                 "select A0J04,A0J03,A0J02,'000' NATION,CP10,CP01||CP02||CP03||CP04 CASENO,a0j01,cp01 from caseprogress, acc0j0 where cp09 = a0j01 (+) and cp60 = '" & Text1 & "' AND CP01='LA' ", adoTaie, adOpenStatic, adLockReadOnly
   
   'Modified by Morgan 2011/12/26 取消a0j03
   adocheck.Open "select A0J04,A0J02,PA09 NATION,CP10,CP01||CP02||CP03||CP04 CASENO,a0j01,cp01 from acc0j0,caseprogress,PATENT where a0j13='" & Text1 & "' and cp09(+)=a0j01 AND CP01 IN ('P','CFP','FCP') AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) UNION " & _
                 "select A0J04,A0J02,TM10 NATION,CP10,CP01||CP02||CP03||CP04 CASENO,a0j01,cp01 from acc0j0,caseprogress,TRADEMARK where a0j13='" & Text1 & "' and cp09(+)=a0j01 AND CP01 IN ('T','CFT','FCT','TF') AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) UNION " & _
                 "select A0J04,A0J02,LC15 NATION,CP10,CP01||CP02||CP03||CP04 CASENO,a0j01,cp01 from acc0j0,caseprogress,LAWCASE where a0j13='" & Text1 & "' and cp09(+)=a0j01 AND CP01 IN ('L','CFL','FCL','LIN','ACS') AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) UNION " & _
                 "select A0J04,A0J02,SP09 NATION,CP10,CP01||CP02||CP03||CP04 CASENO,a0j01,cp01 from acc0j0,caseprogress,SERVICEPRACTICE where a0j13='" & Text1 & "' and cp09(+)=a0j01 AND CP01 NOT IN ('P','CFP','FCP','T','CFT','FCT','TF','L','CFL','FCL','LIN','ACS','LA') AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) UNION " & _
                 "select A0J04,A0J02,'000' NATION,CP10,CP01||CP02||CP03||CP04 CASENO,a0j01,cp01 from acc0j0,caseprogress where a0j13='" & Text1 & "' and cp09(+)=a0j01 AND CP01='LA' ", adoTaie, adOpenStatic, adLockReadOnly
                 
   If adocheck.RecordCount <> 0 Then
      'Modify by Morgan 2011/4/15
      '案件性質或申請國家不同改提示確認更新
      'If adocheck.Fields("a0j04").Value <> adocheck.Fields("NATION").Value Or adocheck.Fields("a0j03").Value <> adocheck.Fields("CP10").Value Or adocheck.Fields("a0j02").Value <> adocheck.Fields("CASENO").Value Then
      '   MsgBox "此收據之申請國家或案件性質或本所案號與案件系統不符, 確認後請電腦中心修改再收款 !!!", vbExclamation + vbOKOnly
      '   adocheck.Close
      '   Exit Sub
      'End If
      Do While Not adocheck.EOF
         If adocheck.Fields("a0j02").Value <> adocheck.Fields("CASENO").Value Then
            MsgBox "此收據之本所案號與案件系統不符, 確認後請電腦中心修改再收款 !!!", vbExclamation + vbOKOnly
            adocheck.Close
            Exit Sub
         'Modified by Morgan 2011/12/26 取消a0j03
         ElseIf adocheck.Fields("a0j04").Value <> adocheck.Fields("NATION").Value Then
            If MsgBox("此收據之申請國家與案件系統不符！是否讓系統自動做修正？", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
               If FixAccData(adocheck("a0j01"), adocheck("cp01"), adocheck("NATION"), adocheck("CP10")) = False Then
                  adocheck.Close
                  Exit Sub
               End If
            Else
               adocheck.Close
               Exit Sub
            End If
         End If
         adocheck.MoveNext
      Loop
   End If
   adocheck.Close
   '2010/3/3 END
   
   With Adodc1.Recordset
   
      If .RecordCount <> 0 Then
         .Find "t0201 = '" & Text2 & "'", 0, adSearchForward, 1
         If .EOF = False Then
            .Find "t0202 = '" & Text1 & "'", 0, adSearchForward, .Bookmark
            If .EOF = False Then
               MsgBox MsgText(9), , MsgText(5)
               Exit Sub
            End If
         End If
      End If
      Text3 = Val(Text3) + Val(Text5)
      .AddNew
      .Fields("t0201").Value = Text2
      .Fields("t0202").Value = Text1
      adoacc0k0.Close
      adoacc0k0.CursorLocation = adUseClient
      'Modify By Sindy 2013/11/29
      'adoacc0k0.Open "select * from acc0k0 where a0k01 = '" & Text1 & "' and (a0k09 = 0 or a0k09 is null)", adoTaie, adOpenDynamic, adLockBatchOptimistic
      adoacc0k0.Open "select * from acc0k0,(select a2801,nvl(max(a2802),0) a2802 from acc280 group by a2801) where a0k01 = '" & Text1 & "' and (a0k09 = 0 or a0k09 is null) and a0k04=a2801(+)", adoTaie, adOpenDynamic, adLockBatchOptimistic
      '2013/11/29 END
      If adoacc0k0.RecordCount <> 0 Then
         '92.3.21 MODIFY BY SONIA 扣繳年度改回預設收款年度
         '.Fields("t0207").Value = Val(Mid(CFDate(adoacc0k0.Fields("a0k02").Value), 1, 3))
         '.Fields("t0207").Value = Val(Mid(ACDate(ServerDate), 1, 3))'Remove by Morgan 2011/10/4 改預設前畫面收款日的年度
         '2005/8/15 ADD BY SONIA
         If Not IsNull(adoacc0k0.Fields("a0k16").Value) And adoacc0k0.Fields("a0k16").Value <> 0 Then
            .Fields("t0207").Value = Val(adoacc0k0.Fields("a0k16").Value)
         'Add by Morgan 2011/10/4 改預設前畫面收款日的年度
         Else
            .Fields("t0207").Value = Val(FCDate(Frmacc1150.MaskEdBox1.Text)) \ 10000
            'Add By Sindy 2013/11/29
            If Val("" & adoacc0k0.Fields("a2802").Value) > 0 Then
               If Val(.Fields("t0207").Value) = Val("" & adoacc0k0.Fields("a2802").Value) Then
                  .Fields("t0207").Value = Val(.Fields("t0207").Value) + 1
               End If
            End If
            '2013/11/29 END
         'end 2011/10/4
         End If
         '2005/8/15 END
         '92.3.21 END
         If strCustNo = MsgText(601) Then
            adoacc0e0.CursorLocation = adUseClient
            adoacc0e0.Open "select * from acc0e0 where a0e05 = '2' and a0e06 = '" & adoacc0k0.Fields("a0k03").Value & "' and (a0e15 <> 0 and a0e15 is not null)", adoTaie, adOpenDynamic, adLockBatchOptimistic
            If adoacc0e0.RecordCount <> 0 Then
               strCustNo = MsgText(602)
               MsgBox "此客戶有退票記錄...", , MsgText(5)
               adoacc0e0.Close
               Exit Sub
            End If
            adoacc0e0.Close
         End If
         .Fields("t0206").Value = 0
         If adoaccsum.State = adStateOpen Then
            adoaccsum.Close
         End If
         adoaccsum.CursorLocation = adUseClient
         adoaccsum.Open "select sum(a1u07), sum(a1u09),sum(a1u08), sum(a1u10) from acc1u0 where a1u02 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoaccsum.RecordCount <> 0 Then
            If IsNull(adoaccsum.Fields(0).Value) Then
               douSAmount = 0
            Else
               douSAmount = Val(adoaccsum.Fields(0).Value)
            End If
            If IsNull(adoaccsum.Fields(1).Value) Then
               douTAmount = 0
            Else
               douTAmount = Val(adoaccsum.Fields(1).Value)
            End If
            '2007/11/7 add by sonia
            If IsNull(adoaccsum.Fields(2).Value) Then
               douBSAmount = 0
            Else
               douBSAmount = Val(adoaccsum.Fields(2).Value)
            End If
            If IsNull(adoaccsum.Fields(3).Value) Then
               douBTAmount = 0
            Else
               douBTAmount = Val(adoaccsum.Fields(3).Value)
            End If
            '2007/11/9 end
         Else
            douSAmount = 0
            douTAmount = 0
            '2007/11/7 add by sonia
            douBSAmount = 0
            douBTAmount = 0
            '2007/11/9 end
         End If
         adoaccsum.Close
         If IsNull(adoacc0k0.Fields("a0k06").Value) Then
            .Fields("t0204").Value = 0
            .Fields("t0208").Value = 0
         Else
            .Fields("t0204").Value = adoacc0k0.Fields("a0k06").Value - douSAmount
            .Fields("t0208").Value = adoacc0k0.Fields("a0k06").Value
         End If
         If IsNull(adoacc0k0.Fields("a0k07").Value) Then
            .Fields("t0205").Value = 0
            .Fields("t0209").Value = 0
         Else
            .Fields("t0205").Value = adoacc0k0.Fields("a0k07").Value - douTAmount
            .Fields("t0209").Value = adoacc0k0.Fields("a0k07").Value
         End If
         'Add by Morgan 2011/8/25
         .Fields("t0219").Value = douSAmount '已銷服務費
         .Fields("t0220").Value = douTAmount '已銷規費
         .Fields("t0221").Value = douBSAmount '已退服務費
         .Fields("t0222").Value = douBTAmount '已退規費
      Else
         .CancelBatch
         MsgBox MsgText(28), , MsgText(5)
         Exit Sub
      End If
      adosubsum.CursorLocation = adUseClient
      adosubsum.Open "select sum(a0m04), sum(a0m05) from acc0m0 where a0m02 = '" & Text1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If adosubsum.RecordCount <> 0 Then
         '已收服務費
         If IsNull(adosubsum.Fields(0).Value) Then
            .Fields("t0210").Value = 0
         Else
            '2007/11/9 modify by sonia 應再扣除已退服務費 E09531311
            '.Fields("t0210").Value = adosubsum.Fields(0).Value
            'Modify by Morgan 2011/8/25 改回不扣除已退(維持與0m0,0k0一致)
            '.Fields("t0210").Value = adosubsum.Fields(0).Value - douBSAmount
            .Fields("t0210").Value = adosubsum.Fields(0).Value
         End If
         '已收規費
         If IsNull(adosubsum.Fields(1).Value) Then
            .Fields("t0211").Value = 0
         Else
            '2007/11/9 modify by sonia 應再扣除已退規費
            '.Fields("t0211").Value = adosubsum.Fields(1).Value
            'Modify by Morgan 2011/8/25 改回不扣除已退(維持與0m0,0k0一致)
            '.Fields("t0211").Value = adosubsum.Fields(1).Value - douBTAmount
            .Fields("t0211").Value = adosubsum.Fields(1).Value
         End If
      Else
         .Fields("t0210").Value = 0
         .Fields("t0211").Value = 0
      End If
      adosubsum.Close
      
      'Added by Morgan 2017/2/21
      adosubsum.Open "select sum(decode(a0j07,'Y',nvl(a0j09,0)+nvl(a0j10,0),nvl(a0j09,0)))*0.1 from acc0j0 where a0j13='" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adosubsum.RecordCount <> 0 Then
         .Fields("t0223") = Val("" & adosubsum.Fields(0))
      End If
      adosubsum.Close
      
      adosubsum.Open "select nvl(sum(decode(a0j07,'Y',nvl(a1u07,0)+nvl(a1u09,0),nvl(a1u07,0))),0)*0.1 from acc1u0,acc0j0 where a1u02='" & Text1 & "' and a0j01(+)=a1u03 and a0j13(+)=a1u02", adoTaie, adOpenStatic, adLockReadOnly
      If adosubsum.RecordCount <> 0 Then
         .Fields("t0223") = Val("" & .Fields("t0223")) - Val("" & adosubsum.Fields(0))
      End If
      adosubsum.Close
      'end 2017/2/21
      
      
'Modify by Morgan 2007/10/17 判斷時要扣除已收
'      '溢收款<=應收規費
'      If Val(Text3) <= (Val(.Fields("t0209").Value) - douTAmount) Then
'         If Val(.Fields("t0209").Value) - Val(.Fields("t0211").Value) = 0 Then
'            '應收服務費<已收服務費
'            If (Val(.Fields("t0208").Value) - Val(.Fields("t0210").Value)) <= 0 Then
'               .Fields("t0204").Value = 0
'               .Fields("t0205").Value = 0
'            '應收服務費>已收服務費
'            Else
'               '93.11.9 MODIFY BY SONIA
'               '.Fields("t0204").Value = Val(Text3)
'               '.Fields("t0205").Value = 0
'               If Val(Text3) <= (Val(.Fields("t0208").Value) - Val(.Fields("t0210").Value)) Then
'                  .Fields("t0204").Value = Val(Text3)
'                  .Fields("t0205").Value = 0
'               Else
'                  .Fields("t0204").Value = (Val(.Fields("t0208").Value) - Val(.Fields("t0210").Value))
'                  .Fields("t0205").Value = 0
'               End If
'               '93.11.9 END
'            End If
'         Else
'            .Fields("t0204").Value = 0
'            .Fields("t0205").Value = Val(Text3)
'         End If
'      '溢收款>應收規費
'      Else
'
'         If (Val(.Fields("t0209").Value) - Val(.Fields("t0211").Value)) <= 0 Then
'            If (Val(.Fields("t0208").Value) - Val(.Fields("t0210").Value)) <= 0 Then
'               .Fields("t0204").Value = 0
'               .Fields("t0205").Value = 0
'            Else
'               If Val(Text3) >= Val(.Fields("t0209").Value) Then
'                  If Val(Text3) <= (Val(.Fields("t0209").Value) - douTAmount + Val(.Fields("t0208").Value) - douSAmount) Then
'                     .Fields("t0204").Value = Val(Text3) - (Val(.Fields("t0209").Value) - douTAmount)
'                  Else
'                     .Fields("t0204").Value = Val(.Fields("t0208").Value) - douSAmount
'                  End If
'                  .Fields("t0205").Value = Val(.Fields("t0209").Value) - douTAmount
'               Else
'                  .Fields("t0204").Value = 0
'                  .Fields("t0205").Value = Val(Text3)
'               End If
'            End If
'         Else
'            .Fields("t0205").Value = Val(.Fields("t0209").Value) - douTAmount
'            If Val(Text3) <= (Val(.Fields("t0209").Value) - douTAmount + Val(.Fields("t0208").Value) - douSAmount) Then
'               .Fields("t0204").Value = Val(Text3) - (Val(.Fields("t0209").Value) - douTAmount)
'            Else
'               .Fields("t0204").Value = Val(.Fields("t0208").Value) - douSAmount
'            End If
'         End If
'      End If
      '溢收款<未收規費(應收-已銷-已收(扣除已退))
      'Modify by Morgan 2011/8/25 未收=應收-已銷-(已收-已退)
      'If Val(Text3) < (Val(.Fields("t0209").Value) - douTAmount - .Fields("t0211")) Then
      If Val(Text3) < .Fields("t0209").Value - .Fields("t0220").Value - (.Fields("t0211") - .Fields("t0222")) Then
         .Fields("t0204").Value = 0
         .Fields("t0205").Value = Val(Text3)
      '溢收款>=未收規費
      Else
         '溢收款-本次收款規費
         'Modify by Morgan 2011/8/25 未收=應收-已銷-(已收-已退)
         '.Fields("t0205").Value = Val(.Fields("t0209").Value) - douTAmount - .Fields("t0211")
         'If Val(Text3) - .Fields("t0205").Value < Val(.Fields("t0208").Value) - douSAmount - .Fields("t0210") Then
         .Fields("t0205").Value = .Fields("t0209").Value - .Fields("t0220").Value - (.Fields("t0211") - .Fields("t0222"))
         If Val(Text3) - .Fields("t0205").Value < .Fields("t0208").Value - .Fields("t0219").Value - (.Fields("t0210") - .Fields("t0221")) Then
            .Fields("t0204").Value = Val(Text3) - .Fields("t0205").Value
         Else
            'Modify by Morgan 2011/8/25
            '.Fields("t0204").Value = Val(.Fields("t0208").Value) - douSAmount - .Fields("t0210")
            .Fields("t0204").Value = .Fields("t0208").Value - .Fields("t0219").Value - (.Fields("t0210") - .Fields("t0221"))
         End If
      End If
'end 2007/10/17
      
      If Text5 <> MsgText(601) Then
         'Modified by Morgan 2014/8/5 扣繳額必須為整數(四捨五入)
         '.Fields("t0206").Value = Val(Text5)
         .Fields("t0206").Value = Round(Val(Text5))
      Else
         .Fields("t0206").Value = Null
      End If
      If IsNull(adoacc0k0.Fields("a0k20").Value) = False Then
         .Fields("t0212").Value = adoacc0k0.Fields("a0k20").Value
      End If
      
      'Modified by Morgan 2013/12/27 a0m11,a0m12,a0m13,a0m14 不再使用,t0214 改放 a0k11
      '.Fields("t0213").Value = .Fields("t0204").Value
      '.Fields("t0215").Value = 0
      '.Fields("t0216").Value = 0
      .Fields("t0214").Value = adoacc0k0.Fields("a0k11").Value
      'end 2013/12/27
      
      adocheck.CursorLocation = adUseClient
      'adocheck.Open "select * from acc0j0 where a0j13 = '" & adoacc0k0.Fields("a0k01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      adocheck.Open "select sum(nvl(a1u07, 0)+nvl(a1u09, 0)) from acc1u0 where a1u02 = '" & adoacc0k0.Fields("a0k01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
         If adocheck.Fields(0).Value > 0 Then
            .Fields("t0217").Value = MsgText(602)
         End If
         'If Mid(adocheck.Fields("a0j02").Value, 1, 1) = "L" And Mid(adocheck.Fields("a0j02").Value, 2, 1) <> "A" Then
         '   .Fields("t0217").Value = "414101"
         'End If
      End If
      adocheck.Close
      
      .Fields("t0218").Value = strUserNum
      .UpdateBatch
   End With
   
   'Add by Morgan 2005/11/2
   Text1 = "E" & Format(Val(Mid(Text1, 2)) + 1, "00000000")
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
Private Sub AdodcRefresh()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   'Modify by Morgan 2006/10/4 不排序,用原來的順序
   'adoadodc1.Open "select * from acctmp02 where t0218 = '" & strUserNum & "' order by t0202 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.Open "select * from acctmp02 where t0218 = '" & strUserNum & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
On Error GoTo Checking

   'Added by Sindy 2021/12/20 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
   If PUB_ChkTrackMode = False Then
       Exit Sub
   End If
   '2021/12/20 END
   
   Select Case KeyCode
      Case vbKeyInsert
         Acctmp02Save
         DataGrid1.Refresh
         SumShow
         'Modify by Morgan 2005/11/2不清除
         'Text1 = MsgText(802)
         Text5 = MsgText(601)
         Text1.SetFocus
         Text1_GotFocus
      Case vbKeyF12
         CalAmount
         Exit Sub
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133) & "/" & MsgText(107)
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgBox(5)
End Sub

'*************************************************
'  刪除收款資料項目
'
'*************************************************
Private Sub Acc0k0Delete()
   Dim lngPos As Long
   lngPos = Adodc1.Recordset.AbsolutePosition
   
On Error GoTo Checking

   If Adodc1.Recordset.RecordCount <> 0 Then
   
      'Added by Morgan 2016/2/18
      '檢查收據若退費金額>已收金額時不可刪除
      If Adodc1.Recordset.Fields("t0221") > Adodc1.Recordset.Fields("t0210") Or Adodc1.Recordset.Fields("t0222") > Adodc1.Recordset.Fields("t0211") Then
         MsgBox "本收據已有退費記錄，不可刪除！", vbExclamation
         Exit Sub
      End If
      'end 2016/2/18
      
      'Added by Morgan 2025/2/24 檢查是否有補扣繳
      adoTaie.Execute "update acc1u0 set a1u04=a1u04  where a1u02 = '" & Text1 & "' and a1u01=a1u03", intI
      If intI > 0 Then
         MsgBox "本收據已有補扣繳紀錄，不可刪除！", vbExclamation
         Exit Sub
      End If
      'end 2025/2/24
      
'Removed by Morgan 2016/3/21 不是部分收款的會沒刪到,改移到下面做
'      'Add by Morgan 2007/2/12 扣繳資料要刪除或修改
'      'Modified by Morgan 2011/10/24 考慮拆收據情形
'      '更新部份收款資料的扣繳資料
'      strSql = "update acc1v0 set a1v05='Y',(a1v06,a1v07)=(select sum(a1u06),a1v04-sum(a1u06) from acc1u0 where a1u03=a1v01 and a1u01<>'" & Text2 & "')" & _
'         " where (a1v01,a1v02) in (select a1u03,a1u02 from acc1u0 where a1u01 = '" & Text2 & "' and a1u02 = '" & Text1 & "')"
'      adoTaie.Execute strSql, intI
'      If intI = 0 Then
'         '刪除沒有收款資料的扣繳資料
'         strSql = "delete from acc1v0" & _
'            " where (a1v01,a1v02) in (select a1u03,a1u02 from acc1u0 where a1u01 = '" & Text2 & "' and a1u02 = '" & Text1 & "')" & _
'            " and not exists(select * from acc1u0 where a1u03=a1v01 and a1u01<>'" & Text2 & "')"
'         adoTaie.Execute strSql, intI
'      End If
'      'End 2007/2/12
      
      adoTaie.Execute "delete from acc1u0 where a1u01 = '" & Text2 & "' and a1u02 = '" & Text1 & "'", intI
     

      
      adoTaie.Execute "delete from acctmp02 where t0201 = '" & Text2 & "' and t0202 = '" & Text1 & "'", intI
      adoTaie.Execute "delete from acc0m0 where a0m01 = '" & Text2 & "' and a0m02 = '" & Text1 & "'", intI
      adoTaie.Execute "update acc0k0 set (a0k17, a0k18) = (select nvl(sum(a0m04), 0),nvl(sum(a0m05), 0) from acc0m0 where a0m02 = a0k01) where a0k01 = '" & Text1 & "'", intI
      'Modify by Morgan 2011/10/4 考慮拆收據情形
      'adoTaie.Execute "update caseprogress set cp73 = (select nvl(sum(a1u04), 0) from acc1u0 where a1u03 = cp09), cp74 = (select nvl(sum(a1u05), 0) from acc1u0 where a1u03 = cp09) where cp60 = '" & Text1 & "'"
      'adoTaie.Execute "update caseprogress set cp75 = cp73+cp74, cp79 = cp16-cp73-cp74-nvl(cp77, 0)+nvl(cp78, 0) where cp60 = '" & Text1 & "'"
      'Modified by Morgan 2025/7/11 +cp76
      adoTaie.Execute "update caseprogress set (cp73,cp74,cp76) = (select nvl(sum(a1u04), 0),nvl(sum(a1u05), 0),nvl(sum(a1u06), 0) from acc1u0 where a1u03 = cp09) where cp09 in (select a0j01 from acc0j0 where a0j13 = '" & Text1 & "')", intI
      adoTaie.Execute "update caseprogress set cp75 = cp73+cp74, cp79 = cp16-cp73-cp74-nvl(cp77, 0)+nvl(cp78, 0) where cp09 in (select a0j01 from acc0j0 where a0j13 = '" & Text1 & "')", intI
      'end 2011/10/4
      
      'Added by Morgan 2016/3/21
      '刪除沒有收款記錄的扣繳資料
      strSql = "delete from acc1v0 where a1v02 = '" & Text1 & "' and not exists(select * from acc1u0 where a1u02=a1v02 and a1u03=a1v01 and substr(a1u01,1,1)='F')"
      adoTaie.Execute strSql, intI
      '更新部份收款資料的扣繳資料
      'Modified by Morgan 2020/1/6 a1v05 應該判斷是否有未收金額而非扣繳
      'strSql = "update acc1v0 set (a1v05,a1v06,a1v07)=(select decode(a1v04-sum(a1u06),0,'N','Y'),sum(a1u06),a1v04-sum(a1u06) from acc1u0 where a1u02=a1v02 and a1u03=a1v01) where a1v02 = '" & Text1 & "'"
      strSql = "update acc1v0 set (a1v05,a1v06,a1v07)=(select decode(max(nvl(a0j09,0)+nvl(a0j10,0))-sum(nvl(a1u07,0)+nvl(a1u09,0))-sum(nvl(a1u04,0)+nvl(a1u05,0)-nvl(a1u08,0)-nvl(a1u10,0)),0,'N','Y')" & _
         ",sum(a1u06),a1v04-sum(a1u06) from acc0j0,acc1u0 where a0j01=a1v01 and a0j13=a1v02 and a1u02(+)=a0j13 and a1u03(+)=a0j01)" & _
         " where a1v02 = '" & Text1 & "'"
      adoTaie.Execute strSql, intI
      'end 2016/3/21
      
      PUB_UpdateReceiptStatus Text1 'Added by Morgan 2017/8/16
   End If
   AdodcRefresh
   'Modify by Morgan 2005/11/2
   'Text1 = ""
   'Modify by Morgan 2006/10/4 預設在下一筆收據--瑞婷
   If Adodc1.Recordset.RecordCount > 0 Then
      If lngPos > 1 And Adodc1.Recordset.RecordCount >= lngPos Then
         Adodc1.Recordset.Move lngPos - 1
      Else
         Adodc1.Recordset.MoveFirst
      End If
      Text1 = Adodc1.Recordset.Fields("t0202")
   Else
      Text1 = "E"
   End If
   Text1.SetFocus
   Text1_GotFocus
   
   SumShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  計算並顯示溢收金額
'
'*************************************************
Private Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(t0204), sum(t0205), sum(t0206) from acctmp02 where t0201 = '" & Text2 & "' and t0218 = '" & strUserNum & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         dou0 = 0
      Else
         dou0 = Val(adoaccsum.Fields(0).Value)
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         dou1 = 0
      Else
         dou1 = Val(adoaccsum.Fields(1).Value)
      End If
      If IsNull(adoaccsum.Fields(2).Value) Then
         dou2 = 0
      Else
         dou2 = Val(adoaccsum.Fields(2).Value)
      End If
      Text3 = Format(dblTotal - (dou0 + dou1) + dou2, FAmount)
   Else
      Text3 = Format(dblTotal, FAmount)
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  儲存資料表(國內收款資料(交易檔))
'
'*************************************************
Private Sub Acc0m0Save()
On Error GoTo Checking
   adoacc0m0.AddNew
   adoacc0m0.Fields("a0m01").Value = Adodc1.Recordset.Fields("t0201").Value
   adoacc0m0.Fields("a0m02").Value = Adodc1.Recordset.Fields("t0202").Value
   If IsNull(Adodc1.Recordset.Fields("t0203").Value) Then
      adoacc0m0.Fields("a0m03").Value = Null
   Else
      adoacc0m0.Fields("a0m03").Value = Adodc1.Recordset.Fields("t0203").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("t0204").Value) Then
      adoacc0m0.Fields("a0m04").Value = 0
   Else
      adoacc0m0.Fields("a0m04").Value = Adodc1.Recordset.Fields("t0204").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("t0205").Value) Then
      adoacc0m0.Fields("a0m05").Value = 0
   Else
      adoacc0m0.Fields("a0m05").Value = Adodc1.Recordset.Fields("t0205").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("t0206").Value) Then
      adoacc0m0.Fields("a0m06").Value = 0
   Else
      adoacc0m0.Fields("a0m06").Value = Adodc1.Recordset.Fields("t0206").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("t0207").Value) Then
      adoacc0m0.Fields("a0m07").Value = Null
   Else
      adoacc0m0.Fields("a0m07").Value = Adodc1.Recordset.Fields("t0207").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("t0210").Value) Then
      adoacc0m0.Fields("a0m08").Value = Null
   Else
      adoacc0m0.Fields("a0m08").Value = Adodc1.Recordset.Fields("t0210").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("t0211").Value) Then
      adoacc0m0.Fields("a0m09").Value = 0
   Else
      adoacc0m0.Fields("a0m09").Value = Adodc1.Recordset.Fields("t0211").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("t0212").Value) Then
      adoacc0m0.Fields("a0m10").Value = Null
   Else
      adoacc0m0.Fields("a0m10").Value = Adodc1.Recordset.Fields("t0212").Value
   End If
   
'Removed by Morgan 2013/12/26 不再使用
'   If IsNull(Adodc1.Recordset.Fields("t0213").Value) Then
'      adoacc0m0.Fields("a0m11").Value = 0
'   Else
'      adoacc0m0.Fields("a0m11").Value = Adodc1.Recordset.Fields("t0213").Value
'   End If
'   If IsNull(Adodc1.Recordset.Fields("t0214").Value) Then
'      adoacc0m0.Fields("a0m12").Value = Null
'   Else
'      adoacc0m0.Fields("a0m12").Value = Adodc1.Recordset.Fields("t0214").Value
'   End If
'   If IsNull(Adodc1.Recordset.Fields("t0215").Value) Then
'      adoacc0m0.Fields("a0m13").Value = 0
'   Else
'      adoacc0m0.Fields("a0m13").Value = Adodc1.Recordset.Fields("t0215").Value
'   End If
'   If IsNull(Adodc1.Recordset.Fields("t0216").Value) Then
'      adoacc0m0.Fields("a0m14").Value = 0
'   Else
'      adoacc0m0.Fields("a0m14").Value = Adodc1.Recordset.Fields("t0216").Value
'   End If
'end 2013/12/26

   If IsNull(Adodc1.Recordset.Fields("t0208").Value) Then
      adoacc0m0.Fields("a0m15").Value = 0
   Else
      adoacc0m0.Fields("a0m15").Value = Adodc1.Recordset.Fields("t0208").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("t0209").Value) Then
      adoacc0m0.Fields("a0m16").Value = 0
   Else
      adoacc0m0.Fields("a0m16").Value = Adodc1.Recordset.Fields("t0209").Value
   End If
   adoacc0m0.Fields("a0m17").Value = Val(strSrvDate(2))
   adoacc0m0.Fields("a0m18").Value = ServerTime
   adoacc0m0.Fields("a0m19").Value = strUserNum
   
'Remove by Morgan 2011/9/30 a0m24 不再使用
'   If IsNull(Adodc1.Recordset.Fields("t0217").Value) Then
'      adoacc0m0.Fields("a0m24").Value = Null
'   Else
'      adoacc0m0.Fields("a0m24").Value = Adodc1.Recordset.Fields("t0217").Value
'   End If
   
   adoacc0m0.UpdateBatch
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Memo by Morgan 2013/12/19 刪除 Acc1p0Save 程式碼

'*************************************************
'  物件使用狀態
'
'*************************************************
Private Sub FormEnabled()
   If Text4 = "" Then
      Command1.Enabled = True
   Else
      Command1.Enabled = False
   End If
End Sub

'*************************************************
'  收款金額不足時之計算
'
'*************************************************
Public Sub CalAmount()
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   'Modified by Morgan 2017/2/16 部分扣繳也要可以分配
   If Val(DataGrid1.Columns(3).Value) < Val(DataGrid1.Columns(8).Value) Or Val(DataGrid1.Columns(4).Value) < Val(DataGrid1.Columns(9).Value) Or (Val("" & Adodc1.Recordset.Fields("t0206")) > 0 And Val("" & Adodc1.Recordset.Fields("t0206")) <> Val("" & Adodc1.Recordset.Fields("t0223"))) Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(106)
      If adocheck.State = adStateOpen Then adocheck.Close
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select a1u01 from acc1u0 where a1u01 = '" & Adodc1.Recordset.Fields("t0201").Value & "' and a1u02 = '" & Adodc1.Recordset.Fields("t0202").Value & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If adocheck.RecordCount = 0 Then
         adocheck.Close
         adocheck.CursorLocation = adUseClient
         'Modify by Morgan 2011/10/4 考慮拆收據情形改先抓0j0
         'adocheck.Open "select * from caseprogress where cp60 = '" & Adodc1.Recordset.Fields("t0202").Value & "' order by cp01||cp02||cp03||cp04 asc, cp09 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         adocheck.Open "select * from acc0j0 where a0j13 = '" & Adodc1.Recordset.Fields("t0202").Value & "' order by a0j02 asc, a0j01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         Do While adocheck.EOF = False
            'Modify by Morgan 2011/10/4 考慮拆收據情形改抓0j0,已收金額改抓 1u0
            'If IsNull(adocheck.Fields("cp16").Value) Then
            '   douAmount = 0
            'Else
            '   douAmount = adocheck.Fields("cp16").Value
            'End If
            'If IsNull(adocheck.Fields("cp17").Value) Then
            '   douFAmount = 0
            'Else
            '   douFAmount = adocheck.Fields("cp17").Value
            'End If
            ''If adocheck.Fields("cp78").Value > 0 Then
            ''   douAmount = douAmount - douFAmount
            ''Else
            '   If IsNull(adocheck.Fields("cp73").Value) = False Then
            '      douAmount = douAmount - Val(adocheck.Fields("cp73").Value) - douFAmount
            '   End If
            '   If IsNull(adocheck.Fields("cp74").Value) = False Then
            '      douFAmount = douFAmount - Val(adocheck.Fields("cp74").Value)
            '   End If
            ''End If
            douAmount = Val("" & adocheck.Fields("a0j09").Value)
            douFAmount = Val("" & adocheck.Fields("a0j10").Value)
               
            If adoaccsum.State = adStateOpen Then adoaccsum.Close
            adoaccsum.CursorLocation = adUseClient
            '2005/10/21 MODIFY BY SONIA
            'adoaccsum.Open "select sum(a1u07), sum(a1u09) from acc1u0 where a1u02 = '" & Adodc1.Recordset.Fields("t0202").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
            '2007/11/9 modify by sonia
            'adoaccsum.Open "select sum(a1u07), sum(a1u09) from acc1u0 where a1u02 = '" & Adodc1.Recordset.Fields("t0202").Value & "' AND a1u03 = '" & adocheck.Fields("cp09").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
            'Modify by Morgan 2011/10/12 +sum(a1u04),sum(a1u05)
            'Modified by Morgan 2014/1/3 +NVL判斷否則若為Null會出錯
            'adoaccsum.Open "select sum(a1u07), sum(a1u09), sum(a1u08), sum(a1u10),sum(a1u04),sum(a1u05) from acc1u0 where a1u02 = '" & Adodc1.Recordset.Fields("t0202").Value & "' AND a1u03 = '" & adocheck.Fields("a0j01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
            adoaccsum.Open "select sum(a1u07), sum(a1u09), nvl(sum(a1u08),0), nvl(sum(a1u10),0),sum(a1u04),sum(a1u05) from acc1u0 where a1u02 = '" & Adodc1.Recordset.Fields("t0202").Value & "' AND a1u03 = '" & adocheck.Fields("a0j01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
            '2007/11/9 end
            '2005/10/21 END
            If adoaccsum.RecordCount <> 0 Then
               If IsNull(adoaccsum.Fields(0).Value) = False Then
                  '2007/11/9 modify by sonia 再加回銷退
                  'douAmount = douAmount - Val(adoaccsum.Fields(0).Value)
                  douAmount = douAmount - Val(adoaccsum.Fields(0).Value) + Val(adoaccsum.Fields(2).Value)
               End If
               If IsNull(adoaccsum.Fields(1).Value) = False Then
                  '2007/11/9 modify by sonia 再加回銷退
                  'douFAmount = douFAmount - Val(adoaccsum.Fields(1).Value)
                  douFAmount = douFAmount - Val(adoaccsum.Fields(1).Value) + Val(adoaccsum.Fields(3).Value)
               End If
               'Add by Morgan 2011/10/12 考慮拆收據情形已收金額改抓 1u0
               If IsNull(adoaccsum.Fields(4).Value) = False Then
                  douAmount = douAmount - adoaccsum.Fields(4).Value
               End If
               If IsNull(adoaccsum.Fields(5).Value) = False Then
                  douFAmount = douFAmount - adoaccsum.Fields(5).Value
               End If
               'end 2011/10/12
            End If
            adoaccsum.Close
            
            If adoacc1u0.State = adStateOpen Then adoacc1u0.Close
            adoacc1u0.CursorLocation = adUseClient
            'Modify by Morgan 2011/10/4 考慮拆收據情形改抓0j0
            adoacc1u0.Open "select a1u01 from acc1u0 where a1u01 = '" & Adodc1.Recordset.Fields("t0201").Value & "' and a1u02 = '" & adocheck.Fields("a0j13").Value & "' and a1u03 = '" & adocheck.Fields("a0j01").Value & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
            If adoacc1u0.RecordCount = 0 Then
               adoTaie.Execute "insert into acc1u0 values ('" & Adodc1.Recordset.Fields("t0201").Value & "', '" & Adodc1.Recordset.Fields("t0202").Value & "', '" & adocheck.Fields("a0j01").Value & "', " & douAmount & ", " & douFAmount & ", 0, 0, 0, 0, 0)"
            End If
            adoacc1u0.Close
            adocheck.MoveNext
         Loop
      End If
      adocheck.Close
      strCon1 = Adodc1.Recordset.Fields("t0201").Value
      strCon2 = Adodc1.Recordset.Fields("t0202").Value
      Frmacc1153.Show
      Me.Enabled = False
   End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim douTax As Double
Dim strTrans As String
Dim strInvoiceNo As String
Dim iRow As Integer
Dim strMsg As String
Dim bolAll As Boolean '是否全額收款
Dim acc1150Y As Double  'Add by Amy 2013/06/21 取Frmacc1150.MaskEdBox1(收款日)及t2070 年度
Dim stSQL As String, intR As Integer, rsQuery As ADODB.Recordset  'add by sonia 2017/3/16
Dim stAccNo As String 'Added by Morgan 2023/8/16

'Added by Morgan 2020/4/15
Dim strA1P01s As String '有傳票號的公司別
Dim strA1P22s As String '傳票號
Dim arrA1p22() As String '傳票號
Dim intPos As Integer
'end 2020/4/15

Dim lngAmt As Long '分錄金額 Added by Morgan 2021/3/10
Dim bolACS As Boolean 'Added by Morgan 2025/4/22 是否ACS案

strSupportCaseList = "" 'Add by Morgan 2005/11/2
strXFeeCaseList = "" 'Add by Morgan 2011/4/26
strAssignCaseList = "" 'Add by Morgan 2011/4/26
strProFeeCaseList = "" 'Add by Morgan 2021/1/20

On Error GoTo Checking
   strTrans = MsgText(601)
   If Adodc1.Recordset.RecordCount <> 0 Then
      If Val(Text3) < 0 Then
         tool3_enabled
         MsgBox MsgText(89), , MsgText(5)
         Cancel = 1
         Exit Sub
      Else
         If Val(Text3) > 0 And Text4 = "" Then
            tool3_enabled
            MsgBox MsgText(132), , MsgText(5)
            Cancel = 1
            Exit Sub
         End If
      End If
      
      With Adodc1.Recordset
         .MoveFirst
         Do While Not .EOF
            'Modify by Morgan 2011/8/25 未收=應收-已銷-(已收-已退)
            'If Val("" & .Fields("t0204")) > Val("" & .Fields("t0208")) - Val("" & .Fields("t0210")) Then
            '   MsgBox "【" & .Fields("t0202") & "】本次服務費>應收服務費-已收服務費！", vbExclamation
            If Val("" & .Fields("t0204")) > Val("" & .Fields("t0208").Value) - Val("" & .Fields("t0219").Value) - (Val("" & .Fields("t0210")) - Val("" & .Fields("t0221"))) Then
               MsgBox "【" & .Fields("t0202") & "】本次服務費>應收服務費-已銷服務費-(已收服務費-已退服務費)！", vbExclamation
               tool3_enabled
               Cancel = 1
               Exit Sub
            End If
            
            'Modify by Morgan 2011/8/25 未收=應收-已銷-(已收-已退)
            'If Val("" & .Fields("t0205")) > Val("" & .Fields("t0209")) - Val("" & .Fields("t0211")) Then
            '   MsgBox "【" & .Fields("t0202") & "】本次規費>應收規費-已收規費！", vbExclamation
            If Val("" & .Fields("t0205")) > Val("" & .Fields("t0209").Value) - Val("" & .Fields("t0220").Value) - (Val("" & .Fields("t0211")) - Val("" & .Fields("t0222"))) Then
               MsgBox "【" & .Fields("t0202") & "】本次規費>應收規費-已銷規費-(已收規費-已退規費)！", vbExclamation
               tool3_enabled
               Cancel = 1
               Exit Sub
            End If
            
            'Added by Morgan 2023/10/19
            'Removed by Morgan 2023/10/20 取消--瑞婷/婉莘
            'If Val("" & .Fields("t0204")) > 0 And Val("" & .Fields("t0205")) <> Val("" & .Fields("t0209").Value) - Val("" & .Fields("t0220").Value) - (Val("" & .Fields("t0211")) - Val("" & .Fields("t0222"))) Then
            '   MsgBox "【" & .Fields("t0202") & "】部分收款須先沖銷規費！", vbExclamation
            '   tool3_enabled
            '   Cancel = 1
            '   Exit Sub
            'End If
            'end 2023/10/19
            
            'Add By Sindy 2013/11/29
            '扣繳確認年度
            adocheck.CursorLocation = adUseClient
            adocheck.Open "select a0k01,a0k04,a2802 from acc0k0,(select a2801,nvl(max(a2802),0) a2802 from acc280 group by a2801) where a0k01 = '" & Adodc1.Recordset.Fields("t0202").Value & "' and a0k04=a2801(+)", adoTaie, adOpenStatic, adLockReadOnly
            If adocheck.RecordCount <> 0 Then
               If Val("" & adocheck.Fields("a2802").Value) > 0 Then
                  If Val("" & .Fields("t0207")) <= Val("" & adocheck.Fields("a2802").Value) Then
                     MsgBox "此收據抬頭【" & adocheck.Fields("a0k04").Value & "】之" & adocheck.Fields("a2802").Value & "年扣繳已確認，不可再輸入" & Val("" & .Fields("t0207")) & "年之扣繳！", vbExclamation
                     tool3_enabled
                     adocheck.Close
                     Cancel = 1
                     Exit Sub
                  End If
               End If
            End If
            adocheck.Close
            '2013/11/29 END
            
            'Added by Morgan 2015/7/7
            'Modified by Morgan 2017/2/16 部分扣繳也要分配
            'If Not (.Fields("t0204").Value + .Fields("t0205").Value = .Fields("t0208").Value + .Fields("t0209").Value) Then
            If Not (.Fields("t0204").Value + .Fields("t0205").Value = .Fields("t0208").Value + .Fields("t0209").Value And (Val("" & .Fields("t0206")) = 0 Or Val("" & .Fields("t0206")) = Val("" & .Fields("t0223")))) Then
            'end 2017/2/16
               strExc(0) = "select sum(a1u04),sum(a1u05),sum(a1u06) from acc1u0 where a1u01='" & Text2 & "' and a1u02='" & Adodc1.Recordset.Fields("t0202").Value & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If Val("" & RsTemp(0)) <> Val("" & .Fields("t0204")) Or Val("" & RsTemp(1)) <> Val("" & .Fields("t0205")) Or Val("" & RsTemp(2)) <> Val("" & .Fields("t0206")) Then
                     MsgBox "收據 " & Adodc1.Recordset.Fields("t0202").Value & " 為部分收款或部分扣繳，請按 F12 依收文號分配收款!!", vbExclamation
                     tool3_enabled
                     Cancel = 1
                     Exit Sub
                  End If
               End If
            End If
            'end 2015/7/7
            .MoveNext
         Loop
      End With
      
      adoTaie.BeginTrans
      
      strTrans = MsgText(602)
      F5639NO = "" 'add by sonia 2016/10/18
      
      'Added by Morgan 2013/12/27
      'Modified by Morgan 2020/4/15
      'strA1P22_1 = "null"
      'strA1P22_J = "null"
      strA1P01s = ""
      strA1P22s = ""
      'end 2020/4/15
      
      '紀錄傳票號
      adoaccsum.CursorLocation = adUseClient
      'Modified by Morgan 2013/12/19 收款會有J公司,取消 a1p01='1' 條件,一張收款單有可能有兩張傳票號分別要記錄
      adoaccsum.Open "select distinct a1p22,a1p01 from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "' and a1p22 is not null", adoTaie, adOpenStatic, adLockReadOnly
      If adoaccsum.RecordCount <> 0 Then
         'Modified by Morgan 2013/12/19
         'stra1p22 = "'" & adoaccsum.Fields("a1p22").Value & "'"
         Do While Not adoaccsum.EOF
            'Modified by Morgan 2020/4/15 考慮會有3家作帳公司別
            'If adoaccsum.Fields("a1p01") = "J" Then
            '   strA1P22_J = "'" & adoaccsum.Fields("a1p22").Value & "'"
            'Else
            '   strA1P22_1 = "'" & adoaccsum.Fields("a1p22").Value & "'"
            'End If
            strA1P01s = strA1P01s & adoaccsum.Fields("a1p01")
            strA1P22s = strA1P22s & adoaccsum.Fields("a1p22") & ";"
            'end 2020/4/15
            adoaccsum.MoveNext
         Loop
         'end 2013/12/19
         
         arrA1p22() = Split(strA1P22s, ";") 'Added by Morgan 2020/4/15
         
         stra1p27 = "'Y'"
         
         'Modified by Morgan 2013/12/19 收款會有J公司,取消 a1p01='1' 條件
         adoTaie.Execute "update acc1p0 set a1p27 = 'Y' where a1p02 = 'A' and a1p04 = '" & Text2 & "'", intI
      Else
         stra1p22 = "null"
         stra1p27 = "null"
      End If
      adoaccsum.Close
      
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(105)
      adoTaie.Execute "delete from acc0m0 where a0m01 = '" & Text2 & "'"
      'Modified by Morgan 2013/12/19 收款會有J公司,取消 a1p01='1' 條件
      adoTaie.Execute "delete from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "' and a1p07 = 0"
      adoTaie.Execute "delete from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "' and a1p05 = '1203'"
      'Add by Morgan 2006/5/30
      'Modified by Morgan 2015/8/21 改刪除所有收入(4開頭)及規費(2201開頭)科目
      'adoTaie.Execute "delete from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "' and instr(a1p14,'點作轉專業')>0", intI
      adoTaie.Execute "delete from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "' and (a1p05 like '4%' or a1p05 like '2201%')", intI
      'end 2015/8/21
      'Added by Morgan 2013/12/27
      'Modified by Morgan 2014/2/12 +2405,2631
      'modify by sonia 2020/4/23 因法律所收款同時將智慧所案源收據同時收款故取消1133
      'Modified by Morgnan 2021/1/15 法律所收款改系統自動產生智慧所分錄1133應收帳款加回
      'adoTaie.Execute "delete from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "' and a1p07 > 0 and a1p05 in ('2141','2405','2631')", intI
      adoTaie.Execute "delete from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "' and a1p07 > 0 and a1p05 in ('2141','2405','2631','1133')", intI
      'Added by Morgan 2021/1/18 主要公司非L的L公司現金
      If strA0L05 <> "L" Then
         adoTaie.Execute "delete from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "' and a1p07 > 0 and a1p01='L' and a1p05='1101' and a1p16='L0100'", intI
      End If
      'end 2021/1/18
      'Added by Morgan 2021/3/10
      adoTaie.Execute "delete from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "' and a1p07 > 0 and a1p01='L' and a1p05='6129'", intI
      
      'Added by Morgan 2023/10/19 -保留點數
      adoTaie.Execute "delete from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "' and a1p07 > 0 and a1p05='2492' and a1p23 is not null", intI
      
      'Added by Morgan 2023/8/16 系統自動產生非主要公司的現金及瑞興銀存科目 '110502','110303','110602'
      adoTaie.Execute "delete from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "' and a1p07 > 0 and a1p01<>'" & strA0L05 & "' and a1p05 in ('1101','110502','110303','110602') and a1p17 is null", intI
      
      Adodc1.Recordset.MoveFirst
      Do While Adodc1.Recordset.EOF = False
         
         Acc0m0Save
         
         'Ken 91/11/25 改成點數所屬之智權人員
         If IsNull(Adodc1.Recordset.Fields("t0212").Value) Then
            strManNo = ""
         Else
            strManNo = Adodc1.Recordset.Fields("t0212").Value
         End If
         
         ' 收款金額足時
         If (Adodc1.Recordset.Fields("t0204").Value + Adodc1.Recordset.Fields("t0205").Value) = (Adodc1.Recordset.Fields("t0208").Value + Adodc1.Recordset.Fields("t0209").Value) Then
            'Modified by Morgan 2017/2/16 沒扣繳或全額扣繳才要刪除
            If Val("" & Adodc1.Recordset.Fields("t0206")) = 0 Or Val("" & Adodc1.Recordset.Fields("t0206")) = Val("" & Adodc1.Recordset.Fields("t0223")) Then
               adoTaie.Execute "delete from acc1u0 where a1u01 = '" & Adodc1.Recordset.Fields("t0201").Value & "' and a1u02 = '" & Adodc1.Recordset.Fields("t0202").Value & "'"
            End If
            'end 2017/2/16
            
            bolAll = True
            
            'Removed by Morgan 2014/2/21 移到下面更新(因為部分收款收尾款時也要更新)
            ''Added by Morgan 2013/12/27
            ''更新是否結清(a0k37),介紹獎金發放日期(a0k36)
            'strSql = "update acc0k0 set a0k36=decode(a0k34,null,null," & strSrvDate(2) & "),a0k37='Y' where a0k01='" & Adodc1.Recordset.Fields("t0202").Value & "'"
            'adoTaie.Execute strSql, intI
            ''end 2013/12/27
            'end 2014/2/21
         Else
            bolAll = False
            'Ken 92/12/22 部份收款註記
            adoTaie.Execute "update acc0k0 set a0k13 = 'Y' where a0k01 = '" & Adodc1.Recordset.Fields("t0202").Value & "'"
         End If
         
         m_bolRcptClear = False 'Added by Morgan 2023/9/26
         
         'Modified by Morgan 2014/3/4 從下面移上來
         If (Val(Adodc1.Recordset.Fields("t0204").Value) + Val(Adodc1.Recordset.Fields("t0205").Value) + Val(Adodc1.Recordset.Fields("t0210").Value) + Val(Adodc1.Recordset.Fields("t0211").Value)) = (Val(Adodc1.Recordset.Fields("t0208").Value) + Val(Adodc1.Recordset.Fields("t0209").Value)) - (Val("" & Adodc1.Recordset.Fields("t0219").Value) + Val("" & Adodc1.Recordset.Fields("t0220").Value) + Val("" & Adodc1.Recordset.Fields("t0221").Value) + Val("" & Adodc1.Recordset.Fields("t0222").Value)) Then
            adoTaie.Execute "update acc0k0 set a0k13 = NULL where a0k01 = '" & Adodc1.Recordset.Fields("t0202").Value & "'"
            adoTaie.Execute "update acc1v0 set a1v05 = 'N' where a1v02 = '" & Adodc1.Recordset.Fields("t0202").Value & "'"
            
            m_bolRcptClear = True 'Added by Morgan 2023/9/26
            
            'Modify by Amy 2022/08/29 +Me.Name
            PUB_InvProc Adodc1.Recordset.Fields("t0202").Value, ChangeTDateStringToTString(Frmacc1150.MaskEdBox1), Text2, , Me.Name 'Added by Morgan 2013/12/26 J公司未開發票者建立發票
            'Added by Morgan 2014/2/21
            '更新是否結清(a0k37),介紹獎金發放日期(a0k36)
            'Removed by Morgan 2015/5/7 改在後面用共用函數更新
            'strSql = "update acc0k0 set a0k36=decode(a0k34,null,null," & strSrvDate(2) & "),a0k37='Y' where a0k01='" & Adodc1.Recordset.Fields("t0202").Value & "' and nvl(a0k37,'N')='N'"
            'adoTaie.Execute strSql, intI
            'end 2015/5/7
            'end 2014/2/21
         End If
         'end 2014/3/4
         
         adocheck.CursorLocation = adUseClient
            
         'Modify by Morgan 2011/9/28 改先抓 acc0j0,全額收與部分收語法一樣故合併做一次
         'strSql = "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, a0k04 as Company, nvl(cpm03, nvl(cpm10, cpm13)) as Property, (cu01||cu02) as CustNo, nvl(na03, na04) as Nation, st03, cp13, cp16, cp17, cp73, cp74, cp76, cp09, pa09 as NaNo, a0k30 from caseprogress, salesno, staff, casepropertyMap, patent, customer, nation, acc0k0 where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and substr(pa26, 1, 8) = cu01 (+) and substr(pa26, 9, 1) = cu02 (+) and pa09 = na01 (+) and cp60 = a0k01 and cp60 = '" & Adodc1.Recordset.Fields("t0202").Value & "' union " & _
           "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, a0k04 as Company, nvl(cpm03, nvl(cpm10, cpm13)) as Property, (cu01||cu02) as CustNo, nvl(na03, na04) as Nation, st03, cp13, cp16, cp17, cp73, cp74, cp76, cp09, tm10 as NaNo, a0k30 from caseprogress, salesno, staff, casepropertyMap, trademark, customer, nation, acc0k0 where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and substr(tm23, 1, 8) = cu01 (+) and substr(tm23, 9, 1) = cu02 (+) and tm10 = na01 (+) and cp60 = a0k01 and cp60 = '" & Adodc1.Recordset.Fields("t0202").Value & "' union " & _
           "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, a0k04 as Company, nvl(cpm03, nvl(cpm10, cpm13)) as Property, (cu01||cu02) as CustNo, nvl(na03, na04) as Nation, st03, cp13, cp16, cp17, cp73, cp74, cp76, cp09, lc15 as NaNo, a0k30 from caseprogress, salesno, staff, casepropertyMap, lawcase, customer, nation, acc0k0 where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = lc01 and cp02 = lc02 and cp03 = lc03 and cp04 = lc04 and substr(lc11, 1, 8) = cu01 (+) and substr(lc11, 9, 1) = cu02 (+) and lc15 = na01 (+) and cp60 = a0k01 and cp60 = '" & Adodc1.Recordset.Fields("t0202").Value & "' union " & _
           "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, a0k04 as Company, nvl(cpm03, nvl(cpm10, cpm13)) as Property, (cu01||cu02) as CustNo, nvl(na03, na04) as Nation, st03, cp13, cp16, cp17, cp73, cp74, cp76, cp09, cu10 as NaNo, a0k30 from caseprogress, salesno, staff, casepropertyMap, hirecase, customer, nation, acc0k0 where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = hc01 and cp02 = hc02 and cp03 = hc03 and cp04 = hc04 and substr(hc05, 1, 8) = cu01 (+) and substr(hc05, 9, 1) = cu02 (+) and cu10 = na01 (+) and cp60 = a0k01 and cp60 = '" & Adodc1.Recordset.Fields("t0202").Value & "' union " & _
           "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, a0k04 as Company, nvl(cpm03, nvl(cpm10, cpm13)) as Property, (cu01||cu02) as CustNo, nvl(na03, na04) as Nation, st03, cp13, cp16, cp17, cp73, cp74, cp76, cp09, sp09 as NaNo, a0k30 from caseprogress, salesno, staff, casepropertyMap, servicepractice, customer, nation, acc0k0 where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and substr(sp08, 1, 8) = cu01 (+) and substr(sp08, 9, 1) = cu02 (+) and sp09 = na01 (+) and cp60 = a0k01 and cp60 = '" & Adodc1.Recordset.Fields("t0202").Value & "'"

         'Modify by Morgan 2011/10/4 +,a0j07,a0j09,a0j10
         'modify by sonia 2016/0/18 +a0k34
         'modify by sonia 2017/9/1 cp13 = sn02 (+)改為 a0k20 = sn02 (+)
         strSql = "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, a0k04 as Company, nvl(cpm03, nvl(cpm10, cpm13)) as Property, (cu01||cu02) as CustNo, nvl(na03, na04) as Nation, st03, cp13, cp16, cp17, cp73, cp74, cp76, cp09, pa09 as NaNo, a0j07,a0j09,a0j10,a0k34 from acc0k0, acc0j0, caseprogress, salesno, staff, casepropertyMap, patent, customer, nation where a0k20 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and substr(pa26, 1, 8) = cu01 (+) and substr(pa26, 9, 1) = cu02 (+) and pa09 = na01 (+) and a0k01(+) ='" & Adodc1.Recordset.Fields("t0202").Value & "' and a0j13(+)=a0k01 and cp09(+)=a0j01 union " & _
           "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, a0k04 as Company, nvl(cpm03, nvl(cpm10, cpm13)) as Property, (cu01||cu02) as CustNo, nvl(na03, na04) as Nation, st03, cp13, cp16, cp17, cp73, cp74, cp76, cp09, tm10 as NaNo, a0j07,a0j09,a0j10,a0k34 from acc0k0, acc0j0, caseprogress, salesno, staff, casepropertyMap, trademark, customer, nation where a0k20 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and substr(tm23, 1, 8) = cu01 (+) and substr(tm23, 9, 1) = cu02 (+) and tm10 = na01 (+) and a0k01(+) = '" & Adodc1.Recordset.Fields("t0202").Value & "' and a0j13(+)=a0k01 and cp09(+)=a0j01 union " & _
           "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, a0k04 as Company, nvl(cpm03, nvl(cpm10, cpm13)) as Property, (cu01||cu02) as CustNo, nvl(na03, na04) as Nation, st03, cp13, cp16, cp17, cp73, cp74, cp76, cp09, lc15 as NaNo, a0j07,a0j09,a0j10,a0k34 from acc0k0, acc0j0, caseprogress, salesno, staff, casepropertyMap, lawcase, customer, nation where a0k20 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = lc01 and cp02 = lc02 and cp03 = lc03 and cp04 = lc04 and substr(lc11, 1, 8) = cu01 (+) and substr(lc11, 9, 1) = cu02 (+) and lc15 = na01 (+) and a0k01(+) = '" & Adodc1.Recordset.Fields("t0202").Value & "' and a0j13(+)=a0k01 and cp09(+)=a0j01 union " & _
           "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, a0k04 as Company, nvl(cpm03, nvl(cpm10, cpm13)) as Property, (cu01||cu02) as CustNo, nvl(na03, na04) as Nation, st03, cp13, cp16, cp17, cp73, cp74, cp76, cp09, cu10 as NaNo, a0j07,a0j09,a0j10,a0k34 from acc0k0, acc0j0, caseprogress, salesno, staff, casepropertyMap, hirecase, customer, nation where a0k20 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = hc01 and cp02 = hc02 and cp03 = hc03 and cp04 = hc04 and substr(hc05, 1, 8) = cu01 (+) and substr(hc05, 9, 1) = cu02 (+) and cu10 = na01 (+) and a0k01(+) = '" & Adodc1.Recordset.Fields("t0202").Value & "' and a0j13(+)=a0k01 and cp09(+)=a0j01 union " & _
           "select sn01 as Man, cp01||cp02||cp03||cp04 CaseNo, a0k04 as Company, nvl(cpm03, nvl(cpm10, cpm13)) as Property, (cu01||cu02) as CustNo, nvl(na03, na04) as Nation, st03, cp13, cp16, cp17, cp73, cp74, cp76, cp09, sp09 as NaNo, a0j07,a0j09,a0j10,a0k34 from acc0k0, acc0j0, caseprogress, salesno, staff, casepropertyMap, servicepractice, customer, nation where a0k20 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and substr(sp08, 1, 8) = cu01 (+) and substr(sp08, 9, 1) = cu02 (+) and sp09 = na01 (+) and a0k01(+) = '" & Adodc1.Recordset.Fields("t0202").Value & "' and a0j13(+)=a0k01 and cp09(+)=a0j01"
           
         adocheck.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
         
         If adocheck.RecordCount > 0 Then
            adocheck.MoveLast
            If IsNull(adocheck.Fields("Man").Value) Then
               strMan = ""
            Else
               strMan = adocheck.Fields("Man").Value
            End If
            If IsNull(adocheck.Fields("CustNo").Value) Then
               strCompanyNo = ""
            Else
               strCompanyNo = adocheck.Fields("CustNo").Value
            End If
            If IsNull(adocheck.Fields("Property").Value) Then
               strProperty = ""
            Else
               strProperty = adocheck.Fields("Property").Value
            End If
            If IsNull(adocheck.Fields("CaseNo").Value) Then
               strCaseNo = ""
            Else
               strCaseNo = adocheck.Fields("CaseNo").Value
            End If
            If IsNull(adocheck.Fields("st03").Value) Then
               strDept = ""
            Else
               strDept = adocheck.Fields("st03").Value
            End If
            If IsNull(adocheck.Fields("Nation").Value) Then
               strNation = ""
            Else
               strNation = adocheck.Fields("Nation").Value
            End If
            If IsNull(adocheck.Fields("NaNo").Value) Then
               strNationNo = ""
            Else
               strNationNo = adocheck.Fields("NaNo").Value
            End If
         End If
         
         'add by sonia 2023/6/9 檢查內商人員若案件為MCT案件,則改為部門點數P2005商標部MCT
         If strManNo <> "" And strCaseNo <> "" Then
            If Left(GetST15(strManNo), 2) = "P2" Then
               If Left(PUB_GetAKindSalesNo(Mid(strCaseNo, 1, Len(strCaseNo) - 9), Mid(strCaseNo, Len(strCaseNo) - 8, 6), Mid(strCaseNo, Len(strCaseNo) - 2, 1), Mid(strCaseNo, Len(strCaseNo) - 1, 2)), 4) = "MCTF" Then
                  strManNo = "P2005"
               End If
            End If
         End If
         'end 2023/6/9
            
         '新增 acc1u0
         If bolAll = True Then
            adocheck.MoveFirst
            Do While adocheck.EOF = False
               'Modify by Morgan 2011/10/4 考慮拆收據情形改抓0j0,沒用程式碼一併取消
               douAmount = Val("" & adocheck.Fields("a0j09").Value) + Val("" & adocheck.Fields("a0j10").Value)
               douFAmount = Val("" & adocheck.Fields("a0j10").Value)
               'end 2011/10/4
                  
               adoacc1u0.CursorLocation = adUseClient
               adoacc1u0.Open "select a1u01 from acc1u0 where a1u01 = '" & Adodc1.Recordset.Fields("t0201").Value & "' and a1u02 = '" & Adodc1.Recordset.Fields("t0202").Value & "' and a1u03 = '" & adocheck.Fields("cp09").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
               If adoacc1u0.RecordCount = 0 Then
                  If IsNull(Adodc1.Recordset.Fields("t0206").Value) = False And Adodc1.Recordset.Fields("t0206").Value <> 0 Then
                     'Modified by Morgan 2011/11/24 是否合併改抓 a0j07
                     If adocheck.Fields("a0j07").Value = MsgText(602) Then
                        douTax = douAmount * 0.1
                     Else
                        douTax = (douAmount - douFAmount) * 0.1
                     End If
                  Else
                     douTax = 0
                  End If
                  If douAmount - douFAmount > 0 Or douFAmount Or douTax Then 'Added by Morgan 2014/1/22 有值才寫
                     adoTaie.Execute "insert into acc1u0(a1u01,a1u02,a1u03,a1u04,a1u05,a1u06,a1u07,a1u08,a1u09,a1u10) values ('" & Adodc1.Recordset.Fields("t0201").Value & "', '" & Adodc1.Recordset.Fields("t0202").Value & "', '" & adocheck.Fields("cp09").Value & "', " & douAmount - douFAmount & ", " & douFAmount & ", " & douTax & ", 0, 0, 0, 0)"
                  End If
               End If
               adoacc1u0.Close
               strF5639NO = "" & adocheck.Fields("a0k34").Value   'add by sonia 2016/10/19
               'Added by Lydia 2025/04/18 TIPS分配比例管制：TIPS請款階段分配比例-年度結算
               If Left("" & adocheck.Fields("caseno"), 3) = "ACS" Then
                  bolACS = True 'Added by Morgan 2025/4/22
                  strExc(0) = "" & adocheck.Fields("caseno")
                  Call ChgCaseNo(strExc(0), strExc)
                  If Len(strExc(2)) = 6 Then
                     Call PUB_ProcAcs_Tips_Rate1(True, strExc(1), strExc(2), strExc(3), strExc(4), "" & adocheck.Fields("cp09"), "" & Adodc1.Recordset.Fields("t0202").Value)
                  End If
               End If
               'end 2025/04/18
               adocheck.MoveNext
            Loop
            'add by sonia 2016/10/18 寰華介紹案件
            If strF5639NO = "F5639" Then
               F5639NO = F5639NO & Adodc1.Recordset.Fields("t0202").Value & ";"
            End If
         End If
         adocheck.Close
         
         'Added by Morgan 2013/12/19
         strA1P01 = "1"
         'Modified by Morgan 2020/4/15
         'stra1p22 = strA1P22_1
         intPos = InStr(strA1P01s, strA1P01)
         If intPos > 0 Then
            stra1p22 = "'" & arrA1p22(intPos - 1) & "'"
         Else
            stra1p22 = "null"
         End If
         'end 2020/4/15
         'end 2013/12/19
         
         '收據抬頭
         adocheck.CursorLocation = adUseClient
         'Modified by Morgan 2021/2/1 +a0j02,a0j01,cp01,cp14
         'strSql = "select a0k04,a0k11 from acc0k0 where a0k01 = '" & Adodc1.Recordset.Fields("t0202").Value & "'"
         strSql = "select a0k04,a0k11,a0j02,a0j01,cp01,cp14" & _
            " from acc0k0, acc0j0,caseprogress where a0k01 = '" & Adodc1.Recordset.Fields("t0202").Value & "' and a0j13(+)=a0k01" & _
            " and cp09(+)=a0j01"
         adocheck.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
         If adocheck.RecordCount <> 0 Then
      
            'Added by Morgan 2013/12/19
            'Modified by Morgan 2020/4/15
            'If adocheck.Fields("a0k11") = "J" Then
            '   strA1P01 = "J"
            '   stra1p22 = strA1P22_J
            'End If
            If adocheck.Fields("a0k11") >= "A" Then
               strA1P01 = adocheck.Fields("a0k11")
               intPos = InStr(strA1P01s, strA1P01)
               If intPos > 0 Then
                  stra1p22 = "'" & arrA1p22(intPos - 1) & "'"
               Else
                  stra1p22 = "null"
               End If
            End If
            'end 2020/4/15
            'end 2013/12/19
            
            SetLOSVar Adodc1.Recordset.Fields("t0202").Value, strA1P01, strLOS02, strTTMan, strTTManSN, strLCaseNo, strLRcpTitle, lngTTAmt, bolB2NeeCourt 'Added by Morgan 2021/1/18 案源變數設定
            'Added by Morgan 2021/4/12
            If strLOS02 <> "" Then
               intPos = InStr(strA1P01s, "1")
               If intPos > 0 Then
                  strA1P22_TT = "'" & arrA1p22(intPos - 1) & "'"
               Else
                  strA1P22_TT = "''"
               End If
               'Added by Morgan 2023/11/20
               intPos = InStr(strA1P01s, "L")
               If intPos > 0 Then
                  strA1P22_L = "'" & arrA1p22(intPos - 1) & "'"
               Else
                  strA1P22_L = "''"
               End If
               'end 2023/11/20
            End If
            'end 2021/4/12
            
            'Added by Morgan 2021/2/1 案源未分案提醒
            'Modified by Morgan 2021/2/3 顧問除外--辜
            If strA1P01 = "L" And strLOS02 <> "" Then
               If adocheck.Fields("cp01") <> "LA" And IsNull(adocheck.Fields("cp14")) Then
                  stNoLawyerAlert = stNoLawyerAlert & adocheck.Fields("a0j02") & "(" & adocheck.Fields("a0j01") & ")" & vbCrLf
               End If
            End If
            'end 2021/1/18
            
            
            If IsNull(adocheck.Fields(0).Value) Then
               strCompany = ""
            Else
               '2010/8/17 modify by sonia 加trim其他就不必再寫了
               'strCompany = MidB(adocheck.Fields(0).Value, 1, 16)
               'Modified by Morgan 2015/4/21 取消Trim,否則遇有造字可能會錯
               'strCompany = Trim(MidB(adocheck.Fields(0).Value, 1, 16))
               strCompany = MidB(adocheck.Fields(0).Value, 1, 16)
            End If
         Else
            strCompany = ""
         End If
         adocheck.Close
         
         '新增分錄
         If Val(Adodc1.Recordset.Fields("t0204").Value) <> 0 Or Val(Adodc1.Recordset.Fields("t0205").Value) <> 0 Then
            'Added by Morgan 2013/12/27
            strExc(0) = Val("" & Adodc1.Recordset.Fields("t0204").Value) + Val("" & Adodc1.Recordset.Fields("t0205").Value) - Val("" & Adodc1.Recordset.Fields("t0206").Value)
            strExc(1) = ChgSQL(strMan & "/" & strCompany)
            'Modified by Morgan 2014/3/12 改摘要
            'Modified by Morgan 2023/7/26 公司改為簡稱
            strExc(2) = ChgSQL(A0802Query(strA1P01, True) & "/" & strMan & "/" & strCompany) '2112 '2017/3/15 改用2116
            strExc(3) = ChgSQL(A0802Query(strA0L05, True) & "/" & strMan & "/" & strCompany) '1133
                        
            '若主要公司別不同於收據公司別時
            'Modified by Morgan 2021/1/18
            'If strA1P01 <> strA0L05 Then
            If strA1P01 <> strA0L05 And Not (strA1P01 = "1" And (strLOS02 = "A1" Or strLOS02 = "A2")) Then
            'end 2021/1/18
            
               strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
               
               'Added by Morgan 2023/8/16 現金1101科目改法律所110502,智權110303,智慧所110602--瑞婷
               If strA1P01 = "L" Then
                  stAccNo = "110502"
               ElseIf strA1P01 = "J" Then
                  stAccNo = "110303"
               Else
                  stAccNo = "110602"
               End If
               
               'Added by Morgan 2021/1/18
               '案源收款
               If strA1P01 = "L" And strLOS02 <> "" Then
                  '借 現金1101
                  'Modified by Morgan 2023/8/16
                  'strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p14, a1p15, a1p16, a1p18, a1p22, a1p27)" & _
                     " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '1101', 'TOT', " & strExc(0) & ", 0, '" & ChgSQL("智慧所代收/" & strCompany) & "', '" & strCompanyNo & "', 'L0100', " & Val(FCDate(strDate)) & ", " & stra1p22 & ", " & stra1p27 & ")"
                  strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p14, a1p15, a1p16, a1p18, a1p22, a1p27)" & _
                     " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & stAccNo & "', 'TOT', " & strExc(0) & ", 0, '" & ChgSQL(A0802Query(strA0L05, True) & "代收/" & strCompany) & "', '" & strCompanyNo & "', 'L0100', " & Val(FCDate(strDate)) & ", " & stra1p22 & ", " & stra1p27 & ")"
                  'end 2023/8/16
               Else
               'end 2021/1/18
               
                  '借 應收帳款1133
                  'modify by sonia 2019/7/30 原借1133應收帳款,貸2116應付帳款-台一VS智權改為借1101現金,貸1101現金
                  'Modified by Morgan 2023/8/16 改科目:1101->stAccNo
                  strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p14, a1p15, a1p16, a1p18, a1p22, a1p27)" & _
                     " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & stAccNo & "', 'TOT', " & strExc(0) & ", 0, '" & strExc(3) & "', '" & strCompanyNo & "', '" & strManNo & "', " & Val(FCDate(strDate)) & ", " & stra1p22 & ", " & stra1p27 & ")"
                     
               End If 'Added by Morgan 2021/1/18
               adoTaie.Execute strSql, intI
               
               '貸 應付帳款2112  '2017/3/15 改用2116
               'add by sonia 2020/4/24 依作帳公司重抓原傳票號碼
               intPos = InStr(strA1P01s, strA0L05)
               If intPos > 0 Then
                  strA0L05A1P22 = "'" & arrA1p22(intPos - 1) & "'"
               Else
                  strA0L05A1P22 = "null"
               End If
               'end 2020/4/24
                  
               strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
               
               'Added by Morgan 2023/8/16 現金科目改法律所110502,智權110303,智慧所110602--瑞婷
               If strA0L05 = "L" Then
                  stAccNo = "110502"
               ElseIf strA0L05 = "J" Then
                  stAccNo = "110303"
               Else
                  stAccNo = "110602"
               End If
               
               'Added by Morgan 2021/1/18
               '案源收款
               If strA1P01 = "L" And strLOS02 <> "" Then
                  '貸 現金1101
                  'Modified by Morgan 2023/8/16 改科目:1101->stAccNo
                  strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p14, a1p15, a1p16, a1p18, a1p22, a1p27)" & _
                     " values ('" & strA0L05 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & stAccNo & "', 'TOT', 0, " & strExc(0) & ", '" & ChgSQL("代收法律所/" & strTTManSN & "/" & strCompany) & "', '" & strCompanyNo & "', '" & strTTMan & "', " & Val(FCDate(strDate)) & ", " & strA0L05A1P22 & ", " & stra1p27 & ")"
               Else
               'end 2021/1/18
                  'modify by sonia 2017/3/15 改用2116
                  'modify by sonia 2019/7/30 原借1133應收帳款,貸2116應付帳款-台一VS智權改為借1101現金,貸1101現金
                  'modify by sonia 2020/4/24 stra1p22改strA0L05A1P22主要公司別傳票號
                  'Modified by Morgan 2023/8/16 改科目:1101->stAccNo
                  strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p14, a1p15, a1p16, a1p18, a1p22, a1p27)" & _
                     " values ('" & strA0L05 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & stAccNo & "', 'TOT', 0, " & strExc(0) & ", '" & strExc(2) & "', '" & strCompanyNo & "', '" & strManNo & "', " & Val(FCDate(strDate)) & ", " & strA0L05A1P22 & ", " & stra1p27 & ")"
                  
               End If 'Added by Morgan 2021/1/18
               adoTaie.Execute strSql, intI
               
'cancel by sonia 2017/3/16 移至下面,不管主要公司別與收據公司別是否相同,只要是未收款已沖帳都要做
'               '未收款已沖帳
'               If CheckInvDone(Adodc1.Recordset.Fields("t0202").Value) = True Then
'                  '貸 應收帳款1133
'                  strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
'                  strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p14, a1p15, a1p16, a1p18, a1p22, a1p27)" & _
'                     " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '1133', 'TOT', 0, " & strExc(0) & ", '" & strExc(3) & "', '" & strCompanyNo & "', '" & strManNo & "', " & Val(FCDate(strDate)) & ", " & stra1p22 & ", " & stra1p27 & ")"
'                  adoTaie.Execute strSql, intI
'
'                  '借 應收未收款2141
'                  strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
'                  strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p14, a1p15, a1p16, a1p18, a1p22, a1p27)" & _
'                     " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '2141', 'TOT', " & strExc(0) & ", 0, '" & strExc(1) & "', '" & strCompanyNo & "', '" & strManNo & "', " & Val(FCDate(strDate)) & ", " & stra1p22 & ", " & stra1p27 & ")"
'                  adoTaie.Execute strSql, intI
'               End If
            End If
            'end 2013/12/27
            
            bolACSNoTaxItem = False 'Added by Morgan 2025/8/19
            
            '未收款已沖帳 add by sonia 2017/3/16 不管主要公司別與收據公司別是否相同,只要是未收款已沖帳都要做,公司別都為J公司,摘要加發票號碼
            If CheckInvDone(Adodc1.Recordset.Fields("t0202").Value) = True Then
               
               '抓發票號碼
               stSQL = "select axc01 from acc431 where axc02='" & Adodc1.Recordset.Fields("t0202").Value & "'"
               intR = 1
               Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
               If intR = 1 Then
                  strExc(1) = strExc(1) & "/" & rsQuery(0)
               End If
               '借 應收未收款2141
               strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
               'Modified by Morgan 2025/8/19 ACS案跨期後收款(未收款已沖帳)，規費(稅款)不用產生分錄--瑞婷
               strExc(4) = 0
               If Left(strCaseNo, 3) = "ACS" Then
                  bolACSNoTaxItem = True
                  strExc(4) = Val("" & Adodc1.Recordset.Fields("t0205").Value)
               End If
               strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p14, a1p15, a1p16, a1p17, a1p18, a1p22, a1p27)" & _
                  " values ('J', 'A', '" & strSerialNo & "', '" & Text2 & "', '2141', 'TOT', " & (Val(strExc(0)) - Val(strExc(4))) & ", 0, '" & strExc(1) & "', '" & strCompanyNo & "', '" & strManNo & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", " & stra1p22 & ", " & stra1p27 & ")"
               adoTaie.Execute strSql, intI
               'end 2025/8/19
               
               '貸 應收帳款1133  modify by sonia 2017/4/5 改用1141未入帳應收帳款
               strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
               strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p14, a1p15, a1p16, a1p17, a1p18, a1p22, a1p27)" & _
                  " values ('J', 'A', '" & strSerialNo & "', '" & Text2 & "', '1141', 'TOT', 0, " & strExc(0) & ", '" & strExc(1) & "', '" & strCompanyNo & "', '" & strManNo & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", " & stra1p22 & ", " & stra1p27 & ")"
               adoTaie.Execute strSql, intI
            End If
            
            'Modify by Morgan 2011/4/25 整理並新增收文分配點數收款
            'Acc1p0Save
            Acc1p0SaveNew
            
         End If
         
         'Modify by Morgan 2007/3/28 若有發票號碼且不為'E'字頭時更新收據公司為9
         strExc(1) = ""
         'Removed by Morgan 2022/5/31 改J公司才有發票
         'If "" & Adodc1.Recordset.Fields(2) <> "" And Left("" & Adodc1.Recordset.Fields(2), 1) <> "E" Then
         '   strExc(1) = ",a0k11='9'"
         'End If
         'end 2022/5/31
         'Modify by Amy 2013/06/21存檔時若扣繳年度A0K16與前畫面Frmacc1150之收款日期的年度不同時，同時更新A0K15為系統日
          'Modify by Morgan 2011/8/25
          'adoTaie.Execute "update acc0k0 set a0k16 = " & IIf(IsNull(Adodc1.Recordset.Fields(6).Value), 0, Adodc1.Recordset.Fields(6).Value) & ", a0k17 = " & IIf(IsNull(Adodc1.Recordset.Fields("t0204").Value), 0, Val(Adodc1.Recordset.Fields("t0204").Value) + Val(Adodc1.Recordset.Fields("t0210").Value)) & ", a0k18 = " & IIf(IsNull(Adodc1.Recordset.Fields("t0205").Value), 0, Val(Adodc1.Recordset.Fields("t0205").Value) + Val(Adodc1.Recordset.Fields("t0211").Value)) & strExc(1) & " where a0k01 = '" & Adodc1.Recordset.Fields("t0202").Value & "'"
          '2012/8/23 modify by sonia 同時將收款智權人員更新回收據智權人員
          'adoTaie.Execute "update acc0k0 set a0k16 = " & IIf(IsNull(Adodc1.Recordset.Fields(6).Value), 0, Adodc1.Recordset.Fields(6).Value) & ", (a0k17, a0k18) = (select nvl(sum(a0m04), 0),nvl(sum(a0m05), 0) from acc0m0 where a0m02 = a0k01)" & strExc(1) & " where a0k01 = '" & Adodc1.Recordset.Fields("t0202").Value & "'"
          'adoTaie.Execute "update acc0k0 set a0k16 = " & IIf(IsNull(Adodc1.Recordset.Fields(6).Value), 0, Adodc1.Recordset.Fields(6).Value) & ", (a0k17, a0k18) = (select nvl(sum(a0m04), 0),nvl(sum(a0m05), 0) from acc0m0 where a0m02 = a0k01), a0k20='" & Adodc1.Recordset.Fields("t0212").Value & "'" & strExc(1) & " where a0k01 = '" & Adodc1.Recordset.Fields("t0202").Value & "'"
          'end 2007/3/28
          If Len(FCDate(Frmacc1150.MaskEdBox1)) = 7 Then '取Frmacc1150.收款日期之年度
            acc1150Y = Val(Left(FCDate(Frmacc1150.MaskEdBox1), 3))
          Else
            acc1150Y = Val(Left(FCDate(Frmacc1150.MaskEdBox1), 2))
          End If
         
          If acc1150Y = Val(DataGrid1.Columns(6).Value) Then
            'modify by sonia 2020/6/4 L公司收據不改A0K20智權人員,因為繳款人員是案源介紹人
            'adoTaie.Execute "update acc0k0 set a0k16 = " & IIf(IsNull(Adodc1.Recordset.Fields(6).Value), 0, Adodc1.Recordset.Fields(6).Value) & ", (a0k17, a0k18) = (select nvl(sum(a0m04), 0),nvl(sum(a0m05), 0) from acc0m0 where a0m02 = a0k01), a0k20='" & Adodc1.Recordset.Fields("t0212").Value & "'" & strExc(1) & " where a0k01 = '" & Adodc1.Recordset.Fields("t0202").Value & "'"
            adoTaie.Execute "update acc0k0 set a0k16 = " & IIf(IsNull(Adodc1.Recordset.Fields(6).Value), 0, Adodc1.Recordset.Fields(6).Value) & ", (a0k17, a0k18) = (select nvl(sum(a0m04), 0),nvl(sum(a0m05), 0) from acc0m0 where a0m02 = a0k01), a0k20=decode(a0k11,'L',a0k20,'" & Adodc1.Recordset.Fields("t0212").Value & "')" & strExc(1) & " where a0k01 = '" & Adodc1.Recordset.Fields("t0202").Value & "'"
          Else
            'modify by sonia 2020/6/4 L公司收據不改A0K20智權人員,因為繳款人員是案源介紹人
            'adoTaie.Execute "update acc0k0 set a0k16 = " & IIf(IsNull(Adodc1.Recordset.Fields(6).Value), 0, Adodc1.Recordset.Fields(6).Value) & ", (a0k17, a0k18) = (select nvl(sum(a0m04), 0),nvl(sum(a0m05), 0) from acc0m0 where a0m02 = a0k01), a0k20='" & Adodc1.Recordset.Fields("t0212").Value & "'" & strExc(1) & ",a0k15='" & strSrvDate(2) & "' where a0k01 = '" & Adodc1.Recordset.Fields("t0202").Value & "'"
            adoTaie.Execute "update acc0k0 set a0k16 = " & IIf(IsNull(Adodc1.Recordset.Fields(6).Value), 0, Adodc1.Recordset.Fields(6).Value) & ", (a0k17, a0k18) = (select nvl(sum(a0m04), 0),nvl(sum(a0m05), 0) from acc0m0 where a0m02 = a0k01), a0k20=decode(a0k11,'L',a0k20,'" & Adodc1.Recordset.Fields("t0212").Value & "')" & strExc(1) & ",a0k15='" & strSrvDate(2) & "' where a0k01 = '" & Adodc1.Recordset.Fields("t0202").Value & "'"
          End If
         'end 2013/06/21
         
         'Add by Morgan 2006/9/6 若發票號為手開收據時回寫收據檔手開收據的扣繳年度
         If Left("" & Adodc1.Recordset.Fields(2), 1) = "E" Then
            adoTaie.Execute "update acc0k0 set a0k16 = " & IIf(IsNull(Adodc1.Recordset.Fields(6).Value), 0, Adodc1.Recordset.Fields(6).Value) & " where a0k01 = '" & Adodc1.Recordset.Fields(2) & "'"
         End If
         
         'Added by Morgan 2023/11/1
         '智權公司收據收款時，若未列印則更新為不列印
         If Adodc1.Recordset.Fields("t0214").Value = "J" Then
            'Modified by Morgan 2023/11/14 N->Z(確定不印)
            adoTaie.Execute "update acc0k0 set a0k32 = 'Z' where a0k01 = '" & Adodc1.Recordset.Fields("t0202") & "' and nvl(a0k19,0)=0 and nvl(a0k32,'Y')<>'Z'", intI
         End If
         'end 2023/11/1
         
         PUB_UpdateReceiptStatus Adodc1.Recordset.Fields("t0202") 'Added by Morgan 2015/5/7
         Adodc1.Recordset.MoveNext
      Loop
      
      adocheck.CursorLocation = adUseClient
      'Modified by Morgan 2013/12/19 收款會有J公司,取消 a1p01='1' 條件
      adocheck.Open "select sum(a1p08) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
         If IsNull(adocheck.Fields(0).Value) = False Then
            If Val(Text3) > 0 Then
               'Added by Morgan 2013/12/27 暫收固定抓主要公司別
               'Modified by Morgan 2015/8/21 其他對沖也要放暫收款單號
               'modify by sonia 2020/4/24 stra1p22改strA0L05A1P22主要公司別傳票號
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27, a1p28, a1p29, a1p30) " & _
               "values ('" & strA0L05 & "', 'A', '" & GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3) & "', '" & Text2 & "', '2401', '" & MsgText(55) & "', 0, " & Val(Text3) & ", null, null, null, null, null, '" & strMan & "/" & strCompany & "/" & Text4 & "', '" & strCompanyNo & "', '" & strManNo & "', null, " & Val(FCDate(strDate)) & ", null, null, null, " & IIf(strA0L05A1P22 = "", "null", strA0L05A1P22) & ", '" & Text4 & "', null, null, null, " & stra1p27 & ", null, null, '" & Text4 & "')"
            End If
         End If
      End If
      adocheck.Close
      
      'Added by Morgan 2023/5/22 若出庭費220113後面有代收款項2407xx，則統一改到最後
      strSql = "update acc1p0 a set a1p03=(select lpad(mxsn+sn,3,'0')" & _
         " from (select a1p03 odsn,rownum sn from acc1p0 where a1p04=a.a1p04 and a1p05=a.a1p05 order by a1p03) x" & _
         ",(select max(a1p03) mxsn from acc1p0 b where  a1p04=a.a1p04 and a1p05<>a.a1p05) y where odsn=a.a1p03)" & _
         " where a1p04='" & Text2 & "' and a1p01='L' and a1p05='220113'" & _
         " and exists(select * From acc1p0 b where a1p01=a.a1p01 and a1p04=a.a1p04 and a1p05 like '2407%' and a1p03>a.a1p03)"
      adoTaie.Execute strSql, intI
      'end 2023/5/22
      
      '更新票據資料
      adocheck.CursorLocation = adUseClient
      'Modified by Morgan 2013/12/19 收款會有J公司,取消 a1p01='1' 條件
      'Modify by Amy 2020/06/30 +a1p11 a0e07因改為key
      adocheck.Open "select a1p09, a1p10, a1p11 from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "' and a1p09 is not null", adoTaie, adOpenStatic, adLockReadOnly
      Do While adocheck.EOF = False
         If strCompanyNo = "" Then
            If Adodc1.Recordset.State = adStateOpen Then
               Adodc1.Recordset.MoveFirst
               adoaccsum.CursorLocation = adUseClient
               adoaccsum.Open "select a0k03 from acc0k0 where a0k01 = '" & Adodc1.Recordset.Fields("t0202").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
               If adoaccsum.RecordCount <> 0 Then
                  If IsNull(adoaccsum.Fields("a0k03").Value) = False Then
                     strCompanyNo = adoaccsum.Fields("a0k03").Value
                  End If
               End If
               adoaccsum.Close
            End If
         End If
         'Modify by Amy 2020/06/30 +a1p11
         adoTaie.Execute "update acc0e0 set a0e05 = '1', a0e06 = '" & strCompanyNo & "' where a0e01 = '" & adocheck.Fields("a1p10").Value & "' and a0e02 = '" & adocheck.Fields("a1p09").Value & "' And a0e07='" & adocheck.Fields("a1p11") & "' "
         adocheck.MoveNext
      Loop
      adocheck.Close
      
   ' 計算並儲存已收金額
      If Adodc1.Recordset.RecordCount <> 0 Then
         Adodc1.Recordset.MoveFirst
      End If
      Do While Adodc1.Recordset.EOF = False
         
         If Adodc1.Recordset.Fields("t0214").Value <> "J" Then 'Added by Morgan 2013/12/27 智權公司不必寫 acc1v0
         
         adocheck.CursorLocation = adUseClient
         'Modify by Morgan 2011/10/4 +a1u02
         'adocheck.Open "select a1u03, sum(a1u04), sum(a1u05), sum(a1u06), sum(a1u07+a1u09), sum(a1u08+a1u10), sum(a1u04+a1u05-a1u08-a1u10+a1u07+a1u09),sum(a1u07), sum(a1u09) from acc1u0 where a1u02 = '" & Adodc1.Recordset.Fields("t0202").Value & "' group by a1u03", adoTaie, adOpenStatic, adLockReadOnly
         adocheck.Open "select a1u03, sum(a1u04), sum(a1u05), sum(a1u06), sum(a1u07+a1u09), sum(a1u08+a1u10), sum(a1u04+a1u05-a1u08-a1u10+a1u07+a1u09),sum(a1u07), sum(a1u09),a1u02 from acc1u0 where a1u02 = '" & Adodc1.Recordset.Fields("t0202").Value & "' group by a1u03,a1u02", adoTaie, adOpenStatic, adLockReadOnly
         Do While adocheck.EOF = False
            If IsNull(adocheck.Fields(1).Value) Then
               douTAmount(0) = 0
            Else
               douTAmount(0) = adocheck.Fields(1).Value
            End If
            If IsNull(adocheck.Fields(2).Value) Then
               douTAmount(1) = 0
            Else
               douTAmount(1) = adocheck.Fields(2).Value
            End If
            If IsNull(adocheck.Fields(3).Value) Then
               douTAmount(2) = 0
            Else
               douTAmount(2) = adocheck.Fields(3).Value
            End If
            If IsNull(adocheck.Fields(4).Value) Then
               douTAmount(3) = 0
            Else
               douTAmount(3) = adocheck.Fields(4).Value
            End If
            If IsNull(adocheck.Fields(5).Value) Then
               douTAmount(4) = 0
            Else
               douTAmount(4) = adocheck.Fields(5).Value
            End If
            If IsNull(adocheck.Fields(6).Value) Then
               douTAmount(5) = 0
            Else
               douTAmount(5) = adocheck.Fields(6).Value
            End If
            '93.12.28 add by sonia
            If IsNull(adocheck.Fields(7).Value) Then
               douTAmount(6) = 0
            Else
               douTAmount(6) = adocheck.Fields(7).Value
            End If
            If IsNull(adocheck.Fields(8).Value) Then
               douTAmount(7) = 0
            Else
               douTAmount(7) = adocheck.Fields(8).Value
            End If
            '93.12.28 end
            'Ken 92/11/20 新增扣繳明細資料
            Dim douFullTax As Double
            Dim douPayTax As Double
            Dim douNonePayTax As Double
            Dim strPayMethod As String
            Dim strYear As String
            Dim strCPN As String
            Dim strANN As String
            
            '93.1.13 MODIFY BY SONIA
            '不管收款時有無扣繳都要新增扣繳明細 ACC1V0
            'If adocheck.Fields(3).Value <> 0 Then
            '93.1.13 END
               adoaccsum.CursorLocation = adUseClient
               'Modified by Morgan 2011/11/10 考慮拆收據情形,條件要含收據號
               'Modified by Morgan 2011/12/27 取消 a0j20
               'adoaccsum.Open "select * from acc0j0, acc0k0, acc0m0 where a0j13 = a0k01 and a0j13 = a0m02 (+) and a0j01 = '" & adocheck.Fields("a1u03").Value & "' and a0j13 = '" & adocheck.Fields("a1u02").Value & "' and a0m01 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
               adoaccsum.Open "select a.*,b.*,c.*,getcp10desc(cp01,cp10,a0j04) cp10N,na03 from acc0j0 a, acc0k0 b, acc0m0 c,caseprogress,nation where a0j13 = a0k01 and a0j13 = a0m02 (+) and a0j01 = '" & adocheck.Fields("a1u03").Value & "' and a0j13 = '" & adocheck.Fields("a1u02").Value & "' and a0m01 = '" & Text2 & "' and cp09(+)=a0j01 and na01(+)=a0j04", adoTaie, adOpenStatic, adLockReadOnly
               If adoaccsum.RecordCount <> 0 Then
                   '93.12.28 待改扣繳資料應扣除銷帳部份
                   'If adoaccsum.Fields("a0j07").Value = MsgText(602) Then
                   '    douFullTax = (Val(adoaccsum.Fields("a0j09").Value) + Val(adoaccsum.Fields("a0j10").Value)) * 0.1
                   'Else
                   '    douFullTax = Val(adoaccsum.Fields("a0j09").Value) * 0.1
                   'End If
                   If adoaccsum.Fields("a0j07").Value = MsgText(602) Then
                      douFullTax = Val(adoaccsum.Fields("a0j09").Value) + Val(adoaccsum.Fields("a0j10").Value) - douTAmount(6) - douTAmount(7)
                   Else
                      douFullTax = Val(adoaccsum.Fields("a0j09").Value) - douTAmount(6)
                   End If
                   douFullTax = douFullTax * 0.1
                   '93.12.28 END
                   douPayTax = douTAmount(2)
                   douNonePayTax = douFullTax - douPayTax
                   If IsNull(adoaccsum.Fields("a0k13").Value) Then
                       strPayMethod = "'N'"
                   Else
                       strPayMethod = "'" & adoaccsum.Fields("a0k13").Value & "'"
                   End If
                   If IsNull(adoaccsum.Fields("a0k16").Value) Then
                       strYear = "null"
                   Else
                       strYear = adoaccsum.Fields("a0k16").Value
                   End If
                   
                   'Modified by Morgan 2011/12/27 取消 a0j20
                   If IsNull(adoaccsum.Fields("cp10N").Value) Then
                       strCPN = "null"
                   Else
                       strCPN = "'" & adoaccsum.Fields("cp10N").Value & "'"
                   End If
                   
                   'Modified by Morgan 2011/12/29 取消 a0j21
                   If IsNull(adoaccsum.Fields("na03").Value) Then
                       strANN = "null"
                   Else
                       strANN = "'" & adoaccsum.Fields("na03").Value & "'"
                   End If
                   
                   If IsNull(adoaccsum.Fields("a0m03").Value) Then
                      strInvoiceNo = "null"
                   Else
                      strInvoiceNo = "'" & adoaccsum.Fields("a0m03").Value & "'"
                   End If
                   If adoquery.State = adStateOpen Then
                      adoquery.Close
                   End If
                   'adoTaie.Execute "delete from acc1v0 where a1v01 = '" & adoaccsum.Fields("a0j01").Value & "'"
                   '92.1.13 MODIFY BY SONIA
                   'adoTaie.Execute "insert into acc1v0 (a1v01, a1v02, a1v03, a1v04, a1v05, a1v06, a1v07, a1v09, a1v12, a1v13, a1v18) values ('" & adoaccsum.Fields("a0j01").Value & "', '" & adoaccsum.Fields("a0j13").Value & "', '" & adoaccsum.Fields("a0k11").Value & "', " & douFullTax & ", " & strPayMethod & ", " & douPayTax & ", " & douNonePayTax & ", " & strYear & ", " & strCPN & ", " & strANN & ", '1')"
                   adoquery.CursorLocation = adUseClient
                   'Modify by Morgan 2011/10/4 +a1v02
                   adoquery.Open "select * from acc1v0 where a1v01 = '" & adoaccsum.Fields("a0j01").Value & "' and a1v02='" & adoaccsum.Fields("a0j13").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
                   If adoquery.RecordCount <> 0 Then
                        'Modify By Cheng 2004/05/03
                        '更新扣繳明細的扣繳年度(A1V09)
'                      adoTaie.Execute "update acc1v0 set a1v04 = " & douFullTax & ", a1v05 = " & strPayMethod & ", a1v06 = " & douPayTax & ", a1v07 = " & douNonePayTax & " where a1v01 = '" & adoaccsum.Fields("a0j01").Value & "'"
                        'Modify by Morgan 2007/1/2 加發票號
                        'Modify by Morgan 2008/1/15 若有扣繳時a1v18要設'1'
                        'adoTaie.Execute "update acc1v0 set a1v04 = " & douFullTax & ", a1v05 = " & strPayMethod & ", a1v06 = " & douPayTax & ", a1v07 = " & douNonePayTax & ", A1V09=" & strYear & ",A1V17=" & strInvoiceNo & " where a1v01 = '" & adoaccsum.Fields("a0j01").Value & "'"
                        If adocheck.Fields(3).Value <> 0 Then
                           'Modify by Morgan 2011/10/4 +a1v02
                           adoTaie.Execute "update acc1v0 set a1v04 = " & douFullTax & ", a1v05 = " & strPayMethod & ", a1v06 = " & douPayTax & ", a1v07 = " & douNonePayTax & ", A1V09=" & strYear & ",A1V17=" & strInvoiceNo & ",A1V18='1' where a1v01 = '" & adoaccsum.Fields("a0j01").Value & "' and a1v02='" & adoaccsum.Fields("a0j13").Value & "'"
                        Else
                           'Modify by Morgan 2011/10/4 +a1v02
                           adoTaie.Execute "update acc1v0 set a1v04 = " & douFullTax & ", a1v05 = " & strPayMethod & ", a1v06 = " & douPayTax & ", a1v07 = " & douNonePayTax & ", A1V09=" & strYear & ",A1V17=" & strInvoiceNo & " where a1v01 = '" & adoaccsum.Fields("a0j01").Value & "' and a1v02='" & adoaccsum.Fields("a0j13").Value & "'"
                        End If
                        'End
                   Else
                      If adocheck.Fields(3).Value <> 0 Then
                         adoTaie.Execute "insert into acc1v0 (a1v01, a1v02, a1v03, a1v04, a1v05, a1v06, a1v07, a1v09, a1v12, a1v13, a1v18, a1v17) values ('" & adoaccsum.Fields("a0j01").Value & "', '" & adoaccsum.Fields("a0j13").Value & "', '" & adoaccsum.Fields("a0k11").Value & "', " & douFullTax & ", " & strPayMethod & ", " & douPayTax & ", " & douNonePayTax & ", " & strYear & ", " & strCPN & ", " & strANN & ", '1', " & strInvoiceNo & ")"
                      Else
                         adoTaie.Execute "insert into acc1v0 (a1v01, a1v02, a1v03, a1v04, a1v05, a1v06, a1v07, a1v09, a1v12, a1v13, a1v17) values ('" & adoaccsum.Fields("a0j01").Value & "', '" & adoaccsum.Fields("a0j13").Value & "', '" & adoaccsum.Fields("a0k11").Value & "', " & douFullTax & ", " & strPayMethod & ", " & douPayTax & ", " & douNonePayTax & ", " & strYear & ", " & strCPN & ", " & strANN & ", " & strInvoiceNo & ")"
                      End If
                   End If
                   adoquery.Close
                   '92.1.13 END
               End If
               adoaccsum.Close
            '93.1.13 MODIFY BY SONIA
            'End If
            '93.1.13 END
            adocheck.MoveNext
         Loop
         adocheck.Close
         
         End If 'Added by Morgan 2013/12/27
         
         Adodc1.Recordset.MoveNext
      Loop
      'Add by Morgan 2011/10/4 更新 CP 財務相關欄位
      strSql = "update caseprogress set (cp73, cp74, cp76, cp77, cp78)=(select nvl(sum(a1u04),0) c73,nvl(sum(a1u05),0) c74" & _
         ",nvl(sum(a1u06),0) c76,nvl(sum(a1u07),0)+nvl(sum(a1u09),0) c77,nvl(sum(a1u08),0)+nvl(sum(a1u10),0) c78 from acc1u0" & _
         " where a1u03=cp09) " & _
         " where cp09 in (select a0j01 from acc0m0,acc0j0 where a0m01='" & Text2 & "' and a0j13(+)=a0m02)"
      adoTaie.Execute strSql, intI
      
      strSql = "update caseprogress set cp75=cp73+cp74,cp79=cp16-cp73-cp74+cp78-cp77" & _
         " where cp09 in (select a0j01 from acc0m0,acc0j0 where a0m01='" & Text2 & "' and a0j13(+)=a0m02)"
      adoTaie.Execute strSql, intI
      'end 2011/10/4
      
      If Text4 <> MsgText(601) Then
         'Modify by Morgan 2005/9/16 若溢收金額=0時刪除暫收資料
         'adoTaie.Execute "update acc0t0 set a0t08 = " & Val(Text3) & " where a0t01 = '" & Text4 & "'"
         If Val(Text3) = 0 Then
            adoTaie.Execute "delete acc0t0 where a0t01 = '" & Text4 & "'"
         Else
            adoTaie.Execute "update acc0t0 set a0t08 = " & Val(Text3) & " where a0t01 = '" & Text4 & "'"
         End If
      End If
      adoTaie.Execute "delete from acctmp02 where t0218 = '" & strUserNum & "'"
   End If
   adoTaie.Execute "update acc0l0 set a0l03 = (select sum(a0m15) from acc0m0 where a0m01 = '" & Text2 & "'), a0l04 = (select sum(a0m16) from acc0m0 where a0m01 = '" & Text2 & "'), a0l10 = (select sum(a0m06) from acc0m0 where a0m01 = '" & Text2 & "'), a0l08 = (select sum(a0m04) from acc0m0 where a0m01 = '" & Text2 & "'), a0l09 = (select sum(a0m05) from acc0m0 where a0m01 = '" & Text2 & "') where a0l01 = '" & Text2 & "'"
   
   adoTaie.Execute "delete from acc1v0 where a1v06=0 and a1v07=0" 'Added by Morgan 2016/3/21
   
   BatchCheck 'Added by Morgan 2013/12/27
   SetProofPrint 'Added by Morgan 2014/1/13
   
   If strTrans = MsgText(602) Then
      adoTaie.CommitTrans
   End If
   PUB_SendMailCache 'Added by Lydia 2025/04/18
   
   'Removed by Morgan 2014/8/18 辜說狀況很多改不通知
   'DetailChangeInform 'Added by Morgan 2013/12/27
   
   strCon1 = "Y"
   StatusClear
   strCustNo = MsgText(601)
   tool1_enabled
   
   'Add by Morgan 2008/10/17
   If strItemNo = MsgText(601) Then
      MsgBox "作業可能有錯，請通知電腦中心！"
      strItemNo = Frmacc1150.Text2
   End If
   
   If strTrans = MsgText(602) Then
      'Modify by Morgan 2011/4/26
      '將出庭費及分配點數訊息併入
      'If strSupportCaseList <> "" Then
      '   'Modify by Morgan 2010/6/2 改訊息
      '   'MsgBox "本次收款收據的收文有支援紀錄，將依收文各轉業務點數5點至專業點數！"
      '   MsgBox "本次收款收據的收文有支援紀錄，將依收文各轉業務點數5點至專業點數！" & vbCrLf & vbCrLf & "案號清單：" & vbCrLf & strSupportCaseList
      'End If
      'Modified by Morgan 2021/1/20 +strProFeeCaseList
      If strSupportCaseList & strXFeeCaseList & strAssignCaseList & strProFeeCaseList <> "" Then
         strMsg = "下列為支援、出庭費及分配點數之案件，若遇【部分收款】" & vbTab & vbCrLf & vbCrLf & _
                  "或【拆收據】致無法扣除或分配時將標示於案號後，請自行" & vbTab & vbCrLf & vbCrLf & _
                  "調整科目並依收款的服務費應先扣支援點數(出庭費)為原則" & vbTab & vbCrLf & vbCrLf & _
                  "，【收多少扣多少】。" & vbCrLf & vbCrLf
            
         If strSupportCaseList <> "" Then
            strMsg = strMsg & vbCrLf & "支援：" & vbCrLf & strSupportCaseList & vbCrLf
         End If
         
         If strXFeeCaseList <> "" Then
            strMsg = strMsg & vbCrLf & "出庭費：" & vbCrLf & strXFeeCaseList & vbCrLf
         End If
         
         If strAssignCaseList <> "" Then
            strMsg = strMsg & vbCrLf & "分配點數：" & vbCrLf & strAssignCaseList & vbCrLf
         End If
         'Added by Morgan 2021/1/20
         If strProFeeCaseList <> "" Then
            strMsg = strMsg & vbCrLf & "專業點數：" & vbCrLf & strProFeeCaseList & vbCrLf
         End If
         'end 2021/1/20
         MsgBox strMsg, vbExclamation
      End If
      
      'Added by Morgan 2021/2/1
      If stNoLawyerAlert <> "" Then
         MsgBox "本次法律案收款有下列收文尚未分案，請自行調整分錄內容。(法務收入、出庭律師...)" & vbCrLf & vbCrLf & stNoLawyerAlert, vbExclamation, "法律案未分案提醒"
      End If
      
      'Added by Morgan 2025/4/22
      If bolACS Then
         strExc(1) = Pub_GetSpecMan("TIPS分配比例不適用案件")
         strExc(0) = "select distinct a0j02 from acc1u0,acc0j0 where a1u01='" & Text2 & "' and a0j01(+)=a1u03 and a0j13(+)=a1u02 and instr('" & strExc(1) & "',a0j02)>0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            RsTemp.MoveFirst
            MsgBox "本次收款含下列ACS案號不適用於TIPS分配新規則，請自行調整分錄！" & vbCrLf & vbCrLf & RsTemp.GetString, vbExclamation
         End If
      End If
      'end 2025/4/22
   End If
   
   DeliverInform 'Add by Morgan 2006/7/19 全額收款發Mail
   PUB_SendMailCache 'Added by Morgan 2013/12/27
   
   'Memo 從上面移下來(因為若有彈訊息會不觸發 Activate 事件)
   Frmacc1150.Show
   Frmacc1150.Adodc1.Recordset.Requery 'Add by Morgan 2010/6/2
   
   Set Frmacc1151 = Nothing
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   If strTrans = MsgText(602) Then
      adoTaie.RollbackTrans
   End If
   adoTaie.Execute "delete from acctmp02 where t0218 = '" & strUserNum & "'"
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub DeliverInform()
'add by sonia 2016/10/18
Dim arrNo, i As Integer
Dim stSQL As String, stSubject As String, stContent As String, intQ As Integer
'end 201/10/18

'Modified by Morgan 2016/8/22 因出納繳款確認也有通知,改共用
   PUB_AccDeliverInform "1", Text2

   'add by sonia 2016/10/18 寰華介紹案件發mail至林柳岑99005之外部信箱st18及Pub_GetSpecMan("財務處總帳人員") (E10524576)
   If F5639NO <> "" Then
      arrNo = Split(F5639NO, ";")
      For i = LBound(arrNo) To UBound(arrNo)
         If arrNo(i) <> "" Then
            strSql = "select cp01||'-'||cp02||'-'||cp03||'-'||cp04 本所案號,nvl(nvl(pa05,pa06),pa07) 案件名稱,na03 申請國家,decode(pa09,'000',cpm03,cpm04) 案件性質,substr(a0j01,10) 佣金,s2.st02 智權人員,nvl(nvl(cu04,cu05),cu06) 申請人,s1.st18 收件人,cp09 from caseprogress,casepropertymap,nation,patent,staff s1,staff s2,customer, (select min(a0j01||a0k17) a0j01 from acc0k0,acc0j0 where a0k01='" & arrNo(i) & "' and a0k01=a0j13(+))" & _
                     " where substr(a0j01,1,9)=cp09(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa01 is not null and pa09=na01(+) and '99005'=s1.st01(+) and cp13=s2.st01(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) " & _
              "union select cp01||'-'||cp02||'-'||cp03||'-'||cp04 本所案號,tm05 案件名稱,na03 申請國家,decode(tm10,'000',cpm03,cpm04) 案件性質,substr(a0j01,10) 佣金,s2.st02 智權人員,nvl(nvl(cu04,cu05),cu06) 申請人,s1.st18 收件人,cp09 from caseprogress,casepropertymap,nation,trademark,staff s1,staff s2,customer, (select min(a0j01||a0k17) a0j01 from acc0k0,acc0j0 where a0k01='" & arrNo(i) & "' and a0k01=a0j13(+))" & _
                     " where substr(a0j01,1,9)=cp09(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and tm01 is not null and tm10=na01(+) and '99005'=s1.st01(+) and cp13=s2.st01(+) and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) " & _
              "union select cp01||'-'||cp02||'-'||cp03||'-'||cp04 本所案號,nvl(nvl(lc05,lc06),lc07) 案件名稱,na03 申請國家,decode(lc15,'000',cpm03,cpm04) 案件性質,substr(a0j01,10) 佣金,s2.st02 智權人員,nvl(nvl(cu04,cu05),cu06) 申請人,s1.st18 收件人,cp09 from caseprogress,casepropertymap,nation,lawcase,staff s1,staff s2,customer, (select min(a0j01||a0k17) a0j01 from acc0k0,acc0j0 where a0k01='" & arrNo(i) & "' and a0k01=a0j13(+))" & _
                     " where substr(a0j01,1,9)=cp09(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and lc01 is not null and lc15=na01(+) and '99005'=s1.st01(+) and cp13=s2.st01(+) and substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+) " & _
              "union select cp01||'-'||cp02||'-'||cp03||'-'||cp04 本所案號,nvl(nvl(sp05,sp06),sp07) 案件名稱,na03 申請國家,decode(sp09,'000',cpm03,cpm04) 案件性質,substr(a0j01,10) 佣金,s2.st02 智權人員,nvl(nvl(cu04,cu05),cu06) 申請人,s1.st18 收件人,cp09 from caseprogress,casepropertymap,nation,servicepractice,staff s1,staff s2,customer, (select min(a0j01||a0k17) a0j01 from acc0k0,acc0j0 where a0k01='" & arrNo(i) & "' and a0k01=a0j13(+))" & _
                     " where substr(a0j01,1,9)=cp09(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and sp01 is not null and sp09=na01(+) and '99005'=s1.st01(+) and cp13=s2.st01(+) and substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+) " & _
              "union select cp01||'-'||cp02||'-'||cp03||'-'||cp04 本所案號,hc06 案件名稱,na03 申請國家,cpm03 案件性質,substr(a0j01,10) 佣金,s2.st02 智權人員,nvl(nvl(cu04,cu05),cu06) 申請人,s1.st18 收件人,cp09 from caseprogress,casepropertymap,nation,hirecase,staff s1,staff s2,customer, (select min(a0j01||a0k17) a0j01 from acc0k0,acc0j0 where a0k01='" & arrNo(i) & "' and a0k01=a0j13(+))" & _
                     " where substr(a0j01,1,9)=cp09(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and hc01 is not null and '000'=na01(+) and '99005'=s1.st01(+) and cp13=s2.st01(+) and substr(hc05,1,8)=cu01(+) and substr(hc05,9,1)=cu02(+)"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               With RsTemp
               If CheckAccDeliverF5639("" & .Fields("收件人"), "" & .Fields("cp09")) = False Then
                  stSubject = .Fields("本所案號") & "【" & .Fields("申請國家") & "】之" & .Fields("案件性質") & "【" & .Fields("cp09") & "】寰華介紹案件已收款通知！"
                  stContent = "本所案號：" & .Fields("本所案號") & vbCrLf & _
                     "案件名稱：" & .Fields("案件名稱") & vbCrLf & _
                     "申請國家：" & .Fields("申請國家") & vbCrLf & _
                     "申請人　：" & .Fields("申請人") & vbCrLf & _
                     "智權人員：" & .Fields("智權人員") & vbCrLf & _
                     "案件性質：" & .Fields("案件性質") & vbCrLf & _
                     "佣　　金：" & Format(Round(Val("" & .Fields("佣金")) / 10, 0), FDollar)
               
                  bolMailSendOk = False
                  PUB_SendMail strUserNum, "" & .Fields("收件人"), "" & .Fields("cp09"), stSubject, stContent, , , , , , Pub_GetSpecMan("財務處總帳人員"), , , , , False
                  If bolMailSendOk = True Then
                     stSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc05,mc06,mc07,mc08)" & _
                        " values( '" & strUserNum & "','" & "" & .Fields("收件人") & "',to_char(sysdate,'yyyymmdd')" & _
                        ",to_char(sysdate,'hh24miss'),to_char(sysdate,'YYYYMMDD'),to_char(sysdate,'HH24MISS'),'" & ChgSQL(stSubject) & "','" & ChgSQL(stContent) & "')"
                     cnnConnection.Execute stSQL, intQ
                  End If
               End If
               End With
            End If
         End If
      Next
   End If
   'end 2016/10/18
   
'   Dim stLstCP09 As String, stSubject As String, stContent As String
'   Dim adoRst As ADODB.Recordset
'
'On Error GoTo ErrHnd
'   'CFP抓無法發文紀錄最後的記錄人員,CFT抓承辦人(無承辦則給陳經理)
'   'Modify by Morgan 2006/8/15 加案件性質
'   'Modify by Morgan 2006/11/1 加申請國家
'   'Modify by Morgan 2011/10/13 考慮拆收據情形
'   'strExc(0) = "select cp01,cp02,cp03,cp04,cp05,pa05,cp09,ud02,ud04,cp10,decode(pa09,'020',cpm04,cpm03) cp10n, na03" & _
'      " From acc0m0, caseprogress, undeliveredrec, patent, casepropertymap, nation" & _
'      " where a0m01='" & Text2 & "' and cp60(+)=a0m02 and cp01='CFP' and cp79=0 and cp27 is null and cp57 is null" & _
'      " and ud01=cp09 and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cpm01(+)=cp01 and cpm02(+)=cp10 and na01(+)=pa09" & _
'      " union select cp01,cp02,cp03,cp04,cp05,tm05,cp09,0,nvl(cp14,'68005'),cp10,decode(tm10,'020',cpm04,cpm03) cp10n, na03" & _
'      " From acc0m0, caseprogress, trademark, casepropertymap, nation" & _
'      " where a0m01='" & Text2 & "' and cp60(+)=a0m02 and cp01='CFT' and cp79=0 and cp27 is null and cp57 is null" & _
'      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and cpm01(+)=cp01 and cpm02(+)=cp10 and na01(+)=tm10" & _
'      " order by cp01,cp09 asc,ud02 desc"
'
'   strExc(0) = "select cp01,cp02,cp03,cp04,cp05,pa05,cp09,ud02,ud04,cp10,decode(pa09,'020',cpm04,cpm03) cp10n, na03" & _
'      " From acc0m0,acc0j0, caseprogress, undeliveredrec, patent, casepropertymap, nation" & _
'      " where a0m01='" & Text2 & "' and a0j13(+)=a0m02 and cp09(+)=a0j01 and cp01='CFP' and cp79=0 and cp27 is null and cp57 is null" & _
'      " and ud01=cp09 and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cpm01(+)=cp01 and cpm02(+)=cp10 and na01(+)=pa09" & _
'      " union select cp01,cp02,cp03,cp04,cp05,tm05,cp09,0,nvl(cp14,'68005'),cp10,decode(tm10,'020',cpm04,cpm03) cp10n, na03" & _
'      " From acc0m0,acc0j0, caseprogress, trademark, casepropertymap, nation" & _
'      " where a0m01='" & Text2 & "' and a0j13(+)=a0m02 and cp09(+)=a0j01 and cp01='CFT' and cp79=0 and cp27 is null and cp57 is null" & _
'      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and cpm01(+)=cp01 and cpm02(+)=cp10 and na01(+)=tm10" & _
'      " order by cp01,cp09 asc,ud02 desc"
'
'   intI = 1
'   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      With adoRst
'      .MoveFirst
'      Do While Not .EOF
'         If .Fields("cp09") <> stLstCP09 Then
'            stSubject = .Fields("cp01") & "-" & .Fields("cp02") & IIf(.Fields("cp03") & .Fields("cp04") = "000", "", "-" & .Fields("cp03") & "-" & .Fields("cp04")) & "收文號【" & .Fields("cp09") & "】申請國家【" & .Fields("na03") & "】已收款可逕行發文！"
'            stContent = "收文號：" & .Fields("cp09") & vbCrLf & _
'               "本所案號：" & .Fields("cp01") & "-" & .Fields("cp02") & IIf(.Fields("cp03") & .Fields("cp04") = "000", "", "-" & .Fields("cp03") & "-" & .Fields("cp04")) & vbCrLf & _
'               "案件名稱：" & .Fields("pa05") & vbCrLf & _
'               "案件性質：" & .Fields("cp10n") & vbCrLf & _
'               "收文日：" & Format(.Fields("cp05") - 19110000, "###/##/##")
'            'Modified by Morgan 2012/7/9 收信人請假，副本發案件職代不要彈訊息(會導致貸方資料顯示不正常)
'            'Modified by Morgan 2016/4/19 財務可能會多次進明細畫面，加判斷當天通知過的不再通知
'            strSql = "update mailcache set mc01=mc01 where mc02='" & .Fields("ud04") & "' and mc03=" & strSrvDate(1) & " and mc07='" & ChgSQL(stSubject) & "' and mc08='" & ChgSQL(stContent) & "'"
'            cnnConnection.Execute strSql, intI
'            If intI = 0 Then
'               bolMailSendOk = False
'               PUB_SendMail strUserNum, "" & .Fields("ud04"), "" & .Fields("cp09"), stSubject, stContent, , , , , , , , , , , False
'               If bolMailSendOk = True Then
'                  strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc05,mc06,mc07,mc08)" & _
'                     " values( '" & strUserNum & "','" & .Fields("ud04") & "',to_char(sysdate,'yyyymmdd')" & _
'                     ",to_char(sysdate,'hh24miss'),to_char(sysdate,'YYYYMMDD'),to_char(sysdate,'HH24MISS'),'" & ChgSQL(stSubject) & "','" & ChgSQL(stContent) & "')"
'                  cnnConnection.Execute strSql, intI
'               End If
'            End If
'            stLstCP09 = .Fields("cp09")
'         End If
'         .MoveNext
'      Loop
'      End With
'   End If
'
'   'Added by Morgan 2015/11/27
'   '所有T*案申請國家非台灣之非FMT案(即CP12非'F'字頭)，或是收款後送件CP141='2'之台灣案，於全額收款時發E-MAIL給承辦人
'   strExc(0) = "SELECT distinct A.*,NA03,CU04,DECODE(A3,'000',CPM03,CPM04) 案件性質,ST02" & _
'      " FROM (select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
'      ",TM05||SP05 案件名稱,SQLDATET(CP05) 收文日,TM10||SP09 A3,TM23||SP08 A4" & _
'      ",CP01,CP10,CP13,CP14" & _
'      " from ( select a0m02 From acc0m0, acc0j0, caseprogress" & _
'      " where a0m01='" & Text2 & "' and a0j13(+)=a0m02 and cp09(+)=a0j01" & _
'      " group by a0m02 having sum(cp79)=0" & _
'      "),acc0j0,caseprogress,trademark,servicepractice" & _
'      " where a0j13(+)=a0m02 and a0j02 like 'T%'" & _
'      " and cp09(+)=a0j01 AND CP27||CP57 IS NULL" & _
'      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
'      " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04" & _
'      " and ((nvl(tm10,sp09)<>'000' and cp12 not like 'F%') or cp141='2')" & _
'      ") A,STAFF,CASEPROPERTYMAP,nation,CUSTOMER" & _
'      " WHERE ST01(+)=CP13 AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
'      " and NA01(+)=A3 AND CU01(+)=SUBSTR(A4,1,8) AND CU02(+)=SUBSTR(A4,9)"
'
'   intI = 1
'   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      With adoRst
'      .MoveFirst
'      Do While Not .EOF
'         '考慮會有拆收據情形,收款金額改不帶出
'         stSubject = .Fields("本所案號") & "之" & .Fields("案件性質") & "已收款，請送件！"
'         stContent = "本所案號：" & .Fields("本所案號") & vbCrLf & _
'            "案件名稱：" & .Fields("案件名稱") & vbCrLf & _
'            "申請國家：" & .Fields("NA03") & vbCrLf & _
'            "申請人　：" & .Fields("CU04") & vbCrLf & _
'            "智權人員：" & .Fields("ST02") & vbCrLf & _
'            "案件性質：" & .Fields("案件性質") & vbCrLf & _
'            "收文日　：" & .Fields("收文日") & vbCrLf & _
'            ""
'         'Modified by Morgan 2016/4/19 財務可能會多次進明細畫面，加判斷當天通知過的不再通知
'         strSql = "update mailcache set mc01=mc01 where mc02='" & .Fields("CP14") & "' and mc03=" & strSrvDate(1) & " and mc07='" & ChgSQL(stSubject) & "' and mc08='" & ChgSQL(stContent) & "'"
'         cnnConnection.Execute strSql, intI
'         If intI = 0 Then
'            bolMailSendOk = False
'            PUB_SendMail strUserNum, "" & .Fields("CP14"), "", stSubject, stContent, , , , , , , , , , , False
'            If bolMailSendOk = True Then
'               strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc05,mc06,mc07,mc08)" & _
'                  " values( '" & strUserNum & "','" & .Fields("CP14") & "',to_char(sysdate,'yyyymmdd')" & _
'                  ",to_char(sysdate,'hh24miss'),to_char(sysdate,'YYYYMMDD'),to_char(sysdate,'HH24MISS'),'" & ChgSQL(stSubject) & "','" & ChgSQL(stContent) & "')"
'               cnnConnection.Execute strSql, intI
'            End If
'         End If
'         .MoveNext
'      Loop
'      End With
'   End If
'   'end 2015/11/27
'
'ErrHnd:
'   If Err.Number <> 0 Then
'      MsgBox Err.Description
'   End If
'end 2016/8/22

End Sub

'Add by Morgan 2005/11/2 檢查是否有支援紀錄需轉專業點數
Private Function CheckSupport(p_CP09 As String) As Boolean
   Dim stSQL As String, intJ As Integer, stVTB As String
   Dim cp(1 To 4) As String
   stSQL = "select cp01,cp02,cp03,cp04 from caseprogress,supporthour" & _
      " where cp09='" & p_CP09 & "'" & _
      " and sh12(+)=cp09 and upper(sh11)='V'"
   intJ = 1
   Set AdoRecordSet3 = ClsLawReadRstMsg(intJ, stSQL)
   If intJ = 1 Then
   
'Remove by Morgan 2007/11/29 改在 PUB_ChkP4SH
      'Add by Morgan 2007/11/1 加判斷本案，國內外案或多國案只要有扣過就不再扣
'      With AdoRecordSet3
'         CP(1) = .Fields("CP01")
'         CP(2) = .Fields("CP02")
'         CP(3) = .Fields("CP03")
'         CP(4) = .Fields("CP04")
'      End With
'      stVTB = "SELECT '" & CP(1) & CP(2) & CP(3) & CP(4) & "' C01 FROM DUAL" & _
'         " UNION SELECT CM01||CM02||CM03||CM04 FROM CASEMAP WHERE CM05='" & CP(1) & "' AND CM06='" & CP(2) & "' AND CM07='" & CP(3) & "' AND CM08='" & CP(4) & "' AND CM10='0'" & _
'         " UNION SELECT CM05||CM06||CM07||CM08 FROM CASEMAP WHERE CM01='" & CP(1) & "' AND CM02='" & CP(2) & "' AND CM03='" & CP(3) & "' AND CM04='" & CP(4) & "' AND CM10='0'" & _
'         " UNION SELECT CR05||CR06||CR07||CR08 FROM CASERELATION WHERE CR01='" & CP(1) & "' AND CR02='" & CP(2) & "' AND CR03='" & CP(3) & "' AND CR04='" & CP(4) & "'" & _
'         " UNION SELECT CR05||CR06||CR07||CR08 FROM CASERELATION WHERE (CR01,CR02,CR03,CR04) IN (SELECT CM01,CM02,CM03,CM04 FROM CASEMAP WHERE CM05='" & CP(1) & "' AND CM06='" & CP(2) & "' AND CM07='" & CP(3) & "' AND CM08='" & CP(4) & "' AND CM10='0')" & _
'         " UNION SELECT CR05||CR06||CR07||CR08 FROM CASERELATION WHERE (CR01,CR02,CR03,CR04) IN (SELECT CM05,CM06,CM07,CM08 FROM CASEMAP WHERE CM01='" & CP(1) & "' AND CM02='" & CP(2) & "' AND CM03='" & CP(3) & "' AND CM04='" & CP(4) & "' AND CM10='0')"
'      stSQL = "select 1 from acc021 where ax207=5000 and ax209='P1001' and substr(ax205,1,1)='4'" & _
'         " and ax214 in (" & stVTB & ")"
'      intJ = 1
'      Set AdoRecordSet3 = ClsLawReadRstMsg(intJ, stSQL)
'      If intJ = 0 Then
'         CheckSupport = True
'      End If
      'end 2007/11/1
      
      CheckSupport = True
   End If
   
End Function

'Add by Morgan 2011/4/18
'考慮收文分配點數問題並將足額與不足額程式合併
'支援、出庭費扣點數規則:全額收才扣,分次收都提醒但不扣(支援要檢查是否已扣)
Private Sub Acc1p0SaveNew()
   Dim strRemark As String '摘要
   Dim strRemarkR As String '收入摘要
   Dim strRemarkF As String '規費摘要
   Dim strRemarkTax As String '扣繳摘要
   Dim strAccNo As String '會計科目
   Dim strR As String '收入科目
   Dim strF As String '規費科目
   Dim strYes As String '是否合併
   Dim strRSystemNo As String '相關案系統別
   Dim strAccNo4SH As String '轉專業點數科目
   Dim bol2SupportOK As Boolean '本次是否可扣支援點數
   Dim bolXFee As Boolean '是否要扣出庭費
   Dim bolXFeeOK As Boolean '本次是否可扣出庭費
   Dim lngAmt As Long '分錄金額
   Dim lngAmtR As Long '服務費
   Dim lngAmtF As Long '規費
   Dim bolRFeeClean As Boolean '服務費是否收全額
   Dim bolAssignDone As Boolean '是否已分配點數
   Dim lngNetAmtR As Long '銷帳後服務費
   Dim lngNetAmtF As Long '銷帳後規費
   'Add by Morgan 2011/10/7
   Dim lngShrAmt As Long '分配金額
   Dim bolLawyerGuei As Boolean '是否委任律師為桂律師
   Dim strLawyerName As String '律師
   Dim strLawyerNameC As String '出庭律師 Added by Morgan 2023/7/18
   Dim strA1P30 As String '其他對沖
   Dim strA1P16 As String '業務對沖
   Dim strLawNo As String '承辦人編號
   Dim lngMaxFee As Long '最大規費
   Dim strMaxFeeNo As String '最大規費科目項次
   Dim strFeeItemNo As String '目前規費科目項次
   Dim bolTaxDone As Boolean '是否稅額分錄已新增
   Dim lngReceiptTax As Long '整張收據稅額
   Dim lngCaseTax As Long '本案稅額
   'Added by Morgan 2021/1/19
   Dim strA1P17 As String '本所案號對沖
   Dim bolLawFeeDone As Boolean '律師出庭費已扣
   Dim bolProFeeDone As Boolean '專業部規費已扣
   Dim lngTemp As Long 'Added by Morgan 2021/7/28
   Dim bolNewCourtFee As Boolean 'Added by Morgan 2022/12/12
   Dim strA1P23 As String, bolRcptClean As Boolean 'Added by Morgan 2023/9/25
   Dim strAMT2492 As String 'Added by Morgan 2023/11/1 點數保留2192金額
   Dim strDeptX As String 'Added by Morgan 2024/11/13 科目對應部門
   
On Error GoTo Checking
   
   'Modify by Morgan 2011/9/28 考慮多對多收據改語法
   'Modified by Morgan 2021/3/31 +lc01,lc47
   'Modified by Morgan 2022/7/13 +收文號排序(早收文通常為主項，法務案源收款出庭費的摘要才會正確，Ex:L-006512)
   strExc(0) = "select a.*,b.*,c.*,d.*,getcp10desc(c.cp01,c.cp10,a0j04) cp10N,na03,st03,st20,st02,lc01,lc47" & _
      " from acc1u0 a, acc0j0 b, caseprogress c, casepropertymap d,nation,staff,lawcase" & _
      " where  a1u01='" & Text2 & "' and a1u02='" & Adodc1.Recordset.Fields("t0202") & "'" & _
      " and a0j01(+)=a1u03 and a0j13(+)=a1u02" & _
      " and c.cp09(+)=a1u03 and cpm01(+)=c.cp01 and cpm02(+)=c.cp10 and na01(+)=a0j04" & _
      " and st01(+)=c.cp14 and lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04 order by a0j01 asc"
   intI = 1
   Set adocheck = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With adocheck
      Do While Not .EOF
         'Added by Morgan 2012/7/12
         '承辦人編號
         'modify by sonia 2019/2/13 +76012,因桂所長的部門由L01改管理部
         If .Fields("st03") = "L01" Or .Fields("st20") = "13" Or .Fields("cp14") = "76012" Then
            strLawNo = .Fields("cp14")
            
         'Added by Morgan 2022/11/9 補收款承辦非律師要再抓相關號的承辦
         ElseIf .Fields("cp01") & .Fields("cp10") = "L78" And Not IsNull(.Fields("cp43")) Then
            strExc(0) = "select cp14,st03,st20 from caseprogress,staff where cp09='" & .Fields("cp43") & "' and st01(+)=cp14"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp.Fields("st03") = "L01" Or RsTemp.Fields("st20") = "13" Or RsTemp.Fields("cp14") = "76012" Then
                  strLawNo = .Fields("cp14")
               End If
            End If
         'end 2022/11/9
         Else
            strLawNo = ""
         End If
         'end 2012/7/12
         
         'Add by Morgan 2011/10/7
         bolLawyerGuei = False
         strLawyerName = ""
         strA1P30 = ""
         If InStr(.Fields("cpm03"), "委任律師") > 0 Then
            If .Fields("cp14") = "76012" Then bolLawyerGuei = True
            'Removed by Morgan 2012/7/12 改判斷科目
            'strA1p30 = "" & .Fields("cp14") 'A1P30對沖代號(其它)存承辦人編號
            strExc(0) = "select s1.st02,s2.st02 from caseprogress,staff s1,CaseLawer,staff s2" & _
               " where cp09='" & .Fields("cp09") & "' and s1.st01(+)=cp14 and cl01(+)=cp09 and cl02(+)<>cp14 and s2.st01(+)=cl02"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If Not IsNull(RsTemp.Fields(0)) Then
                  strLawyerName = "/" & RsTemp.Fields(0)
               End If
               If Not IsNull(RsTemp.Fields(1)) Then
                  'Modified by Morgan 2023/7/17 修正出庭律師為承辦人時會重複的問題
                  'strLawyerName = strLawyerName & "/" & RsTemp.Fields(1)
                  If RsTemp.Fields(1) <> RsTemp.Fields(0) Then
                     strLawyerName = strLawyerName & "/" & RsTemp.Fields(1)
                     If RsTemp.RecordCount > 1 Then
                        RsTemp.MoveNext
                        Do While Not RsTemp.EOF
                           If Not IsNull(RsTemp.Fields(1)) Then
                              If RsTemp.Fields(1) <> RsTemp.Fields(0) Then
                                 strLawyerName = strLawyerName & "/" & RsTemp.Fields(1)
                              End If
                           End If
                           RsTemp.MoveNext
                        Loop
                     End If
                  End If
                  'end 2023/7/17
               End If
            End If
         End If
         'end 2011/10/7
         
         '本所案號
         strCaseNo = "" & .Fields("a0j02")
         '是否合併
         strYes = "" & .Fields("a0j07")
         
         '出庭費
         bolXFee = False
         '台灣案沒有合併且服務費10000以上的需判斷是否有出庭費(判斷是否合併是因為舊的做法出庭費會設為要合併)
         If .Fields("a0j04") = "000" And strYes <> "Y" And (Val("" & .Fields("cp16")) - Val("" & .Fields("cp17"))) >= 10000 Then
            '專利
            If (.Fields("cp01") = "P" Or .Fields("cp01") = "FCP") Then
               If InStr("211,212", .Fields("cp10")) > 0 Then
                  'Modified by Morgan 2018/12/13
                  'bolXFee = True
                  If PUB_ChkNoXFee("" & .Fields("cp09")) = False Then
                     bolXFee = True
                  End If
                  'end 2018/12/13
               ElseIf InStr("503,507,506", .Fields("cp10")) > 0 Then
                  strExc(0) = "select * from caseprogress a where cp09='" & .Fields("a0j01") & "'" & _
                     " and exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05>=a.cp05 and b.cp10 in ('211','212'))" & _
                     " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05>=a.cp05 and b.cp10 in ('211','212') and b.cp16>0)"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     bolXFee = True
                  End If
               End If
            '商標(FCT不扣)
            ElseIf (.Fields("cp01") = "T") Then
               If InStr("204,205", .Fields("cp10")) > 0 Then
                  bolXFee = True
                  '2013/8/19 ADD BY SONIA 葉經理說訴願的言詞辯論為商標處的人處理,故不扣出庭費T-182351
                  strExc(0) = "select b.cp10,c.cp10 from caseprogress a,caseprogress b,caseprogress c where a.cp09='" & .Fields("a0j01") & "'" & _
                     " and a.cp43=b.cp09(+) and b.cp43=c.cp09(+)"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If "" & RsTemp.Fields(0) = "401" Or "" & RsTemp.Fields(1) = "401" Then bolXFee = False
                  End If
                  '2013/8/19 END
               ElseIf InStr("403,408,407", .Fields("cp10")) > 0 Then
                  strExc(0) = "select * from caseprogress a where cp09='" & .Fields("a0j01") & "'" & _
                     " and exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05>=a.cp05 and b.cp10 in ('204','205'))" & _
                     " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05>=a.cp05 and b.cp10 in ('204','205') and b.cp16>0)"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     bolXFee = True
                  End If
               End If
            End If
         End If
         
         '會計科目
         If .Fields("a0j04") = "000" Then
            strR = "" & .Fields("cpm11")
            strF = "" & .Fields("cpm12")
         Else
            strR = "" & .Fields("cpm24")
            strF = "" & .Fields("cpm25")
         End If
         
         '相關案系統別
         strRSystemNo = ""
         strExc(0) = "select cr05 from caserelation1 where cr01 = '" & .Fields("cp01") & "'" & _
            " and cr02 = '" & .Fields("cp02") & "' and cr03 = '" & .Fields("cp03") & "'" & _
            " and cr04 = '" & .Fields("cp04") & "'"
         intI = 1
         Set adoaccsum = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strRSystemNo = "" & adoaccsum.Fields("cr05")
         End If
         
         '法務有建立相關案者改會計科目
         Select Case .Fields("cp01")
            Case "L"
               Select Case strRSystemNo
                  Case "P", "T", "TC", "CFC"
                     strR = "414101"
               End Select
         End Select
         If strR = "" Then
            strR = "41"
         End If
         
         'add by sonia 2022/1/22 從下面移上來，因79075收文一律改M0100
         '作帳智權人員
         strManNo = SalesNoToSales(strManNo, strR)
         If strManNo = "" Then
            strManNo = "M0100"
         End If
         'end 2022/1/22
         
         '智權人員簡稱
         'modify by sonia 2022/1/22 Adodc1.Recordset.Fields("t0212")改用strManNo
         'strExc(0) = "select sn01 from salesno where sn02 = '" & Adodc1.Recordset.Fields("t0212") & "'"
         strExc(0) = "select sn01 from salesno where sn02 = '" & strManNo & "'"
         intI = 1
         Set adoaccsum = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strMan = "" & adoaccsum.Fields("sn01")
         End If
         '案件性質
         'Modified by Morgan 2011/12/27 取消 a0j20
         strProperty = "" & .Fields("cp10N")
         '申請國家
         'Modified by Morgan 2011/12/29 取消 a0j21
         strNation = "" & .Fields("na03")
         
         
         '收入摘要
         strRemarkR = strMan & "/" & strCompany
         'Modify by Morgan 2011/10/7 收入科目摘統一為非台灣的加申請國家
         If .Fields("a0j04") = "000" Then
            strRemarkR = strRemarkR & "/" & strProperty
         Else
            strRemarkR = strRemarkR & "/" & strNation & "/" & strProperty
         End If
         'end 2011/10/7
         
         '出庭費摘要-舊的做法出庭費會設為要合併
         If .Fields("a0j04") = "000" And strYes = "Y" Then
            If ((.Fields("cp01") = "P" Or .Fields("cp01") = "FCP") And InStr("211,212,503,507,506", .Fields("cp10")) > 0) Or _
               ((.Fields("cp01") = "T" Or .Fields("cp01") = "FCT") And InStr("204,205,403,408,407", .Fields("cp10")) > 0) Then
               'modify by sonia 2017/9/8 F10609226 婷說仍要保留智權人員/客戶名稱
               'strRemarkR = strProperty & "/出庭費"
               strRemarkR = strRemarkR & "/出庭費"
            End If
         End If
         
         strRemarkR = strRemarkR & strLawyerName 'Add by Morgan 2011/10/7
         
'cancel by sonia 2022/1/22 移到上面去
'         '作帳智權人員
'         strManNo = SalesNoToSales(strManNo, strR)
'         If strManNo = "" Then
'            strManNo = "M0100"
'         End If
'end 2022/1/22

         '部門
         'MODIFY BY SONIA 2016/1/4
         'Select Case Mid(strR, 1, 4)
         '   Case "4101", "4151"
         '      strDept = "T"
         '   Case "4111"
         '      strDept = "P"
         '   Case "4121"
         '      strDept = "CFT"
         '   Case "4172"
         '      If strR = "417202" Then
         '         strDept = "T"
         '      Else
         '         strDept = "FCT"
         '      End If
         '   Case "4131"
         '      strDept = "CFP"
         '   Case "4141"
         '      strDept = "L"
         '   Case "4171"
         '      strDept = "FCP"
         '   Case "4181"
         '      strDept = "L"
         '   Case "4161"
         '      strDept = "FCL"
         '   Case Else
         '      strDept = "TOT"
         'End Select
         If Left(strR, 1) = "4" Then
            strDept = PUB_GETAccNODept(strR, strDept)
         Else
            strDept = "TOT"
         End If
         'END 2016/1/4
         
'2013/4/8 CANCEL BY SONIA D102031970(L-005113瑞婷說此設定已於二年前取消)
'         '規費科目
'         Select Case .Fields("cp01")
'            Case "L", "LA"
'               If strRSystemNo <> "" Then
'                  strF = "610103"
'                  strDept = "L"
'               End If
'         End Select
'
'         If strF = "" Then
'            strF = "22"
'         End If
'2013/4/8 END
         
         '規費摘要
         'Modify by Morgan 2011/10/7 所有規費科目摘統一為非台灣的加申請國家及收款金額
         If .Fields("a0j04") = "000" Then
            strRemarkF = strMan & "/" & strCompany & "/" & strProperty
         Else
            strRemarkF = strMan & "/" & strCompany & "/" & strNation & "/" & strProperty & "/" & Format(Val("" & adocheck.Fields("a1u04").Value) + Val("" & adocheck.Fields("a1u05").Value), DDollar)
         End If
         'end 2011/10/7
         
         '舊的做法出庭費會設為要合併
         If .Fields("a0j04") = "000" And strYes = "Y" Then
            If ((.Fields("cp01") = "P" Or .Fields("cp01") = "FCP") And InStr("211,212,503,507,506", .Fields("cp10")) > 0) Or _
               ((.Fields("cp01") = "T" Or .Fields("cp01") = "FCT") And InStr("204,205,403,408,407", .Fields("cp10")) > 0) Then
               strRemarkF = strProperty & "/出庭費"
            End If
         End If
         
         strRemarkTax = strRemarkF
         strRemarkF = strRemarkF & strLawyerName 'Add by Morgan 2011/10/7
         
         '銷帳後服務費,銷帳後規費
         lngNetAmtR = Val("" & .Fields("cp16")) - Val("" & .Fields("cp17"))
         lngNetAmtF = Val("" & .Fields("cp17"))
         bolRFeeClean = False
         strExc(0) = "select sum(a1u07) S1,sum(a1u09) S2" & _
            " from acc1u0 where a1u03='" & .Fields("a0j01") & "'"
         intI = 1
         Set adoaccsum = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            lngNetAmtR = lngNetAmtR - Val("" & adoaccsum.Fields("S1"))
            lngNetAmtF = lngNetAmtF - Val("" & adoaccsum.Fields("S2"))
         End If
         
         lngAmtR = Val("" & .Fields("a1u04")) '本次收款服務費
         lngAmtF = Val("" & .Fields("a1u05")) '本次收款規費
                          
         '服務費是否全額收
         If lngAmtR = lngNetAmtR Then
            bolRFeeClean = True
         End If
         
         '服務費
         bolXFeeOK = False
         If bolXFee = True Then
            strXFeeCaseList = strXFeeCaseList & vbCrLf & .Fields("cp01") & "-" & .Fields("cp02") & "-" & .Fields("cp03") & "-" & .Fields("cp04")
            '全額收(服務費)才要扣
            If bolRFeeClean = True And lngAmtR >= 10000 Then
               bolXFeeOK = True
               lngAmtR = lngAmtR - 10000
            Else
               strXFeeCaseList = strXFeeCaseList & "(本次未扣)"
            End If
         End If
         
         '配合出庭分配點數
         bolAssignDone = False
         If .Fields("cp01") = "L" Then
            strExc(0) = "select a0n02,a0n03,a0n04,cp01,cp02,cp03,cp04 from acc0n0,caseprogress where a0n01='" & .Fields("a0j01") & "' and cp09(+)=a0n02 order by decode(a0n01,a0n02,1,2),a0n02"
            intI = 1
            Set adoaccsum = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strAssignCaseList = strAssignCaseList & vbCrLf & .Fields("cp01") & "-" & .Fields("cp02") & "-" & .Fields("cp03") & "-" & .Fields("cp04")
               If bolRFeeClean = True Then
                  Do While Not adoaccsum.EOF
                     '原收文號
                     If .Fields("a0j01") = adoaccsum.Fields("a0n02") Then
                        If adoaccsum.Fields("a0n03") > 0 Then
                           'Added by Morgan 2012/7/12
                           If strR = "414101" Or strR = "416101" Or strR = "416102" Then
                              strA1P30 = strLawNo
                           Else
                              strA1P30 = ""
                           End If
                           'end 2012/7/12
                           'Modify by Morgan 2011/10/7 考慮拆收據情形改依比例分配
                           lngShrAmt = adoaccsum.Fields("a0n03") * (.Fields("a0j09") / (Val("" & .Fields("cp16")) - Val("" & .Fields("cp17"))))
                           'Modified by Morgan 2013/12/19 收款會有J公司,取消 a1p01='1' 條件
                           strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                           'ADD BY SONIA 2016/1/4 105年起法務收入改其他部門收入(傳CP09以判斷案件性質及收文人員)
                           If Val(FCDate(strDate)) >= 1050101 And (Left(strR, 4) = "4141" Or Left(strR, 4) = "4161" Or Left(strR, 4) = "4181") Then
                              'Modified by Morgan 2016/2/1 入帳日期格式未轉
                              InsertLawACC1P0 strA1P01, "A", strSerialNo, Text2, strR, strDept, 0, Val(lngShrAmt), "", "", "", "", "", strRemarkR, strCompanyNo, strManNo, strCaseNo, Val(FCDate(strDate)), "", "", "", IIf(stra1p22 = "null", "", Replace(stra1p22, "'", "")), "", "", "", strYes, IIf(stra1p27 = "null", "", Replace(stra1p27, "'", "")), strA1P30, "", .Fields("a0j01")
                           Else
                           'END 2016/1/4
                              'modify by sonia 2016/1/29 2201xx規費科目的部門一律改用TOT
                              strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                                 " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strR & "', '" & IIf(Left(strR, 4) = "2201", MsgText(55), strDept) & "', 0, " & lngShrAmt & ", null, null, null, null, null, '" & ChgSQL(strRemarkR) & "', '" & strCompanyNo & "', '" & strManNo & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                              adoTaie.Execute strSql, intI
                           End If   'ADD BY SONIA 2016/1/4
                        End If
                        
                        If adoaccsum.Fields("a0n04") > 0 Then
                           'Modify by Morgan 2011/10/7 考慮拆收據情形改依比例分配
                           lngShrAmt = adoaccsum.Fields("a0n04") * (.Fields("a0j10") / Val("" & .Fields("cp17")))
                           strA1P16 = strManNo
                           'Add by Morgan 2011/10/7
                           If bolLawyerGuei Then
                              strA1P16 = "M0100"
                              strF = "414101"
                           End If
                           'Added by Morgan 2012/7/12
                           If strF = "414101" Or strF = "416101" Or strF = "416102" Or strF = "220113" Then
                              strA1P30 = strLawNo
                           Else
                              strA1P30 = ""
                           End If
                           'end 2012/7/12
                           'Modified by Morgan 2013/12/19 收款會有J公司,取消 a1p01='1' 條件
                           strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)

                           'ADD BY SONIA 2016/1/4 105年起法務收入改其他部門收入(傳CP09以判斷案件性質及收文人員)
                           If Val(FCDate(strDate)) >= 1050101 And (Left(strF, 4) = "4141" Or Left(strF, 4) = "4161" Or Left(strF, 4) = "4181") Then
                              InsertLawACC1P0 strA1P01, "A", strSerialNo, Text2, strF, strDept, 0, Val(lngShrAmt), "", "", "", "", "", strRemarkF, strCompanyNo, strA1P16, strCaseNo, strDate, "", "", "", IIf(stra1p22 = "null", "", Replace(stra1p22, "'", "")), "", "", "", strYes, IIf(stra1p27 = "null", "", Replace(stra1p27, "'", "")), strA1P30, "", .Fields("a0j01")
                           Else
                           'END 2016/1/4
                              '2014/2/11 modify by sonia E10228435(F10300038第三項次)部門應仍為L,且收款已不使用610103
                              'strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30) " & _
                                 "values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strF & "', '" & IIf(strF = "610103", strDept, MsgText(55)) & "', 0, " & lngShrAmt & ", null, null, null, null, null, '" & strRemarkF & "', '" & strCompanyNo & "', '" & strA1p16 & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1p30 & "')"
                              'modify by sonia 2016/1/29 2201xx規費科目的部門一律改用TOT
                              strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30) " & _
                                 "values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strF & "', '" & IIf(Left(strF, 4) = "2201", MsgText(55), strDept) & "', 0, " & lngShrAmt & ", null, null, null, null, null, '" & ChgSQL(strRemarkF) & "', '" & strCompanyNo & "', '" & strA1P16 & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                              '2014/2/11 END
                              adoTaie.Execute strSql, intI
                           End If   'ADD BY SONIA 2016/1/4
                           
                           strFeeItemNo = strSerialNo 'Added by Morgan 2014/2/12
                        End If
                        
                     '分配收文號
                     '案號對沖放分配的
                     Else
                        If adoaccsum.Fields("a0n03") > 0 Then
                           'Modify by Morgan 2011/10/7 考慮拆收據情形改依比例分配
                           lngShrAmt = adoaccsum.Fields("a0n03") * (.Fields("a0j09") / (Val("" & .Fields("cp16")) - Val("" & .Fields("cp17"))))
                           
                           '支援收入科目固定用'411104'
                           strRemark = strRemarkR & "/" & adoaccsum.Fields("cp01") & "支援"
                           'Modified by Morgan 2013/12/19 收款會有J公司,取消 a1p01='1' 條件
                           strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                           strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27)" & _
                              " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '411104', '" & adoaccsum.Fields("cp01") & "', 0, " & lngShrAmt & ", null, null, null, null, null, '" & strRemark & "', '" & strCompanyNo & "', '" & strManNo & "', '" & adoaccsum("cp01") & adoaccsum("cp02") & adoaccsum("cp03") & adoaccsum("cp04") & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ")"
                           adoTaie.Execute strSql, intI
                        End If
                        If adoaccsum.Fields("a0n04") > 0 Then
                           'Modify by Morgan 2011/10/7 考慮拆收據情形改依比例分配
                           lngShrAmt = adoaccsum.Fields("a0n04") * (.Fields("a0j10") / Val("" & .Fields("cp17")))
                           
                           strRemark = strRemarkF & "/" & adoaccsum.Fields("cp01") & "支援"
                           'Modified by Morgan 2013/12/19 收款會有J公司,取消 a1p01='1' 條件
                           strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                           '支援規費科目固定用'220102' Add by Morgan 2011/10/11 --辜
                           strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) " & _
                              "values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '220102', '" & MsgText(55) & "', 0, " & lngShrAmt & ", null, null, null, null, null, '" & strRemark & "', '" & strCompanyNo & "', '" & strManNo & "', '" & adoaccsum("cp01") & adoaccsum("cp02") & adoaccsum("cp03") & adoaccsum("cp04") & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ")"
                           adoTaie.Execute strSql, intI
                        End If
                     End If
                     adoaccsum.MoveNext
                  Loop
                  bolAssignDone = True
               Else
                  strAssignCaseList = strAssignCaseList & "(本次未分配)"
               End If
            End If
         End If
         
         'Added by Morgan 2025/7/16
         '律師庭費:不管是否案源，有輸出庭費的都要扣--秀玲
         If lngAmtR > 0 And strA1P01 = "L" Then
            'Modified by Morgan 2025/8/19 直接用收文號抓,否則若收據有兩收文號會重複扣(Ex:E11415536)
            'strExc(0) = "select cl01,cl02,cl03,st02 from acc1u0,caselawer,staff" & _
               " where a1u01='" & Text2 & "' and a1u02='" & Adodc1.Recordset.Fields("t0202") & "'" & _
               " and cl01(+)=a1u03 and cl03>0 and st01(+)=cl02 order by cl03 desc,cl02"
            
            strExc(0) = "select cl01,cl02,cl03,st02 from caselawer,staff where cl01='" & .Fields("cp09") & "' and cl03>0 and st01(+)=cl02 order by cl03 desc,cl02"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               '客戶名稱帶前6字
               strRemarkR = strTTManSN & "/" & Left(strCompany, 6) & "/" & strProperty
               bolNewCourtFee = True
               Do While Not RsTemp.EOF
                  strLawNo = "" & RsTemp("cl02")
                  strLawyerNameC = "" & RsTemp("st02")
                  lngAmt = Val("" & RsTemp("cl03"))
                  
                  strXFeeCaseList = strXFeeCaseList & vbCrLf & strCaseNo & " (" & strLawyerNameC & ":" & lngAmt & ")"
                  If bolRFeeClean = True And lngAmtR >= lngAmt Then
                     strA1P30 = strLawNo
                     strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                     '摘要的律師只需帶扣出庭費的律師
                     strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                        " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '220113', 'TOT', 0, " & lngAmt & ", null, null, null, null, null, '" & ChgSQL(Replace(strRemarkR, strLawyerName, "") & "/" & strLawyerNameC) & "', '" & strCompanyNo & "', '" & strA1P30 & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                     adoTaie.Execute strSql, intI
                     
                     lngAmtR = lngAmtR - lngAmt
                     
                  Else
                     strXFeeCaseList = strXFeeCaseList & " (本次未扣)"
                  End If
                  RsTemp.MoveNext
               Loop
               bolLawFeeDone = True
            End If
         End If
         'end 2025/7/16
               
         '服務費
         If lngAmtR > 0 Then
            If bolAssignDone = False Then
               'Added by Morgan 2021/1/13 案源法律所收據
               '一般法律訴訟-分潤第1次(A1):15%智權,10%案源(規費)
               '一般法律訴訟-分潤第2次(A2):10%智權,5%案源(規費)
               '智財權/一般法律且非訴訟 (A3):全部回智慧所
               '智財權且訴訟(A4):法律所留15000律師費，其他全回智慧所
               'IP案 配合開庭(B1):法律所留15000律師費及相關規費，餘款全部回智慧所；P案-智慧所扣5000給專利處(規費)
               'IP案 行政訴訟(B2):法律所留5000律師費及規費，其他全部回智慧所；智慧所扣5000給商標部(規費)/專利處(規費)
               '           B2例外:專利商標之行政訴訟上訴或答辯都不扣5000律師費及5000出庭費
               If strA1P01 = "L" And (strLOS02 = "A1" Or strLOS02 = "A2" Or strLOS02 = "A3" Or strLOS02 = "A4" Or strLOS02 = "B1" Or strLOS02 = "B2") Then
                  lngAmt = 0
                  '律師庭費
                  'Modified by Morgan 2021/10/22 a1p15原來放L0100改為放出庭律師(同a1p30)
                  'Modified by Morgan 2021/4/29 B2有例外
                  'If strLOS02 <> "A3" And bolLawFeeDone = False Then
                  'Modified by Morgan 2021/10/26 A類要出庭的案件性質才要扣(規費科目非220113的不扣)
'Removed by Morgan 2025/7/17 移到外層(上面)，改不管是否案源，有輸出庭費的都要扣--秀玲
'                  If strLOS02 <> "A3" And bolLawFeeDone = False _
'                     And Not (strLOS02 = "B2" And bolB2NeeCourt = False) _
'                     And Not (Left(strLOS02, 1) = "A" And strF <> "220113") Then
'
'                     '客戶名稱帶前6字
'                     strRemarkR = strTTManSN & "/" & Left(strCompany, 6) & "/" & strProperty
'
'                     '220113 應付規費－律師庭費
'                     'Added by Morgan 2022/12/12 承辦人有在出庭律師檔時直接抓出庭費CL03
'                     bolNewCourtFee = False
'                     'Modified by Morgan 2023/7/5 and cl02=cp14 -> cl03>0, 承辦人可能非出庭律師(ex:蔣律師掛承辦人但不出庭)
'                     strExc(0) = "select cl01,cl02,cl03,st02 from (select decode(a.cp01||a.cp10,'L78',b.cp09,'FCL997',b.cp09,a.cp09) RNo" & _
'                        " from caseprogress a,caseprogress b  where a.cp162='" & .Fields("cp162") & "'" & _
'                        " and a.cp01='" & .Fields("cp01") & "' and b.cp09(+)=a.cp43) X,caseprogress,caselawer,staff" & _
'                        " where cp09(+)=RNo and cl01(+)=cp09 and cl03>0 and st01(+)=cl02 order by cl03 desc,cl02"
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                     If intI = 1 Then
'                        bolNewCourtFee = True
'                        strLawNo = "" & RsTemp("cl02")
'                        strLawyerNameC = "" & RsTemp("st02")
'                        lngAmt = Val("" & RsTemp("cl03"))
'                     Else
'                     'end 2022/12/12
'                        strLawyerNameC = "" & .Fields("st02")
'                        lngAmt = GetLawFee(strLawNo, Adodc1.Recordset.Fields("t0208"), IIf(strLOS02 = "B2", "34", ""))
'                     End If
'
'                     strXFeeCaseList = strXFeeCaseList & vbCrLf & strCaseNo
'                     If bolRFeeClean = True And lngAmtR >= lngAmt Then
'                        strA1P30 = strLawNo
'                        strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
'                        'Modified by Morgan 2023/7/18 摘要的律師只需帶扣出庭費的律師
'                        strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
'                           " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '220113', 'TOT', 0, " & lngAmt & ", null, null, null, null, null, '" & ChgSQL(Replace(strRemarkR, strLawyerName, "") & "/" & strLawyerNameC) & "', '" & strCompanyNo & "', '" & strA1P30 & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
'                        adoTaie.Execute strSql, intI
'
'                        strXFeeCaseList = strXFeeCaseList & " (" & strLawyerNameC & ":" & lngAmt & ")"
'                        lngAmtR = lngAmtR - lngAmt
'
'                        '多律師出庭
'                        'Modified by Morgan 2022/11/9 補收款要抓相關號的承辦律師
'                        'Modified by Morgan 2022/12/12 承辦人有在出庭律師檔時直接抓出庭費CL03
'                        strExc(0) = "select distinct cl02,cl03,st02 from caseprogress a,caseprogress b,caselawer,staff" & _
'                           " where a.cp162='" & .Fields("cp162") & "' and a.cp01='" & .Fields("cp01") & "' and b.cp09(+)=a.cp43" & _
'                           " and cl01=decode(a.cp01||a.cp10,'L78',b.cp09,'FCL997',b.cp09,a.cp09) and cl02<>'" & strLawNo & "' and cl03>0 and st01(+)=cl02 order by cl03 desc,cl02"
'                        intI = 1
'                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                        If intI = 1 Then
'                           Do While Not RsTemp.EOF
'                              strA1P30 = RsTemp("cl02")
'                              'Added by Morgan 2022/12/12
'                              If bolNewCourtFee Then
'                                 lngAmt = RsTemp("cl03")
'                              Else
'                              'end 2022/12/12
'                                 lngAmt = GetLawFee(strA1P30, Adodc1.Recordset.Fields("t0208"), IIf(strLOS02 = "B2", "34", ""))
'                              End If
'                              strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
'                              strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
'                                 " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '220113', 'TOT', 0, " & lngAmt & ", null, null, null, null, null, '" & ChgSQL(strRemarkR & "/" & RsTemp("st02")) & "', '" & strCompanyNo & "', '" & strA1P30 & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
'                              adoTaie.Execute strSql, intI
'
'                              strXFeeCaseList = strXFeeCaseList & " (" & RsTemp("st02") & ":" & lngAmt & ")"
'                              lngAmtR = lngAmtR - lngAmt
'
'                              RsTemp.MoveNext
'                           Loop
'                        End If
'                        bolLawFeeDone = True
'
'                     Else
'                        strXFeeCaseList = strXFeeCaseList & " (本次未扣)"
'                     End If
'                  End If
'end 2025/7/17
                  
                  If strLOS02 = "A1" Or strLOS02 = "A2" Then
                     strRemarkR = strTTManSN & "/" & Left(strCompany, 6) & "/" & strProperty & "/" & .Fields("st02")
                     strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                     strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27)" & _
                        " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strR & "', 'L', 0, " & lngAmtR & ", null, null, null, null, null, '" & ChgSQL(strRemarkR) & "', '" & strCompanyNo & "', 'L0100', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ")"
                     adoTaie.Execute strSql, intI
                     
                  Else
                  
                     If strLOS02 = "A3" Then
                        'Modified by Morgan 2021/10/29 +案件性質為0顧問聘任 Ex:F11009846(LA002889,E11024028)--瑞婷
                        If .Fields("cp01") = "LA" And .Fields("cp10") = "0" Then
                           strAccNo = "240702" '代收款項-顧問
                        Else
                           strAccNo = "240709" '代收款項-其他
                        End If
                     Else
                        strAccNo = "240701" '代收款項-訴訟
                     End If
                     
                     
                     strExc(0) = .Fields("cp01") & .Fields("cp02") & IIf(.Fields("cp04") = "00", IIf(.Fields("cp03") = "0", "", "-" & .Fields("cp03")), "-" & .Fields("cp03") & "-" & .Fields("cp04"))
                     '客戶名稱帶前6字
                     strRemarkR = Left(strCompany, 6) & "/" & strExc(0) & strProperty
                     strRemark = strTTManSN & "/" & Left(strCompany, 6) & "/" & strProperty
                     strA1P30 = ""
                     
                     strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                     'L公司-代收款項
                     strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                        " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strAccNo & "', 'TOT', 0, " & lngAmtR & ", null, null, null, null, null, '" & ChgSQL("智慧所/" & strRemarkR) & "', 'X82357000', 'L0100', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                     adoTaie.Execute strSql, intI
                     
                     'P/T案號、專業點數(規費)
                     lngAmt = 0
                     If strLOS02 = "B1" Or strLOS02 = "B2" Then
                        strA1P17 = GetPTCase(.Fields("cp162"), lngAmt)
                     Else
                        strA1P17 = strCaseNo
                     End If
                     
                     '1公司-應收帳款
                     strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                     strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                        " values ('1', 'A', '" & strSerialNo & "', '" & Text2 & "', '1133', 'TOT', " & lngAmtR & ", 0, null, null, null, null, null, '" & ChgSQL("法律所/" & strTTManSN & "/" & strRemarkR) & "', 'X82357000', '" & strTTMan & "', '" & strA1P17 & "', " & Val(FCDate(strDate)) & ", null, null, null, " & strA1P22_TT & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                     adoTaie.Execute strSql, intI
   
                     'B1 P案-智慧所扣5000給專利處(規費),B2 智慧所扣5000給商標部/專利處(規費)
                     If lngAmt > 0 Then
                        If bolProFeeDone = False Then
                           strProFeeCaseList = strProFeeCaseList & vbCrLf & strA1P17 & " (" & lngAmt & ")"
                           If bolRFeeClean = True And lngAmtR >= lngAmt Then
                              strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                              'Modified by Morgan 2021/3/3 規費部門應為TOT
                              'Modified by Morgan 2024/11/13 改科目 -> 220113 應付規費－律師庭費
                              'If Left(strA1P17, 1) = "T" Then
                              '   strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                              '      " values ('1', 'A', '" & strSerialNo & "', '" & Text2 & "', '220101', 'TOT', 0, " & lngAmt & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & strCompanyNo & "', '" & strTTMan & "', '" & strA1P17 & "', " & Val(FCDate(strDate)) & ", null, null, null, " & strA1P22_TT & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                              'Else
                              '   strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                              '      " values ('1', 'A', '" & strSerialNo & "', '" & Text2 & "', '220102', 'TOT', 0, " & lngAmt & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & strCompanyNo & "', '" & strTTMan & "', '" & strA1P17 & "', " & Val(FCDate(strDate)) & ", null, null, null, " & strA1P22_TT & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                              'End If
                              'end 2024/11/13
                                 strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                                    " values ('1', 'A', '" & strSerialNo & "', '" & Text2 & "', '220113', 'TOT', 0, " & lngAmt & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & strCompanyNo & "', '" & strTTMan & "', '" & strA1P17 & "', " & Val(FCDate(strDate)) & ", null, null, null, " & strA1P22_TT & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                              'end 2024/11/13
                              adoTaie.Execute strSql, intI
                              lngAmtR = lngAmtR - lngAmt
                              bolProFeeDone = True
                              
                           Else
                              strProFeeCaseList = strProFeeCaseList & " (本次未扣)"
                           End If
                        End If
                     End If
                     
                     strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                     'Added by Morgan 2021/3/31
                     If strLOS02 = "A3" And Not IsNull(.Fields("lc01")) And InStr("" & .Fields("LC47"), "專利") + InStr("" & .Fields("LC47"), "商標") + InStr("" & .Fields("LC47"), "著作權") = 0 Then
                        'Modified by Morgan 2022/9/14 回智慧所科目固定用 490102 其他各項收入
                        strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                           " values ('1', 'A', '" & strSerialNo & "', '" & Text2 & "', '490102', 'SAL', 0, " & lngAmtR & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & strCompanyNo & "', '" & strTTMan & "', '" & strA1P17 & "', " & Val(FCDate(strDate)) & ", null, null, null, " & strA1P22_TT & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                        adoTaie.Execute strSql, intI
                     'Added by Morgan 2021/10/25
                     ElseIf strLOS02 = "B1" Then
                        '商標:1/3做CCT-爭議(原案件性質設定科目)點數,2/3做CCT-法務(410110)
                        If Left(strA1P17, 1) = "TC" Then
                           strAccNo = "415102"
                           'Modified by Morgan 2024/8/2
                           'strR = "410110"
                           strR = "490102"
                           'end 2024/8/2
                           strDeptX = "T" 'Added by Morgan 2024/11/13
                        ElseIf Left(strA1P17, 1) = "T" Then
                           strAccNo = "410104"
                           'Modified by Morgan 2024/8/2
                           'strR = "410110"
                           strR = "490102"
                           'end 2024/8/2
                           strDeptX = "T" 'Added by Morgan 2024/11/13
                        '專利:1/3做CCP-爭議(原案件性質設定科目)點數,2/3做CCP-法務(411107)
                        Else
                           strAccNo = "411104"
                           'Modified by Morgan 2024/8/2
                           'strR = "411107"
                           strR = "490102"
                           strDeptX = "P" 'Added by Morgan 2024/11/13
                        End If
                        
                        lngAmt = Round(lngAmtR / 3)
                        'Modified by Morgan 2024/11/13 改部門 SAL-> strDeptX
                        strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                           " values ('1', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strAccNo & "', '" & strDeptX & "', 0, " & lngAmt & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & strCompanyNo & "', '" & strTTMan & "', '" & strA1P17 & "', " & Val(FCDate(strDate)) & ", null, null, null, " & strA1P22_TT & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                        adoTaie.Execute strSql, intI
                        
                        lngAmtR = lngAmtR - lngAmt
                        strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                        strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                           " values ('1', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strR & "', '" & strDeptX & "', 0, " & lngAmtR & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & strCompanyNo & "', '" & strTTMan & "', '" & strA1P17 & "', " & Val(FCDate(strDate)) & ", null, null, null, " & strA1P22_TT & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                        adoTaie.Execute strSql, intI
                     
                     Else
                     'end 2021/3/31
                     
                        InsertLawACC1P0 "1", "A", strSerialNo, Text2, strR, strDept, 0, Val(lngAmtR), "", "", "", "", "", strRemark, strCompanyNo, strTTMan, strA1P17, Val(FCDate(strDate)), "", "", "", Replace(strA1P22_TT, "'", ""), "", "", "", strYes, IIf(stra1p27 = "null", "", Replace(stra1p27, "'", "")), strA1P30, "", .Fields("a0j01")
                     End If
                  End If
                  
               'Added by Morgan 2021/3/9
               ElseIf strA1P01 = "1" And (strLOS02 = "A1" Or strLOS02 = "A2") Then
                  strRemarkR = "法律所/" & strMan & "/" & MidB(strLRcpTitle, 1, 16) & "/" & strLCaseNo & "/" & strProperty
                  strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                  strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                     " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strR & "', 'SAL', 0, " & lngAmtR & ", null, null, null, null, null, '" & ChgSQL(strRemarkR) & "', '" & strCompanyNo & "', '" & strManNo & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                  adoTaie.Execute strSql, intI
               'end 2021/3/9
               
               'Added by Morgan 2021/5/5
               '部門 L
               'LA-999999: 科目414102 , 智權人員L0200
               '非LA-999999：科目依案件性質表設定，智權人員依收據智權人員
               ElseIf strA1P01 = "L" And Left(strCaseNo, 2) = "LA" And strLOS02 = "" Then
                  strAMT2492 = "" 'Added by Morgan 2023/11/1
                  strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                  If strCaseNo = "LA999999000" Then
                     strR = "414102"
                     strManNo = "L0200"
                     strDept = "TOT"
                  'Added by Morgan 2023/11/1
                  Else
                     '顧問聘任簽約多年僅第一年做收入,其他做2492點數保留(部門TOT)
                     If .Fields("cp10") = "0" Then
                        intI = (.Fields("cp54") - .Fields("cp53")) \ 10000 + 1
                        If intI > 1 Then
                           strAMT2492 = lngAmtR - (lngAmtR \ intI)
                        End If
                     End If
                  'end 2023/11/1
                  End If
                  
                  'Modified by Morgan 2023/11/1  - Val(strAMT2492
                  strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                        " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strR & "', '" & strDept & "', 0, " & (lngAmtR - Val(strAMT2492)) & ", null, null, null, null, null, '" & ChgSQL(strRemarkR) & "', '" & strCompanyNo & "', '" & strManNo & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                  adoTaie.Execute strSql, intI
                  
                  'Added by Morgan 2023/11/1
                  '顧問聘任簽約多年僅第一年做收入,其他做2492點數保留(部門TOT)
                  If Val(strAMT2492) > 0 Then
                     strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                     strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                        " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '2492', 'TOT', 0, " & strAMT2492 & ", null, null, null, null, null, '" & ChgSQL(strRemarkR) & "', '" & strCompanyNo & "', '" & strManNo & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                     adoTaie.Execute strSql, intI
                     If .Fields("cp16") > lngAmtR Then '多年聘任簽約且為部分收款
                        MsgBox "顧問聘任簽約多年且為部分收款，請自行調整各收入科目金額！", vbInformation
                     End If
                  End If
                  'end 2023/11/1
   
               'end 2021/5/5
               
               'Added by Morgan 2021/10/22
               '法律所自行收文(非案源),先考慮L案其他要看實際狀況再加
               ElseIf strA1P01 = "L" And .Fields("cp01") = "L" And strLOS02 = "" Then
               
                  '25%案源:10%領現(2121應付費用),15%薪點(收入)
                  '110/10/22跟婉莘確認過用總額計算(不減出庭費)
                  'Removed by Morgan 2025/4/28  律所自行收文案件取消律所分潤--婉莘
                  'strA1P30 = ""
                  'lngAmt = Round(lngAmtR * 0.1)
                  'strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                  'strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                  '   " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '2121', '" & IIf(Left(strR, 4) = "2201", MsgText(55), strDept) & "', 0, " & lngAmt & ", null, null, null, null, null, '" & ChgSQL(strRemarkR & "/案源") & "', '" & strCompanyNo & "', '" & strManNo & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                  'adoTaie.Execute strSql, intI
                  '
                  'lngAmt = Round(lngAmtR * 0.25) - lngAmt
                  'strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                  'strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                  '   " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strR & "', '" & IIf(Left(strR, 4) = "2201", MsgText(55), strDept) & "', 0, " & lngAmt & ", null, null, null, null, null, '" & ChgSQL(strRemarkR & "/案源") & "', '" & strCompanyNo & "', '" & strManNo & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                  'adoTaie.Execute strSql, intI
                  '
                  'lngAmtR = lngAmtR - Round(lngAmtR * 0.25)
                  
                  '律師庭費(目前採固定金額,不考慮20%的選項)
                  '規費科目設220113的是要扣律師庭費,實際規費科目用2403
'Removed by Morgan 2025/7/17 移到外層(上面)，改不管是否案源，有輸出庭費的都要扣--秀玲
'                  If strF = "220113" And bolLawFeeDone = False Then
'                     If strLawNo = "" Then
'                        stNoLawyerAlert = stNoLawyerAlert & adocheck.Fields("a0j02") & "(" & adocheck.Fields("a0j01") & ")" & vbCrLf
'                     End If
'                     'Added by Morgan 2023/7/19 出庭費改先抓caselawer的金額,沒有才走舊規則
'                     strExc(0) = "select cl01,cl02,cl03,st02 from caselawer,staff" & _
'                        " where cl01='" & adocheck.Fields("a0j01") & "' and cl03>0 and st01(+)=cl02 order by cl03 desc,cl02"
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                     If intI = 1 Then
'                        strLawNo = RsTemp("cl02")
'                        strLawyerNameC = RsTemp("st02")
'                        lngAmt = RsTemp("cl03")
'                     Else
'                        strLawyerNameC = .Fields("st02")
'                     'end 2023/7/19
'                        lngAmt = GetLawFee(strLawNo, Adodc1.Recordset.Fields("t0208"), .Fields("cp10"))
'                     End If
'
'                     strXFeeCaseList = strXFeeCaseList & vbCrLf & strCaseNo
'                     If bolRFeeClean = True And lngAmtR >= lngAmt Then
'                        strA1P30 = strLawNo
'                        'Modified by Morgan 2023/7/18 摘要的律師只需帶扣出庭費的律師
'                        strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
'                        strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
'                           " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '220113', 'TOT', 0, " & lngAmt & ", null, null, null, null, null, '" & ChgSQL(Replace(strRemarkR, strLawyerName, "") & "/" & strLawyerNameC) & "', '" & strCompanyNo & "', '" & strA1P30 & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
'                        adoTaie.Execute strSql, intI
'
'                        strXFeeCaseList = strXFeeCaseList & " (" & IIf(strLawNo = "", "未分案", "" & .Fields("st02")) & ":" & lngAmt & ")"
'                        lngAmtR = lngAmtR - lngAmt
'
'                        '多律師出庭
'                        'Modified by Morgan 2023/7/19
'                        'strExc(0) = "select distinct cl02,st02 from caseprogress,caselawer,staff" & _
'                           " where cp162='" & .Fields("cp162") & "' and cp01='" & .Fields("cp01") & "'" & _
'                           " and cl01(+)=cp09 and cl02<>'" & strLawNo & "' and st01(+)=cl02"
'                        strExc(0) = "select cl01,cl02,cl03,st02 from caselawer,staff" & _
'                           " where cl01='" & adocheck.Fields("a0j01") & "' and cl02<>'" & strLawNo & "' and st01(+)=cl02 order by cl03 desc,cl02"
'                        'end 2023/7/19
'                        intI = 1
'                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                        If intI = 1 Then
'                           Do While Not RsTemp.EOF
'                              strA1P30 = RsTemp("cl02")
'                              'Modified by Morgan 2023/7/19
'                              'lngAmt = GetLawFee(strA1P30, Adodc1.Recordset.Fields("t0208"), .Fields("cp10"))
'                              strLawyerNameC = RsTemp("st02")
'                              If IsNull(RsTemp("cl03")) Then
'                                 lngAmt = GetLawFee(strA1P30, Adodc1.Recordset.Fields("t0208"), .Fields("cp10"))
'                              Else
'                                 lngAmt = RsTemp("cl03")
'                              End If
'                              'end 2023/7/19
'                              strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
'                              'Modified by Morgan 2023/7/18 摘要的律師只需帶扣出庭費的律師
'                              strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
'                                 " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '220113', 'TOT', 0, " & lngAmt & ", null, null, null, null, null, '" & ChgSQL(Replace(strRemarkR, strLawyerName, "") & "/" & strLawyerNameC) & "', '" & strCompanyNo & "', '" & strA1P30 & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
'                              adoTaie.Execute strSql, intI
'
'                              strXFeeCaseList = strXFeeCaseList & " (" & RsTemp("st02") & ":" & lngAmt & ")"
'                              lngAmtR = lngAmtR - lngAmt
'
'                              RsTemp.MoveNext
'                           Loop
'                        End If
'                        bolLawFeeDone = True
'
'                     Else
'                        strXFeeCaseList = strXFeeCaseList & " (本次未扣)"
'                     End If
'                  End If
'end 2025/7/17
                                    
                  strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                  strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                     " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strR & "', '" & IIf(Left(strR, 4) = "2201", MsgText(55), strDept) & "', 0, " & lngAmtR & ", null, null, null, null, null, '" & ChgSQL(strRemarkR) & "', '" & strCompanyNo & "', 'L0100', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                  adoTaie.Execute strSql, intI
                     
               Else
               'end 2021/1/15
                  
                  'Added by Morgan 2021/10/22
                  '規費科目設220113的是要扣律師庭費,實際規費科目用2403
                  If strF = "220113" Then
                     strF = "2403"
                  End If
                  'end 2021/10/22
                  
                  'Added by Morgan 2012/7/12
                  If strR = "414101" Or strR = "416101" Or strR = "416102" Then
                     strA1P30 = strLawNo
                  Else
                     strA1P30 = ""
                  End If
                  'end 2012/7/12
                  
                  'Modified by Morgan 2013/12/19 收款會有J公司,取消 a1p01='1' 條件
                  strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                  'ADD BY SONIA 2016/1/4 105年起法務收入改其他部門收入(傳CP09以判斷案件性質及收文人員)
                  'Modified by Morgan 2021/8/9 取消，L公司的收入及規費科目，一律抓案件性質檔的設定--秀玲
                  'If Val(FCDate(strDate)) >= 1050101 And (Left(strR, 4) = "4141" Or Left(strR, 4) = "4161" Or Left(strR, 4) = "4181") Then
                  '   InsertLawACC1P0 strA1P01, "A", strSerialNo, Text2, strR, strDept, 0, Val(lngAmtR), "", "", "", "", "", strRemarkR, strCompanyNo, strManNo, strCaseNo, Val(FCDate(strDate)), "", "", "", IIf(stra1p22 = "null", "", Replace(stra1p22, "'", "")), "", "", "", strYes, IIf(stra1p27 = "null", "", Replace(stra1p27, "'", "")), strA1P30, "", .Fields("a0j01")
                  'Else
                  'end 2021/8/9
                  'END 2016/1/4
                     
                     strA1P16 = strManNo 'Added by Morgan 2023/3/15
                     'Added by Morgan 2021/7/28
                     'ACS之智財顧問112收款，以總收文號讀取智財顧問專業分配比例檔 ACSPFRate，抓出各部門分配比例，
                     '依比例分配至各部門，科目都做原ACS科目但部門改為各專業部門
                     If .Fields("cp01") = "ACS" Then
                        If .Fields("cp10") = "112" Then
                           'Modified by Lydia 2023/11/29 人工調整分配比例(調整比例)
                           'strExc(0) = "select ar02,ar03 from ACSPFRate where ar01='" & .Fields("cp09") & "' and ar03>0 and ar02<>'ACS'"
                           strExc(0) = "select ar02,decode(nvl(ar10,0),0,ar03,ar10) as ar03 from ACSPFRate where ar01='" & .Fields("cp09") & "' and decode(nvl(ar10,0),0,ar03,ar10)>0 and ar02<>'ACS' "
                           intI = 1
                           Set adoaccsum = ClsLawReadRstMsg(intI, strExc(0))
                           If intI = 1 Then
                              'Modified by Morgan 2023/10/6 20%行政流程費用歸顧服組,80%分配
                              'lngTemp = lngAmtR 'Added by Morgan 2023/6/19
                              lngTemp = 0.8 * lngAmtR
                              'end 2023/10/6
                              Do While Not adoaccsum.EOF
                                 'lngTemp = lngAmtR 'Removed by Morgan 2023/6/19
                                 lngAmt = Trunc(lngTemp * adoaccsum("ar03") / 100)
                                 If lngAmt > 0 Then
                                    'Modified by Morgan 2023/3/15 strManNo->strA1P16
                                    strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                                       " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strR & "', '" & adoaccsum("ar02") & "', 0, " & lngAmt & ", null, null, null, null, null, '" & ChgSQL(strRemarkR) & "', '" & strCompanyNo & "', '" & strA1P16 & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                                    adoTaie.Execute strSql, intI
                                    
                                    lngAmtR = lngAmtR - lngAmt
                                    
                                    If lngAmtR > 0 Then
                                       strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                                    End If
                                 End If
                                 adoaccsum.MoveNext
                              Loop
                           End If
                        
                        
                        'Added by Morgan 2023/3/28
                        '後金(702)要做智權人員與顧服組W2001的3:7分點數 --
                        ElseIf .Fields("cp10") = "702" And strManNo <> "W2001" Then
                           'Modified by Morgan 2025/11/3 摘要也要帶智權--珮瑄/秀玲
                           'strRemark = "顧/" & strCompany & "/" & strProperty
                           strRemark = strRemarkR                     '
                           'end 2025/11/3
                           lngAmt = Round(lngAmtR * 0.7)
                           'Modified by Morgan 2025/9/16 --教威/婉莘
                           '業務點數全部歸智權人員，獎金點數智權人員：顧服組為3：7 Ex:F11401138
                           'strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                              " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strR & "', 'W', 0, " & lngAmt & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & strCompanyNo & "', 'W2001', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                           strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                              " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strR & "', 'W', 0, " & lngAmt & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & strCompanyNo & "', '" & strManNo & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'顧服獎金')"
                           adoTaie.Execute strSql, intI
                           
                           lngAmtR = lngAmtR - lngAmt
                           strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                        'Added by Morgan 2023/3/15
                        Else
                           'Modified by Morgan 2023/3/28 改抓特殊設定
                           'strExc(0) = "select cp09 from caseprogress where cp01 = '" & .Fields("cp01") & "' and cp02 = '" & .Fields("cp02") & "' and cp03 = '" & .Fields("cp03") & "' and cp04 = '" & .Fields("cp04") & "' and substr(cp10,1,3) in ('101','103','104','105','106','112','113','121','122','124')"
                           strExc(1) = Pub_GetSpecMan("ACS-C")
                           strExc(0) = "select cp09 from caseprogress where cp01 = '" & .Fields("cp01") & "' and cp02 = '" & .Fields("cp02") & "' and cp03 = '" & .Fields("cp03") & "' and cp04 = '" & .Fields("cp04") & "' and cp10 in ('" & Replace(strExc(1), ";", "','") & "')"
                           'end 2023/3/28
                           intI = 1
                           Set adoaccsum = ClsLawReadRstMsg(intI, strExc(0))
                           If intI = 1 Then
                              'Modified by Morgan 2023/3/28 客戶名稱只抓6個字
                              'If Left(strRemarkR, Len(strMan) + 1) = strMan & "/" Then
                              '   strRemarkR = "專案" & Mid(strRemarkR, Len(strMan) + 1)
                              'End If
                              'strRemarkR = strRemarkR & "/" & strCaseNo & "/" & strManNo 'Added by Morgan 2023/3/23 +本所案號/智權人員編號
                              strRemarkR = "專案" & "/" & Left(strCompany, 6) & "/" & strProperty & "/" & strCaseNo & "/" & strManNo
                              'end 2023/3/28
                              strA1P16 = "M0101"
                              strA1P30 = strManNo
                              
                              'Added by Morgan 2025/4/23
                              '以本所案號讀取ACS_TIPS_Rate的本所案號，依比例拆兩筆，後面的比率要加上"顧服獎金"在對沖其他欄位 'Modified by Morgan 2025/4/28
                              strExc(0) = "select atr08 from ACS_TIPS_Rate where atr01='" & .Fields("cp01") & "' and atr02='" & .Fields("cp02") & "' and atr03='" & .Fields("cp03") & "' and atr04='" & .Fields("cp04") & "' and atr05='1' and atr08>0"
                              intI = 1
                              Set adoaccsum = ClsLawReadRstMsg(intI, strExc(0))
                              If intI = 1 Then
                                 strA1P16 = strManNo
                                 strRemarkR = PUB_GetShortName(strManNo) & "/" & Left(strCompany, 6) & "/" & strProperty & "/" & strCaseNo
                                 lngAmt = Round(lngAmtR * 0.01 * adoaccsum(0))
                                 strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27)" & _
                                    " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strR & "', 'W', 0, " & lngAmt & ", null, null, null, null, null, '" & ChgSQL(strRemarkR) & "', '" & strCompanyNo & "', '" & strManNo & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ")"
                                 adoTaie.Execute strSql, intI
                                 
                                 strA1P30 = "顧服獎金"
                                 lngAmtR = lngAmtR - lngAmt
                                 strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                              End If
                              'end 2025/4/23
                              
                           End If
                        'end 2023/3/15
                        End If
                     End If
                     'end 2021/7/28
                     
                     'Added by Morgan 2023/8/7
                     '顧服組客戶移交專利國內部及商標部處理
                     '顧服組移交客戶是指智權人員為P1004專利智權人員或P2006商標智權人員
                     '1. P1004：若承辦人為工程師，則承辦人與系統特殊設定「P1004業績人員」(郭雅娟)各一半，若承辦人非工程師則全部列「P1004業績人員」；
                     'Modified by Morgan 2024/8/6 +30015比照P1004:承辦人與系統特殊設定「30015業績人員」(創新智權專利)各一半，若承辦人非工程師則全部列「30015業績人員」；
                     If (strManNo = "P1004" Or strManNo = "30015") And lngAmtR > 0 Then
                        '工程師
                        If Left(.Fields("st03"), 2) = "P1" And .Fields("st03") <> "P12" Then
                           strA1P16 = .Fields("cp14")
                           If .Fields("a0j04") = "000" Then
                              strRemark = PUB_GetShortName(strA1P16) & "/" & strCompany & "/" & strProperty
                           Else
                              strRemark = PUB_GetShortName(strA1P16) & "/" & strCompany & "/" & strNation & "/" & strProperty
                           End If
                           lngAmt = Round(lngAmtR * 0.5)
                           strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                              " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strR & "', '" & strDept & "', 0, " & lngAmt & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & strCompanyNo & "', '" & strA1P16 & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                           adoTaie.Execute strSql, intI
                           
                           lngAmtR = lngAmtR - lngAmt
                           strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                        End If
                        
                        'Added by Morgan 2024/8/6
                        If strManNo = "30015" Then
                           strA1P16 = Pub_GetSpecMan("30015業績人員")
                        Else
                        'end 2024/8/6
                        
                           'P1004業績人員
                           strA1P16 = Pub_GetSpecMan("P1004業績人員")
                        End If
                        If .Fields("a0j04") = "000" Then
                           strRemark = PUB_GetShortName(strA1P16) & "/" & strCompany & "/" & strProperty
                        Else
                           strRemark = PUB_GetShortName(strA1P16) & "/" & strCompany & "/" & strNation & "/" & strProperty
                        End If
                        lngAmt = lngAmtR
                        strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                           " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strR & "', '" & strDept & "', 0, " & lngAmt & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & strCompanyNo & "', '" & strA1P16 & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                        adoTaie.Execute strSql, intI
                        lngAmtR = 0
                     
                     '2. P2006：20%列顧服組W2001，剩餘80%，CFT案列「P2006業績CFT案人員」陳蒲璇、T案列「P2006業績T案人員」桂紹禎。
                     ElseIf strManNo = "P2006" And lngAmtR > 0 Then
                        If .Fields("cp01") = "CFT" Then
                           strA1P16 = Pub_GetSpecMan("P2006業績CFT案人員")
                        Else
                           strA1P16 = Pub_GetSpecMan("P2006業績T案人員")
                        End If
                        If .Fields("a0j04") = "000" Then
                           strRemark = PUB_GetShortName(strA1P16) & "/" & strCompany & "/" & strProperty
                        Else
                           strRemark = PUB_GetShortName(strA1P16) & "/" & strCompany & "/" & strNation & "/" & strProperty
                        End If
                        lngAmt = Round(lngAmtR * 0.8)
                        strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                           " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strR & "', '" & strDept & "', 0, " & lngAmt & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & strCompanyNo & "', '" & strA1P16 & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                        adoTaie.Execute strSql, intI
                        lngAmtR = lngAmtR - lngAmt
                        strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                        
                        strA1P16 = "W2001"
                        If .Fields("a0j04") = "000" Then
                           strRemark = PUB_GetShortName(strA1P16) & "/" & strCompany & "/" & strProperty
                        Else
                           strRemark = PUB_GetShortName(strA1P16) & "/" & strCompany & "/" & strNation & "/" & strProperty
                        End If
                        lngAmt = lngAmtR
                        strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                           " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strR & "', 'W', 0, " & lngAmt & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & strCompanyNo & "', '" & strA1P16 & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                        adoTaie.Execute strSql, intI
                        lngAmtR = 0
                     End If
                     'end 2023/8/7
                     
                     If lngAmtR > 0 Then 'Added by Morgan 2022/11/21 ACS可能會全部做專業部門 Ex:F11109981(E11119766)
                        'Added by Morgan 2023/10/18 部分收款且未發文時改作 2492 點數保留
                        strA1P23 = ""
                        If m_bolRcptClear = False And adocheck("cp158") = 0 Then
                           strA1P23 = adocheck("a1u03") & adocheck("a1u02") '收文號+收據號
                           strR = "2492"
                        
                        '有部分收款時將 2492 點數保留轉收入(規費先沖銷,後收款的一定是服務費)
                        Else
                           strExc(0) = "select sum(a1p08)-sum(a1p07) rp from acc1p0" & _
                              " where a1p23='" & adocheck("a1u03") & adocheck("a1u02") & "' and a1p05='2492' and a1p04<>'" & Text2 & "'"
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                           If intI = 1 Then
                              If RsTemp(0) > 0 Then
                                 
                                 strA1P23 = adocheck("a1u03") & adocheck("a1u02") '收文號+收據號
                                 strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                                    " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '2492', '" & strDept & "', " & RsTemp(0) & ", 0, null, null, null, null, null, '" & ChgSQL(strRemarkR) & "', '" & strCompanyNo & "', '" & strA1P16 & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", '" & strA1P23 & "', null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                                 adoTaie.Execute strSql, intI
                                 
                                 strA1P23 = ""
                                 lngAmtR = lngAmtR + RsTemp(0)
                                 strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                              End If
                           End If
                        End If
                        'end 2023/10/18
                        
                        'modify by sonia 2016/1/29 2201xx規費科目的部門一律改用TOT
                        'Modified by Morgan 2023/3/15 strManNo->strA1P16
                        'Modified by Morgan 2023/9/25 +a1p23
                        strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
                           " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strR & "', '" & IIf(Left(strR, 4) = "2201", MsgText(55), strDept) & "', 0, " & lngAmtR & ", null, null, null, null, null, '" & ChgSQL(strRemarkR) & "', '" & strCompanyNo & "', '" & strA1P16 & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", '" & strA1P23 & "', null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                        adoTaie.Execute strSql, intI
                        
                        strA1P23 = "" 'Added by Morgan 2023/10/18
                     End If
                     
                  'End If   'ADD BY SONIA 2016/1/4 'Removed by Morgan 2021/8/9
                  
               End If 'Added by Morgan 2021/1/15
            End If
            
         'Added by Morgan 2024/12/27
         '扣點數時改借方並加註於摘要
         ElseIf lngAmtR < 0 Then
            strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
            strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30)" & _
               " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strR & "', '" & IIf(Left(strR, 4) = "2201", MsgText(55), strDept) & "', " & Abs(lngAmtR) & ", 0, null, null, null, null, null, '" & ChgSQL(strRemarkR) & "扣點數" & "', '" & strCompanyNo & "', '" & strA1P16 & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", '" & strA1P23 & "', null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
            adoTaie.Execute strSql, intI
            strA1P23 = ""
         'end 2024/12/27
         End If
         
         '規費
         If lngAmtF > 0 Then
            'Modified by Morgan 2025/8/19 ACS案跨期後收款(未收款已沖帳)，規費(稅款)不用產生分錄--瑞婷
            'If bolAssignDone = False Then
            If bolAssignDone = False And bolACSNoTaxItem = False Then
            'end 2025/8/19
            
               'Added by Morgan 2021/1/20 案源法律所收據
               If strA1P01 = "L" And strLOS02 <> "" Then
                  strA1P16 = "L0100"
                  strF = "2403" '代收代付款
                  strRemarkF = Left(strCompany, 6) & "/" & strProperty & "規費"
                  strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                  strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30) " & _
                        "values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strF & "', 'TOT', 0, " & lngAmtF & ", null, null, null, null, null, '" & ChgSQL(strRemarkF) & "', '" & strCompanyNo & "', '" & strA1P16 & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                  adoTaie.Execute strSql, intI
               'Added by Morgan 2021/3/9
               ElseIf strA1P01 = "1" And (strLOS02 = "A1" Or strLOS02 = "A2") Then
                  strA1P16 = strManNo
                  strF = "2123" '應付佣金
                  strRemarkF = "法律所/" & strMan & "/" & MidB(strLRcpTitle, 1, 16) & "/" & strLCaseNo & "/" & strProperty
                  strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                  strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30) " & _
                        "values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strF & "', 'TOT', 0, " & lngAmtF & ", null, null, null, null, null, '" & ChgSQL(strRemarkF) & "', '" & strCompanyNo & "', '" & strA1P16 & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                  adoTaie.Execute strSql, intI
               'end 2021/3/9
               Else
               'end 2021/1/20
               
                  strA1P16 = strManNo
                  'Add by Morgan 2011/10/7
                  If bolLawyerGuei Then
                     strA1P16 = "M0100"
                     strF = "414101"
                     'add by sonia 2019/5/27 摘要智權人員要改 F
                     strRemarkF = "總/" & Mid(strRemarkF, 3)
                     'end 2019/5/27
                  End If
                  'Added by Morgan 2012/7/12
                  If strF = "414101" Or strF = "416101" Or strF = "416102" Or strF = "220113" Then
                     strA1P30 = strLawNo
                  Else
                     strA1P30 = ""
                  End If
                  'end 2012/7/12
                  'Modified by Morgan 2013/12/19 收款會有J公司,取消 a1p01='1' 條件
                  strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
                  
                  'ADD BY SONIA 2016/1/4 105年起法務收入改其他部門收入(傳CP09以判斷案件性質及收文人員)
                  'Modified by Morgan 2021/8/9 取消，L公司的收入及規費科目，一律抓案件性質檔的設定--秀玲
                  'If Val(FCDate(strDate)) >= 1050101 And (Left(strF, 4) = "4141" Or Left(strF, 4) = "4161" Or Left(strF, 4) = "4181") Then
                  '   InsertLawACC1P0 strA1P01, "A", strSerialNo, Text2, strF, strDept, 0, Val(lngAmtF), "", "", "", "", "", ChgSQL(strRemarkF), strCompanyNo, strA1P16, strCaseNo, Val(FCDate(strDate)), "", "", "", IIf(stra1p22 = "null", "", Replace(stra1p22, "'", "")), "", "", "", strYes, IIf(stra1p27 = "null", "", Replace(stra1p27, "'", "")), strA1P30, "", .Fields("a0j01")
                  'Else
                  'end 2021/8/9
                  'END 2016/1/4
                  
                     '2014/2/11 modify by sonia E10228435(F10300038第三項次)部門應仍為L,且收款已不使用610103
                     'strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30) " & _
                        "values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strF & "', '" & IIf(strF = "610103", strDept, MsgText(55)) & "', 0, " & lngAmtF & ", null, null, null, null, null, '" & strRemarkF & "', '" & strCompanyNo & "', '" & strA1p16 & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1p30 & "')"
                     'modify by sonia 2016/1/29 2201xx規費科目的部門一律改用TOT
                     'modify by sonia 2020/8/26 A1及A2案源之智慧所收據之規費,科目改用2123應付佣金E10917870
                     'strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30) " & _
                     '   "values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strF & "', '" & IIf(Left(strF, 4) = "2201", MsgText(55), strDept) & "', 0, " & lngAmtF & ", null, null, null, null, null, '" & ChgSQL(strRemarkF) & "', '" & strCompanyNo & "', '" & strA1P16 & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                     strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30) " & _
                        "values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & IIf(strF = "", "2123", strF) & "', '" & IIf(Left(strF, 4) = "2201", MsgText(55), strDept) & "', 0, " & lngAmtF & ", null, null, null, null, null, '" & ChgSQL(strRemarkF) & "', '" & strCompanyNo & "', '" & strA1P16 & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ",'" & strA1P30 & "')"
                     '2014/2/11 END
                     adoTaie.Execute strSql, intI
                     
                  'End If   'ADD BY SONIA 2016/1/4 'Removed by Morgan 2021/8/9
                  
                  strFeeItemNo = strSerialNo 'Added by Morgan 2014/2/12
                  
                  'Removed by Morgan 2014/2/12 改依案號新增銷項稅額
                  
               End If 'Added by Morgan 2021/1/20
            End If
         End If
         
         '出庭費(不可同時有收文分配;目前分配為 L 而出庭費為 P,FCP,T,FCT 故也不會發生)
         If bolXFeeOK = True Then
            'Modified by Morgan 2013/12/19 收款會有J公司,取消 a1p01='1' 條件
            strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
            strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27)" & _
               " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strF & "', '" & MsgText(55) & "', 0, 10000, null, null, null, null, null, '" & strProperty & "/出庭費', '" & strCompanyNo & "', '" & strManNo & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & ")"
            adoTaie.Execute strSql, intI
         End If
         
         '支援
         bol2SupportOK = False
         If CheckSupport(.Fields("cp09")) = True Then
            '判斷相關案都沒扣過點數
            If PUB_ChkP4SH(.Fields("cp09")) = False Then
               strSupportCaseList = strSupportCaseList & vbCrLf & .Fields("cp01") & "-" & .Fields("cp02") & "-" & .Fields("cp03") & "-" & .Fields("cp04")
               'Modified by Morgan 2024/1/3 作 2492 點數保留時不扣支援
               If lngAmtR >= 5000 And strR <> "2492" Then
                  bol2SupportOK = True
                  '借方    '2020/2/14 /點作轉專業改為/專業支援
                  strRemark = strMan & "/" & strCompany & "/專業支援"
                  'Modified by Morgan 2013/12/19 收款會有J公司,取消 a1p01='1' 條件
                  'modify by sonia 2016/1/29 2201xx規費科目的部門一律改用TOT
                  'modify by sonia 2022/1/22 Adodc1.Recordset.Fields("t0212")改用strManNo
                  'strSQLD = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) select '" & strA1P01 & "', 'A', substr(1001+max(a1p03),2), '" & Text2 & "', '" & strR & "', '" & IIf(strDept = "", MsgText(55), IIf(Left(strR, 4) = "2201", MsgText(55), strDept)) & "', 5000, 0, null, null, null, null, null, '" & strRemark & "', '" & strCompanyNo & "', '" & Adodc1.Recordset.Fields("t0212") & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & " from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'"
                  'Modified by Morgan 2025/8/1 摘要+chgsql,會有單引號(F11406032)
                  strSQLD = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) select '" & strA1P01 & "', 'A', substr(1001+max(a1p03),2), '" & Text2 & "', '" & strR & "', '" & IIf(strDept = "", MsgText(55), IIf(Left(strR, 4) = "2201", MsgText(55), strDept)) & "', 5000, 0, null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & strCompanyNo & "', '" & strManNo & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & " from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'"
                  '貸方    '2020/2/14 /點作轉專業改為/專業支援
                  strRemark = "專/" & strCompany & "/專業支援"
                  'Modified by Morgan 2013/12/19 收款會有J公司,取消 a1p01='1' 條件
                  'modify by sonia 2016/1/29 2201xx規費科目的部門一律改用TOT
                  strSQLc = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) select '" & strA1P01 & "', 'A', substr(1001+max(a1p03),2), '" & Text2 & "', '" & strR & "', '" & IIf(strDept = "", MsgText(55), IIf(Left(strR, 4) = "2201", MsgText(55), strDept)) & "', 0, 5000, null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & strCompanyNo & "', 'P1001', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, " & CNULL(strYes) & ", " & stra1p27 & " from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'"
               Else
                  strSupportCaseList = strSupportCaseList & "(本次未扣)"
               End If
            End If
         End If
      
         '轉專業點數科目放在規費科目後面
         If bol2SupportOK = True Then
            adoTaie.Execute strSQLD
            strSQLD = ""
            adoTaie.Execute strSQLc
            strSQLc = ""
         End If
         
         'Added by Morgan 2014/2/12
         'Modified by Morgan 2021/7/26 ACS的706代收代付除外(不開發票)
         'Modified by Morgan 2025/8/19 ACS案跨期後收款(未收款已沖帳)，規費(稅款)不用產生分錄--瑞婷
         'If Adodc1.Recordset.Fields("t0214").Value = "J" And Not (.Fields("cp01") = "ACS" And .Fields("cp10") = "706") Then
         If Adodc1.Recordset.Fields("t0214").Value = "J" And Not (.Fields("cp01") = "ACS" And .Fields("cp10") = "706") And bolACSNoTaxItem = False Then
         'end 2025/8/19
            bolTaxDone = AddTaxItem(.Fields("a0j13"), .Fields("a0j01"), .Fields("a0j02"), Val("" & .Fields("a1u05")), strFeeItemNo)
         End If
         'end 2014/2/12
         
         .MoveNext
      Loop
         
      '扣繳額
      If Adodc1.Recordset.Fields("t0206") <> 0 Then
         'Added by Morgan 2021/1/18 案源收款，L公司的業務對沖固定放 L0100，摘要不必放業務 --辜
         strA1P30 = ""
         If strA1P01 = "L" And strLOS02 <> "" Then
            strRemarkTax = Left(strCompany, 6) & "/" & strProperty
            strManNo = "L0100"
         ElseIf strA1P01 = "1" And (strLOS02 = "A1" Or strLOS02 = "A2") Then
            strRemarkTax = "法律所/" & strMan & "/" & Left(strLRcpTitle, 6) & "/" & strLCaseNo & "/" & strProperty
         'Added by Morgan 2021/5/5
         ElseIf strA1P01 = "L" And Left(strCaseNo, 2) = "LA" And strLOS02 = "" Then
            strRemarkTax = strRemarkTax & "所得稅"
            strA1P30 = Adodc1.Recordset.Fields("t0207") & "法律稅"
         'end 2021/5/5
         End If
         'end 2021/1/18
         'Modified by Morgan 2013/12/19 收款會有J公司,取消 a1p01='1' 條件
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
         adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27, a1p30) values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '1203', '" & MsgText(55) & "', " & Val(Adodc1.Recordset.Fields("t0206")) & ", 0, null, null, null, null, null, '" & strRemarkTax & "', '" & strCompanyNo & "', '" & strManNo & "', '" & strCaseNo & "', " & Val(FCDate(strDate)) & ", null, null, null, " & stra1p22 & ", null, null, null, null, " & stra1p27 & ",'" & strA1P30 & "')"
      End If
      
      'Added by Morgan 2021/3/9
      'A1,A2類案源
      'Modified by Morgan 2023/11/20 改在TT案收據時新增，否則法律案拆收據時會重複
      'If strA1P01 = "L" And (strLOS02 = "A1" Or strLOS02 = "A2") Then
      If Left(strCaseNo, 2) = "TT" And (strLOS02 = "A1" Or strLOS02 = "A2") Then
         '借 勞務費6129
         strExc(2) = "智慧所/" & strTTManSN & "/" & Left(strLRcpTitle, 6) & "/" & strLCaseNo & "/法務案件諮詢"
         lngAmt = lngTTAmt
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
         strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p14, a1p15, a1p16, a1p18, a1p22, a1p27)" & _
            " values ('L', 'A', '" & strSerialNo & "', '" & Text2 & "', '6129', 'L', " & lngAmt & ", 0, '" & ChgSQL(strExc(2)) & "', 'X82357000', 'L0100', " & Val(FCDate(strDate)) & ", " & strA1P22_L & ", " & stra1p27 & ")"
         adoTaie.Execute strSql, intI
         
         '貸 現金1101
         lngAmt = lngTTAmt - 0.1 * lngTTAmt
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
         'Modified by Morgan 2023/8/17 現金1101科目改法律所110502
         strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p14, a1p15, a1p16, a1p18, a1p22, a1p27)" & _
            " values ('L', 'A', '" & strSerialNo & "', '" & Text2 & "', '110502', 'TOT', 0, " & lngAmt & ", '" & ChgSQL(strExc(2)) & "', 'X82357000', 'L0100', " & Val(FCDate(strDate)) & ", " & strA1P22_L & ", " & stra1p27 & ")"
         adoTaie.Execute strSql, intI
         
         '貸 暫收款2401
         lngAmt = lngTTAmt * 0.1
         If lngAmt > 0 Then
            strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
            'Modified by Morgan 2023/8/17 2401->2409
            strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p14, a1p15, a1p16, a1p18, a1p22, a1p27,a1p30)" & _
               " values ('L', 'A', '" & strSerialNo & "', '" & Text2 & "', '2409', 'TOT', 0, " & lngAmt & ", '" & ChgSQL(strExc(2) & "所得稅") & "', 'X82357000', 'L0100', " & Val(FCDate(strDate)) & ", " & strA1P22_L & ", " & stra1p27 & ",'" & Adodc1.Recordset.Fields("t0207") & "法律稅')"
            adoTaie.Execute strSql, intI
         End If
      End If
      'end 2021/3/9
      
      'Added by Morgan 2014/2/21 收據結清時產生銷項稅額科目
      'Removed by Morgan 2014/2/12 改依案號新增銷項稅額
         
      End With
   End If
   
   If adoaccsum.State <> adStateClosed Then adoaccsum.Close
   If adocheck.State <> adStateClosed Then adocheck.Close
   Exit Sub
   
Checking:
   
   MsgBox Err.Description, vbCritical 'Added by Morgan 2019/4/22
   'Resume
   If adoaccsum.State = adStateOpen Then adoaccsum.Close
   If adocheck.State = adStateOpen Then adocheck.Close
   
End Sub

'Modify By Sindy 2015/8/6 已改搬到aacc_fun變共用func,名稱為PUB_ChkIsPerson
'Private Function ChkIsPerson(pNo As String) As Boolean
'   strExc(0) = "select a0k05,a0k11 from acc0k0 where a0k01='" & pNo & "'"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      If RsTemp("a0k05") = "1" Then
'         MsgBox "本收據屬個人不能扣繳!!", vbExclamation, "扣繳檢查"
'         ChkIsPerson = True
'      'Added by Morgan 2013/12/26
'      ElseIf RsTemp("a0k11") = "J" Then
'         MsgBox "智權公司不能扣繳!!", vbExclamation, "扣繳檢查"
'         ChkIsPerson = True
'      'end 2013/12/26
'      End If
'   End If
'End Function

'Added by Morgan 2013/12/27
'檢查智權繳款明細與實際收款是否有差異
Private Sub BatchCheck()
   Dim stSQL As String, intR As Integer
   Dim rsQuery  As ADODB.Recordset
   
   bolDetailChangeMail = False
   
   stSQL = "select * from acc440 where a4416='" & Text2 & "'"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      'stSQL = "select * from acc1p0 where a1p04='" & Text2 & "'"
      'intR = 1
      'Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
      'If intR = 1 Then
     
      'End If
      
      If IsNull(rsQuery("A4424")) Then
         stSQL = "select axd04,axd05 from acc440,acc441 where a4416='" & Text2 & "' and axd01(+)=a4401 and axd02(+)=a4402 and axd03(+)=a4403" & _
            " and not exists(select * from acc1u0 where a1u01=a4416 and a1u02=axd04 and a1u03=axd05 and a1u04=axd06 and a1u05=axd07 and a1u06=axd08)" & _
            " union select a1u02,a1u03 from acc1u0 where a1u01='" & Text2 & "' and not exists(select * from acc440,acc441 where a4416=a1u01" & _
            " and axd01(+)=a4401 and axd02(+)=a4402 and axd03(+)=a4403 and axd04=a1u02 and axd05=a1u03 and axd06=a1u04 and axd07=a1u05 and axd08=a1u06)"
         intR = 1
         Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
         If intR = 1 Then
   'Removed by Morgan 2014/1/10 不要更新繳款明細
   '         '刪除舊明細
   '         stSQL = "delete acc441 where (axd01,axd02,axd03) in (select a4401,a4402,a4403 from acc440 where a4416='" & Text2 & "' and a4424 is null)"
   '         adoTaie.Execute stSQL, intR
   '         '新增明細
   '         stSQL = "insert into acc441(axd01,axd02,axd03,axd04,axd05,axd06,axd07,axd08) select a4401,a4402,a4403,a1u02,a1u03,a1u04,a1u05,a1u06 from acc440,acc1u0 where a4416='" & Text2 & "' and a4424 is null and a1u01(+)=a4416 and a1u02 is not null"
   '         adoTaie.Execute stSQL, intR
            '更新
            stSQL = "update acc440 set a4424='Y' where a4416='" & Text2 & "' and a4424 is null"
            adoTaie.Execute stSQL, intR
            bolDetailChangeMail = True
         End If
      End If
   End If
   Set rsQuery = Nothing
End Sub

'收據是否已開發票且已沖帳
Private Function CheckInvDone(pNo As String) As Boolean
   Dim stSQL As String, intR As Integer, stSales As String, stDate As String
   Dim rsQuery  As ADODB.Recordset
   stSQL = "select 1 from acc431,acc430 where axc02='" & pNo & "' and a4301(+)=axc01 and a4317 is not null"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      CheckInvDone = True
   End If
   Set rsQuery = Nothing
End Function

'2014/11/28 cancel by sonia 已不用
''發票稅額
'Private Function GetInvTax(pNo As String) As String
'   Dim stSQL As String, intR As Integer, stSales As String, stDate As String
'   Dim rsQuery  As ADODB.Recordset
'   stSQL = "select a4305 from acc431,acc430 where axc02='" & pNo & "' and a4301(+)=axc01"
'   intR = 1
'   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
'   If intR = 1 Then
'      GetInvTax = Val("" & rsQuery.Fields(0))
'   End If
'   Set rsQuery = Nothing
'End Function
''end 2013/12/27
'2014/11/28 end

'Added by Morgan 2014/1/10
'考慮通知內容可能超過4K,不寫暫存直接發送
Private Sub DetailChangeInform()
   Dim stVTB As String, stSQL As String, intR As Integer, stSales As String, stDate As String
   Dim rsQuery  As ADODB.Recordset, stRawDetail As String, stRealDetail As String
   
   If bolDetailChangeMail = True Then
      stSQL = "select a4401,sqldatet(a4402)||' '||sqltime(a4403) Rdate from acc440 where a4416='" & Text2 & "' and a4424='Y'"
      intR = 1
      Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         stSales = rsQuery.Fields(0)
         stDate = rsQuery.Fields(1)
         stVTB = "select '1' C0,axd04,axd05,axd06,axd07,axd08 from acc440,acc441 where a4416='" & Text2 & "' and axd01(+)=a4401 and axd02(+)=a4402 and axd03(+)=a4403" & _
            " union all select '2',a1u02,a1u03,a1u04,a1u05,a1u06 from acc1u0 where a1u01='" & Text2 & "'"
         
         stSQL = "select X.*" & _
            ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
            ",nvl(tm05,nvl(pa05,nvl(lc05,nvl(sp05,hc06)))) 案件名稱" & _
            ",decode(a0j04,'000',cpm03,cpm04) 案件性質" & _
            " from (" & stVTB & ") X,acc0j0,caseprogress,casepropertymap" & _
            ",trademark,patent,lawcase,servicepractice,hirecase" & _
            " where a0j01(+)=axd05 and a0j13(+)=axd04 and cp09(+)=axd05 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
            " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
            " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
            " and lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04" & _
            " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04" & _
            " and hc01(+)=cp01 and hc02(+)=cp02 and hc03(+)=cp03 and hc04(+)=cp04" & _
            " order by 1,2,3"
         
         intR = 1
         Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
         If intR = 1 Then
            With rsQuery
            stRawDetail = "&nbsp;智權繳款明細：" & vbCrLf & _
               "<table border=1 cellspacing=0 width=600>" & _
               "<TR style=""background:#E1E1E1""><TD>收據號</TD><TD>本所案號</TD><TD>案件性質</TD><TD>服務費</TD><TD>規費</TD><TD>扣繳</TD><TD>案件名稱</TD></TR>"
            stRealDetail = "&nbsp;財務收款明細：" & vbCrLf & _
               "<table border=1 cellspacing=0 width=600>" & _
               "<TR style=""background:#E1E1E1""><TD>收據號</TD><TD>本所案號</TD><TD>案件性質</TD><TD>服務費</TD><TD>規費</TD><TD>扣繳</TD><TD>案件名稱</TD></TR>"
            .MoveFirst
            Do While Not .EOF
               If .Fields("C0") = "1" Then
                  stRawDetail = stRawDetail & "<TR><TD>" & .Fields("axd04") & "</TD><TD>" & .Fields("本所案號") & "</TD><TD>" & Left("" & .Fields("案件性質"), 6) & "</TD><TD align=right>" & .Fields("axd06") & "</TD><TD align=right>" & .Fields("axd07") & "</TD><TD align=right>" & .Fields("axd08") & "</TD><TD>" & Left("" & .Fields("案件名稱"), 10) & "</TD></TR>"
               Else
                  stRealDetail = stRealDetail & "<TR><TD>" & .Fields("axd04") & "</TD><TD>" & .Fields("本所案號") & "</TD><TD>" & Left("" & .Fields("案件性質"), 6) & "</TD><TD align=right>" & .Fields("axd06") & "</TD><TD align=right>" & .Fields("axd07") & "</TD><TD align=right>" & .Fields("axd08") & "</TD><TD>" & Left("" & .Fields("案件名稱"), 10) & "</TD></TR>"
               End If
               .MoveNext
            Loop
            stRawDetail = stRawDetail & "</TABLE>"
            stRealDetail = stRealDetail & "</TABLE>"
            
            End With
         End If
         PUB_SendMail strUserNum, stSales, "", "您 " & stDate & " 輸入之繳款明細與實際收款不同!!", stRawDetail & vbCrLf & stRealDetail, , , True
      End If
   End If
   Set rsQuery = Nothing
End Sub
'Added by Morgan 2014/1/13
'是否列印憑證
Private Sub SetProofPrint()
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stAddMemo As String
   
   stSQL = "select * from acc440 where a4416='" & Text2 & "'"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      stSQL = ""
      '溢收退客戶--繳款記錄有溢收退客戶且貸方有暫收
      If rsQuery("A4410") > 0 Then
         If rsQuery("A4425") = "2" Then
            If stSQL <> "" Then stSQL = stSQL & " union "
            stSQL = "select '1' C1 from acc1p0 where a1p04='" & Text2 & "' and a1p05='2401' and a1p08>0 and rownum<2"
            stAddMemo = "/溢付款"
         'Added by Morgan 2023/12/6 列暫收:摘要帶智權備註--瑞婷
         Else
            stAddMemo = "/" & rsQuery("A4412")
         'end 2023/12/6
         End If
         
         'Added by Morgan 2023/10/25
         adoTaie.Execute "update acc1p0 set a1p14=a1p14||'" & stAddMemo & "' where a1p04='" & Text2 & "' and a1p05='2401' and a1p08>0", intR
         'end 2023/10/25
      End If
      '電匯--繳款記錄有電匯且借方有特定銀存科目
      If rsQuery("A4406") > 0 Or rsQuery("A4407") > 0 Then
         If stSQL <> "" Then stSQL = stSQL & " union "
         stSQL = "select '2' C1 from acc1p0 where a1p04='" & Text2 & "' and a1p05 in ('110202','110207','110303','110223','110208','110204','110301','110302','1911','1912','1913') and a1p07>0 and rownum<2"
      End If
      '補扣繳--繳款記錄有補扣繳且借方有暫收
      If rsQuery("A4422") > 0 Then
         If stSQL <> "" Then stSQL = stSQL & " union "
         stSQL = "select '3' C1 from acc1p0 where a1p04='" & Text2 & "' and a1p05='2401' and a1p07>0 and rownum<2"
      End If
      If stSQL <> "" Then
         intR = 1
         Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
         If intR = 1 Then
            stSQL = "update acc0l0 set a0l17='Y' where a0l01='" & Text2 & "'"
            adoTaie.Execute stSQL, intR
         End If
      End If
   End If
   Set rsQuery = Nothing
End Sub

'Added by Morgan 2014/1/23
'檢查收款是否需要扣銷項稅額
'In:pReceiptNo=收據號, pReceiveNo=收文號, pCaseNo=本所案號, pRecFee=本次規費金額, pA1P03=本次規費項次
Private Function AddTaxItem(pReceiptNo As String, pReceiveNo As String, pCaseNo As String, pRecFee As Long, pA1P03 As String) As Boolean
   Dim lngTax As Long '本次扣稅額
   Dim strAccCode As String '稅額科目
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   Dim bolServiceOnly As Boolean
   Dim lngReceiptTax As Long '整張收據稅額
   Dim lngCaseTax As Long '本案稅額
   Dim lngTaxLimit As Long '規費科目總額
   Dim oblClear As Boolean '該收據該案號是否已結清
   Dim strFNo As String '結清的收款單號
   Dim bolDone As Boolean '規費已減扣繳
   Dim stCU178 As String, stA4303 As String   'add by sonia 2019/07/30 零稅率
   Dim stSys As String, stA1P30 As String 'Added by Morgan 2023/5/19
 
   bolServiceOnly = False
   
   'add by sonia 2019/7/30 判斷是否零稅率,零稅率不扣
   stSQL = "select A4301,A4323 from acc430,acc431 where axc02='" & pReceiptNo & "' and axc01=a4301(+)"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      If "" & rsQuery("A4323") = "Y" Then
         Exit Function
      End If
   End If
   'end 2019/7/30
   
   '2015/11/12 改結清時抓該次收款的最大規費最小收文號扣,若不夠則借規費科目扣
   
   '1.一張收據同一案號抓最大規費最小收文號的那一筆來扣稅額
   '2.整張收據的稅額與案號稅額加總不同時,差額調至最大案號
   '3.部分收時借 結清時發票無有沖帳貸不同科目
   
   '該收據該案的銷項稅額
   'Modified by Morgan 2016/3/16 結清收款單號需排除非收款的單號(會有銷退單號),但若收款後做銷帳結清則不可再來修改收款否則會誤判
   stSQL = "select Tot-Sub-round((Tot-Sub)/1.05) Tax,Fee-SubFee NetFee,Tot,Amt,FNo" & _
      " from (select a0j02 x1,sum(nvl(a0j09,0)+nvl(a0j10,0)) Tot,nvl(sum(a0j10),0) Fee from acc0j0 where a0j13='" & pReceiptNo & "' and a0j02='" & pCaseNo & "'" & _
      " group by a0j02) x,(select a0j02 y1,nvl(sum(a1u07),0)+nvl(sum(a1u09),0) Sub,nvl(sum(a1u09),0) SubFee,sum(nvl(a1u04,0)+nvl(a1u05,0)+nvl(a1u07,0)-nvl(a1u08,0)+nvl(a1u09,0)-nvl(a1u10,0)) Amt,max(decode(substr(a1u01,1,1),'F',a1u01)) FNo" & _
      " from acc0j0,acc1u0 where a0j13='" & pReceiptNo & "' and a0j02='" & pCaseNo & "' and a1u02(+)=a0j13 and a1u03(+)=a0j01 group by a0j02) y where y1(+)=x1"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      lngCaseTax = rsQuery(0)
      '該案沒有規費(純服務費)
      If rsQuery(1) = 0 Then
         bolServiceOnly = True
      End If
      '該收據該案號是否已結清
      If rsQuery("Tot") = rsQuery("Amt") Then
         oblClear = True
         strFNo = rsQuery("FNo")
      End If
   End If
   
   If Not oblClear Or strFNo <> Text2 Then Exit Function 'Added by Morgan 2015/11/12

   '是否為該次收款該收據該案號的最大規費的最小收文號
   stSQL = "select nvl(max(a0j10),0)-nvl(sum(a1u09),0) Net,a0j01 from acc0j0,acc1u0" & _
      " where a0j13='" & pReceiptNo & "' and a0j02='" & pCaseNo & "' and a0j01 in (select b.a1u03 from acc1u0 b where b.a1u01='" & Text2 & "')" & _
      " and a1u02(+)=a0j13 and a1u03(+)=a0j01 group by a0j01 order by 1 desc,2 asc"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      If rsQuery("a0j01") <> pReceiveNo Then
         Exit Function
      End If
   Else
      Exit Function
   End If

   
   '整張收據的銷項稅額
   stSQL = "select Net-round(Net/1.05) Tax" & _
      " from (select max(nvl(a0k06,0)+nvl(a0k07,0))-nvl(sum(a1u07),0)-nvl(sum(a1u09),0) Net" & _
      " from acc0k0,acc1u0 where a0k01='" & pReceiptNo & "' and a1u02(+)=a0k01)"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      lngReceiptTax = rsQuery(0)
   End If
   
   '1.收據依案號個別計算銷項稅額後加總 2.調整差額至收據的最大案號
   stSQL = "select Sum(Tot-Sub-round((Tot-Sub)/1.05)) Tax,Max(x1) MaxCaseNo" & _
      " from (select a0j02 x1,sum(nvl(a0j09,0)+nvl(a0j10,0)) Tot from acc0j0 where a0j13='" & pReceiptNo & "' group by a0j02" & _
      ") x,(select a0j02 y1,nvl(sum(a1u07),0)+nvl(sum(a1u09),0) Sub  from acc0j0,acc1u0 where a0j13='" & pReceiptNo & "'" & _
      " and a1u02(+)=a0j13 and a1u03(+)=a0j01 group by a0j02) y where y1(+)=x1"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      '本案為最大案號時
      If pCaseNo = rsQuery(1) Then
         '若整張與各案加總不同則將差額調至本案
         If lngReceiptTax <> rsQuery(0) Then
            lngCaseTax = lngReceiptTax - (rsQuery(0) - lngCaseTax)
         End If
      End If
   End If


   If pRecFee > 0 Then
      '本次扣稅額=已收規費總額-規費科目總額[>0,<=本次規費]
      stSQL = "select max(nvl(a0j10,0))-nvl(sum(a1u09),0) Net,nvl(sum(a1u05),0)-nvl(sum(a1u10),0) RecAmt" & _
         " from acc0j0,acc1u0 where a0j13='" & pReceiptNo & "' and a0j01='" & pReceiveNo & "' and a1u02(+)=a0j13 and a1u03(+)=a0j01"
      intR = 1
      Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         '規費科目總額=收文號的應收規費-該案的銷項稅額
         lngTaxLimit = rsQuery(0) - lngCaseTax
         '已收規費總額>規費科目總額
         If rsQuery(1) > lngTaxLimit Then
            lngTax = rsQuery(1) - lngTaxLimit
            If lngTax > pRecFee Then
               lngTax = pRecFee
            End If
         End If
      End If
   End If
   
   strRemark = strMan & "/" & strCompany
   'Added by Morgan 2015/5/28 +發票號碼
   'Modified by Morgan 2015/7/14
   stSQL = "select axc01 from acc431 where axc02='" & pReceiptNo & "'"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      strRemark = strRemark & "/" & rsQuery(0)
   End If
   'end 2015/5/28
   
   'Modified by Morgan 2015/11/12
   'If lngTax > 0 Then
   bolDone = False
   If lngTax = lngCaseTax Then
      '更新規費金額
      If pRecFee = lngTax Then
         strSql = "delete acc1p0 where a1p04='" & Text2 & "' and a1p02='A' and a1p03='" & pA1P03 & "'"
         adoTaie.Execute strSql, intI

      ElseIf pRecFee > lngTax Then
         strSql = "update acc1p0 set a1p08=a1p08-" & lngTax & " where a1p04='" & Text2 & "' and a1p02='A' and a1p03='" & pA1P03 & "'"
         adoTaie.Execute strSql, intI
      End If
      bolDone = True
      
'Removed by Morgan 2015/11/12
'      '預收銷項稅額
'      If oblClear = False Then
'         '單據編號=收文號+收據號
'         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
'         'Modified by Morgan 2015/8/4 +a1p17對沖代號(本所案號)
'         strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p14, a1p15, a1p16, a1p17, a1p18, a1p22, a1p23, a1p27)" & _
'            " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '2405', 'TOT', 0, " & lngTax & ", '" & strRemark & "', '" & strCompanyNo & "', '" & strManNo & "','" & pCaseNo & "', " & Val(FCDate(strDate)) & ", " & stra1p22 & ",'" & pReceiveNo & pReceiptNo & "', " & stra1p27 & ")"
'         adoTaie.Execute strSql, intI
'      End If
'end 2015/11/12

   End If
   
   '該收據該案號是否已結清
   If oblClear = True Then
      'Modified by Morgan 2015/11/12
      ''純服務費
      'If bolServiceOnly = True Then
      If Not bolDone Then
      'end 2015/11/12
      
         'Added by Morgan 2021/3/30 純服務費則改扣收入科目
         If bolServiceOnly = True And Left(pCaseNo, 3) = "ACS" Then
            strSql = "update acc1p0 set a1p08=a1p08-" & lngCaseTax & " where a1p02='A' and a1p04='" & Text2 & "' and a1p03=(select max(a1p03) from acc1p0 b where a1p02='A' and a1p04='" & Text2 & "' and a1p08>" & lngCaseTax & " and a1p05 like '4%' and a1p17='" & pCaseNo & "')"
            adoTaie.Execute strSql, intI
         Else
         'end 2021/3/30
            
            'Modified by Morgan 2014/7/2 辜
            'strAccCode = "2631" '稅捐準備
            'Modified by Morgan 2015/8/21 改以案件性質抓規費科目
            'strAccCode = "2201" '應付規費
            stSQL = "select a0j04,cpm12,cpm25 from acc0j0,caseprogress,casepropertymap where a0j01='" & pReceiveNo & "' and a0j13='" & pReceiptNo & "' and cp09(+)=a0j01 and cpm01(+)=cp01 and cpm02(+)=cp10"
            intR = 1
            Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
            If intR = 1 Then
               If rsQuery.Fields("a0j04") = "000" Then
                  strAccCode = "" & rsQuery.Fields("cpm12")
               Else
                  strAccCode = "" & rsQuery.Fields("cpm25")
               End If
            End If
            If strAccCode = "" Then strAccCode = "2201"
            'end 2015/8/21
            
            'Added by Morgan 2021/3/31
            '1.如果原案件沒有提規費 2. 該案件已算過結餘 則沖應付規費-安全基金2211
            stA1P30 = ""
            stSQL = "select nvl(sum(a1p07),0) LAmt,nvl(sum(a1p08),0) RAmt from acc1p0 where a1p17='" & pCaseNo & "' and a1p05 like '220%'"
            intR = 1
            Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
            If intR = 1 Then
               '借方加總>=貸方加總(已結餘或無規費)
               If rsQuery("LAmt") >= rsQuery("RAmt") Then
                  strAccCode = "2211"
                  stSys = Left(pCaseNo, Len(pCaseNo) - 9)
                  If stSys = "CFP" Or stSys = "CPS" Then
                     stA1P30 = "CFP"
                  ElseIf stSys = "P" Or stSys = "PS" Then
                     stA1P30 = "CCP"
                  ElseIf stSys = "CFT" Or stSys = "CFC" Or stSys = "S" Then
                     stA1P30 = "CFT"
                  Else
                     stA1P30 = "CCT"
                  End If
               End If
            End If
            'end 2021/3/31
            
            strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
            'Modified by Morgan 2015/8/4 +a1p17對沖代號(本所案號)
            'Modified by Morgan 2023/5/19 +a1p30其他對沖
            strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p14, a1p15, a1p16, a1p17, a1p18, a1p22, a1p27,a1p30)" & _
               " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strAccCode & "', 'TOT', " & lngCaseTax & ", 0, '" & strRemark & "', '" & strCompanyNo & "', '" & strManNo & "','" & pCaseNo & "', " & Val(FCDate(strDate)) & ", " & stra1p22 & ", " & stra1p27 & ",'" & stA1P30 & "')"
            adoTaie.Execute strSql, intI
            
         End If
         
      'Removed by Morgan 2015/11/12
      'ElseIf lngTax < lngCaseTax Then
      '   strAccCode = "2405" '預收銷項稅額
      '   strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
      '   'Modified by Morgan 2015/8/4 +a1p17對沖代號(本所案號)
      '   strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p14, a1p15, a1p16, a1p17, a1p18, a1p22, a1p27)" & _
      '      " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strAccCode & "', 'TOT', " & lngCaseTax - lngTax & ", 0, '" & strRemark & "', '" & strCompanyNo & "', '" & strManNo & "','" & pCaseNo & "', " & Val(FCDate(strDate)) & ", " & stra1p22 & ", " & stra1p27 & ")"
      '   adoTaie.Execute strSql, intI
      'end 2015/11/12
      
      End If
      
      'modify by sonia 2017/3/16 未收款沖帳傳票已不用1135應收銷項稅額,所以此處改扣除2141之銷項稅額
'      '有未收款沖帳傳票A4317
'      If CheckInvDone(pReceiptNo) = True Then
'         strAccCode = "1135" '應收銷項稅額
'      Else
'         strAccCode = "2119" '銷項稅額
'      End If
'      strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
'      'Modified by Morgan 2015/8/4 +a1p17對沖代號(本所案號)
'      strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p14, a1p15, a1p16, a1p17, a1p18, a1p22, a1p27)" & _
'         " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strAccCode & "', 'TOT', 0, " & lngCaseTax & ", '" & strRemark & "', '" & strCompanyNo & "', '" & strManNo & "','" & pCaseNo & "', " & Val(FCDate(strDate)) & ", " & stra1p22 & ", " & stra1p27 & ")"
'      adoTaie.Execute strSql, intI
      '有未收款沖帳傳票A4317,則2141要扣除銷項稅額
      If CheckInvDone(pReceiptNo) = True Then
         strSql = "update acc1p0 set a1p07=a1p07-" & lngCaseTax & " where a1p01='J' and a1p04='" & Text2 & "' and a1p05='2141' and a1p07>0 and a1p14='" & strRemark & "'"
         adoTaie.Execute strSql, intI
      Else
         strAccCode = "2119" '銷項稅額
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
         strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p14, a1p15, a1p16, a1p17, a1p18, a1p22, a1p27)" & _
            " values ('" & strA1P01 & "', 'A', '" & strSerialNo & "', '" & Text2 & "', '" & strAccCode & "', 'TOT', 0, " & lngCaseTax & ", '" & strRemark & "', '" & strCompanyNo & "', '" & strManNo & "','" & pCaseNo & "', " & Val(FCDate(strDate)) & ", " & stra1p22 & ", " & stra1p27 & ")"
         adoTaie.Execute strSql, intI
      End If
      'modify by sonia 2017/3/16
   End If
   
   AddTaxItem = True
   Set rsQuery = Nothing
End Function

'add by sonia 2016/10/19
'檢查當天是否有已收款通知林柳岑特助
Public Function CheckAccDeliverF5639(pReceiver As String, pCP09 As String) As Boolean
   Dim stSQL As String, intR As Integer
   stSQL = "update mailcache set mc01=mc01 where mc02='" & pReceiver & "' and instr(mc07,'【" & pCP09 & "】寰華介紹案件已收款通知！')>0"
   cnnConnection.Execute stSQL, intR
   If intR > 0 Then
      CheckAccDeliverF5639 = True
   End If
End Function
'end 2016/10/19

'Added by Morgan 2021/1/19
Private Function GetPTCase(pLOS15 As String, ByRef pCP17 As Long) As String
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   stSQL = "select c2.cp01||c2.cp02||c2.cp03||c2.cp04 CNo,c1.cp17 PFee" & _
      " from lawofficesource,caseprogress c1,caseprogress c2" & _
      " where los15='" & pLOS15 & "' and c1.cp09(+)=los10 and c2.cp09(+)=los01"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      GetPTCase = "" & rsQuery("CNo")
      pCP17 = Val("" & rsQuery("PFee"))
   End If
End Function

'Added by Morgan 2021/3/10
'設定案源變數
Private Sub SetLOSVar(pA0K01 As String, pA0K11 As String, ByRef pLOS02 As String, ByRef pTTMan As String, ByRef pTTManSN As String, ByRef pLCaseNo As String, pLRcpTitle As String, ByRef pTTAmt As Long, ByRef pB2NeeCourt As Boolean)
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   pLOS02 = "" '案源類別
   pTTMan = "" '介紹人
   pTTManSN = "" '介紹簡稱
   pLCaseNo = ""
   pLRcpTitle = ""
   pTTAmt = 0
   
   If pA0K11 = "1" Then
      'Modified by Morgan 2023/8/23
      'stSQL = "select los02,c1.cp13 sal,sn01,a0k04,j1.a0j09+j1.a0j10 TTAmt" & _
         ",c2.cp01||c2.cp02||decode(c2.cp04,'00',decode(c2.cp03,'0','','-'||c2.cp03),'-'||c2.cp04) LCase,los15" & _
         " from acc0j0 j1,lawofficesource,caseprogress c1,salesno,caseprogress c2,caseprogress c3,acc0j0 j2,acc0k0" & _
         " where j1.a0j13='" & pA0K01 & "' and los10(+)=j1.a0j01 and los15 is not null" & _
         " and c1.cp09(+)=los10 and sn02(+)=c1.cp13 and c2.cp09(+)=los06" & _
         " and c3.cp01(+)=c2.cp01 and c3.cp02(+)=c2.cp02 and c3.cp03(+)=c2.cp03 and c3.cp04(+)=c2.cp04" & _
         " and c3.cp162(+)=c2.cp162 and j2.a0j01(+)=c3.cp09 and a0k01(+)=j2.a0j13"
      stSQL = "select los02,c1.cp13 sal,sn01,k2.a0k04,k1.a0k06+k1.a0k07 TTAmt" & _
         ",c2.cp01||c2.cp02||decode(c2.cp04,'00',decode(c2.cp03,'0','','-'||c2.cp03),'-'||c2.cp04) LCase,los15" & _
         " from acc0k0 k1,acc0j0 j1,lawofficesource,caseprogress c1,salesno,caseprogress c2,caseprogress c3,acc0j0 j2,acc0k0 k2" & _
         " where k1.a0k01='" & pA0K01 & "' and j1.a0j13(+)=k1.a0k01 and los10(+)=j1.a0j01 and los15 is not null" & _
         " and c1.cp09(+)=los10 and sn02(+)=c1.cp13 and c2.cp09(+)=los06" & _
         " and c3.cp01(+)=c2.cp01 and c3.cp02(+)=c2.cp02 and c3.cp03(+)=c2.cp03 and c3.cp04(+)=c2.cp04" & _
         " and c3.cp162(+)=c2.cp162 and j2.a0j01(+)=c3.cp09 and k2.a0k01(+)=j2.a0j13"
   Else
      '不確定法律所收據是否開在法律所案源總收文號(LOS06),要用CP162串LOS15檢查
      'Modified by Morgan 2023/8/23
      'stSQL = "select los02,c2.cp13 sal,sn01,a0k04,nvl(j2.a0j09,0)+nvl(j2.a0j10,0) TTAmt" & _
         ",c1.cp01||c1.cp02||decode(c1.cp04,'00',decode(c1.cp03,'0','','-'||c1.cp03),'-'||c1.cp04) LCase,los15" & _
         " from acc0j0 j1,caseprogress c1,lawofficesource,salesno,caseprogress c2,acc0k0,acc0j0 j2" & _
         " where j1.a0j13='" & pA0K01 & "' and c1.cp09(+)=j1.a0j01 and los15(+)=c1.cp162 and los15 is not null" & _
         " and c2.cp09(+)=los10 and sn02(+)=c2.cp13 and a0k01(+)=j1.a0j13 and j2.a0j01(+)=los10"
      stSQL = "select los02,c2.cp13 sal,sn01,k1.a0k04,nvl(k2.a0k06,0)+nvl(k2.a0k07,0) TTAmt" & _
         ",c1.cp01||c1.cp02||decode(c1.cp04,'00',decode(c1.cp03,'0','','-'||c1.cp03),'-'||c1.cp04) LCase,los15" & _
         " from acc0k0 k1,acc0j0 j1,caseprogress c1,lawofficesource,salesno,caseprogress c2,acc0j0 j2,acc0k0 k2" & _
         " where k1.a0k01='" & pA0K01 & "' and j1.a0j13(+)=k1.a0k01 and c1.cp09(+)=j1.a0j01 and los15(+)=c1.cp162 and los15 is not null" & _
         " and c2.cp09(+)=los10 and sn02(+)=c2.cp13 and j2.a0j01(+)=los10 and k2.a0k01(+)=j2.a0j13"
   End If
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      pLOS02 = .Fields("los02")
      pTTMan = .Fields("sal")
      pTTManSN = "" & .Fields("sn01")
      pLCaseNo = .Fields("LCase")
      pLRcpTitle = .Fields("a0k04")
      pTTAmt = .Fields("TTAmt")
      If pLOS02 = "B2" Then pB2NeeCourt = PUB_IsB2NeedCourt(.Fields("los15")) 'Added by Morgan 2021/4/29
      
      'Added by Morgan 2022/11/4
      .MoveNext
      Do While Not .EOF
         'pTTAmt = pTTAmt + .Fields("TTAmt") 'Removed by Morgan 2023/8/23
         If pLOS02 = "B2" And pB2NeeCourt = False Then
            pB2NeeCourt = PUB_IsB2NeedCourt(.Fields("los15"))
         End If
         .MoveNext
      Loop
      'end 2022/11/4
      End With
   End If
End Sub

'Added by Morgan 2021/10/22
'律師出庭費
Private Function GetLawFee(pLawNo As String, pRcptSFee As Long, Optional pCP10 As String) As Long
   Dim lngAmt As Long
   '盧律師出庭費規則:收據金額(服務費)的30% , 最高30000, 最低15000
   If pLawNo = "F5591" Then
      lngAmt = pRcptSFee * 0.3
      If lngAmt > 30000 Then
         lngAmt = 30000
      ElseIf lngAmt < 15000 Then
         lngAmt = 15000
      End If
   '34行政訴訟,每次出庭提5000
   ElseIf pCP10 = "34" Then
      lngAmt = 5000
   Else
      lngAmt = 15000
   End If
   
   GetLawFee = lngAmt
End Function

'Added by Morgan 2024/11/11
'智權公司收款零稅率客戶提醒
Private Sub ZeroTaxCustAlert(pKeyNo As String)
   Dim stSQL As String, intR As Integer
   Dim rsQuery  As ADODB.Recordset
   Dim strMsg As String, stCU178 As String
   
   If Left(pKeyNo, 1) = "F" Then
      stSQL = "select distinct a0k04 from acc0m0,acc0k0 where a0m01='" & pKeyNo & "' and a0k01(+)=a0m02"
   ElseIf Left(pKeyNo, 1) = "E" Then
      stSQL = "select a0k04 from acc0k0 where a0k01='" & pKeyNo & "'"
   End If
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      With rsQuery
      .MoveFirst
      strMsg = ""
      Do While Not .EOF
         stCU178 = ""
         PUB_GetTaxNo .Fields("a0k04"), , , , stCU178
         If stCU178 = "Y" Then
            strMsg = strMsg & .Fields("a0k04") & vbCrLf
         End If
         .MoveNext
      Loop
      If strMsg <> "" Then
         strMsg = "客戶之發票設為零稅率，必須為境外匯款並取得契約書！" & vbCrLf & vbCrLf & strMsg
         MsgBox strMsg, vbExclamation, "零稅率提醒"
      End If
      End With
   End If
   Set rsQuery = Nothing
End Sub
