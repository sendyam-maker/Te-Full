VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc2260 
   AutoRedraw      =   -1  'True
   Caption         =   "國外部智權人員帳款查詢"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5280
   ScaleWidth      =   9360
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "1"
      Top             =   1110
      Width           =   612
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   360
      Left            =   4980
      TabIndex        =   18
      Top             =   30
      Width           =   3765
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   810
         Style           =   2  '單純下拉式
         TabIndex        =   5
         Top             =   30
         Width           =   2910
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "印表機"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.5
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   21
         Top             =   60
         Width           =   735
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2260.frx":0000
      Height          =   3135
      Left            =   90
      TabIndex        =   9
      Top             =   1470
      Width           =   8950
      _ExtentX        =   15796
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "國外部智權人員帳款查詢"
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "請款對象"
         Caption         =   "請款對象"
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
         DataField       =   "請款單號"
         Caption         =   "請款單號"
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
         DataField       =   "請款日"
         Caption         =   "請款日"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "幣別"
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
      BeginProperty Column04 
         DataField       =   "應收外幣"
         Caption         =   "應收外幣"
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
         DataField       =   "已收金額"
         Caption         =   "已收金額"
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
      BeginProperty Column06 
         DataField       =   "未收金額"
         Caption         =   "未收金額"
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
      BeginProperty Column07 
         DataField       =   "本所案號"
         Caption         =   "本所案號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "規費"
         Caption         =   "規費"
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
      BeginProperty Column09 
         DataField       =   "案件最後程序"
         Caption         =   "案件最後程序"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "排序"
         Caption         =   "排序"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
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
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2234.835
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   2970.142
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6120
      TabIndex        =   8
      Top             =   4650
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3840
      TabIndex        =   7
      Top             =   4650
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   6
      Top             =   4650
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "3"
      Top             =   750
      Width           =   612
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   0
      Top             =   30
      Width           =   1065
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Top             =   390
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3120
      TabIndex        =   2
      Top             =   390
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   1380
      Visible         =   0   'False
      Width           =   1200
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
   Begin MSForms.Label lblStaffName 
      Height          =   285
      Left            =   2340
      TabIndex        =   22
      Top             =   45
      Width           =   1965
      VariousPropertyBits=   19
      Size            =   "3466;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "輸出方式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   1140
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "(1.螢幕 2.印表機)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   19
      Top             =   1110
      Width           =   2595
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "未收"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5550
      TabIndex        =   17
      Top             =   4680
      Width           =   510
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "已收"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3270
      TabIndex        =   16
      Top             =   4680
      Width           =   510
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "應收合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   4680
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   -60
      Top             =   1590
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "(1.未收 2.收款 3.往來)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   14
      Top             =   750
      Width           =   2595
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "查詢資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   780
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   390
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "請款日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   420
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   60
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc2260"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/22 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Label1(2)=>lblStaffName
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoadodc1 As New ADODB.Recordset
Public adoadodc2 As New ADODB.Recordset
Public adoacc0z0 As New ADODB.Recordset
Dim PLeft(1 To 10) As Integer
Dim strTemp(1 To 10) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim m_ST05 As String 'Add By Sindy 2011/1/14

Private Sub Form_Activate()
   Dim formCnt As Integer
   For formCnt = 0 To Forms.Count - 1
       If UCase(Forms(formCnt).Name) = "MDIMAIN" Then
             Forms(formCnt).ToolShow
             Exit For
       End If
   Next
   strFormLink = ""
   strFormName = Name
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim strSql As String, i As Integer, j As Integer
Dim strSystemKind As String
   
   MoveFormToCenter Me
   
   m_ST05 = PUB_GetST05(strUserNum) 'Add By Sindy 2011/1/14
   
   strSystemKind = GetSystemKindByNick
   strSql = Printer.DeviceName
   SeekPrintL = Printer.Orientation
   For i = 0 To Printers.Count - 1
      Set Printer = Printers(i)
      Combo1.AddItem Printer.DeviceName, j
      j = j + 1
      If Printer.DeviceName = strSql Then
         SeekPrint = i
      End If
   Next i
   Set Printer = Printers(SeekPrint)
   Combo1.Text = Combo1.List(SeekPrint)
   
   'Modified by Lydia 2021/12/09 改成模組
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   Me.Width = 9500
'   Me.Height = 5400
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath2)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   strFormName = Name
   'Modify by Amy 2023/10/11避免切畫面還要調整修改版面大小, 原9500, 5400 -瑞婷
   PUB_InitForm Me, 9480, 5745, strBackPicPath2
   'end 2021/12/09
   
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   
   'Add By Sindy 2011/1/14 加入權限控管
   Select Case m_ST05
      Case "00", "01", "09" '不限制
      Case Else
         Text1.Text = strUserNum
         'Modified by Lydia 2021/09/22 Label1(2)=>lblStaffName
         lblStaffName = GetPrjSalesNM(strUserNum)
         '2011/2/14 modify by sonia 加入"51"
         'If m_ST05 <> "11" And m_ST05 <> "26" And m_ST05 <> "28" And m_ST05 <> "35" Then
         If m_ST05 <> "11" And m_ST05 <> "26" And m_ST05 <> "28" And m_ST05 <> "35" And m_ST05 <> "51" Then
            Text1.Enabled = False '鎖住
         Else
            Text1.Enabled = True
         End If
   End Select
   '2011/1/14 End
   
   OpenTable
   StatusView MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc2260 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc1k0 where a1k01 = 'Z'", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  查詢資料表(傳票資料)
'
'*************************************************
Public Sub QueryTable()
Dim Cancel As Boolean
On Error GoTo Checking
   
   Cancel = False
   Call Text1_Validate(Cancel)
   If Cancel = True Then
      Text1.SetFocus
      Exit Sub
   End If
   If Trim(Text1.Text) = "" Then
      MsgBox "智權人員不可空白!", , MsgText(5)
      Text1.SetFocus
      Exit Sub
   End If
   If Val(FCDate(MaskEdBox1.Text)) = 0 Then
      MaskEdBox1.Text = "001/01/01"
      'MsgBox "起始請款日不可空白!", , MsgText(5)
      'MaskEdBox1.SetFocus
      'Exit Sub
   End If
   If Val(FCDate(MaskEdBox2.Text)) = 0 Then
      MsgBox "迄止請款日不可空白!", , MsgText(5)
      MaskEdBox2.SetFocus
      Exit Sub
   End If
   
   'Adodc1.Recordset.Close '清除畫面
   Text4.Text = ""
   Text5.Text = ""
   Text6.Text = ""
   strSql = ""
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   If adoadodc2.State = adStateOpen Then
      adoadodc2.Close
   End If
   adoadodc2.CursorLocation = adUseClient
   
   'Modify By Sindy 2012/12/7 A1K08-NVL(A1K06,0) RECAMT ==> A1K08-NVL(A1K31,0) RECAMT
   '                          RECAMT 應收美金 ==> RECAMT 應收外幣
   'Modify By Sindy 2013/1/14 加幣別欄位
   Select Case Text3
      Case "1" '未收
'         strSQL = "SELECT A1K28 請款對象,A1K01 請款單號,A1K02 請款日,RECAMT 應收美金,PAYAMT 已收金額," & _
'                         "decode(A1K29,'Y','',nvl(RECAMT,0)-nvl(PAYAMT,0)) 未收金額,A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16 本所案號,A1K09 規費," & _
'                         "DECODE(C2.CP43,'',SUBSTR(C1.CPM03,18),SUBSTR(C1.CPM03,18)||'-'||SUBSTR(M2.CPM03,1,6)) 案件最後程序,A1K02||A1K01 排序,A1K30" & _
'                         " FROM (SELECT A1K28,A1K01,A1K02,A1K08-NVL(A1K06,0) RECAMT,A1K13,A1K14,A1K15,A1K16,A1K09,SUM(A0Z04) PAYAMT,A1K29,A1K30" & _
'                         " From CASEPROGRESS, ACC1K0, ACC0Z0" & _
'                         " WHERE CP13='" & Text1 & "' AND CP60 IS NOT NULL AND SUBSTR(CP60,1,1)='X' AND CP60=A1K01(+)" & _
'                         " AND a1k02>=" & Val(FCDate(MaskEdBox1.Text)) & " and a1k02<=" & Val(FCDate(MaskEdBox2.Text)) & " AND A1K29 IS NULL AND A1K25 IS NULL AND A1K01=A0Z02(+)" & _
'                         " GROUP BY A1K28,A1K01,A1K02,A1K08-NVL(A1K06,0),A1K13,A1K14,A1K15,A1K16,A1K09,A1K29,A1K30)," & _
'                         " (SELECT CP09,CP43 FROM CASEPROGRESS" & _
'                         " WHERE (CP01,CP02,CP03,CP04) IN (SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS WHERE CP13='" & Text1 & "' AND CP60 IS NOT NULL AND SUBSTR(CP60,1,1)='X')) C2," & _
'                         " (SELECT CP01,CP09,CP10 FROM CASEPROGRESS" & _
'                         " WHERE (CP01,CP02,CP03,CP04) IN (SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS WHERE CP13='" & Text1 & "' AND CP60 IS NOT NULL AND SUBSTR(CP60,1,1)='X')) C3," & _
'                         " CASEPROPERTYMAP M2," & _
'                         " (SELECT CP01,CP02,CP03,CP04,MAX(CP05||CP09||SUBSTR(CPM03,1,6)) CPM03 FROM CASEPROGRESS,CASEPROPERTYMAP" & _
'                         " WHERE (CP01,CP02,CP03,CP04) IN (SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS WHERE CP13='" & Text1 & "' AND CP60 IS NOT NULL AND SUBSTR(CP60,1,1)='X')" & _
'                         " AND SUBSTR(CP09,1,1)<>'B' AND CP01=CPM01(+) AND CP10=CPM02(+)" & _
'                         " GROUP BY CP01,CP02,CP03,CP04) C1" & _
'                         " WHERE A1K13=C1.CP01(+) AND A1K14=C1.CP02(+) AND A1K15=C1.CP03(+) AND A1K16=C1.CP04(+)" & _
'                         " AND SUBSTR(C1.CPM03,9,9)=C2.CP09(+)" & _
'                         " AND C2.CP43=C3.CP09(+)" & _
'                         " AND C3.CP01=M2.CPM01(+) AND C3.CP10=M2.CPM02(+)" & _
'                         " ORDER BY 請款日,請款單號"
         strSql = "SELECT A1K28 請款對象,A1K01 請款單號,A1K02 請款日,A1K18 幣別,RECAMT 應收外幣,PAYAMT 已收金額," & _
                         "decode(A1K29,'Y','',nvl(RECAMT,0)-nvl(PAYAMT,0)) 未收金額,A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16 本所案號,A1K09 規費," & _
                         "DECODE(C2.CP43,'',SUBSTR(C1.CPM03,18),SUBSTR(C1.CPM03,18)||'-'||SUBSTR(M2.CPM03,1,6)) 案件最後程序,A1K02||A1K01 排序,A1K30" & _
                         " FROM (SELECT A1K28,A1K01,A1K02,A1K18,A1K08-NVL(A1K31,0) RECAMT,A1K13,A1K14,A1K15,A1K16,A1K09,SUM(A0Z04) PAYAMT,A1K29,A1K30" & _
                         " From CASEPROGRESS, ACC1K0, ACC0Z0" & _
                         " WHERE CP13='" & Text1 & "' AND CP60 IS NOT NULL AND CP60>'X' AND CP60=A1K01(+)" & _
                         " AND a1k02>=" & Val(FCDate(MaskEdBox1.Text)) & " and a1k02<=" & Val(FCDate(MaskEdBox2.Text)) & " AND A1K29 IS NULL AND A1K25 IS NULL AND A1K01=A0Z02(+)" & _
                         " GROUP BY A1K28,A1K01,A1K02,A1K18,A1K08-NVL(A1K31,0),A1K13,A1K14,A1K15,A1K16,A1K09,A1K29,A1K30)," & _
                         " CASEPROGRESS  C2,CASEPROGRESS C3,CASEPROPERTYMAP M2," & _
                         " (SELECT CP01,CP02,CP03,CP04,MAX(CP05||CP09||SUBSTR(CPM03,1,6)) CPM03 FROM CASEPROGRESS,CASEPROPERTYMAP" & _
                         " WHERE (CP01,CP02,CP03,CP04) IN (SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS WHERE CP13='" & Text1 & "' AND CP60 IS NOT NULL AND CP60>'X')" & _
                         " AND SUBSTR(CP09,1,1)<>'B' AND CP01=CPM01(+) AND CP10=CPM02(+)" & _
                         " GROUP BY CP01,CP02,CP03,CP04) C1" & _
                         " WHERE A1K13=C1.CP01(+) AND A1K14=C1.CP02(+) AND A1K15=C1.CP03(+) AND A1K16=C1.CP04(+)" & _
                         " AND SUBSTR(C1.CPM03,9,9)=C2.CP09(+)" & _
                         " AND C2.CP43=C3.CP09(+)" & _
                         " AND C3.CP01=M2.CPM01(+) AND C3.CP10=M2.CPM02(+)" & _
                         " ORDER BY 請款日,請款單號"
      Case "2" '收款
         strSql = "SELECT A1K28 請款對象,A1K01 請款單號,A1K02 請款日,A1K18 幣別,RECAMT 應收外幣,PAYAMT 已收金額," & _
                         "decode(A1K29,'Y','',nvl(RECAMT,0)-nvl(PAYAMT,0)) 未收金額,A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16 本所案號,A1K09 規費," & _
                         "DECODE(C2.CP43,'',SUBSTR(C1.CPM03,18),SUBSTR(C1.CPM03,18)||'-'||SUBSTR(M2.CPM03,1,6)) 案件最後程序,A1K02||A1K01 排序,A1K30" & _
                         " FROM (SELECT A1K28,A1K01,A1K02,A1K18,A1K08-NVL(A1K31,0) RECAMT,A1K13,A1K14,A1K15,A1K16,A1K09,SUM(A0Z04) PAYAMT,A1K29,A1K30" & _
                         " From CASEPROGRESS, ACC1K0, ACC0Z0" & _
                         " WHERE CP13='" & Text1 & "' AND CP60 IS NOT NULL AND CP60>'X' AND CP60=A1K01(+)" & _
                         " AND a1k02>=" & Val(FCDate(MaskEdBox1.Text)) & " and a1k02<=" & Val(FCDate(MaskEdBox2.Text)) & " AND A1K30>0 AND A1K01=A0Z02(+)" & _
                         " GROUP BY A1K28,A1K01,A1K02,A1K18,A1K08-NVL(A1K31,0),A1K13,A1K14,A1K15,A1K16,A1K09,A1K29,A1K30)," & _
                         " CASEPROGRESS  C2,CASEPROGRESS C3,CASEPROPERTYMAP M2," & _
                         " (SELECT CP01,CP02,CP03,CP04,MAX(CP05||CP09||SUBSTR(CPM03,1,6)) CPM03 FROM CASEPROGRESS,CASEPROPERTYMAP" & _
                         " WHERE (CP01,CP02,CP03,CP04) IN (SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS WHERE CP13='" & Text1 & "' AND CP60 IS NOT NULL AND CP60>'X')" & _
                         " AND SUBSTR(CP09,1,1)<>'B' AND CP01=CPM01(+) AND CP10=CPM02(+)" & _
                         " GROUP BY CP01,CP02,CP03,CP04) C1" & _
                         " WHERE A1K13=C1.CP01(+) AND A1K14=C1.CP02(+) AND A1K15=C1.CP03(+) AND A1K16=C1.CP04(+)" & _
                         " AND SUBSTR(C1.CPM03,9,9)=C2.CP09(+)" & _
                         " AND C2.CP43=C3.CP09(+)" & _
                         " AND C3.CP01=M2.CPM01(+) AND C3.CP10=M2.CPM02(+)" & _
                         " ORDER BY 請款日,請款單號"
      Case "3", "" '往來
         strSql = "SELECT A1K28 請款對象,A1K01 請款單號,A1K02 請款日,A1K18 幣別,RECAMT 應收外幣,PAYAMT 已收金額," & _
                         "decode(A1K29,'Y','',nvl(RECAMT,0)-nvl(PAYAMT,0)) 未收金額,A1K13||'-'||A1K14||'-'||A1K15||'-'||A1K16 本所案號,A1K09 規費," & _
                         "DECODE(C2.CP43,'',SUBSTR(C1.CPM03,18),SUBSTR(C1.CPM03,18)||'-'||SUBSTR(M2.CPM03,1,6)) 案件最後程序,A1K02||A1K01 排序,A1K30" & _
                         " FROM (SELECT A1K28,A1K01,A1K02,A1K18,A1K08-NVL(A1K31,0) RECAMT,A1K13,A1K14,A1K15,A1K16,A1K09,SUM(A0Z04) PAYAMT,A1K29,A1K30" & _
                         " From CASEPROGRESS, ACC1K0, ACC0Z0" & _
                         " WHERE CP13='" & Text1 & "' AND CP60 IS NOT NULL AND CP60>'X' AND CP60=A1K01(+)" & _
                         " AND a1k02>=" & Val(FCDate(MaskEdBox1.Text)) & " and a1k02<=" & Val(FCDate(MaskEdBox2.Text)) & " AND A1K01=A0Z02(+)" & _
                         " GROUP BY A1K28,A1K01,A1K02,A1K18,A1K08-NVL(A1K31,0),A1K13,A1K14,A1K15,A1K16,A1K09,A1K29,A1K30)," & _
                         " CASEPROGRESS  C2,CASEPROGRESS C3,CASEPROPERTYMAP M2," & _
                         " (SELECT CP01,CP02,CP03,CP04,MAX(CP05||CP09||SUBSTR(CPM03,1,6)) CPM03 FROM CASEPROGRESS,CASEPROPERTYMAP" & _
                         " WHERE (CP01,CP02,CP03,CP04) IN (SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS WHERE CP13='" & Text1 & "' AND CP60 IS NOT NULL AND CP60>'X')" & _
                         " AND SUBSTR(CP09,1,1)<>'B' AND CP01=CPM01(+) AND CP10=CPM02(+)" & _
                         " GROUP BY CP01,CP02,CP03,CP04) C1" & _
                         " WHERE A1K13=C1.CP01(+) AND A1K14=C1.CP02(+) AND A1K15=C1.CP03(+) AND A1K16=C1.CP04(+)" & _
                         " AND SUBSTR(C1.CPM03,9,9)=C2.CP09(+)" & _
                         " AND C2.CP43=C3.CP09(+)" & _
                         " AND C3.CP01=M2.CPM01(+) AND C3.CP10=M2.CPM02(+)" & _
                         " ORDER BY 請款日,請款單號"
   End Select
   
   adoadodc1.Open strSql, adoTaie, adOpenStatic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount = 0 Then
      Adodc1.Recordset.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   Else
      If Text3 <> "1" Then
         adoadodc2.Open strSql, adoTaie, adOpenStatic, adLockBatchOptimistic
      End If
      SumShow
   End If
   
   '列印報表
   If Text7 = 2 Then
      Call PrintData
   End If
   
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
Public Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         If FormCheck Then
            Screen.MousePointer = vbHourglass
            QueryTable
            Screen.MousePointer = vbDefault
            Exit Sub
         Else
            MsgBox MsgText(181), , MsgText(5)
         End If
   End Select
   KeyEnter KeyCode
   StatusView MsgText(98)
End Sub

'*************************************************
'  計算並顯示合計
'
'*************************************************
Public Sub SumShow()
Dim dbl_Text4 As Double, dbl_Text5 As Double, dbl_Text6 As Double
Dim i As Integer, strSort As String
   
   dbl_Text4 = 0: dbl_Text5 = 0: dbl_Text6 = 0
   Text4 = "": Text5 = "": Text6 = ""
   
   If adoadodc1.RecordCount <> 0 Then
      If Text3 = "1" Then
         adoadodc1.MoveFirst
         Do While Not adoadodc1.EOF
            'Modify By Sindy 2012/12/7 應收美金 ==> 應收外幣
            dbl_Text4 = dbl_Text4 + Val("" & adoadodc1.Fields("應收外幣"))
            dbl_Text5 = dbl_Text5 + Val("" & adoadodc1.Fields("已收金額"))
            dbl_Text6 = dbl_Text6 + Val("" & adoadodc1.Fields("未收金額"))
            adoadodc1.MoveNext
         Loop
         adoadodc1.MoveFirst
      Else
         adoadodc2.MoveFirst
         Do While Not adoadodc2.EOF
            'Modify By Sindy 2012/12/7 應收美金 ==> 應收外幣
            dbl_Text4 = dbl_Text4 + Val("" & adoadodc2.Fields("應收外幣"))
            dbl_Text5 = dbl_Text5 + Val("" & adoadodc2.Fields("已收金額"))
            dbl_Text6 = dbl_Text6 + Val("" & adoadodc2.Fields("未收金額"))
            '加入收款資料
            If Val("" & adoadodc2.Fields("A1K30")) <> 0 Then
               adoacc0z0.CursorLocation = adUseClient
               strSql = "select a0z01,a0y02,a0z04 from acc0z0,acc0y0 where a0z01=a0y01(+) and a0z02='" & adoadodc2.Fields("請款單號") & "'"
               adoacc0z0.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
               If adoacc0z0.RecordCount <> 0 Then
                  adoacc0z0.MoveFirst
                  i = 0
                  strSort = adoadodc2.Fields("請款日") & adoadodc2.Fields("請款單號")
                  Do While Not adoacc0z0.EOF
                     With Adodc1.Recordset
                        .AddNew
                        i = i + 1
                        .Fields("請款單號").Value = adoacc0z0.Fields("a0z01")
                        .Fields("請款日").Value = adoacc0z0.Fields("a0y02")
                        .Fields("已收金額").Value = adoacc0z0.Fields("a0z04")
                        .Fields("排序").Value = strSort & "-" & i
                     End With
                     adoacc0z0.MoveNext
                  Loop
               End If
               adoacc0z0.Close
            End If
            adoadodc2.MoveNext
         Loop
         adoadodc1.Sort = "排序"
         adoadodc1.MoveFirst
      End If
      Text4 = dbl_Text4
      Text5 = dbl_Text5
      Text6 = dbl_Text6
   End If
End Sub

Private Sub PrintData()
Dim i As Integer

Set Printer = Printers(Combo1.ListIndex)
Printer.EndDoc
Printer.Orientation = 2 '1.直印 2.橫印

If Adodc1.Recordset.RecordCount > 0 Then
   With adoadodc1
      .MoveFirst
      iLine = 1
      strType = ""
      Do While Not .EOF
         For i = 1 To 10
            strTemp(i) = ""
         Next i
         strTemp(1) = CheckStr(.Fields(0))
         strTemp(2) = CheckStr(.Fields(1))
         strTemp(3) = CheckStr(.Fields(2))
         strTemp(4) = CheckStr(.Fields(3))
         strTemp(5) = CheckStr(.Fields(4))
         strTemp(6) = CheckStr(.Fields(5))
         strTemp(7) = CheckStr(.Fields(6))
         strTemp(8) = CheckStr(.Fields(7))
         strTemp(9) = CheckStr(.Fields(8))
         strTemp(10) = CheckStr(.Fields(9))
         If iLine > 37 Or iLine = 1 Then
            If strType <> "" Then Printer.NewPage
            iLine = 1
            PrintTitle '列印表頭
         End If
         PrintDetail
         strType = strTemp(2)
         .MoveNext
      Loop
   End With
End If
Printer.EndDoc
ShowPrintOk
End Sub

Sub GetPleft()
PLeft(1) = 500
PLeft(2) = 2000
PLeft(3) = 3500
PLeft(4) = 4600
PLeft(5) = 6500
PLeft(6) = 8000
PLeft(7) = 9500
PLeft(8) = 10000
PLeft(9) = 12500
PLeft(10) = 13000
End Sub

Sub PrintTitle()
GetPleft

Printer.Font.Size = 16
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("智權人員帳款明細表") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "智權人員帳款明細表"

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = PLeft(1)
Printer.CurrentY = 900
Printer.Print "列印人員：" & strUserName
Printer.CurrentX = 6800
Printer.CurrentY = 900
Printer.Print "帳款日期：" & MaskEdBox1.Text & "-" & MaskEdBox2.Text
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
Printer.CurrentX = PLeft(1)
Printer.CurrentY = 1200
'Modified by Lydia 2021/09/22 Label1(2)=>lblStaffName
Printer.Print "業  務  員：" & lblStaffName
Printer.CurrentX = 6800
Printer.CurrentY = 1200
Printer.Print "部  門  別：" & GetStaffDepartment(Text1) & "　" & GetDepartmentName(GetStaffDepartment(Text1))
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 1200
Printer.Print "頁　　次：" & Printer.Page

iLine = 6
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "請款對象"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "請款單號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "請款日"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "幣別"
'Modify By Sindy 2012/12/7 應收美金 ==> 應收外幣
Printer.CurrentX = PLeft(5) - Printer.TextWidth("應收外幣")
Printer.CurrentY = iLine * 300
Printer.Print "應收外幣"
Printer.CurrentX = PLeft(6) - Printer.TextWidth("已收金額")
Printer.CurrentY = iLine * 300
Printer.Print "已收金額"
Printer.CurrentX = PLeft(7) - Printer.TextWidth("未收金額")
Printer.CurrentY = iLine * 300
Printer.Print "未收金額"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iLine * 300
Printer.Print "本所案號"
Printer.CurrentX = PLeft(9) - Printer.TextWidth("規費")
Printer.CurrentY = iLine * 300
Printer.Print "規費"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iLine * 300
Printer.Print "案件最後程序"

iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print String(205, "-")
iLine = iLine + 1
End Sub

Sub PrintDetail()
Dim m_j As Integer
For m_j = 1 To 10
   If m_j = 5 Or m_j = 6 Or m_j = 7 Or m_j = 9 Then
      Printer.CurrentX = PLeft(m_j) - Printer.TextWidth(strTemp(m_j))
   Else
      Printer.CurrentX = PLeft(m_j)
   End If
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(m_j)
Next m_j
iLine = iLine + 1
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Text1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If Text3 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   'Modified by Lydia 2021/09/22 Label1(2)=>lblStaffName
   lblStaffName.Caption = ""
   If Trim(Text1.Text) <> "" Then
      'Add By Sindy 2011/1/14 加入權限控管
      Select Case m_ST05
         Case "11" '可查ST03='F1'字頭所有人
            If Left(PUB_GetST03(Text1.Text), 2) <> "F1" Then
               MsgBox "沒有查詢此智權人員的權限!", , MsgText(5)
               Text1_GotFocus
               Text1 = ""
               Cancel = True
               Exit Sub
            End If
         Case "26", "28" '可查ST03及ST16相同的所有人
            If PUB_GetStaffST16(strUserNum) <> PUB_GetStaffST16(Text1.Text) Then
               MsgBox "沒有查詢此智權人員的權限!", , MsgText(5)
               Text1_GotFocus
               Text1 = ""
               Cancel = True
               Exit Sub
            End If
         Case "35" '可查ST03='F23'所有人
            If PUB_GetST03(Text1.Text) <> "F23" Then
               MsgBox "沒有查詢此智權人員的權限!", , MsgText(5)
               Text1_GotFocus
               Text1 = ""
               Cancel = True
               Exit Sub
            End If
         '2011/2/14 add by sonia
         Case "51" '可查ST03='F31'及'F41'所有人
            'MODIFY BY SONIA 2015/5/20 加入法務部L02,因P31及F31人員併入L02
            If PUB_GetST03(Text1.Text) <> "F31" And PUB_GetST03(Text1.Text) <> "F41" And PUB_GetST03(Text1.Text) <> "L02" Then
               MsgBox "沒有查詢此智權人員的權限!", , MsgText(5)
               Text1_GotFocus
               Text1 = ""
               Cancel = True
               Exit Sub
            End If
         '2011/2/14 end
         Case Else
      End Select
      '2011/1/14 End
      'Modified by Lydia 2021/09/22 Label1(2)=>lblStaffName
      lblStaffName.Caption = GetPrjSalesNM(Trim(Text1.Text))
   End If
   Cancel = False
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    Select Case KeyAscii
    Case 8, 49, 50, 51
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    Select Case KeyAscii
    Case 49, 50
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
   If Text7 = "" Then Text7 = "1"
End Sub
