VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc1290 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "手開收據資料查詢"
   ClientHeight    =   5115
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   9405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   9405
   Begin VB.TextBox Text11 
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
      Left            =   6690
      MaxLength       =   1
      TabIndex        =   9
      Top             =   1230
      Width           =   300
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  '靠右對齊
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
      Left            =   7605
      MaxLength       =   9
      TabIndex        =   22
      Top             =   4740
      Width           =   1260
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  '靠右對齊
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
      Left            =   5040
      MaxLength       =   9
      TabIndex        =   21
      Top             =   4740
      Width           =   1260
   End
   Begin VB.TextBox Text8 
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
      Left            =   6420
      MaxLength       =   9
      TabIndex        =   7
      Top             =   840
      Width           =   1260
   End
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
      Height          =   315
      Left            =   4755
      MaxLength       =   9
      TabIndex        =   6
      Top             =   840
      Width           =   1260
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Frmacc1290.frx":0000
      Left            =   1080
      List            =   "Frmacc1290.frx":000A
      Style           =   2  '單純下拉式
      TabIndex        =   8
      Top             =   1230
      Width           =   3435
   End
   Begin VB.TextBox Text6 
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
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   5
      Top             =   840
      Width           =   600
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8280
      TabIndex        =   17
      Top             =   540
      Width           =   990
   End
   Begin VB.TextBox Text5 
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
      Left            =   4755
      MaxLength       =   1
      TabIndex        =   4
      Top             =   480
      Width           =   300
   End
   Begin VB.CommandButton Command1 
      Caption         =   "合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8280
      TabIndex        =   10
      Top             =   120
      Width           =   990
   End
   Begin VB.TextBox Text4 
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
      Left            =   6420
      MaxLength       =   9
      TabIndex        =   2
      Top             =   120
      Width           =   1260
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
      Height          =   315
      Left            =   4755
      MaxLength       =   9
      TabIndex        =   1
      Top             =   120
      Width           =   1260
   End
   Begin VB.TextBox Text2 
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
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   3
      Top             =   480
      Width           =   600
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
      Height          =   315
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc1290.frx":0022
      Height          =   3015
      Left            =   90
      TabIndex        =   12
      Top             =   1650
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "手開收據資料查詢"
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "NO"
         Caption         =   "手開收據號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "a0m02"
         Caption         =   "電腦收據號"
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
         DataField       =   "SFee"
         Caption         =   "服務費"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "OFee"
         Caption         =   "規費"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "CUSTOMER"
         Caption         =   "收據抬頭"
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
      BeginProperty Column05 
         DataField       =   "SALES"
         Caption         =   "SALES"
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
         DataField       =   "MEMO2"
         Caption         =   "備註"
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
         DataField       =   "ST02"
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
         DataField       =   "CUSTNO"
         Caption         =   "客戶編號"
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
      BeginProperty Column09 
         DataField       =   "COMP"
         Caption         =   "公司"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "###/##/##"
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
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1980.284
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1470.047
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            ColumnWidth     =   450.142
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   8325
      Top             =   990
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
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "是否含已作廢收據：      (N/Y)"
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
      Left            =   4650
      TabIndex        =   25
      Top             =   1290
      Width           =   3240
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "規費合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6615
      TabIndex        =   24
      Top             =   4740
      Width           =   915
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "服務費合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3870
      TabIndex        =   23
      Top             =   4740
      Width           =   1140
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "客戶編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3675
      TabIndex        =   20
      Top             =   870
      Width           =   915
   End
   Begin VB.Line Line2 
      X1              =   6105
      X2              =   6335
      Y1              =   1005
      Y2              =   1005
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "排序方式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   90
      TabIndex        =   19
      Top             =   1275
      Width           =   915
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收款年度"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   90
      TabIndex        =   18
      Top             =   885
      Width           =   915
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "收據狀態        (1:未使用 2:已使用 空白:全部)"
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
      Left            =   3675
      TabIndex        =   16
      Top             =   510
      Width           =   4620
   End
   Begin VB.Line Line1 
      X1              =   6105
      X2              =   6335
      Y1              =   285
      Y2              =   285
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "收據號碼"
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
      Left            =   3675
      TabIndex        =   15
      Top             =   150
      Width           =   1155
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "年度"
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
      Left            =   90
      TabIndex        =   14
      Top             =   510
      Width           =   660
   End
   Begin MSForms.Label lblSalesName 
      Height          =   300
      Left            =   2115
      TabIndex        =   13
      Top             =   150
      Width           =   1290
      VariousPropertyBits=   19
      Caption         =   "LblFM2"
      Size            =   "2275;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   210
      Left            =   90
      TabIndex        =   11
      Top             =   150
      Width           =   900
   End
End
Attribute VB_Name = "Frmacc1290"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/16 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit

Dim PLeft(0 To 8) As Integer
Dim m_intPage As Integer
Dim m_iPrint As Integer


Private Sub GetPrintLeft()
'Modify by Morgan 2009/9/2 改欄位格式
'   PLeft(0) = 150 '公司別
'   PLeft(1) = PLeft(0) + 800 '客戶編號 1200
'   PLeft(2) = PLeft(1) + 1200 '手開收據號 1200
'   PLeft(3) = PLeft(2) + 1200 '手開金額 1000
'   PLeft(4) = PLeft(3) + 1000 '電腦收據號 1200
'   PLeft(5) = PLeft(4) + 1200 '服務費 1000
'   PLeft(8) = PLeft(5) + 1000 '規費 1000
'   PLeft(6) = PLeft(8) + 1000 '收據抬頭 3500
'   PLeft(7) = PLeft(6) + 3500 '智權人員 810

   PLeft(0) = 150 '手開收據號 1200
   PLeft(1) = PLeft(0) + 1200 '電腦收據號 1200
   PLeft(2) = PLeft(1) + 1200 '服務費 1000
   PLeft(3) = PLeft(2) + 1000 '規費 1000
   PLeft(4) = PLeft(3) + 1000 '收據抬頭 3500
   PLeft(5) = PLeft(4) + 3500 '備註 5500
   PLeft(6) = PLeft(5) + 5500 '智權人員 810
   PLeft(7) = PLeft(6) + 810 '客戶編號 1200
   PLeft(8) = PLeft(7) + 1200 '公司別
End Sub

Private Sub PrintHead()
   m_intPage = m_intPage + 1
   m_iPrint = 150
   'Modify by Morgan 2009/9/2 改橫印
   'Printer.Orientation = 1
   Printer.Orientation = 2
   
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 20
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   'Printer.CurrentX = 3328
   Printer.CurrentX = 5500
   Printer.CurrentY = m_iPrint
   Printer.Print "*** 智權人員手開收據明細 ***"
   m_iPrint = m_iPrint + 500
   Printer.Font.Size = 10
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = m_iPrint
   Printer.Print "列印人：" & GetStaffName(strUserNum)
   
   'Printer.CurrentX = 5500 - Printer.TextWidth("智權人員：")
   Printer.CurrentX = 8000 - Printer.TextWidth("智權人員：")
   Printer.CurrentY = m_iPrint
   Printer.Print "智權人員：" & Text1 & " " & lblSalesName
   
   'Printer.CurrentX = 9500
   Printer.CurrentX = 14000
   Printer.CurrentY = m_iPrint
   Printer.Print "列印日期：" & CFDate(ACDate(ServerDate))
   m_iPrint = m_iPrint + 300
      
   'Printer.CurrentX = 5500 - Printer.TextWidth("收據狀態：")
   Printer.CurrentX = 8000 - Printer.TextWidth("收據狀態：")
   Printer.CurrentY = m_iPrint
   Select Case Text5
      Case 1
         Printer.Print "收據狀態：未使用"
      Case 2
         Printer.Print "收據狀態：已使用"
      Case Else
         Printer.Print "收據狀態：全部"
   End Select
   
   'Printer.CurrentX = 9500
   Printer.CurrentX = 14000
   Printer.CurrentY = m_iPrint
   Printer.Print "頁　　次：" & str(m_intPage)
   m_iPrint = m_iPrint + 300
   
   'Printer.CurrentX = 5500 - Printer.TextWidth("年度：")
   Printer.CurrentX = 8000 - Printer.TextWidth("年度：")
   Printer.CurrentY = m_iPrint
   Printer.Print "年度：" & Text2
   m_iPrint = m_iPrint + 300
   
   If Text3 <> "" Or Text4 <> "" Then
      'Printer.CurrentX = 5500 - Printer.TextWidth("收據號碼：")
      Printer.CurrentX = 8000 - Printer.TextWidth("收據號碼：")
      Printer.CurrentY = m_iPrint
      Printer.Print "收據號碼：" & Text3 & "－" & Text4
      m_iPrint = m_iPrint + 300
   End If
   
   If Text7 <> "" Or Text8 <> "" Then
      'Printer.CurrentX = 5500 - Printer.TextWidth("客戶編號：")
      Printer.CurrentX = 8000 - Printer.TextWidth("客戶編號：")
      Printer.CurrentY = m_iPrint
      Printer.Print "客戶編號：" & Text7 & "－" & Text8
      m_iPrint = m_iPrint + 300
   End If
        
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = m_iPrint
   Printer.Print String(160, "-")
   m_iPrint = m_iPrint + 300
   
   'Modify by Morgan 2009/9/2 調整欄位順序
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = m_iPrint
   Printer.Print "手開收據號"
        
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = m_iPrint
   Printer.Print "電腦收據號"
   
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = m_iPrint
   Printer.Print "服務費"
   
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = m_iPrint
   Printer.Print "規費"
   
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = m_iPrint
   Printer.Print "收據抬頭"
   
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = m_iPrint
   Printer.Print "備註"
   
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = m_iPrint
   Printer.Print "智權人員"
   
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = m_iPrint
   Printer.Print "客戶編號"
   
   Printer.CurrentX = PLeft(8)
   Printer.CurrentY = m_iPrint
   Printer.Print "公司別"
   
'Remove by Morgan 2006/12/21
'   Printer.CurrentX = PLeft(8)
'   Printer.CurrentY = m_iPrint
'   Printer.Print "作廢日期"
   
   m_iPrint = m_iPrint + 300
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = m_iPrint
   Printer.Print String(160, "-")
   m_iPrint = m_iPrint + 300
End Sub

'Add by Morgan 2005/11/17
Private Sub cmdPrint_Click()
   Dim stSales As String
   
   With Adodc1.Recordset
      If .RecordCount > 0 Then
         GetPrintLeft
         m_intPage = 0
         .MoveFirst
         Do While Not .EOF
            'Add by Morgan 2006/12/18 作廢的不印
            If Val("" & .Fields("DT")) > 0 Then GoTo NextRec
            'Modify by Morgan 2006/12/7 要依智權人員跳頁--辜
            If stSales <> .Fields("ST02") Then
               If m_intPage <> 0 Then
                  m_iPrint = m_iPrint + 300
                  Printer.CurrentX = PLeft(0)
                  Printer.CurrentY = m_iPrint
                  Printer.Print String(118, "=")
                  Printer.NewPage
               End If
               PrintHead
            Else
               m_iPrint = m_iPrint + 300
               If m_iPrint > 14000 Then
                  Printer.CurrentX = PLeft(0)
                  Printer.CurrentY = m_iPrint
                  Printer.Print String(160, "-")
                  Printer.NewPage
                  PrintHead
               End If
            End If
            stSales = .Fields("ST02")
            'end 2006/12/7
            
            'Modify by Morgan 2009/9/2 調整欄位順序
            '手開收據號
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = m_iPrint
            Printer.Print "" & .Fields("NO")
            '電腦收據號
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = m_iPrint
            Printer.Print "" & .Fields("a0m02")
            '服務費
            Printer.CurrentX = PLeft(3) - Printer.TextWidth("　") - Printer.TextWidth(Format(" " & .Fields("SFee"), DDollar))
            Printer.CurrentY = m_iPrint
            Printer.Print Format("" & .Fields("SFee"), DDollar)
            '規費
            Printer.CurrentX = PLeft(4) - Printer.TextWidth("　") - Printer.TextWidth(Format(" " & .Fields("OFee"), DDollar))
            Printer.CurrentY = m_iPrint
            Printer.Print Format("" & .Fields("OFee"), DDollar)
            '收據抬頭
            Printer.CurrentX = PLeft(4)
            Printer.CurrentY = m_iPrint
            Printer.Print convForm("" & .Fields("CUSTOMER"), 34)
            '備註
            Printer.CurrentX = PLeft(5)
            Printer.CurrentY = m_iPrint
            Printer.Print convForm("" & .Fields("MEMO2"), 45)
            '智權人員
            Printer.CurrentX = PLeft(6)
            Printer.CurrentY = m_iPrint
            Printer.Print convForm("" & .Fields("ST02"), 6)
            '客戶編號
            Printer.CurrentX = PLeft(7)
            Printer.CurrentY = m_iPrint
            Printer.Print "" & .Fields("CUSTNO")
            '公司別
            Printer.CurrentX = PLeft(8)
            Printer.CurrentY = m_iPrint
            Printer.Print "" & .Fields("COMP")
            
'            If Not IsNull(.Fields("MEMO")) Then
'               strExc(9) = .Fields("MEMO")
'               If IsNumeric(strExc(9)) Then
'                  strExc(9) = Format(strExc(9), DDollar)
'               End If
'               Printer.CurrentX = PLeft(4) - Printer.TextWidth("　") - Printer.TextWidth(strExc(9))
'               Printer.CurrentY = m_iPrint
'               Printer.Print strExc(9)
'            End If
'end 2009/9/2

'Remove by Morgan 2006/12/21 已作廢的不用印
'            Printer.CurrentX = PLeft(8)
'            Printer.CurrentY = m_iPrint
'            Printer.Print Format("" & .Fields("DT"), "###/##/##")
NextRec:
            .MoveNext
         Loop
         m_iPrint = m_iPrint + 300
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = m_iPrint
         Printer.Print String(160, "=")
         
         m_iPrint = m_iPrint + 300
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = m_iPrint
         Printer.Print "合計"
            
         Printer.CurrentX = PLeft(3) - Printer.TextWidth("　") - Printer.TextWidth(Text9.Text)
         Printer.CurrentY = m_iPrint
         Printer.Print Text9.Text
         
         Printer.CurrentX = PLeft(4) - Printer.TextWidth("　") - Printer.TextWidth(Text10.Text)
         Printer.CurrentY = m_iPrint
         Printer.Print Text10.Text
         
         Printer.EndDoc
         ShowPrintOk
      Else
          MsgBox MsgText(28), , MsgText(5)
      End If
   End With
End Sub

Private Sub Command1_Click()
   Dim rstClone As New ADODB.Recordset
   Dim strMsg As String, iCount As Integer
   Dim strPreComp As String, strPreNo As String
   
      If TypeName(Adodc1.Recordset) = "Nothing" Then Exit Sub
      If Adodc1.Recordset.RecordCount > 0 Then
      Set rstClone = Adodc1.Recordset.Clone
      '2005/11/9 MODIFY BY SONIA
      'rstClone.Sort = " SALES,COMP"
      rstClone.Sort = " COMP"
      '2005/11/9 END
      rstClone.MoveFirst
      strPreComp = ""
      Do While Not rstClone.EOF
         'If "" & rstClone.Fields("SALES") = Text1 Then   '2005/11/9 CANCEL BY SONIA
            If "" & rstClone.Fields("COMP") = strPreComp Then
               If "" & rstClone.Fields("NO") <> strPreNo Then
                  iCount = iCount + 1
                  strPreNo = "" & rstClone.Fields("NO")
               End If
            Else
               If iCount > 0 Then
                  strMsg = strMsg & Space(5) & IIf(strPreComp = "", "  ", strPreComp) & "公司：" & iCount & " 張" & Space(5) & vbCrLf
               End If
               iCount = 1
               strPreNo = "" & rstClone.Fields("NO")
               strPreComp = "" & rstClone.Fields("COMP")
            End If
         'End If                                          '2005/11/9 CANCEL BY SONIA
         rstClone.MoveNext
      Loop
      If iCount > 0 Then
         strMsg = strMsg & Space(5) & IIf(strPreComp = "", "  ", strPreComp) & "公司：" & iCount & " 張" & Space(5) & vbCrLf
      End If
      If strMsg <> "" Then
         MsgBox strMsg, vbOKOnly, "合計張數"
      End If
   End If
   Set rstClone = Nothing
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
   PUB_InitForm Me, 9500, 5500, strBackPicPath2
   Combo1.Clear
   Combo1.AddItem "手開收據編號"
   Combo1.AddItem "客戶編號+手開收據編號"
   Combo1.AddItem "智權人員編號+手開收據編號"
   Combo1.AddItem "智權人員編號+客戶編號+手開收據編號"
   Combo1.ListIndex = 0
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc1290 = Nothing
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
   Screen.MousePointer = vbHourglass

   Dim stCon As String, stVTable As String, stConST As String
   
On Error GoTo Checking

   stCon = " to_number(substr(A1.a0k01,5))<2000"
   If Text2 <> "" Then
      'Modify by Morgan 2006/9/25
      'stCon = stCon & " and to_number(substr(A1.a0k01,2,3))=" & Val(Text2)
      If Text5 = "1" Then
         stCon = stCon & " and to_number(substr(A1.a0k01,2,3))=" & Val(Text2)
      ElseIf Text5 = "2" Then
         stCon = stCon & " and to_number(A1.a0k16)=" & Val(Text2)
      Else
         stCon = stCon & " and (to_number(substr(A1.a0k01,2,3))=" & Val(Text2) & " or to_number(A1.a0k16)=" & Val(Text2) & ")"
      End If
   End If
   If Text1 <> "" Then
      stCon = stCon & " and A1.a0k20='" & Text1 & "'"
   End If
   If Text3 <> "" Then
      stCon = stCon & " and A1.a0k01>='" & Text3 & "'"
   End If
   If Text4 <> "" Then
      stCon = stCon & " and A1.a0k01<='" & Text4 & "'"
   End If
   
   'Add by Sindy 2010/12/27 N.不含已作廢收據
   If Text11 = "N" Then
      stCon = stCon & " and (A1.a0k09 is null or A1.a0k09=0) "
   End If
   '2010/12/27 End
   
   '2005/11/9 ADD BY SONIA
   If Text5 = "1" Then
      stCon = stCon & " and A2.a0k01 IS NULL "
   End If
   If Text5 = "2" Then
      stCon = stCon & " and A2.a0k01 IS NOT NULL "
   End If
   '2005/11/9 END
   
   'Add by Morgan 2006/12/20
   If Text6 <> "" Then
      stCon = stCon & " and a0l02>=" & Val(Text6) & "0101 AND a0l02<=" & Val(Text6) & "1231"
   End If
   'end 2006/12/20
   
   'Add by Morgan 2007/4/18
   If Text7 <> "" Then
      stCon = stCon & " and A2.A0K03>='" & Text7 & "'"
   End If
   If Text8 <> "" Then
      stCon = stCon & " and A2.A0K03<='" & Text8 & "'"
   End If
   'end 2007/4/18
      
   'Modify by Morgan 2008/9/15 控制分所只能查該所資料
   stConST = ""
   If pub_strUserOffice <> "1" Then
      stConST = " and st06='" & pub_strUserOffice & "'"
   End If
   
   'Modify by Morgan 2007/4/18 加服務費,規費分開
   'strSQL = " select A1.a0k11 AS COMP,A1.a0k01 AS NO,ST02 ,A1.a0k09 AS DT,a0m02 ,A2.A0K06+A2.A0K07 AS AMOUNT,A2.A0K04 AS CUSTOMER,A1.A0K20 AS SALES,A2.A0K03 CUSTNO,A1.A0K06+A1.A0K07 AS MEMO" & _
            " from acc0k0 A1,acc0m0,acc0l0,ACC0K0 A2,STAFF where " & stCon & " and a0m03(+)=A1.a0k01  AND A1.A0K20=ST01(+) AND A0M02=A2.A0K01(+) AND A0L01(+)=A0M01"
            
   'Modify by Morgan 2008/2/12 收據金額=收款-退費
   'Modified by Morgan 2011/11/25 考慮多案合併收據情形電腦收據是否合併改抓acc0j0
   'stVTable = "select a1u02,sum(a1u08) SFeeX,sum(a1u10) OFeeX,sum(a1u04) SFeeR,sum(a1u05) OFeeR from acc1u0 where a1u02 in (" & _
      " select A2.A0k01 from acc0k0 A1,acc0m0,acc0l0,ACC0K0 A2" & _
      " where a0m03(+)=A1.a0k01 AND A0M02=A2.A0K01(+) AND A0L01(+)=A0M01 and " & stCon & ") group by a1u02"
   stVTable = "select a1u02,sum(a1u08) SFeeX,sum(a1u10) OFeeX,sum(a1u04) SFeeR,sum(a1u05) OFeeR" & _
      ",sum(decode(a0j07,'Y',0,nvl(a1u05,0)-nvl(a1u10,0))) OFee" & _
      " from acc1u0,acc0j0 where a1u02 in ( select A2.A0k01 from acc0k0 A1,acc0m0,acc0l0,ACC0K0 A2" & _
      " where a0m03(+)=A1.a0k01 AND A0M02=A2.A0K01(+) AND A0L01(+)=A0M01 and " & stCon & ")" & _
      " and a0j01(+)=a1u03 and a0j13(+)=a1u02 group by a1u02"
      
   'Modify by Morgan 2007/8/14 一張收據可能會分兩次收款，所以加 distinct-->Ex:E09600319
   'Modify by Morgan 2007/11/26 服務費規費抓已收金額
   'strSQL = " select distinct A1.a0k11 AS COMP,A1.a0k01 AS NO,ST02 ,A1.a0k09 AS DT,a0m02 ,A2.A0K06+A2.A0K07 AS AMOUNT,A2.A0K04 AS CUSTOMER,A1.A0K20 AS SALES,A2.A0K03 CUSTNO,A1.A0K06+A1.A0K07 AS MEMO" & _
   '         ",A2.A0k06-SFeeX+DECODE(A2.A0K30,'Y',A2.A0k07-OFeeX,0) SFee,DECODE(A2.A0K30,'Y',0,A2.A0k07-OFeeX) OFee from acc0k0 A1,acc0m0,acc0l0,ACC0K0 A2,STAFF,(" & stVTable & ") X where " & stCon & " and a0m03(+)=A1.a0k01  AND A1.A0K20=ST01(+) AND A0M02=A2.A0K01(+) AND A0L01(+)=A0M01 and a1u02(+)=A2.a0k01"
   'Modify by Morgan 2009/9/2 +MEMO2
   'Modified by Morgan 2011/11/25 考慮多案合併收據情形電腦收據是否合併改抓acc0j0
   'strSql = " select distinct A1.a0k11 AS COMP,A1.a0k01 AS NO,ST02 ,A1.a0k09 AS DT,a0m02 ,A2.A0K06+A2.A0K07 AS AMOUNT,A2.A0K04 AS CUSTOMER,A1.A0K20 AS SALES,A2.A0K03 CUSTNO,A1.A0K06+A1.A0K07 AS MEMO,A1.A0K08 MEMO2" & _
      ",SFeeR-SFeeX+DECODE(A2.A0K30,'Y',OFeeR-OFeeX,0) SFee,DECODE(A2.A0K30,'Y',0,OFeeR-OFeeX) OFee from acc0k0 A1,acc0m0,acc0l0,ACC0K0 A2,STAFF,(" & stVTable & ") X where " & stCon & " and a0m03(+)=A1.a0k01  AND A1.A0K20=ST01(+) AND A0M02=A2.A0K01(+) AND A0L01(+)=A0M01 and a1u02(+)=A2.a0k01" & stConST
   strSql = " select distinct A1.a0k11 AS COMP,A1.a0k01 AS NO,ST02 ,A1.a0k09 AS DT,a0m02 ,A2.A0K06+A2.A0K07 AS AMOUNT,A2.A0K04 AS CUSTOMER,A1.A0K20 AS SALES,A2.A0K03 CUSTNO,A1.A0K06+A1.A0K07 AS MEMO,A1.A0K08 MEMO2" & _
      ",SFeeR-SFeeX+OFeeR-OFeeX-OFee SFee,OFee from acc0k0 A1,acc0m0,acc0l0,ACC0K0 A2,STAFF,(" & stVTable & ") X where " & stCon & " and a0m03(+)=A1.a0k01  AND A1.A0K20=ST01(+) AND A0M02=A2.A0K01(+) and A0L01(+)=A0M01 and a1u02(+)=A2.a0k01" & stConST
   
   'end 2007/11/26
   'end 2007/4/18
   
   
   'Modify by Morgan 2006/12/20
'   '2005/11/9 ADD BY SONIA
'   If Text1 <> "" Then
'      strSQL = strSQL & " ORDER BY COMP,SALES,NO,A0M02"
'   Else
'      strSQL = strSQL & " ORDER BY COMP,NO,A0M02"
'   End If
'   '2005/11/9 END
   Select Case Combo1.ListIndex
      Case 0 '收據編號
         strSql = strSql & " Order by NO"
      Case 1 '客戶編號+收據編號
         strSql = strSql & " Order by CUSTNO,NO"
      Case 2 '智權人員編號+收據編號
         strSql = strSql & " Order by SALES,NO"
      Case Else
         strSql = strSql & " Order by NO"
   End Select
   'end 2006/12/20
   
   
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoRecordset
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount = 0 Then
      Adodc1.Recordset.Close
      MsgBox MsgText(28), , MsgText(5)
      Command1.Enabled = False
      Text9 = "": Text10 = ""
   Else
      Command1.Enabled = True
      ShowSum 'Add by Morgan 2007/4/18
   End If
   
Checking:
   Screen.MousePointer = vbDefault
   If Err.Number <> 0 Then MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1.Text = ""
   lblSalesName = ""
   Text2.Text = Left(strSrvDate(1), 4) - 1911
   Text3.Text = ""
   Text4.Text = ""
   Text6.Text = ""
   Command1.Enabled = False
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   FormCheck = True
   If Text2 = "" Then
      Text2.SetFocus
      FormCheck = False
      MsgBox "年度不可空白！", vbExclamation, "資料檢查"
      Exit Function
   End If
   
'Remove by Morgan 2006/9/25 可用年度控制;未使用控制收據碼開頭,已使用控制扣繳年度,空白取or
'   If Text1 <> "" Then
'      Exit Function
'   ElseIf Text3 = "" And Text4 = "" Then
'      Text1.SetFocus
'      FormCheck = False
'      MsgBox "智權人員與收據號碼不可同時空白！", vbExclamation, "資料檢查"
'      Exit Function
'   ElseIf Text3 = "" Then
'      Text3.SetFocus
'      FormCheck = False
'      MsgBox "收據號碼條件不完整！", vbExclamation, "資料檢查"
'      Exit Function
'   ElseIf Text4 = "" Then
'      Text4.SetFocus
'      FormCheck = False
'      MsgBox "收據號碼條件不完整！", vbExclamation, "資料檢查"
'      Exit Function
'   End If
End Function

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         If FormCheck Then
            Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
            Screen.MousePointer = vbHourglass
            AdodcRefresh
            Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Text1_Change()
   If Text1 <> "" Then
      lblSalesName = StaffQuery(Text1)
   Else
      lblSalesName = ""
   End If
End Sub

Private Sub Text1_GotFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   'Text1.IMEMode = 2
   CloseIme
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2010/12/27
Private Sub Text11_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text2_GotFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   'Text2.IMEMode = 2
   CloseIme
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii < Asc("0") And KeyAscii > Asc("9") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_GotFocus()
   If Text3.Text <> "" Then
      Text4.Text = Text3.Text
   End If
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text6_GotFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   'Text6.IMEMode = 2
   CloseIme
   TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii < Asc("0") And KeyAscii > Asc("9") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text8
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text8_GotFocus()
   If Text7.Text <> "" Then
      Text8.Text = Text7.Text
   End If
   TextInverse Text8
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub ShowSum()
   Dim lngSFee As Long, lngOFee As Long
   
   Text9 = "": Text10 = ""
   If Adodc1.Recordset.RecordCount > 0 Then
      Set RsTemp = Adodc1.Recordset.Clone
      With RsTemp
      .MoveFirst
      Do While Not .EOF
         lngSFee = lngSFee + Val("" & .Fields("SFee"))
         lngOFee = lngOFee + Val("" & .Fields("OFee"))
         .MoveNext
      Loop
      End With
      Text9 = Format(lngSFee, "#,##0")
      Text10 = Format(lngOFee, "#,##0")
   End If
End Sub
'Remove by Lydia 2018/08/24 與basQuery重複
'Private Function convForm(ByVal p_InStr As String, ByVal p_Num As Integer, Optional ByVal p_Char As String = " ") As String
'   convForm = StrConv(LeftB(StrConv(p_InStr & String(p_Num, p_Char), vbFromUnicode), p_Num), vbUnicode)
'End Function
