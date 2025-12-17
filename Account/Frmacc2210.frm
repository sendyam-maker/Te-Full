VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc2210 
   AutoRedraw      =   -1  'True
   Caption         =   "國外代理人帳目查詢"
   ClientHeight    =   5532
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   9096
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5532
   ScaleWidth      =   9096
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1500
      TabIndex        =   2
      Top             =   330
      Width           =   1572
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3420
      TabIndex        =   3
      Top             =   330
      Width           =   1572
   End
   Begin VB.CommandButton cmdDizhang 
      Caption         =   "帳款處理情形歷史記錄"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5727
      TabIndex        =   37
      Top             =   1290
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1500
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1290
      Width           =   612
   End
   Begin VB.TextBox Text17 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5430
      TabIndex        =   33
      Top             =   4980
      Width           =   1455
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5430
      TabIndex        =   29
      Top             =   4650
      Width           =   1455
   End
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7380
      TabIndex        =   28
      Top             =   4650
      Width           =   1545
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      TabIndex        =   12
      Top             =   4650
      Width           =   1605
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Frmacc2210.frx":0000
      Left            =   1260
      List            =   "Frmacc2210.frx":0002
      TabIndex        =   11
      Top             =   4650
      Width           =   1605
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2210.frx":0004
      Height          =   3000
      Left            =   30
      TabIndex        =   24
      Top             =   1600
      Width           =   8865
      _ExtentX        =   15642
      _ExtentY        =   5292
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   16
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   23
      BeginProperty Column00 
         DataField       =   "A1K28"
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
         DataField       =   "Map"
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
         DataField       =   "CaseNo"
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
      BeginProperty Column03 
         DataField       =   "DocNo"
         Caption         =   "單據編號"
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
         DataField       =   "DocDate"
         Caption         =   "單據日期"
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
      BeginProperty Column05 
         DataField       =   "Currency"
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
      BeginProperty Column06 
         DataField       =   "Famount"
         Caption         =   "外幣金額"
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
      BeginProperty Column08 
         DataField       =   "Namount"
         Caption         =   "台幣金額"
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
         DataField       =   "Close"
         Caption         =   "結清"
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
      BeginProperty Column10 
         DataField       =   "Tamount"
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
      BeginProperty Column11 
         DataField       =   "Oamount"
         Caption         =   "溢收金額"
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
      BeginProperty Column12 
         DataField       =   "DNno"
         Caption         =   "代理人D/N No."
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
         DataField       =   "Nation"
         Caption         =   "國籍"
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
         DataField       =   "FagentNo"
         Caption         =   "代理人編號"
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
         DataField       =   "A1K27"
         Caption         =   "列印對象"
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
      BeginProperty Column16 
         DataField       =   "A1K30"
         Caption         =   "已收金額(台幣)"
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
      BeginProperty Column17 
         DataField       =   "A1K10"
         Caption         =   "美金對台幣匯率"
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
      BeginProperty Column18 
         DataField       =   "A1K12"
         Caption         =   "作廢日期"
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
      BeginProperty Column19 
         DataField       =   "A1K25"
         Caption         =   "銷帳編號"
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
      BeginProperty Column20 
         DataField       =   "CusNo"
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
      BeginProperty Column21 
         DataField       =   "CusName"
         Caption         =   "客戶名稱"
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
      BeginProperty Column22 
         DataField       =   "PA161"
         Caption         =   "特殊出名公司"
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
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1188.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1632.189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1188.284
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   972.284
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   4380.095
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   1488.189
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            ColumnWidth     =   540.284
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1451.906
         EndProperty
         BeginProperty Column13 
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1128.189
         EndProperty
         BeginProperty Column15 
         EndProperty
         BeginProperty Column16 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column17 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column18 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column19 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column20 
         EndProperty
         BeginProperty Column21 
         EndProperty
         BeginProperty Column22 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "單據內容"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   7170
      TabIndex        =   23
      Top             =   60
      Width           =   1212
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3360
      TabIndex        =   10
      Top             =   4980
      Width           =   1605
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7620
      TabIndex        =   9
      Top             =   4980
      Width           =   1275
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1260
      TabIndex        =   8
      Top             =   4980
      Width           =   1605
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1500
      MaxLength       =   1
      TabIndex        =   6
      Top             =   980
      Width           =   612
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3420
      TabIndex        =   1
      Top             =   0
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1500
      TabIndex        =   0
      Top             =   0
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1500
      TabIndex        =   4
      Top             =   660
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.6
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
      Left            =   3420
      TabIndex        =   5
      Top             =   660
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.6
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
      Left            =   3420
      Top             =   1410
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   550
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
   Begin VB.Label Label19 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3180
      TabIndex        =   39
      Top             =   330
      Width           =   255
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "客戶編號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   38
      Top             =   330
      Width           =   1275
   End
   Begin VB.Label LblCompany_t 
      BackStyle       =   0  '透明
      Caption         =   "公司別："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   36
      Top             =   1290
      Width           =   1275
   End
   Begin VB.Label LblCompany 
      BackStyle       =   0  '透明
      Caption         =   "(1:專利商標 2:智權公司 空白:全部)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2190
      TabIndex        =   35
      Top             =   1290
      Width           =   3255
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "未收規費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4980
      TabIndex        =   34
      Top             =   4950
      Width           =   495
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "未收"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   32
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "CF"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4980
      TabIndex        =   31
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "未付"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6930
      TabIndex        =   30
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "外幣合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   27
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "FC"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   26
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "(*作廢、@有折讓、$銷帳、>付款中)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5055
      TabIndex        =   25
      Top             =   660
      Width           =   4305
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "未收"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   22
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "CF已付"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6930
      TabIndex        =   21
      Top             =   5040
      Width           =   705
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "FC"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   20
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "台幣合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   5040
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   -120
      Top             =   4800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "(1.FC往來 2.FC未收 3.CF往來 4.CF未付 5.往來 6.未收未付)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2190
      TabIndex        =   18
      Top             =   980
      Width           =   6495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "查詢資料："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   17
      Top             =   980
      Width           =   1275
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3180
      TabIndex        =   16
      Top             =   660
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "往來日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   15
      Top             =   660
      Width           =   1275
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3180
      TabIndex        =   14
      Top             =   30
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "代理人編號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   13
      Top             =   30
      Width           =   1275
   End
End
Attribute VB_Name = "Frmacc2210"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/09 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

'2006/3/28 整理
Public adoacc1k0 As New ADODB.Recordset
Public adoacc150 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adotmp2210 As New ADODB.Recordset
Dim strSql As String
Dim strWhere(5) As String
Dim bolCus As Boolean 'Add by Amy 2017/02/17 是否下客戶編號條件

Private Sub Command2_Click()
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   
   'Add by Morgan 2010/4/20
   If Adodc1.Recordset.BOF Or Adodc1.Recordset.EOF Then
      MsgBox "尚未點選資料！"
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   tool3_enabled
   strFormLink = Name
   'Modify by Morgan 2006/8/18 要去掉尾巴的符號
   'modify by sonia  2015/4/20 加去掉尾巴的>符號
   strItemNo = Replace(Replace(Replace(Replace(Adodc1.Recordset.Fields("DocNo").Value, "*", ""), "@", ""), "$", ""), ">", "")
   Select Case Mid(Adodc1.Recordset.Fields("DocNo").Value, 1, 1)
      Case MsgText(815)
         'strItemNo = Adodc1.Recordset.Fields("DocNo").Value
         Frmacc2211.Show
      Case MsgText(808)
         'strItemNo = Adodc1.Recordset.Fields("DocNo").Value
         Frmacc2212.Show
      'Modify by Morgan 2004/9/2 國外抵帳單編號 V 的也要
      'Case MsgText(812)
      Case MsgText(812), MsgText(813)
         'strItemNo = Adodc1.Recordset.Fields("DocNo").Value
         Frmacc2213.Show
      Case MsgText(814)
         'strItemNo = Adodc1.Recordset.Fields("DocNo").Value
         Frmacc2214.Show
      '2011/10/31 add by sonia
      Case MsgText(817)    'FC/CF 抵帳編號
         Frmacc2216.Show
      '其他單據先不顯示但要控制,否則無法繼續操作
      Case Else
         MsgBox "此單據顯示畫面尚未完成, 請通知電腦中心！"
         Screen.MousePointer = vbDefault
         Exit Sub
      '2011/10/31 end
   End Select
   
   Me.Enabled = False
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
'edit by nickc 2007/02/08
'   If IsObject(mdiMain) Then
'      mdiMain.toolshow
'   End If
   Dim formCnt As Integer
   For formCnt = 0 To Forms.Count - 1
       If UCase(Forms(formCnt).Name) = "MDIMAIN" Then
             Forms(formCnt).ToolShow
             Exit For
       End If
   Next
   strFormLink = ""
   strFormName = Name
   tool13_enabled 'Add by Amy 2015/08/31
   'Add by Amy 2015/12/04 查詢完後返回此畫面預設往來日期
   If Text1 <> MsgText(601) And Text2 <> MsgText(601) Then MaskEdBox1.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   'Modified by Lydia 2021/12/09 改成模組
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   Me.Width = 9045
'   Me.Height = 5700 '5400
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
   'Modify by Amy 2023/08/18 原:W9045 H5800
   PUB_InitForm Me, 9190, 5970, strBackPicPath2
   'end 2021/12/09
   
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   OpenTable
   StatusView MsgText(98)
   'Add By Sindy 2014/9/5 增加公司別
   If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M31" Then
      LblCompany_t.Visible = True
      LblCompany.Visible = True
      Text4.Visible = True
      'Add by Amy 2017/02/17 先讓財務測
      Label16.Visible = True
      Label19.Visible = True
      Text5.Visible = True
      Text6.Visible = True
   Else
      LblCompany_t.Visible = False
      LblCompany.Visible = False
      Text4.Visible = False
      'Add by Amy 2017/02/17 先讓財務測
      Label16.Visible = False
      Label19.Visible = False
      Text5.Visible = False
      Text6.Visible = False
   End If
   '2014/9/5 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc2210 = Nothing
End Sub

'Add by Amy 2015/08/31 for 查詢帶入值用
Private Sub Text1_Change()
    If Len(Text1) < 6 And Text1 = MsgText(601) Then Exit Sub
    If Text1.Tag = MsgText(601) Then Exit Sub
    
    Text1_LostFocus
    Text1.Tag = MsgText(601)
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus()
    '2009/6/2 MODIFY BY SONIA 預設尾碼999
    'Me.Text2.Text = Me.Text1.Text
    'Modify By Sindy 2014/8/11 999=>ZZZ
    'If Text1.Text <> "" Then Text2.Text = Left(Me.Text1.Text, 6) & "999"
    If Text1.Text <> "" Then Text2.Text = Left(Me.Text1.Text, 6) & "ZZZ"
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    If Len(Text1) = 6 Then
        Text1 = AfterZero(Text1)
    End If
    If Me.Text1.Text <> "" Then
        Me.Text1.Text = Left(Me.Text1.Text & "000000000", 9)
    End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
   CloseIme
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
    If Len(Text2) = 6 Then
        Text2 = AfterZero(Text2)
        Text2.Text = Left(Text2, 8) & "Z"    'add by sonia 2017/4/19 第9碼改為Z,因為有更名前的資料Y51562002
    End If
    If Me.Text2.Text <> "" Then
        Me.Text2.Text = Left(Me.Text2.Text & "000000000", 9)
    End If
   'add by sonia 2017/4/19 第9碼改為Z,因為有更名前的資料Y51562002
   If Mid(Text2, 9, 1) <> "Z" Then
      Text2.Text = Left(Text2, 8) & "Z"
      MsgBox ("可能會有變更名稱前的資料,故代理人編號迄號的第9碼自動改為 Z！")
   End If
   'end 2017/4/19
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
Dim intCounter As Integer
Dim strSelect1 As String
Dim strSelect2 As String
Dim strSelect3 As String
'Add By Cheng 2003/02/11
Dim strCaseNo As String '組合本所案號
'Add by Amy 2017/02/17
Dim strUpd1 As String, strUpd2 As String
   
On Error GoTo Checking

   bolCus = False 'Add by Amy 2017/02/17
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/21 清除查詢印表記錄檔欄位
   If Trim(Text1) <> "" Or Trim(Text2) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1 & Text1 & "-" & Text2 'Add By Sindy 2010/12/21
   End If
   If (MaskEdBox1.Text <> MsgText(29) And MaskEdBox1.Text <> "") Or _
      (MaskEdBox2.Text <> MsgText(29) And MaskEdBox2.Text <> "") Then
      pub_QL05 = pub_QL05 & ";" & Label3 & MaskEdBox1 & "-" & MaskEdBox2 'Add By Sindy 2010/12/21
   End If
   If Trim(Text3) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label5 & Text3 & Label6 'Add By Sindy 2010/12/21
   End If
   'Add by Amy 2017/02/17 +客戶編號
   If Trim(Text5) <> "" Or Trim(Text6) <> "" Then
        bolCus = True
        pub_QL05 = pub_QL05 & ";" & Label16 & Text5 & "-" & Text6
   End If
   '更新特殊出名公司
   strUpd1 = "Update accrpt2210 Set PA161=(" & _
                                "Select Decode(pa161,'T','專利商標','J','智權公司',pa161) From Patent Where pa01=Substr(CaseNo,1,length(CaseNo)-12) And pa02=Substr(CaseNo,(length(CaseNo)-12)+2,6) " & _
                                "And pa03=Substr(CaseNo,(length(CaseNo)-5)+2,1) and pa04=Substr(CaseNo,(length(CaseNo)-3)+2,2) " & _
                     "Union Select Decode(tm130,'J','智權公司',tm130) From Trademark Where tm01=Substr(CaseNo,1,length(CaseNo)-12) And tm02=Substr(CaseNo,(length(CaseNo)-12)+2,6) " & _
                                "And tm03=Substr(CaseNo,(length(CaseNo)-5)+2,1) and tm04=Substr(CaseNo,(length(CaseNo)-3)+2,2) " & _
                     "Union Select Decode(lc48,'J','智權公司',lc48) From Lawcase Where lc01=Substr(CaseNo,1,length(CaseNo)-12) And lc02=Substr(CaseNo,(length(CaseNo)-12)+2,6) " & _
                                "And lc03=Substr(CaseNo,(length(CaseNo)-5)+2,1) and lc04=Substr(CaseNo,(length(CaseNo)-3)+2,2) " & _
                     "Union Select Decode(sp85,'J','智權公司',sp85) From Servicepractice Where sp01=Substr(CaseNo,1,length(CaseNo)-12) And sp02=Substr(CaseNo,(length(CaseNo)-12)+2,6) " & _
                                "And sp03=Substr(CaseNo,(length(CaseNo)-5)+2,1) and sp04=Substr(CaseNo,(length(CaseNo)-3)+2,2) " & _
                    ") Where id='" & strUserNum & "' "
   'end 2017/02/17
   
   strSql = ""
   For intCounter = 0 To 5
      strWhere(intCounter) = ""
   Next intCounter
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   '代理人編號(起)
   If Text1 <> "" Then
      strSelect1 = strSelect1 & " a1k03 >= '" & Text1 & "' "
      strSelect2 = strSelect2 & " a1k28 >= '" & Text1 & "' "
      strSelect3 = strSelect3 & " a1k27 >= '" & Text1 & "' "
      strWhere(1) = strWhere(1) & " and decode(a0y18, 1, a0y07, 2, a0y08, a0y09) >= '" & Text1 & "'"
      strWhere(2) = strWhere(2) & " and a1503 >= '" & Text1 & "'"
      strWhere(3) = strWhere(3) & " and a1803 >= '" & Text1 & "'"
      strWhere(4) = strWhere(4) & " and a1603 >= '" & Text1 & "'"
   End If
   '代理人編號(迄)
   If Text2 <> "" Then
      strSelect1 = strSelect1 & " and a1k03 <= '" & Text2 & "' "
      strSelect2 = strSelect2 & " and a1k28 <= '" & Text2 & "' "
      strSelect3 = strSelect3 & " and a1k27 <= '" & Text2 & "' "
      strWhere(1) = strWhere(1) & " and decode(a0y18, 1, a0y07, 2, a0y08, a0y09) <= '" & Text2 & "'"
      strWhere(2) = strWhere(2) & " and a1503 <= '" & Text2 & "'"
      strWhere(3) = strWhere(3) & " and a1803 <= '" & Text2 & "'"
      strWhere(4) = strWhere(4) & " and a1603 <= '" & Text2 & "'"
   End If
   '往來起日
   If MaskEdBox1.Text <> MsgText(29) And MaskEdBox1.Text <> "" Then
      strWhere(0) = strWhere(0) & " and a1k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      strWhere(1) = strWhere(1) & " and a0y02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      strWhere(2) = strWhere(2) & " and a1502 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      strWhere(3) = strWhere(3) & " and NVL(a1b03,A1802) >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      strWhere(4) = strWhere(4) & " and a1602 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   '往來止日
   If MaskEdBox2.Text <> MsgText(29) And MaskEdBox2.Text <> "" Then
      strWhere(0) = strWhere(0) & " and a1k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      strWhere(1) = strWhere(1) & " and a0y02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      strWhere(2) = strWhere(2) & " and a1502 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      strWhere(3) = strWhere(3) & " and NVL(a1b03,A1802) <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      strWhere(4) = strWhere(4) & " and a1602 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If Text1 <> "" Or Text2 <> "" Then
      strWhere(0) = strWhere(0) & " and ((" & strSelect1 & ") or (" & strSelect2 & ") or (" & strSelect3 & "))"
   End If
   'Modify By Sindy 2010/3/12 增加a1k30,a1k10,a1k12,a1k25,strUserNum
   '2011/9/5 modify by sonia 增加USDamount
   'Modify By Sindy 2012/8/24 有關收款單的規費計算原抓取 sum(A1K09) 改為 sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09)))
   'Modify By Sindy 2014/3/24 +,null as PA161
   Select Case Text3
      Case "1" 'FC往來
'         strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,(a1k08 - nvl(a1k06, 0))) USDamount from acc1k0, fagent, nation, acc140 where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0)
'         '2010/6/29 MODIFY BY SONIA 抵帳幣別不可抓A1K18請款幣別X09901818
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1k17 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, NVL(A1H03,A1I05) as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k08 - nvl(a1k06, 0)) * nvl(a1g02, 0) as Namount, null as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '2' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount from acc1k0, fagent, nation, acc140, acc1g0, acc1h0, acc1i0 where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and fa10 = na01 (+) and a1k01 = a1403 (+) and a1k17 = a1g01 (+) and a1k17 = a1h01 (+) and a1k17 = a1i01 (+) and a1k17 is not null" & strWhere(0)
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,(a1k08 - nvl(a1k06, 0))) USDamount from acc1k0, customer, nation, acc140 where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and cu10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0)
'         '2010/6/29 MODIFY BY SONIA 抵帳幣別不可抓A1K18請款幣別X09901818
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a1k17 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, NVL(A1H03,A1I05) as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k08 - nvl(a1k06, 0)) * nvl(a1g02, 0) as Namount, null as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '2' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount from acc1k0, customer, nation, acc140, acc1g0, acc1h0, acc1i0 where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and cu10 = na01 (+) and a1k01 = a1403 (+) and a1k17 = a1g01 (+) and a1k17 = a1h01 (+) and a1k17 = a1i01 (+) and a1k17 is not null" & strWhere(0)
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, a1k29 as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a0z02 as Map, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc0y0, fagent, nation, acc0z0, acc1k0 where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 and fa10 = na01 (+) and a0y01 = a0z01 and a0z02=a1k01(+) " & strWhere(1) & _
'         " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a0z02, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k29,a1k30,a1k09 "
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, a1k29 as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a0z02 as Map, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc0y0, customer, nation, acc0z0, acc1k0 where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = cu01 and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = cu02 and cu10 = na01 (+) and a0y01 = a0z01 and a0z02=a1k01(+) " & strWhere(1) & _
'         " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(cu05, nvl(cu06, cu04)), na03, a0y01, a0y02, a0y03, a0y06, a0z02, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k29,a1k30,a1k09 "
'請款單-fagent
         'Modify By Sindy 2013/1/14
         strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc1k0, fagent, nation, acc140 where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0)
'抵帳收款-fagent
         '2010/6/29 MODIFY BY SONIA 抵帳幣別不可抓A1K18請款幣別X09901818
         'Modified by Lydia 2025/01/13 union select>> union ALL select; ex: 因為U11400365顯示的欄位值都一樣，在經過Union Select後只顯示一筆
         strSql = strSql & " union ALL select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1k17 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, NVL(A1H03,A1I05) as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k08 - nvl(a1k31, 0)) * nvl(a1g02, 0) as Namount, null as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '2' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc1k0, fagent, nation, acc140, acc1g0, acc1h0, acc1i0 where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and fa10 = na01 (+) and a1k01 = a1403 (+) and a1k17 = a1g01 (+) and a1k17 = a1h01 (+) and a1k17 = a1i01 (+) and a1k17 is not null" & strWhere(0)
'請款單-customer
         strSql = strSql & " union ALL select a1k03 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc1k0, customer, nation, acc140 where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and cu10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0)
'抵帳收款-customer
         '2010/6/29 MODIFY BY SONIA 抵帳幣別不可抓A1K18請款幣別X09901818
         strSql = strSql & " union ALL select a1k03 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a1k17 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, NVL(A1H03,A1I05) as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k08 - nvl(a1k31, 0)) * nvl(a1g02, 0) as Namount, null as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '2' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc1k0, customer, nation, acc140, acc1g0, acc1h0, acc1i0 where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and cu10 = na01 (+) and a1k01 = a1403 (+) and a1k17 = a1g01 (+) and a1k17 = a1h01 (+) and a1k17 = a1i01 (+) and a1k17 is not null" & strWhere(0)
'收款-fagent
         strSql = strSql & " union ALL select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, a1k29 as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a0z02 as Map, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc0y0, fagent, nation, acc0z0, acc1k0 where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 and fa10 = na01 (+) and a0y01 = a0z01 and a0z02=a1k01(+) " & strWhere(1) & _
         " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a0z02, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k29,a1k30,a1k09 "
'收款-customer
         strSql = strSql & " union ALL select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, a1k29 as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a0z02 as Map, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc0y0, customer, nation, acc0z0, acc1k0 where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = cu01 and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = cu02 and cu10 = na01 (+) and a0y01 = a0z01 and a0z02=a1k01(+) " & strWhere(1) & _
         " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(cu05, nvl(cu06, cu04)), na03, a0y01, a0y02, a0y03, a0y06, a0z02, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k29,a1k30,a1k09 "
         'end 2025/01/13 'Modified by Lydia 2025/01/13 union select>> union ALL select; ex: 因為U11400365顯示的欄位值都一樣，在經過Union Select後只顯示一筆
      Case "2" 'FC未收
'Modify by Morgan 2005/1/14 作廢銷帳的不要
'         strSQL = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27 " & _
'                  " from acc1k0, fagent, nation, acc140 where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
'         strSQL = strSQL & " union select a1k03 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27  " & _
'                  " from acc1k0, customer, nation, acc140 where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and cu10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
         '2007/12/10 modify by sonia X09607651分次收款,婧瑄說台幣金額扣除已收金額,外幣改為台幣金額扣除已收金額/請款匯率
         'strSQL = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27 " & _
         '         " from acc1k0, fagent, nation, acc140 where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1401 is null and a1k12 is null"
         'strSQL = strSQL & " union select a1k03 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27  " & _
         '         " from acc1k0, customer, nation, acc140 where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and cu10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1401 is null and a1k12 is null"
         'Modify By Sindy 2012/6/25 X09607651分次收款未收金額部分改回原程式寫法,以X10003936測試
'         strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                  " from acc1k0, fagent, nation, acc140 where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1401 is null and a1k12 is null"
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                  " from acc1k0, customer, nation, acc140 where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and cu10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1401 is null and a1k12 is null"
'         strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                  " from acc1k0, fagent, nation, acc140 where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1401 is null and a1k12 is null"
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                  " from acc1k0, customer, nation, acc140 where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and cu10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1401 is null and a1k12 is null"
'         '2012/6/25 End
'         '2007/12/10 end
'         'Add By Sindy 2012/8/14 +部分收款時要同時帶出收款資料
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, a1k29 as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a0z02 as Map, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc0y0, fagent, nation, acc0z0, acc1k0 where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 and fa10 = na01 (+) and a0y01 = a0z01 and a0z02=a1k01(+) " & strWhere(1) & _
'                  " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a0z02, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k29,a1k30,a1k09 "
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, a1k29 as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a0z02 as Map, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc0y0, customer, nation, acc0z0, acc1k0 where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = cu01 and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = cu02 and cu10 = na01 (+) and a0y01 = a0z01 and a0z02=a1k01(+) " & strWhere(1) & _
'                  " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(cu05, nvl(cu06, cu04)), na03, a0y01, a0y02, a0y03, a0y06, a0z02, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k29,a1k30,a1k09 "
'         '2012/8/14 End
'請款單-fagent
         'Modify By Sindy 2013/1/14
         strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161" & _
                  " from acc1k0, fagent, nation, acc140 where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1401 is null and a1k12 is null"
'請款單-customer
         'Modified by Lydia 2025/01/13 union select>> union ALL select; ex: 因為U11400365顯示的欄位值都一樣，在經過Union Select後只顯示一筆
         strSql = strSql & " union ALL select a1k03 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161" & _
                  " from acc1k0, customer, nation, acc140 where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and cu10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1401 is null and a1k12 is null"
         '2012/6/25 End
         '2007/12/10 end
'Add By Sindy 2012/8/14 +部分收款時要同時帶出收款資料
'部分收款-fagent
         strSql = strSql & " union ALL select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, a1k29 as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a0z02 as Map, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc0y0, fagent, nation, acc0z0, acc1k0 where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 and fa10 = na01 (+) and a0y01 = a0z01 and a0z02=a1k01(+) " & strWhere(1) & _
                  " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a0z02, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k29,a1k30,a1k09 "
'部分收款-customer
         strSql = strSql & " union ALL select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, a1k29 as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a0z02 as Map, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc0y0, customer, nation, acc0z0, acc1k0 where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = cu01 and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = cu02 and cu10 = na01 (+) and a0y01 = a0z01 and a0z02=a1k01(+) " & strWhere(1) & _
                  " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(cu05, nvl(cu06, cu04)), na03, a0y01, a0y02, a0y03, a0y06, a0z02, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k29,a1k30,a1k09 "
         'end 2025/01/13 'Modified by Lydia 2025/01/13 union select>> union ALL select; ex: 因為U11400365顯示的欄位值都一樣，在經過Union Select後只顯示一筆
'2012/8/14 End
'2005/1/14 end
      Case "3" 'CF往來
'         '組合本所案號
'         strCaseNo = "Decode(length(axf03),10,substr(axf03,1,1)||'-'||substr(axf03,2,6)||'-'||substr(axf03,8,1)||'-'||substr(axf03,9,2),11,substr(axf03,1,2)||'-'||substr(axf03,3,6)||'-'||substr(axf03,9,1)||'-'||substr(axf03,10,2),12,substr(axf03,1,3)||'-'||substr(axf03,4,6)||'-'||substr(axf03,10,1)||'-'||substr(axf03,11,2),axf03)"
'         strSql = "select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc150,acc151, fagent, nation where a1501=axf01(+) and substr(a1503, 1, 8) = fa01 and substr(a1503, 9, 1) = fa02 and fa10 = na01 (+)" & strWhere(2)
'         strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, axf04 as Famount, a1506 * nvl(a1g03, 0) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc150,acc151, fagent, nation, acc1g0, acc1h0, acc1i0 where a1501=axf01(+) and substr(a1503, 1, 8) = fa01 and substr(a1503, 9, 1) = fa02 and fa10 = na01 (+) and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2)
'         strSql = strSql & " union select a1503 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc150,acc151, customer, nation where a1501=axf01(+) and substr(a1503, 1, 8) = cu01 and substr(a1503, 9, 1) = cu02 and cu10 = na01 (+)" & strWhere(2)
'         strSql = strSql & " union select a1503 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, axf04 as Famount, a1506 * nvl(a1g03, 0) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc150,acc151, customer, nation, acc1g0, acc1h0, acc1i0 where a1501=axf01(+) and substr(a1503, 1, 8) = cu01 and substr(a1503, 9, 1) = cu02 and cu10 = na01 (+) and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2)
'         '2007/11/27 MODIFY BY SONIA 因W09601477之U09604115有二案號,故將台幣金額由A1905改為axf04*a1906
'         'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, axf04 as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28,'' as a1k27 from acc190, acc180, fagent, nation, acc150, acc151, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 and substr(a1803, 9, 1) = fa02 and fa10 = na01 (+) and a1902=a1501(+) and a1501=axf01(+) and a1908=a1b01(+)" & strWhere(3) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1504,axf03, axf04 "
'         'strSQL = strSQL & " union select a1803 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, axf04 as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27 from acc190, acc180, customer, nation, acc150, acc151, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = cu01 and substr(a1803, 9, 1) = cu02 and cu10 = na01 (+) and a1902=a1501(+) and a1501=axf01(+) and a1908=a1b01(+)" & strWhere(3) & " group by a1803, nvl(cu05, nvl(cu06, cu04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1504,axf03, axf04 "
'         'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
'         'Modify By Sindy 2012/12/6 a1902="U"時抓取acc150,acc151資料
'         strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, axf04 as Famount, sum(axf04*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28,'' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, fagent, nation, acc150, acc151, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 and substr(a1803, 9, 1) = fa02 and fa10 = na01 (+) and a1902=a1501 and a1501=axf01(+) and a1908=a1b01(+) and a1908 is not null" & strWhere(3) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1504,axf03, axf04 "
'         strSql = strSql & " union select a1803 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, axf04 as Famount, sum(axf04*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, customer, nation, acc150, acc151, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = cu01 and substr(a1803, 9, 1) = cu02 and cu10 = na01 (+) and a1902=a1501 and a1501=axf01(+) and a1908=a1b01(+) and a1908 is not null" & strWhere(3) & " group by a1803, nvl(cu05, nvl(cu06, cu04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1504,axf03, axf04 "
'         'Modify By Sindy 2012/12/6 a1902="V"時抓取acc160,acc161資料
'         strCaseNo = "Decode(length(axg03),10,substr(axg03,1,1)||'-'||substr(axg03,2,6)||'-'||substr(axg03,8,1)||'-'||substr(axg03,9,2),11,substr(axg03,1,2)||'-'||substr(axg03,3,6)||'-'||substr(axg03,9,1)||'-'||substr(axg03,10,2),12,substr(axg03,1,3)||'-'||substr(axg03,4,6)||'-'||substr(axg03,10,1)||'-'||substr(axg03,11,2),axg03)"
'         strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, axg04 * (-1) as Famount, sum(axg04*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28,'' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, fagent, nation, acc160, acc161, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 and substr(a1803, 9, 1) = fa02 and fa10 = na01 (+) and a1902=a1601 and a1601=axg01(+) and a1908=a1b01(+) and a1908 is not null" & strWhere(3) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1604,axg03, axg04 "
'         strSql = strSql & " union select a1803 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, axg04 * (-1) as Famount, sum(axg04*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, customer, nation, acc160, acc161, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = cu01 and substr(a1803, 9, 1) = cu02 and cu10 = na01 (+) and a1902=a1601 and a1601=axg01(+) and a1908=a1b01(+) and a1908 is not null" & strWhere(3) & " group by a1803, nvl(cu05, nvl(cu06, cu04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1604,axg03, axg04 "
'         '2012/12/6 End
'         '2007/11/27 end
'         strCaseNo = "Decode(length(axg03),10,substr(axg03,1,1)||'-'||substr(axg03,2,6)||'-'||substr(axg03,8,1)||'-'||substr(axg03,9,2),11,substr(axg03,1,2)||'-'||substr(axg03,3,6)||'-'||substr(axg03,9,1)||'-'||substr(axg03,10,2),12,substr(axg03,1,3)||'-'||substr(axg03,4,6)||'-'||substr(axg03,10,1)||'-'||substr(axg03,11,2),axg03)"
'         'Modify By Sindy 2012/8/13 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'         'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, axg04 as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation where axg01(+)=a1601 and substr(a1603, 1, 8) = fa01(+) and substr(a1603, 9, 1) = fa02(+) and fa10 = na01 (+)" & strWhere(4)
'         strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, axg04 * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation where axg01(+)=a1601 and substr(a1603, 1, 8) = fa01(+) and substr(a1603, 9, 1) = fa02(+) and fa10 = na01 (+)" & strWhere(4)
'         '2012/8/13 End
         'Modify By Sindy 2013/1/14
'帳單ACC150
         '組合本所案號
         strCaseNo = "Decode(length(axf03),10,substr(axf03,1,1)||'-'||substr(axf03,2,6)||'-'||substr(axf03,8,1)||'-'||substr(axf03,9,2),11,substr(axf03,1,2)||'-'||substr(axf03,3,6)||'-'||substr(axf03,9,1)||'-'||substr(axf03,10,2),12,substr(axf03,1,3)||'-'||substr(axf03,4,6)||'-'||substr(axf03,10,1)||'-'||substr(axf03,11,2),axf03)"
         '2014/11/27 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
         'strSql = "select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc150,acc151, fagent, nation where a1501=axf01(+) and substr(a1503, 1, 8) = fa01 and substr(a1503, 9, 1) = fa02 and fa10 = na01 (+)" & strWhere(2)
         strSql = "select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc150,acc151, fagent, nation, acc190 " & _
                  "where a1501=axf01(+) and substr(a1503, 1, 8) = fa01 and substr(a1503, 9, 1) = fa02 and fa10 = na01 (+) and a1501=a1902(+) " & strWhere(2)
         '2014/11/27 end
'抵帳資料
         'Modified by Lydia 2025/01/13 union select>> union ALL select; ex: 因為U11400365顯示的欄位值都一樣，在經過Union Select後只顯示一筆
         strSql = strSql & " union ALL select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, axf04 as Famount, a1506 * nvl(a1g03, 0) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc150,acc151, fagent, nation, acc1g0, acc1h0, acc1i0 where a1501=axf01(+) and substr(a1503, 1, 8) = fa01 and substr(a1503, 9, 1) = fa02 and fa10 = na01 (+) and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2)
'2014/11/27 CANCEL BY SONIA 帳單不會來自於CUSOTMER
'         strSql = strSql & " union select a1503 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc150,acc151, customer, nation where a1501=axf01(+) and substr(a1503, 1, 8) = cu01 and substr(a1503, 9, 1) = cu02 and cu10 = na01 (+)" & strWhere(2)
'         strSql = strSql & " union select a1503 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, axf04 as Famount, a1506 * nvl(a1g03, 0) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc150,acc151, customer, nation, acc1g0, acc1h0, acc1i0 where a1501=axf01(+) and substr(a1503, 1, 8) = cu01 and substr(a1503, 9, 1) = cu02 and cu10 = na01 (+) and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2)
'2014/11/27 END

'帳單結匯(有匯票號)
         '2007/11/27 MODIFY BY SONIA 因W09601477之U09604115有二案號,故將台幣金額由A1905改為axf04*a1906
         'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, axf04 as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28,'' as a1k27 from acc190, acc180, fagent, nation, acc150, acc151, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 and substr(a1803, 9, 1) = fa02 and fa10 = na01 (+) and a1902=a1501(+) and a1501=axf01(+) and a1908=a1b01(+)" & strWhere(3) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1504,axf03, axf04 "
         'strSQL = strSQL & " union select a1803 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, axf04 as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27 from acc190, acc180, customer, nation, acc150, acc151, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = cu01 and substr(a1803, 9, 1) = cu02 and cu10 = na01 (+) and a1902=a1501(+) and a1501=axf01(+) and a1908=a1b01(+)" & strWhere(3) & " group by a1803, nvl(cu05, nvl(cu06, cu04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1504,axf03, axf04 "
         'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
         'Modify By Sindy 2012/12/6 a1902="U"時抓取acc150,acc151資料
'FAGENT
         'Modified by Lydia 2025/01/13 union select>> union ALL select; ex: 因為U11400365顯示的欄位值都一樣，在經過Union Select後只顯示一筆
         strSql = strSql & " union ALL select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, axf04 as Famount, sum(axf04*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28,'' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, fagent, nation, acc150, acc151, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 and substr(a1803, 9, 1) = fa02 and fa10 = na01 (+) and a1902=a1501 and a1501=axf01(+) and a1908=a1b01(+) and a1908 is not null" & strWhere(3) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1504,axf03, axf04 "
'CUSTOMER(國外暫收款退費O單據)
         'Modified by Lydia 2025/01/13 union select>> union ALL select; ex: 因為U11400365顯示的欄位值都一樣，在經過Union Select後只顯示一筆
         strSql = strSql & " union ALL select a1803 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, axf04 as Famount, sum(axf04*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, customer, nation, acc150, acc151, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = cu01 and substr(a1803, 9, 1) = cu02 and cu10 = na01 (+) and a1902=a1501 and a1501=axf01(+) and a1908=a1b01(+) and a1908 is not null" & strWhere(3) & " group by a1803, nvl(cu05, nvl(cu06, cu04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1504,axf03, axf04 "

'抵帳單ACC160
         strCaseNo = "Decode(length(axg03),10,substr(axg03,1,1)||'-'||substr(axg03,2,6)||'-'||substr(axg03,8,1)||'-'||substr(axg03,9,2),11,substr(axg03,1,2)||'-'||substr(axg03,3,6)||'-'||substr(axg03,9,1)||'-'||substr(axg03,10,2),12,substr(axg03,1,3)||'-'||substr(axg03,4,6)||'-'||substr(axg03,10,1)||'-'||substr(axg03,11,2),axg03)"
         'Modify By Sindy 2012/8/13 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
         'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, axg04 as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation where axg01(+)=a1601 and substr(a1603, 1, 8) = fa01(+) and substr(a1603, 9, 1) = fa02(+) and fa10 = na01 (+)" & strWhere(4)
         '2014/11/27 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
         'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, axg04 * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation where axg01(+)=a1601 and substr(a1603, 1, 8) = fa01(+) and substr(a1603, 9, 1) = fa02(+) and fa10 = na01 (+)" & strWhere(4)
         'Modified by Lydia 2025/01/13 union select>> union ALL select; ex: 因為U11400365顯示的欄位值都一樣，在經過Union Select後只顯示一筆
         strSql = strSql & " union ALL select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, axg04 * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, ACC190 " & _
                           " where axg01(+)=a1601 AND A1601=A1902(+) and substr(a1603, 1, 8) = fa01(+) and substr(a1603, 9, 1) = fa02(+) and fa10 = na01 (+)" & strWhere(4)
         '2014/11/27 end
         '2012/8/13 End
'抵帳單結匯(有匯票號)
         'Modify By Sindy 2012/12/6 a1902="V"時抓取acc160,acc161資料
         strCaseNo = "Decode(length(axg03),10,substr(axg03,1,1)||'-'||substr(axg03,2,6)||'-'||substr(axg03,8,1)||'-'||substr(axg03,9,2),11,substr(axg03,1,2)||'-'||substr(axg03,3,6)||'-'||substr(axg03,9,1)||'-'||substr(axg03,10,2),12,substr(axg03,1,3)||'-'||substr(axg03,4,6)||'-'||substr(axg03,10,1)||'-'||substr(axg03,11,2),axg03)"
         'Modified by Lydia 2025/01/13 union select>> union ALL select; ex: 因為U11400365顯示的欄位值都一樣，在經過Union Select後只顯示一筆
         strSql = strSql & " union ALL select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, axg04 * (-1) as Famount, sum(axg04*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28,'' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, fagent, nation, acc160, acc161, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 and substr(a1803, 9, 1) = fa02 and fa10 = na01 (+) and a1902=a1601 and a1601=axg01(+) and a1908=a1b01(+) and a1908 is not null" & strWhere(3) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1604,axg03, axg04 "
         strSql = strSql & " union ALL select a1803 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, axg04 * (-1) as Famount, sum(axg04*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, customer, nation, acc160, acc161, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = cu01 and substr(a1803, 9, 1) = cu02 and cu10 = na01 (+) and a1902=a1601 and a1601=axg01(+) and a1908=a1b01(+) and a1908 is not null" & strWhere(3) & " group by a1803, nvl(cu05, nvl(cu06, cu04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1604,axg03, axg04 "
         '2012/12/6 End
         '2007/11/27 end
      Case "4" 'CF未付
'         '組合本所案號
'         strCaseNo = "Decode(length(axf03),10,substr(axf03,1,1)||'-'||substr(axf03,2,6)||'-'||substr(axf03,8,1)||'-'||substr(axf03,9,2),11,substr(axf03,1,2)||'-'||substr(axf03,3,6)||'-'||substr(axf03,9,1)||'-'||substr(axf03,10,2),12,substr(axf03,1,3)||'-'||substr(axf03,4,6)||'-'||substr(axf03,10,1)||'-'||substr(axf03,11,2),axf03)"
''Modify by Morgan 2005/1/14 作廢銷帳的不要
''         strSQL = "select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27 from acc150,acc151, fagent, nation where a1501=axf01(+) and substr(a1503, 1, 8) = fa01 and substr(a1503, 9, 1) = fa02 and fa10 = na01 (+)" & strWhere(2) & " and (a1520 = 0 or a1520 is null)"
''         strSQL = strSQL & " union select a1503 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27 from acc150, acc151, customer, nation where a1501=axf01(+) and substr(a1503, 1, 8) = cu01 and substr(a1503, 9, 1) = cu02 and cu10 = na01 (+)" & strWhere(2) & " and (a1520 = 0 or a1520 is null)"
''Modify by Morgan 2009/3/23 +已抵帳的也不要
''         strSQL = "select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27 from acc150,acc151, fagent, nation where a1501=axf01(+) and substr(a1503, 1, 8) = fa01 and substr(a1503, 9, 1) = fa02 and fa10 = na01 (+)" & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1507 is null"
''         strSQL = strSQL & " union select a1503 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27 from acc150, acc151, customer, nation where a1501=axf01(+) and substr(a1503, 1, 8) = cu01 and substr(a1503, 9, 1) = cu02 and cu10 = na01 (+)" & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1507 is null"
'         strSql = "select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc150,acc151, fagent, nation where a1501=axf01(+) and substr(a1503, 1, 8) = fa01 and substr(a1503, 9, 1) = fa02 and fa10 = na01 (+)" & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null"
'         strSql = strSql & " union select a1503 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc150, acc151, customer, nation where a1501=axf01(+) and substr(a1503, 1, 8) = cu01 and substr(a1503, 9, 1) = cu02 and cu10 = na01 (+)" & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null"
'         '2006/2/9 ADD BY SONIA
'         strCaseNo = "Decode(length(axg03),10,substr(axg03,1,1)||'-'||substr(axg03,2,6)||'-'||substr(axg03,8,1)||'-'||substr(axg03,9,2),11,substr(axg03,1,2)||'-'||substr(axg03,3,6)||'-'||substr(axg03,9,1)||'-'||substr(axg03,10,2),12,substr(axg03,1,3)||'-'||substr(axg03,4,6)||'-'||substr(axg03,10,1)||'-'||substr(axg03,11,2),axg03)"
'         'Modify By Sindy 2012/8/13 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'         'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, axg04 as Famount, 0 as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation, ACC190 where A1607 IS NULL AND axg01(+)=a1601 AND A1601=A1902(+) AND A1901 IS NULL and substr(a1603, 1, 8) = fa01(+) and substr(a1603, 9, 1) = fa02(+) and fa10 = na01 (+)" & strWhere(4)
'         'Modify By Sindy 2012/9/11 + and a1908 is null
'         strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, axg04 * (-1) as Famount, 0 as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation, ACC190 where A1607 IS NULL AND axg01(+)=a1601 AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null and substr(a1603, 1, 8) = fa01(+) and substr(a1603, 9, 1) = fa02(+) and fa10 = na01 (+)" & strWhere(4)
'         '2012/8/13 End
'         '2006/2/9 END
''2005/1/14 end
         'Modify By Sindy 2013/1/14
'帳單ACC150
         '組合本所案號
         strCaseNo = "Decode(length(axf03),10,substr(axf03,1,1)||'-'||substr(axf03,2,6)||'-'||substr(axf03,8,1)||'-'||substr(axf03,9,2),11,substr(axf03,1,2)||'-'||substr(axf03,3,6)||'-'||substr(axf03,9,1)||'-'||substr(axf03,10,2),12,substr(axf03,1,3)||'-'||substr(axf03,4,6)||'-'||substr(axf03,10,1)||'-'||substr(axf03,11,2),axf03)"
         'Modify by Morgan 2005/1/14 作廢銷帳的不要
         'strSQL = "select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27 from acc150,acc151, fagent, nation where a1501=axf01(+) and substr(a1503, 1, 8) = fa01 and substr(a1503, 9, 1) = fa02 and fa10 = na01 (+)" & strWhere(2) & " and (a1520 = 0 or a1520 is null)"
         'strSQL = strSQL & " union select a1503 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27 from acc150, acc151, customer, nation where a1501=axf01(+) and substr(a1503, 1, 8) = cu01 and substr(a1503, 9, 1) = cu02 and cu10 = na01 (+)" & strWhere(2) & " and (a1520 = 0 or a1520 is null)"
         'Modify by Morgan 2009/3/23 +已抵帳的也不要
         'strSQL = "select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27 from acc150,acc151, fagent, nation where a1501=axf01(+) and substr(a1503, 1, 8) = fa01 and substr(a1503, 9, 1) = fa02 and fa10 = na01 (+)" & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1507 is null"
         'strSQL = strSQL & " union select a1503 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27 from acc150, acc151, customer, nation where a1501=axf01(+) and substr(a1503, 1, 8) = cu01 and substr(a1503, 9, 1) = cu02 and cu10 = na01 (+)" & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1507 is null"
         '2014/11/27 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
         'strSql = "select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc150,acc151, fagent, nation where a1501=axf01(+) and substr(a1503, 1, 8) = fa01 and substr(a1503, 9, 1) = fa02 and fa10 = na01 (+)" & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null"
         strSql = "select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc150,acc151, fagent, nation, acc190 " & _
                  "where a1501=axf01(+) and substr(a1503, 1, 8) = fa01 and substr(a1503, 9, 1) = fa02 and fa10 = na01 (+) and a1501=a1902(+) " & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null"
         '2014/11/27 end
'2014/11/27 CANCEL BY SONIA 帳單不會來自於CUSOTMER
'         strSql = strSql & " union select a1503 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc150, acc151, customer, nation where a1501=axf01(+) and substr(a1503, 1, 8) = cu01 and substr(a1503, 9, 1) = cu02 and cu10 = na01 (+)" & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null"
'2014/11/27 END

'抵帳單ACC160
         '2006/2/9 ADD BY SONIA
         strCaseNo = "Decode(length(axg03),10,substr(axg03,1,1)||'-'||substr(axg03,2,6)||'-'||substr(axg03,8,1)||'-'||substr(axg03,9,2),11,substr(axg03,1,2)||'-'||substr(axg03,3,6)||'-'||substr(axg03,9,1)||'-'||substr(axg03,10,2),12,substr(axg03,1,3)||'-'||substr(axg03,4,6)||'-'||substr(axg03,10,1)||'-'||substr(axg03,11,2),axg03)"
         'Modify By Sindy 2012/8/13 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
         'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, axg04 as Famount, 0 as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation, ACC190 where A1607 IS NULL AND axg01(+)=a1601 AND A1601=A1902(+) AND A1901 IS NULL and substr(a1603, 1, 8) = fa01(+) and substr(a1603, 9, 1) = fa02(+) and fa10 = na01 (+)" & strWhere(4)
         'Modify By Sindy 2012/9/11 + and a1908 is null
         '2014/11/27 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
         'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, axg04 * (-1) as Famount, 0 as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, ACC190 where A1607 IS NULL AND axg01(+)=a1601 AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null and substr(a1603, 1, 8) = fa01(+) and substr(a1603, 9, 1) = fa02(+) and fa10 = na01 (+)" & strWhere(4)
         'Modified by Lydia 2025/01/13 union select>> union ALL select; ex: 因為U11400365顯示的欄位值都一樣，在經過Union Select後只顯示一筆
         strSql = strSql & " union ALL select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, axg04 * (-1) as Famount, 0 as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, ACC190 " & _
                           " where A1607 IS NULL AND axg01(+)=a1601 AND A1601=A1902(+) and a1908 is null and substr(a1603, 1, 8) = fa01(+) and substr(a1603, 9, 1) = fa02(+) and fa10 = na01 (+)" & strWhere(4)
         '2014/11/27 end
         '2012/8/13 End
         '2006/2/9 END
'2005/1/14 end
      Case "6" '未收未付
'         '組合本所案號
'         strCaseNo = "Decode(length(axf03),10,substr(axf03,1,1)||'-'||substr(axf03,2,6)||'-'||substr(axf03,8,1)||'-'||substr(axf03,9,2),11,substr(axf03,1,2)||'-'||substr(axf03,3,6)||'-'||substr(axf03,9,1)||'-'||substr(axf03,10,2),12,substr(axf03,1,3)||'-'||substr(axf03,4,6)||'-'||substr(axf03,10,1)||'-'||substr(axf03,11,2),axf03)"
'         '2007/12/10 modify by sonia X09607651分次收款,婧瑄說台幣金額扣除已收金額,外幣改為台幣金額扣除已收金額/請款匯率
'         'strSQL = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27 " & _
'         '         " from acc1k0, fagent, nation, acc140 where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
'         'strSQL = strSQL & " union select a1k03 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27 " & _
'         '         " from acc1k0, customer, nation, acc140 where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and cu10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
'         'Modify By Sindy 2012/6/25 X09607651分次收款未收金額部分改回原程式寫法,以X10003936測試
''         strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
''                  " from acc1k0, fagent, nation, acc140 where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1401 is null and a1k12 is null"
''         strSql = strSql & " union select a1k03 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
''                  " from acc1k0, customer, nation, acc140 where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and cu10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1401 is null and a1k12 is null"
'         strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                  " from acc1k0, fagent, nation, acc140 where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1401 is null and a1k12 is null"
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                  " from acc1k0, customer, nation, acc140 where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and cu10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1401 is null and a1k12 is null"
'         '2012/6/25 end
'         '2007/12/10 end
'         '2010/10/26 modify by sonia 作廢或已抵帳都不要(前只改cf未付)
'         strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc150,acc151, fagent, nation where a1501=axf01(+) and substr(a1503, 1, 8) = fa01 and substr(a1503, 9, 1) = fa02 and fa10 = na01 (+)" & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null"
'         strSql = strSql & " union select a1503 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc150,acc151, customer, nation where a1501=axf01(+) and substr(a1503, 1, 8) = cu01 and substr(a1503, 9, 1) = cu02 and cu10 = na01 (+)" & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null"
'         '2006/2/9 ADD BY SONIA
'         strCaseNo = "Decode(length(axg03),10,substr(axg03,1,1)||'-'||substr(axg03,2,6)||'-'||substr(axg03,8,1)||'-'||substr(axg03,9,2),11,substr(axg03,1,2)||'-'||substr(axg03,3,6)||'-'||substr(axg03,9,1)||'-'||substr(axg03,10,2),12,substr(axg03,1,3)||'-'||substr(axg03,4,6)||'-'||substr(axg03,10,1)||'-'||substr(axg03,11,2),axg03)"
'         'Modify By Sindy 2012/8/13 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'         'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, axg04 as Famount, 0 as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation, ACC190 where A1607 IS NULL AND axg01(+)=a1601 AND A1601=A1902(+) AND A1901 IS NULL and substr(a1603, 1, 8) = fa01(+) and substr(a1603, 9, 1) = fa02(+) and fa10 = na01 (+)" & strWhere(4)
'         'Modify By Sindy 2012/12/6 + and a1908 is null
'         strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, axg04 * (-1) as Famount, 0 as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation, ACC190 where A1607 IS NULL AND axg01(+)=a1601 AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null and substr(a1603, 1, 8) = fa01(+) and substr(a1603, 9, 1) = fa02(+) and fa10 = na01 (+)" & strWhere(4)
'         '2012/8/13 End
'         '2006/2/9 END
'         'Add By Sindy 2012/8/14 +部分收款時要同時帶出收款資料
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, a1k29 as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a0z02 as Map, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc0y0, fagent, nation, acc0z0, acc1k0 where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 and fa10 = na01 (+) and a0y01 = a0z01 and a0z02=a1k01(+) " & strWhere(1) & _
'                  " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a0z02, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k29,a1k30,a1k09 "
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, a1k29 as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a0z02 as Map, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc0y0, customer, nation, acc0z0, acc1k0 where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = cu01 and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = cu02 and cu10 = na01 (+) and a0y01 = a0z01 and a0z02=a1k01(+) " & strWhere(1) & _
'                  " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(cu05, nvl(cu06, cu04)), na03, a0y01, a0y02, a0y03, a0y06, a0z02, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k29,a1k30,a1k09 "
'         '2012/8/14 End
         '2007/12/10 modify by sonia X09607651分次收款,婧瑄說台幣金額扣除已收金額,外幣改為台幣金額扣除已收金額/請款匯率
         'strSQL = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27 " & _
         '         " from acc1k0, fagent, nation, acc140 where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
         'strSQL = strSQL & " union select a1k03 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27 " & _
         '         " from acc1k0, customer, nation, acc140 where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and cu10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
         'Modify By Sindy 2012/6/25 X09607651分次收款未收金額部分改回原程式寫法,以X10003936測試
'         strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                  " from acc1k0, fagent, nation, acc140 where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1401 is null and a1k12 is null"
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                  " from acc1k0, customer, nation, acc140 where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and cu10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1401 is null and a1k12 is null"
'請款單-fagent
         strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161" & _
                  " from acc1k0, fagent, nation, acc140 where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1401 is null and a1k12 is null"
'請款單-customer
         'Modified by Lydia 2025/01/13 union select>> union ALL select; ex: 因為U11400365顯示的欄位值都一樣，在經過Union Select後只顯示一筆
         strSql = strSql & " union ALL select a1k03 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161" & _
                  " from acc1k0, customer, nation, acc140 where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and cu10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1401 is null and a1k12 is null"
'Add By Sindy 2012/8/14 +部分收款時要同時帶出收款資料
'部分收款-fagent
         'Modified by Lydia 2025/01/13 union select>> union ALL select; ex: 因為U11400365顯示的欄位值都一樣，在經過Union Select後只顯示一筆
         strSql = strSql & " union ALL select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, a1k29 as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a0z02 as Map, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc0y0, fagent, nation, acc0z0, acc1k0 where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 and fa10 = na01 (+) and a0y01 = a0z01 and a0z02=a1k01(+) " & strWhere(1) & _
                  " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a0z02, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k29,a1k30,a1k09 "
'部分收款-customer
         'Modified by Lydia 2025/01/13 union select>> union ALL select; ex: 因為U11400365顯示的欄位值都一樣，在經過Union Select後只顯示一筆
         strSql = strSql & " union ALL select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, a1k29 as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a0z02 as Map, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc0y0, customer, nation, acc0z0, acc1k0 where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = cu01 and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = cu02 and cu10 = na01 (+) and a0y01 = a0z01 and a0z02=a1k01(+) " & strWhere(1) & _
                  " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(cu05, nvl(cu06, cu04)), na03, a0y01, a0y02, a0y03, a0y06, a0z02, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k29,a1k30,a1k09 "
'2012/8/14 End
         '2012/6/25 end
         '2007/12/10 end
         'Modify By Sindy 2013/1/14
'帳單ACC150
         '組合本所案號
         strCaseNo = "Decode(length(axf03),10,substr(axf03,1,1)||'-'||substr(axf03,2,6)||'-'||substr(axf03,8,1)||'-'||substr(axf03,9,2),11,substr(axf03,1,2)||'-'||substr(axf03,3,6)||'-'||substr(axf03,9,1)||'-'||substr(axf03,10,2),12,substr(axf03,1,3)||'-'||substr(axf03,4,6)||'-'||substr(axf03,10,1)||'-'||substr(axf03,11,2),axf03)"
         '2010/10/26 modify by sonia 作廢或已抵帳都不要(前只改cf未付)
         '2014/11/27 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
         'strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc150,acc151, fagent, nation where a1501=axf01(+) and substr(a1503, 1, 8) = fa01 and substr(a1503, 9, 1) = fa02 and fa10 = na01 (+)" & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null"
         'Modified by Lydia 2025/01/13 union select>> union ALL select; ex: 因為U11400365顯示的欄位值都一樣，在經過Union Select後只顯示一筆
         strSql = strSql & " union ALL select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc150,acc151, fagent, nation, acc190 " & _
                           " where a1501=axf01(+) and substr(a1503, 1, 8) = fa01 and substr(a1503, 9, 1) = fa02 and fa10 = na01 (+) and a1501=a1902(+) " & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null"
         '2014/11/27 end
'2014/11/27 CANCEL BY SONIA 帳單不會來自於CUSOTMER
'         strSql = strSql & " union select a1503 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc150,acc151, customer, nation where a1501=axf01(+) and substr(a1503, 1, 8) = cu01 and substr(a1503, 9, 1) = cu02 and cu10 = na01 (+)" & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null"
'2014/11/27 END

'抵帳單ACC160
         '2006/2/9 ADD BY SONIA
         strCaseNo = "Decode(length(axg03),10,substr(axg03,1,1)||'-'||substr(axg03,2,6)||'-'||substr(axg03,8,1)||'-'||substr(axg03,9,2),11,substr(axg03,1,2)||'-'||substr(axg03,3,6)||'-'||substr(axg03,9,1)||'-'||substr(axg03,10,2),12,substr(axg03,1,3)||'-'||substr(axg03,4,6)||'-'||substr(axg03,10,1)||'-'||substr(axg03,11,2),axg03)"
         'Modify By Sindy 2012/8/13 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
         'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, axg04 as Famount, 0 as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation, ACC190 where A1607 IS NULL AND axg01(+)=a1601 AND A1601=A1902(+) AND A1901 IS NULL and substr(a1603, 1, 8) = fa01(+) and substr(a1603, 9, 1) = fa02(+) and fa10 = na01 (+)" & strWhere(4)
         'Modify By Sindy 2012/12/6 + and a1908 is null
         '2014/11/27 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
         'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, axg04 * (-1) as Famount, 0 as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, ACC190 where A1607 IS NULL AND axg01(+)=a1601 AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null and substr(a1603, 1, 8) = fa01(+) and substr(a1603, 9, 1) = fa02(+) and fa10 = na01 (+)" & strWhere(4)
         'Modified by Lydia 2025/01/13 union select>> union ALL select; ex: 因為U11400365顯示的欄位值都一樣，在經過Union Select後只顯示一筆
         strSql = strSql & " union ALL select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, axg04 * (-1) as Famount, 0 as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, ACC190 " & _
                           " where A1607 IS NULL AND axg01(+)=a1601 AND A1601=A1902(+) and a1908 is null and substr(a1603, 1, 8) = fa01(+) and substr(a1603, 9, 1) = fa02(+) and fa10 = na01 (+)" & strWhere(4)
         '2014/11/27 end
         '2012/8/13 End
         '2006/2/9 END
      Case "5", "" '往來
'         '組合本所案號
'         strCaseNo = "Decode(length(axf03),10,substr(axf03,1,1)||'-'||substr(axf03,2,6)||'-'||substr(axf03,8,1)||'-'||substr(axf03,9,2),11,substr(axf03,1,2)||'-'||substr(axf03,3,6)||'-'||substr(axf03,9,1)||'-'||substr(axf03,10,2),12,substr(axf03,1,3)||'-'||substr(axf03,4,6)||'-'||substr(axf03,10,1)||'-'||substr(axf03,11,2),axf03)"
'         strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,(a1k08 - nvl(a1k06, 0))) USDamount " & _
'                  " from acc1k0, fagent, nation, acc140 where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0)
'         '2010/6/29 MODIFY BY SONIA 抵帳幣別不可抓A1K18請款幣別X09901818
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1k17 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, NVL(A1H03,A1I05) as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k08 - nvl(a1k06, 0)) * nvl(a1g02, 0) as Namount, null as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '2' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount from acc1k0, fagent, nation, acc140, acc1g0, acc1h0, acc1i0 " & _
'                " where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and fa10 = na01 (+) and a1k01 = a1403 (+) and a1k17 = a1g01 (+) and a1k17 = a1h01 (+) and a1k17 = a1i01 (+) and a1k17 is not null" & strWhere(0)
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,(a1k08 - nvl(a1k06, 0))) USDamount " & _
'                  " from acc1k0, customer, nation, acc140 where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and cu10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0)
'         '2010/6/29 MODIFY BY SONIA 抵帳幣別不可抓A1K18請款幣別X09901818
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a1k17 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, NVL(A1H03,A1I05) as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k08 - nvl(a1k06, 0)) * nvl(a1g02, 0) as Namount, null as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '2' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount from acc1k0, customer, nation, acc140, acc1g0, acc1h0, acc1i0 " & _
'                " where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and cu10 = na01 (+) and a1k01 = a1403 (+) and a1k17 = a1g01 (+) and a1k17 = a1h01 (+) and a1k17 = a1i01 (+) and a1k17 is not null" & strWhere(0)
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a0z02 as Map, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc0y0, fagent, nation, acc0z0, acc1k0 where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02  and fa10 = na01 (+) and a0y01 = a0z01 and a0z02=a1k01(+) " & strWhere(1) & _
'         " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a0z02, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16,a1k30,a1k09 "
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a0z02 as Map, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc0y0, customer, nation, acc0z0, acc1k0 where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = cu01 and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = cu02 and cu10 = na01 (+) and a0y01 = a0z01 and a0z02=a1k01(+) " & strWhere(1) & _
'         " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(cu05, nvl(cu06, cu04)), na03, a0y01, a0y02, a0y03, a0y06, a0z02, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16,a1k30,a1k09 "
'         strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc150, acc151, fagent, nation where a1501=axf01(+) and substr(a1503, 1, 8) = fa01 and substr(a1503, 9, 1) = fa02 and fa10 = na01 (+)" & strWhere(2)
'         strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, axf04 as Famount, a1506 * nvl(a1g03, 0) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc150,acc151, fagent, nation, acc1g0, acc1h0, acc1i0 where a1501=axf01(+) and substr(a1503, 1, 8) = fa01 and substr(a1503, 9, 1) = fa02 and fa10 = na01 (+) and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2)
'         strSql = strSql & " union select a1503 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc150, acc151, customer, nation where a1501=axf01(+) and substr(a1503, 1, 8) = cu01 and substr(a1503, 9, 1) = cu02 and cu10 = na01 (+)" & strWhere(2)
'         strSql = strSql & " union select a1503 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, axf04 as Famount, a1506 * nvl(a1g03, 0) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc150,acc151, customer, nation, acc1g0, acc1h0, acc1i0 where a1501=axf01(+) and substr(a1503, 1, 8) = cu01 and substr(a1503, 9, 1) = cu02 and cu10 = na01 (+) and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2)
'         '2007/11/27 MODIFY BY SONIA 因W09601477之U09604115有二案號,故將外幣金額由A1904改為axf04*a1906,另sum(a1904)同"3"CF往來改為sum(axf04)
'         'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, acc150, acc151, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 and substr(a1803, 9, 1) = fa02 and fa10 = na01 (+) and a1902=a1501(+) and a1501=axf01(+) and a1908=a1b01(+)" & strWhere(3) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1504, axf03"
'         'strSQL = strSQL & " union select a1803 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27 from acc190, acc180, customer, nation, acc150, acc151, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = cu01 and substr(a1803, 9, 1) = cu02 and cu10 = na01 (+) and a1902=a1501(+) and a1501=axf01(+) and a1908=a1b01(+)" & strWhere(3) & " group by a1803, nvl(cu05, nvl(cu06, cu04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1504, axf03"
'         'Modify By Sindy 2012/12/6 a1902="U"時抓取acc150,acc151資料
'         strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, sum(axf04) as Famount, sum(axf04*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, fagent, nation, acc150, acc151, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 and substr(a1803, 9, 1) = fa02 and fa10 = na01 (+) and a1902=a1501 and a1501=axf01(+) and a1908=a1b01(+) and a1908 is not null" & strWhere(3) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1504, axf03"
'         strSql = strSql & " union select a1803 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, sum(axf04) as Famount, sum(axf04*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, customer, nation, acc150, acc151, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = cu01 and substr(a1803, 9, 1) = cu02 and cu10 = na01 (+) and a1902=a1501 and a1501=axf01(+) and a1908=a1b01(+) and a1908 is not null" & strWhere(3) & " group by a1803, nvl(cu05, nvl(cu06, cu04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1504, axf03"
'         'Modify By Sindy 2012/12/6 a1902="V"時抓取acc160,acc161資料
'         strCaseNo = "Decode(length(axg03),10,substr(axg03,1,1)||'-'||substr(axg03,2,6)||'-'||substr(axg03,8,1)||'-'||substr(axg03,9,2),11,substr(axg03,1,2)||'-'||substr(axg03,3,6)||'-'||substr(axg03,9,1)||'-'||substr(axg03,10,2),12,substr(axg03,1,3)||'-'||substr(axg03,4,6)||'-'||substr(axg03,10,1)||'-'||substr(axg03,11,2),axg03)"
'         strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, sum(axg04) * (-1) as Famount, sum(axg04*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, fagent, nation, acc160, acc161, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 and substr(a1803, 9, 1) = fa02 and fa10 = na01 (+) and a1902=a1601 and a1601=axg01(+) and a1908=a1b01(+) and a1908 is not null" & strWhere(3) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1604, axg03"
'         strSql = strSql & " union select a1803 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, sum(axg04) * (-1) as Famount, sum(axg04*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, customer, nation, acc160, acc161, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = cu01 and substr(a1803, 9, 1) = cu02 and cu10 = na01 (+) and a1902=a1601 and a1601=axg01(+) and a1908=a1b01(+) and a1908 is not null" & strWhere(3) & " group by a1803, nvl(cu05, nvl(cu06, cu04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1604, axg03"
'         '2012/12/6 End
'         '2007/11/27 end
'         strCaseNo = "Decode(length(axg03),10,substr(axg03,1,1)||'-'||substr(axg03,2,6)||'-'||substr(axg03,8,1)||'-'||substr(axg03,9,2),11,substr(axg03,1,2)||'-'||substr(axg03,3,6)||'-'||substr(axg03,9,1)||'-'||substr(axg03,10,2),12,substr(axg03,1,3)||'-'||substr(axg03,4,6)||'-'||substr(axg03,10,1)||'-'||substr(axg03,11,2),axg03)"
'         'Modify By Sindy 2012/8/13 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'         'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, axg04 as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation where axg01(+)=a1601 and substr(a1603, 1, 8) = fa01(+) and substr(a1603, 9, 1) = fa02(+) and fa10 = na01 (+)" & strWhere(4)
'         strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, axg04 * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation where axg01(+)=a1601 and substr(a1603, 1, 8) = fa01(+) and substr(a1603, 9, 1) = fa02(+) and fa10 = na01 (+)" & strWhere(4)
'         '2012/8/13 End
'請款單-fagent
         strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161" & _
                  " from acc1k0, fagent, nation, acc140 where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0)
'抵帳收款-fagent
         '2010/6/29 MODIFY BY SONIA 抵帳幣別不可抓A1K18請款幣別X09901818
         'Modified by Lydia 2025/01/13 union select>> union ALL select; ex: 因為U11400365顯示的欄位值都一樣，在經過Union Select後只顯示一筆
         strSql = strSql & " union ALL select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1k17 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, NVL(A1H03,A1I05) as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k08 - nvl(a1k31, 0)) * nvl(a1g02, 0) as Namount, null as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '2' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc1k0, fagent, nation, acc140, acc1g0, acc1h0, acc1i0 " & _
                " where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and fa10 = na01 (+) and a1k01 = a1403 (+) and a1k17 = a1g01 (+) and a1k17 = a1h01 (+) and a1k17 = a1i01 (+) and a1k17 is not null" & strWhere(0)
'請款單-customer
         strSql = strSql & " union ALL select a1k03 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '1' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161" & _
                  " from acc1k0, customer, nation, acc140 where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and cu10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0)
'抵帳收款-customer
         '2010/6/29 MODIFY BY SONIA 抵帳幣別不可抓A1K18請款幣別X09901818
         strSql = strSql & " union ALL select a1k03 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a1k17 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, NVL(A1H03,A1I05) as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k08 - nvl(a1k31, 0)) * nvl(a1g02, 0) as Namount, null as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Map, '2' as Sort,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc1k0, customer, nation, acc140, acc1g0, acc1h0, acc1i0 " & _
                " where substr(a1k03, 1, 8) = cu01 and substr(a1k03, 9, 1) = cu02 and cu10 = na01 (+) and a1k01 = a1403 (+) and a1k17 = a1g01 (+) and a1k17 = a1h01 (+) and a1k17 = a1i01 (+) and a1k17 is not null" & strWhere(0)
'收款-fagent
         strSql = strSql & " union ALL select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a0z02 as Map, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc0y0, fagent, nation, acc0z0, acc1k0 where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02  and fa10 = na01 (+) and a0y01 = a0z01 and a0z02=a1k01(+) " & strWhere(1) & _
         " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a0z02, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16,a1k30,a1k09 "
'收款-customer
         strSql = strSql & " union ALL select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a0z02 as Map, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc0y0, customer, nation, acc0z0, acc1k0 where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = cu01 and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = cu02 and cu10 = na01 (+) and a0y01 = a0z01 and a0z02=a1k01(+) " & strWhere(1) & _
         " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(cu05, nvl(cu06, cu04)), na03, a0y01, a0y02, a0y03, a0y06, a0z02, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16,a1k30,a1k09 "
         
'帳單ACC150
         'Modify By Sindy 2013/1/14
         '組合本所案號
         strCaseNo = "Decode(length(axf03),10,substr(axf03,1,1)||'-'||substr(axf03,2,6)||'-'||substr(axf03,8,1)||'-'||substr(axf03,9,2),11,substr(axf03,1,2)||'-'||substr(axf03,3,6)||'-'||substr(axf03,9,1)||'-'||substr(axf03,10,2),12,substr(axf03,1,3)||'-'||substr(axf03,4,6)||'-'||substr(axf03,10,1)||'-'||substr(axf03,11,2),axf03)"
         '2014/11/27 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
         'strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc150, acc151, fagent, nation where a1501=axf01(+) and substr(a1503, 1, 8) = fa01 and substr(a1503, 9, 1) = fa02 and fa10 = na01 (+)" & strWhere(2)
         strSql = strSql & " union ALL select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc150, acc151, fagent, nation, acc190 " & _
                           " where a1501=axf01(+) and substr(a1503, 1, 8) = fa01 and substr(a1503, 9, 1) = fa02 and fa10 = na01 (+) and a1501=a1902(+) " & strWhere(2)
         '2014/11/27 end
'抵帳資料
         strSql = strSql & " union ALL select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, axf04 as Famount, a1506 * nvl(a1g03, 0) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc150,acc151, fagent, nation, acc1g0, acc1h0, acc1i0 where a1501=axf01(+) and substr(a1503, 1, 8) = fa01 and substr(a1503, 9, 1) = fa02 and fa10 = na01 (+) and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2)
         'end 2025/01/13 'Modified by Lydia 2025/01/13 union select>> union ALL select; ex: 因為U11400365顯示的欄位值都一樣，在經過Union Select後只顯示一筆
'2014/11/27 CANCEL BY SONIA 帳單不會來自於CUSOTMER
'         strSql = strSql & " union select a1503 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, axf04 as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc150, acc151, customer, nation where a1501=axf01(+) and substr(a1503, 1, 8) = cu01 and substr(a1503, 9, 1) = cu02 and cu10 = na01 (+)" & strWhere(2)
'         strSql = strSql & " union select a1503 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, axf04 as Famount, a1506 * nvl(a1g03, 0) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc150,acc151, customer, nation, acc1g0, acc1h0, acc1i0 where a1501=axf01(+) and substr(a1503, 1, 8) = cu01 and substr(a1503, 9, 1) = cu02 and cu10 = na01 (+) and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2)
'2014/11/27 END

'帳單結匯(有匯票號)
         '2007/11/27 MODIFY BY SONIA 因W09601477之U09604115有二案號,故將外幣金額由A1904改為axf04*a1906,另sum(a1904)同"3"CF往來改為sum(axf04)
         'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, acc150, acc151, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 and substr(a1803, 9, 1) = fa02 and fa10 = na01 (+) and a1902=a1501(+) and a1501=axf01(+) and a1908=a1b01(+)" & strWhere(3) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1504, axf03"
         'strSQL = strSQL & " union select a1803 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27 from acc190, acc180, customer, nation, acc150, acc151, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = cu01 and substr(a1803, 9, 1) = cu02 and cu10 = na01 (+) and a1902=a1501(+) and a1501=axf01(+) and a1908=a1b01(+)" & strWhere(3) & " group by a1803, nvl(cu05, nvl(cu06, cu04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1504, axf03"
         'Modify By Sindy 2012/12/6 a1902="U"時抓取acc150,acc151資料
'FAGENT
         'Modified by Lydia 2025/01/13 union select>> union ALL select; ex: 因為U11400365顯示的欄位值都一樣，在經過Union Select後只顯示一筆
         strSql = strSql & " union ALL select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, sum(axf04) as Famount, sum(axf04*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, fagent, nation, acc150, acc151, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 and substr(a1803, 9, 1) = fa02 and fa10 = na01 (+) and a1902=a1501 and a1501=axf01(+) and a1908=a1b01(+) and a1908 is not null" & strWhere(3) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1504, axf03"
'CUSTOMER(國外暫收款退費O單據)
         strSql = strSql & " union ALL select a1803 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, sum(axf04) as Famount, sum(axf04*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, customer, nation, acc150, acc151, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = cu01 and substr(a1803, 9, 1) = cu02 and cu10 = na01 (+) and a1902=a1501 and a1501=axf01(+) and a1908=a1b01(+) and a1908 is not null" & strWhere(3) & " group by a1803, nvl(cu05, nvl(cu06, cu04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1504, axf03"
         
'抵帳單ACC160
         strCaseNo = "Decode(length(axg03),10,substr(axg03,1,1)||'-'||substr(axg03,2,6)||'-'||substr(axg03,8,1)||'-'||substr(axg03,9,2),11,substr(axg03,1,2)||'-'||substr(axg03,3,6)||'-'||substr(axg03,9,1)||'-'||substr(axg03,10,2),12,substr(axg03,1,3)||'-'||substr(axg03,4,6)||'-'||substr(axg03,10,1)||'-'||substr(axg03,11,2),axg03)"
         'Modify By Sindy 2012/8/13 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
         'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, axg04 as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation where axg01(+)=a1601 and substr(a1603, 1, 8) = fa01(+) and substr(a1603, 9, 1) = fa02(+) and fa10 = na01 (+)" & strWhere(4)
         '2014/11/27 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
         'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, axg04 * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation where axg01(+)=a1601 and substr(a1603, 1, 8) = fa01(+) and substr(a1603, 9, 1) = fa02(+) and fa10 = na01 (+)" & strWhere(4)
         strSql = strSql & " union ALL select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, axg04 * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Map, '1' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, acc190 " & _
                           " where axg01(+)=a1601 AND A1601=A1902(+) and substr(a1603, 1, 8) = fa01(+) and substr(a1603, 9, 1) = fa02(+) and fa10 = na01 (+)" & strWhere(4)
         '2014/11/27 end
         '2012/8/13 End
'抵帳單結匯(有匯票號)
         'Modify By Sindy 2012/12/6 a1902="V"時抓取acc160,acc161資料
         strCaseNo = "Decode(length(axg03),10,substr(axg03,1,1)||'-'||substr(axg03,2,6)||'-'||substr(axg03,8,1)||'-'||substr(axg03,9,2),11,substr(axg03,1,2)||'-'||substr(axg03,3,6)||'-'||substr(axg03,9,1)||'-'||substr(axg03,10,2),12,substr(axg03,1,3)||'-'||substr(axg03,4,6)||'-'||substr(axg03,10,1)||'-'||substr(axg03,11,2),axg03)"
         strSql = strSql & " union ALL select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, sum(axg04) * (-1) as Famount, sum(axg04*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, fagent, nation, acc160, acc161, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 and substr(a1803, 9, 1) = fa02 and fa10 = na01 (+) and a1902=a1601 and a1601=axg01(+) and a1908=a1b01(+) and a1908 is not null" & strWhere(3) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1604, axg03"
         strSql = strSql & " union ALL select a1803 as FagentNo, nvl(cu05, nvl(cu06, cu04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, sum(axg04) * (-1) as Famount, sum(axg04*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1902 as Map, '2' as Sort," & strCaseNo & " as CaseNo, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, customer, nation, acc160, acc161, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = cu01 and substr(a1803, 9, 1) = cu02 and cu10 = na01 (+) and a1902=a1601 and a1601=axg01(+) and a1908=a1b01(+) and a1908 is not null" & strWhere(3) & " group by a1803, nvl(cu05, nvl(cu06, cu04)), na03, a1801, NVL(a1b03,A1802), a1903, a1902, a1604, axg03"
         '2012/12/6 End
         '2007/11/27 end
         'end 2025/01/13 'Modified by Lydia 2025/01/13 union select>> union ALL select; ex: 因為U11400365顯示的欄位值都一樣，在經過Union Select後只顯示一筆
   End Select
   
   Call ClearSumCol 'Add By Sindy 2012/8/14
   
   'Add By Sindy 2010/3/12
   cnnConnection.BeginTrans
   cnnConnection.Execute "delete from accrpt2210 where id='" & strUserNum & "'"
   cnnConnection.Execute "insert into accrpt2210 (FagentNo,FagentName,Nation,DocNo,DocDate,Currency,Famount,Namount,Close,Tamount," & _
                                        "Oamount,Dnno,Map,Sort,Caseno,a1k28,a1k27,a1k30,a1k10,a1k12,a1k25,ID,USDAmount,pa161) " & strSql
   'Add by Amy 2017/02/17
   '更新特殊出名公司
   cnnConnection.Execute strUpd1
   '更新客戶編號
   If bolCus = True Then
        Call UpdCusData
   Else
        strUpd2 = "Update accrpt2210 Set CusNo=(" & _
                                "Select '(1)'||pa26 From Patent Where pa01=Substr(CaseNo,1,length(CaseNo)-12) And pa02=Substr(CaseNo,(length(CaseNo)-12)+2,6) " & _
                                "And pa03=Substr(CaseNo,(length(CaseNo)-5)+2,1) and pa04=Substr(CaseNo,(length(CaseNo)-3)+2,2) And pa26 is not null " & _
                     "Union Select '(1)'||tm23 From Trademark Where tm01=Substr(CaseNo,1,length(CaseNo)-12) And tm02=Substr(CaseNo,(length(CaseNo)-12)+2,6) " & _
                                "And tm03=Substr(CaseNo,(length(CaseNo)-5)+2,1) and tm04=Substr(CaseNo,(length(CaseNo)-3)+2,2) And tm23 is not null " & _
                     "Union Select '(1)'||lc11 From Lawcase Where lc01=Substr(CaseNo,1,length(CaseNo)-12) And lc02=Substr(CaseNo,(length(CaseNo)-12)+2,6) " & _
                                "And lc03=Substr(CaseNo,(length(CaseNo)-5)+2,1) and lc04=Substr(CaseNo,(length(CaseNo)-3)+2,2) And lc11 is not null " & _
                     "Union Select '(1)'||sp08 From Servicepractice Where sp01=Substr(CaseNo,1,length(CaseNo)-12) And sp02=Substr(CaseNo,(length(CaseNo)-12)+2,6) " & _
                                "And sp03=Substr(CaseNo,(length(CaseNo)-5)+2,1) and sp04=Substr(CaseNo,(length(CaseNo)-3)+2,2) And sp08 is not null " & _
                    ") Where id='" & strUserNum & "' "
        cnnConnection.Execute strUpd2
   End If
   
   
   If bolCus = True Then
        '刪除非畫面條件之客戶編號
        strSql = "Delete accrpt2210 Where id='" & strUserNum & "' And CusNo is null "
        cnnConnection.Execute strSql
   End If
   cnnConnection.CommitTrans
   '2010/3/12 End
   'Memo 2017/02/17 改寫法原程式放至UpdSpecCmp()暫存 'Modify By Sindy 2014/9/9 逐筆讀取特殊出名公司
   
   'adoadodc1.Open strSQL & " order by Map asc, Sort asc", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2017/02/17 +顯示客戶資料
   'adoadodc1.Open "select * from accrpt2210 where id='" & strUserNum & "' " & IIf(Text4.Visible = True And Text4.Text = "2", "and PA161='智權公司'", IIf(Text4.Visible = True And Text4.Text = "1", "and (PA161 is null or PA161<>'智權公司')", "")) & " order by Map asc, Sort asc ", adoTaie, adOpenStatic, adLockBatchOptimistic 'Modify By Sindy 2010/3/12 原為adLockReadOnly
   strSql = "Select accrpt2210.*,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) as CusName From accrpt2210,Customer Where id='" & strUserNum & "' " & IIf(Text4.Visible = True And Text4.Text = "2", "and PA161='智權公司'", IIf(Text4.Visible = True And Text4.Text = "1", "and (PA161 is null or PA161<>'智權公司')", "")) & _
                " And Substr(Cusno,4,8)=cu01(+) And Substr(Cusno,12,1)=cu02(+)" & _
                " order by Map asc, Sort asc "
   adoadodc1.Open strSql, adoTaie, adOpenStatic, adLockBatchOptimistic 'Modify By Sindy 2010/3/12 原為adLockReadOnly
   Adodc1.Recordset.ReQuery
   SumShow
   If Adodc1.Recordset.RecordCount = 0 Then
      InsertQueryLog (0) 'Add By Sindy 2010/12/21
      Adodc1.Recordset.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   Else
      InsertQueryLog (Adodc1.Recordset.RecordCount) 'Add By Sindy 2010/12/21
'      'Modify by Amy 2017/02/17 原程式寫至UpdSpecCmp 'Add By Sindy 2014/3/24 +特殊出名公司

      DataGrid1.AllowUpdate = False '鎖住畫面,不可異動資料
'      '2014/3/24 End
      Adodc1.Recordset.MoveFirst
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   '2011/9/6 add by sonia
   Else
      Resume Next
      cnnConnection.RollbackTrans
   '2011/9/6 end
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
            GetDizhang Text1, Text2, True  'Add by Amy 2013/10/30 +帳款處理訊息
            QueryTable
            Screen.MousePointer = vbDefault
            Exit Sub
         Else
            MsgBox MsgText(181) & vbCrLf & "「代理人編號」為必填", , MsgText(5) 'Modify by Amy 2013/10/30 +代理人編號 為必填
         End If
   End Select
   KeyEnter KeyCode
   StatusView MsgText(98)
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

''*************************************************
''  計算並顯示合計 (請款單)
''
''*************************************************
'Public Sub SumShow1()
'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open "select sum(a1k11), sum(a1k30) from acc1k0, fagent where a1k03 = (fa01 || fa02)" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         Text13 = MsgText(601)
'      Else
'         Text13 = Format(adoaccsum.Fields(0).Value, FDollar)
'      End If
'   Else
'      Text13 = MsgText(601)
'   End If
'   adoaccsum.Close
'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open "select sum(a1k11), sum(a1k30) from acc1k0, fagent where a1k03 = (fa01 || fa02) and (a1k29 is null or a1k29 = '')" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         Text11 = MsgText(601)
'      Else
'         If IsNull(adoaccsum.Fields(1).Value) Then
'            Text11 = Format(adoaccsum.Fields(0).Value, FDollar)
'         Else
'            Text11 = Format(adoaccsum.Fields(0).Value - adoaccsum.Fields(1).Value, FDollar)
'         End If
'      End If
'   Else
'      Text11 = MsgText(601)
'   End If
'   adoaccsum.Close
'    'Add By Cheng 2004/04/16
'    '加顯示外幣FC
'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open "select sum(Nvl(a1k08,0)) from acc1k0, fagent where a1k03 = (fa01 || fa02)" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         Text7 = ""
'      Else
'         Text7 = Format(adoaccsum.Fields(0).Value, FDollar)
'      End If
'   Else
'      Text7 = ""
'   End If
'   adoaccsum.Close
'    'End
'End Sub

''*************************************************
''  計算並顯示合計 (帳單)
''
''*************************************************
'Public Sub SumShow2()
'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open "select sum(a1510), sum(a1520) from acc150, fagent where a1503 = (fa01 || fa02)" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         Text12 = MsgText(601)
'      Else
'         Text12 = Format(adoaccsum.Fields(0).Value, FDollar)
'      End If
'   Else
'      Text12 = MsgText(601)
'   End If
'   adoaccsum.Close
'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open "select sum(a1510), sum(a1520) from acc150, fagent where a1503 = (fa01 || fa02) and (a1510 - a1520) > 0" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         Text7 = MsgText(601)
'      Else
'         If IsNull(adoaccsum.Fields(1).Value) Then
'            Text7 = Format(adoaccsum.Fields(0).Value, FDollar)
'         Else
'            Text7 = Format(adoaccsum.Fields(0).Value - adoaccsum.Fields(1).Value, FDollar)
'         End If
'      End If
'   Else
'      Text7 = MsgText(601)
'   End If
'   adoaccsum.Close
'End Sub

''*************************************************
''  儲存請款資料
''
''*************************************************
'Private Sub Acc1k0Query()
'On Error GoTo Checking
'   strSql = ""
'   If Text1 <> MsgText(601) Then
'      strSql = strSql & " and a1k03 >= '" & Text1 & "'"
'   End If
'   If Text2 <> MsgText(601) Then
'      strSql = strSql & " and a1k03 <= '" & Text2 & "'"
'   End If
'   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'      strSql = strSql & " and a1k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'   End If
'   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'      strSql = strSql & " and a1k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'   End If
'   If Text3 <> MsgText(601) Then
'      Select Case Text3
'         Case "2", "6"
'            strSql = strSql & " and (a1k30 = 0 or a1k30 is null)"
'      End Select
'   End If
'   adoacc1k0.CursorLocation = adUseClient
'   adoacc1k0.Open "select * from acc1k0, fagent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+)" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'   Do While adoacc1k0.EOF = False
'      adotmp2210.AddNew
'      If IsNull(adoacc1k0.Fields("a1k03").Value) Then
'         adotmp2210.Fields("t22101").Value = Null
'      Else
'         adotmp2210.Fields("t22101").Value = adoacc1k0.Fields("a1k03").Value
'      End If
'      If IsNull(adoacc1k0.Fields("fa04").Value) Then
'         adotmp2210.Fields("t22102").Value = Null
'      Else
'         adotmp2210.Fields("t22102").Value = adoacc1k0.Fields("fa04").Value
'      End If
'      If IsNull(adoacc1k0.Fields("fa10").Value) Then
'         adotmp2210.Fields("t22103").Value = Null
'      Else
'         adotmp2210.Fields("t22103").Value = NationQuery(adoacc1k0.Fields("fa10").Value, 1)
'      End If
'      adotmp2210.Fields("t22104").Value = adoacc1k0.Fields("a1k01").Value
'      If IsNull(adoacc1k0.Fields("a1k02").Value) Then
'         adotmp2210.Fields("t22105").Value = Null
'      Else
'         adotmp2210.Fields("t22105").Value = adoacc1k0.Fields("a1k02").Value
'      End If
'      If IsNull(adoacc1k0.Fields("a1k18").Value) Then
'         adotmp2210.Fields("t22106").Value = Null
'      Else
'         adotmp2210.Fields("t22106").Value = adoacc1k0.Fields("a1k18").Value
'      End If
'      If IsNull(adoacc1k0.Fields("a1k11").Value) Then
'         adotmp2210.Fields("t22107").Value = 0
'         adotmp2210.Fields("t22108").Value = 0
'      Else
'         If IsNull(adoacc1k0.Fields("a1k10").Value) Then
'            adotmp2210.Fields("t22107").Value = 0
'         Else
'            adotmp2210.Fields("t22107").Value = adoacc1k0.Fields("a1k11").Value / adoacc1k0.Fields("a1k10").Value
'         End If
'         adotmp2210.Fields("t22108").Value = adoacc1k0.Fields("a1k11").Value
'         If IsNull(adoacc1k0.Fields("a1k30").Value) Then
'            adotmp2210.Fields("t22111").Value = 0
'         Else
'            adotmp2210.Fields("t22111").Value = adoacc1k0.Fields("a1k30").Value - adoacc1k0.Fields("a1k11").Value
'         End If
'      End If
'      If IsNull(adoacc1k0.Fields("a1k29").Value) Then
'         adotmp2210.Fields("t22109").Value = Null
'      Else
'         adotmp2210.Fields("t22109").Value = adoacc1k0.Fields("a1k29").Value
'      End If
'      If IsNull(adoacc1k0.Fields("a1k09").Value) Then
'         adotmp2210.Fields("t22110").Value = 0
'      Else
'         adotmp2210.Fields("t22110").Value = adoacc1k0.Fields("a1k09").Value
'      End If
'      adotmp2210.UpdateBatch
'      adoacc1k0.MoveNext
'   Loop
'   adoacc1k0.Close
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
'End Sub

''*************************************************
''  儲存帳單資料
''
''*************************************************
'Private Sub Acc150Query()
'On Error GoTo Checking
'   strSql = ""
'   If Text1 <> MsgText(601) Then
'      strSql = strSql & " and a1503 >= '" & Text1 & "'"
'   End If
'   If Text2 <> MsgText(601) Then
'      strSql = strSql & " and a1503 <= '" & Text2 & "'"
'   End If
'   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'      strSql = strSql & " and a1502 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'   End If
'   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'      strSql = strSql & " and a1502 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'   End If
'   If Text3 <> MsgText(601) Then
'      Select Case Text3
'         Case "4", "6"
'            strSql = strSql & " and (a1510 - a1520) > 0"
'      End Select
'   End If
'   adoacc150.CursorLocation = adUseClient
'   adoacc150.Open "select * from acc150, fagent where a1503 = (fa01 || fa02)" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'   Do While adoacc150.EOF = False
'      adotmp2210.AddNew
'      If IsNull(adoacc150.Fields("a1503").Value) Then
'         adotmp2210.Fields("t22101").Value = Null
'      Else
'         adotmp2210.Fields("t22101").Value = adoacc150.Fields("a1503").Value
'      End If
'      If IsNull(adoacc150.Fields("fa04").Value) Then
'         adotmp2210.Fields("t22102").Value = Null
'      Else
'         adotmp2210.Fields("t22102").Value = adoacc150.Fields("fa04").Value
'      End If
'      If IsNull(adoacc150.Fields("fa10").Value) Then
'         adotmp2210.Fields("t22103").Value = Null
'      Else
'         adotmp2210.Fields("t22103").Value = NationQuery(adoacc150.Fields("fa10").Value, 1)
'      End If
'      adotmp2210.Fields("t22104").Value = adoacc150.Fields("a1501").Value
'      If IsNull(adoacc150.Fields("a1502").Value) Then
'         adotmp2210.Fields("t22105").Value = Null
'      Else
'         adotmp2210.Fields("t22105").Value = adoacc150.Fields("a1502").Value
'      End If
'      If IsNull(adoacc150.Fields("a1505").Value) Then
'         adotmp2210.Fields("t22106").Value = Null
'      Else
'         adotmp2210.Fields("t22106").Value = adoacc150.Fields("a1505").Value
'      End If
'      If IsNull(adoacc150.Fields("a1506").Value) Then
'         adotmp2210.Fields("t22107").Value = 0
'      Else
'         adotmp2210.Fields("t22107").Value = adoacc150.Fields("a1506").Value
'      End If
'      If IsNull(adoacc150.Fields("a1510").Value) Then
'         adotmp2210.Fields("t22108").Value = 0
'      Else
'         adotmp2210.Fields("t22108").Value = adoacc150.Fields("a1510").Value
'         If IsNull(adoacc150.Fields("a1520").Value) Then
'            adotmp2210.Fields("t22109").Value = Null
'         Else
'            If (adoacc150.Fields("a1510").Value - adoacc150.Fields("a1520").Value) > 0 Then
'               adotmp2210.Fields("t22109").Value = Null
'            Else
'               adotmp2210.Fields("t22109").Value = MsgText(602)
'               adotmp2210.Fields("t22111").Value = adoacc150.Fields("a1520").Value = adoacc150.Fields("a1510").Value
'            End If
'         End If
'      End If
'      If IsNull(adoacc150.Fields("a1504").Value) Then
'         adotmp2210.Fields("t22112").Value = Null
'      Else
'         adotmp2210.Fields("t22112").Value = adoacc150.Fields("a1504").Value
'      End If
'      adotmp2210.UpdateBatch
'      adoacc150.MoveNext
'   Loop
'   adoacc150.Close
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
'End Sub

'Add By Sindy 2012/8/14
Private Sub ClearSumCol()
   'FC
   '外幣
   Combo2.Clear
   Combo3.Clear
   '台幣
   Text13 = ""
   Text11 = ""
   Text17 = ""
   'CF
   '外幣
   Combo4.Clear
   Combo5.Clear
   '台幣
   Text12 = ""
End Sub

'*************************************************
'  計算並顯示合計
'
'*************************************************
Public Sub SumShow()
Dim dblRate As Double, i As Integer, intIndex As Integer
Dim dblAmount As Double
Dim CurrencyCnt As Integer                  '幣別數量
Dim CurrencyType(1 To 10) As String   '幣別
Dim dbl_Famount(1 To 10) As Double   '外幣合計
Dim dbl_Famount2(1 To 10) As Double '未收外幣合計
Dim dblSumA1606 As Double 'Add By Sindy 2012/8/14 抵帳單金額
Dim strConP As String, strCont As String, strConL As String, strConS As String 'Add By Sindy 2014/9/9
   
   'Add By Sindy 2014/9/9
   strConP = "": strCont = "": strConL = "": strConS = ""
   
   If Text4 <> "" Then
      If Text4 = "1" Then
         strConP = " and (pa161 is null or pa161<>'J')"
         strCont = " and (tm130 is null or tm130<>'J')"
         strConL = " and (lc48 is null or lc48<>'J')"
         strConS = " and (sp85 is null or sp85<>'J')"
      ElseIf Text4 = "2" Then
         strConP = " and pa161='J'"
         strCont = " and tm130='J'"
         strConL = " and lc48='J'"
         strConS = " and sp85='J'"
      End If
   End If
   '2014/9/9 END
   'Add by Amy 2017/02/17 +客戶編號條件
   If bolCus = True Then
        strConP = strConP & " And (pa26>= '" & Text5 & "' And pa26 <='" & Text6 & "' Or pa27 >= '" & Text5 & "' And pa27<='" & Text6 & "' Or pa28 >= '" & Text5 & "' And pa28 <='" & Text6 & "' Or pa29 >= '" & Text5 & "' And pa29 <='" & Text6 & "' Or pa30 >= '" & Text5 & "' And pa30 <='" & Text6 & "') "
        strCont = strCont & " And (tm23>= '" & Text5 & "' And tm23 <='" & Text6 & "' Or tm78 >= '" & Text5 & "' And tm78<='" & Text6 & "' Or tm79 >= '" & Text5 & "' And tm79 <='" & Text6 & "' Or tm80 >= '" & Text5 & "' And tm80 <='" & Text6 & "' Or tm81 >= '" & Text5 & "' And tm81 <='" & Text6 & "') "
        strConL = strConL & " And (lc11>= '" & Text5 & "' And lc11 <='" & Text6 & "' Or lc43 >= '" & Text5 & "' And lc43<='" & Text6 & "' Or lc44 >= '" & Text5 & "' And lc44 <='" & Text6 & "' Or lc45 >= '" & Text5 & "' And lc45 <='" & Text6 & "' Or lc46 >= '" & Text5 & "' And lc46 <='" & Text6 & "') "
        strConS = strConS & " And (sp08>= '" & Text5 & "' And sp08 <='" & Text6 & "' Or sp58 >= '" & Text5 & "' And sp58<='" & Text6 & "' Or sp59 >= '" & Text5 & "' And sp59 <='" & Text6 & "' Or sp65 >= '" & Text5 & "' And sp65 <='" & Text6 & "' Or sp66 >= '" & Text5 & "' And sp66 <='" & Text6 & "') "
   End If
   
   'Modify By Sindy 2013/1/15 Mark
'   'Add By Sindy 2010/3/12 逐筆計算外幣金額
'   If Adodc1.Recordset.RecordCount <> 0 Then
'      With Adodc1.Recordset
'         .MoveFirst
'         Do While Not .EOF
'            If Left(Trim(.Fields("DocNo")), 1) = "X" Then '請款單號
'               If Trim(.Fields("Currency")) <> "USD" Then
'                  dblRate = PUB_GetUSXRate_1(Replace(.Fields("DocDate"), "/", ""), Trim(.Fields("Currency")))
'                  dblAmount = Val(.Fields("Namount")) / dblRate
'                  .Fields("Famount").Value = dblAmount
'               End If
'            End If
'            .MoveNext
'         Loop
'      End With
'      DataGrid1.AllowUpdate = False '鎖住畫面,不可異動資料
'   End If
'   '2010/3/12 End
   '2013/1/15 End
   
   '計算各欄位合計
   '**************************************
   '外幣合計FC
   '**************************************
   adoaccsum.CursorLocation = adUseClient
   'strSql = "select a1k18,sum(a1k08 - nvl(a1k31, 0)) as Namount from acc1k0 where (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0) & " and a1k18='USD' group by a1k18"
   'Modify By Sindy 2014/9/10 +加公司別,要串基本檔
   'strSql = "select a1k18,sum(a1k08 - nvl(a1k31, 0)) as Namount from acc1k0 where (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0) & " group by a1k18"
   strSql = "select a1k18,sum(a1k08 - nvl(a1k31, 0)) as Namount from acc1k0,patent where (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0) & " and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04" & strConP & " group by a1k18" & _
           " union all" & _
           " select a1k18,sum(a1k08 - nvl(a1k31, 0)) as Namount from acc1k0,trademark where (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0) & " and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04" & strCont & " group by a1k18" & _
           " union all" & _
           " select a1k18,sum(a1k08 - nvl(a1k31, 0)) as Namount from acc1k0,lawcase where (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0) & " and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04" & strConL & " group by a1k18" & _
           " union all" & _
           " select a1k18,sum(a1k08 - nvl(a1k31, 0)) as Namount from acc1k0,servicepractice where (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0) & " and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04" & strConS & " group by a1k18"
   '2014/9/10 END
   'Modify By Sindy 2013/1/15 Mark
   'strSql = strSql & " union select a1k18,sum((a1k11 - nvl(a1k06, 0)) / nvl(DNR03,0)) as Namount from acc1k0, DEBITNOTERATE where (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0) & " and a1k18<>'USD' and DNR01=a1k18 AND DNR02=(SELECT max(DNR02) FROM DEBITNOTERATE WHERE DNR01=a1k18 AND DNR02<=a1k02) group by a1k18"
   '2013/1/15 End
   adoaccsum.Open "select a1k18,sum(Namount) from (" & strSql & ") group by a1k18 order by a1k18", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      adoaccsum.MoveFirst
      Do While Not adoaccsum.EOF
         If Val(" " & adoaccsum.Fields(1)) <> 0 Then
            Combo2.AddItem adoaccsum.Fields(0) & " " & Format(adoaccsum.Fields(1).Value, FDollar)
            Combo2.ListIndex = 0
         End If
         adoaccsum.MoveNext
      Loop
   End If
   adoaccsum.Close
   '**************************************
   '外幣FC未收
   '**************************************
   'strSql = "select a1k18,sum(a1k08 - nvl(a1k31, 0)) as Namount from acc1k0 where (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k18='USD' group by a1k18"
   'Modify By Sindy 2014/9/10 +加公司別,要串基本檔
   'strSql = "select a1k18,sum(a1k08 - nvl(a1k31, 0)) as Namount from acc1k0 where (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0) & " and (a1k29 is null or a1k29 = '') group by a1k18"
   strSql = "select a1k18,sum(a1k08 - nvl(a1k31, 0)) as Namount from acc1k0,patent where (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04" & strConP & " group by a1k18" & _
           " union all" & _
           " select a1k18,sum(a1k08 - nvl(a1k31, 0)) as Namount from acc1k0,trademark where (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04" & strCont & " group by a1k18" & _
           " union all" & _
           " select a1k18,sum(a1k08 - nvl(a1k31, 0)) as Namount from acc1k0,lawcase where (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04" & strConL & " group by a1k18" & _
           " union all" & _
           " select a1k18,sum(a1k08 - nvl(a1k31, 0)) as Namount from acc1k0,servicepractice where (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04" & strConS & " group by a1k18"
   '2014/9/10 END
   'Modify By Sindy 2013/1/15 Mark
   'strSql = strSql & " union select a1k18,sum((a1k11 - nvl(a1k06, 0)) / nvl(DNR03,0)) as Namount from acc1k0, DEBITNOTERATE where (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k18<>'USD' and DNR01=a1k18 AND DNR02=(SELECT max(DNR02) FROM DEBITNOTERATE WHERE DNR01=a1k18 AND DNR02<=a1k02) group by a1k18"
   '2013/1/15 End
   'Modify By Sindy 2014/9/10 +加公司別,要串基本檔
   'strSql = strSql & " union select a1k18,sum(a0z04 * (-1)) as Namount from acc0y0, acc0z0, acc1k0 where a0y01 = a0z01 and a0z02 = a1k01 and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k30>0 group by a1k18"
   'Modify by Amy 2017/02/17 +客戶編號條件
   strSql = strSql & " union all select a1k18,sum(a0z04 * (-1)) as Namount from acc0y0, acc0z0, acc1k0,patent where a0y01 = a0z01 and a0z02 = a1k01 and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k30>0 and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04" & strConP & " group by a1k18"
   strSql = strSql & " union all select a1k18,sum(a0z04 * (-1)) as Namount from acc0y0, acc0z0, acc1k0,trademark where a0y01 = a0z01 and a0z02 = a1k01 and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k30>0 and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04" & strCont & " group by a1k18"
   strSql = strSql & " union all select a1k18,sum(a0z04 * (-1)) as Namount from acc0y0, acc0z0, acc1k0,lawcase where a0y01 = a0z01 and a0z02 = a1k01 and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k30>0 and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04" & strConL & " group by a1k18"
   strSql = strSql & " union all select a1k18,sum(a0z04 * (-1)) as Namount from acc0y0, acc0z0, acc1k0,servicepractice where a0y01 = a0z01 and a0z02 = a1k01 and (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k30>0 and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04" & strConS & " group by a1k18"
   '2014/9/10 END
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select a1k18,sum(Namount) from (" & strSql & ") group by a1k18 order by a1k18", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      adoaccsum.MoveFirst
      Do While Not adoaccsum.EOF
         If Val(" " & adoaccsum.Fields(1)) <> 0 Then
            Combo3.AddItem adoaccsum.Fields(0) & " " & Format(adoaccsum.Fields(1).Value, FDollar)
            Combo3.ListIndex = 0
         End If
         adoaccsum.MoveNext
      Loop
   End If
   adoaccsum.Close
   '**************************************
   '台幣合計FC
   '**************************************
   adoaccsum.CursorLocation = adUseClient
   'Modify By Sindy 2014/9/10 +加公司別,要串基本檔
   'strSql = "select (a1k11 - nvl(a1k06, 0)) as Namount from acc1k0 where (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0)
   'Modify by Amy 2017/02/17 +客戶編號條件
   strSql = "select (a1k11 - nvl(a1k06, 0)) as Namount from acc1k0,patent where (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0) & " and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04" & strConP & _
            " union all" & _
            " select (a1k11 - nvl(a1k06, 0)) as Namount from acc1k0,trademark where (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0) & " and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04" & strCont & _
            " union all" & _
            " select (a1k11 - nvl(a1k06, 0)) as Namount from acc1k0,lawcase where (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0) & " and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04" & strConL & _
            " union all" & _
            " select (a1k11 - nvl(a1k06, 0)) as Namount from acc1k0,servicepractice where (a1k12 is null or a1k12 = 0) and a1k25 is null" & strWhere(0) & " and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04" & strConS
   '2014/9/10 END
   adoaccsum.Open "select sum(Namount) from (" & strSql & ") New", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text13 = MsgText(601)
      Else
         Text13 = Format(adoaccsum.Fields(0).Value, FDollar)
      End If
   Else
      Text13 = MsgText(601)
   End If
   adoaccsum.Close
   '**************************************
   '台幣FC未收,台幣FC未收規費
   '**************************************
   'Modify By Sindy 2014/9/10 +加公司別,要串基本檔
   'strSql = "select decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),2)) as Namount,decode(nvl(a1k30,0),0,a1k09,decode(sign(a1k30-a1k09),-1,a1k09-a1k30,0)) as Lawfee from acc1k0, ACC0Z0 where (a1k12 is null or a1k12 = 0) and a1k25 is null And A1K01=A0Z02(+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
   strSql = "select decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),2)) as Namount,decode(nvl(a1k30,0),0,a1k09,decode(sign(a1k30-a1k09),-1,a1k09-a1k30,0)) as Lawfee from acc1k0, ACC0Z0,patent where (a1k12 is null or a1k12 = 0) and a1k25 is null And A1K01=A0Z02(+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k13||''=pa01 and a1k14||''=pa02 and a1k15||''=pa03 and a1k16||''=pa04" & _
           strConP & _
           " union all" & _
           " select decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),2)) as Namount,decode(nvl(a1k30,0),0,a1k09,decode(sign(a1k30-a1k09),-1,a1k09-a1k30,0)) as Lawfee from acc1k0, ACC0Z0,trademark where (a1k12 is null or a1k12 = 0) and a1k25 is null And A1K01=A0Z02(+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k13||''=tm01 and a1k14||''=tm02 and a1k15||''=tm03 and a1k16||''=tm04" & _
           strCont & _
           " union all" & _
           " select decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),2)) as Namount,decode(nvl(a1k30,0),0,a1k09,decode(sign(a1k30-a1k09),-1,a1k09-a1k30,0)) as Lawfee from acc1k0, ACC0Z0,lawcase where (a1k12 is null or a1k12 = 0) and a1k25 is null And A1K01=A0Z02(+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k13||''=lc01 and a1k14||''=lc02 and a1k15||''=lc03 and a1k16||''=lc04" & _
           strConL & _
           " union all" & _
           " select decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),2)) as Namount,decode(nvl(a1k30,0),0,a1k09,decode(sign(a1k30-a1k09),-1,a1k09-a1k30,0)) as Lawfee from acc1k0, ACC0Z0,servicepractice where (a1k12 is null or a1k12 = 0) and a1k25 is null And A1K01=A0Z02(+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k13||''=sp01 and a1k14||''=sp02 and a1k15||''=sp03 and a1k16||''=sp04" & _
           strConS
   '2014/9/10 END
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(Namount),sum(Lawfee) from (" & strSql & ") New", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      '台幣FC未收
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text11 = MsgText(601)
      Else
         Text11 = Format(adoaccsum.Fields(0).Value, FDollar)
      End If
      '台幣FC未收規費
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text17 = MsgText(601)
      Else
         Text17 = Format(adoaccsum.Fields(1).Value, FDollar)
      End If
   Else
      Text11 = MsgText(601)
      Text17 = MsgText(601)
   End If
   adoaccsum.Close
   '**************************************
   '外幣CF合計
   '**************************************
   'Modify By Sindy 2014/9/10 +加公司別,要串基本檔
   'strSql = "select a1505,sum(axf04) as Namount from acc151, acc150 where axf01 = a1501 and (a1507 is null or a1507 = 0)" & strWhere(2) & " group by a1505"
   'strSql = strSql & " union select a1605 as a1505,sum(axg04 * (-1)) as Namount from acc161, acc160 where axg01(+)=a1601" & strWhere(4) & " and a1607 is not null group by a1605"
   strSql = "select a1505,sum(axf04) as Namount from acc151, acc150,patent where axf01 = a1501 and (a1507 is null or a1507 = 0)" & strWhere(2) & " and substr(AXF03,1,length(AXF03)-9)=pa01 and substr(substr(AXF03,-9),1,6)=pa02 and substr(substr(AXF03,-3),1,1)=pa03 and substr(AXF03,-2)=pa04" & strConP & " group by a1505" & _
           " union all" & _
           " select a1505,sum(axf04) as Namount from acc151, acc150,trademark where axf01 = a1501 and (a1507 is null or a1507 = 0)" & strWhere(2) & " and substr(AXF03,1,length(AXF03)-9)=tm01 and substr(substr(AXF03,-9),1,6)=tm02 and substr(substr(AXF03,-3),1,1)=tm03 and substr(AXF03,-2)=tm04" & strCont & " group by a1505" & _
           " union all" & _
           " select a1505,sum(axf04) as Namount from acc151, acc150,lawcase where axf01 = a1501 and (a1507 is null or a1507 = 0)" & strWhere(2) & " and substr(AXF03,1,length(AXF03)-9)=lc01 and substr(substr(AXF03,-9),1,6)=lc02 and substr(substr(AXF03,-3),1,1)=lc03 and substr(AXF03,-2)=lc04" & strConL & " group by a1505" & _
           " union all" & _
           " select a1505,sum(axf04) as Namount from acc151, acc150,servicepractice where axf01 = a1501 and (a1507 is null or a1507 = 0)" & strWhere(2) & " and substr(AXF03,1,length(AXF03)-9)=sp01 and substr(substr(AXF03,-9),1,6)=sp02 and substr(substr(AXF03,-3),1,1)=sp03 and substr(AXF03,-2)=sp04" & strConS & " group by a1505"
   strSql = strSql & " union all select a1605 as a1505,sum(axg04 * (-1)) as Namount from acc161, acc160,patent where axg01(+)=a1601" & strWhere(4) & " and a1607 is not null and substr(AXG03,1,length(AXG03)-9)=pa01 and substr(substr(AXG03,-9),1,6)=pa02 and substr(substr(AXG03,-3),1,1)=pa03 and substr(AXG03,-2)=pa04" & strConP & " group by a1605"
   strSql = strSql & " union all select a1605 as a1505,sum(axg04 * (-1)) as Namount from acc161, acc160,trademark where axg01(+)=a1601" & strWhere(4) & " and a1607 is not null and substr(AXG03,1,length(AXG03)-9)=tm01 and substr(substr(AXG03,-9),1,6)=tm02 and substr(substr(AXG03,-3),1,1)=tm03 and substr(AXG03,-2)=tm04" & strCont & " group by a1605"
   strSql = strSql & " union all select a1605 as a1505,sum(axg04 * (-1)) as Namount from acc161, acc160,lawcase where axg01(+)=a1601" & strWhere(4) & " and a1607 is not null and substr(AXG03,1,length(AXG03)-9)=lc01 and substr(substr(AXG03,-9),1,6)=lc02 and substr(substr(AXG03,-3),1,1)=lc03 and substr(AXG03,-2)=lc04" & strConL & " group by a1605"
   strSql = strSql & " union all select a1605 as a1505,sum(axg04 * (-1)) as Namount from acc161, acc160,servicepractice where axg01(+)=a1601" & strWhere(4) & " and a1607 is not null and substr(AXG03,1,length(AXG03)-9)=sp01 and substr(substr(AXG03,-9),1,6)=sp02 and substr(substr(AXG03,-3),1,1)=sp03 and substr(AXG03,-2)=sp04" & strConS & " group by a1605"
   '2014/9/10 END
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select a1505,sum(Namount) from (" & strSql & ") group by a1505 order by a1505", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      adoaccsum.MoveFirst
      Do While Not adoaccsum.EOF
         If Val(" " & adoaccsum.Fields(1)) <> 0 Then
            Combo4.AddItem adoaccsum.Fields(0) & " " & adoaccsum.Fields(1)
            Combo4.ListIndex = 0
         End If
         adoaccsum.MoveNext
      Loop
   End If
   adoaccsum.Close
   '**************************************
   '外幣CF未付
   '**************************************
   strSql = ""
   adoaccsum.CursorLocation = adUseClient
   'Modify By Sindy 2012/8/31 Y21399查詢時剛好acc190已有付款資料但未有匯款編號, 因此增加a1908 is null的判斷
   'Modify By Sindy 2014/9/10 +加公司別,要串基本檔
   'strSql = "select a1505,sum(axf04) as Namount from acc151, acc150, acc190 where a1902(+) = axf01 and axf01 = a1501 and (a1507 is null or a1507 = 0)" & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1512 is null and a1908 is null group by a1505"
   'strSql = strSql & " union select a1605 as a1505,sum(axg04 * (-1)) as Namount from acc161, acc160 where axg01(+)=a1601" & strWhere(4) & " and a1607 is null group by a1605"
   strSql = "select a1505,sum(axf04) as Namount from acc151, acc150, acc190,patent where a1902(+) = axf01 and axf01 = a1501 and (a1507 is null or a1507 = 0)" & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1512 is null and a1908 is null and substr(AXF03,1,length(AXF03)-9)=pa01 and substr(substr(AXF03,-9),1,6)=pa02 and substr(substr(AXF03,-3),1,1)=pa03 and substr(AXF03,-2)=pa04" & strConP & " group by a1505" & _
           " union all" & _
           " select a1505,sum(axf04) as Namount from acc151, acc150, acc190,trademark where a1902(+) = axf01 and axf01 = a1501 and (a1507 is null or a1507 = 0)" & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1512 is null and a1908 is null and substr(AXF03,1,length(AXF03)-9)=tm01 and substr(substr(AXF03,-9),1,6)=tm02 and substr(substr(AXF03,-3),1,1)=tm03 and substr(AXF03,-2)=tm04" & strCont & " group by a1505" & _
           " union all" & _
           " select a1505,sum(axf04) as Namount from acc151, acc150, acc190,lawcase where a1902(+) = axf01 and axf01 = a1501 and (a1507 is null or a1507 = 0)" & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1512 is null and a1908 is null and substr(AXF03,1,length(AXF03)-9)=lc01 and substr(substr(AXF03,-9),1,6)=lc02 and substr(substr(AXF03,-3),1,1)=lc03 and substr(AXF03,-2)=lc04" & strConL & " group by a1505" & _
           " union all" & _
           " select a1505,sum(axf04) as Namount from acc151, acc150, acc190,servicepractice where a1902(+) = axf01 and axf01 = a1501 and (a1507 is null or a1507 = 0)" & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1512 is null and a1908 is null and substr(AXF03,1,length(AXF03)-9)=sp01 and substr(substr(AXF03,-9),1,6)=sp02 and substr(substr(AXF03,-3),1,1)=sp03 and substr(AXF03,-2)=sp04" & strConS & " group by a1505"
   strSql = strSql & " union all select a1605 as a1505,sum(axg04 * (-1)) as Namount from acc161, acc160,patent where axg01(+)=a1601" & strWhere(4) & " and a1607 is null and substr(AXG03,1,length(AXG03)-9)=pa01 and substr(substr(AXG03,-9),1,6)=pa02 and substr(substr(AXG03,-3),1,1)=pa03 and substr(AXG03,-2)=pa04" & strConP & " group by a1605"
   strSql = strSql & " union all select a1605 as a1505,sum(axg04 * (-1)) as Namount from acc161, acc160,trademark where axg01(+)=a1601" & strWhere(4) & " and a1607 is null and substr(AXG03,1,length(AXG03)-9)=tm01 and substr(substr(AXG03,-9),1,6)=tm02 and substr(substr(AXG03,-3),1,1)=tm03 and substr(AXG03,-2)=tm04" & strCont & " group by a1605"
   strSql = strSql & " union all select a1605 as a1505,sum(axg04 * (-1)) as Namount from acc161, acc160,lawcase where axg01(+)=a1601" & strWhere(4) & " and a1607 is null and substr(AXG03,1,length(AXG03)-9)=lc01 and substr(substr(AXG03,-9),1,6)=lc02 and substr(substr(AXG03,-3),1,1)=lc03 and substr(AXG03,-2)=lc04" & strConL & " group by a1605"
   strSql = strSql & " union all select a1605 as a1505,sum(axg04 * (-1)) as Namount from acc161, acc160,servicepractice where axg01(+)=a1601" & strWhere(4) & " and a1607 is null and substr(AXG03,1,length(AXG03)-9)=sp01 and substr(substr(AXG03,-9),1,6)=sp02 and substr(substr(AXG03,-3),1,1)=sp03 and substr(AXG03,-2)=sp04" & strConS & " group by a1605"
   '2014/9/10 END
   adoaccsum.Open "select a1505,sum(Namount) from (" & strSql & ") group by a1505 order by a1505", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      adoaccsum.MoveFirst
      Do While Not adoaccsum.EOF
         If Val(" " & adoaccsum.Fields(1)) <> 0 Then
            Combo5.AddItem adoaccsum.Fields(0) & " " & adoaccsum.Fields(1)
            Combo5.ListIndex = 0
         End If
         adoaccsum.MoveNext
      Loop
   End If
   adoaccsum.Close
   '**************************************
   '台幣CF已付
   '**************************************
   'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
   'Modify By Sindy 2014/9/10 +加公司別,要串基本檔
   strSql = "select sum(axf04*a1906) as Namount from acc151, acc150, acc190,patent where axf01 = a1501 and a1501 = a1902 and a1908 is not null and (a1507 is null or a1507 = 0)" & strWhere(2) & " and substr(AXF03,1,length(AXF03)-9)=pa01 and substr(substr(AXF03,-9),1,6)=pa02 and substr(substr(AXF03,-3),1,1)=pa03 and substr(AXF03,-2)=pa04" & strConP & _
           " union all" & _
           " select sum(axf04*a1906) as Namount from acc151, acc150, acc190,trademark where axf01 = a1501 and a1501 = a1902 and a1908 is not null and (a1507 is null or a1507 = 0)" & strWhere(2) & " and substr(AXF03,1,length(AXF03)-9)=tm01 and substr(substr(AXF03,-9),1,6)=tm02 and substr(substr(AXF03,-3),1,1)=tm03 and substr(AXF03,-2)=tm04" & strCont & _
           " union all" & _
           " select sum(axf04*a1906) as Namount from acc151, acc150, acc190,lawcase where axf01 = a1501 and a1501 = a1902 and a1908 is not null and (a1507 is null or a1507 = 0)" & strWhere(2) & " and substr(AXF03,1,length(AXF03)-9)=lc01 and substr(substr(AXF03,-9),1,6)=lc02 and substr(substr(AXF03,-3),1,1)=lc03 and substr(AXF03,-2)=lc04" & strConL & _
           " union all" & _
           " select sum(axf04*a1906) as Namount from acc151, acc150, acc190,servicepractice where axf01 = a1501 and a1501 = a1902 and a1908 is not null and (a1507 is null or a1507 = 0)" & strWhere(2) & " and substr(AXF03,1,length(AXF03)-9)=sp01 and substr(substr(AXF03,-9),1,6)=sp02 and substr(substr(AXF03,-3),1,1)=sp03 and substr(AXF03,-2)=sp04" & strConS
   '2014/9/10 END
   'add by sonia 2021/4/9 加抵帳已付
   strSql = strSql & " union all" & _
           " select sum(axf04*nvl(a1g03,0)) as Namount from acc151, acc150, acc1g0,patent where axf01 = a1501 and a1512 = a1g01 and a1g01 is not null and (a1507 is null or a1507 = 0)" & strWhere(2) & " and substr(AXF03,1,length(AXF03)-9)=pa01 and substr(substr(AXF03,-9),1,6)=pa02 and substr(substr(AXF03,-3),1,1)=pa03 and substr(AXF03,-2)=pa04" & strConP & _
           " union all" & _
           " select sum(axf04*nvl(a1g03,0)) as Namount from acc151, acc150, acc1g0,trademark where axf01 = a1501 and a1512 = a1g01 and a1g01 is not null and (a1507 is null or a1507 = 0)" & strWhere(2) & " and substr(AXF03,1,length(AXF03)-9)=tm01 and substr(substr(AXF03,-9),1,6)=tm02 and substr(substr(AXF03,-3),1,1)=tm03 and substr(AXF03,-2)=tm04" & strCont & _
           " union all" & _
           " select sum(axf04*nvl(a1g03,0)) as Namount from acc151, acc150, acc1g0,lawcase where axf01 = a1501 and a1512 = a1g01 and a1g01 is not null and (a1507 is null or a1507 = 0)" & strWhere(2) & " and substr(AXF03,1,length(AXF03)-9)=lc01 and substr(substr(AXF03,-9),1,6)=lc02 and substr(substr(AXF03,-3),1,1)=lc03 and substr(AXF03,-2)=lc04" & strConL & _
           " union all" & _
           " select sum(axf04*nvl(a1g03,0)) as Namount from acc151, acc150, acc1g0,servicepractice where axf01 = a1501 and a1512 = a1g01 and a1g01 is not null and (a1507 is null or a1507 = 0)" & strWhere(2) & " and substr(AXF03,1,length(AXF03)-9)=sp01 and substr(substr(AXF03,-9),1,6)=sp02 and substr(substr(AXF03,-3),1,1)=sp03 and substr(AXF03,-2)=sp04" & strConS
   'end 2021/4/9
   'add by sonia 2021/5/27 扣除抵帳單ACC160
   strSql = strSql & " union all" & _
           " select sum(axg04*a1906) * (-1) as Namount from acc161, acc160, acc190 where axg01 = a1601 and a1601 = a1902 and a1908 is not null and a1607 is not null " & strWhere(4) & _
           " union all" & _
           " select sum(axg04*nvl(a1g03,0)) * (-1) as Namount from acc161, acc160,acc190,acc1i0 c,acc1i0 d,acc1g0 where axg01 = a1601 and a1607=c.a1i03(+) and a1605=c.a1i05(+) and a1607=d.a1i03 and nvl(c.a1i01,d.a1i01)=a1g01(+) and a1g01 is not null and a1607 is not null and a1601=a1902(+) and a1901 is null " & strWhere(4)
   'end 2021/5/27
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(Namount) from (" & strSql & ") New", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text12 = MsgText(601)
      Else
         Text12 = Format(adoaccsum.Fields(0).Value, FDollar)
      End If
   Else
      Text12 = MsgText(601)
   End If
   adoaccsum.Close
   
   '依查詢資料顯示各欄位值
   Select Case Text3
      Case "1", "2"
         'FC
         '外幣
'         Combo2.Clear: Text10 = ""
'         Combo3.Clear: Text15 = ""
         '台幣
'         Text13 = ""
'         Text11 = ""
'         Text17 = ""
         'CF
         '外幣
         Combo4.Clear
         Combo5.Clear
         '台幣
         Text12 = ""
         If Text3 = "2" Then '2.FC未收
            Combo2.Clear
            Text13 = ""
         End If
      Case "3", "4"
         'FC
         '外幣
         Combo2.Clear
         Combo3.Clear
         '台幣
         Text13 = ""
         Text11 = ""
         Text17 = ""
         'CF
         '外幣
'         Combo4.Clear: Text14 = ""
'         Combo5.Clear: Text16 = ""
         '台幣
'         Text12 = ""
         If Text3 = "4" Then '4.CF未付
            Combo4.Clear
            Text12 = ""
         End If
      Case "", "5", "6"
         'FC
         '外幣
'         Combo2.Clear: Text10 = ""
'         Combo3.Clear: Text15 = ""
         '台幣
'         Text13 = ""
'         Text11 = ""
'         Text17 = ""
         'CF
         '外幣
'         Combo4.Clear: Text14 = ""
'         Combo5.Clear: Text16 = ""
         '台幣
'         Text12 = ""
         If Text3 = "6" Then '未收未付
            Combo2.Clear
            Text13 = ""
            Combo4.Clear
            Text12 = ""
         End If
   End Select
'Modify By Sindy 2012/8/23 Mark 統一寫法
'*****************************************************************
'   'Add By Sindy 2010/3/12 計算外幣金額
'   '預設值
'   CurrencyCnt = 0
'   For i = 1 To 10
'      CurrencyType(i) = ""
'      dbl_Famount(i) = 0
'      dbl_Famount2(i) = 0
'   Next i
'   If Adodc1.Recordset.RecordCount <> 0 And _
'      (Text3 = "1" Or Text3 = "2" Or Text3 = "5" Or Text3 = "6" Or Trim(Text3) = "") Then
'      With Adodc1.Recordset
'         .MoveFirst
'         Do While Not .EOF
'            intIndex = 0
'            If Left(Trim(.Fields("DocNo")), 1) = "X" Then '請款單號
'               '幣別
'               If CurrencyCnt = 0 Then
'                  CurrencyCnt = 1
'                  intIndex = 1
'                  CurrencyType(intIndex) = Trim(.Fields("Currency"))
'               Else
'                  For i = 1 To CurrencyCnt
'                     If CurrencyType(i) = Trim(.Fields("Currency")) Then
'                        intIndex = i
'                        Exit For
'                     End If
'                  Next i
'                  If intIndex = 0 Then
'                     CurrencyCnt = CurrencyCnt + 1
'                     intIndex = CurrencyCnt
'                     CurrencyType(intIndex) = Trim(.Fields("Currency"))
'                  End If
'               End If
'               If CurrencyType(intIndex) <> "USD" Then
'                  dblRate = PUB_GetUSXRate_1(Replace(.Fields("DocDate"), "/", ""), CurrencyType(intIndex))
'               End If
'               '***** 外幣合計 *****
'               If Not IsNull(.Fields("a1k12")) And .Fields("a1k12") <> 0 Then GoTo ReadNext
'               'Modify By Sindy 2012/8/14 Mark
''               If Text3 = "1" Or Text3 = "5" Or Trim(Text3) = "" Then '直接加總
'                  If CurrencyType(intIndex) = "USD" Then
'                     dbl_Famount(intIndex) = dbl_Famount(intIndex) + Val(.Fields("Famount"))
'                  Else
'                     dblAmount = Format(((Val(.Fields("Namount")) * 100 * 100) \ (dblRate * 100)) / 100, FAmount)
'                     .Fields("Famount").Value = dblAmount
'                     dbl_Famount(intIndex) = dbl_Famount(intIndex) + dblAmount
'                  End If
''               End If
''               If Text3 = "2" Or Text3 = "6" Then '2.FC未收 6.未收未付   需計算
''                  If CurrencyType(intIndex) = "USD" Then
''                     If Val(.Fields("a1k30")) = 0 Then
''                        dbl_Famount(intIndex) = dbl_Famount(intIndex) + Val(.Fields("Famount"))
''                     Else
''                        dblAmount = Format((Val(.Fields("Namount")) + Val(.Fields("a1k30"))) / Val(.Fields("a1k10")), FAmount)
''                        dbl_Famount(intIndex) = dbl_Famount(intIndex) + dblAmount
''                     End If
''                  Else
''                     If Val(.Fields("a1k30")) = 0 Then
''                        dblAmount = Format(((Val(.Fields("Namount")) * 100 * 100) \ (dblRate * 100)) / 100, FAmount)
''                        dbl_Famount(intIndex) = dbl_Famount(intIndex) + dblAmount
''                     Else
''                        dblAmount = Format((((Val(.Fields("Namount")) + Val(.Fields("a1k30"))) * 100 * 100) \ (dblRate * 100)) / 100, FAmount)
''                        dbl_Famount(intIndex) = dbl_Famount(intIndex) + dblAmount
''                     End If
''                  End If
''               End If
'               '***** 未收外幣合計 *****
'               If Not IsNull(.Fields("Close")) And .Fields("Close") <> "" Then GoTo ReadNext
'               If Text3 = "1" Or Text3 = "2" Then
'                  If Not IsNull(.Fields("a1k25")) And .Fields("a1k25") <> "" Then GoTo ReadNext
'               End If
'               'Modify By Sindy 2012/8/14 Mark
''               If Text3 = "1" Or Text3 = "5" Or Trim(Text3) = "" Then '1.FC往來 5.往來   需計算
'                  If CurrencyType(intIndex) = "USD" Then
'                     If Val(.Fields("a1k30")) = 0 Then
'                        dbl_Famount2(intIndex) = dbl_Famount2(intIndex) + Val(.Fields("Famount"))
'                     Else
'                        dblAmount = Format((Val(.Fields("Namount")) - Val(.Fields("a1k30"))) / Val(.Fields("a1k10")), FAmount)
'                        dbl_Famount2(intIndex) = dbl_Famount2(intIndex) + dblAmount
'                     End If
'                  Else
'                     If Val(.Fields("a1k30")) = 0 Then
'                        dblAmount = Format(((Val(.Fields("Namount")) * 100 * 100) \ (dblRate * 100)) / 100, FAmount)
'                        dbl_Famount2(intIndex) = dbl_Famount2(intIndex) + dblAmount
'                     Else
'                        dblAmount = Format((((Val(.Fields("Namount")) - Val(.Fields("a1k30"))) * 100 * 100) \ (dblRate * 100)) / 100, FAmount)
'                        dbl_Famount2(intIndex) = dbl_Famount2(intIndex) + dblAmount
'                     End If
'                  End If
''               End If
''               If Text3 = "2" Or Text3 = "6" Then '直接加總
''                  If CurrencyType(intIndex) = "USD" Then
''                     dbl_Famount2(intIndex) = dbl_Famount2(intIndex) + Val(.Fields("Famount"))
''                  Else
''                     dblAmount = Format(((Val(.Fields("Namount")) * 100 * 100) \ (dblRate * 100)) / 100, FAmount)
''                     .Fields("Famount").Value = dblAmount
''                     dbl_Famount2(intIndex) = dbl_Famount2(intIndex) + dblAmount
''                  End If
''               End If
'            End If
'ReadNext:
'            .MoveNext
'         Loop
'      End With
'      DataGrid1.AllowUpdate = False '鎖住畫面,不可異動資料
'      '組下拉式選單
'      If CurrencyCnt <> 0 Then
'         For i = 1 To CurrencyCnt
'            '外幣合計
'            If dbl_Famount(i) > 0 Then
'               Combo2.AddItem CurrencyType(i) & " " & dbl_Famount(i)
'               Combo2.ListIndex = 0
'            End If
'            '未收
'            If dbl_Famount2(i) > 0 Then
'               Combo3.AddItem CurrencyType(i) & " " & dbl_Famount2(i)
'               Combo3.ListIndex = 0
'            End If
'         Next i
'      End If
'   End If
'   '2010/3/12 End
'
'   Select Case Text3
'      Case "1", "2"
'         adoaccsum.CursorLocation = adUseClient
'         'FC台幣
'         adoaccsum.Open "select sum(a1k11 - nvl(a1k06, 0) * a1k10) from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0), adoTaie, adOpenStatic, adLockReadOnly
'         If adoaccsum.RecordCount <> 0 Then
'            If IsNull(adoaccsum.Fields(0).Value) Then
'               Text13 = MsgText(601)
'            Else
'               Text13 = Format(adoaccsum.Fields(0).Value, FDollar)
'            End If
'         Else
'            Text13 = MsgText(601)
'         End If
'         adoaccsum.Close
'         adoaccsum.CursorLocation = adUseClient
'         '93.12.31 MODIFY BY SONIA 加計FC未收外幣
'         'adoaccsum.Open "select sum(a1k11 - nvl(a1k06, 0) * a1k10) from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')", adoTaie, adOpenStatic, adLockReadOnly
'
'         'Modify by Morgan 2005/3/7 扣除銷帳作廢及已收款金額
'         'adoaccsum.Open "select sum(a1k11 - nvl(a1k06, 0) * a1k10),sum(a1k08 - nvl(a1k06, 0)) from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')", adoTaie, adOpenStatic, adLockReadOnly
'         '93.12.31 END
'         '2009/4/27 modify by sonia 改同grid
'         'adoaccsum.Open "select sum(a1k11-nvl(a1k30,0) - nvl(a1k06, 0) * a1k10),sum(a1k08 - nvl(a1k06, 0))- Sum(Nvl(A0Z04,0)) from acc1k0, fagent, nation, ACC0Z0 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0) And A1K01=A0Z02(+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k25 is null and a1k12 is null", adoTaie, adOpenStatic, adLockReadOnly
'         'Modify By Sindy 2010/8/31 增加,sum(nvl(a1k09,0))
'         'adoaccsum.Open "select sum(decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0))),sum(decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) ) from acc1k0, fagent, nation, ACC0Z0 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0) And A1K01=A0Z02(+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k25 is null and a1k12 is null", adoTaie, adOpenStatic, adLockReadOnly
'         adoaccsum.Open "select sum(decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0))),sum(decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))),sum(nvl(a1k09,0)) from acc1k0, fagent, nation, ACC0Z0 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0) And A1K01=A0Z02(+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k25 is null and a1k12 is null", adoTaie, adOpenStatic, adLockReadOnly
'         '2005/3/7 end
'         If adoaccsum.RecordCount <> 0 Then
'            'FC未收台幣
'            If IsNull(adoaccsum.Fields(0).Value) Then
'               Text11 = MsgText(601)
'            Else
'               Text11 = Format(adoaccsum.Fields(0).Value, FDollar)
'            End If
'            '93.12.31 MODIFY BY SONIA 加計FC未收外幣
'            If IsNull(adoaccsum.Fields(1).Value) Then
'               Text8 = MsgText(601)
'            Else
'               Text8 = Format(adoaccsum.Fields(1).Value, FDollar)
'            End If
'            '93.12.31 END
'            'Add By Sindy 2010/8/31 未收規費
'            If IsNull(adoaccsum.Fields(2).Value) Then
'               Text17 = MsgText(601)
'            Else
'               Text17 = Format(adoaccsum.Fields(2).Value, FDollar)
'            End If
'         Else
'            Text11 = MsgText(601)
'            Text8 = MsgText(601)
'            Text17 = MsgText(601) 'Add By Sindy 2010/8/31 未收規費
'         End If
'         adoaccsum.Close
'        'Add By Cheng 2004/04/16
'        '加顯示FC外幣
'         adoaccsum.CursorLocation = adUseClient
'         '93.12.31 MODIFY BY SONIA 扣除折讓
'         'adoaccsum.Open "select sum(Nvl(a1k08,0)) from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0), adoTaie, adOpenStatic, adLockReadOnly
'         adoaccsum.Open "select sum(Nvl(a1k08,0) - nvl(a1k06, 0)) from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0), adoTaie, adOpenStatic, adLockReadOnly
'         '93.12.31 END
'         If adoaccsum.RecordCount <> 0 Then
'            If IsNull(adoaccsum.Fields(0).Value) Then
'               Text7 = ""
'            Else
'               Text7 = Format(adoaccsum.Fields(0).Value, FDollar)
'            End If
'         Else
'            Text7 = ""
'         End If
'         adoaccsum.Close
'        'End
'         If Text3 = "2" Then
'            Combo2.Clear
'            Text13 = ""
'         End If
'      Case "3", "4"
'         adoaccsum.CursorLocation = adUseClient
'         'CF已付台幣
'         '不用扣抵帳單...2012/8/13 Sindy
'         adoaccsum.Open "select sum(a1520 * a1906) from acc150, fagent, nation, acc190 where substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1501 = a1902 and (a1507 is null or a1507 = 0)" & strWhere(2) & " and (a1520 is not null Or a1520 <> 0)", adoTaie, adOpenStatic, adLockReadOnly
'         If adoaccsum.RecordCount <> 0 Then
'            If IsNull(adoaccsum.Fields(0).Value) Then
'               Text12 = MsgText(601)
'            Else
'               Text12 = Format(adoaccsum.Fields(0).Value, FDollar)
'            End If
'         Else
'            Text12 = MsgText(601)
'         End If
'         adoaccsum.Close
'         'Add By Sindy 2010/7/13 增加CF 及 未付合計
'         'CF外幣
'         adoaccsum.CursorLocation = adUseClient
'         adoaccsum.Open "select a1505,sum(axf04) as Namount from acc151, acc150, fagent, nation where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1507 is null or a1507 = 0)" & strWhere(2) & " group by a1505 order by a1505", adoTaie, adOpenStatic, adLockReadOnly
'         If adoaccsum.RecordCount <> 0 Then
'            adoaccsum.MoveFirst
'            Do While Not adoaccsum.EOF
'               If Val(" " & adoaccsum.Fields(1)) <> 0 Then
'                  'Modify By Sindy 2012/8/14
'                  dblSumA1606 = Val(GetACC160Amt(adoaccsum.Fields(0), ""))
'                  Combo4.AddItem adoaccsum.Fields(0) & " " & (Val(adoaccsum.Fields(1)) - dblSumA1606)
'                  '2012/8/14 End
'                  Combo4.ListIndex = 0
'               End If
'               adoaccsum.MoveNext
'            Loop
'         End If
'         adoaccsum.Close
'         'CF未付外幣
'         strSql = ""
'         adoaccsum.CursorLocation = adUseClient
'         strSql = "select a1505,sum(axf04) as Namount from acc151, acc150, fagent, nation where axf01 = a1501 and substr(a1503, 1, 8) = fa01(+) and substr(a1503, 9, 1) = fa02(+) and fa10 = na01(+) " & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1505"
'         'strSql = strSql & " union select a1505,sum(a1904 * (-1)) as Namount from acc190, acc151, acc150, fagent, nation where a1902 = axf01 and axf01 = a1501 and substr(a1503, 1, 8) = fa01(+) and substr(a1503, 9, 1) = fa02(+) and fa10 = na01(+) " & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1505"
'         adoaccsum.Open "select a1505,sum(Namount) from (" & strSql & ") group by a1505 order by a1505", adoTaie, adOpenStatic, adLockReadOnly
'         If adoaccsum.RecordCount <> 0 Then
'            adoaccsum.MoveFirst
'            Do While Not adoaccsum.EOF
'               If Val(" " & adoaccsum.Fields(1)) <> 0 Then
'                  'Modify By Sindy 2012/8/14
'                  dblSumA1606 = Val(GetACC160Amt(adoaccsum.Fields(0), "0"))
'                  Combo5.AddItem adoaccsum.Fields(0) & " " & (Val(adoaccsum.Fields(1)) - dblSumA1606)
'                  '2012/8/14 End
'                  Combo5.ListIndex = 0
'               End If
'               adoaccsum.MoveNext
'            Loop
'         'Add By Sindy 2012/8/15 檢查是否有未抵帳的抵帳單
'         Else
'            Call GetACC160Amt("", "0")
'         '2012/8/15 End
'         End If
'         adoaccsum.Close
'         '2010/7/13 End
'         If Text3 = "4" Then
'            Combo4.Clear
'            Text12 = ""
'         End If
'         Text13 = ""
'         Text11 = ""
'         Text7 = ""
'         Text8 = ""
'         Text17 = "" 'Add By Sindy 2010/8/31 未收規費
'      Case "", "5", "6"
'         If Text3 = "" Or Text3 = "5" Then
'            adoaccsum.CursorLocation = adUseClient
'            'FC台幣
'            adoaccsum.Open "select sum(a1k11 - nvl(a1k06, 0) * a1k10) from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0), adoTaie, adOpenStatic, adLockReadOnly
'            If adoaccsum.RecordCount <> 0 Then
'               If IsNull(adoaccsum.Fields(0).Value) Then
'                  Text13 = MsgText(601)
'               Else
'                  Text13 = Format(adoaccsum.Fields(0).Value, FDollar)
'               End If
'            Else
'               Text13 = MsgText(601)
'            End If
'            adoaccsum.Close
'            adoaccsum.CursorLocation = adUseClient
'            'CF已付台幣
'            adoaccsum.Open "select sum(a1520 * a1906) from acc150, fagent, nation, acc190 where substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1501 = a1902 and (a1507 is null or a1507 = 0)" & strWhere(2) & " and (a1520 is not null Or a1520 <> 0)", adoTaie, adOpenStatic, adLockReadOnly
'            If adoaccsum.RecordCount <> 0 Then
'               If IsNull(adoaccsum.Fields(0).Value) Then
'                  Text12 = MsgText(601)
'               Else
'                  Text12 = Format(adoaccsum.Fields(0).Value, FDollar)
'               End If
'            Else
'               Text12 = MsgText(601)
'            End If
'            adoaccsum.Close
'            'Add By Sindy 2010/7/13 增加CF 及 未付合計
'            'CF外幣
'            adoaccsum.CursorLocation = adUseClient
'            adoaccsum.Open "select a1505,sum(axf04) as Namount from acc151, acc150, fagent, nation where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1507 is null or a1507 = 0)" & strWhere(2) & " group by a1505 order by a1505", adoTaie, adOpenStatic, adLockReadOnly
'            If adoaccsum.RecordCount <> 0 Then
'               adoaccsum.MoveFirst
'               Do While Not adoaccsum.EOF
'                  If Val(" " & adoaccsum.Fields(1)) <> 0 Then
'                     'Modify By Sindy 2012/8/14
'                     dblSumA1606 = Val(GetACC160Amt(adoaccsum.Fields(0), ""))
'                     Combo4.AddItem adoaccsum.Fields(0) & " " & (Val(adoaccsum.Fields(1)) - dblSumA1606)
'                     '2012/8/14 End
'                     Combo4.ListIndex = 0
'                  End If
'                  adoaccsum.MoveNext
'               Loop
'            End If
'            adoaccsum.Close
'            '2010/7/13 End
'            'Add By Cheng 2004/04/16
'            '加顯示FC外幣
'            adoaccsum.CursorLocation = adUseClient
'            '93.12.31 MODIFY BY SONIA 扣除折讓
'            'adoaccsum.Open "select sum(Nvl(a1k08,0)) from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0), adoTaie, adOpenStatic, adLockReadOnly
'            adoaccsum.Open "select sum(Nvl(a1k08,0) - nvl(a1k06, 0)) from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0), adoTaie, adOpenStatic, adLockReadOnly
'            '93.12.31 END
'            If adoaccsum.RecordCount <> 0 Then
'               If IsNull(adoaccsum.Fields(0).Value) Then
'                  Text7 = ""
'               Else
'                  Text7 = Format(adoaccsum.Fields(0).Value, FDollar)
'               End If
'            Else
'               Text7 = ""
'            End If
'            adoaccsum.Close
'            'End
'         Else
'            Combo2.Clear
'            Text13 = ""
'            Text12 = ""
'            Text7 = ""
'         End If
'         adoaccsum.CursorLocation = adUseClient
'         '93.12.31 MODIFY BY SONIA 加計FC未收外幣
'         'adoaccsum.Open "select sum(a1k11 - nvl(a1k06, 0) * a1k10) from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')", adoTaie, adOpenStatic, adLockReadOnly
'         '2009/4/27 modify by sonia 改同grid
'         'adoaccsum.Open "select sum(a1k11 - nvl(a1k06, 0) * a1k10),sum(a1k08 - nvl(a1k06, 0)) from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')", adoTaie, adOpenStatic, adLockReadOnly
'         'Modify By Sindy 2010/8/31 增加,sum(nvl(a1k09,0))
'         'adoaccsum.Open "select sum(decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0))),sum(decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) ) from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')", adoTaie, adOpenStatic, adLockReadOnly
'         adoaccsum.Open "select sum(decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0))),sum(decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))),sum(nvl(a1k09,0)) from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k25 is null and a1k12 is null", adoTaie, adOpenStatic, adLockReadOnly
'         '93.12.31 END
'         If adoaccsum.RecordCount <> 0 Then
'            'FC未收台幣
'            If IsNull(adoaccsum.Fields(0).Value) Then
'               Text11 = MsgText(601)
'            Else
'               Text11 = Format(adoaccsum.Fields(0).Value, FDollar)
'            End If
'            '93.12.31 MODIFY BY SONIA 加計FC未收外幣
'            If IsNull(adoaccsum.Fields(1).Value) Then
'               Text8 = MsgText(601)
'            Else
'               Text8 = Format(adoaccsum.Fields(1).Value, FDollar)
'            End If
'            '93.12.31 END
'            'Add By Sindy 2010/8/31 未收規費
'            If IsNull(adoaccsum.Fields(2).Value) Then
'               Text17 = MsgText(601)
'            Else
'               Text17 = Format(adoaccsum.Fields(2).Value, FDollar)
'            End If
'         Else
'            Text11 = MsgText(601)
'            Text8 = MsgText(601)
'            Text17 = MsgText(601) 'Add By Sindy 2010/8/31 未收規費
'         End If
'         adoaccsum.Close
'         'Add By Sindy 2010/7/13 增加CF 及 未付合計
'         'CF未付外幣
'         strSql = ""
'         adoaccsum.CursorLocation = adUseClient
'         strSql = "select a1505,sum(axf04) as Namount from acc151, acc150, fagent, nation where axf01 = a1501 and substr(a1503, 1, 8) = fa01(+) and substr(a1503, 9, 1) = fa02(+) and fa10 = na01(+) " & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1505"
'         'strSql = strSql & " union select a1505,sum(a1904 * (-1)) as Namount from acc190, acc151, acc150, fagent, nation where a1902 = axf01 and axf01 = a1501 and substr(a1503, 1, 8) = fa01(+) and substr(a1503, 9, 1) = fa02(+) and fa10 = na01(+) " & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1505"
'         adoaccsum.Open "select a1505,sum(Namount) from (" & strSql & ") group by a1505 order by a1505", adoTaie, adOpenStatic, adLockReadOnly
'         If adoaccsum.RecordCount <> 0 Then
'            adoaccsum.MoveFirst
'            Do While Not adoaccsum.EOF
'               If Val(" " & adoaccsum.Fields(1)) <> 0 Then
'                  'Modify By Sindy 2012/8/14
'                  dblSumA1606 = Val(GetACC160Amt(adoaccsum.Fields(0), "0"))
'                  Combo5.AddItem adoaccsum.Fields(0) & " " & (Val(adoaccsum.Fields(1)) - dblSumA1606)
'                  '2012/8/14 End
'                  Combo5.ListIndex = 0
'               End If
'               adoaccsum.MoveNext
'            Loop
'         'Add By Sindy 2012/8/15 檢查是否有未抵帳的抵帳單
'         Else
'            Call GetACC160Amt("", "0")
'         '2012/8/15 End
'         End If
'         adoaccsum.Close
'         '2010/7/13 End
'         '2011/9/6 add by sonia
'         If Text3 = "6" Then
'            Combo2.Clear
'            Text13 = ""
'            Combo4.Clear
'            Text12 = ""
'         End If
'         '2011/9/6 end
'*****************************************************************
'2009/4/28 CANCEL BY SONIA
'      Case Else
'         adoaccsum.CursorLocation = adUseClient
'         adoaccsum.Open "select sum(a1k11 - nvl(a1k06, 0) * a1k10) from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0), adoTaie, adOpenStatic, adLockReadOnly
'         If adoaccsum.RecordCount <> 0 Then
'            If IsNull(adoaccsum.Fields(0).Value) Then
'               Text13 = MsgText(601)
'            Else
'               Text13 = Format(adoaccsum.Fields(0).Value, FDollar)
'            End If
'         Else
'            Text13 = MsgText(601)
'         End If
'         adoaccsum.Close
'         adoaccsum.CursorLocation = adUseClient
'         '93.12.31 MODIFY BY SONIA 加計FC未收外幣
'         'adoaccsum.Open "select sum(a1k11 - nvl(a1k06, 0) * a1k10) from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')", adoTaie, adOpenStatic, adLockReadOnly
'         adoaccsum.Open "select sum(a1k11 - nvl(a1k06, 0) * a1k10),sum(a1k08 - nvl(a1k06, 0)) from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')", adoTaie, adOpenStatic, adLockReadOnly
'         '93.12.31 END
'         If adoaccsum.RecordCount <> 0 Then
'            If IsNull(adoaccsum.Fields(0).Value) Then
'               Text11 = MsgText(601)
'            Else
'               Text11 = Format(adoaccsum.Fields(0).Value, FDollar)
'            End If
'            '93.12.31 MODIFY BY SONIA 加計FC未收外幣
'            If IsNull(adoaccsum.Fields(1).Value) Then
'               Text8 = MsgText(601)
'            Else
'               Text8 = Format(adoaccsum.Fields(1).Value, FDollar)
'            End If
'            '93.12.31 END
'         Else
'            Text11 = MsgText(601)
'            Text8 = MsgText(601)
'         End If
'         adoaccsum.Close
'         adoaccsum.CursorLocation = adUseClient
'         adoaccsum.Open "select sum(a1520 * a1906) from acc150, fagent, nation, acc190 where substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1501 = a1902 and (a1507 is null or a1507 = 0)" & strWhere(2) & " and (a1520 is not null Or a1520 <> 0)", adoTaie, adOpenStatic, adLockReadOnly
'         If adoaccsum.RecordCount <> 0 Then
'            If IsNull(adoaccsum.Fields(0).Value) Then
'               Text12 = MsgText(601)
'            Else
'               Text12 = Format(adoaccsum.Fields(0).Value, FDollar)
'            End If
'         Else
'            Text12 = MsgText(601)
'         End If
'         adoaccsum.Close
'         'Add By Cheng 2004/04/16
'         '加顯示FC外幣
'         adoaccsum.CursorLocation = adUseClient
'         '93.12.31 MODIFY BY SONIA 扣除折讓
'         'adoaccsum.Open "select sum(Nvl(a1k08,0)) from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0), adoTaie, adOpenStatic, adLockReadOnly
'         adoaccsum.Open "select sum(Nvl(a1k08,0) - nvl(a1k06, 0)) from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0), adoTaie, adOpenStatic, adLockReadOnly
'         '93.12.31 END
'         If adoaccsum.RecordCount <> 0 Then
'            If IsNull(adoaccsum.Fields(0).Value) Then
'               Text7 = ""
'            Else
'               Text7 = Format(adoaccsum.Fields(0).Value, FDollar)
'            End If
'         Else
'            Text7 = ""
'         End If
'         adoaccsum.Close
'         'End
'2009/4/28 END
'   End Select
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   'Modify by Amy 2013/10/30 +代理人編號起迄必填(不填跑很慢)
'   If Text1 <> MsgText(601) Then
'      FormCheck = True
'      Exit Function
'   End If
'   If Text2 <> MsgText(601) Then
'      FormCheck = True
'      Exit Function
'   End If
   If Text1 <> MsgText(601) And Text2 <> MsgText(601) Then
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

Private Sub Text3_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2003/05/06
    KeyAscii = UpperCase(KeyAscii)
    Select Case KeyAscii
    Case 8, 49, 50, 51, 52, 53, 54
    Case Else
        KeyAscii = 0
    End Select
End Sub
'2009/4/29 ADD BY SONIA
Private Sub Text3_Validate(Cancel As Boolean)
   If Text3 = "" Then Text3 = "5"
End Sub
'2009/4/29 END

''Add By Sindy 2012/8/14 計算抵帳單金額
'Private Function GetACC160Amt(strA1605 As String, strType As String) As Double
'   Dim rsTmp As New ADODB.Recordset
'   Dim strConSql As String
'
'   GetACC160Amt = 0
'   strConSql = ""
'
'   If strA1605 <> "" Then '幣別
'      strConSql = strConSql & " and A1605='" & strA1605 & "'"
'   End If
'
'   If strType = "0" Then '未付
'      strSql = "select a1605,sum(axg04) from acc161, acc160, ACC190 where A1607 IS NULL AND axg01(+)=a1601 AND A1601=A1902(+) AND A1901 IS NULL" & strWhere(4) & strConSql & " group by a1605"
'   Else
'      strSql = "select a1605,sum(axg04) from acc161, acc160 where axg01(+)=a1601" & strWhere(4) & strConSql & " and a1607 is not null group by a1605"
'   End If
'   intI = 1
'   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      GetACC160Amt = rsTmp.Fields(1)
'      '若無傳入幣別,代表外層無資料,因此在此處直接增加欄位資料
'      If strA1605 = "" And strType = "0" Then
'         Combo5.AddItem rsTmp.Fields(0) & " " & (0 - GetACC160Amt)
'         Combo5.ListIndex = 0
'      End If
'   End If
'   rsTmp.Close
'End Function

'Added by Lydia 2017/01/16 呼叫共用表單->帳款處理情形歷史記錄
Private Sub cmdDizhang_Click()
Dim rsRead As New ADODB.Recordset
Dim sqlB As String
Dim intB As Integer

    If Text1.Text = "" And Text2.Text = "" Then
       MsgBox "請輸入代理人編號 !", vbExclamation
       Text1.SetFocus
       Exit Sub
    End If
    
    sqlB = "select '' V,DR01,DECODE(CU01,NULL,NVL(FA04,NVL(FA05,FA06)),NVL(CU04,NVL(CU05,CU06))) DR01N, SQLDATET(DR04) DR04,DR02,ST02,SQLTIME(DR05) DR05 " & _
          "FROM DizhangRecord,STAFF,customer,fagent " & _
          "WHERE DR03=ST01(+) AND SUBSTR(DR01,1,8)=CU01(+) AND SUBSTR(DR01,9,1)=CU02(+) AND SUBSTR(DR01,1,8)=FA01(+) AND SUBSTR(DR01,9,1)=FA02(+) " & _
          "AND DR01>='" & ChangeCustomerL(Text1.Text) & "' " & IIf(Text2.Text <> "", "AND DR01<='" & ChangeCustomerL(Text2.Text) & "' ", "") & _
          "order by DR01,DR04,DR05"
    intB = 0
    Set rsRead = ClsLawReadRstMsg(intB, sqlB)
    If intB = 1 Then
       Set frm880012.grdDataList.Recordset = rsRead
       Set frm880012.fmParent = Me
       frm880012.iTyp = "3"
       frm880012.Show vbModal
    End If
End Sub

'Add by Amy 2016/02/17
'更新客戶編號
Private Sub UpdCusData()
    Dim RsQ As New ADODB.Recordset, rsA As New ADODB.Recordset
    Dim strQ As String, strSql As String, strUpd As String
    Dim intQ As Integer, intA As Integer
    Dim strNo(4) As String
    
    '無法一次語法更新,多申請人前六碼相同,抓申請人第一個是此客戶編號  ex:Y20804 X21419 多申請人(前六碼相同)
    strQ = "Select CaseNo From accrpt2210 Where Id='" & strUserNum & "' Group by CaseNo Order by CaseNo"
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        With RsQ
            .MoveFirst
            Do While Not .EOF
                'Modify by Amy 2018/11/28 舊資料無案號造成Error
                strNo(0) = "" & .Fields("CaseNo")
                If Trim(strNo(0)) <> MsgText(601) Then
                    strNo(1) = SystemNumber(strNo(0), 1)
                    strNo(2) = SystemNumber(strNo(0), 2)
                    strNo(3) = SystemNumber(strNo(0), 3)
                    strNo(4) = SystemNumber(strNo(0), 4)
                    
                    Select Case strNo(1)
                        Case "CFP", "FCP", "P" '專利
                            strSql = "Select CusNo From (" & _
                                        "Select '(1)'||pa26 as CusNo From Patent Where pa01='" & strNo(1) & "' And pa02='" & strNo(2) & "' " & _
                                    "And pa03='" & strNo(3) & "' and pa04='" & strNo(4) & "' And pa26>= '" & Text5 & "' And pa26 <='" & Text6 & "' And pa26 is not null " & _
                             "Union Select '(2)'||pa27 as CusNo From Patent Where pa01='" & strNo(1) & "' And pa02='" & strNo(2) & "' " & _
                                    "And pa03='" & strNo(3) & "' and pa04='" & strNo(4) & "' And pa27>= '" & Text5 & "' And pa27 <='" & Text6 & "' And pa27 is not null " & _
                             "Union Select '(3)'||pa28 as CusNo From Patent Where pa01='" & strNo(1) & "' And pa02='" & strNo(2) & "' " & _
                                    "And pa03='" & strNo(3) & "' and pa04='" & strNo(4) & "' And pa28>= '" & Text5 & "' And pa28 <='" & Text6 & "' And pa28 is not null " & _
                             "Union Select '(4)'||pa29 as CusNo From Patent Where pa01='" & strNo(1) & "' And pa02='" & strNo(2) & "' " & _
                                    "And pa03='" & strNo(3) & "' and pa04='" & strNo(4) & "' And pa29>= '" & Text5 & "' And pa29 <='" & Text6 & "' And pa29 is not null " & _
                             "Union Select '(5)'||pa30 as CusNo From Patent Where pa01='" & strNo(1) & "' And pa02='" & strNo(2) & "' " & _
                                    "And pa03='" & strNo(3) & "' and pa04='" & strNo(4) & "' And pa30>= '" & Text5 & "' And pa30 <='" & Text6 & "' And pa30 is not null " & _
                                    ") Order by CusNo "
                        Case "CFT", "FCT", "T", "TF" '商標
                            strSql = "Select CusNo From (" & _
                                        "Select '(1)'||tm23 as CusNo From Trademark Where tm01='" & strNo(1) & "' And tm02='" & strNo(2) & "' " & _
                                    "And tm03='" & strNo(3) & "' and tm04='" & strNo(4) & "' And tm23>= '" & Text5 & "' And tm23 <='" & Text6 & "' And tm23 is not null " & _
                             "Union Select '(2)'||tm78 as CusNo From Trademark Where tm01='" & strNo(1) & "' And tm02='" & strNo(2) & "' " & _
                                    "And tm03='" & strNo(3) & "' and tm04='" & strNo(4) & "' And tm78>= '" & Text5 & "' And tm78 <='" & Text6 & "' And tm78 is not null " & _
                             "Union Select '(3)'||tm79 as CusNo From Trademark Where tm01='" & strNo(1) & "' And tm02='" & strNo(2) & "' " & _
                                    "And tm03='" & strNo(3) & "' and tm04='" & strNo(4) & "' And tm79>= '" & Text5 & "' And tm79 <='" & Text6 & "' And tm79 is not null " & _
                             "Union Select '(4)'||tm80 as CusNo From Trademark Where tm01='" & strNo(1) & "' And tm02='" & strNo(2) & "' " & _
                                    "And tm03='" & strNo(3) & "' and tm04='" & strNo(4) & "' And tm80>= '" & Text5 & "' And tm80 <='" & Text6 & "' And tm80 is not null " & _
                             "Union Select '(5)'||tm81 as CusNo From Trademark Where tm01='" & strNo(1) & "' And tm02='" & strNo(2) & "' " & _
                                    "And tm03='" & strNo(3) & "' and tm04='" & strNo(4) & "' And tm81>= '" & Text5 & "' And tm81 <='" & Text6 & "' And tm81 is not null " & _
                                    ") Order by CusNo "
                        Case "CFL", "FCL", "L", "LIN" '法務
                            strSql = "Select CusNo From (" & _
                                        "Select '(1)'||lc11 as CusNo From Lawcase Where lc01='" & strNo(1) & "' And lc02='" & strNo(2) & "' " & _
                                    "And lc03='" & strNo(3) & "' And lc04='" & strNo(4) & "' And lc11>= '" & Text5 & "' And lc11 <='" & Text6 & "' And lc11 is not null " & _
                             "Union Select '(2)'||lc43 as CusNo From Lawcase Where lc01='" & strNo(1) & "' And lc02='" & strNo(2) & "' " & _
                                    "And lc03='" & strNo(3) & "' And lc04='" & strNo(4) & "' And lc43>= '" & Text5 & "' And lc43 <='" & Text6 & "' And lc43 is not null " & _
                             "Union Select '(3)'||lc44 as CusNo From Lawcase Where lc01='" & strNo(1) & "' And lc02='" & strNo(2) & "' " & _
                                    "And lc03='" & strNo(3) & "' And lc04='" & strNo(4) & "' And lc44>= '" & Text5 & "' And lc44 <='" & Text6 & "' And lc44 is not null " & _
                             "Union Select '(4)'||lc45 as CusNo From Lawcase Where lc01='" & strNo(1) & "' And lc02='" & strNo(2) & "' " & _
                                    "And lc03='" & strNo(3) & "' And lc04='" & strNo(4) & "' And lc45>= '" & Text5 & "' And lc45 <='" & Text6 & "' And lc45 is not null " & _
                             "Union Select '(5)'||lc46 as CusNo From Lawcase Where lc01='" & strNo(1) & "' And lc02='" & strNo(2) & "' " & _
                                    "And lc03='" & strNo(3) & "' And lc04='" & strNo(4) & "' And lc46>= '" & Text5 & "' And lc46 <='" & Text6 & "' And lc46 is not null " & _
                                    ") Order by CusNo "
                        Case Else '服務
                            strSql = "Select CusNo From (" & _
                                        "Select '(1)'||sp08 as CusNo From Servicepractice Where sp01='" & strNo(1) & "' And sp02='" & strNo(2) & "' " & _
                                    "And sp03='" & strNo(3) & "' And sp04='" & strNo(4) & "' And sp08>= '" & Text5 & "' And sp08 <='" & Text6 & "' And sp08 is not null " & _
                             "Union Select '(2)'||sp58 as CusNo From Servicepractice Where sp01='" & strNo(1) & "' And sp02='" & strNo(2) & "' " & _
                                    "And sp03='" & strNo(3) & "' And sp04='" & strNo(4) & "' And sp58>= '" & Text5 & "' And sp58 <='" & Text6 & "' And sp58 is not null " & _
                             "Union Select '(3)'||sp59 as CusNo From Servicepractice Where sp01='" & strNo(1) & "' And sp02='" & strNo(2) & "' " & _
                                    "And sp03='" & strNo(3) & "' And sp04='" & strNo(4) & "' And sp59>= '" & Text5 & "' And sp59 <='" & Text6 & "' And sp59 is not null " & _
                             "Union Select '(4)'||sp65 as CusNo From Servicepractice Where sp01='" & strNo(1) & "' And sp02='" & strNo(2) & "' " & _
                                    "And sp03='" & strNo(3) & "' And sp04='" & strNo(4) & "' And sp65>= '" & Text5 & "' And sp65 <='" & Text6 & "' And sp65 is not null " & _
                             "Union Select '(5)'||sp66 as CusNo From Servicepractice Where sp01='" & strNo(1) & "' And sp02='" & strNo(2) & "' " & _
                                    "And sp03='" & strNo(3) & "' And sp04='" & strNo(4) & "' And sp66>= '" & Text5 & "' And sp66 <='" & Text6 & "' And sp66 is not null " & _
                                    ") Order by CusNo "
                    End Select
                     intA = 1
                    Set rsA = ClsLawReadRstMsg(intA, strSql)
                    If intA = 1 Then
                        strUpd = "Update accrpt2210 Set CusNo='" & rsA.Fields("CusNo") & "' Where Id='" & strUserNum & "' " & _
                                        "And Substr(CaseNo,1,length(CaseNo)-12)='" & strNo(1) & "' And Substr(CaseNo,(length(CaseNo)-12)+2,6)='" & strNo(2) & "' " & _
                                        "And Substr(CaseNo,(length(CaseNo)-5)+2,1)='" & strNo(3) & "' And Substr(CaseNo,(length(CaseNo)-3)+2,2)='" & strNo(4) & "' "
                        cnnConnection.Execute strUpd
                    End If
                    rsA.Close
                End If
                .MoveNext
            Loop
        End With
    End If
    RsQ.Close
End Sub

'原程式搬過來備份
Private Sub UpdData()
  Dim Str01 As String, Str02 As String, Str03 As String, Str04 As String 'Add By Sindy 2014/3/24
  Dim rsQuery As ADODB.Recordset 'Add By Sindy 2014/9/9
  Dim strPA161 As String 'Add By Sindy 2014/9/9
  Dim strCaseNo As String
  
  'Modify By Sindy 2014/9/9 逐筆讀取特殊出名公司
   strSql = "select CaseNo from accrpt2210 where id='" & strUserNum & "' group by CaseNo"
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      'With rsQuery
         rsQuery.MoveFirst
         Do While Not rsQuery.EOF
            strCaseNo = rsQuery.Fields("CaseNo")
            Str01 = SystemNumber(strCaseNo, 1)
            Str02 = SystemNumber(strCaseNo, 2)
            Str03 = SystemNumber(strCaseNo, 3)
            Str04 = SystemNumber(strCaseNo, 4)
            strSql = "select decode(pa161,'T','專利商標','J','智權公司',pa161)" & _
                    " From patent" & _
                    " where pa01='" & Str01 & "' and pa02='" & Str02 & "' and pa03='" & Str03 & "' and pa04='" & Str04 & "'" & _
                    " Union" & _
                    " select decode(tm130,'J','智權公司',tm130)" & _
                    " From trademark" & _
                    " where tm01='" & Str01 & "' and tm02='" & Str02 & "' and tm03='" & Str03 & "' and tm04='" & Str04 & "'" & _
                    " Union" & _
                    " select decode(lc48,'J','智權公司',lc48)" & _
                    " From lawcase" & _
                    " where lc01='" & Str01 & "' and lc02='" & Str02 & "' and lc03='" & Str03 & "' and lc04='" & Str04 & "'" & _
                    " Union" & _
                    " select decode(sp85,'J','智權公司',sp85)" & _
                    " From servicepractice" & _
                    " where sp01='" & Str01 & "' and sp02='" & Str02 & "' and sp03='" & Str03 & "' and sp04='" & Str04 & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               strPA161 = "" & RsTemp.Fields(0)
               cnnConnection.Execute "update accrpt2210 set PA161='" & strPA161 & "' where id='" & strUserNum & "' and CaseNo='" & strCaseNo & "'"
            End If
            rsQuery.MoveNext
         Loop
      'End With
   End If
   rsQuery.Close
   Set rsQuery = Nothing
   '2014/9/9 END
End Sub

Private Sub Text5_GotFocus()
    TextInverse Text5
    CloseIme
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_LostFocus()
    If Text5 <> "" Then Text6 = Left(Text5, 6) & "ZZZ"
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
    If Len(Text5) = 6 Then
        Text5 = AfterZero(Text5)
    End If
    If Text5.Text <> "" Then
        Text5 = Left(Text5 & "000000000", 9)
    End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_LostFocus()
    TextInverse Text6
    CloseIme
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
    If Len(Text6) = 6 Then
        Text6 = AfterZero(Text6)
    End If
    If Text6 <> "" Then
        Text6 = Left(Text6 & "000000000", 9)
    End If
End Sub
