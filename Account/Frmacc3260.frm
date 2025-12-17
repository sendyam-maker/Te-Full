VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc3260 
   AutoRedraw      =   -1  'True
   Caption         =   "收票資料查詢"
   ClientHeight    =   5550
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   8760
   Begin VB.ComboBox CboCmp 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Top             =   90
      Width           =   3500
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   810
      TabIndex        =   32
      Top             =   4950
      Width           =   420
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6840
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   450
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4920
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   450
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   10
      Top             =   1650
      Width           =   615
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   6960
      TabIndex        =   23
      Top             =   4995
      Width           =   1275
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   7
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   4
      Top             =   850
      Width           =   615
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   1572
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc3260.frx":0000
      Height          =   2805
      Left            =   240
      TabIndex        =   17
      Top             =   2100
      Width           =   8295
      _ExtentX        =   14623
      _ExtentY        =   4957
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
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
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "收票資料"
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "a0g02"
         Caption         =   "收票銀行"
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
         DataField       =   "a0e13"
         Caption         =   "收票日期"
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
      BeginProperty Column02 
         DataField       =   "a0e02"
         Caption         =   "票據號碼"
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
         DataField       =   "a0e10"
         Caption         =   "到期日期"
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
      BeginProperty Column04 
         DataField       =   "a0e11"
         Caption         =   "票據金額"
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
         DataField       =   "a0e20"
         Caption         =   "存入帳號"
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
         DataField       =   "a0e14"
         Caption         =   "託收日期"
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
      BeginProperty Column07 
         DataField       =   "a0e15"
         Caption         =   "退票日期"
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
      BeginProperty Column08 
         DataField       =   "a0e16"
         Caption         =   "抽票日期"
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
      BeginProperty Column09 
         DataField       =   "a0e17"
         Caption         =   "貼現日期"
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
      BeginProperty Column10 
         DataField       =   "a0e21"
         Caption         =   "兌現日期"
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
      BeginProperty Column11 
         DataField       =   "a0e03"
         Caption         =   "單據編號"
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
      BeginProperty Column12 
         DataField       =   "contect"
         Caption         =   "往來對象"
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
         DataField       =   "a0e12"
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
      BeginProperty Column14 
         DataField       =   "a0e34"
         Caption         =   "轉出日期"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   2709.921
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1120.252
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1480.252
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1120.252
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   1120.252
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   1090.205
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   1090.205
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            ColumnWidth     =   1090.205
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1330.016
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   4089.827
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   5139.78
         EndProperty
         BeginProperty Column14 
            Alignment       =   2
            ColumnWidth     =   1370.268
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   240
      Top             =   2050
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   547
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   4920
      TabIndex        =   5
      Top             =   850
      Width           =   1572
      _ExtentX        =   2769
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   6840
      TabIndex        =   6
      Top             =   850
      Width           =   1572
      _ExtentX        =   2769
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
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   300
      Left            =   4920
      TabIndex        =   8
      Top             =   1200
      Width           =   1572
      _ExtentX        =   2769
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
   Begin MSMask.MaskEdBox MaskEdBox4 
      Height          =   300
      Left            =   6840
      TabIndex        =   9
      Top             =   1200
      Width           =   1572
      _ExtentX        =   2769
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
   Begin MSMask.MaskEdBox MaskEdBox5 
      Height          =   300
      Left            =   4920
      TabIndex        =   11
      Top             =   1650
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
   Begin MSMask.MaskEdBox MaskEdBox6 
      Height          =   300
      Left            =   6840
      TabIndex        =   12
      Top             =   1650
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
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   33
      Top             =   150
      Width           =   972
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "共          張"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   450
      TabIndex        =   31
      Top             =   4980
      Width           =   1380
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "(Y/N)"
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
      Left            =   3000
      TabIndex        =   30
      Top             =   1700
      Width           =   615
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "是否含已兌現票據"
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
      TabIndex        =   29
      Top             =   1700
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      Height          =   375
      Left            =   240
      Top             =   1620
      Width           =   8295
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   28
      Top             =   1635
      Width           =   255
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "兌現日期"
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
      Left            =   3960
      TabIndex        =   27
      Top             =   1700
      Width           =   975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "到期日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3960
      TabIndex        =   26
      Top             =   1200
      Width           =   972
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6600
      TabIndex        =   25
      Top             =   1200
      Width           =   252
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "合計"
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
      Left            =   6360
      TabIndex        =   24
      Top             =   4995
      Width           =   495
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "只查詢託收未兌現"
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
      TabIndex        =   22
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "(Y/N)"
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
      Left            =   3000
      TabIndex        =   21
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "(Y/N)"
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
      Left            =   2520
      TabIndex        =   20
      Top             =   850
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "只查詢未託收"
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
      TabIndex        =   19
      Top             =   850
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "票據號碼"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   18
      Top             =   480
      Width           =   972
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6600
      TabIndex        =   16
      Top             =   850
      Width           =   252
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "收票日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3960
      TabIndex        =   15
      Top             =   850
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6600
      TabIndex        =   14
      Top             =   480
      Width           =   252
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "入帳帳號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3960
      TabIndex        =   13
      Top             =   480
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1515
      Left            =   240
      Top             =   30
      Width           =   8295
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4512
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc3260"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/07 Form2.0已修改 DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoadodc1 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Dim strSql As String
Dim strUnion As String
'92.11.11 ADD BY SONIA
Dim strSQL1 As String
Dim strSQL2 As String           '20140122ADD By eric
Dim strCmp As String, strCmpN As String 'Add by Sindy 2020/04/17


'Add by Sindy 2020/04/17
Private Sub SetCompN()
    strCmpN = "": strCmp = ""
    If Trim(CboCmp) <> MsgText(601) Then
        strCmp = CboCmp
        If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
        End If
    End If
    strCmpN = GetAccReportCmpN(strCmp, True, True)
End Sub

'Add by Sindy 2020/04/17
Private Sub CboCmp_GotFocus()
    TextInverse CboCmp
End Sub

Private Sub CboCmp_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboCmp_Validate(Cancel As Boolean)
    Dim strCmp As String
    
    If Trim(CboCmp) = MsgText(601) Then Exit Sub
    
    strCmp = CboCmp
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If
    If InStr(GetBookKeepCmp & 組合作帳公司 & ",", strCmp) = 0 Then
        MsgBox Label18 & MsgText(63), , MsgText(5)
        Cancel = True
        CboCmp.SetFocus
        Exit Sub
    ElseIf Len(Trim(CboCmp)) = 1 Then
        CboCmp = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/04/17

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 8850
   Me.Height = 6000 'Modify by Amy 2023/08/18 原:5745
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   'Add by Sindy 2020/04/17 公司別改下拉
   CboCmp.AddItem "", 0
   Call Pub_SetCboCmp(CboCmp, True, False, False, , 1)
   'end 2020/04/17
   
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   MaskEdBox3.Mask = DFormat
   MaskEdBox4.Mask = DFormat
   '92.11.10 add by sonia
   MaskEdBox5.Mask = DFormat
   MaskEdBox6.Mask = DFormat
   '92.11.10 end
   'Add by Morgan 2006/7/25
   PUB_SetAccount Combo1
   PUB_SetAccount Combo2
   'end 2006/7/25
   
   'Add by Morgan 2009/9/21 --瑞婷
   Text4 = "Y"
   Text5 = "Y"
   Text7 = "Y"
   'End 2009/9/21
   
   'Add by Morgan 2011/7/18 預設最大到期日--瑞婷
   strExc(0) = "select max(a0e10) from acc0e0 where a0e04='R'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      MaskEdBox4 = Format(RsTemp.Fields(0), MaskEdBox4.Mask)
   End If
   'end 2011/7/18
   
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc3260 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
   adoadodc1.CursorLocation = adUseClient
   '20140122START Modify By eric
   'Modify By Sindy 2020/4/17
   'adoadodc1.Open "select * from acc0e0 where a0e23 = '" & IIf(Text2 = "2", "J", "1") & "' and a0e01 >= '" & Combo1 & "' and a0e01 <= '" & Combo2 & "' and a0e13 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0e13 <= " & Val(FCDate(MaskEdBox2.Text)) & " and a0e25 = 0 and a0e04 = '" & MsgText(18) & "' order by a0e01 asc, a0e02 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc1.Open "select * from acc0e0 where a0e23 = '" & Left(CboCmp, 1) & "' and a0e01 >= '" & Combo1 & "' and a0e01 <= '" & Combo2 & "' and a0e13 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0e13 <= " & Val(FCDate(MaskEdBox2.Text)) & " and a0e25 = 0 and a0e04 = '" & MsgText(18) & "' order by a0e01 asc, a0e02 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2020/4/17 END
   'adoadodc1.Open "select * from acc0e0 where a0e01 >= '" & Combo1 & "' and a0e01 <= '" & Combo2 & "' and a0e13 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0e13 <= " & Val(FCDate(MaskEdBox2.Text)) & " and a0e25 = 0 and a0e04 = '" & MsgText(18) & "' order by a0e01 asc, a0e02 asc", adoTaie, adOpenStatic, adLockReadOnly
   '20140122END
   Set Adodc1.Recordset = adoadodc1
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()

On Error GoTo Checking
   strSql = ""
   strSQL1 = ""
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   
   Call SetCompN 'Add by Sindy 2020/04/17
   
   adoadodc1.CursorLocation = adUseClient
   If Text3 = MsgText(601) Then
      '20140122START Modify By eric
      'Modify By Sindy 2020/4/17 公司別改變數
'      If Text2 <> "" Then
'         strSql = " and a0e23 = '" & IIf(Text2 = "2", "J", "1") & "' "
'      Else
'         strSql = ""
'      End If
      If strCmp <> MsgText(601) Then
          If InStr(strCmp, "+") > 0 Then
             strSql = " And (a0e23 is null or a0e23 In ('" & Replace(strCmp, "+", "','") & "')) "
          Else
             strSql = " And (a0e23 is null or a0e23='" & strCmp & "') "
          End If
      Else
         strSql = ""
      End If
      '2020/4/17 END
      
      If Combo1 <> MsgText(601) Then
         strSql = strSql & " and a0e20 >= '" & Combo1 & "'"
      End If
      'If Combo1 <> MsgText(601) Then
      '   strSql = " and a0e20 >= '" & Combo1 & "'"
      'End If
      '20140122END
      If Combo2 <> MsgText(601) Then
         strSql = strSql & " and a0e20 <= '" & Combo2 & "'"
      End If
      If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
         strSql = strSql & " and a0e13 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      End If
      If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
         strSql = strSql & " and a0e13 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      End If
      If MaskEdBox3.Text <> MsgText(601) And MaskEdBox3.Text <> MsgText(29) Then
         strSql = strSql & " and a0e10 >= " & Val(FCDate(MaskEdBox3.Text)) & ""
      End If
      If MaskEdBox4.Text <> MsgText(601) And MaskEdBox4.Text <> MsgText(29) Then
         strSql = strSql & " and a0e10 <= " & Val(FCDate(MaskEdBox4.Text)) & ""
      End If
      'If Text4 = MsgText(602) And Text5 = MsgText(602) Then
      '   strSQL = strSQL & " and (((a0e14 is null or a0e14 = 0) and (a0e34 = 0 or a0e34 is null) and a0e15 = 0 and a0e16 = 0 and a0e17 = 0 and (a0e21 is null or a0e21 = 0) and (a0e34 = 0 or a0e34 is null)) or ((a0e14 is not null and a0e14 <> 0) and (a0e21 is null or a0e21 = 0) and (a0e34 = 0 or a0e34 is null)))"
      'Else
      '   If Text4 = MsgText(602) Then
      '      strSQL = strSQL & " and (a0e14 is null or a0e14 = 0) and (a0e34 = 0 or a0e34 is null) and a0e15 = 0 and a0e16 = 0 and a0e17 = 0 and (a0e21 is null or a0e21 = 0) and (a0e34 = 0 or a0e34 is null)"
      '   End If
      '   If Text5 = MsgText(602) Then
      '      strSQL = strSQL & " and (a0e14 is not null and a0e14 <> 0) and (a0e21 is null or a0e21 = 0) and (a0e34 = 0 or a0e34 is null)"
      '   Else
      '      If Text5 = MsgText(603) Then
      '         strSQL = strSQL & " and (((a0e14 is not null and a0e14 <> 0) and (a0e21 is not null and a0e21 <> 0)) or (a0e34 <> 0 and a0e34 is not null))"
      '      End If
      '   End If
      'End If
      If Text4 = MsgText(602) And Text5 = MsgText(602) Then
         strSql = strSql & " and (((a0e14 is null or a0e14 = 0) and (a0e21 is null or a0e21 = 0)) or ((a0e14 is not null and a0e14 <> 0) and (a0e21 is null or a0e21 = 0))) and (a0e15 = 0 or a0e15 is null)"
      Else
         If Text4 = MsgText(602) Then
            strSql = strSql & " and (a0e14 is null or a0e14 = 0) and (a0e21 is null or a0e21 = 0) and (a0e15 = 0 or a0e15 is null)"
         End If
         If Text5 = MsgText(602) Then
            strSql = strSql & " and (a0e14 is not null and a0e14 <> 0) and (a0e21 is null or a0e21 = 0 and (a0e15 = 0 or a0e15 is null))"
         End If
      End If
      strUnion = "select a0e01, a0e02, a0g02, a0e13, a0e10, a0e11, a0e14, a0e15, a0e16, a0e03, a0e12, cu04 as contect, a0e17, a0e20, a0e34, a0e21 from acc0e0, acc0g0, customer where a0e01 = a0g01 and a0e05 = '1' and substr(a0e06, 1, 8) = cu01 (+) and substr(a0e06, 9, 1) = cu02 (+) and a0e04 = '" & MsgText(18) & "'" & strSql
      strUnion = strUnion & " union select a0e01, a0e02, a0g02, a0e13, a0e10, a0e11, a0e14, a0e15, a0e16, a0e03, a0e12, a0i02 as contect, a0e17, a0e20, a0e34, a0e21 from acc0e0, acc0g0, acc0i0 where a0e01 = a0g01 and a0e05 = '2' and a0e06 = a0i01 (+) and a0e04 = '" & MsgText(18) & "'" & strSql
      strUnion = strUnion & " union select a0e01, a0e02, a0g02, a0e13, a0e10, a0e11, a0e14, a0e15, a0e16, a0e03, a0e12, st02 as contect, a0e17, a0e20, a0e34, a0e21 from acc0e0, acc0g0, staff where a0e01 = a0g01 and a0e05 = '3' and a0e06 = st01 (+) and a0e04 = '" & MsgText(18) & "'" & strSql
      strUnion = strUnion & " union select a0e01, a0e02, a0g02, a0e13, a0e10, a0e11, a0e14, a0e15, a0e16, a0e03, a0e12, '' as contect, a0e17, a0e20, a0e34, a0e21 from acc0e0, acc0g0 where a0e01 = a0g01 and a0e05 = '4' and a0e04 = '" & MsgText(18) & "'" & strSql
      '92.11.11 add by sonia
      If Text7 = MsgText(602) Then
         '20140122START Modify By eric
         'Modify By Sindy 2020/4/17 公司別改變數
'         If Text2 <> "" Then
'            strSQL1 = " and a0e23 = '" & IIf(Text2 = "2", "J", "1") & "' "
'         Else
'            strSQL1 = ""
'         End If
         If strCmp <> MsgText(601) Then
             If InStr(strCmp, "+") > 0 Then
                strSQL1 = " And (a0e23 is null or a0e23 In ('" & Replace(strCmp, "+", "','") & "')) "
             Else
                strSQL1 = " And (a0e23 is null or a0e23='" & strCmp & "') "
             End If
         Else
            strSQL1 = ""
         End If
         '2020/4/17 END
         
         If Combo1 <> MsgText(601) Then
            strSQL1 = strSQL1 & " and a0e20 >= '" & Combo1 & "'"
         End If
         'If Combo1 <> MsgText(601) Then
         '   strSQL1 = " and a0e20 >= '" & Combo1 & "'"
         'End If
         '20140122END
         If Combo2 <> MsgText(601) Then
            strSQL1 = strSQL1 & " and a0e20 <= '" & Combo2 & "'"
         End If
         If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
            strSQL1 = strSQL1 & " and a0e13 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
         End If
         If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
            strSQL1 = strSQL1 & " and a0e13 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
         End If
         If MaskEdBox5.Text <> MsgText(601) And MaskEdBox5.Text <> MsgText(29) Then
            strSQL1 = strSQL1 & " and a0e21 >= " & Val(FCDate(MaskEdBox5.Text)) & ""
         End If
         If MaskEdBox6.Text <> MsgText(601) And MaskEdBox6.Text <> MsgText(29) Then
            strSQL1 = strSQL1 & " and a0e21 <= " & Val(FCDate(MaskEdBox6.Text)) & ""
         End If
         strUnion = strUnion & " union select a0e01, a0e02, a0g02, a0e13, a0e10, a0e11, a0e14, a0e15, a0e16, a0e03, a0e12, cu04 as contect, a0e17, a0e20, a0e34, a0e21 from acc0e0, acc0g0, customer where a0e01 = a0g01 and a0e05 = '1' and substr(a0e06, 1, 8) = cu01 (+) and substr(a0e06, 9, 1) = cu02 (+) and a0e04 = '" & MsgText(18) & "'" & strSQL1
         strUnion = strUnion & " union select a0e01, a0e02, a0g02, a0e13, a0e10, a0e11, a0e14, a0e15, a0e16, a0e03, a0e12, a0i02 as contect, a0e17, a0e20, a0e34, a0e21 from acc0e0, acc0g0, acc0i0 where a0e01 = a0g01 and a0e05 = '2' and a0e06 = a0i01 (+) and a0e04 = '" & MsgText(18) & "'" & strSQL1
         strUnion = strUnion & " union select a0e01, a0e02, a0g02, a0e13, a0e10, a0e11, a0e14, a0e15, a0e16, a0e03, a0e12, st02 as contect, a0e17, a0e20, a0e34, a0e21 from acc0e0, acc0g0, staff where a0e01 = a0g01 and a0e05 = '3' and a0e06 = st01 (+) and a0e04 = '" & MsgText(18) & "'" & strSQL1
         strUnion = strUnion & " union select a0e01, a0e02, a0g02, a0e13, a0e10, a0e11, a0e14, a0e15, a0e16, a0e03, a0e12, '' as contect, a0e17, a0e20, a0e34, a0e21 from acc0e0, acc0g0 where a0e01 = a0g01 and a0e05 = '4' and a0e04 = '" & MsgText(18) & "'" & strSQL1
      End If
      strUnion = strUnion & " order by a0e13 asc, a0e02 asc"
      '92.11.11 end
      adoadodc1.Open strUnion, adoTaie, adOpenStatic, adLockReadOnly
   Else
      '20140122START Add By eric
      'Modify By Sindy 2020/4/17 公司別改變數
'      If Text2 <> "" Then
'         strSQL2 = " and a0e23 = '" & IIf(Text2 = "2", "J", "1") & "'"
'      Else
'         strSQL2 = ""
'      End If
      If strCmp <> MsgText(601) Then
          If InStr(strCmp, "+") > 0 Then
             strSQL2 = " And (a0e23 is null or a0e23 In ('" & Replace(strCmp, "+", "','") & "')) "
          Else
             strSQL2 = " And (a0e23 is null or a0e23='" & strCmp & "') "
          End If
      Else
         strSQL2 = ""
      End If
      '2020/4/17 END
      strUnion = "select a0e01, a0e02, a0g02, a0e13, a0e10, a0e11, a0e14, a0e15, a0e16, a0e03, a0e12, cu04 as contect, a0e17, a0e20, a0e34, a0e21 from acc0e0, acc0g0, customer where a0e01 = a0g01 and a0e05 = '1' and substr(a0e06, 1, 8) = cu01 (+) and substr(a0e06, 9, 1) = cu02 (+) and a0e02 = '" & Text3 & "' and a0e04 = '" & MsgText(18) & "'" & strSQL2
      strUnion = strUnion & " union select a0e01, a0e02, a0g02, a0e13, a0e10, a0e11, a0e14, a0e15, a0e16, a0e03, a0e12, a0i02 as contect, a0e17, a0e20, a0e34, a0e21 from acc0e0, acc0g0, acc0i0 where a0e01 = a0g01 and a0e05 = '2' and a0e06 = a0i01 (+) and a0e02 = '" & Text3 & "' and a0e04 = '" & MsgText(18) & "'" & strSQL2
      strUnion = strUnion & " union select a0e01, a0e02, a0g02, a0e13, a0e10, a0e11, a0e14, a0e15, a0e16, a0e03, a0e12, st02 as contect, a0e17, a0e20, a0e34, a0e21 from acc0e0, acc0g0, staff where a0e01 = a0g01 and a0e05 = '3' and a0e06 = st01 (+) and a0e02 = '" & Text3 & "' and a0e04 = '" & MsgText(18) & "'" & strSQL2
      strUnion = strUnion & " union select a0e01, a0e02, a0g02, a0e13, a0e10, a0e11, a0e14, a0e15, a0e16, a0e03, a0e12, '' as contect, a0e17, a0e20, a0e34, a0e21 from acc0e0, acc0g0 where a0e01 = a0g01 and a0e05 = '4' and a0e02 = '" & Text3 & "' and a0e04 = '" & MsgText(18) & "' " & strSQL2 & " order by a0e13 asc, a0e02 asc"
      'strUnion = "select a0e01, a0e02, a0g02, a0e13, a0e10, a0e11, a0e14, a0e15, a0e16, a0e03, a0e12, cu04 as contect, a0e17, a0e20, a0e34, a0e21 from acc0e0, acc0g0, customer where a0e01 = a0g01 and a0e05 = '1' and substr(a0e06, 1, 8) = cu01 (+) and substr(a0e06, 9, 1) = cu02 (+) and a0e02 = '" & Text3 & "' and a0e04 = '" & MsgText(18) & "'"
      'strUnion = strUnion & " union select a0e01, a0e02, a0g02, a0e13, a0e10, a0e11, a0e14, a0e15, a0e16, a0e03, a0e12, a0i02 as contect, a0e17, a0e20, a0e34, a0e21 from acc0e0, acc0g0, acc0i0 where a0e01 = a0g01 and a0e05 = '2' and a0e06 = a0i01 (+) and a0e02 = '" & Text3 & "' and a0e04 = '" & MsgText(18) & "'"
      'strUnion = strUnion & " union select a0e01, a0e02, a0g02, a0e13, a0e10, a0e11, a0e14, a0e15, a0e16, a0e03, a0e12, st02 as contect, a0e17, a0e20, a0e34, a0e21 from acc0e0, acc0g0, staff where a0e01 = a0g01 and a0e05 = '3' and a0e06 = st01 (+) and a0e02 = '" & Text3 & "' and a0e04 = '" & MsgText(18) & "'"
      'strUnion = strUnion & " union select a0e01, a0e02, a0g02, a0e13, a0e10, a0e11, a0e14, a0e15, a0e16, a0e03, a0e12, '' as contect, a0e17, a0e20, a0e34, a0e21 from acc0e0, acc0g0 where a0e01 = a0g01 and a0e05 = '4' and a0e02 = '" & Text3 & "' and a0e04 = '" & MsgText(18) & "' order by a0e13 asc, a0e02 asc"
      '20140122END
      adoadodc1.Open strUnion, adoTaie, adOpenStatic, adLockReadOnly
   End If
   Adodc1.Recordset.Requery
   Text1 = Adodc1.Recordset.RecordCount 'Add by Morgan 2009/4/8
   Text6 = MsgText(601)    'add by sonia 2024/4/29 否則無資料時會保留前次結果
   If Adodc1.Recordset.RecordCount = 0 Then
      Adodc1.Recordset.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   Else
      SumShow
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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
            AdodcRefresh
            Screen.MousePointer = vbDefault
            Exit Sub
         Else
            MsgBox MsgText(181), , MsgText(5)
         End If
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  計算並顯示合計
'
'*************************************************
Public Sub SumShow()
'Modified by Morgan 2012/11/12 改抓Grid資料加總
'   adoaccsum.CursorLocation = adUseClient
'   If Text3 = MsgText(601) Then
'      adoaccsum.Open "select SUM(A0E11) from ACC0E0 WHERE a0e04 = '" & MsgText(18) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'   Else
'      adoaccsum.Open "select SUM(A0E11) from ACC0E0 WHERE a0e02 = '" & Text3 & "' and a0e04 = '" & MsgText(18) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'   End If
'   If adoaccsum.RecordCount <> 0 Then
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         Text6 = MsgText(601)
'      Else
'         Text6 = adoaccsum.Fields(0).Value
'      End If
'   Else
'      Text6 = MsgText(601)
'   End If
'   adoaccsum.Close
'   '92.11.11 MODIFY by sonia
'   If Text7 = MsgText(602) Then
'      adoaccsum.CursorLocation = adUseClient
'      adoaccsum.Open "select SUM(A0E11) from ACC0E0 WHERE a0e04 = '" & MsgText(18) & "'" & strSQL1, adoTaie, adOpenStatic, adLockReadOnly
'      If adoaccsum.RecordCount <> 0 Then
'         If IsNull(adoaccsum.Fields(0).Value) Then
'         Else
'            Text6 = Text6 + adoaccsum.Fields(0).Value
'         End If
'      End If
'      adoaccsum.Close
'      End If
'   Text6 = Format(Text6, DDollar)
'   '92.11.11 END
   Text6 = ""
   Set adoaccsum = Adodc1.Recordset.Clone
   With adoaccsum
   .MoveFirst
   Do While Not .EOF
      Text6 = Val(Text6) + .Fields("a0e11")
      .MoveNext
   Loop
   End With
   Text6 = Format(Text6, DDollar)
'end 2012/11/12
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Text3 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Combo1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Combo2 <> MsgText(601) Then
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
   If MaskEdBox3.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox4.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   '92.11.10 add by sonia
   If MaskEdBox5.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox6.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   '92.11.10 end
   FormCheck = False
End Function

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Modify by Sindy 2020/04/17 公司別改下拉
''20140122START By eric
'Private Sub Text2_LostFocus()
'   If Text2.Text <> "1" And Text2.Text <> "2" And Text2.Text <> "" Then
'      MsgBox "公司別僅可為 1 / 2 或不輸入  ! "
'      Text2.Text = ""
'      Text2.SetFocus
'      Exit Sub
'   End If
'End Sub
'
''20140122START By eric
'Private Sub Text2_GotFocus()
'   TextInverse Text2
'   CloseIme
'End Sub
'
''20140122START By eric
'Private Sub Text2_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub
