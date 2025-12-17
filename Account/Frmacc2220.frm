VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc2220 
   AutoRedraw      =   -1  'True
   Caption         =   "國外案件帳目查詢 "
   ClientHeight    =   5436
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8988
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5436
   ScaleWidth      =   8988
   Begin VB.TextBox Text17 
      Alignment       =   1  '靠右對齊
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
      Height          =   330
      Left            =   5370
      TabIndex        =   40
      Top             =   4875
      Width           =   1125
   End
   Begin VB.ComboBox Combo5 
      Height          =   300
      Left            =   7260
      TabIndex        =   38
      Top             =   4515
      Width           =   1575
   End
   Begin VB.ComboBox Combo4 
      Height          =   300
      Left            =   5220
      TabIndex        =   37
      Top             =   4515
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "Frmacc2220.frx":0000
      Left            =   1290
      List            =   "Frmacc2220.frx":0002
      TabIndex        =   36
      Top             =   4515
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      Left            =   3330
      TabIndex        =   35
      Top             =   4515
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   4
      Top             =   495
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   3
      Top             =   495
      Width           =   255
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   2
      Top             =   495
      Width           =   852
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3240
      TabIndex        =   5
      Top             =   495
      Width           =   612
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2220.frx":0004
      Height          =   2385
      Left            =   105
      TabIndex        =   12
      Top             =   2070
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   4212
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "國外案件帳目查詢"
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "DocNo"
         Caption         =   "單據編號"
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
      BeginProperty Column01 
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
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "Famount"
         Caption         =   "外幣金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Namount"
         Caption         =   "台幣金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Tamount"
         Caption         =   "規費"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
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
      BeginProperty Column07 
         DataField       =   "FagentNo"
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
         DataField       =   "a1k1316"
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
      BeginProperty Column09 
         DataField       =   "a1k28"
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
      BeginProperty Column10 
         DataField       =   "a1k27"
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
      BeginProperty Column11 
         DataField       =   "a1k30"
         Caption         =   "已收金額(台幣)"
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
      BeginProperty Column12 
         DataField       =   "a1k10"
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
      BeginProperty Column13 
         DataField       =   "a1k12"
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
      BeginProperty Column14 
         DataField       =   "a1k25"
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
      BeginProperty Column15 
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
         Size            =   275
         BeginProperty Column00 
            ColumnWidth     =   1235.906
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   527.811
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1091.906
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1247.811
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   527.811
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1272.189
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1644.095
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column15 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "單據內容"
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
      Left            =   7200
      TabIndex        =   11
      Top             =   120
      Width           =   1212
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  '靠右對齊
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
      Height          =   330
      Left            =   1290
      TabIndex        =   23
      Top             =   4875
      Width           =   1575
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  '靠右對齊
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
      Height          =   330
      Left            =   7260
      TabIndex        =   22
      Top             =   4875
      Width           =   1575
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  '靠右對齊
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
      Height          =   330
      Left            =   3330
      TabIndex        =   21
      Top             =   4875
      Width           =   1545
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   10
      Top             =   1620
      Width           =   612
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6840
      TabIndex        =   7
      Top             =   870
      Width           =   1572
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   0
      Top             =   120
      Width           =   1572
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      TabIndex        =   6
      Top             =   870
      Width           =   3492
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
      Height          =   330
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   1
      Top             =   495
      Width           =   492
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   330
      Left            =   1320
      TabIndex        =   8
      Top             =   1245
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   572
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   330
      Left            =   3240
      TabIndex        =   9
      Top             =   1245
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   572
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
      Height          =   330
      Left            =   240
      Top             =   1950
      Visible         =   0   'False
      Width           =   1200
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
   Begin MSForms.TextBox Text2 
      Height          =   330
      Left            =   3870
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   495
      Width           =   4575
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "8070;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label20 
      BackStyle       =   0  '透明
      Caption         =   "未收規費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4920
      TabIndex        =   39
      Top             =   4830
      Width           =   495
   End
   Begin VB.Label Label19 
      BackStyle       =   0  '透明
      Caption         =   "(若只輸入代理人D/N No.,             查詢資料請輸入3或4)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   5310
      TabIndex        =   34
      Top             =   1200
      Width           =   3105
   End
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "(*作廢、@有折讓、$銷帳、>付款中)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   33
      Top             =   158
      Width           =   4095
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "未付"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6810
      TabIndex        =   32
      Top             =   4530
      Width           =   495
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "未收"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2850
      TabIndex        =   31
      Top             =   4530
      Width           =   495
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "CF"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   30
      Top             =   4530
      Width           =   375
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "FC"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   990
      TabIndex        =   29
      Top             =   4530
      Width           =   375
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "外幣合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   30
      TabIndex        =   28
      Top             =   4530
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   60
      Top             =   2130
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "台幣合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   30
      TabIndex        =   27
      Top             =   4905
      Width           =   975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "FC"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   990
      TabIndex        =   26
      Top             =   4905
      Width           =   375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "CF已付"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6540
      TabIndex        =   25
      Top             =   4905
      Width           =   855
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "未收"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2850
      TabIndex        =   24
      Top             =   4905
      Width           =   495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "查詢資料："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   1658
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "(1.FC往來 2.FC未收 3.CF往來 4.CF未付 5.往來 6.未收未付)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   19
      Top             =   1658
      Width           =   6495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   18
      Top             =   1230
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "往來日期："
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
      TabIndex        =   17
      Top             =   1284
      Width           =   972
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "代理人D/N No："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   16
      Top             =   908
      Width           =   1665
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "請款單號："
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
      TabIndex        =   15
      Top             =   159
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "申請案號："
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
      TabIndex        =   14
      Top             =   909
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "本所案號："
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
      TabIndex        =   13
      Top             =   534
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc2220"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/09 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Text2
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
'Public adotmp2220 As New ADODB.Recordset 'Removed by Morgan 2025/9/9 沒用了
Dim strSql As String
Dim strWhere(6) As String

Private Sub Combo1_Click()
   CaseQuery
End Sub

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
   
   'Modify by Morgan 2004/8/17 先讀單號去掉符號再判斷
   strItemNo = Adodc1.Recordset.Fields("DocNo").Value
   'modify by sonia  2015/4/21 加去掉尾巴的>符號
   If Right(strItemNo, 1) = "*" Or Right(strItemNo, 1) = "@" Or Right(strItemNo, 1) = "$" Or Right(strItemNo, 1) = ">" Then
      strItemNo = Left(strItemNo, Len(strItemNo) - 1)
   End If
   
   Select Case Left(strItemNo, 1)
      Case MsgText(815)
         Frmacc2211.Show
      Case MsgText(808)
         Frmacc2212.Show
      'Modify by Morgan 2004/9/2 國外抵帳單編號 V 的也要
      'Case MsgText(812)
      Case MsgText(812), MsgText(813)
         Frmacc2213.m_CaseNo = Text1 & Text3 & Text4 & Text6 'Add by Amy 2025/01/14 +m_CaseNo 本所案號
         Frmacc2213.Show
      Case MsgText(814)
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
'   Me.Width = 9030 '8850
'   Me.Height = 5490 '5085 '5500
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
   'Modify by Amy 原:W9030 H5700
   PUB_InitForm Me, 9080, 5880, strBackPicPath2
   'end 2021/12/09
   
   Combo1.AddItem ComboItem(121)
   Combo1.AddItem ComboItem(122)
   Combo1.AddItem ComboItem(123)
   Combo1 = ComboItem(121)
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   OpenTable
   StatusView MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc2220 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   CaseQuery
End Sub

Private Sub Text3_GotFocus()
   'TextInverse Text3
   InverseTextBox Text3
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   Text4 = "0"
   Text6 = "00"
   
   'Add By Sindy 2009/09/18
   If Trim(Text1) = "S" And Trim(Text3) <> "" Then
      strSql = "select * from servicepractice where sp01='" & Trim(Text1) & "' and sp02='" & Trim(Text3) & "' and sp03='" & IIf(Trim(Text4) = "", "0", Trim(Text4)) & "' and sp04='" & IIf(Trim(Text6) = "", "00", Trim(Text6)) & "' " & _
                      "and instr(sp18,'轉入商標') >0 "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         MsgBox "此案已" & Trim(RsTemp.Fields("sp18"))
         Call Text3_GotFocus
         Exit Sub
      End If
   End If
   '2009/09/18 End
   
   CaseQuery
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
   CaseQuery
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   CaseQuery
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
   CloseIme
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
   CloseIme
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   'Removed by Morgan 2025/9/9 沒用了
   'adotmp2220.CursorLocation = adUseClient
   'adotmp2220.Open "select * from tmp2220 order by t22201", adoTemp, adOpenDynamic, adLockBatchOptimistic
   'end 2025/9/9
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
Dim StrSQLa As String
Dim StrSqlB As String
Dim Str01 As String, Str02 As String, Str03 As String, Str04 As String, strCaseNo As String 'Add By Sindy 2014/3/24
   
On Error GoTo Checking
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/21 清除查詢印表記錄檔欄位
   If Trim(Text7) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label3 & Text7 'Add By Sindy 2010/12/21
   End If
   If Trim(Text1) <> "" And Trim(Text3) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1 & Text1 & "-" & Text3 & "-" & Text4 & "-" & Text6 'Add By Sindy 2010/12/21
   End If
   If Trim(Text5) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label2 & Text5 'Add By Sindy 2010/12/21
   End If
   If Trim(Text8) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label4 & Text8 'Add By Sindy 2010/12/21
   End If
   If (MaskEdBox1.Text <> MsgText(29) And MaskEdBox1.Text <> "") Or _
      (MaskEdBox2.Text <> MsgText(29) And MaskEdBox2.Text <> "") Then
      pub_QL05 = pub_QL05 & ";" & Label5 & MaskEdBox1 & "-" & MaskEdBox2 'Add By Sindy 2010/12/21
   End If
   If Trim(Text9) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label8 & Text9 & Label7 'Add By Sindy 2010/12/21
   End If
   
   strSql = ""
   For intCounter = 0 To 5
      strWhere(intCounter) = ""
   Next intCounter
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
    '請款單號
   If Text7 <> "" Then
      strWhere(0) = strWhere(0) & " and a1k01 = '" & Text7 & "'"
      strWhere(1) = strWhere(1) & " and a0z02 = '" & Text7 & "'"
   End If
    '本所案號
   If Text1 <> "" Then
      strWhere(0) = strWhere(0) & " and a1k13 = '" & Text1 & "'"
      strWhere(1) = strWhere(1) & " and a1k13 = '" & Text1 & "'"
      strWhere(2) = strWhere(2) & " and axf03 = '" & Text1 & Text3 & Text4 & Text6 & "'"
      strWhere(3) = strWhere(3) & " and axf03 = '" & Text1 & Text3 & Text4 & Text6 & "'"
      strWhere(5) = strWhere(5) & " and axg03 = '" & Text1 & Text3 & Text4 & Text6 & "'"
   End If
   If Text3 <> "" Then
      strWhere(0) = strWhere(0) & " and a1k14 = '" & Text3 & "'"
      strWhere(1) = strWhere(1) & " and a1k14 = '" & Text3 & "'"
   End If
   If Text3 <> "" And Text4 <> "" Then
      strWhere(0) = strWhere(0) & " and a1k15 = '" & Text4 & "'"
      strWhere(1) = strWhere(1) & " and a1k15 = '" & Text4 & "'"
   End If
   If Text3 <> "" And Text6 <> "" Then
      strWhere(0) = strWhere(0) & " and a1k16 = '" & Text6 & "'"
      strWhere(1) = strWhere(1) & " and a1k16 = '" & Text6 & "'"
   End If
    '代理人D/N No.
   If Text8 <> "" Then
      strWhere(2) = strWhere(2) & " and a1504 = '" & Text8 & "'"
      strWhere(3) = strWhere(3) & " and a1504 = '" & Text8 & "'"
      '94.1.3 add by sonia
      strWhere(5) = strWhere(5) & " and a1604 = '" & Text8 & "'"
      '94.1.3 end
   End If
    '往來日期
   If MaskEdBox1.Text <> MsgText(29) And MaskEdBox1.Text <> "" Then
      strWhere(0) = strWhere(0) & " and a1k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      strWhere(1) = strWhere(1) & " and a0y02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      strWhere(2) = strWhere(2) & " and a1502 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      strWhere(3) = strWhere(3) & " and a1b03 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      strWhere(5) = strWhere(5) & " and a1602 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(29) And MaskEdBox2.Text <> "" Then
      strWhere(0) = strWhere(0) & " and a1k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      strWhere(1) = strWhere(1) & " and a0y02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      strWhere(2) = strWhere(2) & " and a1502 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      strWhere(3) = strWhere(3) & " and a1b03 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      strWhere(5) = strWhere(5) & " and a1602 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   'Modify By Sindy 2010/3/12 增加a1k30,a1k10,a1k12,a1k25,strUserNum
   '2011/9/6 modify by sonia 增加USDamount
   'Modify By Sindy 2012/8/24 有關收款單的規費計算原抓取 sum(A1K09) 改為 sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09)))
   'Modify By Sindy 2014/3/24 +,null as PA161
   Select Case Text9
      Case "1" 'FC往來
'         If Text5 <> "" Then
'            strWhere(4) = " and pa11 = '" & Text5 & "'"
'         End If
'         strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort , a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,(a1k08 - nvl(a1k06, 0))) USDamount from acc1k0, fagent, nation, patent, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4)
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort , a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount from acc0y0, fagent, nation, acc0z0, acc1k0, patent where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04" & strWhere(1) & strWhere(4) & _
'                           " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'         If Text5 <> "" Then
'            strWhere(4) = " and tm12 = '" & Text5 & "'"
'         End If
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,(a1k08 - nvl(a1k06, 0))) USDamount from acc1k0, fagent, nation, trademark, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 (+) and a1k14 = tm02 (+) and a1k15 = tm03 (+) and a1k16 = tm04 (+) and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4)
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                           "from acc0y0, fagent, nation, acc0z0, acc1k0, trademark where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04" & strWhere(1) & strWhere(4) & " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'         If Text5 <> "" Then
'            strWhere(4) = " and sp11 = '" & Text5 & "'"
'         End If
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,(a1k08 - nvl(a1k06, 0))) USDamount from acc1k0, fagent, nation, servicepractice, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 (+) and a1k14 = sp02 (+) and a1k15 = sp03 (+) and a1k16 = sp04 (+) and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4)
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                                    "from acc0y0, fagent, nation, acc0z0, acc1k0, servicepractice where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04" & strWhere(1) & strWhere(4) & " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'         '2005/8/2 ADD BY SONIA 加入LAWCASE
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,(a1k08 - nvl(a1k06, 0))) USDamount from acc1k0, fagent, nation, LAWCASE, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = LC01 (+) and a1k14 = LC02 (+) and a1k15 = LC03 (+) and a1k16 = LC04 (+) and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4)
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                           "from acc0y0, fagent, nation, acc0z0, acc1k0, LAWCASE where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04" & strWhere(1) & strWhere(4) & " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'         '2005/8/2 END
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,(a1k08 - nvl(a1k06, 0))) USDamount from acc1k0, fagent, nation, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0)
'         '2010/6/29 MODIFY BY SONIA 抵帳幣別不可抓A1K18請款幣別X09901818
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1k17 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, NVL(A1H03,A1I05) as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k08 - nvl(a1k06, 0)) * nvl(a1g02, 0) as Namount, null as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort , a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount from acc1k0, fagent, nation, acc140, acc1g0, acc1h0, acc1i0 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k01 = a1403 (+) and a1k17 = a1g01 (+) and a1k17 = a1h01 (+) and a1k17 = a1i01 (+) and a1k17 is not null" & strWhere(0)
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                           "from acc0y0, fagent, nation, acc0z0, acc1k0 where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01" & strWhere(1) & " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'專利
         If Text5 <> "" Then
            strWhere(4) = " and pa11 = '" & Text5 & "'"
         End If
'請款單
         strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort , a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc1k0, fagent, nation, patent, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4)
'收款
         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort , a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc0y0, fagent, nation, acc0z0, acc1k0, patent where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04" & strWhere(1) & strWhere(4) & _
                           " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'商標
         If Text5 <> "" Then
            strWhere(4) = " and tm12 = '" & Text5 & "'"
         End If
'請款單
         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc1k0, fagent, nation, trademark, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 (+) and a1k14 = tm02 (+) and a1k15 = tm03 (+) and a1k16 = tm04 (+) and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4)
'收款
        strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                           "from acc0y0, fagent, nation, acc0z0, acc1k0, trademark where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04" & strWhere(1) & strWhere(4) & " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'服務業務
         If Text5 <> "" Then
            strWhere(4) = " and sp11 = '" & Text5 & "'"
         End If
'請款單
         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc1k0, fagent, nation, servicepractice, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 (+) and a1k14 = sp02 (+) and a1k15 = sp03 (+) and a1k16 = sp04 (+) and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4)
'收款
         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                    "from acc0y0, fagent, nation, acc0z0, acc1k0, servicepractice where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04" & strWhere(1) & strWhere(4) & " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'法務2005/8/2 ADD BY SONIA 加入LAWCASE
'請款單
        'Modified by Lydia 2018/02/13 拿掉strWhere(4); 有基本檔才抓
        'strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc1k0, fagent, nation, LAWCASE, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = LC01 (+) and a1k14 = LC02 (+) and a1k15 = LC03 (+) and a1k16 = LC04 (+) and a1k01 = a1403 (+)" & strWhere(0)
        If Trim(Text7 & Text1 & Text3 & Text8) <> "" Then
            strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161" & _
                                     " from acc1k0, fagent, nation, LAWCASE, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = LC01 (+) and a1k14 = LC02 (+) and a1k15 = LC03 (+) and a1k16 = LC04 (+) and a1k01 = a1403 (+)" & strWhere(0) & _
                                     " and lc01 is not null "
        End If
        'end 2018/02/13
'收款
        'Modified by Lydia 2018/02/13 拿掉strWhere(4) ; 有基本檔才抓
        'strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                  "from acc0y0, fagent, nation, acc0z0, acc1k0, LAWCASE where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04" & strWhere(1) & strWhere(4)  & " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
        If Trim(Text7 & Text1 & Text3 & Text8) <> "" Then
            strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161" & _
                                      " from acc0y0, fagent, nation, acc0z0, acc1k0, LAWCASE where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04" & strWhere(1) & _
                                      " and lc01 is not null" & _
                                      " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
        End If
        'end 2018/02/13
         '2005/8/2 END
'抓不到案號的請款單
         'Modified by Lydia 2018/02/13 比對基本檔無資料才算舊資料
         'strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc1k0, fagent, nation, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0)
         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161" & _
                                   " from acc1k0, fagent, nation, acc140,PATENT,TRADEMARK,SERVICEPRACTICE,LAWCASE" & _
                                   " where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & _
                                   " AND A1K13=PA01(+) AND A1K14=PA02(+) AND A1K15=PA03(+) AND A1K16=PA04(+)" & _
                                   " AND A1K13=TM01(+) AND A1K14=TM02(+) AND A1K15=TM03(+) AND A1K16=TM04(+)" & _
                                   " AND A1K13=SP01(+) AND A1K14=SP02(+) AND A1K15=SP03(+) AND A1K16=SP04(+)" & _
                                   " AND A1K13=LC01(+) AND A1K14=LC02(+) AND A1K15=LC03(+) AND A1K16=LC04(+)" & _
                                   " AND PA01||TM01||SP01||LC01 IS NULL"
'抓不到案號的收款
         'Modified by Lydia 2018/02/13 比對基本檔無資料才算舊資料
         'strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                           "from acc0y0, fagent, nation, acc0z0, acc1k0 where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01" & strWhere(1) & " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161" & _
                           " from acc0y0, fagent, nation, acc0z0, acc1k0,PATENT,TRADEMARK,SERVICEPRACTICE,LAWCASE" & _
                           " where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01" & strWhere(1) & _
                           " AND A1K13=PA01(+) AND A1K14=PA02(+) AND A1K15=PA03(+) AND A1K16=PA04(+)" & _
                           " AND A1K13=TM01(+) AND A1K14=TM02(+) AND A1K15=TM03(+) AND A1K16=TM04(+)" & _
                           " AND A1K13=SP01(+) AND A1K14=SP02(+) AND A1K15=SP03(+) AND A1K16=SP04(+)" & _
                           " AND A1K13=LC01(+) AND A1K14=LC02(+) AND A1K15=LC03(+) AND A1K16=LC04(+)" & _
                           " AND PA01||TM01||SP01||LC01 IS NULL" & _
                           " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'抵帳資料收款
         '2010/6/29 MODIFY BY SONIA 抵帳幣別不可抓A1K18請款幣別X09901818
         If Trim(strWhere(0)) <> "" Then 'Added by Lydia 2018/02/13 無條件不可加入
              strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1k17 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, NVL(A1H03,A1I05) as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k08 - nvl(a1k31, 0)) * nvl(a1g02, 0) as Namount, null as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort , a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc1k0, fagent, nation, acc140, acc1g0, acc1h0, acc1i0 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k01 = a1403 (+) and a1k17 = a1g01 (+) and a1k17 = a1h01 (+) and a1k17 = a1i01 (+) and a1k17 is not null" & strWhere(0)
         End If
      Case "2" 'FC未收
'         If Text5 <> "" Then
'            strWhere(4) = " and pa11 = '" & Text5 & "'"
'         End If
'         '2007/12/10 modify by sonia X09607651分次收款,婧瑄說台幣金額扣除已收金額,外幣改為台幣金額扣除已收金額/請款匯率
'         'strSQL = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27 from acc1k0, fagent, nation, patent, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         'Modify By Sindy 2012/6/25 X09607651分次收款未收金額部分改回原程式寫法,以X10003936測試
''         strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
''                      "from acc1k0, fagent, nation, patent, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '') "
'         strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                  "from acc1k0, fagent, nation, patent, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '') "
'         '2012/6/25 End
'         '2007/12/10 end
'         'Add By Sindy 2012/8/14 +部分收款時要同時帶出收款資料
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort , a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount from acc0y0, fagent, nation, acc0z0, acc1k0, patent where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04" & strWhere(1) & strWhere(4) & _
'                           " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'         '2012/8/14 End
'         If Text5 <> "" Then
'            strWhere(4) = " and tm12 = '" & Text5 & "'"
'         End If
'         '2007/12/10 modify by sonia X09607651分次收款,婧瑄說台幣金額扣除已收金額,外幣改為台幣金額扣除已收金額/請款匯率
'         'strSQL = strSQL & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27 from acc1k0, fagent, nation, trademark, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         'Modify By Sindy 2012/6/25 X09607651分次收款未收金額部分改回原程式寫法,以X10003936測試
''         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
''                                    "from acc1k0, fagent, nation, trademark, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '') "
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                           "from acc1k0, fagent, nation, trademark, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '') "
'         '2012/6/25 End
'         'Add By Sindy 2012/8/14 +部分收款時要同時帶出收款資料
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                           "from acc0y0, fagent, nation, acc0z0, acc1k0, trademark where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04" & strWhere(1) & strWhere(4) & " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'         '2012/8/14 End
'         If Text5 <> "" Then
'            strWhere(4) = " and sp11 = '" & Text5 & "'"
'         End If
'         '2007/12/10 modify by sonia X09607651分次收款,婧瑄說台幣金額扣除已收金額,外幣改為台幣金額扣除已收金額/請款匯率
'         'strSQL = strSQL & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27 from acc1k0, fagent, nation, servicepractice, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         'Modify By Sindy 2012/6/25 X09607651分次收款未收金額部分改回原程式寫法,以X10003936測試
''         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
''                                    "from acc1k0, fagent, nation, servicepractice, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '') "
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                           "from acc1k0, fagent, nation, servicepractice, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '') "
'         '2012/6/25 End
'         'Add By Sindy 2012/8/14 +部分收款時要同時帶出收款資料
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                                    "from acc0y0, fagent, nation, acc0z0, acc1k0, servicepractice where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04" & strWhere(1) & strWhere(4) & " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'         '2012/8/14 End
'         '2005/8/2 ADD BY SONIA 加入LAWCASE
'         '2007/12/10 modify by sonia X09607651分次收款,婧瑄說台幣金額扣除已收金額,外幣改為台幣金額扣除已收金額/請款匯率
'         'strSQL = strSQL & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27 from acc1k0, fagent, nation, LAWCASE, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         'Modify By Sindy 2012/6/25 X09607651分次收款未收金額部分改回原程式寫法,以X10003936測試
''         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
''                                    "from acc1k0, fagent, nation, LAWCASE, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '') "
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                           "from acc1k0, fagent, nation, LAWCASE, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '') "
'         '2012/6/25 End
'         '2005/8/2 END
'         'Add By Sindy 2012/8/14 +部分收款時要同時帶出收款資料
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                           "from acc0y0, fagent, nation, acc0z0, acc1k0, LAWCASE where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04" & strWhere(1) & strWhere(4) & " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'         '2012/8/14 End
'         '2007/12/10 modify by sonia X09607651分次收款,婧瑄說台幣金額扣除已收金額,外幣改為台幣金額扣除已收金額/請款匯率
'         'strSQL = strSQL & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27 from acc1k0, fagent, nation, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
'         'strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID from acc1k0, fagent, nation, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') "
'專利
         If Text5 <> "" Then
            strWhere(4) = " and pa11 = '" & Text5 & "'"
         End If
'未結清請款單
         '2007/12/10 modify by sonia X09607651分次收款,婧瑄說台幣金額扣除已收金額,外幣改為台幣金額扣除已收金額/請款匯率
         'strSQL = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27 from acc1k0, fagent, nation, patent, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
         'Modify By Sindy 2012/6/25 X09607651分次收款未收金額部分改回原程式寫法,以X10003936測試
'         strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                      "from acc1k0, fagent, nation, patent, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '') "
         strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                  "from acc1k0, fagent, nation, patent, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '') "
         '2012/6/25 End
         '2007/12/10 end
'Add By Sindy 2012/8/14 +部分收款時要同時帶出收款資料
         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort , a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc0y0, fagent, nation, acc0z0, acc1k0, patent where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04" & strWhere(1) & strWhere(4) & _
                           " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'2012/8/14 End
'商標
         If Text5 <> "" Then
            strWhere(4) = " and tm12 = '" & Text5 & "'"
         End If
'未結清請款單
         '2007/12/10 modify by sonia X09607651分次收款,婧瑄說台幣金額扣除已收金額,外幣改為台幣金額扣除已收金額/請款匯率
         'strSQL = strSQL & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27 from acc1k0, fagent, nation, trademark, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
         'Modify By Sindy 2012/6/25 X09607651分次收款未收金額部分改回原程式寫法,以X10003936測試
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                                    "from acc1k0, fagent, nation, trademark, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '') "
         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                           "from acc1k0, fagent, nation, trademark, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '') "
         '2012/6/25 End
'Add By Sindy 2012/8/14 +部分收款時要同時帶出收款資料
         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                           "from acc0y0, fagent, nation, acc0z0, acc1k0, trademark where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04" & strWhere(1) & strWhere(4) & " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'2012/8/14 End
'服務業務
         If Text5 <> "" Then
            strWhere(4) = " and sp11 = '" & Text5 & "'"
         End If
'未結清請款單
         '2007/12/10 modify by sonia X09607651分次收款,婧瑄說台幣金額扣除已收金額,外幣改為台幣金額扣除已收金額/請款匯率
         'strSQL = strSQL & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27 from acc1k0, fagent, nation, servicepractice, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
         'Modify By Sindy 2012/6/25 X09607651分次收款未收金額部分改回原程式寫法,以X10003936測試
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                                    "from acc1k0, fagent, nation, servicepractice, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '') "
         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                           "from acc1k0, fagent, nation, servicepractice, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '') "
         '2012/6/25 End
'Add By Sindy 2012/8/14 +部分收款時要同時帶出收款資料
         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                    "from acc0y0, fagent, nation, acc0z0, acc1k0, servicepractice where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04" & strWhere(1) & strWhere(4) & " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'2012/8/14 End
'法務2005/8/2 ADD BY SONIA 加入LAWCASE
'未結清請款單
         '2007/12/10 modify by sonia X09607651分次收款,婧瑄說台幣金額扣除已收金額,外幣改為台幣金額扣除已收金額/請款匯率
         'strSQL = strSQL & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27 from acc1k0, fagent, nation, LAWCASE, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
         'Modify By Sindy 2012/6/25 X09607651分次收款未收金額部分改回原程式寫法,以X10003936測試
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                                    "from acc1k0, fagent, nation, LAWCASE, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '') "
         'Modified by Lydia 2018/02/13 拿掉strWhere(4) ; 有基本檔才抓
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                   "from acc1k0, fagent, nation, LAWCASE, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '') "
         '2012/6/25 End
         '2005/8/2 END
         If Trim(Text7 & Text1 & Text3 & Text8) <> "" Then
            strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161" & _
                                      " from acc1k0, fagent, nation, LAWCASE, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04 and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')" & _
                                      " and lc01 is not null"
         End If
         'end 2018/02/13
'Add By Sindy 2012/8/14 +部分收款時要同時帶出收款資料
         'Modified by Lydia 2018/02/13 拿掉strWhere(4) ; 有基本檔才抓
'        strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
'                          "from acc0y0, fagent, nation, acc0z0, acc1k0, LAWCASE where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04" & strWhere(1) & strWhere(4) & " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'2012/8/14 End
        If Trim(Text7 & Text1 & Text3 & Text8) <> "" Then
            strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161" & _
                              " from acc0y0, fagent, nation, acc0z0, acc1k0, LAWCASE where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04" & strWhere(1) & _
                              " and lc01 is not null and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
        End If
        'end 2018/02/13
         '2007/12/10 modify by sonia X09607651分次收款,婧瑄說台幣金額扣除已收金額,外幣改為台幣金額扣除已收金額/請款匯率
         'strSQL = strSQL & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27 from acc1k0, fagent, nation, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
         'strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID from acc1k0, fagent, nation, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '') "
      Case "3" 'CF往來
'         If Text7 = "" Then
'            If Text5 <> "" Then
'               strWhere(4) = " and pa11 = '" & Text5 & "'"
'            End If
'            strSql = "select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                            "from acc151, acc150, fagent, nation, patent where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1501 "
'            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, sum(axf04) as Famount, sum(axf04 * nvl(a1g03, 0)) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc151, acc150, fagent, nation, patent, acc1g0, acc1h0, acc1i0 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2) & strWhere(4) & _
'                              " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, a1512, decode(a1h02, null, a1i03, a1h02), a1505, decode(a1506, a1520, 'Y', ''), a1504, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1501 "
'            'Modify By Cheng 2003/11/25 acc190串到acc151時, 可能會一筆串到多筆
''            strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, acc151, acc150, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and a1908=a1b01(+)" & strWhere(3) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, PA01||'-'||PA02||'-'||PA03||'-'||PA04"
'            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axf04),台幣金額由sum(a1905)改為sum(axf04*a1906)
'            'StrSQLa = " Select axf01 As Ax1, axf03 As Ax3 From acc190, acc180, fagent, nation, acc151, acc150, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and a1908=a1b01(+)" & strWhere(3) & " group by axf01, axf03 "
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = pa01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = pa02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = pa03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = pa04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, PA01||'-'||PA02||'-'||PA03||'-'||PA04"
'            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
'            StrSQLa = " Select axf01 As Ax1, axf03 As Ax3, sum(axf04) As Ax4 From acc190, acc180, fagent, nation, acc151, acc150, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(3) & " group by axf01, axf03, axf04 "
'            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate," & _
'                         " a1903 as Currency, sum(ax4) as Famount, sum(ax4*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = pa01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = pa02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = pa03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = pa04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, A1802, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, PA01||'-'||PA02||'-'||PA03||'-'||PA04 "
'            '2007/12/6 end
'            '2006/3/29 ADD BY SONIA 抓抵帳單資料
'            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axg04),台幣金額由sum(a1905)改為sum(axg04*a1906)
'            'StrSqlB = " Select axG01 As Ax1, axG03 As Ax3 From acc190, acc180, fagent, nation, acc161, acc160, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = pa01 and substr(axG03, length(axG03) - 8, 6) = pa02 and substr(axG03, length(axG03) - 2, 1) = pa03 and substr(axG03, length(axG03) - 1, 2) = pa04 and a1908=a1b01(+)" & strWhere(5) & " group by axG01, axG03 "
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = pa01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = pa02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = pa03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = pa04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, PA01||'-'||PA02||'-'||PA03||'-'||PA04"
'            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
'            StrSqlB = " Select axG01 As Ax1, axG03 As Ax3, axG04 As Ax4 From acc190, acc180, fagent, nation, acc161, acc160, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = pa01 and substr(axG03, length(axG03) - 8, 6) = pa02 and substr(axG03, length(axG03) - 2, 1) = pa03 and substr(axG03, length(axG03) - 1, 2) = pa04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(5) & " group by axG01, axG03, axG04 "
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(ax4) * (-1) as Famount, sum(ax4*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, patent, acc1b0 " & _
'                              " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = pa01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = pa02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = pa03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = pa04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, PA01||'-'||PA02||'-'||PA03||'-'||PA04 "
'            '2007/12/6 end
'            '2006/3/29 END
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, pa01||'-'||pa02||'-'||pa03||'-'||pa04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation, patent " & _
'                              " where axg01 = a1601 and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = pa01 and substr(axg03, length(axg03) - 8, 6) = pa02 and substr(axg03, length(axg03) - 2, 1) = pa03 and substr(axg03, length(axg03) - 1, 2) = pa04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, pa01||'-'||pa02||'-'||pa03||'-'||pa04, a1601 "
'            If Text5 <> "" Then
'               strWhere(4) = " and tm12 = '" & Text5 & "'"
'            End If
'            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                                       "from acc151, acc150, fagent, nation, trademark where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1501 "
'            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, sum(axf04) as Famount, sum(axf04 * nvl(a1g03, 0)) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc151, acc150, fagent, nation, trademark, acc1g0, acc1h0, acc1i0 " & _
'                              " where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2) & strWhere(4) & _
'                              " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, a1512, decode(a1h02, null, a1i03, a1h02), a1505, decode(a1506, a1520, 'Y', ''), a1504, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1501 "
'            'Modify By Cheng 2003/11/25 acc190串到acc151時, 可能會一筆串到多筆
''            strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, acc151, acc150, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and a1908=a1b01(+)" & strWhere(3) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, TM01||'-'||TM02||'-'||TM03||'-'||TM04"
'            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axf04),台幣金額由sum(a1905)改為sum(axf04*a1906)
'            'StrSQLa = " Select axf01 as Ax1, axf03 As Ax3 From acc190, acc180, fagent, nation, acc151, acc150, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and a1908=a1b01(+)" & strWhere(3) & " group by axf01, axf03 "
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSQLa & " ) A1, acc150, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = tm01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = tm02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = tm03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = tm04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, TM01||'-'||TM02||'-'||TM03||'-'||TM04"
'            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
'            StrSQLa = " Select axf01 as Ax1, axf03 As Ax3, sum(axf04) As Ax4 From acc190, acc180, fagent, nation, acc151, acc150, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(3) & " group by axf01, axf03, axf04 "
'            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, sum(ax4) as Famount, sum(ax4*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, fagent, nation, (" & StrSQLa & " ) A1, acc150, trademark, acc1b0 " & _
'                         " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = tm01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = tm02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = tm03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = tm04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, A1802, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, TM01||'-'||TM02||'-'||TM03||'-'||TM04 "
'            '2007/12/6 end
'            '2006/3/30 ADD BY SONIA 抓抵帳單資料
'            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axg04),台幣金額由sum(a1905)改為sum(axg04*a1906)
'            'StrSqlB = " Select axG01 as Ax1, axG03 As Ax3 From acc190, acc180, fagent, nation, acc161, acc160, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = tm01 and substr(axG03, length(axG03) - 8, 6) = tm02 and substr(axG03, length(axG03) - 2, 1) = tm03 and substr(axG03, length(axG03) - 1, 2) = tm04 and a1908=a1b01(+)" & strWhere(5) & " group by axG01, axG03 "
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSqlB & " ) A1, acc160, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = tm01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = tm02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = tm03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = tm04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, TM01||'-'||TM02||'-'||TM03||'-'||TM04"
'            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
'            StrSqlB = " Select axG01 as Ax1, axG03 As Ax3, axG04 As Ax4 From acc190, acc180, fagent, nation, acc161, acc160, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = tm01 and substr(axG03, length(axG03) - 8, 6) = tm02 and substr(axG03, length(axG03) - 2, 1) = tm03 and substr(axG03, length(axG03) - 1, 2) = tm04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(5) & " group by axG01, axG03, axG04 "
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(ax4) * (-1) as Famount, sum(ax4*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, fagent, nation, (" & StrSqlB & " ) A1, acc160, trademark, acc1b0 " & _
'                              " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = tm01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = tm02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = tm03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = tm04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, TM01||'-'||TM02||'-'||TM03||'-'||TM04 "
'            '2007/12/6 end
'            '2006/3/30 END
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation, trademark " & _
'                              " where axg01 = a1601 and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = tm01 and substr(axg03, length(axg03) - 8, 6) = tm02 and substr(axg03, length(axg03) - 2, 1) = tm03 and substr(axg03, length(axg03) - 1, 2) = tm04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1601 "
'            If Text5 <> "" Then
'               strWhere(4) = " and sp11 = '" & Text5 & "'"
'            End If
'            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                                       "from acc151, acc150, fagent, nation, servicepractice where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1501 "
'            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, sum(axf04) as Famount, sum(axf04 * nvl(a1g03, 0)) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc151, acc150, fagent, nation, servicepractice, acc1g0, acc1h0, acc1i0 " & _
'                              " where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2) & strWhere(4) & _
'                              " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, a1512, decode(a1h02, null, a1i03, a1h02), a1505, decode(a1506, a1520, 'Y', ''), a1504, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1501 "
'            'Modify By Cheng 2003/11/25 acc190串到acc151時, 可能會一筆串到多筆
''            strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, acc151, acc150, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and a1908=a1b01(+)" & strWhere(3) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, SP01||'-'||SP02||'-'||SP03||'-'||SP04"
'            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axf04),台幣金額由sum(a1905)改為sum(axf04*a1906)
'            'StrSQLa = " select axf01 As Ax1, axf03 As Ax3 From acc190, acc180, fagent, nation, acc151, acc150, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and a1908=a1b01(+)" & strWhere(3) & " group by axf01, axf03 "
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = sp01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = sp02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = sp03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = sp04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, SP01||'-'||SP02||'-'||SP03||'-'||SP04"
'            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
'            StrSQLa = " select axf01 As Ax1, axf03 As Ax3, sum(axf04) As Ax4 From acc190, acc180, fagent, nation, acc151, acc150, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(3) & " group by axf01, axf03, axf04 "
'            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate," & _
'                         " a1903 as Currency, sum(ax4) as Famount, sum(ax4*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, servicepractice, acc1b0 " & _
'                         " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = sp01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = sp02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = sp03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = sp04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, A1802, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, SP01||'-'||SP02||'-'||SP03||'-'||SP04 "
'            '2007/12/6 end
'            '2006/3/30 ADD BY SONIA 抓抵帳單資料
'            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axg04),台幣金額由sum(a1905)改為sum(axg04*a1906)
'            'StrSqlB = " select axG01 As Ax1, axG03 As Ax3 From acc190, acc180, fagent, nation, acc161, acc160, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = sp01 and substr(axG03, length(axG03) - 8, 6) = sp02 and substr(axG03, length(axG03) - 2, 1) = sp03 and substr(axG03, length(axG03) - 1, 2) = sp04 and a1908=a1b01(+)" & strWhere(5) & " group by axG01, axG03 "
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = sp01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = sp02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = sp03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = sp04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, SP01||'-'||SP02||'-'||SP03||'-'||SP04"
'            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
'            StrSqlB = " select axG01 As Ax1, axG03 As Ax3, axG04 As Ax4 From acc190, acc180, fagent, nation, acc161, acc160, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = sp01 and substr(axG03, length(axG03) - 8, 6) = sp02 and substr(axG03, length(axG03) - 2, 1) = sp03 and substr(axG03, length(axG03) - 1, 2) = sp04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(5) & " group by axG01, axG03, axG04 "
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(ax4) * (-1) as Famount, sum(ax4*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, servicepractice, acc1b0 " & _
'                              " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = sp01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = sp02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = sp03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = sp04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, SP01||'-'||SP02||'-'||SP03||'-'||SP04 "
'            '2007/12/6 end
'            '2006/3/30 END
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, sp01||'-'||sp02||'-'||sp03||'-'||sp04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation, servicepractice " & _
'                              " where axg01 = a1601 and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = sp01 and substr(axg03, length(axg03) - 8, 6) = sp02 and substr(axg03, length(axg03) - 2, 1) = sp03 and substr(axg03, length(axg03) - 1, 2) = sp04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, sp01||'-'||sp02||'-'||sp03||'-'||sp04, a1601 "
'            '2005/8/2 ADD BY SONIA 加入LAWCASE
'            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                                       "from acc151, acc150, fagent, nation, LAWCASE where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1501 "
'            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, sum(axf04) as Famount, sum(axf04 * nvl(a1g03, 0)) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc151, acc150, fagent, nation, LAWCASE, acc1g0, acc1h0, acc1i0 " & _
'                              " where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04 and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2) & strWhere(4) & _
'                              " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, a1512, decode(a1h02, null, a1i03, a1h02), a1505, decode(a1506, a1520, 'Y', ''), a1504, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1501 "
'            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axf04),台幣金額由sum(a1905)改為sum(axf04*a1906)
'            'StrSQLa = " select axf01 As Ax1, axf03 As Ax3 From acc190, acc180, fagent, nation, acc151, acc150, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04 and a1908=a1b01(+)" & strWhere(3) & " group by axf01, axf03 "
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = LC01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = LC02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = LC03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = LC04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, LC01||'-'||LC02||'-'||LC03||'-'||LC04"
'            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
'            StrSQLa = " select axf01 As Ax1, axf03 As Ax3, sum(axf04) As Ax4 From acc190, acc180, fagent, nation, acc151, acc150, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(3) & " group by axf01, axf03, axf04 "
'            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate," & _
'                         " a1903 as Currency, sum(ax4) as Famount, sum(ax4*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, LAWCASE, acc1b0 " & _
'                         " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = LC01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = LC02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = LC03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = LC04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, A1802, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, LC01||'-'||LC02||'-'||LC03||'-'||LC04 "
'            '2007/12/6 end
'            '2006/3/30 ADD BY SONIA 抓抵帳單資料
'            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axg04),台幣金額由sum(a1905)改為sum(axg04*a1906)
'            'StrSqlB = " select axG01 As Ax1, axG03 As Ax3 From acc190, acc180, fagent, nation, acc161, acc160, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = LC01 and substr(axG03, length(axG03) - 8, 6) = LC02 and substr(axG03, length(axG03) - 2, 1) = LC03 and substr(axG03, length(axG03) - 1, 2) = LC04 and a1908=a1b01(+)" & strWhere(5) & " group by axG01, axG03 "
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = LC01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = LC02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = LC03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = LC04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, LC01||'-'||LC02||'-'||LC03||'-'||LC04"
'            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
'            StrSqlB = " select axG01 As Ax1, axG03 As Ax3, axG04 As Ax4 From acc190, acc180, fagent, nation, acc161, acc160, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = LC01 and substr(axG03, length(axG03) - 8, 6) = LC02 and substr(axG03, length(axG03) - 2, 1) = LC03 and substr(axG03, length(axG03) - 1, 2) = LC04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(5) & " group by axG01, axG03, axG04 "
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(ax4) * (-1) as Famount, sum(ax4*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, LAWCASE, acc1b0 " & _
'                              " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = LC01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = LC02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = LC03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = LC04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, LC01||'-'||LC02||'-'||LC03||'-'||LC04 "
'            '2007/12/6 end
'            '2006/3/30 END
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation, LAWCASE " & _
'                              " where axg01 = a1601 and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = LC01 and substr(axg03, length(axg03) - 8, 6) = LC02 and substr(axg03, length(axg03) - 2, 1) = LC03 and substr(axg03, length(axg03) - 1, 2) = LC04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1601 "
'            '2005/8/2 END
'         End If
         
'專利
         If Text7 = "" Then
            If Text5 <> "" Then
               strWhere(4) = " and pa11 = '" & Text5 & "'"
            End If
'帳單ACC150
            '2014/11/24 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = "select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                            "from acc151, acc150, fagent, nation, patent where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1501 "
            strSql = "select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                            "from acc151, acc150, fagent, nation, patent, acc190 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1501=a1902(+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1501 "
            '2014/11/24 end
'抵帳資料
            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, sum(axf04) as Famount, sum(axf04 * nvl(a1g03, 0)) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc151, acc150, fagent, nation, patent, acc1g0, acc1h0, acc1i0 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2) & strWhere(4) & _
                              " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, a1512, decode(a1h02, null, a1i03, a1h02), a1505, decode(a1506, a1520, 'Y', ''), a1504, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1501 "

'帳單結匯(有匯票號)
            'Modify By Cheng 2003/11/25 acc190串到acc151時, 可能會一筆串到多筆
'            strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, acc151, acc150, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and a1908=a1b01(+)" & strWhere(3) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, PA01||'-'||PA02||'-'||PA03||'-'||PA04"
            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axf04),台幣金額由sum(a1905)改為sum(axf04*a1906)
            'StrSQLa = " Select axf01 As Ax1, axf03 As Ax3 From acc190, acc180, fagent, nation, acc151, acc150, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and a1908=a1b01(+)" & strWhere(3) & " group by axf01, axf03 "
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = pa01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = pa02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = pa03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = pa04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, PA01||'-'||PA02||'-'||PA03||'-'||PA04"
            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
            StrSQLa = " Select axf01 As Ax1, axf03 As Ax3, sum(axf04) As Ax4 From acc190, acc180, fagent, nation, acc151, acc150, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(3) & " group by axf01, axf03, axf04 "
            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate," & _
                         " a1903 as Currency, sum(ax4) as Famount, sum(ax4*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = pa01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = pa02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = pa03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = pa04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, A1802, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, PA01||'-'||PA02||'-'||PA03||'-'||PA04 "
            '2007/12/6 end
'抵帳單ACC160
            '2006/3/29 ADD BY SONIA 抓抵帳單資料
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            '2014/11/26 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, pa01||'-'||pa02||'-'||pa03||'-'||pa04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, patent " & _
                              " where axg01 = a1601 and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = pa01 and substr(axg03, length(axg03) - 8, 6) = pa02 and substr(axg03, length(axg03) - 2, 1) = pa03 and substr(axg03, length(axg03) - 1, 2) = pa04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, pa01||'-'||pa02||'-'||pa03||'-'||pa04, a1601 "
            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, pa01||'-'||pa02||'-'||pa03||'-'||pa04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, patent, ACC190 " & _
                              " where axg01 = a1601 AND A1601=A1902(+) and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = pa01 and substr(axg03, length(axg03) - 8, 6) = pa02 and substr(axg03, length(axg03) - 2, 1) = pa03 and substr(axg03, length(axg03) - 1, 2) = pa04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null), a1602, a1605, a1604, a1607, pa01||'-'||pa02||'-'||pa03||'-'||pa04, a1601 "
            '2014/11/26 end
'抵帳單結匯(有匯票號)
            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axg04),台幣金額由sum(a1905)改為sum(axg04*a1906)
            'StrSqlB = " Select axG01 As Ax1, axG03 As Ax3 From acc190, acc180, fagent, nation, acc161, acc160, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = pa01 and substr(axG03, length(axG03) - 8, 6) = pa02 and substr(axG03, length(axG03) - 2, 1) = pa03 and substr(axG03, length(axG03) - 1, 2) = pa04 and a1908=a1b01(+)" & strWhere(5) & " group by axG01, axG03 "
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = pa01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = pa02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = pa03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = pa04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, PA01||'-'||PA02||'-'||PA03||'-'||PA04"
            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
            StrSqlB = " Select axG01 As Ax1, axG03 As Ax3, axG04 As Ax4 From acc190, acc180, fagent, nation, acc161, acc160, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = pa01 and substr(axG03, length(axG03) - 8, 6) = pa02 and substr(axG03, length(axG03) - 2, 1) = pa03 and substr(axG03, length(axG03) - 1, 2) = pa04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(5) & " group by axG01, axG03, axG04 "
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(ax4) * (-1) as Famount, sum(ax4*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, patent, acc1b0 " & _
                              " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = pa01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = pa02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = pa03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = pa04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, PA01||'-'||PA02||'-'||PA03||'-'||PA04 "
            '2007/12/6 end
            '2006/3/29 END
            
'商標
            If Text5 <> "" Then
               strWhere(4) = " and tm12 = '" & Text5 & "'"
            End If
'帳單ACC150
            '2014/11/24 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                       "from acc151, acc150, fagent, nation, trademark where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1501 "
            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                       "from acc151, acc150, fagent, nation, trademark, acc190 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1501=a1902(+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1501 "
            '2014/11/24 end
'抵帳資料
            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, sum(axf04) as Famount, sum(axf04 * nvl(a1g03, 0)) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc151, acc150, fagent, nation, trademark, acc1g0, acc1h0, acc1i0 " & _
                              " where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2) & strWhere(4) & _
                              " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, a1512, decode(a1h02, null, a1i03, a1h02), a1505, decode(a1506, a1520, 'Y', ''), a1504, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1501 "
'帳單結匯(有匯票號)
            'Modify By Cheng 2003/11/25 acc190串到acc151時, 可能會一筆串到多筆
'            strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, acc151, acc150, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and a1908=a1b01(+)" & strWhere(3) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, TM01||'-'||TM02||'-'||TM03||'-'||TM04"
            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axf04),台幣金額由sum(a1905)改為sum(axf04*a1906)
            'StrSQLa = " Select axf01 as Ax1, axf03 As Ax3 From acc190, acc180, fagent, nation, acc151, acc150, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and a1908=a1b01(+)" & strWhere(3) & " group by axf01, axf03 "
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSQLa & " ) A1, acc150, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = tm01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = tm02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = tm03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = tm04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, TM01||'-'||TM02||'-'||TM03||'-'||TM04"
            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
            StrSQLa = " Select axf01 as Ax1, axf03 As Ax3, sum(axf04) As Ax4 From acc190, acc180, fagent, nation, acc151, acc150, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(3) & " group by axf01, axf03, axf04 "
            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate, a1903 as Currency, sum(ax4) as Famount, sum(ax4*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, fagent, nation, (" & StrSQLa & " ) A1, acc150, trademark, acc1b0 " & _
                         " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = tm01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = tm02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = tm03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = tm04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, A1802, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, TM01||'-'||TM02||'-'||TM03||'-'||TM04 "
            '2007/12/6 end
'抵帳單ACC160
            '2006/3/30 ADD BY SONIA 抓抵帳單資料
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            '2014/11/26 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, trademark " & _
                              " where axg01 = a1601 and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = tm01 and substr(axg03, length(axg03) - 8, 6) = tm02 and substr(axg03, length(axg03) - 2, 1) = tm03 and substr(axg03, length(axg03) - 1, 2) = tm04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1601 "
            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, trademark, ACC190 " & _
                              " where axg01 = a1601 AND A1601=A1902(+) and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = tm01 and substr(axg03, length(axg03) - 8, 6) = tm02 and substr(axg03, length(axg03) - 2, 1) = tm03 and substr(axg03, length(axg03) - 1, 2) = tm04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null), a1602, a1605, a1604, a1607, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1601 "
            '2014/11/26 end
'抵帳單結匯(有匯票號)
            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axg04),台幣金額由sum(a1905)改為sum(axg04*a1906)
            'StrSqlB = " Select axG01 as Ax1, axG03 As Ax3 From acc190, acc180, fagent, nation, acc161, acc160, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = tm01 and substr(axG03, length(axG03) - 8, 6) = tm02 and substr(axG03, length(axG03) - 2, 1) = tm03 and substr(axG03, length(axG03) - 1, 2) = tm04 and a1908=a1b01(+)" & strWhere(5) & " group by axG01, axG03 "
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSqlB & " ) A1, acc160, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = tm01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = tm02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = tm03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = tm04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, TM01||'-'||TM02||'-'||TM03||'-'||TM04"
            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
            StrSqlB = " Select axG01 as Ax1, axG03 As Ax3, axG04 As Ax4 From acc190, acc180, fagent, nation, acc161, acc160, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = tm01 and substr(axG03, length(axG03) - 8, 6) = tm02 and substr(axG03, length(axG03) - 2, 1) = tm03 and substr(axG03, length(axG03) - 1, 2) = tm04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(5) & " group by axG01, axG03, axG04 "
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(ax4) * (-1) as Famount, sum(ax4*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, fagent, nation, (" & StrSqlB & " ) A1, acc160, trademark, acc1b0 " & _
                              " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = tm01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = tm02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = tm03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = tm04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, TM01||'-'||TM02||'-'||TM03||'-'||TM04 "
            '2007/12/6 end
            '2006/3/30 END
'服務業務
            If Text5 <> "" Then
               strWhere(4) = " and sp11 = '" & Text5 & "'"
            End If
'帳單ACC150
            '2014/11/24 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                       "from acc151, acc150, fagent, nation, servicepractice where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1501 "
            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                       "from acc151, acc150, fagent, nation, servicepractice, acc190 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1501=a1902(+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1501 "
            '2014/11/24 end
'抵帳資料
            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, sum(axf04) as Famount, sum(axf04 * nvl(a1g03, 0)) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc151, acc150, fagent, nation, servicepractice, acc1g0, acc1h0, acc1i0 " & _
                              " where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2) & strWhere(4) & _
                              " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, a1512, decode(a1h02, null, a1i03, a1h02), a1505, decode(a1506, a1520, 'Y', ''), a1504, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1501 "
'帳單結匯(有匯票號)
            'Modify By Cheng 2003/11/25 acc190串到acc151時, 可能會一筆串到多筆
'            strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, acc151, acc150, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and a1908=a1b01(+)" & strWhere(3) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, SP01||'-'||SP02||'-'||SP03||'-'||SP04"
            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axf04),台幣金額由sum(a1905)改為sum(axf04*a1906)
            'StrSQLa = " select axf01 As Ax1, axf03 As Ax3 From acc190, acc180, fagent, nation, acc151, acc150, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and a1908=a1b01(+)" & strWhere(3) & " group by axf01, axf03 "
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = sp01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = sp02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = sp03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = sp04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, SP01||'-'||SP02||'-'||SP03||'-'||SP04"
            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
            StrSQLa = " select axf01 As Ax1, axf03 As Ax3, sum(axf04) As Ax4 From acc190, acc180, fagent, nation, acc151, acc150, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(3) & " group by axf01, axf03, axf04 "
            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate," & _
                         " a1903 as Currency, sum(ax4) as Famount, sum(ax4*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, servicepractice, acc1b0 " & _
                         " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = sp01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = sp02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = sp03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = sp04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, A1802, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, SP01||'-'||SP02||'-'||SP03||'-'||SP04 "
            '2007/12/6 end
'抵帳單ACC160
            '2006/3/30 ADD BY SONIA 抓抵帳單資料
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            '2014/11/26 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, sp01||'-'||sp02||'-'||sp03||'-'||sp04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, servicepractice " & _
                              " where axg01 = a1601 and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = sp01 and substr(axg03, length(axg03) - 8, 6) = sp02 and substr(axg03, length(axg03) - 2, 1) = sp03 and substr(axg03, length(axg03) - 1, 2) = sp04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, sp01||'-'||sp02||'-'||sp03||'-'||sp04, a1601 "
            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, sp01||'-'||sp02||'-'||sp03||'-'||sp04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, servicepractice, ACC190 " & _
                              " where axg01 = a1601 AND A1601=A1902(+) and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = sp01 and substr(axg03, length(axg03) - 8, 6) = sp02 and substr(axg03, length(axg03) - 2, 1) = sp03 and substr(axg03, length(axg03) - 1, 2) = sp04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null), a1602, a1605, a1604, a1607, sp01||'-'||sp02||'-'||sp03||'-'||sp04, a1601 "
            '2014/11/26 end
'抵帳單結匯(有匯票號)
            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axg04),台幣金額由sum(a1905)改為sum(axg04*a1906)
            'StrSqlB = " select axG01 As Ax1, axG03 As Ax3 From acc190, acc180, fagent, nation, acc161, acc160, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = sp01 and substr(axG03, length(axG03) - 8, 6) = sp02 and substr(axG03, length(axG03) - 2, 1) = sp03 and substr(axG03, length(axG03) - 1, 2) = sp04 and a1908=a1b01(+)" & strWhere(5) & " group by axG01, axG03 "
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = sp01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = sp02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = sp03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = sp04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, SP01||'-'||SP02||'-'||SP03||'-'||SP04"
            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
            StrSqlB = " select axG01 As Ax1, axG03 As Ax3, axG04 As Ax4 From acc190, acc180, fagent, nation, acc161, acc160, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = sp01 and substr(axG03, length(axG03) - 8, 6) = sp02 and substr(axG03, length(axG03) - 2, 1) = sp03 and substr(axG03, length(axG03) - 1, 2) = sp04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(5) & " group by axG01, axG03, axG04 "
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(ax4) * (-1) as Famount, sum(ax4*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, servicepractice, acc1b0 " & _
                              " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = sp01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = sp02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = sp03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = sp04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, SP01||'-'||SP02||'-'||SP03||'-'||SP04 "
            '2007/12/6 end
            '2006/3/30 END
            
'法務2005/8/2 ADD BY SONIA 加入LAWCASE
'帳單ACC150
            '2014/11/24 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                       "from acc151, acc150, fagent, nation, LAWCASE where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1501 "
             'Modified by Lydia 2018/02/13 拿掉strWhere(4) ; 有基本檔才抓
'            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
'                                       "from acc151, acc150, fagent, nation, LAWCASE, acc190 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1501=a1902(+) and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1501 "
            '2014/11/24 end
            If Trim(Text7 & Text1 & Text3 & Text8) <> "" Then
                strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161" & _
                                           " from acc151, acc150, fagent, nation, LAWCASE, acc190 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1501=a1902(+) and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04" & strWhere(2) & _
                                           " and lc01 is not null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1501 "
            End If
            'end 2018/02/13
'抵帳資料
            'Modified by Lydia 2018/02/13 拿掉strWhere(4) ; 有基本檔才抓
            'strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, sum(axf04) as Famount, sum(axf04 * nvl(a1g03, 0)) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc151, acc150, fagent, nation, LAWCASE, acc1g0, acc1h0, acc1i0 " & _
                              " where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04 and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2) & strWhere(4) & _
                              " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, a1512, decode(a1h02, null, a1i03, a1h02), a1505, decode(a1506, a1520, 'Y', ''), a1504, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1501 "
            If Trim(Text7 & Text1 & Text3 & Text8) <> "" Then
                strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, sum(axf04) as Famount, sum(axf04 * nvl(a1g03, 0)) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc151, acc150, fagent, nation, LAWCASE, acc1g0, acc1h0, acc1i0 " & _
                                  " where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04 and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2) & _
                                  " and lc01 is not null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, a1512, decode(a1h02, null, a1i03, a1h02), a1505, decode(a1506, a1520, 'Y', ''), a1504, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1501 "
            End If
            'end 2018/02/13
'帳單結匯(有匯票號)
            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axf04),台幣金額由sum(a1905)改為sum(axf04*a1906)
            'StrSQLa = " select axf01 As Ax1, axf03 As Ax3 From acc190, acc180, fagent, nation, acc151, acc150, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04 and a1908=a1b01(+)" & strWhere(3) & " group by axf01, axf03 "
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = LC01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = LC02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = LC03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = LC04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, LC01||'-'||LC02||'-'||LC03||'-'||LC04"
            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
            StrSQLa = " select axf01 As Ax1, axf03 As Ax3, sum(axf04) As Ax4 From acc190, acc180, fagent, nation, acc151, acc150, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(3) & " group by axf01, axf03, axf04 "
            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate," & _
                         " a1903 as Currency, sum(ax4) as Famount, sum(ax4*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, LAWCASE, acc1b0 " & _
                         " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = LC01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = LC02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = LC03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = LC04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, A1802, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, LC01||'-'||LC02||'-'||LC03||'-'||LC04 "
            '2007/12/6 end
'抵帳單ACC160
            '2006/3/30 ADD BY SONIA 抓抵帳單資料
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            '2014/11/26 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, LAWCASE " & _
                              " where axg01 = a1601 and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = LC01 and substr(axg03, length(axg03) - 8, 6) = LC02 and substr(axg03, length(axg03) - 2, 1) = LC03 and substr(axg03, length(axg03) - 1, 2) = LC04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1601 "
            'Modified by Lydia 2018/02/13 拿掉strWhere(4) ; 有基本檔才抓
'            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, LAWCASE, ACC190 " & _
'                              " where axg01 = a1601 AND A1601=A1902(+) and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = LC01 and substr(axg03, length(axg03) - 8, 6) = LC02 and substr(axg03, length(axg03) - 2, 1) = LC03 and substr(axg03, length(axg03) - 1, 2) = LC04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null), a1602, a1605, a1604, a1607, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1601 "
            '2014/11/26 end
            If Trim(Text7 & Text1 & Text3 & Text8) <> "" Then
                strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, LAWCASE, ACC190 " & _
                                  " where axg01 = a1601 AND A1601=A1902(+) and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = LC01 and substr(axg03, length(axg03) - 8, 6) = LC02 and substr(axg03, length(axg03) - 2, 1) = LC03 and substr(axg03, length(axg03) - 1, 2) = LC04" & strWhere(5) & _
                                  " and lc01 is not null group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null), a1602, a1605, a1604, a1607, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1601 "
            End If
            'end 2018/02/13
'抵帳單結匯(有匯票號)
            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axg04),台幣金額由sum(a1905)改為sum(axg04*a1906)
            'StrSqlB = " select axG01 As Ax1, axG03 As Ax3 From acc190, acc180, fagent, nation, acc161, acc160, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = LC01 and substr(axG03, length(axG03) - 8, 6) = LC02 and substr(axG03, length(axG03) - 2, 1) = LC03 and substr(axG03, length(axG03) - 1, 2) = LC04 and a1908=a1b01(+)" & strWhere(5) & " group by axG01, axG03 "
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = LC01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = LC02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = LC03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = LC04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, LC01||'-'||LC02||'-'||LC03||'-'||LC04"
            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
            StrSqlB = " select axG01 As Ax1, axG03 As Ax3, axG04 As Ax4 From acc190, acc180, fagent, nation, acc161, acc160, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = LC01 and substr(axG03, length(axG03) - 8, 6) = LC02 and substr(axG03, length(axG03) - 2, 1) = LC03 and substr(axG03, length(axG03) - 1, 2) = LC04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(5) & " group by axG01, axG03, axG04 "
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(ax4) * (-1) as Famount, sum(ax4*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, LAWCASE, acc1b0 " & _
                              " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = LC01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = LC02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = LC03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = LC04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, LC01||'-'||LC02||'-'||LC03||'-'||LC04 "
            '2007/12/6 end
            '2006/3/30 END
            '2005/8/2 END
         End If
      Case "4" 'CF未付
'         If Text7 = "" Then
'            If Text5 <> "" Then
'               strWhere(4) = " and pa11 = '" & Text5 & "'"
'            End If
'            '2010/10/26 modify by sonia 作廢或已抵帳都不要
'            strSql = "select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                         "from acc151, acc150, fagent, nation, patent where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1501 "
'            '2006/4/3 ADD BY SONIA
'            '2010/10/25 modify by sonia 加判斷抵帳日期a1607
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            'Modify By Sindy 2012/12/6 + AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null +, ACC190
'            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation, PATENT, ACC190 " & _
'                              " where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = PA01 and substr(axg03, length(axg03) - 8, 6) = PA02 and substr(axg03, length(axg03) - 2, 1) = PA03 and substr(axg03, length(axg03) - 1, 2) = PA04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1601 "
'            '2006/4/3 END
'            If Text5 <> "" Then
'               strWhere(4) = " and tm12 = '" & Text5 & "'"
'            End If
'            '2010/10/26 modify by sonia 作廢或已抵帳都不要
'            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                            "from acc151, acc150, fagent, nation, trademark where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1501 "
'            '2006/4/3 ADD BY SONIA
'            '2010/10/25 modify by sonia 加判斷抵帳日期a1607
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            'Modify By Sindy 2012/12/6 + AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null +, ACC190
'            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation, TRADEMARK, ACC190 " & _
'                              " where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = TM01 and substr(axg03, length(axg03) - 8, 6) = TM02 and substr(axg03, length(axg03) - 2, 1) = TM03 and substr(axg03, length(axg03) - 1, 2) = TM04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1601 "
'            '2006/4/3 END
'            If Text5 <> "" Then
'               strWhere(4) = " and sp11 = '" & Text5 & "'"
'            End If
'            '2010/10/26 modify by sonia 作廢或已抵帳都不要
'            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                            "from acc151, acc150, fagent, nation, servicepractice where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1501 "
'            '2006/4/3 ADD BY SONIA
'            '2010/10/25 modify by sonia 加判斷抵帳日期a1607
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            'Modify By Sindy 2012/12/6 + AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null +, ACC190
'            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation, SERVICEPRACTICE, ACC190 where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = SP01 and substr(axg03, length(axg03) - 8, 6) = SP02 and substr(axg03, length(axg03) - 2, 1) = SP03 and substr(axg03, length(axg03) - 1, 2) = SP04" & strWhere(5) & strWhere(4) & _
'            " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1601 "
'            '2006/4/3 END
'            '2005/8/2 ADD BY SONIA 加入LAWCASE
'            '2010/10/26 modify by sonia 作廢或已抵帳都不要
'            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                            "from acc151, acc150, fagent, nation, LAWCASE where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1501 "
'            '2005/8/2 END
'            '2006/4/3 ADD BY SONIA
'            '2010/10/25 modify by sonia 加判斷抵帳日期a1607
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            'Modify By Sindy 2012/12/6 + AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null +, ACC190
'            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation, LAWCASE, ACC190 where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = LC01 and substr(axg03, length(axg03) - 8, 6) = LC02 and substr(axg03, length(axg03) - 2, 1) = LC03 and substr(axg03, length(axg03) - 1, 2) = LC04" & strWhere(5) & strWhere(4) & _
'            " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1601 "
'            '2006/4/3 END
'         End If
         
'專利
         If Text7 = "" Then
            If Text5 <> "" Then
               strWhere(4) = " and pa11 = '" & Text5 & "'"
            End If
'帳單ACC150
            '2010/10/26 modify by sonia 作廢或已抵帳都不要
            '2014/11/24 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = "select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                         "from acc151, acc150, fagent, nation, patent where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1501 "
            strSql = "select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                         "from acc151, acc150, fagent, nation, patent, acc190 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1501=a1902(+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1501 "
            '2014/11/24 end
'抵帳單ACC160
            '2006/4/3 ADD BY SONIA
            '2010/10/25 modify by sonia 加判斷抵帳日期a1607
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            'Modify By Sindy 2012/12/6 + AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null +, ACC190
            '2014/11/26 modify by sonia 加>付款中的符號(有acc190但無匯票號碼),故取消A1901 IS NULL條件
            'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, PATENT, ACC190 " & _
                              " where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = PA01 and substr(axg03, length(axg03) - 8, 6) = PA02 and substr(axg03, length(axg03) - 2, 1) = PA03 and substr(axg03, length(axg03) - 1, 2) = PA04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1601 "
            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, PATENT, ACC190 " & _
                              " where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = PA01 and substr(axg03, length(axg03) - 8, 6) = PA02 and substr(axg03, length(axg03) - 2, 1) = PA03 and substr(axg03, length(axg03) - 1, 2) = PA04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null), a1602, a1605, a1604, a1607, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1601 "
            '2014/11/26 end
            '2006/4/3 END
'商標
            If Text5 <> "" Then
               strWhere(4) = " and tm12 = '" & Text5 & "'"
            End If
'帳單ACC150
            '2010/10/26 modify by sonia 作廢或已抵帳都不要
            '2014/11/24 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                            "from acc151, acc150, fagent, nation, trademark where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1501 "
            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                            "from acc151, acc150, fagent, nation, trademark, acc190 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1501=a1902(+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1501 "
            '2014/11/24 end
'抵帳單ACC160
            '2006/4/3 ADD BY SONIA
            '2010/10/25 modify by sonia 加判斷抵帳日期a1607
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            'Modify By Sindy 2012/12/6 + AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null +, ACC190
            '2014/11/26 modify by sonia 加>付款中的符號(有acc190但無匯票號碼),故取消A1901 IS NULL條件
            'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, TRADEMARK, ACC190 " & _
                              " where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = TM01 and substr(axg03, length(axg03) - 8, 6) = TM02 and substr(axg03, length(axg03) - 2, 1) = TM03 and substr(axg03, length(axg03) - 1, 2) = TM04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1601 "
            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, TRADEMARK, ACC190 " & _
                              " where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = TM01 and substr(axg03, length(axg03) - 8, 6) = TM02 and substr(axg03, length(axg03) - 2, 1) = TM03 and substr(axg03, length(axg03) - 1, 2) = TM04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null), a1602, a1605, a1604, a1607, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1601 "
            '2014/11/26 end
            '2006/4/3 END
'服務業務
            If Text5 <> "" Then
               strWhere(4) = " and sp11 = '" & Text5 & "'"
            End If
'帳單ACC150
            '2010/10/26 modify by sonia 作廢或已抵帳都不要
            '2014/11/24 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                            "from acc151, acc150, fagent, nation, servicepractice where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1501 "
            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                            "from acc151, acc150, fagent, nation, servicepractice, acc190 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1501=a1902(+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1501 "
            '2014/11/24 end
'抵帳單ACC160
            '2006/4/3 ADD BY SONIA
            '2010/10/25 modify by sonia 加判斷抵帳日期a1607
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            'Modify By Sindy 2012/12/6 + AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null +, ACC190
            '2014/11/26 modify by sonia 加>付款中的符號(有acc190但無匯票號碼),故取消A1901 IS NULL條件
            'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, SERVICEPRACTICE, ACC190 where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = SP01 and substr(axg03, length(axg03) - 8, 6) = SP02 and substr(axg03, length(axg03) - 2, 1) = SP03 and substr(axg03, length(axg03) - 1, 2) = SP04" & strWhere(5) & strWhere(4) & _
            " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1601 "
            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, SERVICEPRACTICE, ACC190 where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = SP01 and substr(axg03, length(axg03) - 8, 6) = SP02 and substr(axg03, length(axg03) - 2, 1) = SP03 and substr(axg03, length(axg03) - 1, 2) = SP04" & strWhere(5) & strWhere(4) & _
                              " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null), a1602, a1605, a1604, a1607, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1601 "
            '2014/11/26 end
            '2006/4/3 END
'法務2005/8/2 ADD BY SONIA 加入LAWCASE
'帳單ACC150
            '2010/10/26 modify by sonia 作廢或已抵帳都不要
            '2014/11/24 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                            "from acc151, acc150, fagent, nation, LAWCASE where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1501 "
             'Modified by Lydia 2018/02/13 拿掉strWhere(4) ; 有基本檔才抓
            'strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                            "from acc151, acc150, fagent, nation, LAWCASE, acc190 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1501=a1902(+) and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1501 "
            '2014/11/24 end
            '2005/8/2 END
            If Trim(Text7 & Text1 & Text3 & Text8) <> "" Then
                strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                "from acc151, acc150, fagent, nation, LAWCASE, acc190 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1501=a1902(+) and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04" & strWhere(2) & _
                                " and lc01 is not null and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1501 "
            End If
             'end 2018/02/13
'抵帳單ACC160
            '2006/4/3 ADD BY SONIA
            '2010/10/25 modify by sonia 加判斷抵帳日期a1607
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            'Modify By Sindy 2012/12/6 + AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null +, ACC190
            '2014/11/26 modify by sonia 加>付款中的符號(有acc190但無匯票號碼),故取消A1901 IS NULL條件
            'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, LAWCASE, ACC190 where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = LC01 and substr(axg03, length(axg03) - 8, 6) = LC02 and substr(axg03, length(axg03) - 2, 1) = LC03 and substr(axg03, length(axg03) - 1, 2) = LC04" & strWhere(5) & strWhere(4) & _
            " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1601 "
            'Modified by Lydia 2018/02/13 拿掉strWhere(4) ; 有基本檔才抓
'            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, LAWCASE, ACC190 where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = LC01 and substr(axg03, length(axg03) - 8, 6) = LC02 and substr(axg03, length(axg03) - 2, 1) = LC03 and substr(axg03, length(axg03) - 1, 2) = LC04" & strWhere(5) & strWhere(4) & _
            " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null), a1602, a1605, a1604, a1607, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1601 "
            '2014/11/26 end
            '2006/4/3 END
           If Trim(Text7 & Text1 & Text3 & Text8) <> "" Then
                strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161" & _
                                         " from acc161, acc160, fagent, nation, LAWCASE, ACC190 where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = LC01 and substr(axg03, length(axg03) - 8, 6) = LC02 and substr(axg03, length(axg03) - 2, 1) = LC03 and substr(axg03, length(axg03) - 1, 2) = LC04" & strWhere(5) & _
                                         " and lc01 is not null group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null), a1602, a1605, a1604, a1607, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1601 "
           End If
           'end 2018/02/13
         End If
      Case "6" '未收未付
'         If Text5 <> "" Then
'            strWhere(4) = " and pa11 = '" & Text5 & "'"
'         End If
'         '2007/12/10 modify by sonia X09607651分次收款,婧瑄說台幣金額扣除已收金額,外幣改為台幣金額扣除已收金額/請款匯率
'         'strSQL = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27 from acc1k0, fagent, nation, patent, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         'Modify By Sindy 2012/6/25 X09607651分次收款未收金額部分改回原程式寫法,以X10003936測試
''         strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
''                         "from acc1k0, fagent, nation, patent, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                         "from acc1k0, fagent, nation, patent, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         '2012/6/25 End
'         If Text7 = "" Then
'            '2010/10/26 modify by sonia 作廢或已抵帳都不要
'            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                        "from acc151, acc150, fagent, nation, patent where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1501 "
'            '2006/4/3 ADD BY SONIA
'            '2010/10/25 modify by sonia 加判斷抵帳日期a1607
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            'Modify By Sindy 2012/12/6 + AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null +, ACC190
'            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation, PATENT, ACC190 where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = PA01 and substr(axg03, length(axg03) - 8, 6) = PA02 and substr(axg03, length(axg03) - 2, 1) = PA03 and substr(axg03, length(axg03) - 1, 2) = PA04" & strWhere(5) & strWhere(4) & _
'            " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1601 "
'            '2006/4/3 END
'         End If
'         'Add By Sindy 2012/8/14 +部分收款時要同時帶出收款資料
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort , a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount from acc0y0, fagent, nation, acc0z0, acc1k0, patent where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04" & strWhere(1) & strWhere(4) & _
'                           " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'         '2012/8/14 End
'         If Text5 <> "" Then
'            strWhere(4) = " and tm12 = '" & Text5 & "'"
'         End If
'         '2007/12/10 modify by sonia X09607651分次收款,婧瑄說台幣金額扣除已收金額,外幣改為台幣金額扣除已收金額/請款匯率
'         'strSQL = strSQL & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27 from acc1k0, fagent, nation, trademark, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         'Modify By Sindy 2012/6/25 X09607651分次收款未收金額部分改回原程式寫法,以X10003936測試
''         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
''                                    "from acc1k0, fagent, nation, trademark, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                                    "from acc1k0, fagent, nation, trademark, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         '2012/6/25 End
'         If Text7 = "" Then
'            '2010/10/26 modify by sonia 作廢或已抵帳都不要
'            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                        "from acc151, acc150, fagent, nation, trademark where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1501 "
'            '2006/4/3 ADD BY SONIA
'            '2010/10/25 modify by sonia 加判斷抵帳日期a1607
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            'Modify By Sindy 2012/12/6 + AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null +, ACC190
'            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation, TRADEMARK, ACC190 where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = TM01 and substr(axg03, length(axg03) - 8, 6) = TM02 and substr(axg03, length(axg03) - 2, 1) = TM03 and substr(axg03, length(axg03) - 1, 2) = TM04" & strWhere(5) & strWhere(4) & _
'            " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1601 "
'            '2006/4/3 END
'         End If
'         'Add By Sindy 2012/8/14 +部分收款時要同時帶出收款資料
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                           "from acc0y0, fagent, nation, acc0z0, acc1k0, trademark where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04" & strWhere(1) & strWhere(4) & " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'         '2012/8/14 End
'         If Text5 <> "" Then
'            strWhere(4) = " and sp11 = '" & Text5 & "'"
'         End If
'         '2007/12/10 modify by sonia X09607651分次收款,婧瑄說台幣金額扣除已收金額,外幣改為台幣金額扣除已收金額/請款匯率
'         'strSQL = strSQL & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27 from acc1k0, fagent, nation, servicepractice, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         'Modify By Sindy 2012/6/25 X09607651分次收款未收金額部分改回原程式寫法,以X10003936測試
''         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
''                                    "from acc1k0, fagent, nation, servicepractice, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                                    "from acc1k0, fagent, nation, servicepractice, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         '2012/6/25 End
'         If Text7 = "" Then
'            '2010/10/26 modify by sonia 作廢或已抵帳都不要
'            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                                    "from acc151, acc150, fagent, nation, servicepractice where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1501 "
'            '2006/4/3 ADD BY SONIA
'            '2010/10/25 modify by sonia 加判斷抵帳日期a1607
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            'Modify By Sindy 2012/12/6 + AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null +, ACC190
'            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation, SERVICEPRACTICE, ACC190 where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = SP01 and substr(axg03, length(axg03) - 8, 6) = SP02 and substr(axg03, length(axg03) - 2, 1) = SP03 and substr(axg03, length(axg03) - 1, 2) = SP04" & _
'                              strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1601 "
'            '2006/4/3 END
'         End If
'         'Add By Sindy 2012/8/14 +部分收款時要同時帶出收款資料
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                                    "from acc0y0, fagent, nation, acc0z0, acc1k0, servicepractice where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04" & strWhere(1) & strWhere(4) & " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'         '2012/8/14 End
'         '2005/8/3 ADD BY SONIA 加入LAWCASE
'         '2007/12/10 modify by sonia X09607651分次收款,婧瑄說台幣金額扣除已收金額,外幣改為台幣金額扣除已收金額/請款匯率
'         'strSQL = strSQL & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27 from acc1k0, fagent, nation, LAWCASE, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         'Modify By Sindy 2012/6/25 X09607651分次收款未收金額部分改回原程式寫法,以X10003936測試
''         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
''                                    "from acc1k0, fagent, nation, LAWCASE, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                                    "from acc1k0, fagent, nation, LAWCASE, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         '2012/6/25 End
'         If Text7 = "" Then
'            '2010/10/26 modify by sonia 作廢或已抵帳都不要
'            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                                    "from acc151, acc150, fagent, nation, LAWCASE where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1501 "
'            '2006/4/3 ADD BY SONIA
'            '2010/10/25 modify by sonia 加判斷抵帳日期a1607
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            'Modify By Sindy 2012/12/6 + AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null +, ACC190
'            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation, LAWCASE, ACC190 where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = LC01 and substr(axg03, length(axg03) - 8, 6) = LC02 and substr(axg03, length(axg03) - 2, 1) = LC03 and substr(axg03, length(axg03) - 1, 2) = LC04" & strWhere(5) & strWhere(4) & _
'                     " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1601 "
'            '2006/4/3 END
'         End If
'         '2005/8/3 END
'         'Add By Sindy 2012/8/14 +部分收款時要同時帶出收款資料
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                           "from acc0y0, fagent, nation, acc0z0, acc1k0, LAWCASE where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04" & strWhere(1) & strWhere(4) & " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'         '2012/8/14 End

'專利
         If Text5 <> "" Then
            strWhere(4) = " and pa11 = '" & Text5 & "'"
         End If
'未收
         '2007/12/10 modify by sonia X09607651分次收款,婧瑄說台幣金額扣除已收金額,外幣改為台幣金額扣除已收金額/請款匯率
         'strSQL = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27 from acc1k0, fagent, nation, patent, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
         'Modify By Sindy 2012/6/25 X09607651分次收款未收金額部分改回原程式寫法,以X10003936測試
'         strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                         "from acc1k0, fagent, nation, patent, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
         strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                         "from acc1k0, fagent, nation, patent, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
         '2012/6/25 End
'Add By Sindy 2012/8/14 +部分收款時要同時帶出收款資料
         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort , a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc0y0, fagent, nation, acc0z0, acc1k0, patent where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04" & strWhere(1) & strWhere(4) & _
                           " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'2012/8/14 End
         If Text7 = "" Then
'帳單ACC150
            '2010/10/26 modify by sonia 作廢或已抵帳都不要
            '2014/11/24 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                        "from acc151, acc150, fagent, nation, patent where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1501 "
            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                        "from acc151, acc150, fagent, nation, patent, acc190 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1501=a1902(+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1501 "
            '2014/11/24 end
'抵帳單ACC160
            '2006/4/3 ADD BY SONIA
            '2010/10/25 modify by sonia 加判斷抵帳日期a1607
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            'Modify By Sindy 2012/12/6 + AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null +, ACC190
            '2014/11/26 modify by sonia 加>付款中的符號(有acc190但無匯票號碼),故取消A1901 IS NULL條件
            'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, PATENT, ACC190 where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = PA01 and substr(axg03, length(axg03) - 8, 6) = PA02 and substr(axg03, length(axg03) - 2, 1) = PA03 and substr(axg03, length(axg03) - 1, 2) = PA04" & strWhere(5) & strWhere(4) & _
            " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1601 "
            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, PATENT, ACC190 where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = PA01 and substr(axg03, length(axg03) - 8, 6) = PA02 and substr(axg03, length(axg03) - 2, 1) = PA03 and substr(axg03, length(axg03) - 1, 2) = PA04" & strWhere(5) & strWhere(4) & _
            " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null), a1602, a1605, a1604, a1607, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1601 "
            '2014/11/26 end
            '2006/4/3 END
         End If
'商標
         If Text5 <> "" Then
            strWhere(4) = " and tm12 = '" & Text5 & "'"
         End If
'未收
         '2007/12/10 modify by sonia X09607651分次收款,婧瑄說台幣金額扣除已收金額,外幣改為台幣金額扣除已收金額/請款匯率
         'strSQL = strSQL & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27 from acc1k0, fagent, nation, trademark, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
         'Modify By Sindy 2012/6/25 X09607651分次收款未收金額部分改回原程式寫法,以X10003936測試
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                                    "from acc1k0, fagent, nation, trademark, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                    "from acc1k0, fagent, nation, trademark, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
         '2012/6/25 End
'Add By Sindy 2012/8/14 +部分收款時要同時帶出收款資料
         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                           "from acc0y0, fagent, nation, acc0z0, acc1k0, trademark where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04" & strWhere(1) & strWhere(4) & " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'2012/8/14 End
         If Text7 = "" Then
'帳單ACC150
            '2010/10/26 modify by sonia 作廢或已抵帳都不要
            '2014/11/24 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                        "from acc151, acc150, fagent, nation, trademark where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1501 "
            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                        "from acc151, acc150, fagent, nation, trademark, acc190 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1501=a1902(+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1501 "
            '2014/11/24 end
'抵帳單ACC160
            '2006/4/3 ADD BY SONIA
            '2010/10/25 modify by sonia 加判斷抵帳日期a1607
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            'Modify By Sindy 2012/12/6 + AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null +, ACC190
            '2014/11/26 modify by sonia 加>付款中的符號(有acc190但無匯票號碼),故取消A1901 IS NULL條件
            'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, TRADEMARK, ACC190 where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = TM01 and substr(axg03, length(axg03) - 8, 6) = TM02 and substr(axg03, length(axg03) - 2, 1) = TM03 and substr(axg03, length(axg03) - 1, 2) = TM04" & strWhere(5) & strWhere(4) & _
            " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1601 "
            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, TRADEMARK, ACC190 where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = TM01 and substr(axg03, length(axg03) - 8, 6) = TM02 and substr(axg03, length(axg03) - 2, 1) = TM03 and substr(axg03, length(axg03) - 1, 2) = TM04" & strWhere(5) & strWhere(4) & _
            " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null), a1602, a1605, a1604, a1607, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1601 "
            '2014/11/26 end
            '2006/4/3 END
         End If
'服務業務
         If Text5 <> "" Then
            strWhere(4) = " and sp11 = '" & Text5 & "'"
         End If
'未收
         '2007/12/10 modify by sonia X09607651分次收款,婧瑄說台幣金額扣除已收金額,外幣改為台幣金額扣除已收金額/請款匯率
         'strSQL = strSQL & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27 from acc1k0, fagent, nation, servicepractice, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
         'Modify By Sindy 2012/6/25 X09607651分次收款未收金額部分改回原程式寫法,以X10003936測試
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                                    "from acc1k0, fagent, nation, servicepractice, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                    "from acc1k0, fagent, nation, servicepractice, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
         '2012/6/25 End
'Add By Sindy 2012/8/14 +部分收款時要同時帶出收款資料
         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                    "from acc0y0, fagent, nation, acc0z0, acc1k0, servicepractice where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04" & strWhere(1) & strWhere(4) & " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'2012/8/14 End
         If Text7 = "" Then
'帳單ACC150
            '2010/10/26 modify by sonia 作廢或已抵帳都不要
            '2014/11/24 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                    "from acc151, acc150, fagent, nation, servicepractice where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1501 "
            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                    "from acc151, acc150, fagent, nation, servicepractice, acc190 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1501=a1902(+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1501 "
            '2014/11/24 end
'抵帳單ACC160
            '2006/4/3 ADD BY SONIA
            '2010/10/25 modify by sonia 加判斷抵帳日期a1607
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            'Modify By Sindy 2012/12/6 + AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null +, ACC190
            '2014/11/26 modify by sonia 加>付款中的符號(有acc190但無匯票號碼),故取消A1901 IS NULL條件
            'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, SERVICEPRACTICE, ACC190 where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = SP01 and substr(axg03, length(axg03) - 8, 6) = SP02 and substr(axg03, length(axg03) - 2, 1) = SP03 and substr(axg03, length(axg03) - 1, 2) = SP04" & _
                              strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1601 "
            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, SERVICEPRACTICE, ACC190 where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = SP01 and substr(axg03, length(axg03) - 8, 6) = SP02 and substr(axg03, length(axg03) - 2, 1) = SP03 and substr(axg03, length(axg03) - 1, 2) = SP04" & _
                              strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null), a1602, a1605, a1604, a1607, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1601 "
            '2014/11/26 end
            '2006/4/3 END
         End If
'法務2005/8/3 ADD BY SONIA 加入LAWCASE
'未收
         '2007/12/10 modify by sonia X09607651分次收款,婧瑄說台幣金額扣除已收金額,外幣改為台幣金額扣除已收金額/請款匯率
         'strSQL = strSQL & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27 from acc1k0, fagent, nation, LAWCASE, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
         'Modify By Sindy 2012/6/25 X09607651分次收款未收金額部分改回原程式寫法,以X10003936測試
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Famount, decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1))) USDamount " & _
'                                    "from acc1k0, fagent, nation, LAWCASE, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
        'Modified by Lydia 2018/02/13 拿掉strWhere(4) ; 有基本檔才抓
        'strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                   "from acc1k0, fagent, nation, LAWCASE, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
         '2012/6/25 End
         If Trim(Text7 & Text1 & Text3 & Text8) <> "" Then
            strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161" & _
                                      " from acc1k0, fagent, nation, LAWCASE, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04 and a1k01 = a1403 (+)" & strWhere(0) & _
                                      " and lc01 is not null and (a1k29 is null or a1k29 = '')"
         End If
         'end 2018/02/13
'Add By Sindy 2012/8/14 +部分收款時要同時帶出收款資料
         'Modified by Lydia 2018/02/13 拿掉strWhere(4) ; 有基本檔才抓
        'strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                          "from acc0y0, fagent, nation, acc0z0, acc1k0, LAWCASE where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04" & strWhere(1) & strWhere(4) & " and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'2012/8/14 End
        If Trim(Text7 & Text1 & Text3 & Text8) <> "" Then
            strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161" & _
                              " from acc0y0, fagent, nation, acc0z0, acc1k0, LAWCASE where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04" & strWhere(1) & _
                              " and lc01 is not null and (a1k29 is null or a1k29 = '') and a1k30>0 group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
        End If
        'end 2018/02/13
         If Text7 = "" Then
'帳單ACC150
            '2010/10/26 modify by sonia 作廢或已抵帳都不要
            '2014/11/24 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                    "from acc151, acc150, fagent, nation, LAWCASE where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1501 "
            'Modified by Lydia 2018/02/13 拿掉strWhere(4) ; 有基本檔才抓
            'strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                    "from acc151, acc150, fagent, nation, LAWCASE, acc190 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1501=a1902(+) and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1501 "
            '2014/11/24 end
            If Trim(Text7 & Text1 & Text3 & Text8) <> "" Then
                strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161" & _
                                        " from acc151, acc150, fagent, nation, LAWCASE, acc190 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1501=a1902(+) and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04" & strWhere(2) & _
                                        " and lc01 is not null  and (a1520 = 0 or a1520 is null) and a1507 is null and a1512 is null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1501 "
            End If
            'end 2018/02/13
'抵帳單ACC160
            '2006/4/3 ADD BY SONIA
            '2010/10/25 modify by sonia 加判斷抵帳日期a1607
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            'Modify By Sindy 2012/12/6 + AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null +, ACC190
            '2014/11/26 modify by sonia 加>付款中的符號(有acc190但無匯票號碼),故取消A1901 IS NULL條件
            'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, LAWCASE, ACC190 where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) AND A1901 IS NULL and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = LC01 and substr(axg03, length(axg03) - 8, 6) = LC02 and substr(axg03, length(axg03) - 2, 1) = LC03 and substr(axg03, length(axg03) - 1, 2) = LC04" & strWhere(5) & strWhere(4) & _
                     " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1601 "
            'Modified by Lydia 2018/02/13 拿掉strWhere(4) ; 有基本檔才抓
            'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, LAWCASE, ACC190 where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = LC01 and substr(axg03, length(axg03) - 8, 6) = LC02 and substr(axg03, length(axg03) - 2, 1) = LC03 and substr(axg03, length(axg03) - 1, 2) = LC04" & strWhere(5) & strWhere(4) & _
                     " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null), a1602, a1605, a1604, a1607, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1601 "
            '2014/11/26 end
            '2006/4/3 END
            If Trim(Text7 & Text1 & Text3 & Text8) <> "" Then
                strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161" & _
                                         " from acc161, acc160, fagent, nation, LAWCASE, ACC190 where A1607 IS NULL AND axg01 = a1601 AND A1601=A1902(+) and a1908 is null and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = LC01 and substr(axg03, length(axg03) - 8, 6) = LC02 and substr(axg03, length(axg03) - 2, 1) = LC03 and substr(axg03, length(axg03) - 1, 2) = LC04" & strWhere(5) & _
                                         " and lc01 is not null  group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null), a1602, a1605, a1604, a1607, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1601 "
            End If
            'end 2018/02/13
         End If
         '2005/8/3 END
      Case Else '往來
'         If Text5 <> "" Then
'            strWhere(4) = " and pa11 = '" & Text5 & "'"
'         End If
'         strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,(a1k08 - nvl(a1k06, 0))) USDamount from acc1k0, fagent, nation, patent, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4)
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                                    "from acc0y0, fagent, nation, acc0z0, acc1k0, patent where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04" & strWhere(1) & strWhere(4) & " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'         If Text7 = "" Then   '無請款單號時
'            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                                       "from acc151, acc150, fagent, nation, patent where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1501 "
'            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, sum(axf04) as Famount, sum(axf04 * nvl(a1g03, 0)) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc151, acc150, fagent, nation, patent, acc1g0, acc1h0, acc1i0 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2) & strWhere(4) & _
'                              " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, a1512, decode(a1h02, null, a1i03, a1h02), a1505, decode(a1506, a1520, 'Y', ''), a1504, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1501 "
'            '2006/11/20 MODIFY BY SONIA acc190串到acc151時, 可能會一筆串到多筆
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27  from acc190, acc180, fagent, nation, acc151, acc150, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and a1908=a1b01(+)" & strWhere(3) & strWhere(4) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, PA01||'-'||PA02||'-'||PA03||'-'||PA04"
'            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axf04),台幣金額由sum(a1905)改為sum(axf04*a1906)
'            'StrSQLa = " Select axf01 As Ax1, axf03 As Ax3 From acc190, acc180, fagent, nation, acc151, acc150, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and a1908=a1b01(+)" & strWhere(3) & " group by axf01, axf03 "
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = pa01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = pa02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = pa03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = pa04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, PA01||'-'||PA02||'-'||PA03||'-'||PA04"
'            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
'            StrSQLa = " Select axf01 As Ax1, axf03 As Ax3, sum(axf04) As Ax4 From acc190, acc180, fagent, nation, acc151, acc150, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(3) & " group by axf01, axf03, axf04 "
'            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate," & _
'                         " a1903 as Currency, sum(ax4) as Famount, sum(ax4*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = pa01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = pa02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = pa03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = pa04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, A1802, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, PA01||'-'||PA02||'-'||PA03||'-'||PA04 "
'            '2007/12/6 end
'            '2006/11/20 END
'            '2006/4/3 ADD BY SONIA
'            '2006/11/20 MODIFY BY SONIA acc190串到acc151時, 可能會一筆串到多筆
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27  from acc190, acc180, fagent, nation, acc161, acc160, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = pa01 and substr(axG03, length(axG03) - 8, 6) = pa02 and substr(axG03, length(axG03) - 2, 1) = pa03 and substr(axG03, length(axG03) - 1, 2) = pa04 and a1908=a1b01(+)" & strWhere(5) & strWhere(4) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, PA01||'-'||PA02||'-'||PA03||'-'||PA04"
'            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axg04),台幣金額由sum(a1905)改為sum(axg04*a1906)
'            'StrSqlB = " Select axG01 As Ax1, axG03 As Ax3 From acc190, acc180, fagent, nation, acc161, acc160, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = pa01 and substr(axG03, length(axG03) - 8, 6) = pa02 and substr(axG03, length(axG03) - 2, 1) = pa03 and substr(axG03, length(axG03) - 1, 2) = pa04 and a1908=a1b01(+)" & strWhere(5) & " group by axG01, axG03 "
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = pa01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = pa02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = pa03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = pa04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, PA01||'-'||PA02||'-'||PA03||'-'||PA04"
'            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
'            StrSqlB = " Select axG01 As Ax1, axG03 As Ax3, axG04 As Ax4 From acc190, acc180, fagent, nation, acc161, acc160, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = pa01 and substr(axG03, length(axG03) - 8, 6) = pa02 and substr(axG03, length(axG03) - 2, 1) = pa03 and substr(axG03, length(axG03) - 1, 2) = pa04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(5) & " group by axG01, axG03, axG04 "
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(ax4) * (-1) as Famount, sum(ax4*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, patent, acc1b0 " & _
'                              " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = pa01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = pa02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = pa03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = pa04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, PA01||'-'||PA02||'-'||PA03||'-'||PA04 "
'            '2007/12/6 end
'            '2006/11/20 END
'            '2006/4/3 END
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation, patent where axg01 = a1601 and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = pa01 and substr(axg03, length(axg03) - 8, 6) = pa02 and substr(axg03, length(axg03) - 2, 1) = pa03 and substr(axg03, length(axg03) - 1, 2) = pa04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1601 "
'         End If
'         If Text5 <> "" Then
'            strWhere(4) = " and tm12 = '" & Text5 & "'"
'         End If
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,(a1k08 - nvl(a1k06, 0))) USDamount from acc1k0, fagent, nation, trademark, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 (+) and a1k14 = tm02 (+) and a1k15 = tm03 (+) and a1k16 = tm04 (+) and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4)
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                                    "from acc0y0, fagent, nation, acc0z0, acc1k0, trademark where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04" & strWhere(1) & strWhere(4) & " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'         If Text7 = "" Then
'            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                                       "from acc151, acc150, fagent, nation, trademark where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1501 "
'            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, sum(axf04) as Famount, sum(axf04 * nvl(a1g03, 0)) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc151, acc150, fagent, nation, trademark, acc1g0, acc1h0, acc1i0 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2) & strWhere(4) & _
'                              " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, a1512, decode(a1h02, null, a1i03, a1h02), a1505, decode(a1506, a1520, 'Y', ''), a1504, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1501 "
'            '2006/11/20 MODIFY BY SONIA acc190串到acc151時, 可能會一筆串到多筆
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(axf04) as Famount, sum(axf04 * a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, acc151, acc150, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and a1908=a1b01(+)" & strWhere(3) & strWhere(4) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, TM01||'-'||TM02||'-'||TM03||'-'||TM04"
'            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axf04),台幣金額由sum(a1905)改為sum(axf04*a1906)
'            'StrSQLa = " Select axf01 as Ax1, axf03 As Ax3 From acc190, acc180, fagent, nation, acc151, acc150, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and a1908=a1b01(+)" & strWhere(3) & " group by axf01, axf03 "
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSQLa & " ) A1, acc150, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = tm01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = tm02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = tm03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = tm04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, TM01||'-'||TM02||'-'||TM03||'-'||TM04"
'            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
'            StrSQLa = " Select axf01 as Ax1, axf03 As Ax3, sum(axf04) As Ax4 From acc190, acc180, fagent, nation, acc151, acc150, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(3) & " group by axf01, axf03, axf04 "
'            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate," & _
'                         " a1903 as Currency, sum(ax4) as Famount, sum(ax4*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, fagent, nation, (" & StrSQLa & " ) A1, acc150, trademark, acc1b0 " & _
'                         " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = tm01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = tm02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = tm03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = tm04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, A1802, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, TM01||'-'||TM02||'-'||TM03||'-'||TM04 "
'            '2007/12/6 end
'            '2006/11/20 END
'            '2006/4/3 ADD BY SONIA
'            '2006/11/20 MODIFY BY SONIA acc190串到acc151時, 可能會一筆串到多筆
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(axG04) as Famount, sum(axG04 * a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, acc161, acc160, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = tm01 and substr(axG03, length(axG03) - 8, 6) = tm02 and substr(axG03, length(axG03) - 2, 1) = tm03 and substr(axG03, length(axG03) - 1, 2) = tm04 and a1908=a1b01(+)" & strWhere(5) & strWhere(4) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, TM01||'-'||TM02||'-'||TM03||'-'||TM04"
'            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axg04),台幣金額由sum(a1905)改為sum(axg04*a1906)
'            'StrSqlB = " Select axG01 as Ax1, axG03 As Ax3 From acc190, acc180, fagent, nation, acc161, acc160, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = tm01 and substr(axG03, length(axG03) - 8, 6) = tm02 and substr(axG03, length(axG03) - 2, 1) = tm03 and substr(axG03, length(axG03) - 1, 2) = tm04 and a1908=a1b01(+)" & strWhere(5) & " group by axG01, axG03 "
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSqlB & " ) A1, acc160, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = tm01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = tm02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = tm03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = tm04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, TM01||'-'||TM02||'-'||TM03||'-'||TM04"
'            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
'            StrSqlB = " Select axG01 as Ax1, axG03 As Ax3, axG04 As Ax4 From acc190, acc180, fagent, nation, acc161, acc160, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = tm01 and substr(axG03, length(axG03) - 8, 6) = tm02 and substr(axG03, length(axG03) - 2, 1) = tm03 and substr(axG03, length(axG03) - 1, 2) = tm04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(5) & " group by axG01, axG03, axG04 "
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(ax4) * (-1) as Famount, sum(ax4*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, fagent, nation, (" & StrSqlB & " ) A1, acc160, trademark, acc1b0 " & _
'                              " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = tm01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = tm02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = tm03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = tm04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, TM01||'-'||TM02||'-'||TM03||'-'||TM04 "
'            '2007/12/6 end
'            '2006/11/20 END
'            '2006/4/3 END
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc161, acc160, fagent, nation, trademark where axg01 = a1601 and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = tm01 and substr(axg03, length(axg03) - 8, 6) = tm02 and substr(axg03, length(axg03) - 2, 1) = tm03 and substr(axg03, length(axg03) - 1, 2) = tm04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1601 "
'         End If
'         If Text5 <> "" Then
'            strWhere(4) = " and sp11 = '" & Text5 & "'"
'         End If
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,(a1k08 - nvl(a1k06, 0))) USDamount from acc1k0, fagent, nation, servicepractice, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 (+) and a1k14 = sp02 (+) and a1k15 = sp03 (+) and a1k16 = sp04 (+) and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4)
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                                    "from acc0y0, fagent, nation, acc0z0, acc1k0, servicepractice where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04" & strWhere(1) & strWhere(4) & " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'         If Text7 = "" Then
'            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                     " from acc151, acc150, fagent, nation, servicepractice where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1501 "
'            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, sum(axf04) as Famount, sum(axf04 * nvl(a1g03, 0)) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc151, acc150, fagent, nation, servicepractice, acc1g0, acc1h0, acc1i0 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2) & strWhere(4) & _
'                              " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, a1512, decode(a1h02, null, a1i03, a1h02), a1505, decode(a1506, a1520, 'Y', ''), a1504, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1501 "
'            '2006/11/20 MODIFY BY SONIA acc190串到acc151時, 可能會一筆串到多筆
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27 " &
'            '         " from acc190, acc180, fagent, nation, acc151, acc150, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and a1908=a1b01(+)" & strWhere(3) & strWhere(4) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, SP01||'-'||SP02||'-'||SP03||'-'||SP04"
'            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axf04),台幣金額由sum(a1905)改為sum(axf04*a1906)
'            'StrSQLa = " select axf01 As Ax1, axf03 As Ax3 From acc190, acc180, fagent, nation, acc151, acc150, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and a1908=a1b01(+)" & strWhere(3) & " group by axf01, axf03 "
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = sp01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = sp02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = sp03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = sp04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, SP01||'-'||SP02||'-'||SP03||'-'||SP04"
'            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
'            StrSQLa = " select axf01 As Ax1, axf03 As Ax3, sum(axf04) As Ax4 From acc190, acc180, fagent, nation, acc151, acc150, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(3) & " group by axf01, axf03, axf04 "
'            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate," & _
'                         " a1903 as Currency, sum(ax4) as Famount, sum(ax4*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, servicepractice, acc1b0 " & _
'                         " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = sp01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = sp02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = sp03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = sp04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, A1802, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, SP01||'-'||SP02||'-'||SP03||'-'||SP04 "
'            '2007/12/6 end
'            '2006/11/20 END
'            '2006/4/3 ADD BY SONIA
'            '2006/11/20 MODIFY BY SONIA acc190串到acc151時, 可能會一筆串到多筆
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27 " & _
'            '         " from acc190, acc180, fagent, nation, acc161, acc160, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = sp01 and substr(axG03, length(axG03) - 8, 6) = sp02 and substr(aXG03, length(axG03) - 2, 1) = sp03 and substr(axG03, length(axG03) - 1, 2) = sp04 and a1908=a1b01(+)" & strWhere(5) & strWhere(4) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, SP01||'-'||SP02||'-'||SP03||'-'||SP04"
'            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axg04),台幣金額由sum(a1905)改為sum(axg04*a1906)
'            'StrSqlB = " select axG01 As Ax1, axG03 As Ax3 From acc190, acc180, fagent, nation, acc161, acc160, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = sp01 and substr(axG03, length(axG03) - 8, 6) = sp02 and substr(axG03, length(axG03) - 2, 1) = sp03 and substr(axG03, length(axG03) - 1, 2) = sp04 and a1908=a1b01(+)" & strWhere(5) & " group by axG01, axG03 "
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = sp01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = sp02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = sp03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = sp04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, SP01||'-'||SP02||'-'||SP03||'-'||SP04"
'            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
'            StrSqlB = " select axG01 As Ax1, axG03 As Ax3, axG04 As Ax4 From acc190, acc180, fagent, nation, acc161, acc160, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = sp01 and substr(axG03, length(axG03) - 8, 6) = sp02 and substr(axG03, length(axG03) - 2, 1) = sp03 and substr(axG03, length(axG03) - 1, 2) = sp04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(5) & " group by axG01, axG03, axG04 "
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(ax4) * (-1) as Famount, sum(ax4*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, servicepractice, acc1b0 " & _
'                              " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = sp01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = sp02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = sp03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = sp04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, SP01||'-'||SP02||'-'||SP03||'-'||SP04 "
'            '2007/12/6 end
'            '2006/11/20 END
'            '2006/4/3 END
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                     " from acc161, acc160, fagent, nation, servicepractice where axg01 = a1601 and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = sp01 and substr(axg03, length(axg03) - 8, 6) = sp02 and substr(axg03, length(axg03) - 2, 1) = sp03 and substr(axg03, length(axg03) - 1, 2) = sp04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1601 "
'         End If
'         '2005/8/3 ADD BY SONIA 加入LAWCASE
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,(a1k08 - nvl(a1k06, 0))) USDamount from acc1k0, fagent, nation, LAWCASE, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = LC01 (+) and a1k14 = LC02 (+) and a1k15 = LC03 (+) and a1k16 = LC04 (+) and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4)
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                                    "from acc0y0, fagent, nation, acc0z0, acc1k0, LAWCASE where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04" & strWhere(1) & strWhere(4) & " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'         If Text7 = "" Then
'            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                     " from acc151, acc150, fagent, nation, LAWCASE where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1501 "
'            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, sum(axf04) as Famount, sum(axf04 * nvl(a1g03, 0)) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc151, acc150, fagent, nation, LAWCASE, acc1g0, acc1h0, acc1i0 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04 and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2) & strWhere(4) & _
'                              " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, a1512, decode(a1h02, null, a1i03, a1h02), a1505, decode(a1506, a1520, 'Y', ''), a1504, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1501 "
'            '2006/11/20 MODIFY BY SONIA acc190串到acc151時, 可能會一筆串到多筆
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27 " & _
'            '         " from acc190, acc180, fagent, nation, acc151, acc150, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04 and a1908=a1b01(+)" & strWhere(3) & strWhere(4) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, LC01||'-'||LC02||'-'||LC03||'-'||LC04"
'            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axf04),台幣金額由sum(a1905)改為sum(axf04*a1906)
'            'StrSQLa = " select axf01 As Ax1, axf03 As Ax3 From acc190, acc180, fagent, nation, acc151, acc150, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04 and a1908=a1b01(+)" & strWhere(3) & " group by axf01, axf03 "
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = LC01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = LC02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = LC03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = LC04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, LC01||'-'||LC02||'-'||LC03||'-'||LC04"
'            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
'            StrSQLa = " select axf01 As Ax1, axf03 As Ax3, sum(axf04) As Ax4 From acc190, acc180, fagent, nation, acc151, acc150, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(3) & " group by axf01, axf03, axf04 "
'            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate," & _
'                         " a1903 as Currency, sum(ax4) as Famount, sum(ax4*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, LAWCASE, acc1b0 " & _
'                         " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = LC01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = LC02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = LC03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = LC04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, A1802, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, LC01||'-'||LC02||'-'||LC03||'-'||LC04 "
'            '2007/12/6 end
'            '2006/11/20 END
'            '2006/4/3 ADD BY SONIA
'            '2006/11/20 MODIFY BY SONIA acc190串到acc151時, 可能會一筆串到多筆
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27 " & _
'            '         " from acc190, acc180, fagent, nation, acc161, acc160, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = LC01 and substr(axG03, length(axG03) - 8, 6) = LC02 and substr(axG03, length(axG03) - 2, 1) = LC03 and substr(axG03, length(axG03) - 1, 2) = LC04 and a1908=a1b01(+)" & strWhere(5) & strWhere(4) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, LC01||'-'||LC02||'-'||LC03||'-'||LC04"
'            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axg04),台幣金額由sum(a1905)改為sum(axg04*a1906)
'            'StrSqlB = " select axG01 As Ax1, axG03 As Ax3 From acc190, acc180, fagent, nation, acc161, acc160, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = LC01 and substr(axG03, length(axG03) - 8, 6) = LC02 and substr(axG03, length(axG03) - 2, 1) = LC03 and substr(axG03, length(axG03) - 1, 2) = LC04 and a1908=a1b01(+)" & strWhere(5) & " group by axG01, axG03 "
'            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = LC01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = LC02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = LC03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = LC04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, LC01||'-'||LC02||'-'||LC03||'-'||LC04"
'            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
'            StrSqlB = " select axG01 As Ax1, axG03 As Ax3, axG04 As Ax4 From acc190, acc180, fagent, nation, acc161, acc160, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = LC01 and substr(axG03, length(axG03) - 8, 6) = LC02 and substr(axG03, length(axG03) - 2, 1) = LC03 and substr(axG03, length(axG03) - 1, 2) = LC04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(5) & " group by axG01, axG03, axG04 "
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(ax4) * (-1) as Famount, sum(ax4*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, LAWCASE, acc1b0 " & _
'                              " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = LC01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = LC02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = LC03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = LC04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, LC01||'-'||LC02||'-'||LC03||'-'||LC04 "
'            '2007/12/6 end
'            '2006/11/20 END
'            '2006/4/3 END
'            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
'            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                     " from acc161, acc160, fagent, nation, LAWCASE where axg01 = a1601 and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = LC01 and substr(axg03, length(axg03) - 8, 6) = LC02 and substr(axg03, length(axg03) - 2, 1) = LC03 and substr(axg03, length(axg03) - 1, 2) = LC04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1601 "
'         End If
'         '2005/8/3 END
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, decode(a1k18,'USD',null,(a1k08 - nvl(a1k06, 0))) USDamount from acc1k0, fagent, nation, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0)
'         '2010/6/29 MODIFY BY SONIA 抵帳幣別不可抓A1K18請款幣別X09901818
'         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1k17 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, NVL(A1H03,A1I05) as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k08 - nvl(a1k06, 0)) * nvl(a1g02, 0) as Namount, null as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort , a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount from acc1k0, fagent, nation, acc140, acc1g0, acc1h0, acc1i0 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k01 = a1403 (+) and a1k17 = a1g01 (+) and a1k17 = a1h01 (+) and a1k17 = a1i01 (+) and a1k17 is not null" & strWhere(0)
'         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount " & _
'                                         "from acc0y0, fagent, nation, acc0z0, acc1k0 where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01" & strWhere(1) & " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
         
'專利
         If Text5 <> "" Then
            strWhere(4) = " and pa11 = '" & Text5 & "'"
         End If
'請款單
         strSql = "select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc1k0, fagent, nation, patent, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4)
'收款
         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                    "from acc0y0, fagent, nation, acc0z0, acc1k0, patent where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04" & strWhere(1) & strWhere(4) & " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
         If Text7 = "" Then   '無請款單號時
'帳單ACC150
            '2014/11/24 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                       "from acc151, acc150, fagent, nation, patent where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1501 "
            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                       "from acc151, acc150, fagent, nation, patent, acc190 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1501=a1902(+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1501 "
            '2014/11/24 end
'抵帳資料
            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, sum(axf04) as Famount, sum(axf04 * nvl(a1g03, 0)) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc151, acc150, fagent, nation, patent, acc1g0, acc1h0, acc1i0 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2) & strWhere(4) & _
                              " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, a1512, decode(a1h02, null, a1i03, a1h02), a1505, decode(a1506, a1520, 'Y', ''), a1504, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1501 "
'帳單結匯(有匯票號)
            '2006/11/20 MODIFY BY SONIA acc190串到acc151時, 可能會一筆串到多筆
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27  from acc190, acc180, fagent, nation, acc151, acc150, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and a1908=a1b01(+)" & strWhere(3) & strWhere(4) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, PA01||'-'||PA02||'-'||PA03||'-'||PA04"
            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axf04),台幣金額由sum(a1905)改為sum(axf04*a1906)
            'StrSQLa = " Select axf01 As Ax1, axf03 As Ax3 From acc190, acc180, fagent, nation, acc151, acc150, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and a1908=a1b01(+)" & strWhere(3) & " group by axf01, axf03 "
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = pa01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = pa02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = pa03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = pa04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, PA01||'-'||PA02||'-'||PA03||'-'||PA04"
            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
            StrSQLa = " Select axf01 As Ax1, axf03 As Ax3, sum(axf04) As Ax4 From acc190, acc180, fagent, nation, acc151, acc150, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(3) & " group by axf01, axf03, axf04 "
            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate," & _
                         " a1903 as Currency, sum(ax4) as Famount, sum(ax4*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = pa01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = pa02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = pa03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = pa04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, A1802, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, PA01||'-'||PA02||'-'||PA03||'-'||PA04 "
            '2007/12/6 end
            '2006/11/20 END
'抵帳單ACC160
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            '2014/11/26 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, patent where axg01 = a1601 and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = pa01 and substr(axg03, length(axg03) - 8, 6) = pa02 and substr(axg03, length(axg03) - 2, 1) = pa03 and substr(axg03, length(axg03) - 1, 2) = pa04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1601 "
            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, patent, ACC190 " & _
                              " where axg01 = a1601 AND A1601=A1902(+) and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = pa01 and substr(axg03, length(axg03) - 8, 6) = pa02 and substr(axg03, length(axg03) - 2, 1) = pa03 and substr(axg03, length(axg03) - 1, 2) = pa04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null), a1602, a1605, a1604, a1607, PA01||'-'||PA02||'-'||PA03||'-'||PA04, a1601 "
            '2014/11/26 end
'抵帳單結匯(有匯票號)
            '2006/4/3 ADD BY SONIA
            '2006/11/20 MODIFY BY SONIA acc190串到acc151時, 可能會一筆串到多筆
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27  from acc190, acc180, fagent, nation, acc161, acc160, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = pa01 and substr(axG03, length(axG03) - 8, 6) = pa02 and substr(axG03, length(axG03) - 2, 1) = pa03 and substr(axG03, length(axG03) - 1, 2) = pa04 and a1908=a1b01(+)" & strWhere(5) & strWhere(4) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, PA01||'-'||PA02||'-'||PA03||'-'||PA04"
            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axg04),台幣金額由sum(a1905)改為sum(axg04*a1906)
            'StrSqlB = " Select axG01 As Ax1, axG03 As Ax3 From acc190, acc180, fagent, nation, acc161, acc160, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = pa01 and substr(axG03, length(axG03) - 8, 6) = pa02 and substr(axG03, length(axG03) - 2, 1) = pa03 and substr(axG03, length(axG03) - 1, 2) = pa04 and a1908=a1b01(+)" & strWhere(5) & " group by axG01, axG03 "
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = pa01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = pa02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = pa03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = pa04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, PA01||'-'||PA02||'-'||PA03||'-'||PA04"
            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
            StrSqlB = " Select axG01 As Ax1, axG03 As Ax3, axG04 As Ax4 From acc190, acc180, fagent, nation, acc161, acc160, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = pa01 and substr(axG03, length(axG03) - 8, 6) = pa02 and substr(axG03, length(axG03) - 2, 1) = pa03 and substr(axG03, length(axG03) - 1, 2) = pa04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(5) & " group by axG01, axG03, axG04 "
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(ax4) * (-1) as Famount, sum(ax4*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, PA01||'-'||PA02||'-'||PA03||'-'||PA04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, patent, acc1b0 " & _
                              " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = pa01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = pa02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = pa03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = pa04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, PA01||'-'||PA02||'-'||PA03||'-'||PA04 "
            '2007/12/6 end
            '2006/11/20 END
            '2006/4/3 END
         End If
'商標
         If Text5 <> "" Then
            strWhere(4) = " and tm12 = '" & Text5 & "'"
         End If
'請款單
         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc1k0, fagent, nation, trademark, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 (+) and a1k14 = tm02 (+) and a1k15 = tm03 (+) and a1k16 = tm04 (+) and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4)
'收款
         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                    "from acc0y0, fagent, nation, acc0z0, acc1k0, trademark where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04" & strWhere(1) & strWhere(4) & " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
         If Text7 = "" Then
'帳單ACC150
            '2014/11/24 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                       "from acc151, acc150, fagent, nation, trademark where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1501 "
            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                       "from acc151, acc150, fagent, nation, trademark, acc190 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1501=a1902(+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1501 "
            '2014/11/24 end
'抵帳資料
            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, sum(axf04) as Famount, sum(axf04 * nvl(a1g03, 0)) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc151, acc150, fagent, nation, trademark, acc1g0, acc1h0, acc1i0 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2) & strWhere(4) & _
                              " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, a1512, decode(a1h02, null, a1i03, a1h02), a1505, decode(a1506, a1520, 'Y', ''), a1504, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1501 "
'帳單結匯(有匯票號)
            '2006/11/20 MODIFY BY SONIA acc190串到acc151時, 可能會一筆串到多筆
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(axf04) as Famount, sum(axf04 * a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, acc151, acc150, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and a1908=a1b01(+)" & strWhere(3) & strWhere(4) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, TM01||'-'||TM02||'-'||TM03||'-'||TM04"
            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axf04),台幣金額由sum(a1905)改為sum(axf04*a1906)
            'StrSQLa = " Select axf01 as Ax1, axf03 As Ax3 From acc190, acc180, fagent, nation, acc151, acc150, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and a1908=a1b01(+)" & strWhere(3) & " group by axf01, axf03 "
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSQLa & " ) A1, acc150, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = tm01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = tm02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = tm03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = tm04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, TM01||'-'||TM02||'-'||TM03||'-'||TM04"
            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
            StrSQLa = " Select axf01 as Ax1, axf03 As Ax3, sum(axf04) As Ax4 From acc190, acc180, fagent, nation, acc151, acc150, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(3) & " group by axf01, axf03, axf04 "
            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate," & _
                         " a1903 as Currency, sum(ax4) as Famount, sum(ax4*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, fagent, nation, (" & StrSQLa & " ) A1, acc150, trademark, acc1b0 " & _
                         " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = tm01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = tm02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = tm03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = tm04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, A1802, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, TM01||'-'||TM02||'-'||TM03||'-'||TM04 "
            '2007/12/6 end
            '2006/11/20 END
'抵帳單ACC160
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            '2014/11/26 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, trademark where axg01 = a1601 and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = tm01 and substr(axg03, length(axg03) - 8, 6) = tm02 and substr(axg03, length(axg03) - 2, 1) = tm03 and substr(axg03, length(axg03) - 1, 2) = tm04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1601 "
            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, trademark, ACC190 " & _
                              " where axg01 = a1601 AND A1601=A1902(+) and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = tm01 and substr(axg03, length(axg03) - 8, 6) = tm02 and substr(axg03, length(axg03) - 2, 1) = tm03 and substr(axg03, length(axg03) - 1, 2) = tm04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null), a1602, a1605, a1604, a1607, TM01||'-'||TM02||'-'||TM03||'-'||TM04, a1601 "
            '2014/11/26 end
'抵帳單結匯(有匯票號)
            '2006/4/3 ADD BY SONIA
            '2006/11/20 MODIFY BY SONIA acc190串到acc151時, 可能會一筆串到多筆
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(axG04) as Famount, sum(axG04 * a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, acc161, acc160, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = tm01 and substr(axG03, length(axG03) - 8, 6) = tm02 and substr(axG03, length(axG03) - 2, 1) = tm03 and substr(axG03, length(axG03) - 1, 2) = tm04 and a1908=a1b01(+)" & strWhere(5) & strWhere(4) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, TM01||'-'||TM02||'-'||TM03||'-'||TM04"
            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axg04),台幣金額由sum(a1905)改為sum(axg04*a1906)
            'StrSqlB = " Select axG01 as Ax1, axG03 As Ax3 From acc190, acc180, fagent, nation, acc161, acc160, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = tm01 and substr(axG03, length(axG03) - 8, 6) = tm02 and substr(axG03, length(axG03) - 2, 1) = tm03 and substr(axG03, length(axG03) - 1, 2) = tm04 and a1908=a1b01(+)" & strWhere(5) & " group by axG01, axG03 "
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSqlB & " ) A1, acc160, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = tm01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = tm02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = tm03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = tm04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, TM01||'-'||TM02||'-'||TM03||'-'||TM04"
            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
            StrSqlB = " Select axG01 as Ax1, axG03 As Ax3, axG04 As Ax4 From acc190, acc180, fagent, nation, acc161, acc160, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = tm01 and substr(axG03, length(axG03) - 8, 6) = tm02 and substr(axG03, length(axG03) - 2, 1) = tm03 and substr(axG03, length(axG03) - 1, 2) = tm04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(5) & " group by axG01, axG03, axG04 "
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(ax4) * (-1) as Famount, sum(ax4*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, TM01||'-'||TM02||'-'||TM03||'-'||TM04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, fagent, nation, (" & StrSqlB & " ) A1, acc160, trademark, acc1b0 " & _
                              " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = tm01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = tm02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = tm03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = tm04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, TM01||'-'||TM02||'-'||TM03||'-'||TM04 "
            '2007/12/6 end
            '2006/11/20 END
            '2006/4/3 END
         End If
'服務業務
         If Text5 <> "" Then
            strWhere(4) = " and sp11 = '" & Text5 & "'"
         End If
'請款單
         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc1k0, fagent, nation, servicepractice, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 (+) and a1k14 = sp02 (+) and a1k15 = sp03 (+) and a1k16 = sp04 (+) and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4)
'收款
         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                    "from acc0y0, fagent, nation, acc0z0, acc1k0, servicepractice where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04" & strWhere(1) & strWhere(4) & " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
         If Text7 = "" Then
'帳單ACC150
            '2014/11/24 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                     " from acc151, acc150, fagent, nation, servicepractice where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1501 "
            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                     " from acc151, acc150, fagent, nation, servicepractice, acc190 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1501=a1902(+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1501 "
            '2014/11/24 end
'抵帳資料
            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, sum(axf04) as Famount, sum(axf04 * nvl(a1g03, 0)) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc151, acc150, fagent, nation, servicepractice, acc1g0, acc1h0, acc1i0 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2) & strWhere(4) & _
                              " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, a1512, decode(a1h02, null, a1i03, a1h02), a1505, decode(a1506, a1520, 'Y', ''), a1504, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1501 "
'帳單結匯(有匯票號)
            '2006/11/20 MODIFY BY SONIA acc190串到acc151時, 可能會一筆串到多筆
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27 " &
            '         " from acc190, acc180, fagent, nation, acc151, acc150, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and a1908=a1b01(+)" & strWhere(3) & strWhere(4) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, SP01||'-'||SP02||'-'||SP03||'-'||SP04"
            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axf04),台幣金額由sum(a1905)改為sum(axf04*a1906)
            'StrSQLa = " select axf01 As Ax1, axf03 As Ax3 From acc190, acc180, fagent, nation, acc151, acc150, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and a1908=a1b01(+)" & strWhere(3) & " group by axf01, axf03 "
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = sp01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = sp02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = sp03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = sp04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, SP01||'-'||SP02||'-'||SP03||'-'||SP04"
            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
            StrSQLa = " select axf01 As Ax1, axf03 As Ax3, sum(axf04) As Ax4 From acc190, acc180, fagent, nation, acc151, acc150, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(3) & " group by axf01, axf03, axf04 "
            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate," & _
                         " a1903 as Currency, sum(ax4) as Famount, sum(ax4*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, servicepractice, acc1b0 " & _
                         " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = sp01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = sp02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = sp03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = sp04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, A1802, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, SP01||'-'||SP02||'-'||SP03||'-'||SP04 "
            '2007/12/6 end
            '2006/11/20 END
'抵帳單ACC160
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            '2014/11/26 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                     " from acc161, acc160, fagent, nation, servicepractice where axg01 = a1601 and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = sp01 and substr(axg03, length(axg03) - 8, 6) = sp02 and substr(axg03, length(axg03) - 2, 1) = sp03 and substr(axg03, length(axg03) - 1, 2) = sp04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1601 "
            strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, servicepractice, ACC190 " & _
                              " where axg01 = a1601 AND A1601=A1902(+) and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = sp01 and substr(axg03, length(axg03) - 8, 6) = sp02 and substr(axg03, length(axg03) - 2, 1) = sp03 and substr(axg03, length(axg03) - 1, 2) = sp04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null), a1602, a1605, a1604, a1607, SP01||'-'||SP02||'-'||SP03||'-'||SP04, a1601 "
            '2014/11/26 end
'抵帳單結匯(有匯票號)
            '2006/4/3 ADD BY SONIA
            '2006/11/20 MODIFY BY SONIA acc190串到acc151時, 可能會一筆串到多筆
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27 " & _
            '         " from acc190, acc180, fagent, nation, acc161, acc160, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = sp01 and substr(axG03, length(axG03) - 8, 6) = sp02 and substr(aXG03, length(axG03) - 2, 1) = sp03 and substr(axG03, length(axG03) - 1, 2) = sp04 and a1908=a1b01(+)" & strWhere(5) & strWhere(4) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, SP01||'-'||SP02||'-'||SP03||'-'||SP04"
            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axg04),台幣金額由sum(a1905)改為sum(axg04*a1906)
            'StrSqlB = " select axG01 As Ax1, axG03 As Ax3 From acc190, acc180, fagent, nation, acc161, acc160, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = sp01 and substr(axG03, length(axG03) - 8, 6) = sp02 and substr(axG03, length(axG03) - 2, 1) = sp03 and substr(axG03, length(axG03) - 1, 2) = sp04 and a1908=a1b01(+)" & strWhere(5) & " group by axG01, axG03 "
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = sp01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = sp02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = sp03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = sp04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, SP01||'-'||SP02||'-'||SP03||'-'||SP04"
            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
            StrSqlB = " select axG01 As Ax1, axG03 As Ax3, axG04 As Ax4 From acc190, acc180, fagent, nation, acc161, acc160, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = sp01 and substr(axG03, length(axG03) - 8, 6) = sp02 and substr(axG03, length(axG03) - 2, 1) = sp03 and substr(axG03, length(axG03) - 1, 2) = sp04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(5) & " group by axG01, axG03, axG04 "
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(ax4) * (-1) as Famount, sum(ax4*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, SP01||'-'||SP02||'-'||SP03||'-'||SP04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, servicepractice, acc1b0 " & _
                              " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = sp01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = sp02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = sp03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = sp04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, SP01||'-'||SP02||'-'||SP03||'-'||SP04 "
            '2007/12/6 end
            '2006/11/20 END
            '2006/4/3 END
         End If
'法務2005/8/3 ADD BY SONIA 加入LAWCASE
'請款單
        'Modified by Lydia 2018/02/13 拿掉strWhere(4) ; 有基本檔才抓
        'strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc1k0, fagent, nation, LAWCASE, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = LC01 (+) and a1k14 = LC02 (+) and a1k15 = LC03 (+) and a1k16 = LC04 (+) and a1k01 = a1403 (+)" & strWhere(0) & strWhere(4)
        If Trim(Text7 & Text1 & Text3 & Text8) <> "" Then
            strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161" & _
                                      " from acc1k0, fagent, nation, LAWCASE, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = LC01 (+) and a1k14 = LC02 (+) and a1k15 = LC03 (+) and a1k16 = LC04 (+) and a1k01 = a1403 (+)" & strWhere(0) & _
                                      " and lc01 is not null"
        End If
        'end 2018/02/13
'收款
        'Modified by Lydia 2018/02/13 拿掉strWhere(4) ; 有基本檔才抓
        'strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                   "from acc0y0, fagent, nation, acc0z0, acc1k0, LAWCASE where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04" & strWhere(1) & strWhere(4) & " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
        If Trim(Text7 & Text1 & Text3 & Text8) <> "" Then
            strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                      "from acc0y0, fagent, nation, acc0z0, acc1k0, LAWCASE where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04" & strWhere(1) & _
                                      " and lc01 is not null group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
        End If
        'end 2018/02/13
         If Text7 = "" Then
'帳單ACC150
            '2014/11/24 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501, a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                     " from acc151, acc150, fagent, nation, LAWCASE where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501, a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1501 "
            'Modified by Lydia 2018/02/13 拿掉strWhere(4) ; 有基本檔才抓
 '           strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                     " from acc151, acc150, fagent, nation, LAWCASE, acc190 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1501=a1902(+) and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04" & strWhere(2) & strWhere(4) & " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1501 "
            '2014/11/24 end
            strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*') as DocNo, a1502 as DocDate, a1505 as Currency, sum(axf04) as Famount, 0 as Namount, decode(a1506, a1520, 'Y', '') as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                     " from acc151, acc150, fagent, nation, LAWCASE, acc190 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1501=a1902(+) and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04" & strWhere(2) & _
                                     " and lc01 is not null group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, decode(a1507, null, a1501||decode(a1908||a1512,null,decode(a1901,null,null,'>'),null), a1501||'*'), a1502, a1505, decode(a1506, a1520, 'Y', ''), a1504, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1501 "
            'end 2018/02/13
'抵帳資料
            'Modified by Lydia 2018/02/13 拿掉strWhere(4) ; 有基本檔才抓
            'strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, sum(axf04) as Famount, sum(axf04 * nvl(a1g03, 0)) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc151, acc150, fagent, nation, LAWCASE, acc1g0, acc1h0, acc1i0 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04 and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2) & strWhere(4) & _
                              " group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, a1512, decode(a1h02, null, a1i03, a1h02), a1505, decode(a1506, a1520, 'Y', ''), a1504, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1501 "
            If Trim(Text7 & Text1 & Text3 & Text8) <> "" Then
                strSql = strSql & " union select a1503 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1512 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, a1505 as Currency, sum(axf04) as Famount, sum(axf04 * nvl(a1g03, 0)) as Namount, null as Close, 0 as Tamount, 0 as Oamount, a1504 as DNno, a1501 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161" & _
                                         " from acc151, acc150, fagent, nation, LAWCASE, acc1g0, acc1h0, acc1i0 where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04 and a1512 = a1g01 (+) and a1512 = a1h01 (+) and a1512 = a1i01 (+) and a1512 is not null" & strWhere(2) & _
                                        " and lc01 is not null  group by a1503, nvl(fa05, nvl(fa06, fa04)), na03, a1512, decode(a1h02, null, a1i03, a1h02), a1505, decode(a1506, a1520, 'Y', ''), a1504, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1501 "
            End If
            'end 2018/02/13
'帳單結匯(有匯票號)
            '2006/11/20 MODIFY BY SONIA acc190串到acc151時, 可能會一筆串到多筆
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27 " & _
            '         " from acc190, acc180, fagent, nation, acc151, acc150, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04 and a1908=a1b01(+)" & strWhere(3) & strWhere(4) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, LC01||'-'||LC02||'-'||LC03||'-'||LC04"
            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axf04),台幣金額由sum(a1905)改為sum(axf04*a1906)
            'StrSQLa = " select axf01 As Ax1, axf03 As Ax3 From acc190, acc180, fagent, nation, acc151, acc150, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04 and a1908=a1b01(+)" & strWhere(3) & " group by axf01, axf03 "
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = LC01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = LC02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = LC03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = LC04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, LC01||'-'||LC02||'-'||LC03||'-'||LC04"
            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
            StrSQLa = " select axf01 As Ax1, axf03 As Ax3, sum(axf04) As Ax4 From acc190, acc180, fagent, nation, acc151, acc150, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(3) & " group by axf01, axf03, axf04 "
            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, NVL(a1b03,A1802) as DocDate," & _
                         " a1903 as Currency, sum(ax4) as Famount, sum(ax4*a1906) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1501 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, fagent, nation, (" & StrSQLa & ") A1, acc150, LAWCASE, acc1b0 " & _
                         " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = LC01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = LC02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = LC03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = LC04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, A1802, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1501, LC01||'-'||LC02||'-'||LC03||'-'||LC04 "
            '2007/12/6 end
            '2006/11/20 END
'抵帳單ACC160
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            '2014/11/26 modify by sonia 加>付款中的符號(有acc190但無匯票號碼)
            'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601 as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                     " from acc161, acc160, fagent, nation, LAWCASE where axg01 = a1601 and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = LC01 and substr(axg03, length(axg03) - 8, 6) = LC02 and substr(axg03, length(axg03) - 2, 1) = LC03 and substr(axg03, length(axg03) - 1, 2) = LC04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601, a1602, a1605, a1604, a1607, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1601 "
            'Modified by Lydia 2018/02/13 拿掉strWhere(4) ; 有基本檔才抓
            'strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, LAWCASE, ACC190 " & _
                              " where axg01 = a1601 AND A1601=A1902(+) and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = LC01 and substr(axg03, length(axg03) - 8, 6) = LC02 and substr(axg03, length(axg03) - 2, 1) = LC03 and substr(axg03, length(axg03) - 1, 2) = LC04" & strWhere(5) & strWhere(4) & " group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null), a1602, a1605, a1604, a1607, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1601 "
            '2014/11/26 end
            If Trim(Text7 & Text1 & Text3 & Text8) <> "" Then
                strSql = strSql & " union select a1603 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null) as DocNo, a1602 as DocDate, a1605 as Currency, sum(axg04) * (-1) as Famount, 0 as Namount, decode(a1607, NULL, '', 'Y') as Close, 0 as Tamount, 0 as Oamount, a1604 as DNno, a1601 as Sno, '1' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc161, acc160, fagent, nation, LAWCASE, ACC190 " & _
                                         " where axg01 = a1601 AND A1601=A1902(+) and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axg03, 1, length(axg03) - 9) = LC01 and substr(axg03, length(axg03) - 8, 6) = LC02 and substr(axg03, length(axg03) - 2, 1) = LC03 and substr(axg03, length(axg03) - 1, 2) = LC04" & strWhere(5) & _
                                         " and lc01 is not null group by a1603, nvl(fa05, nvl(fa06, fa04)), na03, a1601||decode(a1908,null,decode(a1901,null,null,'>'),null), a1602, a1605, a1604, a1607, LC01||'-'||LC02||'-'||LC03||'-'||LC04, a1601 "
            End If
            'end 2018/02/13
'抵帳單結匯(有匯票號)
            '2006/4/3 ADD BY SONIA
            '2006/11/20 MODIFY BY SONIA acc190串到acc151時, 可能會一筆串到多筆
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27 " & _
            '         " from acc190, acc180, fagent, nation, acc161, acc160, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = LC01 and substr(axG03, length(axG03) - 8, 6) = LC02 and substr(axG03, length(axG03) - 2, 1) = LC03 and substr(axG03, length(axG03) - 1, 2) = LC04 and a1908=a1b01(+)" & strWhere(5) & strWhere(4) & " group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, LC01||'-'||LC02||'-'||LC03||'-'||LC04"
            '2007/12/6 modify by sonia 外幣金額由sum(a1904)改為sum(axg04),台幣金額由sum(a1905)改為sum(axg04*a1906)
            'StrSqlB = " select axG01 As Ax1, axG03 As Ax3 From acc190, acc180, fagent, nation, acc161, acc160, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = LC01 and substr(axG03, length(axG03) - 8, 6) = LC02 and substr(axG03, length(axG03) - 2, 1) = LC03 and substr(axG03, length(axG03) - 1, 2) = LC04 and a1908=a1b01(+)" & strWhere(5) & " group by axG01, axG03 "
            'strSQL = strSQL & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(a1904) as Famount, sum(a1905) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27 from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = LC01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = LC02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = LC03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = LC04 and a1908=a1b01(+) group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, LC01||'-'||LC02||'-'||LC03||'-'||LC04"
            'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
            StrSqlB = " select axG01 As Ax1, axG03 As Ax3, axG04 As Ax4 From acc190, acc180, fagent, nation, acc161, acc160, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = axG01 and substr(axG03, 1, length(axG03) - 9) = LC01 and substr(axG03, length(axG03) - 8, 6) = LC02 and substr(axG03, length(axG03) - 2, 1) = LC03 and substr(axG03, length(axG03) - 1, 2) = LC04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(5) & " group by axG01, axG03, axG04 "
            'Modify By Sindy 2012/8/15 明細資料抓Acc160資料抵帳單時,金額欄請改以負數表示
            strSql = strSql & " union select a1803 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1801 as DocNo, a1b03 as DocDate, a1903 as Currency, sum(ax4) * (-1) as Famount, sum(ax4*a1906) * (-1) as Namount, '' as Close, 0 as Tamount, 0 as Oamount, '' as DNno, a1601 as Sno, '2' as Sort, LC01||'-'||LC02||'-'||LC03||'-'||LC04 as a1k1316, '' as a1k28, '' as a1k27,0 as a1k30,0 as a1k10,0 as a1k12,'' as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc190, acc180, fagent, nation, (" & StrSqlB & ") A1, acc160, LAWCASE, acc1b0 " & _
                              " where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1601 and a1601 = A1.Ax1 and substr(A1.Ax3, 1, length(A1.Ax3) - 9) = LC01 and substr(A1.Ax3, length(A1.Ax3) - 8, 6) = LC02 and substr(A1.Ax3, length(A1.Ax3) - 2, 1) = LC03 and substr(A1.Ax3, length(A1.Ax3) - 1, 2) = LC04 and a1908=a1b01(+) and A1908 IS not NULL group by a1803, nvl(fa05, nvl(fa06, fa04)), na03, a1801, a1b03, a1903, a1601, LC01||'-'||LC02||'-'||LC03||'-'||LC04 "
            '2007/12/6 end
            '2006/11/20 END
            '2006/4/3 END
         End If
         '2005/8/3 END
'抓不到案號的請款單
         'Modified by Lydia 2018/02/13 比對基本檔無資料才算舊資料
         'strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc1k0, fagent, nation, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0)
         strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, decode(a1401,null,decode(a1k12,null,decode(nvl(a1k07,0),0,a1k01,a1k01||'@'),a1k01||'*'),a1k01||'$') as DocNo, a1k02 as DocDate, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k29 as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '1' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161" & _
                                  " from acc1k0, fagent, nation, acc140,PATENT,TRADEMARK,SERVICEPRACTICE,LAWCASE" & _
                                  " where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k01 = a1403 (+)" & strWhere(0) & _
                                  " AND A1K13=PA01(+) AND A1K14=PA02(+) AND A1K15=PA03(+) AND A1K16=PA04(+)" & _
                                  " AND A1K13=TM01(+) AND A1K14=TM02(+) AND A1K15=TM03(+) AND A1K16=TM04(+)" & _
                                  " AND A1K13=SP01(+) AND A1K14=SP02(+) AND A1K15=SP03(+) AND A1K16=SP04(+)" & _
                                  " AND A1K13=LC01(+) AND A1K14=LC02(+) AND A1K15=LC03(+) AND A1K16=LC04(+)" & _
                                  " AND PA01||TM01||SP01||LC01 IS NULL"
'抓不到案號的收款
         'Modified by Lydia 2018/02/13 比對基本檔無資料才算舊資料
         'strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                         "from acc0y0, fagent, nation, acc0z0, acc1k0 where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01" & strWhere(1) & " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
         strSql = strSql & " union select decode(a0y18, 1, a0y07, 2, a0y08, a0y09) as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a0y01 as DocNo, a0y02 as DocDate, a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, '' as Close, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 " & _
                                  " from acc0y0, fagent, nation, acc0z0, acc1k0,PATENT,TRADEMARK,SERVICEPRACTICE,LAWCASE" & _
                                  " where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01" & strWhere(1) & _
                                  " AND A1K13=PA01(+) AND A1K14=PA02(+) AND A1K15=PA03(+) AND A1K16=PA04(+)" & _
                                  " AND A1K13=TM01(+) AND A1K14=TM02(+) AND A1K15=TM03(+) AND A1K16=TM04(+)" & _
                                  " AND A1K13=SP01(+) AND A1K14=SP02(+) AND A1K15=SP03(+) AND A1K16=SP04(+)" & _
                                  " AND A1K13=LC01(+) AND A1K14=LC02(+) AND A1K15=LC03(+) AND A1K16=LC04(+)" & _
                                  " AND PA01||TM01||SP01||LC01 IS NULL" & _
                                  " group by decode(a0y18, 1, a0y07, 2, a0y08, a0y09), nvl(fa05, nvl(fa06, fa04)), na03, a0y01, a0y02, a0y03, a0y06, a1k01, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16, a1k28, a1k27,a1k30,a1k10,a1k12,a1k25 "
'抵帳資料收款
         '2010/6/29 MODIFY BY SONIA 抵帳幣別不可抓A1K18請款幣別X09901818
         If Trim(strWhere(0)) <> "" Then 'Added by Lydia 2018/02/13 無條件不可加入
              strSql = strSql & " union select a1k03 as FagentNo, nvl(fa05, nvl(fa06, fa04)) as FagentName, na03 as Nation, a1k17 as DocNo, decode(a1h02, null, a1i03, a1h02) as DocDate, NVL(A1H03,A1I05) as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k08 - nvl(a1k31, 0)) * nvl(a1g02, 0) as Namount, null as Close, a1k09 as Tamount, 0 as Oamount, '' as DNno, a1k01 as Sno, '2' as Sort , a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k28, a1k27,nvl(a1k30,0) as a1k30,a1k10,nvl(a1k12,0) as a1k12,nvl(a1k25,'') as a1k25,'" & strUserNum & "' as ID, null USDamount,null as PA161 from acc1k0, fagent, nation, acc140, acc1g0, acc1h0, acc1i0 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k01 = a1403 (+) and a1k17 = a1g01 (+) and a1k17 = a1h01 (+) and a1k17 = a1i01 (+) and a1k17 is not null" & strWhere(0)
         End If
   End Select
   
   Call ClearSumCol 'Add By Sindy 2012/8/14
   
   If strSql = "" Then
      InsertQueryLog (0) 'Add By Sindy 2010/12/21
      DataGrid1.Refresh
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   
   'Add By Sindy 2010/3/12
   cnnConnection.BeginTrans
   cnnConnection.Execute "delete from accrpt220 where id='" & strUserNum & "' "
   cnnConnection.Execute "insert into accrpt220 " & strSql
   cnnConnection.CommitTrans
   '2010/3/12 End
      
   If strSql <> "" Then
      'adoadodc1.Open strSQL & " order by Sno asc, Sort asc", adoTaie, adOpenStatic, adLockReadOnly
      'modify by sonia X09707856 再加日期排序條件
      adoadodc1.Open "select * from accrpt220 where id='" & strUserNum & "' order by Sno asc, Sort asc, DocDate asc ", adoTaie, adOpenStatic, adLockBatchOptimistic 'Modify By Sindy 2010/3/12 原為adLockReadOnly
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/21
      DataGrid1.Refresh
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Adodc1.Recordset.ReQuery
   SumShow
   If Adodc1.Recordset.RecordCount = 0 Then
      InsertQueryLog (0) 'Add By Sindy 2010/12/21
      Adodc1.Recordset.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   'Add by Morgan 2010/4/22 預設第一筆
   Else
      InsertQueryLog (Adodc1.Recordset.RecordCount) 'Add By Sindy 2010/12/21
      
      'Add By Sindy 2014/3/24 +特殊出名公司
      With Adodc1.Recordset
         .MoveFirst
         Do While Not .EOF
            strCaseNo = .Fields("a1k1316")
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
               .Fields("PA161") = "" & RsTemp.Fields(0)
            End If
            .MoveNext
         Loop
      End With
      DataGrid1.AllowUpdate = False '鎖住畫面,不可異動資料
      '2014/3/24 End
      
      Adodc1.Recordset.MoveFirst
      'Add by Amy 2013/10/30 +帳款處理訊息
      strExc(2) = ""
      strExc(0) = GetDizhang("" & Adodc1.Recordset.Fields("FagentNo"), , False) '代理人
      strExc(1) = GetDizhang("" & Adodc1.Recordset.Fields("a1k28"), , False) '請款對象
      
      If strExc(0) <> "" Or strExc(1) <> "" Then
        If strExc(0) = strExc(1) Then
            strExc(2) = "代理人/ 請款對象 編號 " & strExc(0)
        ElseIf strExc(0) <> "" And strExc(1) <> "" And strExc(0) <> strExc(1) Then
            strExc(2) = "    代理人編號 " & strExc(0) & "請款對象編號 " & strExc(1)
        ElseIf strExc(0) <> "" Then
            strExc(2) = "代理人編號 " & strExc(0)
        ElseIf strExc(1) <> "" Then
            strExc(2) = "請款對象編號 " & strExc(1)
        End If
            MsgBox strExc(2) & vbCrLf & "詳細情形請與財務處聯繫!!", vbInformation
      End If
      'end 2013/10/16
   End If
   
   Exit Sub
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   '2011/9/6 add by sonia
   Else
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
            If Text4 = "" Then
               Text4 = "0"
            End If
            If Text6 = "" Then
               Text6 = "00"
            End If
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
'         Text10 = MsgText(601)
'      Else
'         If IsNull(adoaccsum.Fields(1).Value) Then
'            Text10 = Format(adoaccsum.Fields(0).Value, FDollar)
'         Else
'            Text10 = Format(adoaccsum.Fields(0).Value - adoaccsum.Fields(1).Value, FDollar)
'         End If
'      End If
'   Else
'      Text10 = MsgText(601)
'   End If
'   adoaccsum.Close
'End Sub

'*************************************************
'  顯示本所案號名稱
'
'*************************************************
Private Sub CaseQuery()
   Text2 = "" 'add by sonia 2015/4/20
   Select Case Combo1
      Case ComboItem(121)
         Text2 = CaseNameShow(Text1, Text3, Text4, Text6, 1)
      Case ComboItem(122)
         Text2 = CaseNameShow(Text1, Text3, Text4, Text6, 2)
      Case ComboItem(123)
         Text2 = CaseNameShow(Text1, Text3, Text4, Text6, 3)
   End Select
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

''*************************************************
''  儲存請款資料
''
''*************************************************
'Private Sub Acc1k0Query()
'On Error GoTo Checking
'   strSql = ""
'   If Text7 <> MsgText(601) Then
'      strSql = strSql & " and a1k01 = '" & Text7 & "'"
'   End If
'   If Text1 <> MsgText(601) Then
'      strSql = strSql & " and a1k13 = '" & Text1 & "'"
'   End If
'   If Text3 <> MsgText(601) Then
'      strSql = strSql & " and a1k14 = '" & Text3 & "'"
'   End If
'   If Text4 <> MsgText(601) Then
'      strSql = strSql & " and a1k15 = '" & Text4 & "'"
'   End If
'   If Text6 <> MsgText(601) Then
'      strSql = strSql & " and a1k16 = '" & Text6 & "'"
'   End If
'   If Text5 <> MsgText(601) Then
'      Select Case Text1
'         Case "FCP", "CFP"
'            strSql = strSql & "and pa11 = '" & Text5 & "'"
'         Case "FCT", "CFT"
'            strSql = strSql & "and tm12 = '" & Text5 & "'"
'      End Select
'   End If
'   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'      strSql = strSql & " and a1k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'   End If
'   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'      strSql = strSql & " and a1k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'   End If
'   If Text9 <> MsgText(601) Then
'      Select Case Text3
'         Case "2", "6"
'            strSql = strSql & " and (a1k30 = 0 or a1k30 is null)"
'      End Select
'   End If
'   adoacc1k0.CursorLocation = adUseClient
'   Select Case Text1
'      Case "FCP"
'         adoacc1k0.Open "select * from acc1k0, patent, fagent where (a1k13 || a1k14 || a1k15 || a1k16) = (pa01 || pa02 || pa03 || pa04) and a1k03 = (fa01 || fa02)" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'      Case "FCT"
'         adoacc1k0.Open "select * from acc1k0, trademark, fagent where (a1k13 || a1k14 || a1k15 || a1k16) = (tm01 || tm02 || tm03 || tm04) and a1k03 = (fa01 || fa02)" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'      Case Else
'         adoacc1k0.Open "select * from acc1k0, patent, fagent where (a1k13 || a1k14 || a1k15 || a1k16) = (pa01 || pa02 || pa03 || pa04) and a1k03 = (fa01 || fa02)" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'   End Select
'   Do While adoacc1k0.EOF = False
'      adotmp2220.AddNew
'      adotmp2220.Fields("t22201").Value = adoacc1k0.Fields("a1k01").Value
'      If IsNull(adoacc1k0.Fields("a1k02").Value) Then
'         adotmp2220.Fields("t22202").Value = Null
'      Else
'         adotmp2220.Fields("t22202").Value = adoacc1k0.Fields("a1k02").Value
'      End If
'      If IsNull(adoacc1k0.Fields("a1k18").Value) Then
'         adotmp2220.Fields("t22203").Value = Null
'      Else
'         adotmp2220.Fields("t22203").Value = adoacc1k0.Fields("a1k18").Value
'      End If
'      If IsNull(adoacc1k0.Fields("a1k11").Value) Then
'         adotmp2220.Fields("t22204").Value = 0
'         adotmp2220.Fields("t22205").Value = 0
'      Else
'         If IsNull(adoacc1k0.Fields("a1k10").Value) Then
'            adotmp2220.Fields("t22204").Value = 0
'         Else
'            adotmp2220.Fields("t22204").Value = adoacc1k0.Fields("a1k11").Value / adoacc1k0.Fields("a1k10").Value
'         End If
'         adotmp2220.Fields("t22205").Value = adoacc1k0.Fields("a1k11").Value
'      End If
'      If IsNull(adoacc1k0.Fields("a1k09").Value) Then
'         adotmp2220.Fields("t22206").Value = 0
'      Else
'         adotmp2220.Fields("t22206").Value = adoacc1k0.Fields("a1k09").Value
'      End If
'      If IsNull(adoacc1k0.Fields("a1k29").Value) Then
'         adotmp2220.Fields("t22207").Value = Null
'      Else
'         adotmp2220.Fields("t22207").Value = adoacc1k0.Fields("a1k29").Value
'      End If
'      If IsNull(adoacc1k0.Fields("a1k03").Value) Then
'         adotmp2220.Fields("t22208").Value = Null
'      Else
'         adotmp2220.Fields("t22208").Value = adoacc1k0.Fields("a1k03").Value
'      End If
'      adotmp2220.UpdateBatch
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
'   If Text7 <> MsgText(601) Then
'      strSql = strSql & " and a1501 = '" & Text7 & "'"
'   End If
'   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'      strSql = strSql & " and a1502 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'   End If
'   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'      strSql = strSql & " and a1502 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'   End If
'   If Text8 <> MsgText(601) Then
'      strSql = strSql & " and a1504 = '" & Text8 & "'"
'   End If
'   If Text9 <> MsgText(601) Then
'      Select Case Text9
'         Case "4", "6"
'            strSql = strSql & " and (a1520 - a1510) > 0"
'      End Select
'   End If
'   adoacc150.CursorLocation = adUseClient
'   adoacc150.Open "select * from acc150, fagent where a1503 = (fa01 || fa02)" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'   Do While adoacc150.EOF = False
'      adotmp2220.AddNew
'      adotmp2220.Fields("t22201").Value = adoacc150.Fields("a1501").Value
'      If IsNull(adoacc150.Fields("a1502").Value) Then
'         adotmp2220.Fields("t22202").Value = Null
'      Else
'         adotmp2220.Fields("t22202").Value = adoacc150.Fields("a1502").Value
'      End If
'      If IsNull(adoacc150.Fields("a1505").Value) Then
'         adotmp2220.Fields("t22203").Value = Null
'      Else
'         adotmp2220.Fields("t22203").Value = adoacc150.Fields("a1505").Value
'      End If
'      If IsNull(adoacc150.Fields("a1510").Value) Then
'         adotmp2220.Fields("t22204").Value = 0
'         adotmp2220.Fields("t22205").Value = 0
'      Else
'         If IsNull(adoacc150.Fields("a1513").Value) Then
'            adotmp2220.Fields("t22204").Value = 0
'         Else
'            adotmp2220.Fields("t22204").Value = adoacc150.Fields("a1510").Value / adoacc150.Fields("a1513").Value
'         End If
'         adotmp2220.Fields("t22205").Value = adoacc150.Fields("a1k10").Value
'         If IsNull(adoacc150.Fields("a1520").Value) Then
'            adotmp2220.Fields("t22207").Value = Null
'         Else
'            If (adoacc150.Fields("a1510").Value - adoacc150.Fields("a1520").Value) < 0 Then
'               adotmp2220.Fields("t22207").Value = MsgText(602)
'            Else
'               adotmp2220.Fields("t22207").Value = Null
'            End If
'         End If
'      End If
'      If IsNull(adoacc150.Fields("a1503").Value) Then
'         adotmp2220.Fields("t22208").Value = Null
'      Else
'         adotmp2220.Fields("t22208").Value = adoacc150.Fields("a1503").Value
'      End If
'      adotmp2220.UpdateBatch
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
   
'   'Add By Sindy 2010/3/12 計算外幣金額
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
   
   '計算各欄位合計
   'Add By Sindy 2012/10/1
   If Text8 = "" Then '無代理人D/N No
   '2012/10/1 End
      '**************************************
      '外幣合計FC
      '**************************************
      strSql = ""
      adoaccsum.CursorLocation = adUseClient
      If Text5 <> "" Then
         strWhere(4) = " and pa11 = '" & Text5 & "'"
      End If
      '2013/9/24 modify by sonia 以下語法加入a1k01,否則相同金額會少算
      strSql = "select a1k01,a1k18,(a1k08 - nvl(a1k31, 0)) as Namount from acc1k0, patent where a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) '& " and a1k18='USD'"
      'Modify By Sindy 2013/1/15 Mark
      'strSql = strSql & " union select a1k18,(a1k11 - nvl(a1k06, 0)) / nvl(DNR03,0) as Namount from acc1k0, patent, DEBITNOTERATE where a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and a1k18<>'USD' and DNR01=a1k18 AND DNR02=(SELECT max(DNR02) FROM DEBITNOTERATE WHERE DNR01=a1k18 AND DNR02<=a1k02)"
      '2013/1/15 End
      If Text5 <> "" Then
         strWhere(4) = " and tm12 = '" & Text5 & "'"
      End If
      strSql = strSql & " union select a1k01,a1k18,(a1k08 - nvl(a1k31, 0)) as Namount from acc1k0, trademark where a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) '& " and a1k18='USD'"
      'Modify By Sindy 2013/1/15 Mark
      'strSql = strSql & " union select a1k18,(a1k11 - nvl(a1k06, 0)) / nvl(DNR03,0) as Namount from acc1k0, trademark, DEBITNOTERATE where a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and a1k18<>'USD' and DNR01=a1k18 AND DNR02=(SELECT max(DNR02) FROM DEBITNOTERATE WHERE DNR01=a1k18 AND DNR02<=a1k02)"
      '2013/1/15 End
      If Text5 <> "" Then
         strWhere(4) = " and sp11 = '" & Text5 & "'"
      End If
      strSql = strSql & " union select a1k01,a1k18,(a1k08 - nvl(a1k31, 0)) as Namount from acc1k0, servicepractice where a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) '& " and a1k18='USD'"
      'Modify By Sindy 2013/1/15 Mark
      'strSql = strSql & " union select a1k18,(a1k11 - nvl(a1k06, 0)) / nvl(DNR03,0) as Namount from acc1k0, servicepractice, DEBITNOTERATE where a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and a1k18<>'USD' and DNR01=a1k18 AND DNR02=(SELECT max(DNR02) FROM DEBITNOTERATE WHERE DNR01=a1k18 AND DNR02<=a1k02)"
      '2013/1/15 End
      strSql = strSql & " union select a1k01,a1k18,(a1k08 - nvl(a1k31, 0)) as Namount from acc1k0 where (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) '& " and a1k18='USD'"
      'Modify By Sindy 2013/1/15 Mark
      'strSql = strSql & " union select a1k18,(a1k11 - nvl(a1k06, 0)) / nvl(DNR03,0) as Namount from acc1k0, DEBITNOTERATE where (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & " and a1k18<>'USD' and DNR01=a1k18 AND DNR02=(SELECT max(DNR02) FROM DEBITNOTERATE WHERE DNR01=a1k18 AND DNR02<=a1k02)"
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
      strSql = ""
      adoaccsum.CursorLocation = adUseClient
      If Text5 <> "" Then
         strWhere(4) = " and pa11 = '" & Text5 & "'"
      End If
      '2013/9/24 modify by sonia 以下語法加入a1k01,否則相同金額會少算
      strSql = "select a1k01,a1k18,(a1k08 - nvl(a1k31, 0)) as Namount from acc1k0, patent where a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')" ' and a1k18='USD'"
      'Modify By Sindy 2013/1/15 Mark
      'strSql = strSql & " union select a1k18,(a1k11 - nvl(a1k06, 0)) / nvl(DNR03,0) as Namount from acc1k0, patent, DEBITNOTERATE where a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '') and a1k18<>'USD' and DNR01=a1k18 AND DNR02=(SELECT max(DNR02) FROM DEBITNOTERATE WHERE DNR01=a1k18 AND DNR02<=a1k02)"
      '2013/1/15 End
      strSql = strSql & " union select a1k01,a1k18,a0z04 * (-1) as Namount from acc0y0, acc0z0, acc1k0, patent where a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(1) & strWhere(4) & " and (a1k29 is null or a1k29 = '') and a1k30>0"
      If Text5 <> "" Then
         strWhere(4) = " and tm12 = '" & Text5 & "'"
      End If
      strSql = strSql & " union select a1k01,a1k18,(a1k08 - nvl(a1k31, 0)) as Namount from acc1k0, trademark where a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')" ' and a1k18='USD'"
      'Modify By Sindy 2013/1/15 Mark
      'strSql = strSql & " union select a1k18,(a1k11 - nvl(a1k06, 0)) / nvl(DNR03,0) as Namount from acc1k0, trademark, DEBITNOTERATE where a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '') and a1k18<>'USD' and DNR01=a1k18 AND DNR02=(SELECT max(DNR02) FROM DEBITNOTERATE WHERE DNR01=a1k18 AND DNR02<=a1k02)"
      '2013/1/15 End
      strSql = strSql & " union select a1k01,a1k18,a0z04 * (-1) as Namount from acc0y0, acc0z0, acc1k0, trademark where a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(1) & strWhere(4) & " and (a1k29 is null or a1k29 = '') and a1k30>0"
      If Text5 <> "" Then
         strWhere(4) = " and sp11 = '" & Text5 & "'"
      End If
      strSql = strSql & " union select a1k01,a1k18,(a1k08 - nvl(a1k31, 0)) as Namount from acc1k0, servicepractice where a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')" ' and a1k18='USD'"
      'Modify By Sindy 2013/1/15 Mark
      'strSql = strSql & " union select a1k18,(a1k11 - nvl(a1k06, 0)) / nvl(DNR03,0) as Namount from acc1k0, servicepractice, DEBITNOTERATE where a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '') and a1k18<>'USD' and DNR01=a1k18 AND DNR02=(SELECT max(DNR02) FROM DEBITNOTERATE WHERE DNR01=a1k18 AND DNR02<=a1k02)"
      '2013/1/15 End
      strSql = strSql & " union select a1k01,a1k18,a0z04 * (-1) as Namount from acc0y0, acc0z0, acc1k0, servicepractice where a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(1) & strWhere(4) & " and (a1k29 is null or a1k29 = '') and a1k30>0"
      strSql = strSql & " union select a1k01,a1k18,(a1k08 - nvl(a1k31, 0)) as Namount from acc1k0 where (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & " and (a1k29 is null or a1k29 = '')" ' and a1k18='USD'"
      'Modify By Sindy 2013/1/15 Mark
      'strSql = strSql & " union select a1k18,(a1k11 - nvl(a1k06, 0)) / nvl(DNR03,0) as Namount from acc1k0, DEBITNOTERATE where (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k18<>'USD' and DNR01=a1k18 AND DNR02=(SELECT max(DNR02) FROM DEBITNOTERATE WHERE DNR01=a1k18 AND DNR02<=a1k02)"
      '2013/1/15 End
      'Modified by Lydia 2018/02/13 拿掉strWhere(4)
      strSql = strSql & " union select a1k01,a1k18,a0z04 * (-1) as Namount from acc0y0, acc0z0, acc1k0 where a0y01 = a0z01 and a0z02 = a1k01 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(1) & " and (a1k29 is null or a1k29 = '') and a1k30>0"
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
      strSql = ""
      adoaccsum.CursorLocation = adUseClient
      If Text5 <> "" Then
         strWhere(4) = " and pa11 = '" & Text5 & "'"
      End If
      '2013/9/24 modify by sonia 以下語法加入a1k01,否則相同金額會少算
      strSql = "select a1k01,(a1k11 - nvl(a1k06, 0)) as Namount from acc1k0, patent where a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4)
      If Text5 <> "" Then
         strWhere(4) = " and tm12 = '" & Text5 & "'"
      End If
      strSql = strSql & " union select a1k01,(a1k11 - nvl(a1k06, 0)) as Namount from acc1k0, trademark where a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4)
      If Text5 <> "" Then
         strWhere(4) = " and sp11 = '" & Text5 & "'"
      End If
      strSql = strSql & " union select a1k01,(a1k11 - nvl(a1k06, 0)) as Namount from acc1k0, servicepractice where a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4)
      strSql = strSql & " union select a1k01,(a1k11 - nvl(a1k06, 0)) as Namount from acc1k0 where (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0)
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
      strSql = ""
      adoaccsum.CursorLocation = adUseClient
      If Text5 <> "" Then
         strWhere(4) = " and pa11 = '" & Text5 & "'"
      End If
      '2013/9/24 modify by sonia 以下語法加入a1k01,否則相同金額會少算
      strSql = "select a1k01,decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),2)) as Namount,decode(nvl(a1k30,0),0,a1k09,decode(sign(a1k30-a1k09),-1,a1k09-a1k30,0)) as Lawfee from acc1k0, patent where a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
      If Text5 <> "" Then
         strWhere(4) = " and tm12 = '" & Text5 & "'"
      End If
      strSql = strSql & " union select a1k01,decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),2)) as Namount,decode(nvl(a1k30,0),0,a1k09,decode(sign(a1k30-a1k09),-1,a1k09-a1k30,0)) as Lawfee from acc1k0, trademark where a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
      If Text5 <> "" Then
         strWhere(4) = " and sp11 = '" & Text5 & "'"
      End If
      strSql = strSql & " union select a1k01,decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),2)) as Namount,decode(nvl(a1k30,0),0,a1k09,decode(sign(a1k30-a1k09),-1,a1k09-a1k30,0)) as Lawfee from acc1k0, servicepractice where a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
      strSql = strSql & " union select a1k01,decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),2)) as Namount,decode(nvl(a1k30,0),0,a1k09,decode(sign(a1k30-a1k09),-1,a1k09-a1k30,0)) as Lawfee from acc1k0 where (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
      adoaccsum.Open "select sum(Namount),sum(Lawfee) from (" & strSql & ") New", adoTaie, adOpenStatic, adLockReadOnly
      If adoaccsum.RecordCount <> 0 Then
         If IsNull(adoaccsum.Fields(0).Value) Then
            Text11 = MsgText(601)
         Else
            Text11 = Format(adoaccsum.Fields(0).Value, FDollar)
         End If
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
   End If
   If Text7 = "" Then '無請款單號
      '**************************************
      '外幣CF合計
      '**************************************
      strSql = ""
      adoaccsum.CursorLocation = adUseClient
      If Text5 <> "" Then
         strWhere(4) = " and pa11 = '" & Text5 & "'"
      End If
      strSql = "select a1505,sum(axf04) as Namount from acc151, acc150, patent where axf01 = a1501 and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " group by a1505 "
      strSql = strSql & " union select a1605 as a1505,sum(axg04 * (-1)) as Namount from acc161, acc160, patent where axg01 = a1601 and substr(axg03, 1, length(axg03) - 9) = pa01 and substr(axg03, length(axg03) - 8, 6) = pa02 and substr(axg03, length(axg03) - 2, 1) = pa03 and substr(axg03, length(axg03) - 1, 2) = pa04 and a1607 is not null" & strWhere(5) & strWhere(4) & " group by a1605"
      If Text5 <> "" Then
         strWhere(4) = " and tm12 = '" & Text5 & "'"
      End If
      strSql = strSql & " union select a1505,sum(axf04) as Namount from acc151, acc150, trademark where axf01 = a1501 and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " group by a1505 "
      strSql = strSql & " union select a1605 as a1505,sum(axg04 * (-1)) as Namount from acc161, acc160, trademark where axg01 = a1601 and substr(axg03, 1, length(axg03) - 9) = tm01 and substr(axg03, length(axg03) - 8, 6) = tm02 and substr(axg03, length(axg03) - 2, 1) = tm03 and substr(axg03, length(axg03) - 1, 2) = tm04 and a1607 is not null" & strWhere(5) & strWhere(4) & " group by a1605"
      If Text5 <> "" Then
         strWhere(4) = " and sp11 = '" & Text5 & "'"
      End If
      strSql = strSql & " union select a1505,sum(axf04) as Namount from acc151, acc150, servicepractice where axf01 = a1501 and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " group by a1505 "
      strSql = strSql & " union select a1605 as a1505,sum(axg04 * (-1)) as Namount from acc161, acc160, servicepractice where axg01 = a1601 and substr(axg03, 1, length(axg03) - 9) = sp01 and substr(axg03, length(axg03) - 8, 6) = sp02 and substr(axg03, length(axg03) - 2, 1) = sp03 and substr(axg03, length(axg03) - 1, 2) = sp04 and a1607 is not null" & strWhere(5) & strWhere(4) & " group by a1605"
       'Modified by Lydia 2018/02/13 法務無申請案號
'      If Text5 = "" Then
'         strWhere(4) = ""
'         strSql = strSql & " union select a1505,sum(axf04) as Namount from acc151, acc150, lawcase where axf01 = a1501 and substr(axf03, 1, length(axf03) - 9) = lc01 and substr(axf03, length(axf03) - 8, 6) = lc02 and substr(axf03, length(axf03) - 2, 1) = lc03 and substr(axf03, length(axf03) - 1, 2) = lc04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " group by a1505 "
'         strSql = strSql & " union select a1605 as a1505,sum(axg04 * (-1)) as Namount from acc161, acc160, LAWCASE where axg01 = a1601 and substr(axg03, 1, length(axg03) - 9) = lc01 and substr(axg03, length(axg03) - 8, 6) = lc02 and substr(axg03, length(axg03) - 2, 1) = lc03 and substr(axg03, length(axg03) - 1, 2) = lc04 and a1607 is not null" & strWhere(5) & " group by a1605"
'      End If
       If Trim(strWhere(2)) <> "" Then strSql = strSql & " union select a1505,sum(axf04) as Namount from acc151, acc150, lawcase where axf01 = a1501 and substr(axf03, 1, length(axf03) - 9) = lc01 and substr(axf03, length(axf03) - 8, 6) = lc02 and substr(axf03, length(axf03) - 2, 1) = lc03 and substr(axf03, length(axf03) - 1, 2) = lc04 and (a1507 is null or a1507 = 0)" & strWhere(2) & " group by a1505 "
       If Trim(strWhere(5)) <> "" Then strSql = strSql & " union select a1605 as a1505,sum(axg04 * (-1)) as Namount from acc161, acc160, LAWCASE where axg01 = a1601 and substr(axg03, 1, length(axg03) - 9) = lc01 and substr(axg03, length(axg03) - 8, 6) = lc02 and substr(axg03, length(axg03) - 2, 1) = lc03 and substr(axg03, length(axg03) - 1, 2) = lc04 and a1607 is not null" & strWhere(5) & " group by a1605"
      'end 2018/02/13
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
      If Text5 <> "" Then
         strWhere(4) = " and pa11 = '" & Text5 & "'"
      End If
      'Modify By Sindy 2012/8/31 Y21399查詢時剛好acc190已有付款資料但未有匯款編號, 因此增加a1908 is null的判斷
      strSql = "select a1505,sum(axf04) as Namount from acc151, acc150, patent, acc190 where a1902(+) = axf01 and axf01 = a1501 and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1512 is null and a1908 is null group by a1505"
      strSql = strSql & " union select a1605 as a1505,sum(axg04 * (-1)) as Namount from acc161, acc160, patent where axg01 = a1601 and substr(axg03, 1, length(axg03) - 9) = pa01 and substr(axg03, length(axg03) - 8, 6) = pa02 and substr(axg03, length(axg03) - 2, 1) = pa03 and substr(axg03, length(axg03) - 1, 2) = pa04 and a1607 is null" & strWhere(5) & strWhere(4) & " group by a1605"
      If Text5 <> "" Then
         strWhere(4) = " and tm12 = '" & Text5 & "'"
      End If
      strSql = strSql & " union select a1505,sum(axf04) as Namount from acc151, acc150, trademark, acc190 where a1902(+) = axf01 and axf01 = a1501 and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1512 is null and a1908 is null group by a1505"
      strSql = strSql & " union select a1605 as a1505,sum(axg04 * (-1)) as Namount from acc161, acc160, trademark where axg01 = a1601 and substr(axg03, 1, length(axg03) - 9) = tm01 and substr(axg03, length(axg03) - 8, 6) = tm02 and substr(axg03, length(axg03) - 2, 1) = tm03 and substr(axg03, length(axg03) - 1, 2) = tm04 and a1607 is null" & strWhere(5) & strWhere(4) & " group by a1605"
      If Text5 <> "" Then
         strWhere(4) = " and sp11 = '" & Text5 & "'"
      End If
      strSql = strSql & " union select a1505,sum(axf04) as Namount from acc151, acc150, servicepractice, acc190 where a1902(+) = axf01 and axf01 = a1501 and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1512 is null and a1908 is null group by a1505"
      strSql = strSql & " union select a1605 as a1505,sum(axg04 * (-1)) as Namount from acc161, acc160, servicepractice where axg01 = a1601 and substr(axg03, 1, length(axg03) - 9) = sp01 and substr(axg03, length(axg03) - 8, 6) = sp02 and substr(axg03, length(axg03) - 2, 1) = sp03 and substr(axg03, length(axg03) - 1, 2) = sp04 and a1607 is null" & strWhere(5) & strWhere(4) & " group by a1605"
      'Modified by Lydia 2018/02/13 法務無申請案號
      'If Text5 = "" Then
      '   strWhere(4) = ""
     '    strSql = strSql & " union select a1505,sum(axf04) as Namount from acc151, acc150, lawcase, acc190 where a1902(+) = axf01 and axf01 = a1501 and substr(axf03, 1, length(axf03) - 9) = lc01 and substr(axf03, length(axf03) - 8, 6) = lc02 and substr(axf03, length(axf03) - 2, 1) = lc03 and substr(axf03, length(axf03) - 1, 2) = lc04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1520 = 0 or a1520 is null) and a1512 is null and a1908 is null group by a1505"
    '     strSql = strSql & " union select a1605 as a1505,sum(axg04 * (-1)) as Namount from acc161, acc160, LAWCASE where axg01 = a1601 and substr(axg03, 1, length(axg03) - 9) = lc01 and substr(axg03, length(axg03) - 8, 6) = lc02 and substr(axg03, length(axg03) - 2, 1) = lc03 and substr(axg03, length(axg03) - 1, 2) = lc04 and a1607 is null" & strWhere(5) & " group by a1605"
      'End If
       If Trim(strWhere(2)) <> "" Then strSql = strSql & " union select a1505,sum(axf04) as Namount from acc151, acc150, lawcase, acc190 where a1902(+) = axf01 and axf01 = a1501 and substr(axf03, 1, length(axf03) - 9) = lc01 and substr(axf03, length(axf03) - 8, 6) = lc02 and substr(axf03, length(axf03) - 2, 1) = lc03 and substr(axf03, length(axf03) - 1, 2) = lc04 and (a1507 is null or a1507 = 0)" & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1512 is null and a1908 is null group by a1505"
       If Trim(strWhere(5)) <> "" Then strSql = strSql & " union select a1605 as a1505,sum(axg04 * (-1)) as Namount from acc161, acc160, LAWCASE where axg01 = a1601 and substr(axg03, 1, length(axg03) - 9) = lc01 and substr(axg03, length(axg03) - 8, 6) = lc02 and substr(axg03, length(axg03) - 2, 1) = lc03 and substr(axg03, length(axg03) - 1, 2) = lc04 and a1607 is null" & strWhere(5) & " group by a1605"
      'end 2018/02/13
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
      strSql = ""
      adoaccsum.CursorLocation = adUseClient
      If Text5 <> "" Then
         strWhere(4) = " and pa11 = '" & Text5 & "'"
      End If
      'Modify By Sindy 2012/9/11 查CF往來或往來時 : 抓ACC190(W單號), A1908 IS NULL的資料不要抓出來,例 P-102095
      strSql = "select sum(axf04*a1906) as Namount from acc151, acc150, acc190, patent where axf01 = a1501 and a1501 = a1902 and A1908 IS not NULL and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
      'add by sonia 2021/4/9 加抵帳已付
      strSql = strSql & " union select sum(axf04*nvl(a1g03,0)) as Namount from acc151, acc150, acc1g0, patent where axf01 = a1501 and a1512 = a1g01 and a1g01 is not null and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
      'end 2021/4/9
      If Text5 <> "" Then
         strWhere(4) = " and tm12 = '" & Text5 & "'"
      End If
      strSql = strSql & " union select sum(axf04*a1906) as Namount from acc151, acc150, acc190, trademark where axf01 = a1501 and a1501 = a1902 and A1908 IS not NULL and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
      'add by sonia 2021/4/9 加抵帳已付
      strSql = strSql & " union select sum(axf04*nvl(a1g03,0)) as Namount from acc151, acc150, acc1g0, trademark where axf01 = a1501 and a1512 = a1g01 and a1g01 is not null and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
      'end 2021/4/9
      If Text5 <> "" Then
         strWhere(4) = " and sp11 = '" & Text5 & "'"
      End If
      strSql = strSql & " union select sum(axf04*a1906) as Namount from acc151, acc150, acc190, servicepractice where axf01 = a1501 and a1501 = a1902 and A1908 IS not NULL and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
      'add by sonia 2021/4/9 加抵帳已付
      strSql = strSql & " union select sum(axf04*nvl(a1g03,0)) as Namount from acc151, acc150, acc1g0, servicepractice where axf01 = a1501 and a1512 = a1g01 and a1g01 is not null and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
      'end 2021/4/9
      'Modified by Lydia 2018/02/13 法務無申請案號
      'If Text5 = "" Then
      '   strWhere(4) = ""
      '   strSql = strSql & " union select sum(axf04*a1906) as Namount from acc151, acc150, acc190, lawcase where axf01 = a1501 and a1501 = a1902 and A1908 IS not NULL and substr(axf03, 1, length(axf03) - 9) = lc01 and substr(axf03, length(axf03) - 8, 6) = lc02 and substr(axf03, length(axf03) - 2, 1) = lc03 and substr(axf03, length(axf03) - 1, 2) = lc04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
      'End If
      If Trim(strWhere(2)) <> "" Then strSql = strSql & " union select sum(axf04*a1906) as Namount from acc151, acc150, acc190, lawcase where axf01 = a1501 and a1501 = a1902 and A1908 IS not NULL and substr(axf03, 1, length(axf03) - 9) = lc01 and substr(axf03, length(axf03) - 8, 6) = lc02 and substr(axf03, length(axf03) - 2, 1) = lc03 and substr(axf03, length(axf03) - 1, 2) = lc04 and (a1507 is null or a1507 = 0)" & strWhere(2)
      'end 2018/02/13
      'add by sonia 2021/4/9 加抵帳已付
      If Trim(strWhere(2)) <> "" Then strSql = strSql & " union select sum(axf04*nvl(a1g03,0)) as Namount from acc151, acc150, acc1g0, lawcase where axf01 = a1501 and a1512 = a1g01 and a1g01 is not null and substr(axf03, 1, length(axf03) - 9) = lc01 and substr(axf03, length(axf03) - 8, 6) = lc02 and substr(axf03, length(axf03) - 2, 1) = lc03 and substr(axf03, length(axf03) - 1, 2) = lc04 and (a1507 is null or a1507 = 0)" & strWhere(2)
      'end 2021/4/9
      'add by sonia 2021/5/24 加抵帳單已付P-097213之V10300025及抵帳單已抵帳CFP-025854之之V10600009,算合計不需要串基本檔
      If Trim(strWhere(2)) <> "" Then strSql = strSql & " union select sum(axg04*a1906) * (-1) as Namount from acc161, acc160, acc190 where axg01 = a1601 and a1601 = a1902 and a1908 is not null and a1607 is not null " & strWhere(5) & strWhere(4)
      If Trim(strWhere(2)) <> "" Then strSql = strSql & " union select sum(axg04*nvl(a1g03,0)) * (-1) as Namount from acc161, acc160,acc190,acc1i0 c,acc1i0 d,acc1g0 where axg01 = a1601 and a1607=c.a1i03(+) and a1605=c.a1i05(+) and a1607=d.a1i03 and nvl(c.a1i01,d.a1i01)=a1g01(+) and a1g01 is not null and a1607 is not null and a1601=a1902(+) and a1901 is null " & strWhere(5) & strWhere(4)
      'end 2021/5/24
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
   End If
   
   '依查詢資料顯示各欄位值
   Select Case Text9
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
         If Text9 = "2" Then '2.FC未收
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
         If Text9 = "4" Then '4.CF未付
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
         If Text9 = "6" Then '未收未付
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
'      (Text9 = "1" Or Text9 = "2" Or Text9 = "5" Or Text9 = "6" Or Trim(Text9) = "") Then
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
''               If Text9 = "1" Or Text9 = "5" Or Trim(Text9) = "" Then '直接加總
'                  If CurrencyType(intIndex) = "USD" Then
'                     dbl_Famount(intIndex) = dbl_Famount(intIndex) + Val(.Fields("Famount"))
'                  Else
'                     dblAmount = Format(((Val(.Fields("Namount")) * 100 * 100) \ (dblRate * 100)) / 100, FAmount)
'                     .Fields("Famount").Value = dblAmount
'                     dbl_Famount(intIndex) = dbl_Famount(intIndex) + dblAmount
'                  End If
''               End If
''               If Text9 = "2" Or Text9 = "6" Then '2.FC未收 6.未收未付   需計算
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
'               If Text9 = "1" Or Text9 = "2" Then
'                  If Not IsNull(.Fields("a1k25")) And .Fields("a1k25") <> "" Then GoTo ReadNext
'               End If
'               'Modify By Sindy 2012/8/14 Mark
''               If Text9 = "1" Or Text9 = "5" Or Trim(Text9) = "" Then '1.FC往來 5.往來   需計算
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
''               If Text9 = "2" Or Text9 = "6" Then '直接加總
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
'   strSql = ""
'   Select Case Text9
'      Case "1", "2"
'         adoaccsum.CursorLocation = adUseClient
'         If Text5 <> "" Then
'            strWhere(4) = " and pa11 = '" & Text5 & "'"
'         End If
'         'Modify by Morgan 2005/3/8 扣除銷帳
'         'strSQL = "select (a1k11 - nvl(a1k06, 0) * a1k10) as Namount from acc1k0, fagent, nation, patent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4)
'         strSql = "select (a1k11 - nvl(a1k06, 0) * a1k10) as Namount from acc1k0, fagent, nation, patent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4)
'
'         If Text5 <> "" Then
'            strWhere(4) = " and tm12 = '" & Text5 & "'"
'         End If
'         'Modify by Morgan 2005/3/8 扣除銷帳
'         'strSQL = strSQL & " union select (a1k11 - nvl(a1k06, 0) * a1k10) as Namount from acc1k0, fagent, nation, trademark where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4)
'         strSql = strSql & " union select (a1k11 - nvl(a1k06, 0) * a1k10) as Namount from acc1k0, fagent, nation, trademark where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4)
'
'         If Text5 <> "" Then
'            strWhere(4) = " and sp11 = '" & Text5 & "'"
'         End If
'         'Modify by Morgan 2005/3/8 扣除銷帳
'         'strSQL = strSQL & " union select (a1k11 - nvl(a1k06, 0) * a1k10) as Namount from acc1k0, fagent, nation, servicepractice where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4)
'         'strSQL = strSQL & " union select (a1k11 - nvl(a1k06, 0) * a1k10) as Namount from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0)
'         strSql = strSql & " union select (a1k11 - nvl(a1k06, 0) * a1k10) as Namount from acc1k0, fagent, nation, servicepractice where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4)
'         strSql = strSql & " union select (a1k11 - nvl(a1k06, 0) * a1k10) as Namount from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0)
'         '台幣合計
'         adoaccsum.Open "select sum(Namount) from (" & strSql & ") New", adoTaie, adOpenStatic, adLockReadOnly
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
'         strSql = ""
'         adoaccsum.CursorLocation = adUseClient
'         If Text5 <> "" Then
'            strWhere(4) = " and pa11 = '" & Text5 & "'"
'         End If
'
'         'Modify by Morgan 2005/3/8 扣除銷帳
'         'strSQL = "select (a1k08 - nvl(a1k06, 0)) as Namount from acc1k0, fagent, nation, patent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4)
'         strSql = "select (a1k08 - nvl(a1k06, 0)) as Namount from acc1k0, fagent, nation, patent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4)
'
'         If Text5 <> "" Then
'            strWhere(4) = " and tm12 = '" & Text5 & "'"
'         End If
'         'Modify by Morgan 2005/3/8 扣除銷帳
'         'strSQL = strSQL & " union select (a1k08 - nvl(a1k06, 0)) as Namount from acc1k0, fagent, nation, trademark where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4)
'         strSql = strSql & " union select (a1k08 - nvl(a1k06, 0)) as Namount from acc1k0, fagent, nation, trademark where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4)
'
'         If Text5 <> "" Then
'            strWhere(4) = " and sp11 = '" & Text5 & "'"
'         End If
'         'Modify by Morgan 2005/3/8 扣除銷帳
'         'strSQL = strSQL & " union select (a1k08 - nvl(a1k06, 0)) as Namount from acc1k0, fagent, nation, servicepractice where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4)
'         'strSQL = strSQL & " union select (a1k08 - nvl(a1k06, 0)) as Namount from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0)
'         strSql = strSql & " union select (a1k08 - nvl(a1k06, 0)) as Namount from acc1k0, fagent, nation, servicepractice where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4)
'         strSql = strSql & " union select (a1k08 - nvl(a1k06, 0)) as Namount from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0)
'         '外幣合計
'         adoaccsum.Open "select sum(Namount) from (" & strSql & ") New", adoTaie, adOpenStatic, adLockReadOnly
'         If adoaccsum.RecordCount <> 0 Then
'            If IsNull(adoaccsum.Fields(0).Value) Then
'               Text10 = MsgText(601)
'            Else
'               Text10 = Format(adoaccsum.Fields(0).Value, FDollar)
'            End If
'         Else
'            Text10 = MsgText(601)
'         End If
'         adoaccsum.Close
'         strSql = ""
'         adoaccsum.CursorLocation = adUseClient
'         If Text5 <> "" Then
'            strWhere(4) = " and pa11 = '" & Text5 & "'"
'         End If
'         'Modify by Morgan 2005/3/8 扣除銷帳
'         'strSQL = "select (a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)) as Namount from acc1k0, fagent, nation, patent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         '2009/4/24 modify by sonia 改同grid
'         'strSQL = "select (a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)) as Namount from acc1k0, fagent, nation, patent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         'Add By Sindy 2010/8/31 增加,nvl(a1k09,0) as Lawfee
'         strSql = "select a1k01,decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount,nvl(a1k09,0) as Lawfee from acc1k0, fagent, nation, patent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'
'         If Text5 <> "" Then
'            strWhere(4) = " and tm12 = '" & Text5 & "'"
'         End If
'         'Modify by Morgan 2005/3/8 扣除銷帳
'         'strSQL = strSQL & " union select (a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)) as Namount from acc1k0, fagent, nation, trademark where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         '2009/4/24 modify by sonia 改同grid
'         'strSQL = strSQL & " union select (a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)) as Namount from acc1k0, fagent, nation, trademark where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         strSql = strSql & " union select a1k01,decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount,nvl(a1k09,0) as Lawfee from acc1k0, fagent, nation, trademark where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'
'         If Text5 <> "" Then
'            strWhere(4) = " and sp11 = '" & Text5 & "'"
'         End If
'         'Modify by Morgan 2005/3/8 扣除銷帳
'         'strSQL = strSQL & " union select (a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)) as Namount from acc1k0, fagent, nation, servicepractice where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         'strSQL = strSQL & " union select (a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)) as Namount from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
'         '2009/4/24 modify by sonia 改同grid
'         'strSQL = strSQL & " union select (a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)) as Namount from acc1k0, fagent, nation, servicepractice where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         'strSQL = strSQL & " union select (a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)) as Namount from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
'         strSql = strSql & " union select a1k01,decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount,nvl(a1k09,0) as Lawfee from acc1k0, fagent, nation, servicepractice where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         strSql = strSql & " union select a1k01,decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount,nvl(a1k09,0) as Lawfee from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
'         '2009/4/24 end
'         '台幣未收
'         adoaccsum.Open "select sum(Namount),sum(Lawfee) from (" & strSql & ") New", adoTaie, adOpenStatic, adLockReadOnly
'         If adoaccsum.RecordCount <> 0 Then
'            If IsNull(adoaccsum.Fields(0).Value) Then
'               Text11 = MsgText(601)
'            Else
'               Text11 = Format(adoaccsum.Fields(0).Value, FDollar)
'            End If
'            'Add By Sindy 2010/8/31 未收規費
'            If IsNull(adoaccsum.Fields(1).Value) Then
'               Text17 = MsgText(601)
'            Else
'               Text17 = Format(adoaccsum.Fields(1).Value, FDollar)
'            End If
'         Else
'            Text11 = MsgText(601)
'            Text17 = MsgText(601) 'Add By Sindy 2010/8/31 未收規費
'         End If
'         adoaccsum.Close
'         strSql = ""
'         adoaccsum.CursorLocation = adUseClient
'         If Text5 <> "" Then
'            strWhere(4) = " and pa11 = '" & Text5 & "'"
'         End If
'
'         'Modify by Morgan 2005/3/8 扣除銷帳
'         'strSQL = "select (a1k08 - nvl(a1k06, 0) - nvl(a1k30, 0) / decode(a1k10, 0, 1, nvl(a1k10, 1))) as Namount from acc1k0, fagent, nation, patent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         '2009/4/24 modify by sonia 改同grid
'         'strSQL = "select (a1k08 - nvl(a1k06, 0) - nvl(a1k30, 0) / decode(a1k10, 0, 1, nvl(a1k10, 1))) as Namount from acc1k0, fagent, nation, patent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         strSql = "select decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Namount from acc1k0, fagent, nation, patent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'
'         If Text5 <> "" Then
'            strWhere(4) = " and tm12 = '" & Text5 & "'"
'         End If
'         'Modify by Morgan 2005/3/8 扣除銷帳
'         'strSQL = strSQL & " union select (a1k08 - nvl(a1k06, 0) - nvl(a1k30, 0) / decode(a1k10, 0, 1, nvl(a1k10, 1))) as Namount from acc1k0, fagent, nation, trademark where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         '2009/4/24 modify by sonia 改同grid
'         'strSQL = strSQL & " union select (a1k08 - nvl(a1k06, 0) - nvl(a1k30, 0) / decode(a1k10, 0, 1, nvl(a1k10, 1))) as Namount from acc1k0, fagent, nation, trademark where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         strSql = strSql & " union select decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Namount from acc1k0, fagent, nation, trademark where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'
'         If Text5 <> "" Then
'            strWhere(4) = " and sp11 = '" & Text5 & "'"
'         End If
'         'Modify by Morgan 2005/3/8 扣除銷帳
'         'strSQL = strSQL & " union select (a1k08 - nvl(a1k06, 0) - nvl(a1k30, 0) / decode(a1k10, 0, 1, nvl(a1k10, 1))) as Namount from acc1k0, fagent, nation, servicepractice where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         'strSQL = strSQL & " union select (a1k08 - nvl(a1k06, 0) - nvl(a1k30, 0) / decode(a1k10, 0, 1, nvl(a1k10, 1))) as Namount from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
'         '2009/4/24 modify by sonia 改同grid
'         'strSQL = strSQL & " union select (a1k08 - nvl(a1k06, 0) - nvl(a1k30, 0) / decode(a1k10, 0, 1, nvl(a1k10, 1))) as Namount from acc1k0, fagent, nation, servicepractice where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         'strSQL = strSQL & " union select (a1k08 - nvl(a1k06, 0) - nvl(a1k30, 0) / decode(a1k10, 0, 1, nvl(a1k10, 1))) as Namount from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
'         strSql = strSql & " union select decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Namount from acc1k0, fagent, nation, servicepractice where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         strSql = strSql & " union select decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Namount from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
'         '2009/4/24 end
'         '外幣未收
'         adoaccsum.Open "select sum(Namount) from (" & strSql & ") New", adoTaie, adOpenStatic, adLockReadOnly
'         If adoaccsum.RecordCount <> 0 Then
'            If IsNull(adoaccsum.Fields(0).Value) Then
'               Text15 = MsgText(601)
'            Else
'               Text15 = Format(adoaccsum.Fields(0).Value, FDollar)
'            End If
'         Else
'            Text15 = MsgText(601)
'         End If
'         adoaccsum.Close
'         Text12 = ""
'         Text14 = ""
'         Text16 = ""
'         If Text9 = "2" Then
'            Combo2.Clear
'            Text13 = ""
'         End If
'      Case "3", "4"
'         If Text7 = "" Then
'            adoaccsum.CursorLocation = adUseClient
'            If Text5 <> "" Then
'               strWhere(4) = " and pa11 = '" & Text5 & "'"
'            End If
'            '2011/9/6 modify by sonia台幣金額由sum(a1905)改為sum(axf04*a1906)
'            'strSql = "select sum(a1905) as Namount from acc151, acc150, acc190, fagent, nation, patent where axf01 = a1501 and a1501 = a1902 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            strSql = "select sum(axf04*a1906) as Namount from acc151, acc150, acc190, fagent, nation, patent where axf01 = a1501 and a1501 = a1902 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            If Text5 <> "" Then
'               strWhere(4) = " and tm12 = '" & Text5 & "'"
'            End If
'            '2011/9/6 modify by sonia台幣金額由sum(a1905)改為sum(axf04*a1906)
'            'strSql = strSql & " union select sum(a1905) as Namount from acc151, acc150, acc190, fagent, nation, trademark where axf01 = a1501 and a1501 = a1902 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            strSql = strSql & " union select sum(axf04*a1906) as Namount from acc151, acc150, acc190, fagent, nation, trademark where axf01 = a1501 and a1501 = a1902 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            If Text5 <> "" Then
'               strWhere(4) = " and sp11 = '" & Text5 & "'"
'            End If
'            '2011/9/6 modify by sonia台幣金額由sum(a1905)改為sum(axf04*a1906)
'            'strSql = strSql & " union select sum(a1905) as Namount from acc151, acc150, acc190, fagent, nation, servicepractice where axf01 = a1501 and a1501 = a1902 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            strSql = strSql & " union select sum(axf04*a1906) as Namount from acc151, acc150, acc190, fagent, nation, servicepractice where axf01 = a1501 and a1501 = a1902 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            '2011/9/6 add by sonia 加入法務案件
'            If Text5 = "" Then
'               strWhere(4) = ""
'               strSql = strSql & " union select sum(axf04*a1906) as Namount from acc151, acc150, acc190, fagent, nation, lawcase where axf01 = a1501 and a1501 = a1902 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = lc01 and substr(axf03, length(axf03) - 8, 6) = lc02 and substr(axf03, length(axf03) - 2, 1) = lc03 and substr(axf03, length(axf03) - 1, 2) = lc04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            End If
'            '2011/9/6 end
'            adoaccsum.Open "select sum(Namount) from (" & strSql & ") New", adoTaie, adOpenStatic, adLockReadOnly
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
'
'            strSql = ""
'            adoaccsum.CursorLocation = adUseClient
'            If Text5 <> "" Then
'               strWhere(4) = " and pa11 = '" & Text5 & "'"
'            End If
'            'Modify By Sindy 2010/7/13 CF 各幣別分開顯示
'            'strSql = "select sum(axf04) as Namount from acc151, acc150, fagent, nation, patent where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            strSql = "select a1505,sum(axf04) as Namount from acc151, acc150, fagent, nation, patent where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " group by a1505 "
'            If Text5 <> "" Then
'               strWhere(4) = " and tm12 = '" & Text5 & "'"
'            End If
'            'strSql = strSql & " union select sum(axf04) as Namount from acc151, acc150, fagent, nation, trademark where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            strSql = strSql & " union select a1505,sum(axf04) as Namount from acc151, acc150, fagent, nation, trademark where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " group by a1505 "
'            If Text5 <> "" Then
'               strWhere(4) = " and sp11 = '" & Text5 & "'"
'            End If
'            'strSql = strSql & " union select sum(axf04) as Namount from acc151, acc150, fagent, nation, servicepractice where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            strSql = strSql & " union select a1505,sum(axf04) as Namount from acc151, acc150, fagent, nation, servicepractice where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " group by a1505 "
'            'adoaccsum.Open "select sum(Namount) from (" & strSql & ") New", adoTaie, adOpenStatic, adLockReadOnly
'            adoaccsum.Open "select a1505,sum(Namount) from (" & strSql & ") group by a1505 order by a1505", adoTaie, adOpenStatic, adLockReadOnly
'            If adoaccsum.RecordCount <> 0 Then
'               adoaccsum.MoveFirst
'               Do While Not adoaccsum.EOF
'                  If Val(" " & adoaccsum.Fields(1)) <> 0 Then
'                     'Modify By Sindy 2012/8/15
'                     dblSumA1606 = Val(GetACC160Amt(adoaccsum.Fields(0), ""))
'                     Combo4.AddItem adoaccsum.Fields(0) & " " & (Val(adoaccsum.Fields(1)) - dblSumA1606)
'                     '2012/8/15 End
'                     Combo4.ListIndex = 0
'                  End If
'                  adoaccsum.MoveNext
'               Loop
''               If IsNull(adoaccsum.Fields(0).Value) Then
''                  Text14 = MsgText(601)
''               Else
''                  Text14 = Format(adoaccsum.Fields(0).Value, FDollar)
''               End If
''            Else
''               Text14 = MsgText(601)
'            End If
'            adoaccsum.Close
'            strSql = ""
'            adoaccsum.CursorLocation = adUseClient
'            If Text5 <> "" Then
'               strWhere(4) = " and pa11 = '" & Text5 & "'"
'            End If
'            'Modify By Sindy 2010/7/13 未付 各幣別分開顯示
'            'strSql = "select sum(axf04) as Namount from acc151, acc150, fagent, nation, patent where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null)"
'            'strSql = strSql & " union select sum(a1904 * (-1)) as Namount from acc190, acc151, acc150, fagent, nation, patent where a1902 = axf01 and axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null)"
'            strSql = "select a1505,sum(axf04) as Namount from acc151, acc150, fagent, nation, patent where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null) group by a1505"
'            strSql = strSql & " union select a1505,sum(a1904 * (-1)) as Namount from acc190, acc151, acc150, fagent, nation, patent where a1902 = axf01 and axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null) group by a1505"
'            If Text5 <> "" Then
'               strWhere(4) = " and tm12 = '" & Text5 & "'"
'            End If
'            'strSql = strSql & " union select sum(axf04) as Namount from acc151, acc150, fagent, nation, trademark where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null)"
'            'strSql = strSql & " union select sum(a1904 * (-1)) as Namount from acc190, acc151, acc150, fagent, nation, trademark where a1902 = axf01 and axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null)"
'            strSql = strSql & " union select a1505,sum(axf04) as Namount from acc151, acc150, fagent, nation, trademark where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null) group by a1505"
'            strSql = strSql & " union select a1505,sum(a1904 * (-1)) as Namount from acc190, acc151, acc150, fagent, nation, trademark where a1902 = axf01 and axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null) group by a1505"
'            If Text5 <> "" Then
'               strWhere(4) = " and sp11 = '" & Text5 & "'"
'            End If
'            'strSql = strSql & " union select sum(axf04) as Namount from acc151, acc150, fagent, nation, servicepractice where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null)"
'            'strSql = strSql & " union select sum(a1904 * (-1)) as Namount from acc190, acc151, acc150, fagent, nation, servicepractice where a1902 = axf01 and axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null)"
'            strSql = strSql & " union select a1505,sum(axf04) as Namount from acc151, acc150, fagent, nation, servicepractice where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null) group by a1505"
'            strSql = strSql & " union select a1505,sum(a1904 * (-1)) as Namount from acc190, acc151, acc150, fagent, nation, servicepractice where a1902 = axf01 and axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null) group by a1505"
'            'adoaccsum.Open "select sum(Namount) from (" & strSql & ") New", adoTaie, adOpenStatic, adLockReadOnly
'            adoaccsum.Open "select a1505,sum(Namount) from (" & strSql & ") group by a1505 order by a1505", adoTaie, adOpenStatic, adLockReadOnly
'            If adoaccsum.RecordCount <> 0 Then
'               adoaccsum.MoveFirst
'               Do While Not adoaccsum.EOF
'                  If Val(" " & adoaccsum.Fields(1)) <> 0 Then
'                     'Modify By Sindy 2012/8/15
'                     dblSumA1606 = Val(GetACC160Amt(adoaccsum.Fields(0), "0"))
'                     Combo5.AddItem adoaccsum.Fields(0) & " " & (Val(adoaccsum.Fields(1)) - dblSumA1606)
'                     '2012/8/15 End
'                     Combo5.ListIndex = 0
'                  End If
'                  adoaccsum.MoveNext
'               Loop
''               If IsNull(adoaccsum.Fields(0).Value) Then
''                  Text16 = MsgText(601)
''               Else
''                  If adoaccsum.Fields(0).Value < 0 Then
''                     Text16 = MsgText(601)
''                  Else
''                     Text16 = Format(adoaccsum.Fields(0).Value, FDollar)
''                  End If
''               End If
''            Else
''               Text16 = MsgText(601)
'            'Add By Sindy 2012/8/15 檢查是否有未抵帳的抵帳單
'            Else
'               Call GetACC160Amt("", "0")
'            '2012/8/15 End
'            End If
'            adoaccsum.Close
'            Text13 = ""
'            Text11 = ""
'            Text17 = MsgText(601) 'Add By Sindy 2010/8/31 未收規費
'            Text10 = ""
'            Text15 = ""
'            If Text9 = "4" Then
'               Combo4.Clear
'               Text12 = ""
'            End If
'         End If
'      Case "", "5", "6"
'         'edit by nickc 2007/02/08
'         'stSQL = ""
'         strSql = ""
'         adoaccsum.CursorLocation = adUseClient
'         If Text5 <> "" Then
'            strWhere(4) = " and pa11 = '" & Text5 & "'"
'         End If
'         strSql = "select (a1k11 - nvl(a1k06, 0) * a1k10) as Namount from acc1k0, fagent, nation, patent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4)
'         If Text5 <> "" Then
'            strWhere(4) = " and tm12 = '" & Text5 & "'"
'         End If
'         strSql = strSql & " union select (a1k11 - nvl(a1k06, 0) * a1k10) as Namount from acc1k0, fagent, nation, trademark where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4)
'         If Text5 <> "" Then
'            strWhere(4) = " and sp11 = '" & Text5 & "'"
'         End If
'         strSql = strSql & " union select (a1k11 - nvl(a1k06, 0) * a1k10) as Namount from acc1k0, fagent, nation, servicepractice where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4)
'         strSql = strSql & " union select (a1k11 - nvl(a1k06, 0) * a1k10) as Namount from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0)
'         adoaccsum.Open "select sum(Namount) from (" & strSql & ") New", adoTaie, adOpenStatic, adLockReadOnly
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
'         strSql = ""
'         adoaccsum.CursorLocation = adUseClient
'         If Text5 <> "" Then
'            strWhere(4) = " and pa11 = '" & Text5 & "'"
'         End If
'         '2009/4/24 modify by sonia 改同grid
'         'strSQL = "select (a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)) as Namount from acc1k0, fagent, nation, patent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         'Add By Sindy 2010/8/31 增加,nvl(a1k09,0) as Lawfee
'         strSql = "select a1k01,decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount,nvl(a1k09,0) as Lawfee from acc1k0, fagent, nation, patent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         If Text5 <> "" Then
'            strWhere(4) = " and tm12 = '" & Text5 & "'"
'         End If
'         '2009/4/24 modify by sonia 改同grid
'         'strSQL = strSQL & " union select (a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)) as Namount from acc1k0, fagent, nation, trademark where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         strSql = strSql & " union select a1k01,decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount,nvl(a1k09,0) as Lawfee from acc1k0, fagent, nation, trademark where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         If Text5 <> "" Then
'            strWhere(4) = " and sp11 = '" & Text5 & "'"
'         End If
'         '2009/4/24 modify by sonia 改同grid
'         'strSQL = strSQL & " union select (a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)) as Namount from acc1k0, fagent, nation, servicepractice where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         'strSQL = strSQL & " union select (a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)) as Namount from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
'         strSql = strSql & " union select a1k01,decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount,nvl(a1k09,0) as Lawfee from acc1k0, fagent, nation, servicepractice where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         strSql = strSql & " union select a1k01,decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0) * a1k10),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)),0)) as Namount,nvl(a1k09,0) as Lawfee from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
'         '台幣未收
'         adoaccsum.Open "select sum(Namount),sum(Lawfee) from (" & strSql & ") New", adoTaie, adOpenStatic, adLockReadOnly
'         If adoaccsum.RecordCount <> 0 Then
'            If IsNull(adoaccsum.Fields(0).Value) Then
'               Text11 = MsgText(601)
'            Else
'               Text11 = Format(adoaccsum.Fields(0).Value, FDollar)
'            End If
'            'Add By Sindy 2010/8/31 未收規費
'            If IsNull(adoaccsum.Fields(1).Value) Then
'               Text17 = MsgText(601)
'            Else
'               Text17 = Format(adoaccsum.Fields(1).Value, FDollar)
'            End If
'         Else
'            Text11 = MsgText(601)
'            Text17 = MsgText(601) 'Add By Sindy 2010/8/31 未收規費
'         End If
'         adoaccsum.Close
'
'         strSql = ""
'         If Text7 = "" Then
'            adoaccsum.CursorLocation = adUseClient
'            If Text5 <> "" Then
'               strWhere(4) = " and pa11 = '" & Text5 & "'"
'            End If
'            '2011/9/6 modify by sonia台幣金額由sum(a1905)改為sum(axf04*a1906)
'            'strSql = "select sum(a1905) as Namount from acc151, acc150, acc190, fagent, nation, patent where axf01 = a1501 and a1501 = a1902 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            strSql = "select sum(axf04*a1906) as Namount from acc151, acc150, acc190, fagent, nation, patent where axf01 = a1501 and a1501 = a1902 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            If Text5 <> "" Then
'               strWhere(4) = " and tm12 = '" & Text5 & "'"
'            End If
'            '2011/9/6 modify by sonia台幣金額由sum(a1905)改為sum(axf04*a1906)
'            'strSql = strSql & " union select sum(a1905) as Namount from acc151, acc150, acc190, fagent, nation, trademark where axf01 = a1501 and a1501 = a1902 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            strSql = strSql & " union select sum(axf04*a1906) as Namount from acc151, acc150, acc190, fagent, nation, trademark where axf01 = a1501 and a1501 = a1902 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            If Text5 <> "" Then
'               strWhere(4) = " and sp11 = '" & Text5 & "'"
'            End If
'            '2011/9/6 modify by sonia台幣金額由sum(a1905)改為sum(axf04*a1906)
'            'strSql = strSql & " union select sum(a1905) as Namount from acc151, acc150, acc190, fagent, nation, servicepractice where axf01 = a1501 and a1501 = a1902 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            strSql = strSql & " union select sum(axf04*a1906) as Namount from acc151, acc150, acc190, fagent, nation, servicepractice where axf01 = a1501 and a1501 = a1902 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            '2011/9/6 add by sonia 加入法務案件
'            If Text5 = "" Then
'               strWhere(4) = ""
'               strSql = strSql & " union select sum(axf04*a1906) as Namount from acc151, acc150, acc190, fagent, nation, lawcase where axf01 = a1501 and a1501 = a1902 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = lc01 and substr(axf03, length(axf03) - 8, 6) = lc02 and substr(axf03, length(axf03) - 2, 1) = lc03 and substr(axf03, length(axf03) - 1, 2) = lc04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            End If
'            '2011/9/6 end
'            adoaccsum.Open "select sum(Namount) from (" & strSql & ") New", adoTaie, adOpenStatic, adLockReadOnly
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
'         End If
'
'         strSql = ""
'         adoaccsum.CursorLocation = adUseClient
'         If Text5 <> "" Then
'            strWhere(4) = " and pa11 = '" & Text5 & "'"
'         End If
'         strSql = "select (a1k08 - nvl(a1k06, 0)) as Namount from acc1k0, fagent, nation, patent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4)
'         If Text5 <> "" Then
'            strWhere(4) = " and tm12 = '" & Text5 & "'"
'         End If
'         strSql = strSql & " union select (a1k08 - nvl(a1k06, 0)) as Namount from acc1k0, fagent, nation, trademark where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4)
'         If Text5 <> "" Then
'            strWhere(4) = " and sp11 = '" & Text5 & "'"
'         End If
'         strSql = strSql & " union select (a1k08 - nvl(a1k06, 0)) as Namount from acc1k0, fagent, nation, servicepractice where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4)
'         strSql = strSql & " union select (a1k08 - nvl(a1k06, 0)) as Namount from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0)
'         adoaccsum.Open "select sum(Namount) from (" & strSql & ") New", adoTaie, adOpenStatic, adLockReadOnly
'         If adoaccsum.RecordCount <> 0 Then
'            If IsNull(adoaccsum.Fields(0).Value) Then
'               Text10 = MsgText(601)
'            Else
'               Text10 = Format(adoaccsum.Fields(0).Value, FDollar)
'            End If
'         Else
'            Text10 = MsgText(601)
'         End If
'         adoaccsum.Close
'         strSql = ""
'         adoaccsum.CursorLocation = adUseClient
'         If Text5 <> "" Then
'            strWhere(4) = " and pa11 = '" & Text5 & "'"
'         End If
'         '2009/4/24 modify by sonia 改同grid
'         'strSQL = "select (a1k08 - nvl(a1k06, 0) - nvl(a1k30, 0) / decode(a1k10, 0, 1, nvl(a1k10, 1))) as Namount from acc1k0, fagent, nation, patent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         strSql = "select decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Namount from acc1k0, fagent, nation, patent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         If Text5 <> "" Then
'            strWhere(4) = " and tm12 = '" & Text5 & "'"
'         End If
'         '2009/4/24 modify by sonia 改同grid
'         'strSQL = strSQL & " union select (a1k08 - nvl(a1k06, 0) - nvl(a1k30, 0) / decode(a1k10, 0, 1, nvl(a1k10, 1))) as Namount from acc1k0, fagent, nation, trademark where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         strSql = strSql & " union select decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Namount from acc1k0, fagent, nation, trademark where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         If Text5 <> "" Then
'            strWhere(4) = " and sp11 = '" & Text5 & "'"
'         End If
'         '2009/4/24 modify by sonia 改同grid
'         'strSQL = strSQL & " union select (a1k08 - nvl(a1k06, 0) - nvl(a1k30, 0) / decode(a1k10, 0, 1, nvl(a1k10, 1))) as Namount from acc1k0, fagent, nation, servicepractice where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         strSql = strSql & " union select decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0))/a1k10,1)) as Namount from acc1k0, fagent, nation, servicepractice where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         '外幣未收
'         adoaccsum.Open "select sum(Namount) from (" & strSql & ") New", adoTaie, adOpenStatic, adLockReadOnly
'         If adoaccsum.RecordCount <> 0 Then
'            If IsNull(adoaccsum.Fields(0).Value) Then
'               Text15 = MsgText(601)
'            Else
'               Text15 = Format(adoaccsum.Fields(0).Value, FDollar)
'            End If
'         Else
'            Text15 = MsgText(601)
'         End If
'         adoaccsum.Close
'         strSql = ""
'         If Text7 = "" Then
'            adoaccsum.CursorLocation = adUseClient
'            If Text5 <> "" Then
'               strWhere(4) = " and pa11 = '" & Text5 & "'"
'            End If
'            'Modify By Sindy 2010/7/13 CF 各幣別分開顯示
'            'strSql = "select sum(axf04) as Namount from acc151, acc150, fagent, nation, patent where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            strSql = "select a1505,sum(axf04) as Namount from acc151, acc150, fagent, nation, patent where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " group by a1505 "
'            If Text5 <> "" Then
'               strWhere(4) = " and tm12 = '" & Text5 & "'"
'            End If
'            'strSql = strSql & " union select sum(axf04) as Namount from acc151, acc150, fagent, nation, trademark where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            strSql = strSql & " union select a1505,sum(axf04) as Namount from acc151, acc150, fagent, nation, trademark where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " group by a1505 "
'            If Text5 <> "" Then
'               strWhere(4) = " and sp11 = '" & Text5 & "'"
'            End If
'            'strSql = strSql & " union select sum(axf04) as Namount from acc151, acc150, fagent, nation, servicepractice where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            strSql = strSql & " union select a1505,sum(axf04) as Namount from acc151, acc150, fagent, nation, servicepractice where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " group by a1505 "
'            'adoaccsum.Open "select sum(Namount) from (" & strSql & ") New", adoTaie, adOpenStatic, adLockReadOnly
'            adoaccsum.Open "select a1505,sum(Namount) from (" & strSql & ") group by a1505", adoTaie, adOpenStatic, adLockReadOnly
'            If adoaccsum.RecordCount <> 0 Then
'               adoaccsum.MoveFirst
'               Do While Not adoaccsum.EOF
'                  If Val(" " & adoaccsum.Fields(1)) <> 0 Then
'                     'Modify By Sindy 2012/8/15
'                     dblSumA1606 = Val(GetACC160Amt(adoaccsum.Fields(0), ""))
'                     Combo4.AddItem adoaccsum.Fields(0) & " " & (Val(adoaccsum.Fields(1)) - dblSumA1606)
'                     '2012/8/15 End
'                     Combo4.ListIndex = 0
'                  End If
'                  adoaccsum.MoveNext
'               Loop
''               If IsNull(adoaccsum.Fields(0).Value) Then
''                  Text14 = MsgText(601)
''               Else
''                  Text14 = Format(adoaccsum.Fields(0).Value, FDollar)
''               End If
''            Else
''               Text14 = MsgText(601)
'            End If
'            adoaccsum.Close
'            strSql = ""
'            adoaccsum.CursorLocation = adUseClient
'            If Text5 <> "" Then
'               strWhere(4) = " and pa11 = '" & Text5 & "'"
'            End If
'            'Modify By Sindy 2010/7/13 未付 各幣別分開顯示
'            'strSql = "select sum(axf04) as Namount from acc151, acc150, fagent, nation, patent where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null)"
'            'strSql = strSql & " union select sum(a1904 * (-1)) as Namount from acc190, acc151, acc150, fagent, nation, patent where a1902 = axf01 and axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null)"
'            strSql = "select a1505,sum(axf04) as Namount from acc151, acc150, fagent, nation, patent where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null) group by a1505"
'            strSql = strSql & " union select a1505,sum(a1904 * (-1)) as Namount from acc190, acc151, acc150, fagent, nation, patent where a1902 = axf01 and axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null) group by a1505"
'            If Text5 <> "" Then
'               strWhere(4) = " and tm12 = '" & Text5 & "'"
'            End If
'            'strSql = strSql & " union select sum(axf04) as Namount from acc151, acc150, fagent, nation, trademark where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null)"
'            'strSql = strSql & " union select sum(a1904 * (-1)) as Namount from acc190, acc151, acc150, fagent, nation, trademark where a1902 = axf01 and axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null)"
'            strSql = strSql & " union select a1505,sum(axf04) as Namount from acc151, acc150, fagent, nation, trademark where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null) group by a1505"
'            strSql = strSql & " union select a1505,sum(a1904 * (-1)) as Namount from acc190, acc151, acc150, fagent, nation, trademark where a1902 = axf01 and axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null) group by a1505"
'            If Text5 <> "" Then
'               strWhere(4) = " and sp11 = '" & Text5 & "'"
'            End If
'            'strSql = strSql & " union select sum(axf04) as Namount from acc151, acc150, fagent, nation, servicepractice where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null)"
'            'strSql = strSql & " union select sum(a1904 * (-1)) as Namount from acc190, acc151, acc150, fagent, nation, servicepractice where a1902 = axf01 and axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null)"
'            strSql = strSql & " union select a1505,sum(axf04) as Namount from acc151, acc150, fagent, nation, servicepractice where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null) group by a1505"
'            strSql = strSql & " union select a1505,sum(a1904 * (-1)) as Namount from acc190, acc151, acc150, fagent, nation, servicepractice where a1902 = axf01 and axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null) group by a1505"
'            'adoaccsum.Open "select sum(Namount) from (" & strSql & ") New", adoTaie, adOpenStatic, adLockReadOnly
'            adoaccsum.Open "select a1505,sum(Namount) from (" & strSql & ") group by a1505 order by a1505", adoTaie, adOpenStatic, adLockReadOnly
'            If adoaccsum.RecordCount <> 0 Then
'               adoaccsum.MoveFirst
'               Do While Not adoaccsum.EOF
'                  If Val(" " & adoaccsum.Fields(1)) <> 0 Then
'                     'Modify By Sindy 2012/8/15
'                     dblSumA1606 = Val(GetACC160Amt(adoaccsum.Fields(0), "0"))
'                     Combo5.AddItem adoaccsum.Fields(0) & " " & (Val(adoaccsum.Fields(1)) - dblSumA1606)
'                     '2012/8/15 End
'                     Combo5.ListIndex = 0
'                  End If
'                  adoaccsum.MoveNext
'               Loop
''               If IsNull(adoaccsum.Fields(0).Value) Then
''                  Text16 = MsgText(601)
''               Else
''                  If adoaccsum.Fields(0).Value < 0 Then
''                     Text16 = MsgText(601)
''                  Else
''                     Text16 = Format(adoaccsum.Fields(0).Value, FDollar)
''                  End If
''               End If
''            Else
''               Text16 = MsgText(601)
'            'Add By Sindy 2012/8/15 檢查是否有未抵帳的抵帳單
'            Else
'               Call GetACC160Amt("", "0")
'            '2012/8/15 End
'            End If
'            adoaccsum.Close
'         End If
'         If Text9 = "6" Then
'            Combo2.Clear
'            Text13 = ""
'            Combo4.Clear
'            Text12 = ""
'         End If
'*****************************************************************
'2009/4/28 CANCEL BY SONIA
'      Case Else
'         adoaccsum.CursorLocation = adUseClient
'         If Text5 <> "" Then
'            strWhere(4) = " and pa11 = '" & Text5 & "'"
'         End If
'         strSQL = "select (a1k11 - nvl(a1k06, 0) * a1k10) as Namount from acc1k0, fagent, nation, patent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4)
'         If Text5 <> "" Then
'            strWhere(4) = " and tm12 = '" & Text5 & "'"
'         End If
'         strSQL = strSQL & " union select (a1k11 - nvl(a1k06, 0) * a1k10) as Namount from acc1k0, fagent, nation, trademark where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4)
'         If Text5 <> "" Then
'            strWhere(4) = " and sp11 = '" & Text5 & "'"
'         End If
'         strSQL = strSQL & " union select (a1k11 - nvl(a1k06, 0) * a1k10) as Namount from acc1k0, fagent, nation, servicepractice where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4)
'         strSQL = strSQL & " union select (a1k11 - nvl(a1k06, 0) * a1k10) as Namount from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0)
'         adoaccsum.Open "select sum(Namount) from (" & strSQL & ") New", adoTaie, adOpenStatic, adLockReadOnly
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
'         strSQL = ""
'         adoaccsum.CursorLocation = adUseClient
'         If Text5 <> "" Then
'            strWhere(4) = " and pa11 = '" & Text5 & "'"
'         End If
'         strSQL = "select (a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)) as Namount from acc1k0, fagent, nation, patent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         If Text5 <> "" Then
'            strWhere(4) = " and tm12 = '" & Text5 & "'"
'         End If
'         strSQL = strSQL & " union select (a1k11 - nvl(a1k06, 0) * a1k10 - nvl(a1k30,0)) as Namount from acc1k0, fagent, nation, trademark where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         If Text5 <> "" Then
'            strWhere(4) = " and sp11 = '" & Text5 & "'"
'         End If
'         strSQL = strSQL & " union select (a1k11 - nvl(a1k06, 0) * a1k10- nvl(a1k30,0)) as Namount from acc1k0, fagent, nation, servicepractice where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         strSQL = strSQL & " union select (a1k11 - nvl(a1k06, 0) * a1k10- nvl(a1k30,0)) as Namount from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
'         adoaccsum.Open "select sum(Namount) from (" & strSQL & ") New", adoTaie, adOpenStatic, adLockReadOnly
'         If adoaccsum.RecordCount <> 0 Then
'            If IsNull(adoaccsum.Fields(0).Value) Then
'               Text11 = MsgText(601)
'            Else
'               Text11 = Format(adoaccsum.Fields(0).Value, FDollar)
'            End If
'         Else
'            Text11 = MsgText(601)
'         End If
'         adoaccsum.Close
'         strSQL = ""
'         If Text7 = "" Then
'            adoaccsum.CursorLocation = adUseClient
'            If Text5 <> "" Then
'               strWhere(4) = " and pa11 = '" & Text5 & "'"
'            End If
'            strSQL = "select sum(a1905) as Namount from acc151, acc150, acc190, fagent, nation, patent where axf01 = a1501 and a1501 = a1902 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            If Text5 <> "" Then
'               strWhere(4) = " and tm12 = '" & Text5 & "'"
'            End If
'            strSQL = strSQL & " union select sum(a1905) as Namount from acc151, acc150, acc190, fagent, nation, trademark where axf01 = a1501 and a1501 = a1902 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            If Text5 <> "" Then
'               strWhere(4) = " and sp11 = '" & Text5 & "'"
'            End If
'            strSQL = strSQL & " union select sum(a1905) as Namount from acc151, acc150, acc190, fagent, nation, servicepractice where axf01 = a1501 and a1501 = a1902 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            adoaccsum.Open "select sum(Namount) from (" & strSQL & ") New", adoTaie, adOpenStatic, adLockReadOnly
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
'         End If
'         strSQL = ""
'         adoaccsum.CursorLocation = adUseClient
'         If Text5 <> "" Then
'            strWhere(4) = " and pa11 = '" & Text5 & "'"
'         End If
'         strSQL = "select (a1k08 - nvl(a1k06, 0)) as Namount from acc1k0, fagent, nation, patent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4)
'         If Text5 <> "" Then
'            strWhere(4) = " and tm12 = '" & Text5 & "'"
'         End If
'         strSQL = strSQL & " union select (a1k08 - nvl(a1k06, 0)) as Namount from acc1k0, fagent, nation, trademark where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4)
'         If Text5 <> "" Then
'            strWhere(4) = " and sp11 = '" & Text5 & "'"
'         End If
'         strSQL = strSQL & " union select (a1k08 - nvl(a1k06, 0)) as Namount from acc1k0, fagent, nation, servicepractice where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4)
'         strSQL = strSQL & " union select (a1k08 - nvl(a1k06, 0)) as Namount from acc1k0, fagent, nation where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and (a1k12 is null or a1k12 = 0)" & strWhere(0)
'         adoaccsum.Open "select sum(Namount) from (" & strSQL & ") New", adoTaie, adOpenStatic, adLockReadOnly
'         If adoaccsum.RecordCount <> 0 Then
'            If IsNull(adoaccsum.Fields(0).Value) Then
'               Text10 = MsgText(601)
'            Else
'               Text10 = Format(adoaccsum.Fields(0).Value, FDollar)
'            End If
'         Else
'            Text10 = MsgText(601)
'         End If
'         adoaccsum.Close
'         strSQL = ""
'         adoaccsum.CursorLocation = adUseClient
'         If Text5 <> "" Then
'            strWhere(4) = " and pa11 = '" & Text5 & "'"
'         End If
'         strSQL = "select (a1k08 - nvl(a1k06, 0) - nvl(a1k30, 0) / decode(a1k10, 0, 1, nvl(a1k10, 1))) as Namount from acc1k0, fagent, nation, patent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         If Text5 <> "" Then
'            strWhere(4) = " and tm12 = '" & Text5 & "'"
'         End If
'         strSQL = strSQL & " union select (a1k08 - nvl(a1k06, 0) - nvl(a1k30, 0) / decode(a1k10, 0, 1, nvl(a1k10, 1))) as Namount from acc1k0, fagent, nation, trademark where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         If Text5 <> "" Then
'            strWhere(4) = " and sp11 = '" & Text5 & "'"
'         End If
'         strSQL = strSQL & " union select (a1k08 - nvl(a1k06, 0) - nvl(a1k30, 0) / decode(a1k10, 0, 1, nvl(a1k10, 1))) as Namount from acc1k0, fagent, nation, servicepractice where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0)" & strWhere(0) & strWhere(4) & " and (a1k29 is null or a1k29 = '')"
'         adoaccsum.Open "select sum(Namount) from (" & strSQL & ") New", adoTaie, adOpenStatic, adLockReadOnly
'         If adoaccsum.RecordCount <> 0 Then
'            If IsNull(adoaccsum.Fields(0).Value) Then
'               Text15 = MsgText(601)
'            Else
'               Text15 = Format(adoaccsum.Fields(0).Value, FDollar)
'            End If
'         Else
'            Text15 = MsgText(601)
'         End If
'         adoaccsum.Close
'         strSQL = ""
'         If Text7 = "" Then
'            adoaccsum.CursorLocation = adUseClient
'            If Text5 <> "" Then
'               strWhere(4) = " and pa11 = '" & Text5 & "'"
'            End If
'            strSQL = "select sum(axf04) as Namount from acc151, acc150, fagent, nation, patent where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            If Text5 <> "" Then
'               strWhere(4) = " and tm12 = '" & Text5 & "'"
'            End If
'            strSQL = strSQL & " union select sum(axf04) as Namount from acc151, acc150, fagent, nation, trademark where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            If Text5 <> "" Then
'               strWhere(4) = " and sp11 = '" & Text5 & "'"
'            End If
'            strSQL = strSQL & " union select sum(axf04) as Namount from acc151, acc150, fagent, nation, servicepractice where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4)
'            adoaccsum.Open "select sum(Namount) from (" & strSQL & ") New", adoTaie, adOpenStatic, adLockReadOnly
'            If adoaccsum.RecordCount <> 0 Then
'               If IsNull(adoaccsum.Fields(0).Value) Then
'                  Text14 = MsgText(601)
'               Else
'                  Text14 = Format(adoaccsum.Fields(0).Value, FDollar)
'               End If
'            Else
'               Text14 = MsgText(601)
'            End If
'            adoaccsum.Close
'            strSQL = ""
'            adoaccsum.CursorLocation = adUseClient
'            If Text5 <> "" Then
'               strWhere(4) = " and pa11 = '" & Text5 & "'"
'            End If
'            strSQL = "select sum(axf04) as Namount from acc151, acc150, fagent, nation, patent where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null)"
'            strSQL = strSQL & " union select sum(a1904 * (-1)) as Namount from acc190, acc151, acc150, fagent, nation, patent where a1902 = axf01 and axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null)"
'            If Text5 <> "" Then
'               strWhere(4) = " and tm12 = '" & Text5 & "'"
'            End If
'            strSQL = strSQL & " union select sum(axf04) as Namount from acc151, acc150, fagent, nation, trademark where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null)"
'            strSQL = strSQL & " union select sum(a1904 * (-1)) as Namount from acc190, acc151, acc150, fagent, nation, trademark where a1902 = axf01 and axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null)"
'            If Text5 <> "" Then
'               strWhere(4) = " and sp11 = '" & Text5 & "'"
'            End If
'            strSQL = strSQL & " union select sum(axf04) as Namount from acc151, acc150, fagent, nation, servicepractice where axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null)"
'            strSQL = strSQL & " union select sum(a1904 * (-1)) as Namount from acc190, acc151, acc150, fagent, nation, servicepractice where a1902 = axf01 and axf01 = a1501 and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and fa10 = na01 (+) and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & strWhere(4) & " and (a1506 <> a1520 or a1520 is null)"
'            adoaccsum.Open "select sum(Namount) from (" & strSQL & ") New", adoTaie, adOpenStatic, adLockReadOnly
'            If adoaccsum.RecordCount <> 0 Then
'               If IsNull(adoaccsum.Fields(0).Value) Then
'                  Text16 = MsgText(601)
'               Else
'                  If adoaccsum.Fields(0).Value < 0 Then
'                     Text16 = MsgText(601)
'                  Else
'                     Text16 = Format(adoaccsum.Fields(0).Value, FDollar)
'                  End If
'               End If
'            Else
'               Text16 = MsgText(601)
'            End If
'            adoaccsum.Close
'         End If
'2009/4/28 END
'   End Select
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Text7 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text3 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text4 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text6 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text5 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text8 <> MsgText(601) Then
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
   If Text9 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

Private Sub Text9_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2003/05/06
    KeyAscii = UpperCase(KeyAscii)
    Select Case KeyAscii
    Case 8, 49, 50, 51, 52, 53, 54
    Case Else
        KeyAscii = 0
    End Select
End Sub

'2009/4/28 ADD BY SONIA
Private Sub Text9_Validate(Cancel As Boolean)
   If Text9 = "" Then Text9 = "5"
End Sub
'2009/4/28 END

''Add By Sindy 2012/8/15 計算抵帳單金額
'Private Function GetACC160Amt(strA1605 As String, strType As String) As Double
'   Dim rsTmp As New ADODB.Recordset
'   Dim strConSql As String
'
'   GetACC160Amt = 0
'   strConSql = ""
'
'   If strType = "0" Then '未付
'      strConSql = strConSql & " and a1607 is null"
'   Else
'      strConSql = strConSql & " and a1607 is not null"
'   End If
'   If strA1605 <> "" Then '幣別
'      strConSql = strConSql & " and A1605='" & strA1605 & "'"
'   End If
'
'   If Text5 <> "" Then
'      strWhere(4) = " and pa11 = '" & Text5 & "'"
'   End If
'                   strSql = "select a1605,sum(axg04) from acc161, acc160, patent where axg01 = a1601 and substr(axg03, 1, length(axg03) - 9) = pa01 and substr(axg03, length(axg03) - 8, 6) = pa02 and substr(axg03, length(axg03) - 2, 1) = pa03 and substr(axg03, length(axg03) - 1, 2) = pa04" & strWhere(5) & strWhere(4) & strConSql & " group by a1605"
'   If Text5 <> "" Then
'      strWhere(4) = " and tm12 = '" & Text5 & "'"
'   End If
'   strSql = strSql & " union select a1605,sum(axg04) from acc161, acc160, trademark where axg01 = a1601 and substr(axg03, 1, length(axg03) - 9) = tm01 and substr(axg03, length(axg03) - 8, 6) = tm02 and substr(axg03, length(axg03) - 2, 1) = tm03 and substr(axg03, length(axg03) - 1, 2) = tm04" & strWhere(5) & strWhere(4) & strConSql & " group by a1605"
'   If Text5 <> "" Then
'      strWhere(4) = " and sp11 = '" & Text5 & "'"
'   End If
'   strSql = strSql & " union select a1605,sum(axg04) from acc161, acc160, servicepractice where axg01 = a1601 and substr(axg03, 1, length(axg03) - 9) = sp01 and substr(axg03, length(axg03) - 8, 6) = sp02 and substr(axg03, length(axg03) - 2, 1) = sp03 and substr(axg03, length(axg03) - 1, 2) = sp04" & strWhere(5) & strWhere(4) & strConSql & " group by a1605"
'   strSql = strSql & " union select a1605,sum(axg04) from acc161, acc160, LAWCASE where axg01 = a1601 and substr(axg03, 1, length(axg03) - 9) = lc01 and substr(axg03, length(axg03) - 8, 6) = lc02 and substr(axg03, length(axg03) - 2, 1) = lc03 and substr(axg03, length(axg03) - 1, 2) = lc04" & strWhere(5) & strConSql & " group by a1605"
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
