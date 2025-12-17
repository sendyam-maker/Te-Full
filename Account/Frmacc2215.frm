VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc2215 
   AutoRedraw      =   -1  'True
   Caption         =   "結匯資料查詢 "
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4860
   ScaleWidth      =   8760
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Height          =   330
      Left            =   1344
      MaxLength       =   9
      TabIndex        =   10
      Top             =   216
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
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
      Height          =   330
      Left            =   6864
      MaxLength       =   15
      TabIndex        =   9
      Top             =   573
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00FFFFFF&
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
      Height          =   330
      Left            =   1344
      MaxLength       =   14
      TabIndex        =   8
      Top             =   930
      Width           =   1572
   End
   Begin VB.TextBox Text14 
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
      Height          =   330
      Left            =   3780
      TabIndex        =   7
      Top             =   4425
      Width           =   1332
   End
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4104
      TabIndex        =   6
      Top             =   566
      Width           =   1572
   End
   Begin VB.TextBox Text12 
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
      Height          =   330
      Left            =   5145
      TabIndex        =   4
      Top             =   4425
      Width           =   1380
   End
   Begin VB.TextBox Text20 
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
      Height          =   330
      Left            =   1380
      MaxLength       =   12
      TabIndex        =   3
      Top             =   4425
      Width           =   855
   End
   Begin VB.TextBox Text15 
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
      Height          =   330
      Left            =   6555
      TabIndex        =   2
      Top             =   4425
      Width           =   1788
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2215.frx":0000
      Height          =   810
      Left            =   1335
      TabIndex        =   5
      Top             =   1605
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1402
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483624
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "a1902"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1785.26
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Frmacc2215.frx":0015
      Height          =   1755
      Left            =   240
      TabIndex        =   1
      Top             =   2610
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3096
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
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
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "a0102"
         Caption         =   "會計科目"
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
         DataField       =   "a1p07"
         Caption         =   "借方金額"
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
      BeginProperty Column02 
         DataField       =   "a1p08"
         Caption         =   "貸方金額"
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
      BeginProperty Column03 
         DataField       =   "a1p21"
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
      BeginProperty Column04 
         DataField       =   "a1p11"
         Caption         =   "銀行帳號"
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
         DataField       =   "a0g02"
         Caption         =   "銀行名稱"
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
         DataField       =   "a1p20"
         Caption         =   "銀存匯率"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "a1p17"
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
      BeginProperty Column08 
         DataField       =   "a1p14"
         Caption         =   "摘要"
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
            ColumnWidth     =   3270.047
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1769.953
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1769.953
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   5864.882
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   330
      Left            =   1344
      TabIndex        =   11
      Top             =   573
      Width           =   1572
      _ExtentX        =   2752
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
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
      Height          =   330
      Left            =   6864
      TabIndex        =   12
      Top             =   216
      Width           =   1572
      _ExtentX        =   2752
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
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
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "Frmacc2215.frx":002A
      Height          =   810
      Left            =   5280
      TabIndex        =   0
      Top             =   1605
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1402
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "a1c03"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1785.26
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   1320
      Top             =   1455
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   315
      Left            =   5280
      Top             =   1455
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   315
      Left            =   240
      Top             =   2490
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
      Caption         =   "Adodc3"
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
      Left            =   2940
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   216
      Width           =   2775
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "4895;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   480
      Left            =   4110
      TabIndex        =   24
      Top             =   930
      Width           =   4335
      VariousPropertyBits=   -1467989985
      ScrollBars      =   2
      Size            =   "7641;847"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   24
      Top             =   4272
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "代理人"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   384
      TabIndex        =   23
      Top             =   255
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "結匯日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   22
      Top             =   612
      Width           =   972
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "匯票號碼"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5880
      TabIndex        =   21
      Top             =   612
      Width           =   972
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "手續費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   384
      TabIndex        =   20
      Top             =   969
      Width           =   972
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "備註"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3144
      TabIndex        =   19
      Top             =   969
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1365
      Left            =   270
      Top             =   150
      Width           =   8295
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
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
      Height          =   252
      Left            =   2904
      TabIndex        =   18
      Top             =   4471
      Width           =   612
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "作業日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5904
      TabIndex        =   17
      Top             =   255
      Width           =   972
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "匯票方式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3120
      TabIndex        =   16
      Top             =   612
      Width           =   972
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   8748
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   0
      X2              =   8748
      Y1              =   2490
      Y2              =   2490
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "未結匯單據"
      Height          =   975
      Left            =   1080
      TabIndex        =   15
      Top             =   1530
      Width           =   255
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "結匯單據"
      Height          =   855
      Left            =   5040
      TabIndex        =   14
      Top             =   1575
      Width           =   255
   End
   Begin VB.Label Label22 
      BackStyle       =   0  '透明
      Caption         =   "筆數"
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
      Left            =   420
      TabIndex        =   13
      Top             =   4470
      Width           =   855
   End
End
Attribute VB_Name = "Frmacc2215"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/09 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Text2、Text5
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoacc1b0 As New ADODB.Recordset
Public adoacc190 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoadodc2 As New ADODB.Recordset
Public adoadodc3 As New ADODB.Recordset
Public adoacc1p0 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Dim strSerialNo As String
Dim strCurrency As String
Dim strNo As String
Dim douTAmount As Double
Dim douLAmount As Double
Dim strAccNo As String
Dim strYes As String
Dim strA1917 As String   'add by sonia 2021/7/8

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
'   Me.Width = 8850
'   Me.Height = 5050
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
   PUB_InitForm Me, 8850, 5300, strBackPicPath1
   'end 2021/12/09
   
   Text3 = strItemNo
   Text1 = strCompanyNo
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   OpenTable
   AdodcRefresh
   SumShow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strItemNo = ""
   tool3_enabled
   Frmacc2214.Enabled = True
   Set Frmacc2215 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
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
   adoacc1b0.CursorLocation = adUseClient
   adoacc1b0.Open "select * from acc1b0 where a1b01 = '" & Text3 & "' and a1b02 = '" & Text1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc1b0.RecordCount <> 0 Then
      FormShow
   End If
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc190, acc180 where acc190.a1901 = acc180.a1801 and a1915 = " & Val(FCDate(MaskEdBox2.Text)) & " and a1803 = '" & Text1 & "' and a1908 is null order by a1902 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
   adoadodc2.CursorLocation = adUseClient
   adoadodc2.Open "select * from acc1c0, acc150, acc190 where a1c03 = a1501 (+) and a1c03 = a1902 and a1c01 = '" & Text3 & "' and a1c02 like '" & Text1 & "%" & "' order by decode(substr(a1c03, 1, 1), 'U', 1, 'V', 2, 'O', 3) asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc2.Recordset = adoadodc2
   adoadodc3.CursorLocation = adUseClient
   adoadodc3.Open "select * from acc1p0, acc010, acc0g0 where a1p05 = a0101 (+) and a1p10 = a0g01 (+) and a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "' order by a1p03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc3.Recordset = adoadodc3
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
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      adoadodc1.Open "select * from acc190, acc180 where a1901 = a1801 and a1803 = '" & Text1 & "' and (a1908 is null or a1908 = '') order by a1902 asc", adoTaie, adOpenStatic, adLockReadOnly
   Else
      adoadodc1.Open "select * from acc190, acc180 where a1901 = a1801 and a1803 = 'A' and (a1908 is null or a1908 = '') order by a1902 asc", adoTaie, adOpenStatic, adLockReadOnly
   End If
   Adodc1.Recordset.Requery
   adoadodc2.Close
   adoadodc2.CursorLocation = adUseClient
   adoadodc2.Open "select * from acc1c0, acc150, acc190 where a1c03 = a1501 (+) and a1c03 = a1902 and a1c01 = '" & Text3 & "' and a1c02 like '" & Text1 & "%" & "' order by decode(substr(a1c03, 1, 1), 'U', 1, 'V', 2, 'O', 3) asc", adoTaie, adOpenStatic, adLockReadOnly
   Adodc2.Recordset.Requery
   strA1917 = ""   'add by sonia 2021/7/8
   If Adodc2.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc2.Recordset.Fields("a1505").Value) Then
         strCurrency = "USD"
         strA1917 = "1"
      Else
         strCurrency = Adodc2.Recordset.Fields("a1505").Value
         strA1917 = Adodc2.Recordset.Fields("a1917").Value
      End If
   Else
      strCurrency = "USD"
   End If
   adoadodc3.Close
   adoadodc3.CursorLocation = adUseClient
   'modify by sonia 2021/7/8 傳票公司別要抓A1917
   'adoadodc3.Open "select * from acc1p0, acc010, acc0g0 where a1p05 = a0101 (+) and a1p10 = a0g01 (+) and a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "' order by a1p03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc3.Open "select * from acc1p0, acc010, acc0g0 where a1p05 = a0101 (+) and a1p10 = a0g01 (+) and a1p01 = '" & strA1917 & "' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "' order by a1p03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc3.Recordset.Requery
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表
'
'*************************************************
Public Sub FormShow()
   Text1 = adoacc1b0.Fields("a1b02").Value
   If Len(Text1) = 6 Then
      Text2 = FagentQuery(AfterZero(Text1), 2)
      Text1 = AfterZero(Text1)
   Else
      Text2 = FagentQuery(Text1, 2)
   End If
   Text3 = adoacc1b0.Fields("a1b01").Value
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc1b0.Fields("a1b03").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(adoacc1b0.Fields("a1b03").Value)
   End If
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(adoacc1b0.Fields("a1b05").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(adoacc1b0.Fields("a1b05").Value)
   End If
   MaskEdBox2.Mask = DFormat
   If IsNull(adoacc1b0.Fields("a1b06").Value) Then
      Combo2 = MsgText(601)
   Else
      Select Case adoacc1b0.Fields("a1b06").Value
         Case "1"
            Combo2 = "票匯"
         Case "2"
            Combo2 = "電匯"
      End Select
   End If
   If IsNull(adoacc1b0.Fields("a1b04").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = adoacc1b0.Fields("a1b04").Value
   End If
   If IsNull(adoacc1b0.Fields("a1b07").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = adoacc1b0.Fields("a1b07").Value
   End If
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("fagent", "fa01 || fa02", IIf(Len(Text1) = 6, AfterZero(Text1), Text1), Label1) = False Then
      Cancel = True
      Exit Sub
   End If
   If Len(Text1) = 6 Then
      Text2 = FagentQuery(AfterZero(Text1), 2)
      Text1 = AfterZero(Text1)
   Else
      If Len(Text1) = 8 Then
         Text1 = Text1 & "0"
      End If
      Text2 = FagentQuery(Text1, 2)
   End If
End Sub


Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   AdodcRefresh
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
   'add by nickc 2007/07/13 將輸入法改成使用API
   OpenIme
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   KeyEnter KeyCode
End Sub

'*************************************************
'  清除顯示資料
'
'*************************************************
Public Sub Adodc3Clear()
   'edit by nickc 2007/02/08
   'Text6 = ""
   'Text13 = ""
   'Text16 = ""
   'Text7 = ""
   'Text11 = ""
   'Text9 = ""
   'Text10 = ""
   'Text8 = ""
   'Combo1 = ""
End Sub


'*************************************************
'  計算並顯示合計
'
'*************************************************
Public Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   'modify by sonia 2021/7/8 傳票公司別要抓A1917
   'adoaccsum.Open "select sum(a1p07), sum(a1p08), count(*), sum(decode(substr(a1p05, 1, 1), '2', a1p21, 0)) from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   adoaccsum.Open "select sum(a1p07), sum(a1p08), count(*), sum(decode(substr(a1p05, 1, 1), '2', a1p21, 0)) from acc1p0 where a1p01 = '" & strA1917 & "' and a1p02 = 'I' and a1p04 = '" & Text3 & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.EOF = False And adoaccsum.BOF = False Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text14 = MsgText(601)
      Else
         Text14 = Format(Val(adoaccsum.Fields(0).Value), FAmount)
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text12 = MsgText(601)
      Else
         Text12 = Format(Val(adoaccsum.Fields(1).Value), FAmount)
      End If
      If IsNull(adoaccsum.Fields(2).Value) Then
         Text20 = MsgText(601)
      Else
         Text20 = Format(adoaccsum.Fields(2).Value, DAmount)
      End If
      If IsNull(adoaccsum.Fields(3).Value) Then
         Text15 = MsgText(601)
      Else
         Text15 = Format(adoaccsum.Fields(3).Value, FAmount)
      End If
   Else
      Text14 = MsgText(601)
      Text12 = MsgText(601)
      Text20 = MsgText(601)
      Text15 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   If adoacc1b0.RecordCount = 0 Then
      Exit Sub
   End If
   CountShow adoacc1b0.Bookmark, adoacc1b0.RecordCount
End Sub


'add by nickc 2007/07/13 將輸入法改成使用API
Private Sub Text5_Validate(Cancel As Boolean)
CloseIme
End Sub
