VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc21f0 
   AutoRedraw      =   -1  'True
   Caption         =   "抵帳作業"
   ClientHeight    =   5650
   ClientLeft      =   50
   ClientTop       =   280
   ClientWidth     =   8780
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5650
   ScaleWidth      =   8780
   Begin VB.CommandButton Command6 
      Caption         =   "匯入付款單"
      Height          =   324
      Left            =   2928
      TabIndex        =   11
      Top             =   3144
      Width           =   1404
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   3360
      Style           =   2  '單純下拉式
      TabIndex        =   14
      Top             =   5130
      Width           =   2900
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00FFFFFF&
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
      Height          =   345
      Left            =   1125
      MaxLength       =   12
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   540
      Width           =   1092
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00FFFFFF&
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
      Height          =   345
      Left            =   7335
      MaxLength       =   12
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   540
      Width           =   1092
   End
   Begin VB.CommandButton Command4 
      Caption         =   "列印抵帳資料"
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
      Left            =   6510
      TabIndex        =   31
      Top             =   5130
      Width           =   1692
   End
   Begin VB.TextBox Text11 
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
      Left            =   7125
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1080
      Width           =   480
   End
   Begin VB.TextBox Text10 
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
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   13
      Top             =   4785
      Width           =   612
   End
   Begin VB.CommandButton Command5 
      Height          =   300
      Left            =   3480
      Picture         =   "Frmacc21f0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   170
      Width           =   350
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00C0FFFF&
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
      Left            =   2040
      MaxLength       =   15
      TabIndex        =   0
      Top             =   156
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Frmacc21f0.frx":0102
      Height          =   1200
      Left            =   210
      TabIndex        =   16
      Top             =   3555
      Width           =   8295
      _ExtentX        =   14623
      _ExtentY        =   2117
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "a1501"
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
      BeginProperty Column01 
         DataField       =   "axf03"
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
      BeginProperty Column02 
         DataField       =   "a1505"
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
         DataField       =   "a1510"
         Caption         =   "台幣金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "a1506"
         Caption         =   "外幣金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1730.268
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1759.748
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   560.126
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1840.252
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1860.095
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc21f0.frx":0117
      Height          =   1200
      Left            =   240
      TabIndex        =   15
      Top             =   1410
      Width           =   8295
      _ExtentX        =   14623
      _ExtentY        =   2117
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "A1K01"
         Caption         =   "請款編號"
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
         DataField       =   "A1K13||A1K14||A1K15||A1K16"
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
      BeginProperty Column02 
         DataField       =   "A1K08"
         Caption         =   "外幣金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "A1K11"
         Caption         =   "外幣現值(NT)"
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
         DataField       =   "A1K09"
         Caption         =   "規費"
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
         DataField       =   "Property"
         Caption         =   "案件性質"
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
         Size            =   344
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   1429.795
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1429.795
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1280.126
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1450.205
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   1080
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7785
      Picture         =   "Frmacc21f0.frx":012C
      Style           =   1  '圖片外觀
      TabIndex        =   12
      ToolTipText     =   "取消"
      Top             =   3120
      Width           =   492
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7800
      Picture         =   "Frmacc21f0.frx":0796
      Style           =   1  '圖片外觀
      TabIndex        =   8
      ToolTipText     =   "取消"
      Top             =   990
      Width           =   492
   End
   Begin VB.CommandButton Command2 
      Caption         =   "抵帳結果資料"
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
      Left            =   6696
      TabIndex        =   9
      Top             =   2655
      Width           =   1692
   End
   Begin VB.TextBox Text8 
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
      Left            =   6348
      TabIndex        =   26
      Top             =   4770
      Width           =   1932
   End
   Begin VB.TextBox Text7 
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
      Left            =   4536
      TabIndex        =   25
      Top             =   4770
      Width           =   1812
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  '靠右對齊
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
      Height          =   315
      Left            =   7335
      MaxLength       =   13
      TabIndex        =   3
      Top             =   170
      Width           =   1092
   End
   Begin VB.TextBox Text5 
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
      Height          =   315
      Left            =   1296
      MaxLength       =   15
      TabIndex        =   10
      Top             =   3130
      Width           =   1572
   End
   Begin VB.TextBox Text4 
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
      Left            =   4632
      TabIndex        =   21
      Top             =   2655
      Width           =   1440
   End
   Begin VB.TextBox Text3 
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
      Left            =   3300
      TabIndex        =   20
      Top             =   2655
      Width           =   1332
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '靠右對齊
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
      Left            =   5040
      MaxLength       =   12
      TabIndex        =   2
      Top             =   170
      Width           =   1092
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
      Height          =   315
      Left            =   1350
      MaxLength       =   15
      TabIndex        =   6
      Top             =   1000
      Width           =   1572
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   315
      Left            =   210
      Top             =   3435
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   2117
      _ExtentY        =   564
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   240
      Top             =   1290
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   2117
      _ExtentY        =   564
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
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "匯入後刪除付款單資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   228
      Left            =   4368
      TabIndex        =   38
      Top             =   3192
      Width           =   2796
   End
   Begin MSForms.TextBox Text14 
      Height          =   345
      Left            =   2220
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   540
      Width           =   3930
      VariousPropertyBits=   671105055
      MaxLength       =   80
      Size            =   "6932;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "印表機:"
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
      TabIndex        =   36
      Top             =   5130
      Width           =   855
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "代理人"
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
      Left            =   360
      TabIndex        =   35
      Top             =   540
      Width           =   675
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "抵帳日期"
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
      Left            =   6300
      TabIndex        =   34
      Top             =   540
      Width           =   900
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "D09300001"
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
      Left            =   750
      TabIndex        =   33
      Top             =   2655
      Width           =   1515
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "傳票"
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
      Left            =   240
      TabIndex        =   32
      Top             =   2655
      Width           =   525
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "收款(1)或結匯(2)"
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
      Left            =   5265
      TabIndex        =   30
      Top             =   1095
      Width           =   2100
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "是否結清"
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
      Left            =   225
      TabIndex        =   29
      Top             =   4785
      Width           =   975
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
      Left            =   1905
      TabIndex        =   28
      Top             =   4785
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00404040&
      Height          =   915
      Left            =   225
      Top             =   45
      Width           =   8295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "FC/CF 抵帳編號"
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
      TabIndex        =   27
      Top             =   170
      Width           =   1812
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   -96
      Top             =   4728
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label6 
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
      Left            =   3630
      TabIndex        =   24
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "CF匯率"
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
      Left            =   6360
      TabIndex        =   23
      Top             =   170
      Width           =   852
   End
   Begin VB.Label Label4 
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
      Left            =   330
      TabIndex        =   22
      Top             =   3195
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   -30
      X2              =   8730
      Y1              =   3015
      Y2              =   3015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   -30
      X2              =   8730
      Y1              =   3030
      Y2              =   3030
   End
   Begin VB.Label Label3 
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
      Left            =   2370
      TabIndex        =   19
      Top             =   2655
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "FC匯率"
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
      Left            =   4080
      TabIndex        =   18
      Top             =   170
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "請款編號"
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
      TabIndex        =   17
      Top             =   1050
      Width           =   1215
   End
End
Attribute VB_Name = "Frmacc21f0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/07 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、DataGrid2改字型=新細明體-ExtB、Text14; Printer列印未改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoacc1g0 As New ADODB.Recordset
Public adoacc150 As New ADODB.Recordset
Public adoacc1k0 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoadodc2 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adocheck As New ADODB.Recordset
Dim intCounter As Integer
Dim m_intPage As Integer '頁數
Const m_dblLeft As Double = 500 '橫軸偏移值
'Add by Morgan 2011/9/6
Dim m_bolAlert As Boolean '檢查分錄提醒
Dim m_strAlertMsg As String
'Added by Lydia 2018/11/05
Dim strPrinter As String '系統預設印表機
Dim strPrtOrt As Integer '系統預設印表機的紙張方向

Private Sub Command1_Click()
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then Exit Sub 'Added by Morgan 2023/10/23
         
   AdodcDelete1
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine1 KeyCode
End Sub

Private Sub Command2_Click()
'Add By Cheng 2004/02/27
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim blnReImportData As Boolean '判斷是否重新匯入資料
'End
'Add by Amy 2013/10/30
Dim strUpd As String
   
   If Text9 = MsgText(601) Then
      Exit Sub
   End If
    'Add By Cheng 2004/03/09
    '判斷傳票是否過帳
    If ChkPosting(Me.Label12.Caption) = True Then Exit Sub
    'End
    
    'add by sonia 2025/4/10
    If Val(Text2) = 0 Or Val(Text6) = 0 Then
       MsgBox "FC匯率或CF匯率尚未輸入 無法產生抵帳結果資料!!!", vbExclamation + vbOKOnly
       Exit Sub
    End If
    'end 2025/4/10
    
    'Add by Amy 2013/10/30 +抵帳訊息
    strExc(0) = GetDizhang(Left(Text13, 8), , False, 2)
    If strExc(0) <> MsgText(601) Then
       If MsgBox("此編號有設定帳款處理情形 (" & strExc(0) & ")" & vbCrLf & "此次抵帳後是否維持原狀態？ 是：維持，N：取消設定" & vbCrLf, vbYesNo + vbDefaultButton2) = vbNo Then
          '取消設定將 CU142 或 fa103 更新為null
          If Left(Text13, 1) = "X" Then
               strUpd = "Update Customer set CU142=Null Where CU01='" & Left(Text13, 8) & "' "
          Else
              strUpd = "Update Fagent set FA103=Null Where FA01='" & Left(Text13, 8) & "' "
          End If
          'Pub_SeekTbLog strUpd 'Modify by amy 2013/11/01 取消記log(目前只有在frmaacc21r0選宣告破產才記log)
          adoTaie.Execute strUpd
          
          If strExc(0) = "宣告破產" Then '宣告破產取消時發mail 給秀玲
            PUB_SendMail strUserNum, "83002", "", Left(Text13, 8) & " 財務取消宣告破產設定", "請至客戶或代理人檔取消狀態、呆帳記錄及備註欄之加註！"
          End If
       End If
    End If
   'end 2013/10/30
   
   Screen.MousePointer = vbHourglass
   
   strItemNo = Text9
   If Text10 <> "" Then
      strCon1 = Text10
   Else
      strCon1 = ""
   End If
    'Add By Cheng 2004/02/27
    blnReImportData = False
    If Me.Text11.Text = "1" Then
        '進入Frmacc21f2
        StrSQLa = "Select * From acc1p0, acc010 where a1p05 = a0101 and a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & strItemNo & "' order by a1p03 asc"
    Else
        '進入Frmacc21f1
        StrSQLa = "Select * From acc1p0, acc010, acc0g0 Where a1p05 = a0101 and a1p10 = a0g01 (+) and a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & strItemNo & "' order by a1p03 asc "
    End If
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
    '若未輸入抵帳結果資料
    If rsA.EOF = True Then
        blnReImportData = True
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    If blnReImportData = True Then
      ProcessData
    End If
    
   'Added by Lydia 2018/11/05 若有變動印表機, 先更新列印設定,令三畫面(Frmacc21f0~Frmacc21f2)的預設印表機一致
    If Me.cmbPrinter.Text <> Me.cmbPrinter.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, "0", "0", Me.cmbPrinter.Text
    End If
   'end 2018/11/05
   
   tool7_enabled
   If Text11 = "1" Then
      Frmacc21f2.Show
   Else
      Frmacc21f1.Show
   End If
   Me.Hide
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command3_Click()
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then Exit Sub 'Added by Morgan 2023/10/23
   
   AdodcDelete2
End Sub

Private Sub Command3_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine2 KeyCode
End Sub
'Memo by Lydia 2018/11/05 列印抵帳資料
Private Sub Command4_Click()
   If Text9 = MsgText(601) Then
      Exit Sub
   End If
   strItemNo = Text9
   Screen.MousePointer = vbHourglass
   
   PUB_RestorePrinter cmbPrinter 'Added by Lydia 2018/11/05 改印表機
   '2009/7/2 MODIFY BY SONIA
   'If Me.Text11.Text = "1" Then
   If Val(Text7) > Val(Text4) Then
       PrintDataf2
   Else
       PrintDataf1
   End If
   PUB_RestorePrinter strPrinter, strPrtOrt 'Added by Lydia 2018/11/05 還原系統印表機
   
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command5_Click()
   If adoacc1g0.RecordCount = 0 Or Text9 = MsgText(601) Then
      Exit Sub
   End If
   adoacc1g0.Find "a1g01 = '" & Text9 & "'", 0, adSearchForward, 1
   If adoacc1g0.EOF Then
      MsgBox MsgText(33), , MsgText(5)
      adoacc1g0.MoveFirst
   End If
   FormShow
   AdodcRefresh1
   AdodcRefresh2
   SumShow1
   SumShow2
   RecordShow
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    'Add By Cheng 2004/03/10
    If ColIndex = 2 Then
        SendKeys "{Tab}"
        Me.Adodc1.Recordset("A1K11").Value = Format(Val("" & Me.Adodc1.Recordset("A1K08").Value) * Val(Me.Text2.Text), FAmount)
    End If
    'End
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine1 KeyCode
End Sub

Private Sub DataGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine2 KeyCode
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   If adoacc1g0.RecordCount <> 0 Then
      adoacc1g0.MoveFirst
   End If
   adoacc1g0.Find "a1g01 = '" & strItemNo & "'", 0, adSearchForward, 1
   If adoacc1g0.EOF = False Then
      FormShow
      AdodcRefresh1
      AdodcRefresh2
      SumShow1
      SumShow2
      RecordShow
   End If
   strItemNo = MsgText(601)
End Sub

'Added by Lydia 2021/12/07
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Added by Lydia 2021/12/07 Form2.0 記錄鍵盤傳入順序
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call PUB_SaveTrackMode(1, KeyCode)  'Added by Lydia 2021/12/07 Form2.0 記錄鍵盤傳入順序
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   'Added by Lydia 2018/11/05 預設印表機選項
   strPrtOrt = Printer.Orientation
   PUB_SetPrinter Me.Name, cmbPrinter, strPrinter
   '2018/11/05
   
   'Modified by Lydia 2021/12/07 改成模組
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   Me.Width = 8850
'   'Modified by Lydia 2018/11/05
'   'Me.Height = 5820
'   Me.Height = 5900
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath1)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
'   Next
   strFormName = Name
   'Modify by Amy 2023/08/18 W8850 H5900
   PUB_InitForm Me, 8880, 6100, strBackPicPath1
   'end 2021/12/07

   OpenTable
   If adoacc1g0.RecordCount <> 0 Then
      adoacc1g0.MoveLast
      adoacc1g0.MoveFirst
      RecordShow
   End If
   FormDisabled
    'Add By Cheng 2004/03/09
    '傳票號碼
    Label12 = ""
    'End
    'Add by Morgan 2004/11/25
    Text1.Text = "X"
    Text5.Text = "U"
End Sub

Private Sub Form_Resize()
   tool1_enabled
   strFormName = Name
End Sub

Private Sub Form_Unload(Cancel As Integer)

   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   strTrackMode = "" 'Added by Lydia 2021/12/07 Form2.0 記錄鍵盤傳入順序(清除)
   
   'Added by Lydia 2018/11/05 若有變動印表機, 則更新列印設定
    If Me.cmbPrinter.Text <> Me.cmbPrinter.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, "0", "0", Me.cmbPrinter.Text
    End If
   'end 2018/11/05
   
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc21f0 = Nothing
End Sub

Private Sub Text1_GotFocus()
   'MODIFY BY SONIA 2015/5/18
   'TextInverse Text1
   If Len(Text1) > 0 Then
      Text1.SelStart = 1
      Text1.SelLength = Len(Text1) - 1
   End If
   '2015/5/18
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine1 KeyCode
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc1g0.CursorLocation = adUseClient
   adoacc1g0.Open "select * from acc1g0 order by a1g01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
'   adoadodc1.Open "select a1k01, a1k13 || a1k14 || a1k15 || a1k16, a1k11, a1k08, a1k09, cp10, nvl(cpm03, cpm04) as Property from caseprogress, acc1k0, casepropertymap where cp60 = a1k01 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and a1k17 = '" & Text9 & "' order by a1k01 asc", adoTaie, adOpenStatic, adLockReadOnly
   '93.9.21 MODIFY BY SONIA
   'adoadodc1.Open "select a1k01, a1k13 || a1k14 || a1k15 || a1k16, a1k11, a1k08, a1k09, cp10, nvl(cpm03, cpm04) as Property, A1K30 from caseprogress, acc1k0, casepropertymap where cp60 = a1k01 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and a1k17 = '" & Text9 & "' order by a1k01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modified by Lydia 2022/12/14 抵帳作業Frmacc21f0的傳票，請依本所案號排序
   'adoadodc1.Open "select a1k01, a1k13 || a1k14 || a1k15 || a1k16, a1k11, a1k08-nvl(a1k06, 0), a1k09, cp10, nvl(cpm03, cpm04) as Property, A1K30 from caseprogress, acc1k0, casepropertymap where cp60 = a1k01 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and a1k17 = '" & Text9 & "' order by a1k01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.Open "select a1k01, a1k13 || a1k14 || a1k15 || a1k16, a1k11, a1k08-nvl(a1k06, 0), a1k09, cp10, nvl(cpm03, cpm04) as Property, A1K30 from caseprogress, acc1k0, casepropertymap where cp60 = a1k01 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and a1k17 = '" & Text9 & "' order by a1k13 asc, a1k14 asc, a1k15 asc, a1k16 asc, a1k01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '93.9.21 END
   Set Adodc1.Recordset = adoadodc1
   adoadodc2.CursorLocation = adUseClient
   'Modified by Lydia 2022/12/14 抵帳作業Frmacc21f0的傳票，請依本所案號排序
   'adoadodc2.Open "select * from acc151, acc150 where axf01 = a1501 and a1512 = '" & Text9 & "' order by axf01 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc2.Open "select * from acc151, acc150 where axf01=a1501 and a1512='" & Text9 & "' order by axf03 asc, axf02 asc, axf01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc2.Recordset = adoadodc2
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  重新整理 Adodc1 之資料
'
'*************************************************
Public Sub AdodcRefresh1()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   'Modify by Morgan 2011/9/2 +a1k02,a1k10,a1k18
   'Modified by Morgan 2015/10/15 a1k06改放台幣折讓金額,請款幣別折讓金額;+SFee
   'adoadodc1.Open "select a1k01, a1k13 || a1k14 || a1k15 || a1k16, round((a1k08-nvl(a1k06, 0)) * " & Val(Text2) & ", 2) as a1k11, (a1k08-nvl(a1k06, 0)) AS A1K08, a1k09, nvl(cpm03, cpm04) as Property, a1k30, a1k03,a1k02,a1k10,a1k18 from acc1k0, (select cp01, cp60, min(cp10) as cp10 from acc1k0, caseprogress where a1k01 = cp60 and a1k17 = '" & Text9 & "' group by cp01, cp60) new, casepropertymap where a1k01 = cp60 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and a1k17 = '" & Text9 & "' order by a1k01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modified by Lydia 2022/12/14 抵帳作業Frmacc21f0的傳票，請依本所案號排序
   'adoadodc1.Open "select a1k01, a1k13 || a1k14 || a1k15 || a1k16, round((a1k08-nvl(a1k31, 0)) * " & Val(Text2) & ", 2) as a1k11, (a1k08-nvl(a1k31, 0)) AS A1K08, a1k09, nvl(cpm03, cpm04) as Property, a1k30, a1k03,a1k02,a1k10,a1k18,a1k11-nvl(a1k06,0)-nvl(a1k09,0) SFee from acc1k0, (select cp01, cp60, min(cp10) as cp10 from acc1k0, caseprogress where a1k01 = cp60 and a1k17 = '" & Text9 & "' group by cp01, cp60) new, casepropertymap where a1k01 = cp60 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and a1k17 = '" & Text9 & "' order by a1k01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.Open "select a1k01, a1k13 || a1k14 || a1k15 || a1k16, round((a1k08-nvl(a1k31, 0)) * " & Val(Text2) & ", 2) as a1k11, (a1k08-nvl(a1k31, 0)) AS A1K08, a1k09, nvl(cpm03, cpm04) as Property, a1k30, a1k03,a1k02,a1k10,a1k18,a1k11-nvl(a1k06,0)-nvl(a1k09,0) SFee from acc1k0, (select cp01, cp60, min(cp10) as cp10 from acc1k0, caseprogress where a1k01 = cp60 and a1k17 = '" & Text9 & "' group by cp01, cp60) new, casepropertymap where a1k01 = cp60 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and a1k17 = '" & Text9 & "' order by a1k13 asc, a1k14 asc, a1k15 asc, a1k16 asc, a1k01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      Text13 = "" & adoadodc1.Fields("a1k03").Value 'Add by Morgan 2006/7/17
      Adodc1.Recordset.Find "a1k01 = '" & Text1 & "'", 0, adSearchForward, 1
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  重新整理 Adodc2 之資料
'
'*************************************************
Public Sub AdodcRefresh2()
On Error GoTo Checking
   adoadodc2.Close
   adoadodc2.CursorLocation = adUseClient
   'Modified by Lydia 2022/12/14 抵帳作業Frmacc21f0的傳票，請依本所案號排序
   'adoadodc2.Open "select a1501, axf03, a1505, round(axf04 * '" & Val(Text6) & "') as a1510, axf04 as a1506, axf02, a1503 from acc151, acc150 where axf01 = a1501 and a1512 = '" & Text9 & "' order by a1501, axf03, a1505, a1510, a1506", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc2.Open "select a1501, axf03, a1505, round(axf04 * '" & Val(Text6) & "') as a1510, axf04 as a1506, axf02, a1503 from acc151, acc150 where axf01 = a1501 and a1512 = '" & Text9 & "' order by axf03, axf02 ,a1501, a1505, a1510, a1506", adoTaie, adOpenStatic, adLockReadOnly
   Adodc2.Recordset.Requery
   If Adodc2.Recordset.RecordCount <> 0 Then
      Text13 = "" & adoadodc2.Fields("a1503").Value 'Add by Morgan 2006/7/17
      Adodc2.Recordset.Find "a1501 = '" & Text5 & "'", 0, adSearchForward, 1
   End If
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
   Text9 = adoacc1g0.Fields("a1g01").Value
   SetDate Text9 'Add by Morgan 2006/7/17
   If IsNull(adoacc1g0.Fields("a1g02").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = adoacc1g0.Fields("a1g02").Value
   End If
   If IsNull(adoacc1g0.Fields("a1g03").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = adoacc1g0.Fields("a1g03").Value
   End If
   If IsNull(adoacc1g0.Fields("a1g10").Value) Then
      Text10 = MsgText(601)
   Else
      Text10 = adoacc1g0.Fields("a1g10").Value
   End If
   'Add By Cheng 2004/03/09
   Me.Label12.Caption = GetSummonsNo("1", "K", Me.Text9.Text)
   'End
   'add by sonia 2017/3/30 已過帳不可修改抵帳明細 Z092
   If ChkPosting(Me.Label12.Caption, "N") = True Then
      Command1.Enabled = False
      Command3.Enabled = False
   Else
      Command1.Enabled = True
      Command3.Enabled = True
   End If
   'end 2017/3/30
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text13_Change()
   Text14 = FagentQuery_1(Text13, 2)
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text5_GotFocus()
   'MODIFY BY SONIA 2015/5/18
   'TextInverse Text5
   If Len(Text5) > 0 Then
      Text5.SelStart = 1
      Text5.SelLength = Len(Text5) - 1
   End If
   '2015/5/18
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine2 KeyCode
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  儲存資料表(國外請款單資料)
'
'*************************************************
Private Sub Acc1k0Save()
On Error GoTo Checking
   If Text1 = "" Then
      Exit Sub
   End If
   adoacc1k0.CursorLocation = adUseClient
   '2006/7/4 MODIFY BY SONIA 加入未作廢控制
   'adoacc1k0.Open "select * from acc1k0 where a1k01 = '" & Text1 & "' and a1k17 is null", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '2011/7/22 modify by sonia 加入已結清a1k29,部分收款a1k30控制(自KeyDefine1移過來)
   adoacc1k0.Open "select * from acc1k0 where a1k01 = '" & Text1 & "' and a1k17 is null and a1k12 is null and a1k29 is null and (a1k30 is null or a1k30=0) ", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc1k0.RecordCount <> 0 Then
      adoacc1k0.Fields("a1k17").Value = Text9
      adoacc1k0.Fields("a1k29").Value = MsgText(602)
        '已收金額(台幣)取小數兩位
'      adoacc1k0.Fields("a1k30").Value = Format(Val(Text2) * Val(IIf(IsNull(adoacc1k0.Fields("a1k08").Value), 0, adoacc1k0.Fields("a1k08").Value)), DAmount)
      '93.9.21 MODIFY BY SONIA
      'adoacc1k0.Fields("a1k30").Value = Format(Val(Text2) * Val(IIf(IsNull(adoacc1k0.Fields("a1k08").Value), 0, adoacc1k0.Fields("a1k08").Value)), 2)
      adoacc1k0.Fields("a1k30").Value = Format(Val(Text2) * Val(IIf(IsNull(adoacc1k0.Fields("a1k08").Value), 0, (adoacc1k0.Fields("a1k08").Value - IIf(IsNull(adoacc1k0.Fields("a1k06").Value), 0, adoacc1k0.Fields("a1k06").Value)))), FAmount)
      '93.9.21 END
        'End
      adoacc1k0.UpdateBatch
      AdodcRefresh1
   Else
      MsgBox MsgText(34), , MsgText(5)
   End If
   adoacc1k0.Close
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  功能鍵定義(1)
'
'*************************************************
Private Sub KeyDefine1(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyInsert
         If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
            Exit Sub
         End If
         
'2011/7/22 cancel by sonia 移到Acc1k0Save
'         'Add by Morgan 2006/8/15
'         strExc(0) = "select * from acc1k0 where a1k01='" & Text1.Text & "'"
'         intI = 1
'         'edit by nickc 2007/02/07 不用 dll 了
'         'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            If "" & RsTemp.Fields("a1k29") = "Y" Then
'               MsgBox "請款單已結清，不可進行抵帳作業！"
'               Exit Sub
'            ElseIf Val("" & RsTemp.Fields("a1k30")) > 0 Then
'               MsgBox "請款單已有部分收款，不可進行抵帳作業！"
'               Exit Sub
'            End If
'         End If
'         'end 2006/8/15
'2011/7/22 end
        'Add by Amy 2014/04/09 +抵帳訊息
        Dim strTp(2) As String '記錄代理人a1k03/列印對象a1k27/請款對象a1k28 抵帳情況
        strExc(0) = "Select a1k03,a1k27,a1k28 From acc1k0 Where a1k01 = '" & Text1 & "' and a1k17 is null " & _
                        "and a1k12 is null and a1k29 is null and (a1k30 is null or a1k30=0) "
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
            strExc(1) = ""
            If Not IsNull(RsTemp("a1k03")) Then
                If GetDizhang(Left(RsTemp("a1k03"), 8), , False, 2, False) = "不同意抵帳" Then
                    strTp(0) = Left(RsTemp("a1k03"), 8)
                End If
            End If
            If Not IsNull(RsTemp("a1k27")) Then
                If GetDizhang(Left(RsTemp("a1k27"), 8), , False, 2, False) = "不同意抵帳" Then
                    strTp(1) = Left(RsTemp("a1k27"), 8)
                End If
            End If
            If Not IsNull(RsTemp("a1k28")) Then
                If GetDizhang(Left(RsTemp("a1k28"), 8), , False, 2, False) = "不同意抵帳" Then
                    strTp(2) = Left(RsTemp("a1k28"), 8)
                End If
            End If
            
            If strTp(0) <> MsgText(601) Then strExc(1) = strTp(0) & ","
            If strTp(1) <> MsgText(601) And InStr(strExc(1), strTp(1)) = 0 Then strExc(1) = strExc(1) & strTp(1) & ","
            If strTp(2) <> MsgText(601) And InStr(strExc(1), strTp(2)) = 0 Then strExc(1) = strExc(1) & strTp(2) & ","
            If strExc(1) <> MsgText(601) Then
                If Len(strExc(1)) = 9 Then
                    strExc(1) = "代理人「" & Left(strExc(1), Len(strExc(1)) - 1) & "」不同意抵帳！"
                Else
                    strExc(1) = "代理人: " & vbCrLf & Left(strExc(1), Len(strExc(1)) - 1) & vbCrLf & "不同意抵帳！"
                End If
                MsgBox strExc(1), , MsgText(5)
                Exit Sub
            End If
            'add by sonia 2021/11/19
            'Modified by Morgan 2022/1/11 有編號才檢查否則第1筆都會彈訊息
            'If "" & RsTemp("a1k03") <> Text13 And "" & RsTemp("a1k28") <> Text13 Then
            If Text13 <> "" And "" & RsTemp("a1k03") <> Text13 And "" & RsTemp("a1k28") <> Text13 Then
            'end 2022/1/11
                If MsgBox("此請款單之代理人及請款對象都與上方代理人不符，仍要做此筆請款單的抵帳嗎？ Y：維持，N：取消設定" & vbCrLf, vbYesNo + vbDefaultButton2) = vbNo Then
                   Exit Sub
                End If
            End If
            'end 2021/11/19
            
        End If
        'end 2014/04/09
         
         Frmacc21f0_Save
         If strControlButton <> MsgText(602) Then
            Acc1k0Save
         End If
         If strControlButton <> MsgText(602) Then
            SumShow1
            Text1.Text = "X"
            'modify by sonia 2015/5/20
            'Text1.SetFocus
            Text1_GotFocus
         End If
         strControlButton = MsgText(601)
   End Select
   KeyEnter KeyCode
End Sub

'*************************************************
'  儲存資料表(國外帳單資料(主檔))
'
'*************************************************
'Modified by Lydia 2024/09/03 strNoList>>U單號
Private Sub Acc150Save(Optional ByVal strNoList As String)
On Error GoTo Checking
   If Text5 = "" Then
      Exit Sub
   End If
   adoacc150.CursorLocation = adUseClient
   '2006/7/4 MODIFY BY SONIA 加入未作廢控制
   'adoacc150.Open "select * from acc150 where a1501 = '" & Text5 & "' and a1512 is null", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '2010/9/30 modify by sonia 加入未付款控制 U09800026已付,此處誤輸再刪除會清掉已付金額
   'adoacc150.Open "select * from acc150 where a1501 = '" & Text5 & "' and a1512 is null and a1507 is null", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '2011/7/22 MODIFY BY SONIA再加入是否審核A1521控管
   'Modified by Lydia 2024/09/03 a1501 = '" & Text5 & "'>>
   'adoacc150.Open "select * from acc150 where a1501 = '" & Text5 & "' and a1512 is null and a1507 is null AND (A1520 IS NULL OR A1520=0) AND (A1521 IS NULL OR A1521<>'N') ", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc150.Open "select * from acc150 where " & IIf(strNoList = "", "a1501 = '" & Text5 & "'", "a1501 in (" & GetAddStr(strNoList) & ") ") & " and a1512 is null and a1507 is null AND (A1520 IS NULL OR A1520=0) AND (A1521 IS NULL OR A1521<>'N') order by a1501 ", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc150.RecordCount <> 0 Then
      'Added by Lydia 2024/09/03
      adoacc150.MoveFirst
      Do While Not adoacc150.EOF
      'end 2024/09/03
         adoacc150.Fields("a1512").Value = Text9
         If Text6 <> MsgText(601) Then
            adoacc150.Fields("a1513").Value = Val(Text6)
            If IsNull(adoacc150.Fields("a1506").Value) Then
               adoacc150.Fields("a1510").Value = 0
               adoacc150.Fields("a1520").Value = 0
            Else
               adoacc150.Fields("a1510").Value = Val(Format(Val(Text6) * Val(adoacc150.Fields("a1506").Value), DAmount))
               adoacc150.Fields("a1520").Value = Val(adoacc150.Fields("a1506").Value)
            End If
         Else
            adoacc150.Fields("a1513").Value = 0
         End If
      'Added by Lydia 2024/09/03
         adoacc150.MoveNext
      Loop
      'end 2024/09/03
      adoacc150.UpdateBatch
      AdodcRefresh2
   Else
      MsgBox MsgText(34), , MsgText(5)
   End If
   adoacc150.Close
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  功能鍵定義(2)
'
'*************************************************
Private Sub KeyDefine2(KeyCode As Integer)
    Dim strCompNo As String 'Add by Amy 2015/07/22
    
   Select Case KeyCode
      Case vbKeyInsert
         If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
            Exit Sub
         End If
        
        'Add by Amy 2014/04/09 +抵帳訊息
        'Modify by Amy 2014/07/22 +抓acc151.axf03 for 抵帳不可抵J公司資料
        'Modified by Morgan 2016/7/29 因增加W(待審核)已審核改判斷 A1521='Y'

        'strExc(0) = "Select * From acc150,acc151 Where a1501 = '" & Text5 & "' and a1512 is null and a1507 is null " & _
                         "And (A1520 IS NULL OR A1520=0) AND (A1521 IS NULL OR A1521<>'N') And a1501=axf01 "
        strExc(0) = "Select * From acc150,acc151 Where a1501 = '" & Text5 & "' and a1512 is null and a1507 is null " & _
                         "And (A1520 IS NULL OR A1520=0) AND (A1521 IS NULL OR A1521='Y') And a1501=axf01 "
        'end 2016/7/29
        
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
            'Add by Amy 2015/07/22 +輸J公司抵帳時不可insert
            If Not IsNull(RsTemp("axf03")) Then
                 strExc(0) = GetSpecialComp(Mid(RsTemp("axf03"), 1, Len(RsTemp("axf03")) - 9), Mid(RsTemp("axf03"), (Len(RsTemp("axf03")) - 9) + 1, 6), Mid(RsTemp("axf03"), (Len(RsTemp("axf03")) - 3) + 1, 1), Mid(RsTemp("axf03"), (Len(RsTemp("axf03")) - 2) + 1, 2), strCompNo, 0)
                If strCompNo = "J" Then
                    MsgBox "抵帳不可為J公司！"
                    Exit Sub
                End If
            End If
            'end 2015/07/22
            If Not IsNull(RsTemp("a1503")) Then
                If GetDizhang(Left(RsTemp("a1503"), 8), , False, 2, False) = "不同意抵帳" Then
                    MsgBox "代理人「" & Left(RsTemp("a1503"), 8) & "」不同意抵帳！"
                    Exit Sub
                End If
            End If
            'add by sonia 2021/11/19
            'Modified by Morgan 2022/1/11 有編號才檢查否則第1筆都會彈訊息
            'If "" & RsTemp("a1503") <> Text13 Then
            If Text13 <> "" And "" & RsTemp("a1503") <> Text13 Then
            'end 2022/1/11
                If MsgBox("此帳單之代理人與上方代理人不符，仍要做此筆帳單的抵帳嗎？ Y：維持，N：取消設定" & vbCrLf, vbYesNo + vbDefaultButton2) = vbNo Then
                   Exit Sub
                End If
            End If
            'end 2021/11/19
                    
            'Added by Morgan 2023/10/23 排除結匯中的帳單
            strExc(0) = "select * from acc170 where a1702='" & Text5 & "' and a1701='1'"
            intI = 1
            Set adocheck = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MsgBox "此帳單已有結匯資料，不可抵帳！", vbCritical
               Exit Sub
            End If
            adocheck.Close
            'end 2023/10/29
        End If
        'end 2014/04/09
        
         Frmacc21f0_Save
         If strControlButton <> MsgText(602) Then
            Acc150Save
         End If
         If strControlButton <> MsgText(602) Then
            SumShow2
            Text5 = "U"
            'modify by sonia 2015/5/20
            'Text5.SetFocus
            Text5_GotFocus
         End If
         strControlButton = MsgText(601)
   End Select
   KeyEnter KeyCode
End Sub

'*************************************************
'  計算並顯示 Adodc1 之合計
'
'*************************************************
Public Sub SumShow1()
   adoaccsum.CursorLocation = adUseClient
   '93.9.21 MODIFY BY SONIA
   'adoaccsum.Open "select sum(a1k08), sum(round(a1k08 * " & Val(Text2) & ", 2)) from acc1k0 where a1k17 = '" & Text9 & "'", adoTaie, adOpenStatic, adLockReadOnly
   adoaccsum.Open "select sum(a1k08-nvl(a1k06, 0)), sum(round((a1k08-nvl(a1k06, 0)) * " & Val(Text2) & ", 2)) from acc1k0 where a1k17 = '" & Text9 & "'", adoTaie, adOpenStatic, adLockReadOnly
   '93.9.21 END
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text3 = MsgText(601)
      Else
         Text3 = Format(adoaccsum.Fields(0).Value, FAmount)
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text4 = MsgText(601)
      Else
         Text4 = Format(adoaccsum.Fields(1).Value, FAmount)
      End If
   Else
      Text3 = MsgText(601)
      Text4 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  計算並顯示 Adodc2 之合計
'
'*************************************************
Public Sub SumShow2()
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(round(axf04 * '" & Val(Text6) & "')), sum(axf04) from acc151, acc150 where axf01 = a1501 and a1512 = '" & Text9 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text7 = MsgText(601)
      Else
         Text7 = Format(adoaccsum.Fields(0).Value, DAmount)
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text8 = MsgText(601)
      Else
         Text8 = Format(adoaccsum.Fields(1).Value, FAmount)
      End If
   Else
     Text7 = MsgText(601)
     Text8 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  刪除 Adodc1 之資料
'
'*************************************************
Private Sub AdodcDelete1()
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   adoTaie.Execute "update acc1k0 set a1k17 = null, a1k29 = null, a1k30 = 0 where a1k01 = '" & Adodc1.Recordset.Fields("a1k01").Value & "'"
   AdodcRefresh1
   SumShow1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  刪除 Adodc2 之資料
'
'*************************************************
Private Sub AdodcDelete2()
On Error GoTo Checking
   If Adodc2.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   adoTaie.Execute "update acc150 set a1512 = null, a1520 = 0 where a1501 = '" & Adodc2.Recordset.Fields("a1501").Value & "'"
   AdodcRefresh2
   SumShow2
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = adoacc1g0.Bookmark & MsgText(35) & adoacc1g0.RecordCount
End Sub

'*************************************************
'  關閉分錄欄位輸入狀態
'
'*************************************************
Public Sub FormDisabled()
   Text1.Enabled = False
   Text5.Enabled = False
   Command1.Enabled = False
   Command2.Enabled = True
   Command3.Enabled = False
End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'
'*************************************************
Public Sub FormEnabled()
   Text1.Enabled = True
   Text5.Enabled = True
   Command1.Enabled = True
   Command2.Enabled = False
   Command3.Enabled = True
End Sub

'*************************************************
'  列印抵帳資料
'
'*************************************************
Public Sub PrintDataf2()
Dim strAmount As String
Dim intLength As Integer
Dim strCurrency As String
'Add By Cheng 2003/05/27
Dim strCaseNo As String '本所案號
Dim strFaNo As String '代理人編號

   intCounter = 0
   m_intPage = 1
   Printer.FontSize = 12
   '帳單資料
   adoquery.CursorLocation = adUseClient
    'Modify By Cheng 2003/05/21
'   adoquery.Open "select * from acc150 where a1512 = '" & strItemNo & "' order by a1504 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoquery.Open "select * from acc151, acc150 where axf01 = a1501 and a1512 = '" & strItemNo & "' order by a1504 asc", adoTaie, adOpenStatic, adLockReadOnly
    'Add By Cheng 2003/05/21
    strCaseNo = "" & adoquery.Fields("axf03").Value
    strFaNo = "" & adoquery.Fields("a1503").Value
    PrintHeadf2 strCaseNo, strFaNo
   Printer.CurrentX = 0 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Your Debit Notes"
    'Add By Cheng 2003/05/21
   Printer.CurrentX = 2000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Our Ref"
   Printer.CurrentX = 5000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Date"
   If IsNull(adoquery.Fields("a1505").Value) Then
      strCurrency = ""
   Else
      strCurrency = adoquery.Fields("a1505").Value
   End If
   Printer.CurrentX = 7000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Amount(" & strCurrency & ")"
   intCounter = intCounter + 1
   Printer.Line (0 + m_dblLeft, 0 + intCounter * 300 - 50)-(9000 + m_dblLeft, 0 + intCounter * 300 - 50)
   Do While adoquery.EOF = False
      If intCounter > 48 Then
        Printer.CurrentX = 5000
        Printer.CurrentY = 0 + intCounter * 300
        Printer.Print "**" & m_intPage & "**"
        m_intPage = m_intPage + 1
         Printer.NewPage
        'Add By Cheng 2003/05/27
        PrintHeadf2 strCaseNo, strFaNo
        Printer.CurrentX = 0 + m_dblLeft
        Printer.CurrentY = 0 + intCounter * 300
        Printer.Print "Your Debit Notes"
         'Add By Cheng 2003/05/21
        Printer.CurrentX = 2000 + m_dblLeft
        Printer.CurrentY = 0 + intCounter * 300
        Printer.Print "Our Ref"
        Printer.CurrentX = 5000 + m_dblLeft
        Printer.CurrentY = 0 + intCounter * 300
        Printer.Print "Date"
        If IsNull(adoquery.Fields("a1505").Value) Then
           strCurrency = ""
        Else
           strCurrency = adoquery.Fields("a1505").Value
        End If
        Printer.CurrentX = 7000 + m_dblLeft
        Printer.CurrentY = 0 + intCounter * 300
        Printer.Print "Amount(" & strCurrency & ")"
        intCounter = intCounter + 1
        Printer.Line (0 + m_dblLeft, 0 + intCounter * 300 - 50)-(9000 + m_dblLeft, 0 + intCounter * 300 - 50)
      End If
      Printer.CurrentX = 0 + m_dblLeft
      Printer.CurrentY = 0 + intCounter * 300
      If IsNull(adoquery.Fields("a1504").Value) Then
         Printer.Print ""
      Else
         Printer.Print adoquery.Fields("a1504").Value
      End If
        'Add By Cheng 2003/05/21
      Printer.CurrentX = 2000 + m_dblLeft
      Printer.CurrentY = 0 + intCounter * 300
      If IsNull(adoquery.Fields("axf03").Value) Then
         Printer.Print ""
      Else
         Printer.Print adoquery.Fields("axf03").Value
      End If
      Printer.CurrentX = 5000 + m_dblLeft
      Printer.CurrentY = 0 + intCounter * 300
      If IsNull(adoquery.Fields("a1502").Value) Then
         Printer.Print ""
      Else
         Printer.Print Format(CADate(adoquery.Fields("a1502").Value), "####-##-##")
      End If
        'Modify By Cheng 2003/05/21
'      If IsNull(adoquery.Fields("a1506").Value) = False Then
'         strAmount = Format(Val(adoquery.Fields("a1506").Value), FDollar)
      If IsNull(adoquery.Fields("axf04").Value) = False Then
         strAmount = Format(Val(adoquery.Fields("axf04").Value), FDollar)
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = 9000 - intLength + m_dblLeft
         Printer.CurrentY = 0 + intCounter * 300
         Printer.Print strAmount
      End If
      intCounter = intCounter + 1
      adoquery.MoveNext
   Loop
   adoquery.Close
   intCounter = intCounter + 1
   Printer.Line (0 + m_dblLeft, 0 + intCounter * 300 - 50)-(9000 + m_dblLeft, 0 + intCounter * 300 - 50)
   Printer.CurrentX = 0 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Total amount due in your favor is"
   Printer.CurrentX = 6000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print strCurrency
   adoaccsum.CursorLocation = adUseClient
    'Modify By Cheng 2003/05/21
'   adoaccsum.Open "select sum(a1506) from acc150 where a1512 = '" & strItemNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   adoaccsum.Open "select sum(axf04) from acc151, acc150 where axf01 = a1501 and a1512 = '" & strItemNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) = False Then
         strAmount = Format(Val(adoaccsum.Fields(0).Value), FDollar)
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = 9000 - intLength + m_dblLeft
         Printer.CurrentY = 0 + intCounter * 300
         Printer.Print strAmount
      End If
   End If
   adoaccsum.Close
    intCounter = intCounter + 1
    Printer.CurrentX = 5000
    Printer.CurrentY = 0 + intCounter * 300
    Printer.Print "**" & m_intPage & "**"
    m_intPage = m_intPage + 1
   
   intCounter = 0
   m_intPage = 1
   Printer.NewPage
   '請款單資料
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from acc1k0 where a1k17 = '" & strItemNo & "' order by a1k01 asc", adoTaie, adOpenStatic, adLockReadOnly
    'Add By Cheng 2003/05/21
    PrintHeadf2 strCaseNo, strFaNo
   Printer.CurrentX = 0 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Our Debit Notes"
   Printer.CurrentX = 2000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Our Ref"
   Printer.CurrentX = 5000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Date"
   If IsNull(adoquery.Fields("a1k18").Value) Then
      strCurrency = ""
   Else
      strCurrency = adoquery.Fields("a1k18").Value
   End If
   Printer.CurrentX = 7000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   '2012/5/17 MODIFY BY SONIA 固定用美金收款 Z10100011
   'Printer.Print "Amount(" & strCurrency & ")"
   Printer.Print "Amount(USD)"
   intCounter = intCounter + 1
   Printer.Line (0 + m_dblLeft, 0 + intCounter * 300 - 50)-(9000 + m_dblLeft, 0 + intCounter * 300 - 50)
   Do While adoquery.EOF = False
      If intCounter > 48 Then
        Printer.CurrentX = 5000
        Printer.CurrentY = 0 + intCounter * 300
        Printer.Print "**" & m_intPage & "**"
        m_intPage = m_intPage + 1
         Printer.NewPage
        'Add By Cheng 2003/05/27
        PrintHeadf2 strCaseNo, strFaNo
        Printer.CurrentX = 0 + m_dblLeft
        Printer.CurrentY = 0 + intCounter * 300
        Printer.Print "Our Debit Notes"
        Printer.CurrentX = 2000 + m_dblLeft
        Printer.CurrentY = 0 + intCounter * 300
        Printer.Print "Our Ref"
        Printer.CurrentX = 5000 + m_dblLeft
        Printer.CurrentY = 0 + intCounter * 300
        Printer.Print "Date"
        If IsNull(adoquery.Fields("a1k18").Value) Then
           strCurrency = ""
        Else
           strCurrency = adoquery.Fields("a1k18").Value
        End If
        Printer.CurrentX = 7000 + m_dblLeft
        Printer.CurrentY = 0 + intCounter * 300
        '2012/5/17 MODIFY BY SONIA 固定用美金收款 Z10100011
        'Printer.Print "Amount(" & strCurrency & ")"
        Printer.Print "Amount(USD)"
        intCounter = intCounter + 1
        Printer.Line (0 + m_dblLeft, 0 + intCounter * 300 - 50)-(9000 + m_dblLeft, 0 + intCounter * 300 - 50)
      End If
      Printer.CurrentX = 0 + m_dblLeft
      Printer.CurrentY = 0 + intCounter * 300
      If IsNull(adoquery.Fields("a1k01").Value) Then
         Printer.Print ""
      Else
         '2009/7/2 MODIFY BY SONIA 改印完整編號
         'Printer.Print Mid(adoquery.Fields("a1k01").Value, 2, Len(adoquery.Fields("a1k01").Value) - 1)
         Printer.Print adoquery.Fields("a1k01").Value
      End If
      'Add By Cheng 2003/05/21
      Printer.CurrentX = 2000 + m_dblLeft
      Printer.CurrentY = 0 + intCounter * 300
      If IsNull(adoquery.Fields("a1k13").Value) Then
         Printer.Print ""
      Else
         Printer.Print adoquery.Fields("a1k13").Value & adoquery.Fields("a1k14").Value & adoquery.Fields("a1k15").Value * adoquery.Fields("a1k16").Value
      End If
      Printer.CurrentX = 5000 + m_dblLeft
      Printer.CurrentY = 0 + intCounter * 300
      If IsNull(adoquery.Fields("a1k02").Value) Then
         Printer.Print ""
      Else
         Printer.Print Format(CADate(adoquery.Fields("a1k02").Value), "####-##-##")
      End If
      If IsNull(adoquery.Fields("a1k08").Value) = False Then
         strAmount = Format(Val(adoquery.Fields("a1k08").Value) - Val(IIf(IsNull(adoquery.Fields("a1k06").Value), 0, adoquery.Fields("a1k06").Value)), FDollar)
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = 9000 - intLength + m_dblLeft
         Printer.CurrentY = 0 + intCounter * 300
         Printer.Print strAmount
      End If
      intCounter = intCounter + 1
      adoquery.MoveNext
   Loop
   adoquery.Close
   intCounter = intCounter + 1
   Printer.Line (0 + m_dblLeft, 0 + intCounter * 300 - 50)-(9000 + m_dblLeft, 0 + intCounter * 300 - 50)
   Printer.CurrentX = 0 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Total amount due in our favor is"
   Printer.CurrentX = 6000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   '2012/5/17 MODIFY BY SONIA 固定用美金收款 Z10100011
   'Printer.Print strCurrency
   Printer.Print "USD"
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(a1k08 - nvl(a1k06, 0)) from acc1k0 where a1k17 = '" & strItemNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) = False Then
         strAmount = Format(Val(adoaccsum.Fields(0).Value), FDollar)
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = 9000 - intLength + m_dblLeft
         Printer.CurrentY = 0 + intCounter * 300
         Printer.Print strAmount
      End If
   End If
   adoaccsum.Close
    intCounter = intCounter + 1
    Printer.CurrentX = 5000
    Printer.CurrentY = 0 + intCounter * 300
    Printer.Print "**" & m_intPage & "**"
    m_intPage = m_intPage + 1
   Printer.EndDoc
End Sub

'*************************************************
'  列印抵帳資料
'
'*************************************************
Public Sub PrintDataf1()
Dim strAmount As String
Dim intLength As Integer
Dim strCurrency As String
'Add By Cheng 2003/05/21
Dim strCaseNo As String '本所案號
Dim strFaNo As String '代理人編號

   intCounter = 0
   m_intPage = 1
   Printer.FontSize = 12
   '帳單資料
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from acc151, acc150 where axf01 = a1501 and a1512 = '" & strItemNo & "' order by a1504 asc", adoTaie, adOpenStatic, adLockReadOnly
    'Add By Cheng 2003/05/21
    strCaseNo = "" & adoquery.Fields("axf03").Value
    strFaNo = "" & adoquery.Fields("a1503").Value
    PrintHeadf1 strCaseNo, strFaNo
   Printer.CurrentX = 0 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Your Debit Notes"
   Printer.CurrentX = 2000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Our Ref"
   Printer.CurrentX = 5000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Date"
   If IsNull(adoquery.Fields("a1505").Value) Then
      strCurrency = ""
   Else
      strCurrency = adoquery.Fields("a1505").Value
   End If
   Printer.CurrentX = 7000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Amount(" & strCurrency & ")"
   intCounter = intCounter + 1
   Printer.Line (0 + m_dblLeft, 0 + intCounter * 300 - 50)-(9000 + m_dblLeft, 0 + intCounter * 300 - 50)
   Do While adoquery.EOF = False
        If intCounter > 48 Then
              Printer.CurrentX = 5000
              Printer.CurrentY = 0 + intCounter * 300
              Printer.Print "**" & m_intPage & "**"
              m_intPage = m_intPage + 1
              Printer.NewPage
            'Add By Cheng 2003/05/27
            PrintHeadf1 strCaseNo, strFaNo
            Printer.CurrentX = 0 + m_dblLeft
            Printer.CurrentY = 0 + intCounter * 300
            Printer.Print "Your Debit Notes"
            Printer.CurrentX = 2000 + m_dblLeft
            Printer.CurrentY = 0 + intCounter * 300
            Printer.Print "Our Ref"
            Printer.CurrentX = 5000 + m_dblLeft
            Printer.CurrentY = 0 + intCounter * 300
            Printer.Print "Date"
            If IsNull(adoquery.Fields("a1505").Value) Then
               strCurrency = ""
            Else
               strCurrency = adoquery.Fields("a1505").Value
            End If
            Printer.CurrentX = 7000 + m_dblLeft
            Printer.CurrentY = 0 + intCounter * 300
            Printer.Print "Amount(" & strCurrency & ")"
            intCounter = intCounter + 1
            Printer.Line (0 + m_dblLeft, 0 + intCounter * 300 - 50)-(9000 + m_dblLeft, 0 + intCounter * 300 - 50)
        End If
      Printer.CurrentX = 0 + m_dblLeft
      Printer.CurrentY = 0 + intCounter * 300
      If IsNull(adoquery.Fields("a1504").Value) Then
         Printer.Print ""
      Else
         Printer.Print adoquery.Fields("a1504").Value
      End If
      Printer.CurrentX = 2000 + m_dblLeft
      Printer.CurrentY = 0 + intCounter * 300
      If IsNull(adoquery.Fields("axf03").Value) Then
         Printer.Print ""
      Else
         Printer.Print adoquery.Fields("axf03").Value
      End If
      Printer.CurrentX = 5000 + m_dblLeft
      Printer.CurrentY = 0 + intCounter * 300
      If IsNull(adoquery.Fields("a1502").Value) Then
         Printer.Print ""
      Else
         Printer.Print Format(CADate(adoquery.Fields("a1502").Value), "####-##-##")
      End If
      If IsNull(adoquery.Fields("axf04").Value) = False Then
         strAmount = Format(Val(adoquery.Fields("axf04").Value), FDollar)
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = 9000 - intLength + m_dblLeft
         Printer.CurrentY = 0 + intCounter * 300
         Printer.Print strAmount
      End If
      intCounter = intCounter + 1
      adoquery.MoveNext
   Loop
   adoquery.Close
   intCounter = intCounter + 1
   Printer.Line (0 + m_dblLeft, 0 + intCounter * 300 - 50)-(9000 + m_dblLeft, 0 + intCounter * 300 - 50)
   Printer.CurrentX = 0 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Total amount due in your favor is"
   Printer.CurrentX = 6000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print strCurrency
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(axf04) from acc151, acc150 where axf01 = a1501 and a1512 = '" & strItemNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) = False Then
         strAmount = Format(Val(adoaccsum.Fields(0).Value), FDollar)
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = 9000 - intLength + m_dblLeft
         Printer.CurrentY = 0 + intCounter * 300
         Printer.Print strAmount
      End If
   End If
   adoaccsum.Close
    intCounter = intCounter + 1
    Printer.CurrentX = 5000
    Printer.CurrentY = 0 + intCounter * 300
    Printer.Print "**" & m_intPage & "**"
    m_intPage = m_intPage + 1
   
    m_intPage = 1
    intCounter = 0
    Printer.NewPage
    '請款單資料
    adoquery.CursorLocation = adUseClient
    adoquery.Open "select * from acc1k0 where a1k17 = '" & strItemNo & "' order by a1k01 asc", adoTaie, adOpenStatic, adLockReadOnly
    'Add By Cheng 2003/05/21
    strCaseNo = "" & adoquery.Fields("a1k13").Value & adoquery.Fields("a1k14").Value & adoquery.Fields("a1k15").Value & adoquery.Fields("a1k16").Value
    '2010/6/18 MODIFY BY SONIA 婧瑄說應改為請款對象
    'strFANo = "" & adoquery.Fields("a1k03").Value
    strFaNo = "" & adoquery.Fields("a1k03").Value
    '2010/6/19 END
    PrintHeadf1 strCaseNo, strFaNo
   Printer.CurrentX = 0 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Our Debit Notes"
   Printer.CurrentX = 2000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Our Ref"
   Printer.CurrentX = 5000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Date"
   If IsNull(adoquery.Fields("a1k18").Value) Then
      strCurrency = ""
   Else
      strCurrency = adoquery.Fields("a1k18").Value
   End If
   Printer.CurrentX = 7000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   '2012/5/17 MODIFY BY SONIA 固定用美金收款 Z10100011
   'Printer.Print "Amount(" & strCurrency & ")"
   Printer.Print "Amount(USD)"
   intCounter = intCounter + 1
   Printer.Line (0 + m_dblLeft, 0 + intCounter * 300 - 50)-(9000 + m_dblLeft, 0 + intCounter * 300 - 50)
   Do While adoquery.EOF = False
        If intCounter > 48 Then
              Printer.CurrentX = 5000
              Printer.CurrentY = 0 + intCounter * 300
              Printer.Print "**" & m_intPage & "**"
              m_intPage = m_intPage + 1
              Printer.NewPage
            'Add By Cheng 2003/05/27
            PrintHeadf1 strCaseNo, strFaNo
            Printer.CurrentX = 0 + m_dblLeft
            Printer.CurrentY = 0 + intCounter * 300
            Printer.Print "Our Debit Notes"
            Printer.CurrentX = 2000 + m_dblLeft
            Printer.CurrentY = 0 + intCounter * 300
            Printer.Print "Our Ref"
            Printer.CurrentX = 5000 + m_dblLeft
            Printer.CurrentY = 0 + intCounter * 300
            Printer.Print "Date"
            If IsNull(adoquery.Fields("a1k18").Value) Then
                strCurrency = ""
            Else
                strCurrency = adoquery.Fields("a1k18").Value
            End If
            Printer.CurrentX = 7000 + m_dblLeft
            Printer.CurrentY = 0 + intCounter * 300
            '2012/5/17 MODIFY BY SONIA 固定用美金收款 Z10100011
            'Printer.Print "Amount(" & strCurrency & ")"
            Printer.Print "Amount(USD)"
            intCounter = intCounter + 1
            Printer.Line (0 + m_dblLeft, 0 + intCounter * 300 - 50)-(9000 + m_dblLeft, 0 + intCounter * 300 - 50)
        End If
      Printer.CurrentX = 0 + m_dblLeft
      Printer.CurrentY = 0 + intCounter * 300
      If IsNull(adoquery.Fields("a1k01").Value) Then
         Printer.Print ""
      Else
         'Modify by Morgan 2004/8/10
         '改印完整編號
         'Printer.Print Mid(adoquery.Fields("a1k01").Value, 2, Len(adoquery.Fields("a1k01").Value) - 1)
         Printer.Print adoquery.Fields("a1k01").Value
      End If
      Printer.CurrentX = 2000 + m_dblLeft
      Printer.CurrentY = 0 + intCounter * 300
      If IsNull(adoquery.Fields("a1k13").Value) Then
         Printer.Print ""
      Else
         Printer.Print adoquery.Fields("a1k13").Value & adoquery.Fields("a1k14").Value & adoquery.Fields("a1k15").Value & adoquery.Fields("a1k16").Value
      End If
      Printer.CurrentX = 5000 + m_dblLeft
      Printer.CurrentY = 0 + intCounter * 300
      If IsNull(adoquery.Fields("a1k02").Value) Then
         Printer.Print ""
      Else
         Printer.Print Format(CADate(adoquery.Fields("a1k02").Value), "####-##-##")
      End If
      If IsNull(adoquery.Fields("a1k08").Value) = False Then
         strAmount = Format(Val(adoquery.Fields("a1k08").Value) - Val(IIf(IsNull(adoquery.Fields("a1k06").Value), 0, adoquery.Fields("a1k06").Value)), FDollar)
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = 9000 - intLength + m_dblLeft
         Printer.CurrentY = 0 + intCounter * 300
         Printer.Print strAmount
      End If
      intCounter = intCounter + 1
      adoquery.MoveNext
   Loop
   adoquery.Close
   intCounter = intCounter + 1
   Printer.Line (0 + m_dblLeft, 0 + intCounter * 300 - 50)-(9000 + m_dblLeft, 0 + intCounter * 300 - 50)
   Printer.CurrentX = 0 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Total amount due in our favor is"
   Printer.CurrentX = 6000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   '2012/5/17 MODIFY BY SONIA 固定用美金收款 Z10100011
   'Printer.Print strCurrency
   Printer.Print "USD"
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(a1k08 - nvl(a1k06, 0)) from acc1k0 where a1k17 = '" & strItemNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) = False Then
         strAmount = Format(Val(adoaccsum.Fields(0).Value), FDollar)
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = 9000 - intLength + m_dblLeft
         Printer.CurrentY = 0 + intCounter * 300
         Printer.Print strAmount
      End If
   End If
   adoaccsum.Close
    intCounter = intCounter + 1
    Printer.CurrentX = 5000
    Printer.CurrentY = 0 + intCounter * 300
    Printer.Print "**" & m_intPage & "**"
    m_intPage = m_intPage + 1
   Printer.EndDoc
End Sub

'*************************************************
'  抬頭列印
'
'*************************************************
Private Sub PrintHeadf2(strCaseNo As String, strFaNo As String)
Dim intRow As Integer
Dim StrSQLa As String
Dim strLanguage As String
   
    intRow = 0
    strLanguage = ""       '2012/2/22 ADD BY SONIA 案件無定稿語文抓代理人的定稿語文
    
    adocheck.CursorLocation = adUseClient
    adocheck.Open "select pa85 as Lang from patent, customer where substr(pa26, 1, 8) = cu01 and substr(pa26, 9, 1) = cu02 and " & ChgPatent(strCaseNo) & _
                  " union select tm53 as Lang from trademark, customer where substr(tm23, 1, 8) = cu01 and substr(tm23, 9, 1) = cu02 and " & ChgTradeMark(strCaseNo) & _
                  " union select sp34 as Lang from servicepractice, customer where substr(sp08, 1, 8) = cu01 and substr(sp08, 9, 1) = cu02 and " & ChgService(strCaseNo), adoTaie, adOpenStatic, adLockReadOnly
    If adocheck.RecordCount <> 0 Then
        If IsNull(adocheck.Fields("Lang").Value) = False Then
           strLanguage = adocheck.Fields("Lang").Value
        '2012/2/22 CANCEL BY SONIA
        'Else
        '   strLanguage = "2"
        '2012/2/22 END
        End If
    Else
        strLanguage = "2"
    End If
    adocheck.Close
    Printer.CurrentX = 7000 + m_dblLeft
    Printer.CurrentY = 0 + intRow * 300
'    Printer.Print Format(AFDate(ServerDate), "mmm. d, yyyy")
    intRow = intRow + 1
    adocheck.CursorLocation = adUseClient
    StrSQLa = "Select * From Fagent Where FA01='" & Mid(strFaNo, 1, 8) & "' And FA02='" & Mid(strFaNo, 9, 1) & "' "
    adocheck.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
    If adocheck.RecordCount > 0 Then
        '2012/2/22 ADD BY SONIA 案件無定稿語文抓代理人的定稿語文
        If strLanguage = "" Then
           If IsNull(adocheck.Fields("fa31").Value) = False Then strLanguage = adocheck.Fields("fa31").Value
        End If
        '2012/2/22 END
        
        Select Case strLanguage
           Case "2"
              If IsNull(adocheck.Fields("fa05").Value) = False Then
                 If m_intPage = 1 Then
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa05").Value
                 End If
              Else '2012/2/22 ADD BY SONIA 無英文印中文
                 If m_intPage = 1 Then
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa04").Value
                 End If
              End If
              If IsNull(adocheck.Fields("fa63").Value) = False Then
                 If m_intPage = 1 Then
                    intRow = intRow + 1
                    intCounter = intCounter + 1
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa63").Value
                 End If
              End If
              If IsNull(adocheck.Fields("fa64").Value) = False Then
                 If m_intPage = 1 Then
                    intRow = intRow + 1
                    intCounter = intCounter + 1
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa64").Value
                 End If
              End If
              If IsNull(adocheck.Fields("fa65").Value) = False Then
                 If m_intPage = 1 Then
                    intRow = intRow + 1
                    intCounter = intCounter + 1
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa65").Value
                 End If
              End If
              intRow = intRow + 1
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa32").Value) Then
                    If IsNull(adocheck.Fields("fa18").Value) = False Then
                       Printer.CurrentX = 0 + m_dblLeft
                       Printer.CurrentY = 0 + intRow * 300
                       Printer.Print adocheck.Fields("fa18").Value
                    'Add By Cheng 2003/03/26
                    '若無英文地址時,  印中文地址
                    Else
                       Printer.CurrentX = 0 + m_dblLeft
                       Printer.CurrentY = 0 + intRow * 300
                       Printer.Print "" & adocheck.Fields("fa17").Value
                    End If
                 Else
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa32").Value
                 End If
              End If
              intRow = intRow + 1
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa32").Value) Then
                    If IsNull(adocheck.Fields("fa19").Value) = False Then
                       Printer.CurrentX = 0 + m_dblLeft
                       Printer.CurrentY = 0 + intRow * 300
                       Printer.Print adocheck.Fields("fa19").Value
                    End If
                 Else
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print "" & adocheck.Fields("fa33").Value
                 End If
              End If
              intRow = intRow + 1
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa32").Value) Then
                    If IsNull(adocheck.Fields("fa20").Value) = False Then
                       Printer.CurrentX = 0 + m_dblLeft
                       Printer.CurrentY = 0 + intRow * 300
                       Printer.Print adocheck.Fields("fa20").Value
                    End If
                 Else
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print "" & adocheck.Fields("fa34").Value
                 End If
              End If
              intRow = intRow + 1
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa32").Value) Then
                    If IsNull(adocheck.Fields("fa21").Value) = False Then
                       Printer.CurrentX = 0 + m_dblLeft
                       Printer.CurrentY = 0 + intRow * 300
                       Printer.Print adocheck.Fields("fa21").Value
                    End If
                 Else
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print "" & adocheck.Fields("fa35").Value
                 End If
              End If
              intRow = intRow + 1
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa32").Value) Then
                    If IsNull(adocheck.Fields("fa22").Value) = False Then
                       Printer.CurrentX = 0 + m_dblLeft
                       Printer.CurrentY = 0 + intRow * 300
                       Printer.Print adocheck.Fields("fa22").Value
                    End If
                 Else
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print "" & adocheck.Fields("fa36").Value
                 End If
              End If
              
              'Add by Morgan 2011/5/25
              '英文地址6
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa32").Value) Then
                    If IsNull(adocheck.Fields("fa70").Value) = False Then
                       intRow = intRow + 1
                       Printer.CurrentX = 0 + m_dblLeft
                       Printer.CurrentY = 0 + intRow * 300
                       Printer.Print adocheck.Fields("fa70").Value
                    End If
                 End If
              End If
              
           Case "3"
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa06").Value) = False Then
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa06").Value
                 End If
              End If
              intRow = intRow + 1
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa23").Value) = False Then
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa23").Value
                 End If
              End If
              
           '2012/2/22 ADD BY SONIA Y47804無英文
           Case "1"
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa04").Value) = False Then
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa04").Value
                 End If
              End If
              intRow = intRow + 1
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa17").Value) = False Then
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa17").Value
                 End If
              End If
              intRow = intRow + 1
           '2012/2/22 END
        End Select
    End If
    adocheck.Close
    intRow = intRow + 1
    intCounter = intRow

End Sub

'*************************************************
'  抬頭列印
'
'*************************************************
Private Sub PrintHeadf1(strCaseNo As String, strFaNo As String)
Dim intRow As Integer
Dim StrSQLa As String
Dim strLanguage As String
   
    intRow = 0
    strLanguage = ""       '2012/2/22 ADD BY SONIA 案件無定稿語文抓代理人的定稿語文
    
    adocheck.CursorLocation = adUseClient
    adocheck.Open "select pa85 as Lang from patent, customer where substr(pa26, 1, 8) = cu01 and substr(pa26, 9, 1) = cu02 and " & ChgPatent(strCaseNo) & _
                  " union select tm53 as Lang from trademark, customer where substr(tm23, 1, 8) = cu01 and substr(tm23, 9, 1) = cu02 and " & ChgTradeMark(strCaseNo) & _
                  " union select sp34 as Lang from servicepractice, customer where substr(sp08, 1, 8) = cu01 and substr(sp08, 9, 1) = cu02 and " & ChgService(strCaseNo), adoTaie, adOpenStatic, adLockReadOnly
    If adocheck.RecordCount <> 0 Then
        If IsNull(adocheck.Fields("Lang").Value) = False Then
           strLanguage = adocheck.Fields("Lang").Value
        '2012/2/22 CANCEL BY SONIA
        'Else
        '   strLanguage = "2"
        '2012/2/22 END
        End If
    Else
        strLanguage = "2"
    End If
    adocheck.Close
    Printer.CurrentX = 7000 + m_dblLeft
    Printer.CurrentY = 0 + intRow * 300
'    Printer.Print Format(AFDate(ServerDate), "mmm. d, yyyy")
    intRow = intRow + 1
    adocheck.CursorLocation = adUseClient
    'Modify By Cheng 2004/02/27
'    strSQLA = "Select * From Fagent Where FA01='" & Mid(strFANo, 1, 8) & "' And FA02='" & Mid(strFANo, 9, 1) & "' "
    '2012/2/22 modify by sonia 加fa31,cu64
    StrSQLa = "Select FA04, FA05, FA63, FA64, FA65, FA06, FA17, FA18, FA19, FA20, FA21, FA22, FA70, FA32, FA33, FA34, FA35, FA36, FA23, FA31 From Fagent Where FA01='" & Mid(strFaNo, 1, 8) & "' And FA02='" & Mid(strFaNo, 9, 1) & "' "
    StrSQLa = StrSQLa & " Union Select CU04 As FA04, CU05 As FA05, CU88 As FA63, CU89 As FA64, CU90 As FA65, CU06 As FA06, CU23 As FA17, CU24 As FA18, CU25 As FA19, CU26 As FA20, CU27 As FA21, CU28 As FA22, Cu102 As FA70, CU65 As FA32, CU66 As FA33, CU67 As FA34, CU68 As FA35, CU69 As FA36, CU29 As FA23, CU64 As FA31 From Customer Where CU01='" & Mid(strFaNo, 1, 8) & "' And CU02='" & Mid(strFaNo, 9, 1) & "' "
    'End
    adocheck.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
    If adocheck.RecordCount > 0 Then
        '2012/2/22 ADD BY SONIA 案件無定稿語文抓代理人的定稿語文
        If strLanguage = "" Then
           If IsNull(adocheck.Fields("fa31").Value) = False Then strLanguage = adocheck.Fields("fa31").Value
        End If
        '2012/2/22 END
        
        Select Case strLanguage
           Case "2"
              If IsNull(adocheck.Fields("fa05").Value) = False Then
                 If m_intPage = 1 Then
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa05").Value
                 End If
              Else '無英文印中文
                 If m_intPage = 1 Then
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa04").Value
                 End If
              End If
              If IsNull(adocheck.Fields("fa63").Value) = False Then
                 If m_intPage = 1 Then
                    intRow = intRow + 1
                    intCounter = intCounter + 1
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa63").Value
                 End If
              End If
              If IsNull(adocheck.Fields("fa64").Value) = False Then
                 If m_intPage = 1 Then
                    intRow = intRow + 1
                    intCounter = intCounter + 1
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa64").Value
                 End If
              End If
              If IsNull(adocheck.Fields("fa65").Value) = False Then
                 If m_intPage = 1 Then
                    intRow = intRow + 1
                    intCounter = intCounter + 1
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa65").Value
                 End If
              End If
              intRow = intRow + 1
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa32").Value) Then
                    If IsNull(adocheck.Fields("fa18").Value) = False Then
                       Printer.CurrentX = 0 + m_dblLeft
                       Printer.CurrentY = 0 + intRow * 300
                       Printer.Print adocheck.Fields("fa18").Value
                    'Add By Cheng 2003/03/26
                    '若無英文地址時,  印中文地址
                    Else
                       Printer.CurrentX = 0 + m_dblLeft
                       Printer.CurrentY = 0 + intRow * 300
                       Printer.Print "" & adocheck.Fields("fa17").Value
                    End If
                 Else
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa32").Value
                 End If
              End If
              intRow = intRow + 1
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa32").Value) Then
                    If IsNull(adocheck.Fields("fa19").Value) = False Then
                       Printer.CurrentX = 0 + m_dblLeft
                       Printer.CurrentY = 0 + intRow * 300
                       Printer.Print adocheck.Fields("fa19").Value
                    End If
                 Else
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print "" & adocheck.Fields("fa33").Value
                 End If
              End If
              intRow = intRow + 1
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa32").Value) Then
                    If IsNull(adocheck.Fields("fa20").Value) = False Then
                       Printer.CurrentX = 0 + m_dblLeft
                       Printer.CurrentY = 0 + intRow * 300
                       Printer.Print adocheck.Fields("fa20").Value
                    End If
                 Else
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print "" & adocheck.Fields("fa34").Value
                 End If
              End If
              intRow = intRow + 1
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa32").Value) Then
                    If IsNull(adocheck.Fields("fa21").Value) = False Then
                       Printer.CurrentX = 0 + m_dblLeft
                       Printer.CurrentY = 0 + intRow * 300
                       Printer.Print adocheck.Fields("fa21").Value
                    End If
                 Else
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print "" & adocheck.Fields("fa35").Value
                 End If
              End If
              intRow = intRow + 1
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa32").Value) Then
                    If IsNull(adocheck.Fields("fa22").Value) = False Then
                       Printer.CurrentX = 0 + m_dblLeft
                       Printer.CurrentY = 0 + intRow * 300
                       Printer.Print adocheck.Fields("fa22").Value
                    End If
                 Else
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print "" & adocheck.Fields("fa36").Value
                 End If
              End If
              'Add by Morgan 2011/5/25
              '英文地址6
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa32").Value) Then
                    If IsNull(adocheck.Fields("fa70").Value) = False Then
                       intRow = intRow + 1
                       Printer.CurrentX = 0 + m_dblLeft
                       Printer.CurrentY = 0 + intRow * 300
                       Printer.Print adocheck.Fields("fa70").Value
                    End If
                 End If
              End If
              
           Case "3"
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa06").Value) = False Then
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa06").Value
                 End If
              End If
              intRow = intRow + 1
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa23").Value) = False Then
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa23").Value
                 End If
              End If
              
           '2012/2/22 ADD BY SONIA Y47804無英文
           Case "1"
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa04").Value) = False Then
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa04").Value
                 End If
              End If
              intRow = intRow + 1
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa17").Value) = False Then
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa17").Value
                 End If
              End If
              intRow = intRow + 1
           '2012/2/22 END
        End Select
    End If
    adocheck.Close
    intRow = intRow + 1
    intCounter = intRow

End Sub

'Add By Cheng 2004/03/09
'抓傳票號碼
Private Function GetSummonsNo(strA1P01 As String, stra1p02 As String, strA1P04 As String)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

StrSQLa = "Select A1P22 From ACC1P0 Where A1P01='" & strA1P01 & "' And A1P02='" & stra1p02 & "' And A1P04='" & strA1P04 & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetSummonsNo = "" & rsA.Fields(0).Value
Else
    GetSummonsNo = ""
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'Add By Cheng 2004/03/09
'判斷傳票是否過過帳
'modify by sonia 2017/3/30 加是否顯示訊息的參數strMSG
Private Function ChkPosting(strax202 As String, Optional ByRef strMsg As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

StrSQLa = "Select AX210 From ACC021 Where AX202='" & strax202 & "' And AX210 Is Not Null "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    ChkPosting = True
    'modify by sonia 2017/3/30
    'MsgBox "傳票 " & strax202 & " 已經過帳!!!", vbExclamation + vbOKOnly
    If strMsg = "" Then MsgBox "傳票 " & strax202 & " 已經過帳!!!", vbExclamation + vbOKOnly
    'end 2017/3/30
Else
    ChkPosting = False
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function
'Add by Morgan 2006/7/17
Private Sub SetDate(p_a1g01 As String)
   Text12 = ""
   strExc(0) = "select a1p18 from acc1g0 x,acc1p0 where a1p04(+)=a1g01 and  a1g01 = '" & p_a1g01 & "' and a1p18>0 and rownum<2"
   intI = 1
   'edit by nickc 2007/02/07 不用 dll 了
   'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Text12 = Format(RsTemp.Fields(0), "###/##/##")
   End If
End Sub
'*************************************************
'  產生分錄資料
'
'*************************************************
Public Sub ProcessData()
Dim strDept As String
Dim strAccNo As String
Dim strCaseNo As String
Dim strProperty As String
Dim strRemark As String
Dim strSalesNo As String
Dim strCustomerNo As String
Dim douService As Double
Dim douFee As Double
Dim douBalance As Double
Dim douDebitAmount As Double
Dim strName As String
Dim strName1 As String
'2005/5/16 ADD BY SONIA
Dim strNation As String
'2005/5/27 ADD BY SONIA
Dim StrStaff As String
Dim m_strDomAmt  As String    '國內收款金額  2010/6/30 add by sonia
Dim strAccNo1 As String 'Added by Lydia 2016/07/18

   'Add by Morgan 2011/9/6
   m_bolAlert = False
   m_strAlertMsg = ""
   'end 2011/9/6
   
   If Adodc1.Recordset.RecordCount = 0 Or Adodc2.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   Adodc2.Recordset.MoveFirst
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from acc1x0 where a1x01 = '" & Adodc2.Recordset.Fields("a1505").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields("a1x02").Value) = False Then
         douBalance = Val(adoquery.Fields("a1x02").Value)
      Else
         douBalance = 0
      End If
   Else
      douBalance = 0
   End If
   adoquery.Close
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "' and a1p22 is not null", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      adoquery.Close
      Exit Sub
   End If
   adoquery.Close
   
   'Modified by Morgan 2022/1/21 改依輸入選項決定(有特殊情形,分錄自行調整)--婧瑄
   'If Val(Text7) > Val(Text4) Then
   If Text11 = "1" Then
   'end 2022/1/21
      'adoTaie.Execute "delete from acc1h0 where a1h01 = '" & Text9 & "'"
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select * from acc1h0 where a1h01 = '" & Text9 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount = 0 Then
         adoTaie.Execute "insert into acc1h0 (a1h01, a1h02, a1h03, a1h04, a1h05, a1h06, a1h07, a1h08) values ('" & Text9 & "', " & Val(strSrvDate(2)) & ", 'USD', " & Val(Text2) & ", " & CNULL(Text10) & ", " & Val(strSrvDate(2)) & ", " & ServerTime & ", '" & strUserNum & "')"
         adoTaie.Execute "delete from acc1i0 where a1i01 = '" & Text9 & "'"   'add by sonia 2021/4/22
      End If
      adoquery.Close
   Else
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select * from acc1i0 where a1i01 = '" & Text9 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount = 0 Then
         adoTaie.Execute "insert into acc1i0 (a1i01, a1i03, a1i05, a1i06, a1i07, a1i08, a1i09, a1i10, a1i11) values ('" & Text9 & "', " & Val(strSrvDate(2)) & ", '" & Adodc2.Recordset.Fields("a1505").Value & "', " & douBalance & ", " & Val(Format(Val(Text8) - Val(Text3), FAmount)) & ", " & CNULL(Text10) & ", " & Val(strSrvDate(2)) & ", " & ServerTime & ", '" & strUserNum & "')"
         adoTaie.Execute "delete from acc1h0 where a1h01 = '" & Text9 & "'"   'add by sonia 2021/4/22
      End If
      adoquery.Close
   End If
   adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'"
   Adodc2.Recordset.MoveFirst
   Do While Adodc2.Recordset.EOF = False
      adoaccsum.CursorLocation = adUseClient
      '2005/5/16 MODIFY BY SONIA 加 casepropertymap
      'adoaccsum.Open "select * from caseprogress, acc151 where cp09 = axf02 and cp61 = axf01 and cp61 = '" & Adodc2.Recordset.Fields("a1501").Value & "' and cp09 = '" & Adodc2.Recordset.Fields("axf02").Value & "' union " & _
      '               "select * from caseprogress, acc151 where cp09 = axf02 and cp62 = axf01 and cp62 = '" & Adodc2.Recordset.Fields("a1501").Value & "' and cp09 = '" & Adodc2.Recordset.Fields("axf02").Value & "' union " & _
      '               "select * from caseprogress, acc151 where cp09 = axf02 and cp63 = axf01 and cp63 = '" & Adodc2.Recordset.Fields("a1501").Value & "' and cp09 = '" & Adodc2.Recordset.Fields("axf02").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      '2007/3/2 modify by sonia 加入cp87,cp88
      'adoaccsum.Open "select * from caseprogress, acc151, casepropertymap where cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp09 = axf02 and cp61 = axf01 and cp61 = '" & Adodc2.Recordset.Fields("a1501").Value & "' and cp09 = '" & Adodc2.Recordset.Fields("axf02").Value & "' union " & _
      '               "select * from caseprogress, acc151, casepropertymap where cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp09 = axf02 and cp62 = axf01 and cp62 = '" & Adodc2.Recordset.Fields("a1501").Value & "' and cp09 = '" & Adodc2.Recordset.Fields("axf02").Value & "' union " & _
      '               "select * from caseprogress, acc151, casepropertymap where cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp09 = axf02 and cp63 = axf01 and cp63 = '" & Adodc2.Recordset.Fields("a1501").Value & "' and cp09 = '" & Adodc2.Recordset.Fields("axf02").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      adoaccsum.Open "select * from caseprogress, acc151, casepropertymap where cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp09 = axf02 and cp61 = axf01 and cp61 = '" & Adodc2.Recordset.Fields("a1501").Value & "' and cp09 = '" & Adodc2.Recordset.Fields("axf02").Value & "' union " & _
                     "select * from caseprogress, acc151, casepropertymap where cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp09 = axf02 and cp62 = axf01 and cp62 = '" & Adodc2.Recordset.Fields("a1501").Value & "' and cp09 = '" & Adodc2.Recordset.Fields("axf02").Value & "' union " & _
                     "select * from caseprogress, acc151, casepropertymap where cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp09 = axf02 and cp63 = axf01 and cp63 = '" & Adodc2.Recordset.Fields("a1501").Value & "' and cp09 = '" & Adodc2.Recordset.Fields("axf02").Value & "' union " & _
                     "select * from caseprogress, acc151, casepropertymap where cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp09 = axf02 and cp87 = axf01 and cp87 = '" & Adodc2.Recordset.Fields("a1501").Value & "' and cp09 = '" & Adodc2.Recordset.Fields("axf02").Value & "' union " & _
                     "select * from caseprogress, acc151, casepropertymap where cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp09 = axf02 and cp88 = axf01 and cp88 = '" & Adodc2.Recordset.Fields("a1501").Value & "' and cp09 = '" & Adodc2.Recordset.Fields("axf02").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      '2007/3/2 end
      '2005/5/16 END
      Do While adoaccsum.EOF = False
         strCaseNo = adoaccsum.Fields("cp01").Value & adoaccsum.Fields("cp02").Value & adoaccsum.Fields("cp03").Value & adoaccsum.Fields("cp04").Value
         '2010/6/30 modify by sonia
         'strRemark = strCaseNo
         'If IsNull(Adodc2.Recordset.Fields("a1505").Value) = False Then
         '   strRemark = strRemark & "/" & Adodc2.Recordset.Fields("a1505").Value
         'End If
         'If IsNull(adoaccsum.Fields("axf04").Value) = False Then
         '   strRemark = strRemark & " " & adoaccsum.Fields("axf04").Value
         'End If
         'Mid(Right("00" & adoquery.Fields("Caseno").Value, 12), 4, 6) & " " & Left("" & GetA0K04("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value), 4) & " " & GetA1l02("" & adoquery.Fields("Caseno").Value, "" & adoquery.Fields("axf02").Value) & "/" & m_strDomAmt & " " & strCurrency & " " & Format(Val(adoquery.Fields("Famount").Value), FDollar)
         '2011/8/19 modify by sonia GetA1l02加傳收款金額m_strDomAmt
         '2011/11/23 modify by sonia 把收款日/收款金額改到摘要的最前面,以便外帳核對資料
         'strRemark = Mid(Right("00" & strCaseNo, 12), 4, 6) & " " & Left("" & GetA0K04("" & strCaseNo, "" & adoaccsum.Fields("axf02").Value), 4) & " " & GetA1l02("" & strCaseNo, "" & adoaccsum.Fields("axf02").Value, m_strDomAmt) & "/" & m_strDomAmt
         strRemark = GetA1l02("" & strCaseNo, "" & adoaccsum.Fields("axf02").Value, m_strDomAmt) & "/" & m_strDomAmt & " " & Mid(Right("00" & strCaseNo, 12), 4, 6) & " " & Left("" & GetA0K04("" & strCaseNo, "" & adoaccsum.Fields("axf02").Value), 4)
         '2011/11/23 end
         
         If IsNull(Adodc2.Recordset.Fields("a1505").Value) = False Then
            strRemark = strRemark & " " & Adodc2.Recordset.Fields("a1505").Value
         End If
         If IsNull(adoaccsum.Fields("axf04").Value) = False Then
            strRemark = strRemark & " " & adoaccsum.Fields("axf04").Value
         End If
         '2010/6/30 end
         If IsNull(adoaccsum.Fields("cp13").Value) Then
            strSalesNo = ""
         Else
            strSalesNo = adoaccsum.Fields("cp13").Value
         End If
         adoquery.CursorLocation = adUseClient
         '2005/5/16 MODIFY BY SONIA 加申請國家
         'adoquery.Open "select pa26 as CustNo from patent where pa01 = '" & adoaccsum.Fields("cp01").Value & "' and pa02 = '" & adoaccsum.Fields("cp02").Value & "' and pa03 = '" & adoaccsum.Fields("cp03").Value & "' and pa04 = '" & adoaccsum.Fields("cp04").Value & "' union " & _
         '              "select tm23 as CustNo from trademark where tm01 = '" & adoaccsum.Fields("cp01").Value & "' and tm02 = '" & adoaccsum.Fields("cp02").Value & "' and tm03 = '" & adoaccsum.Fields("cp03").Value & "' and tm04 = '" & adoaccsum.Fields("cp04").Value & "' union " & _
         '              "select lc11 as CustNo from lawcase where lc01 = '" & adoaccsum.Fields("cp01").Value & "' and lc02 = '" & adoaccsum.Fields("cp02").Value & "' and lc03 = '" & adoaccsum.Fields("cp03").Value & "' and lc04 = '" & adoaccsum.Fields("cp04").Value & "' union " & _
         '              "select hc05 as CustNo from hirecase where hc01 = '" & adoaccsum.Fields("cp01").Value & "' and hc02 = '" & adoaccsum.Fields("cp02").Value & "' and hc03 = '" & adoaccsum.Fields("cp03").Value & "' and hc04 = '" & adoaccsum.Fields("cp04").Value & "' union " & _
         '              "select sp08 as CustNo from servicepractice where sp01 = '" & adoaccsum.Fields("cp01").Value & "' and sp02 = '" & adoaccsum.Fields("cp02").Value & "' and sp03 = '" & adoaccsum.Fields("cp03").Value & "' and sp04 = '" & adoaccsum.Fields("cp04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         adoquery.Open "select pa26 as CustNo, PA09 AS strNation from patent where pa01 = '" & adoaccsum.Fields("cp01").Value & "' and pa02 = '" & adoaccsum.Fields("cp02").Value & "' and pa03 = '" & adoaccsum.Fields("cp03").Value & "' and pa04 = '" & adoaccsum.Fields("cp04").Value & "' union " & _
                       "select tm23 as CustNo, TM10 AS strNation from trademark where tm01 = '" & adoaccsum.Fields("cp01").Value & "' and tm02 = '" & adoaccsum.Fields("cp02").Value & "' and tm03 = '" & adoaccsum.Fields("cp03").Value & "' and tm04 = '" & adoaccsum.Fields("cp04").Value & "' union " & _
                       "select lc11 as CustNo, LC15 AS strNation from lawcase where lc01 = '" & adoaccsum.Fields("cp01").Value & "' and lc02 = '" & adoaccsum.Fields("cp02").Value & "' and lc03 = '" & adoaccsum.Fields("cp03").Value & "' and lc04 = '" & adoaccsum.Fields("cp04").Value & "' union " & _
                       "select hc05 as CustNo, '000' AS strNation from hirecase where hc01 = '" & adoaccsum.Fields("cp01").Value & "' and hc02 = '" & adoaccsum.Fields("cp02").Value & "' and hc03 = '" & adoaccsum.Fields("cp03").Value & "' and hc04 = '" & adoaccsum.Fields("cp04").Value & "' union " & _
                       "select sp08 as CustNo, SP09 AS strNation from servicepractice where sp01 = '" & adoaccsum.Fields("cp01").Value & "' and sp02 = '" & adoaccsum.Fields("cp02").Value & "' and sp03 = '" & adoaccsum.Fields("cp03").Value & "' and sp04 = '" & adoaccsum.Fields("cp04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         '2005/5/16 END
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields("CustNo").Value) Then
               strCustomerNo = ""
            Else
               strCustomerNo = adoquery.Fields("CustNo").Value
            End If
            '2005/5/16 ADD BY SONIA
            If IsNull(adoquery.Fields("strNation").Value) Then
               strNation = ""
            Else
               strNation = adoquery.Fields("strNation").Value
            End If
            '2005/5/16 END
         Else
            strCustomerNo = ""
            strNation = ""
         End If
         adoquery.Close
         '2005/5/27 CANCEL BY SONIA 因為規費之部門都掛 TOT
         'Select Case adoaccsum.Fields("cp01").Value
         '   Case "T", "TF"             '2005/5/16 加入TF
         '      strDept = "T"
         '   Case "P", "PS"             '2005/5/16 加入PS
         '      strDept = "P"
         '   Case "FCT"
         '      strDept = "FCT"
         '   Case "FCP", "FG"           '2005/5/16 加入FG
         '      strDept = "FCP"
         '   Case "CFT", "CFC"          '2005/5/16 加入CFC
         '      strDept = "CFT"
         '   Case "CFP", "CPS"          '2005/5/16 加入CPS
         '      strDept = "CFP"
         '   '2005/5/16 ADD BY SONIA
         '   Case "L"
         '      strDept = "L"
         '   Case "FCL", "CFL"
         '      strDept = "FCL"
         '   Case "S"
         '      If strNation = "000" Then
         '         strDept = "FCT"
         '      Else
         '         strDept = "CFT"
         '      End If
         '   '2005/5/16 END
         '   Case Else
         '      '2005/5/16 MODIFY BYS SONIA
         '      'strDept = "TOT"
         '      strDept = "T"
         '      '2005/5/16 END
         'End Select
         '2005/3/23 modify by sonia
         'If Len(adoaccsum.Fields("cp01").Value) = 3 Then
         '   If adoaccsum.Fields("cp01").Value = "CFT" Or adoaccsum.Fields("cp01").Value = "CFC" Then
         '      adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p18, a1p19, a1p20, a1p21, a1p28, a1p29, a1p14, a1p17, a1p16, a1p15) values ('1', 'K', '" & GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'", 3) & "', '" & Text9 & "', '220105', '" & MsgText(55) & "', " & Val(Format(adoaccsum.Fields("axf04").Value * Val(Text6), DAmount)) & ", 0, " & Val(ACDate(ServerDate)) & ", '" & Adodc2.Recordset.Fields("a1505").Value & "', " & Val(Text6) & ", " & Val(adoaccsum.Fields("axf04").Value) & ", " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", '" & strRemark & "', '" & strCaseNo & "', '" & strSalesNo & "', '" & strCustomerNo & "')"
         '   Else
         '      adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p18, a1p19, a1p20, a1p21, a1p28, a1p29, a1p14, a1p17, a1p16, a1p15) values ('1', 'K', '" & GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'", 3) & "', '" & Text9 & "', '220106', '" & MsgText(55) & "', " & Val(Format(adoaccsum.Fields("axf04").Value * Val(Text6), DAmount)) & ", 0, " & Val(ACDate(ServerDate)) & ", '" & Adodc2.Recordset.Fields("a1505").Value & "', " & Val(Text6) & ", " & Val(adoaccsum.Fields("axf04").Value) & ", " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", '" & strRemark & "', '" & strCaseNo & "', '" & strSalesNo & "', '" & strCustomerNo & "')"
         '   End If
         'Else
         '   If adoaccsum.Fields("cp01").Value = "T" Then
         '      adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p18, a1p19, a1p20, a1p21, a1p28, a1p29, a1p14, a1p17, a1p16, a1p15) values ('1', 'K', '" & GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'", 3) & "', '" & Text9 & "', '220111', '" & MsgText(55) & "', " & Val(Format(adoaccsum.Fields("axf04").Value * Val(Text6), DAmount)) & ", 0, " & Val(ACDate(ServerDate)) & ", '" & Adodc2.Recordset.Fields("a1505").Value & "', " & Val(Text6) & ", " & Val(adoaccsum.Fields("axf04").Value) & ", " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", '" & strRemark & "', '" & strCaseNo & "', '" & strSalesNo & "', '" & strCustomerNo & "')"
         '   Else
         '      adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p18, a1p19, a1p20, a1p21, a1p28, a1p29, a1p14, a1p17, a1p16, a1p15) values ('1', 'K', '" & GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'", 3) & "', '" & Text9 & "', '220112', '" & MsgText(55) & "', " & Val(Format(adoaccsum.Fields("axf04").Value * Val(Text6), DAmount)) & ", 0, " & Val(ACDate(ServerDate)) & ", '" & Adodc2.Recordset.Fields("a1505").Value & "', " & Val(Text6) & ", " & Val(adoaccsum.Fields("axf04").Value) & ", " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", '" & strRemark & "', '" & strCaseNo & "', '" & strSalesNo & "', '" & strCustomerNo & "')"
         '   End If
         'End If
         '2005/5/16 MODIFY BY SONIA
         'If Len(adoaccsum.Fields("cp01").Value) = 3 Then
         '   If adoaccsum.Fields("cp01").Value = "CFT" Or adoaccsum.Fields("cp01").Value = "CFC" Then
         '      strAccNo = "220105"
         '   Else
         '      strAccNo = "220106"
         '   End If
         'Else
         '   If adoaccsum.Fields("cp01").Value = "T" Then
         '      strAccNo = "220111"
         '   Else
         '      strAccNo = "220112"
         '   End If
         'End If
         strAccNo = adoaccsum.Fields("cpm12").Value
         Select Case adoaccsum.Fields("cp01").Value
            '2007/8/14 MODIFY BY SONIA Z09600013
            'Case "T", "TS"
            Case "T", "TS", "TB", "TC", "TD", "TM", "TR", "TT"
               If strNation <> "000" Then
                  strAccNo = "220111"
               End If
            Case "P", "PS"
               If strNation <> "000" Then
                  strAccNo = "220112"
               End If
            Case "S"
               Select Case strNation
                  Case "000"
                     strAccNo = "220103"
                  Case "020"
                     strAccNo = "220112"
                  Case Else
                     strAccNo = "220105"
               End Select
         End Select
         '2005/5/16 END
         
         'modify by sonia 2021/3/12 加傳日期
         strSalesNo = SalesNoToAccSales(strSalesNo, strAccNo, strCaseNo, Val(strSrvDate(2)))
         If strSalesNo <> "" Then
         Else
            strSalesNo = "M0100"
         End If
         '2012/10/16 MODIFY BY SONIA 加存A1P23,轉帳務時才能判斷資料來源Z10100007(P097900000)為貸方退費,Z10100029(P098968000)為貸方收款
         'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p18, a1p19, a1p20, a1p21, a1p28, a1p29, a1p14, a1p17, a1p16, a1p15) values ('1', 'K', '" & GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'", 3) & "', '" & Text9 & "', '" & strAccNo & "', '" & MsgText(55) & "', " & Val(Format(adoaccsum.Fields("axf04").Value * Val(Text6), DAmount)) & ", 0, " & Val(strSrvDate(2)) & ", '" & Adodc2.Recordset.Fields("a1505").Value & "', " & Val(Text6) & ", " & Val(adoaccsum.Fields("axf04").Value) & ", " & Val(strSrvDate(2)) & ", " & ServerTime & ", '" & strRemark & "', '" & strCaseNo & "', '" & strSalesNo & "', '" & strCustomerNo & "')"
         'modify sonia 2017/4/17 北京寰華介紹案源(a1803='Y53374000' & a0k34='F5639'之結匯,改借方規費為收入(扣業務點數)
         'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p18, a1p19, a1p20, a1p21, a1p28, a1p29, a1p14, a1p17, a1p16, a1p15, a1p23) values ('1', 'K', '" & GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'", 3) & "', '" & Text9 & "', '" & strAccNo & "', '" & MsgText(55) & "', " & Val(Format(adoaccsum.Fields("axf04").Value * Val(Text6), DAmount)) & ", 0, " & Val(strSrvDate(2)) & ", '" & Adodc2.Recordset.Fields("a1505").Value & "', " & Val(Text6) & ", " & Val(adoaccsum.Fields("axf04").Value) & ", " & Val(strSrvDate(2)) & ", " & ServerTime & ", '" & strRemark & "', '" & strCaseNo & "', '" & strSalesNo & "', '" & strCustomerNo & "','" & Adodc2.Recordset.Fields("a1501").Value & "')"
         adoacc150.CursorLocation = adUseClient
         adoacc150.Open "select axf03,a0k20,sn01,cpm24,cp10,cp09,a0k03,substr(nvl(nvl(cu04,cu05),cu06),6) cu04 from acc150,acc151,acc0j0,acc0k0,customer,salesno,caseprogress,casepropertymap " & _
                       "where a1501='" & Adodc2.Recordset.Fields("a1501").Value & "' and a1503='Y53374000' and a1501=axf01(+) and axf02=a0j01(+) and a0j13=a0k01(+) and a0k34='F5639' " & _
                       "and substr(a0k03,1,8)=cu01(+) and substr(a0k03,9,1)=cu02(+) and a0k20=sn02(+) and axf02=cp09(+) and cp01=cpm01(+) and cp10=cpm02(+)", adoTaie, adOpenStatic, adLockReadOnly
         If adoacc150.RecordCount <> 0 Then
            strRemark = "" & adoacc150.Fields("sn01") & "/" & adoacc150.Fields("cu04") & "/" & "北京寰華介紹案源" & "/" & adoacc150.Fields("axf03")
            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p18, a1p19, a1p20, a1p21, a1p28, a1p29, a1p14, a1p17, a1p16, a1p15, a1p23) values ('1', 'K', '" & GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'", 3) & "', '" & Text9 & "', '" & "" & adoacc150.Fields("cpm24") & "', '" & MsgText(55) & "', " & Val(Format(adoaccsum.Fields("axf04").Value * Val(Text6), DAmount)) & ", 0, " & Val(strSrvDate(2)) & ", '" & Adodc2.Recordset.Fields("a1505").Value & "', " & Val(Text6) & ", " & Val(adoaccsum.Fields("axf04").Value) & ", " & Val(strSrvDate(2)) & ", " & ServerTime & ", '" & strRemark & "', '" & strCaseNo & "', '" & strSalesNo & "', '" & strCustomerNo & "','" & Adodc2.Recordset.Fields("a1501").Value & "')"
         Else
            adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p18, a1p19, a1p20, a1p21, a1p28, a1p29, a1p14, a1p17, a1p16, a1p15, a1p23) values ('1', 'K', '" & GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'", 3) & "', '" & Text9 & "', '" & strAccNo & "', '" & MsgText(55) & "', " & Val(Format(adoaccsum.Fields("axf04").Value * Val(Text6), DAmount)) & ", 0, " & Val(strSrvDate(2)) & ", '" & Adodc2.Recordset.Fields("a1505").Value & "', " & Val(Text6) & ", " & Val(adoaccsum.Fields("axf04").Value) & ", " & Val(strSrvDate(2)) & ", " & ServerTime & ", '" & strRemark & "', '" & strCaseNo & "', '" & strSalesNo & "', '" & strCustomerNo & "','" & Adodc2.Recordset.Fields("a1501").Value & "')"
         End If
         adoacc150.Close
         'end 2017/4/17
         '2005/3/23 end
         adoaccsum.MoveNext
      Loop
      adoaccsum.Close
      adoTaie.Execute "update acc150 set a1512 = '" & Text9 & "', a1513 = " & Val(Text6) & ", a1520 = a1506 where a1501 = '" & Adodc2.Recordset.Fields("a1501").Value & "'"
      Adodc2.Recordset.MoveNext
   Loop
   Adodc2.Recordset.MoveFirst
   Adodc1.Recordset.MoveFirst
   Do While Adodc1.Recordset.EOF = False
      If strName <> (Adodc1.Recordset.Fields(0).Value & Adodc1.Recordset.Fields(1).Value) Then
         
         'Add by Morgan 2011/9/6 若有分配資料則跑新程式
         strExc(0) = "select * from acc1n0 where a1n01='" & Adodc1.Recordset.Fields("a1k01") & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Acc1p0Save
            GoTo NextRec
         End If
         'end 2011/9/6
      
         adoaccsum.CursorLocation = adUseClient
         adoaccsum.Open "select * from caseprogress, casepropertymap where cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp60 = '" & Adodc1.Recordset.Fields("a1k01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         Do While adoaccsum.EOF = False
            If strName1 <> adoaccsum.Fields("cp60").Value Then
               If IsNull(adoaccsum.Fields("cpm03").Value) Then
                  If IsNull(adoaccsum.Fields("cpm04").Value) Then
                     strProperty = ""
                  Else
                     strProperty = adoaccsum.Fields("cpm04").Value
                  End If
               Else
                  strProperty = adoaccsum.Fields("cpm03").Value
               End If
               strCaseNo = adoaccsum.Fields("cp01").Value & adoaccsum.Fields("cp02").Value & adoaccsum.Fields("cp03").Value & adoaccsum.Fields("cp04").Value
               If IsNull(adoaccsum.Fields("cp13").Value) Then
                  strSalesNo = ""
               Else
                  strSalesNo = adoaccsum.Fields("cp13").Value
               End If
               '2005/5/27 ADD BY SONIA
               If IsNull(adoaccsum.Fields("cp14").Value) Then
                  StrStaff = ""
               Else
                  StrStaff = adoaccsum.Fields("cp14").Value
               End If
               '2005/5/27 END
               adoquery.CursorLocation = adUseClient
               '2005/5/16 MODIFY BY SONIA 加申請國家
               adoquery.Open "select pa26 as CustNo, PA09 AS strNation from patent where pa01 = '" & adoaccsum.Fields("cp01").Value & "' and pa02 = '" & adoaccsum.Fields("cp02").Value & "' and pa03 = '" & adoaccsum.Fields("cp03").Value & "' and pa04 = '" & adoaccsum.Fields("cp04").Value & "' union " & _
                             "select tm23 as CustNo, TM10 AS strNation from trademark where tm01 = '" & adoaccsum.Fields("cp01").Value & "' and tm02 = '" & adoaccsum.Fields("cp02").Value & "' and tm03 = '" & adoaccsum.Fields("cp03").Value & "' and tm04 = '" & adoaccsum.Fields("cp04").Value & "' union " & _
                             "select lc11 as CustNo, LC15 AS strNation from lawcase where lc01 = '" & adoaccsum.Fields("cp01").Value & "' and lc02 = '" & adoaccsum.Fields("cp02").Value & "' and lc03 = '" & adoaccsum.Fields("cp03").Value & "' and lc04 = '" & adoaccsum.Fields("cp04").Value & "' union " & _
                             "select hc05 as CustNo, '000' AS strNation from hirecase where hc01 = '" & adoaccsum.Fields("cp01").Value & "' and hc02 = '" & adoaccsum.Fields("cp02").Value & "' and hc03 = '" & adoaccsum.Fields("cp03").Value & "' and hc04 = '" & adoaccsum.Fields("cp04").Value & "' union " & _
                             "select sp08 as CustNo, SP09 AS strNation from servicepractice where sp01 = '" & adoaccsum.Fields("cp01").Value & "' and sp02 = '" & adoaccsum.Fields("cp02").Value & "' and sp03 = '" & adoaccsum.Fields("cp03").Value & "' and sp04 = '" & adoaccsum.Fields("cp04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
               '2005/5/16 END
               If adoquery.RecordCount <> 0 Then
                  If IsNull(adoquery.Fields("CustNo").Value) Then
                     strCustomerNo = ""
                  Else
                     strCustomerNo = adoquery.Fields("CustNo").Value
                  End If
                  '2005/5/16 ADD BY SONIA
                  If IsNull(adoquery.Fields("strNation").Value) Then
                     strNation = ""
                  Else
                     strNation = adoquery.Fields("strNation").Value
                  End If
                  '2005/5/16 END
               Else
                  strCustomerNo = ""
                  strNation = ""
               End If
               adoquery.Close

               Select Case adoaccsum.Fields("cp01").Value
                  '2005/5/27 MODIFY BY SONIA
                  'Case "T"
                  '   strDept = "T"
                  'Case "P"
                  '   strDept = "P"
                  'Case "FCT"
                  '   strDept = "FCT"
                  'Case "FCP"
                  '   strDept = "FCP"
                  'Case "CFT"
                  '   strDept = "CFT"
                  'Case "CFP"
                  '   strDept = "CFP"
                  'Case Else
                  '   strDept = "TOT"
                  Case "T", "TF"
                     strDept = "T"
                  Case "P", "PS"
                     strDept = "P"
                  Case "FCT"
                     strDept = "FCT"
                     If adocheck.State = adStateOpen Then
                        adocheck.Close
                     End If
                     adocheck.CursorLocation = adUseClient
                     adocheck.Open "select st03 from staff where st01 = '" & StrStaff & "'", adoTaie, adOpenStatic, adLockReadOnly
                     If adocheck.RecordCount <> 0 Then
                        If IsNull(adocheck.Fields("st03").Value) = False Then
                           If Mid(adocheck.Fields("st03").Value, 1, 2) = "P2" Then
                              strDept = "T"
                           End If
                        End If
                     End If
                     adocheck.Close
                  Case "FCP", "FG"
                     strDept = "FCP"
                  Case "CFT", "CFC"
                     strDept = "CFT"
                  Case "CFP", "CPS"
                     strDept = "CFP"
                  Case "L"
                     strDept = "L"
                  Case "FCL", "CFL"
                     strDept = "FCL"
                  Case "S"
                     If strNation = "000" Then
                        strDept = "FCT"
                     Else
                        strDept = "CFT"
                     End If
                  Case Else
                     strDept = "T"
                  '2005/5/27 END
               End Select
              
               If IsNull(Adodc1.Recordset.Fields("a1k09").Value) = False Then
                  douFee = Val(Adodc1.Recordset.Fields("a1k09").Value)
               Else
                  douFee = Val(adoaccsum.Fields("cp17").Value)
               End If
               
               'Added by Lydia 2016/07/18 抵帳作業的規費只到整數
               strAccNo1 = GetFeeAccNo(adoaccsum.Fields("cp01"), strNation)
               If strAccNo1 = "" Then strAccNo1 = "" & adoaccsum.Fields("cpm12") 'Added by Lydia 2019/01/11 抵帳X10800002的國籍為台灣,又沒有分配點數
               If Left(strAccNo1, 4) = "2201" Then douFee = Val(Format(douFee, DAmount))
               
               If IsNull(Adodc1.Recordset.Fields("a1k11").Value) = False Then
                  If Adodc1.Recordset.Fields("a1k11").Value = 0 Then
                     douService = Val(Adodc1.Recordset.Fields("a1k30").Value) - douFee
                  Else
                     douService = Val(Adodc1.Recordset.Fields("a1k11").Value) - douFee
                  End If
      '         Else
      '            douService = Val(Adodc1.Recordset.Fields("a1k11").Value) - Val(Adodc1.Recordset.Fields("a1k09").Value)
               End If
               If IsNull(adoaccsum.Fields("cpm11").Value) Then
                  strAccNo = MsgText(601)
               Else
                  strAccNo = adoaccsum.Fields("cpm11").Value
               End If
               '2005/3/23 modify by sonia
               'If AccNoToSalesNo(strAccNo) <> "" Then
               '   strSalesNo = AccNoToSalesNo(strAccNo)
               'End If
               'modify by sonia 2021/3/12 加傳日期
               strSalesNo = SalesNoToAccSales(strSalesNo, strAccNo, strCaseNo, Val(strSrvDate(2)))
               If strSalesNo <> "" Then
               Else
                  strSalesNo = "M0100"
               End If
               '2005/3/23 end
               strRemark = strCaseNo & "/" & strProperty
               If douService <> 0 Then
                  '2012/10/16 MODIFY BY SONIA 加存A1P23,轉帳務時才能判斷資料來源Z10100007(P097900000)為貸方退費,Z10100029(P098968000)為貸方收款
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p18, a1p19, a1p20, a1p21, a1p28, a1p29, a1p14, a1p17, a1p16, a1p15) values ('1', 'K', '" & GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'", 3) & "', '" & Text9 & "', '" & strAccNo & "', '" & strDept & "', 0, " & douService & ", " & Val(strSrvDate(2)) & ", 'USD', " & Val(Text2) & ", " & Val(Format(douService / Val(Text2), FAmount)) & ", " & Val(strSrvDate(2)) & ", " & ServerTime & ", '" & strRemark & "', '" & strCaseNo & "', '" & strSalesNo & "', '" & strCustomerNo & "')"
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p18, a1p19, a1p20, a1p21, a1p28, a1p29, a1p14, a1p17, a1p16, a1p15, a1p23) values ('1', 'K', '" & GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'", 3) & "', '" & Text9 & "', '" & strAccNo & "', '" & strDept & "', 0, " & douService & ", " & Val(strSrvDate(2)) & ", 'USD', " & Val(Text2) & ", " & Val(Format(douService / Val(Text2), FAmount)) & ", " & Val(strSrvDate(2)) & ", " & ServerTime & ", '" & strRemark & "', '" & strCaseNo & "', '" & strSalesNo & "', '" & strCustomerNo & "','" & Adodc1.Recordset.Fields("a1k01").Value & "')"
               End If
               If IsNull(adoaccsum.Fields("cpm12").Value) Then
                  strAccNo = MsgText(601)
               Else
                  strAccNo = adoaccsum.Fields("cpm12").Value
               End If
               '2005/5/27 ADD BY SONIA
               'Added by Lydia 2016/07/18 改成模組
               'Select Case adoaccsum.Fields("cp01").Value
               '   Case "T", "TS"
               '      If strNation <> "000" Then
               '         strAccNo = "220111"
               '      End If
               '   Case "P", "PS"
               '      If strNation <> "000" Then
               '         strAccNo = "220112"
               '      End If
               '   Case "S"
               '      Select Case strNation
               '         Case "000"
               '            strAccNo = "220103"
               '         Case "020"
               '            strAccNo = "220112"
               '         Case Else
               '            strAccNo = "220105"
               '      End Select
               'End Select
               ''2005/5/27 END
               strAccNo = strAccNo1
               
               'strRemark = strCaseNo & "/USD " & Val(Format(douFee / Val(Text2), FAmount))
               '2005/3/23 add by sonia
               'modify by sonia 2021/3/12 加傳日期
               strSalesNo = SalesNoToAccSales(strSalesNo, strAccNo, strCaseNo, Val(strSrvDate(2)))
               If strSalesNo <> "" Then
               Else
                  strSalesNo = "M0100"
               End If
               '2005/3/23 end
               If douFee <> 0 Then
                  '2012/10/16 MODIFY BY SONIA 加存A1P23,轉帳務時才能判斷資料來源Z10100007(P097900000)為貸方退費,Z10100029(P098968000)為貸方收款
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p18, a1p19, a1p20, a1p21, a1p28, a1p29, a1p14, a1p17, a1p16, a1p15) values ('1', 'K', '" & GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'", 3) & "', '" & Text9 & "', '" & strAccNo & "', '" & MsgText(55) & "', 0, " & douFee & ", " & Val(strSrvDate(2)) & ", 'USD', " & Val(Text2) & ", " & Val(Format(douFee / Val(Text2), FAmount)) & ", " & Val(strSrvDate(2)) & ", " & ServerTime & ", '" & strRemark & "', '" & strCaseNo & "', '" & strSalesNo & "', '" & strCustomerNo & "')"
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p18, a1p19, a1p20, a1p21, a1p28, a1p29, a1p14, a1p17, a1p16, a1p15, a1p23) values ('1', 'K', '" & GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'", 3) & "', '" & Text9 & "', '" & strAccNo & "', '" & MsgText(55) & "', 0, " & douFee & ", " & Val(strSrvDate(2)) & ", 'USD', " & Val(Text2) & ", " & Val(Format(douFee / Val(Text2), FAmount)) & ", " & Val(strSrvDate(2)) & ", " & ServerTime & ", '" & strRemark & "', '" & strCaseNo & "', '" & strSalesNo & "', '" & strCustomerNo & "','" & Adodc1.Recordset.Fields("a1k01").Value & "')"
               End If
               strName1 = adoaccsum.Fields("cp60").Value
            End If
            adoaccsum.MoveNext
         Loop
         adoaccsum.Close
         
NextRec: 'Add by Morgan 2011/9/6

         'adoTaie.Execute "update acc1k0 set a1k29 = '" & MsgText(602) & "', a1k30 = a1k11 where a1k01 = '" & Adodc1.Recordset.Fields("a1k01").Value & "'"
         adoTaie.Execute "update acc1k0 set a1k17 = '" & Text9 & "' where a1k01 = '" & Adodc1.Recordset.Fields("a1k01").Value & "'"
         strName = (Adodc1.Recordset.Fields(0).Value & Adodc1.Recordset.Fields(1).Value)
      End If
      Adodc1.Recordset.MoveNext
   Loop
   Adodc1.Recordset.MoveFirst
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(a1p07) as Debit, sum(a1p08) as Credit from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If adoaccsum.Fields("Debit").Value > adoaccsum.Fields("Credit").Value Then
         douBalance = Val(Text2)
      Else
         douBalance = Val(Text6)
         douDebitAmount = Val(Format((Val(Text8) - Val(Text3)) * douBalance, DAmount))
         If Text10 <> MsgText(602) Then
            If douDebitAmount > 0 Then
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p18, a1p28, a1p29, a1p19, a1p20, a1p21, a1p14) values ('1', 'K', '" & GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'", 3) & "', '" & Text9 & "', '110205', '" & MsgText(55) & "', 0, " & douDebitAmount & ", " & Val(strSrvDate(2)) & ", " & Val(strSrvDate(2)) & ", " & ServerTime & ", 'USD', " & douBalance & ", " & (Val(Text8) - Val(Text3)) & ", '" & "結匯" & "/USD " & Format((Val(Text8) - Val(Text3)), FDollar) & "')"
            Else
               If douDebitAmount < 0 Then
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p18, a1p28, a1p29, a1p19, a1p20, a1p21, a1p14) values ('1', 'K', '" & GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'", 3) & "', '" & Text9 & "', '110205', '" & MsgText(55) & "', " & douDebitAmount * (-1) & ", 0, " & Val(strSrvDate(2)) & ", " & Val(strSrvDate(2)) & ", " & ServerTime & ", 'USD', " & douBalance & ", " & (Val(Text8) - Val(Text3)) * (-1) & ", '" & "結匯" & "/USD " & Format((Val(Text8) - Val(Text3)) * (-1), FDollar) & "')"
               End If
            End If
         Else
            If douDebitAmount > 0 Then
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p18, a1p28, a1p29, a1p19, a1p20, a1p21, a1p14) values ('1', 'K', '" & GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'", 3) & "', '" & Text9 & "', '7128', '" & MsgText(55) & "', 0, " & douDebitAmount & ", " & Val(strSrvDate(2)) & ", " & Val(strSrvDate(2)) & ", " & ServerTime & ", 'USD', " & douBalance & ", " & (Val(Text8) - Val(Text3)) & ", '" & "結匯" & "/USD " & Format((Val(Text8) - Val(Text3)), FDollar) & "')"
            Else
               If douDebitAmount < 0 Then
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p18, a1p28, a1p29, a1p19, a1p20, a1p21, a1p14) values ('1', 'K', '" & GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'", 3) & "', '" & Text9 & "', '7128', '" & MsgText(55) & "', " & douDebitAmount * (-1) & ", 0, " & Val(strSrvDate(2)) & ", " & Val(strSrvDate(2)) & ", " & ServerTime & ", 'USD', " & douBalance & ", " & (Val(Text8) - Val(Text3)) * (-1) & ", '" & "結匯" & "/USD " & Format((Val(Text8) - Val(Text3)) * (-1), FDollar) & "')"
               End If
            End If
         End If
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select sum(a1p07) as Debit, sum(a1p08) as Credit from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "' and a1p05 <> '7128'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            If Text10 <> MsgText(602) Then
               If adoquery.Fields("Debit").Value < adoaccsum.Fields("Credit").Value Then
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p18, a1p28, a1p29, a1p19, a1p20, a1p21, a1p14) values ('1', 'K', '" & GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'", 3) & "', '" & Text9 & "', '7128', '" & MsgText(55) & "', " & Val(adoquery.Fields("Credit").Value) - Val(adoquery.Fields("Debit").Value) & ", 0, " & Val(strSrvDate(2)) & ", " & Val(strSrvDate(2)) & ", " & ServerTime & ", 'NTD', 1, " & Val(adoquery.Fields("Credit").Value) - Val(adoquery.Fields("Debit").Value) & ", '結匯')"
               Else
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p18, a1p28, a1p29, a1p19, a1p20, a1p21, a1p14) values ('1', 'K', '" & GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'", 3) & "', '" & Text9 & "', '7128', '" & MsgText(55) & "', 0, " & Val(adoquery.Fields("Debit").Value) - Val(adoquery.Fields("Credit").Value) & ", " & Val(strSrvDate(2)) & ", " & Val(strSrvDate(2)) & ", " & ServerTime & ", 'NTD', 1, " & Val(adoquery.Fields("Credit").Value) - Val(adoquery.Fields("Debit").Value) & ", '結匯')"
               End If
            Else
               If adoquery.Fields("Debit").Value < adoaccsum.Fields("Credit").Value Then
                  adoTaie.Execute "update acc1p0 set a1p07 = " & Val(adoquery.Fields("Credit").Value) - Val(adoquery.Fields("Debit").Value) & ", a1p08 = 0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "' and a1p05 = '7128'"
               Else
                  adoTaie.Execute "update acc1p0 set a1p08 = " & Val(adoquery.Fields("Debit").Value) - Val(adoquery.Fields("Credit").Value) & ", a1p07 = 0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "' and a1p05 = '7128'"
               End If
            End If
         End If
         adoquery.Close
      End If
   End If
   adoaccsum.Close
   
   If m_bolAlert Then MsgBox "下列請款單號有分配點數，請確認結果是否無誤！" & vbCrLf & vbCrLf & m_strAlertMsg
End Sub

'Add by Morgan 2011/9/6
'產生貸方分錄(有分配資料)
Private Sub Acc1p0Save()
   Dim douService As Double
   Dim douFee As Double
   Dim strCaseNo As String '本所案號
   Dim strCaseProperty As String '案件性質
   Dim strSalesMan As String '智權人員
   Dim strCurrency As String '請款幣別
   Dim strExchange As String '請款匯率
   Dim strSerialNo As String '分錄序次
   Dim strSystemType As String '系統別
   Dim strAccNo As String '科目
   Dim strDept As String '承辦人會計部門
   Dim strEngDept As String '承辦人部門
   Dim strSalesDept As String '智權人員部門
   Dim strCustNo As String '客戶編號
   Dim strProperty As String '案件性質碼
   Dim strR As String '收入科目
   Dim strF As String '規費科目
   Dim strNation As String '申請國家
   Dim bolXFee As Boolean '服務費是否含出庭費
   Dim bolXFeeDone As Boolean '出庭費是否已扣除
   Dim strCP09 As String '收文號
   Dim strA1P30 As String '對沖-其他
   Dim strAmt As String '分錄金額
   Dim strAmtTot As String '收款服務費金額
   Dim strA1p14 As String '摘要
   Dim strA1p08 As String '借方金額
   Dim strAmtRest As String '未分配金額
   Dim adoacc1n0 As ADODB.Recordset
   Dim strPtTot As String '請款單總點數
   Dim strA1p16s As String '智權人代碼清單
   Dim strSerialNoFrom As String '分錄序次起號
   Dim strNetAmount As String '可分配金額(收款點數會大於請款點數)
   Dim strShareP As String
   Dim strShareT As String
   Dim strShareL As String
   Dim strShareFCP As String
   Dim strShareFCT As String
   Dim strShareFCL As String
   Dim strSharePointMemo As String '跨部門點數分配摘要
   Dim strMemoFrom As String '本次更新項次
   
   '收文資料
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select * from caseprogress, casepropertymap where cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp60 = '" & Adodc1.Recordset.Fields("a1k01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.EOF = False Then
      
      strSalesDept = "" & adoaccsum.Fields("cp12").Value 'Added by Morgan 2012/9/19
      If IsNull(adoaccsum.Fields("cpm03").Value) Then
         If IsNull(adoaccsum.Fields("cpm04").Value) Then
            strCaseProperty = ""
         Else
            strCaseProperty = adoaccsum.Fields("cpm04").Value
         End If
      Else
         strCaseProperty = adoaccsum.Fields("cpm03").Value
      End If
      strCaseNo = adoaccsum.Fields("cp01").Value & adoaccsum.Fields("cp02").Value & adoaccsum.Fields("cp03").Value & adoaccsum.Fields("cp04").Value
      If IsNull(adoaccsum.Fields("cp13").Value) Then
         strSalesMan = ""
      Else
         strSalesMan = adoaccsum.Fields("cp13").Value
      End If

      strA1p14 = strCaseNo & "/" & strCaseProperty
      
      '申請人編號,申請國家
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select pa26 as CustNo, PA09 AS strNation from patent where pa01 = '" & adoaccsum.Fields("cp01").Value & "' and pa02 = '" & adoaccsum.Fields("cp02").Value & "' and pa03 = '" & adoaccsum.Fields("cp03").Value & "' and pa04 = '" & adoaccsum.Fields("cp04").Value & "' union " & _
                    "select tm23 as CustNo, TM10 AS strNation from trademark where tm01 = '" & adoaccsum.Fields("cp01").Value & "' and tm02 = '" & adoaccsum.Fields("cp02").Value & "' and tm03 = '" & adoaccsum.Fields("cp03").Value & "' and tm04 = '" & adoaccsum.Fields("cp04").Value & "' union " & _
                    "select lc11 as CustNo, LC15 AS strNation from lawcase where lc01 = '" & adoaccsum.Fields("cp01").Value & "' and lc02 = '" & adoaccsum.Fields("cp02").Value & "' and lc03 = '" & adoaccsum.Fields("cp03").Value & "' and lc04 = '" & adoaccsum.Fields("cp04").Value & "' union " & _
                    "select hc05 as CustNo, '000' AS strNation from hirecase where hc01 = '" & adoaccsum.Fields("cp01").Value & "' and hc02 = '" & adoaccsum.Fields("cp02").Value & "' and hc03 = '" & adoaccsum.Fields("cp03").Value & "' and hc04 = '" & adoaccsum.Fields("cp04").Value & "' union " & _
                    "select sp08 as CustNo, SP09 AS strNation from servicepractice where sp01 = '" & adoaccsum.Fields("cp01").Value & "' and sp02 = '" & adoaccsum.Fields("cp02").Value & "' and sp03 = '" & adoaccsum.Fields("cp03").Value & "' and sp04 = '" & adoaccsum.Fields("cp04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         If IsNull(adoquery.Fields("CustNo").Value) Then
            strCustNo = ""
         Else
            strCustNo = adoquery.Fields("CustNo").Value
         End If
         If IsNull(adoquery.Fields("strNation").Value) Then
            strNation = ""
         Else
            strNation = adoquery.Fields("strNation").Value
         End If
      Else
         strCustNo = ""
         strNation = ""
      End If
      adoquery.Close
      
      '預設科目
      With adoaccsum
      If strNation <> "000" Then
         strR = "" & .Fields("cpm24").Value
         strF = "" & .Fields("cpm25").Value
      Else
         strR = "" & .Fields("cpm11").Value
         strF = "" & .Fields("cpm12").Value
      End If
      
      '台灣案專利商標出庭費控管
      Do While Not .EOF
         strSystemType = "" & .Fields("cp01").Value
         If strNation = "000" And Val(Adodc1.Recordset("a1k02")) >= 960815 Then
            strProperty = "" & .Fields("cp10").Value
            strCP09 = "" & .Fields("cp09").Value
            '專利
            If (strSystemType = "P" Or strSystemType = "FCP") Then
               If InStr("211,212", strProperty) > 0 Then
                  bolXFee = True
                  Exit Do
               ElseIf InStr("503,507,506", strProperty) > 0 Then
                  strExc(0) = "select * from caseprogress a where cp09='" & strCP09 & "'" & _
                     " and exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05>=a.cp05 and b.cp10 in ('211','212'))" & _
                     " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05>=a.cp05 and b.cp10 in ('211','212') and b.cp16>0)"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     bolXFee = True
                     Exit Do
                  End If
               End If
            '商標(FCT不扣)
            ElseIf (strSystemType = "T") Then
               If InStr("204,205", strProperty) > 0 Then
                  bolXFee = True
                  Exit Do
               ElseIf InStr("403,408,407", strProperty) > 0 Then
                  strExc(0) = "select * from caseprogress a where cp09='" & strCP09 & "'" & _
                     " and exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05>=a.cp05 and b.cp10 in ('204','205'))" & _
                     " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05>=a.cp05 and b.cp10 in ('204','205') and b.cp16>0)"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     bolXFee = True
                     Exit Do
                  End If
               End If
            End If
         End If
         .MoveNext
      Loop
      .MoveFirst
      End With
            
      With Adodc1.Recordset
               
      '固定用美金收款
      'strCurrency = "" & .Fields("a1k18")
      'strExchange = "" & .Fields("a1k10")
      strCurrency = "USD"
      strExchange = Val(Text2)
      
      douFee = Val(.Fields("a1k09").Value)
      'Added by Lydia 2016/07/18 抵帳作業的規費只到整數
      If Left(strF, 4) = "2201" Then douFee = Val(Format(douFee, DAmount))
      
      '此處的 a1k11 = round((a1k08-nvl(a1k06, 0)) * " & Val(Text2) & ", 2)
      'Modified by Lydia 2016/07/18
      'douService = Val(.Fields("a1k11")) - Val(.Fields("a1k09").Value)
      douService = Val(.Fields("a1k11")) - douFee
      
      strAmtTot = douService
      strNetAmount = strAmtTot
      '是否要扣出庭費10000
      If bolXFee = True Then
         strAmtTot = Val(strAmtTot - 10000)
      End If
      strAmtRest = strAmtTot
      
      'add by sonia 2016/6/13 法務收入改科目
      '收款
      'Modified by Morgan 2020/4/15 請款單日期>=智慧所更名日者改回依案件性質表設定之科目收入；
      If DBDATE(.Fields("a1k02")) < 智慧所更名日 And Val(strSrvDate(2)) > 1050000 And (Left(strR, 4) = "4141" Or Left(strR, 4) = "4161" Or Left(strR, 4) = "4181") Then
         'modify by sonia 2021/3/12 加傳日期
         strSalesMan = SalesNoToAccSales(strSalesMan, strR, strCaseNo, strSrvDate(2))
         If strSalesMan = "" Then
            strSalesMan = "M0100"
         End If
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'", 3)
         InsertLawACC1P0 "1", "K", strSerialNo, strItemNo, strR, IIf(strDept = "", MsgText(55), strDept), 0, Val(strAmtTot), "", "", "", "", "", _
            strA1p14 & IIf(strR = "416101", "/" & .Fields("A1K01").Value & "/" & strCurrency & Format("" & .Fields("A1K08").Value, "0.00"), ""), _
            strCustNo, strSalesMan, strCaseNo, strSrvDate(2), strCurrency, strExchange, "" & Format(Val(strAmtTot) / Val(Text2), FAmount), "", Adodc1.Recordset.Fields("a1k01").Value, "", "", "", "", "", "", ""
      Else
      'end 2016/6/13
      
         '部門,點數
         strExc(0) = "select a0910,max(st03) st03,sum(a1n05) pts from acc1n0,staff,acc090" & _
            " where a1n01='" & .Fields("a1k01") & "' and a1n02='2' and st01(+)=a1n04" & _
            " and a0901(+)=st15 group by a0910 order by 3 desc,2,1"
         intI = 1
         Set adoacc1n0 = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
         
            strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'", 3)
            strMemoFrom = strSerialNo
            '請款單點數
            'Modified by Morgan 2015/10/15 要抓請款點數非收款點數
            'strPtTot = Val(strAmtTot) / 1000
            strPtTot = Val("" & .Fields("SFee")) / 1000
            'end 2015/10/15
            Do While Not adoacc1n0.EOF
               strAccNo = strR
               '只有一筆時不必分配
               If adoacc1n0.RecordCount = 1 Then
                  strAmt = strAmtTot
               '尚有未分配點數時
               Else
                  m_bolAlert = True '承辦點數有跨部門
                  If InStr(m_strAlertMsg, .Fields("a1k01")) = 0 Then m_strAlertMsg = m_strAlertMsg & vbCrLf & .Fields("a1k01")
                  
                  If adoacc1n0.AbsolutePosition > 1 Then
                     strSerialNo = Format(Val(strSerialNo) + 1, "000")
                  End If
                  
                  If strAmtRest > 0 Then
                     '若有出庭費固定扣第一筆(點數最多的)
                     If bolXFee = True And bolXFeeDone = False Then
                        strAmt = Round(strNetAmount * adoacc1n0.Fields("pts") / strPtTot, 2) - 10000
                        bolXFeeDone = True
                        
                     'Added by Morgan 2015/10/15
                     ElseIf adoacc1n0.AbsolutePosition = adoacc1n0.RecordCount Then
                        strAmt = strAmtRest
                     'end 2015/10/15
                     
                     '不足額也照比例分
                     Else
                        strAmt = Round(strNetAmount * adoacc1n0.Fields("pts") / strPtTot, 2)
                     End If
                  Else
                     strAmt = 0
                  End If
               End If
               
               strAmtRest = strAmtRest - strAmt
               
               '承辦人會計部門
               strDept = "" & adoacc1n0.Fields("a0910")
               '承辦人部門
               strEngDept = "" & adoacc1n0.Fields("st03")
               
               'FMP
               If strR = "411103" And Left(strSalesDept, 1) = "F" Then
                  If strDept = "FCP" Then
                     strAccNo = "417102"
                  Else
                     strAccNo = "411106"
                  End If
               'FMT
               ElseIf strR = "410103" And Left(strSalesDept, 1) = "F" Then
                  strAccNo = "410109"
            
               '417201 FCT收入,若為內商人員承辦時改科目為 417202 FCT爭議
               '417202 FCT爭議,若為國外部承辦時改科目為 417201 FCT收入
               'Modify by Morgan 2010/6/21 非 FCT,T 時要依跨部門規則
               ElseIf strSystemType = "FCT" And strDept = "T" Then
                  strAccNo = "417202"
                  
               ElseIf strSystemType = "FCT" And strDept = "FCT" Then
                  strAccNo = "417201"
               'end 2010/6/21
               
               'Add by Morgan 2010/10/8 CFT&FCT 或 CFL&FCL  不算跨部門
               ElseIf strSystemType = "CFT" And strDept = "FCT" Then
                  strDept = strSystemType
               ElseIf strSystemType = "CFL" And strDept = "FCL" Then
                  strDept = strSystemType
               'end 2010/10/8
               
               'Add by Morgan 2010/10/27 CFP & P 不算跨部門
               ElseIf strSystemType = "CFP" And strDept = "P" Then
                  strDept = strSystemType
               '跨部門分點數
               ElseIf strSystemType <> strDept Then
                  Select Case strDept
                     Case "P"
                        strAccNo = "411101"
                        strShareP = Val(strShareP) + Val(strAmt)
                     Case "T"
                        strAccNo = "410101"
                        strShareT = Val(strShareT) + Val(strAmt)
                     Case "L"
                        strAccNo = "414101"
                        strShareL = Val(strShareL) + Val(strAmt)
                     Case "FCP"
                        'modify by sonia 2016/8/4
                        'strAccNo = "417101"
                        Select Case strSystemType
                           Case "FCL", "CFL", "LIN"
                              'modify by sonia 2022/3/3 417103-->417109
                              strAccNo = "417109"
                           Case Else
                              strAccNo = "417101"
                        End Select
                        'end 2016/8/4
                        strShareFCP = Val(strShareFCP) + Val(strAmt)
                     Case "FCT"
                        'modify by sonia 2016/8/4
                        'strAccNo = "417201"
                        Select Case strSystemType
                           Case "FCL", "CFL", "LIN"
                              'modify by sonia 2022/3/3 417203-->417202
                              strAccNo = "417202"
                           Case Else
                              strAccNo = "417201"
                        End Select
                        'end 2016/8/4
                        strShareFCT = Val(strShareFCT) + Val(strAmt)
                     Case "FCL"
                        strAccNo = "416101"
                        strShareFCL = Val(strShareFCL) + Val(strAmt)
                  End Select
               End If
               
               strA1p16s = ""
               strSerialNoFrom = strSerialNo
               strExc(0) = "select a1n04,sum(a1n05) pts from acc1n0" & _
                  " where a1n01='" & .Fields("a1k01") & "' and a1n02='1'" & _
                  " group by a1n04 order by 2,1"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  Do While Not RsTemp.EOF
                     strSalesMan = "" & RsTemp("a1n04")
                     '作帳智權人代碼
                     'modify by sonia 2021/3/12 加傳日期
                     strSalesMan = SalesNoToAccSales(strSalesMan, strAccNo, strCaseNo, strSrvDate(2))
                     If strSalesMan = "" Then
                        strSalesMan = "M0100"
                        
                     End If
                     '只有一筆時不必分配
                     If RsTemp.RecordCount = 1 Then
                        strA1p08 = strAmt
                        '2012/10/16 MODIFY BY SONIA 加存A1P23,轉帳務時才能判斷資料來源Z10100007(P097900000)為貸方退費,Z10100029(P098968000)為貸方收款
                        'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p21) values " & _
                           "('1', 'K', '" & strSerialNo & "', '" & Text9 & "', '" & strAccNo & "', 0, " & strA1p08 & ", '" & strA1p14 & IIf(strAccNo = "416101", "/" & .Fields("A1K01").Value & "/" & strCurrency & Format("" & .Fields("A1K08").Value, FAmount), "") & "', '" & strSalesMan & "', '" & strCaseNo & "', " & strSrvDate(2) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & Format(strA1p08 / Val(Text2), FAmount) & ")"
                        adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p21, a1p23) values " & _
                           "('1', 'K', '" & strSerialNo & "', '" & Text9 & "', '" & strAccNo & "', 0, " & strA1p08 & ", '" & strA1p14 & IIf(strAccNo = "416101", "/" & .Fields("A1K01").Value & "/" & strCurrency & Format("" & .Fields("A1K08").Value, FAmount), "") & "', '" & strSalesMan & "', '" & strCaseNo & "', " & strSrvDate(2) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & Format(strA1p08 / Val(Text2), FAmount) & ",'" & Adodc1.Recordset.Fields("a1k01").Value & "')"
                     Else
                        strA1p08 = Round(strAmt * RsTemp("pts") / strPtTot, 2)
                        '若分配智權人的作帳智權人代碼已有資料時累加
                        If InStr(strA1p16s, strSalesMan) > 0 Then
                           adoTaie.Execute "update acc1p0 set a1p08=a1p08+" & strA1p08 & " where a1p01='1' and a1p02='K' and a1p03>='" & strSerialNoFrom & "' and a1p04='" & Text9 & "' and a1p16='" & strSalesMan & "'"
                        Else
                           If RsTemp.AbsolutePosition > 1 Then
                              strSerialNo = Format(Val(strSerialNo) + 1, "000")
                           End If
                           '2012/10/16 MODIFY BY SONIA 加存A1P23,轉帳務時才能判斷資料來源Z10100007(P097900000)為貸方退費,Z10100029(P098968000)為貸方收款
                           'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p21) values " & _
                              "('1', 'K', '" & strSerialNo & "', '" & Text9 & "', '" & strAccNo & "', 0, " & strA1p08 & ", '" & strA1p14 & IIf(strAccNo = "416101", "/" & .Fields("A1K01").Value & "/" & strCurrency & Format("" & .Fields("A1K08").Value, FAmount), "") & "', '" & strSalesMan & "', '" & strCaseNo & "', " & strSrvDate(2) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & Format(strA1p08 / Val(Text2), FAmount) & ")"
                           adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p21, a1p23) values " & _
                              "('1', 'K', '" & strSerialNo & "', '" & Text9 & "', '" & strAccNo & "', 0, " & strA1p08 & ", '" & strA1p14 & IIf(strAccNo = "416101", "/" & .Fields("A1K01").Value & "/" & strCurrency & Format("" & .Fields("A1K08").Value, FAmount), "") & "', '" & strSalesMan & "', '" & strCaseNo & "', " & strSrvDate(2) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & IIf(strDept = "", MsgText(55), strDept) & "', '" & strCustNo & "', " & Format(strA1p08 / Val(Text2), FAmount) & ",'" & Adodc1.Recordset.Fields("a1k01").Value & "')"
                           strA1p16s = strA1p16s & "," & strSalesMan
                        End If
                     End If
                     strAmt = strAmt - strA1p08
                     RsTemp.MoveNext
                  Loop
                  '差額放在最後一筆
                  If Val(strAmt) > 0 Then
                     adoTaie.Execute "update acc1p0 set a1p08=a1p08+" & strAmt & " where a1p01='1' and a1p02='K' and a1p03='" & strSerialNo & "' and a1p04='" & Text9 & "'"
                  End If
                  
                  '若項次有增加表示智權人員點數有做分配
                  If strSerialNo <> strSerialNoFrom Then
                     m_bolAlert = True
                     If InStr(m_strAlertMsg, .Fields("a1k01")) = 0 Then m_strAlertMsg = m_strAlertMsg & vbCrLf & .Fields("a1k01")
                  End If
                  
                  'Added by Morgan 2022/10/20 FCP收入再細分科目
                  If Left(strAccNo, 4) = "4171" And strAccNo <> "417102" And strAccNo <> "417103" Then
                     UpdateFCPACC1P0 "1", "K", strSerialNo, Text9, strAccNo, .Fields("a1k01"), strNation
                  End If
                  'end 2022/10/20
                        
               End If
               adoacc1n0.MoveNext
            Loop
            'Add by Morgan 2010/6/11
            '點數分配摘要
            If Val(strShareP) > 0 Then
               strSharePointMemo = strSharePointMemo & " P" & Round(100 * Val(strShareP) / Val(strAmtTot), 0) & "%"
            End If
            If Val(strShareT) > 0 Then
               strSharePointMemo = strSharePointMemo & " T" & Round(100 * Val(strShareT) / Val(strAmtTot), 0) & "%"
            End If
            If Val(strShareL) > 0 Then
               strSharePointMemo = strSharePointMemo & " L" & Round(100 * Val(strShareL) / Val(strAmtTot), 0) & "%"
            End If
            If Val(strShareFCP) > 0 Then
               strSharePointMemo = strSharePointMemo & " FCP" & Round(100 * Val(strShareFCP) / Val(strAmtTot), 0) & "%"
            End If
            If Val(strShareFCT) > 0 Then
               strSharePointMemo = strSharePointMemo & " FCT" & Round(100 * Val(strShareFCT) / Val(strAmtTot), 0) & "%"
            End If
            If Val(strShareFCL) > 0 Then
               strSharePointMemo = strSharePointMemo & " FCL" & Round(100 * Val(strShareFCL) / Val(strAmtTot), 0) & "%"
            End If
            If strSharePointMemo <> "" Then
               adoTaie.Execute "update acc1p0 set a1p14 = a1p14 ||'/'||'" & Trim(strSharePointMemo) & "' where a1p01='1' and a1p02='K' and a1p03>='" & strMemoFrom & "' and a1p04='" & Text9 & "' and substr(a1p05,1,1)='4' and a1p17='" & strCaseNo & "'"
            End If
         End If
      End If
      
      '規費
      strAccNo = strF
      
      Select Case strSystemType
         Case "S"
            If strNation = "000" Then
               strDept = "FCT"
            Else
               strDept = "CFT"
            End If
         Case "T", "TF"
            strDept = "T"
         Case "P", "PS"
            strDept = "P"
         Case "FCT"
            strDept = "FCT"
         Case "FCP", "FG"
            strDept = "FCP"
         Case "CFT", "CFC"
            strDept = "CFT"
         Case "CFP", "CPS"
            strDept = "CFP"
         Case "L"
            strDept = "L"
         Case "FCL", "CFL"
            strDept = "FCL"
         Case Else
            strDept = "T"
      End Select
      
      '台灣案專利商標出庭費控管
      If bolXFee = True Then
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'", 3)
         '2012/10/16 MODIFY BY SONIA 加存A1P23,轉帳務時才能判斷資料來源Z10100007(P097900000)為貸方退費,Z10100029(P098968000)為貸方收款
         'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p21) values " & _
            "('1', 'K', '" & strSerialNo & "', '" & Text9 & "', '" & strAccNo & "', 0, 10000, '" & strA1p14 & "/出庭費', null, '" & strCaseNo & "', " & strSrvDate(2) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & Format(10000 / Val(Text2), FAmount) & ")"
         adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14, a1p16, a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p21, a1p23) values " & _
            "('1', 'K', '" & strSerialNo & "', '" & Text9 & "', '" & strAccNo & "', 0, 10000, '" & strA1p14 & "/出庭費', null, '" & strCaseNo & "', " & strSrvDate(2) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & Format(10000 / Val(Text2), FAmount) & ",'" & Adodc1.Recordset.Fields("a1k01").Value & "')"
      End If
      
      '規費>0
      If douFee > 0 Then
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & Text9 & "'", 3)
         If strAccNo = "220105" Or strAccNo = "220106" Then
             'CFT, CFP摘要帶的金額為總額(規費+服務費)
             strRemark = strA1p14 & "/" & IIf(strDept = "CFT" Or strDept = "CFP", (douService + douFee), douFee)
         ElseIf strAccNo = "220111" Or strAccNo = "220112" Then
            '大陸專利商標摘要帶的金額為總額(規費+服務費)
            strRemark = strA1p14 & "/" & (douService + douFee)
         Else
            strRemark = strA1p14
         End If
         
         '2012/10/16 MODIFY BY SONIA 加存A1P23,轉帳務時才能判斷資料來源Z10100007(P097900000)為貸方退費,Z10100029(P098968000)為貸方收款
         'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14,  a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p21, a1p16) values " & _
            "('1', 'K', '" & strSerialNo & "', '" & Text9 & "', '" & strAccNo & "', 0, " & douFee & ", '" & strRemark & "','" & strCaseNo & "', " & strSrvDate(2) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & Format(douFee / Val(Text2), FAmount) & ", '" & strSalesMan & "')"
         'modify by sonia 2017/1/17 規費科目不放智權人員a1p16
         adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p07, a1p08, a1p14,  a1p17, a1p18, a1p19, a1p20, a1p06, a1p15, a1p21, a1p23) values " & _
            "('1', 'K', '" & strSerialNo & "', '" & Text9 & "', '" & strAccNo & "', 0, " & douFee & ", '" & strRemark & "','" & strCaseNo & "', " & strSrvDate(2) & ", '" & strCurrency & "', " & Val(strExchange) & ", '" & MsgText(55) & "', '" & strCustNo & "', " & Format(douFee / Val(Text2), FAmount) & ",'" & Adodc1.Recordset.Fields("a1k01").Value & "')"
         
      End If
      adoaccsum.Close
      End With
   End If
   Set adoacc1n0 = Nothing
End Sub

'Add by Amy 2014/11/05 由aacc_sav搬回
'Modified by Lydia 2024/09/03 +bolA1811付款單
Public Sub Frmacc21f0_Save(Optional ByVal bolA1811 As Boolean = False)
   On Error GoTo Checking
   With Frmacc21f0
      If .Text9 = MsgText(601) Then
         MsgBox MsgText(10) & .Label7, , MsgText(5)
         strControlButton = MsgText(602)
         .Text9.SetFocus
         Exit Sub
      End If
      If strSaveConfirm = MsgText(3) Then
         If .adoacc1g0.RecordCount <> 0 Then
            .adoacc1g0.Find "a1g01 = '" & .Text9 & "'", 0, adSearchForward, 1
            If .adoacc1g0.EOF = False Then
'               MsgBox MsgText(9), , MsgText(5)
'               strControlButton = MsgText(602)
'               .Text9.SetFocus
               Exit Sub
            End If
         End If
         .adoacc1g0.AddNew
      End If
      .adoacc1g0.Fields("a1g01").Value = .Text9
      If .Text2 <> MsgText(601) Then   'FC匯率
         .adoacc1g0.Fields("a1g02").Value = Val(.Text2)
      Else
         .adoacc1g0.Fields("a1g02").Value = 0
      End If
      If .Text6 <> MsgText(601) Then   'CF匯率
         .adoacc1g0.Fields("a1g03").Value = Val(.Text6)
      Else
         .adoacc1g0.Fields("a1g03").Value = 0
      End If
      If .Text10 <> MsgText(601) Then  '是否結清
         .adoacc1g0.Fields("a1g10").Value = .Text10
      Else
         .adoacc1g0.Fields("a1g10").Value = Null
      End If
      If strSaveConfirm = MsgText(3) Then
         .adoacc1g0.Fields("a1g04").Value = Val(strSrvDate(2))
         .adoacc1g0.Fields("a1g05").Value = ServerTime
         .adoacc1g0.Fields("a1g06").Value = strUserNum
      Else
         .adoacc1g0.Fields("a1g07").Value = Val(strSrvDate(2))
         .adoacc1g0.Fields("a1g08").Value = ServerTime
         .adoacc1g0.Fields("a1g09").Value = strUserNum
      End If
      .adoacc1g0.UpdateBatch
        'Modify By Cheng 2004/03/23
'        .Adodc1.Recordset.MoveFirst
        If bolA1811 = False Then 'Added by Lydia 2024/09/03
           If .Adodc1.Recordset.RecordCount > 0 Then .Adodc1.Recordset.MoveFirst
           'End
           While Not .Adodc1.Recordset.EOF
   '            adoTaie.Execute "update acc1k0 set a1k08 = " & Val("" & .Adodc1.Recordset("A1K08").Value) & " Where a1k01 = '" & .Adodc1.Recordset("A1K01").Value & "'"
               adoTaie.Execute "update acc1k0 set a1k29 = '" & MsgText(602) & "', a1k30 = round(" & Val("" & .Adodc1.Recordset("A1K08").Value) & " * " & Val(.Text2) & ", 2) where a1k01 = '" & .Adodc1.Recordset("A1K01").Value & "' And a1k17 = '" & .Text9 & "'"
               .Adodc1.Recordset.MoveNext
           Wend
           'End
   '      adoTaie.Execute "update acc1k0 set a1k29 = '" & MsgText(602) & "', a1k30 = round(a1k08 * " & Val(.Text2) & ", 2) where a1k17 = '" & .Text9 & "'"
        End If 'Added by Lydia 2024/09/03
      .RecordShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub
'Added by Lydia 2016/07/18 判斷規費的會計科目
Private Function GetFeeAccNo(ByVal iCP01 As String, iNa01 As String) As String
    Select Case iCP01
       Case "T", "TS"
          If iNa01 <> "000" Then
             GetFeeAccNo = "220111"
          End If
       Case "P", "PS"
          If iNa01 <> "000" Then
             GetFeeAccNo = "220112"
          End If
       Case "S"
          Select Case iNa01
             Case "000"
                GetFeeAccNo = "220103"
             Case "020"
                GetFeeAccNo = "220112"
             Case Else
                GetFeeAccNo = "220105"
          End Select
    End Select
End Function

'Added by Lydia 2024/09/03 匯入付款單=6.抵帳
Private Sub Command6_Click()

Dim strTmpA As String
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   If Left(Text5, 1) = "W" Then
      strExc(0) = "select acc180.*, acc190.* from acc180, acc190 where a1801='" & Trim(Text5) & "' and a1801=a1901(+) and a1908 is null and a1811='6' order by a1902 "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 0 Then
         MsgBox "付款單號的匯款方式非抵帳！"
      Else
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            '參考KeyDefine2的檢查
            If "" & RsTemp.Fields("a1917") = "J" Then '直接用付款單的設定
               MsgBox RsTemp.Fields("a1902") & "抵帳不可為J公司！"
               GoTo JumpToNext
            End If
            If "" & RsTemp.Fields("a1803") <> "" Then
                If GetDizhang(Left("" & RsTemp("a1803"), 8), , False, 2, False) = "不同意抵帳" Then
                    MsgBox "代理人「" & Left(RsTemp("a1803"), 8) & "」不同意抵帳！"
                    Exit Sub
                End If
            End If
            If RsTemp.AbsolutePosition = 1 Then
               If Text13 = "" Then
                  Text13 = "" & RsTemp.Fields("a1803")
               ElseIf "" & RsTemp("a1803") <> Text13 Then
                  If MsgBox("此付款單之代理人與上方代理人不符，仍要做此筆帳單的抵帳嗎？ Y：維持，N：取消設定" & vbCrLf, vbYesNo + vbDefaultButton2) = vbNo Then
                     Exit Sub
                  End If
               End If
            End If

            strTmpA = strTmpA & "," & RsTemp.Fields("a1902")
JumpToNext:
            RsTemp.MoveNext
         Loop
         
         If strTmpA <> "" Then
            strTmpA = Mid(strTmpA, 2)
            
            Frmacc21f0_Save True
            If strControlButton <> MsgText(602) Then
               '先刪除付款單
               strSql = "delete from acc170 where a1702 in (select a1902 from acc190 where a1901='" & Text5 & "' ) and a1701='1' "
               cnnConnection.Execute strSql, intI
               strSql = "delete from acc190 where a1901='" & Text5 & "' "
               cnnConnection.Execute strSql, intI
               strSql = "delete from acc180 where a1801='" & Text5 & "' "
               cnnConnection.Execute strSql, intI
                              
               Acc150Save strTmpA
            End If
            If strControlButton <> MsgText(602) Then
               SumShow2
               Text5 = "U"
               Text5_GotFocus
            End If
            strControlButton = MsgText(601)
         End If
      End If
   Else
      MsgBox "請輸入付款單號！"
   End If
End Sub
