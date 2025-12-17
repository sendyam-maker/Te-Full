VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc1180 
   AutoRedraw      =   -1  'True
   Caption         =   "付款作業"
   ClientHeight    =   5712
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5712
   ScaleWidth      =   8760
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   7140
      Picture         =   "Frmacc1180.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   69
      ToolTipText     =   "清除畫面"
      Top             =   3090
      Width           =   400
   End
   Begin VB.ComboBox Combo7 
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
      ItemData        =   "Frmacc1180.frx":08CA
      Left            =   6450
      List            =   "Frmacc1180.frx":08CC
      Style           =   2  '單純下拉式
      TabIndex        =   7
      Top             =   380
      Width           =   2030
   End
   Begin VB.ComboBox Combo6 
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
      ItemData        =   "Frmacc1180.frx":08CE
      Left            =   6930
      List            =   "Frmacc1180.frx":08D8
      Style           =   2  '單純下拉式
      TabIndex        =   67
      Top             =   1380
      Width           =   1545
   End
   Begin VB.TextBox Text22 
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
      MaxLength       =   8
      TabIndex        =   26
      Top             =   5370
      Width           =   1572
   End
   Begin VB.TextBox Text21 
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
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   25
      Top             =   5340
      Width           =   1572
   End
   Begin VB.TextBox Text19 
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
      Left            =   1335
      MaxLength       =   12
      TabIndex        =   24
      Top             =   5340
      Width           =   1572
   End
   Begin VB.TextBox Text18 
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
      Left            =   1335
      MaxLength       =   3
      TabIndex        =   23
      Top             =   5040
      Width           =   528
   End
   Begin VB.TextBox Text17 
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
      Left            =   4545
      TabIndex        =   60
      Top             =   30
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox Text20 
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
      Height          =   300
      Left            =   1230
      MaxLength       =   12
      TabIndex        =   57
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox Text15 
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
      Height          =   300
      Left            =   5580
      TabIndex        =   56
      Top             =   3120
      Width           =   1332
   End
   Begin VB.TextBox Text13 
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
      Left            =   6840
      MaxLength       =   14
      TabIndex        =   13
      Top             =   3768
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   2580
      Picture         =   "Frmacc1180.frx":08E8
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.ComboBox Combo4 
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
      TabIndex        =   22
      Top             =   4692
      Width           =   1575
   End
   Begin VB.TextBox Text7 
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
      Height          =   315
      Left            =   4080
      TabIndex        =   21
      Top             =   4692
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
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
      TabIndex        =   14
      Top             =   4068
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   1335
      TabIndex        =   20
      Top             =   4692
      Width           =   1750
   End
   Begin VB.ComboBox Combo1 
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
      TabIndex        =   19
      Top             =   4380
      Width           =   1575
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00C0FFFF&
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
      Left            =   360
      MaxLength       =   6
      TabIndex        =   11
      Top             =   3768
      Width           =   1572
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc1180.frx":09EA
      Height          =   1005
      Left            =   150
      TabIndex        =   28
      Top             =   2070
      Width           =   8430
      _ExtentX        =   14880
      _ExtentY        =   1778
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   20
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
         Name            =   "新細明體"
         Size            =   10.8
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
      BeginProperty Column02 
         DataField       =   "a1p07"
         Caption         =   "借方金額"
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
      BeginProperty Column03 
         DataField       =   "a1p08"
         Caption         =   "貸方金額"
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
         DataField       =   "a1p18"
         Caption         =   "開票日期"
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
         DataField       =   "a1p12"
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
      BeginProperty Column06 
         DataField       =   "a1p09"
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
      BeginProperty Column07 
         DataField       =   "a1p13"
         Caption         =   "票別"
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
         DataField       =   "a1p24"
         Caption         =   "領款方式"
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
            ColumnWidth     =   2448
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1344.189
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1332.284
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1368
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1235.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1607.811
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   587.906
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   1031.811
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   150
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
   Begin VB.TextBox Text10 
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
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   16
      Top             =   4068
      Width           =   1572
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
      Height          =   400
      Left            =   7590
      Picture         =   "Frmacc1180.frx":09FF
      Style           =   1  '圖片外觀
      TabIndex        =   27
      ToolTipText     =   "取消"
      Top             =   3090
      Width           =   400
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
      Height          =   300
      Left            =   6840
      MaxLength       =   8
      TabIndex        =   15
      Top             =   4068
      Width           =   1572
   End
   Begin VB.TextBox Text8 
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
      Left            =   5040
      MaxLength       =   14
      TabIndex        =   12
      Top             =   3768
      Width           =   1575
   End
   Begin VB.TextBox Text6 
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
      Height          =   300
      Left            =   4230
      TabIndex        =   40
      Top             =   3120
      Width           =   1308
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
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
      Left            =   1380
      MaxLength       =   9
      TabIndex        =   6
      Top             =   1056
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   1380
      MaxLength       =   1
      TabIndex        =   0
      Top             =   48
      Width           =   612
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1380
      TabIndex        =   2
      Top             =   390
      Width           =   1575
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3300
      TabIndex        =   3
      Top             =   390
      Width           =   1575
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
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   300
      Left            =   1380
      TabIndex        =   4
      Top             =   720
      Width           =   1575
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
   Begin MSMask.MaskEdBox MaskEdBox4 
      Height          =   300
      Left            =   3300
      TabIndex        =   5
      Top             =   720
      Width           =   1575
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
   Begin MSMask.MaskEdBox MaskEdBox5 
      Height          =   300
      Left            =   1335
      TabIndex        =   17
      Top             =   4395
      Width           =   1575
      _ExtentX        =   2794
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
   Begin MSMask.MaskEdBox MaskEdBox6 
      Height          =   300
      Left            =   4080
      TabIndex        =   18
      Top             =   4380
      Width           =   1575
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
   Begin MSMask.MaskEdBox MaskEdBox7 
      Height          =   300
      Left            =   6450
      TabIndex        =   1
      Top             =   45
      Width           =   1530
      _ExtentX        =   2688
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   12648447
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
   Begin MSForms.ComboBox Combo5 
      Height          =   330
      Left            =   4080
      TabIndex        =   71
      Top             =   5010
      Width           =   4335
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "7638;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text23 
      Height          =   600
      Left            =   6930
      TabIndex        =   70
      Top             =   750
      Width           =   1545
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "2725;1058"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text16 
      Height          =   315
      Left            =   1380
      TabIndex        =   10
      Top             =   1728
      Width           =   4110
      VariousPropertyBits=   679493659
      BackColor       =   16777215
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text14 
      Height          =   300
      Left            =   1950
      TabIndex        =   55
      Top             =   3765
      Width           =   2775
      VariousPropertyBits=   679493659
      BackColor       =   14737632
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text4 
      Height          =   315
      Left            =   1380
      TabIndex        =   9
      Top             =   1392
      Width           =   4110
      VariousPropertyBits=   679493659
      BackColor       =   16777215
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text3 
      Height          =   315
      Left            =   2940
      TabIndex        =   36
      Top             =   1056
      Width           =   3945
      VariousPropertyBits=   679493663
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      Caption         =   "智財局退費款項類別請選其他"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   5490
      TabIndex        =   68
      Top             =   1770
      Width           =   3120
   End
   Begin VB.Label Label32 
      BackStyle       =   0  '透明
      Caption         =   "主要付款銀行"
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
      Left            =   5520
      TabIndex        =   66
      Top             =   1440
      Width           =   1350
   End
   Begin VB.Label Label31 
      BackStyle       =   0  '透明
      Caption         =   "付款備註"
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
      Left            =   5520
      TabIndex        =   65
      Top             =   750
      Width           =   1215
   End
   Begin VB.Label Label30 
      BackStyle       =   0  '透明
      Caption         =   "對沖(業)"
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
      Left            =   5880
      TabIndex        =   64
      Top             =   5370
      Width           =   975
   End
   Begin VB.Label Label29 
      BackStyle       =   0  '透明
      Caption         =   "對沖(其)"
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
      Left            =   3120
      TabIndex        =   63
      Top             =   5340
      Width           =   975
   End
   Begin VB.Label Label28 
      BackStyle       =   0  '透明
      Caption         =   "對沖(本)"
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
      Left            =   375
      TabIndex        =   62
      Top             =   5340
      Width           =   975
   End
   Begin VB.Label Label27 
      BackStyle       =   0  '透明
      Caption         =   "部門別"
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
      Left            =   375
      TabIndex        =   61
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label26 
      BackStyle       =   0  '透明
      Caption         =   "聯絡地址 "
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
      Left            =   180
      TabIndex        =   59
      Top             =   1725
      Width           =   1215
   End
   Begin VB.Label Label25 
      BackStyle       =   0  '透明
      Caption         =   "筆數"
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
      Left            =   270
      TabIndex        =   58
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label24 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "貸方金額"
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
      Left            =   6840
      TabIndex        =   54
      Top             =   3525
      Width           =   1575
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "款項類別"
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
      Left            =   5880
      TabIndex        =   53
      Top             =   4695
      Width           =   975
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "手續費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3120
      TabIndex        =   52
      Top             =   4695
      Width           =   675
   End
   Begin VB.Label Label22 
      BackStyle       =   0  '透明
      Caption         =   "備註"
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
      Left            =   3120
      TabIndex        =   51
      Top             =   5010
      Width           =   975
   End
   Begin VB.Label Label21 
      BackStyle       =   0  '透明
      Caption         =   "領款方式"
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
      Left            =   375
      TabIndex        =   50
      Top             =   4695
      Width           =   975
   End
   Begin VB.Label Label20 
      BackStyle       =   0  '透明
      Caption         =   "票別"
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
      Left            =   5880
      TabIndex        =   49
      Top             =   4380
      Width           =   975
   End
   Begin VB.Label Label19 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "會計科目"
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
      Left            =   480
      TabIndex        =   48
      Top             =   3525
      Width           =   4215
   End
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "開票銀行"
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
      Left            =   3120
      TabIndex        =   47
      Top             =   4065
      Width           =   975
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "付款日期"
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
      Left            =   5480
      TabIndex        =   46
      Top             =   45
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   -60
      Top             =   3360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   2220
      Left            =   180
      Top             =   3495
      Width           =   8430
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "到期日期"
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
      Left            =   3120
      TabIndex        =   45
      Top             =   4380
      Width           =   975
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "開票日期"
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
      Left            =   375
      TabIndex        =   44
      Top             =   4395
      Width           =   975
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "票據號碼"
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
      Left            =   5880
      TabIndex        =   43
      Top             =   4065
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "借方金額"
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
      Left            =   5040
      TabIndex        =   42
      Top             =   3525
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "開票帳號"
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
      TabIndex        =   41
      Top             =   4065
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "合計"
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
      Left            =   3390
      TabIndex        =   39
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "應付總金額"
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
      TabIndex        =   38
      Top             =   420
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "支票抬頭"
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
      Left            =   180
      TabIndex        =   37
      Top             =   1395
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "往來對象"
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
      Left            =   180
      TabIndex        =   35
      Top             =   1050
      Width           =   1215
   End
   Begin VB.Label Label5 
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
      Left            =   3060
      TabIndex        =   34
      Top             =   720
      Width           =   255
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
      Height          =   255
      Left            =   180
      TabIndex        =   33
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label7 
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
      Left            =   3060
      TabIndex        =   32
      Top             =   390
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "入帳日期"
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
      Left            =   180
      TabIndex        =   31
      Top             =   390
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "(1.廠商 2.客戶 3.員工)"
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
      Left            =   2100
      TabIndex        =   30
      Top             =   45
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "往來類別"
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
      Left            =   180
      TabIndex        =   29
      Top             =   45
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc1180"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/08/20 Form2.0已修改 Text3/Text4/Text10/Text14/Text16/Text23/Combo5
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoacc0q0 As New ADODB.Recordset
Public adoacc1p0 As New ADODB.Recordset
Public adoacc0o0 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoacctotal As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adocase As New ADODB.Recordset
Dim strSerialNo As String
Public strDocNo As String
Public lngDate As Long

Dim bNoCreditSide As Boolean
Dim StrA0H01 As String, StrA0H02 As String 'Add by Amy 2014/02/10

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Combo2_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Combo3_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Combo4_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'Modify by Amy 2021/08/20 改Form2.0 原:Integer
Private Sub Combo5_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyDefine Val(KeyCode)
End Sub

'Add by Amy 2014/01/21 由Text2_Validate搬過來並改寫
Private Sub Combo7_Change()
    If strSaveConfirm <> MsgText(3) Then
        Exit Sub
    End If
    If Combo7.ListCount = 0 Then
        Exit Sub
    End If
    
    SetCheckTitle 'Add by Morgan 2006/10/27 預設支票抬頭
End Sub

'由Text2_Validate搬過來並改寫
Private Sub Combo7_Validate(Cancel As Boolean)
    If strSaveConfirm <> MsgText(3) Then
        Exit Sub
    End If
    If Combo7.ListCount = 0 Then
        Exit Sub
    End If

    If Combo7.ListCount > 0 And Combo7 = "" Then
        MsgBox Label9 & "必要欄位,請選取...", , MsgText(5)
        Cancel = True
        Combo7.SetFocus
        Exit Sub
    End If

    Dim strSQL1 As String, strSql As String, adoTmp As ADODB.Recordset
    
    'Modify by Amy 2021/04/12 從下面搬上來,因先輸L公司資料,換輸1公司資料時,應帶華銀-辜
    'Add by Amy 2021/03/18 L公司付款銀行預帶瑞興(原:大台北)
    If Mid(Combo7, 1, 1) = "L" Then
        Combo6.ListIndex = 1
    Else
        Combo6.ListIndex = 0
    End If
    
    strExc(1) = Mid(Combo7, 1, 1)
    strSql = GetStrWhere
    'Add by Morgan 2009/8/31
    'strSQL1 = "select a1p05,sum(a1p08) a1p08,a1p14 from acc1p0, acc0o0 where a1p04 = a0o01 and a1p01 = '" & strExc(0) & "' and a1p02 = 'B' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a1p15 = '" & Text2 & "'" & strSql & " group by a1p04,a1p05,a1p14" & _
                    " union all select a1p05,sum(a1p08) a1p08,a1p14 from acc1p0, acc0o0 where a1p23 = a0o01 and a1p01 = '" & strExc(0) & "' and a1p02 = 'E' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a0o03 = '" & Text2 & "'" & strSql & " group by a1p04,a1p05,a1p14" & _
                    " union all select a1p05,sum(a1p08) a1p08,a1p14 from acc1p0, acc0o0 where a1p23 = a0o01 and a1p01 = '" & strExc(0) & "' and a1p02 = 'Z' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a0o03 = '" & Text2 & "'" & strSql & " group by a1p04,a1p05,a1p14"
    '2014/02/10 +if及公司別 並修改語法讓速度變快
    If strExc(1) = "J" Then
        strSQL1 = "select a1p05,sum(a1p08) a1p08,a1p14 from acc1p0, acc0o0 where a1p01(+) =a0o07 and a1p04(+) = a0o01 and a1p01 = '" & strExc(1) & "' and a1p02= 'B' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a0o03 = '" & Text2 & "' And a1p04 not in (Select  a2505 From acc250 Where a2505(+)=a0o01 And a2510 is null And a2519 is null And a0o19='2') " & strSql & " group by a1p04,a1p05,a1p14" & _
                        " union all select a1p05,sum(a1p08) a1p08,a1p14 from acc1p0, acc0o0 where a1p01(+) = a0o07 and a1p23(+) = a0o01 and a1p01 = '" & strExc(1) & "' and a1p02 = 'E' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a0o03 = '" & Text2 & "' And a1p04 not in (Select  a2505 From acc250 Where a2505(+)=a0o01 And a2510 is null And a2519 is null And a0o19='2') " & strSql & " group by a1p04,a1p05,a1p14" & _
                        " union all select a1p05,sum(a1p08) a1p08,a1p14 from acc1p0, acc0o0 where a1p01(+) = a0o07 and a1p23(+) = a0o01 and a1p01 = '" & strExc(1) & "' and a1p02 = 'Z' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a0o03 = '" & Text2 & "' And a1p04 not in (Select  a2505 From acc250 Where a2505(+)=a0o01 And a2510 is null And a2519 is null And a0o19='2') " & strSql & " group by a1p04,a1p05,a1p14"
    Else
        strSQL1 = "select a1p05,sum(a1p08) a1p08,a1p14 from acc1p0, acc0o0 where a1p01(+) =a0o07 and a1p04(+) = a0o01 and a1p01 = '" & strExc(1) & "' and a1p02 = 'B' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a0o03 = '" & Text2 & "'" & strSql & " group by a1p04,a1p05,a1p14" & _
                        " union all select a1p05,sum(a1p08) a1p08,a1p14 from acc1p0, acc0o0 where a1p01(+) = a0o07 and a1p23(+) = a0o01 and a1p01 = '" & strExc(1) & "' and a1p02 = 'E' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a0o03 = '" & Text2 & "'" & strSql & " group by a1p04,a1p05,a1p14" & _
                        " union all select a1p05,sum(a1p08) a1p08,a1p14 from acc1p0, acc0o0 where a1p01(+) = a0o07 and a1p23(+) = a0o01 and a1p01 = '" & strExc(1) & "' and a1p02 = 'Z' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a0o03 = '" & Text2 & "'" & strSql & " group by a1p04,a1p05,a1p14"
    End If
    'end 2014/02/10
    intI = 1
    Set adoTmp = ClsLawReadRstMsg(intI, strSQL1)
    If intI = 1 Then
        With adoTmp
            Do While Not .EOF
                .MoveNext
                If .EOF Then
                    bNoCreditSide = False
                Else
                    bNoCreditSide = True
                End If
                .MovePrevious

                Text11 = .Fields("a1p05")
                Text8 = .Fields("a1p08")
                Combo5 = .Fields("a1p14")
                KeyDefine vbKeyInsert
                .MoveNext
            Loop
        End With
    End If
    Combo7.Tag = Mid(Combo7, 1, 1) 'Add by Amy 2020/07/01
    
    bNoCreditSide = False
    'end 2009/8/31
End Sub

'Add by Amy 2020/07/03
Private Sub Command1_Click()
    AdodcClear
End Sub

Private Sub Command2_Click()
   If Adodc1.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc1.Recordset.Fields("a1p22").Value) = False Then
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select ax210 from acc021 where ax201 = '" & Adodc1.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc1.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(155), , MsgText(5)
            Text11.SetFocus
            adoquery.Close
            Exit Sub
         End If
         adoquery.Close
      End If
   End If
   AdodcDelete
   'Add by Amy 2014/01/21 新增時acc1p0 有資料不可改往來對象及應付總金額
   If strSaveConfirm = MsgText(3) Then
        If CheckAcc1p0(Mid(Combo7, 1, 1), Text17) = True Then
            Text2.Enabled = False
            Combo7.Enabled = False
        Else
            Text2.Enabled = True
            Combo7.Enabled = True
        End If
   End If
   'end 2014/01/21
End Sub

Private Sub Command2_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Command3_Click()
   'Modify by Amy 2014/01/21 查詢鈕(望遠鏡)隱藏,因無從得知公司別
   Acc0q0Refresh
   If adoacc0q0.RecordCount <> 0 Then
      FormShow
      AdodcRefresh
      SumShow
      RecordShow
   End If
End Sub

Private Sub Command3_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command3_Click
         Exit Sub
   End Select
   KeyDefine KeyCode
End Sub

Private Sub DataGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
   Select Case ColIndex
      Case 2
         DataGrid1.Columns(2) = OldValue
   End Select
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   strSerialNo = Adodc1.Recordset.Fields("a1p03").Value
   AdodcShow
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strItemNo = MsgText(601) Or strCustNo = MsgText(601) Then
      Exit Sub
   End If
   MaskEdBox7.Mask = ""
   MaskEdBox7.Text = CFDate(strItemNo)
   MaskEdBox7.Mask = DFormat
   Combo7.Tag = strCompanyNo 'Add by Amy 2014/01/21 +公司別
   Text2 = strCustNo
   Acc0q0Refresh
   If adoacc0q0.RecordCount <> 0 Then
      FormShow
      AdodcRefresh
      SumShow
      RecordShow
      'Add by Amy 2014/09/30 +a1p22有值不可修改付款日
      If CheckExistA1p22(Mid(Combo7, 1, 1), "C", Text17) = True Then
         MaskEdBox7.Enabled = False
      Else
         MaskEdBox7.Enabled = True
      End If
      'end 2014/09/30
   End If
   strCompanyNo = MsgText(601) 'Add by Amy 2014/01/21 +公司別
   strItemNo = MsgText(601)
   strCustNo = MsgText(601)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
   'Modify by Amy 2023/07/19 調整大小
   'PUB_InitForm Me, 8850, 6100, strBackPicPath1
   PUB_InitForm Me, 9000, 6450, strBackPicPath1
   
   strItemNo = MsgText(601)
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   MaskEdBox3.Mask = DFormat
   MaskEdBox4.Mask = DFormat
   MaskEdBox5.Mask = DFormat
   MaskEdBox6.Mask = DFormat
   MaskEdBox7.Mask = DFormat
   Combo1.AddItem ComboItem(11)
   Combo1.AddItem ComboItem(12)
   Combo1.AddItem ComboItem(13)
   Combo2.AddItem ComboItem(81)
   Combo2.AddItem ComboItem(82)
   Combo2.AddItem ComboItem(83)
   Combo2.AddItem ComboItem(84)
   Combo2.AddItem ComboItem(85)
   Combo2.AddItem "6--寄出地址特別" 'Add by Morgan 2006/10/31
   Combo4.AddItem ComboItem(111)
   Combo4.AddItem ComboItem(112)
   Combo4.AddItem ComboItem(113)
   Combo4.AddItem ComboItem(114)
   Combo4.AddItem ComboItem(115)
   Combo4.AddItem ComboItem(116)
   Combo4.AddItem ComboItem(117)
   
   FormDisabled
   OpenTable
   If adoacc0q0.RecordCount <> 0 Then
      adoacc0q0.MoveLast
      adoacc0q0.MoveFirst
      RecordShow
   End If
   
   Combo6.ListIndex = 0 'Added by Morgan 2011/11/15 預設華銀
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   If AmountCheck <> MsgText(602) Then
      tool1_enabled
      MsgBox MsgText(30), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc1180 = Nothing
End Sub

Private Sub MaskEdBox1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label3 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
End Sub

Private Sub MaskEdBox2_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
   If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
      Exit Sub
   End If
   If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
      MsgBox Label3 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox2.SetFocus
      Exit Sub
   End If
End Sub

Private Sub MaskEdBox3_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub MaskEdBox3_Validate(Cancel As Boolean)
   If MaskEdBox3.Text = MsgText(601) Or MaskEdBox3.Text = MsgText(29) Then
      Exit Sub
   End If
   If DateCheck(MaskEdBox3.Text) = MsgText(603) Then
      MsgBox Label4 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox3.SetFocus
      Exit Sub
   End If
End Sub

Private Sub MaskEdBox4_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub MaskEdBox4_Validate(Cancel As Boolean)
   If MaskEdBox4.Text = MsgText(601) Or MaskEdBox4.Text = MsgText(29) Then
      Exit Sub
   End If
   If DateCheck(MaskEdBox4.Text) = MsgText(603) Then
      MsgBox Label4 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox4.SetFocus
      Exit Sub
   End If
End Sub

Private Sub MaskEdBox5_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub MaskEdBox5_Validate(Cancel As Boolean)
   If MaskEdBox5.Text = MsgText(601) Or MaskEdBox5.Text = MsgText(29) Then
      Exit Sub
   End If
   If DateCheck(MaskEdBox5.Text) = MsgText(603) Then
      MsgBox Label14 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox5.SetFocus
      Exit Sub
   End If
End Sub

Private Sub MaskEdBox6_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub MaskEdBox6_Validate(Cancel As Boolean)
   If MaskEdBox6.Text = MsgText(601) Or MaskEdBox6.Text = MsgText(29) Then
      Exit Sub
   End If
   If DateCheck(MaskEdBox6.Text) = MsgText(603) Then
      MsgBox Label5 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox6.SetFocus
      Exit Sub
   End If
   Select Case Text11
      Case "2111", "110201", "110214"
         Combo5 = IIf(MaskEdBox6.Text <> MsgText(29), FCDate(MaskEdBox6.Text), "") & "/" & Text9 & "/" & Text4
   End Select
End Sub

Private Sub MaskEdBox7_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub MaskEdBox7_Validate(Cancel As Boolean)
   If MaskEdBox7.Text = MsgText(601) Or MaskEdBox7.Text = MsgText(29) Then
      MsgBox Label16 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox7.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox7.Text) = MsgText(603) Then
      MsgBox Label16 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox7.SetFocus
      Exit Sub
   End If
   If strSaveConfirm <> MsgText(3) Then
      Exit Sub
   End If
   If Mid(MaskEdBox7.Text, 1, 3) <> Mid(CFDate(ACDate(ServerDate)), 1, 3) Then
      Text17 = UpdateNo("acc0q0", "a0q17", 4, MaskEdBox7.Text, MsgText(818))
   Else
      Text17 = strDocNo
   End If
End Sub
'93.12.8 cancel by sonia 辜說不限制
'Private Sub Text1_Change()
'   If Text1 = Mid(ComboItem(91), 1, 1) Or Text1 = Mid(ComboItem(93), 1, 1) Then
'      MaskEdBox1.Enabled = False
'      MaskEdBox2.Enabled = False
'      MaskEdBox3.Enabled = True
'      MaskEdBox4.Enabled = True
'   Else
'      MaskEdBox1.Enabled = True
'      MaskEdBox2.Enabled = True
'      MaskEdBox3.Enabled = False
'      MaskEdBox4.Enabled = False
'   End If
'End Sub
'93.12.8 end

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

'2009/12/10 add by sonia
Private Sub Text1_KeyPress(KeyAscii As Integer)
   If (KeyAscii < Asc("1") Or KeyAscii > Asc("3")) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub
'2009/12/10 end

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc0q0.CursorLocation = adUseClient
   adoacc0q0.MaxRecords = intMax
   'Modify by Amy 2014/01/21  +公司別且不先load資料避免過久
   'adoacc0q0.Open "select * from acc0q0 where a0q01||a0q03 >= '" & Val(FCDate(MaskEdBox7.Text)) & Text2 & "' order by a0q17 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   strExc(0) = "select * from acc0q0 where a0q01= " & Val(FCDate(MaskEdBox7.Text)) & " And a0q03 = '" & Text2 & "' And a0q19='' order by a0q17 asc"
   adoacc0q0.Open strExc(0), adoTaie, adOpenDynamic, adLockBatchOptimistic
   
   adoacc1p0.CursorLocation = adUseClient
   '2014/01/21 改公司別 原:'1'
   strExc(0) = "select * from acc1p0 where a1p01 = '' and a1p02 = 'C' and a1p03 = '" & Text2 & "'" & _
                  " and a1p04 = '" & FCDate(MaskEdBox7.Text) & "' order by a1p05 asc"
    adoacc1p0.Open strExc(0), adoTaie, adOpenDynamic, adLockBatchOptimistic
    
   adoacc0o0.CursorLocation = adUseClient
   adoacc0o0.Open "select * from acc0o0 Where a0o07='' order by a0o01 desc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   
   adoadodc1.CursorLocation = adUseClient
   strExc(0) = "select * from acc1p0 where a1p01 = '' and a1p02 = 'C' and a1p03 = '" & Text2 & "'" & _
                    " and a1p04 = '" & FCDate(MaskEdBox7.Text) & "' order by a1p05 asc"
   adoadodc1.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
   'end 2014/01/21
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表(國內付款資料(主檔))
'
'*************************************************
Public Sub FormShow()
   Dim strCmp As String 'Add by Amy 2020/04/09
   
   If IsNull(adoacc0q0.Fields("a0q04").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = adoacc0q0.Fields("a0q04").Value
   End If
   If IsNull(adoacc0q0.Fields("a0q17").Value) Then
      Text17 = MsgText(601)
   Else
      Text17 = adoacc0q0.Fields("a0q17").Value
   End If
   Text23 = "" & adoacc0q0.Fields("a0q18").Value 'Add by Morgan 2006/10/31
   
   lngDate = adoacc0q0.Fields("a0q01").Value
   MaskEdBox7.Mask = MsgText(601)
   MaskEdBox7.Text = CFDate(adoacc0q0.Fields("a0q01").Value)
   MaskEdBox7.Mask = DFormat
   MaskEdBox1.Mask = MsgText(601)
   MaskEdBox1.Text = MsgText(601)
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = MsgText(601)
   MaskEdBox2.Text = MsgText(601)
   MaskEdBox2.Mask = DFormat
   MaskEdBox3.Mask = MsgText(601)
   MaskEdBox3.Text = MsgText(601)
   MaskEdBox3.Mask = DFormat
   MaskEdBox4.Mask = MsgText(601)
   MaskEdBox4.Text = MsgText(601)
   MaskEdBox4.Mask = DFormat
   'Modify by Amy 2014/01/21 改為combo7顯示(+公司別) 原:Text5
   If IsNull(adoacc0q0.Fields("a0q06").Value) Then
      Combo7 = MsgText(601)
   Else
      Combo7.Clear
      'Modify by Amy 2020/04/09 公司名稱改抓function 原:IIf(adoacc0q0.Fields("a0q19") = "1", "台一", "智權")
      strCmp = adoacc0q0.Fields("a0q19") & "." & A0802Query(adoacc0q0.Fields("a0q19"), True)
      Combo7.AddItem strCmp & ":" & adoacc0q0.Fields("a0q06").Value
      Combo7 = strCmp & ":" & adoacc0q0.Fields("a0q06").Value
      'end 2020/04/09
   End If
   'end 2014/01/21
   Text2 = adoacc0q0.Fields("a0q03").Value
   If IsNull(adoacc0q0.Fields("a0q05").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = adoacc0q0.Fields("a0q05").Value
   End If
   If IsNull(adoacc0q0.Fields("a0q16").Value) Then
      Text16 = MsgText(601)
   Else
      Text16 = adoacc0q0.Fields("a0q16").Value
   End If
   Text16.Tag = Text16 'Add by Amy 2015/11/10
   Select Case Text1
      Case Mid(ComboItem(91), 1, 1)
         Text3 = A0i02Query(Text2)
      Case Mid(ComboItem(92), 1, 1)
         If Len(Text2) = 6 Then
            Text2 = AfterZero(Text2)
         'Add by Morgan 2007/3/1 八碼時要補'0'
         ElseIf Len(Text2) = 8 Then
            Text2 = Text2 & "0"
         'End 2007/3/1
         End If
         Text3 = CustomerQuery(Text2, 1)
      Case Mid(ComboItem(93), 1, 1)
         Text3 = StaffQuery(Text2)
      Case Else
         Text3 = MsgText(601)
   End Select
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub

Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 <> MsgText(601) Then
      If ExistCheck("acc0g0", "a0g01", Text10, Label18) = False Then
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub Text11_Change()
   Text14 = A0102Query(Text11)
   'Add by Amy 2014/01/21
   If Text14 = MsgText(601) Or strSaveConfirm = MsgText(601) Then
        Exit Sub
   End If
   
   'Modify 2014/02/10 以會計科目抓Acc0H0開票帳號及銀行,若有資料鎖住
   StrA0H01 = ""
   StrA0H02 = GetBankData(Text11, StrA0H01)
   If StrA0H02 = MsgText(601) And StrA0H01 = MsgText(601) Then
        Combo3.Enabled = True
        Text10.Enabled = True
    Else
        Combo3 = StrA0H02
        Text10 = StrA0H01
        Combo3.Enabled = False
        Text10.Enabled = False
    End If
    'end 2014/02/10
   'end 2014/01/21
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub

Private Sub Text11_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
   If Text11 <> MsgText(601) Then
      'Modify by Amy 2011/01/21 if判斷的Function
      'If ExistCheck("acc010", "a0101", Text11, Label19) = False Then
      If PUB_CheckCompany(Text11, Mid(Combo7, 1, 1)) = False Then
         Cancel = True
         Exit Sub
      End If
   End If
   Select Case Text11
      Case "2111", "110201", "110214"
         Combo5 = IIf(MaskEdBox6.Text <> MsgText(29), FCDate(MaskEdBox6.Text), "") & "/" & Text9 & "/" & Text4
   End Select
   
   'Added by Morgan 2006/10/30 檢查款項類別及領款方式
   Select Case Text11
      'Modified by Morgan 2011/10/4+"1101", "110202", "1911", "1912", "1913"
      'Modified by Morgan 2011/11/15 +110207
      'modify by sonia 2020/5/12 +110602
      Case "2111", "1101", "110202", "1911", "1912", "1913", "110207", "110602"
         SetPayAffair
   End Select
   
   Select Case Text11
      Case "2112", "2113"
         adoquery.CursorLocation = adUseClient
         '93.9.14 MODIFY BY SONIA
         'adoquery.Open "select * from acc1p0, acc0o0 where a1p04 = a0o01 and a1p05 = '" & Text11 & "' and a0o03 = '" & Text2 & "' and (a0o11 is null or a0o11 = " & Val(FCDate(MaskEdBox7.Text)) & ") union " & _
         '              "select * from acc1p0, acc0o0 where a1p23 = a0o01 and a1p05 = '2112' and a0o03 = '" & Text2 & "' and (a0o11 is null or a0o11 = " & Val(FCDate(MaskEdBox7.Text)) & ")", adoTaie, adOpenStatic, adLockReadOnly
         'Modify by Amy 2014/01/21 +公司別
         adoquery.Open "select * from acc1p0, acc0o0 where a1p01='" & Mid(Combo7, 1, 1) & "' And A1P02='B' AND a1p04 = a0o01 and a1p05 = '" & Text11 & "' and a0o03 = '" & Text2 & "' and (a0o11 is null or a0o11 = " & Val(FCDate(MaskEdBox7.Text)) & ") union " & _
                       "select * from acc1p0, acc0o0 where a1p01='" & Mid(Combo7, 1, 1) & "' And A1P02 IN ('E','Z') AND a1p23 = a0o01 and a1p05 = '2112' and a0o03 = '" & Text2 & "' and (a0o11 is null or a0o11 = " & Val(FCDate(MaskEdBox7.Text)) & ")", adoTaie, adOpenStatic, adLockReadOnly
         '93.9.14 END
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields("a1p14").Value) = False Then
               Combo5 = adoquery.Fields("a1p14").Value
            Else
               Combo5 = MsgText(601)
            End If
         Else
            Combo5 = MsgText(601)
         End If
         adoquery.Close
   End Select
   'add by sonia 2021/1/29 以本所案號以判別FCP,FCT英日文組
   If AccNoToSalesNo(Text11, Text19) <> "" Then
      Text22 = AccNoToSalesNo(Text11, Text19)
   End If
   'end 2021/1/29
End Sub

Private Sub Text13_GotFocus()
   TextInverse Text13
End Sub

Private Sub Text13_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text16_GotFocus()
   StatusView MsgText(65) & "100"
   TextInverse Text16
   'edit by nickc 2007/06/11  切換輸入法改用API
   OpenIme
End Sub

'Modify by Amy 2021/08/20 改Form2.0 原:Integer
Private Sub Text16_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyDefine Val(KeyCode)
End Sub

Private Sub Text16_LostFocus()
   StatusView MsgText(601)
End Sub

Private Sub Text16_Validate(Cancel As Boolean)
   If CheckLen(Label26, Text16, 100) = MsgText(603) Then
      Cancel = True
      Text16.SetFocus
      Exit Sub
   End If
   'edit by nickc 2007/06/11  切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub

Private Sub Text18_GotFocus()
   TextInverse Text18
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text18_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text18_Validate(Cancel As Boolean)
   If Text18 <> MsgText(601) Then
      If ExistCheck("acc090", "a0901", Text18, Label27) = False Then
         Cancel = True
         Exit Sub
      End If
   End If
   If CheckDept(Text11, Text18) = False Then
      MsgBox MsgText(103), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text19_GotFocus()
   TextInverse Text19
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text19_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text19_Validate(Cancel As Boolean)
On Error GoTo Checking
   If Text19 <> MsgText(601) Then
      Text19 = CaseNoZero(Text19)
      adocase.CursorLocation = adUseClient
      adocase.Open "select pa01 as SystemNo from patent where pa01 = '" & Mid(Text19, 1, Len(Text19) - 9) & "' and pa02 = '" & Mid(Text19, Len(Text19) - 8, 6) & "' and pa03 = '" & Mid(Text19, Len(Text19) - 2, 1) & "' and pa04 = '" & Mid(Text19, Len(Text19) - 1, 2) & "' union " & _
                   "select tm01 as SystemNo from trademark where tm01 = '" & Mid(Text19, 1, Len(Text19) - 9) & "' and tm02 = '" & Mid(Text19, Len(Text19) - 8, 6) & "' and tm03 = '" & Mid(Text19, Len(Text19) - 2, 1) & "' and tm04 = '" & Mid(Text19, Len(Text19) - 1, 2) & "' union " & _
                   "select lc01 as SystemNo from lawcase where lc01 = '" & Mid(Text19, 1, Len(Text19) - 9) & "' and lc02 = '" & Mid(Text19, Len(Text19) - 8, 6) & "' and lc03 = '" & Mid(Text19, Len(Text19) - 2, 1) & "' and lc04 = '" & Mid(Text19, Len(Text19) - 1, 2) & "' union " & _
                   "select hc01 as SystemNo from hirecase where hc01 = '" & Mid(Text19, 1, Len(Text19) - 9) & "' and hc02 = '" & Mid(Text19, Len(Text19) - 8, 6) & "' and hc03 = '" & Mid(Text19, Len(Text19) - 2, 1) & "' and hc04 = '" & Mid(Text19, Len(Text19) - 1, 2) & "' union " & _
                   "select sp01 as SystemNo from servicepractice where sp01 = '" & Mid(Text19, 1, Len(Text19) - 9) & "' and sp02 = '" & Mid(Text19, Len(Text19) - 8, 6) & "' and sp03 = '" & Mid(Text19, Len(Text19) - 2, 1) & "' and sp04 = '" & Mid(Text19, Len(Text19) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocase.RecordCount = 0 Then
         MsgBox MsgText(28) & Label28, , MsgText(5)
         Cancel = True
         adocase.Close
         Exit Sub
      End If
      adocase.Close
      'add by sonia 2021/1/29 以本所案號以判別FCP,FCT英日文組
      If AccNoToSalesNo(Text11, Text19) <> "" Then
         Text22 = AccNoToSalesNo(Text11, Text19)
      End If
      'end 2021/1/29
   End If
   Exit Sub
Checking:
   MsgBox MsgText(128), , MsgText(5)
   Exit Sub
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Mark by Amy 2014/01/21 沒使用
'Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
'   Select Case KeyCode
'      Case vbKeyF12
'         Select Case Text1
'            Case Mid(ComboItem(91), 1, 1)
'               Calculate1
'            Case Mid(ComboItem(93), 1, 1)
'               Calculate1
'            Case Else
'               Calculate2
'         End Select
'         Exit Sub
'   End Select
'   KeyDefine KeyCode
'End Sub
'end 2014/01/21

Private Sub Text2_Validate(Cancel As Boolean)
   Dim strSql As String
   'Dim strSQL1 As String, adoTmp As ADODB.Recordset 'Mark 2014/01/21

   If Text2 = MsgText(601) Then
      Exit Sub
   End If
   If Text1 <> MsgText(601) Then
      Select Case Text1
         Case Mid(ComboItem(91), 1, 1)
            If ExistCheck("acc0i0", "a0i01", Text2, Label6) = False Then
               Cancel = True
               Exit Sub
            End If
         Case Mid(ComboItem(92), 1, 1)
            If ExistCheck("customer", "cu01", Mid(IIf(Len(Text2) = 6, AfterZero(Text2), Text2), 1, 8), Label6) = False Then
               Cancel = True
               Exit Sub
            End If
         Case Mid(ComboItem(93), 1, 1)
            If ExistCheck("staff", "st01", Text2, Label6) = False Then
               Cancel = True
               Exit Sub
            End If
      End Select
   End If
   Select Case Text1
      Case Mid(ComboItem(91), 1, 1)
         Text3 = A0i02Query(Text2)
         adoquery.CursorLocation = adUseClient
         'Modify by Morgan 2009/1/21 廠商郵遞區號欄位自地址欄拆開a0i04
         adoquery.Open "select a0i04||a0i03 from acc0i0 where a0i01 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields(0).Value) Then
               Text16 = ""
            Else
               Text16 = adoquery.Fields(0).Value
            End If
         Else
            Text16 = ""
         End If
         adoquery.Close
      Case Mid(ComboItem(92), 1, 1)
         If Len(Text2) = 6 Then
            Text2 = AfterZero(Text2)
         ElseIf Len(Text2) = 8 Then
            Text2 = Text2 & "0"
         End If
         Text3 = CustomerQuery(Text2, 1)
         
         adoquery.CursorLocation = adUseClient
         'Modify by Morgan 2005/10/17加郵遞區號
         'adoquery.Open "select cu31, cu23 from customer where cu01 = '" & Mid(Text2, 1, 8) & "' and cu02 = '" & Mid(Text2, 9, 1) & "'", adoTaie, adOpenStatic, adLockReadOnly
         '2011/1/21 MODIFY BY SONIA加CU23的郵遞區號CU112
         'adoquery.Open "select cu30||cu31, cu23,cu80 from customer where cu01 = '" & Mid(Text2, 1, 8) & "' and cu02 = '" & Mid(Text2, 9, 1) & "'", adoTaie, adOpenStatic, adLockReadOnly
         adoquery.Open "select cu30||cu31, CU112||cu23,cu80 from customer where cu01 = '" & Mid(Text2, 1, 8) & "' and cu02 = '" & Mid(Text2, 9, 1) & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            'Add by Morgan 2007/1/22 加判斷客戶狀態
            If Not IsNull(adoquery.Fields(2)) Then
               Text16 = adoquery.Fields(2)
            'end 2007/1/22
            ElseIf IsNull(adoquery.Fields(0).Value) Then
               If IsNull(adoquery.Fields(1).Value) Then
                  Text16 = ""
               Else
                  Text16 = adoquery.Fields(1).Value
               End If
            Else
               Text16 = adoquery.Fields(0).Value
            End If
         Else
            Text16 = ""
         End If
         adoquery.Close
      Case Mid(ComboItem(93), 1, 1)
         Text3 = StaffQuery(Text2)
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select st08 from staff where st01 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields(0).Value) Then
               Text16 = ""
            Else
               Text16 = adoquery.Fields(0).Value
            End If
         Else
            Text16 = ""
         End If
         adoquery.Close
      Case Else
         Text3 = MsgText(601)
   End Select
   Text4 = Text3
   strSql = strSql & GetStrWhere 'Modify by Amy 原程式寫至function
   
   If strSaveConfirm = MsgText(3) Then
      'Modify by Amy 2014/01/21
      If Combo7 = MsgText(601) Then
        adoaccsum.CursorLocation = adUseClient
       'Modify by Amy 2014/02/10
'       adoaccsum.Open "select sum(nvl(Amount, 0)) from (select sum(a1p08) as Amount from acc1p0, acc0o0 where a1p04 = a0o01 and a1p01 = '1' and a1p02 = 'B' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a1p15 = '" & Text2 & "'" & strSql & _
'                     " union all select sum(a1p08) as Amount from acc1p0, acc0o0 where a1p23 = a0o01 and a1p01 = '1' and a1p02 = 'E' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a0o03 = '" & Text2 & "'" & strSql & _
'                     " union all select sum(a1p08) as Amount from acc1p0, acc0o0 where a1p23 = a0o01 and a1p01 = '1' and a1p02 = 'Z' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a0o03 = '" & Text2 & "'" & strSql & ") new", adoTaie, adOpenStatic, adLockReadOnly
        '+抓取J公司應付總金額(排除銷貨折讓單未收回者)
'        strExc(0) = "select A1P01,sum(nvl(Amount, 0)) Amount from " & _
'                         "(select A1P01,sum(a1p08) as Amount from acc1p0, acc0o0 where a1p01=a0o07(+) and a1p04 = a0o01(+) and a1p02 = 'B' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a1p15 = '" & Text2 & "'" & strSql & " GROUP BY A1P01 " & _
'             "union all select A1P01,sum(a1p08) as Amount from acc1p0, acc0o0 where a1p01=a0o07(+) and a1p23 = a0o01(+) and a1p02 = 'E' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a0o03 = '" & Text2 & "'" & strSql & " GROUP BY A1P01 " & _
'             "union all select A1P01,sum(a1p08) as Amount from acc1p0, acc0o0 where a1p01=a0o07(+) and a1p23 = a0o01(+) and a1p02 = 'Z' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a0o03 = '" & Text2 & "'" & strSql & " GROUP BY A1P01) " & _
'             "GROUP BY A1P01"

        strExc(0) = "select '1." & A0802Query("1", True) & ":',sum(nvl(Amount, 0)) TAmount from (select sum(a1p08) as Amount from acc1p0, acc0o0 where a1p01(+)=a0o07 and a1p04(+) = a0o01 and a1p01 = '1' and a1p02 = 'B' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a0o03 = '" & Text2 & "'" & strSql & _
                         " union all select sum(a1p08) as Amount from acc1p0, acc0o0 where a1p01(+)=a0o07 and a1p23(+) = a0o01 and a1p01 = '1' and a1p02 = 'E' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a0o03 = '" & Text2 & "'" & strSql & _
                         " union all select sum(a1p08) as Amount from acc1p0, acc0o0 where a1p01(+)=a0o07 and a1p23(+) = a0o01 and a1p01 = '1' and a1p02 = 'Z' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a0o03 = '" & Text2 & "'" & strSql & ") "
        '+抓取J公司應付總金額(排除銷貨折讓單未收回者)
        strExc(0) = strExc(0) & " Union " & _
                        "select 'J." & A0802Query("J", True) & ":',sum(nvl(Amount, 0)) TAmount from (select sum(a1p08) as Amount from acc1p0, acc0o0 where a1p01(+)=a0o07 and a1p04(+) = a0o01 and a1p01 = 'J' and a1p02 = 'B' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a0o03 = '" & Text2 & "' And a1p04 not in (Select a2505 From acc250 Where a2505(+)=a0o01 And a2510 is null And a2519 is null And a0o19='2') " & strSql & _
                        " union all select sum(a1p08) as Amount from acc1p0, acc0o0 where a1p01(+)=a0o07 and a1p23(+) = a0o01 and a1p01 = 'J' and a1p02 = 'E' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a0o03 = '" & Text2 & "' And a1p04 not in (Select a2505 From acc250 Where a2505(+)=a0o01 And a2510 is null And a2519 is null And a0o19='2') " & strSql & _
                        " union all select sum(a1p08) as Amount from acc1p0, acc0o0 where a1p01(+)=a0o07 and a1p23(+) = a0o01 and a1p01 = 'J' and a1p02 = 'Z' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a0o03 = '" & Text2 & "' And a1p04 not in (Select a2505 From acc250 Where a2505(+)=a0o01 And a2510 is null And a2519 is null And a0o19='2') " & strSql & ") "
        'Add by Amy 2020/04/09 增加L公司
        strExc(0) = strExc(0) & " Union " & _
                         "select 'L." & A0802Query("L", True) & ":',sum(nvl(Amount, 0)) TAmount from (select sum(a1p08) as Amount from acc1p0, acc0o0 where a1p01(+)=a0o07 and a1p04(+) = a0o01 and a1p01 = 'L' and a1p02 = 'B' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a0o03 = '" & Text2 & "'" & strSql & _
                         " union all select sum(a1p08) as Amount from acc1p0, acc0o0 where a1p01(+)=a0o07 and a1p23(+) = a0o01 and a1p01 = 'L' and a1p02 = 'E' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a0o03 = '" & Text2 & "'" & strSql & _
                         " union all select sum(a1p08) as Amount from acc1p0, acc0o0 where a1p01(+)=a0o07 and a1p23(+) = a0o01 and a1p01 = 'L' and a1p02 = 'Z' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a0o03 = '" & Text2 & "'" & strSql & ") "

        strExc(0) = "Select * From (" & strExc(0) & ") Where TAmount >0"
        'end 2014/02/10
        
        adoaccsum.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
      '2014/01/21 預設支票抬頭改至選完 應付總金額才帶(combo7_change)
'      If adoaccsum.RecordCount <> 0 Then
'         If IsNull(adoaccsum.Fields(0).Value) Then
'            Text5 = ""
'         Else
'            Text5 = adoaccsum.Fields(0).Value
'            'Add by Morgan 2006/10/27 預設支票抬頭
'            'If Val(Text5) > 0 Then
'            '   SetCheckTitle
'            'End If
'            'end 2006/10/27
'         End If
'      Else
'         Text5 = ""
'      End If
        Combo7.Clear
        Combo7.AddItem MsgText(601)
        If adoaccsum.RecordCount > 0 Then
            adoaccsum.MoveFirst
            Do While Not adoaccsum.EOF
                If Val(adoaccsum.Fields(1)) <> 0 Then
                    'Modify by Amy 2014/02/06
                    'Combo7.AddItem IIf(adoaccsum.Fields(0) = "1", "1.台一:", "J.智權:") & Format(adoaccsum.Fields(1), DDollar2)
                    Combo7.AddItem adoaccsum.Fields(0) & Format(adoaccsum.Fields(1), DDollar2)
                    If adoaccsum.RecordCount = 1 Then
                        '若只有一筆資料預設
                        'Combo7 = IIf(adoaccsum.Fields(0) = "1", "1.台一:", "J.智權:") & Format(adoaccsum.Fields(1), DDollar2)
                        Combo7 = adoaccsum.Fields(0) & Format(adoaccsum.Fields(1), DDollar2)
                    End If
                    'end 2014/02/06
                End If
                adoaccsum.MoveNext
            Loop
        End If
        adoaccsum.Close
      End If
      If Combo7.ListCount = 1 Then
            '無應付款資料
            MsgBox MsgText(142), , MsgText(5)
            Cancel = True
            Text2.SetFocus
      ElseIf Combo7.ListCount = 2 Then
            FormEnabled
            Combo7_Change
            Combo7_Validate (False)
      Else
            Combo7.Enabled = True
      End If
   End If
   'end 2014/01/21
   'Mark by Amy 2014/01/21 往上搬並修改
'   If Val(Text5) = 0 Then
'      MsgBox MsgText(142), , MsgText(5)
'      Cancel = True
'      Text2.SetFocus
'   ElseIf Left(Text2, 1) = "F" Then
   'end 2014/01/21
   'Add by Morgan 2007/6/6 若為外翻時需檢查證明單已收回才可付款
   If Left(Text2, 1) = "F" Then
      If CheckNotRec = True Then
         MsgBox Text3 & "尚有翻譯費證明單未收回！", vbExclamation
      End If
   'end 2007/6/6
   End If
    
   'Modify by Amy 2014/01/21 搬到 選完應付總金額再帶
   'Add by Morgan 2009/8/31
'   If Val(Text5) > 0 And Cancel = False Then
'      strSQL1 = "select a1p05,sum(a1p08) a1p08,a1p14 from acc1p0, acc0o0 where a1p04 = a0o01 and a1p01 = '1' and a1p02 = 'B' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a1p15 = '" & Text2 & "'" & strSql & " group by a1p04,a1p05,a1p14" & _
'               " union all select a1p05,sum(a1p08) a1p08,a1p14 from acc1p0, acc0o0 where a1p23 = a0o01 and a1p01 = '1' and a1p02 = 'E' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a0o03 = '" & Text2 & "'" & strSql & " group by a1p04,a1p05,a1p14" & _
'               " union all select a1p05,sum(a1p08) a1p08,a1p14 from acc1p0, acc0o0 where a1p23 = a0o01 and a1p01 = '1' and a1p02 = 'Z' and (a0o11 is null or a0o11 = 0) and a1p05 in ('2112', '2113') and a0o03 = '" & Text2 & "'" & strSql & " group by a1p04,a1p05,a1p14"
'      intI = 1
'      Set adoTmp = ClsLawReadRstMsg(intI, strSQL1)
'      If intI = 1 Then
'         With adoTmp
'         Do While Not .EOF
'            .MoveNext
'            If .EOF Then
'               bNoCreditSide = False
'            Else
'               bNoCreditSide = True
'            End If
'            .MovePrevious
'
'            Text11 = .Fields("a1p05")
'            Text8 = .Fields("a1p08")
'            Combo5 = .Fields("a1p14")
'            KeyDefine vbKeyInsert
'            .MoveNext
'         Loop
'         End With
'      End If
'   End If
'   bNoCreditSide = False
'   'end 2009/8/31
End Sub
'Add by Morgan 2007/6/6
Private Function CheckNotRec() As Boolean
   strExc(0) = "select 1 from acc250 where a2502='5' and a2503='" & Text2 & "' and a2509 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      CheckNotRec = True
   End If
End Function

'Add by Morgan 2006/10/27 預設支票抬頭
Private Sub SetCheckTitle(Optional bolReset As Boolean = False)
   If bolReset = True Then
      Text4.Text = ""
   End If
   If Text2 <> "" Then
      'Modify by Amy 2014/01/24 +公司別
      strExc(0) = "select A0S18 from acc0O0, acc0S0 where a0o07='" & Mid(Combo7, 1, 1) & "' And (a0o11 is null or a0o11 = 0) and a0o09 is not null and a0s01(+)=a0o09" & _
         " and a0o03 = '" & Text2 & "' and A0S18 is not null order by a0s14,a0s15,a0s11,a0s12"
      intI = 1
      'edit by nickc 2007/02/08 不用 dll 了
      'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Text4.Text = "" & RsTemp(0)
      'Add by Morgan 2006/12/8 扣繳退費抓應付帳款分錄的摘要
      Else
         strExc(0) = "select a1p14 from acc0O0, acc1p0 where a0o07='" & Mid(Combo7, 1, 1) & "' And (a0o11 is null or a0o11 = 0) and a0o19='3' and a1p23(+)=a0o01 and a0o03 = '" & Text2 & "' and a1p14 is not null and a1p05='2112'"
         intI = 1
         'edit by nickc 2007/02/08 不用 dll 了
         'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Text4.Text = "" & RsTemp(0)
         End If
      'end 2006/12/8
      End If
   End If
End Sub
'Add by Morgan 2006/10/30 預設款項類別,領款方式
Private Sub SetPayAffair()
   If Text2 <> "" Then
     'Modify by Amy 2014/01/21 +公司別
      strExc(0) = "select A0O19,A0O09 from acc0O0 where a0o07='" & Mid(Combo7, 1, 1) & "' And a0o11 = " & Val(FCDate(MaskEdBox7.Text)) & " and a0o03 = '" & Text2 & "'"
      intI = 1
      'edit by nickc 2007/02/08 不用 dll 了
      'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Select Case "" & RsTemp(0)
            Case "3" '扣繳
               Combo4.ListIndex = 3
            Case "2" '銷退
               'Modify by Amy 2014/01/21 +公司別
               strExc(0) = "select A0S02,A1P14 from acc0S0, acc1p0 where a0s01='" & RsTemp(1) & "'" & _
                  " And a1p01='" & Mid(Combo7, 1, 1) & "' and a1p04(+)=a0s01 and a1p05(+)='2112' order by a0s14,a0s15,a0s11,a0s12"
               intI = 1
               'edit by nickc 2007/02/08 不用 dll 了
               'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  '暫收
                  If Left("" & RsTemp(0), 1) = "J" Then
                     Combo4.ListIndex = 1
                  '收據款
                  ElseIf Left("" & RsTemp(0), 1) = "E" Then
                     Combo4.ListIndex = 2
                  End If
               
                  If InStr("" & RsTemp(1), "寄出特別") > 0 Then
                     Combo2.ListIndex = 5
                  ElseIf InStr("" & RsTemp(1), "寄出") > 0 Then
                     Combo2.ListIndex = 1
                  ElseIf InStr("" & RsTemp(1), "寄分所") > 0 Then
                     Combo2.ListIndex = 2
                  ElseIf InStr("" & RsTemp(1), "交智權人員") > 0 Then
                     Combo2.ListIndex = 0
                  End If
               End If
         End Select
      End If
   End If
End Sub
Private Sub Text21_GotFocus()
   TextInverse Text21
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text21_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text22_GotFocus()
   TextInverse Text22
End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text22_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text22_Validate(Cancel As Boolean)
   If Text22 <> MsgText(601) Then
      If ExistCheck("staff", "st01", Text22, Label30) = False Then
         Cancel = True
         Exit Sub
      End If
   End If
   'add by sonia 2021/1/29
   If SalesNoCheckAccNo(Text11, Text22) = False Then
   End If
   'end 2021/1/29
End Sub

Private Sub Text23_GotFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
OpenIme
End Sub

'Add by Morgan 2006/12/5
'Modify by Amy 改Form2.0 原:Integer
Private Sub Text23_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyDefine Val(KeyCode)
End Sub

Private Sub Text23_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
CloseIme
End Sub

Private Sub Text4_GotFocus()
   StatusView MsgText(65) & "100"
   TextInverse Text4
   'edit by nickc 2007/06/11  切換輸入法改用API
   OpenIme
   
End Sub

'Modify by Amy 2021/08/20 改Form2.0 原:Integer
Private Sub Text4_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyDefine Val(KeyCode)
End Sub

Private Sub Text4_LostFocus()
   StatusView MsgText(601)
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
   If CheckLen(Label8, Text4, 100) = MsgText(603) Then
      Cancel = True
      Text4.SetFocus
      Exit Sub
   End If
   'edit by nickc 2007/06/11  切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub

Private Sub Text8_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text9_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'*************************************************
'  重新整裡 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   'Modify by Amy 2014/01/21 改公司別 原:'1'
   adoadodc1.Open "select * from acc1p0, acc010 where a1p05 = a0101 and a1p01 = '" & Mid(Combo7, 1, 1) & "' and a1p02 = 'C'" & _
      " and a1p04 = '" & Text17 & "' order by a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.Requery
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  儲存資料表(國內付款資料(交易檔))
'
'*************************************************
Private Sub Acc1p0Save()
Dim strPayTime As String
Dim strType As String
Dim strAccCode As String 'Added by Morgan 2011/11/15 乙存會計科目
Dim strSql As String 'Add by Amy 2020/07/01

On Error GoTo Checking
   If Text11 = MsgText(601) Then
      MsgBox MsgText(10) & Label19, , MsgText(5)
      strControlButton = MsgText(602)
      Text11.SetFocus
      Exit Sub
   Else
      If ExistCheck("acc010", "a0101", Text11, Label19) = False Then
         strControlButton = MsgText(602)
         Text11.SetFocus
         Exit Sub
      End If
      If Text10 <> MsgText(601) Then
         If ExistCheck("acc0g0", "a0g01", Text10, Label18) = False Then
            strControlButton = MsgText(602)
            Text10.SetFocus
            Exit Sub
         End If
      End If
      'Modify by Amy 2014/01/21 原text5 改combo7顯示
'      If Text5 = MsgText(601) Then
'         strControlButton = MsgText(602)
'         MsgBox MsgText(136), , MsgText(5)
'         Text2.SetFocus
'         Exit Sub
'      End If
      If Combo7.ListCount > 0 And Combo7 = "" Then
        MsgBox Label9 & "必要欄位,請選取...", , MsgText(5)
        strControlButton = MsgText(602)
        Combo7.SetFocus
        Exit Sub
      End If
      'end 2014/01/21
      If CheckDept(Text11, Text18) = False Then
         MsgBox MsgText(103), , MsgText(5)
         strControlButton = MsgText(602)
         Text18.SetFocus
         Exit Sub
      End If
      If MaskEdBox5.Text <> MsgText(601) And MaskEdBox5.Text <> MsgText(29) Then
         If DateCheck(MaskEdBox5.Text) = MsgText(603) Then
            MsgBox Label14 & MsgText(63), , MsgText(5)
            strControlButton = MsgText(602)
            MaskEdBox5.SetFocus
            Exit Sub
         End If
      End If
      If MaskEdBox6.Text <> MsgText(601) And MaskEdBox6.Text <> MsgText(29) Then
         If DateCheck(MaskEdBox6.Text) = MsgText(603) Then
            MsgBox Label15 & MsgText(63), , MsgText(5)
            strControlButton = MsgText(602)
            MaskEdBox6.SetFocus
            Exit Sub
         End If
      End If
      If Text19 <> MsgText(601) Then
         Text19 = CaseNoZero(Text19)
         adocase.CursorLocation = adUseClient
         adocase.Open "select pa01 as SystemNo from patent where pa01 = '" & Mid(Text19, 1, Len(Text19) - 9) & "' and pa02 = '" & Mid(Text19, Len(Text19) - 8, 6) & "' and pa03 = '" & Mid(Text19, Len(Text19) - 2, 1) & "' and pa04 = '" & Mid(Text19, Len(Text19) - 1, 2) & "' union " & _
                        "select tm01 as SystemNo from trademark where tm01 = '" & Mid(Text19, 1, Len(Text19) - 9) & "' and tm02 = '" & Mid(Text19, Len(Text19) - 8, 6) & "' and tm03 = '" & Mid(Text19, Len(Text19) - 2, 1) & "' and tm04 = '" & Mid(Text19, Len(Text19) - 1, 2) & "' union " & _
                        "select lc01 as SystemNo from lawcase where lc01 = '" & Mid(Text19, 1, Len(Text19) - 9) & "' and lc02 = '" & Mid(Text19, Len(Text19) - 8, 6) & "' and lc03 = '" & Mid(Text19, Len(Text19) - 2, 1) & "' and lc04 = '" & Mid(Text19, Len(Text19) - 1, 2) & "' union " & _
                        "select hc01 as SystemNo from hirecase where hc01 = '" & Mid(Text19, 1, Len(Text19) - 9) & "' and hc02 = '" & Mid(Text19, Len(Text19) - 8, 6) & "' and hc03 = '" & Mid(Text19, Len(Text19) - 2, 1) & "' and hc04 = '" & Mid(Text19, Len(Text19) - 1, 2) & "' union " & _
                        "select sp01 as SystemNo from servicepractice where sp01 = '" & Mid(Text19, 1, Len(Text19) - 9) & "' and sp02 = '" & Mid(Text19, Len(Text19) - 8, 6) & "' and sp03 = '" & Mid(Text19, Len(Text19) - 2, 1) & "' and sp04 = '" & Mid(Text19, Len(Text19) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adocase.RecordCount = 0 Then
            MsgBox MsgText(28) & Label28, , MsgText(5)
            strControlButton = MsgText(602)
            adocase.Close
            Exit Sub
         End If
         adocase.Close
      End If
      If Text22 <> MsgText(601) Then
         If ExistCheck("staff", "st01", Text22, Label30) = False Then
            strControlButton = MsgText(602)
            Exit Sub
         End If
      End If
      Select Case Text11
         Case "2111", "110201", "110214"
            If Combo3 = MsgText(601) Then
               MsgBox MsgBox(63) & Label11, , MsgText(5)
               strControlButton = MsgText(602)
               Combo3.SetFocus
               Exit Sub
            End If
            If Text10 = MsgText(601) Then
               MsgBox MsgText(63) & Label18, , MsgText(5)
               strControlButton = MsgText(602)
               Text10.SetFocus
               Exit Sub
            End If
            If Text9 = MsgText(601) Then
               MsgBox MsgText(63) & Label13, , MsgText(5)
               strControlButton = MsgText(602)
               Text9.SetFocus
               Exit Sub
            End If
            If MaskEdBox5.Text = MsgText(601) Or MaskEdBox5.Text = MsgText(29) Then
               MsgBox MsgText(63) & Label14, , MsgText(5)
               strControlButton = MsgText(602)
               MaskEdBox5.SetFocus
               Exit Sub
            End If
            If MaskEdBox6.Text = MsgText(601) Or MaskEdBox6.Text = MsgText(29) Then
               MsgBox MsgText(63) & Label15, , MsgText(5)
               strControlButton = MsgText(602)
               MaskEdBox6.SetFocus
               Exit Sub
            End If
            If Combo1 = MsgText(601) Then
               MsgBox MsgText(63) & Label20, , MsgText(5)
               strControlButton = MsgText(602)
               Combo1.SetFocus
               Exit Sub
            End If
            If Combo2 = MsgText(601) Then
               MsgBox MsgText(63) & Label21, , MsgText(5)
               strControlButton = MsgText(602)
               Combo2.SetFocus
               Exit Sub
            End If
            If Combo4 = MsgText(601) Then
               MsgBox MsgText(63) & Label17, , MsgText(5)
               strControlButton = MsgText(602)
               Combo4.SetFocus
               Exit Sub
            End If
            
            'Add by Morgan 2010/6/2 檢查票據是否存在
            If adoquery.State = adStateOpen Then adoquery.Close
            'Modify by Amy 2014/01/21 +公司別
            'Modify  by Amy 2020/07/01 +a0e07 因改為key
            strSql = "select * from acc0e0 where a0e01 = '" & Text10 & "' and a0e02 = '" & Text9 & "' And a0e23='" & Mid(Combo7, 1, 1) & "' And a0e07='" & Combo3 & "' "
            adoquery.CursorLocation = adUseClient
            adoquery.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
            If adoquery.RecordCount <> 0 Then
               If adoquery.Fields("a0e03").Value <> Text17 Then
                  MsgBox MsgText(196), , MsgText(5)
                  strControlButton = MsgText(602)
                  Text9.SetFocus
                  adoquery.Close
                  Exit Sub
               End If
            End If
            adoquery.Close
            'end 2010/6/2
      End Select
   End If
   
   'add by sonia 2015/12/30 檢查民國105年起法務收入科目不可使用
   intI = PUB_AccNoEnable(Text11, Val(FCDate(MaskEdBox7.Text)))
   If intI <> 0 Then
      strControlButton = MsgText(602)
      Text11.SetFocus
      Exit Sub
   End If
   'end 2015/12/30
   'Add by Morgan 2007/10/2 檢查科目部門&智權人員是否正確
   intI = PUB_AccNoGood(Text11, Text18, Text22)
   If intI <> 0 Then
      strControlButton = MsgText(602)
      If intI = 1 Then
         Text11.SetFocus
      ElseIf intI = 2 Then
         Text18.SetFocus
      ElseIf intI = 3 Then
         Text22.SetFocus
      End If
      Exit Sub
   End If
   'end 2007/10/2
   'Add by Amy 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me) = False Then
        strControlButton = MsgText(602)
        Exit Sub
   End If
   
   If Adodc1.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc1.Recordset.Fields("a1p22").Value) = False Then
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select ax210 from acc021 where ax201 = '" & Adodc1.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc1.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(155), , MsgText(5)
            strControlButton = MsgText(602)
            Text11.SetFocus
            adoquery.Close
            Exit Sub
         End If
         adoquery.Close
      End If
   End If
   'Modify by Amy 2014/02/10 以會計科目抓Acc0H0開票帳號及銀行
    StrA0H01 = ""
    StrA0H02 = GetBankData(strAccCode, StrA0H01)
   If StrA0H02 <> MsgText(601) And StrA0H01 <> MsgText(601) Then
        Combo3 = StrA0H02
        Text10 = StrA0H01
    End If
   'end 2014/02/10
   adoacc1p0.Close
   adoacc1p0.CursorLocation = adUseClient
   'Modify by Amy 2014/01/21 改公司別 原:'1'
   adoacc1p0.Open "select * from acc1p0 where a1p01 = '" & Mid(Combo7, 1, 1) & "' and a1p02 = 'C' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text17 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc1p0.RecordCount = 0 Then
      adoacc1p0.AddNew
      adoacc1p0.Fields("a1p01").Value = Mid(Combo7, 1, 1)  '2014/01/21 原:"1"
      adoacc1p0.Fields("a1p02").Value = "C"
      adoacc1p0.Fields("a1p03").Value = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '" & Mid(Combo7, 1, 1) & "' and a1p02 = 'C' and a1p04 = '" & Text17 & "'", 3) '2014/01/21 原:'1'
      adoacc1p0.Fields("a1p04").Value = Text17
   End If
   adoacc1p0.Fields("a1p05").Value = Text11
   adoacc1p0.Fields("a1p06").Value = MsgText(55)
   If Combo3 <> MsgText(601) Then
      adoacc1p0.Fields("a1p11").Value = Combo3
   Else
      adoacc1p0.Fields("a1p11").Value = Null
   End If
   If Text8 <> MsgText(601) Then
      adoacc1p0.Fields("a1p07").Value = Val(Format(Text8, DAmount))
   Else
      adoacc1p0.Fields("a1p07").Value = 0
   End If
   If Text13 <> MsgText(601) Then
      adoacc1p0.Fields("a1p08").Value = Val(Format(Text13, DAmount))
   Else
      adoacc1p0.Fields("a1p08").Value = 0
   End If
   If MaskEdBox7.Text <> MsgText(601) And MaskEdBox7.Text <> MsgText(29) Then
      adoacc1p0.Fields("a1p18").Value = Val(FCDate(MaskEdBox7.Text))
   Else
      adoacc1p0.Fields("a1p18").Value = Null
   End If
   If MaskEdBox6.Text <> MsgText(601) And MaskEdBox6.Text <> MsgText(29) Then
      adoacc1p0.Fields("a1p12").Value = Val(FCDate(MaskEdBox6.Text))
   Else
      adoacc1p0.Fields("a1p12").Value = Null
   End If
   If Combo1 <> MsgText(601) Then
      adoacc1p0.Fields("a1p13").Value = Mid(Combo1, 1, 1)
   Else
      adoacc1p0.Fields("a1p13").Value = Null
   End If
   If Combo2 <> MsgText(601) Then
      adoacc1p0.Fields("a1p24").Value = Mid(Combo2, 1, 1)
   Else
      adoacc1p0.Fields("a1p24").Value = Null
   End If
   If Text7 <> MsgText(601) Then
      adoacc1p0.Fields("a1p25").Value = Val(Text7)
   Else
      adoacc1p0.Fields("a1p25").Value = 0
   End If
   If Combo5 <> MsgText(601) Then
      adoacc1p0.Fields("a1p14").Value = Combo5
      Combo5.AddItem Combo5
   Else
      adoacc1p0.Fields("a1p14").Value = Null
   End If
   If Text2 <> MsgText(601) Then
      adoacc1p0.Fields("a1p15").Value = Text2
   End If
   'modify by sonia 2021/1/29 加傳本所案號以判別FCP,FCT英日文組
   'If AccNoToSalesNo(Text11) = "" Then
   If AccNoToSalesNo(Text11, Text19) = "" Then
      adoacc1p0.Fields("a1p16").Value = Null
   Else
      'modify by sonia 2021/1/29 加傳本所案號以判別FCP,FCT英日文組
      'adoacc1p0.Fields("a1p16").Value = AccNoToSalesNo(Text11)
      adoacc1p0.Fields("a1p16").Value = AccNoToSalesNo(Text11, Text19)
   End If
   If Text10 <> MsgText(601) Then
      adoacc1p0.Fields("a1p10").Value = Text10
   Else
      adoacc1p0.Fields("a1p10").Value = Null
   End If
   If Text9 <> MsgText(601) Then
      adoacc1p0.Fields("a1p09").Value = Text9
   Else
      adoacc1p0.Fields("a1p09").Value = Null
   End If
   If Combo4 <> MsgText(601) Then
      adoacc1p0.Fields("a1p26").Value = Mid(Combo4, 1, 1)
   Else
      adoacc1p0.Fields("a1p26").Value = Null
   End If
   If Text18 <> MsgText(601) Then
      adoacc1p0.Fields("a1p06").Value = Text18
   Else
      adoacc1p0.Fields("a1p06").Value = MsgText(55)
   End If
   If Text19 <> MsgText(601) Then
      adoacc1p0.Fields("a1p17").Value = Text19
   Else
      adoacc1p0.Fields("a1p17").Value = Null
   End If
   If Text21 <> MsgText(601) Then
      adoacc1p0.Fields("a1p30").Value = Text21
   Else
      adoacc1p0.Fields("a1p30").Value = Null
   End If
   If Text22 <> MsgText(601) Then
      adoacc1p0.Fields("a1p16").Value = Text22
   Else
      adoacc1p0.Fields("a1p16").Value = Null
   End If
   If IsNull(adoacc1p0.Fields("a1p22").Value) = False Then
      adoacc1p0.Fields("a1p27").Value = MsgText(602)
   End If
   adoacc1p0.UpdateBatch
   strSerialNo = MsgText(601)
   Select Case Text1
      Case "1"
         strType = "2"
      Case "2"
         strType = "1"
      Case "3"
         strType = "3"
   End Select
   Select Case Text11
      Case "2111", "110201", "110214"
         adoquery.CursorLocation = adUseClient
         'Modify by Amy 2014/01/21 +公司別
         'Modify  by Amy 2020/07/01 +a0e07 因改為key
         strSql = "select * from acc0e0 where a0e01 = '" & Text10 & "' and a0e02 = '" & Text9 & "' And a0e23='" & Mid(Combo7, 1, 1) & "' And a0e07='" & Combo3 & "' "
         adoquery.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            If adoquery.Fields("a0e03").Value = Text17 Then
               'Modify by Amy 2014/01/21 +公司別
               'Modify by Amy 2020/07/01 拿掉 a0e07 = '" & Combo3 & "'
               'adoTaie.Execute "update acc0e0 set a0e07 = '" & Combo3 & "', a0e13 = " & Val(FCDate(MaskEdBox5.Text)) & ", a0e10 = " & Val(FCDate(MaskEdBox6.Text)) & ", a0e08 = '" & Mid(Combo1, 1, 1) & "', a0e12 = '" & Text4 & "', a0e11 = " & Val(Text13) & " where a0e01 = '" & Text10 & "' and a0e02 = '" & Text9 & "' And a0e23='" & Mid(Combo7, 1, 1) & "' "
               adoTaie.Execute "update acc0e0 set  a0e13 = " & Val(FCDate(MaskEdBox5.Text)) & ", a0e10 = " & Val(FCDate(MaskEdBox6.Text)) & ", a0e08 = '" & Mid(Combo1, 1, 1) & "', a0e12 = '" & Text4 & "', a0e11 = " & Val(Text13) & " where a0e01 = '" & Text10 & "' and a0e02 = '" & Text9 & "' And a0e23='" & Mid(Combo7, 1, 1) & "' And a0e07='" & Combo3 & "' "
            Else
               MsgBox MsgText(196), , MsgText(5)
               adoquery.Close
               Text9.SetFocus
               Exit Sub
            End If
         Else
            'Add by Amy 2020/07/01 若修改票據號碼按 Insert,原資料不會被刪除
            If Text9 & Text10 & Combo3 & Mid(Combo7, 1, 1) <> Text9.Tag & Text10.Tag & Combo3.Tag & Combo7.Tag Then
                adoTaie.Execute "delete from acc0e0 where a0e01 = '" & Text10.Tag & "' and a0e02 = '" & Text9.Tag & "' And a0e23='" & Combo7.Tag & "' And a0e07='" & Combo3.Tag & "' "
            End If
            'Modify by Amy 2014/01/21 +公司別 Insert語法 原:null
            'Modify  by Amy 2020/07/01 +a0e07 因改為key
            adoTaie.Execute "delete from acc0e0 where a0e01 = '" & Text10 & "' and a0e02 = '" & Text9 & "' And a0e23='" & Mid(Combo7, 1, 1) & "' And a0e07='" & Combo3 & "' "
            adoTaie.Execute "insert into acc0e0 values ('" & Text9 & "', '" & Text10 & "', '" & Text17 & "', 'P', '" & strType & "', '" & Text2 & "', '" & Combo3 & "', '" & Mid(Combo1, 1, 1) & "', " & _
                            "" & Val(FCDate(MaskEdBox6.Text)) & ", " & Val(Text13) & ", '" & Text4 & "', " & Val(FCDate(MaskEdBox5.Text)) & ", 0, 0, 0, null, 0, null, null, null, 0, 0, 0, 0, '" & Mid(Combo7, 1, 1) & "', null, null, null, null, 0, null, " & Val(Text7) & ", null, null, '" & strUserNum & "', " & Val(strSrvDate(2)) & ", " & ServerTime & ", null, 0, 0, 0, 0, 0, null, 0, null, null)"
         End If
         adoquery.Close
      'Add by Morgan 2006/6/28 當借方科目為 "2112", "2113" 時,若廠商的付款方式設定為1或2時自動新增貸方科目110202,摘要帶廠商名稱
      Case "2112", "2113"
         If bNoCreditSide = False Then 'Add by Morgan 2009/9/1 借方批次新增時最後一筆才新增貸方
         
            adoquery.CursorLocation = adUseClient
            'Modified by Morgan 2011/11/15 +a0i20
            adoquery.Open "select a0i12,a0i20 from acc0i0 where a0i01 = '" & Text2 & "' and a0i12 in ('1','2')", adoTaie, adOpenForwardOnly, adLockReadOnly
            If adoquery.RecordCount <> 0 Then
               adoacc1p0.MoveFirst
               'Modified by Morgan 2011/11/15 +華銀
               'adoacc1p0.Find "a1p05='110202'"
               'Modify by Amy 2014/01/29 +J公司
'               If adoquery.Fields("a0i12") = "2" Or Combo6.ListIndex = 1 Then
'                  strAccCode = "110202"
'               Else
'                  strAccCode = "110207"
'               End If
               'Modify by Amy 2021/03/18 +L公司,會計科目固定帶110502
               If Mid(Combo7, 1, 1) = "L" Then
                   strAccCode = "110502"
               'Add by Amy 2023/06/28 J公司,會計科目固定帶 110303-瑞婷
               ElseIf Mid(Combo7, 1, 1) = "J" Then
                   strAccCode = "110303"
               '一信直存 or 主要付款銀行為瑞興
               ElseIf adoquery.Fields("a0i12") = "2" Or Combo6.ListIndex = 1 Then
                  'Mark by Amy 2023/06/28
'                  If Mid(Combo7, 1, 1) = "J" Then
'                     strAccCode = "110303"
'                  Else
                     'modify by sonia 2020/5/12 110202->110602
                     strAccCode = "110602"
'                  End If
               Else
                  'Mark by Amy 2023/06/28
'                  If Mid(Combo7, 1, 1) = "J" Then
'                     strAccCode = "110304"
'                  Else
                     strAccCode = "110207"
'                  End If
               End If
               'end 2014/01/29
               'modify by sonia 2021/6/10 修改資料時應重新讀資料,否則會又新增一個項次
               'adoacc1p0.Find "a1p05='" & strAccCode & "'"
               adoacc1p0.Close
               adoacc1p0.CursorLocation = adUseClient
               adoacc1p0.Open "select * from acc1p0 where a1p01 = '" & Mid(Combo7, 1, 1) & "' and a1p02 = 'C' and a1p05='" & strAccCode & "' and a1p04 = '" & Text17 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
               'end 2021/6/10
               'end 2011/11/15
               'Modify by Amy 2014/02/10 以會計科目抓Acc0H0開票帳號及銀行
                StrA0H01 = ""
                StrA0H02 = GetBankData(strAccCode, StrA0H01)
                If StrA0H02 <> MsgText(601) And StrA0H01 <> MsgText(601) Then
                    Combo3 = StrA0H02
                    Text10 = StrA0H01
                End If
                'end 2014/02/10
               If adoacc1p0.EOF Then
                  adoacc1p0.AddNew
                  adoacc1p0.Fields("a1p01") = Mid(Combo7, 1, 1) '2014/01/21 原:"1"
                  adoacc1p0.Fields("a1p02") = "C"
                  adoacc1p0.Fields("a1p03").Value = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '" & Mid(Combo7, 1, 1) & "' and a1p02 = 'C' and a1p04 = '" & Text17 & "'", 3) '2014/01/21 原:'1'
                  adoacc1p0.Fields("a1p04").Value = Text17
                  'Modified by Morgan 2011/11/15 +華銀
                  'adoacc1p0.Fields("a1p05").Value = "110202"
                  adoacc1p0.Fields("a1p05").Value = strAccCode
                  'end 2011/11/15
                  adoacc1p0.Fields("a1p06").Value = MsgText(55)
                  adoacc1p0.Fields("a1p07").Value = 0
                  'Modify by Amy 2014/01/21 原Text5 改Combo7顯示
                  'adoacc1p0.Fields("a1p08").Value = Val(Format(Text5, DAmount))
                  adoacc1p0.Fields("a1p08").Value = Val(Format(Mid(Combo7, InStr(1, Combo7, ":") + 1), DAmount))
                  'end 2014/01/21
                  If Text10 <> MsgText(601) Then
                     adoacc1p0.Fields("a1p10").Value = Text10
                  Else
                     adoacc1p0.Fields("a1p10").Value = Null
                  End If
                  If Combo3 <> MsgText(601) Then
                     adoacc1p0.Fields("a1p11").Value = Combo3
                  Else
                     adoacc1p0.Fields("a1p11").Value = Null
                  End If
                  If Combo1 <> MsgText(601) Then
                     adoacc1p0.Fields("a1p13").Value = Mid(Combo1, 1, 1)
                  Else
                     adoacc1p0.Fields("a1p13").Value = Null
                  End If
                  If MaskEdBox7.Text <> MsgText(601) And MaskEdBox7.Text <> MsgText(29) Then
                     adoacc1p0.Fields("a1p18").Value = Val(FCDate(MaskEdBox7.Text))
                  Else
                     adoacc1p0.Fields("a1p18").Value = Null
                  End If
                  If Text3 <> MsgText(601) Then
                     adoacc1p0.Fields("a1p14").Value = Text3
                  End If
                  If Text2 <> MsgText(601) Then
                     adoacc1p0.Fields("a1p15").Value = Text2
                  End If
                  adoacc1p0.UpdateBatch
               End If
            End If
            adoquery.Close
         End If
   End Select
   AdodcRefresh
   Adodc1.Recordset.MoveLast
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料(國內付款資料(交易檔))
'
'*************************************************
Private Sub AdodcShow()
   Text11 = Adodc1.Recordset.Fields("a1p05").Value
   If IsNull(Adodc1.Recordset.Fields("a1p11").Value) Then
      Combo3 = MsgText(601)
   Else
      Combo3 = Adodc1.Recordset.Fields("a1p11").Value
   End If
   Combo3.Tag = Combo3 'Add by Amy 2020/07/01 開票帳號
   If IsNull(Adodc1.Recordset.Fields("a1p10").Value) Then
      Text10 = MsgText(601)
   Else
      Text10 = Adodc1.Recordset.Fields("a1p10").Value
   End If
   Text10.Tag = Text10 'Add by Amy 2020/07/01
   'Add by Amy 2014/01/21
   If strSaveConfirm <> MsgText(601) Then
        'Modify by Amy 2014/02/10 原判斷Combo3及Text10
        Combo3.Enabled = True
        Text10.Enabled = True
        StrA0H01 = ""
        If GetBankData(Text11, StrA0H01) <> "" Then
            Combo3.Enabled = False
            Text10.Enabled = False
        End If
        'end 2014/02/10
   End If
   'end 2014/01/21
   If IsNull(Adodc1.Recordset.Fields("a1p09").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = Adodc1.Recordset.Fields("a1p09").Value
   End If
   Text9.Tag = Text9 'Add by Amy 2020/07/01
   If IsNull(Adodc1.Recordset.Fields("a1p07").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = Adodc1.Recordset.Fields("a1p07").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p08").Value) Then
      Text13 = MsgText(601)
   Else
      Text13 = Adodc1.Recordset.Fields("a1p08").Value
   End If
   MaskEdBox5.Mask = MsgText(601)
   If IsNull(Adodc1.Recordset.Fields("a1p18").Value) Then
      MaskEdBox5.Text = MsgText(601)
   Else
      MaskEdBox5.Text = CFDate(Adodc1.Recordset.Fields("a1p18").Value)
   End If
   MaskEdBox5.Mask = DFormat
   MaskEdBox6.Mask = MsgText(601)
   If IsNull(Adodc1.Recordset.Fields("a1p12").Value) Then
      MaskEdBox6.Text = MsgText(601)
   Else
      MaskEdBox6.Text = CFDate(Adodc1.Recordset.Fields("a1p12").Value)
   End If
   MaskEdBox6.Mask = DFormat
   If IsNull(Adodc1.Recordset.Fields("a1p13").Value) Then
      Combo1 = MsgText(601)
   Else
      Combo1 = Combo1.List(Val(Adodc1.Recordset.Fields("a1p13").Value) - 1)
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p24").Value) Then
      Combo2 = MsgText(601)
   Else
      Combo2 = Combo2.List(Val(Adodc1.Recordset.Fields("a1p24").Value) - 1)
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p25").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = Adodc1.Recordset.Fields("a1p25").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p26").Value) Then
      Combo4 = MsgText(601)
   Else
      Combo4 = Combo4.List(Val(Adodc1.Recordset.Fields("a1p26").Value) - 1)
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p14").Value) Then
      Combo5 = MsgText(601)
   Else
      Combo5 = Adodc1.Recordset.Fields("a1p14").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p06").Value) Then
      Text18 = MsgText(601)
   Else
      Text18 = Adodc1.Recordset.Fields("a1p06").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p17").Value) Then
      Text19 = MsgText(601)
   Else
      Text19 = Adodc1.Recordset.Fields("a1p17").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p30").Value) Then
      Text21 = MsgText(601)
   Else
      Text21 = Adodc1.Recordset.Fields("a1p30").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p16").Value) Then
      Text22 = MsgText(601)
   Else
      Text22 = Adodc1.Recordset.Fields("a1p16").Value
   End If
End Sub

'*************************************************
'  清除顯示資料
'
'*************************************************
Public Sub AdodcClear()
  
  Text11 = ""
  Text10 = "011010075"
  Text10.Tag = Text10 'Add by Amy 2020/07/02
  '2010/6/21 MODIFY BY SONIA
  'Combo3 = "0149950"
  'modify by sonia 2020/6/19
  'Modify by Amy 2020/07/02 Mark 0149951,+.tag上次Insert資料
  'Combo3 = "0149951"
  Combo3 = "1756650"
  Combo3.Tag = Combo3
  Text9 = ""
  Text9.Tag = Text9
  'end 2020/07/02
  Text8 = ""
  Text13 = ""
  MaskEdBox5.Mask = ""
  MaskEdBox5.Text = CFDate(Val(strSrvDate(2)))
  MaskEdBox5.Mask = DFormat
  MaskEdBox6.Mask = ""
  MaskEdBox6.Text = ""
  MaskEdBox6.Mask = DFormat
  Combo1 = ComboItem(11)
  Combo2 = ""
  Text7 = ""
  Combo4 = ""
  Combo5 = ""
  Text18 = ""
  Text19 = ""
  Text21 = ""
  Text22 = ""
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyInsert
         If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
            Exit Sub
         End If
         'Add by Morgan 2006/10/31
         If Text11 = "2111" And Text23 = "" And Left(Combo4, 1) = 7 Then
            MsgBox "款項類別選【其他】時付款備註不可空白！(要印在付款通知單)"
            Text23.SetFocus
            Exit Sub
         End If
         'end 2006/10/31
         Frmacc1180_Save
         If strControlButton <> MsgText(602) Then
            Acc1p0Save
         End If
         If strControlButton <> MsgText(602) Then
            SumShow
            AdodcClear
            Text11.SetFocus
         End If
         'Modify by Amy 2014/01/21 新增時acc1p0 有資料不可改往來對象及應付總金額
         If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
            If CheckAcc1p0(Mid(Combo7, 1, 1), Text17) = True Then
                Text2.Enabled = False
                Combo7.Enabled = False
            Else
                Text2.Enabled = True
                Combo7.Enabled = True
            End If
         End If
         'end 2014/01/21
         strControlButton = MsgText(601)
   End Select
   KeyEnter KeyCode
End Sub

'*************************************************
'  計算並顯示合計
'
'*************************************************
Public Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   'Modify by Amy 2014/01/21 改公司別 原:'1'
   adoaccsum.Open "select sum(a1p07), sum(a1p08), count(*) from acc1p0 where a1p01 = '" & Mid(Combo7, 1, 1) & "' and a1p02 = 'C'" & _
      " and a1p04 = '" & Text17 & "' order by a1p05 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text6 = MsgText(601)
      Else
         Text6 = adoaccsum.Fields(0).Value
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text15 = MsgText(601)
      Else
         Text15 = adoaccsum.Fields(1).Value
      End If
      If IsNull(adoaccsum.Fields(2).Value) Then
         Text20 = MsgText(601)
      Else
         Text20 = adoaccsum.Fields(2).Value
      End If
   Else
      Text6 = MsgText(601)
      Text15 = MsgText(601)
      Text20 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  計算廠商或員工之應付總金額
'
'*************************************************
Private Sub Calculate1()
Dim strSql As String
   
   '93.12.8 add by sonia
   strSql = ""
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0o05 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0o05 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If MaskEdBox3.Text <> MsgText(601) And MaskEdBox3.Text <> MsgText(29) Then
      strSql = strSql & " and a0o06 >= " & Val(FCDate(MaskEdBox3.Text)) & ""
   End If
   If MaskEdBox4.Text <> MsgText(601) And MaskEdBox4.Text <> MsgText(29) Then
      strSql = strSql & " and a0o06 <= " & Val(FCDate(MaskEdBox4.Text)) & ""
   End If
   '93.12.8 end
   adoacctotal.CursorLocation = adUseClient
   '93.12.8 modify by sonia
   'adoacctotal.Open "select sum(a1p08) from acc1p0, acc0o0 where acc1p0.a1p03 = acc0o0.a0o01 and a0o05 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0o05 <= " & Val(FCDate(MaskEdBox2.Text)) & " and a0o03 = '" & Text2 & "' and a1p01 = '1' and a1p02 = 'B' and a0o11 is null", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2014/01/21 改公司別 原:'1' 並拿掉多的and 原:and " & strSql & "
   adoacctotal.Open "select sum(a1p08) from acc1p0, acc0o0 where acc1p0.a1p03 = acc0o0.a0o01 " & strSql & " and a0o03 = '" & Text2 & "' and a1p01 = '" & Mid(Combo7, 1, 1) & "' and a1p02 = 'B' and a0o11 is null", adoTaie, adOpenStatic, adLockReadOnly
   '93.12.8 end
   If adoacctotal.RecordCount <> 0 Then
      'Modify by Amy 2014/01/21 目前此function 未使用且text5拿掉改combo7 故mark
      If IsNull(adoacctotal.Fields(0).Value) Then
         'Text5 = MsgText(601)
      Else
         'Text5 = adoacctotal.Fields(0).Value
      End If
   Else
      'Text5 = MsgText(601)
   End If
   adoacctotal.Close
End Sub

'*************************************************
'  計算客戶之應付總金額
'
'*************************************************
Private Sub Calculate2()
Dim strSql As String
   
   '93.12.8 add by sonia
   strSql = ""
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0o05 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0o05 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If MaskEdBox3.Text <> MsgText(601) And MaskEdBox3.Text <> MsgText(29) Then
      strSql = strSql & " and a0o06 >= " & Val(FCDate(MaskEdBox3.Text)) & ""
   End If
   If MaskEdBox4.Text <> MsgText(601) And MaskEdBox4.Text <> MsgText(29) Then
      strSql = strSql & " and a0o06 <= " & Val(FCDate(MaskEdBox4.Text)) & ""
   End If
   '93.12.8 end
   adoacctotal.CursorLocation = adUseClient
   '93.12.8 modify by sonia
   'adoacctotal.Open "select sum(a1p08) from acc1p0, acc0o0 where acc1p0.a1p03 = acc0o0.a0o01 and a0o06 >= " & Val(FCDate(MaskEdBox3.Text)) & " and a0o06 <= " & Val(FCDate(MaskEdBox4.Text)) & " and a0o03 = '" & Text2 & "' and a1p01 = '1' and a1p02 = 'B' and a0o11 is null ", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2014/01/21 改公司別 原:'1'  拿掉多的and 原:and " & strSql & "
   adoacctotal.Open "select sum(a1p08) from acc1p0, acc0o0 where acc1p0.a1p03 = acc0o0.a0o01 " & strSql & " and a0o03 = '" & Text2 & "' and a1p01 = '" & Mid(Combo7, 1, 1) & "' and a1p02 = 'B' and a0o11 is null ", adoTaie, adOpenStatic, adLockReadOnly
   '93.12.8 end
   'Modify by Amy 2014/01/21 目前此function 未使用且text5拿掉改combo7 故mark
   If adoacctotal.RecordCount <> 0 Then
      If IsNull(adoacctotal.Fields(0).Value) Then
         'Text5 = MsgText(601)
      Else
         'Text5 = adoacctotal.Fields(0).Value
      End If
   Else
      'Text5 = MsgText(601)
   End If
   adoacctotal.Close
   adoacctotal.CursorLocation = adUseClient
   '93.12.8 modify by sonia
   'adoacctotal.Open "select sum(a1p08) from acc1p0, acc0o0 where a1p23 = a0o01 and a0o06 >= " & Val(FCDate(MaskEdBox3.Text)) & " and a0o06 <= " & Val(FCDate(MaskEdBox4.Text)) & " and a0o03 = '" & Text2 & "' and a1p01 = '1' and a1p02 = 'E' and a0o11 is null ", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2014/01/21 改公司別 原:'1'
   adoacctotal.Open "select sum(a1p08) from acc1p0, acc0o0 where a1p23 = a0o01 and " & strSql & " and a0o03 = '" & Text2 & "' and a1p01 = '" & Mid(Combo7, 1, 1) & "' and a1p02 = 'E' and a0o11 is null ", adoTaie, adOpenStatic, adLockReadOnly
   '93.12.8 end
   If adoacctotal.RecordCount <> 0 Then
      If IsNull(adoacctotal.Fields(0).Value) = False Then
         'Text5 = Val(Text5) + Val(adoacctotal.Fields(0).Value)
      End If
   End If
   'end 2014/01/21
   adoacctotal.Close
End Sub

'*************************************************
'  刪除資料表(國內付款資料(交易檔))
'
'*************************************************
Private Sub AdodcDelete()
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   'Modify by Amy 2014/01/21 公司別
   adoTaie.Execute "delete from acc1p0 where a1p01 = '" & Mid(Combo7, 1, 1) & "' and a1p02 = 'C' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text17 & "' "
   'Modify  by Amy 2020/07/01 +a0e07 因改為key
   adoTaie.Execute "delete from acc0e0 where a0e23='" & Mid(Combo7, 1, 1) & "' And a0e01 = '" & Text10 & "' and a0e02 = '" & Text9 & "' And a0e07='" & Combo3 & "' "
   SumShow
   AdodcRefresh
   AdodcClear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  金額合計檢核
'
'*************************************************
Public Function AmountCheck() As String
   Dim dblAmount As Double
   Dim strCombo7 As String 'Add by Amy 2014/01/21
   
   'Modify by Amy  改combo7顯示 原:Text5
   If Combo7 = "" Then
        strCombo7 = ""
   Else
        strCombo7 = Replace(Mid(Combo7, InStr(1, Combo7, ":") + 1), ",", "")
   End If
   
   If strCombo7 = Text6 Then
         AmountCheck = MsgText(602)
   'Add by Morgan 2006/10/30
   ElseIf Adodc1.Recordset.State = adStateOpen Then
      'Modify by Morgan 2009/11/9 會有多筆應付帳款
      'If Adodc1.Recordset.EOF = False Then
      '   If Val(Text5) = Val("" & Adodc1.Recordset.Fields("a1p07")) Then
      Adodc1.Recordset.MoveFirst
      Do While Not Adodc1.Recordset.EOF
         'Modified by Morgan 2023/9/11
         'Adodc1.Recordset.Find "a1p05='2112'", , adSearchForward, Adodc1.Recordset.AbsolutePosition
         'If Not Adodc1.Recordset.EOF Then
         '   dblAmount = dblAmount + Val("" & Adodc1.Recordset.Fields("a1p07"))
         '   Adodc1.Recordset.MoveNext
         'End If
         If Adodc1.Recordset.Fields("a1p05") = "2112" Then
            dblAmount = dblAmount + Val("" & Adodc1.Recordset.Fields("a1p07"))
         End If
         Adodc1.Recordset.MoveNext
         'end 2023/9/11
      Loop
      If dblAmount > 0 Then
         'Modify by Amy 2014/01/21 原:Val(Text5)
         If Val(strCombo7) = dblAmount Then
      'end 2009/11/9
            AmountCheck = MsgText(602)
         End If
      End If
   'end 2006/10/30
   End If
End Function
'Add by Morgan 2011/8/10 檢查開票日期與付款日期是否一致
Public Function CheckBillDate() As Boolean
   With Adodc1.Recordset
   .MoveFirst
   Do While Not .EOF
      If "" & .Fields("a1p18") <> ChangeTDateStringToTString(MaskEdBox7) Then
         If MsgBox("開票日期與付款日期不同，是否要繼續？", vbYesNo + vbDefaultButton2) = vbYes Then
            Exit Do
         Else
            Exit Function
         End If
      End If
      .MoveNext
   Loop
   End With
   CheckBillDate = True
End Function

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = adoacc0q0.Bookmark & MsgText(35) & adoacc0q0.RecordCount
End Sub

'*************************************************
'  關閉分錄欄位輸入狀態
'
'*************************************************
Public Sub FormDisabled()
   Text11.Enabled = False
   Combo3.Enabled = False
   Text10.Enabled = False
   Text9.Enabled = False
   Text8.Enabled = False
   Text13.Enabled = False
   MaskEdBox5.Enabled = False
   MaskEdBox6.Enabled = False
   Combo1.Enabled = False
   Combo2.Enabled = False
   Text7.Enabled = False
   Combo4.Enabled = False
   Combo5.Enabled = False
   Command2.Enabled = False
   Text18.Enabled = False
   Text19.Enabled = False
   Text21.Enabled = False
   Text22.Enabled = False
   SetToolBar  'Add by Amy 2014/01/21 上、下、第一、最後一筆不可使用
End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'
'*************************************************
Public Sub FormEnabled()
   'Add by Amy 2014/09/30
   If strSaveConfirm = MsgText(3) Then
        '新增
        MaskEdBox7.Enabled = True
   End If
   'end 2014/09/30
   Text11.Enabled = True
   Combo3.Enabled = True
   Text10.Enabled = True
   'Add by Amy 2014/02/10
   StrA0H01 = ""
   If GetBankData(Text11, StrA0H01) <> "" Then
        Combo3.Enabled = False
        Text10.Enabled = False
   End If
   'end 2014/02/10
   Text9.Enabled = True
   Text8.Enabled = True
   Text13.Enabled = True
   MaskEdBox5.Enabled = True
   MaskEdBox6.Enabled = True
   Combo1.Enabled = True
   Combo2.Enabled = True
   Text7.Enabled = True
   Combo4.Enabled = True
   Combo5.Enabled = True
   Command2.Enabled = True
   Text18.Enabled = True
   Text19.Enabled = True
   Text21.Enabled = True
   Text22.Enabled = True
   SetToolBar 'Add by Amy 2014/01/21 上、下、第一、最後一筆不可使用
End Sub

'*************************************************
'  重新整理國內付款資料
'
'*************************************************
Public Sub Acc0q0Refresh()
On Error GoTo Checking
   If adoacc0q0.State = adStateOpen Then
      adoacc0q0.Close
   End If
   adoacc0q0.CursorLocation = adUseClient
   adoacc0q0.MaxRecords = intMax
   '93.9.13 MODIFY BY SONIA
   'adoacc0q0.Open "select * from acc0q0 where a0q01||a0q03 >= '" & Val(FCDate(MaskEdBox7.Text)) & Text2 & "' order by a0q17 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modify by Amy 2014/01/21 +公司別
   adoacc0q0.Open "select * from acc0q0 where a0q01 = " & Val(FCDate(MaskEdBox7.Text)) & _
      " AND a0q03 = '" & Text2 & "' And a0q19='" & Combo7.Tag & "' order by A0Q03,a0q17 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '93.9.13 END
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  借貸方檢核
'
'*************************************************
Public Function CreDebCheck() As String
   If Text6 = Text15 Then
      CreDebCheck = MsgText(602)
   End If
End Function

Private Sub Text9_Validate(Cancel As Boolean)
   Select Case Text11
      Case "2111", "110201", "110214"
         Combo5 = IIf(MaskEdBox6.Text <> MsgText(29), FCDate(MaskEdBox6.Text), "") & "/" & Text9 & "/" & Text4
   End Select
End Sub

'Add by Amy 2014/01/17 由.bas 搬回
Public Sub Frmacc1180_Save()
Dim adocheck As New ADODB.Recordset
Dim strPayTime As String
Dim strSql As String
Dim bolAddNew As Boolean 'Add by Morgan 2007/1/15 是否新增acc1q0

On Error GoTo Checking
      If Text2 = MsgText(601) Then
         MsgBox MsgText(10) & Label6, , MsgText(5)
         strControlButton = MsgText(602)
         Text2.SetFocus
         Exit Sub
      Else
         If MaskEdBox7.Text = MsgText(601) Or MaskEdBox7.Text = MsgText(29) Then
            MsgBox Label16 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            MaskEdBox7.SetFocus
            Exit Sub
         Else
            If DateCheck(MaskEdBox7.Text) = MsgText(603) Then
               MsgBox Label16 & MsgText(63), , MsgText(5)
               strControlButton = MsgText(602)
               MaskEdBox7.SetFocus
               Exit Sub
            End If
         End If
         
         Select Case Text1
            Case Mid(ComboItem(91), 1, 1)
               If ExistCheck("acc0i0", "a0i01", Text2, Label6) = False Then
                  strControlButton = MsgText(602)
                  Text2.SetFocus
                  Exit Sub
               End If
            Case Mid(ComboItem(92), 1, 1)
               If ExistCheck("customer", "cu01 || cu02", IIf(Len(Text2) = 6, AfterZero(Text2), Text2), Label6) = False Then
                  strControlButton = MsgText(602)
                  Text2.SetFocus
                  Exit Sub
               End If
            Case Mid(ComboItem(93), 1, 1)
               If ExistCheck("staff", "st01", Text2, Label6) = False Then
                  strControlButton = MsgText(602)
                  Text2.SetFocus
                  Exit Sub
               End If
            '2006/1/4 ADD BY SONIA
            Case Else
               MsgBox MsgText(161), , MsgText(5)
               strControlButton = MsgText(602)
               Text1.SetFocus
               Exit Sub
            '2006/1/4 END
         End Select
         If CheckLen(Label8, Text4, 100) = MsgText(603) Then
            strControlButton = MsgText(602)
            Text4.SetFocus
            Exit Sub
         End If
         '2011/10/18 add by sonia 檢查地址
         'Modify by Amy 2015/11/10 不為空且是新增或有修改 才檢查
         If Text16 <> MsgText(601) And Text16.Tag <> Text16 Then
            If CheckTaiwanAddr(Text16, "000", "聯絡地址") = False Then
                strControlButton = MsgText(602)
                Text16.SetFocus
                Exit Sub
            End If
         End If
         'end 2015/11/10
         '2011/10/18 end
      End If
      '2006/1/10 ADD BY SONIA
      If adoquery.State = adStateOpen Then
         adoquery.Close
      End If
      adoquery.CursorLocation = adUseClient
      'Modify by Amy 2014/01/21 +公司別
      adoquery.Open "select ax210 from acc1p0, acc021 where a1p01 = ax201 and a1p22 = ax202 and a1p03 = ax203 and ax210 is not null and a1p04 = '" & Text17 & "' And a1p01='" & Mid(Combo7, 1, 1) & "' ", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         MsgBox MsgText(14), , MsgText(5)
         strControlButton = MsgText(602)
         adoquery.Close
         AdodcClear 'Add by Amy 2014/01/21 清除從text2_validate Insert 預帶的資料
         Exit Sub
      End If
      adoquery.Close
      '2006/1/10 END
      bolAddNew = True 'Add by Morgan 2007/1/15
      If strSaveConfirm = MsgText(3) Then
         If adoacc0q0.RecordCount <> 0 Then
            adoacc0q0.Find "a0q17 = '" & Text17 & "'", 0, adSearchForward, 1
            If adoacc0q0.EOF = False Then
'               .adoacc0q0.Find "a0q03 = '" & .Text2 & "'", 0, adSearchForward, .adoacc0q0.Bookmark
'               If .adoacc0q0.EOF = False Then
'                  .adoacc0q0.Find "a0q01 = '" & Val(FCDate(.MaskEdBox7.Text)) & "'", 0, adSearchForward, .adoacc0q0.Bookmark
'                  If .adoacc0q0.EOF = False Then

                     'Modify by Morgan 2007/1/15 還是要Update acc0q0 否則後來改的資料便不會更新
                     'Exit Sub
                     bolAddNew = False
                     'end 2007/1/15
                     
'                  End If
'               End If
            End If
         End If
         
         If bolAddNew = True Then 'Add by Morgan 2007/1/15
            adoquery.CursorLocation = adUseClient
            'Modify by Amy 2014/01/21 +公司別
            adoquery.Open "select a0q01 from acc0q0 where a0q01 = " & Val(FCDate(MaskEdBox7.Text)) & " and a0q03 = '" & Text2 & "' And a0q19='" & Mid(Combo7, 1, 1) & "' ", adoTaie, adOpenStatic, adLockReadOnly
            If adoquery.RecordCount <> 0 Then
               adoquery.Close
               MsgBox MsgText(88), , MsgText(5)
               strControlButton = MsgText(602)
               Text2.SetFocus
               Exit Sub
            End If
            adoquery.Close
            adoacc0q0.AddNew
         End If
      End If
      
      adoacc0q0.Fields("a0q19").Value = Mid(Combo7, 1, 1) 'Add by Amy 2014/01/21 +公司別
      If Text1 <> MsgText(601) Then
         adoacc0q0.Fields("a0q04").Value = Text1
      Else
         adoacc0q0.Fields("a0q04").Value = Null
      End If
      adoacc0q0.Fields("a0q01").Value = Val(FCDate(MaskEdBox7.Text))
'      If .Text16 <> MsgText(601) Then
'         .adoacc0q0.Fields("a0q15").Value = .Text16
'      Else
'         .adoacc0q0.Fields("a0q15").Value = Null
'      End If
      'Modify by Amy 2014/01/21 改combo7 原:Val(Text5)
      If Combo7 <> MsgText(601) Then
         adoacc0q0.Fields("a0q06").Value = Val(Format(Mid(Combo7, InStr(1, Combo7, ":") + 1), DAmount))
      Else
         adoacc0q0.Fields("a0q06").Value = 0
      End If
      'end 2014/01/21
      adoacc0q0.Fields("a0q03").Value = Text2
      If Text4 <> MsgText(601) Then
         adoacc0q0.Fields("a0q05").Value = Text4
      Else
         adoacc0q0.Fields("a0q05").Value = Null
      End If
      If Text16 <> MsgText(601) Then
         adoacc0q0.Fields("a0q16").Value = Text16
      Else
         adoacc0q0.Fields("a0q16").Value = Null
      End If
      If Text17 <> MsgText(601) Then
         adoacc0q0.Fields("a0q17").Value = Text17
      Else
         adoacc0q0.Fields("a0q17").Value = Null
      End If
      
      'Add by Morgan 2006/10/31
      If Text23 <> "" Then
         adoacc0q0.Fields("a0q18") = Text23
      Else
         adoacc0q0.Fields("a0q18") = Null
      End If
      
      If strSaveConfirm = MsgText(3) Then
         adoacc0q0.Fields("a0q09").Value = Val(strSrvDate(2))
         adoacc0q0.Fields("a0q10").Value = ServerTime
         adoacc0q0.Fields("a0q11").Value = strUserNum
      Else
         adoacc0q0.Fields("a0q12").Value = Val(strSrvDate(2))
         adoacc0q0.Fields("a0q13").Value = ServerTime
         adoacc0q0.Fields("a0q14").Value = strUserNum
      End If
      adoacc0q0.UpdateBatch
      strSql = ""
      Select Case Text1
         Case "1", "2", "3"
            If MaskEdBox1.Text <> MsgText(29) Then
               strSql = strSql & " and a0o05 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
            End If
            If MaskEdBox2.Text <> MsgText(29) Then
               strSql = strSql & " and a0o05 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
            End If
            '93.12.14 ADD BY SONIA
            If MaskEdBox3.Text <> MsgText(29) Then
               strSql = strSql & " and a0o06 >= " & Val(FCDate(MaskEdBox3.Text)) & ""
            End If
            If MaskEdBox4.Text <> MsgText(29) Then
               strSql = strSql & " and a0o06 <= " & Val(FCDate(MaskEdBox4.Text)) & ""
            End If
            '93.12.14 END
'         Case "2"
'            If .MaskEdBox3.Text <> MsgText(29) Then
'               strSQL = strSQL & " and a0o06 >= " & Val(FCDate(.MaskEdBox3.Text)) & ""
'            End If
'            If .MaskEdBox4.Text <> MsgText(29) Then
'               strSQL = strSQL & " and a0o06 <= " & Val(FCDate(.MaskEdBox4.Text)) & ""
'            End If
         End Select
      'Modify by Amy 2014/01/21 +公司別
      If strSaveConfirm = MsgText(3) Then
         adoTaie.Execute "update acc0o0 set a0o11 = " & Val(FCDate(MaskEdBox7.Text)) & " where a0o07='" & Mid(Combo7, 1, 1) & "' And a0o03 = '" & Text2 & "' and (a0o11 is null or a0o11 = '')" & strSql
      Else
         adoTaie.Execute "update acc0o0 set a0o11 = " & Val(FCDate(MaskEdBox7.Text)) & " where a0o07='" & Mid(Combo7, 1, 1) & "' And a0o03 = '" & Text2 & "' and (a0o11 = " & lngDate & ")" & strSql
      End If
      'end 2014/01/21
      lngDate = Val(FCDate(MaskEdBox7.Text)) 'Add by Morgan 2006/7/25
      RecordShow
      
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)

End Sub

Public Sub Frmacc1180_Delete()
On Error GoTo Checking
      'Modify by Amy 2014/01/21 +公司別
      If DeleteCheck("select a0q01 from acc0q0 where a0q19='" & Mid(Combo7, 1, 1) & "' And a0q01 = " & Val(FCDate(MaskEdBox7.Text)) & " and a0q03 = '" & Text2 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "delete from acc1p0 where  a1p01 ='" & Mid(Combo7, 1, 1) & "' and a1p02 = 'C' and a1p04 = '" & Text17 & "'"  '原:a1p01='1'
      adoacc1p0.Requery
      adoTaie.Execute "update acc0o0 set a0o11 = null where a0o07='" & Mid(Combo7, 1, 1) & "' And a0o03 = '" & Text2 & "' and a0o11 = " & Val(FCDate(MaskEdBox7.Text)) & ""
      adoTaie.Execute "delete from acc0q0 where a0q19='" & Mid(Combo7, 1, 1) & "' And a0q01 = " & Val(FCDate(MaskEdBox7.Text)) & " and a0q03 = '" & Text2 & "'"
      'end 2014/01/21
      adoacc0q0.Requery
      AdodcRefresh
      If adoacc0q0.RecordCount <> 0 Then
         adoacc0q0.MoveFirst
         RecordShow
      Else
         StatusClear
      End If
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Public Sub Frmacc1180_Clear(Optional bolClsCombo7 As Boolean = False)
      Text1 = ""
      Text17 = ""
      If MaskEdBox7.Text = MsgText(29) Or MaskEdBox7.Text = MsgText(601) Then
         MaskEdBox7.Mask = ""
         MaskEdBox7.Text = ""
         MaskEdBox7.Mask = DFormat
      End If
      MaskEdBox7.Tag = "" 'Add by Amy 2014/10/27
      If MaskEdBox1.Text = MsgText(29) Or MaskEdBox1.Text = MsgText(601) Then
         MaskEdBox1.Mask = ""
         MaskEdBox1.Text = ""
         MaskEdBox1.Mask = DFormat
      End If
      If MaskEdBox2.Text = MsgText(29) Or MaskEdBox2.Text = MsgText(29) Then
         MaskEdBox2.Mask = ""
         MaskEdBox2.Text = ""
         MaskEdBox2.Mask = DFormat
      End If
      If MaskEdBox3.Text = MsgText(29) Or MaskEdBox3.Text = MsgText(601) Then
         MaskEdBox3.Mask = ""
         MaskEdBox3.Text = ""
         MaskEdBox3.Mask = DFormat
      End If
      If MaskEdBox4.Text = MsgText(29) Or MaskEdBox4.Text = MsgText(601) Then
         MaskEdBox4.Mask = ""
         MaskEdBox4.Text = ""
         MaskEdBox4.Mask = DFormat
      End If
'      .Text16 = ""
      'Modify by Amy 2014/01/21 +if 改顯示 原:Text5 = ""
      If bolClsCombo7 = True And Combo7 <> "" Then
         Combo7.Clear
      End If
      Text2.Enabled = True
      'end 2014/01/21
      Text2 = ""
      Text3 = ""
      Text6 = ""
      Text15 = ""
      Text20 = ""
      Text4 = ""
      Text16 = ""
      Text16.Tag = "" 'Add by Amy 2015/11/10
      Text23 = "" 'Add by Morgan 2006/12/5
      AdodcRefresh
      AdodcClear
      Text1.SetFocus
End Sub

Public Sub Frmacc1180_Last()
      If AmountCheck <> MsgText(602) Then
         MsgBox MsgText(30), , MsgText(5)
         Exit Sub
      End If
      If adoacc0q0.RecordCount <> 0 Then
         adoacc0q0.MoveLast
         FormShow
         SumShow
         AdodcRefresh
         AdodcClear
         RecordShow
      End If
End Sub

Public Sub Frmacc1180_Previous()
      If AmountCheck <> MsgText(602) Then
         MsgBox MsgText(30), , MsgText(5)
         Exit Sub
      End If
      If adoacc0q0.BOF = False Then
         adoacc0q0.MovePrevious
         If adoacc0q0.BOF Then
            adoacc0q0.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         FormShow
         SumShow
         AdodcRefresh
         AdodcClear
         RecordShow
      End If
End Sub

Public Sub Frmacc1180_Next()
      If AmountCheck <> MsgText(602) Then
         MsgBox MsgText(30), , MsgText(5)
         Exit Sub
      End If
      If adoacc0q0.EOF = False Then
         adoacc0q0.MoveNext
         If adoacc0q0.EOF Then
            adoacc0q0.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         FormShow
         SumShow
         AdodcRefresh
         AdodcClear
         RecordShow
      End If
End Sub

Public Sub Frmacc1180_First()
      If AmountCheck <> MsgText(602) Then
         MsgBox MsgText(30), , MsgText(5)
         Exit Sub
      End If
      If adoacc0q0.RecordCount <> 0 Then
         adoacc0q0.MoveFirst
         FormShow
         SumShow
         AdodcRefresh
         AdodcClear
         RecordShow
      End If
End Sub
'end 2014/01/17

'Add by Amy 2014/01/21 由acc_var.bas 搬回 並修改
Public Sub FormCheck()
    Dim strMsg As String 'Add by Amy 2014/09/30
    
    If Combo7 = "" Then
        MsgBox Label9 & "必要欄位,請選取...", , MsgText(5)
        Combo7.SetFocus
        strControlButton = MsgText(602)
        Exit Sub
    End If
    
    If CreDebCheck <> MsgText(602) Or Val(Text6) = 0 Or Val(Text15) = 0 Then
        MsgBox MsgText(11), , MsgText(5)
        strControlButton = MsgText(602)
        Exit Sub
    End If
    'Modify by Morgan 2006/11/10
    'If .Text6 <> .Text5 Then
    If AmountCheck <> MsgText(602) Then
        MsgBox MsgText(59), , MsgText(5)
        strControlButton = MsgText(602)
        Text1.SetFocus
        Exit Sub
    End If

    'Add by Morgan 2011/8/10
    If CheckBillDate = False Then
        strControlButton = MsgText(602)
        Exit Sub
    End If
    'Add by Amy 2014/09/30 +系統日檢查
    If MaskEdBox7.Enabled = True Then
        If ChkWorkData(Mid(Combo7, 1, 1), DBDATE(MaskEdBox7), strMsg) = False Then
            MsgBox Label16 & strMsg, , MsgText(5)
            strControlButton = MsgText(602)
            MaskEdBox7.SetFocus
            Exit Sub
        End If
    End If
    'end 2014/09/30
End Sub

Private Sub SetToolBar()
   '上、下、第一、最後一筆不可使用 (因抓資料的寫法永遠只抓一筆)
   With Forms(0)
      .Toolbar1.Buttons.Item(13).Enabled = False
      .Toolbar1.Buttons.Item(14).Enabled = False
      .Toolbar1.Buttons.Item(15).Enabled = False
      .Toolbar1.Buttons.Item(16).Enabled = False
   End With
End Sub

Private Function GetStrWhere() As String
    '畫面上條件
    GetStrWhere = ""
    'Modify by Amy 2014/01/28 因抓a1p18使速度變慢所以改抓a0o05
    If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
       GetStrWhere = GetStrWhere & " and a0o05 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
    End If
    If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
       GetStrWhere = GetStrWhere & " and a0o05 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
    End If
    'end 2014/01/28
    '93.12.8 add by sonia
    If MaskEdBox3.Text <> MsgText(601) And MaskEdBox3.Text <> MsgText(29) Then
       GetStrWhere = GetStrWhere & " and a0o06 >= " & Val(FCDate(MaskEdBox3.Text)) & ""
    End If
    If MaskEdBox4.Text <> MsgText(601) And MaskEdBox4.Text <> MsgText(29) Then
       GetStrWhere = GetStrWhere & " and a0o06 <= " & Val(FCDate(MaskEdBox4.Text)) & ""
    End If
    '93.12.8 end
End Function

Private Function CheckAcc1p0(strNo1 As String, strNo2 As String) As Boolean
    '確認Acc1p0 是否有資料
    Dim strSql As String, adoTmp As ADODB.Recordset
    
    CheckAcc1p0 = False
    strSql = "Select * From Acc1p0 Where a1p01='" & strNo1 & "' And a1p04='" & strNo2 & "' "
    intI = 1
    Set adoTmp = ClsLawReadRstMsg(intI, strSql)
    If intI = 1 Then
        CheckAcc1p0 = True
    End If
    adoTmp.Close
End Function

Private Function GetBankData(strA0H08 As String, ByRef StrA0H01 As String) As String
    '依會計科目取得  開票帳號及開票銀行(代號)
    GetBankData = ""
    Dim strSql As String, adoTmp As ADODB.Recordset
    
    strSql = "Select * From Acc0H0 Where A0H08='" & strA0H08 & "' "
    intI = 1
    Set adoTmp = ClsLawReadRstMsg(intI, strSql)
    If intI = 1 Then
        GetBankData = adoTmp.Fields("A0H02")
        StrA0H01 = adoTmp.Fields("A0H01")
    End If
    adoTmp.Close
End Function
'end 2014/01/21

'Add by Amy 2014/10/27 為資料一致更新acc1p0
Public Sub UpdateAcc1p0()
    Dim strUpd As String
    
On Error GoTo ChkHand
    
    If strSaveConfirm = MsgText(3) Or (strSaveConfirm = MsgText(4) And Val(MaskEdBox7.Tag) <> Val(FCDate(MaskEdBox7))) Then
       strUpd = "Update Acc1p0 set a1p18=" & Val(FCDate(MaskEdBox7)) & _
                     " Where a1p01='" & Mid(Combo7, 1, 1) & "' And a1p04='" & Text17 & "' And a1p02='C' "
        adoTaie.Execute strUpd
    End If

ChkHand:
    If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox "UpdateAcc1p0 錯誤:" & Err.Description, , MsgText(5)
   strControlButton = MsgText(602)
End Sub

Public Sub SetData(ByVal strKeyCode As String)
    Select Case strKeyCode
        Case "F3", "F9"
            '解改日期存檔再修改不會存acc1p0 (因tag只記錄前一次改前資料)
            MaskEdBox7.Tag = Val(FCDate(MaskEdBox7))
        Case Else
    End Select
End Sub
'end 2014/10/27

