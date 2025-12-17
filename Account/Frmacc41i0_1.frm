VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc41i0_1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000B&
   BorderStyle     =   1  '單線固定
   Caption         =   "財產報廢作業"
   ClientHeight    =   5505
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   8730
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   8730
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7950
      Picture         =   "Frmacc41i0_1.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   22
      ToolTipText     =   "取消"
      Top             =   4905
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7350
      Picture         =   "Frmacc41i0_1.frx":066A
      Style           =   1  '圖片外觀
      TabIndex        =   21
      ToolTipText     =   "清除畫面"
      Top             =   4905
      Width           =   495
   End
   Begin VB.TextBox Txt1 
      BackColor       =   &H00E0E0E0&
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
      Height          =   315
      Index           =   3
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   1635
      Width           =   1215
   End
   Begin VB.TextBox txtA2B07 
      BackColor       =   &H00FFFFFF&
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
      Height          =   330
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   13
      Top             =   1635
      Width           =   615
   End
   Begin VB.TextBox txtA2B06 
      BackColor       =   &H00FFFFFF&
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
      Height          =   315
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   12
      Top             =   1320
      Width           =   1215
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
      Left            =   1800
      TabIndex        =   10
      Text            =   "Combo2"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtA2B03 
      BackColor       =   &H00FFFFFF&
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
      Height          =   330
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   9
      Top             =   1320
      Width           =   612
   End
   Begin VB.TextBox txtA2B16 
      BackColor       =   &H00FFFFFF&
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
      Height          =   330
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   5
      Top             =   525
      Width           =   612
   End
   Begin VB.TextBox txtA2B02 
      BackColor       =   &H00FFFFFF&
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
      Height          =   330
      Left            =   6360
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   6
      Top             =   510
      Width           =   612
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   6960
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   525
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Default         =   -1  'True
      Height          =   300
      Left            =   2280
      Picture         =   "Frmacc41i0_1.frx":0F34
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   120
      Width           =   350
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc41i0_1.frx":1036
      Height          =   1050
      Left            =   270
      TabIndex        =   23
      Top             =   2715
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   1852
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   14
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
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "a1p03"
         Caption         =   "項次"
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
         DataField       =   "a1p05"
         Caption         =   "科目代號"
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
         DataField       =   "a0102"
         Caption         =   "科目名稱"
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
      BeginProperty Column04 
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
      BeginProperty Column05 
         DataField       =   "a0902"
         Caption         =   "部門別"
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
            Alignment       =   2
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2580.095
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column06 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   30
      Top             =   2790
      Visible         =   0   'False
      Width           =   990
      _ExtentX        =   1746
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
   Begin VB.TextBox txtA1P03 
      BackColor       =   &H00E0E0E0&
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
      Height          =   300
      Left            =   270
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   24
      Top             =   4545
      Width           =   492
   End
   Begin VB.TextBox txtA1P07 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3390
      MaxLength       =   7
      TabIndex        =   17
      Top             =   4545
      Width           =   1572
   End
   Begin VB.TextBox txtA1P08 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5070
      MaxLength       =   7
      TabIndex        =   18
      Top             =   4545
      Width           =   1572
   End
   Begin VB.TextBox txtTot1 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
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
      Height          =   300
      Left            =   4890
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   3795
      Width           =   1368
   End
   Begin VB.TextBox txtTot2 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
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
      Height          =   300
      Left            =   6270
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   3795
      Width           =   1356
   End
   Begin VB.TextBox txtA1P05 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   870
      MaxLength       =   6
      TabIndex        =   16
      Top             =   4545
      Width           =   972
   End
   Begin VB.TextBox Txt1 
      BackColor       =   &H00E0E0E0&
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
      Height          =   300
      Index           =   4
      Left            =   1830
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   4545
      Width           =   1452
   End
   Begin VB.TextBox txtA1P06 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6750
      MaxLength       =   3
      TabIndex        =   19
      Top             =   4545
      Width           =   612
   End
   Begin VB.TextBox Txt1 
      BackColor       =   &H00E0E0E0&
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
      Height          =   300
      Index           =   5
      Left            =   7350
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   4545
      Width           =   1212
   End
   Begin VB.TextBox txtA2B01 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Txt1 
      BackColor       =   &H00E0E0E0&
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
      Height          =   315
      Index           =   0
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   525
      Width           =   3135
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   315
      Left            =   4440
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   315
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.25
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
      Left            =   3840
      TabIndex        =   14
      Top             =   1635
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.25
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
      Left            =   4890
      TabIndex        =   15
      Top             =   1635
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.TextBox Txt1 
      BackColor       =   &H00E0E0E0&
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
      Height          =   315
      Index           =   2
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Txt1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   6360
      MaxLength       =   7
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSForms.TextBox txtA2B04 
      Height          =   380
      Left            =   1200
      TabIndex        =   8
      Top             =   900
      Width           =   6615
      VariousPropertyBits=   -1466941409
      BackColor       =   16777215
      MaxLength       =   100
      ScrollBars      =   2
      Size            =   "11668;670"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtA2B09 
      Height          =   585
      Left            =   1200
      TabIndex        =   3
      Top             =   1980
      Width           =   7335
      VariousPropertyBits=   -1467989989
      BackColor       =   16777215
      MaxLength       =   500
      ScrollBars      =   2
      Size            =   "12938;1032"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo3 
      Height          =   345
      Left            =   870
      TabIndex        =   20
      Top             =   4890
      Width           =   6015
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "10610;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "報廢部門"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   5400
      TabIndex        =   51
      Top             =   150
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label16 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "報廢日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   2640
      TabIndex        =   50
      Top             =   150
      Width           =   975
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "財產名稱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   49
      Top             =   945
      Width           =   900
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "摘要"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   270
      TabIndex        =   48
      Top             =   4950
      Width           =   975
   End
   Begin VB.Label Label16 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "未折減餘額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6000
      TabIndex        =   46
      Top             =   1665
      Width           =   1215
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "備　註"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   45
      Top             =   2010
      Width           =   900
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "使用月份"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   44
      Top             =   1695
      Width           =   900
   End
   Begin VB.Label Label16 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "取得原價"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   6240
      TabIndex        =   43
      Top             =   1350
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "類　別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   42
      Top             =   555
      Width           =   750
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "所在地"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   41
      Top             =   1350
      Width           =   900
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   40
      Top             =   1710
      Width           =   255
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "攤提期間"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   39
      Top             =   1665
      Width           =   900
   End
   Begin VB.Label Label9 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "項次"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   38
      Top             =   4305
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "會計科目"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1590
      TabIndex        =   37
      Top             =   4305
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "借方金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3630
      TabIndex        =   36
      Top             =   4305
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "貸方金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5310
      TabIndex        =   35
      Top             =   4305
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "部門別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7230
      TabIndex        =   34
      Top             =   4305
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1245
      Left            =   150
      Top             =   4185
      Width           =   8535
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   30
      Top             =   4905
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label15 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4110
      TabIndex        =   33
      Top             =   3825
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "取得日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   28
      Top             =   1350
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "編　號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   150
      Width           =   900
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   555
      Width           =   900
   End
End
Attribute VB_Name = "Frmacc41i0_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2017/03/06 財產報廢作業
Option Explicit
Public adoacc2b0 As New ADODB.Recordset '財產目錄資料
Public adoacc2b0a As New ADODB.Recordset
Public adoacc1p0 As New ADODB.Recordset '報廢傳票資料
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Dim m_A2B17 As String
Dim m_A2B22 As String '首筆折舊傳票號碼(財產目錄)

Dim bolDelAll As Boolean '刪除全部明細

'報廢傳票的設定
Private Const m_A1P02 = "N" '傳票分錄別
Dim bolA1P27 As Boolean  '是否更新-報廢傳票
Dim strAccNo As String   '傳票號碼

Private Sub Combo3_GotFocus()
   OpenIme
End Sub

'清除畫面(新增項次)
Private Sub Command1_Click()
   AdodcClear
   txtA1P03 = GetSeqNo(txtA2B01, MaskEdBox2)
   txtA1P05.SetFocus
End Sub

'取消(刪除項次)
Private Sub Command2_Click()
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   
   If bolDelAll = True Then
        adoTaie.Execute " delete from acc1p0 where a1p02 = '" & m_A1P02 & "' and a1p04 = '" & txtA2B01 & MaskEdBox2.Tag & "' "
   Else
        Adodc1.Recordset.Find "A1P03 = " & CNULL(txtA1P03), 0, adSearchForward, 1
        If Adodc1.Recordset.EOF Then
           Exit Sub
        End If
        adoTaie.Execute " delete from acc1p0 where a1p02 = '" & m_A1P02 & "' and a1p03 = '" & txtA1P03 & "' and a1p04 = '" & CompA1P04(txtA2B01, MaskEdBox2) & "' "
   End If
   
   AdodcRefresh
   AdodcClear
   SumShow
   txtA1P03 = GetSeqNo(txtA2B01, MaskEdBox2)  '重抓項次
   
   If adoacc1p0.RecordCount = 0 Then
      StatusClear
   Else
      RecordShow
   End If

End Sub

Private Sub Command3_Click()
   If txtA2B01 = MsgText(601) Then
      Exit Sub
   End If

   txtA2B01 = PUB_ChgNumeralStyle(txtA2B01)  'Added by Lydia 2021/03/16 轉換全形數字變成半形數字

   Acc2b0Refresh
   If adoacc2b0.RecordCount <> 0 Then
      adoacc2b0.Find "a2b01 = '" & txtA2B01 & "'", 0, adSearchForward, 1
      If adoacc2b0.EOF = False Then
         FormShow
         AdodcRefresh
         SumShow
         RecordShow
      End If
   Else
      MsgBox MsgText(33), , MsgText(5)
   End If
   AdodcClear
End Sub

'編號查詢
Private Sub Command3_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command3_Click
         Exit Sub
   End Select
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   AdodcShow
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   
   adoacc2b0.Find "a2b01 = '" & strItemNo & "'", 0, adSearchForward, 1
   If adoacc2b0.EOF = False Then
      FormShow
      AdodcRefresh
      SumShow
      RecordShow
      FormDisabled
   End If
   strItemNo = MsgText(601)
End Sub

'Added by Lydia 2021/12/01
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Added by Lydia 2021/12/01 Form2.0 記錄鍵盤傳入順序
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'檢查是否可修改明細項(報廢傳票)
'iAct:0=修改,1=刪除
Public Function EditCheck(Optional iAct As Integer = 0, Optional bolMsg As Boolean = True) As Boolean
Dim inJ As Integer
Dim rsR1 As New ADODB.Recordset

bolA1P27 = False

If strAccNo <> MsgText(601) Then
    If iAct = 1 Then
       MsgBox "此筆資料已轉傳票，不可刪除！"
       Exit Function
    Else
       '檢查傳票是否已過帳
       If PUB_CheckPosted(strAccNo, bolMsg, txtA2B16.Text) = True Then
          Exit Function
       Else
          bolA1P27 = True
       End If
    End If
End If

EditCheck = True

End Function

Private Sub Form_Load()
   '表單初始化
   'Modified by Lydia 2021/12/01 height 5730=>5940
   PUB_InitForm Me, 8850, 5940, strBackPicPath1

   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   MaskEdBox3.Mask = Mid(DFormat, 1, 6)
   MaskEdBox4.Mask = Mid(DFormat, 1, 6)
      
   '類別
   Combo1.Clear
   Combo1.AddItem "", 0
   strExc(1) = Pub_GetA2b02Name
   For intI = 1 To Val(strExc(1))
       Combo1.AddItem Pub_GetA2b02Name(Trim(intI)), intI
   Next intI
   
   '所在地
   Combo2.Clear
   Combo2.AddItem "", 0
   strExc(1) = Pub_GetA2b03Sname
   For intI = 1 To Val(strExc(1))
       Combo2.AddItem Pub_GetA2b03Sname(Trim(intI)), intI
   Next intI
   
   Combo3.Text = ""
   OpenTable
   
   'Remove by Lydia 2017/05/18 預設不顯示
   'If adoacc2b0.RecordCount <> 0 Then
   '   Frmacc41i0_1_First
   'End If
   
   FormDisabled
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   strTrackMode = "" 'Added by Lydia 2021/12/01 Form2.0 記錄鍵盤傳入順序(清除)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc41i0_1 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking

   adoacc2b0.CursorLocation = adUseClient
   adoacc2b0.Open "select * from acc2b0 where nvl(a2b19,0) > 0 order by a2b01 asc ", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc2b0a.CursorLocation = adUseClient
   adoacc2b0a.Open "select * from acc2b0 where nvl(a2b19,0) = 0 order by a2b01 asc ", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc1p0.CursorLocation = adUseClient
   adoacc1p0.Open "select * from acc1p0 where a1p02='" & m_A1P02 & "' and a1p04 = '" & CompA1P04(txtA2B01, MaskEdBox2) & "' order by a1p03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc1p0, acc010, acc090 where a1p05=a0101(+) and a1p06=a0901(+) and a1p02='" & m_A1P02 & "' and a1p04 = '" & CompA1P04(txtA2B01, MaskEdBox2) & "' order by a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  重新整理財產目錄主檔
'
'*************************************************
Public Sub Acc2b0Refresh()
On Error GoTo Checking
   If adoacc2b0.State = adStateOpen Then
      adoacc2b0.Close
   End If
   adoacc2b0.CursorLocation = adUseClient
   adoacc2b0.Open "select * from acc2b0 where nvl(a2b19,0) > 0 and a2b01>=" & txtA2B01 & " order by a2b01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'組合報廢傳票-單據號碼
Private Function CompA1P04(ByVal SNo As String, ByVal Sdate As String) As String
    If SNo <> "" Then
       CompA1P04 = SNo & Val(FCDate(Sdate))
    End If
End Function

'*************************************************
'  顯示財產目錄資料(報廢)
'
'*************************************************
Public Sub FormShow()
   
   If strSaveConfirm = MsgText(3) Then
       Call ShowAccDept(True)
   Else
       Call ShowAccDept(False)
   End If
   
   '編號
   txtA2B01 = "" & adoacc2b0.Fields("a2b01").Value
   txtA2B01.Tag = txtA2B01.Text
   '類別
   txtA2B02 = "" & adoacc2b0.Fields("a2b02").Value
   txtA2B02.Tag = txtA2B02.Text
   '所在地
   txtA2B03 = "" & adoacc2b0.Fields("a2b03").Value
   '財產名稱
   txtA2B04 = "" & adoacc2b0.Fields("a2b04").Value
   '取得日期
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc2b0.Fields("a2b05").Value) Then
      MaskEdBox1.Text = MsgText(29)
   Else
      MaskEdBox1.Text = CFDate(adoacc2b0.Fields("a2b05").Value)
   End If
   MaskEdBox1.Tag = "" & adoacc2b0.Fields("a2b05").Value
   MaskEdBox1.Mask = DFormat
   
   '取得原價
   txtA2B06 = "" & adoacc2b0.Fields("a2b06").Value
   txtA2B06.Tag = txtA2B06.Text
   '使用月份
   txtA2B07 = "" & adoacc2b0.Fields("a2b07").Value
   txtA2B07.Tag = txtA2B07.Text

   '報廢傳票日期
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(adoacc2b0.Fields("a2b19").Value) Then
      MaskEdBox2.Text = MsgText(29)
   Else
      MaskEdBox2.Text = CFDate(adoacc2b0.Fields("a2b19").Value)
   End If
   MaskEdBox2.Tag = "" & adoacc2b0.Fields("a2b19").Value
   MaskEdBox2.Mask = DFormat
   '備註
   txtA2B09 = "" & adoacc2b0.Fields("a2b09").Value
   '公司別
   txtA2B16 = "" & adoacc2b0.Fields("a2b16").Value
   txtA2B16.Tag = txtA2B16.Text
   '每月固定傳票編號
   m_A2B17 = "" & adoacc2b0.Fields("a2b17").Value
   '攤提期間
   MaskEdBox3.Mask = MsgText(601)
   If IsNull(adoacc2b0.Fields("a2b20").Value) Then
      MaskEdBox3.Text = MsgText(601)
   Else
      MaskEdBox3.Text = IIf(Len(adoacc2b0.Fields("a2b20").Value) < 5, "0" & Mid(adoacc2b0.Fields("a2b20").Value, 1, 2) & "/" & Mid(adoacc2b0.Fields("a2b20").Value, 3, 2), Mid(adoacc2b0.Fields("a2b20").Value, 1, 3) & "/" & Mid(adoacc2b0.Fields("a2b20").Value, 4, 2))
   End If
   MaskEdBox3.Mask = Mid(DFormat, 1, 6)
   MaskEdBox4.Mask = MsgText(601)
   If IsNull(adoacc2b0.Fields("a2b21").Value) Then
      MaskEdBox4.Text = MsgText(601)
   Else
      MaskEdBox4.Text = IIf(Len(adoacc2b0.Fields("a2b21").Value) < 5, "0" & Mid(adoacc2b0.Fields("a2b21").Value, 1, 2) & "/" & Mid(adoacc2b0.Fields("a2b21").Value, 3, 2), Mid(adoacc2b0.Fields("a2b21").Value, 1, 3) & "/" & Mid(adoacc2b0.Fields("a2b21").Value, 4, 2))
   End If
   MaskEdBox4.Mask = Mid(DFormat, 1, 6)
   
   '首筆折舊傳票號碼(財產目錄)
   m_A2B22 = "" & adoacc2b0.Fields("a2b22")
   
   '未折減金額
   Call CaculateAmt
End Sub

'*************************************************
'  顯示財產目錄資料
'
'*************************************************
Public Sub Adoacc2b0aRefresh()
  
   If strSaveConfirm = MsgText(3) Then
       Call ShowAccDept(True)
   Else
       Call ShowAccDept(False)
   End If
   
   '編號
   txtA2B01 = "" & adoacc2b0a.Fields("a2b01").Value
   txtA2B01.Tag = txtA2B01.Text
   '類別
   txtA2B02 = "" & adoacc2b0a.Fields("a2b02").Value
   txtA2B02.Tag = txtA2B02.Text
   '所在地
   txtA2B03 = "" & adoacc2b0a.Fields("a2b03").Value
   '財產名稱
   txtA2B04 = "" & adoacc2b0a.Fields("a2b04").Value
   '取得日期
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc2b0a.Fields("a2b05").Value) Then
      MaskEdBox1.Text = MsgText(29)
   Else
      MaskEdBox1.Text = CFDate(adoacc2b0a.Fields("a2b05").Value)
   End If
   MaskEdBox1.Tag = "" & adoacc2b0a.Fields("a2b05").Value
   MaskEdBox1.Mask = DFormat
   
   '取得原價
   txtA2B06 = "" & adoacc2b0a.Fields("a2b06").Value
   txtA2B06.Tag = txtA2B06.Text
   '使用月份
   txtA2B07 = "" & adoacc2b0a.Fields("a2b07").Value
   txtA2B07.Tag = txtA2B07.Text

   '報廢傳票日期
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(adoacc2b0a.Fields("a2b19").Value) Then
      MaskEdBox2.Text = MsgText(29)
   Else
      MaskEdBox2.Text = CFDate(adoacc2b0a.Fields("a2b19").Value)
   End If
   MaskEdBox2.Tag = "" & adoacc2b0a.Fields("a2b19").Value
   MaskEdBox2.Mask = DFormat
   '備註
   txtA2B09 = "" & adoacc2b0a.Fields("a2b09").Value
   '公司別
   txtA2B16 = "" & adoacc2b0a.Fields("a2b16").Value
   txtA2B16.Tag = txtA2B16.Text
   '每月固定傳票編號
   m_A2B17 = "" & adoacc2b0a.Fields("a2b17").Value
   '攤提期間
   MaskEdBox3.Mask = MsgText(601)
   If IsNull(adoacc2b0a.Fields("a2b20").Value) Then
      MaskEdBox3.Text = MsgText(601)
   Else
      MaskEdBox3.Text = IIf(Len(adoacc2b0a.Fields("a2b20").Value) < 5, "0" & Mid(adoacc2b0a.Fields("a2b20").Value, 1, 2) & "/" & Mid(adoacc2b0a.Fields("a2b20").Value, 3, 2), Mid(adoacc2b0a.Fields("a2b20").Value, 1, 3) & "/" & Mid(adoacc2b0a.Fields("a2b20").Value, 4, 2))
   End If
   MaskEdBox3.Mask = Mid(DFormat, 1, 6)
   MaskEdBox4.Mask = MsgText(601)
   If IsNull(adoacc2b0a.Fields("a2b21").Value) Then
      MaskEdBox4.Text = MsgText(601)
   Else
      MaskEdBox4.Text = IIf(Len(adoacc2b0a.Fields("a2b21").Value) < 5, "0" & Mid(adoacc2b0a.Fields("a2b21").Value, 1, 2) & "/" & Mid(adoacc2b0a.Fields("a2b21").Value, 3, 2), Mid(adoacc2b0a.Fields("a2b21").Value, 1, 3) & "/" & Mid(adoacc2b0a.Fields("a2b21").Value, 4, 2))
   End If
   MaskEdBox4.Mask = Mid(DFormat, 1, 6)
   
   '首筆折舊傳票號碼(財產目錄)
   m_A2B22 = "" & adoacc2b0a.Fields("a2b22")
   
   '未折減金額
   Call CaculateAmt
End Sub

'*************************************************
'  清除欄位資料
'
'*************************************************
Public Sub AdodcClear()
   txtA1P03 = ""
   txtA1P05 = ""
   txtA1P06 = ""
   txtA1P07 = ""
   txtA1P08 = ""
   Combo3 = ""
   txt1(4) = ""
   txt1(5) = ""
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   
   Call PUB_SaveTrackMode(1, KeyCode)  'Added by Lydia 2021/12/01 Form2.0 記錄鍵盤傳入順序
   
   Select Case KeyCode
      Case vbKeyInsert
        
         If strSaveConfirm <> MsgText(4) Then
            Exit Sub
         End If
         'Remove by Lydia 2017/05/22
         'If Val(Txt1(3)) = 0 Then
         '   MsgBox "未折減餘額為零！", vbCritical
         '   Exit Sub
         'End If
         'Added by Lydia 2021/12/01 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
         If PUB_ChkTrackMode = False Then
             Exit Sub
         End If
         'Added by Lydia 2021/12/01 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
         If PUB_ChkUniText(Me, , True, "ComboBox") = False Then
             Exit Sub
         End If
         'end 2021/12/01
         'end 2021/12/01
         Frmacc41i0_1_Save
         If strControlButton <> MsgText(602) Then
            Acc1p0Save
         End If
         If strControlButton <> MsgText(602) Then
            Combo3.AddItem Combo3
            AdodcClear
            txtA1P03 = GetSeqNo(txtA2B01, MaskEdBox2)
            SumShow
            txtA1P05.SetFocus
         End If
         strControlButton = MsgText(601)
   End Select
   KeyEnter KeyCode
End Sub

'*************************************************
'  顯示Grid資料(報廢傳票資料)
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc1p0,acc010, acc090 where a1p05=a0101(+) and a1p06=a0901(+) and a1p02='" & m_A1P02 & "' and a1p04 = '" & CompA1P04(txtA2B01, MaskEdBox2) & "' order by a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.Requery
   '移動到現在項次
   If Adodc1.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc1.Recordset.Fields("a1p22").Value) = False Then
         strAccNo = Adodc1.Recordset.Fields("a1p22").Value
      '若修改把明細全剪掉造成acc1p0沒資料,而產生新的傳票號碼
      ElseIf strSaveConfirm <> MsgText(4) Then
         strAccNo = MsgText(601)
      End If
      Adodc1.Recordset.Find "A1P03 = " & CNULL(txtA1P03), 0, adSearchForward, 1
      If Adodc1.Recordset.EOF Then
         Adodc1.Recordset.MoveFirst
         Exit Sub
      Else
         DataGrid1.SelBookmarks.add Adodc1.Recordset.Bookmark
      End If
   Else
        If strSaveConfirm = MsgText(4) Then
        Else
            strAccNo = MsgText(601)
        End If
   End If
   
   SumShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示欄位資料(Grid資料)
'
'*************************************************
Private Sub AdodcShow()
   '項次
   txtA1P03 = "" & Adodc1.Recordset.Fields("A1P03").Value

   '會計科目
   txtA1P05 = "" & Adodc1.Recordset.Fields("A1P05").Value

   '借方金額
   txtA1P07 = "" & Adodc1.Recordset.Fields("A1P07").Value

   '貸方金額
   txtA1P08 = "" & Adodc1.Recordset.Fields("A1P08").Value

   '部門別
   txtA1P06 = "" & Adodc1.Recordset.Fields("A1P06").Value
   If txtA1P06 = MsgText(55) Then txtA1P06 = MsgText(601)

   '摘要
   Combo3.Text = "" & Adodc1.Recordset.Fields("A1P14").Value

End Sub

'*************************************************
'  計算並顯示合計資料
'
'*************************************************
Public Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(A1P07), sum(A1P08) from acc1p0 where a1p02='" & m_A1P02 & "' and a1p04 = '" & CompA1P04(txtA2B01, MaskEdBox2) & "' ", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         txtTot1 = MsgText(601)
      Else
         txtTot1 = Format(adoaccsum.Fields(0).Value, DDollar)
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         txtTot2 = MsgText(601)
      Else
         txtTot2 = Format(adoaccsum.Fields(1).Value, DDollar)
      End If
   Else
      txtTot1 = MsgText(601)
      txtTot2 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   If adoacc2b0.RecordCount <> 0 And strSaveConfirm = MsgText(601) Then
      Frmacc0000.StatusBar1.Panels(2).Text = adoacc2b0.Bookmark & MsgText(35) & adoacc2b0.RecordCount
   End If
End Sub


'*************************************************
'  儲存欄位資料(傳票資料--交易檔)
'
'*************************************************
Private Sub Acc1p0Save()
On Error GoTo Checking

   If txtA1P05 = MsgText(601) Then
      MsgBox MsgText(10) & Label5, , MsgText(5)
      strControlButton = MsgText(602)
      txtA1P05.SetFocus
      Exit Sub
   Else
      '報廢日期
      If MaskEdBox2.Text = MsgText(29) Or Val(Replace(FCDate(MaskEdBox2), "_", "")) = 0 Then
         MsgBox MsgText(52) & MsgText(46), , MsgText(5)
         strControlButton = MsgText(602)
         MaskEdBox2.SetFocus
         Exit Sub
      End If
      
      '檢查會計科目
      If PUB_CheckCompany(txtA1P05, txtA2B16) = False Then
         strControlButton = MsgText(602)
         txtA1P05.SetFocus
         Exit Sub
      End If
      '檢查金額
      Call CaculateAmt
      If Val(txtA1P07) <> 0 And Val(txtA1P08) <> 0 Then
         MsgBox MsgText(47) & MsgText(46), , MsgText(5)
         strControlButton = MsgText(602)
         txtA1P07.SetFocus
         Exit Sub
      End If
      
      'Remove by Lydia 2017/05/22
      'If Val(Txt1(3)) = 0 Then
      '   MsgBox "未折減餘額為零！", vbCritical
      '   Exit Sub
      'End If
      
      '檢查部門別
      If txtA1P06 <> MsgText(601) Then
         If ExistCheck("acc090", "a0901", txtA1P06, Label8) = False Then
            strControlButton = MsgText(602)
            txtA1P06.SetFocus
            Exit Sub
         End If
      End If
      '檢查部門&科目別
      If CheckDept(txtA1P05, txtA1P06) = False Then
         MsgBox MsgText(103), , MsgText(5)
         strControlButton = MsgText(602)
         txtA1P06.SetFocus
         Exit Sub
      End If
   End If

   '檢查民國105年起法務收入科目不可使用
   intI = PUB_AccNoEnable(txtA1P05, Val(Replace(FCDate(MaskEdBox2), "_", "")))
   If intI <> 0 Then
      strControlButton = MsgText(602)
      txtA1P05.SetFocus
      Exit Sub
   End If
   
   '檢查會計科目&部門是否正確
   intI = PUB_AccNoGood(txtA1P05, txtA1P06)
   If intI <> 0 Then
      strControlButton = MsgText(602)
      If intI = 1 Then
         txtA1P05.SetFocus
      ElseIf intI = 2 Then
         txtA1P06.SetFocus
      End If
      Exit Sub
   End If
   
   '重抓項次
   If txtA1P03 = MsgText(601) Then
      txtA1P03 = GetSeqNo(txtA2B01, MaskEdBox2)
   End If
      
   adoacc1p0.Close
   adoacc1p0.CursorLocation = adUseClient
   adoacc1p0.Open "select * from acc1p0 where a1p02='" & m_A1P02 & "' and a1p04 = '" & CompA1P04(txtA2B01, MaskEdBox2.Text) & "' and A1P03 = '" & txtA1P03 & "' ", adoTaie, adOpenDynamic, adLockBatchOptimistic

   If adoacc1p0.RecordCount = 0 Then
        adoacc1p0.AddNew
        adoacc1p0.Fields("a1p01").Value = txtA2B16
        adoacc1p0.Fields("a1p02").Value = m_A1P02
        adoacc1p0.Fields("a1p03").Value = Trim(txtA1P03)
        adoacc1p0.Fields("a1p04").Value = CompA1P04(txtA2B01, MaskEdBox2.Text)
   End If
        
   If txtA1P05 <> MsgText(601) Then
      adoacc1p0.Fields("A1P05").Value = txtA1P05
   Else
      adoacc1p0.Fields("A1P05").Value = Null
   End If
   If txtA1P06 <> MsgText(601) Then
      adoacc1p0.Fields("A1P06").Value = txtA1P06
   Else
      adoacc1p0.Fields("A1P06").Value = Null
   End If
   If txtA1P07 <> MsgText(601) Then
      adoacc1p0.Fields("A1P07").Value = Val(txtA1P07)
   Else
      adoacc1p0.Fields("A1P07").Value = 0
   End If
   If txtA1P08 <> MsgText(601) Then
      adoacc1p0.Fields("A1P08").Value = Val(txtA1P08)
   Else
      adoacc1p0.Fields("A1P08").Value = 0
   End If
   '入帳日期
   adoacc1p0.Fields("A1P18").Value = Val(FCDate(MaskEdBox2.Text))
   '摘要
   adoacc1p0.Fields("A1P14").Value = "" & PUB_RepToOneSpace(PUB_StringFilter(Combo3.Text))
   
   adoacc1p0.UpdateBatch
   AdodcRefresh
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
   If txtTot1 = txtTot2 Then
      CreDebCheck = MsgText(602)
   End If
End Function

'*************************************************
'  關閉分錄欄位輸入狀態
'*************************************************
Public Sub FormDisabled()

   txtA2B01.Enabled = True
   MaskEdBox2.Enabled = False
   txtA2B09.Enabled = False
   Command3.Enabled = True

   txtA1P05.Enabled = False
   txtA1P07.Enabled = False
   txtA1P08.Enabled = False
   txtA1P06.Enabled = False
   Combo3.Enabled = False
   Command1.Enabled = False
   Command2.Enabled = False

End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'*************************************************
Public Sub FormEnabled()

   txtA2B09.Enabled = True
   MaskEdBox2.Enabled = True
   
   '新增
   If strSaveConfirm = MsgText(3) Then
      txtA2B01.Enabled = True
   '修改
   ElseIf strSaveConfirm = MsgText(4) Then
      txtA2B01.Enabled = False
      Command3.Enabled = False
      If EditCheck = False Then '檢查傳票已過帳，不可修改
         MaskEdBox2.Enabled = False  '報廢傳票日期
         Exit Sub
      End If
   End If
   
   txtA1P05.Enabled = True
   txtA1P07.Enabled = True
   txtA1P08.Enabled = True
   txtA1P06.Enabled = True
   Combo3.Enabled = True
   Command1.Enabled = True
   Command2.Enabled = True

End Sub
Public Sub Frmacc41i0_1_Clear()
    txtA2B01 = ""
    txtA2B02 = ""
    Combo1.Text = ""
    txtA2B03 = ""
    Combo2.Text = ""
    Combo3.Text = ""
    txtA2B04 = ""
    MaskEdBox1.Mask = ""
    MaskEdBox1.Text = ""
    MaskEdBox1.Mask = DFormat
    MaskEdBox1.Tag = ""
    
    txtA2B06 = ""
    txtA2B07 = ""
    MaskEdBox2.Mask = ""
    MaskEdBox2.Text = ""
    MaskEdBox2.Mask = DFormat
    MaskEdBox2.Tag = ""
    txtA2B09 = ""
    txtA2B16 = ""
    txtA2B01.Tag = txtA2B01.Text
    txtA2B02.Tag = txtA2B02.Text
    txtA2B06.Tag = txtA2B06.Text
    txtA2B07.Tag = txtA2B07.Text
    txtA2B16.Tag = txtA2B16.Text
    MaskEdBox3.Mask = ""
    MaskEdBox3.Text = ""
    MaskEdBox3.Mask = Mid(DFormat, 1, 6)
    MaskEdBox4.Mask = ""
    MaskEdBox4.Text = ""
    MaskEdBox4.Mask = Mid(DFormat, 1, 6)
    txt1(0) = "":  txt1(1) = "":  txt1(2) = ""
    txt1(3) = "":  txt1(4) = "":  txt1(5) = ""
    
    strAccNo = ""
    m_A2B22 = ""
    
    Call ShowAccDept(False)
    AdodcRefresh
End Sub

Public Sub Frmacc41i0_1_Save()
Dim rsAD As New ADODB.Recordset

On Error GoTo Checking
   
    If strSaveConfirm = MsgText(3) Then
        '報廢傳票日期
        If Val(Replace(FCDate(MaskEdBox2), "_", "")) <> 0 Then
           adoacc2b0a.Fields("a2b19").Value = Val(FCDate(MaskEdBox2.Text))
        Else
           adoacc2b0a.Fields("a2b19").Value = Null
        End If
         
        '備註
        adoacc2b0a.Fields("a2b09").Value = PUB_RepToOneSpace(PUB_StringFilter(Trim(txtA2B09)))
        
        '修改人員,日期
        adoacc2b0a.Fields("a2b13").Value = strUserNum
        adoacc2b0a.Fields("a2b14").Value = Val(strSrvDate(2))
        adoacc2b0a.Fields("a2b15").Value = Val(Format(ServerTime, "000000"))
    
        adoacc2b0a.UpdateBatch
        adoacc2b0a.Resync
        Call Command3_Click '重整資料
    Else
        '報廢傳票日期
        If Val(Replace(FCDate(MaskEdBox2), "_", "")) <> 0 Then
           adoacc2b0.Fields("a2b19").Value = Val(FCDate(MaskEdBox2.Text))
        Else
           adoacc2b0.Fields("a2b19").Value = Null
        End If
         
        '備註
        adoacc2b0.Fields("a2b09").Value = PUB_RepToOneSpace(PUB_StringFilter(Trim(txtA2B09)))
        
        '修改人員,日期
        adoacc2b0.Fields("a2b13").Value = strUserNum
        adoacc2b0.Fields("a2b14").Value = Val(strSrvDate(2))
        adoacc2b0.Fields("a2b15").Value = Val(Format(ServerTime, "000000"))
    
        adoacc2b0.UpdateBatch
        adoacc2b0.Resync
    
    End If
    RecordShow
      
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'取得最大明細項次
Private Function GetSeqNo(strA2b01 As String, strA2B19 As String) As String
    Dim adoaccmax As New ADODB.Recordset
    
    If adoaccmax.State = adStateOpen Then
         adoaccmax.Close
    End If
    adoaccmax.CursorLocation = adUseClient
    adoaccmax.Open "select nvl(max(A1P03),0) from acc1p0 where a1p02='" & m_A1P02 & "' and a1p04 = '" & CompA1P04(strA2b01, strA2B19) & "'  ", adoTaie, adOpenStatic, adLockReadOnly

    If adoaccmax.RecordCount = 0 Then
        GetSeqNo = ZeroBeforeNo(0, 3)
    Else
        GetSeqNo = ZeroBeforeNo(Val(adoaccmax.Fields(0).Value), 3)
    End If
    adoaccmax.Close
End Function

Private Sub MaskEdBox4_Validate(Cancel As Boolean)
   If MaskEdBox4 <> MsgText(601) Then
      If Not MaskEdBox3.Text < MaskEdBox4.Text Then
         MsgBox "攤提期間範圍不正確 !", vbCritical
         MaskEdBox4.SetFocus
         Cancel = True
      End If
   End If
End Sub

Private Sub Txt1_Change(Index As Integer)
  If Index = 1 Then
     txt1(2) = A0902Query(txt1(1))
  End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
  If Index = 1 Then
     TextInverse txt1(Index)
     CloseIme
  End If
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 1 Then
     KeyAscii = UpperCase(KeyAscii)
  End If
End Sub

Private Sub txt1_LostFocus(Index As Integer)
If Index = 1 Then
   'Modified by Lydia 2017/05/22 沒有餘額,直接沖資產
   'If strSaveConfirm = MsgText(3) And Val(FCDate(MaskEdBox2.Text)) >= Pub_A2b05Begin And Val(Txt1(3)) <> 0 Then
   If strSaveConfirm = MsgText(3) And Val(FCDate(MaskEdBox2.Text)) >= Pub_A2b05Begin Then
      If txt1(1) = MsgText(601) Then
         txt1(1) = "TOT"
      End If
      If GetDeptA09(txt1(1), "04") <> MsgText(602) Then
         MsgBox "不可輸入非分攤部門 !"
         txt1(1).SetFocus
         Exit Sub
      End If
      '報廢日期(無) -> 報廢日期(有) ->有餘額，自動產生明細
      If Val(MaskEdBox2.Tag) = 0 And Val(Replace(FCDate(MaskEdBox2), "_", "")) > 0 And Adodc1.Recordset.RecordCount = 0 Then
         Call CaculateAmt
         'If Val(Txt1(3)) <> 0 Then 'Remove by Lydia 2017/05/22 沒有餘額,直接沖資產
           Call CreateAcc1p0
         'End If 'Remove by Lydia 2017/05/22
      End If
   End If
End If
End Sub

Private Sub txtA1P07_GotFocus()
   TextInverse txtA1P07
   CloseIme
End Sub

Private Sub txtA1P08_GotFocus()
   TextInverse txtA1P08
   CloseIme
End Sub

Private Sub txtA2B02_Change()
   If Val(txtA2B02) > 0 And Val(txtA2B02) < Combo1.ListCount Then
      Combo1.ListIndex = Val(txtA2B02)
   End If
End Sub

Private Sub txtA2B02_GotFocus()
    TextInverse txtA2B02
    CloseIme
End Sub

Private Sub txtA2B02_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm = MsgText(4) Then
      Exit Sub
   End If
   
   If InStr("1,2,3,4", txtA2B02) = 0 Then
       MsgBox "類別只可輸入 1 ~ 4", , MsgText(5)
       Cancel = True
       txtA2B02.SetFocus
       Exit Sub
   End If
End Sub

Private Sub txtA2B03_Change()
   If Val(txtA2B03) > 0 And Val(txtA2B03) < Combo2.ListCount Then
      Combo2.ListIndex = Val(txtA2B03)
   End If
End Sub

Private Sub txtA2B03_GotFocus()
   TextInverse txtA2B03
   CloseIme
End Sub

Private Sub txtA2B03_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm = MsgText(4) Then
      Exit Sub
   End If
   
   If InStr("1,2,3,4,5", txtA2B03) = 0 Then
       MsgBox "類別只可輸入 1 ~ 5", , MsgText(5)
       Cancel = True
       txtA2B03.SetFocus
       Exit Sub
   End If
End Sub

Private Sub txtA2B04_GotFocus()
   TextInverse txtA2B04
   CloseIme
End Sub

Private Sub txtA2B06_GotFocus()
   TextInverse txtA2B06
   CloseIme
End Sub

Private Sub txtA2B07_GotFocus()
   TextInverse txtA2B07
   CloseIme
End Sub

'計算各項金額
Private Sub CaculateAmt()
Dim rsAD As New ADODB.Recordset
Dim intA As Integer

    '未折減餘額 = 取得原價－(依固定傳票產生的Acc1p0總額+首筆折舊傳票金額)
    strExc(1) = "0"
    '固定傳票產生的Acc1p0總額
    If m_A2B17 <> "" Then
       'Modified by Lydia 2017/05/24 直接抓固定傳票的餘額
       'strExc(1) = PUB_SumA1PtoU(txtA2B16, txtA2B17, , , "6126")
       strSql = "SELECT AXD14 FROM ACC0D1 WHERE AXD01='" & txtA2B16 & "' AND AXD02=" & m_A2B17
       intA = 1
       Set rsAD = ClsLawReadRstMsg(intA, strSql)
       If intA = 1 Then
          strExc(1) = "" & rsAD.Fields("AXD14")
       End If
       'end 2017/05/24
    End If
   '抓首筆折舊傳票金額(舊資料)
   strExc(2) = "0"
   If m_A2B22 <> "" Then
      strSql = "select sum(ax206) tot from acc021 where ax201='" & txtA2B16 & "' and ax202='" & m_A2B22 & "' and substr(ax205,1,4)='6126' "
      intA = 1
      Set rsAD = ClsLawReadRstMsg(intA, strSql)
      If intA = 1 Then
         strExc(2) = rsAD.Fields("tot")
      End If
   Else '首筆折舊傳票金額ACC1P0
       strSql = "select nvl(sum(a1p07),0) amt1 from acc1p0 where a1p01='" & txtA2B16 & "' and a1p02='M' and a1p04='" & CompA1P04(txtA2B01, MaskEdBox1) & "' "
       intA = 1
       Set rsAD = ClsLawReadRstMsg(intA, strSql)
       If intA = 1 Then
          strExc(2) = rsAD.Fields("amt1")
       End If
       'Added by Lydia 2019/09/27 首筆傳票的折舊部門帶入報廢部門
       If Val(strExc(2)) > 0 And txt1(1).Text = "" And txt1(1).Visible = True And txt1(1).Enabled = True Then
            strSql = "select a1p06 from acc1p0 where a1p05 like '6%' and a1p01='" & txtA2B16 & "' and a1p02='M' and a1p04='" & CompA1P04(txtA2B01, MaskEdBox1) & "' "
            intA = 1
            Set rsAD = ClsLawReadRstMsg(intA, strSql)
            If intA = 1 Then
               txt1(1).Text = "" & rsAD.Fields("a1p06")
            End If
       End If
       'end 2019/09/27
   End If
   
   'Modified by Lydia 2017/05/26 未折減餘額若有固定傳票,則用AXD14做餘額
   'Txt1(3) = Val(txtA2B06) - Val(strExc(1)) - Val(strExc(2))
   If m_A2B17 <> "" Then
       txt1(3) = Val(strExc(1))
   Else
       '無固定傳票,並且已過攤提期間,餘額設為零
       If Val(Replace(FCDate(MaskEdBox4 & "/01"), "_", "")) <= strSrvDate(2) Then
           txt1(3) = "0"
       Else
           txt1(3) = Val(txtA2B06) - Val(strExc(1)) - Val(strExc(2))
       End If
   End If
   'end 2017/05/26

End Sub

Private Sub txtA2B09_GotFocus()
   TextInverse txtA2B09
   OpenIme
End Sub

Private Sub txtA1P03_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         txtA1P05.SetFocus
         Exit Sub
   End Select
End Sub

Private Sub txtA1P05_Change()
   If txtA1P05 = MsgText(601) Then
      Exit Sub
   End If
   txt1(4) = A0102Query(txtA1P05)
End Sub

Private Sub txtA1P05_GotFocus()
   TextInverse txtA1P05
   CloseIme
End Sub

Private Sub txtA1P05_Validate(Cancel As Boolean)
   If strSaveConfirm = MsgText(601) Then
        Exit Sub
   End If
   If txtA1P05 <> MsgText(601) Then
      If PUB_CheckCompany(txtA1P05, txtA2B16) = False Then
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub txtA1P08_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub txtA2B01_GotFocus()
   TextInverse txtA2B01
   CloseIme
End Sub
 'Added by Lydia 2021/03/16 轉換全形數字變成半形數字
Private Sub txtA2B01_LostFocus()
   If txtA2B01 <> "" Then
       txtA2B01 = PUB_ChgNumeralStyle(txtA2B01)
   End If
End Sub

Private Sub txtA2B01_Validate(Cancel As Boolean)
   
   If strSaveConfirm = MsgText(3) Then
      adoacc2b0a.Close
      adoacc2b0a.CursorLocation = adUseClient
      adoacc2b0a.Open "select * from acc2b0 where nvl(a2b19,0) = 0 and a2b01 = '" & txtA2B01 & "' order by a2b01", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If adoacc2b0a.RecordCount <> 0 Then
         Adoacc2b0aRefresh
         AdodcRefresh
         SumShow
      Else
         MsgBox MsgText(28), , MsgText(5)
         AdodcClear
         AdodcRefresh
         Cancel = True
      End If
   End If
End Sub

Private Sub txtA2B16_Change()
   If txtA2B16 = MsgText(601) Then
      Exit Sub
   End If
   'Add by Amy 2020/04/14 +只能輸作帳公司別
   If InStr(GetBookKeepCmp, txtA2B16) = 0 Then
     txt1(0) = ""
     Exit Sub
   End If
   'end 2020/04/14
   txt1(0) = A0802Query(txtA2B16)
End Sub

Private Sub txtA2B16_GotFocus()
   TextInverse txtA2B16
   CloseIme
End Sub

Private Sub txtA2B16_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtA2B16_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm = MsgText(4) Then
      Exit Sub
   End If
   'Modify by Amy 2020/04/14
   'If txtA2B16 <> "1" And txtA2B16 <> "J" Then
   If InStr(GetBookKeepCmp, txtA2B16) = 0 Then
         'MsgBox "公司別只可輸入 1 或 J", , MsgText(5)
         MsgBox Label4 & MsgText(63), , MsgText(5)
         Cancel = True
         txtA2B16.SetFocus
         Exit Sub
   End If
   'end 2020/04/14
End Sub

Private Sub txtA1P06_Change()
   txt1(5) = A0902Query(txtA1P06)
End Sub

Private Sub txtA1P06_GotFocus()
   TextInverse txtA1P06
   CloseIme
End Sub

Private Sub txtA1P06_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtA1P06_Validate(Cancel As Boolean)
   If CheckDept(txtA1P05, txtA1P06) = False Then
      MsgBox MsgText(103), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
   If txtA1P06 <> MsgText(601) Then
      If ExistCheck("acc090", "a0901", txtA1P06, Label8) = False Then
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
    If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
       Exit Sub
    End If
    If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
       MsgBox Label16(7) & MsgText(63), , MsgText(5)
       MaskEdBox2.SetFocus
       Cancel = True
       Exit Sub
    End If
    
    '新增
    If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
       If Val(Replace(FCDate(MaskEdBox2), "_", "")) > 0 And Val(Replace(FCDate(MaskEdBox2), "_", "")) < Val(FCDate(MaskEdBox1.Text)) Then
          MsgBox "報廢傳票日期不可小於取得日期!"
          MaskEdBox2.SetFocus
          Cancel = True
          Exit Sub
       End If
       If DBDATE(MaskEdBox2.Text) <> DBDATE(MaskEdBox2.Tag) Then
        If ChkWorkData(txtA2B16, DBDATE(MaskEdBox2), strExc(1)) = False Then
            MsgBox "報廢日期" & strExc(1), , MsgText(5)
            MaskEdBox2.SetFocus
            Cancel = True
            Exit Sub
        End If
       End If
    End If
    
    '修改
    If MaskEdBox2.Enabled = True And strSaveConfirm = MsgText(4) Then
        '報廢日期(有) -> 報廢日期(無)
        If Val(MaskEdBox2.Tag) > 0 And Val(Replace(FCDate(MaskEdBox2), "_", "")) = 0 Then
           If bolA1P27 Then
               If MsgBox("已有傳票編號，確定是否取消傳票?", vbYesNo + vbDefaultButton2) = vbYes Then
                  bolDelAll = True
                  Call Command2_Click
                  bolDelAll = False
               Else
                  MaskEdBox2.Mask = MsgText(601)
                  MaskEdBox2.Text = CFDate(MaskEdBox2.Tag)
                  MaskEdBox2.Mask = DFormat
               End If
           Else
              bolDelAll = True
              Call Command2_Click
              bolDelAll = False
           End If
        End If
    End If
End Sub

Public Sub Frmacc41i0_1_First()
    If FormCheck = False Then
       Exit Sub
    End If
    If adoacc2b0.RecordCount <> 0 Then
       adoacc2b0.MoveFirst
       FormShow
       AdodcRefresh
       SumShow
       RecordShow
       FormDisabled
    End If
    AdodcClear
End Sub

Public Sub Frmacc41i0_1_Last()
    If FormCheck = False Then
         Exit Sub
    End If
    If adoacc2b0.RecordCount <> 0 Then
       adoacc2b0.MoveLast
       FormShow
       AdodcRefresh
       SumShow
       RecordShow
       FormDisabled
    End If
    AdodcClear
  End Sub
  
Public Sub Frmacc41i0_1_Next()
    If FormCheck = False Then
       Exit Sub
    End If
    If adoacc2b0.EOF = False Then
       adoacc2b0.MoveNext
       If adoacc2b0.EOF Then
          adoacc2b0.MoveLast
          MsgBox MsgText(8), , MsgText(5)
       End If
       FormShow
       AdodcRefresh
       SumShow
       RecordShow
       FormDisabled
    End If
    AdodcClear
End Sub

Public Sub Frmacc41i0_1_Previous()
    If FormCheck = False Then
        Exit Sub
    End If
    If adoacc2b0.BOF = False Then
       adoacc2b0.MovePrevious
       If adoacc2b0.BOF Then
          adoacc2b0.MoveFirst
          MsgBox MsgText(7), , MsgText(5)
       End If
       FormShow
       AdodcRefresh
       SumShow
       RecordShow
       FormDisabled
    End If
    AdodcClear
End Sub

Public Function FormCheck() As Boolean

Dim bCancel As Boolean

FormCheck = False

    If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      If Val(Replace(FCDate(MaskEdBox2), "_", "")) = 0 Then
          MsgBox "報廢傳票日期不可空白!"
          MaskEdBox2.SetFocus
          Exit Function
      End If
      If Val(Replace(FCDate(MaskEdBox2), "_", "")) > 0 And Val(Replace(FCDate(MaskEdBox2), "_", "")) < Val(FCDate(MaskEdBox1.Text)) Then
          MsgBox "報廢傳票日期不可小於取得日期!"
          MaskEdBox2.SetFocus
          Exit Function
      End If
      If DBDATE(MaskEdBox2.Text) <> DBDATE(MaskEdBox2.Tag) Then
        If ChkWorkData(txtA2B16, DBDATE(MaskEdBox2), strExc(1)) = False Then
            MsgBox "報廢日期" & strExc(1), , MsgText(5)
            MaskEdBox2.SetFocus
            Exit Function
        End If
      End If
      
      'Modified by Lydia 2017/05/22
      'If Val(Txt1(3)) > 0 And Trim(Txt1(1)) = "" Then
      'Modified by Lydia 2017/12/13 修改會出錯
      'If Trim(Txt1(1)) = "" Then
      If Trim(txt1(1)) = "" And txt1(1).Visible = True Then
         MsgBox "請輸入報廢部門!!"
         txt1(1).SetFocus
         Exit Function
      End If
      'Remove by Lydia 2017/05/22
      'If Adodc1.Recordset.RecordCount > 0 Then
        'If Val(Txt1(3)) = 0 Then
        '    MsgBox "請刪除報廢明細資料!!"
        '    Exit Function
        'End If
        'Re by Lydia 2017/05/22 檢查傳票金額與總額
        'If Val(Format(txtTot1, "#####0")) <> Val(Txt1(3)) Or Val(Format(txtTot2, "#####0")) <> Val(Txt1(3)) Then
        '    If MsgBox("未折減餘額與報廢傳票金額不一致，是否繼續?", vbYesNo + vbDefaultButton2) = vbNo Then
        '       Exit Function
        '    End If
        'End If
      'End If
      'Added by Lydia 2017/05/22
      If Adodc1.Recordset.RecordCount = 0 Then
         MsgBox "請輸入傳票明細!!"
         Exit Function
      End If
      'end 2017/05/22
        'Added by Lydia 2021/12/01 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
        If PUB_ChkUniText(Me, , True, "TextBox") = False Then
            Exit Function
        End If
        'end 2021/12/01
    End If
 
    '借貸平衡
    If CreDebCheck <> MsgText(602) Then
        MsgBox MsgText(11), , MsgText(5)
        Exit Function
    End If
    'Added by Lydia 2017/05/22 檢查傳票金額與總額
    If Val(Format(txtTot1, "#####0")) <> Val(txtA2B06) Or Val(Format(txtTot2, "#####0")) <> Val(txtA2B06) Then
        MsgBox "報廢傳票金額與取得原價不一致!!", vbCritical
        Exit Function
    End If
    'end 2017/05/22
    
FormCheck = True

End Function

'為資料一致更新acc1p0
Public Sub UpdateAcc1p0()
    Dim strUpd As String
    
On Error GoTo ChkHand
    
    '傳票已過帳，不可修改
    If EditCheck(0, False) = False Then
       Exit Sub
    End If
    
    '更新傳票
    If strSaveConfirm = MsgText(4) Then
       If strAccNo <> "" Then
          strUpd = strUpd & " ,a1p22= " & CNULL(strAccNo)
       End If
       If bolA1P27 = True Then
          strUpd = strUpd & " ,a1p27='Y' "
       End If
    End If

    If strSaveConfirm = MsgText(4) Then
        strUpd = "Update Acc1p0 set a1p04='" & CompA1P04(txtA2B01, MaskEdBox2.Text) & "' " & _
                 ",a1p18=" & Val(FCDate(MaskEdBox2.Text)) & ", a1p28=" & strSrvDate(2) & ",a1p29=" & Val(Format(ServerTime, "000000")) & strUpd & _
                 " Where a1p02='" & m_A1P02 & "' and substr(a1p04,1,6)='" & txtA2B01 & "' "
        adoTaie.Execute strUpd
    End If
    
ChkHand:
    If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox "UpdateAcc1p0 錯誤:" & Err.Description, , MsgText(5)
   strControlButton = MsgText(602)
End Sub

'依類別新增折舊傳票
Private Function CreateAcc1p0() As Boolean
Dim strFirst As String
Dim strCase1 As String, strCase2 As String

CreateAcc1p0 = False
If txtA2B01 = "" Or Val(Replace(FCDate(MaskEdBox2), "_", "")) = 0 Or txtA2B02 = "" Then Exit Function

   If (strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4)) Then
      Select Case txtA2B02
          Case "1" '交通運輸設備
              strCase1 = "1542": strCase2 = "1541"
          Case "2" '生財器具
              strCase1 = "1512": strCase2 = "1511"
          Case "3" '電腦硬體
              strCase1 = "1502": strCase2 = "1501"
          Case "4" '電腦軟體
              strCase1 = "1504": strCase2 = "1503"
      End Select

      strSql = "delete from acc1p0 Where a1p02='" & m_A1P02 & "' and substr(a1p04,1,6)='" & txtA2B01 & "' "
      adoTaie.Execute strSql
      strFirst = "INSERT INTO ACC1P0(A1P01,A1P02,A1P03,A1P04,A1P05,A1P06,A1P07,A1P08,A1P14,A1P18,A1P27) VALUES "
      
      'Added by Lydia 2017/05/22 若有餘額,借方掛其他損失7203
      '借方(總公司):備抵折舊、其他損失
      If Val(txt1(3)) <> 0 Then
         'Modified by Lydia 2019/09/27 財務目錄 報廢傳票, 在摘要後面都加上(已報廢沖銷)
         strSql = strFirst & "('" & txtA2B16 & "','" & m_A1P02 & "','001','" & CompA1P04(txtA2B01, MaskEdBox2.Text) & "','" & strCase1 & "','TOT'," & Val(txtA2B06) - Val(txt1(3)) & ",0," & CNULL(Trim(txtA2B04) & "/" & txtA2B01 & "(已報廢沖銷)") & "," & Val(FCDate(MaskEdBox2.Text)) & ",'" & IIf(bolA1P27 = True, "Y", "") & "')"
         adoTaie.Execute strSql
         strSql = strFirst & "('" & txtA2B16 & "','" & m_A1P02 & "','002','" & CompA1P04(txtA2B01, MaskEdBox2.Text) & "','7203','TOT'," & Val(txt1(3)) & ",0," & CNULL(Trim(txtA2B04) & "/" & txtA2B01 & "(已報廢沖銷)") & "," & Val(FCDate(MaskEdBox2.Text)) & ",'" & IIf(bolA1P27 = True, "Y", "") & "')"
         adoTaie.Execute strSql
      Else
      'end 2017/05/22
         'Modified by Lydia 2017/05/22 無餘額沖總額 Val(Txt1(3))=> Val(txtA2B06)
         strSql = strFirst & "('" & txtA2B16 & "','" & m_A1P02 & "','001','" & CompA1P04(txtA2B01, MaskEdBox2.Text) & "','" & strCase1 & "','TOT'," & Val(txtA2B06) & ",0," & CNULL(Trim(txtA2B04) & "/" & txtA2B01 & "(已報廢沖銷)") & "," & Val(FCDate(MaskEdBox2.Text)) & ",'" & IIf(bolA1P27 = True, "Y", "") & "')"
         adoTaie.Execute strSql
      End If 'end 2017/05/22
      
      'Modified by Lydia 2017/05/22 貸方沖總額
      'strSql = strFirst & "('" & txtA2B16 & "','" & m_A1P02 & "','002','" & CompA1P04(txtA2B01, MaskEdBox2.Text) & "','" & strCase2 & "','" & IIf(Txt1(1) <> "", Txt1(1), "TOT") & "',0," & Val(Txt1(3)) & "," & CNULL(Trim(txtA2B04) & "/" & txtA2B01) & "," & Val(FCDate(MaskEdBox2.Text)) & ",'" & IIf(bolA1P27 = True, "Y", "") & "')"
      strSql = strFirst & "('" & txtA2B16 & "','" & m_A1P02 & "','" & IIf(Val(txt1(3)) <> 0, "003", "002") & "' ,'" & CompA1P04(txtA2B01, MaskEdBox2.Text) & "','" & strCase2 & "','" & IIf(txt1(1) <> "", txt1(1), "TOT") & "',0," & Val(txtA2B06) & "," & CNULL(Trim(txtA2B04) & "/" & txtA2B01 & "(已報廢沖銷)") & "," & Val(FCDate(MaskEdBox2.Text)) & ",'" & IIf(bolA1P27 = True, "Y", "") & "')"
      adoTaie.Execute strSql '貸方
      
      txtA2B02.Tag = txtA2B02.Text
      Call AdodcRefresh
      SumShow
      RecordShow
      CreateAcc1p0 = True
   End If
End Function

Private Sub ShowAccDept(ByVal bShow As Boolean)
    '非新增時,不顯示報廢部門
    Label16(9).Visible = bShow
    txt1(1).Visible = bShow
    txt1(2).Visible = bShow
    
    txt1(1).Text = ""
    txt1(2).Text = ""
End Sub

