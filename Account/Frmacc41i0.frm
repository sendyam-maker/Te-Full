VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc41i0 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000B&
   BorderStyle     =   1  '單線固定
   Caption         =   "財產目錄資料"
   ClientHeight    =   5850
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   8730
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8730
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
      Index           =   6
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   60
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtA2B17 
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
      Left            =   5160
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   975
   End
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
      Left            =   7920
      Picture         =   "Frmacc41i0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   25
      ToolTipText     =   "取消"
      Top             =   5235
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
      Left            =   7320
      Picture         =   "Frmacc41i0.frx":066A
      Style           =   1  '圖片外觀
      TabIndex        =   24
      ToolTipText     =   "清除畫面"
      Top             =   5235
      Width           =   495
   End
   Begin VB.CommandButton CmdCall 
      BackColor       =   &H00C0FFC0&
      Caption         =   "每月固定傳票資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      MaskColor       =   &H80000000&
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   90
      Width           =   1935
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
      Height          =   330
      Index           =   3
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   2010
      Width           =   1215
   End
   Begin VB.TextBox txtA2B07 
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
      Height          =   330
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   12
      Top             =   1635
      Width           =   615
   End
   Begin VB.TextBox txtA2B18 
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
      Height          =   330
      Left            =   4470
      MaxLength       =   2
      TabIndex        =   13
      Top             =   1643
      Width           =   612
   End
   Begin VB.TextBox txtA2B06 
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
      Height          =   330
      Left            =   7350
      MaxLength       =   7
      TabIndex        =   11
      Top             =   1268
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   1830
      TabIndex        =   9
      Text            =   "Combo2"
      Top             =   1253
      Width           =   1335
   End
   Begin VB.TextBox txtA2B03 
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
      Height          =   330
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1260
      Width           =   612
   End
   Begin VB.TextBox txtA2B16 
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
      Height          =   330
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   4
      Top             =   525
      Width           =   612
   End
   Begin VB.TextBox txtA2B02 
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
      Height          =   330
      Left            =   6360
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   525
      Width           =   612
   End
   Begin VB.ComboBox Combo1 
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
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   518
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Default         =   -1  'True
      Height          =   300
      Left            =   2280
      Picture         =   "Frmacc41i0.frx":0F34
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   120
      Width           =   350
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc41i0.frx":1036
      Height          =   1050
      Left            =   240
      TabIndex        =   26
      Top             =   3045
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
      Left            =   0
      Top             =   3120
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
      Left            =   240
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   27
      Top             =   4875
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
      Left            =   3360
      MaxLength       =   7
      TabIndex        =   20
      Top             =   4875
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
      Left            =   5040
      MaxLength       =   7
      TabIndex        =   21
      Top             =   4875
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
      Left            =   4860
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   4125
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
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   4125
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
      Left            =   840
      MaxLength       =   6
      TabIndex        =   19
      Top             =   4875
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   4875
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
      Left            =   6720
      MaxLength       =   3
      TabIndex        =   22
      Top             =   4875
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
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   4875
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
      Height          =   330
      Index           =   0
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   525
      Width           =   3135
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   315
      Left            =   4470
      TabIndex        =   10
      Top             =   1268
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   315
      Left            =   7350
      TabIndex        =   14
      Top             =   1643
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
      Left            =   1200
      TabIndex        =   15
      Top             =   2025
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   529
      _Version        =   393216
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
      Left            =   2280
      TabIndex        =   16
      Top             =   2025
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   529
      _Version        =   393216
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
      Left            =   7260
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   2265
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
      Left            =   6660
      MaxLength       =   7
      TabIndex        =   17
      Top             =   2265
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSForms.TextBox txtA2B04 
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   870
      Width           =   6615
      VariousPropertyBits=   -1467989989
      BackColor       =   16777215
      MaxLength       =   100
      ScrollBars      =   2
      Size            =   "11668;661"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtA2B09 
      Height          =   585
      Left            =   1200
      TabIndex        =   18
      Top             =   2370
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
      Left            =   840
      TabIndex        =   23
      Top             =   5220
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
      Caption         =   "折舊部門"
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
      Left            =   5700
      TabIndex        =   58
      Top             =   2220
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblA2B19 
      BackStyle       =   0  '透明
      Caption         =   "lblA2B19"
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
      Left            =   7440
      TabIndex        =   57
      Top             =   2048
      Width           =   1095
   End
   Begin VB.Label lblA2b0622 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "首筆折舊傳票日期"
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
      Left            =   5430
      TabIndex        =   56
      Top             =   1673
      Width           =   1815
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
      Left            =   6120
      TabIndex        =   55
      Top             =   2048
      Width           =   1095
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
      TabIndex        =   54
      Top             =   930
      Width           =   900
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "攤提固定傳票流水號"
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
      Index           =   5
      Left            =   3120
      TabIndex        =   53
      Top             =   150
      Width           =   2175
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
      Left            =   240
      TabIndex        =   52
      Top             =   5280
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
      Left            =   3120
      TabIndex        =   50
      Top             =   2048
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
      TabIndex        =   49
      Top             =   2400
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
      TabIndex        =   48
      Top             =   1673
      Width           =   900
   End
   Begin VB.Label Label17 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "每月攤提日期"
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
      Left            =   3030
      TabIndex        =   47
      Top             =   1673
      Width           =   1335
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
      Left            =   6270
      TabIndex        =   46
      Top             =   1298
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
      TabIndex        =   45
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
      TabIndex        =   44
      Top             =   1298
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
      Left            =   2040
      TabIndex        =   43
      Top             =   2048
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
      Left            =   240
      TabIndex        =   42
      Top             =   2048
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
      Left            =   240
      TabIndex        =   41
      Top             =   4635
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
      Left            =   1560
      TabIndex        =   40
      Top             =   4635
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
      Left            =   3600
      TabIndex        =   39
      Top             =   4635
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
      Left            =   5280
      TabIndex        =   38
      Top             =   4635
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
      Left            =   7200
      TabIndex        =   37
      Top             =   4635
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1245
      Left            =   120
      Top             =   4515
      Width           =   8535
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   5235
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
      Left            =   4080
      TabIndex        =   36
      Top             =   4155
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
      Left            =   3390
      TabIndex        =   31
      Top             =   1298
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
      TabIndex        =   30
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
      TabIndex        =   28
      Top             =   563
      Width           =   900
   End
End
Attribute VB_Name = "Frmacc41i0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2017/02/10 財產目錄資料
Option Explicit
Public adoacc2b0 As New ADODB.Recordset '財產目錄資料
Public adoacc1p0 As New ADODB.Recordset '首筆折舊傳票資料
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public m_Auto As Boolean '是否自動新增每月固定傳票
Dim MonAmt As Long '每月折舊金額

'首筆折舊傳票的設定
Private Const m_A1P02 = "M" '傳票分錄別
Dim bolA1P27 As Boolean  '是否更新-首筆折舊傳票
Dim strAccNo As String   '傳票號碼
Dim FstAmt As Long '金額
Dim m_A2B22 As String '首筆折舊傳票號碼(財產目錄)

Public Sub cmdCall_Click()
Dim bolAdd As Boolean
   If strSaveConfirm = MsgText(601) And txtA2B16 <> "" Then
      If CheckUse("Frmacc4170", strExec) = False Then
          Exit Sub
      End If
      
      '已有每月固定傳票
      If txtA2B17 <> "" Then
         strCompanyNo = txtA2B16
         strItemNo = txtA2B17
        '新增每月固定傳票
        If m_Auto = True Then bolAdd = True
      End If
      
      If bolAdd Or strItemNo <> "" Then
         Me.Hide
         'Call Frmacc4170.SetFmForm(Me, txtA2B01) 'Remove  by Lydia 2021/12/22
         If bolAdd Then
            Call Frmacc4170.SetFmForm(Me, txtA2B01, True) 'Added by Lydia 2021/12/ 22 跳過一次KeyF9檢查; 因為從frmacc41i0新增時自動呼叫frmacc4170會對frmacc4170再執行一次F9
            Frmacc4170.Show    '開啟表單
            Frmacc4170.Text1.SetFocus '為了能觸發Form_Active
            Call KeyEnter(vbKeyF3)  '按F3修改
         Else
            Call Frmacc4170.SetFmForm(Me, txtA2B01) 'Added by Lydia 2021/12/22
            Frmacc4170.Show    '開啟表單
         End If
      End If
   End If

End Sub

Private Sub Combo1_LostFocus()
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      If Combo1.ListIndex < 1 Then
         MsgBox "請選擇正確類別！"
         Combo1.SetFocus
         Exit Sub
      End If
      txtA2B02.Text = Combo1.ListIndex
   End If
End Sub

Private Sub Combo2_LostFocus()
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      If Combo2.ListIndex < 1 Then
         MsgBox "請選擇正確所在地！"
         Combo2.SetFocus
         Exit Sub
      End If
      txtA2B03.Text = Combo2.ListIndex
   End If
End Sub

Private Sub Combo3_GotFocus()
    OpenIme
End Sub

'清除畫面(新增項次)
Private Sub Command1_Click()
   AdodcClear
   txtA1P03 = GetSeqNo(txtA2B01, MaskEdBox1)
   txtA1P05.SetFocus
End Sub

'取消(刪除項次)
Private Sub Command2_Click()
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   Adodc1.Recordset.Find "A1P03 = " & CNULL(txtA1P03), 0, adSearchForward, 1
   If Adodc1.Recordset.EOF Then
      Exit Sub
   End If
   
   adoTaie.Execute "delete from acc1p0 where a1p02 = '" & m_A1P02 & "' and a1p03 = '" & txtA1P03 & "' and a1p04 = '" & CompA1P04(txtA2B01, MaskEdBox1) & "'"
   
   AdodcRefresh
   AdodcClear
   SumShow
   txtA1P03 = GetSeqNo(txtA2B01, MaskEdBox1)  '重抓項次
   
   If adoacc1p0.RecordCount = 0 Then
      StatusClear
   Else
      RecordShow
   End If

End Sub

Private Sub Command3_Click()
   If txtA2B01 = MsgText(601) Or strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Exit Sub
   End If
   'Added by Lydia 2021/03/16 轉換全形數字變成半形數字
   If txtA2B01 <> "" Then
       txtA2B01 = PUB_ChgNumeralStyle(txtA2B01)
   End If
   
   adoacc2b0.Find "a2b01 = '" & txtA2B01 & "'", 0, adSearchForward, 1
   If adoacc2b0.EOF = False Then
      FormShow
      AdodcRefresh
      SumShow
      RecordShow
      FormDisabled
   Else
      MsgBox MsgText(33), , MsgText(5)
      adoacc2b0.MoveFirst
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
   If KeyCode = vbKeyF3 Then '修改

   End If
   If KeyCode = vbKeyF5 Then '刪除
      If EditCheck = False Then Exit Sub
   End If
   KeyDefine KeyCode
End Sub

'檢查是否可修改明細項(首筆折舊傳票)
'iAct:0=修改,1=刪除
Public Function EditCheck(Optional iAct As Integer = 0, Optional bolMsg As Boolean = True) As Boolean
Dim inJ As Integer
Dim rsR1 As New ADODB.Recordset

bolA1P27 = False

'Added by Lydia 2021/03/16
If lblA2B19.Caption <> "" Then
     MsgBox "此筆資料已報廢，不可" & IIf(iAct = 0, "修改", "刪除") & "！"
     Exit Function
End If
'end 2021/03/16

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
   'Modified by Lydia 2021/12/01 height 5730 =>6285
   PUB_InitForm Me, 8850, 6285, strBackPicPath1
   strFormName = Name
   
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   MaskEdBox3.Mask = Mid(DFormat, 1, 6)
   MaskEdBox4.Mask = Mid(DFormat, 1, 6)
   '設定折舊部門的位置
   Label16(9).Top = Label16(3).Top: txt1(1).Top = txt1(3).Top: txt1(2).Top = txt1(3).Top
      
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
   
   'Added by Lydia 2017/05/24 預設顯示首筆傳票日期
   lblA2b0622.Caption = "首筆折舊傳票日期"
   txt1(6).Visible = False
   txt1(6).Top = MaskEdBox2.Top
   txt1(6).Left = MaskEdBox2.Left
   'end 2017/05/24
   
   Combo3.Text = ""
   lblA2B19.Caption = ""
   OpenTable
   
   'Remove by Lydia 2017/05/18 預設不顯示
   'If adoacc2b0.RecordCount <> 0 Then
   '   Frmacc41i0_First
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
   Set Frmacc41i0 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking

   adoacc2b0.CursorLocation = adUseClient
   adoacc2b0.Open "select * from acc2b0 order by a2b01 asc ", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc1p0.CursorLocation = adUseClient
   adoacc1p0.Open "select * from acc1p0 where a1p02='" & m_A1P02 & "' and a1p04 = '" & CompA1P04(txtA2B01, MaskEdBox1) & "' order by a1p03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc1p0, acc010, acc090 where a1p05=a0101(+) and a1p06=a0901(+) and a1p02='" & m_A1P02 & "' and a1p04 = '" & CompA1P04(txtA2B01, MaskEdBox1) & "' order by a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'組合首筆折舊傳票-單據號碼
Private Function CompA1P04(ByVal SNo As String, ByVal Sdate As String) As String
    If SNo <> "" Then
       CompA1P04 = SNo & Val(FCDate(Sdate))
    End If
End Function
'*************************************************
'  重新整理財產目錄主檔
'
'*************************************************
Public Sub Acc2b0Refresh(Optional ByVal strA2b01 As String)
On Error GoTo Checking

   If adoacc2b0.State = adStateOpen Then
      adoacc2b0.Close
   End If
   adoacc2b0.CursorLocation = adUseClient
   strSql = "select * from acc2b0 order by a2b01 asc"
   adoacc2b0.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   
   If adoacc2b0.RecordCount <> 0 And strA2b01 <> "" Then
        adoacc2b0.Find "a2b01 = '" & strA2b01 & "'", 0, adSearchForward, 1
        If adoacc2b0.EOF = False Then
           FormShow
           AdodcRefresh
           SumShow
           RecordShow
        End If
   End If
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub
'*************************************************
'  顯示財產目錄資料
'
'*************************************************
Public Sub FormShow()

    '非新增時,不顯示折舊部門
   Call ShowAccDept(False)
   
   Call TxtEnabled(True)
   
   '編號
   txtA2B01 = "" & adoacc2b0.Fields("a2b01").Value
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
   '計算攤提期間
   'If Val(txtA2B07) > 0 Then Call txtA2B07_Validate(False)
   '首筆折舊傳票日期
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(adoacc2b0.Fields("a2b08").Value) Then
      MaskEdBox2.Text = MsgText(29)
   Else
      MaskEdBox2.Text = CFDate(adoacc2b0.Fields("a2b08").Value)
   End If
   MaskEdBox2.Tag = "" & adoacc2b0.Fields("a2b08").Value
   MaskEdBox2.Mask = DFormat
   '備註
   txtA2B09 = "" & adoacc2b0.Fields("a2b09").Value
   '公司別
   txtA2B16 = "" & adoacc2b0.Fields("a2b16").Value
   txtA2B16.Tag = txtA2B16.Text
   '攤提固定傳票流水號
   If IsNull(adoacc2b0.Fields("a2b17").Value) Then
      txtA2B17 = MsgText(601)
   Else
      txtA2B17 = Format("" & adoacc2b0.Fields("a2b17").Value, "000")
   End If
   '每月攤提日期
   txtA2B18 = "" & adoacc2b0.Fields("a2b18").Value
   txtA2B18.Tag = txtA2B18.Text
   '報廢日期
   If IsNull(adoacc2b0.Fields("a2b19").Value) Then
      lblA2B19.Caption = MsgText(601)
   Else
      lblA2B19.Caption = CFDate(adoacc2b0.Fields("a2b19").Value)
   End If
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
   
   'Added by Lydia 2017/05/24 顯示首筆折舊傳票號碼
   txt1(6) = m_A2B22
   If m_A2B22 <> "" Then
       lblA2b0622.Caption = "首筆折舊傳票號碼"
       txt1(6).Visible = True
       MaskEdBox2.Visible = False
   Else
       lblA2b0622.Caption = "首筆折舊傳票日期"
       txt1(6).Visible = False
       MaskEdBox2.Visible = True
   End If
   'end 2017/05/24
   
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
         If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
            Exit Sub
         End If
         'Added by Lydia 2021/12/01 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
         If PUB_ChkTrackMode = False Then
             Exit Sub
         End If
         'Added by Lydia 2021/12/01 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
         If PUB_ChkUniText(Me, , True, "ComboBox") = False Then
             Exit Sub
         End If
         'end 2021/12/01
         
         Frmacc41i0_Save
         If strControlButton <> MsgText(602) Then
            Acc1p0Save
         End If
         If strControlButton <> MsgText(602) Then
            '記錄公司別
            If strSaveConfirm = MsgText(3) And txtA2B16.Tag = "" Then
               txtA2B16.Tag = txtA2B16.Text
            End If
            Combo3.AddItem Combo3
            AdodcClear
            txtA1P03 = GetSeqNo(txtA2B01, MaskEdBox1)
            SumShow
            txtA2B01.Locked = True
            txtA1P05.SetFocus
         End If
         strControlButton = MsgText(601)
   End Select
   KeyEnter KeyCode
End Sub

'*************************************************
'  顯示Grid資料(首筆折舊傳票資料)
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc1p0,acc010, acc090 where a1p05=a0101(+) and a1p06=a0901(+) and a1p02='" & m_A1P02 & "' and a1p04 = '" & CompA1P04(txtA2B01, MaskEdBox1) & "' order by a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
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
   adoaccsum.Open "select sum(A1P07), sum(A1P08) from acc1p0 where a1p02='" & m_A1P02 & "' and a1p04 = '" & CompA1P04(txtA2B01, MaskEdBox1) & "' ", adoTaie, adOpenStatic, adLockReadOnly
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
      
      If FstAmt = 0 Then
         MsgBox "首筆折舊傳票金額為零！", vbCritical
         txtA2B07.SetFocus
         Exit Sub
      End If
      
      '檢查部門別
      If txtA1P06 <> MsgText(601) Then
         If ExistCheck("acc090", "a0901", txtA1P06, Label8) = False Then
            strControlButton = MsgText(602)
            txtA1P06.SetFocus
            Exit Sub
         End If
      Else
         '未輸入部門，預設為TOT
         If Mid(txtA1P05, 1, 1) <> "6" Then
            txtA1P06.Text = "TOT"
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
   intI = PUB_AccNoEnable(txtA1P05, Val(FCDate(MaskEdBox2.Text)))
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
      txtA1P03 = GetSeqNo(txtA2B01, MaskEdBox1)
   End If
      
   adoacc1p0.Close
   adoacc1p0.CursorLocation = adUseClient
   adoacc1p0.Open "select * from acc1p0 where a1p02='" & m_A1P02 & "' and a1p04 = '" & CompA1P04(txtA2B01, MaskEdBox1.Text) & "' and A1P03 = '" & txtA1P03 & "' ", adoTaie, adOpenDynamic, adLockBatchOptimistic

   If adoacc1p0.RecordCount = 0 Then
        adoacc1p0.AddNew
        adoacc1p0.Fields("a1p01").Value = txtA2B16
        adoacc1p0.Fields("a1p02").Value = m_A1P02
        adoacc1p0.Fields("a1p03").Value = Trim(txtA1P03)
        adoacc1p0.Fields("a1p04").Value = CompA1P04(txtA2B01, MaskEdBox1.Text)
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
   adoacc1p0.Fields("A1P14").Value = "" & ChgSQL(PUB_RepToOneSpace(PUB_StringFilter(Combo3.Text)))
   
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
   
   txtA2B01.Locked = False
   'txtA2B02.Locked = True  'Mark by Lydia 2022/12/30 都改成下拉選單
   'txtA2B03.Locked = True  'Mark by Lydia 2022/12/30 都改成下拉選單
   txtA2B04.Locked = True
   MaskEdBox1.Enabled = False
   txtA2B06.Locked = True
   txtA2B07.Locked = True
   MaskEdBox2.Enabled = False
   txtA2B09.Locked = True
   txtA2B16.Locked = True
   txtA2B18.Locked = True
   Command3.Enabled = True
   cmdCall.Enabled = True
   MaskEdBox3.Enabled = False
   MaskEdBox4.Enabled = False
   
   txtA1P05.Enabled = False
   txtA1P07.Enabled = False
   txtA1P08.Enabled = False
   txtA1P06.Enabled = False
   Combo1.Enabled = False
   Combo2.Enabled = False
   Combo3.Enabled = False
   Command1.Enabled = False
   Command2.Enabled = False

End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'*************************************************
Public Sub FormEnabled()
   txtA2B01.Locked = True
   'txtA2B02.Locked = False 'Mark by Lydia 2022/12/30 都改成下拉選單
   'txtA2B03.Locked = False 'Mark by Lydia 2022/12/30 都改成下拉選單
   txtA2B04.Locked = False
   MaskEdBox1.Enabled = True
   txtA2B06.Locked = False
   txtA2B07.Locked = False
   MaskEdBox2.Enabled = True
   txtA2B09.Locked = False
   txtA2B16.Locked = False
   txtA2B18.Locked = False
   Command3.Enabled = False
   
   Call TxtEnabled(True)
           
   cmdCall.Enabled = False
   
    '新增
    If strSaveConfirm = MsgText(3) Then
        Combo1.Enabled = True
        txtA2B16.Locked = False
        Call ShowAccDept(True)
    '修改
    Else
        Combo1.Enabled = True
        txtA2B16.Locked = True
        If EditCheck = False Then '檢查傳票已過帳，不可修改
           MaskEdBox1.Enabled = False  '取得日期
           MaskEdBox2.Enabled = False  '首筆折舊傳票日期
           Call TxtEnabled(False)
           txtA1P05.Enabled = False
           txtA1P07.Enabled = False
           txtA1P08.Enabled = False
           txtA1P06.Enabled = False
           Combo1.Enabled = False
           Combo3.Enabled = False
           Command1.Enabled = False
           Command2.Enabled = False
           Exit Sub
        End If
    End If
   
   If m_A2B22 <> "" Then MaskEdBox2.Enabled = False
   txtA1P05.Enabled = True
   txtA1P07.Enabled = True
   txtA1P08.Enabled = True
   txtA1P06.Enabled = True
   Combo2.Enabled = True
   Combo3.Enabled = True
   Command1.Enabled = True
   Command2.Enabled = True
End Sub

Public Sub Frmacc41i0_Clear()
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
    txtA2B17 = ""
    txtA2B18 = ""
    lblA2B19.Caption = ""
    txtA2B02.Tag = txtA2B02.Text
    txtA2B06.Tag = txtA2B06.Text
    txtA2B07.Tag = txtA2B07.Text
    txtA2B16.Tag = txtA2B16.Text
    txtA2B18.Tag = txtA2B18.Text
    MaskEdBox3.Mask = ""
    MaskEdBox3.Text = ""
    MaskEdBox3.Mask = Mid(DFormat, 1, 6)
    MaskEdBox4.Mask = ""
    MaskEdBox4.Text = ""
    MaskEdBox4.Mask = Mid(DFormat, 1, 6)
    
    txt1(0) = "":  txt1(1) = "":  txt1(2) = ""
    txt1(3) = "":  txt1(4) = "":  txt1(5) = ""
    
    strAccNo = ""
    m_Auto = False
    FstAmt = 0
    MonAmt = 0
    m_A2B22 = ""
    
    Call ShowAccDept(False)
    Call TxtEnabled(True)
    AdodcRefresh
End Sub

Public Sub Frmacc41i0_Save()
Dim rsAD As New ADODB.Recordset

On Error GoTo Checking
   
      '新增
      If strSaveConfirm = MsgText(3) Then
         If adoacc2b0.RecordCount <> 0 Then
            adoacc2b0.Find "a2b01 = '" & txtA2B01 & "'", 0, adSearchForward, 1
            If adoacc2b0.EOF = False Then
               GoTo NextRecord
            End If
         End If
         adoacc2b0.AddNew
      End If
      
NextRecord:
      '編號
      adoacc2b0.Fields("a2b01").Value = txtA2B01
      '類別
      adoacc2b0.Fields("a2b02").Value = txtA2B02
      '所在地
      adoacc2b0.Fields("a2b03").Value = txtA2B03
      '財產名稱
      adoacc2b0.Fields("a2b04").Value = ChgSQL(Trim(txtA2B04))
      '取得日期
      adoacc2b0.Fields("a2b05").Value = Val(FCDate(MaskEdBox1.Text))
      '取得原價
      If txtA2B06 <> MsgText(601) Then
         adoacc2b0.Fields("a2b06").Value = Val(txtA2B06)
      Else
         adoacc2b0.Fields("a2b06").Value = 0
      End If
      '使用月份
      If txtA2B07 <> MsgText(601) Then
         adoacc2b0.Fields("a2b07").Value = Val(txtA2B07)
      Else
         adoacc2b0.Fields("a2b07").Value = 0
      End If
      '首筆折舊傳票日期
      adoacc2b0.Fields("a2b08").Value = Val(FCDate(MaskEdBox2.Text))
      '備註
      adoacc2b0.Fields("a2b09").Value = ChgSQL(PUB_RepToOneSpace(PUB_StringFilter(Trim(txtA2B09))))
      
      '記錄人員,日期
      If strSaveConfirm = MsgText(3) Then
         adoacc2b0.Fields("a2b10").Value = strUserNum
         adoacc2b0.Fields("a2b11").Value = Val(strSrvDate(2))
         adoacc2b0.Fields("a2b12").Value = Val(Format(ServerTime, "000000"))
      '修改人員,日期
      Else
         adoacc2b0.Fields("a2b13").Value = strUserNum
         adoacc2b0.Fields("a2b14").Value = Val(strSrvDate(2))
         adoacc2b0.Fields("a2b15").Value = Val(Format(ServerTime, "000000"))
      End If
      
      '公司別
      adoacc2b0.Fields("a2b16").Value = txtA2B16

      '每月攤提日期
      adoacc2b0.Fields("a2b18").Value = Val(txtA2B18)
      '攤提期間
      If MaskEdBox3.Text <> Mid(MsgText(29), 1, 6) Then
         adoacc2b0.Fields("a2b20").Value = Val(Mid(MaskEdBox3.Text, 1, 3) & Mid(MaskEdBox3.Text, 5, 2))
      Else
         adoacc2b0.Fields("a2b20").Value = Null
      End If
      If MaskEdBox4.Text <> Mid(MsgText(29), 1, 6) Then
         adoacc2b0.Fields("a2b21").Value = Val(Mid(MaskEdBox4.Text, 1, 3) & Mid(MaskEdBox4.Text, 5, 2))
      Else
         adoacc2b0.Fields("a2b21").Value = Null
      End If
      
      adoacc2b0.UpdateBatch
      adoacc2b0.Resync
      RecordShow
      
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'刪除-財產目錄
Public Sub Frmacc41i0_Delete()
Dim bolDel As Boolean
Dim intT As Integer
Dim rsD1 As New ADODB.Recordset

On Error GoTo Checking

      If DeleteCheck("select a2b01 from acc2b0 where a2b01 = '" & txtA2B01 & "'") = MsgText(603) Then
         Exit Sub
      End If
      If EditCheck(1) = False Then
         Exit Sub
      End If
      bolDel = True
      If Trim(txtA2B17) <> "" Then
         strSql = "select * from acc0d1 where axd01 = '" & txtA2B16 & "' and axd02= " & Val(txtA2B17)
         intT = 1
         Set rsD1 = ClsLawReadRstMsg(intT, strSql)
         If intT = 1 Then
            If MsgBox("請問是否一併刪除每月固定傳票？", vbYesNo + vbDefaultButton2) = vbNo Then
               bolDel = False
            End If
         End If
         Set rsD1 = Nothing
      End If

      adoTaie.BeginTrans
        '首筆折舊傳票
        adoTaie.Execute "delete from acc1p0 where a1p01='" & txtA2B16 & "' and a1p02 = '" & m_A1P02 & "' and a1p04 = '" & CompA1P04(txtA2B01, MaskEdBox1.Text) & "'"
        If bolDel = True Then
            '每月固定傳票(主檔)
            adoTaie.Execute "delete from acc0d1 where axd01 = '" & txtA2B16 & "' and axd02= " & Val(txtA2B17)
            '每月固定傳票(交易檔)
            adoTaie.Execute "delete from acc0d0 where a0d01 = '" & txtA2B16 & "' and a0d02= " & Val(txtA2B17)
        End If
        '財產目錄
        adoTaie.Execute "delete from acc2b0 where a2b01 = '" & txtA2B01 & "'"
      adoTaie.CommitTrans
      adoacc1p0.Requery
      adoacc2b0.Requery
      AdodcRefresh
      If adoacc2b0.RecordCount <> 0 Then
         adoacc2b0.MoveFirst
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

'取得最大明細項次
Private Function GetSeqNo(strA2b01 As String, strA2b05 As String) As String
    Dim adoaccmax As New ADODB.Recordset
    
    If adoaccmax.State = adStateOpen Then
         adoaccmax.Close
    End If
    adoaccmax.CursorLocation = adUseClient
    adoaccmax.Open "select nvl(max(A1P03),0) from acc1p0 where a1p02='" & m_A1P02 & "' and a1p04 = '" & CompA1P04(strA2b01, strA2b05) & "'  ", adoTaie, adOpenStatic, adLockReadOnly

    If adoaccmax.RecordCount = 0 Then
        GetSeqNo = ZeroBeforeNo(0, 3)
    Else
        GetSeqNo = ZeroBeforeNo(Val(adoaccmax.Fields(0).Value), 3)
    End If
    adoaccmax.Close
End Function

'取得財產目錄的最大編號(BY 系統日)
Public Function GetA2b01No() As String
Dim adoaccmax As New ADODB.Recordset

On Error GoTo TransErr
   
    If adoaccmax.State = adStateOpen Then
         adoaccmax.Close
    End If
    adoaccmax.CursorLocation = adUseClient
    adoaccmax.Open "select nvl(max(a2b01),0) mno from acc2b0 where substr(a2b11,1,3)='" & Mid(strSrvDate(2), 1, 3) & "' ", adoTaie, adOpenStatic, adLockReadOnly

    If adoaccmax.RecordCount <> 0 Then
       If adoaccmax(0) = 0 Then
          GetA2b01No = Mid(strSrvDate(2), 1, 3) & "001"
       Else
          GetA2b01No = Format(Val(adoaccmax.Fields(0).Value) + 1, "000")
       End If
    End If
    adoaccmax.Close
    
    strSql = "INSERT INTO ACC2B0(A2B01,A2B10,A2B11,A2B12) VALUES (" & Val(GetA2b01No) & ",'" & strUserNum & "'," & Val(strSrvDate(2)) & "," & Val(Format(ServerTime, "000000")) & " ) "
    adoTaie.Execute strSql
    
    Exit Function
TransErr:
   If Err.Number = -2147168237 Then
      Resume Next
   End If
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
     txt1(2) = A0902Query(txt1(1).Text)
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
   'Modified by Lydia 2019/12/03 新增時之折舊部門, 請檢查不可空白 (ex.傳票acc1p0.a1p04=3681070625')
   'If strSaveConfirm = MsgText(3) And Val(FCDate(MaskEdBox1.Text)) >= Pub_A2b05Begin And FstAmt <> 0 Then
   If strSaveConfirm = MsgText(3) Then
      If txt1(1) = MsgText(601) Then
          'Modified by Lydia 2019/12/03 新增時之折舊部門, 請檢查不可空白
         'If MsgBox("若不輸入折舊部門代號，不會自動產生明細，是否繼續輸入？", vbCritical + vbYesNo + vbDefaultButton1) = vbYes Then
         '   Txt1(1).SetFocus
         '   Exit Sub
         'Else
            MsgBox "折舊部門不可空白 !"
            txt1(1).SetFocus
            Exit Sub
         'End If 'Remove by Lydia 2019/12/03
      End If
      If GetDeptA09(txt1(1), "04") <> MsgText(602) Then
         MsgBox "不可輸入非分攤部門 !"
         txt1(1).SetFocus
         Exit Sub
      End If
      If txt1(1) = MsgText(55) Then
         MsgBox "不可輸入TOT !"
         txt1(1).SetFocus
         Exit Sub
      End If

      If Val(Replace(FCDate(MaskEdBox2), "_", "")) = 0 Then
         MsgBox "首筆折舊傳票日期不可空白！", vbCritical
         MaskEdBox2.SetFocus
         Exit Sub
      End If
      If txtA2B02 = "" Then
         MsgBox "類別不可空白！", vbCritical
         txtA2B02.SetFocus
         Exit Sub
      End If
      If CreateAcc1p0 = False Then
         txt1(1).SetFocus
         Exit Sub
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

'Added by Lydia 2021/03/16 轉換全形數字變成半形數字
Private Sub txtA2B01_LostFocus()
   If txtA2B01 <> "" Then
       txtA2B01 = PUB_ChgNumeralStyle(txtA2B01)
   End If
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
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
    If InStr("1,2,3,4", txtA2B02) = 0 Then
        MsgBox "類別只可輸入 1 ~ 4", , MsgText(5)
        Cancel = True
        txtA2B02.SetFocus
        Exit Sub
    End If
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
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
    If InStr("1,2,3,4,5", txtA2B03) = 0 Then
        MsgBox "類別只可輸入 1 ~ 5", , MsgText(5)
        txtA2B03.SetFocus
        Cancel = True
        Exit Sub
    End If
   End If
End Sub

Private Sub txtA2B04_GotFocus()
   TextInverse txtA2B04
   OpenIme
End Sub

Private Sub txtA2B04_Validate(Cancel As Boolean)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      If Trim(txtA2B04) = MsgText(601) Then
         MsgBox MsgText(10) & Label16(6), , MsgText(5)
         txtA2B04.SetFocus
         Cancel = True
         Exit Sub
      Else
         txtA2B04 = ChgSQL(PUB_StringFilter(Trim(txtA2B04.Text)))
      End If
   End If
End Sub

Private Sub txtA2B06_GotFocus()
   TextInverse txtA2B06
   CloseIme
End Sub

Private Sub txtA2B06_Validate(Cancel As Boolean)
 '首筆折舊傳票已過帳不可改原價
 If strSaveConfirm = MsgText(4) And txtA2B06.Text <> txtA2B06.Tag Then
    Call CaculateAmt
 End If
End Sub

Private Sub txtA2B07_GotFocus()
   TextInverse txtA2B07
   CloseIme
End Sub

Private Sub txtA2B07_Validate(Cancel As Boolean)
'攤提期間=取得日期的次月～取得日期的次月＋(使用月份-2);因為起始月份也算
MaskEdBox3.Text = Mid(CFDate(TransDate(CompDate(1, 1, FCDate(MaskEdBox1.Text)), 1)), 1, 6)
If Val(txtA2B07) - 2 > 0 Then
   MaskEdBox4.Text = Mid(CFDate(TransDate(CompDate(1, Val(txtA2B07) - 2, Replace(MaskEdBox3.Text & "/01", "/", "")), 1)), 1, 6)
Else
   MaskEdBox4.Text = MaskEdBox3.Text
End If

If (strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4)) And txtA2B07.Text <> txtA2B07.Tag Then
   Call CaculateAmt
End If
End Sub

'計算各項金額
Private Sub CaculateAmt()
Dim rsAD As New ADODB.Recordset
Dim intA As Integer
   '新增-補資料不建立首筆折舊傳票->金額歸零
   If Val(txtA2B06) = 0 Or Val(txtA2B07) = 0 Or (Val(Replace(FCDate(MaskEdBox1), "_", "")) < Pub_A2b05Begin) Then
      MonAmt = 0
      FstAmt = 0
   Else
      '每月折舊金額(取整數)
      MonAmt = Val(txtA2B06) \ Val(txtA2B07)
      '首筆折舊傳票金額＝取得原價－（每月折舊金額＊使用月數）＋每月折舊金額。
      FstAmt = Val(txtA2B06) - (MonAmt * Val(txtA2B07)) + MonAmt
   End If
   
   '抓首筆折舊傳票金額(舊資料)
   If m_A2B22 <> "" Then
      '若有首筆折舊傳票號碼(財產目錄),改抓傳票日期=首筆傳票日期
      strSql = "select a0205,sum(ax206) tot from acc020,acc021 where a0201='" & txtA2B16 & "' and a0202='" & m_A2B22 & "' and a0201=ax201(+) and a0202=ax202(+) and substr(ax205,1,4)='6126' group by a0205 "
      intA = 1
      Set rsAD = ClsLawReadRstMsg(intA, strSql)
      If intA = 1 Then
         If rsAD.Fields("tot") <> 0 Then
            FstAmt = rsAD.Fields("tot")
            MaskEdBox2.Mask = MsgText(601)
            MaskEdBox2.Text = CFDate(rsAD.Fields("a0205"))
            MaskEdBox2.Tag = "" & adoacc2b0.Fields("a2b08").Value
            MaskEdBox2.Mask = DFormat
         End If
      End If
   End If
   
   Set rsAD = Nothing
    '未折減餘額
    strExc(1) = ""
    If txtA2B17 <> "" Then
       'Modified by Lydia 2017/05/26 直接抓固定傳票的餘額
       'strExc(1) = PUB_SumA1PtoU(txtA2B16, txtA2B17, , , "6126")
       strSql = "SELECT AXD14 FROM ACC0D1 WHERE AXD01='" & txtA2B16 & "' AND AXD02=" & txtA2B17
       intA = 1
       Set rsAD = ClsLawReadRstMsg(intA, strSql)
       If intA = 1 Then
          strExc(1) = "" & rsAD.Fields("AXD14")
       End If
       txt1(3) = Val(strExc(1))
       'end 2017/05/24
    'Added by Lydia 2017/05/26
    Else
       '無固定傳票,並且已過攤提期間,餘額設為零
       If Val(Replace(FCDate(MaskEdBox4 & "/01"), "_", "")) <= strSrvDate(2) Then
           txt1(3) = "0"
       Else
           txt1(3) = Val(txtA2B06) - Val(strExc(1)) - FstAmt
       End If
    'end 2017/05/26
    End If
    
    'Remove by Lydia 2017/05/26 未折減餘額若有固定傳票,則用AXD14做餘額
    'Txt1(3) = Val(txtA2B06) - Val(strExc(1)) - FstAmt

End Sub
Private Sub txtA2B09_GotFocus()
   TextInverse txtA2B09
   CloseIme
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
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
    'Modify by Amy 2020/04/14
    'If txtA2B16 <> "1" And txtA2B16 <> "J" Then
    If InStr(GetBookKeepCmp, txtA2B16) = 0 Then
       'MsgBox "公司別只可輸入 1 或 J", , MsgText(5)
       MsgBox Label4 & MsgText(63), , MsgText(5)
       txtA2B16.SetFocus
       Cancel = True
       Exit Sub
    End If
    'end 2020/04/14
   End If
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

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label3 & MsgText(63), , MsgText(5)
      MaskEdBox1.SetFocus
      Cancel = True
      Exit Sub
   End If
   
   If strSaveConfirm = MsgText(3) Then
      '補資料不建立首筆折舊傳票
      If Val(Replace(FCDate(MaskEdBox1), "_", "")) < Pub_A2b05Begin Then
         MaskEdBox2.Text = MsgText(29)
      End If
      If Val(txtA2B07) > 0 Then Call txtA2B07_Validate(False)
   ElseIf strSaveConfirm = MsgText(4) Then
      If Val(Replace(FCDate(MaskEdBox1), "_", "")) < Pub_A2b05Begin And Val(Replace(FCDate(MaskEdBox2), "_", "")) > 0 And m_A2B22 = "" Then
         MsgBox Pub_A2b05Begin & "以前為補資料,不建立首筆折舊傳票!!", vbCritical
         MaskEdBox2.Text = MsgText(29)
      End If
      If Val(txtA2B07) > 0 Then Call txtA2B07_Validate(False)
   End If
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
   Cancel = False
   If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
      Exit Sub
   End If
   If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
      MsgBox Label16(8) & MsgText(63), , MsgText(5)
      MaskEdBox2.SetFocus
      Cancel = True
      Exit Sub
   End If
   
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      If Val(FCDate(MaskEdBox1.Text)) >= Pub_A2b05Begin And Val(FCDate(MaskEdBox2.Text)) < Val(FCDate(MaskEdBox1.Text)) Then
         MsgBox "首筆折舊傳票日期不可小於取得日期!"
         MaskEdBox2.SetFocus
         Cancel = True
         Exit Sub
      End If
      
      If DBDATE(MaskEdBox2.Text) <> DBDATE(MaskEdBox2.Tag) And m_A2B22 = "" Then
        If ChkWorkData(txtA2B16, DBDATE(MaskEdBox2), strExc(1)) = False Then
            MsgBox "首筆折舊傳票日期" & strExc(1), , MsgText(5)
            MaskEdBox2.SetFocus
            Cancel = True
            Exit Sub
        End If
      End If
   End If
End Sub

Private Sub txtA2B18_Validate(Cancel As Boolean)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
    If Val(txtA2B18.Text) < 1 Or Val(txtA2B18.Text) > 25 Then
       MsgBox Label17 & "限於1至25...;因2月26,27,28有可能為放假日, 無法產生傳票!", , MsgText(5)
       Cancel = True
    End If
   End If
End Sub

Public Sub Frmacc41i0_First()
    If strSaveConfirm <> MsgText(601) Then
      If FormCheck = False Then
         Exit Sub
      End If
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

Public Sub Frmacc41i0_Last()
    If strSaveConfirm <> MsgText(601) Then
      If FormCheck = False Then
         Exit Sub
      End If
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
  
Public Sub Frmacc41i0_Next()
    If strSaveConfirm <> MsgText(601) Then
      If FormCheck = False Then
         Exit Sub
      End If
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

Public Sub Frmacc41i0_Previous()
    If strSaveConfirm <> MsgText(601) Then
      If FormCheck = False Then
         Exit Sub
      End If
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

   '公司別
   If txtA2B16 = MsgText(601) Then
      MsgBox MsgText(10) & Label4, , MsgText(5)
      strControlButton = MsgText(602)
      txtA2B16.SetFocus
      Exit Function
   End If
   '類別
   If txtA2B02 = MsgText(601) Then
      MsgBox MsgText(10) & Label10, , MsgText(5)
      strControlButton = MsgText(602)
      txtA2B02.SetFocus
      Exit Function
   Else
      Call txtA2B02_Validate(bCancel)
      If bCancel = True Then
        txtA2B02.SetFocus
        Exit Function
      End If
   End If
   '所在地
   If txtA2B03 = MsgText(601) Then
      MsgBox MsgText(10) & Label14, , MsgText(5)
      strControlButton = MsgText(602)
      txtA2B03.SetFocus
      Exit Function
   Else
      Call txtA2B03_Validate(bCancel)
      If bCancel = True Then
        txtA2B03.SetFocus
        Exit Function
      End If
   End If
   
   '財產名稱
   Call txtA2B04_Validate(bCancel)
   If bCancel = True Then
      txtA2B04.SetFocus
      Exit Function
   End If

   '取得日期
   If Val(FCDate(MaskEdBox1.Text)) = 0 Then
      MsgBox MsgText(10) & Label3, , MsgText(5)
      strControlButton = MsgText(602)
      MaskEdBox1.SetFocus
      Exit Function
   End If
   
   '取得原價
   If txtA2B06 = MsgText(601) Or Val(txtA2B06) = 0 Then
      MsgBox MsgText(10) & Label16(0), , MsgText(5)
      strControlButton = MsgText(602)
      txtA2B06.SetFocus
      Exit Function
   Else
      If PUB_CheckStrNEC(txtA2B06.Text, "N") = False Then
         MsgBox MsgText(130) & Label16(0), , MsgText(5)
         strControlButton = MsgText(602)
         txtA2B06.SetFocus
         Exit Function
      End If
   End If
   
   '使用月份
   If txtA2B07 = MsgText(601) Or Val(txtA2B07) = 0 Then
      MsgBox MsgText(10) & Label16(1), , MsgText(5)
      strControlButton = MsgText(602)
      txtA2B07.SetFocus
      Exit Function
   Else
      If PUB_CheckStrNEC(txtA2B07.Text, "N") = False Then
         MsgBox MsgText(130) & Label16(1), , MsgText(5)
         strControlButton = MsgText(602)
         txtA2B07.SetFocus
         Exit Function
      End If
   End If
   
   'Added by Lydia 2017/05/22 使用月份
    Call txtA2B07_Validate(bCancel)  '避免未變更攤提期間
    If bCancel = True Then
       txtA2B07.SetFocus
       Exit Function
    End If
   'end 2017/05/22
   
   '攤提期間
   If MaskEdBox3.Text = Mid(MsgText(29), 1, 6) Then
      MsgBox MsgText(10) & Label11, , MsgText(5)
      strControlButton = MsgText(602)
      MaskEdBox3.SetFocus
      Exit Function
   End If
   If MaskEdBox4.Text = Mid(MsgText(29), 1, 6) Then
      MsgBox MsgText(10) & Label11, , MsgText(5)
      strControlButton = MsgText(602)
      MaskEdBox4.SetFocus
      Exit Function
   End If
   
   '取得日期＜判斷日期時為補資料，無首筆折舊傳票
   If Val(FCDate(MaskEdBox1.Text)) >= Pub_A2b05Begin Then
     Call MaskEdBox2_Validate(bCancel)
     If bCancel = True Then
        Exit Function
     End If
     If Val(DateDiff("m", MaskEdBox3.Text & "/01", MaskEdBox4.Text & "/01")) + 1 <> Val(txtA2B07) - 1 Then
        MsgBox "攤提期間不等於使用月份，請洽電腦中心協助 !", vbCritical
     End If
   End If
   
   If strSaveConfirm = MsgText(3) Then
      '補資料不建立首筆折舊傳票
      If Val(Replace(FCDate(MaskEdBox1), "_", "")) < Pub_A2b05Begin And m_A2B22 = "" Then
         MaskEdBox2.Text = MsgText(29)
         If adoacc1p0.RecordCount <> 0 Then
            MsgBox "請刪除折舊明細資料!!"
            Exit Function
         End If
      End If
   ElseIf strSaveConfirm = MsgText(4) Then
      If Val(Replace(FCDate(MaskEdBox1), "_", "")) < Pub_A2b05Begin And Val(Replace(FCDate(MaskEdBox2), "_", "")) > 0 And m_A2B22 = "" Then
         MsgBox Pub_A2b05Begin & "以前為補資料,不建立首筆折舊傳票!!", vbCritical
         MaskEdBox2.Text = MsgText(29)
         If adoacc1p0.RecordCount <> 0 Then
            MsgBox "請刪除折舊明細資料!!"
            Exit Function
         End If
      End If
   End If
   
   '每月攤提日期
   If txtA2B18 = MsgText(601) Or Val(txtA2B18) = 0 Then
      MsgBox MsgText(10) & Label17, , MsgText(5)
      strControlButton = MsgText(602)
      txtA2B18.SetFocus
      Exit Function
   Else
      If PUB_CheckStrNEC(txtA2B18.Text, "N") = False Then
         MsgBox MsgText(130) & Label17, , MsgText(5)
         strControlButton = MsgText(602)
         txtA2B18.SetFocus
         Exit Function
      End If
      Call txtA2B18_Validate(bCancel)
      If bCancel = True Then
         txtA2B18.SetFocus
         Exit Function
      End If
   End If

    '借貸平衡
    If CreDebCheck <> MsgText(602) Then
        MsgBox MsgText(11), , MsgText(5)
        Exit Function
    End If
    
    If (Val(Format(txtTot1, "#####0")) <> FstAmt Or Val(Format(txtTot2, "#####0")) <> FstAmt) And m_A2B22 = "" Then
        MsgBox "首筆折舊傳票金額應為" & FstAmt & "，請確認傳票明細金額!", vbCritical
        Exit Function
    End If
    
    'Added by Lydia 2021/12/01 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
    End If
    'end 2021/12/01
    
FormCheck = True

End Function

'為資料一致更新acc1p0
Public Sub UpdateAcc1p0()
    Dim strUpd As String
    
On Error GoTo ChkHand
    
    '傳票已過帳，後續不可修改
    If EditCheck(0, False) = False Then
       Exit Sub
    End If
    
    '非補資料-更新固定傳票的摘要
    If strSaveConfirm = MsgText(4) And Val(FCDate(MaskEdBox2.Text)) >= Pub_A2b05Begin And Trim(txtA2B17) <> "" Then
        strUpd = "UPDATE ACC0D0 SET A0D10=" & CNULL(ChgSQL(Trim(txtA2B04)) & "/" & txtA2B01) & " WHERE A0D01='" & txtA2B16 & "' AND A0D02='" & txtA2B17 & "' "
        adoTaie.Execute strUpd
        strUpd = ""
    End If
    
    '更新傳票的相關欄位
    If strSaveConfirm = MsgText(4) Then
       If strAccNo <> "" Then
          strUpd = strUpd & " ,a1p22= " & CNULL(strAccNo)
       End If
       If bolA1P27 = True Then
          strUpd = strUpd & " ,a1p27='Y' "
       End If
    End If
    '補資料不變更首筆折舊傳票日期
    If m_A2B22 = "" Then
       strUpd = strUpd & ", a1p18=" & Val(FCDate(MaskEdBox2.Text))
    End If
    
    If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
        strUpd = "Update Acc1p0 set a1p01='" & txtA2B16 & "', a1p04='" & CompA1P04(txtA2B01, MaskEdBox1.Text) & "', a1p14=" & CNULL(ChgSQL(Trim(txtA2B04)) & "/" & txtA2B01) & _
                  ", a1p28=" & strSrvDate(2) & ",a1p29=" & Val(Format(ServerTime, "000000")) & strUpd & _
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

   If (strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4)) And txtA2B02.Text <> txtA2B02.Tag Then
      
      GetKindCase txtA2B02, strCase1, strCase2
      
      If CheckDept(strCase1, txt1(1)) = False Then
         MsgBox "請輸入正確部門代號！", , MsgText(5)
         Exit Function
      End If
      If ExistCheck("acc090", "a0901", txt1(1), "請輸入正確部門代號！") = False Then
         Exit Function
      End If
    
      strSql = "delete from acc1p0 Where a1p02='" & m_A1P02 & "' and substr(a1p04,1,6)='" & txtA2B01 & "' "
      adoTaie.Execute strSql
      strFirst = "INSERT INTO ACC1P0(A1P01,A1P02,A1P03,A1P04,A1P05,A1P06,A1P07,A1P08,A1P14,A1P18,A1P27) VALUES "
      
      strSql = strFirst & "('" & txtA2B16 & "','" & m_A1P02 & "','001','" & CompA1P04(txtA2B01, MaskEdBox1.Text) & "','" & strCase1 & "','" & txt1(1) & "'," & FstAmt & ",0," & CNULL(ChgSQL(Trim(txtA2B04)) & "/" & txtA2B01) & "," & Val(FCDate(MaskEdBox2.Text)) & ",'" & IIf(bolA1P27 = True, "Y", "") & "')"
      adoTaie.Execute strSql '借方
      strSql = strFirst & "('" & txtA2B16 & "','" & m_A1P02 & "','002','" & CompA1P04(txtA2B01, MaskEdBox1.Text) & "','" & strCase2 & "','TOT',0," & FstAmt & "," & CNULL(ChgSQL(Trim(txtA2B04)) & "/" & txtA2B01) & "," & Val(FCDate(MaskEdBox2.Text)) & ",'" & IIf(bolA1P27 = True, "Y", "") & "')"
      adoTaie.Execute strSql '貸方:備抵折舊掛總公司
      
      txtA2B02.Tag = txtA2B02.Text
      Call AdodcRefresh
      SumShow
      RecordShow
      CreateAcc1p0 = True
   Else
      CreateAcc1p0 = True '避免無法跳離欄位
   End If
End Function

'取得借/貸方科目
Private Sub GetKindCase(ByVal aKind As String, ByRef Cno1 As String, ByRef Cno2 As String)
    Select Case aKind
        Case "1" '交通運輸設備
            Cno1 = "612602": Cno2 = "1542"
        Case "2" '生財器具
            Cno1 = "612601": Cno2 = "1512"
        Case "3" '電腦硬體
            Cno1 = "612603": Cno2 = "1502"
        Case "4" '電腦軟體
            Cno1 = "612604": Cno2 = "1504"
    End Select
End Sub

'自動新增固定傳票
Public Function Acc0dxSave() As Boolean
Dim tmpNo As String
Dim strFirst As String
Dim strCase1 As String, strCase2 As String

    tmpNo = Pub_GetDefColMaxNo("acc0d1", "axd02")
    
    GetKindCase txtA2B02, strCase1, strCase2
    '固定傳票主檔
    strSql = "INSERT INTO ACC0D1 (AXD01, AXD02, AXD03, AXD05, AXD06, AXD07, AXD11, AXD12, AXD13, AXD14) " & _
             "VALUES ('" & txtA2B16 & "'," & tmpNo & ",'" & txtA2B18 & "','" & strUserNum & "'," & CNULL(strSrvDate(2), True) & "," & CNULL(Format(ServerTime, "000000"), True) & _
             "," & CNULL(Trim(Replace(MaskEdBox3.Text, "/", "")), True) & "," & CNULL(Trim(Replace(MaskEdBox4.Text, "/", "")), True) & "," & MonAmt * (Val(txtA2B07) - 1) & "," & MonAmt * (Val(txtA2B07) - 1) & " )"
    adoTaie.Execute strSql
    '固定傳票交易檔
    strFirst = "INSERT INTO ACC0D0 (A0D01,A0D02,A0D03,A0D05,A0D06,A0D07,A0D08,A0D10) "
    
    strSql = strFirst & "VALUES ('" & txtA2B16 & "'," & tmpNo & ",'001','" & strCase1 & "'," & MonAmt & ",0,'" & txt1(1) & "','" & Trim(txtA2B04) & "/" & txtA2B01 & "') "
    adoTaie.Execute strSql '借方
    strSql = strFirst & "VALUES ('" & txtA2B16 & "'," & tmpNo & ",'002','" & strCase2 & "',0," & MonAmt & ",'TOT','" & Trim(txtA2B04) & "/" & txtA2B01 & "') "
    adoTaie.Execute strSql '貸方
    
    txtA2B17 = tmpNo
    
    '更新財產目錄
    If adoacc2b0.RecordCount <> 0 Then
       adoacc2b0.Find "a2b01 = '" & txtA2B01 & "'", 0, adSearchForward, 1
       If adoacc2b0.EOF = False Then
          adoacc2b0.Fields("a2b17").Value = Val(txtA2B17)
          adoacc2b0.UpdateBatch
          adoacc2b0.Resync
       End If
    End If
End Function

Private Sub ShowAccDept(ByVal bShow As Boolean)
    '新增時,不顯示報廢日期,改顯示折舊部門(6開頭的科目)
    Label16(7).Visible = Not bShow
    lblA2B19.Visible = Not bShow
    Label16(9).Visible = bShow
    txt1(1).Visible = bShow
    txt1(2).Visible = bShow
    
    txt1(1).Text = ""
    txt1(2).Text = ""
End Sub

'控制是否能點選
Private Sub TxtEnabled(ByVal bolType As Boolean)
   txtA2B16.Enabled = bolType      '公司別
   txtA2B02.Enabled = bolType      '類別
   txtA2B06.Enabled = bolType      '取得原價
   txtA2B07.Enabled = bolType      '使用月份
   txtA2B04.Enabled = bolType      '財產名稱
End Sub


