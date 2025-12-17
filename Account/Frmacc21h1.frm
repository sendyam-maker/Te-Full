VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc21h1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "請款單內容輸入"
   ClientHeight    =   5400
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8784
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   8784
   Begin VB.ComboBox Combo5 
      Height          =   260
      ItemData        =   "Frmacc21h1.frx":0000
      Left            =   4710
      List            =   "Frmacc21h1.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   63
      Top             =   4980
      Width           =   890
   End
   Begin VB.CommandButton Command4 
      Caption         =   "插入"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7470
      Picture         =   "Frmacc21h1.frx":0004
      TabIndex        =   19
      ToolTipText     =   "插入"
      Top             =   4620
      Width           =   680
   End
   Begin VB.TextBox Text25 
      BackColor       =   &H00C0FFFF&
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
      Height          =   324
      Left            =   2700
      MaxLength       =   15
      TabIndex        =   60
      Top             =   0
      Width           =   1368
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Height          =   324
      Left            =   1125
      MaxLength       =   15
      TabIndex        =   59
      Top             =   0
      Width           =   1368
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      Height          =   800
      Left            =   3180
      TabIndex        =   55
      Top             =   3720
      Width           =   1520
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
         Height          =   315
         Left            =   0
         MaxLength       =   14
         TabIndex        =   14
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label Label26 
         Caption         =   "(RMB金額不一定要輸入)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   165
         Left            =   0
         TabIndex        =   57
         Top             =   30
         Width           =   2025
      End
      Begin VB.Label Label25 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "RMB金額"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   56
         Top             =   210
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   795
      Left            =   3210
      TabIndex        =   52
      Top             =   4470
      Width           =   1490
      Begin VB.TextBox Text9 
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
         Left            =   0
         MaxLength       =   4
         TabIndex        =   15
         Top             =   495
         Width           =   690
      End
      Begin VB.TextBox Text10 
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
         Left            =   690
         MaxLength       =   1
         TabIndex        =   16
         Top             =   495
         Width           =   750
      End
      Begin VB.Label Label21 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "超過　商品數"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   7.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   -10
         TabIndex        =   54
         Top             =   30
         Width           =   670
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label22 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "商品名稱可減免(Y)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   7.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   660
         TabIndex        =   53
         Top             =   30
         Width           =   800
         WordWrap        =   -1  'True
      End
   End
   Begin VB.ComboBox Combo4 
      Height          =   300
      ItemData        =   "Frmacc21h1.frx":08CE
      Left            =   6525
      List            =   "Frmacc21h1.frx":08DE
      Style           =   2  '單純下拉式
      TabIndex        =   8
      Top             =   1380
      Width           =   2070
   End
   Begin VB.TextBox Text19 
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
      Height          =   300
      Left            =   7650
      TabIndex        =   49
      Top             =   930
      Width           =   945
   End
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
      Left            =   7065
      MaxLength       =   1
      TabIndex        =   4
      Top             =   630
      Width           =   300
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   270
      Top             =   3060
      Visible         =   0   'False
      Width           =   960
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
      Height          =   300
      Left            =   1800
      TabIndex        =   46
      Top             =   4125
      Width           =   1212
   End
   Begin VB.ComboBox Combo3 
      Height          =   260
      ItemData        =   "Frmacc21h1.frx":0910
      Left            =   6060
      List            =   "Frmacc21h1.frx":0912
      Style           =   2  '單純下拉式
      TabIndex        =   6
      Top             =   930
      Width           =   990
   End
   Begin VB.TextBox Text2 
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
      Height          =   300
      Left            =   1125
      TabIndex        =   28
      Top             =   330
      Width           =   1572
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
      Left            =   7875
      MaxLength       =   1
      TabIndex        =   1
      Top             =   30
      Width           =   300
   End
   Begin VB.TextBox Text7 
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
      Height          =   300
      Left            =   6525
      TabIndex        =   27
      Top             =   330
      Width           =   492
   End
   Begin VB.TextBox Text14 
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
      Height          =   300
      Left            =   7155
      TabIndex        =   26
      Top             =   4125
      Width           =   1212
   End
   Begin VB.TextBox Text15 
      BackColor       =   &H00C0FFFF&
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
      Height          =   315
      Left            =   80
      MaxLength       =   3
      TabIndex        =   11
      Top             =   4965
      Width           =   480
   End
   Begin VB.TextBox Text16 
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
      Left            =   560
      MaxLength       =   6
      TabIndex        =   12
      Top             =   4965
      Width           =   660
   End
   Begin VB.TextBox Text18 
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
      Left            =   5610
      MaxLength       =   14
      TabIndex        =   17
      Top             =   4965
      Width           =   1080
   End
   Begin VB.TextBox Text20 
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
      Left            =   6710
      MaxLength       =   14
      TabIndex        =   18
      Top             =   4965
      Width           =   720
   End
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
      Height          =   660
      Left            =   8180
      Picture         =   "Frmacc21h1.frx":0914
      Style           =   1  '圖片外觀
      TabIndex        =   21
      ToolTipText     =   "刪除"
      Top             =   4620
      Width           =   525
   End
   Begin VB.TextBox Text21 
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
      Height          =   300
      Left            =   7005
      TabIndex        =   24
      Top             =   330
      Width           =   852
   End
   Begin VB.TextBox Text22 
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
      Height          =   300
      Left            =   7845
      TabIndex        =   23
      Top             =   330
      Width           =   252
   End
   Begin VB.TextBox Text23 
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
      Height          =   300
      Left            =   8085
      TabIndex        =   22
      Top             =   330
      Width           =   372
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
      Height          =   300
      ItemData        =   "Frmacc21h1.frx":0F7E
      Left            =   2580
      List            =   "Frmacc21h1.frx":0F80
      TabIndex        =   13
      Top             =   4965
      Width           =   620
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
      Left            =   1125
      MaxLength       =   9
      TabIndex        =   2
      Top             =   630
      Width           =   1572
   End
   Begin VB.TextBox Text8 
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
      Left            =   3885
      MaxLength       =   9
      TabIndex        =   3
      Top             =   630
      Width           =   1572
   End
   Begin VB.CommandButton Command2 
      Caption         =   "新增"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7470
      Picture         =   "Frmacc21h1.frx":0F82
      TabIndex        =   20
      ToolTipText     =   "新增"
      Top             =   4950
      Width           =   680
   End
   Begin VB.CommandButton Command3 
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6570
      TabIndex        =   10
      Top             =   1950
      Width           =   2025
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   1125
      TabIndex        =   9
      Top             =   2010
      Width           =   5415
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc21h1.frx":184C
      Height          =   1815
      Left            =   180
      TabIndex        =   25
      Top             =   2310
      Width           =   8430
      _ExtentX        =   14880
      _ExtentY        =   3196
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "a1l02"
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
         DataField       =   "a1l04"
         Caption         =   "請款項目"
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
         DataField       =   "a1j03"
         Caption         =   "中文名稱"
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
         DataField       =   "a1l05"
         Caption         =   "請款金額(台幣)"
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
         DataField       =   "a1l07"
         Caption         =   "折扣(台幣)"
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
         DataField       =   "a1l16"
         Caption         =   "輸入幣別"
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
         DataField       =   "a1l17"
         Caption         =   "輸入幣別金額"
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
         DataField       =   "a1l18"
         Caption         =   "輸入RMB金額"
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
            ColumnWidth     =   564.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   947.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1992.189
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1044.284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   875.906
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1284.095
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   5085
      TabIndex        =   0
      Top             =   0
      Width           =   1170
      _ExtentX        =   2053
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
   Begin VB.Label Label11 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "輸入幣別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   4740
      TabIndex        =   64
      Top             =   4710
      Width           =   860
      WordWrap        =   -1  'True
   End
   Begin MSForms.TextBox Text17 
      Height          =   320
      Left            =   1230
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   4970
      Width           =   1340
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      Size            =   "2355;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text24 
      Height          =   630
      Left            =   1125
      TabIndex        =   7
      Top             =   1380
      Width           =   4380
      VariousPropertyBits=   -1466941413
      BackColor       =   16777215
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "7726;1111"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text11 
      Height          =   420
      Left            =   1125
      TabIndex        =   5
      Top             =   930
      Width           =   4380
      VariousPropertyBits=   -1466941413
      BackColor       =   16777215
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "7726;741"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text3 
      Height          =   300
      Left            =   2730
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   330
      Width           =   2715
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      Size            =   "4789;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      X1              =   2430
      X2              =   2745
      Y1              =   150
      Y2              =   150
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "注意事項"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   58
      Top             =   1380
      Width           =   840
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "列印幣別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5580
      TabIndex        =   51
      Top             =   1380
      Width           =   840
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "匯率"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7155
      TabIndex        =   50
      Top             =   983
      Width           =   420
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "是否特殊請款單        (Y:是 C:整批)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   190
      Left            =   5570
      TabIndex        =   48
      Top             =   690
      Width           =   3070
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "外幣"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1125
      TabIndex        =   47
      Top             =   4170
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "請款編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   45
      Top             =   83
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "請款日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4185
      TabIndex        =   44
      Top             =   60
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "代理人"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   43
      Top             =   383
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "是否列印申請人        (Y/N)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   190
      Left            =   6380
      TabIndex        =   42
      Top             =   90
      Width           =   2360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "幣別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5565
      TabIndex        =   41
      Top             =   983
      Width           =   420
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "本所案號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5565
      TabIndex        =   40
      Top             =   390
      Width           =   840
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "列印備註"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   39
      Top             =   983
      Width           =   840
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   38
      Top             =   4170
      Width           =   420
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "折扣後 NT$"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5925
      TabIndex        =   37
      Top             =   4170
      Width           =   1065
   End
   Begin VB.Label Label12 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "項次"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   110
      TabIndex        =   36
      Top             =   4700
      Width           =   420
   End
   Begin VB.Label Label13 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "請款項目"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   1230
      TabIndex        =   35
      Top             =   4700
      Width           =   840
   End
   Begin VB.Label Label14 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "請款金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   5730
      TabIndex        =   34
      Top             =   4700
      Width           =   840
   End
   Begin VB.Label Label16 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "折扣(%)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   6690
      TabIndex        =   33
      Top             =   4700
      Width           =   750
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   890
      Left            =   30
      Top             =   4460
      Width           =   8730
   End
   Begin VB.Image Image1 
      Height          =   140
      Left            =   60
      Top             =   4490
      Visible         =   0   'False
      Width           =   140
   End
   Begin VB.Label Label6 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "規費內容"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2580
      TabIndex        =   32
      Top             =   4500
      Width           =   480
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "列印對象"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   31
      Top             =   683
      Width           =   840
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "請款對象"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2925
      TabIndex        =   30
      Top             =   683
      Width           =   840
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "印表機"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   29
      Top             =   2070
      Width           =   630
   End
End
Attribute VB_Name = "Frmacc21h1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/08 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Text3、Text24、Text17、Text11
'Modified by Morgan 2014/8/6 改與相同案件性質整批共用,原用到 frmacc1h0 的程式都改抓本畫面或變數
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
'2005/7/8整理

'****************************************************
'* 注意：規則若有變時要檢查下列程式是否也要同步修改
'* basUpdate.PUB_UpdateA1k08
'* Frmacc21p1
'* frm06010602_3.frm
'* frm060303.frm
'* frm060304.frm
'* frm060306_6.frm
'* frm060306_7.frm
'* frm060307.frm
'* frm060319.frm
'* frm110102_2.frm
'* frm03020401_04.frm
'* frm03020404_03.frm
'****************************************************
Option Explicit

Public adoacc1k0 As New ADODB.Recordset
Public adoacc1l0 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoaccmax As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adocheck As New ADODB.Recordset
Public adoselect As New ADODB.Recordset
Public strOriDate As String
Public m_strCP10 As String '案件性質
Public IsPrintAddress  As Boolean
Public m_bolFMPnewcase As Boolean 'Added by Morgan 2013/2/20 是否為FMP新案請款(要扣安全基金)
'Added by Morgan 2014/8/15
Public m_FromForm As Form '來源表單
Public m_bolIsBatch As Boolean '是否整批請款
Public strDNoArray As Variant '整批請款單號
Public m_Discount As String
Public m_CP10List As String 'Added by Morgan 2016/11/1 該次輸入的案件性質
'end 2014/8/15
Public m_AutoSave As Boolean 'Added by Morgan 2023/8/9 批次請款是否自動存檔
'Add by Amy 2025/11/12 要更新之總收文號/結案單號/不續辦or閉卷/結案單下一程序
Public stUpdCP09 As String, stF0301 As String, stNowCP10 As String, stNotInCP10 As String, stNP07 As String

Dim strSql As String
Dim strNo As String
'Dim douAmount As Double
Dim doua1k11 As Double 'Add By Sindy 2009/09/29
Dim strAmount As String
Dim intLength As Integer
Dim intCounter As Integer
Dim douUSDollar As Double
Dim strLanguage As String
Dim strMaxNo As String
Dim strDiscount As String
Private Const intDefault As Integer = 500
Private Const intTop As Integer = 600
Dim strNewPage As String
Dim prnPrint As Printer
Dim strPrint As String
Dim strCurr As String
Dim strRemark As String
Dim intAddSpaceRow As Integer
Dim m_blnClkPrintButton As Boolean '是否有按列印按鈕
Dim m_strDisc As String '折扣
Dim m_blnAcc1l0NoData As Boolean 'ACC1L0是否無資料(即新增請款項目狀態)
Dim m_strItemNo As String '請款項目序號
Dim m_bolActivated As Boolean
Dim CP10have926 As Boolean   '2010/10/25 add by sonia X55778之核對已准專利請款要預設請款備註
Dim m_boleFiling As Boolean 'Added by Morgan 2011/3/23 商標申請案是否電子送件
Dim m_bolAfterLoad As Boolean 'Added by Morgan 2012/12/6
'Add By Sindy 2013/1/24
Dim dblInputRate As Double
Dim strA1K18 As String
Dim bolIsFMP As Boolean
Dim int_NTD As Integer 'Add By Sindy 2025/3/25
Dim dblDisc As Double
Dim dblDiscAmt As Double
'2013/1/24 End
'Dim m_bolNoDisbursements As Boolean 'Added by Morgan 2015/7/13 'Removed by Morgan 2015/11/27
Dim m_AppNo As String 'Added by Morgan 2015/8/5 申請人1
Dim m_strTM08 As String 'Added by Morgan 2015/11/19 商標種類,目前為FCT判斷申請規費是否需考慮商品數用
Dim m_bolChkDate As Boolean 'Added by Morgan 2024/1/19
Dim strFA10 As String 'Added by Morgan 2025/7/3
Dim stRemindMsg As String, stShowMsg As String  'Add by Amy 2025/11/12

'Added by Morgan 2019/11/25
Private Sub Command4_Click()
   Dim strNo As String
   If Not Adodc1.Recordset.EOF Then
      AdodcClear
      strNo = "" & Adodc1.Recordset.Fields(1).Value
      Text15 = strNo
      
      strSql = "update acc1l0 set a1l02=lpad(a1l02+1,3,'0') where a1l01='" & Text1 & "' and a1l02>='" & strNo & "'"
      cnnConnection.Execute strSql, intI
      AdodcRefresh
      adoacc1l0.ReQuery
      Do While Not Adodc1.Recordset.EOF
        If "" & Adodc1.Recordset.Fields(1).Value > strNo Then
           Exit Do
        End If
        Adodc1.Recordset.MoveNext
      Loop
      
      Text16.SetFocus
   End If
End Sub

Private Sub Form_Activate()
   Static bolF22Set As Boolean
       
   'Added by Morgan 2019/8/20 外專程序輸入時自動帶出目前資料且游標停在"請款金額"--淑華,敏莉
   If Pub_StrUserSt03 = "F22" Then
      If bolF22Set = False Then
         If adoadodc1.RecordCount > 0 Then
            DataGrid1_SelChange 0
            Text18.SetFocus
         End If
         bolF22Set = True
      End If
   End If
   'end 2019/8/20
   
   If m_AutoSave Then Unload Me 'Added by Morgan 2023/8/9
End Sub

Private Sub Form_Load()
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim ii As Integer
   
   PUB_InitForm Me, Me.Width, Me.Height
   
   PUB_SetPrinter Me.Name, Combo2, strPrint 'Modified by Morgan 2017/11/7 設定請款單印表機,原程式移除改呼叫公用函數
   
    Me.Combo1.Clear
    Me.Combo1.AddItem ""
    Me.Combo1.AddItem "1"
    Me.Combo1.AddItem "2"
    Me.Combo1.AddItem "3"
    Me.Combo1.AddItem "4"
    Me.Combo1.AddItem "5"
    Me.Combo1.AddItem "6"
    Me.Combo1.AddItem "7"
    Me.Combo1.AddItem "8"
    Me.Combo1.AddItem "9"
    Me.Combo1.AddItem "A"
    Me.Combo1.AddItem "B"
    Me.Combo1.AddItem "C"
    Me.Combo1.AddItem "D"
    Me.Combo1.AddItem "E"
    
   'Added by Morgan 2012/12/6
   '順序不可變更(a1k33=listindex+1)
   Me.Combo4.Clear
   Me.Combo4.AddItem "純台幣", 0
   Me.Combo4.AddItem "台幣+外幣合計", 1
   Me.Combo4.AddItem "純外幣", 2
   Me.Combo4.AddItem "外幣+美金合計", 3
   'Me.Combo4.ListIndex = 1 '預設 2:台幣+外幣合計
   'end 2012/12/6


   'Added by Morgan 2014/8/18
   m_bolIsBatch = False
   If Not m_FromForm Is Nothing Then
      If UCase(m_FromForm.Name) = UCase("frmacc21p0") Then
         m_bolIsBatch = True
      End If
   End If
   '整批請款用
   If m_bolIsBatch = True Then
      Line1.Visible = True
      Text25.Visible = True
      Text25.Text = strCon7
      MaskEdBox1.Enabled = False
      MaskEdBox1.BackColor = &HE0E0E0
      Text6.Enabled = False
      Text6.BackColor = &HE0E0E0
      Text8.Enabled = False
      Text8.BackColor = &HE0E0E0
   Else
      Line1.Visible = False
      Text25.Visible = False
   End If
   'end 2014/8/18

   CP10have926 = False '2010/10/25 add by sonia
    'edit by nickc 2005/06/28 修正
   'OpenTable
   If OpenTable = False Then Exit Sub
   SumShow
   'Add by Amy 2025/11/12 結案單請款資料新增完成提醒
   If stF0301 <> "" Then
      If stRemindMsg <> "" Then
         MsgBox stRemindMsg, vbExclamation, "提醒！"
      End If
      If stShowMsg <> "" Then
         MsgBox stShowMsg, vbExclamation, "警告！"
      End If
   End If
   'end 2025/11/12
   
   Text15 = GetMaxNo(strItemNo)
    'Add By Cheng 2003/02/26
    '預設未按列印按鈕
    m_blnClkPrintButton = False
    
    'Add By Cheng 2004/01/28
    '記錄原請款日期
    Me.MaskEdBox1.Tag = Me.MaskEdBox1.Text
    'End
    Me.Combo3.Tag = Me.Combo3.Text   '2009/4/23 ADD BY SONIA 記錄原請款幣別
    'add by nick 2004/11/17
    IsPrintAddress = True
    
   
   'Added by Morgan 2012/9/18
   'Removed by Morgan 2024/1/19 併入SetText19
   'If Text19.Enabled = True Then
   '   'Modified by Morgan 2022/11/8
   '   'MsgBox Text8 & " 請款匯率特別請留意！", vbExclamation
   '   If Text19.Tag = "Y" Then
   '      MsgBox "請款對象 " & Text8 & FagentQuery(Text8, 2) & " 請款匯率特別，請留意！", vbExclamation
   '   Else
   '      MsgBox "客戶 " & m_AppNo & CustomerQuery(m_AppNo, 2) & " 請款匯率特別，請留意！", vbExclamation
   '   End If
   '   'end 2022/11/8
   'End If
   'end 2024/1/19
   
   'Added by Morgan 2012/12/6
   If IsNull(adoacc1k0.Fields("a1k33")) Then
      'Modify By Sindy 2016/12/16 + , , Text21, Text22, Text23
      'Combo4.ListIndex = PUB_GetDefaultCurrPrintType(Text7, Text6, Combo3) - 1
      'Modified by Morgan 2018/4/27
      'Combo4.ListIndex = PUB_GetDefaultCurrPrintType(Text7, Text6, Combo3, , Text21, Text22, Text23) - 1
      Combo4.ListIndex = PUB_GetDefaultCurrPrintType(Text7, Text8, Combo3, , Text21, Text22, Text23, Text6) - 1
      'end 2018/4/27
   Else
      Combo4.ListIndex = Val(adoacc1k0.Fields("a1k33")) - 1
   End If
   m_bolAfterLoad = True
   'end 2012/12/6
   
   'Add By Sindy 2012/12/27
   'Modify By Sindy 2025/3/20
'   Me.Frame1.Top = 4470  'Modified by Lydia 2021/12/08 4230 => 4470
'   Me.Frame1.Left = 3370
'   Me.Frame2.Top = 4470  'Modified by Lydia 2021/12/08 4230 => 4470
'   Me.Frame2.Left = 3370
   Me.Frame2.Top = Me.Frame1.Top
   Me.Frame2.Left = Me.Frame1.Left
   '2025/3/20 END
   '2012/12/27 End
   
'   If Pub_StrUserSt03 = "M51" Then
'      Combo4.Enabled = True
'   End If
   
   'Added by Morgan 2019/11/5 駐點設定(請款項目按TAB直接跳請款金額)--敏莉
   If Not (Text7 = "FCT" Or Text7 = "S") Then
      Combo1.TabStop = False
      Text9.TabStop = False
      Text10.TabStop = False
   End If
   'end 2019/11/5
End Sub

'Removed by Morgan 2012/12/6 combo3 改單純式下拉,取消此事件控制(應不需要每輸1字母都檢查)
''2009/4/23 ADD BY SONIA 開放請款幣別
'Private Sub Combo3_Change()
'   '若更改請款幣別
'   If Me.Combo3.Text <> Me.Combo3.Tag Then
'       If Text19.Enabled = False Then 'Added by Morgan 2012/9/18
'         '重抓請款匯率
'         dblRate = PUB_GetUSXRate_1(Replace(Me.MaskEdBox1.Text, "/", ""), Me.Combo3.Text)
'         Text19 = dblRate 'Added by Morgan 2012/9/18
'       End If
'       SumShow
'       Me.MaskEdBox1.Tag = Me.MaskEdBox1.Text
'       Me.Combo3.Tag = Me.Combo3.Text
'       Label19 = Me.Combo3.Text & "$"
'       m_blnClkPrintButton = False    '2009/10/13 ADD BY SONIA否則印完改日期再印會錯
'   End If
'End Sub

Private Sub Combo3_Click()
'   '若更改請款幣別
'   If Me.Combo3.Text <> Me.Combo3.Tag Then
'       If Text19.Enabled = False Then 'Added by Morgan 2012/9/18
'         '重抓請款匯率對台幣匯率
'         dblRate = PUB_GetUSXRate_1(Replace(Me.MaskEdBox1.Text, "/", ""), Me.Combo3.Text)
'         Text19 = dblRate 'Added by Morgan 2012/9/18
'       End If
'       SumShow
'       Me.MaskEdBox1.Tag = Me.MaskEdBox1.Text
'       Me.Combo3.Tag = Me.Combo3.Text
'       Label19 = Me.Combo3.Text & "$"
'       m_blnClkPrintButton = False    '2009/10/13 ADD BY SONIA否則印完改日期再印會錯
'
'      'Added by Morgan 2012/12/6
'      '列印對象或請款幣別變更要重新預設列印幣別
'      If m_bolAfterLoad Then Combo4.ListIndex = PUB_GetDefaultCurrPrintType(Text7, Text6, Combo3) - 1
'   End If
   Call GetCurrType
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo3_Validate(Cancel As Boolean)
   If Combo3 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc1y0", "a1y01", Combo3, Label5) = False Then
      Cancel = True
      Combo3.SetFocus
   End If
   If Combo3 <> "USD" Then
      If ExistCheck("DebitNoteRate", "DNR01", Combo3, "請款" & Label5 & "匯率") = False Then
         Cancel = True
         Combo3.SetFocus
      End If
   End If
'   '若更改請款幣別
'   If Me.Combo3.Text <> Me.Combo3.Tag Then
'      If Text19.Enabled = False Then 'Added by Morgan 2012/9/18
'         '重抓請款匯率對台幣匯率
'         dblRate = PUB_GetUSXRate_1(Replace(Me.MaskEdBox1.Text, "/", ""), Me.Combo3.Text)
'         Text19 = dblRate 'Added by Morgan 2012/9/18
'      End If
'      SumShow
'      Me.MaskEdBox1.Tag = Me.MaskEdBox1.Text
'      Me.Combo3.Tag = Me.Combo3.Text
'      Label19 = Me.Combo3.Text & "$"
'      m_blnClkPrintButton = False    '2009/10/13 ADD BY SONIA否則印完改日期再印會錯
'
'      'Added by Morgan 2012/12/6
'      '列印對象或請款幣別變更要重新預設列印幣別
'      If m_bolAfterLoad Then Combo4.ListIndex = PUB_GetDefaultCurrPrintType(Text7, Text6, Combo3) - 1
'   End If
   Call GetCurrType
End Sub
'2009/4/23 END

'Modify By Sindy 2013/1/18
Private Sub GetCurrType()
   '若更改請款幣別
   If Me.Combo3.Text <> Me.Combo3.Tag Then
       If Text19.Enabled = False Then 'Added by Morgan 2012/9/18
         'Modify By Sindy 2015/2/25
         If Me.Combo3.Tag = "" And Text19 <> "" Then
            'X10400893 的請款日已用WORKSHEET改為 29, (正常104/1/19請款日之匯率應為 29.5)
            '重新進入Frmacc21h1應維持原匯率,不可重抓
         Else
         '2015/2/25 END
            '重抓請款匯率對台幣匯率
            'Modified by Morgan 2024/1/19
            'dblRate = PUB_GetUSXRate_1(Replace(Me.MaskEdBox1.Text, "/", ""), Me.Combo3.Text)
            dblRate = PUB_GetRate(Replace(Me.MaskEdBox1.Text, "/", ""), Me.Combo3.Text, Text8, Text7, m_AppNo)
            'end 2024/1/19
            Text19 = dblRate 'Added by Morgan 2012/9/18
         End If
       End If
       SumShow
       Me.MaskEdBox1.Tag = Me.MaskEdBox1.Text
       Me.Combo3.Tag = Me.Combo3.Text
       Label19 = Me.Combo3.Text & "$"
       m_blnClkPrintButton = False    '2009/10/13 ADD BY SONIA否則印完改日期再印會錯
       
      'Added by Morgan 2012/12/6
      '列印對象或請款幣別變更要重新預設列印幣別
      'Modify By Sindy 2016/12/16 + , , Text21, Text22, Text23
      'If m_bolAfterLoad And bolIsFMP = False Then Combo4.ListIndex = PUB_GetDefaultCurrPrintType(Text7, Text6, Combo3) - 1
      'Modified by Morgan 2018/4/27
      'If m_bolAfterLoad And bolIsFMP = False Then Combo4.ListIndex = PUB_GetDefaultCurrPrintType(Text7, Text6, Combo3, , Text21, Text22, Text23) - 1
      If m_bolAfterLoad And bolIsFMP = False Then Combo4.ListIndex = PUB_GetDefaultCurrPrintType(Text7, Text8, Combo3, , Text21, Text22, Text23, Text6) - 1
      'end 2018/4/27
      
      '若為台幣時則格式欄位鎖住不可修改
      If Me.Combo3.Text = "NTD" Then
         Combo4.ListIndex = 0 'Add By Sindy 2015/5/12 NTD時必須為純台幣
         Combo4.Enabled = False
      Else
         Combo4.Enabled = True
      End If
   End If
End Sub

'Added by Morgan 2012/12/6
Private Sub Combo4_Click()
   '若列印模式改變時重新計算外幣
   If Combo4.Text <> Combo4.Tag Then
      If Combo3.Text = "USD" And Combo4.ListIndex = 3 Then
         MsgBox "美金請款不可選 ""外幣+美金合計"" 格式！", vbCritical
         'Combo4.Text = Combo4.Tag
         Exit Sub
      End If
      'Add By Sindy 2013/1/18
      '同時檢查請款幣別<>NTD時不可輸入1
      If Trim(Combo3.Text) <> "NTD" And Combo4.ListIndex = 0 Then
         MsgBox "請款幣別<>NTD時,幣別格式不可點選純台幣！", vbCritical
         'Combo4.Text = Combo4.Tag
         Exit Sub
      End If
      '請款幣別<>RMB時不可輸入4
      If Trim(Combo3.Text) <> "RMB" And Combo4.ListIndex = 3 Then
         MsgBox "請款幣別<>RMB時,幣別格式不可點選外幣+美金合計！", vbCritical
         'Combo4.Text = Combo4.Tag
         Exit Sub
      End If
      '2012/12/27 End
      SumShow
      m_blnClkPrintButton = False
   End If
   Combo4.Tag = Combo4
End Sub

'Add By Sindy 2013/4/16
Private Sub Combo5_Click()
   If Combo5.Text = "USD" And Right(Trim(Text16), 2) = "99" Then
      Text13.Enabled = True '輸入RMB金額
   Else
      Text13.Enabled = False
      Text13.Text = ""
   End If
End Sub

'Add By Sindy 2012/12/27
Private Sub Combo5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2012/12/27
Private Sub Combo5_Validate(Cancel As Boolean)
   If Combo5 = MsgText(601) Then
      Exit Sub
   End If
   If Trim(Combo3.Text) = "" Then
      MsgBox "請款幣別不可空白！"
      Cancel = True
      Exit Sub
   End If
   'Modify By Sindy 2025/3/20 FMP的控管
   If bolIsFMP = True Then
   '2025/3/20 END
      '大陸官方規費及代理人服務費只能輸入RMB或USD
      If Right(Trim(Text16), 2) = "99" Or Right(Trim(Text16), 2) = "98" Then
         If Combo5 <> "RMB" And Combo5 <> "USD" Then
            MsgBox "大陸官方規費及代理人服務費只能輸入RMB或USD！"
            Cancel = True
            Exit Sub
         End If
      '本所服務費只可輸入NTD或請款幣別
      Else
         If Combo5 <> "NTD" And Combo5 <> Combo3 Then
            MsgBox "本所服務費只可輸入NTD或請款幣別！"
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub Command1_Click()
   AdodcDelete
   AdodcClear
   DataGrid1.Refresh
   
   'Added by Morgan 2019/8/8 外專程序輸入時改停在下一項目--淑華,敏莉
   'If Pub_StrUserSt03 = "F22" Then 'Removed by Morgan 2019/11/25 應該可不限制
      Do While Not Adodc1.Recordset.EOF
        If "" & Adodc1.Recordset.Fields(1).Value > m_strItemNo Then
           Exit Do
        End If
        Adodc1.Recordset.MoveNext
      Loop
      If Adodc1.Recordset.EOF Then
         Command2.Value = True '剪掉最後一筆時自動新增
      Else
         DataGrid1_SelChange 0
         Text18.SetFocus 'Added by Morgan 2019/8/19 --敏莉
      End If
   'End If 'Removed by Morgan 2019/11/25
   'end 2019/8/8
End Sub

Private Sub Command2_Click()
   AdodcClear
   Text15 = GetMaxNo(Text1)
   SumShow
   'Modify by Amy 2025/11/12 由結案單傳更新資料會錯
   If stUpdCP09 = "" Then
      Text16.SetFocus
   End If
End Sub

Private Sub Command3_Click()
   Dim iCancel As Integer, iUnloadMode As Integer
   
   'Added by Lydia 2021/12/08 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Sub
    End If
   'end 2021/12/08
   
   Command3.Enabled = False 'Add by Morgan 2006/11/15 避免重複點到
   Screen.MousePointer = vbHourglass
    'Add By Cheng 2003/02/26
    '按列印按鈕
    If m_blnClkPrintButton = False Then
        'Added by Morgan 2013/2/20
        'Modified by Morgan 2025/2/2/14
        'If ChkFMPItem() = False Then
        '    Command3.Enabled = True
        '    Screen.MousePointer = vbDefault
        '    Exit Sub
        'End If
        Form_QueryUnload iCancel, iUnloadMode
        If iCancel = True Then
            Command3.Enabled = True
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        'end 2025/2/14
        'end 2013/2/20

        'Add By Cheng 2003/02/26
        '若有按列印請款單時, 要先存檔
        Frmacc21h1_Save
        
        'Added by Morgan 2014/8/13
        If ChkMoney() = False Then
            Command3.Enabled = True
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        'end 2014/8/13
        m_blnClkPrintButton = True
    End If
    SetAcc1n0 adoacc1k0.Fields("a1k01").Value 'Added by Morgan 2018/11/27
    
    'Modify By Cheng 2003/02/26
    '列印請款單改用單筆列印的程式
    Load Frmacc2480
    With Frmacc2480
      .Visible = False 'Add by Morgan 2010/12/1
      .Text1.Text = Me.Text1.Text
      'Modified by Morgan 2014/8/18 +整批請款判斷
      '.Text2.Text = Me.Text1.Text
      If Text25 <> "" Then
         .Text2.Text = Me.Text25.Text
         .Combo2.Text = m_FromForm.cboAddrPrinter
         .txtAdd.Text = "Y"
         .m_bolOneAddr = True
      Else
         .Text2.Text = Me.Text1.Text
      End If
      'end 2014/8/18
      .Combo1.Text = Me.Combo2.Text
      'Modified by Lydia 2015/04/15 為了區別整批請款單,+C
      'Modified by Morgan 2015/11/12 取消C(還原),整批列印不能按此鈕
      If Text5 = "Y" Then
         .m_bEditDoc = True
      Else
         'Add by Morgan 2006/10/14
         '中文請款單印2份
         If .GetLanguage(Text7, Text21, Text22, Text23, Text1) = "1" Then
            'Modify By Sindy 2021/1/22 內商請款單列印1份
            If Left(Pub_StrUserSt03, 2) = "P2" Then
               .txtCopy = "1"
            Else
            '2021/1/22 END
               .txtCopy = "2"
            End If
         End If
         'end 2006/10/14
      End If
      .Command2_Click: DoEvents
      If .m_bEditDoc = False Then 'Added by Morgan 2015/7/20 編輯模式控制駐點停在Word
        Me.SetFocus
      End If
    End With
    Unload Frmacc2480
    Screen.MousePointer = vbDefault
    Command3.Enabled = True 'Add by Morgan 2006/11/15 避免重複點到
    'Add by Morgan 2009/7/21 工具列恢復,表單名稱重設(夕陽才會有作用)
    tool3_enabled
    strFormName = Me.Name
    'end 2009/7/21
    'add by sonia 2017/6/15 專利處或商標處人員操作提醒
    If Command3.Caption = "Word(&W)" And (Left(Pub_StrUserSt03, 2) = "P1" Or Left(Pub_StrUserSt03, 2) = "P2") Then
      MsgBox "特殊請款單請記得存入 Typing2\國外部特殊請款單 !", vbExclamation + vbOKOnly
    End If
    'end 2017/6/15
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   AdodcShow
End Sub

'Added by Lydia 2021/12/08
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Added by Lydia 2021/12/08 Form2.0 記錄鍵盤傳入順序
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim rsA As New ADODB.Recordset
Dim bolCancel As Boolean 'Add By Sindy 2015/3/12
   
   If Me.Text2.Text = "" Then
       MsgBox "請輸入代理人!!!", vbExclamation + vbOKOnly
       Me.Text2.SetFocus
       Cancel = True
       GoTo ExitFlag
   End If
   If Me.Text6.Text = "" Then
       MsgBox "請輸入列印對象!!!", vbExclamation + vbOKOnly
       Me.Text6.SetFocus
       Cancel = True
       GoTo ExitFlag
   End If
   If Me.Text8.Text = "" Then
       MsgBox "請輸入請款對象!!!", vbExclamation + vbOKOnly
       Me.Text8.SetFocus
       Cancel = True
       GoTo ExitFlag
   End If
   'Add By Sindy 2013/1/24
   If Combo3 = "" Then
      MsgBox "請款幣別不可空白!!!", vbExclamation + vbOKOnly
      Cancel = True
      GoTo ExitFlag
   End If
   '同時檢查請款幣別<>NTD時不可輸入1
   If Trim(Combo3.Text) <> "NTD" And Combo4.ListIndex = 0 Then
      MsgBox "請款幣別<>NTD時,幣別格式不可點選純台幣！"
      Combo4.SetFocus
      Cancel = True
      GoTo ExitFlag
   End If
   '請款幣別<>RMB時不可輸入4
   If Trim(Combo3.Text) <> "RMB" And Combo4.ListIndex = 3 Then
      MsgBox "請款幣別<>RMB時,幣別格式不可點選外幣+美金合計！"
      Combo4.SetFocus
      Cancel = True
      GoTo ExitFlag
   End If
   '2013/1/24 End
   
   'Add By Sindy 2015/3/12
   bolCancel = False
   Call MaskEdBox1_Validate(bolCancel)
   If bolCancel = True Then
      MaskEdBox1.SetFocus
      Cancel = True
      GoTo ExitFlag
   End If
   '2015/3/12 END
   
   'Add By Sindy 2025/7/21 防止匯率有修改,沒有計算到請款幣別請款金額
   bolCancel = False
   Call Text19_Validate(bolCancel)
   If bolCancel = True Then
      If Text19.Enabled = True Then Text19.SetFocus
      Cancel = True
      GoTo ExitFlag
   End If
   '2025/7/21 END
   
'   'Add By Sindy 2013/4/17
'   strExc(0) = "select * from acc1L0 where a1L01='" & Text1 & "' and substr(a1L04,-2)='98'"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      RsTemp.MoveFirst
'      '代表有代理人服務費
'      Do While Not RsTemp.EOF
'         '檢查有無本所服務費
'         If rsA.State <> adStateClosed Then rsA.Close
'         strExc(0) = "select * from acc1L0 where a1L01='" & Text1 & "' and a1L04='" & Left(Trim(RsTemp.Fields("a1L04")), Len(Trim(RsTemp.Fields("a1L04"))) - 2) & "' and a1L05 is not null"
'         rsA.CursorLocation = adUseClient
'         rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
'         If rsA.RecordCount = 0 Then
'            MsgBox "請先輸入本所服務費資料後，才能輸入代收代付！", , MsgText(5)
'            Set rsA = Nothing
'            Cancel = True
'            GoTo ExitFlag
'         End If
'         RsTemp.MoveNext
'      Loop
'   End If
'   Set rsA = Nothing
'   '2013/4/17 End
   
   'Added by Morgan 2013/2/20
   If ChkFMPItem() = False Then
       Cancel = True
       GoTo ExitFlag
   End If
   'end 2013/2/20

   'Add by Morgan 2008/5/29 從 Frmacc21h1_Save 搬來
   With Frmacc21h1
      If .Text1 = MsgText(601) Then
         Cancel = True
         MsgBox MsgText(10) & .Label1, , MsgText(5)
         .Text1.SetFocus
         GoTo ExitFlag
      Else
         If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
            Cancel = True
            MsgBox .Label2 & MsgText(52), , MsgText(5)
            .MaskEdBox1.SetFocus
            GoTo ExitFlag
         Else
            If DateCheck(.MaskEdBox1.Text) = MsgText(603) Then
               Cancel = True
               MsgBox .Label2 & MsgText(63), , MsgText(5)
               .MaskEdBox1.SetFocus
               GoTo ExitFlag
            End If
         End If
         If .Text6 <> MsgText(601) Then
            If Mid(.Text6, 1, 1) = "X" Then
               If ExistCheck("customer", "cu01", Mid(.Text6, 1, 8), .Label15) = False Then
                  Cancel = True
                  .Text6.SetFocus
                  GoTo ExitFlag
               End If
            Else
               If ExistCheck("fagent", "fa01", Mid(.Text6, 1, 8), .Label5) = False Then
                  Cancel = True
                  .Text6.SetFocus
                  GoTo ExitFlag
               End If
            End If
         End If
         If .Text8 <> MsgText(601) Then
            .Text8 = ChangeCustomerL(.Text8) 'Add by Morgan 2008/5/29 補9碼
            If Mid(.Text8, 1, 1) = "X" Then
               If ExistCheck("customer", "cu01", Mid(.Text8, 1, 8), .Label17) = False Then
                  Cancel = True
                  .Text8.SetFocus
                  GoTo ExitFlag
               End If
            Else
               If ExistCheck("fagent", "fa01", Mid(.Text8, 1, 8), .Label17) = False Then
                  Cancel = True
                  .Text8.SetFocus
                  GoTo ExitFlag
               End If
            End If
         End If
      End If
    End With
    
   'Added by Morgan 2011/11/24
   If Adodc1.Recordset.RecordCount > 0 Then
      Adodc1.Recordset.MoveFirst
      Adodc1.Recordset.Find ("a1l05= 0")
      If Not Adodc1.Recordset.EOF Then
        Cancel = True
        MsgBox "【" & Adodc1.Recordset("a1j03") & "】請款金額不可為 0 !!", vbExclamation
        AdodcShow
        Text18.SetFocus
        GoTo ExitFlag
      End If
   End If
   'end 2011/11/24
   
   'Added by Morgan 2019/5/29 因計算比例改要含打字費,檢查從 KeyDefine 移來
   '檢查翻譯費用是否超過比例
   If Text7 = "FCP" Or Text7 = "FG" Or Text7 = "P" Or Text7 = "CFP" Then
      If Adodc1.Recordset.RecordCount > 0 Then
         Adodc1.Recordset.MoveFirst
         Adodc1.Recordset.Find ("a1l04='201'")
         If Not Adodc1.Recordset.EOF Then
           If PUB_ChkTranslationFee(Text1, , False) = False Then
               Cancel = True
               AdodcShow
               Text18.SetFocus
               GoTo ExitFlag
            End If
         End If
      End If
   End If
   'end 2019/5/29
   
   'Added by Morgan 2022/3/1 整批列印後又進明細畫面(會自動重算)，導致催款與請款金額不同 Ex:X11102139
   '整批列印的請款單增加提醒
   If Text5 = "C" Then
      If adoacc1k0.Fields("a1k08").Value <> Val(Text12) Then
         MsgBox "本請款單為「整批列印」的請款單且請款金額[" & Val(Text12) & "]與原紀錄[" & adoacc1k0.Fields("a1k08").Value & "]不同，請重新執行「請款單整批列印作業」！", vbExclamation
      End If
   End If
   'end 2022/3/1
ExitFlag:
   If Cancel = True Then
      tool3_enabled
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim strSql As String, dou1n0Cnt As Double, ii As Integer
   Screen.MousePointer = vbHourglass
   
   'Added by Morgan 2015/11/27
   '避免請款項目加入後才改請款對象,此處再檢查一次
   '目前只有專利案的需求,為免無謂的檢查加判斷系統別
   If Text7 = "FCP" Or Text7 = "P" Then
      If adoacc1l0.RecordCount > 0 Then
      adoacc1l0.MoveFirst
      With adoacc1l0
      Do While Not .EOF
         If chkItem(.Fields("a1l04")) = False Then
            Cancel = 1
            tool3_enabled
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         .MoveNext
      Loop
      End With
      End If
   End If
   'end 2015/11/27
   
   'Modified by Morgan 2013/9/17 Eaton 請款單都需人工修改,存檔時自動設定為特殊
   If Left(Text8 & "000", 9) = "Y20438000" And Text5 = "" Then
      Text5.Text = "Y"
      Frmacc21h1_Save
   End If
         
   'Modify By Cheng 2003/02/26
   '若未按列印按鈕, 才要存檔
'   Frmacc21h1_Save
   If m_blnClkPrintButton = False Then
      Frmacc21h1_Save
      'Added by Morgan 2014/8/20
      If ChkMoney() = False Then
         Cancel = 1
         tool3_enabled
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
      'end 2014/8/20
      
      'add by sonia 2024/8/9 FCP及FMP案B類收文927其他翻譯且承辦人為外翻編號且相關總收文號為C類之結匯金額,OA委外翻譯請款單注意事項A1K34也要加註，收款時才會注意
      If adoacc1k0.Fields("a1k34").Value <> "" Then adoacc1k0.Fields("a1k34").Value = Replace(adoacc1k0.Fields("a1k34").Value, "此為OA委外翻譯請款；", "")
      adoselect.CursorLocation = adUseClient
      adoselect.Open "select a1p07,a1w01,a1w02,cp60,cp61 from acc1w0,caseprogress,acc1p0 where a1w01='" & adoacc1k0.Fields("a1k01").Value & "' and substr(a1w02,1,1)='B' and a1w02=cp09(+) " & _
                      "and cp01 in ('P','FCP') and cp10='927' and substr(cp14,1,1)='F' and substr(cp43,1,1)='C' and cp61||a1w02=a1p23 and a1p07>0", adoTaie, adOpenStatic, adLockReadOnly
      If adoselect.RecordCount <> 0 Then
         adoacc1k0.Fields("a1k34").Value = "此為OA委外翻譯請款；" & adoacc1k0.Fields("a1k34").Value
      End If
      adoselect.Close
      adoacc1k0.UpdateBatch
      adoacc1k0.ReQuery
      'end 2024/8/9
      
      'Modify By Sindy 2011/3/3
      'PUB_CheckDNMemo Text8 'Add by Morgan 2008/6/11 D/N備註提醒
      PUB_CheckDNMemo Text8, , Text7 'Add by Morgan 2008/6/11 D/N備註提醒
   End If
   
   'Add By Cheng 2003/02/05
   '若印表機變動, 則更新列印設定
   If Me.Combo2.Text <> Me.Combo2.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
   End If
    
   'edit by nick 2004/11/30  印地址條
   If pub_blnARPrintAddress = True Then
      pub_AddressListSN = pub_AddressListSN + 1
      'edit by nick 2004/11/10
      'PUB_AddNewAddressList strUserNum, "" & Me.Text7.Text, "" & Me.Text21.Text, "" & Me.Text22.Text, "" & Me.Text23.Text, "" & pub_AddressListSN, "0", m_strCP10
      PUB_AddNewAddressList strUserNum, "" & Me.Text7.Text, "" & Me.Text21.Text, "" & Me.Text22.Text, "" & Me.Text23.Text, "" & pub_AddressListSN, "0", IIf(UCase(Me.Text7.Text) = "FCT", IIf(m_strCP10 = "102", m_strCP10, ""), m_strCP10)
   End If
   
   'Added by Morgan 2014/9/23
   If m_bolIsBatch = True Then
      PUB_AddBatch Text1, Text25
   'Added by Morgan 2018/11/27
   Else
      SetAcc1n0 adoacc1k0.Fields("a1k01").Value
   'end 2018/11/27
   End If
   'end 2014/9/23
   
'Removed by Morgan 2018/11/27 移到上面改叫 SetAcc1n0
'   'Modify by Morgan 2010/5/17 分配點數上線
'   '先檢查ACC1N0是否有資料
'   strSql = "SELECT count(*) FROM ACC1N0 WHERE a1n01='" & strItemNo & "' "
'   intI = 1: dou1n0Cnt = 0
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      dou1n0Cnt = RsTemp.Fields(0)
'   End If
'
'   If dou1n0Cnt = 0 Then
'      '點數分配規則
'      'Modify by Morgan 2010/4/1 改規則
'      'If RunPointPei = False Then
'      If PUB_PointAutoassign(strItemNo) = True Then
'         If PUB_ChkPointOk(strItemNo) = False Then
'            Frmacc21h3.Show vbModal
'         End If
'      Else
'         Frmacc21h3.Show vbModal
'      End If
'   '若請款點數有異動, 則需進入點數分配作業
'   'ElseIf doua1k11 <> Val(Text14) Then
'   ElseIf PUB_ChkPointOk(strItemNo) = False Then
'      Frmacc21h3.Show vbModal
'   End If
'   'end 2010/5/17
'end 2018/11/27
   
   'Added by Lydia 2016/11/17 以請款對象檢查是否存在於國外固定寄催款單代理人檔(ACC225)且下次寄發日期＞系統日，若存在則顯示訊息提醒操作人員
   If PUB_ChkAcc225MsgList(Text1 & IIf(m_bolIsBatch = True, "-" & Text25, ""), Text8, Text7, Text21, Text22, Text23) Then
      'Added by Lydia 2017/01/11 用執行檔執行時,回上一表單可能會無法觸發Form_Activate
      If Not m_FromForm Is Nothing Then
         strFormName = m_FromForm.Name  '列印後會清空
      End If
      'end 2017/01/11
   End If
   
   strTrackMode = "" 'Added by Lydia 2021/12/08 Form2.0 記錄鍵盤傳入順序(清除)
   stUpdCP09 = "": stF0301 = "": stNowCP10 = "": stNotInCP10 = "": stNP07 = "" 'Add by Amy 2025/10/22
   
   tool1_enabled
   Screen.MousePointer = vbHourglass
'Modified by Morgan 2014/8/15
'   Select Case strFormLink
'      Case "Frmacc21h0"
'         Frmacc21h0.Show
'   End Select
   If Not m_FromForm Is Nothing Then
      m_FromForm.Show
   End If
'end 2014/8/18
   
   PUB_SendMailCache 'Added by Lydia 2019/07/03
   Screen.MousePointer = vbDefault
   Set Frmacc21h1 = Nothing
End Sub

''點數分配規則 Add By Sindy 2009/09/29
'Private Function RunPointPei() As Boolean
'Dim strSql As String
'Dim douF4101 As Double, douF4102 As Double, douF4103 As Double
'Dim douP1001 As Double, douP2001 As Double
'Dim dou97099 As Double, dou97098 As Double
'Dim strCP14(20) As String, douCP14(20) As Double, intCP14 As Integer
'Dim bolPcp12F As Boolean
'Dim douXXX As Double, douDotNum As Double, strTemp As String, douTemp As Double
'Dim douRuleFee As Double
'Dim i As Integer
'
'On Error GoTo ErrorHandler
'RunPointPei = True
'cnnConnection.BeginTrans
'
'   '預設值
'   bolPcp12F = False '是否有系統別為P且CP12為F字頭者
'   douRuleFee = 0: douXXX = 0: douDotNum = 0
'   douF4101 = 0: douF4102 = 0: douF4103 = 0
'   douP1001 = 0: douP2001 = 0
'   dou97099 = 0: dou97098 = 0
'   intCP14 = 0
'
'   '規費 : 項目代號含99 & (系統別=T & 項目代號=03)
'   'Modify By Sindy 2012/12/27 規費 : +項目代號含98
''   strSql = "SELECT sum(a1L05) FROM acc1L0" & _
''                   " WHERE a1L01='" & strItemNo & "' and (instr(a1L04,'99')<>0 or (a1L03='T' and a1L04='03'))"
'   strSql = "SELECT sum(a1L05) FROM acc1L0" & _
'                   " WHERE a1L01='" & strItemNo & "' and (instr(a1L04,'98')<>0 or instr(a1L04,'99')<>0 or (a1L03='T' and a1L04='03'))"
'   '2012/12/27 End
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      If Not IsNull(RsTemp.Fields(0)) Then
'         douRuleFee = Val(RsTemp.Fields(0))
'      End If
'   End If
'   douXXX = (Val(Text14) - douRuleFee) / 1000 '總分配點數
'
'   'Modify By Sindy 2013/1/24 +and instr(a1L04,'98')=0
'   strSql = "SELECT a1L03,sum((a1L05-a1L07)/1000) as a1,cp14,st15,cp12" & _
'                   " FROM acc1L0,CaseProgress,staff" & _
'                   " WHERE a1L01='" & strItemNo & "' and a1L01=cp60(+) and a1L04=cp10(+) and instr(a1L04,'99')=0 and instr(a1L04,'98')=0 and (a1L03<>'T' and a1L04<>'03') and cp14=st01(+) group by a1L03,cp14,st15,cp12"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      With RsTemp
'         .MoveFirst
'         Do While Not .EOF
'            If Trim(.Fields("a1L03")) = "P" And Left(Trim(.Fields("cp12")), 1) = "F" Then bolPcp12F = True
'
'            '以分配點數人員抓ST15判斷
'            If Trim(.Fields("st15")) = "F12" Then
'               douF4103 = douF4103 + Val(Trim(.Fields("a1")))
'            ElseIf Trim(.Fields("st15")) = "F22" Then
'               douF4102 = douF4102 + Val(Trim(.Fields("a1")))
'            ElseIf Trim(.Fields("st15")) = "P1" Then
'               douP1001 = douP1001 + Val(Trim(.Fields("a1")))
'            ElseIf Trim(.Fields("st15")) = "P2" Then
'               douP2001 = douP2001 + Val(Trim(.Fields("a1")))
'            '法務投資例外狀況者
'            ElseIf (Trim(.Fields("a1L03")) = "FCL" Or Trim(.Fields("a1L03")) = "LIN" Or Trim(.Fields("a1L03")) = "CFL") _
'                  And Trim(.Fields("cp14")) = "97009" Then
'               dou97099 = dou97099 + Val(Trim(.Fields("a1")))
'            ElseIf (Trim(.Fields("a1L03")) = "FCL" Or Trim(.Fields("a1L03")) = "LIN" Or Trim(.Fields("a1L03")) = "CFL") _
'                  And Trim(.Fields("cp14")) = "97005" Then
'               dou97098 = dou97098 + Val(Trim(.Fields("a1")))
'            '非上列例外狀況, 有CP14者
'            ElseIf Trim(.Fields("cp14")) <> "" Then
'               intCP14 = intCP14 + 1
'               strCP14(intCP14) = Trim(.Fields("cp14"))
'               douCP14(intCP14) = Val(Trim(.Fields("a1")))
'            '無CP14者
'            ElseIf Trim(.Fields("a1L03")) = "FCP" Or Trim(.Fields("a1L03")) = "FG" Then
'               douF4102 = douF4102 + Val(Trim(.Fields("a1")))
'            ElseIf Trim(.Fields("a1L03")) = "P" Or Trim(.Fields("a1L03")) = "PS" Or Trim(.Fields("a1L03")) = "CFP" Or Trim(.Fields("a1L03")) = "CPS" Then
'               douP1001 = douP1001 + Val(Trim(.Fields("a1")))
'            ElseIf Trim(.Fields("a1L03")) = "FCT" Or Trim(.Fields("a1L03")) = "CFT" Or Trim(.Fields("a1L03")) = "CFC" Or Trim(.Fields("a1L03")) = "S" Then
'               douF4103 = douF4103 + Val(Trim(.Fields("a1")))
'            ElseIf Trim(.Fields("a1L03")) = "T" Then
'               douP2001 = douP2001 + Val(Trim(.Fields("a1")))
'            ElseIf Trim(.Fields("a1L03")) = "CFL" Or Trim(.Fields("a1L03")) = "FCL" Or Trim(.Fields("a1L03")) = "LIN" Then
'               douF4101 = douF4101 + Val(Trim(.Fields("a1")))
'            Else
'               MsgBox "點數分配有誤，請洽資訊系統人員！"
'               GoTo ErrorHandler
'            End If
'
'            .MoveNext
'         Loop
'      End With
'
'      '[異動資料庫]
'      For i = 1 To intCP14
'         If douCP14(i) <> 0 Then
'            If bolPcp12F = True Then '每個分配點數人員都要扣20%給P1001
'               douP1001 = douP1001 + (douCP14(i) * 20 / 100)
'               douCP14(i) = douCP14(i) - (douCP14(i) * 20 / 100)
'            End If
'            '分配點數之總計
'            douDotNum = douDotNum + douCP14(i)
'            '存檔
'            strSql = "INSERT INTO acc1N0 (a1N01,a1N02,a1N03,a1N04,a1N05,a1N06) " & _
'               "VALUES ('" & strItemNo & "','" & strCP14(i) & "'," & douCP14(i) & ",'" & strUserNum & "'," & strSrvDate(1) & ",to_char(sysdate, 'HH24MISS'))"
'            cnnConnection.Execute strSql
'         End If
'      Next i
'
'      For i = 1 To 7
'         strTemp = "": douTemp = 0
'         If i = 1 And douF4101 <> 0 Then strTemp = "F4101": douTemp = douF4101
'         If i = 2 And douF4102 <> 0 Then strTemp = "F4102": douTemp = douF4102
'         If i = 3 And douF4103 <> 0 Then strTemp = "F4103": douTemp = douF4103
'         If i = 4 And douP2001 <> 0 Then strTemp = "P2001": douTemp = douP2001
'         If i = 5 And dou97099 <> 0 Then strTemp = "97099": douTemp = dou97099
'         If i = 6 And dou97098 <> 0 Then strTemp = "97098": douTemp = dou97098
'         If i = 7 And douP1001 <> 0 Then strTemp = "P1001": douTemp = douP1001
'         If douTemp <> 0 Then
'            If i < 7 Then
'               If bolPcp12F = True Then '每個分配點數人員都要扣20%給P1001
'                  douP1001 = douP1001 + (douTemp * 20 / 100)
'                  douTemp = douTemp - (douTemp * 20 / 100)
'               End If
'            End If
'            '分配點數之總計
'            douDotNum = douDotNum + douTemp
'            '存檔
'            strSql = "INSERT INTO acc1N0 (a1N01,a1N02,a1N03,a1N04,a1N05,a1N06) " & _
'               "VALUES ('" & strItemNo & "','" & strTemp & "'," & douTemp & ",'" & strUserNum & "'," & strSrvDate(1) & ",to_char(sysdate, 'HH24MISS'))"
'            cnnConnection.Execute strSql
'         End If
'      Next i
'
'      '檢查分配點數之總計是否=總分配點數
'      If Val(douDotNum) <> Val(douXXX) Then
'         '刪檔
'         strSql = "delete from acc1N0 where a1N01='" & strItemNo & "'"
'         cnnConnection.Execute strSql
'
'         MsgBox "分配點數之總計不等於總分配點數，請洽資訊系統人員！", vbExclamation, "點數分配有誤"
'         GoTo ErrorHandler
'      End If
'   End If
'
'cnnConnection.CommitTrans
'Exit Function
'ErrorHandler:
'    cnnConnection.RollbackTrans
'    RunPointPei = False
'End Function

'*************************************************
'  開啟資料表
'
'*************************************************
'edit by nick 2005/06/28
'Private sub OpenTable()
Private Function OpenTable() As Boolean
Dim strSystemKind As String
Dim douFee As Double
Dim strCP10Text As String '組案件性質代號字串
Dim dblFCT10101 As Double
Dim ii As Integer
Dim intCnt As Integer
Dim stA1L04 As String, stA1L05 As String 'Add by Morgan 2004/10/4
Dim stDisc As String 'Add by Morgan 2004/12/17 折扣
'Add by Morgan 2011/3/4
Dim bolHave35 As Boolean '商標是否有申請35類
Dim str1stItemCode As String '第一項商標申請規費代碼
Dim intOverItemCnt As Integer '超過商品數
'Add By Sindy 2013/1/24
Dim i As Integer
Dim strCompDate As Double
'2013/1/24 End
Dim bolOverTime As Boolean 'Added by Morgan 2016/7/1 逾期
'Dim strFA10 As String 'Added by Morgan 2018/1/15 代理人國籍 'Removed by Morgan 2025/7/3 改全域變數
Dim strFA76 As String 'Added by Morgan 2019/1/28 代理人性質
Dim dblRate As Double 'Added by Morgan 2020/3/6 由收文金額推算的折扣
'Add by Amy 2025/11/12
Dim jj As Integer, stOldCP09 As String, stOldCP10 As String, stCP10Item As String, stChkCP10Item As String, stAmtNotSame(1) As String, stTpMsg As String
'         列印申請人              /          列印對象         /          請款對象        /          帳款已清         /      固定請款對象    /     固定列印對象    / 是否列印申請人
Dim o_A1K04 As String, o_A1K27 As String, o_A1K28 As String, o_A1K29 As String, o_TM56 As String, o_TM69 As String, o_CCM20 As String
'end 2025/11/12

'add by nickc 2005/06/28
OpenTable = False
On Error GoTo Checking
   stRemindMsg = "": stShowMsg = "" 'Add by Amy 2025/11/12
   
   'Added by Morgan 2014/8/6 從下面搬上來並加以讀取的資料設定畫面欄位
   adoacc1k0.CursorLocation = adUseClient
   adoacc1k0.Open "select * from acc1k0 where a1k01 = '" & strItemNo & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   
   '從 formshow 搬來
   Text7 = adoacc1k0.Fields("a1k13").Value
   Text21 = adoacc1k0.Fields("a1k14").Value
   Text22 = adoacc1k0.Fields("a1k15").Value
   Text23 = adoacc1k0.Fields("a1k16").Value
   m_strCP10 = GetCP10(strCon9)
   'end 2014/8/6
   
   strFA10 = GetPrjNationNumber(adoacc1k0.Fields("a1k03")) 'Added by Morgan 2018/1/15
   strFA76 = PUB_GetFAgentFA76(adoacc1k0.Fields("a1k03")) 'Added by Morgan 2019/1/28
   
   'Added  by Morgan 2015/11/19
   '商標種類
   m_strTM08 = ""
   If CheckSys(Text7) = "2" Then
      strExc(0) = "select tm08 from trademark where tm01='" & Text7 & "' and tm02='" & Text21 & "' and tm03='" & Text22 & "' and tm04='" & Text23 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_strTM08 = "" & RsTemp(0)
      End If
   End If
   'end 2015/11/19
   
   'Add By Sindy 2012/12/27
   Frame1.Visible = False
   Frame2.Visible = False
   bolIsFMP = False
   'DataGrid1.Columns(5).Visible = False '輸入幣別
   'DataGrid1.Columns(6).Visible = False '輸入幣別金額
   DataGrid1.Columns(7).Visible = False '輸入RMB金額
   strExc(0) = "select cp01,cp12,a1k22,a1k19 from caseprogress,acc1k0 where cp60='" & strItemNo & "' and cp60=a1k01(+) and cp01='P' and substr(cp12,1,1)='F'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      '記錄日期
      If Not IsNull(RsTemp.Fields("a1k19")) And RsTemp.Fields("a1k19") > 0 Then
         strCompDate = DBDATE(RsTemp.Fields("a1k19"))
      '更新日期
      ElseIf Not IsNull(RsTemp.Fields("a1k22")) And RsTemp.Fields("a1k22") > 0 Then
         strCompDate = DBDATE(RsTemp.Fields("a1k22"))
      End If
      'FMP開放可以輸入各幣別金額
      If strCompDate >= AccFMPImputCurrStarDate Then
         Frame2.Visible = True
         bolIsFMP = True
         'DataGrid1.Columns(5).Visible = True
         'DataGrid1.Columns(6).Visible = True
         DataGrid1.Columns(7).Visible = True
      Else
         Frame1.Visible = True
      End If
   Else
      Frame1.Visible = True
   End If
   '2012/12/27 End
   
   'Add by Morgan 2011/3/23
   'Modify By Sindy 2020/1/10 and cp118='Y' => and cp118 is not null
   'strExc(0) = "select 1 from caseprogress where cp60='" & strItemNo & "' and cp01 in ('T','FCT') and cp10='101' and cp118='Y'"
   strExc(0) = "select 1 from caseprogress where cp60='" & strItemNo & "' and cp01 in ('T','FCT') and cp10='101' and cp118 is not null"
   '2020/1/10 END
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      m_boleFiling = True
      Text10.Enabled = True
   Else
      m_boleFiling = False
      Text10.Enabled = False
   End If
   'end 2011/3/23
   
   adocheck.CursorLocation = adUseClient
   adocheck.Open "select a1l01 from acc1l0 where a1l01 = '" & strItemNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount = 0 Then
       'Acc1l0無資料
       m_blnAcc1l0NoData = True
   Else
       'Acc1l0有資料
       m_blnAcc1l0NoData = False
   End If
   
   'Added by Morgan 2025/7/22 從 formshow 搬來,ChkItemFCP02會用
    '列印對象
    '若無資料(新增)
    'Modified by Morgan 2014/8/18 考慮整批請款
    'If m_blnAcc1l0NoData = True Then
    If m_blnAcc1l0NoData = True And m_bolIsBatch = False Then
    'end 2014/8/18
        Me.Text6.Text = PUB_GetA1K27("" & adoacc1k0.Fields("a1k13").Value, "" & adoacc1k0.Fields("a1k14").Value, "" & adoacc1k0.Fields("a1k15").Value, "" & adoacc1k0.Fields("a1k16").Value, m_strCP10)
        
        'Added by Morgan 2017/7/7 列印對象為 Y54443 OMG Electronic Chemicals, LLC 時預設特殊帳單 -- 陳增廣
        If Text6 = "Y54443000" Then
            Text5 = "Y"
        End If
        'end 2017/7/7
    '若有資料(修改)
    Else
        If IsNull(adoacc1k0.Fields("a1k27").Value) Then
            Text6 = ""
        Else
            Text6 = adoacc1k0.Fields("a1k27").Value
        End If
    End If
    Text6.Tag = Text6 'Added by Morgan 2012/12/6
    
    '請款對象
    '若無資料(新增)
    'Modified by Morgan 2014/8/18 考慮整批請款
    'If m_blnAcc1l0NoData = True Then
    If m_blnAcc1l0NoData = True And m_bolIsBatch = False Then
    'end 2014/8/18
        Me.Text8.Text = PUB_GetA1K28("" & adoacc1k0.Fields("a1k13").Value, "" & adoacc1k0.Fields("a1k14").Value, "" & adoacc1k0.Fields("a1k15").Value, "" & adoacc1k0.Fields("a1k16").Value, m_strCP10)
        If Me.Text8.Text = "" Then Me.Text8.Text = Me.Text2.Text
    '若有資料(修改)
    Else
        If IsNull(adoacc1k0.Fields("a1k28").Value) Then
            Text8 = ""
        Else
            Text8 = adoacc1k0.Fields("a1k28").Value
            'If Text8 = "Y48292000" Then Text19 = adoacc1k0.Fields("a1k10"): dblRate = Val(Text19) 'HP用報價匯率 Added by Morgan 2012/9/18
        End If
    End If
    Text8.Tag = Text8.Text
   'end 2025/7/22
   
   'Added by Morgan 2015/10/30 從下面搬上來
   '2009/4/23 ADD BY SONIA 抓有輸入過匯率的請款幣別
   Combo3.Clear
   Combo3.AddItem "USD"
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select distinct DNR01 from DebitNoteRate order by DNR01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoquery.EOF = False
      Combo3.AddItem adoquery.Fields("DNR01").Value
      adoquery.MoveNext
   Loop
   'Add By Sindy 2013/1/24
   'FMP不可為NTD或RMB
'Removed by Morgan 2016/8/1 取消限制--David
'   If bolIsFMP = True Then
'      For i = Combo3.ListCount - 1 To 0 Step -1
'         If Combo3.List(i) = "NTD" Or Combo3.List(i) = "RMB" Then
'            Combo3.RemoveItem i
'         End If
'      Next i
'   End If
'end 2016/8/1
   '2013/1/24 End
   adoquery.Close
   'Add By Sindy 2025/3/20 "輸入幣別"與"幣別"下拉選單內容相同
   Combo5.Clear
   If bolIsFMP = True Then
      Combo5.AddItem "RMB"
      Combo5.AddItem "USD"
      Combo5.ListIndex = 1
   Else
      For ii = 0 To Me.Combo3.ListCount - 1
         Combo5.AddItem Combo3.List(ii)
         If Trim(Combo3.List(ii)) = "NTD" Then
            int_NTD = ii
         End If
      Next ii
      Combo5.ListIndex = int_NTD
   End If
   '2025/3/20 END
   
   'FormShow 'Removed by Morgan 2016/6/27 此時變數 CP10have926 尚未設定判斷有誤,改移回下面
   'end 2015/10/30
   
   m_AppNo = GetPrjPeopleNum1(Text7 & "-" & Text21 & "-" & Text22 & "-" & Text23) 'Added by Morgan 2015/8/5
   
'Removed by Morgan 2015/11/27 檢查統一移到更新時做並改用函數
'   'Added by Morgan 2015/7/13
'   m_bolNoDisbursements = False
'   If Text7 = "FCP" Then
'      'Modified by Morgan 2015/8/5 改用m_AppNo
'      'strExc(1) = GetPrjPeopleNum1(Text7 & "-" & Text21 & "-" & Text22 & "-" & Text23)
'      'Modified by Morgan 2015/8/3 +X49346,X72101
'      'Modified by Morgan 2015/11/27 +X62773
'      If m_AppNo = "X56842000" Or m_AppNo = "X49346000" Or m_AppNo = "X72101000" Then
'         m_bolNoDisbursements = True
'      End If
'   End If
'   'end 2015/7/13
'end 2015/11/27
   
    strCP10Text = ","
   If adocheck.RecordCount = 0 Then
    '系統類別為T者不預設請款項目 2009/10/13 cancel by sonia因T本用01,02,03請款項目,2009/5改為以案件性質為請款項目
   'If adocheck.RecordCount = 0 And text7.text <> "T" Then
      adoselect.CursorLocation = adUseClient
      '92.7.8 MODIFY BY SONIA
      'adoselect.Open "select a1j01, a1j02 as No, a1j17, nvl(cp17, 0) as cp17, cp01, cp02, cp03, cp04, nvl(cp16, 0) as cp16 from acc1j0, caseprogress where a1j01 = cp01 and substr(a1j02, 1, 3) = cp10 and (substr(a1j02, 4, 2) = '99' or length(a1j02) = 3) and cp60 = '" & strItemNo & "' order by a1j02 asc", adoTaie, adOpenStatic, adLockReadOnly
'      adoselect.Open "select a1j01, a1j02 as No, a1j17, nvl(cp17, 0) as cp17, cp01, cp02, cp03, cp04, nvl(cp16, 0) as cp16 from acc1j0, caseprogress where a1j01 = cp01 and (a1j02 = cp10 OR (substr(a1j02, 1, 3) = CP10 AND substr(a1j02, 4, 2) = '99')) and cp60 = '" & strItemNo & "' order by a1j02 asc", adoTaie, adOpenStatic, adLockReadOnly
      strSql = ""
      Select Case Text7.Text
         Case "FCT"
            strSql = strSql & " Order By CP09, A1J02 "
         Case Else
            'Modify by Morgan 2010/4/13 FCP的940要排在前面(因列印時會改為101/102/103/105)
            'strSql = strSql & " Order By A1J02 "
            '2013/5/17 modify by sonia FMP案也依收文號順序,P-104767
            'strSql = strSql & " Order By decode(A1j01||substr(A1j02,1,3),'FCP940','101'||substrb(A1J02,4),A1J02) "
            '2013/5/20 modify by sonia 再改先依cp05順序
            If bolIsFMP = True Then
               strSql = strSql & " Order By CP05,CP09, A1J02 "
            Else
               strSql = strSql & " Order By decode(A1j01||substr(A1j02,1,3),'FCP940','101'||substrb(A1J02,4),A1J02) "
            End If
            '2013/5/17 end
      End Select
      
      '2006/1/27 MODIFY BY SONIA 加CFP之160399(X09410601)
      'StrSql = "select a1j01, a1j02 as No, a1j17, nvl(cp17, 0) as cp17, cp01, cp02, cp03, cp04, nvl(cp16, 0) as cp16, CP10 from acc1j0, caseprogress where a1j01 = cp01 and (a1j02 = cp10 OR (substr(a1j02, 1, 3) = CP10 AND substr(a1j02, 4, 2) = '99')) and cp60 = '" & strItemNo & "' " & StrSql
      'Modified by Morgan 2014/8/12 +cp09,cp07,cp27
      'Modify By Sindy 2017/3/1 FMP案的請款項目,增加自動預設(98代收代付)項目
      If bolIsFMP = True Then
         strSql = "select a1j01, a1j02 as No, a1j17, nvl(cp17, 0) as cp17, cp01, cp02, cp03, cp04, nvl(cp16, 0) as cp16, CP10,cp09,cp07,cp27 from acc1j0, caseprogress where a1j01 = cp01 and (a1j02 = cp10 OR (substr(a1j02, 1, 3) = CP10 AND substr(a1j02, 4, 2) = '99') OR (substr(a1j02, 1, 3) = CP10 AND substr(a1j02, 4, 2) = '98') OR (substr(a1j02, 1, 4) = CP10 AND substr(a1j02, 5, 2) = '99') OR (substr(a1j02, 1, 4) = CP10 AND substr(a1j02, 5, 2) = '98')) and cp60 = '" & strItemNo & "' " & strSql
      Else
      '2017/3/1 END
         strSql = "select a1j01, a1j02 as No, a1j17, nvl(cp17, 0) as cp17, cp01, cp02, cp03, cp04, nvl(cp16, 0) as cp16, CP10,cp09,cp07,cp27 from acc1j0, caseprogress where a1j01 = cp01 and (a1j02 = cp10 OR (substr(a1j02, 1, 3) = CP10 AND substr(a1j02, 4, 2) = '99') OR (substr(a1j02, 1, 4) = CP10 AND substr(a1j02, 5, 2) = '99')) and cp60 = '" & strItemNo & "' " & strSql
      End If
      '2006/1/27 END
      adoselect.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      '92.7.8 END
      Do While adoselect.EOF = False
         'Add by Amy 2025/11/12 由結案單進入
         If stF0301 <> "" And stUpdCP09 = "" Then
            '若同時請 102-延期,會自動產生 02之請款項目,若結案單請款項目也有此項目,無法知道屬於哪個總收文號,故於此比對
            If stOldCP09 <> "" And stOldCP09 <> "" & adoselect.Fields("CP09") Then
               If ChkAndSetCCDItem("1", stOldCP09, stOldCP10, strItemNo, Mid(stCP10Item, 2), stTpMsg) = False Then
                  stAmtNotSame(0) = stAmtNotSame(0) & stTpMsg
               End If
               stCP10Item = ""
            End If
            '請款項目前3碼不同,檢查結案單及請款單項目 總金額ex:請2道303
            If Left(stOldCP10, 3) <> "" And Left(stOldCP10, 3) <> "" & adoselect.Fields("CP10") Then
               '檢查請款項目前3碼與結案單請款項目前3碼相同之總金額不相同,彈訊息
               If ChkCCDAndAcc1L0NotSame(1, Me.Name, stF0301, strItemNo, stTpMsg, Mid(stChkCP10Item, 2), Left(stOldCP10, 3)) = True Then
                  stAmtNotSame(1) = stAmtNotSame(1) & "," & stTpMsg
               End If
               stChkCP10Item = ""
            End If
         End If
         
         strCP10Text = strCP10Text & "" & adoselect.Fields("CP10").Value & ","
         strMaxNo = GetMaxNo(strItemNo)
         'Add by Amy 2025/11/12
         stCP10Item = stCP10Item & "," & strMaxNo
         stChkCP10Item = stChkCP10Item & "," & strMaxNo
         'end 2025/11/22
         
         '2006/1/27 MODIFY BY SONIA 因CFP之1603(X09410601)
         'If Len(adoselect.Fields("No").Value) <> 3 Then
         If Len(adoselect.Fields("No").Value) > 4 Then
         '2006/1/27 END
            'If Mid(adoselect.Fields("No").Value, 4, 2) = "99" Then
               strDiscount = "0"
            'End If
            douFee = Val(adoselect.Fields("cp17").Value)
            
            'Add by Morgan 2004/10/5 71599,71699,71799 第一類規費
            If "" & adoselect.Fields("CP01").Value = "FCT" And InStr("71599,71699,71799", "" & adoselect.Fields("No").Value) > 0 Then
               If adoselect.Fields("No").Value = "71599" Then
                  douFee = 1000
               ElseIf adoselect.Fields("No").Value = "71699" Then
                  douFee = 1500
               Else
                  douFee = 2500
               End If
            End If
            '2004/10/5 END
            
         Else
'            strDiscount = DiscountShow(adoselect.Fields("cp01").Value, adoselect.Fields("cp02").Value, adoselect.Fields("cp03").Value, adoselect.Fields("cp04").Value)
            'Added by Morgan 2014/8/18 考慮整批請款
            If m_bolIsBatch Then
               strDiscount = 100 - Val(m_Discount)
            Else
            'end 2014/8/18
               strDiscount = 100 - Val(PUB_GetA1L07Disc(adoselect.Fields("cp01").Value, adoselect.Fields("cp02").Value, adoselect.Fields("cp03").Value, adoselect.Fields("cp04").Value, adoselect.Fields("cp10").Value, strSrvDate(2)))
            End If 'Added by Morgan 2014/8/18
            
            If strDiscount = "100" Then strDiscount = 0
            douFee = Val(adoselect.Fields("cp16").Value) - Val(adoselect.Fields("cp17").Value)
            'Add by Morgan 2004/10/4 查名服務費 7000
            'Modified by Morgan 2011/10/31 查名服務費改 6000--陳金蓮
            'Modified by Morgan 2018/1/15 +FCT(轉案後請款)
            If ("" & adoselect.Fields("CP01").Value = "S" Or "" & adoselect.Fields("CP01").Value = "FCT") And "" & adoselect.Fields("No").Value = "001" Then
               'Modified by Moran 2018/1/15 日本區查名第1類預設5000(餘額),第2類以後3000,雜費500 --陳金蓮
               'douFee = 6000
               If Left(strFA10, 3) = "011" Then
                  douFee = 5000
               Else
                  douFee = 6000
               End If
               'end 2018/1/15
            '2004/10/4 END
            'Add By Cheng 2003/09/30
            ElseIf "" & adoselect.Fields("CP01").Value = "FCT" Then
                If adoselect.Fields("No").Value = "101" Then
                  'Added by Morgan 2020/3/6
                  '日本區商申 雜費：500, 第1類：9000, 第2類以上：各7000 ; 折扣預設相同由收文金額推算--湘嫻
                  If Left(strFA10, 3) = "011" Then
                     intCnt = GetTMKindCnt(Text7.Text, Text21.Text, Text22.Text, Text23.Text) - 1
                     dblFCT10101 = Val("" & adoselect.Fields("CP16").Value - Val("" & adoselect.Fields("CP17").Value) - 500)
                     '收文金額大於正常費用時表示有超項費，因無法預知(將來考慮發文輸入儲存)，維持不帶折扣
                     If dblFCT10101 > Val(9000# + intCnt * 7000#) Then
                        dblRate = 1
                        douFee = 9000
                     Else
                        dblRate = dblFCT10101 / Val(9000# + intCnt * 7000#)
                        '若折扣有小數時可能是第1類9折第2類以後折扣不同情形
                        If dblRate <> Round(dblRate, 2) Then
                           If intCnt = 0 Then
                              dblRate = 1
                              douFee = 9000
                           Else
                              douFee = Val(9000# * 0.9)
                              dblRate = Val(dblFCT10101 - douFee) / Val(intCnt * 7000#)
                           End If
                        Else
                           douFee = Val(9000# * dblRate)
                        End If
                     End If
                     dblFCT10101 = dblFCT10101 - douFee
                  Else
                  'end 2020/3/6
                  
                     'Modify by Morgan 2011/2/9 第1類服務費9000
                     'douFee = 10000
                     douFee = 9000
                     'end 2011/2/9
                     dblFCT10101 = Val("" & adoselect.Fields("CP16").Value - douFee - Val("" & adoselect.Fields("CP17").Value) - 600)
                     
                  End If 'Added by Morgan 2020/3/6
                
                'Add by Morgan 2004/10/5 715,716,717 服務費
                ElseIf InStr("715,716,717", "" & adoselect.Fields("No").Value) > 0 Then
                     
                     'Modify by Morgan 2007/3/6 第一期&全期也要控制 -- 陳金蓮
                     'If adoselect.Fields("No").Value = "715" Then
                     '   douFee = 5500
                     'ElseIf adoselect.Fields("No").Value = "716" Then
                        'Modify by Morgan 2006/6/2 超過的金額加到716
                        'douFee = 3000
                        
                        'Modified by Morgan 2019/1/28 FCT日本區「註冊費」(717)請款 => 雜費:700
                        'douFee = Val("" & adoselect.Fields("CP16").Value - Val("" & adoselect.Fields("CP17").Value) - 500)
                        If adoselect.Fields("No").Value = "717" And Left(strFA10, 3) = "011" Then
                           douFee = Val("" & adoselect.Fields("CP16").Value - Val("" & adoselect.Fields("CP17").Value) - 700)
                        Else
                           douFee = Val("" & adoselect.Fields("CP16").Value - Val("" & adoselect.Fields("CP17").Value) - 500)
                        End If
                        'end 2019/1/28
                        
                     'Else
                     '   douFee = 5500
                     'End If
                     
                     'Add by Morgan 2007/3/6 還要減掉跨類的服務費
                     intCnt = GetTMKindCnt(Text7.Text, Text21.Text, Text22.Text, Text23.Text) - 1
                     'Modify by Morgan 2011/2/9 跨類第二期2500 其他3000(第1類 第二期3500 其他5500)
                     'douFee = douFee - intCnt * 1000#
                     If adoselect.Fields("No").Value = "716" Then
                        douFee = douFee - intCnt * 2500#
                     'Added by Morgan 2019/1/28
                     'FCT日本區「註冊費」(717)請款第２個類別以上 => 經代理人委辦:1,500/每類;自行來所:3,000/每類
                     ElseIf adoselect.Fields("No").Value = "717" And Left(strFA10, 3) = "011" Then
                        If strFA76 = "A" Then
                           douFee = douFee - intCnt * 1500#
                        Else
                           douFee = douFee - intCnt * 3000#
                        End If
                     'end 2019/1/28
                     Else
                        douFee = douFee - intCnt * 3000#
                     End If
                     'end 2011/2/9
                     'end 2007/3/6
                
                '2006/5/25 ADD BY SONIA
                'Modify by Morgan 2010/8/26 +501移轉 -- 陳金蓮
                'ElseIf adoselect.Fields("No").Value = "102" Then
                'Modified by Morgan 2017/6/1 延展移轉雜費不同不可合併--陳金蓮 Ex.X10608529(FCT-14135)
                'ElseIf (adoselect.Fields("No").Value = "102" Or adoselect.Fields("No").Value = "501") Then
                ElseIf adoselect.Fields("No").Value = "102" Then
                'end 2017/6/1
                    'Modified by Morgan 2014/8/14 延展雜費改 500
                    'Modified by Morgan 2014/8/15 延展雜費改回 600
                    'Modified by Morgan 2014/12/2 延展雜費再改 500
                     douFee = Val("" & adoselect.Fields("CP16").Value - Val("" & adoselect.Fields("CP17").Value) - 500)
                    
                     'Add by Morgan 2011/2/9 延展跨類4000(第1類8000)
                     If adoselect.Fields("No").Value = "102" Then
                        intCnt = GetTMKindCnt(Text7.Text, Text21.Text, Text22.Text, Text23.Text) - 1
                        'Added by Morgan 2020/3/4 日本區延展跨類500/1000
                        If Left(strFA10, 3) = "011" Then
                           '經代理人委辦
                           If strFA76 = "A" Then
                              douFee = douFee - intCnt * 500#
                           '自行來所
                           Else
                              douFee = douFee - intCnt * 1000#
                           End If
                        Else
                        'end 2020/3/4
                        
                           douFee = douFee - intCnt * 4000#
                        End If 'Added by Morgan 2020/3/4
                     End If
                     'end 2011/2/9
                     
                'Added by Moragn 2017/6/1
                ElseIf adoselect.Fields("No").Value = "501" Then
                  douFee = Val("" & adoselect.Fields("CP16").Value - Val("" & adoselect.Fields("CP17").Value) - 600)
                'end 2017/6/1
                
                'Added by Morgan 2018/10/9
                'FCT「補正」之日文請款項目預設為201=扣掉(02)雜費後之金額,02=500 -- 陳金蓮
                ElseIf Left(strFA10, 3) = "011" And adoselect.Fields("No").Value = "201" Then
                  douFee = Val("" & adoselect.Fields("CP16").Value - Val("" & adoselect.Fields("CP17").Value) - 500)
                End If
            End If
            'End
         End If
         strSystemKind = adoselect.Fields("cp01").Value
        'Modify By Cheng 2004/04/23
        '規費項目(xxx99)不預設折扣
'         adoTaie.Execute "insert into acc1l0 values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & adoselect.Fields("No").Value & "', " & douFee & ", null, " & douFee * Val(strDiscount) / 100 & ", '" & strUserNum & "', " & Val(FCDate(ServerDate)) & ", " & ServerTime & ", null, null, null)"
         'Modify by Morgan 2004/12/17 規費項目且請款金額為0的不新增
         '2006/1/27 MODIFY BY SONIA 因CFP之1603(X09410601)
         'If Not (Len(adoselect.Fields("No").Value) <> 3 And douFee = 0) Then
         'modify by sonia 2013/5/17 FMP案要新增P-104767
         'If Not (Len(adoselect.Fields("No").Value) > 4 And douFee = 0) Then
         If Not (Len(adoselect.Fields("No").Value) > 4 And douFee = 0) Or bolIsFMP = True Then
         '2006/1/27 END
            '2005/7/8 MODIFY BY SONIA A1L08應存民國年
            'adoTaie.Execute "insert into acc1l0 values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & adoselect.Fields("No").Value & "', " & douFee & ", null, " & douFee * Val(IIf(InStr("" & adoselect.Fields("No").Value, "99,") > 0, 100, strDiscount)) / 100 & ", '" & strUserNum & "', " & Val(FCDate(ServerDate)) & ", " & ServerTime & ", null, null, null)"
            'Modify By Sindy 2010/01/28 改為每個類別之規費金額先預設為3000元,剩下的金額則掛在第1個規費
            'Modified by Morgan 2015/11/19 證明標章(系統商標種類:7)及團體標章(系統商標種類:8)除外--陳金蓮
            If Text7.Text = "FCT" And adoselect.Fields("No").Value = "10199" And m_strTM08 <> "7" And m_strTM08 <> "8" Then
               
               intCnt = GetTMKindCnt(Text7.Text, Text21.Text, Text22.Text, Text23.Text, bolHave35)
               Dim douAmt As Double, douOverAmt As Double
               douAmt = douFee - (Val(intCnt) * 3000)
               
               'Add by Morgan 2011/3/4 若有申請 35 類則第一項代碼改 A0199
               If bolHave35 Then
                  str1stItemCode = "A0199"
               Else
                  str1stItemCode = adoselect.Fields("No").Value
               End If
               'end 2011/3/4
               
'Modified by Morgan 2014/2/25 電子送件第一項規費預設2700,若只有一項再考慮超過商品數,若無法算出時,差額不必再調整到第一項,由人工調整(可能商品名稱有減免)--陳金蓮
'Modified by Morgan 2014/5/1 金額與收文不符不好,取消--陳金蓮
               If douAmt > 0 Then
                  'Modify by Morgan 2011/3/4 若只有一類時要預設超過商品數(35類除300,其他除200)
                  'adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & adoselect.Fields("No").Value & "', " & (3000 + douAmt) & ", null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
                  If intCnt = 1 Then
                     intOverItemCnt = 0
                     douOverAmt = douAmt
                     If m_boleFiling Then
                        douOverAmt = douOverAmt + 300
                     End If
                     
                     If bolHave35 Then
                        If douAmt Mod 500 = 0 Then
                           intOverItemCnt = douOverAmt / 500
                        End If
                     Else
                        If douAmt Mod 200 = 0 Then
                           intOverItemCnt = douOverAmt / 200
                        End If
                     End If

                     adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09,A1L14) values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & str1stItemCode & "', " & (3000 + douAmt) & ", null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS')," & intOverItemCnt & ")"
                  Else
                     adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & str1stItemCode & "', " & (3000 + douAmt) & ", null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
                  End If
                  'end 2011/3/4
               Else
                  adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & str1stItemCode & "', " & (3000 + douAmt) & ", null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
               End If

'               intOverItemCnt = 0
'               If intCnt = 1 Then
'                  douOverAmt = douAmt
'                  If m_boleFiling Then
'                     douOverAmt = douOverAmt + 300
'                  End If
'                  If douOverAmt > 0 Then
'                     If bolHave35 Then
'                        If douOverAmt Mod 500 = 0 Then
'                           intOverItemCnt = douOverAmt / 500
'                        End If
'                     Else
'                        If douOverAmt Mod 200 = 0 Then
'                           intOverItemCnt = douOverAmt / 200
'                        End If
'                     End If
'                  End If
'               End If
'
'               If intOverItemCnt > 0 Then
'                  adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09,A1L14) values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & str1stItemCode & "', " & douFee & ", null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS')," & intOverItemCnt & ")"
'               Else
'                  adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & str1stItemCode & "', " & IIf(m_boleFiling = True, 2700, 3000) & ", null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
'               End If
'end 2014/2/25

               'Add by Morgan 2010/3/1 從下面移上來
               '預設第二個類別以後的規費項目(10199)
               For ii = 2 To intCnt
                  strMaxNo = GetMaxNo(strItemNo)
                  'Add by Amy 2025/11/12
                  stCP10Item = stCP10Item & "," & strMaxNo
                  stChkCP10Item = stChkCP10Item & "," & strMaxNo
                  'end 2025/11/22
                  adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & Text7.Text & "', '10199', 3000, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
               Next ii
               'End 2010/3/1
               
            '2010/01/28 End
            '2010/10/25 ADD BY SONIA FCP超頁費,超項費無服務費不必新增
            ElseIf "" & adoselect.Fields("CP01").Value = "FCP" And douFee = 0 And (adoselect.Fields("No").Value = "939" Or adoselect.Fields("No").Value = "938") Then
            '2010/10/25 end
            
            'Added by Morgan 2014/8/12
            '延展規費1類4000,過期為8000
            ElseIf Text7.Text = "FCT" And adoselect.Fields("No").Value = "10299" Then
               intCnt = GetTMKindCnt(Text7.Text, Text21.Text, Text22.Text, Text23.Text) - 1
               bolOverTime = False
               If Not IsNull(adoselect("cp07")) And Not IsNull(adoselect("cp27")) Then
                  '逾期
                  If adoselect("cp27") > PUB_GetWorkDay1(adoselect("cp07"), False) Then
                     bolOverTime = True
                     douFee = douFee - intCnt * 8000#
                  Else
                     douFee = douFee - intCnt * 4000#
                  End If
               Else
                  douFee = douFee - intCnt * 4000#
               End If
               strMaxNo = GetMaxNo(strItemNo)
               'Add by Amy 2025/11/12
               stCP10Item = stCP10Item & "," & strMaxNo
               stChkCP10Item = stChkCP10Item & "," & strMaxNo
               'end 2025/11/22
               adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & adoselect.Fields("No").Value & "', " & douFee & ", null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
                  
               For ii = 1 To intCnt
                  strMaxNo = GetMaxNo(strItemNo)
                  'Add by Amy 2025/11/12
                  stCP10Item = stCP10Item & "," & strMaxNo
                  stChkCP10Item = stChkCP10Item & "," & strMaxNo
                  'end 2025/11/22
                  'Modified by Morgan 2016/7/1 逾期8000
                  adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & adoselect.Fields("No").Value & "', " & IIf(bolOverTime, 8000, 4000) & ", null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
               Next
            'end 2014/8/12
            
            Else
               'Modified by Morgan 2013/5/2
               'adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & adoselect.Fields("No").Value & "', " & douFee & ", null, " & douFee * Val(IIf(InStr("" & adoselect.Fields("No").Value, "99,") > 0, 100, strDiscount)) / 100 & ", '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
               'Modified by Morgan 2015/8/5 +申請人 m_AppNo
               'strExc(1) = 100 * (1 - PUB_GetDiscX(strCon1, adoselect.Fields("CP01").Value, adoselect.Fields("No").Value, 1 - Val(strDiscount) / 100))
               'adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & adoselect.Fields("No").Value & "', " & douFee & ", null, " & douFee * Val(IIf(InStr("" & adoselect.Fields("No").Value, "99,") > 0, 100, strExc(1))) / 100 & ", '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
               'Modified by Lydia 2016/09/07 FMP案不預設折扣
               'Modified by Morgan 2017/9/30 FMP案改回要預設--David
               'If Right("" & adoselect.Fields("No").Value, 2) = "99" Or bolIsFMP Then
               If Right("" & adoselect.Fields("No").Value, 2) = "99" Then
                  strExc(1) = 0
                  'Add By Sindy 2017/3/1 代收代付項目預設為0
                  If bolIsFMP = True And _
                     (Len("" & adoselect.Fields("No").Value) = 5 And Right("" & adoselect.Fields("No").Value, 2) = "98" _
                      Or Len("" & adoselect.Fields("No").Value) = 6 And Right("" & adoselect.Fields("No").Value, 2) = "98") Then
                     douFee = 0
                  End If
                  '2017/3/1 END
               
               'Added by Morgan 2020/3/4 日本區延展折扣=(收文金額/8000)
               ElseIf Text7.Text = "FCT" And adoselect.Fields("No").Value = "102" And Left(strFA10, 3) = "011" Then
                  strExc(1) = 1 - douFee / 8000
                  douFee = 8000
               
               'Added by Morgan 2020/3/6
               '日本區商申 雜費：500, 第1類：9000, 第2類以上：各7000 ; 折扣預設相同由收文金額推算--湘嫻
               ElseIf Text7.Text = "FCT" And adoselect.Fields("No").Value = "101" And Left(strFA10, 3) = "011" Then
                  strExc(1) = 1 - douFee / 9000
                  douFee = 9000
               Else
                  '參數1不可改 Text2 (此時尚未有值,後面才設定)
                  strExc(1) = 1 - PUB_GetDiscX(strCon1, adoselect.Fields("CP01").Value, adoselect.Fields("No").Value, 1 - Val(strDiscount) / 100, m_AppNo)
               End If
               'modify by sonia 2021/3/25 FCP,FG之補收款911要改用相關總收文號之案件性質請款
               'adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & adoselect.Fields("No").Value & "', " & douFee & ", null, " & douFee * Val(strExc(1)) & ", '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
               If (Text7.Text = "FCP" Or Text7.Text = "FG") And Left(adoselect.Fields("No").Value, 3) = "911" Then
                  If douFee > 0 Then
                     strExc(0) = "select b.cp09 cp09,b.cp10||substr('" & adoselect.Fields("No").Value & "',4) cp10 from caseprogress a,caseprogress b where a.cp09='" & adoselect.Fields("cp09").Value & "' and a.cp43=b.cp09(+)"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & RsTemp.Fields("cp10").Value & "', " & douFee & ", null, " & douFee * Val(strExc(1)) & ", '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
                     End If
                  End If
               Else
                  adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & adoselect.Fields("No").Value & "', " & douFee & ", null, " & douFee * Val(strExc(1)) & ", '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
               End If
               'end 2021/3/25
               'end 2015/8/5
               'end 2013/5/2
            End If
            '2005/7/8 END
         End If
         
        'End
        'Add by Morgan 2004/10/4 查名服務費(001) 2~N->7000, 雜費(02) -> 600*N
        'Modified by Morgan 2011/10/31 查名服務費改 6000--陳金蓮
        'Modified by Morgan 2015/1/23 雜費(02)改固定列1個,傳真(01)改不要,差額調整到查名(001)
        'Modified by Morgan 2018/1/15 +FCT(轉案後請款)
        If (adoselect.Fields("CP01").Value = "S" Or adoselect.Fields("CP01").Value = "FCT") And adoselect.Fields("No").Value = "001" Then
            intCnt = GetSPKindCnt(Text7.Text, Text21.Text, Text22.Text, Text23.Text) - 1
            
            'Added by Moran 2018/1/15 日本區查名第1類預設5000(餘額),第2類以後3000,雜費500 --陳金蓮
            If Left(strFA10, 3) = "011" Then
               '第1類
               stA1L05 = Format(Val("" & adoselect.Fields("CP16").Value) - (5000# + 3000# * intCnt) - (500#))
               If Val(stA1L05) <> 0 Then
                  adoTaie.Execute "update acc1l0 set a1l05=a1l05+(" & stA1L05 & ") where a1l01='" & strItemNo & "' and a1l04='001' and a1l05=5000 and rownum<2", intI
               End If
               '第2類以後
               For ii = 1 To intCnt
                   strMaxNo = GetMaxNo(strItemNo)
                   'Add by Amy 2025/11/12
                   stCP10Item = stCP10Item & "," & strMaxNo
                   stChkCP10Item = stChkCP10Item & "," & strMaxNo
                   'end 2025/11/22
                   adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & adoselect.Fields("No").Value & "', 3000, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
               Next ii
               '雜費
               strMaxNo = GetMaxNo(strItemNo)
               'Add by Amy 2025/11/12
               stCP10Item = stCP10Item & "," & strMaxNo
               stChkCP10Item = stChkCP10Item & "," & strMaxNo
               'end 2025/11/22
               stA1L05 = 500
               adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & Text7.Text & "', '02'," & stA1L05 & " , null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
            Else
            'end 2018/1/15
               'Added by Morgan 2015/1/23
               stA1L05 = Format(Val("" & adoselect.Fields("CP16").Value) - (6000# + 6000# * intCnt) - (600#))
               
               If Val(stA1L05) <> 0 Then
                  adoTaie.Execute "update acc1l0 set a1l05=a1l05+(" & stA1L05 & ") where a1l01='" & strItemNo & "' and a1l04='001' and a1l05=6000 and rownum<2", intI
               End If
               'end 2015/1/23
               
               For ii = 1 To intCnt
                   strMaxNo = GetMaxNo(strItemNo)
                   'Add by Amy 2025/11/12
                   stCP10Item = stCP10Item & "," & strMaxNo
                   stChkCP10Item = stChkCP10Item & "," & strMaxNo
                   'end 2025/11/22
                   adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & adoselect.Fields("No").Value & "', 6000, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
               Next ii
               '雜費 "02" 600*N
               strMaxNo = GetMaxNo(strItemNo)
               'Add by Amy 2025/11/12
               stCP10Item = stCP10Item & "," & strMaxNo
               stChkCP10Item = stChkCP10Item & "," & strMaxNo
               'end 2025/11/22
               'Modified by Morgan 2015/1/23
               'stA1L05 = Format(600# + 600# * intCnt)
               stA1L05 = 600
               'end 2015/1/23
               adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & Text7.Text & "', '02'," & stA1L05 & " , null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
            End If
            'Removed by Morgan 2015/1/23
            ''傳真 "01" CP16-服務費-雜費
            'strMaxNo = GetMaxNo(strItemNo)
            'stA1L05 = Format(Val("" & adoselect.Fields("CP16").Value) - (6000# + 6000# * intCnt) - (600# + 600# * intCnt))
            'adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & Text7.Text & "', '01'," & stA1L05 & " , null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
            'end 2015/1/23
        '2004/10/4 end
        'Add By Cheng 2003/12/11
        ElseIf "" & adoselect.Fields("CP01").Value = "FCT" Then
            If adoselect.Fields("No").Value = "101" Then
                '依(商品類別數-1)預設商申項目(101)數
               intCnt = GetTMKindCnt(Text7.Text, Text21.Text, Text22.Text, Text23.Text) - 1
               For ii = 1 To intCnt
                  strMaxNo = GetMaxNo(strItemNo)
                  'Add by Amy 2025/11/12
                  stCP10Item = stCP10Item & "," & strMaxNo
                  stChkCP10Item = stChkCP10Item & "," & strMaxNo
                  'end 2025/11/22
                  'Added by Morgan 2020/3/6
                  '日本區商申 雜費：500, 第1類：9000, 第2類以上：各7000 ; 折扣預設相同由收文金額推算--湘嫻
                  If Left(strFA10, 3) = "011" Then
                     adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & adoselect.Fields("No").Value & "', 7000, null," & 7000 * (1 - dblRate) & ", '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
                     dblFCT10101 = dblFCT10101 - Val(7000# * dblRate)
                  Else
                  'end 2020/3/6
                  
                     'Modify by Morgan 2011/2/9 第2類以後服務費 7000
                     'adoTaie.Execute "insert into acc1l0 values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & adoselect.Fields("No").Value & "', 5000, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ", " & ServerTime & ", null, null, null)"
                     'dblFCT10101 = dblFCT10101 - 5000 'Add by Morgan 2010/7/1
                     adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & adoselect.Fields("No").Value & "', 7000, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
                     dblFCT10101 = dblFCT10101 - 7000
                     'end 2011/2/9
                     
                  End If 'Added by Morgan 2020/3/6
               Next ii
            'Add by Morgan 2004/10/4 715,716,717 註冊費第二類以後服務費 1000
            ElseIf InStr("715,716,717", "" & adoselect.Fields("No").Value) > 0 Then
               '依(商品類別數-1)預設商申項目(101)數
                intCnt = GetTMKindCnt(Text7.Text, Text21.Text, Text22.Text, Text23.Text) - 1
                For ii = 1 To intCnt
                    strMaxNo = GetMaxNo(strItemNo)
                    'Add by Amy 2025/11/12
                    stCP10Item = stCP10Item & "," & strMaxNo
                    stChkCP10Item = stChkCP10Item & "," & strMaxNo
                    'end 2025/11/22
                    'Modify by Morgan 2011/2/9 跨類第二期2500 其他3000
                    'adoTaie.Execute "insert into acc1l0 values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & adoselect.Fields("No").Value & "', 1000, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ", " & ServerTime & ", null, null, null)"
                    If adoselect.Fields("No").Value = "716" Then
                        adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & adoselect.Fields("No").Value & "', 2500, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
                    'Added by Morgan 2019/1/28
                    'FCT日本區「註冊費」(717)請款第２個類別以上服務費 => 經代理人委辦:1,500/每類;自行來所:3,000/每類
                    ElseIf adoselect.Fields("No").Value = "717" And Left(strFA10, 3) = "011" Then
                        If strFA76 = "A" Then
                           adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & adoselect.Fields("No").Value & "', 1500, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
                        Else
                           adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & adoselect.Fields("No").Value & "', 3000, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
                        End If
                    'end 2019/1/28
                    Else
                        adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & adoselect.Fields("No").Value & "', 3000, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
                    End If
                    'end 2011/2/9
                Next ii
            '2004/10/4 end
            'Add by Morgan 2011/2/9 延展跨類4000
            ElseIf adoselect.Fields("No").Value = "102" Then
               intCnt = GetTMKindCnt(Text7.Text, Text21.Text, Text22.Text, Text23.Text) - 1
               For ii = 1 To intCnt
                  strMaxNo = GetMaxNo(strItemNo)
                  'Add by Amy 2025/11/12
                  stCP10Item = stCP10Item & "," & strMaxNo
                  stChkCP10Item = stChkCP10Item & "," & strMaxNo
                  'end 2025/11/22
                  adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & adoselect.Fields("a1j01").Value & "', '" & adoselect.Fields("No").Value & "', 4000, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
               Next
            End If
        
        End If
        'End
            'Remove by Morgan 2007/5/23 移到下面，否則acc1j0沒有設定請款項目的案件性質會沒有折扣
            'strDiscount = 100 - Val(PUB_GetA1L07Disc(adoselect.Fields("cp01").Value, adoselect.Fields("cp02").Value, adoselect.Fields("cp03").Value, adoselect.Fields("cp04").Value, adoselect.Fields("cp10").Value, strSrvDate(2)))
            'If strDiscount = "100" Then strDiscount = 0
            'If Val(strDiscount) > 0 Then stDisc = strDiscount 'Add by Morgan 2004/12/17
            'end 2007/5/23
         'Add by Amy 2025/11/12
         stOldCP09 = "" & adoselect.Fields("CP09")
         stOldCP10 = "" & adoselect.Fields("CP10")
         'end 2025/11/12
         adoselect.MoveNext
      Loop
      'Add by Amy 2025/11/12 由結案單進入且第1畫面由使用者自行選擇總收文號者
      If stF0301 <> "" And stOldCP10 <> "" Then
         '最後寫入其他道對應不到進度
         If ChkAndSetCCDItem("2", stOldCP09, stOldCP10, strItemNo, Mid(stCP10Item, 2), stTpMsg) = False Then
            stAmtNotSame(0) = stAmtNotSame(0) & "," & stTpMsg
         End If
      End If
      
      'Modify by Morgan 2007/5/23 判斷acc1j0沒有時也要抓折扣
      'adoselect.Close
      If adoselect.RecordCount = 0 Then
         adoselect.Close
         strSql = "select cp01, cp02, cp03, cp04, CP10 from caseprogress where cp60 = '" & strItemNo & "' order by cp09"
         adoselect.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      End If
      If adoselect.RecordCount > 0 Then
         adoselect.MoveLast
         stDisc = 100 - Val(PUB_GetA1L07Disc(adoselect.Fields("cp01").Value, adoselect.Fields("cp02").Value, adoselect.Fields("cp03").Value, adoselect.Fields("cp04").Value, adoselect.Fields("cp10").Value, strSrvDate(2)))
         If stDisc = "100" Then stDisc = 0
      End If
      'end 2007/5/23
      adoselect.Close
      
        'Add By Cheng 2003/09/30
        Select Case Text7.Text
        Case "FCP"
            '若有翻譯
            '92.11.24 ADD BY SONIA 增加 檢視中說
            'If InStr(strCP10Text, ",201,") > 0 Then
           '2007/6/29 modify BY SONIA 增加 製作中說
           '2007/12/4 modify BY SONIA 增加 新式樣創作說明
           'Modified by Morgan 2020/2/3 改新案申請時自動新增202補文件、106主張國際優先權(原03打字費還是掛在新案翻譯不動)--敏莉
           If InStr(strCP10Text, ",101,") > 0 Or InStr(strCP10Text, ",102,") > 0 Or InStr(strCP10Text, ",103,") > 0 Then
                '新增"202"
                strMaxNo = GetMaxNo(strItemNo)
                'Modify by  Morgan 2004/12/17 要計算折扣
                'adoTaie.Execute "insert into acc1l0 values ('" & strItemNo & "', '" & strMaxNo & "', '" & text7.text & "', '202', 1600, null, 0, '" & strUserNum & "', " & Val(FCDate(ServerDate)) & ", " & ServerTime & ", null, null, null)"
                adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & Text7.Text & "', '202', 1600, null, " & 1600 * Val(stDisc) / 100 & " , '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
                '新增"106"
                strMaxNo = GetMaxNo(strItemNo)
                'Modify by  Morgan 2004/12/17 要計算折扣
                'adoTaie.Execute "insert into acc1l0 values ('" & strItemNo & "', '" & strMaxNo & "', '" & text7.text & "', '106', 2500, null, 0, '" & strUserNum & "', " & Val(FCDate(ServerDate)) & ", " & ServerTime & ", null, null, null)"
                '2007/1/2 MODIFY BY SONIA 調服務費2500->3000
                'adoTaie.Execute "insert into acc1l0 values ('" & strItemNo & "', '" & strMaxNo & "', '" & text7.text & "', '106', 2500, null, " & 2500 * Val(stDisc) / 100 & ", '" & strUserNum & "', " & strSrvDate(2) & ", " & ServerTime & ", null, null, null)"
                adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & Text7.Text & "', '106', 3000, null, " & 3000 * Val(stDisc) / 100 & ", '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
                '2007/1/2 END
           End If
           
           If InStr(strCP10Text, ",201,") > 0 Or InStr(strCP10Text, ",209,") > 0 Or InStr(strCP10Text, ",210,") > 0 Or InStr(strCP10Text, ",223,") > 0 Then
            '92.11.24 END
                '新增"03"
                strMaxNo = GetMaxNo(strItemNo)
                adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & Text7.Text & "', '03', 0, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
            End If
            'end 2020/2/3
            
            '若非領證, 年費
            'Modified by Lydia 2016/09/07 FCP案的核對已准專利926,預設不加傳真和雜費
            If InStr(strCP10Text, ",601,") <= 0 And InStr(strCP10Text, ",605,") <= 0 And InStr(strCP10Text, ",926,") <= 0 Then
               '新增"01" 傳真
               'Removed by Morgan 2022/7/19 改不帶出--Phoebe
               'strMaxNo = GetMaxNo(strItemNo)
               'adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & Text7.Text & "', '01', 0, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
               'end 2022/7/19
               
               If ChkItemFCP02() = True Then 'Added by Morgan 2025/2/10
                  '新增"02" 雜費
                  strMaxNo = GetMaxNo(strItemNo)
                  adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & Text7.Text & "', '02', 0, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
               End If
            End If
            '2010/10/25 ADD BY SONIA 申請人X55778且請款程序為"核對已准專利"926時預設請款備註
            If CUISX55778 = True And InStr(strCP10Text, ",926,") > 0 Then
               CP10have926 = True
            End If
            '2010/10/25 END
        Case "FCT"
            '若有商申
            If InStr(strCP10Text, ",101,") > 0 Then
                
                'Remove by Morgan 2010/3/1 移到上面接在第1個規費後面
                ''Add By Cheng 2003/12/11
                ''依(商品類別數-1)預設規費項目(10199)數
                'intCnt = GetTMKindCnt(text7.text, text21.text, text22.text, text23.text) - 1
                'For ii = 1 To intCnt
                '    strMaxNo = GetMaxNo(strItemNo)
                '    'Modify By Sindy 2010/01/28 改為每個類別之規費金額先預設為3000元,剩下的金額則掛在第1個規費
                '    'adoTaie.Execute "insert into acc1l0 values ('" & strItemNo & "', '" & strMaxNo & "', '" & text7.text & "', '10199', 0, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ", " & ServerTime & ", null, null, null)"
                '    adoTaie.Execute "insert into acc1l0 values ('" & strItemNo & "', '" & strMaxNo & "', '" & text7.text & "', '10199', 3000, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ", " & ServerTime & ", null, null, null)"
                'Next ii
                ''End
                'end 2010/3/1
               
               If dblFCT10101 > 0 Then 'Add by Morgan 2010/7/1
                  '新增"01"
                  strMaxNo = GetMaxNo(strItemNo)
                  adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & Text7.Text & "', '01', " & dblFCT10101 & ", null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
               End If 'Add by Morgan 2010/7/1
               
                '新增"02"
                strMaxNo = GetMaxNo(strItemNo)
                'Added by Morgan 2020/3/6
                '日本區商申 雜費：500, 第1類：9000, 第2類以上：各7000 ; 折扣預設相同由收文金額推算--湘嫻
                If Left(strFA10, 3) = "011" Then
                  adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & Text7.Text & "', '02', 500, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
                Else
                'end 2020/3/6
                
                  adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & Text7.Text & "', '02', 600, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
                  
                End If 'Added by Morgan 2020/3/6
            
            'Add by Morgan 2004/10/4 註冊費依類別數產生多筆項目 715->5000, 716->1500, 717->2500
            ElseIf InStr(strCP10Text, ",715,") > 0 Or InStr(strCP10Text, ",716,") > 0 Or InStr(strCP10Text, ",717,") > 0 Then
               If InStr(strCP10Text, ",715,") > 0 Then
                  stA1L04 = "71599"
                  stA1L05 = "1000"
               ElseIf InStr(strCP10Text, ",716,") > 0 Then
                  stA1L04 = "71699"
                  stA1L05 = "1500"
               Else
                  stA1L04 = "71799"
                  stA1L05 = "2500"
               End If
                  
                '依(商品類別數-1)預設規費項目(10199)數
                'intCnt 服務費已設定過,次處不再重抓
                For ii = 1 To intCnt
                    strMaxNo = GetMaxNo(strItemNo)
                    adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & Text7.Text & "', '" & stA1L04 & "'," & stA1L05 & ", null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
                Next ii
                
               '新增"02"
               strMaxNo = GetMaxNo(strItemNo)
               'Modified by Morgan 2019/1/28 FCT日本區「註冊費」(717)請款 => 雜費:700
               'adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & Text7.Text & "', '02', 500, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
               stA1L05 = "500"
               If InStr(strCP10Text, ",717,") > 0 And Left(strFA10, 3) = "011" Then
                  stA1L05 = "700"
               End If
               adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & Text7.Text & "', '02', " & stA1L05 & ", null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
               'end 2019/1/28
               
            '2004/10/4 end
            
            '2006/5/18 ADD BY SONIA 延展加 02 雜費600元
            ElseIf InStr(strCP10Text, ",102,") > 0 Then
                '新增"02"
                strMaxNo = GetMaxNo(strItemNo)
                'Modified by Morgan 2014/8/12 延展雜項改預設 500 --陳金蓮
                'Modified by Morgan 2014/8/15 延展雜項改回預設 600 --陳金蓮
                'Modified by Morgan 2014/12/2 延展雜費再改 500 --陳金蓮
                adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & Text7.Text & "', '02', 500, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
            '2006/5/18 END
            
            'Add by Morgan 2010/8/26 +501移轉 -- 陳金蓮
            ElseIf InStr(strCP10Text, ",501,") > 0 Then
               '新增"02"
               strMaxNo = GetMaxNo(strItemNo)
               adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & Text7.Text & "', '02', 600, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
            
            'Added by Morgan 2018/10/9
            'FCT「補正」之日文請款項目預設為201=扣掉(02)雜費後之金額,02=500 -- 陳金蓮
            ElseIf Left(strFA10, 3) = "011" And InStr(strCP10Text, ",201,") > 0 Then
                '新增"02"
                strMaxNo = GetMaxNo(strItemNo)
                adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & Text7.Text & "', '02', 500, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
            End If
            
            '抓核准案件性質
            stA1L04 = ""
            Select Case GetRelCaseProperty(strItemNo, "1001")
            Case "102" '延展
               '新增"1023"
               stA1L04 = "1023"
            Case "103" '補發註冊證
               '新增"1031"
               stA1L04 = "1031"
            Case "301" '變更
               '新增"3012"
               stA1L04 = "3012"
            '2007/6/7 ADD BY SONIA
            Case "313" '減縮商品
               '新增"3132"
               stA1L04 = "3132"
            '2007/6/7 END
            Case "501" '移轉
                '新增"5012"
                stA1L04 = "5012"
            Case "502" '授權
                '新增"5022"
                stA1L04 = "5022"
            Case "504" '再授權
                '新增"5023"
                stA1L04 = "5023"
            Case "506" '設定質權
                '新增"5061"
                stA1L04 = "5061"
            End Select
            'End
            If stA1L04 <> "" Then
               '新增
               strMaxNo = GetMaxNo(strItemNo)
               'Modify by Morgan 2007/5/23 加折扣
               'adoTaie.Execute "insert into acc1l0 values ('" & strItemNo & "', '" & strMaxNo & "', '" & text7.text & "', '" & stA1L04 & "', 3000, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ", " & ServerTime & ", null, null, null)"
               adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & Text7.Text & "', '" & stA1L04 & "', 3000, null, " & 3000 * Val(stDisc) / 100 & ", '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
               'end 2007/5/23
               '新增"02"
               strMaxNo = GetMaxNo(strItemNo)
               adoTaie.Execute "insert into acc1l0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L10,A1L08,A1L09) values ('" & strItemNo & "', '" & strMaxNo & "', '" & Text7.Text & "', '02', 500, null, 0, '" & strUserNum & "', " & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'))"
            End If
        End Select
   End If
   adocheck.Close
   
'2009/6/24 開放使用
'   If Pub_StrUserSt03 = "M51" Then
'      Combo3.Enabled = True
'   Else
'      Combo3.Enabled = False
'   End If
   '2009/4/23 END
   
   'Removed by Morgan 2014/8/6 Form要共用,原呼叫畫面的資料改直接以acc1k0設定,故移到最上面
   'adoacc1k0.CursorLocation = adUseClient
   'adoacc1k0.Open "select * from acc1k0 where a1k01 = '" & strItemNo & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'end 2014/8/6
   adoacc1l0.CursorLocation = adUseClient
   adoacc1l0.Open "select * from acc1l0 where a1l01 = '" & strItemNo & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   FormShow 'Modified by Morgan 2016/6/27 恢復
   strCustNo = Left(Text8, 8) 'Add by Amy 2013/10/30 記錄代理人
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc1l0, acc1j0 where a1l03 = a1j01 and a1l04 = a1j02 and a1l01 = '" & Text1 & "' order by a1l02 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
   
   'Add by Amy 2025/11/12 由結案單進入更新703/704 開頭及2碼
   '只有請別道,仍保留703/704 當道讓user 自行刪除(因frmacc21H0只能選703/704 之總收文號)-秀玲
   ' ex:X11400533 FCT-052536 銷[612-補充理由]期限,但請款代號:614
   If stF0301 <> "" Then
      If stOldCP10 <> "" Then
         If ChkAndSetCCDItem("3", stOldCP09, stOldCP10, strItemNo, Mid(stCP10Item, 2), stTpMsg) = False Then
            stAmtNotSame(0) = Mid(stAmtNotSame(0) & "," & stTpMsg, 2)
         Else
            AdodcRefresh
         End If
      End If
      '最後總金額是否相同
      If ChkCCDAndAcc1L0NotSame("0", Me.Name, stF0301, strItemNo, stTpMsg) = True Then
         stAmtNotSame(1) = Replace(Mid(stAmtNotSame(1) & "," & stTpMsg, 2), ";", vbCrLf)
      End If
      '訊息
      If stAmtNotSame(0) = "" And stAmtNotSame(1) = "" Then
         stShowMsg = "結案單請款項目已新增完畢,請確認！"
      Else
         If stAmtNotSame(0) <> "" Then stShowMsg = Replace(Mid(stAmtNotSame(0), 2), ",", vbCrLf)
         If stAmtNotSame(1) <> "" Then
            If stShowMsg <> "" Then stShowMsg = stShowMsg & vbCrLf
            stShowMsg = stShowMsg & Replace(Mid(stAmtNotSame(1), 2), ",", vbCrLf)
         End If
         stShowMsg = "結案單請項目與目前寫入之請款項目不同如下：" & vbCrLf & _
                                    stShowMsg & vbCrLf & "請確認！"
      End If
      '*** 抓取目前設定,有不同彈訊息 ***
      Call Pub_GetCloseA1KData(3, Me.Name, Text7, Text21, Text22, Text23, o_A1K29, stNP07, o_A1K04, o_A1K27, o_A1K28, o_TM56, o_TM69, stF0301, o_CCM20, stTpMsg)
      If o_CCM20 <> "" Then
         Text4 = o_CCM20
      End If
      If stTpMsg <> "" Then
         stRemindMsg = "此結案單有設定請款單資訊如下：" & Replace(stTpMsg, ";", vbCrLf) & vbCrLf & _
                                       "請確認！"
      End If
      '*** End 抓取目前設定,有不同彈訊息 ***
   End If
   'end 2025/11/12
   
Checking:
   If Err.Number = 0 Then
      OpenTable = True
      Exit Function
   End If
   MsgBox Err.Description, , MsgText(5)
   OpenTable = False
End Function

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc1l0, acc1j0 where acc1l0.a1l03 = acc1j0.a1j01 and acc1l0.a1l04 = acc1j0.a1j02 and a1l01 = '" & Text1 & "' order by a1l02 asc", adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.ReQuery
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
   Dim iPrintCurrType As Integer
   
   Text1 = strItemNo
   'Add By Sindy 2009/09/29
   doua1k11 = 0
   If Not IsNull(adoacc1k0.Fields("a1k11").Value) Then
      doua1k11 = adoacc1k0.Fields("a1k11").Value
   End If
   '2009/09/29 End
   MaskEdBox1.Mask = MsgText(601)
   
   'Modified by Morgan 2025/8/20
   '請款對象為Y56199 Coupang Corp.的所有帳單，請款日皆為當月16號--Kahn
   '系統自動產生的請款單在Trigger(ACC1K0_BEFORE)更新請款日
   'If IsNull(adoacc1k0.Fields("a1k02").Value) Then
   '   MaskEdBox1.Text = CFDate(Val(strSrvDate(2)))
   'Else
   '   MaskEdBox1.Text = CFDate(adoacc1k0.Fields("a1k02").Value)
   'End If
   strExc(2) = "" & adoacc1k0("a1k02")
   If strExc(2) = "" Then strExc(2) = strSrvDate(2)
   If Text8 = "Y56199000" And strExc(2) = adoacc1k0("a1k19") And adoacc1k0("a1k19") = strSrvDate(2) Then
      strExc(1) = Val(strExc(2)) \ 100 & "16"
      If strExc(1) < strExc(2) Then
         strExc(1) = TransDate(CompDate(1, 1, strExc(1)), 1)
      End If
      If strExc(2) <> strExc(1) Then
         strExc(2) = strExc(1)
         MsgBox "請款對象為Y56199 Coupang Corp.的所有帳單，當月16號以前 (含當天)產生的帳單，請款日自動設定當月16號，當月16號以後 (至月底)產生的帳單，請款日自動設定隔月的16號！", vbInformation
      End If
   End If
   MaskEdBox1.Text = CFDate(Val(strExc(2)))
   'end 2025/8/20
   
   strOriDate = MaskEdBox1.Text
   MaskEdBox1.Mask = DFormat
   If IsNull(adoacc1k0.Fields("a1k03").Value) Then
      If strCon1 <> "" Then
         If Len(strCon1) = 6 Then
            Text2 = strCon1 & "000"
         Else
            Text2 = strCon1
         End If
      Else
         Text2 = MsgText(601)
      End If
   Else
      If Len(adoacc1k0.Fields("a1k03").Value) = 6 Then
         Text2 = adoacc1k0.Fields("a1k03").Value & "000"
      Else
         Text2 = adoacc1k0.Fields("a1k03").Value
      End If
   End If
   Select Case Mid(Text2, 1, 1)
      Case "Y"
         Text3 = FagentQuery(Text2, 2)
         '2005/7/8 ADD BY SONIA
         If Text3 = "" Then
            Text3 = FagentQuery(Text2, 1)
         End If
         If Text3 = "" Then
            Text3 = FagentQuery(Text2, 3)
         End If
         '2005/7/8 END
      Case "X"
         Text3 = CustomerQuery(Text2, 2)
         '2005/7/8 ADD BY SONIA
         If Text3 = "" Then
            Text3 = CustomerQuery(Text2, 1)
         End If
         If Text3 = "" Then
            Text3 = CustomerQuery(Text2, 3)
         End If
         '2005/7/8 END
   End Select
'   If IsNull(adoacc1k0.Fields("a1k18").Value) Then
'      Combo3 = MsgText(601)
'   Else
'      Combo3 = adoacc1k0.Fields("a1k18").Value
'   End If
   
   If IsNull(adoacc1k0.Fields("a1k05").Value) Then
      Text11 = MsgText(601)
   Else
      Text11 = adoacc1k0.Fields("a1k05").Value
   End If
   Text24 = "" & adoacc1k0.Fields("a1k34").Value 'Added by Morgan 2013/10/18
   
'Removed by Morgan 2014/8/6 移到 OpenTable
'   Text7 = adoacc1k0.Fields("a1k13").Value
'   Text21 = adoacc1k0.Fields("a1k14").Value
'   Text22 = adoacc1k0.Fields("a1k15").Value
'   Text23 = adoacc1k0.Fields("a1k16").Value
''取得案件性質
'   m_strCP10 = GetCP10("" & Frmacc21h0.Adodc2.Recordset("CP09").Value)
'end 2014/8/6

   Text5 = "" & adoacc1k0("a1k32") 'Add by Morgan 2010/5/18
   'Modified by Lydia 2015/04/15 整批請款單時,a1k32不可修改
   If Text5 = "C" Then
      Text5.Enabled = False
   Else
      Text5.Enabled = True
   End If
    
    'Add By Sindy 2013/1/24
    '預設請款幣別及列印幣別格式及請款匯款
    If IsNull(adoacc1k0.Fields("a1k18")) Or "" & adoacc1k0.Fields("a1k18") = "" Then
      '依請款對象抓請款幣別
      'Modify By Sindy 2016/12/16 + , Text21, Text22, Text23
      'Call PUB_GetDefaultCurrPrintType(Text7, Text8, "", strA1K18)
      'Modified by Morgan 2018/4/27 +傳列印對象,同時讀取列印幣別格式
      'Call PUB_GetDefaultCurrPrintType(Text7, Text8, "", strA1K18, Text21, Text22, Text23, Text6)
      iPrintCurrType = PUB_GetDefaultCurrPrintType(Text7, Text8, "", strA1K18, Text21, Text22, Text23, Text6) - 1
      'end 2018/4/27
      
      'Modified by Morgan 2018/6/29 取消限制,IPSIDE Y19357B30 P案要用台幣請款--David
      ''FMP不可為NTD或RMB
      'If bolIsFMP = True Then
      '   If strA1K18 = "NTD" Or strA1K18 = "RMB" Then
      '      strA1K18 = ""
      '   End If
      'End If
      'end 2018/6/29
      
      '抓請款匯率
      If strA1K18 <> "" Then
         Combo3 = strA1K18
         adoacc1k0.Fields("a1k18").Value = strA1K18
         'Modified by Morgan 2024/1/19 改函數以便共用
         ''Added by Morgan 2021/8/10 +Y51345北京正理商標事務所並預設匯率為 1/0.036=27.77(兩位後面捨去) -- 桂英
         'If Text8 = "Y51345000" And strA1K18 = "USD" Then
         '   Text19 = 27.77
         ''Added by Morgan 2022/11/8
         'ElseIf m_AppNo = "X55070010" And strA1K18 = "USD" And Text7 = "FCT" Then
         '   Text19 = 26.5
         ''Added by Morgan 2024/1/17 加坡代理人SPRUSON & FERGUSON (AISA) PTE LTD (Y21071)商標案件固定請款匯率為USD1.00 = NTD28.39 --洪琬姿
         'ElseIf Text8 = "Y21071000" And Text7 = "FCT" And DBDATE(MaskEdBox1) <= "20241231" Then
         '   Text19 = 28.39
         'Else
         ''end 2021/8/10
         '   Text19 = PUB_GetUSXRate_1(Val(adoacc1k0.Fields("a1k02").Value), strA1K18)
         'End If 'Added by Morgan 2021/8/10
         Text19 = PUB_GetRate(Val(adoacc1k0.Fields("a1k02").Value), strA1K18, Text8, Text7, m_AppNo)
         'end 2024/1/19
         adoacc1k0.Fields("a1k10").Value = Text19
      End If
      '依列印對象抓列印幣別格式
      If bolIsFMP = False Then
         'Modify By Sindy 2016/12/16 + , , Text21, Text22, Text23
         'Combo4.ListIndex = PUB_GetDefaultCurrPrintType(Text7, Text6, Combo3) - 1
         'Modified by Morgan 2018/4/27
         'Combo4.ListIndex = PUB_GetDefaultCurrPrintType(Text7, Text6, Combo3, , Text21, Text22, Text23) - 1
         Combo4.ListIndex = iPrintCurrType
         'end 2018/4/27
      Else
         'FMP列印幣別格式固定為3.純外幣
         If iPrintCurrType <> 0 Then 'Added by Morgan 2018/6/29 取消 NTD 限制,IPSIDE Y19357B30 P案要用台幣請款--David
            Combo4.ListIndex = 2
            Combo4.Enabled = False
         End If
      End If
      adoacc1k0.Fields("a1k33").Value = Combo4.ListIndex + 1
    Else
      Text19 = "" & adoacc1k0.Fields("a1k10").Value 'Add By Sindy 2015/2/25 顯示DB資料
      Combo3 = adoacc1k0.Fields("a1k18").Value
    End If
    'Modify By Sindy 2013/1/24 從上頭程式段Move至此
    If Text8 = "Y48292000" Then Text19 = adoacc1k0.Fields("a1k10"): dblRate = Val(Text19)  'Added by Morgan 2012/9/18 HP用報價匯率
    '2013/1/24 End
    
    SetText19 'Added by Morgan 2012/9/18
    
    '是否列印申請人
    '若無資料(新增)
    'Modified by Morgan 2014/8/18 考慮整批請款
    'If m_blnAcc1l0NoData = True Then
    If m_blnAcc1l0NoData = True And m_bolIsBatch = False Then
    'end 2014/8/18
        'Modify by Morgan 2004/12/16
        'Me.Text4.Text = PUB_GetA1K04("" & adoacc1k0.Fields("a1k13").Value, "" & adoacc1k0.Fields("a1k14").Value, "" & adoacc1k0.Fields("a1k15").Value, "" & adoacc1k0.Fields("a1k16").Value)
        Me.Text4.Text = PUB_GetA1K04("" & adoacc1k0.Fields("a1k13").Value, "" & adoacc1k0.Fields("a1k14").Value, "" & adoacc1k0.Fields("a1k15").Value, "" & adoacc1k0.Fields("a1k16").Value, Me.Text8.Text, m_strCP10)
    '若有資料(修改)
    Else
        If IsNull(adoacc1k0.Fields("a1k04").Value) Then
           Text4 = MsgText(601)
        Else
           Text4 = adoacc1k0.Fields("a1k04").Value
        End If
    End If
    
    '取得折扣
    'Added by Morgan 2014/8/18 考慮整批請款
    If m_bolIsBatch Then
      m_strDisc = 100 - Val(m_Discount)
    Else
    'end 2014/8/18
      m_strDisc = PUB_GetA1L07Disc("" & adoacc1k0.Fields("a1k13").Value, "" & adoacc1k0.Fields("a1k14").Value, "" & adoacc1k0.Fields("a1k15").Value, "" & adoacc1k0.Fields("a1k16").Value, m_strCP10, Replace(Me.MaskEdBox1.Text, "/", ""))
    End If 'Added by Morgan 2014/8/18
    
    If m_strDisc = "100" Then m_strDisc = ""
    
    '2010/10/25 ADD BY SONIA 申請人X55778提示要輸入請款備註
    If CUISX55778 = True Then
       MsgBox "申請人為X55778(NIPPON SODA CO., LTD.), 請輸入請款單備註！", vbExclamation + vbOKOnly
       If CP10have926 Then SetRemark Text7.Text, "X55778000", Text11
    End If
    '2010/10/25 END
    
    'Add by Morgan 2010/9/10
    If m_blnAcc1l0NoData = True And Text11.Text = "" Then
      SetRemark Text7.Text, Text8.Text, Text11
    End If
    
End Sub

'Added by Morgan 2012/9/18
'HP用報價匯率
Private Sub SetText19()
   m_bolChkDate = False 'Added by Morgan 2024/1/19
   
   'Modified by Morgan 2015/11/18 +Y54332000
   'Modified by Morgan 2019/10/3 +Y53475(ADASTRA Intellectual Property Sdn Bhd)--Tim
   'Modified by Morgan 2019/10/4 +Y54975 (YAKIMA (NAN JING) PRECISION INDUSTRY CO., LTD.)--Ali
   'Modified by Morgan 2020/6/24 +Y55294 --Ryan
   'Modified by Morgan 2020/10/12 +Y45622 --郭怡瑩
   'Modified by Morgan 2021/3/18 +X82693,X83843,X71117,X7111702,X7111703 -- 潘韻丞
   'Modified by Morgan 2021/3/26 +m_AppNo <> ""
   'Modified by Morgan 2021/8/10 +Y51345北京正理商標事務所並預設匯率為 1/0.036=27.78 -- 桂英
   'Modified by Morgan 2021/8/26 +X76135,Y54600 --Monica
   'Modified by Morgan 2021/10/1 +Y55540 --Tim
   'Modified by Morgan 2022/5/20 +Y54339,Y54339B10,Y54339B20 --Tim
   'Modified by Morgan 2022/9/23 +X86352 -- 潘韻丞
   'Modified by Morgan 2024/9/16 +Y25061及其關聯編號、Y54470、Y54469、Y55363 --沈佳穎
   'Modified by Morgan 2025/6/24 +Y56059,Y5605901 --韻丞
   If Text8 <> "" And (InStr("Y48292000,Y54332000,Y53475000,Y54975000,Y55294000,Y45622000,Y51345000,Y54600000,Y55540000,Y54339000,Y54339B10,Y54339B20,Y54470000,Y54469000,Y55363000,Y56059000,Y56059010", Text8) > 0 Or Left(Text8, 6) = "Y25061") Then
      Text19.Enabled = True: Text19.BackColor = Text8.BackColor
      Text19.Tag = "Y"
   ElseIf (m_AppNo <> "" And InStr("X82693000,X83843000,X71117000,X71117020,X71117030,X76135000,X86352000", m_AppNo) > 0) Then
      Text19.Enabled = True: Text19.BackColor = Text8.BackColor
      Text19.Tag = "X"
   'Added by Morgan 2022/11/8 +X55070010 --黃咸達/孫季仙
   ElseIf m_AppNo = "X55070010" And Text7 = "FCT" Then
      Text19.Enabled = True: Text19.BackColor = Text8.BackColor
      Text19.Tag = "X"
   'end 2022/11/8
   'Modified by Morgan 2024/1/19 --洪琬姿
   'Modified by Morgan 2024/5/28 +S --宜儒
   ElseIf Text8 = "Y21071000" And (Text7 = "FCT" Or Text7 = "S") And DBDATE(MaskEdBox1) >= "20240117" And DBDATE(MaskEdBox1) <= "20241231" Then
      Text19.Enabled = True: Text19.BackColor = Text8.BackColor
      Text19.Tag = "Y"
      m_bolChkDate = True
      
   'Added by Morgan 2025/3/25 FCP-071354 以客戶要求的付款匯率進行請款--Kahn
   ElseIf Text7 & Text21 & Text22 & Text23 = "FCP071354000" Then
      Text19.Enabled = True: Text19.BackColor = Text8.BackColor
      Text19.Tag = Text7 & "-" & Text21 & IIf(Text22 & Text23 = "000", "", "-" & Text22 & "-" & Text23)
      m_bolChkDate = True
   'end 2025/3/25
   
   'Added by Morgan 2025/7/3 FCT案請款單日本區所有Ｙ編號開放可手動調整匯率--May
   'Modified by Morgan 2025/7/10 +CFT也要(改判斷外商部門)--Sasa/May
   ElseIf Left(strFA10, 3) = "011" And Left(Pub_StrUserSt03, 2) = "F1" Then
      Text19.Enabled = True: Text19.BackColor = Text8.BackColor
      Text19.Tag = ""
      m_bolChkDate = True
   'end 2025/7/3
   Else
   'end 2024/1/19
      Text19.Enabled = False: Text19.BackColor = Text12.BackColor
   End If
   
   'Added by Morgan 2012/9/18
   If Text19.Enabled = True Then
      'Modified by Morgan 2022/11/8
      'MsgBox Text8 & " 請款匯率特別請留意！", vbExclamation
      If Text19.Tag = "Y" Then
         MsgBox "請款對象 " & Text8 & FagentQuery(Text8, 2) & " 請款匯率特別，請留意！", vbExclamation
      ElseIf Text19.Tag = "X" Then
         MsgBox "客戶 " & m_AppNo & CustomerQuery(m_AppNo, 2) & " 請款匯率特別，請留意！", vbExclamation
      ElseIf Text19.Tag <> "" Then
         MsgBox Text19.Tag & " 請款匯率特別，請留意！", vbExclamation
      End If
      'end 2022/11/8
   End If
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox Label2 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label2 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   'Add By Cheng 2004/01/28
   '若更改請款日期
   If Me.MaskEdBox1.Text <> Me.MaskEdBox1.Tag Then
      'Modified by Morgan 2024/1/19
      'If Text19.Enabled = False Then 'Added by Morgan 2012/9/18
      If Text19.Enabled = False Or m_bolChkDate Then
      'end 2024/1/19
         'Modify By Sindy 2015/2/25 增加詢問
         If MsgBox("是否要重抓新日期的匯率？", vbYesNo + vbDefaultButton1) = vbYes Then
         '2015/2/25 END
            '2009/4/23 MODIFY BY SONIA 重抓請款匯率
            'dblRate = PUB_GetUSXRate_1(Replace(Me.MaskEdBox1.Text, "/", ""))
            'Modified by Morgan 2024/1/19
            'dblRate = PUB_GetUSXRate_1(Replace(Me.MaskEdBox1.Text, "/", ""), Me.Combo3.Text)
            SetText19
            dblRate = PUB_GetRate(Replace(Me.MaskEdBox1.Text, "/", ""), Me.Combo3.Text, Text8, Text7, m_AppNo)
            'end 2024/1/19
            '2009/4/23 END
            Text19 = dblRate 'Added by Morgan 2012/9/18
         End If
       End If
       SumShow
       Me.MaskEdBox1.Tag = Me.MaskEdBox1.Text
       m_blnClkPrintButton = False    '2009/10/13 ADD BY SONIA否則印完改日期再印會錯
   End If
   'End
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add by Morgan 2011/3/22
Private Sub Text10_Change()
   If Text7 = "FCT" And (Text16 = "10199" Or Text16 = "A0199") Then
      Text18 = GetFCT10199Fee
   End If
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub
'end 2011/3/23

Private Sub Text11_GotFocus()
   TextInverse Text11
   'add by nickc 2007/07/13 將輸入法改成使用API
   OpenIme
End Sub

'add by nickc 2007/07/13 將輸入法改成使用API
Private Sub Text11_Validate(Cancel As Boolean)
   CloseIme
End Sub

Private Sub Text15_GotFocus()
   TextInverse Text15
End Sub

Private Sub Text16_GotFocus()
   TextInverse Text16
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text16_Validate(Cancel As Boolean)
   
   If Text16.Tag = Text16 And Val(Text18) > 0 Then Exit Sub 'Added by Morgan 2018/7/19 已有金額且請款項目沒有改維持舊資料
   
   If Text16 = MsgText(601) Then
      Exit Sub
   End If
   
   If ExistCheck("acc1j0", "a1j01 || a1j02", Text7 & Text16, Label13) = False Then
      MsgBox MsgText(28) & Label13, , MsgText(5)
      Cancel = True
      Exit Sub
   End If
   
   Text17 = A1j03Query(Text7, Text16)
   
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select a1j17 from acc1j0 where a1j01 = '" & Text7 & "' and a1j02 = '" & Text16 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text18 = MsgText(601)
      Else
         Text18 = adoaccsum.Fields(0).Value
      End If
   Else
      Text18 = MsgText(601)
   End If
   adoaccsum.Close
   
   If Mid(Text16, 1, 2) = "99" Then
      Text20 = "0"
   'Added by Morgan 2011/10/31 雜費預設不折扣--陳金蓮
   ElseIf (Text7 = "FCT" Or Text7 = "CFT" Or Text7 = "S") And Text16 = "02" Then
      Text20 = "0"
   'end 2011/10/31
   
   'Added by Morgan 2018/12/18 傳真01,雜費02預設不折扣--吳若芬
   ElseIf (Text7 = "FCP" Or Text7 = "FG" Or bolIsFMP) And (Text16 = "01" Or Text16 = "02") Then
      Text20 = "0"
   'end 2018/12/18
'Removed by Morgan 2017/9/30 FMP案改回預設--David
'   'Added by Lydia 2016/09/07 FMP案不預設折扣
'   ElseIf bolIsFMP Then
'      Text20 = "0"
'   'end 2016/09/07
   
   Else
'      Text20 = Val(DiscountShow(Text7, Text21, Text22, Text23)) * Val(Text18) / 100
      'Modify by Morgan 2005/1/10 帶出折扣數,不是金額
      'Text20 = Val(m_strDisc) * Val(Text18) / 100
      'Modified by Morgan 2013/5/1 代理人+案件性質特殊折扣
      'Text20 = m_strDisc
      'Modified by Morgan 2015/8/5 +申請人 m_AppNo,有值不預設
      'Text20 = 100 * PUB_GetDiscX(Text2, Text7, Text16, IIf(Val(m_strDisc) = 0, 1, Val(m_strDisc) / 100))
      If Text20 = "" Or Text20 = "0" Or Text20 = "100" Then
         Text20 = 100 * PUB_GetDiscX(Text2, Text7, Text16, IIf(Val(m_strDisc) = 0, 1, Val(m_strDisc) / 100), m_AppNo)
      End If
      'end 2013/5/1
   End If
   
   'Add by Morgan 2011/3/2
   'Modified by Morgan 2015/11/19 證明標章(系統商標種類:7)及團體標章(系統商標種類:8)除外--陳金蓮
   'Modified by Morgan 2020/1/9 +1013(超過商品數)--陳金蓮
   If Text7 = "FCT" And (Text16 = "10199" Or Text16 = "A0199" Or Text16 = "1013") And m_strTM08 <> "7" And m_strTM08 <> "8" Then
      Text9.TabStop = True
      If Text10.Enabled Then Text10.TabStop = True
   'Added by Morgan 2018/11/20
   ElseIf Text7 = "S" And Text16 = "0011" Then
      Text9.TabStop = True
   'end 2018/11/20
   Else
      Text9.TabStop = False
      If Text10.Enabled Then Text10.TabStop = False
   End If
   
   'Add By Sindy 2013/1/24
   If Right(Trim(Text16), 2) = "99" Or Right(Trim(Text16), 2) = "98" Then
      Text20.Enabled = False
      Text20.Text = ""
'      'Add By Sindy 2013/4/17
'      If Right(Trim(Text16), 2) = "98" Then
'         strExc(0) = "select * from acc1L0 where a1L01='" & Text1 & "' and a1L04='" & Left(Trim(Text16), Len(Trim(Text16)) - 2) & "' and a1L05 is not null"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 0 Or RsTemp.RecordCount = 0 Then
'            MsgBox "請先輸入本所服務費資料後，才能輸入代收代付！", , MsgText(5)
'            Cancel = True
'            Exit Sub
'         End If
'      End If
'      '2013/4/17 End
   Else
      Text20.Enabled = True
   End If
   If bolIsFMP = True Then Call SetCombo5
   '2013/1/24 End
End Sub

'Add By Sindy 2013/1/24
Private Sub SetCombo5(Optional pOldCurrrency As String)
Dim i As Integer
   
   If Trim(Combo3.Text) = "" Then
      MsgBox "請款幣別不可空白！"
      Exit Sub
   End If
   If bolIsFMP = True Then
      Text13.Enabled = False '輸入RMB金額
      Text13.Text = ""
      '大陸官方規費及代理人服務費只能輸入RMB或USD
      '但是二者的輸入幣別必須相同
      If Right(Trim(Text16), 2) = "99" Or Right(Trim(Text16), 2) = "98" Then
         strCurr = ""
         If Adodc1.Recordset.RecordCount > 0 Then
            strExc(0) = "select * from acc1L0 where a1L01='" & Text1 & "' and (substr(a1L04,-2)='99' or substr(a1L04,-2)='98') and a1L16 is not null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp.RecordCount > 0 Then
                  strCurr = "" & RsTemp.Fields("a1L16")
               End If
            End If
         End If
         Combo5.Clear
         If strCurr <> "" Then
            Combo5.AddItem strCurr
            Combo5.ListIndex = 0
            Combo5.Enabled = False
            If strCurr = "USD" And Right(Trim(Text16), 2) = "99" Then
               Text13.Enabled = True
            End If
         Else
            Combo5.AddItem "RMB"
            Combo5.AddItem "USD"
            Combo5.ListIndex = 1
            Combo5.Enabled = True
         End If
      Else '本所服務費只可輸入NTD或請款幣別
         Combo5.Clear
         Combo5.AddItem "NTD"
         Combo5.AddItem Combo3
         Combo5.ListIndex = 1
         Combo5.Enabled = True
         'Added by Morgan 2016/8/1
         '請款幣別變更要新增原請款幣別否則會當
         If pOldCurrrency <> "" And pOldCurrrency <> Combo3 And pOldCurrrency <> "NTD" Then
            Combo5.AddItem pOldCurrrency
         End If
         'end 2016/8/1
      End If
   End If
End Sub

Private Sub Text18_GotFocus()
   TextInverse Text18
End Sub

Private Sub Text18_Validate(Cancel As Boolean)
   If Text16 = MsgText(601) Then
      Exit Sub
   End If
   adocheck.CursorLocation = adUseClient
   adocheck.Open "select a1j07, a1j08 from acc1j0 where a1j01 = '" & Text7 & "' and a1j02 = '" & Text16 & "' and ((a1j07 >= " & Val(Text18) & " and a1j08 <= " & Val(Text18) & ") or (a1j07 = 0 and a1j08 = 0))", adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount = 0 Then
      MsgBox MsgText(59), , MsgText(5)
      Cancel = True
   End If
   adocheck.Close
End Sub

Private Sub Text19_GotFocus()
   TextInverse Text19
End Sub

Private Sub Text19_Validate(Cancel As Boolean)
   'Added by Morgan 2021/10/1
   If Val(Text19) = 0 And Text19.Enabled Then
      MsgBox "請輸入請款匯率！", vbExclamation
      Cancel = True
      Exit Sub
   End If
   'end 2021/10/1
   
   If dblRate <> Val(Text19) Then
      dblRate = Val(Text19)
      SumShow
      m_blnClkPrintButton = False
   End If
End Sub

Private Sub Text20_GotFocus()
   TextInverse Text20
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  儲存資料表(國外請款資料(交易檔))
'
'*************************************************
Private Sub Acc1l0Save()
On Error GoTo Checking
      
      'Add By Sindy 2025/4/10
      If bolIsFMP = False And Combo5.Text <> "NTD" Then
         If Combo5.Text <> Combo3.Text Then
            MsgBox "輸入幣別和請款幣別要一致!", , MsgText(5)
            strControlButton = MsgText(602)
            Combo5.SetFocus
            Exit Sub
         End If
      End If
      '2025/4/10 END
      
      If Text15 = MsgText(601) Then
         MsgBox MsgText(10) & Label12, , MsgText(5)
         strControlButton = MsgText(602)
         Text15.SetFocus
         Exit Sub
      Else
         If ExistCheck("acc1j0", "a1j01 || a1j02", Text7 & Text16, Label13) = False Then
            strControlButton = MsgText(602)
            Text16.SetFocus
            Exit Sub
         End If
         adocheck.CursorLocation = adUseClient
         adocheck.Open "select a1j07, a1j08 from acc1j0 where a1j01 = '" & Text7 & "' and a1j02 = '" & Text16 & "' and ((a1j07 >= " & Val(Text18) & " and a1j08 <= " & Val(Text18) & ") or (a1j07 = 0 and a1j08 = 0))", adoTaie, adOpenStatic, adLockReadOnly
         If adocheck.RecordCount = 0 Then
            MsgBox MsgText(59), , MsgText(5)
            strControlButton = MsgText(602)
            Text18.SetFocus
            adocheck.Close
            Exit Sub
         End If
         adocheck.Close
      End If
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select a0z02 from acc0z0 where a0z02 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         MsgBox MsgText(126), , MsgText(5)
         strControlButton = MsgText(602)
         Text18.SetFocus
         adoquery.Close
         Exit Sub
      End If
      adoquery.Close
      
      If adoacc1l0.RecordCount <> 0 Then
         adoacc1l0.Find "a1l02 = '" & Text15 & "'", 0, adSearchForward, 1
         If adoacc1l0.EOF Then
            adoacc1l0.AddNew
         End If
      Else
         adoacc1l0.AddNew
      End If
      adoacc1l0.Fields("a1l01").Value = Text1
      adoacc1l0.Fields("a1l02").Value = Text15
      adoacc1l0.Fields("a1l03").Value = Text7
      If Text16 <> MsgText(601) Then
         adoacc1l0.Fields("a1l04").Value = Text16
      Else
         adoacc1l0.Fields("a1l04").Value = Null
      End If
      If Combo1 <> MsgText(601) Then
         adoacc1l0.Fields("a1l06").Value = Combo1
      Else
         adoacc1l0.Fields("a1l06").Value = Null
      End If
      
      '存入請款金額(台幣)
      If Text18 <> MsgText(601) Then
         'Modify by Morgan 2004/6/28
         'Y48673000,Y49575000 存小數兩位(後面無條件捨去)
         If m_strCP10 = 605 And (Text8 = "Y48673000" Or Text8 = "Y49575000") Then
            adoacc1l0.Fields("a1l05").Value = Format(Val(Text18), FAmount)
         Else
            adoacc1l0.Fields("a1l05").Value = Val(Text18)
         End If
      Else
         adoacc1l0.Fields("a1l05").Value = 0
      End If
      '折扣(%)
      If Val(Text20) = 100 Then Text20 = 0 'Add By Sindy 2014/2/12 折扣100和無折扣應該是一樣的結果 ex:2002.5
      If Val(Text20) <> 0 Then
         'Modify by Morgan 2004/11/8 改無條件捨去
         'adoacc1l0.Fields("a1l07").Value = Val(Text18) * (100 - Val(Text20)) / 100
         'Modified by Morgan 2014/2/10
         'adoacc1l0.Fields("a1l07").Value = Int(Val(Text18) * (100 - Val(Text20)) / 100)
         adoacc1l0.Fields("a1l07").Value = Val(Text18) - Trunc(Val(Text18) * Val(Text20) / 100)
      Else
         adoacc1l0.Fields("a1l07").Value = 0
      End If
      
      'Modify By Sindy 2025/3/21 非FMP 或 FMP本所服務費 均使用請款匯率計算台幣
      If bolIsFMP = False Or _
         (bolIsFMP = True And Right(Trim(Text16), 2) <> "99" And Right(Trim(Text16), 2) <> "98") Then
         If Combo5.Text <> "NTD" Then
            '以請款匯率換算NTD
            'Modified by Morgan 2025/7/21 要抓畫面上的匯率,否則特殊匯率時會有問題
            'dblInputRate = PUB_GetUSXRate_1(Replace(Me.MaskEdBox1.Text, "/", ""), Me.Combo5.Text)
            dblInputRate = Val(Text19)
            'end 2025/7/21
            If Text18 <> MsgText(601) Then
               'Modify By Sindy 2025/4/9 Anny說外幣換算台幣都要加1元(無條件進位),再換算回來時才不會有落差
               'adoacc1l0.Fields("a1l05").Value = Trunc(Val(Text18) * dblInputRate)
               'Modified by Morgan 2025/4/28
               'adoacc1l0.Fields("a1l05").Value = Round(Val(Text18) * dblInputRate)
               adoacc1l0.Fields("a1l05").Value = -1 * Int(-1 * Val(Text18) * dblInputRate)
            Else
               adoacc1l0.Fields("a1l05").Value = 0
            End If
            If Val(Text20) <> 0 Then
               'Modified by Morgan 2018/4/20 應與非FMP案同一規則--David確認
               'adoacc1l0.Fields("a1l07").Value = Trunc(Val(adoacc1l0.Fields("a1l05").Value) * (100 - Val(Text20)) / 100)
               'Modified by Morgan 2025/7/28 折扣改四捨五入,否則轉外幣可能會差1 Ex:X11409565
               'adoacc1l0.Fields("a1l07").Value = Val(adoacc1l0.Fields("a1l05").Value) - Trunc(Val(adoacc1l0.Fields("a1l05").Value) * Val(Text20) / 100)
               adoacc1l0.Fields("a1l07").Value = Val(adoacc1l0.Fields("a1l05").Value) - Round(Val(adoacc1l0.Fields("a1l05").Value) * Val(Text20) / 100)
            Else
               adoacc1l0.Fields("a1l07").Value = 0
            End If
         End If
      End If
      '2025/3/21 END
      'Add By Sindy 2013/1/24
      If bolIsFMP = True Then
         '大陸官方規費 或 大陸代理人服務費
         If Right(Trim(Text16), 2) = "99" Or Right(Trim(Text16), 2) = "98" Then
            '以輸入之RMB＊當時之RMB報價匯率(只抓小數三位)換算成 NTD
            If Combo5.Text = "RMB" Then
               dblInputRate = PUB_GetAcc210(1, Combo5, Replace(Me.MaskEdBox1.Text, "/", ""))
               If Text18 <> MsgText(601) Then
                  adoacc1l0.Fields("a1l05").Value = Trunc(Val(Text18) * dblInputRate)
               Else
                  adoacc1l0.Fields("a1l05").Value = 0
               End If
            '以輸入之USD＊當時之USD預估結匯匯率(只抓小數三位)換算成 NTD
            ElseIf Combo5.Text = "USD" Then
               dblInputRate = PUB_GetAcc210(2, Combo5, Replace(Me.MaskEdBox1.Text, "/", ""))
               If Text18 <> MsgText(601) Then
                  adoacc1l0.Fields("a1l05").Value = Trunc(Val(Text18) * dblInputRate)
               Else
                  adoacc1l0.Fields("a1l05").Value = 0
               End If
            End If
            If Right(Trim(Text16), 2) = "98" Then
               '檢查其本所服務費是否有輸入折扣
               dblDisc = 0
               'Modified by Morgan 2016/10/17 +a1l07>0
               strExc(0) = "select * from acc1L0 where a1L01='" & Text1 & "' and a1L04='" & Left(Trim(Text16), Len(Trim(Text16)) - 2) & "' and a1l07>0"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  'Added by Morgan 2016/10/17 有記錄折扣數則直接抓該值
                  If RsTemp.Fields("a1l19") > 0 Then
                     dblDisc = 100 * RsTemp.Fields("a1l19")
                  Else
                  'end 2016/10/17
                  
                     If Val(RsTemp.Fields("a1l05").Value) = 0 Then
                        dblDisc = Round(100 - (Val(RsTemp.Fields("a1l07").Value) / 1 * 100))
                     Else
                        dblDisc = Round(100 - (Val(RsTemp.Fields("a1l07").Value) / Val(RsTemp.Fields("a1l05").Value) * 100))
                     End If
                     
                  End If 'Added by Morgan 2016/10/17
               End If
               If Val(dblDisc) > 0 Then
                  'Modified by Morgan 2018/4/20 應與非FMP案同一規則--David確認
                  'adoacc1l0.Fields("a1l07").Value = Trunc(Val(adoacc1l0.Fields("a1l05").Value) * (100 - Val(dblDisc)) / 100)
                  adoacc1l0.Fields("a1l07").Value = Val(adoacc1l0.Fields("a1l05").Value) - Trunc(Val(adoacc1l0.Fields("a1l05").Value) * Val(dblDisc) / 100)
               End If
            End If
         '本所服務費
         Else
            'Modify By Sindy 2025/3/21 Mark,程式往上移
'            If Combo5.Text <> "NTD" Then
'               '以請款匯率換算NTD
'               dblInputRate = PUB_GetUSXRate_1(Replace(Me.MaskEdBox1.Text, "/", ""), Me.Combo5.Text)
'               If Text18 <> MsgText(601) Then
'                  adoacc1l0.Fields("a1l05").Value = Trunc(Val(Text18) * dblInputRate)
'               Else
'                  adoacc1l0.Fields("a1l05").Value = 0
'               End If
'               If Val(Text20) <> 0 Then
'                  'Modified by Morgan 2018/4/20 應與非FMP案同一規則--David確認
'                  'adoacc1l0.Fields("a1l07").Value = Trunc(Val(adoacc1l0.Fields("a1l05").Value) * (100 - Val(Text20)) / 100)
'                  adoacc1l0.Fields("a1l07").Value = Val(adoacc1l0.Fields("a1l05").Value) - Trunc(Val(adoacc1l0.Fields("a1l05").Value) * Val(Text20) / 100)
'               Else
'                  adoacc1l0.Fields("a1l07").Value = 0
'               End If
'            End If
            '檢查是否有98資料,因折扣會互相影響
            strExc(0) = "select * from acc1L0 where a1L01='" & Text1 & "' and a1L04='" & Trim(Text16) & "98'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               '折扣不同時,要重新計算98折扣金額
               Text20 = IIf(Val(Text20) = 0, 100, Val(Text20))
               'Modify by Sindy 2017/3/9 不然會出現溢位
               If Val(RsTemp.Fields("a1l05").Value) > 0 Then
               '2017/3/9 END
                  If Round(100 - (Val(RsTemp.Fields("a1l07").Value) / Val(RsTemp.Fields("a1l05").Value) * 100)) <> Val(Text20) Then
                     dblDiscAmt = Int(Val(RsTemp.Fields("a1l05").Value) * (100 - Val(Text20)) / 100)
                     strSql = "update acc1L0 set a1L07=" & dblDiscAmt & " where a1L01='" & Text1 & "' and a1L04='" & Trim(Text16) & "98'"
                     cnnConnection.Execute strSql
                  End If
               End If
            End If
         End If
      End If
      '2013/1/24 End
      
      'Add by Morgan 2004/8/26 人員時間也要紀錄
      adoacc1l0.Fields("a1l10").Value = strUserNum
      adoacc1l0.Fields("a1l08").Value = strSrvDate(2)
      adoacc1l0.Fields("a1l09").Value = ServerTime
      'Add end
      adoacc1l0.Fields("a1l14").Value = Val(Text9) 'Add by Morgan 2011/3/2
      'Add by Morgan 2011/3/23
      If Text10 <> "" Then
         adoacc1l0.Fields("a1l15").Value = Text10
      Else
         adoacc1l0.Fields("a1l15").Value = Null
      End If
      
      'Add By Sindy 2012/12/27
      'Modify By Sindy 2025/3/21 開放輸入幣別及請款金額各系統都可以使用,不限制只用於FMP
      'Modify By Sindy 2025/4/1 若幣別為NTD就維持不要儲存到a1l16,a1l17,a1l18
      '                         因其他程式的判斷會受影響
      '                         + And Combo5.Text = "NTD"
      If bolIsFMP = False And Combo5.Text = "NTD" Then
      '2025/4/1 END
         adoacc1l0.Fields("a1l16").Value = Null
         adoacc1l0.Fields("a1l17").Value = Null
         adoacc1l0.Fields("a1l18").Value = Null
      Else
         adoacc1l0.Fields("a1l16").Value = Combo5.Text
         If Text18 <> MsgText(601) Then
            adoacc1l0.Fields("a1l17").Value = Val(Text18)
         Else
            adoacc1l0.Fields("a1l17").Value = 0
         End If
         'Modify By Sindy 2025/3/21
         If bolIsFMP = False Then
'            adoacc1l0.Fields("a1l16").Value = Null
'            adoacc1l0.Fields("a1l17").Value = Null
            adoacc1l0.Fields("a1l18").Value = Null
         Else
         '2025/3/21 END
            If Text13 <> MsgText(601) Then
               adoacc1l0.Fields("a1l18").Value = Val(Text13)
            Else
               adoacc1l0.Fields("a1l18").Value = 0
            End If
         End If
      End If
      '2012/12/27 End
      
      adoacc1l0.UpdateBatch
      adoacc1l0.ReQuery 'Added by Morgan 2017/11/21 此處須重新查詢，因新增時若有更新到代收代付(98)的資料會導致錯誤-2147217864"找不到要更新的資料列。最後讀取的值已被變更。"
      
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   
End Sub

'*************************************************
'  顯示 Adodc 之資料
'
'*************************************************
Private Sub AdodcShow()
   Text15 = Adodc1.Recordset.Fields("a1l02").Value
   If IsNull(Adodc1.Recordset.Fields("a1l04").Value) Then
      Text16 = MsgText(601)
   Else
      Text16 = Adodc1.Recordset.Fields("a1l04").Value
   End If
   
   'Add By Sindy 2013/1/24
   If Right(Trim(Text16), 2) = "99" Or Right(Trim(Text16), 2) = "98" Then
      Text20.Enabled = False
      Text20.Text = ""
   Else
      Text20.Enabled = True
   End If
   
   Text17 = A1j03Query(Text7, Text16)
   If IsNull(Adodc1.Recordset.Fields("a1l06").Value) Then
      Combo1 = MsgText(601)
   Else
      Combo1 = Adodc1.Recordset.Fields("a1l06").Value
   End If
   
   Text9 = "" & Adodc1.Recordset.Fields("a1l14").Value 'Add by Morgan 2011/3/2 要在請款金額前面否則該金額會被覆蓋
   Text10 = "" & Adodc1.Recordset.Fields("a1l15").Value 'Add by Morgan 2011/3/23 要在請款金額前面否則該金額會被覆蓋
   
   'Modified by Morgan 2016/8/1
   'If bolIsFMP = True Then Call SetCombo5
   If bolIsFMP = True Then Call SetCombo5("" & Adodc1.Recordset.Fields("a1l16").Value)
   'end 2016/8/1
   '2013/1/24 End
   'Add by Sindy 2012/12/27
   If "" & Adodc1.Recordset.Fields("a1l16").Value > "" Then
      Combo5.Text = "" & Adodc1.Recordset.Fields("a1l16").Value
   'Add By Sindy 2025/3/21
   Else
      If bolIsFMP = False Then
         Combo5.ListIndex = int_NTD '預設台幣
      End If
   '2025/3/21 END
   End If
   '請款金額
   If Val("" & Adodc1.Recordset.Fields("a1l17").Value) > 0 Then
      Text18 = Val(Adodc1.Recordset.Fields("a1l17").Value)
   Else
   '2012/12/27 End
      If IsNull(Adodc1.Recordset.Fields("a1l05").Value) Then
         Text18 = MsgText(601)
      Else
         Text18 = Val(Adodc1.Recordset.Fields("a1l05").Value)
      End If
   End If
   'Add By Sindy 2013/4/17
   If IsNull(Adodc1.Recordset.Fields("a1l18").Value) Then
      Text13 = MsgText(601)
   Else
      If Adodc1.Recordset.Fields("a1l18").Value = 0 Then
         Text13 = MsgText(601)
      Else
         Text13 = Val(Adodc1.Recordset.Fields("a1l18").Value)
      End If
   End If
   '2013/4/17 End
   If IsNull(Adodc1.Recordset.Fields("a1l07").Value) Then
      Text20 = "100"
   'Added by Morgan 2015/8/5
   ElseIf Not IsNull(Adodc1.Recordset.Fields("a1l19").Value) Then
      If Adodc1.Recordset.Fields("a1l19").Value = 1 Then
         Text20 = 0
      Else
         Text20 = Adodc1.Recordset.Fields("a1l19").Value * 100
      End If
   'end 2015/8/5
   Else
      If Text18 <> MsgText(601) Then
         If Adodc1.Recordset.Fields("a1l05").Value <> 0 Then
            'Modify by Morgan 2011/3/2 折扣數顯示整數
            'Text20 = 100 - (Val(Adodc1.Recordset.Fields("a1l07").Value) / Val(Adodc1.Recordset.Fields("a1l05").Value) * 100)
            Text20 = Round(100 - (Val(Adodc1.Recordset.Fields("a1l07").Value) / Val(Adodc1.Recordset.Fields("a1l05").Value) * 100))
         Else
            Text20 = "100"
         End If
      Else
         Text20 = "100"
      End If
   End If
   
   'Added by Morgan 2018/7/19 請款項目跳離有相關控制如預設折扣...
   'Modify by Amy 2025/11/12 +stF0301,由結案單傳更新資料進入會錯
   If Text16.Enabled = True And stF0301 = "" Then
      Text16.SetFocus
   End If
   Text16.Tag = Text16
   'end 2018/7/19
   
   m_strItemNo = Me.Text15.Text 'Added by Morgan 2019/8/8
End Sub

'*************************************************
'  刪除 Adodc 之資料
'
'*************************************************
Private Sub AdodcDelete()
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   adoacc1l0.Find "a1l02 = '" & Text15 & "'", 0, adSearchForward, 1
   If adoacc1l0.EOF = False Then
      adoacc1l0.Delete
      adoacc1l0.UpdateBatch
      AdodcRefresh
      SumShow
      AdodcClear
   End If
   'Add by Morgan 2004/8/26 同步更新主檔
   Frmacc21h1_Save
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Added by Morgan 2015/11/27 檢查項目是否可請款
Private Function chkItem(pItemNo As String) As Boolean
   
   If Text7 = "FCP" Or Text7 = "P" Then
      'Added by Morgan 2025/2/10
      '檢查專利不請雜費設定
      If pItemNo = "02" Then
         If ChkItemFCP02(True) = False Then
            Exit Function
         End If
      End If
      'end 2025/2/10
   
      'Added by Morgan 2015/7/13 'X56842谷歌公司不可請02雜費
      'Modified by Morgan 2015/8/3 +X49346,X72101
      'Modified by Morgan 2017/3/10 +X70722010,X70286,X74494
      'Removed by Morgan 2025/2/11 已增加專利不得請雜費欄位設定
      'If pItemNo = "02" And (m_AppNo = "X56842000" Or m_AppNo = "X49346000" Or m_AppNo = "X72101000" Or m_AppNo = "X70722010" Or m_AppNo = "X70286000" Or m_AppNo = "X74494000") Then
      '   MsgBox "案件申請人【" & m_AppNo & " " & CustomerQuery(m_AppNo, 2) & "】不得請【" & pItemNo & A1j03Query(Text7, pItemNo) & "】！", vbExclamation
      '   Exit Function
      'End If
      'end 2025/2/11
    
      'Added by Morgan 2015/10/30 Y54037000 Mondelez Global LLC 不可請雜費02及影印費08
      'Modified by Morgan 2025/2/11 已增加專利不得請雜費欄位設定
      'If Text8 = "Y54037000" And (pItemNo = "02" Or pItemNo = "08") Then
      If Text8 = "Y54037000" And (pItemNo = "08") Then
      'end 2025/2/11
         MsgBox "請款對象【" & Text8 & " " & FagentQuery(Text8, 2) & "】不得請【" & pItemNo & A1j03Query(Text7, pItemNo) & "】！", vbExclamation
         Exit Function
      End If
      
      'Added by Morgan 2015/11/27
      'X62773卡夫特食品研究發展公司 (MONDELEZ 相關企業) 的案件下案件性質皆不得請款
      '1.202補文件 2.01傳真 3.02雜費 4.03打字費 5.08影印費
      'Modified by Morgan 2025/2/11 已增加專利不得請雜費欄位設定
      'If m_AppNo = "X62773000" And (pItemNo = "202" Or pItemNo = "01" Or pItemNo = "02" Or pItemNo = "03" Or pItemNo = "08") Then
      If m_AppNo = "X62773000" And (pItemNo = "202" Or pItemNo = "01" Or pItemNo = "03" Or pItemNo = "08") Then
      'end 2025/2/11
         MsgBox "案件申請人No.為【" & ChangeCustomerS(m_AppNo) & "】" & pItemNo & A1j03Query(Text7, pItemNo) & "不得請款！", vbExclamation
         Exit Function
      End If
      
      'Added by Morgan 2015/11/3 Y54179 Longitude 案件不可請雜費02
      'Modified by Morgan 2017/11/6 +Y53495,Y53495010 --洪培堯
      'Removed by Morgan 2025/2/11 已增加專利不得請雜費欄位設定
      'If (Text2 = "Y54179000" Or Text2 = "Y53495000" Or Text2 = "Y53495010") And (pItemNo = "02") Then
      '   MsgBox "案件代理人【" & Text2 & " " & FagentQuery(Text2, 2) & "】不得請【" & pItemNo & A1j03Query(Text7, pItemNo) & "】！", vbExclamation
      '   Exit Function
      'End If
      'end 2025/2/11
      
      'Added by Morgan 2017/11/3 Y54869 Albemarle Corporation 不可請雜費02
      'Modified by Morgan 2019/8/5 +Y55134 Albemarle Germany GmbH --Ryan
      'Modified by Morgan 2020/3/18 +Y53942 Xperi Corporation --Ali
      'Removed by Morgan 2025/2/11 已增加專利不得請雜費欄位設定
      'If (Text8 = "Y54869000" Or Text8 = "Y55134000" Or Text8 = "Y53942000") And pItemNo = "02" Then
      '   MsgBox "請款對象【" & Text8 & " " & FagentQuery(Text8, 2) & "】不得請【" & pItemNo & A1j03Query(Text7, pItemNo) & "】！", vbExclamation
      '   Exit Function
      'End If
      'end 2025/2/11
      
      'Added by Lydia 2018/12/22
      ' Y52878B10 Allergan Plc : 如遇到下述請款項目 , 請設定請款彈跳提醒: 以下請款項目不可以請款:  --羅暐曄
      ' 01傳真 , 02雜費 , 03打字費
      'Modified by Morgan 2025/2/11 已增加專利不得請雜費欄位設定
      'If Text2 = "Y52878B10" And (pItemNo = "01" Or pItemNo = "02" Or pItemNo = "03") Then
      If Text2 = "Y52878B10" And (pItemNo = "01" Or pItemNo = "03") Then
      'end 2025/2/11
         MsgBox "案件代理人【" & Text2 & " " & FagentQuery(Text2, 2) & "】不得請【" & pItemNo & A1j03Query(Text7, pItemNo) & "】！", vbExclamation
         Exit Function
      End If
      'end 2018/12/22
      
   End If
   chkItem = True
End Function
'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
Dim Cancel As Boolean
   
   Call PUB_SaveTrackMode(1, KeyCode)  'Added by Lydia 2021/12/08 Form2.0 記錄鍵盤傳入順序
   
   Select Case KeyCode
      Case vbKeyInsert
         'Added by Lydia 2021/12/08 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
         If PUB_ChkTrackMode = False Then
            Exit Sub
         End If
         
         'Added by Lydia 2021/12/08 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
          If PUB_ChkUniText(Me, , True, "TextBox") = False Then
              Exit Sub
          End If
         'end 2021/12/08
    
'Modified by Morgan 2015/11/27 改呼叫共用函數檢查
'         'Added by Morgan 2015/7/13 不可請02雜費控制
'         'Modified by Morgan 2015/11/3 Y54179案件不可請雜費02
'         If Text7 = "FCP" And Text16 = "02" And (m_bolNoDisbursements Or Text2 = "Y54179000") Then
'            MsgBox "本案不可請02雜費！", vbExclamation
'            If Text16.Enabled = True Then Text16.SetFocus
'            Exit Sub
'         End If
'         'end 2015/7/13
         
'         'Added by Morgan 2015/10/30 Y54037000 不可請雜費02及影印費08
'          If Text8 = "Y54037000" And Text7 = "FCP" And (Text16 = "02" Or Text16 = "08") Then
'               MsgBox "本案不可請02雜費或08影印費！", vbExclamation
'               If Text16.Enabled = True Then Text16.SetFocus
'               Exit Sub
'          End If
'         'end 2015/10/30
         If chkItem(Text16) = False Then
            If Text16.Enabled = True Then Text16.SetFocus
            Exit Sub
         End If
'end 2015/11/27

         'Add by Morgan 2007/5/24
         If Val(Text20) > 100 Then
            MsgBox "折扣不可大於100%！", vbExclamation
            If Text20.Enabled = True Then Text20.SetFocus
            Exit Sub
         End If
         'end 2007/5/24
         
         'Add by Morgan 2011/3/15 商品數金額檢查
         'Modified by Morgan 2015/11/19 證明標章(系統商標種類:7)及團體標章(系統商標種類:8)除外--陳金蓮
         If Text7 = "FCT" And (Text16 = "10199" Or Text16 = "A0199") And m_strTM08 <> "7" And m_strTM08 <> "8" Then
            If Val(Text18) <> GetFCT10199Fee Then
               If MsgBox("請款金額與計算規則不符，是否確定要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
                  If Text9.Enabled = True Then Text9.SetFocus: Text9_GotFocus
                  Exit Sub
               End If
            End If
         End If
         '規費內容
         If CheckLengthIsOK(Combo1, 1) = False Then
            MsgBox "規費內容長度超過！"
            Exit Sub
         End If
         'END 2011/3/15
         
         'Add By Sindy 2012/12/27
         'If bolIsFMP = True Then 'Modify By Sindy 2025/3/20 均開放輸入幣別
            If Trim(Combo5.Text) = "" Then
               MsgBox "輸入幣別不可空白！"
               Exit Sub
            End If
            Cancel = False
            Call Combo5_Validate(Cancel)
            If Cancel = True Then
               Exit Sub
            End If
         'End If
         '2012/12/27 End
                  
         'Added by Morgan 2015/9/23
         '檢查翻譯費用是否超過比例
         'Removed by Morgan 2019/5/29 因計算比例改要含打字費,移到 Form_QueryUnload 檢查
         'If Text16 = "201" And (Text7 = "FCP" Or Text7 = "FG" Or Text7 = "P" Or Text7 = "CFP") Then
         '   'Mofieid by Morgan 2015/10/29
         '   'Mofieid by Morgan 2016/5/3 輸入幣別要判斷有選才要乘匯率
         '   If PUB_ChkTranslationFee(Text1, Val(Text18) * IIf(Val(Text20) > 0, Val(Text20) / 100, 1) * IIf(Combo5.Visible = True And Combo5.Text <> "" And Combo5.Text <> "NTD", Val(Text19), 1), False) = False Then
         '      Exit Sub
         '   End If
         'End If
         'end 2019/5/29
         'end 2015/9/23
         
         'Add By Cheng 2004/05/13
         '記錄請款項目序號
         m_strItemNo = Me.Text15.Text
'         If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
'            Exit Sub
'         End If
         If strControlButton <> MsgText(602) Then
            Acc1l0Save
         End If
         If strControlButton <> MsgText(602) Then
            AdodcRefresh
            'Add By Cheng 2004/05/13
            '調整Grid畫面, 讓新增或修改的該筆資料能顯示在畫面上, 不用捲動捲軸
            If Me.Adodc1.Recordset.RecordCount > 0 Then Me.Adodc1.Recordset.MoveLast
            Do While Not Adodc1.Recordset.EOF
                If "" & Adodc1.Recordset.Fields(1).Value = m_strItemNo Then
                  Exit Do
                End If
                Adodc1.Recordset.MovePrevious
            Loop
            
            'End
            'Removed by Morgan2017/11/9 Command2_Click 內有執行不必重複
            'SumShow
            'AdodcClear
            'end 2017/11/9
            
            'Added by Morgan 2019/8/8 外專程序輸入時改自動帶出下一項目--淑華,敏莉
            If Pub_StrUserSt03 = "F22" Then
               Adodc1.Recordset.MoveNext
               If Adodc1.Recordset.EOF Then
                  Adodc1.Recordset.MoveLast
                  Command2_Click
               Else
                  SumShow
                  DataGrid1_SelChange 0
                  'Added by Morgan 2021/2/23
                  Cancel = False
                  Call Text16_Validate(Cancel)
                  If Cancel = True Then
                     Text16.SetFocus
                  Else
                  'end 2021/2/23
                     Text18.SetFocus 'Added by Morgan 2019/8/19 --敏莉
                  End If
               End If
            Else
            'end 2019/8/8
            
               Command2_Click
               
            End If 'Added by Morgan 2019/8/8
                        
            'Add by Morgan 2004/8/26 同步更新主檔
            Frmacc21h1_Save
            'add by nick 2004/11/30  印地址條
            If pub_blnARPrintAddress = True Then
                pub_AddressListSN = pub_AddressListSN + 1
                'edit by nick 2004/11/10
                'PUB_AddNewAddressList strUserNum, "" & Me.Text7.Text, "" & Me.Text21.Text, "" & Me.Text22.Text, "" & Me.Text23.Text, "" & pub_AddressListSN, "0", m_strCP10
                PUB_AddNewAddressList strUserNum, "" & Me.Text7.Text, "" & Me.Text21.Text, "" & Me.Text22.Text, "" & Me.Text23.Text, "" & pub_AddressListSN, "0", IIf(UCase(Me.Text7.Text) = "FCT", IIf(m_strCP10 = "102", m_strCP10, ""), m_strCP10)
            End If
         End If
         strControlButton = MsgText(601)
   End Select
   KeyEnter KeyCode
End Sub

'*************************************************
'  清除查詢顯示
'
'*************************************************
Public Sub AdodcClear()
   Text15 = ""
   Text16 = ""
   Text17 = ""
   Combo1 = ""
   Text18 = ""
   Text20 = ""
   Text9 = "" 'Add by Morgan 2011/3/2
   Text10 = "" 'Add by Morgan 2011/3/23
   'Combo5 = "" 'Add by Sindy 2013/4/17
   Text13 = "" 'Add by Sindy 2013/4/17
   Text16.Tag = Text16 'Added by Morgan 2018/7/19
   'Add By Sindy 2025/3/25
   If Combo5.ListIndex = -1 Or Combo5.Text = "" Then
      If bolIsFMP = False Then
         Combo5.ListIndex = int_NTD
      Else
         Combo5.ListIndex = 1
      End If
   End If
   '2025/3/25 END
End Sub

'Add by Morgan 2010/7/30
Private Sub Text5_Change()
   m_blnClkPrintButton = False
   'Add by Morgan 2010/12/1
   'Modified by Lydia 2015/04/15 為了區別整批請款單,a1k32=C
   'Modified by Morgan 2015/11/12 還原,此處為特殊請款單控制而不是指整批列印
   If Text5 = "Y" Then
   'If Text5 = "C" Then
   'end 2015/11/12
      Command3.Caption = "Word(&W)"
      'Added by Morgan 2014/8/18 整批請款只能單筆修改
      If Me.m_bolIsBatch Then
         Command3.Enabled = False
      Else
         Command3.Enabled = True
      End If
      'end 2014/8/18
   Else
      Command3.Caption = "列印(&P)"
      Command3.Enabled = True 'Added by Morgan 2014/8/18
   End If
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Modified by Lydia 2015/04/15 為了區別整批請款單,+C
   'Modified by Morgan 2015/11/12 取消C(還原),該標記只能由整批列印回寫
   If KeyAscii <> 8 And Chr(KeyAscii) <> "Y" Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   If Text6 = MsgText(601) Then
      Exit Sub
   End If
   'Add By Cheng 2003/06/25
   '若有輸入列印對象則自動補滿9碼
   'Modified by Morgan 2023/12/5 第9碼若是空白或錯誤會檢查不出來
   'Me.Text6.Text = Left(Me.Text6.Text & "000000000", 9)
   'If Len(Text6) = 6 Then
   '   Text6 = AfterZero(Text6)
   'Else
   '   If Len(Text6) = 8 Then
   '      Text6 = Text6 & "0"
   '   End If
   'End If
   'If Mid(Text6, 1, 1) = "X" Then
   '   If ExistCheck("customer", "cu01", Mid(Text6, 1, 8), Label15) = False Then
   '      MsgBox MsgText(28), , MsgText(5)
   '      Cancel = True
   '      Exit Sub
   '   End If
   'Else
   '   If ExistCheck("fagent", "fa01", Mid(Text6, 1, 8), Label15) = False Then
   '      MsgBox MsgText(28), , MsgText(5)
   '      Cancel = True
   '      Exit Sub
   '   End If
   'End If
   Text6.Text = Left(Trim(Text6.Text) & "000000000", 9)
   If Left(Text6.Text, 1) = "X" Then
      If ClsPDGetCustomer(Text6.Text, strExc(1)) = False Then
         MsgBox MsgText(28), , Label15
         Cancel = True
         Exit Sub
      End If
   Else
      If ClsPDGetAgent(Text6.Text, strExc(1)) = False Then
         MsgBox MsgText(28), , Label15
         Cancel = True
         Exit Sub
      End If
   End If
   'end 2023/12/5
   
   'Added by Morgan 2012/12/6
   '列印對象或請款幣別變更要重新預設列印幣別格式
   If Text6.Tag <> Text6 Then
      'Modify By Sindy 2016/12/16 + , , Text21, Text22, Text23
      'If m_bolAfterLoad And bolIsFMP = False Then Combo4.ListIndex = PUB_GetDefaultCurrPrintType(Text7, Text6, Combo3) - 1
      'Modified by Morgan 2018/4/27
      'If m_bolAfterLoad And bolIsFMP = False Then Combo4.ListIndex = PUB_GetDefaultCurrPrintType(Text7, Text6, Combo3, , Text21, Text22, Text23) - 1
      If m_bolAfterLoad And bolIsFMP = False Then Combo4.ListIndex = PUB_GetDefaultCurrPrintType(Text7, Text8, Combo3, , Text21, Text22, Text23, Text6) - 1
      'end 2018/4/27
   End If
   Text6.Tag = Text6
   'end 2012/12/6
   
   If Text6 = "Y54443000" And Text5 = "" Then Text5 = "Y" 'Added by Morgan 2017/7/7
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  合計顯示
'
'*************************************************
Public Sub SumShow()
Dim douAmount As Double
Dim douDiscount As Double
Dim USRate As Double   '2009/4/23 ADD BY SONIA
Dim dblAmt As Double 'Add By Sindy 2013/3/28
'Added by Morgan 2013/11/1
Dim bolNewForm As Boolean '用新的請款單格式
Dim dblItemAmt As Double '單項金額
Dim dblOfficialAmt As Double '服務費小計
Dim dblServiceAmt As Double '規費小計
'end 2013/11/1
   
   'Added by Morgan 2013/12/4 請款日期>=1021205改用新格式
   bolNewForm = False
   If Val(Replace(Me.MaskEdBox1.Text, "/", "")) >= 1021205 Then
      'Modified by Morgan 2014/2/25 +傳本所案號
      If PUB_GetBillFormat(Text8, Text7, Text21, Text22, Text23) = 0 Then
         bolNewForm = True
      End If
   End If
   'end 2013/12/4
   
   Label19 = Me.Combo3.Text & "$"   '2009/4/23 ADD BY SONIA
   
   'Modify By Sindy 2013/1/24
'   If bolIsFMP = True Then
      '以請款幣別預估結匯匯率換算外幣金額
      'dblInputRate = PUB_GetAcc210(2, Combo3, Replace(Me.MaskEdBox1.Text, "/", ""))
      adoaccsum.CursorLocation = adUseClient
      '明細外幣金額捨去小數後加總
      'adoaccsum.Open "select sum(a1l05), sum(a1l07), sum(trunc((a1l05-nvl(a1l07,0))/" & IIf(dblInputRate = 0, 1, dblInputRate) & ",0)) from acc1l0 where a1l01 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
      adoaccsum.Open "select * from acc1l0 where a1l01 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoaccsum.RecordCount <> 0 Then
         adoaccsum.MoveFirst
         Text12 = MsgText(601): dblAmt = 0
         Text14 = MsgText(601)
         Do While Not adoaccsum.EOF
            If IsNull(adoaccsum.Fields("a1l05").Value) Then
               douAmount = 0
            Else
               douAmount = Val(adoaccsum.Fields("a1l05").Value)
            End If
            If IsNull(adoaccsum.Fields("a1l07").Value) Then
               douDiscount = 0
            Else
               douDiscount = Val(adoaccsum.Fields("a1l07").Value)
            End If
            Text14 = Val(Text14) + (douAmount - douDiscount) '折扣後 NT$
            
            'Modified by Morgan 2013/11/1
            'dblAmt = dblAmt + GetDebitNoteFAmt(Text1, Combo3, Replace(Me.MaskEdBox1.Text, "/", ""), "" & adoaccsum.Fields("a1l04"), "" & adoaccsum.Fields("a1l05"), _
                     "" & adoaccsum.Fields("a1l07"), "" & adoaccsum.Fields("a1l16"), "" & adoaccsum.Fields("a1l17"), Combo4.ListIndex + 1, Text19)
             'Added by Morgan 2019/8/30 BASF 翻譯費 美金要顯示至小數第2位
             'Modified by Morgan 2022/2/18 +927其他翻譯 Ex:X11102382-- Ryan
             'Modified by Morgan 2022/9/1 +209檢視中說-- Tim
             'Modified by Morgan 2025/10/31 +FG Ex:X11404796
             If (Text8 = "Y45814010" Or Text8 = "Y33268010") And (Text7 = "FCP" Or Text7 = "FG" Or Text7 = "P" Or Text7 = "CFP") And (adoaccsum.Fields("a1l04") = "201" Or adoaccsum.Fields("a1l04") = "927" Or adoaccsum.Fields("a1l04") = "209") Then
               'Modified by Morgan 2021/12/8 補減折扣金額 Ex:X11017688
               'Modify By Sindy 2025/4/2 開放可以輸入幣別,幣別相同又沒有折扣問題,不用再換算直接使用A1L17
               If "" & adoaccsum.Fields("a1l16").Value = Combo3.Text And Val("" & adoaccsum.Fields("a1l07")) = 0 Then
                  dblItemAmt = Val("" & adoaccsum.Fields("a1l17"))
               Else
               '2025/4/2 END
                  dblItemAmt = Format((Val("" & adoaccsum.Fields("a1l05")) - Val("" & adoaccsum.Fields("a1l07"))) / Val(Text19), "#.00")
               End If
             Else
             'end 2019/8/30
             
             dblItemAmt = GetDebitNoteFAmt(Text1, Combo3, Replace(Me.MaskEdBox1.Text, "/", ""), "" & adoaccsum.Fields("a1l04"), "" & adoaccsum.Fields("a1l05"), _
                       "" & adoaccsum.Fields("a1l07"), "" & adoaccsum.Fields("a1l16"), "" & adoaccsum.Fields("a1l17"), Combo4.ListIndex + 1, Text19)
                       
             End If 'Added by Morgan 2019/8/30
             dblAmt = dblAmt + dblItemAmt
             If Right(adoaccsum("a1l04"), 2) = "99" Then '新格式啟用後規費只會是99結尾
               dblOfficialAmt = dblOfficialAmt + dblItemAmt
             Else
               dblServiceAmt = dblServiceAmt + dblItemAmt
             End If
             'end 2013/11/1
                
'            '以請款幣別輸入者以A1L17計算
'            If Trim("" & adoaccsum.Fields("a1l16")) = Combo3 Then
'               dblAmt = dblAmt + Val(adoaccsum.Fields("a1l17"))
'            '以RMB輸入者以A1L05/請款幣別預估結匯匯率計算
'            ElseIf Trim("" & adoaccsum.Fields("a1l16")) = "RMB" Then
'               '1.純台幣,2.台幣+外幣合計 : 逐筆加總不去小數,最後才一次去小數
'               If Combo4.ListIndex = 0 Or Combo4.ListIndex = 1 Then
'                  dblAmt = dblAmt + (Val(adoaccsum.Fields("a1l05")) - Val(adoaccsum.Fields("a1l07"))) / dblInputRate
'               '3.純外幣,4.外幣+美金合計 : 逐筆加總並且去小數
'               Else
'                  dblAmt = dblAmt + Trunc(((Val(adoaccsum.Fields("a1l05")) - Val(adoaccsum.Fields("a1l07"))) / dblInputRate), 0)
'               End If
'            '以請款幣別請款匯率計算
'            Else
'               '1.純台幣,2.台幣+外幣合計 : 逐筆加總不去小數,最後才一次去小數
'               If Combo4.ListIndex = 0 Or Combo4.ListIndex = 1 Then
'                  dblAmt = dblAmt + (Val(adoaccsum.Fields("a1l05")) - Val(adoaccsum.Fields("a1l07"))) / dblRate
'               '3.純外幣,4.外幣+美金合計 : 逐筆加總並且去小數
'               Else
'                  dblAmt = dblAmt + Trunc(((Val(adoaccsum.Fields("a1l05")) - Val(adoaccsum.Fields("a1l07"))) / dblRate), 0)
'               End If
'            End If
            
            adoaccsum.MoveNext
         Loop
         If dblAmt > 0 Then
            '1.純台幣,2.台幣+外幣合計
            If Combo4.ListIndex = 0 Or Combo4.ListIndex = 1 Then
               'Modified by Morgan 2013/11/1 新格式服務費規費各自小計故都要捨去
               'Text12 = Format(Fix(Format(dblAmt)), FAmount) '最後才一次去小數
               If bolNewForm = True Then
                  Text12 = Format(Fix(Format(dblOfficialAmt)) + Fix(Format(dblServiceAmt)), FAmount)
               Else
                  Text12 = Format(Fix(Format(dblAmt)), FAmount)
               End If
               'end 2013/11/1
               
            '3.純外幣,4.外幣+美金合計
            Else
               Text12 = dblAmt
            End If
         End If
      Else
         Text12 = MsgText(601)
         Text14 = MsgText(601)
      End If
      adoaccsum.Close
'   Else
'   '2013/1/24 End
'      adoaccsum.CursorLocation = adUseClient
'      'Modified by Morgan 2012/1/31 +明細外幣金額四捨五入後加總
'      'Modified by Morgan 2012/12/6 改明細外幣金額捨去小數後加總
'      'adoaccsum.Open "select sum(a1l05), sum(a1l07), sum(round((a1l05-nvl(a1l07,0))/" & IIf(dblRate = 0, 1, dblRate) & ",2)) from acc1l0 where a1l01 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
'      adoaccsum.Open "select sum(a1l05), sum(a1l07), sum(trunc((a1l05-nvl(a1l07,0))/" & IIf(dblRate = 0, 1, dblRate) & ",0)) from acc1l0 where a1l01 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adoaccsum.RecordCount <> 0 Then
'         If IsNull(adoaccsum.Fields(0).Value) Then
'            douAmount = 0
'         Else
'            douAmount = Val(adoaccsum.Fields(0).Value)
'         End If
'         If IsNull(adoaccsum.Fields(1).Value) Then
'            douDiscount = 0
'         Else
'            douDiscount = Val(adoaccsum.Fields(1).Value)
'         End If
'         If dblRate <> 0 Then
'
'            'Modified by Morgan 2012/12/6
'            '統一改判斷列印幣別
'            '1.純台幣,2.台幣+外幣合計
'            '台幣總額換算外幣後捨去小數
'            If Combo4.ListIndex = 0 Or Combo4.ListIndex = 1 Then
'               Text12 = Format(Fix(Format((douAmount - douDiscount) / dblRate)), FAmount)
'   '            Text13 = Text12 'Modify By Sindy 2012/12/27 Mark 美金 USD$
'
'            '3.純外幣,4.外幣+美金合計
'            '明細外幣金額捨去小數後加總,美金
'            Else
'               Text12 = Format(Val("" & adoaccsum.Fields(2).Value), FAmount)
'               'Modify By Sindy 2012/12/27 Mark
'   '            If Combo3.Text <> "USD" Then
'   '               USRate = PUB_GetDNRate(Replace(Me.MaskEdBox1.Text, "/", ""), Me.Combo3.Text)
'   '               Text13 = Fix(Format(Text12 * USRate))
'   '            Else
'   '               Text13 = Text12
'   '            End If
'               '2012/12/27 End
'            End If
'            'end 2012/12/6
'
'   'Removed by Morgan 2012/12/6
'   '         '2009/4/23 ADD BY SONIA 加非美金請款
'   '         If Me.Combo3.Text <> "USD" Then
'   '            '計算請款幣別合計
'   '            Text12 = Format((((douAmount - douDiscount) * 100 * 100) \ (dblRate * 100)) / 100, FAmount)
'   '            '抓請款幣別對美金匯率
'   '            USRate = PUB_GetDNRate(Replace(Me.MaskEdBox1.Text, "/", ""), Me.Combo3.Text)
'   '            '計算美金合計取至整數位(無條件捨去),注意若不捨去則Frmacc21h1_Save要修改
'   '            'Morgan 2012/11/12 改四捨五入到小數兩位--葉經理
'   '            'Text13 = Format(Text12 * USRate * 100 \ 100, FAmount)
'   '            Text13 = Round(Text12 * USRate, 2)
'   '         Else
'   '         '2009/4/23 END
'   '
'   ''Modified by Morgan 2012/11/2 外幣合計抓明細外幣加總(小數兩位)--David,Frances
'   ''
'   ''           'Modify By Cheng 2004/04/27
'   ''           '美金取至整數位(無條件捨去)
'   ''   '         Text13 = Format((douAmount - douDiscount) / dblRate, FAmount)
'   ''
'   ''            'Modify by Morgan 2004/7/20
'   ''            'Y48673000,Y49575000 存小數兩位(後面無條件捨去)
'   ''            If m_strCP10 = "605" And (Text8 = "Y48673000" Or Text8 = "Y49575000") Then
'   ''               'Modify by Morgan 2005/3/28 fix有bug fix(849/28.3)=29  須先轉dbl
'   ''               'Modify by Morgan 2005/4/29 fix bug再修改
'   ''               'Text13 = Format(Fix(CDbl((douAmount - douDiscount) * 100 / dblRate)) / 100, FAmount)
'   ''               Text13 = Format((((douAmount - douDiscount) * 100 * 100) \ (dblRate * 100)) / 100, FAmount)
'   ''
'   ''            'Added by Morgan 2012/1/31 Y52218 PanKorea Patent & Law Firm 美金加總保留小數點--David
'   ''            'Modified by Morgan 2012/7/6 +Y34126 L'AIR LIQUIDE SA DIRECTION DE LA PROPRIETE INTELLECTUELLE--David
'   ''            'Modified by Morgan 2012/8/31 +Y48292000 HP
'   ''            'Modified by Morgan 2012/10/11 +Y45149000,Y45149010
'   ''            'Modified by Morgan 2012/11/2 +Y23045000 --陳芊穎
'   ''            ElseIf Text8 = "Y52218000" Or Text8 = "Y34126000" Or Text8 = "Y48292000" Or Text8 = "Y45149000" Or Text8 = "Y45149010" Or Text8 = "Y23045000" Then
'   ''               Text13 = Format(Val("" & adoaccsum.Fields(2).Value), FAmount)
'   ''
'   ''            'end 2012/1/31
'   ''            Else
'   ''               'Modify by Morgan 2005/3/28 fix有bug fix(849/28.3)=29  須先轉dbl
'   ''               'Modify by Morgan 2005/4/29 fix bug再修改
'   ''               'Text13 = Format(Fix(CDbl((douAmount - douDiscount) / dblRate)), FAmount)
'   ''               Text13 = Format(((douAmount - douDiscount) * 100) \ (dblRate * 100), FAmount)
'   ''            End If
'   ''
'   ''           'End
'   '
'   '           Text13 = Round(Val("" & adoaccsum.Fields(2).Value), 2)
'   '
'   ''end 2012/11/2
'   '           Text12 = Text13    '2009/4/23 ADD BY SONIA
'   '         End If
'         Else
'            Text12 = 0           '2009/4/23 ADD BY SONIA
'   '         Text13 = 0 'Modify By Sindy 2012/12/27 Mark
'         End If
'         Text14 = douAmount - douDiscount
'      Else
'         Text12 = MsgText(601)   '2009/4/23 ADD BY SONIA
'   '      Text13 = MsgText(601) 'Modify By Sindy 2012/12/27 Mark
'         Text14 = MsgText(601)
'      End If
'      adoaccsum.Close
'   End If
End Sub

'*************************************************
'  關閉分錄欄位輸入狀態
'
'*************************************************
Public Sub FormDisabled()
   Text15.Enabled = False
   Text16.Enabled = False
   Combo1.Enabled = False
   Text18.Enabled = False
   Text20.Enabled = False
   Command1.Enabled = False
   Command2.Enabled = False
   'Add By Sindy 2012/12/27
   Frame1.Enabled = False
   Frame2.Enabled = False
   '2012/12/27 End
End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'
'*************************************************
Public Sub FormEnabled()
   Text15.Enabled = True
   Text16.Enabled = True
   Combo1.Enabled = True
   Text18.Enabled = True
   Text20.Enabled = True
   Command1.Enabled = True
   Command2.Enabled = True
   'Add By Sindy 2012/12/27
   Frame1.Enabled = True
   Frame2.Enabled = True
   '2012/12/27 End
End Sub

Private Sub Text8_Validate(Cancel As Boolean)
   If Text8 = MsgText(601) Then
      Text8.Tag = ""
      Exit Sub
   End If
   'Add By Cheng 2003/06/25
   '若有輸入請款對象自動補滿9碼
   'Modified by Morgan 2023/12/5 第9碼若是空白或錯誤會檢查不出來
   'Me.Text8.Text = Left(Me.Text8.Text & "000000000", 9)
   'If Len(Text8) = 6 Then
   '   Text8 = AfterZero(Text8)
   'Else
   '   If Len(Text8) = 8 Then
   '      Text8 = Text8 & "0"
   '   End If
   'End If
   'If Mid(Text8, 1, 1) = "X" Then
   '   If ExistCheck("customer", "cu01", Mid(Text8, 1, 8), Label17) = False Then
   '      MsgBox MsgText(28), , MsgText(5)
   '      Cancel = True
   '      Exit Sub
   '   End If
   'Else
   '   If ExistCheck("fagent", "fa01", Mid(Text8, 1, 8), Label17) = False Then
   '      MsgBox MsgText(28), , MsgText(5)
   '      Cancel = True
   '      Exit Sub
   '   End If
   'End If
   Text8.Text = Left(Trim(Text8.Text) & "000000000", 9)
   If Left(Text8.Text, 1) = "X" Then
      If ClsPDGetCustomer(Text8.Text, strExc(1)) = False Then
         MsgBox MsgText(28), , Label15
         Cancel = True
         Exit Sub
      End If
   Else
      If ClsPDGetAgent(Text8.Text, strExc(1)) = False Then
         MsgBox MsgText(28), , Label15
         Cancel = True
         Exit Sub
      End If
   End If
   'end 2023/12/5
   
   'Add by Morgan 2010/9/10
   If Text8.Tag <> Text8.Text Then
      SetText19 'Added by Morgan 2012/9/18
      SetRemark Text7.Text, Text8.Text, Text11
      'Add By Sindy 2013/1/24
      If m_bolAfterLoad Then
         '依請款對象抓請款幣別
         'Modify By Sindy 2016/12/16 + , Text21, Text22, Text23
         'If PUB_GetDefaultCurrPrintType(Text7, Text8, "", strA1K18) <> 0 Then
         'Modified by Morgan 2018/4/27
         'If PUB_GetDefaultCurrPrintType(Text7, Text8, "", strA1K18, Text21, Text22, Text23) <> 0 Then
         If PUB_GetDefaultCurrPrintType(Text7, Text8, "", strA1K18, Text21, Text22, Text23, Text6) <> 0 Then
         'end 2018/4/27
            'FMP不可為NTD或RMB
            If bolIsFMP = True Then
               If strA1K18 = "NTD" Or strA1K18 = "RMB" Then
                  strA1K18 = ""
               End If
            End If
            Combo3 = strA1K18
            '抓請款匯率
            If strA1K18 <> "" Then
               'If Text19.Enabled = False Then
                 'Modified by Morgan 2024/1/19
                 'dblRate = PUB_GetUSXRate_1(Replace(Me.MaskEdBox1.Text, "/", ""), Me.Combo3.Text)
                 dblRate = PUB_GetRate(Replace(Me.MaskEdBox1.Text, "/", ""), Me.Combo3.Text, Text8, Text7, m_AppNo)
                 'end 2024/1/19
                 Text19 = dblRate
               'End If
            Else
               Text19 = 0
            End If
            SumShow 'Added by Morgan 2024/9/25 匯率設定後合計要重算
         End If
      End If
      '2013/1/24 End
   End If
   Text8.Tag = Text8.Text
End Sub

'*************************************************
'  代出折扣金額
'
'*************************************************
'Remove by Morgan 2011/1/17 沒用了
'Public Function DiscountShow(strValue1 As String, strValue2 As String, strValue3 As String, strValue4 As String) As String
'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open "select nvl(pa49, pa50) as Discount, nvl(pa75, pa26) as No from patent where pa01 = '" & strValue1 & "' and pa02 = '" & strValue2 & "' and pa03 = '" & strValue3 & "' and pa04 = '" & strValue4 & "' union " & _
'                  "select nvl(tm36, tm37) as Discount, nvl(tm44, tm23) as No from trademark where tm01 = '" & strValue1 & "' and tm02 = '" & strValue2 & "' and tm03 = '" & strValue3 & "' and tm04 = '" & strValue4 & "' union " & _
'                  "select lc24 as Discount, nvl(lc22, lc11) as No from lawcase where lc01 = '" & strValue1 & "' and lc02 = '" & strValue2 & "' and lc03 = '" & strValue3 & "' and lc04 = '" & strValue4 & "' union " & _
'                  "select sp31 as Discount, nvl(sp26, sp03) as No from servicepractice where sp01 = '" & strValue1 & "' and sp02 = '" & strValue2 & "' and sp03 = '" & strValue3 & "' and sp04 = '" & strValue4 & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      If IsNull(adoaccsum.Fields(0).Value) = False Then
'         adoquery.CursorLocation = adUseClient
'         adoquery.Open "select fa27 as Ldate from fagent where fa01 = '" & Mid(adoaccsum.Fields(1).Value, 1, 8) & "' and fa02 = '" & Mid(adoaccsum.Fields(1).Value, 9, 1) & "'", adoTaie, adOpenStatic, adLockReadOnly
'         If adoquery.RecordCount <> 0 Then
'            If IsNull(adoquery.Fields(0).Value) = False Then
'               If Val(strSrvDate(1)) >= Val(adoquery.Fields(0).Value) Then
'                  DiscountShow = adoaccsum.Fields(0).Value
'                  adoquery.Close
'                  adoaccsum.Close
'                  Exit Function
'               End If
'            End If
'         End If
'         adoquery.Close
'         adoquery.CursorLocation = adUseClient
'         adoquery.Open "select cu38 as Ldate from customer where cu01 = '" & Mid(adoaccsum.Fields(1).Value, 1, 8) & "' and cu02 = '" & Mid(adoaccsum.Fields(1).Value, 9, 1) & "'", adoTaie, adOpenStatic, adLockReadOnly
'         If adoquery.RecordCount <> 0 Then
'            If IsNull(adoquery.Fields(0).Value) = False Then
'               If Val(strSrvDate(1)) >= Val(adoquery.Fields(0).Value) Then
'                  DiscountShow = adoaccsum.Fields(0).Value
'                  adoquery.Close
'                  adoaccsum.Close
'                  Exit Function
'               Else
'                  DiscountShow = "0"
'               End If
'            Else
'               DiscountShow = "0"
'            End If
'         End If
'         adoquery.Close
'      End If
'      adoquery.CursorLocation = adUseClient
'      adoquery.Open "select nvl(fa25, fa26) as Discount, fa27 as Ldate from fagent where fa01 = '" & Mid(adoaccsum.Fields("No").Value, 1, 8) & "' and fa02 = '" & Mid(adoaccsum.Fields("No").Value, 9, 1) & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adoquery.RecordCount <> 0 Then
'         If IsNull(adoquery.Fields("Ldate").Value) = False Then
'            If Val(strSrvDate(1)) >= Val(adoquery.Fields("Ldate").Value) Then
'               DiscountShow = IIf(IsNull(adoquery.Fields("Discount").Value), 0, adoquery.Fields("Discount").Value)
'               adoquery.Close
'               adoaccsum.Close
'               Exit Function
'            Else
'               DiscountShow = "0"
'            End If
'         Else
'            DiscountShow = "0"
'         End If
'      End If
'      adoquery.Close
'      adoquery.CursorLocation = adUseClient
'      adoquery.Open "select nvl(cu36, cu37) as Discount, cu38 as Ldate from customer where cu01 = '" & Mid(adoaccsum.Fields("No").Value, 1, 8) & "' and cu02 = '" & Mid(adoaccsum.Fields("No").Value, 9, 1) & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adoquery.RecordCount <> 0 Then
'         If IsNull(adoquery.Fields("Ldate").Value) = False Then
'            If Val(strSrvDate(1)) >= Val(adoquery.Fields("Ldate").Value) Then
'               DiscountShow = IIf(IsNull(adoquery.Fields("Discount").Value), 0, adoquery.Fields("Discount").Value)
'            Else
'               DiscountShow = "0"
'            End If
'         Else
'            DiscountShow = "0"
'         End If
'      End If
'      adoquery.Close
'   End If
'   adoaccsum.Close
'End Function

'*************************************************
' 取得最大流水號
'
'*************************************************
Public Function GetMaxNo(strValue As String) As String
   adoaccmax.CursorLocation = adUseClient
   adoaccmax.Open "select nvl(max(a1l02), 0) as No from acc1l0 where a1l01 = '" & strValue & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccmax.RecordCount <> 0 Then
      If IsNull(adoaccmax.Fields(0).Value) Then
         GetMaxNo = ZeroBeforeNo(0, 3)
      Else
         GetMaxNo = ZeroBeforeNo(Val(adoaccmax.Fields(0).Value), 3)
      End If
   Else
      GetMaxNo = ZeroBeforeNo(0, 3)
   End If
   adoaccmax.Close
End Function

'Add By Cheng 2003/09/26
'取得案件性質
Private Function GetCP10(strCP09 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

GetCP10 = ""
'Modify by Morgan 2006/10/25 "T", "FCT", "CFT", "TF"案且相關總收文號的案件性質為102時回傳102
'StrSQLa = "Select * From CaseProgress Where CP09='" & strCP09 & "'"
StrSQLa = "Select C1.CP10,C2.CP01 CP01 From CaseProgress C1, caseprogress C2 Where C1.CP09='" & strCP09 & "' and C2.cp09(+)=C1.cp43 and C2.cp10(+)='102'"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   'Modify by Morgan 2006/10/25
   'GetCP10 = "" & rsA("CP10").Value
   Select Case "" & rsA("CP01")
      Case "T", "FCT", "CFT", "TF"
         GetCP10 = "102"
      Case Else
         GetCP10 = "" & rsA("CP10").Value
   End Select
   'end 2006/10/25
End If

If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'Removed by Morgan 2014/8/6 沒有在用
''取得請款對象
'Private Function GetA1K28(strA1k01 As String) As String
'Dim StrSQLa As String
'Dim rsA As New ADODB.Recordset
'
'GetA1K28 = ""
'StrSQLa = "Select * From Acc1k0 Where a1k01='" & strA1k01 & "' "
'rsA.CursorLocation = adUseClient
'rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'If rsA.RecordCount > 0 Then
'    GetA1K28 = "" & rsA("A1K28").Value
'End If
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'End Function

'Add By Morgan 2003/10/4
'取得商品類別數
Private Function GetSPKindCnt(strSP01 As String, strSP02 As String, strSP03 As String, strSP04 As String) As Integer

   GetSPKindCnt = 0
   'Modify by Morgan 2007/12/17 改抓商品類別
   'strSQL = "Select SP18 From ServicePractice Where " & ChgService(strSP01 & strSP02 & strSP03 & strSP04)
   strSql = "Select SP73 From ServicePractice Where " & ChgService(strSP01 & strSP02 & strSP03 & strSP04)
   'Added by Morgan 2018/5/31 也會有FCT案(Ex:FCT-39870)
   strSql = strSql & " union Select TM09 From TradeMark Where " & ChgTradeMark(strSP01 & strSP02 & strSP03 & strSP04)
   'end 2018/5/31
   'end 2007/12/17
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      '若有資料
      If .RecordCount > 0 Then
         If "" & .Fields(0).Value = "" Then
            GetSPKindCnt = 1
         Else
            GetSPKindCnt = UBound(Split("" & .Fields(0).Value, ",")) + 1
         End If
      End If
   End With
   CheckOC3
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

'Add By Cheng 2004/02/10
'取得相關總收文號的案件性質
Private Function GetRelCaseProperty(strCP60 As String, strCP10 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   GetRelCaseProperty = ""
   StrSQLa = "Select * From Caseprogress Where CP09 In ( Select CP43 From Caseprogress Where CP60='" & strCP60 & "' And CP10='" & strCP10 & "' )"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   While Not rsA.EOF
      GetRelCaseProperty = "" & rsA("CP10").Value
      rsA.MoveNext
   Wend
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Function

'Add by Morgan 2010/9/10
'Modified by Lydia 2021/12/09 TextBox=> object
Private Function SetRemark(p_Sys As String, p_AgentNo As String, p_TextBox As Object) As String
   
   'Modify by Morgan 2011/4/18 其他請款作業也要用故改抓公用函數
   'If p_AgentNo = "Y33268020" Then
   '   p_TextBox = """Services were not performed in the US"""
   'Modified by Morgan 2014/12/3
   'p_TextBox = PUB_GetDNRemark(p_AgentNo)
   p_TextBox = PUB_GetDNRemark(p_AgentNo, Text7.Text, Text21.Text, Text22.Text, Text23.Text)
   If p_TextBox <> "" Then
   
   '2010/10/25 add by sonia
   ElseIf p_AgentNo = "X55778000" Then
      'Modified by morgan 2014/7/17 新版word格式跳行後不需前置空白
      'p_TextBox = "NT$4,500 for the specification less 2 pages and drawings do not exceed 10 pages." & Chr(13) & _
                  "       Surcharge NT$300 for each additional page of specification or  each additional 5 pages of " & Chr(13) & _
                  "       drawings (less than 5 shall be counted as 5)" & Chr(13)
      'Modified by Morgan 2015/5/29 跳行也取消
      'p_TextBox = "NT$4,500 for the specification less 2 pages and drawings do not exceed 10 pages." & vbCrLf & _
                  "Surcharge NT$300 for each additional page of specification or  each additional 5 pages of " & vbCrLf & _
                  "drawings (less than 5 shall be counted as 5)"
      'Modified by Lydai 2020/08/07 改備註
      'p_TextBox = "NT$4,500 for the specification less 2 pages and drawings do not exceed 10 pages." & _
                  "Surcharge NT$300 for each additional page of specification or  each additional 5 pages of " & _
                  "drawings (less than 5 shall be counted as 5)"
      ''end 2015/5/29
      p_TextBox = "NT$4,500 for the specification less 2 pages and drawings do not exceed 10 pages." & _
                  "Surcharge NT$300 for each additional page of specification or each additional 10 pages of " & _
                  "drawings (less than 10 shall be counted as 10)"
                  
   '2010/10/25 end
   End If
End Function

'2010/10/25 ADD BY SONIA 判斷申請人是否為X55778
Private Function CUISX55778() As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   CUISX55778 = False
   'Modify By Sindy 2011/2/21 增加LC43,LC44,LC45,LC46
   'Add By Sindy 2011/2/21 增加HC05,HC24,HC25,HC26,HC27
   StrSQLa = "select PA26 CU1,PA27 CU2,PA28 CU3,PA29 CU4,PA30 CU5 from patent where pa01 = '" & Text7 & "' and pa02 = '" & Text21 & "' and pa03 = '" & Text22 & "' and pa04 = '" & Text23 & "' union " & _
                 "select tm23 CU1,tm78 CU2,tm79 CU3,tm80 CU4,tm81 CU5 from trademark where tm01 = '" & Text7 & "' and tm02 = '" & Text21 & "' and tm03 = '" & Text22 & "' and tm04 = '" & Text23 & "' union " & _
                 "select lc11 CU1,lc43 CU2,lc44 CU3,lc45 CU4,lc46 CU5 from lawcase where lc01 = '" & Text7 & "' and lc02 = '" & Text21 & "' and lc03 = '" & Text22 & "' and lc04 = '" & Text23 & "' union " & _
                 "select hc05 CU1,hc24 CU2,hc25 CU3,hc26 CU4,hc27 CU5 from hirecase where hc01 = '" & Text7 & "' and hc02 = '" & Text21 & "' and hc03 = '" & Text22 & "' and hc04 = '" & Text23 & "' union " & _
                 "select sp08 CU1,sp58 CU2,sp59 CU3,sp65 CU4,sp66 CU5 from servicepractice where sp01 = '" & Text7 & "' and sp02 = '" & Text21 & "' and sp03 = '" & Text22 & "' and sp04 = '" & Text23 & "'"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      If "" & rsA.Fields(0) = "X55778000" Or "" & rsA.Fields(1) = "X55778000" Or "" & rsA.Fields(2) = "X55778000" Or "" & rsA.Fields(3) = "X55778000" Or "" & rsA.Fields(4) = "X55778000" Then
         CUISX55778 = True
      End If
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Function
'2010/10/25 END

Private Function GetFCT10199Fee() As Long
   Dim bFirstFeeItem As Boolean, lngNetFee As Long
   
   'Modify by Morgan 2011/3/22 +商標名稱可減免計算
   lngNetFee = 3000 + Val(Text9) * IIf(Text16 = "A0199", 500, 200) - IIf(Text10 = "Y", 300, 0)
   If m_boleFiling Then
      Set RsTemp = Adodc1.Recordset.Clone
      With RsTemp
      .MoveFirst
      bFirstFeeItem = True
      Do While Not .EOF
         If .Fields("a1l02") < Text15 And (.Fields("a1l04") = "A0199" Or .Fields("a1l04") = "10199") Then
            bFirstFeeItem = False
            Exit Do
         End If
         .MoveNext
      Loop
      End With
      If bFirstFeeItem = True Then
         lngNetFee = lngNetFee - 300
      End If
   End If
   GetFCT10199Fee = lngNetFee
End Function

'Add by Morgan 2011/3/3 超過商品數
Private Sub Text9_Change()
   
   If Text7 = "FCT" And (Text16 = "10199" Or Text16 = "A0199") Then
      '請款金額
      Text18 = GetFCT10199Fee
   'Added by Morgan 2018/11/20
   ElseIf Text7 = "S" And Text16 = "0011" Then
      Text18 = 500 * Val(Text9)
   'Added by Morgan 2020/1/9
   ElseIf Text7 = "FCT" And Text16 = "1013" Then
      Text18 = 80 * Val(Text9)
   End If
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

'Add by Morgan 2011/3/2
Private Sub Text9_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub


'Added by Morgan 2013/2/20
'檢查FMP案的101,102請款必須有打字費且大於NT2000(安全基金),103請款必須大於NT2000(安全基金)
Private Function ChkFMPItem() As Boolean
   Dim dblAmount As Double
   
   m_bolFMPnewcase = False
   'Modified by Morgan 2019/8/8 沒明細時也要跳過否則會當
   'Modified by Morgan 2020/4/30
   'If strSrvDate(1) < AccFMPImputCurrStarDate Or RsTemp.RecordCount = 0 Then
   If strSrvDate(1) < AccFMPImputCurrStarDate Or Adodc1.Recordset.RecordCount = 0 Then
   'end 2020/4/30
   
      ChkFMPItem = True
      Exit Function
   Else
      '非FMP案跳過
      'Modified by Morgan 2013/8/15 改判斷國外部收文的A類的新案請款(寫成函數以便與點數分配共用)
      'strExc(0) = "select * from caseprogress where cp60='" & Text1 & "' and cp01='P' and cp12 like 'F%'"
      'intI = 1
      'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      'If intI = 0 Then
      '   ChkFMPItem = True
      '   Exit Function
      'End If
      m_bolFMPnewcase = PUB_FMPNewCaseInvoice(Text1.Text)
      If m_bolFMPnewcase = False Then
         ChkFMPItem = True
         Exit Function
      End If
      'end 2013/8/15
      
      Set RsTemp = Adodc1.Recordset.Clone
      
'Modified by Morgan 2013/8/15
'改判斷若為新案請款則服務費加總需大於2000
'      '設計
'      RsTemp.MoveFirst
'      RsTemp.Find "a1l04='103'"
'      If Not RsTemp.EOF Then
'         If RsTemp.Fields("a1l05") >= 2000 Then
'            m_bolFMPnewcase = True
'            ChkFMPItem = True
'            Exit Function
'         Else
'            MsgBox "FMP案【" & RsTemp.Fields("a1j03") & "】請款不可低於NT2000!!", vbExclamation, "FMP新案請款安全基金檢查"
'            Exit Function
'         End If
'      End If
'
'      '發明或新型
'      RsTemp.MoveFirst
'      RsTemp.Find "a1l04='101'"
'      If RsTemp.EOF Then
'         RsTemp.MoveFirst
'         RsTemp.Find "a1l04='102'"
'      End If
'      If Not RsTemp.EOF Then
'         strExc(1) = "FMP案【" & RsTemp.Fields("a1j03") & "】請款必須同時請【打字費】且不可低於NT2000!!"
'         RsTemp.MoveFirst
'         RsTemp.Find "a1l04='03'"
'         If RsTemp.EOF Then
'            MsgBox strExc(1), vbExclamation, "FMP新案請款安全基金檢查"
'            Exit Function
'         ElseIf RsTemp.Fields("a1l05") >= 2000 Then
'            m_bolFMPnewcase = True
'            ChkFMPItem = True
'            Exit Function
'         Else
'            MsgBox strExc(1), vbExclamation, "FMP新案請款安全基金檢查"
'            Exit Function
'         End If
'      Else
'         ChkFMPItem = True
'      End If
      
      dblAmount = 0
      
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         If Right(RsTemp("a1l04"), 2) <> "98" And Right(RsTemp("a1l04"), 2) <> "99" Then
            dblAmount = dblAmount + RsTemp("a1l05") - Val("" & RsTemp("a1l07"))
         End If
         RsTemp.MoveNext
      Loop
      
      If dblAmount > 2000 Then
         ChkFMPItem = True
      Else
         MsgBox "FMP新案請款服務費不可低於 NT2000!!", vbExclamation, "FMP新案請款安全基金檢查"
      End If
   End If
End Function

'Added by Morgan 2014/8/13
'檢查請款金額與收文金額是否相符--陳金蓮
Private Function ChkMoney() As Boolean
   ChkMoney = True
   '外商
   If Left(Pub_StrUserSt03, 2) = "F1" Then
      'Modofied by Morgan 2014/8/20 控制只有A類收文時才檢查--陳金蓮
      strExc(0) = "select * from caseprogress where cp60='" & Text1 & "' and cp09>'B'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 0 Then
         strExc(0) = "select * from acc1k0,(select cp60,sum(cp16) TOT,sum(cp17) FEE from caseprogress where cp60='" & Text1 & "' group by cp60) A"
         strExc(0) = strExc(0) & " where a1k01='" & Text1 & "' and cp60(+)=a1k01 and (a1k11<>nvl(TOT,0) or nvl(a1k09,0)<>nvl(FEE,0))"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(1) = ""
            If Val("" & RsTemp("a1k09")) <> Val("" & RsTemp("FEE")) Then
               strExc(1) = "規費"
            Else
               strExc(1) = "金額"
            End If
            If MsgBox("請款" & strExc(1) & "與收文" & strExc(1) & "不同是否要繼續？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
               ChkMoney = False
            End If
         End If
      End If
   End If
End Function

'Added by Morgan 2014/8/18
'整批請款存檔
Public Sub Frmacc21p1_Save()
   Dim ii As Integer
   Dim stSQL As String, intR As Integer
   
   If m_bolIsBatch Then
      For ii = LBound(strDNoArray) To UBound(strDNoArray)
         If strDNoArray(ii) > Text1 Then
            If CheckCP10(CStr(strDNoArray(ii)), m_CP10List) Then 'Added by Morgan 2016/11/1 開放多個案件性質合併請款--陳金蓮
               'Modified by Morgan 2023/8/9 規費相同的才更新,不同則要再進明細計算金額
               stSQL = "update acc1k0 a set (a1k08,a1k09,a1k10,a1k11,a1k18,a1k33)=(select b.a1k08,b.a1k09,b.a1k10,b.a1k11,b.a1k18,b.a1k33 from acc1k0 b where b.a1k01='" & Text1 & "') where a1k01='" & strDNoArray(ii) & "'" & _
                  " and exists(select 1 from (select sum(cp16) x1,sum(nvl(cp17,0)) x2 from caseprogress where cp60='" & strDNoArray(ii) & "') x" & _
                  " ,(select sum(cp16) y1,sum(nvl(cp17,0)) y2 from caseprogress where cp60='" & Text1 & "') y where y1=x1 and y2=x2)"
               adoTaie.Execute stSQL, intR
               If intR = 1 Then
                  stSQL = "delete acc1l0 where a1l01='" & strDNoArray(ii) & "'"
                  adoTaie.Execute stSQL, intR
                  stSQL = "insert into acc1l0(a1l01,a1l02,a1l03,a1l04,a1l05,a1l06,a1l07,a1l08,a1l09,a1l10,a1l11,a1l12,a1l13,a1l14,a1l15,a1l16,a1l17,a1l18)" & _
                     " select '" & strDNoArray(ii) & "',a1l02,a1l03,a1l04,a1l05,a1l06,a1l07,a1l08,a1l09,a1l10,a1l11,a1l12,a1l13,a1l14,a1l15,a1l16,a1l17,a1l18 from acc1l0 where a1l01='" & Text1 & "'"
                  adoTaie.Execute stSQL, intR
               End If
            End If 'Added by Morgan 2016/11/1
         End If
      Next
   End If
End Sub

'Added by Morgan 2016/11/1
Private Function CheckCP10(pDebitNo As String, pCP10List As String) As Boolean
   Dim arrCP10() As String
   Dim iCP10Count As Integer, iRow As Integer, iRec As Integer
   
   If pCP10List <> "" Then
      arrCP10 = Split(pCP10List, ",")
      iCP10Count = UBound(arrCP10) - LBound(arrCP10) + 1
      cnnConnection.Execute "update caseprogress set cp60=cp60 where cp60='" & pDebitNo & "'", iRec
      cnnConnection.Execute "update caseprogress set cp60=cp60 where cp60='" & pDebitNo & "' and cp10 in ('" & Replace(pCP10List, ",", "','") & "')", iRow
      '案件性質符合收文數要等於該請款單的總收文數也要等於案件性質數
      If iRow <> iRec Or iRow <> iCP10Count Then
         CheckCP10 = False
         Exit Function
      End If
   End If
   CheckCP10 = True
End Function
'Added  by Morgan 2018/11/27
'點數分配檢查
Public Sub SetAcc1n0(pDnNo As String)
   '先檢查ACC1N0是否有資料
   strSql = "SELECT * FROM ACC1N0 WHERE a1n01='" & pDnNo & "' and rownum<2"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 0 Then
      '自動點數分配
      If PUB_PointAutoassign(pDnNo) = False Then
         Frmacc21h3.Show vbModal
      End If
   End If
   
   '若請款點數有異動, 則需進入點數分配作業
   If PUB_ChkPointOk(pDnNo) = False Then
      Frmacc21h3.Show vbModal
   End If
End Sub

'Added by Morgan 2025/2/10
Private Function ChkItemFCP02(Optional pShowMsg As Boolean = False) As Boolean
   Dim strQ As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   Dim strMsg As String
   
   ChkItemFCP02 = True
   strQ = "select decode(pa181,'Y','個案',decode(fa136,'Y','代理人 '||pa75,'申請人'||decode(c1.cu202,'Y','1 '||pa26" & _
      ",decode(c2.cu202,'Y','2 '||pa27,decode(c3.cu202,'Y','3 '||pa28,decode(c4.cu202,'Y','4 '||pa29,'5 '||pa30)))))) MSG,1 SRT" & _
      " from patent,customer c1,customer c2,customer c3,customer c4,customer c5,fagent" & _
      " where pa01='" & Text7 & "' and pa02='" & Text21 & "' and pa03='" & Text22 & "' and pa04='" & Text23 & "'" & _
      " and c1.cu01(+)=substr(pa26,1,8) and c1.cu02(+)=substr(pa26,9)" & _
      " and c2.cu01(+)=substr(pa27,1,8) and c2.cu02(+)=substr(pa27,9)" & _
      " and c3.cu01(+)=substr(pa28,1,8) and c3.cu02(+)=substr(pa28,9)" & _
      " and c4.cu01(+)=substr(pa29,1,8) and c4.cu02(+)=substr(pa29,9)" & _
      " and c5.cu01(+)=substr(pa30,1,8) and c5.cu02(+)=substr(pa30,9)" & _
      " and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9) and pa181||c1.cu202||c2.cu202||c3.cu202||c4.cu202||c5.cu202||fa136 is not null" & _
      " union select '請款對象 " & Text8 & "',2 SRT from fagent where fa01='" & Left(Text8, 8) & "' and fa02='" & Mid(Text8, 9) & "' and fa136 is not null" & _
      " union select '請款對象 " & Text8 & "',2 SRT from customer where cu01='" & Left(Text8, 8) & "' and cu02='" & Mid(Text8, 9) & "' and cu202 is not null" & _
      " union select '列印對象 " & Text6 & "',3 SRT from fagent where fa01='" & Left(Text6, 8) & "' and fa02='" & Mid(Text6, 9) & "' and fa136 is not null" & _
      " union select '列印對象 " & Text6 & "',3 SRT from customer where cu01='" & Left(Text6, 8) & "' and cu02='" & Mid(Text6, 9) & "' and cu202 is not null order by SRT"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      ChkItemFCP02 = False
      If pShowMsg Then
         MsgBox "【" & rsQuery(0) & "】有設定專利不得請雜費！", vbExclamation
      End If
   End If
   Set rsQuery = Nothing
End Function

'確認其他道進度是否有規費 Or 其他道對應不到進度者及703/704 開頭
'stChoose:1-確認其他道進度是否有規費 /2-其他道對應不到進度者 /3-更新703/704 開頭及結案單款項目2碼
Public Function ChkAndSetCCDItem(stChoose As String, stNowCP09 As String, stNowCP10 As String, stInvNo As String, stA1l02 As String, ByRef stErrMsg As String) As Boolean
   Dim RsRun As New ADODB.Recordset, intRun As Integer, strRun As String, strCmd As String, strSeq As String, strShowMsg(2) As String
   Dim rsA As New ADODB.Recordset, intA As Integer, strA As String, strWhrA As String, i As Integer, arrOther
   Dim bolUpd As Boolean, stCP10 As String, stCCD04 As String, stCCD05 As String, stCCD06 As String, stA1L07 As String
On Error GoTo oErr

   ChkAndSetCCDItem = False: stErrMsg = ""
'*** 確認其他道進度是否有 規費及更新 折扣 ***
   If stChoose = "1" Then
      strShowMsg(2) = "結案單與請款單折扣是否一致"
      strRun = "Select CCD04,CCD05,CCD06 From CloseCaseDetail Where CCD01='" & stF0301 & "' And CCD02='1' And Substr(CCD04,1,3)='" & stNowCP10 & "' " & _
                      "And CCD06 is not Null Order by Decode(length(CCD04),3,1,2),CCD04 "
      intRun = 1
      Set RsRun = ClsLawReadRstMsg(intRun, strRun)
      If intRun = 1 Then
         If RsRun.RecordCount > 0 Then
            RsRun.MoveFirst
            Do While Not RsRun.EOF
               stCCD04 = "" & RsRun.Fields("CCD04")
               strA = "Select A1L02,A1L04,A1L19 From Acc1L0 Where a1L01='" & strItemNo & "' And A1L02 in('" & Replace(stA1l02, ",", "','") & "') " & _
                           "And SubStr(A1L04,1,3)='" & stNowCP10 & "' Order by Decode(A1L04,'" & RsRun.Fields("CCD04") & "',1,2),A1L04"
               intA = 1
               Set rsA = ClsLawReadRstMsg(intA, strA)
               If intA = 1 Then
                  If rsA.RecordCount > 0 Then
                     stCCD05 = "" & RsRun.Fields("CCD05") '金額
                     stCCD06 = Val("" & RsRun.Fields("CCD06")) / 100 '折扣
                     stA1L07 = Val(stCCD05) * (1 - Val(stCCD06)) '折扣金額
                     If stCCD06 <> "" & rsA.Fields("A1L19") Then
                        strCmd = "A1L07=" & stA1L07 & ",A1L19=" & stCCD06
                        strCmd = "Update Acc1L0 Set " & strCmd & " Where A1L01='" & stInvNo & "' And A1L02='" & rsA.Fields("A1L02") & "' "
                        adoTaie.Execute strCmd
                     End If
                  Else
                     strShowMsg(1) = "．結案單請款項目" & stCCD04 & "[" & GetCaseTypeName(Text7, stNowCP10, 0) & "] 有折扣,但請款單項目無可對應"
                  End If
               End If
               RsRun.MoveNext
            Loop
         End If
         If strShowMsg(0) <> "" Then stErrMsg = stErrMsg & "," & strShowMsg(2) & "有誤" & vbCrLf & Mid(strShowMsg(0), 2)
      End If
      
      strShowMsg(2) = "確認進度是否有規費"
      strRun = "Select cp09,cp10,cp16,cp17 From CaseProgress " & _
                       "Where Nvl(cp17,0)>0 And cp09 ='" & stNowCP09 & "' " & _
                       "And Not Exists(Select * From Acc1L0 Where a1L01='" & strItemNo & "' And SubStr(a1L04,1,3)=cp10 " & _
                                                    "And A1L02 in('" & Replace(stA1l02, ",", "','") & "') " & "And (SubStr(a1L04,-2)='98' Or SubStr(a1L04,-2)='99') )"
      intRun = 1
      Set RsRun = ClsLawReadRstMsg(intRun, strRun)
      If intRun = 1 Then
         If RsRun.RecordCount > 0 Then
            RsRun.MoveFirst
            Do While Not RsRun.EOF
               strShowMsg(1) = "": bolUpd = False
               stCP10 = "" & RsRun.Fields("cp10")
               stCCD05 = "" & RsRun.Fields("cp17") '規費
        '*** 確認是否有相關案件性質之規費請款項目 ***
               strWhrA = "And a1j01='" & Text7 & "' And (a1j02='" & stCP10 & "98' Or a1j02='" & stCP10 & "99' )"
               strA = "Select a1j02,A1J02Cnt From Acc1j0,(Select Count(a1j02) as A1J02Cnt From Acc1j0 Where 1=1 " & strWhrA & " ) " & "Where 1=1 " & strWhrA
               intA = 1
               Set rsA = ClsLawReadRstMsg(intA, strA)
               '只有1筆xxx98 or xxx99 才更新
               If intA = 1 Then
                  If Val("" & rsA.Fields("A1J02Cnt")) = 1 Then
                     bolUpd = True
                     stCCD04 = rsA.Fields("a1j02")
                  Else
                     strShowMsg(1) = "其 請款項目代號有xxx98及xxx99"
                  End If
               Else
                  strShowMsg(1) = "無對應之請款項目代號"
               End If
      '*** End 確認是否有相關案件性質之規費請款項目 ***
               If bolUpd = True Then
                  'CCD04=A1L04=請款項目/CCD05=A1L05=金額/CCD06=A1L19=折扣%
                  If InsertAcc1L0("A", stInvNo, stCCD04, stCCD05, "" & RsRun.Fields("CCD06"), strShowMsg(1)) = False Then
                     strShowMsg(0) = strShowMsg(0) & "," & strShowMsg(1)
                  End If
               Else
                  strShowMsg(0) = strShowMsg(0) & ",．[" & GetCaseTypeName(Text7, stCP10, 0) & "] 有規費," & strShowMsg(1)
               End If
               RsRun.MoveNext
            Loop
         End If
      End If
      If strShowMsg(0) <> "" Then stErrMsg = stErrMsg & "," & strShowMsg(2) & "有誤" & vbCrLf & Mid(strShowMsg(0), 2)
   End If
'*** End 確認進度是否有規費 ***
   
'*** 其他道對應不到進度者 ***
   If stChoose = "2" Then
      '進度沒有的其他道,直接寫入Acc1L0
      If stNotInCP10 <> "" Then
         strShowMsg(2) = "更新進度沒有的其他道"
         bolUpd = False
         
         arrOther = Split(stNotInCP10, ",")
         '只有請1道且不是703/704
         If UBound(arrOther) = 0 And stNowCP10 <> "" Then
            strA = GetAcc21H0Sql("1.1", "ChkAndSetCCDItem", stF0301, , arrOther(i), , stInvNo)
            intA = 1
            Set rsA = ClsLawReadRstMsg(intA, strA)
            If intA = 1 Then
               If rsA.RecordCount = 1 Then
                  bolUpd = True
                  strShowMsg(0) = "": strShowMsg(1) = ""
                  'CCD04=A1L04=請款項目/CCD05=A1L05=金額/CCD06=A1L19=折扣%
                  If InsertAcc1L0("E", stInvNo, "" & rsA.Fields("CCD04"), "" & rsA.Fields("CCD05"), "" & rsA.Fields("CCD06"), strShowMsg(1), stA1l02) = False Then
                     strShowMsg(0) = strShowMsg(0) & "," & strShowMsg(1)
                  End If
               End If
            End If
         End If
         
         If bolUpd = False Then
            '*** 進度沒有的其他道-多筆 ***
            For i = LBound(arrOther) To UBound(arrOther)
               strRun = GetAcc21H0Sql("1", "ChkAndSetCCDItem", stF0301, , arrOther(i), , stInvNo)
               intRun = 1
               Set RsRun = ClsLawReadRstMsg(intRun, strRun)
               If intRun = 1 Then
                  If RsRun.RecordCount > 0 Then
                     RsRun.MoveFirst
                     Do While Not RsRun.EOF
                        strShowMsg(0) = "": strShowMsg(1) = ""
                        'CCD04=A1L04=請款項目/CCD05=A1L05=金額/CCD06=A1L19=折扣%
                        If InsertAcc1L0("A", stInvNo, "" & RsRun.Fields("CCD04"), "" & RsRun.Fields("CCD05"), "" & RsRun.Fields("CCD06"), strShowMsg(1)) = False Then
                           strShowMsg(0) = strShowMsg(0) & "," & strShowMsg(1)
                        End If
                        RsRun.MoveNext
                     Loop
                  End If
               End If
            Next i
            '*** End 進度沒有的其他道-多筆 ***
         End If
         If strShowMsg(0) <> "" Then stErrMsg = stErrMsg & "," & strShowMsg(2) & "有誤" & vbCrLf & Mid(strShowMsg(0), 2)
      End If
   End If
'*** End 其他道對應不到進度者 ***

'*** 更新703/704 開頭及結案單請款項目2碼 ***
   If stChoose = "3" Then
      strShowMsg(2) = "更新703/704 開頭及請款項目2碼 "
      'Memo 此寫法需於 OpenTable 設定好 adoacc1l0 才可使用,否則於 Acc1l0Save 時會出錯
      strRun = GetAcc21H0Sql("1", Me.Name, stF0301, , stNowCP10)
      intRun = 1
      Set RsRun = ClsLawReadRstMsg(intRun, strRun)
      If intRun = 1 Then
         bolUpd = True
         RsRun.MoveFirst
         Do While Not RsRun.EOF
            stCCD04 = "" & RsRun.Fields("ccd04") '請款項目代號
            stCCD05 = "" & RsRun.Fields("ccd05") '金額
            stCCD06 = "" & RsRun.Fields("ccd06") '折扣
   
            adoadodc1.Find "a1L04 = '" & Left(stCCD04, 3) & "'"
            If adoadodc1.EOF = False Then
               '第1筆都更新703/704,其他筆才新增
               If Left(stCCD04, 3) = stNowCP10 And bolUpd = True Then
                  DataGrid1.SelBookmarks.add Adodc1.Recordset.Bookmark
                  AdodcShow
               End If
            End If
            Text16 = stCCD04 '請款項目代號
            Call Text16_Validate(False)
            Text18 = stCCD05 '金額
            Text20 = stCCD06 '折扣
            Call KeyDefine(vbKeyInsert)
            bolUpd = False
           
            strTrackMode = "" '不清空會記錄前筆資料,而第2筆無法Insert
            RsRun.MoveNext
         Loop
      End If
      If strShowMsg(0) <> "" Then stErrMsg = stErrMsg & "," & strShowMsg(2) & "有誤" & vbCrLf & Mid(strShowMsg(0), 2)
   End If
   
 '*** End 更新703/704 開頭及結案單請款項目2碼 ***
   ChkAndSetCCDItem = True
   
oErr:
   If Err.Number <> 0 Then
      Resume
      stErrMsg = "更新資料有誤(Frmacc21h1.ChkAndSet9899rOther-" & strShowMsg(2) & ")！" & vbCrLf & _
                           "請通知電腦中心" & vbCrLf & Err.Description
   End If
   Set RsRun = Nothing
   Set rsA = Nothing
End Function

Private Function InsertAcc1L0(stEditMod As String, stInvNo As String, stCCD04 As String, stCCD05 As String, stCCD06 As String, stErrMsg As String _
 , Optional ByVal stA1l02 As String) As Boolean
   Dim intRun As Integer, stSeq As String, stA1L07 As String, stCmd As String
On Error GoTo ErrAcc1L0
   
   InsertAcc1L0 = False
   If stEditMod = "A" Then stSeq = GetMaxNo(stInvNo) '項次
   'CCD04=A1L04=請款項目/CCD05=A1L05=金額/CCD06=A1L19=折扣%
   '折扣/折扣金額
   If Val(stCCD06) = 0 Then
      stCCD06 = "1"
      stA1L07 = "0"
   Else
      stCCD06 = Val(stCCD06) / 100
      stA1L07 = Val(stCCD05) * (1 - Val(stCCD06)) '折扣金額
   End If
   
   '新增
   If stEditMod = "A" Then
      stCmd = "'" & stInvNo & "','" & stSeq & "','" & Text7 & "','" & stCCD04 & "','" & stCCD05 & "' " & _
                           "," & stA1L07 & ",'" & strUserNum & "'," & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS')," & stCCD06 & " "
      stCmd = "Insert into Acc1L0 (A1L01,A1L02,A1L03,A1L04,A1L05,A1L07,A1L10,A1L08,A1L09,A1L19) Values (" & stCmd & ")"
   '修改
   Else
      stCmd = "A1L04='" & stCCD04 & "',A1L05=" & stCCD05 & ",A1L07=" & stA1L07 & ",A1L19=" & stCCD06 & " "
      stCmd = "Update Acc1L0 Set " & stCmd & "Where A1L01='" & stInvNo & "' And A1L02='" & stA1l02 & "' "
   End If
   adoTaie.Execute stCmd, intRun
   
   If intRun > 0 Then
      InsertAcc1L0 = True
   Else
      stErrMsg = "新增Acc1L0有誤(Frmacc21h1.InsertAcc1L0-請款項目" & stCCD04 & ")！" & vbCrLf & _
                           "請通知電腦中心" & vbCrLf & Err.Description
   End If
   
ErrAcc1L0:
   If Err.Number <> 0 Then
      stErrMsg = "新增Acc1L0有誤(Frmacc21h1.InsertAcc1L0)！" & vbCrLf & _
                           "請通知電腦中心" & vbCrLf & Err.Description
   End If
End Function
