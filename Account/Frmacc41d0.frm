VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc41d0 
   AutoRedraw      =   -1  'True
   Caption         =   "應收付分錄調整"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   8760
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
      Height          =   600
      Left            =   7896
      Picture         =   "Frmacc41d0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   16
      ToolTipText     =   "取消"
      Top             =   3900
      Width           =   550
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1308
      MaxLength       =   1
      TabIndex        =   0
      Top             =   60
      Width           =   612
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4056
      TabIndex        =   2
      Top             =   420
      Width           =   4032
   End
   Begin VB.TextBox Text3 
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
      Height          =   315
      Left            =   240
      TabIndex        =   23
      Top             =   3156
      Width           =   500
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '靠右對齊
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
      Left            =   3840
      MaxLength       =   14
      TabIndex        =   5
      Top             =   3156
      Width           =   1572
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  '靠右對齊
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
      Left            =   5400
      MaxLength       =   14
      TabIndex        =   6
      Top             =   3156
      Width           =   1572
   End
   Begin VB.TextBox Text11 
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
      Height          =   324
      Left            =   4212
      TabIndex        =   22
      Top             =   2316
      Width           =   1500
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
      Height          =   324
      Left            =   5700
      TabIndex        =   21
      Top             =   2316
      Width           =   1500
   End
   Begin VB.TextBox Text13 
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
      Height          =   300
      Left            =   1920
      TabIndex        =   20
      Top             =   60
      Width           =   6492
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
      Height          =   300
      Left            =   2400
      MaxLength       =   9
      TabIndex        =   10
      Top             =   3876
      Width           =   1572
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
      Height          =   300
      Left            =   5616
      MaxLength       =   5
      TabIndex        =   11
      Top             =   3900
      Width           =   1572
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5616
      MaxLength       =   10
      TabIndex        =   13
      Top             =   4260
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2400
      MaxLength       =   12
      TabIndex        =   8
      Top             =   3516
      Width           =   1572
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H00C0FFFF&
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
      Left            =   828
      MaxLength       =   6
      TabIndex        =   4
      Top             =   3156
      Width           =   972
   End
   Begin VB.TextBox Text16 
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
      Left            =   6960
      MaxLength       =   3
      TabIndex        =   7
      Top             =   3156
      Width           =   612
   End
   Begin VB.TextBox Text17 
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
      Height          =   315
      Left            =   7560
      TabIndex        =   18
      Top             =   3156
      Width           =   950
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
      Height          =   600
      Left            =   7308
      Picture         =   "Frmacc41d0.frx":066A
      Style           =   1  '圖片外觀
      TabIndex        =   15
      ToolTipText     =   "清除畫面"
      Top             =   3900
      Width           =   550
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   8040
      Picture         =   "Frmacc41d0.frx":0F34
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   420
      Width           =   350
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
      Height          =   300
      Left            =   2400
      MaxLength       =   9
      TabIndex        =   12
      Top             =   4248
      Width           =   1572
   End
   Begin VB.TextBox Text18 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   14
      Top             =   4572
      Visible         =   0   'False
      Width           =   540
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1296
      TabIndex        =   1
      Top             =   420
      Width           =   1572
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc41d0.frx":1036
      Height          =   1500
      Left            =   252
      TabIndex        =   17
      Top             =   768
      Width           =   8292
      _ExtentX        =   14631
      _ExtentY        =   2646
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
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
      ColumnCount     =   10
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
            SubFormatType   =   1
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
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
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
      BeginProperty Column06 
         DataField       =   "a1p06"
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
      BeginProperty Column07 
         DataField       =   "a1p15"
         Caption         =   "對沖代號(客)"
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
         DataField       =   "a1p16"
         Caption         =   "對沖代號(業)"
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
         DataField       =   "a1p17"
         Caption         =   "對沖代號(本所案號)"
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
            Alignment       =   2
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2399.811
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1470.047
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3390.236
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1950.236
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   660
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
   Begin MSForms.TextBox Text15 
      Height          =   315
      Left            =   1800
      TabIndex        =   19
      Top             =   3156
      Width           =   2000
      VariousPropertyBits=   679493657
      BackColor       =   14737632
      MaxLength       =   8
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   312
      Left            =   4776
      TabIndex        =   9
      Top             =   3516
      Width           =   3765
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "7646;591"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "入帳日期"
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
      TabIndex        =   39
      Top             =   408
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "單據編號"
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
      Left            =   3048
      TabIndex        =   38
      Top             =   432
      Width           =   972
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
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
      TabIndex        =   37
      Top             =   60
      Width           =   852
   End
   Begin VB.Label Label4 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "項次"
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
      Left            =   240
      TabIndex        =   36
      Top             =   2916
      Width           =   492
   End
   Begin VB.Label Label5 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "會計科目"
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
      Left            =   1560
      TabIndex        =   35
      Top             =   2916
      Width           =   972
   End
   Begin VB.Label Label6 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "借方金額"
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
      Left            =   4080
      TabIndex        =   34
      Top             =   2910
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "貸方金額"
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
      Left            =   5640
      TabIndex        =   33
      Top             =   2910
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "部門別"
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
      Left            =   7200
      TabIndex        =   32
      Top             =   2916
      Width           =   852
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "摘要"
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
      Left            =   4188
      TabIndex        =   31
      Top             =   3552
      Width           =   612
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   2196
      Left            =   120
      Top             =   2808
      Width           =   8532
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4716
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label15 
      Alignment       =   2  '置中對齊
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
      Left            =   3264
      TabIndex        =   30
      Top             =   2316
      Width           =   732
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(客)"
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
      Left            =   240
      TabIndex        =   29
      Top             =   3876
      Width           =   1452
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(業)"
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
      Left            =   4176
      TabIndex        =   28
      Top             =   3900
      Width           =   1452
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "沖帳傳票號碼"
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
      Left            =   4176
      TabIndex        =   27
      Top             =   4260
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(本所案號)"
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
      Left            =   240
      TabIndex        =   26
      Top             =   3516
      Width           =   2172
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(其它)"
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
      Left            =   240
      TabIndex        =   25
      Top             =   4272
      Width           =   1632
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "作帳公司"
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
      Left            =   228
      TabIndex        =   24
      Top             =   4620
      Visible         =   0   'False
      Width           =   1632
   End
End
Attribute VB_Name = "Frmacc41d0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/25 Form2.0已修改 Text15/Combo1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc020 As New ADODB.Recordset
Public adoacc021 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adocheck As New ADODB.Recordset
Public stra1p02 As String
Dim adocase As New ADODB.Recordset
Dim strDocuNo As String

'add by nickc 2007/07/13 將輸入法改成使用API
Private Sub Combo1_GotFocus()
OpenIme
End Sub
'add by nickc 2007/07/13 將輸入法改成使用API
Private Sub Combo1_Validate(Cancel As Boolean)
CloseIme
End Sub

Private Sub Command1_Click()
   If Adodc1.Recordset.RecordCount <> 0 Then
      adoTaie.Execute "delete from acc1p0 where a1p01 = '" & Adodc1.Recordset.Fields("a1p01").Value & "' and a1p04 = '" & Adodc1.Recordset.Fields("a1p04").Value & "' and a1p03 = '" & Adodc1.Recordset.Fields("a1p03").Value & "'"
      AdodcRefresh
      AdodcClear
      SumShow
   End If
End Sub

Private Sub Command2_Click()
Dim adoaccmax As New ADODB.Recordset

   AdodcClear
   adoaccmax.CursorLocation = adUseClient
   adoaccmax.Open "select max(a1p03) from acc1p0 where a1p01 = '" & Text1 & "' and a1p04 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccmax.RecordCount <> 0 Then
      Text3 = ZeroBeforeNo(Val(adoaccmax.Fields(0).Value), 3)
   Else
      Text3 = ZeroBeforeNo(0, 3)
   End If
   adoaccmax.Close
   SumShow
   Text14.SetFocus
End Sub

Private Sub Command3_Click()
   'If adoacc020.RecordCount = 0 Or Text1 = MsgText(601) Or Text2 = MsgText(601) Then
   '   Exit Sub
   'End If
   'adoacc020.Find "a0201 = '" & Text1 & "'", 0, adSearchForward, 1
   'If adoacc020.EOF = False Then
   '   adoacc020.Find "a0202 = '" & Text2 & "'", 0, adSearchForward, adoacc020.Bookmark
   '   If adoacc020.EOF Then
   '      MsgBox MsgText(33), , MsgText(5)
   '      adoacc020.MoveFirst
   '   End If
   'Else
   '   MsgBox MsgText(33), , MsgText(5)
   '   adoacc020.MoveFirst
   'End If
   'Modify by Amy 2014/02/11
   If ChkData = True Then
      MsgBox "資料為多筆,請至應付分錄調整查詢", , MsgText(5)
      Frmacc41d0_Clear
      stra1p02 = ""
   End If
   Acc020Refresh
   If adoacc020.RecordCount <> 0 Then
      FormShow
      AdodcRefresh
      SumShow
      RecordShow
   Else
      AdodcRefresh
      Text11 = ""
      Text12 = ""
      AdodcClear
   End If
   'end 2014/02/11
End Sub

Private Sub Command3_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command3_Click
         Exit Sub
   End Select
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   AdodcShow
'   SumShow
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strCompanyNo = MsgText(601) Then
      Exit Sub
   End If
   'adoacc020.Find "a0201 = '" & strCompanyNo & "'", 0, adSearchForward, 1
   'If adoacc020.EOF = False Then
   '   adoacc020.Find "a0202 = '" & strItemNo & "'", 0, adSearchForward, adoacc020.Bookmark
   '   If adoacc020.EOF = False Then
   '      FormShow
   '      AdodcRefresh
   '      SumShow
   '      RecordShow
   '   End If
   'End If
   Text1 = strCompanyNo
   Text2 = strItemNo
   'Remove by Lydia 2019/11/07 Frmacc41d1回傳
   'stra1p02 = strExc(0) 'Add by Amy 2014/02/06
   Acc020Refresh , True
   If adoacc020.RecordCount <> 0 Then
      FormShow
      AdodcRefresh
      SumShow
      RecordShow
   End If
   strCompanyNo = MsgText(601)
   strItemNo = MsgText(601)
   'Remove by Lydia 2019/11/07 執行後,清空變數
   'strExc(0) = MsgText(601) 'Add by Amy 2014/02/06
   stra1p02 = MsgText(601)
End Sub

'Add by Amy 2021/10/25
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Form2.0 記錄鍵盤傳入順序
End Sub

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
   Me.Height = 5500
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   Text1 = MsgText(601)
   Text2 = MsgText(601)
   MaskEdBox1.Mask = DFormat
   OpenTable
   If adoacc020.RecordCount <> 0 Then
      RecordShow
   End If
   FormDisabled
   'Modify by Amy 2014/02/11 新增會錯/刪除會run Frmacc4120_Delete 但帶空值不會刪 因不會用到鎖住按鈕
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   CreDebCheck
   If CreDebCheck <> MsgText(602) Then
      tool1_enabled
      MsgBox MsgText(11), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
   StatusClear
   strConTitle = MsgText(601)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   strTrackMode = "" 'Add by Amy 2021/10/25 Form2.0 記錄鍵盤傳入順序(清除)
   Set Frmacc41d0 = Nothing
End Sub

'2014/1/23 cancel by sonia
'Private Sub MaskEdBox1_Change()
'   If strSaveConfirm <> MsgText(3) Then
'      Exit Sub
'   End If
'   Text2 = AccAutoNo(MsgText(801), 4, Val(Mid(MaskEdBox1.Text, 1, 3)), Val(Mid(MaskEdBox1.Text, 5, 2)))
'End Sub
'2014/1/23 end

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   'Add by Amy 2014/02/11
   If strSaveConfirm = MsgText(601) Then
        Exit Sub
   End If
   'end 2014/02/11
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox Label9 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label9 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Text1_Change()
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
   'Add by Amy 2020/04/14 判斷只可輸作帳公司
   If InStr(GetBookKeepCmp, Text1) = 0 Then
      Text13 = ""
      Exit Sub
   End If
   'end 2020/04/14
   Text13 = A0802Query(Text1)
'2014/1/23 cancel by sonia
'   If strSaveConfirm <> MsgText(3) Then
'      Exit Sub
'   End If
'   Text2 = AccAutoNo(MsgText(801), 4, Val(Mid(MaskEdBox1.Text, 1, 3)), Val(Mid(MaskEdBox1.Text, 5, 2)))
'2014/1/23 end
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc020.CursorLocation = adUseClient
   'Modify by Morgan 2004/10/26 不必抓資料
   adoacc020.MaxRecords = intMax
   'adoacc020.Open "select distinct a1p01, a1p04, a1p18, a1p22 from acc1p0 where a1p02 in ('Z', 'E', 'W', 'L', 'G') and a1p04 >= '" & Text2 & "' order by a1p01 asc, a1p04 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoacc020.Open "select distinct a1p01, a1p04, a1p18, a1p22 from acc1p0 where rownum<1", adoTaie, adOpenStatic, adLockReadOnly
   adoacc021.CursorLocation = adUseClient
   'adoacc021.Open "select * from acc1p0 where a1p01 = '" & Text1 & "' and a1p04 = '" & Text2 & "' and a1p03 = '" & Text3 & "' order by a1p01 asc, a1p04 asc, a1p03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc021.Open "select * from acc1p0 where rownum<1", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   'adoadodc1.Open "select * from acc1p0, acc010 where a1p05 = a0101 (+) and a1p01 = '" & Text1 & "' and a1p04 = '" & Text2 & "' order by a1p01 asc, a1p04 asc, a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc1.Open "select * from acc1p0, acc010 where rownum<1", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示欄位資料(傳票資料--主檔)
'
'*************************************************
Public Sub FormShow()
   If IsNull(adoacc020.Fields("a1p01").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = adoacc020.Fields("a1p01").Value
   End If
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc020.Fields("a1p18").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(Trim(str(adoacc020.Fields("a1p18").Value)))
   End If
   MaskEdBox1.Mask = DFormat
   If IsNull(adoacc020.Fields("a1p04").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = adoacc020.Fields("a1p04").Value
   End If
   stra1p02 = adoacc020.Fields("a1p02").Value
End Sub

'*************************************************
'  儲存欄位資料(傳票資料--交易檔)
'
'*************************************************
Public Sub acc021Save()
Dim strCombo1 As String

On Error GoTo Checking
   If Text14 = MsgText(601) Then
      MsgBox MsgText(10) & Label5, , MsgText(5)
      strControlButton = MsgText(602)
      Text14.SetFocus
      Exit Sub
   Else
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select * from acc010 where a0101 = '" & Text14 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount = 0 Then
         MessageShow Label5
         strControlButton = MsgText(602)
         adocheck.Close
         Text14.SetFocus
         Exit Sub
      Else
         If IsNull(adocheck.Fields("a0105").Value) And Left(Text14, 1) = "6" Then
            If Text16 = MsgText(601) Or Text16 = MsgText(55) Then
               MsgBox MsgText(198), , MsgText(5)
               strControlButton = MsgText(602)
               adocheck.Close
               Text16.SetFocus
            End If
         End If
      End If
      adocheck.Close
      If Val(Text4) <> 0 And Val(Text5) <> 0 Then
         MsgBox MsgText(47) & MsgText(46), , MsgText(5)
         strControlButton = MsgText(602)
         Text4.SetFocus
         Exit Sub
      End If
      If Val(Text4) = 0 And Val(Text5) = 0 Then
         MsgBox MsgText(58) & MsgText(46), , MsgText(5)
         strControlButton = MsgText(602)
         Text4.SetFocus
         Exit Sub
      End If
      If CheckDept(Text14, Text16) = False Then
         MsgBox MsgText(103), , MsgText(5)
         strControlButton = MsgText(602)
         Text16.SetFocus
         Exit Sub
      End If
      If Text16 <> MsgText(601) Then
         adocheck.CursorLocation = adUseClient
         adocheck.Open "select a0901 from acc090 where a0901 = '" & Text16 & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adocheck.RecordCount = 0 Then
            MessageShow Label8
            strControlButton = MsgText(602)
            adocheck.Close
            Text16.SetFocus
            Exit Sub
         End If
         adocheck.Close
      End If
      If Text7 <> MsgText(601) Then
         adocheck.CursorLocation = adUseClient
         'modify by sonia 2021/2/1 +FAGENT檢查
         adocheck.Open "select cu01 as Name from customer where cu01 = '" & Mid(Text7, 1, 8) & "' union " & _
                       "select fa01 as Name from fagent where fa01 = '" & Mid(Text7, 1, 8) & "' union " & _
                       "select a0i01 as Name from acc0i0 where a0i01 = '" & Text7 & "' union " & _
                       "select st01 as Name from staff where st01 = '" & Text7 & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adocheck.RecordCount = 0 Then
            MessageShow Label11
            strControlButton = MsgText(602)
            adocheck.Close
            Text7.SetFocus
            Exit Sub
         End If
         adocheck.Close
      End If
      If Text8 <> MsgText(601) Then
         adocheck.CursorLocation = adUseClient
         adocheck.Open "select st01 from staff where st01 = '" & Text8 & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adocheck.RecordCount = 0 Then
            MessageShow Label12
            strControlButton = MsgText(602)
            adocheck.Close
            Text8.SetFocus
            Exit Sub
         End If
         adocheck.Close
      End If
      If Text10 <> MsgText(601) Then
         adocheck.CursorLocation = adUseClient
         If Len(Mid(Text10, 2, Len(Text10) - 1)) > 6 Then
            adocheck.Open "select cp09 from caseprogress where cp01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and cp02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and cp03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and cp04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
            If adocheck.RecordCount = 0 Then
               MessageShow Label14
               strControlButton = MsgText(602)
               adocheck.Close
               Text10.SetFocus
               Exit Sub
            End If
            adocheck.Close
         Else
            MessageShow Label14
            strControlButton = MsgText(602)
            Text10.SetFocus
            Exit Sub
         End If
      End If
      
      'add by sonia 2015/12/30 檢查民國105年起法務收入科目不可使用
      intI = PUB_AccNoEnable(Text14, Val(FCDate(MaskEdBox1.Text)))
      If intI <> 0 Then
         strControlButton = MsgText(602)
         Text14.SetFocus
         Exit Sub
      End If
      'end 2015/12/30
      'Add by Morgan 2007/10/2 檢查科目部門&智權人員是否正確
      intI = PUB_AccNoGood(Text14, Text16, Text8)
      If intI <> 0 Then
         strControlButton = MsgText(602)
         If intI = 1 Then
            Text14.SetFocus
         ElseIf intI = 2 Then
            Text16.SetFocus
         ElseIf intI = 3 Then
            Text8.SetFocus
         End If
         Exit Sub
      End If
      'end 2007/10/2
   End If
   adoacc021.Close
   adoacc021.CursorLocation = adUseClient
   'Modify by Amy 2014/02/06 +a1p02
   'adoacc021.Open "select * from acc1p0 where a1p01 = '" & Text1 & "' and a1p04 = '" & Text2 & "' and a1p03 = '" & Text3 & "' order by a1p01 asc, a1p04 asc, a1p03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc021.Open "select * from acc1p0 where a1p01 = '" & Text1 & "' and a1p04 = '" & Text2 & "' and a1p03 = '" & Text3 & "' And a1p02='" & stra1p02 & "' order by a1p01 asc, a1p04 asc, a1p03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc021.RecordCount = 0 Then
      adoacc021.AddNew
   End If
   adoacc021.Fields("a1p01").Value = Text1
   adoacc021.Fields("a1p02").Value = stra1p02
   adoacc021.Fields("a1p04").Value = Text2
   adoacc021.Fields("a1p03").Value = Text3
   If Text14 <> MsgText(601) Then
      adoacc021.Fields("a1p05").Value = Text14
   Else
      adoacc021.Fields("a1p05").Value = Null
   End If
   If Text4 <> MsgText(601) Then
      adoacc021.Fields("a1p07").Value = Val(Text4)
   Else
      adoacc021.Fields("a1p07").Value = 0
   End If
   If Text5 <> MsgText(601) Then
      adoacc021.Fields("a1p08").Value = Val(Text5)
   Else
      adoacc021.Fields("a1p08").Value = 0
   End If
   If Text16 <> MsgText(601) Then
      adoacc021.Fields("a1p06").Value = Text16
   Else
      adoacc021.Fields("a1p06").Value = MsgText(55)
   End If
   If Combo1 <> MsgText(601) Then
      adoacc021.Fields("a1p14").Value = Replace(Combo1, "'", "''")
      strCombo1 = Combo1
      Combo1.Clear
      Combo1.AddItem strCombo1
   Else
      adoacc021.Fields("a1p14").Value = Null
   End If
   If Text7 <> MsgText(601) Then
      adoacc021.Fields("a1p15").Value = Text7
   Else
      adoacc021.Fields("a1p15").Value = Null
   End If
   If Text8 <> MsgText(601) Then
      adoacc021.Fields("a1p16").Value = Text8
   Else
      adoacc021.Fields("a1p16").Value = Null
   End If
   
   'Add by Morgan 2005/1/10 入帳日期
   If MaskEdBox1.Text <> MsgText(29) Then
      adoacc021.Fields("a1p18").Value = Val(ChangeTDateStringToTString(MaskEdBox1.Text))
   Else
      adoacc021.Fields("a1p18").Value = Null
   End If
   
   If Text6 <> MsgText(601) Then
      adoacc021.Fields("a1p30").Value = Text6
   Else
      adoacc021.Fields("a1p30").Value = Null
   End If
   Text10 = CaseNoZero(Text10)
   If Text10 <> MsgText(601) Then
      adoacc021.Fields("a1p17").Value = Text10
   Else
      adoacc021.Fields("a1p17").Value = Null
   End If
   'Mark by Amy 2014/02/11 取消作帳公司
'   If Text18 <> MsgText(601) Then
'      adoacc021.Fields("a1p31").Value = Text18
'   Else
'      adoacc021.Fields("a1p31").Value = Null
'   End If
   'end 2014/02/11
   If strDocuNo <> MsgText(601) Then
      adoacc021.Fields("a1p22").Value = adoacc020.Fields("a1p22").Value
      adoacc021.Fields("a1p27").Value = MsgText(602)
      'adoTaie.Execute "update acc1p0 set a1p27 = 'Y' where a1p01 = '" & Text1 & "' and a1p22 = '" & strDocuNo & "'"
   Else
      adoacc021.Fields("a1p22").Value = Null
   End If
   adoacc021.UpdateBatch
   adoTaie.Execute "update acc1p0 set a1p27 = 'Y' where a1p01 = '" & Text1 & "' and a1p22 = '" & strDocuNo & "'"
   AdodcRefresh
'   Adodc1.Recordset.Find "ax203 = '" & Text3 & "'", 0, adSearchForward, 1
'   If Adodc1.Recordset.EOF Then
'      Adodc1.Recordset.MoveFirst
'   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料(傳票資料--交易檔)
'
'*************************************************
Public Sub AdodcShow()
   If IsNull(Adodc1.Recordset.Fields("a1p03").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = Adodc1.Recordset.Fields("a1p03").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p05").Value) Then
      Text14 = MsgText(601)
   Else
      Text14 = Adodc1.Recordset.Fields("a1p05").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p07").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = Adodc1.Recordset.Fields("a1p07").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p08").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = Adodc1.Recordset.Fields("a1p08").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p06").Value) Then
      Text16 = MsgText(601)
      Text17 = MsgText(601)
   Else
      If Adodc1.Recordset.Fields("a1p06").Value = MsgText(55) Then
         Text16 = MsgText(601)
         Text17 = MsgText(601)
      Else
         Text16 = Adodc1.Recordset.Fields("a1p06").Value
      End If
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p14").Value) Then
      Combo1 = MsgText(601)
   Else
      Combo1 = Adodc1.Recordset.Fields("a1p14").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p15").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = Adodc1.Recordset.Fields("a1p15").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p16").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = Adodc1.Recordset.Fields("a1p16").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p30").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = Adodc1.Recordset.Fields("a1p30").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p17").Value) Then
      Text10 = MsgText(601)
   Else
      Text10 = Adodc1.Recordset.Fields("a1p17").Value
   End If
   'Mark by Amy 2014/02/07 取消作帳公司
'   If IsNull(Adodc1.Recordset.Fields("a1p31").Value) Then
'      Text18 = MsgText(601)
'   Else
'      Text18 = Adodc1.Recordset.Fields("a1p31").Value
'   End If
    'end 2014/02/07
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   If Text1 <> MsgText(601) Then
     If InStr(GetBookKeepCmp, Text1) = 0 Then
         MsgBox Label3 & MsgText(63), , MsgText(5)
        Cancel = True
        Text1.SetFocus
        Exit Sub
     End If
   End If
   If ExistCheck("acc080", "a0801", Text1, Label3, False) = False Then
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
On Error GoTo Checking
   If Text10 <> MsgText(601) Then
      Text10 = CaseNoZero(Text10)
      adocase.CursorLocation = adUseClient
      adocase.Open "select cp09 from caseprogress where cp01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and cp02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and cp03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and cp04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocase.RecordCount = 0 Then
         MsgBox MsgText(28) & Label14, , MsgText(5)
         Cancel = True
         adocase.Close
         Exit Sub
      End If
      adocase.Close
      'add by sonia 2021/1/28 以本所案號以判別FCP,FCT英日文組
      If AccNoToSalesNo(Text14, Text10) <> "" Then
         Text8 = AccNoToSalesNo(Text14, Text10)
      End If
      'end 2021/1/28
   End If
   QueryCustomer
   Exit Sub
Checking:
   MsgBox MsgText(128), , MsgText(5)
   Exit Sub
End Sub

Private Sub Text14_Change()
   If Text14 = MsgText(601) Then
      Exit Sub
   End If
   Text15 = A0102Query(Text14)
End Sub

Private Sub Text14_GotFocus()
   TextInverse Text14
End Sub

Private Sub Text14_Validate(Cancel As Boolean)
   If Text14 <> MsgText(601) Then
      If ExistCheck("acc010", "a0101", Text14, Label5) = False Then
         Cancel = True
         Exit Sub
      End If
   End If
   'modify by sonia 2021/1/28 加傳本所案號以判別FCP,FCT英日文組
   'If AccNoToSalesNo(Text14) <> "" Then
   '   Text8 = AccNoToSalesNo(Text14)
   If AccNoToSalesNo(Text14, Text10) <> "" Then
      Text8 = AccNoToSalesNo(Text14, Text10)
   'end 2021/1/28
   End If
End Sub

Private Sub Text16_Change()
   If Text16 = MsgText(601) Then
      Exit Sub
   End If
   Text17 = A0902Query(Text16)
End Sub

Private Sub Text16_GotFocus()
   TextInverse Text16
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text16_Validate(Cancel As Boolean)
   If Text16 <> MsgText(601) Then
      If ExistCheck("acc090", "a0901", Text16, Label8) = False Then
         Cancel = True
         Exit Sub
      End If
   End If
   If CheckDept(Text14, Text16) = False Then
      MsgBox MsgText(103), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
End Sub

'Mark by Amy 2014/02/07 取消作帳公司
'Private Sub Text18_GotFocus()
'   TextInverse Text18
'End Sub
'
'Private Sub Text18_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub
'end 2014/02/07

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  清除欄位資料(傳票資料--交易檔)
'
'*************************************************
Public Sub AdodcClear()
   Text14 = ""
   Text15 = ""
   Text4 = ""
   Text5 = ""
   Text16 = ""
   Text17 = ""
   Combo1 = ""
   Text7 = ""
   Text8 = ""
   Text9 = ""
   Text6 = ""
   Text10 = ""
   'Text18 = "" 'Mark by Amy 2014/02/07 取消作帳公司
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   'Add by Amy 2021/10/25
   Call PUB_SaveTrackMode(1, KeyCode)
    
   'Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
   If PUB_ChkTrackMode = False Then
        Exit Sub
   End If
   'end 2021/10/25
 
   Select Case KeyCode
      Case vbKeyInsert
         If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
            Exit Sub
         End If
         If strControlButton <> MsgText(602) Then
            acc021Save
         End If
         If strControlButton <> MsgText(602) Then
            AdodcClear
            If adocheck.State = adStateOpen Then
               adocheck.Close
            End If
            adocheck.CursorLocation = adUseClient
            adocheck.Open "select max(a1p03) from acc1p0 where a1p01 = '" & Text1 & "' and a1p04 = '" & Text2 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
            If adocheck.RecordCount <> 0 Then
               If IsNull(adocheck.Fields(0).Value) Then
                  Text3 = MsgText(601)
               Else
                  Text3 = adocheck.Fields(0).Value
               End If
            Else
               Text3 = MsgText(601)
            End If
            adocheck.Close
            Text3 = ZeroBeforeNo(Text3, 3)
            SumShow
            Text14.SetFocus
         End If
         strControlButton = MsgText(601)
      Case vbKeyDown
         If Adodc1.Recordset.EOF = False Then
            Adodc1.Recordset.MoveNext
            If Adodc1.Recordset.EOF = False Then
               AdodcShow
            End If
         End If
      Case vbKeyUp
         If Adodc1.Recordset.BOF = False Then
            Adodc1.Recordset.MovePrevious
            If Adodc1.Recordset.BOF = False Then
               AdodcShow
            End If
         End If
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 = MsgText(601) Then
      MsgBox Label2 & MsgText(52), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
   If Text7 <> MsgText(601) Then
      If Len(Text7) = 6 Then
         Text7 = AfterZero(Text7)
      'Add by Morgan 2007/3/1 八碼時要補'0'
      ElseIf Len(Text7) = 8 Then
         Text7 = Text7 & "0"
      'End 2007/3/1
      End If
      If ExistCheck("customer", "cu01", Mid(Text7, 1, 8), Label11, False) = False Then
         If ExistCheck("acc0i0", "a0i01", Text7, Label11, False) = False Then
            If ExistCheck("staff", "st01", Text7, Label11, False) = False Then
               If ExistCheck("fagent", "fa01", Mid(Text7, 1, 8), Label11, False) = False Then   'add by sonia 2021/2/1 +FAGENT檢查
                  MsgBox MsgText(28) & Label11, , MsgText(5)
                  Cancel = True
                  Exit Sub
               End If   'add by sonia 2021/2/1
            End If
         End If
      End If
   End If
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text8_Validate(Cancel As Boolean)
   If Text8 <> MsgText(601) Then
      If ExistCheck("staff", "st01", Text8, Label12) = False Then
         Cancel = True
         Exit Sub
      End If
      'add by sonia 2021/1/28
      If SalesNoCheckAccNo(Text14, Text8) = False Then
      End If
      'end 2021/1/28
   End If
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  顯示欄位資料(傳票合計)
'
'*************************************************
Public Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(a1p07), sum(a1p08) from acc1p0 where a1p01 = '" & Text1 & "' and a1p04 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text11 = MsgText(601)
      Else
         Text11 = Format(adoaccsum.Fields(0).Value, FDollar)
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text12 = MsgText(601)
      Else
         Text12 = Format(adoaccsum.Fields(1).Value, FDollar)
      End If
   Else
      Text11 = MsgText(601)
      Text12 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  借貸方檢核
'
'*************************************************
Public Function CreDebCheck() As String
   If Text11 = Text12 Then
      CreDebCheck = MsgText(602)
   End If
End Function

'*************************************************
'  關閉分錄欄位輸入狀態
'
'*************************************************
Public Sub FormDisabled()
   Text14.Enabled = False
   Text4.Enabled = False
   Text5.Enabled = False
   Text16.Enabled = False
   Combo1.Enabled = False
   Text7.Enabled = False
   Text8.Enabled = False
   Text9.Enabled = False
   Text6.Enabled = False
   Text10.Enabled = False
   'Text18.Enabled = False 'Mark by Amy 2014/02/07 取消作帳公司
   Command1.Enabled = False
   Command2.Enabled = False
End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'
'*************************************************
Public Sub FormEnabled()
   'Add by Amy 2014/02/11 不可修改公司別及單據編號
   Text1.Enabled = False
   Text2.Enabled = False
   'end 2014/02/11
   Text14.Enabled = True
   Text4.Enabled = True
   Text5.Enabled = True
   Text16.Enabled = True
   Combo1.Enabled = True
   Text7.Enabled = True
   Text8.Enabled = True
   Text9.Enabled = True
   Text6.Enabled = True
   Text10.Enabled = True
   'Text18.Enabled = True 'Mark by Amy 2014/02/07 取消作帳公司
   Command1.Enabled = True
   Command2.Enabled = True
End Sub

'*************************************************
'  重新整理傳票資料
'
'*************************************************
'Modify by Amy 2014/02/10 加參數IsBack '由Frmacc41d1查詢返回
'Modify by Morgan 2004/10/27 加參數iMov 0:a1p04=Text2, 1:a1p04>=Text2
Public Sub Acc020Refresh(Optional ByVal iMov As Integer = 0, Optional ByVal IsBack As Boolean = False)
   Dim strSql As String
   Screen.MousePointer = vbHourglass
On Error GoTo Checking
   If adoacc020.State = adStateOpen Then
      adoacc020.Close
   End If
   adoacc020.CursorLocation = adUseClient
   adoacc020.MaxRecords = intMax
   'Modify by Morgan 2004/10/27
   'adoacc020.Open "select distinct a1p01, a1p04, a1p18, a1p22 from acc1p0 where a1p02 in ('Z', 'E', 'W', 'L', 'G') and a1p04 >= '" & Text2 & "' order by a1p01 asc, a1p04 asc", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2014/02/11 改語法修正同單據編號 a1p02及公司別不同抓錯的問題(塑贊有限公司1021)
   If iMov = 0 Then
      'adoacc020.Open "select distinct a1p01, a1p04, a1p18, a1p22 from acc1p0 where ''||a1p02 in ('Z', 'E', 'W', 'L', 'G') and a1p04 >= '" & Text2 & "' and rownum<2", adoTaie, adOpenStatic, adLockReadOnly
      If IsBack = True Then
        strSql = "select distinct a1p01, a1p04, a1p18, a1p22,a1p02 from acc1p0 where a1p01='" & Text1 & "' and a1p02 ='" & stra1p02 & "' and a1p04 = '" & Text2 & "' "
      Else
        strSql = "select distinct a1p01, a1p04, a1p18, a1p22,a1p02 from acc1p0 where a1p04 = '" & Text2 & "' "
      End If
      adoacc020.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   Else
      'adoacc020.Open "select distinct a1p01, a1p04, a1p18, a1p22 from acc1p0 where ''||a1p02 in ('Z', 'E', 'W', 'L', 'G') and a1p04 >= '" & Text2 & "' order by a1p01 asc, a1p04 asc", adoTaie, adOpenStatic, adLockReadOnly
      strSql = "select distinct a1p01, a1p04, a1p18, a1p22,a1p02 from acc1p0 where ''||a1p02 in ('Z', 'E', 'W', 'L', 'G') and a1p04 >= '" & Text2 & "' order by a1p04 asc,a1p01 asc"
      adoacc020.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   End If
   'end 2014/02/11
Checking:
   Screen.MousePointer = vbDefault
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   
End Sub

'**************************************************
'  重新整理 Adodc 之資料
'
'**************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   adoacc021.Close
   adoacc021.CursorLocation = adUseClient
   'Modify by Amy 2014/02/07 +a1p02
   'adoacc021.Open "select * from acc1p0 where a1p01 = '" & Text1 & "' and a1p04 = '" & Text2 & "' and a1p03 = '" & Text3 & "' order by a1p01 asc, a1p04 asc, a1p03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc021.Open "select * from acc1p0 where a1p01 = '" & Text1 & "' and a1p04 = '" & Text2 & "' and a1p03 = '" & Text3 & "' And a1p02='" & stra1p02 & "' order by a1p01 asc, a1p04 asc, a1p03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   'adoadodc1.Open "select * from acc1p0, acc010 where a1p05 = a0101 (+) and a1p01 = '" & Text1 & "' and a1p04 = '" & Text2 & "' order by a1p01 asc, a1p04 asc, a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc1.Open "select * from acc1p0, acc010 where a1p05 = a0101 (+) and a1p01 = '" & Text1 & "' and a1p04 = '" & Text2 & "' And a1p02='" & stra1p02 & "' order by a1p01 asc, a1p04 asc, a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
   'end 2014/02/07
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      'stra1p02 = Adodc1.Recordset.Fields("a1p02").Value 'Mark by Amy 2014/02/07
      If IsNull(Adodc1.Recordset.Fields("a1p22").Value) Then
         strDocuNo = MsgText(601)
      Else
         strDocuNo = Adodc1.Recordset.Fields("a1p22").Value
      End If
      Adodc1.Recordset.Find "a1p03 = '" & Text3 & "'", 0, adSearchForward, 1
      If Adodc1.Recordset.EOF Then
         Exit Sub
      Else
         DataGrid1.SelBookmarks.add Adodc1.Recordset.Bookmark
      End If
   End If
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
On Error GoTo Checking
   Frmacc0000.StatusBar1.Panels(2).Text = adoacc020.Bookmark & MsgText(35) & adoacc020.RecordCount
Checking:
   Exit Sub
End Sub

'*************************************************
'  以本所案號查詢客戶名稱
'
'*************************************************
Public Sub QueryCustomer()
Dim strSql As String

   If Text10 = MsgText(601) Then
      Exit Sub
   End If
   strSql = "select cu01||cu02 as Name, cu04, cu05, cu06, cu88, cu89, cu90 from patent, customer where substr(pa26, 1, 8) = cu01 and nvl(substr(pa26, 9, 1), '0') = cu02 and pa01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and pa02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and pa03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and pa04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
            "select cu01||cu02 as Name, cu04, cu05, cu06, cu88, cu89, cu90 from trademark, customer where substr(tm23, 1, 8) = cu01 and nvl(substr(tm23, 9, 1), '0') = cu02 and tm01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and tm02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and tm03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and tm04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
            "select cu01||cu02 as Name, cu04, cu05, cu06, cu88, cu89, cu90 from lawcase, customer where substr(lc11, 1, 8) = cu01 and nvl(substr(lc11, 9, 1), '0') = cu02 and lc01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and lc02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and lc03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and lc04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
            "select cu01||cu02 as Name, cu04, cu05, cu06, cu88, cu89, cu90 from hirecase, customer where substr(hc05, 1, 8) = cu01 and nvl(substr(hc05, 9, 1), '0') = cu02 and hc01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and hc02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and hc03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and hc04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
            "select cu01||cu02 as Name, cu04, cu05, cu06, cu88, cu89, cu90 from servicepractice, customer where substr(sp08, 1, 8) = cu01 and nvl(substr(sp08, 9, 1), '0') = cu02 and sp01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and sp02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and sp03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and sp04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "'"
   adocase.CursorLocation = adUseClient
   adocase.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adocase.RecordCount <> 0 Then
      If IsNull(adocase.Fields(0).Value) Then
         Text7 = MsgText(601)
      Else
         Text7 = adocase.Fields(0).Value
      End If
      If IsNull(adocase.Fields("cu04").Value) Then
         If IsNull(adocase.Fields("cu05").Value) Then
            If IsNull(adocase.Fields("cu06").Value) Then
               Combo1 = MsgText(601)
            Else
               Combo1 = adocase.Fields("cu06").Value
            End If
         Else
            Combo1 = adocase.Fields("cu05").Value
            If IsNull(adocase.Fields("cu88").Value) = False Then
               Combo1 = Combo1 & adocase.Fields("cu88").Value
            End If
            If IsNull(adocase.Fields("cu89").Value) = False Then
               Combo1 = Combo1 & adocase.Fields("cu89").Value
            End If
            If IsNull(adocase.Fields("cu90").Value) = False Then
               Combo1 = Combo1 & adocase.Fields("cu90").Value
            End If
         End If
      Else
         Combo1 = adocase.Fields("cu04").Value
      End If
   Else
      Text7 = MsgText(601)
      Combo1 = MsgText(601)
   End If
   adocase.Close
End Sub

'Add by Amy 2014/02/06 由bas 搬回
Public Sub FormF2Set()
    CreDebCheck
    If CreDebCheck <> MsgText(602) Then
        MsgBox MsgText(11), , MsgText(5)
        strSaveConfirm = MsgText(601)
        Exit Sub
    End If
    Text1.Enabled = False
    Text2.Enabled = False
    Combo1.Clear
    Frmacc41d0_Clear
    AdodcClear
    AdodcRefresh
    Text1 = "1"
    If MaskEdBox1.Text = MsgText(29) Or MaskEdBox1.Text = MsgText(601) Then
        MaskEdBox1.Mask = MsgText(601)
        MaskEdBox1.Text = CFDate(strSrvDate(2))
        MaskEdBox1.Mask = DFormat
    End If
    FormEnabled
    Text3 = ZeroBeforeNo(MsgText(12), 3)
    SumShow
    Text14.SetFocus
End Sub

Public Sub Frmacc41d0_Clear()
    Text1 = ""
    Text13 = ""
    If MaskEdBox1.Text = MsgText(29) Or MaskEdBox1.Text = MsgText(601) Then
        MaskEdBox1.Mask = ""
        MaskEdBox1.Text = ""
        MaskEdBox1.Mask = DFormat
    End If
    Text2 = ""
    Text11 = ""
    Text12 = ""
    'MaskEdBox1.SetFocus 'Modify by Amy 2014/02/11
End Sub

Public Sub Frmacc41d0_First()
    CreDebCheck
    If CreDebCheck <> MsgText(602) Then
        MsgBox MsgText(11), , MsgText(5)
        Exit Sub
    End If
    If adoacc020.RecordCount <> 0 Then
        adoacc020.MoveFirst
        FormShow
        AdodcRefresh
        SumShow
    End If
    AdodcClear
    RecordShow
End Sub

Public Sub Frmacc41d0_Last()
    CreDebCheck
    If CreDebCheck <> MsgText(602) Then
        MsgBox MsgText(11), , MsgText(5)
        Exit Sub
    End If
    If adoacc020.RecordCount <> 0 Then
         adoacc020.MoveLast
         FormShow
         AdodcRefresh
         SumShow
    End If
    AdodcClear
    RecordShow
End Sub

Public Sub Frmacc41d0_Next()
      CreDebCheck
      If CreDebCheck <> MsgText(602) Then
         MsgBox MsgText(11), , MsgText(5)
         Exit Sub
      End If
      If adoacc020.EOF = False Then
         adoacc020.MoveNext
         'Add by Morgan 2004/10/27
         If adoacc020.EOF Then
            Acc020Refresh 1
            If adoacc020.RecordCount > 0 Then
               adoacc020.MoveNext
            Else
               Exit Sub
            End If
         End If
         '2004/10/27 end
         If adoacc020.EOF Then
            adoacc020.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         FormShow
         AdodcRefresh
         SumShow
      End If
      AdodcClear
      RecordShow
End Sub

Public Sub Frmacc41d0_Previous()
      CreDebCheck
      If CreDebCheck <> MsgText(602) Then
         MsgBox MsgText(11), , MsgText(5)
         Exit Sub
      End If
      If adoacc020.BOF = False Then
         adoacc020.MovePrevious
         If adoacc020.BOF Then
            adoacc020.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         FormShow
         AdodcRefresh
         SumShow
         RecordShow
      End If
      AdodcClear
End Sub

'Add by Amy 2014/02/11
'依單據編號(a1p04)判斷是否有重覆資料
Private Function ChkData() As Boolean
    Dim StrSqlChk As String
    ChkData = False
    StrSqlChk = "Select count(distinct a1p01) From acc1p0 Where ''||a1p02 in ('Z', 'E', 'W', 'L', 'G') and a1p04 = '" & Text2 & "' Having count(Distinct a1p01) >1 " & _
       "Union All Select count(distinct a1p02) From acc1p0 Where ''||a1p02 in ('Z', 'E', 'W', 'L', 'G') and a1p04 = '" & Text2 & "' Having count(Distinct a1p02) >1 "
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, StrSqlChk)
    If intI = 1 Then
        ChkData = True
    End If
End Function
'end 2014/02/11

'Add by Amy 2020/04/14 從aacc_var搬回來修改,加作帳公司判斷
Public Function FormCheck() As Boolean
    Dim bCancel As Boolean
    
    If Text1 <> MsgText(601) Then
        Call Text1_Validate(bCancel)
        If bCancel = True Then
            strControlButton = MsgText(602)
            Exit Function
        End If
    End If
    If CreDebCheck <> MsgText(602) Or Val(Text11) = 0 Or Val(Text12) = 0 Then
        MsgBox MsgText(11), , MsgText(5)
        strControlButton = MsgText(602)
        Exit Function
    End If
    FormDisabled
    Text1.Enabled = True
    Text2.Enabled = True
    Text1.SetFocus
    FormCheck = True
End Function

