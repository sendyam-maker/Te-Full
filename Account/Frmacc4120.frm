VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc4120 
   AutoRedraw      =   -1  'True
   Caption         =   "傳票輸入"
   ClientHeight    =   5112
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5112
   ScaleWidth      =   8760
   Begin VB.TextBox Text3 
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
      Height          =   315
      Left            =   240
      TabIndex        =   43
      Top             =   3216
      Width           =   500
   End
   Begin VB.CommandButton CmdChgComp 
      Appearance      =   0  '平面
      Caption         =   "更換公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7050
      TabIndex        =   42
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "發票明細"
      Height          =   300
      Left            =   252
      TabIndex        =   41
      Top             =   2400
      Width           =   1200
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
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   15
      Top             =   4632
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   8040
      Picture         =   "Frmacc4120.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   480
      Width           =   350
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
      Height          =   600
      Left            =   7440
      Picture         =   "Frmacc4120.frx":0102
      Style           =   1  '圖片外觀
      TabIndex        =   16
      ToolTipText     =   "清除畫面"
      Top             =   3960
      Width           =   550
   End
   Begin VB.TextBox Text17 
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
      Height          =   312
      Left            =   7560
      TabIndex        =   36
      Top             =   3216
      Width           =   950
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
      Height          =   312
      Left            =   6960
      MaxLength       =   3
      TabIndex        =   7
      Top             =   3216
      Width           =   612
   End
   Begin VB.TextBox Text14 
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
      Height          =   315
      Left            =   810
      MaxLength       =   6
      TabIndex        =   4
      Top             =   3216
      Width           =   972
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
      Left            =   2400
      MaxLength       =   12
      TabIndex        =   8
      Top             =   3576
      Width           =   1572
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
      Height          =   315
      Left            =   5616
      MaxLength       =   10
      TabIndex        =   14
      Top             =   4320
      Width           =   1710
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
      Height          =   315
      Left            =   5616
      MaxLength       =   5
      TabIndex        =   11
      Top             =   3960
      Width           =   765
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
      Height          =   300
      Left            =   2415
      MaxLength       =   9
      TabIndex        =   10
      Top             =   3936
      Width           =   1572
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
      Height          =   324
      Left            =   5688
      TabIndex        =   29
      Top             =   2376
      Width           =   1500
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
      Height          =   324
      Left            =   4212
      TabIndex        =   28
      Top             =   2376
      Width           =   1500
   End
   Begin VB.TextBox Text5 
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
      Height          =   312
      Left            =   5400
      MaxLength       =   14
      TabIndex        =   6
      Top             =   3216
      Width           =   1572
   End
   Begin VB.TextBox Text4 
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
      Height          =   312
      Left            =   3840
      MaxLength       =   14
      TabIndex        =   5
      Top             =   3216
      Width           =   1572
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
      Left            =   6510
      MaxLength       =   10
      TabIndex        =   2
      Top             =   480
      Width           =   1572
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
      Left            =   1308
      MaxLength       =   1
      TabIndex        =   0
      Top             =   120
      Width           =   612
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
      Height          =   600
      Left            =   8025
      Picture         =   "Frmacc4120.frx":09CC
      Style           =   1  '圖片外觀
      TabIndex        =   17
      ToolTipText     =   "取消"
      Top             =   3960
      Width           =   550
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1296
      TabIndex        =   1
      Top             =   480
      Width           =   1572
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc4120.frx":1036
      Height          =   1500
      Left            =   252
      TabIndex        =   39
      Top             =   828
      Width           =   8292
      _ExtentX        =   14626
      _ExtentY        =   2646
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
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "ax203"
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
         DataField       =   "ax205"
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
         DataField       =   "ax206"
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
         DataField       =   "ax207"
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
         DataField       =   "ax212"
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
         DataField       =   "ax204"
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
         DataField       =   "ax208"
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
         DataField       =   "ax209"
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
         DataField       =   "ax214"
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
      BeginProperty Column10 
         DataField       =   "ax213"
         Caption         =   "對沖代號(其他)"
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
            ColumnWidth     =   492.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   852.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1307.906
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1307.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3395.906
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   708.095
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1307.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1955.906
         EndProperty
         BeginProperty Column10 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   720
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
   Begin MSForms.TextBox Text6 
      Height          =   300
      Left            =   2400
      TabIndex        =   13
      Top             =   4308
      Width           =   1572
      VariousPropertyBits=   679493659
      MaxLength       =   10
      Size            =   "10927;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text19 
      Height          =   315
      Left            =   6345
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3960
      Width           =   960
      VariousPropertyBits=   679493661
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
      Height          =   330
      Left            =   4776
      TabIndex        =   9
      Top             =   3576
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
   Begin MSForms.TextBox Text15 
      Height          =   315
      Left            =   1800
      TabIndex        =   35
      Top             =   3216
      Width           =   2000
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text13 
      Height          =   315
      Left            =   1920
      TabIndex        =   30
      Top             =   120
      Width           =   5000
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "美國公開費之退費摘要中, 退公開費 四個字不可分開"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   3250
      TabIndex        =   40
      Top             =   4680
      Width           =   5160
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "作帳公司"
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
      Left            =   228
      TabIndex        =   38
      Top             =   4680
      Visible         =   0   'False
      Width           =   1632
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(其它)"
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
      Left            =   240
      TabIndex        =   37
      Top             =   4332
      Width           =   1632
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(本所案號)"
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
      Left            =   240
      TabIndex        =   34
      Top             =   3576
      Width           =   2172
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "沖帳傳票號碼"
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
      Left            =   4125
      TabIndex        =   33
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(業)"
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
      Left            =   4110
      TabIndex        =   32
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(客)"
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
      Left            =   240
      TabIndex        =   31
      Top             =   3936
      Width           =   1452
   End
   Begin VB.Label Label15 
      Alignment       =   2  '置中對齊
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
      Height          =   252
      Left            =   3264
      TabIndex        =   27
      Top             =   2376
      Width           =   732
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4776
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   2196
      Left            =   120
      Top             =   2868
      Width           =   8532
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "摘要"
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
      Left            =   4140
      TabIndex        =   26
      Top             =   3615
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   2  '置中對齊
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
      Left            =   7320
      TabIndex        =   25
      Top             =   2970
      Width           =   855
   End
   Begin VB.Label Label7 
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
      Left            =   5640
      TabIndex        =   24
      Top             =   2970
      Width           =   1095
   End
   Begin VB.Label Label6 
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
      Left            =   4080
      TabIndex        =   23
      Top             =   2970
      Width           =   1095
   End
   Begin VB.Label Label5 
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
      Left            =   1680
      TabIndex        =   22
      Top             =   2970
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "項次"
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
      Left            =   240
      TabIndex        =   21
      Top             =   2976
      Width           =   492
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
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
      TabIndex        =   20
      Top             =   120
      Width           =   852
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "傳票編號"
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
      TabIndex        =   19
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "傳票日期"
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
      TabIndex        =   18
      Top             =   480
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc4120"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/07 Form2.0已修改 Text13/Text15/Text19/Combo1/DataGrid1/Text6(1110607改)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit

Public adoacc020 As New ADODB.Recordset
Public adoacc021 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adocheck As New ADODB.Recordset
Dim adocase As New ADODB.Recordset
Dim adocase1 As New ADODB.Recordset
Dim strA1R01 As String 'Add by Amy 2014/01/15 設定字串以取得傳票編號
Dim strSpecComp As String 'Add by Amy 2014/11/17 特殊出名公司
Dim strAxb(1 To 17) As String 'Add by Amy 2017/04/18
Dim HasUpdTag As Boolean, strContent As String 'Add by Amy 2021/09/23 有修改結餘資料/發信內文
Public bolF3 As Boolean 'Add by Amy 2022/05/16
Dim m_CurrKEY As String, strNation As String 'Add by Amy 2024/08/05 傳票號/國籍
Dim bolFirst As Boolean 'Added by Lydia 2025/08/08 True = Focus在摘要Combo，在按下Insert鍵時先重送Insert鍵於第2次才執行更新明細

'Add by Amy 2014/01/15
Private Sub CmdChgComp_Click()
   Dim strNewComp As String
   
   'Modify by Amy 2014/02/20 改不用InputBox輸公司別
'   Text1.Tag = Text1
'   Call ChgCompany(Text1)
'   strNewComp = UCase(strExc(0))
    Frmacc41c2.Show vbModal
    strNewComp = strCompanyNo
    'end 2014/02/20
    If strNewComp <> "" And strNewComp <> Me.Text1 Then
        Text1 = strNewComp
        MaskEdBox1.Mask = MsgText(601)
        MaskEdBox1.Mask = DFormat
        Text2 = MsgText(601)
        AdodcClear
        Acc020Refresh
        AdodcRefresh
        
        '依公司別設定
        If Text1 = "J" Then
            strA1R01 = MsgText(819) 'JD
            cmdDetail.Visible = True
        'Add by Amy 2020/12/24 +L公司
        ElseIf Text1 = "L" Then
            strA1R01 = MsgText(820) 'LD
            cmdDetail.Visible = False
        Else
            strA1R01 = MsgText(801) 'D
            cmdDetail.Visible = False
        End If
    End If
    strCompanyNo = MsgText(601) 'Add by Amy 2014/02/20
    m_CurrKEY = MsgText(601) 'Add by Amy 2024/08/05
End Sub

Private Sub cmdDetail_Click()
    If strSaveConfirm = MsgText(3) Then
        Frmacc4120_Save
    End If
    
    Frmacc1172.strBackForm = "Frmacc4120"  '記錄返回畫面
    Frmacc1172.Show
    Screen.MousePointer = vbDefault
    Me.Hide
End Sub
'end 2014/01/15

'add by nickc 2007/07/13 將輸入法改成使用API
Private Sub Combo1_GotFocus()
   OpenIme
   bolFirst = True  'Added by Lydia 2025/08/08 記錄Form 2.0元件的Focus
End Sub

Private Sub Combo1_LostFocus()
   bolFirst = False  'Added by Lydia 2025/08/08 記錄Form 2.0元件的Focus
End Sub

'add by nickc 2007/07/13 將輸入法改成使用API
Private Sub Combo1_Validate(Cancel As Boolean)
   CloseIme
End Sub

'剪刀
Private Sub Command1_Click()
Dim BookThisRec 'Add By Sindy 2013/12/19
   If Adodc1.Recordset.RecordCount <> 0 Then
      If Not IsNull(Adodc1.Recordset.Fields("ax210").Value) Then
         MsgBox MsgText(14), , MsgText(5)
         Text14.SetFocus
         Exit Sub
      End If
      'Add By Sindy 2013/12/19
      Adodc1.Recordset.MovePrevious
      If Adodc1.Recordset.BOF Then
         Adodc1.Recordset.MoveFirst
         BookThisRec = Adodc1.Recordset.Bookmark
      Else
         BookThisRec = Adodc1.Recordset.Bookmark
         Adodc1.Recordset.MoveNext
      End If
      '2013/12/19 END
      Call ChkSetAxb16("Del") 'Add by Amy 2021/09/23 判斷結餘資料寫Tag
      adoTaie.Execute "delete from acc021 where ax201 = '" & Adodc1.Recordset.Fields("ax201").Value & "' and ax202 = '" & Adodc1.Recordset.Fields("ax202").Value & "' and ax203 = '" & Adodc1.Recordset.Fields("ax203").Value & "'"
      AdodcRefresh
      'Modify by Amy 2017/05/12 +if 若剪完最後一筆會Errror
      If Adodc1.Recordset.BOF = False Then Adodc1.Recordset.Bookmark = BookThisRec 'Add By Sindy 2013/12/19
      AdodcClear
      Text3 = GetSeqNo(Text1, Text2) 'Add by Amy 2014/01/15 重抓流水,若沒重抓按完剪刀鈕再insert會錯
      SumShow
   End If
End Sub

'垃圾桶
Private Sub Command2_Click()
'Dim adoaccmax As New ADODB.Recordset
   AdodcClear
   'Modify by Amy 2014/01/15 改寫至function
'   adoaccmax.CursorLocation = adUseClient
'   adoaccmax.Open "select max(ax203) from acc021 where ax201 = '" & Text1 & "' and ax202 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccmax.RecordCount <> 0 Then
'      Text3 = ZeroBeforeNo(Val(adoaccmax.Fields(0).Value), 3)
'   Else
'      Text3 = ZeroBeforeNo(0, 3)
'   End If
'   adoaccmax.Close
   Text3 = GetSeqNo(Text1, Text2)
   'end 2014/01/15
   SumShow
   Text14.SetFocus
End Sub

'Modify by Amy 2017/04/18 原:Private
'查詢傳票號
Public Sub Command3_Click()
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
    
   'Add by Amy 2014/01/15
   If Text2 = MsgText(601) Then
      'Modify by Amy 2024/08/05 原MsgText(181)
      MsgBox "輸入" & Label2 & "再按查詢", , MsgText(5)
      Exit Sub
    End If
    'end 2014/01/15
    'Add by Amy 2017/04/18
    'Modify by Amy 2023/04/07 +frmacc41L0
    If UCase(Me.Tag) = "FRMACC41G0" Or UCase(Me.Tag) = "FRMACC41H0" Or UCase(Me.Tag) = "FRMACC41L0" Then
        strSaveConfirm = MsgText(4)
        Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = True
        'Modify by Amy 2023/05/16 +if '按修改->Insert->取消->修改->Insert->存檔會出現「找不到要更新的資料列。最後取的值已被變更」,無法修改
        If bolF3 = True Then KeyEnter vbKeyF3
    End If
    
   Acc020Refresh
   If adoacc020.RecordCount <> 0 Then
      FormShow
      AdodcRefresh
      SumShow
      RecordShow
   'Add by Amy 2024/08/05
      m_CurrKEY = Text2
   Else
      MsgBox MsgText(33), , MsgText(5)
   End If
   strControlButton = ""
   'end 2024/08/05
End Sub

Private Sub Command3_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command3_Click
         Exit Sub
   End Select
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
    
   'Modify By Sindy 2013/12/19 加EOF判斷
   If Not Adodc1.Recordset.EOF Then
      AdodcShow
      SumShow
   End If
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
   
   If strSaveConfirm = MsgText(601) Then 'Add by Amy 2014/01/15
        Text2.SetFocus
   End If
  
   Text1 = strCompanyNo
   Text2 = strItemNo
   Acc020Refresh
   If adoacc020.RecordCount <> 0 Then
      FormShow
      AdodcRefresh
      SumShow
      RecordShow
      'Add by Amy 2014/01/15 +if
       If Text1 = "J" Then
            strA1R01 = MsgText(819) 'JD
            cmdDetail.Visible = True
       'Add by Amy 2020/03/17 L公司
       ElseIf Text1 = "L" Then
            strA1R01 = MsgText(820) 'LD
            cmdDetail.Visible = False
       Else
            strA1R01 = MsgText(801) 'D
            cmdDetail.Visible = False
       End If
    
       If strSaveConfirm = MsgText(601) Then
           '由查詢或frmacc1172/返回且strSaveConfirm=""時 (與frmacc1170共用1172)
            FormDisabled
       Else
            FormEnabled
       End If
       'end 2014/01/15
       m_CurrKEY = Text2 'Add by Amy 2024/08/05
   End If
   strCompanyNo = MsgText(601)
   strItemNo = MsgText(601)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  ' Add by Amy 2021/12/07 Form2.0 記錄鍵盤傳入順序
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
   Me.Width = 8976
   Me.Height = 5676
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'Modify by Amy 2014/01/15 從Menu 輸公司別並鎖住
   Text1 = strCompanyNo '原:MsgText(601)
   Text1.Enabled = False
   strCompanyNo = MsgText(601)
   
    If Text1 = "J" Then
        strA1R01 = MsgText(819) 'JD
        cmdDetail.Visible = True
    ElseIf Text1 = "L" Then
        strA1R01 = MsgText(820) 'LD
        cmdDetail.Visible = False
    Else
        strA1R01 = MsgText(801) 'D
        cmdDetail.Visible = False
    End If
   'end 2014/01/15
   Text2 = MsgText(601)
   MaskEdBox1.Mask = DFormat
   OpenTable
   If adoacc020.RecordCount <> 0 Then
      RecordShow
   End If
   FormDisabled
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   CreDebCheck
   If CreDebCheck <> MsgText(602) Then
      'Modify by Amy 2023/04/07 +frmacc41L0
      If UCase(Me.Tag) = "FRMACC41H0" Or UCase(Me.Tag) = "FRMACC41G0" Or UCase(Me.Tag) = "FRMACC41L0" Then
            MenuDisabled
            tool7_enabled
      Else
         'Modify by Amy 2024/08/05 程式已判斷傳票不連號問題,故開放都可使用刪除
'         'Modify by Amy 2024/02/05 M51可使用刪除
'         If Pub_StrUserSt03 = "M51" Then
            tool1_enabled
'         Else
'            'Add by Amy 2023/12/06 避免傳票不連號,取消刪除鈕 原:tool1_enabled
'            tool14_enabled
'         End If
      End If
      MsgBox MsgText(11), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
   'Add by Amy 2017/04/18
   If UCase(Me.Tag) = "FRMACC41H0" Then
      '自動產生結餘傳票4191科目之ax204為null 不可離開
      If ChkAx204Null = True Then
        MsgBox "結餘轉撥傳票請補輸入部門欄 ！", , MsgText(5)
        Cancel = True
        MenuDisabled
        tool7_enabled
        Exit Sub
      End If
   End If
   'end 2017/04/18
   'Add by Amy 2022/05/16 結餘轉撥傳票存檔時需檢查每個人的總額必須與SalesPoint相符
    If UCase(Me.Tag) = "FRMACC41H0" Then
        If ChkSalesPointVal = True Then
            Cancel = True
            MenuDisabled
            tool7_enabled
            Exit Sub
        End If
    End If
    
   StatusClear
   strConTitle = MsgText(601)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   If Len(Me.Tag) = 0 Then MenuEnabled
   'Modify by Amy 2017/10/11
   If UCase(Me.Tag) = "FRMACC41G0" Then
       strFormName = "Frmacc41g0"
       tool1_enabled
       Frmacc41g0.Show
   ElseIf UCase(Me.Tag) = "FRMACC41H0" Then
       strFormName = "Frmacc41h0"
       tool1_enabled
       Frmacc41h0.Show
   'Add by Amy 2023/04/07 ACS分潤
   ElseIf UCase(Me.Tag) = "FRMACC41L0" Then
       strFormName = "Frmacc41l0"
       tool1_enabled
       Frmacc41l0.Show
   End If
   Call PUB_GetLock("", "Frmacc4120")
   'end 2017/10/11
   Me.Tag = MsgText(601)
   strTrackMode = "" 'Add by Amy 2021/12/07 Form2.0 記錄鍵盤傳入順序(清除)
   Set Frmacc4120 = Nothing
End Sub

Private Sub MaskEdBox1_Change()
   'Modify by Amy 2024/08/05 修改日期時,strControlButton設回空,避免無法點Grid
   If strSaveConfirm = MsgText(4) And strControlButton = MsgText(602) Then
      strControlButton = MsgText(601)
   End If
   'end 2024/08/05
   
'   If strSaveConfirm <> MsgText(3) Then
'      Exit Sub
'   End If
   'Modify by Amy 2014/01/15 改至setformF2 做 MsgText(801) 改strA1R01
   'Text2 = AccAutoNo(strA1R01, 4, Val(Mid(MaskEdBox1.Text, 1, 3)), Val(Mid(MaskEdBox1.Text, 5, 2)))
  
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   'Modify by Amy 2023/03/07 避免按新增後已產生傳票號又去改日期,造成日期與傳票號不連號
'   'Add by Amy 2014/01/15
'   If strSaveConfirm = MsgText(601) Then
'      Exit Sub
'   End If
'   'end 2014/01/15
'   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
'      MsgBox Label9 & MsgText(52), vbExclamation, "日期錯誤！"
'      Cancel = True
'      MaskEdBox1.SetFocus
'      Exit Sub
'   End If
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
        Exit Sub
   End If
   'end 2023/03/07
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label9 & MsgText(63), vbExclamation, "日期錯誤！"
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   'Add by Amy 2014/11/20 +工作日判斷
   If ChkWorkDay(FCDate(MaskEdBox1.Text) + 19110000) = False Then
      MsgBox Label9 & "請輸入工作日！", vbExclamation, "日期錯誤！"
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   'Memo by Amy 2024/08/05 原ChkA0205檢查改至ChkForm,讓新增時可使用
   '查D113070029 (傳票日7/1)->改日期(7/19)->按新增鈕 ->彈 只能輸7/2(因先觸發此檢查)->抓到新的傳票號,卻仍可操作
   'end 2014/11/20
End Sub

Private Sub Text1_Change()
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
  
   Text13 = A0802Query(Text1)
   If strSaveConfirm <> MsgText(3) Then
      Exit Sub
   End If
    'Modify by Amy 2014/01/15 改至setformF2 做MsgText(801) 改strA1R01
    'Text2 = AccAutoNo(strA1R01, 4, Val(Mid(MaskEdBox1.Text, 1, 3)), Val(Mid(MaskEdBox1.Text, 5, 2)))
   
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
   adoacc020.MaxRecords = intMax
   'Modify by Amy 2014/01/15 +公司別
   adoacc020.Open "select * from acc020 where a0201='" & Text1 & "' And a0202 >= '" & Text2 & "' order by a0201 asc, a0202 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc021.CursorLocation = adUseClient
   adoacc021.Open "select * from acc021 where ax201 = '" & Text1 & "' and ax202 = '" & Text2 & "' and ax203 = '" & Text3 & "' order by ax201 asc, ax202 asc, ax203 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc021, acc010 where ax205 = a0101 (+) and ax201 = '" & Text1 & "' and ax202 = '" & Text2 & "' order by ax201 asc, ax202 asc, ax203 asc", adoTaie, adOpenStatic, adLockReadOnly
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
   If IsNull(adoacc020.Fields("a0201").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = adoacc020.Fields("a0201").Value
   End If
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc020.Fields("a0205").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(Trim(str(adoacc020.Fields("a0205").Value)))
   End If
   MaskEdBox1.Tag = "" & Trim(str(adoacc020.Fields("a0205").Value)) 'Add byAmy 2014/11/20
   MaskEdBox1.Mask = DFormat
   If IsNull(adoacc020.Fields("a0202").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = adoacc020.Fields("a0202").Value
   End If
End Sub

'*************************************************
'  儲存欄位資料(傳票資料--交易檔)
'
'*************************************************
Private Sub acc021Save()
Dim strCombo1 As String
Dim StrSQLa As String 'Add by Amy 2014/07/10

On Error GoTo Checking
'Mark by Amy 2024/08/05 改至ChkForm
'   If Text14 = MsgText(601) Then
'      MsgBox MsgText(10) & Label5, , MsgText(5)
'      strControlButton = MsgText(602)
'      Text14.SetFocus
'      Exit Sub
'   Else
'      adocheck.CursorLocation = adUseClient
'      adocheck.Open "select * from acc010 where a0101 = '" & Text14 & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adocheck.RecordCount = 0 Then
'         MessageShow Label5
'         strControlButton = MsgText(602)
'         adocheck.Close
'         Text14.SetFocus
'         Exit Sub
'      Else
'         If IsNull(adocheck.Fields("a0105").Value) And Left(Text14, 1) = "6" Then
'            If Text16 = MsgText(601) Or Text16 = MsgText(55) Then
'               MsgBox MsgText(198), , MsgText(5)
'               strControlButton = MsgText(602)
'               adocheck.Close
'               Text16.SetFocus
'            End If
'         End If
'      End If
'      adocheck.Close
'      If Val(Text4) <> 0 And Val(Text5) <> 0 Then
'         MsgBox MsgText(47) & MsgText(46), , MsgText(5)
'         strControlButton = MsgText(602)
'         Text4.SetFocus
'         Exit Sub
'      End If
'      If Val(Text4) = 0 And Val(Text5) = 0 Then
'         MsgBox MsgText(58) & MsgText(46), , MsgText(5)
'         strControlButton = MsgText(602)
'         Text4.SetFocus
'         Exit Sub
'      End If
'      'Modify by Amy 2021/05/17 +FormName
'      If CheckDept(Text14, Text16, Me.Name) = False Then
'         MsgBox MsgText(103), , MsgText(5)
'         strControlButton = MsgText(602)
'         Text16.SetFocus
'         Exit Sub
'      End If
'      If Text16 <> MsgText(601) Then
'         adocheck.CursorLocation = adUseClient
'         adocheck.Open "select a0901 from acc090 where a0901 = '" & Text16 & "'", adoTaie, adOpenStatic, adLockReadOnly
'         If adocheck.RecordCount = 0 Then
'            MessageShow Label8
'            strControlButton = MsgText(602)
'            adocheck.Close
'            Text16.SetFocus
'            Exit Sub
'         End If
'         adocheck.Close
'      End If
'      If Text7 <> MsgText(601) Then
'         adocheck.CursorLocation = adUseClient
'         adocheck.Open "select cu01 as Name from customer where cu01 = '" & Mid(Text7, 1, 8) & "' union " & _
'                       "select a0i01 as Name from acc0i0 where a0i01 = '" & Text7 & "' union " & _
'                       "select st01 as Name from staff where st01 = '" & Text7 & "' union " & _
'                       "select fa01 as Name from fagent where fa01 = '" & Mid(Text7, 1, 8) & "'", adoTaie, adOpenStatic, adLockReadOnly
'         If adocheck.RecordCount = 0 Then
'            MessageShow Label11
'            strControlButton = MsgText(602)
'            adocheck.Close
'            Text7.SetFocus
'            Exit Sub
'         End If
'         adocheck.Close
'      End If
'      If Text8 <> MsgText(601) Then
'         adocheck.CursorLocation = adUseClient
'         adocheck.Open "select st01 from staff where st01 = '" & Text8 & "'", adoTaie, adOpenStatic, adLockReadOnly
'         If adocheck.RecordCount = 0 Then
'            MessageShow Label12
'            strControlButton = MsgText(602)
'            adocheck.Close
'            Text8.SetFocus
'            Exit Sub
'         End If
'         adocheck.Close
'      End If
'      If Text10 <> MsgText(601) Then
'         adocheck.CursorLocation = adUseClient
''         If Len(Mid(Text10, 2, Len(Text10) - 1)) > 6 Then
'         'Ken 92/06/12 改用案件基本資料檢查
'         'adocheck.Open "select cp09 from caseprogress where cp01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and cp02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and cp03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and cp04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
'         adocheck.Open "select pa01 as SystemNo from patent where pa01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and pa02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and pa03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and pa04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
'                       "select tm01 as SystemNo from trademark where tm01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and tm02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and tm03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and tm04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
'                       "select lc01 as SystemNo from lawcase where lc01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and lc02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and lc03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and lc04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
'                       "select hc01 as SystemNo from hirecase where hc01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and hc02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and hc03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and hc04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
'                       "select sp01 as SystemNo from servicepractice where sp01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and sp02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and sp03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and sp04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
'         If adocheck.RecordCount = 0 Then
'            MessageShow Label14
'            strControlButton = MsgText(602)
'            adocheck.Close
'            Text10.SetFocus
'            Exit Sub
'         End If
'         adocheck.Close
''         Else
''            MessageShow Label14
''            strControlButton = MsgText(602)
''            Text10.SetFocus
''            Exit Sub
''         End If
'      End If
'
'      'add by sonia 2015/12/30 檢查民國105年起法務收入科目不可使用
'      intI = PUB_AccNoEnable(Text14, Val(FCDate(MaskEdBox1.Text)))
'      If intI <> 0 Then
'         strControlButton = MsgText(602)
'         Text14.SetFocus
'         Exit Sub
'      End If
'      'end 2015/12/30
'      'Add by Amy 2022/09/21 科目是2201開頭且本所案號是S開頭案號時,抓服務業務案件基本檔之申請國家為台灣000者, 科目必須為220103,非台灣者必須為220105
'      If Left(Text14, 4) = "2201" And Trim(Text10) <> MsgText(601) Then
'        If Mid(Text10, 1, Len(Text10) - 9) = "S" And ChkSCaseAccNO(Text14, Text10) = False Then
'           strControlButton = MsgText(602)
'           Text14.SetFocus
'           Exit Sub
'        End If
'      End If
'      'end 2022/09/21
'      'Add by Morgan 2007/10/2 檢查科目部門&智權人員是否正確
'      intI = PUB_AccNoGood(Text14, Text16, Text8)
'      If intI <> 0 Then
'         strControlButton = MsgText(602)
'         If intI = 1 Then
'            Text14.SetFocus
'         ElseIf intI = 2 Then
'            Text16.SetFocus
'         ElseIf intI = 3 Then
'            Text8.SetFocus
'         End If
'         Exit Sub
'      End If
'      'end 2007/10/2
'
'      If Text9 <> MsgText(601) Then
'         adocheck.CursorLocation = adUseClient
'         adocheck.Open "select a0201, a0202 from acc020 where a0201 = '" & Text1 & "' and a0202 = '" & Text9 & "'", adoTaie, adOpenStatic, adLockReadOnly
'         If adocheck.RecordCount = 0 Then
'            MessageShow Label13
'            strControlButton = MsgText(602)
'            adocheck.Close
'            Text9.SetFocus
'            Exit Sub
'         End If
'         adocheck.Close
'      End If
'   End If
'   adoacc021.Close
   If strSaveConfirm = MsgText(4) And Text3 = MsgText(601) Then
        Text3 = GetSeqNo(Text1, Text2) '解按修改直接insert的錯誤
   End If
   StrSQLa = "select * from acc021 where ax201 = '" & Text1 & "' and ax202 = '" & Text2 & "' and ax203 = '" & Text3 & "'"
   If adoacc021.State = adStateOpen Then adoacc021.Close 'Add by Amy 2024/08/05
   adoacc021.CursorLocation = adUseClient
   adoacc021.Open StrSQLa, adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modify by Amy 2014/07/10 有造字無法發現，所以改寫法
   Text10 = CaseNoZero(Text10)
   strCombo1 = Combo1 'Memo by Amy 2022/05/13 換行已於PUB_ChkUniText換掉
   If adoacc021.RecordCount = 0 Then
      'adoacc021.AddNew
      StrSQLa = "Insert Into acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax208,ax209,ax211,ax212,ax213,ax214) Values (" & _
                  "'" & Text1 & "' ,'" & Text2 & "' ,'" & Text3 & "' ,'" & IIf(Text16 <> MsgText(601), Text16, MsgText(55)) & "' ," & CNULL(ChgSQL(Text14)) & "," & _
                  IIf(Text4 <> MsgText(601), Val(Text4), 0) & "," & IIf(Text5 <> MsgText(601), Val(Text5), 0) & "," & CNULL(ChgSQL(Text7)) & "," & CNULL(ChgSQL(Text8)) & "," & _
                  CNULL(ChgSQL(Text9)) & "," & CNULL(ChgSQL(strCombo1)) & "," & CNULL(ChgSQL(Text6)) & "," & CNULL(ChgSQL(Text10)) & ")"

   Else
    StrSQLa = "Update acc021 set ax204='" & IIf(Text16 <> MsgText(601), Text16, MsgText(55)) & "' ,ax205=" & CNULL(ChgSQL(Text14)) & ",ax206=" & IIf(Text4 <> MsgText(601), Val(Text4), 0) & _
                 ",ax207=" & IIf(Text5 <> MsgText(601), Val(Text5), 0) & ",ax208=" & CNULL(ChgSQL(Text7)) & ",ax209=" & CNULL(ChgSQL(Text8)) & _
                 ",ax211=" & CNULL(ChgSQL(Text9)) & ",ax212=" & CNULL(ChgSQL(strCombo1)) & ",ax213=" & CNULL(ChgSQL(Text6)) & ",ax214=" & CNULL(ChgSQL(Text10)) & _
               " Where ax201= '" & Text1 & "' And ax202='" & Text2 & "'  And ax203='" & Text3 & "' "
   End If
   adoTaie.Execute StrSQLa
'   adoacc021.Fields("ax201").Value = Text1
'   adoacc021.Fields("ax202").Value = Text2
'   adoacc021.Fields("ax203").Value = Text3
'   If Text14 <> MsgText(601) Then
'      adoacc021.Fields("ax205").Value = Text14
'   Else
'      adoacc021.Fields("ax205").Value = Null
'   End If
'   If Text4 <> MsgText(601) Then
'      adoacc021.Fields("ax206").Value = Val(Text4)
'   Else
'      adoacc021.Fields("ax206").Value = 0
'   End If
'   If Text5 <> MsgText(601) Then
'      adoacc021.Fields("ax207").Value = Val(Text5)
'   Else
'      adoacc021.Fields("ax207").Value = 0
'   End If
'   If Text16 <> MsgText(601) Then
'      adoacc021.Fields("ax204").Value = Text16
'   Else
'      adoacc021.Fields("ax204").Value = MsgText(55)
'   End If
   'Modify by Amy 2022/05/13 原判斷Combo1 <> MsgText(601)
   If strCombo1 <> MsgText(601) Then
      'adoacc021.Fields("ax212").Value = Replace(Combo1, "'", "''")
      'strCombo1 = Combo1
      Combo1.Clear
      Combo1.AddItem strCombo1
'   Else
'      adoacc021.Fields("ax212").Value = Null
   End If
   'end 2022/05/13
'   If Text7 <> MsgText(601) Then
'      adoacc021.Fields("ax208").Value = Text7
'   Else
'      adoacc021.Fields("ax208").Value = Null
'   End If
'   If Text8 <> MsgText(601) Then
'      adoacc021.Fields("ax209").Value = Text8
'   Else
'      adoacc021.Fields("ax209").Value = Null
'   End If
'   If Text9 <> MsgText(601) Then
'      adoacc021.Fields("ax211").Value = Text9
'   Else
'      adoacc021.Fields("ax211").Value = Null
'   End If
'   If Text6 <> MsgText(601) Then
'      adoacc021.Fields("ax213").Value = Text6
'   Else
'      adoacc021.Fields("ax213").Value = Null
'   End If
'   Text10 = CaseNoZero(Text10)
'   If Text10 <> MsgText(601) Then
'      adoacc021.Fields("ax214").Value = Text10
'   Else
'      adoacc021.Fields("ax214").Value = Null
'   End If
   'Modify by Amy 2014/01/15 取消作帳公司
'   If Text18 <> MsgText(601) Then
'      adoacc021.Fields("ax215").Value = Text18
'   Else
'      adoacc021.Fields("ax215").Value = Null
'   End If
   'end 2014/01/15
'   adoacc021.UpdateBatch
   'end 2014/07/10
   'Add by Amy 2021/06/02 判斷結餘資料需發mail
   'Modify by Amy 2021/09/23 原程式改寫至ChkSetAxb16
   Call ChkSetAxb16("Ins")
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
    Dim strCaseNo(1 To 4)  As String 'Add by Amy 2014/11/19
    
   If IsNull(Adodc1.Recordset.Fields("ax203").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = Adodc1.Recordset.Fields("ax203").Value
   End If
   '會計科目
   If IsNull(Adodc1.Recordset.Fields("ax205").Value) Then
      Text14 = MsgText(601)
   Else
      Text14 = Adodc1.Recordset.Fields("ax205").Value
   End If
   Text14.Tag = Text14 'Add by Amy 2021/06/02
   '借方
   If IsNull(Adodc1.Recordset.Fields("ax206").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = Adodc1.Recordset.Fields("ax206").Value
   End If
   Text4.Tag = Text4 'Add by Amy 2021/06/02
   '貸方
   If IsNull(Adodc1.Recordset.Fields("ax207").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = Adodc1.Recordset.Fields("ax207").Value
   End If
   Text5.Tag = Text5 'Add by Amy 2021/06/02
   If IsNull(Adodc1.Recordset.Fields("ax204").Value) Then
      Text16 = MsgText(601)
      Text17 = MsgText(601)
   Else
      'Modify by Amy 2023/06/14 +Mid(Adodc1.Recordset.Fields("ax205").Value, 1, 2) <> "49",490101-安全基金撥補部門一定為TOT
      'Modify by Amy 2023/07/05 490102 部門一定[不]可以是TOT
      If Adodc1.Recordset.Fields("ax204").Value = MsgText(55) And Adodc1.Recordset.Fields("ax205") <> "490101" Then
            Text16 = MsgText(601)
            Text17 = MsgText(601)
      Else
         Text16 = Adodc1.Recordset.Fields("ax204").Value
      End If
   End If
   '摘要
   If IsNull(Adodc1.Recordset.Fields("ax212").Value) Then
      Combo1 = MsgText(601)
   Else
      Combo1 = Adodc1.Recordset.Fields("ax212").Value
   End If
   '對沖-客
   If IsNull(Adodc1.Recordset.Fields("ax208").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = Adodc1.Recordset.Fields("ax208").Value
   End If
   '對沖-業務
   If IsNull(Adodc1.Recordset.Fields("ax209").Value) Then
      Text8 = MsgText(601)
      Text19 = ""
   Else
      Text8 = Adodc1.Recordset.Fields("ax209").Value
      Text19 = StaffQuery(Text8)
   End If
   Text8.Tag = Text8 'Add by Amy 2021/06/02
   If IsNull(Adodc1.Recordset.Fields("ax211").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = Adodc1.Recordset.Fields("ax211").Value
   End If
   '對沖-其他
   If IsNull(Adodc1.Recordset.Fields("ax213").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = Adodc1.Recordset.Fields("ax213").Value
   End If
   Text6.Tag = Text6 'Add by Amy 2021/06/02
   If IsNull(Adodc1.Recordset.Fields("ax214").Value) Then
      Text10 = MsgText(601)
   Else
      Text10 = Adodc1.Recordset.Fields("ax214").Value
   End If
   Text10.Tag = Text10 'Add by Amy 2024/08/05 ex:1公司 D113080118 修改時無法改摘要(因預帶)
   'Add byAmy 2014/11/19
    ChgCaseNo Text10, strCaseNo()
    'Modify by Amy 2024/04/03 顧問固定帶L公司,法務只有ACS案可能為J公司其他為1公司
    'If ChkPatentNameCompany(strCaseNo(1), strCaseNo(2), strCaseNo(3), strCaseNo(4)) = "J" Then
    'Modify by Amy 2024/08/05 原:ChkPatentNameCompany改其他地方也可用
    strExc(2) = ChkPatentNameCompany("1", strCaseNo(1), strCaseNo(2), strCaseNo(3), strCaseNo(4))
    If strExc(2) = "J" Or strExc(2) = "L" Then
        strSpecComp = strExc(2)
    'end 2024/04/02
    Else
        strSpecComp = "1"
    End If
   'Modify by Amy 2014/01/15 取消作帳公司
'   If IsNull(Adodc1.Recordset.Fields("ax215").Value) Then
'      Text18 = MsgText(601)
'   Else
'      Text18 = Adodc1.Recordset.Fields("ax215").Value
'   End If
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   If ExistCheck("acc080", "a0801", Text1, Label3, False) = False Then
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
   CloseIme
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Mark by Amy 2024/08/05 避免未檢查到,將程式改至 Validate
Private Sub Text10_LostFocus()
'   'add by nick 2004/07/01
'On Error GoTo Checking
'If Text10 <> MsgText(601) Then
'      Dim strNation As String
'      Text10 = CaseNoZero(Text10)
'      adocase.CursorLocation = adUseClient
'      adocase.Open "select pa01 as SystemNo,pa09,pa26  from patent where pa01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and pa02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and pa03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and pa04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
'                   "select tm01 as SystemNo,tm10,tm23 from trademark where tm01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and tm02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and tm03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and tm04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
'                   "select lc01 as SystemNo,lc15,lc11 from lawcase where lc01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and lc02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and lc03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and lc04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
'                   "select hc01 as SystemNo,'000',hc07 from hirecase where hc01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and hc02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and hc03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and hc04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
'                   "select sp01 as SystemNo,sp09,sp08 from servicepractice where sp01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and sp02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and sp03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and sp04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adocase.RecordCount > 0 Then
'         '檢查當科目是 220112 220102 220111 220101 220103 220104 220105 220106 時，要檢查申請國家級系統別
'         strNation = CheckStr(adocase.Fields(1).Value)
'         'add by sonia 2021/5/3 加傳本所案號以判斷英日文組FCP-058897
'         If AccNoToSalesNo(Text14, Text10) <> "" Then
'            Text8 = AccNoToSalesNo(Text14, Text10)
'         End If
'         'end 2021/5/3
'      End If
'      adocase.Close
'         Select Case Text14
'         Case "220101"
'                 'edit by nick 2004/07/07 加系統別
'                 'If (Mid(Text10, 1, Len(Text10) - 9) = "T" Or Mid(Text10, 1, Len(Text10) - 9) = "TB") And strNation = "000" Then
'                 If (Mid(Text10, 1, Len(Text10) - 9) = "T" Or Mid(Text10, 1, Len(Text10) - 9) = "TB" Or Mid(Text10, 1, Len(Text10) - 9) = "TS" Or Mid(Text10, 1, Len(Text10) - 9) = "TD" Or Mid(Text10, 1, Len(Text10) - 9) = "TM" Or Mid(Text10, 1, Len(Text10) - 9) = "TR" Or Mid(Text10, 1, Len(Text10) - 9) = "TT") And strNation = "000" Then
'                 Else
'                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
'                       Text14.SetFocus
'                       Text14.SelStart = 0
'                       Text14.SelLength = Len(Text14)
'                       Exit Sub
'                 End If
'         Case "220102"
'                 'edit by nick 2004/07/07 加系統別
'                 'If Mid(Text10, 1, Len(Text10) - 9) = "P" And strNation = "000" Then
'                 If (Mid(Text10, 1, Len(Text10) - 9) = "P" Or Mid(Text10, 1, Len(Text10) - 9) = "PS") And strNation = "000" Then
'                 Else
'                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
'                       Text14.SetFocus
'                       Text14.SelStart = 0
'                       Text14.SelLength = Len(Text14)
'                       Exit Sub
'                 End If
'         Case "220103"
'                 'edit by nick 2004/07/07 加系統別
'                 'If Mid(Text10, 1, Len(Text10) - 9) = "FCT" Then
'                 If (Mid(Text10, 1, Len(Text10) - 9) = "FCT" Or Mid(Text10, 1, Len(Text10) - 9) = "S") And strNation = "000" Then
'                 Else
'                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
'                       Text14.SetFocus
'                       Text14.SelStart = 0
'                       Text14.SelLength = Len(Text14)
'                       Exit Sub
'                 End If
'         Case "220104"
'                 'edit by nick 2004/07/07 加系統別
'                 'If Mid(Text10, 1, Len(Text10) - 9) = "FCP" Then
'                 If Mid(Text10, 1, Len(Text10) - 9) = "FCP" Or Mid(Text10, 1, Len(Text10) - 9) = "FG" Then
'                 Else
'                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
'                       Text14.SetFocus
'                       Text14.SelStart = 0
'                       Text14.SelLength = Len(Text14)
'                       Exit Sub
'                 End If
'         Case "220105"
'                 'edit by nick 2004/07/07 加系統別
'                 'If Mid(Text10, 1, Len(Text10) - 9) = "CFT" Then
'                 If (Mid(Text10, 1, Len(Text10) - 9) = "CFT" Or Mid(Text10, 1, Len(Text10) - 9) = "CFC" Or Mid(Text10, 1, Len(Text10) - 9) = "S") And strNation <> "000" Then
'                 Else
'                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
'                       Text14.SetFocus
'                       Text14.SelStart = 0
'                       Text14.SelLength = Len(Text14)
'                       Exit Sub
'                 End If
'         Case "220106"
'                 'edit by nick 2004/07/07 加系統別
'                 'If Mid(Text10, 1, Len(Text10) - 9) = "CFP" Then
'                 '2012/4/24 MODIFY BY SONIA 加系統類別 LIN
'                 '2014/5/28 modify by sonia 加系統類別 L 但必須為非台灣
'                 If Mid(Text10, 1, Len(Text10) - 9) = "CFP" Or Mid(Text10, 1, Len(Text10) - 9) = "FCL" Or Mid(Text10, 1, Len(Text10) - 9) = "LIN" Or Mid(Text10, 1, Len(Text10) - 9) = "CFL" Or Mid(Text10, 1, Len(Text10) - 9) = "CPS" Or Mid(Text10, 1, Len(Text10) - 9) = "L" Then
'                    '2014/5/28 ADD BY SONIA L案必須為非台灣
'                    If Mid(Text10, 1, Len(Text10) - 9) = "L" Then
'                        If strNation = "000" Then
'                           MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
'                           Text14.SetFocus
'                           Text14.SelStart = 0
'                           Text14.SelLength = Len(Text14)
'                           Exit Sub
'                        End If
'                    End If
'                    'END 2014/5/28
'                 Else
'                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
'                       Text14.SetFocus
'                       Text14.SelStart = 0
'                       Text14.SelLength = Len(Text14)
'                       Exit Sub
'                 End If
'         Case "220107"
'                 'add by nick 2004/07/07 加系統別
'                 '2012/4/24 MODIFY BY SONIA 加申請國家條件
'                 If Mid(Text10, 1, Len(Text10) - 9) = "TC" And strNation = "000" Then
'                 Else
'                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
'                       Text14.SetFocus
'                       Text14.SelStart = 0
'                       Text14.SelLength = Len(Text14)
'                       Exit Sub
'                 End If
'         Case "220111"
'                 'edit by nick 2004/07/07 加系統別
'                 'If (Mid(Text10, 1, Len(Text10) - 9) = "T" Or Mid(Text10, 1, Len(Text10) - 8) = "TF") And strNation <> "000" Then
'                 '2012/4/24 modify by sonia 加系統類別TC,TD,TM,TT,TB,TR
'                 If (Mid(Text10, 1, Len(Text10) - 9) = "TS" Or Mid(Text10, 1, Len(Text10) - 9) = "T" Or Mid(Text10, 1, Len(Text10) - 9) = "TF" Or Mid(Text10, 1, Len(Text10) - 9) = "TC" Or Mid(Text10, 1, Len(Text10) - 9) = "TD" Or Mid(Text10, 1, Len(Text10) - 9) = "TM" Or Mid(Text10, 1, Len(Text10) - 9) = "TT" Or Mid(Text10, 1, Len(Text10) - 9) = "TB" Or Mid(Text10, 1, Len(Text10) - 9) = "TR") And strNation <> "000" Then
'                 Else
'                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
'                       Text14.SetFocus
'                       Text14.SelStart = 0
'                       Text14.SelLength = Len(Text14)
'                       Exit Sub
'                 End If
'         Case "220112"
'                 'edit by nick 2004/07/07 加系統別
'                 'If Mid(Text10, 1, Len(Text10) - 9) = "P" And strNation <> "000" Then
'                 If (Mid(Text10, 1, Len(Text10) - 9) = "P" Or Mid(Text10, 1, Len(Text10) - 9) = "PS") And strNation <> "000" Then
'                 Else
'                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
'                       Text14.SetFocus
'                       Text14.SelStart = 0
'                       Text14.SelLength = Len(Text14)
'                       Exit Sub
'                 End If
'         '2012/4/24 ADD BY SONIA
'         Case "220108"
'                 If (Mid(Text10, 1, Len(Text10) - 9) = "P" Or Mid(Text10, 1, Len(Text10) - 9) = "PS") Then
'                 Else
'                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
'                       Text14.SetFocus
'                       Text14.SelStart = 0
'                       Text14.SelLength = Len(Text14)
'                       Exit Sub
'                 End If
'         Case "220113"
'                 If Mid(Text10, 1, Len(Text10) - 9) = "L" Or Mid(Text10, 1, Len(Text10) - 9) = "LA" Or Mid(Text10, 1, Len(Text10) - 9) = "FCL" Or Mid(Text10, 1, Len(Text10) - 9) = "LIN" Then
'                    '2014/5/28 ADD BY SONIA L案必須為台灣
'                    If Mid(Text10, 1, Len(Text10) - 9) = "L" Then
'                        If strNation <> "000" Then
'                           MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
'                           Text14.SetFocus
'                           Text14.SelStart = 0
'                           Text14.SelLength = Len(Text14)
'                           Exit Sub
'                        End If
'                    End If
'                    'END 2014/5/28
'                 Else
'                    'add by sonia 2024/4/3 因財務處2024/3/26要求：法律所案案源案件之專業部門所提列(T, P, FCT, FCP)的出庭費, 想統一科目改成 220113
'                    If strNation = "000" And (Mid(Text10, 1, Len(Text10) - 9) = "P" Or Mid(Text10, 1, Len(Text10) - 9) = "T" Or Mid(Text10, 1, Len(Text10) - 9) = "FCP" Or Mid(Text10, 1, Len(Text10) - 9) = "FCT") Then
'                    Else
'                    'end 2024/4/3
'                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
'                       Text14.SetFocus
'                       Text14.SelStart = 0
'                       Text14.SelLength = Len(Text14)
'                       Exit Sub
'                    End If    'add by sonia 2024/4/3
'                 End If
'         '2012/4/24 END
'         Case "610103"
'                 'add by nick 2004/07/07 加系統別
'                 '2012/4/24 MODIFY BY SONIA 加系統類別LIN
'                 If Mid(Text10, 1, Len(Text10) - 9) = "L" Or Mid(Text10, 1, Len(Text10) - 9) = "LA" Or Mid(Text10, 1, Len(Text10) - 9) = "FCL" Or Mid(Text10, 1, Len(Text10) - 9) = "LIN" Or Mid(Text10, 1, Len(Text10) - 9) = "CFL" Then
'                 Else
'                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
'                       Text14.SetFocus
'                       Text14.SelStart = 0
'                       Text14.SelLength = Len(Text14)
'                       Exit Sub
'                 End If
'        Case Else
'         End Select
' End If
' Exit Sub
'Checking:
'   MsgBox MsgText(128), , MsgText(5)
'   Exit Sub
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   
On Error GoTo Checking
   strNation = "" 'Add by Amy 2024/08/05
   
   If Text10 <> MsgText(601) Then
      Text10 = CaseNoZero(Text10)
      'Modify by Amy 2024/08/05 避免有未改到,將資料改至ChkPatentNameCompany
       strExc(3) = ChkPatentNameCompany(2, Mid(Text10, 1, Len(Text10) - 9), Mid(Text10, Len(Text10) - 8, 6), Mid(Text10, Len(Text10) - 2, 1), Mid(Text10, Len(Text10) - 1, 2))
      If adocase.State = adStateOpen Then
         adocase.Close
      End If
      adocase.CursorLocation = adUseClient
      'Ken 92/06/12 改用案件基本資料檢查
      'adocase.Open "select cp09 from caseprogress where cp01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and cp02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and cp03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and cp04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      'edit by nick 2004/07/01
      'adocase.Open "select pa01 as SystemNo  from patent where pa01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and pa02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and pa03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and pa04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
                   "select tm01 as SystemNo from trademark where tm01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and tm02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and tm03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and tm04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
                   "select lc01 as SystemNo from lawcase where lc01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and lc02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and lc03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and lc04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
                   "select hc01 as SystemNo from hirecase where hc01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and hc02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and hc03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and hc04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
                   "select sp01 as SystemNo from servicepractice where sp01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and sp02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and sp03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and sp04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      'Modify by Amy 2014/11/17 +特殊出名公司
      'Modify by Amy 2024/04/03 顧問'1' as SpecComp->'L' as SpecComp,lc48->Decode(Instr(lc01,'L'),0,lc48,'L')
      adocase.Open strExc(3), adoTaie, adOpenStatic, adLockReadOnly
      'end 2024/08/05
      If adocase.RecordCount = 0 Then
         MsgBox MsgText(28) & Label14, , MsgText(5)
         Cancel = True
         adocase.Close
         Exit Sub
      Else
        'Modify by Amy 2024/08/05 從LostFocus搬過來
        strNation = "" & adocase.Fields("Nation")
        'add by sonia 2021/5/3 加傳本所案號以判斷英日文組FCP-058897
        'Modify by Amy 2024/08/05 有修改案號才預帶
        If Text10.Tag <> Text10 Then
            If AccNoToSalesNo(Text14, Text10) <> "" Then
               Text8 = AccNoToSalesNo(Text14, Text10)
            End If
         End If
         'end 2021/5/3
         'end 2024/08/05
         
        'Add by Amy 2014/11/17 +記錄特殊出名公司
        'Modify by Amy 2024/04/03 出名稱公司J及L,其餘設1
        If "" & adocase.Fields("SpecComp") = "J" Or "" & adocase.Fields("SpecComp") = "L" Then
            'strSpecComp = "J"
            strSpecComp = "" & adocase.Fields("SpecComp")
        'end 2024/04/03
        Else
            strSpecComp = "1"
        End If
        'end 2014/11/17
        'edit by nick 2004/07/13 財務處說有些不秀公司別
         QueryCustomer 1
               '2004/07/01 nick
         '針對   P  T  TF  CFT  CFP  加入客戶名稱
         '  FCT  FCP 加入本所案號
         '這些案號的摘要 , 要在前面加入資訊
         Select Case Mid(Text10, 1, Len(Text10) - 9)
         'edit by nick 2004/07/06 總帳原本就有抓案件名稱，所以不用在做一次
'         Case "P", "T", "TF", "CFT", "CFP"
'                Dim strCustomer As String
'                strCustomer = CheckStr(adocase.Fields(2).Value)
'                AdoRecordSet3.CursorLocation = adUseClient
'                AdoRecordSet3.Open "SELECT cu04 FROM Customer " & _
'               "WHERE CU01 = '" & Mid(strCustomer, 1, 8) & "' AND " & _
'                     "CU02 = '" & Mid(strCustomer, 9, 1) & "'", cnnConnection, adOpenStatic, adLockReadOnly
'                If AdoRecordSet3.RecordCount > 0 Then
'                       Combo1 = CheckStr(AdoRecordSet3.Fields(0).Value) & "/" & Combo1
'                End If
'                AdoRecordSet3.Close
         Case "FCT", "FCP"
            'Modify by Amy 2024/08/05 有修改案號才預帶
            If Text10.Tag <> Text10 Then
                 Combo1 = Text10 & "/" & Combo1
            End If
         Case Else
            'add by nick 2004/07/13
            QueryCustomer 2
         End Select
         
      End If
      adocase.Close
      Text10.Tag = Text10 'Add by Amy 2024/08/05 預帶後立即更新Tag,避免其他資料被不能修改
      'Memo by Amy 2024/08/05 原:Text10_LostFocus 以案號檢查會計科目 程式改至ChkForm
      '                                                    寫於此檢查為Cancel則會無法跳至會計科目欄
      
   End If
   Exit Sub
   
Checking:
   MsgBox MsgText(128), , MsgText(5)
   Exit Sub
End Sub

Private Sub Text14_GotFocus()
   TextInverse Text14
End Sub

Private Sub Text14_Change()
   If Text14 = MsgText(601) Then
      Exit Sub
   End If
   Text15 = A0102Query(Text14)
End Sub

Private Sub Text14_Validate(Cancel As Boolean)
   If Text14 <> MsgText(601) Then
      'Modify by Amy 2014/01/15 改檢查會計科目的公司別是否正確
      'If ExistCheck("acc010", "A0101", Text14, Label5) = False Then
      If PUB_CheckCompany(Text14, Text1) = False Then
         Cancel = True
         Exit Sub
      End If
   End If
  
   If AccNoToSalesNo(Text14, Text10) <> "" Then
      'modify by sonia 2021/5/3 加傳本所案號以判斷英日文組FCP-058897
      Text8 = AccNoToSalesNo(Text14, Text10)
   End If
 
   '2005/6/6 ADD BY SONIA
   If Text14 = "1134" Or Text14 = "1305" Or Text14 = "1307" Or Text14 = "1308" Or Text14 = "1309" Or Text14 = "1310" Or Text14 = "110211" Or Text14 = "110212" Then
      MsgBox "此科目已不再使用!!", , "User 輸入錯誤!!"
      Cancel = True
      Text14.SelStart = 0
      Text14.SelLength = Len(Text14)
      Exit Sub
   End If
   '2005/6/6 END
   
End Sub

Private Sub Text16_Change()
   If Text16 = MsgText(601) Then
      Exit Sub
   End If
   Text17 = A0902Query(Text16)
End Sub

Private Sub Text16_GotFocus()
   TextInverse Text16
   CloseIme
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text16_Validate(Cancel As Boolean)
   Dim stMsg As String 'Add by Amy 2020/08/12
   
   '部門
   If Text16 <> MsgText(601) Then
      If ExistCheck("acc090", "a0901", Text16, Label8) = False Then
         Cancel = True
         Exit Sub
      End If
   End If
   'Modify by Amy 2021/05/17 +FormName
   'Modify by Amy 2023/06/14 +stMsg 部門只能是 TOT-婉莘
   '    ex: 1 公司 D112011982輸M部門,部門損益表會列示於管理部,改為一致
   If CheckDept(Text14, Text16, Me.Name, stMsg) = False Then
      If stMsg <> MsgText(601) Then
         MsgBox stMsg, , MsgText(5)
      Else
         MsgBox MsgText(103), , MsgText(5)
      End If
      Cancel = True
      Exit Sub
   End If
   'Add by Amy 2020/08/12 L公司6及72字頭部門別不可輸TOT或空值
   If Text1 = "L" And (Left(Text14, 1) = "6" Or Left(Text14, 2) = "72") And (Text16 = "TOT" Or Trim(Text16) = MsgText(601)) Then
      MsgBox "L公司6或72字頭會計科目之部門別不可輸TOT或空白", , MsgText(5)
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text18_GotFocus()
   CloseIme
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Modify by Amy 2014/01/15 取消作帳公司
'Private Sub Text18_GotFocus()
'   TextInverse Text18
'End Sub
'
'Private Sub Text18_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub
'end 2014/01/15

Private Sub Text2_GotFocus()
   'Modify by Morgan 2006/11/6 還原--瑞婷
   'Modify by Morgan 2006/7/20--瑞婷
   'TextInverse Text2
'   If Text2 = "" Then
'      Text2 = "D"
'      Text2.SelStart = 1
'   Else
'      Text2.SelStart = 1
'      Text2.SelLength = Len(Text2) - 1
'   End If
   TextInverse Text2
   CloseIme
   'End 2006/11/6
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  清除欄位資料(傳票資料--交易檔)
'
'*************************************************
Public Sub AdodcClear()
   Text3 = ""
   Text14 = ""
   Text14.Tag = "" 'Add by Amy 2021/06/02
   Text15 = ""
   Text4 = ""
   Text4.Tag = "" 'Add by Amy 2021/06/02
   Text5 = ""
   Text5.Tag = "" 'Add by Amy 2021/06/02
   Text16 = ""
   Text17 = ""
   Combo1 = ""
   Text7 = ""
   Text8 = ""
   Text8.Tag = "" 'Add by Amy 2021/06/02
   Text9 = ""
   Text6 = ""
   Text6.Tag = "" 'Add by Amy 2021/06/02
   Text10 = ""
   Text10.Tag = "" 'Add by Amy 2024/08/05
   'Text18 = "" Modify by Amy 2014/01/15
   Text19 = ""
   strControlButton = "" 'Add by Amy 2024/08/05
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   'Add byAmy 2014/11/17
   Dim strMsg As String
   Dim nResponse
   Dim bCancel As Boolean 'Add by Amy 2020/08/12
   
   Call PUB_SaveTrackMode(1, KeyCode) 'Add by Amy 2021/12/07 Form2.0
   
   Select Case KeyCode
      Case vbKeyInsert
         'Added by Lydia 2025/08/08 明細要按Insert才會更新資料，但是Form2.0元件支援Insert鍵會切換”新增/覆寫模式”
         If bolFirst = True Then  '在按下Insert鍵時先重送Insert鍵於第2次才執行更新明細
             Call PUB_SendSKey("KeyInsert")
             bolFirst = False
             Exit Sub
         End If
         'end 2025/08/08
         
         'Add by Amy 2024/08/05 檢查時不應可按Insert
         '    ex:查 D113070003->修改->傳票日改 7/4 ->彈訊息 ->點明細->Insert
         If strSaveConfirm = MsgText(601) Then
            Exit Sub
         End If
         'Modify by Amy 2024/08/05 原檢查程式改至ChkForm,避免有未改到的,抓項次編號改抓GetSeqNo
         If ChkForm("Ins") = False Then
            strControlButton = MsgText(602)
            Exit Sub
         End If
         SaveData ("Ins")
         If strControlButton <> MsgText(602) Then
            AdodcClear
            Text3 = GetSeqNo(Text1, Text2)
         'end 2024/08/05
            SumShow
            Text14.SetFocus
         End If
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
   'Add by Amy 2024/08/05
   If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
      m_CurrKEY = Text2
   End If
End Sub

'Modify by Amy 2014/01/15 +Mark  因進入按換公司別會彈訊息 改至command3
'Private Sub Text2_Validate(Cancel As Boolean)
'    If Text2 = MsgText(601) Then
'        MsgBox Label2 & MsgText(52), , MsgText(5)
'        Cancel = True
'        Exit Sub
'    End If
'End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
   OpenIme
End Sub
'Modify by Amy 2022/06/07 原:Integer
Private Sub Text6_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
   CloseIme
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
   If Text7 <> MsgText(601) Then
      If Len(Text7) = 6 Then
         Text7 = AfterZero(Text7)
      End If
      If ExistCheck("customer", "cu01", Mid(Text7, 1, 8), Label11, False) = False Then
         If ExistCheck("acc0i0", "a0i01", Text7, Label11, False) = False Then
            If ExistCheck("staff", "st01", Text7, Label11, False) = False Then
               If ExistCheck("fagent", "fa01", Mid(Text7, 1, 8), Label11, False) = False Then
                  MsgBox MsgText(28) & Label11, , MsgText(5)
                  Cancel = True
                  Exit Sub
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
   CloseIme
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text8_Validate(Cancel As Boolean)
   Text19 = ""
   If Text8 <> MsgText(601) Then
'Modify by Morgan 2007/2/5 員工已離職要提醒
'      If ExistCheck("staff", "st01", Text8, Label12) = False Then
'         Cancel = True
'         Exit Sub
'      End If
      'add by sonia 2016/11/15 不可輸入S29,否則實績結餘點數報表抓不到
      'Modify by Amy 2019/08/30 W1001/W2001要可以輸
      'modify by sonia 2023/12/26 W1001/W2001改寫法W開頭
      'If Text8 > "S" And Text8 <> "W1001" And Text8 <> "W2001" Then
      If Text8 > "S" And Left(Text8, 1) <> "W" Then
         MsgBox "不可輸入S字頭的編號,否則實績結餘點數報表抓不到 !", , MsgText(5)
         Cancel = True
         TextInverse Text8
      'end 2016/11/15
      ElseIf PUB_GetStaffState(Text8.Text, strExc(1), True) = 0 Then
         Cancel = True
         TextInverse Text8
      Else
         Text19.Text = strExc(1)
      End If
      'add by sonia 2023/11/30
      If SalesNoCheckAccNo(Text14, Text8) = False Then
      End If
      'end 2023/11/30
'end 2007/2/5
   End If
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
   CloseIme
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
   adoaccsum.Open "select sum(ax206), sum(ax207) from acc021 where ax201 = '" & Text1 & "' and ax202 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
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
   CmdChgComp.Enabled = True 'Add by Amy 2014/01/15 +換公司鈕
   'Modify by Amy 2014/11/20
   'MaskEdBox1.Enabled = False
   Text2.Enabled = True
   Text2.Locked = False
   'end 2014/11/20
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
   'Text18.Enabled = False Modify by Amy 2014/01/15
   Command1.Enabled = False
   Command2.Enabled = False
End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'
'*************************************************
Public Sub FormEnabled()
   CmdChgComp.Enabled = False 'Add by Amy 2014/01/15
   'Modify by Amy 2014/11/20 傳票日可修改
   MaskEdBox1.Enabled = True
   Text2.Locked = True
   'end 2014/11/20
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
   'Text18.Enabled = True Modify by Amy 2014/01/15
   Command1.Enabled = True
   Command2.Enabled = True
End Sub

'*************************************************
'  重新整理傳票資料
'
'*************************************************
'Modify by Amy 2024/08/05 +m_A0202
Public Sub Acc020Refresh(Optional ByVal m_A0202 As String = "")
   Dim stA0202 As String 'Add by Amy 2024/08/05
On Error GoTo Checking
   'Modify by Amy 2024/08/05
   stA0202 = Text2
   If m_A0202 <> MsgText(601) Then
      stA0202 = m_A0202
   End If
   'end 2024/08/05
   If adoacc020.State = adStateOpen Then
      adoacc020.Close
   End If
   adoacc020.CursorLocation = adUseClient
   adoacc020.MaxRecords = intMax
   'Modify by Amy 2014/01/15 +公司別
   'Modify by Amy 2024/08/05 原:And a0202 >= '" & Text2 & "'
   strExc(0) = "select * from acc020 where a0201='" & Text1 & "' And a0202 >= '" & stA0202 & "' order by a0201 asc, a0202 asc"
   adoacc020.Open strExc(0), adoTaie, adOpenDynamic, adLockBatchOptimistic
Checking:
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
   adoacc021.Open "select * from acc021 where ax202 = '" & Text2 & "' and ax201 = '" & Text1 & "' and ax203 = '" & Text3 & "' order by ax201 asc, ax202 asc, ax203 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc021, acc010 where ax205 = a0101 (+) and ax202 = '" & Text2 & "' and ax201 = '" & Text1 & "' order by ax201 asc, ax202 asc, ax203 asc", adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.Find "ax203 = '" & Text3 & "'", 0, adSearchForward, 1
      If Adodc1.Recordset.EOF Then
         Adodc1.Recordset.MoveFirst
         DataGrid1.SelBookmarks.add Adodc1.Recordset.Bookmark
         Exit Sub
      Else
         DataGrid1.SelBookmarks.add Adodc1.Recordset.Bookmark
      End If
   End If
   AdodcClear
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

Private Sub Text9_Validate(Cancel As Boolean)
   If Text9 <> MsgText(601) Then
      adocheck.CursorLocation = adUseClient
      'Modify by Amy 2014/01/15  改公司別 原a0201 = '1'
      adocheck.Open "select a0201 from acc020 where a0201 = '" & Text1 & "' and a0202 = '" & Text9 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount = 0 Then
         MessageShow Label13
         adocheck.Close
         Cancel = True
         Exit Sub
      End If
      adocheck.Close
   End If
End Sub

'*************************************************
'  以本所案號查詢客戶名稱
'
'*************************************************
Public Sub QueryCustomer(Optional State As Integer = 1)
Dim strSql As String

   If Text10 = MsgText(601) Then
      Exit Sub
   End If
   'Add by Amy 2024/08/05 +案號不同才預帶 ex:1公司 D113080118 修改時無法改摘要(因預帶)
   If Text10.Tag = Text10 Then Exit Sub
 
   strSql = "select cu01||cu02 as Name, cu04, cu05, cu06, cu88, cu89, cu90 from patent, customer where substr(pa26, 1, 8) = cu01 and nvl(substr(pa26, 9, 1), '0') = cu02 and pa01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and pa02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and pa03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and pa04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
            "select cu01||cu02 as Name, cu04, cu05, cu06, cu88, cu89, cu90 from trademark, customer where substr(tm23, 1, 8) = cu01 and nvl(substr(tm23, 9, 1), '0') = cu02 and tm01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and tm02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and tm03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and tm04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
            "select cu01||cu02 as Name, cu04, cu05, cu06, cu88, cu89, cu90 from lawcase, customer where substr(lc11, 1, 8) = cu01 and nvl(substr(lc11, 9, 1), '0') = cu02 and lc01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and lc02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and lc03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and lc04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
            "select cu01||cu02 as Name, cu04, cu05, cu06, cu88, cu89, cu90 from hirecase, customer where substr(hc05, 1, 8) = cu01 and nvl(substr(hc05, 9, 1), '0') = cu02 and hc01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and hc02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and hc03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and hc04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
            "select cu01||cu02 as Name, cu04, cu05, cu06, cu88, cu89, cu90 from servicepractice, customer where substr(sp08, 1, 8) = cu01 and nvl(substr(sp08, 9, 1), '0') = cu02 and sp01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and sp02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and sp03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and sp04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "'"
   Set adocase1 = New ADODB.Recordset
   adocase1.CursorLocation = adUseClient
   adocase1.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adocase1.RecordCount <> 0 Then
      'add by nick 2004/07/13
      If State = 1 Then
         If IsNull(adocase1.Fields(0).Value) Then
            Text7 = MsgText(601)
         Else
            Text7 = adocase1.Fields(0).Value
         End If
      End If
      'add by nick 2004/07/13
      If State = 2 Then
         If IsNull(adocase1.Fields("cu04").Value) Then
            If IsNull(adocase1.Fields("cu05").Value) Then
               If IsNull(adocase1.Fields("cu06").Value) Then
                  Combo1 = MsgText(601)
               Else
                  Combo1 = adocase1.Fields("cu06").Value
               End If
            Else
               Combo1 = adocase1.Fields("cu05").Value
               If IsNull(adocase1.Fields("cu88").Value) = False Then
                  Combo1 = Combo1 & adocase1.Fields("cu88").Value
               End If
               If IsNull(adocase1.Fields("cu89").Value) = False Then
                  Combo1 = Combo1 & adocase1.Fields("cu89").Value
               End If
               If IsNull(adocase1.Fields("cu90").Value) = False Then
                  Combo1 = Combo1 & adocase1.Fields("cu90").Value
               End If
            End If
         Else
            Combo1 = adocase1.Fields("cu04").Value
         End If
      End If
   Else
      Text7 = MsgText(601)
      Combo1 = MsgText(601)
   End If
  
   adocase1.Close
End Sub

'Mark by Amy 2024/08/05 不使用,檢查改至ChkForm,存檔改至SaveData
'Add by Amy 2014/01/06 搬回.bas Function
Public Sub FormCheck() 'F9 鈕
'   Dim bCancel As Boolean 'Add by Amy 2023/06/14
'
'    '合計確認
'    If CreDebCheck <> MsgText(602) Or Val(Text11) = 0 Or Val(Text12) = 0 Then
'       MsgBox MsgText(11), , MsgText(5)
'       strControlButton = MsgText(602)
'       Exit Sub
'    End If
'    'Add by Amy 2023/06/14 若部門欄位未跳離開,可能會沒檢查到,故存檔前再檢查一次
'    Call Text16_Validate(bCancel)
'    If bCancel = True Then
'         strControlButton = MsgText(602)
'         TextInverse Text16
'         Text16.SetFocus
'         Exit Sub
'    End If
'    'Add by Amy 2018/03/27 檢查同一張傳票若借貸都有2491XX的科目時,[不能]有4XXX的貸方
'    'Memo 若未剔除4字頭,會[智權人員結餘點數查詢](frmacc4270) 1放出 資料可能抓的不正確 ex:1 公司 D106092547 需剔除
'    If Chk2491InCome = True Then
'        MsgBox "結餘放收入及結餘調整人員不可做在同一傳票！", , MsgText(5)
'        strControlButton = MsgText(602)
'        Exit Sub
'    End If
'    'Add by Amy 2023/06/14 避免從傳票輸入key 隱藏版傳票資料輸錯,導致期末結餘資料有誤,故彈訊息提醒ex:1公司D112052529
'    'Memo                                   因結餘保留放出產生傳票(frmacc41f0) 1放出,也會產生同樣借貸科目,且財務也可能做調整傳票(借:24910x 貸:4字頭 1120614 問瑞婷),故不可鎖住
'    If Chk24910xAnd41xxxx = True Then
'         If MsgBox("此為「非當月結餘轉撥傳票產生(隱藏版人員)」所用之傳票科目" & vbCrLf & _
'                              "確定[不是]要在隱藏版輸,按「是」繼續存檔,按「否」不存回前畫面", vbYesNo + vbDefaultButton2) = vbNo Then
'
'               strControlButton = MsgText(602)
'               Exit Sub
'         End If
'    End If
'    'Add by Amy 2022/05/16 結餘轉撥傳票存檔時需檢查每個人的總額必須與SalesPoint相符
'    If UCase(Me.Tag) = "FRMACC41H0" Then
'        If ChkSalesPointVal = True Then
'            strControlButton = MsgText(602)
'            Exit Sub
'        End If
'    End If
'    '2010/6/22 ADD BY SONIA 新增時若傳票日已月結則提醒
'    Dim m_A0B02 As String
'    If strSaveConfirm = MsgText(3) And MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'       adocheck.CursorLocation = adUseClient
'       '2014/1/22 modify by sonia 加公司別a0b04
'       'adocheck.Open "select A0B02 from acc0B0", adoTaie, adOpenStatic, adLockReadOnly
'       adocheck.Open "select a0b02 from acc0b0 where a0b04='" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
'       If adocheck.RecordCount <> 0 Then
'          If IsNull(adocheck.Fields("A0B02").Value) = False Then
'             m_A0B02 = adocheck.Fields("A0B02").Value
'          End If
'       End If
'       adocheck.Close
'       If Val(FCDate(MaskEdBox1.Text)) <= Val(m_A0B02) Then
'          If MsgBox("此日期已月結, 確定要輸此日期的傳票嗎 ??", vbYesNo + vbDefaultButton2) = vbYes Then
'          Else
'             strControlButton = MsgText(602)
'             MaskEdBox1.SetFocus
'             Exit Sub
'          End If
'       End If
'    End If
'    '2010/6/22 END
'    'Add by Amy 2014/07/16 J公司傳票借方為進項稅額時,判斷A1P04及進項發票無資料時不可存檔
'    If Text1 = "J" Then
'        If CheckIs1211(Text1, Text2) = True Then
'            If CheckExistA1p04(Text1, Text2) = False Then
'                If ExistCheck("Acc450", "A4501", Text2, "", False) = False Then
'                    MsgBox "有進項稅額科目, 請輸入發票明細 !", , MsgText(5)
'                    strControlButton = MsgText(602)
'                    Exit Sub
'                End If
'            End If
'        End If
'    End If
'    'end 2014/07/16
'    If strSaveConfirm = MsgText(4) Then
'       Frmacc4120_Save
'    End If
'    'Add by Amy 2014/11/20 因frmacc4120_Save 跳離會鎖住要key的欄位
'    If strControlButton = MsgText(602) Then Exit Sub
'    'Add by Amy 2014/11/24 解新增後直接修改(貸方金額改借方金額)及傳票日.tag未更新
'    AdodcRefresh
'    MaskEdBox1.Tag = Val(FCDate(MaskEdBox1))
'    'end 2014/11/24
'    FormDisabled
'    'Text1.Enabled = True 'Modify by Amy 2014/01/15
'    Text2.Enabled = True
'    Text2.SetFocus 'Modify by Amy 2014/01/17
End Sub

'由aacc_save.bas 搬回
Private Sub Frmacc4120_Save()
Dim strSave, strSql As String
Dim bCancel As Boolean 'Add by Amy 2014/11/20

On Error GoTo Checking
'Mark by Amy 2024/08/05 改至ChkForm
'      If strSaveConfirm = MsgText(4) Then
'         strSql = "select a0209, a0210 from acc020 where a0201 = '" & Text1 & "' and a0202 = '" & Text2 & "'"
'         If CheckRecord(strSql, IIf(IsNull(adoacc020.Fields("a0209").Value), 0, adoacc020.Fields("a0209").Value), IIf(IsNull(adoacc020.Fields("a0210").Value), 0, adoacc020.Fields("a0210").Value)) = False Then
'            strControlButton = MsgText(602)
'            Text1.SetFocus
'            Exit Sub
'         End If
'      End If
'      If Text1 = MsgText(601) Then
'         MsgBox MsgText(10) & Label3, , MsgText(5)
'         strControlButton = MsgText(602)
'         Text1.SetFocus
'         Exit Sub
'      Else
'         If Text2 = MsgText(601) Then
'            MsgBox MsgText(10) & Label2, , MsgText(5)
'            strControlButton = MsgText(602)
'            Text2.SetFocus
'            Exit Sub
'         End If
'         'Modify by Amy 2014/11/20 原程式修改至MaskEdBox1_Validate
'         Call MaskEdBox1_Validate(bCancel)
'         If bCancel = True Then
'            strControlButton = MsgText(602)
'            MaskEdBox1.SetFocus
'            Exit Sub
'         End If
'         'end 2014/11/20
'         If ExistCheck("acc080", "a0801", Text1, Label3) = False Then
'            strControlButton = MsgText(602)
'            Text1.SetFocus
'            Exit Sub
'         End If
'      End If
'      'add by sonia 2015/12/30 檢查民國105年起法務收入科目不可使用
'      intI = PUB_AccNoEnable(Text14, Val(FCDate(MaskEdBox1.Text)))
'      If intI <> 0 Then
'         strControlButton = MsgText(602)
'         Text14.SetFocus
'         Exit Sub
'      End If
'      'end 2015/12/30
'      'Add by Morgan 2007/2/5 檢查科目部門&智權人員是否正確
'      'Modify by Amy 2021/03/08 +傳票號
'      intI = PUB_AccNoGood(Text14, Text16, Text8, , Text2)
'      If intI <> 0 Then
'         strControlButton = MsgText(602)
'         If intI = 1 Then
'            Text14.SetFocus
'         ElseIf intI = 2 Then
'            Text16.SetFocus
'         ElseIf intI = 3 Then
'            Text8.SetFocus
'         End If
'         Exit Sub
'      End If
'      'end 2007/2/5

   'Modify by Amy 2024/08/05 原判斷"And ax205 = '" & Text14 & "' ",存Acc021時已判斷ax203,故拿掉,判斷是否過帳改至ChkForm
   'Memo by Amy  判斷Acc021 因為曾經財務操作有Acc021但沒Acc020-秀玲
   strSql = "Select * From Acc021 Where ax201||''= '" & Text1 & "' and ax202 = '" & Text2 & "' " & _
                  "Order by ax201 asc, ax202 asc, ax203 asc"
   If adoacc021.State = adStateOpen Then adoacc021.Close
   adoacc021.CursorLocation = adUseClient
   adoacc021.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   '新增
   If strSaveConfirm = MsgText(3) Then
      If adoacc020.RecordCount <> 0 Then
         adoacc020.Find "a0201 = '" & Text1 & "'", 0, adSearchForward, 1
         If adoacc020.EOF = False Then
            adoacc020.Find "a0202 = '" & Text2 & "'", 0, adSearchForward, adoacc020.Bookmark
            If adoacc020.EOF = False Then
               Exit Sub
            End If
         End If
      End If
      adoacc020.AddNew
   End If
      
   adoacc020.Fields("a0201").Value = Text1
   If Text2 <> MsgText(601) Then
      adoacc020.Fields("a0202").Value = Text2
   Else
      adoacc020.Fields("a0202").Value = Null
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      adoacc020.Fields("a0205").Value = Val(FCDate(MaskEdBox1.Text))
   Else
      adoacc020.Fields("a0205").Value = Null
   End If
   If strSaveConfirm = MsgText(3) Then
      adoacc020.Fields("a0206").Value = Val(strSrvDate(2))
      adoacc020.Fields("a0207").Value = ServerTime
      adoacc020.Fields("a0208").Value = strUserNum
      '911021 nick 因為少傳年和月
      'strSave = AccSaveAutoNo(MsgText(801), Mid(.Text2, 7, 4))
      'Modify by Amy 2014/01/15 原:MsgText(801)
      strSave = AccSaveAutoNo(strA1R01, Mid(Text2, 7, 4), Mid(Text2, 2, 3), Mid(Text2, 5, 2))
   Else
      adoacc020.Fields("a0209").Value = Val(strSrvDate(2))
      adoacc020.Fields("a0210").Value = ServerTime
      adoacc020.Fields("a0211").Value = strUserNum
   End If
   adoacc020.UpdateBatch
   RecordShow
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub
'end 2013/01/06

'Add by Amy 2014/01/15
'Modify by Amy  2014/02/20 Mark 因改由frmacc41c2
'Public Sub ChgCompany(Optional strDefault As String = "")
'   Do
'      strExc(0) = InputBox("請輸入主要公司別!! ( 1 or J )", , strDefault)
'
'      If strExc(0) = "" Then
'         Exit Do
'      ElseIf strExc(0) <> "1" And UCase(strExc(0)) <> "J" Then
'         MsgBox "只可輸入 1 或 J", vbCritical
'      Else
'         Exit Do
'      End If
'   Loop
'End Sub
'end 2014/02/20

Private Function GetSeqNo(strax201 As String, strax202 As String) As String
    '取得項次
    Dim adoaccmax As New ADODB.Recordset
    
    If adoaccmax.State = adStateOpen Then
         adoaccmax.Close
    End If
    adoaccmax.CursorLocation = adUseClient
    'Modify by Amy 2017/05/12 +Having 避免max() 沒資料產生E-Fail的錯誤
    adoaccmax.Open "select max(ax203) from acc021 where ax201 = '" & strax201 & "' and ax202 = '" & strax202 & "' Having max(ax203) is not null", adoTaie, adOpenStatic, adLockReadOnly
    If adoaccmax.RecordCount <> 0 Then
       GetSeqNo = ZeroBeforeNo(Val(adoaccmax.Fields(0).Value), 3)
    Else
       GetSeqNo = ZeroBeforeNo(0, 3)
    End If
    adoaccmax.Close
End Function
'end 2014/01/15

'Mark by Amy 2024/08/05 不使用,檢查改至ChkForm
'由acc_var搬回並改寫
Public Function FormF2Check() As Boolean
'   Dim bCancel As Boolean 'Add by Amy 2023/03/07
'
'   FormF2Check = True
'   '2009/7/27 ADD BY SONIA 正在轉批次時不可按新增
'   If adocheck.State = adStateOpen Then
'      adocheck.Close
'   End If
'   adocheck.CursorLocation = adUseClient
'   '2014/1/22 modify by sonia 加公司別a0b04
'   'adocheck.Open "select a0b10 from acc0b0 where a0b10 = '01'", adoTaie, adOpenStatic, adLockReadOnly
'   adocheck.Open "select a0b10 from acc0b0 where a0b04 = '" & Text1 & "' and a0b10 = '01'", adoTaie, adOpenStatic, adLockReadOnly
'   If adocheck.RecordCount <> 0 Then
'      MsgBox MsgText(197), , MsgText(5)
'      adocheck.Close
'      FormF2Check = False 'Add by Amy 2014/01/14
'      strSaveConfirm = MsgText(601)
'      Exit Function
'   End If
'   adocheck.Close
'   '2009/7/27 END
'    If MaskEdBox1.Text = MsgText(29) Or MaskEdBox1.Text = MsgText(601) Then
'        MaskEdBox1.Mask = MsgText(601)
'        MaskEdBox1.Text = CFDate(strSrvDate(2))
'        MaskEdBox1.Mask = DFormat
'    Else
'        Call MaskEdBox1_Validate(bCancel)
'        If bCancel = True Then
'            FormF2Check = False
'            strSaveConfirm = MsgText(601)
'            Exit Function
'        End If
'    End If
'
'   CreDebCheck
'   If CreDebCheck <> MsgText(602) Then
'      MsgBox MsgText(11), , MsgText(5)
'      FormF2Check = False 'Add by Amy 2014/01/14
'      strSaveConfirm = MsgText(601)
'      Exit Function
'   End If
End Function

Private Sub Frmacc4120_Delete()
   Dim bolInTrans As Boolean
   Dim stSQL As String, intR As Integer
   
On Error GoTo Checking
   
   'Memo by Amy 2024/08/05 原DeleteCheck 檢查 改至ChkForm
   'Modified by Morgan 2023/5/23 加 Transaction
   adoTaie.BeginTrans
   bolInTrans = True
   'end 2023/5/23
   
   adoTaie.Execute "delete from acc021 where ax201 = '" & Text1 & "' and ax202 = '" & Text2 & "'"
   'Add by Amy 2017/09/14 刪acc450進項發票資料
   adoTaie.Execute "Delete From acc450 Where a4501 = '" & Text2 & "'"
   adoTaie.Execute "delete from acc020 where a0201 = '" & Text1 & "' and a0202 = '" & Text2 & "'"
   
   'Added by Morgan 2023/5/23 若刪除為最大傳票號時還原自動編號
   stSQL = "update acc1r0 a set a1r04=a1r04-1" & _
      " where a1r01='" & strA1R01 & "'" & _
      " and a1r02=" & (Val(Mid(Text2, 2, 3)) + 1911) & _
      " and a1r03=" & Val(Mid(Text2, 5, 2)) & _
      " and a1r04=" & Val(Mid(Text2, 7))
   adoTaie.Execute stSQL, intR
   
   adoTaie.CommitTrans
   bolInTrans = False
   'end 2023/5/23
   
   'Mark by Amy 2024/08/05 以下程式改至SaveData並修改
'   AdodcRefresh
'   AdodcClear
'   adoacc020.Requery
'   If adoacc020.RecordCount <> 0 Then
'      adoacc020.MoveFirst
'      RecordShow
'   Else
'      StatusClear
'   End If
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   If bolInTrans Then adoTaie.RollbackTrans 'Added by Morgan 2023/5/23
   
   MsgBox Err.Description, , MsgText(5)
End Sub

Public Sub Frmacc4120_Clear()
   'Modify by Amy 2014/01/15
   'Text1 = ""
   'Text13 = ""
   If MaskEdBox1.Text = MsgText(29) Or MaskEdBox1.Text = MsgText(601) Then
      MaskEdBox1.Mask = ""
      MaskEdBox1.Text = ""
      MaskEdBox1.Mask = DFormat
   End If
   'Text2 = ""  'Modify by Amy 2014/01/15
   Text11 = ""
   Text12 = ""
   'Modify by Amy 2014/11/20
   If strSaveConfirm = MsgText(601) Then
        Text2.SetFocus
   End If
   strControlButton = "" 'Add by Amy 2024/08/05
   
End Sub

Public Sub Frmacc4120_Previous()
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
      m_CurrKEY = Text2 'Add by Amy 2024/08/05
   End If
   AdodcClear
End Sub

Public Sub Frmacc4120_Next()
   CreDebCheck
   If CreDebCheck <> MsgText(602) Then
      MsgBox MsgText(11), , MsgText(5)
      Exit Sub
   End If
   If adoacc020.EOF = False Then
      adoacc020.MoveNext
      If adoacc020.EOF Then
         adoacc020.MoveLast
         MsgBox MsgText(8), , MsgText(5)
      End If
      FormShow
      AdodcRefresh
      SumShow
      m_CurrKEY = Text2 'Add by Amy 2024/08/05
   End If
   AdodcClear
   RecordShow
End Sub

Public Sub Frmacc4120_First()
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
      m_CurrKEY = Text2 'Add by Amy 2024/08/05
   End If
   AdodcClear
   RecordShow
End Sub

Public Sub Frmacc4120_Last()
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
      m_CurrKEY = Text2 'Add by Amy 2024/08/05
   End If
   AdodcClear
   RecordShow
End Sub
'end 2014/01/14

'Modify by Amy 2014/11/20 重新整理
Public Sub SetData(ByVal strKeyCode As String)
    Select Case strKeyCode
        Case "F2-1" '新增(檢查前設定)
            'Text1.Enabled = False 'Modify by Amy 2014/01/15
            Text2 = "" 'Add by Amy 2024/08/05
            Combo1.Clear
            Frmacc4120_Clear
            AdodcClear
            AdodcRefresh
        Case "F2-2" '新增(檢查完後設定)
            'Text1 = "1"  Modify by Amy 2014/01/15
            Text2.Enabled = False
            If MaskEdBox1.Text = MsgText(29) Or MaskEdBox1.Text = MsgText(601) Then
               MaskEdBox1.Mask = MsgText(601)
               MaskEdBox1.Text = CFDate(strSrvDate(2))
               MaskEdBox1.Mask = DFormat
            End If
            MaskEdBox1.Tag = FCDate(MaskEdBox1.Text) 'Add by Amy 2023/04/19
            '傳票號
            Text2 = AccAutoNo(strA1R01, 4, Val(Mid(MaskEdBox1.Text, 1, 3)), Val(Mid(MaskEdBox1.Text, 5, 2))) 'Add by Amy 2014/01/15
            AdodcRefresh
            FormEnabled
            Text3 = ZeroBeforeNo(MsgText(12), 3)
            SumShow
            Text14.SetFocus
        Case "F3" '修改
            'Memo by Amy 2024/08/05 原'Add by Amy 2017/04/18 系統產生的傳票不可任意修改 檢查改至ChkForm
            FormEnabled
        Case "F5" '刪除
            'Memo by Amy 2024/08/05 開放都可刪除,原'Add by Amy 2023/12/06 是否過帳及 2022/05/13 系統產生的傳票不可任意修改 檢查改至ChkForm
            'Memo by Amy 2023/12/06 有刪除鈕是因財務可能刪上個月最後一筆未過帳傳票,若改程式需刪除及新增(刪上個月最後一筆又加傳票)都需加判斷
            '                                                     故避免傳票不連號,拿掉刪除鈕,由電腦中心刪 ex:1121206 旻霖 刪了1公司D112120003,但已有04傳票
            strSaveConfirm = "D" '加ChkForm,檢查有誤strSaveConfirm=""會跳離,需先設值否則會無法往下做
        Case "F10" '取消
            'Modify by Amy 2024/08/05 避免資料殘留
            '  ex:一進入 按新增->已產生傳票號->取消->修改->Insert 資料會產生EOF錯誤
            '  ex:查 D113060031 (傳票日1130603)->改傳票日1130628 ->新增->彈已月結訊息 (保留於先前資料,避免日期與傳票號不一致,故記錄前一次查詢資料)
            '  ex:修改某資料->增加項次(004)->Insert (項次005) ->取消 項次編號不會刪除->再修改(項次仍是005)
            Text2 = m_CurrKEY
            If Text2 <> MsgText(601) Then
               Command3_Click
            Else
               'Memo by Amy 2024/08/05 原Mark by Amy 2022/05/13 因改寫法,不會再有錯誤
               'Mark by Amy 2022/05/13 '按修改->Insert->取消->修改->Insert->存檔會出現「找不到要更新的資料列。最後取的值已被變更」,無法修改
               Frmacc4120_Clear
               AdodcRefresh
            End If
            strSaveConfirm = MsgText(601) 'Add by Amy 2022/05/16
            AdodcClear
            'end 2027/07/19
            FormDisabled
            'end 2022/05/13
            'Text1.Enabled = True 'Modify by Amy 2014/01/15
            If Text2.Enabled = True Then Text2.SetFocus
            'Add by Amy 2024/08/05
            If HasUpdTag = True Then
               strContent = "": HasUpdTag = False
            End If
        Case "BtExit"
            strControlButton = MsgText(601) 'Add by Amy 2024/08/05
            'Add by Amy 2017/04/18 從其他支進入Menu鎖住,取消後只能用離開/修改 鈕
            If Len(Me.Tag) > 0 Then
                MenuDisabled
                tool7_enabled
            'Modify by Amy 2024/08/05 程式已判斷傳票不連號問題,故開放都可使用刪除
            Else
               tool1_enabled
            End If
'            'Add by Amy 2024/02/05 M51可使用刪除
'            ElseIf Pub_StrUserSt03 = "M51" Then
'               tool1_enabled
'            'Add by Amy 2017/08/23 非自動產生傳票bug修正
'            Else
'                'Add by Amy 2023/12/06 避免傳票不連號,取消刪除鈕 原:tool1_enabled
'               tool14_enabled
'            End If
             'end 2024/08/05
        Case Else
    End Select
End Sub

'Add by Amy 2014/11/19 特殊出名公司 copy frmacc1121 Function 修改
'Modify by Amy 2024/08/05 加stChoose,讓其他檢查也可用
Private Function ChkPatentNameCompany(stChoose As String, pPA01 As String, pPA02 As String, pPA03 As String, pPA04 As String) As String
   Dim stSQL As String, adoRst As ADODB.Recordset, intR As Integer
   Dim stField(4) As String, stWhere(4) As String, stFixWhr As String, bolSql As Boolean 'Add by Amy 2024/08/05
   
   ChkPatentNameCompany = ""
   bolSql = False
   stFixWhr = "PA01='" & pPA01 & "' And PA02='" & pPA02 & "' And PA03='" & pPA03 & "' And PA04='" & pPA04 & "' "
   stWhere(0) = stFixWhr
   stWhere(1) = Replace(stFixWhr, "PA", "TM")
   stWhere(2) = Replace(stFixWhr, "PA", "SP")
   stWhere(3) = Replace(stFixWhr, "PA", "LC")
   stWhere(4) = Replace(stFixWhr, "PA", "HC")
   Select Case stChoose
      Case "1" '回傳 特殊出名公司
         stField(0) = "pa161"
         stField(1) = "tm130"
         stField(2) = "sp85"
         stField(3) = "Decode(LC01,'ACS',LC48,'L')" '法務-ACS案抓LC48,其餘都設L公司
         stField(4) = "'L' as SpecComp" '顧問
         
         stWhere(0) = stWhere(0) & "And pa161 is not null "
         stWhere(1) = stWhere(1) & "And tm130 is not null "
         stWhere(2) = stWhere(2) & "And sp85 is not null "
      Case "2" '回傳 語法
         bolSql = True
         stField(0) = "pa09 as Nation,pa26 as Apply,pa01 as SystemNo,pa161 as SpecComp"
         stField(1) = "tm10 as Nation,tm23 as Apply,tm01 as SystemNo,tm130 as SpecComp"
         stField(2) = "sp09 as Nation,sp08 as Apply,sp01 as SystemNo,sp85 as SpecComp"
         stField(3) = "lc15 as Nation,lc11 as Apply,lc01 as SystemNo,Decode(LC01,'ACS',LC48,'L') as SpecComp"
         stField(4) = "'000' as Nation,hc07 as Apply,hc01 as SystemNo,'L' as SpecComp"
         
   End Select
   'Modify by Amy 2024/04/03 +顧問
   stSQL = "Select " & stField(0) & " From patent Where " & stWhere(0) & _
     " Union Select " & stField(1) & " From trademark Where " & stWhere(1) & _
     " Union Select " & stField(2) & " From servicepractice Where " & stWhere(2) & _
     " Union Select " & stField(3) & " From lawcase Where " & stWhere(3) & _
     " Union Select " & stField(4) & " From hirecase Where " & stWhere(4)
     
   If bolSql = True Then ChkPatentNameCompany = stSQL: Exit Function
   
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      ChkPatentNameCompany = Trim("" & adoRst.Fields(0).Value)
   End If
End Function

'Add by Amy 2014/11/20 傳票日期檢查
'Modify by Amy 2024/08/05 原:Private /+IsF2
Public Function ChkA0205(Optional ByVal IsF2 As Boolean = False) As Boolean
    Dim strQuery As String, RsQ As ADODB.Recordset, intQ As Integer
    Dim strMsg As String, strTpDate As String
    Dim m_A0B02 As String, m_A0B03 As String, bCancel As Boolean 'Add by Amy 2024/08/05
    
    ChkA0205 = False
    
    'Add by Amy 2024/08/05 原新增及新增存檔判斷 看FormCheck及FormF2Check
    '按 新增 鈕
    If IsF2 = True Then
      '正在轉批次時不可按新增
      If Pub_GetField("acc0b0", "a0b04 = '" & Text1 & "' and a0b10 = '01'", "a0b10") <> MsgText(601) Then
         MsgBox MsgText(197), , MsgText(5)
         Exit Function
      End If
      '新增時若傳票日已月結則提醒
      If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
         m_A0B03 = Pub_GetField("acc0b0", "a0b04 = '" & Text1 & "'", "a0b03")
         m_A0B02 = Pub_GetField("acc0b0", "a0b04 = '" & Text1 & "'", "a0b02")
         '畫面傳票日小於等於[年]結日,不可再輸[當年]以前的傳票-秀玲
         '畫面傳票日小於等於[月]結日,不可再輸[月結日當月]以前的傳票
         If Left(Val(FCDate(MaskEdBox1.Text)) + 19110000, 4) <= Left(Val(m_A0B03) + 19110000, 4) Then
            MsgBox "目前年結日為" & CFDate(m_A0B03) & vbCrLf & _
                              "故不可再輸" & Val(Left(Val(m_A0B03) + 19110000, 4)) - 1911 & "年以前的傳票", , MsgText(5)
            Exit Function
         ElseIf Left(Val(FCDate(MaskEdBox1.Text)) + 19110000, 6) <= Left(Val(m_A0B02) + 19110000, 6) Then
            MsgBox "目前月結日為" & CFDate(m_A0B02) & vbCrLf & _
                             "故不可再輸月結日當月以前的傳票", , MsgText(5)
            Exit Function
         End If
      End If
    End If
    
    Call MaskEdBox1_Validate(bCancel)
    If bCancel = True Then
       Exit Function
    End If
    'end 2024/08/05
    
    'Modify by Amy 2023/04/19 日期與傳票號檢查
    'ex:一進入先輸11204月傳票日->按新增->產生傳票號 ->改傳票日為11203月日期
    If Text2 <> MsgText(601) Then
        If Val(Mid(FCDate(MaskEdBox1), 1, Len(FCDate(MaskEdBox1)) - 2)) <> Val(Mid(Text2, 2, 5)) Then
            '不能跨月(因傳票編號, 所以一定要重key)
            MsgBox Label9 & "跨月不可修改", , MsgText(5)
            MaskEdBox1.SetFocus
            Exit Function
        End If
    End If
    If strSaveConfirm = MsgText(3) Then
        'ex:目前最大傳票日為1120418,一進入查1120411傳票日後->按新增->產生傳票號(日期會是1120411,日期與傳票號會不連續）
        strExc(1) = Pub_GetMaxA0205(Text1, Val(Left(Val(FCDate(MaskEdBox1)) + 19110000, 6)) - 191100)
        If Val(FCDate(MaskEdBox1.Text)) < Val(strExc(1)) Then
            MsgBox Label9 & "不可小於目前傳票日(" & CFDate(strExc(1)) & ")", , MsgText(5)
            Exit Function
        End If
    'end 2023/04/19
    ElseIf MaskEdBox1.Tag <> MsgText(601) Then
        '修改
        'Modify by Amy 2014/11/24 再改往前及往後日期修改判斷
        If Val(MaskEdBox1.Tag) > Val(FCDate(MaskEdBox1)) Then
            '日期改往前且有前一筆,日期只能為 前一筆日期<=改後日期<原日期
            strQuery = "Select A0205 From acc020 Where A0201='" & Text1 & "' " & _
                              "And A0202=(Select Max(A0202) From acc020 Where A0201='" & Text1 & "' And A0202<'" & Text2 & "') "
            intQ = 1
            Set RsQ = ClsLawReadRstMsg(intQ, strQuery)
            If intQ = 1 Then
                'If Val(FCDate(MaskEdBox1)) <> Val(RsQ.Fields("A0205")) Then
                If Val(RsQ.Fields("A0205")) <= Val(FCDate(MaskEdBox1)) And Val(FCDate(MaskEdBox1)) < Val(MaskEdBox1.Tag) Then
                Else
                    If Val(RsQ.Fields("A0205")) = Val(MaskEdBox1.Tag) Then
                        strMsg = CFDate(RsQ.Fields("A0205"))
                    Else
                         strMsg = CFDate(RsQ.Fields("A0205")) & "~" & CFDate(MaskEdBox1.Tag)
                    End If
                    MsgBox Label9 & "只能輸 " & strMsg, , MsgText(5)
                    Exit Function
                End If
            End If
        End If
        
        If Val(MaskEdBox1.Tag) < Val(FCDate(MaskEdBox1)) Then
            strQuery = "Select A0205 From acc020 Where A0201='" & Text1 & "' " & _
                             "And A0202=(Select Min(A0202) From acc020 Where A0201='" & Text1 & "' And A0202>'" & Text2 & "') "
            intQ = 1
            Set RsQ = ClsLawReadRstMsg(intQ, strQuery)
            If intQ = 1 Then
                '日期改往後,有 後一筆資料, 原日期<改後日期<=後一筆日期
                'If Val(FCDate(MaskEdBox1)) <> Val(RsQ.Fields("A0205")) Then
                If Val(MaskEdBox1.Tag) < Val(FCDate(MaskEdBox1)) And Val(FCDate(MaskEdBox1)) <= Val(RsQ.Fields("A0205")) Then
                Else
                    If Val(RsQ.Fields("A0205")) = Val(MaskEdBox1.Tag) Then
                        strMsg = CFDate(RsQ.Fields("A0205"))
                    Else
                         strMsg = CFDate(MaskEdBox1.Tag) & "~" & CFDate(RsQ.Fields("A0205"))
                    End If
                    MsgBox Label9 & "只能輸 " & strMsg, , MsgText(5)
                    Exit Function
                End If
            Else
                '日期改往後,無 後一筆資料,只能改原日期的下一個工作日
                strTpDate = PUB_GetWorkDayAfterSysDate(CDbl(MaskEdBox1.Tag) + 19110000, 1)
                If Val(strTpDate) <> Val(FCDate(MaskEdBox1)) Then
                    MsgBox Label9 & "只能輸 " & CFDate(strTpDate), , MsgText(5)
                    Exit Function
                End If
            End If
        End If
        'end 2014/11/24
    End If
   ChkA0205 = True
End Function

'Add by Amy 2017/04/18 確認自動轉傳票 4194科目之Ax204是否為 null
Private Function ChkAx204Null() As Boolean
    Dim adoQ As New ADODB.Recordset
    Dim strQ As String
    Dim intQ As Integer
    
    ChkAx204Null = False
    strQ = "Select Ax204 From Acc021 Where ax201='1' And ax202='" & Text2 & "' And ax204 is null And ax205='4194' "
    intQ = 1
    Set adoQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        ChkAx204Null = True
    End If
    adoQ.Close
    
End Function

Private Function Chk2491InCome() As Boolean
    Dim i As Integer
    Dim bol2491Debit As Boolean, bol2491Credit As Boolean '2491借/貸
    Dim bol4Credit As Boolean '收入貸
  
    Chk2491InCome = False
    
    Adodc1.Recordset.MoveFirst
    For i = 1 To Adodc1.Recordset.RecordCount
        If Mid(Adodc1.Recordset.Fields("ax205").Value, 1, 4) = "2491" Then
            If Val("" & Adodc1.Recordset.Fields("ax206").Value) > 0 Then bol2491Debit = True
            If Val("" & Adodc1.Recordset.Fields("ax207").Value) > 0 Then bol2491Credit = True
        End If
        If Mid(Adodc1.Recordset.Fields("ax205").Value, 1, 1) = "4" And Val("" & Adodc1.Recordset.Fields("ax207").Value) > 0 Then
            bol4Credit = True
        End If
        Adodc1.Recordset.MoveNext
    Next i
    If bol2491Debit = True And bol2491Credit = True And bol4Credit = True Then
        Chk2491InCome = True
    End If
End Function

'Add by Amy 2021/06/02
Private Sub MailData(ByRef strContent As String)
    Dim strTo As String, strSubject As String
    
    strContent = ""
    If Text14.Tag <> Text14 Then
        strContent = strContent & "會計科目：" & Text14.Tag & " --> " & Text14 & vbCrLf
    End If
    If Text4.Tag <> Text4 Then
        strContent = strContent & "借　　方：" & Text4.Tag & " --> " & Text4 & vbCrLf
    End If
    If Text5.Tag <> Text5 Then
        strContent = strContent & "貸　　方：" & Text5.Tag & " --> " & Text5 & vbCrLf
    End If
    If Text6.Tag <> Text6 Then
        strContent = strContent & "對沖其他：" & Text6.Tag & " --> " & Text6 & vbCrLf
    End If
    If Text8.Tag <> Text8 Then
        strContent = strContent & "對沖業務：" & Text8.Tag & " --> " & Text8 & vbCrLf
    End If
    
    'Mark by Amy 2021/06/02 先不使用,無法正確判定是否存檔,目前先發給Amy 好判斷財務做了什麼,才能知道點數如何處理!
'    If strContent <> MsgText(601) Then
'        strTo = Pub_GetSpecMan("財務處總帳人員") & ";A2004"
'        strSubject = Text1 & "公司　傳票" & Text2 & "修改"
'        strContent = strSubject & "如下：" & vbCrLf & _
'                            strContent & vbCrLf & _
'                            "智權輸入已開放且已產生SalesBalance資料" & vbCrLf & _
'                            "請告知電腦中心修改項目，判斷是否刪除已產生之SalesBalance資料！！"
'
'        PUB_SendMail strUserNum, strTo, "", strSubject, strContent
'    End If
End Sub

'Modify by Amy 2021/09/23 從 acc021Save搬過來修改,結餘資料有修改寫Tag (Axb16)
Private Sub ChkSetAxb16(ByVal stState As String)
    Dim strYM As String, strSPYM As String

    strYM = Left(Val(FCDate(MaskEdBox1)) + 19110000, 6)
    strSPYM = Left(strSrvDate(1), 6)
    If Right(strSPYM, 2) = "01" Then
         strSPYM = Val(Left(strSPYM, 4)) - 1912 & "12"
    Else
         strSPYM = Val(strSPYM) - 1
    End If
    '傳票為系統月-1 且 SalesPoint及SalesBalance 已有系統月-1資料 之需新增SalesPoint人員 且 會科為4字頭且為「結餘」資料,判斷修改欄位為結餘相關寫Tag(Axb16)
    If strYM = strSPYM Then
        If ExistCheck("SalesPoint", "SP01", strSPYM, "", False) = True And ExistCheck("SalesBalance", "SB01", Val(strSPYM) - 191100, "", False) = True And InStr(不需新增SalesPoint人員, Text8) = 0 Then
            'Insert 鍵
            If stState = "Ins" Then
                 '原為 4字頭 對沖-其他 是結餘 -> 4字頭 「不是」結餘 Or 「不是」4字頭 / 原「不是」4字頭->4字頭 結餘
                 If ((Left(Text14.Tag, 1) = "4" And InStr(Text6.Tag, "結餘") > 0) And ((Left(Text14, 1) = "4" Or InStr(Text6, "結餘") = 0) Or (Left(Text14.Tag, 1) = "4" And Left(Text14, 1) <> "4"))) _
                   Or (Left(Text14.Tag, 1) <> "4" And (Left(Text14, 1) = "4" And InStr(Text6, "結餘") > 0)) Then
                     Call MailData(strContent)
                     If WirteAxb16(Val(strYM) - 191100, "Y") = True Then
                        HasUpdTag = True
                     End If
                 '4字頭修改 借 貸 對沖-業
                 ElseIf Left(Text14.Tag, 1) = "4" And InStr(Text6.Tag, "結餘") > 0 And Left(Text14, 1) = "4" And InStr(Text6, "結餘") > 0 _
                   And (Text4.Tag <> Text4 Or Text5.Tag <> Text5 Or Text8.Tag <> Text8) Then
                     Call MailData(strContent)
                     If WirteAxb16(Val(strYM) - 191100, "Y") = True Then
                        HasUpdTag = True
                     End If
                 End If
            '刪除(剪刀)
            ElseIf Left(Text14, 1) = "4" And InStr(Text6, "結餘") > 0 Then
                strContent = "刪除 4字頭 結餘 資料"
                If WirteAxb16(Val(strYM) - 191100, "Y") = True Then
                    HasUpdTag = True
                End If
            End If
        End If
    End If
End Sub

'Add by Amy 2022/05/16 檢查每個人的總額必須與SalesPoint相符
Private Function ChkSalesPointVal() As Boolean
    Dim rsQ1 As New ADODB.Recordset, RsQ2 As New ADODB.Recordset
    Dim intQ1 As Integer, intQ2 As Integer, strQ1 As String, strQ2 As String, strQ2_Fix As String
    Dim strAcDate As String, strField As String, stMsg As String
  
    ChkSalesPointVal = False: stMsg = ""
    
    strQ2_Fix = "Select Sum(ax207-ax206) stVal,ax209,st02 From Acc021,Staff " & _
                      "Where ax201='1' And ax202='" & Text2.Text & "' " & _
                      "And ax209=st01(+) "
    
    strAcDate = Val(Mid(Val(FCDate(MaskEdBox1.Text)) + 19110000, 1, 6)) - 191100
    strQ1 = GetPoint_SP(strAcDate, strAcDate, , , "SP40", False, "Frmacc41h0", , True)
    intQ1 = 1
    Set rsQ1 = ClsLawReadRstMsg(intQ1, strQ1)
    If intQ1 = 1 Then
        rsQ1.MoveFirst
        Do While rsQ1.EOF = False
            intQ2 = 1
            strQ2 = strQ2_Fix & " And ax209='" & rsQ1.Fields("SP02") & "' Group by ax209,st02"
            Set RsQ2 = ClsLawReadRstMsg(intQ2, strQ2)
            If intQ2 = 1 Then
                If Val("" & rsQ1.Fields("SP40")) <> Val("" & RsQ2.Fields("stVal")) Then
                    stMsg = stMsg & "," & RsQ2.Fields("st02")
                End If
            End If
            rsQ1.MoveNext
        Loop
        If stMsg <> MsgText(601) Then
            ChkSalesPointVal = True
            stMsg = Mid(stMsg, 2)
            If InStr(stMsg, ",") > 0 Then stMsg = stMsg & vbCrLf '多人加換行
            MsgBox stMsg & "個人總額不符合結餘點數轉撥", , MsgText(5)
        End If
    End If
    Set rsQ1 = Nothing
    Set RsQ2 = Nothing
End Function

'Add by Amy 2022/09/21 確認 S案申請國家之會計科目
'Mark by Amy 2024/08/05 不使用ChkCaseNoAndAccNo已檢查
'Private Function ChkSCaseAccNO(ByVal stAccNo As String, ByVal stCaseNo As String) As Boolean
'    Dim RsQ As New ADODB.Recordset
'    Dim intQ As Integer, strQ As String
'
'    'Add by Amy 2022/10/26 只判斷S案
'    If Mid(stCaseNo, 1, Len(stCaseNo) - 9) <> "S" Then ChkSCaseAccNO = True: Exit Function
'
'    strQ = "Select * From ServicePractice " & _
'              "Where SP01='" & Mid(stCaseNo, 1, Len(stCaseNo) - 9) & "' And SP02='" & Mid(stCaseNo, Len(stCaseNo) - 8, 6) & "' And SP03='" & Mid(stCaseNo, Len(stCaseNo) - 2, 1) & "' And SP04='" & Mid(stCaseNo, Len(stCaseNo) - 1, Len(stCaseNo)) & "' "
'    intQ = 1
'    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
'    If intQ = 1 Then
'        If "" & RsQ.Fields("SP09") = "000" Then
'            If stAccNo = "220103" Then
'                ChkSCaseAccNO = True
'            Else
'                MsgBox "此S案申請國家為台灣,會計科目只能使用220103"
'            End If
'        Else
'            If stAccNo = "220105" Then
'                ChkSCaseAccNO = True
'            Else
'                MsgBox "此S案申請國家非台灣,會計科目只能使用220105"
'            End If
'        End If
'    End If
'    Set RsQ = Nothing
'End Function

'Add by Amy 2023/06/14 檢查有借方24910及貸方41xxxx,避免直接Key隱藏版傳票的科目資料key錯
Private Function Chk24910xAnd41xxxx() As Boolean
    Dim i As Integer, bol24910Debit As Boolean, bol4Credit As Boolean '24910借/41字頭收入貸
    Dim stChk24910x As String, stChk41xxxx As String
  
    Chk24910xAnd41xxxx = False
    stChk24910x = "249101,249102,249103,249104" 'Memo 此有增加需確認 隱藏版是否需修改
    
    Adodc1.Recordset.MoveFirst
    For i = 1 To Adodc1.Recordset.RecordCount
        '有 借方 24910x
        If InStr(stChk24910x, Adodc1.Recordset.Fields("ax205").Value) > 0 And Val("" & Adodc1.Recordset.Fields("ax206").Value) > 0 Then
            bol24910Debit = True
            Select Case Adodc1.Recordset.Fields("ax205").Value
                Case "249101" 'T
                   stChk41xxxx = "410103"
                Case "249102" 'P
                  stChk41xxxx = "411103"
                Case "249103" 'CFT
                  stChk41xxxx = "412101"
                Case "249104" 'CFP
                  stChk41xxxx = "413101"
            End Select
        End If
        '有 貸方 相對應4字頭科目
        If stChk41xxxx = Adodc1.Recordset.Fields("ax205").Value And Val("" & Adodc1.Recordset.Fields("ax207").Value) > 0 Then
            bol4Credit = True
        End If
        Adodc1.Recordset.MoveNext
    Next i
    If bol24910Debit = True And bol4Credit = True Then
        Chk24910xAnd41xxxx = True
    End If
End Function

'Add by Amy 2024/08/05
'整合檢查程式,避免有未改到的(目前檢查 F3/Insert/F5/F9)
Public Function ChkForm(ByVal m_State As String) As Boolean
   Dim stSQL As String, stMsg As String, stTP As String, bCancel As Boolean
   Dim nResponse
   
   ChkForm = False
   'add by nickc 2005/08/26 將控制字元歸位
    strControlButton = ""
   
'*** 都要檢查 ***
   '公司別
   If Text1 = MsgText(601) Then
      MsgBox MsgText(10) & Label3, , MsgText(5)
      Text1.SetFocus
      Exit Function
   End If
   If ExistCheck("acc080", "a0801", Text1, Label3) = False Then
      Text1.SetFocus
      Exit Function
   End If
   '傳票日
   If MaskEdBox1.Text = MsgText(29) Or MaskEdBox1.Text = MsgText(601) Then
      MsgBox MsgText(10) & Label9, , MsgText(5)
      MaskEdBox1.SetFocus
      Exit Function
   End If
   '傳票號
   If Text2 = MsgText(601) Then
      MsgBox MsgText(10) & Label2, , MsgText(5)
      Text2.SetFocus
      Exit Function
   End If
'*** End 都要檢查 ***
   
'*** 個別鈕 優先檢查 ***
   '按[確定] 鈕
   If m_State = "F9" Then
      'Add by Amy 2024/08/05  明細仍有資料未Insert不可存,因目前明細資料檢查只寫於Insert
      If Text14 & Text4 & Text5 <> MsgText(601) Then
         MsgBox "仍有資料未【Insert】不可存檔" & vbCrLf & _
                           "若不 Insert 請按【垃圾桶】清除"
         Exit Function
      End If
      '修改確定 檢查是否有異動
      If strSaveConfirm = MsgText(4) Then
         stSQL = "select a0209, a0210 from acc020 where a0201 = '" & Text1 & "' and a0202 = '" & Text2 & "'"
         If CheckRecord(stSQL, IIf(IsNull(adoacc020.Fields("a0209").Value), 0, adoacc020.Fields("a0209").Value), IIf(IsNull(adoacc020.Fields("a0210").Value), 0, adoacc020.Fields("a0210").Value)) = False Then
            Text1.SetFocus
            Exit Function
         End If
      End If
   '按 Insert 鈕
   ElseIf m_State = "Ins" Then
      If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
         Exit Function
      End If
      'Add by Amy 2021/12/07 Form2.0控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
      If PUB_ChkTrackMode = False Then
         Exit Function
      End If
      'Add by Amy 2021/12/07 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
      If PUB_ChkUniText(Me) = False Then
         Exit Function
      End If
   '按 修改 / 刪除 鈕
   Else
      'Memo  by Amy 2023/12/06 有刪除鈕是因財務可能刪上個月最後一筆未過帳傳票,若改程式需刪除及新增(刪上個月最後一筆又加傳票)都需加判斷
      '                                                     故避免傳票不連號,拿掉刪除鈕,由電腦中心刪 ex:1121206 旻霖 刪了1公司D112120003,但已有04傳票
      '判斷是否過帳
      If Adodc1.Recordset.RecordCount <> 0 Then
         Adodc1.Recordset.MoveFirst
         If Not IsNull(Adodc1.Recordset.Fields("ax210").Value) Then
            MsgBox MsgText(14), , MsgText(5)
            Exit Function
         End If
      End If
      'Add by Amy 2017/04/18 系統產生的傳票不可任意修改(若需刪除需由電腦中心改axbXX的資料後再由電腦中心刪)
      '                                                從傳票輸入進入且是１公司修改的資料需判斷是否為智權點數產生之傳票
      If UCase(Me.Tag) = MsgText(601) And Text1 = "1" Then
         '抓智權點數傳票起始值
         If bolAcc0b1(0, Left(FCDate(MaskEdBox1.Text), 5), strAxb()) = True Then
            'Modify by Amy 2023/04/14 +strAxb(17)
            If (Text2 >= strAxb(4) And Text2 <= strAxb(5)) Or Text2 = strAxb(6) Or (Text2 >= strAxb(7) And Text2 <= strAxb(8)) _
               Or (Text2 >= strAxb(9) And Text2 <= strAxb(10)) Or Text2 = strAxb(11) Or (Text2 >= strAxb(12) And Text2 <= strAxb(13)) _
               Or Text2 = strAxb(14) Or Text2 = strAxb(17) Then
                  stMsg = "此為系統產生之傳票" & vbCrLf
                  If m_State = "F3" Or (m_State = "F5" And Pub_StrUserSt03 <> "M51") Then
                     If Text2.Locked = True Then Text2.Locked = False
                     MsgBox stMsg & "請勿任意修改！", , MsgText(5)
                     Exit Function
                  'Add by Amy 2024/08/05 電腦中心 刪除 彈訊息(目前刪除只有電腦中心可刪)
                  ElseIf MsgBox(stMsg & "[已]調整當月智權點數傳票資料(Acc0b1)？" & vbCrLf & _
                                            "確定刪除？" & vbCrLf & _
                                            "是:刪除　　否:回前畫面", vbYesNo + vbQuestion) = vbNo Then
                     Exit Function
                  End If
            End If
         End If
      End If
      '刪除
      If m_State = "F5" Then
         If DeleteCheck("select a0201 from acc020 where a0201 = '" & Text1 & "' and a0202 = '" & Text2 & "'") = MsgText(603) Then
            Exit Function
         End If
         'Add by Amy 2024/08/05 判斷畫面傳票編號後面仍有傳票不可刪
         stSQL = "a0201= '" & Text1 & "' And a0202>'" & Text2 & "' And SubStr(a0205+19110000,1,6)=" & Left(Val(FCDate(MaskEdBox1.Text)) + 19110000, 6)
         If Pub_GetField("Acc020", stSQL, "Min(a0202)") <> MsgText(601) Then
            MsgBox "此傳票後仍有其他傳票資料不可刪除！" & vbCrLf & "請移做他用", , MsgText(5)
            Exit Function
         End If
      End If
   End If
'*** End 個別鈕  優先檢查 ***
   
   'Add by Amy 2024/08/05 從MaskEdBox1_Validate搬過來(原因看MaskEdBox1_Validate Memo by Amy 2024/08/05 )
   If ChkA0205 = False Then
      MaskEdBox1.SetFocus
      Exit Function
   End If
   
   '按 修改 / 刪除 鈕
   If m_State = "F3" Or m_State = "F5" Then
      '修改 / 刪除 鈕 檢查只到此
      ChkForm = True: Exit Function
   End If

   If m_State = "Ins" Then
   '*** 會計科目 ***
      If Text14 = MsgText(601) Then
         MsgBox MsgText(10) & Label5, vbExclamation, "資料錯誤"
         TextInverse Text14
         Text14.SetFocus
         Exit Function
      End If
      Call Text14_Validate(bCancel)
      If bCancel = True Then
         Exit Function
      End If
      'Memo by Amy 2024/08/05 原:檢查民國105年起法務收入科目不可使用-秀玲
      '                                                    因PUB_AccNoEnable已無作用,故不需檢查(改每日檢查舊收據資料)
      
      'add by sonia 2017/3/22 借方的預付稅捐1203科目且並非應收/付轉來的傳票,一定要輸入對沖代號(其它)欄
      If Text14 = "1203" And Text6 = "" And Val(Text4) > 0 Then
         strExc(0) = "select a1p01,a1p22 from acc1p0 where a1p01='" & Text1.Text & "' and a1p22='" & Text2 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then  '應收/付轉來的傳票
         Else
            MsgBox "預付稅捐1203科目, 請輸入對沖代號(其它)欄！", vbExclamation, "資料錯誤"
            If Text6.Enabled = True Then
               TextInverse Text6
               Text6.SetFocus
            End If
            Exit Function
         End If
      End If
      'end 2017/3/22
      'add by sonia 2020/9/8 2491或2211對沖其他欄一定要輸
      If Text14 = "2211" And Text6 = "" Then
         MsgBox "此科目一定要輸對沖代號(其)欄！", vbExclamation, "資料錯誤"
         TextInverse Text6
         Text6.SetFocus
         Exit Function
      End If
      If Left(Text14, 4) = "2491" And InStr(Text6, "結餘") = 0 Then
         MsgBox "此科目一定要輸對沖代號(其)欄, 而且要有「結餘」字樣！", vbExclamation, "資料錯誤"
         TextInverse Text6
         Text6.SetFocus
         Exit Function
      End If
      '2020/9/8 end
      'Add by Amy 2023/04/20 490101(安全基金撥補)屬於「結餘」資料,若ax213(對沖-其他)未輸需彈訊息不可存檔
      'Memo 未加結餘 智權點數實績與結餘分析表 [不應]列於「當月實績」,而專業達成點數表(秘書用「當月實績」為0,不一致 ex:11201月
      If Text14 = "490101" Then
         'Memo by Amy 2024/08/05 原判斷 490101 部門只能是 TOT,於下方 部門 跳離開檢查(CheckDept)
         If InStr(Text6, "結餘") = 0 Then
            MsgBox "【" & Text15 & "】科目一定要輸對沖代號(其)欄" & vbCrLf & _
                              "而且要有「結餘」字樣！", vbExclamation, "資料錯誤"
            TextInverse Text6
            Text6.SetFocus
            Exit Function
         End If
         'Add by Amy 2023/06/14 彈提醒
         MsgBox "每月業績關閉後又加【" & Text15 & "】科目" & vbCrLf & _
                        "需至每月業績開放/關閉輸入 按「關閉後又加安全基金撥補請按此鈕」"
      End If
      'add by sonia 2015/4/22 41XX(除4191,4192,4194)外或7121摘要有[結餘],對沖其他欄也要有
      'Modify by Amy 2024/08/05 改4字頭
      If (Left(Text14, 1) = "4" And Text14 <> "4191" And Text14 <> "4192" And Text14 <> "4194") Or Text14 = "7121" Then
         If InStr(Combo1, "結餘") > 0 And InStr(Text6, "結餘") = 0 Then
            MsgBox "收文科目摘要欄內有【結餘】字樣" & vbCrLf & _
                              "對沖代號(其它)欄也要輸結餘！", vbExclamation, "資料錯誤"
            TextInverse Text6
            Text6.SetFocus
            Exit Function
         End If
      End If
      '2015/4/22 end
      'add by sonia 2019/9/5
      'Modify by Amy 2022/05/16 排除由結餘/實績 轉撥進入
      If Left(Text14, 1) = "4" And InStr(Combo1, "轉撥") > 0 And UCase(Me.Tag) <> "FRMACC41H0" And UCase(Me.Tag) <> "FRMACC41G0" Then
         nResponse = MsgBox("非轉撥傳票摘要不可輸入【轉撥】二字 " & vbCrLf & _
                                                   "否則會影響實績點數, 是否存檔?", vbOKCancel + vbDefaultButton2, MsgText(5))
         If nResponse = vbCancel Then
            Exit Function
         End If
      End If
      'end 2019/9/5
      'Add by Amy 2024/08/05 從acc021Save搬過來
      '6字頭無分攤類別,則 部門  不可為 空 或 TOT
      If Left(Text14, 1) = "6" Then
         If Pub_GetField("Acc010", "a0101 = '" & Text14 & "'", "a0105") = MsgText(601) _
           And (Text16 = MsgText(601) Or Text16 = MsgText(55)) Then
            MsgBox MsgText(198), , MsgText(5)
            Text16.SetFocus
            Exit Function
         End If
      End If
      'Add by Morgan 2007/4/4 翻譯費之本所案號檢查
      If Text14 = "6130" Then
         If Text10 = "" Then
            If MsgBox("你輸入的科目為【" & Text15 & "】但未輸本所案號，是否確定要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
               If Text10.Enabled = True Then
                  Text10.SetFocus
               End If
               Exit Function
            End If
         'Modified by Morgan 2019/2/27 +控制借方才檢查(因為有可能是要沖掉 Ex:FCP-59927)
         'Else
         ElseIf Val(Text4) > 0 Then
         'end 2019/2/27
         'Modified by Morgan 2007/8/8 加判斷不是自己這張傳票號
            strExc(0) = "select ax202 from acc021 where ax201='" & Text1.Text & "' and  ax205='" & Text14.Text & "' and ax214='" & Text10.Text & "'" & IIf(Text2 <> "", " and ax202<>'" & Text2 & "'", "")
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MsgBox "【" & Text10 & "】之【" & Text15 & "】已存在於傳票【" & RsTemp(0) & "】", vbExclamation, "資料重複"
               If Text10.Enabled = True Then
                  TextInverse Text10
                  Text10.SetFocus
               End If
               Exit Function
            End If
         End If
      End If
      'end 2007/4/4
   '*** 其他欄位 ***
      '借貸金額
      '金額不可同時輸入
      If Val(Text4) <> 0 And Val(Text5) <> 0 Then
         MsgBox MsgText(47) & MsgText(46), , MsgText(5)
         Text4.SetFocus
         Exit Function
      End If
      '金額不可為0...
      If Val(Text4) = 0 And Val(Text5) = 0 Then
         MsgBox MsgText(58) & MsgText(46), , MsgText(5)
         Text4.SetFocus
         Exit Function
      End If
      'Add by Amy 2020/08/12 部門
      Call Text16_Validate(bCancel)
      If bCancel = True Then
         TextInverse Text16
         Text16.SetFocus
         Exit Function
      End If
      'Add by Morgan 2007/10/2 檢查科目部門&智權人員是否正確
      'Modify by Amy 2021/03/08 +傳票號
      intI = PUB_AccNoGood(Text14, Text16, Text8, , Text2)
      If intI <> 0 Then
         If intI = 1 Then
            Text14.SetFocus
         ElseIf intI = 2 Then
            Text16.SetFocus
         ElseIf intI = 3 Then
            Text8.SetFocus
         End If
         Exit Function
      End If
      'end 2007/10/2
      '對沖(本所案號)
      Call Text10_Validate(bCancel)
      If bCancel = True Then
         TextInverse Text10
         Text10.SetFocus
         Exit Function
      End If
      'Add by Amy 2024/08/05 避免有未檢查到,從Text10_LostFocus搬過來
      If Text10 <> MsgText(601) Then
         '以案號檢查會計科目 (strNation於Text10_Validate設定)
         If ChkCaseNoAndAccNo(strNation) = False Then
            stMsg = Mid(Text10, 1, Len(Text10) - 9) & " " & GetNationName(strNation) & " 案 不可使用【" & Text14 & "】"
            MsgBox stMsg, , "輸入錯誤!!"
            TextInverse Text14
            Text14.SetFocus
            Exit Function
         End If
      End If
      'end 2024/08/05
      '對沖(客)
      Call Text7_Validate(bCancel)
      If bCancel = True Then
         TextInverse Text7
         Text7.SetFocus
         Exit Function
      End If
      '對沖(業)
      Call Text8_Validate(bCancel)
      If bCancel = True Then
         TextInverse Text8
         Text8.SetFocus
         Exit Function
      End If
      'Add by Amy 2018/06/06 41字頭及7121且st15是S部門不可輸小數,避免智權人員實績與結餘輸入因點數四捨五入後,導致智權人員實績與結餘分析表出現負數
      'Modify by Amy 2024/08/05 改4字頭
      If Text8 <> MsgText(601) Then
         If (Left(Text14, 1) = "4" Or Text14 = "7121") And Left(GetST15(Text8), 1) = "S" Then
            If Text4 <> MsgText(601) And InStr(Text4, ".") > 0 Then
               MsgBox IIf(Left(Text14, 1) = "4", "4字頭", "7121") & "科目不可輸小數！", vbExclamation, "資料錯誤"
               TextInverse Text4
               Text4.SetFocus
               Exit Function
            End If
            If Text5 <> MsgText(601) And InStr(Text5, ".") > 0 Then
               MsgBox IIf(Left(Text14, 1) = "4", "4字頭", "7121") & "科目不可輸小數！", vbExclamation, "資料錯誤"
               TextInverse Text5
               Text5.SetFocus
               Exit Function
            End If
         End If
      End If
      'end 2018/06/06
      '對沖傳票號
      Call Text9_Validate(bCancel)
      If bCancel = True Then
         TextInverse Text9
         Text9.SetFocus
         Exit Function
      End If
   '*** 其他提醒 ***
      'Add by Amy 2014/11/17 +公司別與特殊出名公司不同彈訊息提醒但可存-瑞婷
      'Modify by Amy 2023/04/07 排除ACS分潤進入者不需彈
      If UCase(Me.Tag) <> "FRMACC41L0" Then
         If Trim(Text10) <> MsgText(601) And Text1 <> strSpecComp Then
            nResponse = MsgBox("特殊出名公司與傳票公司別不符是否存檔?", vbOKCancel + vbDefaultButton2, MsgText(5))
            If nResponse = vbCancel Then
               Exit Function
            End If
         End If
      End If
      'end 2014/11/17
      
   End If 'm_State = "Ins"
   
   '確定 鈕
   If m_State = "F9" Then
      '合計(借貸平衡)確認
      If CreDebCheck <> MsgText(602) Or Val(Text11) = 0 Or Val(Text12) = 0 Then
         MsgBox MsgText(11), , MsgText(5)
         Exit Function
      End If
      'Add by Amy 2018/03/27 檢查同一張傳票若借貸都有2491XX的科目時,[不能]有4XXX的貸方
      'Memo 若未剔除4字頭,會[智權人員結餘點數查詢](frmacc4270) 1放出 資料可能抓的不正確 ex:1 公司 D106092547 需剔除
      If Chk2491InCome = True Then
         MsgBox "結餘放收入及結餘調整人員不可做在同一傳票！", , MsgText(5)
         Exit Function
      End If
      'Add by Amy 2023/06/14 避免從傳票輸入key 隱藏版傳票資料輸錯,導致期末結餘資料有誤,故彈訊息提醒ex:1公司D112052529
      'Memo                                   因結餘保留放出產生傳票(frmacc41f0) 1放出,也會產生同樣借貸科目,且財務也可能做調整傳票(借:24910x 貸:4字頭 1120614 問瑞婷),故不可鎖住
      If Chk24910xAnd41xxxx = True Then
         If MsgBox("此為「非當月結餘轉撥傳票產生(隱藏版人員)」所用之傳票科目" & vbCrLf & _
                              "確定[不是]要在隱藏版輸,按「是」繼續存檔,按「否」不存回前畫面", vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Function
         End If
      End If
      'Add by Amy 2022/05/16 結餘轉撥傳票存檔時需檢查每個人的總額必須與SalesPoint相符
      If UCase(Me.Tag) = "FRMACC41H0" Then
         If ChkSalesPointVal = True Then
            Exit Function
         End If
      End If
      'Add by Amy 2014/07/16 J公司傳票借方為進項稅額時,判斷A1P04及進項發票無資料時不可存檔
      If Text1 = "J" Then
          If CheckIs1211(Text1, Text2) = True Then
              If CheckExistA1p04(Text1, Text2) = False Then
                  If ExistCheck("Acc450", "A4501", Text2, "", False) = False Then
                      MsgBox "有進項稅額科目, 請輸入發票明細 !", , MsgText(5)
                      Exit Function
                  End If
              End If
          End If
      End If
      'end 2014/07/16
   End If 'm_State = "F9"

   ChkForm = True
End Function

Public Sub SaveData(ByVal m_State As String)
   Dim stTP As String
   
   Select Case m_State
      Case "Ins"
         Call Frmacc4120_Save
         Call acc021Save
      Case "Save" '確定
         If strSaveConfirm = MsgText(4) Then
            Frmacc4120_Save
         End If
         '新增後直接修改(貸方金額改借方金額)及傳票日.tag 要更新
         AdodcRefresh
         MaskEdBox1.Tag = Val(FCDate(MaskEdBox1))
         FormDisabled
         Text2.Enabled = True
         Text2.SetFocus
         'Modify by Amy 2024/08/05  從Form_Unload搬過來 原'Add by Amy 2021/09/23 若有修改發mail給電腦中心
         If HasUpdTag = True Then
              PUB_SendMail strUserNum, "A2004", "", "智權點數開放後傳票有修改「結餘」資料", strContent
              strContent = "": HasUpdTag = False
         End If
      Case "F5" '刪除
         Frmacc4120_Delete
         Frmacc4120_Clear
         adoacc020.MoveFirst
         If adoacc020.BOF = True Then
            Text2 = ""
         Else
            stTP = adoacc020.Fields("a0202")
            Acc020Refresh (stTP)
            Frmacc4120_Last
            If adoacc020.RecordCount = 0 Then
               Text2 = ""
               AdodcRefresh
            End If
         End If
   End Select
End Sub

'依案號判斷會計科目是否符合(從Text10_LostFocus搬過來修改)
Private Function ChkCaseNoAndAccNo(ByVal strNation As String) As Boolean
'Memo by Amy 2024/08/05 拿掉 ChkSCaseAccNO,可能原寫於Text10_LostFocus 直接按Inset不會檢查
'    原 Add by Amy 2022/09/21 ChkSCaseAccNO函數,判斷科目是2201開頭且S案,申請國家為台灣000者, 科目必須為220103,非台灣者必須為220105

   ChkCaseNoAndAccNo = False
   
   Select Case Text14
      Case "220101"
         If (Mid(Text10, 1, Len(Text10) - 9) = "T" Or Mid(Text10, 1, Len(Text10) - 9) = "TB" Or Mid(Text10, 1, Len(Text10) - 9) = "TS" _
            Or Mid(Text10, 1, Len(Text10) - 9) = "TD" Or Mid(Text10, 1, Len(Text10) - 9) = "TM" Or Mid(Text10, 1, Len(Text10) - 9) = "TR" _
            Or Mid(Text10, 1, Len(Text10) - 9) = "TT") And strNation = "000" Then
         Else
            Exit Function
         End If
      Case "220102"
         If (Mid(Text10, 1, Len(Text10) - 9) = "P" Or Mid(Text10, 1, Len(Text10) - 9) = "PS") And strNation = "000" Then
         Else
            Exit Function
         End If
      Case "220103"
         If (Mid(Text10, 1, Len(Text10) - 9) = "FCT" Or Mid(Text10, 1, Len(Text10) - 9) = "S") And strNation = "000" Then
         Else
            Exit Function
         End If
      Case "220104"
         If Mid(Text10, 1, Len(Text10) - 9) = "FCP" Or Mid(Text10, 1, Len(Text10) - 9) = "FG" Then
         Else
            Exit Function
         End If
      Case "220105"
         If (Mid(Text10, 1, Len(Text10) - 9) = "CFT" Or Mid(Text10, 1, Len(Text10) - 9) = "CFC" Or Mid(Text10, 1, Len(Text10) - 9) = "S") _
            And strNation <> "000" Then
         Else
            Exit Function
         End If
      Case "220106"
         If Mid(Text10, 1, Len(Text10) - 9) = "CFP" Or Mid(Text10, 1, Len(Text10) - 9) = "FCL" Or Mid(Text10, 1, Len(Text10) - 9) = "LIN" _
            Or Mid(Text10, 1, Len(Text10) - 9) = "CFL" Or Mid(Text10, 1, Len(Text10) - 9) = "CPS" Or Mid(Text10, 1, Len(Text10) - 9) = "L" Then
            'L案必須為[非]台灣
            If Mid(Text10, 1, Len(Text10) - 9) = "L" And strNation = "000" Then
               Exit Function
            End If
         Else
            Exit Function
         End If
      Case "220107"
         If Mid(Text10, 1, Len(Text10) - 9) = "TC" And strNation = "000" Then
         Else
            Exit Function
         End If
      Case "220108"
         If (Mid(Text10, 1, Len(Text10) - 9) = "P" Or Mid(Text10, 1, Len(Text10) - 9) = "PS") Then
         Else
            Exit Function
         End If
      Case "220111"
         If (Mid(Text10, 1, Len(Text10) - 9) = "T" Or Mid(Text10, 1, Len(Text10) - 9) = "TS" Or Mid(Text10, 1, Len(Text10) - 9) = "TF" _
            Or Mid(Text10, 1, Len(Text10) - 9) = "TC" Or Mid(Text10, 1, Len(Text10) - 9) = "TD" Or Mid(Text10, 1, Len(Text10) - 9) = "TM" _
            Or Mid(Text10, 1, Len(Text10) - 9) = "TT" Or Mid(Text10, 1, Len(Text10) - 9) = "TB" Or Mid(Text10, 1, Len(Text10) - 9) = "TR") _
            And strNation <> "000" Then
         Else
            Exit Function
         End If
      Case "220112"
         If (Mid(Text10, 1, Len(Text10) - 9) = "P" Or Mid(Text10, 1, Len(Text10) - 9) = "PS") And strNation <> "000" Then
         Else
            Exit Function
         End If
      Case "220113"
         If Mid(Text10, 1, Len(Text10) - 9) = "L" Or Mid(Text10, 1, Len(Text10) - 9) = "LA" Or Mid(Text10, 1, Len(Text10) - 9) = "FCL" Or Mid(Text10, 1, Len(Text10) - 9) = "LIN" Then
            'L案必須為台灣
            If Mid(Text10, 1, Len(Text10) - 9) = "L" And strNation <> "000" Then
               Exit Function
            End If
         Else
            '因財務處2024/3/26要求：法律所案案源案件之專業部門所提列(T, P, FCT, FCP)的出庭費, 想統一科目改成 220113
            If (Mid(Text10, 1, Len(Text10) - 9) = "P" Or Mid(Text10, 1, Len(Text10) - 9) = "T" Or Mid(Text10, 1, Len(Text10) - 9) = "FCP" Or Mid(Text10, 1, Len(Text10) - 9) = "FCT") _
               And strNation = "000" Then
            Else
               Exit Function
            End If
         End If
      Case "610103"
         If Mid(Text10, 1, Len(Text10) - 9) = "L" Or Mid(Text10, 1, Len(Text10) - 9) = "LA" Or Mid(Text10, 1, Len(Text10) - 9) = "FCL" _
            Or Mid(Text10, 1, Len(Text10) - 9) = "LIN" Or Mid(Text10, 1, Len(Text10) - 9) = "CFL" Then
         Else
            Exit Function
         End If
      Case Else
   End Select
      
   ChkCaseNoAndAccNo = True
End Function
'end 2024/08/05
