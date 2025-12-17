VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc41c0 
   AutoRedraw      =   -1  'True
   Caption         =   "傳票過帳後摘要修改"
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
      Top             =   3924
      Width           =   765
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "發票明細"
      Height          =   300
      Left            =   252
      TabIndex        =   41
      Top             =   2400
      Width           =   1200
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
      Left            =   7115
      TabIndex        =   40
      Top             =   84
      Width           =   1335
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
      Left            =   8010
      Picture         =   "Frmacc41c0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   23
      ToolTipText     =   "取消"
      Top             =   3924
      Visible         =   0   'False
      Width           =   550
   End
   Begin VB.TextBox Text1 
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
      Left            =   1308
      MaxLength       =   1
      TabIndex        =   0
      Top             =   84
      Width           =   612
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
      Top             =   444
      Width           =   1572
   End
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
      Height          =   300
      Left            =   240
      TabIndex        =   22
      Top             =   3180
      Width           =   492
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '靠右對齊
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
      Left            =   3348
      Locked          =   -1  'True
      MaxLength       =   14
      TabIndex        =   5
      Top             =   3180
      Width           =   1572
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  '靠右對齊
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
      Left            =   5040
      Locked          =   -1  'True
      MaxLength       =   14
      TabIndex        =   6
      Top             =   3180
      Width           =   1572
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
      TabIndex        =   21
      Top             =   2340
      Width           =   1500
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
      Left            =   5700
      TabIndex        =   20
      Top             =   2340
      Width           =   1488
   End
   Begin VB.TextBox Text13 
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
      Left            =   1920
      TabIndex        =   19
      Top             =   84
      Width           =   5100
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
      Left            =   2400
      MaxLength       =   9
      TabIndex        =   10
      Top             =   3900
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
      Height          =   300
      Left            =   5616
      MaxLength       =   10
      TabIndex        =   13
      Top             =   4284
      Width           =   1572
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
      Top             =   3540
      Width           =   1572
   End
   Begin VB.TextBox Text14 
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
      Height          =   300
      Left            =   840
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   4
      Top             =   3180
      Width           =   972
   End
   Begin VB.TextBox Text16 
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
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   7
      Top             =   3180
      Width           =   612
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
      Height          =   300
      Left            =   7320
      TabIndex        =   17
      Top             =   3180
      Width           =   1212
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
      Left            =   7425
      Picture         =   "Frmacc41c0.frx":066A
      Style           =   1  '圖片外觀
      TabIndex        =   16
      ToolTipText     =   "清除畫面"
      Top             =   3924
      Visible         =   0   'False
      Width           =   550
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   8100
      Picture         =   "Frmacc41c0.frx":0F34
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   444
      Width           =   350
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
      Left            =   2412
      MaxLength       =   1
      TabIndex        =   14
      Top             =   4596
      Visible         =   0   'False
      Width           =   540
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1296
      TabIndex        =   1
      Top             =   444
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
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
      Bindings        =   "Frmacc41c0.frx":1036
      Height          =   1500
      Left            =   252
      TabIndex        =   15
      Top             =   792
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
            ColumnWidth     =   1692.284
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1476.284
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1404.284
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
      Top             =   684
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   550
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
      TabIndex        =   12
      Top             =   4272
      Width           =   1572
      VariousPropertyBits=   679493659
      MaxLength       =   9
      Size            =   "10927;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   312
      Left            =   4776
      TabIndex        =   9
      Top             =   3540
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
   Begin MSForms.TextBox Text19 
      Height          =   315
      Left            =   6390
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   3930
      Width           =   960
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
   Begin MSForms.TextBox Text15 
      Height          =   300
      Left            =   1830
      TabIndex        =   18
      Top             =   3180
      Width           =   1455
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
      TabIndex        =   39
      Top             =   444
      Width           =   972
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
      Height          =   252
      Left            =   5520
      TabIndex        =   38
      Top             =   444
      Width           =   972
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
      TabIndex        =   37
      Top             =   84
      Width           =   852
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
      TabIndex        =   36
      Top             =   2940
      Width           =   492
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
      Height          =   252
      Left            =   1560
      TabIndex        =   35
      Top             =   2940
      Width           =   972
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
      Height          =   252
      Left            =   3600
      TabIndex        =   34
      Top             =   2940
      Width           =   1092
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
      Height          =   252
      Left            =   5280
      TabIndex        =   33
      Top             =   2940
      Width           =   1092
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
      Height          =   252
      Left            =   7200
      TabIndex        =   32
      Top             =   2940
      Width           =   852
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
      Height          =   252
      Left            =   4188
      TabIndex        =   31
      Top             =   3576
      Width           =   612
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   2196
      Left            =   120
      Top             =   2832
      Width           =   8532
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4740
      Visible         =   0   'False
      Width           =   132
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
      TabIndex        =   30
      Top             =   2340
      Width           =   732
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
      TabIndex        =   29
      Top             =   3900
      Width           =   1452
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
      Height          =   252
      Left            =   4176
      TabIndex        =   28
      Top             =   3924
      Width           =   1452
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
      Height          =   252
      Left            =   4176
      TabIndex        =   27
      Top             =   4284
      Width           =   1452
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
      TabIndex        =   26
      Top             =   3540
      Width           =   2172
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
      TabIndex        =   25
      Top             =   4296
      Width           =   1632
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
      TabIndex        =   24
      Top             =   4644
      Visible         =   0   'False
      Width           =   1632
   End
End
Attribute VB_Name = "Frmacc41c0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/25 Form2.0已修改 Text15/Text19/Combo1/DataGrid1
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
Dim bolFirst As Boolean 'Added by Lydia 2025/08/08 True = Focus在摘要Combo，在按下Insert鍵時先重送Insert鍵於第2次才執行更新明細

'Modify by Amy 2014/02/20
Private Sub CmdChgComp_Click()
    Dim strNewComp As String
    
    Frmacc41c2.Show vbModal
    strNewComp = strCompanyNo
    If strNewComp <> "" And strNewComp <> Me.Text1 Then
        Text1 = strNewComp
        Text2 = MsgText(601)
        Text11 = MsgText(601)
        Text12 = MsgText(601)
        AdodcClear
        Acc020Refresh
        AdodcRefresh
    End If
    strCompanyNo = MsgText(601)
End Sub
'end 2014/02/20

'Add by Amy 2014/07/09
Private Sub cmdDetail_Click()
    Frmacc1172.strBackForm = "Frmacc41c0"  '記錄返回畫面
    Frmacc1172.Show
    Screen.MousePointer = vbDefault
    Me.Hide
End Sub
'end 2014/07/09

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

Private Sub Command1_Click()
   If Adodc1.Recordset.RecordCount <> 0 Then
      If Not IsNull(Adodc1.Recordset.Fields("ax210").Value) Then
         MsgBox MsgText(14), , MsgText(5)
         Text14.SetFocus
         Exit Sub
      End If
      adoTaie.Execute "delete from acc021 where ax201 = '" & Adodc1.Recordset.Fields("ax201").Value & "' and ax202 = '" & Adodc1.Recordset.Fields("ax202").Value & "' and ax203 = '" & Adodc1.Recordset.Fields("ax203").Value & "'"
      AdodcRefresh
      AdodcClear
      SumShow
   End If
End Sub

Private Sub Command2_Click()
Dim adoaccmax As New ADODB.Recordset

   AdodcClear
   adoaccmax.CursorLocation = adUseClient
   adoaccmax.Open "select max(ax203) from acc021 where ax201 = '" & Text1 & "' and ax202 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
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
    'Add by Amy 2014/02/19 由text2_validate搬過來
    If Text2 = MsgText(601) Then
      MsgBox Label2 & MsgText(52), , MsgText(5)
      Text2.SetFocus
      Exit Sub
   End If
   'end 2014/02/19
   
   Acc020Refresh
   If adoacc020.RecordCount <> 0 Then
      FormShow
      AdodcRefresh
      SumShow
      RecordShow
   End If
   AdodcClear
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
   Acc020Refresh
   If adoacc020.RecordCount <> 0 Then
      FormShow
      AdodcRefresh
      SumShow
      RecordShow
   End If
   strCompanyNo = MsgText(601)
   strItemNo = MsgText(601)
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
   'Modify by Amy 2014/02/20
   'Text1 = MsgText(601)
   Text1 = strCompanyNo
   strCompanyNo = MsgText(601)
   'end 2014/02/20
   Text2 = MsgText(601)
   MaskEdBox1.Mask = DFormat
   OpenTable
   If adoacc020.RecordCount <> 0 Then
      RecordShow
   End If
'   FormDisabled
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
   Set Frmacc41c0 = Nothing
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
   adoacc020.MaxRecords = intMax
   Select Case strAccount
      Case "2"
         adoacc020.Open "select a0301 as a0201, a0302 as a0202, a0305 as a0205, a0306 as a0206, a0307 as a0207, a0308 as a0208, a0309 as a0209, a0310 as a0210, a0311 as a0211 from acc030 order by a0201 asc, a0202 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      Case Else
         'Modify by Amy 2014/02/19 +公司別
         adoacc020.Open "select * from acc020 where a0201='" & Text1 & "' And a0202 >= '" & Text2 & "' order by a0201 asc, a0202 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   End Select
   adoacc021.CursorLocation = adUseClient
   Select Case strAccount
      Case "2"
         adoacc021.Open "select ax301 as ax201, ax302 as ax202, ax303 as ax203, ax304 as ax204, ax305 as ax205, ax306 as ax206, ax307 as ax207, ax308 as ax208, ax309 as ax209, ax310 as ax210, ax311 as ax211, ax312 as ax212, ax314 as ax214, ax313 as ax213 from acc031 where ax301 = '" & Text1 & "' and ax302 = '" & Text2 & "' and ax303 = '" & Text3 & "' order by ax201 asc, ax202 asc, ax203 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      Case Else
         adoacc021.Open "select * from acc021 where ax201 = '" & Text1 & "' and ax202 = '" & Text2 & "' and ax203 = '" & Text3 & "' order by ax201 asc, ax202 asc, ax203 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   End Select
   adoadodc1.CursorLocation = adUseClient
   Select Case strAccount
      Case "2"
         adoadodc1.Open "select ax301 as ax201, ax302 as ax202, ax303 as ax203, ax304 as ax204, ax305 as ax205, ax306 as ax206, ax307 as ax207, ax308 as ax208, ax309 as ax209, ax310 as ax210, ax311 as ax211, ax312 as ax212, ax314 as ax214, ax313 as ax213, a0102 from acc031, acc010 where ax305 = a0101 (+) and ax301 = '" & Text1 & "' and ax302 = '" & Text2 & "' order by ax201 asc, ax202 asc, ax203 asc", adoTaie, adOpenStatic, adLockReadOnly
      Case Else
         adoadodc1.Open "select * from acc021, acc010 where ax205 = a0101 (+) and ax201 = '" & Text1 & "' and ax202 = '" & Text2 & "' order by ax201 asc, ax202 asc, ax203 asc", adoTaie, adOpenStatic, adLockReadOnly
   End Select
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
Public Sub acc021Save()
Dim strCombo1 As String

On Error GoTo Checking
   'Add by Amy 2014/02/19
   If Text2 = MsgText(601) Then
      MsgBox Label2 & MsgText(52), , MsgText(5)
      strControlButton = MsgText(602)
      Text2.SetFocus
      Exit Sub
   End If
   'end 2014/02/19
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
      'Add by Amy 2021/10/25 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
      If PUB_ChkUniText(Me) = False Then
         strControlButton = MsgText(602)
         Exit Sub
      End If

'cancel by sonia 2016/5/20 不能改部門不必檢查,否則有錯時setfocus會有誤
'      If CheckDept(Text14, Text16) = False Then
'         MsgBox MsgText(103), , MsgText(5)
'         strControlButton = MsgText(602)
'         Text16.SetFocus
'         Exit Sub
'      End If
'end 2016/5/20
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
         '2009/1/17 modify by sonia 加可輸代理人編號
         'adocheck.Open "select cu01 as Name from customer where cu01 = '" & Mid(Text7, 1, 8) & "' union " & _
                       "select a0i01 as Name from acc0i0 where a0i01 = '" & Text7 & "' union " & _
                       "select st01 as Name from staff where st01 = '" & Text7 & "'", adoTaie, adOpenStatic, adLockReadOnly
         adocheck.Open "select cu01 as Name from customer where cu01 = '" & Mid(Text7, 1, 8) & "' union " & _
                       "select a0i01 as Name from acc0i0 where a0i01 = '" & Text7 & "' union " & _
                       "select st01 as Name from staff where st01 = '" & Text7 & "' union " & _
                       "select fa01 as Name from fagent where fa01 = '" & Mid(Text7, 1, 8) & "'", adoTaie, adOpenStatic, adLockReadOnly
         '2009/1/17 end
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
'         If Len(Mid(Text10, 2, Len(Text10) - 1)) > 6 Then
            '2006/10/19 MODIFY BY SONIA 改抓案件基本檔
            'adocheck.Open "select cp09 from caseprogress where cp01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and cp02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and cp03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and cp04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
            Select Case Mid(Text10, 1, Len(Text10) - 9)
               Case "P", "CFP", "FCP"
                  adocheck.Open "select * from PATENT where PA01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and PA02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and PA03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and PA04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
               Case "T", "CFT", "FCT", "TF"
                  adocheck.Open "select * from TRADEMARK where TM01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and TM02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and TM03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and TM04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
               'modify by sonia 2019/8/14 +ACS,LIN系統類別
               Case "L", "CFL", "FCL", "LIN", "ACS"
                  adocheck.Open "select * from LAWCASE where LC01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and LC02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and LC03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and LC04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
               Case "LA"
                  adocheck.Open "select * from HIRECASE where HC01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and HC02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and HC03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and HC04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
               Case Else
                  adocheck.Open "select * from SERVICEPRACTICE where SP01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and SP02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and SP03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and SP04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
            End Select
            '2006/10/19 END
            If adocheck.RecordCount = 0 Then
               MessageShow Label14
               strControlButton = MsgText(602)
               adocheck.Close
               Text10.SetFocus
               Exit Sub
            End If
            adocheck.Close
'         Else
'            MessageShow Label14
'            strControlButton = MsgText(602)
'            Text10.SetFocus
'            Exit Sub
'         End If
      End If
      If Text9 <> MsgText(601) Then
         adocheck.CursorLocation = adUseClient
         adocheck.Open "select a0201, a0202 from acc020 where a0201 = '" & Text1 & "' and a0202 = '" & Text9 & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adocheck.RecordCount = 0 Then
            MessageShow Label13
            strControlButton = MsgText(602)
            adocheck.Close
            Text9.SetFocus
            Exit Sub
         End If
         adocheck.Close
      End If
   End If
   adoacc021.Close
   adoacc021.CursorLocation = adUseClient
   Select Case strAccount
      Case "2"
         adoacc021.Open "select ax301 as ax201, ax302 as ax202, ax303 as ax203, ax304 as ax204, ax305 as ax205, ax306 as ax206, ax307 as ax207, ax308 as ax208, ax309 as ax209, ax310 as ax210, ax311 as ax211, ax312 as ax212, ax314 as ax214, ax313 as ax213 from acc031 where ax301 = '" & Text1 & "' and ax302 = '" & Text2 & "' and ax303 = '" & Text3 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      Case Else
         adoacc021.Open "select * from acc021 where ax201 = '" & Text1 & "' and ax202 = '" & Text2 & "' and ax203 = '" & Text3 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   End Select
   If adoacc021.RecordCount = 0 Then
      adoacc021.AddNew
   End If
   adoacc021.Fields("ax201").Value = Text1
   adoacc021.Fields("ax202").Value = Text2
   adoacc021.Fields("ax203").Value = Text3
   If Text14 <> MsgText(601) Then
      adoacc021.Fields("ax205").Value = Text14
   Else
      adoacc021.Fields("ax205").Value = Null
   End If
   If Text4 <> MsgText(601) Then
      adoacc021.Fields("ax206").Value = Val(Text4)
   Else
      adoacc021.Fields("ax206").Value = 0
   End If
   If Text5 <> MsgText(601) Then
      adoacc021.Fields("ax207").Value = Val(Text5)
   Else
      adoacc021.Fields("ax207").Value = 0
   End If
   If Text16 <> MsgText(601) Then
      adoacc021.Fields("ax204").Value = Text16
   Else
      adoacc021.Fields("ax204").Value = MsgText(55)
   End If
    'Modify By Cheng 2004/03/12
    '避免單引號存檔錯誤
'   adoTaie.Execute "update acc1p0 set a1p14 = " & IIf(Combo1 = "", "null", "'" & Combo1 & "'") & ", a1p15 = " & IIf(Text7 = "", "null", "'" & Text7 & "'") & ", a1p16 = " & IIf(Text8 = "", "null", "'" & Text8 & "'") & ", a1p17 = " & IIf(Text10 = "", "null", "'" & Text10 & "'") & ", a1p30 = " & IIf(Text6 = "", "null", "'" & Text6 & "'") & ", a1p31 = " & IIf(Text18 = "", "null", "'" & Text18 & "'") & " where a1p01 = '" & Text1 & "' and a1p03 = '" & Text3 & "' and a1p22 = '" & Text2 & "'"
   'Modify by Amy 2014/02/19 取消作帳公司
   'adoTaie.Execute "update acc1p0 set a1p14 = " & IIf(Combo1 = "", "null", "'" & ChgSQL(Combo1) & "'") & ", a1p15 = " & IIf(Text7 = "", "null", "'" & Text7 & "'") & ", a1p16 = " & IIf(Text8 = "", "null", "'" & Text8 & "'") & ", a1p17 = " & IIf(Text10 = "", "null", "'" & Text10 & "'") & ", a1p30 = " & IIf(Text6 = "", "null", "'" & Text6 & "'") & ", a1p31 = " & IIf(Text18 = "", "null", "'" & Text18 & "'") & " where a1p01 = '" & Text1 & "' and a1p03 = '" & Text3 & "' and a1p22 = '" & Text2 & "'"
   adoTaie.Execute "update acc1p0 set a1p14 = " & IIf(Combo1 = "", "null", "'" & ChgSQL(Combo1) & "'") & ", a1p15 = " & IIf(Text7 = "", "null", "'" & Text7 & "'") & ", a1p16 = " & IIf(Text8 = "", "null", "'" & Text8 & "'") & ", a1p17 = " & IIf(Text10 = "", "null", "'" & Text10 & "'") & ", a1p30 = " & IIf(Text6 = "", "null", "'" & Text6 & "'") & " where a1p01 = '" & Text1 & "' and a1p03 = '" & Text3 & "' and a1p22 = '" & Text2 & "'"
    'End
   If Combo1 <> MsgText(601) Then
      adoacc021.Fields("ax212").Value = Replace(Combo1, "'", "''")
      strCombo1 = Combo1
      Combo1.Clear
      Combo1.AddItem strCombo1
   Else
      adoacc021.Fields("ax212").Value = Null
   End If
   If Text7 <> MsgText(601) Then
      adoacc021.Fields("ax208").Value = Text7
   Else
      adoacc021.Fields("ax208").Value = Null
   End If
   If Text8 <> MsgText(601) Then
      adoacc021.Fields("ax209").Value = Text8
   Else
      adoacc021.Fields("ax209").Value = Null
   End If
   If Text9 <> MsgText(601) Then
      adoacc021.Fields("ax211").Value = Text9
   Else
      adoacc021.Fields("ax211").Value = Null
   End If
   If Text6 <> MsgText(601) Then
      adoacc021.Fields("ax213").Value = Text6
   Else
      adoacc021.Fields("ax213").Value = Null
   End If
   Text10 = CaseNoZero(Text10)
   If Text10 <> MsgText(601) Then
      adoacc021.Fields("ax214").Value = Text10
   Else
      adoacc021.Fields("ax214").Value = Null
   End If
   'Modify by Amy 2013/12/12 Mark 取消作帳公司欄位
'   If Text18 <> MsgText(601) Then
'      adoacc021.Fields("ax215").Value = Text18
'   Else
'      adoacc021.Fields("ax215").Value = Null
'   End If
   'end 2013/12/12
   '2008/7/17 add by sonia
   If Text6 <> MsgText(601) Then
      adoacc021.Fields("ax213").Value = Text6
   Else
      adoacc021.Fields("ax213").Value = Null
   End If
   '2008/7/17 end
   adoacc021.UpdateBatch
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
   If IsNull(Adodc1.Recordset.Fields("ax203").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = Adodc1.Recordset.Fields("ax203").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("ax205").Value) Then
      Text14 = MsgText(601)
   Else
      Text14 = Adodc1.Recordset.Fields("ax205").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("ax206").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = Adodc1.Recordset.Fields("ax206").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("ax207").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = Adodc1.Recordset.Fields("ax207").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("ax204").Value) Then
      Text16 = MsgText(601)
      Text17 = MsgText(601)
   Else
      If Adodc1.Recordset.Fields("ax204").Value = MsgText(55) Then
         Text16 = MsgText(601)
         Text17 = MsgText(601)
      Else
         Text16 = Adodc1.Recordset.Fields("ax204").Value
      End If
   End If
   If IsNull(Adodc1.Recordset.Fields("ax212").Value) Then
      Combo1 = MsgText(601)
   Else
      Combo1 = Adodc1.Recordset.Fields("ax212").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("ax208").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = Adodc1.Recordset.Fields("ax208").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("ax209").Value) Then
      Text8 = MsgText(601)
      Text19 = "" 'Add by Amy 2017/09/14
   Else
      Text8 = Adodc1.Recordset.Fields("ax209").Value
      Text19 = StaffQuery(Text8) 'Add by Amy 2017/09/14
   End If
   If IsNull(Adodc1.Recordset.Fields("ax211").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = Adodc1.Recordset.Fields("ax211").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("ax213").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = Adodc1.Recordset.Fields("ax213").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("ax214").Value) Then
      Text10 = MsgText(601)
   Else
      Text10 = Adodc1.Recordset.Fields("ax214").Value
   End If
   'Modify by Amy 2013/12/12 Mark 取消作帳公司欄位
'   If IsNull(Adodc1.Recordset.Fields("ax215").Value) Then
'      Text18 = MsgText(601)
'   Else
'      Text18 = Adodc1.Recordset.Fields("ax215").Value
'   End If
   'end 2013/12/12
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

Private Sub Text10_Validate(Cancel As Boolean)
On Error GoTo Checking
   If Text10 <> MsgText(601) Then
      Text10 = CaseNoZero(Text10)
      adocase.CursorLocation = adUseClient
      '2006/10/19 MODIFY BY SONIA 改抓案件基本檔
      'adocase.Open "select cp09 from caseprogress where cp01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and cp02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and cp03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and cp04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      'modify by sonia 2019/8/14 因ACS案件改寫法
      'Select Case Mid(Text10, 1, Len(Text10) - 9)
      '   Case "P", "CFP", "FCP"
      '      adocase.Open "select * from PATENT where PA01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and PA02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and PA03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and PA04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      '   Case "T", "CFT", "FCT", "TF"
      '      adocase.Open "select * from TRADEMARK where TM01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and TM02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and TM03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and TM04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      '   Case "L", "CFL", "FCL"
      '      adocase.Open "select * from LAWCASE where LC01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and LC02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and LC03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and LC04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      '   Case "LA"
      '      adocase.Open "select * from HIRECASE where HC01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and HC02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and HC03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and HC04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      '   Case Else
      '      adocase.Open "select * from SERVICEPRACTICE where SP01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and SP02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and SP03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and SP04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      'End Select
      adocase.Open "select pa01 as SystemNo,pa09,pa26  from patent where pa01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and pa02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and pa03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and pa04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
                   "select tm01 as SystemNo,tm10,tm23 from trademark where tm01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and tm02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and tm03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and tm04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
                   "select lc01 as SystemNo,lc15,lc11 from lawcase where lc01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and lc02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and lc03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and lc04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
                   "select hc01 as SystemNo,'000',hc07 from hirecase where hc01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and hc02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and hc03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and hc04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "' union " & _
                   "select sp01 as SystemNo,sp09,sp08 from servicepractice where sp01 = '" & Mid(Text10, 1, Len(Text10) - 9) & "' and sp02 = '" & Mid(Text10, Len(Text10) - 8, 6) & "' and sp03 = '" & Mid(Text10, Len(Text10) - 2, 1) & "' and sp04 = '" & Mid(Text10, Len(Text10) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      '2006/10/19 END
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
         Text8_Validate True
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
      Text8_Validate True
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
   CloseIme
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

Private Sub Text18_GotFocus()
   TextInverse Text18
   CloseIme
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Modify by Amy 2013/12/12 Mark 取消作帳公司欄位
'Private Sub Text18_GotFocus()
'   TextInverse Text18
'End Sub
'
'Private Sub Text18_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub
'end 2013/12/12

Private Sub Text2_GotFocus()
   TextInverse Text2
   CloseIme
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
   Text15 = ""
   Text4 = ""
   Text5 = ""
   Text16 = ""
   Text17 = ""
   Combo1 = ""
   Text7 = ""
   Text8 = ""
   Text19 = "" 'Add by Amy 2017/09/14
   Text9 = ""
   Text6 = ""
   Text10 = ""
   'Text18 = "" 'Modify by Amy 2013/12/12 Mark 取消作帳公司欄位
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
Dim nResponse
   'Add by Amy 2021/10/25
   Call PUB_SaveTrackMode(1, KeyCode)
    
    'Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
    If PUB_ChkTrackMode = False Then
        Exit Sub
    End If
    'end 2021/10/25
 
   
   Select Case KeyCode
      Case vbKeyInsert
         'Added by Lydia 2025/08/08 明細要按Insert才會更新資料，但是Form2.0元件支援Insert鍵會切換”新增/覆寫模式”
         If bolFirst = True Then  '在按下Insert鍵時先重送Insert鍵於第2次才執行更新明細
             Call PUB_SendSKey("KeyInsert")
             bolFirst = False
             Exit Sub
         End If
         'end 2025/08/08
         
'         If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
'            Exit Sub
'         End If
'         Frmacc4120_Save
         'Add by Amy 2014/07/16 J公司傳票借方為進項稅額時,判斷A1P04及進項發票無資料時不可存檔
         If Text1 = "J" Then
            If CheckIs1211(Text1, Text2) = True Then
                If CheckExistA1p04(Text1, Text2) = False Then
                    If ExistCheck("Acc450", "A4501", Text2, "", False) = False Then
                        MsgBox "有進項稅額科目, 請輸入發票明細 !", , MsgText(5)
                        Exit Sub
                    End If
                End If
            End If
         End If
        
        'add by sonia 2015/4/22 41XX(除4191,4192,4194)外或7121摘要有結餘,對沖其他欄也要有
        If (Left(Text14, 2) = "41" And Text14 <> "4191" And Text14 <> "4192" And Text14 <> "4194") Or Text14 = "7121" Then
           If InStr(Combo1, "結餘") > 0 And InStr(Text6, "結餘") = 0 Then
              MsgBox "收文科目摘要欄內有 結餘 字樣, 對沖代號(其它)欄也要輸結餘！", vbExclamation, "資料錯誤"
              TextInverse Text6
              Text6.SetFocus
              Exit Sub
           End If
        End If
        '2015/4/22 end
         
         'add by sonia 2019/9/5
         If Left(Text14, 1) = "4" And InStr(Combo1, "轉撥") > 0 Then
            nResponse = MsgBox("非轉撥傳票摘要不可輸入 轉撥 二字, 否則會影響實績點數, 是否存檔?", vbOKCancel + vbDefaultButton2, MsgText(5))
            If nResponse = vbCancel Then
               Exit Sub
            End If
         End If
         'end 2019/9/5
         
        'end 2014/07/16
         If strControlButton <> MsgText(602) Then
            acc021Save
         End If
         If strControlButton <> MsgText(602) Then
            AdodcClear
            If adocheck.State = adStateOpen Then
               adocheck.Close
            End If
            adocheck.CursorLocation = adUseClient
            Select Case strAccount
               Case "2"
                  adocheck.Open "select max(ax303) from acc031 where ax301 = '" & Text1 & "' and ax302 = '" & Text2 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
               Case Else
                  adocheck.Open "select max(ax203) from acc021 where ax201 = '" & Text1 & "' and ax202 = '" & Text2 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
            End Select
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
'            Text14.SetFocus
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

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

'Modify by Amy 2022/07/20 原:Integer
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
      'Add by Morgan 2007/3/1 八碼時要補'0'
      ElseIf Len(Text7) = 8 Then
         Text7 = Text7 & "0"
      'End 2007/3/1
      End If
      If ExistCheck("customer", "cu01", Mid(Text7, 1, 8), Label11, False) = False Then
         If ExistCheck("acc0i0", "a0i01", Text7, Label11, False) = False Then
            If ExistCheck("staff", "st01", Text7, Label11, False) = False Then
               'Modify by Morgan 2006/8/15 加檢查代理人擋
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
   If Text8 <> MsgText(601) Then
      'modify by sonia 員工已離職要提醒;不可輸入S29,否則實績結餘點數報表抓不到
      'If ExistCheck("staff", "st01", Text8, Label12) = False Then
      '   Cancel = True
      '   Exit Sub
      'End If
      If Text8 > "S" Then
         MsgBox "不可輸入S字頭的編號,否則實績結餘點數報表抓不到 !", , MsgText(5)
         Cancel = True
         TextInverse Text8
      ElseIf PUB_GetStaffState(Text8.Text, strExc(1), True) = 0 Then
         Cancel = True
         TextInverse Text8
      'Add by Amy 2017/09/14
      Else
        Text19.Text = strExc(1)
      End If
      'add by sonia 2021/1/28
      If SalesNoCheckAccNo(Text14, Text8) = False Then
      End If
      'end 2021/1/28
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
   Select Case strAccount
      Case "2"
         adoaccsum.Open "select sum(ax306), sum(ax307) from acc031 where ax301 = '" & Text1 & "' and ax302 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
      Case Else
         adoaccsum.Open "select sum(ax206), sum(ax207) from acc021 where ax201 = '" & Text1 & "' and ax202 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
   End Select
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
   'Text18.Enabled = False 'Modify by Amy 2013/12/12 Mark 取消作帳公司欄位
   Command1.Enabled = False
   Command2.Enabled = False
End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'
'*************************************************
Public Sub FormEnabled()
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
   'Text18.Enabled = True 'Modify by Amy 2013/12/12 Mark 取消作帳公司欄位
   Command1.Enabled = True
   Command2.Enabled = True
End Sub

'*************************************************
'  重新整理傳票資料
'
'*************************************************
'Modify by Morgan 2004/10/27 加參數iMov 0:ax202=Text2, 1:ax202>=Text2
Public Sub Acc020Refresh(Optional ByVal iMov As Integer = 0)
   Screen.MousePointer = vbHourglass
On Error GoTo Checking
   If adoacc020.State = adStateOpen Then
      adoacc020.Close
   End If
   adoacc020.CursorLocation = adUseClient
   adoacc020.MaxRecords = intMax
   Select Case strAccount
      Case "2"
         adoacc020.Open "select a0301 as a0201, a0302 as a0202, a0305 as a0205, a0306 as a0206, a0307 as a0207, a0308 as a0208, a0309 as a0209, a0310 as a0210, a0311 as a0211 from acc030 order by a0201 asc, a0202 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      Case Else
         '93.3.10 MODIFY BY SONIA
         'adoacc020.Open "select * from acc020 where a0202 >= '" & Text2 & "' order by a0201 asc, a0202 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         
         'Modify by Morgan 2004/10/27
         'adoacc020.Open "select * from acc020 where a0202 >= '" & Text2 & "' AND ROWNUM<11 order by a0201 asc, a0202 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         'Modify by Amy 2014/02/19 +公司別
         If iMov = 0 Then
            adoacc020.Open "select * from acc020 where a0201='" & Text1 & "' And a0202 >= '" & Text2 & "' AND ROWNUM<2", adoTaie, adOpenDynamic, adLockBatchOptimistic
         Else
            adoacc020.Open "select * from acc020 where a0201='" & Text1 & "' And a0202 >= '" & Text2 & "' order by a0201 asc, a0202 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         End If
         'end 2014/02/19
         '2004/10/27 end
         
         '93.3.10 END
   End Select
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
   Select Case strAccount
      Case "2"
         adoacc021.Open "select ax301 as ax201, ax302 as ax202, ax303 as ax203, ax304 as ax204, ax305 as ax205, ax306 as ax206, ax307 as ax207, ax308 as ax208, ax309 as ax209, ax310 as ax210, ax311 as ax211, ax312 as ax212, ax314 as ax214, ax313 as ax213 from acc031 where ax302 = '" & Text2 & "' and ax301 = '" & Text1 & "' and ax303 = '" & Text3 & "' order by ax201 asc, ax202 asc, ax203 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      Case Else
         adoacc021.Open "select * from acc021 where ax202 = '" & Text2 & "' and ax201 = '" & Text1 & "' and ax203 = '" & Text3 & "' order by ax201 asc, ax202 asc, ax203 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   End Select
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   Select Case strAccount
      Case "2"
         adoadodc1.Open "select ax301 as ax201, ax302 as ax202, ax303 as ax203, ax304 as ax204, ax305 as ax205, ax306 as ax206, ax307 as ax207, ax308 as ax208, ax309 as ax209, ax310 as ax210, ax311 as ax211, ax312 as ax212, ax314 as ax214, a0102, ax313 as ax213 from acc031, acc010 where ax305 = a0101 (+) and ax302 = '" & Text2 & "' and ax301 = '" & Text1 & "' order by ax201 asc, ax202 asc, ax203 asc", adoTaie, adOpenStatic, adLockReadOnly
      Case Else
         adoadodc1.Open "select * from acc021, acc010 where ax205 = a0101 (+) and ax202 = '" & Text2 & "' and ax201 = '" & Text1 & "' order by ax201 asc, ax202 asc, ax203 asc", adoTaie, adOpenStatic, adLockReadOnly
   End Select
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.Find "ax203 = '" & Text3 & "'", 0, adSearchForward, 1
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

Private Sub Text9_Validate(Cancel As Boolean)
   If Text9 <> MsgText(601) Then
      adocheck.CursorLocation = adUseClient
      'Modify by Amy 2014/02/19 改公司別 原:'1'
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


