VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc2110 
   AutoRedraw      =   -1  'True
   Caption         =   "收款作業"
   ClientHeight    =   5430
   ClientLeft      =   50
   ClientTop       =   280
   ClientWidth     =   8820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   8820
   Begin VB.TextBox Text19 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3456
      Width           =   390
   End
   Begin VB.TextBox Text17 
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
      Height          =   315
      Left            =   1290
      MaxLength       =   10
      TabIndex        =   6
      Top             =   4470
      Width           =   1455
   End
   Begin VB.TextBox Text16 
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
      Left            =   1290
      MaxLength       =   3
      TabIndex        =   14
      Top             =   4140
      Width           =   528
   End
   Begin VB.TextBox Text15 
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
      Left            =   1290
      TabIndex        =   19
      Top             =   4815
      Width           =   1350
   End
   Begin VB.TextBox Text14 
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
      Left            =   3930
      MaxLength       =   9
      TabIndex        =   17
      Top             =   4476
      Width           =   1700
   End
   Begin VB.TextBox Text13 
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
      Left            =   6690
      MaxLength       =   5
      TabIndex        =   16
      Top             =   4140
      Width           =   900
   End
   Begin VB.TextBox Text12 
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
      Left            =   3930
      MaxLength       =   12
      TabIndex        =   15
      Top             =   4140
      Width           =   1700
   End
   Begin VB.TextBox Text11 
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
      Left            =   6690
      MaxLength       =   10
      TabIndex        =   18
      Top             =   4476
      Width           =   1845
   End
   Begin VB.TextBox Text1 
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
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   8
      Top             =   3456
      Width           =   390
   End
   Begin VB.ComboBox Combo4 
      Height          =   300
      Left            =   6720
      TabIndex        =   3
      Top             =   97
      Width           =   1692
   End
   Begin VB.TextBox Text10 
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
      Left            =   6744
      TabIndex        =   38
      Top             =   3024
      Width           =   1575
   End
   Begin VB.TextBox Text20 
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
      Left            =   1704
      MaxLength       =   12
      TabIndex        =   35
      Top             =   3024
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   2640
      Picture         =   "Frmacc2110.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   97
      Width           =   350
   End
   Begin VB.TextBox Text5 
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
      Left            =   5064
      TabIndex        =   34
      Top             =   3024
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
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
      Height          =   330
      Left            =   3930
      TabIndex        =   10
      Top             =   3456
      Width           =   1572
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "Frmacc2110.frx":0102
      Left            =   6690
      List            =   "Frmacc2110.frx":0104
      TabIndex        =   13
      Top             =   3816
      Width           =   1845
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2110.frx":0106
      Height          =   1700
      Left            =   240
      TabIndex        =   22
      Top             =   1200
      Width           =   8295
      _ExtentX        =   14623
      _ExtentY        =   2981
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.5
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "a0102"
         Caption         =   "會計科目"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "a1p06"
         Caption         =   "部門"
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
         DataField       =   "a1p21"
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
      BeginProperty Column03 
         DataField       =   "a1p07"
         Caption         =   "借方金額(台幣)"
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
         Caption         =   "貸方金額(台幣)"
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
         DataField       =   "a1p23"
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
      BeginProperty Column06 
         DataField       =   "a1p24"
         Caption         =   "收款類別"
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
            ColumnWidth     =   2589.732
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   610.016
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1390.11
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1679.811
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1670.173
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1539.78
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "收款資料"
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
      Left            =   7320
      TabIndex        =   9
      Top             =   888
      Width           =   1212
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
      Height          =   350
      Left            =   264
      Picture         =   "Frmacc2110.frx":011B
      Style           =   1  '圖片外觀
      TabIndex        =   21
      ToolTipText     =   "取消"
      Top             =   2964
      Width           =   350
   End
   Begin VB.TextBox Text9 
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
      Left            =   3624
      TabIndex        =   32
      Top             =   3024
      Width           =   1335
   End
   Begin VB.TextBox Text8 
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
      Left            =   3930
      MaxLength       =   20
      TabIndex        =   12
      Top             =   3816
      Width           =   1700
   End
   Begin VB.TextBox Text7 
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
      Left            =   1290
      MaxLength       =   14
      TabIndex        =   11
      Top             =   3816
      Width           =   1350
   End
   Begin VB.TextBox Text3 
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
      Left            =   1290
      MaxLength       =   13
      TabIndex        =   4
      Top             =   420
      Width           =   1668
   End
   Begin VB.TextBox Text2 
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
      Height          =   315
      Left            =   1290
      MaxLength       =   15
      TabIndex        =   0
      Top             =   90
      Width           =   1335
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   4080
      TabIndex        =   2
      Top             =   97
      Width           =   1575
      _ExtentX        =   2787
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   1080
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
   Begin MSForms.TextBox Text18 
      Height          =   330
      Left            =   7590
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   4140
      Width           =   960
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "1693;582"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text6 
      Height          =   330
      Left            =   5520
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   3456
      Width           =   3015
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "5318;582"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text4 
      Height          =   405
      Left            =   4080
      TabIndex        =   5
      Top             =   405
      Width           =   4305
      VariousPropertyBits=   -1467989989
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "7594;714"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   345
      Left            =   3930
      TabIndex        =   20
      Top             =   4815
      Width           =   4605
      VariousPropertyBits=   679495707
      BackColor       =   16777215
      DisplayStyle    =   3
      Size            =   "8123;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label20 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   47
      Top             =   3456
      Width           =   825
   End
   Begin VB.Label Label19 
      BackStyle       =   0  '透明
      Caption         =   "傳票號碼"
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
      Left            =   270
      TabIndex        =   46
      Top             =   4470
      Width           =   975
   End
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "部門別"
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
      Left            =   300
      TabIndex        =   45
      Top             =   4140
      Width           =   975
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "台幣金額"
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
      Left            =   270
      TabIndex        =   44
      Top             =   4815
      Width           =   975
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "對沖(客)"
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
      Left            =   2970
      TabIndex        =   43
      Top             =   4485
      Width           =   960
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "對沖(業)"
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
      Left            =   5745
      TabIndex        =   42
      Top             =   4140
      Width           =   1230
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "對沖(本)"
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
      Left            =   2955
      TabIndex        =   41
      Top             =   4140
      Width           =   960
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "對沖(其)"
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
      Left            =   5745
      TabIndex        =   40
      Top             =   4485
      Width           =   1185
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "借1/貸2"
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
      Left            =   1590
      TabIndex        =   39
      Top             =   3450
      Width           =   855
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "摘要"
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
      Left            =   2955
      TabIndex        =   37
      Top             =   4830
      Width           =   855
   End
   Begin VB.Label Label22 
      BackStyle       =   0  '透明
      Caption         =   "筆數"
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
      Left            =   1104
      TabIndex        =   36
      Top             =   3024
      Width           =   852
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "收款類別"
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
      Left            =   5730
      TabIndex        =   33
      Top             =   3810
      Width           =   1245
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   24
      Top             =   4464
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label9 
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
      Height          =   252
      Left            =   3024
      TabIndex        =   31
      Top             =   3024
      Width           =   612
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1845
      Left            =   165
      Top             =   3390
      Width           =   8475
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "單據編號"
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
      Left            =   2940
      TabIndex        =   30
      Top             =   3810
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "外幣金額"
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
      Left            =   300
      TabIndex        =   29
      Top             =   3810
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "會計科目"
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
      Left            =   2940
      TabIndex        =   28
      Top             =   3450
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "備註"
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
      Left            =   3120
      TabIndex        =   27
      Top             =   450
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "匯率(NT)"
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
      Left            =   300
      TabIndex        =   26
      Top             =   450
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "幣別"
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
      Left            =   5880
      TabIndex        =   25
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "收款日期"
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
      Left            =   3120
      TabIndex        =   24
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收款單號"
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
      Left            =   300
      TabIndex        =   23
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc2110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/03 改成Form2.0 ; Combo2、DataGrid1改字型=新細明體-ExtB、Text4、Text6、Text18
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoacc0y0 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoacc120 As New ADODB.Recordset
Public adoacc1p0 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adocase As New ADODB.Recordset
Dim strSerialNo As String
Public strDocNo As String
'Add by Morgan 2006/6/20
Public bolForm2 As Boolean '是否要進第二畫面
'Added by Lydia 2015/10/19
Public m_A1k28 As String '提醒催款的請款對象
'Added by Morgan 2021/5/27
Public strA1P01s As String '有傳票號的公司別
Public strA1P22s As String '傳票號
'end  2021/5/27

'2012/3/8 ADD BY SONIA
Private Sub Combo1_LostFocus()
   If Combo1 = ComboItem(53) Then
      Combo2 = Combo2 & "  託收"
   End If
End Sub
'2012/3/8 END

Private Sub Combo2_GotFocus()
   TextInverse Combo2  'Added by Lydia 2021/12/14 Form 2.0的ComboBox的GotFocus不會全選反白
End Sub

Private Sub Combo3_Change()
   If Combo3 = MsgText(601) Then
      Exit Sub
   End If
   Text6 = A0102Query(Combo3)
End Sub

Private Sub Combo3_Click()
   If Combo3 = MsgText(601) Then
      Exit Sub
   End If
   Text6 = A0102Query(Combo3)
   Combo2 = Combo4 & "  " & Format(Text7, DDollar)
End Sub

Private Sub Combo3_Validate(Cancel As Boolean)
   If Combo3 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc010", "a0101", Combo3, Label6) = False Then
      Cancel = True
      Exit Sub
   End If
   'modify by sonia 2021/1/27 加傳本所案號以判別FCP,FCT英日文組
   'If AccNoToSalesNo(Combo3) <> "" Then
   '   Text13 = AccNoToSalesNo(Combo3)
   If AccNoToSalesNo(Combo3, Text12) <> "" Then
      Text13 = AccNoToSalesNo(Combo3, Text12)
      Text13_Validate True
   'end 2021/1/27
   End If
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo4_Validate(Cancel As Boolean)
   If Combo4 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc1y0", "a1y01", Combo4, Label3) = False Then
      Cancel = True
      Combo4.SetFocus
   End If
End Sub

Private Sub Command1_Click()
   If Adodc1.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc1.Recordset.Fields("a1p22").Value) = False Then
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select ax210 from acc021 where ax201 = '" & Adodc1.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc1.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(155), , MsgText(5)
            Combo3.SetFocus
            adoquery.Close
            Exit Sub
         End If
         adoquery.Close
      End If
   End If
   AdodcDelete
   SumShow
End Sub

'Add by Morgan 2005/4/27 檢查是否已過帳
Public Function AX210Exist() As Boolean
   If Adodc1.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc1.Recordset.Fields("a1p22").Value) = False Then
         adoquery.CursorLocation = adUseClient
         'Modified by Morgan 2021/4/12 案源會有L公司
         'adoquery.Open "select ax210 from acc021 where ax201 = '" & Adodc1.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc1.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         adoquery.Open "select ax210 from acc021 where (ax201,ax202) in (select a1p01,a1p22 from acc1p0 where a1p04='" & Adodc1.Recordset.Fields("a1p04").Value & "') and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         'end 2021/4/12
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(155), , MsgText(5)
            adoquery.Close
            AX210Exist = True
         Else
            adoquery.Close
         End If
         
      End If
   End If
End Function

Private Sub Command2_Click()
   'Modify by Morgan 2005/4/27
   '改Call Function 檢查
'   If Adodc1.Recordset.RecordCount <> 0 Then
'      If IsNull(Adodc1.Recordset.Fields("a1p22").Value) = False Then
'         adoquery.CursorLocation = adUseClient
'         adoquery.Open "select ax210 from acc021 where ax201 = '" & Adodc1.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc1.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
'         If adoquery.RecordCount <> 0 Then
'            MsgBox MsgText(155), , MsgText(5)
'            adoquery.Close
'            Exit Sub
'         End If
'         adoquery.Close
'      End If
'   End If
   If AX210Exist = True Then Exit Sub
   '2005/4/27 end
   
   If Text2 <> MsgText(601) Then
      strItemNo = Text2
   Else
      strItemNo = MsgText(601)
   End If
   If MaskEdBox1.Text <> MsgText(29) Then
      strCon1 = FCDate(MaskEdBox1.Text)
   Else
      strCon1 = ""
   End If
   If Combo4 <> MsgText(601) Then
      strCon2 = Combo4
   Else
      strCon2 = ""
   End If
   If Text3 <> MsgText(601) Then
      strCon3 = Text3
   Else
      strCon3 = ""
   End If
   If Text5 <> MsgText(601) Then
      strCon4 = Text5
   Else
      strCon4 = ""
   End If
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   'Modified by Morgan 2021/4/12 考慮案源收款，排除應收帳款(1133)
   adoquery.Open "select sum(a1p21) from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & Text2 & "' and a1p05 <> '1203' and not (a1p05 = '1133' and instr(a1p14||' ','法律所/')=1 ) and a1p07 <> 0", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields(0).Value) Then
         dblTotal = 0
      Else
         dblTotal = adoquery.Fields(0).Value
      End If
   Else
      dblTotal = 0
   End If
   adoquery.Close
   tool3_enabled
   Screen.MousePointer = vbHourglass
   m_A1k28 = "" 'Added by Lydia 2016/07/04 清空前一筆提醒催款的請款對象
   Frmacc2111.Show
   Frmacc2111.m_Currency = Combo4 'Add By Sindy 2014/12/10
   'Add by Morgan 2006/6/20
   bolForm2 = False
   Text2.Locked = False
   'end 2006/6/20
   Screen.MousePointer = vbDefault
   Me.Hide
End Sub

Private Sub Command3_Click()
   'Add by Morgan 2006/6/20
   If Frmacc2110.bolForm2 = True Then
      MsgBox "匯率有異動，請點【收款資料】以便重新計算台幣收款金額！", vbExclamation
      Exit Sub
   End If
   'end 2006/6/20
   Acc0y0Refresh
   If adoacc0y0.RecordCount <> 0 Then
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
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   
   'Modify by Amy 2020/09/21 從下面搬上來,只要有進第二個面畫回來,有較早帳款未付都彈
   'Added by Lydia 2015/10/19 提醒催款(較早帳款未付即時催款)
   If m_A1k28 <> "" Then
      If MsgBox("是否進行催款?", vbInformation + vbYesNo, "較早帳款未付即時催款") = vbYes Then
         Frmacc2470.SetParent Me
         Frmacc2470.Text1 = m_A1k28
         Frmacc2470.Text2 = m_A1k28
         'Added by Lydia 2015/11/02 大陸代理人要+國別
         strExc(1) = GetPrjNationNumber(m_A1k28)
         If strExc(1) = "020" Then
            Frmacc2470.Text3 = "020"
            Frmacc2470.Text4 = "020"
         End If
         'end 2015/11/02
         
         Frmacc2470.Show
         Frmacc2470.MaskEdBox2.Text = ChangeTStringToTDateString(TransDate(CompDate(2, -1, (Left(strSrvDate(1), 6) & "01")), 1))
         'Modified by Lydia 2019/11/04 數字需有千位分號
         'Frmacc2470.currAmount = Trim(Combo4.Text) & Trim(Text9.Text)
         Frmacc2470.currAmount = Trim(Combo4.Text) & Format(Val(Text9.Text), FDollar)
         Frmacc2470.strLDate = TransDate(ChangeTDateStringToTString(Me.MaskEdBox1.Text), 2)
         Call Frmacc2470.Command2_Click
         Unload Frmacc2470
         tool1_enabled
      End If
   End If
   'end 2015/10/19
   m_A1k28 = "" 'Added by Lydia 2019/11/04 已處理,清空變數
   
   If adoacc0y0.RecordCount <> 0 Then
      adoacc0y0.MoveFirst
   End If
   'adoacc0y0.Find "a0y01 = '" & strItemNo & "'", 0, adSearchForward, 1
   'If adoacc0y0.EOF = False Then
   '   FormShow
   '   AdodcRefresh
   '   SumShow
   '   RecordShow
   'End If
   If Adodc1.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc1.Recordset.Fields("a1p22").Value) = False Then
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select ax210 from acc021 where ax201 = '" & Adodc1.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc1.Recordset.Fields("a1p22").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            Text2.SetFocus
            adoquery.Close
            Acc0y0Refresh
            AdodcRefresh
            SumShow
            strCon1 = ""
            strItemNo = MsgText(601)
            Exit Sub
         End If
         adoquery.Close
      End If
   End If
   If strCon1 = "Y" Then
      SumShow
      
      If Text5 <> Text10 Then
         AdodcRefresh
         strDelConfirm = MsgBox(MsgText(11), vbOKCancel + vbDefaultButton1, MsgText(5))
         If strDelConfirm = vbCancel Then
            adoaccsum.CursorLocation = adUseClient
            adoaccsum.Open "select a0z02 from acc0z0 where a0z01 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
            Do While adoaccsum.EOF = False
               adoTaie.Execute "update acc1k0 set a1k29 = null, a1k30 = 0 where a1k01 = '" & adoaccsum.Fields(0).Value & "'"
               adoaccsum.MoveNext
            Loop
            adoaccsum.Close
            adoTaie.Execute "delete from acc1p0 where a1p02 = 'F' and a1p04 = '" & Text2 & "'"
            adoTaie.Execute "delete from acc0y0 where a0y01 = '" & Text2 & "'"
            adoTaie.Execute "delete from acc0z0 where a0z01 = '" & Text2 & "'"
            adoTaie.Execute "delete from acc120 where a1210 = '" & Text2 & "'"
            Frmacc2110_Clear
            Acc0y0Refresh
            AdodcRefresh
            SumShow
            strCon1 = ""
            strItemNo = MsgText(601)
            Exit Sub
         End If
      Else
         AdodcRefresh
'Removed by Morgan 2012/6/1 一定要存不必問--婧瑄
'         strDelConfirm = MsgBox(MsgText(131), vbOKCancel + vbDefaultButton1, MsgText(5))
'         If strDelConfirm = vbCancel Then
'            adoaccsum.CursorLocation = adUseClient
'            adoaccsum.Open "select a0z02 from acc0z0 where a0z01 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
'            Do While adoaccsum.EOF = False
'               adoTaie.Execute "update acc1k0 set a1k29 = null, a1k30 = 0 where a1k01 = '" & adoaccsum.Fields(0).Value & "'"
'               adoaccsum.MoveNext
'            Loop
'            adoaccsum.Close
'            adoTaie.Execute "delete from acc1p0 where a1p02 = 'F' and a1p04 = '" & Text2 & "'"
'            adoTaie.Execute "delete from acc0y0 where a0y01 = '" & Text2 & "'"
'            adoTaie.Execute "delete from acc0z0 where a0z01 = '" & Text2 & "'"
'            adoTaie.Execute "delete from acc120 where a1210 = '" & Text2 & "'"
'            Frmacc2110_Clear
'            Acc0y0Refresh
'            AdodcRefresh
'            SumShow
'            strCon1 = ""
'            strItemNo = MsgText(601)
'            Exit Sub
'         End If
      End If
   End If
   Text2 = strItemNo
   Acc0y0Refresh
   If adoacc0y0.RecordCount <> 0 Then
      FormShow
      AdodcRefresh
      SumShow
      RecordShow
   End If
   strItemNo = MsgText(601)
   
   'Modify by Amy 2020/09/21 原:提醒催款(較早帳款未付即時催款) 往上搬 ex:M10903988 有未付款,因ax210為null 不會彈
   
End Sub

'Added by Lydia 2021/12/03
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Added by Lydia 2021/12/03 Form2.0 記錄鍵盤傳入順序
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   'Modified by Lydia 2021/12/07 改成模組
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   Me.Width = 8850
'   Me.Height = 5600
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
   'Modify by Amy 2023/08/18 W8880 H5700
   PUB_InitForm Me, 8920, 5880, strBackPicPath1
   'end 2021/12/07
   Combo1.AddItem ComboItem(51)
   Combo1.AddItem ComboItem(52)
   Combo1.AddItem ComboItem(53)
   Combo1.AddItem ComboItem(54)
   Combo1.AddItem ComboItem(55)
   Combo3.AddItem ComboItem(61)
   Combo3.AddItem ComboItem(62)
   Combo3.AddItem ComboItem(63)
   Combo3.AddItem ComboItem(64)
   Combo3.AddItem ComboItem(65)
   Combo3.AddItem ComboItem(66)
   Combo3.AddItem ComboItem(67)
   Combo3.AddItem ComboItem(68)
   Combo3.AddItem ComboItem(69)
   Combo3.AddItem ComboItem(70)
   MaskEdBox1.Mask = DFormat
   OpenTable
   If adoacc0y0.RecordCount <> 0 Then
      adoacc0y0.MoveLast
      adoacc0y0.MoveFirst
      RecordShow
   End If
   FormDisabled
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   
   'Add by Morgan 2006/6/20
   If bolForm2 = True Then
      MsgBox "匯率有異動，請點【收款資料】以便重新計算台幣收款金額！", vbExclamation
      Cancel = 1
      Exit Sub
   End If
   'end 2006/6/20
   
   'CreDebCheck 'Removed by Morgan 2021/5/31
   If CreDebCheck <> MsgText(602) Then
      tool1_enabled
      MsgBox MsgText(11), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   strTrackMode = "" 'Added by Lydia 2021/12/03 Form2.0 記錄鍵盤傳入順序(清除)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc2110 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Public Sub OpenTable()
On Error GoTo Checking
   adoacc0y0.CursorLocation = adUseClient
   adoacc0y0.MaxRecords = intMax
   adoacc0y0.Open "select * from acc0y0 where a0y01 >= '" & Text2 & "' order by a0y01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p03 = '" & Text2 & "' and a1p04 = 'F' order by a1p05 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from acc1y0 order by a1y01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoquery.EOF = False
      Combo4.AddItem adoquery.Fields("a1y01").Value
      adoquery.MoveNext
   Loop
   adoquery.Close
   Combo4 = "USD"
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
  If adoacc0y0.RecordCount = 0 Then
     Exit Sub
  End If
  Text2 = adoacc0y0.Fields("a0y01").Value
  MaskEdBox1.Mask = MsgText(601)
  If IsNull(adoacc0y0.Fields("a0y02").Value) Then
     MaskEdBox1.Text = MsgText(601)
  Else
     MaskEdBox1.Text = CFDate(adoacc0y0.Fields("a0y02").Value)
  End If
  MaskEdBox1.Mask = DFormat
  If IsNull(adoacc0y0.Fields("a0y03").Value) Then
     Combo4 = MsgText(601)
  Else
     Combo4 = adoacc0y0.Fields("a0y03").Value
  End If
  If IsNull(adoacc0y0.Fields("a0y04").Value) Then
     Text3 = MsgText(601)
  Else
     Text3 = adoacc0y0.Fields("a0y04").Value
  End If
  If IsNull(adoacc0y0.Fields("a0y11").Value) Then
     Text4 = MsgText(601)
  Else
     Text4 = adoacc0y0.Fields("a0y11").Value
  End If
'Removed by Morgan 2021/4/12 傳票號碼可能不只一張，改顯示在明細
'    'Add By Cheng 2003/06/02
'    '顯示傳票號碼
'    Me.Text17.Text = GetA1P22(Me.Text2.Text)
'end 2021/4/12
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
   If strSaveConfirm <> MsgText(3) Then
      Exit Sub
   End If
   If Mid(MaskEdBox1.Text, 1, 3) <> Mid(CFDate(ACDate(ServerDate)), 1, 3) Then
      Text2 = UpdateNo("acc0y0", "a0y01", 5, MaskEdBox1.Text, MsgText(808))
   Else
      'Text2 = AutoNo(MsgText(808), 5)
      Text2 = strDocNo
   End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
On Error GoTo Checking
   If Text12 <> MsgText(601) Then
      Select Case Len(Text12)
         Case 7, 8, 9
            Text12 = Text12 & "000"
'         Case 10
'            Text12 = Text12 & "00"
      End Select
   End If
   If Text12 <> MsgText(601) And Text12 <> "000" Then
   '   Text12 = CaseNoZero(Text10)
      adocase.CursorLocation = adUseClient
      adocase.Open "select pa01 as SystemNo from patent where pa01 = '" & Mid(Text12, 1, Len(Text12) - 9) & "' and pa02 = '" & Mid(Text12, Len(Text12) - 8, 6) & "' and pa03 = '" & Mid(Text12, Len(Text12) - 2, 1) & "' and pa04 = '" & Mid(Text12, Len(Text12) - 1, 2) & "' union " & _
                   "select tm01 as SystemNo from trademark where tm01 = '" & Mid(Text12, 1, Len(Text12) - 9) & "' and tm02 = '" & Mid(Text12, Len(Text12) - 8, 6) & "' and tm03 = '" & Mid(Text12, Len(Text12) - 2, 1) & "' and tm04 = '" & Mid(Text12, Len(Text12) - 1, 2) & "' union " & _
                   "select lc01 as SystemNo from lawcase where lc01 = '" & Mid(Text12, 1, Len(Text12) - 9) & "' and lc02 = '" & Mid(Text12, Len(Text12) - 8, 6) & "' and lc03 = '" & Mid(Text12, Len(Text12) - 2, 1) & "' and lc04 = '" & Mid(Text12, Len(Text12) - 1, 2) & "' union " & _
                   "select hc01 as SystemNo from hirecase where hc01 = '" & Mid(Text12, 1, Len(Text12) - 9) & "' and hc02 = '" & Mid(Text12, Len(Text12) - 8, 6) & "' and hc03 = '" & Mid(Text12, Len(Text12) - 2, 1) & "' and hc04 = '" & Mid(Text12, Len(Text12) - 1, 2) & "' union " & _
                   "select sp01 as SystemNo from servicepractice where sp01 = '" & Mid(Text12, 1, Len(Text12) - 9) & "' and sp02 = '" & Mid(Text12, Len(Text12) - 8, 6) & "' and sp03 = '" & Mid(Text12, Len(Text12) - 2, 1) & "' and sp04 = '" & Mid(Text12, Len(Text12) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocase.RecordCount = 0 Then
         MsgBox MsgText(28) & Label14, , MsgText(5)
         Cancel = True
         TextInverse Text12
         adocase.Close
         Exit Sub
      End If
      adocase.Close
      'add by sonia 2021/1/27 以本所案號以判別FCP,FCT英日文組
      If AccNoToSalesNo(Combo3, Text12) <> "" Then
         Text13 = AccNoToSalesNo(Combo3, Text12)
         Text13_Validate True
      End If
      'end 2021/1/27
   End If
   Exit Sub
Checking:
   MsgBox MsgText(128), , MsgText(5)
   Exit Sub
End Sub

Private Sub Text13_GotFocus()
   TextInverse Text13
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text13_Validate(Cancel As Boolean)
   Text18.Text = ""
   If Text13 <> MsgText(601) Then
'Modify by Morgan 2007/2/5 員工已離職要提醒
'      If ExistCheck("staff", "st01", Text13, Label13) = False Then
'         Cancel = True
'         Exit Sub
'      End If
      If PUB_GetStaffState(Text13.Text, strExc(1), True) = 0 Then
         Cancel = True
         TextInverse Text13
      Else
         Text18.Text = strExc(1)
      End If
      'add by sonia 2021/1/28
      If SalesNoCheckAccNo(Combo3, Text13) = False Then
      End If
      'end 2021/1/28
'end 2007/2/5
   End If
End Sub

Private Sub Text14_GotFocus()
   TextInverse Text14
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text14_Validate(Cancel As Boolean)
   If Text14 <> MsgText(601) Then
      If Len(Text14) = 6 Then
         Text14 = AfterZero(Text14)
      'Add by Morgan 2007/3/1 八碼時要補'0'
      ElseIf Len(Text14) = 8 Then
         Text14 = Text14 & "0"
      'End 2007/3/1
      End If
      If ExistCheck("customer", "cu01", Mid(Text14, 1, 8), Label16, False) = False Then
         If ExistCheck("fagent", "fa01", Mid(Text14, 1, 8), Label16, False) = False Then
            If ExistCheck("acc0i0", "a0i01", Text14, Label16, False) = False Then
               If ExistCheck("staff", "st01", Text14, Label16, False) = False Then
                  MsgBox MsgText(28) & Label16, , MsgText(5)
                  Cancel = True
                  Exit Sub
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub Text15_GotFocus()
    TextInverse Me.Text15
End Sub
Private Sub Text15_LostFocus()
    'Added by Lydia 2016/08/02 規費只能輸入整數
    If Left(Trim(Combo3), 4) = "2201" And Text15 <> "" And Text15 <> Format(Val(Text15), DAmount) Then
        MsgBox "規費只能輸入整數!", vbCritical
        Text15.SetFocus
    End If
    'end 2016/08/02
End Sub

Private Sub Text16_GotFocus()
   TextInverse Text16
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text16_Validate(Cancel As Boolean)
   If Text16 <> MsgText(601) Then
      If ExistCheck("acc090", "a0901", Text16, Label18) = False Then
         Cancel = True
         Exit Sub
      End If
   End If
   If CheckDept(Combo3, Text16) = False Then
      MsgBox MsgText(103), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text19_GotFocus()
   TextInverse Text19
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  儲存資料表(國外收款資料(分錄檔))
'
'*************************************************
Private Sub Acc1p0Save()
Dim blnCancel As Boolean
Dim StrSQLa As String   '2011/11/9 ADD BY SONIA

On Error GoTo Checking

   'Added by Morgan 2021/4/14
   '公司別檢查
   If Text19 = MsgText(601) Then
      MsgBox MsgText(10) & Label20, , MsgText(5)
      strControlButton = MsgText(602)
      Text19.SetFocus
      Exit Sub
   End If
   'end 2021/4/14
      
   If Combo3 = MsgText(601) Then
      MsgBox MsgText(10) & Label6, , MsgText(5)
      strControlButton = MsgText(602)
      Combo3.SetFocus
      Exit Sub
   Else
      '2013/8/19 modify by sonia 加入110222,113001
      If Combo3 = "113002" Or Combo3 = "110205" Or Combo3 = "110222" Or Combo3 = "113001" Then
         If Combo1 = MsgText(601) Then
            MsgBox MsgText(10) & Label10, , MsgText(5)
            strControlButton = MsgText(602)
            Combo1.SetFocus
            Exit Sub
         End If
      End If
        'Add By Cheng 2004/04/22
        If Me.Text8.Enabled = True Then
            blnCancel = False
            Text8_Validate blnCancel
            If blnCancel = True Then
                strControlButton = MsgText(602)
                Exit Sub
            End If
        End If
        'End
      If ExistCheck("acc010", "a0101", Combo3, Label6) = False Then
         strControlButton = MsgText(602)
         Combo3.SetFocus
         Exit Sub
      End If
      If CheckDept(Combo3, Text16) = False Then
         MsgBox MsgText(103), , MsgText(5)
         strControlButton = MsgText(602)
         Text16.SetFocus
         Exit Sub
      End If
      
      'add by sonia 2015/12/30 檢查民國105年起法務收入科目不可使用
      intI = PUB_AccNoEnable(Combo3, Val(FCDate(MaskEdBox1.Text)))
      If intI <> 0 Then
         strControlButton = MsgText(602)
         Combo3.SetFocus
         Exit Sub
      End If
      'end 2015/12/30
      'Add by Morgan 2007/2/5 檢查科目部門&智權人員是否正確
      intI = PUB_AccNoGood(Combo3, Text16, Text13)
      If intI <> 0 Then
         strControlButton = MsgText(602)
         If intI = 1 Then
            Combo3.SetFocus
         ElseIf intI = 2 Then
            Text16.SetFocus
         ElseIf intI = 3 Then
            Text13.SetFocus
         End If
         Exit Sub
      End If
      'end 2007/2/5
      
      'Added by Lydia 2016/08/02 規費只能輸入整數
      If Left(Trim(Combo3), 4) = "2201" And Text15 <> "" And Text15 <> Format(Val(Text15), DAmount) Then
         MsgBox "規費只能輸入整數!", vbCritical
         Text15.SetFocus
         Exit Sub
      End If
      'end 2016/08/02
      
      'add by sonia 2021/3/4 收入及規費科目一定要有單據編號
      If (Left(Trim(Combo3), 1) = "4" Or Left(Trim(Combo3), 4) = "2201") And Text8 = "" Then
         MsgBox "收入及規費科目一定要有單據編號!", vbCritical
         Text8.SetFocus
         Exit Sub
      End If
      'end 2021/3/4
     
      If Text12 <> MsgText(601) Then
         adocase.CursorLocation = adUseClient
         adocase.Open "select pa01 as SystemNo from patent where pa01 = '" & Mid(Text12, 1, Len(Text12) - 9) & "' and pa02 = '" & Mid(Text12, Len(Text12) - 8, 6) & "' and pa03 = '" & Mid(Text12, Len(Text12) - 2, 1) & "' and pa04 = '" & Mid(Text12, Len(Text12) - 1, 2) & "' union " & _
                        "select tm01 as SystemNo from trademark where tm01 = '" & Mid(Text12, 1, Len(Text12) - 9) & "' and tm02 = '" & Mid(Text12, Len(Text12) - 8, 6) & "' and tm03 = '" & Mid(Text12, Len(Text12) - 2, 1) & "' and tm04 = '" & Mid(Text12, Len(Text12) - 1, 2) & "' union " & _
                        "select lc01 as SystemNo from lawcase where lc01 = '" & Mid(Text12, 1, Len(Text12) - 9) & "' and lc02 = '" & Mid(Text12, Len(Text12) - 8, 6) & "' and lc03 = '" & Mid(Text12, Len(Text12) - 2, 1) & "' and lc04 = '" & Mid(Text12, Len(Text12) - 1, 2) & "' union " & _
                        "select hc01 as SystemNo from hirecase where hc01 = '" & Mid(Text12, 1, Len(Text12) - 9) & "' and hc02 = '" & Mid(Text12, Len(Text12) - 8, 6) & "' and hc03 = '" & Mid(Text12, Len(Text12) - 2, 1) & "' and hc04 = '" & Mid(Text12, Len(Text12) - 1, 2) & "' union " & _
                        "select sp01 as SystemNo from servicepractice where sp01 = '" & Mid(Text12, 1, Len(Text12) - 9) & "' and sp02 = '" & Mid(Text12, Len(Text12) - 8, 6) & "' and sp03 = '" & Mid(Text12, Len(Text12) - 2, 1) & "' and sp04 = '" & Mid(Text12, Len(Text12) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adocase.RecordCount = 0 Then
            MsgBox MsgText(28) & Label14, , MsgText(5)
            strControlButton = MsgText(602)
            Text12.SetFocus
            adocase.Close
            Exit Sub
         End If
         adocase.Close
      End If
      If Text11 <> MsgText(601) Then
         If Mid(Text11, 1, 1) = "N" Then
            If adoquery.State = adStateOpen Then
               adoquery.Close
            End If
            adoquery.CursorLocation = adUseClient
            '2011/11/8 MODIFY BY SONIA IPO退費為NTD故不檢查外幣金額改檢查台幣金額
'            adoquery.Open "select a1207 from acc120 where a1201 = '" & Text11 & "'", adoTaie, adOpenStatic, adLockReadOnly
'            If adoquery.RecordCount <> 0 Then
'               If Val(Text7) <> Val(adoquery.Fields("a1207").Value) Then
'                  MsgBox MsgText(112), , MsgText(5)
'                  strControlButton = MsgText(602)
'                  Text7.SetFocus
'                  adoquery.Close
'                  Exit Sub
'               End If
'            End If
            adoquery.Open "select a1204,a1207 from acc120 where a1201 = '" & Text11 & "'", adoTaie, adOpenStatic, adLockReadOnly
            If adoquery.RecordCount <> 0 Then
               Select Case adoquery.Fields("a1204")
                  Case "NTD"
                     If Val(Text15) <> Val(adoquery.Fields("a1207").Value) Then
                        MsgBox MsgText(112), , MsgText(5)
                        strControlButton = MsgText(602)
                        Text7.SetFocus
                        adoquery.Close
                        Exit Sub
                     End If
                  Case Else
                     If Val(Text7) <> Val(adoquery.Fields("a1207").Value) Then
                        MsgBox MsgText(112), , MsgText(5)
                        strControlButton = MsgText(602)
                        Text7.SetFocus
                        adoquery.Close
                        Exit Sub
                     End If
               End Select
            End If
            '2011/11/8 END
            adoquery.Close
         End If
      End If
      If Text14 <> MsgText(601) Then
         If ExistCheck("customer", "cu01", Mid(Text14, 1, 8), Label16, False) = False Then
            If ExistCheck("fagent", "fa01", Mid(Text14, 1, 8), Label16, False) = False Then
               If ExistCheck("acc0i0", "a0i01", Text14, Label16, False) = False Then
                  If ExistCheck("staff", "st01", Text14, Label16, False) = False Then
                     MsgBox MsgText(28) & Label16, , MsgText(5)
                     strControlButton = MsgText(602)
                     Exit Sub
                  End If
               End If
            End If
         End If
      End If
      If IsNumeric(Text7) = False Then
         MsgBox MsgText(130), , MsgText(5)
         strControlButton = MsgText(602)
         Text7.SetFocus
         Exit Sub
      End If
   End If
   If Adodc1.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc1.Recordset.Fields("a1p22").Value) = False Then
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select ax210 from acc021 where ax201 = '" & Adodc1.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc1.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(155), , MsgText(5)
            strControlButton = MsgText(602)
            Combo3.SetFocus
            adoquery.Close
            Exit Sub
         End If
         adoquery.Close
      End If
   End If
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.Find "a1p03 = '" & strSerialNo & "'", 0, adSearchForward, 1
      If Adodc1.Recordset.EOF Then
         strSerialNo = ""
      End If
   Else
      strSerialNo = ""
   End If
'   adoacc1p0.CursorLocation = adUseClient
'   adoacc1p0.Open "select * from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoacc1p0.RecordCount = 0 Then
   If strSerialNo = "" Then
      Adodc1.Recordset.AddNew
      'Modified by Morgan 2021/4/14 案源會有L公司
      'Adodc1.Recordset.Fields("a1p01").Value = "1"
      Adodc1.Recordset.Fields("a1p01").Value = Text19
      
      Adodc1.Recordset.Fields("a1p02").Value = "F"
      Adodc1.Recordset.Fields("a1p03").Value = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'F' and a1p04 = '" & Text2 & "'", 3)
      Adodc1.Recordset.Fields("a1p04").Value = Text2
      
   'Added by Morgan 2021/5/31 會改公司別
   Else
      Adodc1.Recordset.Fields("a1p01").Value = Text19
   'end 2021/5/31
   End If
'   adoacc1p0.Close
   Adodc1.Recordset.Fields("a1p05").Value = Combo3
'   Adodc1.Recordset.Fields("a1p06").Value = MsgText(55)
   If Text7 <> MsgText(601) Then
      Adodc1.Recordset.Fields("a1p21").Value = Val(Text7)
   Else
      Adodc1.Recordset.Fields("a1p21").Value = 0
   End If
   If Text8 <> MsgText(601) Then
      Adodc1.Recordset.Fields("a1p23").Value = Text8
   Else
      Adodc1.Recordset.Fields("a1p23").Value = Null
   End If
   If Combo1 <> MsgText(601) Then
      Adodc1.Recordset.Fields("a1p24").Value = Mid(Combo1, 1, 1)
   Else
      Adodc1.Recordset.Fields("a1p24").Value = Null
   End If
   If Combo4 <> MsgText(601) Then
      Adodc1.Recordset.Fields("a1p19").Value = Combo4
   Else
      Adodc1.Recordset.Fields("a1p19").Value = Null
   End If
   If Combo2 <> MsgText(601) Then
      Adodc1.Recordset.Fields("a1p14").Value = Combo2
      Combo2.AddItem Adodc1.Recordset.Fields("a1p14").Value
   Else
      Adodc1.Recordset.Fields("a1p14").Value = Null
   End If
   If MaskEdBox1.Text <> MsgText(29) Then
      Adodc1.Recordset.Fields("a1p18").Value = FCDate(MaskEdBox1.Text)
   Else
      Adodc1.Recordset.Fields("a1p18").Value = Null
   End If
   Select Case Text1
      Case "1"
         If Text3 <> MsgText(601) Then
            Adodc1.Recordset.Fields("a1p20").Value = Val(Text3)
            'Adodc1.Recordset.Fields("a1p07").Value = Val(Format(Val(Text7) * Val(Text3), FAmount))
            'If Text7 <> MsgText(601) And Text7 <> "0" Then
            '   Adodc1.Recordset.Fields("A1P08").Value = 0
            'End If
            '2011/11/9 ADD BY SONIA 暫收款單號存暫收款單之匯率及暫收款單金額
            If Combo3 = "2401" Then
               '2011/11/24 modify by sonia 加a1204
               StrSQLa = "select a1205,a1207,a1204 from acc120 where a1201 = '" & Text8 & "'"
               adoacc120.CursorLocation = adUseClient
               adoacc120.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
               '若有資料
               If adoacc120.RecordCount <> 0 Then
                  Adodc1.Recordset.Fields("a1p20").Value = "" & adoacc120.Fields(0).Value
                  '2011/11/24 modify by sonia NTD暫收款不帶A1207改帶畫面輸的
                  'Adodc1.Recordset.Fields("a1p21").Value = "" & adoacc120.Fields(1).Value
                  If adoacc120.Fields("a1204") <> "NTD" Then Adodc1.Recordset.Fields("a1p21").Value = "" & adoacc120.Fields(1).Value
                  '2011/11/24 end
               End If
               adoacc120.Close
            End If
            '2011/11/9 END
         Else
            Adodc1.Recordset.Fields("a1p20").Value = 0
            'Adodc1.Recordset.Fields("a1p07").Value = 0
            'Adodc1.Recordset.Fields("a1p08").Value = 0
         End If
         If Text15 <> MsgText(601) Then
            Adodc1.Recordset.Fields("a1p07").Value = Val(Text15)
            Adodc1.Recordset.Fields("a1p08").Value = 0
         Else
            Adodc1.Recordset.Fields("a1p07").Value = 0
            Adodc1.Recordset.Fields("a1p08").Value = 0
         End If
      Case "2"
         If Mid(Combo3, 1, 1) <> "2" And Mid(Combo3, 1, 1) <> "4" Then
            If Text3 <> MsgText(601) Then
               Adodc1.Recordset.Fields("a1p20").Value = Val(Text3)
               'Adodc1.Recordset.Fields("a1p08").Value = Val(Format(Val(Text7) * Val(Text3), FAmount))
               'If Text7 <> MsgText(601) And Text7 <> "0" Then
               '   Adodc1.Recordset.Fields("A1P07").Value = 0
               'End If
            Else
               Adodc1.Recordset.Fields("a1p20").Value = 0
               'Adodc1.Recordset.Fields("a1p07").Value = 0
               'Adodc1.Recordset.Fields("a1p08").Value = 0
            End If
         End If
         If Text15 <> MsgText(601) Then
            Adodc1.Recordset.Fields("a1p08").Value = Val(Text15)
            Adodc1.Recordset.Fields("a1p07").Value = 0
         Else
            Adodc1.Recordset.Fields("a1p08").Value = 0
            Adodc1.Recordset.Fields("a1p07").Value = 0
         End If
   End Select
   If Text12 <> MsgText(601) Then
      Adodc1.Recordset.Fields("a1p17").Value = Text12
   Else
      Adodc1.Recordset.Fields("a1p17").Value = Null
   End If
   If Text13 <> MsgText(601) Then
      Adodc1.Recordset.Fields("a1p16").Value = Text13
   Else
      Adodc1.Recordset.Fields("a1p16").Value = Null
   End If
   If Text14 <> MsgText(601) Then
      Adodc1.Recordset.Fields("a1p15").Value = Text14
   Else
      Adodc1.Recordset.Fields("a1p15").Value = Null
   End If
   If Text11 <> MsgText(601) Then
      Adodc1.Recordset.Fields("a1p30").Value = Text11
   Else
      Adodc1.Recordset.Fields("a1p30").Value = Null
   End If
   If Text16 <> MsgText(601) Then
      Adodc1.Recordset.Fields("a1p06").Value = Text16
   Else
      Adodc1.Recordset.Fields("a1p06").Value = MsgText(55)
   End If
   Adodc1.Recordset.UpdateBatch
   AdodcRefresh
   strSerialNo = MsgText(601)
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
   Text19 = "" & Adodc1.Recordset.Fields("a1p01").Value  'Added by Morgan 2021/4/14
   If Adodc1.Recordset.Fields("a1p07").Value = 0 Then
      Text1 = "2"
   Else
      Text1 = "1"
   End If
   Combo3 = Adodc1.Recordset.Fields("a1p05").Value
   If IsNull(Adodc1.Recordset.Fields("a1p21").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = Adodc1.Recordset.Fields("a1p21").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p23").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = Adodc1.Recordset.Fields("a1p23").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p24").Value) Then
      Combo1 = MsgText(601)
   Else
      Combo1 = Combo1.List(Val(Adodc1.Recordset.Fields("a1p24").Value) - 1)
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p17").Value) Then
      Text12 = MsgText(601)
   Else
      Text12 = Adodc1.Recordset.Fields("a1p17").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p16").Value) Then
      Text13 = MsgText(601)
      Text18 = ""
   Else
      Text13 = Adodc1.Recordset.Fields("a1p16").Value
      Text18 = StaffQuery(Text13)
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p15").Value) Then
      Text14 = MsgText(601)
   Else
      Text14 = Adodc1.Recordset.Fields("a1p15").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p30").Value) Then
      Text11 = MsgText(601)
   Else
      Text11 = Adodc1.Recordset.Fields("a1p30").Value
   End If
   If Text1 = "1" Then
      If IsNull(Adodc1.Recordset.Fields("a1p07").Value) Then
         Text15 = MsgText(601)
      Else
         Text15 = Adodc1.Recordset.Fields("a1p07").Value
      End If
   Else
      If IsNull(Adodc1.Recordset.Fields("a1p08").Value) Then
         Text15 = MsgText(601)
      Else
         Text15 = Adodc1.Recordset.Fields("a1p08").Value
      End If
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p14").Value) Then
      Combo2 = ""
   Else
      Combo2 = Adodc1.Recordset.Fields("a1p14").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p06").Value) Then
      Text16 = MsgText(601)
   Else
      Text16 = Adodc1.Recordset.Fields("a1p06").Value
   End If
   
   Text17 = "" & Adodc1.Recordset.Fields("a1p22").Value 'Added by Morgan 2021/4/12
   
   'Added by Morgan 2022/4/25
   '已有傳票號時不可改公司別
   If Text17 <> "" Then
      Text19.Enabled = False
   Else
      Text19.Enabled = Text1.Enabled
   End If
   'end 2022/4/25
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   'Modified by Morgan 2021/4/12 可能有不只1家公司,取消 and a1p01 = '1'
   adoadodc1.Open "select * from acc1p0, acc010 where a1p05 = a0101 and a1p02 = 'F' and a1p04 = '" & Text2 & "' order by a1p01 asc, a1p03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.Find "a1p03 = '" & strSerialNo & "'", 0, adSearchForward, 1
      If Adodc1.Recordset.EOF Then
         Adodc1.Recordset.MoveFirst
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
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)

   Call PUB_SaveTrackMode(1, KeyCode)  'Added by Lydia 2021/12/03 Form2.0 記錄鍵盤傳入順序
   
   Select Case KeyCode
      Case vbKeyInsert
         If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
            Exit Sub
         End If
         'Added by Lydia 2021/12/03 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
         If PUB_ChkTrackMode = False Then
             Exit Sub
         End If
         'end 2021/12/03
         'Added by Lydia 2021/12/03 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
         If PUB_ChkUniText(Me, , True, "ComboBox") = False Then
             Exit Sub
         End If
         'end 2021/12/03
         'Frmacc2110_Save
         If strControlButton <> MsgText(602) Then
            Acc1p0Save
         End If
         If strControlButton <> MsgText(602) Then
            SumShow
            AdodcClear
            Text1.SetFocus
         End If
         strControlButton = MsgText(601)
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Calculate
   End If
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
   'edit by nickc 2007/06/11  切換輸入法改用API
   OpenIme
End Sub

Private Sub Text4_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
CloseIme
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
Dim StrSQLa As String

   '2009/10/20 MODIFY BY SONIA 瑞婷說 611301摘要預設 手續費, 不要帶金額
   'Combo2 = Combo4 & "  " & Format(Text7, FDollar)
   If Combo3 = "611301" Then
      Combo2 = "手續費"
   Else
      Combo2 = Combo4 & "  " & Format(Text7, FDollar)
   End If
   Text15 = Val(Format(Val(Text7) * Val(Text3), FAmount))
   
   '2015/7/2 ADD BY SONIA 暫收款抓原產生時傳票的貸方金額 N10200206 抓 M10203020
   If Combo3 = "2401" Then
      StrSQLa = "select a1p08,a1204 from acc120,acc1p0 where a1201 = '" & Text8 & "' and a1p30 = '" & Text8 & "' and a1p08>0"
      adoacc120.CursorLocation = adUseClient
      adoacc120.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
      '若有資料
      If adoacc120.RecordCount <> 0 Then
         If adoacc120.Fields("a1204") <> "NTD" Then Text15 = "" & adoacc120.Fields(0).Value
      End If
      adoacc120.Close
   End If
   '2015/7/2 END

End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  清除 Adodc 顯示資料
'
'*************************************************
Public Sub AdodcClear()
   Text1 = ""
   Combo3 = ""
   Text6 = ""
   Text7 = ""
   Text8 = ""
   Combo1 = ""
   Text12 = ""
   Text13 = ""
   Text14 = ""
   Text11 = ""
   Text15 = ""
   Text16 = ""
   Combo2 = ""
   m_A1k28 = "" 'Added by Lydia 2015/10/19
   Text19 = "1" 'Added by Morgan 2021/4/15
   'Added by Morgan 2022/4/25
   Text19.Enabled = Text1.Enabled
   Text17 = ""
   'end 2022/4/25
   
End Sub

'*************************************************
'  刪除 Adodc 之資料
'
'*************************************************
Private Sub AdodcDelete()
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount <> 0 Then
      adoTaie.Execute "delete from acc1p0 where a1p02 = 'F' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text2 & "'"
      AdodcRefresh
      AdodcClear
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  計算並顯示合計資料
'
'*************************************************
Public Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   'Modified by Morgan 2021/4/14 案源收款會有L公司,取消 a1p01='1' 條件
   adoaccsum.Open "select sum(decode(a1p07, 0, 0, a1p21)), sum(nvl(a1p07, 0)), sum(nvl(a1p08, 0)), count(*) from acc1p0 where a1p02 = 'F' and a1p04 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text9 = MsgText(601)
      Else
         Text9 = Format(adoaccsum.Fields(0).Value, FAmount)
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text5 = MsgText(601)
      Else
         Text5 = Format(adoaccsum.Fields(1).Value, FAmount)
      End If
      If IsNull(adoaccsum.Fields(2).Value) Then
         Text10 = MsgText(601)
      Else
         Text10 = Format(adoaccsum.Fields(2).Value, FAmount)
      End If
      If IsNull(adoaccsum.Fields(3).Value) Then
         Text20 = MsgText(601)
      Else
         Text20 = adoaccsum.Fields(3).Value
      End If
   Else
      Text9 = MsgText(601)
      Text5 = MsgText(601)
      Text10 = MsgText(601)
      Text20 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = adoacc0y0.Bookmark & MsgText(35) & adoacc0y0.RecordCount
End Sub

'*************************************************
'  關閉分錄欄位輸入狀態
'
'*************************************************
Public Sub FormDisabled()
   Text19.Enabled = False 'Added by Morgan 2021/4/12
   Text1.Enabled = False
   Combo3.Enabled = False
   Text7.Enabled = False
   Text8.Enabled = False
   Combo1.Enabled = False
   Text12.Enabled = False
   Text13.Enabled = False
   Text14.Enabled = False
   Text11.Enabled = False
   Text15.Enabled = False
   Text16.Enabled = False
   Combo2.Enabled = False
   Command1.Enabled = False
   Command2.Enabled = True
End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'
'*************************************************
Public Sub FormEnabled()
   'Modified by Morgan 2022/4/25
   'Text19.Enabled = True 'Added by Morgan 2021/4/12
   If Text17.Text <> "" Then
      Text19.Enabled = False
   Else
      Text19.Enabled = True
   End If
   'end 2022/4/25
   Text1.Enabled = True
   Combo3.Enabled = True
   Text7.Enabled = True
   Text8.Enabled = True
   Combo1.Enabled = True
   Text12.Enabled = True
   Text13.Enabled = True
   Text14.Enabled = True
   Text11.Enabled = True
   Text15.Enabled = True
   Text16.Enabled = True
   Combo2.Enabled = True
   Command1.Enabled = True
   Command2.Enabled = False
   
   'Added by Morgan 2022/6/29 有傳票號不可改收款日期
   If Text2 <> "" Then
      MaskEdBox1.Enabled = False
      If CheckExistA1p22("", "F", Text2.Text) = False Then
         MaskEdBox1.Enabled = True
      End If
   End If
   'end 2022/6/29
End Sub

Private Sub Text8_Validate(Cancel As Boolean)
Dim StrSQLa As String
    
   If Me.Text8.Text = "" Then Exit Sub
   If Combo3 = "2401" Then
        '2011/11/9 MODIFY BY SONIA CU05,FA05未帶出,再加A1207
        'StrSQLa = "select a1202, a1204, a1205, '' from acc120 where a1201 = '" & Text8 & "' And A1203 Is Null "
        'StrSQLa = StrSQLa & " Union select a1202, a1204, a1205, '' As CU05 from acc120, Customer where substr(A1203,1,8)=CU01 And substr(A1203,9,1)=CU02 And a1201 = '" & Text8 & "' "
        'StrSQLa = StrSQLa & " Union select a1202, a1204, a1205, '' As FA05 from acc120, Fagent where substr(A1203,1,8)=FA01 And substr(A1203,9,1)=FA02 And a1201 = '" & Text8 & "' "
        '2015/7/2 ADD BY SONIA 暫收款抓原產生時傳票的貸方金額 N10200206 抓 M10203020
        'StrSQLa = "select a1202, a1204, a1205, '', a1207 from acc120 where a1201 = '" & Text8 & "' And A1203 Is Null "
        'StrSQLa = StrSQLa & " Union select a1202, a1204, a1205, CU05, a1207 from acc120, Customer where substr(A1203,1,8)=CU01 And substr(A1203,9,1)=CU02 And a1201 = '" & Text8 & "' "
        'StrSQLa = StrSQLa & " Union select a1202, a1204, a1205, FA05,a1207 from acc120, Fagent where substr(A1203,1,8)=FA01 And substr(A1203,9,1)=FA02 And a1201 = '" & Text8 & "' "
        'modify by sonia 2024/6/17 +a1203
        StrSQLa = "select a1202, a1204, a1205, '', a1207,a1p08,a1203 from acc120,acc1p0 where a1201 = '" & Text8 & "' And A1203 Is Null and a1p30 = '" & Text8 & "' and a1p08>0"
        StrSQLa = StrSQLa & " Union select a1202, a1204, a1205, CU05, a1207,a1p08,a1203 from acc120,acc1p0, Customer where substr(A1203,1,8)=CU01 And substr(A1203,9,1)=CU02 And a1201 = '" & Text8 & "' and a1p30 = '" & Text8 & "' and a1p08>0"
        StrSQLa = StrSQLa & " Union select a1202, a1204, a1205, FA05,a1207,a1p08,a1203 from acc120,acc1p0, Fagent where substr(A1203,1,8)=FA01 And substr(A1203,9,1)=FA02 And a1201 = '" & Text8 & "' and a1p30 = '" & Text8 & "' and a1p08>0"
        '2015/7/2 END
        '2011/11/9 END
      adoacc120.CursorLocation = adUseClient
      adoacc120.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
      '若有資料
      If adoacc120.RecordCount <> 0 Then
        'Add By Cheng 2004/04/27
        '對沖(其)=單據編號
        Me.Text11.Text = Me.Text8.Text
        'End
'         If IsNull(adoacc120.Fields(1).Value) Then
'            Text1 = ""
'         Else
'            Text1 = adoacc120.Fields(1).Value
'         End If
'         If IsNull(adoacc120.Fields(2).Value) Then
'            Text3 = ""
'         Else
'            Text3 = adoacc120.Fields(2).Value
'         End If
'         Combo2 = Text1 & "  " & Format(Text7, DDollar)
         '2011/11/9 改摘要順序,原為幣別+外幣金額/暫收款日期/客戶或代理人英文名稱, 再加暫收款單號
         Combo2 = ""
         If IsNull(adoacc120.Fields(3).Value) = False Then
            '2013/2/20 modify by sonia 只抓前20碼,否則外帳的科目分類帳會蓋掉金額D101020305
            'Combo2 = Combo2 & adoacc120.Fields(3).Value  '客戶或代理人英文名稱
            Combo2 = Combo2 & Left(adoacc120.Fields(3).Value, 20) '客戶或代理人英文名稱
         End If
         'Combo2 = Combo2 & adoacc120.Fields(1).Value & " " & Text7  '幣別+外幣金額
         Combo2 = Combo2 & "/" & adoacc120.Fields(1).Value & " " & adoacc120.Fields(4).Value  '幣別+外幣金額
         If IsNull(adoacc120.Fields(0).Value) = False Then
            Combo2 = Combo2 & "/" & CFDate(adoacc120.Fields(0).Value) '暫收款日期
         End If
         Combo2 = Combo2 & "/" & Text8
         If "" & adoacc120.Fields(1).Value = "NTD" Then
            Text15 = Val(Format(Val(adoacc120.Fields(4).Value), FAmount))
         Else
            '2015/7/2 ADD BY SONIA 暫收款抓原產生時傳票的貸方金額 N10200206 抓 M10203020
            'Text15 = Val(Format(Val(Text7) * Val(adoacc120.Fields(2).Value), FAmount))
            Text15 = "" & adoacc120.Fields(5).Value
            '2015/7/2 END
         End If
         '2011/11/9 END
         Text14 = "" & adoacc120.Fields("a1203").Value  'add by sonia 2024/6/17 預設對沖(客)
        '若無資料
      Else
'         Text1 = ""
'         Text3 = ""
            'Add By Cheng 2004/04/22
            MsgBox "單據編號輸入錯誤!!!", vbExclamation + vbOKOnly
            TextInverse Me.Text8
            Cancel = True
            'End
      End If
      adoacc120.Close
   End If
   '2006/3/13 ADD BY SONIA
   '2013/8/15 MODIFY BY SONIA 加入110205科目
   '2013/8/19 modify by sonia 加入110222,113001
   If Combo3 = "113002" Or Combo3 = "110205" Or Combo3 = "110222" Or Combo3 = "113001" Then
      If InStr(1, Combo2, Me.Text8.Text) > 0 Then
      Else
         Combo2 = Combo2 & "/" & Me.Text8.Text '單據編號
      End If
   End If
   '2006/3/13 END
End Sub

'*************************************************
'  重新整理國外收款資料
'
'*************************************************
Public Sub Acc0y0Refresh()
On Error GoTo Checking
   If adoacc0y0.State = adStateOpen Then
      adoacc0y0.Close
   End If
   adoacc0y0.CursorLocation = adUseClient
   adoacc0y0.MaxRecords = intMax
   adoacc0y0.Open "select * from acc0y0 where a0y01 >= '" & Text2 & "' order by a0y01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
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
   Dim strQ As String, intQ As Integer
   Dim rstQ As ADODB.Recordset
   
   If Text5 = Text10 Then
      'Modified by Morgan 2021/5/31 所有公司別都沒有不平
      'CreDebCheck = MsgText(602)
      strQ = "select * from (select a1p01 COMP,sum(a1p07) AMT1,sum(a1p08) AMT2,sum(a1p07)-sum(a1p08) AMT3 from acc1p0 where a1p04='" & Text2 & "' group by a1p01) where AMT3<>0"
      intQ = 1
      Set rstQ = ClsLawReadRstMsg(intQ, strQ)
      If intQ = 0 Then
         CreDebCheck = MsgText(602)
      End If
      'end 2021/5/31
   End If
End Function

'*************************************************
'  重新計算借貸方資料
'
'*************************************************
Public Sub Calculate()
End Sub

'Removed by Morgan 2021/4/12
''Add By Cheng 2003/06/02
'Private Function GetA1P22(strA1P04 As String) As String
'Dim StrSQLa As String
'Dim rsA As New ADODB.Recordset
'
'GetA1P22 = ""
'StrSQLa = "Select * From ACC1P0 Where A1P01='1' And A1P02='F' And A1P04='" & strA1P04 & "' "
'rsA.CursorLocation = adUseClient
'rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'If rsA.RecordCount > 0 Then
'    GetA1P22 = "" & rsA("A1P22").Value
'End If
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'End Function
