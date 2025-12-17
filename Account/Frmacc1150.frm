VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc1150 
   AutoRedraw      =   -1  'True
   Caption         =   "收款作業"
   ClientHeight    =   5100
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   8844
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5100
   ScaleWidth      =   8844
   Begin VB.CommandButton Command5 
      Appearance      =   0  '平面
      Caption         =   "主要公司別"
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
      Left            =   5220
      TabIndex        =   51
      Top             =   75
      Width           =   1335
   End
   Begin VB.TextBox Text21 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6570
      MaxLength       =   1
      TabIndex        =   3
      Top             =   75
      Width           =   345
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8010
      Picture         =   "Frmacc1150.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   21
      ToolTipText     =   "取消"
      Top             =   2490
      Width           =   350
   End
   Begin VB.CommandButton Command4 
      Caption         =   "公司別合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6525
      TabIndex        =   50
      Top             =   2490
      Width           =   1290
   End
   Begin VB.TextBox Text19 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1230
      MaxLength       =   1
      TabIndex        =   5
      Top             =   3165
      Width           =   390
   End
   Begin VB.TextBox Text15 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3972
      MaxLength       =   12
      TabIndex        =   18
      Top             =   4440
      Width           =   1572
   End
   Begin VB.TextBox Text14 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1230
      MaxLength       =   3
      TabIndex        =   17
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox Text20 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   41
      Top             =   2508
      Width           =   855
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5100
      TabIndex        =   40
      Top             =   2508
      Width           =   1308
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6840
      MaxLength       =   14
      TabIndex        =   8
      Top             =   3172
      Width           =   1572
   End
   Begin VB.CommandButton Command3 
      Height          =   315
      Left            =   2385
      Picture         =   "Frmacc1150.frx":066A
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   75
      Width           =   350
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1230
      TabIndex        =   15
      Top             =   4128
      Width           =   1665
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc1150.frx":076C
      Height          =   2010
      Left            =   135
      TabIndex        =   22
      Top             =   450
      Width           =   7275
      _ExtentX        =   12827
      _ExtentY        =   3535
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   10
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
      BeginProperty Column04 
         DataField       =   "a1p17"
         Caption         =   "對沖(本)"
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
      BeginProperty Column06 
         DataField       =   "a0g02"
         Caption         =   "收票銀行"
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
         DataField       =   "a1p11"
         Caption         =   "收票帳號"
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
      BeginProperty Column09 
         DataField       =   "a1p23"
         Caption         =   "暫收款單號"
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
            ColumnWidth     =   2171.906
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1307.906
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1272.189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   5567.812
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1332.284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1751.811
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2831.811
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1451.906
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   1247.811
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1548.284
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6840
      MaxLength       =   30
      TabIndex        =   14
      Top             =   3828
      Width           =   1572
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   11
      Top             =   3495
      Width           =   1572
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1230
      MaxLength       =   12
      TabIndex        =   12
      Top             =   3828
      Width           =   1665
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3960
      MaxLength       =   8
      TabIndex        =   10
      Top             =   3495
      Width           =   1572
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5190
      MaxLength       =   14
      TabIndex        =   7
      Top             =   3172
      Width           =   1572
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2955
      TabIndex        =   30
      Top             =   3165
      Width           =   2160
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1665
      MaxLength       =   6
      TabIndex        =   6
      Top             =   3150
      Width           =   1305
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3765
      TabIndex        =   28
      Top             =   2508
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "收款明細"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7155
      TabIndex        =   23
      Top             =   75
      Width           =   1425
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1065
      MaxLength       =   15
      TabIndex        =   0
      Top             =   75
      Width           =   1305
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   315
      Left            =   3795
      TabIndex        =   2
      Top             =   75
      Width           =   1215
      _ExtentX        =   2159
      _ExtentY        =   572
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.4
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
      Left            =   3960
      TabIndex        =   13
      Top             =   3828
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.4
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
      Left            =   252
      Top             =   624
      Visible         =   0   'False
      Width           =   972
      _ExtentX        =   2117
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
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1230
      MaxLength       =   8
      TabIndex        =   9
      Top             =   3495
      Width           =   720
   End
   Begin MSForms.TextBox Text18 
      Height          =   330
      Left            =   1230
      TabIndex        =   20
      Top             =   4740
      Width           =   1695
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "2990;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text16 
      Height          =   330
      Left            =   6825
      TabIndex        =   19
      Top             =   4455
      Width           =   1605
      VariousPropertyBits=   671105051
      MaxLength       =   9
      Size            =   "2831;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   330
      Left            =   3960
      TabIndex        =   16
      Top             =   4125
      Width           =   4485
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "7911;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   1755
      Left            =   7425
      TabIndex        =   4
      Top             =   705
      Width           =   1275
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "2249;3096"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text17 
      Height          =   330
      Left            =   1935
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   3495
      Width           =   990
      VariousPropertyBits=   671105051
      MaxLength       =   8
      Size            =   "1746;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblA1P22 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   3975
      TabIndex        =   53
      Top             =   4740
      Width           =   1572
   End
   Begin VB.Label Label21 
      BackStyle       =   0  '透明
      Caption         =   "傳票號碼"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   52
      Top             =   4770
      Width           =   975
   End
   Begin VB.Label Label19 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   49
      Top             =   3180
      Width           =   975
   End
   Begin VB.Label Label20 
      BackStyle       =   0  '透明
      Caption         =   "對沖(其)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   47
      Top             =   4740
      Width           =   975
   End
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "對沖(客)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5880
      TabIndex        =   46
      Top             =   4452
      Width           =   972
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "對沖(本)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3000
      TabIndex        =   45
      Top             =   4440
      Width           =   972
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "部門別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   44
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   43
      Top             =   3525
      Width           =   975
   End
   Begin VB.Label Label22 
      BackStyle       =   0  '透明
      Caption         =   "筆數"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   42
      Top             =   2508
      Width           =   852
   End
   Begin VB.Label Label14 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "貸方金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6840
      TabIndex        =   39
      Top             =   2928
      Width           =   1572
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "摘要"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3000
      TabIndex        =   38
      Top             =   4128
      Width           =   612
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "票別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   37
      Top             =   4125
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4560
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "暫收款單號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5640
      TabIndex        =   36
      Top             =   3852
      Width           =   1212
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "收票銀行"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5640
      TabIndex        =   35
      Top             =   3528
      Width           =   972
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "到期日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3000
      TabIndex        =   34
      Top             =   3828
      Width           =   972
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "收票帳號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   33
      Top             =   3825
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "票據號碼"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3000
      TabIndex        =   32
      Top             =   3528
      Width           =   972
   End
   Begin VB.Label Label6 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "借方金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5190
      TabIndex        =   31
      Top             =   2925
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "會計科目"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1710
      TabIndex        =   29
      Top             =   2925
      Width           =   1260
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3075
      TabIndex        =   27
      Top             =   2505
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收款單號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "備註"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7500
      TabIndex        =   25
      Top             =   420
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收款日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2835
      TabIndex        =   24
      Top             =   120
      Width           =   900
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   2190
      Left            =   135
      Top             =   2910
      Width           =   8475
   End
End
Attribute VB_Name = "Frmacc1150"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit

Public adoacc0l0 As New ADODB.Recordset
Public adoacc1p0 As New ADODB.Recordset
Public adoacc010 As New ADODB.Recordset
Public adoacc0g0 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Dim douTotalAmount As Double
Dim strSerialNo As String
Public strDocNo As String

'Added by Morgan 2013/12/19
Public m_A4401 As String '繳款人員
Public m_A4402 As String '繳款日期
Public m_A4403 As String '繳款時間
Public m_AutoRun As Boolean
Dim m_Rebuild As Boolean, m_Activated As Boolean, m_SaveCheck As Boolean

'Added by Morgan 2014/6/25
'Public m_A1P22_1 As String 'J公司傳票號  'cancel by sonia 2020/4/24
'Public m_A1P22_J As String '1公司傳票號  'cancel by sonia 2020/4/24

'add by sonia 2020/4/24
Public strA1P01s As String '有傳票號的公司別
Public strA1P22s As String '傳票號
'end 2020/4/24
Dim m_TTRcpSQL As String 'Added by Morgan 2023/11/17 整批收款的案源智慧所收據語法
Dim m_bolAlert As Boolean 'Added by Morgan 2025/6/13

Private Sub Combo2_GotFocus()
   TextInverse Text2
   'edit by nickc 2007/06/11  切換輸入法改用API
   OpenIme
End Sub

Private Sub Combo2_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
CloseIme
End Sub

Private Sub Command1_Click()
   'Added by Morgan 2025/6/13
   If m_bolAlert Then
      MsgBox "票期超過規定,是否已有核准簽呈?", vbExclamation
      m_bolAlert = False
   End If
   'end 2025/6/13
   
   'Added by Morgan 2013/12/26
   If m_AutoRun = True Then
      m_AutoRun = False
      strSql = "update acc440 set a4416='" & Text2 & "' where A4401='" & m_A4401 & "' and A4402=" & m_A4402 & " and A4403=" & m_A4403
      adoTaie.Execute strSql, intI
      Frmacc1155.m_iReturn = 1
   End If
   'end 2013/12/26
   
   If Adodc1.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc1.Recordset.Fields("a1p22").Value) = False Then
         If adoquery.State = adStateOpen Then adoquery.Close       'Added by Lydia 2024/11/28
         adoquery.CursorLocation = adUseClient
         'Modified by Morgan 2013/12/19 一張收款單有可能有兩張傳票號
         'adoQuery.Open "select ax210 from acc021 where ax201 = '" & Adodc1.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc1.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         adoquery.Open "select ax210 from acc021 where (AX201,AX202) in (select a1p01,a1p22 from acc1p0 where a1p04='" & Adodc1.Recordset.Fields("a1p04").Value & "') and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(155), , MsgText(5)
            adoquery.Close
            Exit Sub
         End If
         adoquery.Close
      End If
   End If
   If Text3 = "" Then
      MsgBox MsgText(179), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Exit Sub
   End If
   If Text2 <> MsgText(601) Then
      strItemNo = Text2
   Else
      Exit Sub
   End If
'   If Text3 <> MsgText(601) Then
'      dblTotal = douTotalAmount
'   Else
'      dblTotal = 0
'   End If
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   'Modify by Morgan 2006/5/30 排除"點作專業"
   'adoquery.Open "select sum(a1p07) from acc1p0 where a1p01 = '1' and a1p02 = 'A' and a1p04 = '" & Text2 & "' and a1p05 <> '1203' and a1p07 <> 0", adoTaie, adOpenStatic, adLockReadOnlyadoquery.Open "select sum(a1p07) from acc1p0 where a1p01 = '1' and a1p02 = 'A' and a1p04 = '" & Text2 & "' and a1p05 <> '1203' and a1p07 <> 0", adoTaie, adOpenStatic, adLockReadOnly
   'Modified by Morgan 2013/12/27 收款會有J公司,取消 a1p01='1' 條件, 排除 1133,2141 科目
   'adoQuery.Open "select sum(a1p07) from acc1p0 where a1p01 = '1' and a1p02 = 'A' and a1p04 = '" & Text2 & "' and a1p05 <> '1203' and a1p07 <> 0 and (a1p14 is null or instr(a1p14,'點作轉專業')=0)", adoTaie, adOpenStatic, adLockReadOnly
   'Modified by Morgan 2014/2/12 +排除 2405,2631 科目
   'Modified by Morgan 2015/8/21 +排除 4開頭,2201開頭 科目
   'modify by sonia 2020/2/14 已排除4開頭就已排除(點作轉專業)故不再改(專業支援)
   'modify by sonia 2020/4/23 因法律所收款同時將智慧所案源收據同時收款故取消and a1p05 <> '1133'
   'Modified by Morgnan 2021/1/15 法律所收款改系統自動產生智慧所分錄1133應收帳款加回,並排除主要公司非L的L公司現金
   'adoquery.Open "select sum(a1p07) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "' and a1p05 <> '1203' and a1p05 <> '2141'  and a1p05 <> '2405'  and a1p05 <> '2613' and a1p07 <> 0 and a1p05 not like '4%' and a1p05 not like '2201%' and (a1p14 is null or instr(a1p14,'點作轉專業')=0)", adoTaie, adOpenStatic, adLockReadOnly
   'Modfied by Morgan 2021/3/10 排除L公司的勞務費
   'Modified by Morgan 2023/8/16 抓主要公司 或 TT999999案1公司瑞興銀存110602
   'adoquery.Open "select sum(a1p07) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "' and a1p05 <> '1203' and a1p05 <> '2141'  and a1p05 <> '2405'  and a1p05 <> '2613'  and a1p05 <> '1133' and a1p07 <> 0 and a1p05 not like '4%' and a1p05 not like '2201%' and (a1p14 is null or instr(a1p14,'點作轉專業')=0) and not (a1p01<>'" & Text21 & "' and a1p01='L' and a1p05='1101' and a1p16='L0100') and not (a1p01='L' and a1p05='6129')", adoTaie, adOpenStatic, adLockReadOnly
   'Modified by Morgan 2023/10/19 -保留點數  and not (a1p05='2492' and a1p23 is not null)
   adoquery.Open "select sum(a1p07) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "' and a1p05 <> '1203' and a1p05 <> '2141'  and a1p05 <> '2405'  and a1p05 <> '2613'  and a1p05 <> '1133' and a1p07 <> 0 and a1p05 not like '4%' and a1p05 not like '2201%' and (a1p14 is null or instr(a1p14,'點作轉專業')=0) and not (a1p01<>'" & Text21 & "' and a1p01='L' and a1p05='1101' and a1p16='L0100') and not (a1p01='L' and a1p05='6129') and (a1p01='" & Text21 & "' or (a1p01='1' and a1p05='110602' and a1p17='TT999999000')) and not (a1p05='2492' and a1p23 is not null)", adoTaie, adOpenStatic, adLockReadOnly
   
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
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strDate = MaskEdBox1.Text
   Else
      strDate = MsgText(601)
   End If
   tool3_enabled
   Frmacc1151.Show
   Screen.MousePointer = vbDefault
   Me.Hide
   
   'Added by Morgan 2013/12/25
   If m_Rebuild = True Then
      m_Rebuild = False
      Unload Frmacc1151
   End If
   'end 2013/12/25
   
End Sub

Private Sub Command2_Click()
   If Adodc1.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc1.Recordset.Fields("a1p22").Value) = False Then
         If adoquery.State = adStateOpen Then adoquery.Close       'Added by Lydia 2024/11/28
         adoquery.CursorLocation = adUseClient
         'Modified by Morgan 2013/12/19 一張收款單有可能有兩張傳票號
         'adoQuery.Open "select ax210 from acc021 where ax201 = '" & Adodc1.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc1.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         adoquery.Open "select ax210 from acc021 where (AX201,AX202) in (select a1p01,a1p22 from acc1p0 where a1p04='" & Adodc1.Recordset.Fields("a1p04").Value & "') and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(155), , MsgText(5)
            Text4.SetFocus
            adoquery.Close
            Exit Sub
         End If
         adoquery.Close
      End If
   End If
   AdodcDelete
   SumShow
End Sub
'Add by Morgan 2005/8/12
Public Sub RefreshData()
   
   Acc0l0Refresh
   If adoacc0l0.RecordCount <> 0 Then
      FormShow
      AdodcRefresh
      SumShow
      RecordShow
   End If
   
   'Add by Morgan 2005/9/29 檢查借貸平衡
   If Frmacc1150.CreDebCheck <> MsgText(602) Then
      MsgBox MsgText(11), , MsgText(5)
      'Added by Morgan 2013/6/28
      Text2.Enabled = False
      Command3.Enabled = False
      'end 2013/6/28
      Exit Sub
   End If
   '2005/9/29 end
   
End Sub
'Modify by Morgan 2005/8/12 程式碼抽出改成call sub 以方便共用
Private Sub Command3_Click()
   RefreshData
End Sub

Private Sub Command3_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command3_Click
         Exit Sub
   End Select
   KeyDefine KeyCode
End Sub

Private Sub Command4_Click()
   If Text2 <> "" Then
      strExc(0) = "select a1p01 COMP,sum(a1p07) AMT1,sum(a1p08) AMT2,sum(a1p07)-sum(a1p08) AMT3 from acc1p0 where a1p04='" & Text2 & "' group by a1p01"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With Frmacc1154
            Set .Adodc1.Recordset = RsTemp
            .Show vbModal
         End With
         strFormName = Name
      End If
   End If
End Sub

Private Sub Command5_Click()
   Dim strNewComp As String
   
   'Added by Morgan 2014/1/27
   If Adodc1.Recordset.RecordCount > 0 Then
      If Not IsNull(Adodc1.Recordset.Fields("a1p22").Value) Then
         MsgBox "已有傳票號碼，不可更改主要公司別!!", vbExclamation
         Exit Sub
      End If
   End If
   'end 2014/1/27
   
   Text21.Tag = Text21
   Do
      'Modified by Morgan 2020/4/13 +L
      strExc(0) = InputBox("請輸入主要公司別!! ( 1 or J or L )", , Frmacc1150.Text21)
      If strExc(0) = "" Then
         Exit Do
      ElseIf strExc(0) <> "1" And UCase(strExc(0)) <> "J" And UCase(strExc(0)) <> "L" Then
         MsgBox "只可輸入 1 或 J 或 L", vbCritical
      Else
         Exit Do
      End If
   Loop
   strNewComp = UCase(strExc(0))
   If strNewComp <> "" And strNewComp <> Text21 Then
      'Added by Morgan 2023/8/25
      Set RsTemp = Adodc1.Recordset.Clone
      If RsTemp.RecordCount > 0 Then
         RsTemp.MoveFirst
         If Val(RsTemp("a1p03")) = 1 Then
            If strNewComp <> RsTemp("a1p01") Then
               MsgBox "主要公司別必須和第1筆分錄的公司別相同！", vbCritical
               Exit Sub
            End If
         End If
      End If
      'end 2023/8/25
   
      Text21 = strNewComp
      If Command1.Enabled = True Then
         Set RsTemp = Adodc1.Recordset.Clone
         RsTemp.Find "a1p08>0"
         If Not RsTemp.EOF Then
            m_Rebuild = True
            m_Activated = False
            Command1.Value = True
            If m_Rebuild = False Then
               strSql = "update acc0l0 set a0l05='" & Text21 & "' where a0l01='" & Text2 & "'"
               adoTaie.Execute strSql, intI
               If m_Activated = False Then
                  Form_Activate
               End If
            Else
               Text21 = Text21.Tag
               m_Rebuild = False
            End If
         End If
      'Added by Morgan 2023/8/25
      ElseIf Not adoacc0l0.EOF Then
         adoacc0l0.Fields("a0l05").Value = Text21
         adoacc0l0.UpdateBatch
      'end 2023/8/25
      End If
   End If
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   strSerialNo = Adodc1.Recordset.Fields("a1p03").Value
   AdodcShow
End Sub

'Added by Morgan 2013/12/17
Public Sub AutoProcess()
   '新增
   strTrackMode = "" 'Added by Morgan 2021/12/22
   KeyDefine vbKeyF2
   If Text2 <> "" Then
      AutoAdd1
      Frmacc1150_Save '要放在AutoAdd1後面執行備註才會存到acc0l0
      AutoAdd2
      AdodcRefresh
      SumShow
      m_AutoRun = True
   End If
End Sub
'新增 acc0m0 收據資料
Private Sub AutoAdd2()
   Dim stVTB As String
   '批次新增 1u0
   'Modified by Morgan 2014/5/16 沒有值要放 0 否則有報表會錯
   strSql = "insert into acc1u0(a1u01,a1u02,a1u03,a1u04,a1u05,a1u06,a1u07,a1u08,a1u09,a1u10)" & _
      " select '" & Text2 & "',axd04,axd05,axd06,axd07,axd08,0,0,0,0 from acc441 where AXD01='" & m_A4401 & "' and AXD02=" & m_A4402 & " and AXD03=" & m_A4403
   adoTaie.Execute strSql, intI
   
   '批次新增 0m0
   stVTB = "select axd04 C1,sum(axd06) C2,sum(axd07) C3,sum(axd08) C4 from acc441" & _
      " where AXD01='" & m_A4401 & "' and AXD02=" & m_A4402 & " and AXD03=" & m_A4403 & _
      " group by axd04"
      
   'Modified by Morgan 2015/1/21 若分次收款時扣繳年度必須相同
   strSql = "insert into ACC0M0(a0m01,a0m02,a0m04,a0m05,a0m06" & _
      ",a0m07,a0m08,a0m09,a0m10,a0m15,a0m16,a0m17,a0m18,a0m19)" & _
      " select '" & Text2 & "',C1,max(C2),max(C3),max(C4),nvl(min(a0m07)," & Val(MaskEdBox1) & ")" & _
      ",nvl(sum(a0m04),0),nvl(sum(a0m05),0),'" & m_A4401 & "',max(a0k06),max(a0k07)," & strSrvDate(2) & ",to_char(sysdate,'hh24miss')" & _
      ",'" & strUserNum & "' from (" & stVTB & "),acc0k0,acc0m0 where a0k01(+)=C1 and a0m02(+)=C1 group by C1"
   adoTaie.Execute strSql, intI
   
   'add by sonia 2020/5/20 法律所案源之二事務所之間收據也要同時收款
   '新增 1u0
   'Modified by Morgan 2023/11/17 修正法律案拆收據智慧所收據會重複計算問題 Ex:L006723000(AB2042462)-E11222711,E11223528->AB2042384
   'strSql = "insert into acc1u0 (a1u01,a1u02,a1u03,a1u04,a1u05,a1u06,a1u07,a1u08,a1u09,a1u10)" & _
      " SELECT '" & Text2 & "',A0J13,A0J01,SUM(A0J09),SUM(A0J10),sum(decode(a0j07,'Y',nvl(a0j09,0)+nvl(a0j10,0),nvl(a0j09,0)))/10,0,0,0,0 FROM " & _
      "(SELECT AXD04,AXD05,AXD06,L1.LOS01,L2.LOS06,A0J13,A0J01,A0J09-NVL(A1U07,0) A0J09,A0J10-NVL(A1U08,0) A0J10,A0K11,a0j07 FROM ACC441,LAWOFFICESOURCE L1,LAWOFFICESOURCE L2,ACC0J0,ACC0K0," & _
      " (SELECT A1U03,SUM(A1U07) A1U07,SUM(A1U08) A1U08 FROM ACC1U0 WHERE A1U03 IN " & _
      " (SELECT DISTINCT A0J01 FROM ACC441,LAWOFFICESOURCE L1,LAWOFFICESOURCE L2,ACC0J0 WHERE AXD01='" & m_A4401 & "' AND AXD02=" & m_A4402 & " AND AXD03=" & m_A4403 & " " & _
      "         AND AXD05=L1.LOS01(+) AND AXD05=L2.LOS06(+) AND L1.LOS01||L2.LOS01 IS NOT NULL AND DECODE(SUBSTR(L1.LOS02,1,1),'B',L1.LOS06,'C',L1.LOS06,L2.LOS10)=A0J01(+) AND A0J01 IS NOT NULL) GROUP BY A1U03)" & _
      " WHERE AXD01='" & m_A4401 & "' AND AXD02=" & m_A4402 & " AND AXD03=" & m_A4403 & " AND AXD05=L1.LOS01(+) AND AXD05=L2.LOS06(+) AND L1.LOS01||L2.LOS01 IS NOT NULL " & _
      " AND DECODE(SUBSTR(L1.LOS02,1,1),'B',L1.LOS06,'C',L1.LOS06,L2.LOS10)=A0J01(+) AND A0J01 IS NOT NULL AND A0J01=A1U03(+) AND A0J13=A0K01(+)" & _
      " ) GROUP BY A0J13,A0J01"
   strSql = "insert into acc1u0 (a1u01,a1u02,a1u03,a1u04,a1u05,a1u06,a1u07,a1u08,a1u09,a1u10)" & _
      " SELECT '" & Text2 & "',A0J13,A0J01,A0J09-A1U07,A0J10-A1U09,decode(a0j07,'Y',nvl(a0j09,0)-A1U07+nvl(a0j10,0)-A1U09,nvl(a0j09,0)-A1U07)/10,0,0,0,0" & _
      " FROM (SELECT  a0j13 X1,a0j01 X2,NVL(SUM(A1U07),0) A1U07,NVL(SUM(A1U08),0) A1U09" & _
      " FROM acc0j0,ACC1U0 WHERE a0j13 IN (" & m_TTRcpSQL & ") AND A1U02(+)=a0j13 and A1U03(+)=a0j01 group by a0j13,a0j01" & _
      ") X,ACC0J0 WHERE A0J13(+)=X1 and A0J01(+)=X2"
    'end 2023/11/17
   adoTaie.Execute strSql, intI
   
   '新增 0m0
   'Modified by Morgan 2023/11/17 修正法律案拆收據智慧所收據會重複計算問題 Ex:L006723000(AB2042462)-E11222711,E11223528->AB2042384
   'strSql = "insert into ACC0M0 (a0m01,a0m02,a0m04,a0m05,a0m06" & _
      ",a0m07,a0m08,a0m09,a0m10,a0m15,a0m16,a0m17,a0m18,a0m19)" & _
      " select '" & Text2 & "',A0J13,SUM(A0J09),SUM(A0J10),sum(decode(a0j07,'Y',nvl(a0j09,0)+nvl(a0j10,0),nvl(a0j09,0)))/10," & Val(MaskEdBox1) & "" & _
      ",0,0,'" & m_A4401 & "',SUM(A0J09),SUM(A0J10)," & strSrvDate(2) & ",to_char(sysdate,'hh24miss')" & _
      ",'" & strUserNum & "' FROM " & _
      "(SELECT AXD04,AXD05,AXD06,L1.LOS01,L2.LOS06,A0J13,A0J01,A0J09-NVL(A1U07,0) A0J09,A0J10-NVL(A1U08,0) A0J10,A0K11,A0J07 FROM ACC441,LAWOFFICESOURCE L1,LAWOFFICESOURCE L2,ACC0J0,ACC0K0," & _
      " (SELECT A1U03,SUM(A1U07) A1U07,SUM(A1U08) A1U08 FROM ACC1U0 WHERE A1U03 IN " & _
      " (SELECT DISTINCT A0J01 FROM ACC441,LAWOFFICESOURCE L1,LAWOFFICESOURCE L2,ACC0J0 WHERE AXD01='" & m_A4401 & "' AND AXD02=" & m_A4402 & " AND AXD03=" & m_A4403 & " " & _
      "         AND AXD05=L1.LOS01(+) AND AXD05=L2.LOS06(+) AND L1.LOS01||L2.LOS01 IS NOT NULL AND DECODE(SUBSTR(L1.LOS02,1,1),'B',L1.LOS06,'C',L1.LOS06,L2.LOS10)=A0J01(+) AND A0J01 IS NOT NULL) GROUP BY A1U03)" & _
      " WHERE AXD01='" & m_A4401 & "' AND AXD02=" & m_A4402 & " AND AXD03=" & m_A4403 & " AND AXD05=L1.LOS01(+) AND AXD05=L2.LOS06(+) AND L1.LOS01||L2.LOS01 IS NOT NULL " & _
      " AND DECODE(SUBSTR(L1.LOS02,1,1),'B',L1.LOS06,'C',L1.LOS06,L2.LOS10)=A0J01(+) AND A0J01 IS NOT NULL AND A0J01=A1U03(+) AND A0J13=A0K01(+)" & _
      " ) GROUP BY A0J13"
   strSql = "insert into ACC0M0 (a0m01,a0m02,a0m04,a0m05,a0m06" & _
      ",a0m07,a0m08,a0m09,a0m10,a0m15,a0m16,a0m17,a0m18,a0m19)" & _
      " select '" & Text2 & "',A0J13,SUM(A0J09-A1U07),SUM(A0J10-A1U09),sum(decode(a0j07,'Y',nvl(a0j09,0)-A1U07+nvl(a0j10,0)-A1U09,nvl(a0j09,0)-A1U07))/10," & Val(MaskEdBox1) & "" & _
      ",0,0,'" & m_A4401 & "',SUM(A0J09),SUM(A0J10)," & strSrvDate(2) & ",to_char(sysdate,'hh24miss'),'" & strUserNum & "'" & _
      " FROM (SELECT  a0j13 X1,NVL(SUM(A1U07),0) A1U07,NVL(SUM(A1U08),0) A1U09" & _
      " FROM acc0j0,ACC1U0 WHERE a0j13 IN (" & m_TTRcpSQL & ") AND A1U02(+)=a0j13 and A1U03(+)=a0j01 group by a0j13" & _
      ") X,ACC0J0 WHERE A0J13(+)=X1 GROUP BY A0J13"
   'end 2023/11/17
   adoTaie.Execute strSql, intI
   'end 2020/5/20
End Sub


'新增借方科目
Private Sub AutoAdd1()
   Dim A1P03 As String, A1P05 As String, A1P07 As String, A1P09 As String, A1P10 As String, A1P11 As String, A1P12 As String, A1P14 As String, A1P15 As String, A1P18 As String
   Dim bolDone As Boolean
   Dim stMsg As String  'add by sonia 2020/4/28
   Dim A1P16 As String, LOS02 As String, TTMan As String 'Added by Morgan 2021/1/18
   Dim A1P17 As String 'Added by Morgan 2023/8/17
   Dim bolNew As Boolean 'Added by Morgan 2025/6/16
   
   A1P03 = 0
   A1P18 = Val(FCDate(MaskEdBox1.Text))
   
   'strExc(0) = "select * from acc440,acc230,staff where a4401='" & m_A4401 & "' and a4402=" & m_A4402 & " and a4403=" & m_A4403 & " and a2301(+)=a4421 and st01(+)=A4401"
   'Modified by Morgan 2015/6/17 電匯日期改抓簽收輸入日期 --辜
   'decode(A2324,null,'',to_char(A2324,'yyyymmdd')-19110000) EDate
   'Modified by Morgan 2017/11/27 所別先判斷簽收輸入人員的所別--辜 Ex.E10626506南所業務的款項但高所簽收
   'Modified by Morgan 2025/6/16 所別改依簽收時點選(A23205)--瑞婷
   'strExc(0) = "select A.*,B.*,C.*,D.*,NVL(J.st06,E.st06) st06,sn01,G.cu04 CName1,H.cu04 CName2,decode(a2325,'','',a2325||'/'||a2326||'/'||a2328||'/'||a0g02) Memo" & _
      " from ( select axd01,axd02,axd03,min(axd04) axd04 from acc441" & _
      " where axd01='" & m_A4401 & "' and axd02=" & m_A4402 & " and axd03=" & m_A4403 & "" & _
      " group by axd01,axd02,axd03) A,acc440 B,acc0k0 C,acc230 D,staff E" & _
      ",salesno F ,customer G,customer H,acc0g0 I,staff J" & _
      " where a4401(+)=axd01 and a4402(+)=axd02 and a4403(+)=axd03" & _
      " and a0k01(+)=axd04 and a2301(+)=a4421 and E.st01(+)=A4401 and sn02(+)=A4401" & _
      " and G.cu01(+)=substr(A2304,1,8) and G.cu02(+)=substr(A2304,9)" & _
      " and H.cu01(+)=substr(a0k03,1,8) and H.cu02(+)=substr(a0k03,9) and a0g01(+)=a2327 and J.st01(+)=A2311"
   strExc(0) = "select A.*,B.*,C.*,D.*,a2305 st06,sn01,G.cu04 CName1,H.cu04 CName2,decode(a2325,'','',a2325||'/'||a2326||'/'||a2328||'/'||a0g02) Memo" & _
      " from ( select axd01,axd02,axd03,min(axd04) axd04 from acc441" & _
      " where axd01='" & m_A4401 & "' and axd02=" & m_A4402 & " and axd03=" & m_A4403 & "" & _
      " group by axd01,axd02,axd03) A,acc440 B,acc0k0 C,acc230 D,staff E" & _
      ",salesno F ,customer G,customer H,acc0g0 I,staff J" & _
      " where a4401(+)=axd01 and a4402(+)=axd02 and a4403(+)=axd03" & _
      " and a0k01(+)=axd04 and a2301(+)=a4421 and E.st01(+)=A4401 and sn02(+)=A4401" & _
      " and G.cu01(+)=substr(A2304,1,8) and G.cu02(+)=substr(A2304,9)" & _
      " and H.cu01(+)=substr(a0k03,1,8) and H.cu02(+)=substr(a0k03,9) and a0g01(+)=a2327"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      'Added by Morgan 2021/1/18 案源收款，L公司的業務對沖固定放 L0100 --辜
      A1P16 = m_A4401
      If Text21 = "L" Then
         If ChkIsLawCase(.Fields("a0k01")) = True Then
            A1P16 = "L0100"
         End If
      End If
      'end 2021/1/18
      
      '摘要=智權人員備註+出納備註
      Text1 = "" & .Fields("A4412")
      If Not IsNull(.Fields("A4415")) Then
         Text1 = Text1 & IIf(Text1 <> "", ";", "") & .Fields("A4415")
      End If
      
      'Added by Morgan 2023/10/31
      '溢收款退客戶備註帶出退客戶--瑞婷
      If .Fields("A4425") = "2" And .Fields("A4410") > 0 Then
         Text1 = Text1 & IIf(Text1 <> "", ";", "") & "退客戶"
      End If
      'end 2023/10/31
      
      A1P15 = .Fields("a0k03")
      '票據
      If .Fields("A4405") > 0 Then
         bolDone = False
         A1P03 = Format(Val(A1P03) + 1, "000")
         A1P05 = "113001"

         If .Fields("A2306") > 0 Then
            A1P07 = .Fields("A2306")
            A1P14 = "" & .Fields("Memo")
            A1P09 = "" & .Fields("A2326")
            A1P10 = "" & .Fields("A2327")
            A1P11 = "" & .Fields("A2328")
            A1P12 = "" & .Fields("A2325")
            strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p09,a1p10,a1p11,a1p12,a1p13,a1p14,a1p16,a1p18) values('" & Text21 & "','A','" & A1P03 & "','" & Text2 & "','" & A1P05 & "','TOT'," & A1P07 & ",0,'" & A1P09 & "','" & A1P10 & "','" & A1P11 & "'," & CNULL(A1P12, True) & ",'1','" & ChgSQL(A1P14) & "','" & A1P16 & "'," & A1P18 & ")"
            adoTaie.Execute strSql, intI
            bolDone = True
            
            'Added by Morgan 2025/6/13
            '來自簽收作業的支票若超過規定於按【收款明細】時彈提醒
            If A1P12 <> "" Then
               'Modified by Morgan 2025/6/18 規定有修正，改用函數
               'strExc(1) = CompDate(1, 2, DBDATE(.Fields("A2302")))
               strExc(1) = PUB_GetCheckMaxDate(.Fields("A2302"))
               If DBDATE(A1P12) > strExc(1) Then
                  m_bolAlert = True
               End If
            End If
            'end 2025/6/13
         End If
         
         '多筆簽收
         If Not IsNull(.Fields("A4427")) Then
            strExc(0) = "select A.*,decode(a2325,'','',a2325||'/'||a2326||'/'||a2328||'/'||a0g02) Memo from acc230 A,acc0g0 where a2301 in ('" & Replace(.Fields("A4427"), ";", "','") & "') and A2306>0 and a0g01(+)=a2327"
            intI = 1
            Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               Do While Not adoRecordset.EOF
                  A1P07 = adoRecordset.Fields("A2306")
                  A1P14 = "" & adoRecordset.Fields("Memo")
                  A1P09 = "" & adoRecordset.Fields("A2326")
                  A1P10 = "" & adoRecordset.Fields("A2327")
                  A1P11 = "" & adoRecordset.Fields("A2328")
                  A1P12 = "" & adoRecordset.Fields("A2325")
                  A1P03 = Format(Val(A1P03) + 1, "000")
                  strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p09,a1p10,a1p11,a1p12,a1p13,a1p14,a1p16,a1p18) values('" & Text21 & "','A','" & A1P03 & "','" & Text2 & "','" & A1P05 & "','TOT'," & A1P07 & ",0,'" & A1P09 & "','" & A1P10 & "','" & A1P11 & "'," & CNULL(A1P12, True) & ",'1','" & ChgSQL(A1P14) & "','" & A1P16 & "'," & A1P18 & ")"
                  adoTaie.Execute strSql, intI
                  bolDone = True
                  
                  'Added by Morgan 2025/6/13
                  '來自簽收作業的支票若超過(輸入日期+2個月)於按【收款明細】時彈提醒
                  If A1P12 <> "" Then
                     'Modified by Morgan 2025/8/28 規定有修正，改用函數
                     'strExc(1) = CompDate(1, 2, DBDATE(adoRecordset.Fields("A2302")))
                     strExc(1) = PUB_GetCheckMaxDate(.Fields("A2302"))
                     If DBDATE(A1P12) > strExc(1) Then
                        m_bolAlert = True
                     End If
                  End If
                  'end 2025/6/13
                  
                  adoRecordset.MoveNext
               Loop
            End If
         End If
         
         If Not bolDone Then
            A1P07 = .Fields("A4405")
            A1P14 = Text1
            strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p09,a1p10,a1p11,a1p12,a1p13,a1p14,a1p16,a1p18) values('" & Text21 & "','A','" & A1P03 & "','" & Text2 & "','" & A1P05 & "','TOT'," & A1P07 & ",0,'" & A1P09 & "','" & A1P10 & "','" & A1P11 & "'," & CNULL(A1P12, True) & ",'1','" & ChgSQL(A1P14) & "','" & A1P16 & "'," & A1P18 & ")"
            adoTaie.Execute strSql, intI
         End If

      End If
      
      '電匯
      If .Fields("A2318") > 0 Then
         A1P07 = .Fields("A2318")
         A1P03 = Format(Val(A1P03) + 1, "000")
         'Modified by Morgan 2015/7/17
         '若收款月份大於簽收輸入月份科目改為暫收2401
         If Val(FCDate(MaskEdBox1.Text)) \ 100 > .Fields("A2302") \ 100 Then
            A1P05 = "2401"
         Else
            A1P05 = "" & .Fields("A2322")
         End If
         'end 2015/7/17
         'modify by sonia 2017/12/12 +A2330
         'Modified by Morgan 2021/1/19 案源收款，L公司摘要不帶業務
         A1P14 = IIf(A1P16 = "L0100", "", .Fields("sn01") & "/") & Left(.Fields("a0k04"), 10) & " " & .Fields("A2302") & " " & .Fields("A2330")
         strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p15,a1p16,a1p18) values('" & Text21 & "','A','" & A1P03 & "','" & Text2 & "','" & A1P05 & "','TOT'," & A1P07 & ",0,'" & ChgSQL(A1P14) & "','" & A1P15 & "','" & A1P16 & "'," & A1P18 & ")"
         adoTaie.Execute strSql, intI
      End If
      
      If Not IsNull(.Fields("A4427")) Then
         strExc(0) = "select A2318,A2322,A2302 from acc230 where a2301 in ('" & Replace(.Fields("A4427"), ";", "','") & "') and a2318>0"
         intI = 1
         Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Do While Not adoRecordset.EOF
               A1P07 = adoRecordset.Fields("A2318")
               A1P03 = Format(Val(A1P03) + 1, "000")
               '若收款月份大於簽收輸入月份科目改為暫收2401
               If Val(FCDate(MaskEdBox1.Text)) \ 100 > adoRecordset.Fields("A2302") \ 100 Then
                  A1P05 = "2401"
               Else
                  A1P05 = adoRecordset.Fields("A2322")
               End If
               'modify by sonia 2017/12/12 +A2330
               'Modified by Morgan 2021/1/19 案源收款，L公司摘要不帶業務
               A1P14 = IIf(A1P16 = "L0100", "", .Fields("sn01") & "/") & Left(.Fields("a0k04"), 10) & " " & adoRecordset.Fields("A2302") & " " & .Fields("A2330")
               strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p15,a1p16,a1p18) values('" & Text21 & "','A','" & A1P03 & "','" & Text2 & "','" & A1P05 & "','TOT'," & A1P07 & ",0,'" & ChgSQL(A1P14) & "','" & A1P15 & "','" & A1P16 & "'," & A1P18 & ")"
               adoTaie.Execute strSql, intI
               adoRecordset.MoveNext
            Loop
         End If
      End If
      
      '現金
      If .Fields("A4408") > 0 Then
         'Added by Morgan 2025/6/16
         strExc(0) = "select * from acc010 where a0101='191101'"
         intI = 1
         Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
           bolNew = True
         End If
         'end 2025/6/16
    
         bolDone = False
         A1P03 = Format(Val(A1P03) + 1, "000")
         '中所
         If .Fields("st06") = "2" Then
            If bolNew Then
               A1P05 = "191101"
            Else
               A1P05 = "1911"
            End If
         ElseIf .Fields("st06") = "3" Then
            If bolNew Then
               A1P05 = "191201"
            Else
               A1P05 = "1912"
            End If
         ElseIf .Fields("st06") = "4" Then
            If bolNew Then
               A1P05 = "191301"
            Else
               A1P05 = "1913"
            End If
         Else
            A1P05 = "1101"
         End If
         'Modified by Morgan 2021/1/19 案源收款，L公司摘要不帶業務
         A1P14 = IIf(A1P16 = "L0100", "", .Fields("sn01") & "/") & Left(.Fields("a0k04"), 10)
         
         'Added by Morgan 2015/7/17
         '簽收
         If .Fields("A2317") > 0 Then
            A1P07 = .Fields("A2317")
            '摘要同電匯
            A1P14 = A1P14 & " " & .Fields("A2302")
            '若收款月份大於簽收輸入月份科目改為暫收2401
            If Val(FCDate(MaskEdBox1.Text)) \ 100 > .Fields("A2302") \ 100 Then
               A1P05 = "2401"
            End If
            strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p15,a1p16,a1p18) values('" & Text21 & "','A','" & A1P03 & "','" & Text2 & "','" & A1P05 & "','TOT'," & A1P07 & ",0,'" & ChgSQL(A1P14) & "','" & A1P15 & "','" & A1P16 & "'," & A1P18 & ")"
            adoTaie.Execute strSql, intI
            bolDone = True
         End If
         'end 2015/7/17
         
         'Added by Morgan 2015/7/23 多筆簽收
         If Not IsNull(.Fields("A4427")) Then
            'Modified by Morgan 2017/11/27 所別先判斷簽收輸入人員的所別--辜 Ex.E10626506南所業務的款項但高所簽收
            'Modified by Morgan 2025/6/16 所別改依簽收時點選(A23205)--瑞婷
            'strExc(0) = "select A.*,st06 from acc230 A,staff where a2301 in ('" & Replace(.Fields("A4427"), ";", "','") & "') and A2317>0 and st01(+)=A2311"
            strExc(0) = "select A.*,a2305 st06 from acc230 A where a2301 in ('" & Replace(.Fields("A4427"), ";", "','") & "') and A2317>0"
            intI = 1
            Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               Do While Not adoRecordset.EOF
                  A1P07 = adoRecordset.Fields("A2317")
                  A1P03 = Format(Val(A1P03) + 1, "000")
                  'modify by sonia 2017/12/12 +A2330
                  'Modified by Morgan 2021/1/19 案源收款，L公司摘要不帶業務
                  A1P14 = IIf(A1P16 = "L0100", "", .Fields("sn01") & "/") & Left(.Fields("a0k04"), 10) & " " & adoRecordset.Fields("A2302") & " " & .Fields("A2330")
                  '若收款月份大於簽收輸入月份科目改為暫收2401
                  If Val(FCDate(MaskEdBox1.Text)) \ 100 > adoRecordset.Fields("A2302") \ 100 Then
                     A1P05 = "2401"
                  ElseIf adoRecordset.Fields("st06") = "2" Then
                     If bolNew Then
                        A1P05 = "191101"
                     Else
                        A1P05 = "1911"
                     End If
                  ElseIf adoRecordset.Fields("st06") = "3" Then
                     If bolNew Then
                        A1P05 = "191201"
                     Else
                        A1P05 = "1912"
                     End If
                  ElseIf adoRecordset.Fields("st06") = "4" Then
                     If bolNew Then
                        A1P05 = "191301"
                     Else
                        A1P05 = "1913"
                     End If
                  Else
                     A1P05 = "1101"
                  End If
                  strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p15,a1p16,a1p18) values('" & Text21 & "','A','" & A1P03 & "','" & Text2 & "','" & A1P05 & "','TOT'," & A1P07 & ",0,'" & ChgSQL(A1P14) & "','" & A1P15 & "','" & A1P16 & "'," & A1P18 & ")"
                  adoTaie.Execute strSql, intI
                  bolDone = True
                  adoRecordset.MoveNext
               Loop
            End If
         End If
         'end 2015/7/23
         
         If Not bolDone Then
            A1P07 = .Fields("A4408")
            strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p15,a1p16,a1p18) values('" & Text21 & "','A','" & A1P03 & "','" & Text2 & "','" & A1P05 & "','TOT'," & A1P07 & ",0,'" & ChgSQL(A1P14) & "','" & A1P15 & "','" & A1P16 & "'," & A1P18 & ")"
            adoTaie.Execute strSql, intI
            bolDone = True
         End If
         
      End If
      '抵暫收款
      If .Fields("A4409") > 0 Then
         A1P07 = .Fields("A4409")
         A1P03 = Format(Val(A1P03) + 1, "000")
         A1P05 = "2401"
         A1P14 = Text1
         
         strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p16,a1p18) values('" & Text21 & "','A','" & A1P03 & "','" & Text2 & "','" & A1P05 & "','TOT'," & A1P07 & ",0,'" & ChgSQL(A1P14) & "','" & A1P16 & "'," & A1P18 & ")"
         adoTaie.Execute strSql, intI
      End If
      
      'Added by Morgan 2015/7/15
      '其他
      If .Fields("A4430") > 0 Then
         A1P07 = .Fields("A4430")
         A1P03 = Format(Val(A1P03) + 1, "000")
         A1P05 = "611602"
         A1P14 = "" & .Fields("A4431")
          strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p16,a1p18) values('" & Text21 & "','A','" & A1P03 & "','" & Text2 & "','" & A1P05 & "','SAL'," & A1P07 & ",0,'" & ChgSQL(A1P14) & "','" & A1P16 & "'," & A1P18 & ")"
         adoTaie.Execute strSql, intI
      End If
      'end 2015/7/15
      
      '手續費
      If .Fields("A4411") > 0 Then
         A1P07 = .Fields("A4411")
         A1P03 = Format(Val(A1P03) + 1, "000")
         'Added by Morgan 2014/1/9
         If .Fields("A4405") > 0 Then
            'Modified by Morgan 2021/1/19 案源收款，L公司摘要不帶業務
            A1P14 = IIf(A1P16 = "L0100", "", .Fields("sn01") & "/") & .Fields("CName2") & "/寄支票"
            A1P05 = "611001"
         Else
            'Modified by Morgan 2021/1/19 案源收款，L公司摘要不帶業務
            A1P14 = IIf(A1P16 = "L0100", "", .Fields("sn01") & "/") & .Fields("CName1") & "/匯費"
         'end 2014/1/9
            A1P05 = "611301"
         End If 'Added by Morgan 2014/1/9
         'modify by sonia 2020/11/9 L公司部門改L
         'strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p16,a1p18) values('" & Text21 & "','A','" & A1P03 & "','" & Text2 & "','" & A1P05 & "','TOT'," & A1P07 & ",0,'" & ChgSQL(A1P14) & "','" & A1P16 & "'," & A1P18 & ")"
         strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p16,a1p18) values('" & Text21 & "','A','" & A1P03 & "','" & Text2 & "','" & A1P05 & "','" & IIf(Text21 = "L", "L", "TOT") & "'," & A1P07 & ",0,'" & ChgSQL(A1P14) & "','" & A1P16 & "'," & A1P18 & ")"
         adoTaie.Execute strSql, intI
      End If
      
      '外幣
      If .Fields("A4426") > 0 Then
         A1P07 = .Fields("A4426")
         A1P03 = Format(Val(A1P03) + 1, "000")
         A1P05 = "110208"
         'Modified by Morgan 2021/1/19 案源收款，L公司摘要不帶業務
         A1P14 = IIf(A1P16 = "L0100", "", .Fields("sn01") & "/") & Left(.Fields("a0k04"), 10)
         strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p15,a1p16,a1p18) values('" & Text21 & "','A','" & A1P03 & "','" & Text2 & "','" & A1P05 & "','TOT'," & A1P07 & ",0,'" & ChgSQL(A1P14) & "','" & A1P15 & "','" & A1P16 & "'," & A1P18 & ")"
         adoTaie.Execute strSql, intI
      End If
      
      If .Fields("A4422") > 0 Then
         MsgBox "本次繳款記錄有補扣繳 " & Format(.Fields("A4422"), "#,###") & " 元，請留意！", vbExclamation
      End If
      End With
   End If

   'Added by Morgan 2023/11/17
   '案源智慧所收據號
   m_TTRcpSQL = "SELECT distinct A0J13 from ACC441,ACC0K0,CASEPROGRESS,LAWOFFICESOURCE,ACC0J0" & _
      " WHERE AXD01='" & m_A4401 & "' AND AXD02=" & m_A4402 & " AND AXD03=" & m_A4403 & _
      " AND A0K01(+)=AXD04 AND A0K11='L' and cp09(+)=axd05 and los15(+)=cp162 and los02 in ('A1','A2')" & _
      " and a0j01(+)=los10"
   'end 2023/11/17
      
   'Added by Morgan 2021/3/8 A1,A2類案源法律所收據收款自動新增智慧所現金科目
   'Modified by Morgan 2023/8/23 現金科目改為 110602 瑞興銀行乙存(智慧所)
   'Modified by Morgan 2023/11/17 修正法律所拆收據問題資料重複問題,加扣除銷帳金額並考慮若多張收據收款加跑迴圈產生分錄
   'strExc(0) = "select 0.9*(j1.A0J09+j1.A0J10) a1p07,c2.cp13,getcp10desc(c2.cp01,c2.cp10,j1.a0j04) cp10N,a0k04" & _
      ",c1.cp01||c1.cp02||decode(c1.cp04,'00',decode(c1.cp03,'0','','-'||c1.cp03),'-'||c1.cp04) CaseNo" & _
      " FROM ACC441,acc0k0,caseprogress c1,LAWOFFICESOURCE,acc0j0 j1,caseprogress c2" & _
      " WHERE AXD01='" & m_A4401 & "' AND AXD02=" & m_A4402 & " AND AXD03=" & m_A4403 & _
      " and a0k01(+)=axd04 and a0k11='L' and c1.cp09(+)=axd05 and los15(+)=c1.cp162 and los02 in ('A1','A2')" & _
      " and j1.a0j01(+)=los10 and c2.cp09(+)=los10 and j1.a0j13 is not null"
   
   '法律所+智慧所收文號
   strExc(1) = "select distinct los06,los10,a0k04 from ACC441,acc0k0,caseprogress,LAWOFFICESOURCE" & _
      " WHERE AXD01='" & m_A4401 & "' AND AXD02=" & m_A4402 & " AND AXD03=" & m_A4403 & _
      " and a0k01(+)=axd04 and a0k11='L' and cp09(+)=axd05 and los15(+)=cp162 and los02 in ('A1','A2')"
   
   '智慧所收文號銷帳金額
   strExc(2) = "SELECT a0j13 X1,a0j01 X2,SUM(A1U07) A1U07,SUM(A1U08) A1U09" & _
      " FROM acc0j0,ACC1U0 WHERE a0j13 IN (" & m_TTRcpSQL & ") AND A1U02(+)=a0j13 and A1U03(+)=a0j01 group by a0j13,a0j01"
   
   strExc(0) = "select 0.9*(nvl(A0J09,0)+nvl(A0J10,0)-nvl(a1u07,0)-nvl(a1u09,0)) a1p07,c2.cp13,getcp10desc(c2.cp01,c2.cp10,a0j04) cp10N,a0k04" & _
      ",c1.cp01||c1.cp02||decode(c1.cp04,'00',decode(c1.cp03,'0','','-'||c1.cp03),'-'||c1.cp04) CaseNo,c2.cp01||c2.cp02||c2.cp03||c2.cp04 TTNo" & _
      " FROM (" & strExc(1) & ") A,caseprogress c1,acc0j0,(" & strExc(2) & ") X,caseprogress c2" & _
      " where c1.cp09(+)=los06 and a0j01(+)=los10 and a0j13 is not null and X1(+)=a0j13 and X2(+)=a0j01 and c2.cp09(+)=a0j01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF 'Added by Morgan 2023/11/17
         A1P03 = Format(Val(A1P03) + 1, "000")
         A1P07 = Val("" & .Fields("a1p07"))
         A1P14 = "法律所/" & Left(.Fields("a0k04"), 6) & .Fields("CaseNo") & "/" & .Fields("cp10N")
         A1P15 = "X82357000"
         A1P16 = .Fields("cp13")
         'Modified by Morgan 2023/11/17
         'A1P17 = "TT999999000" 'Added by Morgan 2023/8/17
         A1P17 = .Fields("TTNo")
         'end 2023/11/17
         strSql = "insert into acc1p0(a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p15,a1p16,a1p17,a1p18) values('1','A','" & A1P03 & "','" & Text2 & "','110602','TOT'," & A1P07 & ",0,'" & ChgSQL(A1P14) & "','" & A1P15 & "','" & A1P16 & "','" & A1P17 & "'," & A1P18 & ")"
         adoTaie.Execute strSql, intI
         
         .MoveNext 'Added by Morgan 2023/11/17
      Loop 'Added by Morgan 2023/11/17
      End With
   End If
   'end 2021/3/8

   'add by sonia 2020/4/28 提醒法律所案號案件之二事務所之間收據也要同時收款
   'Modified by Morgan 2023/11/17 修正法律案拆收據智慧所收據會重複計算問題 Ex:L006723000(AB2042462)-E11222711,E11223528->AB2042384
   'strExc(0) = "SELECT A0K11,A0J13,TO_CHAR(SUM(A0J09+A0J10),'999,999,999') AMT FROM ( " & _
      "SELECT AXD04,AXD05,AXD06,L1.LOS01,L2.LOS06,A0J13,A0J01,A0J09-NVL(A1U07,0) A0J09,A0J10-NVL(A1U08,0) A0J10,A0K11 FROM ACC441,LAWOFFICESOURCE L1,LAWOFFICESOURCE L2,ACC0J0,ACC0K0," & _
      " (SELECT A1U03,SUM(A1U07) A1U07,SUM(A1U08) A1U08 FROM ACC1U0 WHERE A1U03 IN " & _
      " (SELECT DISTINCT A0J01 FROM ACC441,LAWOFFICESOURCE L1,LAWOFFICESOURCE L2,ACC0J0 WHERE AXD01='" & m_A4401 & "' AND AXD02=" & m_A4402 & " AND AXD03=" & m_A4403 & " " & _
      "         AND AXD05=L1.LOS01(+) AND AXD05=L2.LOS06(+) AND L1.LOS01||L2.LOS01 IS NOT NULL AND DECODE(SUBSTR(L1.LOS02,1,1),'B',L1.LOS06,'C',L1.LOS06,L2.LOS10)=A0J01(+) AND A0J01 IS NOT NULL) GROUP BY A1U03)" & _
      " WHERE AXD01='" & m_A4401 & "' AND AXD02=" & m_A4402 & " AND AXD03=" & m_A4403 & " AND AXD05=L1.LOS01(+) AND AXD05=L2.LOS06(+) AND L1.LOS01||L2.LOS01 IS NOT NULL " & _
      " AND DECODE(SUBSTR(L1.LOS02,1,1),'B',L1.LOS06,'C',L1.LOS06,L2.LOS10)=A0J01(+) AND A0J01 IS NOT NULL AND A0J01=A1U03(+) AND A0J13=A0K01(+)" & _
      " ) GROUP BY A0K11,A0J13 ORDER BY 1,2"
   strExc(0) = "SELECT A0K11,A0K01,TO_CHAR(A0K06-A1U07+A0K07-A1U09,'999,999,999') AMT" & _
      " FROM (SELECT A0K01 X1,NVL(SUM(A1U07),0) A1U07,NVL(SUM(A1U08),0) A1U09" & _
      " FROM ACC0K0,ACC1U0 WHERE A0K01 IN (" & m_TTRcpSQL & ") AND A1U02(+)=A0K01 group by A0K01" & _
      ") X,ACC0K0 WHERE A0K01(+)=X1 ORDER BY 1,2"
   'end 2023/11/17
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      stMsg = "有案源收據，要同時收款："
      Do While Not RsTemp.EOF
         'Modified by Morgan 2023/11/17
         'stMsg = stMsg & vbCrLf & RsTemp("A0K11") & " 公司   " & RsTemp("A0J13") & " 金額 " & RsTemp("AMT")
         stMsg = stMsg & vbCrLf & RsTemp("A0K11") & " 公司   " & RsTemp("A0K01") & " 金額 " & RsTemp("AMT")
         'end 2023/11/17
         RsTemp.MoveNext
      Loop
      MsgBox stMsg
   End If
   'end 2020/4/28
End Sub
'end 2013/12/17

Private Sub Form_Activate()
   m_Activated = True 'Added by Morgan 2014/1/2
   
   strFormName = Name
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   'If adoacc0l0.RecordCount <> 0 Then
   '   adoacc0l0.MoveFirst
   'End If
   'adoacc0l0.Find "a0l01 = '" & strItemNo & "'", 0, adSearchForward, 1
   'If adoacc0l0.EOF = False Then
   '   FormShow
   '   AdodcRefresh
   '   SumShow
   '   RecordShow
   'End If
   
   If Adodc1.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc1.Recordset.Fields("a1p22").Value) = False Then
         If adoquery.State = adStateOpen Then adoquery.Close       'Added by Lydia 2024/11/28
         adoquery.CursorLocation = adUseClient
         'Modified by Morgan 2013/12/19 一張收款單有可能有兩張傳票號
         'adoQuery.Open "select ax210 from acc021 where ax201 = '" & Adodc1.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc1.Recordset.Fields("a1p22").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         adoquery.Open "select ax210 from acc021 where (AX201,AX202) in (select a1p01,a1p22 from acc1p0 where a1p04='" & Adodc1.Recordset.Fields("a1p04").Value & "')", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            If Text2.Enabled Then Text2.SetFocus
            adoquery.Close
            Acc0l0Refresh
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
      If Text3 <> Text12 Then
         AdodcRefresh
         strDelConfirm = MsgBox(MsgText(11), vbOKCancel + vbDefaultButton1, MsgText(5))
         If strDelConfirm = vbCancel Then
            If adoquery.State = adStateOpen Then adoquery.Close       'Added by Lydia 2024/11/28
            adoquery.CursorLocation = adUseClient
            adoquery.Open "select * from acc0m0 where a0m01 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
            Do While adoquery.EOF = False
               adoTaie.Execute "update acc0k0 set a0k17 = a0k17 - " & Val(adoquery.Fields("a0m04").Value) & ", a0k18 = a0k18 - " & Val(adoquery.Fields("a0m05").Value) & " where a0k01 = '" & adoquery.Fields("a0m02").Value & "'"
               adoquery.MoveNext
            Loop
            adoquery.Close
            adoquery.CursorLocation = adUseClient
            'Modify by Amy 2020/06/30 +a0e01/a0e07因改為key,故需抓 a1p10/a1p11
            adoquery.Open "select a1p09,a1p10,a1p11 from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
            Do While adoquery.EOF = False
               If IsNull(adoquery.Fields("a1p09").Value) = False Then
                  'adoTaie.Execute "delete from acc0e0 where a0e02 = '" & adoquery.Fields("a1p09").Value & "'"
                  adoTaie.Execute "delete from acc0e0 where a0e02 = '" & adoquery.Fields("a1p09").Value & "' And a0e01='" & adoquery.Fields("a1p10") & "' And a0e07='" & adoquery.Fields("a1p11") & "' "
            'end 2020/06/30
               End If
               adoquery.MoveNext
            Loop
            adoquery.Close
            adoTaie.Execute "delete from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'"
            adoTaie.Execute "delete from acc0l0 where a0l01 = '" & Text2 & "'"
            'Modify by Morgan 2011/8/18 移到下面改抓acc1u0更新(修正部分收款的已收金額也會被清除問題)
            'adoTaie.Execute "update acc0k0 set a0k17 = 0, a0k18 = 0 where a0k01 in (select a0m02 from acc0m0 where a0m01 = '" & Text2 & "')"
            'adoTaie.Execute "update caseprogress set cp73 = 0, cp74 = 0, cp75 = 0, cp79 = cp16 where cp60 in (select a0m02 from acc0m0 where a0m01 = '" & Text2 & "')"
            strSql = "update acc0k0 set (a0k17,a0k18)=(select nvl(sum(a1u04),0),nvl(sum(a1u05),0)" & _
               " from acc1u0 where a1u02=a0k01 and a1u01<>'" & Text2 & "')" & _
               " where a0k01 in (select a0m02 from acc0m0 where a0m01 = '" & Text2 & "')"
            adoTaie.Execute strSql, intI
            strSql = "update caseprogress set (cp73,cp74,cp75)=(select nvl(sum(a1u04),0),nvl(sum(a1u05),0)" & _
               ",nvl(sum(a1u04),0)+nvl(sum(a1u05),0) from acc1u0 where a1u03=cp09 and a1u01<>'" & Text2 & "')" & _
               " where cp09 in (select a0j01 from acc0m0,acc0j0 where a0m01 = '" & Text2 & "' and a0j13(+)=a0m02)"
            adoTaie.Execute strSql, intI
            adoTaie.Execute "update caseprogress set cp79 = nvl(cp16, 0) - nvl(cp75, 0) - nvl(cp77, 0) + nvl(cp78, 0) where cp09 in (select a0j01 from acc0m0,acc0j0 where a0m01 = '" & Text2 & "' and a0j13(+)=a0m02)"
            'end 2011/8/18
            adoTaie.Execute "delete from acc0m0 where a0m01 = '" & Text2 & "'"
            adoTaie.Execute "delete from acc1u0 where a1u01 = '" & Text2 & "'"
            adoTaie.Execute "delete from acc0t0 where a0t01 = '" & adoacc0l0.Fields("a0l06").Value & "'"
            adoTaie.Execute "Update acc440 set a4416=null where a4416='" & Text2 & "'", intI 'Added by Morgan 2018/10/1
            Frmacc1150_Clear
            Acc0l0Refresh
            SumShow
            strCon1 = ""
            Exit Sub
         End If
      Else
         AdodcRefresh
         'strDelConfirm = MsgBox(MsgText(131), vbOKCancel + vbDefaultButton1, MsgText(5))
         'If strDelConfirm = vbCancel Then
         '   adoquery.CursorLocation = adUseClient
         '   adoquery.Open "select * from acc0m0 where a0m01 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
         '   Do While adoquery.EOF = False
         '      adoTaie.Execute "update acc0k0 set a0k17 = a0k17 - " & Val(adoquery.Fields("a0m04").Value) & ", a0k18 = a0k18 - " & Val(adoquery.Fields("a0m05").Value) & " where a0k01 = '" & adoquery.Fields("a0m02").Value & "'"
         '      adoquery.MoveNext
         '   Loop
         '   adoquery.Close
         '   adoquery.CursorLocation = adUseClient
         '   adoquery.Open "select a1p09 from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
         '   Do While adoquery.EOF = False
         '      If IsNull(adoquery.Fields("a1p09").Value) = False Then
         '         adoTaie.Execute "delete from acc0e0 where a0e02 = '" & adoquery.Fields("a1p09").Value & "'"
         '      End If
         '      adoquery.MoveNext
         '   Loop
         '   adoquery.Close
         '   adoTaie.Execute "delete from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'"
         '   adoTaie.Execute "delete from acc0l0 where a0l01 = '" & Text2 & "'"
         '   adoTaie.Execute "update acc0k0 set a0k17 = 0, a0k18 = 0 where a0k01 in (select a0m02 from acc0m0 where a0m01 = '" & Text2 & "')"
         '   adoTaie.Execute "update caseprogress set cp73 = 0, cp74 = 0, cp75 = 0, cp79 = cp16 where cp60 in (select a0m02 from acc0m0 where a0m01 = '" & Text2 & "')"
         '   adoTaie.Execute "delete from acc0m0 where a0m01 = '" & Text2 & "'"
         '   adoTaie.Execute "delete from acc1u0 where a1u01 = '" & Text2 & "'"
         '   adoTaie.Execute "delete from acc0t0 where a0t01 = '" & adoacc0l0.Fields("a0l06").Value & "'"
         '   Frmacc1150_Clear
         '   Acc0l0Refresh
         '   AdodcRefresh
         '   SumShow
         '   strCon1 = ""
         '   Exit Sub
         'End If
      End If
      
   End If
   Text2 = strItemNo
   Acc0l0Refresh
   
   If adoacc0l0.RecordCount <> 0 Then
      FormShow
      AdodcRefresh
      SumShow
      RecordShow
   End If
   strItemNo = MsgText(601)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call PUB_SaveTrackMode(1, KeyCode)  'Form2.0 記錄鍵盤傳入順序
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 8940 'Modify by Amy 2023/08/18 原:8850
   'Modified by Morgan 2023/8/7 win10邊框較粗
   'Me.Height = 5500
   Me.Height = 5770
   'end 2023/8/7
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   Combo1.AddItem ComboItem(11)
   Combo1.AddItem ComboItem(12)
   Combo1.AddItem ComboItem(13)
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   strItemNo = MsgText(601)
   OpenTable
   If adoacc0l0.RecordCount <> 0 Then
      adoacc0l0.MoveLast
      adoacc0l0.MoveFirst
      RecordShow
   End If
   FormDisabled
   Command1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   'Add by Amy 2018/06/06 不可輸小數判斷,避免智權人員實績與結餘輸入因點數四捨五入後,導致智權人員實績與結餘分析表出現負數
   If ChkDot = True Then
      tool1_enabled
      MsgBox "41字頭或7121科目不可輸入小數！", , MsgText(5)
      Cancel = True
      Exit Sub
   End If
   'end 2018/06/06
   If CreDebCheck <> MsgText(602) Then
      tool1_enabled
      MsgBox MsgText(11), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc1150 = Nothing
   
   'Added by Morgan 2013/12/26
   If m_A4401 <> "" Then
      If m_AutoRun = True Then
         Frmacc1155.m_iReturn = -1
      End If
      Frmacc1155.Show
   End If
   'end 2013/12/26
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox Label1 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label1 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If strSaveConfirm <> MsgText(3) Then
      Exit Sub
   End If
   If Mid(MaskEdBox1.Text, 1, 3) <> Mid(CFDate(ACDate(ServerDate)), 1, 3) Then
      Text2 = UpdateNo("acc0l0", "a0l01", 5, MaskEdBox1.Text, MsgText(803))
   Else
      'Text2 = AutoNo(MsgText(803), 5)
      Text2 = strDocNo
   End If
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
   If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
      Exit Sub
   End If
   If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
      MsgBox Label9 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox2.SetFocus
      Exit Sub
   End If
   RemarkShow
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   'edit by nickc 2007/06/11  切換輸入法改用API
   OpenIme
End Sub

Private Sub Text1_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
CloseIme
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc0l0.CursorLocation = adUseClient
   adoacc0l0.MaxRecords = intMax
   adoacc0l0.Open "select * from acc0l0 where a0l01 >= '" & Text2 & "' order by a0l01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc1p0.CursorLocation = adUseClient
   'Modified by Morgan 2013/12/19 收款會有J公司,取消 a1p01='1' 條件
   adoacc1p0.Open "select * from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "' order by a1p03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   'Modified by Morgan 2013/12/19 收款會有J公司,取消 a1p01='1' 條件
   adoadodc1.Open "select * from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "' order by a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表(國內收款資料(主檔))
'
'*************************************************
Public Sub FormShow()
   Text2 = adoacc0l0.Fields("a0l01").Value
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc0l0.Fields("a0l02").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(adoacc0l0.Fields("a0l02").Value)
   End If
   MaskEdBox1.Mask = DFormat
   If IsNull(adoacc0l0.Fields("a0l07").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = adoacc0l0.Fields("a0l07").Value
   End If
   'Added by Morgan 2013/12/27
   Text21 = "" & adoacc0l0.Fields("a0l05").Value
   If Text21 = "" Then Text21 = "1"
   'end 2013/12/27
End Sub

'*************************************************
'  顯示資料表(國內收款資料(分錄檔))
'
'*************************************************
Private Sub AdodcShow()
   Text4 = Adodc1.Recordset.Fields("a1p05").Value
   Text4.Tag = Text4.Text 'Added by Morgan 2015/1/30
   If IsNull(Adodc1.Recordset.Fields("a1p07").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = Adodc1.Recordset.Fields("a1p07").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p08").Value) Then
      Text11 = MsgText(601)
   Else
      Text11 = Adodc1.Recordset.Fields("a1p08").Value
   End If
   
   If IsNull(Adodc1.Recordset.Fields("a1p09").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = Adodc1.Recordset.Fields("a1p09").Value
   End If
   Text7.Tag = Text7 'Add by Amy 2020/06/30
   If IsNull(Adodc1.Recordset.Fields("a1p10").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = Adodc1.Recordset.Fields("a1p10").Value
   End If
   Text9.Tag = Text9 'Add by Amy 2020/06/30
   If IsNull(Adodc1.Recordset.Fields("a1p11").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = Adodc1.Recordset.Fields("a1p11").Value
   End If
   Text8.Tag = Text8 'Add by Amy 2020/06/30
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(Adodc1.Recordset.Fields("a1p12").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(Adodc1.Recordset.Fields("a1p12").Value)
   End If
   MaskEdBox2.Mask = DFormat
   If IsNull(Adodc1.Recordset.Fields("a1p23").Value) Then
      Text10 = MsgText(601)
   Else
      Text10 = Adodc1.Recordset.Fields("a1p23").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p13").Value) Then
      Combo1 = MsgText(601)
   Else
      Combo1 = Combo1.List(Val(Adodc1.Recordset.Fields("a1p13").Value) - 1)
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p16").Value) Then
      Text13 = MsgText(601)
      Text17 = ""
   Else
      Text13 = Adodc1.Recordset.Fields("a1p16").Value
      Text17 = StaffQuery(Text13)
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p14").Value) Then
      Combo2 = MsgText(601)
   Else
      Combo2 = Adodc1.Recordset.Fields("a1p14").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p06").Value) Then
      Text14 = MsgText(601)
   Else
      Text14 = Adodc1.Recordset.Fields("a1p06").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p17").Value) Then
      Text15 = MsgText(601)
   Else
      Text15 = Adodc1.Recordset.Fields("a1p17").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p15").Value) Then
      Text16 = MsgText(601)
   Else
      Text16 = Adodc1.Recordset.Fields("a1p15").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p30").Value) Then
      Text18 = MsgText(601)
   Else
      Text18 = Adodc1.Recordset.Fields("a1p30").Value
   End If
   
   Text19 = "" & Adodc1.Recordset.Fields("a1p01").Value 'Added by Morgan 2012/12/19
   'Added by Morgan 2020/12/25
   lblA1P22 = "" & Adodc1.Recordset.Fields("a1p22").Value
   If lblA1P22 <> "" Then Text19.Enabled = False
   'end 2020/12/25
   
   'add by sonia 2025/5/9 2401且有暫收單號時鎖住業務編號及客戶編號
   If Text4 = "2401" And Text10 <> "" Then
      Text13.Enabled = False
      Text16.Enabled = False
   Else
      Text13.Enabled = True
      Text16.Enabled = True
   End If
   'end 2025/5/9
End Sub

'*************************************************
'  計算並顯示總計
'
'*************************************************
Public Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   'Modified by Morgan 2013/12/19 收款會有J公司,取消 a1p01='1' 條件
   adoaccsum.Open "select sum(a1p07), sum(a1p08), count(*) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text3 = MsgText(601)
         douTotalAmount = 0
      Else
         Text3 = Format(adoaccsum.Fields(0).Value, FDollar)
         douTotalAmount = Val(adoaccsum.Fields(0).Value)
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text12 = MsgText(601)
      Else
         Text12 = Format(adoaccsum.Fields(1).Value, FDollar)
      End If
      If IsNull(adoaccsum.Fields(2).Value) Then
         Text20 = MsgText(601)
      Else
         Text20 = Format(adoaccsum.Fields(2).Value, DDollar)
      End If
   Else
      Text3 = MsgText(601)
      douTotalAmount = 0
      Text12 = MsgText(601)
      Text20 = MsgText(601)
   End If
   adoaccsum.Close
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Exit Sub
   Else
      If Text3 = MsgText(601) Or Val(Text3) = 0 Then
         Command1.Enabled = False
         Command5.Enabled = False 'Added by Morgan 2013/12/27
      Else
         Command1.Enabled = True
         Command5.Enabled = True 'Added by Morgan 2013/12/27
      End If
   End If
End Sub

'*************************************************
'  儲存資料表(國內收款資料(分錄檔))
'
'*************************************************
Private Function Acc1p0Save() As Boolean
   Dim bCancel As Boolean
   Dim strSql As String 'Add by Amy 2020/06/30
   
On Error GoTo Checking
   If Text4 = MsgText(601) Then
      MsgBox MsgText(10) & Label5, , MsgText(5)
      strControlButton = MsgText(602)
      Text4.SetFocus
      Exit Function
   Else
   
      'Added by Morgan 2025/3/19
      '票據檢查
      'Modified by Morgan 2025/4/8 加判斷有傳票才檢查，否則隔月(換收據)會無法修改
      If Adodc1.Recordset.RecordCount > 0 Then
         If IsNull(Adodc1.Recordset.Fields("a1p22").Value) = False Then
            If Text7 <> "" Then
               strExc(0) = "select * from acc0e0 where a0e01 = '" & Text9 & "' and a0e02 = '" & Text7 & "' and a0e07='" & Text8 & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If RsTemp.Fields("a0e03").Value <> Text2 Then
                     MsgBox MsgText(196), , MsgText(5)
                     strControlButton = MsgText(602)
                     Text7.SetFocus
                     Exit Function
                  ElseIf RsTemp("a0e19") & RsTemp("a0e20") <> "" Then
                     MsgBox "票據[" & Text7 & "/" & Text8 & "/" & Text9 & "]已處理(託收,兌現...)，不可修改！", vbExclamation, "輸入錯誤"
                     strControlButton = MsgText(602)
                     Exit Function
                  End If
               End If
            End If
         End If
      End If
      'end 2025/4/8
      
      If Text7.Tag <> "" And Text7 & Text8 & Text9 <> Text7.Tag & Text8.Tag & Text9.Tag Then
         strExc(0) = "select * from acc0e0 where a0e01 = '" & Text9.Tag & "' and a0e02 = '" & Text7.Tag & "' and a0e07='" & Text8.Tag & "' and a0e19||a0e20 is not null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            MsgBox "原票據[" & Text7.Tag & "/" & Text8.Tag & "/" & Text9.Tag & "]已處理(託收,兌現...)，不可刪除！", vbExclamation, "輸入錯誤"
            strControlButton = MsgText(602)
            Exit Function
         End If
      End If
      'end 2025/3/19
   
      'Added by Morgan 2013/12/26
      '票據
      If Text4 = "113001" Then
         If Text7 = "" Then
            MsgBox "請輸入票據號碼！", vbExclamation, MsgText(5)
            strControlButton = MsgText(602)
            Text7.SetFocus
            Exit Function
         End If
         If Text9 = "" Then
            MsgBox "請輸入收票銀行！", vbExclamation, MsgText(5)
            strControlButton = MsgText(602)
            Text9.SetFocus
            Exit Function
         End If
         If Text8 = "" Then
            MsgBox "請輸入收票帳號！", vbExclamation, MsgText(5)
            strControlButton = MsgText(602)
            Text8.SetFocus
            Exit Function
         End If
         
      'Removed by Morgan 2014/1/2 有沒有單號的情形,改回不必檢查
      '暫收款
'      ElseIf Text4 = "2401" Then
'         If Text10 = "" Then
'            MsgBox "請輸入暫收款單號！", vbExclamation, MsgText(5)
'            strControlButton = MsgText(602)
'            Text10.SetFocus
'            Exit Function
'         End If

      End If
      'end 2013/12/26
      
      'Added by Morgan 2015/1/30
      If Left(Text4.Tag, 1) = "4" And Text4.Text = "2401" Then
         MsgBox "收入科目暫不做點數時, 請改用 2492點數保留 科目！", vbExclamation
         strControlButton = MsgText(602)
         Text4.SetFocus
         Exit Function
      End If
      'end 2015/1/30
      
      'add by sonia 2020/11/9
      If Left(Text4, 1) = "6" And Text19 = "L" And (Text14 = "TOT" Or Trim(Text14) = MsgText(601)) Then
         MsgBox "L公司費用科目的部門不可輸TOT或空白！", vbExclamation
         strControlButton = MsgText(602)
         Text14.SetFocus
         Exit Function
      End If
      'end 2020/11/9
         
      'Add by Morgan 2008/7/8 本所案號檢查
      Text15_Validate bCancel
      If bCancel = True Then
         strControlButton = MsgText(602)
         Text15.SetFocus
         Exit Function
      End If
      Text16_Validate bCancel
      If bCancel = True Then
         strControlButton = MsgText(602)
         Text16.SetFocus
         Exit Function
      End If
      'end 2008/7/8
      If ExistCheck("acc010", "a0101", Text4, Label5) = False Then
         strControlButton = MsgText(602)
         Text4.SetFocus
         Exit Function
      End If
      If Mid(Text4, 1, 2) = "41" Then
         If ExistCheck("staff", "st01", Text13, Label15) = False Then
            MsgBox MsgText(10) & Label15, , MsgText(5)
            strControlButton = MsgText(602)
            Text13.SetFocus
            Exit Function
         End If
      End If
      If Val(Text6) = 0 And Val(Text11) = 0 Then
         MsgBox MsgText(58), , MsgText(5)
         strControlButton = MsgText(602)
         Text6.SetFocus
         Exit Function
      End If
      If CheckDept(Text4, Text14) = False Then
         MsgBox MsgText(103), , MsgText(5)
         strControlButton = MsgText(602)
         Text14.SetFocus
         Exit Function
      End If
      'add by sonia 2015/12/30 檢查民國105年起法務收入科目不可使用
      intI = PUB_AccNoEnable(Text4, Val(FCDate(MaskEdBox1.Text)))
      If intI <> 0 Then
         strControlButton = MsgText(602)
         Text4.SetFocus
         Exit Function
      End If
      'end 2015/12/30
      'Add by Morgan 2007/2/2 檢查科目部門&智權人員是否正確
      intI = PUB_AccNoGood(Text4, Text14, Text13)
      If intI <> 0 Then
         strControlButton = MsgText(602)
         If intI = 1 Then
            Text4.SetFocus
         ElseIf intI = 2 Then
            Text14.SetFocus
         ElseIf intI = 3 Then
            Text13.SetFocus
         End If
         Exit Function
      End If
      'end 2007/2/2
      If strSaveConfirm = MsgText(3) And m_SaveCheck = False Then
         If adoquery.State = adStateOpen Then adoquery.Close       'Added by Lydia 2024/11/28
         adoquery.CursorLocation = adUseClient
         '2012/2/13 MODIFY BY SONIA 加 A0E04(票號0006148)
         'adoquery.Open "select a0e02 from acc0e0 where a0e02 = '" & Text7 & "' and a0e01 = '" & Text9 & "'", adoTaie, adOpenStatic, adLockReadOnly
         'Modify by Amy 2020/06/30 +a0e07因改為key
         adoquery.Open "select a0e02 from acc0e0 where a0e02 = '" & Text7 & "' and a0e01 = '" & Text9 & "' and a0e07='" & Text8 & "' and a0e04 = 'R' ", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(148), , MsgText(5)
            adoquery.Close
            strControlButton = MsgText(602)
            Text7.SetFocus
            Exit Function
         End If
         adoquery.Close
      End If
      If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
         If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
            MsgBox Label9 & MsgText(63), , MsgText(5)
            strControlButton = MsgText(602)
            MaskEdBox2.SetFocus
            Exit Function
         End If
      Else
         If Text4 = "113001" Then
            MsgBox Label9 & MsgText(63), , MsgText(5)
            strControlButton = MsgText(602)
            MaskEdBox2.SetFocus
            Exit Function
         End If
      End If
      If Text9 <> MsgText(601) Then
         If ExistCheck("acc0g0", "a0g01", Text9, Label10) = False Then
            strControlButton = MsgText(602)
            Text9.SetFocus
            Exit Function
         End If
      End If
   End If
   
   'Added by Morgan 2013/12/26
   '公司別檢查
   If Text19 = MsgText(601) Then
      MsgBox MsgText(10) & Label19, , MsgText(5)
      strControlButton = MsgText(602)
      Text19.SetFocus
      Exit Function
      
'Removed by Morgan 2014/1/27 改最後存檔前檢查分錄的公司有主要公司別就好
'   'Added by Morgan 2014/1/20
'   ElseIf strSerialNo <> "" Then
'      If Adodc1.Recordset.Fields("a1p01") <> Text19 Then
'         If Not IsNull(Adodc1.Recordset.Fields("a1p22")) Then
'            MsgBox "已有傳票號，公司別不可變更！", vbExclamation
'            strControlButton = MsgText(602)
'            Text19.SetFocus
'            Exit Function
'         End If
'      End If
'end 2014/1/27

   End If
   
   If PUB_CheckCompany(Text4, Text19) = False Then
      strControlButton = MsgText(602)
      Text19.SetFocus
      Exit Function
   End If
   'end 2013/12/26
   
   'Added by Morgan 2014/2/12
   'Modified by Morgan 2024/1/3 +2492保留點數
   If Text4 = "2405" Or Text4 = "2492" Then
      'Modified by Morgan 2024/1/22 有可能手動改作保留(如票期太長時F11300464)--瑞婷
      'If Text10 = "" Then
      If Text10 = "" And Text4 = "2405" Then
      'end 2024/1/22
         MsgBox "請於[暫收款單號]欄位輸入[收文號+收據號碼]!!", vbInformation, "輸入提醒"
         strControlButton = MsgText(602)
         Text10.SetFocus
         Exit Function
      End If
      
      If Text10 <> "" Then
         strExc(0) = "select 1 from acc0j0 where a0j01||a0j13='" & Text10 & "' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            MsgBox "請於[暫收款單號]欄位輸入[收文號+收據號碼]!!", vbExclamation, "輸入錯誤"
            strControlButton = MsgText(602)
            Text10.SetFocus
            Exit Function
         End If
      End If
   'end 2014/2/12
   ElseIf Text10 <> MsgText(601) Then
      If Text4 <> "2401" Then
         MsgBox MsgText(113), , MsgText(5)
         strControlButton = MsgText(602)
         Text10.SetFocus
         Exit Function
      End If
      
      'Added by Morgan 2014/1/20
      '暫收款的公司別必須與分錄的相同
      strExc(0) = "select a0t18 from acc0t0 where a0t01='" & Text10 & "' and a0t18 is not null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp(0) <> Text19 Then
            MsgBox "暫收單號的公司別與收款公司別必須相同!!", vbExclamation
            strControlButton = MsgText(602)
            Text10.SetFocus
            Exit Function
         End If
      End If
      'end 2014/1/20
      
      If Val(Text11) <> 0 Then
         If CheckAmount(Val(Text11)) = False Then
            MsgBox MsgText(112), , MsgText(5)
            strControlButton = MsgText(602)
            Text10.SetFocus
            Exit Function
         End If
      End If
      If Val(Text6) <> 0 Then
         If CheckAmount(Val(Text6)) = False Then
            MsgBox MsgText(112), , MsgText(5)
            strControlButton = MsgText(602)
            Text10.SetFocus
            Exit Function
         End If
      End If
      If Text15 <> MsgText(601) Then
         Text15 = CaseNoZero(Text15)
         If adoquery.State = adStateOpen Then adoquery.Close       'Added by Lydia 2024/11/28
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select pa01 as SystemNo from patent where pa01 = '" & Mid(Text15, 1, Len(Text15) - 9) & "' and pa02 = '" & Mid(Text15, Len(Text15) - 8, 6) & "' and pa03 = '" & Mid(Text15, Len(Text15) - 2, 1) & "' and pa04 = '" & Mid(Text15, Len(Text15) - 1, 2) & "' union " & _
                        "select tm01 as SystemNo from trademark where tm01 = '" & Mid(Text15, 1, Len(Text15) - 9) & "' and tm02 = '" & Mid(Text15, Len(Text15) - 8, 6) & "' and tm03 = '" & Mid(Text15, Len(Text15) - 2, 1) & "' and tm04 = '" & Mid(Text15, Len(Text15) - 1, 2) & "' union " & _
                        "select lc01 as SystemNo from lawcase where lc01 = '" & Mid(Text15, 1, Len(Text15) - 9) & "' and lc02 = '" & Mid(Text15, Len(Text15) - 8, 6) & "' and lc03 = '" & Mid(Text15, Len(Text15) - 2, 1) & "' and lc04 = '" & Mid(Text15, Len(Text15) - 1, 2) & "' union " & _
                        "select hc01 as SystemNo from hirecase where hc01 = '" & Mid(Text15, 1, Len(Text15) - 9) & "' and hc02 = '" & Mid(Text15, Len(Text15) - 8, 6) & "' and hc03 = '" & Mid(Text15, Len(Text15) - 2, 1) & "' and hc04 = '" & Mid(Text15, Len(Text15) - 1, 2) & "' union " & _
                        "select sp01 as SystemNo from servicepractice where sp01 = '" & Mid(Text15, 1, Len(Text15) - 9) & "' and sp02 = '" & Mid(Text15, Len(Text15) - 8, 6) & "' and sp03 = '" & Mid(Text15, Len(Text15) - 2, 1) & "' and sp04 = '" & Mid(Text15, Len(Text15) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount = 0 Then
            MsgBox MsgText(28) & Label17, , MsgText(5)
            strControlButton = MsgText(602)
            adoquery.Close
            Exit Function
         End If
         adoquery.Close
      End If
      If Text16 <> MsgText(601) Then
         If Len(Text16) = 6 Then
            Text16 = AfterZero(Text16)
         'Add by Morgan 2007/3/1 八碼時要補'0'
         ElseIf Len(Text16) = 8 Then
            Text16 = Text16 & "0"
         'End 2007/3/1
         End If
         If ExistCheck("customer", "cu01", Mid(Text16, 1, 8), Label11, False) = False Then
            If ExistCheck("acc0i0", "a0i01", Text16, Label11, False) = False Then
               If ExistCheck("staff", "st01", Text16, Label11, False) = False Then
                  MsgBox MsgText(28) & Label18, , MsgText(5)
                  strControlButton = MsgText(602)
                  Exit Function
               End If
            End If
         End If
      End If
      
      If Text18 = "" Then Text18 = Text10 'Added by Morgan 2015/8/21 暫收款單號寫到其他對沖
   End If
   
   If Adodc1.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc1.Recordset.Fields("a1p22").Value) = False Then
         If adoquery.State = adStateOpen Then adoquery.Close       'Added by Lydia 2024/11/28
         adoquery.CursorLocation = adUseClient
         'Modified by Morgan 2013/12/19 一張收款單有可能有兩張傳票號
         'adoQuery.Open "select ax210 from acc021 where ax201 = '" & Adodc1.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc1.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         adoquery.Open "select ax210 from acc021 where (ax201,ax202) in(select a1p01,a1p22 from acc1p0 where a1p04='" & Adodc1.Recordset.Fields("a1p04").Value & "') and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(155), , MsgText(5)
            strControlButton = MsgText(602)
            Text4.SetFocus
            adoquery.Close
            Exit Function
         End If
         adoquery.Close
      End If
   End If
   If Adodc1.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc1.Recordset.Fields("a1p09").Value) = False Then
         If adoquery.State = adStateOpen Then adoquery.Close       'Added by Lydia 2024/11/28
         adoquery.CursorLocation = adUseClient
         'Modify by Amy 2020/06/30 +a0e07因改為key
         strSql = "select a0e02 from acc0e0 where a0e01 = '" & Text9 & "' and a0e02 = '" & Adodc1.Recordset.Fields("a1p09").Value & "' And a0e07='" & Text8 & "' " & _
                    "and a0e03 <> '" & Text2 & "' and ((a0e14 is not null and a0e14 <> 0) or (a0e15 is not null and a0e15 <> 0) or (a0e17 is not null and a0e17 <> 0) or (a0e21 is not null and a0e21 <> 0) or (a0e34 is not null and a0e34 <> 0))"
         adoquery.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(193), , MsgText(5)
            strControlButton = MsgText(602)
            Text7.SetFocus
            adoquery.Close
            Exit Function
         End If
         adoquery.Close
      End If
   End If
   adoacc1p0.Close
   adoacc1p0.CursorLocation = adUseClient
   'Modified by Morgan 2013/12/19 收款會有J公司,取消 a1p01='1' 條件
   adoacc1p0.Open "select * from acc1p0 where a1p02 = 'A' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text2 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc1p0.RecordCount = 0 Then
      adoacc1p0.AddNew
      'Modified by Morgan 2013/12/19
      'adoacc1p0.Fields("a1p01").Value = "1"
      adoacc1p0.Fields("a1p01").Value = Text19
      adoacc1p0.Fields("a1p02").Value = "A"
      'Modified by Morgan 2013/12/19 收款會有J公司,取消 a1p01='1' 條件
      adoacc1p0.Fields("a1p03").Value = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'A' and a1p04 = '" & Text2 & "'", 3)
      strSerialNo = adoacc1p0.Fields("a1p03").Value
      adoacc1p0.Fields("a1p04").Value = Text2
   End If
   adoacc1p0.Fields("a1p01").Value = Text19 'Added by Morgan 2014/1/3
   adoacc1p0.Fields("a1p05").Value = Text4
   adoacc1p0.Fields("a1p06").Value = MsgText(55)
   If Text6 <> MsgText(601) Then
      adoacc1p0.Fields("a1p07").Value = Val(Text6)
   Else
      adoacc1p0.Fields("a1p07").Value = 0
   End If
   If Text11 <> MsgText(601) Then
      adoacc1p0.Fields("a1p08").Value = Val(Text11)
   Else
      adoacc1p0.Fields("a1p08").Value = 0
   End If
   If Text7 <> MsgText(601) Then
      adoacc1p0.Fields("a1p09").Value = Text7
   Else
      adoacc1p0.Fields("a1p09").Value = Null
   End If
   If Text9 <> MsgText(601) Then
      adoacc1p0.Fields("a1p10").Value = Text9
   Else
      adoacc1p0.Fields("a1p10").Value = Null
   End If
   If Text8 <> MsgText(601) Then
      adoacc1p0.Fields("a1p11").Value = Text8
   Else
      adoacc1p0.Fields("a1p11").Value = Null
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      adoacc1p0.Fields("a1p12").Value = Val(FCDate(MaskEdBox2.Text))
   Else
      adoacc1p0.Fields("a1p12").Value = Null
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      adoacc1p0.Fields("a1p18").Value = Val(FCDate(MaskEdBox1.Text))
   Else
      adoacc1p0.Fields("a1p18").Value = Null
   End If
   If Text10 <> MsgText(601) Then
      adoacc1p0.Fields("a1p23").Value = Text10
   Else
      adoacc1p0.Fields("a1p23").Value = Null
   End If
   If Combo1 <> MsgText(601) Then
      adoacc1p0.Fields("a1p13").Value = Mid(Combo1, 1, 1)
   Else
      adoacc1p0.Fields("a1p13").Value = Null
   End If
   If Text13 <> MsgText(601) Then
      adoacc1p0.Fields("a1p16").Value = Text13
   Else
      adoacc1p0.Fields("a1p16").Value = Null
   End If
   If Combo2 <> MsgText(601) Then
      adoacc1p0.Fields("a1p14").Value = Combo2
      Combo2.AddItem Combo2
   Else
      adoacc1p0.Fields("a1p14").Value = Null
   End If
   If Text14 <> MsgText(601) Then
      adoacc1p0.Fields("a1p06").Value = Text14
   Else
      adoacc1p0.Fields("a1p06").Value = MsgText(55)
   End If
   If Text15 <> MsgText(601) Then
      adoacc1p0.Fields("a1p17").Value = Text15
   Else
      adoacc1p0.Fields("a1p17").Value = Null
   End If
   If Text16 <> MsgText(601) Then
      adoacc1p0.Fields("a1p15").Value = Text16
   Else
      adoacc1p0.Fields("a1p15").Value = Null
   End If
   If Text18 <> MsgText(601) Then
      adoacc1p0.Fields("a1p30").Value = Text18
   Else
      adoacc1p0.Fields("a1p30").Value = Null
   End If
   adoacc1p0.UpdateBatch
   If Text7 <> MsgText(601) Then
      If adoquery.State = adStateOpen Then adoquery.Close       'Added by Lydia 2024/11/28
      adoquery.CursorLocation = adUseClient
      'Modify by Amy 2020/06/30 +a0e07因改為key, 更新時拿掉a0e07
      adoquery.Open "select * from acc0e0 where a0e01 = '" & Text9 & "' and a0e02 = '" & Text7 & "' and a0e07='" & Text8 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         If adoquery.Fields("a0e03").Value = Text2 Then
            'Modified by Morgan 2014/1/22 +a0e23
            'adoTaie.Execute "update acc0e0 set a0e05 = '1', a0e06 = '" & Text16 & "', a0e07 = '" & Text8 & "', a0e13 = " & Val(FCDate(MaskEdBox1.Text)) & ", a0e10 = " & Val(FCDate(MaskEdBox2.Text)) & ", a0e08 = '" & Mid(Combo1, 1, 1) & "', a0e12 = '" & Combo2 & "', a0e11 = " & Val(Text6) & ",a0e23='" & Text19 & "' where a0e01 = '" & Text9 & "' and a0e02 = '" & Text7 & "'"
            'Modified by Morgan 2022/12/8 Text16空白時不要回寫
            'strSql = "update acc0e0 set a0e05 = '1', a0e06 = '" & Text16 & "', a0e13 = " & Val(FCDate(MaskEdBox1.Text)) & ", a0e10 = " & Val(FCDate(MaskEdBox2.Text)) & ", a0e08 = '" & Mid(Combo1, 1, 1) & "', a0e12 = '" & Combo2 & "', a0e11 = " & Val(Text6) & ",a0e23='" & Text19 & "' " & _
                         "where a0e01 = '" & Text9 & "' and a0e02 = '" & Text7 & "' And a0e07 = '" & Text8 & "' "
            strSql = "update acc0e0 set a0e05 = '1'" & IIf(Text16 <> "", ", a0e06 = '" & Text16 & "'", "") & ", a0e13 = " & Val(FCDate(MaskEdBox1.Text)) & ", a0e10 = " & Val(FCDate(MaskEdBox2.Text)) & ", a0e08 = '" & Mid(Combo1, 1, 1) & "', a0e12 = '" & Combo2 & "', a0e11 = " & Val(Text6) & ",a0e23='" & Text19 & "' " & _
                         "where a0e01 = '" & Text9 & "' and a0e02 = '" & Text7 & "' And a0e07 = '" & Text8 & "' "
            adoTaie.Execute strSql
         'end 2020/06/30
         Else
            MsgBox MsgText(196), , MsgText(5)
            adoquery.Close
            Text7.SetFocus
            Exit Function
         End If
      Else
         'Add by Amy 2020/06/30 若修改票據號碼按 Insert,原資料不會被刪除
         If Text7 & Text8 & Text9 <> Text7.Tag & Text8.Tag & Text9.Tag Then
             adoTaie.Execute "delete from acc0e0 where a0e01 = '" & Text9.Tag & "' and a0e02 = '" & Text7.Tag & "' and a0e07='" & Text8.Tag & "'"
         End If
         'Modify by Amy 2020/06/30 +a0e07因改為key
         adoTaie.Execute "delete from acc0e0 where a0e01 = '" & Text9 & "' and a0e02 = '" & Text7 & "' and a0e07='" & Text8 & "'"
         '93.10.14 MODIFY BY SONIA 原A0E06未存
         'adoTaie.Execute "insert into acc0e0 values ('" & Text7 & "', '" & Text9 & "', '" & Text2 & "', 'R', '4', null, '" & Text8 & "', '" & Mid(Combo1, 1, 1) & "', " & _
         '                "" & Val(FCDate(MaskEdBox2.Text)) & ", " & Val(Text6) & ", '" & Combo2 & "', " & Val(FCDate(MaskEdBox1.Text)) & ", 0, 0, 0, null, 0, null, null, null, 0, 0, 0, 0, null, null, null, null, null, 0, null, 0, null, null, '" & strUserNum & "', " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", null, 0, 0, 0, 0, 0, null, 0, null, null)"
         'Modified by Morgan 2014/1/22 改指定欄位以免新增欄位後語法會錯,a0e23也要寫入
         'adoTaie.Execute "insert into acc0e0 values ('" & Text7 & "', '" & Text9 & "', '" & Text2 & "', 'R', '4', '" & Text16 & "', '" & Text8 & "', '" & Mid(Combo1, 1, 1) & "', " & _
                         "" & Val(FCDate(MaskEdBox2.Text)) & ", " & Val(Text6) & ", '" & Combo2 & "', " & Val(FCDate(MaskEdBox1.Text)) & ", 0, 0, 0, null, 0, null, null, null, 0, 0, 0, 0, null, null, null, null, null, 0, null, 0, null, null, '" & strUserNum & "', " & Val(strSrvDate(2)) & ", " & ServerTime & ", null, 0, 0, 0, 0, 0, null, 0, null, null)"
         adoTaie.Execute "insert into acc0e0(A0E02,A0E01,A0E03,A0E04,A0E05,A0E06,A0E07,A0E08,A0E10,A0E11,A0E12,A0E13,A0E14,A0E15,A0E16,A0E33,A0E17,A0E18,A0E19,A0E20,A0E21,A0E37,A0E34,A0E22,A0E23,A0E24,A0E32,A0E40,A0E41,A0E25,A0E35,A0E36,A0E38,A0E39,A0E28,A0E26,A0E27,A0E31,A0E29,A0E30,A0E42,A0E43,A0E44,A0E45,A0E46,A0E47,A0E48 )" & _
            " values ('" & Text7 & "', '" & Text9 & "', '" & Text2 & "', 'R', '4', '" & Text16 & "', '" & Text8 & "', '" & Mid(Combo1, 1, 1) & "', " & Val(FCDate(MaskEdBox2.Text)) & ", " & Val(Text6) & ", '" & Combo2 & "', " & Val(FCDate(MaskEdBox1.Text)) & ", 0, 0, 0, null, 0, null, null, null, 0, 0, 0, 0, '" & Text19 & "', null, null, null, null, 0, null, 0, null, null, '" & strUserNum & "', " & Val(strSrvDate(2)) & ", " & ServerTime & ", null, 0, 0, 0, 0, 0, null, 0, null, null)"
         'end 2014/1/22
         '93.10.14 END
      End If
      adoquery.Close
   End If
   AdodcRefresh
   Adodc1.Recordset.Find "a1p03 = '" & strSerialNo & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF Then
      Adodc1.Recordset.MoveFirst
   Else
      DataGrid1.SelBookmarks.add Adodc1.Recordset.Bookmark
   End If
   strSerialNo = MsgText(601)
   
   Acc1p0Save = True 'Added by Morgan 2013/12/26
   
Checking:
   If Err.Number = 0 Then
      Exit Function
   End If
   MsgBox Err.Description, , MsgText(5)
End Function

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   'Modified by Morgan 2013/12/19 收款會有J公司,取消 a1p01='1' 條件
   'modify by sonia 2020/5/12 排序+a1p01
   adoadodc1.Open "select * from acc1p0, acc010, acc0g0 where a1p05 = a0101 and a1p10 = a0g01 (+) and a1p02 = 'A' and a1p04 = '" & Text2 & "' order by a1p01,a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.Requery
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
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
   
   'Added by Morgan 2016/2/18
   '貸方暫收款要檢查若有銷退記錄時不可刪除
   If Text4 = "2401" And Val(Text11) > 0 And Text10 <> "" Then
      strExc(0) = "select * from acc0s0 where a0s02='" & Text10 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         MsgBox "暫收款已有銷退記錄不可刪除！", vbExclamation
         Exit Sub
      End If
   End If
   'end 2016/2/18
   
   'Added by Morgan 2023/8/25
   If Val(strSerialNo) = "1" Then
      MsgBox "第1筆分錄不可刪除！", vbExclamation
      Exit Sub
   End If
   'end 2023/8/25

   'Modify by Amy 2020/06/30 +Text8 <> "" /a0e07因改為key
   If Text9 <> "" And Text7 <> "" And Text8 <> "" Then
      adoTaie.Execute "delete from acc0e0 where a0e01 = '" & Text9 & "' and a0e02 = '" & Text7 & "' And a0e07='" & Text8 & "' "
   End If
   'Modified by Morgan 2013/12/19 收款會有J公司,取消 a1p01='1' 條件
   adoTaie.Execute "delete from acc1p0 where a1p02 = 'A' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text2 & "'"
   'If Text10 <> MsgText(601) Then
   '   adoTaie.Execute "delete from acc0t0 where a0t01 = '" & Text10 & "'"
   'End If
   AdodcRefresh
   AdodcClear
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
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   'Added by Sindy 2021/12/20 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
   If PUB_ChkTrackMode = False Then
       Exit Sub
   End If
   '2021/12/20 END
   
   Select Case KeyCode
      Case vbKeyInsert
         If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
            Exit Sub
         End If
        'Add By Cheng 2004/05/06
        '在修改狀態若傳票已過帳, 則不可更新
        If strSaveConfirm = MsgText(4) And strCon10 <> "" Then
            'Modified by Morgan 2013/12/19 一張收款單有可能有兩張傳票號
            'StrSQLa = "Select Count(*) From ACC021 Where AX202='" & strCon10 & "' And AX210 Is Not Null "
            StrSQLa = "Select Count(*) From ACC021 Where (AX201,AX202) in (select a1p01,a1p22 from acc1p0 where a1p04='" & Text2 & "') And AX210 Is Not Null "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
            If Val("" & rsA.Fields(0).Value) > 0 Then
                MsgBox "傳票" & strCon10 & "已過帳, 不可更改資料!!!", vbExclamation + vbOKOnly
                If rsA.State <> adStateClosed Then rsA.Close
                Set rsA = Nothing
                Exit Sub
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
        End If
        'End
        
         'add by sonia 2019/9/5
         If InStr(Combo2, "轉撥") > 0 Then
            MsgBox "摘要不可輸入 轉撥 二字, 否則會影響實績點數 !!!", vbExclamation + vbOKOnly
            Exit Sub
         End If
         'end 2019/9/5
         
         'add by sonia 2019/9/5
         If Left(Text4, 1) = "4" And Text13 = "" Then
            MsgBox "收入科目，智權人員不可空白 !!!", vbExclamation + vbOKOnly
            Exit Sub
         End If
         'end 2019/9/5
         
         'Add by Sindy 2021/12/14 檢查畫面上的物件是否含有Unicode文字
         If PUB_ChkUniText(Me, True, True) = False Then
            Exit Sub
         End If
         
         Frmacc1150_Save
         If strControlButton <> MsgText(602) Then
            Acc1p0Save
         End If
         If strControlButton <> MsgText(602) Then
            AdodcClear
            SumShow
            'Modified by Morgan 2014/1/2
            'Text4.SetFocus
            If Text19.Enabled Then Text19.SetFocus
         End If
         strControlButton = MsgText(601)
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10.Enabled = False Then
      Exit Sub
   End If
   If Text10 = MsgText(601) Then
      Exit Sub
   End If
   
   'Added by Morgan 2023/9/26 2492 也會放單號 收文號+收據號
   If Text4 = "2492" Then
      Exit Sub
   End If
   'end 2023/9/26
   
   If Text4 <> "2401" Then
      MsgBox MsgText(113), , MsgText(5)
      Cancel = True
      Text10.SetFocus
      Exit Sub
   End If
   If Val(Text11) <> 0 Then
      If CheckAmount(Val(Text11)) = False Then
         MsgBox MsgText(112), , MsgText(5)
         Cancel = True
         Text10.SetFocus
         Exit Sub
      End If
   End If
   If Val(Text6) <> 0 Then
      If CheckAmount(Val(Text6)) = False Then
         MsgBox MsgText(112), , MsgText(5)
         Cancel = True
         Text10.SetFocus
         Exit Sub
      End If
      Text18 = Text10 '2012/9/25 add by sonia 暫收單號放至對沖其他
   End If
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   'Modify by Morgan 2004/2/10
   '摘要都要含客戶名稱
   '2004/3/1還原
   'Modified by Morgan 2014/1/20 +a0t18
   'Modified by Morgan 2014/2/21 銷退產生的暫收摘要加銷退單號
   'adoquery.Open "select sn01||'/'||a0t17||cu04||'/'||'" & Text10 & "' as Remark, a0t05, a0t06 from acc0t0, salesno, customer where a0t05 = sn02 (+) and substr(a0t06, 1, 8) = cu01 (+) and substr(a0t06, 9, 1) = cu02 (+) and a0t01 = '" & Text10 & "'", adoTaie, adOpenStatic, adLockReadOnly
   'Modified by Lydia 2024/11/28 +客戶名稱cuname
   adoquery.Open "select sn01||'/'||decode(a0t17, null, cu04, a0t17)||decode(substr(a0t07,1,1),'I','/'||(a0t07))||'/'||'" & Text10 & "' as Remark, a0t05, a0t06,a0t18,nvl(cu04,nvl(cu05,cu06)) as cuname " & _
                "from acc0t0, salesno, customer where a0t05 = sn02 (+) and substr(a0t06, 1, 8) = cu01 (+) and substr(a0t06, 9, 1) = cu02 (+) and a0t01 = '" & Text10 & "' ", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      'Added by Lydia 2024/11/28 未收文客戶暫收款管制
      If "" & adoquery.Fields("a0t06") = "X03072010" Then
         MsgBox "此暫收款之客戶為" & adoquery.Fields("A0t06") & adoquery.Fields("cuname") & "，不可沖帳 !", vbExclamation
         adoquery.Close
         Cancel = True
         Text10.SetFocus
         Exit Sub
      End If
      'end 2024/11/28
      'add by sonia 2025/5/9 鎖住業務編號及客戶編號
      Text13.Enabled = False
      Text16.Enabled = False
      'end 2025/5/9
      If IsNull(adoquery.Fields("Remark").Value) Then
         Combo2 = MsgText(601)
      Else
         Combo2 = adoquery.Fields("Remark").Value
      End If
      If IsNull(adoquery.Fields("a0t05").Value) Then
         Text13 = MsgText(601)
         Text17 = ""
      Else
         Text13 = adoquery.Fields("a0t05").Value
         Text17 = StaffQuery(Text13)
      End If
      If IsNull(adoquery.Fields("a0t06").Value) Then
         Text16 = MsgText(601)
      Else
         Text16 = adoquery.Fields("a0t06").Value
      End If
      
      'Added by Morgan 2014/1/20
      If Text19 <> "" And Not IsNull(adoquery.Fields("a0t18")) Then
         If adoquery.Fields("a0t18") <> Text19 Then
            MsgBox "暫收單號的公司與收款公司別不同!!", vbExclamation
         End If
      End If
      'end 2014/1/20
      
   Else
      Combo2 = MsgText(601)
      Text13 = MsgText(601)
      Text17 = ""
      Text16 = MsgText(601)
      'add by sonia 2025/5/9 前面有暫收單號時會鎖住
      Text13.Enabled = True
      Text16.Enabled = True
      'end 2025/5/9
   End If
   adoquery.Close
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub

Private Sub Text13_GotFocus()
   TextInverse Text13
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text13_Validate(Cancel As Boolean)
   Text17 = ""
   If Text13 <> MsgText(601) Then
'Modify by Morgan 2007/2/5 員工已離職要提醒
'      If ExistCheck("staff", "st01", Text13, Label15, False) = False Then
'         MsgBox MsgText(45) & Label15, , MsgText(5)
'         Text13.SetFocus
'         Cancel = True
'         TextInverse Text13
'         Exit Sub
'      Else
'         Text17 = StaffQuery(Text13)
'      End If
      If PUB_GetStaffState(Text13.Text, strExc(1), True) = 0 Then
         'modify by sonia 2025/5/9
         'Text13.SetFocus
         If Text13.Enabled = True Then Text13.SetFocus
         'end 2025/5/9
         Cancel = True
         TextInverse Text13
      Else
         Text17 = strExc(1)
      End If
      'add by sonia 2021/1/29
      If SalesNoCheckAccNo(Text4, Text13) = False Then
      End If
      'end 2021/1/29
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
      If ExistCheck("acc090", "a0901", Text14, Label16) = False Then
         Cancel = True
         Exit Sub
      End If
   End If
   If CheckDept(Text4, Text14) = False Then
      MsgBox MsgText(103), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text15_GotFocus()
   TextInverse Text15
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'add by sonia 2023/11/30 從Text15_Validate移過來,否則F11209363第18項次改智權人員一直會被還原
'以本所案號以判別FCP,FCT英日文組
Private Sub Text15_LostFocus()
   If Text15 <> MsgText(601) Then
      If AccNoToSalesNo(Text4, Text15) <> "" Then
         Text13 = AccNoToSalesNo(Text4, Text15)
         Text13_Validate True
      End If
   End If
End Sub
'end 2023/11/30

Private Sub Text15_Validate(Cancel As Boolean)
On Error GoTo Checking
   If Text15 <> MsgText(601) Then
      Text15 = CaseNoZero(Text15)
      If adoquery.State = adStateOpen Then adoquery.Close       'Added by Lydia 2024/11/28
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select pa01 as SystemNo from patent where pa01 = '" & Mid(Text15, 1, Len(Text15) - 9) & "' and pa02 = '" & Mid(Text15, Len(Text15) - 8, 6) & "' and pa03 = '" & Mid(Text15, Len(Text15) - 2, 1) & "' and pa04 = '" & Mid(Text15, Len(Text15) - 1, 2) & "' union " & _
                   "select tm01 as SystemNo from trademark where tm01 = '" & Mid(Text15, 1, Len(Text15) - 9) & "' and tm02 = '" & Mid(Text15, Len(Text15) - 8, 6) & "' and tm03 = '" & Mid(Text15, Len(Text15) - 2, 1) & "' and tm04 = '" & Mid(Text15, Len(Text15) - 1, 2) & "' union " & _
                   "select lc01 as SystemNo from lawcase where lc01 = '" & Mid(Text15, 1, Len(Text15) - 9) & "' and lc02 = '" & Mid(Text15, Len(Text15) - 8, 6) & "' and lc03 = '" & Mid(Text15, Len(Text15) - 2, 1) & "' and lc04 = '" & Mid(Text15, Len(Text15) - 1, 2) & "' union " & _
                   "select hc01 as SystemNo from hirecase where hc01 = '" & Mid(Text15, 1, Len(Text15) - 9) & "' and hc02 = '" & Mid(Text15, Len(Text15) - 8, 6) & "' and hc03 = '" & Mid(Text15, Len(Text15) - 2, 1) & "' and hc04 = '" & Mid(Text15, Len(Text15) - 1, 2) & "' union " & _
                   "select sp01 as SystemNo from servicepractice where sp01 = '" & Mid(Text15, 1, Len(Text15) - 9) & "' and sp02 = '" & Mid(Text15, Len(Text15) - 8, 6) & "' and sp03 = '" & Mid(Text15, Len(Text15) - 2, 1) & "' and sp04 = '" & Mid(Text15, Len(Text15) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount = 0 Then
         MsgBox MsgText(28) & Label17, , MsgText(5)
         Cancel = True
         adoquery.Close
         Exit Sub
      End If
      adoquery.Close
'modify by sonia 2023/11/30 移到Text15_LostFocus,否則F11209363第18項次改智權人員一直會被還原
'      'add by sonia 2021/1/29 以本所案號以判別FCP,FCT英日文組
'      If AccNoToSalesNo(Text4, Text15) <> "" Then
'         Text13 = AccNoToSalesNo(Text4, Text15)
'         Text13_Validate True
'      End If
'      'end 2021/1/29
'end 2023/11/30
   End If
   Exit Sub
Checking:
   MsgBox MsgText(128), , MsgText(5)
   Exit Sub
End Sub

Private Sub Text16_GotFocus()
   TextInverse Text16
End Sub

Private Sub Text16_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text16_Validate(Cancel As Boolean)
   If Text16 <> MsgText(601) Then
      If Len(Text16) = 6 Then
         Text16 = AfterZero(Text16)
      'Add by Morgan 2007/3/1 八碼時要補'0'
      ElseIf Len(Text16) = 8 Then
         Text16 = Text16 & "0"
      'End 2007/3/1
      End If
      If ExistCheck("customer", "cu01", Mid(Text16, 1, 8), Label11, False) = False Then
         If ExistCheck("acc0i0", "a0i01", Text16, Label11, False) = False Then
            If ExistCheck("staff", "st01", Text16, Label11, False) = False Then
               MsgBox MsgText(28) & Label18, , MsgText(5)
               Cancel = True
               Exit Sub
            End If
         End If
      End If
   End If
End Sub

Private Sub Text18_GotFocus()
   TextInverse Text18
End Sub

Private Sub Text18_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text19_GotFocus()
   TextInverse Text19
   CloseIme
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Modified by Morgan 2020/4/13 +l
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("J") And KeyAscii <> Asc("L") Then
      KeyAscii = 0
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_Change()
   Text5 = A0102Query(Text4)
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
   If Text4 <> MsgText(601) Then
      If ExistCheck("acc010", "a0101", Text4, Label5, False) = False Then
         MsgBox MsgText(45) & Label5, , MsgText(5)
         Cancel = True
         Exit Sub
      End If
   Else
      Exit Sub
   End If
   RemarkShow
   Select Case Text4
      Case "2401"
         Text10 = MsgText(806)
      Case Else
         Text10 = MsgText(601)
         'add by sonia 2025/5/9 2401且有暫收單號時鎖住業務編號及客戶編號，非此情形要解開
         Text13.Enabled = True
         Text16.Enabled = True
         'end 2025/5/9
   End Select
   'modify by sonia 2021/1/29 加傳本所案號以判別FCP,FCT英日文組
   'If AccNoToSalesNo(Text4) <> "" Then
   '   Text13 = AccNoToSalesNo(Text4)
   If AccNoToSalesNo(Text4, Text15) <> "" Then
      Text13 = AccNoToSalesNo(Text4, Text15)
   'end 2021/1/29
      Text17 = StaffQuery(Text13)
   End If
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
'   If strSaveConfirm = MsgText(3) Then
'      adoquery.CursorLocation = adUseClient
'      adoquery.Open "select a0e02 from acc0e0 where a0e02 = '" & Text7 & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adoquery.RecordCount <> 0 Then
'         MsgBox MsgText(148), , MsgText(5)
'         adoquery.Close
'         Cancel = True
'         Text7.SetFocus
'         Exit Sub
'      End If
'      adoquery.Close
'   End If
   RemarkShow
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub

Private Sub Text8_Validate(Cancel As Boolean)
   RemarkShow
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

'*************************************************
'  清除查詢顯示
'
'*************************************************
Public Sub AdodcClear()
   Text3 = ""
   Text12 = ""
   Text20 = ""
   Text4 = ""
   Text4.Tag = "" 'Added by Morgan 2015/1/30
   Text6 = ""
   Text11 = ""
   Text7 = ""
   Text9 = ""
   Text8 = ""
   'Add by Amy 2020/06/30
   Text7.Tag = ""
   Text8.Tag = ""
   Text9.Tag = ""
   'end 2020/06/30
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text10 = ""
   Combo1 = ""
   Combo2 = ""
   Text13 = ""
   Text17 = ""
   Text14 = ""
   Text15 = ""
   Text16 = ""
   Text18 = ""
   Text19 = "" 'Added by Morgan 2013/12/19
   'Added by Morgan 2020/12/25
   lblA1P22 = ""
   Text19.Enabled = Text4.Enabled
   'end 2020/12/25
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = adoacc0l0.Bookmark & MsgText(35) & adoacc0l0.RecordCount
End Sub

'*************************************************
'  關閉分錄欄位輸入狀態
'
'*************************************************
Public Sub FormDisabled()
   Text2.Enabled = True
   Text4.Enabled = False
   Text6.Enabled = False
   Text11.Enabled = False
   Text7.Enabled = False
   Text9.Enabled = False
   Text13.Enabled = False
   Text8.Enabled = False
   MaskEdBox2.Enabled = False
   Combo1.Enabled = False
   Text10.Enabled = False
   Combo2.Enabled = False
   Text14.Enabled = False
   Text15.Enabled = False
   Text16.Enabled = False
   Text18.Enabled = False
   Text19.Enabled = False 'Added by Morgan 2013/12/19
   Command2.Enabled = False
   Command1.Enabled = True
   Command5.Enabled = True 'Added by Morgan 2014/1/27
   MaskEdBox1.Enabled = False 'Added by Morgan 2022/6/29
End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'
'*************************************************
Public Sub FormEnabled()
   Text2.Enabled = False
   Text4.Enabled = True
   Text6.Enabled = True
   Text11.Enabled = True
   Text7.Enabled = True
   Text9.Enabled = True
   Text13.Enabled = True
   Text8.Enabled = True
   MaskEdBox2.Enabled = True
   Combo1.Enabled = True
   Text10.Enabled = True
   Combo2.Enabled = True
   Text14.Enabled = True
   Text15.Enabled = True
   Text16.Enabled = True
   Text18.Enabled = True
   If lblA1P22 = "" Then Text19.Enabled = True 'Added by Morgan 2013/12/19 'Modified by Morgan 2020/12/25 有傳票號不可改公司別
   Command2.Enabled = True
   Command1.Enabled = False
   Command5.Enabled = False 'Added by Morgan 2014/1/27
   'Added by Morgan 2022/6/29 有傳票號不可改收款日期
   If Text2 <> "" Then
      If CheckExistA1p22("", "A", Text2.Text) = False Then
         MaskEdBox1.Enabled = True
      End If
      'add by sonia 2024/1/5 已過帳後換收據的情形會由電腦中心拿掉A1P22由使用者自行換收據，此時收款日期也不能改F11208888
      If Mid(Val(FCDate(MaskEdBox1.Text)), 1, 5) < Mid(CompDate(1, -1, strSrvDate(1)) - 19110000, 1, 5) Then
         MaskEdBox1.Enabled = False
      End If
      'end 2024/1/5
   End If
   'end 2022/6/29
   'add by sonia 2025/5/9 2401且有暫收單號時鎖住業務編號及客戶編號
   If Text4 = "2401" And Text10 <> "" Then
      Text13.Enabled = False
      Text16.Enabled = False
   End If
   'end 2025/5/9
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   If Text9 <> MsgText(601) Then
      If ExistCheck("acc0g0", "a0g01", Text9, Label10, False) = False Then
         MsgBox MsgText(45) & Label10, , MsgText(5)
         Cancel = True
         Text9.SetFocus
         TextInverse Text9
         Exit Sub
      End If
   End If
   RemarkShow
End Sub

'*************************************************
'  摘要顯示
'
'*************************************************
Public Sub RemarkShow()
   If Mid(Text4, 1, 4) = "1130" Then
      Combo2 = FCDate(MaskEdBox2.Text) & "/" & Text7 & "/" & Text8 & "/" & A0g02Query(Text9)
      Exit Sub
   End If
End Sub

'*************************************************
'  檢查暫收款金額
'
'*************************************************
Public Function CheckAmount(douValue As Double) As Boolean
   If adoquery.State = adStateOpen Then adoquery.Close       'Added by Lydia 2024/11/28
   adoquery.CursorLocation = adUseClient
   'Modify by Morgan 2004/2/6
   '加會計科目條件
   adoquery.Open "select a0t08, a1p07 from acc0t0, acc1p0 where a0t01 = a1p23 (+) and a0t01 = '" & Text10 & "' AND a1p05(+)='" & Text4 & "' and a1p04(+)<>'" & Text2 & "' order by a1p07 desc", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields(0).Value) Then
         CheckAmount = False
      'Modify by Morgan 2004/2/6
      '要檢查金額
      ElseIf adoquery.Fields(0).Value <> douValue Then
         CheckAmount = False
      Else
         If IsNull(adoquery.Fields("a1p07").Value) Then
            CheckAmount = True
         Else
            If adoquery.Fields("a1p07").Value = 0 Then
               CheckAmount = True
            Else
               CheckAmount = False
            End If
         End If
      End If
   Else
      CheckAmount = False
   End If
   adoquery.Close
End Function

'*************************************************
'  重新整理國內收款資料
'
'*************************************************
Public Sub Acc0l0Refresh()
On Error GoTo Checking
   If adoacc0l0.State = adStateOpen Then
      adoacc0l0.Close
   End If
   adoacc0l0.CursorLocation = adUseClient
   adoacc0l0.MaxRecords = intMax
   adoacc0l0.Open "select * from acc0l0 where a0l01 >= '" & Text2 & "' order by a0l01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
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
   If Text3 = Text12 Then
      'Added by Morgan 2013/12/27 所有公司別都沒有不平
      'CreDebCheck = MsgText(602)
      strExc(0) = "select * from (select a1p01 COMP,sum(a1p07) AMT1,sum(a1p08) AMT2,sum(a1p07)-sum(a1p08) AMT3 from acc1p0 where a1p04='" & Text2 & "' group by a1p01) where AMT3<>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 0 Then
         CreDebCheck = MsgText(602)
      End If
   End If
End Function

'Added by Morgan 2013/12/26
'存檔前分錄內容檢查
Public Function SaveCheck() As Boolean
   Dim bCancel As Boolean, bMainComponyExists As Boolean
   Dim strMsg As String 'Added by Morgan 2022/6/29
   
   'Added by Morgan 2022/6/29 分錄存檔跑較久，日期先檢查
   If MaskEdBox1.Enabled = True Then
      If ChkWorkData("1", DBDATE(MaskEdBox1), strMsg) = False Then
          MsgBox Label1 & strMsg, , MsgText(5)
          MaskEdBox1.SetFocus
          Exit Function
      End If
   End If
   'end 2022/6/29
            
   m_SaveCheck = True
   bMainComponyExists = False
   With Adodc1.Recordset
   .MoveFirst
   Do While Not .EOF
      strSerialNo = .Fields("a1p03").Value
      'Added by Morgan 2023/8/25
      If Val(strSerialNo) = 1 And Text21 <> .Fields("a1p01").Value Then
         m_SaveCheck = False
         MsgBox "主要公司別必須和第1筆分錄的公司別相同!!", vbCritical
         Exit Function
      End If
      'end 2023/8/25
      AdodcShow
      'Added by Morgan 2023/5/30 暫收款無單號提醒
      If Text4 = "2401" Then
         If Text10 = "" Then
            If MsgBox("請輸入暫收款單號，是否輸入？", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
               Text10.SetFocus
               m_SaveCheck = False
               Exit Function
            End If
         End If
      End If
      'end 2023/5/30
      If Acc1p0Save = False Then
         m_SaveCheck = False
         Exit Function
      End If
      If Me.Text19 = Me.Text21 Then
         bMainComponyExists = True
      End If
      .MoveNext
   Loop
   .MoveFirst
   End With
   
   'Added by Morgan 2014/1/27
   If bMainComponyExists = False Then
      m_SaveCheck = False
      MsgBox "主要公司別的分錄不存在!!", vbCritical
      Exit Function
   End If
   'end 2014/1/27
   
   SaveCheck = True
   m_SaveCheck = False
End Function

'Added by Morgan 2015/5/28
Public Function DeleteCheck() As Boolean
   Dim stSQL As String, intR As Integer
   Dim stMsg As String
   
   'Added by Morgan 2025/2/25 檢查是否有補扣繳
   strExc(0) = "select distinct a1u02 from acc1u0 where a1u02 in (select a0m02 from acc0m0 where a0m01='" & Text2 & "') and a1u01=a1u03"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      MsgBox "下列收據已有補扣繳紀錄，本次收款不可刪除！" & vbCrLf & RsTemp.GetString(), vbExclamation
      Exit Function
   End If
   'end 2025/2/25
   
   'Added by Morgan 2016/2/18
   '貸方暫收款要檢查若有銷退記錄時不可刪除
   strExc(0) = "select a0s01,a0s02 from acc0s0 where a0s02 in (select a0t01 from acc0T0 where a0t07='" & Text2 & "')"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      MsgBox "本次收款轉出之暫收款[ " & RsTemp(1) & " ]已有銷退記錄[ " & RsTemp(0) & " ]，不可刪除！", vbExclamation
      Exit Function
   End If
      
   '檢查收據若有退費記錄時提醒
   strExc(0) = "select a0s01,a0s02 from acc0s0 where a0s02 in (select a0m02 from acc0m0 where a0m01='" & Text2 & "') and (a0s06>0 or a0s07>0)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If MsgBox("本次收款的收據[ " & RsTemp(1) & " ]有退費記錄[ " & RsTemp(0) & " ]，是否確定要刪除？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
         Exit Function
      End If
   End If
   'end 2016/2/18
      
   stSQL = "select axc01,axc02 from acc0m0,acc431 where a0m01='" & Text2 & "' and axc02(+)=a0m02 and axc03(+)=a0m01 and axc01 is not null"
   intR = 1
   Set adoquery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      stMsg = "請款單號" & vbTab & "發票號"
      Do While Not adoquery.EOF
         stMsg = stMsg & vbCrLf & adoquery("axc02") & vbTab & adoquery("axc01")
         adoquery.MoveNext
      Loop
      MsgBox "下列請款單已開發票，本次收款不可刪除!" & vbCrLf & vbCrLf & stMsg & vbCrLf & vbCrLf & "(若要取消收款請先作廢發票)", vbExclamation, "收款刪除檢查"
      DeleteCheck = False
   Else
      DeleteCheck = True
   End If
   If adoquery.State = adStateOpen Then adoquery.Close
End Function

'Add by Amy 2018/06/05 從aacc_lst/next/pre/fst 搬回
Public Sub Frmacc1150_First()
    'Add by Amy 2018/06/05 41字頭及7121不可輸小數
    If ChkDot = True Then
        MsgBox "41字頭或7121科目不可輸入小數！", , MsgText(5)
        Exit Sub
     End If
     'end 2018/06/05
      'CreDebCheck 'Removed by Morgan 2021/5/31
      If CreDebCheck <> MsgText(602) Then
         MsgBox MsgText(11), , MsgText(5)
         Exit Sub
      End If
      If adoacc0l0.RecordCount <> 0 Then
'         .adoacc0l0.MoveFirst
         adoaccsum.CursorLocation = adUseClient
         adoaccsum.Open "select min(a0l01) from acc0l0", adoTaie, adOpenStatic, adLockReadOnly
         If adoaccsum.EOF = False Then
            If IsNull(adoaccsum.Fields(0).Value) = False Then
              Text2 = adoaccsum.Fields(0).Value
            End If
         End If
         adoaccsum.Close
         Acc0l0Refresh
         FormShow
         AdodcRefresh
         SumShow
         RecordShow
      End If
End Sub

Public Sub Frmacc1150_Last()
    'Add by Amy 2018/06/05 41字頭及7121不可輸小數
    If ChkDot = True Then
        MsgBox "41字頭或7121科目不可輸入小數！", , MsgText(5)
        Exit Sub
     End If
     'end 2018/06/05
      CreDebCheck
      If CreDebCheck <> MsgText(602) Then
         MsgBox MsgText(11), , MsgText(5)
         Exit Sub
      End If
      If adoacc0l0.RecordCount <> 0 Then
'         .adoacc0l0.MoveLast
         adoaccsum.CursorLocation = adUseClient
         'Modified by Morgan 2022/6/30
         'adoaccsum.Open "select max(a0l01) from acc0l0", adoTaie, adOpenStatic, adLockReadOnly
         adoaccsum.Open "select max(a0l01) from acc0l0 where a0l02=(select max(a0l02) from acc0l0)", adoTaie, adOpenStatic, adLockReadOnly
         'end 2022/6/30
         If adoaccsum.EOF = False Then
            If IsNull(adoaccsum.Fields(0).Value) = False Then
              Text2 = adoaccsum.Fields(0).Value
            End If
         End If
         adoaccsum.Close
         Acc0l0Refresh
         FormShow
         AdodcRefresh
         SumShow
         RecordShow
      End If
End Sub

Public Sub Frmacc1150_Next()
    'Add by Amy 2018/06/05 41字頭及7121不可輸小數
    If ChkDot = True Then
        MsgBox "41字頭或7121科目不可輸入小數！", , MsgText(5)
        Exit Sub
     End If
     'end 2018/06/05
      CreDebCheck
      If CreDebCheck <> MsgText(602) Then
         MsgBox MsgText(11), , MsgText(5)
         Exit Sub
      End If
      If adoacc0l0.EOF = False Then
         adoacc0l0.MoveNext
         If adoacc0l0.EOF Then
'            .adoacc0l0.MoveLast
'            MsgBox MsgText(8), , MsgText(5)
            Acc0l0Refresh
         End If
         FormShow
         AdodcRefresh
         SumShow
         RecordShow
      End If
End Sub

Public Sub Frmacc1150_Previous()
    'Add by Amy 2018/06/05 41字頭及7121不可輸小數
    If ChkDot = True Then
        MsgBox "41字頭或7121科目不可輸入小數！", , MsgText(5)
        Exit Sub
     End If
     'end 2018/06/05
      CreDebCheck
      If CreDebCheck <> MsgText(602) Then
         MsgBox MsgText(11), , MsgText(5)
         Exit Sub
      End If
      If adoacc0l0.BOF = False Then
         adoacc0l0.MovePrevious
         If adoacc0l0.BOF Then
'            .adoacc0l0.MoveFirst
             adoaccsum.CursorLocation = adUseClient
             adoaccsum.Open "select max(a0l01) from acc0l0 where a0l01 <  '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
            If adoaccsum.EOF = False Then
                If IsNull(adoaccsum.Fields(0).Value) = False Then
                  Text2 = adoaccsum.Fields(0).Value
               End If
            Else
               MsgBox MsgText(7), , MsgText(5)
            End If
             adoaccsum.Close
             Acc0l0Refresh
         End If
         FormShow
         AdodcRefresh
         SumShow
         RecordShow
      End If
End Sub
'end 2018/06/05

'Add by Amy 2018/06/06 41字頭及7121且業務是S部門不可輸小數
Public Function ChkDot() As Boolean
    strExc(0) = "Select * From Acc1p0,Staff Where a1p04='" & Text2 & "' " & _
                      "And (Substr(a1p05,1,2)='41' or a1p05='7121' ) And (instr(a1p07,'.')>0 or instr(a1p08,'.')>0) " & _
                      "And a1p16 is not null And a1p16=st01(+) And SubStr(St15,1,1)='S' "
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        ChkDot = True
    End If
End Function

'Add by Amy 2020/06/30 由acc_sav搬過來
Public Sub Frmacc1150_Save()
Dim strYes As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'add by sonia 2020/4/24
Dim arrA1p01() As String '傳票號
Dim arrA1p22() As String '傳票號
Dim intPos As Integer
'end 2020/4/24

   On Error GoTo Checking
      
   With Frmacc1150
      
      If Text2 = MsgText(601) Then
         MsgBox MsgText(10) & Label3, , MsgText(5)
         strControlButton = MsgText(602)
         Text2.SetFocus
         Exit Sub
      Else
         If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
            MsgBox Label1 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            MaskEdBox1.SetFocus
            Exit Sub
         Else
            If DateCheck(.MaskEdBox1.Text) = MsgText(603) Then
               MsgBox Label1 & MsgText(63), , MsgText(5)
               strControlButton = MsgText(602)
               MaskEdBox1.SetFocus
               Exit Sub
            End If
         End If
      End If
      
      'Added by Morgan 2014/1/2
      If Text21 = "" Then
         MsgBox "主要公司別不可空白！"
         strControlButton = MsgText(602)
         Exit Sub
      End If
      'end 2014/1/2
      
      
        'Add By Cheng 2004/05/06
        '在修改狀態若傳票已過帳, 則不可更新
        If strSaveConfirm = MsgText(4) Then
            'Modified by Morgan 2014/6/24 會有1或J公司兩家
            'If strCon10 <> "" Then
            'modify by sonia 2020/4/24
'            If .m_A1P22_1 & .m_A1P22_J <> "" Then
'               'StrSQLa = "Select Count(*) From ACC021 Where AX202='" & strCon10 & "' And AX210 Is Not Null "
'               StrSQLa = "Select ax202||'(1公司)' From ACC021 Where  AX201='1' and ax202='" & .m_A1P22_1 & "' And AX210 Is Not Null "
'               StrSQLa = StrSQLa & " union Select ax202||'(J公司)' From ACC021 Where  AX201='J' and ax202='" & .m_A1P22_J & "' And AX210 Is Not Null "
'               rsA.CursorLocation = adUseClient
'               'rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'               rsA.Open StrSQLa, adoTaie, adOpenForwardOnly, adLockReadOnly
'               'If Val("" & rsA.Fields(0).Value) > 0 Then
'               If Not (rsA.EOF And rsA.BOF) Then
'                   'MsgBox "傳票" & strCon10 & "已過帳, 不可更改資料!!!", vbExclamation + vbOKOnly
'                   MsgBox "傳票" & rsA(0) & "已過帳, 不可更改資料!!!", vbExclamation + vbOKOnly
'                   If rsA.State <> adStateClosed Then rsA.Close
'                   Set rsA = Nothing
'                   strControlButton = MsgText(602)
'                   Exit Sub
'               End If
'            'end 2014/6/24
'               If rsA.State <> adStateClosed Then rsA.Close
'               Set rsA = Nothing
'            End If
            arrA1p01 = Split(.strA1P01s, ";")
            arrA1p22 = Split(.strA1P22s, ";")
            For intPos = LBound(arrA1p22) To UBound(arrA1p22)
               If arrA1p22(intPos) <> "" Then
                  StrSQLa = "Select ax202||'(" & arrA1p01(intPos) & "公司)' From ACC021 Where AX201='" & arrA1p01(intPos) & "' and ax202='" & arrA1p22(intPos) & "' And AX210 Is Not Null "
                  rsA.CursorLocation = adUseClient
                  rsA.Open StrSQLa, adoTaie, adOpenForwardOnly, adLockReadOnly
                  If Not (rsA.EOF And rsA.BOF) Then
                      MsgBox "傳票" & rsA(0) & "已過帳, 不可更改資料!!!", vbExclamation + vbOKOnly
                      If rsA.State <> adStateClosed Then rsA.Close
                      Set rsA = Nothing
                      strControlButton = MsgText(602)
                      Exit Sub
                  End If
                  If rsA.State <> adStateClosed Then rsA.Close
                  Set rsA = Nothing
               End If
            Next
            'end 2020/4/24
        End If
        'End
           
      If strSaveConfirm = MsgText(3) Then
         If adoacc0l0.RecordCount <> 0 Then
            adoacc0l0.Find "a0l01 = '" & Text2 & "'", 0, adSearchForward, 1
            If adoacc0l0.EOF = False Then
               Exit Sub
            End If
         End If

         adoacc0l0.AddNew
         'Add by Morgan 2007/7/18
         adoacc0l0.Fields("a0l03").Value = 0
         adoacc0l0.Fields("a0l04").Value = 0
         '.adoacc0l0.Fields("a0l05").Value = 0   '2013/10/11 CANCEL BY SONIA 無用
         adoacc0l0.Fields("a0l05").Value = Text21 'Added by Morgan 2013/12/27
         adoacc0l0.Fields("a0l08").Value = 0
         adoacc0l0.Fields("a0l09").Value = 0
         adoacc0l0.Fields("a0l10").Value = 0
         'end 2007/7/18
      'Added by Morgan 2012/3/23
      Else
         strExc(1) = Val(FCDate(MaskEdBox1.Text)) \ 10000
         If Val(strExc(1)) > 0 Then
            strExc(0) = "select a0m02 from acc0m0 where a0m01='" & Text2 & "' and a0m07<>" & strExc(1)
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If MsgBox("下列收據之扣繳年度與收款日期年度不符，是否應更正??" & vbCrLf & vbCrLf & RsTemp.GetString, vbYesNo + vbDefaultButton1) = vbYes Then
                  strControlButton = MsgText(602)
                  Exit Sub
               End If
            End If
         End If
      'end 2012/3/23
      End If
      
      adoacc0l0.Fields("a0l01").Value = Text2
      If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
         adoacc0l0.Fields("a0l02").Value = Val(FCDate(MaskEdBox1.Text))
      Else
         adoacc0l0.Fields("a0l02").Value = Null
      End If
      If Text1 <> MsgText(601) Then
         adoacc0l0.Fields("a0l07").Value = Text1
      Else
         adoacc0l0.Fields("a0l07").Value = Null
      End If
      'Remove by Morgan 2007/7/18 新增時才要設0
      '.adoacc0l0.Fields("a0l03").Value = 0
      '.adoacc0l0.Fields("a0l04").Value = 0
      '.adoacc0l0.Fields("a0l05").Value = 0
      '.adoacc0l0.Fields("a0l08").Value = 0
      '.adoacc0l0.Fields("a0l09").Value = 0
      '.adoacc0l0.Fields("a0l10").Value = 0
      'end 2007/7/18
      If strSaveConfirm = MsgText(3) Then
         adoacc0l0.Fields("a0l11").Value = Val(strSrvDate(2))
         adoacc0l0.Fields("a0l12").Value = ServerTime
         adoacc0l0.Fields("a0l13").Value = strUserNum
      Else
         adoacc0l0.Fields("a0l14").Value = Val(strSrvDate(2))
         adoacc0l0.Fields("a0l15").Value = ServerTime
         adoacc0l0.Fields("a0l16").Value = strUserNum
      End If
      adoacc0l0.UpdateBatch
      If strSaveConfirm <> MsgText(3) Then
         'Modified by Morgan 2014/6/24 會有1或J公司兩家
         'If strCon10 <> "" Then
         '   adoTaie.Execute "update acc1p0 set a1p22 = '" & strCon10 & "', a1p27 = 'Y' where a1p01 = '1' and a1p02 = 'A' and a1p04 = '" & .Text2 & "'"
         'End If
         'modify by sonia 2020/4/24 考慮會有3家作帳公司別
         'If .m_A1P22_1 <> "" Then
         '   adoTaie.Execute "update acc1p0 set a1p22 = '" & .m_A1P22_1 & "', a1p27 = 'Y' where a1p01 = '1' and a1p02 = 'A' and a1p04 = '" & .Text2 & "'", intI
         'End If
         'If .m_A1P22_J <> "" Then
         '   adoTaie.Execute "update acc1p0 set a1p22 = '" & .m_A1P22_J & "', a1p27 = 'Y' where a1p01 = 'J' and a1p02 = 'A' and a1p04 = '" & .Text2 & "'", intI
         'End If
         arrA1p01 = Split(strA1P01s, ";")
         arrA1p22 = Split(strA1P22s, ";")
         For intPos = LBound(arrA1p22) To UBound(arrA1p22)
            If arrA1p22(intPos) <> "" Then
               adoTaie.Execute "update acc1p0 set a1p22 = '" & arrA1p22(intPos) & "', a1p27 = 'Y' where a1p01 = '" & arrA1p01(intPos) & "' and a1p02 = 'A' and a1p04 = '" & Text2 & "'", intI
            End If
         Next
         'end 2020/4/24
         'end 2014/6/24
      End If

      RecordShow
      If Text2.Enabled Then
         Text2.SetFocus
      End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If

   MsgBox Err.Description, , MsgText(5)
   End With
End Sub
'Added by Morgan 2021/1/18
'檢查是否為案源收據
Private Function ChkIsLawCase(pA0K01 As String) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   stSQL = "select los02 from acc0j0,caseprogress,lawofficesource where a0j13='" & pA0K01 & "' and cp09(+)=a0j01 and los15(+)=cp162 and los15 is not null"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      ChkIsLawCase = True
   End If
End Function
