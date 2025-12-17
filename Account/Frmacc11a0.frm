VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc11a0 
   AutoRedraw      =   -1  'True
   Caption         =   "暫收款作業"
   ClientHeight    =   5440
   ClientLeft      =   50
   ClientTop       =   460
   ClientWidth     =   8940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5440
   ScaleWidth      =   8940
   Begin VB.TextBox Text26 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   7070
      TabIndex        =   62
      Top             =   360
      Width           =   1500
   End
   Begin VB.TextBox Text25 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   2175
      TabIndex        =   61
      Top             =   330
      Width           =   2750
   End
   Begin VB.TextBox Text24 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1590
      MaxLength       =   1
      TabIndex        =   3
      Top             =   330
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "未沖"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Left            =   5970
      TabIndex        =   59
      Top             =   3090
      Width           =   735
   End
   Begin VB.OptionButton Option2 
      Caption         =   "全部"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Left            =   6720
      TabIndex        =   58
      Top             =   3090
      Width           =   780
   End
   Begin VB.TextBox Text23 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   21
      Top             =   4620
      Width           =   1572
   End
   Begin VB.TextBox Text22 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   19
      Top             =   4620
      Width           =   1572
   End
   Begin VB.TextBox Text21 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4080
      MaxLength       =   9
      TabIndex        =   20
      Top             =   4620
      Visible         =   0   'False
      Width           =   1572
   End
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
      Height          =   300
      Left            =   6840
      MaxLength       =   3
      TabIndex        =   18
      Top             =   4320
      Width           =   528
   End
   Begin VB.TextBox Text17 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1590
      TabIndex        =   8
      Top             =   1320
      Width           =   1572
   End
   Begin VB.TextBox Text20 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   49
      Top             =   3075
      Width           =   855
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4560
      TabIndex        =   48
      Top             =   3075
      Width           =   1092
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6840
      MaxLength       =   14
      TabIndex        =   12
      Top             =   3690
      Width           =   1572
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1920
      TabIndex        =   46
      Top             =   3690
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   2760
      Picture         =   "Frmacc11a0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   24
      Width           =   350
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4080
      TabIndex        =   17
      Top             =   4305
      Width           =   1572
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc11a0.frx":0102
      Height          =   1365
      Left            =   240
      TabIndex        =   24
      Top             =   1680
      Width           =   8295
      _ExtentX        =   14623
      _ExtentY        =   2417
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
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
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
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
            SubFormatType   =   1
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
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column03 
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
      BeginProperty Column04 
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
            ColumnWidth     =   1690.016
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1250.079
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   970.016
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   3899.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   6529.89
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   30
      Top             =   1620
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   2117
      _ExtentY        =   635
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
   Begin VB.TextBox Text11 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   7110
      TabIndex        =   25
      Top             =   1005
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7680
      Picture         =   "Frmacc11a0.frx":0117
      Style           =   1  '圖片外觀
      TabIndex        =   23
      ToolTipText     =   "取消"
      Top             =   3060
      Width           =   450
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      TabIndex        =   40
      Top             =   3075
      Width           =   1092
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   14
      Top             =   4005
      Width           =   1572
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6840
      MaxLength       =   12
      TabIndex        =   15
      Top             =   4005
      Width           =   1572
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   13
      Top             =   4005
      Width           =   1572
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5040
      MaxLength       =   14
      TabIndex        =   11
      Top             =   3690
      Width           =   1572
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      MaxLength       =   6
      TabIndex        =   10
      Top             =   3690
      Width           =   1572
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   4788
      MaxLength       =   9
      TabIndex        =   5
      Top             =   660
      Width           =   1236
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   1590
      MaxLength       =   5
      TabIndex        =   4
      Top             =   660
      Width           =   876
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3840
      MaxLength       =   1
      TabIndex        =   2
      Top             =   24
      Width           =   612
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1572
      MaxLength       =   15
      TabIndex        =   0
      Top             =   24
      Width           =   1215
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1590
      TabIndex        =   6
      Top             =   1005
      Width           =   1575
      _ExtentX        =   2787
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.5
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
      Left            =   4470
      TabIndex        =   7
      Top             =   1005
      Width           =   1575
      _ExtentX        =   2787
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.5
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
      Left            =   1320
      TabIndex        =   16
      Top             =   4305
      Width           =   1575
      _ExtentX        =   2787
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "沖銷傳票或備註"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Left            =   5300
      TabIndex        =   63
      Top             =   390
      Width           =   1680
   End
   Begin MSForms.TextBox Text18 
      Height          =   336
      Left            =   2472
      TabIndex        =   52
      Top             =   660
      Width           =   972
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      MaxLength       =   5
      Size            =   "7223;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text13 
      Height          =   330
      Left            =   6030
      TabIndex        =   44
      Top             =   660
      Width           =   2535
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      Size            =   "4471;582"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   330
      Left            =   1320
      TabIndex        =   22
      Top             =   4935
      Width           =   7095
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "12515;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text12 
      Height          =   330
      Left            =   4470
      TabIndex        =   9
      Top             =   1320
      Width           =   4095
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "7223;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label27 
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
      TabIndex        =   60
      Top             =   330
      Width           =   1245
   End
   Begin VB.Label Label26 
      BackStyle       =   0  '透明
      Caption         =   "對沖(其)"
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
      Left            =   5880
      TabIndex        =   57
      Top             =   4620
      Width           =   975
   End
   Begin VB.Label Label24 
      BackStyle       =   0  '透明
      Caption         =   "對沖(本)"
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
      Left            =   360
      TabIndex        =   56
      Top             =   4620
      Width           =   975
   End
   Begin VB.Label Label23 
      BackStyle       =   0  '透明
      Caption         =   "對沖(客)"
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
      Left            =   3120
      TabIndex        =   55
      Top             =   4620
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label22 
      BackStyle       =   0  '透明
      Caption         =   " "
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
      Left            =   5670
      TabIndex        =   54
      Top             =   4740
      Width           =   975
   End
   Begin VB.Label Label21 
      BackStyle       =   0  '透明
      Caption         =   "部門別"
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
      Left            =   5880
      TabIndex        =   53
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label20 
      BackStyle       =   0  '透明
      Caption         =   "暫收款金額"
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
      TabIndex        =   51
      Top             =   1350
      Width           =   1245
   End
   Begin VB.Label Label25 
      BackStyle       =   0  '透明
      Caption         =   "筆數"
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
      Left            =   360
      TabIndex        =   50
      Top             =   3075
      Width           =   855
   End
   Begin VB.Label Label19 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "貸方金額"
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
      Left            =   6840
      TabIndex        =   47
      Top             =   3450
      Width           =   1575
   End
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "票別"
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
      Left            =   3120
      TabIndex        =   45
      Top             =   4305
      Width           =   975
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "摘要"
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
      Left            =   360
      TabIndex        =   43
      Top             =   4935
      Width           =   855
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "備註"
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
      Left            =   3270
      TabIndex        =   42
      Top             =   1350
      Width           =   975
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "處理單號"
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
      Left            =   6150
      TabIndex        =   41
      Top             =   1035
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   4560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1920
      Left            =   225
      Top             =   3405
      Width           =   8295
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "合計"
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
      Left            =   2640
      TabIndex        =   39
      Top             =   3075
      Width           =   615
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "收票銀行"
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
      Left            =   3120
      TabIndex        =   38
      Top             =   4005
      Width           =   975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "到期日期"
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
      Left            =   360
      TabIndex        =   37
      Top             =   4305
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "收票帳號"
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
      Left            =   5880
      TabIndex        =   36
      Top             =   4005
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "票據號碼"
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
      Left            =   360
      TabIndex        =   35
      Top             =   4005
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "借方金額"
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
      Left            =   5040
      TabIndex        =   34
      Top             =   3450
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "會計科目"
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
      Left            =   360
      TabIndex        =   33
      Top             =   3450
      Width           =   4335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "欲處理日期"
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
      Left            =   3270
      TabIndex        =   32
      Top             =   1035
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "輸入日期"
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
      TabIndex        =   31
      Top             =   1035
      Width           =   1245
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "客戶編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3588
      TabIndex        =   30
      Top             =   696
      Width           =   972
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
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
      TabIndex        =   29
      Top             =   660
      Width           =   1245
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "(1.暫收款 2.溢收轉入 3.銷帳退費轉入)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4560
      TabIndex        =   28
      Top             =   24
      Width           =   4212
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "類別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3240
      TabIndex        =   27
      Top             =   24
      Width           =   612
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "暫收款單號"
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
      TabIndex        =   26
      Top             =   30
      Width           =   1245
   End
End
Attribute VB_Name = "Frmacc11a0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoacc0t0 As New ADODB.Recordset
Public adoacc1p0 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Dim strSerialNo As String
Public strDocNo As String
Dim m_ActiveControlName As String 'Added by Morgan 2023/10/6
Dim m_UpdStatus As String, m_UpdMsg As String 'Added by Lydia 2024/11/28 未收文客戶暫收款管制：是否為特殊修改，彈訊息

Private Sub Combo1_GotFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   OpenIme
End Sub

Private Sub Combo1_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   m_ActiveControlName = Me.ActiveControl.Name 'Added by Morgan 2023/10/12
End Sub

Private Sub Combo1_LostFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   CloseIme
   m_ActiveControlName = "" 'Added by Morgan 2023/10/12
End Sub

Private Sub Command1_Click()
   If Adodc1.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc1.Recordset.Fields("a1p22").Value) = False Then
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select ax210 from acc021 where ax201 = '" & Adodc1.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc1.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(155), , MsgText(5)
            Text5.SetFocus
            adoquery.Close
            Exit Sub
         End If
         adoquery.Close
      End If
   End If
   AdodcDelete
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Command3_Click()

   'Modified by Morgan 2023/8/1 取消上下筆功能
   '   '2022/7/28 add by sonia
   '   If Text1 <> "" Then
   '      OpenTable
   '      If adoacc0t0.RecordCount <> 0 Then
   '          adoacc0t0.MoveFirst
   '          'RecordShow
   '      End If
   '   End If
   '   'end 2022/7/28
   '   If adoacc0t0.RecordCount = 0 Or Text1 = MsgText(601) Then
   '      'add by sonia 2022/7/28
   '      Text2 = "": Text24 = "": Text25 = "": Text3 = "": Text18 = "": Text4 = "": Text13 = ""
   '      Text11 = "": Text12 = "": Text17 = "": MaskEdBox1.Text = MsgText(29): MaskEdBox2.Text = MsgText(29)
   '      MsgBox MsgText(33), , MsgText(5)
   '      'end 2022/7/28
   '      Exit Sub
   '   End If
   '   adoacc0t0.Find "a0t01 = '" & Text1 & "'", 0, adSearchForward, 1
   '   If adoacc0t0.EOF = False Then
   '      FormShow
   '      AdodcRefresh
   '      SumShow
   '      RecordShow
   '   Else
   '      MsgBox MsgText(33), , MsgText(5)
   '      adoacc0t0.MoveFirst
   '   End If
   If Text1 <> "" Then
      OpenTable
      If adoacc0t0.EOF Then
         Text2 = "": Text24 = "": Text25 = "": Text3 = "": Text18 = "": Text4 = "": Text13 = ""
         Text11 = "": Text12 = "": Text17 = "": Text26 = "": MaskEdBox1.Text = MsgText(29): MaskEdBox2.Text = MsgText(29)
         AdodcRefresh
         SumShow
         MsgBox MsgText(33), , MsgText(5)
      Else
         FormShow
         AdodcRefresh
         SumShow
         RecordShow
      End If
   Else
      MsgBox "請先輸入暫收款單號！", vbExclamation
   End If
   'end 2023/8/1
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
   'Modified by Morgan 2023/8/1
   'If adoacc0t0.RecordCount <> 0 Then
   '   adoacc0t0.MoveFirst
   'End If
   'adoacc0t0.Find "a0t01 = '" & strItemNo & "'", 0, adSearchForward, 1
   'If adoacc0t0.EOF = False Then
   '   FormShow
   '   AdodcRefresh
   '   SumShow
   '   RecordShow
   'End If
   Text1 = strItemNo
   Command3.Value = True
   'end 2023/8/1
   strItemNo = MsgText(601)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   'Added by Morgan 2023/10/12
   'Debug.Print Now & ": Form_KeyUp :" & Me.ActiveControl.Name
   If m_ActiveControlName <> "" And m_ActiveControlName <> Me.ActiveControl.Name Then
      Exit Sub
   End If
   'end 2023/10/12
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   'Modify by Amy 2023/10/06 原W:9045
   Me.Width = 9060
   Me.Height = 6015 'Modify by Amy 原:5700
   'Modify by Amy 2023/10/06 原(lngWidth - Me.Width) 讓切畫面時不需再調
   Me.Move 0, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   strItemNo = MsgText(601)
   Combo2.AddItem ComboItem(11)
   Combo2.AddItem ComboItem(12)
   Combo2.AddItem ComboItem(13)
         
   'Modified by Morgan 2023/7/27 取消上下筆功能
   'OpenTable 'Modified by Morgan 2023/9/19
   'If adoacc0t0.RecordCount <> 0 Then
   '   adoacc0t0.MoveLast
   '   adoacc0t0.MoveFirst
   '   RecordShow
   'End If
   OpenTable 'Added by Morgan 2023/9/19 新增會用
   m_ToolBarNobrowse = True
   tool1_enabled
   'end 2023/7/27
   
   ObjectEnabled_2
   Option2.Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   m_ToolBarNobrowse = False 'Added by Morgan 2023/7/27
   Set Frmacc11a0 = Nothing
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label6 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If strSaveConfirm <> MsgText(3) Then
      Exit Sub
   End If
   If Mid(MaskEdBox1.Text, 1, 3) <> Mid(CFDate(ACDate(ServerDate)), 1, 3) Then
      Text1 = UpdateNo("acc0t0", "a0t01", 5, MaskEdBox1.Text, MsgText(806))
   Else
      'Text1 = AutoNo(MsgText(806), 5)
      Text1 = strDocNo
   End If
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
'   If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
'      Exit Sub
'   End If
   If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
      MsgBox Label7 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox2.SetFocus
      Exit Sub
   End If
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
Dim strSql As String  'add by sonia 2022/7/28

On Error GoTo Checking
   Screen.MousePointer = vbHourglass
   If adoacc0t0.State = adStateOpen Then
      adoacc0t0.Close
   End If
   adoacc0t0.CursorLocation = adUseClient
   If Option1.Value Then
      '2007/10/30 modify by sonia 因J09400660轉國外收款沖到,故A1P02再加'F'
      'modify by sonia 2021/8/23 很多暫收款財務處自行以總帳傳票沖銷,故人工上A0T10以區別,故加入A0T10為判斷條件.又因速度慢加入ROWNUM<5
      'modify by sonia 2022/7/28
      'adoacc0t0.Open "select * from acc0t0 where a0t10 is null and a0t01 not in (select a1p23 from acc1p0 where a1p02 in ('A', 'Z', 'W', 'F')" & _
         " and a1p05 = '2401' and a1p07 <> 0 and a1p23 is not null) and a0t01 not in (select a0s02 from acc0s0" & _
         " where substr(a0s02, 1, 1) = 'J') and ROWNUM<5 order by a0t01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If Text1 <> "" Then
         strSql = "and a0t01='" & Text1 & "'"
      Else
         strSql = ""
      End If
      'modify by sonia 2024/5/14 a1p02加入'E'，因為a1p04='聯米企業股份有限公司971'有沖銷J09700301
      adoacc0t0.Open "select * from acc0t0 where a0t10 is null and a0t01 not in (select a1p23 from acc1p0 where a1p02 in ('A', 'Z', 'W', 'F', 'E')" & _
         " and a1p05 = '2401' and a1p07 <> 0 and a1p23 is not null) and a0t01 not in (select a0s02 from acc0s0" & _
         " where substr(a0s02, 1, 1) = 'J') " & strSql & " and ROWNUM<5 order by a0t01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      'end 2022/7/28
   Else
      'Modified by Morgan 2023/8/1 取消上下筆功能
      'adoacc0t0.Open "select * from acc0t0 where  order by a0t01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If Text1 <> "" Then
         strSql = "and a0t01='" & Text1 & "'"
      Else
         strSql = " and rownum<1"
      End If
      adoacc0t0.Open "select * from acc0t0 where 1=1 " & strSql & " order by a0t01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      'end 2023/8/1
   End If

   If adoacc1p0.State = adStateOpen Then
      adoacc1p0.Close
   End If
   adoacc1p0.CursorLocation = adUseClient
   'Modify By Sindy 2013/12/30
   'adoacc1p0.Open "select * from acc1p0 where a1p01 = '1' and a1p02 = 'D' and a1p04 = '" & Text1 & "' order by a1p05 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc1p0.Open "select * from acc1p0 where a1p01 = '" & Text24 & "' and a1p02 = 'D' and a1p04 = '" & Text1 & "' order by a1p05 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '2013/12/30 END

'Removed by Morgan 2023/8/1
'   If adoadodc1.State = adStateOpen Then
'      adoadodc1.Close
'   End If
'   adoadodc1.CursorLocation = adUseClient
'   'Modify By Sindy 2013/12/30
'   'adoadodc1.Open "select * from acc1p0 where a1p01 = '1' and a1p02 = 'D' and a1p04 = '" & Text1 & "' order by a1p05 asc", adoTaie, adOpenStatic, adLockReadOnly
'   adoadodc1.Open "select * from acc1p0 where a1p01 = '" & Text24 & "' and a1p02 = 'D' and a1p04 = '" & Text1 & "' order by a1p05 asc", adoTaie, adOpenStatic, adLockReadOnly
'   '2013/12/30 END
'   Set Adodc1.Recordset = adoadodc1
'end 2023/8/1

   Screen.MousePointer = vbDefault
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表(國內暫收款資料(主檔))
'
'*************************************************
Public Sub FormShow()
   Text1 = adoacc0t0.Fields("a0t01").Value
   If IsNull(adoacc0t0.Fields("a0t02").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = adoacc0t0.Fields("a0t02").Value
   End If
   If IsNull(adoacc0t0.Fields("a0t05").Value) Then
      Text3 = MsgText(601)
      Text18 = MsgText(601)
   Else
      Text3 = adoacc0t0.Fields("a0t05").Value
      Text18 = StaffQuery(Text3)
   End If
   If IsNull(adoacc0t0.Fields("a0t06").Value) Then
      Text4 = MsgText(601)
      Text13 = MsgText(601)
   Else
      Text4 = adoacc0t0.Fields("a0t06").Value
      Text13 = CustomerQuery(Text4, 1)
   End If
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc0t0.Fields("a0t03").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(adoacc0t0.Fields("a0t03").Value)
   End If
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(adoacc0t0.Fields("a0t04").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(adoacc0t0.Fields("a0t04").Value)
   End If
   MaskEdBox2.Mask = DFormat
   If IsNull(adoacc0t0.Fields("a0t07").Value) Then
      Text11 = MsgText(601)
   Else
      Text11 = adoacc0t0.Fields("a0t07").Value
   End If
   If IsNull(adoacc0t0.Fields("a0t17").Value) Then
      Text12 = MsgText(601)
   Else
      Text12 = adoacc0t0.Fields("a0t17").Value
   End If
   If IsNull(adoacc0t0.Fields("a0t08").Value) Then
      Text17 = MsgText(601)
   Else
      Text17 = adoacc0t0.Fields("a0t08").Value
   End If
   'Add By Sindy 2013/12/30
   If IsNull(adoacc0t0.Fields("a0t18").Value) Then
      Text24 = "1"
   Else
      Text24 = adoacc0t0.Fields("a0t18").Value
   End If
   '2013/12/30 END
   'add by sonia 2024/5/22
   If IsNull(adoacc0t0.Fields("a0t10").Value) Then
      Text26 = MsgText(601)
   Else
      Text26 = adoacc0t0.Fields("a0t10").Value
   End If
   'end 2024/5/22
End Sub

'*************************************************
'  顯示資料表(國內暫收款資料(分錄檔))
'
'*************************************************
Public Sub AdodcShow()
   Text5 = Adodc1.Recordset.Fields("a1p05").Value
   If IsNull(Adodc1.Recordset.Fields("a1p07").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = Adodc1.Recordset.Fields("a1p07").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p08").Value) Then
      Text15 = MsgText(601)
   Else
      Text15 = Adodc1.Recordset.Fields("a1p08").Value
   End If
   '票據號碼
   If IsNull(Adodc1.Recordset.Fields("a1p09").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = Adodc1.Recordset.Fields("a1p09").Value
   End If
   Text7.Tag = Text7 'Add by Amy 2020/07/03
   '收票銀行
   If IsNull(Adodc1.Recordset.Fields("a1p11").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = Adodc1.Recordset.Fields("a1p11").Value
   End If
   Text8.Tag = Text8 'Add by Amy 2020/07/03
   MaskEdBox3.Mask = MsgText(601)
   If IsNull(Adodc1.Recordset.Fields("a1p12").Value) Then
      MaskEdBox3.Text = MsgText(601)
   Else
      MaskEdBox3.Text = CFDate(Adodc1.Recordset.Fields("a1p12").Value)
   End If
   MaskEdBox3.Mask = DFormat
   '收票帳號
   If IsNull(Adodc1.Recordset.Fields("a1p10").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = Adodc1.Recordset.Fields("a1p10").Value
   End If
   Text9.Tag = Text9 'Add by Amy 2020/07/03
   If IsNull(Adodc1.Recordset.Fields("a1p13").Value) Then
      Combo2 = MsgText(601)
   Else
      Combo2 = Combo2.List(Val(Adodc1.Recordset.Fields("a1p13").Value) - 1)
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p14").Value) Then
      Combo1 = MsgText(601)
   Else
      Combo1 = Adodc1.Recordset.Fields("a1p14").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p06").Value) Then
      Text19 = MsgText(601)
   Else
      Text19 = Adodc1.Recordset.Fields("a1p06").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p17").Value) Then
      Text22 = MsgText(601)
   Else
      Text22 = Adodc1.Recordset.Fields("a1p17").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p15").Value) Then
      Text21 = MsgText(601)
   Else
      Text21 = Adodc1.Recordset.Fields("a1p15").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p30").Value) Then
      Text23 = MsgText(601)
   Else
      Text23 = Adodc1.Recordset.Fields("a1p30").Value
   End If
End Sub

'*************************************************
'  清除 Adodc 之顯示資料
'
'*************************************************
Public Sub AdodcClear()
   Text5 = ""
   Text6 = ""
   Text15 = ""
   Text7 = ""
   Text7.Tag = "" 'Add by Amy 2020/07/03
   Text8 = ""
   Text8.Tag = "" 'Add by Amy 2020/07/03
   MaskEdBox3.Mask = ""
   MaskEdBox3.Text = ""
   MaskEdBox3.Mask = DFormat
   Text9 = ""
   Text9.Tag = "" 'Add by Amy 2020/07/03
   Text20 = ""
   Text10 = ""
   Text16 = ""
   Combo2 = ""
   Combo1 = ""
   Text19 = ""
   Text22 = ""
   Text21 = ""
   Text23 = ""
   'Added by Lydia 2024/11/28
   m_UpdStatus = ""
   m_UpdMsg = ""
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   
   If adoadodc1.State = adStateOpen Then 'Added by Morgan 2023/8/1
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   'Modify By Sindy 2013/12/30
'   adoadodc1.Open "select * from acc1p0, acc010, acc0g0 where a1p05 = a0101 and a1p10 = a0g01 (+) and a1p01 = '1' and a1p02 = 'D'" & _
'      " and a1p04 = '" & Text1 & "' order by a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc1.Open "select * from acc1p0, acc010, acc0g0 where a1p05 = a0101 and a1p10 = a0g01 (+) and a1p01 = '" & Text24 & "' and a1p02 = 'D'" & _
      " and a1p04 = '" & Text1 & "' order by a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2013/12/30 END
   
   Set Adodc1.Recordset = adoadodc1 'Added by Morgan 2023/8/1
   Adodc1.Recordset.Requery
   
   
   SetData ("Refresh") 'Add by Amy 2014/10/29
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  儲存資料表(國內暫收款資料(分錄檔))
'
'*************************************************
Private Sub Acc1p0Save()
Dim Cancel As Boolean
   
On Error GoTo Checking
   'Add By Sindy 2013/12/30
   If Text24 = MsgText(601) Then
      MsgBox MsgText(10) & Label27, , MsgText(5)
      strControlButton = MsgText(602)
      Text24.SetFocus
      Exit Sub
   'Add by Amy 2020/04/07
    Else
        Call Text24_Validate(Cancel)
        If Cancel = True Then
            strControlButton = MsgText(602)
            Text24.SetFocus
        End If
      End If
      'end 2020/04/07
   '2013/12/30 END
   'Add by Amy 2014/11/04
   If Text2 = MsgText(601) Then
      MsgBox MsgText(10) & Label2, , MsgText(5)
      strControlButton = MsgText(602)
      Text2.SetFocus
      Exit Sub
   End If
   'end 2014/11/04
   If Text5 = MsgText(601) Then
      MsgBox MsgText(10) & Label8, , MsgText(5)
      strControlButton = MsgText(602)
      Text5.SetFocus
      Exit Sub
   Else
      'Modify By Sindy 2013/12/30
'      If ExistCheck("acc010", "a0101", Text5, Label8) = False Then
'         strControlButton = MsgText(602)
'         Text5.SetFocus
'         Exit Sub
'      End If
      Cancel = False
      Call Text5_Validate(Cancel)
      If Cancel = True Then
         strControlButton = MsgText(602)
         Text5.SetFocus
         Exit Sub
      End If
      '2013/12/30 END
      If Text9 <> MsgText(601) Then
         If ExistCheck("acc0g0", "a0g01", Text9, Label13) = False Then
            strControlButton = MsgText(602)
            Text9.SetFocus
            Exit Sub
         End If
      End If
      If MaskEdBox3.Text <> MsgText(601) And MaskEdBox3.Text <> MsgText(29) Then
         If DateCheck(MaskEdBox3.Text) = MsgText(603) Then
            MsgBox Label12 & MsgText(63), , MsgText(5)
            strControlButton = MsgText(602)
            MaskEdBox3.SetFocus
            Exit Sub
         End If
      End If
   End If
   
   If Text22 <> MsgText(601) Then
      Text22 = CaseNoZero(Text22)
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select pa01 as SystemNo from patent where pa01 = '" & Mid(Text22, 1, Len(Text22) - 9) & "' and pa02 = '" & Mid(Text22, Len(Text22) - 8, 6) & "' and pa03 = '" & Mid(Text22, Len(Text22) - 2, 1) & "' and pa04 = '" & Mid(Text22, Len(Text22) - 1, 2) & "' union " & _
                     "select tm01 as SystemNo from trademark where tm01 = '" & Mid(Text22, 1, Len(Text22) - 9) & "' and tm02 = '" & Mid(Text22, Len(Text22) - 8, 6) & "' and tm03 = '" & Mid(Text22, Len(Text22) - 2, 1) & "' and tm04 = '" & Mid(Text22, Len(Text22) - 1, 2) & "' union " & _
                     "select lc01 as SystemNo from lawcase where lc01 = '" & Mid(Text22, 1, Len(Text22) - 9) & "' and lc02 = '" & Mid(Text22, Len(Text22) - 8, 6) & "' and lc03 = '" & Mid(Text22, Len(Text22) - 2, 1) & "' and lc04 = '" & Mid(Text22, Len(Text22) - 1, 2) & "' union " & _
                     "select hc01 as SystemNo from hirecase where hc01 = '" & Mid(Text22, 1, Len(Text22) - 9) & "' and hc02 = '" & Mid(Text22, Len(Text22) - 8, 6) & "' and hc03 = '" & Mid(Text22, Len(Text22) - 2, 1) & "' and hc04 = '" & Mid(Text22, Len(Text22) - 1, 2) & "' union " & _
                     "select sp01 as SystemNo from servicepractice where sp01 = '" & Mid(Text22, 1, Len(Text22) - 9) & "' and sp02 = '" & Mid(Text22, Len(Text22) - 8, 6) & "' and sp03 = '" & Mid(Text22, Len(Text22) - 2, 1) & "' and sp04 = '" & Mid(Text22, Len(Text22) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount = 0 Then
         MsgBox MsgText(28) & Label24, , MsgText(5)
         strControlButton = MsgText(602)
         adoquery.Close
         Exit Sub
      End If
      adoquery.Close
   End If
   If Text21 <> MsgText(601) Then
      If Len(Text21) = 6 Then
         Text21 = AfterZero(Text21)
      'Add by Morgan 2007/3/1 八碼時要補'0'
      ElseIf Len(Text21) = 8 Then
         Text21 = Text21 & "0"
      'End 2007/3/1
      End If
      If ExistCheck("customer", "cu01", Mid(Text21, 1, 8), Label23, False) = False Then
         If ExistCheck("acc0i0", "a0i01", Text21, Label23, False) = False Then
            If ExistCheck("staff", "st01", Text21, Label23, False) = False Then
               MsgBox MsgText(28) & Label23, , MsgText(5)
               strControlButton = MsgText(602)
               Exit Sub
            End If
         End If
      End If
   End If
   
   'add by sonia 2024/5/23 因為有外幣暫收款情形，有可能有借方或貸方的匯差手續費,故改為只檢查貸方2401科目金額必須與上方暫收款金額Text17相同即可
   If Text5 = "2401" And Val(Text15) <> Val(Text17) And Val(Text15) > 0 Then
      MsgBox "貸方暫收款科目金額必須與上方暫收款金額相同！", , MsgText(5)
      strControlButton = MsgText(602)
      Text15.SetFocus
      Exit Sub
   End If
   'end 2024/5/23
   
   'add by sonia 2015/12/30 檢查民國105年起法務收入科目不可使用
   intI = PUB_AccNoEnable(Text5, Val(FCDate(MaskEdBox1.Text)))
   If intI <> 0 Then
      strControlButton = MsgText(602)
      Text5.SetFocus
      Exit Sub
   End If
   'end 2015/12/30
   'Add by Morgan 2007/10/2 檢查科目部門&智權人員是否正確
   intI = PUB_AccNoGood(Text5, Text19, Text3)
   If intI <> 0 Then
      strControlButton = MsgText(602)
      If intI = 1 Then
         Text5.SetFocus
      ElseIf intI = 2 Then
         Text19.SetFocus
      ElseIf intI = 3 Then
         Text3.SetFocus
      End If
      Exit Sub
   End If
   'end 2007/10/2
   
   If Adodc1.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc1.Recordset.Fields("a1p22").Value) = False Then
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select ax210 from acc021 where ax201 = '" & Adodc1.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc1.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(155), , MsgText(5)
            strControlButton = MsgText(602)
            Text5.SetFocus
            adoquery.Close
            Exit Sub
         End If
         adoquery.Close
      End If
   End If
   
   If adoacc1p0.State = adStateOpen Then 'Added by Morgan 2023/8/1
      adoacc1p0.Close
   End If
   adoacc1p0.CursorLocation = adUseClient
   'Modify By Sindy 2013/12/30
   'adoacc1p0.Open "select * from acc1p0 where a1p01 = '1' and a1p02 = 'D' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc1p0.Open "select * from acc1p0 where a1p01 = '" & Text24 & "' and a1p02 = 'D' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '2013/12/30 END
   If adoacc1p0.RecordCount = 0 Then
      adoacc1p0.AddNew
      'Modify By Sindy 2013/12/30
      'adoacc1p0.Fields("a1p01").Value = "1"
      'adoacc1p0.Fields("a1p03").Value = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'D' and a1p04 = '" & Text1 & "'", 3)
      adoacc1p0.Fields("a1p01").Value = Text24
      adoacc1p0.Fields("a1p03").Value = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '" & Text24 & "' and a1p02 = 'D' and a1p04 = '" & Text1 & "'", 3)
      '2013/12/30 END
      adoacc1p0.Fields("a1p02").Value = "D"
      adoacc1p0.Fields("a1p04").Value = Text1
   End If
   adoacc1p0.Fields("a1p05").Value = Text5
   'adoacc1p0.Fields("a1p06").Value = StaffDeptQuery(Text3)
   adoacc1p0.Fields("a1p06").Value = MsgText(55)
   If Text6 <> MsgText(601) Then
      adoacc1p0.Fields("a1p07").Value = Val(Text6)
   Else
      adoacc1p0.Fields("a1p07").Value = 0
   End If
   If Text15 <> MsgText(601) Then
      adoacc1p0.Fields("a1p08").Value = Val(Text15)
   Else
      adoacc1p0.Fields("a1p08").Value = 0
   End If
   If Text7 <> MsgText(601) Then
      adoacc1p0.Fields("a1p09").Value = Text7
   Else
      adoacc1p0.Fields("a1p09").Value = Null
   End If
   If Text8 <> MsgText(601) Then
      adoacc1p0.Fields("a1p11").Value = Text8
   Else
      adoacc1p0.Fields("a1p11").Value = Null
   End If
   If MaskEdBox3.Text <> MsgText(601) And MaskEdBox3.Text <> MsgText(29) Then
      adoacc1p0.Fields("a1p12").Value = Val(FCDate(MaskEdBox3.Text))
   Else
      adoacc1p0.Fields("a1p12").Value = Null
   End If
   If Text9 <> MsgText(601) Then
      adoacc1p0.Fields("a1p10").Value = Text9
   Else
      adoacc1p0.Fields("a1p10").Value = Null
   End If
   If Combo2 <> MsgText(601) Then
      adoacc1p0.Fields("a1p13").Value = Mid(Combo2, 1, 1)
   Else
      adoacc1p0.Fields("a1p13").Value = Null
   End If
   If Combo1 <> MsgText(601) Then
      adoacc1p0.Fields("a1p14").Value = Combo1
      Combo1.AddItem Combo1
   Else
      adoacc1p0.Fields("a1p14").Value = Null
   End If
   'modify by sonia 2021/1/29 加傳本所案號以判別FCP,FCT英日文組
   'If AccNoToSalesNo(Text5) = "" Then
   If AccNoToSalesNo(Text5, Text22) = "" Then
      If Text3 <> MsgText(601) Then
         adoacc1p0.Fields("a1p16").Value = Text3
      Else
         adoacc1p0.Fields("a1p16").Value = Null
      End If
   Else
      'modify by sonia 2021/1/29 加傳本所案號以判別FCP,FCT英日文組
      'adoacc1p0.Fields("a1p16").Value = AccNoToSalesNo(Text5)
      adoacc1p0.Fields("a1p16").Value = AccNoToSalesNo(Text5, Text22)
   End If
   If Text4 <> MsgText(601) Then
      adoacc1p0.Fields("a1p15").Value = Text4
   Else
      adoacc1p0.Fields("a1p15").Value = Null
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      adoacc1p0.Fields("a1p18").Value = Val(FCDate(MaskEdBox1.Text))
   Else
      adoacc1p0.Fields("a1p18").Value = Null
   End If
   If Text19 <> MsgText(601) Then
      adoacc1p0.Fields("a1p06").Value = Text19
   Else
      adoacc1p0.Fields("a1p06").Value = MsgText(55)
   End If
   If Text22 <> MsgText(601) Then
      adoacc1p0.Fields("a1p17").Value = Text22
   Else
      adoacc1p0.Fields("a1p17").Value = Null
   End If
   If Text23 <> MsgText(601) Then
      adoacc1p0.Fields("a1p30").Value = Text23
   Else
      adoacc1p0.Fields("a1p30").Value = Null
   End If
   If IsNull(adoacc1p0.Fields("a1p22").Value) = False Then
      adoacc1p0.Fields("a1p27").Value = MsgText(602)
   End If
   
   'Add by Morgan 2005/10/19 修改時要上更新時間
   If strSaveConfirm = MsgText(4) Then
      adoacc1p0.Fields("a1p28") = strSrvDate(2)
      adoacc1p0.Fields("a1p29") = ServerTime
   End If
   
   adoacc1p0.UpdateBatch
   
   If Text7 <> MsgText(601) And Text5 = "113001" Then
      'Add by Amy 2020/07/03 若修改票據號碼按 Insert,原資料不會被刪除
      If Text7 & Text8 & Text9 <> Text7.Tag & Text8.Tag & Text9.Tag Then
        adoTaie.Execute "delete from acc0e0 where a0e01 = '" & Text9.Tag & "' and a0e02 = '" & Text7.Tag & "' And a0e07='" & Text8.Tag & "' "
      End If
      'Modify 2020/07/03 +a0e07 因改為key
      adoTaie.Execute "delete from acc0e0 where a0e01 = '" & Text9 & "' and a0e02 = '" & Text7 & "' And a0e07='" & Text8 & "' "
      adoTaie.Execute "insert into acc0e0 values ('" & Text7 & "', '" & Text9 & "', '" & Text1 & "', 'R', '1', '" & Text4 & "', '" & Text8 & "', '" & Mid(Combo2, 1, 1) & "', " & _
                      "" & Val(FCDate(MaskEdBox3.Text)) & ", " & Val(Text6) & ", '" & Combo1 & "', " & Val(FCDate(MaskEdBox1.Text)) & ", 0, 0, 0, null, 0, null, null, null, 0, 0, 0, 0, '" & Text24 & "', null, null, null, null, 0, null, 0, null, null, '" & strUserNum & "', " & Val(strSrvDate(2)) & ", " & ServerTime & ", null, 0, 0, 0, 0, 0, null, 0, null, null)"
   End If
   'Remove by Morgan 2005/10/31 存檔時更新即可
   'adoTaie.Execute "update acc0t0 set a0t08 = " & Val(Replace(Text16, ",", "")) & " where a0t01 = '" & Text1 & "'"
   strSerialNo = MsgText(601)
   AdodcRefresh
   SumShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   Err.Clear
End Sub

'*************************************************
'  刪除 Adodc 之資料
'
'*************************************************
Private Sub AdodcDelete()
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount <> 0 Then
      'Modify By Sindy 2013/12/30
      'adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'D' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text1 & "'"
      adoTaie.Execute "delete from acc1p0 where a1p01 = '" & Text24 & "' and a1p02 = 'D' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text1 & "'"
      '2013/12/30 END
      SumShow
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
'  計算並顯示合計
'
'*************************************************
Public Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   'Modify By Sindy 2013/12/30
   'adoaccsum.Open "select sum(a1p07), sum(a1p08), count(*) from acc1p0 where a1p01 = '1' and a1p02 = 'D' and a1p04 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   adoaccsum.Open "select sum(a1p07), sum(a1p08), count(*) from acc1p0 where a1p01 = '" & Text24 & "' and a1p02 = 'D' and a1p04 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   '2013/12/30 END
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text10 = MsgText(601)
      Else
         Text10 = Format(adoaccsum.Fields(0).Value, FAmount)
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text16 = MsgText(601)
      Else
         Text16 = Format(adoaccsum.Fields(1).Value, FAmount)
      End If
      If IsNull(adoaccsum.Fields(2).Value) Then
         Text20 = MsgText(601)
      Else
         Text20 = Format(adoaccsum.Fields(2).Value, DDollar)
      End If
   Else
      Text10 = MsgText(601)
      Text16 = MsgText(601)
      Text20 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyInsert
         'Add by Sindy 2021/12/14 檢查畫面上的物件是否含有Unicode文字
         If PUB_ChkUniText(Me, True, True) = False Then
            Exit Sub
         End If

         If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
            Exit Sub
         End If
         If strControlButton <> MsgText(602) Then
            Acc1p0Save
         End If
         If strControlButton <> MsgText(602) Then
            AdodcClear
            Text5.SetFocus
         End If
         SumShow
         strControlButton = MsgText(601)
   End Select
   KeyEnter KeyCode
End Sub

Private Sub MaskEdBox3_Validate(Cancel As Boolean)
   If MaskEdBox3.Text = MsgText(601) Or MaskEdBox3.Text = MsgText(29) Then
      Exit Sub
   End If
   If DateCheck(MaskEdBox3.Text) = MsgText(603) Then
      MsgBox Label12 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox3.SetFocus
      Exit Sub
   End If
   RemarkShow
End Sub

Private Sub Option1_Click()
   'Removed by Morgan 2023/7/27 取消上下筆功能
   'OpenTable
   'If adoacc0t0.RecordCount <> 0 Then
   '    adoacc0t0.MoveFirst
   '    RecordShow
   'End If
   'end 2023/7/27
End Sub

Private Sub Option2_Click()
   'Removed by Morgan 2023/7/27 取消上下筆功能
   'OpenTable
   'If adoacc0t0.RecordCount <> 0 Then
   '    adoacc0t0.MoveFirst
   '    RecordShow
   'End If
   'end 2023/7/27
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
   'edit by nickc 2007/06/11  切換輸入法改用API
   OpenIme
End Sub

Private Sub Text12_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
CloseIme
End Sub

Private Sub Text15_GotFocus()
   TextInverse Text15
End Sub

'add by sonia 2025/5/9 貸方2401才預設
Private Sub Text15_Validate(Cancel As Boolean)
   If Text5 = "2401" And Val(Text15) > 0 Then
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select sn01 from salesno where sn02 = '" & Text3 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         If IsNull(adoquery.Fields(0).Value) = False And InStr(Combo1, adoquery.Fields(0)) = 0 Then
            Combo1 = Combo1 & adoquery.Fields(0).Value
         End If
      End If
      adoquery.Close
      
      If Trim(Text12.Text) <> "" Then
         If InStr(Combo1, Trim(Text12.Text)) = 0 Then
            Combo1 = Combo1 & "/" & Trim(Text12.Text) & "/" & Text1
         End If
      Else
         If InStr(Combo1, Text13 & "/" & Text1) = 0 Then
            Combo1 = Combo1 & "/" & Text13 & "/" & Text1
         End If
      End If
      '對沖其他放暫收款單號
      Text23 = Text1
   End If
End Sub
'end 2025/5/9

Private Sub Text17_GotFocus()
   TextInverse Text17
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

Private Sub Text19_Validate(Cancel As Boolean)
   If Text19 <> MsgText(601) Then
      If ExistCheck("acc090", "a0901", Text19, Label21) = False Then
         Cancel = True
         Exit Sub
      End If
   End If
   If CheckDept(Text5, Text19) = False Then
      MsgBox MsgText(103), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text2_Change()
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   '1.暫收款 2.溢收轉入 3.銷帳退費轉入
   Select Case Text2
      Case Mid(ComboItem(91), 1, 1)
         '1.暫收款
         ObjectEnabled_1
      Case Mid(ComboItem(92), 1, 1)
         '2.溢收轉入
         ObjectEnabled_2
      Case Mid(ComboItem(93), 1, 1)
         '3.銷帳退費轉入
         ObjectEnabled_2
   End Select
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text21_GotFocus()
   TextInverse Text21
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text21_Validate(Cancel As Boolean)
   If Text21 <> MsgText(601) Then
      If Len(Text21) = 6 Then
         Text21 = AfterZero(Text21)
      'Add by Morgan 2007/3/1 八碼時要補'0'
      ElseIf Len(Text21) = 8 Then
         Text21 = Text21 & "0"
      'End 2007/3/1
      End If
      If ExistCheck("customer", "cu01", Mid(Text21, 1, 8), Label23, False) = False Then
         If ExistCheck("acc0i0", "a0i01", Text21, Label23, False) = False Then
            If ExistCheck("staff", "st01", Text21, Label23, False) = False Then
               MsgBox MsgText(28) & Label23, , MsgText(5)
               Cancel = True
               Exit Sub
            End If
         End If
      End If
   End If
End Sub

Private Sub Text22_GotFocus()
   TextInverse Text22
End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text22_Validate(Cancel As Boolean)
On Error GoTo Checking
   If Text22 <> MsgText(601) Then
      Text22 = CaseNoZero(Text22)
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select pa01 as SystemNo from patent where pa01 = '" & Mid(Text22, 1, Len(Text22) - 9) & "' and pa02 = '" & Mid(Text22, Len(Text22) - 8, 6) & "' and pa03 = '" & Mid(Text22, Len(Text22) - 2, 1) & "' and pa04 = '" & Mid(Text22, Len(Text22) - 1, 2) & "' union " & _
                   "select tm01 as SystemNo from trademark where tm01 = '" & Mid(Text22, 1, Len(Text22) - 9) & "' and tm02 = '" & Mid(Text22, Len(Text22) - 8, 6) & "' and tm03 = '" & Mid(Text22, Len(Text22) - 2, 1) & "' and tm04 = '" & Mid(Text22, Len(Text22) - 1, 2) & "' union " & _
                   "select lc01 as SystemNo from lawcase where lc01 = '" & Mid(Text22, 1, Len(Text22) - 9) & "' and lc02 = '" & Mid(Text22, Len(Text22) - 8, 6) & "' and lc03 = '" & Mid(Text22, Len(Text22) - 2, 1) & "' and lc04 = '" & Mid(Text22, Len(Text22) - 1, 2) & "' union " & _
                   "select hc01 as SystemNo from hirecase where hc01 = '" & Mid(Text22, 1, Len(Text22) - 9) & "' and hc02 = '" & Mid(Text22, Len(Text22) - 8, 6) & "' and hc03 = '" & Mid(Text22, Len(Text22) - 2, 1) & "' and hc04 = '" & Mid(Text22, Len(Text22) - 1, 2) & "' union " & _
                   "select sp01 as SystemNo from servicepractice where sp01 = '" & Mid(Text22, 1, Len(Text22) - 9) & "' and sp02 = '" & Mid(Text22, Len(Text22) - 8, 6) & "' and sp03 = '" & Mid(Text22, Len(Text22) - 2, 1) & "' and sp04 = '" & Mid(Text22, Len(Text22) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount = 0 Then
         MsgBox MsgText(28) & Label24, , MsgText(5)
         Cancel = True
         adoquery.Close
         Exit Sub
      End If
      adoquery.Close
   End If
   Exit Sub
Checking:
   MsgBox MsgText(128), , MsgText(5)
   Exit Sub
End Sub

Private Sub Text23_GotFocus()
   TextInverse Text23
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2013/12/27
Private Sub Text24_Change()
   If Text24 = MsgText(601) Then
      Text25 = ""
      Exit Sub
   End If
   'Modify by Amy 2020/04/07 改抓function
   If InStr(GetBookKeepCmp, Text24) = 0 Then
     Text25 = ""
     Exit Sub
   End If
   Text25 = A0802Query(Text24)
'   Select Case Text24
'      Case "1"
'         Text25 = MsgText(901)
''      Case "2"
''         Text25 = MsgText(902)
''      Case "3"
''         Text25 = MsgText(903)
''      Case "5"
''         Text25 = MsgText(904)
''      Case "7"
''         Text25 = MsgText(905)
''      Case "8"
''         Text25 = MsgText(906)
''      Case "9"
''         Text25 = MsgText(908)
'      Case "J"
'         Text25 = MsgText(907)
'   End Select
   'end 2020/04/07
End Sub

Private Sub Text24_GotFocus()
   TextInverse Text24
End Sub

Private Sub Text24_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text24_Validate(Cancel As Boolean)
   If Text24 <> MsgText(601) Then
'      MsgBox "請輸入公司別!!", , MsgText(5)
'      Cancel = True
'      Text24.SetFocus
'      Exit Sub
'   Else
      'Modify by Amy 2020/04/07
      'If Text24 <> "1" And Text24 <> "J" Then
      If InStr(GetBookKeepCmp, Text24) = 0 Then
         MsgBox Label27 & MsgText(63), , MsgText(5) '原:"公司別只可輸入 1 或 J"
      'end 2020/04/07
         Cancel = True
         Text24.SetFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If Text3 <> MsgText(601) Then
      'modify by sonia 2025/5/9 員工已離職要提醒
      'If ExistCheck("staff", "st01", Text3, Label4, False) = False Then
      '   MsgBox MsgText(45) & Label4, , MsgText(5)
      If PUB_GetStaffState(Text3.Text, strExc(1), True) = 0 Then
      'end 2025/5/9
         Cancel = True
         Text3.SetFocus
         TextInverse Text3
         Exit Sub
      End If
   End If
   Text18 = StaffQuery(Text3)
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
   If Text4 <> MsgText(601) Then
      If ExistCheck("customer", "cu01", Mid(IIf(Len(Text4) = 6, AfterZero(Text4), Text4), 1, 8), Label5, False) = False Then
         MsgBox MsgText(45) & Label5, , MsgText(5)
         Text4.SetFocus
         Cancel = True
         TextInverse Text4
         Exit Sub
      End If
   Else
      MsgBox MsgText(45) & Label5, , MsgText(5)
      Text4.SetFocus
      Cancel = True
      TextInverse Text4
      Exit Sub
   End If
   If Len(Text4) = 6 Then
      Text4 = AfterZero(Text4)
   ElseIf Len(Text4) = 8 Then
      Text4 = Text4 & "0"
   End If
   Text13 = CustomerQuery(Text4, 1)
   
   'add by sonia 2024/7/5
   If Right(Text4, 1) <> "0" Then
      MsgBox "客戶編號不可輸入更名前的編號，即第9碼必須為0 !", , MsgText(5)
      Text4.SetFocus
      Cancel = True
      TextInverse Text4
      Exit Sub
   End If
   'end 2024/7/5
End Sub

Private Sub Text5_Change()
   Text14 = A0102Query(Text5)
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   If Text5 <> MsgText(601) Then
      If ExistCheck("acc010", "a0101", Text5, Label8) = False Then
         Cancel = True
         Text5.SetFocus
         TextInverse Text5
         Exit Sub
      End If
      
      'Add By Sindy 2013/12/30
      If PUB_CheckCompany(Text5, Text24) = False Then
         Cancel = True
         Text24.SetFocus
         TextInverse Text24
         Exit Sub
      End If
'      adoquery.CursorLocation = adUseClient
'      adoquery.Open "select a0109 from acc010 where a0101 = '" & Text5 & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adoquery.RecordCount <> 0 Then
'         If IsNull(adoquery.Fields(0).Value) = False Then
'            If Text24 <> adoquery.Fields(0).Value Then
'               MsgBox "公司別必須為 " & adoquery.Fields(0).Value
'               adoquery.Close
'               Cancel = True
'               Text24.SetFocus
'               TextInverse Text24
'               Exit Sub
'            End If
'         End If
'      End If
'      adoquery.Close
      '2013/12/30 END
   End If
   
'cancel by sonia 2025/5/9 移到Text15_Validate
'   'modify by sonia 2024/5/28 貸方2401才做
'   'If Text5 = "2401" Then
'   If Text5 = "2401" And Val(Text15) > 0 Then
'      'Modify by Amy 2024/01/03 摘要要可加字-辜
'      'Combo1 = ""
'      adoquery.CursorLocation = adUseClient
'      adoquery.Open "select sn01 from salesno where sn02 = '" & Text3 & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adoquery.RecordCount <> 0 Then
'         If IsNull(adoquery.Fields(0).Value) = False And InStr(Combo1, adoquery.Fields(0)) = 0 Then
'            Combo1 = Combo1 & adoquery.Fields(0).Value
'         End If
'      End If
'      adoquery.Close
'
'      'Modify by Morgan 2004/11/15 若備註有輸入資料時預設為摘要
'      'Combo1 = Combo1 & "/" & Text13 & "/" & Text1
'      If Trim(Text12.Text) <> "" Then
'         If InStr(Combo1, Trim(Text12.Text)) = 0 Then
'            Combo1 = Combo1 & "/" & Trim(Text12.Text)
'         End If
'      Else
'         If InStr(Combo1, Text13 & "/" & Text1) = 0 Then
'            Combo1 = Combo1 & "/" & Text13 & "/" & Text1
'         End If
'      End If
'      '2004/11/15
'      'end 2024/01/03
'
'      '2012/8/29 ADD BY SONIA 對沖其他放暫收款單號
'      Text23 = Text1
'      '2012/8/29 END
'   End If
'end 2025/5/9

   'add by sonia 2021/1/29
   If SalesNoCheckAccNo(Text5, Text3) = False Then
   End If
   'end 2021/1/29
   RemarkShow
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

'2025/5/9 add by sonia 借方不論科目都帶摘要智權人員簡碼+客戶名稱,但不帶暫收款單號
Private Sub Text6_Validate(Cancel As Boolean)
   If Val(Text6) > 0 Then
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select sn01 from salesno where sn02 = '" & Text3 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         If IsNull(adoquery.Fields(0).Value) = False And InStr(Combo1, adoquery.Fields(0)) = 0 Then
            Combo1 = Combo1 & adoquery.Fields(0).Value
         End If
      End If
      adoquery.Close
      
      If Trim(Text12.Text) <> "" Then
         If InStr(Combo1, Trim(Text12.Text)) = 0 Then
            Combo1 = Combo1 & "/" & Trim(Text12.Text)
         End If
      Else
         If InStr(Combo1, Text13 & "/" & Text1) = 0 Then
            Combo1 = Combo1 & "/" & Text13
         End If
      End If
   End If
End Sub
'end 2025/5/9

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
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
'  物件使用狀態設定1
'
'*************************************************
Public Sub ObjectEnabled_1()
   Adodc1.Enabled = True
   Command1.Enabled = True
   Text5.Enabled = True
   Text6.Enabled = True
   Text7.Enabled = True
   Text8.Enabled = True
   MaskEdBox3.Enabled = True
   Text9.Enabled = True
   Combo1.Enabled = True
   Combo2.Enabled = True
   Text19.Enabled = True
   Text22.Enabled = True
   Text21.Enabled = True
   Text23.Enabled = True
   'Added by Lydia 2024/11/28 增加控制欄位
   Text4.Enabled = True
   Text2.Enabled = True
   Text3.Enabled = True
   MaskEdBox2.Enabled = True
   Text17.Enabled = True
   Text12.Enabled = True
   'end 2024/11/28
End Sub

'*************************************************
'  物件使用狀態設定2
'
'*************************************************
Public Sub ObjectEnabled_2()
   Adodc1.Enabled = False
   Command1.Enabled = False
   Text5.Enabled = False
   Text6.Enabled = False
   Text7.Enabled = False
   Text8.Enabled = False
   MaskEdBox3.Enabled = False
   Text9.Enabled = False
   Combo1.Enabled = False
   Combo2.Enabled = False
   Text19.Enabled = False
   Text22.Enabled = False
   Text21.Enabled = False
   Text23.Enabled = False
   'Added by Lydia 2024/11/28 增加控制欄位
   Text4.Enabled = False
   Text2.Enabled = False
   Text3.Enabled = False
   MaskEdBox2.Enabled = False
   Text17.Enabled = False
   Text12.Enabled = False
   'end 2024/11/28
End Sub

'Added by Lydia 2024/11/28
'*************************************************
'  物件使用狀態設定3--僅開放客戶欄位
'
'*************************************************
Public Sub ObjectEnabled_3()
   Adodc1.Enabled = False
   Command1.Enabled = False
   Text5.Enabled = False
   Text6.Enabled = False
   Text7.Enabled = False
   Text8.Enabled = False
   MaskEdBox3.Enabled = False
   Text9.Enabled = False
   Combo1.Enabled = False
   Combo2.Enabled = False
   Text19.Enabled = False
   Text22.Enabled = False
   Text21.Enabled = False
   Text23.Enabled = False
   '*****增加控制欄位****
   Text4.Enabled = True
   Text2.Enabled = False
   Text3.Enabled = False
   MaskEdBox2.Enabled = False
   Text17.Enabled = False
   Text12.Enabled = False
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = adoacc0t0.Bookmark & MsgText(35) & adoacc0t0.RecordCount
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   If Text9 <> MsgText(601) Then
      If ExistCheck("acc0g0", "a0g01", Text9, Label13, False) = False Then
         MsgBox MsgText(45) & Label13, , MsgText(5)
         Cancel = True
         Text9.SetFocus
         Exit Sub
         TextInverse Text9
      End If
   End If
   RemarkShow
End Sub

'*************************************************
'  摘要顯示
'*************************************************
Public Sub RemarkShow()
   If Mid(Text5, 1, 4) = "1130" Then
      Combo1 = FCDate(MaskEdBox3.Text) & "/" & Text7 & "/" & Text8 & "/" & A0g02Query(Text9)
      Exit Sub
   End If
End Sub

'Add by Morgan 2005/10/19
'Modified by Lydia 2024/11/28 +pStatus
Public Function CheckUsed(Optional ByRef pStatus As String) As Boolean

   'Added by Lydia 2024/11/28 debug
   pStatus = ""
   If Text2 = "1" Then  '未收文客戶暫收款管制
      strSql = "select a1p01,a1p15,a1p22,ax210 from acc1p0,acc021 where a1p01='" & Trim(Text24) & "' and a1p04='" & Trim(Text1) & "' " & _
               "and a1p01=ax201(+) and a1p22=ax202(+) and a1p03=ax203(+) and nvl(ax210,0)>0 group by a1p01,a1p15,a1p22,ax210 "
      CheckOC3
      With AdoRecordSet3
         .CursorLocation = adUseClient
         .Open strSql, adoTaie, adOpenForwardOnly, adLockReadOnly
         If .RecordCount > 0 Then
            If .Fields("a1p15") = Trim(Text4) And Trim(Text4) = "X03072010" Then
               '未收文客戶暫收款作業：X03072010用於新客戶未建檔，因沒有客戶代號卻先收款
               pStatus = "1"  '僅開放修改客戶欄位
               'Added by Lydia 2025/02/07 若屬尚無人認領,中,則輸入M0100
               If Trim(Text3) = "" Then
                  Text3 = "M0100"
                  Text18 = StaffQuery(Text3)
               End If
               'end 2025/02/07
            Else
               '傳票已過帳, 不可變更原始資料...
               MsgBox MsgText(155), , MsgText(5)
               CheckUsed = True
            End If
         End If
      End With
   End If
   m_UpdStatus = pStatus & IIf(pStatus <> "", Trim(Text4), "") '特殊修改
   m_UpdMsg = ""
   
   If m_UpdStatus = "" Then
   'end 2024/11/28
      '檢查是否已沖收款
      'Modify by Morgan 2007/4/27 加判斷科目為2401
      'strSQL = "select * from acc1p0 where a1p23='" & Text1 & "' and a1p07>0 and rownum<2"
      'modify by sonia 2021/8/23 很多暫收款財務處自行以總帳傳票沖銷,故人工上A0T10以區別,故加入A0T10為判斷條件
      'strSql = "select * from acc1p0 where a1p23='" & Text1 & "' and a1p07>0 and a1p05='2401' and rownum<2"
      strSql = "select a0t01 from acc0t0 where a0t01='" & Text1 & "'  And a0t10 is not null " & _
                "union select a1p23 from acc1p0 where a1p23='" & Text1 & "' and a1p07>0 and a1p05='2401' and rownum<2"
      'end 2021/8/18
      CheckOC3
      With AdoRecordSet3
         .CursorLocation = adUseClient
         .Open strSql, adoTaie, adOpenForwardOnly, adLockReadOnly
         If .RecordCount > 0 Then
            MsgBox "本暫收款已沖不可修改或刪除！", vbExclamation
            CheckUsed = True
         End If
      End With
   End If 'Added by Lydia 2024/11/28
End Function

Public Sub Frmacc11a0_Save()
Dim adocheck As New ADODB.Recordset
Dim strYes As String
Dim strMsg As String 'Add by Amy 2014/10/29
Dim bolCancel As Boolean 'Add by Amy 2020/04/07
   
   On Error GoTo Checking
   With Frmacc11a0
      'Add By Sindy 2013/12/30
      If .Text24 = MsgText(601) Then
         MsgBox MsgText(10) & .Label27, , MsgText(5)
         strControlButton = MsgText(602)
         .Text24.SetFocus
         Exit Sub
      'Add by Amy 2020/04/07
      Else
        Call Text24_Validate(bolCancel)
        If bolCancel = True Then
            strControlButton = MsgText(602)
            .Text24.SetFocus
        End If
      End If
      'end 2020/04/07
      '2013/12/30 END
      'Add by Amy 2014/11/04
      If Text2 = MsgText(601) Then
          MsgBox MsgText(10) & Label2, , MsgText(5)
          strControlButton = MsgText(602)
          Text2.SetFocus
          Exit Sub
      End If
      'end 2014/11/04
      If .Text1 = MsgText(601) Then
         MsgBox MsgText(10) & .Label1, , MsgText(5)
         strControlButton = MsgText(602)
         .Text1.SetFocus
         Exit Sub
      Else
         If .Text3 <> MsgText(601) Then
            If ExistCheck("staff", "st01", .Text3, .Label4) = False Then
               strControlButton = MsgText(602)
               .Text3.SetFocus
               Exit Sub
            End If
         'Added by Lydia 2025/02/07
         Else
            '未收文客戶暫收款作業：X03072010用於新客戶未建檔，因沒有客戶代號卻先收款;若屬尚無人認領,中,則輸入M0100
            If Left(m_UpdStatus, 1) = "1" And Trim(Text4) = "X03072010" Then
               Text3 = "M0100"
               Text18 = StaffQuery(Text3)
            End If
         'end 2025/02/07
         End If
         
         'cancel by sonia 2024/5/23 因為有外幣暫收款情形，有可能有借方或貸方的匯差手續費,故改為只檢查貸方2401科目金額必須與上方暫收款金額Text17相同即可
         'If Val(.Text17) <> Val(.Text16) And Val(.Text16) <> 0 Then
         '   MsgBox MsgText(59), , MsgText(5)
         '   strControlButton = MsgText(602)
         '   .Text17.SetFocus
         '   Exit Sub
         'End If
         If Val(.Text17) = 0 Then
            MsgBox MsgText(59), , MsgText(5)
            strControlButton = MsgText(602)
            .Text17.SetFocus
            Exit Sub
         End If
         If .Text4 <> MsgText(601) Then
            If ExistCheck("customer", "cu01", Mid(IIf(Len(.Text4) = 6, AfterZero(.Text4), .Text4), 1, 8), .Label5) = False Then
               strControlButton = MsgText(602)
               .Text4.SetFocus
               Exit Sub
            End If
         Else
            MsgBox MsgText(45) & .Label5, , MsgText(5)
            strControlButton = MsgText(602)
            .Text4.SetFocus
            Exit Sub
         End If
         'Modify by Amy 2014/10/29 +必填及系統日檢查
         If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
            MsgBox .Label6 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            .MaskEdBox1.SetFocus
            Exit Sub
         End If
         If DateCheck(.MaskEdBox1.Text) = MsgText(603) Then
            MsgBox .Label6 & MsgText(63), , MsgText(5)
            strControlButton = MsgText(602)
            .MaskEdBox1.SetFocus
            Exit Sub
         End If
         If MaskEdBox1.Enabled = True Then
            If ChkWorkData(Text24, DBDATE(MaskEdBox1), strMsg) = False Then
                MsgBox Label6 & strMsg, , MsgText(5)
                strControlButton = MsgText(602)
                MaskEdBox1.SetFocus
                Exit Sub
            End If
         End If
         'end 2014/10/29
'         If .MaskEdBox2.Text <> MsgText(601) And .MaskEdBox2.Text <> MsgText(29) Then
            If DateCheck(.MaskEdBox2.Text) = MsgText(603) Then
               MsgBox .Label7 & MsgText(63), , MsgText(5)
               strControlButton = MsgText(602)
               .MaskEdBox2.SetFocus
               Exit Sub
            End If
'         End If
      End If
      'add by sonia 2024/5/28
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select a1p04,count(*) from acc1p0 where a1p01 = '" & Text24 & "' and a1p02 = 'D' and a1p04 = '" & Text1 & "' and a1p05='2401' and a1p08>0 group by a1p04", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         If adoquery.Fields(1) > 1 Then
            MsgBox "分錄中貸方暫收款資料不可超過１筆！", , MsgText(5)
            strControlButton = MsgText(602)
            adoquery.Close
            Exit Sub
         End If
      Else
         MsgBox "分錄中無貸方暫收款資料！", , MsgText(5)
         strControlButton = MsgText(602)
         adoquery.Close
         Exit Sub
      End If
      adoquery.Close
      'end 2024/5/28
      If strSaveConfirm = MsgText(3) Then
         If .adoacc0t0.RecordCount <> 0 Then
            .adoacc0t0.MoveFirst
            .adoacc0t0.Find "a0t01 = '" & .Text1 & "'", 0, adSearchForward, 1
            If .adoacc0t0.EOF = False Then
               strControlButton = MsgText(602)
               Exit Sub
            End If
         End If
         .adoacc0t0.AddNew
      End If
      .adoacc0t0.Fields("a0t01").Value = .Text1
      If .Text2 <> MsgText(601) Then
         .adoacc0t0.Fields("a0t02").Value = .Text2
      Else
         .adoacc0t0.Fields("a0t02").Value = Null
      End If
      If .Text3 <> MsgText(601) Then
         .adoacc0t0.Fields("a0t05").Value = .Text3
      Else
         .adoacc0t0.Fields("a0t05").Value = Null
      End If
      If .Text4 <> MsgText(601) Then
         .adoacc0t0.Fields("a0t06").Value = .Text4
      Else
         .adoacc0t0.Fields("a0t06").Value = Null
      End If
      If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
         .adoacc0t0.Fields("a0t03").Value = Val(FCDate(.MaskEdBox1.Text))
      Else
         .adoacc0t0.Fields("a0t03").Value = Null
      End If
      If .MaskEdBox2.Text <> MsgText(601) And .MaskEdBox2.Text <> MsgText(29) Then
         .adoacc0t0.Fields("a0t04").Value = Val(FCDate(.MaskEdBox2.Text))
      Else
         .adoacc0t0.Fields("a0t04").Value = Null
      End If
'      If .Text16 <> MsgText(601) Then
'         .adoacc0t0.Fields("a0t08").Value = Val(.Text16)
'      Else
'         .adoacc0t0.Fields("a0t08").Value = 0
'      End If
      If .Text12 <> MsgText(601) Then
         .adoacc0t0.Fields("a0t17").Value = .Text12
      Else
         .adoacc0t0.Fields("a0t17").Value = Null
      End If
      If .Text17 <> MsgText(601) Then
         .adoacc0t0.Fields("a0t08").Value = Val(.Text17)
      Else
         .adoacc0t0.Fields("a0t08").Value = 0
      End If
      If strSaveConfirm = MsgText(3) Then
         .adoacc0t0.Fields("a0t11").Value = Val(strSrvDate(2))
         .adoacc0t0.Fields("a0t12").Value = ServerTime
         .adoacc0t0.Fields("a0t13").Value = strUserNum
      Else
         .adoacc0t0.Fields("a0t14").Value = Val(strSrvDate(2))
         .adoacc0t0.Fields("a0t15").Value = ServerTime
         .adoacc0t0.Fields("a0t16").Value = strUserNum
      End If
      'Add By Sindy 2013/12/30
      .adoacc0t0.Fields("a0t18").Value = .Text24 '公司別
      '2013/12/30 END
      .adoacc0t0.UpdateBatch
      .RecordShow
      
      'Added by Lydia 2024/11/28 未收文客戶暫收款管制：特殊修改
      If Left(m_UpdStatus, 1) = "1" And Mid(m_UpdStatus, 2) <> Trim(Text4) Then
         '未收文客戶暫收款作業：X03072010用於新客戶未建檔，因沒有客戶代號卻先收款
         strExc(1) = CustomerQuery(Mid(m_UpdStatus, 2), 1)
         strExc(2) = CustomerQuery(Trim(Text4), 1)
         strExc(3) = ""
         strSql = "select a1p01,a1p15,a1p22,ax210 from acc1p0,acc021 where a1p01='" & Trim(Text24) & "' and a1p04='" & Trim(Text1) & "' " & _
                  "and a1p01=ax201(+) and a1p22=ax202(+) and a1p03=ax203(+) group by a1p01,a1p15,a1p22,ax210 "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            RsTemp.MoveFirst
            '更新A1p15對沖客戶、a1p14摘要中之客戶名稱中文
            strSql = "Update acc1p0 set a1p15='" & Trim(Text4) & "', a1p14=replace(a1p14,'" & ChgSQL(strExc(1)) & "','" & strExc(2) & "') where a1p01='" & Trim(Text24) & "' and a1p04='" & Trim(Text1) & "' and a1p15='" & Mid(m_UpdStatus, 2) & "' "
            cnnConnection.Execute strSql
            Do While Not RsTemp.EOF
               If "" & RsTemp.Fields("a1p22") <> "" And "" & RsTemp.Fields("ax210") <> "" Then
                  '更新ax208對沖客戶、ax212摘要中之客戶名稱中文
                  strSql = "Update acc021 set ax208='" & Trim(Text4) & "',ax212=replace(ax212,'" & ChgSQL(strExc(1)) & "','" & strExc(2) & "') where ax201='" & Trim(Text24) & "' and ax202='" & RsTemp.Fields("a1p22") & "' and ax208='" & Mid(m_UpdStatus, 2) & "' "
                  cnnConnection.Execute strSql
                  strExc(3) = strExc(3) & IIf(strExc(3) <> "", ",", "") & RsTemp.Fields("a1p22")
               End If
               RsTemp.MoveNext
            Loop
            If strExc(3) <> "" Then
               m_UpdMsg = "已修改" & Trim(Text24) & "公司傳票" & strExc(3) & "之對沖客戶，帳務傳票請自行修改！"
            End If
         End If
      End If
      'end 2024/11/28
      
Checking:
   'Add by Morgan 2005/10/31 所有錯誤都要控制
   'If Err.Number = 0 Or Err.Number = -2147217864 Then
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   strControlButton = MsgText(602) 'Add by Morgan 2005/10/31
   End With
End Sub

'Add by Amy 2014/10/29
'為資料一致更新acc1p0
Public Sub UpdateAcc1p0()
    Dim strUpd As String
    
On Error GoTo ChkHand

    If Text2 = "1" And Val(MaskEdBox1.Tag) <> Val(FCDate(MaskEdBox1)) Then
        strUpd = "Update Acc1p0 set a1p18=" & Val(FCDate(MaskEdBox1)) & _
                     " Where a1p01='" & Text24 & "' And a1p04='" & Text1 & "' And a1p02='D' "
        adoTaie.Execute strUpd
    End If

ChkHand:
    If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox "UpdateAcc1p0 錯誤:" & Err.Description, , MsgText(5)
   strControlButton = MsgText(602)
End Sub

Public Sub SetData(ByVal strWork As String)
    Select Case strWork
        Case "Refresh"
            Text24.Enabled = True
            MaskEdBox1.Enabled = True
            If CheckExistA1p22(Text24, "D", Text1) = True Then
                'a1p22有值不可修改輸入日及公司別
                Text24.Enabled = False
                MaskEdBox1.Enabled = False
            ElseIf adoadodc1.RecordCount > 0 Then
                'acc1p0有資料公司別不可改
                Text24.Enabled = False
            End If
        Case "F2"
            Text24.Enabled = True
            MaskEdBox1.Enabled = True
        Case "F3", "F9"
            '解改日期存檔再修改不會存acc1p0 (因tag只記錄前一次改前資料)
            MaskEdBox1.Tag = Val(FCDate(MaskEdBox1))
            'Added by Lydia 2024/11/28 未收文客戶暫收款管制
            If m_UpdMsg <> "" Then
               MsgBox m_UpdMsg, vbInformation + vbOKOnly
            End If
            'end 2024/11/28
        Case Else
    End Select
End Sub
'end 2014/10/29

'Add by Amy 2014/10/29  由acc_cls搬回
Public Sub Frmacc11a0_Clear()
   With Frmacc11a0
      .Text1 = ""
      .Text2 = ""
      .Text3 = ""
      .Text18 = ""
      .Text4 = "X"
      .Text13 = ""
      If .MaskEdBox1.Text = MsgText(29) Or .MaskEdBox1.Text = MsgText(601) Then
         .MaskEdBox1.Mask = ""
         .MaskEdBox1.Text = CFDate(ACDate(ServerDate))
         .MaskEdBox1.Mask = DFormat
      End If
      .MaskEdBox1.Tag = "" 'Add by Amy 2014/10/29
      If .MaskEdBox2.Text = MsgText(29) Or .MaskEdBox2.Text = MsgText(601) Then
         .MaskEdBox2.Mask = ""
         .MaskEdBox2.Text = ""
         .MaskEdBox2.Mask = DFormat
      End If
      .Text11 = ""
      .Text10 = ""
      .Text12 = ""
      .Text17 = ""
      .Text26 = ""    'add by sonia 2024/5/22
      .AdodcRefresh
      .AdodcClear
      .Text2.Enabled = True 'Added by Lydia 2024/11/28
      .Text2.SetFocus
   End With
End Sub

'由acc_del搬回
Public Sub Frmacc11a0_Delete()
On Error GoTo Checking
   With Frmacc11a0
      If DeleteCheck("select a0t01 from acc0t0 where a0t01 = '" & .Text1 & "'") = MsgText(603) Then
         Exit Sub
      End If
      'Modify by Amy 2014/10/29 原a1p01 = '1'
      adoTaie.Execute "delete from acc1p0 where a1p01 = '" & Text24 & "' and a1p02 = 'D' and a1p04 = '" & .Text1 & "'"
      .adoacc1p0.Requery
      adoTaie.Execute "delete from acc0t0 where a0t01 = '" & .Text1 & "'"
      .adoacc0t0.Requery
      .AdodcRefresh
      If .adoacc0t0.RecordCount <> 0 Then
         .adoacc0t0.MoveFirst
         .RecordShow
      Else
         StatusClear
      End If
   End With
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub
'end 2014/10/29
