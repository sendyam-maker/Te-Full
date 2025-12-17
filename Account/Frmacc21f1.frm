VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc21f1 
   AutoRedraw      =   -1  'True
   Caption         =   "抵帳結匯輸入"
   ClientHeight    =   5510
   ClientLeft      =   50
   ClientTop       =   270
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5510
   ScaleWidth      =   8760
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   3720
      Style           =   2  '單純下拉式
      TabIndex        =   21
      Top             =   5100
      Width           =   2900
   End
   Begin VB.TextBox Text23 
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
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   19
      Top             =   4590
      Width           =   765
   End
   Begin VB.TextBox Text22 
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
      Left            =   6870
      MaxLength       =   10
      TabIndex        =   18
      Top             =   4267
      Width           =   1572
   End
   Begin VB.TextBox Text21 
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
      Left            =   4068
      MaxLength       =   9
      TabIndex        =   17
      Top             =   4267
      Width           =   1572
   End
   Begin VB.TextBox Text19 
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
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   16
      Top             =   4260
      Width           =   1665
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
      Height          =   315
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   14
      Top             =   3930
      Width           =   528
   End
   Begin VB.TextBox Text3 
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
      Left            =   6840
      MaxLength       =   14
      TabIndex        =   11
      Top             =   3277
      Width           =   1572
   End
   Begin VB.ComboBox Combo3 
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
      Left            =   1320
      TabIndex        =   3
      Top             =   397
      Width           =   1572
   End
   Begin VB.TextBox Text20 
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
      Height          =   300
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   45
      Top             =   2496
      Width           =   855
   End
   Begin VB.TextBox Text17 
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
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   13
      Top             =   3607
      Width           =   1572
   End
   Begin VB.TextBox Text15 
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
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   12
      Top             =   3600
      Width           =   1665
   End
   Begin VB.TextBox Text11 
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
      Left            =   4080
      MaxLength       =   14
      TabIndex        =   10
      Top             =   3277
      Width           =   1572
   End
   Begin VB.TextBox Text13 
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
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   8
      Top             =   2947
      Width           =   1572
   End
   Begin VB.TextBox Text14 
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
      Height          =   300
      Left            =   4080
      TabIndex        =   37
      Top             =   2496
      Width           =   1332
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
      Height          =   350
      Left            =   7932
      Picture         =   "Frmacc21f1.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   23
      ToolTipText     =   "取消"
      Top             =   2460
      Width           =   350
   End
   Begin VB.TextBox Text9 
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
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   9
      Top             =   3270
      Width           =   1665
   End
   Begin VB.TextBox Text8 
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
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   7
      Top             =   2940
      Width           =   612
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
      Height          =   300
      Left            =   5520
      TabIndex        =   36
      Top             =   2496
      Width           =   1332
   End
   Begin VB.TextBox Text7 
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
      Height          =   315
      Left            =   4200
      TabIndex        =   35
      Top             =   795
      Width           =   1572
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   1320
      TabIndex        =   6
      Top             =   780
      Width           =   1572
   End
   Begin VB.CommandButton Command2 
      Caption         =   "列印抵帳資料"
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
      Left            =   6840
      TabIndex        =   22
      Top             =   5130
      Width           =   1692
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   20
      Top             =   5130
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00FFFFFF&
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
      MaxLength       =   14
      TabIndex        =   5
      Top             =   432
      Width           =   1572
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00FFFFFF&
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
      Left            =   4200
      MaxLength       =   13
      TabIndex        =   4
      Top             =   432
      Width           =   1572
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00FFFFFF&
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
      Left            =   4200
      MaxLength       =   14
      TabIndex        =   1
      Top             =   60
      Width           =   1572
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
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   0
      Top             =   60
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   6960
      TabIndex        =   2
      Top             =   60
      Width           =   1572
      _ExtentX        =   2769
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Frmacc21f1.frx":066A
      Height          =   1275
      Left            =   240
      TabIndex        =   24
      Top             =   1170
      Width           =   8295
      _ExtentX        =   14623
      _ExtentY        =   2258
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
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
      BeginProperty Column04 
         DataField       =   "a0g02"
         Caption         =   "銀行名稱"
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
            ColumnWidth     =   3569.953
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1340.221
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1399.748
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   4919.811
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   5859.78
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   336
      Left            =   240
      Top             =   1020
      Visible         =   0   'False
      Width           =   1200
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
   Begin MSForms.TextBox Text24 
      Height          =   330
      Left            =   2100
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   4590
      Width           =   1110
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "1958;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text16 
      Height          =   330
      Left            =   5670
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   2947
      Width           =   2775
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "4895;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text10 
      Height          =   330
      Left            =   5670
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   3607
      Width           =   2775
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "4895;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   345
      Left            =   4080
      TabIndex        =   15
      Top             =   3915
      Width           =   4350
      VariousPropertyBits=   679495707
      BackColor       =   16777215
      DisplayStyle    =   3
      Size            =   "7673;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label26 
      BackStyle       =   0  '透明
      Caption         =   "印表機:"
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
      Left            =   2880
      TabIndex        =   54
      Top             =   5130
      Width           =   975
   End
   Begin VB.Label Label25 
      BackStyle       =   0  '透明
      Caption         =   "對沖(業)"
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
      Left            =   372
      TabIndex        =   53
      Top             =   4621
      Width           =   972
   End
   Begin VB.Label Label24 
      BackStyle       =   0  '透明
      Caption         =   "對沖(其)"
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
      Left            =   5892
      TabIndex        =   52
      Top             =   4291
      Width           =   972
   End
   Begin VB.Label Label23 
      BackStyle       =   0  '透明
      Caption         =   "對沖(客)"
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
      Left            =   3120
      TabIndex        =   51
      Top             =   4291
      Width           =   972
   End
   Begin VB.Label Label21 
      BackStyle       =   0  '透明
      Caption         =   "對沖(本)"
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
      TabIndex        =   50
      Top             =   4291
      Width           =   972
   End
   Begin VB.Label Label20 
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
      Left            =   360
      TabIndex        =   49
      Top             =   3961
      Width           =   972
   End
   Begin VB.Label Label19 
      BackStyle       =   0  '透明
      Caption         =   "台幣金額"
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
      Left            =   5880
      TabIndex        =   48
      Top             =   3301
      Width           =   1032
   End
   Begin VB.Label Label18 
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
      Left            =   3144
      TabIndex        =   47
      Top             =   3961
      Width           =   840
   End
   Begin VB.Label Label22 
      BackStyle       =   0  '透明
      Caption         =   "筆數"
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
      TabIndex        =   46
      Top             =   2496
      Width           =   852
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      Height          =   2085
      Left            =   240
      Top             =   2880
      Width           =   8295
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "銀行代號"
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
      Left            =   3120
      TabIndex        =   44
      Top             =   3631
      Width           =   972
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "銀行帳號"
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
      TabIndex        =   43
      Top             =   3631
      Width           =   972
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "外幣金額"
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
      Left            =   3120
      TabIndex        =   42
      Top             =   3216
      Width           =   972
   End
   Begin VB.Label Label14 
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
      Left            =   3120
      TabIndex        =   41
      Top             =   2971
      Width           =   972
   End
   Begin VB.Label Label13 
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
      Left            =   3240
      TabIndex        =   40
      Top             =   2496
      Width           =   612
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "銀存匯率"
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
      Top             =   3301
      Width           =   972
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "借1/貸2"
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
      TabIndex        =   38
      Top             =   2971
      Width           =   972
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "台幣金額"
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
      Left            =   3120
      TabIndex        =   34
      Top             =   819
      Width           =   972
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "付款方式"
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
      TabIndex        =   33
      Top             =   819
      Width           =   972
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "(Y/N)"
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
      Left            =   2040
      TabIndex        =   32
      Top             =   5130
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   5070
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "是否結清"
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
      Left            =   360
      TabIndex        =   31
      Top             =   5130
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   15
      X2              =   8775
      Y1              =   5070
      Y2              =   5070
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   15
      X2              =   8775
      Y1              =   5025
      Y2              =   5025
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "外幣金額"
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
      Left            =   6000
      TabIndex        =   30
      Top             =   456
      Width           =   972
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "匯率"
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
      Left            =   3120
      TabIndex        =   29
      Top             =   420
      Width           =   972
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "幣別"
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
      TabIndex        =   28
      Top             =   456
      Width           =   972
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "結匯日期"
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
      Left            =   6000
      TabIndex        =   27
      Top             =   60
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "手續費"
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
      Left            =   3120
      TabIndex        =   26
      Top             =   84
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "匯票號碼"
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
      TabIndex        =   25
      Top             =   84
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc21f1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/07 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Combo2、Text10、Text16、Text24; Printer列印未改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoacc1i0 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoacc1p0 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adocheck As New ADODB.Recordset
Dim strSerialNo As String
Dim strAccNo As String
Dim strYes As String
Dim strDocuNo As String
Dim strDYes As String
Dim intCounter As Integer
'Add By Cheng 2003/05/21
Const m_dblLeft As Double = 500 '橫軸偏移值
Dim m_intPage As Integer '頁數
Dim m_lLastPos As Long 'Add by Morgan 2004/11/26 Grid 游標搜尋用
'Added by Lydia 2018/11/05
Dim strPrinter As String '系統預設印表機
Dim strPrtOrt As Integer '系統預設印表機的紙張方向

Private Sub Combo1_Validate(Cancel As Boolean)
   If Text1 <> MsgText(601) Then
      Exit Sub
   End If
   Select Case Mid(Combo1, 1, 1)
      Case "2", "3", "4"
         strAccNo = AccAutoNo("A", 5)
         strYes = AccSaveAutoNo(Mid(strAccNo, 1, 1), Mid(strAccNo, 5, 5))
         Text1 = strAccNo
   End Select
End Sub

'edit by nickc 2007/07/13
Private Sub Combo2_GotFocus()
OpenIme
TextInverse Combo2  'Added by Lydia 2021/12/14 Form 2.0的ComboBox的GotFocus不會全選反白
End Sub
Private Sub Combo2_Validate(Cancel As Boolean)
CloseIme
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo3_Validate(Cancel As Boolean)
   If Combo3 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc1y0", "a1y01", Combo3, Label4) = False Then
      Cancel = True
      Combo3.SetFocus
   End If
End Sub

Private Sub Command1_Click()
   If Adodc1.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc1.Recordset.Fields("a1p22").Value) = False Then
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select ax210 from acc021 where ax201 = '" & Adodc1.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc1.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            MsgBox MsgText(155), , MsgText(5)
            Text8.SetFocus
            adoquery.Close
            Exit Sub
         End If
         adoquery.Close
      End If
   End If
   AdodcDelete
   AdodcClear
   SumShow
End Sub
'Memo by Lydia 2018/11/05 列印抵帳資料
Private Sub Command2_Click()
   Screen.MousePointer = vbHourglass
   PUB_RestorePrinter cmbPrinter 'Added by Lydia 2018/11/05 改印表機
   PrintData
   PUB_RestorePrinter strPrinter, strPrtOrt 'Added by Lydia 2018/11/05 還原系統印表機
   Screen.MousePointer = vbDefault
End Sub

Private Sub DataGrid2_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   strSerialNo = Adodc1.Recordset.Fields("a1p03").Value
   m_lLastPos = Adodc1.Recordset.AbsolutePosition 'Add by Morgan 2006/12/5
   AdodcShow
End Sub

'Added by Lydia 2021/12/07
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Added by Lydia 2021/12/07 Form2.0 記錄鍵盤傳入順序
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   'Added by Lydia 2018/11/05 預設印表機選項
   strPrtOrt = Printer.Orientation
   PUB_SetPrinter "Frmacc21f0", cmbPrinter, strPrinter
   '2018/11/05
   
   'Modified by Lydia 2021/12/07 改成模組
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   Me.Caption = Me.Caption & " --> " & strItemNo
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
   Me.Caption = Me.Caption & " --> " & strItemNo
   PUB_InitForm Me, 8850, 6070, strBackPicPath1
   'end 2021/12/07
   
   MaskEdBox1.Mask = MsgText(601)
   MaskEdBox1.Text = MsgText(601)
   MaskEdBox1.Mask = DFormat
   Combo1.AddItem ComboItem(71)
   Combo1.AddItem ComboItem(72)
   Combo1.AddItem ComboItem(73)
   Combo1.AddItem ComboItem(74)
   OpenTable
   If adoacc1i0.RecordCount <> 0 Then
      FormShow
      AdodcRefresh
   End If
   'Add by Amy 2014/11/06 a1p22有值不可改結匯日(若adoacc1i0沒資料但已有傳票所以寫於此)
   If IsNull(adoadodc1.Fields("a1p22")) Then
      MaskEdBox1.Enabled = True
   Else
      MaskEdBox1.Enabled = False
   End If
   'end 2014/11/06
   SumShow
   FormDisabled
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim strUpd As String 'Add by Amy 2014/11/06
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   CreDebCheck
   If CreDebCheck <> MsgText(602) Then
      tool7_enabled
      MsgBox MsgText(11), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
   'Modify by Amy 2014/11/06 避免修改日期為空 按取消仍更新
   If Val(MaskEdBox1.Tag) <> Val(FCDate(MaskEdBox1)) Then
      strUpd = strUpd & ", a1p18 = " & Val(FCDate(MaskEdBox1.Text))
   End If
   If strDocuNo <> "" Then '1.若acc1i0沒資料strDocuNo及strDYes 為空更新會error 2.新增項目需更新a1p22
      'adoTaie.Execute "update acc1p0 set a1p22 = " & strDocuNo & ", a1p27 = " & strDYes & ", a1p18 = " & Val(FCDate(MaskEdBox1.Text)) & " where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & strItemNo & "'"
      strUpd = "update acc1p0 set a1p22 = " & strDocuNo & ", a1p27 = " & strDYes & strUpd & " where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & strItemNo & "'"
      adoTaie.Execute strUpd
   End If
   'end 2014/11/06
   
   strTrackMode = "" 'Added by Lydia 2021/12/07 Form2.0 記錄鍵盤傳入順序(清除)
   
   'Added by Lydia 2018/11/05 若有變動印表機, 則更新列印設定
    If Me.cmbPrinter.Text <> Me.cmbPrinter.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, "Frmacc21f0", Me.cmbPrinter.Name, "0", "0", Me.cmbPrinter.Text
      Frmacc21f0.cmbPrinter.Text = Me.cmbPrinter.Text '三畫面(Frmacc21f0~Frmacc21f2)的預設印表機一致
    End If
   'end 2018/11/05
   
   tool1_enabled
   Frmacc21f0.Show
   Set Frmacc21f1 = Nothing
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox Label3 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label3 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc1i0.CursorLocation = adUseClient
   adoacc1i0.Open "select * from acc1i0 where a1i01 = '" & strItemNo & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc1p0, acc010, acc0g0 where a1p05 = a0101 (+) and a1p10 = a0g01 (+) and a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & strItemNo & "' order by a1p03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from acc1y0 order by a1y01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoquery.EOF = False
      Combo3.AddItem adoquery.Fields("a1y01").Value
      adoquery.MoveNext
   Loop
   adoquery.Close
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
   If IsNull(adoacc1i0.Fields("a1i02").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = adoacc1i0.Fields("a1i02").Value
   End If
   If IsNull(adoacc1i0.Fields("a1i04").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = adoacc1i0.Fields("a1i04").Value
   End If
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc1i0.Fields("a1i03").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(adoacc1i0.Fields("a1i03").Value)
   End If
   MaskEdBox1.Tag = "" & adoacc1i0.Fields("a1i03").Value 'Add by Amy 2014/11/06
   MaskEdBox1.Mask = DFormat
   If IsNull(adoacc1i0.Fields("a1i05").Value) Then
      Combo3 = MsgText(601)
   Else
      Combo3 = adoacc1i0.Fields("a1i05").Value
   End If
   If IsNull(adoacc1i0.Fields("a1i06").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = adoacc1i0.Fields("a1i06").Value
   End If
   If IsNull(adoacc1i0.Fields("a1i07").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = adoacc1i0.Fields("a1i07").Value
   End If
   If IsNull(adoacc1i0.Fields("a1i15").Value) Then
      Combo1 = MsgText(601)
   Else
      Combo1 = Combo1.List(Val(adoacc1i0.Fields("a1i15").Value) - 1)
   End If
   'If IsNull(adoacc1i0.Fields("a1i08").Value) Then
   '   Text6 = MsgText(601)
   'Else
   '   Text6 = adoacc1i0.Fields("a1i08").Value
   'End If
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub

Private Sub Text13_Change()
   If Text13 = MsgText(601) Then
      Exit Sub
   End If
   Text16 = A0102Query(Text13)
End Sub

Private Sub Text13_GotFocus()
   TextInverse Text13
End Sub

Private Sub Text13_Validate(Cancel As Boolean)
   If Text13 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc010", "a0101", Text13, Label14) = False Then
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

Private Sub Text17_Change()
   If Text17 = MsgText(601) Then
      Exit Sub
   End If
   Text10 = A0g02Query(Text17)
End Sub

Private Sub Text17_GotFocus()
   TextInverse Text17
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text17_Validate(Cancel As Boolean)
   If Text17 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc0g0", "a0g01", Text17, Label17) = False Then
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text18_GotFocus()
   TextInverse Text18
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text18_Validate(Cancel As Boolean)
   If Text18 <> MsgText(601) Then
      If ExistCheck("acc090", "a0901", Text18, Label20) = False Then
         Cancel = True
         Exit Sub
      End If
   End If
   If CheckDept(Text13, Text18) = False Then
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

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text21_GotFocus()
   TextInverse Text21
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text22_GotFocus()
   TextInverse Text22
End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text23_GotFocus()
   TextInverse Text23
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'Add by Morgan 2007/2/6 員工已離職要提醒
Private Sub Text23_Validate(Cancel As Boolean)
   Text24 = ""
   If Text23 <> MsgText(601) Then
      If PUB_GetStaffState(Text23.Text, strExc(1), True) = 0 Then
         Cancel = True
         TextInverse Text23
      Else
         Text24.Text = strExc(1)
      End If
   End If
End Sub
'end 2007/2/6
Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'Added by Lydia 2016/07/18
Private Sub Text3_LostFocus()
    'Added by Lydia 2016/07/18 規費只能輸入整數
    If Left(Trim(Text13), 4) = "2201" And Text3 <> "" And Text3 <> Format(Val(Text3), DAmount) Then
        MsgBox "規費只能輸入整數!", vbCritical
        Text3.SetFocus
    End If
    'end 2016/07/18
End Sub

Private Sub Text4_Change()
   If Text4 = MsgText(601) Then
      Exit Sub
   End If
   Text7 = Int(Val(Text4) * Val(Text5))
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text5_Change()
   If Text5 = MsgText(601) Then
      Exit Sub
   End If
   Text7 = Int(Val(Text4) * Val(Text5))
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

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc1p0, acc010, acc0g0 where a1p05 = a0101 and a1p10 = a0g01 (+) and a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & strItemNo & "' order by a1p03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc1.Recordset.Fields("a1p22").Value) Then
         strDocuNo = "null"
         strDYes = "null"
      Else
         strDocuNo = "'" & Adodc1.Recordset.Fields("a1p22").Value & "'"
         strDYes = "'Y'"
      End If
      'Modify by Morgan 2006/12/5
      If m_lLastPos > 1 Then
         If m_lLastPos < Adodc1.Recordset.RecordCount Then
            Adodc1.Recordset.Move m_lLastPos - 1, adBookmarkFirst
         Else
            Adodc1.Recordset.MoveLast
         End If
      Else
         Adodc1.Recordset.MoveFirst
      End If
   Else
      strDocuNo = "null"
      strDYes = "null"
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  儲存 Adodc 之資料
'
'*************************************************
Private Sub Acc1p0Save()
On Error GoTo Checking
      If Text13 = MsgText(601) Then
         MsgBox MsgText(10) & Label14, , MsgText(5)
         strControlButton = MsgText(602)
         Text13.SetFocus
         Exit Sub
      Else
         If ExistCheck("acc010", "a0101", Text13, Label14) = False Then
            strControlButton = MsgText(602)
            Text13.SetFocus
            Exit Sub
         End If
         If CheckDept(Text13, Text18) = False Then
            MsgBox MsgText(103), , MsgText(5)
            strControlButton = MsgText(602)
            Text18.SetFocus
            Exit Sub
         End If
         If Text17 <> MsgText(601) Then
            If ExistCheck("acc0g0", "a0g01", Text17, Label17) = False Then
               Text17.SetFocus
               Exit Sub
            End If
         End If
      End If
      
      'add by sonia 2015/12/30 檢查民國105年起法務收入科目不可使用
      intI = PUB_AccNoEnable(Text13, Val(FCDate(MaskEdBox1.Text)))
      If intI <> 0 Then
         strControlButton = MsgText(602)
         Text13.SetFocus
         Exit Sub
      End If
      'end 2015/12/30
      'Add by Morgan 2007/2/5 檢查科目部門&智權人員是否正確
      intI = PUB_AccNoGood(Text13, Text18, Text23)
      If intI <> 0 Then
         strControlButton = MsgText(602)
         If intI = 1 Then
            Text13.SetFocus
         ElseIf intI = 2 Then
            Text18.SetFocus
         ElseIf intI = 3 Then
            Text23.SetFocus
         End If
         Exit Sub
      End If
      'end 2007/2/5
      'Added by Lydia 2016/07/18 規費只能輸入整數
      If Left(Trim(Text13), 4) = "2201" And Text3 <> "" And Text3 <> Format(Val(Text3), DAmount) Then
          MsgBox "規費只能輸入整數!", vbCritical
          Text3.SetFocus
          Exit Sub
      End If
      'end 2016/07/18
      
      If Adodc1.Recordset.RecordCount <> 0 Then
         If IsNull(Adodc1.Recordset.Fields("a1p22").Value) = False Then
            adoquery.CursorLocation = adUseClient
            adoquery.Open "select ax210 from acc021 where ax201 = '" & Adodc1.Recordset.Fields("a1p01").Value & "' and ax202 = '" & Adodc1.Recordset.Fields("a1p22").Value & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
            If adoquery.RecordCount <> 0 Then
               MsgBox MsgText(155), , MsgText(5)
               strControlButton = MsgText(602)
               Text8.SetFocus
               adoquery.Close
               Exit Sub
            End If
            adoquery.Close
         End If
      End If
      If adoacc1p0.State = adStateOpen Then
         adoacc1p0.Close
      End If
      adoacc1p0.CursorLocation = adUseClient
      adoacc1p0.Open "select * from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & strItemNo & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If adoacc1p0.RecordCount = 0 Then
         adoacc1p0.AddNew
         adoacc1p0.Fields("a1p01").Value = "1"
         adoacc1p0.Fields("a1p02").Value = "K"
         adoacc1p0.Fields("a1p03").Value = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & strItemNo & "'", 3)
         adoacc1p0.Fields("a1p04").Value = strItemNo
         m_lLastPos = Adodc1.Recordset.AbsolutePosition 'Add by Morgan 2006/12/5
      End If
      adoacc1p0.Fields("a1p05").Value = Text13
      If Text9 <> MsgText(601) Then
         adoacc1p0.Fields("a1p20").Value = Val(Text9)
      Else
         adoacc1p0.Fields("a1p20").Value = 0
      End If
      If Text11 <> MsgText(601) Then
      '   Select Case Text8
      '      Case "1"
      '         adoacc1p0.Fields("a1p07").Value = Val(Text11) * Val(Text9)
      '         adoacc1p0.Fields("a1p08").Value = 0
      '      Case "2"
      '         adoacc1p0.Fields("a1p08").Value = Val(Text11) * Val(Text9)
      '         adoacc1p0.Fields("a1p07").Value = 0
      '      Case Else
      '         adoacc1p0.Fields("a1p07").Value = 0
      '         adoacc1p0.Fields("a1p08").Value = 0
      '   End Select
         adoacc1p0.Fields("a1p21").Value = Val(Text11)
      Else
      '   adoacc1p0.Fields("a1p07").Value = 0
      '   adoacc1p0.Fields("a1p08").Value = 0
         adoacc1p0.Fields("a1p21").Value = 0
      End If
      Select Case Text8
         Case "1"
            If Text3 <> MsgText(601) Then
               adoacc1p0.Fields("a1p07").Value = Val(Text3)
            Else
               adoacc1p0.Fields("a1p07").Value = 0
            End If
            adoacc1p0.Fields("a1p08").Value = 0
         Case Else
            If Text3 <> MsgText(601) Then
               adoacc1p0.Fields("a1p08").Value = Val(Text3)
            Else
               adoacc1p0.Fields("a1p08").Value = 0
            End If
            adoacc1p0.Fields("a1p07").Value = 0
      End Select
      If Text15 <> MsgText(601) Then
         adoacc1p0.Fields("a1p11").Value = Text15
      Else
         adoacc1p0.Fields("a1p11").Value = Null
      End If
      If Text17 <> MsgText(601) Then
         adoacc1p0.Fields("a1p10").Value = Text17
      Else
         adoacc1p0.Fields("a1p10").Value = Null
      End If
'      If Text1 <> MsgText(601) Then
'         adoacc1p0.Fields("a1p09").Value = Text1
'      Else
'         adoacc1p0.Fields("a1p09").Value = Null
'      End If
      'modify by sonia 2021/1/28 加傳本所案號以判別FCP,FCT英日文組
      'If AccNoToSalesNo(Text13) <> "" Then
      '   adoacc1p0.Fields("a1p16").Value = AccNoToSalesNo(Text13)
      If AccNoToSalesNo(Text13, Text19) <> "" Then
         adoacc1p0.Fields("a1p16").Value = AccNoToSalesNo(Text13, Text19)
      'end 2021/1/28
      End If
      If MaskEdBox1.Text <> MsgText(29) Then
         adoacc1p0.Fields("a1p18").Value = Val(FCDate(MaskEdBox1.Text))
      Else
         adoacc1p0.Fields("a1p18").Value = Null
      End If
      If Combo2 <> MsgText(601) Then
         adoacc1p0.Fields("a1p14").Value = Combo2
        'Add By Cheng 2004/02/27
        Me.Combo2.AddItem Combo2.Text
        'End
      Else
         adoacc1p0.Fields("a1p14").Value = Null
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
         adoacc1p0.Fields("a1p15").Value = Text21
      Else
         adoacc1p0.Fields("a1p15").Value = Null
      End If
      If Text22 <> MsgText(601) Then
         adoacc1p0.Fields("a1p30").Value = Text22
      Else
         adoacc1p0.Fields("a1p30").Value = Null
      End If
      If Text23 <> MsgText(601) Then
         adoacc1p0.Fields("a1p16").Value = Text23
      Else
         adoacc1p0.Fields("a1p16").Value = Null
      End If
      adoacc1p0.UpdateBatch
      adoacc1p0.Close
      adoTaie.Execute "update acc1p0 set a1p09 = '" & Text1 & "' where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & strItemNo & "'"
      strSerialNo = MsgText(601)
      m_lLastPos = Adodc1.Recordset.AbsolutePosition 'Add by Morgan 2006/12/5
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
   Text13 = Adodc1.Recordset.Fields("a1p05").Value
   If IsNull(Adodc1.Recordset.Fields("a1p20").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = Adodc1.Recordset.Fields("a1p20").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p07").Value) Or Adodc1.Recordset.Fields("a1p07").Value = 0 Then
      If IsNull(Adodc1.Recordset.Fields("a1p08").Value) Then
'         Text11 = MsgText(601)
         Text8 = MsgText(601)
         Text3 = MsgText(601)
      Else
'         Text11 = Adodc1.Recordset.Fields("a1p08").Value
         Text8 = "2"
         Text3 = Adodc1.Recordset.Fields("a1p08").Value
      End If
   Else
'      Text11 = Adodc1.Recordset.Fields("a1p07").Value
      Text8 = "1"
      Text3 = Adodc1.Recordset.Fields("a1p07").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p21").Value) Then
      Text11 = MsgText(601)
   Else
      Text11 = Adodc1.Recordset.Fields("a1p21").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p11").Value) Then
      Text15 = MsgText(601)
   Else
      Text15 = Adodc1.Recordset.Fields("a1p11").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p10").Value) Then
      Text17 = MsgText(601)
      Text10 = MsgText(601)
   Else
      Text17 = Adodc1.Recordset.Fields("a1p10").Value
      Text10 = A0g02Query(Text17)
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p14").Value) Then
      Combo2 = MsgText(601)
   Else
      Combo2 = Adodc1.Recordset.Fields("a1p14").Value
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
   If IsNull(Adodc1.Recordset.Fields("a1p15").Value) Then
      Text21 = MsgText(601)
   Else
      Text21 = Adodc1.Recordset.Fields("a1p15").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p30").Value) Then
      Text22 = MsgText(601)
   Else
      Text22 = Adodc1.Recordset.Fields("a1p30").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p16").Value) Then
      Text23 = MsgText(601)
      Text24 = ""
   Else
      Text23 = Adodc1.Recordset.Fields("a1p16").Value
      Text24 = StaffQuery(Text23)
   End If
End Sub

'*************************************************
'  清除顯示資料
'
'*************************************************
Public Sub AdodcClear()
   Text8 = ""
   Text13 = ""
   Text16 = ""
   Text9 = "1"
   Text11 = ""
   Text3 = ""
   Text15 = ""
   Text17 = ""
   Text10 = ""
   Combo2 = ""
   Text18 = ""
   Text19 = ""
   Text21 = ""
   Text22 = ""
   Text23 = ""
   Text8.SetFocus
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
Dim strCase(1 To 4) As String
Dim ii As Integer
   
   Call PUB_SaveTrackMode(1, KeyCode)  'Added by Lydia 2021/12/07 Form2.0 記錄鍵盤傳入順序
   
   Select Case KeyCode
      Case vbKeyInsert
         'Added by Lydia 2021/12/07 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
         If PUB_ChkTrackMode = False Then
             Exit Sub
         End If
         'end 2021/12/07
         If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
            Exit Sub
         End If
         'Added by Lydia 2021/12/07 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
         If PUB_ChkUniText(Me, , True, "ComboBox") = False Then
             Exit Sub
         End If
         'end 2021/12/07
         
        'Add By Cheng 2004/02/27
        '對沖(本)
        If Me.Text19.Text <> "" Then
            ChgCaseNo Me.Text19.Text, strCase
            Me.Text19.Text = ""
            For ii = LBound(strCase) To UBound(strCase)
                Me.Text19.Text = Me.Text19.Text & strCase(ii)
            Next ii
            If ChkOurCase(Me.Text19.Text) = False Then
                Me.Text19.SetFocus
                Text19_GotFocus
                Exit Sub
            End If
        End If
        '對沖(客)
        If Me.Text21.Text <> "" Then
            Me.Text21.Text = Left(Me.Text21.Text & "00000000", 9)
            If ChkOurCust(Me.Text21.Text) = False Then
                Me.Text21.SetFocus
                Text21_GotFocus
                Exit Sub
            End If
        End If
        '對沖(業)
        If Me.Text23.Text <> "" Then
            If ChkOurStaff(Me.Text23.Text) = False Then
                Me.Text23.SetFocus
                Text23_GotFocus
                Exit Sub
            End If
        End If
        'End
         Frmacc21f1_Save
         If strControlButton <> MsgText(602) Then
            Acc1p0Save
         End If
         If strControlButton <> MsgText(602) Then
            AdodcRefresh
            AdodcClear
            SumShow
            Text8.SetFocus
         End If
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
   adoaccsum.Open "select sum(a1p07), sum(a1p08), count(*) from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p04 = '" & strItemNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text14 = MsgText(601)
      Else
         Text14 = Format(adoaccsum.Fields(0).Value, FDollar)
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
      Text14 = MsgText(601)
      Text12 = MsgText(601)
      Text20 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  刪除 Adodc1 之資料
'
'*************************************************
Private Sub AdodcDelete()
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & strItemNo & "'"
   AdodcRefresh
   AdodcClear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

'*************************************************
'  關閉分錄欄位輸入狀態
'
'*************************************************
Public Sub FormDisabled()
   Text8.Enabled = False
   Text13.Enabled = False
   Text9.Enabled = False
   Text11.Enabled = False
   Text3.Enabled = False
   Text15.Enabled = False
   Text17.Enabled = False
   Command1.Enabled = False
   Text18.Enabled = False
   Text19.Enabled = False
   Text21.Enabled = False
   Text22.Enabled = False
   Text23.Enabled = False
End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'
'*************************************************
Public Sub FormEnabled()
   Text8.Enabled = True
   Text13.Enabled = True
   Text9.Enabled = True
   Text11.Enabled = True
   Text3.Enabled = True
   Text15.Enabled = True
   Text17.Enabled = True
   Command1.Enabled = True
   Text18.Enabled = True
   Text19.Enabled = True
   Text21.Enabled = True
   Text22.Enabled = True
   Text23.Enabled = True
End Sub

'*************************************************
'  借貸方檢核
'
'*************************************************
Public Function CreDebCheck() As String
   If Text12 = Text14 Then
      CreDebCheck = MsgText(602)
      Exit Function
   End If
   CreDebCheck = MsgText(603)
End Function

'*************************************************
'  列印抵帳資料
'
'*************************************************
Public Sub PrintData()
Dim strAmount As String
Dim intLength As Integer
Dim strCurrency As String
'Add By Cheng 2003/05/21
Dim strCaseNo As String '本所案號
Dim strFaNo As String '代理人編號

   intCounter = 0
   m_intPage = 1
   Printer.FontSize = 12
   '帳單資料
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from acc151, acc150 where axf01 = a1501 and a1512 = '" & strItemNo & "' order by a1504 asc", adoTaie, adOpenStatic, adLockReadOnly
    'Add By Cheng 2003/05/21
    strCaseNo = "" & adoquery.Fields("axf03").Value
    strFaNo = "" & adoquery.Fields("a1503").Value
    PrintHead strCaseNo, strFaNo
   Printer.CurrentX = 0 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Your Debit Notes"
   Printer.CurrentX = 2000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Our Ref"
   Printer.CurrentX = 5000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Date"
   If IsNull(adoquery.Fields("a1505").Value) Then
      strCurrency = ""
   Else
      strCurrency = adoquery.Fields("a1505").Value
   End If
   Printer.CurrentX = 7000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Amount(" & strCurrency & ")"
   intCounter = intCounter + 1
   Printer.Line (0 + m_dblLeft, 0 + intCounter * 300 - 50)-(9000 + m_dblLeft, 0 + intCounter * 300 - 50)
   Do While adoquery.EOF = False
        If intCounter > 48 Then
              Printer.CurrentX = 5000
              Printer.CurrentY = 0 + intCounter * 300
              Printer.Print "**" & m_intPage & "**"
              m_intPage = m_intPage + 1
              Printer.NewPage
            'Add By Cheng 2003/05/27
            PrintHead strCaseNo, strFaNo
            Printer.CurrentX = 0 + m_dblLeft
            Printer.CurrentY = 0 + intCounter * 300
            Printer.Print "Your Debit Notes"
            Printer.CurrentX = 2000 + m_dblLeft
            Printer.CurrentY = 0 + intCounter * 300
            Printer.Print "Our Ref"
            Printer.CurrentX = 5000 + m_dblLeft
            Printer.CurrentY = 0 + intCounter * 300
            Printer.Print "Date"
            If IsNull(adoquery.Fields("a1505").Value) Then
               strCurrency = ""
            Else
               strCurrency = adoquery.Fields("a1505").Value
            End If
            Printer.CurrentX = 7000 + m_dblLeft
            Printer.CurrentY = 0 + intCounter * 300
            Printer.Print "Amount(" & strCurrency & ")"
            intCounter = intCounter + 1
            Printer.Line (0 + m_dblLeft, 0 + intCounter * 300 - 50)-(9000 + m_dblLeft, 0 + intCounter * 300 - 50)
        End If
      Printer.CurrentX = 0 + m_dblLeft
      Printer.CurrentY = 0 + intCounter * 300
      If IsNull(adoquery.Fields("a1504").Value) Then
         Printer.Print ""
      Else
         Printer.Print adoquery.Fields("a1504").Value
      End If
      Printer.CurrentX = 2000 + m_dblLeft
      Printer.CurrentY = 0 + intCounter * 300
      If IsNull(adoquery.Fields("axf03").Value) Then
         Printer.Print ""
      Else
         Printer.Print adoquery.Fields("axf03").Value
      End If
      Printer.CurrentX = 5000 + m_dblLeft
      Printer.CurrentY = 0 + intCounter * 300
      If IsNull(adoquery.Fields("a1502").Value) Then
         Printer.Print ""
      Else
         Printer.Print Format(CADate(adoquery.Fields("a1502").Value), "####-##-##")
      End If
      If IsNull(adoquery.Fields("axf04").Value) = False Then
         strAmount = Format(Val(adoquery.Fields("axf04").Value), FDollar)
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = 9000 - intLength + m_dblLeft
         Printer.CurrentY = 0 + intCounter * 300
         Printer.Print strAmount
      End If
      intCounter = intCounter + 1
      adoquery.MoveNext
   Loop
   adoquery.Close
   intCounter = intCounter + 1
   Printer.Line (0 + m_dblLeft, 0 + intCounter * 300 - 50)-(9000 + m_dblLeft, 0 + intCounter * 300 - 50)
   Printer.CurrentX = 0 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Total amount due in your favor is"
   Printer.CurrentX = 6000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print strCurrency
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(axf04) from acc151, acc150 where axf01 = a1501 and a1512 = '" & strItemNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) = False Then
         strAmount = Format(Val(adoaccsum.Fields(0).Value), FDollar)
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = 9000 - intLength + m_dblLeft
         Printer.CurrentY = 0 + intCounter * 300
         Printer.Print strAmount
      End If
   End If
   adoaccsum.Close
    intCounter = intCounter + 1
    Printer.CurrentX = 5000
    Printer.CurrentY = 0 + intCounter * 300
    Printer.Print "**" & m_intPage & "**"
    m_intPage = m_intPage + 1
   
    m_intPage = 1
    intCounter = 0
    Printer.NewPage
    '請款單資料
    adoquery.CursorLocation = adUseClient
    adoquery.Open "select * from acc1k0 where a1k17 = '" & strItemNo & "' order by a1k01 asc", adoTaie, adOpenStatic, adLockReadOnly
    'Add By Cheng 2003/05/21
    strCaseNo = "" & adoquery.Fields("a1k13").Value & adoquery.Fields("a1k14").Value & adoquery.Fields("a1k15").Value & adoquery.Fields("a1k16").Value
    '2010/6/18 MODIFY BY SONIA 婧瑄說應改為請款對象
    'strFANo = "" & adoquery.Fields("a1k03").Value
    strFaNo = "" & adoquery.Fields("a1k28").Value
    '2010/6/19 END
    PrintHead strCaseNo, strFaNo
   Printer.CurrentX = 0 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Our Debit Notes"
   Printer.CurrentX = 2000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Our Ref"
   Printer.CurrentX = 5000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Date"
   If IsNull(adoquery.Fields("a1k18").Value) Then
      strCurrency = ""
   Else
      strCurrency = adoquery.Fields("a1k18").Value
   End If
   Printer.CurrentX = 7000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   '2012/5/17 MODIFY BY SONIA 固定用美金收款 Z10100011
   'Printer.Print "Amount(" & strCurrency & ")"
   Printer.Print "Amount(USD)"
   intCounter = intCounter + 1
   Printer.Line (0 + m_dblLeft, 0 + intCounter * 300 - 50)-(9000 + m_dblLeft, 0 + intCounter * 300 - 50)
   Do While adoquery.EOF = False
        If intCounter > 48 Then
              Printer.CurrentX = 5000
              Printer.CurrentY = 0 + intCounter * 300
              Printer.Print "**" & m_intPage & "**"
              m_intPage = m_intPage + 1
              Printer.NewPage
            'Add By Cheng 2003/05/27
            PrintHead strCaseNo, strFaNo
            Printer.CurrentX = 0 + m_dblLeft
            Printer.CurrentY = 0 + intCounter * 300
            Printer.Print "Our Debit Notes"
            Printer.CurrentX = 2000 + m_dblLeft
            Printer.CurrentY = 0 + intCounter * 300
            Printer.Print "Our Ref"
            Printer.CurrentX = 5000 + m_dblLeft
            Printer.CurrentY = 0 + intCounter * 300
            Printer.Print "Date"
            If IsNull(adoquery.Fields("a1k18").Value) Then
                strCurrency = ""
            Else
                strCurrency = adoquery.Fields("a1k18").Value
            End If
            Printer.CurrentX = 7000 + m_dblLeft
            Printer.CurrentY = 0 + intCounter * 300
            '2012/5/17 MODIFY BY SONIA 固定用美金收款 Z10100011
            'Printer.Print "Amount(" & strCurrency & ")"
            Printer.Print "Amount(USD)"
            intCounter = intCounter + 1
            Printer.Line (0 + m_dblLeft, 0 + intCounter * 300 - 50)-(9000 + m_dblLeft, 0 + intCounter * 300 - 50)
        End If
      Printer.CurrentX = 0 + m_dblLeft
      Printer.CurrentY = 0 + intCounter * 300
      If IsNull(adoquery.Fields("a1k01").Value) Then
         Printer.Print ""
      Else
         '2012/5/17 MODIFY BY SONIA 改印完整編號
         'Printer.Print Mid(adoquery.Fields("a1k01").Value, 2, Len(adoquery.Fields("a1k01").Value) - 1)
         Printer.Print adoquery.Fields("a1k01").Value
      End If
      Printer.CurrentX = 2000 + m_dblLeft
      Printer.CurrentY = 0 + intCounter * 300
      If IsNull(adoquery.Fields("a1k13").Value) Then
         Printer.Print ""
      Else
         Printer.Print adoquery.Fields("a1k13").Value & adoquery.Fields("a1k14").Value & adoquery.Fields("a1k15").Value * adoquery.Fields("a1k16").Value
      End If
      Printer.CurrentX = 5000 + m_dblLeft
      Printer.CurrentY = 0 + intCounter * 300
      If IsNull(adoquery.Fields("a1k02").Value) Then
         Printer.Print ""
      Else
         Printer.Print Format(CADate(adoquery.Fields("a1k02").Value), "####-##-##")
      End If
      If IsNull(adoquery.Fields("a1k08").Value) = False Then
         strAmount = Format(Val(adoquery.Fields("a1k08").Value) - Val(IIf(IsNull(adoquery.Fields("a1k06").Value), 0, adoquery.Fields("a1k06").Value)), FDollar)
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = 9000 - intLength + m_dblLeft
         Printer.CurrentY = 0 + intCounter * 300
         Printer.Print strAmount
      End If
      intCounter = intCounter + 1
      adoquery.MoveNext
   Loop
   adoquery.Close
   intCounter = intCounter + 1
   Printer.Line (0 + m_dblLeft, 0 + intCounter * 300 - 50)-(9000 + m_dblLeft, 0 + intCounter * 300 - 50)
   Printer.CurrentX = 0 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   Printer.Print "Total amount due in our favor is"
   Printer.CurrentX = 6000 + m_dblLeft
   Printer.CurrentY = 0 + intCounter * 300
   '2012/5/17 MODIFY BY SONIA 固定用美金收款 Z10100011
   'Printer.Print strCurrency
   Printer.Print "USD"
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(a1k08 - nvl(a1k06, 0)) from acc1k0 where a1k17 = '" & strItemNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) = False Then
         strAmount = Format(Val(adoaccsum.Fields(0).Value), FDollar)
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = 9000 - intLength + m_dblLeft
         Printer.CurrentY = 0 + intCounter * 300
         Printer.Print strAmount
      End If
   End If
   adoaccsum.Close
    intCounter = intCounter + 1
    Printer.CurrentX = 5000
    Printer.CurrentY = 0 + intCounter * 300
    Printer.Print "**" & m_intPage & "**"
    m_intPage = m_intPage + 1
   Printer.EndDoc
End Sub

'*************************************************
'  抬頭列印
'
'*************************************************
Private Sub PrintHead(strCaseNo As String, strFaNo As String)
Dim intRow As Integer
Dim StrSQLa As String
'add by nickc 2007/02/08
Dim strLanguage As String
   
    intRow = 0
    adocheck.CursorLocation = adUseClient
    adocheck.Open "select pa85 as Lang from patent, customer where substr(pa26, 1, 8) = cu01 and substr(pa26, 9, 1) = cu02 and " & ChgPatent(strCaseNo) & _
                  " union select tm53 as Lang from trademark, customer where substr(tm23, 1, 8) = cu01 and substr(tm23, 9, 1) = cu02 and " & ChgTradeMark(strCaseNo) & _
                  " union select sp34 as Lang from servicepractice, customer where substr(sp08, 1, 8) = cu01 and substr(sp08, 9, 1) = cu02 and " & ChgService(strCaseNo), adoTaie, adOpenStatic, adLockReadOnly
    If adocheck.RecordCount <> 0 Then
        If IsNull(adocheck.Fields("Lang").Value) = False Then
           strLanguage = adocheck.Fields("Lang").Value
        Else
           strLanguage = "2"
        End If
    Else
        strLanguage = "2"
    End If
    adocheck.Close
    Printer.CurrentX = 7000 + m_dblLeft
    Printer.CurrentY = 0 + intRow * 300
'    Printer.Print Format(AFDate(ServerDate), "mmm. d, yyyy")
    intRow = intRow + 1
    adocheck.CursorLocation = adUseClient
    'Modify By Cheng 2004/02/27
'    strSQLA = "Select * From Fagent Where FA01='" & Mid(strFANo, 1, 8) & "' And FA02='" & Mid(strFANo, 9, 1) & "' "
    StrSQLa = "Select FA04, FA05, FA63, FA64, FA65, FA06, FA17, FA18, FA19, FA20, FA21, FA22, FA70, FA32, FA33, FA34, FA35, FA36, FA23 From Fagent Where FA01='" & Mid(strFaNo, 1, 8) & "' And FA02='" & Mid(strFaNo, 9, 1) & "' "
    StrSQLa = StrSQLa & " Union Select CU04 As FA04, CU05 As FA05, CU88 As FA63, CU89 As FA64, CU90 As FA65, CU06 As FA06, CU23 As FA17, CU24 As FA18, CU25 As FA19, CU26 As FA20, CU27 As FA21, CU28 As FA22, Cu102 As FA70, CU65 As FA32, CU66 As FA33, CU67 As FA34, CU68 As FA35, CU69 As FA36, CU29 As FA23 From Customer Where CU01='" & Mid(strFaNo, 1, 8) & "' And CU02='" & Mid(strFaNo, 9, 1) & "' "
    'End
    adocheck.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
    If adocheck.RecordCount > 0 Then
        Select Case strLanguage
           Case "2"
              If IsNull(adocheck.Fields("fa05").Value) = False Then
                 If m_intPage = 1 Then
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa05").Value
                 End If
              Else '無英文印中文
                 If m_intPage = 1 Then
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa04").Value
                 End If
              End If
              If IsNull(adocheck.Fields("fa63").Value) = False Then
                 If m_intPage = 1 Then
                    intRow = intRow + 1
                    intCounter = intCounter + 1
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa63").Value
                 End If
              End If
              If IsNull(adocheck.Fields("fa64").Value) = False Then
                 If m_intPage = 1 Then
                    intRow = intRow + 1
                    intCounter = intCounter + 1
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa64").Value
                 End If
              End If
              If IsNull(adocheck.Fields("fa65").Value) = False Then
                 If m_intPage = 1 Then
                    intRow = intRow + 1
                    intCounter = intCounter + 1
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa65").Value
                 End If
              End If
              intRow = intRow + 1
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa32").Value) Then
                    If IsNull(adocheck.Fields("fa18").Value) = False Then
                       Printer.CurrentX = 0 + m_dblLeft
                       Printer.CurrentY = 0 + intRow * 300
                       Printer.Print adocheck.Fields("fa18").Value
                    'Add By Cheng 2003/03/26
                    '若無英文地址時,  印中文地址
                    Else
                       Printer.CurrentX = 0 + m_dblLeft
                       Printer.CurrentY = 0 + intRow * 300
                       Printer.Print "" & adocheck.Fields("fa17").Value
                    End If
                 Else
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa32").Value
                 End If
              End If
              intRow = intRow + 1
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa32").Value) Then
                    If IsNull(adocheck.Fields("fa19").Value) = False Then
                       Printer.CurrentX = 0 + m_dblLeft
                       Printer.CurrentY = 0 + intRow * 300
                       Printer.Print adocheck.Fields("fa19").Value
                    End If
                 Else
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print "" & adocheck.Fields("fa33").Value
                 End If
              End If
              intRow = intRow + 1
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa32").Value) Then
                    If IsNull(adocheck.Fields("fa20").Value) = False Then
                       Printer.CurrentX = 0 + m_dblLeft
                       Printer.CurrentY = 0 + intRow * 300
                       Printer.Print adocheck.Fields("fa20").Value
                    End If
                 Else
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print "" & adocheck.Fields("fa34").Value
                 End If
              End If
              intRow = intRow + 1
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa32").Value) Then
                    If IsNull(adocheck.Fields("fa21").Value) = False Then
                       Printer.CurrentX = 0 + m_dblLeft
                       Printer.CurrentY = 0 + intRow * 300
                       Printer.Print adocheck.Fields("fa21").Value
                    End If
                 Else
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print "" & adocheck.Fields("fa35").Value
                 End If
              End If
              intRow = intRow + 1
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa32").Value) Then
                    If IsNull(adocheck.Fields("fa22").Value) = False Then
                       Printer.CurrentX = 0 + m_dblLeft
                       Printer.CurrentY = 0 + intRow * 300
                       Printer.Print adocheck.Fields("fa22").Value
                    End If
                 Else
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print "" & adocheck.Fields("fa36").Value
                 End If
                 
                  'Add by Morgan 2011/5/25
                  '英文地址6
                 If IsNull(adocheck.Fields("fa32").Value) Then
                    If IsNull(adocheck.Fields("fa70").Value) = False Then
                       intRow = intRow + 1
                       Printer.CurrentX = 0 + m_dblLeft
                       Printer.CurrentY = 0 + intRow * 300
                       Printer.Print adocheck.Fields("fa70").Value
                    End If
                 End If
                 
              End If
           Case "3"
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa06").Value) = False Then
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa06").Value
                 End If
              End If
              intRow = intRow + 1
              If m_intPage = 1 Then
                 If IsNull(adocheck.Fields("fa23").Value) = False Then
                    Printer.CurrentX = 0 + m_dblLeft
                    Printer.CurrentY = 0 + intRow * 300
                    Printer.Print adocheck.Fields("fa23").Value
                 End If
              End If
        End Select
    End If
    adocheck.Close
    intRow = intRow + 1
    intCounter = intRow

End Sub

'Add By Cheng 204/02/27
'檢查案件基本資料
Private Function ChkOurCase(strCaseNo As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

ChkOurCase = True
StrSQLa = "Select PA01 From Patent Where " & ChgPatent(strCaseNo)
StrSQLa = StrSQLa & " Union Select TM01 From Trademark Where " & ChgTradeMark(strCaseNo)
StrSQLa = StrSQLa & " Union Select LC01 From Lawcase Where " & ChgLawcase(strCaseNo)
StrSQLa = StrSQLa & " Union Select HC01 From Hirecase Where " & ChgHirecase(strCaseNo)
StrSQLa = StrSQLa & " Union Select SP01 From Servicepractice Where " & ChgService(strCaseNo)
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
If rsA.EOF = True Then
    ChkOurCase = False
    MsgBox "查無此本所案號資料, 請重新輸入!!!", vbExclamation + vbOKOnly
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'Add By Cheng 204/02/27
'檢查客戶或代理人基本資料
Private Function ChkOurCust(strCustNumber As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

ChkOurCust = True
StrSQLa = "Select CU01 From Customer Where " & ChgCustomer(strCustNumber)
StrSQLa = StrSQLa & " Union Select FA01 From Fagent Where " & ChgFagent(strCustNumber)
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
If rsA.EOF = True Then
    ChkOurCust = False
    MsgBox "查無此客戶或代理人資料, 請重新輸入!!!", vbExclamation + vbOKOnly
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'Add By Cheng 204/02/27
'檢查客戶或代理人基本資料
Private Function ChkOurStaff(strStaffNumber As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

ChkOurStaff = True
StrSQLa = "Select ST01 From Staff Where ST01='" & strStaffNumber & "' And ST04='1' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
If rsA.EOF = True Then
    ChkOurStaff = False
    MsgBox "查無此智權人員資料, 請重新輸入!!!", vbExclamation + vbOKOnly
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'Add by Amy 2014/11/05 由aacc_sav搬回
Public Sub Frmacc21f1_Save()
Dim strMsg As String 'Add by Amy 2014/11/06

On Error GoTo Checking
   With Frmacc21f1
      If .Text1 = MsgText(601) Then
         MsgBox MsgText(10) & .Label1, , MsgText(5)
         strControlButton = MsgText(602)
         .Text1.SetFocus
         Exit Sub
      Else
         If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
             MsgBox .Label3 & MsgText(52), , MsgText(5)
             strControlButton = MsgText(602)
            'Add by Amy 2014/11/06 +if 避免有有傳票資料又進錯畫面按insert的錯誤
            If .MaskEdBox1.Enabled = True Then .MaskEdBox1.SetFocus
            Exit Sub
         Else
            If DateCheck(.MaskEdBox1.Text) = MsgText(603) Then
               MsgBox .Label3 & MsgText(63), , MsgText(5)
               strControlButton = MsgText(602)
               .MaskEdBox1.SetFocus
               Exit Sub
            End If
            'Add by Amy 2014/11/06 +系統日檢查
            If .MaskEdBox1.Enabled = True Then
                If ChkWorkData("1", DBDATE(.MaskEdBox1), strMsg) = False Then
                    MsgBox .Label3 & strMsg, , MsgText(5)
                    .MaskEdBox1.SetFocus
                    Exit Sub
                End If
            End If
         End If
      End If
      If .adoacc1i0.RecordCount = 0 Then
         .adoacc1i0.AddNew
         .adoacc1i0.Fields("a1i01").Value = strItemNo
         .adoacc1i0.Fields("a1i09").Value = Val(strSrvDate(2))
         .adoacc1i0.Fields("a1i10").Value = ServerTime
         .adoacc1i0.Fields("a1i11").Value = strUserNum
      Else
         .adoacc1i0.Find "a1i01 = '" & strItemNo & "'", 0, adSearchForward, 1
         If .adoacc1i0.EOF Then
            .adoacc1i0.AddNew
            .adoacc1i0.Fields("a1i01").Value = strItemNo
            .adoacc1i0.Fields("a1i09").Value = Val(strSrvDate(2))
            .adoacc1i0.Fields("a1i10").Value = ServerTime
            .adoacc1i0.Fields("a1i11").Value = strUserNum
         End If
      End If
      '.adoacc1i0.Fields("a1i02").Value = .Text1
      If .Text2 <> MsgText(601) Then
         .adoacc1i0.Fields("a1i04").Value = Val(.Text2)
      Else
         .adoacc1i0.Fields("a1i04").Value = 0
      End If
      If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
         .adoacc1i0.Fields("a1i03").Value = Val(FCDate(.MaskEdBox1.Text))
      Else
         .adoacc1i0.Fields("a1i03").Value = Null
      End If
      If .Combo3 <> MsgText(601) Then
         .adoacc1i0.Fields("a1i05").Value = .Combo3
      Else
         .adoacc1i0.Fields("a1i05").Value = Null
      End If
      If .Text4 <> MsgText(601) Then
         .adoacc1i0.Fields("a1i06").Value = Val(.Text4)
      Else
         .adoacc1i0.Fields("a1i06").Value = 0
      End If
      If .Text5 <> MsgText(601) Then
         .adoacc1i0.Fields("a1i07").Value = Val(.Text5)
      Else
         .adoacc1i0.Fields("a1i07").Value = 0
      End If
      If .Combo1 <> MsgText(601) Then
         .adoacc1i0.Fields("a1i15").Value = Mid(.Combo1, 1, 1)
      Else
         .adoacc1i0.Fields("a1i15").Value = Null
      End If
      If strCon1 <> MsgText(601) Then
         .adoacc1i0.Fields("a1i08").Value = strCon1
      Else
         .adoacc1i0.Fields("a1i08").Value = Null
      End If
      .adoacc1i0.Fields("a1i12").Value = Val(strSrvDate(2))
      .adoacc1i0.Fields("a1i13").Value = ServerTime
      .adoacc1i0.Fields("a1i14").Value = strUserNum
      .adoacc1i0.UpdateBatch
      Select Case strCon1
         Case "Y"
            adoTaie.Execute "update acc1k0 set a1k29 = 'Y' where a1k17 = '" & strItemNo & "'"
         Case Else
            adoTaie.Execute "update acc1k0 set a1k29 = null where a1k17 = '" & strItemNo & "'"
      End Select
      adoTaie.Execute "update acc1i0 set a1i02 = '" & .Text1 & "' where a1i01 = '" & strItemNo & "'"
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

