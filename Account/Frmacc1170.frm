VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc1170 
   AutoRedraw      =   -1  'True
   Caption         =   "應付款資料"
   ClientHeight    =   5184
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5184
   ScaleWidth      =   8760
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
      Height          =   550
      Left            =   7390
      Picture         =   "Frmacc1170.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   51
      ToolTipText     =   "清除畫面"
      Top             =   3000
      Width           =   550
   End
   Begin VB.TextBox Text15 
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
      Left            =   1320
      TabIndex        =   50
      Top             =   4600
      Width           =   1572
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "明細"
      Height          =   300
      Left            =   2784
      TabIndex        =   49
      Top             =   1330
      Width           =   600
   End
   Begin VB.TextBox Text21 
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
      Left            =   7080
      TabIndex        =   48
      Top             =   600
      Width           =   1200
   End
   Begin VB.TextBox Text19 
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
      Left            =   3840
      MaxLength       =   1
      TabIndex        =   2
      Top             =   240
      Width           =   450
   End
   Begin VB.TextBox Text17 
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
      Left            =   6650
      MaxLength       =   8
      TabIndex        =   16
      Top             =   4365
      Width           =   800
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
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   15
      Top             =   4365
      Width           =   1572
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
      Height          =   300
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   14
      Top             =   4320
      Width           =   1572
   End
   Begin VB.TextBox Text20 
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
      Height          =   300
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   40
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   2670
      Picture         =   "Frmacc1170.frx":08CA
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   240
      Width           =   350
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   1572
   End
   Begin VB.TextBox Text14 
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
      Left            =   5415
      MaxLength       =   1
      TabIndex        =   3
      Top             =   240
      Width           =   612
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
      Height          =   550
      Left            =   7990
      Picture         =   "Frmacc1170.frx":09CC
      Style           =   1  '圖片外觀
      TabIndex        =   18
      ToolTipText     =   "取消"
      Top             =   3000
      Width           =   550
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
      Height          =   300
      Left            =   4920
      TabIndex        =   34
      Top             =   4032
      Width           =   3492
   End
   Begin VB.TextBox Text12 
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
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   13
      Top             =   4032
      Width           =   852
   End
   Begin VB.TextBox Text11 
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
      Height          =   300
      Left            =   1320
      MaxLength       =   14
      TabIndex        =   12
      Top             =   4032
      Width           =   1572
   End
   Begin VB.TextBox Text9 
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
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   11
      Top             =   3720
      Width           =   1572
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
      Height          =   300
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   10
      Top             =   3720
      Width           =   612
   End
   Begin VB.TextBox Text7 
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
      Height          =   300
      Left            =   5916
      TabIndex        =   28
      Top             =   3120
      Width           =   1416
   End
   Begin VB.TextBox Text6 
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
      Height          =   300
      Left            =   4452
      TabIndex        =   27
      Top             =   3120
      Width           =   1440
   End
   Begin VB.TextBox Text4 
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
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   8
      Top             =   1320
      Width           =   1300
   End
   Begin VB.TextBox Text2 
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
      Left            =   1440
      MaxLength       =   9
      TabIndex        =   4
      Top             =   600
      Width           =   1200
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
      Height          =   315
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   4080
      TabIndex        =   6
      Top             =   960
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   6720
      TabIndex        =   7
      Top             =   960
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
      Bindings        =   "Frmacc1170.frx":1036
      Height          =   1300
      Left            =   240
      TabIndex        =   39
      Top             =   1680
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   2307
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
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
      ColumnCount     =   5
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
            Format          =   "#,##0"
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
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
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
      BeginProperty Column04 
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
         Size            =   254
         BeginProperty Column00 
            ColumnWidth     =   3971.906
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1428.095
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   5532.095
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   1560
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
   Begin MSForms.TextBox Text22 
      Height          =   315
      Left            =   7470
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   4360
      Width           =   960
      VariousPropertyBits=   679493659
      BackColor       =   14737632
      MaxLength       =   8
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text10 
      Height          =   300
      Left            =   5640
      TabIndex        =   31
      Top             =   3720
      Width           =   2770
      VariousPropertyBits=   679493657
      BackColor       =   14737632
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   335
      Left            =   4080
      TabIndex        =   17
      Top             =   4680
      Width           =   4335
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
   Begin MSForms.TextBox Text5 
      Height          =   315
      Left            =   4080
      TabIndex        =   9
      Top             =   1320
      Width           =   4212
      VariousPropertyBits=   -1466941413
      MaxLength       =   35
      ScrollBars      =   2
      Size            =   "7429;556"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text3 
      Height          =   315
      Left            =   2700
      TabIndex        =   21
      Top             =   600
      Width           =   3235
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label21 
      BackStyle       =   0  '透明
      Caption         =   "統一編號"
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
      Left            =   6120
      TabIndex        =   47
      Top             =   630
      Width           =   975
   End
   Begin VB.Label Label19 
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
      Height          =   255
      Left            =   3120
      TabIndex        =   46
      Top             =   270
      Width           =   780
   End
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "對沖(業)"
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
      Left            =   5775
      TabIndex        =   45
      Top             =   4365
      Width           =   975
   End
   Begin VB.Label Label20 
      BackStyle       =   0  '透明
      Caption         =   "對沖(其)"
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
      Left            =   3120
      TabIndex        =   44
      Top             =   4365
      Width           =   975
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "對沖(本)"
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
      Left            =   360
      TabIndex        =   43
      Top             =   4350
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "對沖(客)"
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
      Left            =   360
      TabIndex        =   42
      Top             =   4650
      Width           =   1020
   End
   Begin VB.Label Label22 
      BackStyle       =   0  '透明
      Caption         =   "筆數"
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
      TabIndex        =   41
      Top             =   3120
      Width           =   852
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "應付款類別"
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
      Left            =   240
      TabIndex        =   38
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label15 
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
      Left            =   3135
      TabIndex        =   37
      Top             =   4650
      Width           =   855
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "(1:廠商 2:客戶 3:員工)"
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
      Left            =   6120
      TabIndex        =   36
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "往來類別"
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
      Left            =   4440
      TabIndex        =   35
      Top             =   240
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1440
      Left            =   252
      Top             =   3588
      Width           =   8292
   End
   Begin VB.Label Label12 
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
      Left            =   3120
      TabIndex        =   33
      Top             =   4032
      Width           =   972
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "金額"
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
      TabIndex        =   32
      Top             =   4032
      Width           =   852
   End
   Begin VB.Label Label10 
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
      Left            =   3120
      TabIndex        =   30
      Top             =   3720
      Width           =   972
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "借1/貸2"
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
      TabIndex        =   29
      Top             =   3720
      Width           =   852
   End
   Begin VB.Label Label7 
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
      Left            =   3720
      TabIndex        =   26
      Top             =   3120
      Width           =   612
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "備註"
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
      Left            =   3590
      TabIndex        =   25
      Top             =   1320
      Width           =   550
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "欲處理日期"
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
      TabIndex        =   24
      Top             =   960
      Width           =   1212
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "入帳日期"
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
      Left            =   3120
      TabIndex        =   23
      Top             =   960
      Width           =   972
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "發票號碼"
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
      TabIndex        =   22
      Top             =   1320
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "往來對象"
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
      Left            =   240
      TabIndex        =   20
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "應付款單號"
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
      Left            =   240
      TabIndex        =   19
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Frmacc1170"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/08/20 Form2.0已修改 Text3/Text5/Text10/Text22/Combo1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit
Public adoacc0o0 As New ADODB.Recordset
Public adoacc1p0 As New ADODB.Recordset
Public adoacc010 As New ADODB.Recordset
Public adoacc090 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adocase As New ADODB.Recordset
Dim strSerialNo As String
Public strDocNo As String
Public strAccNo As String
'add by nick 2004/07/06
Public adocase1 As New ADODB.Recordset
'Add by Amy 2013/12/26
Dim strFAId As String '統一編號
Dim bolFirstIns As Boolean 'Add by Amy 2014/10/27 是否按過insert
Dim bolDelAll As Boolean '修改時是否將資料全部剪掉
Public bolBack As Boolean 'Add by Amy 2018/09/13

'Add by Amy 2013/12/26
Private Sub cmdDetail_Click()
    'Add by Amy 2025/02/26 避免未輸入帳日及欲處理日進入會Error 彈訊息
    If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
         MsgBox Label4 & "不可為空！"
         Exit Sub
    End If
    If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
         MsgBox Label5 & "不可為空！"
         Exit Sub
    End If
    'end 2025/02/26
    If strSaveConfirm = MsgText(3) Then
        Frmacc1170_Save
    End If
    strItemNo = Text1     '記錄應付款單號返回才會重Acc0o0Refresh
    'Add by Amy 2014/10/27
    If bolFirstIns = False Then
        Text19.Tag = Text19 '記錄公司別以免返回後修改
    End If
    'end 2014/10/27
    Frmacc1172.strBackForm = "Frmacc1170" '記錄返回畫面
    Frmacc1172.Show
    Screen.MousePointer = vbDefault
    Me.Hide
End Sub
'end 2013/12/26

Private Sub Combo1_Change()
   If Mid(Text9, 1, 1) = "2" Then
      Text5 = Combo1
   End If
End Sub

Private Sub Combo1_GotFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   OpenIme
End Sub

Private Sub Combo1_LostFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   CloseIme
End Sub

Private Sub Combo2_Change()
    If strSaveConfirm = MsgText(3) Then
        If Text19 = "J" And Text14 = "1" And Mid(Combo2, 1, 1) = "1" Then
            cmdDetail.Enabled = True
        Else
            cmdDetail.Enabled = False
        End If
    End If
End Sub

'Add by Amy 2014/01/07
Private Sub Command1_Click()
    AdodcClear
    strSerialNo = MsgText(601)
    AdodcRefresh
End Sub
'end 2014/01/07

Private Sub Command2_Click()
   If Adodc1.Recordset.RecordCount <> 0 Then
      'Modify by Amy 2016/07/01 若Grid忘了選再按剪下程式會錯
      'If IsNull(Adodc1.Recordset.Fields("a1p22").Value) = False Then
      If strAccNo <> MsgText(601) Then
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
End Sub

Private Sub Command3_Click()
   'If adoacc0o0.RecordCount = 0 Or Text1 = MsgText(601) Then
   '   Exit Sub
   'End If
   'adoacc0o0.Find "a0o01 = '" & Text1 & "'", 0, adSearchForward, 1
   'If adoacc0o0.EOF = False Then
   '   FormShow
   '   AdodcRefresh
   '   RecordShow
   'Else
   '   MsgBox MsgText(33), , MsgText(5)
   '   adoacc0o0.MoveFirst
   'End If
   Acc0o0Refresh
   If adoacc0o0.RecordCount <> 0 Then
      FormShow
      AdodcRefresh
      SumShow
      RecordShow
      FormDisabled 'Add by Amy 2013/12/26
   End If
   AdodcClear 'Add by Amy 2014/01/07
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
   
   'If adoacc0o0.RecordCount <> 0 Then
   '   adoacc0o0.MoveFirst
   'End If
   'adoacc0o0.Find "a0o01 = '" & strItemNo & "'", 0, adSearchForward, 1
   'If adoacc0o0.EOF = False Then
   '   FormShow
   '   AdodcRefresh
   '   SumShow
   '   RecordShow
   'End If
   
   Text1 = strItemNo
   Acc0o0Refresh
   If adoacc0o0.RecordCount <> 0 Then
       FormShow
       AdodcRefresh
       SumShow
       RecordShow
       'Add by Amy 2014/01/07 +if
       If strSaveConfirm = MsgText(601) Then
           '由查詢或frmacc1172返回且strSaveConfirm=""時 (與frmacc4120共用1172)
            FormDisabled
       Else
            FormEnabled
       End If
   End If
   strItemNo = MsgText(601)
   
End Sub
'Add by Morgan 2004/9/27
'iAct:0=修改,1=刪除
Public Function EditCheck(Optional iAct As Integer = 0) As Boolean
'Add by Amy 2013/12/26 +if 因一進入就按修改造成錯誤
If adoacc0o0.RecordCount > 0 Then
   If "" & adoacc0o0.Fields("a0o11").Value <> "" Then
      MsgBox "此筆資料已付款，不可修改(刪除)！"
   Else
      'Add by Morgan 2008/1/2
      '2010/9/24 MODIFY BY SONIA G09901020
      'strExc(0) = "select a1p22 from acc1p0 where a1p04='" & Text1 & "'"
      strExc(0) = "select a1p22 from acc1p0 where a1p04='" & Text1 & "' AND A1P02='B'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If Not IsNull(RsTemp.Fields(0)) Then
            If iAct = 1 Then
               MsgBox "此筆資料已轉傳票，不可刪除！"
               Exit Function
            End If
            If PUB_CheckPosted(RsTemp.Fields(0)) = True Then
               Exit Function
            End If
         End If
      End If
      'end 2008/1/2
      EditCheck = True
   End If
End If
'end 2013/12/26
End Function

'Add by Amy 2021/08/20 避免於Form2.0 物件上按Insert 觸發2次 KeyDefine KeyCode事件
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Form2.0 記錄鍵盤傳入順序
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   'Add by Morgan 2004/9/27
   If KeyCode = vbKeyF3 Then
      If EditCheck = False Then Exit Sub
   End If
   '2004/9/27 end
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 8976 'Moidy by Amy 2023/07/19  原:8850
   Me.Height = 5750 'Moidy by Amy 2023/07/19  原:5500
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   strItemNo = MsgText(601)
   Combo2.AddItem ComboItem(101)
   Combo2.AddItem ComboItem(102)
   Combo2.AddItem ComboItem(103)
   OpenTable
   If adoacc0o0.RecordCount <> 0 Then
      adoacc0o0.MoveLast
      adoacc0o0.MoveFirst
      RecordShow
   End If
   FormDisabled
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   'Modify  by Amy 2014/01/15
'   CreDebCheck
'   If CreDebCheck <> MsgText(602) Then
'      tool1_enabled
'      MsgBox MsgText(11), , MsgText(5)
'      Cancel = True
'      Exit Sub
'   End If
    If FormCheck = False Then
        tool1_enabled
        Cancel = True
        Exit Sub
    End If
    'end 2013/01/15
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   strTrackMode = "" 'Add by Amy 2021/08/20 Form2.0 記錄鍵盤傳入順序(清除)
   MenuEnabled
   Set Frmacc1170 = Nothing
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
  
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label4 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If strSaveConfirm <> MsgText(3) Then
      Exit Sub
   End If
   If Mid(MaskEdBox1.Text, 1, 3) <> Mid(CFDate(ACDate(ServerDate)), 1, 3) Then
      Text1 = UpdateNo("acc0o0", "a0o01", 5, MaskEdBox1.Text, MsgText(818))
   Else
      'Text1 = AutoNo(MsgText(804), 5)
      Text1 = strDocNo
   End If
   
   'Add by Morgan 2005/1/15 若往來類別為廠商時，欲處理日期預設次月10日
   If Text14.Text = "1" Then
      MaskEdBox2.Mask = ""
      'Modify by Morgan 2008/3/3
      'MaskEdBox2.Text = Format(Format(DateAdd("M", 1, MaskEdBox1.Text), "YYMM"), "0##/##") & "/10"
      'Modify by Morgan 2010/12/6
      'MaskEdBox2.Text = Format(Format(DateAdd("M", 1, Left(MaskEdBox1.Text, 7) & "01"), "YYMM"), "0##/##") & "/10"
      MaskEdBox2.Text = Format(TransDate(CompDate(1, 1, Trim(DBDATE(MaskEdBox1))), 1) \ 100, "0##/##") & "/10"
      'end 2008/3/3
      MaskEdBox2.Mask = DFormat
   End If
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc0o0.CursorLocation = adUseClient
   adoacc0o0.MaxRecords = intMax
   adoacc0o0.Open "select * from acc0o0 where a0o01 >= '" & Text1 & "' order by a0o01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc1p0.CursorLocation = adUseClient
   'Modify by Amy 2013/12/26 改公司別 原:'1'
   adoacc1p0.Open "select * from acc1p0 where  a1p01 = '" & Text19 & "' and a1p02 = 'B' and a1p03 = '" & Text1 & "' and a1p04 = 'B' order by a1p05 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc1p0 where  a1p01 = '" & Text19 & "' and a1p02 = 'B' and a1p03 = '" & Text1 & "' and a1p04 = 'B' order by a1p05 asc", adoTaie, adOpenStatic, adLockReadOnly
   'end 2013/12/13
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表(國內應付帳款資料(主檔))
'
'*************************************************
Public Sub FormShow()
   Text1 = adoacc0o0.Fields("a0o01").Value
   Text19 = adoacc0o0.Fields("a0o07").Value 'Add by Amy 2013/12/26
   If IsNull(adoacc0o0.Fields("a0o02").Value) Then
      Text14 = MsgText(601)
   Else
      Text14 = adoacc0o0.Fields("a0o02").Value
   End If
   If IsNull(adoacc0o0.Fields("a0o03").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = adoacc0o0.Fields("a0o03").Value
   End If
   Text2.Tag = Text2 'Add by Amy 2014/10/27
   If IsNull(adoacc0o0.Fields("a0o19").Value) Then
      Combo2 = MsgText(601)
   Else
      Combo2 = Combo2.List(Val(adoacc0o0.Fields("a0o19").Value) - 1)
   End If
   'Modify by Amy 2018/09/13 +if (frmacc1172返回不要再帶a0o04值,以明細發票號碼為主)
   If bolBack = False Then
        If IsNull(adoacc0o0.Fields("a0o04").Value) Then
           Text4 = MsgText(601)
        Else
           Text4 = adoacc0o0.Fields("a0o04").Value
        End If
   End If
   'end 2018/09/13
   Text4.Tag = Text4 'Add by Amy 2018/09/13
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc0o0.Fields("a0o05").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(adoacc0o0.Fields("a0o05").Value)
   End If
   MaskEdBox1.Tag = "" & adoacc0o0.Fields("a0o05").Value 'Add by Amy 2014/10/27
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(adoacc0o0.Fields("a0o06").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(adoacc0o0.Fields("a0o06").Value)
   End If
   MaskEdBox2.Mask = DFormat
   If IsNull(adoacc0o0.Fields("a0o10").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = adoacc0o0.Fields("a0o10").Value
   End If
   Text21 = "" 'Add by Amy 2013/12/26 +統一編號 欄
   Select Case Text14
      Case Mid(ComboItem(91), 1, 1)
         'Modify by Amy 2013/12/26 往來類別為1-廠商抓其名稱及統一編號
         'Text3 = A0i02Query(Text2)
'         If strSaveConfirm = MsgText(601) Then
            Text3 = GetFactoryName(Text2, strFAId)
'         Else
'            Text3 = GetFactoryName(Text2, strFAId, True)
'         End If
         Text21 = strFAId
      Case Mid(ComboItem(92), 1, 1)
         If Len(Text2) = 6 Then
            Text2 = AfterZero(Text2)
         'Add by Morgan 2007/3/1 八碼時要補'0'
         ElseIf Len(Text2) = 8 Then
            Text2 = Text2 & "0"
         'End 2007/3/1
         End If
         Text3 = CustomerQuery(Text2, 1)
         
      Case Mid(ComboItem(93), 1, 1)
         Text3 = StaffQuery(Text2)
      Case Else
         Text3 = MsgText(601)
   End Select
   
End Sub

'*************************************************
'  顯示資料表(國內應付帳款資料(分錄檔))
'
'*************************************************
Public Sub AdodcShow()
   Text9 = Adodc1.Recordset.Fields("a1p05").Value
   If IsNull(Adodc1.Recordset.Fields("a1p07").Value) Or Adodc1.Recordset.Fields("a1p07").Value = 0 Then
      If IsNull(Adodc1.Recordset.Fields("a1p08").Value) Then
         Text8 = MsgText(601)
         Text11 = MsgText(601)
      Else
         Text8 = "2"
         Text11 = Adodc1.Recordset.Fields("a1p08").Value
      End If
   Else
      Text8 = "1"
      Text11 = Adodc1.Recordset.Fields("a1p07").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p06").Value) Then
      Text12 = MsgText(601)
   Else
      Text12 = Adodc1.Recordset.Fields("a1p06").Value
   End If
   'Modify by Amy 2014/01/07 取消作帳公司改顯示a1p15
   'If IsNull(Adodc1.Recordset.Fields("a1p31").Value) Then
   If IsNull(Adodc1.Recordset.Fields("a1p15").Value) Then
      Text15 = MsgText(601)
   Else
      'Text15 = Adodc1.Recordset.Fields("a1p31").Value
      Text15 = Adodc1.Recordset.Fields("a1p15").Value
   End If
   'end 2014/01/07
   If IsNull(Adodc1.Recordset.Fields("a1p14").Value) Then
      Combo1 = MsgText(601)
   Else
      Combo1 = Adodc1.Recordset.Fields("a1p14").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p17").Value) Then
      Text16 = MsgText(601)
   Else
      Text16 = Adodc1.Recordset.Fields("a1p17").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p30").Value) Then
      Text18 = MsgText(601)
   Else
      Text18 = Adodc1.Recordset.Fields("a1p30").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p16").Value) Then
      Text17 = MsgText(601)
   Else
      Text17 = Adodc1.Recordset.Fields("a1p16").Value
   End If
End Sub

'*************************************************
'  計算並顯示借/貸方總金額
'
'*************************************************
Public Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   'Modify by Amy 2013/12/26 +公司別 原:'1'
   adoaccsum.Open "select sum(a1p07), sum(a1p08), count(*) from acc1p0 where a1p01 = '" & Text19 & "' and a1p02 = 'B' and a1p04 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text6 = MsgText(601)
      Else
         Text6 = Format(adoaccsum.Fields(0).Value, DDollar)
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text7 = MsgText(601)
      Else
         Text7 = Format(adoaccsum.Fields(1).Value, DDollar)
      End If
      If IsNull(adoaccsum.Fields(2).Value) Then
         Text20 = MsgText(601)
      Else
         Text20 = adoaccsum.Fields(2).Value
      End If
   Else
      Text6 = MsgText(601)
      Text7 = MsgText(601)
      Text20 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  清除顯示資料(藍色框部分)
'
'*************************************************
Public Sub AdodcClear()
   Text8 = ""
   Text9 = ""
   Text10 = ""
   Text11 = ""
   Text12 = ""
   Text15 = ""
   Combo1 = ""
   Text16 = ""
   Text18 = ""
   Text17 = ""
   Text22 = "" 'Add by Amy 2017/12/05
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Dim bolCancel As Boolean 'Add by Amy 2020/11/06
   Call PUB_SaveTrackMode(1, KeyCode)  'Add by Amy 2021/08/20 Form2.0 記錄鍵盤傳入順序
   
   Select Case KeyCode
      Case vbKeyInsert
         If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
            Exit Sub
         End If
         
         'Add by Amy 2021/08/20 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
         If PUB_ChkTrackMode = False Then
             Exit Sub
         End If

         'Add by Amy 2020/11/06 L公司6及72字頭部門別不可輸TOT或空值
         Call Text12_Validate(bolCancel)
         If bolCancel = True Then
            TextInverse Text12
            Text12.SetFocus
            Exit Sub
         End If
         'end 2020/11/06
         Frmacc1170_Save
         If strControlButton <> MsgText(602) Then
            Acc1p0Save
            'Add by Amy 2014/10/27 +新增-是否按過insert (for 最後更新acc1p0)
            If strSaveConfirm = MsgText(3) And bolFirstIns = False And Text19.Tag = "" Then
                Text19.Tag = Text19
                bolFirstIns = True
            End If
            'end 2014/10/27
         End If
         If strControlButton <> MsgText(602) Then
            SumShow
            AdodcClear
            Text8.SetFocus
         End If
         strControlButton = MsgText(601)
   End Select
   KeyEnter KeyCode
   
End Sub

'*************************************************
'  儲存資料表(國內應付帳款資料(分錄檔))
'
'*************************************************
Private Sub Acc1p0Save()
    Dim strSql As String 'Add by Amy 2016/07/01
    
On Error GoTo Checking
   If Text9 = MsgText(601) Then
      MsgBox MsgText(10) & Label10, , MsgText(5)
      strControlButton = MsgText(602)
      Text9.SetFocus
      Exit Sub
   Else
      'Modify by Amy 2014/01/07 +公司別確認
'      If ExistCheck("acc010", "a0101", Text9, Label10) = False Then
'         strControlButton = MsgText(602)
'         Text9.SetFocus
'         Exit Sub
'      End If
      If PUB_CheckCompany(Text9, Text19) = False Then
         strControlButton = MsgText(602)
         Text9.SetFocus
         Exit Sub
      End If
      'end 2014/01/07
      If CheckDept(Text9, Text12) = False Then
         MsgBox MsgText(103), , MsgText(5)
         strControlButton = MsgText(602)
         Text12.SetFocus
         Exit Sub
      End If
      If Text11 = MsgText(601) Or Val(Text11) = 0 Then
         MsgBox MsgText(58), , MsgText(5)
         strControlButton = MsgText(602)
         Text11.SetFocus
         Exit Sub
      End If
      If Text12 <> MsgText(601) Then
         If ExistCheck("acc090", "a0901", Text12, Label12) = False Then
            strControlButton = MsgText(602)
            Text12.SetFocus
            Exit Sub
         End If
      End If
      If Text16 <> MsgText(601) Then
         Text16 = CaseNoZero(Text16)
         adocase.CursorLocation = adUseClient
         adocase.Open "select pa01 as SystemNo from patent where pa01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and pa02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and pa03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and pa04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
                     "select tm01 as SystemNo from trademark where tm01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and tm02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and tm03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and tm04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
                     "select lc01 as SystemNo from lawcase where lc01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and lc02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and lc03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and lc04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
                     "select hc01 as SystemNo from hirecase where hc01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and hc02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and hc03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and hc04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
                     "select sp01 as SystemNo from servicepractice where sp01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and sp02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and sp03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and sp04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adocase.RecordCount = 0 Then
            MsgBox MsgText(28) & Label17, , MsgText(5)
            strControlButton = MsgText(602)
            adocase.Close
            Exit Sub
         End If
         adocase.Close
      End If
      If Text17 <> MsgText(601) Then
         If ExistCheck("staff", "st01", Text17, Label18) = False Then
            strControlButton = MsgText(602)
            Exit Sub
         End If
      End If
   End If
      
   'add by sonia 2015/12/30 檢查民國105年起法務收入科目不可使用
   intI = PUB_AccNoEnable(Text9, Val(FCDate(MaskEdBox1.Text)))
   If intI <> 0 Then
      strControlButton = MsgText(602)
      Text9.SetFocus
      Exit Sub
   End If
   'end 2015/12/30
   'Add by Morgan 2007/10/2 檢查科目部門&智權人員是否正確
   intI = PUB_AccNoGood(Text9, Text12, Text17)
   If intI <> 0 Then
      strControlButton = MsgText(602)
      If intI = 1 Then
         Text9.SetFocus
      ElseIf intI = 2 Then
         Text12.SetFocus
      ElseIf intI = 3 Then
         Text17.SetFocus
      End If
      Exit Sub
   End If
   'end 2007/10/2
   'Add by Amy 2013/12/26 +if 避免Adodc1 Find a1p03 未找到資料產生錯誤
   If strSaveConfirm <> MsgText(3) Then
        If Adodc1.Recordset.RecordCount <> 0 Then
            If Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst 'Add by Amy 2014/01/07 '避免修改直接新增的錯誤
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
   End If
   'end 2013/12/26
   adoacc1p0.Close
   adoacc1p0.CursorLocation = adUseClient
   'Modify by Amy 2013/12/26 +公司別 原:'1'
   adoacc1p0.Open "select * from acc1p0 where a1p01 = '" & Text19 & "' and a1p02 = 'B' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modify by Amy 2016/07/01 +strAccNo=空 判斷 解:若修改把明細全剪掉造成acc1p0沒資料,而產生新的傳票號碼-婉莘
   'Modify by Amy 2016/07/21 未產生傳票號碼時,修改某筆會多新增一筆(拿掉 strAccNo = MsgText(601) 判斷)
   If adoacc1p0.RecordCount = 0 Then
        adoacc1p0.AddNew
        adoacc1p0.Fields("a1p01").Value = Text19 'Modify by Amy 2013/12/26 +公司別 原:"1"
        adoacc1p0.Fields("a1p02").Value = "B"
        'Modify by Amy 2013/12/26 +公司別 原:'1'
        adoacc1p0.Fields("a1p03").Value = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '" & Text19 & "' and a1p02 = 'B' and a1p04 = '" & Text1 & "'", 3)
        strSerialNo = adoacc1p0.Fields("a1p03").Value
        adoacc1p0.Fields("a1p04").Value = Text1
   End If
   adoacc1p0.Fields("a1p05").Value = Text9
   If Text8 <> MsgText(601) Then
      Select Case Text8
         Case "1"
            If Text11 <> MsgText(601) Then
               adoacc1p0.Fields("a1p07").Value = Val(Format(Text11, DAmount))
            Else
               adoacc1p0.Fields("a1p07").Value = 0
            End If
            adoacc1p0.Fields("a1p08").Value = 0
         Case "2"
            If Text11 <> MsgText(601) Then
               adoacc1p0.Fields("a1p08").Value = Val(Format(Text11, DAmount))
            Else
               adoacc1p0.Fields("a1p08").Value = 0
            End If
            adoacc1p0.Fields("a1p07").Value = 0
         Case Else
            adoacc1p0.Fields("a1p07").Value = 0
            adoacc1p0.Fields("a1p08").Value = 0
      End Select
   Else
      adoacc1p0.Fields("a1p07").Value = 0
      adoacc1p0.Fields("a1p08").Value = 0
   End If
   If Text12 <> MsgText(601) Then
      adoacc1p0.Fields("a1p06").Value = Text12
   Else
      adoacc1p0.Fields("a1p06").Value = MsgText(55)
   End If
   'Modify by Amy 2013/12/26 取消作帳公司
'   If Text15 <> MsgText(601) Then
'      adoacc1p0.Fields("a1p31").Value = Text15
'   Else
'      adoacc1p0.Fields("a1p31").Value = Null
'   End If
   If Combo1 <> MsgText(601) Then
      adoacc1p0.Fields("a1p14").Value = Combo1
      Combo1.AddItem Combo1
   Else
      adoacc1p0.Fields("a1p14").Value = ""
   End If
   If Text2 <> MsgText(601) Then
      adoacc1p0.Fields("a1p15").Value = Text2
   End If
   'modify by sonia 2021/1/29 加傳本所案號以判別FCP,FCT英日文組
   'If AccNoToSalesNo(Text9) = "" Then
   If AccNoToSalesNo(Text9, Text16) = "" Then
      adoacc1p0.Fields("a1p16").Value = Null
   Else
      'modify by sonia 2021/1/29 加傳本所案號以判別FCP,FCT英日文組
      'adoacc1p0.Fields("a1p16").Value = AccNoToSalesNo(Text9)
      adoacc1p0.Fields("a1p16").Value = AccNoToSalesNo(Text9, Text16)
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      adoacc1p0.Fields("A1P18").Value = Val(FCDate(MaskEdBox1.Text))
   Else
      adoacc1p0.Fields("A1P18").Value = ""
   End If
   If Text16 <> MsgText(601) Then
      adoacc1p0.Fields("a1p17").Value = Text16
   Else
      adoacc1p0.Fields("a1p17").Value = Null
   End If
   If Text18 <> MsgText(601) Then
      adoacc1p0.Fields("a1p30").Value = Text18
   Else
      adoacc1p0.Fields("a1p30").Value = Null
   End If
   If Text17 <> MsgText(601) Then
      adoacc1p0.Fields("a1p16").Value = Text17
   Else
      adoacc1p0.Fields("a1p16").Value = Null
   End If
   'Modify by Amy 2016/07/21 全部剪掉再Insert a1p27需Y傳票才會更新
'   If IsNull(adoacc1p0.Fields("a1p22").Value) = False Then
'      adoacc1p0.Fields("a1p27").Value = MsgText(602)
'   End If
   If strAccNo <> MsgText(601) Or IsNull(adoacc1p0.Fields("a1p22").Value) = False Then
      If bolDelAll = True Then
        adoacc1p0.Fields("a1p22").Value = strAccNo
      End If
      adoacc1p0.Fields("a1p27").Value = MsgText(602)
   End If
   'end 2016/07/21
   
   adoacc1p0.UpdateBatch
   AdodcRefresh
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   'Modify by Amy 2013/12/26 改公司別 原:'1'
   adoadodc1.Open "select * from acc1p0, acc010, acc090 where a1p05 = a0101 and a1p06 = a0901 and a1p01 = '" & Text19 & "' and a1p02 = 'B' and a1p04 = '" & Text1 & "' order by a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc1.Recordset.Fields("a1p22").Value) = False Then
         strAccNo = Adodc1.Recordset.Fields("a1p22").Value
      'Modify by Amy 2016/07/01 +If strSaveConfirm <> MsgText(4) 判斷 解:若修改把明細全剪掉造成acc1p0沒資料,而產生新的傳票號碼-婉莘
      ElseIf strSaveConfirm <> MsgText(4) Then
         strAccNo = MsgText(601)
      End If
      Adodc1.Recordset.Find "a1p03 = '" & strSerialNo & "'", 0, adSearchForward, 1
      If Adodc1.Recordset.EOF Then
         Exit Sub
      Else
         DataGrid1.SelBookmarks.add Adodc1.Recordset.Bookmark
      End If
   'Modify by Amy 2016/07/21
   Else
        If strSaveConfirm = MsgText(4) Then
            If Adodc1.Recordset.RecordCount = 0 Then bolDelAll = True
        Else
            strAccNo = MsgText(601)
        End If
   End If
   strSerialNo = MsgText(601)
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
   If Adodc1.Recordset.RecordCount <> 0 Then
      'Modify by Amy 2013/12/26 +公司別 原:'1'
      adoTaie.Execute "delete from acc1p0 where a1p01 = '" & Text19 & "' and a1p02 = 'B' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text1 & "'"
   End If
   AdodcRefresh
   SumShow
   AdodcClear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Add by Amy 2013/12/26
Private Sub MaskEdBox2_LostFocus()
    '判斷J公司發票號碼若非輸"00"時,自動切換至明細
    If Text19 = "J" And Text4 <> "00" And Text4 <> "" And Mid(Combo2, 1, 1) = "1" Then
        cmdDetail_Click
    End If
End Sub
'end 2013/12/26

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
   If MaskEdBox2.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      Exit Sub
   End If
   If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
      MsgBox Label5 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox2.SetFocus
      Exit Sub
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

Private Sub Text12_Change()
   Text13 = A0902Query(Text12)
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
   If CheckDept(Text9, Text12) = False Then
      MsgBox MsgText(103), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
   If Text12 <> MsgText(601) Then
      If ExistCheck("acc090", "a0901", Text12, Label12) = False Then
         Cancel = True
         Exit Sub
      End If
   End If
   'Add by Amy 2020/11/06 L公司6及72字頭部門別不可輸TOT或空值
   If Text19 = "L" And (Left(Text9, 1) = "6" Or Left(Text9, 2) = "72") And (Text12 = "TOT" Or Trim(Text12) = MsgText(601)) Then
      MsgBox "L公司6或72字頭會計科目之部門別不可輸TOT或空白", , MsgText(5)
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text14_Change()
    If strSaveConfirm = MsgText(3) Then
        If Text19 = "J" And Text14 = "1" And Mid(Combo2, 1, 1) = "1" Then
            cmdDetail.Enabled = True
        Else
            cmdDetail.Enabled = False
        End If
    End If
End Sub

Private Sub Text14_GotFocus()
   TextInverse Text14
End Sub

'2009/12/10 add by sonia
Private Sub Text14_KeyPress(KeyAscii As Integer)
   If (KeyAscii < Asc("1") Or KeyAscii > Asc("3")) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub
'2009/12/10 end

Private Sub Text14_Validate(Cancel As Boolean)
   'Add by Amy 2013/12/26
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   'end 2013/12/26
   Select Case Text14
      '2013/1/9 add by sonia
      Case ""
         MsgBox MsgText(10), , MsgText(5)
         Text14.SetFocus
         Cancel = True
      '2013/1/9 end
      Case Mid(ComboItem(91), 1, 1)
         If MaskEdBox1.Text = MsgText(29) Or MaskEdBox1.Text = MsgText(601) Then
            MaskEdBox1.Mask = ""
            MaskEdBox1.Text = CFDate(ACDate(ServerDate))
            MaskEdBox1.Mask = DFormat
         End If
         If MaskEdBox2.Text = MsgText(29) Or MaskEdBox2.Text = MsgText(601) Then
            MaskEdBox2.Mask = ""
            MaskEdBox2.Text = ""
            MaskEdBox2.Mask = DFormat
         End If
      Case Else
         If MaskEdBox1.Text = MsgText(29) Or MaskEdBox2.Text = MsgText(601) Then
            MaskEdBox1.Mask = ""
            MaskEdBox1.Text = CFDate(ACDate(ServerDate))
            MaskEdBox1.Mask = DFormat
         End If
         If MaskEdBox2.Text = MsgText(29) Or MaskEdBox2.Text = MsgText(601) Then
            MaskEdBox2.Mask = ""
            MaskEdBox2.Text = CFDate(ACDate(ServerDate))
            MaskEdBox2.Mask = DFormat
         End If
   End Select
End Sub

Private Sub Text16_GotFocus()
   TextInverse Text16
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text16_LostFocus()
   'add by nick 2004/07/06
On Error GoTo Checking
   If Text16 <> MsgText(601) Then
      Dim strNation As String
      Text16 = CaseNoZero(Text16)
      Set adocase1 = New ADODB.Recordset
      adocase1.CursorLocation = adUseClient
      adocase1.Open "select pa01 as SystemNo,pa09,pa26  from patent where pa01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and pa02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and pa03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and pa04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
                   "select tm01 as SystemNo,tm10,tm23 from trademark where tm01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and tm02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and tm03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and tm04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
                   "select lc01 as SystemNo,lc15,lc11 from lawcase where lc01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and lc02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and lc03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and lc04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
                   "select hc01 as SystemNo,'000',hc07 from hirecase where hc01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and hc02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and hc03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and hc04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
                   "select sp01 as SystemNo,sp09,sp08 from servicepractice where sp01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and sp02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and sp03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and sp04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocase1.RecordCount > 0 Then
      'add by nick 2004/07/07
      '檢查當科目是 220112 220102 220111 220101 220103 220104 220105 220106 時，要檢查申請國家級系統別
      strNation = CheckStr(adocase1.Fields(1).Value)
      End If
      adocase1.Close
         Select Case Text9
         Case "220101"
                 'edit by nick 2004/07/07 加系統別
                 'If (Mid(Text16, 1, Len(Text16) - 9) = "T" Or Mid(Text16, 1, Len(Text16) - 9) = "TB") And strNation = "000" Then
                 If (Mid(Text16, 1, Len(Text16) - 9) = "T" Or Mid(Text16, 1, Len(Text16) - 9) = "TB" Or Mid(Text16, 1, Len(Text16) - 9) = "TS" Or Mid(Text16, 1, Len(Text16) - 9) = "TD" Or Mid(Text16, 1, Len(Text16) - 9) = "TM" Or Mid(Text16, 1, Len(Text16) - 9) = "TR" Or Mid(Text16, 1, Len(Text16) - 9) = "TT") And strNation = "000" Then
                 Else
                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
                       Text9.SetFocus
                       Text9.SelStart = 0
                       Text9.SelLength = Len(Text9)
                       Exit Sub
                 End If
         Case "220102"
                 'edit by nick 2004/07/07 加系統別
                 'If Mid(Text16, 1, Len(Text16) - 9) = "P" And strNation = "000" Then
                 If (Mid(Text16, 1, Len(Text16) - 9) = "P" Or Mid(Text16, 1, Len(Text16) - 9) = "PS") And strNation = "000" Then
                 Else
                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
                       Text9.SetFocus
                       Text9.SelStart = 0
                       Text9.SelLength = Len(Text9)
                       Exit Sub
                 End If
         Case "220103"
                 'edit by nick 2004/07/07 加系統別
                 'If Mid(Text16, 1, Len(Text16) - 9) = "FCT" Then
                 If (Mid(Text16, 1, Len(Text16) - 9) = "FCT" Or Mid(Text16, 1, Len(Text16) - 9) = "S") And strNation = "000" Then
                 Else
                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
                       Text9.SetFocus
                       Text9.SelStart = 0
                       Text9.SelLength = Len(Text9)
                       Exit Sub
                 End If
         Case "220104"
                 'edit by nick 2004/07/07 加系統別
                 'If Mid(Text16, 1, Len(Text16) - 9) = "FCP" Then
                 If Mid(Text16, 1, Len(Text16) - 9) = "FCP" Or Mid(Text16, 1, Len(Text16) - 9) = "FG" Then
                 Else
                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
                       Text9.SetFocus
                       Text9.SelStart = 0
                       Text9.SelLength = Len(Text9)
                       Exit Sub
                 End If
         Case "220105"
                 'edit by nick 2004/07/07 加系統別
                 'If Mid(Text16, 1, Len(Text16) - 9) = "CFT" Then
                 If (Mid(Text16, 1, Len(Text16) - 9) = "CFT" Or Mid(Text16, 1, Len(Text16) - 9) = "CFC" Or Mid(Text16, 1, Len(Text16) - 9) = "S") And strNation <> "000" Then
                 Else
                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
                       Text9.SetFocus
                       Text9.SelStart = 0
                       Text9.SelLength = Len(Text9)
                       Exit Sub
                 End If
         Case "220106"
                 'edit by nick 2004/07/07 加系統別
                 'If Mid(Text16, 1, Len(Text16) - 9) = "CFP" Then
                 If Mid(Text16, 1, Len(Text16) - 9) = "CFP" Or Mid(Text16, 1, Len(Text16) - 9) = "FCL" Or Mid(Text16, 1, Len(Text16) - 9) = "CFL" Or Mid(Text16, 1, Len(Text16) - 9) = "CPS" Then
                 Else
                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
                       Text9.SetFocus
                       Text9.SelStart = 0
                       Text9.SelLength = Len(Text9)
                       Exit Sub
                 End If
         Case "220107"
                 'add by nick 2004/07/07 加系統別
                 If Mid(Text16, 1, Len(Text16) - 9) = "TC" Then
                 Else
                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
                       Text9.SetFocus
                       Text9.SelStart = 0
                       Text9.SelLength = Len(Text9)
                       Exit Sub
                 End If
         Case "220111"
                 'edit by nick 2004/07/07 加系統別
                 'If (Mid(Text16, 1, Len(Text16) - 9) = "T" Or Mid(Text16, 1, Len(Text16) - 8) = "TF") And strNation <> "000" Then
                 If (Mid(Text16, 1, Len(Text16) - 9) = "TS" Or Mid(Text16, 1, Len(Text16) - 9) = "T" Or Mid(Text16, 1, Len(Text16) - 8) = "TF") And strNation <> "000" Then
                 Else
                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
                       Text9.SetFocus
                       Text9.SelStart = 0
                       Text9.SelLength = Len(Text9)
                       Exit Sub
                 End If
         Case "220112"
                 'edit by nick 2004/07/07 加系統別
                 'If Mid(Text16, 1, Len(Text16) - 9) = "P" And strNation <> "000" Then
                 If (Mid(Text16, 1, Len(Text16) - 9) = "P" Or Mid(Text16, 1, Len(Text16) - 9) = "PS") And strNation <> "000" Then
                 Else
                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
                       Text9.SetFocus
                       Text9.SelStart = 0
                       Text9.SelLength = Len(Text9)
                       Exit Sub
                 End If
         Case "610103"
                 'add by nick 2004/07/07 加系統別
                 If Mid(Text16, 1, Len(Text16) - 9) = "L" Or Mid(Text16, 1, Len(Text16) - 9) = "LA" Or Mid(Text16, 1, Len(Text16) - 9) = "FCL" Or Mid(Text16, 1, Len(Text16) - 9) = "CFL" Then
                 Else
                       MsgBox "科目輸入錯誤!!", , "User 輸入錯誤!!"
                       Text9.SetFocus
                       Text9.SelStart = 0
                       Text9.SelLength = Len(Text9)
                       Exit Sub
                 End If
         Case Else
         End Select
   End If
   Exit Sub
Checking:
   MsgBox MsgText(128), , MsgText(5)
   Exit Sub
End Sub

Private Sub Text16_Validate(Cancel As Boolean)
On Error GoTo Checking
   If Text16 <> MsgText(601) Then
      Text16 = CaseNoZero(Text16)
      Set adocase1 = New ADODB.Recordset
      adocase1.CursorLocation = adUseClient
      adocase1.Open "select pa01 as SystemNo from patent where pa01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and pa02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and pa03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and pa04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
                   "select tm01 as SystemNo from trademark where tm01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and tm02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and tm03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and tm04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
                   "select lc01 as SystemNo from lawcase where lc01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and lc02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and lc03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and lc04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
                   "select hc01 as SystemNo from hirecase where hc01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and hc02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and hc03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and hc04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
                   "select sp01 as SystemNo from servicepractice where sp01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and sp02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and sp03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and sp04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocase1.RecordCount = 0 Then
         MsgBox MsgText(28) & Label17, , MsgText(5)
         Cancel = True
         adocase1.Close
         Exit Sub
      End If
      adocase1.Close
   End If
   QueryCustomer
   'add by sonia 2021/1/29 以本所案號以判別FCP,FCT英日文組
   If AccNoToSalesNo(Text9, Text16) <> "" Then
      Text17 = AccNoToSalesNo(Text9, Text16)
      Text22 = StaffQuery(Text17)
   End If
   'end 2021/1/29
          '2004/07/06 nick
         '針對   P  T  TF  CFT  CFP  加入客戶名稱
         '  FCT  FCP 加入本所案號
         '這些案號的摘要 , 要在前面加入資訊
   If Text16 = MsgText(601) Then
      Exit Sub
   End If
         Select Case Mid(Text16, 1, Len(Text16) - 9)
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
                 Combo1 = Text16 & "/" & Combo1
         Case Else
         End Select
   Exit Sub
Checking:
   MsgBox MsgText(128), , MsgText(5)
   Exit Sub
End Sub

Private Sub Text17_GotFocus()
   TextInverse Text17
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text17_Validate(Cancel As Boolean)
   Text22 = ""
   If Text17 <> MsgText(601) Then
      'Modify by Amy 2017/12/05 加顯示智權人員姓名
'      If ExistCheck("staff", "st01", Text17, Label18) = False Then
'         Cancel = True
'         Exit Sub
'      End If
      If PUB_GetStaffState(Text17.Text, strExc(1), True) = 0 Then
        Cancel = True
        TextInverse Text17
      Else
        Text22.Text = strExc(1)
      End If
      'add by sonia 2021/1/29
      If SalesNoCheckAccNo(Text9, Text17) = False Then
      End If
      'end 2021/1/29
   End If
End Sub

Private Sub Text18_GotFocus()
   TextInverse Text18
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add by Amy 2013/12/26
Private Sub Text19_GotFocus()
    TextInverse Text19
    CloseIme
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text19_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   'Modify by Amy 2020/04/09
   'If Text19 <> "1" And Text19 <> "J" Then
   If InStr(GetBookKeepCmp, Text19) = 0 Then
      MsgBox Label19 & MsgText(63), , MsgText(5) '原:"公司別輸入錯誤請確認 ！"
      Cancel = True
      Exit Sub
   End If
End Sub
'end 2013/12/26

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text14 <> MsgText(601) Then
      Select Case Text14
         Case Mid(ComboItem(91), 1, 1)
            If ExistCheck("acc0i0", "a0i01", Text2, Label2) = False Then
               Cancel = True
               Exit Sub
            End If
         Case Mid(ComboItem(92), 1, 1)
            If ExistCheck("customer", "cu01", Mid(IIf(Len(Text2) = 6, AfterZero(Text2), Text2), 1, 8), Label2) = False Then
               Cancel = True
               Exit Sub
            End If
         Case Mid(ComboItem(93), 1, 1)
            If ExistCheck("staff", "st01", Text2, Label2) = False Then
               Cancel = True
               Exit Sub
            End If
      End Select
   End If
   Select Case Text14
      Case Mid(ComboItem(91), 1, 1)
         'Modify by Amy 2014/0114 往來類別為1-廠商抓其名稱及統一編號
         'Text3 = A0i02Query(Text2)
         If strSaveConfirm <> MsgText(601) And Text19 = "J" Then
            '公司別為J需show訊息
            Text3 = GetFactoryName(Text2, strFAId, True)
         Else
            Text3 = GetFactoryName(Text2, strFAId)
         End If
         Text21 = strFAId
      Case Mid(ComboItem(92), 1, 1)
         If Len(Text2) = 6 Then
            Text2 = AfterZero(Text2)
         ElseIf Len(Text2) = 8 Then
            Text2 = Text2 & "0"
         End If
         Text3 = CustomerQuery(Text2, 1)
      Case Mid(ComboItem(93), 1, 1)
         Text3 = StaffQuery(Text2)
      Case Else
         Text3 = MsgText(601)
   End Select
End Sub

Private Sub Text4_Change()
    If strSaveConfirm = MsgText(3) Then
        If Text19 = "J" And Text14 = "1" And Mid(Combo2, 1, 1) = "1" Then
            cmdDetail.Enabled = True
        Else
            cmdDetail.Enabled = False
        End If
    End If
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'2014/1/27 add by sonia
Private Sub Text4_Validate(Cancel As Boolean)
   If Text19 = "J" And Text14 = "1" And Mid(Combo2, 1, 1) = "1" Then
      If Text4 = "" Then
         'Modify by Amy 2019/06/24 +確定無發票號碼請輸入00 訊息
         MsgBox MsgText(10) & Label3 & vbCrLf & "確定無發票號碼請輸入00", , MsgText(5)
         Cancel = True
         Exit Sub
      End If
   End If
   'Add by Amy 2018/09/13 G10700329 只改此畫面之發票號未更新明細,造成資料不一致
   'Modify by Amy 2018/10/29 發票號碼為00略過不跳下一個畫面
    If strSaveConfirm = MsgText(4) And cmdDetail.Enabled = True And Text4 <> "00" Then
        If bolBack = False Then
            Call cmdDetail_Click
        End If
   End If
End Sub
'2014/1/27 end

Private Sub Text5_GotFocus()
   TextInverse Text5
   'edit by nickc 2007/06/11  切換輸入法改用API
   OpenIme
End Sub

Private Sub Text5_LostFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   CloseIme
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub

Private Sub Text9_Change()
   Text10 = A0102Query(Text9)
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

'*************************************************
'  借貸方檢核
'
'*************************************************
Public Function CreDebCheck() As String
   If Text6 = Text7 Then
      CreDebCheck = MsgText(602)
   End If
End Function

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = adoacc0o0.Bookmark & MsgText(35) & adoacc0o0.RecordCount
End Sub

'*************************************************
'  關閉分錄欄位輸入狀態
'  2013/12/26 +應付款資料
'*************************************************
Public Sub FormDisabled()
   'Add by Amy 2013/12/26
    Text1.Locked = False
    Text19.Locked = True
    Text14.Locked = True
    Text2.Locked = True
    Combo2.Locked = True
    Text4.Locked = True
    MaskEdBox1.Enabled = False
    Text5.Locked = True
    MaskEdBox2.Enabled = False
  
   '+J公司且應付款類別為1 明細鈕才開放
   If Text19 = "J" And Text14 = "1" And Mid(Combo2, 1, 1) = "1" Then
        cmdDetail.Enabled = True
   Else
        cmdDetail.Enabled = False
   End If
   'end 2013/12/26
   Text8.Enabled = False
   Text9.Enabled = False
   Text11.Enabled = False
   Text12.Enabled = False
   'Text15.Enabled = False 'Modify 2013/12/26
   Combo1.Enabled = False
   Command1.Enabled = False 'Add by Amy 2014/01/07
   Command2.Enabled = False
   Text16.Enabled = False
   Text18.Enabled = False
   Text17.Enabled = False
End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'  2013/12/26 +應付款資料欄位
'*************************************************
Public Sub FormEnabled()
    'Add by Amy 2013/12/26
    Text1.Locked = True
    Text2.Locked = False
    'Combo2.Locked = False 'Modify by Amy 2014/01/07 往下搬
    Text4.Locked = False
    MaskEdBox1.Enabled = True
    Text5.Locked = False
    MaskEdBox2.Enabled = True
    If strSaveConfirm = MsgText(3) Then
       '新增
       Combo2.Locked = False
       Text19.Locked = False
       Text14.Locked = False
    Else
        Combo2.Locked = True
       Text19.Locked = True
       Text14.Locked = True
       'Add by Amy 2014/11/04 +a1p22有值不可修改入帳日
        If CheckExistA1p22(Text19, "B", Text1) = True Then
            MaskEdBox1.Enabled = False
        End If
        'end 2014/11/04
    End If
   '+J公司且應付款類別為1 明細鈕才開放
   If Text19 = "J" And Text14 = "1" And Mid(Combo2, 1, 1) = "1" Then
        cmdDetail.Enabled = True
   Else
        cmdDetail.Enabled = False
   End If
   'end 2013/12/26
   Text8.Enabled = True
   Text9.Enabled = True
   Text11.Enabled = True
   Text12.Enabled = True
   'Text15.Enabled = True 'Modify 2013/12/26
   Combo1.Enabled = True
   Command1.Enabled = True 'Add by Amy 2014/01/07
   Command2.Enabled = True
   Text16.Enabled = True
   Text18.Enabled = True
   Text17.Enabled = True
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   If Text9 <> MsgText(601) Then
      'Modify by Amy 2013/12/26 +公司別確認
'      If ExistCheck("acc010", "a0101", Text9, Label10) = False Then
'         Cancel = True
'         Exit Sub
'      End If
      If PUB_CheckCompany(Text9, Text19) = False Then
         Cancel = True
         Exit Sub
      End If
      'end 2013/12/26
   End If
   '93.12.29 ADD BY SONIA
   '2005/3/29 MODIFY BY SONIA
   'If Text9 = "2112" And (Text2 <= "V" Or Text2 >= "W") Then
   '   MsgBox "會計科目與往來對象不符!!", , "User 輸入錯誤!!"
   '   Cancel = True
   '   Exit Sub
   'End If
   If Text9 = "2112" And (Mid(Text2, 1, 1) <> "V" And Mid(Text2, 1, 1) <> "X") Then
      MsgBox "會計科目與往來對象不符!!", , "User 輸入錯誤!!"
      Cancel = True
      Exit Sub
   End If
   '2005/3/29 END
   If Text9 = "2113" And (Text2 <= "F" Or Text2 >= "G") Then
      MsgBox "會計科目與往來對象不符!!", , "User 輸入錯誤!!"
      Cancel = True
      Exit Sub
   End If
   '93.12.29 END
   If Combo1 = "" Then  '2014/3/24 add by sonia 無摘要時才預設
      If Mid(Text9, 1, 1) = "2" Then
         Combo1 = Text3
      End If
      Select Case Text9
         Case "2112", "2113"
            adoaccsum.CursorLocation = adUseClient
            'Modify by Amy 2013/12/26 +公司別 原:'1'
            adoaccsum.Open "select a1p14 from acc1p0, acc010, acc090 where a1p05 = a0101 and a1p06 = a0901 and a1p01 = '" & Text19 & "' and a1p02 = 'B' and a1p04 = '" & Text1 & "' and a1p07 <> 0 order by a1p03 desc", adoTaie, adOpenStatic, adLockReadOnly
            If adoaccsum.RecordCount <> 0 Then
               If IsNull(adoaccsum.Fields(0).Value) = False Then
                  Combo1 = Combo1 & " / " & adoaccsum.Fields(0).Value
               End If
            End If
            adoaccsum.Close
      End Select
   End If           '2014/3/24 end
   'add by sonia 2021/1/29 以科目帶智權人員,加傳本所案號以判別FCP,FCT英日文組
   If AccNoToSalesNo(Text9, Text16) <> "" Then
      Text17 = AccNoToSalesNo(Text9, Text16)
      Text22 = StaffQuery(Text17)
   End If
   'end 2021/1/29
End Sub

'*************************************************
'  重新整理國內應付資料
'
'*************************************************
Public Sub Acc0o0Refresh()
On Error GoTo Checking
   If adoacc0o0.State = adStateOpen Then
      adoacc0o0.Close
   End If
   adoacc0o0.CursorLocation = adUseClient
   adoacc0o0.MaxRecords = intMax
   adoacc0o0.Open "select * from acc0o0 where a0o01 >= '" & Text1 & "' order by a0o01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  以本所案號查詢客戶名稱
'
'*************************************************
Public Sub QueryCustomer()
Dim strSql As String

   If Text16 = MsgText(601) Then
      Exit Sub
   End If
   strSql = "select cu01||cu02 as Name, cu04, cu05, cu06, cu88, cu89, cu90 from patent, customer where substr(pa26, 1, 8) = cu01 and nvl(substr(pa26, 9, 1), '0') = cu02 and pa01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and pa02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and pa03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and pa04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
            "select cu01||cu02 as Name, cu04, cu05, cu06, cu88, cu89, cu90 from trademark, customer where substr(tm23, 1, 8) = cu01 and nvl(substr(tm23, 9, 1), '0') = cu02 and tm01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and tm02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and tm03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and tm04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
            "select cu01||cu02 as Name, cu04, cu05, cu06, cu88, cu89, cu90 from lawcase, customer where substr(lc11, 1, 8) = cu01 and nvl(substr(lc11, 9, 1), '0') = cu02 and lc01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and lc02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and lc03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and lc04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
            "select cu01||cu02 as Name, cu04, cu05, cu06, cu88, cu89, cu90 from hirecase, customer where substr(hc05, 1, 8) = cu01 and nvl(substr(hc05, 9, 1), '0') = cu02 and hc01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and hc02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and hc03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and hc04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "' union " & _
            "select cu01||cu02 as Name, cu04, cu05, cu06, cu88, cu89, cu90 from servicepractice, customer where substr(sp08, 1, 8) = cu01 and nvl(substr(sp08, 9, 1), '0') = cu02 and sp01 = '" & Mid(Text16, 1, Len(Text16) - 9) & "' and sp02 = '" & Mid(Text16, Len(Text16) - 8, 6) & "' and sp03 = '" & Mid(Text16, Len(Text16) - 2, 1) & "' and sp04 = '" & Mid(Text16, Len(Text16) - 1, 2) & "'"
    'add by nick 2004/07/07
   adocase1.CursorLocation = adUseClient
   adocase1.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adocase1.RecordCount <> 0 Then
      If IsNull(adocase1.Fields("cu04").Value) Then
         If IsNull(adocase1.Fields("cu05").Value) Then
            If IsNull(adocase1.Fields("cu06").Value) Then
               Combo1 = Combo1 & MsgText(601)
            Else
               Combo1 = Combo1 & IIf(Me.Combo1.Text <> "", " / ", "") & adocase1.Fields("cu06").Value
            End If
         Else
            Combo1 = Combo1 & IIf(Me.Combo1.Text <> "", " / ", "") & adocase1.Fields("cu05").Value
            If IsNull(adocase1.Fields("cu88").Value) = False Then
               Combo1 = Combo1 & " " & adocase1.Fields("cu88").Value
            End If
            If IsNull(adocase1.Fields("cu89").Value) = False Then
               Combo1 = Combo1 & " " & adocase1.Fields("cu89").Value
            End If
            If IsNull(adocase1.Fields("cu90").Value) = False Then
               Combo1 = Combo1 & " " & adocase1.Fields("cu90").Value
            End If
         End If
      Else
         Combo1 = Combo1 & IIf(Me.Combo1.Text <> "", " / ", "") & adocase1.Fields("cu04").Value
      End If
   End If
   adocase1.Close
End Sub

'Add by Amy 2013/12/26
'廠商編號取得廠商名稱及身份證字號/統編,無統編可秀訊息
Private Function GetFactoryName(strFactoryNo As String, ByRef strFactoryID As String, Optional bolMsg As Boolean = False) As String
    GetFactoryName = ""
    strFactoryID = ""
    strExc(0) = "Select * From Acc0I0 Where A0I01='" & strFactoryNo & "' "
    intI = 0
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
          GetFactoryName = "" & RsTemp.Fields("a0i02")
          strFactoryID = "" & RsTemp.Fields("a0i18")
          'Modify  by Amy 2014/01/13 +編號不為F才秀訊息
          'Add by Amy 2014/01/07 +是否秀訊息
          If bolMsg = True And strFactoryID = "" And Left(strFactoryNo, 1) <> "F" Then
                MsgBox "該廠商無統一編號,請自行至廠商基本資料補輸", , MsgText(5)
          End If
    End If
End Function

'由aacc_sav.bas搬回
Public Sub Frmacc1170_Save()
Dim adocheck As New ADODB.Recordset
Dim strYes As String
Dim bCancel As Boolean 'Add by Amy 2014/01/07
Dim strMsg As String 'Add by Amy 2014/10/27

On Error GoTo Checking
    bCancel = False
   
      If Text1 = MsgText(601) Then
         MsgBox MsgText(10) & Label1, , MsgText(5)
         strControlButton = MsgText(602)
         Text1.SetFocus
         Exit Sub
      Else
         'Add by Amy 2013/12/26 +公司別必填
         If Text19 = MsgText(601) Then
            MsgBox MsgText(10) & Label19, , MsgText(5)
            strControlButton = MsgText(602)
            Text19.SetFocus
            Exit Sub
         End If
         'Add by Amy 2014/01/07
         Call Text19_Validate(bCancel)
         If bCancel = True Then
            strControlButton = MsgText(602)
            Text19.SetFocus
            Exit Sub
         End If
         'end 2014/01/07
         If strSaveConfirm = MsgText(3) And Text19.Tag <> "" Then
            If Text19.Tag <> Text19 Then
                MsgBox Label19 & "與明細資料公司別不一致，請確認！", , MsgText(5)
                strControlButton = MsgText(602)
                Text19.SetFocus
                Exit Sub
            End If
         End If
         'end 2013/12/26
         If Text14 <> MsgText(601) Then
            Select Case Text14
               Case Mid(ComboItem(91), 1, 1)
                  If ExistCheck("acc0i0", "a0i01", Text2, Label2) = False Then
                     strControlButton = MsgText(602)
                     Text2.SetFocus
                     Exit Sub
                  End If
               Case Mid(ComboItem(92), 1, 1)
                  If ExistCheck("customer", "cu01 || cu02", IIf(Len(Text2) = 6, AfterZero(Text2), Text2), Label2) = False Then
                     strControlButton = MsgText(602)
                     Text2.SetFocus
                     Exit Sub
                  End If
               Case Mid(ComboItem(93), 1, 1)
                  If ExistCheck("staff", "st01", Text2, Label2) = False Then
                     strControlButton = MsgText(602)
                     Text2.SetFocus
                     Exit Sub
                  End If
            End Select
         '2013/1/9 add by sonia
         Else
            MsgBox "請輸入往來類別 !", , MsgText(5)
            strControlButton = MsgText(602)
            Text14.SetFocus
            Exit Sub
         '2013/1/9 end
         End If
         'Modify by Amy 2014/10/27 +入帳日設必填及與系統日的檢查
         If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
            MsgBox Label4 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            MaskEdBox1.SetFocus
            Exit Sub
         End If
         If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
            MsgBox Label4 & MsgText(63), , MsgText(5)
            strControlButton = MsgText(602)
            MaskEdBox1.SetFocus
            Exit Sub
         End If
         If MaskEdBox1.Enabled = True Then
            If ChkWorkData(Text19, DBDATE(MaskEdBox1), strMsg) = False Then
                MsgBox Label4 & strMsg, , MsgText(5)
                strControlButton = MsgText(602)
                MaskEdBox1.SetFocus
                Exit Sub
            End If
         End If
         'end 2014/10/27
         'Add by Amy 2023/12/08 欲處理日不可空白
         If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
            MsgBox Label5 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            MaskEdBox2.SetFocus
            Exit Sub
         End If
         If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
            If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
               MsgBox Label5 & MsgText(63), , MsgText(5)
               strControlButton = MsgText(602)
               MaskEdBox2.SetFocus
               Exit Sub
            End If
         End If
         '2007/10/30 add by sonia
         If Combo2 = MsgText(601) Then
            MsgBox MsgText(10) & Label16, , MsgText(5)
            strControlButton = MsgText(602)
            Combo2.SetFocus
            Exit Sub
         ElseIf Mid(Combo2, 1, 1) <> "1" Then
            MsgBox MsgText(63) & Label16, , MsgText(5)
            strControlButton = MsgText(602)
            Combo2.SetFocus
            Exit Sub
         End If
         '2007/10/30 end
      End If
      'Add by Amy 2014/01/07 +檢查會計科目
      Call Text9_Validate(bCancel)
      If bCancel = True Then
            strControlButton = MsgText(602)
            Text9.SetFocus
            Exit Sub
      End If
      'end 2014/01/07
      'add by sonia 2015/12/30 檢查民國105年起法務收入科目不可使用
      intI = PUB_AccNoEnable(Text9, Val(FCDate(MaskEdBox1.Text)))
      If intI <> 0 Then
         strControlButton = MsgText(602)
         Text9.SetFocus
         Exit Sub
      End If
      'end 2015/12/30
      '2007/8/8 ADD BY SONIA 檢查科目部門&智權人員是否正確
      intI = PUB_AccNoGood(Text9, Text12, Text17)
      If intI <> 0 Then
         strControlButton = MsgText(602)
         If intI = 1 Then
            Text9.SetFocus
         ElseIf intI = 2 Then
            Text12.SetFocus
         ElseIf intI = 3 Then
            Text17.SetFocus
         End If
         Exit Sub
      End If
      '2007/8/8 END
      'Add by Amy 2021/09/01 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
      If PUB_ChkUniText(Me) = False Then
         strControlButton = MsgText(602)
         Exit Sub
      End If
       
      If strSaveConfirm = MsgText(3) Then
         If adoacc0o0.RecordCount <> 0 Then
            adoacc0o0.Find "a0o01 = '" & Text1 & "'", 0, adSearchForward, 1
            If adoacc0o0.EOF = False Then
               'Modify by Amy 2014/10/27 解新增insert完又改畫面有Acc0o0資料不會更新的問題
               'Exit Sub
               GoTo NextRecord
            End If
         End If
         adoacc0o0.AddNew
      End If
      
NextRecord:
      adoacc0o0.Fields("a0o01").Value = Text1
      adoacc0o0.Fields("a0o07").Value = Text19 'Add by Amy 2013/12/26 +公司別
      If Text14 <> MsgText(601) Then
         adoacc0o0.Fields("a0o02").Value = Text14
      Else
         adoacc0o0.Fields("a0o02").Value = Null
      End If
      If Text2 <> MsgText(601) Then
         adoacc0o0.Fields("a0o03").Value = Text2
      Else
         adoacc0o0.Fields("a0o03").Value = Null
      End If
      If Combo2 <> MsgText(601) Then
         adoacc0o0.Fields("a0o19").Value = Mid(Combo2, 1, 1)
      Else
         adoacc0o0.Fields("a0o19").Value = Null
      End If
      If Text4 <> MsgText(601) Then
         adoacc0o0.Fields("a0o04").Value = Text4
      Else
         adoacc0o0.Fields("a0o04").Value = Null
      End If
      If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
         adoacc0o0.Fields("a0o05").Value = Val(FCDate(MaskEdBox1.Text))
      Else
         adoacc0o0.Fields("a0o05").Value = Null
      End If
      If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
         adoacc0o0.Fields("a0o06").Value = Val(FCDate(MaskEdBox2.Text))
      Else
         adoacc0o0.Fields("a0o06").Value = Null
      End If
      If Text5 <> MsgText(601) Then
         adoacc0o0.Fields("a0o10").Value = Text5
      Else
         adoacc0o0.Fields("a0o10").Value = Null
      End If
      If strSaveConfirm = MsgText(3) Then
         adoacc0o0.Fields("a0o13").Value = Val(strSrvDate(2))
         adoacc0o0.Fields("a0o14").Value = ServerTime
         adoacc0o0.Fields("a0o15").Value = strUserNum
      Else
         adoacc0o0.Fields("a0o16").Value = Val(strSrvDate(2))
         adoacc0o0.Fields("a0o17").Value = ServerTime
         adoacc0o0.Fields("a0o18").Value = strUserNum
      End If
      If strSaveConfirm = MsgText(4) Then
         If strAccNo <> MsgText(601) Then
            'Modify by Amy 2013/12/26 改公司別 原:'1'
            'adoTaie.Execute "update acc1p0 set a1p22 = '" & strAccNo & "', a1p27 = '" & MsgText(602) & "' where a1p01 = '1' and a1p02 = 'B' and a1p04 = '" & Text1 & "'"
            adoTaie.Execute "update acc1p0 set a1p22 = '" & strAccNo & "', a1p27 = '" & MsgText(602) & "' where a1p01 = '" & Text19 & "' and a1p02 = 'B' and a1p04 = '" & Text1 & "'"
         End If
      End If
      adoacc0o0.UpdateBatch
      RecordShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'由aacc_cls.bas搬回
Public Sub Frmacc1170_Clear()
      bolFirstIns = False 'Add by Amy 2014/10/27
      Text1 = ""
      'Add by Amy 2013/12/26
      Text19 = ""  '公司別
      Text19.Tag = ""   '2014/3/24 add by sonia
      Text21 = "" '統一編號
      'end 2013/12/26
      Text14 = ""
      Text2 = ""
      Text2.Tag = "" 'Add by Amy 2014/10/27
      Combo2 = ""
      Text3 = ""
      Text4 = ""
      Text4.Tag = "" 'Add by Amy 2018/09/13
      If MaskEdBox1.Text = MsgText(29) Or MaskEdBox1.Text = MsgText(601) Then
         MaskEdBox1.Mask = ""
         MaskEdBox1.Text = ""
         MaskEdBox1.Mask = DFormat
      End If
      MaskEdBox1.Tag = "" 'Add by Amy 2014/10/27
      If MaskEdBox2.Text = MsgText(29) Or MaskEdBox2.Text = MsgText(601) Then
         MaskEdBox2.Mask = ""
         MaskEdBox2.Text = ""
         MaskEdBox2.Mask = DFormat
      End If
      Text5 = ""
      Text6 = ""
      Text7 = ""
      Text20 = ""
      AdodcRefresh
      AdodcClear
      Text19.SetFocus 'Modify by Amy 2013/12/26 原:Text14.SetFocus
End Sub

'由aacc_del.bas搬回
Public Sub Frmacc1170_Delete()
On Error GoTo Checking
      If DeleteCheck("select a0o01 from acc0o0 where a0o01 = '" & Text1 & "'") = MsgText(603) Then
         Exit Sub
      End If
      'Modify by Amy 2013/12/26 改公司別 原:'1'
      adoTaie.Execute "delete from acc1p0 where a1p01 = '" & Text19 & "' and a1p02 = 'B' and a1p04 = '" & Text1 & "'"
      adoacc1p0.Requery
      adoTaie.Execute "delete from acc0o0 where a0o01 = '" & Text1 & "'"
      adoacc0o0.Requery
      'Add by Amy 2017/09/14 刪acc450進項發票資料
      adoTaie.Execute "Delete From acc450 Where a4501 = '" & Text1 & "'"
      AdodcRefresh
      If adoacc0o0.RecordCount <> 0 Then
         adoacc0o0.MoveFirst
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

Public Sub Frmacc1170_First()
      'Modify  by Amy 2014/01/15
'      CreDebCheck
'       If CreDebCheck <> MsgText(602) Then
'          MsgBox MsgText(11), , MsgText(5)
'          Exit Sub
'       End If
       If FormCheck = False Then
            Exit Sub
       End If
       'end 2013/01/15
      If adoacc0o0.RecordCount <> 0 Then
         adoacc0o0.MoveFirst
         FormShow
         AdodcRefresh
         SumShow
         RecordShow
         FormDisabled 'Add by Amy  2013/12/26
      End If
      AdodcClear 'Add by Amy 2014/01/07
End Sub

Public Sub Frmacc1170_Last()
      'Modify  by Amy 2014/01/15
'      CreDebCheck
'      If CreDebCheck <> MsgText(602) Then
'         MsgBox MsgText(11), , MsgText(5)
'         Exit Sub
'      End If
       If FormCheck = False Then
            Exit Sub
       End If
       'end 2013/01/15
      If adoacc0o0.RecordCount <> 0 Then
         adoacc0o0.MoveLast
         FormShow
         AdodcRefresh
         SumShow
         RecordShow
         FormDisabled 'Add by Amy  2013/12/26
      End If
      AdodcClear 'Add by Amy 2014/01/07
  End Sub
  
  Public Sub Frmacc1170_Next()
      'Modify  by Amy 2014/01/15
'      CreDebCheck
'      If CreDebCheck <> MsgText(602) Then
'         MsgBox MsgText(11), , MsgText(5)
'         Exit Sub
'      End If
       If FormCheck = False Then
            Exit Sub
       End If
       'end 2013/01/15
      If adoacc0o0.EOF = False Then
         adoacc0o0.MoveNext
         If adoacc0o0.EOF Then
            adoacc0o0.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         FormShow
         AdodcRefresh
         SumShow
         RecordShow
         FormDisabled 'Add by Amy  2013/12/26
      End If
      AdodcClear 'Add by Amy 2014/01/07
End Sub

Public Sub Frmacc1170_Previous()
      'Modify  by Amy 2014/01/15
'      CreDebCheck
'      If CreDebCheck <> MsgText(602) Then
'         MsgBox MsgText(11), , MsgText(5)
'         Exit Sub
'      End If
       'Add by Amy 2014/01/15
       If FormCheck = False Then
            Exit Sub
       End If
       'end 2013/01/15
      If adoacc0o0.BOF = False Then
         adoacc0o0.MovePrevious
         If adoacc0o0.BOF Then
            adoacc0o0.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         FormShow
         AdodcRefresh
         SumShow
         RecordShow
         FormDisabled 'Add by Amy  2013/12/26
      End If
      AdodcClear 'Add by Amy 2014/01/07
End Sub
'end 2013/12/26

'Add by Amy 2014/01/15
'由acc_var 搬回 上/下/第一/最後 筆 FormUnload 都有用到
Public Function FormCheck() As Boolean
    Dim bolCancel As Boolean 'Add by Amy 2020/04/09
    
    FormCheck = True
    'Add by Amy 2020/04/09
    If Text19 <> MsgText(601) Then
        Call Text19_Validate(bolCancel)
        If bolCancel = True Then
            FormCheck = False
            Exit Function
        End If
    End If
    'end 2020/04/09
    'Add by Amy 2020/11/06 L公司6及72字頭部門別不可輸TOT或空值
    Call Text12_Validate(bolCancel)
    If bolCancel = True Then
        FormCheck = False
        Exit Function
    End If
    
    If CreDebCheck <> MsgText(602) Then
        MsgBox MsgText(11), , MsgText(5)
        FormCheck = False
        Exit Function
    End If

    If Text19 = "J" And Text14 = "1" And Mid(Combo2, 1, 1) = "1" Then
        If InComeTaxChk = False Then
            MsgBox "進項稅額合計與明細營業稅合計不同！", , MsgText(5)
            FormCheck = False
            Exit Function
        End If
    End If
End Function

Public Function InComeTaxChk() As Boolean
    '進項稅額與明細營業稅總額是否相同
    InComeTaxChk = True
    strExc(0) = "Select Nvl(Tax1,0) as Tax1,Nvl(Tax2,0) as Tax2 From (Select a1p04,sum(a1p07) as Tax1 From acc1p0 Where a1p01='" & Text19 & "' and a1p04='" & Text1 & "' and a1p05='1211' Group by a1p04), " & _
                      "( Select a4501,sum(a4508) as Tax2 From acc450 WHERE a4501 = '" & Text1 & "' Group by a4501) Where a1p04=a4501(+) "
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        If Val(RsTemp.Fields("Tax1")) <> Val(RsTemp.Fields("Tax2")) Then
            InComeTaxChk = False
        End If
    End If
End Function
'end 2014/01/15

'Add by Amy 2014/10/27 為資料一致更新acc1p0
Public Sub UpdateAcc1p0()
    Dim strUpd As String
    
On Error GoTo ChkHand
  
    If strSaveConfirm = MsgText(3) Then
        strUpd = "Update Acc1p0 set a1p15='" & Text2 & "',a1p18='" & Val(FCDate(MaskEdBox1)) & "' " & _
                     "Where a1p01='" & Text19 & "' And a1p04='" & Text1 & "' And a1p02='B' "
        adoTaie.Execute strUpd
    ElseIf strSaveConfirm = MsgText(4) Then
        If Text2.Tag <> Text2 Then strUpd = strUpd & ",a1p15='" & Text2 & "' "
        If Val(MaskEdBox1.Tag) <> Val(FCDate(MaskEdBox1)) Then strUpd = strUpd & ",a1p18='" & Val(FCDate(MaskEdBox1)) & "' "
        If strUpd <> "" Then
            strUpd = "Update Acc1p0 set " & Mid(strUpd, 2) & IIf(MaskEdBox1.Enabled = True, "", ",a1p27='Y'") & " Where a1p01='" & Text19 & "' And a1p04='" & Text1 & "' And a1p02='B' "
            adoTaie.Execute strUpd
        End If
    End If
    
ChkHand:
    If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox "UpdateAcc1p0 錯誤:" & Err.Description, , MsgText(5)
   strControlButton = MsgText(602)
End Sub

Public Sub SetData(ByVal strKeyCode As String)
    Select Case strKeyCode
        Case "F9"
            '解改日期存檔再修改不會存acc1p0 (因tag只記錄前一次改前資料)
            MaskEdBox1.Tag = Val(FCDate(MaskEdBox1))
            Text2.Tag = Text2
            bolDelAll = False 'Add by Amy 2016/07/21
        Case "F10"
            bolDelAll = False 'Add by Amy 2016/07/21
        Case Else
    End Select
End Sub
'end 2014/10/27

