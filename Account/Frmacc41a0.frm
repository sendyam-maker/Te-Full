VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc41a0 
   AutoRedraw      =   -1  'True
   Caption         =   "CF案件結餘結算作業"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   9560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   9560
   Begin VB.TextBox Text16 
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
      Left            =   7800
      TabIndex        =   34
      Top             =   960
      Width           =   1550
   End
   Begin VB.TextBox Text14 
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
      Left            =   4065
      TabIndex        =   32
      Top             =   960
      Width           =   2745
   End
   Begin VB.TextBox Text8 
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
      Left            =   4584
      TabIndex        =   28
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "輸入傳票資料"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7464
      TabIndex        =   2
      Top             =   4728
      Width           =   1575
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
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   26
      Top             =   4728
      Width           =   1575
   End
   Begin VB.TextBox Text12 
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
      Left            =   2910
      TabIndex        =   25
      Top             =   4728
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Height          =   300
      Left            =   2520
      Picture         =   "Frmacc41a0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   240
      Width           =   350
   End
   Begin VB.TextBox Text11 
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
      MaxLength       =   5
      TabIndex        =   4
      Top             =   4368
      Width           =   1575
   End
   Begin VB.TextBox Text9 
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
      Left            =   7464
      TabIndex        =   21
      Top             =   4008
      Width           =   1572
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
      Height          =   300
      Left            =   7824
      TabIndex        =   18
      Top             =   3504
      Width           =   1215
   End
   Begin VB.TextBox Text6 
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
      Left            =   3456
      TabIndex        =   17
      Top             =   3480
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc41a0.frx":0102
      Height          =   1755
      Left            =   240
      TabIndex        =   7
      Top             =   1665
      Width           =   5895
      _ExtentX        =   10389
      _ExtentY        =   3104
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
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
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "cp09"
         Caption         =   "收文號"
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
         DataField       =   "property"
         Caption         =   "案件性質"
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
         DataField       =   "st02"
         Caption         =   "智權人員"
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
         DataField       =   "Ramount"
         Caption         =   "收款金額"
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
         DataField       =   "Samount"
         Caption         =   "已作收入金額"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   950.173
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1120.252
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1319.811
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text5 
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
      TabIndex        =   15
      Top             =   4008
      Width           =   1572
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
      Height          =   300
      Left            =   1320
      TabIndex        =   13
      Top             =   600
      Width           =   1572
   End
   Begin VB.TextBox Text1 
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
      Left            =   4050
      TabIndex        =   10
      Top             =   240
      Width           =   1572
   End
   Begin VB.TextBox Text10 
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
      Top             =   240
      Width           =   1215
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   1575
      _ExtentX        =   2787
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   14737632
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   7464
      TabIndex        =   5
      Top             =   4368
      Width           =   1572
      _ExtentX        =   2787
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   570
      Top             =   2085
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Frmacc41a0.frx":0117
      Height          =   1755
      Left            =   6135
      TabIndex        =   29
      Top             =   1665
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   3104
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   20
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
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "a0102"
         Caption         =   "支出科目"
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
         DataField       =   "Amount"
         Caption         =   "支出金額"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1470.047
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1310.173
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   6120
      Top             =   2085
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
   Begin MSForms.TextBox Text15 
      Height          =   300
      Left            =   1305
      TabIndex        =   33
      Top             =   1320
      Width           =   8040
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text4 
      Height          =   300
      Left            =   2910
      TabIndex        =   24
      Top             =   4365
      Width           =   1575
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   312
      Left            =   2880
      TabIndex        =   6
      Top             =   600
      Width           =   6300
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "7646;591"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text2 
      Height          =   300
      Left            =   5640
      TabIndex        =   11
      Top             =   240
      Width           =   3540
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label13 
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
      Height          =   255
      Left            =   6960
      TabIndex        =   35
      Top             =   1005
      Width           =   735
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "申請國家"
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
      Left            =   3030
      TabIndex        =   31
      Top             =   1005
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "代理人"
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
      TabIndex        =   30
      Top             =   1365
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "部門"
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
      TabIndex        =   27
      Top             =   4728
      Width           =   852
   End
   Begin VB.Label Label10 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "結算日期"
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
      Left            =   6384
      TabIndex        =   23
      Top             =   4368
      Width           =   972
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   22
      Top             =   4365
      Width           =   900
   End
   Begin VB.Label Label8 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "結轉收入金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   5904
      TabIndex        =   20
      Top             =   4008
      Width           =   1452
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "小計"
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
      Left            =   2544
      TabIndex        =   19
      Top             =   3480
      Width           =   492
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "填表日期"
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
      TabIndex        =   16
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "浮動金額"
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
      TabIndex        =   14
      Top             =   4008
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "本所案號"
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
      TabIndex        =   12
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "申請人"
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
      Left            =   3120
      TabIndex        =   9
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "結餘單號"
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
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4608
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc41a0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/25 Form2.0已修改 Text2/text4/Text15/Combo1/DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc240 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoadodc2 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Dim strAddNo As String 'Add by Morgan 2011/6/23

Private Sub Command1_Click()
Dim bCancel As Boolean
   'edit by nickc 2005/11/04 修改成新增及修改才可以輸傳票
   'If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   If Text10 = "" Then
      Exit Sub
   End If
   '2012/6/27 add by sonia
   MaskEdBox2_Validate bCancel
   If bCancel Then Exit Sub
   '2012/6/27 end
   Dim oCP01 As String
   Dim oCP02 As String
   Dim oCP03 As String
   Dim oCP04 As String
   If Text3 <> "" Then
      oCP01 = Mid(Text3, 1, Len(Text3) - 9)
      oCP02 = Mid(Text3, Len(Text3) - 8, 6)
      oCP03 = Mid(Text3, Len(Text3) - 2, 1)
      oCP04 = Mid(Text3, Len(Text3) - 1, 2)
   End If
   Screen.MousePointer = vbHourglass
   'add by nickc 2005/09/08 先預設傳票資料
   Dim StrSqlB As String
   Dim ChkRs As New ADODB.Recordset
   Dim tmpSN01 As String
   Dim tmpSNGp As String
   'add by nickc 2005/09/29
   Dim bolIsTransReserve As Boolean
   
   Set ChkRs = New ADODB.Recordset
   StrSqlB = "select sn01 from salesno where sn02='" & Text11 & "' "
   ChkRs.CursorLocation = adUseClient
   ChkRs.Open StrSqlB, adoTaie, adOpenStatic, adLockReadOnly
   tmpSN01 = ""
   If ChkRs.RecordCount <> 0 Then
      tmpSN01 = CheckStr(ChkRs.Fields("sn01").Value)
   End If
   Set ChkRs = Nothing
   tmpSNGp = ""
   If Text11 = "M0100" Then
      tmpSNGp = "結餘總"
   Else
      Set ChkRs = New ADODB.Recordset
      StrSqlB = "select decode(st06,'1','結餘北','2','結餘中','3','結餘南','4','結餘高','') as st06 from staff where st01='" & Text11 & "' "
      ChkRs.CursorLocation = adUseClient
      ChkRs.Open StrSqlB, adoTaie, adOpenStatic, adLockReadOnly
      If ChkRs.RecordCount <> 0 Then
         tmpSNGp = CheckStr(ChkRs.Fields("st06").Value)
      End If
   End If
   Set ChkRs = Nothing
   Set ChkRs = New ADODB.Recordset
   'Modify by Amy +公司別
   'StrSqlB = "select * from acc1p0 where a1p01='1' and a1p02='S' and a1p04='" & Text10 & "' "
   StrSqlB = "select * from acc1p0 where a1p01='" & Text16.Tag & "' and a1p02='S' and a1p04='" & Text10 & "' "
   ChkRs.CursorLocation = adUseClient
   ChkRs.Open StrSqlB, adoTaie, adOpenStatic, adLockReadOnly
   If ChkRs.RecordCount = 0 Then
      'add by nickc 2005/09/29 收入 > 0 問轉收入還是保留
      '2011/3/25 MODIFY BY SONIA 辜說一律轉保留,這樣才能區分結餘轉收入及正常收入
'      bolIsTransReserve = False
'      If Val(Text9.Text) > 0 Then
'         If MsgBox("轉收入(Y)？轉保留(N)？", vbQuestion + vbYesNo + vbDefaultButton1, "一定要選喔！") = vbNo Then
'            bolIsTransReserve = True
'         End If
'      End If
      bolIsTransReserve = True
      '2011/3/25 END
      If Val(Text9.Text) + Val(Text5.Text) <> 0 Then
            '2009/7/22 MODIFY BY SONIA 摘要最前面加本所案號
            'Modify by Amy 2013/12/17 原:acc1p0 values '1'
            'modify by sonia 2022/6/30 摘要取消智權人員簡稱欄tmpSN01
            If (Val(Text9.Text) > 0) Or (Val(Text9.Text) = 0 And Val(Text5.Text) > 0) Then      'Text9結轉金額 >= 0 的
               adoTaie.Execute "insert into acc1p0 (a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p15,a1p16,a1p17,a1p18,a1p30) " & _
                                         " values ('" & Text16.Tag & "','S','001','" & Text10 & "','" & IIf(oCP01 = "CFP", "220106", IIf(oCP01 = "CPS", "220106", IIf(oCP01 = "P", "220112", IIf(oCP01 = "PS", "220112", IIf(oCP01 = "CFT", "220105", IIf(oCP01 = "CFC", "220105", IIf(oCP01 = "S", "220105", "220111"))))))) & "' " & _
                                         ",'TOT'," & Abs(Text5.Text) + Abs(Text9.Text) & ",0,'" & Text3 & "/" & StrToStr(ChgSQL(Text2), 8) & "/結餘','" & Text1 & "','" & Text11 & "','" & Text3 & "'," & MaskEdBox2.ClipText & ",null )"
               If Abs(Text9.Text) > 0 Then
                   'modify by sonia 2016/1/5 4121改412101,4131改413101
                   adoTaie.Execute "insert into acc1p0 (a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p15,a1p16,a1p17,a1p18,a1p30) " & _
                                         " values ('" & Text16.Tag & "','S','002','" & Text10 & "','" & IIf(oCP01 = "CFP", IIf(bolIsTransReserve = True, "249104", "413101"), IIf(oCP01 = "CPS", IIf(bolIsTransReserve = True, "249104", "413101"), IIf(oCP01 = "P", IIf(bolIsTransReserve = True, "249102", "411103"), IIf(oCP01 = "PS", IIf(bolIsTransReserve = True, "249102", "411103"), IIf(oCP01 = "CFT", IIf(bolIsTransReserve = True, "249103", "412101"), IIf(oCP01 = "CFC", IIf(bolIsTransReserve = True, "249103", "412101"), IIf(oCP01 = "S", IIf(bolIsTransReserve = True, "249103", "412101"), IIf(bolIsTransReserve = True, "249101", "410103")))))))) & "' " & _
                                         "," & IIf(bolIsTransReserve = True, "Null", " '" & IIf(oCP01 = "CFP", "CFP", IIf(oCP01 = "CPS", "CFP", IIf(oCP01 = "P", "P", IIf(oCP01 = "PS", "P", IIf(oCP01 = "CFT", "CFT", IIf(oCP01 = "CFC", "CFT", IIf(oCP01 = "S", "CFT", "T"))))))) & "' ") & ",0," & Text9.Text & ",'" & Text3 & "/" & StrToStr(ChgSQL(Text2), 8) & "/結餘','" & Text1 & "','" & Text11 & "','" & Text3 & "'," & MaskEdBox2.ClipText & ",'" & tmpSNGp & "')"
               End If
               If Abs(Text5.Text) > 0 Then
                  adoTaie.Execute "insert into acc1p0 (a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p15,a1p16,a1p17,a1p18,a1p30) " & _
                                         " values ('" & Text16.Tag & "','S','003','" & Text10 & "','2211' " & _
                                         ",'TOT',0," & Abs(Text5.Text) & ",'" & Text3 & "/" & StrToStr(ChgSQL(Text2), 8) & "/結餘','" & Text1 & "','" & Text11 & "','" & Text3 & "'," & MaskEdBox2.ClipText & ",'" & IIf(oCP01 = "CFP", "CFP", IIf(oCP01 = "CPS", "CFP", IIf(oCP01 = "P", "CCP", IIf(oCP01 = "PS", "CCP", IIf(oCP01 = "CFT", "CFT", IIf(oCP01 = "CFC", "CFT", IIf(oCP01 = "S", "CFT", "CCT"))))))) & "' )"
               End If
            Else
               adoTaie.Execute "insert into acc1p0 (a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p15,a1p16,a1p17,a1p18,a1p30) " & _
                                         " values ('" & Text16.Tag & "','S','001','" & Text10 & "','2211' " & _
                                         ",'TOT'," & Abs(Text5.Text) & ",0,'" & Text3 & "/" & StrToStr(ChgSQL(Text2), 8) & "/結餘','" & Text1 & "','" & Text11 & "','" & Text3 & "'," & MaskEdBox2.ClipText & ",'" & IIf(oCP01 = "CFP", "CFP", IIf(oCP01 = "CPS", "CFP", IIf(oCP01 = "P", "CCP", IIf(oCP01 = "PS", "CCP", IIf(oCP01 = "CFT", "CFT", IIf(oCP01 = "CFC", "CFT", IIf(oCP01 = "S", "CFT", "CCT"))))))) & "' )"
               adoTaie.Execute "insert into acc1p0 (a1p01,a1p02,a1p03,a1p04,a1p05,a1p06,a1p07,a1p08,a1p14,a1p15,a1p16,a1p17,a1p18,a1p30) " & _
                                         " values ('" & Text16.Tag & "','S','002','" & Text10 & "','" & IIf(oCP01 = "CFP", "220106", IIf(oCP01 = "CPS", "220106", IIf(oCP01 = "P", "220112", IIf(oCP01 = "PS", "220112", IIf(oCP01 = "CFT", "220105", IIf(oCP01 = "CFC", "220105", IIf(oCP01 = "S", "220105", "220111"))))))) & "' " & _
                                         ",'TOT',0," & Abs(Text5.Text) & ",'" & Text3 & "/" & StrToStr(ChgSQL(Text2), 8) & "/結餘','" & Text1 & "','" & Text11 & "','" & Text3 & "'," & MaskEdBox2.ClipText & ",null )"
            End If
        End If
   End If
   'add by nick end
   strItemNo = Text10
   strCon1 = MaskEdBox2.Text
   strCon2 = Text9
   strCon3 = Text3
   strCon4 = Text11
   strCon5 = Text13
   tool3_enabled
   Frmacc41a2.Show
   Me.Hide
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      'Modify by Morgan 2011/6/23
      If GetAcc240(Text10) = True Then
      'end 2011/6/23
         FormShow
         AdodcRefresh
         SumShow
         RecordShow
         
      End If 'Add by Morgan 2011/6/23
   'edit by nickc 2005/11/16 新增才可以查未結算資料
   'Else
   ElseIf strSaveConfirm = MsgText(3) Then
     Acc240Query
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Activate()
On Error GoTo Checking
'Frmacc0000.Toolbar1.Buttons.Item(8).Enabled = False 'Removed by Morgan 2022/9/30
DisabledMoveRecord
   strFormName = Name
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   'Modify by Morgan 2011/6/23
   If GetAcc240(strItemNo) = True Then
      FormShow
      AdodcRefresh
      SumShow
      RecordShow
   End If
   strItemNo = MsgText(601)
Checking:
   Exit Sub
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 9780
   Me.Height = 5780
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   OpenTable
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strConTitle = MsgText(601)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc41a0 = Nothing
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
   If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
      MsgBox Label10 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox2.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
      MsgBox Label10 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox2.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Text10_GotFocus()
   'add by nickc 2008/03/12
   If Text10 = "R" Then
        Text10.SelStart = 1
   Else
        TextInverse Text10
   End If
   CloseIme  'add by sonia 2017/2/22
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'2012/6/27 modify by sonia 改用LostFocus
'Private Sub Text10_Validate(Cancel As Boolean)
'   If strSaveConfirm = MsgText(3) Then
'      If Text10 = "" Then
'         Exit Sub
'      End If
'      Acc240Query
'      'AdodcRefresh 'Remove by Morgan 2009/7/20 Acc240Query 內已經有執行
'      'SumShow
'   End If
'End Sub
Private Sub Text10_LostFocus()
   If strSaveConfirm = MsgText(3) Then
      If Text10 = "" Then
         Exit Sub
      End If
      Acc240Query
      'add by sonia 2014/6/10 J公司加提醒
      If Text16.Tag = "J" Then
         MsgBox "請注意！此為智權公司案件！"
      End If
      'end 2014/6/10
   End If
End Sub
'2012/6/27 end

Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub

'*************************************************
'  依結餘單號顯示相關資料
'
'*************************************************
Public Sub TableQuery()
Dim strName As String

   Combo1.Clear
   strName = CaseNameShow(Mid(Text3, 1, Len(Text3) - 9), Mid(Text3, Len(Text3) - 8, 6), Mid(Text3, Len(Text3) - 2, 1), Mid(Text3, Len(Text3) - 1, 2), 1)
   If strName <> "" Then
      Combo1.AddItem strName
      Combo1 = strName
   Else
      Combo1 = ""
   End If
   strName = CaseNameShow(Mid(Text3, 1, Len(Text3) - 9), Mid(Text3, Len(Text3) - 8, 6), Mid(Text3, Len(Text3) - 2, 1), Mid(Text3, Len(Text3) - 1, 2), 2)
   If strName <> "" Then
      Combo1.AddItem strName
   End If
   strName = CaseNameShow(Mid(Text3, 1, Len(Text3) - 9), Mid(Text3, Len(Text3) - 8, 6), Mid(Text3, Len(Text3) - 2, 1), Mid(Text3, Len(Text3) - 1, 2), 3)
   If strName <> "" Then
      Combo1.AddItem strName
   End If
   strName = CaseCustShow(Mid(Text3, 1, Len(Text3) - 9), Mid(Text3, Len(Text3) - 8, 6), Mid(Text3, Len(Text3) - 2, 1), Mid(Text3, Len(Text3) - 1, 2), 1)
   Text1 = strName
   strName = CustomerQuery(Text1, 1)
   If strName <> "" Then
      Text2 = strName
   Else
      strName = CustomerQuery(Text1, 2)
      If strName <> "" Then
         Text2 = strName
      Else
         strName = CustomerQuery(Text1, 3)
         If strName <> "" Then
            Text2 = strName
         Else
            Text2 = ""
         End If
      End If
   End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Modify By Sindy 2014/2/21
'Private Sub Text11_Validate(Cancel As Boolean)
Public Sub Text11_Validate(Cancel As Boolean)
'2014/2/21 END
   If Text11 = MsgText(601) Then
      MsgBox MsgText(10), , MsgText(5)
      Cancel = True
      Text11.SetFocus
      Exit Sub
   Else
      If ExistCheck("staff", "st01", Text11, Label9, False) = False Then
         MsgBox MsgText(45) & Label9, , MsgText(5)
         Cancel = True
         Text11.SetFocus
         Exit Sub
      End If
   End If
   'add by nickc 2005/08/04 加檢查不可離職
   If adoquery.State = 1 Then adoquery.Close
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select st02, st03,st04 from staff where st01 = '" & Text11 & "' ", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      'Modify By Sindy 2014/2/24
'      If CheckStr(adoquery.Fields("st04").Value) = "2" Then
'         adoquery.Close 'Add By Sindy 2014/2/21
'         MsgBox "不可輸入離職人員！", vbInformation, "錯誤！"
'         Cancel = True
'         Text11.SetFocus
'         Exit Sub
'      End If
      If IsNull(adoquery.Fields("st02").Value) Then
         Text4 = ""
      Else
         Text4 = adoquery.Fields("st02").Value
      End If
      If IsNull(adoquery.Fields("st03").Value) Then
         Text13 = ""
      Else
         Text13 = adoquery.Fields("st03").Value
         Text12 = A0902Query(Text13)
      End If
   Else
      Text4 = ""
      Text12 = ""
      Text13 = ""
   End If
   adoquery.Close
'   StaffShow
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   TableQuery
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking

   'Remove by Morgan 2011/6/23 進畫面太慢,取消上下筆功能(沒在用)

'add by nickc 2005/11/15
Dim MaxDay  As String
Dim Str41a0 As String
'抓最大結餘日
Dim oCP01 As String
Dim oCP02 As String
Dim oCP03 As String
Dim oCP04 As String

'add by nickc 2007/02/28
Dim tRS As New ADODB.Recordset

If Text3 <> "" Then
   oCP01 = Mid(Text3, 1, Len(Text3) - 9)
   oCP02 = Mid(Text3, Len(Text3) - 8, 6)
   oCP03 = Mid(Text3, Len(Text3) - 2, 1)
   oCP04 = Mid(Text3, Len(Text3) - 1, 2)
End If
MaxDay = ""
'Modify by Amy 2017/10/24 E_Fail Err
If oCP01 = "CFP" Then
   '2010/3/26 MODIFY BY SONIA 僅EPC的子案與母案合併,接續案不可合併,集體設計暫時也不合併
   'strSQL = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240002<>'" & Text10 & "'  and a240003 is null "
   strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "' and A240002<>'" & Text10 & "' and (a240003 is null or a240003=0) having Max(a240001) is not null "
ElseIf oCP01 = "TF" Then
   '2010/3/26 MODIFY BY SONIA 馬德里案母案與延土延伸分開算
   'strSQL = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "'  and A240002<>'" & Text10 & "'  and a240003 is null "
   strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240002<>'" & Text10 & "'  and a240003 is null having Max(a240001) is not null "
Else
   strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "' and A240008='" & oCP04 & "' and A240002<>'" & Text10 & "'  and a240003 is null having Max(a240001) is not null "
End If
'end 2017/10/24
CheckOC2
Set tRS = New ADODB.Recordset
tRS.CursorLocation = adUseClient
tRS.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If Not tRS.EOF And Not tRS.BOF Then
   MaxDay = CheckStr(tRS.Fields(0))
End If
Str41a0 = ""
Dim strSQL1 As String
Dim StrSQLa As String
StrSQLa = ""
strSQL1 = ""
If oCP01 = "TF" Then
   Str41a0 = " and ax214>='" & oCP01 & oCP02 & "000' and ax214<='" & oCP01 & oCP02 & "ZZZ'  " & IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ") & " and a0205 <= " & Val(FCDate(MaskEdBox1.Text)) & " "
   strSQL1 = " and ax214>='" & oCP01 & oCP02 & "000' and ax214<='" & oCP01 & oCP02 & "ZZZ' and C2.cp01='" & oCP01 & "' and c2.cp02='" & oCP02 & "'  "
ElseIf oCP01 = "CFP" Then
   Str41a0 = " and ax214>='" & oCP01 & oCP02 & oCP03 & "00' and ax214<='" & oCP01 & oCP02 & oCP03 & "ZZ'  " & IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ") & " and a0205 <= " & Val(FCDate(MaskEdBox1.Text)) & " "
   strSQL1 = " and ax214>='" & oCP01 & oCP02 & oCP03 & "00' and ax214<='" & oCP01 & oCP02 & oCP03 & "ZZ' and C2.cp01='" & oCP01 & "' and c2.cp02='" & oCP02 & "' and c2.cp03='" & oCP03 & "' "
Else
   Str41a0 = " and ax214='" & oCP01 & oCP02 & oCP03 & oCP04 & "' " & IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ") & " and a0205 <= " & Val(FCDate(MaskEdBox1.Text)) & " "
   strSQL1 = " and ax214='" & oCP01 & oCP02 & oCP03 & oCP04 & "' and C2.cp01='" & oCP01 & "' and c2.cp02='" & oCP02 & "' and c2.cp03='" & oCP03 & "' and c2.cp04='" & oCP04 & "' "
End If
Dim NewAcc020021 As String
   '2011/3/25 MODIFY BY SONIA 剔除結餘傳票,否則第二次以上的結餘會抓到
   NewAcc020021 = "(select distinct ax202,ax214,newa1,newa2,newa3,newa4 from (" & _
                  " select ax202,ax212,ax214,sum(A1) as newA1,sum(A2) as newA2,sum(A3) as newA3,sum(A4) as newA4 from (" & _
                  " select ax202,ax212,ax214, " & _
                  " (DECODE(substr(ax205,1,1),'4',nvl(ax207,0)-nvl(ax206,0),decode(substr(ax205,1,4),'2201',decode(instr(ax212,'退費'),0,nvl(ax207,0),0),0))) as A1, " & _
                  "  (decode(substr(ax205,1,1),'4',nvl(ax207,0)-nvl(ax206,0),0)) as A2, " & _
                  " (decode(substr(ax205,1,4),'2201',decode(ax206,0,0,nvl(ax206,0)))) as A3, " & _
                  " (decode(substr(ax205,1,4),'2201',decode(ax206,0,decode(instr(ax212,'退費'),0,0,nvl(ax207,0)),0))) as A4 " & _
                  " From acc020, acc021 where  ax201=a0201(+) and ax202=a0202(+) AND INSTR(AX212,'結餘')=0 " & Str41a0 & ") NewTable " & _
                  " group by NewTable.ax202,NewTable.ax212,NewTable.ax214) NewTable2) NewTable3 "
                            
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select distinct cp09, decode(substr(pa09, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, patent, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc1u0 where ax202=a1p22(+) and a1p04=a1u01(+) and a1u03=cp09(+) and cp01||cp02||cp03||cp04='" & Text3.Text & "' and cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+)  union " & _
                  "select cp09, decode(substr(tm10, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, trademark, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc1u0 where ax202=a1p22(+) and a1p04=a1u01(+) and a1u03=cp09(+) and cp01||cp02||cp03||cp04='" & Text3.Text & "' and cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+)  union " & _
                  "select cp09, decode(substr(lc15, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, lawcase, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc1u0 where  ax202=a1p22(+) and a1p04=a1u01(+) and a1u03=cp09(+) and cp01||cp02||cp03||cp04='" & Text3.Text & "' and cp01 = lc01 and cp02 = lc02 and cp03 = lc03 and cp04 = lc04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+)  union " & _
                  "select cp09, decode(substr(sp09, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, servicepractice, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc1u0 where  ax202=a1p22(+) and a1p04=a1u01 and a1u03=cp09(+) and cp01||cp02||cp03||cp04='" & Text3.Text & "' and cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+)  union " & _
                  "select cp09, decode(substr(pa09, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, patent, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc0z0 where  ax202=a1p22(+) and a1p04=a0z01(+) and a0z02=cp60(+) and cp01||cp02||cp03||cp04='" & Text3.Text & "' and cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+)  union " & _
                  "select cp09, decode(substr(tm10, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, trademark, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc0z0 where ax202=a1p22(+) and a1p04=a0z01(+) and a0z02=cp60(+) and cp01||cp02||cp03||cp04='" & Text3.Text & "' and  cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+)  union " & _
                  "select cp09, decode(substr(lc15, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, lawcase, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc0z0 where ax202=a1p22(+) and a1p04=a0z01(+) and a0z02=cp60(+) and cp01||cp02||cp03||cp04='" & Text3.Text & "' and  cp01 = lc01 and cp02 = lc02 and cp03 = lc03 and cp04 = lc04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+)  union " & _
                  "select cp09, decode(substr(sp09, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, servicepractice, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc0z0 where ax202=a1p22(+) and a1p04=a0z01(+) and a0z02=cp60(+) and cp01||cp02||cp03||cp04='" & Text3.Text & "' and cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+)  order by cp09 asc", adoTaie, adOpenStatic, adLockReadOnly

   Set Adodc1.Recordset = adoadodc1
   OpenTableRight
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Public Sub OpenTableRight()

On Error GoTo Checking
Dim tRS As New ADODB.Recordset
Dim MaxDay  As String
Dim Str41a0 As String
'抓最大結餘日
Dim oCP01 As String
Dim oCP02 As String
Dim oCP03 As String
Dim oCP04 As String
If Text3 <> "" Then
   oCP01 = Mid(Text3, 1, Len(Text3) - 9)
   oCP02 = Mid(Text3, Len(Text3) - 8, 6)
   oCP03 = Mid(Text3, Len(Text3) - 2, 1)
   oCP04 = Mid(Text3, Len(Text3) - 1, 2)
End If
MaxDay = ""
'Modify by Amy 2017/10/24 E_Fail Err
If oCP01 = "CFP" Then
   '2010/3/26 MODIFY BY SONIA 僅EPC的子案與母案合併,接續案不可合併,集體設計暫時也不合併
   'strSQL = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240002<>'" & Text10 & "'  and a240003 is null "
   strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "' and A240002<>'" & Text10 & "'  and a240003 is null having Max(a240001) is not null "
ElseIf oCP01 = "TF" Then
   strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "'  and A240002<>'" & Text10 & "'  and a240003 is null having Max(a240001) is not null "
Else
   strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "' and A240008='" & oCP04 & "' and A240002<>'" & Text10 & "'  and a240003 is null having Max(a240001) is not null "
End If
'end 2017/10/24
CheckOC2
Set tRS = New ADODB.Recordset
tRS.CursorLocation = adUseClient
tRS.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If Not tRS.EOF And Not tRS.BOF Then
   MaxDay = CheckStr(tRS.Fields(0))
End If
Str41a0 = ""
If oCP01 = "TF" Then
   Str41a0 = " and ax214>='" & oCP01 & oCP02 & "000' and ax214<='" & oCP01 & oCP02 & "ZZZ'  " & IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ")
ElseIf oCP01 = "CFP" Then
   Str41a0 = " and ax214>='" & oCP01 & oCP02 & oCP03 & "00' and ax214<='" & oCP01 & oCP02 & oCP03 & "ZZ'  " & IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ")
Else
   Str41a0 = " and ax214='" & oCP01 & oCP02 & oCP03 & oCP04 & "' " & IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ")
End If
  If adoadodc2.State = 1 Then adoadodc2.Close
   adoadodc2.CursorLocation = adUseClient
   '2011/3/25 MODIFY BY SONIA 剔除結餘傳票,否則第二次以上的結餘會抓到
   'Modified by Morgan 2011/12/13 調整語法
   'adoadodc2.Open "select a0201, a0202,a0102, sum(decode(substr(ax205,1,4),'2201',decode(ax206,0,decode(instr(ax212,'退費'),0,0,nvl(ax207,0)) * -1,nvl(ax206,0)))) as Amount from acc021, acc020,Acc010 where ax201 = a0201 and ax202 = a0202 and (substr(ax205, 1, 4) = '2201' or substr(ax205,1,1)='4') " & Str41a0 & " and a0205 <= " & Val(FCDate(MaskEdBox1.Text)) & " and ax205=a0101 AND INSTR(AX212,'結餘')=0 and decode(substr(ax205,1,4),'2201',decode(ax206,0,decode(instr(ax212,'退費'),0,0,nvl(ax207,0)) * -1,nvl(ax206,0)))<>0  group by a0201, a0202,A0102 ", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc2.Open "select a0201, a0202,a0102, sum(decode(substr(ax205,1,4),'2201',decode(ax206,0,decode(instr(ax212,'退費'),0,0,nvl(ax207,0)) * -1,nvl(ax206,0)))) as Amount from acc021, acc020,Acc010 where ax201 = a0201(+) and ax202 = a0202(+) and (substr(ax205, 1, 4) = '2201' or substr(ax205,1,1)='4') " & Str41a0 & " and a0205 <= " & Val(FCDate(MaskEdBox1.Text)) & " and ax205=a0101(+) AND INSTR(AX212,'結餘')=0 and decode(substr(ax205,1,4),'2201',decode(ax206,0,decode(instr(ax212,'退費'),0,0,nvl(ax207,0)) * -1,nvl(ax206,0)))<>0  group by a0201, a0202,A0102 ", adoTaie, adOpenStatic, adLockReadOnly
   'adoadodc2.Requery 'Removed by Morgan 2011/12/13 上面才抓的資料無須再重新查詢
   Set Adodc2.Recordset = adoadodc2
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
'add by nickc 2007/02/08
Dim tRS As New ADODB.Recordset

On Error GoTo Checking

   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   If adoadodc2.State = adStateOpen Then
      adoadodc2.Close
   End If
'add by nickc 2005/11/15
Dim MaxDay  As String
Dim Str41a0 As String
'抓最大結餘日
Dim oCP01 As String
Dim oCP02 As String
Dim oCP03 As String
Dim oCP04 As String
If Text3 <> "" Then
oCP01 = Mid(Text3, 1, Len(Text3) - 9)
oCP02 = Mid(Text3, Len(Text3) - 8, 6)
oCP03 = Mid(Text3, Len(Text3) - 2, 1)
oCP04 = Mid(Text3, Len(Text3) - 1, 2)
End If
MaxDay = ""
'Modify by Amy 2017/10/24 E_Fail Err
If oCP01 = "CFP" Then
   '2010/3/26 MODIFY BY SONIA 僅EPC的子案與母案合併,接續案不可合併,集體設計暫時也不合併
   'strSQL = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240002<>'" & Text10 & "' and a240003 is null "
   strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "' and A240002<>'" & Text10 & "' and a240003 is null having Max(a240001) is not null "
ElseIf oCP01 = "TF" Then
   strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "'  and A240002<>'" & Text10 & "'  and a240003 is null having Max(a240001) is not null "
Else
   strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "' and A240008='" & oCP04 & "' and A240002<>'" & Text10 & "'  and a240003 is null having Max(a240001) is not null "
End If
'end 2017/10/24
CheckOC2
Set tRS = New ADODB.Recordset
tRS.CursorLocation = adUseClient
tRS.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If Not tRS.EOF And Not tRS.BOF Then
   MaxDay = CheckStr(tRS.Fields(0))
End If
Str41a0 = ""
Dim strSQL1 As String
Dim StrSQLa As String
StrSQLa = ""
strSQL1 = ""
If oCP01 = "TF" Then
   Str41a0 = " and ax214>='" & oCP01 & oCP02 & "000' and ax214<='" & oCP01 & oCP02 & "ZZZ'  " & IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ") & " and a0205 <= " & Val(FCDate(MaskEdBox1.Text)) & " "
   strSQL1 = " and ax214>='" & oCP01 & oCP02 & "000' and ax214<='" & oCP01 & oCP02 & "ZZZ' and C2.cp01='" & oCP01 & "' and c2.cp02='" & oCP02 & "'  "
ElseIf oCP01 = "CFP" Then
   Str41a0 = " and ax214>='" & oCP01 & oCP02 & oCP03 & "00' and ax214<='" & oCP01 & oCP02 & oCP03 & "ZZ'  " & IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ") & " and a0205 <= " & Val(FCDate(MaskEdBox1.Text)) & " "
   strSQL1 = " and ax214>='" & oCP01 & oCP02 & oCP03 & "00' and ax214<='" & oCP01 & oCP02 & oCP03 & "ZZ' and C2.cp01='" & oCP01 & "' and c2.cp02='" & oCP02 & "' and c2.cp03='" & oCP03 & "' "
Else
   Str41a0 = " and ax214='" & oCP01 & oCP02 & oCP03 & oCP04 & "' " & IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ") & " and a0205 <= " & Val(FCDate(MaskEdBox1.Text)) & " "
   strSQL1 = " and ax214='" & oCP01 & oCP02 & oCP03 & oCP04 & "' and C2.cp01='" & oCP01 & "' and c2.cp02='" & oCP02 & "' and c2.cp03='" & oCP03 & "' and c2.cp04='" & oCP04 & "' "
End If
Dim NewAcc020021 As String
   '2011/3/25 MODIFY BY SONIA 剔除結餘傳票,否則第二次以上的結餘會抓到
   NewAcc020021 = "(select distinct ax202,ax214,newa1,newa2,newa3,newa4 from (" & _
                  " select ax202,ax212,ax214,sum(A1) as newA1,sum(A2) as newA2,sum(A3) as newA3,sum(A4) as newA4 from (" & _
                  " select ax202,ax212,ax214, " & _
                  " (DECODE(substr(ax205,1,1),'4',nvl(ax207,0)-nvl(ax206,0),decode(substr(ax205,1,4),'2201',decode(instr(ax212,'退費'),0,nvl(ax207,0),0),0))) as A1, " & _
                  "  (decode(substr(ax205,1,1),'4',nvl(ax207,0)-nvl(ax206,0),0)) as A2, " & _
                  " (decode(substr(ax205,1,4),'2201',decode(ax206,0,0,nvl(ax206,0)))) as A3, " & _
                  " (decode(substr(ax205,1,4),'2201',decode(ax206,0,decode(instr(ax212,'退費'),0,0,nvl(ax207,0)),0))) as A4 " & _
                  " From acc020, acc021 where  ax201=a0201(+) and ax202=a0202(+) AND INSTR(AX212,'結餘')=0 " & Str41a0 & ") NewTable " & _
                  " group by NewTable.ax202,NewTable.ax212,NewTable.ax214) NewTable2) NewTable3 "
                           
   adoadodc1.CursorLocation = adUseClient
   'Modify by Morgan 2009/7/20 改語法
   'adoadodc1.Open "select distinct cp09, decode(substr(pa09, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, patent, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc1u0 where ax202=a1p22(+) and a1p04=a1u01(+) and a1u03=cp09(+) and cp01||cp02||cp03||cp04='" & Text3.Text & "' and cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+)  union " & _
                  "select cp09, decode(substr(tm10, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, trademark, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc1u0 where ax202=a1p22(+) and a1p04=a1u01(+) and a1u03=cp09(+) and cp01||cp02||cp03||cp04='" & Text3.Text & "' and cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+)  union " & _
                  "select cp09, decode(substr(lc15, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, lawcase, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc1u0 where  ax202=a1p22(+) and a1p04=a1u01(+) and a1u03=cp09(+) and cp01||cp02||cp03||cp04='" & Text3.Text & "' and cp01 = lc01 and cp02 = lc02 and cp03 = lc03 and cp04 = lc04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+)  union " & _
                  "select cp09, decode(substr(sp09, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, servicepractice, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc1u0 where  ax202=a1p22(+) and a1p04=a1u01 and a1u03=cp09(+) and cp01||cp02||cp03||cp04='" & Text3.Text & "' and cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+)  union " & _
                  "select cp09, decode(substr(pa09, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, patent, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc0z0 where  ax202=a1p22(+) and a1p04=a0z01(+) and a0z02=cp60(+) and cp01||cp02||cp03||cp04='" & Text3.Text & "' and cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+)  union " & _
                  "select cp09, decode(substr(tm10, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, trademark, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc0z0 where ax202=a1p22(+) and a1p04=a0z01(+) and a0z02=cp60(+) and cp01||cp02||cp03||cp04='" & Text3.Text & "' and  cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+)  union " & _
                  "select cp09, decode(substr(lc15, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, lawcase, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc0z0 where ax202=a1p22(+) and a1p04=a0z01(+) and a0z02=cp60(+) and cp01||cp02||cp03||cp04='" & Text3.Text & "' and  cp01 = lc01 and cp02 = lc02 and cp03 = lc03 and cp04 = lc04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+)  union " & _
                  "select cp09, decode(substr(sp09, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, servicepractice, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc0z0 where ax202=a1p22(+) and a1p04=a0z01(+) and a0z02=cp60(+) and cp01||cp02||cp03||cp04='" & Text3.Text & "' and cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+)  order by cp09 asc", adoTaie, adOpenStatic, adLockReadOnly

   'Adodc1.Recordset.Requery
   adoadodc1.Open "select distinct cp09, decode(substr(pa09, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, patent, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc1u0 where cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=cp13 and a1u03(+)=cp09 and a1p04(+)=a1u01 and ax202(+)=a1p22 and ax202 is not null  union " & _
                  "select cp09, decode(substr(tm10, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, trademark, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc1u0 where cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=cp13 and a1u03(+)=cp09 and a1p04(+)=a1u01 and ax202(+)=a1p22 and ax202 is not null  union " & _
                  "select cp09, decode(substr(lc15, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, lawcase, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc1u0 where  cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04 and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=cp13 and a1u03(+)=cp09 and a1p04(+)=a1u01 and ax202(+)=a1p22 and ax202 is not null union " & _
                  "select cp09, decode(substr(sp09, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, servicepractice, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc1u0 where  cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=cp13 and a1u03(+)=cp09 and a1p04(+)=a1u01 and ax202(+)=a1p22 and ax202 is not null  union " & _
                  "select cp09, decode(substr(pa09, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, patent, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc0z0 where cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=cp13 and a0z02(+)=cp60 and a1p04(+)=a0z01 and ax202(+)=a1p22 and ax202 is not null  union " & _
                  "select cp09, decode(substr(tm10, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, trademark, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc0z0 where cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=cp13 and a0z02(+)=cp60 and a1p04(+)=a0z01 and ax202(+)=a1p22 and ax202 is not null  union " & _
                  "select cp09, decode(substr(lc15, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, lawcase, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc0z0 where cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04 and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=cp13 and a0z02(+)=cp60 and a1p04(+)=a0z01 and ax202(+)=a1p22 and ax202 is not null  union " & _
                  "select cp09, decode(substr(sp09, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, newa1 as Ramount, newa2 as Samount from  caseprogress, servicepractice, casepropertymap, staff," & NewAcc020021 & ",acc1p0,acc0z0 where cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=cp13 and a0z02(+)=cp60 and a1p04(+)=a0z01 and ax202(+)=a1p22 and ax202 is not null  order by cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
   
   Set DataGrid1.DataSource = Adodc1 '重新連結比重新執行語法快
   'end 2009/7/20
   
   '2015/3/13 add by sonia 讀取FC代理人欄,有FC代理人欄之案件,智權人員設定為總所 CFP023092,P103938
   If Text10 <> "" Then
      If Left(PUB_GetStaffST15(Text11, 1), 1) <> "S" Then   'add by sonia 2016/5/16 智權人員為智權部者一律不管是否有FC代理人欄CFP-015488(瑞婷)
         If GetPrjPeopleNum6(oCP01 & "-" & oCP02 & "-" & oCP03 & "-" & oCP04) <> "" Then
            Text11 = "M0100"
         End If
      End If                                                'add by sonia 2016/5/16
   End If
   '2015/3/13 end
   
   'add by sonia 2017/6/19 若屬於業績列入P1001之專利處人員則智權人員改為P1001
   If adoquery.State = 1 Then adoquery.Close
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * FROM SetSpecMan where ocode='P1001' and instr(oman,'" & Text11 & "')>0 ", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      Text11 = "P1001"
   End If
   adoquery.Close
   'end 2017/6/19
   
   OpenTableRight
'   adoadodc2.CursorLocation = adUseClient
'   adoadodc2.Open "select a0201, a0202, a0102, sum(ax206 - ax207) as Amount from acc021, acc020, acc010 where ax201 = a0201 and ax202 = a0202 and ax205 = a0101 and substr(ax205, 1, 4) = '2201' and ax214 = '" & Text3 & "' and a0205 > " & Val(FCDate(MaskEdBox1.Text)) & " group by a0201, a0202, a0102", adoTaie, adOpenStatic, adLockReadOnly
'   Adodc2.Recordset.Requery
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示欄位資料(結餘資料)
'
'*************************************************
Public Sub FormShow()
'add by sonia 2025/5/23
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'end 2025/5/23

   Text10 = adoacc240.Fields("A240002").Value
   Text10.Tag = Text10 'Add by Morgan 2011/6/23
   If IsNull(adoacc240.Fields("A240005").Value) And IsNull(adoacc240.Fields("A240006").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = adoacc240.Fields("A240005").Value & adoacc240.Fields("A240006").Value & adoacc240.Fields("A240007").Value & adoacc240.Fields("A240008").Value
      TableQuery
   End If
   If IsNull(adoacc240.Fields("A241006").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = adoacc240.Fields("A241006").Value
   End If
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc240.Fields("A240001").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(adoacc240.Fields("A240001").Value)
   End If
   MaskEdBox1.Mask = DFormat
   If IsNull(adoacc240.Fields("A240010").Value) Then
      Text11 = MsgText(601)
   Else
      Text11 = adoacc240.Fields("A240010").Value
      StaffShow
   End If
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(adoacc240.Fields("A240015").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(adoacc240.Fields("A240015").Value)
   End If
   MaskEdBox2.Mask = DFormat
   If IsNull(adoacc240.Fields("a240011").Value) Then
      Text14 = MsgText(601)
   Else
      Text14 = adoacc240.Fields("a240011").Value
   End If
   If IsNull(adoacc240.Fields("a240012").Value) Then
      Text15 = MsgText(601)
   Else
      Text15 = adoacc240.Fields("a240012").Value
   End If
   SpecialCompShow 'Add by Amy 2013/12/17
   'add by sonia 2025/5/23 若分錄已產生傳票A1P22 is not null則鎖住結算日期欄，以免資料不一致
   StrSQLa = "select distinct a1p22 from acc1p0 where a1p04 = '" & Text10 & "' and a1p22 is not null"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      MaskEdBox2.Enabled = False
   Else
      MaskEdBox2.Enabled = True
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   'end 2025/5/23
End Sub

'*************************************************
'  顯示智權人員相關資料
'
'*************************************************
Public Sub StaffShow()
   'add by nickc 2005/11/15
   If adoquery.State = 1 Then adoquery.Close
   adoquery.CursorLocation = adUseClient
   'Modify By Sindy 2014/2/24
   'adoquery.Open "select st02, st03 from staff where st01 = '" & Text11 & "' and st04 = '1'", adoTaie, adOpenStatic, adLockReadOnly
   adoquery.Open "select st02, st03 from staff where st01 = '" & Text11 & "'", adoTaie, adOpenStatic, adLockReadOnly
   '2014/2/24 END
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields("st02").Value) Then
         Text4 = ""
      Else
         Text4 = adoquery.Fields("st02").Value
      End If
      If IsNull(adoquery.Fields("st03").Value) Then
         Text13 = ""
      Else
         Text13 = adoquery.Fields("st03").Value
         Text12 = A0902Query(Text13)
      End If
   Else
      Text4 = ""
      Text12 = ""
      Text13 = ""
   End If
   adoquery.Close
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = adoacc240.Bookmark & MsgText(35) & adoacc240.RecordCount
End Sub

'*************************************************
'  顯示欄位資料(傳票合計)
'
'*************************************************
Public Sub SumShow()
Dim tRS As New ADODB.Recordset
Dim MaxDay  As String
Dim Str41a0 As String
'抓最大結餘日
Dim oCP01 As String
Dim oCP02 As String
Dim oCP03 As String
Dim oCP04 As String
   
   If Text3 <> "" Then
      oCP01 = Mid(Text3, 1, Len(Text3) - 9)
      oCP02 = Mid(Text3, Len(Text3) - 8, 6)
      oCP03 = Mid(Text3, Len(Text3) - 2, 1)
      oCP04 = Mid(Text3, Len(Text3) - 1, 2)
   End If
   MaxDay = ""
   'Modify by Amy 2017/10/24 E_Fail Err
   If oCP01 = "CFP" Then
      '2010/3/26 MODIFY BY SONIA 僅EPC的子案與母案合併,接續案不可合併,集體設計暫時也不合併
      'strSQL = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240002<>'" & Text10 & "'  and a240003 is null "
      strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "' and A240002<>'" & Text10 & "'  and a240003 is null having Max(a240001) is not null "
   ElseIf oCP01 = "TF" Then
      strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "'  and A240002<>'" & Text10 & "'  and a240003 is null having Max(a240001) is not null "
   Else
      strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "' and A240008='" & oCP04 & "' and A240002<>'" & Text10 & "'  and a240003 is null having Max(a240001) is not null "
   End If
   CheckOC2
   Set tRS = New ADODB.Recordset
   tRS.CursorLocation = adUseClient
   tRS.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not tRS.EOF And Not tRS.BOF Then
      MaxDay = CheckStr(tRS.Fields(0))
   End If
   Str41a0 = ""
   If oCP01 = "TF" Then
      Str41a0 = " and ax214>='" & oCP01 & oCP02 & "000' and ax214<='" & oCP01 & oCP02 & "ZZZ'  " & IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ")
   ElseIf oCP01 = "CFP" Then
      Str41a0 = " and ax214>='" & oCP01 & oCP02 & oCP03 & "00' and ax214<='" & oCP01 & oCP02 & oCP03 & "ZZ'  " & IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ")
   Else
      Str41a0 = " and ax214='" & oCP01 & oCP02 & oCP03 & oCP04 & "' " & IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ")
   End If
   
   If adoaccsum.State = 1 Then adoaccsum.Close
   adoaccsum.CursorLocation = adUseClient
   'edit by nickc 2005/11/15
   'adoaccsum.Open "select sum(Amount) from (select a0201, a0202,  sum(decode(substr(ax205,1,4),'2201',decode(ax206,0,decode(instr(ax212,'退費'),0,0,nvl(ax207,0)) * -1,nvl(ax206,0)))) as Amount from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and (substr(ax205, 1, 4) = '2201' or substr(ax205,1,1)='4') " & Str41a0 & " and a0205 <= " & Val(FCDate(MaskEdBox1.Text)) & " group by a0201, a0202) new", adoTaie, adOpenStatic, adLockReadOnly
   '2011/3/25 MODIFY BY SONIA 剔除結餘傳票,否則第二次以上的結餘會抓到
   'Modified by Morgan 2011/12/13 調整語法
'   adoaccsum.Open "select sum(Amount),sum(InAmount),sum(OutAmount) from (select a0201, a0202,sum(DECODE(substr(ax205,1,1),'4',nvl(ax207,0)-nvl(ax206,0),decode(substr(ax205,1,4),'2201',decode(instr(ax212,'退費'),0,nvl(ax207,0),0),0))) as INAmount,sum(decode(substr(ax205,1,1),'4',nvl(ax207,0)-nvl(ax206,0),0)) as OutAmount,  sum(decode(substr(ax205,1,4),'2201',decode(ax206,0,decode(instr(ax212,'退費'),0,0,nvl(ax207,0)) * -1,nvl(ax206,0)))) as Amount from acc021, acc020 where ax201 = a0201 and ax202 = a0202 " & Str41a0 & " and a0205 <= " & Val(FCDate(MaskEdBox1.Text)) & " AND INSTR(AX212,'結餘')=0 group by a0201, a0202) new", adoTaie, adOpenStatic, adLockReadOnly
   adoaccsum.Open "select sum(Amount),sum(InAmount),sum(OutAmount) from (select a0201, a0202,sum(DECODE(substr(ax205,1,1),'4',nvl(ax207,0)-nvl(ax206,0),decode(substr(ax205,1,4),'2201',decode(instr(ax212,'退費'),0,nvl(ax207,0),0),0))) as INAmount,sum(decode(substr(ax205,1,1),'4',nvl(ax207,0)-nvl(ax206,0),0)) as OutAmount,  sum(decode(substr(ax205,1,4),'2201',decode(ax206,0,decode(instr(ax212,'退費'),0,0,nvl(ax207,0)) * -1,nvl(ax206,0)))) as Amount from acc021, acc020 where ax201 = a0201(+) and ax202 = a0202(+) " & Str41a0 & " and a0205 <= " & Val(FCDate(MaskEdBox1.Text)) & " AND INSTR(AX212,'結餘')=0 group by a0201, a0202) new", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text7 = MsgText(601)
      Else
         Text7 = adoaccsum.Fields(0).Value
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text6 = MsgText(601)
      Else
         Text6 = adoaccsum.Fields(1).Value
      End If
      If IsNull(adoaccsum.Fields(2).Value) Then
         Text8 = MsgText(601)
      Else
         Text8 = adoaccsum.Fields(2).Value
      End If
   Else
      Text7 = MsgText(601)
      Text6 = MsgText(601)
      Text8 = MsgText(601)
   End If
   adoaccsum.Close
   If Command1.Enabled = True Then Command1.SetFocus    '2012/6/27 add by sonia
End Sub

Private Sub Text4_Change()
   Text9 = Format(Val(Text6) - Val(Text8) - Val(Text7) - Val(Text5), FAmount)
   'add by Val(Text6) - Val(Text8) - Val(Text7)nickc 2005/10/31 新增或修改時
   If strSaveConfirm = MsgText(4) Or strSaveConfirm = MsgText(3) Then
      If Val(Text9) = 0 And Val(Text5) = 0 Then Command1.Enabled = False Else Command1.Enabled = True
   Else
      Command1.Enabled = False
   End If
End Sub

Private Sub Text6_Change()
   Text9 = Format(Val(Text6) - Val(Text8) - Val(Text7) - Val(Text5), FAmount)
   'add by nickc 2005/10/31 新增或修改時
   If strSaveConfirm = MsgText(4) Or strSaveConfirm = MsgText(3) Then
      If Val(Text9) = 0 And Val(Text5) = 0 Then Command1.Enabled = False Else Command1.Enabled = True
   Else
      Command1.Enabled = False
   End If
End Sub

Private Sub Text7_Change()
   Text9 = Format(Val(Text6) - Val(Text8) - Val(Text7) - Val(Text5), FAmount)
   If strSaveConfirm = MsgText(4) Or strSaveConfirm = MsgText(3) Then
      If Val(Text9) = 0 And Val(Text5) = 0 Then Command1.Enabled = False Else Command1.Enabled = True
   Else
      Command1.Enabled = False
   End If
End Sub

Private Sub Text8_Change()
   Text9 = Format(Val(Text6) - Val(Text8) - Val(Text7) - Val(Text5), FAmount)
   If strSaveConfirm = MsgText(4) Or strSaveConfirm = MsgText(3) Then
      If Val(Text9) = 0 And Val(Text5) = 0 Then Command1.Enabled = False Else Command1.Enabled = True
   Else
      Command1.Enabled = False
   End If
End Sub

'*************************************************
'  關閉分錄欄位輸入狀態
'
'*************************************************
Public Sub FormDisabled()
   Text9 = Format(Val(Text6) - Val(Text8) - Val(Text7) - Val(Text5), FAmount)
   Command1.Enabled = False
   Command2.Enabled = True
   Text10.Enabled = True
End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'
'*************************************************
Public Sub FormEnabled()
   Text9 = Format(Val(Text6) - Val(Text8) - Val(Text7) - Val(Text5), FAmount)
   If strSaveConfirm = MsgText(4) Or strSaveConfirm = MsgText(3) Then
      If Val(Text9) = 0 And Val(Text5) = 0 Then Command1.Enabled = False Else Command1.Enabled = True
      If strSaveConfirm = MsgText(4) Then
         Command2.Enabled = False
         Text10.Enabled = False
      End If
   Else
      Command1.Enabled = False
   End If
End Sub

'*************************************************
'  顯示欄位資料(未結算之結餘明細)
'
'*************************************************
Public Sub Acc240Query()
'2011/7/12 add by sonia
Dim oCP01 As String
Dim oCP02 As String
Dim oCP03 As String
Dim oCP04 As String
'2011/7/12 end
   
   If adoquery.State = 1 Then adoquery.Close 'Add By Sindy 2014/2/21
   adoquery.CursorLocation = adUseClient
   '2009/8/5 modify by sonia 離職智權人員或國外部結餘,智權人員改為M0100
   'adoquery.Open "select * from acc240,acc241 where a240002 = '" & Text10 & "' and (a240003 is null or a240003 = 0) and (a240015 is null or a240015=0) and A240002=A241001 and A241002=998 ", adoTaie, adOpenStatic, adLockReadOnly
   adoquery.Open "select * from acc240,acc241,staff where a240002 = '" & Text10 & "' and (a240003 is null or a240003 = 0) and (a240015 is null or a240015=0) and A240002=A241001 and A241002=998 and a240010=st01(+) ", adoTaie, adOpenStatic, adLockReadOnly
   '2009/8/5 END
   If adoquery.RecordCount <> 0 Then
      strAddNo = Text10 'Add by Morgan 2011/6/23
      If IsNull(adoquery.Fields("A240005").Value) And IsNull(adoquery.Fields("A240006").Value) Then
         Text3 = MsgText(601)
      Else
         Text3 = adoquery.Fields("A240005").Value & adoquery.Fields("A240006").Value & adoquery.Fields("A240007").Value & adoquery.Fields("A240008").Value
         '2011/7/12 add by sonia
         oCP01 = adoquery.Fields("A240005").Value
         oCP02 = adoquery.Fields("A240006").Value
         oCP03 = adoquery.Fields("A240007").Value
         oCP04 = adoquery.Fields("A240008").Value
         '2011/7/12 end
      End If
      If IsNull(adoquery.Fields("a241006").Value) Then
         Text5 = MsgText(601)
      Else
         Text5 = adoquery.Fields("a241006").Value
      End If
      MaskEdBox1.Mask = MsgText(601)
      If IsNull(adoquery.Fields("a240001").Value) Then
         MaskEdBox1.Text = MsgText(601)
      Else
         MaskEdBox1.Text = CFDate(adoquery.Fields("a240001").Value)
      End If
      MaskEdBox1.Mask = DFormat
      If IsNull(adoquery.Fields("a240010").Value) Then
         Text11 = MsgText(601)
      Else
         Text11 = adoquery.Fields("a240010").Value
      End If
      '2009/8/5 add by sonia
      If CheckStr(adoquery.Fields("st04").Value) = "2" Or Mid(adoquery.Fields("st03").Value, 1, 1) = "F" Then
         Text11 = "M0100"
      End If
      '2009/8/5 end
      'add by sonia 2023/10/3 智權部暫存區且非無效客戶者皆轉總所
      If Mid(adoquery.Fields("st03").Value, 1, 1) = "S" And CheckStr(adoquery.Fields("a240010").Value) < "6" And Mid(adoquery.Fields("a240010").Value, 5, 1) <> "9" Then
         Text11 = "M0100"
      End If
      'end 2023/10/3
      If adoquery.Fields("a240010").Value = "A4023" Then Text11 = "M0100"    'add by sonia 2023/5/10 林文雄的一律改為M0100
      If adoquery.Fields("a240010").Value = "79075" Then Text11 = "M0100"    'add by sonia 2023/8/11 郭雅娟的一律改為M0100
      MaskEdBox2.Mask = MsgText(601)
      '2012/6/27 modify by sonia 辜說預設前一筆
      'If IsNull(adoquery.Fields("a240003").Value) Then
      '   MaskEdBox2.Text = CFDate(ACDate(ServerDate))
      'Else
      '   MaskEdBox2.Text = CFDate(adoquery.Fields("a240003").Value)
      'End If
      If Not IsNull(adoquery.Fields("a240015").Value) Then
         MaskEdBox2.Text = CFDate(adoquery.Fields("a240015").Value)
      End If
      '2012/6/27 end
      MaskEdBox2.Mask = DFormat
      If IsNull(adoquery.Fields("a240011").Value) Then
         Text14 = MsgText(601)
      Else
         Text14 = adoquery.Fields("a240011").Value
      End If
      If IsNull(adoquery.Fields("a240012").Value) Then
         Text15 = MsgText(601)
      Else
         Text15 = adoquery.Fields("a240012").Value
      End If
      '2011/7/12 ADD BY SONIA 檢查若非本案最後結餘單則提醒作之,R096030071已有R100010046
      adoquery.Close
      adoquery.CursorLocation = adUseClient
      If oCP01 = "CFP" Then
         adoquery.Open "select * from acc240 where a240005 = '" & oCP01 & "' and a240006 = '" & oCP02 & "' and a240007 = '" & oCP03 & "' and A240002<>'" & Text10 & "' and (a240003 is null or a240003 = 0) and (a240015 is null or a240015=0)  ", adoTaie, adOpenStatic, adLockReadOnly
      ElseIf Text3 = "TF" Then
         adoquery.Open "select * from acc240 where a240005 = '" & oCP01 & "' and a240006 = '" & oCP02 & "' and A240002<>'" & Text10 & "' and (a240003 is null or a240003 = 0) and (a240015 is null or a240015=0)  ", adoTaie, adOpenStatic, adLockReadOnly
      Else
         adoquery.Open "select * from acc240 where a240005 = '" & oCP01 & "' and a240006 = '" & oCP02 & "' and a240007 = '" & oCP03 & "' and A240008='" & oCP04 & "' and A240002<>'" & Text10 & "' and (a240003 is null or a240003 = 0) and (a240015 is null or a240015=0)  ", adoTaie, adOpenStatic, adLockReadOnly
      End If
      If Not adoquery.EOF Then
         MsgBox "此案尚有其他結餘單, 此單不可結算, 請先作廢 !", , MsgText(5)
         adoquery.Close
         Exit Sub
      End If
      '2011/7/12 END
   Else
      strAddNo = "" 'Add by Morgan 2011/6/23
      Text3 = ""
      Text5 = ""
      MaskEdBox1.Mask = ""
      MaskEdBox1.Text = ""
      MaskEdBox1.Mask = DFormat
      Text11 = ""
      MaskEdBox2.Mask = ""        '2012/6/27 cancel by sonia 辜說不要清
      'MaskEdBox2.Text = ""        '2012/6/27 cancel by sonia 辜說不要清
      MaskEdBox2.Mask = DFormat   '2012/6/27 cancel by sonia 辜說不要清
      Text1 = ""
      Text2 = ""
      Combo1.Clear
      Text4 = ""
      Text13 = ""
      Text12 = ""
      'edit by nickc 2008/03/12
      'Text10 = ""
      Text10 = "R"
      MsgBox MsgText(33), , MsgText(5)
      'add by nickc 2008/03/12
      Text10.SetFocus
      Text10.SelStart = 1
      
      adoquery.Close
      Exit Sub
   End If
   adoquery.Close
   AdodcRefresh
   'OpenTableLeft
   OpenTableRight
   TableQuery
   SumShow
   StaffShow
   SpecialCompShow 'Add by Amy 2013/12/17
End Sub

Public Sub CheckBalance()
Dim stSQL As String, iR As Integer
Dim adoRst As ADODB.Recordset
'add by sonia 2018/5/14
Dim oCP01 As String
Dim oCP02 As String
Dim oCP03 As String
Dim oCP04 As String
'end 2018/5/14

   If Text3 <> "" Then
      oCP01 = Mid(Text3, 1, Len(Text3) - 9)
      oCP02 = Mid(Text3, Len(Text3) - 8, 6)
      oCP03 = Mid(Text3, Len(Text3) - 2, 1)
      oCP04 = Mid(Text3, Len(Text3) - 1, 2)
   End If
   'Modify by Morgan 2011/6/23 +還要加上本結餘單的規費
   'modify by sonia 2018/5/14 本所案號讀ACC021要分TF,CFP
   'stSQL = "select sum(ax206)-sum(ax207) from (select ax206,ax207 from acc021 WHERE AX214='" & Text3 & "' AND ax205 in ('220105','220106','220111','220112')" & _
      " union all select a1p07,a1p08 from acc1p0 where a1p04='" & Text10 & "' and a1p05 in ('220105','220106','220111','220112'))"
   If oCP01 = "TF" Then
      stSQL = "select sum(ax206)-sum(ax207) from (select ax206,ax207 from acc021 WHERE AX214>='" & oCP01 & oCP02 & "' AND AX214<='" & oCP01 & oCP02 & "ZZZ' AND ax205 in ('220105','220106','220111','220112')"
   ElseIf oCP01 = "CFP" Then
      stSQL = "select sum(ax206)-sum(ax207) from (select ax206,ax207 from acc021 WHERE AX214>='" & oCP01 & oCP02 & "' AND AX214<='" & oCP01 & oCP02 & oCP03 & "ZZ' AND ax205 in ('220105','220106','220111','220112')"
   Else
      stSQL = "select sum(ax206)-sum(ax207) from (select ax206,ax207 from acc021 WHERE AX214='" & oCP01 & oCP02 & oCP03 & oCP04 & "' AND ax205 in ('220105','220106','220111','220112')"
   End If
   stSQL = stSQL & " union all select a1p07,a1p08 from acc1p0 where a1p04='" & Text10 & "' and a1p05 in ('220105','220106','220111','220112'))"
   'end 2018/5/14
   iR = 1
   Set adoRst = ClsLawReadRstMsg(iR, stSQL)
   If iR = 1 Then
      'Modified by Morgan 2021/7/16 改不平都要彈訊息--辜
      'If adoRst(0) > 0 Then
      '   MsgBox "本案規費餘額負 " & Format(adoRst(0), DDollar) & " !!"
      'end 2021/7/16
      If adoRst(0) <> 0 Then
         MsgBox "本案規費不平, 餘額為 " & Format(adoRst(0), DDollar) & " !!"
      End If
   End If
   Set adoRst = Nothing
End Sub

'Add by Morgan 2011/6/23
Public Function GetAcc240(pNo As String) As Boolean
   strExc(0) = "select * from acc240,Acc241 where a240002='" & pNo & "' and (a240003 is null or a240003 = 0) and (a240015 is not null or a240015<>0) and A240002=A241001 and A241002=998 order by a240002 asc"
   intI = 1
   Set adoacc240 = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      GetAcc240 = True
   End If
End Function

'Add by Morgan 2011/6/23
'取消查詢功能
Public Sub DisabledMoveRecord()
   Frmacc0000.Toolbar1.Buttons.Item(13).Enabled = False
   Frmacc0000.Toolbar1.Buttons.Item(14).Enabled = False
   Frmacc0000.Toolbar1.Buttons.Item(15).Enabled = False
   Frmacc0000.Toolbar1.Buttons.Item(16).Enabled = False
End Sub

'Add by Morgan 2011/6/23
Public Function EditCheck() As Boolean
   If Text10.Tag = "" Then
      MsgBox "請先查詢資料後再作業!!"
   ElseIf Text10.Tag <> Text10 Then
      MsgBox "單號已異動，請重新查詢資料後再作業!!"
   Else
      EditCheck = True
   End If
End Function

'Add by Morgan 2011/6/23
Public Function AddCheck() As Boolean
   If strAddNo = "" Then
      MsgBox "請先查詢資料後再作業!!"
   ElseIf strAddNo <> Text10 Then
      MsgBox "單號已異動，請重新查詢資料後再作業!!"
   Else
      AddCheck = True
   End If
End Function

'Add by Amy 2013/12/17
'特殊出名公司為null 或 T 公司別=1, 特殊出名公司為J 公司別=J
Private Sub SpecialCompShow()
    Dim strCompNo As String
    strCompNo = ""
    If Val(strSrvDate(1)) >= Val(InvoiceStartDate) Then
        Text16 = GetSpecialComp(Mid(Text3, 1, Len(Text3) - 9), Mid(Text3, Len(Text3) - 8, 6), Mid(Text3, Len(Text3) - 2, 1), Mid(Text3, Len(Text3) - 1, 2), strCompNo, 6)
        Text16.Tag = strCompNo
    Else
        Text16 = "1-台一國際專利"
        Text16.Tag = "1"
    End If
End Sub

'將寫於.bas 的function 搬回
Public Sub FormClear()
'      If .MaskEdBox2.Text = MsgText(29) Then
'         .MaskEdBox2.Text = CFDate(ACDate(ServerDate))
'      End If
'      .MaskEdBox2.SetFocus
      Text10 = ""
      Text3 = ""
      MaskEdBox1.Mask = ""
      MaskEdBox1.Text = ""
      MaskEdBox1.Mask = DFormat
      Text5 = ""
      Text11 = ""
      Text1 = ""
      MaskEdBox2.Mask = ""
'     .MaskEdBox2.Text = ""    '2012/3/27 cancel by sonia 帶前一筆
      MaskEdBox2.Mask = DFormat
      Text2 = ""
      Combo1.Clear
      Text4 = ""
      Text13 = ""
      Text12 = ""
      Text10.SetFocus
      Text14 = ""
      Text15 = ""
      Text7 = ""
      Text6 = ""
      Text8 = ""
      Text9 = ""
      Text16 = "" 'Add by Amy 2013/12/17
      AdodcRefresh
End Sub

Public Sub MoveLastRecord()
    If adoacc240.RecordCount <> 0 Then
         adoacc240.MoveLast
         FormShow
         AdodcRefresh
         SumShow
         RecordShow
      End If
End Sub

Public Sub MoveFirstRecord()
    If adoacc240.RecordCount <> 0 Then
         adoacc240.MoveFirst
         FormShow
         AdodcRefresh
         SumShow
         RecordShow
      End If
End Sub

Public Sub MoveNextRecord()
    If adoacc240.EOF = False Then
         adoacc240.MoveNext
         If adoacc240.EOF Then
            adoacc240.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         FormShow
         AdodcRefresh
         SumShow
         RecordShow
      End If
End Sub

Public Sub MovePreviousRecord()
    If adoacc240.BOF = False Then
         adoacc240.MovePrevious
         If adoacc240.BOF Then
            adoacc240.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         FormShow
         AdodcRefresh
         SumShow
         RecordShow
      End If
End Sub

Public Sub FormSave()
On Error GoTo Checking
'''edit by nickc 2007/08/24 將檢查放在 acc_var 以免資料檢查不過，又被 commit 掉
'''      If .Text10 = MsgText(601) Then
'''         MsgBox MsgText(10), , MsgText(5)
'''         strControlButton = MsgText(602)
'''         .Text10.SetFocus
'''         Exit Sub
'''      Else
'''         If .Text11 = MsgText(601) Then
'''            MsgBox MsgText(10), , MsgText(5)
'''            strControlButton = MsgText(602)
'''            .Text11.SetFocus
'''            Exit Sub
'''         Else
'''            If ExistCheck("staff", "st01", .Text11, .Label9) = False Then
'''               MsgBox MsgText(45) & .Label9, , MsgText(5)
'''               strControlButton = MsgText(602)
'''               .Text11.SetFocus
'''               Exit Sub
'''            End If
'''         End If
'''         If .MaskEdBox2.Text = MsgText(601) Or .MaskEdBox2.Text = MsgText(29) Then
'''            MsgBox .Label10 & MsgText(52), , MsgText(5)
'''            strControlButton = MsgText(602)
'''            .MaskEdBox2.SetFocus
'''            Exit Sub
'''         End If
'''         If DateCheck(.MaskEdBox2.Text) = MsgText(603) Then
'''            MsgBox .Label10 & MsgText(63), , MsgText(5)
'''            strControlButton = MsgText(602)
'''            .MaskEdBox2.SetFocus
'''            Exit Sub
'''         End If
'''      End If
      adoTaie.Execute "update acc240 set a240015=" & Val(FCDate(MaskEdBox2.Text)) & ",A240016='" & strUserNum & "',A240010='" & Text11 & "' where A240002='" & Text10 & "' "
      'modify by sonia 2025/5/21 分錄未產生傳票
      If GetAcc240(Text10) = True Then
         Text10.Tag = Text10
         AdodcRefresh
         RecordShow
      End If
      'end 2011/7/7
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub
'end 2013/12/17
