VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc41b0 
   AutoRedraw      =   -1  'True
   Caption         =   "CF案件結餘作廢作業"
   ClientHeight    =   5110
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   9420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5110
   ScaleWidth      =   9420
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
      TabIndex        =   33
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
      Left            =   4050
      TabIndex        =   31
      Top             =   975
      Width           =   2745
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
      Left            =   4020
      TabIndex        =   17
      Top             =   255
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
      Height          =   315
      Left            =   1320
      TabIndex        =   15
      Top             =   600
      Width           =   1572
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
      TabIndex        =   14
      Top             =   3888
      Width           =   1572
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
      Left            =   3408
      TabIndex        =   13
      Top             =   3480
      Width           =   1095
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
      Left            =   7896
      TabIndex        =   12
      Top             =   3504
      Width           =   1140
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
      Left            =   7680
      TabIndex        =   11
      Top             =   3912
      Width           =   1572
   End
   Begin VB.TextBox Text11 
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
      TabIndex        =   10
      Top             =   4248
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Height          =   300
      Left            =   2520
      Picture         =   "Frmacc41b0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   240
      Width           =   350
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
      Height          =   315
      Left            =   2910
      TabIndex        =   7
      Top             =   4608
      Width           =   1555
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
      TabIndex        =   6
      Top             =   4608
      Width           =   1575
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
      Left            =   4608
      TabIndex        =   5
      Top             =   3480
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc41b0.frx":0102
      Height          =   1755
      Left            =   240
      TabIndex        =   3
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
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Samount"
         Caption         =   "已作收入金額"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1179.78
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
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
            ColumnWidth     =   1289.764
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   18
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
      Left            =   7680
      TabIndex        =   2
      Top             =   4272
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
      Bindings        =   "Frmacc41b0.frx":0117
      Height          =   1770
      Left            =   6135
      TabIndex        =   4
      Top             =   1665
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   3122
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
            Format          =   "#,##0"
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
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1149.732
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   255
      Top             =   2235
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   6165
      Top             =   2040
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
      TabIndex        =   32
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
   Begin MSForms.TextBox Text2 
      Height          =   300
      Left            =   5640
      TabIndex        =   16
      Top             =   255
      Width           =   3705
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
      Height          =   315
      Left            =   2910
      TabIndex        =   9
      Top             =   600
      Width           =   6435
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "11351;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text4 
      Height          =   300
      Left            =   2910
      TabIndex        =   8
      Top             =   4245
      Width           =   1560
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      Size            =   "2743;529"
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
      TabIndex        =   34
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
      TabIndex        =   30
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
      TabIndex        =   29
      Top             =   1365
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4488
      Visible         =   0   'False
      Width           =   132
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
      TabIndex        =   28
      Top             =   240
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
      TabIndex        =   27
      Top             =   240
      Width           =   975
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
      TabIndex        =   26
      Top             =   600
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
      TabIndex        =   25
      Top             =   3888
      Width           =   972
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
      TabIndex        =   24
      Top             =   960
      Width           =   975
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
      Left            =   2568
      TabIndex        =   23
      Top             =   3480
      Width           =   492
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
      Left            =   6120
      TabIndex        =   22
      Top             =   3912
      Width           =   1452
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
      TabIndex        =   21
      Top             =   4245
      Width           =   900
   End
   Begin VB.Label Label10 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "作廢日期"
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
      Left            =   6600
      TabIndex        =   20
      Top             =   4272
      Width           =   972
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
      TabIndex        =   19
      Top             =   4608
      Width           =   852
   End
End
Attribute VB_Name = "Frmacc41b0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/25 Form2.0已修改 Text2/Text4/Text5/Combo1/DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc240 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoadodc2 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset

Private Sub Command2_Click()
   If strSaveConfirm <> MsgText(3) Then
      If adoacc240.RecordCount = 0 Or Text10 = MsgText(601) Then
         Exit Sub
      End If
      'edit by nickc 2005/07/28
      adoacc240.Find "A240002 = '" & Text10 & "'", 0, adSearchForward, 1
      If adoacc240.EOF Then
         MsgBox MsgText(33), , MsgText(5)
         adoacc240.MoveFirst
      End If
      FormShow
      AdodcRefresh
      SumShow
      RecordShow
   Else
     Acc240Query
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Activate()
On Error GoTo Checking
Frmacc0000.Toolbar1.Buttons.Item(8).Enabled = False
   strFormName = Name
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   'edit by nickc 2005/07/29
   adoacc240.Find "a240002 = '" & strItemNo & "'", 0, adSearchForward, 1
   If adoacc240.EOF = False Then
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
   Me.Width = 9640
   Me.Height = 5670
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
   Set Frmacc41b0 = Nothing
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
'edit by nickc 2005/07/28 存檔時會檢查
'   If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
'      MsgBox Label10 & MsgText(52), , MsgText(5)
'      Cancel = True
'      MaskEdBox2.SetFocus
'      Exit Sub
'   End If
If MaskEdBox2.Text = MsgText(29) Then Exit Sub
   If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
      MsgBox Label10 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox2.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
   CloseIme  'add by sonia 2017/2/22
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If strSaveConfirm = MsgText(3) Then
      If Text10 = "" Then
         Exit Sub
      End If
      Acc240Query
      'AdodcRefresh 'Remove by Morgan 2009/7/21 Acc240Query 已有執行
      'SumShow
   End If
End Sub

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

Private Sub Text11_Validate(Cancel As Boolean)
   StaffShow
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
   adoacc240.CursorLocation = adUseClient
   'edit by nickc 2005/07/28
   adoacc240.Open "select * from acc240,Acc241 where (a240003 is not null or a240003 <> 0) and A240002=A241001 and A241002=998 order by a240002 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   OpenTableLeft
   OpenTableRight
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Public Sub OpenTableLeft()
   If adoadodc1.State = 1 Then adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
'edit by nickc 2005/07/28
   adoadodc1.Open "select cp09, decode(substr(pa09, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, (nvl(cp75, 0) - nvl(cp78, 0)) as Ramount, nvl(cp73, 0) as Samount from  caseprogress, patent, casepropertymap, staff where  cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+) and cp59 = '" & Text10 & "' union " & _
                  "select cp09, decode(substr(tm10, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, (nvl(cp75, 0) - nvl(cp78, 0)) as Ramount, nvl(cp73, 0) as Samount from  caseprogress, trademark, casepropertymap, staff where  cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+) and cp59 = '" & Text10 & "' union " & _
                  "select cp09, decode(substr(lc15, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, (nvl(cp75, 0) - nvl(cp78, 0)) as Ramount, nvl(cp73, 0) as Samount from  caseprogress, lawcase, casepropertymap, staff where cp01 = lc01 and cp02 = lc02 and cp03 = lc03 and cp04 = lc04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+) and cp59 = '" & Text10 & "' union " & _
                  "select cp09, nvl(cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, (nvl(cp75, 0) - nvl(cp78, 0)) as Ramount, nvl(cp73, 0) as Samount from  caseprogress, hirecase, casepropertymap, staff where cp01 = hc01 and cp02 = hc02 and cp03 = hc03 and cp04 = hc04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+) and cp59 = '" & Text10 & "' union " & _
                  "select cp09, decode(substr(sp09, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, (nvl(cp75, 0) - nvl(cp78, 0)) as Ramount, nvl(cp73, 0) as Samount from  caseprogress, servicepractice, casepropertymap, staff where cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+) and cp59 = '" & Text10 & "' order by cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc1.Requery
   Set Adodc1.Recordset = adoadodc1
End Sub

Public Sub OpenTableRight()

On Error GoTo Checking
Dim tRS As New ADODB.Recordset
Dim MaxDay  As String
'add by nickc 2006/01/17
Dim SecDay As String
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
   strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240002<>'" & Text10 & "' having Max(a240001) is not null "
ElseIf oCP01 = "TF" Then
   strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "'  and A240002<>'" & Text10 & "' having Max(a240001) is not null "
Else
   strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "' and A240008='" & oCP04 & "' and A240002<>'" & Text10 & "' having Max(a240001) is not null "
End If
'end 2017/10/24
CheckOC2
Set tRS = New ADODB.Recordset
tRS.CursorLocation = adUseClient
tRS.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If Not tRS.EOF And Not tRS.BOF Then
   MaxDay = CheckStr(tRS.Fields(0))     '最後一次結餘日期
End If
'add by nickc 2005/01/17
SecDay = ""
'Modify by Amy 2017/10/24 E_Fail Err
If MaxDay <> "" Then
    If oCP01 = "CFP" Then
       strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240002<>'" & Text10 & "' and A240001<>" & MaxDay & " and A240003 is null having Max(a240001) is not null "
    ElseIf oCP01 = "TF" Then
       strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "'  and A240002<>'" & Text10 & "' and A240001<>" & MaxDay & " and A240003 is null having Max(a240001) is not null "
    Else
       strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "' and A240008='" & oCP04 & "' and A240002<>'" & Text10 & "' and A240001<>" & MaxDay & " and A240003 is null having Max(a240001) is not null "
    End If
    'end 2017/10/24
    CheckOC2
    Set tRS = New ADODB.Recordset
    tRS.CursorLocation = adUseClient
    tRS.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If Not tRS.EOF And Not tRS.BOF Then
       SecDay = CheckStr(tRS.Fields(0))     '最後一次結餘日期的上一次
    End If
End If
Str41a0 = ""
If oCP01 = "TF" Then
   'edit by nickc 2006/01/17
   'Str41a0 = " and ax214>='" & oCP01 & oCP02 & "000' and ax214<='" & oCP01 & oCP02 & "ZZZ'  " & IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ")
   Str41a0 = " and ax214>='" & oCP01 & oCP02 & "000' and ax214<='" & oCP01 & oCP02 & "ZZZ'  " & IIf(Trim(SecDay) = "", "", " and a0205>=" & SecDay & " ") & IIf(Trim(MaxDay) = "", "", " and a0205<=" & MaxDay & " ")
ElseIf oCP01 = "CFP" Then
   'edit by nickc 2006/01/17
   'Str41a0 = " and ax214>='" & oCP01 & oCP02 & oCP03 & "00' and ax214<='" & oCP01 & oCP02 & oCP03 & "ZZ'  " & IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ")
   Str41a0 = " and ax214>='" & oCP01 & oCP02 & oCP03 & "00' and ax214<='" & oCP01 & oCP02 & oCP03 & "ZZ'  " & IIf(Trim(SecDay) = "", "", " and a0205>=" & SecDay & " ") & IIf(Trim(MaxDay) = "", "", " and a0205<=" & MaxDay & " ")
Else
   'edit by nickc 206/01/17
   'Str41a0 = " and ax214='" & oCP01 & oCP02 & oCP03 & oCP04 & "' " & IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ")
   Str41a0 = " and ax214='" & oCP01 & oCP02 & oCP03 & oCP04 & "' " & IIf(Trim(SecDay) = "", "", " and a0205>=" & SecDay & " ") & IIf(Trim(MaxDay) = "", "", " and a0205<=" & MaxDay & " ")
End If
  If adoadodc2.State = 1 Then adoadodc2.Close
   adoadodc2.CursorLocation = adUseClient
   'edit by nickc 2006/01/17
   'adoadodc2.Open "select a0201, a0202,a0102, sum(decode(substr(ax205,1,4),'2201',decode(ax206,0,decode(instr(ax212,'退費'),0,0,nvl(ax207,0)) * -1,nvl(ax206,0)))) as Amount from acc021, acc020,Acc010 where ax201 = a0201 and ax202 = a0202 and (substr(ax205, 1, 4) = '2201' or substr(ax205,1,1)='4') " & Str41a0 & " and a0205 <= " & Val(FCDate(MaskEdBox1.Text)) & " and ax205=a0101 and decode(substr(ax205,1,4),'2201',decode(ax206,0,decode(instr(ax212,'退費'),0,0,nvl(ax207,0)) * -1,nvl(ax206,0)))<>0  group by a0201, a0202,A0102 ", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc2.Open "select a0201, a0202,a0102, sum(decode(substr(ax205,1,4),'2201',decode(ax206,0,decode(instr(ax212,'退費'),0,0,nvl(ax207,0)) * -1,nvl(ax206,0)))) as Amount from acc021, acc020,Acc010 where ax201 = a0201 and ax202 = a0202 and (substr(ax205, 1, 4) = '2201' or substr(ax205,1,1)='4') " & Str41a0 & " and ax205=a0101 and decode(substr(ax205,1,4),'2201',decode(ax206,0,decode(instr(ax212,'退費'),0,0,nvl(ax207,0)) * -1,nvl(ax206,0)))<>0  group by a0201, a0202,A0102 ", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc2.Requery
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
On Error GoTo Checking
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   If adoadodc2.State = adStateOpen Then
      adoadodc2.Close
   End If
'   adoadodc1.CursorLocation = adUseClient
''edit by nickc 2005/07/28
'   adoadodc1.Open "select cp09, decode(substr(pa09, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, (nvl(cp75, 0) - nvl(cp78, 0)) as Ramount, nvl(cp73, 0) as Samount from  caseprogress, patent, casepropertymap, staff where  cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+) and cp59 = '" & Text10 & "' union " & _
'                  "select cp09, decode(substr(tm10, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, (nvl(cp75, 0) - nvl(cp78, 0)) as Ramount, nvl(cp73, 0) as Samount from  caseprogress, trademark, casepropertymap, staff where  cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+) and cp59 = '" & Text10 & "' union " & _
'                  "select cp09, decode(substr(lc15, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, (nvl(cp75, 0) - nvl(cp78, 0)) as Ramount, nvl(cp73, 0) as Samount from  caseprogress, lawcase, casepropertymap, staff where  cp01 = lc01 and cp02 = lc02 and cp03 = lc03 and cp04 = lc04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+) and cp59 = '" & Text10 & "' union " & _
'                  "select cp09, nvl(cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, (nvl(cp75, 0) - nvl(cp78, 0)) as Ramount, nvl(cp73, 0) as Samount from  caseprogress, hirecase, casepropertymap, staff where  cp01 = hc01 and cp02 = hc02 and cp03 = hc03 and cp04 = hc04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+) and cp59 = '" & Text10 & "' union " & _
'                  "select cp09, decode(substr(sp09, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, (nvl(cp75, 0) - nvl(cp78, 0)) as Ramount, nvl(cp73, 0) as Samount from  caseprogress, servicepractice, casepropertymap, staff where  cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+) and cp59 = '" & Text10 & "' order by cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
'
'   Adodc1.Recordset.Requery
   OpenTableLeft
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
   Text10 = adoacc240.Fields("A240002").Value
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
   If IsNull(adoacc240.Fields("A240003").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(adoacc240.Fields("A240003").Value)
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
End Sub

'*************************************************
'  顯示智權人員相關資料
'
'*************************************************
Public Sub StaffShow()
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select st02, st03 from staff where st01 = '" & Text11 & "' and st04 = '1'", adoTaie, adOpenStatic, adLockReadOnly
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
'  顯示欄位資料(結餘合計)
'
'*************************************************
Public Sub SumShow()

Dim tRS As New ADODB.Recordset
Dim MaxDay  As String
'add by nickc 2006/01/17
Dim SecDay As String
Dim Str41a0 As String
   adoaccsum.CursorLocation = adUseClient
'edit by nickc 2005/07/28
   adoaccsum.Open "select sum(Ramount), sum(Samount) from (select cp09, decode(substr(pa09, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, (nvl(cp75, 0) - nvl(cp78, 0)) as Ramount, nvl(cp73, 0) as Samount from  caseprogress, patent, casepropertymap, staff where  cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+) and cp59 = '" & Text10 & "' union " & _
                  "select cp09, decode(substr(tm10, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, (nvl(cp75, 0) - nvl(cp78, 0)) as Ramount, nvl(cp73, 0) as Samount from  caseprogress, trademark, casepropertymap, staff where  cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+) and cp59 = '" & Text10 & "' union " & _
                  "select cp09, decode(substr(lc15, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, (nvl(cp75, 0) - nvl(cp78, 0)) as Ramount, nvl(cp73, 0) as Samount from  caseprogress, lawcase, casepropertymap, staff where  cp01 = lc01 and cp02 = lc02 and cp03 = lc03 and cp04 = lc04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+) and cp59 = '" & Text10 & "' union " & _
                  "select cp09, nvl(cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, (nvl(cp75, 0) - nvl(cp78, 0)) as Ramount, nvl(cp73, 0) as Samount from  caseprogress, hirecase, casepropertymap, staff where  cp01 = hc01 and cp02 = hc02 and cp03 = hc03 and cp04 = hc04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+) and cp59 = '" & Text10 & "' union " & _
                  "select cp09, decode(substr(sp09, 1, 2), '00', cpm03, cpm04) as property, st02, nvl(cp05 - 19110000, 0) as Rdate, (nvl(cp75, 0) - nvl(cp78, 0)) as Ramount, nvl(cp73, 0) as Samount from  caseprogress, servicepractice, casepropertymap, staff where  cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp13 = st01 (+) and cp59 = '" & Text10 & "') new", adoTaie, adOpenStatic, adLockReadOnly

   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text6 = MsgText(601)
      Else
         Text6 = adoaccsum.Fields(0).Value
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text8 = MsgText(601)
      Else
         Text8 = adoaccsum.Fields(1).Value
      End If
   Else
      Text6 = MsgText(601)
      Text8 = MsgText(601)
   End If
'抓最大結餘日
Dim oCP01 As String
Dim oCP02 As String
Dim oCP03 As String
Dim oCP04 As String
If Trim(Text3) <> "" Then
oCP01 = Mid(Text3, 1, Len(Text3) - 9)
oCP02 = Mid(Text3, Len(Text3) - 8, 6)
oCP03 = Mid(Text3, Len(Text3) - 2, 1)
oCP04 = Mid(Text3, Len(Text3) - 1, 2)
End If
MaxDay = ""
'Modify by Amy 2014/10/24 E_Fail Err
If oCP01 = "CFP" Then
   strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240002<>'" & Text10 & "' having Max(a240001) is not null "
ElseIf oCP01 = "TF" Then
   strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "'  and A240002<>'" & Text10 & "' having Max(a240001) is not null "
Else
   strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "' and A240008='" & oCP04 & "' and A240002<>'" & Text10 & "' having Max(a240001) is not null "
End If
'end 2017/10/24
CheckOC2
Set tRS = New ADODB.Recordset
tRS.CursorLocation = adUseClient
tRS.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If Not tRS.EOF And Not tRS.BOF Then
   MaxDay = CheckStr(tRS.Fields(0))
End If
'add by nickc 2005/01/17
SecDay = ""
'Modify by Amy 2017/10/24 E_Fail Err
If MaxDay <> "" Then
    If oCP01 = "CFP" Then
       strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240002<>'" & Text10 & "' and A240001<>" & MaxDay & " and A240003 is null having Max(a240001) is not null "
    ElseIf oCP01 = "TF" Then
       strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "'  and A240002<>'" & Text10 & "' and A240001<>" & MaxDay & " and A240003 is null having Max(a240001) is not null "
    Else
       strSql = "SELECT MAX(a240001) FROM ACC240 WHERE a240005='" & oCP01 & "' and A240006='" & oCP02 & "' and A240007='" & oCP03 & "' and A240008='" & oCP04 & "' and A240002<>'" & Text10 & "' and A240001<>" & MaxDay & " and A240003 is null having Max(a240001) is not null "
    End If
    'end 2017/10/24
    CheckOC2
    Set tRS = New ADODB.Recordset
    tRS.CursorLocation = adUseClient
    tRS.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If Not tRS.EOF And Not tRS.BOF Then
       SecDay = CheckStr(tRS.Fields(0))     '最後一次結餘日期的上一次
    End If
End If
Str41a0 = ""
If oCP01 = "TF" Then
   'edit by nickc 2006/01/17
   'Str41a0 = " and ax214>='" & oCP01 & oCP02 & "000' and ax214<='" & oCP01 & oCP02 & "ZZZ'  " & IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ")
   Str41a0 = " and ax214>='" & oCP01 & oCP02 & "000' and ax214<='" & oCP01 & oCP02 & "ZZZ'  " & IIf(Trim(SecDay) = "", "", " and a0205>=" & SecDay & " ") & IIf(Trim(MaxDay) = "", "", " and a0205<=" & MaxDay & " ")
ElseIf oCP01 = "CFP" Then
   'edit by nickc 2006/01/17
   'Str41a0 = " and ax214>='" & oCP01 & oCP02 & oCP03 & "00' and ax214<='" & oCP01 & oCP02 & oCP03 & "ZZ'  " & IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ")
   Str41a0 = " and ax214>='" & oCP01 & oCP02 & oCP03 & "00' and ax214<='" & oCP01 & oCP02 & oCP03 & "ZZ'  " & IIf(Trim(SecDay) = "", "", " and a0205>=" & SecDay & " ") & IIf(Trim(MaxDay) = "", "", " and a0205<=" & MaxDay & " ")
Else
   'edit by nickc 206/01/17
   'Str41a0 = " and ax214='" & oCP01 & oCP02 & oCP03 & oCP04 & "' " & IIf(Trim(MaxDay) = "", "", " and a0205>=" & MaxDay & " ")
   Str41a0 = " and ax214='" & oCP01 & oCP02 & oCP03 & oCP04 & "' " & IIf(Trim(SecDay) = "", "", " and a0205>=" & SecDay & " ") & IIf(Trim(MaxDay) = "", "", " and a0205<=" & MaxDay & " ")
End If
'若是已作廢資料，數字代資料庫的
If strSaveConfirm <> MsgText(3) Then
   Text6 = adoacc240.Fields("A241003").Value
   Text8 = adoacc240.Fields("A241004").Value
End If
   adoaccsum.Close
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(Amount) from (select a0201, a0202,  sum(decode(substr(ax205,1,4),'2201',decode(ax206,0,decode(instr(ax212,'退費'),0,0,nvl(ax207,0)) * -1,nvl(ax206,0)))) as Amount from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and (substr(ax205, 1, 4) = '2201' or substr(ax205,1,1)='4') " & Str41a0 & " and a0205 <= " & Val(FCDate(MaskEdBox1.Text)) & " group by a0201, a0202) new", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text7 = MsgText(601)
      Else
         Text7 = adoaccsum.Fields(0).Value
      End If
   Else
      Text7 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

Private Sub Text4_Change()
   Text9 = Val(Text6) - Val(Text8) - Val(Text7) - Val(Text5)
End Sub

Private Sub Text6_Change()
   Text9 = Val(Text6) - Val(Text8) - Val(Text7) - Val(Text5)
End Sub

Private Sub Text7_Change()
   Text9 = Val(Text6) - Val(Text8) - Val(Text7) - Val(Text5)
End Sub

Private Sub Text8_Change()
   Text9 = Val(Text6) - Val(Text8) - Val(Text7) - Val(Text5)
End Sub

'*************************************************
'  顯示欄位資料(未作廢之結餘明細)
'
'*************************************************
Public Sub Acc240Query()
   If adoquery.State <> adStateClosed Then adoquery.Close
   
   adoquery.CursorLocation = adUseClient
   'edit by nickc 2005/07/29
   adoquery.Open "select * from acc240,acc241 where a240002 = '" & Text10 & "' and (a240003 is null or a240003 = 0) and A240002=A241001 and A241002=998 ", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      'Add by Morgan 2009/7/21 已結算不可再作廢
      If Not IsNull(adoquery.Fields("a240015")) Then
         MsgBox "【" & Text10 & "】已結算不可作廢！"
         FormReset
         adoquery.Close
         Exit Sub
      End If
      
      If IsNull(adoquery.Fields("A240005").Value) And IsNull(adoquery.Fields("A240006").Value) Then
         Text3 = MsgText(601)
      Else
         Text3 = adoquery.Fields("A240005").Value & adoquery.Fields("A240006").Value & adoquery.Fields("A240007").Value & adoquery.Fields("A240008").Value
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
      MaskEdBox2.Mask = MsgText(601)
      If IsNull(adoquery.Fields("a240003").Value) Then
         MaskEdBox2.Text = CFDate(ACDate(ServerDate))
      Else
         MaskEdBox2.Text = CFDate(adoquery.Fields("a240003").Value)
      End If
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
   Else
      FormReset
      MsgBox MsgText(33), , MsgText(5)
      adoquery.Close
      Exit Sub
   End If
   adoquery.Close
   OpenTableLeft
   OpenTableRight
   TableQuery
   SumShow
   StaffShow
   SpecialCompShow 'Add by Amy 2013/12/17
End Sub

Private Sub FormReset()
   Text3 = ""
   Text5 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   Text11 = ""
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text1 = ""
   Text2 = ""
   Combo1.Clear
   Text4 = ""
   Text13 = ""
   Text12 = ""
   Text10 = ""
End Sub

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

'將寫於.bas 的function搬回
Public Sub FormClear()
      Text10 = ""
      Text3 = ""
      MaskEdBox1.Mask = ""
      MaskEdBox1.Text = ""
      MaskEdBox1.Mask = DFormat
      Text5 = ""
      Text11 = ""
      Text1 = ""
      MaskEdBox2.Mask = ""
      MaskEdBox2.Text = ""
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
If Text10 = MsgText(601) Then
         MsgBox MsgText(10), , MsgText(5)
         strControlButton = MsgText(602)
         Text10.SetFocus
         Exit Sub
      Else
         If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
            MsgBox Label10 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            MaskEdBox2.SetFocus
            Exit Sub
         End If
         If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
            MsgBox Label10 & MsgText(63), , MsgText(5)
            strControlButton = MsgText(602)
            MaskEdBox2.SetFocus
            Exit Sub
         End If
      End If
      'edit by nickc 2005/07/28
      adoTaie.Execute "update acc240 set a240003 = " & Val(FCDate(MaskEdBox2.Text)) & ",A240016='" & strUserNum & "'  where a240002 = '" & Text10 & "'"
      adoacc240.Requery
      adoacc240.MoveFirst
      adoacc240.Find "a240002 = '" & Text10 & "'", 0, adSearchForward, 1
      If adoacc240.EOF Then
         adoacc240.MoveFirst
      End If
      FormShow
      AdodcRefresh
      'edit by nickc 2005/07/28
      'Do While .adoquery.EOF = False
         'adoTaie.Execute "update caseprogress set cp59 = null where cp09 = '" & .adoquery.Fields("a1t02").Value & "'"
         adoTaie.Execute "update caseprogress set cp59 = null where cp59 = '" & Text10 & "'"
      '   .adoquery.MoveNext
      'Loop
'      .adoquery.Close
      OpenTableLeft
      SumShow
      RecordShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub
'end 2013/12/17
