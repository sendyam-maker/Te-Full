VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc3150 
   AutoRedraw      =   -1  'True
   Caption         =   "退票作業"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5475
   ScaleWidth      =   8790
   Begin VB.CommandButton Command2 
      Height          =   300
      Left            =   2940
      Picture         =   "Frmacc3150.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   570
      Width           =   350
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
      Left            =   6840
      TabIndex        =   31
      Top             =   2700
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
      TabIndex        =   28
      Top             =   2700
      Width           =   1572
   End
   Begin VB.TextBox Text10 
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
      Left            =   4110
      TabIndex        =   26
      Top             =   1980
      Width           =   1572
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4080
      TabIndex        =   5
      Top             =   900
      Width           =   1572
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc3150.frx":0102
      Height          =   1995
      Left            =   240
      TabIndex        =   7
      Top             =   3150
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3545
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
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "退票資料"
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "a0e02"
         Caption         =   "票據號碼"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "a0e01"
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
      BeginProperty Column02 
         DataField       =   "a0e07"
         Caption         =   "銀行帳號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "a0e13"
         Caption         =   "開票日期"
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
      BeginProperty Column04 
         DataField       =   "a0e10"
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
      BeginProperty Column05 
         DataField       =   "a0e03"
         Caption         =   "單據號碼"
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
         DataField       =   "a0e11"
         Caption         =   "票據金額"
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
      BeginProperty Column07 
         DataField       =   "a0e06"
         Caption         =   "往來對象"
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
         DataField       =   "cu04"
         Caption         =   "客戶名稱"
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
      BeginProperty Column10 
         DataField       =   "a0e15"
         Caption         =   "退票日期"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1874.835
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            ColumnWidth     =   1289.764
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   240
      Top             =   3030
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
   Begin VB.TextBox Text5 
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
      Left            =   6840
      TabIndex        =   22
      Top             =   2340
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.TextBox Text6 
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
      Left            =   4110
      MaxLength       =   10
      TabIndex        =   1
      Top             =   240
      Width           =   1565
   End
   Begin VB.TextBox Text4 
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
      TabIndex        =   18
      Top             =   2340
      Width           =   1572
   End
   Begin VB.TextBox Text3 
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
      TabIndex        =   16
      Top             =   1620
      Width           =   1572
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
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   0
      Top             =   240
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
      TabIndex        =   2
      Top             =   570
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   4110
      TabIndex        =   12
      Top             =   1620
      Width           =   1575
      _ExtentX        =   2778
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
      Left            =   1320
      TabIndex        =   14
      Top             =   1980
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   300
      Left            =   1320
      TabIndex        =   4
      Top             =   900
      Width           =   1575
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
   Begin MSForms.TextBox Text12 
      Height          =   315
      Left            =   2910
      TabIndex        =   29
      Top             =   2700
      Width           =   2775
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "4895;556"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text9 
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Top             =   1230
      Width           =   7095
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "12509;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text8 
      Height          =   315
      Left            =   2910
      TabIndex        =   20
      Top             =   2310
      Width           =   2775
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "4895;556"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   315
      Left            =   5670
      TabIndex        =   19
      Top             =   240
      Width           =   2775
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "4895;556"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "託收帳號"
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
      Left            =   5880
      TabIndex        =   30
      Top             =   2700
      Width           =   975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "託收銀行"
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
      TabIndex        =   27
      Top             =   2700
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "備註"
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
      TabIndex        =   25
      Top             =   1260
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "後續處理"
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
      TabIndex        =   24
      Top             =   900
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "退票日期"
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
      Left            =   390
      TabIndex        =   23
      Top             =   900
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "票別"
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
      TabIndex        =   21
      Top             =   1980
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "往來對象"
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
      TabIndex        =   17
      Top             =   2340
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "票據金額"
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
      TabIndex        =   15
      Top             =   1620
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "到期日期"
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
      TabIndex        =   13
      Top             =   1980
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "開票日期"
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
      TabIndex        =   11
      Top             =   1620
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      Height          =   2952
      Left            =   240
      Top             =   120
      Width           =   8292
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "票據號碼"
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
      TabIndex        =   10
      Top             =   240
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收票帳號"
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
      TabIndex        =   9
      Top             =   570
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "收票銀行"
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
      TabIndex        =   8
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc3150"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/19 Form2.0已修改 Text7/Text8/Text9/DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc0e0 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adocheck As New ADODB.Recordset

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Command2_Click()
   'Modify by Amy 2020/07/17 +Text1不可為空及PKey 判斷
   If Adodc1.Recordset.RecordCount = 0 Or Text2 = MsgText(601) Or Text6 = MsgText(601) Or Text1 = MsgText(601) Then
      Exit Sub
   End If
   Adodc1.Recordset.Find "a0e01 = '" & Text6 & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      'Adodc1.Recordset.Find "a0e02 = '" & Text2 & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
      Adodc1.Recordset.Find "PKey = '" & Text2 & Text6 & Text1 & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
   'end 2020/07/17
      If Adodc1.Recordset.EOF = False Then
         '2014/2/26 add by sonia 0014386,011010075
         If "" & Adodc1.Recordset.Fields("a0e04").Value = "P" Then
            MsgBox MsgText(33), , MsgText(5)
            Adodc1.Recordset.MoveFirst
            Exit Sub
         End If
         '2014/2/26 end
         FormShow
         AdodcRefresh
         RecordShow
      Else
         MsgBox MsgText(33), , MsgText(5)
         Adodc1.Recordset.MoveFirst
      End If
   Else
      MsgBox MsgText(33), , MsgText(5)
      Adodc1.Recordset.MoveFirst
   End If
End Sub

Private Sub Command2_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command2_Click
         Exit Sub
   End Select
   KeyDefine KeyCode
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   FormShow
   RecordShow
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strCompanyNo = MsgText(601) Then
      Exit Sub
   End If
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.MoveFirst
   End If
   Adodc1.Recordset.Find "a0e01 = '" & strCompanyNo & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      'Modify by Amy 2020/07/17 +PKey 因a0e07改為key,PKey For Find
      'Adodc1.Recordset.Find "a0e02 = '" & strItemNo & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
      Adodc1.Recordset.Find "PKey = '" & strItemNo & strCompanyNo & strBankAcc & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
      If Adodc1.Recordset.EOF = False Then
         FormShow
         RecordShow
      End If
   End If
   strCompanyNo = MsgText(601)
   strBankAcc = MsgText(601) 'Add by Amy 2020/07/17
End Sub

'Add by Amy 2021/10/19
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
   'Modify by Amy 2023/10/11 原:W8850 H5600
   Me.Width = 8910
   Me.Height = 5940
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   Combo1.AddItem ComboItem(31)
   Combo1.AddItem ComboItem(32)
   OpenTable
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.MoveLast
      Adodc1.Recordset.MoveFirst
      RecordShow
   End If
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
   strTrackMode = "" 'Add by Amy 2021/10/19 Form2.0 記錄鍵盤傳入順序(清除)
   Set Frmacc3150 = Nothing
End Sub

Private Sub MaskEdBox3_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub MaskEdBox3_Validate(Cancel As Boolean)
   If MaskEdBox3.Text = MsgText(601) Or MaskEdBox3.Text = MsgText(29) Then
      MsgBox Label8 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox3.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox3.Text) = MsgText(603) Then
      MsgBox Label8 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox3.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         MaskEdBox3.SetFocus
         Exit Sub
   End Select
   KeyDefine KeyCode
   KeyEnter KeyCode
End Sub

'Add by Amy 2020/07/17
Private Sub Text1_Validate(Cancel As Boolean)
    If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
        Exit Sub
    End If
    adocheck.CursorLocation = adUseClient
    adocheck.Open "select a0e01, a0e02 from acc0e0 where a0e02 = '" & Text2 & "' and a0e01 = '" & Text6 & "' And a0e07='" & Text1 & "' and (a0e15 = 0 or a0e15 is null) and (a0e17 = 0 or a0e17 is null) and a0e04 = '" & MsgText(18) & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
    KeyDefine vbKeyF12
    adocheck.Close
End Sub

Private Sub Text11_Change()
   If Text11 = MsgText(601) Then
      Exit Sub
   End If
   Text12 = A0g02Query(Text11)
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   'Modify by Morgan 2008/1/11 新增才帶否則修改時因已有退票日期會清掉收票銀行
   'If Text2 <> MsgText(601) Then
   If strSaveConfirm = "A" And Text2 <> MsgText(601) Then
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select a0e01, a0e02 from acc0e0 where a0e02 = '" & Text2 & "' and (a0e15 = 0 or a0e15 is null) and (a0e17 = 0 or a0e17 is null) and a0e04 = '" & MsgText(18) & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If adocheck.RecordCount <> 0 Then
         If IsNull(adocheck.Fields(0).Value) Then
            Text6 = MsgText(601)
         Else
            Text6 = adocheck.Fields(0).Value
            KeyDefine vbKeyF12
         End If
      Else
         Text6 = MsgText(601)
      End If
      adocheck.Close
   End If
End Sub

Private Sub Text4_Change()
   If Text4 = MsgText(601) Then
      Exit Sub
   End If
   Select Case Mid(Text5, 1, 1)
      Case Mid(ComboItem(131), 1, 1)
         Text8 = CustomerQuery(Text4, 1)
      Case Mid(ComboItem(132), 1, 1)
         Text8 = A0i02Query(Text4)
      Case Mid(ComboItem(133), 1, 1)
         Text8 = StaffQuery(Text4)
      Case Else
         Text8 = MsgText(601)
   End Select
End Sub

Private Sub Text6_Change()
   If Text6 = MsgText(601) Then
      Exit Sub
   End If
   Text7 = A0g02Query(Text6)
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc0e0.CursorLocation = adUseClient
   'Modify by Morgan 2004/11/1 加 and rownum<1
   'Modify by Amy 2020/07/17 +a0e07 因改為key
   adoacc0e0.Open "select * from acc0e0 where a0e01 = '" & Text6 & "' and a0e02 = '" & Text2 & "' And a0e07='" & Text1 & "' and rownum<1", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   'Modify by Morgan 2006/2/23 加智權人員,客戶名稱
   'adoadodc1.Open "select * from acc0e0 where a0e04 = '" & MsgText(18) & "' and a0e15 <> 0 and a0e25 = 0 order by a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modify by Amy 2020/07/17 +PKey 因a0e07改為key,用於Find
   adoadodc1.Open "select a.*,st02,cu04,a0e02||a0e01||a0e07 as PKey from acc0e0 a,acc1p0,staff,customer where a0e04 = '" & MsgText(18) & "' and a0e15 <> 0 and a0e25 = 0 and a1p09(+)=a0e02 and a1p07(+)>0 and st01(+)=a1p16 and cu01(+)=substr(a0e06,1,8) and cu02(+)=substr(a0e06,9) order by a0e01 asc,PKey asc, a0e02 asc", adoTaie, adOpenForwardOnly, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表(票據資料--退票)
'
'*************************************************
Public Sub FormShow()
   Text6 = Adodc1.Recordset.Fields("a0e01").Value
   If IsNull(Adodc1.Recordset.Fields("a0e07").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = Adodc1.Recordset.Fields("a0e07").Value
   End If
   Text2 = Adodc1.Recordset.Fields("a0e02").Value
   MaskEdBox3.Mask = MsgText(601)
   If IsNull(Adodc1.Recordset.Fields("a0e15").Value) Or Adodc1.Recordset.Fields("a0e02").Value = 0 Then
      MaskEdBox3.Text = MsgText(601)
   Else
      MaskEdBox3.Text = CFDate(Adodc1.Recordset.Fields("a0e15").Value)
   End If
   MaskEdBox3.Mask = DFormat
   If IsNull(Adodc1.Recordset.Fields("a0e33").Value) Then
      Combo1 = MsgText(601)
   Else
      Combo1 = Combo1.List(Val(Adodc1.Recordset.Fields("a0e33").Value) - 1)
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e12").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = Adodc1.Recordset.Fields("a0e12").Value
   End If
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(Adodc1.Recordset.Fields("a0e13").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(Adodc1.Recordset.Fields("a0e13").Value)
   End If
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(Adodc1.Recordset.Fields("a0e10").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(Adodc1.Recordset.Fields("a0e10").Value)
   End If
   MaskEdBox2.Mask = DFormat
   If IsNull(Adodc1.Recordset.Fields("a0e08").Value) Then
      Text10 = MsgText(601)
   Else
      Select Case Adodc1.Recordset.Fields("a0e08").Value
         Case Mid(ComboItem(11), 1, 1)
            Text10 = ComboItem(11)
         Case Mid(ComboItem(12), 1, 1)
            Text10 = ComboItem(12)
         Case Mid(ComboItem(13), 1, 1)
            Text10 = ComboItem(13)
         Case Else
            Text10 = MsgText(601)
      End Select
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e11").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = Adodc1.Recordset.Fields("a0e11").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e05").Value) Then
      Text5 = MsgText(601)
   Else
      Select Case Adodc1.Recordset.Fields("a0e05").Value
         Case "1"
            Text5 = ComboItem(91)
         Case "2"
            Text5 = ComboItem(92)
         Case "3"
            Text5 = ComboItem(93)
      End Select
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e06").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = Adodc1.Recordset.Fields("a0e06").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e19").Value) Then
      Text11 = MsgText(601)
   Else
      Text11 = Adodc1.Recordset.Fields("a0e19").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e20").Value) Then
      Text13 = MsgText(601)
   Else
      Text13 = Adodc1.Recordset.Fields("a0e20").Value
   End If
End Sub

'*************************************************
'  清除查詢顯示
'
'*************************************************
Private Sub DataClear()
   Text1 = ""
   MaskEdBox3.Mask = ""
   MaskEdBox3.Text = ""
   MaskEdBox3.Mask = DFormat
   Combo1 = ""
   Text9 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text10 = ""
   Text3 = ""
   Text5 = ""
   Text4 = ""
   Text8 = ""
   Text11 = ""
   Text12 = ""
   Text13 = ""
End Sub

'*************************************************
'  查詢顯示(票據資料)
'
'*************************************************
Private Sub DataShow()
   If IsNull(adoacc0e0.Fields("a0e07").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = adoacc0e0.Fields("a0e07").Value
   End If
   MaskEdBox3.Mask = MsgText(601)
   If IsNull(adoacc0e0.Fields("a0e15").Value) Or adoacc0e0.Fields("a0e15").Value = 0 Then
      MaskEdBox3.Text = MsgText(601)
   Else
      MaskEdBox3.Text = CFDate(adoacc0e0.Fields("a0e15").Value)
   End If
   MaskEdBox3.Mask = DFormat
   If IsNull(adoacc0e0.Fields("a0e33").Value) Then
      Combo1 = MsgText(601)
   Else
      Combo1 = Combo1.List(Val(adoacc0e0.Fields("a0e33").Value) - 1)
   End If
   If IsNull(adoacc0e0.Fields("a0e12").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = adoacc0e0.Fields("a0e12").Value
   End If
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc0e0.Fields("a0e13").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(adoacc0e0.Fields("a0e13").Value)
   End If
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(adoacc0e0.Fields("a0e10").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(adoacc0e0.Fields("a0e10").Value)
   End If
   MaskEdBox2.Mask = DFormat
   If IsNull(adoacc0e0.Fields("a0e08").Value) Then
      Text10 = MsgText(601)
   Else
      Select Case adoacc0e0.Fields("a0e08").Value
         Case Mid(ComboItem(11), 1, 1)
            Text10 = ComboItem(11)
         Case Mid(ComboItem(12), 1, 1)
            Text10 = ComboItem(12)
         Case Mid(ComboItem(13), 1, 1)
            Text10 = ComboItem(13)
         Case Else
            Text10 = MsgText(601)
      End Select
   End If
   If IsNull(adoacc0e0.Fields("a0e11").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = adoacc0e0.Fields("a0e11").Value
   End If
   If IsNull(adoacc0e0.Fields("a0e05").Value) Then
      Text5 = MsgText(601)
   Else
      Select Case adoacc0e0.Fields("a0e05").Value
         Case "1"
            Text5 = ComboItem(91)
         Case "2"
            Text5 = ComboItem(92)
         Case "3"
            Text5 = ComboItem(93)
      End Select
   End If
   If IsNull(adoacc0e0.Fields("a0e06").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = adoacc0e0.Fields("a0e06").Value
      Select Case Mid(Text5, 1, 1)
         Case Mid(ComboItem(131), 1, 1)
            Text8 = CustomerQuery(Text4, 1)
         Case Mid(ComboItem(132), 1, 1)
            Text8 = A0i02Query(Text4)
         Case Mid(ComboItem(133), 1, 1)
            Text8 = StaffQuery(Text4)
         Case Else
            Text8 = MsgText(601)
      End Select
   End If
   If IsNull(adoacc0e0.Fields("a0e19").Value) Then
      Text11 = MsgText(601)
   Else
      Text11 = adoacc0e0.Fields("a0e19").Value
   End If
   Text12 = A0g02Query(Text11)
   If IsNull(adoacc0e0.Fields("a0e20").Value) Then
      Text13 = MsgText(601)
   Else
      Text13 = adoacc0e0.Fields("a0e20").Value
   End If
End Sub

'*************************************************
'  搜尋票據資料
'
'*************************************************
Private Sub QueryAcc0e0()
   adoacc0e0.Close
   adoacc0e0.CursorLocation = adUseClient
   '2014/2/26 modify by sonia 0014386,011010075
   'adoacc0e0.Open "select * from acc0e0 where a0e01 = '" & Text6 & "' and a0e02 = '" & Text2 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modify by Amy 2020/07/17 +a0e07 因改為key
   adoacc0e0.Open "select * from acc0e0 where a0e01 = '" & Text6 & "' and a0e02 = '" & Text2 & "' And a0e07='" & Text1 & "' and a0e04 = '" & MsgText(18) & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc0e0.RecordCount <> 0 Then
      DataShow
   Else
      DataClear
   End If
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
    Dim strQ As String 'Add by Amy 2020/07/17
    
On Error GoTo Checking
   'Modify by Amy 2020/07/17 +串a0e01/a0e07 避免acc1p0資料抓太多  ex:票號:0004130/銀行:010065388
   strQ = "Select a.*,st02,cu04,a0e02||a0e01||a0e07 as PKey From acc0e0 a,acc1p0,staff,customer " & _
             "Where a0e04 = '" & MsgText(18) & "' and a1p09(+)=a0e02 and a1p10(+)=a0e01 and a1p11(+)=a0e07 and a0e15 <> 0 and a0e25 = 0 and a1p07(+)>0 " & _
             "and st01(+)=a1p16 and cu01(+)=substr(a0e06,1,8) and cu02(+)=substr(a0e06,9) "
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   'Modify by Amy 2020/07/17 +PKey 因a0e07改為key,PKey For Find
   If strConTitle = MsgText(31) Or strConTitle = MsgText(601) Then
      'Modify by Morgan 2006/2/23 加智權人員,客戶名稱
      'adoadodc1.Open "select * from acc0e0 where a0e04 = '" & MsgText(18) & "' and a0e15 <> 0 and a0e25 = 0 order by a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      adoadodc1.Open strQ & " order by a0e01 asc, a0e02 asc", adoTaie, adOpenForwardOnly, adLockReadOnly
   Else
      If strConTitle <> strCon6 And strConTitle <> strCon7 Then
         'Modify by Morgan 2006/2/23 加智權人員,客戶名稱
         'adoadodc1.Open "select * from acc0e0 where a0e04 = '" & MsgText(18) & "' and a0e15 <> 0 and a0e25 = 0 and " & strConTitle & " >= '" & strCondition1 & "' and " & strConTitle & " <= '" & strCondition2 & "' order by a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         adoadodc1.Open strQ & strConTitle & " >= '" & strCondition1 & "' and " & strConTitle & " <= '" & strCondition2 & "' order by a0e01 asc, a0e02 asc", adoTaie, adOpenForwardOnly, adLockReadOnly
      Else
         'Modify by Morgan 2006/2/23 加智權人員,客戶名稱
         'adoadodc1.Open "select * from acc0e0 where a0e04 = '" & MsgText(18) & "' and a0e15 <> 0 and a0e25 = 0 and " & strConTitle & " >= " & Val(strCondition1) & " and " & strConTitle & " <= " & Val(strCondition2) & " order by a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         adoadodc1.Open strQ & strConTitle & " >= " & Val(strCondition1) & " and " & strConTitle & " <= " & Val(strCondition2) & " order by a0e01 asc, a0e02 asc", adoTaie, adOpenForwardOnly, adLockReadOnly
      End If
   End If
   'end 2020/07/17
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      If Text6 <> MsgText(601) And Text2 <> MsgText(601) Then
         Adodc1.Recordset.Find "a0e01 = '" & Text6 & "'", 0, adSearchForward, 1
         If Adodc1.Recordset.EOF = False Then
            'Modify by Amy 2020/07/17 改抓PKey
            'Adodc1.Recordset.Find "a0e02 = '" & Text2 & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
            Adodc1.Recordset.Find "PKey = '" & Text2 & Text6 & Text1 & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
            If Adodc1.Recordset.EOF = False Then
               FormShow
               RecordShow
            End If
         End If
      End If
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   If ExistCheck("acc0g0", "a0g01", Text6, Label9) = False Then
      Cancel = True
      Exit Sub
   End If
   '2007/7/26 add by sonia 因票號0004130有二個銀行
   adocheck.CursorLocation = adUseClient
   adocheck.Open "select a0e01, a0e02 from acc0e0 where a0e02 = '" & Text2 & "' and a0e01 = '" & Text6 & "' and (a0e15 = 0 or a0e15 is null) and (a0e17 = 0 or a0e17 is null) and a0e04 = '" & MsgText(18) & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   KeyDefine vbKeyF12
   adocheck.Close
   '2007/7/26 end
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
   'add by nickc 2007/07/13 將輸入法改成使用API
   OpenIme
End Sub

'Modify by Amy 2021/10/19 原:KeyCode As Integer
Private Sub Text9_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyDefine Val(KeyCode)
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   'Add by Amy 2018/02/12 進來直接新增一筆會error
   If Adodc1.Recordset.EOF = False And Adodc1.Recordset.RecordCount > 0 Then
    Frmacc0000.StatusBar1.Panels(2).Text = Adodc1.Recordset.Bookmark & MsgText(35) & Adodc1.Recordset.RecordCount
   End If
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   'Add by Amy 2021/10/19
   Call PUB_SaveTrackMode(1, KeyCode)
    
    'Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
    If PUB_ChkTrackMode = False Then
            Exit Sub
    End If
    'end 2021/10/19
    
   Select Case KeyCode
      Case vbKeyF12
         QueryAcc0e0
   End Select
   KeyEnter KeyCode
End Sub
'add by nickc 2007/07/13 將輸入法改成使用API
Private Sub Text9_Validate(Cancel As Boolean)
CloseIme
End Sub

'Add by Amy 2020/07/17
'從aacc_sav搬回
Public Sub Frmacc3150_Save()
   On Error GoTo Checking
   With Frmacc3150
      If .Text6 = MsgText(601) Then
         MsgBox MsgText(10) & .Label9, , MsgText(5)
         strControlButton = MsgText(602)
         .Text6.SetFocus
         Exit Sub
      Else
         If .Text2 = MsgText(602) Then
            MsgBox MsgText(10) & .Label2, , MsgText(5)
            strControlButton = MsgText(602)
            .Text2.SetFocus
            Exit Sub
         End If
         If ExistCheck("acc0g0", "a0g01", .Text6, .Label9) = False Then
            strControlButton = MsgText(602)
            .Text6.SetFocus
            Exit Sub
         End If
         If .MaskEdBox3.Text = MsgText(601) Or .MaskEdBox3.Text = MsgText(29) Then
            MsgBox .Label8 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            .MaskEdBox3.SetFocus
            Exit Sub
         Else
            If DateCheck(.MaskEdBox3.Text) = MsgText(603) Then
               MsgBox .Label8 & MsgText(63), , MsgText(5)
               strControlButton = MsgText(602)
               .MaskEdBox3.SetFocus
               Exit Sub
            End If
         End If
      End If
      'Add by Amy 2021/10/19
      If PUB_ChkUniText(Me) = False Then
         strControlButton = MsgText(602)
         Exit Sub
      End If

      If strSaveConfirm = MsgText(3) Then
         .adoacc0e0.Close
         .adoacc0e0.CursorLocation = adUseClient
         'Modify by Amy 2020/07/17 +a0e07 因改為key
         .adoacc0e0.Open "select * from acc0e0 where a0e01 = '" & .Text6 & "' and a0e02 = '" & .Text2 & "' And a0e07='" & Text1 & "' and a0e15 = 0 and a0e17 = 0 and a0e21 = 0 and a0e25 = 0", adoTaie, adOpenDynamic, adLockBatchOptimistic
         If .adoacc0e0.RecordCount = 0 Then
            MsgBox MsgText(33) & " " & MsgText(39), , MsgText(5)
            Exit Sub
         End If
      Else
         .adoacc0e0.Close
         .adoacc0e0.CursorLocation = adUseClient
         'Modify by Amy 2020/07/17 +a0e07 因改為key
         .adoacc0e0.Open "select * from acc0e0 where a0e01 = '" & .Text6 & "' and a0e02 = '" & .Text2 & "' And a0e07='" & Text1 & "' ", adoTaie, adOpenDynamic, adLockBatchOptimistic
         If .adoacc0e0.RecordCount = 0 Then
            MsgBox MsgText(33) & " " & MsgText(39), , MsgText(5)
            Exit Sub
         End If
      End If
      If .Text1 <> MsgText(601) Then
         .adoacc0e0.Fields("a0e07").Value = .Text1
      Else
         .adoacc0e0.Fields("a0e07").Value = Null
      End If
      If .MaskEdBox3.Text <> MsgText(601) And .MaskEdBox3.Text <> MsgText(29) Then
         .adoacc0e0.Fields("a0e15").Value = Val(FCDate(.MaskEdBox3.Text))
      Else
         .adoacc0e0.Fields("a0e15").Value = 0
      End If
      If .Combo1 <> MsgText(601) Then
         .adoacc0e0.Fields("a0e33").Value = Mid(.Combo1, 1, 1)
      Else
         .adoacc0e0.Fields("a0e33").Value = Null
      End If
      If .Text9 <> MsgText(601) Then
         .adoacc0e0.Fields("a0e12").Value = .Text9
      Else
         .adoacc0e0.Fields("a0e12").Value = Null
      End If
      .adoacc0e0.Fields("a0e14").Value = 0
      If strSaveConfirm = MsgText(3) Then
         .adoacc0e0.Fields("a0e26").Value = Val(strSrvDate(2))
         .adoacc0e0.Fields("a0e27").Value = ServerTime
         .adoacc0e0.Fields("a0e28").Value = strUserNum
      Else
         .adoacc0e0.Fields("a0e29").Value = Val(strSrvDate(2))
         .adoacc0e0.Fields("a0e30").Value = ServerTime
         .adoacc0e0.Fields("a0e31").Value = strUserNum
      End If
      .adoacc0e0.UpdateBatch
      .AdodcRefresh
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .RecordShow
      End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

'從aacc_del搬回
Public Sub Frmacc3150_Delete()
On Error GoTo Checking
   With Frmacc3150
      ''Modify by Amy 2020/07/17 +a0e07 因改為key
      If DeleteCheck("select a0e01 from acc0e0 where a0e01 = '" & .Text6 & "' and a0e02 = '" & .Text2 & "' And a0e07='" & Text1 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "update acc0e0 set a0e15 = 0, a0e33 = null, a0e19 = null, a0e20 = null where a0e01 = '" & .Text6 & "' and a0e02 = '" & .Text2 & "' And a0e07='" & Text1 & "' "
      'end 2020/07/17
      .AdodcRefresh
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveFirst
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

'從aacc_cls搬回
Public Sub Frmacc3150_Clear()
   With Frmacc3150
      .Text6 = ""
      .Text7 = ""
      .Text1 = ""
      .Text2 = ""
      .MaskEdBox3.Mask = ""
      .MaskEdBox3.Text = ""
      .MaskEdBox3.Mask = DFormat
      .Combo1 = ""
      .Text9 = ""
      .MaskEdBox1.Mask = ""
      .MaskEdBox1.Text = ""
      .MaskEdBox1.Mask = DFormat
      .MaskEdBox2.Mask = ""
      .MaskEdBox2.Text = ""
      .MaskEdBox2.Mask = DFormat
      .Text10 = ""
      .Text3 = ""
      .Text5 = ""
      .Text4 = ""
      .Text8 = ""
      .Text11 = ""
      .Text12 = ""
      .Text13 = ""
      .Text2.SetFocus
   End With
End Sub
