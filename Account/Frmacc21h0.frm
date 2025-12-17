VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc21h0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "請款單輸入"
   ClientHeight    =   5400
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8784
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   8784
   Begin VB.TextBox Text10 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   70
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   70
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00FFFFFF&
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
      Left            =   7245
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   22
      Top             =   60
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "使用預留單號"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5445
      TabIndex        =   21
      Top             =   90
      Width           =   1725
   End
   Begin VB.CommandButton Command6 
      Caption         =   "預留單號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   225
      TabIndex        =   20
      Top             =   30
      Width           =   1020
   End
   Begin VB.TextBox Text6 
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
      Left            =   2220
      MaxLength       =   6
      TabIndex        =   1
      Top             =   540
      Width           =   852
   End
   Begin VB.TextBox Text1 
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
      Left            =   1740
      MaxLength       =   3
      TabIndex        =   0
      Top             =   540
      Width           =   492
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   315
      Left            =   1530
      Top             =   4590
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
   Begin VB.CommandButton Command4 
      Caption         =   "點數分配"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7395
      TabIndex        =   7
      Top             =   480
      Width           =   1020
   End
   Begin VB.CommandButton Command2 
      Caption         =   "內容輸入"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6345
      TabIndex        =   6
      Top             =   480
      Width           =   1020
   End
   Begin VB.TextBox Text5 
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
      Left            =   4740
      MaxLength       =   15
      TabIndex        =   4
      Top             =   540
      Width           =   1215
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
      Height          =   315
      Left            =   3060
      MaxLength       =   1
      TabIndex        =   2
      Top             =   540
      Width           =   252
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
      Left            =   3300
      MaxLength       =   2
      TabIndex        =   3
      Top             =   540
      Width           =   372
   End
   Begin VB.CommandButton Command5 
      Height          =   300
      Left            =   5970
      Picture         =   "Frmacc21h0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   540
      Width           =   350
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3300
      Picture         =   "Frmacc21h0.frx":0102
      Style           =   1  '圖片外觀
      TabIndex        =   14
      ToolTipText     =   "確定"
      Top             =   3540
      Width           =   450
   End
   Begin VB.CommandButton Command3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4620
      Picture         =   "Frmacc21h0.frx":0544
      Style           =   1  '圖片外觀
      TabIndex        =   16
      ToolTipText     =   "取消"
      Top             =   3540
      Width           =   450
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6300
      TabIndex        =   18
      Top             =   3600
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc21h0.frx":0986
      Height          =   1425
      Left            =   180
      TabIndex        =   8
      Top             =   2100
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   2519
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
      Caption         =   "未開請款單"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "cp09"
         Caption         =   "總收文號"
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
         DataField       =   "Rdate"
         Caption         =   "收文日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "####/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "cpm03"
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
      BeginProperty Column03 
         DataField       =   "st02"
         Caption         =   "承辦人"
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
         DataField       =   "pa75"
         Caption         =   "代理人"
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
         DataField       =   "cp45"
         Caption         =   "彼所案號"
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
         DataField       =   "Sdate"
         Caption         =   "發文日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "####/##/##"
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
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1307.906
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1247.811
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   1235.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   120
      Top             =   2070
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Frmacc21h0.frx":099B
      Height          =   1320
      Left            =   180
      TabIndex        =   12
      Top             =   4005
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   2328
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
      Caption         =   "已開請款單"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "cp09"
         Caption         =   "總收文號"
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
         DataField       =   "Rdate"
         Caption         =   "收文日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "####/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "cpm03"
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
      BeginProperty Column03 
         DataField       =   "st02"
         Caption         =   "承辦人"
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
         DataField       =   "pa75"
         Caption         =   "代理人"
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
         DataField       =   "cp45"
         Caption         =   "彼所案號"
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
         DataField       =   "Sdate"
         Caption         =   "發文日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "####/##/##"
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
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1307.906
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1247.811
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   1235.906
         EndProperty
      EndProperty
   End
   Begin MSForms.TextBox Text2 
      Height          =   330
      Left            =   1740
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   900
      Width           =   6615
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      Size            =   "11668;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text3 
      Height          =   330
      Left            =   1740
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1260
      Width           =   6615
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      Size            =   "11668;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text4 
      Height          =   330
      Left            =   1740
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1620
      Width           =   6615
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      Size            =   "11668;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      X1              =   2475
      X2              =   2700
      Y1              =   210
      Y2              =   210
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "本所案號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   19
      Top             =   570
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "案件名稱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   300
      TabIndex        =   17
      Top             =   968
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(中)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1260
      TabIndex        =   15
      Top             =   968
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(英)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1260
      TabIndex        =   13
      Top             =   1328
      Width           =   360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(日)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1260
      TabIndex        =   11
      Top             =   1688
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   -60
      Top             =   3810
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1575
      Left            =   135
      Top             =   450
      Width           =   8385
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "請款編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3780
      TabIndex        =   10
      Top             =   570
      Width           =   975
   End
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "地址條"
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
      Left            =   5550
      TabIndex        =   9
      Top             =   3630
      Width           =   735
   End
End
Attribute VB_Name = "Frmacc21h0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/08 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、DataGrid2改字型=新細明體-ExtB、Text2、Text3、Text4
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoacc1k0 As New ADODB.Recordset
Public adorate As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoadodc2 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public frmlink As Form
'Add By Cheng 2003/04/02
Dim strPrint As String '記錄預設印表機名稱
'add by nick 2004/11/17
Public IsPrintAddress As Boolean
'add by nickc 2007/02/08
Dim prnPrint
Dim m_bolActivated As Boolean
Public stF0301 As String, stCP09 As String, stUpdCP09 As String, stNotInCP10 As String, stNowCP10 As String, stNP07 As String 'Add by Amy 2025/11/11

'Added by Morgan 2014/6/6
Private Sub Check1_Click()
   'Removed by Morgan 2023/7/19
   'If Check1.Value = 1 Then
   '   Text11.Enabled = True
   '   Text11.SetFocus
   'Else
   '   Text11.Enabled = False
   'End If
End Sub

'Modify by Amy 2025/10/17 原:Private
Public Sub Command1_Click()
   Dim stNextCP09 As String
   Dim stMsgTxt As String    'add by sonia 2021/3/26
   
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   If IsNull(Adodc1.Recordset.Fields("pa75").Value) Then
      MsgBox MsgText(149), , MsgText(5)
      Exit Sub
   End If
   If Text5 <> MsgText(601) Then
      'Added by Morgan 2024/8/19
      'Modified by Morgan 2025/4/8 +Y45697000 BASF Schweiz AG --Franny
      If strSaveConfirm = MsgText(3) And Val(Right(strSrvDate(1), 2)) > 20 Then
         strExc(1) = PUB_GetA1K28(Adodc1.Recordset("cp01"), Adodc1.Recordset("cp02"), Adodc1.Recordset("cp03"), Adodc1.Recordset("cp04"), Adodc1.Recordset("cp10"))
         If strExc(1) = "Y45814010" Or strExc(1) = "Y33268010" Or strExc(1) = "Y45697000" Then
            MsgBox "BASF案件，每月21日至月底不得請款及寄帳單，故請退回此請款，待下個月一日再請款！", vbExclamation
            Exit Sub
         End If
      End If
      'end 2024/8/19
   
      'Added by Morgan 2019/12/13 核對已准專利是第1個被點選且無發文日時彈訊息
      If Adodc2.Recordset.RecordCount = 0 And Adodc1.Recordset.Fields("cp01") = "FCP" And Adodc1.Recordset.Fields("cp10") = "926" Then
         strExc(0) = "select 1 from caseprogress where cp09='" & Adodc1.Recordset.Fields("cp09") & "' and cp27 is null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            MsgBox "核對已准專利尚未發文，請程序人員上發文後，再繼續請款作業!", vbExclamation
            Exit Sub
         End If
      End If
      'end 2019/12/13
       
       'Modify by Amy 2025/10/17 +if 由結案單開啟
      If stCP09 <> "" And UCase(TypeName(frmlink)) <> "NOTHING" Then
         adoTaie.Execute "update caseprogress set cp60 = '" & Text5 & "' where cp09 = '" & stCP09 & "'"
         adoTaie.Execute "insert into acc1w0 (a1w01, a1w02) values ('" & Text5 & "', '" & stCP09 & "')"
         stCP09 = ""
      Else
'      If IsNull(Adodc1.Recordset.Fields("pa75").Value) = False Then
         adoTaie.Execute "update caseprogress set cp60 = '" & Text5 & "' where cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "'"
         adoTaie.Execute "insert into acc1w0 (a1w01, a1w02) values ('" & Text5 & "', '" & Adodc1.Recordset.Fields("cp09").Value & "')"
'      Else
'         MsgBox MsgText(149), , MsgText(5)
'         Exit Sub
'      End If
      End If
      'end 2025/10/17

      'add by sonia 2021/3/25 選非補收款(專911,商705)或延期(專404,商303)時若有補收款或延期未請款也要一併請FCP-062918
      'modify by sonia 2021/4/22 再加是否向客戶收款條件and cp20 is null
      strExc(0) = "select cpm03,cp09,cp10,sk02 from systemkind,caseprogress,casepropertymap where cp43='" & Adodc1.Recordset.Fields("cp09") & "' and cp01=sk01(+) and nvl(cp16,0)+nvl(cp84,0)>0 and cp20 is null and cp60 is null and cp01=cpm01(+) and cp10=cpm02(+)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If ((RsTemp.Fields("sk02") = "1" Or RsTemp.Fields("sk02") = "5") And (RsTemp.Fields("cp10") = "911" Or RsTemp.Fields("cp10") = "404")) _
         Or ((RsTemp.Fields("sk02") = "2" Or RsTemp.Fields("sk02") = "6") And (RsTemp.Fields("cp10") = "705" Or RsTemp.Fields("cp10") = "303")) Then
            MsgBox "此收文號尚有 " & RsTemp.Fields("cpm03") & " 未請款，將一併列入請款!", vbExclamation
            adoTaie.Execute "update caseprogress set cp60 = '" & Text5 & "' where cp09 = '" & RsTemp.Fields("cp09") & "'"
            adoTaie.Execute "insert into acc1w0 (a1w01, a1w02) values ('" & Text5 & "', '" & RsTemp.Fields("cp09") & "')"
         End If
      End If
      '選補收款(專911,商705)或延期(專404,商303)其相關總收文號未請款也要一併請款，
      'modify by sonia 2021/4/22 再加已發文條件and cp158>0
      'modify by sonia 2021/4/28 再加是否向客戶收款條件and cp20 is null (FCP-063441延期-申復之CP43為審查意見書)
      strExc(0) = "select cpm03,b.cp09,sk02,a.cp10 cp10 from systemkind,caseprogress a,caseprogress b,casepropertymap where a.cp09='" & Adodc1.Recordset.Fields("cp09") & "' and a.cp43=b.cp09(+) and b.cp01=sk01(+) and nvl(b.cp16,0)+nvl(b.cp84,0)>0 and b.cp158>0 and b.cp20 is null and b.cp60 is null and b.cp01=cpm01(+) and b.cp10=cpm02(+)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If ((RsTemp.Fields("sk02") = "1" Or RsTemp.Fields("sk02") = "5") And (RsTemp.Fields("cp10") = "911" Or RsTemp.Fields("cp10") = "404")) _
         Or ((RsTemp.Fields("sk02") = "2" Or RsTemp.Fields("sk02") = "6") And (RsTemp.Fields("cp10") = "705" Or RsTemp.Fields("cp10") = "303")) Then
            MsgBox "相關總收文號 " & RsTemp.Fields("cpm03") & " 尚未請款，將一併列入請款!", vbExclamation
            adoTaie.Execute "update caseprogress set cp60 = '" & Text5 & "' where cp09 = '" & RsTemp.Fields("cp09") & "'"
            adoTaie.Execute "insert into acc1w0 (a1w01, a1w02) values ('" & Text5 & "', '" & RsTemp.Fields("cp09") & "')"
         End If
      End If
      '選FCP面詢408若有請求面詢407或請求面詢之補收款911未請款時也要一併請款
      If Adodc1.Recordset.Fields("cp01") = "FCP" And Adodc1.Recordset.Fields("cp10") = "408" Then
         strExc(0) = "select b.cp09 no407,b.cp60 db407,c.cp09 no911,c.cp60 db911 from caseprogress a,caseprogress b,caseprogress c where a.cp09='" & Adodc1.Recordset.Fields("cp09") & "' and a.cp01=b.cp01(+) and a.cp02=b.cp02(+) and a.cp03=b.cp03(+) and a.cp04=b.cp04(+) and '407'=b.cp10(+) and b.cp09=c.cp43(+) and '911'=c.cp10(+)"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            stMsgTxt = ""
            If IsNull(RsTemp.Fields("db407")) = True Then
               If Not IsNull(RsTemp.Fields("no407")) Then 'Added by Morgan 2025/2/25 不一定會有407 Ex:FCP-039857--敏莉
                  stMsgTxt = "請求面詢"
                  adoTaie.Execute "update caseprogress set cp60 = '" & Text5 & "' where cp09 = '" & RsTemp.Fields("no407") & "'"
                  adoTaie.Execute "insert into acc1w0 (a1w01, a1w02) values ('" & Text5 & "', '" & RsTemp.Fields("no407") & "')"
               End If
            End If
            If IsNull(RsTemp.Fields("db911")) = True And IsNull(RsTemp.Fields("no911")) = False Then
               If stMsgTxt = "請求面詢" Then
                  stMsgTxt = "請求面詢及其補收款"
               Else
                  stMsgTxt = "請求面詢之補收款"
               End If
               adoTaie.Execute "update caseprogress set cp60 = '" & Text5 & "' where cp09 = '" & RsTemp.Fields("no911") & "'"
               adoTaie.Execute "insert into acc1w0 (a1w01, a1w02) values ('" & Text5 & "', '" & RsTemp.Fields("no911") & "')"
            End If
            If stMsgTxt <> "" Then
               MsgBox stMsgTxt & " 尚未請款，將一併列入請款!", vbExclamation
            End If
         End If
      End If
      'end 2021/3/25
      
      
      'Added by Morgan 2019/10/7
      '游標能一直往下跑，不要再跳回第一道--敏莉
      Adodc1.Recordset.MoveNext
      If Not Adodc1.Recordset.EOF Then
         stNextCP09 = Adodc1.Recordset.Fields("cp09")
      End If
      'end 2019/10/7
      
   End If
   AdodcRefresh
   
   'Added by Morgan 2019/10/7
   '游標能一直往下跑，不要再跳回第一道--敏莉
   'modify by sonia 2021/4/6
   'If stNextCP09 <> "" Then
   If stNextCP09 <> "" And adoadodc1.RecordCount > 0 Then
      Adodc1.Recordset.Find "cp09='" & stNextCP09 & "'"
      If Adodc1.Recordset.EOF Then
         Adodc1.Recordset.MoveFirst
      End If
   End If
   'end 2019/10/7
   
End Sub

'Modified by Morgan 2017/1/7 改共用(原來寫在Command2_Click中)
Public Function ModifyCheck() As Boolean
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   
    'Add By Cheng 2003/11/14
    StrSQLa = "Select * From ACC1K0 Where A1K01='" & Me.Text5.Text & "' "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        If "" & rsA("A1K12").Value <> "" And "" & rsA("A1K12").Value <> "0" Then
            MsgBox "此筆請款單已作廢!!!", vbExclamation + vbOKOnly
            Me.Text5.SetFocus
            Text5_GotFocus
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            Exit Function
        End If
        If "" & rsA("A1K29").Value = "Y" Then
            MsgBox "此筆請款單已結清, 若要列印請至請款單列印!!!", vbExclamation + vbOKOnly
            Me.Text5.SetFocus
            Text5_GotFocus
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            Exit Function
        End If
        
        'Added by Morgan 2013/1/9
        If Val("" & rsA("A1K06").Value) > 0 Then
            MsgBox "此筆請款單已有折讓不可修改!!!", vbExclamation + vbOKOnly
            Me.Text5.SetFocus
            Text5_GotFocus
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            Exit Function
        End If
        'end 2013/1/9
        
    Else
        MsgBox "查無資料!!!", vbExclamation + vbOKOnly
        Me.Text5.SetFocus
        Text5_GotFocus
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        Exit Function
    End If
    
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    'End
    
    ModifyCheck = True
End Function

'Modify by Amy 2025/10/17 原:Private
Public Sub Command2_Click()
Dim Cancel As Boolean 'Add By Sindy 2009/07/15
Dim stMsg As String, arrData 'Add by Amy 2025/10/17
   
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Exit Sub
   End If
   If Text5 = MsgText(601) Then
      Exit Sub
   End If
   
   'Add By Sindy 2009/07/15
   Cancel = False
   Call Text1_Validate(Cancel)
   If Cancel = True Then Exit Sub
   '2009/07/15 End
   
   If ModifyCheck = False Then Exit Sub
   
   '請款單號
   strItemNo = Text5
   '請款匯率
   If adoacc1k0.RecordCount <> 0 Then
      If IsNull(adoacc1k0.Fields("a1k10").Value) Then
         dblRate = 0
      Else
         dblRate = adoacc1k0.Fields("a1k10").Value
      End If
   End If
   '代理人
   If Adodc2.Recordset.RecordCount <> 0 Then
      If IsNull(Adodc2.Recordset.Fields("pa75").Value) Then
         strCon1 = ""
      Else
         strCon1 = Adodc2.Recordset.Fields("pa75").Value
      End If
   End If
   'add by nickc 2005/06/28 要先查到資料
      'modified by Lydia 2014/12/16 從讀CP09下方移到上方
   If adoadodc2.EOF And adoadodc2.BOF Then
      MsgBox "請先按望遠鏡查詢資料！", , "警告！"
      Exit Sub
   End If
   
   'Added by Morgan 2022/3/1 整批列印後又進明細畫面(會自動重算)，導致催款與請款金額不同 Ex:X11102139
   '整批列印的請款單增加提醒
   If adoacc1k0("a1k32") = "C" Then
      If MsgBox("本請款單為「整批列印」的請款單，若只查看而沒有要修改資料，請執行「國外案件帳目查詢」！" & vbCrLf & vbCrLf & "是否確定要修改資料？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
         Exit Sub
      End If
   End If
   'end 2022/3/1
   
   '收文號
   strCon9 = Adodc2.Recordset("CP09").Value 'Added by Morgan 2014/8/6 收文號
   '表單名稱
   strFormLink = Name
   tool3_enabled

   Set Frmacc21h1.m_FromForm = Me 'Added by Morgan 2014/8/15
   'Add by Amy 2025/11/11 傳入更新資料(結案單用)
   Frmacc21h1.stF0301 = stF0301
   Frmacc21h1.stNowCP10 = stNowCP10 '不續辦or閉卷
   Frmacc21h1.stUpdCP09 = stUpdCP09 '寫於Frmacc21h1.Show 前,才能抓到總收文號對應之項次
   Frmacc21h1.stNotInCP10 = stNotInCP10 '進度沒有之案件性質
   Frmacc21h1.stNP07 = stNP07
   'end 2025/11/11
   Frmacc21h1.Show
   
   'add by nick 2004/11/17
   Frmacc21h0.IsPrintAddress = IsPrintAddress
   Me.Hide
End Sub

Private Sub Command3_Click()
   If Adodc2.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   adoTaie.Execute "update caseprogress set cp60 = null where cp09 = '" & Adodc2.Recordset.Fields("cp09").Value & "'"
   adoTaie.Execute "delete from acc1w0 where a1w01 = '" & Text5 & "' and a1w02 = '" & Adodc2.Recordset.Fields("cp09").Value & "'"
   AdodcRefresh
End Sub

Private Sub Command4_Click()
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Exit Sub
   End If
   If Text5 = MsgText(601) Then
      Exit Sub
   End If
   strExc(0) = "Select * From ACC1K0 Where A1K01='" & Me.Text5.Text & "' "
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(RsTemp("A1K12")) Then
         MsgBox "此筆請款單已作廢!!!", vbExclamation + vbOKOnly
         Me.Text5.SetFocus
         Text5_GotFocus
         Exit Sub
      End If
      If RsTemp("A1K29").Value = "Y" Then
         strExc(1) = ""
         If Pub_StrUserSt03 = "M51" Then
            strExc(1) = MsgBox("此筆請款單已結清, 是否仍然要調整分配點數!!!", vbExclamation + vbYesNo + vbDefaultButton2)
         Else
            strExc(1) = MsgBox("此筆請款單已結清, 不可調整分配點數!!!", vbExclamation + vbOKOnly)
         End If
         If strExc(1) <> vbYes Then
            Me.Text5.SetFocus
            Text5_GotFocus
            Exit Sub
         End If
      End If
   Else
      MsgBox "查無資料!!!", vbExclamation + vbOKOnly
      Me.Text5.SetFocus
      Text5_GotFocus
      Exit Sub
   End If
   If adoadodc2.EOF And adoadodc2.BOF Then
      MsgBox "請先按望遠鏡查詢資料！", , "警告！"
      Exit Sub
   End If
   strItemNo = Text5
   Frmacc21h3.Show vbModal
   strItemNo = ""
   strFormName = Me.Name
   tool1_enabled
End Sub

Private Sub Command5_Click()
   'If adoacc1k0.RecordCount = 0 Or Text5 = MsgText(601) Then
   '   Exit Sub
   'End If
   'adoacc1k0.Find "a1k01 = '" & Text5 & "'", 0, adSearchForward, 1
   'If adoacc1k0.EOF = False Then
   '   FormShow
   '   AdodcRefresh
   '   RecordShow
   'Else
   '   MsgBox MsgText(33), , MsgText(5)
   '   adoacc1k0.MoveFirst
   'End If
   Acc1k0Refresh
   If adoacc1k0.RecordCount <> 0 Then
      FormShow
      'AdodcRefresh 'Removed by Morgan 2021/4/15 FormShow裡面的CaseQuery有執行了，不必重複
      RecordShow
   'Added by Morgan 2015/8/28
   Else
      Text5.Tag = ""
      Text1.Text = ""
      Text6.Text = ""
      Text7.Text = ""
      Text8.Text = ""
      Text2.Text = ""
      Text3.Text = ""
      Text4.Text = ""
      AdodcRefresh
      RecordShow
   'end 2015/8/28
   End If
End Sub
'Added by Morgan 2014/6/6
Private Sub Command6_Click()
   
   'Added by Morgan 2023/7/19
   If GetRsvDN() = True Then
      MsgBox "您尚有預留單號未使用，已自動載入！", vbInformation
   Else
   'end 2023/7/19
   
      strExc(0) = InputBox("請輸入要預留單號的數量：", Me.Caption & "-" & Command6.Caption)
      If Val(strExc(0)) > 0 Then
         If Val(strExc(0)) > 10 Then
            If MsgBox("系統將預留 " & Val(strExc(0)) & " 個單號，是否確定要繼續？", vbYesNo) = vbNo Then
               Exit Sub
            End If
         End If
         'Modified by Morgan 2023/6/28
         'ReserveNo Val(strExc(0))
         If PUB_RsvDN(CInt(strExc(0)), strExc(1), strExc(2)) = True Then
            Text9 = strExc(1)
            Text10 = strExc(2)
            Check1.Enabled = True
            Check1.Value = 1
            Text11 = Text9
         End If
         'end 2023/6/28
      End If
   End If
End Sub

'Added by Morgan 2023/7/19
Private Function GetRsvDN() As Boolean
   If PUB_GetRsvDN(strExc(1), strExc(2)) = True Then
      Text9 = strExc(1)
      Text10 = strExc(2)
      Check1.Enabled = True
      Check1.Value = 1
      Text11 = Text9
      GetRsvDN = True
   Else
      Text9 = ""
      Text10 = ""
      Text11 = ""
      Check1.Enabled = False
      Check1.Value = 0
   End If
End Function

'Removed by Morgan 2023/6/28 改寫公用 PUB_RsvDnNo
'Added by Morgan 2014/6/6
'Private Sub ReserveNo(pQty As Integer)
'
'   adoTaie.BeginTrans
'
'On Error GoTo ErrHnd
'
'   adoTaie.Execute "update acc1r0 set a1r04 = a1r04 where a1r01 = 'X'"
'   strExc(1) = AccAutoNo("X", 5)
'   If AccSaveAutoNo("X", Right(strExc(1), 5)) = "Y" Then
'      adoTaie.Execute "update acc1r0 set a1r04 = a1r04+" & (pQty - 1) & " where a1r01 = 'X'"
'   End If
'   adoTaie.CommitTrans
'   Text9 = strExc(1)
'   If pQty > 1 Then
'      Text10 = Left(Text9, 1) & (Val(Mid(Text9, 2)) + pQty - 1)
'   Else
'      Text10 = Text9
'   End If
'   Check1.Enabled = True
'   Check1.Value = 1
'   Text11 = Text9
'   Exit Sub
'
'ErrHnd:
'   adoTaie.RollbackTrans
'   MsgBox Err.Description, vbCritical
'
'End Sub

'Added by Morgan 2014/6/6
Public Function AddCheck() As Boolean
   If Check1.Value = 1 Then
      'Modified by Morgan 2023/7/19 舊程式已刪除,只保留新的
      If GetRsvDN() = False Then
         MsgBox "預留單號讀取失敗！", vbCritical
         Exit Function
      ElseIf PUB_UpdRsvDN(Text11.Text) = False Then
         MsgBox "預留單號使用失敗！", vbCritical
         Exit Function
      End If
   End If
   AddCheck = True
End Function

Private Sub Form_Activate()
   If Not m_bolActivated Then
      'Modified by Lydia 2021/12/08 調高
      'PUB_InitForm Me
      PUB_InitForm Me, , 5850
      m_bolActivated = True
   End If
   
'edit by nickc 2007/02/08
'   '93.3.16 ADD BY SONIA
'   If IsObject(mdiMain) Then
'      mdiMain.toolshow
'   End If
'   '93.3.16 END
   Dim formCnt As Integer
   For formCnt = 0 To Forms.Count - 1
       If UCase(Forms(formCnt).Name) = "MDIMAIN" Then
             Forms(formCnt).ToolShow
             Exit For
       End If
   Next
   strFormName = Name
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   If adoacc1k0.RecordCount <> 0 Then
      adoacc1k0.MoveFirst
   End If
   'adoacc1k0.Find "a1k01 = '" & strItemNo & "'", 0, adSearchForward, 1
   'If adoacc1k0.EOF = False Then
   '   FormShow
   '   AdodcRefresh
   '   RecordShow
   'End If
   Text5 = strItemNo
   Acc1k0Refresh
   If adoacc1k0.RecordCount <> 0 Then
      FormShow
      AdodcRefresh
      RecordShow
   End If
   strItemNo = MsgText(601)
   
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
   
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim ii As Integer

   OpenTable
   If adoacc1k0.RecordCount <> 0 Then
      adoacc1k0.MoveLast
      adoacc1k0.MoveFirst
      RecordShow
   End If
   
   PUB_SetPrinter Me.Name, Combo2, strPrint, False 'Modified by Morgan 2017/11/8 設定印表機改呼叫公用函數,原程式移除
       
   'Add by Morgan 2004/9/10 操作人員為 'F1x' 部門的結束時印地址條
   If Left(GetStaffDepartment(strUserNum), 2) = "F1" Then pub_blnARPrintAddress = True
   'add by nick 2004/11/17
   IsPrintAddress = True
   
   'Added by Morgan 2014/6/9
   Check1.Enabled = False
   Text11.Enabled = False
   'end 2014/6/9
   
   '2015/10/21 add by sonia 加入總經理權限(等級01),可使用所有程式,但維護程式只有查詢功能,不可新增刪除修改)
   If PUB_GetST05(strUserNum) = "01" Then
      Command2.Enabled = False
      Command4.Enabled = False
      Command6.Enabled = False
   End If
   '2015/10/21 end
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Add By Cheng 2003/04/02
    If pub_blnARPrintAddress = True Then
        '列印地址條
        PUB_PrintAddressList strUserNum, Me.Combo2.Text
        '刪除地址條列表資料
        PUB_DeleteAddressList strUserNum
        '初始化序號
        pub_AddressListSN = 0
    End If
    pub_blnARPrintAddress = False
    'Add By Cheng 2003/02/05
    '若印表機變動, 則更新列印設定
    If Me.Combo2.Text <> Me.Combo2.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
    End If
   For Each prnPrint In Printers
      If prnPrint.DeviceName = strPrint Then
         Set Printer = prnPrint
      End If
   Next
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   If TypeName(frmlink) <> "Nothing" Then
      frmlink.Show
      'Modify by Amy 2025/08/19 +if 外商結案單,於 解除期限 or 閉卷完成後回待處理區
      If UCase(frmlink.Name) = "FRM210149_1" Then
         Call frmlink.cmdExit_Click
      Else
         frmlink.Clear
      End If
   End If
   stF0301 = "": stCP09 = "": stUpdCP09 = "": stNotInCP10 = "": stNowCP10 = "": stNP07 = "" 'Add by Amy 2025/11/11
   Set frmlink = Nothing
   Set Frmacc21h0 = Nothing
End Sub

Private Sub Text1_GotFocus()
   CloseIme
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
   
   'Add By Sindy 2025/2/20
   If adoacc1k0.State = adStateOpen Then
      adoacc1k0.Close
   End If
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   If adoadodc2.State = adStateOpen Then
      adoadodc2.Close
   End If
   '2025/2/20 END
   adoacc1k0.CursorLocation = adUseClient
   adoacc1k0.MaxRecords = intMax
   adoacc1k0.Open "select * from acc1k0 where a1k01 >= '" & Text5 & "' and (a1k12 is null or a1k12 <> '') order by a1k01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      adoadodc1.Open "select * from caseprogress where cp01 = '" & Text1 & "' and cp02 = '" & Text6 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text8 & "' order by cp05 desc, cp09 desc", adoTaie, adOpenStatic, adLockReadOnly
   Else
      adoadodc1.Open "select * from caseprogress where cp01 = 'Z'", adoTaie, adOpenStatic, adLockReadOnly
   End If
   adoadodc2.CursorLocation = adUseClient
   adoadodc2.Open "select * from caseprogress where cp01 = '" & Text1 & "' and cp02 = '" & Text6 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text8 & "' and cp60 = '" & Text5 & "' and (cp27 is not null and cp27 <> 0) order by cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
   Set Adodc2.Recordset = adoadodc2
   
   Exit Sub 'Add By Sindy 2025/2/20
   
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
Dim StrSQLa As String, i As Integer

On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   adoadodc2.Close
   adoadodc2.CursorLocation = adUseClient
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      'Modify by Amy 2025/10/03 搬至共用,調整讓結案單也可使用,因函數尚未確認完,先保留
      StrSQLa = "select cp09, nvl(cp05 - 19110000, 0) Rdate, decode(pa09, '000', cpm03, cpm04)||Decode(CP10,'1001',A.A1,'1002',A.A1,'1008',A.A1,'')||GetRelateCasePropertyName(cp09,'1') as cpm03, st02, nvl(pa75,pa26) as pa75, cp45, nvl(cp27 - 19110000, 0) Sdate, cp60, cp01, cp02, cp03, cp04, cp17, cp16, cp05,cp10 from caseprogress, patent, casepropertymap, staff, nation, (Select ' - '||decode(pa09, '000', cpm03, cpm04) As A1, CP09 As A2 From Patent, Caseprogress, CasePropertyMap Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP01=CPM01 And CP10=CPM02 And PA01='" & Me.Text1.Text & "' And PA02='" & Me.Text6.Text & "' And PA03='" & Me.Text7.Text & "' And PA04='" & Me.Text8.Text & "' ) A " & _
                                "where cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and cp01 = cpm01 and cp10 = cpm02 and cp14 = st01 (+) and pa09 = na01 (+) And CP43=A.A2(+) and cp01 = '" & Text1 & "' and cp02 = '" & Text6 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text8 & "' and (cp60 is null or cp60 = '') and (cp27 is not null and cp27 <> 0) and cp20 is null  union " & _
                "select cp09, nvl(cp05 - 19110000, 0) Rdate, decode(pa09, '000', cpm03, cpm04)||Decode(CP10,'1001',A.A1,'1002',A.A1,'1008',A.A1,'')||GetRelateCasePropertyName(cp09,'1') as cpm03, st02, nvl(pa75,pa26) as pa75, cp45, nvl(cp27 - 19110000, 0) Sdate, cp60, cp01, cp02, cp03, cp04, cp17, cp16, cp05,cp10 from caseprogress, patent, casepropertymap, staff, nation, (Select ' - '||decode(pa09, '000', cpm03, cpm04) As A1, CP09 As A2 From Patent, Caseprogress, CasePropertyMap Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP01=CPM01 And CP10=CPM02 And PA01='" & Me.Text1.Text & "' And PA02='" & Me.Text6.Text & "' And PA03='" & Me.Text7.Text & "' And PA04='" & Me.Text8.Text & "' ) A " & _
                                "where cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and cp01 = cpm01 and cp10 = cpm02 and cp14 = st01 (+) and pa09 = na01 (+) And CP43=A.A2(+) and cp01 = '" & Text1 & "' and cp02 = '" & Text6 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text8 & "' and (cp60 is null or cp60 = '') and ((((cp01='P' or cp01='CFP') and substr(nvl(cp12,'S'),1,1)<>'F') or cp10 = '201' or cp10 = '209' or cp10 = '210' or cp10 = '223' or cp10 = '926') and cp27 is null AND CP57 IS NULL) and cp20 is null union " & _
                "select cp09, nvl(cp05 - 19110000, 0) Rdate, decode(tm10, '000', cpm03, cpm04)||Decode(CP10,'1001',A.A1,'1002',A.A1,'1003',A.A1,'1004',A.A1,'1008',A.A1,'')||GetRelateCasePropertyName(cp09,'1') as cpm03, st02, nvl(tm44,tm23) as pa75, cp45, nvl(cp27 - 19110000, 0) Sdate, cp60, cp01, cp02, cp03, cp04, cp17, cp16, cp05,cp10 from caseprogress, trademark, casepropertymap, staff, nation, (Select ' - '||decode(tm10, '000', cpm03, cpm04) As A1, CP09 As A2 From Trademark, Caseprogress, CasePropertyMap Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP01=CPM01 And CP10=CPM02 And TM01='" & Me.Text1.Text & "' And TM02='" & Me.Text6.Text & "' And TM03='" & Me.Text7.Text & "' And TM04='" & Me.Text8.Text & "' ) A " & _
                                "where cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and cp01 = cpm01 and cp10 = cpm02 and cp14 = st01 (+) and tm10 = na01 (+) And CP43=A.A2(+) and cp01 = '" & Text1 & "' and cp02 = '" & Text6 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text8 & "' and (cp60 is null or cp60 = '') and (cp27 is not null and cp27 <> 0) and cp20 is null  union " & _
                "select cp09, nvl(cp05 - 19110000, 0) Rdate, decode(tm10, '000', cpm03, cpm04)||Decode(CP10,'1001',A.A1,'1002',A.A1,'1003',A.A1,'1004',A.A1,'1008',A.A1,'')||GetRelateCasePropertyName(cp09,'1') as cpm03, st02, nvl(tm44,tm23) as pa75, cp45, nvl(cp27 - 19110000, 0) Sdate, cp60, cp01, cp02, cp03, cp04, cp17, cp16, cp05,cp10 from caseprogress, trademark, casepropertymap, staff, nation, (Select ' - '||decode(tm10, '000', cpm03, cpm04) As A1, CP09 As A2 From Trademark, Caseprogress, CasePropertyMap Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP01=CPM01 And CP10=CPM02 And TM01='" & Me.Text1.Text & "' And TM02='" & Me.Text6.Text & "' And TM03='" & Me.Text7.Text & "' And TM04='" & Me.Text8.Text & "' ) A " & _
                                "where cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and cp01 = cpm01 and cp10 = cpm02 and cp14 = st01 (+) and tm10 = na01 (+) And CP43=A.A2(+) and cp01 = '" & Text1 & "' and cp02 = '" & Text6 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text8 & "' and (cp60 is null or cp60 = '') and (cp01='FCT' and cp10 = '601') and cp27 is null AND CP57 IS NULL and cp20 is null  union " & _
                "select cp09, nvl(cp05 - 19110000, 0) Rdate, decode(tm10, '000', cpm03, cpm04)||Decode(CP10,'1001',A.A1,'1002',A.A1,'1003',A.A1,'1004',A.A1,'1008',A.A1,'')||GetRelateCasePropertyName(cp09,'1') as cpm03, st02, nvl(tm44,tm23) as pa75, cp45, nvl(cp27 - 19110000, 0) Sdate, cp60, cp01, cp02, cp03, cp04, cp17, cp16, cp05,cp10 from caseprogress, trademark, casepropertymap, staff, nation, (Select ' - '||decode(tm10, '000', cpm03, cpm04) As A1, CP09 As A2 From Trademark, Caseprogress, CasePropertyMap Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP01=CPM01 And CP10=CPM02 And TM01='" & Me.Text1.Text & "' And TM02='" & Me.Text6.Text & "' And TM03='" & Me.Text7.Text & "' And TM04='" & Me.Text8.Text & "' ) A " & _
                                "where cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and cp01 = cpm01 and cp10 = cpm02 and cp14 = st01 (+) and tm10 = na01 (+) And CP43=A.A2(+) and cp01 = '" & Text1 & "' and cp02 = '" & Text6 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text8 & "' and (cp60 is null or cp60 = '') and (substr(cp12,1,2)='F1' and cp09<'B' and nvl(cp16,0)>0) AND CP57 IS NULL and cp20 is null  union " & _
                "select cp09, nvl(cp05 - 19110000, 0) Rdate, decode(lc15, '000', cpm03, cpm04)||GetRelateCasePropertyName(cp09,'1') as cpm03, st02, nvl(lc22,lc11) as pa75, cp45, nvl(cp27 - 19110000, 0) Sdate, cp60, cp01, cp02, cp03, cp04, cp17, cp16, cp05,cp10 from caseprogress, lawcase, casepropertymap, staff, nation where cp01 = lc01 and cp02 = lc02 and cp03 = lc03 and cp04 = lc04 and cp01 = cpm01 and cp10 = cpm02 and cp14 = st01 (+) and lc15 = na01 (+) and cp01 = '" & Text1 & "' and cp02 = '" & Text6 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text8 & "' and (cp60 is null or cp60 = '') and (cp27 is not null and cp27 <> 0) and cp20 is null  union " & _
                "select cp09, nvl(cp05 - 19110000, 0) Rdate, decode(sp09, '000', cpm03, cpm04)||GetRelateCasePropertyName(cp09,'1') as cpm03, st02, nvl(sp26,sp08) as pa75, cp45, nvl(cp27 - 19110000, 0) Sdate, cp60, cp01, cp02, cp03, cp04, cp17, cp16, cp05,cp10 from caseprogress, servicepractice, casepropertymap, staff, nation where cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and cp01 = cpm01 and cp10 = cpm02 and cp14 = st01 (+) and sp09 = na01 (+) and cp01 = '" & Text1 & "' and cp02 = '" & Text6 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text8 & "' and (cp60 is null or cp60 = '') and (cp27 is not null and cp27 <> 0) and cp20 is null  union " & _
                "select cp09, nvl(cp05 - 19110000, 0) Rdate, decode(sp09, '000', cpm03, cpm04)||GetRelateCasePropertyName(cp09,'1') as cpm03, st02, nvl(sp26,sp08) as pa75, cp45, nvl(cp27 - 19110000, 0) Sdate, cp60, cp01, cp02, cp03, cp04, cp17, cp16, cp05,cp10 from caseprogress, servicepractice, casepropertymap, staff, nation where cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and cp01 = cpm01 and cp10 = cpm02 and cp14 = st01 (+) and sp09 = na01 (+) and cp01 = '" & Text1 & "' and cp02 = '" & Text6 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text8 & "' and (cp60 is null or cp60 = '') and (substr(cp12,1,2)='F1' and cp09<'B' and nvl(cp16,0)>0) and cp20 is null order by cp05 desc, cp09 desc"
      'StrSQLa = GetAcc21H0Sql(0, Me.Name, "", Me.Text1 & "-" & Me.Text6 & "-" & Me.Text7 & "-" & Me.Text8)
      'end 2025/10/03
      
      'Modify By Sindy 2011/5/23
      'adoadodc1.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
      adoadodc1.Open StrSQLa, adoTaie, adOpenStatic, adLockBatchOptimistic
   Else
      'Modify By Sindy 2011/5/23
      'adoadodc1.Open "select cp09, nvl(cp05 - 19110000, 0) Rdate, cpm03, st02, '' as pa75, cp45, nvl(cp27 - 19110000, 0) Sdate, cp60, cp01, cp02, cp03, cp04, cp17, cp16 from caseprogress, casepropertymap, staff where cp01 = cpm01 and cp10 = cpm02 and cp14 = st01 (+) and cp01 = 'Z'", adoTaie, adOpenStatic, adLockReadOnly
      'Modified by Morgan 2021/4/16 +GetRelateCasePropertyName
      adoadodc1.Open "select cp09, nvl(cp05 - 19110000, 0) Rdate, cpm03||GetRelateCasePropertyName(cp09,'1') as cpm03, st02, '' as pa75, cp45, nvl(cp27 - 19110000, 0) Sdate, cp60, cp01, cp02, cp03, cp04, cp17, cp16 from caseprogress, casepropertymap, staff where cp01 = cpm01 and cp10 = cpm02 and cp14 = st01 (+) and cp01 = 'Z'", adoTaie, adOpenStatic, adLockBatchOptimistic
   End If
      'Modified by Lydia 2015/10/05  + 1008
      'Modified by Morgan 2021/4/16 +GetRelateCasePropertyName
   StrSQLa = "select cp09, nvl(cp05 - 19110000, 0) Rdate, decode(pa09, '000', cpm03, cpm04)||Decode(CP10,'1001',A.A1,'1002',A.A1,'1008',A.A1,'')||GetRelateCasePropertyName(cp09,'1') as cpm03, st02, pa75, cp45, nvl(cp27 - 19110000, 0) Sdate, cp60, cp01, cp02, cp03, cp04, cp17, cp16 from caseprogress, patent, casepropertymap, staff, nation, (Select ' - '||decode(pa09, '000', cpm03, cpm04) As A1, CP09 As A2 From Patent, Caseprogress, CasePropertyMap Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP01=CPM01 And CP10=CPM02 And PA01='" & Me.Text1.Text & "' And PA02='" & Me.Text6.Text & "' And PA03='" & Me.Text7.Text & "' And PA04='" & Me.Text8.Text & "' ) A where cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and cp01 = cpm01 and cp10 = cpm02 and cp14 = st01 (+) and pa09 = na01 (+) And CP43=A.A2(+) and cp01 = '" & Text1 & "' and cp02 = '" & Text6 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text8 & "' and cp60 = '" & Text5 & "' union " & _
                  "select cp09, nvl(cp05 - 19110000, 0) Rdate, decode(tm10, '000', cpm03, cpm04)||Decode(CP10,'1001',A.A1,'1002',A.A1,'1003',A.A1,'1004',A.A1,'1008',A.A1,'')||GetRelateCasePropertyName(cp09,'1') as cpm03, st02, tm44 as pa75, cp45, nvl(cp27 - 19110000, 0) Sdate, cp60, cp01, cp02, cp03, cp04, cp17, cp16 from caseprogress, trademark, casepropertymap, staff, nation, (Select ' - '||decode(tm10, '000', cpm03, cpm04) As A1, CP09 As A2 From Trademark, Caseprogress, CasePropertyMap Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP01=CPM01 And CP10=CPM02 And TM01='" & Me.Text1.Text & "' And TM02='" & Me.Text6.Text & "' And TM03='" & Me.Text7.Text & "' And TM04='" & Me.Text8.Text & "' ) A where cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and cp01 = cpm01 and cp10 = cpm02 and cp14 = st01 (+) and tm10 = na01 (+) And CP43=A.A2(+) and cp01 = '" & Text1 & "' and cp02 = '" & Text6 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text8 & "' and cp60 = '" & Text5 & "' union " & _
                  "select cp09, nvl(cp05 - 19110000, 0) Rdate, decode(lc15, '000', cpm03, cpm04)||GetRelateCasePropertyName(cp09,'1') as cpm03, st02, lc22 as pa75, cp45, nvl(cp27 - 19110000, 0) Sdate, cp60, cp01, cp02, cp03, cp04, cp17, cp16 from caseprogress, lawcase, casepropertymap, staff, nation where cp01 = lc01 and cp02 = lc02 and cp03 = lc03 and cp04 = lc04 and cp01 = cpm01 and cp10 = cpm02 and cp14 = st01 (+) and lc15 = na01 (+) and cp01 = '" & Text1 & "' and cp02 = '" & Text6 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text8 & "' and cp60 = '" & Text5 & "' union " & _
                  "select cp09, nvl(cp05 - 19110000, 0) Rdate, decode(sp09, '000', cpm03, cpm04)||GetRelateCasePropertyName(cp09,'1') as cpm03, st02, sp26 as pa75, cp45, nvl(cp27 - 19110000, 0) Sdate, cp60, cp01, cp02, cp03, cp04, cp17, cp16 from caseprogress, servicepractice, casepropertymap, staff, nation where  cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and cp01 = cpm01 and cp10 = cpm02 and cp14 = st01 (+) and sp09 = na01 (+) and cp01 = '" & Text1 & "' and cp02 = '" & Text6 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text8 & "' and cp60 = '" & Text5 & "' order by cp09 asc"

   
   
   'Modify By Sindy 2011/5/23
   'adoadodc2.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   adoadodc2.Open StrSQLa, adoTaie, adOpenStatic, adLockBatchOptimistic
   Adodc1.Recordset.ReQuery
   Adodc2.Recordset.ReQuery
   
   'Add By Sindy 2011/5/23
   If adoadodc1.RecordCount > 0 Then adoadodc1.MoveFirst
'Removed by Morgan 2021/4/16 OraOLEDB會有錯誤，改語法內直接用db函數抓'
'   For i = 1 To adoadodc1.RecordCount
'      adoadodc1.Fields(2) = adoadodc1.Fields(2) & PUB_GetRelateCasePropertyName(adoadodc1.Fields(0), "1")
'      adoadodc1.MoveNext
'   Next i
'   If adoadodc1.RecordCount > 0 Then adoadodc1.MoveFirst
'end 2021/4/16
   
   If adoadodc2.RecordCount > 0 Then
      adoadodc2.MoveFirst
   'Added by Morgan 2015/8/28
      Text1.Enabled = False
      Text6.Enabled = False
      Text7.Enabled = False
      Text8.Enabled = False
   Else
      Text1.Enabled = True
      Text6.Enabled = True
      Text7.Enabled = True
      Text8.Enabled = True
   'end 2015/8/28
   End If
   
'Removed by Morgan 2021/4/16 OraOLEDB會有錯誤，改語法內直接用db函數抓
'   For i = 1 To adoadodc2.RecordCount
'      adoadodc2.Fields(2) = adoadodc2.Fields(2) & PUB_GetRelateCasePropertyName(adoadodc2.Fields(0), "1")
'      adoadodc2.MoveNext
'   Next i
'   If adoadodc2.RecordCount > 0 Then adoadodc2.MoveFirst
'end 2021/4/16
   
   '2011/5/23 End
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示查詢資料(案件基本資料)
'
'*************************************************
'Modify by Amy 2025/08/19 原:Private
Public Sub CaseQuery()
Dim strTM10 As String, strFA119 As String, strTo As String
   
   If Text1 = MsgText(601) Or Text6 = MsgText(601) Or Text7 = MsgText(601) Or Text8 = MsgText(601) Then
      Exit Sub
   End If
   Text2 = CaseNameShow(Text1, Text6, Text7, Text8, 1)
   Text3 = CaseNameShow(Text1, Text6, Text7, Text8, 2)
   Text4 = CaseNameShow(Text1, Text6, Text7, Text8, 3)
   
   'Add By Sindy 2021/1/5 T案要顯示【陸代定稿加註】
   If Text8.Enabled = True And Command5.Enabled = False And TypeName(frmlink) = "Nothing" And Left(Text1, 1) = "T" Then
      strTM10 = GetPrjNation1(Text1 & "-" & Text6 & "-" & Text7 & "-" & Text8)
      strTo = "": strFA119 = ""
      'Modify By Sindy 2020/1/8 請款單只管FC代理人
'      If strTM10 <> "000" Then
'         strTo = PUB_GetFCeMailConText("Main_EMail", Text1, Text6, Text7, Text8, "CF", , True)
'      Else
         strTo = PUB_GetFCeMailConText("Main_EMail", Text1, Text6, Text7, Text8, "FC", , True)
'      End If
      If strTo <> "" Then
         CheckOC3
         strExc(0) = "select fa01,fa02,fa119" & _
                     " from fagent" & _
                     " where fa01='" & Left(strTo, 8) & "' and fa02='" & Mid(strTo, 9, 1) & "'"
         intI = 1
         Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strFA119 = "" & AdoRecordSet3.Fields("FA119")
         End If
         CheckOC3
      End If
      If strFA119 <> "" Then
         MsgBox "【陸代定稿加註】" & vbCrLf & vbCrLf & strFA119, vbInformation
      End If
   End If
   '2021/1/5 END
   
   Call PUB_ChkTemporaryReceipts(Text1, Text6, Text7, Text8) 'Add By Sindy 2014/5/28 檢查是否有暫收款
   AdodcRefresh
End Sub

'Add By Sindy 2009/07/15
Public Sub Text1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(Text1) = False Then
      ' 檢查系統類別
      If IsCorrectSysKind(Text1) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "本所案號中的系統別不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Text1_GotFocus
         GoTo EXITSUB
      End If
      ' 檢查使用者權限
      If IsUserHasRightOfSystem(strUserNum, Text1) = False Then
         '2009/8/21 ADD BY SONIA FCP程序可輸入P及CFP請款單
         If Not (GetStaffDepartment(strUserNum) = "F22" And (Text1 = "P" Or Text1 = "PS" Or Text1 = "CFP" Or Text1 = "CPS")) Then
         '2009/8/21 END
            Cancel = True
            strTit = "資料檢核"
            strMsg = "您沒有使用該系統類別的權限"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            Text1_GotFocus
            GoTo EXITSUB
         End If
      End If
   Else
'      Cancel = True
'      strTit = "資料檢核"
'      strMsg = "本所案號中的系統別不可空白"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      Text1_GotFocus
'      GoTo EXITSUB
   End If
'   Text6.Enabled = True
'   Text7.Enabled = True
'   Text8.Enabled = True
'   Exit Sub
EXITSUB:
'   Text6.Enabled = False
'   Text7.Enabled = False
'   Text8.Enabled = False
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("X") And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
   End If
End Sub

Private Sub Text5_GotFocus()
   CloseIme
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_GotFocus()
   CloseIme
   TextInverse Text6
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
'   Text7 = "0"
'   Text8 = "00"
'   CaseQuery
End Sub

Private Sub Text7_GotFocus()
   CloseIme
   TextInverse Text7
End Sub

'*************************************************
'  顯示資料表
'
'*************************************************
Public Sub FormShow()
   If IsNull(adoacc1k0.Fields("a1k13").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = adoacc1k0.Fields("a1k13").Value
   End If
   If IsNull(adoacc1k0.Fields("a1k14").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = adoacc1k0.Fields("a1k14").Value
   End If
   If IsNull(adoacc1k0.Fields("a1k15").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = adoacc1k0.Fields("a1k15").Value
   End If
   If IsNull(adoacc1k0.Fields("a1k16").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = adoacc1k0.Fields("a1k16").Value
   End If
   Text5 = adoacc1k0.Fields("a1k01").Value
   CaseQuery
   If Len(Text5) = 10 Then
      Command2.Enabled = False
   Else
      Command2.Enabled = True
   End If
   Text5.Tag = Text5 'Added by Morgan 2015/8/28
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
'   CaseQuery
End Sub

Private Sub Text8_GotFocus()
   CloseIme
   TextInverse Text8
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   If adoacc1k0.RecordCount = 0 Then
      Exit Sub
   End If
   CountShow adoacc1k0.Bookmark, adoacc1k0.RecordCount
End Sub

Private Sub Text8_Validate(Cancel As Boolean)
   If Text7 = MsgText(601) Then
      Text7 = "0"
   End If
   If Text8 = MsgText(601) Then
      Text8 = "00"
   End If
   CaseQuery
End Sub

'*************************************************
'  重新整理國外請款資料
'*************************************************
Public Sub Acc1k0Refresh()
On Error GoTo Checking
   If adoacc1k0.State = adStateOpen Then
      adoacc1k0.Close
   End If
   adoacc1k0.CursorLocation = adUseClient
   adoacc1k0.MaxRecords = intMax
   adoacc1k0.Open "select * from acc1k0 where a1k01 >= '" & Text5 & "' and (a1k12 is null or a1k12 <> '') order by a1k01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

