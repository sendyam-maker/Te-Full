VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc41j0 
   AutoRedraw      =   -1  'True
   Caption         =   "非智權結餘轉撥報出傳票產生(隱藏版)"
   ClientHeight    =   4704
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8772
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4704
   ScaleWidth      =   8772
   Begin VB.TextBox TxtSum 
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
      Height          =   315
      Left            =   5050
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2740
      Width           =   1320
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
      Index           =   7
      Left            =   7575
      MaxLength       =   9
      TabIndex        =   6
      Top             =   3600
      Width           =   1000
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
      Index           =   5
      Left            =   1380
      MaxLength       =   5
      TabIndex        =   4
      Top             =   3600
      Width           =   765
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
      Index           =   3
      Left            =   1380
      MaxLength       =   5
      TabIndex        =   2
      Top             =   3240
      Width           =   765
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
      Left            =   8025
      Picture         =   "Frmacc41j0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   10
      ToolTipText     =   "取消"
      Top             =   3975
      Width           =   550
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
      Index           =   6
      Left            =   4560
      MaxLength       =   9
      TabIndex        =   5
      Top             =   3600
      Width           =   1572
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
      Height          =   312
      Index           =   4
      Left            =   4560
      MaxLength       =   3
      TabIndex        =   3
      Top             =   3225
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
      Height          =   600
      Left            =   7320
      Picture         =   "Frmacc41j0.frx":066A
      Style           =   1  '圖片外觀
      TabIndex        =   9
      ToolTipText     =   "清除畫面"
      Top             =   3960
      Width           =   550
   End
   Begin VB.CommandButton CmdUpdAcc 
      BackColor       =   &H00C0FFC0&
      Caption         =   "更正傳票"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7440
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   330
      Width           =   1150
   End
   Begin VB.CommandButton CmdSaveAcc 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生傳票"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6210
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   330
      Width           =   1150
   End
   Begin VB.CommandButton CndRead 
      BackColor       =   &H00C0FFC0&
      Caption         =   "讀取資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3120
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   50
      Width           =   1200
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1800
      TabIndex        =   1
      Top             =   45
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
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
      Bindings        =   "Frmacc41j0.frx":0F34
      Height          =   1995
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   8535
      _ExtentX        =   15050
      _ExtentY        =   3514
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   16
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
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "Dept"
         Caption         =   "業務區"
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
         DataField       =   "Sales"
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
      BeginProperty Column02 
         DataField       =   "SBT04"
         Caption         =   "結餘部門"
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
         DataField       =   "Obj"
         Caption         =   "轉撥對象"
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
         DataField       =   "SBT06"
         Caption         =   "轉撥金額"
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
         DataField       =   "SBT07"
         Caption         =   "對沖其他"
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
         DataField       =   "SBT08"
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
      BeginProperty Column07 
         DataField       =   "SBT01"
         Caption         =   "SBT01"
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
         DataField       =   "SBT02"
         Caption         =   "SBT02"
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
         DataField       =   "SBT03"
         Caption         =   "SBT03"
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
         DataField       =   "SBT05"
         Caption         =   "SBT05"
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
            ColumnWidth     =   1008
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1391.811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1008
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1391.811
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1091.906
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   0
      Top             =   600
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   1800
      TabIndex        =   25
      Top             =   390
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSForms.TextBox Text8 
      Height          =   630
      Left            =   1410
      TabIndex        =   7
      Top             =   3990
      Width           =   5700
      VariousPropertyBits=   -1466941413
      MaxLength       =   200
      ScrollBars      =   2
      Size            =   "10054;1111"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text2 
      Height          =   315
      Index           =   1
      Left            =   2200
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3600
      Width           =   960
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      MaxLength       =   8
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text2 
      Height          =   315
      Index           =   0
      Left            =   2200
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3240
      Width           =   960
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "轉撥金額合計："
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
      Left            =   3480
      TabIndex        =   29
      Top             =   2740
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "智權公司結餘月底轉撥不適用本程式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4680
      TabIndex        =   27
      Top             =   75
      Width           =   3800
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "轉期末傳票日期："
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
      Left            =   120
      TabIndex        =   26
      Top             =   390
      Width           =   1680
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  '透明
      Caption         =   "Lbl1"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   24
      Top             =   390
      Width           =   1680
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "轉傳票號碼："
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
      Left            =   3510
      TabIndex        =   23
      Top             =   390
      Width           =   1680
   End
   Begin VB.Label Label8 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "摘　   要："
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
      TabIndex        =   22
      Top             =   3960
      Width           =   1200
   End
   Begin VB.Label Label7 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "對沖其他："
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
      Left            =   6360
      TabIndex        =   21
      Top             =   3600
      Width           =   1200
   End
   Begin VB.Label Label6 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "轉撥金額："
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
      Left            =   3360
      TabIndex        =   20
      Top             =   3600
      Width           =   1200
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "轉撥對象："
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
      Top             =   3600
      Width           =   1200
   End
   Begin VB.Label Label4 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "(T、P、CFT、CFP)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   17
      Top             =   3240
      Width           =   1995
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "智權人員："
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
      TabIndex        =   16
      Top             =   3240
      Width           =   1200
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "結餘部門："
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
      Left            =   3360
      TabIndex        =   13
      Top             =   3225
      Width           =   1200
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1545
      Left            =   120
      Top             =   3120
      Width           =   8535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "　　　資料年月："
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
      Left            =   120
      TabIndex        =   8
      Top             =   45
      Width           =   1680
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc41j0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/08 Form2.0已修改 Text2(0)/Text2(1)/Text8/DataGrid1
'Create by Amy 2017/04/20
Option Explicit
Const MskFormat As String = "###/##"
' 第一筆資料/最後一筆資料/目前正在顯示
Dim m_FirstKEY As String, m_LastKEY As String, m_CurrKEY As String
Public adoSBT As New ADODB.Recordset
Dim strQ As String, strCmd As String, strAcDate As String
Dim strAxb(14 To 15) As String
Dim strA0b01 As String, strA0b05 As String '會計過帳日/業績輸入關閉年月
Dim bolHasAx210 As Boolean '是否已過帳
Dim oTxt 'Modify by Amy 2021/12/08 原:As TextBox
Dim strMsg As String
'Add by Amy 2017/09/21
Dim bolFirst As Boolean
Dim stDate As String, stYM As String '畫面當月期末保留傳票日/畫面年月
Dim bol0b1HasIns As Boolean 'Acc0b1 是否有Insert 當月資料
Dim HasUpdTag As Boolean 'Add by Amy 2021/09/22 有更新修改Tag

Private Sub CmdSaveAcc_Click()
    
    If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    '存檔前檢查
    If FormCheck(1) = False Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If FormSave(1) = True Then
        MsgBox "已產生傳票！"
        Call ShowBt(0)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdUpdAcc_Click()

    If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    '存檔前檢查
    If FormCheck(2) = False Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If FormSave(2) = True Then
        MsgBox "傳票已更正！"
        Call ShowBt(0)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub CndRead_Click()
    Screen.MousePointer = vbHourglass
    strMsg = "資料！"
    If ExistCheck("SalesBalanceTran", "SBT01", Replace(MaskEdBox1.Text, "/", ""), strMsg, True) = False Then
        'function 會彈訊息
    End If
    Call QueryData
    If strSaveConfirm = MsgText(601) Then Call ShowBt(0)
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub Command1_Click()
    Dim ThisRec
    Dim intI As Integer
    
    Screen.MousePointer = vbHourglass
   
    Adodc1.Recordset.MovePrevious
    If Adodc1.Recordset.BOF = True Then
        Adodc1.Recordset.MoveFirst
        ThisRec = Adodc1.Recordset.Bookmark
    Else
        ThisRec = Adodc1.Recordset.Bookmark
        Adodc1.Recordset.MoveNext
    End If
    strCmd = "Delete SalesBalanceTran Where SBT01=" & Val(Replace(MaskEdBox1.Text, "/", "")) & _
                    " And SBT03='" & ChgSQL(Text1(3)) & "' And SBT04='" & ChgSQL(Text1(4)) & "' And SBT05='" & ChgSQL(Text1(5)) & "' "
    adoTaie.Execute strCmd, intI
    'Add by Amy 2021/09/22 +ChkSetAxb16
    If intI > 0 Then
        Call ChkSetAxb16("Del")
    End If
    Call QueryData
    If strSaveConfirm = MsgText(601) Then Call ShowBt(0)
    Adodc1.Recordset.Requery
    If Adodc1.Recordset.BOF = False Then Adodc1.Recordset.Bookmark = ThisRec
    TextClear
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
    TextClear
    Text1(3).SetFocus
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
    If Not Adodc1.Recordset.EOF Then
        AdodcShow
    End If
End Sub

Private Sub Form_Activate()
    tool1_enabled
    Forms(0).Toolbar1.Buttons.Item(9).Enabled = False
    'Add by Amy 2021/09/08 業績輸入關閉後或當月已過帳,不可再按修改鈕
    If CmdSaveAcc.Enabled = False And CmdUpdAcc.Enabled = False Then
        Forms(0).Toolbar1.Buttons.Item(5).Enabled = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  ' Add by Amy 2021/12/08 Form2.0 記錄鍵盤傳入順序
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(1, KeyCode) 'Add by Amy 2021/12/08 Form2.0
    KeyDefine KeyCode
End Sub

Private Sub Form_Load()
    Dim intX As Integer
    Dim intY As Integer
    Dim sglWidth As Single
    Dim sglHeight As Single
    
    '不使用查詢
    Frmacc0000.Toolbar1.Buttons.Item(9).Enabled = False
    strFormName = Name
    Me.Width = 8895
    Me.Height = 5220
    Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
    
    strA0b01 = GetA0b01(strA0b05)
    TextClear
    FormDisabled
    bolFirst = True
    Call ShowBt(0)
    'Add by Amy 2021/09/08 業績輸入關閉後或當月已過帳,不可再按修改鈕
    If CmdSaveAcc.Enabled = False And CmdUpdAcc.Enabled = False Then
        Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = False
    End If
    RefreshRange
    QueryData
    bolFirst = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    '未存檔或取消不可離開
    If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
        Cancel = 1
        Exit Sub
    End If
    
    If strAxb(14) = MsgText(601) Then
        MsgBox "當月傳票尚未產生相關傳票！", , MsgText(5)
    Else
        If bolNotSave = True Then
            MsgBox "有修改資料但尚未更正傳票,請按「更正傳票」鈕！", , MsgText(5)
            Cancel = True
            Exit Sub
        End If
    End If
    
    'Mark by Amy 2017/09/21 尚未過帳彈訊息-拿掉秀玲
'    If bolHasAx210 = False And strAxB(14) <> MsgText(601) Then
'        MsgBox "請記得過帳！" & vbCrLf & _
'                "在確認專業點數及業務點數相同後,再通知智權主管寫報告！"
'    End If
    '當月業績輸入尚未關閉彈訊息
    If strA0b05 < Left(GetPreMonLastDate(strSrvDate(1)), 5) Then
        MsgBox "請重新產生「實績與結餘分析表」之Excel檔案後" & vbCrLf & _
                "再關閉業績輸入！"
    End If
    'end 2017/09/21
    'Add by Amy 2021/09/22 若有修改發mail給電腦中心
    If HasUpdTag = True Then
        PUB_SendMail strUserNum, "A2004", "", "非當月結餘轉撥傳票產生(隱藏版) 資料有修改", "如摘要"
    End If
    
    strFormName = MsgText(601)
    KeyEnter vbKeyEscape
    tool4_enabled
    MenuEnabled
    strTrackMode = "" 'Add by Amy 2021/12/08 Form2.0 記錄鍵盤傳入順序(清除)
    Call PUB_GetLock("", "Frmacc41j0") 'Add by Amy 2017/09/21
    Set Frmacc41j0 = Nothing
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
    Dim strMsg As String
    
    If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MskFormat Then
        Exit Sub
    End If
    
    strMsg = "轉期末傳票日期"
    If Val(Replace(MaskEdBox1.Text, "/", "")) < Val(Left(業績自動轉傳票啟用年月, 5)) Then
        MsgBox strMsg & "無舊資料！", , MsgText(5)
        Call ClearMaskBox(MaskEdBox2)
        Cancel = True
        MaskEdBox1.SetFocus
        Exit Sub
    End If
    If IsDate(ChangeTStringToWDateString(Replace(MaskEdBox1.Text, "/", "") & "01")) = False Then
        MsgBox strMsg & "輸入錯誤！", , MsgText(5)
        Call ClearMaskBox(MaskEdBox2)
        Cancel = True
        MaskEdBox1.SetFocus
        Exit Sub
    End If
    If Val(Replace(MaskEdBox1.Text, "/", "")) >= Val(Left(strSrvDate(2), 5)) Then
        MsgBox strMsg & "需小於系統月份！", , MsgText(5)
        Call ClearMaskBox(MaskEdBox2)
        Cancel = True
        MaskEdBox1.SetFocus
        Exit Sub
    End If
    
    If strSaveConfirm = MsgText(601) Then Call ShowBt(0)
End Sub

'intChoose:0:判斷當月Axb15及Axb14/1:判斷是否開放Command1/2鈕
Private Sub ShowBt(ByVal intChoose As Integer)
    Dim strTemp As String 'Add by Amy 2022/01/04
    
    CmdSaveAcc.Enabled = False
    CmdUpdAcc.Enabled = False
    Command1.Enabled = False
    Command2.Enabled = False
   
    If intChoose = 0 Then
        'Modify by Amy 2017/09/21 由From_Load搬過來修改
        If bolFirst = True Then
            stDate = GetPreMonLastDate(strSrvDate(1)) '取得上個月最後一天工作日
            stYM = Left(stDate, 5)
        Else
            stYM = Left(Replace(MaskEdBox1.Text, "/", ""), 5)
        End If
    
        '抓非當月轉撥傳票號碼,傳票已過帳或當月傳票未產生,更正傳票鈕不可使用
        Lbl1 = ""
        strExc(0) = bolAcc0b1(3, stYM, strAxb())
        If strExc(0) = "False" Then
            bol0b1HasIns = ExistCheck("Acc0b1", "Axb01", stYM, "", False)
        Else
            bol0b1HasIns = True
        End If
        bolHasAx210 = Pub_ChkAxbPost(strAxb(14), strAxb(14))
        
        '當月期末保留傳票日(strAxb15) 有值 則不可修改,沒值 預設上個月最後一個工作日
        If strAxb(15) = MsgText(601) Then
            MaskEdBox2.Enabled = True
        Else
            stDate = strAxb(15)
            MaskEdBox2.Enabled = False
        End If
        
        '預設上個月
        MaskEdBox1.Mask = ""
        MaskEdBox1.Text = Mid(stYM, 1, 3) & "/" & Mid(stYM, 4, 2)
        MaskEdBox1.Mask = MskFormat
        
        MaskEdBox2.Mask = ""
        MaskEdBox2.Text = CFDate(stDate)
        MaskEdBox2.Mask = DFormat
        'end 2017/09/21
        
        'Modify by Amy 2021/02/05 業績輸入關閉後不可再輸入 原:Val(stYM) >= Val(strA0b05)
        '11001 關閉後(區主管已全確認完)又加SalesPoint 沒有的人員(10051),有輸業績輸入部門一定要重新於點數那支確認是否要報出
        '畫面日期大於業績輸入關閉年月(開放尚未關閉)
        'Modify by Amy 2021/01/04 原:Val(Left(strSrvDate(1), 6)) - 191101 ,11101月輸11012月會判斷錯
        strTemp = Left(GetPreMonLastDate(strSrvDate(1)), 5)
        If Val(stYM) > Val(strA0b05) And Val(stYM) = Val(strTemp) Then
        'end 2021/01/04
            If strAxb(15) <> MsgText(601) Then
                '結餘期末保留傳票日期(Axb15)有值(結餘保留傳票產生需先做),可按「產生傳票鈕」
                If strAxb(14) = MsgText(601) Then
                    CmdSaveAcc.Enabled = True
                '已有傳票號只能按「更正傳票鈕」
                ElseIf bolHasAx210 = False Then
                    CmdUpdAcc.Enabled = True
                End If
            Else
                CmdSaveAcc.Enabled = True
            End If
        End If
    Else
         If bolHasAx210 = False Then
            Command1.Enabled = True '剪刀
            Command2.Enabled = True '垃圾筒
        End If
    End If
   
    Lbl1 = strAxb(14)
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
    Dim strMsg As String
    Dim nResponse
       
    Select Case KeyCode
        Case vbKeyInsert
            'Add by Amy 2021/12/08 Form2.0控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
            If PUB_ChkTrackMode = False Then
                Exit Sub
            End If
            'end 2021/12/08

            If FormCheck(0, "Insert") = False Then Exit Sub
            FormIns
            TextClear
            QueryData
            If Text1(3).Enabled = True Then
                Text1(3).SetFocus
            Else
                Text1(4).SetFocus
            End If
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

Public Function FormCheck(ByVal intCmd As Integer, Optional strKey As String = "") As Boolean
    Dim strLabel As String, bolCancel As Boolean
    Dim strMsg As String, strTmp(1) As String
   
    FormCheck = False: bolCancel = False

    'Modify by Amy 2019/08/05 +strSaveConfirm = MsgText(601),已過帳再按修改彈完訊息無法做任何動作
    If strKey = "F2" And strAxb(14) <> MsgText(601) Then
        MsgBox "此年月已產生傳票,不可再新增！", , MsgText(5)
        MaskEdBox1.SetFocus
        strSaveConfirm = MsgText(601)
        Exit Function
    End If
    
    If strKey = "F3" And bolHasAx210 = True Then
        MsgBox "此年月傳票已過帳,不可再修改！", , MsgText(5)
        MaskEdBox1.SetFocus
        strSaveConfirm = MsgText(601)
        Exit Function
    End If
    
    If strKey = "F5" And strAxb(14) <> MsgText(601) Then
        MsgBox "此年月已產生傳票,不可刪除！", , MsgText(5)
        MaskEdBox1.SetFocus
        strSaveConfirm = MsgText(601)
        Exit Function
    End If
    'end 2019/08/05
    
    strLabel = "資料年月"
    If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
        MsgBox strLabel & "不可為空值！", , MsgText(5)
        MaskEdBox1.SetFocus
        Exit Function
    End If
    '未讀取資料時不可刪除
    If Adodc1.Recordset.RecordCount = 0 And strKey = "F5" Then Exit Function
    
    Call MaskEdBox1_Validate(bolCancel)
    If bolCancel = True Then Exit Function
    
    'Add by Amy 2017/09/21
    strLabel = "轉期末傳票日期"
    If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
        MsgBox strLabel & "不可為空值！", , MsgText(5)
        MaskEdBox2.SetFocus
        Exit Function
    End If
    Call MaskEdBox2_Validate(bolCancel)
    If bolCancel = True Then Exit Function
    'end 2017/09/21
       
        
    If intCmd > 0 Then
        '判斷是否已存檔
        If strSaveConfirm <> MsgText(601) Then
            MsgBox "資料尚未存檔,請先存檔再按「" & IIf(intCmd = 1, "產生", "更正") & "傳票」鈕！", , MsgText(5)
            Exit Function
        End If
        
    End If
        
    If strKey = "Insert" Then
        'Add by Amy 2021/12/08 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
        If PUB_ChkUniText(Me) = False Then
           strControlButton = MsgText(602)
           Exit Function
        End If

        '智權人員
        If Text1(3) = MsgText(601) Then
            MsgBox Left(Label2, Len(Label2) - 1) & "不可為空值！", , MsgText(5)
            If Text1(3).Enabled = True Then Text1(3).SetFocus
            Exit Function
        End If
        Call Text1_Validate(3, bolCancel)
        If bolCancel = True Then Exit Function
        '轉撥對象
        If Text1(5) = MsgText(601) Then
            MsgBox Left(Label5, Len(Label5) - 1) & "不可為空值！", , MsgText(5)
            Call Text1_GotFocus(5)
            Exit Function
        End If
        Call Text1_Validate(5, bolCancel)
        If bolCancel = True Then Exit Function
        'Modify by Amy 2018/09/07 20091不限制-瑞婷
        'Modify by Amy 2019/08/05 +M0100不限制-瑞婷
        If Text1(3) = Text1(5) And Text1(3) <> "20091" And Text1(3) <> "M0100" Then
            MsgBox "轉撥對象不可與智權人員相同！", , MsgText(5)
            Call Text1_GotFocus(5)
            Exit Function
        End If
        '結餘部門
        If Text1(4) = MsgText(601) Then
            MsgBox Left(Label3, Len(Label3) - 1) & "不可為空值！", , MsgText(5)
            Call Text1_GotFocus(4)
            Exit Function
        End If
        Call Text1_Validate(4, bolCancel)
        If bolCancel = True Then Exit Function
        '對沖其他
        If Text1(7) = MsgText(601) Then
            MsgBox Left(Label7, Len(Label7) - 1) & "不可為空值！", , MsgText(5)
            Call Text1_GotFocus(7)
            Exit Function
        End If
        '轉撥金額
        If Val(Text1(6)) = 0 Then
            MsgBox Left(Label6, Len(Label6) - 1) & "不可為 0 或 空值！", , MsgText(5)
            Call Text1_GotFocus(6)
            Exit Function
        Else
            strTmp(0) = SetTranVal
            If Val(Text1(6)) > Val(strTmp(0)) Then
                MsgBox Left(Label6, Len(Label6) - 1) & "輸入值大於結餘轉撥金額！", , MsgText(5)
                Call Text1_GotFocus(6)
                Exit Function
            End If
            strTmp(1) = Val(GetSBT06()) + Val(Text1(6)) - Val(Text1(6).Tag) '減Text1(6).Tag 若修改金額要先減
            If Val(strTmp(1)) > Val(strTmp(0)) And strTmp(1) <> MsgText(601) Then
                MsgBox Left(Label6, Len(Label6) - 1) & "輸入總值大於個人結餘轉撥金額！", , MsgText(5)
                Call Text1_GotFocus(6)
                Exit Function
            End If
        End If
        '判斷 Key重覆不可修改
        If (Text1(3).Tag <> Text1(3) Or Text1(4).Tag <> Text1(4) Or Text1(5).Tag <> Text1(5)) And Text1(3) <> MsgText(601) And ChkSBT() = True Then
            MsgBox "資料已存在！", , MsgText(5)
            If Text1(3).Tag <> Text1(3) And Text1(3).Enabled = True Then
                Call Text1_GotFocus(3)
            ElseIf Text1(4).Tag <> Text1(4) Then
                Call Text1_GotFocus(4)
            Else
                Call Text1_GotFocus(5)
            End If
            Exit Function
        End If
    End If
    
    '判斷A0b10 是否有值
    If Pub_GetAcc0b0("a0b10", "1") <> MsgText(601) Then
        MsgBox MsgText(197), , MsgText(5)
        Exit Function
    End If
    
    
    FormCheck = True
End Function

Public Sub SetData(ByVal strKeyCode As String)
    Select Case strKeyCode
        Case "F2", "F3" '新增/修改
            Call ShowBt(1)
            FormEnabled
            If strKeyCode = "F2" Then
                Lbl1 = ""
                TextClear
                Text1(3).SetFocus
            Else
                Text1(4).SetFocus
            End If
           adoTaie.BeginTrans
        Case "F9" '存檔
            adoTaie.CommitTrans
            RefreshRange
            Lbl1 = ""
            TextClear
            QueryData
            FormDisabled
            Frmacc0000.Toolbar1.Buttons.Item(9).Enabled = False
        Case "F10" '取消
            adoTaie.RollbackTrans
            Lbl1 = ""
            TextClear
            QueryData
            FormDisabled
            Frmacc0000.Toolbar1.Buttons.Item(9).Enabled = False
        Case "FirstRec"
            If m_FirstKEY <> MsgText(601) Then
                MaskEdBox1.Text = Mid(m_FirstKEY, 1, 3) & "/" & Mid(m_FirstKEY, 4, 2)
            End If
            Call QueryData
        Case "LastRec"
            If m_LastKEY <> MsgText(601) Then
                MaskEdBox1.Text = Mid(m_LastKEY, 1, 3) & "/" & Mid(m_LastKEY, 4, 2)
            End If
            Call QueryData
        Case "NextRec"
            Call GetCurrRecord(strKeyCode)
        Case "PreRec"
            Call GetCurrRecord(strKeyCode)
        Case Else
    End Select
End Sub

'intCmd:1-產生傳票/2-更正傳票
 Private Function FormSave(ByVal intCmd As Integer) As Boolean
    Dim stA0202 As String, stAx203 As String, stAx205(1) As String
    
On Error GoTo ErrHand

    FormSave = False
    
    adoTaie.BeginTrans
    adoTaie.Execute "Update Acc0b0 Set a0b10= '01' Where a0b04 = '1'"
    '按「更正傳票」鈕(刪除傳票檔資料再新增加)
    If intCmd = 2 Then
        strCmd = "Update Acc020 set A0209=" & Val(strSrvDate(2)) & ",A0210=" & ServerTime & ",A0211='" & strUserNum & "' " & _
                       "Where A0201='1' And A0202='" & strAxb(14) & "' "
        adoTaie.Execute strCmd
        strCmd = "Delete From Acc021 Where Ax201='1' And Ax202='" & strAxb(14) & "' "
        adoTaie.Execute strCmd
    End If
        
    With Adodc1.Recordset
        .MoveFirst
        If strAxb(14) = MsgText(601) Then
            stA0202 = AccAutoNo(MsgText(801), 4, Val(Left(strAcDate, 3)), Val(Mid(strAcDate, 4, 2)))
            strCmd = AccSaveAutoNo(MsgText(801), Right(stA0202, 4), Val(Left(strAcDate, 3)), Val(Mid(strAcDate, 4, 2)))
        Else
            stA0202 = strAxb(14)
        End If
        If intCmd = 1 Then
            strCmd = "Insert Into Acc020 (a0201,a0202,a0205,a0208,a0206,a0207) " & _
                     "Values('1','" & stA0202 & "', " & Val(stDate) & ",'" & strUserNum & "'," & Val(strSrvDate(2)) & "," & ServerTime & ")"
            adoTaie.Execute strCmd
        End If
            
        Do While .EOF = False
            'Memo 此有會計科目有增加需確認 [傳票輸入] 是否需修改
            Select Case "" & .Fields("SBT04")
                Case "T"
                    stAx205(0) = "249101"
                    stAx205(1) = "410103"
                Case "P"
                    stAx205(0) = "249102"
                    stAx205(1) = "411103"
                Case "CFT"
                    stAx205(0) = "249103"
                    stAx205(1) = "412101"
                Case "CFP"
                    stAx205(0) = "249104"
                    stAx205(1) = "413101"
            End Select
            
            '取得流水號
            stAx203 = GetSeqNo("1", stA0202)
            '借方
            strCmd = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax209,ax212,ax213) " & _
                            "Values('1','" & stA0202 & "', '" & stAx203 & "','" & .Fields("SBT04") & "','" & stAx205(0) & "'," & _
                            Val(.Fields("SBT06")) & ",0,'" & .Fields("SBT03") & "','" & .Fields("SBT08") & "','" & .Fields("SBT07") & "') "
            adoTaie.Execute strCmd
            
            '取得流水號
            stAx203 = GetSeqNo("1", stA0202)
            '貸方
            strCmd = "Insert Into Acc021 (ax201,ax202,ax203,ax204,ax205,ax206,ax207,ax209,ax212,ax213) " & _
                            "Values('1','" & stA0202 & "', '" & stAx203 & "','" & .Fields("SBT04") & "','" & stAx205(1) & "'," & _
                             "0," & Val(.Fields("SBT06")) & ",'" & .Fields("SBT05") & "','" & .Fields("SBT08") & "','" & .Fields("SBT07") & "') "
            adoTaie.Execute strCmd
            .MoveNext
        Loop
    End With
    
    '更新Acc01b對應傳票日期及號碼
    strCmd = ""
    If intCmd = 1 Then
        If bol0b1HasIns = False Then
             strCmd = "Insert Into Acc0b1 (axb01,axb15,axb14) Values(" & Val(strAcDate) & "," & Val(stDate) & ",'" & stA0202 & "')"
        Else
            If strAxb(15) = MsgText(601) Then strCmd = ",axb15=" & stDate
            strCmd = "Update Acc0b1 Set axb14='" & stA0202 & "'" & strCmd & " Where axb01=" & Val(strAcDate)
        End If
        adoTaie.Execute strCmd
    End If
    strAxb(14) = stA0202: Lbl1 = strAxb(14)
    
    adoTaie.Execute "Update Acc0b0 Set a0b10= null Where a0b04 = '1'"
    
    adoTaie.CommitTrans
    FormSave = True
    Exit Function
ErrHand:
    adoTaie.RollbackTrans
    MsgBox Err.Description, , MsgText(5)
End Function

Public Sub FormDel(ByVal intCmd As Integer)
    
    '刪除畫面資料年月SalesBalanceTran當月全部資料
    strCmd = "Delete SalesBalanceTran Where SBT01=" & strAcDate
    adoTaie.Execute strCmd
    
    Adodc1.Recordset.Requery
    Lbl1 = ""
    TextClear
    FormDisabled
    MsgBox strAcDate & " 全部資料已刪除！", vbInformation
    If strAcDate = m_FirstKEY Or strAcDate = m_LastKEY Then
        RefreshRange
    End If
    GetCurrRecord ("PreRec")
End Sub

Private Function FormIns() As Boolean
    Dim stDept As String 'Add by Amy 2017/09/21
    
    If ChkSBT(True) = True Then
        strCmd = "Delete SalesBalanceTran Where SBT01=" & Val(Replace(MaskEdBox1.Text, "/", "")) & _
                    " And SBT03='" & ChgSQL(Text1(3).Tag) & "' And SBT04='" & ChgSQL(Text1(4).Tag) & "' And SBT05='" & ChgSQL(Text1(5).Tag) & "' "
        adoTaie.Execute strCmd
    End If
       
    stDept = GetST15(Text1(3), Val(Replace(MaskEdBox1.Text, "/", "")) + 191100) 'Modify by Amy 2021/02/19 原:PUB_GetStaffST15
    strCmd = "Insert Into SalesBalanceTran (SBT01,SBT02,SBT03,SBT04,SBT05,SBT06,SBT07,SBT08,SBT09,SBT10) Values" & _
                "(" & Val(Replace(MaskEdBox1.Text, "/", "")) & ",'" & ChgSQL(stDept) & "','" & ChgSQL(Text1(3)) & "','" & ChgSQL(Text1(4)) & "','" & ChgSQL(Text1(5)) & "'," & _
                Val(Text1(6)) & ",'" & ChgSQL(Text1(7)) & "','" & ChgSQL(Text8) & "'," & Val(strSrvDate(2)) & "," & ServerTime & ")"
    adoTaie.Execute strCmd
    'Modify by Amy 2017/09/21 智權點數開放後若SalesPoint無資料,則需新增人員至SalesPoint
    If ExistCheck("SalesPoint", "SP01", Val(Replace(MaskEdBox1.Text, "/", "") + 191100), strMsg, False) = True Then
        'Modify by Amy 2021/07/22 +if M0100/M0109不需新增
        If InStr(不需新增SalesPoint人員, Text1(3)) = 0 Then
            '若SalesPoint無此轉發人員資料,則需新增人員至SalesPoint
            If ExistCheck("SalesPoint", "SP01||SP02", Val(Replace(MaskEdBox1.Text, "/", "") + 191100) & Text1(3), strMsg, False) = False Then
                strCmd = "Insert Into SalesPoint (SP01,SP02,SP48) Values(" & Val(Replace(MaskEdBox1.Text, "/", "")) + 191100 & ",'" & ChgSQL(Text1(3)) & "','" & ChgSQL(GetST15(Text1(3))) & "') "
                adoTaie.Execute strCmd
            End If
        End If
        'Modify by Amy 2021/07/22 +if M0100/M0109不需新增
        If InStr(不需新增SalesPoint人員, Text1(5)) = 0 Then
            '若SalesPoint無此轉發人員資料,則需新增人員至SalesPoint
            If ExistCheck("SalesPoint", "SP01||SP02", Val(Replace(MaskEdBox1.Text, "/", "") + 191100) & Text1(5), strMsg, False) = False Then
                strCmd = "Insert Into SalesPoint (SP01,SP02,SP48) Values(" & Val(Replace(MaskEdBox1.Text, "/", "")) + 191100 & ",'" & ChgSQL(Text1(5)) & "','" & ChgSQL(GetST15(Text1(5))) & "') "
                adoTaie.Execute strCmd
            End If
        End If
        'Add by Amy 2021/07/22 +判斷新增或修改人員且已產生傳票,發信給財務
        'Modify by Amy 2021/08/03 原判斷已產生SalesBalance,才發mail
        'Modify by Amy 2021/09/22 改寫 Tag
        Call ChkSetAxb16("Ins")
'        If ExistCheck("SalesBalance", "SB01", Replace(MaskEdBox1.Text, "/", ""), strMsg, False) = True Then
'            If InStr(不需新增SalesPoint人員, Text1(3)) = 0 Or InStr(不需新增SalesPoint人員, Text1(5)) = 0 Then
'                Call ChkSendMail
'            End If
'        End If
    End If
End Function

Private Sub AdodcShow()
    TextClear
    If Adodc1.Recordset.RecordCount > 0 Then
        Text1(3) = "" & Adodc1.Recordset.Fields("SBT03")
        Text1(3).Tag = Text1(3)
        Text2(0) = "" & Adodc1.Recordset.Fields("Sales")
        Text1(4) = "" & Adodc1.Recordset.Fields("SBT04")
        Text1(4).Tag = Text1(4)
        Text1(5) = "" & Adodc1.Recordset.Fields("SBT05")
        Text1(5).Tag = Text1(5)
        Text2(1) = "" & Adodc1.Recordset.Fields("Obj")
        Text1(6) = "" & Adodc1.Recordset.Fields("SBT06")
        Text1(6).Tag = Text1(6)
        Text1(7) = "" & Adodc1.Recordset.Fields("SBT07")
        Text8 = "" & Adodc1.Recordset.Fields("SBT08") 'Modify by Amy 2021/12/08 原:Text1(8)
    End If
End Sub

Private Sub FormEnabled()
    For Each oTxt In Text1
        oTxt.Enabled = True
    Next
End Sub

Private Sub FormDisabled()
    For Each oTxt In Text1
        oTxt.Enabled = False
    Next
End Sub

Private Sub TextClear()
    For Each oTxt In Text1
        oTxt.Text = ""
        oTxt.Tag = ""
    Next
    For Each oTxt In Text2
        oTxt.Text = ""
    Next
    'Add by Amy 2021/12/08 Text1(8)改為Text8
    Text8 = ""
    Text8.Tag = ""
    TxtSum = "" 'Add by Amy 2019/01/14
End Sub

'Add by Amy 2017/09/21 當月期末保留傳票日
Private Sub MaskEdBox2_Validate(Cancel As Boolean)
    Dim strLabel As String, stMsg As String
    
    If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
        Exit Sub
    End If
    
    strLabel = "轉期末傳票日期"
    If IsDate(ChangeTStringToWDateString(FCDate(MaskEdBox2.Text))) = False Then
        MsgBox strLabel & "輸入錯誤！", , MsgText(5)
        Cancel = True
        MaskEdBox2.SetFocus
        Exit Sub
    End If
    If Val(Left(FCDate(MaskEdBox2.Text), 5)) >= Val(Left(strSrvDate(2), 5)) _
     And Val(FCDate(MaskEdBox2.Text)) >= Val(strA0b05) Then
        MsgBox strLabel & "需小於系統月份且大於業績輸入關閉日！", , MsgText(5)
        Cancel = True
        MaskEdBox2.SetFocus
        Exit Sub
    End If
    If ChkWorkDay(Val(FCDate(MaskEdBox2.Text)) + 19110000) = False Then
        MsgBox strLabel & "需是工作日！", , MsgText(5)
        Cancel = True
        MaskEdBox2.SetFocus
        Exit Sub
    End If
    If MaskEdBox2.Enabled = True Then
        If ChkWorkData("1", DBDATE(MaskEdBox2.Text), stMsg) = False Then
            MsgBox strLabel & stMsg, , MsgText(5)
            Cancel = True
            MaskEdBox2.SetFocus
            Exit Sub
        End If
    End If
   
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    If Index = 8 Then
        OpenIme
    Else
        CloseIme
    End If
    TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    '智權人員/轉撥對象/結餘部門(T、P、CFT、CFP)
    If Index >= 3 And Index <= 5 Then
        KeyAscii = UpperCase(KeyAscii)
        'Modify by Amy 2021/12/08 Text1(8)改為Text8
        If Index = 3 Then Text1(6) = "": Text1(7) = "": Text8 = "": Text2(0) = ""
        If Index = 4 Then Text1(6) = ""
        If Index = 5 Then Text8 = "": Text2(1) = ""
        'end 2021/12/08
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Dim stDep As String
    
    If Text1(Index) = MsgText(601) Then Exit Sub
    
    '智權人員有輸 預帶「對沖其他」欄位
    If Index = 3 Then
        If Text1(7) = MsgText(601) Then
            stDep = PUB_GetStaffST15(Text1(3), 1)
            If Left(stDep, 1) = "S" Then
                Text1(7) = "結餘" & PUB_GetZone(stDep, True)
            Else
                Text1(7) = "結餘總"
            End If
        End If
    'Modify by Amy 2021/12/08 Text1(8)改為Text8
        If Text1(5) <> MsgText(601) And Text8 = MsgText(601) Then
            Text8 = Text2(0) & "結餘轉" & Text2(1)
        End If
    ElseIf Index = 5 And Text1(3) <> MsgText(601) And Text8 = MsgText(601) Then
        Text8 = Text2(0) & "結餘轉" & Text2(1)
    End If
    'end 2021/12/08
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
    Dim strName As String, strTmp As String

    If Text1(Index) = "" Then Exit Sub
    
    Select Case Index
        '智權人員/轉撥對象
        Case 3, 5
            'Modify by Amy 2019/02/13 不可輸 S字頭編號 ex:S29 目標
            If Left(Text1(Index), 1) = "S" Then
                MsgBox IIf(Index = 3, "智權人員", "轉撥對象") & "不可輸入S字頭編號 ！"
                Call Text1_GotFocus(Index)
                Cancel = True
                Exit Sub
            '智權人員可以為離職人員,但轉撥對象不可
            ElseIf PUB_GetStaffNameDept(Text1(Index), strName, strTmp, True, IIf(Index = 3, True, False)) = False Then
                If Index = 5 Then
                    Call Text1_GotFocus(Index)
                    Cancel = True
                    Exit Sub
                End If
            End If
            If Index = 3 Then
                Text2(0) = strName
                If Text1(4) <> MsgText(601) And Text1(6) = MsgText(601) Then
                    Text1(6) = Val(SetTranVal) - Val(GetSBT06) + Val(Text1(6).Tag)
                End If
                'Add by Amy 2018/12/04 20091預帶摘要若欄位未跳離不會預帶
                'Modify by Amy 2021/12/08 Text1(8)改為Text8
                If Trim(Text8) = MsgText(601) Then Call Text1_LostFocus(Index)
            Else
                Text2(1) = strName
            End If
        '結餘部門
        Case 4
            If Text1(Index) <> "T" And Text1(Index) <> "P" And Text1(Index) <> "CFT" And Text1(Index) <> "CFP" Then
                Cancel = True
                MsgBox Left(Label3, Len(Label3) - 1) & "輸入錯誤請修正！", , MsgText(5)
                Call Text1_GotFocus(Index)
                Exit Sub
            End If
            '預帶轉撥金額
            If Trim(Text1(6)) = MsgText(601) Then
                If Text1(4).Tag <> MsgText(601) And Text1(4) <> Text1(4).Tag Then
                    Text1(6) = Val(SetTranVal)
                Else
                    Text1(6) = Val(SetTranVal) - Val(GetSBT06) + Val(Text1(6).Tag)
                End If
            End If
        '轉撥金額
        Case 6
            If IsNumeric(Text1(Index)) = False Then
                Cancel = True
                MsgBox Left(Label6, Len(Label6) - 1) & "需輸入數字", , MsgText(5)
                Call Text1_GotFocus(Index)
                Exit Sub
            End If
    End Select
End Sub

'以結餘部門之相對應科目及智權人員抓Acc021之餘額預設轉撥金額
Private Function SetTranVal() As String
    Dim RsQ As New ADODB.Recordset
    Dim stQ As String
    Dim intQ As Integer
    
    SetTranVal = ""
    If Text1(3) = MsgText(601) And Text1(4) = MsgText(601) Then Exit Function
    
    Select Case Text1(4)
        Case "T"
            stQ = stQ & "And AX205='249101' "
        Case "P"
            stQ = stQ & "And AX205='249102' "
        Case "CFT"
            stQ = stQ & "And AX205='249103' "
        Case "CFP"
            stQ = stQ & "And AX205='249104' "
    End Select
    stQ = stQ & "And AX209='" & Text1(3) & "' "
    If strAxb(14) <> MsgText(601) Then stQ = stQ & " And AX202<>'" & strAxb(14) & "' "
    
    'Modify by Amy 2018/01/18 只抓1公司
    stQ = "Select Sum(ax207-ax206) as stSum From Acc021,Acc020 " & _
             "Where A0201(+)=AX201 And A0202(+)=AX202 And AX201='1' " & stQ & _
             "Group by AX209"
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, stQ)
    If intQ = 1 Then
        SetTranVal = Val("" & RsQ.Fields("stSum"))
    End If
    RsQ.Close
    
End Function

Private Sub RefreshRange()
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset

    strSql = "Select * From (" & _
               "Select Min(SBT01) SBT01 From SalesBalanceTran " & _
               "Where  SBT01>=" & Val(業績自動轉傳票啟用年月) & _
               ") Where SBT01 is not null"
   
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount > 0 Then
        If IsNull(rsTmp.Fields("SBT01")) = False Then: m_FirstKEY = rsTmp.Fields("SBT01")
    End If
    rsTmp.Close

    strSql = "Select * From (" & _
                "Select Max(SBT01) SBT01 From SalesBalanceTran " & _
                "Where  SBT01>=" & Val(業績自動轉傳票啟用年月) & _
                ") Where SBT01 is not null"
 
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("SBT01")) = False Then: m_LastKEY = rsTmp.Fields("SBT01")
    End If
    rsTmp.Close
    
    Set rsTmp = Nothing
End Sub

'設定m_CurrKEY 上/下一筆用
Private Sub GetCurrRecord(ByVal stKey As String)
    Dim strSql As String, strSign As String
    Dim rsTmp As New ADODB.Recordset
            
    If stKey = "PreRec" Then
        strSql = "Max(SBT01)"
        strSign = "<"
    Else
        strSql = "Min(SBT01)"
        strSign = ">"
    End If

    strSql = "Select * From (" & _
                "Select " & strSql & "SBT01 From SalesBalanceTran " & _
               "Where  SBT01" & strSign & Val(m_CurrKEY) & _
               ") Where SBT01 is not null"
   
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount > 0 Then
        If IsNull(rsTmp.Fields("SBT01")) = False Then
            m_CurrKEY = rsTmp.Fields("SBT01")
            MaskEdBox1.Text = Mid(m_CurrKEY, 1, 3) & "/" & Mid(m_CurrKEY, 4, 2)
            Call QueryData
        End If
    Else
        If stKey = "PreRec" Then
            MsgBox MsgText(9008), , MsgText(5)
        Else
            MsgBox MsgText(9009), , MsgText(5)
        End If
    End If
    
    rsTmp.Close
    
    Set rsTmp = Nothing
End Sub

'顯示資料
Private Sub QueryData()

On Error GoTo Checking
    
    strAcDate = Replace(MaskEdBox1.Text, "/", "")
    
    If adoSBT.State = adStateOpen Then adoSBT.Close
    adoSBT.CursorLocation = adUseClient
    strQ = "Select A0902 as Dept,S.St02 as Sales,SBT04,SBT05,O.St02 as Obj,SBT06,SBT07,SBT08,SBT01,SBT02,SBT03,SBT05 " & _
                "From SalesBalanceTran,Staff S,Staff O,Acc090 " & _
                "Where SBT01=" & Val(strAcDate) & " And SBT03=S.St01(+) And SBT05=O.St01(+) And S.St15=A0901(+) " & _
                "Order by S.St15,SBT02"
    adoSBT.Open strQ, adoTaie, adOpenStatic, adLockReadOnly

    Set Adodc1.Recordset = adoSBT
    Adodc1.Recordset.Requery
    If Adodc1.Recordset.RecordCount > 0 Then
        m_CurrKEY = strAcDate
    End If
    Call SumShow 'Add by Amy 2019/01/14
    If strSaveConfirm = MsgText(601) Then Call ShowBt(0)
    Exit Sub
    
Checking:
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Function bolNotSave() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim intQ As Integer
    Dim stSBDT As String
    
    '取得非當月結餘轉撥修改時間
    strQ = "Select Max(SBT09||Decode(length(SBT10),5,'0'||SBT10,SBT10)) as DT From SalesBalanceTran " & _
            "Where SBT01=" & strAcDate & " Having Max(SBT09||Decode(length(SBT10),5,'0'||SBT10,SBT10)) is not null"
    intQ = 1
    Set rsTmp = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
         stSBDT = "" & rsTmp.Fields("DT")
    End If
    rsTmp.Close
    If stSBDT = "" Then Exit Function
    
    strQ = "Select A0206||Decode(length(A0207),5,'0'||A0207,A0207) as DT From Acc020 Where A0201='1' And A0202='" & strAxb(14) & "' " & _
              "And A0206||Decode(length(A0207),5,'0'||A0207,A0207)<" & Val(stSBDT) & " And A0209||A0210 is null " & _
    "Union Select A0209||Decode(length(A0210),5,'0'||A0210,A0210) as DT From Acc020 Where A0201='1' And A0202='" & strAxb(14) & "' " & _
              "And A0209||Decode(length(A0210),5,'0'||A0210,A0210)<" & Val(stSBDT) & " And A0209||A0210 is not null "
    intQ = 1
    Set rsTmp = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        bolNotSave = True
    End If
    rsTmp.Close
End Function

Private Function ChkSBT(Optional ByVal bolOldData As Boolean = False) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim intQ As Integer
    
    ChkSBT = False
  
    If bolOldData = True Then
         strQ = "Select * From SalesBalanceTran Where SBT01=" & Val(Replace(MaskEdBox1.Text, "/", "")) & _
                  " And SBT03='" & Text1(3).Tag & "' And SBT04='" & Text1(4).Tag & "' And SBT05='" & Text1(5).Tag & "' "
    Else
        strQ = "Select * From SalesBalanceTran Where SBT01=" & Val(Replace(MaskEdBox1.Text, "/", "")) & _
                  " And SBT03='" & Text1(3) & "' And SBT04='" & Text1(4) & "' And SBT05='" & Text1(5) & "' "
    End If
    intQ = 1
    Set rsTmp = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        ChkSBT = True
    End If
    rsTmp.Close
    
End Function

'取得某智權人員同一結餘部門轉撥出去之轉撥金額
Private Function GetSBT06() As String
    Dim rsTmp As New ADODB.Recordset
    Dim intQ As Integer
    
    GetSBT06 = ""
    
    strQ = "Select Sum(SBT06) as SBT06 From SalesBalanceTran Where SBT01=" & Val(Replace(MaskEdBox1.Text, "/", "")) & _
                " And SBT03='" & Text1(3) & "' And SBT04='" & Text1(4) & "' Group by SBT01,SBT03,SBT04 "
    intQ = 1
    Set rsTmp = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        GetSBT06 = Val("" & rsTmp.Fields("SBT06"))
    End If
    rsTmp.Close
End Function

Private Sub ClearMaskBox(ByRef MaskBox As MaskEdBox)
    MaskBox.Mask = ""
    MaskBox.Text = ""
    MaskBox.Mask = DFormat
End Sub

'Add by Amy 2019/01/14 增加合計
Private Sub SumShow()
    Dim RsSum As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    strQ = "Select Sum(SBT06) as SBT06 From SalesBalanceTran Where SBT01=" & _
              Val(Replace(MaskEdBox1.Text, "/", ""))
    intQ = 1
    Set RsSum = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        TxtSum = Format("" & RsSum.Fields("SBT06"), FDollar)
    End If
    RsSum.Close
End Sub

'Add by Amy 2021/02/05 SalesPoint 已確認
Private Function ChkIsAccept(ByVal stDept As String) As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, strWhere As String, intQ As Integer
    
    ChkIsAccept = False
    strWhere = "And SP48='" & stDept & "' And SP45 is not null "
    strQ = ChkPointAcceptSql(Val(Replace(MaskEdBox1.Text, "/", "")) + 191100, Me.Name, 9, strWhere)
    If InStr(strQ, "請洽電腦中心") > 0 Then
        Exit Function
    End If
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then ChkIsAccept = True
    RsQ.Close
    Set RsQ = Nothing
End Function

'Add by Amy 2021/07/22 寄信設定
Private Sub ChkSendMail()
    Dim stDept(1) As String, bolAccept(1) As Boolean, stTO(1) As String, stAreaManNo(1) As String, stSubject As String
    Dim bolSpecMail(1) As Boolean
    Dim stTmp As String, stTmp2 As String, stTmp3 As String
    
    
    '智權人員
    stDept(0) = GetST15(Text1(3), , Val(Replace(MaskEdBox1.Text, "/", "")) + 191100)
    bolAccept(0) = ChkIsAccept(stDept(0))
    
    '轉撥人員
    stDept(1) = GetST15(Text1(5), , Val(Replace(MaskEdBox1.Text, "/", "")) + 191100)
    bolAccept(1) = ChkIsAccept(stDept(1))
        
    '**** 目前只發財務,因未存檔(存檔需判斷條件太多)按Insert就觸發,故先Mark 下方程式 ***
'    '判斷江郁仁是否請假一整天
'    If stDept(0) = "F11" Or stDept(1) = "F11" Then stTmp = Set98020Ag
'    If stTmp <> MsgText(601) Then stTmp = Mid(stTmp, 2)
'
'    '有修改智權人員 or 轉撥對象 or 轉撥金額才發 mail
    If Text1(3).Tag & Text1(5).Tag & Text1(6).Tag <> Text1(3) & Text1(5) & Text1(6) Then
'        stAreaManNo(0) = GetDeptMan(stDept(0)) '智權人員的區主管
'        stAreaManNo(1) = GetDeptMan(stDept(1)) '轉撥對象的區主管
'        If stDept(0) = "F11" Then stAreaManNo(0) = "98020"
'        If stDept(1) = "F11" Then stAreaManNo(1) = "98020"
'
'
'        '不需輸 智權點數實績與結餘輸入部門 且 需輸但不是特殊員編(F11部門要通知) 不需通知
'        If InStr(智權點數實績與結餘輸入部門, Left(stDept(0), 1)) = 0 Or stDept(0) = "P10" Or Text1(3) = "20091" _
'          Or (InStr(Replace(智權點數實績與結餘輸入部門, "S", ""), Left(stDept(0), 1)) > 0 And InStr("F4106;F4107", Text1(3)) = 0) Then
'           stAreaManNo(0) = ""
'        End If
'        If InStr(智權點數實績與結餘輸入部門, Left(stDept(1), 1)) = 0 Or stDept(1) = "P10" Or Text1(5) = "20091" _
'          Or (InStr(Replace(智權點數實績與結餘輸入部門, "S", ""), Left(stDept(1), 1)) > 0 And InStr("F4106;F4107", Text1(5)) = 0) Then
'            stAreaManNo(1) = ""
'        End If
'
'        '智權部/外商 加發個人
'        If stAreaManNo(0) <> MsgText(601) Then
'            If InStr("F4106;F4107", Text1(3)) > 0 Then
'                stTmp2 = PUB_GetST14(Text1(3))
'            Else
'                stTmp2 = Text1(3)
'            End If
'        End If
'        '智權部/外商 加發個人
'        If stAreaManNo(1) <> MsgText(601) Then
'            If InStr("F4106;F4107", Text1(5)) > 0 Then
'                stTmp3 = PUB_GetST14(Text1(5))
'            Else
'                stTmp3 = Text1(5)
'            End If
'        End If
'        'F11部門區主管為江郁仁,若請假要發江郁仁及葉易雲(不發人事職代)
'        If stTmp <> MsgText(601) Then
'            If stDept(0) = "F11" Then
'                stAreaManNo(0) = stTmp
'                bolSpecMail(0) = True
'            End If
'            If stDept(1) = "F11" Then
'                stAreaManNo(1) = stTmp
'                bolSpecMail(1) = True
'            End If
'        End If
'
'        If stAreaManNo(0) <> MsgText(601) Then stTo(0) = ";" & stAreaManNo(0)
'        If stAreaManNo(1) <> MsgText(601) Then stTo(1) = ";" & stAreaManNo(1)
'        If stTo(0) = stTo(1) Then
'            stTo(1) = "" '同部門只發一次
'            If InStr(stTo(0), stTmp2) = 0 Then stTo(0) = stTo(0) & ";" & stTmp2
'            If InStr(stTo(0), stTmp3) = 0 Then stTo(0) = stTo(0) & ";" & stTmp3
'        Else
'            If stTo(0) <> MsgText(601) Then
'                If InStr(stTo(0), stTmp2) = 0 Then stTo(0) = stTo(0) & ";" & stTmp2
'            End If
'            If stTo(1) <> MsgText(601) Then
'                If InStr(stTo(1), stTmp3) = 0 Then stTo(1) = stTo(1) & ";" & stTmp3
'            End If
'        End If
'
'        stSubject = "點數有新增或修改 請重新進入每月點數查詢／輸入作業操作並存檔"
'        If stTo(0) <> MsgText(601) Then
'            PUB_SendMail strUserNum, Mid(stTo(0), 2), "", GetPrjSalesNM(Text1(3)) & "(" & Text1(3) & ") " & stSubject, "如主旨", , , , , , , , , , bolSpecMail(0)
'        End If
'        If stTo(1) <> MsgText(601) Then
'            PUB_SendMail strUserNum, Mid(stTo(1), 2), "", GetPrjSalesNM(Text1(5)) & "(" & Text1(5) & ") " & stSubject, "如主旨", , , , , , , , , , bolSpecMail(1)
'        End If
        '**** End 目前只發財務,因未存檔(存檔需判斷條件太多)按Insert就觸發 ***
        
        '通知財務是否需刪畫面年月 智權報出結餘 (SalesBalance) 資料
        stTO(0) = Pub_GetSpecMan("財務處總帳人員")
        stTO(1) = "A2004"
        stSubject = GetPrjSalesNM(Text1(3)) & "(" & Text1(3) & ") 與" & GetPrjSalesNM(Text1(5)) & "(" & Text1(5) & ") 資料有修改" & "請確認後續處理！"
        stTmp = "系統提醒：" & vbCrLf & _
                      "非當月結餘轉撥傳票產生 資料有變動，若有更改人員或點數，請於先開放區主管輸入，待區主管再確認後，" & vbCrLf & _
                      "請通知電腦中心修改項目，判斷是否刪除 " & MaskEdBox1.Text & "月 智權報出結餘 (SalesBalance) 資料。"
        PUB_SendMail strUserNum, stTO(0), "", stSubject, stTmp, , , , , , stTO(1)
    End If
End Sub

'Add by Amy 2021/09/22 結餘資料有修改寫Tag (Axb16)
Private Sub ChkSetAxb16(ByVal stState As String)
    '已產畫面當月結餘保留資料
    If ExistCheck("SalesBalance", "SB01", Replace(MaskEdBox1.Text, "/", ""), strMsg, False) = True Then
        'Mark by Amy 2024/06/04 M0100也可以輸,只是SalesPoint 不需新增(用算的不需再記錄),若關閉後又加不會更新Axb16,會導致SalesBalance(frmacc41H0)不會重抓
        'If InStr(不需新增SalesPoint人員, Text1(3)) = 0 Or InStr(不需新增SalesPoint人員, Text1(5)) = 0 Then
            '有修改智權人員 or 轉撥對象 or 轉撥金額 寫Axb16
            If stState = "Ins" And Text1(3).Tag & Text1(5).Tag & Text1(6).Tag <> Text1(3) & Text1(5) & Text1(6) Then
                If WirteAxb16(stYM, "Y") = True Then
                    HasUpdTag = True
                End If
            '刪除 寫Axb16
            ElseIf WirteAxb16(stYM, "Y") = True Then
                HasUpdTag = True
            End If
        'End If
    End If
End Sub

