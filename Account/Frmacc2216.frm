VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc2216 
   AutoRedraw      =   -1  'True
   Caption         =   "抵帳作業"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   8760
   Begin VB.TextBox Text13 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00FFFFFF&
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
      Height          =   330
      Left            =   1125
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   540
      Width           =   1092
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00FFFFFF&
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
      Height          =   330
      Left            =   7335
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   540
      Width           =   1092
   End
   Begin VB.TextBox Text10 
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
      Height          =   330
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   3
      Top             =   4950
      Width           =   612
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00C0FFFF&
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
      Height          =   330
      Left            =   2040
      MaxLength       =   15
      TabIndex        =   0
      Top             =   156
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Frmacc2216.frx":0000
      Height          =   1560
      Left            =   240
      TabIndex        =   5
      Top             =   3255
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   2752
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "a1501"
         Caption         =   "帳單編號"
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
         DataField       =   "axf03"
         Caption         =   "本所案號"
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
         DataField       =   "a1505"
         Caption         =   "幣別"
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
         DataField       =   "a1510"
         Caption         =   "台幣金額"
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
      BeginProperty Column04 
         DataField       =   "a1506"
         Caption         =   "外幣金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.000"
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
            ColumnWidth     =   1725.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1769.953
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   569.764
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1830.047
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1860.095
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2216.frx":0015
      Height          =   1560
      Left            =   240
      TabIndex        =   4
      Top             =   1110
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   2752
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
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
         DataField       =   "A1K01"
         Caption         =   "請款編號"
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
         DataField       =   "A1K13||A1K14||A1K15||A1K16"
         Caption         =   "本所案號"
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
         DataField       =   "A1K08"
         Caption         =   "外幣金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "A1K11"
         Caption         =   "外幣現值(NT)"
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
         DataField       =   "A1K09"
         Caption         =   "規費"
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
      BeginProperty Column05 
         DataField       =   "Property"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         Size            =   344
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   1080
         EndProperty
      EndProperty
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
      Height          =   330
      Left            =   6348
      TabIndex        =   13
      Top             =   4950
      Width           =   1932
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
      Height          =   330
      Left            =   4536
      TabIndex        =   12
      Top             =   4950
      Width           =   1812
   End
   Begin VB.TextBox Text6 
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
      Height          =   330
      Left            =   7335
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   2
      Top             =   170
      Width           =   1092
   End
   Begin VB.TextBox Text4 
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
      Height          =   330
      Left            =   4632
      TabIndex        =   9
      Top             =   2700
      Width           =   1440
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
      Height          =   330
      Left            =   3300
      TabIndex        =   8
      Top             =   2700
      Width           =   1332
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
      Height          =   330
      Left            =   5040
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   1
      Top             =   170
      Width           =   1092
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   315
      Left            =   210
      Top             =   3135
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   240
      Top             =   930
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
   Begin MSForms.TextBox Text14 
      Height          =   330
      Left            =   2220
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   540
      Width           =   3930
      VariousPropertyBits=   671105055
      BackColor       =   16777215
      MaxLength       =   50
      Size            =   "6932;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
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
      Height          =   210
      Left            =   360
      TabIndex        =   21
      Top             =   540
      Width           =   675
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "抵帳日期"
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
      Left            =   6300
      TabIndex        =   19
      Top             =   540
      Width           =   900
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "D093010001"
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
      Left            =   1350
      TabIndex        =   18
      Top             =   2745
      Width           =   1200
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "傳票號碼"
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
      Left            =   240
      TabIndex        =   17
      Top             =   2760
      Width           =   900
   End
   Begin VB.Label Label9 
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
      Left            =   225
      TabIndex        =   16
      Top             =   4995
      Width           =   975
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
      Left            =   1905
      TabIndex        =   15
      Top             =   4995
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00404040&
      Height          =   930
      Left            =   225
      Top             =   45
      Width           =   8295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "FC/CF 抵帳編號"
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
      Top             =   170
      Width           =   1812
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   -90
      Top             =   4785
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label6 
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
      Height          =   255
      Left            =   3870
      TabIndex        =   11
      Top             =   4995
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "CF匯率"
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
      Left            =   6360
      TabIndex        =   10
      Top             =   170
      Width           =   852
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   -30
      X2              =   8730
      Y1              =   3075
      Y2              =   3075
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   -30
      X2              =   8730
      Y1              =   3090
      Y2              =   3090
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   2730
      TabIndex        =   7
      Top             =   2745
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "美金匯率"
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
      Left            =   4080
      TabIndex        =   6
      Top             =   170
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc2216"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/09 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、DataGrid2改字型=新細明體-ExtB、Text14
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2011/10/31 CREATE BY SONIA
Option Explicit
Public adoacc1g0 As New ADODB.Recordset
Public adoacc150 As New ADODB.Recordset
Public adoacc1k0 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoadodc2 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adocheck As New ADODB.Recordset
Dim intCounter As Integer
Dim m_intPage As Integer '頁數
Const m_dblLeft As Double = 500 '橫軸偏移值
Dim m_bolAlert As Boolean '檢查分錄提醒
Dim m_strAlertMsg As String

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   'Modified by Lydia 2021/12/09 改成模組
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   Me.Width = 8850
'   Me.Height = 5685
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
   PUB_InitForm Me, 8850, 5800, strBackPicPath1
   'end 2021/12/09
      
   OpenTable
   SumShow1
   SumShow2
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strItemNo = ""
   tool3_enabled
   Select Case strFormLink
      Case "Frmacc2210"
         Frmacc2210.Enabled = True
      Case "Frmacc2220"
         Frmacc2220.Enabled = True
   End Select
   Set Frmacc2216 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc1g0.CursorLocation = adUseClient
   adoacc1g0.Open "select * from acc1g0 where a1g01 = '" & strItemNo & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   FormShow
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select a1k01, a1k13 || a1k14 || a1k15 || a1k16, round((a1k08-nvl(a1k06, 0)) * " & Val(Text2) & ", 2) as a1k11, (a1k08-nvl(a1k06, 0)) AS A1K08, a1k09, nvl(cpm03, cpm04) as Property, a1k30, a1k03 from acc1k0, (select cp01, cp60, min(cp10) as cp10 from acc1k0, caseprogress where a1k01 = cp60 and a1k17 = '" & strItemNo & "' group by cp01, cp60) new, casepropertymap where a1k01 = cp60 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and a1k17 = '" & strItemNo & "' order by a1k01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      Text13 = "" & adoadodc1.Fields("a1k03").Value
   End If
   Text14 = FagentQuery(Text13, 2)
   If Text14 = "" Then
      Text14 = FagentQuery(Text13, 1)
   End If
   If Text14 = "" Then
      Text14 = FagentQuery(Text13, 3)
   End If
   adoadodc2.CursorLocation = adUseClient
   adoadodc2.Open "select * from acc151, acc150 where axf01 = a1501 and a1512 = '" & strItemNo & "' order by a1501, axf03, a1505, a1510, a1506", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc2.Recordset = adoadodc2
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
   Text9 = adoacc1g0.Fields("a1g01").Value
   SetDate Text9
   If IsNull(adoacc1g0.Fields("a1g02").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = adoacc1g0.Fields("a1g02").Value
   End If
   If IsNull(adoacc1g0.Fields("a1g03").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = adoacc1g0.Fields("a1g03").Value
   End If
   If IsNull(adoacc1g0.Fields("a1g10").Value) Then
      Text10 = MsgText(601)
   Else
      Text10 = adoacc1g0.Fields("a1g10").Value
   End If
   
   Me.Label12.Caption = GetSummonsNo("1", "K", Me.Text9.Text)
End Sub

'*************************************************
'  計算並顯示 Adodc1 之合計
'
'*************************************************
Public Sub SumShow1()
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(a1k08-nvl(a1k06, 0)), sum(round((a1k08-nvl(a1k06, 0)) * " & Val(Text2) & ", 2)) from acc1k0 where a1k17 = '" & Text9 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text3 = MsgText(601)
      Else
         Text3 = Format(adoaccsum.Fields(0).Value, FAmount)
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text4 = MsgText(601)
      Else
         Text4 = Format(adoaccsum.Fields(1).Value, FAmount)
      End If
   Else
      Text3 = MsgText(601)
      Text4 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  計算並顯示 Adodc2 之合計
'
'*************************************************
Public Sub SumShow2()
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(round(axf04 * '" & Val(Text6) & "')), sum(axf04) from acc151, acc150 where axf01 = a1501 and a1512 = '" & Text9 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text7 = MsgText(601)
      Else
         Text7 = Format(adoaccsum.Fields(0).Value, DAmount)
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text8 = MsgText(601)
      Else
         Text8 = Format(adoaccsum.Fields(1).Value, FAmount)
      End If
   Else
     Text7 = MsgText(601)
     Text8 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

'抓傳票號碼
Private Function GetSummonsNo(strA1P01 As String, stra1p02 As String, strA1P04 As String)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

StrSQLa = "Select A1P22 From ACC1P0 Where A1P01='" & strA1P01 & "' And A1P02='" & stra1p02 & "' And A1P04='" & strA1P04 & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetSummonsNo = "" & rsA.Fields(0).Value
Else
    GetSummonsNo = ""
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

Private Sub SetDate(p_a1g01 As String)
   Text12 = ""
   strExc(0) = "select a1p18 from acc1g0 x,acc1p0 where a1p04(+)=a1g01 and  a1g01 = '" & p_a1g01 & "' and a1p18>0 and rownum<2"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Text12 = Format(RsTemp.Fields(0), "###/##/##")
   End If
End Sub

