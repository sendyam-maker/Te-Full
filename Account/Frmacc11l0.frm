VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc11l0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "收文金額分配作業"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   7485
   Begin VB.TextBox txtDate 
      BackColor       =   &H00E0E0E0&
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
      Left            =   4410
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   98
      Width           =   990
   End
   Begin VB.CommandButton cmdSearch1 
      Height          =   300
      Left            =   3465
      Picture         =   "Frmacc11l0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4125
      Width           =   350
   End
   Begin VB.TextBox txtA0N02 
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
      Left            =   2340
      TabIndex        =   16
      Top             =   4110
      Width           =   1125
   End
   Begin VB.TextBox txtCNo1 
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
      Index           =   4
      Left            =   1920
      MaxLength       =   2
      TabIndex        =   15
      Top             =   4110
      Width           =   372
   End
   Begin VB.TextBox txtCNo1 
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
      Index           =   3
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   14
      Top             =   4110
      Width           =   252
   End
   Begin VB.TextBox txtCNo1 
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
      Index           =   2
      Left            =   840
      MaxLength       =   6
      TabIndex        =   13
      Top             =   4110
      Width           =   852
   End
   Begin VB.TextBox txtCNo1 
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
      Index           =   1
      Left            =   360
      MaxLength       =   3
      TabIndex        =   12
      Top             =   4110
      Width           =   492
   End
   Begin VB.CommandButton cmdSearch 
      Height          =   300
      Left            =   2745
      Picture         =   "Frmacc11l0.frx":0102
      Style           =   1  '圖片外觀
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   495
      Width           =   350
   End
   Begin VB.TextBox txtSFee 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
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
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   900
      Width           =   1572
   End
   Begin VB.TextBox txtOFee 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
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
      Left            =   4410
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   900
      Width           =   1572
   End
   Begin VB.TextBox txtProperty 
      BackColor       =   &H00E0E0E0&
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
      Left            =   4410
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   495
      Width           =   2835
   End
   Begin VB.CommandButton cmdClear 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   5940
      Picture         =   "Frmacc11l0.frx":0204
      Style           =   1  '圖片外觀
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "清除畫面"
      Top             =   3855
      Width           =   555
   End
   Begin VB.CommandButton cmdCut 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   6525
      Picture         =   "Frmacc11l0.frx":0ACE
      Style           =   1  '圖片外觀
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "取消"
      Top             =   3855
      Width           =   555
   End
   Begin VB.TextBox txtA0N04 
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
      Left            =   4905
      MaxLength       =   14
      TabIndex        =   19
      Top             =   4110
      Width           =   960
   End
   Begin VB.TextBox txtA0N03 
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
      Left            =   3870
      MaxLength       =   14
      TabIndex        =   18
      Top             =   4110
      Width           =   960
   End
   Begin VB.TextBox txtOFeeTot 
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
      Left            =   4995
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3270
      Width           =   1212
   End
   Begin VB.TextBox txtSFeeTot 
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
      Left            =   2070
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3270
      Width           =   1212
   End
   Begin VB.TextBox txtCNo 
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
      Index           =   1
      Left            =   1170
      MaxLength       =   3
      TabIndex        =   0
      Top             =   98
      Width           =   492
   End
   Begin VB.TextBox txtCNo 
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
      Index           =   2
      Left            =   1650
      MaxLength       =   6
      TabIndex        =   1
      Top             =   98
      Width           =   852
   End
   Begin VB.TextBox txtCNo 
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
      Index           =   3
      Left            =   2490
      MaxLength       =   1
      TabIndex        =   2
      Top             =   98
      Width           =   252
   End
   Begin VB.TextBox txtCNo 
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
      Index           =   4
      Left            =   2730
      MaxLength       =   2
      TabIndex        =   3
      Top             =   98
      Width           =   372
   End
   Begin VB.TextBox txtA0N01 
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
      Left            =   1170
      TabIndex        =   4
      Top             =   495
      Width           =   1572
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   300
      Top             =   2250
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc11l0.frx":1138
      Height          =   1905
      Left            =   105
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1320
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   3360
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   16
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "C00"
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
      BeginProperty Column01 
         DataField       =   "A0N02"
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
      BeginProperty Column02 
         DataField       =   "C02"
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
         DataField       =   "A0N03"
         Caption         =   "服務費"
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
         DataField       =   "A0N04"
         Caption         =   "規費"
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
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1980.284
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1094.74
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收文日"
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
      Left            =   3465
      TabIndex        =   34
      Top             =   150
      Width           =   675
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
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
      Height          =   210
      Left            =   225
      TabIndex        =   33
      Top             =   3330
      Width           =   450
   End
   Begin VB.Label Label8 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "服務費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3870
      TabIndex        =   32
      Top             =   3840
      Width           =   630
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "服務費"
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
      Left            =   195
      TabIndex        =   31
      Top             =   945
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "規費"
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
      Left            =   3420
      TabIndex        =   30
      Top             =   945
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "案件性質"
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
      Left            =   3420
      TabIndex        =   29
      Top             =   540
      Width           =   900
   End
   Begin VB.Label Label6 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "收文號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2340
      TabIndex        =   28
      Top             =   3840
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   4305
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   915
      Left            =   180
      Top             =   3720
      Width           =   7080
   End
   Begin VB.Label Label16 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "規費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4905
      TabIndex        =   27
      Top             =   3840
      Width           =   420
   End
   Begin VB.Label Label13 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "本所案號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   26
      Top             =   3840
      Width           =   840
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "規費"
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
      Left            =   4320
      TabIndex        =   25
      Top             =   3330
      Width           =   450
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "服務費"
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
      Left            =   1305
      TabIndex        =   24
      Top             =   3330
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Height          =   210
      Left            =   195
      TabIndex        =   23
      Top             =   150
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收文號"
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
      Left            =   195
      TabIndex        =   22
      Top             =   540
      Width           =   675
   End
End
Attribute VB_Name = "Frmacc11l0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Create by Morgan 2011/4/6
Option Explicit

Dim m_bReadGrid As Boolean
Dim m_adoRst As ADODB.Recordset
Dim m_bolSel As Boolean
'Add by Morgan 2011/4/13
Public m_sCallType As String 'A:新增,E:修改
Public m_fCallForm As Form
Public m_bRefresh As Boolean
Public m_sAssignNo As String


Private Sub cmdClear_Click()
   ClearField
End Sub

Private Sub cmdSearch_Click()
   Dim bolUserClick As Boolean
   bolUserClick = IIf(Me.ActiveControl.Name = "cmdSearch", True, False)
   If txtCNo(3) = "" Then txtCNo(3) = "0"
   If txtCNo(4) = "" Then txtCNo(4) = "00"
   strExc(0) = GetRecNo(txtCNo(1), txtCNo(2), txtCNo(3), txtCNo(4), , bolUserClick)
   If strExc(0) <> "" Then
      txtA0N01 = strExc(0)
      txtProperty = ""
      txtOFee = ""
      txtSFee = ""
      txtDate = ""
      QueryData
      If bolUserClick And (strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4)) Then
         txtCNo1(1).SetFocus
      End If
   End If
End Sub

Private Sub cmdSearch1_Click()
   Dim bolUserClick As Boolean
   bolUserClick = IIf(Me.ActiveControl.Name = "cmdSearch1", True, False)
   If txtCNo1(3) = "" Then txtCNo1(3) = "0"
   If txtCNo1(4) = "" Then txtCNo1(4) = "00"
   strExc(0) = GetRecNo(txtCNo1(1), txtCNo1(2), txtCNo1(3), txtCNo1(4), True, bolUserClick)
   If strExc(0) <> "" Then
      txtA0N02 = strExc(0)
      If bolUserClick And (strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4)) Then
         txtA0N03.SetFocus
      End If
   End If
End Sub

Private Sub cmdCut_Click()
   If Not (m_adoRst.EOF Or m_adoRst.BOF) Then
      m_adoRst.Delete
      m_adoRst.UpdateBatch
      DataGrid1.Refresh
      ClearField
      ShowSum
   End If
End Sub

Public Function TxtValidate() As Boolean
   If txtA0N01 = "" Then
      MsgBox "收文號不可空白！"
      If txtA0N01.Enabled Then txtA0N01.SetFocus
      Exit Function
   Else
      'Modified by Morgan 2014/3/19 收費有可能為0,改提醒後可繼續
      'strExc(0) = "select cp09 from caseprogress where cp09='" & txtA0N01 & "' and cp16>0"
      strExc(0) = "select cp09,nvl(cp16,0) from caseprogress where cp09='" & txtA0N01 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp(1) = 0 Then
            'MsgBox "收文號輸入錯誤！"
            If MsgBox("收文號 " & txtA0N01 & " 費用為 0，是否仍要繼續！", vbYesNo + vbDefaultButton2 + vbQuestion, "費用檢查") = vbNo Then
               If txtA0N01.Enabled Then txtA0N01.SetFocus
               Exit Function
            End If
         End If
      End If
   End If
   
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      If Val(txtSFee) <> Val(txtSFeeTot) Then
         MsgBox "服務費不符！"
         Exit Function
      End If
      
      If Val(txtOFee) <> Val(txtOFeeTot) Then
         MsgBox "規費不符！"
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function

Public Function FormDelete() As Boolean
   If TxtValidate = False Then
      Exit Function
   End If
   
On Error GoTo ErrHnd

   strSql = "delete acc0n0 where a0n01='" & txtA0N01 & "'"
      adoTaie.Execute strSql, intI
   FormDelete = True
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function


Public Function FormSave() As Boolean
   
   If TxtValidate = False Then Exit Function
   
   adoTaie.BeginTrans
   
   If strSaveConfirm = MsgText(4) Then
      strSql = "delete acc0n0 where a0n01='" & txtA0N01 & "'"
      adoTaie.Execute strSql, intI
   End If
   
   With m_adoRst
   If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
         strSql = "insert into acc0n0(a0n01,a0n02,a0n03,a0n04,a0n05,a0n06,a0n07)" & _
            " values('" & txtA0N01 & "','" & .Fields("a0n02") & "'," & .Fields("a0n03") & _
            "," & .Fields("a0n04") & ",'" & strUserNum & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'))"
         adoTaie.Execute strSql, intI
         .MoveNext
      Loop
   End If
   End With
   
   adoTaie.CommitTrans
   FormSave = True
   Exit Function
   
ErrHnd:
   adoTaie.RollbackTrans
   MsgBox Err.Description
   
End Function

Private Sub Form_Activate()
   strFormName = Name
   '新增呼叫
   If m_sCallType = "A" Then
      tool1_enabled
      KeyEnter vbKeyF2
      txtA0N01 = m_sAssignNo
      QueryData
      If txtCNo1(1).Enabled Then txtCNo1(1).SetFocus
      
   '修改呼叫
   ElseIf m_sCallType = "E" Then
      tool1_enabled
      txtA0N01 = m_sAssignNo
      If QueryData = True Then
         KeyEnter vbKeyF3
         Frmacc0000.Toolbar1.Buttons.Item(7).Enabled = False '不可取消
      Else
         Unload Me
      End If
      
   ElseIf m_bRefresh Then
      m_bRefresh = False
      If strItemNo = MsgText(601) Then
         txtA0N01 = txtA0N01.Tag
      Else
         txtA0N01 = strItemNo
      End If
      QueryData
   End If
End Sub

Private Sub Form_Load()
   '表單初始化
   PUB_InitForm Me, Me.Width, Me.Height
   If m_sCallType = "" Then
      QueryData -2
   End If
   FormEnable
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   strSaveConfirm = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   If m_sCallType <> "" Then
      m_fCallForm.Enabled = True
      strFormName = m_fCallForm.Name
      m_fCallForm.Show
   End If
   Set Frmacc11l0 = Nothing
End Sub

Public Sub FormClear()
   Dim oControl As Control
   For Each oControl In Me.Controls
      If TypeName(oControl) = "TextBox" Then
        oControl.Text = Empty
      End If
   Next
   QueryData , False
End Sub

Public Sub FormEnable()
   Select Case strSaveConfirm
      Case MsgText(3) '新增
         txtCNo(1).Enabled = True
         txtCNo(2).Enabled = True
         txtCNo(3).Enabled = True
         txtCNo(4).Enabled = True
         txtA0N01.Enabled = True
         cmdSearch.Enabled = True
         
         txtCNo1(1).Enabled = True
         txtCNo1(2).Enabled = True
         txtCNo1(3).Enabled = True
         txtCNo1(4).Enabled = True
         txtA0N02.Enabled = True
         cmdSearch1.Enabled = True
         txtA0N03.Enabled = True
         txtA0N04.Enabled = True
         cmdClear.Enabled = True
         cmdCut.Enabled = True
         
      Case MsgText(4) '修改
         txtCNo(1).Enabled = False
         txtCNo(2).Enabled = False
         txtCNo(3).Enabled = False
         txtCNo(4).Enabled = False
         txtA0N01.Enabled = False
         cmdSearch.Enabled = False
         
         txtCNo1(1).Enabled = True
         txtCNo1(2).Enabled = True
         txtCNo1(3).Enabled = True
         txtCNo1(4).Enabled = True
         txtA0N02.Enabled = True
         cmdSearch1.Enabled = True
         txtA0N03.Enabled = True
         txtA0N04.Enabled = True
         cmdClear.Enabled = True
         cmdCut.Enabled = True
         
      Case Else
         txtCNo(1).Enabled = True
         txtCNo(2).Enabled = True
         txtCNo(3).Enabled = True
         txtCNo(4).Enabled = True
         txtA0N01.Enabled = True
         cmdSearch.Enabled = True
         
         txtCNo1(1).Enabled = False
         txtCNo1(2).Enabled = False
         txtCNo1(3).Enabled = False
         txtCNo1(4).Enabled = False
         txtA0N02.Enabled = False
         cmdSearch1.Enabled = False
         txtA0N03.Enabled = False
         txtA0N04.Enabled = False
         cmdClear.Enabled = False
         cmdCut.Enabled = False
   End Select
End Sub

Private Function GetRecNo(ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String _
   , Optional pbolNoMoney As Boolean = False, Optional ByVal pbolForceShow As Boolean = True) As String
   Dim stCon As String
   
   If pbolNoMoney Then stCon = " and (cp09='" & txtA0N01.Text & "' or (nvl(cp16,0)=0 and cp09<'B' and cp01='P' and cp10 in ('225','226') and cp57||cp20 is null))"
   
   If pCP03 = "" Then pCP03 = "0"
   If pCP04 = "" Then pCP04 = "00"
   
   strExc(0) = "select '',cp09,sqldatet(cp05) cp05,cpm03,st02,nvl(cp16,0)-nvl(cp17,0) 服務費,cp17 規費" & _
      " from caseprogress,casepropertymap,staff where cp01='" & pCP01 & "'" & _
      " and cp02='" & pCP02 & "' and cp03='" & pCP03 & "' and cp04='" & pCP04 & "'" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=cp14" & stCon
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.RecordCount = 1 And pbolForceShow = False Then
         GetRecNo = RsTemp("cp09")
      Else
         Set Frmacc21h4.grdDataList.Recordset = RsTemp
         Set Frmacc21h4.fmParent = Me
         Frmacc21h4.Show vbModal
         strFormName = Me.Name
         If Me.Tag <> "" Then
            GetRecNo = Me.Tag
         End If
      End If
      If GetRecNo <> "" Then
         m_bolSel = True
      End If
   Else
      MsgBox "無符合之收文資料！"
   End If
End Function

Private Sub txtA0N01_GotFocus()
   CloseIme
   If m_bolSel Then
      If txtCNo1(1).Enabled Then txtCNo1(1).SetFocus
   Else
      TextInverse txtA0N01
   End If
End Sub

Private Sub txtA0N01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub ReadMainData()
   '要考慮銷帳金額
   strExc(0) = "select cpm03,cp16-nvl(a1u07,0)-nvl(a1u09,0) cp16,cp17-nvl(a1u09,0) cp17" & _
      ",cp01,cp02,cp03,cp04,sqldatet(cp05) cp05" & _
      " from caseprogress,casepropertymap,(select a1u03,sum(a1u07) a1u07,sum(a1u09) a1u09" & _
      " from acc1u0 where a1u03='" & txtA0N01 & "' group by a1u03) X" & _
      " where cp09='" & txtA0N01 & "' and a1u03(+)=cp09" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      txtCNo(1) = "" & RsTemp("cp01")
      txtCNo(2) = "" & RsTemp("cp02")
      txtCNo(3) = "" & RsTemp("cp03")
      txtCNo(4) = "" & RsTemp("cp04")
      txtProperty = "" & RsTemp("cpm03")
      txtOFee = "" & RsTemp("cp17")
      If Not IsNull(RsTemp("cp16")) Then
         txtSFee = RsTemp("cp16") - Val(txtOFee)
      End If
      txtDate = "" & RsTemp("cp05")
   Else
      txtProperty = ""
      txtOFee = ""
      txtSFee = ""
      txtDate = ""
   End If
End Sub

Private Sub txtA0N01_Validate(Cancel As Boolean)
   If txtA0N01 = "" Then
      txtProperty = ""
      txtSFee = ""
      txtOFee = ""
   ElseIf txtA0N01.Tag <> txtA0N01.Text Then
      OpenTable
   End If
End Sub

Private Sub txtA0N02_GotFocus()
   CloseIme
   If m_bolSel Then
      SetFee
      txtA0N03.SetFocus
   Else
      TextInverse txtA0N02
   End If
End Sub

Private Sub SetFee()
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      If txtA0N03 = "" And txtA0N04 = "" Then
         txtA0N03 = Val(txtSFee) - Val(txtSFeeTot)
         txtA0N04 = Val(txtOFee) - Val(txtOFeeTot)
      End If
   End If
End Sub

Private Sub txtA0N02_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtA0N02_Validate(Cancel As Boolean)
   SetFee
End Sub

Private Sub txtA0N03_GotFocus()
   CloseIme
   m_bolSel = False
   TextInverse txtA0N03
End Sub

Private Sub txtA0N04_GotFocus()
   CloseIme
   TextInverse txtA0N04
End Sub

Private Sub txtCNo_GotFocus(Index As Integer)
   CloseIme
   TextInverse txtCNo(Index)
End Sub

Private Sub txtCNo_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCNo_Validate(Index As Integer, Cancel As Boolean)
   If strSaveConfirm <> MsgText(4) Then
      If Index = 4 Then
         cmdSearch.Value = True
      End If
   End If
End Sub

Private Sub txtCNo1_GotFocus(Index As Integer)
   CloseIme
   m_bolSel = False
   TextInverse txtCNo1(Index)
End Sub

Private Sub txtCNo1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub ReadData()
   With Adodc1.Recordset
      If Not (.EOF Or .BOF) Then
         txtCNo1(1) = "" & .Fields("CP01")
         txtCNo1(2) = "" & .Fields("CP02")
         txtCNo1(3) = "" & .Fields("CP03")
         txtCNo1(4) = "" & .Fields("CP04")
         txtA0N02 = "" & .Fields("a0n02")
         txtA0N03 = "" & .Fields("a0n03")
         txtA0N04 = "" & .Fields("a0n04")
         If txtCNo1(1).Enabled Then txtCNo1(1).SetFocus
      End If
   End With
End Sub

Private Sub DataGrid1_Click()
   '點選同一列可能不會觸發RowColChange
   If DataGrid1.col = -1 Then
      ReadData
   End If
   m_bReadGrid = True
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If m_bReadGrid = True Then
      ReadData
   End If
End Sub

Private Sub DataGrid1_Validate(Cancel As Boolean)
   m_bReadGrid = False
End Sub

Private Sub ClearField()
   txtCNo1(1) = ""
   txtCNo1(2) = ""
   txtCNo1(3) = ""
   txtCNo1(4) = ""
   txtA0N02 = ""
   txtA0N03 = ""
   txtA0N04 = ""
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyInsert
         UpdateRec
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Add1stRec()
   With m_adoRst
      .AddNew
      .Fields("C00") = txtCNo(1) & "-" & txtCNo(2) & IIf(txtCNo(3) & txtCNo(4) <> "000", "-" & txtCNo(3) & "-" & txtCNo(4), "")
      .Fields("a0n02") = txtA0N01
      .Fields("C02") = txtProperty
      .Fields("a0n03") = Val(txtSFee) - (Val(txtSFee) \ 3000) * 1000#
      .Fields("a0n04") = Val(txtOFee) - (Val(txtOFee) \ 3000) * 1000#
      .Fields("cp01") = txtCNo(1)
      .Fields("cp02") = txtCNo(2)
      .Fields("cp03") = txtCNo(3)
      .Fields("cp04") = txtCNo(4)
      .UPDATE
   End With
End Sub

Private Sub UpdateRec()
   
   
   strExc(0) = "select cpm03,cp16,cp01,cp02,cp03,cp04" & _
   " from caseprogress,casepropertymap where cp09='" & txtA0N02 & "'" & _
   " and cpm01(+)=cp01 and cpm02(+)=cp10"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If txtA0N02 <> txtA0N01 Then
         If Val("" & RsTemp("cp16")) > 0 Then
            MsgBox "分配的收文號不可有費用!!"
            Exit Sub
         End If
      End If
      txtCNo1(1) = RsTemp("cp01")
      txtCNo1(2) = RsTemp("cp02")
      txtCNo1(3) = RsTemp("cp03")
      txtCNo1(4) = RsTemp("cp04")
      strExc(1) = "" & RsTemp("cpm03")
      '目前只開放L案分配給P(因為收款作業也要同步)
      If txtCNo1(1) <> "L" And txtCNo1(1) <> "P" Then
         MsgBox "系統別輸入錯誤！(目前只開放 L 與 P)"
         Exit Sub
      End If
   Else
      MsgBox "收文號輸入錯誤!!"
      Exit Sub
   End If
   
   '2014/2/21 add by sonia 已收款提醒-瑞婷
   strExc(0) = "select nvl(cp79,0) cp79,nvl(cp16,0) cp16 from caseprogress where cp09='" & txtA0N01 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Val(RsTemp("cp16")) > 0 And Val(RsTemp("cp79")) = 0 Then
         MsgBox txtCNo(1) & txtCNo(2) & txtCNo(3) & txtCNo(4) & " 的 " & txtA0N01 & " 已收款!!"
      End If
   End If
   '2014/2/21 end
      
   With m_adoRst
   If .RecordCount > 0 Then
      m_bReadGrid = False
      .MoveFirst
      .Find "a0n02='" & txtA0N02 & "'"
      If .EOF Then
         .AddNew
      End If
   Else
      .AddNew
   End If
   .Fields("C00") = txtCNo1(1) & "-" & txtCNo1(2) & IIf(txtCNo1(3) & txtCNo1(4) <> "000", "-" & txtCNo1(3) & "-" & txtCNo1(4), "")
   .Fields("a0n02") = txtA0N02
   .Fields("C02") = strExc(1)
   .Fields("a0n03") = txtA0N03
   .Fields("a0n04") = txtA0N04
   .Fields("cp01") = txtCNo1(1)
   .Fields("cp02") = txtCNo1(2)
   .Fields("cp03") = txtCNo1(3)
   .Fields("cp04") = txtCNo1(4)
   .UPDATE
   End With
   txtA0N01.Tag = txtA0N01
   ClearField
   ShowSum
   txtCNo1(1).SetFocus
End Sub

Private Sub ShowSum()
   txtSFeeTot = ""
   txtOFeeTot = ""
   Set RsTemp = m_adoRst.Clone
   If RsTemp.RecordCount > 0 Then
      With RsTemp
      .MoveFirst
      Do While Not .EOF
         txtSFeeTot = Val(txtSFeeTot) + Val("" & .Fields("a0n03"))
         txtOFeeTot = Val(txtOFeeTot) + Val("" & .Fields("a0n04"))
         .MoveNext
      Loop
      End With
   End If
   Set RsTemp = Nothing
End Sub

Private Sub txtCNo1_Validate(Index As Integer, Cancel As Boolean)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      If Index = 4 Then
         cmdSearch1.Value = True
      End If
   End If
End Sub

'p_iDir: 0=本筆 1=下筆 2=末筆 -1=上筆 -2=首筆
Public Function QueryData(Optional p_iDir As Integer = 0, Optional bolShowErrMsg As Boolean = True) As Boolean
   Dim stCon As String
   
On Error GoTo Checking

   Select Case p_iDir
      Case 0
         stCon = " and a0n01='" & txtA0N01 & "'"
      Case 1
         stCon = " and a0n01=(select min(b.a0n01) from acc0n0 b where b.a0n01>'" & txtA0N01.Tag & "')"
      Case 2
         stCon = " and a0n01=(select max(b.a0n01) from acc0n0 b)"
      Case -1
         stCon = " and a0n01=(select max(b.a0n01) from acc0n0 b where b.a0n01<'" & txtA0N01.Tag & "')"
      Case -2
         stCon = " and a0n01=(select min(b.a0n01) from acc0n0 b)"
   End Select
   
   strExc(0) = "select a0n01,cp01,cp02,cp03,cp04 from acc0n0,caseprogress where cp09(+)=a0n01" & stCon
   intI = 1
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      txtA0N01 = RsTemp("a0n01")
      txtCNo(1) = RsTemp("cp01")
      txtCNo(2) = RsTemp("cp02")
      txtCNo(3) = RsTemp("cp03")
      txtCNo(4) = RsTemp("cp04")
      QueryData = True
      OpenTable
   '沒資料
   ElseIf intI = 0 Then
      OpenTable bolShowErrMsg
   End If
   
   
Checking:
   If Err.Number <> 0 Then
      MsgBox Err.Description, , MsgText(5)
   End If

End Function

Public Sub OpenTable(Optional bolShowErrMsg As Boolean = True)
   
   ReadMainData
   
   strExc(0) = "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) C00" & _
      ",a0n02,cpm03 C02,a0n03 ,a0n04 ,CP01,CP02,CP03,CP04" & _
      " From acc0n0 a, caseprogress, casepropertymap" & _
      " where a0n01='" & txtA0N01 & "' and cp09(+)=a0n02 and cpm01(+)=cp01" & _
      " and cpm02(+)=cp10" & _
      " order by 1,2"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'Modify by Amy 2014/06/30 +FormName 改暫存TB
   Set m_adoRst = PUB_CreateRecordset(RsTemp, , , , Me.Name)
   Set RsTemp = Nothing
   Set Adodc1.Recordset = m_adoRst
   Set DataGrid1.DataSource = Adodc1
   DataGrid1.Refresh
   DataGrid1.col = 0
   DataGrid1.CurrentCellVisible = True
   ClearField
   If Adodc1.Recordset.RecordCount > 0 Then
      txtA0N01.Tag = txtA0N01.Text
      If strSaveConfirm = MsgText(3) Then
         If bolShowErrMsg Then
            MsgBox "該收文號已有分配資料，作業模式已改為修改!!"
         End If
         strSaveConfirm = MsgText(4)
      End If
   Else
      If strSaveConfirm = MsgText(3) Then
         If txtA0N01 <> "" Then
            If txtCNo(1) <> "L" Then
               '目前只開放L案分配給P(因為收款作業也要同步)
               MsgBox "系統別輸入錯誤！(目前只開放 L)"
            Else
               Add1stRec
            End If
         End If
      Else
         If bolShowErrMsg Then
            MsgBox "分配資料不存在！"
         End If
      End If
   End If
   ShowSum
End Sub

Public Function MoveFirst(Optional bolShowErrMsg As Boolean = True) As Boolean
   MoveFirst = QueryData(-2, bolShowErrMsg)
End Function

Public Function MoveLast(Optional bolShowErrMsg As Boolean = True) As Boolean
   MoveLast = QueryData(2, bolShowErrMsg)
End Function

Public Function MoveNext(Optional bolShowErrMsg As Boolean = True) As Boolean
   MoveNext = QueryData(1, bolShowErrMsg)
End Function

Public Function MovePrevious(Optional bolShowErrMsg As Boolean = True) As Boolean
   MovePrevious = QueryData(-1, bolShowErrMsg)
End Function

Public Function EditCheck() As Boolean
   If txtA0N01.Tag = txtA0N01 And Adodc1.Recordset.RecordCount > 0 Then
      EditCheck = True
   Else
      MsgBox "資料不正確，請重新查詢！"
   End If
End Function
