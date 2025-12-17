VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc41a2 
   AutoRedraw      =   -1  'True
   Caption         =   "傳票資料輸入"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   8730
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5250
   ScaleWidth      =   8730
   Begin VB.TextBox Text10 
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
      Left            =   2370
      MaxLength       =   9
      TabIndex        =   6
      Top             =   4335
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
      Left            =   5400
      MaxLength       =   10
      TabIndex        =   7
      Top             =   4335
      Width           =   1572
   End
   Begin VB.TextBox Text9 
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
      Left            =   7155
      TabIndex        =   27
      Top             =   3600
      Width           =   1320
   End
   Begin VB.TextBox Text8 
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
      Left            =   6510
      MaxLength       =   5
      TabIndex        =   3
      Top             =   3600
      Width           =   600
   End
   Begin VB.TextBox Text1 
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
      Left            =   6840
      MaxLength       =   15
      TabIndex        =   25
      Top             =   240
      Width           =   1572
   End
   Begin VB.TextBox Text2 
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
      MaxLength       =   15
      TabIndex        =   15
      Top             =   240
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
      Left            =   3960
      TabIndex        =   14
      Top             =   2784
      Width           =   1215
   End
   Begin VB.TextBox Text4 
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
      Left            =   360
      MaxLength       =   6
      TabIndex        =   0
      Top             =   3600
      Width           =   1005
   End
   Begin VB.TextBox Text6 
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
      Left            =   3270
      MaxLength       =   14
      TabIndex        =   1
      Top             =   3600
      Width           =   1572
   End
   Begin VB.TextBox Text7 
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
      Left            =   2385
      MaxLength       =   12
      TabIndex        =   4
      Top             =   4005
      Width           =   1572
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   7080
      Picture         =   "Frmacc41a2.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   9
      ToolTipText     =   "取消"
      Top             =   2724
      Width           =   612
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
      Left            =   4890
      MaxLength       =   14
      TabIndex        =   2
      Top             =   3600
      Width           =   1572
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
      Left            =   5208
      TabIndex        =   12
      Top             =   2784
      Width           =   1215
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
      TabIndex        =   11
      Top             =   2784
      Width           =   855
   End
   Begin VB.TextBox Text13 
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
      Left            =   5400
      MaxLength       =   8
      TabIndex        =   5
      Top             =   4005
      Width           =   825
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc41a2.frx":066A
      Height          =   1980
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   8292
      _ExtentX        =   14623
      _ExtentY        =   3493
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
      ColumnCount     =   7
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
      BeginProperty Column04 
         DataField       =   "a1p16"
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
      BeginProperty Column05 
         DataField       =   "a1p17"
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
      BeginProperty Column06 
         DataField       =   "a1p06"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   3330.142
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1310.173
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1280.126
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   5559.875
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   760.252
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   4080
      TabIndex        =   16
      Top             =   240
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2134
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
   Begin MSForms.TextBox Text5 
      Height          =   315
      Left            =   1425
      TabIndex        =   13
      Top             =   3600
      Width           =   1755
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   330
      Left            =   1020
      TabIndex        =   8
      Top             =   4665
      Width           =   7425
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
   Begin MSForms.Label lblsaname 
      Height          =   285
      Left            =   6315
      TabIndex        =   33
      Top             =   4005
      Width           =   2115
      BackColor       =   14737632
      VariousPropertyBits=   19
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(本所案號)"
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
      TabIndex        =   32
      Top             =   4005
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(業)"
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
      Left            =   4050
      TabIndex        =   31
      Top             =   4005
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(客)"
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
      Left            =   405
      TabIndex        =   30
      Top             =   4335
      Width           =   1455
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(其)"
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
      Left            =   4050
      TabIndex        =   29
      Top             =   4335
      Width           =   1275
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
      Height          =   255
      Left            =   7185
      TabIndex        =   28
      Top             =   3330
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "結轉金額"
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
      TabIndex        =   26
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "入帳日期"
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
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label3 
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
      TabIndex        =   23
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label4 
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
      Left            =   3120
      TabIndex        =   22
      Top             =   2784
      Width           =   492
   End
   Begin VB.Label Label5 
      Alignment       =   2  '置中對齊
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
      Height          =   255
      Left            =   375
      TabIndex        =   21
      Top             =   3360
      Width           =   2460
   End
   Begin VB.Label Label6 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "借方金額"
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
      Left            =   3270
      TabIndex        =   20
      Top             =   3375
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1800
      Left            =   225
      Top             =   3270
      Width           =   8295
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4824
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label13 
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
      Height          =   255
      Left            =   390
      TabIndex        =   19
      Top             =   4740
      Width           =   615
   End
   Begin VB.Label Label14 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "貸方金額"
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
      Left            =   4890
      TabIndex        =   18
      Top             =   3360
      Width           =   1575
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
      TabIndex        =   17
      Top             =   2784
      Width           =   852
   End
End
Attribute VB_Name = "Frmacc41a2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/25 Form2.0已修改 lblsaname/Text5/Combo2
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoacc1p0 As New ADODB.Recordset
Public adocheck As New ADODB.Recordset
Dim strSerialNo As String
Dim IsLockData As Boolean

Private Sub Combo2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Command2_Click()
   AdodcDelete
   SumShow
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   strSerialNo = Adodc1.Recordset.Fields("a1p03").Value
   AdodcShow
End Sub

'Add by Amy 2021/10/25
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Form2.0 記錄鍵盤傳入順序
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   'add by nickc 2005/09/13
   If IsLockData = False Then
      KeyDefine KeyCode
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 8950
   Me.Height = 5810
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   Text2 = strItemNo
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = strCon1
   MaskEdBox1.Mask = DFormat
   Text1 = strCon2
   Text7 = strCon3
   Text13 = strCon4
'   Text8 = strCon5
   lblsaname = GetPrjSalesNM(Text13)
   OpenTable
   SumShow
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim adoquery As New ADODB.Recordset   'add by sonia 2017/6/19

   If Text3 <> Text12 Then
'避免失敗後產生錯誤 edit by nickc 2005/11/21
'      tool2_enabled
      MsgBox MsgText(11), , MsgText(5)
      Cancel = 1
      Exit Sub
   End If
   
   'add by sonia 2017/6/19 保留科目249X不可有小數位
   If adoquery.State = 1 Then adoquery.Close
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from acc1p0 where a1p01 = '" & Frmacc41a0.Text16.Tag & "' and a1p02 = 'S' and a1p04 = '" & Text2 & "' and substr(a1p05,1,3)='249' and (a1p07<>trunc(a1p07,0) or a1p08<>trunc(a1p08,0)) ", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      MsgBox "結餘保留科目不可有小數位！", , MsgText(5)
      Cancel = 1
      Exit Sub
   End If
   adoquery.Close
   'end 2017/6/19
   
   strItemNo = ""
   strCon1 = ""
   strCon2 = ""
   strCon3 = ""
   strCon4 = ""
   strCon5 = ""
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   'edit by nickc 2005/11/04
   'tool1_enabled
   tool2_enabled
   Frmacc41a0.Show
   strTrackMode = "" 'Add by Amy 2021/10/25 Form2.0 記錄鍵盤傳入順序(清除)
   Set Frmacc41a2 = Nothing
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox Label1 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label1 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   'Modify by Amy 2013/12/17 +公司別
   'adoadodc1.Open "select * from acc1p0, acc010 where a1p05 = a0101 and a1p01 = '1' and a1p02 = 'S' and a1p04 = '" & Text2 & "' order by a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc1.Open "select * from acc1p0, acc010 where a1p05 = a0101 and a1p01 = '" & Frmacc41a0.Text16.Tag & "' and a1p02 = 'S' and a1p04 = '" & Text2 & "' order by a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
   IsLockData = False
   'add by nickc 2005/09/13 檢查有無過帳資料
   If Not IsNull(adoadodc1.Fields("A1P22").Value) Then
      Dim strSQLc As String
      Dim Chk021Rs As New ADODB.Recordset
      Set Chk021Rs = New ADODB.Recordset
      'Modify by Amy 2013/12/17 +公司別
      'strSQLc = "select * from acc021 where ax201='1' and ax202='" & adoadodc1.Fields("A1P22").Value & "' "
      strSQLc = "select * from acc021 where ax201='" & Frmacc41a0.Text16.Tag & "' and ax202='" & adoadodc1.Fields("A1P22").Value & "' "
      Chk021Rs.CursorLocation = adUseClient
      Chk021Rs.Open strSQLc, adoTaie, adOpenStatic, adLockReadOnly
      If Chk021Rs.RecordCount <> 0 Then
         If Not IsNull(Chk021Rs.Fields("AX210").Value) Then
            '已過帳
            IsLockData = True
            Text4.Enabled = False
            Text6.Enabled = False
            Text11.Enabled = False
            Text13.Enabled = False
            Text7.Enabled = False
            Text8.Enabled = False
            Combo2.Enabled = False
         End If
      End If
   End If
   Set Chk021Rs = Nothing
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表(國內收款資料(分錄檔))
'
'*************************************************
Private Sub AdodcShow()
   If IsNull(Adodc1.Recordset.Fields("a1p05").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = Adodc1.Recordset.Fields("a1p05").Value
   End If
   'add by nickc 2005/11/08
   If IsNull(Adodc1.Recordset.Fields("a1p06").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = Adodc1.Recordset.Fields("a1p06").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p07").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = Adodc1.Recordset.Fields("a1p07").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p08").Value) Then
      Text11 = MsgText(601)
   Else
      Text11 = Adodc1.Recordset.Fields("a1p08").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p16").Value) = False Then
      Text13 = Adodc1.Recordset.Fields("a1p16").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p17").Value) = False Then
      Text7 = Adodc1.Recordset.Fields("a1p17").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p14").Value) Then
      Combo2 = MsgText(601)
   Else
      Combo2 = Adodc1.Recordset.Fields("a1p14").Value
   End If
   'add by nickc 2005/09/27
   If IsNull(Adodc1.Recordset.Fields("a1p30").Value) Then
      Text19 = MsgText(601)
   Else
      Text19 = Adodc1.Recordset.Fields("a1p30").Value
   End If
   'add by nickc 2005/09/29
   If IsNull(Adodc1.Recordset.Fields("a1p15").Value) Then
      Text10 = MsgText(601)
   Else
      Text10 = Adodc1.Recordset.Fields("a1p15").Value
   End If
   lblsaname = GetPrjSalesNM(Text13)
End Sub

'*************************************************
'  計算並顯示總計
'
'*************************************************
Public Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   'Modify by Amy 2013/12/17 +公司別
   'adoaccsum.Open "select sum(a1p07), sum(a1p08), count(*) from acc1p0 where a1p01 = '1' and a1p02 = 'S' and a1p04 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
   adoaccsum.Open "select sum(a1p07), sum(a1p08), count(*) from acc1p0 where a1p01 = '" & Frmacc41a0.Text16.Tag & "' and a1p02 = 'S' and a1p04 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text3 = MsgText(601)
      Else
         Text3 = adoaccsum.Fields(0).Value
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text12 = MsgText(601)
      Else
         Text12 = adoaccsum.Fields(1).Value
      End If
      If IsNull(adoaccsum.Fields(2).Value) Then
         Text20 = MsgText(601)
      Else
         Text20 = adoaccsum.Fields(2).Value
      End If
   Else
      Text3 = MsgText(601)
      Text12 = MsgText(601)
      Text20 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  儲存資料表(國內收款資料(分錄檔))
'
'*************************************************
Private Sub Acc1p0Save()
On Error GoTo Checking
   If Text4 = MsgText(601) Then
      MsgBox MsgText(10) & Label5, , MsgText(5)
      strControlButton = MsgText(602)
      Text4.SetFocus
      Exit Sub
   Else
      If ExistCheck("acc010", "a0101", Text4, Label5) = False Then
         strControlButton = MsgText(602)
         Text4.SetFocus
         Exit Sub
      End If
      If Text13 <> "" Then
         If ExistCheck("staff", "st01", Text13, Label12) = False Then
            MsgBox MsgText(10) & Label12, , MsgText(5)
            strControlButton = MsgText(602)
            Text13.SetFocus
            Exit Sub
         End If
      End If
      If Text7 <> MsgText(601) Then
         adocheck.CursorLocation = adUseClient
         adocheck.Open "select cp09 from caseprogress where cp01 = '" & Mid(Text7, 1, Len(Text7) - 9) & "' and cp02 = '" & Mid(Text7, Len(Text7) - 8, 6) & "' and cp03 = '" & Mid(Text7, Len(Text7) - 2, 1) & "' and cp04 = '" & Mid(Text7, Len(Text7) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adocheck.RecordCount = 0 Then
            MessageShow Label9
            strControlButton = MsgText(602)
            adocheck.Close
            Text7.SetFocus
            Exit Sub
         End If
         adocheck.Close
      End If
      If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
         If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
            MsgBox Label1 & MsgText(63), , MsgText(5)
            strControlButton = MsgText(602)
            Exit Sub
         End If
      End If
   End If
   
   'add by sonia 2015/12/30 檢查民國105年起法務收入科目不可使用
   intI = PUB_AccNoEnable(Text4, Val(FCDate(MaskEdBox1.Text)))
   If intI <> 0 Then
      strControlButton = MsgText(602)
      Text4.SetFocus
      Exit Sub
   End If
   'end 2015/12/30
   'Add by Morgan 2007/2/5 檢查科目部門&智權人員是否正確
   intI = PUB_AccNoGood(Text4, Text8, Text13)
   If intI <> 0 Then
      strControlButton = MsgText(602)
      If intI = 1 Then
         Text4.SetFocus
      ElseIf intI = 2 Then
         Text8.SetFocus
      ElseIf intI = 3 Then
         Text13.SetFocus
      End If
      Exit Sub
   End If
   'end 2007/2/5
      
   adoacc1p0.CursorLocation = adUseClient
   'Modify by Amy 2013/12/17 +公司別
   'adoacc1p0.Open "select * from acc1p0 where a1p01 = '1' and a1p02 = 'S' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text2 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc1p0.Open "select * from acc1p0 where a1p01 = '" & Frmacc41a0.Text16.Tag & "' and a1p02 = 'S' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text2 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc1p0.RecordCount = 0 Then
      adoacc1p0.AddNew
      adoacc1p0.Fields("a1p01").Value = Frmacc41a0.Text16.Tag 'Modify 2013/12/17 原: "1"
      adoacc1p0.Fields("a1p02").Value = "S"
      'Modify 2013/12/17
      'adoacc1p0.Fields("a1p03").Value = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'S' and a1p04 = '" & Text2 & "'", 3)
      adoacc1p0.Fields("a1p03").Value = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '" & Frmacc41a0.Text16.Tag & "' and a1p02 = 'S' and a1p04 = '" & Text2 & "'", 3)
      strSerialNo = adoacc1p0.Fields("a1p03").Value
      adoacc1p0.Fields("a1p04").Value = Text2
   End If
   'end 2013/12/17
   adoacc1p0.Fields("a1p05").Value = Text4
   If Text6 <> MsgText(601) Then
      adoacc1p0.Fields("a1p07").Value = Val(Text6)
   Else
      adoacc1p0.Fields("a1p07").Value = 0
   End If
   If Text11 <> MsgText(601) Then
      adoacc1p0.Fields("a1p08").Value = Val(Text11)
   Else
      adoacc1p0.Fields("a1p08").Value = 0
   End If
   If Text13 <> MsgText(601) Then
      adoacc1p0.Fields("a1p16").Value = Text13
   Else
      adoacc1p0.Fields("a1p16").Value = Null
   End If
   If Text7 <> MsgText(601) Then
      adoacc1p0.Fields("a1p17").Value = Text7
   Else
      adoacc1p0.Fields("a1p17").Value = Null
   End If
   If Text8 <> MsgText(601) Then
      adoacc1p0.Fields("a1p06").Value = Text8
   Else
      adoacc1p0.Fields("a1p06").Value = MsgText(55)
   End If
   If Combo2 <> MsgText(601) Then
      adoacc1p0.Fields("a1p14").Value = Combo2
      Combo2.AddItem Combo2
   Else
      adoacc1p0.Fields("a1p14").Value = Null
   End If
   If MaskEdBox1.Text <> MsgText(29) Then
      adoacc1p0.Fields("a1p18").Value = Val(FCDate(MaskEdBox1.Text))
   Else
      adoacc1p0.Fields("a1p18").Value = Null
   End If
   'add by nickc 2005/09/27
   If Text19 <> MsgText(601) Then
      adoacc1p0.Fields("a1p30").Value = Text19
   Else
      adoacc1p0.Fields("a1p30").Value = Null
   End If
   'add by nickc 2005/09/29
   If Text10 <> MsgText(601) Then
      adoacc1p0.Fields("a1p15").Value = Text10
   Else
      adoacc1p0.Fields("a1p15").Value = Null
   End If
   adoacc1p0.UpdateBatch
   adoacc1p0.Close
   'add by nickc 2005/09/13 有更新時，要上更新註記日期時間
   'edit by nickc 2005/11/08 還沒有傳票的不更新
   'adoTaie.Execute "update acc1p0 set a1p27='Y',a1p28=to_char(to_number(to_char(sysdate,'YYYY'))-1911)||to_char(sysdate, 'MMDD'),a1p29=to_number(to_char(sysdate,'hh24miss')) where a1p01 = '1' and a1p02 = 'S' and a1p04 = '" & Text2 & "' "
   'Modify by Amy 2013/12/17 +公司別
   'adoTaie.Execute "update acc1p0 set a1p27='Y',a1p28=to_char(to_number(to_char(sysdate,'YYYY'))-1911)||to_char(sysdate, 'MMDD'),a1p29=to_number(to_char(sysdate,'hh24miss')) where a1p01 = '1' and a1p02 = 'S' and a1p04 = '" & Text2 & "' and a1p22 is not null "
   adoTaie.Execute "update acc1p0 set a1p27='Y',a1p28=to_char(to_number(to_char(sysdate,'YYYY'))-1911)||to_char(sysdate, 'MMDD'),a1p29=to_number(to_char(sysdate,'hh24miss')) where a1p01 = '" & Frmacc41a0.Text16.Tag & "' and a1p02 = 'S' and a1p04 = '" & Text2 & "' and a1p22 is not null "
   AdodcRefresh
   Adodc1.Recordset.Find "a1p03 = '" & strSerialNo & "' ", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF Then
      Adodc1.Recordset.MoveFirst
   End If
   strSerialNo = MsgText(601)
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
   'Modify by Amy 2013/12/17 +公司別
   'adoadodc1.Open "select * from acc1p0, acc010, acc0g0 where a1p05 = a0101 and a1p10 = a0g01 (+) and a1p01 = '1' and a1p02 = 'S' and a1p04 = '" & Text2 & "' order by a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc1.Open "select * from acc1p0, acc010, acc0g0 where a1p05 = a0101 and a1p10 = a0g01 (+) and a1p01 = '" & Frmacc41a0.Text16.Tag & "' and a1p02 = 'S' and a1p04 = '" & Text2 & "' order by a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.Requery
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
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   'Modify by Amy 2013/12/17 +公司別
   'adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'S' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text2 & "'"
   adoTaie.Execute "delete from acc1p0 where a1p01 = '" & Frmacc41a0.Text16.Tag & "' and a1p02 = 'S' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & Text2 & "'"
   AdodcRefresh
   AdodcClear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   'Add by Amy 2021/10/25
   Call PUB_SaveTrackMode(1, KeyCode)
   'Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
   If PUB_ChkTrackMode = False Then
        Exit Sub
   End If
   'end 2021/10/25
   Select Case KeyCode
      Case vbKeyInsert
         'add by nickc 2007/03/06 若是沒有指定部門，則自動指定 tot
         If Trim(Text8) = "" Then Text8.Text = "TOT"
         
         'add by sonia 2015/4/22 41XX(除4191,4192,4194)外或7121摘要有結餘,對沖其他欄也要有
         If (Left(Text4, 2) = "41" And Text4 <> "4191" And Text4 <> "4192" And Text4 <> "4194") Or Text4 = "7121" Then
            If InStr(Combo2, "結餘") > 0 And InStr(Text19, "結餘") = 0 Then
               MsgBox "收文科目摘要欄內有 結餘 字樣, 對沖代號(其)欄也要輸結餘！", vbExclamation, "資料錯誤"
               TextInverse Text19
               Text19.SetFocus
               Exit Sub
            End If
         End If
         '2015/4/22 end
         
         'add by sonia 2020/9/8 2491或2211對沖其他欄一定要輸
         If Left(Text4, 4) = "2491" And InStr(Text19, "結餘") = 0 Then
            MsgBox "此科目一定要輸對沖代號(其)欄, 而且要有「結餘」字樣！", vbExclamation, "資料錯誤"
            TextInverse Text19
            Text19.SetFocus
            Exit Sub
         End If
         If Text4 = "2211" And Text19 = "" Then
            MsgBox "此科目一定要輸對沖代號(其)欄！", vbExclamation, "資料錯誤"
            TextInverse Text19
            Text19.SetFocus
            Exit Sub
         End If
         '2020/9/8 end
         'Add by Amy 2021/10/25 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
         If PUB_ChkUniText(Me) = False Then
            strControlButton = MsgText(602)
            Exit Sub
         End If
         'end 2021/10/25
         
         Acc1p0Save
         If strControlButton <> MsgText(602) Then 'Add by Morgan 2007/3/5 加錯誤判斷
            AdodcClear
            SumShow
            Text4.SetFocus
         End If
         'add by nickc 2008/03/12 取消其他功能，剩下  insert 和 esc
         KeyEnter KeyCode
      Case vbKeyEscape
         KeyEnter KeyCode
   End Select
   'edit by nickc 2008/03/12 取消所有的功能
   'KeyEnter KeyCode
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 <> MsgText(601) Then
      If Len(Text10) = 6 Then
         Text10 = AfterZero(Text10)
      'Add by Morgan 2007/3/1 八碼時要補'0'
      ElseIf Len(Text10) = 8 Then
         Text10 = Text10 & "0"
      'End 2007/3/1
      End If
      If ExistCheck("customer", "cu01", Mid(Text10, 1, 8), Label8, False) = False Then
         If ExistCheck("acc0i0", "a0i01", Text10, Label8, False) = False Then
            If ExistCheck("staff", "st01", Text10, Label8, False) = False Then
               If ExistCheck("fagent", "fa01", Mid(Text10, 1, 8), Label8, False) = False Then
                  MsgBox MsgText(28) & Label8, , MsgText(5)
                  Cancel = True
                  Exit Sub
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub

Private Sub Text13_GotFocus()
   TextInverse Text13
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text13_Validate(Cancel As Boolean)
   If Text13 <> MsgText(601) Then
      If ExistCheck("staff", "st01", Text13, Label12) = False Then
         MsgBox MsgText(45) & Label12, , MsgText(5)
         Cancel = True
         Exit Sub
      'add by nickc 2005/09/29
      Else
         lblsaname = GetPrjSalesNM(Text13)
      End If
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_Change()
   Text5 = A0102Query(Text4)
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
   If Text4 <> MsgText(601) Then
      If ExistCheck("acc010", "a0101", Text4, Label5) = False Then
         MsgBox MsgText(45) & Label5, , MsgText(5)
         Cancel = True
         Exit Sub
      End If
   Else
      Exit Sub
   End If
   RemarkShow
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  清除查詢顯示
'
'*************************************************
Public Sub AdodcClear()
   Text3 = ""
   Text12 = ""
   Text20 = ""
   Text4 = ""
   Text6 = ""
   Text11 = ""
   Combo2 = ""
   'add by nickc 2005/09/27
   Text19 = ""
   Text10 = ""
End Sub

'*************************************************
'  摘要顯示
'
'*************************************************
Public Sub RemarkShow()
   If Mid(Text4, 1, 4) = "1130" Then
      Combo2 = StaffQuery(Text7) & " / " & Text13
      Exit Sub
   End If
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
   If Text7 = MsgText(601) Then
      Exit Sub
   End If
   adocheck.CursorLocation = adUseClient
   adocheck.Open "select cp09 from caseprogress where cp01 = '" & Mid(Text7, 1, Len(Text7) - 9) & "' and cp02 = '" & Mid(Text7, Len(Text7) - 8, 6) & "' and cp03 = '" & Mid(Text7, Len(Text7) - 2, 1) & "' and cp04 = '" & Mid(Text7, Len(Text7) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount = 0 Then
      MessageShow Label9
      adocheck.Close
      Cancel = True
      Text7.SetFocus
      Exit Sub
   End If
   adocheck.Close
End Sub

Private Sub Text8_Change()
   Text9 = A0902Query(Text8)
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text8_Validate(Cancel As Boolean)
   If Text8 <> MsgText(601) Then
      If ExistCheck("acc090", "a0901", Text8, Label11) = False Then
         MsgBox MsgText(45) & Label11, , MsgText(5)
         Cancel = True
         Text8.SetFocus
         Exit Sub
      End If
   End If
End Sub
