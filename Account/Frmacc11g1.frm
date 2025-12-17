VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc11g1 
   AutoRedraw      =   -1  'True
   Caption         =   "沖轉分錄輸入"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   8740
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4920
   ScaleWidth      =   8740
   Begin VB.TextBox TextCompNo 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   4188
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   36
      Top             =   72
      Width           =   300
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5616
      MaxLength       =   3
      TabIndex        =   33
      Top             =   4392
      Width           =   528
   End
   Begin VB.CommandButton Command2 
      Caption         =   "沖暫收"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   7068
      TabIndex        =   30
      Top             =   48
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5028
      MaxLength       =   14
      TabIndex        =   4
      Top             =   3180
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   7668
      Picture         =   "Frmacc11g1.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   10
      ToolTipText     =   "取消"
      Top             =   2436
      Width           =   350
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3348
      TabIndex        =   16
      Top             =   2496
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   4188
      TabIndex        =   2
      Top             =   432
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2388
      TabIndex        =   6
      Top             =   3660
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1908
      TabIndex        =   14
      Top             =   3180
      Width           =   2775
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6828
      MaxLength       =   14
      TabIndex        =   5
      Top             =   3180
      Width           =   1572
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4860
      TabIndex        =   13
      Top             =   2496
      Width           =   1455
   End
   Begin VB.TextBox Text20 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1308
      MaxLength       =   12
      TabIndex        =   12
      Top             =   2496
      Width           =   855
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7068
      MaxLength       =   15
      TabIndex        =   11
      Top             =   432
      Width           =   1335
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   348
      TabIndex        =   3
      Top             =   3180
      Width           =   1575
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   1548
      TabIndex        =   0
      Top             =   72
      Width           =   1572
      _ExtentX        =   2769
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc11g1.frx":066A
      Height          =   1596
      Left            =   228
      TabIndex        =   15
      Top             =   792
      Width           =   8292
      _ExtentX        =   14623
      _ExtentY        =   2822
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.5
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "a0102"
         Caption         =   "會計科目"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
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
            ColumnWidth     =   2780.221
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   5449.89
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   228
      Top             =   672
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   2117
      _ExtentY        =   564
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1548
      TabIndex        =   1
      Top             =   432
      Width           =   1572
      _ExtentX        =   2769
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSForms.TextBox Text14 
      Height          =   330
      Left            =   2388
      TabIndex        =   31
      Top             =   4368
      Width           =   1572
      VariousPropertyBits=   671105051
      Size            =   "2773;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text4 
      Height          =   330
      Left            =   5625
      TabIndex        =   7
      Top             =   3660
      Width           =   1575
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "2778;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   330
      Left            =   2388
      TabIndex        =   8
      Top             =   4020
      Width           =   1572
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "2773;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   330
      Left            =   5628
      TabIndex        =   9
      Top             =   4020
      Width           =   2772
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "4890;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "作帳公司"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3240
      TabIndex        =   35
      Top             =   72
      Width           =   972
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "部門別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4200
      TabIndex        =   34
      Top             =   4404
      Width           =   972
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(其他)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   348
      TabIndex        =   32
      Top             =   4380
      Width           =   1644
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "欲處理日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   348
      TabIndex        =   29
      Top             =   432
      Width           =   1212
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "會計科目"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   348
      TabIndex        =   28
      Top             =   2940
      Width           =   4332
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "借方金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5028
      TabIndex        =   27
      Top             =   2940
      Width           =   1572
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1944
      Left            =   228
      Top             =   2832
      Width           =   8292
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2508
      TabIndex        =   26
      Top             =   2496
      Width           =   732
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   -12
      Top             =   4260
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "沖轉方式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3228
      TabIndex        =   25
      Top             =   432
      Width           =   972
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(客)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4188
      TabIndex        =   24
      Top             =   3660
      Width           =   1452
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(業)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   348
      TabIndex        =   23
      Top             =   4020
      Width           =   1452
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(本所案號)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   348
      TabIndex        =   22
      Top             =   3660
      Width           =   2052
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "摘要"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4188
      TabIndex        =   21
      Top             =   4020
      Width           =   492
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "沖轉日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   348
      TabIndex        =   20
      Top             =   72
      Width           =   1212
   End
   Begin VB.Label Label11 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "貸方金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6828
      TabIndex        =   19
      Top             =   2940
      Width           =   1572
   End
   Begin VB.Label Label22 
      BackStyle       =   0  '透明
      Caption         =   "筆數"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   348
      TabIndex        =   18
      Top             =   2496
      Width           =   852
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "暫收款單號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5868
      TabIndex        =   17
      Top             =   432
      Width           =   1212
   End
End
Attribute VB_Name = "Frmacc11g1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/30 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoaccsum As New ADODB.Recordset
Public adoadodc3 As New ADODB.Recordset
Public adoacc1p0 As New ADODB.Recordset
Public adocase As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public strCompNo As String    'add by sonia 2021/1/28 作帳公司別

Dim strSerialNo As String
Dim strAutoNo As String
Dim strOutputNo As String
Dim intNumber As Integer


Private Sub Combo1_Click()
   Combo3.Clear
   Select Case Mid(Combo1, 1, 1)
      Case "1"
         FormDisabled
         Combo3.AddItem "2401"
         Combo3 = "2401"
      Case "2"
         FormEnabled
         Combo3.AddItem "1101"
         Combo3.AddItem "1911"
         Combo3.AddItem "1912"
         Combo3.AddItem "1913"
         Combo3.AddItem "110202"
         Combo3.AddItem "113001"
      Case Else
         FormDisabled
   End Select
End Sub

Private Sub Combo2_GotFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
OpenIme
End Sub

Private Sub Combo2_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
CloseIme
End Sub

Private Sub Combo3_Validate(Cancel As Boolean)
   If Combo3 <> MsgText(601) Then
      If ExistCheck("acc010", "a0101", Combo3, Label2) = False Then
         MsgBox MsgText(28) & Label12, , MsgText(5)
         Cancel = True
         Exit Sub
      End If
   Else
     Combo3 = MsgText(601)
   End If
   Text7 = A0102Query(Combo3)
   Text4 = strCon2
   Text5 = strCon3
   If Combo2 = MsgText(601) Then
      Combo2 = strItemNo & "/" & Text11
   End If
   'modify by sonia 2021/1/28 加傳本所案號以判別FCP,FCT英日文組
   'If AccNoToSalesNo(Combo3) <> "" Then
   '   Text5 = AccNoToSalesNo(Combo3)
   If AccNoToSalesNo(Combo3, Text6) <> "" Then
      Text5 = AccNoToSalesNo(Combo3, Text6)
   'end 2021//28
   End If
End Sub

Private Sub Command1_Click()
   AdodcDelete
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Command2_Click()
   If Val(Text9) = 0 Then
      MsgBox MsgText(124), , MsgText(5)
      Exit Sub
   End If
   If Text3 <> Text9 Then
      MsgBox MsgText(11), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   OutputNo
   Screen.MousePointer = vbDefault
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   strSerialNo = Adodc1.Recordset.Fields("a1p03").Value
   AdodcShow
End Sub
'Added by Morgan 2023/12/1
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Call PUB_SaveTrackMode(0, KeyCode)
End Sub
'Added by Morgan 2023/12/1
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call PUB_SaveTrackMode(1, KeyCode)
   If PUB_ChkTrackMode = True Then
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
   Me.Width = 8970
   Me.Height = 5505
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
   Combo1.AddItem ComboItem(221)
   Combo1.AddItem ComboItem(222)
   Combo1 = ComboItem(222)
   Text11 = "J"
   Combo1_Click
   OpenTable
   SumShow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Text3 <> Text9 Then
      MsgBox MsgText(11), , MsgText(5)
      tool3_enabled
      Cancel = True
      Exit Sub
   End If
   If Mid(Combo1, 1, 1) = "1" Then
      If Text11 = MsgText(601) Then
         MsgBox MsgText(123), , MsgText(5)
         tool3_enabled
         Cancel = True
         Exit Sub
      End If
   End If
   'Add by Amy 2014/10/30 +FormCheck
   If FormCheck = False Then
      tool3_enabled
      Cancel = True
      Exit Sub
   End If
   
   'Add by Sindy 2021/11/30 檢查畫面上的物件是否含有Unicode文字
   If PUB_ChkUniText(Me, True, True) = False Then
      Cancel = True
      Exit Sub
   End If

   adoTaie.Execute "insert into acc1u0 select a1v02, a1v02, a1v01, 0, 0, a1v11 * (-1), 0, 0, 0, 0 from acc1v0 where a1v08 is not null and a1v11 <> 0 and a1v16 is null and a1v01 in " & strCon5
   'Added by Morgan 2011/11/9 更新進度檔已扣繳金額
   'Modified by Morgan 2012/4/19 修正 a1u01=cp09-->a1u03=cp09
   adoTaie.Execute "update caseprogress set cp76=(select nvl(sum(a1u06),0) from acc1u0 where a1u03=cp09) where cp09 in (select a1v01 from acc1v0 where a1v08 is not null and a1v11 <> 0 and a1v16 is null and a1v01 in " & strCon5 & " )", intI
   'end 2011/11/9
   adoTaie.Execute "update acc1v0 set a1v16 = '" & MsgText(602) & "', a1v08 = null, a1v07 = a1v11, a1v06 = a1v06 - a1v11 where a1v08 is not null and a1v11 <> 0 and a1v16 is null and a1v01 in " & strCon5
   'Modify by Amy 2014/10/30 +intNumber (未更新到正確的資料)
   'modify by sonia 2021/1/28 a1p01='1'改a1p01='" & TextCompNo & "'
   adoTaie.Execute "update acc1p0 set a1p12 = " & Val(FCDate(MaskEdBox1.Text)) & ", a1p18 = " & Val(FCDate(MaskEdBox2.Text)) & ", a1p23 = '" & Text11 & "', a1p24 = '" & Mid(Combo1, 1, 1) & "' where a1p01='" & TextCompNo & "' and a1p02 = 'W' and a1p04 = '" & strItemNo & intNumber & "'"
   adoTaie.Execute "update acc1p0 set a1p14 = a1p14||'" & Text11 & "' where a1p01='" & TextCompNo & "' and a1p02 = 'W' and a1p04 = '" & strItemNo & intNumber & "' and a1p05 = '2401'"
   'end 2021/1/28
   'end 2014/10/30
   lngTotal = 0
   tool3_enabled
   Select Case strFormLink
      Case "Frmacc11g0"
         Frmacc11g0.Enabled = True
   End Select
   Set Frmacc11g1 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   
   TextCompNo = strCompNo   'add by sonia 2021/1/28
   adoquery.CursorLocation = adUseServer
   'modify by sonia 2021/1/28 a1p01='1'改a1p01='" & TextCompNo & "'
   adoquery.Open "select max(a1p04) from acc1p0 where a1p01='" & TextCompNo & "' and a1p02 = 'W' and a1p04 like '" & strItemNo & "%" & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields(0).Value) Then
         intNumber = 1
      Else
         intNumber = Val(Mid(adoquery.Fields(0).Value, Len(strItemNo) + 1, Len(adoquery.Fields(0).Value) - Len(strItemNo))) + 1
      End If
   Else
      intNumber = 1
   End If
   MaskEdBox2.Enabled = True 'Add by Amy 2014/10/30
   adoquery.Close
   adoquery.CursorLocation = adUseClient
   'modify by sonia 2021/1/28 a1p01='1'改a1p01='" & TextCompNo & "'
   adoquery.Open "select * from acc1p0 where a1p01='" & TextCompNo & "' and a1p02 = 'W' and a1p04 = '" & strItemNo & intNumber & "' order by a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount = 0 Then
      'modify by sonia 2021/1/28 a1p01='1'改a1p01='" & TextCompNo & "'
      strAutoNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01='" & TextCompNo & "' and a1p02 = 'W' and a1p04 = '" & strItemNo & intNumber & "'", 3)
      'Modify By Sindy 2015/11/11 + a1p30=strCon7
      'modify by sonia 2021/1/28 a1p01固定'1'改TextCompNo
      adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13" & _
                      ", a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27, a1p30)" & _
                      " values ('" & TextCompNo & "', 'W', '" & strAutoNo & "', '" & strItemNo & intNumber & "', '1203', '" & MsgText(55) & "', 0, " & Val(strCon1) & ", null, null, null, null, null" & _
                      ", '" & StaffQuery(strCon3) & "/" & strItemNo & "', '" & strCon2 & "', '" & strCon3 & "', null, null, null, null, null, null, " & _
                      "null, null, null, null, null,'" & strCon7 & "')"
      MaskEdBox2.Mask = MsgText(601)
      MaskEdBox2.Text = CFDate(ACDate(ServerDate))
      MaskEdBox2.Mask = DFormat
      MaskEdBox1.Mask = MsgText(601)
      MaskEdBox1.Text = CFDate(ACDate(ServerDate))
      MaskEdBox1.Mask = DFormat
   Else
      MaskEdBox2.Mask = MsgText(601)
      If IsNull(adoquery.Fields("a1p18").Value) Then
         MaskEdBox2.Text = MsgText(601)
      Else
         MaskEdBox2.Text = CFDate(adoquery.Fields("a1p18").Value)
      End If
      MaskEdBox2.Mask = DFormat
      'Add by Amy 2014/10/30 a1p22有值不可修改
      If Not IsNull(adoquery.Fields("a1p22").Value) Then
         MaskEdBox2.Enabled = False
      End If
      MaskEdBox1.Mask = MsgText(601)
      If IsNull(adoquery.Fields("a1p12").Value) Then
         MaskEdBox1.Text = MsgText(601)
      Else
         MaskEdBox1.Text = CFDate(adoquery.Fields("a1p12").Value)
      End If
      MaskEdBox1.Mask = DFormat
      If IsNull(adoquery.Fields("a1p24").Value) Then
         Combo1 = MsgText(601)
      Else
         Combo1 = Combo1.List(Val(adoquery.Fields("a1p24").Value) - 1)
         Select Case adoquery.Fields("a1p24").Value
            Case "1"
               FormDisabled
            Case "2"
               FormEnabled
         End Select
      End If
      If IsNull(adoquery.Fields("a1p23").Value) Then
         Text11 = MsgText(601)
      Else
         Text11 = adoquery.Fields("a1p23").Value
         Command2.Enabled = False
      End If
   End If
   adoquery.Close
   adoadodc3.CursorLocation = adUseClient
   'modify by sonia 2021/1/28 a1p01='1'改a1p01='" & TextCompNo & "'
   adoadodc3.Open "select * from acc1p0, acc010 where a1p05 = a0101 and a1p01='" & TextCompNo & "' and a1p02 = 'W' and a1p04 = '" & strItemNo & intNumber & "' order by a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc3
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox MsgText(10) & Label1, , MsgText(5)
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

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
   Dim strMsg As String
    
   'Modify by Amy 2014/10/29 設必填 +系統日比較
   If MaskEdBox2.Enabled = False Then Exit Sub
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
   If ChkWorkData("1", DBDATE(MaskEdBox2), strMsg) = False Then
        MsgBox Label10 & strMsg, , MsgText(5)
        Cancel = True
        MaskEdBox2.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

'*************************************************
'  儲存資料表(扣繳稅款退費資料(分錄檔))
'
'*************************************************
Private Sub Acc1p0Save()
On Error GoTo Checking
   If Combo3 = MsgText(601) Then
      MsgBox MsgText(10), , MsgText(5)
      strControlButton = MsgText(602)
      Combo3.SetFocus
      Exit Sub
   Else
      If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
         MsgBox MsgText(10) & Label1, , MsgText(5)
         strControlButton = MsgText(602)
         MaskEdBox1.SetFocus
         Exit Sub
      Else
         If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
            MsgBox Label1 & MsgText(63), , MsgText(5)
            strControlButton = MsgText(602)
            MaskEdBox1.SetFocus
            Exit Sub
         End If
      End If
      If ExistCheck("acc010", "a0101", Combo3, Label2) = False Then
         MsgBox MsgText(28) & Label2, , MsgText(5)
         strControlButton = MsgText(602)
         Combo3.SetFocus
         Exit Sub
      End If
   End If
   
   'add by sonia 2015/12/30 檢查民國105年起法務收入科目不可使用
   intI = PUB_AccNoEnable(Combo3, Val(FCDate(MaskEdBox2.Text)))
   If intI <> 0 Then
      strControlButton = MsgText(602)
      Combo3.SetFocus
      Exit Sub
   End If
   'end 2015/12/30
   'Add by Morgan 2007/10/2 檢查科目部門&智權人員是否正確
   intI = PUB_AccNoGood(Combo3, Text2, Text5)
   If intI <> 0 Then
      strControlButton = MsgText(602)
      If intI = 1 Then
         Combo3.SetFocus
      ElseIf intI = 2 Then
         Text2.SetFocus
      ElseIf intI = 3 Then
         Text5.SetFocus
      End If
      Exit Sub
   End If
   'end 2007/10/2
   
   adoacc1p0.CursorLocation = adUseClient
   'modify by sonia 2021/1/28 a1p01='1'改a1p01='" & TextCompNo & "'
   adoacc1p0.Open "select * from acc1p0 where a1p01='" & TextCompNo & "' and a1p02 = 'W' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & strItemNo & intNumber & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc1p0.RecordCount = 0 Then
      adoacc1p0.AddNew
      'modify by sonia 2021/1/28 a1p01='1'改a1p01='" & TextCompNo & "'
      adoacc1p0.Fields("a1p01").Value = TextCompNo
      adoacc1p0.Fields("a1p02").Value = "W"
      'modify by sonia 2021/1/28 a1p01='1'改a1p01='" & TextCompNo & "'
      adoacc1p0.Fields("a1p03").Value = GetSerialNo("select max(a1p03) from acc1p0 where a1p01='" & TextCompNo & "' and a1p02 = 'W' and a1p04 = '" & strItemNo & intNumber & "'", 3)
      adoacc1p0.Fields("a1p04").Value = strItemNo & intNumber
   End If
   adoacc1p0.Fields("a1p05").Value = Combo3
   adoacc1p0.Fields("a1p06").Value = MsgText(55)
   If Text1 <> MsgText(601) Then
      adoacc1p0.Fields("a1p07").Value = Val(Text1)
   Else
      adoacc1p0.Fields("a1p07").Value = 0
   End If
   If Text8 <> MsgText(601) Then
      adoacc1p0.Fields("a1p08").Value = Val(Text8)
   Else
      adoacc1p0.Fields("a1p08").Value = 0
   End If
   If Text6 <> MsgText(601) Then
      adoacc1p0.Fields("a1p17").Value = Text6
   Else
      adoacc1p0.Fields("a1p17").Value = Null
   End If
   If Text4 <> MsgText(601) Then
      adoacc1p0.Fields("a1p15").Value = Text4
   Else
      adoacc1p0.Fields("a1p15").Value = Null
   End If
   If Text5 <> MsgText(601) Then
      adoacc1p0.Fields("a1p16").Value = Text5
   Else
      adoacc1p0.Fields("a1p16").Value = Null
   End If
   If Combo2 <> MsgText(601) Then
      adoacc1p0.Fields("a1p14").Value = Combo2
      Combo2.AddItem Combo2
   Else
      adoacc1p0.Fields("a1p14").Value = Null
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      adoacc1p0.Fields("a1p18").Value = Val(FCDate(MaskEdBox2.Text))
   Else
      adoacc1p0.Fields("a1p18").Value = Null
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      adoacc1p0.Fields("a1p12").Value = Val(FCDate(MaskEdBox1.Text))
   Else
      adoacc1p0.Fields("a1p12").Value = Null
   End If
   If Text11 <> MsgText(601) Then
      adoacc1p0.Fields("a1p23").Value = Text11
   Else
      adoacc1p0.Fields("a1p23").Value = Null
   End If
   If Combo1 <> MsgText(601) Then
      adoacc1p0.Fields("a1p24").Value = Mid(Combo1, 1, 1)
   Else
      adoacc1p0.Fields("a1p24").Value = Null
   End If
   If Text2 <> MsgText(601) Then
      adoacc1p0.Fields("a1p06").Value = Text2
   Else
      adoacc1p0.Fields("a1p06").Value = MsgText(55)
   End If
   If Text14 <> MsgText(601) Then
      adoacc1p0.Fields("a1p30").Value = Text14
   Else
      adoacc1p0.Fields("a1p30").Value = Null
   End If
   adoacc1p0.UpdateBatch
   adoacc1p0.Close
   'modify by sonia 2021/1/28 a1p01='1'改a1p01='" & TextCompNo & "'
   adoTaie.Execute "update acc1p0 set a1p18 = " & Val(FCDate(MaskEdBox2.Text)) & " where a1p01 = '" & TextCompNo & "' and a1p02 = 'W' and a1p04 = '" & strItemNo & intNumber & "'"
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
Private Sub AdodcRefresh()
On Error GoTo Checking
   adoadodc3.Close
   adoadodc3.CursorLocation = adUseClient
   'modify by sonia 2021/1/28 a1p01='1'改a1p01='" & TextCompNo & "'
   adoadodc3.Open "select * from acc1p0, acc010 where a1p05 = a0101 and a1p01='" & TextCompNo & "' and a1p02 = 'W' and a1p04 = '" & strItemNo & intNumber & "' order by a1p03 asc", adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.Requery
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
   Combo3 = Adodc1.Recordset.Fields("a1p05").Value
   If IsNull(Adodc1.Recordset.Fields("a1p07").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = Adodc1.Recordset.Fields("a1p07").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p08").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = Adodc1.Recordset.Fields("a1p08").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p17").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = Adodc1.Recordset.Fields("a1p17").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p15").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = Adodc1.Recordset.Fields("a1p15").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p16").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = Adodc1.Recordset.Fields("a1p16").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p14").Value) Then
      Combo2 = MsgText(601)
   Else
      Combo2 = Adodc1.Recordset.Fields("a1p14").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p06").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = Adodc1.Recordset.Fields("a1p06").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1p30").Value) Then
      Text14 = MsgText(601)
   Else
      Text14 = Adodc1.Recordset.Fields("a1p30").Value
   End If
End Sub

'*************************************************
'  刪除資料表(扣繳稅款退費資料(分錄檔))
'
'*************************************************
Private Sub AdodcDelete()
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   'modify by sonia 2021/1/28 a1p01='1'改a1p01='" & TextCompNo & "'
   adoTaie.Execute "delete from acc1p0 where a1p01='" & TextCompNo & "' and a1p02 = 'W' and a1p03 = '" & strSerialNo & "' and a1p04 = '" & strItemNo & intNumber & "'"
   SumShow
   AdodcRefresh
   AdodcClear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  清除顯示資料
'
'*************************************************
Private Sub AdodcClear()
   Combo3 = ""
   Text7 = ""
   Text1 = ""
   Text8 = ""
   Text6 = ""
   Text4 = ""
   Text5 = ""
   Combo2 = ""
   Text2 = ""
   Text14 = ""
End Sub

'*************************************************
'  計算並顯示合計
'
'*************************************************
Private Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   'modify by sonia 2021/1/28 a1p01='1'改a1p01='" & TextCompNo & "'
   adoaccsum.Open "select sum(a1p07), sum(a1p08), count(*) from acc1p0 where a1p01='" & TextCompNo & "' and a1p02 = 'W' and a1p04 = '" & strItemNo & intNumber & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text3 = MsgText(601)
      Else
         Text3 = Format(adoaccsum.Fields(0).Value, DAmount)
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text9 = MsgText(601)
      Else
         Text9 = Format(adoaccsum.Fields(1).Value, DAmount)
      End If
      If IsNull(adoaccsum.Fields(2).Value) Then
         Text20 = MsgText(601)
      Else
         Text20 = Format(adoaccsum.Fields(2).Value, DAmount)
      End If
   Else
      Text3 = MsgText(601)
      Text9 = MsgText(601)
      Text20 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyInsert
         Acc1p0Save
         AdodcRefresh
         SumShow
         AdodcClear
         Combo3.SetFocus
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
Dim bolValue As Boolean
   
   If Text11 = MsgText(601) Then
      Exit Sub
   End If
   'Modified by Lydia 2024/11/28 未收文客戶暫收款管制
   'bolValue = CheckData("select * from acc0t0 where a0t01 = '" & Text11 & "'", Label13)
   'If bolValue = False Then
   '   Cancel = True
   '   Text11.SetFocus
   'End If
   If adoquery.State = adStateOpen Then adoquery.Close
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select a.*,nvl(cu04,nvl(cu05,cu06)) cuname from acc0t0 a, customer where a0t01 = '" & Text11 & "' and substr(a0t06,1,8)=cu01(+) and substr(a0t06,9,1)=cu02(+) ", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If "" & adoquery.Fields("a0t06") = "X03072010" Then
         MsgBox "此暫收款之客戶為" & adoquery.Fields("A0t06") & adoquery.Fields("cuname") & "，不可沖帳 !", vbExclamation
         adoquery.Close
         Cancel = True
         Text11.SetFocus
         Exit Sub
      End If
   Else
      Cancel = True
      Text11.SetFocus
   End If
   adoquery.Close
   'end 2024/11/28
   
End Sub

Private Sub Text14_GotFocus()
   TextInverse Text14
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text14_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'Add by Morgan 2007/3/1
Private Sub Text4_Validate(Cancel As Boolean)
   If Text4 <> MsgText(601) Then
      If Len(Text4) = 6 Then
         Text4 = AfterZero(Text4)
      ElseIf Len(Text4) = 8 Then
         Text4 = Text4 & "0"
      End If
      If ExistCheck("customer", "cu01", Mid(Text4, 1, 8), Label6, False) = False Then
         If ExistCheck("acc0i0", "a0i01", Text4, Label6, False) = False Then
            If ExistCheck("staff", "st01", Text4, Label6, False) = False Then
               MsgBox MsgText(28) & Label6, , MsgText(5)
               Cancel = True
               Exit Sub
            End If
         End If
      End If
   End If
End Sub
'End 2007/3/1
Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'add by sonia 2021/1/28
Private Sub Text5_Validate(Cancel As Boolean)
   If Text5 <> MsgText(601) Then
      If PUB_GetStaffState(Text5.Text, strExc(1), True) = 0 Then
         Cancel = True
         TextInverse Text5
      End If
      If SalesNoCheckAccNo(Combo3, Text5) = False Then
      End If
   End If
End Sub
'end 2021/1/28

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   'Modify By Sindy 2021/11/30 控管案號長度至少為10,不然下列程式會出錯
   If Text6 <> MsgText(601) And Len(Trim(Text6)) >= 10 Then
      adocase.CursorLocation = adUseClient
      adocase.Open "select pa01 as SystemNo from patent where pa01 = '" & Mid(Text6, 1, Len(Text6) - 9) & "' and pa02 = '" & Mid(Text6, Len(Text6) - 8, 6) & "' and pa03 = '" & Mid(Text6, Len(Text6) - 2, 1) & "' and pa04 = '" & Mid(Text6, Len(Text6) - 1, 2) & "' union " & _
                   "select tm01 as SystemNo from trademark where tm01 = '" & Mid(Text6, 1, Len(Text6) - 9) & "' and tm02 = '" & Mid(Text6, Len(Text6) - 8, 6) & "' and tm03 = '" & Mid(Text6, Len(Text6) - 2, 1) & "' and tm04 = '" & Mid(Text6, Len(Text6) - 1, 2) & "' union " & _
                   "select lc01 as SystemNo from lawcase where lc01 = '" & Mid(Text6, 1, Len(Text6) - 9) & "' and lc02 = '" & Mid(Text6, Len(Text6) - 8, 6) & "' and lc03 = '" & Mid(Text6, Len(Text6) - 2, 1) & "' and lc04 = '" & Mid(Text6, Len(Text6) - 1, 2) & "' union " & _
                   "select hc01 as SystemNo from hirecase where hc01 = '" & Mid(Text6, 1, Len(Text6) - 9) & "' and hc02 = '" & Mid(Text6, Len(Text6) - 8, 6) & "' and hc03 = '" & Mid(Text6, Len(Text6) - 2, 1) & "' and hc04 = '" & Mid(Text6, Len(Text6) - 1, 2) & "' union " & _
                   "select sp01 as SystemNo from servicepractice where sp01 = '" & Mid(Text6, 1, Len(Text6) - 9) & "' and sp02 = '" & Mid(Text6, Len(Text6) - 8, 6) & "' and sp03 = '" & Mid(Text6, Len(Text6) - 2, 1) & "' and sp04 = '" & Mid(Text6, Len(Text6) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocase.RecordCount = 0 Then
         MsgBox MsgText(28) & Label8, , MsgText(5)
         Cancel = True
         adocase.Close
         Exit Sub
      End If
      adocase.Close
      'add by sonia 2021/1/28 以本所案號以判別FCP,FCT英日文組
      If AccNoToSalesNo(Combo3, Text6) <> "" Then
         Text5 = AccNoToSalesNo(Combo3, Text6)
      End If
      'end 2021/1/28
   End If
   QueryCustomer
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub

'*************************************************
'  關閉分錄欄位輸入狀態
'
'*************************************************
Public Sub FormDisabled()
   Text4.Enabled = False
   Text5.Enabled = False
   Text6.Enabled = False
End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'
'*************************************************
Public Sub FormEnabled()
   Text4.Enabled = True
   Text5.Enabled = True
   Text6.Enabled = True
End Sub

'*************************************************
'  以本所案號查詢客戶名稱
'
'*************************************************
Public Sub QueryCustomer()
Dim strSql As String

   If Text6 = MsgText(601) Then
      Exit Sub
   End If
   strSql = "select cu01||cu02 as Name, cu04, cu05, cu06, cu88, cu89, cu90 from patent, customer where substr(pa26, 1, 8) = cu01 and nvl(substr(pa26, 9, 1), '0') = cu02 and pa01 = '" & Mid(Text6, 1, Len(Text6) - 9) & "' and pa02 = '" & Mid(Text6, Len(Text6) - 8, 6) & "' and pa03 = '" & Mid(Text6, Len(Text6) - 2, 1) & "' and pa04 = '" & Mid(Text6, Len(Text6) - 1, 2) & "' union " & _
            "select cu01||cu02 as Name, cu04, cu05, cu06, cu88, cu89, cu90 from trademark, customer where substr(tm23, 1, 8) = cu01 and nvl(substr(tm23, 9, 1), '0') = cu02 and tm01 = '" & Mid(Text6, 1, Len(Text6) - 9) & "' and tm02 = '" & Mid(Text6, Len(Text6) - 8, 6) & "' and tm03 = '" & Mid(Text6, Len(Text6) - 2, 1) & "' and tm04 = '" & Mid(Text6, Len(Text6) - 1, 2) & "' union " & _
            "select cu01||cu02 as Name, cu04, cu05, cu06, cu88, cu89, cu90 from lawcase, customer where substr(lc11, 1, 8) = cu01 and nvl(substr(lc11, 9, 1), '0') = cu02 and lc01 = '" & Mid(Text6, 1, Len(Text6) - 9) & "' and lc02 = '" & Mid(Text6, Len(Text6) - 8, 6) & "' and lc03 = '" & Mid(Text6, Len(Text6) - 2, 1) & "' and lc04 = '" & Mid(Text6, Len(Text6) - 1, 2) & "' union " & _
            "select cu01||cu02 as Name, cu04, cu05, cu06, cu88, cu89, cu90 from hirecase, customer where substr(hc05, 1, 8) = cu01 and nvl(substr(hc05, 9, 1), '0') = cu02 and hc01 = '" & Mid(Text6, 1, Len(Text6) - 9) & "' and hc02 = '" & Mid(Text6, Len(Text6) - 8, 6) & "' and hc03 = '" & Mid(Text6, Len(Text6) - 2, 1) & "' and hc04 = '" & Mid(Text6, Len(Text6) - 1, 2) & "' union " & _
            "select cu01||cu02 as Name, cu04, cu05, cu06, cu88, cu89, cu90 from servicepractice, customer where substr(sp08, 1, 8) = cu01 and nvl(substr(sp08, 9, 1), '0') = cu02 and sp01 = '" & Mid(Text6, 1, Len(Text6) - 9) & "' and sp02 = '" & Mid(Text6, Len(Text6) - 8, 6) & "' and sp03 = '" & Mid(Text6, Len(Text6) - 2, 1) & "' and sp08 = '" & Mid(Text6, Len(Text6) - 1, 2) & "'"
   adocase.CursorLocation = adUseClient
   adocase.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adocase.RecordCount <> 0 Then
      If IsNull(adocase.Fields(0).Value) Then
         Text4 = MsgText(601)
      Else
         Text4 = adocase.Fields(0).Value
      End If
      If IsNull(adocase.Fields("cu04").Value) Then
         If IsNull(adocase.Fields("cu05").Value) Then
            If IsNull(adocase.Fields("cu06").Value) Then
               Combo2 = MsgText(601)
            Else
               Combo2 = adocase.Fields("cu06").Value
            End If
         Else
            Combo2 = adocase.Fields("cu05").Value
            If IsNull(adocase.Fields("cu88").Value) = False Then
               Combo2 = Combo2 & adocase.Fields("cu88").Value
            End If
            If IsNull(adocase.Fields("cu89").Value) = False Then
               Combo2 = Combo2 & adocase.Fields("cu89").Value
            End If
            If IsNull(adocase.Fields("cu90").Value) = False Then
               Combo2 = Combo2 & adocase.Fields("cu90").Value
            End If
         End If
      Else
         Combo2 = adocase.Fields("cu04").Value
      End If
   Else
      Text4 = MsgText(601)
      Combo2 = MsgText(601)
   End If
   adocase.Close
End Sub

'*************************************************
'  產生轉出應付或暫收款單號
'
'*************************************************
Public Sub OutputNo()
'   Select Case Mid(Combo1, 1, 1)
'      Case "1"
'         strOutputNo = AutoNo(MsgText(806), 5)
'         adoTaie.Execute "insert into acc0t0 values ('" & strOutputNo & "', '3', " & Val(FCDate(MaskEdBox2.Text)) & ", " & Val(FCDate(MaskEdBox1.Text)) & ", null, '" & strCon2 & "', null, " & Val(Text9) & ", null, null, null, '" & strUserNum & "', " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", null, null, null)"
'   End Select
'   Text11 = strOutputNo
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   'modify by sonia 2021/1/28 加公司別檢查
   'adoquery.Open "select * from acc0t0 where a0t01 = '" & Text11 & "'", adoTaie, adOpenStatic, adLockReadOnly
   adoquery.Open "select * from acc0t0 where a0t01 = '" & Text11 & "' and a0t18 = '" & TextCompNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      'modify by sonia 2021/1/28 a1p01='1'改a1p01='" & TextCompNo & "'
      'modify by sonia 2025/2/26 同時更新a1p30以利傳送到AX213
      adoTaie.Execute "update acc1p0 set a1p23 = '" & Text11 & "',a1p30 = '" & Text11 & "' where a1p01='" & TextCompNo & "' and a1p02 = 'W' and a1p04 = '" & strItemNo & intNumber & "'"
   Else
      MsgBox MsgText(28) & Label13, , MsgText(5)
      adoquery.Close
      Exit Sub
   End If
   adoquery.Close
   If Text11 <> MsgText(601) Then
      Command2.Enabled = False
   End If
End Sub

'Add by Amy 2014/10/30
Private Function FormCheck() As Boolean
    Dim bCancel As Boolean
    
    MaskEdBox2_Validate bCancel
    If bCancel = True Then
        MaskEdBox2.SetFocus
        Exit Function
    End If
    FormCheck = True
End Function

