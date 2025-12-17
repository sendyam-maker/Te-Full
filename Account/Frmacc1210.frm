VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc1210 
   AutoRedraw      =   -1  'True
   Caption         =   "收據資料查詢"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5415
   ScaleWidth      =   8940
   Begin VB.TextBox txtInvoice 
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
      Height          =   310
      Left            =   4920
      TabIndex        =   27
      Top             =   2280
      Width           =   2000
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
      Height          =   310
      Left            =   1320
      TabIndex        =   24
      Top             =   2280
      Width           =   2000
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
      Height          =   310
      Left            =   1320
      TabIndex        =   21
      Top             =   1920
      Width           =   675
   End
   Begin VB.CommandButton Command1 
      Caption         =   "回上一畫面"
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
      Left            =   6960
      TabIndex        =   1
      Top             =   120
      Width           =   1452
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc1210.frx":0000
      Height          =   2475
      Left            =   90
      TabIndex        =   2
      Top             =   2670
      Width           =   8500
      _ExtentX        =   15002
      _ExtentY        =   4366
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   16777215
      HeadLines       =   1
      RowHeight       =   18
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
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "收據號碼查詢"
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "a0j02"
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
         DataField       =   "a0j01"
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
         DataField       =   "na03"
         Caption         =   "申請國家"
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
         DataField       =   "cp10N"
         Caption         =   "帳款類別"
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
         DataField       =   "RAmount"
         Caption         =   "應收金額"
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
         DataField       =   "EAmount"
         Caption         =   "已收金額"
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
      BeginProperty Column06 
         DataField       =   "a1u01"
         Caption         =   "單據編號"
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
         DataField       =   "a0l02"
         Caption         =   "單據日期"
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
      BeginProperty Column08 
         DataField       =   "a1u07"
         Caption         =   "銷帳服務費"
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
         DataField       =   "a1u09"
         Caption         =   "銷帳規費"
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
         DataField       =   "a1u08"
         Caption         =   "銷退服務費"
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
      BeginProperty Column11 
         DataField       =   "a1u10"
         Caption         =   "銷退規費"
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
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1335.118
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "收款資料"
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
      Left            =   5640
      TabIndex        =   0
      Top             =   120
      Width           =   1212
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
      Height          =   310
      Left            =   4440
      TabIndex        =   16
      Top             =   1560
      Width           =   852
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   14
      Top             =   1560
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
   Begin VB.TextBox Text6 
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
      Height          =   310
      Left            =   1320
      TabIndex        =   9
      Top             =   840
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
      Height          =   310
      Left            =   2160
      TabIndex        =   7
      Top             =   480
      Width           =   2295
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
      Height          =   310
      Left            =   1320
      TabIndex        =   6
      Top             =   480
      Width           =   852
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
      Height          =   310
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   6840
      TabIndex        =   20
      Top             =   1560
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   2664
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
   Begin MSForms.TextBox Text13 
      Height          =   315
      Left            =   1995
      TabIndex        =   25
      Top             =   1920
      Width           =   1320
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      Size            =   "2328;556"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text10 
      Height          =   375
      Left            =   4440
      TabIndex        =   18
      Top             =   1890
      Width           =   3975
      VariousPropertyBits=   -1466941409
      BackColor       =   14737632
      ScrollBars      =   2
      Size            =   "7011;661"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text8 
      Height          =   315
      Left            =   1320
      TabIndex        =   12
      Top             =   1200
      Width           =   7095
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      Size            =   "12515;556"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   315
      Left            =   2880
      TabIndex        =   10
      Top             =   840
      Width           =   5535
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      Size            =   "9763;556"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "智權發票號碼"
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
      Left            =   3480
      TabIndex        =   26
      Top             =   2310
      Width           =   1395
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "發票號碼"
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
      Top             =   2310
      Width           =   975
   End
   Begin VB.Label Label10 
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
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label9 
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
      Height          =   255
      Left            =   5880
      TabIndex        =   19
      Top             =   1560
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label8 
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
      Left            =   3480
      TabIndex        =   17
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "列印次數"
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
      Left            =   3480
      TabIndex        =   15
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "收據日期"
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
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭"
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
      TabIndex        =   11
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "客戶代號"
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
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
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
      Left            =   360
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收據號碼"
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
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc1210"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/16 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit

Public adoacc0k0 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset

Private Sub Command1_Click()
   strExitControl = MsgText(601)
   strCondition1 = MsgText(601)
   strCondition2 = MsgText(601)
   'Mark by Amy 2014/03/18 搬至form_load
'   tool3_enabled
'   Frmacc1211.Show
   'end 2014/03/18
   Unload Me
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Command2_Click()
   If Text1 = "" Then
      Exit Sub
   End If
   strCon1 = Text1
   tool3_enabled
   Frmacc1212.Show
   Me.Enabled = False
End Sub

Private Sub Command2_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Activate()
   strFormName = Name
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   'Modify by Amy 2023/10/06 原W:8850/H:5700切畫面不用再調-瑞婷
   Me.Width = 9060
   'Modify By Cheng 2002/03/28
'   Me.Height = 5400
   Me.Height = 5880
   'Modify by Amy 2023/10/06 原(lngWidth - Me.Width) / 2
   Me.Move 0, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   If strItemNo <> "" Then
      Text1 = strItemNo
   End If
   OpenTable
   If adoacc0k0.RecordCount <> 0 Then
      FormShow
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strExitControl = MsgText(602) Then
      strFormName = MsgText(601)
      KeyEnter vbKeyEscape
      MenuEnabled
      Set Frmacc1210 = Nothing
      Exit Sub
   End If
   'Add by  Amy 2014/03/18
   tool3_enabled
   Frmacc1211.Show
   'end 2014/03/18
   strExitControl = MsgText(602)
   Set Frmacc1210 = Nothing
End Sub


Private Sub Text6_Change()
   If Text6 = MsgText(601) Then
      Exit Sub
   End If
   Text7 = CustomerQuery(Text6, 1)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc0k0.CursorLocation = adUseClient
   adoacc0k0.Open "select * from acc0k0 where a0k01 = '" & strItemNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc1.CursorLocation = adUseClient
   'Modified by Morgan 2011/10/25 考慮拆收據應收改抓0j0
   'adoadodc1.Open "select a0j02, a0j21, a0j20, (cp16) as RAmount, 0 as EAmount, '' as a1u01, 0 as a0l02, '' as a1u03, a0j01 from acc0j0, caseprogress where a0j01 = cp09 and a0j13 = '" & Text1 & "' " & _
                  "union select a0j02, a0j21, a0j20, 0 as RAmount, nvl(a1u04, 0)+nvl(a1u05, 0) as EAmount, a1u01, a0l02, a1u03, a0j01 from acc1u0, acc0l0, acc0j0 where a1u01 = a0l01 and a1u03 = a0j01 and a0j13 = '" & Text1 & "' order by a0j01 asc, a1u01 desc", adoTaie, adOpenStatic, adLockReadOnly
   'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
   'Modified by Morgan 2012/4/27 收款資料要判斷收據號(因為可拆收據一收文號可能對多張收據號)
   'modify by sonia 2013/12/12 收款單號改單據編號,收款日期改單據日期,並加顯示銷帳銷退金額,並改排序條件
   'adoadodc1.Open "select a0j02, na03, getcp10desc(cp01,cp10,a0j04) cp10N, nvl(a0j09,0)+nvl(a0j10,0) as RAmount, 0 as EAmount, '' as a1u01, 0 as a0l02, '' as a1u03, a0j01 from acc0j0, caseprogress,nation where a0j01 = cp09(+) and a0j13 = '" & Text1 & "' and na01(+)=a0j04 " & _
                  "union select a0j02, na03, getcp10desc(cp01,cp10,a0j04) cp10N, 0 as RAmount, nvl(a1u04, 0)+nvl(a1u05, 0) as EAmount, a1u01, a0l02, a1u03, a0j01 from acc0l0, acc1u0, acc0j0,caseprogress,nation where a1u01 = a0l01(+) and a1u03(+) = a0j01 and a1u02(+)=a0j13 and a0j13 = '" & Text1 & "' and cp09(+)=a0j01 and na01(+)=a0j04 and a1u01 is not null order by a0j01 asc, a1u01 desc", adoTaie, adOpenStatic, adLockReadOnly
   'adoadodc1.Open "select a0j02, na03, getcp10desc(cp01,cp10,a0j04) cp10N, nvl(a0j09,0)+nvl(a0j10,0) as RAmount, 0 as EAmount, '' as a1u01, a0k02 as a0l02,0 as a1u07,0 as a1u09,0 as a1u08,0 as a1u10, '' as a1u03, a0j01 from acc0j0,acc0k0,caseprogress,nation where a0j01 = cp09(+) and a0j13 = '" & Text1 & "' and na01(+)=a0j04 and a0j13=a0k01(+) " & _
                  "union select a0j02, na03, getcp10desc(cp01,cp10,a0j04) cp10N, 0 as RAmount, nvl(a1u04, 0)+nvl(a1u05, 0) as EAmount, a1u01,nvl(a0s03,a0l02) a0l02,a1u07,a1u09,a1u08,a1u10, a1u03, a0j01 from acc0l0,acc0s0,acc1u0,acc0j0,caseprogress,nation where a1u01 = a0l01(+) and a1u01 = a0s01(+) and a1u03(+) = a0j01 and a1u02(+)=a0j13 and a0j13 = '" & Text1 & "' and cp09(+)=a0j01 and na01(+)=a0j04 and a1u01 is not null " & _
                  "order by a0j01, a0l02,a1u01 ", adoTaie, adOpenStatic, adLockReadOnly
   '2014/12/8 modify by sonia E10328594補扣繳之單據編號改顯示'補扣繳'
   'modify by sonia 2021/10/22 Grid中之案件性質改為帳款類別getcp10desc(cp01,cp10,a0j04)改為nvl(a0j22,Getcp10desc(Cp01,Cp10,A0j04))
   adoadodc1.Open "select a0j02, na03, nvl(a0j22,Getcp10desc(Cp01,Cp10,A0j04)) cp10N, nvl(a0j09,0)+nvl(a0j10,0) as RAmount, 0 as EAmount, '' as a1u01, a0k02 as a0l02,0 as a1u07,0 as a1u09,0 as a1u08,0 as a1u10, '' as a1u03, a0j01 from acc0j0,acc0k0,caseprogress,nation where a0j01 = cp09(+) and a0j13 = '" & Text1 & "' and na01(+)=a0j04 and a0j13=a0k01(+) " & _
                  "union select a0j02, na03, nvl(a0j22,Getcp10desc(Cp01,Cp10,A0j04)) cp10N, 0 as RAmount, nvl(a1u04, 0)+nvl(a1u05, 0) as EAmount, decode(a1u01,a1u03,'補扣繳',a1u01) a1u01,nvl(a0s03,a0l02) a0l02,a1u07,a1u09,a1u08,a1u10, a1u03, a0j01 from acc0l0,acc0s0,acc1u0,acc0j0,caseprogress,nation where a1u01 = a0l01(+) and a1u01 = a0s01(+) and a1u03(+) = a0j01 and a1u02(+)=a0j13 and a0j13 = '" & Text1 & "' and cp09(+)=a0j01 and na01(+)=a0j04 and a1u01 is not null " & _
                  "order by a0j01, a0l02,a1u01 ", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
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
   'Add By Cheng 2002/03/28
   Dim rs As New ADODB.Recordset
   
   Text1 = adoacc0k0.Fields("a0k01").Value
   If IsNull(adoacc0k0.Fields("a0k11").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = adoacc0k0.Fields("a0k11").Value
      Select Case Text2
         Case "1"
            Text3 = MsgText(901)
         Case "2"
            'Modify by Amy 2021/08/17
            'Text3 = MsgText(902)
            Text3 = A0802Query(Text2, True)
         Case "3"
            Text3 = MsgText(903)
         Case "5"
            Text3 = MsgText(904)
         Case "7"
            Text3 = MsgText(905)
         Case "8"
            Text3 = MsgText(906)
      End Select
   End If

   If IsNull(adoacc0k0.Fields("a0k03").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = adoacc0k0.Fields("a0k03").Value
   End If
   If IsNull(adoacc0k0.Fields("a0k04").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = adoacc0k0.Fields("a0k04").Value
   End If
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc0k0.Fields("a0k02").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(adoacc0k0.Fields("a0k02").Value)
   End If
   MaskEdBox1.Mask = DFormat
   If IsNull(adoacc0k0.Fields("a0k19").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = adoacc0k0.Fields("a0k19").Value
   End If
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(adoacc0k0.Fields("a0k09").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(adoacc0k0.Fields("a0k09").Value)
   End If
   MaskEdBox2.Mask = DFormat
   If IsNull(adoacc0k0.Fields("a0k20").Value) Then
      Text11 = MsgText(601)
      Text13 = MsgText(601)
   Else
      Text11 = adoacc0k0.Fields("a0k20").Value
      Text13 = StaffQuery(Text11)
   End If
   If IsNull(adoacc0k0.Fields("a0k08").Value) Then
      Text10 = MsgText(601)
   Else
      Text10 = adoacc0k0.Fields("a0k08").Value
   End If
   
   'Add By Cheng 2002/03/28
   If rs.State <> adStateClosed Then rs.Close
   rs.CursorLocation = adUseClient
   rs.Open "Select Distinct A0M03 FROM ACC0M0 WHERE A0M02='" & Me.Text1.Text & "' ORDER BY A0M03", _
            cnnConnection, adOpenStatic, adLockReadOnly
   If rs.RecordCount > 0 Then
      rs.MoveFirst
      Me.Text12.Text = "" & rs.Fields(0).Value
      rs.MoveNext
      While Not rs.EOF
         Me.Text12.Text = Me.Text12.Text & "," & rs.Fields(0).Value
         rs.MoveNext
      Wend
   Else
      Me.Text12.Text = ""
   End If
   If rs.State <> adStateClosed Then rs.Close
   'Add by Amy 2013/12/17 +智權發票號
   rs.CursorLocation = adUseClient
   rs.Open "Select Axc01 From Acc431 Where Axc02='" & Me.Text1 & "' ", cnnConnection, adOpenStatic, adLockReadOnly
   If rs.RecordCount > 0 Then
      Me.txtInvoice = "" & rs.Fields(0).Value
   Else
      Me.txtInvoice = ""
   End If
   'end 2013/12/17
   If rs.State <> adStateClosed Then rs.Close
   Set rs = Nothing
End Sub
