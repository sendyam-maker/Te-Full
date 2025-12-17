VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc3140 
   AutoRedraw      =   -1  'True
   Caption         =   "支票未領備註說明"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   8805
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
      Left            =   4080
      TabIndex        =   23
      Top             =   1230
      Width           =   1572
   End
   Begin VB.CommandButton Command2 
      Height          =   300
      Left            =   2940
      Picture         =   "Frmacc3140.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   570
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
      Left            =   6840
      MaxLength       =   12
      TabIndex        =   5
      Top             =   1950
      Width           =   1572
   End
   Begin VB.TextBox Text9 
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
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1950
      Width           =   1572
   End
   Begin VB.TextBox Text8 
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
      Left            =   6840
      TabIndex        =   19
      Top             =   1590
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
      Left            =   1320
      TabIndex        =   15
      Top             =   1590
      Width           =   1572
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc3140.frx":0102
      Height          =   2700
      Left            =   240
      TabIndex        =   6
      Top             =   2430
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4763
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
      Caption         =   "未調整支票資料"
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "a0e02"
         Caption         =   "票據號碼"
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
         Caption         =   "收票帳號"
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
         DataField       =   "a0e13"
         Caption         =   "收票日期"
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
         DataField       =   "a0e37"
         Caption         =   "兌領日期"
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
         DataField       =   "a0e08"
         Caption         =   "票別"
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
         DataField       =   "a0e12"
         Caption         =   "備註"
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
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1470.047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1470.047
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1470.047
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1470.047
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   555.024
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   4050.142
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   240
      Top             =   2310
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
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   1
      Top             =   240
      Width           =   1515
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
      TabIndex        =   11
      Top             =   1230
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   4080
      TabIndex        =   17
      Top             =   1590
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
   Begin MSForms.TextBox Text6 
      Height          =   315
      Left            =   1320
      TabIndex        =   25
      Top             =   870
      Width           =   7112
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "12545;556"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text10 
      Height          =   300
      Left            =   2880
      TabIndex        =   21
      Top             =   1950
      Width           =   2772
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   300
      Left            =   5640
      TabIndex        =   12
      Top             =   240
      Width           =   2772
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "連絡電話"
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
      ToolTipText     =   "(1.執行兌現工作 2.沖轉已兌現票據)"
      Top             =   1230
      Width           =   975
   End
   Begin VB.Label Label11 
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
      TabIndex        =   22
      ToolTipText     =   "(1.執行兌現工作 2.沖轉已兌現票據)"
      Top             =   1950
      Width           =   975
   End
   Begin VB.Label Label10 
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
      TabIndex        =   20
      ToolTipText     =   "(1.執行兌現工作 2.沖轉已兌現票據)"
      Top             =   1950
      Width           =   975
   End
   Begin VB.Label Label9 
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
      Left            =   5880
      TabIndex        =   18
      ToolTipText     =   "(1.執行兌現工作 2.沖轉已兌現票據)"
      Top             =   1590
      Width           =   975
   End
   Begin VB.Label Label8 
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
      Left            =   3120
      TabIndex        =   16
      ToolTipText     =   "(1.執行兌現工作 2.沖轉已兌現票據)"
      Top             =   1590
      Width           =   975
   End
   Begin VB.Label Label7 
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
      TabIndex        =   14
      ToolTipText     =   "(1.執行兌現工作 2.沖轉已兌現票據)"
      Top             =   1590
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "備　　註"
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
      Top             =   900
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   2200
      Left            =   225
      Top             =   120
      Width           =   8295
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "帳戶餘額"
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
      TabIndex        =   10
      ToolTipText     =   "(1.執行兌現工作 2.沖轉已兌現票據)"
      Top             =   1230
      Width           =   975
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
      TabIndex        =   9
      Top             =   240
      Width           =   972
   End
   Begin VB.Label Label3 
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
      TabIndex        =   8
      Top             =   570
      Width           =   975
   End
   Begin VB.Label Label1 
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
      TabIndex        =   7
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc3140"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/19 Form2.0已修改 Text5/Text6/Text10/DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc0e0 As New ADODB.Recordset
Public adoacc0h0 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adocheck As New ADODB.Recordset

Private Sub Command2_Click()
   'Modify by Amy 2020/07/16 +Text1不可為空及PKey 判斷
   If Adodc1.Recordset.RecordCount = 0 Or Text2 = MsgText(601) Or Text4 = MsgText(601) Or Text1 = MsgText(601) Then
      Exit Sub
   End If
   Adodc1.Recordset.Find "a0e01 = '" & Text4 & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
'      Adodc1.Recordset.Find "a0e02 = '" & Text2 & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
       Adodc1.Recordset.Find "PKey = '" & Text2 & Text4 & Text1 & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
      If Adodc1.Recordset.EOF = False Then
           FormShow
           AdodcRefresh
           RecordShow
      Else
           MsgBox MsgText(33), , MsgText(5)
           Adodc1.Recordset.MoveFirst
      End If
   'end 2020/07/16
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
   FormShow
End Sub

Private Sub Form_Activate()
   Frmacc0000.Toolbar1.Buttons.Item(9).Enabled = False 'Add by Amy 2020/07/16 不使用查詢
   strFormName = Name
   If strCompanyNo = MsgText(601) Then
      Exit Sub
   End If
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.MoveFirst
   End If
   Adodc1.Recordset.Find "a0e01 = '" & strCompanyNo & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      'Modify by Amy 2020/07/16 +PKey 因a0e07改為key
'      Adodc1.Recordset.Find "a0e02 = '" & strItemNo & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
      Adodc1.Recordset.Find "PKey = '" & strItemNo & strCompanyNo & strBankAcc & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
      If Adodc1.Recordset.EOF = False Then
            FormShow
            RecordShow
      End If
   End If
   
   strCompanyNo = MsgText(601)
   strBankAcc = MsgText(601) 'Add by Amy 2020/07/16
End Sub

'Add by Amy 2021/10/19
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Form2.0 記錄鍵盤傳入順序
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer, intY As Integer, sglWidth As Single, sglHeight As Single
   
   Frmacc0000.Toolbar1.Buttons.Item(9).Enabled = False 'Add by Amy 2020/07/16 不使用查詢
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   'Modify by Amy 2023/10/11 原:W8850 H5595
   Me.Width = 8925
   Me.Height = 5895 'Modify by Amy 2020/07/16 原:5500
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
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
   Set Frmacc3140 = Nothing
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Text6.SetFocus
         Exit Sub
   End Select
   KeyDefine KeyCode
   KeyEnter KeyCode
End Sub

'Add by Amy 2020/07/16 票據號碼+收票銀行 多筆需再抓資料
Private Sub Text1_Validate(Cancel As Boolean)
    Dim strQ As String
    
    If Trim(Text2) = MsgText(601) Or Trim(Text4) = MsgText(601) Or Trim(Text1) = MsgText(601) Then
        Exit Sub
    End If
    strQ = "Select a0e01, a0e02,a0e07 From acc0e0 where a0e02 = '" & Text2 & "' And a0e01='" & Text4 & "' And a0e07='" & Text1 & "' And A0E04='P' "
    adocheck.CursorLocation = adUseClient
    adocheck.Open strQ, adoTaie, adOpenDynamic, adLockBatchOptimistic
    If adocheck.RecordCount <> 0 Then
         If IsNull(adocheck.Fields(2).Value) Then
             Text1 = MsgText(601)
         Else
             Text1 = adocheck.Fields(2).Value
             KeyDefine vbKeyF12
         End If
    Else
         Call DataClear
         Adodc1.Recordset.MoveFirst
    End If
    adocheck.Close
End Sub
'end 2020/07/16

Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub

Private Sub Text11_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
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
   If Text2 <> MsgText(601) Then
      adocheck.CursorLocation = adUseClient
      '2005/12/7 MODIFY BY SONIA 加A0E04條件
      'adocheck.Open "select a0e01, a0e02 from acc0e0 where a0e02 = '" & Text2 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      adocheck.Open "select a0e01, a0e02 from acc0e0 where a0e02 = '" & Text2 & "' AND A0E04='P' ", adoTaie, adOpenDynamic, adLockBatchOptimistic
      'Modify by Amy 2020/07/16 一筆才預帶 收票銀行
      If adocheck.RecordCount <> 0 And adocheck.RecordCount = 1 Then
         If IsNull(adocheck.Fields(0).Value) Then
            Text4 = MsgText(601)
         Else
            Text4 = adocheck.Fields(0).Value
            'KeyDefine vbKeyF12
      'end 2020/07/16
         End If
      Else
         Text4 = MsgText(601)
      End If
      adocheck.Close
   End If
End Sub

Private Sub Text4_Change()
   If Text4 = MsgText(601) Then
      Exit Sub
   End If
   Text5 = A0g02Query(Text4)
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
   Dim strQ As String 'Add by Amy 2020/07/16

   'Add by Amy 2020/07/16
    If Trim(Text2) = MsgText(601) Or Trim(Text4) = MsgText(601) Then
        Exit Sub
    End If
    strQ = "Select a0e01, a0e02,a0e07 From acc0e0 where a0e02 = '" & Text2 & "' And a0e01='" & Text4 & "' And A0E04='P' "
    adocheck.CursorLocation = adUseClient
    adocheck.Open strQ, adoTaie, adOpenDynamic, adLockBatchOptimistic
    If adocheck.RecordCount <> 0 Then
        If IsNull(adocheck.Fields(2).Value) Then
            Text1 = MsgText(601)
        '只有一筆直接預帶
        ElseIf adocheck.RecordCount = 1 Then
            Text1 = adocheck.Fields(2).Value
            KeyDefine vbKeyF12
        End If
    Else
        Text1 = MsgText(601)
    End If
    adocheck.Close
    'end 2020/07/16
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   If ExistCheck("acc0g0", "a0g01", Text4, Label1) = False Then
      Cancel = True
      Exit Sub
   End If
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc0e0.CursorLocation = adUseClient
   '2005/12/7 MODIFY BY SONIA 加A0E04條件
   'adoacc0e0.Open "select * from acc0e0 where a0e01 = '" & Text4 & "' and a0e02 = '" & Text2 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modify by Amy 2020/07/16 +a0e07 因改為key
   adoacc0e0.Open "select * from acc0e0 where a0e01 = '" & Text4 & "' and a0e02 = '" & Text2 & "' And a0e07='" & Text1 & "' AND A0E04='P' ", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0h0.CursorLocation = adUseClient
   adoacc0h0.Open "select a0h01, a0h02, a0h08 from acc0h0 where a0h01 = '" & Text4 & "' and a0h02 = '" & Text1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   '2005/12/7 MODIFY BY SONIA 加A0E04條件
   'adoadodc1.Open "select * from acc0e0 where a0e22 <> 0 and a0e25 = 0 order by a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '2005/12/13 MODIFY BY SONIA 只出現未退票未作廢未調整且已兌領之應付票據
   'adoadodc1.Open "select * from acc0e0 where a0e22 <> 0 and a0e25 = 0 AND A0E04='P' order by a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modify by Amy +PKey 因a0e07改為key,用於Find
   adoadodc1.Open "select acc0e0.*,a0e02||a0e01||a0e07 as PKey from acc0e0 where a0e15 = 0 and a0e25 = 0 and (a0e37 is not null and a0e37 <> 0) and (a0e22 = 0 or a0e22 is null) AND a0e04 = '" & MsgText(19) & "' order by a0e01 asc,PKey asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表(票據資料--調節)
'
'*************************************************
Public Sub FormShow()
   Text4 = Adodc1.Recordset.Fields("a0e01").Value
   Text2 = Adodc1.Recordset.Fields("a0e02").Value
   '2005/12/13 CANCEL BY SONIA 取消調整日期欄
   'MaskEdBox1.Mask = MsgText(601)
   'If IsNull(Adodc1.Recordset.Fields("a0e22").Value) Or Adodc1.Recordset.Fields("a0e22").Value = 0 Then
   '   MaskEdBox1.Text = MsgText(601)
   'Else
   '   MaskEdBox1.Text = CFDate(Adodc1.Recordset.Fields("a0e22").Value)
   'End If
   'MaskEdBox1.Mask = DFormat
   If IsNull(Adodc1.Recordset.Fields("a0e12").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = Adodc1.Recordset.Fields("a0e12").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e07").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = Adodc1.Recordset.Fields("a0e07").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e11").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = Adodc1.Recordset.Fields("a0e11").Value
   End If
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(Adodc1.Recordset.Fields("a0e10").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(Adodc1.Recordset.Fields("a0e10").Value)
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e06").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = Adodc1.Recordset.Fields("a0e06").Value
      Text12 = GetConTel(Text8.Text)
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e19").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = Adodc1.Recordset.Fields("a0e19").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e20").Value) Then
      Text11 = MsgText(601)
   Else
      Text11 = Adodc1.Recordset.Fields("a0e20").Value
   End If
   QueryAcc0h0
End Sub

'*************************************************
'  搜尋銀行帳戶資料
'
'*************************************************
Private Sub QueryAcc0h0()
   adoacc0h0.Close
   adoacc0h0.CursorLocation = adUseClient
   adoacc0h0.Open "select a0h01, a0h02, a0h08 from acc0h0 where a0h01 = '" & Text4 & "' and a0h02 = '" & Text1 & "'"
End Sub

'*************************************************
'  查詢顯示(票據資料)
'
'*************************************************
Private Sub DataShow()
   If IsNull(adoacc0e0.Fields("a0e12").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = adoacc0e0.Fields("a0e12").Value
   End If
   If IsNull(adoacc0e0.Fields("a0e07").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = adoacc0e0.Fields("a0e07").Value
   End If
   If IsNull(adoacc0e0.Fields("a0e11").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = adoacc0e0.Fields("a0e11").Value
   End If
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(adoacc0e0.Fields("a0e10").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(adoacc0e0.Fields("a0e10").Value)
   End If
   If IsNull(adoacc0e0.Fields("a0e06").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = adoacc0e0.Fields("a0e06").Value
   End If
   If IsNull(adoacc0e0.Fields("a0e19").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = adoacc0e0.Fields("a0e19").Value
   End If
   If IsNull(adoacc0e0.Fields("a0e20").Value) Then
      Text11 = MsgText(601)
   Else
      Text11 = adoacc0e0.Fields("a0e20").Value
   End If
End Sub

'*************************************************
'  清除查詢顯示
'
'*************************************************
Private Sub DataClear()
   'Text1 = "" 'Mark by Amy 2020/07/16
   Text3 = ""
   Text6 = ""
   'Add by Amy 2020/07/16
   Text12 = ""
   Text7 = ""
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text8 = ""
   Text9 = ""
   Text10 = ""
   Text11 = ""
   'end 2020/07/16
End Sub

'*************************************************
'  搜尋票據資料
'
'*************************************************
Private Sub QueryAcc0e0()
   adoacc0e0.Close
   adoacc0e0.CursorLocation = adUseClient
   '2005/12/7 MODIFY BY SONIA 加A0E04條件
   'adoacc0e0.Open "select * from acc0e0 where a0e01 = '" & Text4 & "' and a0e02 = '" & Text2 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modify by Amy 2020/07/16 +a0e07 因改為key
   adoacc0e0.Open "select * from acc0e0 where a0e01 = '" & Text4 & "' and a0e02 = '" & Text2 & "' And a0e07='" & Text1 & "' AND A0E04='P' ", adoTaie, adOpenDynamic, adLockBatchOptimistic
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
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   If strConTitle = MsgText(31) Or strConTitle = MsgText(601) Then
      '2005/12/7 MODIFY BY SONIA 加A0E04條件
      'adoadodc1.Open "select * from acc0e0 where a0e22 <> 0 and a0e25 = 0 order by a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      '2005/12/13 MODIFY BY SONIA 只出現未退票未作廢未調整且已兌領之應付票據
      'adoadodc1.Open "select * from acc0e0 where a0e22 <> 0 and a0e25 = 0 AND A0E04='P' order by a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      'Modify by Amy 2020/07/16 +PKey 因a0e07改為key
      adoadodc1.Open "select acc0e0.*,a0e02||a0e01||a0e07 as PKey from acc0e0 where a0e15 = 0 and a0e25 = 0 and (a0e37 is not null and a0e37 <> 0) and (a0e22 = 0 or a0e22 is null) AND a0e04 = '" & MsgText(19) & "' order by a0e01 asc,PKey asc, a0e02 asc,a0e07 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Else
      If strConTitle <> strCon6 And strConTitle <> strCon7 Then
         '2005/12/7 MODIFY BY SONIA 加A0E04條件
         'adoadodc1.Open "select * from acc0e0 where a0e22 <> 0 and a0e25 = 0 and " & strConTitle & " >= '" & strCondition1 & "' and " & strConTitle & " <= '" & strCondition2 & "' order by a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         '2005/12/13 MODIFY BY SONIA 只出現未退票未作廢未調整且已兌領之應付票據
         'adoadodc1.Open "select * from acc0e0 where a0e22 <> 0 and a0e25 = 0 AND A0E04='P' and " & strConTitle & " >= '" & strCondition1 & "' and " & strConTitle & " <= '" & strCondition2 & "' order by a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         adoadodc1.Open "select acc0e0.*,a0e02||a0e01||a0e07 as PKey from acc0e0 where a0e15 = 0 and a0e25 = 0 and (a0e37 is not null and a0e37 <> 0) and (a0e22 = 0 or a0e22 is null) AND a0e04 = '" & MsgText(19) & "' and " & strConTitle & " >= '" & strCondition1 & "' and " & strConTitle & " <= '" & strCondition2 & "' order by a0e01 asc,PKey asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      Else
         '2005/12/7 MODIFY BY SONIA 加A0E04條件
         'adoadodc1.Open "select * from acc0e0 where a0e22 <> 0 and a0e25 = 0 and " & strConTitle & " >= " & Val(strCondition1) & " and " & strConTitle & " <= " & Val(strCondition2) & " order by a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         '2005/12/13 MODIFY BY SONIA 只出現未退票未作廢未調整且已兌領之應付票據
         'adoadodc1.Open "select * from acc0e0 where a0e22 <> 0 and a0e25 = 0 AND A0E04='P' and " & strConTitle & " >= " & Val(strCondition1) & " and " & strConTitle & " <= " & Val(strCondition2) & " order by a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         adoadodc1.Open "select acc0e0.*,a0e02||a0e01||a0e07 as PKey from acc0e0 where a0e15 = 0 and a0e25 = 0 and (a0e37 is not null and a0e37 <> 0) and (a0e22 = 0 or a0e22 is null) AND a0e04 = '" & MsgText(19) & "' and " & strConTitle & " >= " & Val(strCondition1) & " and " & strConTitle & " <= " & Val(strCondition2) & " order by a0e01 asc,PKey asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      End If
   End If
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      'Modify by Amy 2020/07/16 +Text1不可為空及PKey 判斷
      If Text4 <> MsgText(601) And Text2 <> MsgText(601) And Text1 <> MsgText(601) Then
         Adodc1.Recordset.Find "a0e01 = '" & Text4 & "'", 0, adSearchForward, 1
         If Adodc1.Recordset.EOF = False Then
'            Adodc1.Recordset.Find "a0e02 = '" & Text2 & "'", 0, adSearchForward, , Adodc1.Recordset.Bookmark
             Adodc1.Recordset.Find "PKey= '" & Text2 & Text4 & Text1 & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
            If Adodc1.Recordset.EOF = False Then
                  FormShow
                  RecordShow
            End If
         End If
      End If
      'end 2020/07/16
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
   'add by nickc 2007/07/13 將輸入法改成使用API
   OpenIme
End Sub

'Modify by Amy 021/10/19 原:KeyCode As Integer
Private Sub Text6_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
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
   Select Case KeyCode
      Case vbKeyF12
         QueryAcc0e0
         QueryAcc0h0
   End Select
   KeyEnter KeyCode
End Sub

'add by nickc 2007/07/13 將輸入法改成使用API
Private Sub Text6_Validate(Cancel As Boolean)
CloseIme
End Sub

Private Sub Text9_Change()
   If Text9 = MsgText(601) Then
      Exit Sub
   End If
   Text10 = A0g02Query(Text9)
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

Private Sub Text9_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   If Text9 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc0g0", "a0g01", Text9, Label10) = False Then
      Cancel = True
      Exit Sub
   End If
End Sub

Private Function GetConTel(ByRef p_CustNo As String) As String
   strExc(0) = "select cu16 from customer where cu01||cu02='" & p_CustNo & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      GetConTel = "" & RsTemp.Fields(0)
   End If
   
End Function

'Add by Amy 2020/07/16
'從aacc_sav搬回來
Public Sub Frmacc3140_Save()
Dim strAuto As String

   On Error GoTo Checking
   With Frmacc3140
      If .Text4 = MsgText(601) Then
         MsgBox MsgText(10) & .Label1, , MsgText(5)
         strControlButton = MsgText(602)
         .Text4.SetFocus
         Exit Sub
      Else
         If .Text2 = MsgText(601) Then
            MsgBox MsgText(10) & .Label2, , MsgText(5)
            strControlButton = MsgText(602)
            .Text2.SetFocus
            Exit Sub
         End If
         If ExistCheck("acc0g0", "a0g01", .Text4, .Label1) = False Then
            strControlButton = MsgText(602)
            .Text4.SetFocus
            Exit Sub
         End If
         If .Text9 <> MsgText(601) Then
            If ExistCheck("acc0g0", "a0g01", .Text9, .Label10) = False Then
               strControlButton = MsgText(602)
               .Text9.SetFocus
               Exit Sub
            End If
         End If
      End If
      'Add by Amy 2021/10/19
      If PUB_ChkUniText(Me) = False Then
         strControlButton = MsgText(602)
         Exit Sub
      End If

      adoTaie.BeginTrans
      .adoacc0e0.Close
      .adoacc0e0.CursorLocation = adUseClient
      '2005/12/7 MODIFY BY SONIA 加A0E04條件
      '.adoacc0e0.Open "select * from acc0e0 where a0e01 = '" & .Text4 & "' and a0e02 = '" & .Text2 & "' and a0e25 = 0", adoTaie, adOpenDynamic, adLockBatchOptimistic
      'Modify by Amy 2020/07/16 +a0e07 因改為key
      .adoacc0e0.Open "select * from acc0e0 where a0e01 = '" & .Text4 & "' and a0e02 = '" & .Text2 & "' And a0e07='" & Text1 & "' and a0e25 = 0 AND A0E04='P' ", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If .adoacc0e0.RecordCount = 0 Then
         MsgBox MsgText(33) & " " & MsgText(39), , MsgText(5)
         Exit Sub
      End If
      If .Text1 <> MsgText(601) Then
         .adoacc0e0.Fields("a0e07").Value = .Text1
      Else
         .adoacc0e0.Fields("a0e07").Value = Null
      End If
      '2005/12/13 CANCEL BY SONIA 取消此欄位
      'If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
      '   .adoacc0e0.Fields("a0e22").Value = Val(FCDate(.MaskEdBox1.Text))
      'Else
      '   .adoacc0e0.Fields("a0e22").Value = Null
      'End If
      '2005/12/13 END
      If .Text6 <> MsgText(601) Then
         .adoacc0e0.Fields("a0e12").Value = .Text6
      Else
         .adoacc0e0.Fields("a0e12").Value = Null
      End If
      If .Text9 <> MsgText(601) Then
         .adoacc0e0.Fields("a0e19").Value = .Text9
      Else
         .adoacc0e0.Fields("a0e19").Value = Null
      End If
      If .Text11 <> MsgText(601) Then
         .adoacc0e0.Fields("a0e20").Value = .Text11
      Else
         .adoacc0e0.Fields("a0e20").Value = Null
      End If
      .adoacc0e0.Fields("a0e21").Value = 0
      .adoacc0e0.Fields("a0e29").Value = Val(strSrvDate(2))
      .adoacc0e0.Fields("a0e30").Value = ServerTime
      .adoacc0e0.Fields("a0e31").Value = strUserNum
'      strAutoNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'L' and a1p04 = '" & .Text2 & .Text4 & "4" & "' and a1p05 = '1120'", 3)
'      adoTaie.Execute "insert into acc1p0 values ('1', 'L', '" & strAutoNo & "', '" & .Text2 & .Text4 & "4" & "', '1120', '" & MsgText(55) & "', 0, " & .adoacc0e0.Fields("a0e11").Value & ", '" & .Text2 & "', '" & .Text9 & "', '" & .Text11 & "', " & _
                      "" & Val(.adoacc0e0.Fields("a0e10").Value) & ", '" & .adoacc0e0.Fields("a0e08").Value & "', '" & .adoacc0e0.Fields("a0e12").Value & "', null, null, null, " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, null, null, " & _
                      "'" & .adoacc0e0.Fields("a0e03").Value & "', null, null, null)"
      .adoacc0e0.UpdateBatch
      adoTaie.CommitTrans
      .AdodcRefresh
      .RecordShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

'從aacc_cls搬回
Public Sub Frmacc3140_Clear()
   With Frmacc3140
      .Text4 = ""
      .Text5 = ""
      .Text1 = ""
      .Text2 = ""
      '2005/12/13 CANCEL BY SONIA
      '.MaskEdBox1.Mask = ""
      '.MaskEdBox1.Text = CFDate(ACDate(ServerDate))
      '.MaskEdBox1.Mask = DFormat
      .Text3 = ""
      .Text6 = ""
      .Text7 = ""
      .MaskEdBox2.Mask = ""
      .MaskEdBox2.Text = ""
      .MaskEdBox2.Mask = DFormat
      .Text8 = ""
      .Text9 = ""
      .Text10 = ""
      .Text11 = ""
      .Text2.SetFocus
   End With
End Sub

'從aacc_del搬回
Public Sub Frmacc3140_Delete()
On Error GoTo Checking
   With Frmacc3140
      'Modify by Amy 2020/07/16 +a0e07 因改為key
      If DeleteCheck("select a0e01 from acc0e0 where a0e01 = '" & .Text4 & "' and a0e02 = '" & .Text2 & "' And a0e07='" & Text1 & "' ") = MsgText(603) Then
         Exit Sub
      End If
      'Modify by Amy 2020/07/16 a1p04 加開票帳號 因a0e07改為key,避免key重覆
      adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'L' and a1p04 = '" & .Text2 & .Text4 & Text1 & "4" & "' and a1p05 = '1120'"
      adoTaie.Execute "update acc0e0 set a0e22 = 0, a0e33 = '' where a0e01 = '" & .Text4 & "' and a0e02 = '" & .Text2 & "' And a0e07='" & Text1 & "' "
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

