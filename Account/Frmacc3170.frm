VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc3170 
   AutoRedraw      =   -1  'True
   Caption         =   "票據貼現作業"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   8760
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
      Left            =   3840
      TabIndex        =   26
      Top             =   3528
      Width           =   1140
   End
   Begin VB.TextBox Text9 
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
      Height          =   300
      Left            =   6840
      TabIndex        =   4
      Top             =   600
      Width           =   1572
   End
   Begin VB.CommandButton Command2 
      Height          =   300
      Left            =   5280
      Picture         =   "Frmacc3170.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   600
      Width           =   350
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc3170.frx":0102
      Height          =   2400
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4233
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
      Caption         =   "票據貼現資料"
      ColumnCount     =   8
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
      BeginProperty Column02 
         DataField       =   "a0e44"
         Caption         =   "貼現利息"
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
         DataField       =   "a0e43"
         Caption         =   "貼現金額"
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
      BeginProperty Column06 
         DataField       =   "a0e05"
         Caption         =   "往來類別"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1649.764
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1530.142
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   555.024
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column07 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   960
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
   Begin VB.TextBox Text7 
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
      Left            =   360
      MaxLength       =   8
      TabIndex        =   5
      Top             =   4368
      Width           =   1452
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   1572
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   12648447
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
      Height          =   300
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   1
      Top             =   240
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      Height          =   492
      Left            =   7920
      Picture         =   "Frmacc3170.frx":0117
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   4248
      Width           =   492
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
      Height          =   300
      Left            =   4800
      TabIndex        =   17
      Top             =   4368
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
      Left            =   3240
      TabIndex        =   16
      Top             =   4368
      Width           =   1572
   End
   Begin VB.TextBox Text5 
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
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   6
      Top             =   4368
      Width           =   1452
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
      Height          =   300
      Left            =   5040
      MaxLength       =   14
      TabIndex        =   10
      Top             =   3528
      Width           =   1455
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
      Left            =   2160
      TabIndex        =   14
      Top             =   3528
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
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
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
      TabIndex        =   0
      Top             =   240
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   6360
      TabIndex        =   24
      Top             =   4368
      Width           =   1452
      _ExtentX        =   2566
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
   Begin MSForms.TextBox Text11 
      Height          =   300
      Left            =   5670
      TabIndex        =   21
      Top             =   240
      Width           =   2745
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "4854;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
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
      Height          =   252
      Left            =   1320
      TabIndex        =   27
      Top             =   3528
      Width           =   492
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "貼現利率"
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
      TabIndex        =   25
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  '置中對齊
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
      Left            =   2040
      TabIndex        =   23
      Top             =   4128
      Width           =   972
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "貼現日期"
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
      TabIndex        =   22
      Top             =   600
      Width           =   1212
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4008
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   852
      Left            =   240
      Top             =   4008
      Width           =   8292
   End
   Begin VB.Label Label10 
      Alignment       =   2  '置中對齊
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
      Height          =   252
      Left            =   6480
      TabIndex        =   20
      Top             =   4128
      Width           =   1212
   End
   Begin VB.Label Label8 
      Alignment       =   2  '置中對齊
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
      Height          =   252
      Left            =   4920
      TabIndex        =   19
      Top             =   4128
      Width           =   1332
   End
   Begin VB.Label Label6 
      Alignment       =   2  '置中對齊
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
      Height          =   252
      Left            =   3480
      TabIndex        =   18
      Top             =   4128
      Width           =   1092
   End
   Begin VB.Label Label5 
      Alignment       =   2  '置中對齊
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
      Left            =   600
      TabIndex        =   15
      Top             =   4128
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "貼現序號"
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
      TabIndex        =   13
      Top             =   600
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "貼現帳號"
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
      TabIndex        =   12
      Top             =   240
      Width           =   972
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "貼現銀行"
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
      TabIndex        =   11
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc3170"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/20 Form2.0已修改 Text11/DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc0e0 As New ADODB.Recordset
Public adoacc0f0 As New ADODB.Recordset
Public adoacc0g0 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoaccnum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adocheck As New ADODB.Recordset

Private Sub Command1_Click()
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.Fields("a0e17").Value = 0
      Adodc1.Recordset.Fields("a0e18").Value = Null
      Adodc1.Recordset.UpdateBatch
      AdodcRefresh
      SumShow
      SumSave
      AdodcClear
   End If
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Command2_Click()
   If adoacc0f0.RecordCount = 0 Or MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Or Text2 = MsgText(601) Then
      Exit Sub
   End If
   adoacc0f0.Find "a0f01 = " & Val(FCDate(MaskEdBox1.Text)) & "", 0, adSearchForward, 1
   If adoacc0f0.EOF = False Then
      adoacc0f0.Find "a0f02 = '" & Text2 & "'", 0, adSearchForward, adoacc0f0.Bookmark
      If adoacc0f0.EOF = False Then
         FormShow
         AdodcRefresh
         RecordShow
      Else
         MsgBox MsgText(33), , MsgText(5)
         adoacc0f0.MoveFirst
      End If
   Else
      MsgBox MsgText(33), , MsgText(5)
      adoacc0f0.MoveFirst
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
   AdodcShow
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strCompanyNo = MsgText(601) Then
      Exit Sub
   End If
   If adoacc0f0.RecordCount <> 0 Then
      adoacc0f0.MoveFirst
   End If
   adoacc0f0.Find "a0f01 = " & Val(strCompanyNo) & "", 0, adSearchForward, 1
   If adoacc0f0.EOF = False Then
      adoacc0f0.Find "a0f02 = '" & strItemNo & "'", 0, adSearchForward, adoacc0f0.Bookmark
      If adoacc0f0.EOF = False Then
         FormShow
         AdodcRefresh
         RecordShow
      End If
   End If
   strCompanyNo = MsgText(601)
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
   Me.Width = 8850
   Me.Height = 5500
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
   OpenTable
   If adoacc0f0.RecordCount <> 0 Then
      adoacc0f0.MoveLast
      adoacc0f0.MoveFirst
      RecordShow
   End If
   FormDisabled
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
   Set Frmacc3170 = Nothing
End Sub

Private Sub MaskEdBox1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox Label12 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label12 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 = MsgText(601) Then
      MsgBox Label1 & MsgText(52), , MsgText(5)
      Cancel = True
      Exit Sub
   Else
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select a0h01, a0h02 from acc0h0 where a0h02 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
         If IsNull(adocheck.Fields(0).Value) = False Then
            Text10 = adocheck.Fields(0).Value
            adocheck.Close
            Exit Sub
         End If
      End If
      MessageShow Label1
      Cancel = True
      adocheck.Close
      Exit Sub
   End If
End Sub

Private Sub Text10_Change()
   If Text10 = MsgText(601) Then
      Exit Sub
   End If
   Text11 = A0g02Query(Text10)
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub

Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc0e0.CursorLocation = adUseClient
   adoacc0e0.Open "select * from acc0e0 where a0e01 = '" & Text5 & "' and a0e02 = '" & Text7 & "' and a0e25 = 0", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0f0.CursorLocation = adUseClient
   adoacc0f0.Open "select * from acc0f0 order by a0f01 desc, a0f02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc0e0 where a0e17 = " & Val(FCDate(MaskEdBox1.Text)) & " and a0e18 = '" & Text2 & "' order by a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表(貼現票據資料)
'
'*************************************************
Public Sub FormShow()
   If IsNull(adoacc0f0.Fields("a0f03").Value) Then
      Text10 = MsgText(601)
   Else
      Text10 = adoacc0f0.Fields("a0f03").Value
   End If
   If IsNull(adoacc0f0.Fields("a0f04").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = adoacc0f0.Fields("a0f04").Value
   End If
   Text2 = adoacc0f0.Fields("a0f02").Value
   MaskEdBox1.Mask = MsgText(601)
   MaskEdBox1.Text = CFDate(adoacc0f0.Fields("a0f01").Value)
   MaskEdBox1.Mask = DFormat
   If IsNull(adoacc0f0.Fields("a0f15").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = adoacc0f0.Fields("a0f15").Value
   End If
End Sub

'*************************************************
'  顯示資料表(票據資料)
'
'*************************************************
Public Sub AdodcShow()
   Text5 = Adodc1.Recordset.Fields("a0e01").Value
   Text7 = Adodc1.Recordset.Fields("a0e02").Value
   If IsNull(Adodc1.Recordset.Fields("a0e11").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = Adodc1.Recordset.Fields("a0e11").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e06").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = Adodc1.Recordset.Fields("a0e06").Value
   End If
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(Adodc1.Recordset.Fields("a0e10").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(Adodc1.Recordset.Fields("a0e10").Value)
   End If
   MaskEdBox2.Mask = DFormat
End Sub

'*************************************************
'  清除顯示(票據資料)
'
'*************************************************
Public Sub AdodcClear()
   Text5 = ""
   Text7 = ""
   Text6 = ""
   Text8 = ""
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
End Sub

'*************************************************
'  儲存資料表(票據資料)
'
'*************************************************
Public Sub AdodcSave()
Dim LngDays As Long

On Error GoTo Checking
   If Text5 = MsgText(601) Then
      MsgBox MsgText(10) & Label7, , MsgText(5)
      strControlButton = MsgText(602)
      Text5.SetFocus
      Exit Sub
   Else
      If Text7 = MsgText(601) Then
         MsgBox MsgText(10) & Label5, , MsgText(5)
         strControlButton = MsgText(602)
         Text7.SetFocus
         Exit Sub
      End If
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select a0g01 from acc0g0 where a0g01 = '" & Text5 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If adocheck.RecordCount = 0 Then
         MessageShow Label5
         strControlButton = MsgText(602)
         adocheck.Close
         Exit Sub
      End If
      adocheck.Close
   End If
   adoacc0e0.Close
   adoacc0e0.CursorLocation = adUseClient
   adoacc0e0.Open "select * from acc0e0 where a0e01 = '" & Text5 & "' and a0e02 = '" & Text7 & "' and a0e14 = 0 and a0e15 = 0 and a0e21 = 0 and a0e25 = 0 and a0e17 = 0", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc0e0.RecordCount = 0 Then
      MsgBox MsgText(61), , MsgText(5)
      strControlButton = MsgText(602)
      Exit Sub
   End If
'   If adoacc0e0.Fields("a0e17").Value <> 0 Then
'      MsgBox MsgText(9), , MsgText(5)
'      strControlButton = MsgText(602)
'      Exit Sub
'   End If
   adoacc0e0.Fields("a0e17").Value = Val(FCDate(MaskEdBox1.Text))
   adoacc0e0.Fields("a0e18").Value = Text2
   If Text10 <> MsgText(601) Then
      adoacc0e0.Fields("a0e19").Value = Text10
   Else
      adoacc0e0.Fields("a0e19").Value = Null
   End If
   If Text1 <> MsgText(601) Then
      adoacc0e0.Fields("a0e20").Value = Text1
   Else
      adoacc0e0.Fields("a0e20").Value = Null
   End If
   If Text9 <> MsgText(601) Then
      adoacc0e0.Fields("a0e42").Value = Val(Text9)
   Else
      adoacc0e0.Fields("a0e42").Value = 0
   End If
   LngDays = CDays(FCDate(MaskEdBox1.Text), IIf(IsNull(adoacc0e0.Fields("a0e10").Value), 0, adoacc0e0.Fields("a0e10").Value))
   adoacc0e0.Fields("a0e44").Value = Val(Format(adoacc0e0.Fields("a0e11").Value * Val(Text9) * LngDays / 365, DDollar))
   adoacc0e0.Fields("a0e43").Value = Val(adoacc0e0.Fields("a0e11").Value) - Val(adoacc0e0.Fields("a0e44").Value)
   adoacc0e0.UpdateBatch
   AdodcRefresh
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  計算並顯示票載總金額
'
'*************************************************
Private Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(a0e11), sum(a0e43), sum(a0e44) from acc0e0 where a0e17 = " & Val(FCDate(MaskEdBox1.Text)) & " and a0e18 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text3 = MsgText(601)
      Else
         Text3 = adoaccsum.Fields(0).Value
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text4 = MsgText(601)
      Else
         Text4 = adoaccsum.Fields(1).Value
      End If
      If IsNull(adoaccsum.Fields(2).Value) Then
         Text12 = MsgText(601)
      Else
         Text12 = adoaccsum.Fields(2).Value
      End If
   Else
      Text3 = MsgText(601)
      Text4 = MsgText(601)
      Text12 = MsgText(601)
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
         If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
            Exit Sub
         End If
         Frmacc3170_Save
         If strControlButton <> MsgText(602) Then
            AdodcSave
         End If
         If strControlButton <> MsgText(602) Then
            SumSave
            SumShow
            AdodcClear
            Text7.SetFocus
         End If
         strControlButton = MsgText(601)
   End Select
   KeyEnter KeyCode
End Sub

'*************************************************
'  儲存票載總金額
'
'*************************************************
Private Sub SumSave()
   If Text3 <> MsgText(601) Then
      adoacc0f0.Fields("a0f05").Value = Val(Text3)
   Else
      adoacc0f0.Fields("a0f05").Value = 0
   End If
   If Text4 <> MsgText(601) Then
      adoacc0f0.Fields("a0f06").Value = Val(Text4)
   Else
      adoacc0f0.Fields("a0f06").Value = 0
   End If
   If Text9 <> MsgText(601) Then
      adoacc0f0.Fields("a0f15").Value = Val(Text9)
   Else
      adoacc0f0.Fields("a0f15").Value = 0
   End If
   adoacc0f0.UpdateBatch
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   If ExistCheck("acc0g0", "a0g01", Text10, Label9) = False Then
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
   If Val(Text4) <= 0 Then
      MsgBox MsgText(58), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text5_Change()
   QueryAcc0e0
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   If Text5 = MsgText(601) Then
      Exit Sub
   End If
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   If ExistCheck("acc0g0", "a0g01", Text5, Label7) = False Then
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text7_Change()
   QueryAcc0e0
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'*************************************************
'  搜尋票據資料
'
'*************************************************
Private Sub QueryAcc0e0()
   adoacc0e0.Close
   adoacc0e0.CursorLocation = adUseClient
   adoacc0e0.Open "select * from acc0e0 where a0e02 = '" & Text7 & "' and a0e25 = 0", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc0e0.RecordCount > 1 Then
      adoacc0e0.Close
      adoacc0e0.CursorLocation = adUseClient
      adoacc0e0.Open "select * from acc0e0 where a0e01 = '" & Text5 & "' and a0e02 = '" & Text7 & "' and a0e25 = 0", adoTaie, adOpenDynamic, adLockBatchOptimistic
   End If
   If adoacc0e0.RecordCount <> 0 Then
      Text5 = adoacc0e0.Fields("a0e01").Value
      If IsNull(adoacc0e0.Fields("a0e11").Value) Then
         Text6 = MsgText(601)
      Else
         Text6 = adoacc0e0.Fields("a0e11").Value
      End If
      If IsNull(adoacc0e0.Fields("a0e06").Value) Then
         Text8 = MsgText(601)
      Else
         Text8 = adoacc0e0.Fields("a0e06").Value
      End If
      MaskEdBox2.Mask = MsgText(601)
      If IsNull(adoacc0e0.Fields("a0e10").Value) Then
         MaskEdBox2.Text = MsgText(601)
      Else
         MaskEdBox2.Text = CFDate(adoacc0e0.Fields("a0e10").Value)
      End If
      MaskEdBox2.Mask = DFormat
   Else
      Text5 = MsgText(601)
      Text6 = MsgText(601)
      Text8 = MsgText(601)
      MaskEdBox2.Mask = MsgText(601)
      MaskEdBox2.Text = MsgText(601)
      MaskEdBox2.Mask = DFormat
   End If
End Sub

'*************************************************
'  重新整理傳票資料
'
'*************************************************
Public Sub Acc0f0Refresh()
On Error GoTo Checking
   adoacc0f0.Close
   adoacc0f0.CursorLocation = adUseClient
   If strConTitle = MsgText(31) Or strConTitle = MsgText(601) Then
      adoacc0f0.Open "select * from acc0f0 order by a0f01 desc, a0f02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Else
      If strConTitle <> strCon3 Then
         adoacc0f0.Open "select * from acc0f0 where " & strConTitle & " >= '" & strCondition1 & "' and " & strConTitle & " <= '" & strCondition2 & "' order by a0f01 desc, a0f02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      Else
         adoacc0f0.Open "select * from acc0f0 where " & strConTitle & " >= " & Val(strCondition1) & " and " & strConTitle & " <= " & Val(strCondition2) & " order by a0f01 desc, a0f02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      End If
   End If
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
   adoadodc1.Open "select * from acc0e0 where a0e17 = " & Val(FCDate(MaskEdBox1.Text)) & " and a0e18 = '" & Text2 & "' order by a0e01 asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
   SumShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = adoacc0f0.Bookmark & MsgText(35) & adoacc0f0.RecordCount
End Sub

'*************************************************
'  關閉分錄欄位輸入狀態
'
'*************************************************
Public Sub FormDisabled()
   Text7.Enabled = False
   Text5.Enabled = False
   Command1.Enabled = False
End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'
'*************************************************
Public Sub FormEnabled()
   Text7.Enabled = True
   Text5.Enabled = True
   Command1.Enabled = True
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   If Text7 <> MsgText(601) Then
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select a0e01, a0e02 from acc0e0 where a0e02 = '" & Text7 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
         If IsNull(adocheck.Fields(0).Value) Then
            Text5 = MsgText(601)
         Else
            Text5 = adocheck.Fields(0).Value
            QueryAcc0e0
         End If
      Else
         Text5 = MsgText(601)
      End If
      adocheck.Close
   End If
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

'Add by Amy 2020/07/17
'從aacc_sav搬回
Public Sub Frmacc3170_Save()
   On Error GoTo Checking
   With Frmacc3170
      If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
         MsgBox MsgText(10) & .Label12, , MsgText(5)
         strControlButton = MsgText(602)
         .MaskEdBox1.SetFocus
         Exit Sub
      Else
         If DateCheck(.MaskEdBox1.Text) = MsgText(603) Then
            MsgBox .Label12 & MsgText(63), , MsgText(5)
            strControlButton = MsgText(602)
            .MaskEdBox1.SetFocus
            Exit Sub
         End If
         If .Text2 = MsgText(602) Then
            MsgBox MsgText(10) & .Label2, , MsgText(5)
            strControlButton = MsgText(602)
            .Text2.SetFocus
            Exit Sub
         End If
         If .Text10 <> MsgText(601) Then
            If ExistCheck("acc0g0", "a0g01", .Text10, .Label9) = False Then
               strControlButton = MsgText(602)
               .Text10.SetFocus
               Exit Sub
            End If
         End If
         If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
            MsgBox .Label12 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            .MaskEdBox1.SetFocus
            Exit Sub
         End If
      End If
      If strSaveConfirm = MsgText(3) Then
         If .adoacc0f0.RecordCount <> 0 Then
            .adoacc0f0.Find "a0f01 = " & Val(FCDate(.MaskEdBox1.Text)) & "", 0, adSearchForward, 1
            If .adoacc0f0.EOF = False Then
               .adoacc0f0.Find "a0f02 = '" & .Text2 & "'", 0, adSearchForward, .adoacc0f0.Bookmark
               If .adoacc0f0.EOF = False Then
                  Exit Sub
               End If
            End If
         End If
         .adoacc0f0.AddNew
      End If
      
      .adoacc0f0.Fields("a0f01").Value = Val(FCDate(.MaskEdBox1.Text))
      .adoacc0f0.Fields("a0f02").Value = .Text2
      If .Text10 <> MsgText(601) Then
         .adoacc0f0.Fields("a0f03").Value = .Text10
      Else
         .adoacc0f0.Fields("a0f03").Value = Null
      End If
      If .Text1 <> MsgText(601) Then
         .adoacc0f0.Fields("a0f04").Value = .Text1
      Else
         .adoacc0f0.Fields("a0f04").Value = Null
      End If
      If .Text3 <> MsgText(601) Then
         .adoacc0f0.Fields("a0f05").Value = Val(.Text3)
      Else
         .adoacc0f0.Fields("a0f05").Value = 0
      End If
      If .Text4 <> MsgText(601) Then
         .adoacc0f0.Fields("a0f06").Value = Val(.Text4)
      Else
         .adoacc0f0.Fields("a0f06").Value = 0
      End If
      If .Text9 <> MsgText(601) Then
         .adoacc0f0.Fields("A0f15").Value = Val(.Text9)
      Else
         .adoacc0f0.Fields("a0f15").Value = 0
      End If
      If strSaveConfirm = MsgText(3) Then
         .adoacc0f0.Fields("a0f09").Value = Val(strSrvDate(2))
         .adoacc0f0.Fields("a0f10").Value = ServerTime
         .adoacc0f0.Fields("a0f11").Value = strUserNum
      Else
         .adoacc0f0.Fields("a0f12").Value = Val(strSrvDate(2))
         .adoacc0f0.Fields("a0f13").Value = ServerTime
         .adoacc0f0.Fields("a0f14").Value = strUserNum
      End If
      .adoaccnum.CursorLocation = adUseClient
      .adoaccnum.Open "select * from autonumber where au01 = '" & MsgText(816) & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If .adoaccnum.RecordCount = 0 Then
         .adoaccnum.AddNew
         .adoaccnum.Fields("au01").Value = MsgText(816)
      End If
      .adoaccnum.Fields("au02").Value = Mid(.MaskEdBox1.Text, 5, 2) & Mid(.MaskEdBox1.Text, 8, 2)
      .adoaccnum.Fields("au03").Value = Val(.Text2)
      .adoaccnum.UpdateBatch
      .adoaccnum.Close
      .adoacc0f0.UpdateBatch

      .AdodcRefresh
      .RecordShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If

   MsgBox Err.Description, , MsgText(5)
   End With
End Sub

'從aacc_del搬回
Public Sub Frmacc3170_Delete()
On Error GoTo Checking
   With Frmacc3170
      If DeleteCheck("select a0f01 from acc0f0 where a0f01 = " & Val(FCDate(.MaskEdBox1.Text)) & " and a0f02 = '" & .Text2 & "'") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "update acc0e0 set a0e17 = 0, a0e18 = '' where a0e17 = " & Val(FCDate(.MaskEdBox1.Text)) & " and a0e18 = '" & .Text2 & "'"
      .AdodcRefresh
      adoTaie.Execute "delete from acc0f0 where a0f01 = " & Val(FCDate(.MaskEdBox1.Text)) & " and a0f02 = '" & .Text2 & "'"
      .adoacc0f0.Requery
      Frmacc3170_Clear
      .AdodcClear
      If .adoacc0f0.RecordCount <> 0 Then
         .adoacc0f0.MoveFirst
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
Public Sub Frmacc3170_Clear()
   With Frmacc3170
      .Text10 = ""
      .Text11 = ""
      .Text1 = ""
      .MaskEdBox1.Mask = ""
      .MaskEdBox1.Text = ""
      .MaskEdBox1.Mask = DFormat
      .Text2 = ""
      .Text9 = ""
      .Text3 = ""
      .Text4 = ""
      .Text12 = ""
      .Text1.SetFocus
   End With
End Sub
