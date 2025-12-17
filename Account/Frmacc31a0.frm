VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc31a0 
   AutoRedraw      =   -1  'True
   Caption         =   "票據作廢作業"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   8760
   Begin VB.CommandButton Command2 
      Height          =   300
      Left            =   2550
      Picture         =   "Frmacc31a0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   600
      Width           =   350
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
      Left            =   1320
      TabIndex        =   27
      Top             =   2040
      Width           =   1572
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc31a0.frx":0102
      Height          =   2400
      Left            =   240
      TabIndex        =   25
      Top             =   2520
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
      Caption         =   "票據作廢資料"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "PKey"
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
         Caption         =   "開票銀行"
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
         Caption         =   "開票帳號"
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
         DataField       =   "a0e25"
         Caption         =   "作廢日期"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1289.764
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   2400
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
   Begin VB.TextBox Text14 
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
      TabIndex        =   24
      Top             =   1680
      Width           =   1572
   End
   Begin VB.TextBox Text3 
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
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   2
      Top             =   600
      Width           =   1200
   End
   Begin VB.TextBox Text2 
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
      Left            =   4080
      TabIndex        =   11
      Top             =   1680
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
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   0
      Top             =   240
      Width           =   1572
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
      Height          =   300
      Left            =   4080
      TabIndex        =   10
      Top             =   2040
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
      Height          =   300
      Left            =   1320
      TabIndex        =   9
      Top             =   1320
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
      Left            =   4080
      TabIndex        =   8
      Top             =   1320
      Width           =   1572
   End
   Begin VB.TextBox Text12 
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
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   4080
      TabIndex        =   4
      Top             =   600
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   1320
      TabIndex        =   12
      Top             =   1680
      Width           =   1572
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
   Begin MSForms.TextBox Text4 
      Height          =   300
      Left            =   1320
      TabIndex        =   5
      Top             =   960
      Width           =   7092
      VariousPropertyBits=   -1467989989
      ScrollBars      =   2
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text11 
      Height          =   300
      Left            =   5670
      TabIndex        =   7
      Top             =   1320
      Width           =   2740
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      Size            =   "4833;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text13 
      Height          =   300
      Left            =   5670
      TabIndex        =   6
      Top             =   240
      Width           =   2772
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   2295
      Left            =   240
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "手續費"
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
      TabIndex        =   26
      Top             =   2040
      Width           =   852
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
      Height          =   252
      Left            =   360
      TabIndex        =   23
      Top             =   960
      Width           =   732
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "開票帳號"
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
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "開票銀行"
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
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label8 
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
      Height          =   252
      Left            =   5880
      TabIndex        =   20
      Top             =   1680
      Width           =   612
   End
   Begin VB.Label Label4 
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
      Left            =   3120
      TabIndex        =   19
      Top             =   1680
      Width           =   972
   End
   Begin VB.Label Label7 
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
      Left            =   360
      TabIndex        =   18
      Top             =   1680
      Width           =   972
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
      Height          =   252
      Left            =   3120
      TabIndex        =   17
      Top             =   1320
      Width           =   972
   End
   Begin VB.Label Label5 
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
      TabIndex        =   16
      Top             =   240
      Width           =   972
   End
   Begin VB.Label Label2 
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
      Left            =   3120
      TabIndex        =   15
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "單據號碼"
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
      TabIndex        =   14
      Top             =   2040
      Width           =   972
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "往來類別"
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
      TabIndex        =   13
      Top             =   1320
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc31a0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/20 Form2.0已修改 Text4/Text11/Text13
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc0e0 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adocheck As New ADODB.Recordset

Private Sub Command2_Click()
   'Modify by Amy 2020/07/21 +Text3不可為空及PKey 判斷
   If Adodc1.Recordset.RecordCount = 0 Or Text5 = MsgText(601) Or Text12 = MsgText(601) Or Text3 = MsgText(601) Then
      Exit Sub
   End If
   Adodc1.Recordset.Find "a0e01 = '" & Text12 & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      'Adodc1.Recordset.Find "a0e02 = '" & Text5 & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
      Adodc1.Recordset.Find "PKey = '" & Text5 & Text12 & Text3 & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
   'end 2020/07/21
      If Adodc1.Recordset.EOF = False Then
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
      'Modify by Amy 2020/07/21 +PKey 因a0e07改為key,用於Find
      'Adodc1.Recordset.Find "a0e02 = '" & strItemNo & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
      Adodc1.Recordset.Find "PKey = '" & strItemNo & strCompanyNo & strBankAcc & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
      If Adodc1.Recordset.EOF = False Then
         FormShow
         RecordShow
      End If
   End If
   strCompanyNo = MsgText(601)
   strBankAcc = MsgText(601) 'Add by Amy 2020/07/21
End Sub

'Add by Amy 2021/10/20
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
   strTrackMode = "" 'Add by Amy 2021/10/20 Form2.0 記錄鍵盤傳入順序(清除)
   Set Frmacc31a0 = Nothing
End Sub

Private Sub MaskEdBox1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox Label2 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label2 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Text10_Change()
   If Text10 = MsgText(601) Then
      Exit Sub
   End If
   Select Case Mid(Text9, 1, 1)
      Case Mid(ComboItem(131), 1, 1)
         Text11 = CustomerQuery(Text10, 1)
      Case Mid(ComboItem(132), 1, 1)
         Text11 = A0i02Query(Text10)
      Case Mid(ComboItem(133), 1, 1)
         Text11 = StaffQuery(Text10)
      Case Else
         Text11 = MsgText(601)
   End Select
End Sub

Private Sub Text12_Change()
   If Text12 = MsgText(601) Then
      Exit Sub
   End If
   Text13 = A0g02Query(Text12)
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub

Private Sub Text12_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   If ExistCheck("acc0g0", "a0g01", Text12, Label9) = False Then
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'Add by Amy 2020/07/21
Private Sub Text3_Validate(Cancel As Boolean)
    If Text5 = MsgText(601) Or Text12 = MsgText(601) Or Text3 = MsgText(601) Then
        Exit Sub
    End If
    If adocheck.State = adStateOpen Then adocheck.Close
    adocheck.CursorLocation = adUseClient
    adocheck.Open "select a0e01, a0e02,a0e07 from acc0e0 where a0e04 = '" & MsgText(19) & "' and a0e02 = '" & Text5 & "' And a0e01='" & Text12 & "' And a0e07='" & Text3 & "' and a0e14 = 0 and a0e17 = 0 and a0e21 = 0", adoTaie, adOpenStatic, adLockReadOnly
    If adocheck.RecordCount <> 0 Then
        QueryAcc0e0
        adocheck.Close
        Exit Sub
    End If
    
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
   'add by nickc 2007/07/13 將輸入法改成使用API
   OpenIme
End Sub

'Modify by Amy 2021/10/20 原:KeyCode As Integer
Private Sub Text4_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyDefine Val(KeyCode)
End Sub
'add by nickc 2007/07/13 將輸入法改成使用API
Private Sub Text4_Validate(Cancel As Boolean)
CloseIme
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc0e0.CursorLocation = adUseClient
   'Modify by Amy 2020/07/21 +a0e07 因改為key
   adoacc0e0.Open "select * from acc0e0 where a0e01 = '" & Text12 & "' and a0e02 = '" & Text5 & "' And a0e07='" & Text3 & "' and a0e04 = '" & MsgText(19) & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   'Modify by Amy 2020/07/21 +PKey 因a0e07改為key,用於Find
   adoadodc1.Open "select acc0e0.*,a0e02||a0e01||a0e07 as PKey from acc0e0 where a0e04 = '" & MsgText(19) & "' and a0e25 <> 0 order by a0e25 desc, a0e01 asc,PKey asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
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
   Text5 = Adodc1.Recordset.Fields("a0e02").Value
   Text12 = Adodc1.Recordset.Fields("a0e01").Value
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(Adodc1.Recordset.Fields("a0e25").Value) Or Adodc1.Recordset.Fields("a0e25").Value = 0 Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(Adodc1.Recordset.Fields("a0e25").Value)
   End If
   MaskEdBox1.Mask = DFormat
   If IsNull(Adodc1.Recordset.Fields("a0e07").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = Adodc1.Recordset.Fields("a0e07").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e03").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = Adodc1.Recordset.Fields("a0e03").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e05").Value) Then
      Text9 = MsgText(601)
   Else
      Select Case Adodc1.Recordset.Fields("a0e05").Value
         Case Mid(ComboItem(131), 1, 1)
            Text9 = ComboItem(131)
         Case Mid(ComboItem(132), 1, 1)
            Text9 = ComboItem(132)
         Case Mid(ComboItem(133), 1, 1)
            Text9 = ComboItem(133)
         Case Mid(ComboItem(134), 1, 1)
            Text9 = ComboItem(134)
      End Select
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e06").Value) Then
      Text10 = MsgText(601)
   Else
      Text10 = Adodc1.Recordset.Fields("a0e06").Value
   End If
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(Adodc1.Recordset.Fields("a0e10").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(Adodc1.Recordset.Fields("a0e10").Value)
   End If
   MaskEdBox2.Mask = DFormat
   If IsNull(Adodc1.Recordset.Fields("a0e11").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = Adodc1.Recordset.Fields("a0e11").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e08").Value) Then
      Text14 = MsgText(601)
   Else
      Select Case Adodc1.Recordset.Fields("a0e08").Value
         Case Mid(ComboItem(11), 1, 1)
            Text14 = ComboItem(11)
         Case Mid(ComboItem(12), 1, 1)
            Text14 = ComboItem(12)
         Case Mid(ComboItem(13), 1, 1)
            Text14 = ComboItem(13)
      End Select
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e36").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = Adodc1.Recordset.Fields("a0e36").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e12").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = Adodc1.Recordset.Fields("a0e12").Value
   End If
End Sub

'*************************************************
'  清除查詢顯示
'
'*************************************************
Private Sub DataClear()
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   Text3 = ""
   Text1 = ""
   Text9 = ""
   Text10 = ""
   Text11 = ""
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text2 = ""
   Text14 = ""
   Text4 = ""
   Text6 = ""
End Sub

'*************************************************
'  查詢顯示(票據資料)
'
'*************************************************
Private Sub DataShow()
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc0e0.Fields("a0e25").Value) Or adoacc0e0.Fields("a0e25").Value = 0 Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(adoacc0e0.Fields("a0e25").Value)
   End If
   MaskEdBox1.Mask = DFormat
   If IsNull(adoacc0e0.Fields("a0e07").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = adoacc0e0.Fields("a0e07").Value
   End If
   If IsNull(adoacc0e0.Fields("a0e03").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = adoacc0e0.Fields("a0e03").Value
   End If
   If IsNull(adoacc0e0.Fields("a0e05").Value) Then
      Text9 = MsgText(601)
   Else
      Select Case adoacc0e0.Fields("a0e05").Value
         Case Mid(ComboItem(131), 1, 1)
            Text9 = ComboItem(131)
         Case Mid(ComboItem(132), 1, 1)
            Text9 = ComboItem(132)
         Case Mid(ComboItem(133), 1, 1)
            Text9 = ComboItem(133)
         Case Mid(ComboItem(134), 1, 1)
            Text9 = ComboItem(134)
      End Select
   End If
   If IsNull(adoacc0e0.Fields("a0e06").Value) Then
      Text10 = MsgText(601)
   Else
      Text10 = adoacc0e0.Fields("a0e06").Value
   End If
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(adoacc0e0.Fields("a0e10").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(adoacc0e0.Fields("a0e10").Value)
   End If
   MaskEdBox2.Mask = DFormat
   If IsNull(adoacc0e0.Fields("a0e11").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = adoacc0e0.Fields("a0e11").Value
   End If
   If IsNull(adoacc0e0.Fields("a0e08").Value) Then
      Text14 = MsgText(601)
   Else
      Select Case adoacc0e0.Fields("a0e08").Value
         Case Mid(ComboItem(11), 1, 1)
            Text14 = ComboItem(11)
         Case Mid(ComboItem(12), 1, 1)
            Text14 = ComboItem(12)
         Case Mid(ComboItem(13), 1, 1)
            Text14 = ComboItem(13)
      End Select
   End If
   If IsNull(adoacc0e0.Fields("a0e36").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = adoacc0e0.Fields("a0e36").Value
   End If
   If IsNull(adoacc0e0.Fields("a0e12").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = adoacc0e0.Fields("a0e12").Value
   End If
End Sub

'*************************************************
'  搜尋票據資料
'
'*************************************************
Private Sub QueryAcc0e0()
   adoacc0e0.Close
   adoacc0e0.CursorLocation = adUseClient
   'Modify by Amy 2020/07/21 +a0e07 因改為key
   adoacc0e0.Open "select * from acc0e0 where a0e01 = '" & Text12 & "' and a0e02 = '" & Text5 & "' And a0e07='" & Text3 & "' and a0e04 = '" & MsgText(19) & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
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
   'Modify by Amy 2020/07/21 +PKey 因a0e07改為key,用於Find
   If strConTitle = MsgText(31) Or strConTitle = MsgText(601) Then
      adoadodc1.Open "select acc0e0.*,a0e02||a0e01||a0e07 as PKey from acc0e0 where a0e04 = '" & MsgText(19) & "' and a0e25 <> 0 order by a0e25 desc, a0e01 asc,PKey asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Else
      If strConTitle <> strCon6 And strConTitle <> strCon7 Then
         adoadodc1.Open "select acc0e0.*,a0e02||a0e01||a0e07 as PKey from acc0e0 where a0e04 = '" & MsgText(19) & "' and a0e25 <> 0 and " & strConTitle & " >= '" & strCondition1 & "' and " & strConTitle & " <= '" & strCondition2 & "' order by a0e25 desc, a0e01 asc,PKey asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      Else
         adoadodc1.Open "select acc0e0.*,a0e02||a0e01||a0e07 as PKey from acc0e0 where a0e04 = '" & MsgText(19) & "' and a0e25 <> 0 and " & strConTitle & " >= " & Val(strCondition1) & " and " & strConTitle & " <= " & Val(strCondition2) & " order by a0e25 desc, a0e01 asc,PKey asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      End If
   End If
   'end 2020/07/21
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      If Text5 <> MsgText(601) And Text12 <> MsgText(601) Then
         Adodc1.Recordset.Find "a0e01 = '" & Text12 & "'", 0, adSearchForward, 1
         If Adodc1.Recordset.EOF = False Then
            'Modify by Amy 2020/07/21 +PKey 因a0e07改為key,用於Find
            'Adodc1.Recordset.Find "a0e02 = '" & Text5 & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
            Adodc1.Recordset.Find "PKey = '" & Text5 & Text12 & Text3 & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
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
   'Add by Amy 2021/10/20
   Call PUB_SaveTrackMode(1, KeyCode)
   'Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
   If PUB_ChkTrackMode = False Then
        Exit Sub
   End If
   'end 2021/10/20

   Select Case KeyCode
      Case vbKeyF12
         QueryAcc0e0
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   If Text5 <> MsgText(601) Then
      If adocheck.State = adStateOpen Then adocheck.Close
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select a0e01, a0e02 from acc0e0 where a0e04 = '" & MsgText(19) & "' and a0e02 = '" & Text5 & "' and a0e14 = 0 and a0e17 = 0 and a0e21 = 0", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
         If IsNull(adocheck.Fields(0).Value) = False Then
            Text12 = adocheck.Fields(0).Value
            'QueryAcc0e0 'Mark by Amy 2020/07/21 改至開票帳號做
            adocheck.Close
            Exit Sub
         End If
      End If
      'Mark by Amy 2020/07/21 改至開票帳號做
'      MsgBox MsgText(62), , MsgText(5)
'      Cancel = True
'      adocheck.Close
'      Exit Sub
   End If
End Sub

'Add by Amy 2014/11/05 由aacc_sav搬回
Public Sub Frmacc31a0_Save()

On Error GoTo Checking

   With Frmacc31a0
      If .Text5 = MsgText(601) Then
         MsgBox MsgText(10) & .Label5, , MsgText(5)
         strControlButton = MsgText(602)
         .Text5.SetFocus
         Exit Sub
      Else
         If .Text12 = MsgText(601) Then
            MsgBox MsgText(10) & .Label10, , MsgText(5)
            strControlButton = MsgText(602)
            .Text12.SetFocus
            Exit Sub
         End If
         If ExistCheck("acc0g0", "a0g01", .Text12, .Label10) = False Then
            strControlButton = MsgText(602)
            .Text12.SetFocus
            Exit Sub
         End If
         If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
            MsgBox .Label2 & MsgText(52), , MsgText(5)
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
         End If
         'Add by Amy 2021/10/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
         If PUB_ChkUniText(Me) = False Then
            strControlButton = MsgText(602)
            Exit Sub
        End If

         .adocheck.CursorLocation = adUseClient
         .adocheck.Open "select a0h01, a0h02 from acc0h0 where a0h01 = '" & .Text12 & "' and a0h02 = '" & .Text3 & "'", adoTaie, adOpenStatic, adLockReadOnly
         If .adocheck.RecordCount = 0 Then
            MessageShow .Label10
            strControlButton = MsgText(602)
            .adocheck.Close
            .Text3.SetFocus
            Exit Sub
         End If
         .adocheck.Close
      End If
      .adoacc0e0.Close
      .adoacc0e0.CursorLocation = adUseClient
      'Modify by Amy 2020/07/21 +a0e07 因改為key
      .adoacc0e0.Open "select * from acc0e0 where a0e01 = '" & .Text12 & "' and a0e02 = '" & .Text5 & "' And a0e07='" & Text3 & "' and a0e04 = '" & MsgText(19) & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If .adoacc0e0.RecordCount = 0 Then
         MsgBox MsgText(33), , MsgText(5)
         Exit Sub
      End If
      If .Text3 <> MsgText(601) Then
         .adoacc0e0.Fields("a0e07").Value = .Text3
      Else
         .adoacc0e0.Fields("a0e07").Value = Null
      End If
      If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
         .adoacc0e0.Fields("a0e25").Value = Val(FCDate(.MaskEdBox1.Text))
      Else
         .adoacc0e0.Fields("a0e25").Value = 0
      End If
      If .Text4 <> MsgText(601) Then
         .adoacc0e0.Fields("a0e12").Value = .Text4
      Else
         .adoacc0e0.Fields("a0e12").Value = Null
      End If
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
      
'Add by Morgan 2004/11/11 加transaction

On Error GoTo TranErrHnd

      adoTaie.BeginTrans
'2004/11/11 end
      
      'Modify by Amy 2020/07/21 a1p04加a0e07 因改為key
      adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) (select a1p01, a1p02, a1p03, '" & .Text5 & .Text12 & Text3 & "9" & "', a1p05, a1p06, a1p08, 0, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27 from acc1p0 where a1p01 = '1' and a1p02 = 'L' and a1p04 = '" & .Text5 & .Text12 & Text3 & "2" & "' and a1p08 <> 0)"
      adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) (select a1p01, a1p02, a1p03, '" & .Text5 & .Text12 & Text3 & "9" & "', a1p05, a1p06, 0, a1p07, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27 from acc1p0 where a1p01 = '1' and a1p02 = 'L' and a1p04 = '" & .Text5 & .Text12 & Text3 & "2" & "' and a1p07 <> 0)"
      
      'Add by Morgan 2004/11/11 若是開給智慧局的票同時要清除送件資料
      If .Text10.Text = "V0001" Then
         'Add by Morgan 2005/7/26 改更新 AppList
         strSql = "update applist set al06=null" & _
            " where al01=" & strSrvDate(1) & _
            " and al06='" & .Text5.Text & "'"
         adoTaie.Execute strSql
         
         'Add by Morgan 2011/6/2 更新 AppListe
         strSql = "update appliste set al06=null" & _
            " where al01=" & strSrvDate(1) & _
            " and al06='" & .Text5.Text & "'"
         adoTaie.Execute strSql
      End If
      adoTaie.CommitTrans
      '2004/11/11 END
      
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   End With
   
'Add by Morgan 2004/11/11
   Err.Clear
TranErrHnd:
   If Err.Number <> 0 Then
      adoTaie.RollbackTrans
      MsgBox Err.Description, , MsgText(5)
   End If
End Sub

'Add by Amy 2020/07/21
'從acc_cls搬回
Public Sub Frmacc31a0_Clear()
   With Frmacc31a0
      .Text5 = ""
      .Text12 = ""
      .Text13 = ""
      .MaskEdBox1.Mask = ""
      .MaskEdBox1.Text = ""
      .MaskEdBox1.Mask = DFormat
      .Text3 = ""
      .Text1 = ""
      .Text9 = ""
      .Text10 = ""
      .Text11 = ""
      .MaskEdBox2.Mask = ""
      .MaskEdBox2.Text = ""
      .MaskEdBox2.Mask = DFormat
      .Text2 = ""
      .Text14 = ""
      .Text4 = ""
      .Text5.SetFocus
   End With
End Sub

'從acc_del搬回
Public Sub Frmacc31a0_Delete()
On Error GoTo Checking
   With Frmacc31a0
      'Modify by Amy 2020/07/21 +a0e07 因改為key
      If DeleteCheck("select a0e01 from acc0e0 where a0e01 = '" & .Text12 & "' and a0e02 = '" & .Text5 & "' And a0e07='" & Text3 & "' ") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "update acc0e0 set a0e25 = 0 where a0e01 = '" & .Text12 & "' and a0e02 = '" & .Text5 & "' And a0e07='" & Text3 & "' "
      adoTaie.Execute "delete acc1p0 where a1p01 = '1' and a1p02 = 'L' and a1p04 = '" & .Text5 & .Text12 & Text3 & "9" & "'"
      'end 2020/07/21
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
