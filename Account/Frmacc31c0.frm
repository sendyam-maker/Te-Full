VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc31c0 
   AutoRedraw      =   -1  'True
   Caption         =   "票據轉出作業"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5385
   ScaleWidth      =   8820
   Begin VB.ComboBox Combo1 
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
      Left            =   6870
      TabIndex        =   5
      Top             =   600
      Width           =   1572
   End
   Begin VB.CommandButton Command2 
      Height          =   300
      Left            =   2550
      Picture         =   "Frmacc31c0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   620
      Width           =   350
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
      Height          =   315
      Left            =   1320
      TabIndex        =   25
      Top             =   2040
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
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   1215
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
      TabIndex        =   12
      Top             =   1680
      Width           =   1572
   End
   Begin VB.TextBox Text4 
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
      TabIndex        =   11
      Top             =   2040
      Width           =   1572
   End
   Begin VB.TextBox Text6 
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
   Begin VB.TextBox Text5 
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
      Top             =   1680
      Visible         =   0   'False
      Width           =   1572
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc31c0.frx":0102
      Height          =   2300
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4075
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
      Caption         =   "票據轉出資料"
      ColumnCount     =   9
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
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "a0e34"
         Caption         =   "轉出日期"
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
      BeginProperty Column05 
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
      BeginProperty Column06 
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
      BeginProperty Column07 
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
      BeginProperty Column08 
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
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1319.811
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   13
      Top             =   1320
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   4080
      TabIndex        =   14
      Top             =   1320
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
   Begin MSMask.MaskEdBox MaskEdBox3 
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   2520
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
   Begin MSForms.TextBox Text7 
      Height          =   315
      Left            =   5670
      TabIndex        =   10
      Top             =   240
      Width           =   2775
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
   Begin MSForms.TextBox Text8 
      Height          =   315
      Left            =   5670
      TabIndex        =   9
      Top             =   2040
      Width           =   2775
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "4895;556"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text9 
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Top             =   960
      Width           =   7140
      VariousPropertyBits=   -1466941413
      MaxLength       =   35
      ScrollBars      =   2
      Size            =   "12594;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "轉出所別"
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
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label9 
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
      TabIndex        =   24
      Top             =   240
      Width           =   972
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label1 
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
      TabIndex        =   23
      Top             =   600
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
      TabIndex        =   22
      Top             =   240
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      Height          =   2412
      Left            =   240
      Top             =   120
      Width           =   8292
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "收票日期"
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
      TabIndex        =   21
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label4 
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
      TabIndex        =   20
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label5 
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
      Left            =   360
      TabIndex        =   19
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
      TabIndex        =   18
      Top             =   2040
      Width           =   972
   End
   Begin VB.Label Label7 
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
      Left            =   360
      TabIndex        =   17
      Top             =   2040
      Width           =   972
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "轉出日期"
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
      Left            =   3150
      TabIndex        =   16
      Top             =   600
      Width           =   975
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
      TabIndex        =   15
      Top             =   960
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc31c0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/07 Form2.0已修改 Text7/Text8/Text9/DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc0e0 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adocheck As New ADODB.Recordset

Private Sub Combo1_Validate(Cancel As Boolean)
    If Trim(Combo1) = MsgText(601) Then Exit Sub
    If Asc(Combo1) < 49 Or Asc(Combo1) > 51 Then
        MsgBox Label14.Caption & "輸入錯誤", , MsgText(5)
        Combo1.SetFocus
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub Command2_Click()
   'Modify by Amy 2020/07/22 +text1 不可為空及PKey 判斷
   If Adodc1.Recordset.RecordCount = 0 Or Text6 = MsgText(601) Or Text2 = MsgText(601) Or Text1 = MsgText(601) Then
      Exit Sub
   End If
   Adodc1.Recordset.Find "a0e01 = '" & Text6 & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      'Adodc1.Recordset.Find "a0e02 = '" & Text2 & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
      Adodc1.Recordset.Find "PKey = '" & Text2 & Text6 & Text1 & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
   'end 2020/07/22
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
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   FormShow
   RecordShow
End Sub

Private Sub Form_Activate()
   strFormName = Name
   SetToolBar 'Add by Amy 2014/11/14
   If strCompanyNo = MsgText(601) Then
      Exit Sub
   End If
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.MoveFirst
   End If
   Adodc1.Recordset.Find "a0e01 = '" & strCompanyNo & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      'Modify by Amy 2020/07/22 +PKey 因a0e07改為key,用於Find
      'Adodc1.Recordset.Find "a0e02 = '" & strItemNo & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
      Adodc1.Recordset.Find "PKey = '" & strItemNo & strCompanyNo & strBankAcc & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
      If Adodc1.Recordset.EOF = False Then
         FormShow
         RecordShow
      End If
   End If
   strCompanyNo = MsgText(601)
   strBankAcc = MsgText(601) 'Add by Amy 2020/07/22
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer, intY As Integer, sglWidth As Single, sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   'Modify by Amy 2023/10/11 原W8850 H5500
   Me.Width = 8940
   Me.Height = 5850
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   SetToolBar 'Add by Amy 2014/11/14
   Combo1.AddItem ComboItem(191)
   Combo1.AddItem ComboItem(192)
   Combo1.AddItem ComboItem(193)
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   MaskEdBox3.Mask = DFormat
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
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc31c0 = Nothing
End Sub

Private Sub MaskEdBox3_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub MaskEdBox3_Validate(Cancel As Boolean)
   If MaskEdBox3.Text = MsgText(601) Or MaskEdBox3.Text = MsgText(29) Then
      MsgBox Label8 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox3.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox3.Text) = MsgText(603) Then
      MsgBox Label8 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox3.SetFocus
      Exit Sub
   End If
End Sub

'Add by Amy 2020/07/22
Private Sub Text1_GotFocus()
    TextInverse Text1
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    If strSaveConfirm <> MsgText(3) Then Exit Sub
    If Text2 = MsgText(601) Or Text6 = MsgText(601) Or Text1 = MsgText(601) Then Exit Sub
    Acc0e0Query
End Sub
'end 2020/07/22

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
   If strSaveConfirm <> MsgText(3) Then
      Exit Sub
   End If
   If Text2 <> MsgText(601) Then
      adocheck.CursorLocation = adUseClient
      'Modify by Morgan 2004/10/29  加應收過濾條件 and a0e04='R'
      adocheck.Open "select a0e01, a0e02 from acc0e0 where a0e02 = '" & Text2 & "' and a0e14 = 0 and a0e15 = 0 and a0e17 = 0 and a0e21 = 0 and a0e25 = 0 and a0e04='R'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
         'Modify by Amy 2020/07/22 一筆才預帶
         If IsNull(adocheck.Fields(0).Value) = False And adocheck.RecordCount = 1 Then
            Text6 = adocheck.Fields(0).Value
            adocheck.Close
            'Acc0e0Query'Mark by Amy 2020/07/22 至 銀行帳號跳離開做
            Exit Sub
         End If
      End If
      'Mark by Amy 2020/07/22
'      MessageShow Label2
'      adocheck.Close
'      Cancel = True
'      Exit Sub
   End If
End Sub

Private Sub Text4_Change()
   If Text4 = MsgText(601) Then
      Exit Sub
   End If
   Select Case Mid(Text5, 1, 1)
      Case Mid(ComboItem(131), 1, 1)
         Text8 = CustomerQuery(Text4, 1)
      Case Mid(ComboItem(132), 1, 1)
         Text8 = A0i02Query(Text4)
      Case Mid(ComboItem(133), 1, 1)
         Text8 = StaffQuery(Text4)
      Case Else
         Text8 = MsgText(601)
   End Select
End Sub

Private Sub Text6_Change()
   If Text6 = MsgText(601) Then
      Exit Sub
   End If
   Text7 = A0g02Query(Text6)
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'Add by Amy 2020/07/22
Private Sub Text6_Validate(Cancel As Boolean)
    If strSaveConfirm <> MsgText(3) Then Exit Sub
    If Text6 = MsgText(601) Then Exit Sub
    adocheck.CursorLocation = adUseClient
    adocheck.Open "select a0e01, a0e02,a0e07 from acc0e0 where a0e02 = '" & Text2 & "' And a0e01='" & Text6 & "' and a0e14 = 0 and a0e15 = 0 and a0e17 = 0 and a0e21 = 0 and a0e25 = 0 and a0e04='R'", adoTaie, adOpenStatic, adLockReadOnly
    If adocheck.RecordCount <> 0 Then
        '一筆才預帶
         If IsNull(adocheck.Fields(0).Value) = False And adocheck.RecordCount = 1 Then
            Text1 = adocheck.Fields(2).Value
            adocheck.Close
        End If
    End If
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   'Modify by Amy 2020/07/22 +PKey 因a0e07改為key,用於Find
   adoadodc1.Open "select acc0e0.*,a0e02||a0e01||a0e07 as PKey from acc0e0 where (a0e34 is not null AND A0E34 <> 0) order by a0e34 desc, a0e01 asc,PKey asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
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
   'Modify by Amy 2020/07/22 +PKey 因a0e07改為key,用於Find
   adoadodc1.Open "select acc0e0.*,a0e02||a0e01||a0e07 as PKey from acc0e0 where (a0e34 is not null AND A0E34 <> 0) order by a0e34 desc, a0e01 asc,PKey asc, a0e02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      If Text2 <> MsgText(601) And Text6 <> MsgText(601) Then
         Adodc1.Recordset.Find "a0e01 = '" & Text6 & "'", 0, adSearchForward, 1
         If Adodc1.Recordset.EOF = False Then
            'Modify by Amy 2020/07/21 +PKey 因a0e07改為key,用於Find
            'Adodc1.Recordset.Find "a0e02 = '" & Text2 & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
            Adodc1.Recordset.Find "PKey = '" & Text2 & Text6 & Text1 & "'", 0, adSearchForward, Adodc1.Recordset.Bookmark
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
'  顯示資料表
'
'*************************************************
Public Sub FormShow()
   Text6 = Adodc1.Recordset.Fields("a0e01").Value
   If IsNull(Adodc1.Recordset.Fields("a0e07").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = Adodc1.Recordset.Fields("a0e07").Value
   End If
   Text2 = Adodc1.Recordset.Fields("a0e02").Value
   Select Case IIf(IsNull(Adodc1.Recordset.Fields("a0e45").Value), "", Adodc1.Recordset.Fields("a0e45").Value)
      Case "1"
         Combo1 = ComboItem(191)
      Case "2"
         Combo1 = ComboItem(192)
      Case "3"
         Combo1 = ComboItem(193)
      Case Else
         Combo1 = ""
   End Select
   MaskEdBox3.Mask = MsgText(601)
   If IsNull(Adodc1.Recordset.Fields("a0e34").Value) Then
      MaskEdBox3.Text = MsgText(601)
   Else
      MaskEdBox3.Text = CFDate(Adodc1.Recordset.Fields("a0e34").Value)
   End If
   MaskEdBox3.Mask = DFormat
   If IsNull(Adodc1.Recordset.Fields("a0e12").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = Adodc1.Recordset.Fields("a0e12").Value
   End If
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(Adodc1.Recordset.Fields("a0e13").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(Adodc1.Recordset.Fields("a0e13").Value)
   End If
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(Adodc1.Recordset.Fields("a0e10").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(Adodc1.Recordset.Fields("a0e10").Value)
   End If
   MaskEdBox2.Mask = DFormat
   If IsNull(Adodc1.Recordset.Fields("a0e11").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = Adodc1.Recordset.Fields("a0e11").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e05").Value) Then
      Text5 = MsgText(601)
   Else
      Select Case Adodc1.Recordset.Fields("a0e05").Value
         Case Mid(ComboItem(131), 1, 1)
            Text5 = ComboItem(131)
         Case Mid(ComboItem(132), 1, 1)
            Text5 = ComboItem(132)
         Case Mid(ComboItem(133), 1, 1)
            Text5 = ComboItem(133)
      End Select
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e08").Value) Then
      Text10 = MsgText(601)
   Else
      Select Case Adodc1.Recordset.Fields("a0e08").Value
         Case Mid(ComboItem(11), 1, 1)
            Text10 = ComboItem(11)
         Case Mid(ComboItem(12), 1, 1)
            Text10 = ComboItem(12)
         Case Mid(ComboItem(13), 1, 1)
            Text10 = ComboItem(13)
         Case Else
            Text10 = MsgText(601)
      End Select
   End If
   If IsNull(Adodc1.Recordset.Fields("a0e06").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = Adodc1.Recordset.Fields("a0e06").Value
      Select Case Adodc1.Recordset.Fields("A0E05").Value
         Case Mid(ComboItem(131), 1, 1)
            Text8 = CustomerQuery(AfterZero(Text4), 1)
         Case Mid(ComboItem(132), 1, 1)
            Text8 = A0i02Query(Text4)
         Case Mid(ComboItem(133), 1, 1)
            Text8 = StaffQuery(Text4)
         Case Else
            Text8 = MsgText(601)
      End Select
   End If
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = Adodc1.Recordset.Bookmark & MsgText(35) & Adodc1.Recordset.RecordCount
End Sub

'*************************************************
'  顯示查詢資料
'
'*************************************************
Private Sub Acc0e0Query()
On Error GoTo Checking
   adoacc0e0.CursorLocation = adUseClient
   'Modify by Amy 2020/07/22 +a0e07 因改為key
   adoacc0e0.Open "select * from acc0e0 where a0e01 = '" & Text6 & "' and a0e02 = '" & Text2 & "' And a0e07='" & Text1 & "' and (a0e25 is null or a0e25 = 0)", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0e0.RecordCount = 0 Then
      MsgBox MsgText(33) & " " & MsgText(39), , MsgText(5)
      adoacc0e0.Close
      Exit Sub
   End If
   If IsNull(adoacc0e0.Fields("a0e07").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = adoacc0e0.Fields("a0e07").Value
   End If
   MaskEdBox3.Mask = MsgText(601)
   If IsNull(adoacc0e0.Fields("a0e34").Value) Then
      MaskEdBox3.Text = MsgText(601)
   Else
      MaskEdBox3.Text = CFDate(adoacc0e0.Fields("a0e34").Value)
   End If
   MaskEdBox3.Mask = DFormat
'   If IsNull(adoacc0e0.Fields("a0e12").Value) Then
'      Text9 = MsgText(601)
'   Else
'      Text9 = adoacc0e0.Fields("a0e12").Value
'   End If
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc0e0.Fields("a0e13").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(adoacc0e0.Fields("a0e13").Value)
   End If
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(adoacc0e0.Fields("a0e10").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(adoacc0e0.Fields("a0e10").Value)
   End If
   MaskEdBox2.Mask = DFormat
   If IsNull(adoacc0e0.Fields("a0e11").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = adoacc0e0.Fields("a0e11").Value
   End If
   If IsNull(adoacc0e0.Fields("a0e05").Value) Then
      Text5 = MsgText(601)
   Else
      Select Case adoacc0e0.Fields("a0e05").Value
         Case Mid(ComboItem(131), 1, 1)
            Text5 = ComboItem(131)
         Case Mid(ComboItem(132), 1, 1)
            Text5 = ComboItem(132)
         Case Mid(ComboItem(133), 1, 1)
            Text5 = ComboItem(133)
      End Select
   End If
   If IsNull(adoacc0e0.Fields("a0e08").Value) Then
      Text10 = MsgText(601)
   Else
      Select Case adoacc0e0.Fields("a0e08").Value
         Case Mid(ComboItem(11), 1, 1)
            Text10 = ComboItem(11)
         Case Mid(ComboItem(12), 1, 1)
            Text10 = ComboItem(12)
         Case Mid(ComboItem(13), 1, 1)
            Text10 = ComboItem(13)
         Case Else
            Text10 = MsgText(601)
      End Select
   End If
   If IsNull(adoacc0e0.Fields("a0e06").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = adoacc0e0.Fields("a0e06").Value
   End If
   Text9 = MaskEdBox2.Text & " / " & Text2 & " / " & Text1 & " / " & Text7
   adoacc0e0.Close
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
   'add by nickc 2007/07/13 將輸入法改成使用API
   OpenIme
End Sub

'Modify by Amy 2021/10/07 原:KeyCode As Integer
Private Sub Text9_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   KeyDefine Val(KeyCode)
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         Acc0e0Query
   End Select
   KeyEnter KeyCode
End Sub
'add by nickc 2007/07/13 將輸入法改成使用API
Private Sub Text9_Validate(Cancel As Boolean)
CloseIme
End Sub

'Add by Amy 2014/11/05 由aacc_sav 搬回form
Public Sub Frmacc31c0_Save()
Dim adoacc0g0 As New ADODB.Recordset
Dim strAutoNo As String
Dim strPosition As String
'Add by Amy 2014/11/14
Dim strMsg As String
Dim bCancel As Boolean

   On Error GoTo Checking
   
      If Text6 = MsgText(601) Then
         MsgBox MsgText(10) & Label9, , MsgText(5)
         strControlButton = MsgText(602)
         Text6.SetFocus
         Exit Sub
      Else
         If Text2 = MsgText(601) Then
            MsgBox MsgText(10) & Label2, , MsgText(5)
            strControlButton = MsgText(602)
            Text2.SetFocus
            Exit Sub
         End If
         If ExistCheck("acc0g0", "a0g01", Text6, Label9) = False Then
            strControlButton = MsgText(602)
            Text6.SetFocus
            Exit Sub
         End If
         If MaskEdBox3.Text = MsgText(601) Or MaskEdBox3.Text = MsgText(29) Then
            MsgBox Label8 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            MaskEdBox3.SetFocus
            Exit Sub
         Else
            If DateCheck(MaskEdBox3.Text) = MsgText(603) Then
               MsgBox Label8 & MsgText(63), , MsgText(5)
               strControlButton = MsgText(602)
               MaskEdBox3.SetFocus
               Exit Sub
            End If
            'Add byAmy 2014/11/11 +系統日檢查
            If MaskEdBox3.Enabled = True Then
                If ChkWorkData("1", DBDATE(MaskEdBox3), strMsg) = False Then
                    MsgBox Label8 & strMsg, , MsgText(5)
                    strControlButton = MsgText(602)
                    MaskEdBox3.SetFocus
                    Exit Sub
                End If
            End If
            'end 2014/11/11
         End If
         'Add by Amy 2014/11/14 轉出所別沒輸Insert Acc1p0會error
         If Trim(Combo1) = MsgText(601) Then
            MsgBox Label14 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            Combo1.SetFocus
            Exit Sub
         End If
         Call Combo1_Validate(bCancel)
         If bCancel = True Then
            strControlButton = MsgText(602)
            Exit Sub
         End If
         'end 2014/11/14
      End If
      'Add by Amy 2021/10/07 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
      If PUB_ChkUniText(Me) = False Then
         strControlButton = MsgText(602)
         Exit Sub
      End If
      
      adoTaie.BeginTrans
      adoacc0e0.CursorLocation = adUseClient
      'Modify by Morgan 2004/10/29 加應收過濾條件 and a0e04='R'
      'Modify by Amy 2020/07/22 +a0e07 因改為key
      adoacc0e0.Open "select * from acc0e0 where a0e01 = '" & Text6 & "' and a0e02 = '" & Text2 & "' And a0e07='" & Text1 & "' and a0e14 = 0 and a0e15 = 0 and a0e17 = 0 and a0e21 = 0 and a0e25 = 0 and a0e04='R'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If adoacc0e0.RecordCount = 0 Then
         adoTaie.RollbackTrans
         Exit Sub
      End If
      If IsNull(adoacc0e0.Fields("a0e34").Value) = False And adoacc0e0.Fields("a0e34").Value <> 0 Then
         MsgBox MsgText(9), , MsgText(5)
         adoTaie.RollbackTrans
         adoacc0e0.Close
        Exit Sub
      End If
      If Text1 <> MsgText(601) Then
         adoacc0e0.Fields("a0e07").Value = Text1
      Else
         adoacc0e0.Fields("a0e07").Value = Null
      End If
      If MaskEdBox3.Text <> MsgText(601) And MaskEdBox3.Text <> MsgText(29) Then
         adoacc0e0.Fields("a0e34").Value = Val(FCDate(MaskEdBox3.Text))
         adoacc0e0.Fields("a0e21").Value = Val(FCDate(MaskEdBox3.Text))
      Else
         adoacc0e0.Fields("a0e34").Value = 0
         adoacc0e0.Fields("a0e21").Value = 0
      End If
      If Combo1 <> MsgText(601) Then
         adoacc0e0.Fields("a0e45").Value = Mid(Combo1, 1, 1)
      Else
         adoacc0e0.Fields("a0e45").Value = Null
      End If
      If Text9 <> MsgText(601) Then
         adoacc0e0.Fields("a0e12").Value = Text9
      Else
         adoacc0e0.Fields("a0e12").Value = Null
      End If
      Select Case Mid(Combo1, 1, 1)
         Case "1"
            strPosition = "1911"
         Case "2"
            strPosition = "1912"
         Case "3"
            strPosition = "1913"
         Case Else
      End Select
      adoacc0e0.Fields("a0e29").Value = Val(strSrvDate(2))
      adoacc0e0.Fields("a0e30").Value = ServerTime
      adoacc0e0.Fields("a0e31").Value = strUserNum
      'Modify by Amy 2020/07/22 +a0e07 因改為key
      strAutoNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'L' and a1p04 = '" & Text2 & Text6 & Text1 & "7" & "'", 3)
      adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values " & _
                    "('1', 'L', '" & strAutoNo & "', '" & Text2 & Text6 & Text1 & "7" & "', '" & strPosition & "', '" & MsgText(55) & "', " & adoacc0e0.Fields("a0e11").Value & ", 0, '" & Text2 & "', '" & Text6 & "', '" & Text1 & "', " & _
                      "" & Val(adoacc0e0.Fields("a0e10").Value) & ", '" & adoacc0e0.Fields("a0e08").Value & "', '" & adoacc0e0.Fields("a0e12").Value & "', null, null, null, " & Val(FCDate(MaskEdBox3.Text)) & ", null, null, null, null, " & _
                      "'" & adoacc0e0.Fields("a0e03").Value & "', null, null, null, null)"
      strAutoNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'L' and a1p04 = '" & Text2 & Text6 & Text1 & "7" & "'", 3)
      adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values " & _
                    "('1', 'L', '" & strAutoNo & "', '" & Text2 & Text6 & Text1 & "7" & "', '113001', '" & MsgText(55) & "', 0, " & adoacc0e0.Fields("a0e11").Value & ", '" & Text2 & "', '" & Text6 & "', '" & Text1 & "', " & _
                      "" & Val(adoacc0e0.Fields("a0e10").Value) & ", '" & adoacc0e0.Fields("a0e08").Value & "', '" & adoacc0e0.Fields("a0e12").Value & "', null, null, null, " & Val(FCDate(MaskEdBox3.Text)) & ", null, null, null, null, " & _
                      "'" & adoacc0e0.Fields("a0e03").Value & "', null, null, null, null)"
      'end 2020/07/22
      adoacc0e0.UpdateBatch
      adoTaie.CommitTrans
      AdodcRefresh
      adoacc0e0.Close
      RecordShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   adoTaie.RollbackTrans
   MsgBox Err.Description, , MsgText(5)
End Sub

'Add by Amy 2014/11/14 鎖toolbar
'新增Save時會將轉出日寫入a0e21, 而修改Save會判斷a0e21有值會跳離不會存,所以先將修改鎖住
'刪除時若已有傳票仍會刪除Acc1p0,且a0e21不會被更新,所以先將刪除鎖住
Public Sub SetToolBar()
    With Forms(0)
      .Toolbar1.Buttons.Item(5).Enabled = False
      .Toolbar1.Buttons.Item(8).Enabled = False
   End With
End Sub

'Add by Amy 2020/07/21
'從acc_cls搬回
Public Sub Frmacc31c0_Clear()
   With Frmacc31c0
      .Text6 = ""
      .Combo1 = ""
      .Text7 = ""
      .Text1 = ""
      .Text2 = ""
      .MaskEdBox3.Mask = ""
      .MaskEdBox3.Text = ""
      .MaskEdBox3.Mask = DFormat
      .Text9 = ""
      .Text10 = ""
      .MaskEdBox1.Mask = ""
      .MaskEdBox1.Text = ""
      .MaskEdBox1.Mask = DFormat
      .MaskEdBox2.Mask = ""
      .MaskEdBox2.Text = ""
      .MaskEdBox2.Mask = DFormat
      .Text3 = ""
      .Text5 = ""
      .Text4 = ""
      .Text8 = ""
      .Text2.SetFocus
   End With
End Sub

'從acc_del搬回
Public Sub Frmacc31c0_Delete()
On Error GoTo Checking
   With Frmacc31c0
      'Modify by Amy 2020/07/22 +a0e07 因改為key
      If DeleteCheck("select a0e01 from acc0e0 where a0e01 = '" & .Text6 & "' and a0e02 = '" & .Text2 & "' And a0e07='" & Text1 & "' ") = MsgText(603) Then
         Exit Sub
      End If
      adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'L' and a1p04 = '" & .Text2 & .Text6 & Text1 & "7" & "' and a1p05 = '1911'"
      adoTaie.Execute "update acc0e0 set a0e34 = '' where a0e01 = '" & .Text6 & "' and a0e02 = '" & .Text2 & "' And a0e07='" & Text1 & "' "
      'end 2020/07/22
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
