VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc2160 
   AutoRedraw      =   -1  'True
   Caption         =   "抵帳單輸入"
   ClientHeight    =   5328
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5328
   ScaleWidth      =   8760
   Begin VB.CommandButton Command2 
      Caption         =   "電子檔"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3555
      TabIndex        =   31
      Top             =   189
      Width           =   855
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1275
      TabIndex        =   9
      Top             =   4635
      Width           =   480
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6765
      MaxLength       =   14
      TabIndex        =   14
      Top             =   4635
      Width           =   1572
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1740
      TabIndex        =   10
      Top             =   4635
      Width           =   780
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2520
      TabIndex        =   11
      Top             =   4635
      Width           =   240
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2775
      TabIndex        =   12
      Top             =   4635
      Width           =   348
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3120
      TabIndex        =   13
      Top             =   4635
      Visible         =   0   'False
      Width           =   396
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1560
      TabIndex        =   5
      Top             =   1260
      Width           =   1596
   End
   Begin VB.TextBox Text8 
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
      Height          =   330
      Left            =   3144
      TabIndex        =   27
      Top             =   4110
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   3120
      Picture         =   "Frmacc2160.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   165
      Width           =   350
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
      Height          =   450
      Left            =   6840
      Picture         =   "Frmacc2160.frx":0102
      Style           =   1  '圖片外觀
      TabIndex        =   15
      ToolTipText     =   "取消"
      Top             =   4080
      Width           =   450
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4200
      TabIndex        =   6
      Top             =   1267
      Width           =   1332
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1920
      MaxLength       =   15
      TabIndex        =   3
      Top             =   890
      Width           =   3612
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   2
      Top             =   520
      Width           =   1572
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6840
      MaxLength       =   15
      TabIndex        =   8
      Top             =   150
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      MaxLength       =   15
      TabIndex        =   0
      Top             =   150
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   330
      Left            =   6840
      TabIndex        =   4
      Top             =   890
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   593
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
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
      Bindings        =   "Frmacc2160.frx":076C
      Height          =   1860
      Left            =   240
      TabIndex        =   16
      Top             =   2130
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   3281
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   17
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
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "axg02"
         Caption         =   "總收文號"
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
         DataField       =   "axg03"
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
         DataField       =   "axg04"
         Caption         =   "抵帳金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "axg12"
         Caption         =   "案件名稱"
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
         DataField       =   "axg13"
         Caption         =   "收據抬頭"
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
         Size            =   284
         BeginProperty Column00 
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1404.284
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1307.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3767.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   4500.284
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   1920
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
      Height          =   330
      Left            =   6840
      TabIndex        =   26
      Top             =   1267
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   572
      _Version        =   393216
      BackColor       =   14737632
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSForms.TextBox Text9 
      Height          =   405
      Left            =   1560
      TabIndex        =   7
      Top             =   1650
      Width           =   6915
      VariousPropertyBits=   -1467989989
      ScrollBars      =   2
      Size            =   "12197;714"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text3 
      Height          =   330
      Left            =   3150
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   520
      Width           =   5295
      VariousPropertyBits=   671105049
      BackColor       =   16777215
      MaxLength       =   50
      Size            =   "9340;582"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "本所案號"
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
      Left            =   300
      TabIndex        =   30
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "抵帳單金額"
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
      Left            =   5565
      TabIndex        =   29
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "合計"
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
      Left            =   2310
      TabIndex        =   28
      Top             =   4148
      Width           =   855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "抵帳日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5640
      TabIndex        =   25
      Top             =   1306
      Width           =   1212
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      Height          =   495
      Left            =   240
      Top             =   4575
      Width           =   8295
   End
   Begin VB.Image Image2 
      Height          =   132
      Left            =   0
      Top             =   4752
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4752
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "備註"
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
      Left            =   360
      TabIndex        =   24
      Top             =   1650
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "抵帳總額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3240
      TabIndex        =   23
      Top             =   1306
      Width           =   972
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "抵帳幣別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   22
      Top             =   1306
      Width           =   972
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "抵帳單日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5640
      TabIndex        =   21
      Top             =   929
      Width           =   1212
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "代理人C/N No."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   20
      Top             =   929
      Width           =   1572
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "代理人"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   19
      Top             =   559
      Width           =   852
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "原帳單編號"
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
      Left            =   5625
      TabIndex        =   18
      Top             =   188
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "抵帳單編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   17
      Top             =   189
      Width           =   1212
   End
End
Attribute VB_Name = "Frmacc2160"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/07 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Text3、Text9
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoacc160 As New ADODB.Recordset
Public adoacc150 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adocaseprogress As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public strDocNo As String
Dim RQstr As String '   'Add by Lydia 2014/10/31
'Add By Sindy 2018/2/22
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
Public m_RDate As String
Dim m_Done As Boolean
Public m_PrevForm As Form  '前一畫面
'2018/2/22 END


Private Sub Combo1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
   If Combo1 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc1y0", "a1y01", Combo1, Label6) = False Then
      Cancel = True
      Combo1.SetFocus
   End If
End Sub

Private Sub Command1_Click()
   AdodcDelete
   AdodcClear
End Sub

Private Sub Command2_Click()
   If Text2 <> "" Then
      strExc(0) = "select a1601 from acc160 where a1601='" & Text2 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strItemNo = Text2
         Frmacc2154.Show vbModal
         strItemNo = ""
         strFormName = Me.Name
      Else
         MsgBox "抵帳單不存在！", vbCritical
      End If
   Else
      MsgBox "請先輸入抵帳單編號！", vbExclamation
   End If
End Sub

Private Sub Command3_Click()
'   If adoacc160.RecordCount = 0 Or Text2 = MsgText(601) Then
'      Exit Sub
'   End If
'   adoacc160.Find "a1601 = '" & Text2 & "'", 0, adSearchForward, 1

   'Add by Lydia 2014/10/31 改為frmacc2150的方式
   Acc160Refresh
   AdodcClear
   If adoacc160.EOF = False Then
      FormShow
      AdodcRefresh
      RecordShow
   Else
      If FMP2open = True Then
        MsgBox "權限不足或查無符合資料 !", vbInformation
      Else
        MsgBox MsgText(33), , MsgText(5)
      End If
      If adoacc160.EOF <> adoacc160.BOF Then adoacc160.MoveFirst
   End If
End Sub

Private Sub Command3_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command3_Click
         Exit Sub
   End Select
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   AdodcShow
End Sub

Private Sub Form_Activate()
'edit by nickc 2007/02/08
'   '93.3.16 ADD BY SONIA
'   If IsObject(mdiMain) Then
'      mdiMain.toolshow
'   End If
'   '93.3.16 END
   
   Dim formCnt As Integer
   For formCnt = 0 To Forms.Count - 1
       If UCase(Forms(formCnt).Name) = "MDIMAIN" Then
             Forms(formCnt).ToolShow
             Exit For
       End If
   Next
   strFormName = Name
   
   'Added by Sindy 2018/2/22
   If m_strIR01 <> "" And m_Done = False Then
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "＜" & m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 & "＞）"
      'KeyEnter vbKeyF2 'Set新增狀態 : 使用者自行決定啟動何功能,因新增則自動取號了,易生多空號
   End If
   '2018/2/22 END
   
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      If strItemNo = MsgText(601) Then
         Exit Sub
      End If
      'Modified by Morgan 2018/3/2
      'If adoacc160.RecordCount <> 0 Then
      '   adoacc160.MoveFirst
      'End If
      'adoacc160.Find "a1601 = '" & strItemNo & "'", 0, adSearchForward, 1
      'If adoacc160.EOF = False Then
      Text2 = strItemNo
      Acc160Refresh
      If adoacc160.RecordCount <> 0 Then
      'end 2018/3/2
         FormShow
         AdodcRefresh
         RecordShow
      End If
      strItemNo = MsgText(601)
      Exit Sub
   End If
   If strCon9 <> "" Then
      If strControlButton <> MsgText(602) Then
         adoTaie.Execute strCon9
'         adoTaie.Execute strCon10
      End If
      If strControlButton <> MsgText(602) Then
         AdodcRefresh
         AdodcClear
         Text14.SetFocus
      End If
      If strCustNo <> MsgText(601) Then
         Text1 = strCustNo
         Text3 = FagentQuery(Text1, 2)
         If Text3 = MsgText(601) Then
            Text3 = FagentQuery(Text1, 1)
         End If
      End If
      Frmacc2160_Save
      strControlButton = MsgText(601)
   End If
End Sub

'Added by Lydia 2021/12/07
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Added by Lydia 2021/12/07 Form2.0 記錄鍵盤傳入順序
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   'Modified by Lydia 2021/12/07 改成模組
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   Me.Width = 8850
'   Me.Height = 5500
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
   'Modify by Amy 2023/08/18 H5550
   PUB_InitForm Me, 8850, 5770, strBackPicPath1
   'end 2021/12/07
   
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
    FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
   OpenTable
   If adoacc160.RecordCount <> 0 Then
      adoacc160.MoveLast
      adoacc160.MoveFirst
      RecordShow
   End If
   FormDisabled
   
   'Added by Morgan 2016/8/26
   'Add By Sindy 2021/1/18 + Or Pub_StrUserSt03 = "P22"
   'Removed by Morgan 2023/11/10 開放都可上傳
   'If Pub_StrUserSt03 = "P12" Or Pub_StrUserSt03 = "M51" Or _
   '   Pub_StrUserSt03 = "M31" Or Pub_StrUserSt03 = "P22" Then
   '   Command2.Visible = True
   'Else
   '   Command2.Visible = False
   'End If
   'end 2023/11/10
   'end 2016/8/26
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   strTrackMode = "" 'Added by Lydia 2021/12/07 Form2.0 記錄鍵盤傳入順序(清除)
   
   KeyEnter vbKeyEscape
   MenuEnabled
   
   'Add By Sindy 2018/2/23
   If m_strIR01 <> "" Then
      If Not m_PrevForm Is Nothing Then
         Call m_PrevForm.GoNext
      End If
   End If
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   '2018/2/23 END
   
   Set Frmacc2160 = Nothing
End Sub

Private Sub MaskEdBox1_LostFocus()
   'Add by Morgan 2007/5/18
   '檢查抵帳單資料是否重覆
   'Modify By Sindy 2009/06/17 若為專利處只須以代理人+代理人D/N No.做重覆檢核
'   If Left(Trim(GetStaffDepartment(strUserNum)), 2) <> "P1" Then
'      If PUB_ChkDNDup(MaskEdBox1.Text, Text1.Text, Text5.Text, Text2.Text, , 1) = True Then
'         Text5.SetFocus
'      End If
'   End If
   'end 2007/5/18
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox Label5 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label5 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If strSaveConfirm <> MsgText(3) Then
      Exit Sub
   End If
   If Mid(MaskEdBox1.Text, 1, 3) <> Mid(CFDate(ACDate(ServerDate)), 1, 3) Then
      Text2 = UpdateNo("acc160", "a1601", 5, MaskEdBox1.Text, MsgText(813))
   Else
      'Text2 = AutoNo(MsgText(813), 5)
      Text2 = strDocNo
   End If

End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
   Select Case Len(Text1)
      Case 6
         Text1 = Text1 & "000"
      Case 8
         Text1 = Text1 & "0"
   End Select
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
    If FMP2open = True Then
       If InStr(1, FMP2openSQL, Trim(Text1)) = 0 Then  '限定特定代理人
          MsgBox "權限不足 !", vbInformation
          Cancel = True
          Text1.SetFocus
          TextInverse Text1
          Exit Sub
       End If
    End If
   If ExistCheck("fagent", "fa01", Mid(Text1, 1, 8), Label3) = False Then
      Cancel = True
      Text1.SetFocus
      TextInverse Text1
      Exit Sub
   End If
   Text3 = FagentQuery(Text1, 2)
   If Text3 = MsgText(601) Then
      Text3 = FagentQuery(Text1, 1)
   End If
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   'Add By Sindy 2012/6/5 改為先抓該代理人是否有設定帳單幣別,若有,則抓代理人的帳單幣別,若沒有才抓NA52
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select fa113 from fagent where fa01='" & Mid(Text1, 1, 8) & "' and fa02='" & Mid(Text1, 9, 1) & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields("fa113").Value) = False Then
         Combo1 = adoquery.Fields("fa113").Value
         adoquery.Close
         Exit Sub
      End If
   End If
   adoquery.Close
   '2012/6/5 End
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select na52 from fagent, nation where fa10 = na01 (+) and fa01 = '" & Mid(Text1, 1, 8) & "' and fa02 = '" & Mid(Text1, 9, 1) & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields("na52").Value) = False Then
         Combo1 = adoquery.Fields("na52").Value
      End If
   End If
   adoquery.Close
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
   Select Case Text14
      Case "TF"
         Text6 = "0"
         Text12 = "00"
      Case Else
         Text10 = "0"
         Text6 = "00"
   End Select
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text13_GotFocus()
   TextInverse Text13
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text14_GotFocus()
   TextInverse Text14
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text14_Validate(Cancel As Boolean)
   If Text14 = "TF" Then
      Text12.Visible = True
   Else
      Text12.Visible = False
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking

   
   adoacc160.CursorLocation = adUseClient
   'adoacc160.Open "select * from acc160 order by a1601 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   Dim midSql As String
   midSql = " select m0.* from acc160 m0 where a1601>='" & Text2 & "' "
   If FMP2open = True Then
      RQstr = " select m1.axg01 from acc161 m1,caseprogress f0 where m0.a1601=m1.axg01(+) and m1.axg02=f0.cp09(+) " & FMP2openSQL
      midSql = midSql & " and a1601 in (" & RQstr & ") "
   End If
   midSql = midSql & " order by 1 asc "
   adoacc160.Open midSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc161 where axg01 = '" & Text2 & "' order by axg02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from acc1y0 order by a1y01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoquery.EOF = False
      Combo1.AddItem adoquery.Fields("a1y01").Value
      adoquery.MoveNext
   Loop
   adoquery.Close
   Combo1 = "USD"
Checking:
   If Err.NUMBER = 0 Then
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
   adoadodc1.Open "select * from acc161 where axg01 = '" & Text2 & "' order by axg02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.ReQuery
   SumShow
Checking:
   If Err.NUMBER = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表
'
'*************************************************
Public Sub FormShow()
   Text2 = adoacc160.Fields("a1601").Value
   If IsNull(adoacc160.Fields("a1603").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = adoacc160.Fields("a1603").Value
   End If
'   If Len(Text1) = 6 Then
'      Text3 = FagentQuery(AfterZero(Text1), 2)
'   Else
'      Text3 = FagentQuery(Text1, 2)
'      '2012/8/9 add by sonia
'      If Text3 = "" Then
'         Text3 = FagentQuery(Text1, 1)
'      End If
'      '2012/8/9 end
'   End If
'Add by Lydia 2014/11/13 改變讀取代理人名稱的方式
   If ClsPDGetAgent(Text1, strExc(0)) = True Then
      Text3 = strExc(0)
   Else
      Text3 = ""
   End If
   
   If IsNull(adoacc160.Fields("a1604").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = adoacc160.Fields("a1604").Value
   End If
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc160.Fields("a1602").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(adoacc160.Fields("a1602").Value)
   End If
   MaskEdBox1.Mask = DFormat
   If IsNull(adoacc160.Fields("a1605").Value) Then
      Combo1 = MsgText(601)
   Else
      Combo1 = adoacc160.Fields("a1605").Value
   End If
   If IsNull(adoacc160.Fields("a1606").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = Format(adoacc160.Fields("a1606").Value, FAmount)
   End If
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(adoacc160.Fields("a1607").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(adoacc160.Fields("a1607").Value)
   End If
   MaskEdBox2.Mask = DFormat
   If IsNull(adoacc160.Fields("a1608").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = adoacc160.Fields("a1608").Value
   End If
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from acc170 where a1702 = '" & Text2 & "' and a1709 is not null", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      tool15_enabled
   Else
      tool1_enabled
   End If
   adoquery.Close
End Sub

'*************************************************
'  顯示 Adodc 之資料
'
'*************************************************
Public Sub AdodcShow()
   Text14 = Mid(Adodc1.Recordset.Fields("axg03").Value, 1, Len(Adodc1.Recordset.Fields("axg03").Value) - 9)
   Select Case Text10
      Case "TF"
      Case Else
         Text11 = Mid(Adodc1.Recordset.Fields("axg03").Value, Len(Adodc1.Recordset.Fields("axg03").Value) - 8, 6)
         Text10 = Mid(Adodc1.Recordset.Fields("axg03").Value, Len(Adodc1.Recordset.Fields("axg03").Value) - 2, 1)
         Text6 = Mid(Adodc1.Recordset.Fields("axg03").Value, Len(Adodc1.Recordset.Fields("axg03").Value) - 1, 2)
   End Select
   If IsNull(Adodc1.Recordset.Fields("axg04").Value) Then
      Text13 = MsgText(601)
   Else
      Text13 = Format(Adodc1.Recordset.Fields("axg04").Value, FAmount)
   End If
End Sub

'*************************************************
'  重新整理抵帳單資料
'
'*************************************************
Public Sub Acc160Refresh()
On Error GoTo Checking
   If adoacc160.State = adStateOpen Then
      adoacc160.Close
   End If
   adoacc160.CursorLocation = adUseClient
   adoacc160.MaxRecords = intMax
   'adoacc160.Open "select * from acc160 where a1601 >= '" & Text2 & "' order by a1601 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   Dim midSql As String
   midSql = " select m0.* from acc160 m0 where a1601>='" & Text2 & "' "
   If FMP2open = True Then
      RQstr = " select m1.axg01 from acc161 m1,caseprogress f0 where m1.axg01=m0.a1601 and f0.cp09=m1.axg02 " & FMP2openSQL
      midSql = midSql & " and a1601 in (" & RQstr & ") "
   End If
   midSql = midSql & " order by 1 asc "
   adoacc160.Open midSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   
Checking:
   If Err.NUMBER = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  清除顯示資料
'
'*************************************************
Public Sub AdodcClear()
   Text14 = ""
   Text11 = ""
   Text10 = ""
   Text6 = ""
   Text12 = ""
   Text13 = ""
End Sub
'edit by nickc 2007/02/08 早就沒用了
''*************************************************
''  儲存資料表(國外抵帳單資料(交易檔))
''
''*************************************************
'Private Sub Acc161Save()
'On Error GoTo Checking
'      If Text14 = MsgText(601) Then
'         MsgBox MsgText(10) & Label11, , MsgText(5)
'         strControlButton = MsgText(602)
'         Text14.SetFocus
'         Exit Sub
'      End If
'      If Adodc1.Recordset.RecordCount <> 0 Then
'         Adodc1.Recordset.Find "axg02 = '" & Text15 & "'", 0, adSearchForward, 1
'         If Adodc1.Recordset.EOF Then
'            Adodc1.Recordset.AddNew
'         End If
'      Else
'         Adodc1.Recordset.AddNew
'      End If
'      Adodc1.Recordset.Fields("axg01").Value = Text2
'      Adodc1.Recordset.Fields("axg02").Value = Text15
'      If Text14 <> MsgText(601) Then
'         Adodc1.Recordset.Fields("axg03").Value = Text14
'      Else
'         Adodc1.Recordset.Fields("axg03").Value = Null
'      End If
'      If Text10 <> MsgText(610) Then
'         Adodc1.Recordset.Fields("axg12").Value = Text10
'      Else
'         Adodc1.Recordset.Fields("axg12").Value = Null
'      End If
'      If Text11 <> MsgText(601) Then
'         Adodc1.Recordset.Fields("axg04").Value = Val(Text11)
'      Else
'         Adodc1.Recordset.Fields("axg04").Value = 0
'      End If
'      If Text12 <> MsgText(601) Then
'         Adodc1.Recordset.Fields("axg05").Value = Text12
'      Else
'         Adodc1.Recordset.Fields("axg05").Value = Null
'      End If
'      adoquery.CursorLocation = adUseClient
'      adoquery.Open "select a0k04 from caseprogress, acc0k0 where cp60 = a0k01 (+) and cp09 = '" & Text15 & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adoquery.RecordCount <> 0 Then
'         If IsNull(adoquery.Fields("a0k04").Value) Then
'            Adodc1.Recordset.Fields("axg13").Value = Null
'         Else
'            Adodc1.Recordset.Fields("axg13").Value = adoquery.Fields("a0k04").Value
'         End If
'      End If
'      adoquery.Close
'      Adodc1.Recordset.Fields("axg06").Value = Val(ACDate(ServerDate))
'      Adodc1.Recordset.Fields("axg07").Value = ServerTime
'      Adodc1.Recordset.Fields("axg08").Value = strUserNum
'      Adodc1.Recordset.UpdateBatch
'      AdodcRefresh
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
'End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)

   Call PUB_SaveTrackMode(1, KeyCode)  'Added by Lydia 2021/12/07 Form2.0 記錄鍵盤傳入順序
   
   Select Case KeyCode
      Case vbKeyInsert
         'Added by Lydia 2021/12/07 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
         If PUB_ChkTrackMode = False Then
             Exit Sub
         End If
         'end 2021/12/07
         
         If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
            Exit Sub
         End If
'         'Add by Morgan 2007/5/18
'         '檢查抵帳單資料是否重覆
'         'Modify By Sindy 2009/06/17 若為專利處只須以代理人+代理人D/N No.做重覆檢核
'         If Left(Trim(GetStaffDepartment(strUserNum)), 2) = "P1" Then
'            If PUB_ChkDNDup("", Text1.Text, Text5.Text, Text2.Text, , 1) = True Then
'               Text5.SetFocus
'               Exit Sub
'            End If
'         Else
'            If PUB_ChkDNDup(MaskEdBox1.Text, Text1.Text, Text5.Text, Text2.Text, , 1) = True Then
'               Text5.SetFocus
'               Exit Sub
'            End If
'         End If
'         'end 2007/5/18
         
        'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
         If FMP2open = True Then
            If PUB_FMPtoCheck(0, 1, Pub_strUserST05, Text14, Text11, Text10, Text6) = False Then
              Text11.SetFocus
              Exit Sub
            End If
         End If
              
         If Val(Text13) = 0 Then
            MsgBox MsgText(162) & Label12, , MsgText(5)
            Text13.SetFocus
            Exit Sub
         End If
         If adoquery.State = adStateOpen Then
            adoquery.Close
         End If
         adoquery.CursorLocation = adUseClient
         Select Case Text14
            Case "TF"
               adoquery.Open "select cp09 from caseprogress where cp01 = '" & Text14 & "' and cp02 = '" & Text11 & Text10 & "' and cp03 = '" & Text6 & "' and cp04 = '" & Text12 & "'", adoTaie, adOpenStatic, adLockReadOnly
            Case Else
               adoquery.Open "select cp09 from caseprogress where cp01 = '" & Text14 & "' and cp02 = '" & Text11 & "' and cp03 = '" & Text10 & "' and cp04 = '" & Text6 & "'", adoTaie, adOpenStatic, adLockReadOnly
         End Select
         If adoquery.RecordCount = 0 Then
            MsgBox MsgText(188) & Label11, , MsgText(5)
            Text11.SetFocus
            adoquery.Close
            Exit Sub
         End If
         adoquery.Close
         If Text1 <> MsgText(601) Then
            strCustNo = Text1
         Else
            strCustNo = ""
         End If
         If Text3 <> MsgText(601) Then
            strCon1 = Text3
         Else
            strCon1 = ""
         End If
         If Text14 <> MsgText(601) Then
            strCon2 = Text14
         Else
            strCon2 = ""
         End If
         If Text11 <> MsgText(601) Then
            strCon3 = Text11
         Else
            strCon3 = ""
         End If
         If Text10 <> MsgText(601) Then
            strCon4 = Text10
         Else
            strCon4 = ""
         End If
         If Text6 <> MsgText(601) Then
            strCon5 = Text6
         Else
            strCon5 = ""
         End If
         If Text12 <> MsgText(601) Then
            strCon6 = Text12
         Else
            strCon6 = ""
         End If
         If Text13 <> MsgText(601) Then
            strCon7 = Text13
         Else
            strCon7 = ""
         End If
         If Text2 <> MsgText(601) Then
            strCon8 = Text2
         Else
            strCon8 = ""
         End If
         tool3_enabled
         Screen.MousePointer = vbHourglass
         Frmacc2162.Show
         Screen.MousePointer = vbDefault
         Me.Hide
'         Frmacc2160_Save
'         If strControlButton <> MsgText(602) Then
'            Acc161Save
'         End If
'         If strControlButton <> MsgText(602) Then
'            AdodcClear
'            Text15.SetFocus
'         End If
'         strControlButton = MsgText(601)
   End Select
   KeyEnter KeyCode
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
   adoTaie.Execute "delete from acc161 where axg01 = '" & Text2 & "' and axg02 = '" & Adodc1.Recordset.Fields("axg02").Value & "'"
   'Adodc1.Recordset.Delete
   'Adodc1.Recordset.UpdateBatch
   AdodcRefresh
Checking:
   If Err.NUMBER = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
   If Text4 = "" Then
      Exit Sub
   End If

   If FMP2open = False Then
        If ExistCheck("acc150", "a1501", Text4, Label2) = False Then
           Cancel = True
           Exit Sub
        End If
   Else
        'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
        strExc(0) = " select m0.* from acc150 m0 where a1501='" & Text4 & "' "
        If FMP2open = True Then
           strExc(1) = " select m1.axf01 from acc151 m1,caseprogress f0  where m0.a1501=m1.axf01(+) and m1.axf02=f0.cp09(+) " & FMP2openSQL
           strExc(0) = strExc(0) & " and a1501 in (" & strExc(1) & ") "
        End If
        strExc(0) = strExc(0) & " order by 1 asc "
   
        If PUB_FMPtoCheck(0, 1, Pub_strUserST05, "CHANGE_SQL", strExc(0)) = False Then
           Cancel = True
           Exit Sub
        End If
   End If
   Acc150Query
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

''Add By Sindy 2009/06/17
'Private Sub Text5_LostFocus()
'   '檢查抵帳單資料是否重覆
'   '若為專利處只須以代理人+代理人D/N No.做重覆檢核
'   If Left(Trim(GetStaffDepartment(strUserNum)), 2) = "P1" Then
'      If PUB_ChkDNDup("", Text1.Text, Text5.Text, Text2.Text, , 1) = True Then
'         Text5.SetFocus
'      End If
'   End If
'End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
   'edit by nickc 2007/06/11  切換輸入法改用API
   OpenIme
End Sub

'*************************************************
'  顯示查詢資料
'
'*************************************************
Private Sub Acc150Query()
   adoacc150.CursorLocation = adUseClient
   adoacc150.Open "select * from acc150 where a1501 = '" & Text4 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc150.RecordCount <> 0 Then
      If IsNull(adoacc150.Fields("a1503").Value) Then
         Text1 = MsgText(601)
      Else
         Text1 = adoacc150.Fields("a1503").Value
      End If
   Else
      Text1 = MsgText(601)
   End If
   adoacc150.Close
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   If adoacc160.RecordCount = 0 Then
      Exit Sub
   End If
   CountShow adoacc160.Bookmark, adoacc160.RecordCount
End Sub

'*************************************************
'  關閉分錄欄位輸入狀態
'
'*************************************************
Public Sub FormDisabled()
   Text2.Enabled = True
   Text14.Enabled = False
   Text11.Enabled = False
   Text10.Enabled = False
   Text6.Enabled = False
   Text12.Enabled = False
   Text13.Enabled = False
   Command1.Enabled = False
End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'
'*************************************************
Public Sub FormEnabled()
   Text2.Enabled = False
   Text14.Enabled = True
   Text11.Enabled = True
   Text10.Enabled = True
   Text6.Enabled = True
   Text12.Enabled = True
   Text13.Enabled = True
   Command1.Enabled = True
End Sub

'*************************************************
'  計算並顯示合計
'
'*************************************************
Public Sub SumShow()
   If adoaccsum.State = adStateOpen Then
      adoaccsum.Close
   End If
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(axg04) from acc161 where axg01 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text8 = MsgText(601)
      Else
         Text8 = Format(adoaccsum.Fields(0).Value, FAmount)
      End If
   Else
      Text8 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

Private Sub Text9_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
CloseIme
End Sub
