VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc1250 
   AutoRedraw      =   -1  'True
   Caption         =   "收款單號查詢"
   ClientHeight    =   5172
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9528
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5172
   ScaleWidth      =   9528
   Begin VB.TextBox Text12 
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
      Height          =   315
      Left            =   6450
      TabIndex        =   19
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text11 
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
      Height          =   300
      Left            =   4080
      TabIndex        =   18
      Top             =   1680
      Width           =   1572
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc1250.frx":0000
      Height          =   2505
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   9150
      _ExtentX        =   16150
      _ExtentY        =   4424
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "收款單號查詢"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "a0k02"
         Caption         =   "收據日期"
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
      BeginProperty Column01 
         DataField       =   "a0k01"
         Caption         =   "收據編號"
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
         DataField       =   "CaseNo"
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
      BeginProperty Column03 
         DataField       =   "cp09"
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
      BeginProperty Column04 
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
      BeginProperty Column05 
         DataField       =   "cp10N"
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
      BeginProperty Column06 
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
      BeginProperty Column07 
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
      BeginProperty Column08 
         DataField       =   "NAmount"
         Caption         =   "未收金額"
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
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1235.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   912.189
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1128.189
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   924.095
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   1080
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   2280
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
   Begin VB.CommandButton Command2 
      Caption         =   "收款情形"
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
      Left            =   7200
      TabIndex        =   15
      Top             =   240
      Width           =   1212
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   1572
   End
   Begin VB.TextBox Text2 
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
      Height          =   315
      Left            =   1320
      TabIndex        =   7
      Top             =   600
      Width           =   852
   End
   Begin VB.TextBox Text6 
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
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Top             =   960
      Width           =   1572
   End
   Begin VB.TextBox Text9 
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
      Height          =   300
      Left            =   6840
      TabIndex        =   2
      Top             =   1680
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   1680
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   529
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
   Begin MSForms.TextBox Text3 
      Height          =   315
      Left            =   7290
      TabIndex        =   22
      Top             =   600
      Width           =   1125
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      Size            =   "1984;556"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   300
      Left            =   2910
      TabIndex        =   5
      Top             =   960
      Width           =   5505
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      Size            =   "9710;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text8 
      Height          =   300
      Left            =   1320
      TabIndex        =   4
      Top             =   1320
      Width           =   7095
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      Size            =   "12515;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text10 
      Height          =   330
      Left            =   1320
      TabIndex        =   1
      Top             =   2040
      Width           =   7095
      VariousPropertyBits=   -1466941409
      BackColor       =   14737632
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "12515;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "(1. 個人 2. 公司)"
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
      Left            =   2280
      TabIndex        =   21
      Top             =   600
      Width           =   1875
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
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
      Left            =   5490
      TabIndex        =   20
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "收款金額"
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
      Left            =   3120
      TabIndex        =   17
      Top             =   1680
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
      Caption         =   "收款單號"
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
      TabIndex        =   14
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "個人/公司"
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
      TabIndex        =   13
      Top             =   600
      Width           =   1185
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "客戶代號"
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
      TabIndex        =   12
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭"
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
      TabIndex        =   11
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "收款日期"
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
      TabIndex        =   10
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "溢收金額"
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
      Left            =   5880
      TabIndex        =   9
      Top             =   1680
      Width           =   972
   End
   Begin VB.Label Label8 
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
      Left            =   300
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc1250"
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
Public adoacc0m0 As New ADODB.Recordset
Public adoacc0n0 As New ADODB.Recordset
Public adoacc0l0 As New ADODB.Recordset
Public adoacctmp05 As New ADODB.Recordset
Public ado0m0sum As New ADODB.Recordset
Public ado0n0sum As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Dim stSQL As String


Private Sub Command2_Click()
   strCon1 = Text1
   strFormLink = Name
   Frmacc1221.Show
   Me.Enabled = False
End Sub

Private Sub Form_Activate()
   strFormName = Name
   strFormLink = ""
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
   'Modify by Amy 2023/08/18 W9500 H5400
   Me.Width = 9630
   Me.Height = 5650
   'end 2023/08/18
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   Text1 = MsgText(803)
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   Acctmp05Delete
'   adoacctmp05.Close
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc1250 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   TableQuery
   AdodcRefresh
End Sub

Private Sub Text6_Change()
   Text7 = CustomerQuery(Text6, 1)
   If Text7 = MsgText(601) Then
      Text7 = CustomerQuery(Text6, 2)
      If Text7 = MsgText(601) Then
         Text7 = CustomerQuery(Text6, 3)
      End If
   End If
End Sub

'*************************************************
'  查詢資料表(國內收據資料)
'
'*************************************************
Private Sub TableQuery()
   If Text1 = MsgText(601) Or Text1 = MsgText(803) Then
      Exit Sub
   End If
   adoacc0k0.CursorLocation = adUseClient
   adoacc0k0.Open "select * from acc0m0, acc0k0, acc0l0 where a0m02 = a0k01 and a0m01 = a0l01 and a0m01 = '" & Text1 & "' order by a0k01 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0k0.RecordCount <> 0 Then
      adoacc0k0.MoveFirst
      FormShow
   Else
      FormClear
   End If
   adoacc0k0.Close
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   'Modify by Morgan 2011/8/23 改從 0j0 抓 cp
   'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
   adoadodc1.Open "select a0k02, a0k01, a0j02, a0j01, na03, getcp10desc(cp01,cp10,a0j04) cp10N, (a0j09 + a0j10) as RAmount, cp75 as EAmount, cp79 as NAmount from acc0k0, acc0j0, caseprogress,nation where a0j13(+) = a0k01 and cp09(+) = a0j01 and a0k03 = '1' and na01(+)=a0j04", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
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
Private Sub FormClear()
   Text2 = ""
   Text6 = ""
   Text7 = ""
   Text8 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   Text9 = ""
   Text10 = ""
   Text11 = ""
   Text12 = ""
End Sub

'*************************************************
'  顯示資料表(國內收據資料)
'
'*************************************************
Private Sub FormShow()
Dim lngSum As Long

   lngSum = 0
   If IsNull(adoacc0k0.Fields("a0k05").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = adoacc0k0.Fields("a0k05").Value
   End If
   If IsNull(adoacc0k0.Fields("a0k20").Value) Then
      Text12 = MsgText(601)
      Text3 = MsgText(601)
   Else
      Text12 = adoacc0k0.Fields("a0k20").Value
      Text3 = StaffQuery(adoacc0k0.Fields("a0k20").Value)
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
   If IsNull(adoacc0k0.Fields("a0l02").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(adoacc0k0.Fields("a0l02").Value)
   End If
'   adoacc0n0.CursorLocation = adUseClient
'   adoacc0n0.Open "select sum(a0n03) from acc0n0 where a0n01 = '" & Text1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   If adoacc0n0.RecordCount <> 0 Then
'      adoacc0m0.CursorLocation = adUseClient
'      adoacc0m0.Open "select sum(a0m04), sum(a0m05), sum(a0m06) from acc0m0 where a0m01 = '" & Text1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'      If adoacc0m0.RecordCount <> 0 Then
'          If IsNull(adoacc0m0.Fields(0).Value) Then
'             lngSum = 0
'          Else
'             lngSum = adoacc0m0.Fields(0).Value
'          End If
'          If IsNull(adoacc0m0.Fields(1).Value) = False Then
'             lngSum = lngSum + Val(adoacc0m0.Fields(1).Value)
'          End If
'          If IsNull(adoacc0m0.Fields(2).Value) = False Then
'             lngSum = lngSum + Val(adoacc0m0.Fields(2).Value)
'          End If
'          Text9 = Val(adoacc0n0.Fields(0).Value) - lngSum
'       Else
'          Text9 = adoacc0n0.Fields(0).Value
'       End If
'       adoacc0m0.Close
'    Else
'       Text9 = MsgText(601)
'    End If
'    adoacc0n0.Close
   adoacc0l0.CursorLocation = adUseClient
   adoacc0l0.Open "select sum(a1u04+a1u05) from acc1u0 where a1u01 = '" & Text1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc0l0.RecordCount <> 0 Then
      If IsNull(adoacc0l0.Fields(0).Value) Then
         Text11 = ""
      Else
         Text11 = Format(adoacc0l0.Fields(0).Value, DDollar)
      End If
   Else
      Text11 = 0
   End If
   adoacc0l0.Close
   adoacc0l0.CursorLocation = adUseClient
   adoacc0l0.Open "select a1p08 from acc1p0 where a1p04 = '" & Text1 & "' and a1p08 <> 0 and a1p05 = '2401'", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0l0.RecordCount <> 0 Then
      If IsNull(adoacc0l0.Fields("a1p08").Value) Then
         Text9 = MsgText(601)
      Else
         Text9 = Format(adoacc0l0.Fields("a1p08").Value, DDollar)
      End If
   Else
      Text9 = MsgText(601)
   End If
   adoacc0l0.Close
   If IsNull(adoacc0k0.Fields("a0l07").Value) Then
      Text10 = MsgText(601)
   Else
      Text10 = adoacc0k0.Fields("a0l07").Value
   End If
End Sub

'*************************************************
'  刪除資料表之記錄(智權人員帳款查詢暫存檔)
'
'*************************************************
Private Sub Acctmp05Delete()
   adoTaie.Execute "delete from acctmp05"
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
   
On Error GoTo Checking
   strSql = ""
   If Text1 <> "" Then
      strSql = strSql & " and a1u01 = '" & Text1 & "'"
   Else
      MsgBox MsgText(162) & Label1, , MsgText(5)
      Exit Sub
   End If
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   'Modify by Morgan 2011/8/23 改從 0j0 抓 cp
   'Modify by Morgan 2011/10/27 考慮拆收據情形改先寫暫存
   'adoadodc1.Open "select a0k02, a0k01, cp01||cp02||cp03||cp04 CaseNo, cp09, a0j21, a0j20, (cp16) as RAmount, cp75 as EAmount, cp79 as NAmount from acc1u0, acc0j0, acc0k0, caseprogress where a0j01(+)=a1u03 and a0j13(+)=a1u02 and a0k01(+) = a1u02 and cp09(+) = a1u03" & strSql, adoTaie, adOpenStatic, adLockReadOnly
   adoTaie.Execute "delete ACCTMP08 where T05='" & Me.Name & "' and T14='" & strUserNum & "'"
   
   strExc(0) = "select a1u02,a1u03,'" & Me.Name & "',a1u01,'" & strUserNum & "' T14 from acc1u0 where 1=1 " & strSql
      
   adoTaie.Execute "insert into ACCTMP08(T01,T02,T05,T06,T14) " & strExc(0), intI
   '更新金額欄位
   strSql = "update ACCTMP08 set T08=(select nvl(a0j09, 0)+nvl(a0j10, 0) from acc0j0 where a0j13=T01 and a0j01=T02)" & _
      ",(T09,T10,T11,T12)=(select nvl(sum(a1u04),0)+nvl(sum(a1u05),0) T09,nvl(sum(a1u06),0) T10" & _
      ",nvl(sum(a1u07),0)+nvl(sum(a1u09),0) T11,nvl(sum(a1u08),0)+nvl(sum(a1u10),0) T12 " & _
      " from acc1u0 where a1u02=T01 and a1u03=T02) where T05='" & Me.Name & "' and T14='" & strUserNum & "'"
   adoTaie.Execute strSql, intI
   
   'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
   strExc(0) = "select a0k02, a0k01, a0j02 CaseNo, a0j01 cp09, na03, getcp10desc(cp01,cp10,a0j04) cp10N, T08 as RAmount" & _
      ",T09 as EAmount, T08-T09-T11+T12 as NAmount" & _
      " from ACCTMP08, acc0j0, acc0k0,caseprogress,nation where T05='" & Me.Name & "' and T14='" & strUserNum & "' and a0j13(+)=T01 and a0j01(+)=T02 and a0k01(+)=T01 and cp09(+)=a0j01 and na01(+)=a0j04"
   
   adoadodc1.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
   'end 2011/10/27
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount = 0 Then
      Adodc1.Recordset.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
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
Public Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         If FormCheck Then
            Screen.MousePointer = vbHourglass
            'Add by Morgan 2005/6/8 加所別控制
            Erase strExc: strExc(1) = Text1
            If PUB_CheckCaseZone(strExc, pub_strUserOffice, "3") = True Then
               TableQuery
               AdodcRefresh
            End If
            '2005/6/8
            Screen.MousePointer = vbDefault
            Exit Sub
         Else
            MsgBox MsgText(181), , MsgText(5)
         End If
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Text1 <> MsgText(601) And Text1 <> MsgText(803) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

