VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc1270 
   AutoRedraw      =   -1  'True
   Caption         =   "應付款查詢"
   ClientHeight    =   5112
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9444
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5112
   ScaleWidth      =   9444
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1110
      TabIndex        =   8
      Top             =   1380
      Width           =   3500
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   7200
      TabIndex        =   22
      Top             =   4650
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   6192
      TabIndex        =   20
      Top             =   4650
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   5184
      TabIndex        =   19
      Top             =   4650
      Width           =   975
   End
   Begin VB.TextBox Text3 
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
      Left            =   1110
      TabIndex        =   0
      Top             =   60
      Width           =   612
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc1270.frx":0000
      Height          =   2790
      Left            =   240
      TabIndex        =   16
      Top             =   1800
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   4911
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
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
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "廠商帳款查詢"
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "t0610"
         Caption         =   "公司"
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
         DataField       =   "t0603"
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
      BeginProperty Column02 
         DataField       =   "name"
         Caption         =   "名稱"
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
         DataField       =   "t0602"
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
      BeginProperty Column04 
         DataField       =   "t0601"
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
      BeginProperty Column05 
         DataField       =   "t0611"
         Caption         =   "欲處理日"
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
         DataField       =   "t0604"
         Caption         =   "發票號碼"
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
         DataField       =   "t0605"
         Caption         =   "應付金額"
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
         DataField       =   "t0606"
         Caption         =   "已付金額"
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
      BeginProperty Column09 
         DataField       =   "t0612"
         Caption         =   "傳票編號"
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
         DataField       =   "t0607"
         Caption         =   "未付金額"
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
      BeginProperty Column11 
         DataField       =   "t0608"
         Caption         =   "支票抬頭"
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
      BeginProperty Column12 
         DataField       =   "t0609"
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
            ColumnWidth     =   540.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1152
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1344.189
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   3504.189
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   3564.284
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "單據內容"
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
   Begin VB.TextBox Text6 
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
      Left            =   1110
      TabIndex        =   1
      Top             =   390
      Width           =   612
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1110
      TabIndex        =   4
      Top             =   1050
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
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
      Left            =   1110
      TabIndex        =   2
      Top             =   720
      Width           =   1572
   End
   Begin VB.TextBox Text2 
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
      Left            =   3030
      TabIndex        =   3
      Top             =   720
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3030
      TabIndex        =   5
      Top             =   1050
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   4650
      Visible         =   0   'False
      Width           =   1200
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
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   300
      Left            =   5910
      TabIndex        =   6
      Top             =   1050
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
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
   Begin MSMask.MaskEdBox MaskEdBox4 
      Height          =   300
      Left            =   7710
      TabIndex        =   7
      Top             =   1050
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
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
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   25
      Top             =   1050
      Width           =   255
   End
   Begin VB.Label Label11 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "欲處理日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4725
      TabIndex        =   24
      Top             =   1050
      Width           =   1125
   End
   Begin VB.Label Label10 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   " 公 司 別 "
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   75
      TabIndex        =   23
      Top             =   1410
      Width           =   975
   End
   Begin VB.Label Label7 
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
      Left            =   4410
      TabIndex        =   21
      Top             =   4650
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "(1.廠商 2.客戶)"
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
      Left            =   1830
      TabIndex        =   18
      Top             =   60
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "往來類別"
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
      Left            =   60
      TabIndex        =   17
      Top             =   60
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "(1.未付 2.已付 3.往來)"
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
      Left            =   1950
      TabIndex        =   14
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label8 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "查詢資料"
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
      Left            =   60
      TabIndex        =   13
      Top             =   390
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2790
      TabIndex        =   12
      Top             =   1050
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "付款日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   135
      TabIndex        =   11
      Top             =   1050
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "往來對象"
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
      Left            =   60
      TabIndex        =   10
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2790
      TabIndex        =   9
      Top             =   720
      Width           =   255
   End
End
Attribute VB_Name = "Frmacc1270"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/07 Form2.0已修改 DataGrid1
'Memo by Amy 2022/04/01 畫面「往來類別」拿掉 3.員工 - 瑞婷
'2014/01/08 增加 "公司"欄 顯示公司別  add by eric
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoacc0o0 As New ADODB.Recordset
Public adoacc0q0 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoacctmp06 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Dim strSql As String
Dim strCmp As String 'Add by Amy 2022/04/01

'Add by Amy 2022/04/01
Private Sub Combo2_GotFocus()
    TextInverse Combo2
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
    strCmp = ""
    If Trim(Combo2) = MsgText(601) Then Exit Sub
    
    strCmp = Combo2
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If
    If InStr(GetBookKeepCmp, strCmp) = 0 Then
        MsgBox Label8 & MsgText(63), , MsgText(5)
        Cancel = True
        Combo2.SetFocus
        Exit Sub
    ElseIf Len(Trim(Combo2)) = 1 Then
        Combo2 = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2022/04/01

Private Sub Command1_Click()
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   tool3_enabled
   If Adodc1.Recordset.Fields("t0606").Value = 0 Then
      strCon1 = Adodc1.Recordset.Fields("t0601").Value
      strCon3 = Adodc1.Recordset.Fields("t0610").Value 'Add by Amy 2014/02/05 +公司別
      Frmacc1271.Show
   Else
      strCon3 = Adodc1.Recordset.Fields("t0610").Value 'Add by Amy 2014/02/05 +公司別
      If IsNull(Adodc1.Recordset.Fields("t0602").Value) = False Then
         strCon1 = Adodc1.Recordset.Fields("t0602").Value
      Else
         strCon1 = ""
      End If
      If IsNull(Adodc1.Recordset.Fields("t0603").Value) = False Then
         strCon2 = Adodc1.Recordset.Fields("t0603").Value
      Else
         strCon2 = ""
      End If
      Frmacc1272.Show
   End If
   Me.Enabled = False
Checking:
   Exit Sub
End Sub

Private Sub Form_Activate()
   strFormName = Name
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
   Me.Width = 9660 'Moidy by Amy 2023/07/19  原:9500
   Me.Height = 5670 'Moidy by Amy 2023/07/19  原:5500
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'Add by Amy 2022/04/01 +公司別
   Combo2.AddItem "", 0
   Call Pub_SetCboCmp(Combo2, False, False, False, , 1)
   'end 2022/04/01
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   'Add by Amy 2022/04/08 +欲處理日
   MaskEdBox3.Enabled = False
   MaskEdBox4.Enabled = False
   MaskEdBox3.Mask = DFormat
   MaskEdBox4.Mask = DFormat
   'end 2022/04/08
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc1270 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  儲存資料表(國內應付款資料暫存檔)
'
'*************************************************
Private Sub Acctmp06Save()
Dim lngPreAmount, lngPayAmount As Long
Dim strQ As String 'Add by Amy 2022/04/01
Dim stra1p22 As String 'Add by Amy 2022//04/15

On Error GoTo Checking
   strSql = MsgText(601)
   '往來類別
   If Text3 <> MsgText(601) Then
      strSql = strSql & " and a0o02 = '" & Text3 & "'"
   End If
   '往來對象起
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and a0o03 >= '" & Text1 & "'"
   End If
   '往來對象迄
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a0o03 <= '" & Text2 & "'"
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      '往來類別為 1.廠商 or 3.員工
      If Text3 = Mid(ComboItem(91), 1, 1) Or Text3 = Mid(ComboItem(93), 1, 1) Then
         strSql = strSql & " and a0o05 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      Else
         'a0o06-欲處理日,查詢條件無誤 (a0o11-付款日:真正付款回寫的日期)
         strSql = strSql & " and a0o06 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      End If
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      '往來類別為 1.廠商 or 3.員工
      If Text3 = Mid(ComboItem(91), 1, 1) Or Text3 = Mid(ComboItem(93), 1, 1) Then
         strSql = strSql & " and a0o05 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      Else
         strSql = strSql & " and a0o06 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      End If
   End If
   'Add by Amy 2022/04/08 +欲處理日(只有查詢資料為 1.未付才會出現 )
   If MaskEdBox3.Text <> MsgText(601) And MaskEdBox3.Text <> MsgText(29) Then
        strSql = strSql & " and a0o06 >= " & Val(FCDate(MaskEdBox3.Text)) & ""
   End If
   If MaskEdBox4.Text <> MsgText(601) And MaskEdBox4.Text <> MsgText(29) Then
        strSql = strSql & " and a0o06 <= " & Val(FCDate(MaskEdBox4.Text)) & ""
   End If
   'end 2022/04/08
   'Add by Amy 2022/04/01 +公司別
   strCmp = ""
   If Trim(Combo2) <> MsgText(601) Then
      strCmp = Combo2
      If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
      End If
      strSql = strSql & " and a0o07='" & strCmp & "' "
   End If
   'end 2022/04/01
   If strSql <> MsgText(601) Then
      strSql = " where " & Mid(strSql, 5, Len(strSql) - 4)
   End If
   adoacctmp06.CursorLocation = adUseClient
   'Modify by Amy 2022/04/01 +ID,避免同時操作
   adoacctmp06.Open "select * from acctmp06 Where ID='" & strUserNum & "' ", adoTaie, adOpenDynamic, adLockBatchOptimistic
' 應付資料
   adoacc0o0.CursorLocation = adUseClient
   '2011/8/16 modify by sonia 加名稱name
   'adoacc0o0.Open "select * from acc0o0" & strSql & " order by a0o01 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoacc0o0.Open "select acc0o0.*,substr(nvl(nvl(st02,nvl(nvl(cu04,cu05||decode(cu88,NULL,NULL,' '||cu88) ||decode(cu89,NULL,NULL,' '||cu89) ||decode(cu90,NULL,NULL,' '||cu90)),cu06)),a0I02),1,40) name from acc0o0,staff,customer,acc0i0 " & strSql & " and a0o03=st01(+) and a0o03=a0i01(+) and substr(a0o03,1,8)=cu01(+) and substr(a0o03,9,1)=cu02(+) order by a0o01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc0o0.EOF = False
      adoacctmp06.AddNew
      adoacctmp06.Fields("t0601").Value = adoacc0o0.Fields("a0o01").Value
      '往來類別為 1.廠商 or 3.員工
      If Text3 = Mid(ComboItem(91), 1, 1) Or Text3 = Mid(ComboItem(93), 1, 1) Then
         '入帳日
         If IsNull(adoacc0o0.Fields("a0o05").Value) Then
            adoacctmp06.Fields("t0602").Value = adoacc0o0.Fields("a0o06").Value
         Else
            adoacctmp06.Fields("t0602").Value = adoacc0o0.Fields("a0o05").Value
         End If
      Else
         '欲處理日
         If IsNull(adoacc0o0.Fields("a0o06").Value) Then
            adoacctmp06.Fields("t0602").Value = adoacc0o0.Fields("a0o05").Value
         Else
            adoacctmp06.Fields("t0602").Value = adoacc0o0.Fields("a0o06").Value
         End If
      End If
      adoacctmp06.Fields("t0610").Value = adoacc0o0.Fields("a0o07").Value '2014/01/08 add by eric
      adoacctmp06.Fields("t0603").Value = adoacc0o0.Fields("a0o03").Value
      adoacctmp06.Fields("name").Value = adoacc0o0.Fields("name").Value  '2011/8/16 add by sonia
      adoacctmp06.Fields("t0604").Value = adoacc0o0.Fields("a0o04").Value
      adoacctmp06.Fields("t0609").Value = adoacc0o0.Fields("a0o10").Value
      'Add by Amy 2022/04/01 +欲處理日,前面判斷「入帳日」為空寫入「欲處理日」,不是空寫入「入帳日」,秀玲說入帳日一定會有值,故寫入欲處理日
      adoacctmp06.Fields("t0611").Value = adoacc0o0.Fields("a0o06").Value
      adoacctmp06.Fields("ID").Value = strUserNum
      'end 2022/04/01
      adoaccsum.CursorLocation = adUseClient
      '2014/1/9 modify by sonia a1p01 = '1'改用a0o07
      adoaccsum.Open "select sum(a1p08) from acc1p0 where a1p01 = '" & adoacc0o0.Fields("a0o07").Value & "' and a1p02 = 'B' and a1p04 = '" & adoacc0o0.Fields("a0o01").Value & "' and a1p05 in ('2112', '2113') UNION " & _
                     "select sum(a1p08) from acc1p0 where a1p01 = '" & adoacc0o0.Fields("a0o07").Value & "' and a1p02 = 'E' and a1p23 = '" & adoacc0o0.Fields("a0o01").Value & "' and a1p05 in ('2112', '2113') union " & _
                     "select sum(a1p08) from acc1p0 where a1p01 = '" & adoacc0o0.Fields("a0o07").Value & "' and a1p02 = 'Z' and a1p23 = '" & adoacc0o0.Fields("a0o01").Value & "' and a1p05 in ('2112', '2113')", adoTaie, adOpenStatic, adLockReadOnly
      If adoaccsum.RecordCount <> 0 Then
         If IsNull(adoaccsum.Fields(0).Value) Then
            adoacctmp06.Fields("t0605").Value = 0
            lngPreAmount = 0
         Else
            adoacctmp06.Fields("t0605").Value = adoaccsum.Fields(0).Value
            lngPreAmount = Val(adoaccsum.Fields(0).Value)
         End If
      Else
         adoacctmp06.Fields("t0605").Value = 0
         lngPreAmount = 0
      End If
      adoaccsum.Close
      
      'Add by Amy 2022/04/01 寫入應付傳票資料
      'Modify by Amy 2022/04/14 改寫至function
      'Modify by Amy 2022/04/15 避免不止回傳一筆傳票號,導致錯誤
      stra1p22 = GetVoucherNo(1, "" & adoacc0o0.Fields("a0o07"), "" & adoacc0o0.Fields("a0o01"))
      If InStr(stra1p22, ",") > 0 Then
        adoacctmp06.Fields("t0612").Value = "多筆"
        MsgBox "回傳「傳票號碼」資料有問題,請洽電腦中心"
      Else
        adoacctmp06.Fields("t0612").Value = stra1p22
      End If
      
      adoacc0q0.CursorLocation = adUseClient
      adoacc0q0.Open "select a0q06 from acc0q0, acc0o0 where a0q01 = a0o11 and a0q03 = a0o03 and a0o01 = '" & adoacc0o0.Fields("a0o01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoacc0q0.RecordCount <> 0 Then
'         adoacctmp06.Fields("t0606").Value = lngPreAmount
         adoacctmp06.Fields("t0606").Value = 0
         lngPayAmount = lngPreAmount
      Else
         adoacctmp06.Fields("t0606").Value = 0
         lngPayAmount = 0
      End If
      adoacc0q0.Close
      
      adoacctmp06.Fields("t0607").Value = lngPreAmount - lngPayAmount
      adoacctmp06.UpdateBatch
      adoacc0o0.MoveNext
   Loop
   adoacc0o0.Close
   
' 付款資料
   strSql = MsgText(601)
   If Text3 <> MsgText(601) Then
      strSql = strSql & " and a0q04 = '" & Text3 & "'"
   End If
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and a0q03 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a0q03 <= '" & Text2 & "'"
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0q01 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0q01 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   'Add by Amy 2022/04/01 +公司別
   If Trim(Combo2) <> MsgText(601) Then
       strSql = strSql & " and a0q19='" & strCmp & "' "
   End If
   'end 2022/04/01
   If strSql <> MsgText(601) Then
      strSql = " where " & Mid(strSql, 5, Len(strSql) - 4)
   End If
   adoacc0o0.CursorLocation = adUseClient
   '2011/8/16 modify by sonia 加名稱name
   'adoacc0o0.Open "select * from acc0q0" & strSql & " order by a0q01 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoacc0o0.Open "select acc0q0.*,substr(nvl(nvl(st02,cu04),a0I02),1,40) name from acc0q0,staff,customer,acc0i0 " & strSql & " and a0q03=st01(+) and a0q03=a0i01(+) and substr(a0q03,1,8)=cu01(+) and substr(a0q03,9,1)=cu02(+) order by a0q01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc0o0.EOF = False
      adoacctmp06.AddNew
      adoacctmp06.Fields("t0610").Value = adoacc0o0.Fields("a0q19").Value '2014/01/08 add by eric
      'adoacctmp06.Fields("t0601").Value = adoacc0o0.Fields("a0q01").Value 'Mak by Amy 2022/04/01 已收不顯示G單號,避免與未收搞混-秀玲
      adoacctmp06.Fields("t0602").Value = adoacc0o0.Fields("a0q01").Value
      adoacctmp06.Fields("t0603").Value = adoacc0o0.Fields("a0q03").Value
      adoacctmp06.Fields("name").Value = adoacc0o0.Fields("name").Value  '2011/8/16 add by sonia
      adoacctmp06.Fields("t0605").Value = 0
      adoacctmp06.Fields("t0606").Value = adoacc0o0.Fields("a0q06").Value
      adoacctmp06.Fields("t0607").Value = 0
      If IsNull(adoacc0o0.Fields("a0q05").Value) = False Then
         adoacctmp06.Fields("t0608").Value = adoacc0o0.Fields("a0q05").Value
      End If
      'Add by Amy 2022/04/14 寫入付款傳票資料
      'Modify by Amy 2022/04/15 避免不止回傳一筆傳票號,導致錯誤
      stra1p22 = GetVoucherNo(2, "" & adoacc0o0.Fields("a0q19"), "" & adoacc0o0.Fields("a0q17"))
      If InStr(stra1p22, ",") > 0 Then
        adoacctmp06.Fields("t0612").Value = "多筆"
        MsgBox "回傳「傳票號碼」資料有問題,請洽電腦中心"
      Else
        adoacctmp06.Fields("t0612").Value = stra1p22
      End If
      
      adoacctmp06.Fields("ID").Value = strUserNum 'Add by Amy 2022/04/01 +ID
      adoacctmp06.UpdateBatch
      adoacc0o0.MoveNext
   Loop
   adoacc0o0.Close
   adoacctmp06.Close
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   adoacc0o0.Close
   adoacctmp06.Close
   MsgBox Err.Description, , MsgText(5)
End Sub

'Add by Amy 2022/04/14 抓取1.應付/2.付款 傳票編號
Private Function GetVoucherNo(ByVal intChoose, ByVal stCmp As String, ByVal stDocNo As String) As String
    Dim RsQ As New ADODB.Recordset, strQ As String, intQ As Integer
    
    GetVoucherNo = ""
    If intChoose = 1 Then
        strQ = "Select a1p22 from acc1p0 where a1p01 = '" & stCmp & "'  and a1p02 ='B' and a1p04 = '" & stDocNo & "' and  a1p05 in ('2112', '2113') Union " & _
                   "Select a1p22 from acc1p0 where a1p01 = '" & stCmp & "' and a1p02 In ('E','Z') and a1p23 = '" & stDocNo & "' and a1p05 in ('2112', '2113') "
    Else
        'Modify by Amy 2022/04/15 +Distinct 因1公司 D111010596 有兩項2112
        strQ = "Select Distinct a1p22 from acc1p0 where a1p01 = '" & stCmp & "'  and a1p02 ='C' and a1p04 = '" & stDocNo & "' and  a1p05 in ('2112', '2113') "
    End If
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        RsQ.MoveFirst
        Do While RsQ.EOF = False
            GetVoucherNo = GetVoucherNo & "," & RsQ.Fields("a1p22")
            RsQ.MoveNext
        Loop
        If GetVoucherNo <> MsgText(601) Then
            GetVoucherNo = Mid(GetVoucherNo, 2)
        End If
    End If
    Set RsQ = Nothing
End Function

'*************************************************
'  刪除資料表之記錄(收據作廢查詢暫存檔)
'
'*************************************************
Private Sub Acctmp06Delete()
   '避免多人操作 +ID
   adoTaie.Execute "Delete from acctmp06 Where ID='" & strUserNum & "' "
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         Call Text6_Validate(False) 'Add by Amy 2022/04/08 避免 於Text6修改後,直接按F12 欲處理日資料仍在
         If FormCheck Then
            Screen.MousePointer = vbHourglass
            Acctmp06Delete
            Acctmp06Save
            AdodcRefresh
            SumShow
            Screen.MousePointer = vbDefault
            Exit Sub
         'Mark by Amy 2022/04/01 FormCheck判斷必填欄位
'         Else
'            MsgBox MsgText(181), , MsgText(5)
         End If
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text3 = Mid(ComboItem(92), 1, 1) Then
      If Len(Text1) = 6 Then
         Text1 = AfterZero(Text1)
      Else
         If Len(Text1) = 8 Then
            Text1 = Text1 & "0"
         End If
      End If
   End If
   Text2 = Text1
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text3 = Mid(ComboItem(92), 1, 1) Then
      If Len(Text2) = 6 Then
         Text2 = AfterZero(Text2)
      End If
   End If
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
    Dim strWhere As String
    
On Error GoTo Checking
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   '查詢資料
   'Modify by Amy 2022/04/01 +ID
   strSql = "Select * From Acctmp06 Where ID='" & strUserNum & "' "
   Select Case Text6
      Case "1" '未付
         strWhere = "And  t0607 <> 0 "
         'adoadodc1.Open "select * from acctmp06 where t0607 <> 0  order by t0603 asc, t0602 asc, t0601 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      Case "2" '已付
         strWhere = "And  t0606 <> 0 "
         'adoadodc1.Open "select * from acctmp06 where t0606 <> 0 order by t0603 asc, t0602 asc, t0601 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      Case Else '往來
         'adoadodc1.Open "select * from acctmp06 order by t0603 asc, t0602 asc, t0601 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   End Select
   strSql = strSql & strWhere & " Order by t0603 asc, t0602 asc, t0601 asc"
   adoadodc1.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   'end 2022/04/01
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
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acctmp06 where t0601 = '1' order by t0601 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  計算並顯示總金額
'
'*************************************************
Public Sub SumShow()
   Text4 = ""
   Text7 = ""
   Text5 = ""
   If Adodc1.Recordset.State <> adStateOpen Then
      Exit Sub
   End If
   Set adoaccsum = Adodc1.Recordset.Clone
   Do While adoaccsum.EOF = False
      Text4 = Val(Text4) + adoaccsum.Fields("t0605").Value
      Text7 = Val(Text7) + adoaccsum.Fields("t0606").Value
      Text5 = Val(Text5) + adoaccsum.Fields("t0607").Value
      adoaccsum.MoveNext
   Loop
   adoaccsum.Close
   Text4 = Format(Text4, DDollar)
   Text7 = Format(Text7, DDollar)
   Text5 = Format(Text5, DDollar)
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   
   'Modify by Amy 2022/04/01 有下往來類別有資料,沒下反而沒資料,秀玲說改必填
   If Text3 = MsgText(601) Then
      'FormCheck = True
      MsgBox Label5.Caption & "不可為空！"
      Exit Function
   End If
   'Modify by Amy 2022/04/15 因2022/04/01 加「欲處理日」,若未付沒下「入帳日」條件,導致抓付款資料時,資料範圍太大,
   '                                             且有重覆資料出現「多重步驟操作錯誤…」的訊息(避免未下日期資料量過大出現錯誤,日期一定要下)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) _
    Or MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
        MsgBox "「" & Label3.Caption & "」起迄不可為空！"
        If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
           MaskEdBox1.SetFocus
        Else
           MaskEdBox2.SetFocus
        End If
        Exit Function
   End If
 
   If Text1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text2 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If Text6 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

'Add by Amy 2022/04/01 判斷輸 1.未付  畫面「付款日期」改為「入帳日期」
Private Sub Text6_Validate(Cancel As Boolean)
    'Modify by Amy 2022/04/08 +欲處理日
    Label3.Caption = "付款日期"
    'If Trim(Text6) = MsgText(601) Then Exit Sub
    
    If Trim(Text6) = "1" Then
        Label3.Caption = "入帳日期"
       
        MaskEdBox3.Enabled = True
        MaskEdBox4.Enabled = True
    Else
        MaskEdBox3.Mask = ""
        MaskEdBox3.Text = ""
        MaskEdBox3.Mask = DFormat
        MaskEdBox4.Mask = ""
        MaskEdBox4.Text = ""
        MaskEdBox4.Mask = DFormat
        MaskEdBox3.Enabled = False
        MaskEdBox4.Enabled = False
    End If
End Sub
