VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc21s0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "其他幣別請款匯率資料維護"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   6435
   Begin VB.CommandButton Command5 
      Height          =   300
      Left            =   2970
      Picture         =   "Frmacc21s0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   210
      Width           =   350
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   1380
      TabIndex        =   0
      Top             =   210
      Width           =   1572
   End
   Begin VB.TextBox textDNR04 
      Alignment       =   1  '靠右對齊
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.000000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   1
      EndProperty
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
      Left            =   4320
      MaxLength       =   15
      TabIndex        =   3
      Top             =   570
      Width           =   1572
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc21s0.frx":0102
      Height          =   4365
      Left            =   180
      TabIndex        =   4
      Top             =   1200
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   7699
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   20
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
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "dnr01"
         Caption         =   "請款幣別"
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
         DataField       =   "dnr02"
         Caption         =   "日期"
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
      BeginProperty Column02 
         DataField       =   "dnr03"
         Caption         =   "對台幣匯率"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "dnr04"
         Caption         =   "對美金匯率"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1409.953
         EndProperty
      EndProperty
   End
   Begin VB.TextBox textDNR03 
      Alignment       =   1  '靠右對齊
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.000000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   1
      EndProperty
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
      Left            =   1380
      MaxLength       =   15
      TabIndex        =   2
      Top             =   570
      Width           =   1572
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   180
      Top             =   1080
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   4320
      TabIndex        =   1
      Top             =   210
      Width           =   1575
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
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "日期"
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
      Left            =   3840
      TabIndex        =   8
      Top             =   210
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "對美金匯率 "
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
      TabIndex        =   7
      Top             =   570
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "對台幣匯率"
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
      Left            =   210
      TabIndex        =   6
      Top             =   570
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "請款幣別"
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
      Left            =   420
      TabIndex        =   5
      Top             =   210
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc21s0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/12/01 Form2.0已檢查 (無需修改的物件)
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo By Sindy 2010/8/12 日期欄已修改
Option Explicit

Public adousxrate As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset

Private Sub Combo1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'檢查請款幣別是否存在
Private Sub Combo1_Validate(Cancel As Boolean)
   If Combo1 = MsgText(601) Then
      Exit Sub
   End If
   If Combo1 = "USD" Or ExistCheck("acc1y0", "a1y01", Combo1, Label1) = False Then
      If Combo1 = "USD" Then MsgBox "幣別不可為美金！", , MsgText(5)
      Cancel = True
      Combo1.SetFocus
   End If
End Sub

'Added by Morgan 2019/10/5
Private Sub Command5_Click()
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   If Combo1 <> MsgText(601) Then
      Adodc1.Recordset.Find "dnr01 = '" & Trim(Combo1.Text) & "'", 0, adSearchForward, 1
      If Adodc1.Recordset.EOF Then
         Adodc1.Recordset.MoveFirst
         Adodc1.Recordset.Find "dnr01 = '" & Trim(Combo1.Text) & "'", 0, adSearchForward, 1
      End If
      If Adodc1.Recordset.EOF Then
         MsgBox "查無資料！", vbExclamation
         Adodc1.Recordset.MoveFirst
      End If
      FormShow
      RecordShow
   End If
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   FormShow
   RecordShow
End Sub

Private Sub Form_Activate()
   '93.3.16 ADD BY SONIA
   'Modified by Morgan 2019/10/5
   'If IsObject(mdiMain) Then
   '   ToolShow
   'End If
   If UCase(Forms(0).Name) = "MDIMAIN" Then
      Forms(0).ToolShow
   End If
   'end 2019/10/5
   '93.3.16 END
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   PUB_InitForm Me, Me.Width, Me.Height
   
   MaskEdBox1.Mask = DFormat
   
   '預設請款幣別下拉選單
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from acc1y0 where a1y01<>'USD' order by a1y01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoquery.EOF = False
      Combo1.AddItem adoquery.Fields("a1y01").Value
      adoquery.MoveNext
   Loop
   adoquery.Close
   
   OpenTable
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
   Set Frmacc21s0 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from debitnoterate order by dnr01,dnr02 desc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示欄位資料(美金匯率資料表)
'
'*************************************************
Public Sub FormShow()
   If IsNull(Adodc1.Recordset.Fields("dnr01").Value) Then
      Combo1.Text = MsgText(601)
   Else
      Combo1.Text = Adodc1.Recordset.Fields("dnr01").Value
   End If
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(Adodc1.Recordset.Fields("dnr02").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(Adodc1.Recordset.Fields("dnr02").Value)
   End If
   MaskEdBox1.Mask = DFormat
   If IsNull(Adodc1.Recordset.Fields("dnr03").Value) Then
      textDNR03 = MsgText(601)
   Else
      textDNR03 = Adodc1.Recordset.Fields("dnr03").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("dnr04").Value) Then
      textDNR04 = MsgText(601)
   Else
      textDNR04 = Adodc1.Recordset.Fields("dnr04").Value
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
   adoadodc1.Open "select * from debitnoterate order by dnr01,dnr02 desc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) And _
         Combo1.Text <> MsgText(601) Then
         Adodc1.Recordset.Find "dnr01 = '" & Trim(Combo1.Text) & "'", 0, adSearchForward, 1
         If Adodc1.Recordset.EOF = False Then
            Adodc1.Recordset.Find "dnr02 = " & Val(FCDate(MaskEdBox1.Text)) & "", 0, adSearchForward, adoadodc1.Bookmark
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
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   CountShow Adodc1.Recordset.Bookmark, Adodc1.Recordset.RecordCount
End Sub

Private Sub MaskEdBox1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub textDNR03_GotFocus()
   TextInverse textDNR03
End Sub

Private Sub textDNR03_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub textDNR04_GotFocus()
   TextInverse textDNR04
End Sub

Private Sub textDNR04_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub
