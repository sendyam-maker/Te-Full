VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc3250 
   AutoRedraw      =   -1  'True
   Caption         =   "到期日別票據明細查詢"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   9285
   Begin VB.TextBox Text8 
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
      Left            =   6525
      TabIndex        =   20
      Top             =   4464
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc3250.frx":0000
      Height          =   3300
      Left            =   120
      TabIndex        =   18
      Top             =   1080
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   5821
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
      Caption         =   "到期日別票據明細資料"
      ColumnCount     =   8
      BeginProperty Column00 
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
      BeginProperty Column01 
         DataField       =   "a0e07"
         Caption         =   "銀行帳號"
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
      BeginProperty Column03 
         DataField       =   "a0g02"
         Caption         =   "銀行代號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "a0e04"
         Caption         =   "應收/應付"
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
      BeginProperty Column05 
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
      BeginProperty Column06 
         DataField       =   "a0e08"
         Caption         =   "票別"
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
      BeginProperty Column07 
         DataField       =   "contect"
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
            Alignment       =   2
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   3825.071
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   120
      Top             =   960
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
   Begin VB.TextBox Text6 
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
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   1572
   End
   Begin VB.TextBox Text5 
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
      Left            =   6720
      TabIndex        =   16
      Top             =   1440
      Width           =   1572
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
      Left            =   1200
      TabIndex        =   14
      Top             =   1440
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
      Left            =   6720
      TabIndex        =   12
      Top             =   1080
      Width           =   1572
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
      Left            =   1200
      TabIndex        =   10
      Top             =   1080
      Width           =   1572
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
      Left            =   6720
      TabIndex        =   4
      Top             =   600
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1200
      TabIndex        =   0
      Top             =   240
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
      Left            =   3120
      TabIndex        =   1
      Top             =   240
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
   Begin MSForms.TextBox Text7 
      Height          =   300
      Left            =   2790
      TabIndex        =   17
      Top             =   600
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
   Begin VB.Label Label10 
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
      Height          =   255
      Left            =   5760
      TabIndex        =   21
      Top             =   4470
      Width           =   735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "應收/付"
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
      Left            =   5760
      TabIndex        =   19
      Top             =   240
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   -120
      Top             =   4350
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "應付票額"
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
      Left            =   5760
      TabIndex        =   15
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "應收票額"
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
      Left            =   240
      TabIndex        =   13
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "預期餘額"
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
      Left            =   5760
      TabIndex        =   11
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "存款餘額"
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
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   855
      Left            =   120
      Top             =   120
      Width           =   8850
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "銀行帳號"
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
      Left            =   5760
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "銀行別"
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
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
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
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc3250"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/07 Form2.0已修改 Text7/DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc0h0 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Dim strSql As String

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
   'Modify by Amy 2023/10/11 避免切畫面仍要調整,故調大小 原W9500 H5500/(lngWidth - Me.Width) / 2-瑞婷
   Me.Width = 9400
   Me.Height = 5580
   Me.Move 0, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   Combo1.AddItem ComboItem(181)
   Combo1.AddItem ComboItem(182)
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc3250 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc0h0.CursorLocation = adUseClient
   adoacc0h0.Open "select * from acc0h0 where a0h01 = '" & Text5 & "' and a0h02 = '" & Text1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc0e0 where a0e15 = 0 and a0e17 = 0 and a0e21 = 0 and a0e25 = 0 and a0e01 = '" & Text6 & "' and a0e07 = '" & Text1 & "'  and a0e10 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0e10 <= " & Val(FCDate(MaskEdBox2.Text)) & " order by a0e01 asc, a0e02 asc", adoTaie, adOpenStatic, adLockReadOnly
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
Public Sub FormClear()
   Text2 = ""
   Text3 = ""
   Text4 = ""
   Text5 = ""
End Sub

'*************************************************
'  顯示資料(銀行帳戶資料)
'
'*************************************************
Public Sub FormShow()
Dim adoacc0b0 As New ADODB.Recordset
Dim adoaccsum As New ADODB.Recordset
Dim adoacc040 As New ADODB.Recordset
Dim intYear, intMonth As Integer

   adoacc0b0.CursorLocation = adUseClient
   adoacc0b0.Open "select * from acc0b0", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0b0.RecordCount <> 0 Then
      If IsNull(adoacc0b0.Fields("a0b01").Value) = False Then
         intYear = Val(Mid(CFDate(adoacc0b0.Fields("a0b01").Value), 1, 3))
         intMonth = Val(Mid(CFDate(adoacc0b0.Fields("a0b01").Value), 5, 2))
      Else
         If Mid(ServerDate, 5, 2) = 12 Then
            intYear = Val(Mid(CFDate(ACDate(ServerDate)), 1, 3)) - 1
            intMonth = 1
         Else
            intYear = Val(Mid(CFDate(ACDate(ServerDate)), 1, 3))
            intMonth = Val(Mid(CFDate(ACDate(ServerDate)), 5, 2)) - 1
         End If
      End If
   Else
      If Mid(ServerDate, 5, 2) = 12 Then
         intYear = Val(Mid(CFDate(ACDate(ServerDate)), 1, 3)) - 1
         intMonth = 1
      Else
         intYear = Val(Mid(CFDate(ACDate(ServerDate)), 1, 3))
         intMonth = Val(Mid(CFDate(ACDate(ServerDate)), 5, 2)) - 1
      End If
   End If
   adoacc0b0.Close
   adoacc040.CursorLocation = adUseClient
   adoacc040.Open "select a0408 from acc040 where a0401 = " & intYear & " and a0404 = '" & MsgText(55) & "' and a0403 = '1' and a0405 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0402 in (select max(a0402) from acc040 where a0401 = " & intYear & " and a0404 = '" & MsgText(55) & "' and a0403 = '1' and a0405 = '" & adoacc0h0.Fields("a0h08").Value & "')", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc040.RecordCount <> 0 Then
      If IsNull(adoacc040.Fields(0).Value) Then
         Text2 = MsgText(601)
      Else
         Text2 = adoacc040.Fields(0).Value
      End If
   Else
      Text2 = MsgText(601)
   End If
   adoacc040.Close
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(a0e11) from acc0e0 where a0e04 = '" & MsgText(18) & "' and a0e15 = 0 and a0e17 = 0 and a0e21 = 0 and a0e25 = 0 " & strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text4 = MsgText(601)
      Else
         Text4 = adoaccsum.Fields(0).Value
      End If
   Else
      Text4 = MsgText(601)
   End If
   adoaccsum.Close
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(a0e11) from acc0e0 where a0e04 = '" & MsgText(19) & "' and a0e15 = 0 and a0e17 = 0 and a0e21 = 0 and a0e25 = 0 and (a0e37 = 0 or a0e37 is null) and a0e22 = 0" & strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text5 = MsgText(601)
      Else
         Text5 = adoaccsum.Fields(0).Value
      End If
   Else
      Text5 = MsgText(601)
   End If
   adoaccsum.Close
   Text3 = Val(Text2) + Val(Text4) - Val(Text5)
End Sub

'*************************************************
'  查詢資料(銀行帳戶資料)
'
'*************************************************
Private Sub QueryAcc0h0()
   adoacc0h0.Close
   adoacc0h0.CursorLocation = adUseClient
   adoacc0h0.Open "select * from acc0h0 where a0h01 = '" & Text6 & "' and a0h02 = '" & Text1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc0h0.RecordCount <> 0 Then
      FormShow
   Else
      FormClear
   End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text6_Change()
   Text7 = A0g02Query(Text6)
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
Dim strUnion As String

On Error GoTo Checking
   strSql = MsgText(601)
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = " and a0e10 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0e10 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If Text6 <> MsgText(601) Then
      strSql = strSql & " and (a0e01 = '" & Text6 & "' or a0e19 = '" & Text6 & "')"
   End If
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and (a0e07 = '" & Text1 & "' or a0e20 = '" & Text1 & "')"
   End If
   strUnion = "select a0e01, a0e02, a0e10, a0e07, a0g02, a0e04, a0e11, a0e08, cu04 as contect from acc0e0, acc0g0, customer where a0e01 = a0g01 and a0e05 = '1' and substr(a0e06, 1, 8) = cu01 (+) and substr(a0e06, 9, 1) = cu02 (+) and a0e15 = 0 and a0e17 = 0 and a0e25 = 0 and a0e22 = 0 and a0e04 = '" & Mid(Combo1, 1, 1) & "'" & strSql
   strUnion = strUnion & " union select a0e01, a0e02, a0e10, a0e07, a0g02, a0e04, a0e11, a0e08, a0i02 as contect from acc0e0, acc0g0, acc0i0 where a0e01 = a0g01 and a0e05 = '2' and a0e06 = a0i01 (+) and a0e15 = 0 and a0e17 = 0 and a0e25 = 0 and a0e22 = 0 and a0e04 = '" & Mid(Combo1, 1, 1) & "'" & strSql
   strUnion = strUnion & " union select a0e01, a0e02, a0e10, a0e07, a0g02, a0e04, a0e11, a0e08, st02 as contect from acc0e0, acc0g0, staff where a0e01 = a0g01 and a0e05 = '3' and a0e06 = st01 (+) and a0e15 = 0 and a0e17 = 0 and a0e25 = 0 and a0e22 = 0 and a0e04 = '" & Mid(Combo1, 1, 1) & "'" & strSql
   strUnion = strUnion & " union select a0e01, a0e02, a0e10, a0e07, a0g02, a0e04, a0e11, a0e08, '' as contect from acc0e0, acc0g0 where a0e01 = a0g01 and a0e05 = '4' and a0e15 = 0 and a0e17 = 0 and a0e25 = 0 and a0e22 = 0 and a0e04 = '" & Mid(Combo1, 1, 1) & "'" & strSql & " order by a0e10 asc, a0e02 asc"
   adoadodc1.Open strUnion, adoTaie, adOpenStatic, adLockReadOnly
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
            AdodcRefresh
            QueryAcc0h0
            SumShow
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
'  計算並顯示合計
'
'*************************************************
Public Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(a0e11) from acc0e0 where a0e15 = 0 and a0e17 = 0 and a0e25 = 0 and a0e22 = 0 and a0e04 = '" & Mid(Combo1, 1, 1) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text8 = MsgText(601)
      Else
         Text8 = Format(adoaccsum.Fields(0).Value, DDollar)
      End If
   Else
      Text8 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Text1 <> MsgText(601) Then
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

