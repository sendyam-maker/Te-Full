VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc4130 
   AutoRedraw      =   -1  'True
   Caption         =   "公司基本資料"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   8760
   Begin VB.TextBox Text15 
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
      Left            =   4290
      MaxLength       =   5
      TabIndex        =   33
      Top             =   120
      Width           =   285
   End
   Begin VB.TextBox Text14 
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
      Left            =   2100
      MaxLength       =   6
      TabIndex        =   14
      Top             =   3210
      Width           =   1575
   End
   Begin VB.TextBox Text13 
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
      Left            =   5880
      MaxLength       =   9
      TabIndex        =   13
      Top             =   3210
      Width           =   1572
   End
   Begin VB.TextBox txtAddr2 
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
      Left            =   2100
      MaxLength       =   45
      TabIndex        =   7
      Top             =   1820
      Width           =   6480
   End
   Begin VB.TextBox txtAddr1 
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
      Left            =   2100
      MaxLength       =   45
      TabIndex        =   6
      Top             =   1480
      Width           =   6480
   End
   Begin VB.TextBox Text12 
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
      Left            =   6570
      MaxLength       =   9
      TabIndex        =   12
      Top             =   2840
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
      Height          =   315
      Left            =   5880
      MaxLength       =   50
      TabIndex        =   8
      Top             =   2160
      Width           =   2700
   End
   Begin VB.CommandButton Command1 
      Height          =   300
      Left            =   2880
      Picture         =   "Frmacc4130.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   127
      Width           =   345
   End
   Begin VB.TextBox Text8 
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
      Left            =   2100
      MaxLength       =   50
      TabIndex        =   15
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox Text7 
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
      Left            =   2100
      MaxLength       =   10
      TabIndex        =   11
      Top             =   2840
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc4130.frx":0102
      Height          =   1365
      Left            =   240
      TabIndex        =   24
      Top             =   3720
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   2408
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   20
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.25
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
         DataField       =   "a0801"
         Caption         =   "公司別"
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
         DataField       =   "a0802"
         Caption         =   "公司名稱(中)"
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
         DataField       =   "st02"
         Caption         =   "負責人"
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
         DataField       =   "a0807"
         Caption         =   "統一編號"
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
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4169.764
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1409.953
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   990
      Top             =   -30
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
   Begin VB.TextBox Text2 
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
      Left            =   6030
      MaxLength       =   5
      TabIndex        =   2
      Top             =   127
      Width           =   1092
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
      Height          =   315
      Left            =   5880
      MaxLength       =   15
      TabIndex        =   10
      Top             =   2500
      Width           =   1572
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
      Height          =   315
      Left            =   2100
      MaxLength       =   8
      TabIndex        =   9
      Top             =   2500
      Width           =   1575
   End
   Begin VB.TextBox Text4 
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
      Left            =   2100
      MaxLength       =   50
      TabIndex        =   4
      Top             =   800
      Width           =   6480
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
      Left            =   2100
      MaxLength       =   1
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin MSForms.TextBox Text5 
      Height          =   315
      Left            =   7110
      TabIndex        =   23
      Top             =   127
      Width           =   1452
      VariousPropertyBits=   679493659
      BackColor       =   14737632
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text6 
      Height          =   315
      Left            =   2100
      TabIndex        =   5
      Top             =   1140
      Width           =   6480
      VariousPropertyBits=   679493659
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text3 
      Height          =   315
      Left            =   2100
      TabIndex        =   3
      Top             =   460
      Width           =   6480
      VariousPropertyBits=   679493659
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label16 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "作帳公司　　Y:是"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3330
      TabIndex        =   34
      Top             =   150
      Width           =   1800
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "公司簡稱"
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
      TabIndex        =   32
      Top             =   3255
      Width           =   2295
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "勞工保險證號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4425
      TabIndex        =   31
      Top             =   3255
      Width           =   1350
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "英文地址2"
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
      TabIndex        =   30
      Top             =   1845
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "英文地址1"
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
      TabIndex        =   29
      Top             =   1515
      Width           =   1455
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "健保局投保單位代號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4425
      TabIndex        =   28
      Top             =   2892
      Width           =   2025
   End
   Begin VB.Label Label7 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "電話"
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
      Left            =   4680
      TabIndex        =   27
      Top             =   2190
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "網　　址"
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
      TabIndex        =   26
      Top             =   2190
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "營利事業稅籍編號"
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
      TabIndex        =   25
      Top             =   2865
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   3585
      Left            =   150
      Top             =   45
      Width           =   8505
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label14 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "房屋稅籍編號"
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
      Left            =   4320
      TabIndex        =   22
      Top             =   2530
      Width           =   1455
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "統一編號"
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
      TabIndex        =   21
      Top             =   2535
      Width           =   1095
   End
   Begin VB.Label Label12 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "負責人"
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
      Left            =   5190
      TabIndex        =   20
      Top             =   150
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "發票地址"
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
      TabIndex        =   19
      Top             =   1170
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "公司名稱(英)"
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
      TabIndex        =   18
      Top             =   825
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "公司名稱(中)"
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
      TabIndex        =   17
      Top             =   495
      Width           =   1455
   End
   Begin VB.Label Label3 
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
      Left            =   210
      TabIndex        =   16
      Top             =   150
      Width           =   855
   End
End
Attribute VB_Name = "Frmacc4130"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/08/02 Form2.0已修改 Text3/Text5/Text6/DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc080 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset

Private Sub Command1_Click()
   If Adodc1.Recordset.RecordCount = 0 Or Text1 = MsgText(601) Then
      Exit Sub
   End If
   Adodc1.Recordset.Find "a0801 = '" & Text1 & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      FormShow
      RecordShow
   Else
      MsgBox MsgText(33), , MsgText(5)
      Adodc1.Recordset.MoveFirst
   End If
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
   Adodc1.Recordset.Find "a0801 = '" & strCompanyNo & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      FormShow
      RecordShow
   End If
   strCompanyNo = MsgText(601)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Form2.0 記錄鍵盤傳入順序
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
   
   'Modified by Lydia 2017/09/06 表單初始化
'   Me.Width = 8850
'   Me.Height = 5500
'
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath1)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   PUB_InitForm Me, 8850, 5700, strBackPicPath1
   'end 2017/09/06
   
   OpenTable
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.MoveLast
      Adodc1.Recordset.MoveFirst
      RecordShow
   End If
   'Add by Amy 2020/03/13 M51可看見作帳公司
   Label16.Visible = False
   Text15.Visible = False
   If Pub_StrUserSt03 = "M51" Then
        Label16.Visible = True
        Text15.Visible = True
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
   strTrackMode = "" 'Add by Amy 2021/10/26 Form2.0 記錄鍵盤傳入順序(清除)
   Set Frmacc4130 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc080, staff where a0806 = st01 (+) order by a0801 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示欄位資料(公司別資料)
'
'*************************************************
Public Sub FormShow()
   If IsNull(Adodc1.Recordset.Fields("a0801").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = Adodc1.Recordset.Fields("a0801").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0806").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = Adodc1.Recordset.Fields("a0806").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0802").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = Adodc1.Recordset.Fields("a0802").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0803").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = Adodc1.Recordset.Fields("a0803").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0804").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = Adodc1.Recordset.Fields("a0804").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0807").Value) Then
      Text10 = MsgText(601)
   Else
      Text10 = Adodc1.Recordset.Fields("a0807").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0809").Value) Then
      Text11 = MsgText(601)
   Else
      Text11 = Adodc1.Recordset.Fields("a0809").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0808").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = Adodc1.Recordset.Fields("a0808").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0805").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = Adodc1.Recordset.Fields("a0805").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0813").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = Adodc1.Recordset.Fields("a0813").Value
   End If
   
   '2008/12/15 CANCEL BY SONIA 與婧瑄確認不需此欄故取消
   'If IsNull(Adodc1.Recordset.Fields("a0814").Value) Then
   '   Text12 = MsgText(601)
   'Else
   '   Text12 = Adodc1.Recordset.Fields("a0814").Value
   'End If
   '2008/12/15 END
   
   'Added by Morgan 2013/3/13
   Text12 = "" & Adodc1.Recordset.Fields("a0821").Value
   
   'Added by Lydia 2017/09/06 +英文地址1,2
   If IsNull(Adodc1.Recordset.Fields("a0822").Value) Then
      txtAddr1 = MsgText(601)
   Else
      txtAddr1 = Adodc1.Recordset.Fields("a0822").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a0823").Value) Then
      txtAddr2 = MsgText(601)
   Else
      txtAddr2 = Adodc1.Recordset.Fields("a0823").Value
   End If
   'end 2017/09/06
   
   Text13 = "" & Adodc1.Recordset.Fields("a0824").Value 'Added by Morgan 2017/10/26
   'Add by Amy 2020/03/13
   Text14 = "" & Adodc1.Recordset.Fields("a0820").Value '公司簡稱
   '是作帳公司
   If IsNull(Adodc1.Recordset.Fields("a0827").Value) Then
        Text15 = MsgText(601)
   Else
        Text15 = "" & Adodc1.Recordset.Fields("a0827").Value
   End If
   'end 2020/03/13
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   If Text1 = MsgText(601) Then
      MsgBox Label3 & MsgText(52), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 = MsgText(601) Then
      Exit Sub
   End If
   If Len(Text10) <> 8 Then
      Cancel = True
      Exit Sub
   End If
   If UnionCode(Text10) = MsgText(603) Then
      MsgBox Label13 & MsgText(63), , MsgText(5)
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
   CloseIme
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text13_GotFocus()
   TextInverse Text13
   CloseIme
End Sub

'Add by Amy 2020/03/13
Private Sub Text15_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
        KeyAscii = 0
    End If
End Sub

'2008/12/15 cancal by sonia
'Private Sub Text12_GotFocus()
'   TextInverse Text12
'End Sub
'2008/12/15 end

Private Sub Text2_Change()
   Text5 = StaffQuery(Text2)
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 <> MsgText(601) Then
      If ExistCheck("staff", "st01", Text2, Label12) = False Then
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub Text3_GotFocus()
   StatusView MsgText(65) & "30"
   TextInverse Text3
   'add by nickc 2007/07/13 將輸入法改成使用API
   OpenIme
End Sub

Private Sub Text3_LostFocus()
   StatusView MsgText(601)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If CheckLen(Label1, Text3, 30) = MsgText(603) Then
      Cancel = True
      Exit Sub
   End If
   'add by nickc 2007/07/13 將輸入法改成使用API
   If Cancel = False Then CloseIme
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_GotFocus()
   StatusView MsgText(65) & "70"
   TextInverse Text6
   'add by nickc 2007/07/13 將輸入法改成使用API
   OpenIme
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc080, staff where a0806 = st01 (+) order by a0801 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      If Text1 <> MsgText(601) Then
         Adodc1.Recordset.Find "a0801 = '" & Text1 & "'", 0, adSearchForward, 1
         If Adodc1.Recordset.EOF = False Then
            FormShow
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
   Frmacc0000.StatusBar1.Panels(2).Text = Adodc1.Recordset.Bookmark & MsgText(35) & Adodc1.Recordset.RecordCount
End Sub

Private Sub Text6_LostFocus()
   StatusView MsgText(601)
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   If CheckLen(Label6, Text6, 70) = MsgText(603) Then
      Cancel = True
      Exit Sub
   End If
   'add by nickc 2007/07/13 將輸入法改成使用API
   If Cancel = False Then CloseIme
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

'Added by Lydia 2017/09/06
Private Sub txtAddr1_GotFocus()
   TextInverse txtAddr1
End Sub

Private Sub txtAddr1_KeyPress(KeyAscii As Integer)
   'KeyAscii = UpperCase(KeyAscii) 'Mark by Lydia 2017/09/07 可小寫
End Sub

Private Sub txtAddr2_GotFocus()
   TextInverse txtAddr2
End Sub

Private Sub txtAddr2_KeyPress(KeyAscii As Integer)
   'KeyAscii = UpperCase(KeyAscii) 'Mark by Lydia 2017/09/07 可小寫
End Sub
'Added by Lydia 2017/09/07
Private Sub txtAddr1_Validate(Cancel As Boolean)
   If txtAddr1.Text <> "" Then
     If PUB_CheckStrNEC(Trim(txtAddr1.Text)) Then
        MsgBox "不可輸入中文 !!"
        txtAddr1.SetFocus
        Cancel = True
        Exit Sub
     End If
   End If
End Sub

Private Sub txtAddr2_Validate(Cancel As Boolean)
   If txtAddr2.Text <> "" Then
     If PUB_CheckStrNEC(Trim(txtAddr2.Text)) Then
        MsgBox "不可輸入中文 !!"
        txtAddr2.SetFocus
        Cancel = True
        Exit Sub
     End If
   End If
End Sub

'Add by Amy 2020/03/13 從 aacc_sav搬過來修改
Public Sub Frmacc4130_Save()
    Dim strSql As String

On Error GoTo Checking
 
      If strSaveConfirm = MsgText(4) Then
         strSql = "select a0811, a0812 from acc080 where a0801 = '" & Text1 & "'"
         If CheckRecord(strSql, IIf(IsNull(Adodc1.Recordset.Fields("a0811").Value), 0, Adodc1.Recordset.Fields("a0811").Value), IIf(IsNull(Adodc1.Recordset.Fields("a0812").Value), 0, Adodc1.Recordset.Fields("a0812").Value)) = False Then
            strControlButton = MsgText(602)
            Text1.SetFocus
            Exit Sub
         End If
      End If
      '公司別
      If Text1 = MsgText(601) Then
         MsgBox MsgText(10) & Label3, , MsgText(5)
         strControlButton = MsgText(602)
         Text1.SetFocus
         Exit Sub
      Else
         '負責人
         If Text2 <> MsgText(601) Then
            If ExistCheck("staff", "st01", Text2, Label12) = False Then
               strControlButton = MsgText(602)
               Text2.SetFocus
               Exit Sub
            End If
         End If
         If Text10 <> MsgText(601) Then
            If UnionCode(Text10) = MsgText(603) Then
               MsgBox Label13 & MsgText(63), , MsgText(5)
               strControlButton = MsgText(602)
               Text10.SetFocus
               Exit Sub
            End If
         End If
         If CheckLen(Label1, Text3, 30) = MsgText(603) Then
            strControlButton = MsgText(602)
            Text3.SetFocus
            Exit Sub
         End If
         If CheckLen(Label6, Text6, 70) = MsgText(603) Then
            strControlButton = MsgText(602)
            Text6.SetFocus
            Exit Sub
         End If
         'Add by Amy 2020/03/13 +公司簡稱
          If CheckLen(Label15, Text14, Text14.MaxLength) = MsgText(603) Then
            strControlButton = MsgText(602)
            Text14.SetFocus
            Exit Sub
         End If
      End If
      'Add by Amy 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
      If PUB_ChkUniText(Me) = False Then
        strControlButton = MsgText(602)
        Exit Sub
      End If
      
      adoacc080.CursorLocation = adUseClient
      adoacc080.Open "select * from acc080 where a0801 = '" & Text1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If strSaveConfirm = MsgText(3) Then
         If adoacc080.RecordCount <> 0 Then
            MsgBox MsgText(9), , MsgText(5)
            strControlButton = MsgText(602)
            Text1.SetFocus
            Exit Sub
         End If
         adoacc080.AddNew
      Else
         If strSaveConfirm = MsgText(4) Then
            If adoacc080.RecordCount = 0 Then
               MsgBox MsgText(28), , MsgText(5)
               strControlButton = MsgText(602)
               Text1.SetFocus
               Exit Sub
            End If
         End If
      End If
      adoacc080.Fields("a0801").Value = Text1
      If Text2 <> MsgText(601) Then
         adoacc080.Fields("a0806").Value = Text2
      Else
         adoacc080.Fields("a0806").Value = Null
      End If
      If Text3 <> MsgText(601) Then
         adoacc080.Fields("a0802").Value = Text3
      Else
         adoacc080.Fields("a0802").Value = Null
      End If
      If Text4 <> MsgText(601) Then
         adoacc080.Fields("a0803").Value = Text4
      Else
         adoacc080.Fields("a0803").Value = Null
      End If
      If Text6 <> MsgText(601) Then
         adoacc080.Fields("a0804").Value = Text6
      Else
         adoacc080.Fields("a0804").Value = Null
      End If
      If Text10 <> MsgText(601) Then
         adoacc080.Fields("a0807").Value = Text10
      Else
         adoacc080.Fields("a0807").Value = Null
      End If
      If Text11 <> MsgText(601) Then
         adoacc080.Fields("a0809").Value = Text11
      Else
         adoacc080.Fields("a0809").Value = Null
      End If
      If Text7 <> MsgText(601) Then
         adoacc080.Fields("a0808").Value = Text7
      Else
         adoacc080.Fields("a0808").Value = Null
      End If
      If Text8 <> MsgText(601) Then
         adoacc080.Fields("a0805").Value = Text8
      Else
         adoacc080.Fields("a0805").Value = Null
      End If
      If Text9 <> MsgText(601) Then
         adoacc080.Fields("a0813").Value = Text9
      Else
         adoacc080.Fields("a0813").Value = Null
      End If
      '2008/12/15 CANCEL BY SONIA 與婧瑄確認不需此欄故取消
      'If .Text12 <> MsgText(601) Then
      '   .adoacc080.Fields("a0814").Value = .Text12
      'Else
      '   .adoacc080.Fields("a0814").Value = Null
      'End If
      '2008/12/15 END
      
      'Added by Morgan 2013/3/13
      If Text12 <> MsgText(601) Then
         adoacc080.Fields("a0821").Value = Text12
      Else
         adoacc080.Fields("a0821").Value = Null
      End If
      'end 2013/3/13
      
      'Added by Lydia 2017/09/06 英文地址1,2
      If txtAddr1 <> MsgText(601) Then
         adoacc080.Fields("a0822").Value = PUB_StringFilter(txtAddr1)
      Else
         adoacc080.Fields("a0822").Value = Null
      End If
      If txtAddr2 <> MsgText(601) Then
         adoacc080.Fields("a0823").Value = PUB_StringFilter(txtAddr2)
      Else
         adoacc080.Fields("a0823").Value = Null
      End If
      'end 2017/09/06
      
      'Added by Morgan 2017/10/26
      If Text13 <> MsgText(601) Then
         adoacc080.Fields("a0824").Value = Text13
      Else
         adoacc080.Fields("a0824").Value = Null
      End If
      'end 2017/10/26
      
      'Add by Amy 2020/03/13
      '公司簡稱
      If Text14 <> MsgText(601) Then
         adoacc080.Fields("a0820").Value = Text14
      Else
         adoacc080.Fields("a0820").Value = Null
      End If
      '作帳公司(M51用)
      If Text15.Visible = True Then
        If Text15 <> MsgText(601) Then
            adoacc080.Fields("a0827").Value = Text15
        Else
            adoacc080.Fields("a0827").Value = Null
        End If
      End If
      
      If strSaveConfirm = MsgText(4) Then
         adoacc080.Fields("a0810").Value = strUserNum
         adoacc080.Fields("a0811").Value = Val(strSrvDate(2))
         adoacc080.Fields("a0812").Value = ServerTime
      End If
      adoacc080.UpdateBatch
      adoacc080.Close
      AdodcRefresh
      RecordShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
 
End Sub
