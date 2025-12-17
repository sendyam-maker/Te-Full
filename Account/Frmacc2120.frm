VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc2120 
   AutoRedraw      =   -1  'True
   Caption         =   "暫收款作業"
   ClientHeight    =   5472
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   8784
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5472
   ScaleWidth      =   8784
   Begin VB.OptionButton Option2 
      Caption         =   "全部"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Left            =   5076
      TabIndex        =   30
      Top             =   1791
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "未沖"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Left            =   4368
      TabIndex        =   29
      Top             =   1791
      Width           =   720
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   4068
      TabIndex        =   4
      Top             =   627
      Width           =   1596
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   2760
      Picture         =   "Frmacc2120.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   255
      Width           =   350
   End
   Begin VB.ComboBox Combo1 
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
      TabIndex        =   8
      Top             =   1356
      Width           =   1452
   End
   Begin VB.TextBox Text13 
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
      Left            =   6840
      TabIndex        =   27
      Top             =   1356
      Width           =   1572
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
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   10
      Top             =   1728
      Width           =   612
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2120.frx":0102
      Height          =   2295
      Left            =   240
      TabIndex        =   12
      Top             =   2910
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   4043
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "a1201"
         Caption         =   "暫收款單號"
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
         DataField       =   "a1203"
         Caption         =   "代理人"
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
         DataField       =   "a1202"
         Caption         =   "暫收款日期"
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
      BeginProperty Column03 
         DataField       =   "a1204"
         Caption         =   "幣別"
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
         DataField       =   "a1205"
         Caption         =   "匯率(NT)"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "a1207"
         Caption         =   "外幣金額"
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
      BeginProperty Column06 
         DataField       =   "a1208"
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
      BeginProperty Column07 
         DataField       =   "a1211"
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
            ColumnWidth     =   1488.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1476.284
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1284.095
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1679.811
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1572.095
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   4944.189
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text11 
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
      Left            =   6840
      TabIndex        =   24
      Top             =   1728
      Width           =   1572
   End
   Begin VB.TextBox Text9 
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
      Left            =   4080
      MaxLength       =   12
      TabIndex        =   9
      Top             =   1356
      Width           =   1572
   End
   Begin VB.TextBox Text8 
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
      MaxLength       =   6
      TabIndex        =   6
      Top             =   984
      Width           =   1452
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
      Left            =   6840
      MaxLength       =   14
      TabIndex        =   7
      Top             =   984
      Width           =   1572
   End
   Begin VB.TextBox Text5 
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
      Left            =   6840
      MaxLength       =   13
      TabIndex        =   5
      Top             =   612
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4080
      MaxLength       =   9
      TabIndex        =   2
      Top             =   240
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
      Top             =   240
      Width           =   1215
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   330
      Left            =   1560
      TabIndex        =   3
      Top             =   612
      Width           =   1452
      _ExtentX        =   2561
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   90
      Top             =   2760
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
   Begin MSForms.TextBox Text10 
      Height          =   555
      Left            =   1560
      TabIndex        =   11
      Top             =   2100
      Width           =   6855
      VariousPropertyBits=   -1467989989
      ScrollBars      =   2
      Size            =   "12091;979"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text6 
      Height          =   330
      Left            =   3030
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   984
      Width           =   2655
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "4683;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text3 
      Height          =   330
      Left            =   5670
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   240
      Width           =   2775
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "4895;582"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "收款類別"
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
      TabIndex        =   28
      Top             =   1395
      Width           =   972
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "台幣金額"
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
      TabIndex        =   26
      Top             =   1395
      Width           =   972
   End
   Begin VB.Label Label11 
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
      Height          =   252
      Left            =   360
      TabIndex        =   25
      Top             =   2100
      Width           =   972
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4920
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   2640
      Left            =   255
      Top             =   105
      Width           =   8295
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "處理單號"
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
      TabIndex        =   23
      Top             =   1767
      Width           =   972
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "(1. 暫收 2.溢收轉入)"
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
      Left            =   2280
      TabIndex        =   22
      Top             =   1767
      Width           =   3492
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "暫收款類別"
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
      TabIndex        =   21
      Top             =   1767
      Width           =   1212
   End
   Begin VB.Label Label8 
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
      Height          =   252
      Left            =   3120
      TabIndex        =   20
      Top             =   1395
      Width           =   972
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "外幣金額"
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
      TabIndex        =   19
      Top             =   1023
      Width           =   972
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "會計科目"
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
      TabIndex        =   18
      Top             =   1023
      Width           =   972
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "匯率(NT)"
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
      TabIndex        =   17
      Top             =   651
      Width           =   972
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "幣別"
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
      Left            =   3240
      TabIndex        =   16
      Top             =   650
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "暫收款日期"
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
      TabIndex        =   15
      Top             =   651
      Width           =   1212
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   278
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "暫收款單號"
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
      TabIndex        =   13
      Top             =   279
      Width           =   1212
   End
End
Attribute VB_Name = "Frmacc2120"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/03 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Text3、Text6、Text10
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoadodc1 As New ADODB.Recordset
Public adoacc1p0 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public strDocNo As String

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
   If Combo2 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc1y0", "a1y01", Combo2, Label4) = False Then
      Cancel = True
      Combo2.SetFocus
   End If
End Sub

Private Sub Command3_Click()
   If Adodc1.Recordset.RecordCount = 0 Or Text2 = MsgText(601) Then
      Exit Sub
   End If
   Adodc1.Recordset.Find "a1201 = '" & Text2 & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      FormShow
      AdodcRefresh
      RecordShow
   Else
      MsgBox MsgText(33), , MsgText(5)
      Adodc1.Recordset.MoveFirst
   End If
End Sub

Private Sub Command3_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command3_Click
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Command3.Enabled Then
      FormShow
   End If
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.MoveFirst
   End If
   Adodc1.Recordset.Find "a1201 = '" & strItemNo & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      FormShow
      RecordShow
   End If
   strItemNo = MsgText(601)
End Sub

'Added by Lydia 2021/12/03
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Added by Lydia 2021/12/03 Form2.0 記錄鍵盤傳入順序
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call PUB_SaveTrackMode(1, KeyCode)  'Added by Lydia 2021/12/03 Form2.0 記錄鍵盤傳入順序
   
   KeyEnter KeyCode
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
   'Modify by Amy 2023/08/18 W8850 H5700
   PUB_InitForm Me, 8880, 5920, strBackPicPath1
   'end 2021/12/07
   Combo1.AddItem ComboItem(51)
   Combo1.AddItem ComboItem(52)
   Combo1.AddItem ComboItem(53)
   Combo1.AddItem ComboItem(54)
   Combo1.AddItem ComboItem(55)
   MaskEdBox1.Mask = DFormat
   
   OpenTable
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.MoveLast
      Adodc1.Recordset.MoveFirst
      RecordShow
   End If
   
   Option1.Value = True
   FormEnable
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   strTrackMode = "" 'Added by Lydia 2021/12/03 Form2.0 記錄鍵盤傳入順序(清除)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc2120 = Nothing
End Sub

Private Sub MaskEdBox1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox Label3 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label3 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If strSaveConfirm <> MsgText(3) Then
      Exit Sub
   End If
   If Mid(MaskEdBox1.Text, 1, 3) <> Mid(CFDate(ACDate(ServerDate)), 1, 3) Then
      Text2 = UpdateNo("acc120", "a1201", 5, MaskEdBox1.Text, MsgText(809))
   Else
      'Text2 = AutoNo(MsgText(809), 5)
      Text2 = strDocNo
   End If
End Sub

Private Sub Option1_Click()
   AdodcRefresh
   If Adodc1.Recordset.RecordCount <> 0 Then
        Adodc1.Recordset.MoveFirst
        RecordShow
    End If
End Sub

Private Sub Option2_Click()
   AdodcRefresh
   If Adodc1.Recordset.RecordCount <> 0 Then
        Adodc1.Recordset.MoveFirst
       RecordShow
    End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   
   adoadodc1.CursorLocation = adUseClient
   
   'Removed by Morgan 2017/9/21 此處不必抓資料 AdodcRefresh 抓就好
   'If Option1.Value Then
   '   '92.6.16 modify by sonia
   '   'adoadodc1.Open "select * from acc120 where a1201 not in (select a1303 from acc130) and a1201 not in (select a1p30 from acc1p0 where a1p02 = 'F' and a1p05 = '2401' and a1p07 <> 0 and a1p30 is not null) order by a1201 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '   adoadodc1.Open "select * from acc120 where a1201 not in (select a1303 from acc130) and a1201 not in (select a1p30 from acc1p0 where a1p02 in ('F','K', 'I') and a1p05 = '2401' and a1p07 <> 0 and a1p30 is not null) order by a1201 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Else
   '   adoadodc1.Open "select * from acc120 order by a1201 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'End If
   adoadodc1.Open "select * from acc120 where rownum<1 order by a1201 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'end 2017/9/21
   Set Adodc1.Recordset = adoadodc1
   
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from acc1y0 order by a1y01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoquery.EOF = False
      Combo2.AddItem adoquery.Fields("a1y01").Value
      adoquery.MoveNext
   Loop
   adoquery.Close
   Combo2 = "USD"
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
   If Option1.Value Then
      '92.6.16 modify by sonia
      'adoadodc1.Open "select * from acc120 where a1201 not in (select a1303 from acc130) and a1201 not in (select a1p30 from acc1p0 where a1p02 = 'F' and a1p05 = '2401' and a1p07 <> 0 and a1p30 is not null) order by a1201 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      'Modified by Morgan 2017/9/21 太慢改語法
      'adoadodc1.Open "select * from acc120 where a1201 not in (select a1303 from acc130) and a1201 not in (select a1p30 from acc1p0 where a1p02 in ('F','K', 'I') and a1p05 = '2401' and a1p07 <> 0 and a1p30 is not null) order by a1201 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      adoadodc1.Open "select * from acc120 where not exists(select * from acc130 where a1303=a1201) and not exists( select * from acc1p0 where a1p30=a1201 and a1p02||'' in ('F','K', 'I') and a1p05||'' = '2401' and a1p07 <> 0) order by a1201 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Else
      adoadodc1.Open "select * from acc120 order by a1201 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   End If
   
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.Find "a1201 = '" & Text2 & "'", 0, adSearchForward, 1
      If Adodc1.Recordset.EOF = False Then
         FormShow
         RecordShow
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
   Text2 = Adodc1.Recordset.Fields("a1201").Value
   If IsNull(Adodc1.Recordset.Fields("a1203").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = Adodc1.Recordset.Fields("a1203").Value
   End If
   Text3 = FagentQuery(Text1, 2)
   If Text3 = MsgText(601) Then
      Text3 = FagentQuery(Text1, 1)
   End If
   If Text3 = MsgText(601) Then
      Text3 = FagentQuery(Text1, 3)
   End If
   If Text3 = MsgText(601) Then
      Text3 = CustomerQuery(Text1, 2)
   End If
   If Text3 = MsgText(601) Then
      Text3 = CustomerQuery(Text1, 1)
   End If
   If Text3 = MsgText(601) Then
      Text3 = CustomerQuery(Text1, 3)
   End If
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(Adodc1.Recordset.Fields("a1202").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(Adodc1.Recordset.Fields("a1202").Value)
   End If
   MaskEdBox1.Mask = DFormat
   If IsNull(Adodc1.Recordset.Fields("a1204").Value) Then
      Combo2 = MsgText(601)
   Else
      Combo2 = Adodc1.Recordset.Fields("a1204").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1205").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = Adodc1.Recordset.Fields("a1205").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1206").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = Adodc1.Recordset.Fields("a1206").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1207").Value) Then
      Text7 = MsgText(601)
   Else
      Text7 = Adodc1.Recordset.Fields("a1207").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1220").Value) Then
      Combo1 = MsgText(601)
   Else
      Combo1 = Combo1.List(Val(Adodc1.Recordset.Fields("a1220").Value) - 1)
   End If
   If IsNull(Adodc1.Recordset.Fields("a1208").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = Adodc1.Recordset.Fields("a1208").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1209").Value) Then
      Text12 = MsgText(601)
   Else
      Text12 = Adodc1.Recordset.Fields("a1209").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1210").Value) Then
      Text11 = MsgText(601)
   Else
      Text11 = Adodc1.Recordset.Fields("a1210").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("a1211").Value) Then
      Text10 = MsgText(601)
   Else
      Text10 = Adodc1.Recordset.Fields("a1211").Value
   End If
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
   If Text1 <> MsgText(601) Then
      If ExistCheck("fagent", "fa01", Mid(Text1, 1, 8), Label2, False) = False Then
         If ExistCheck("customer", "cu01", Mid(Text1, 1, 8), Label2, False) = False Then
            MsgBox MsgText(45) & Label2, , MsgText(5)
            Cancel = True
            Text1.SetFocus
            Exit Sub
         End If
      End If
   End If
   Text3 = FagentQuery(Text1, 2)
   If Text3 = MsgText(601) Then
      Text3 = FagentQuery(Text1, 1)
   End If
   If Text3 = MsgText(601) Then
      Text3 = FagentQuery(Text1, 3)
   End If
   If Text3 = MsgText(601) Then
      Text3 = CustomerQuery(Text1, 2)
   End If
   If Text3 = MsgText(601) Then
      Text3 = CustomerQuery(Text1, 1)
   End If
   If Text3 = MsgText(601) Then
      Text3 = CustomerQuery(Text1, 3)
   End If
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
   'edit by nickc 2007/06/11  切換輸入法改用API
   OpenIme
End Sub

'Modified by Lydia 2021/12/03 改成Form 2.0; KeyCode As Integer=>MSForms.ReturnInteger
Private Sub Text10_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
  
   'Modified by Lydia 2021/12/03 +val()
   KeyEnter Val(KeyCode)
End Sub

Private Sub Text10_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
CloseIme
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub

Private Sub Text12_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Text5_Change()
   If Text5 = MsgText(601) Then
      Exit Sub
   End If
   Text13 = Format(Val(Text5) * Val(Text7), FAmount)
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Text7_Change()
   If Text7 = MsgText(601) Then
      Exit Sub
   End If
   Text13 = Format(Val(Text5) * Val(Text7), FAmount)
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Text8_Change()
   If Text8 = MsgText(601) Then
      Text6 = MsgText(601)
      Exit Sub
   End If
   Text6 = A0102Query(Text8)
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub

Private Sub Text8_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Text8_Validate(Cancel As Boolean)
   If Text8 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc010", "a0101", Text8, Label6) = False Then
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text9_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = Adodc1.Recordset.Bookmark & MsgText(35) & Adodc1.Recordset.RecordCount
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   '2012/10/22 modify by sonia
   'Select Case Len(Text9)
   '   Case 7, 8, 9
   '      Text9 = Text9 & "000"
   'End Select
   If Text9 <> MsgText(601) Then
      Text9 = CaseNoZero(Text9)
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select pa01 as SystemNo from patent where pa01 = '" & Mid(Text9, 1, Len(Text9) - 9) & "' and pa02 = '" & Mid(Text9, Len(Text9) - 8, 6) & "' and pa03 = '" & Mid(Text9, Len(Text9) - 2, 1) & "' and pa04 = '" & Mid(Text9, Len(Text9) - 1, 2) & "' union " & _
                    "select tm01 as SystemNo from trademark where tm01 = '" & Mid(Text9, 1, Len(Text9) - 9) & "' and tm02 = '" & Mid(Text9, Len(Text9) - 8, 6) & "' and tm03 = '" & Mid(Text9, Len(Text9) - 2, 1) & "' and tm04 = '" & Mid(Text9, Len(Text9) - 1, 2) & "' union " & _
                    "select lc01 as SystemNo from lawcase where lc01 = '" & Mid(Text9, 1, Len(Text9) - 9) & "' and lc02 = '" & Mid(Text9, Len(Text9) - 8, 6) & "' and lc03 = '" & Mid(Text9, Len(Text9) - 2, 1) & "' and lc04 = '" & Mid(Text9, Len(Text9) - 1, 2) & "' union " & _
                    "select hc01 as SystemNo from hirecase where hc01 = '" & Mid(Text9, 1, Len(Text9) - 9) & "' and hc02 = '" & Mid(Text9, Len(Text9) - 8, 6) & "' and hc03 = '" & Mid(Text9, Len(Text9) - 2, 1) & "' and hc04 = '" & Mid(Text9, Len(Text9) - 1, 2) & "' union " & _
                    "select sp01 as SystemNo from servicepractice where sp01 = '" & Mid(Text9, 1, Len(Text9) - 9) & "' and sp02 = '" & Mid(Text9, Len(Text9) - 8, 6) & "' and sp03 = '" & Mid(Text9, Len(Text9) - 2, 1) & "' and sp04 = '" & Mid(Text9, Len(Text9) - 1, 2) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount = 0 Then
         MsgBox MsgText(28) & Label8, , MsgText(5)
         Cancel = True
         adoquery.Close
         Exit Sub
      End If
      adoquery.Close
   End If
   Exit Sub
Checking:
   MsgBox MsgText(128), , MsgText(5)
   Exit Sub
End Sub
'Add by Morgan 2011/3/10
Private Sub SetControl(bolOpenState As Boolean)
   Dim oControl As Control
   For Each oControl In Me.Controls
      'Debug.Print TypeName(oControl)
      If TypeName(oControl) = "TextBox" Then
         If oControl.Enabled = True Then
            oControl.Locked = Not bolOpenState
         End If
      End If
      If TypeName(oControl) = "ComboBox" Then
         oControl.Enabled = bolOpenState
      End If

      If TypeName(oControl) = "MaskEdBox" Then
         oControl.Enabled = bolOpenState
      End If
   Next
End Sub

'Add by Morgan 2011/3/10
Public Sub FormEnable(Optional ByVal strState As String)
   
   '新增
   If strState = MsgText(3) Then
      SetControl True
      Text2.Locked = True
      Command3.Enabled = False
      
   '修改
   ElseIf strState = MsgText(4) Then
   
      If Text12 = "2" Then
         SetControl False
         Text10.Locked = False
      Else
         SetControl True
        'Add by Amy 2014/11/03  +a1p22有值不可修改入帳日
        If CheckExistA1p22("1", "G", Text2) = True Then
            MaskEdBox1.Enabled = False
        End If
        'end 2014/11/03
      End If
      Text2.Locked = True
      Command3.Enabled = False
   Else
   
      SetControl False
      Text2.Locked = False
      Command3.Enabled = True
   End If
End Sub

'Add by Amy 2014/11/03 由aacc_sav搬回
Public Sub Frmacc2120_Save()
Dim strAccNo As String
Dim strMsg As String 'Add by Amy 2014/11/03

   'Added by Lydia 2021/12/03 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
      strControlButton = MsgText(602)
      Exit Sub
   End If
   'end 2021/12/03
   
   On Error GoTo Checking
   With Frmacc2120
   
      If .Text2 = MsgText(601) Then
         MsgBox MsgText(10), , MsgText(5)
         strControlButton = MsgText(602)
         .Text2.SetFocus
         Exit Sub
      'Added by Lydia 2018/07/19 檢查幣別(ex.N10600046沒有輸入幣別,人工補USD)
      ElseIf Trim(.Combo2.Text) = MsgText(601) Then
         MsgBox MsgText(10), , MsgText(5)
         strControlButton = MsgText(602)
         .Combo2.SetFocus
         Exit Sub
      'end 2018/07/19
      Else
         If .Text1 <> "" Then
            If ExistCheck("fagent", "fa01", Mid(.Text1, 1, 8), .Label2, False) = False Then
               If ExistCheck("customer", "cu01", Mid(.Text1, 1, 8), .Label2, False) = False Then
                  MsgBox MsgText(45) & .Label2, , MsgText(5)
                  strControlButton = MsgText(602)
                  .Text1.SetFocus
                  Exit Sub
               End If
            End If
         End If
         If .Text8.Locked = False Then 'Add by Morgan 2011/3/16
            If ExistCheck("acc010", "a0101", .Text8, .Label6) = False Then
               strControlButton = MsgText(602)
               .Text8.SetFocus
               Exit Sub
            End If
         End If
         If .MaskEdBox1.Enabled = True Then 'Add by Morgan 2011/3/16
            If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
               MsgBox .Label3 & MsgText(52), , MsgText(5)
               strControlButton = MsgText(602)
               .MaskEdBox1.SetFocus
               Exit Sub
            Else
               If DateCheck(.MaskEdBox1.Text) = MsgText(603) Then
                  MsgBox .Label3 & MsgText(63), , MsgText(5)
                  strControlButton = MsgText(602)
                  .MaskEdBox1.SetFocus
                  Exit Sub
               End If
               'Add by Amy 2014/11/03 +系統日期檢查
               If ChkWorkData("1", DBDATE(MaskEdBox1.Text), strMsg) = False Then
                    MsgBox Label3 & strMsg, , MsgText(5)
                    strControlButton = MsgText(602)
                    MaskEdBox1.SetFocus
                    Exit Sub
                End If
                'end 2014/11/03
            End If
         End If
      End If
      If strSaveConfirm = MsgText(3) Then
         If .Adodc1.Recordset.RecordCount <> 0 Then
            .Adodc1.Recordset.Find "a1201 = '" & .Text2 & "'", 0, adSearchForward, 1
            If .Adodc1.Recordset.EOF = False Then
               MsgBox MsgText(9), , MsgText(5)
               strControlButton = MsgText(602)
               .Text2.SetFocus
               Exit Sub
            End If
         End If
         .Adodc1.Recordset.AddNew
      End If
      .Adodc1.Recordset.Fields("a1201").Value = .Text2
      If .Text1 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a1203").Value = .Text1
      Else
         .Adodc1.Recordset.Fields("a1203").Value = Null
      End If
      If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
         .Adodc1.Recordset.Fields("a1202").Value = Val(FCDate(.MaskEdBox1.Text))
      Else
         .Adodc1.Recordset.Fields("a1202").Value = Null
      End If
      If .Combo2 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a1204").Value = .Combo2
      Else
         .Adodc1.Recordset.Fields("a1204").Value = Null
      End If
      If .Text5 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a1205").Value = Val(.Text5)
      Else
         .Adodc1.Recordset.Fields("a1205").Value = 0
      End If
      If .Text8 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a1206").Value = .Text8
      Else
         .Adodc1.Recordset.Fields("a1206").Value = Null
      End If
      If .Text7 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a1207").Value = Val(.Text7)
      Else
         .Adodc1.Recordset.Fields("a1207").Value = 0
      End If
      If .Combo1 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a1220").Value = Mid(.Combo1, 1, 1)
      Else
         .Adodc1.Recordset.Fields("a1220").Value = Null
      End If
      If .Text9 <> MsgText(601) Then
         .Text9 = CaseNoZero(.Text9)    'add by sonia 2017/9/27 N10600202未補零,導致後續結匯408645474Y52292000錯誤
         .Adodc1.Recordset.Fields("a1208").Value = .Text9
      Else
         .Adodc1.Recordset.Fields("a1208").Value = Null
      End If
      If .Text12 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a1209").Value = .Text12
      Else
         .Adodc1.Recordset.Fields("a1209").Value = Null
      End If
      If .Combo2 = MsgText(601) And .Text7 = MsgText(601) Then
         .Adodc1.Recordset.Fields("a1211").Value = Null
      Else
         .Adodc1.Recordset.Fields("a1211").Value = .Combo2 & " " & .Text7
      End If
      If .Text10 <> MsgText(601) Then
         .Adodc1.Recordset.Fields("a1211").Value = .Text10
      Else
         .Adodc1.Recordset.Fields("a1211").Value = Null
      End If
      If strSaveConfirm = MsgText(3) Then
         .Adodc1.Recordset.Fields("a1214").Value = Val(strSrvDate(2))
         .Adodc1.Recordset.Fields("a1215").Value = ServerTime
         .Adodc1.Recordset.Fields("a1216").Value = strUserNum
      Else
         .Adodc1.Recordset.Fields("a1217").Value = Val(strSrvDate(2))
         .Adodc1.Recordset.Fields("a1218").Value = ServerTime
         .Adodc1.Recordset.Fields("a1219").Value = strUserNum
      End If
      
      'Add by Morgan 2011/3/10 修改備註
      If strSaveConfirm = MsgText(4) And .Text12 = "2" Then GoTo flgUpdateBatch
      
      .adoquery.CursorLocation = adUseClient
      .adoquery.Open "select distinct a1p22 from acc1p0 where a1p01 = '1' and a1p02 = 'G' and a1p04 = '" & .Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If .adoquery.RecordCount <> 0 Then
         If IsNull(.adoquery.Fields("a1p22").Value) = False Then
            strAccNo = .adoquery.Fields("a1p22").Value
         Else
            strAccNo = ""
         End If
      Else
         strAccNo = ""
      End If
      .adoquery.Close
      adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'G' and a1p04 = '" & .Text2 & "'"
      If .Text8 <> MsgText(601) Then
'借方------------------------------------------------
         .adoacc1p0.CursorLocation = adUseClient
         .adoacc1p0.Open "select * from acc1p0 where a1p01 = '1' and a1p02 = 'G' and a1p04 = '" & .Text2 & "' and a1p05 = '" & .Text8 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
         If .adoacc1p0.RecordCount = 0 Then
            .adoacc1p0.AddNew
         End If
         .adoacc1p0.Fields("a1p01").Value = "1"
         .adoacc1p0.Fields("a1p02").Value = "G"
         .adoacc1p0.Fields("a1p03").Value = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'G' and a1p04 = '" & .Text2 & "'", 3)
         .adoacc1p0.Fields("a1p04").Value = .Text2
         .adoacc1p0.Fields("a1p05").Value = .Text8
         .adoacc1p0.Fields("a1p06").Value = MsgText(55)
         '2009/1/17 add by sonia 婧瑄說借貸方都放對沖客戶
         If .Text1 <> MsgText(601) Then
            .adoacc1p0.Fields("a1p15").Value = .Text1
         Else
            .adoacc1p0.Fields("a1p15").Value = Null
         End If
         '2009/1/17 end
         'modify by sonia 2021/1/27 加傳本所案號以判別FCP,FCT英日文組
         'If AccNoToSalesNo(.Text8) = "" Then
         If AccNoToSalesNo(.Text8, .Text9) = "" Then
            .adoacc1p0.Fields("a1p16").Value = Null
         Else
            'modify by sonia 2021/1/27 加傳本所案號以判別FCP,FCT英日文組
            '.adoacc1p0.Fields("a1p16").Value = AccNoToSalesNo(.Text8)
            .adoacc1p0.Fields("a1p16").Value = AccNoToSalesNo(.Text8, .Text9)
         End If
         If .Combo2 <> MsgText(601) Then
            .adoacc1p0.Fields("a1p19").Value = .Combo2
         Else
            .adoacc1p0.Fields("a1p19").Value = Null
         End If
         If .Text5 <> MsgText(601) Then
            .adoacc1p0.Fields("a1p20").Value = Val(.Text5)
         Else
            .adoacc1p0.Fields("a1p20").Value = 0
         End If
         If .Text7 <> MsgText(601) Then
            .adoacc1p0.Fields("a1p21").Value = Val(.Text7)
            .adoacc1p0.Fields("a1p07").Value = Val(Format(Val(.Text7) * Val(.Text5), FAmount))
         Else
            .adoacc1p0.Fields("a1p21").Value = 0
            .adoacc1p0.Fields("a1p07").Value = 0
         End If
         .adoacc1p0.Fields("a1p08").Value = 0
         If strAccNo <> "" Then
            .adoacc1p0.Fields("a1p22").Value = strAccNo
         Else
            .adoacc1p0.Fields("a1p22").Value = Null
         End If
         If .Combo1 <> MsgText(601) Then
            .adoacc1p0.Fields("a1p24").Value = Mid(.Combo1, 1, 1)
         Else
            .adoacc1p0.Fields("a1p24").Value = Null
         End If
         If .Text12 <> MsgText(601) Then
            .adoacc1p0.Fields("a1p26").Value = .Text12
         Else
            .adoacc1p0.Fields("a1p26").Value = Null
         End If
         If .Text11 <> MsgText(601) Then
            .adoacc1p0.Fields("a1p23").Value = .Text11
         Else
            .adoacc1p0.Fields("a1p23").Value = Null
         End If
         'If .Text10 <> MsgText(601) Then
            .adoacc1p0.Fields("a1p14").Value = .Combo2 & " " & Format(.Text7, FDollar)
         'Else
         '   .adoacc1p0.Fields("a1p14").Value = Null
         'End If
         If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
            .adoacc1p0.Fields("a1p18").Value = Val(FCDate(.MaskEdBox1.Text))
         Else
            .adoacc1p0.Fields("a1p18").Value = Null
         End If
         If IsNull(.adoacc1p0.Fields("a1p22").Value) = False Then
            .adoacc1p0.Fields("a1p27").Value = MsgText(602)
         End If
         If strSaveConfirm <> MsgText(3) And IsNull(.adoacc1p0.Fields("a1p27").Value) = False Then
            .adoacc1p0.Fields("a1p27").Value = MsgText(602)
         End If
        'Add By Cheng 2004/05/11
        '將暫收款單號存入對沖代號--其他
         .adoacc1p0.Fields("a1p30").Value = .Text2
        'End
         .adoacc1p0.UpdateBatch
         .adoacc1p0.Close
'貸方------------------------------------------------
         .adoacc1p0.CursorLocation = adUseClient
         'modify by sonia 2013/12/25 加貸方金額>0條件 N10200468
         '.adoacc1p0.Open "select * from acc1p0 where a1p01 = '1' and a1p02 = 'G' and a1p04 = '" & .Text2 & "' and a1p05 = '2401'", adoTaie, adOpenDynamic, adLockBatchOptimistic
         .adoacc1p0.Open "select * from acc1p0 where a1p01 = '1' and a1p02 = 'G' and a1p04 = '" & .Text2 & "' and a1p05 = '2401' and a1p08>0", adoTaie, adOpenDynamic, adLockBatchOptimistic
         If .adoacc1p0.RecordCount = 0 Then
            .adoacc1p0.AddNew
         End If
         .adoacc1p0.Fields("a1p01").Value = "1"
         .adoacc1p0.Fields("a1p02").Value = "G"
         .adoacc1p0.Fields("a1p03").Value = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '1' and a1p02 = 'G' and a1p04 = '" & .Text2 & "'", 3)
         .adoacc1p0.Fields("a1p04").Value = .Text2
         .adoacc1p0.Fields("a1p05").Value = "2401"
         .adoacc1p0.Fields("a1p06").Value = MsgText(55)
         '2009/1/17 add by sonia 婧瑄說借貸方都放對沖客戶
         If .Text1 <> MsgText(601) Then
            .adoacc1p0.Fields("a1p15").Value = .Text1
         Else
            .adoacc1p0.Fields("a1p15").Value = Null
         End If
         '2009/1/17 end
         If .Combo2 <> MsgText(601) Then
            .adoacc1p0.Fields("a1p19").Value = .Combo2
         Else
            .adoacc1p0.Fields("a1p19").Value = Null
         End If
         If .Text5 <> MsgText(601) Then
            .adoacc1p0.Fields("a1p20").Value = Val(.Text5)
         Else
            .adoacc1p0.Fields("a1p20").Value = 0
         End If
         If .Text7 <> MsgText(601) Then
            .adoacc1p0.Fields("a1p21").Value = Val(.Text7)
            .adoacc1p0.Fields("a1p08").Value = Val(Format(Val(.Text7) * Val(.Text5), FAmount))
         Else
            .adoacc1p0.Fields("a1p21").Value = 0
            .adoacc1p0.Fields("a1p08").Value = 0
         End If
         .adoacc1p0.Fields("a1p07").Value = 0
         If strAccNo <> "" Then
            .adoacc1p0.Fields("a1p22").Value = strAccNo
         Else
            .adoacc1p0.Fields("a1p22").Value = Null
         End If
         If .Combo1 <> MsgText(601) Then
            .adoacc1p0.Fields("a1p24").Value = Mid(.Combo1, 1, 1)
         Else
            .adoacc1p0.Fields("a1p24").Value = Null
         End If
         If .Text12 <> MsgText(601) Then
            .adoacc1p0.Fields("a1p26").Value = .Text12
         Else
            .adoacc1p0.Fields("a1p26").Value = Null
         End If
         If .Text11 <> MsgText(601) Then
            .adoacc1p0.Fields("a1p23").Value = .Text11
         Else
            .adoacc1p0.Fields("a1p23").Value = Null
         End If
         'If .Text10 <> MsgText(601) Then
            '2011/11/9 modify by sonia 加暫收款單號
            '.adoacc1p0.Fields("a1p14").Value = .Text3 & "/" & .Combo2 & " " & Format(.Text7, FDollar)
            'modify by sonia 2018/5/18
            '.adoacc1p0.Fields("a1p14").Value = .Text3 & "/" & .Combo2 & " " & Format(.Text7, FDollar) & "/" & .Text2
            .adoacc1p0.Fields("a1p14").Value = .Text3 & "/" & .Combo2 & " " & Format(.Text7, FDollar) & "/" & .Text2 & "/" & Text10 & "/" & Text9
         'Else
         '   .adoacc1p0.Fields("a1p14").Value = Null
         'End If
         If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
            .adoacc1p0.Fields("a1p18").Value = Val(FCDate(.MaskEdBox1.Text))
         Else
            .adoacc1p0.Fields("a1p18").Value = Null
         End If
         If IsNull(.adoacc1p0.Fields("a1p22").Value) = False Then
            .adoacc1p0.Fields("a1p27").Value = MsgText(602)
         End If
         If strSaveConfirm <> MsgText(3) And IsNull(.adoacc1p0.Fields("a1p27").Value) = False Then
            .adoacc1p0.Fields("a1p27").Value = MsgText(602)
         End If
        'Add By Cheng 2004/05/11
        '將暫收款單號存入對沖代號--其他
         .adoacc1p0.Fields("a1p30").Value = .Text2
        'End
         .adoacc1p0.UpdateBatch
         .adoacc1p0.Close
      End If
      
flgUpdateBatch:

      .Adodc1.Recordset.UpdateBatch
      .AdodcRefresh
      .RecordShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   strControlButton = MsgText(602)
   MsgBox Err.Description, , MsgText(5)
   End With
End Sub
