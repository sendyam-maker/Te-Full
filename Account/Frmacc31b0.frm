VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc31b0 
   AutoRedraw      =   -1  'True
   Caption         =   "即期票存入作業"
   ClientHeight    =   5316
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5316
   ScaleWidth      =   8760
   Begin VB.TextBox Text1 
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
      Height          =   300
      Left            =   6810
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1020
      Width           =   1572
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1260
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   180
      Width           =   2375
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
      Height          =   300
      Left            =   864
      MaxLength       =   10
      TabIndex        =   16
      Top             =   4845
      Width           =   855
   End
   Begin VB.TextBox Text7 
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
      Left            =   3384
      MaxLength       =   10
      TabIndex        =   15
      Top             =   4845
      Width           =   1455
   End
   Begin VB.TextBox Text6 
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
      Height          =   315
      Left            =   1260
      TabIndex        =   14
      Top             =   1350
      Width           =   1572
   End
   Begin VB.TextBox Text4 
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
      Height          =   300
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1020
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
      Height          =   315
      Left            =   4650
      MaxLength       =   10
      TabIndex        =   1
      Top             =   180
      Width           =   1272
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc31b0.frx":0000
      Height          =   3105
      Left            =   240
      TabIndex        =   7
      Top             =   1710
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   5440
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   20
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
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "即期票存入資料"
      ColumnCount     =   7
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
         DataField       =   "a0g02"
         Caption         =   "收票銀行"
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
      BeginProperty Column02 
         DataField       =   "a0e07"
         Caption         =   "收票帳號"
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
      BeginProperty Column03 
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
      BeginProperty Column04 
         DataField       =   "a0e11"
         Caption         =   "票據金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
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
         DataField       =   "custname"
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
         Size            =   344
         BeginProperty Column00 
            ColumnWidth     =   1535.811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1607.811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1607.811
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1307.906
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   552.189
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   6600.189
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7920
      Picture         =   "Frmacc31b0.frx":0015
      Style           =   1  '圖片外觀
      TabIndex        =   6
      ToolTipText     =   "取消"
      Top             =   560
      Width           =   492
   End
   Begin VB.TextBox Text5 
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
      Height          =   315
      Left            =   1260
      MaxLength       =   8
      TabIndex        =   3
      Top             =   1020
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1260
      TabIndex        =   2
      Top             =   600
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
      Height          =   315
      Left            =   240
      Top             =   1620
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   550
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
   Begin MSForms.TextBox Text3 
      Height          =   315
      Left            =   5940
      TabIndex        =   12
      Top             =   180
      Width           =   2580
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      Size            =   "4542;556"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "收票帳號"
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
      Left            =   5850
      TabIndex        =   20
      Top             =   1020
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "票據金額"
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
      TabIndex        =   19
      Top             =   1380
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "筆數"
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
      Left            =   270
      TabIndex        =   18
      Top             =   4845
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "金額合計"
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
      TabIndex        =   17
      Top             =   4845
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   852
      Left            =   240
      Top             =   108
      Width           =   8292
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "收票銀行"
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
      Left            =   3120
      TabIndex        =   13
      Top             =   1020
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "存入銀行"
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
      Left            =   3720
      TabIndex        =   11
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "銀行帳號"
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
      Top             =   240
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   24
      Top             =   4248
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "票據號碼"
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
      TabIndex        =   9
      Top             =   1020
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "存入日期"
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
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc31b0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/07 Form2.0已修改 Text3/DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc0e0 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adocheck As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Dim dayDate As Date
Dim strRemark As String
Private Const strType As String = "99"
Public strA0E23 As String   '2014/1/24 add by sonia

Private Sub Command2_Click()
   AdodcDelete
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 8850
   Me.Height = 5760  'Modify by Amy 2023/08/16 原:5600
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'Add by Morgan 2006/7/21
   MaskEdBox1 = CFDate(strSrvDate(2))
   MaskEdBox1.Mask = DFormat
   PUB_SetAccount Combo1
   'end 2006/7/21
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc31b0 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   'Modify by Morgan 2004/11/1 加 and rownum<1
   adoadodc1.Open "select * from acc0e0 where a0e19 = '" & Text2 & "' and a0e20 = '" & Combo1 & "' and a0e18 = " & Val(FCDate(MaskEdBox1.Text)) & " and rownum<1 order by a0e21 asc, a0e48 desc", adoTaie, adOpenStatic, adLockReadOnly
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
Dim strUnion As String
Dim lngDate As Long

On Error GoTo Checking
   dayDate = Format(Val(FCDate(MaskEdBox1.Text)) + 19110000, "####/##/##")
   lngDate = ACDate(Format(dayDate + 3, "YYYYMMDD"))
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   'Modify by Morgan 2005/1/27 需考慮不是客戶的往來對象(Ex.銀行)
   'strUnion = "select a0e01, a0e02, a0e10, a0e11, a0g02, a0e20, a0e08, cu04 as contect, a0e21, a0e48 from acc0e0, acc0g0, customer where a0e01 = a0g01 and a0e05 = '1' and substr(a0e06, 1, 8) = cu01 (+) and substr(a0e06, 9, 1) = cu02 (+) and substr(a0e06, 9, 1) = decode(a0e05, '1', cu02) and a0e45 = '" & strType & "' and a0e19 = '" & Text2 & "' and a0e20 = '" & Combo1 & "' and a0e21 = " & Val(FCDate(MaskEdBox1.Text)) & " and (a0e14 is null or a0e14 = 0)"
   '2014/1/24 modify by sonia 加入公司別a0e23條件
   'Modify by Amy 2020/07/22 +a0e07 因改為key
   'Modify by Amy 2023/05/23 原:Combo1
   strExc(1) = Left(Combo1.Text, InStr(Combo1, " ") - 1)
   strUnion = "select a0e01, a0e02, a0e10, a0e11, a0g02, a0e20, a0e08, cu04 as custname, a0e21, a0e48,a0e07 from acc0e0, acc0g0, customer where a0e23='" & strA0E23 & "' and a0e01 = a0g01 and a0e05 = '1' and substr(a0e06, 1, 8) = cu01 (+) and substr(a0e06, 9, 1) = cu02 (+) and a0e45 = '" & strType & "' and a0e19 = '" & Text2 & "' and a0e20 = '" & strExc(1) & "' and a0e21 = " & Val(FCDate(MaskEdBox1.Text)) & " and (a0e14 is null or a0e14 = 0)"
   strUnion = strUnion & " union select a0e01, a0e02, a0e10, a0e11, a0g02, a0e20, a0e08, a0i02 as custname, a0e21, a0e48,a0e07 from acc0e0, acc0g0, acc0i0 where a0e23='" & strA0E23 & "' and a0e01 = a0g01 and a0e05 = '2' and a0e06 = a0i01 (+) and a0e45 = '" & strType & "' and a0e19 = '" & Text2 & "' and a0e20 = '" & strExc(1) & "' and a0e21 = " & Val(FCDate(MaskEdBox1.Text)) & " and (a0e14 is null or a0e14 = 0)"
   strUnion = strUnion & " union select a0e01, a0e02, a0e10, a0e11, a0g02, a0e20, a0e08, st02 as custname, a0e21, a0e48,a0e07 from acc0e0, acc0g0, staff where a0e23='" & strA0E23 & "' and a0e01 = a0g01 and a0e05 = '3' and a0e06 = st01 (+) and a0e45 = '" & strType & "' and a0e19 = '" & Text2 & "' and a0e20 = '" & strExc(1) & "' and a0e21 = " & Val(FCDate(MaskEdBox1.Text)) & " and (a0e14 is null or a0e14 = 0)"
   strUnion = strUnion & " union select a0e01, a0e02, a0e10, a0e11, a0g02, a0e20, a0e08, '' as custname, a0e21, a0e48,a0e07 from acc0e0, acc0g0 where a0e23='" & strA0E23 & "' and a0e01 = a0g01 and a0e45 = '" & strType & "' and a0e05 = '4' and a0e19 = '" & Text2 & "' and a0e20 = '" & strExc(1) & "' and a0e21 = " & Val(FCDate(MaskEdBox1.Text)) & " and (a0e14 is null or a0e14 = 0) order by a0e21 asc, a0e48 desc"
   'end 2023/05/23
'   strUnion = "select a0e01, a0e02, a0e10, a0e11, a0g02, a0e20, a0e08, a0e06 from acc0e0, acc0g0 where a0e01 = a0g01 and a0e19 = '" & Text2 & "' and a0e20 = '" & Combo1 & "' and a0e18 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0e18 <= " & lngDate & ""
   adoadodc1.Open strUnion, adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.Requery
   SumShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  儲存資料表(票據資料)
'
'*************************************************
'Modify by Morgan 2004/10/29
'Private Sub Acc0e0Save()
Private Function Acc0e0Save() As Boolean
Dim strAutoNo As String
Dim strMsg As String 'Add by Amy 2014/11/12
Dim strCombo1 As String 'Add by Amy 2023/05/23

On Error GoTo Checking

   If Text4 = "" Then 'Add by Morgan 2004/10/29
      adocheck.CursorLocation = adUseClient
      'Modify by Morgan 2004/10/29 加應收過濾條件 and a0e04='R'
      'adocheck.Open "select a0e01, a0e02 from acc0e0 where a0e02 = '" & Text5 & "'", adoTaie, adOpenStatic, adLockReadOnly
      'Modified by Morgan 2020/8/7+a0e07
      adocheck.Open "select a0e01,a0e07 from acc0e0 where a0e02 = '" & Text5 & "' and a0e04='R'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
         'Add by Morgan 2004/10/29
         If adocheck.RecordCount > 1 Then
            MsgBox "收票銀行無法確定，請自行輸入！【 " & adocheck.GetString(, , , ";") & "】" 'Add by Morgan 2004/10/29
            adocheck.Close
            Text4.SetFocus
            Exit Function
         End If
         '2004/10/29 end
         If IsNull(adocheck.Fields(0).Value) = False Then
            Text4 = adocheck.Fields(0).Value
            Text1 = adocheck.Fields("a0e07").Value 'Added by Morgan 2020/8/7
         End If
      End If
      adocheck.Close
   End If
   If Text5 = MsgText(601) Then
      MsgBox MsgText(10) & Label1, , MsgText(5)
      strControlButton = MsgText(602)
      Exit Function
   Else
      If Text4 = MsgText(601) Then
         MsgBox MsgText(10) & Label4, , MsgText(5)
         strControlButton = MsgText(602)
         Text4.SetFocus 'Add by Morgan 2004/10/29
         Exit Function
      End If
      If Combo1 = MsgText(601) Then
         MsgBox Label2 & MsgText(52), , MsgText(5)
         strControlButton = MsgText(602)
         Exit Function
      End If
      If Text2 = MsgText(601) Then
         MsgBox Label3 & MsgText(52), , MsgText(5)
         strControlButton = MsgText(602)
         Exit Function
      End If
      If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
         MsgBox Label5 & MsgText(52), , MsgText(5)
         strControlButton = MsgText(602)
         MaskEdBox1.SetFocus
         Exit Function
      Else
         If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
            MsgBox Label5 & MsgText(63), , MsgText(5)
            strControlButton = MsgText(602)
            MaskEdBox1.SetFocus
            Exit Function
         End If
         'Add by Amy 2014/11/12 +系統日檢查
         If ChkWorkData(strA0E23, DBDATE(MaskEdBox1), strMsg) = False Then
            MsgBox Label5 & strMsg, , MsgText(5)
            strControlButton = MsgText(602)
            MaskEdBox1.SetFocus
            Exit Function
        End If
        'end 2014/11/12
      End If
      adocheck.CursorLocation = adUseClient
      'Modify by Amy 2023/05/23 原:Combo1
      strCombo1 = Left(Combo1.Text, InStr(Combo1, " ") - 1)
      adocheck.Open "select a0h01, a0h02 from acc0h0 where a0h01 = '" & Text2 & "' and a0h02 = '" & strCombo1 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount = 0 Then
         MessageShow Label2
         strControlButton = MsgText(602)
         adocheck.Close
         Exit Function
      End If
      adocheck.Close
      adocheck.CursorLocation = adUseClient
      'Modify by Amy 2020/07/22 +a0e07 因改為key
      adocheck.Open "select a0e01, a0e02 from acc0e0 where a0e01 = '" & Text4 & "' and a0e02 = '" & Text5 & "' And a0e07='" & Text1 & "' ", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount = 0 Then
         MessageShow Label1
         strControlButton = MsgText(602)
         adocheck.Close
         Exit Function
      End If
      adocheck.Close
   End If
   adoTaie.BeginTrans
   adoacc0e0.CursorLocation = adUseClient
   'Modify by Morgan 2004/10/29 加應收過濾條件 and a0e04='R'
   'Modify by Amy 2020/07/22 +a0e07 因改為key
   adoacc0e0.Open "select * from acc0e0 where a0e01 = '" & Text4 & "' and a0e02 = '" & Text5 & "' And a0e07='" & Text1 & "' and a0e14 = 0 and a0e15 = 0 and a0e17 = 0 and a0e21 = 0 and a0e25 = 0  and a0e04='R' order by a0e10 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   If adoacc0e0.RecordCount = 0 Then
      MsgBox MsgText(33) & " " & MsgText(39), , MsgText(5)
      adoacc0e0.Close
      adoTaie.RollbackTrans
      Exit Function
   End If
   '2014/1/24 modify by sonia 加判斷公司別a0e23
   If adoacc0e0.Fields("a0e23").Value <> strA0E23 Then
      MsgBox "票據公司別與託收銀行帳號不合！", , MsgText(5)
      strControlButton = MsgText(602)
      adoacc0e0.Close
      adoTaie.RollbackTrans
      Text5.SetFocus
      Exit Function
   End If
   '2014/1/24 end
   If Text2 <> MsgText(601) Then
      adoacc0e0.Fields("a0e19").Value = Text2
   Else
      adoacc0e0.Fields("a0e19").Value = Null
   End If
   'Modify by Amy 2023/05/23 原:combo1
   If strCombo1 <> MsgText(601) Then
      adoacc0e0.Fields("a0e20").Value = strCombo1
   Else
      adoacc0e0.Fields("a0e20").Value = Null
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select * from acc0g0 where a0g01 = '" & Text4 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
'         If adocheck.Fields("a0g09").Value = MsgText(602) Then
            adoacc0e0.Fields("a0e21").Value = Val(FCDate(MaskEdBox1.Text))
'         Else
'            dayDate = Format(Val(FCDate(MaskEdBox1.Text)) + 19110000, "####/##/##")
'            adoacc0e0.Fields("a0e21").Value = ACDate(Format(dayDate + 3, "YYYYMMDD"))
'         End If
      Else
         adoacc0e0.Fields("a0e21").Value = 0
      End If
      adocheck.Close
   Else
      adoacc0e0.Fields("a0e21").Value = 0
   End If
   adoacc0e0.Fields("a0e45").Value = strType
   adoacc0e0.Fields("a0e29").Value = Val(strSrvDate(2))
   adoacc0e0.Fields("a0e30").Value = ServerTime
   adoacc0e0.Fields("a0e31").Value = strUserNum
   adoacc0e0.Fields("a0e48").Value = ServerTime
   adoquery.CursorLocation = adUseClient
   'Modify by Amy 2020/07/22 +a0e07 因改為key
   adoquery.Open "select a0e10||'/'||a0e02||'/'||a0e07||'/'||substr(a0g02, 1, 12) from acc0e0, acc0g0 where a0e01 = a0g01 (+) and a0e02 = '" & Text5 & "' and a0e01 = '" & Text4 & "' And a0e07='" & Text1 & "' ", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields(0).Value) Then
         strRemark = ""
      Else
         strRemark = adoquery.Fields(0).Value
      End If
   End If
   adoquery.Close
   '2014/1/24 modify by sonia a1p01='1' 改為a1p01='" & strA0E23 & "
   'Modify by Amy 2020/07/22 +a0e07 因改為key
   strAutoNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '" & strA0E23 & "' and a1p02 = 'L' and a1p04 = '" & adoacc0e0.Fields("a0e02").Value & adoacc0e0.Fields("a0e01").Value & adoacc0e0.Fields("a0e07").Value & "3" & "'", 3)
   adoquery.CursorLocation = adUseClient
   'Modify by Amy 2023/05/23 原:Combo1
   adoquery.Open "select a0h08 from acc0h0 where a0h01 = '" & Text2 & "' and a0h02 = '" & strCombo1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values " & _
                    "('" & strA0E23 & "', 'L', '" & strAutoNo & "', '" & adoacc0e0.Fields("a0e02").Value & adoacc0e0.Fields("a0e01").Value & adoacc0e0.Fields("a0e07").Value & "3" & "', '" & adoquery.Fields(0).Value & "', '" & MsgText(55) & "', " & adoacc0e0.Fields("a0e11").Value & ", 0, '" & Text5 & "', '" & Text2 & "', '" & strCombo1 & "', " & _
                      "" & Val(adoacc0e0.Fields("a0e10").Value) & ", '" & adoacc0e0.Fields("a0e08").Value & "', '" & strRemark & "', null, null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", null, null, null, null, " & _
                      "'" & adoacc0e0.Fields("a0e03").Value & "', null, null, '1', null)"
   Else
      adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values " & _
                    "('" & strA0E23 & "', 'L', '" & strAutoNo & "', '" & adoacc0e0.Fields("a0e02").Value & adoacc0e0.Fields("a0e01").Value & adoacc0e0.Fields("a0e07").Value & "3" & "', '110201', '" & MsgText(55) & "', " & adoacc0e0.Fields("a0e11").Value & ", 0, '" & Text5 & "', '" & Text2 & "', '" & strCombo1 & "', " & _
                      "" & Val(adoacc0e0.Fields("a0e10").Value) & ", '" & adoacc0e0.Fields("a0e08").Value & "', '" & strRemark & "', null, null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", null, null, null, null, " & _
                      "'" & adoacc0e0.Fields("a0e03").Value & "', null, null, '1', null)"
   End If
   adoquery.Close
   strAutoNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01 = '" & strA0E23 & "' and a1p02 = 'L' and a1p04 = '" & adoacc0e0.Fields("a0e02").Value & adoacc0e0.Fields("a0e01").Value & adoacc0e0.Fields("a0e07").Value & "3" & "'", 3)
   adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values " & _
                "('" & strA0E23 & "', 'L', '" & strAutoNo & "', '" & adoacc0e0.Fields("a0e02").Value & adoacc0e0.Fields("a0e01").Value & adoacc0e0.Fields("a0e07").Value & "3" & "', '113001', '" & MsgText(55) & "', 0, " & adoacc0e0.Fields("a0e11").Value & ", '" & Text5 & "', '" & Text2 & "', '" & strCombo1 & "', " & _
                   "" & Val(adoacc0e0.Fields("a0e10").Value) & ", '" & adoacc0e0.Fields("a0e08").Value & "', '" & strRemark & "', null, null, null, " & Val(FCDate(MaskEdBox1.Text)) & ", null, null, null, null, " & _
                   "'" & adoacc0e0.Fields("a0e03").Value & "', null, null, '1', null)"
   'end 2023/05/23
   'end 2020/07/22
   adoacc0e0.UpdateBatch
   adoTaie.CommitTrans
   Acc0e0Save = True 'Add by Morgan 2004/10/29
   AdodcRefresh
   adoacc0e0.Close
Checking:
   If Err.Number = 0 Then
      Exit Function
   End If
   MsgBox Err.Description, , MsgText(5)
End Function

'*************************************************
'  刪除 Adodc 之資料
'
'*************************************************
Private Sub AdodcDelete()
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   adocheck.CursorLocation = adUseClient
   'Modify by Amy 2020/07/22 +a0e07 因改為key
   adocheck.Open "select ax210 from acc1p0, acc021 where a1p22 = ax202 and a1p01 = '" & strA0E23 & "' and a1p02 = 'L' and a1p04 = '" & Adodc1.Recordset.Fields("a0e02").Value & Adodc1.Recordset.Fields("a0e01").Value & Adodc1.Recordset.Fields("a0e07").Value & "3" & "' and ax210 is not null", adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount <> 0 Then
      MsgBox MsgText(180), , MsgText(5)
      adocheck.Close
      Exit Sub
   End If
   adocheck.Close
   adoTaie.BeginTrans
   adoTaie.Execute "delete from acc1p0 where a1p01 = '" & strA0E23 & "' and a1p02 = 'L' and a1p04 = '" & Adodc1.Recordset.Fields("a0e02").Value & Adodc1.Recordset.Fields("a0e01").Value & Adodc1.Recordset.Fields("a0e07").Value & "3" & "'"
   adoTaie.Execute "update acc0e0 set a0e21 = 0, a0e19 = null, a0e20 = null where a0e01 = '" & Adodc1.Recordset.Fields("a0e01").Value & "' and a0e02 = '" & Adodc1.Recordset.Fields("a0e02").Value & "' And a0e07='" & Adodc1.Recordset.Fields("a0e07").Value & "' "
   'end 2020/07/22
   adoTaie.CommitTrans
   AdodcRefresh
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
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         AdodcRefresh
      Case vbKeyInsert
         'Modify by Morgan 2004/10/29
         'Acc0e0Save
         If Acc0e0Save = True Then
            Text5 = MsgText(601)
            Text4 = MsgText(601)
            Text6 = MsgText(601)
            Text1 = MsgText(601) 'Added by Morgan 2020/8/7
            Text5.SetFocus
         End If
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133)
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox Label5 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   Else
      If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
         MsgBox Label5 & MsgText(63), , MsgText(5)
         Cancel = True
         MaskEdBox1.SetFocus
         Exit Sub
      End If
      AdodcRefresh
   End If
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
   Dim strCombo1 As String 'Add by Amy 2023/05/23
   
   If Combo1 <> MsgText(601) Then
      'Modify by Amy 2023/05/23 原:Combo1
      strCombo1 = Left(Combo1.Text, InStr(Combo1, " ") - 1)
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select a0h01, a0h02 from acc0h0 where a0h02 = '" & strCombo1 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
         If IsNull(adocheck.Fields(0).Value) = False Then
            Text2 = adocheck.Fields(0).Value
            adocheck.Close
            '2014/1/24 add by sonia
            'modify by sonia 2015/5/12 加智權 華銀長安0236819
            If strCombo1 = "1607750" Or strCombo1 = "0236819" Then
               strA0E23 = "J"
            'add by sonia 2020/4/7 加法律所
            ElseIf strCombo1 = "1756890" Then
               strA0E23 = "L"
            'end 2020/4/7
            Else
               strA0E23 = "1"
            End If
            '2014/1/24 end
            Exit Sub
         End If
         'end 2023/05/23
      End If
      MessageShow Label2
      Cancel = True
      adocheck.Close
   End If
End Sub

'Add by Amy 2020/07/22
Private Sub Text1_GotFocus()
    TextInverse Text1
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    If Text1 <> "" Then 'Added by Morgan 2020/8/7
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select a0e01, a0e02, a0e11 from acc0e0 where a0e02 = '" & Text5 & "' and a0e01 = '" & Text4 & "' And a0e07='" & Text1 & "' and a0e04='R'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
         If IsNull(adocheck.Fields("a0e11").Value) Then
               Text6 = ""
            Else
               Text6 = Format(adocheck.Fields("a0e11").Value, DDollar)
            End If
         adocheck.Close
         Exit Sub
      End If
      MessageShow Label1
      Cancel = True
      adocheck.Close
   End If 'Added by Morgan 2020/8/7
End Sub
'end 2020/07/22

Private Sub Text2_Change()
   If Text2 = MsgText(601) Then
      Exit Sub
   End If
   Text3 = A0g02Query(Text2)
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 <> MsgText(601) Then
      If ExistCheck("acc0g0", "a0g01", Text2, Label3) = False Then
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
   If Text4 <> MsgText(601) Then
      If ExistCheck("acc0g0", "a0g01", Text4, Label4) = False Then
         Cancel = True
         Exit Sub
      End If
      'ADD BY SONIA 2013/11/15 票號0271438有二筆(011030754及050040118)
      adocheck.CursorLocation = adUseClient
      'Modify by Amy 2020/07/22 +a0e07及判斷只有一筆才預帶
      adocheck.Open "select a0e01, a0e02, a0e11,a0e07 from acc0e0 where a0e02 = '" & Text5 & "' and a0e01 = '" & Text4 & "' and a0e04='R'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
         If adocheck.RecordCount = 1 Then
            If IsNull(adocheck.Fields("a0e11").Value) Then
               Text6 = ""
            Else
               Text6 = Format(adocheck.Fields("a0e11").Value, DDollar)
            End If
            Text1 = adocheck.Fields("a0e07")
         End If
         adocheck.Close
         Exit Sub
      End If
      'end 2020/07/22
'      MessageShow Label1
'      Cancel = True
      adocheck.Close
      '2013/11/15 END
   End If
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   If Text5 <> MsgText(601) Then
      adocheck.CursorLocation = adUseClient
      'Modify by Morgan 2004/10/29  加應收過濾條件 and a0e04='R'
      adocheck.Open "select a0e01, a0e02, a0e11 from acc0e0 where a0e02 = '" & Text5 & "' and a0e04='R'", adoTaie, adOpenStatic, adLockReadOnly
      'Modify by Amy 2020/07/22 只有一筆預帶
      If adocheck.RecordCount <> 0 And adocheck.RecordCount = 1 Then
         If IsNull(adocheck.Fields("a0e01").Value) = False Then
            Text4 = adocheck.Fields("a0e01").Value
         Else
            Text4 = ""
         End If
'         If IsNull(adocheck.Fields("a0e11").Value) Then
'            Text6 = ""
'         Else
'            Text6 = Format(adocheck.Fields("a0e11").Value, DDollar)
'         End If
         adocheck.Close
         Exit Sub
      End If
'      MessageShow Label1
'      Cancel = True
      adocheck.Close
'      Exit Sub
   End If
End Sub

'*************************************************
'  筆數及金額合計
'
'*************************************************
Private Sub SumShow()
Dim adoaccsum As New ADODB.Recordset

   adoaccsum.CursorLocation = adUseClient
   '2014/1/24 modify by sonia 加判斷公司別a0e23
   'adoaccsum.Open "select count(a0e02), sum(a0e11) from acc0e0 where a0e19 = '" & Text2 & "' and a0e20 = '" & Combo1 & "' and a0e21 = " & Val(FCDate(MaskEdBox1.Text)) & " and a0e45 = '" & strType & "' and a0e15 = 0 and a0e14 = 0", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2023/05/23 原:Combo1
   strExc(1) = Left(Combo1.Text, InStr(Combo1, " ") - 1)
   adoaccsum.Open "select count(a0e02), sum(a0e11) from acc0e0 where a0e23='" & strA0E23 & "' and a0e19 = '" & Text2 & "' and a0e20 = '" & strExc(1) & "' and a0e21 = " & Val(FCDate(MaskEdBox1.Text)) & " and a0e45 = '" & strType & "' and a0e15 = 0 and a0e14 = 0", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text8 = MsgText(601)
      Else
         Text8 = Format(adoaccsum.Fields(0).Value, DDollar)
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text7 = MsgText(601)
      Else
         Text7 = Format(adoaccsum.Fields(1).Value, DDollar)
      End If
   Else
      Text8 = MsgText(601)
      Text7 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

