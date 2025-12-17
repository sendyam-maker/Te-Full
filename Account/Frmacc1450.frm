VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc1450 
   AutoRedraw      =   -1  'True
   Caption         =   "客戶帳款明細表"
   ClientHeight    =   4584
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4584
   ScaleWidth      =   8760
   Begin VB.TextBox TxtDeptE 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   2640
      MaxLength       =   9
      TabIndex        =   2
      Top             =   615
      Width           =   1200
   End
   Begin VB.TextBox TxtDeptS 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   1080
      MaxLength       =   9
      TabIndex        =   1
      Top             =   615
      Width           =   1200
   End
   Begin VB.CommandButton cmdLikeSearch 
      Caption         =   "搜尋"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7830
      TabIndex        =   5
      Top             =   990
      Width           =   675
   End
   Begin VB.ComboBox cboTitle 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "Frmacc1450.frx":0000
      Left            =   1080
      List            =   "Frmacc1450.frx":0002
      TabIndex        =   4
      Top             =   960
      Width           =   6720
   End
   Begin VB.TextBox Text16 
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
      Left            =   2070
      MaxLength       =   1
      TabIndex        =   22
      Text            =   "Y"
      Top             =   3090
      Width           =   690
   End
   Begin VB.ComboBox cboPrinters 
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
      Left            =   1080
      Style           =   2  '單純下拉式
      TabIndex        =   0
      Top             =   120
      Width           =   3525
   End
   Begin VB.TextBox txtSys 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   5520
      MaxLength       =   100
      TabIndex        =   3
      Top             =   615
      Width           =   2850
   End
   Begin VB.TextBox Text15 
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
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   17
      Text            =   "L"
      Top             =   1950
      Width           =   390
   End
   Begin VB.TextBox Text14 
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
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   16
      Text            =   "1"
      Top             =   1950
      Width           =   390
   End
   Begin VB.TextBox Text13 
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
      Left            =   6090
      MaxLength       =   1
      TabIndex        =   21
      Top             =   2760
      Width           =   690
   End
   Begin VB.TextBox Text12 
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
      Left            =   2070
      MaxLength       =   1
      TabIndex        =   20
      Text            =   "Y"
      Top             =   2760
      Width           =   690
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
      Height          =   300
      Left            =   1860
      MaxLength       =   6
      TabIndex        =   11
      Top             =   1620
      Width           =   2712
   End
   Begin VB.TextBox Text10 
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
      Left            =   7470
      MaxLength       =   2
      TabIndex        =   15
      Top             =   1620
      Width           =   495
   End
   Begin VB.TextBox Text9 
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
      Left            =   7110
      MaxLength       =   1
      TabIndex        =   14
      Top             =   1620
      Width           =   375
   End
   Begin VB.TextBox Text8 
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
      Left            =   6150
      MaxLength       =   6
      TabIndex        =   13
      Top             =   1620
      Width           =   975
   End
   Begin VB.TextBox Text7 
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
      Left            =   5550
      MaxLength       =   3
      TabIndex        =   12
      Top             =   1620
      Width           =   615
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
      Height          =   315
      Left            =   5550
      MaxLength       =   9
      TabIndex        =   18
      Top             =   1950
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1920
      Style           =   1  '圖片外觀
      TabIndex        =   23
      Top             =   3540
      Width           =   4692
   End
   Begin VB.TextBox Text5 
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
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   19
      Top             =   2310
      Width           =   612
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
      Height          =   300
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   10
      Top             =   1620
      Width           =   780
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
      Left            =   1080
      MaxLength       =   9
      TabIndex        =   6
      Top             =   1290
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
      Left            =   3000
      MaxLength       =   9
      TabIndex        =   7
      Top             =   1290
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   5550
      TabIndex        =   8
      Top             =   1290
      Width           =   1245
      _ExtentX        =   2180
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   7140
      TabIndex        =   9
      Top             =   1290
      Width           =   1245
      _ExtentX        =   2201
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
   Begin VB.Label Label21 
      BackStyle       =   0  '透明
      Caption         =   "部　門"
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
      Left            =   150
      TabIndex        =   44
      Top             =   648
      Width           =   1000
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
      Height          =   252
      Left            =   2400
      TabIndex        =   43
      Top             =   615
      Width           =   252
   End
   Begin VB.Label Label20 
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
      Left            =   150
      TabIndex        =   42
      Top             =   1020
      Width           =   975
   End
   Begin VB.Label Label19 
      BackStyle       =   0  '透明
      Caption         =   "函證用不包含未列印收據"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   3660
      TabIndex        =   41
      Top             =   3000
      Width           =   1365
   End
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "是否含未列印收據            ( Y:含 )"
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
      Left            =   150
      TabIndex        =   40
      Top             =   3150
      Width           =   4605
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "印表機"
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
      Left            =   150
      TabIndex        =   39
      Top             =   150
      Width           =   975
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "系統類別"
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
      Left            =   4596
      TabIndex        =   38
      Top             =   648
      Width           =   972
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "PS : 列印類別選 4 時不適用於本所案號列印           列印類別選 5 時不含未列印收據資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   180
      TabIndex        =   37
      Top             =   4020
      Width           =   4695
   End
   Begin VB.Label Label14 
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
      Left            =   1650
      TabIndex        =   36
      Top             =   1980
      Width           =   255
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
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
      Left            =   150
      TabIndex        =   35
      Top             =   2010
      Width           =   975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "是否列印客戶案件案號            ( Y:是 )"
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
      TabIndex        =   34
      Top             =   2760
      Width           =   4605
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "是否依公司別跳頁            ( Y:是 )"
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
      Left            =   150
      TabIndex        =   33
      Top             =   2760
      Width           =   4605
   End
   Begin VB.Label Label10 
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
      Left            =   4620
      TabIndex        =   32
      Top             =   1695
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "合併列印客戶代號"
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
      Left            =   3630
      TabIndex        =   31
      Top             =   2010
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   6900
      Top             =   4020
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "(1.未收 2.收回 3.往來日期      4.交易情形 5.函證用未收)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1800
      TabIndex        =   30
      Top             =   2310
      Width           =   2985
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "列印類別"
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
      Left            =   150
      TabIndex        =   29
      Top             =   2370
      Width           =   975
   End
   Begin VB.Label Label5 
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
      Left            =   6930
      TabIndex        =   28
      Top             =   1290
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "作業日期"
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
      Left            =   4620
      TabIndex        =   27
      Top             =   1350
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Height          =   210
      Left            =   150
      TabIndex        =   26
      Top             =   1665
      Width           =   900
   End
   Begin VB.Label Label1 
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
      Left            =   150
      TabIndex        =   25
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label7 
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
      Left            =   2760
      TabIndex        =   24
      Top             =   1290
      Width           =   255
   End
End
Attribute VB_Name = "Frmacc1450"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
Option Explicit

Public adoacc0k0 As New ADODB.Recordset
Public adoacc0m0 As New ADODB.Recordset
Public adocaseprogress As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoaccrpt105 As New ADODB.Recordset
Dim strCustomerNo As String
Dim lngCounter As Long
Dim lngAmount As Long
Dim strAmount As String
Dim intLength As Integer
Dim intCounter As Integer
Dim intPage As Integer
Dim strSameName As String
'Add By Cheng 2003/04/24
Dim m_strA0k11 As String  '公司別
Dim m_strCompany As String
Dim strProperty As String
Dim strNation As String
'Add By Cheng 2003/06/10
Dim m_lngMaxNo As Integer '最大序號
Dim ii As Double '序號
'Add by Morgan 2004/11/23
Dim lngCurrentX As Long, lngCurrentY As Long
Dim m_iDefaultPrinter As Integer '預設印表機
Dim m_sColumn(0 To 15) As String '欄位名稱
Dim m_iColWidth(0 To 15) As Integer '欄位寬度
Dim m_iVPad As Integer '行距
Dim m_iHPad As Integer '列距
Dim m_iMaxLine As Integer '可印行數
Dim lngSubTot(1 To 5) As Long, lngTot(1 To 5) As Long, jj As Integer '小計,合計

Private Sub cboPrinters_Click()
   Set Printer = Printers(cboPrinters.ListIndex)
End Sub

'Add By Sindy 2016/6/8
Private Sub cmdLikeSearch_Click()
   If cboTitle.Text = "" Then
      MsgBox "請輸入收據抬頭！", vbCritical
   Else
      PUB_AddItem2CboTitle cboTitle, Text1, Text2, "", True
   End If
End Sub

'Add By Sindy 2016/6/7
Private Sub cboTitle_Click()
   If cboTitle.ListIndex > 0 Then
      If Text1.Text = "" Then
         Text1.Text = Right(cboTitle.Text, 9)
      ElseIf Text2.Text = "" Then
         Text2.Text = Right(cboTitle.Text, 9)
      End If
      strExc(1) = cboTitle.List(cboTitle.ListIndex)
      cboTitle.List(0) = RTrim(Left(strExc(1), Len(strExc(1)) - 9))
   End If
   cboTitle.ListIndex = 0
End Sub
Private Sub cboTitle_GotFocus()
   OpenIme
End Sub
Private Sub cboTitle_KeyPress(KeyAscii As Integer)
   If Text1 <> "" Or Text2 <> "" Or cboTitle.ListCount > 0 Then
      Text1 = "": Text2 = ""
      Text4 = "": Text11 = ""
      cboTitle.Clear
   End If
End Sub
Private Sub cboTitle_Validate(Cancel As Boolean)
   If CheckLen(Label1, cboTitle, 100) = MsgText(603) Then
      Cancel = True
      Exit Sub
   End If
   '切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub
'2016/6/7 END

Private Sub Command1_Click()
Dim bolShowMsg As Boolean 'Add By Sindy 2016/6/8
   
   'Add by Morgan 2004/11/23 印表機提示
   If MsgBox("請選擇雷射印表機列印！", vbOKCancel + vbDefaultButton2, "選擇印表機") = vbCancel Then
      Exit Sub
   End If
   SetColumnName '設定表頭欄位
   '2004/11/23 end
   
   'Add by Morgan 2005/6/3
   '選函證用未收時作業日期迄日不可空白(票據到期日的參考日期)
   If Text5 = "5" Then
      If MaskEdBox2.Text = MsgText(29) Then
         MsgBox "選函證用未收時作業日期迄日不可空白！", vbExclamation
         MaskEdBox2.SetFocus
         Exit Sub
      End If
   End If
   'Modify By Sindy 2016/6/8
   If Text5 = MsgText(601) Then
      MsgBox "列印類別不可空白！", vbExclamation
      Text5.SetFocus
      Exit Sub
   End If
   '2016/6/8 END
   
   If FormCheck() = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Accrpt105Delete
   ProduceData
   FormPrint
   Screen.MousePointer = vbDefault
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = "請選雷射印表機列印..."
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(114) & " / " & MsgText(102)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 8880
    'Modify By Cheng 2003/03/12
    '預設表單高度
'   Me.Height = 4100
'   Me.Height = 4455
   'Modify by Amy 2023/08/14 原:4890
   Me.Height = 5030 '6120 '5775 'Modify by Morgan 2004/11/23 加系統類別,印表機
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   Text1 = "X"
   Text2 = "X"
   Text5 = "1"
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(114) & " / " & MsgText(102)
   
   PUB_SetPrinter Me.Name, cboPrinters, , , m_iDefaultPrinter  'Modified by Morgan 2017/11/9 設定印表機改呼叫公用函數,原程式移除
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   '若印表機變動, 則更新列印設定
   If Me.cboPrinters.Text <> Me.cboPrinters.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, cboPrinters.Name, "0", "0", Me.cboPrinters.Text
   End If
   Set Printer = Printers(m_iDefaultPrinter)
   Set Frmacc1450 = Nothing
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
'   MaskEdBox2.Mask = ""
'   MaskEdBox2.Text = MaskEdBox1.Text
'   MaskEdBox2.Mask = DFormat
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Len(Text1) = 6 Then
      Text1 = AfterZero(Text1)
   End If
   'Modify by Morgan 2004/11/26 改帶出關係企業(尾3碼改999)
   'Text2 = Text1
   'Modify By Sindy 2014/8/11 999=>ZZZ
   'If Len(Text1.Text) >= 6 Then Text2.Text = Left(Text1.Text, 6) & "999"
   If Len(Text1.Text) >= 6 Then Text2.Text = Left(Text1.Text, 6) & "ZZZ"
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub

Private Sub Text12_GotFocus()
    'Add By Cheng 2003/03/12
    TextInverse Text12
    CloseIme
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
'    'Add By Cheng 2003/03/12
'    '只可輸入1或2或不輸
    KeyAscii = UpperCase(KeyAscii)
    Select Case KeyAscii
    Case 89, 8
        'Do Nothing
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub Text13_GotFocus()
    TextInverse Me.Text13
    CloseIme
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 8 And KeyAscii <> 89 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text14_GotFocus()
   TextInverse Text14
   CloseIme
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text15_GotFocus()
   TextInverse Text15
   CloseIme
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2012/3/27
Private Sub Text16_GotFocus()
    TextInverse Text16
End Sub

'Add By Sindy 2012/3/27
Private Sub Text16_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    Select Case KeyAscii
    Case 89, 8
        'Do Nothing
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Len(Text2) = 6 Then
      Text2 = AfterZero(Text2)
   End If
End Sub

'Private Sub Text3_GotFocus()
'   TextInverse Text3
'End Sub

Private Sub Text4_Change()
   If Len(Text4) = 5 Then
      Text11 = StaffQuery(Text4)
   Else
      Text11 = MsgText(601)
   End If
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   If Not (KeyAscii = 8 Or KeyAscii >= Asc("1") And KeyAscii <= Asc("5")) Then
      KeyAscii = 0
      Beep
   ElseIf KeyAscii = Asc("4") Then
      MsgBox "由於本系統暫收款資料還無法完全串聯，故類別4暫停使用！", vbExclamation
      KeyAscii = 0
   End If
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
On Error GoTo Checking
   lngCounter = 0
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   If adoaccrpt105.State = adStateOpen Then
      adoaccrpt105.Close
   End If
   adoaccrpt105.CursorLocation = adUseClient
   adoaccrpt105.Open "select * from accrpt105", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Select Case Text5 '列印類別
      Case "1" '未收
         Select1New
      Case "2" '收回
         Select2New
      Case "3" '往來日期
         Select3New
      Case "4" '交易情形
         'Select4 '目前不開放使用
      Case "5" '函證用未收
         Select5New
      Case Else
         Select1New
   End Select
   adoaccrpt105.Close
   StatusClear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  刪除報表資料
'
'*************************************************
Private Sub Accrpt105Delete()
   adoTaie.Execute "delete from accrpt105"
End Sub


'*************************************************
'  畫線
'
'*************************************************
Private Sub PaintLine(strSign As String)
Dim intCounter As Integer

   Select Case Text5
      Case "1"
         For intCounter = 12 To 16
            adoaccrpt105.Fields(intCounter).Value = strSign
         Next intCounter
      Case "2", "3", "4"
         For intCounter = 12 To 15
            adoaccrpt105.Fields(intCounter).Value = strSign
         Next intCounter
      Case Else
         For intCounter = 12 To 16
            adoaccrpt105.Fields(intCounter).Value = strSign
         Next intCounter
   End Select
End Sub

'*************************************************
'  計數器
'
'*************************************************
Private Function Counter() As Long
   lngCounter = lngCounter + 1
   Counter = lngCounter
    'Add By Cheng 2003/06/10
    m_lngMaxNo = lngCounter
End Function

Private Function GetSysSQL(ByVal stSys As String) As String
   Dim arrSys, i As Integer
   arrSys = Split(stSys, ",")
   For i = LBound(arrSys) To UBound(arrSys)
      arrSys(i) = "'" & arrSys(i) & "'"
   Next
   GetSysSQL = Join(arrSys, ",")
   
End Function

'2013/5/13 cancel by sonia 沒有用到
''*************************************************
''  列印類別-選擇未收之小計計算
''
''*************************************************
'Private Sub SubSelect1()
'Dim strSql As String
''Add By Cheng 2003/04/24
'Dim StrSQLa As String
'
'   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'      strSql = " and a0k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'   End If
'   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'      strSql = strSql & " and a0k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'   End If
'   If Text7 <> MsgText(601) Then
'      strSql = strSql & " and cp01 = '" & Text7 & "'"
'   End If
'   If Text8 <> MsgText(601) Then
'      strSql = strSql & " and cp02 = '" & Text8 & "'"
'   End If
'   If Text9 <> MsgText(601) Then
'      strSql = strSql & " and cp03 = '" & Text9 & "'"
'   End If
'   If Text10 <> MsgText(601) Then
'      strSql = strSql & " and cp04 = '" & Text10 & "'"
'   End If
'   If Text4 <> MsgText(601) Then
'      strSql = strSql & " and a0k20 = '" & Text4 & "'"
'   End If
'   adoacc0k0.MovePrevious
'   adoaccsum.CursorLocation = adUseClient
'   'adoaccsum.Open "select sum(decode(a0k30, 'Y', nvl(cp16, 0), nvl(cp16, 0) - nvl(cp17, 0))), sum(decode(a0k30, 'Y', 0, nvl(cp17, 0))), sum(nvl(cp16, 0)), sum(nvl(cp75, 0)), sum(nvl(cp16, 0) - nvl(cp75, 0)) from caseprogress, acc0j0, acc0k0 where cp09 = a0j01 (+) and cp60 = a0k01 (+) and (cp79 <> 0 or cp79 is null) and (a0k09 is null or a0k09 = 0) and (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0)) and a0k03 = '" & adoacc0k0.Fields("a0k03").Value & "' and instr(a0k04, '" & Trim(adoacc0k0.Fields("a0k04").Value) & "') > 0" & strSQL, adoTaie, adOpenStatic, adLockReadOnly
'    'Modify By Cheng 2003/04/24
'    '依公司別跳頁
'    If Me.Text12.Text = "Y" Then
'        'Modify By Cheng 2003/06/10
''        strSQLA = "select sum(to_number(nvl(r10513, 0))), sum(to_number(nvl(r10514, 0))), sum(to_number(nvl(r10515, 0))), sum(to_number(nvl(r10516, 0))), sum(to_number(nvl(r10517, 0))) from accrpt105 where r10501 = '" & strUserNum & "' and r10507 is not null and r10522='" & m_strA0k11 & "' and r10504 = '" & Trim(adoacc0k0.Fields("a0k04").Value) & "'"
'        StrSQLa = "select sum(to_number(nvl(r10513, 0))), sum(to_number(nvl(r10514, 0))), sum(to_number(nvl(r10515, 0))), sum(to_number(nvl(r10516, 0))), sum(to_number(nvl(r10517, 0))) from accrpt105 where r10501 = '" & strUserNum & "' and r10507 is not null and r10522='" & m_strA0k11 & "' and r10503 = '" & Trim(adoacc0k0.Fields("r10503").Value) & "' and r10504 = '" & Trim(adoacc0k0.Fields("r10504").Value) & "'"
'        adoaccsum.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'    '不依公司別跳頁
'    Else
'        'Modify By Cheng 2003/06/10
''        strSQLA = "select sum(to_number(nvl(r10513, 0))), sum(to_number(nvl(r10514, 0))), sum(to_number(nvl(r10515, 0))), sum(to_number(nvl(r10516, 0))), sum(to_number(nvl(r10517, 0))) from accrpt105 where r10501 = '" & strUserNum & "' and r10507 is not null and r10504 = '" & Trim(adoacc0k0.Fields("a0k04").Value) & "'"
'        StrSQLa = "select sum(to_number(nvl(r10513, 0))), sum(to_number(nvl(r10514, 0))), sum(to_number(nvl(r10515, 0))), sum(to_number(nvl(r10516, 0))), sum(to_number(nvl(r10517, 0))) from accrpt105 where r10501 = '" & strUserNum & "' and r10507 is not null and r10503 = '" & Trim(adoacc0k0.Fields("r10503").Value) & "' and r10504 = '" & Trim(adoacc0k0.Fields("r10504").Value) & "'"
'        adoaccsum.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'    End If
'   If adoaccsum.RecordCount <> 0 And (adoaccsum.Fields(0).Value <> 0 Or adoaccsum.Fields(1).Value <> 0 Or adoaccsum.Fields(2).Value <> 0 Or adoaccsum.Fields(3).Value <> 0 Or adoaccsum.Fields(4).Value <> 0) Then
'      ii = ii + 1
'      adoaccrpt105.AddNew
'      adoaccrpt105.Fields("r10501").Value = strUserNum
'      adoaccrpt105.Fields("r10502").Value = ii
'      If Text6 <> "" Then
'         adoaccrpt105.Fields("r10503").Value = Text6
'         adoaccrpt105.Fields("r10504").Value = Text3
'      Else
'         adoaccrpt105.Fields("r10503").Value = adoacc0k0.Fields("r10503").Value
'         adoaccrpt105.Fields("r10504").Value = Trim(adoacc0k0.Fields("r10504").Value)
'      End If
'       'Add By Cheng 2003/04/24
'       '公司別
'       adoaccrpt105.Fields("r10522").Value = adoacc0k0.Fields("r10522").Value
'      PaintLine ReportSum(4)
'      adoaccrpt105.UpdateBatch
'      ii = ii + 1
'      adoaccrpt105.AddNew
'      adoaccrpt105.Fields("r10501").Value = strUserNum
'      adoaccrpt105.Fields("r10502").Value = ii
'      If Text6 <> "" Then
'         adoaccrpt105.Fields("r10503").Value = Text6
'         adoaccrpt105.Fields("r10504").Value = Text3
'      Else
'         adoaccrpt105.Fields("r10503").Value = adoacc0k0.Fields("r10503").Value
'         adoaccrpt105.Fields("r10504").Value = Trim(adoacc0k0.Fields("r10504").Value)
'      End If
'      adoaccrpt105.Fields("r10510").Value = ReportSum(24)
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         adoaccrpt105.Fields("r10513").Value = 0
'      Else
''         If adoacc0k0.Fields("a0k30").Value = MsgText(602) Then
''            adoaccrpt105.Fields("r10513").Value = adoaccsum.Fields(2).Value
''         Else
'            adoaccrpt105.Fields("r10513").Value = adoaccsum.Fields(0).Value
''         End If
'      End If
'      If IsNull(adoaccsum.Fields(1).Value) Then
'         adoaccrpt105.Fields("r10514").Value = 0
'      Else
''         If adoacc0k0.Fields("a0k30").Value = MsgText(602) Then
''            adoaccrpt105.Fields("r10514").Value = 0
''         Else
'            adoaccrpt105.Fields("r10514").Value = adoaccsum.Fields(1).Value
''         End If
'      End If
'      If IsNull(adoaccsum.Fields(2).Value) Then
'         adoaccrpt105.Fields("r10515").Value = 0
'      Else
'         adoaccrpt105.Fields("r10515").Value = adoaccsum.Fields(2).Value
'      End If
'      If IsNull(adoaccsum.Fields(3).Value) Then
'         adoaccrpt105.Fields("r10516").Value = 0
'      Else
'         adoaccrpt105.Fields("r10516").Value = adoaccsum.Fields(3).Value
'      End If
'      If IsNull(adoaccsum.Fields(4).Value) Then
'         adoaccrpt105.Fields("r10517").Value = 0
'      Else
'         adoaccrpt105.Fields("r10517").Value = adoaccsum.Fields(4).Value
'      End If
'    'Add By Cheng 2003/04/24
'    '公司別
'      adoaccrpt105.Fields("r10522").Value = adoacc0k0.Fields("r10522").Value
'      adoaccrpt105.UpdateBatch
'   End If
'   adoaccsum.Close
'   adoacc0k0.MoveNext
'End Sub

'2013/5/13 cancel by sonia 沒有用到
''*************************************************
''  列印類別-選擇收回之小計計算
''
''*************************************************
'Private Sub SubSelect2()
'Dim strSql As String
'
'   If Text4 <> MsgText(601) Then
'      strSql = " and a0k20 = '" & Text4 & "'"
'   End If
'   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'      strSql = strSql & " and a0l02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'   End If
'   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'      strSql = strSql & " and a0l02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'   End If
'   adocaseprogress.CursorLocation = adUseClient
'   adocaseprogress.Open "select cp60 from caseprogress where cp01 = '" & Text7 & "' and cp02 = '" & Text8 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text10 & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adocaseprogress.RecordCount <> 0 Then
'      strSql = strSql & " and a0k01 in ("
'      Do While adocaseprogress.EOF = False
'         If IsNull(adocaseprogress.Fields("cp60").Value) = False Then
'            strSql = strSql & "'" & adocaseprogress.Fields("cp60").Value & "', "
'         End If
'         adocaseprogress.MoveNext
'      Loop
'      strSql = Mid(strSql, 1, Len(strSql) - 2) & ")"
'   End If
'   adocaseprogress.Close
'   adoacc0m0.MovePrevious
'   adoaccsum.CursorLocation = adUseClient
'   'adoaccsum.Open "select sum(a0m04), sum(a0m05), sum(a0m06) from acc0m0, acc0l0, acc0k0 where a0m01 = a0l01 and a0m02 = a0k01 and a0k03 = '" & adoacc0m0.Fields("a0k03").Value & "' and a0k04 = '" & adoacc0m0.Fields("a0k04").Value & "'" & strSQL, adoTaie, adOpenStatic, adLockReadOnly
'    'Modify By Cheng 2003/04/24
'    '依公司別跳頁
'    If Me.Text12.Text = "Y" Then
'        'Modify By Cheng 2003/07/11
''        adoaccsum.Open "select sum(to_number(nvl(r10513, 0))), sum(to_number(nvl(r10515, 0))) from accrpt105 where r10501 = '" & strUserNum & "' and r10507 is not null and r10503 = '" & Trim(adoacc0m0.Fields("a0k03").Value) & "' and r10504 = '" & Trim(adoacc0m0.Fields("a0k04").Value) & "' and r10522='" & m_strA0k11 & "' ", adoTaie, adOpenStatic, adLockReadOnly
'        adoaccsum.Open "select sum(to_number(nvl(r10513, 0))), sum(to_number(nvl(r10515, 0))), sum(to_number(nvl(r10514, 0))), sum(to_number(nvl(r10516, 0))) from accrpt105 where r10501 = '" & strUserNum & "' and r10507 is not null and r10503 = '" & Trim(adoacc0m0.Fields("a0k03").Value) & "' and r10504 = '" & Trim(adoacc0m0.Fields("a0k04").Value) & "' and r10522='" & m_strA0k11 & "' ", adoTaie, adOpenStatic, adLockReadOnly
'    '不依公司別跳頁
'    Else
'        'Modify By Cheng 2003/07/11
''        adoaccsum.Open "select sum(to_number(nvl(r10513, 0))), sum(to_number(nvl(r10515, 0))) from accrpt105 where r10501 = '" & strUserNum & "' and r10507 is not null and r10503 = '" & Trim(adoacc0m0.Fields("a0k03").Value) & "' and r10504 = '" & Trim(adoacc0m0.Fields("a0k04").Value) & "'", adoTaie, adOpenStatic, adLockReadOnly
'        adoaccsum.Open "select sum(to_number(nvl(r10513, 0))), sum(to_number(nvl(r10515, 0))), sum(to_number(nvl(r10514, 0))), sum(to_number(nvl(r10516, 0))) from accrpt105 where r10501 = '" & strUserNum & "' and r10507 is not null and r10503 = '" & Trim(adoacc0m0.Fields("a0k03").Value) & "' and r10504 = '" & Trim(adoacc0m0.Fields("a0k04").Value) & "'", adoTaie, adOpenStatic, adLockReadOnly
'    End If
'   If adoaccsum.RecordCount <> 0 And (adoaccsum.Fields(0).Value <> 0 Or adoaccsum.Fields(1).Value <> 0 Or adoaccsum.Fields(2).Value <> 0 Or adoaccsum.Fields(3).Value <> 0) Then
'      adoaccrpt105.AddNew
'      adoaccrpt105.Fields("r10501").Value = strUserNum
'      adoaccrpt105.Fields("r10502").Value = Counter
'      If Text6 <> "" Then
'         adoaccrpt105.Fields("r10503").Value = Text6
'         adoaccrpt105.Fields("r10504").Value = Text3
'      Else
'         adoaccrpt105.Fields("r10503").Value = adoacc0m0.Fields("a0k03").Value
'         adoaccrpt105.Fields("r10504").Value = adoacc0m0.Fields("a0k04").Value
'      End If
'       'Add By Cheng 2003/04/24
'       '公司別
'       adoaccrpt105.Fields("r10522").Value = m_strA0k11
'      PaintLine ReportSum(4)
'      adoaccrpt105.UpdateBatch
'      adoaccrpt105.AddNew
'      adoaccrpt105.Fields("r10501").Value = strUserNum
'      adoaccrpt105.Fields("r10502").Value = Counter
'      If Text6 <> "" Then
'         adoaccrpt105.Fields("r10503").Value = Text6
'         adoaccrpt105.Fields("r10504").Value = Text3
'      Else
'         adoaccrpt105.Fields("r10503").Value = adoacc0m0.Fields("a0k03").Value
'         adoaccrpt105.Fields("r10504").Value = adoacc0m0.Fields("a0k04").Value
'      End If
'      adoaccrpt105.Fields("r10510").Value = ReportSum(24)
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         adoaccrpt105.Fields("r10513").Value = "0"
'      Else
'         adoaccrpt105.Fields("r10513").Value = adoaccsum.Fields(0).Value
'      End If
'      If IsNull(adoaccsum.Fields(1).Value) Then
'         adoaccrpt105.Fields("r10515").Value = "0"
'      Else
'         adoaccrpt105.Fields("r10515").Value = adoaccsum.Fields(1).Value
'      End If
'      If IsNull(adoaccsum.Fields(2).Value) Then
'         adoaccrpt105.Fields("r10514").Value = "0"
'      Else
'         adoaccrpt105.Fields("r10514").Value = adoaccsum.Fields(2).Value
'      End If
'      If IsNull(adoaccsum.Fields(3).Value) Then
'         adoaccrpt105.Fields("r10516").Value = "0"
'      Else
'         adoaccrpt105.Fields("r10516").Value = adoaccsum.Fields(3).Value
'      End If
'    'Add By Cheng 2003/04/24
'    '公司別
'      adoaccrpt105.Fields("r10522").Value = m_strA0k11
'      adoaccrpt105.UpdateBatch
'   End If
'   adoaccsum.Close
'   adoacc0m0.MoveNext
'End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   If Len(Text6) = 6 Then
      Text6 = AfterZero(Text6)
   End If
End Sub

'2013/5/13 cancel by sonia 沒有用到
''*************************************************
''  抬頭列印--未收
''
''*************************************************
'Private Sub PrintHead1()
'   Printer.FontSize = 16
'   Printer.CurrentX = 5000
'   Printer.CurrentY = 1000
'   Printer.Print ReportTitle(105)
'   Printer.CurrentX = 9000
'   Printer.CurrentY = 1000
''    'Modify By Cheng 2003/03/12
''    '加個人或公司別
'   Printer.Print "(未收)"
''   Printer.Print "(未收)" & IIf(Me.Text12.Text = "1", " (個人)", IIf(Me.Text12.Text = "2", " (公司)", ""))
'   Printer.FontSize = 12
'   Printer.CurrentX = 5500
'   Printer.CurrentY = 1500
'   Printer.Print "帳款日期: "
'   Printer.CurrentX = 6600
'   Printer.CurrentY = 1500
'   Printer.Print MaskEdBox1.Text & " ~ " & MaskEdBox2.Text
'   Printer.CurrentX = 500
'   Printer.CurrentY = 2100
'   Printer.Print "列印人員: "
'   Printer.CurrentX = 1800
'   Printer.CurrentY = 2100
'   Printer.Print StaffQuery(strUserNum)
'   Printer.CurrentX = 12000
'   Printer.CurrentY = 2100
'   Printer.Print "列印日期: "
'   Printer.CurrentX = 13300
'   Printer.CurrentY = 2100
'   Printer.Print IIf(Mid(CFDate(ACDate(ServerDate)), 1, 1) = "0", Mid(CFDate(ACDate(ServerDate)), 2, 8), CFDate(ACDate(ServerDate)))
'    'Add By Cheng 2003/04/24
'    '依公司別跳頁
'    If Me.Text12.Text = "Y" Then
'        Printer.CurrentX = 500
'        Printer.CurrentY = 2400
'        Printer.Print "公  司  別:　" & GetCompanyName("" & adoaccrpt105.Fields("r10522").Value)
'    End If
'   Printer.CurrentX = 12000
'   Printer.CurrentY = 2400
'   Printer.Print "頁次: "
'   Printer.CurrentX = 13300
'   Printer.CurrentY = 2400
'   Printer.Print intPage
'   Printer.CurrentX = 500
'   Printer.CurrentY = 2700
'   Printer.Print "客戶代號: "
'   Printer.CurrentX = 1800
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt105.Fields("r10503").Value) = False Then
'      Printer.Print adoaccrpt105.Fields("r10503").Value
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 3000
'   Printer.CurrentY = 2700
'   Printer.Print "客戶名稱: "
'   Printer.CurrentX = 4100
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt105.Fields("cu04").Value) = False Then
'      Printer.Print adoaccrpt105.Fields("cu04").Value
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 8000
'   Printer.CurrentY = 2700
'   Printer.Print "收據抬頭: "
'   Printer.CurrentX = 9100
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt105.Fields("r10504").Value) = False Then
'      Printer.Print adoaccrpt105.Fields("r10504").Value
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 500
'   Printer.CurrentY = 3300
'   Printer.Print "收據日期"
'   Printer.CurrentX = 1600
'   Printer.CurrentY = 3300
'   Printer.Print "收據號碼"
'   Printer.CurrentX = 2700
'   Printer.CurrentY = 3300
'   Printer.Print "智權人員"
'   Printer.CurrentX = 3600
'   Printer.CurrentY = 3300
'   Printer.Print "本所案號"
'    'Add By Cheng 2003/12/12
'    If Me.Text13.Text = "Y" Then
'        Printer.CurrentX = 5400
'        Printer.CurrentY = 3300
'        Printer.Print "客戶案件案號"
'    End If
'    'End
'   Printer.CurrentX = 5400 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "案件名稱"
'   Printer.CurrentX = 8500 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "案件性質"
'   Printer.CurrentX = 9600 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "申請國家"
'   Printer.CurrentX = 10900 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "服務費"
'   Printer.CurrentX = 12400 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "規費"
'   Printer.CurrentX = 13900 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "應收金額"
'   Printer.CurrentX = 15400 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "已收金額"
'   Printer.CurrentX = 16900 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "未收金額"
'   Printer.Line (500, 3700)-(17800 + IIf(Me.Text13.Text = "Y", 1920, 0), 3700)
'End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1 = "X"
   Text2 = "X"
   'Modify By Sindy 2016/6/8
   cboTitle.Text = ""
'   Text3 = ""
   '2016/6/8 END
   Text4 = ""
   Text11 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text5 = ""
   Text6 = ""
   Text7 = ""
   Text8 = ""
   Text9 = ""
   Text10 = ""
   Me.Text12.Text = "Y"
   Me.Text13.Text = ""
   Me.Text14.Text = "1"
   Me.Text15.Text = "J"  'modify by sonia 2014/11/13 原預設1-8
   'Add by Amy 2023/08/14
   Me.TxtDeptS = ""
   Me.TxtDeptE = ""
   'end 2023/08/14
   Text1.SetFocus
End Sub

'2013/5/13 cancel by sonia 沒有用到
''*************************************************
'' 列印報表 (未收)
''
''*************************************************
'Public Sub FormPrint1()
'   intCounter = 0
'   intPage = 0
'   strCustomerNo = ""
'    m_strCompany = ""
'   adoaccrpt105.CursorLocation = adUseClient
'   adoaccrpt105.Open "select * from accrpt105, customer where substr(r10503, 1, 8) = cu01 (+) and substr(r10503, 9, 1) = cu02 (+) and r10501 = '" & strUserNum & "' order by r10501 asc, r10502 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Do While adoaccrpt105.EOF = False
'        'Modify By Cheng 2003/07/10
''        If strCustomerNo <> adoaccrpt105.Fields("r10504").Value Then
'        If strCustomerNo <> "" & adoaccrpt105.Fields("r10503").Value & adoaccrpt105.Fields("r10504").Value Then
'            intCounter = 0
'            intPage = intPage + 1
'            If strCustomerNo <> "" Then
'                Printer.NewPage
'            End If
'            PrintHead1
'            'Modify By Cheng 2003/07/10
''            strCustomerNo = adoaccrpt105.Fields("r10504").Value
'            strCustomerNo = "" & adoaccrpt105.Fields("r10503").Value & adoaccrpt105.Fields("r10504").Value
'            m_strCompany = "" & adoaccrpt105.Fields("r10522").Value
'        'Add By Cheng 2003/04/24
'        '若公司別不同
'        ElseIf m_strCompany <> "" & adoaccrpt105.Fields("r10522").Value Then
'            If Me.Text12.Text = "Y" Then
'                intCounter = 0
'                intPage = intPage + 1
'                Printer.NewPage
'                PrintHead1
'                m_strCompany = "" & adoaccrpt105.Fields("r10522").Value
'            End If
'        End If
'      If intCounter >= 35 Then
'         intCounter = 0
'         intPage = intPage + 1
'         Printer.NewPage
'         PrintHead1
'      End If
'      Printer.CurrentX = 500
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10505").Value) = False Then
'         Printer.Print IIf(Mid(CFDate(adoaccrpt105.Fields("r10505").Value), 1, 1) = "0", Mid(CFDate(adoaccrpt105.Fields("r10505").Value), 2, 8), CFDate(adoaccrpt105.Fields("r10505").Value))
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 1600
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10507").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10507").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 2700
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10508").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10508").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 3600
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10509").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10509").Value
'      Else
'         Printer.Print ""
'      End If
'        'Add By Cheng 2003/12/12
'        '若要列印客戶案件案號
'        If Me.Text13.Text = "Y" Then
'            Printer.CurrentX = 5400
'            Printer.CurrentY = 3800 + intCounter * 300
'            If IsNull(adoaccrpt105.Fields("r10523").Value) = False Then
'               Printer.Print adoaccrpt105.Fields("r10523").Value
'            Else
'               Printer.Print ""
'            End If
'        End If
'        'End
'      Printer.CurrentX = 5400 + IIf(Me.Text13.Text = "Y", 1920, 0)
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10510").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10510").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 8500 + IIf(Me.Text13.Text = "Y", 1920, 0)
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10511").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10511").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 9600 + IIf(Me.Text13.Text = "Y", 1920, 0)
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10512").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10512").Value
'      Else
'         Printer.Print ""
'      End If
'      If IsNull(adoaccrpt105.Fields("r10513").Value) = False Then
'         strAmount = Format(adoaccrpt105.Fields("r10513").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 11800 - intLength + IIf(Me.Text13.Text = "Y", 1920, 0)
'         Printer.CurrentY = 3800 + intCounter * 300
'         Printer.Print strAmount
'      End If
'      If IsNull(adoaccrpt105.Fields("r10514").Value) = False Then
'         strAmount = Format(adoaccrpt105.Fields("r10514").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 13300 - intLength + IIf(Me.Text13.Text = "Y", 1920, 0)
'         Printer.CurrentY = 3800 + intCounter * 300
'         Printer.Print strAmount
'      Else
'         Printer.Print ""
'      End If
'      If IsNull(adoaccrpt105.Fields("r10515").Value) = False Then
'         strAmount = Format(adoaccrpt105.Fields("r10515").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 14800 - intLength + IIf(Me.Text13.Text = "Y", 1920, 0)
'         Printer.CurrentY = 3800 + intCounter * 300
'         Printer.Print strAmount
'      Else
'         Printer.Print ""
'      End If
'      If IsNull(adoaccrpt105.Fields("r10516").Value) = False Then
'         strAmount = Format(adoaccrpt105.Fields("r10516").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 16300 - intLength + IIf(Me.Text13.Text = "Y", 1920, 0)
'         Printer.CurrentY = 3800 + intCounter * 300
'         Printer.Print strAmount
'      Else
'         Printer.Print ""
'      End If
'      If IsNull(adoaccrpt105.Fields("r10517").Value) = False Then
'         strAmount = Format(adoaccrpt105.Fields("r10517").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 17800 - intLength + IIf(Me.Text13.Text = "Y", 1920, 0)
'         Printer.CurrentY = 3800 + intCounter * 300
'         Printer.Print strAmount
'      Else
'         Printer.Print ""
'      End If
'      intCounter = intCounter + 1
'      adoaccrpt105.MoveNext
'   Loop
'   adoaccrpt105.Close
'   Printer.EndDoc
'End Sub

'Add by Morgan 2004/11/19
Private Sub PrintDetail(ByRef arrData() As String)

   Dim stAmount As String
   
   lngCurrentY = lngCurrentY + m_iVPad
   
   lngCurrentX = 300
   Printer.CurrentY = lngCurrentY: Printer.CurrentX = lngCurrentX
   Printer.Print arrData(1)   '"收款日期"
   
   lngCurrentX = lngCurrentX + m_iColWidth(1)
   Printer.CurrentY = lngCurrentY: Printer.CurrentX = lngCurrentX
   Printer.Print arrData(2)   '"收據號碼"
   
   lngCurrentX = lngCurrentX + m_iColWidth(2)
   Printer.CurrentY = lngCurrentY: Printer.CurrentX = lngCurrentX
   Printer.Print arrData(3)   '"智權人員"
   
   lngCurrentX = lngCurrentX + m_iColWidth(3)
   Printer.CurrentY = lngCurrentY: Printer.CurrentX = lngCurrentX
   'If Me.Text13.Text = "Y" Then   '2013/5/13 cancel by sonia
      Printer.Print arrData(4)   '"客戶案件案號"
      lngCurrentX = lngCurrentX + m_iColWidth(4)
      Printer.CurrentY = lngCurrentY: Printer.CurrentX = lngCurrentX
   'End If                        '2013/5/13 cancel by sonia
   Printer.Print Left(arrData(5), 10)  '"案件名稱"
   
   lngCurrentX = lngCurrentX + m_iColWidth(5)
   Printer.CurrentY = lngCurrentY: Printer.CurrentX = lngCurrentX
   Printer.Print Left(arrData(6), 4)  '"案件性質"
   
   lngCurrentX = lngCurrentX + m_iColWidth(6)
   Printer.CurrentY = lngCurrentY: Printer.CurrentX = lngCurrentX
   Printer.Print Left(arrData(7), 5)  '"申請國家"
   
   If Val(arrData(8)) > 0 Then
      stAmount = Format(Val(arrData(8)), DDollar)
   Else
      stAmount = arrData(8)
   End If
   lngCurrentX = lngCurrentX + m_iColWidth(7)
   Printer.CurrentX = lngCurrentX + m_iColWidth(8) - 50 - Printer.TextWidth(stAmount)
   Printer.CurrentY = lngCurrentY
   Printer.Print stAmount   '"收款金額"
   
   If Val(arrData(9)) > 0 Then
      stAmount = Format(Val(arrData(9)), DDollar)
   Else
      stAmount = arrData(9)
   End If
   lngCurrentX = lngCurrentX + m_iColWidth(8)
   Printer.CurrentX = lngCurrentX + m_iColWidth(9) - 50 - Printer.TextWidth(stAmount)
   Printer.CurrentY = lngCurrentY
   Printer.Print stAmount   '"可扣稅額"
   
   If Val(arrData(10)) > 0 Then
      stAmount = Format(Val(arrData(10)), DDollar)
   Else
      stAmount = arrData(10)
   End If
   lngCurrentX = lngCurrentX + m_iColWidth(9)
   Printer.CurrentX = lngCurrentX + m_iColWidth(10) - 50 - Printer.TextWidth(stAmount)
   Printer.CurrentY = lngCurrentY
   Printer.Print stAmount  '"收款扣繳稅　　額"
   
   If Val(arrData(11)) > 0 Then
      stAmount = Format(Val(arrData(11)), DDollar)
   Else
      stAmount = arrData(11)
   End If
   lngCurrentX = lngCurrentX + m_iColWidth(10)
   Printer.CurrentX = lngCurrentX + m_iColWidth(11) - 50 - Printer.TextWidth(stAmount)
   Printer.CurrentY = lngCurrentY
   Printer.Print stAmount  '"未扣稅額"
      
   lngCurrentX = lngCurrentX + m_iColWidth(11)
   If Text5.Text = "1" Or Text5.Text = "5" Then
      If Val(arrData(12)) > 0 Then
         stAmount = Format(Val(arrData(12)), DDollar)
      Else
         stAmount = arrData(12)
      End If
      Printer.CurrentX = lngCurrentX + m_iColWidth(12) - 50 - Printer.TextWidth(stAmount)
      Printer.CurrentY = lngCurrentY
      Printer.Print stAmount  '"未收金額"
      
      If Text5.Text = "5" Then
         lngCurrentX = lngCurrentX + m_iColWidth(12)
         If Val(arrData(13)) > 0 Then
            stAmount = Format(Val(arrData(13)), DDollar)
         Else
            stAmount = arrData(13)
         End If
         Printer.CurrentX = lngCurrentX + m_iColWidth(13) - 50 - Printer.TextWidth(stAmount)
         Printer.CurrentY = lngCurrentY
         Printer.Print stAmount  '"票據金額"
         
         lngCurrentX = lngCurrentX + m_iColWidth(13)
         Printer.CurrentY = lngCurrentY: Printer.CurrentX = lngCurrentX
         Printer.Print arrData(14)
         
         lngCurrentX = lngCurrentX + m_iColWidth(14)
         Printer.CurrentY = lngCurrentY: Printer.CurrentX = lngCurrentX
         Printer.Print arrData(15)
      End If
      
   ElseIf Text5.Text = "2" Then
      Printer.CurrentY = lngCurrentY: Printer.CurrentX = lngCurrentX
      Printer.Print arrData(12)  '"傳票編號"
   End If
   
End Sub

Private Sub SetColumnWidth()
   Erase m_iColWidth
   If Text5.Text = "5" Then
      m_iColWidth(1) = 800
      m_iColWidth(2) = 950
      m_iColWidth(3) = 800
      m_iColWidth(4) = 1200
      m_iColWidth(5) = 1900
      m_iColWidth(6) = 900
      m_iColWidth(7) = 900
      m_iColWidth(8) = 950
      m_iColWidth(9) = 950
      m_iColWidth(10) = 950
      m_iColWidth(11) = 950
      m_iColWidth(12) = 950
      m_iColWidth(13) = 950
      m_iColWidth(14) = 950
      m_iColWidth(15) = 900
   Else
      m_iColWidth(1) = 1000
      m_iColWidth(2) = 1300
      m_iColWidth(3) = 1000
      m_iColWidth(4) = 1500
      m_iColWidth(5) = 2400
      m_iColWidth(6) = 1100
      m_iColWidth(7) = 1100
      m_iColWidth(8) = 1200
      m_iColWidth(9) = 1200
      m_iColWidth(10) = 1200
      m_iColWidth(11) = 1200
      m_iColWidth(12) = 1200
   End If
End Sub
Private Sub SetColumnName()
   Erase m_sColumn
   
   m_iVPad = 300: m_iHPad = 100: m_iMaxLine = 22
   m_sColumn(0) = ReportTitle(105)  '報表名稱
   m_sColumn(1) = "收據日期"     '欄位1
   m_sColumn(2) = "收據號碼"     '欄位2
   m_sColumn(3) = "智權人員"       '欄位3
   '2013/5/13 modify by sonia
   'm_sColumn(4) = "客戶案件案號" '欄位4
   If Me.Text13.Text = "Y" Then
      m_sColumn(4) = "客戶案件案號" '欄位4
   Else
      m_sColumn(4) = "本所案號" '欄位4
   End If
   '2013/5/13 end
   m_sColumn(5) = "案件名稱"     '欄位5
   m_sColumn(6) = "案件性質"     '欄位6
   m_sColumn(7) = "申請國家"     '欄位7
   m_sColumn(8) = "服務費"       '欄位8
   m_sColumn(9) = "規費"         '欄位9
   m_sColumn(10) = "應收金額"    '欄位10
   m_sColumn(11) = "已收金額"    '欄位11
   m_sColumn(12) = "未收金額"    '欄位12
   
   Select Case Text5.Text
      Case "1"
         m_sColumn(0) = m_sColumn(0) & "　(未收)"
      Case "2"
         m_sColumn(0) = m_sColumn(0) & "　(收回)"
         m_sColumn(1) = "收款日期"
         m_sColumn(8) = "收款金額"
         m_sColumn(9) = "可扣稅額"
         m_sColumn(10) = "收款扣繳稅　　額"
         m_sColumn(11) = "未扣稅額"
         m_sColumn(12) = "傳票編號"    '欄位12
      Case "3"
         m_sColumn(0) = m_sColumn(0) & "　(往來日期)"
         m_sColumn(1) = "單據日期"
         m_sColumn(8) = "應收金額"
         m_sColumn(9) = "收款金額"
         m_sColumn(10) = "收款扣繳稅　　額"
         m_sColumn(11) = "銷帳退費"
         m_sColumn(12) = ""
      Case "4"
         m_sColumn(0) = m_sColumn(0) & "　(交易情形)"
         m_sColumn(1) = "單據日期"
         m_sColumn(8) = "應收金額" & "應付金額"
         m_sColumn(9) = "收款金額" & "付款金額"
         m_sColumn(10) = "收款扣繳稅　　額"
         m_sColumn(11) = "銷帳退費"
         m_sColumn(12) = ""
      Case "5"
         m_sColumn(0) = m_sColumn(0) & "　(函證用未收)"
         m_sColumn(13) = "票據金額"
         m_sColumn(14) = "單據到期日"
         m_sColumn(15) = "收款單據"
         m_iVPad = 250: m_iHPad = 75: m_iMaxLine = 30
   End Select
   
   SetColumnWidth
End Sub
'*************************************************
'  抬頭列印
'
'*************************************************
'Modify  by Morgan 2004/11/19 改成A4橫印格式
Private Sub PrintHead()
   Dim iNext As Integer
      
   If Text5.Text = "5" Then
      Printer.FontSize = 12
      lngCurrentY = 1000
      Printer.CurrentY = lngCurrentY: Printer.CurrentX = 5000
      Printer.Print m_sColumn(0) '報表名稱
      
      Printer.FontSize = 9
      lngCurrentY = lngCurrentY + 300
   Else
      Printer.FontSize = 14
      lngCurrentY = 1000
      Printer.CurrentY = lngCurrentY: Printer.CurrentX = 5000
      Printer.Print m_sColumn(0) '報表名稱
      
      Printer.FontSize = 11
      lngCurrentY = lngCurrentY + 500
   End If
   
   'Modify by Morgan 2005/5/23 所有條件都要印,照畫面順序
   'Printer.CurrentY = lngCurrentY: Printer.CurrentX = 5500
   'Printer.Print "帳款日期:　" & MaskEdBox1.Text & " ~ " & MaskEdBox2.Text
   
   iNext = 0
   If txtSys <> "" Then
      Printer.CurrentY = lngCurrentY + iNext * m_iVPad: Printer.CurrentX = 5500
      Printer.Print "系統類別:　" & txtSys
      iNext = iNext + 1
   End If
   
   If Text1 <> "" Or Text2 <> "" Then
      Printer.CurrentY = lngCurrentY + iNext * m_iVPad: Printer.CurrentX = 5500
      Printer.Print "客戶代號:　" & Text1 & " ~ " & Text2
      iNext = iNext + 1
   End If
   
   'Modify By Sindy 2016/6/8
   '收據抬頭
'   If Text3 <> "" Then
'      Printer.CurrentY = lngCurrentY + iNext * m_iVPad: Printer.CurrentX = 5500
'      Printer.Print "收據抬頭:　" & Text3
'      iNext = iNext + 1
'   End If
   If cboTitle.Text <> "" Then
      Printer.CurrentY = lngCurrentY + iNext * m_iVPad: Printer.CurrentX = 5500
      Printer.Print "收據抬頭:　" & cboTitle.Text
      iNext = iNext + 1
   End If
   '2016/6/8 END
   
   If Text4 <> "" Then
      Printer.CurrentY = lngCurrentY + iNext * m_iVPad: Printer.CurrentX = 5500
      Printer.Print "智權人員:　" & Text4 & " " & Text11
      iNext = iNext + 1
   End If
   
   If MaskEdBox1.Text <> "___/__/__" Or MaskEdBox2.Text <> "___/__/__" Then
      Printer.CurrentY = lngCurrentY + iNext * m_iVPad: Printer.CurrentX = 5500
      Printer.Print "帳款日期:　" & MaskEdBox1.Text & " ~ " & MaskEdBox2.Text
      iNext = iNext + 1
   End If
   
   If Text7 <> "" Then
      Printer.CurrentY = lngCurrentY + iNext * m_iVPad: Printer.CurrentX = 5500
      Printer.Print "本所案號:　" & Text7 & Text8 & Text9 & Text10
      iNext = iNext + 1
   End If
   '2005/5/23 end
   
   
   lngCurrentY = lngCurrentY + 2 * m_iVPad
   Printer.CurrentY = lngCurrentY: Printer.CurrentX = 300
   Printer.Print "列印人員:　" & StaffQuery(strUserNum)

   Printer.CurrentY = lngCurrentY: Printer.CurrentX = 11000
   Printer.Print "列印日期:　" & IIf(Mid(CFDate(ACDate(ServerDate)), 1, 1) = "0", Mid(CFDate(ACDate(ServerDate)), 2, 8), CFDate(ACDate(ServerDate)))
   
   lngCurrentY = lngCurrentY + m_iVPad
   '依公司別跳頁
   If Me.Text12.Text = "Y" Then
       Printer.CurrentY = lngCurrentY: Printer.CurrentX = 300
       'Modify By Sindy 2020/3/30
       'Printer.Print "公  司  別:　" & GetCompanyName("" & adoaccrpt105.Fields("r10522").Value)
       Printer.Print "公  司  別:　" & A0802Query("" & adoaccrpt105.Fields("r10522").Value)
       '2020/3/30 END
   End If
   
   Printer.CurrentY = lngCurrentY: Printer.CurrentX = 11000
   Printer.Print "頁次:　" & Format(intPage)
   
   lngCurrentY = lngCurrentY + m_iVPad
   Printer.CurrentY = lngCurrentY: Printer.CurrentX = 300
   Printer.Print "客戶名稱:　" & adoaccrpt105.Fields("r10503").Value & "　" & adoaccrpt105.Fields("cu04").Value
      
   lngCurrentY = lngCurrentY + m_iVPad
   Printer.CurrentY = lngCurrentY: Printer.CurrentX = 300
   Printer.Print "收據抬頭:　" & adoaccrpt105.Fields("r10504").Value
   
   lngCurrentY = lngCurrentY + 2 * m_iVPad
   
   lngCurrentX = 300
   Printer.CurrentY = lngCurrentY: Printer.CurrentX = lngCurrentX
   Printer.Print m_sColumn(1)
   
   lngCurrentX = lngCurrentX + m_iColWidth(1)
   Printer.CurrentY = lngCurrentY: Printer.CurrentX = lngCurrentX
   Printer.Print m_sColumn(2)
   
   lngCurrentX = lngCurrentX + m_iColWidth(2)
   Printer.CurrentY = lngCurrentY: Printer.CurrentX = lngCurrentX
   Printer.Print m_sColumn(3)
   
   lngCurrentX = lngCurrentX + m_iColWidth(3)
   Printer.CurrentY = lngCurrentY: Printer.CurrentX = lngCurrentX
   
   'If Me.Text13.Text = "Y" Then   '2013/5/13 cancel by sonia
      Printer.Print m_sColumn(4)
      lngCurrentX = lngCurrentX + m_iColWidth(4)
      Printer.CurrentY = lngCurrentY: Printer.CurrentX = lngCurrentX
   'End If                         '2013/5/13 cancel by sonia
   Printer.Print m_sColumn(5)
   
   lngCurrentX = lngCurrentX + m_iColWidth(5)
   Printer.CurrentY = lngCurrentY: Printer.CurrentX = lngCurrentX
   Printer.Print m_sColumn(6)
   
   lngCurrentX = lngCurrentX + m_iColWidth(6)
   Printer.CurrentY = lngCurrentY: Printer.CurrentX = lngCurrentX
   Printer.Print m_sColumn(7)
      
   lngCurrentX = lngCurrentX + m_iColWidth(7)
   Printer.CurrentX = lngCurrentX + m_iColWidth(8) - m_iHPad - Printer.TextWidth(Left(m_sColumn(8), 4))
   Printer.CurrentY = lngCurrentY
   Printer.Print Left(m_sColumn(8), 4)
   Printer.CurrentX = lngCurrentX + m_iColWidth(8) - m_iHPad - Printer.TextWidth(Mid(m_sColumn(8), 5))
   Printer.CurrentY = lngCurrentY + m_iVPad
   Printer.Print Mid(m_sColumn(8), 5)
   
   
   lngCurrentX = lngCurrentX + m_iColWidth(8)
   Printer.CurrentX = lngCurrentX + m_iColWidth(9) - m_iHPad - Printer.TextWidth(Left(m_sColumn(9), 4))
   Printer.CurrentY = lngCurrentY
   Printer.Print Left(m_sColumn(9), 4)
   Printer.CurrentX = lngCurrentX + m_iColWidth(9) - m_iHPad - Printer.TextWidth(Mid(m_sColumn(9), 5))
   Printer.CurrentY = lngCurrentY + m_iVPad
   Printer.Print Mid(m_sColumn(9), 5)
   
   lngCurrentX = lngCurrentX + m_iColWidth(9)
   Printer.CurrentX = lngCurrentX + m_iColWidth(10) - m_iHPad - Printer.TextWidth(Left(m_sColumn(10), 4))
   Printer.CurrentY = lngCurrentY
   Printer.Print Left(m_sColumn(10), 4)
   Printer.CurrentX = lngCurrentX + m_iColWidth(10) - m_iHPad - Printer.TextWidth(Mid(m_sColumn(10), 5))
   Printer.CurrentY = lngCurrentY + m_iVPad
   Printer.Print Mid(m_sColumn(10), 5)
   
   lngCurrentX = lngCurrentX + m_iColWidth(10)
   Printer.CurrentX = lngCurrentX + m_iColWidth(11) - m_iHPad - Printer.TextWidth(m_sColumn(11))
   Printer.CurrentY = lngCurrentY
   Printer.Print m_sColumn(11)
   
   Select Case Text5.Text
      Case "1", "5"
         lngCurrentX = lngCurrentX + m_iColWidth(11)
         Printer.CurrentY = lngCurrentY
         Printer.CurrentX = lngCurrentX + m_iColWidth(12) - m_iHPad - Printer.TextWidth(m_sColumn(12))
         Printer.Print m_sColumn(12)
         
         If Text5.Text = "1" Then
            lngCurrentY = lngCurrentY + m_iVPad - 100
            Printer.Line (300, lngCurrentY + 150)-(lngCurrentX + m_iColWidth(12), lngCurrentY + 150)
         Else
            lngCurrentX = lngCurrentX + m_iColWidth(12)
            Printer.CurrentY = lngCurrentY
            Printer.CurrentX = lngCurrentX + m_iColWidth(13) - m_iHPad - Printer.TextWidth(m_sColumn(13))
            Printer.Print m_sColumn(13)
            
            lngCurrentX = lngCurrentX + m_iColWidth(13)
            Printer.CurrentY = lngCurrentY: Printer.CurrentX = lngCurrentX
            Printer.Print m_sColumn(14)
            
            lngCurrentX = lngCurrentX + m_iColWidth(14)
            Printer.CurrentY = lngCurrentY: Printer.CurrentX = lngCurrentX
            Printer.Print m_sColumn(15)
            lngCurrentY = lngCurrentY + m_iVPad - 100
            Printer.Line (300, lngCurrentY + 100)-(lngCurrentX + m_iColWidth(11), lngCurrentY + 100)
         End If
         
      Case "2"
         lngCurrentX = lngCurrentX + m_iColWidth(11)
         Printer.CurrentY = lngCurrentY: Printer.CurrentX = lngCurrentX
         Printer.Print m_sColumn(12)
         
         lngCurrentY = lngCurrentY + 2 * m_iVPad - 100
         Printer.Line (300, lngCurrentY + 150)-(lngCurrentX + m_iColWidth(12), lngCurrentY + 150)
         
      Case Else
         lngCurrentY = lngCurrentY + 2 * m_iVPad - 100
         Printer.Line (300, lngCurrentY + 150)-(lngCurrentX + m_iColWidth(11), lngCurrentY + 150)
   End Select
   
End Sub

Private Sub PrintTail()
   lngCurrentY = lngCurrentY + m_iVPad
   Select Case Text5.Text
      Case "1", "2"
         Printer.Line (300, lngCurrentY + 150)-(lngCurrentX + m_iColWidth(12), lngCurrentY + 150)
      Case "5"
         Printer.Line (300, lngCurrentY + 150)-(lngCurrentX + m_iColWidth(11), lngCurrentY + 150)
      Case Else
         Printer.Line (300, lngCurrentY + 150)-(lngCurrentX, lngCurrentY + 150)
   End Select
End Sub

'2013/5/13 cancel by sonia 沒有用到
''*************************************************
''  抬頭列印
''
''*************************************************
'Private Sub PrintHead2()
'   Printer.FontSize = 16
'   Printer.CurrentX = 5000
'   Printer.CurrentY = 1000
'   Printer.Print ReportTitle(105)
'   Printer.CurrentX = 9000
'   Printer.CurrentY = 1000
'    'Modify By Cheng 2003/03/12
'    '加個人或公司別
'   Printer.Print "(收回)"
''   Printer.Print "(收回)" & IIf(Me.Text12.Text = "1", " (個人)", IIf(Me.Text12.Text = "2", " (公司)", ""))
'   Printer.FontSize = 12
'   Printer.CurrentX = 5500
'   Printer.CurrentY = 1500
'   Printer.Print "帳款日期: "
'   Printer.CurrentX = 6600
'   Printer.CurrentY = 1500
'   Printer.Print MaskEdBox1.Text & " ~ " & MaskEdBox2.Text
'   Printer.CurrentX = 500
'   Printer.CurrentY = 2100
'   Printer.Print "列印人員: "
'   Printer.CurrentX = 1800
'   Printer.CurrentY = 2100
'   Printer.Print StaffQuery(strUserNum)
'   Printer.CurrentX = 13000
'   Printer.CurrentY = 2100
'   Printer.Print "列印日期: "
'   Printer.CurrentX = 14300
'   Printer.CurrentY = 2100
'   Printer.Print IIf(Mid(CFDate(ACDate(ServerDate)), 1, 1) = "0", Mid(CFDate(ACDate(ServerDate)), 2, 8), CFDate(ACDate(ServerDate)))
'    'Add By Cheng 2003/04/24
'    '依公司別跳頁
'    If Me.Text12.Text = "Y" Then
'        Printer.CurrentX = 500
'        Printer.CurrentY = 2400
'        Printer.Print "公  司  別:　" & GetCompanyName("" & adoaccrpt105.Fields("r10522").Value)
'    End If
'   Printer.CurrentX = 13000
'   Printer.CurrentY = 2400
'   Printer.Print "頁次: "
'   Printer.CurrentX = 14300
'   Printer.CurrentY = 2400
'   Printer.Print intPage
'   Printer.CurrentX = 500
'   Printer.CurrentY = 2700
'   Printer.Print "客戶代號: "
'   Printer.CurrentX = 1800
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt105.Fields("r10503").Value) = False Then
'      Printer.Print adoaccrpt105.Fields("r10503").Value
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 3000
'   Printer.CurrentY = 2700
'   Printer.Print "客戶名稱: "
'   Printer.CurrentX = 4100
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt105.Fields("cu04").Value) = False Then
'      Printer.Print adoaccrpt105.Fields("cu04").Value
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 8000
'   Printer.CurrentY = 2700
'   Printer.Print "收據抬頭: "
'   Printer.CurrentX = 9100
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt105.Fields("r10504").Value) = False Then
'      Printer.Print adoaccrpt105.Fields("r10504").Value
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 0
'   Printer.CurrentY = 3300
'   Printer.Print "收款日期"
'   Printer.CurrentX = 1100
'   Printer.CurrentY = 3300
'   Printer.Print "收款單號"
'   Printer.CurrentX = 2200
'   Printer.CurrentY = 3300
'   Printer.Print "收據號碼"
'   Printer.CurrentX = 3300
'   Printer.CurrentY = 3300
'   Printer.Print "智權人員"
'   Printer.CurrentX = 4100
'   Printer.CurrentY = 3300
'   Printer.Print "本所案號"
'    'Add By Cheng 2003/12/12
'    If Me.Text13.Text = "Y" Then
'        Printer.CurrentX = 5500
'        Printer.CurrentY = 3300
'        Printer.Print "客戶案件案號"
'    End If
'    'End
'   Printer.CurrentX = 5500 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "案件名稱"
'   Printer.CurrentX = 9100 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "案件性質"
'   Printer.CurrentX = 10400 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "申請國家"
''   Printer.CurrentX = 11500
'   Printer.CurrentX = 11500 + 250 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "收款金額"
''   Printer.CurrentX = 12500
'   Printer.CurrentX = 12500 + 500 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "可扣稅額"
''   Printer.CurrentX = 13700
'   Printer.CurrentX = 13700 + 750 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "收款扣繳"
''   Printer.CurrentX = 13700
'   Printer.CurrentX = 13700 + 750 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3600
'   Printer.Print "稅　　額"
''   Printer.CurrentX = 14800
'   Printer.CurrentX = 14800 + 900 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "未扣稅額"
''   Printer.CurrentX = 15950
'   Printer.CurrentX = 15950 + 1150 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "傳票編號"
''   Printer.Line (0, 4000)-(17200, 4000)
'   Printer.Line (0, 4000)-(17200 + 1150 + IIf(Me.Text13.Text = "Y", 1920, 0), 4000)
'End Sub

'*************************************************
' 列印報表
'
'*************************************************
'Add by Morgan 2004/11/19
Public Sub FormPrint()
   
   Dim arrData(1 To 15) As String   '明細資料
   intCounter = 0
   intPage = 0
   strCustomerNo = ""
   m_strCompany = "" '初始化公司別
   adoaccrpt105.CursorLocation = adUseClient
   adoaccrpt105.Open "select a.*,cu04 from accrpt105 a, customer where substr(r10503, 1, 8) = cu01 (+) and substr(r10503, 9, 1) = cu02 (+) and r10501 = '" & strUserNum & "' order by r10501 asc, r10502 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccrpt105.RecordCount > 0 Then
      Printer.Orientation = vbPRORLandscape '橫印
      Do While adoaccrpt105.EOF = False
         Erase arrData
         If strCustomerNo <> "" & adoaccrpt105.Fields("r10503").Value & adoaccrpt105.Fields("r10504").Value Then
            If strCustomerNo <> "" Then
               PrintTail
               Printer.NewPage
            End If
            intCounter = 0
            intPage = intPage + 1
            PrintHead
            strCustomerNo = "" & adoaccrpt105.Fields("r10503").Value & adoaccrpt105.Fields("r10504").Value
            m_strCompany = "" & adoaccrpt105.Fields("r10522").Value
           '若公司別不同
           ElseIf m_strCompany <> "" & adoaccrpt105.Fields("r10522").Value Then
               If Me.Text12.Text = "Y" Then
                  PrintTail
                  Printer.NewPage
                  intCounter = 0
                  intPage = intPage + 1
                  PrintHead
                  m_strCompany = "" & adoaccrpt105.Fields("r10522").Value
               End If
         End If
         If intCounter >= m_iMaxLine Then
            PrintTail
            Printer.NewPage
            intCounter = 0
            intPage = intPage + 1
            PrintHead
         End If
         If IsNull(adoaccrpt105.Fields("r10505").Value) = False Then
            arrData(1) = IIf(Mid(CFDate(adoaccrpt105.Fields("r10505").Value), 1, 1) = "0", Mid(CFDate(adoaccrpt105.Fields("r10505").Value), 2, 8), CFDate(adoaccrpt105.Fields("r10505").Value))
         End If
         arrData(2) = "" & adoaccrpt105.Fields("r10507").Value
         arrData(3) = "" & adoaccrpt105.Fields("r10508").Value
         arrData(4) = "" & adoaccrpt105.Fields("r10523").Value
         arrData(5) = "" & adoaccrpt105.Fields("r10510").Value
         arrData(6) = "" & adoaccrpt105.Fields("r10511").Value
         arrData(7) = "" & adoaccrpt105.Fields("r10512").Value
         arrData(8) = "" & adoaccrpt105.Fields("r10513").Value
         arrData(9) = "" & adoaccrpt105.Fields("r10514").Value
         arrData(10) = "" & adoaccrpt105.Fields("r10515").Value
         arrData(11) = "" & adoaccrpt105.Fields("r10516").Value
         If Text5.Text = "1" Or Text5.Text = "5" Then
            arrData(12) = "" & adoaccrpt105.Fields("r10517").Value
         Else
            arrData(12) = "" & adoaccrpt105.Fields("r10518").Value
         End If
         If Text5.Text = "5" Then
            arrData(13) = "" & adoaccrpt105.Fields("r10519").Value
            arrData(14) = "" & adoaccrpt105.Fields("r10520").Value
            arrData(15) = "" & adoaccrpt105.Fields("r10521").Value
         End If
         intCounter = intCounter + 1
         Call PrintDetail(arrData)
         adoaccrpt105.MoveNext
      Loop
      
      PrintTail
      Printer.EndDoc
   End If
   adoaccrpt105.Close
End Sub

'2013/5/13 cancel by sonia 沒有用到
''*************************************************
'' 列印報表 (收回)
''
''*************************************************
'Public Sub FormPrint2()
'   intCounter = 0
'   intPage = 0
'   strCustomerNo = ""
'    'Add By Cheng 2003/04/24
'    '初始化公司別
'    m_strCompany = ""
'   adoaccrpt105.CursorLocation = adUseClient
'   adoaccrpt105.Open "select * from accrpt105, customer where substr(r10503, 1, 8) = cu01 (+) and substr(r10503, 9, 1) = cu02 (+) and r10501 = '" & strUserNum & "' order by r10501 asc, r10502 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Do While adoaccrpt105.EOF = False
'        'Modify By Cheng 2003/07/14
''      If strCustomerNo <> adoaccrpt105.Fields("r10504").Value Then
'      If strCustomerNo <> "" & adoaccrpt105.Fields("r10503").Value & adoaccrpt105.Fields("r10504").Value Then
'         intCounter = 0
'         intPage = intPage + 1
'         If strCustomerNo <> "" Then
'            Printer.NewPage
'         End If
'         PrintHead2
'        'Modify By Cheng 2003/07/14
''         strCustomerNo = adoaccrpt105.Fields("r10504").Value
'         strCustomerNo = "" & adoaccrpt105.Fields("r10503").Value & adoaccrpt105.Fields("r10504").Value
'            m_strCompany = "" & adoaccrpt105.Fields("r10522").Value
'        'Add By Cheng 2003/04/24
'        '若公司別不同
'        ElseIf m_strCompany <> "" & adoaccrpt105.Fields("r10522").Value Then
'            If Me.Text12.Text = "Y" Then
'                intCounter = 0
'                intPage = intPage + 1
'                Printer.NewPage
'                PrintHead2
'                m_strCompany = "" & adoaccrpt105.Fields("r10522").Value
'            End If
'      End If
'      If intCounter = 35 Then
'         intCounter = 0
'         intPage = intPage + 1
'         Printer.NewPage
'         PrintHead2
'      End If
'      Printer.CurrentX = 0
'      Printer.CurrentY = 4100 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10505").Value) = False Then
'         Printer.Print IIf(Mid(CFDate(adoaccrpt105.Fields("r10505").Value), 1, 1) = "0", Mid(CFDate(adoaccrpt105.Fields("r10505").Value), 2, 8), CFDate(adoaccrpt105.Fields("r10505").Value))
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 1000
'      Printer.CurrentY = 4100 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10506").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10506").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 2200
'      Printer.CurrentY = 4100 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10507").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10507").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 3300
'      Printer.CurrentY = 4100 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10508").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10508").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 4100
'      Printer.CurrentY = 4100 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10509").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10509").Value
'      Else
'         Printer.Print ""
'      End If
'        'Add By Cheng 2003/12/12
'        '若要列印客戶案件案號
'        If Me.Text13.Text = "Y" Then
'            Printer.CurrentX = 5500
'            Printer.CurrentY = 4100 + intCounter * 300
'            If IsNull(adoaccrpt105.Fields("r10523").Value) = False Then
'               Printer.Print adoaccrpt105.Fields("r10523").Value
'            Else
'               Printer.Print ""
'            End If
'        End If
'        'End
'      Printer.CurrentX = 5500 + IIf(Me.Text13.Text = "Y", 1920, 0)
'      Printer.CurrentY = 4100 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10510").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10510").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 9100 + IIf(Me.Text13.Text = "Y", 1920, 0)
'      Printer.CurrentY = 4100 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10511").Value) = False Then
'         Printer.Print MidB(adoaccrpt105.Fields("r10511").Value, 1, 10)
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 10400 + IIf(Me.Text13.Text = "Y", 1920, 0)
'      Printer.CurrentY = 4100 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10512").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10512").Value
'      Else
'         Printer.Print ""
'      End If
'      If IsNull(adoaccrpt105.Fields("r10513").Value) = False Then
'         strAmount = Format(adoaccrpt105.Fields("r10513").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
''         Printer.CurrentX = 12400 - intLength
'         Printer.CurrentX = 12400 - intLength + 250 + IIf(Me.Text13.Text = "Y", 1920, 0)
'         Printer.CurrentY = 4100 + intCounter * 300
'         Printer.Print strAmount
'      End If
'      If IsNull(adoaccrpt105.Fields("r10514").Value) = False Then
'         strAmount = Format(adoaccrpt105.Fields("r10514").Value, DDollar)
'         If strAmount = "" Then
'            strAmount = "0"
'         End If
'         intLength = Printer.TextWidth(strAmount)
''         Printer.CurrentX = 13500 - intLength
'         Printer.CurrentX = 13500 - intLength + 500 + IIf(Me.Text13.Text = "Y", 1920, 0)
'         Printer.CurrentY = 4100 + intCounter * 300
'         Printer.Print strAmount
'      End If
'      If IsNull(adoaccrpt105.Fields("r10515").Value) = False Then
'         strAmount = Format(adoaccrpt105.Fields("r10515").Value, DDollar)
'         If strAmount = "" Then
'            strAmount = "0"
'         End If
'         intLength = Printer.TextWidth(strAmount)
''         Printer.CurrentX = 14700 - intLength
'         Printer.CurrentX = 14700 - intLength + 750 + IIf(Me.Text13.Text = "Y", 1920, 0)
'         Printer.CurrentY = 4100 + intCounter * 300
'         Printer.Print strAmount
'      End If
'      If IsNull(adoaccrpt105.Fields("r10516").Value) = False Then
'         strAmount = Format(adoaccrpt105.Fields("r10516").Value, DDollar)
'         If strAmount = "" Then
'            strAmount = "0"
'         End If
'         intLength = Printer.TextWidth(strAmount)
''         Printer.CurrentX = 15900 - intLength
'         Printer.CurrentX = 15900 - intLength + 900 + IIf(Me.Text13.Text = "Y", 1920, 0)
'         Printer.CurrentY = 4100 + intCounter * 300
'         Printer.Print strAmount
'      End If
''      Printer.CurrentX = 15950
'      Printer.CurrentX = 15950 + 1150 + IIf(Me.Text13.Text = "Y", 1920, 0)
'      Printer.CurrentY = 4100 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10518").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10518").Value
'      Else
'         Printer.Print ""
'      End If
'      intCounter = intCounter + 1
'      adoaccrpt105.MoveNext
'   Loop
'   adoaccrpt105.Close
'   Printer.EndDoc
'End Sub

'2013/5/13 cancel by sonia 沒有用到
''*************************************************
''  列印類別-選擇往來日期之小計計算
''
''*************************************************
'Private Sub SubSelect3()
'Dim strSql As String
'
'   adoacc0k0.MovePrevious
'   adoaccsum.CursorLocation = adUseClient
'   If Text6 = "" Then
'      'Modify By Cheng 2003/04/28
'      If Me.Text12.Text = "Y" Then
'        adoaccsum.Open "select sum(to_number(nvl(r10513, 0))), sum(to_number(nvl(r10514, 0))), sum(to_number(nvl(r10515, 0))), sum(to_number(nvl(r10516, 0))) from accrpt105 where r10501 = '" & strUserNum & "' and r10507 is not null and r10504 = '" & adoacc0k0.Fields("Title").Value & "' and r10522='" & m_strA0k11 & "' ", adoTaie, adOpenStatic, adLockReadOnly
'      Else
'        adoaccsum.Open "select sum(to_number(nvl(r10513, 0))), sum(to_number(nvl(r10514, 0))), sum(to_number(nvl(r10515, 0))), sum(to_number(nvl(r10516, 0))) from accrpt105 where r10501 = '" & strUserNum & "' and r10507 is not null and r10504 = '" & adoacc0k0.Fields("Title").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'      End If
'   Else
'      'Modify By Cheng 2003/04/28
'      If Me.Text12.Text = "Y" Then
'        adoaccsum.Open "select sum(to_number(nvl(r10513, 0))), sum(to_number(nvl(r10514, 0))), sum(to_number(nvl(r10515, 0))), sum(to_number(nvl(r10516, 0))) from accrpt105 where r10501 = '" & strUserNum & "' and r10507 is not null and r10504 = '" & adoacc0k0.Fields("Title").Value & "' and r10522='" & m_strA0k11 & "' ", adoTaie, adOpenStatic, adLockReadOnly
'      Else
'        adoaccsum.Open "select sum(to_number(nvl(r10513, 0))), sum(to_number(nvl(r10514, 0))), sum(to_number(nvl(r10515, 0))), sum(to_number(nvl(r10516, 0))) from accrpt105 where r10501 = '" & strUserNum & "' and r10507 is not null and r10504 = '" & Text3 & "'", adoTaie, adOpenStatic, adLockReadOnly
'      End If
'   End If
'   If adoaccsum.RecordCount <> 0 And (adoaccsum.Fields(0).Value <> 0 Or adoaccsum.Fields(1).Value <> 0 Or adoaccsum.Fields(2).Value <> 0 Or adoaccsum.Fields(3).Value <> 0) Then
'      adoaccrpt105.AddNew
'      adoaccrpt105.Fields("r10501").Value = strUserNum
'      adoaccrpt105.Fields("r10502").Value = Counter
'      If Text6 <> "" Then
'         adoaccrpt105.Fields("r10503").Value = Text6
'         adoaccrpt105.Fields("r10504").Value = Text3
'      Else
'         adoaccrpt105.Fields("r10503").Value = adoacc0k0.Fields("Cust").Value
'         adoaccrpt105.Fields("r10504").Value = adoacc0k0.Fields("Title").Value
'      End If
'       'Add By Cheng 2003/04/28
'       adoaccrpt105.Fields("r10522").Value = m_strA0k11
'      PaintLine ReportSum(4)
'      adoaccrpt105.UpdateBatch
'      adoaccrpt105.AddNew
'      adoaccrpt105.Fields("r10501").Value = strUserNum
'      adoaccrpt105.Fields("r10502").Value = Counter
'      If Text6 <> "" Then
'         adoaccrpt105.Fields("r10503").Value = Text6
'         adoaccrpt105.Fields("r10504").Value = Text3
'      Else
'         adoaccrpt105.Fields("r10503").Value = adoacc0k0.Fields("Cust").Value
'         adoaccrpt105.Fields("r10504").Value = adoacc0k0.Fields("Title").Value
'      End If
'      adoaccrpt105.Fields("r10510").Value = ReportSum(24)
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         adoaccrpt105.Fields("r10513").Value = 0
'      Else
'         adoaccrpt105.Fields("r10513").Value = adoaccsum.Fields(0).Value
'      End If
'      If IsNull(adoaccsum.Fields(1).Value) Then
'         adoaccrpt105.Fields("r10514").Value = 0
'      Else
'         adoaccrpt105.Fields("r10514").Value = adoaccsum.Fields(1).Value
'      End If
'      If IsNull(adoaccsum.Fields(2).Value) Then
'         adoaccrpt105.Fields("r10515").Value = 0
'      Else
'         adoaccrpt105.Fields("r10515").Value = adoaccsum.Fields(2).Value
'      End If
'      If IsNull(adoaccsum.Fields(3).Value) Then
'         adoaccrpt105.Fields("r10516").Value = 0
'      Else
'         adoaccrpt105.Fields("r10516").Value = adoaccsum.Fields(3).Value
'      End If
'    'Add By Cheng 2003/04/28
'      adoaccrpt105.Fields("r10522").Value = m_strA0k11
'      adoaccrpt105.UpdateBatch
'   End If
'   adoaccsum.Close
'   adoacc0k0.MoveNext
'End Sub

'2013/5/13 cancel by sonia 沒有用到
''*************************************************
''  抬頭列印
''
''*************************************************
'Private Sub PrintHead3()
'   Printer.FontSize = 16
'   Printer.CurrentX = 5000
'   Printer.CurrentY = 1000
'   Printer.Print ReportTitle(105)
'   Printer.CurrentX = 9000
'   Printer.CurrentY = 1000
''    'Modify By Cheng 2003/03/12
''    '加個人或公司別
'   Printer.Print "(往來日期)"
''   Printer.Print "(往來日期)" & IIf(Me.Text12.Text = "1", " (個人)", IIf(Me.Text12.Text = "2", " (公司)", ""))
'   Printer.FontSize = 12
'   Printer.CurrentX = 5500
'   Printer.CurrentY = 1500
'   Printer.Print "帳款日期: "
'   Printer.CurrentX = 6600
'   Printer.CurrentY = 1500
'   Printer.Print MaskEdBox1.Text & " ~ " & MaskEdBox2.Text
'   Printer.CurrentX = 500
'   Printer.CurrentY = 2100
'   Printer.Print "列印人員: "
'   Printer.CurrentX = 1800
'   Printer.CurrentY = 2100
'   Printer.Print StaffQuery(strUserNum)
'   Printer.CurrentX = 12000
'   Printer.CurrentY = 2100
'   Printer.Print "列印日期: "
'   Printer.CurrentX = 13300
'   Printer.CurrentY = 2100
'   Printer.Print IIf(Mid(CFDate(ACDate(ServerDate)), 1, 1) = "0", Mid(CFDate(ACDate(ServerDate)), 2, 8), CFDate(ACDate(ServerDate)))
'    'Add By Cheng 2003/04/28
'    '依公司別跳頁
'    If Me.Text12.Text = "Y" Then
'        Printer.CurrentX = 500
'        Printer.CurrentY = 2400
'        Printer.Print "公  司  別:　" & GetCompanyName("" & adoaccrpt105.Fields("r10522").Value)
'    End If
'   Printer.CurrentX = 12000
'   Printer.CurrentY = 2400
'   Printer.Print "頁次: "
'   Printer.CurrentX = 13300
'   Printer.CurrentY = 2400
'   Printer.Print intPage
'   Printer.CurrentX = 500
'   Printer.CurrentY = 2700
'   Printer.Print "客戶代號: "
'   Printer.CurrentX = 1800
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt105.Fields("r10503").Value) = False Then
'      Printer.Print adoaccrpt105.Fields("r10503").Value
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 3000
'   Printer.CurrentY = 2700
'   Printer.Print "客戶名稱: "
'   Printer.CurrentX = 4100
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt105.Fields("cu04").Value) = False Then
'      Printer.Print adoaccrpt105.Fields("cu04").Value
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 8000
'   Printer.CurrentY = 2700
'   Printer.Print "收據抬頭: "
'   Printer.CurrentX = 9100
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt105.Fields("r10504").Value) = False Then
'      Printer.Print adoaccrpt105.Fields("r10504").Value
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 500
'   Printer.CurrentY = 3300
'   Printer.Print "單據日期"
'   Printer.CurrentX = 1600
'   Printer.CurrentY = 3300
'   Printer.Print "單據號碼"
'   Printer.CurrentX = 2700
'   Printer.CurrentY = 3300
'   Printer.Print "收據號碼"
'   Printer.CurrentX = 3800
'   Printer.CurrentY = 3300
'   Printer.Print "智權人員"
'   Printer.CurrentX = 4700
'   Printer.CurrentY = 3300
'   Printer.Print "本所案號"
'    'Add By Cheng 2003/12/12
'    If Me.Text13.Text = "Y" Then
'        Printer.CurrentX = 6100
'        Printer.CurrentY = 3300
'        Printer.Print "客戶案件案號"
'    End If
'    'End
'   Printer.CurrentX = 6100 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "案件名稱"
'   Printer.CurrentX = 9200 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "案件性質"
'   Printer.CurrentX = 10500 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "申請國家"
'   Printer.CurrentX = 11600 + 250 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "應收金額"
'   Printer.CurrentX = 13200 + 250 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "收款金額"
'   Printer.CurrentX = 14300 + 500 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "收款扣繳"
'   Printer.CurrentX = 15400 + 750 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "銷帳退費"
'   Printer.Line (500, 3700)-(16500 + 750 + IIf(Me.Text13.Text = "Y", 1920, 0), 3700)
'End Sub

'2013/5/13 cancel by sonia 沒有用到
''*************************************************
'' 列印報表 (往來日期)
''
''*************************************************
'Public Sub FormPrint3()
'Dim strSameDocNo As String
'
'   intCounter = 0
'   intPage = 0
'   strCustomerNo = ""
'    'Add By Cheng 2003/04/28
'    '初始化公司別
'    m_strCompany = ""
'   adoaccrpt105.CursorLocation = adUseClient
'   adoaccrpt105.Open "select * from accrpt105, customer where substr(r10503, 1, 8) = cu01 (+) and substr(r10503, 9, 1) = cu02 (+) and r10501 = '" & strUserNum & "' order by r10501 asc, r10502 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Do While adoaccrpt105.EOF = False
'      If strCustomerNo <> adoaccrpt105.Fields("r10504").Value Then
'         intCounter = 0
'         intPage = intPage + 1
'         If strCustomerNo <> "" Then
'            Printer.NewPage
'         End If
'         PrintHead3
'         strCustomerNo = adoaccrpt105.Fields("r10504").Value
'            m_strCompany = "" & adoaccrpt105.Fields("r10522").Value
'        'Add By Cheng 2003/04/24
'        '若公司別不同
'        ElseIf m_strCompany <> "" & adoaccrpt105.Fields("r10522").Value Then
'            If Me.Text12.Text = "Y" Then
'                intCounter = 0
'                intPage = intPage + 1
'                Printer.NewPage
'                PrintHead3
'                m_strCompany = "" & adoaccrpt105.Fields("r10522").Value
'            End If
'      End If
'      If intCounter = 35 Then
'         intCounter = 0
'         intPage = intPage + 1
'         Printer.NewPage
'         PrintHead3
'      End If
'      Printer.CurrentX = 500
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10505").Value) = False Then
'         Printer.Print IIf(Mid(CFDate(adoaccrpt105.Fields("r10505").Value), 1, 1) = "0", Mid(CFDate(adoaccrpt105.Fields("r10505").Value), 2, 8), CFDate(adoaccrpt105.Fields("r10505").Value))
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 1600
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10506").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10506").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 2700
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10507").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10507").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 3800
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10508").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10508").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 4700
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10509").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10509").Value
'      Else
'         Printer.Print ""
'      End If
'        'Add By Cheng 2003/12/12
'        '若要列印客戶案件案號
'        If Me.Text13.Text = "Y" Then
'            Printer.CurrentX = 6100
'            Printer.CurrentY = 3800 + intCounter * 300
'            If IsNull(adoaccrpt105.Fields("r10523").Value) = False Then
'               Printer.Print adoaccrpt105.Fields("r10523").Value
'            Else
'               Printer.Print ""
'            End If
'        End If
'        'End
'      Printer.CurrentX = 6100 + IIf(Me.Text13.Text = "Y", 1920, 0)
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10510").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10510").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 9200 + IIf(Me.Text13.Text = "Y", 1920, 0)
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10511").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10511").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 10500 + IIf(Me.Text13.Text = "Y", 1920, 0)
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10512").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10512").Value
'      Else
'         Printer.Print ""
'      End If
'      If IsNull(adoaccrpt105.Fields("r10513").Value) = False Then
'         strAmount = Format(adoaccrpt105.Fields("r10513").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
''         Printer.CurrentX = 13100 - intLength
'         Printer.CurrentX = 11600 + 250 + Printer.TextWidth("應收金額") - intLength + IIf(Me.Text13.Text = "Y", 1920, 0)
'         Printer.CurrentY = 3800 + intCounter * 300
'         Printer.Print strAmount
'      End If
'      If IsNull(adoaccrpt105.Fields("r10514").Value) = False Then
'         strAmount = Format(adoaccrpt105.Fields("r10514").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
''         Printer.CurrentX = 14200 - intLength
'         Printer.CurrentX = 13200 + 250 + Printer.TextWidth("收款金額") - intLength + IIf(Me.Text13.Text = "Y", 1920, 0)
'         Printer.CurrentY = 3800 + intCounter * 300
'         Printer.Print strAmount
'      Else
'         Printer.Print ""
'      End If
'      If IsNull(adoaccrpt105.Fields("r10515").Value) = False Then
'         strAmount = Format(adoaccrpt105.Fields("r10515").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
''         Printer.CurrentX = 15300 - intLength
'         Printer.CurrentX = 14300 + 500 + Printer.TextWidth("收款扣繳") - intLength + IIf(Me.Text13.Text = "Y", 1920, 0)
'         Printer.CurrentY = 3800 + intCounter * 300
'         Printer.Print strAmount
'      Else
'         Printer.Print ""
'      End If
'      If IsNull(adoaccrpt105.Fields("r10516").Value) = False Then
'         strAmount = Format(adoaccrpt105.Fields("r10516").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
''         Printer.CurrentX = 16400 - intLength
'         Printer.CurrentX = 15400 + 750 + Printer.TextWidth("銷帳退費") - intLength + IIf(Me.Text13.Text = "Y", 1920, 0)
'         Printer.CurrentY = 3800 + intCounter * 300
'         Printer.Print strAmount
'      Else
'         Printer.Print ""
'      End If
'      intCounter = intCounter + 1
'      adoaccrpt105.MoveNext
'   Loop
'   adoaccrpt105.Close
'   Printer.EndDoc
'End Sub
'
'*************************************************
'  列印類別-選擇未收
'
'*************************************************
Private Sub Select4()
Dim douCal1 As Double
Dim douCal2 As Double
Dim strSql As String
'Add by Morgan 2004/12/7
Dim strSQLK As String, strSQLL As String, strSQLT As String, strSQLS As String, strSQLO As String, strSqlq As String
'add by nickc 2007/02/08
Dim StrSQLa As String, strSameDocNo

    strSql = ""
   strSameName = ""
    'Add By Cheng 2003/04/28
    '初始化公司別
    m_strA0k11 = ""
   If Text1 <> MsgText(601) Then
      strSql = " and a0k03 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a0k03 <= '" & Text2 & "'"
   End If
   If Text4 <> MsgText(601) Then
      strSql = strSql & " and a0k20 = '" & Text4 & "'"
   End If

'   If Text6 = "" Then
      'Modify By Sindy 2016/6/8
      'If Text3 <> MsgText(601) Then
      If cboTitle.Text <> MsgText(601) Then
         '2011/10/20 MODIFY BY SONIA E10023515
         'strSql = strSql & " and instr(a0k04, '" & Text3 & "') > 0"
         'strSql = strSql & " and instr(UPPER(a0k04), UPPER('" & Text3 & "')) > 0"
         strSql = strSql & " and instr(UPPER(a0k04), UPPER('" & cboTitle.Text & "')) > 0"
      '2016/6/8 END
      End If
'   End If
    'Modify By Cheng 2003/04/24
'    'Add By Cheng 2003/03/12
'    '個人/公司別
'    If Me.Text12.Text <> "" Then
'        strSQL = strSQL & " and a0k05 = '" & Me.Text12.Text & "' "
'    End If
    'Add By Cheng 2004/01/09
    '公司別
    If Me.Text14.Text <> "" Then
        strSql = strSql & " and A0K11 >= '" & Me.Text14.Text & "' "
    End If
    If Me.Text15.Text <> "" Then
        strSql = strSql & " and A0K11 <= '" & Me.Text15.Text & "' "
    End If
    'End
    'Add By Cheng 2004/01/12
    '若非北所員工, 只能列印該所資料
    If pub_strUserOffice <> "1" Then
        strSql = strSql & " And ''||ST06='" & pub_strUserOffice & "' "
    End If
    'End

   'Add by Morgan 2004/11/25
   If txtSys.Text <> "" Then
      strSql = strSql & " and exists(select * from caseprogress where cp01 in (" & GetSysSQL(txtSys.Text) & ") and cp60=a0k01)"
   End If
   '2004/11/25 end

   adocaseprogress.CursorLocation = adUseClient
   adocaseprogress.Open "select cp60 from caseprogress where cp01 = '" & Text7 & "' and cp02 = '" & Text8 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text10 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adocaseprogress.RecordCount <> 0 Then
      strSql = strSql & " and a0k01 in ("
      Do While adocaseprogress.EOF = False
         If IsNull(adocaseprogress.Fields("cp60").Value) = False Then
            strSql = strSql & "'" & adocaseprogress.Fields("cp60").Value & "', "
         End If
         adocaseprogress.MoveNext
      Loop
      strSql = Mid(strSql, 1, Len(strSql) - 2) & ")"
   End If

   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      'Modify by Morgan 2004/12/7 改依單據抓作業日期
      'strSQL = strSQL & " and a0k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      strSQLK = strSql & " and a0k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      strSQLL = strSql & " and a0l02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      strSQLT = strSql & " and a0t03 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      strSQLS = strSql & " and a0s03 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      strSQLO = strSql & " and a0o05 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      strSqlq = strSql & " and a0q01 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      '2004/12/7 end
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      'Modify by Morgan 2004/12/7 改依單據抓作業日期
      'strSQL = strSQL & " and a0k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      strSQLK = strSQLK & " and a0k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      strSQLL = strSQLL & " and a0l02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      strSQLT = strSQLT & " and a0t03 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      strSQLS = strSQLS & " and a0s03 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      strSQLO = strSQLO & " and a0o05 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      strSqlq = strSqlq & " and a0q01 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      '2004/12/7 end
   End If

   adocaseprogress.Close
   adoacc0k0.CursorLocation = adUseClient
    'Modify By Cheng 2003/04/28
    '若依公司別跳頁
    If Me.Text12.Text = "Y" Then
        StrSQLa = "select a0k03 as Cust, a0k04 as Title, a0k02 as DocDate, a0k01 as DocNo, a0k01, a0k20 as Sales,  a0k23, (a0k06 + a0k07) as TAmount, 0 as RAmount, 0 as VAmount, 0 as BAmount, 1 as Serial, a0k11 from acc0k0, Staff where a0k06 <> 0 And A0K20=ST01(+) " & strSQLK & _
                       " union select a0k03 as Cust, a0k04 as Title, a0l02 as DocDate, a0m01 as DocNo, a0k01, a0k20 as Sales,  a0k23, 0 as TAmount, (a0m04 + a0m05) as RAmount, a0m06 as VAmount, 0 as BAmount, 2 as Serial, a0k11 from acc0m0, acc0l0, acc0k0, Staff where a0m01 = a0l01 and a0m02 = a0k01 And A0K20=ST01(+) " & strSQLL & _
                       " union select a0k03 as Cust, a0k04 as Title, a0t03 as DocDate, a0t01 as DocNo, a0k01, a0k20 as Sales,  a0k23, 0 as TAmount, a0t08 as RAmount, 0 as VAmount, 0 as BAmount, 3 as Serial, a0k11 from acc0m0, acc0t0, acc0k0, Staff where a0m01 = a0t07 and a0m02 = a0k01 And A0K20=ST01(+) " & strSQLT & _
                       " union select a0k03 as Cust, a0k04 as Title, a0s03 as DocDate, a0s01 as DocNo, a0k01, a0k20 as Sales,  a0k23, 0 as TAmount, 0 as RAmount, 0 as VAmount, decode(a0s05, 0, (a0s06 + a0s07), a0s05) as BAmount, 4 as Serial, a0k11 from acc0s0, acc0k0, Staff where a0s02 = a0k01 And A0K20=ST01(+) " & strSQLS & _
                       " union select a0k03 as Cust, a0k04 as Title, a0o05 as DocDate, a0o01 as DocNo, a0k01, a0k20 as Sales,  a0k23, (a0s06 + a0s07) as TAmount, 0 as RAmount, 0 as VAmount, 0 as BAmount, 5 as Serial, a0k11 from acc0o0, acc0s0, acc0k0, Staff where a0o01 = a0s10 and a0s02 = a0k01 And A0K20=ST01(+) " & strSQLO & _
                       " union select a0k03 as Cust, a0k04 as Title, a0q01 as DocDate, to_char(a0q01) as DocNo, a0k01, a0k20 as Sales,  a0k23, 0 as TAmount, a0q06 as RAmount, 0 as VAmount, 0 as BAmount, 6 as Serial, a0k11 from acc0o0, acc0q0, acc0s0, acc0k0, Staff where a0o01 = a0s10 and a0o11 = a0q01 and a0o02 = a0q03 and a0s02 = a0k01 And A0K20=ST01(+) " & strSqlq & _
                       " order by Cust asc, a0k11 asc, Title asc, a0k01 asc, Serial asc, DocDate asc"
        'End
        adoacc0k0.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
    '不依公司別跳頁
    Else
        adoacc0k0.Open "select a0k03 as Cust, a0k04 as Title, a0k02 as DocDate, a0k01 as DocNo, a0k01, a0k20 as Sales,  a0k23, (a0k06 + a0k07) as TAmount, 0 as RAmount, 0 as VAmount, 0 as BAmount, 1 as Serial, A0K11 from acc0k0, Staff where a0k06 <> 0 And A0K20=ST01(+) " & strSQLK & _
                       " union select a0k03 as Cust, a0k04 as Title, a0l02 as DocDate, a0m01 as DocNo, a0k01, a0k20 as Sales,  a0k23, 0 as TAmount, (a0m04 + a0m05) as RAmount, a0m06 as VAmount, 0 as BAmount, 2 as Serial, A0K11 from acc0m0, acc0l0, acc0k0, Staff where a0m01 = a0l01 and a0m02 = a0k01 And A0K20=ST01(+) " & strSQLL & _
                       " union select a0k03 as Cust, a0k04 as Title, a0t03 as DocDate, a0t01 as DocNo, a0k01, a0k20 as Sales,  a0k23, 0 as TAmount, a0t08 as RAmount, 0 as VAmount, 0 as BAmount, 3 as Serial, A0K11 from acc0m0, acc0t0, acc0k0, Staff where a0m01 = a0t07 and a0m02 = a0k01 And A0K20=ST01(+) " & strSQLT & _
                       " union select a0k03 as Cust, a0k04 as Title, a0s03 as DocDate, a0s01 as DocNo, a0k01, a0k20 as Sales,  a0k23, 0 as TAmount, 0 as RAmount, 0 as VAmount, decode(a0s05, 0, (a0s06 + a0s07), a0s05) as BAmount, 4 as Serial, A0K11 from acc0s0, acc0k0, Staff where a0s02 = a0k01 And A0K20=ST01(+) " & strSQLS & _
                       " union select a0k03 as Cust, a0k04 as Title, a0o05 as DocDate, a0o01 as DocNo, a0k01, a0k20 as Sales,  a0k23, (a0s06 + a0s07) as TAmount, 0 as RAmount, 0 as VAmount, 0 as BAmount, 5 as Serial, A0K11 from acc0o0, acc0s0, acc0k0, Staff where a0o01 = a0s10 and a0s02 = a0k01 And A0K20=ST01(+) " & strSQLO & _
                       " union select a0k03 as Cust, a0k04 as Title, a0q01 as DocDate, to_char(a0q01) as DocNo, a0k01, a0k20 as Sales,  a0k23, 0 as TAmount, a0q06 as RAmount, 0 as VAmount, 0 as BAmount, 6 as Serial, A0K11 from acc0o0, acc0q0, acc0s0, acc0k0, Staff where a0o01 = a0s10 and a0o11 = a0q01 and a0o02 = a0q03 and a0s02 = a0k01 And A0K20=ST01(+) " & strSqlq & _
                       " order by Cust asc, Title asc, a0k01 asc, Serial asc, DocDate asc", adoTaie, adOpenStatic, adLockReadOnly
    End If
   If adoacc0k0.RecordCount = 0 Then
      adoacc0k0.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc0k0.EOF = False
      If adoacc0k0.Fields("a0k01").Value <> strSameDocNo Then
         strSameDocNo = adoacc0k0.Fields("a0k01").Value
      End If
      If adoacc0k0.Fields("Title").Value <> strSameName Then
         strSameName = adoacc0k0.Fields("Title").Value
      End If
      If "" & adoacc0k0.Fields("a0k11").Value <> m_strA0k11 Then
         m_strA0k11 = "" & adoacc0k0.Fields("a0k11").Value
      End If
      adoaccrpt105.AddNew
      adoaccrpt105.Fields("r10501").Value = strUserNum
      adoaccrpt105.Fields("r10502").Value = Counter
      If Text6 <> "" Then
         adoaccrpt105.Fields("r10503").Value = Text6
         'Modify By Sindy 2016/6/8
         'adoaccrpt105.Fields("r10504").Value = Text3
         adoaccrpt105.Fields("r10504").Value = cboTitle.Text
         '2016/6/8 END
      Else
         If IsNull(adoacc0k0.Fields("Cust").Value) Then
            adoaccrpt105.Fields("r10503").Value = Null
         Else
            adoaccrpt105.Fields("r10503").Value = adoacc0k0.Fields("Cust").Value
         End If
         If IsNull(adoacc0k0.Fields("Title").Value) Then
            adoaccrpt105.Fields("r10504").Value = Null
         Else
            adoaccrpt105.Fields("r10504").Value = adoacc0k0.Fields("Title").Value
         End If
      End If
      If IsNull(adoacc0k0.Fields("DocDate").Value) Then
         adoaccrpt105.Fields("r10505").Value = Null
      Else
         adoaccrpt105.Fields("r10505").Value = adoacc0k0.Fields("DocDate").Value
      End If
      adoaccrpt105.Fields("r10506").Value = adoacc0k0.Fields("DocNo").Value
      adoaccrpt105.Fields("r10507").Value = adoacc0k0.Fields("a0k01").Value
      If IsNull(adoacc0k0.Fields("Sales").Value) Then
         adoaccrpt105.Fields("r10508").Value = Null
      Else
         adoaccrpt105.Fields("r10508").Value = StaffQuery(adoacc0k0.Fields("Sales").Value)
      End If
      adocaseprogress.CursorLocation = adUseClient
      'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
      If Text7 <> MsgText(601) Then
         adocaseprogress.Open "select cp01||'-'||cp02||'-'||cp03||'-'||cp04, cp09, cp01, getcp10desc(cp01,cp10,a0j04) cp10N, na03 from caseprogress, acc0j0,nation where cp09 = a0j01 (+) and cp60 = '" & adoacc0k0.Fields("a0k01").Value & "' and cp01 = '" & Text7 & "' and cp02 = '" & Text8 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text10 & "' and na01(+)=a0j04", adoTaie, adOpenStatic, adLockReadOnly
      Else
         adocaseprogress.Open "select cp01||'-'||cp02||'-'||cp03||'-'||cp04, cp09, cp01, getcp10desc(cp01,cp10,a0j04) cp10N, na03 from caseprogress, acc0j0,nation where cp09 = a0j01 (+) and cp60 = '" & adoacc0k0.Fields("a0k01").Value & "' and na01(+)=a0j04", adoTaie, adOpenStatic, adLockReadOnly
      End If
      If adocaseprogress.RecordCount <> 0 Then
         adoaccrpt105.Fields("r10509").Value = Replace(adocaseprogress.Fields(0).Value, "-0-00", "")
         adoaccrpt105.Fields("r10510").Value = MidB(CaseNameQuery(adocaseprogress.Fields(1).Value, 1), 1, 24)
         If adoaccrpt105.Fields("r10510").Value = "" Then
            adoaccrpt105.Fields("r10510").Value = Mid(CaseNameQuery(adocaseprogress.Fields(1).Value, 2), 1, 24)
         End If

         'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
         If IsNull(adocaseprogress.Fields("cp10N").Value) Then
            adoaccrpt105.Fields("r10511").Value = Null
         Else
            adoaccrpt105.Fields("r10511").Value = adocaseprogress.Fields("cp10N").Value
         End If
         If IsNull(adocaseprogress.Fields("na03").Value) Then
            adoaccrpt105.Fields("r10512").Value = Null
         Else
            adoaccrpt105.Fields("r10512").Value = adocaseprogress.Fields("na03").Value
         End If
      Else
         adoaccrpt105.Fields("r10509").Value = Null
         adoaccrpt105.Fields("r10510").Value = Null
         adoaccrpt105.Fields("r10511").Value = Null
         adoaccrpt105.Fields("r10512").Value = Null
      End If
      adocaseprogress.Close
      If IsNull(adoacc0k0.Fields("TAmount").Value) Then
         adoaccrpt105.Fields("r10513").Value = Null
      Else
         adoaccrpt105.Fields("r10513").Value = Format(adoacc0k0.Fields("TAmount").Value)
      End If
      If IsNull(adoacc0k0.Fields("RAmount").Value) Then
         adoaccrpt105.Fields("r10514").Value = Null
      Else
         adoaccrpt105.Fields("r10514").Value = Format(adoacc0k0.Fields("RAmount").Value)
      End If
      If IsNull(adoacc0k0.Fields("VAmount").Value) Then
         adoaccrpt105.Fields("r10515").Value = Null
      Else
         adoaccrpt105.Fields("r10515").Value = Format(adoacc0k0.Fields("VAmount").Value)
      End If
      If IsNull(adoacc0k0.Fields("BAmount").Value) Then
         adoaccrpt105.Fields("r10516").Value = Null
      Else
         adoaccrpt105.Fields("r10516").Value = Format(adoacc0k0.Fields("BAmount").Value)
      End If
      adoaccrpt105.Fields("r10522").Value = m_strA0k11
        '客戶案件案號
        '2013/5/13 modify by sonia 不列印客戶案件案號改印本所案號
        'If Me.Text13.Text = "Y" Then
        '    adoaccrpt105.Fields("r10523").Value = GetCustCaseNo("" & adoacc0k0("a0k01").Value)
        'Else
        '    adoaccrpt105.Fields("r10523").Value = ""
        'End If
        adoaccrpt105.Fields("r10523").Value = GetCustCaseNo("" & adoacc0k0("a0k01").Value, Me.Text13.Text)
        '2013/5/13 end
      adoaccrpt105.UpdateBatch
      adoacc0k0.MoveNext
'-------------------------------------------------
' 小計計算
'-------------------------------------------------
      If adoacc0k0.EOF = False Then
         If Text6 = "" Then
            If adoacc0k0.Fields("Title").Value <> strSameName Then
               SubSelect4
            ElseIf "" & adoacc0k0.Fields("a0k11").Value <> m_strA0k11 Then
               If Me.Text12.Text = "Y" Then
                    SubSelect4
               End If
            End If
         Else
            If "" & adoacc0k0.Fields("a0k11").Value <> m_strA0k11 Then
               If Me.Text12.Text = "Y" Then
                    SubSelect4
               End If
            End If
         End If
      Else
'-------------------------------------------------
'  合計計算
'-------------------------------------------------
         SubSelect4
         adoacc0k0.MoveLast
         adoaccsum.CursorLocation = adUseClient
         adoaccsum.Open "select sum(to_number(nvl(r10513, 0))), sum(to_number(nvl(r10514, 0))), sum(to_number(nvl(r10515, 0))), sum(to_number(nvl(r10516, 0))) from accrpt105 where r10501 = '" & strUserNum & "' and r10509 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If adoaccsum.RecordCount <> 0 And (adoaccsum.Fields(0).Value <> 0 Or adoaccsum.Fields(1).Value <> 0 Or adoaccsum.Fields(2).Value <> 0 Or adoaccsum.Fields(3).Value <> 0) Then
             adoaccrpt105.AddNew
             adoaccrpt105.Fields("r10501").Value = strUserNum
             adoaccrpt105.Fields("r10502").Value = Counter
             If Text6 <> "" Then
                adoaccrpt105.Fields("r10503").Value = Text6
                'Modify By Sindy 2016/6/8
                'adoaccrpt105.Fields("r10504").Value = Text3
                adoaccrpt105.Fields("r10504").Value = cboTitle.Text
                '2016/6/8 END
             Else
                adoaccrpt105.Fields("r10503").Value = adoacc0k0.Fields("Cust").Value
                adoaccrpt105.Fields("r10504").Value = adoacc0k0.Fields("Title").Value
             End If
             adoaccrpt105.Fields("r10522").Value = m_strA0k11
             PaintLine ReportSum(4)
             adoaccrpt105.UpdateBatch
             adoaccrpt105.AddNew
             adoaccrpt105.Fields("r10501").Value = strUserNum
             adoaccrpt105.Fields("r10502").Value = Counter
             If Text6 <> "" Then
                adoaccrpt105.Fields("r10503").Value = Text6
                'Modify By Sindy 2016/6/8
                'adoaccrpt105.Fields("r10504").Value = Text3
                adoaccrpt105.Fields("r10504").Value = cboTitle.Text
                '2016/6/8 END
             Else
                adoaccrpt105.Fields("r10503").Value = adoacc0k0.Fields("Cust").Value
                adoaccrpt105.Fields("r10504").Value = adoacc0k0.Fields("Title").Value
             End If
             adoaccrpt105.Fields("r10510").Value = ReportSum(25)
            If IsNull(adoaccsum.Fields(0).Value) Then
               adoaccrpt105.Fields("r10513").Value = 0
            Else
               adoaccrpt105.Fields("r10513").Value = Format(adoaccsum.Fields(0).Value)
            End If
            If IsNull(adoaccsum.Fields(1).Value) Then
               adoaccrpt105.Fields("r10514").Value = 0
            Else
               adoaccrpt105.Fields("r10514").Value = Format(adoaccsum.Fields(1).Value)
            End If
            If IsNull(adoaccsum.Fields(2).Value) Then
               adoaccrpt105.Fields("r10515").Value = 0
            Else
               adoaccrpt105.Fields("r10515").Value = Format(adoaccsum.Fields(2).Value)
            End If
            If IsNull(adoaccsum.Fields(3).Value) Then
               adoaccrpt105.Fields("r10516").Value = 0
            Else
               adoaccrpt105.Fields("r10516").Value = Format(adoaccsum.Fields(3).Value)
            End If
            'Add By Cheng 2003/04/28
             adoaccrpt105.Fields("r10522").Value = m_strA0k11
             adoaccrpt105.UpdateBatch
             adoaccrpt105.AddNew
             adoaccrpt105.Fields("r10501").Value = strUserNum
             adoaccrpt105.Fields("r10502").Value = Counter
             If Text6 <> "" Then
                adoaccrpt105.Fields("r10503").Value = Text6
                'Modify By Sindy 2016/6/8
                'adoaccrpt105.Fields("r10504").Value = Text3
                adoaccrpt105.Fields("r10504").Value = cboTitle.Text
                '2016/6/8 END
             Else
                adoaccrpt105.Fields("r10503").Value = adoacc0k0.Fields("Cust").Value
                adoaccrpt105.Fields("r10504").Value = adoacc0k0.Fields("Title").Value
             End If
            'Add By Cheng 2003/04/28
             adoaccrpt105.Fields("r10522").Value = m_strA0k11
             PaintLine ReportSum(8)
             adoaccrpt105.UpdateBatch
         End If
         adoaccsum.Close
         adoacc0k0.MoveNext
      End If
   Loop
   adoacc0k0.Close
End Sub

'*************************************************
'  列印類別-選擇交易情形之小計計算
'
'*************************************************
Private Sub SubSelect4()
Dim strSql As String
  
   adoacc0k0.MovePrevious
   adoaccsum.CursorLocation = adUseClient
   If Text6 = "" Then
      'Modify By Cheng 2003/04/28
      '依公司別跳頁
      If Me.Text12.Text = "Y" Then
        adoaccsum.Open "select sum(to_number(nvl(r10513, 0))), sum(to_number(nvl(r10514, 0))), sum(to_number(nvl(r10515, 0))), sum(to_number(nvl(r10516, 0))) from accrpt105 where r10501 = '" & strUserNum & "' and r10507 is not null and r10504 = '" & adoacc0k0.Fields("Title").Value & "' and r10522='" & m_strA0k11 & "' ", adoTaie, adOpenStatic, adLockReadOnly
      '不依公司別跳頁
      Else
        adoaccsum.Open "select sum(to_number(nvl(r10513, 0))), sum(to_number(nvl(r10514, 0))), sum(to_number(nvl(r10515, 0))), sum(to_number(nvl(r10516, 0))) from accrpt105 where r10501 = '" & strUserNum & "' and r10507 is not null and r10504 = '" & adoacc0k0.Fields("Title").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      End If
   Else
      'Modify By Cheng 2003/04/28
      '依公司別跳頁
      If Me.Text12.Text = "Y" Then
        'Modify By Sindy 2016/6/8 Text3 ==> cboTitle.Text
        adoaccsum.Open "select sum(to_number(nvl(r10513, 0))), sum(to_number(nvl(r10514, 0))), sum(to_number(nvl(r10515, 0))), sum(to_number(nvl(r10516, 0))) from accrpt105 where r10501 = '" & strUserNum & "' and r10507 is not null and r10504 = '" & cboTitle.Text & "' and r10522='" & m_strA0k11 & "' ", adoTaie, adOpenStatic, adLockReadOnly
      '不依公司別跳頁
      Else
        'Modify By Sindy 2016/6/8 Text3 ==> cboTitle.Text
        adoaccsum.Open "select sum(to_number(nvl(r10513, 0))), sum(to_number(nvl(r10514, 0))), sum(to_number(nvl(r10515, 0))), sum(to_number(nvl(r10516, 0))) from accrpt105 where r10501 = '" & strUserNum & "' and r10507 is not null and r10504 = '" & cboTitle.Text & "'", adoTaie, adOpenStatic, adLockReadOnly
      End If
   End If
   If adoaccsum.RecordCount <> 0 And (adoaccsum.Fields(0).Value <> 0 Or adoaccsum.Fields(1).Value <> 0 Or adoaccsum.Fields(2).Value <> 0 Or adoaccsum.Fields(3).Value <> 0) Then
      adoaccrpt105.AddNew
      adoaccrpt105.Fields("r10501").Value = strUserNum
      adoaccrpt105.Fields("r10502").Value = Counter
      If Text6 <> "" Then
         adoaccrpt105.Fields("r10503").Value = Text6
         'Modify By Sindy 2016/6/8
         'adoaccrpt105.Fields("r10504").Value = Text3
         adoaccrpt105.Fields("r10504").Value = cboTitle.Text
         '2016/6/8 END
      Else
         adoaccrpt105.Fields("r10503").Value = adoacc0k0.Fields("Cust").Value
         adoaccrpt105.Fields("r10504").Value = adoacc0k0.Fields("Title").Value
      End If
       'Add By Cheng 2003/04/28
       adoaccrpt105.Fields("r10522").Value = m_strA0k11
      PaintLine ReportSum(4)
      adoaccrpt105.UpdateBatch
      adoaccrpt105.AddNew
      adoaccrpt105.Fields("r10501").Value = strUserNum
      adoaccrpt105.Fields("r10502").Value = Counter
      If Text6 <> "" Then
         adoaccrpt105.Fields("r10503").Value = Text6
         'Modify By Sindy 2016/6/8
         'adoaccrpt105.Fields("r10504").Value = Text3
         adoaccrpt105.Fields("r10504").Value = cboTitle.Text
         '2016/6/8 END
      Else
         adoaccrpt105.Fields("r10503").Value = adoacc0k0.Fields("Cust").Value
         adoaccrpt105.Fields("r10504").Value = adoacc0k0.Fields("Title").Value
      End If
      adoaccrpt105.Fields("r10510").Value = ReportSum(24)
      If IsNull(adoaccsum.Fields(0).Value) Then
         adoaccrpt105.Fields("r10513").Value = 0
      Else
         adoaccrpt105.Fields("r10513").Value = Format(adoaccsum.Fields(0).Value)
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         adoaccrpt105.Fields("r10514").Value = 0
      Else
         adoaccrpt105.Fields("r10514").Value = Format(adoaccsum.Fields(1).Value)
      End If
      If IsNull(adoaccsum.Fields(2).Value) Then
         adoaccrpt105.Fields("r10515").Value = 0
      Else
         adoaccrpt105.Fields("r10515").Value = Format(adoaccsum.Fields(2).Value)
      End If
      If IsNull(adoaccsum.Fields(3).Value) Then
         adoaccrpt105.Fields("r10516").Value = 0
      Else
         adoaccrpt105.Fields("r10516").Value = Format(adoaccsum.Fields(3).Value)
      End If
    'Add By Cheng 2003/04/28
      adoaccrpt105.Fields("r10522").Value = m_strA0k11
      adoaccrpt105.UpdateBatch
   End If
   adoaccsum.Close
   adoacc0k0.MoveNext
End Sub

'2013/5/13 cancel by sonia 沒有用到
''*************************************************
''  抬頭列印
''
''*************************************************
'Private Sub PrintHead4()
'   Printer.FontSize = 16
'   Printer.CurrentX = 5000
'   Printer.CurrentY = 1000
'   Printer.Print ReportTitle(105)
'   Printer.CurrentX = 9000
'   Printer.CurrentY = 1000
'    'Modify By Cheng 2003/03/12
'    '加個人或公司別
'   Printer.Print "(交易情形)"
''   Printer.Print "(交易情形)" & IIf(Me.Text12.Text = "1", " (個人)", IIf(Me.Text12.Text = "2", " (公司)", ""))
'   Printer.FontSize = 12
'   Printer.CurrentX = 5500
'   Printer.CurrentY = 1500
'   Printer.Print "帳款日期: "
'   Printer.CurrentX = 6600
'   Printer.CurrentY = 1500
'   Printer.Print MaskEdBox1.Text & " ~ " & MaskEdBox2.Text
'   Printer.CurrentX = 500
'   Printer.CurrentY = 2100
'   Printer.Print "列印人員: "
'   Printer.CurrentX = 1800
'   Printer.CurrentY = 2100
'   Printer.Print StaffQuery(strUserNum)
'   Printer.CurrentX = 12000
'   Printer.CurrentY = 2100
'   Printer.Print "列印日期: "
'   Printer.CurrentX = 13300
'   Printer.CurrentY = 2100
'   Printer.Print IIf(Mid(CFDate(ACDate(ServerDate)), 1, 1) = "0", Mid(CFDate(ACDate(ServerDate)), 2, 8), CFDate(ACDate(ServerDate)))
'    'Add By Cheng 2003/04/28
'    '依公司別跳頁
'    If Me.Text12.Text = "Y" Then
'        Printer.CurrentX = 500
'        Printer.CurrentY = 2400
'        Printer.Print "公  司  別:　" & GetCompanyName("" & adoaccrpt105.Fields("r10522").Value)
'    End If
'   Printer.CurrentX = 12000
'   Printer.CurrentY = 2400
'   Printer.Print "頁次: "
'   Printer.CurrentX = 13300
'   Printer.CurrentY = 2400
'   Printer.Print intPage
'   Printer.CurrentX = 500
'   Printer.CurrentY = 2700
'   Printer.Print "客戶代號: "
'   Printer.CurrentX = 1800
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt105.Fields("r10503").Value) = False Then
'      Printer.Print adoaccrpt105.Fields("r10503").Value
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 3000
'   Printer.CurrentY = 2700
'   Printer.Print "客戶名稱: "
'   Printer.CurrentX = 4100
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt105.Fields("cu04").Value) = False Then
'      Printer.Print adoaccrpt105.Fields("cu04").Value
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 8000
'   Printer.CurrentY = 2700
'   Printer.Print "收據抬頭: "
'   Printer.CurrentX = 9100
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt105.Fields("r10504").Value) = False Then
'      Printer.Print adoaccrpt105.Fields("r10504").Value
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 500
'   Printer.CurrentY = 3300
'   Printer.Print "單據日期"
'   Printer.CurrentX = 1600
'   Printer.CurrentY = 3300
'   Printer.Print "單據號碼"
'   Printer.CurrentX = 2900
'   Printer.CurrentY = 3300
'   Printer.Print "收據號碼"
'   Printer.CurrentX = 4100
'   Printer.CurrentY = 3300
'   Printer.Print "智權人員"
'   Printer.CurrentX = 5000
'   Printer.CurrentY = 3300
'   Printer.Print "本所案號"
'    'Add By Cheng 2003/12/12
'    If Me.Text13.Text = "Y" Then
'        Printer.CurrentX = 6700
'        Printer.CurrentY = 3300
'        Printer.Print "客戶案件案號"
'    End If
'    'End
'   Printer.CurrentX = 6700 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "案件名稱"
'   Printer.CurrentX = 9800 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "案件性質"
'   Printer.CurrentX = 11100 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "申請國家"
'   Printer.CurrentX = 12500 + 200 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "應收金額"
'   Printer.CurrentX = 12500 + 200 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3600
'   Printer.Print "應付金額"
'   Printer.CurrentX = 13800 + 200 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "收款金額"
'   Printer.CurrentX = 13800 + 200 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3600
'   Printer.Print "付款金額"
'   Printer.CurrentX = 14900 + 200 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "收款扣繳"
'   Printer.CurrentX = 16000 + 200 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "銷帳退費"
'   Printer.Line (500, 4000)-(17100 + 200 + IIf(Me.Text13.Text = "Y", 1920, 0), 4000)
'End Sub

'2013/5/13 cancel by sonia 沒有用到
''*************************************************
'' 列印報表 (交易情形)
''
''*************************************************
'Public Sub FormPrint4()
'   intCounter = 0
'   intPage = 0
'   strCustomerNo = ""
'    'Add By Cheng 2003/04/28
'    '初始化公司別變數
'    m_strCompany = ""
'   adoaccrpt105.CursorLocation = adUseClient
'   adoaccrpt105.Open "select * from accrpt105, customer where substr(r10503, 1, 8) = cu01 (+) and substr(r10503, 9, 1) = cu02 (+) and r10501 = '" & strUserNum & "' order by r10501 asc, r10502 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Do While adoaccrpt105.EOF = False
'      If strCustomerNo <> adoaccrpt105.Fields("r10504").Value Then
'         intCounter = 0
'         intPage = intPage + 1
'         If strCustomerNo <> "" Then
'            Printer.NewPage
'         End If
'         PrintHead4
'         strCustomerNo = adoaccrpt105.Fields("r10504").Value
'            m_strCompany = "" & adoaccrpt105.Fields("r10522").Value
'        'Add By Cheng 2003/04/24
'        '若公司別不同
'        ElseIf m_strCompany <> "" & adoaccrpt105.Fields("r10522").Value Then
'            If Me.Text12.Text = "Y" Then
'                intCounter = 0
'                intPage = intPage + 1
'                Printer.NewPage
'                PrintHead4
'                m_strCompany = "" & adoaccrpt105.Fields("r10522").Value
'            End If
'      End If
'      If intCounter = 35 Then
'         intCounter = 0
'         intPage = intPage + 1
'         Printer.NewPage
'         PrintHead4
'      End If
'      Printer.CurrentX = 500
'      Printer.CurrentY = 4100 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10505").Value) = False Then
'         Printer.Print IIf(Mid(CFDate(adoaccrpt105.Fields("r10505").Value), 1, 1) = "0", Mid(CFDate(adoaccrpt105.Fields("r10505").Value), 2, 8), CFDate(adoaccrpt105.Fields("r10505").Value))
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 1600
'      Printer.CurrentY = 4100 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10506").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10506").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 2900
'      Printer.CurrentY = 4100 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10507").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10507").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 4100
'      Printer.CurrentY = 4100 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10508").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10508").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 5000
'      Printer.CurrentY = 4100 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10509").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10509").Value
'      Else
'         Printer.Print ""
'      End If
'        'Add By Cheng 2003/12/12
'        '若要列印客戶案件案號
'        If Me.Text13.Text = "Y" Then
'            Printer.CurrentX = 6700
'            Printer.CurrentY = 4100 + intCounter * 300
'            If IsNull(adoaccrpt105.Fields("r10523").Value) = False Then
'               Printer.Print adoaccrpt105.Fields("r10523").Value
'            Else
'               Printer.Print ""
'            End If
'        End If
'        'End
'      Printer.CurrentX = 6700 + IIf(Me.Text13.Text = "Y", 1920, 0)
'      Printer.CurrentY = 4100 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10510").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10510").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 9800 + IIf(Me.Text13.Text = "Y", 1920, 0)
'      Printer.CurrentY = 4100 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10511").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10511").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 11100 + IIf(Me.Text13.Text = "Y", 1920, 0)
'      Printer.CurrentY = 4100 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10512").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10512").Value
'      Else
'         Printer.Print ""
'      End If
'      If IsNull(adoaccrpt105.Fields("r10513").Value) = False Then
'         strAmount = Format(adoaccrpt105.Fields("r10513").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
''         Printer.CurrentX = 13800 - intLength
'         Printer.CurrentX = 12500 + 200 + Printer.TextWidth("應收金額") - intLength + IIf(Me.Text13.Text = "Y", 1920, 0)
'         Printer.CurrentY = 4100 + intCounter * 300
'         Printer.Print strAmount
'      End If
'      If IsNull(adoaccrpt105.Fields("r10514").Value) = False Then
'         strAmount = Format(adoaccrpt105.Fields("r10514").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
''         Printer.CurrentX = 14800 - intLength
'         Printer.CurrentX = 13800 + 200 + Printer.TextWidth("收款金額") - intLength + IIf(Me.Text13.Text = "Y", 1920, 0)
'         Printer.CurrentY = 4100 + intCounter * 300
'         Printer.Print strAmount
'      Else
'         Printer.Print ""
'      End If
'      If IsNull(adoaccrpt105.Fields("r10515").Value) = False Then
'         strAmount = Format(adoaccrpt105.Fields("r10515").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
''         Printer.CurrentX = 15900 - intLength
'         Printer.CurrentX = 14900 + 200 + Printer.TextWidth("收款扣繳") - intLength + IIf(Me.Text13.Text = "Y", 1920, 0)
'         Printer.CurrentY = 4100 + intCounter * 300
'         Printer.Print strAmount
'      Else
'         Printer.Print ""
'      End If
'      If IsNull(adoaccrpt105.Fields("r10516").Value) = False Then
'         strAmount = Format(adoaccrpt105.Fields("r10516").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
''         Printer.CurrentX = 17000 - intLength
'         Printer.CurrentX = 16000 + 200 + Printer.TextWidth("銷帳退費") - intLength + IIf(Me.Text13.Text = "Y", 1920, 0)
'         Printer.CurrentY = 4100 + intCounter * 300
'         Printer.Print strAmount
'      Else
'         Printer.Print ""
'      End If
'      intCounter = intCounter + 1
'      adoaccrpt105.MoveNext
'   Loop
'   adoaccrpt105.Close
'   Printer.EndDoc
'End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub

Private Sub Text8_Validate(Cancel As Boolean)
   If Text8 = MsgText(601) Then
      Exit Sub
   End If
   Text9 = "0"
   Text10 = "00"
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   'Add By Cheng 2004/01/09
   '檢查公司別
   If Me.Text14.Text = "" Then Me.Text14.Text = "1"
   If Me.Text15.Text = "" Then Me.Text15.Text = "8"
   If Me.Text14.Text > Me.Text15.Text Then
'        MsgBox "公司別範圍輸入錯誤!!!", vbExclamation + vbOKOnly
       Exit Function
   End If
   'End
   If Text1 <> MsgText(601) And Text1 <> "X" Then
      FormCheck = True
      Exit Function
   End If
   If Text2 <> MsgText(601) And Text2 <> "X" Then
      FormCheck = True
      Exit Function
   End If
   'Modify By Sindy 2016/6/8 Text3 ==> cboTitle.Text
   'If Text3 <> MsgText(601) Then
   If cboTitle.Text <> MsgText(601) Then
   '2016/6/8 END
      FormCheck = True
      Exit Function
   End If
   If Text4 <> MsgText(601) Then
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
   If Text7 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text8 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text9 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text10 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
'   If Text5 <> MsgText(601) Then
'      FormCheck = True
'      Exit Function
'   End If
   FormCheck = False
End Function

'2013/5/13 cancel by sonia 沒有用到
''*************************************************
''  抬頭列印--函證用未收
''
''*************************************************
'Private Sub PrintHead5()
'   Printer.FontSize = 16
'   Printer.CurrentX = 5000
'   Printer.CurrentY = 1000
'   Printer.Print ReportTitle(105)
'   Printer.CurrentX = 9000
'   Printer.CurrentY = 1000
''    'Modify By Cheng 2003/03/12
''    '加個人或公司別
'   Printer.Print "(函證用未收)"
''   Printer.Print "(函證用未收)" & IIf(Me.Text12.Text = "1", " (個人)", IIf(Me.Text12.Text = "2", " (公司)", ""))
'   Printer.FontSize = 10
'   Printer.CurrentX = 5500
'   Printer.CurrentY = 1500
'   Printer.Print "帳款日期: "
'   Printer.CurrentX = 6600
'   Printer.CurrentY = 1500
'   Printer.Print MaskEdBox1.Text & " ~ " & MaskEdBox2.Text
'   Printer.CurrentX = 100
'   Printer.CurrentY = 2100
'   Printer.Print "列印人員: "
'   Printer.CurrentX = 1400
'   Printer.CurrentY = 2100
'   Printer.Print StaffQuery(strUserNum)
'   Printer.CurrentX = 12000
'   Printer.CurrentY = 2100
'   Printer.Print "列印日期: "
'   Printer.CurrentX = 13300
'   Printer.CurrentY = 2100
'   Printer.Print IIf(Mid(CFDate(ACDate(ServerDate)), 1, 1) = "0", Mid(CFDate(ACDate(ServerDate)), 2, 8), CFDate(ACDate(ServerDate)))
'   Printer.CurrentX = 12000
'   Printer.CurrentY = 2400
'   Printer.Print "頁次: "
'   Printer.CurrentX = 13300
'   Printer.CurrentY = 2400
'   Printer.Print intPage
'   Printer.CurrentX = 100
'   Printer.CurrentY = 2700
'   Printer.Print "客戶代號: "
'   Printer.CurrentX = 1400
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt105.Fields("r10503").Value) = False Then
'      Printer.Print adoaccrpt105.Fields("r10503").Value
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 3000
'   Printer.CurrentY = 2700
'   Printer.Print "客戶名稱: "
'   Printer.CurrentX = 4100
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt105.Fields("cu04").Value) = False Then
'      Printer.Print adoaccrpt105.Fields("cu04").Value
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 8000
'   Printer.CurrentY = 2700
'   Printer.Print "收據抬頭: "
'   Printer.CurrentX = 9100
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt105.Fields("r10504").Value) = False Then
'      Printer.Print adoaccrpt105.Fields("r10504").Value
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 100
'   Printer.CurrentY = 3300
'   Printer.Print "收據日期"
'   Printer.CurrentX = 1200
'   Printer.CurrentY = 3300
'   Printer.Print "收據號碼"
'   Printer.CurrentX = 2300
'   Printer.CurrentY = 3300
'   Printer.Print "智權人員"
'   Printer.CurrentX = 3000
'   Printer.CurrentY = 3300
'   Printer.Print "本所案號"
'    'Add By Cheng 2003/12/12
'    If Me.Text13.Text = "Y" Then
'        Printer.CurrentX = 4400
'        Printer.CurrentY = 3300
'        Printer.Print "客戶案件案號"
'    End If
'    'End
'   Printer.CurrentX = 4400 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "案件名稱"
'   Printer.CurrentX = 7200 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "案件性質"
'   Printer.CurrentX = 8300 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "申請國家"
'   Printer.CurrentX = 9400 + 200 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "服務費"
'   Printer.CurrentX = 10700 + 200 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "規費"
'   Printer.CurrentX = 11900 + 200 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "應收金額"
'   Printer.CurrentX = 13300 + 200 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "已收金額"
'   Printer.CurrentX = 14400 + 200 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "未收金額"
'   Printer.CurrentX = 15600 + 200 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "票據金額"
'   Printer.CurrentX = 16800 + 200 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "票據到期日"
'   Printer.CurrentX = 17900 + 200 + IIf(Me.Text13.Text = "Y", 1920, 0)
'   Printer.CurrentY = 3300
'   Printer.Print "收款單號"
'   Printer.Line (100, 3700)-(18800 + 200 + IIf(Me.Text13.Text = "Y", 1920, 0), 3700)
'End Sub

'2013/5/13 cancel by sonia 沒有用到
''*************************************************
'' 列印報表 (函證用未收)
''
''*************************************************
'Public Sub FormPrint5()
'   intCounter = 0
'   intPage = 0
'   strCustomerNo = ""
'    'Add By Cheng 2003/04/28
'    '初始化公司別變數
'    m_strCompany = ""
'   adoaccrpt105.CursorLocation = adUseClient
'   adoaccrpt105.Open "select * from accrpt105, customer where substr(r10503, 1, 8) = cu01 (+) and substr(r10503, 9, 1) = cu02 (+) and r10501 = '" & strUserNum & "' order by r10501 asc, r10502 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Do While adoaccrpt105.EOF = False
'      If strCustomerNo <> adoaccrpt105.Fields("r10503").Value Then
'         intCounter = 0
'         intPage = intPage + 1
'         If strCustomerNo <> "" Then
'            Printer.NewPage
'         End If
'         PrintHead5
'         strCustomerNo = adoaccrpt105.Fields("r10503").Value
'            m_strCompany = "" & adoaccrpt105.Fields("r10522").Value
'        'Add By Cheng 2003/04/24
'        '若公司別不同
'        ElseIf m_strCompany <> "" & adoaccrpt105.Fields("r10522").Value Then
'            If Me.Text12.Text = "Y" Then
'                intCounter = 0
'                intPage = intPage + 1
'                Printer.NewPage
'                PrintHead5
'                m_strCompany = "" & adoaccrpt105.Fields("r10522").Value
'            End If
'      End If
'      If intCounter = 35 Then
'         intCounter = 0
'         intPage = intPage + 1
'         Printer.NewPage
'         PrintHead5
'      End If
'      Printer.CurrentX = 100
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10505").Value) = False Then
'         Printer.Print IIf(Mid(CFDate(adoaccrpt105.Fields("r10505").Value), 1, 1) = "0", Mid(CFDate(adoaccrpt105.Fields("r10505").Value), 2, 8), CFDate(adoaccrpt105.Fields("r10505").Value))
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 1200
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10507").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10507").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 2300
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10508").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10508").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 3000
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10509").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10509").Value
'      Else
'         Printer.Print ""
'      End If
'        'Add By Cheng 2003/12/12
'        '若要列印客戶案件案號
'        If Me.Text13.Text = "Y" Then
'            Printer.CurrentX = 4400
'            Printer.CurrentY = 3800 + intCounter * 300
'            If IsNull(adoaccrpt105.Fields("r10523").Value) = False Then
'               Printer.Print adoaccrpt105.Fields("r10523").Value
'            Else
'               Printer.Print ""
'            End If
'        End If
'        'End
'      Printer.CurrentX = 4400 + IIf(Me.Text13.Text = "Y", 1920, 0)
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10510").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10510").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 7200 + IIf(Me.Text13.Text = "Y", 1920, 0)
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10511").Value) = False Then
'         Printer.Print MidB(adoaccrpt105.Fields("r10511").Value, 1, 10)
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 8300 + IIf(Me.Text13.Text = "Y", 1920, 0)
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10512").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10512").Value
'      Else
'         Printer.Print ""
'      End If
'      If IsNull(adoaccrpt105.Fields("r10513").Value) = False Then
'         strAmount = Format(adoaccrpt105.Fields("r10513").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
''         Printer.CurrentX = 10500 - intLength
'         Printer.CurrentX = 9400 + 200 + Printer.TextWidth("服務費") - intLength + IIf(Me.Text13.Text = "Y", 1920, 0)
'         Printer.CurrentY = 3800 + intCounter * 300
'         Printer.Print strAmount
'      End If
'      If IsNull(adoaccrpt105.Fields("r10514").Value) = False Then
'         strAmount = Format(adoaccrpt105.Fields("r10514").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
''         Printer.CurrentX = 11700 - intLength
'         Printer.CurrentX = 10700 + 200 + Printer.TextWidth("規費") - intLength + IIf(Me.Text13.Text = "Y", 1920, 0)
'         Printer.CurrentY = 3800 + intCounter * 300
'         Printer.Print strAmount
'      Else
'         Printer.Print ""
'      End If
'      If IsNull(adoaccrpt105.Fields("r10515").Value) = False Then
'         strAmount = Format(adoaccrpt105.Fields("r10515").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
''         Printer.CurrentX = 13100 - intLength
'         Printer.CurrentX = 11900 + 200 + Printer.TextWidth("應收金額") - intLength + IIf(Me.Text13.Text = "Y", 1920, 0)
'         Printer.CurrentY = 3800 + intCounter * 300
'         Printer.Print strAmount
'      Else
'         Printer.Print ""
'      End If
'      If IsNull(adoaccrpt105.Fields("r10516").Value) = False Then
'         strAmount = Format(adoaccrpt105.Fields("r10516").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
''         Printer.CurrentX = 14200 - intLength
'         Printer.CurrentX = 13300 + 200 + Printer.TextWidth("已收金額") - intLength + IIf(Me.Text13.Text = "Y", 1920, 0)
'         Printer.CurrentY = 3800 + intCounter * 300
'         Printer.Print strAmount
'      Else
'         Printer.Print ""
'      End If
'      If IsNull(adoaccrpt105.Fields("r10517").Value) = False Then
'         strAmount = Format(adoaccrpt105.Fields("r10517").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
''         Printer.CurrentX = 15400 - intLength
'         Printer.CurrentX = 14400 + 200 + Printer.TextWidth("未收金額") - intLength + IIf(Me.Text13.Text = "Y", 1920, 0)
'         Printer.CurrentY = 3800 + intCounter * 300
'         Printer.Print strAmount
'      Else
'         Printer.Print ""
'      End If
'      If IsNull(adoaccrpt105.Fields("r10519").Value) = False Then
'         strAmount = Format(adoaccrpt105.Fields("r10519").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
''         Printer.CurrentX = 16600 - intLength
'         Printer.CurrentX = 15600 + 200 + Printer.TextWidth("票據金額") - intLength + IIf(Me.Text13.Text = "Y", 1920, 0)
'         Printer.CurrentY = 3800 + intCounter * 300
'         Printer.Print strAmount
'      Else
'         Printer.Print ""
'      End If
''      Printer.CurrentX = 16800
'      Printer.CurrentX = 16800 + 200 + IIf(Me.Text13.Text = "Y", 1920, 0)
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10520").Value) = False Then
'         Printer.Print CFDate(adoaccrpt105.Fields("r10520").Value)
'      Else
'         Printer.Print ""
'      End If
''      Printer.CurrentX = 17900
'      Printer.CurrentX = 17900 + 200 + IIf(Me.Text13.Text = "Y", 1920, 0)
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt105.Fields("r10521").Value) = False Then
'         Printer.Print adoaccrpt105.Fields("r10521").Value
'      Else
'         Printer.Print ""
'      End If
'      intCounter = intCounter + 1
'      adoaccrpt105.MoveNext
'   Loop
'   adoaccrpt105.Close
'   Printer.FontSize = 12
'   Printer.EndDoc
'End Sub

'2013/5/13 cancel by sonia 沒有用到
''*************************************************
''  列印類別-選擇函證用未收之小計計算
''
''*************************************************
'Private Sub SubSelect5()
' Dim strSql As String
'
'   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'      strSql = " and a0k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'   End If
'   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'      strSql = strSql & " and a0k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'   End If
'   If Text7 <> MsgText(601) Then
'      strSql = strSql & " and cp01 = '" & Text7 & "'"
'   End If
'   If Text8 <> MsgText(601) Then
'      strSql = strSql & " and cp02 = '" & Text8 & "'"
'   End If
'   If Text9 <> MsgText(601) Then
'      strSql = strSql & " and cp03 = '" & Text9 & "'"
'   End If
'   If Text10 <> MsgText(601) Then
'      strSql = strSql & " and cp04 = '" & Text10 & "'"
'   End If
'   If Text4 <> MsgText(601) Then
'      strSql = strSql & " and a0k20 = '" & Text4 & "'"
'   End If
'   adoacc0k0.MovePrevious
'   adoaccsum.CursorLocation = adUseClient
'   'adoaccsum.Open "select sum(A), sum(B), sum(C), sum(D), sum(E) from (select sum(decode(a0k30, 'Y', nvl(cp16, 0), nvl(cp16, 0) - nvl(cp17, 0))) as A, sum(decode(a0k30, 'Y', 0, nvl(cp17, 0))) as B, sum(nvl(cp16, 0)) as C, sum(nvl(cp75, 0)) as D, sum(nvl(cp16, 0) - nvl(cp75, 0)) as E from caseprogress, acc0j0, acc0k0 where cp09 = a0j01 (+) and cp60 = a0k01 (+) and (cp79 <> 0 or cp79 is null) and (a0k09 is null or a0k09 = 0) and (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0)) and a0k03 = '" & adoacc0k0.Fields("a0k03").Value & "' and instr(a0k04, '" & Trim(adoacc0k0.Fields("a0k04").Value) & "') > 0" & strSQL & _
'   '               " union select sum(decode(a0k30, 'Y', nvl(cp16, 0), nvl(cp16, 0) - nvl(cp17, 0))) as A, sum(decode(a0k30, 'Y', 0, nvl(cp17, 0))) as B, sum(nvl(cp16, 0)) as C, sum(nvl(cp75, 0)) as D, sum(nvl(cp16, 0) - nvl(cp75, 0)) as E from caseprogress, acc0j0, acc0k0, acc0m0, acc0e0 where cp09 = a0j01 (+) and cp60 = a0k01 (+) and cp60 = a0m02 (+) and a0m01 = a0e03 and (a0k09 is null or a0k09 = 0) and a0e10 > " & Val(FCDate(MaskEdBox2.Text)) & " and a0k03 = '" & adoacc0k0.Fields("a0k03").Value & "' and instr(a0k04, '" & Trim(adoacc0k0.Fields("a0k04").Value) & "') > 0" & strSQL & ") new", adoTaie, adOpenStatic, adLockReadOnly
'   'Modify By Cheng 2003/04/28
'   '依公司別跳頁
'   If Me.Text12.Text = "Y" Then
'       adoaccsum.Open "select sum(to_number(nvl(r10513, 0))), sum(to_number(nvl(r10514, 0))), sum(to_number(nvl(r10515, 0))), sum(to_number(nvl(r10516, 0))), sum(to_number(nvl(r10517, 0))) from accrpt105 where r10507 is not null and r10501 = '" & strUserNum & "' and r10503 = '" & strSameName & "' and r10522='" & m_strA0k11 & "' ", adoTaie, adOpenStatic, adLockReadOnly
'   '不依公司別跳頁
'   Else
'       adoaccsum.Open "select sum(to_number(nvl(r10513, 0))), sum(to_number(nvl(r10514, 0))), sum(to_number(nvl(r10515, 0))), sum(to_number(nvl(r10516, 0))), sum(to_number(nvl(r10517, 0))) from accrpt105 where r10507 is not null and r10501 = '" & strUserNum & "' and r10503 = '" & strSameName & "'", adoTaie, adOpenStatic, adLockReadOnly
'   End If
'   If adoaccsum.RecordCount <> 0 And (adoaccsum.Fields(0).Value <> 0 Or adoaccsum.Fields(1).Value <> 0 Or adoaccsum.Fields(2).Value <> 0 Or adoaccsum.Fields(3).Value <> 0 Or adoaccsum.Fields(4).Value <> 0) Then
'      adoaccrpt105.AddNew
'      adoaccrpt105.Fields("r10501").Value = strUserNum
'      adoaccrpt105.Fields("r10502").Value = Counter
'      If Text6 <> "" Then
'         adoaccrpt105.Fields("r10503").Value = Text6
'         adoaccrpt105.Fields("r10504").Value = Text3
'      Else
'         adoaccrpt105.Fields("r10503").Value = adoacc0k0.Fields("a0k03").Value
'         adoaccrpt105.Fields("r10504").Value = Trim(adoacc0k0.Fields("a0k04").Value)
'      End If
'       'Add By Cheng 2003/04/28
'      adoaccrpt105.Fields("r10522").Value = m_strA0k11
'      PaintLine ReportSum(4)
'      adoaccrpt105.UpdateBatch
'      adoaccrpt105.AddNew
'      adoaccrpt105.Fields("r10501").Value = strUserNum
'      adoaccrpt105.Fields("r10502").Value = Counter
'      If Text6 <> "" Then
'         adoaccrpt105.Fields("r10503").Value = Text6
'         adoaccrpt105.Fields("r10504").Value = Text3
'      Else
'         adoaccrpt105.Fields("r10503").Value = adoacc0k0.Fields("a0k03").Value
'         adoaccrpt105.Fields("r10504").Value = Trim(adoacc0k0.Fields("a0k04").Value)
'      End If
'      adoaccrpt105.Fields("r10510").Value = ReportSum(24)
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         adoaccrpt105.Fields("r10513").Value = 0
'      Else
''         If adoacc0k0.Fields("a0k30").Value = MsgText(602) Then
''            adoaccrpt105.Fields("r10513").Value = adoaccsum.Fields(2).Value
''         Else
'            adoaccrpt105.Fields("r10513").Value = adoaccsum.Fields(0).Value
''         End If
'      End If
'      If IsNull(adoaccsum.Fields(1).Value) Then
'         adoaccrpt105.Fields("r10514").Value = 0
'      Else
''         If adoacc0k0.Fields("a0k30").Value = MsgText(602) Then
''            adoaccrpt105.Fields("r10514").Value = 0
''         Else
'            adoaccrpt105.Fields("r10514").Value = adoaccsum.Fields(1).Value
''         End If
'      End If
'      If IsNull(adoaccsum.Fields(2).Value) Then
'         adoaccrpt105.Fields("r10515").Value = 0
'      Else
'         adoaccrpt105.Fields("r10515").Value = adoaccsum.Fields(2).Value
'      End If
'      If IsNull(adoaccsum.Fields(3).Value) Then
'         adoaccrpt105.Fields("r10516").Value = 0
'      Else
'         adoaccrpt105.Fields("r10516").Value = adoaccsum.Fields(3).Value
'      End If
'      If IsNull(adoaccsum.Fields(4).Value) Then
'         adoaccrpt105.Fields("r10517").Value = 0
'      Else
'         adoaccrpt105.Fields("r10517").Value = adoaccsum.Fields(4).Value
'      End If
'      'Add By Cheng 2003/04/28
'      adoaccrpt105.Fields("r10522").Value = m_strA0k11
'      adoaccrpt105.UpdateBatch
'   End If
'   adoaccsum.Close
'   adoacc0k0.MoveNext
'End Sub

'Add By Cheng 2003/06/10
'重整資料
Private Sub ReorganizeDatas1()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

If Me.Text12.Text = "Y" Then
    StrSQLa = "Select R10501, R10503, R10504, R10522 ,Count(R10501) From accrpt105 Where R10501='" & strUserNum & "' And R10502<=(" & m_lngMaxNo - 3 & ") Having Count(R10501) < 3 Group By R10501, R10503, R10504, R10522 "
Else
    StrSQLa = "Select R10501, R10503, R10504, Count(R10501) From accrpt105 Where R10501='" & strUserNum & "' And R10502<=(" & m_lngMaxNo - 3 & ") Having Count(R10501) < 3 Group By R10501, R10503, R10504 "
End If
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    While Not rsA.EOF
        If Me.Text12.Text = "Y" Then
            StrSQLa = "Delete From accrpt105 Where R10501='" & rsA.Fields(0).Value & "' And R10503='" & rsA.Fields(1).Value & "' And R10504='" & rsA.Fields(2).Value & "' And R10522='" & rsA.Fields(3).Value & "' And R10502<= " & m_lngMaxNo - 3
            cnnConnection.Execute StrSQLa
        Else
            StrSQLa = "Delete From accrpt105 Where R10501='" & rsA.Fields(0).Value & "' And R10503='" & rsA.Fields(1).Value & "' And R10504='" & rsA.Fields(2).Value & "' And R10502<= " & m_lngMaxNo - 3
            cnnConnection.Execute StrSQLa
        End If
        rsA.MoveNext
    Wend
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Sub

'Add By Cheng 2003/06/18
'若同一張收據資料全部收款了, 則刪除之
Private Sub ReorganizeDatas1_1()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

StrSQLa = "Select R10507, Sum(R10515), Sum(R10516) from accrpt105 Where R10501='" & strUserNum & "' Group By R10501, R10507 Having Sum(R10515)=Sum(R10516) "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    While Not rsA.EOF
        StrSQLa = "Delete From accrpt105 Where R10501='" & strUserNum & "' And R10507='" & rsA.Fields(0).Value & "' "
        adoTaie.Execute StrSQLa
        rsA.MoveNext
    Wend
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Sub

'Add By Cheng 2003/12/12
'取得客戶案件案號
'2013/5/13 modify by sonia, 加strYN, Y取得客戶案件案號,否則取本所案號
Private Function GetCustCaseNo(strCP60 As String, strYN As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

GetCustCaseNo = ""
If strYN = "Y" Then '2013/5/13 ADD BY SONIA
   StrSQLa = "Select PA48 From CaseProgress, Patent Where CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 And CP60='" & strCP60 & "' "
   StrSQLa = StrSQLa & " Union Select TM35 From CaseProgress, Trademark Where CP01=TM01 And CP02=TM02 And CP03=TM03 And CP04=TM04 And CP60='" & strCP60 & "' "
   StrSQLa = StrSQLa & " Union Select LC17 From CaseProgress, Lawcase Where CP01=LC01 And CP02=LC02 And CP03=LC03 And CP04=LC04 And CP60='" & strCP60 & "' "
   StrSQLa = StrSQLa & " Union Select SP29 From CaseProgress, Servicepractice Where CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04 And CP60='" & strCP60 & "' "
'2013/5/13 ADD BY SONIA
Else
   StrSQLa = "Select a0j02 From acc0j0 Where a0j13='" & strCP60 & "' "
End If
'2013/5/13 END
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetCustCaseNo = "" & rsA.Fields(0).Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

Private Sub Text9_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add by Amy 2023/08/14
Private Sub TxtDeptE_GotFocus()
   TextInverse TxtDeptE
End Sub

Private Sub TxtDeptE_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub TxtDeptS_GotFocus()
   TextInverse TxtDeptS
End Sub

Private Sub TxtDeptS_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'end 2023/08/14

Private Sub TxtDeptS_LostFocus()
   If Trim(TxtDeptS) = MsgText(601) Then Exit Sub
   
   TxtDeptE = TxtDeptS
End Sub

Private Sub txtSys_GotFocus()
   TextInverse txtSys
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtSys.IMEMode = 2
   CloseIme
End Sub

Private Sub txtSys_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add by Morgan 2005/5/31
'*************************************************
'  列印類別-選擇未收
'
'*************************************************
Private Sub Select1New()
   Dim stVTBx As String
   Dim stConK As String, stConCP As String
   Dim stChoice As String
   Dim stCustNo As String, stAddressee As String
   Dim stPreCustNo As String, stPreAddressee As String
   Dim bolAddItem As Boolean
   Dim stLstItem As String, stLstDocNo As String 'Added by Morgan 2011/12/21 記錄案件性質、收據號碼
   Dim dblAddAmt(3) As Double 'Added by Morgan 2011/12/21
   Dim stLstCaseNo As String 'Added by Lydia 2019/10/15 記錄本所案號
   
   '語法選擇 預設1 acc0k0為主
   stChoice = "1"
   strSql = ""
   stConCP = ""
   stConK = ""
   '系統類別
   If txtSys.Text <> "" Then
      stConCP = stConCP & " and cp01 in (" & GetSysSQL(txtSys.Text) & ")"
   End If
   If Text1 = "X" Then Text1 = MsgText(601)
   If Text2 = "X" Then Text2 = MsgText(601)
   '客戶代號
   If Text1 <> MsgText(601) And Text1 <> "X" Then
      stConK = stConK & " and a0k03 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) And Text2 <> "X" Then
      stConK = stConK & " and a0k03 <= '" & Text2 & "'"
   End If
   '收據抬頭
   'Modify By Sindy 2016/6/8 Text3 ==> cboTitle.Text
   'If Text3 <> MsgText(601) Then
   If cboTitle.Text <> MsgText(601) Then
      '2011/10/20 MODIFY BY SONIA E10023515
      'stConK = stConK & " and instr(a0k04, '" & Text3 & "') > 0"
      'stConK = stConK & " and instr(UPPER(a0k04), UPPER('" & Text3 & "')) > 0"
      stConK = stConK & " and instr(UPPER(a0k04), UPPER('" & cboTitle.Text & "')) > 0"
   End If
   '智權人員
   If Text4 <> MsgText(601) Then
      stConK = stConK & " and a0k20 = '" & Text4 & "'"
   End If
   '作業日期
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      stConK = stConK & " and a0k02 >= " & Val(FCDate(MaskEdBox1.Text))
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      stConK = stConK & " and a0k02 <= " & Val(FCDate(MaskEdBox2.Text))
   End If
   '本所案號
   If Text7 <> "" And Text8 <> "" Then
      stConCP = stConCP & " and cp01='" & Text7 & "' and cp02='" & Text8 & "' and cp03='" & Text9 & "' and cp04='" & Text10 & "'"
      stChoice = "2"
   End If
   '公司別
   If Me.Text14.Text <> "" Then
       stConK = stConK & " and A0K11 >= '" & Me.Text14.Text & "'"
   End If
   If Me.Text15.Text <> "" Then
       stConK = stConK & " and A0K11 <= '" & Me.Text15.Text & "'"
   End If
   
   '2012/3/27 Add by Sindy +不包含未列印收據故加 and a0k32 is null條件
   If Me.Text16.Text = "" Then
       stConK = stConK & " and A0K32 is null"
   End If
   
   'Modified by Morgan 2011/11/1 考慮拆收據情形
   'stVTBx = "select cp60,cp09,sum(a1u04) a1u04,sum(a1u05) a1u05,sum(a1u07) a1u07" & _
         ",sum(a1u08) a1u08,sum(a1u09) a1u09,sum(a1u10) a1u10"
   stVTBx = "select a0j13 as a1u02,a0j01 as a1u03,sum(a1u04) a1u04,sum(a1u05) a1u05,sum(a1u07) a1u07" & _
         ",sum(a1u08) a1u08,sum(a1u09) a1u09,sum(a1u10) a1u10"
         
   '有本所號時以caseprogress為主
   If stChoice = "2" Then
      'Modified by Morgan 2011/11/1 考慮拆收據情形
      'stVTBx = stVTBx & " from caseprogress,acc1u0 where cp60 is not null"
      stVTBx = stVTBx & " from caseprogress,acc0j0,acc1u0 where cp60 is not null and a0j01(+)=cp09"
   '一般以acc0k0為主
   Else
      'Modified by Morgan 2011/11/1 考慮拆收據情形
      'stVTBx = stVTBx & " from acc0k0,caseprogress,acc1u0 where cp60(+)=a0k01" & stConK
      stVTBx = stVTBx & " from acc0k0,acc0j0,caseprogress,acc1u0 where a0j13(+)=a0k01 and cp09(+)=a0j01" & stConK
   End If
   'Modified by Morgan 2011/11/1 考慮拆收據情形
   'stVTBx = stVTBx & " and a1u03(+)=cp09 and cp79>0" & stConCP & " group by cp60,cp09"
   stVTBx = stVTBx & " and a1u02(+)=a0j13 and a1u03(+)=a0j01 and cp79>0" & stConCP & " group by a0j13,a0j01"
   
   'Modified by Morgan 2011/11/1 考慮拆收據情形
   'strSql = "select a0k03 CustNo,a0k04 Addressee,a0k11 Comp,a0K02 DocDate,a0k01 DocNo,a0j01 RecNo,st02 Sales" & _
      ",a0j20,a0j21,nvl(a0j09,0)-nvl(a1u07,0) Fee1,nvl(A0j10,0)-nvl(a1u09,0) Fee2" & _
      ",nvl(a1u04,0)+nvl(a1u05,0)-nvl(a1u08,0)-nvl(a1u10,0) Fee3,a0k30" & _
      " from (" & stVTBx & ") Vx,acc0k0,acc0j0,staff" & _
      " where a0k01(+)=cp60 and a0j01(+)=cp09 and st01(+)=a0k20" & stConK
   'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
   'Modified by Lydia 2019/10/15 +本所案號CaseNo
   strSql = "select a0k03 CustNo,a0k04 Addressee,a0k11 Comp,a0K02 DocDate,a0k01 DocNo,a0j01 RecNo,st02 Sales" & _
      ",getcp10desc(cp01,cp10,a0j04) cp10N,na03,nvl(a0j09,0)-nvl(a1u07,0) Fee1,nvl(A0j10,0)-nvl(a1u09,0) Fee2" & _
      ",nvl(a1u04,0)+nvl(a1u05,0)-nvl(a1u08,0)-nvl(a1u10,0) Fee3,a0j07,a0k33,a0j22,a0j25" & _
      ",cp01||cp02||cp03||cp04 as CaseNo" & _
      " from (" & stVTBx & ") Vx,acc0j0,acc0k0,staff,caseprogress,nation" & _
      " where a0j13(+)=a1u02 and a0j01(+)=a1u03 and a0k01(+)=a1u02 and st01(+)=a0k20" & stConK & " and cp09(+)=a0j01 and na01(+)=a0j04"
   
   '若非北所員工, 只能列印該所資料
   If pub_strUserOffice <> "1" Then
       strSql = strSql & " And ''||ST06='" & pub_strUserOffice & "'"
   End If
   
   'Add by Amy 2023/08/14 +部門
   If TxtDeptS <> MsgText(601) Then
      strSql = strSql & " And ST15>='" & TxtDeptS & "' "
   End If
   If TxtDeptE <> MsgText(601) Then
      strSql = strSql & " And ST15<='" & TxtDeptE & "' "
   End If
   'end 2023/08/14
   
   strSql = "select * from (" & strSql & ") where Fee1+Fee2-Fee3>0" 'Added by Morgan 2011/11/1 考慮拆收據情形
   
   '依公司別跳頁
   If Me.Text12.Text = "Y" Then
      strSql = strSql & " order by CustNo asc, Comp asc, Addressee asc, DocDate asc"
   Else
      strSql = strSql & " order by CustNo asc, Addressee asc, DocDate asc"
   End If
   
   'Added by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
   'Modified by Lydia 2019/10/15 比照智權人員-客戶請款明細表的順序: 收據日期a0k02->收據號碼a0k01->收文號
   'strSql = strSql & ",DocNo asc,a0j25 asc"
   strSql = strSql & ",DocNo asc, RecNo asc, a0j25 asc"
   
On Error GoTo ErrHnd

   If adoacc0m0.State = adStateOpen Then adoacc0m0.Close
   With adoacc0m0
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         Erase lngSubTot
         Erase lngTot
         If Text6 <> "" Then
            stPreCustNo = Text6.Text
            'Modify By Sindy 2016/6/8 Text3 ==> cboTitle.Text
            stPreAddressee = cboTitle.Text
            '2016/6/8 END
         Else
            stPreCustNo = "" & .Fields("CustNo").Value
            stPreAddressee = "" & .Fields("Addressee").Value
         End If
         m_strA0k11 = "" & .Fields("Comp").Value
         
         Do While Not .EOF
            If Text6 <> "" Then
               stCustNo = Text6.Text
               'Modify By Sindy 2016/6/8 Text3 ==> cboTitle.Text
               'stAddressee = Text3.Text
               stAddressee = cboTitle.Text
               '2016/6/8 END
            Else
               stCustNo = "" & .Fields("CustNo").Value
               stAddressee = "" & .Fields("Addressee").Value
            End If
            If stCustNo & stAddressee <> stPreCustNo & stPreAddressee Then
               SubSelect stPreCustNo, stPreAddressee, lngSubTot()
               stPreCustNo = stCustNo: stPreAddressee = stAddressee
               For jj = 1 To 5
                  lngTot(jj) = lngTot(jj) + lngSubTot(jj)
               Next
               Erase lngSubTot
            '若依公司別跳頁 公司別不同時
            ElseIf Me.Text12.Text = "Y" And "" & .Fields("Comp").Value <> m_strA0k11 Then
               SubSelect stCustNo, stAddressee, lngSubTot()
               For jj = 1 To 5
                  lngTot(jj) = lngTot(jj) + lngSubTot(jj)
               Next
               Erase lngSubTot
            End If
            m_strA0k11 = "" & .Fields("Comp").Value
            
            'Added by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
            'Modified by Lydia 2019/10/15 +本所案號
            'If stLstDocNo = .Fields("DocNo") And .Fields("a0k33") = "Y" And stLstItem = .Fields("a0j22") Then
            If stLstDocNo = .Fields("DocNo") And .Fields("a0k33") = "Y" And stLstItem = .Fields("a0j22") And stLstCaseNo = "" & .Fields("CaseNo") Then
               bolAddItem = False
            Else
               bolAddItem = True
            End If
            stLstDocNo = "" & .Fields("DocNo")
            stLstItem = "" & .Fields("a0j22")
            stLstCaseNo = "" & .Fields("CaseNo").Value 'Added by Lydia 2019/10/15
            If bolAddItem = True Then
            'end 2011/12/21
            
               adoaccrpt105.AddNew
               adoaccrpt105.Fields("r10501").Value = strUserNum
               adoaccrpt105.Fields("r10502").Value = Counter
               adoaccrpt105.Fields("r10503").Value = stCustNo
               adoaccrpt105.Fields("r10504").Value = stAddressee
               adoaccrpt105.Fields("r10505").Value = .Fields("DocDate").Value
               adoaccrpt105.Fields("r10507").Value = .Fields("DocNo").Value
               adoaccrpt105.Fields("r10508").Value = .Fields("Sales").Value
               adoaccrpt105.Fields("r10510").Value = MidB(CaseNameQuery(.Fields("RecNo").Value, 1), 1, 24)
               'Modified by Morgan 2011/12/21
               If .Fields("a0k33") = "Y" Then
                  adoaccrpt105.Fields("r10511").Value = .Fields("a0j22").Value
               Else
                  'Modified by Morgan 2011/12/27 取消 a0j20
                  adoaccrpt105.Fields("r10511").Value = .Fields("cp10N").Value
               End If
               'end 2011/12/21
               
               'Modified by Morgan 2011/12/30 取消 a0j21
               adoaccrpt105.Fields("r10512").Value = .Fields("na03").Value
               '是否合併
               If "" & .Fields("a0j07").Value = "Y" Then
                  '服務費
                  adoaccrpt105.Fields("r10513").Value = Format(.Fields("Fee1").Value + .Fields("Fee2").Value)
                  '規費
                  adoaccrpt105.Fields("r10514").Value = 0
               Else
                  '服務費
                  adoaccrpt105.Fields("r10513").Value = Format(.Fields("Fee1").Value)
                  '規費
                  adoaccrpt105.Fields("r10514").Value = Format(.Fields("Fee2").Value)
               End If
               adoaccrpt105.Fields("r10515").Value = Format(Val(adoaccrpt105.Fields("r10513").Value) + Val(adoaccrpt105.Fields("r10514").Value))
               adoaccrpt105.Fields("r10516").Value = Format(.Fields("Fee3").Value)
               adoaccrpt105.Fields("r10517").Value = Format(Val(adoaccrpt105.Fields("r10515").Value) - Val(adoaccrpt105.Fields("r10516").Value))
               lngSubTot(1) = lngSubTot(1) + adoaccrpt105.Fields("r10513").Value
               lngSubTot(2) = lngSubTot(2) + adoaccrpt105.Fields("r10514").Value
               lngSubTot(3) = lngSubTot(3) + Val(adoaccrpt105.Fields("r10515").Value)
               lngSubTot(4) = lngSubTot(4) + adoaccrpt105.Fields("r10516").Value
               lngSubTot(5) = lngSubTot(5) + adoaccrpt105.Fields("r10517").Value
               adoaccrpt105.Fields("r10522").Value = "" & .Fields("Comp").Value
               '客戶案件案號
               '2013/5/13 modify by sonia 不列印客戶案件案號改印本所案號
               'If Me.Text13.Text = "Y" Then
               '    adoaccrpt105.Fields("r10523").Value = GetCustCaseNo("" & .Fields("DocNo").Value)
               'Else
               '    adoaccrpt105.Fields("r10523").Value = ""
               'End If
               'Modified by Lydia 2019/10/15 直接抓本所案號
               'adoaccrpt105.Fields("r10523").Value = GetCustCaseNo("" & .Fields("DocNo").Value, Me.Text13.Text)
               adoaccrpt105.Fields("r10523").Value = "" & .Fields("CaseNo")
               '2013/5/13 end
            
            'Added by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
            Else
               Erase dblAddAmt
               
               '是否合併
               If "" & .Fields("a0j07").Value = "Y" Then
                  '服務費
                  dblAddAmt(1) = .Fields("Fee1").Value + .Fields("Fee2").Value
                  '規費
                  dblAddAmt(2) = 0
               Else
                  '服務費
                  dblAddAmt(1) = .Fields("Fee1").Value
                  '規費
                  dblAddAmt(2) = .Fields("Fee2").Value
               End If
               dblAddAmt(3) = .Fields("Fee3").Value
               
               adoaccrpt105.Fields("r10513").Value = Format(adoaccrpt105.Fields("r10513").Value + dblAddAmt(1))
               adoaccrpt105.Fields("r10514").Value = Format(adoaccrpt105.Fields("r10514").Value + dblAddAmt(2))
               adoaccrpt105.Fields("r10515").Value = Format(Val(adoaccrpt105.Fields("r10513").Value) + Val(adoaccrpt105.Fields("r10514").Value))
               adoaccrpt105.Fields("r10516").Value = Format(adoaccrpt105.Fields("r10516").Value + dblAddAmt(3))
               adoaccrpt105.Fields("r10517").Value = Format(Val(adoaccrpt105.Fields("r10515").Value) - Val(adoaccrpt105.Fields("r10516").Value))
               
               lngSubTot(1) = lngSubTot(1) + dblAddAmt(1)
               lngSubTot(2) = lngSubTot(2) + dblAddAmt(2)
               lngSubTot(3) = lngSubTot(3) + dblAddAmt(1) + dblAddAmt(2)
               lngSubTot(4) = lngSubTot(4) + dblAddAmt(3)
               lngSubTot(5) = lngSubTot(5) + dblAddAmt(1) + dblAddAmt(2) - dblAddAmt(3)
            End If
            'end 2011/12/21
            
            adoaccrpt105.UpdateBatch
            .MoveNext
         Loop
         '小計
         SubSelect stCustNo, stAddressee, lngSubTot()
         For jj = 1 To 5
            lngTot(jj) = lngTot(jj) + lngSubTot(jj)
         Next
         '合計
         SubSelect stCustNo, stAddressee, lngTot(), 2
      Else
         MsgBox MsgText(28), , MsgText(5)
      End If
   End With
   If adoacc0m0.State = adStateOpen Then adoacc0m0.Close
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   'Resume
End Sub
'Add by Morgan 2005/5/27
'*************************************************
'  列印類別-選擇收回
'
'*************************************************
Private Sub Select2New()

   Dim stVTBx As String
   Dim stConJ As String, stConK As String, stConL As String
   Dim stChoice As String
   Dim stCustNo As String, stAddressee As String
   Dim stPreCustNo As String, stPreAddressee As String
   Dim bolAddItem As Boolean
   Dim stLstItem As String, stLstDocNo As String 'Added by Morgan 2011/12/21 記錄案件性質、收據號碼
   Dim dblAddAmt(3) As Double 'Added by Morgan 2011/12/21
   Dim stLstCaseNo As String 'Added by Lydia 2019/10/15 記錄本所案號

   'add by nickc 2007/02/08
   Dim stCon
   '語法選擇 預設1 acc0k0為主
   stChoice = "1"
   strSql = ""
   stConJ = ""
   stConK = ""
   stConL = ""
   '系統類別
   If txtSys.Text <> "" Then
      stConJ = stConJ & " and ltrim(substr(lpad(a0j02,12,' '),1,3)) in (" & GetSysSQL(txtSys.Text) & ")"
   End If
   If Text1 = "X" Then Text1 = MsgText(601)
   If Text2 = "X" Then Text2 = MsgText(601)
   '客戶代號
   If Text1 <> MsgText(601) And Text1 <> "X" Then
      stConK = stConK & " and a0k03 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) And Text2 <> "X" Then
      stConK = stConK & " and a0k03 <= '" & Text2 & "'"
   End If
   '收據抬頭
   'Modify By Sindy 2016/6/8 Text3 ==> cboTitle.Text
   'If Text3 <> MsgText(601) Then
   If cboTitle.Text <> MsgText(601) Then
      '2011/10/20 MODIFY BY SONIA E10023515
      'stConK = stConK & " and instr(a0k04, '" & Text3 & "') > 0"
      'stConK = stConK & " and instr(UPPER(a0k04), UPPER('" & Text3 & "')) > 0"
      stConK = stConK & " and instr(UPPER(a0k04), UPPER('" & cboTitle.Text & "')) > 0"
   End If
   '智權人員
   If Text4 <> MsgText(601) Then
      stConK = stConK & " and a0k20 = '" & Text4 & "'"
   End If
   '作業日期
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      stConL = stConL & " and a0l02 >= " & Val(FCDate(MaskEdBox1.Text))
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      stConL = stConL & " and a0l02 <= " & Val(FCDate(MaskEdBox2.Text))
   End If
   '本所案號
   If Text7 <> "" And Text8 <> "" Then
      stConJ = stConJ & " and a0j02='" & Text7 & Text8 & Text9 & Text10 & "'"
      stChoice = "2"
   End If
   '公司別
   If Me.Text14.Text <> "" Then
       stConK = stConK & " and A0K11 >= '" & Me.Text14.Text & "'"
   End If
   If Me.Text15.Text <> "" Then
       stConK = stConK & " and A0K11 <= '" & Me.Text15.Text & "'"
   End If
   
   '2012/3/27 Add by Sindy +不包含未列印收據故加 and a0k32 is null條件
   If Me.Text16.Text = "" Then
       'Modified by Lydia 2025/06/10 a0k32 is null 改用函數判斷：geta0k32type(a0k01)='1'
       stConK = stConK & " and geta0k32type(a0k01)='1'"
   End If
   
   'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
   strSql = "select a0k03 CustNo,a0k04 Addressee,a0k11 Comp,a0l02 DocDate,a0k01 DocNo,a0j01 RecNo,st02 Sales" & _
      ",a0j07,getcp10desc(cp01,cp10,a0j04) cp10N,na03,a1u04,a1u05,a1u06,a1p22,a0k33,a0j22,a0j25"
   '有本所號時以acc0j0為主
   If stChoice = "2" Then
      '抓傳票號
      stVTBx = " select distinct a1p04,a1p22" & _
         " from acc0j0,acc1u0,acc1p0" & _
         " where a1u03(+)=a0j01" & stConJ & _
         " and a1p04(+)=a1u01 and a1p01 = '1' and a1p02 = 'A'"
      'Modified by Lydia 2019/10/15 +本所案號CaseNo
      strSql = strSql & _
         ",cp01||cp02||cp03||cp04 as CaseNo" & _
         " from acc0j0,acc0k0,acc1u0, acc0l0, staff" & _
         " ,(" & stVTBx & " ) X,caseprogress,nation where a0k01(+)=a0j13 and cp09(+)=a0j01 and na01(+)=a0j04 "
   '一般以acc0k0為主
   Else
      '抓傳票號
      stVTBx = " select distinct a1p04,a1p22" & _
         " from acc0k0,acc1u0,acc1p0" & _
         " where nvl(a0k09,0)=0" & stConK & stCon & _
         " and a1u02(+)=a0k01 and a1u01 LIKE 'F%' and a1p04(+)=a1u01" & _
         " and a1p01 = '1' and a1p02 = 'A'"
      'Modified by Lydia 2019/10/15 +本所案號CaseNo
      strSql = strSql & _
         ",cp01||cp02||cp03||cp04 as CaseNo" & _
         " from acc0k0,acc0j0,acc1u0, acc0l0, staff" & _
         " ,(" & stVTBx & " ) X,caseprogress,nation where a0j13(+)=a0k01 and cp09(+)=a0j01 and na01(+)=a0j04 "
   End If
   'Modified by Morgan 2011/11/1 考慮拆收據情形
   'strSql = strSql & _
      " and a1u03(+)=a0j01 and a1u01 LIKE 'F%'" & stConJ & stConK & stConL & _
      " and a0l01(+)=a1u01 and st01(+)=a0k20 and a1p04(+)=a1u01"
   strSql = strSql & _
      " and a1u02(+)=a0j13 and a1u03(+)=a0j01 and a1u01 LIKE 'F%'" & stConJ & stConK & stConL & _
      " and a0l01(+)=a1u01 and st01(+)=a0k20 and a1p04(+)=a1u01"
   
   '若非北所員工, 只能列印該所資料
   If pub_strUserOffice <> "1" Then
       strSql = strSql & " And ''||ST06='" & pub_strUserOffice & "'"
   End If
   
   'Add by Amy 2023/08/14 +部門
   If TxtDeptS <> MsgText(601) Then
      strSql = strSql & " And ST15>='" & TxtDeptS & "' "
   End If
   If TxtDeptE <> MsgText(601) Then
      strSql = strSql & " And ST15<='" & TxtDeptE & "' "
   End If
   'end 2023/08/14
   
   '依公司別跳頁
   If Me.Text12.Text = "Y" Then
      strSql = strSql & " order by CustNo asc, Comp asc, Addressee asc, DocDate asc"
   Else
      strSql = strSql & " order by CustNo asc, Addressee asc, DocDate asc"
   End If
   
   'Added by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
   'Modified by Lydia 2019/10/15 比照智權人員-客戶請款明細表的順序: 收據日期a0k02->收據號碼a0k01->收文號
   'strSql = strSql & ",DocNo asc,a0j25 asc"
   strSql = strSql & ",DocNo asc, RecNo asc, a0j25 asc"
   
On Error GoTo ErrHnd

   If adoacc0m0.State = adStateOpen Then adoacc0m0.Close
   With adoacc0m0
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         Erase lngSubTot
         Erase lngTot
         If Text6 <> "" Then
            stPreCustNo = Text6.Text
            'Modify By Sindy 2016/6/8 Text3 ==> cboTitle.Text
            'stPreAddressee = Text3.Text
            stPreAddressee = cboTitle.Text
            '2016/6/8 END
         Else
            stPreCustNo = "" & .Fields("CustNo").Value
            stPreAddressee = "" & .Fields("Addressee").Value
         End If
         m_strA0k11 = "" & .Fields("Comp").Value
         Do While Not .EOF
            If Text6 <> "" Then
               stCustNo = Text6.Text
               'Modify By Sindy 2016/6/8 Text3 ==> cboTitle.Text
               'stAddressee = Text3.Text
               stAddressee = cboTitle.Text
               '2016/6/8 END
            Else
               stCustNo = "" & .Fields("CustNo").Value
               stAddressee = "" & .Fields("Addressee").Value
            End If
            If stCustNo & stAddressee <> stPreCustNo & stPreAddressee Then
               SubSelect stPreCustNo, stPreAddressee, lngSubTot()
               stPreCustNo = stCustNo: stPreAddressee = stAddressee
               For jj = 1 To 5
                  lngTot(jj) = lngTot(jj) + lngSubTot(jj)
               Next
               Erase lngSubTot
            '若依公司別跳頁 公司別不同時
            ElseIf Me.Text12.Text = "Y" And "" & .Fields("Comp").Value <> m_strA0k11 Then
               SubSelect stCustNo, stAddressee, lngSubTot()
               For jj = 1 To 5
                  lngTot(jj) = lngTot(jj) + lngSubTot(jj)
               Next
               Erase lngSubTot
            End If
            m_strA0k11 = "" & .Fields("Comp").Value
            
            'Added by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
            'Modified by Lydia 2019/10/15 +本所案號
            If stLstDocNo = .Fields("DocNo") And .Fields("a0k33") = "Y" And stLstItem = .Fields("a0j22") And stLstCaseNo = "" & .Fields("CaseNo") Then
               bolAddItem = False
            Else
               bolAddItem = True
            End If
            stLstDocNo = "" & .Fields("DocNo")
            stLstItem = "" & .Fields("a0j22")
            stLstCaseNo = "" & .Fields("CaseNo").Value 'Added by Lydia 2019/10/15
            If bolAddItem = True Then
            'end 2011/12/21
            
               adoaccrpt105.AddNew
               adoaccrpt105.Fields("r10501").Value = strUserNum
               adoaccrpt105.Fields("r10502").Value = Counter
               adoaccrpt105.Fields("r10503").Value = stCustNo
               adoaccrpt105.Fields("r10504").Value = stAddressee
               adoaccrpt105.Fields("r10505").Value = .Fields("DocDate").Value
               adoaccrpt105.Fields("r10507").Value = .Fields("DocNo").Value
               adoaccrpt105.Fields("r10508").Value = .Fields("Sales").Value
               adoaccrpt105.Fields("r10510").Value = MidB(CaseNameQuery(.Fields("RecNo").Value, 1), 1, 24)
               'Modified by Morgan 2011/12/21
               If .Fields("a0k33") = "Y" Then
                  adoaccrpt105.Fields("r10511").Value = .Fields("a0j22").Value
               Else
                  'Modified by Morgan 2011/12/27 取消 a0j20
                  adoaccrpt105.Fields("r10511").Value = .Fields("cp10N").Value
               End If
               
               'Modified by Morgan 2011/12/30 取消 a0j21
               adoaccrpt105.Fields("r10512").Value = .Fields("na03").Value
               
               adoaccrpt105.Fields("r10513").Value = Format(Val("" & .Fields("a1u04").Value) + Val("" & .Fields("a1u05").Value))
               If "" & .Fields("a0j07").Value = "Y" Then
                  adoaccrpt105.Fields("r10514").Value = Format(adoaccrpt105.Fields("r10513").Value / 10)
               Else
                  adoaccrpt105.Fields("r10514").Value = Format(Val("" & .Fields("a1u04").Value) / 10)
               End If
               adoaccrpt105.Fields("r10515").Value = Format(.Fields("a1u06").Value)
               adoaccrpt105.Fields("r10516").Value = Format(Val("" & adoaccrpt105.Fields("r10514").Value) - Val("" & adoaccrpt105.Fields("r10515").Value))
               
               lngSubTot(1) = lngSubTot(1) + adoaccrpt105.Fields("r10513").Value
               lngSubTot(2) = lngSubTot(2) + adoaccrpt105.Fields("r10514").Value
               lngSubTot(3) = lngSubTot(3) + adoaccrpt105.Fields("r10515").Value
               lngSubTot(4) = lngSubTot(4) + adoaccrpt105.Fields("r10516").Value
               
               adoaccrpt105.Fields("r10518").Value = .Fields("a1p22").Value
               adoaccrpt105.Fields("r10522").Value = "" & .Fields("Comp").Value
               '客戶案件案號
               '2013/5/13 modify by sonia 不列印客戶案件案號改印本所案號
               'If Me.Text13.Text = "Y" Then
               '    adoaccrpt105.Fields("r10523").Value = GetCustCaseNo("" & .Fields("DocNo").Value)
               'Else
               '    adoaccrpt105.Fields("r10523").Value = ""
               'End If
               'Modified by Lydia 2019/10/15 直接抓本所案號
               'adoaccrpt105.Fields("r10523").Value = GetCustCaseNo("" & .Fields("DocNo").Value, Me.Text13.Text)
               adoaccrpt105.Fields("r10523").Value = "" & .Fields("CaseNo")
               '2013/5/13 end
            
            'Added by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
            Else
               Erase dblAddAmt
               
               dblAddAmt(1) = Val("" & .Fields("a1u04").Value) + Val("" & .Fields("a1u05").Value)
               If "" & .Fields("a0j07").Value = "Y" Then
                  dblAddAmt(2) = Format(dblAddAmt(1) / 10)
               Else
                  dblAddAmt(2) = Format(Val("" & .Fields("a1u04").Value) / 10)
               End If
               dblAddAmt(3) = .Fields("a1u06").Value
               'Modified by Morgan 2023/11/27 +Format否則遇小數會造成多重步驟錯誤(欄位型態轉換問題)
               adoaccrpt105.Fields("r10513").Value = Format(adoaccrpt105.Fields("r10513").Value + dblAddAmt(1))
               adoaccrpt105.Fields("r10514").Value = Format(adoaccrpt105.Fields("r10514").Value + dblAddAmt(2))
               adoaccrpt105.Fields("r10515").Value = Format(adoaccrpt105.Fields("r10515").Value + dblAddAmt(3))
               adoaccrpt105.Fields("r10516").Value = Format(adoaccrpt105.Fields("r10516").Value + dblAddAmt(2) - dblAddAmt(3))
               
               lngSubTot(1) = lngSubTot(1) + dblAddAmt(1)
               lngSubTot(2) = lngSubTot(2) + dblAddAmt(2)
               lngSubTot(3) = lngSubTot(3) + dblAddAmt(3)
               lngSubTot(4) = lngSubTot(4) + dblAddAmt(2) - dblAddAmt(3)
            End If
            'end 2011/12/21
            
            adoaccrpt105.UpdateBatch
            .MoveNext
         Loop
         '小計
         SubSelect stCustNo, stAddressee, lngSubTot()
         For jj = 1 To 5
            lngTot(jj) = lngTot(jj) + lngSubTot(jj)
         Next
         '合計
         SubSelect stCustNo, stAddressee, lngTot(), 2
      Else
         MsgBox MsgText(28), , MsgText(5)
      End If
   End With
   If adoacc0m0.State = adStateOpen Then adoacc0m0.Close
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   'Resume
End Sub

'Add by Morgan 2005/5/30
'*************************************************
'  列印類別-選擇往來日期
'
'*************************************************
Private Sub Select3New()

   Dim stVTBx As String
   Dim stConJ As String, stConK As String, stConL As String, stConS As String, stConCP As String
   Dim stChoice As String
   Dim stCustNo As String, stAddressee As String
   Dim stPreCustNo As String, stPreAddressee As String
   Dim bolAddItem As Boolean
   Dim stLstItem As String, stLstDocNo As String 'Added by Morgan 2011/12/21 記錄案件性質、收據號碼
   Dim dblAddAmt(4) As Double 'Added by Morgan 2011/12/21
   Dim stLstCaseNo As String 'Added by Lydia 2019/10/15 記錄本所案號
   Dim stConST As String 'Add by Amy 2023/08/14
   
   '語法選擇 預設1 acc0k0為主
   stChoice = "1"
   strSql = ""
   stConJ = ""
   stConK = ""
   stConL = ""
   stConS = ""
   stConCP = ""
   stConST = "" 'Add by Amy 2023/08/14
   '系統類別
   If txtSys.Text <> "" Then
      stConJ = stConJ & " and ltrim(substr(lpad(a0j02,12,' '),1,3)) in (" & GetSysSQL(txtSys.Text) & ")"
      stConCP = stConCP & " and cp01 in (" & GetSysSQL(txtSys.Text) & ")"
   End If
   
   'Add by Amy 2023/08/14 +部門
   If TxtDeptS <> MsgText(601) Then
      stConST = stConST & " And ST15>='" & TxtDeptS & "' "
   End If
   If TxtDeptE <> MsgText(601) Then
      stConST = stConST & " And ST15<='" & TxtDeptE & "' "
   End If
   'end 2023/08/14
   
   '客戶代號
   If Text1 = "X" Then Text1 = MsgText(601)
   If Text2 = "X" Then Text2 = MsgText(601)
   If Text1 <> MsgText(601) Then
      stConK = stConK & " and a0k03 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) And Text2 <> "X" Then
      stConK = stConK & " and a0k03 <= '" & Text2 & "'"
   End If
   '收據抬頭
   'Modify By Sindy 2016/6/8 Text3 ==> cboTitle.Text
   'If Text3 <> MsgText(601) Then
   If cboTitle.Text <> MsgText(601) Then
      '2011/10/20 MODIFY BY SONIA E10023515
      'stConK = stConK & " and instr(a0k04, '" & Text3 & "') > 0"
      'stConK = stConK & " and instr(UPPER(a0k04), UPPER('" & Text3 & "')) > 0"
      stConK = stConK & " and instr(UPPER(a0k04), UPPER('" & cboTitle.Text & "')) > 0"
   End If
   '智權人員
   If Text4 <> MsgText(601) Then
      stConK = stConK & " and a0k20 = '" & Text4 & "'"
   End If
   '作業日期
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      stConK = stConK & " and a0k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      stConL = stConL & " and a0l02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      stConS = stConS & " and a0s03 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      stConK = stConK & " and a0k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      stConL = stConL & " and a0l02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      stConS = stConS & " and a0s03 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   '本所案號
   If Text7 <> "" And Text8 <> "" Then
      stConCP = stConCP & " and cp01='" & Text7 & "' and cp02='" & Text8 & "' and cp03='" & Text9 & "' and cp04='" & Text10 & "'"
      stConJ = stConJ & " and a0j02='" & Text7 & Text8 & Text9 & Text10 & "'"
      stChoice = "2"
   End If
   '公司別
   If Me.Text14.Text <> "" Then
       stConK = stConK & " and A0K11 >= '" & Me.Text14.Text & "' "
   End If
   If Me.Text15.Text <> "" Then
       stConK = stConK & " and A0K11 <= '" & Me.Text15.Text & "'"
   End If
   
   '2012/3/27 Add by Sindy +不包含未列印收據故加 and a0k32 is null條件
   If Me.Text16.Text = "" Then
       'Modified by Lydia 2025/06/10 a0k32 is null 改用函數判斷：geta0k32type(a0k01)='1'
       stConK = stConK & " and geta0k32type(a0k01)='1'"
   End If
   
   '有本所號時以acc0j0為主
   If stChoice = "2" Then
      '收據資料
      'Modified by Morgan 2011/11/2 考慮拆收據情形(應收只要減銷帳不必加退費金額--辜)
      'strSql = "select a0k03 CustNo,a0k04 Addressee,a0k11 Comp,a0k02 DocDate,a0k01 DocNo,a0j01 RecNo,st02 Sales" & _
         ",a0j20,a0j21,nvl(cp16,0)-nvl(cp77,0)+nvl(cp78,0) TAmount, 0 RAmount, 0 VAmount, 0 BAmount" & _
         " From caseprogress,acc0k0,acc0j0,Staff" & _
         " where a0k01(+)=cp60 and a0j01(+)=cp09 and ST01(+)=a0k20 and nvl(a0k09,0)=0" & stConCP & stConK
      'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
      'Modified by Lydia 2019/10/15 +本所案號CaseNo
      'Modify by Amy 2023/08/14 +部門條件stConST
      strSql = "select a0k03 CustNo,a0k04 Addressee,a0k11 Comp,a0k02 DocDate" & _
         ",a0k01 DocNo,a0j01 RecNo,st02 Sales,getcp10desc(cp01,cp10,a0j04) cp10N,na03" & _
         ",nvl(a0j09,0)+nvl(a0j10,0)-x3 TAmount" & _
         ", 0 RAmount, 0 VAmount, 0 BAmount" & _
         ",a0k33,a0j22,a0j25,cp01||cp02||cp03||cp04 as CaseNo" & _
         " From (select cp09 x1,a0j13 x2,nvl(sum(a1u07),0)+nvl(sum(a1u09),0) x3" & _
         " from caseprogress,acc0j0,acc0k0,acc1u0 where a0j01(+)=cp09 and a0k01(+)=a0j13" & _
         " and nvl(a0k09,0)=0" & stConCP & stConK & _
         " and a1u02(+)=a0j13 and a1u03(+)=a0j01 group by cp09,a0j13),caseprogress,acc0j0,acc0k0,Staff,nation" & _
         " where cp09(+)=x1 and a0j01(+)=x1 and a0j13(+)=x2 and a0k01(+)=x2 and ST01(+)=a0k20 and na01(+)=a0j04 " & stConST
      '收款資料
      'Modified by Morgan 2011/11/2 考慮拆收據情形 +and a1u02(+)=a0j13
      'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
      'Modified by Lydia 2019/10/15 +本所案號CaseNo
      'Modify by Amy 2023/08/14 +部門條件stConST
      strSql = strSql & " Union All" & _
         " select a0k03 CustNo, a0k04 Addressee,a0k11 Comp,a0l02 DocDate,a0k01 DocNo,a0j01 RecNo,st02 Sales" & _
         ",getcp10desc(cp01,cp10,a0j04) cp10N,na03,0 TAmount, nvl(a1u04,0)+nvl(a1u05,0) RAmount, a1u06 VAmount, 0 BAmount" & _
         ",a0k33,a0j22,a0j25,cp01||cp02||cp03||cp04 as CaseNo" & _
         " from acc0j0,acc0k0,acc1u0,acc0l0,staff,caseprogress,nation" & _
         " where a0k01(+)=a0j13" & stConJ & stConK & stConL & stConST & _
         " and a1u03(+)=a0j01 and a1u02(+)=a0j13 and a1u01 LIKE 'F%'" & _
         " and a0l01(+)=a1u01 and st01(+)=a0k20 and cp09(+)=a0j01 and na01(+)=a0j04 "
         
      '退費資料
      'Modified by Morgan 2011/11/2 考慮拆收據情形　+and a1u02(+)=a0j13
      'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
      'Modified by Lydia 2019/10/15 +本所案號CaseNo
      'Modify by Amy 2023/08/14 +部門條件stConST
      strSql = strSql & " Union All" & _
         " select a0k03 CustNo, a0k04 Addressee,a0k11 Comp,a0s03 DocDate,a0k01 DocNo,a0j01 RecNo,st02 Sales" & _
         ",getcp10desc(cp01,cp10,a0j04) cp10N,na03,0 TAmount,0 RAmount, a1u06 VAmount,nvl(a1u08,0)+nvl(a1u10,0) BAmount" & _
         ",a0k33,a0j22,a0j25,cp01||cp02||cp03||cp04 as CaseNo" & _
         " from acc0j0,acc0k0,acc1u0,acc0s0,staff,caseprogress,nation" & _
         " where a0k01(+)=a0j13" & stConJ & stConK & stConS & stConST & _
         " and a1u03(+)=a0j01 and a1u02(+)=a0j13 and a1u01 LIKE 'I%'" & _
         " and a0s02(+)=a0k01 and st01(+)=a0k20 and nvl(a1u08,0)+nvl(a1u10,0)>0 and cp09(+)=a0j01 and na01(+)=a0j04 "
   '一般以acc0k0為主
   Else
      '收據資料
      'Modified by Morgan 2011/11/2 考慮拆收據情形(應收只要減銷帳不必加退費金額--辜)
      'strSql = "select a0k03 CustNo,a0k04 Addressee,a0k11 Comp,a0k02 DocDate,a0k01 DocNo,a0j01 RecNo,st02 Sales" & _
         ",a0j20,a0j21,nvl(cp16,0)-nvl(cp77,0)+nvl(cp78,0) TAmount, 0 RAmount, 0 VAmount, 0 BAmount" & _
         " From acc0k0,caseprogress,acc0j0,Staff" & _
         " where cp60(+)=a0k01 and a0j01(+)=cp09 and ST01(+)=a0k20 and nvl(a0k09,0)=0" & stConCP & stConK
      'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
      'Modified by Lydia 2019/10/15 +本所案號CaseNo
      'Modify by Amy 2023/08/14 +部門條件stConST
      strSql = "select a0k03 CustNo,a0k04 Addressee,a0k11 Comp,a0k02 DocDate" & _
         ",a0k01 DocNo,a0j01 RecNo,st02 Sales,getcp10desc(cp01,cp10,a0j04) cp10N, na03" & _
         ",nvl(a0j09,0)+nvl(a0j10,0)-x3 TAmount" & _
         ",0 RAmount, 0 VAmount, 0 BAmount" & _
         ",a0k33,a0j22,a0j25,cp01||cp02||cp03||cp04 as CaseNo" & _
         " From (select cp09 x1,a0j13 x2,nvl(sum(a1u07),0)+nvl(sum(a1u09),0) x3" & _
         " from acc0k0,acc0j0,caseprogress,acc1u0 where a0j13(+)=a0k01 and cp09(+)=a0j01" & _
         " and nvl(a0k09,0)=0" & stConCP & stConK & _
         " and a1u02(+)=a0j13 and a1u03(+)=a0j01 group by cp09,a0j13),acc0k0,acc0j0,caseprogress,Staff,nation" & _
         " where cp09(+)=x1 and a0j01(+)=x1 and a0j13(+)=x2 and a0k01(+)=x2 and ST01(+)=a0k20 and na01(+)=a0j04 " & stConST
         
      '收款資料
      'Modified by Morgan 2011/11/2 考慮拆收據情形
      'strSql = strSql & " Union All" & _
         " select a0k03 CustNo, a0k04 as Addressee,a0k11 Comp,a0l02 DocDate,a0k01 DocNo,a0j01 RecNo, st02 Sales" & _
         " ,a0j20,a0j21,0 as TAmount, nvl(a1u04,0)+nvl(a1u05,0) as RAmount, a1u06 as VAmount, 0 as BAmount" & _
         " from acc0k0,acc0j0,acc1u0,acc0l0,staff" & _
         " where a0j13(+)=a0k01" & stConJ & stConK & stConL & _
         " and a1u03(+)=a0j01 and a1u02(+)=a0j13 and a1u01 LIKE 'F%'" & _
         " and a0l01(+)=a1u01 and st01(+)=a0k20"
      'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
      'Modified by Lydia 2019/10/15 +本所案號CaseNo
      'Modify by Amy 2023/08/14 +部門條件stConST
      strSql = strSql & " Union All" & _
         " select a0k03 CustNo, a0k04 as Addressee,a0k11 Comp,a0l02 DocDate,a0k01 DocNo,a0j01 RecNo, st02 Sales" & _
         ",getcp10desc(cp01,cp10,a0j04) cp10N,na03,0 as TAmount, nvl(a1u04,0)+nvl(a1u05,0) as RAmount, a1u06 as VAmount, 0 as BAmount" & _
         ",a0k33,a0j22,a0j25,cp01||cp02||cp03||cp04 as CaseNo" & _
         " from acc0k0,acc0j0,acc1u0,acc0l0,staff,caseprogress,nation" & _
         " where a0j13(+)=a0k01" & stConJ & stConK & stConL & stConST & _
         " and a1u03(+)=a0j01 and a1u02(+)=a0j13 and a1u01 LIKE 'F%'" & _
         " and a0l01(+)=a1u01 and st01(+)=a0k20 and cp09(+)=a0j01 and na01(+)=a0j04 "
         
      '退費資料
      'Modified by Morgan 2011/11/2 考慮拆收據情形　+and a1u02(+)=a0j13
      'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
      'Modified by Lydia 2019/10/15 +本所案號CaseNo
      'Modify by Amy 2023/08/14 +部門條件stConST
      strSql = strSql & " Union All" & _
         " select a0k03 CustNo, a0k04 Addressee,a0k11 Comp,a0s03 DocDate,a0k01 DocNo,a0j01 RecNo,st02 Sales" & _
         ",getcp10desc(cp01,cp10,a0j04) cp10N,na03,0 TAmount,0 RAmount, a1u06 VAmount,nvl(a1u08,0)+nvl(a1u10,0) BAmount" & _
         ",a0k33,a0j22,a0j25,cp01||cp02||cp03||cp04 as CaseNo" & _
         " from acc0k0,acc0j0,acc1u0,acc0s0,staff,caseprogress,nation" & _
         " where a0j13(+)=a0k01" & stConJ & stConK & stConS & stConST & _
         " and a1u03(+)=a0j01 and a1u02(+)=a0j13 and a1u01 LIKE 'I%'" & _
         " and a0s01(+)=a1u01 and st01(+)=a0k20 and nvl(a1u08,0)+nvl(a1u10,0)>0 and cp09(+)=a0j01 and na01(+)=a0j04 "
         
   End If
   '若非北所員工, 只能列印該所資料
   If pub_strUserOffice <> "1" Then
       strSql = strSql & " And ''||ST06='" & pub_strUserOffice & "' "
   End If
   
   '依公司別跳頁
   If Me.Text12.Text = "Y" Then
      strSql = strSql & " order by CustNo asc, Comp asc, Addressee asc, DocDate asc"
   Else
      strSql = strSql & " order by CustNo asc, Addressee asc, DocDate asc"
   End If
   
   'Added by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
   'Modified by Lydia 2019/10/15 比照智權人員-客戶請款明細表的順序: 收據日期a0k02->收據號碼a0k01->收文號
   'strSql = strSql & ",DocNo asc,a0j25 asc"
   strSql = strSql & ",DocNo asc, RecNo asc, a0j25 asc"
   
On Error GoTo ErrHnd

   If adoacc0m0.State = adStateOpen Then adoacc0m0.Close
   With adoacc0m0
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         Erase lngSubTot
         Erase lngTot
         If Text6 <> "" Then
            stPreCustNo = Text6.Text
            'Modify By Sindy 2016/6/8 Text3 ==> cboTitle.Text
            'stPreAddressee = Text3.Text
            stPreAddressee = cboTitle.Text
            '2016/6/8 END
         Else
            stPreCustNo = "" & .Fields("CustNo").Value
            stPreAddressee = "" & .Fields("Addressee").Value
         End If
         m_strA0k11 = "" & .Fields("Comp").Value
         Do While Not .EOF
            If Text6 <> "" Then
               stCustNo = Text6.Text
               'Modify By Sindy 2016/6/8 Text3 ==> cboTitle.Text
               'stAddressee = Text3.Text
               stAddressee = cboTitle.Text
               '2016/6/8 END
            Else
               stCustNo = "" & .Fields("CustNo").Value
               stAddressee = "" & .Fields("Addressee").Value
            End If
            If stCustNo & stAddressee <> stPreCustNo & stPreAddressee Then
               SubSelect stPreCustNo, stPreAddressee, lngSubTot()
               stPreCustNo = stCustNo: stPreAddressee = stAddressee
               For jj = 1 To 5
                  lngTot(jj) = lngTot(jj) + lngSubTot(jj)
               Next
               Erase lngSubTot
            '若依公司別跳頁 公司別不同時
            ElseIf Me.Text12.Text = "Y" And "" & .Fields("Comp").Value <> m_strA0k11 Then
               SubSelect stCustNo, stAddressee, lngSubTot()
               For jj = 1 To 5
                  lngTot(jj) = lngTot(jj) + lngSubTot(jj)
               Next
               Erase lngSubTot
            End If
            m_strA0k11 = "" & .Fields("Comp").Value
            
            'Added by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
            'Modified by Lydia 2019/10/15 +本所案號
            'If stLstDocNo = .Fields("DocNo") And .Fields("a0k33") = "Y" And stLstItem = .Fields("a0j22") Then
            If stLstDocNo = .Fields("DocNo") And .Fields("a0k33") = "Y" And stLstItem = .Fields("a0j22") And stLstCaseNo = "" & .Fields("CaseNo") Then
               bolAddItem = False
            Else
               bolAddItem = True
            End If
            stLstDocNo = "" & .Fields("DocNo")
            stLstItem = "" & .Fields("a0j22")
            stLstCaseNo = "" & .Fields("CaseNo").Value 'Added by Lydia 2019/10/15
            If bolAddItem = True Then
            'end 2011/12/21
            
               adoaccrpt105.AddNew
               adoaccrpt105.Fields("r10501").Value = strUserNum
               adoaccrpt105.Fields("r10502").Value = Counter
               adoaccrpt105.Fields("r10503").Value = stCustNo
               adoaccrpt105.Fields("r10504").Value = stAddressee
               adoaccrpt105.Fields("r10505").Value = .Fields("DocDate").Value
               adoaccrpt105.Fields("r10507").Value = .Fields("DocNo").Value
               adoaccrpt105.Fields("r10508").Value = .Fields("Sales").Value
               adoaccrpt105.Fields("r10510").Value = MidB(CaseNameQuery("" & .Fields("RecNo").Value, 1), 1, 24)
               'Modified by Morgan 2011/12/21
               If .Fields("a0k33") = "Y" Then
                  adoaccrpt105.Fields("r10511").Value = .Fields("a0j22").Value
               Else
                  'Modified by Morgan 2011/12/27 取消 a0j20
                  adoaccrpt105.Fields("r10511").Value = .Fields("cp10N").Value
               End If
               'end 2011/12/21
               
               'Modified by Morgan 2011/12/30 取消 a0j21
               adoaccrpt105.Fields("r10512").Value = "" & .Fields("na03").Value
               
               'Modify by Amy 2018/12/24 資料null 會error ex:X66340 1071222 往來 資料E10025019 1040721 AA0042368 VAmount 為null
               adoaccrpt105.Fields("r10513").Value = Format(Val("" & .Fields("TAmount").Value))
               adoaccrpt105.Fields("r10514").Value = Format(Val("" & .Fields("RAmount").Value))
               adoaccrpt105.Fields("r10515").Value = Format(Val("" & .Fields("VAmount").Value))
               adoaccrpt105.Fields("r10516").Value = Format(Val("" & .Fields("BAmount").Value))
               'end 2018/12/24
               
               lngSubTot(1) = lngSubTot(1) + adoaccrpt105.Fields("r10513").Value
               lngSubTot(2) = lngSubTot(2) + adoaccrpt105.Fields("r10514").Value
               lngSubTot(3) = lngSubTot(3) + adoaccrpt105.Fields("r10515").Value
               lngSubTot(4) = lngSubTot(4) + adoaccrpt105.Fields("r10516").Value
               
               adoaccrpt105.Fields("r10522").Value = "" & .Fields("Comp").Value
               '客戶案件案號
               '2013/5/13 modify by sonia 不列印客戶案件案號改印本所案號
               'If Me.Text13.Text = "Y" Then
               '    adoaccrpt105.Fields("r10523").Value = GetCustCaseNo("" & .Fields("DocNo").Value)
               'Else
               '    adoaccrpt105.Fields("r10523").Value = ""
               'End If
               'Modified by Lydia 2019/10/15 直接抓本所案號
               'adoaccrpt105.Fields("r10523").Value = GetCustCaseNo("" & .Fields("DocNo").Value, Me.Text13.Text)
               adoaccrpt105.Fields("r10523").Value = "" & .Fields("CaseNo")
               '2013/5/13 end
               
            'Added by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
            Else
               Erase dblAddAmt
               
               dblAddAmt(1) = Val(.Fields("TAmount").Value)
               dblAddAmt(2) = Val(.Fields("RAmount").Value)
               dblAddAmt(3) = Val(.Fields("VAmount").Value)
               dblAddAmt(4) = Val(.Fields("BAmount").Value)
               
               adoaccrpt105.Fields("r10513").Value = Format(adoaccrpt105.Fields("r10513").Value + dblAddAmt(1))
               adoaccrpt105.Fields("r10514").Value = Format(adoaccrpt105.Fields("r10514").Value + dblAddAmt(2))
               adoaccrpt105.Fields("r10515").Value = Format(adoaccrpt105.Fields("r10515").Value + dblAddAmt(3))
               adoaccrpt105.Fields("r10516").Value = Format(adoaccrpt105.Fields("r10516").Value + dblAddAmt(4))
               
               lngSubTot(1) = lngSubTot(1) + dblAddAmt(1)
               lngSubTot(2) = lngSubTot(2) + dblAddAmt(2)
               lngSubTot(3) = lngSubTot(3) + dblAddAmt(3)
               lngSubTot(4) = lngSubTot(4) + dblAddAmt(4)
            End If
            'end 2011/12/21
            adoaccrpt105.UpdateBatch
            .MoveNext
         Loop
         '小計
         SubSelect stCustNo, stAddressee, lngSubTot()
         For jj = 1 To 5
            lngTot(jj) = lngTot(jj) + lngSubTot(jj)
         Next
         '合計
         SubSelect stCustNo, stAddressee, lngTot(), 2
      Else
         MsgBox MsgText(28), , MsgText(5)
      End If
   End With
   If adoacc0m0.State = adStateOpen Then adoacc0m0.Close
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   'Resume
End Sub

'Add by Morgan 2005/5/31
'*************************************************
'  列印類別-選擇函證用未收
'
'*************************************************
Private Sub Select5New()
   Dim stVTBx As String, stVTBy As String
   Dim stConK As String, stConCP As String
   Dim stChoice As String
   Dim stCustNo As String, stAddressee As String
   Dim stPreCustNo As String, stPreAddressee As String
   Dim stDDate As String
   Dim bolAddItem As Boolean
   Dim stLstItem As String, stLstDocNo As String 'Added by Morgan 2011/12/21 記錄案件性質、收據號碼
   Dim dblAddAmt(4) As Double 'Added by Morgan 2011/12/21
   Dim stLstCaseNo As String 'Added by Lydia 2019/10/15 記錄本所案號
   
   '語法選擇 預設1 acc0k0為主
   stChoice = "1"
   strSql = ""
   stConCP = ""
   stConK = ""
   '系統類別
   If txtSys.Text <> "" Then
      stConCP = stConCP & " and cp01 in (" & GetSysSQL(txtSys.Text) & ")"
   End If
   If Text1 = "X" Then Text1 = MsgText(601)
   If Text2 = "X" Then Text2 = MsgText(601)
   '客戶代號
   If Text1 <> MsgText(601) And Text1 <> "X" Then
      stConK = stConK & " and a0k03 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) And Text2 <> "X" Then
      stConK = stConK & " and a0k03 <= '" & Text2 & "'"
   End If
   '收據抬頭
   'Modify By Sindy 2016/6/8 Text3 ==> cboTitle.Text
   'If Text3 <> MsgText(601) Then
   If cboTitle.Text <> MsgText(601) Then
      '2011/10/20 MODIFY BY SONIA E10023515
      'stConK = stConK & " and instr(a0k04, '" & Text3 & "') > 0"
      'stConK = stConK & " and instr(UPPER(a0k04), UPPER('" & Text3 & "')) > 0"
      stConK = stConK & " and instr(UPPER(a0k04), UPPER('" & cboTitle.Text & "')) > 0"
   End If
   '智權人員
   If Text4 <> MsgText(601) Then
      stConK = stConK & " and a0k20 = '" & Text4 & "'"
   End If
   '作業日期
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      stConK = stConK & " and a0k02 >= " & Val(FCDate(MaskEdBox1.Text))
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      stConK = stConK & " and a0k02 <= " & Val(FCDate(MaskEdBox2.Text))
      stDDate = Val(FCDate(MaskEdBox2.Text))
   End If
   '本所案號
   If Text7 <> "" And Text8 <> "" Then
      stConCP = stConCP & " and cp01='" & Text7 & "' and cp02='" & Text8 & "' and cp03='" & Text9 & "' and cp04='" & Text10 & "'"
      stChoice = "2"
   End If
   '公司別
   If Me.Text14.Text <> "" Then
       stConK = stConK & " and A0K11 >= '" & Me.Text14.Text & "'"
   End If
   If Me.Text15.Text <> "" Then
       stConK = stConK & " and A0K11 <= '" & Me.Text15.Text & "'"
   End If
   
   'Modified by Morgan 2011/12/21 +a0k33,a0j22,a0j25
   'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
   strSql = "select a0k01 DocNo, a0k04 Addressee, a0k03 CustNo, a0k02 DocDate, st02 Sales" & _
      ",a0j01 RecNo, getcp10desc(cp01,cp10,a0j04) cp10N, na03,nvl(a0j09,0)-nvl(a1u07,0) Fee1,nvl(A0j10,0)-nvl(a1u09,0) Fee2" & _
      ",nvl(a1u04,0)+nvl(a1u05,0)-nvl(a1u08,0)-nvl(a1u10,0) Fee3,a0k11 Comp" & _
      ",substr(VxC0,8,7) Amount, substr(VxC0,1,7) DueDate,substr(VxC0,15) ColNo,a0j07,a0k33,a0j22,a0j25"
      
   '排除 未收金額=0 and ( 無票據 or 票據到期日<作業迄日者)
   '有本所號時以caseprogress為主
   If stChoice = "2" Then
      '票據資料
      'Modified by Morgan 2011/11/2 考慮拆收據情形
      'stVTBx = "select cp09 VxC1,max(lpad(nvl(a0e10,0),7,'0')||lpad(nvl(a0e11,0),7,'0')||a0e03) VxC0" & _
         " From caseprogress, acc0m0, acc0e0" & _
         " where a0m02(+)=cp60 and a0e03(+)=a0m01" & stConCP & _
         " group by cp09"
      stVTBx = "select a0j01 VxC1,a0j13 VxC2,max(lpad(nvl(a0e10,0),7,'0')||lpad(nvl(a0e11,0),7,'0')||a0e03) VxC0" & _
         " From caseprogress,acc0j0, acc0m0, acc0e0" & _
         " where a0j01(+)=cp09 and a0m02(+)=a0j13 and a0e03(+)=a0m01" & stConCP & _
         " group by a0j01,a0j13"
         
      '銷帳
      'Modified by Morgan 2011/11/2 考慮拆收據情形
      'stVTBy = "select a1u03,sum(a1u04) a1u04,sum(a1u05) a1u05,sum(a1u07) a1u07" & _
         ",sum(a1u08) a1u08,sum(a1u09) a1u09,sum(a1u10) a1u10" & _
         " From caseprogress, acc1u0" & _
         " where a1u02(+)=cp60" & stConCP & _
         " group by a1u03"
      stVTBy = "select a1u03,a1u02,sum(a1u04) a1u04,sum(a1u05) a1u05,sum(a1u07) a1u07" & _
         ",sum(a1u08) a1u08,sum(a1u09) a1u09,sum(a1u10) a1u10" & _
         " From caseprogress, acc1u0" & _
         " where a1u03(+)=cp09" & stConCP & _
         " group by a1u03,a1u02"
         
   '一般以acc0k0為主
   Else
      '票據資料
      'Modified by Morgan 2011/11/2 考慮拆收據情形
      'stVTBx = "select cp09 VxC1,cp60 VxC2,max(lpad(nvl(a0e10,0),7,'0')||lpad(nvl(a0e11,0),7,'0')||a0e03) VxC0" & _
         " From acc0k0,caseprogress, acc0m0, acc0e0" & _
         " where cp60(+)=a0k01 and a0m02(+)=a0k01 and a0e03(+)=a0m01" & stConCP & stConK & _
         " group by cp09,cp60"
      stVTBx = "select a0j01 VxC1,a0j13 VxC2,max(lpad(nvl(a0e10,0),7,'0')||lpad(nvl(a0e11,0),7,'0')||a0e03) VxC0" & _
         " From acc0k0,acc0j0,caseprogress, acc0m0, acc0e0" & _
         " where a0j13(+)=a0k01 and cp09(+)=a0j01 and a0m02(+)=a0k01 and a0e03(+)=a0m01" & stConCP & stConK & _
         " group by a0j01,a0j13"
      
      '銷帳
      'Modified by Morgan 2011/11/2 考慮拆收據情形
      'stVTBy = "select a1u03,sum(a1u04) a1u04,sum(a1u05) a1u05,sum(a1u07) a1u07" & _
         ",sum(a1u08) a1u08,sum(a1u09) a1u09,sum(a1u10) a1u10" & _
         " From acc0k0, acc1u0,caseprogress" & _
         " where a1u02(+)=a0k01 and cp60(+)=a1u03" & stConCP & stConK & _
         " group by a1u03"
      stVTBy = "select a1u03,a1u02,sum(a1u04) a1u04,sum(a1u05) a1u05,sum(a1u07) a1u07" & _
         ",sum(a1u08) a1u08,sum(a1u09) a1u09,sum(a1u10) a1u10" & _
         " From acc0k0, acc1u0,caseprogress" & _
         " where a1u02(+)=a0k01 and cp09(+)=a1u03" & stConCP & stConK & _
         " group by a1u03,a1u02"
         
   End If
   'Modified by Morgan 2011/11/2 考慮拆收據情形
   'strSql = strSql & _
      " from (" & stVTBx & ") Vx,(" & stVTBy & ") Vy,caseprogress,acc0j0,acc0k0, staff" & _
      " where a1u03(+)=VxC1 and a1u03(+)=VxC1 and cp09(+)=VxC1 and a0j01(+)=VxC1" & _
      " and a0k01(+)=VxC2 and st01(+)=a0k20" & stConK & _
      " and not (cp79=0 and ( VxC0 is null or to_number(substr(VxC0,1,7))<" & stDDate & "))"
   'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
   '2012/1/12 modify by sonia 辜說函證用不包含未列印收據故加 and a0k32 is null條件
   'Modified by Lydia 2019/10/15 +本所案號CaseNo
   'Modified by Lydia 2025/06/10 a0k32 is null 改用函數判斷：geta0k32type(a0k01)='1'
   strSql = strSql & _
      ",cp01||cp02||cp03||cp04 as CaseNo" & _
      " from (" & stVTBx & ") Vx,(" & stVTBy & ") Vy,caseprogress,acc0j0,acc0k0, staff,nation" & _
      " where a1u03(+)=VxC1 and a1u02(+)=VxC2 and cp09(+)=VxC1 and a0j01(+)=VxC1 and a0j13(+)=VxC2" & _
      " and a0k01(+)=VxC2 and st01(+)=a0k20 and geta0k32type(a0k01)='1'" & stConK & _
      " and not (cp79=0 and ( VxC0 is null or to_number(substr(VxC0,1,7))<" & stDDate & ")) and nvl(a0j09,0)-nvl(a1u07,0) +nvl(A0j10,0)-nvl(a1u09,0)>0 and na01(+)=a0j04 "
      
   '若非北所員工, 只能列印該所資料
   If pub_strUserOffice <> "1" Then
       strSql = strSql & " And ''||ST06='" & pub_strUserOffice & "'"
   End If
   
   'Add by Amy 2023/08/14 +部門
   If TxtDeptS <> MsgText(601) Then
      strSql = strSql & " And ST15>='" & TxtDeptS & "' "
   End If
   If TxtDeptE <> MsgText(601) Then
      strSql = strSql & " And ST15<='" & TxtDeptE & "' "
   End If
   'end 2023/08/14
   
   '依公司別跳頁
   'Modified by Morgan 2011/12/27 取消 a0j20
   If Me.Text12.Text = "Y" Then
      'Modified by Lydia 2019/10/15 比照智權人員-客戶請款明細表的順序: 收據日期a0k02->收據號碼a0k01->收文號
      'strSql = strSql & " order by CustNo asc, Comp asc, Addressee asc, DocDate asc,DocNo asc,cp10N asc"
      strSql = strSql & " order by CustNo asc, Comp asc, Addressee asc, DocDate asc"
   Else
      'Modified by Lydia 2019/10/15 比照智權人員-客戶請款明細表的順序: 收據日期a0k02->收據號碼a0k01->收文號
      'strSql = strSql & " order by CustNo asc, Addressee asc, DocDate asc,DocNo asc,cp10N asc"
      strSql = strSql & " order by CustNo asc, Addressee asc, DocDate asc"
   End If
   
   'Added by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
   'Modified by Lydia 2019/10/15 比照智權人員-客戶請款明細表的順序: 收據日期a0k02->收據號碼a0k01->收文號
   'strSql = strSql & ",DocNo asc,a0j25 asc"
   strSql = strSql & ",DocNo asc, RecNo asc, a0j25 asc"
   
On Error GoTo ErrHnd

   If adoacc0m0.State = adStateOpen Then adoacc0m0.Close
   With adoacc0m0
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         Erase lngSubTot
         Erase lngTot
         If Text6 <> "" Then
            stPreCustNo = Text6.Text
            'Modify By Sindy 2016/6/8 Text3 ==> cboTitle.Text
            'stPreAddressee = Text3.Text
            stPreAddressee = cboTitle.Text
            '2016/6/8 END
         Else
            stPreCustNo = "" & .Fields("CustNo").Value
            stPreAddressee = "" & .Fields("Addressee").Value
         End If
         m_strA0k11 = "" & .Fields("Comp").Value
         Do While Not .EOF
            If Text6 <> "" Then
               stCustNo = Text6.Text
               'Modify By Sindy 2016/6/8 Text3 ==> cboTitle.Text
               'stAddressee = Text3.Text
               stAddressee = cboTitle.Text
               '2016/6/8 END
            Else
               stCustNo = "" & .Fields("CustNo").Value
               stAddressee = "" & .Fields("Addressee").Value
            End If
            If stCustNo & stAddressee <> stPreCustNo & stPreAddressee Then
               SubSelect stPreCustNo, stPreAddressee, lngSubTot()
               stPreCustNo = stCustNo: stPreAddressee = stAddressee
               For jj = 1 To 5
                  lngTot(jj) = lngTot(jj) + lngSubTot(jj)
               Next
               Erase lngSubTot
            '若依公司別跳頁 公司別不同時
            ElseIf Me.Text12.Text = "Y" And "" & .Fields("Comp").Value <> m_strA0k11 Then
               SubSelect stCustNo, stAddressee, lngSubTot()
               For jj = 1 To 5
                  lngTot(jj) = lngTot(jj) + lngSubTot(jj)
               Next
               Erase lngSubTot
            End If
            m_strA0k11 = "" & .Fields("Comp").Value
            
            'Added by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
            'Modified by Lydia 2019/10/15 +本所案號
            'If stLstDocNo = .Fields("DocNo") And .Fields("a0k33") = "Y" And stLstItem = .Fields("a0j22") Then
            If stLstDocNo = .Fields("DocNo") And .Fields("a0k33") = "Y" And stLstItem = .Fields("a0j22") And stLstCaseNo = "" & .Fields("CaseNo") Then
               bolAddItem = False
            Else
               bolAddItem = True
            End If
            stLstDocNo = "" & .Fields("DocNo")
            stLstItem = "" & .Fields("a0j22")
            stLstCaseNo = "" & .Fields("CaseNo").Value 'Added by Lydia 2019/10/15
            If bolAddItem = True Then
            'end 2011/12/21
               
               adoaccrpt105.AddNew
               adoaccrpt105.Fields("r10501").Value = strUserNum
               adoaccrpt105.Fields("r10502").Value = Counter
               adoaccrpt105.Fields("r10503").Value = stCustNo
               adoaccrpt105.Fields("r10504").Value = stAddressee
               adoaccrpt105.Fields("r10505").Value = .Fields("DocDate").Value
               adoaccrpt105.Fields("r10507").Value = .Fields("DocNo").Value
               adoaccrpt105.Fields("r10508").Value = .Fields("Sales").Value
               adoaccrpt105.Fields("r10510").Value = MidB(CaseNameQuery(.Fields("RecNo").Value, 1), 1, 24)
               'Modified by Morgan 2011/12/21
               If .Fields("a0k33") = "Y" Then
                  adoaccrpt105.Fields("r10511").Value = .Fields("a0j22").Value
               Else
                  'Modified by Morgan 2011/12/27 取消 a0j20
                  adoaccrpt105.Fields("r10511").Value = .Fields("cp10N").Value
               End If
               'end 2011/12/21
               
               'Modified by Morgan 2011/12/30 取消 a0j21
               adoaccrpt105.Fields("r10512").Value = .Fields("na03").Value
               
               '是否合併
               If "" & .Fields("a0j07").Value = "Y" Then
                  '服務費
                  adoaccrpt105.Fields("r10513").Value = Format(.Fields("Fee1").Value + .Fields("Fee2").Value)
                  '規費
                  adoaccrpt105.Fields("r10514").Value = 0
               Else
                  '服務費
                  adoaccrpt105.Fields("r10513").Value = Format(.Fields("Fee1").Value)
                  '規費
                  adoaccrpt105.Fields("r10514").Value = Format(.Fields("Fee2").Value)
               End If
               
               '應收
               adoaccrpt105.Fields("r10515").Value = Format(Val(adoaccrpt105.Fields("r10513").Value) + Val(adoaccrpt105.Fields("r10514").Value))
               '已收
               adoaccrpt105.Fields("r10516").Value = Format(.Fields("Fee3").Value)
               '未收
               adoaccrpt105.Fields("r10517").Value = Format(Val(adoaccrpt105.Fields("r10515").Value) - Val(adoaccrpt105.Fields("r10516").Value))
               
               lngSubTot(1) = lngSubTot(1) + adoaccrpt105.Fields("r10513").Value
               lngSubTot(2) = lngSubTot(2) + adoaccrpt105.Fields("r10514").Value
               lngSubTot(3) = lngSubTot(3) + Val(adoaccrpt105.Fields("r10515").Value)
               lngSubTot(4) = lngSubTot(4) + adoaccrpt105.Fields("r10516").Value
               lngSubTot(5) = lngSubTot(5) + adoaccrpt105.Fields("r10517").Value
               
               adoaccrpt105.Fields("r10519").Value = Format("" & .Fields("Amount").Value, "#")
               adoaccrpt105.Fields("r10520").Value = Format("" & .Fields("DueDate").Value, "#")
               adoaccrpt105.Fields("r10521").Value = "" & .Fields("ColNo").Value
               adoaccrpt105.Fields("r10522").Value = "" & .Fields("Comp").Value
               '客戶案件案號
               '2013/5/13 modify by sonia 不列印客戶案件案號改印本所案號
               'If Me.Text13.Text = "Y" Then
               '    adoaccrpt105.Fields("r10523").Value = GetCustCaseNo("" & .Fields("DocNo").Value)
               'Else
               '    adoaccrpt105.Fields("r10523").Value = ""
               'End If
               'Modified by Lydia 2019/10/15 直接抓本所案號
               'adoaccrpt105.Fields("r10523").Value = GetCustCaseNo("" & .Fields("DocNo").Value, Me.Text13.Text)
               adoaccrpt105.Fields("r10523").Value = "" & .Fields("CaseNo")
               '2013/5/13 end
               
            'Added by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
            Else
               Erase dblAddAmt
               
               '是否合併
               If "" & .Fields("a0j07").Value = "Y" Then
                  '服務費
                  dblAddAmt(1) = .Fields("Fee1").Value + .Fields("Fee2").Value
                  '規費
                  dblAddAmt(2) = 0
               Else
                  '服務費
                  dblAddAmt(1) = .Fields("Fee1").Value
                  '規費
                  dblAddAmt(2) = .Fields("Fee2").Value
               End If
               dblAddAmt(3) = .Fields("Fee3").Value
               
               adoaccrpt105.Fields("r10513").Value = Format(adoaccrpt105.Fields("r10513").Value + dblAddAmt(1))
               adoaccrpt105.Fields("r10514").Value = Format(adoaccrpt105.Fields("r10514").Value + dblAddAmt(2))
               adoaccrpt105.Fields("r10515").Value = Format(Val(adoaccrpt105.Fields("r10513").Value) + Val(adoaccrpt105.Fields("r10514").Value))
               adoaccrpt105.Fields("r10516").Value = Format(adoaccrpt105.Fields("r10516").Value + dblAddAmt(3))
               adoaccrpt105.Fields("r10517").Value = Format(Val(adoaccrpt105.Fields("r10515").Value) - Val(adoaccrpt105.Fields("r10516").Value))
               
               lngSubTot(1) = lngSubTot(1) + dblAddAmt(1)
               lngSubTot(2) = lngSubTot(2) + dblAddAmt(2)
               lngSubTot(3) = lngSubTot(3) + dblAddAmt(1) + dblAddAmt(2)
               lngSubTot(4) = lngSubTot(4) + dblAddAmt(3)
               lngSubTot(5) = lngSubTot(5) + dblAddAmt(1) + dblAddAmt(2) - dblAddAmt(3)
            End If
            'end 2011/12/21

            adoaccrpt105.UpdateBatch
            .MoveNext
         Loop
         '小計
         SubSelect stCustNo, stAddressee, lngSubTot()
         For jj = 1 To 5
            lngTot(jj) = lngTot(jj) + lngSubTot(jj)
         Next
         '合計
         SubSelect stCustNo, stAddressee, lngTot(), 2
      Else
         MsgBox MsgText(28), , MsgText(5)
      End If
   End With
   If adoacc0m0.State = adStateOpen Then adoacc0m0.Close
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   'Resume
End Sub
'Add by Morgan 2005/5/30
'*************************************************
'  列印類別-選擇收回之小計計算
'
'*************************************************
Private Sub SubSelect(ByVal p_stCustNo As String, ByVal p_stAddressee As String, ByRef p_lngSubTot() As Long, Optional p_iMode As Integer = 1, Optional p_stSelect As String = "2")

   adoaccrpt105.AddNew
   adoaccrpt105.Fields("r10501").Value = strUserNum
   adoaccrpt105.Fields("r10502").Value = Counter
   adoaccrpt105.Fields("r10503").Value = p_stCustNo
   adoaccrpt105.Fields("r10504").Value = p_stAddressee
   adoaccrpt105.Fields("r10522").Value = m_strA0k11
   PaintLine ReportSum(4)
   adoaccrpt105.UpdateBatch
   
   adoaccrpt105.AddNew
   adoaccrpt105.Fields("r10501").Value = strUserNum
   adoaccrpt105.Fields("r10502").Value = Counter
   adoaccrpt105.Fields("r10503").Value = p_stCustNo
   adoaccrpt105.Fields("r10504").Value = p_stAddressee
   If p_iMode = 2 Then
      adoaccrpt105.Fields("r10510").Value = ReportSum(25)
   Else
      adoaccrpt105.Fields("r10510").Value = ReportSum(24)
   End If
   adoaccrpt105.Fields("r10513").Value = p_lngSubTot(1)
   adoaccrpt105.Fields("r10514").Value = p_lngSubTot(2)
   adoaccrpt105.Fields("r10515").Value = p_lngSubTot(3)
   adoaccrpt105.Fields("r10516").Value = p_lngSubTot(4)
   adoaccrpt105.Fields("r10517").Value = p_lngSubTot(5)
   adoaccrpt105.Fields("r10522").Value = m_strA0k11
   adoaccrpt105.UpdateBatch
   '合計線
   If p_iMode = 2 Then
      adoaccrpt105.AddNew
      adoaccrpt105.Fields("r10501").Value = strUserNum
      adoaccrpt105.Fields("r10502").Value = Counter
      adoaccrpt105.Fields("r10503").Value = p_stCustNo
      adoaccrpt105.Fields("r10504").Value = p_stAddressee
      adoaccrpt105.Fields("r10522").Value = m_strA0k11
      PaintLine ReportSum(8)
      adoaccrpt105.UpdateBatch
   End If
End Sub
