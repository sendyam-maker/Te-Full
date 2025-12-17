VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc1460 
   AutoRedraw      =   -1  'True
   Caption         =   "智權人員帳款明細表"
   ClientHeight    =   3972
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8664
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3972
   ScaleWidth      =   8664
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
      Left            =   5190
      TabIndex        =   6
      Top             =   810
      Width           =   945
   End
   Begin VB.CheckBox Check5 
      Caption         =   "是否列出申請案號"
      Height          =   195
      Left            =   2280
      TabIndex        =   45
      Top             =   2310
      Width           =   2500
   End
   Begin VB.OptionButton Option1 
      Caption         =   "提供客戶用 (服務費、規費合併)"
      Height          =   180
      Index           =   0
      Left            =   5280
      TabIndex        =   44
      Top             =   2310
      Value           =   -1  'True
      Width           =   3000
   End
   Begin VB.OptionButton Option1 
      Caption         =   "提供智權用 (服務費、規費分開)"
      Height          =   180
      Index           =   1
      Left            =   5280
      TabIndex        =   43
      Top             =   2550
      Width           =   3000
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
      Height          =   315
      Left            =   5280
      MaxLength       =   3
      TabIndex        =   11
      Top             =   1530
      Width           =   615
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
      Height          =   315
      Left            =   5880
      MaxLength       =   6
      TabIndex        =   12
      Top             =   1530
      Width           =   975
   End
   Begin VB.TextBox Text11 
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
      Left            =   6840
      MaxLength       =   1
      TabIndex        =   13
      Top             =   1530
      Width           =   375
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
      Height          =   315
      Left            =   7200
      MaxLength       =   2
      TabIndex        =   14
      Top             =   1530
      Width           =   495
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
      ItemData        =   "Frmacc1460.frx":0000
      Left            =   960
      List            =   "Frmacc1460.frx":0002
      TabIndex        =   7
      Top             =   1200
      Width           =   6720
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
      Left            =   7800
      TabIndex        =   8
      Top             =   1200
      Width           =   675
   End
   Begin VB.CheckBox Check4 
      Caption         =   "是否列出客戶案件案號欄"
      Height          =   195
      Left            =   2280
      TabIndex        =   22
      Top             =   2550
      Width           =   2500
   End
   Begin VB.CheckBox Check3 
      Caption         =   "是否含預定收款日未到期者"
      Height          =   195
      Left            =   2280
      TabIndex        =   20
      Top             =   2790
      Visible         =   0   'False
      Width           =   2500
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
      Height          =   315
      Left            =   2670
      MaxLength       =   9
      TabIndex        =   10
      Top             =   1530
      Width           =   1425
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
      Height          =   315
      Left            =   960
      MaxLength       =   9
      TabIndex        =   9
      Top             =   1530
      Width           =   1425
   End
   Begin VB.CheckBox Check2 
      Caption         =   "僅列印巳送件者"
      Height          =   195
      Left            =   150
      TabIndex        =   21
      Top             =   2550
      Width           =   2000
   End
   Begin VB.TextBox Text5 
      Height          =   300
      Left            =   6930
      TabIndex        =   4
      Top             =   450
      Width           =   1425
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   5190
      TabIndex        =   3
      Top             =   450
      Width           =   1425
   End
   Begin VB.CheckBox Check1 
      Caption         =   "僅列印有備註者"
      Height          =   195
      Left            =   150
      TabIndex        =   19
      Top             =   2310
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生Excel檔(&P)"
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
      Left            =   1860
      Style           =   1  '圖片外觀
      TabIndex        =   23
      Top             =   3270
      Width           =   4692
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
      Height          =   315
      Left            =   960
      MaxLength       =   1
      TabIndex        =   5
      Top             =   810
      Width           =   612
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
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   945
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   1425
      _ExtentX        =   2519
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
      Left            =   2670
      TabIndex        =   2
      Top             =   480
      Width           =   1425
      _ExtentX        =   2498
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
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   300
      Left            =   960
      TabIndex        =   15
      Top             =   1890
      Width           =   1425
      _ExtentX        =   2519
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
      Left            =   2670
      TabIndex        =   16
      Top             =   1890
      Width           =   1425
      _ExtentX        =   2498
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
   Begin MSMask.MaskEdBox MaskEdBox5 
      Height          =   300
      Left            =   5280
      TabIndex        =   17
      Top             =   1890
      Width           =   1425
      _ExtentX        =   2519
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
   Begin MSMask.MaskEdBox MaskEdBox6 
      Height          =   300
      Left            =   6960
      TabIndex        =   18
      Top             =   1890
      Width           =   1425
      _ExtentX        =   2519
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
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "若是提供客戶 須將 台灣 以外的申請國家  服務及規費合併"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   42
      Top             =   3000
      Width           =   5505
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "請留意報表用途"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   870
      TabIndex        =   41
      Top             =   2850
      Width           =   1800
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "本所案號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4290
      TabIndex        =   40
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label20 
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   39
      Top             =   1230
      Width           =   975
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4440
      TabIndex        =   38
      Top             =   810
      Width           =   630
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "發文日"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4290
      TabIndex        =   37
      Top             =   1920
      Width           =   975
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
      Left            =   6780
      TabIndex        =   36
      Top             =   1890
      Width           =   255
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "收文日"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   35
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label10 
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
      Left            =   2460
      TabIndex        =   34
      Top             =   1890
      Width           =   255
   End
   Begin VB.Label Label9 
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
      Left            =   2460
      TabIndex        =   33
      Top             =   1530
      Width           =   255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "客戶代號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   32
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "收據號碼"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4290
      TabIndex        =   31
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label6 
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
      Left            =   6720
      TabIndex        =   30
      Top             =   450
      Width           =   255
   End
   Begin VB.Label lblSalesName 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "智權人員名稱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2010
      TabIndex        =   29
      Top             =   180
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   1590
      Top             =   3000
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "(1.未收 2.收回 3.往來)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1650
      TabIndex        =   28
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "列印類別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   27
      Top             =   810
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
      Left            =   2460
      TabIndex        =   26
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "帳款日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   25
      Top             =   510
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   24
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc1460"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
Option Explicit

'Modify by Amy 2017/08/01 原public 若其他有用到相同名稱會造成error
Dim adoacc0k0 As New ADODB.Recordset
Dim adoacc0l0 As New ADODB.Recordset
Dim adocaseprogress As New ADODB.Recordset
Dim adoaccsum As New ADODB.Recordset
Dim adoaccrpt106 As New ADODB.Recordset
'end 2017/08/01
Dim strSalesNo As String
Dim strSameName As String
Dim lngCounter As Long
Dim lngAmount As Long
Dim strAmount As String
Dim intLength As Integer
Dim intCounter As Integer
Dim intPage As Integer
'Add By Cheng 2003/05/20
Const m_dblLeftDiff As Double = 500
'Add By Cheng 2004/02/05
Dim m_strMaxComp As String '最大的公司別
'Add by Morgan 2005/5/13
Dim m_strSalesList As String '智權人員清單
Dim xlsAnnuity As New Excel.Application
Dim wksAnnuity As New Worksheet
'Add by Amy 2015/03/03
Dim intField As Integer '起始欄位
Dim strFieldN(), intWidth()  '欄位名稱/大小
Dim intXlsSheet As Integer '工作表
Dim strText As String
Dim strFileName As String, bolHasData As Boolean 'Add by Amy 2015/05/14
Dim strOldArea As String 'Add by Amy 2016/04/27 業務區
Dim strWkName As String 'Add by Amy 2017/09/25 for 2010 工作表名稱為中文
'Add by Amy 2020/04/27
Dim rsA As New ADODB.Recordset
Dim i As Integer
Dim strAppNo As String 'Add by Amy 2020/05/28
Dim intDelF As Integer 'Add by Amy 2020/06/30 欄位全顯示,若未勾「客戶案件案號」及「申請案號」欄,刪申請案號欄時會刪錯
Dim stTpCmp As String, intTitleR_Fix As Integer, intTitleR As Integer 'Add by Amy 2022/07/04
Dim strF As String 'Add by Amy 2022/07/20

'Mark by Amy 2020/07/08 不用下拉
''Add by Amy 2020/04/27
'Private Sub CboComp_GotFocus()
'    TextInverse CboComp
'End Sub
'
'Private Sub CboComp_KeyPress(KeyAscii As Integer)
'    KeyAscii = UpperCase(KeyAscii)
'End Sub
'
'Private Sub CboComp_Validate(Cancel As Boolean)
'    Dim strCmp As String
'
'    strCmp = CboComp
'    If InStr(strCmp, "　") > 0 Then
'        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
'    End If
'    If InStr(GetBookKeepCmp, strCmp) = 0 Then
'        MsgBox Label14 & MsgText(63), , MsgText(5)
'        Cancel = True
'        CboComp.SetFocus
'        Exit Sub
'    ElseIf Len(Trim(CboComp)) = 1 Then
'        CboComp = Trim(strCmp) & "　" & A0802Query(strCmp)
'    End If
'End Sub
''end 2020/04/27
'end 2020/07/08

'Add By Sindy 2016/6/8
Private Sub cboTitle_Click()
   If cboTitle.ListIndex > 0 Then
      If Text7.Text = "" Then
         Text7.Text = Right(cboTitle.Text, 9)
      ElseIf Text8.Text = "" Then
         Text8.Text = Right(cboTitle.Text, 9)
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
   If Text7 <> "" Or Text8 <> "" Or cboTitle.ListCount > 0 Then
      Text7 = "": Text8 = ""
      Text1 = "": lblSalesName = ""
      cboTitle.Clear
   End If
End Sub
Private Sub cboTitle_Validate(Cancel As Boolean)
   If CheckLen(Label8, cboTitle, 100) = MsgText(603) Then
      Cancel = True
      Exit Sub
   End If
   '切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub
'2016/6/8 END

'Add By Sindy 2016/6/8
Private Sub cmdLikeSearch_Click()
   If cboTitle.Text = "" Then
      MsgBox "請輸入收據抬頭！", vbCritical
   Else
      'Modify by Amy 2024/10/08 bug-帶錯欄位
      'PUB_AddItem2CboTitle cboTitle, Text1, Text2, "", True
      PUB_AddItem2CboTitle cboTitle, Text7, Text8, "", True
   End If
End Sub

'【ACCRPT106 欄位說明】
'r10601 UserID
'r10602 序號　Counter
'r10603 公司別
'r10604 員工部門
'r10605 員工姓名
'r10606 客戶編號
'r10607 收據抬頭
'r10608 收據日期 / 收款日期
'r10609 收款單號
'r10610 收據號碼
'r10611 本所案號
'r10612 案件性質名稱
'r10613 申請國家名稱
'r10614 應收金額: 應收金額 -銷帳 / 本次服務費 + 本次規費
'r10615 已收金額：(本次服務費+本次規費)-(本次退費服務費-本次退費規費)
'r10616 未收金額
'r10617 本次規費
'r10618 案件規費餘額
'r10619 備註
'r10620 點數
'r10621 發文日
'r10622 預定收款日   '2012/5/2 add by sonia
'r10623 A0K32.收據暫不列印 'Add By Sindy 2013/6/14
'r10624 客戶案件案號 'Add By Sindy 2013/12/5
'r10625 應收服務費 / 本次服務費 'Add by Amy 2016/09/19 本次服務費 原寫入r10616
'r10626 扣繳 'Add by Amy 2016/09/19

Private Sub Command1_Click()
   Dim bolCancel As Boolean 'Add by Amy 2020/04/27
   Dim stMsg As String 'Add by Amy 2022/07/04
   'Mark by Amy 2020/03/12
   'Add By Sindy 2013/12/5
'   If Text3 <> "1" And Text3 <> "2" Then
'      MsgBox MsgText(146), , MsgText(5)
'      Text3.SetFocus
'      Exit Sub
'   End If
   '2013/12/5 END
   'Mark by Amy 2020/07/08 不用下拉
'   'Modify by Amy 2020/04/27 公司別改下拉 原:Text2
'   Call CboComp_Validate(bolCancel)
'   If bolCancel = True Then
'      cboComp.SetFocus
'      Exit Sub
'   End If
'   'end 2020/04/27
   'end 2020/07/08
   stTpCmp = "": intTitleR_Fix = 1: intTitleR = 1 'Add by Amy 2022/07/04
   If FormCheck(stMsg) = False Then
      If stMsg = MsgText(601) Then
        stMsg = MsgText(181)
      End If
      MsgBox stMsg, , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   bolHasData = False 'Add by Amy 2015/05/14
   Accrpt106Delete
   ProduceData
   Select Case Text3
        Case "1" '未收
            'PrintReport1
            PrintExcel1 'Add By Sindy 2010/5/20
        Case "2" '收回
            'PrintReport2
            PrintExcel2 'Add By Sindy 2010/5/20
        'Add by Amy 2020/04/27 往來
        Case "3"
            PrintExcel3
    End Select
    
   'Add by Morgan 2005/5/13
   '沒下智權人員條件時印智權人員清單
   'Modify by Morgan 2007/10/1 智權人員範圍改成一個
   'If Text1 = "" And Text2 = "" And m_strSalesList <> "" Then
'   If Text1 = "" And m_strSalesList <> "" Then
'      PrintSalesList
'   End If
   Screen.MousePointer = vbDefault
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
End Sub

'Private Sub PrintSalesList()
'Dim ii As Integer, iIdx As Integer, bolContinue As Boolean, stNames() As String
'Dim iY As Integer, iTabWidth As Integer, iPos As Integer
'
'   stNames = Split(m_strSalesList, vbCrLf)
'   bolContinue = True
'   iIdx = 0
'   intPage = 0
'   Do While bolContinue
'      intPage = intPage + 1
'      Printer.FontSize = 16
'      Printer.CurrentX = 4000 - m_dblLeftDiff
'      Printer.CurrentY = 1000
'      Printer.Print ReportTitle(106)
'      Printer.CurrentX = 8000 - m_dblLeftDiff
'      Printer.CurrentY = 1000
'      Printer.Print IIf(Text3 = "1", "(未收)", "(已收)")
'      Printer.FontSize = 10
'      Printer.CurrentX = 4500 - m_dblLeftDiff
'      Printer.CurrentY = 1800
'      Printer.Print "帳款日期: "
'      Printer.CurrentX = 5700 - m_dblLeftDiff
'      Printer.CurrentY = 1800
'      Printer.Print MaskEdBox1.Text & " ~ " & MaskEdBox2.Text
'      Printer.CurrentX = 500 - m_dblLeftDiff
'      Printer.CurrentY = 2100
'      Printer.Print "列印人員: "
'      Printer.CurrentX = 1800 - m_dblLeftDiff
'      Printer.CurrentY = 2100
'      Printer.Print StaffQuery(strUserNum)
'      Printer.CurrentX = 10000 - 200 - m_dblLeftDiff
'      Printer.CurrentY = 2100
'      Printer.Print "列印日期: "
'      Printer.CurrentX = 11300 - 400 - m_dblLeftDiff
'      Printer.CurrentY = 2100
'      Printer.Print CFDate(ACDate(ServerDate))
'      Printer.CurrentX = 500 - m_dblLeftDiff
'      Printer.CurrentY = 2400
'      Printer.Print "智權人員繳回清單"
'      Printer.CurrentX = 10000 - 200 - m_dblLeftDiff
'      Printer.CurrentY = 2400
'      Printer.Print "頁　　次: "
'      Printer.CurrentX = 11300 - 400 - m_dblLeftDiff
'      Printer.CurrentY = 2400
'      Printer.Print intPage
'      Printer.CurrentX = 500 - m_dblLeftDiff
'      Printer.CurrentY = 2700
'      Printer.FontSize = 12
'      Printer.Print "智權人員　　　繳回日　　　智權人員　　　繳回日　　　智權人員　　　繳回日　　　智權人員　　　繳回日"
'      Printer.Line (500 - m_dblLeftDiff, 3000)-(12500 - m_dblLeftDiff - 400, 3000)
'      iTabWidth = Printer.TextWidth("智權人員　　　繳回日　　　")
'      iY = 3000
'      '一頁印100個智權人員
'      For ii = 1 To 100
'         If iIdx < UBound(stNames) Then
'            iPos = iIdx Mod 4
'            Printer.CurrentX = 500 - m_dblLeftDiff + iPos * iTabWidth
'            '第一欄
'            If iPos = 0 Then
'               iY = iY + 100
'               Printer.CurrentY = iY
'               Printer.Print stNames(iIdx)
'            '第四欄
'            ElseIf iIdx Mod 4 = 3 Then
'               Printer.CurrentY = iY
'               Printer.Print stNames(iIdx)
'               '畫分格線
'               iY = iY + 400
'               Printer.Line (500 - m_dblLeftDiff, iY)-(12500 - m_dblLeftDiff - 400, iY)
'            '中間欄位
'            Else
'               Printer.CurrentY = iY
'               Printer.Print stNames(iIdx)
'            End If
'         Else
'            bolContinue = False
'            Exit For
'         End If
'         iIdx = iIdx + 1
'      Next
'      '還有資料時跳頁
'      If bolContinue Then Printer.NewPage
'   Loop
'   Printer.EndDoc
'End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   'Modify by Amy 2023/10/11 原W8655 H4125
   Me.Width = 8790
   Me.Height = 4440
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'Mark by Amy 2020/07/08 改回輸入
'   'Add by Amy 2020/04/27 公司別改下拉
'   cboComp.Clear
'   cboComp.AddItem "", 0
'   Call Pub_SetCboCmp(cboComp, False, False, False, , 1)
'   'end 2020/04/27
   'end 2020/07/08
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   'Add By Sindy 2010/9/13
   MaskEdBox3.Mask = DFormat
   MaskEdBox4.Mask = DFormat
   MaskEdBox5.Mask = DFormat
   MaskEdBox6.Mask = DFormat
   '2010/9/13 End
   
   lblSalesName = ""
   'Add by Amy 2017/08/01 預設 未收及勾選「是否含預定收款日未到期者」-瑞婷
   Text3 = "1"
   'Check3.Value = 1 'Remove by Lydia 2018/09/12
   'end 2017/08/01
    Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc1460 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme 'Add by Morgan 2007/10/2
End Sub

Private Sub Text1_Change()
   If Len(Text1) = 5 Then
      lblSalesName = StaffQuery(Text1)
   Else
      lblSalesName = MsgText(601)
   End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_GotFocus()
    TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    'Mark by Amy 2020/07/08 2020/04/27 改成下拉,但瑞婷說要可輸2公司,秀玲說改不判斷公司別,都可輸
'    If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
'      Beep
'      KeyAscii = 0
'    End If
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()

   lngCounter = 0
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   If adoaccrpt106.State = adStateOpen Then
      adoaccrpt106.Close
   End If
   If adoaccrpt106.State = adStateOpen Then adoaccrpt106.Close
   adoaccrpt106.CursorLocation = adUseClient
   'Modify by Amy 2016/09/19 避免同時執行 +where
   adoaccrpt106.Open "select * from accrpt106 Where R10601='" & strUserNum & "' ", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Select Case Text3
      Case "1" '未收
         Select1
      Case "2" '收回
         Select2
      'Add by Amy 2020/04/27 往來
      Case "3"
         Select3
      Case Else
         'MsgBox MsgText(146), , MsgText(5)'Mark by Amy 2022/07/04 改至FormCheck
   End Select
   StatusClear
End Sub

'*************************************************
'  刪除報表資料
'
'*************************************************
Private Sub Accrpt106Delete()
    'Modify by Amy 2016/09/19 避免同時執行抓錯資料 +where
    adoTaie.Execute "Delete from accrpt106 Where R10601='" & strUserNum & "' "
    'Add by Amy 2020/04/27
    If Text3 = "3" Then
        adoTaie.Execute "Delete from accrpt106_1 Where ID='" & strUserNum & "' "
    End If
End Sub

'*************************************************
'  畫線
'
'*************************************************
Private Sub PaintLine(strSign As String)
Dim intCounter As Integer

   For intCounter = 13 To 16
      adoaccrpt106.Fields(intCounter).Value = strSign
   Next intCounter
End Sub

'*************************************************
'  計數器
'
'*************************************************
Private Function Counter() As Long
   lngCounter = lngCounter + 1
   Counter = lngCounter
End Function

'*************************************************
'  列印類別-選擇未收
'
'*************************************************
Private Sub Select1()
Dim strSameName As String
'Modify by Amy 2016/08/01 Double型態改String(Err:多重步驟操作發生錯誤請檢查每一個狀態值)
Dim strCal1 As String
Dim strCal2 As String
Dim strSql As String
'Add By Cheng 2003/05/06
Dim StrSQLa As String
'Add By Cheng 2003/05/12
Dim strCompany As String '公司別
Dim strSalesNo As String '智權人員
Dim strIDSql As String
Dim strCustCaseNo As String 'Add By Sindy 2013/12/5
Dim strQ As String, strCal3 As String, strCal4 As String 'Add by Amy 2016/09/19
Dim strTmp As String 'Add by Amy 2017/08/01 R10626欄位若=Val(R10625)/10 會(Err:多重步驟操作發生錯誤請檢查每一個狀態值)
Dim strCmp As String 'Add by Amy 2020/04/27
On Error GoTo ErrorHandler
   
   strSql = ""
   strAppNo = "" 'Add by Amy 2020/05/28
   'Modify by Amy 2022/07/20 原畫面條件改至function
   strSql = GetWhere(False)
   
   'Modify By Sindy 2010/5/20
'   StrSQLa = "select * from acc0k0, caseprogress, acc0j0, Staff where a0J01 = cp09 (+) and A0K01 = A0J13 (+) And a0k20=st01(+) and (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0)) and (a0k09 is null or a0k09 = 0)" & strSql & " order by st03 asc, a0k20 asc, a0k11 asc, a0k03 asc, a0k04 asc, a0k02 asc, a0k01 asc, cp01 asc, cp02 asc, cp03 asc, cp04 asc "
   'Modified by Morgan 2011/11/2 考慮拆收據情形
   'StrSQLa = "select cp01,cp02,cp03,cp04,cp16,cp18,cp27,a0k01,a0k02,a0k03,a0k04,a0k08,a0k11,a0k20,a0j01,a0j02,a0j20,a0j21 " & _
                     " from acc0k0, caseprogress, acc0j0, Staff where a0J01 = cp09 (+) and A0K01 = A0J13 (+) And a0k20=st01(+) and (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0)) and (a0k09 is null or a0k09 = 0) " & strSql & _
                     " order by a0k22 asc, a0k20 asc, a0k04 asc, a0k02 asc, a0k01 asc, cp09 asc "
   'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
   '2012/5/2 modify by sonia 加預定收款日
   'StrSQLa = "select cp01,cp02,cp03,cp04,nvl(a0j09,0)+nvl(a0j10,0) cp16,nvl(round(a0j09/1000,1),0) cp18,cp27,a0k01,a0k02,a0k03,a0k04,a0k08,a0k11,a0k20,a0j01,a0j02,getcp10desc(cp01,cp10,a0j04) cp10N,na03 " & _
                     " from acc0k0, caseprogress, acc0j0, Staff,nation where a0J01 = cp09 (+) and A0K01 = A0J13 (+) And a0k20=st01(+) and (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0)) and (a0k09 is null or a0k09 = 0) " & strSql & _
                     " and na01(+)=a0j04 order by a0k22 asc, a0k20 asc, a0k04 asc, a0k02 asc, a0k01 asc, cp09 asc "
   'Add By Sindy 2013/6/14 +A0K32
   'Modify by Amy 2016/04/27 +A0k22
   'Modify by Amy 2016/09/19 +nvl(a0j09,0)/nvl(a0j10,0)/a0k05/a0k30
   'Modify by Amy 2017/06/13 +a0k37 is null
   'Modify by Amy 2018/05/21 +a0j07 P119680 a0k30不等於a0j07,先抓a0j07否則非台灣案扣繳金額會錯
   'modify by sonia 2018/5/22 取消A0K30,統一改用A0J07
   'Modified by Lydia 2018/09/12 抓客戶檔-付款週期月份cu175
   'StrSQLa = "select cp01,cp02,cp03,cp04,nvl(a0j09,0)+nvl(a0j10,0) cp16,nvl(round(a0j09/1000,1),0) cp18,cp27,a0k01,a0k02,a0k03,a0k04,a0k08,a0k11,a0k20,a0j01,a0j02,getcp10desc(cp01,cp10,a0j04) cp10N,na03,rd05,a0k32,a0k22,nvl(a0j09,0) a0j09,nvl(a0j10,0) a0j10,a0k05,a0j07" & _
                     " from acc0k0, caseprogress, acc0j0, Staff,nation,(select rd01,rd05 from ReceivablesDay where (rd01,rd02,rd03) in (select rd01,rd02,max(rd03) from ReceivablesDay where (rd01,rd02) in (select rd01,max(rd02) from ReceivablesDay group by rd01 ) group by rd01,rd02)) BB " & _
                     " where a0J01 = cp09 (+) and a0J01=BB.RD01(+) and A0K01 = A0J13 (+) And a0k20=st01(+) and a0k37 is null and (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0)) and (a0k09 is null or a0k09 = 0) " & strSql & _
                     " and na01(+)=a0j04 order by a0k22 asc, a0k20 asc, a0k04 asc, a0k02 asc, a0k01 asc, cp09 asc "
   '2012/5/2 end
   'Modify by Amy 2022/07/20 有下智權人員,需抓案源資料
   strF = "cp01,cp02,cp03,cp04,Nvl(a0j09,0)+nvl(a0j10,0) cp16,Nvl(Round(a0j09/1000,1),0) cp18,cp27,a0k01,a0k02,a0k03,a0k04,a0k08,a0k11,a0k20,a0j01,a0j02" & _
            ",GetCP10Desc(cp01,cp10,a0j04) cp10N,na03,a0k32,a0k22,nvl(a0j09,0) a0j09,nvl(a0j10,0) a0j10,a0k05,a0j07,Decode(Substr(a0k03,1,1),'X',Nvl(cu175,2),null) cu175,cp09 "
                        
   StrSQLa = "Select " & strF & " From Acc0k0, Caseprogress,Acc0j0, Staff,Nation,Customer " & _
                     " where a0J01=cp09 (+) And A0K01= A0J13(+) And a0k37 is null And (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0)) And (a0k09 is null or a0k09 = 0) " & strSql & _
                     " And A0K20=st01(+) And substr(a0k03,1,8)=cu01(+) And substr(a0k03,9,1)=cu02(+) And na01(+)=a0j04 "
   '有下智權人員,抓案源資料
   If Trim(Text1) <> MsgText(601) Then
        StrSQLa = StrSQLa & " Union " & _
                    "Select " & strF & " From Staff,Nation,Customer,(" & GetCaseSource(Text3, GetWhere(True)) & ") " & _
                    "Where Los04=st01(+) And substr(a0k03,1,8)=cu01(+) And substr(a0k03,9,1)=cu02(+) And na01(+)=a0j04 "
        If pub_strUserOffice <> "1" Then
            StrSQLa = StrSQLa & " And ''||ST06='" & pub_strUserOffice & "' "
        End If
   End If
   StrSQLa = StrSQLa & " Order by a0k22 asc, a0k20 asc, a0k04 asc, a0k02 asc, a0k01 asc, cp09 asc "
   'end 2022/07/20
   'end 2018/09/12
   If adoacc0k0.State = adStateOpen Then adoacc0k0.Close
   adoacc0k0.CursorLocation = adUseClient
   adoacc0k0.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0k0.RecordCount = 0 Then
      adoacc0k0.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   bolHasData = True 'Add by Amy 2020/04/27
   Do While adoacc0k0.EOF = False
      If IIf(IsNull(adoacc0k0.Fields("a0k11").Value), MsgText(601), adoacc0k0.Fields("a0k11").Value) & IIf(IsNull(adoacc0k0.Fields("a0k20").Value), MsgText(601), adoacc0k0.Fields("a0k20").Value) & IIf(IsNull(adoacc0k0.Fields("a0k03").Value), MsgText(601), adoacc0k0.Fields("a0k03").Value) <> strSameName Then
         strSameName = IIf(IsNull(adoacc0k0.Fields("a0k03").Value), MsgText(601), adoacc0k0.Fields("a0k11").Value) & IIf(IsNull(adoacc0k0.Fields("a0k20").Value), MsgText(601), adoacc0k0.Fields("a0k20").Value) & IIf(IsNull(adoacc0k0.Fields("a0k03").Value), MsgText(601), adoacc0k0.Fields("a0k03").Value)
      End If
      'Add By Cheng 2003/05/12
      '記錄公司別
      strCompany = "" & adoacc0k0("a0k11").Value
      '記錄智權人員
      strSalesNo = "" & adoacc0k0("a0k20").Value
      adoaccrpt106.AddNew
      adoaccrpt106.Fields("r10601").Value = strUserNum
      '序號
      adoaccrpt106.Fields("r10602").Value = Counter
      '公司別
      If IsNull(adoacc0k0.Fields("a0k11").Value) Then
         adoaccrpt106.Fields("r10603").Value = Null
      Else
         adoaccrpt106.Fields("r10603").Value = adoacc0k0.Fields("a0k11").Value
      End If
      '員工部門
      'Modify by Amy 2016/04/27 原以員編(a0k20)抓目前區別,因10501開始中三區人員調為中四區,故不可抓目前區別
      If IsNull(adoacc0k0.Fields("a0k22").Value) Then
         adoaccrpt106.Fields("r10604").Value = Null
      Else
         adoaccrpt106.Fields("r10604").Value = adoacc0k0.Fields("a0k22").Value 'StaffDeptQuery(adoacc0k0.Fields("a0k20").Value)
      End If
      '員工姓名
      'Modify by Amy 2016/04/27 暫存檔中改存員編
      If IsNull(adoacc0k0.Fields("a0k20").Value) Then
         adoaccrpt106.Fields("r10605").Value = Null
      Else
         adoaccrpt106.Fields("r10605").Value = adoacc0k0.Fields("a0k20").Value 'StaffQuery(adoacc0k0.Fields("a0k20").Value)
      End If
      '客戶編號
      If IsNull(adoacc0k0.Fields("a0k03").Value) Then
         adoaccrpt106.Fields("r10606").Value = Null
      Else
         adoaccrpt106.Fields("r10606").Value = adoacc0k0.Fields("a0k03").Value
      End If
      '收據抬頭
      If IsNull(adoacc0k0.Fields("a0k04").Value) Then
         adoaccrpt106.Fields("r10607").Value = Null
      Else
         adoaccrpt106.Fields("r10607").Value = adoacc0k0.Fields("a0k04").Value
      End If
      '收據日期
      If IsNull(adoacc0k0.Fields("a0k02").Value) Then
         adoaccrpt106.Fields("r10608").Value = Null
      Else
         adoaccrpt106.Fields("r10608").Value = adoacc0k0.Fields("a0k02").Value
      End If
      '收據號碼
      adoaccrpt106.Fields("r10610").Value = adoacc0k0.Fields("a0k01").Value
      '本所案號
      If "" & adoacc0k0.Fields("a0j02").Value <> "" Then
         If Mid(adoacc0k0.Fields("a0j02").Value, Len(adoacc0k0.Fields("a0j02").Value) - 2, 3) = "000" Then
            adoaccrpt106.Fields("r10611").Value = Mid(adoacc0k0.Fields("a0j02").Value, 1, Len(adoacc0k0.Fields("a0j02").Value) - 3)
         Else
            adoaccrpt106.Fields("r10611").Value = adoacc0k0.Fields("a0j02").Value
         End If
      Else
         adoaccrpt106.Fields("r10611").Value = ""
      End If
      
      '案件性質名稱
      'Modified by Morgan 2011/12/27 取消 a0j20
      If IsNull(adoacc0k0.Fields("cp10N").Value) Then
         adoaccrpt106.Fields("r10612").Value = Null
      Else
         adoaccrpt106.Fields("r10612").Value = adoacc0k0.Fields("cp10N").Value
      End If
      '申請國家名稱
      'Modified by Morgan 2011/12/30 取消 a0j21
      If IsNull(adoacc0k0.Fields("na03").Value) Then
         adoaccrpt106.Fields("r10613").Value = Null
      Else
         adoaccrpt106.Fields("r10613").Value = adoacc0k0.Fields("na03").Value
      End If
      '應收金額
      If IsNull(adoacc0k0.Fields("cp16").Value) Then
         adoaccrpt106.Fields("r10614").Value = 0
      Else
         adoaccrpt106.Fields("r10614").Value = adoacc0k0.Fields("cp16").Value
      End If
      'Add By Sindy 2013/6/14 +A0K32
      If IsNull(adoacc0k0.Fields("A0K32").Value) Then
         adoaccrpt106.Fields("r10623").Value = Null
      Else
         adoaccrpt106.Fields("r10623").Value = adoacc0k0.Fields("A0K32").Value
      End If
      '2013/6/14 END
      'Modify by Amy 2020/05/28 +strAppNo 抓申請案號
      strAppNo = "申請案號"
      'Add By Sindy 2013/12/5
      If CaseNameQuery(adoacc0k0.Fields("a0j01").Value, 1, strCustCaseNo, strAppNo) <> "" Then
         If Trim(strCustCaseNo) = "" Then
            adoaccrpt106.Fields("r10624").Value = Null
         Else
            adoaccrpt106.Fields("r10624").Value = strCustCaseNo
         End If
      End If
      '2013/12/5 END
      '申請案號
      If Trim(strAppNo) = MsgText(601) Then
         adoaccrpt106.Fields("r10630").Value = Null
      Else
        adoaccrpt106.Fields("r10630").Value = strAppNo
      End If
      'end 2020/05/28
      
      'Add by Amy 2022/07/04 +介紹人
      If "" & adoacc0k0.Fields("a0k11").Value = "L" Then
        adoaccrpt106.Fields("r10631").Value = GetLos04("" & adoacc0k0.Fields("a0j01"))
      End If
      
      If adocaseprogress.State = adStateOpen Then adocaseprogress.Close
      adocaseprogress.CursorLocation = adUseClient
      'Modified by Morgan 2014/5/16 要考慮欄位值可能為Null
      'Modify by Amy  2016/09/19 +Sum(a1u07)/Sum(a1u09)
      strQ = "select sum(nvl(a1u04,0)+nvl(a1u05,0)-nvl(a1u08,0)-nvl(a1u10,0)), sum(nvl(a1u07,0)+nvl(a1u09,0)),Sum(Nvl(a1u07,0)) as a1u07,Sum(Nvl(a1u09,0)) as a1u09 " & _
                "from acc1u0 where a1u02 = '" & adoacc0k0.Fields("a0k01").Value & "' and a1u03 = '" & adoacc0k0.Fields("a0j01").Value & "'"
      adocaseprogress.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
      If adocaseprogress.RecordCount <> 0 Then
         '已收金額：(本次服務費+本次規費)-(本次退費服務費-本次退費規費)
         If IsNull(adocaseprogress.Fields(0).Value) Then
            adoaccrpt106.Fields("r10615").Value = 0
         Else
            adoaccrpt106.Fields("r10615").Value = adocaseprogress.Fields(0).Value
         End If
         If IsNull(adocaseprogress.Fields(1).Value) = False Then
            '應收金額-銷帳
            If Val(adoaccrpt106.Fields("r10614").Value) <= Val(adocaseprogress.Fields(1).Value) Then
               adoaccrpt106.Fields("r10614").Value = 0
            Else
               adoaccrpt106.Fields("r10614").Value = Val(adoaccrpt106.Fields("r10614").Value) - Val(adocaseprogress.Fields(1).Value)
            End If
         End If
         'Add by Amy 2016/09/19 +應收規費/應收服務費
         adoaccrpt106.Fields("r10617").Value = Val("" & adocaseprogress.Fields("a1u09").Value)
         adoaccrpt106.Fields("r10625").Value = Val("" & adocaseprogress.Fields("a1u07").Value)
      Else
         adoaccrpt106.Fields("r10615").Value = 0
         'Add by Amy 2016/09/19 +應收規費/應收服務費
         adoaccrpt106.Fields("r10617").Value = 0
         adoaccrpt106.Fields("r10625").Value = 0
      End If
      adocaseprogress.Close
      
      strCal1 = Val(adoaccrpt106.Fields("r10614").Value)
      strCal2 = Val(adoaccrpt106.Fields("r10615").Value)
      '未收金額
      adoaccrpt106.Fields("r10616").Value = Val(strCal1) - Val(strCal2)
      If (Val(adoaccrpt106.Fields("r10614").Value) = 0 And IsNull(adoaccrpt106.Fields("r10611").Value) = False) Or Val(adoaccrpt106.Fields("r10614").Value) = Val(adoaccrpt106.Fields("r10615").Value) Then
         adoaccrpt106.Delete
      Else
         'Add By Sindy 2010/5/20 增加欄位
         '案件規費餘額
         'Modify by Amy 2020/04/27 原程式搬至 GetCaseFeesBalance
'         If Trim(adoacc0k0.Fields("cp01").Value) = "TF" Then
'            'Modify By Sindy 2014/8/11 999=>ZZZ
'            'strIDSql = "ax214>='" & Trim(adoacc0k0.Fields("cp01").Value) & Trim(adoacc0k0.Fields("cp02").Value) & "000' AND ax214<='" & Trim(adoacc0k0.Fields("cp01").Value) & Trim(adoacc0k0.Fields("cp02").Value) & "999' "
'            strIDSql = "ax214>='" & Trim(adoacc0k0.Fields("cp01").Value) & Trim(adoacc0k0.Fields("cp02").Value) & "000' AND ax214<='" & Trim(adoacc0k0.Fields("cp01").Value) & Trim(adoacc0k0.Fields("cp02").Value) & "ZZZ' "
'         ElseIf Trim(adoacc0k0.Fields("cp01").Value) = "CFP" Then
'            strIDSql = "ax214>='" & Trim(adoacc0k0.Fields("cp01").Value) & Trim(adoacc0k0.Fields("cp02").Value) & Trim(adoacc0k0.Fields("cp03").Value) & "00' AND ax214<='" & Trim(adoacc0k0.Fields("cp01").Value) & Trim(adoacc0k0.Fields("cp02").Value) & Trim(adoacc0k0.Fields("cp03").Value) & "99' "
'         Else
'            strIDSql = "ax214='" & Trim(adoacc0k0.Fields("cp01").Value) & Trim(adoacc0k0.Fields("cp02").Value) & Trim(adoacc0k0.Fields("cp03").Value) & Trim(adoacc0k0.Fields("cp04").Value) & "'"
'         End If
         strQ = GetCaseFeesBalance(Trim(adoacc0k0.Fields("cp01").Value), Trim(adoacc0k0.Fields("cp02").Value), Trim(adoacc0k0.Fields("cp03").Value), Trim(adoacc0k0.Fields("cp04").Value))
         If adocaseprogress.State = adStateOpen Then adocaseprogress.Close
         adocaseprogress.CursorLocation = adUseClient
         'adocaseprogress.Open "select sum(ax207)-sum(ax206) from acc021 where " & strIDSql & " and SUBSTR(ax205,1,4) = '2201'", adoTaie, adOpenStatic, adLockReadOnly
         adocaseprogress.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
         'end 2020/04/27
         If adocaseprogress.RecordCount <> 0 Then
            If IsNull(adocaseprogress.Fields(0).Value) Then
               'adoaccrpt106.Fields("r10618").Value = 0
            Else
               If Val(adocaseprogress.Fields(0).Value) < 0 Then
                  'Modify By Sindy 2016/12/8 多重步驟...Error
                  'adoaccrpt106.Fields("r10618").Value = Val(adocaseprogress.Fields(0).Value)
                  adoaccrpt106.Fields("r10618").Value = adocaseprogress.Fields(0).Value
                  '2016/12/8 END
                  '2011/9/27 add by sonia 此案號此區間的第一筆寫入規費餘額即可E09921672(P096224)
                  If adocaseprogress.State = adStateOpen Then adocaseprogress.Close
                  adocaseprogress.CursorLocation = adUseClient
                  adocaseprogress.Open "select * from accrpt106 where R10601='" & strUserNum & "' AND R10611='" & adoaccrpt106.Fields("r10611").Value & "' AND nvl(R10618,0)<>0 ", adoTaie, adOpenStatic, adLockReadOnly
                  If adocaseprogress.RecordCount <> 0 Then
                     adoaccrpt106.Fields("r10618").Value = Null
                  End If
                  '2011/9/27 END
               End If
            End If
         Else
            'adoaccrpt106.Fields("r10618").Value = 0
         End If
         adocaseprogress.Close
         
        'Add by Amy 2016/09/19
        '服務費、規費
        'Modify by Amy 2020/04/27 改選 提供客戶用=規費併入服務費顯示,也顯示應收金額,由User自行刪不要欄位 原:Check5.Value=1
        If Option1(0).Value = True Then
            'Modify by Amy 2018/05/21 原抓A0K30,P119680 a0k30不等於a0j07,先抓a0j07否則非台灣案扣繳金額會錯
            If "" & adoacc0k0.Fields("A0j07").Value = "Y" Then
                '合併時,應收規費0 金額列於應收服務費
                adoaccrpt106.Fields("r10617").Value = 0
                adoaccrpt106.Fields("r10625").Value = adoaccrpt106.Fields("r10614").Value '(應收-銷帳)
            Else
                adoaccrpt106.Fields("r10617").Value = Val("" & adoacc0k0.Fields("a0j10").Value) - Val(adoaccrpt106.Fields("r10617").Value)
                adoaccrpt106.Fields("r10625").Value = Val("" & adoacc0k0.Fields("a0j09").Value) - Val(adoaccrpt106.Fields("r10625").Value)
            End If
        Else
            adoaccrpt106.Fields("r10617").Value = Val("" & adoacc0k0.Fields("a0j10").Value)
            adoaccrpt106.Fields("r10625").Value = Val("" & adoacc0k0.Fields("a0j09").Value)
        End If
        
        '扣繳
        '公司別為 J or a0k05為 個人,扣繳固定為 0
        If "" & adoacc0k0.Fields("A0K11").Value = "J" Or "" & adoacc0k0.Fields("A0K05").Value = "1" Then
           adoaccrpt106.Fields("r10626").Value = 0
        'A0K30為合併(Y)時,應收金額/10
        'Modify by Amy 2018/05/21 原抓A0K30,P119680 a0k30不等於a0j07,先抓a0j07否則非台灣案扣繳金額會錯
        ElseIf "" & adoacc0k0.Fields("A0j07").Value = "Y" Then
           'Modify by Amy 2018/11/27 需扣除銷帳
           'strTmp = Val("" & adoacc0k0.Fields("cp16").Value) / 10
           strTmp = Val("" & adoaccrpt106.Fields("r10614").Value) / 10
           adoaccrpt106.Fields("r10626").Value = strTmp
        'A0K30為Null時,應收服務費/10
        Else
            If IsNull(adoaccrpt106.Fields("r10625").Value) = True Then
                adoaccrpt106.Fields("r10626").Value = 0
           Else
                strTmp = Val(adoaccrpt106.Fields("r10625").Value) / 10
                adoaccrpt106.Fields("r10626").Value = strTmp
           End If
        End If
        'end 2016/09/19

         '備註
         If IsNull(adoacc0k0.Fields("a0k08").Value) Then
            adoaccrpt106.Fields("r10619").Value = Null
         Else
            adoaccrpt106.Fields("r10619").Value = adoacc0k0.Fields("a0k08").Value
         End If
         '點數
         If IsNull(adoacc0k0.Fields("cp18").Value) Then
            adoaccrpt106.Fields("r10620").Value = Null
         Else
            adoaccrpt106.Fields("r10620").Value = adoacc0k0.Fields("cp18").Value
         End If
         '發文日
         If IsNull(adoacc0k0.Fields("cp27").Value) Then
            adoaccrpt106.Fields("r10621").Value = Null
         Else
            adoaccrpt106.Fields("r10621").Value = adoacc0k0.Fields("cp27").Value
         End If
         '2010/5/20 End
         '2012/5/2 add by sonia
         '預定收款日
         'Modified by Lydia 2018/09/12 改成付款週期月份
'         If IsNull(adoacc0k0.Fields("rd05").Value) Then
'            adoaccrpt106.Fields("r10622").Value = Null
'         Else
'            adoaccrpt106.Fields("r10622").Value = adoacc0k0.Fields("rd05").Value
'         End If
'         '未達預定收款日
'         If Check3.Value = 0 And IsNull(adoaccrpt106.Fields("r10622").Value) = False Then
'            If Val(adoaccrpt106.Fields("r10622").Value) > Val(strSrvDate(1)) Then adoaccrpt106.Delete
'         End If
'         '2012/5/2 end
         If IsNull(adoacc0k0.Fields("cu175").Value) Then
            adoaccrpt106.Fields("r10622").Value = Null
         Else
            adoaccrpt106.Fields("r10622").Value = adoacc0k0.Fields("cu175").Value
         End If
         'end 2018/09/12
      End If 'end 未收金額
      
      adoaccrpt106.UpdateBatch
      adoacc0k0.MoveNext
''-------------------------------------------------
'' 小計計算
''-------------------------------------------------
'      If adoacc0k0.EOF = False Then
'        '若智權人員不同時
'        If "" & adoacc0k0("a0k20").Value <> strSalesNo Then
'            SubSelect1
'            SubSelect1_2
'            SubSelect1_1
'        '若公司別不同時
'        ElseIf "" & adoacc0k0("a0k11").Value <> strCompany Then
'            SubSelect1
'            SubSelect1_2
'        '若客戶編號不同時
'         ElseIf IIf(IsNull(adoacc0k0.Fields("a0k11").Value), MsgText(601), adoacc0k0.Fields("a0k11").Value) & IIf(IsNull(adoacc0k0.Fields("a0k20").Value), MsgText(601), adoacc0k0.Fields("a0k20").Value) & IIf(IsNull(adoacc0k0.Fields("a0k03").Value), MsgText(601), adoacc0k0.Fields("a0k03").Value) <> strSameName Then
'            SubSelect1
'         End If
'      Else
''-------------------------------------------------
''  合計計算
''-------------------------------------------------
'         SubSelect1
'         SubSelect1_2
'         SubSelect1_1
'
'         adoacc0k0.MoveLast
'         adoaccrpt106.AddNew
'         adoaccrpt106.Fields("r10601").Value = strUserNum
'         adoaccrpt106.Fields("r10602").Value = Counter
'         adoaccrpt106.Fields("r10603").Value = m_strMaxComp
'         If IsNull(adoacc0k0.Fields("a0k20").Value) Then
'            adoaccrpt106.Fields("r10604").Value = Null
'         Else
'            adoaccrpt106.Fields("r10604").Value = StaffDeptQuery(adoacc0k0.Fields("a0k20").Value)
'         End If
'         If IsNull(adoacc0k0.Fields("a0k20").Value) Then
'            adoaccrpt106.Fields("r10605").Value = Null
'         Else
'            adoaccrpt106.Fields("r10605").Value = StaffQuery(adoacc0k0.Fields("a0k20").Value)
'         End If
'         If IsNull(adoacc0k0.Fields("a0k03").Value) Then
'            adoaccrpt106.Fields("r10606").Value = Null
'         Else
'            adoaccrpt106.Fields("r10606").Value = adoacc0k0.Fields("a0k03").Value
'         End If
'         adoaccrpt106.Fields("r10611").Value = ReportSum(25)
'         adoaccsum.CursorLocation = adUseClient
'         StrSQLa = "select nvl(sum(to_number(nvl(r10614, 0))), 0), nvl(sum(to_number(nvl(r10615, 0))), 0) from accrpt106 where r10601 = '" & strUserNum & "' and r10611 = '" & ReportSum(24) & "'"
'         adoaccsum.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'         If adoaccsum.RecordCount <> 0 Then
'            If IsNull(adoaccsum.Fields(0).Value) Then
'               adoaccrpt106.Fields("r10614").Value = 0
'               douCal1 = 0
'            Else
'               adoaccrpt106.Fields("r10614").Value = adoaccsum.Fields(0).Value
'               douCal1 = adoaccsum.Fields(0).Value
'            End If
'            If IsNull(adoaccsum.Fields(1).Value) = False Then
'               adoaccrpt106.Fields("r10615").Value = adoaccsum.Fields(1).Value
'               douCal2 = Val(adoaccsum.Fields(1).Value)
'            Else
'               adoaccrpt106.Fields("r10615").Value = 0
'               douCal2 = 0
'            End If
'         Else
'            adoaccrpt106.Fields("r10614").Value = 0
'            adoaccrpt106.Fields("r10615").Value = 0
'            douCal1 = 0
'            douCal2 = 0
'         End If
'         adoaccsum.Close
'         adoaccrpt106.Fields("r10616").Value = douCal1 - douCal2
'         adoaccrpt106.UpdateBatch
'         adoaccrpt106.AddNew
'         adoaccrpt106.Fields("r10601").Value = strUserNum
'         adoaccrpt106.Fields("r10602").Value = Counter
'        adoaccrpt106.Fields("r10603").Value = m_strMaxComp
'         If IsNull(adoacc0k0.Fields("a0k20").Value) Then
'            adoaccrpt106.Fields("r10604").Value = Null
'         Else
'            adoaccrpt106.Fields("r10604").Value = StaffDeptQuery(adoacc0k0.Fields("a0k20").Value)
'         End If
'         If IsNull(adoacc0k0.Fields("a0k20").Value) Then
'            adoaccrpt106.Fields("r10605").Value = Null
'         Else
'            adoaccrpt106.Fields("r10605").Value = StaffQuery(adoacc0k0.Fields("a0k20").Value)
'         End If
'         If IsNull(adoacc0k0.Fields("a0k03").Value) Then
'            adoaccrpt106.Fields("r10606").Value = Null
'         Else
'            adoaccrpt106.Fields("r10606").Value = adoacc0k0.Fields("a0k03").Value
'         End If
'         PaintLine ReportSum(8)
'         adoaccrpt106.UpdateBatch
'         adoacc0k0.MoveNext
'      End If
   Loop
   adoacc0k0.Close
   Exit Sub
ErrorHandler:
   If adoacc0k0.State <> adStateClosed Then adoacc0k0.Close
   Set adoacc0k0 = Nothing
   If adocaseprogress.State <> adStateClosed Then adocaseprogress.Close
   Set adocaseprogress = Nothing
   MsgBox Err.Description, , MsgText(5)
End Sub

''*************************************************
''  列印類別-選擇未收之小計計算
''
''*************************************************
'Private Sub SubSelect1()
'Dim douCal1, douCal2 As Double
'Dim strSql As String
'
'   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'      strSql = " and a0k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'   End If
'   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'      strSql = strSql & " and a0k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'   End If
'
'   adoacc0k0.MovePrevious
'   adoaccrpt106.AddNew
'   adoaccrpt106.Fields("r10601").Value = strUserNum
'   adoaccrpt106.Fields("r10602").Value = Counter
'   If IsNull(adoacc0k0.Fields("a0k11").Value) Then
'      adoaccrpt106.Fields("r10603").Value = Null
'   Else
'      adoaccrpt106.Fields("r10603").Value = adoacc0k0.Fields("a0k11").Value
'   End If
'   If IsNull(adoacc0k0.Fields("a0k20").Value) Then
'      adoaccrpt106.Fields("r10604").Value = Null
'   Else
'      adoaccrpt106.Fields("r10604").Value = StaffDeptQuery(adoacc0k0.Fields("a0k20").Value)
'   End If
'   If IsNull(adoacc0k0.Fields("a0k20").Value) Then
'      adoaccrpt106.Fields("r10605").Value = Null
'   Else
'      adoaccrpt106.Fields("r10605").Value = StaffQuery(adoacc0k0.Fields("a0k20").Value)
'   End If
'   If IsNull(adoacc0k0.Fields("a0k03").Value) Then
'      adoaccrpt106.Fields("r10606").Value = Null
'   Else
'      adoaccrpt106.Fields("r10606").Value = adoacc0k0.Fields("a0k03").Value
'   End If
'   PaintLine ReportSum(4)
'   adoaccrpt106.AddNew
'   adoaccrpt106.Fields("r10601").Value = strUserNum
'   adoaccrpt106.Fields("r10602").Value = Counter
'   If IsNull(adoacc0k0.Fields("a0k11").Value) Then
'      adoaccrpt106.Fields("r10603").Value = Null
'   Else
'      adoaccrpt106.Fields("r10603").Value = adoacc0k0.Fields("a0k11").Value
'   End If
'   If IsNull(adoacc0k0.Fields("a0k20").Value) Then
'      adoaccrpt106.Fields("r10604").Value = Null
'   Else
'      adoaccrpt106.Fields("r10604").Value = StaffDeptQuery(adoacc0k0.Fields("a0k20").Value)
'   End If
'   If IsNull(adoacc0k0.Fields("a0k20").Value) Then
'      adoaccrpt106.Fields("r10605").Value = Null
'   Else
'      adoaccrpt106.Fields("r10605").Value = StaffQuery(adoacc0k0.Fields("a0k20").Value)
'   End If
'   If IsNull(adoacc0k0.Fields("a0k03").Value) Then
'      adoaccrpt106.Fields("r10606").Value = Null
'   Else
'      adoaccrpt106.Fields("r10606").Value = adoacc0k0.Fields("a0k03").Value
'   End If
'   adoaccrpt106.Fields("r10611").Value = ReportSum(24)
'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open "select sum(to_number(nvl(r10614, 0))), sum(to_number(nvl(r10615, 0))) from accrpt106 where r10601 = '" & strUserNum & "' and r10607 is not null and r10605 = '" & StaffQuery(IIf(IsNull(adoacc0k0.Fields("a0k20").Value), MsgText(601), adoacc0k0.Fields("a0k20").Value)) & "' and r10606 = '" & adoacc0k0.Fields("a0k03").Value & "' and r10603 = '" & adoacc0k0.Fields("a0k11").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         adoaccrpt106.Fields("r10614").Value = 0
'         douCal1 = 0
'      Else
'         adoaccrpt106.Fields("r10614").Value = adoaccsum.Fields(0).Value
'         douCal1 = adoaccsum.Fields(0).Value
'      End If
'      If IsNull(adoaccsum.Fields(1).Value) = False Then
'         adoaccrpt106.Fields("r10615").Value = adoaccsum.Fields(1).Value
'         douCal2 = Val(adoaccsum.Fields(1).Value)
'      Else
'         adoaccrpt106.Fields("r10615").Value = 0
'         douCal2 = 0
'      End If
'   Else
'      adoaccrpt106.Fields("r10614").Value = 0
'      adoaccrpt106.Fields("r10615").Value = 0
'      douCal1 = 0
'      douCal2 = 0
'   End If
'   adoaccsum.Close
'   adoaccrpt106.Fields("r10616").Value = douCal1 - douCal2
'   adoaccrpt106.AddNew
'   adoaccrpt106.Fields("r10601").Value = strUserNum
'   adoaccrpt106.Fields("r10602").Value = Counter
'   If IsNull(adoacc0k0.Fields("a0k11").Value) Then
'      adoaccrpt106.Fields("r10603").Value = Null
'   Else
'      adoaccrpt106.Fields("r10603").Value = adoacc0k0.Fields("a0k11").Value
'   End If
'   If IsNull(adoacc0k0.Fields("a0k20").Value) Then
'      adoaccrpt106.Fields("r10604").Value = Null
'   Else
'      adoaccrpt106.Fields("r10604").Value = StaffDeptQuery(adoacc0k0.Fields("a0k20").Value)
'   End If
'   If IsNull(adoacc0k0.Fields("a0k20").Value) Then
'      adoaccrpt106.Fields("r10605").Value = Null
'   Else
'      adoaccrpt106.Fields("r10605").Value = StaffQuery(adoacc0k0.Fields("a0k20").Value)
'   End If
'   If IsNull(adoacc0k0.Fields("a0k03").Value) Then
'      adoaccrpt106.Fields("r10606").Value = Null
'   Else
'      adoaccrpt106.Fields("r10606").Value = adoacc0k0.Fields("a0k03").Value
'   End If
'   PaintLine ReportSum(4)
'   If douCal1 = 0 Then
'      adoaccrpt106.CancelBatch
'   Else
'      adoaccrpt106.UpdateBatch
'   End If
'   adoacc0k0.MoveNext
'End Sub
'
''Add By Cheng 2003/05/12
''*************************************************
''  列印類別-選擇未收之智權人員小計計算
''
''*************************************************
'Private Sub SubSelect1_1()
'Dim douCal1, douCal2 As Double
'Dim strSql As String
'Dim StrSQLa As String
''Add By Cheng 2004/02/05
'Dim StrSqlB As String
'Dim rsB As New ADODB.Recordset
'
'   adoacc0k0.MovePrevious
'   adoaccrpt106.AddNew
'   adoaccrpt106.Fields("r10601").Value = strUserNum
'   adoaccrpt106.Fields("r10602").Value = Counter
'    '取得智權人員目前有資料的最大公司別
'    StrSqlB = "Select Max(R10603) From ACCRPT106 Where R10601 = '" & strUserNum & "' And R10605 = '" & StaffQuery(IIf(IsNull(adoacc0k0.Fields("a0k20").Value), MsgText(601), adoacc0k0.Fields("a0k20").Value)) & "' "
'    rsB.CursorLocation = adUseClient
'    rsB.Open StrSqlB, adoTaie, adOpenStatic, adLockReadOnly
'    m_strMaxComp = ""
'    While Not rsB.EOF
'        m_strMaxComp = "" & rsB.Fields(0).Value
'        rsB.MoveNext
'    Wend
'    If rsB.State <> adStateClosed Then rsB.Close
'    Set rsB = Nothing
'    adoaccrpt106.Fields("r10603").Value = m_strMaxComp
'   If IsNull(adoacc0k0.Fields("a0k20").Value) Then
'      adoaccrpt106.Fields("r10604").Value = Null
'   Else
'      adoaccrpt106.Fields("r10604").Value = StaffDeptQuery(adoacc0k0.Fields("a0k20").Value)
'   End If
'   If IsNull(adoacc0k0.Fields("a0k20").Value) Then
'      adoaccrpt106.Fields("r10605").Value = Null
'   Else
'      adoaccrpt106.Fields("r10605").Value = StaffQuery(adoacc0k0.Fields("a0k20").Value)
'   End If
'   If IsNull(adoacc0k0.Fields("a0k03").Value) Then
'      adoaccrpt106.Fields("r10606").Value = Null
'   Else
'      adoaccrpt106.Fields("r10606").Value = adoacc0k0.Fields("a0k03").Value
'   End If
'   adoaccrpt106.Fields("r10611").Value = "智權人員小計:"
'   adoaccsum.CursorLocation = adUseClient
'    StrSQLa = "select sum(to_number(nvl(r10614, 0))), sum(to_number(nvl(r10615, 0))) from accrpt106 where r10601 = '" & strUserNum & "' and r10607 is not null and r10605 = '" & StaffQuery(IIf(IsNull(adoacc0k0.Fields("a0k20").Value), MsgText(601), adoacc0k0.Fields("a0k20").Value)) & "' "
'   adoaccsum.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         adoaccrpt106.Fields("r10614").Value = 0
'         douCal1 = 0
'      Else
'         adoaccrpt106.Fields("r10614").Value = adoaccsum.Fields(0).Value
'         douCal1 = adoaccsum.Fields(0).Value
'      End If
'      If IsNull(adoaccsum.Fields(1).Value) = False Then
'         adoaccrpt106.Fields("r10615").Value = adoaccsum.Fields(1).Value
'         douCal2 = Val(adoaccsum.Fields(1).Value)
'      Else
'         adoaccrpt106.Fields("r10615").Value = 0
'         douCal2 = 0
'      End If
'   Else
'      adoaccrpt106.Fields("r10614").Value = 0
'      adoaccrpt106.Fields("r10615").Value = 0
'      douCal1 = 0
'      douCal2 = 0
'   End If
'   adoaccsum.Close
'   adoaccrpt106.Fields("r10616").Value = douCal1 - douCal2
'   adoaccrpt106.AddNew
'   adoaccrpt106.Fields("r10601").Value = strUserNum
'   adoaccrpt106.Fields("r10602").Value = Counter
'    adoaccrpt106.Fields("r10603").Value = m_strMaxComp
'   If IsNull(adoacc0k0.Fields("a0k20").Value) Then
'      adoaccrpt106.Fields("r10604").Value = Null
'   Else
'      adoaccrpt106.Fields("r10604").Value = StaffDeptQuery(adoacc0k0.Fields("a0k20").Value)
'   End If
'   If IsNull(adoacc0k0.Fields("a0k20").Value) Then
'      adoaccrpt106.Fields("r10605").Value = Null
'   Else
'      adoaccrpt106.Fields("r10605").Value = StaffQuery(adoacc0k0.Fields("a0k20").Value)
'   End If
'   If IsNull(adoacc0k0.Fields("a0k03").Value) Then
'      adoaccrpt106.Fields("r10606").Value = Null
'   Else
'      adoaccrpt106.Fields("r10606").Value = adoacc0k0.Fields("a0k03").Value
'   End If
'   PaintLine ReportSum(4)
'   If douCal1 = 0 Then
'      adoaccrpt106.CancelBatch
'   Else
'      adoaccrpt106.UpdateBatch
'   End If
'   adoacc0k0.MoveNext
'End Sub
'
''Add By Cheng 2003/05/12
''*************************************************
''  列印類別-選擇未收之公司別小計計算
''
''*************************************************
'Private Sub SubSelect1_2()
'Dim douCal1, douCal2 As Double
'Dim strSql As String
'Dim StrSQLa As String
'
'   adoacc0k0.MovePrevious
'   adoaccrpt106.AddNew
'   adoaccrpt106.Fields("r10601").Value = strUserNum
'   adoaccrpt106.Fields("r10602").Value = Counter
'   If IsNull(adoacc0k0.Fields("a0k11").Value) Then
'      adoaccrpt106.Fields("r10603").Value = Null
'   Else
'      adoaccrpt106.Fields("r10603").Value = adoacc0k0.Fields("a0k11").Value
'   End If
'   If IsNull(adoacc0k0.Fields("a0k20").Value) Then
'      adoaccrpt106.Fields("r10604").Value = Null
'   Else
'      adoaccrpt106.Fields("r10604").Value = StaffDeptQuery(adoacc0k0.Fields("a0k20").Value)
'   End If
'   If IsNull(adoacc0k0.Fields("a0k20").Value) Then
'      adoaccrpt106.Fields("r10605").Value = Null
'   Else
'      adoaccrpt106.Fields("r10605").Value = StaffQuery(adoacc0k0.Fields("a0k20").Value)
'   End If
'   If IsNull(adoacc0k0.Fields("a0k03").Value) Then
'      adoaccrpt106.Fields("r10606").Value = Null
'   Else
'      adoaccrpt106.Fields("r10606").Value = adoacc0k0.Fields("a0k03").Value
'   End If
'   adoaccrpt106.Fields("r10611").Value = "公司別小計:"
'   adoaccsum.CursorLocation = adUseClient
'    StrSQLa = "select sum(to_number(nvl(r10614, 0))), sum(to_number(nvl(r10615, 0))) from accrpt106 where r10601 = '" & strUserNum & "' and r10607 is not null and r10605 = '" & StaffQuery(IIf(IsNull(adoacc0k0.Fields("a0k20").Value), MsgText(601), adoacc0k0.Fields("a0k20").Value)) & "' and r10603 = '" & adoacc0k0.Fields("a0k11").Value & "'"
'   adoaccsum.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         adoaccrpt106.Fields("r10614").Value = 0
'         douCal1 = 0
'      Else
'         adoaccrpt106.Fields("r10614").Value = adoaccsum.Fields(0).Value
'         douCal1 = adoaccsum.Fields(0).Value
'      End If
'      If IsNull(adoaccsum.Fields(1).Value) = False Then
'         adoaccrpt106.Fields("r10615").Value = adoaccsum.Fields(1).Value
'         douCal2 = Val(adoaccsum.Fields(1).Value)
'      Else
'         adoaccrpt106.Fields("r10615").Value = 0
'         douCal2 = 0
'      End If
'   Else
'      adoaccrpt106.Fields("r10614").Value = 0
'      adoaccrpt106.Fields("r10615").Value = 0
'      douCal1 = 0
'      douCal2 = 0
'   End If
'   adoaccsum.Close
'   adoaccrpt106.Fields("r10616").Value = douCal1 - douCal2
'   adoaccrpt106.AddNew
'   adoaccrpt106.Fields("r10601").Value = strUserNum
'   adoaccrpt106.Fields("r10602").Value = Counter
'   If IsNull(adoacc0k0.Fields("a0k11").Value) Then
'      adoaccrpt106.Fields("r10603").Value = Null
'   Else
'      adoaccrpt106.Fields("r10603").Value = adoacc0k0.Fields("a0k11").Value
'   End If
'   If IsNull(adoacc0k0.Fields("a0k20").Value) Then
'      adoaccrpt106.Fields("r10604").Value = Null
'   Else
'      adoaccrpt106.Fields("r10604").Value = StaffDeptQuery(adoacc0k0.Fields("a0k20").Value)
'   End If
'   If IsNull(adoacc0k0.Fields("a0k20").Value) Then
'      adoaccrpt106.Fields("r10605").Value = Null
'   Else
'      adoaccrpt106.Fields("r10605").Value = StaffQuery(adoacc0k0.Fields("a0k20").Value)
'   End If
'   If IsNull(adoacc0k0.Fields("a0k03").Value) Then
'      adoaccrpt106.Fields("r10606").Value = Null
'   Else
'      adoaccrpt106.Fields("r10606").Value = adoacc0k0.Fields("a0k03").Value
'   End If
'   PaintLine ReportSum(4)
'   If douCal1 = 0 Then
'      adoaccrpt106.CancelBatch
'   Else
'      adoaccrpt106.UpdateBatch
'   End If
'   adoacc0k0.MoveNext
'End Sub

'*************************************************
'  列印類別-選擇收回
'
'*************************************************
Private Sub Select2()
Dim strSameName As String
Dim strSql As String
'Add By Cheng 2003/05/12
Dim StrSQLa As String
Dim strCompany As String '公司別
Dim strSalesNo As String '智權人員
Dim strCustCaseNo As String 'Add By Sindy 2013/12/5
Dim strCmp As String 'Add by Amy 2020/04/27

On Error GoTo Checking
   strSql = ""
   strAppNo = "" 'Add by Amy 2020/05/28
   'Modify by Amy 2022/07/20 原畫面條件改至function
   strSql = GetWhere(False)
   
   strSameName = ""
   'Modify By Sindy 2010/5/20
'   StrSQLa = "select * from acc0k0, caseprogress, acc0j0, acc0l0, acc0m0, Staff where a0J01 = cp09 (+) and a0k01 = a0j13 (+) and a0k01 = a0m02 and a0m01 = a0l01 And a0k20=st01(+) and (a0k09 is null or a0k09 = 0)" & strSql & _
'                    " order by st03 asc, a0k20 asc, a0k11 asc, a0k03 asc, a0k04 asc, a0k02 asc, a0l01 asc, a0k01 asc, cp01 asc, cp02 asc, cp03 asc, cp04 asc "
   'Modified by Morgan 2011/11/2 考慮拆收據情形
   'StrSQLa = "select a0k01,a0k03,a0k04,a0k11,a0k20,a0l02,a0l01,a0j01,a0j02,a0j20,a0j21 " & _
                    " from acc0k0, caseprogress, acc0j0, acc0l0, acc0m0, Staff where a0J01 = cp09 (+) and a0k01 = a0j13 (+) and a0k01 = a0m02 and a0m01 = a0l01 And a0k20=st01(+) and (a0k09 is null or a0k09 = 0) " & strSql & _
                    " order by a0k22 asc, a0k20 asc, a0k04 asc, a0k02 asc, a0k01 asc, cp09 asc "
   'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
   'Modify by Amy 2016/04/27 +A0k22
   'Modify by Amy 2016/09/19 +A1u06
   'Modify by Amy 2020/04/27 +A0j07
   'Modify by Amy 2022/07/20 有下智權人員,需抓案源資料
   strF = "a0k01,a0k03,a0k04,a0k11,a0k20,a0l02,a0l01,a0j01,a0j02,GetCP10Desc(cp01,cp10,a0j04) cp10N,na03,a0k22,a1u06,a0j07,cp09,a0k02 "
   
   StrSQLa = "Select " & strF & " From Acc0l0,Acc1u0,Acc0k0,Caseprogress,Acc0j0,Staff,Nation" & _
                    " Where a1u01(+)=a0l01 and a1u02=a0k01(+) and (a0k09 is null or a0k09 = 0) and cp09(+)=a1u03 " & _
                    " and a0j01(+)=a1u03 and a0j13(+)=a1u02 And st01(+)=a0k20  and na01(+)=a0j04 " & strSql
                  
   '有下智權人員,抓案源資料
   If Trim(Text1) <> MsgText(601) Then
        strSql = Replace(strSql, "a0k20", "los04")
        StrSQLa = StrSQLa & " Union " & _
                        "Select " & strF & " From Acc0j0,Staff,Nation,(" & GetCaseSource(Text3, GetWhere(True)) & ") " & _
                       " Where a0j01(+)=a1u03 And a0j13(+)=a1u02 And st01(+)=los04 And na01(+)=a0j04 "
        If pub_strUserOffice <> "1" Then
            StrSQLa = StrSQLa & " And ''||ST06='" & pub_strUserOffice & "' "
        End If
   End If
   StrSQLa = StrSQLa & " Order by a0k22 asc, a0k20 asc, a0k04 asc, a0k02 asc, a0k01 asc, cp09 asc "
   'end 2022/07/20
   If adoacc0l0.State = adStateOpen Then adoacc0l0.Close
   adoacc0l0.CursorLocation = adUseClient
   adoacc0l0.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0l0.RecordCount = 0 Then
      adoacc0l0.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   
   bolHasData = True 'Add by Amy 2020/04/27
   Do While adoacc0l0.EOF = False
      If IIf(IsNull(adoacc0l0.Fields("a0k11").Value), MsgText(601), adoacc0l0.Fields("a0k11").Value) & IIf(IsNull(adoacc0l0.Fields("a0k20").Value), MsgText(601), adoacc0l0.Fields("a0k20").Value) & IIf(IsNull(adoacc0l0.Fields("a0k03").Value), MsgText(601), adoacc0l0.Fields("a0k03").Value) <> strSameName Then
         strSameName = IIf(IsNull(adoacc0l0.Fields("a0k03").Value), MsgText(601), adoacc0l0.Fields("a0k11").Value) & IIf(IsNull(adoacc0l0.Fields("a0k20").Value), MsgText(601), adoacc0l0.Fields("a0k20").Value) & IIf(IsNull(adoacc0l0.Fields("a0k03").Value), MsgText(601), adoacc0l0.Fields("a0k03").Value)
      End If
      '記錄公司別
      strCompany = "" & adoacc0l0("a0k11").Value
      '記錄智權人員
      strSalesNo = "" & adoacc0l0("a0k20").Value
      adoaccrpt106.AddNew
      adoaccrpt106.Fields("r10601").Value = strUserNum
      adoaccrpt106.Fields("r10602").Value = Counter
      '公司別
      If IsNull(adoacc0l0.Fields("a0k11").Value) Then
         adoaccrpt106.Fields("r10603").Value = Null
      Else
         adoaccrpt106.Fields("r10603").Value = adoacc0l0.Fields("a0k11").Value
      End If
      '員工部門
      'Modify by Amy 2016/04/27 原以員編(a0k20)抓目前區別,因10501開始中三區人員調為中四區,故不可抓目前區別
      If IsNull(adoacc0l0.Fields("a0k22").Value) Then
         adoaccrpt106.Fields("r10604").Value = Null
      Else
         adoaccrpt106.Fields("r10604").Value = adoacc0l0.Fields("a0k22").Value 'StaffDeptQuery(adoacc0l0.Fields("a0k20").Value)
      End If
      '員工姓名
      'Modify by Amy 2016/04/27 暫存檔中改存員編
      If IsNull(adoacc0l0.Fields("a0k20").Value) Then
         adoaccrpt106.Fields("r10605").Value = Null
      Else
         'adoaccrpt106.Fields("r10605").Value = adoacc0l0.Fields("a0k20").Value & " " & StaffQuery(adoacc0l0.Fields("a0k20").Value)
         adoaccrpt106.Fields("r10605").Value = adoacc0l0.Fields("a0k20").Value 'StaffQuery(adoacc0l0.Fields("a0k20").Value)
      End If
      'Add By Cheng 2003/05/14
      '客戶編號
      If IsNull(adoacc0l0.Fields("a0k03").Value) Then
         adoaccrpt106.Fields("r10606").Value = Null
      Else
         adoaccrpt106.Fields("r10606").Value = adoacc0l0.Fields("a0k03").Value
      End If
      If IsNull(adoacc0l0.Fields("a0k04").Value) Then
         adoaccrpt106.Fields("r10607").Value = Null
      Else
         adoaccrpt106.Fields("r10607").Value = adoacc0l0.Fields("a0k04").Value
      End If
      '收款日期
      If IsNull(adoacc0l0.Fields("a0l02").Value) Then
         adoaccrpt106.Fields("r10608").Value = Null
      Else
         adoaccrpt106.Fields("r10608").Value = adoacc0l0.Fields("a0l02").Value
      End If
      '收款單號
      adoaccrpt106.Fields("r10609").Value = "" & adoacc0l0.Fields("a0l01").Value
      '收據號碼
      adoaccrpt106.Fields("r10610").Value = "" & adoacc0l0.Fields("a0k01").Value
      
      If IsNull(adoacc0l0.Fields("a0j02").Value) Then
         adoaccrpt106.Fields("r10611").Value = Null
      Else
         If Mid(adoacc0l0.Fields("a0j02").Value, Len(adoacc0l0.Fields("a0j02").Value) - 2, 3) = "000" Then
            adoaccrpt106.Fields("r10611").Value = Mid(adoacc0l0.Fields("a0j02").Value, 1, Len(adoacc0l0.Fields("a0j02").Value) - 3)
         Else
            adoaccrpt106.Fields("r10611").Value = adoacc0l0.Fields("a0j02").Value
         End If
      End If
      'Modified by Morgan 2011/12/27 取消 a0j20
      If IsNull(adoacc0l0.Fields("cp10N").Value) Then
         adoaccrpt106.Fields("r10612").Value = Null
      Else
         adoaccrpt106.Fields("r10612").Value = adoacc0l0.Fields("cp10N").Value
      End If
      
      'Modified by Morgan 2011/12/30 取消 a0j21
      If IsNull(adoacc0l0.Fields("na03").Value) Then
         adoaccrpt106.Fields("r10613").Value = Null
      Else
         adoaccrpt106.Fields("r10613").Value = adoacc0l0.Fields("na03").Value
      End If
      
      'Modify by Amy 2020/05/28 +strAppNo 抓申請案號
      strAppNo = "申請案號"
      'Add By Sindy 2013/12/5
      If CaseNameQuery(adoacc0l0.Fields("a0j01").Value, 1, strCustCaseNo, strAppNo) <> "" Then
         If Trim(strCustCaseNo) = "" Then
            adoaccrpt106.Fields("r10624").Value = Null
         Else
            adoaccrpt106.Fields("r10624").Value = strCustCaseNo
         End If
      End If
      '2013/12/5 END
      If Trim(strAppNo) = MsgText(601) Then
        adoaccrpt106.Fields("r10630").Value = Null
      Else
        adoaccrpt106.Fields("r10630").Value = strAppNo
      End If
      'end 2020/05/28
      
      'Add by Amy 2022/07/04 +介紹人
      If "" & adoacc0l0.Fields("a0k11").Value = "L" Then
        adoaccrpt106.Fields("r10631").Value = GetLos04("" & adoacc0l0.Fields("a0j01"))
      End If
      
      'Add  by Amy 2016/09/19 +扣繳
      If IsNull(adoacc0l0.Fields("a1u06").Value) Then
         adoaccrpt106.Fields("r10626").Value = Null
      Else
         adoaccrpt106.Fields("r10626").Value = adoacc0l0.Fields("a1u06").Value
      End If
      
      StrSQLa = "select sum(a1u04), sum(a1u05), sum(a1u04+a1u05) from acc1u0 where a1u02 = '" & adoacc0l0.Fields("a0k01").Value & "' and a1u03 = '" & adoacc0l0.Fields("a0j01").Value & "'"
      If adocaseprogress.State = adStateOpen Then adocaseprogress.Close
      adocaseprogress.CursorLocation = adUseClient
      adocaseprogress.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
      If adocaseprogress.RecordCount <> 0 Then
         '本次服務費
         If IsNull(adocaseprogress.Fields(0).Value) Then
            adoaccrpt106.Fields("r10625").Value = 0
         Else
            adoaccrpt106.Fields("r10625").Value = Val(adocaseprogress.Fields(0).Value)
         End If
         '本次規費
         If IsNull(adocaseprogress.Fields(1).Value) Then
            adoaccrpt106.Fields("r10617").Value = 0
         Else
            adoaccrpt106.Fields("r10617").Value = Val(adocaseprogress.Fields(1).Value)
         End If
         '收款金額:本次服務費+本次規費
         If IsNull(adocaseprogress.Fields(2).Value) Then
            adoaccrpt106.Fields("r10614").Value = 0
         Else
            adoaccrpt106.Fields("r10614").Value = Val(adocaseprogress.Fields(2).Value)
         End If
         'Add by Amy 2020/04/27 提供客戶用=規費併入服務費顯示
         If Option1(0).Value = True And "" & adoacc0l0.Fields("A0j07").Value = "Y" Then
            '合併時,服務費=應收金額,規費0
            adoaccrpt106.Fields("r10625").Value = adoaccrpt106.Fields("r10614").Value '收款金額
            adoaccrpt106.Fields("r10617").Value = 0
         End If
         'end 2020/04/27
      End If
      adocaseprogress.Close
      
      adoaccrpt106.UpdateBatch
      adoacc0l0.MoveNext
''-------------------------------------------------
'' 小計計算
''-------------------------------------------------
'      If adoacc0l0.EOF = False Then
'        '若智權人員不同時
'        If "" & adoacc0l0("a0k20").Value <> strSalesNo Then
'            SubSelect2
'            SubSelect2_2
'            SubSelect2_1
'        '若公司別不同時
'        ElseIf "" & adoacc0l0("a0k11").Value <> strCompany Then
'            SubSelect2
'            SubSelect2_2
'        '若客戶編號不同時
'         ElseIf IIf(IsNull(adoacc0l0.Fields("a0k11").Value), MsgText(601), adoacc0l0.Fields("a0k11").Value) & IIf(IsNull(adoacc0l0.Fields("a0k20").Value), MsgText(601), adoacc0l0.Fields("a0k20").Value) & IIf(IsNull(adoacc0l0.Fields("a0k03").Value), MsgText(601), adoacc0l0.Fields("a0k03").Value) <> strSameName Then
'            SubSelect2
'         End If
'      Else
''-------------------------------------------------
''  合計計算
''-------------------------------------------------
'         SubSelect2
'         SubSelect2_2
'         SubSelect2_1
'
'         adoacc0l0.MoveLast
'         adoaccrpt106.UpdateBatch
'         adoaccrpt106.AddNew
'         adoaccrpt106.Fields("r10601").Value = strUserNum
'         adoaccrpt106.Fields("r10602").Value = Counter
'         adoaccrpt106.Fields("r10603").Value = "" & adoacc0l0.Fields("a0k11").Value
'         If IsNull(adoacc0l0.Fields("a0k20").Value) Then
'            adoaccrpt106.Fields("r10605").Value = Null
'         Else
'            adoaccrpt106.Fields("r10605").Value = adoacc0l0.Fields("a0k20").Value & " " & StaffQuery(adoacc0l0.Fields("a0k20").Value)
'         End If
'         adoaccrpt106.Fields("r10611").Value = ReportSum(25)
'         adoaccsum.CursorLocation = adUseClient
'         StrSQLa = "select sum(to_number(nvl(r10614, 0))), sum(to_number(nvl(r10616, 0))), sum(to_number(nvl(r10617, 0))) from accrpt106 where r10601 = '" & strUserNum & "' and r10607 is not null "
'         adoaccsum.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'         If adoaccsum.RecordCount <> 0 Then
'            adoaccrpt106.Fields("r10614").Value = IIf("" & adoaccsum.Fields(0).Value = "", 0, adoaccsum.Fields(0).Value)
'            adoaccrpt106.Fields("r10616").Value = IIf("" & adoaccsum.Fields(1).Value = "", 0, adoaccsum.Fields(1).Value)
'            adoaccrpt106.Fields("r10617").Value = IIf("" & adoaccsum.Fields(2).Value = "", 0, adoaccsum.Fields(2).Value)
'         Else
'            adoaccrpt106.Fields("r10614").Value = 0
'            adoaccrpt106.Fields("r10616").Value = 0
'            adoaccrpt106.Fields("r10617").Value = 0
'         End If
'         adoaccsum.Close
'         adoaccrpt106.UpdateBatch
'         adoaccrpt106.AddNew
'         adoaccrpt106.Fields("r10601").Value = strUserNum
'         adoaccrpt106.Fields("r10602").Value = Counter
'         If IsNull(adoacc0l0.Fields("a0k20").Value) Then
'            adoaccrpt106.Fields("r10605").Value = Null
'         Else
'            adoaccrpt106.Fields("r10605").Value = adoacc0l0.Fields("a0k20").Value & " " & StaffQuery(adoacc0l0.Fields("a0k20").Value)
'         End If
'         PaintLine ReportSum(8)
'         adoacc0l0.MoveNext
'      End If
   Loop
   adoacc0l0.Close
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

''*************************************************
''  列印類別-選擇收回之小計計算
''
''*************************************************
'Private Sub SubSelect2()
'Dim strSql As String
''Add By Cheng 2003/05/12
'Dim StrSQLa As String
'
'   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'      strSql = " and a0l02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'   End If
'   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'      strSql = strSql & " and a0l02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'   End If
'
'   adoacc0l0.MovePrevious
'   adoaccrpt106.AddNew
'   adoaccrpt106.Fields("r10601").Value = strUserNum
'   adoaccrpt106.Fields("r10602").Value = Counter
'   adoaccrpt106.Fields("r10603").Value = "" & adoacc0l0.Fields("a0k11").Value
'   If IsNull(adoacc0l0.Fields("a0k20").Value) Then
'      adoaccrpt106.Fields("r10605").Value = Null
'   Else
'      adoaccrpt106.Fields("r10605").Value = adoacc0l0.Fields("a0k20").Value & " " & StaffQuery(adoacc0l0.Fields("a0k20").Value)
'   End If
'   PaintLine ReportSum(4)
'   adoaccrpt106.AddNew
'   adoaccrpt106.Fields("r10601").Value = strUserNum
'   adoaccrpt106.Fields("r10602").Value = Counter
'   adoaccrpt106.Fields("r10603").Value = "" & adoacc0l0.Fields("a0k11").Value
'   If IsNull(adoacc0l0.Fields("a0k20").Value) Then
'      adoaccrpt106.Fields("r10605").Value = Null
'   Else
'      adoaccrpt106.Fields("r10605").Value = adoacc0l0.Fields("a0k20").Value & " " & StaffQuery(adoacc0l0.Fields("a0k20").Value)
'   End If
'   adoaccrpt106.Fields("r10611").Value = ReportSum(24)
'   adoaccsum.CursorLocation = adUseClient
'    StrSQLa = "select sum(to_number(nvl(r10614, 0))), sum(to_number(nvl(r10616, 0))), sum(to_number(nvl(r10617, 0))) from accrpt106 where r10601 = '" & strUserNum & "' and r10607 is not null and r10605 = '" & adoacc0l0.Fields("a0k20").Value & " " & StaffQuery(IIf(IsNull(adoacc0l0.Fields("a0k20").Value), MsgText(601), adoacc0l0.Fields("a0k20").Value)) & "' and r10606 = '" & adoacc0l0.Fields("a0k03").Value & "' and r10603 = '" & adoacc0l0.Fields("a0k11").Value & "'"
'   adoaccsum.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      adoaccrpt106.Fields("r10614").Value = IIf("" & adoaccsum.Fields(0).Value = "", 0, adoaccsum.Fields(0).Value)
'      adoaccrpt106.Fields("r10616").Value = IIf("" & adoaccsum.Fields(1).Value = "", 0, adoaccsum.Fields(1).Value)
'      adoaccrpt106.Fields("r10617").Value = IIf("" & adoaccsum.Fields(2).Value = "", 0, adoaccsum.Fields(2).Value)
'   Else
'      adoaccrpt106.Fields("r10614").Value = 0
'      adoaccrpt106.Fields("r10616").Value = 0
'      adoaccrpt106.Fields("r10617").Value = 0
'   End If
'   adoaccsum.Close
'   adoaccrpt106.AddNew
'   adoaccrpt106.Fields("r10601").Value = strUserNum
'   adoaccrpt106.Fields("r10602").Value = Counter
'   adoaccrpt106.Fields("r10603").Value = "" & adoacc0l0.Fields("a0k11").Value
'   If IsNull(adoacc0l0.Fields("a0k20").Value) Then
'      adoaccrpt106.Fields("r10605").Value = Null
'   Else
'      adoaccrpt106.Fields("r10605").Value = adoacc0l0.Fields("a0k20").Value & " " & StaffQuery(adoacc0l0.Fields("a0k20").Value)
'   End If
'   PaintLine ReportSum(4)
'   adoaccrpt106.UpdateBatch
'   adoacc0l0.MoveNext
'End Sub
'
''*************************************************
''  列印類別-選擇收回之智權人員小計計算
''
''*************************************************
'Private Sub SubSelect2_1()
'Dim strSql As String
'Dim StrSQLa As String
'
'   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'      strSql = " and a0l02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'   End If
'   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'      strSql = strSql & " and a0l02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'   End If
'
'   adoacc0l0.MovePrevious
'   adoaccrpt106.AddNew
'   adoaccrpt106.Fields("r10601").Value = strUserNum
'   adoaccrpt106.Fields("r10602").Value = Counter
'   adoaccrpt106.Fields("r10603").Value = "" & adoacc0l0.Fields("a0k11").Value
'   If IsNull(adoacc0l0.Fields("a0k20").Value) Then
'      adoaccrpt106.Fields("r10605").Value = Null
'   Else
'      adoaccrpt106.Fields("r10605").Value = adoacc0l0.Fields("a0k20").Value & " " & StaffQuery(adoacc0l0.Fields("a0k20").Value)
'   End If
'   adoaccrpt106.Fields("r10611").Value = "智權人員小計:"
'   adoaccsum.CursorLocation = adUseClient
'    StrSQLa = "select sum(to_number(nvl(r10614, 0))), sum(to_number(nvl(r10616, 0))), sum(to_number(nvl(r10617, 0))) from accrpt106 where r10601 = '" & strUserNum & "' and r10607 is not null and r10605 = '" & adoacc0l0.Fields("a0k20").Value & " " & StaffQuery(IIf(IsNull(adoacc0l0.Fields("a0k20").Value), MsgText(601), adoacc0l0.Fields("a0k20").Value)) & "' "
'   adoaccsum.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      adoaccrpt106.Fields("r10614").Value = IIf("" & adoaccsum.Fields(0).Value = "", 0, adoaccsum.Fields(0).Value)
'      adoaccrpt106.Fields("r10616").Value = IIf("" & adoaccsum.Fields(1).Value = "", 0, adoaccsum.Fields(1).Value)
'      adoaccrpt106.Fields("r10617").Value = IIf("" & adoaccsum.Fields(2).Value = "", 0, adoaccsum.Fields(2).Value)
'   Else
'      adoaccrpt106.Fields("r10614").Value = 0
'      adoaccrpt106.Fields("r10616").Value = 0
'      adoaccrpt106.Fields("r10617").Value = 0
'   End If
'   adoaccsum.Close
'   adoaccrpt106.AddNew
'   adoaccrpt106.Fields("r10601").Value = strUserNum
'   adoaccrpt106.Fields("r10602").Value = Counter
'   adoaccrpt106.Fields("r10603").Value = "" & adoacc0l0.Fields("a0k11").Value
'   If IsNull(adoacc0l0.Fields("a0k20").Value) Then
'      adoaccrpt106.Fields("r10605").Value = Null
'   Else
'      adoaccrpt106.Fields("r10605").Value = adoacc0l0.Fields("a0k20").Value & " " & StaffQuery(adoacc0l0.Fields("a0k20").Value)
'   End If
'   PaintLine ReportSum(4)
'   adoaccrpt106.UpdateBatch
'   adoacc0l0.MoveNext
'End Sub
'
''*************************************************
''  列印類別-選擇收回之公司別小計計算
''
''*************************************************
'Private Sub SubSelect2_2()
'Dim strSql As String
'Dim StrSQLa As String
'
'   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'      strSql = " and a0l02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'   End If
'   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'      strSql = strSql & " and a0l02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'   End If
'
'   adoacc0l0.MovePrevious
'   adoaccrpt106.AddNew
'   adoaccrpt106.Fields("r10601").Value = strUserNum
'   adoaccrpt106.Fields("r10602").Value = Counter
'   adoaccrpt106.Fields("r10603").Value = "" & adoacc0l0.Fields("a0k11").Value
'   If IsNull(adoacc0l0.Fields("a0k20").Value) Then
'      adoaccrpt106.Fields("r10605").Value = Null
'   Else
'      adoaccrpt106.Fields("r10605").Value = adoacc0l0.Fields("a0k20").Value & " " & StaffQuery(adoacc0l0.Fields("a0k20").Value)
'   End If
'   adoaccrpt106.Fields("r10611").Value = "公司別小計:"
'   adoaccsum.CursorLocation = adUseClient
'    StrSQLa = "select sum(to_number(nvl(r10614, 0))), sum(to_number(nvl(r10616, 0))), sum(to_number(nvl(r10617, 0))) from accrpt106 where r10601 = '" & strUserNum & "' and r10607 is not null and r10605 = '" & adoacc0l0.Fields("a0k20").Value & " " & StaffQuery(IIf(IsNull(adoacc0l0.Fields("a0k20").Value), MsgText(601), adoacc0l0.Fields("a0k20").Value)) & "' and r10603 = '" & adoacc0l0.Fields("a0k11").Value & "'"
'   adoaccsum.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      adoaccrpt106.Fields("r10614").Value = IIf("" & adoaccsum.Fields(0).Value = "", 0, adoaccsum.Fields(0).Value)
'      adoaccrpt106.Fields("r10616").Value = IIf("" & adoaccsum.Fields(1).Value = "", 0, adoaccsum.Fields(1).Value)
'      adoaccrpt106.Fields("r10617").Value = IIf("" & adoaccsum.Fields(2).Value = "", 0, adoaccsum.Fields(2).Value)
'   Else
'      adoaccrpt106.Fields("r10614").Value = 0
'      adoaccrpt106.Fields("r10616").Value = 0
'      adoaccrpt106.Fields("r10617").Value = 0
'   End If
'   adoaccsum.Close
'   adoaccrpt106.AddNew
'   adoaccrpt106.Fields("r10601").Value = strUserNum
'   adoaccrpt106.Fields("r10602").Value = Counter
'   adoaccrpt106.Fields("r10603").Value = "" & adoacc0l0.Fields("a0k11").Value
'   If IsNull(adoacc0l0.Fields("a0k20").Value) Then
'      adoaccrpt106.Fields("r10605").Value = Null
'   Else
'      adoaccrpt106.Fields("r10605").Value = adoacc0l0.Fields("a0k20").Value & " " & StaffQuery(adoacc0l0.Fields("a0k20").Value)
'   End If
'   PaintLine ReportSum(4)
'   adoaccrpt106.UpdateBatch
'   adoacc0l0.MoveNext
'End Sub
'
''*************************************************
''  抬頭列印
''
''*************************************************
'Private Sub PrintHead1()
'   Printer.FontSize = 16
'   Printer.CurrentX = 4000 - m_dblLeftDiff
'   Printer.CurrentY = 1000
'   Printer.Print ReportTitle(106)
'   Printer.CurrentX = 8000 - m_dblLeftDiff
'   Printer.CurrentY = 1000
'   Printer.Print "(未收)"
'   Printer.FontSize = 10
'   Printer.CurrentX = 4500 - m_dblLeftDiff
'   Printer.CurrentY = 1800
'   Printer.Print "帳款日期: "
'   Printer.CurrentX = 5700 - m_dblLeftDiff
'   Printer.CurrentY = 1800
'   Printer.Print MaskEdBox1.Text & " ~ " & MaskEdBox2.Text
'   Printer.CurrentX = 500 - m_dblLeftDiff
'   Printer.CurrentY = 2100
'   Printer.Print "列印人員: "
'   Printer.CurrentX = 1800 - m_dblLeftDiff
'   Printer.CurrentY = 2100
'   Printer.Print StaffQuery(strUserNum)
'   Printer.CurrentX = 10000 - 200 - m_dblLeftDiff
'   Printer.CurrentY = 2100
'   Printer.Print "列印日期: "
'   Printer.CurrentX = 11300 - 400 - m_dblLeftDiff
'   Printer.CurrentY = 2100
'   Printer.Print CFDate(ACDate(ServerDate))
'   Printer.CurrentX = 10000 - 200 - m_dblLeftDiff
'   Printer.CurrentY = 2400
'   Printer.Print "頁　　次: "
'   Printer.CurrentX = 11300 - 400 - m_dblLeftDiff
'   Printer.CurrentY = 2400
'   Printer.Print intPage
'   Printer.CurrentX = 500 - m_dblLeftDiff
'   Printer.CurrentY = 2700
'   Printer.Print "公司別: "
'   Printer.CurrentX = 1300 - m_dblLeftDiff
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt106.Fields("r10603").Value) = False Then
'      Printer.Print adoaccrpt106.Fields("r10603").Value
'      Printer.CurrentX = 1500 - m_dblLeftDiff
'      Printer.CurrentY = 2700
'      Printer.Print A0802Query(adoaccrpt106.Fields("r10603").Value)
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 4500 - m_dblLeftDiff
'   Printer.CurrentY = 2700
'   Printer.Print "部門別: "
'   Printer.CurrentX = 5600 - m_dblLeftDiff
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt106.Fields("r10604").Value) = False Then
'      Printer.Print adoaccrpt106.Fields("r10604").Value
'      Printer.CurrentX = 6300 - m_dblLeftDiff
'      Printer.CurrentY = 2700
'      Printer.Print A0902Query(adoaccrpt106.Fields("r10604").Value)
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 10000 - 200 - m_dblLeftDiff
'   Printer.CurrentY = 2700
'   Printer.Print "智權人員: "
'   Printer.CurrentX = 11100 - 200 - m_dblLeftDiff
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt106.Fields("r10605").Value) = False Then
'      Printer.Print adoaccrpt106.Fields("r10605").Value
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 500 - m_dblLeftDiff
'   Printer.CurrentY = 3300
'   Printer.Print "客戶編號"
'   Printer.CurrentX = 1700 - m_dblLeftDiff
'   Printer.CurrentY = 3300
'   Printer.Print "收據抬頭"
'   Printer.CurrentX = 3500 - m_dblLeftDiff - 400
'   Printer.CurrentY = 3300
'   Printer.Print "收據日期"
'   Printer.CurrentX = 4500 - m_dblLeftDiff - 400
'   Printer.CurrentY = 3300
'   Printer.Print "收據號碼"
'   Printer.CurrentX = 5700 - m_dblLeftDiff - 400
'   Printer.CurrentY = 3300
'   Printer.Print "本所案號"
'   Printer.CurrentX = 7100 - m_dblLeftDiff - 400
'   Printer.CurrentY = 3300
'   Printer.Print "案件性質"
'   Printer.CurrentX = 8100 - m_dblLeftDiff - 400
'   Printer.CurrentY = 3300
'   Printer.Print "申請國家"
'   Printer.CurrentX = 9100 - m_dblLeftDiff - 400
'   Printer.CurrentY = 3300
'   Printer.Print "應收金額"
'   Printer.CurrentX = 10200 - m_dblLeftDiff - 400
'   Printer.CurrentY = 3300
'   Printer.Print "已收金額"
'   Printer.CurrentX = 11300 - m_dblLeftDiff - 400
'   Printer.CurrentY = 3300
'   Printer.Print "未收金額"
'   Printer.Line (500 - m_dblLeftDiff, 3700)-(12500 - m_dblLeftDiff - 400, 3700)
'End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1 = ""
   'Add by Amy 2015/03/03
   'Modify by Amy 2020/04/27 公司別改下拉
   'Modify by Amy 2020/07/08 改回輸入
   Text2 = ""
   Text2.Tag = ""
'   CboComp = ""
'   CboComp.Tag = ""
   'end 2020/07/08
   'end 2020/04/27
   lblSalesName = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text3 = ""
   'Add by Sindy 2010/8/17
   Text4 = ""
   Text5 = ""
   Check1.Value = 0
   Check2.Value = 0
   '2010/8/17 End
   'Check5.Value = 0 'Mark by Amy 2020/03/12 Add by Amy 2016/09/19
   cboTitle.Text = "" 'Add By Sindy 2016/6/13
   'Add by Sindy 2010/9/13
   Text7 = ""
   Text8 = ""
   'Add By Sindy 2016/8/22
   Text9 = ""
   Text10 = ""
   Text11 = ""
   Text12 = ""
   '2016/8/22 END
   MaskEdBox3.Mask = ""
   MaskEdBox3.Text = ""
   MaskEdBox3.Mask = DFormat
   MaskEdBox4.Mask = ""
   MaskEdBox4.Text = ""
   MaskEdBox4.Mask = DFormat
   MaskEdBox5.Mask = ""
   MaskEdBox5.Text = ""
   MaskEdBox5.Mask = DFormat
   MaskEdBox6.Mask = ""
   MaskEdBox6.Text = ""
   MaskEdBox6.Mask = DFormat
   '2010/9/13 End
   Text1.SetFocus
End Sub

'Private Sub Old_PrintSalesListExcel()
'Dim ii As Integer, iIdx As Integer, bolContinue As Boolean, stNames() As String
'Dim iY As Integer, iTabWidth As Integer, iPos As Integer
'Dim i As Integer, strTemp As String
'Dim strText As String
'   stNames = Split(m_strSalesList, vbCrLf)
'   If UBound(stNames) <= 0 Then Exit Sub
'
'   bolContinue = True
'   iIdx = 0
'
'   With wksAnnuity
'   Do While bolContinue
'      intPage = intPage + 1
'      '換頁
'      .Range("A" & intCounter).Select
'      .HPageBreaks.Add Before:=.Application.ActiveCell
'      For i = 1 To 2
'         If i = 1 Then
'            .Range("F" & intCounter).Value = "智權人員帳款明細表 (" & IIf(Trim(Text3) = "1", "未收", "收回") & ")"
'         ElseIf i = 2 Then
'            intCounter = intCounter + 1
'            strText = ""
'            If MaskEdBox1.Text <> MsgText(29) Then
'               If strText <> "" Then strText = strText & "　"
'               strText = strText & "帳款日期：" & MaskEdBox1.Text & "~" & MaskEdBox2.Text
'            End If
'            If Text4 <> "" Then
'               If strText <> "" Then strText = strText & "　"
'               strText = strText & "收據號碼：" & Text4.Text & "~" & Text5.Text
'            End If
'            If Text7 <> "" Then
'               If strText <> "" Then strText = strText & "　"
'               strText = strText & "客戶代號：" & Text7.Text & "~" & Text8.Text
'            End If
'            If MaskEdBox3.Text <> MsgText(29) Then
'               If strText <> "" Then strText = strText & "　"
'               strText = strText & "收文日：" & MaskEdBox3.Text & "~" & MaskEdBox4.Text
'            End If
'            If MaskEdBox5.Text <> MsgText(29) Then
'               If strText <> "" Then strText = strText & "　"
'               strText = strText & "發文日：" & MaskEdBox5.Text & "~" & MaskEdBox6.Text
'            End If
'            .Range("F" & intCounter).Value = strText
'         End If
'         strTemp = "A" & intCounter & ":L" & intCounter
'         .Range(strTemp).Select
'         With .Application.Selection
'             .HorizontalAlignment = xlCenter
'             '.MergeCells = True
'         End With
'         If i = 1 Then
'           With .Application.Selection
'             .Font.Size = 18
'            End With
'         End If
'      Next i
'      intCounter = intCounter + 1
'      .Range("A" & intCounter).Value = "列印日期：" & CFDate(ACDate(ServerDate))
'      intCounter = intCounter + 1
'      .Range("A" & intCounter).Value = "頁　　次：" & intPage
'      intCounter = intCounter + 1
'      .Range("A" & intCounter).Value = "列印人員：" & strUserName
'      intCounter = intCounter + 1
'      .Range("A" & intCounter).Value = "智權人員繳回清單"
'      intCounter = intCounter + 1
'      .Range("A" & intCounter).Value = "智權人員"
'      .Range("B" & intCounter).Value = "繳回日"
'      .Range("C" & intCounter).Value = "智權人員"
'      .Range("D" & intCounter).Value = "繳回日"
'      .Range("E" & intCounter).Value = "智權人員"
'      .Range("F" & intCounter).Value = "繳回日"
'      .Range("G" & intCounter).Value = "智權人員"
'      .Range("H" & intCounter).Value = "繳回日"
'      .Range("I" & intCounter).Value = "智權人員"
'      .Range("J" & intCounter).Value = "繳回日"
'      .Range("K" & intCounter).Value = "智權人員"
'      .Range("L" & intCounter).Value = "繳回日"
'      strTemp = "A" & intCounter & ":L" & intCounter
'      .Range(strTemp).Select
'      With .Application.Selection.Borders(xlEdgeBottom)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'       End With
'       intCounter = intCounter + 1
'      '一頁印162個智權人員
'      For ii = 1 To 162
'         If iIdx < UBound(stNames) Then
'            iPos = iIdx Mod 6
'            '第一欄
'            If iPos = 0 Then
'               .Range("A" & intCounter).Value = stNames(iIdx)
'            '第二欄
'            ElseIf iPos = 1 Then
'               .Range("C" & intCounter).Value = stNames(iIdx)
'            '第三欄
'            ElseIf iPos = 2 Then
'               .Range("E" & intCounter).Value = stNames(iIdx)
'            '第四欄
'            ElseIf iPos = 3 Then
'               .Range("G" & intCounter).Value = stNames(iIdx)
'            '第五欄
'            ElseIf iPos = 4 Then
'               .Range("I" & intCounter).Value = stNames(iIdx)
'            '第六欄
'            ElseIf iPos = 5 Then
'               .Range("K" & intCounter).Value = stNames(iIdx)
'               intCounter = intCounter + 1
'            End If
'         Else
'            bolContinue = False
'            Exit For
'         End If
'         iIdx = iIdx + 1
'      Next
'   Loop
'   End With
'End Sub

'Add By Sindy 2010/5/20
'Public Sub Old_PrintExcelTitle1(strDept As String, strSales)
'Dim i As Integer, strTemp As String
'Dim strText As String
'   intPage = intPage + 1
'   With wksAnnuity
'      For i = 1 To 2
'         If i = 1 Then
'            .Range("G" & intCounter).Value = "智權人員帳款明細表 (未收)"
'         ElseIf i = 2 Then
'            intCounter = intCounter + 1
'            strText = ""
'            If MaskEdBox1.Text <> MsgText(29) Then
'               If strText <> "" Then strText = strText & "　"
'               strText = strText & "帳款日期：" & MaskEdBox1.Text & "~" & MaskEdBox2.Text
'            End If
'            If Text4 <> "" Then
'               If strText <> "" Then strText = strText & "　"
'               strText = strText & "收據號碼：" & Text4.Text & "~" & Text5.Text
'            End If
'            If Text7 <> "" Then
'               If strText <> "" Then strText = strText & "　"
'               strText = strText & "客戶代號：" & Text7.Text & "~" & Text8.Text
'            End If
'            If MaskEdBox3.Text <> MsgText(29) Then
'               If strText <> "" Then strText = strText & "　"
'               strText = strText & "收文日：" & MaskEdBox3.Text & "~" & MaskEdBox4.Text
'            End If
'            If MaskEdBox5.Text <> MsgText(29) Then
'               If strText <> "" Then strText = strText & "　"
'               strText = strText & "發文日：" & MaskEdBox5.Text & "~" & MaskEdBox6.Text
'            End If
'            .Range("G" & intCounter).Value = strText
'         End If
'         '2012/5/2 modify by sonia
'         'strTemp = "A" & intCounter & ":M" & intCounter
'         strTemp = "A" & intCounter & ":Q" & intCounter 'Modify By Sindy 2013/6/14
'         .Range(strTemp).Select
'         With .Application.Selection
'             .HorizontalAlignment = xlCenter
'             '.MergeCells = False
'         End With
'         If i = 1 Then
'           With .Application.Selection
'             .Font.Size = 18
'            End With
'         End If
'      Next i
'      intCounter = intCounter + 1
'      '.Range("K" & intCounter).Value = "頁　次：" & intPage
'      .Range("A" & intCounter).Value = "頁　次：" & intPage
'      intCounter = intCounter + 1
'      '.Range("A" & intCounter).Value = "部門別：" & strDept
'      '.Range("K" & intCounter).Value = "智權人員：" & strSales
'      .Range("A" & intCounter).Value = "智權人員：" & strDept & " " & strSales
'      intCounter = intCounter + 1
'      .Range("A" & intCounter).Value = "收據抬頭"
'      .Range("B" & intCounter).Value = "收據日期"
'      'Add By Sindy 2013/6/14
'      .Range("C" & intCounter).Value = "客戶代號"
'      .Range("D" & intCounter).Value = "印"
'      '2013/6/14 END
'      .Range("E" & intCounter).Value = "收據號碼"
'      'Add By Sindy 2013/12/6
'      .Range("F" & intCounter).Value = "客戶案件案號"
'      '2013/12/6 END
'      .Range("G" & intCounter).Value = "本所案號"
'      .Range("H" & intCounter).Value = "案件性質"
'      .Range("I" & intCounter).Value = "申請國家"
'      .Range("J" & intCounter).Value = "應收金額"
'      .Range("K" & intCounter).Value = "已收金額"
'      .Range("L" & intCounter).Value = "未收金額"
'      .Range("M" & intCounter).Value = "案件規費餘額"
'      .Range("N" & intCounter).Value = "備註"
'      .Range("O" & intCounter).Value = "點數"
'      .Range("P" & intCounter).Value = "發文日"
'      .Range("Q" & intCounter).Value = "預定收款日"  '2012/5/2 add by sonia
'      '2012/5/2 modify by sonia
'      'strTemp = "A" & intCounter & ":M" & intCounter
'      strTemp = "A" & intCounter & ":Q" & intCounter 'Modify By Sindy 2013/6/14
'      .Range(strTemp).Select
'      With .Application.Selection.Borders(xlEdgeBottom)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'       End With
'       intCounter = intCounter + 1
'   End With
'End Sub

'Add by Amy 2020/04/27 +L 公司,程式調整
'*************************************************
' 產生Excel資料(未收)
'
'*************************************************
Public Sub PrintExcel1()
Dim strStart As Integer, CountData_S As Integer
Dim bolIsFirst As Boolean
Dim strQuery As String, strA As String, strCmp As String
Dim intMaxTitle As Integer 'Add by Amy 2022/07/04

On Error GoTo ErrHnd
   
   'Modify by Amy 2020/07/08 公司別改回輸入(不用下拉)
    strCmp = Trim(Text2) 'Trim(CboComp)
'    If strCmp <> MsgText(60) Then
'        If InStr(strCmp, "　") > 0 Then
'            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
'        End If
'    End If
    If strCmp <> MsgText(601) Then
'        If strCmp <> "J" And strCmp <> "L" Then
'            strA = " And r10603 Not In ('J','L') "
'        Else
            strA = " And r10603='" & strCmp & "' "
'        End If
    End If
    
    'strA = "Select Distinct Decode(r10603,'J','J','L','L','1') as Cmp From accrpt106 Where r10601 = '" & strUserNum & "' " & strA
    strA = "Select Distinct r10603 as Cmp From accrpt106 Where r10601 = '" & strUserNum & "' " & strA
    'Add by Amy 2022/07/04 法律所再產生一份以介紹人為主的資料
    If Trim(Text2) = MsgText(601) Or Trim(Text2) = "L" Then
        If ChkLaw = True Then
            strA = strA & " Union Select 'LZ' as Cmp From Dual Order by Cmp "
        End If
    End If
    'end 2020/07/08
    If rsA.State = adStateOpen Then rsA.Close
    rsA.CursorLocation = adUseClient
    rsA.Open strA, adoTaie, adOpenDynamic, adLockBatchOptimistic
    If rsA.RecordCount = 0 Then
        rsA.Close
        Exit Sub
    End If
    
    ReDim strFieldN(20)
    ReDim intWidth(20)
    'Modify by Amy 2020/05/28 +申請案號
    'Modify by Amy 2022/07/04 +介紹人
    strFieldN = Array("收據抬頭", "公司", "收據日期", "客戶代號", "印", _
                                "介紹人", "收據號碼", "客戶案件案號", "申請案號", "本所案號", _
                                 "案件性質", "申請國家", "案件名稱", "應收金額", "應收服務費", _
                                "應收規費", "已收金額", "未收金額", "案件規費餘額", "備註", _
                                "點數", "扣繳", "發文日", "付款週期月份")
    intWidth = Array(13, 4.5, 10, 10, 5, 8, 10, 13, 10, 10, _
                                10, 10, 13, 10, 10, 10, 10, 10, 10, 14, _
                                13, 8, 10, 12)
    'end 2022/07/04
    'end 2020/05/28
    bolIsFirst = True: intXlsSheet = 1: strText = "": CountData_S = 0: intField = 65: bolHasData = True
    intDelF = 0 'Add by Amy 2020/06/30
    
    rsA.MoveFirst
    For i = 0 To rsA.RecordCount - 1
        '*** 設定Excel檔名 ***
        If i = 0 Then
            strFileName = strExcelPath & "智權人員帳款明細-未收" & Format(Now, "yyyymmddhhmmss") & MsgText(43)
            If Dir(strFileName) = MsgText(601) Then
               If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
                  MkDir strExcelPath
               End If
            Else
               Kill strFileName
            End If
            xlsAnnuity.SheetsInNewWorkbook = 3
            xlsAnnuity.Workbooks.add
            xlsAnnuity.Application.WindowState = xlMinimized
        End If
        
        If strWkName = MsgText(601) Then strWkName = Left(xlsAnnuity.Worksheets(1).Name, Len(xlsAnnuity.Worksheets(1).Name) - 1)
        'Add by Amy 2020/05/28 畫面公司別為空白=全部公司,若全公司都有資料需有6個Sheet,故需加工作表
        'Modify by Amy 2021/05/13 2010需增加工作表會錯,目前財務只有2010及2013版本 拿掉 And intXlsSheet <> 1 And Val(xlsAnnuity.Version) = 15
        If intXlsSheet > 3 Then
            'Modify by Amy 2022/07/04 +After:=wksAnnuity 加於最後
            xlsAnnuity.Worksheets.add After:=wksAnnuity '插入sheet
        End If
        Set wksAnnuity = xlsAnnuity.Worksheets(strWkName & intXlsSheet)
        wksAnnuity.Activate
        'xlsAnnuity.Visible = True
    
        '*** 抓取各公司別資料 ***
        'Modify by Amy 2022/07/04 法律所再產生一份以介紹人為主的資料
        If "" & rsA.Fields("Cmp") = "LZ" Then
            '介紹人以目前區為主-瑞婷
            strQuery = "Select r10603,r10606,r10607,r10608,r10610,r10611,r10612,r10613,r10614,r10615," & _
                                "r10616,r10617,r10618,r10619,r10620,r10621,r10622,r10623,r10624,r10625,r10626," & _
                                "r10630,r10631 as r10605,st15 as r10604,'' as st02,'' as r10631 From accrpt106,Staff " & _
                                "Where r10601 = '" & strUserNum & "' And r10631=st01(+) And r10631 Is Not Null "
            '勾選「客戶案件案號」
            If Check4.Value = 1 Then
                strQuery = strQuery & " Order by st15 asc,r10631,r10624 asc,r10602 asc"
            Else
                strQuery = strQuery & " Order by st15 asc,r10631,r10602 asc"
            End If
        Else
            'Modify by Amy 2020/07/28 每一個公司一個sheet顯示-瑞婷
    '        If "" & rsA.Fields("Cmp") = "1" Then
    '            strQuery = " And r10603<>'J' And r10603<>'L' "
    '        Else
                strQuery = " And r10603='" & rsA.Fields("Cmp") & "' "
    '        End If
            'end 2020/07/28
            'Modify by Amy 2022/07/04 +介紹人
            strQuery = "Select accrpt106.*,st02 From accrpt106,Staff Where r10601 = '" & strUserNum & "' And r10631=st01(+) " & strQuery
            '勾選「客戶案件案號」
            If Check4.Value = 1 Then
                strQuery = strQuery & " Order by r10624 asc,r10601 asc,r10602 asc"
            Else
                strQuery = strQuery & " Order by r10601 asc,r10602 asc"
            End If
        End If
        'end 2022/07/04
        If adoaccrpt106.State = adStateOpen Then adoaccrpt106.Close
        adoaccrpt106.CursorLocation = adUseClient
        adoaccrpt106.Open strQuery, adoTaie, adOpenDynamic, adLockBatchOptimistic
        With wksAnnuity
            If adoaccrpt106.RecordCount > 0 Then
                strSalesNo = "": strSameName = "":  m_strSalesList = "": strOldArea = "": CountData_S = 0
                intCounter = 1
                adoaccrpt106.MoveFirst
                Do While adoaccrpt106.EOF = False
                    If strSalesNo <> "" & adoaccrpt106.Fields("r10605").Value Then
                        m_strSalesList = m_strSalesList & "" & adoaccrpt106.Fields("r10605").Value & vbCrLf
                        '顯示員編
                        If strSalesNo = "" Then
                            'Modify by Amy 2020/04/27 +公司別
                            Call SetExcel(1, bolIsFirst, A0902Query("" & adoaccrpt106.Fields("r10604").Value) & StaffQuery("" & adoaccrpt106.Fields("r10605").Value) & " " & adoaccrpt106.Fields("r10605").Value, "" & rsA.Fields("Cmp"))
                            strStart = intCounter '合計公式
                        End If
                    End If
                    '智權人員/業務區不同也換頁(ex:86052 104年及105年不同區的資料)
                    'Modify by Amy 2022/07/04 改變數 intCounter > 6/CountData_S = 27 ->改顯示資料25列
                    intMaxTitle = intTitleR_Fix + intTitleR + 1
                    If (intCounter > intMaxTitle + 1 And strSalesNo <> "" & adoaccrpt106.Fields("r10605").Value) Or (intCounter > intMaxTitle + 1 And strOldArea <> "" & adoaccrpt106.Fields("r10604").Value) _
                        Or CountData_S = 32 - (intMaxTitle + 1) Then
                        If intCounter > intMaxTitle + 1 And (strSalesNo <> "" & adoaccrpt106.Fields("r10605").Value Or strOldArea <> "" & adoaccrpt106.Fields("r10604").Value) Then
                    'end 2022/07/04
                            '合計
                            Call PrintSum1(strStart)
                            intCounter = intCounter + 1
                            strStart = intCounter
                        End If
        
                        '換頁
                        .Range("A" & intCounter).Select
                        .HPageBreaks.add Before:=.Application.ActiveCell
                        'Modify by Amy 2020/04/27 +公司別
                        Call SetExcel(1, , A0902Query("" & adoaccrpt106.Fields("r10604").Value) & StaffQuery("" & adoaccrpt106.Fields("r10605").Value) & " " & adoaccrpt106.Fields("r10605").Value, "" & rsA.Fields("Cmp"))
                        CountData_S = 0
                    End If
                    '收據抬頭不同時
                    If IsNull(adoaccrpt106.Fields("r10607").Value) = False Then
                        .Range(Chr(GetValue("收據抬頭") + 65) & intCounter).Value = Left("" & adoaccrpt106.Fields("r10607").Value, 6)
                    End If
                    strSameName = ("" & adoaccrpt106.Fields("r10607").Value)
                    If IsNull(adoaccrpt106.Fields("r10603").Value) = False Then
                        .Range(Chr(GetValue("公司") + 65) & intCounter).Value = adoaccrpt106.Fields("r10603").Value
                        .Range(Chr(GetValue("公司") + 65) & intCounter).HorizontalAlignment = xlCenter
                    End If
                    If IsNull(adoaccrpt106.Fields("r10608").Value) = False Then
                        .Range(Chr(GetValue("收據日期") + 65) & intCounter).Value = CFDate(adoaccrpt106.Fields("r10608").Value)
                    End If
                    If IsNull(adoaccrpt106.Fields("r10606").Value) = False Then
                        .Range(Chr(GetValue("客戶代號") + 65) & intCounter).Value = adoaccrpt106.Fields("r10606").Value
                    End If
                    If IsNull(adoaccrpt106.Fields("r10623").Value) = False Then
                       .Range(Chr(GetValue("印") + 65) & intCounter).Value = "N"
                    End If
                    'Add by Amy 2022/07/04 公司別為L顯示,介紹人
                    If "" & rsA.Fields("Cmp") = "L" Then
                        If IsNull(adoaccrpt106.Fields("r10631").Value) = False Then
                            .Range(Chr(GetValue("介紹人") + 65) & intCounter).Value = "" & adoaccrpt106.Fields("st02").Value
                        End If
                    End If
                    'end 2022/07/04
                    If IsNull(adoaccrpt106.Fields("r10610").Value) = False Then
                       .Range(Chr(GetValue("收據號碼") + 65) & intCounter).Value = adoaccrpt106.Fields("r10610").Value
                    End If
                    If IsNull(adoaccrpt106.Fields("r10624").Value) = False Then
                       .Range(Chr(GetValue("客戶案件案號") + 65) & intCounter).Value = adoaccrpt106.Fields("r10624").Value
                    End If
                    'Add by Amy 2020/05/28 +申請案號
                    If IsNull(adoaccrpt106.Fields("r10630").Value) = False Then
                       .Range(Chr(GetValue("申請案號") + 65) & intCounter).NumberFormatLocal = "@"
                       .Range(Chr(GetValue("申請案號") + 65) & intCounter).Value = adoaccrpt106.Fields("r10630").Value
                    End If
                    'end 2020/05/28
                    If IsNull(adoaccrpt106.Fields("r10611").Value) = False Then
                       .Range(Chr(GetValue("本所案號") + 65) & intCounter).Value = adoaccrpt106.Fields("r10611").Value
                    End If
                    If IsNull(adoaccrpt106.Fields("r10612").Value) = False Then
                       .Range(Chr(GetValue("案件性質") + 65) & intCounter).Value = adoaccrpt106.Fields("r10612").Value 'Modify by Amy 2024/12/24 原只取4個字
                    End If
                    If IsNull(adoaccrpt106.Fields("r10613").Value) = False Then
                       .Range(Chr(GetValue("申請國家") + 65) & intCounter).Value = Left(adoaccrpt106.Fields("r10613").Value, 4)
                    End If
                    .Range(Chr(GetValue("案件名稱") + 65) & intCounter).Value = StrToStr(GetCaseName("" & adoaccrpt106.Fields("r10611").Value), 20)
                    
                    'Excel資料的數字為0時以 0顯示之(DDollar => DDollar2)
                    If IsNull(adoaccrpt106.Fields("r10614").Value) = False Then
                       .Range(Chr(GetValue("應收金額") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10614").Value, DDollar2)
                    Else
                       .Range(Chr(GetValue("應收金額") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
                    End If
                    If IsNull(adoaccrpt106.Fields("r10625").Value) = False Then
                       .Range(Chr(GetValue("應收服務費") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10625").Value, DDollar2)
                    Else
                       .Range(Chr(GetValue("應收服務費") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
                    End If
                    If IsNull(adoaccrpt106.Fields("r10617").Value) = False Then
                       .Range(Chr(GetValue("應收規費") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10617").Value, DDollar2)
                    Else
                       .Range(Chr(GetValue("應收規費") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
                    End If
                    If IsNull(adoaccrpt106.Fields("r10626").Value) = False Then
                       .Range(Chr(GetValue("扣繳") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10626").Value, DDollar2)
                    Else
                       .Range(Chr(GetValue("扣繳") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
                    End If
                    If IsNull(adoaccrpt106.Fields("r10615").Value) = False Then
                       .Range(Chr(GetValue("已收金額") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10615").Value, DDollar2)
                    Else
                       .Range(Chr(GetValue("已收金額") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
                    End If
                    If IsNull(adoaccrpt106.Fields("r10616").Value) = False Then
                       .Range(Chr(GetValue("未收金額") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10616").Value, DDollar2)
                    Else
                       .Range(Chr(GetValue("未收金額") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
                    End If
                    If IsNull(adoaccrpt106.Fields("r10618").Value) = False Then
                       .Range(Chr(GetValue("案件規費餘額") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10618").Value, DDollar2)
                    Else
                       .Range(Chr(GetValue("案件規費餘額") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
                    End If
                    
                    If IsNull(adoaccrpt106.Fields("r10619").Value) = False Then
                       .Range(Chr(GetValue("備註") + 65) & intCounter).Value = adoaccrpt106.Fields("r10619").Value
                    End If
                    If IsNull(adoaccrpt106.Fields("r10620").Value) = False Then
                       .Range(Chr(GetValue("點數") + 65) & intCounter).Value = Val(adoaccrpt106.Fields("r10620").Value)
                    End If
                    If IsNull(adoaccrpt106.Fields("r10621").Value) = False Then
                       .Range(Chr(GetValue("發文日") + 65) & intCounter).Value = CFDate(Val(adoaccrpt106.Fields("r10621").Value) - 19110000)
                    End If
                    If IsNull(adoaccrpt106.Fields("r10622").Value) = False Then
                        .Range(Chr(GetValue("付款週期月份") + 65) & intCounter).Value = Val(adoaccrpt106.Fields("r10622").Value)
                    End If
                    intCounter = intCounter + 1
                    CountData_S = CountData_S + 1 '計算智權人員明細筆數 for 換頁
                    strSalesNo = adoaccrpt106.Fields("r10605").Value
                    strOldArea = adoaccrpt106.Fields("r10604").Value '業務區
                    adoaccrpt106.MoveNext
                Loop
                '合計
                Call PrintSum1(strStart) '合計公式
            End If
            'Add by Amy 2022/07/04 公司別非L公司,刪除「介紹人」欄
            If "" & rsA.Fields("Cmp") <> "L" Then
                .Range(Chr(GetValue("介紹人") + 65 - intDelF) & ":" & Chr(GetValue("介紹人") + 65 - intDelF)).Delete Shift:=xlToLeft
                intDelF = intDelF + 1
            End If
            'Modify by Amy 2020/06/30 加- intDelF 因欄位全顯示,若未勾「客戶案件案號」及「申請案號」欄,刪申請案號欄時會刪錯
            If Check4.Value = 0 Then
                .Range(Chr(GetValue("客戶案件案號") + 65 - intDelF) & ":" & Chr(GetValue("客戶案件案號") + 65 - intDelF)).Delete Shift:=xlToLeft
                intDelF = intDelF + 1
            End If
            'Add by Amy 2020/05/28 +申請案號
            If Check5.Value = 0 Then
                .Range(Chr(GetValue("申請案號") + 65 - intDelF) & ":" & Chr(GetValue("申請案號") + 65 - intDelF)).Delete Shift:=xlToLeft
                intDelF = intDelF + 1
            End If
            'end 2020/06/30
        End With
        'Modify by Amy 2022/07/04 固定抬頭改抓變數 /+ rsA.Fields("Cmp") = "LZ" 判斷/SetExcel(3)
        Call SetExcel(3)
        wksAnnuity.Range(Chr(intField) & intTitleR_Fix + 1 & ":" & Chr(UBound(strFieldN) + intField) & intCounter).Font.Size = 12
        If "" & rsA.Fields("Cmp") = "LZ" Then
            wksAnnuity.Name = CompNameQuery("L", 4) & "-介紹人"
        Else
            wksAnnuity.Name = CompNameQuery("" & rsA.Fields("Cmp"), 4) 'Modify by Amy 2020/07/08 要可輸2公司 原:A0802Query("" & rsA.Fields("Cmp"), True)
        End If
        'end 2022/07/04
         intXlsSheet = intXlsSheet + 1
         intDelF = 0 'Add by Amy 2020/07/28 每印完一個公司欄位刪除欄位要設0
         bolIsFirst = False
         '繳回清單
         If Text1 = "" And m_strSalesList <> "" Then
            intCounter = 1
            Call PrintSalesListExcel("" & rsA.Fields("Cmp"))
         End If
        rsA.MoveNext
    Next i

    '判斷版本
    If Val(xlsAnnuity.Version) < 12 Then
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
    Else
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
    End If

    xlsAnnuity.Workbooks.Close
    xlsAnnuity.Quit
    'Modify by Amy 2021/06/22 路徑改中文字顯示
    If bolHasData = True Then MsgBox "Excel檔已產生!" & vbCrLf & vbCrLf & strExcelPathN & Replace(strFileName, strExcelPath, ""), vbInformation, Me.Caption
    adoaccrpt106.Close
    Set xlsAnnuity = Nothing
    Set wksAnnuity = Nothing
    Exit Sub
   
ErrHnd:
    adoaccrpt106.Close
    '判斷版本
    If Val(xlsAnnuity.Version) < 12 Then
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
    Else
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
    End If
    xlsAnnuity.Workbooks.Close
    xlsAnnuity.Quit
    Set xlsAnnuity = Nothing
    Set wksAnnuity = Nothing
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Public Sub PrintExcel1_Old()
''Add by Amy 2016/09/19 合計改公式
''Dim strTot14 As String, strTot15 As String, strTot16 As String, strTot18 As String, strTot20 As String
'Dim strStart As Integer
''end 2016/09/19
'Dim strQuery As String, bolIsFirst As Boolean, CountData_S As Integer
'
'On Error GoTo ErrHnd
'
'    bolIsFirst = True: intXlsSheet = 1: strText = "": CountData_S = 0
'
'    If Text2 = "" Then
'        '未輸公司別 台一及智權分開顯示,先產生台一的資料
'        Text2.Tag = "1"
'    Else
'        Text2.Tag = IIf(Text2 = "2", "J", "1")
'    End If
'
'    strQuery = "Select * From accrpt106 Where r10601 = '" & strUserNum & "' And r10603" & IIf(Text2.Tag = "J", "='J'", "<>'J'")
'    If Check4.Value = 1 Then
'        strQuery = strQuery & " Order by r10624 asc,r10601 asc,r10602 asc"
'    Else
'        strQuery = strQuery & " Order by r10601 asc,r10602 asc"
'    End If
'
'    If adoaccrpt106.State = adStateOpen Then adoaccrpt106.Close
'    adoaccrpt106.CursorLocation = adUseClient
'    adoaccrpt106.Open strQuery, adoTaie, adOpenDynamic, adLockBatchOptimistic
'
'    If adoaccrpt106.RecordCount = 0 Then
'        adoaccrpt106.Close
'        MsgBox "無資料產生！" 'Add by Amy 2017/06/08
'        Exit Sub
'    End If
'
'    'Modify by Amy 2016/09/19 +應收服務費/應收規費/扣繳
'    ReDim strFieldN(20)
'    ReDim intWidth(20)
'    'Modified by Lydia 2018/09/12 預定收款日=> 付款週期月份
'    strFieldN = Array("收據抬頭", "公司", "收據日期", "客戶代號", "印", _
'                                "收據號碼", "客戶案件案號", "本所案號", "案件性質", "申請國家", _
'                                "案件名稱", "應收金額", "應收服務費", "應收規費", "扣繳", _
'                                "已收金額", "未收金額", "案件規費餘額", "備註", "點數", _
'                                "發文日", "付款週期月份")
'    intWidth = Array(13, 4.5, 10, 10, 5, 10, 13, 10, 10, 10, _
'                                13, 10, 10, 10, 10, 10, 10, 14, 13, 8, _
'                                10, 12)
'    'end 2016/09/19
'    intField = 65: bolHasData = True
'
'    strFileName = strExcelPath & "智權人員帳款明細-未收" & Format(Now, "yyyymmddhhmmss") & MsgText(43)
'   If Dir(strFileName) = MsgText(601) Then
'      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
'         MkDir strExcelPath
'      End If
'   Else
'      Kill strFileName
'   End If
'
'NextComp:
'
'    strSalesNo = "": strSameName = "":  m_strSalesList = ""
'    strOldArea = "": CountData_S = 0 'Add by Amy 2016/09/19
'    intCounter = 1
'    'strTot14 = 0: strTot15 = 0: strTot16 = 0: strTot18 = 0: strTot20 = 0 'Mark by Amy 2016/09/19 合計改公式
'
'    If bolIsFirst = True Then
'        xlsAnnuity.SheetsInNewWorkbook = 3 'Added by Lydia 2019/04/08 預設工作表數量
'        xlsAnnuity.Workbooks.add
'        xlsAnnuity.Application.WindowState = xlMinimized
'    End If
'
'    'Modify by Amy 2017/09/25 for 工作表名稱改為中文
'    If strWkName = MsgText(601) Then strWkName = Left(xlsAnnuity.Worksheets(1).Name, Len(xlsAnnuity.Worksheets(1).Name) - 1)
'    Set wksAnnuity = xlsAnnuity.Worksheets(strWkName & intXlsSheet)
'    'end 2017/09/25
'    wksAnnuity.Activate
'
'    With wksAnnuity
'        If adoaccrpt106.RecordCount > 0 Then
'            '逐筆填值
'            Do While adoaccrpt106.EOF = False
'                If strSalesNo <> "" & adoaccrpt106.Fields("r10605").Value Then
'                    m_strSalesList = m_strSalesList & "" & adoaccrpt106.Fields("r10605").Value & vbCrLf
'                    'Modify by Amy 2016/04/27 +顯示員編
'                    If strSalesNo = "" Then
'                        Call SetExcel(1, bolIsFirst, A0902Query("" & adoaccrpt106.Fields("r10604").Value) & StaffQuery("" & adoaccrpt106.Fields("r10605").Value) & " " & adoaccrpt106.Fields("r10605").Value)
'                        strStart = intCounter 'Add by Amy 2016/09/19 合計改公式
'                    End If
'                End If
'                '智權人員不同時換頁
'                'Modfiy by Amy 2016/04/27 +業務區不同也換頁(ex:86052 104年及105年不同區的資料)
'                If (intCounter > 6 And strSalesNo <> "" & adoaccrpt106.Fields("r10605").Value) Or (intCounter > 6 And strOldArea <> "" & adoaccrpt106.Fields("r10604").Value) _
'                  Or CountData_S = 27 Then
'                    If intCounter > 6 And (strSalesNo <> "" & adoaccrpt106.Fields("r10605").Value Or strOldArea <> "" & adoaccrpt106.Fields("r10604").Value) Then
'                        '合計
'                        'Modify by Amy 2016/09/19 合計改公式
'                        'Call PrintSum1(strTot14, strTot15, strTot16, strTot18, strTot20, strTot21)
'                        'strTot14 = 0: strTot15 = 0: strTot16 = 0: strTot18 = 0: strTot20 = 0
'                        Call PrintSum1(strStart)
'                        intCounter = intCounter + 1
'                        strStart = intCounter  'Add by Amy 2016/09/19
'                    End If
'
'                    '換頁
'                    .Range("A" & intCounter).Select
'                    .HPageBreaks.add Before:=.Application.ActiveCell
'                    Call SetExcel(1, , A0902Query("" & adoaccrpt106.Fields("r10604").Value) & StaffQuery("" & adoaccrpt106.Fields("r10605").Value) & " " & adoaccrpt106.Fields("r10605").Value)
'                    CountData_S = 0
'                End If
'
'                '若收據抬頭不同時
'                If IsNull(adoaccrpt106.Fields("r10607").Value) = False Then
'                   .Range(Chr(GetValue("收據抬頭") + 65) & intCounter).Value = Left("" & adoaccrpt106.Fields("r10607").Value, 6)
'                End If
'                strSameName = ("" & adoaccrpt106.Fields("r10607").Value)
'                'Add by Amy 2016/09/19 +公司
'                If IsNull(adoaccrpt106.Fields("r10603").Value) = False Then
'                    .Range(Chr(GetValue("公司") + 65) & intCounter).Value = adoaccrpt106.Fields("r10603").Value
'                    .Range(Chr(GetValue("公司") + 65) & intCounter).HorizontalAlignment = xlCenter
'                End If
'                If IsNull(adoaccrpt106.Fields("r10608").Value) = False Then
'                    .Range(Chr(GetValue("收據日期") + 65) & intCounter).Value = CFDate(adoaccrpt106.Fields("r10608").Value)
'                End If
'                If IsNull(adoaccrpt106.Fields("r10606").Value) = False Then
'                    .Range(Chr(GetValue("客戶代號") + 65) & intCounter).Value = adoaccrpt106.Fields("r10606").Value
'                End If
'                If IsNull(adoaccrpt106.Fields("r10623").Value) = False Then
'                   .Range(Chr(GetValue("印") + 65) & intCounter).Value = "N"
'                End If
'                If IsNull(adoaccrpt106.Fields("r10610").Value) = False Then
'                   .Range(Chr(GetValue("收據號碼") + 65) & intCounter).Value = adoaccrpt106.Fields("r10610").Value
'                End If
'                If IsNull(adoaccrpt106.Fields("r10624").Value) = False Then
'                   .Range(Chr(GetValue("客戶案件案號") + 65) & intCounter).Value = adoaccrpt106.Fields("r10624").Value
'                End If
'                If IsNull(adoaccrpt106.Fields("r10611").Value) = False Then
'                   .Range(Chr(GetValue("本所案號") + 65) & intCounter).Value = adoaccrpt106.Fields("r10611").Value
'                End If
'                If IsNull(adoaccrpt106.Fields("r10612").Value) = False Then
'                   .Range(Chr(GetValue("案件性質") + 65) & intCounter).Value = Left(adoaccrpt106.Fields("r10612").Value, 4)
'                End If
'                If IsNull(adoaccrpt106.Fields("r10613").Value) = False Then
'                   .Range(Chr(GetValue("申請國家") + 65) & intCounter).Value = Left(adoaccrpt106.Fields("r10613").Value, 4)
'                End If
'                .Range(Chr(GetValue("案件名稱") + 65) & intCounter).Value = StrToStr(GetCaseName("" & adoaccrpt106.Fields("r10611").Value), 20)
'
'                'Excel資料的數字為0時以 0顯示之(DDollar => DDollar2)
'                If IsNull(adoaccrpt106.Fields("r10614").Value) = False Then
'                   .Range(Chr(GetValue("應收金額") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10614").Value, DDollar2)
'                   'strTot14 = strTot14 + Val(adoaccrpt106.Fields("r10614").Value)
'                Else
'                   .Range(Chr(GetValue("應收金額") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
'                End If
'                'Add by Amy 2016/09/19 +應收服務費/應收規費/扣繳
'                If IsNull(adoaccrpt106.Fields("r10625").Value) = False Then
'                   .Range(Chr(GetValue("應收服務費") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10625").Value, DDollar2)
'                Else
'                   .Range(Chr(GetValue("應收服務費") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
'                End If
'                If IsNull(adoaccrpt106.Fields("r10617").Value) = False Then
'                   .Range(Chr(GetValue("應收規費") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10617").Value, DDollar2)
'                Else
'                   .Range(Chr(GetValue("應收規費") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
'                End If
'                If IsNull(adoaccrpt106.Fields("r10626").Value) = False Then
'                   .Range(Chr(GetValue("扣繳") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10626").Value, DDollar2)
'                Else
'                   .Range(Chr(GetValue("扣繳") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
'                End If
'                'end 2016/09/19
'                If IsNull(adoaccrpt106.Fields("r10615").Value) = False Then
'                   .Range(Chr(GetValue("已收金額") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10615").Value, DDollar2)
'                   'strTot15 = strTot15 + Val(adoaccrpt106.Fields("r10615").Value)
'                Else
'                   .Range(Chr(GetValue("已收金額") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
'                End If
'                If IsNull(adoaccrpt106.Fields("r10616").Value) = False Then
'                   .Range(Chr(GetValue("未收金額") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10616").Value, DDollar2)
'                   'strTot16 = strTot16 + Val(adoaccrpt106.Fields("r10616").Value)
'                Else
'                   .Range(Chr(GetValue("未收金額") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
'                End If
'                If IsNull(adoaccrpt106.Fields("r10618").Value) = False Then
'                   .Range(Chr(GetValue("案件規費餘額") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10618").Value, DDollar2)
'                   'strTot18 = strTot18 + Val(adoaccrpt106.Fields("r10618").Value)
'                Else
'                   .Range(Chr(GetValue("案件規費餘額") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
'                End If
'
'                If IsNull(adoaccrpt106.Fields("r10619").Value) = False Then
'                   .Range(Chr(GetValue("備註") + 65) & intCounter).Value = adoaccrpt106.Fields("r10619").Value
'                End If
'                If IsNull(adoaccrpt106.Fields("r10620").Value) = False Then
'                   .Range(Chr(GetValue("點數") + 65) & intCounter).Value = Val(adoaccrpt106.Fields("r10620").Value)
'                   'strTot20 = strTot20 + Val(adoaccrpt106.Fields("r10620").Value)
'                End If
'                If IsNull(adoaccrpt106.Fields("r10621").Value) = False Then
'                   .Range(Chr(GetValue("發文日") + 65) & intCounter).Value = CFDate(Val(adoaccrpt106.Fields("r10621").Value) - 19110000)
'                End If
'                If IsNull(adoaccrpt106.Fields("r10622").Value) = False Then
'                    'Modified by Lydia 2018/09/12 預定收款日=>付款週期月份
'                   '.Range(Chr(GetValue("預定收款日") + 65) & intCounter).Value = CFDate(Val(adoaccrpt106.Fields("r10622").Value) - 19110000)
'                    .Range(Chr(GetValue("付款週期月份") + 65) & intCounter).Value = Val(adoaccrpt106.Fields("r10622").Value)
'                End If
'                intCounter = intCounter + 1
'                CountData_S = CountData_S + 1 '計算智權人員明細筆數 for 換頁
'                strSalesNo = adoaccrpt106.Fields("r10605").Value
'                strOldArea = adoaccrpt106.Fields("r10604").Value 'Add by Amy 2016/04/27 +業務區
'                adoaccrpt106.MoveNext
'            Loop
'            '合計
'            Call PrintSum1(strStart) 'Modify by Amy 2016/09/19 合計改公式
'        Else
'            '沒資料只印表頭
'            Call SetExcel(1, True, "")
'        End If
'        'Add by Amy 2016/09/19
'        If Check5.Value = 1 Then
'            .Range(Chr(GetValue("應收金額") + 65) & ":" & Chr(GetValue("應收金額") + 65)).Delete Shift:=xlToLeft
'        Else
'            .Range(Chr(GetValue("應收規費") + 65) & ":" & Chr(GetValue("應收規費") + 65)).Delete Shift:=xlToLeft
'            .Range(Chr(GetValue("應收服務費") + 65) & ":" & Chr(GetValue("應收服務費") + 65)).Delete Shift:=xlToLeft
'        End If
'        If Check4.Value = 0 Then
'            .Range(Chr(GetValue("客戶案件案號") + 65) & ":" & Chr(GetValue("客戶案件案號") + 65)).Delete Shift:=xlToLeft
'        End If
'    End With
'
'    intXlsSheet = intXlsSheet + 1
'    '公司別為空先跑台一,再跑智權
'    If Trim(Text2) = "" Then
'        If bolIsFirst = True Then
''            If Check4.Value = 0 Then
''                wksAnnuity.Range(Chr(GetValue("客戶案件案號") + 65) & ":" & Chr(GetValue("客戶案件案號") + 65)).Delete Shift:=xlToLeft
''            End If
'            wksAnnuity.Range(Chr(intField) & "2:" & Chr(UBound(strFieldN) + intField) & intCounter).Font.Size = 12
'            wksAnnuity.Name = "台一"
'            bolIsFirst = False
'            '繳回清單
'            If Text1 = "" And m_strSalesList <> "" Then
'                intCounter = 1
'                Call PrintSalesListExcel
'            End If
'
'            '智權公司
'            strQuery = "Select * From accrpt106 Where r10601 = '" & strUserNum & "' And r10603='J' "
'            If Check4.Value = 1 Then
'                strQuery = strQuery & "Order by r10624 asc,r10601 asc,r10602 asc"
'            Else
'                strQuery = strQuery & "Order by r10601 asc,r10602 asc"
'            End If
'            If adoaccrpt106.State <> adStateClosed Then adoaccrpt106.Close
'            adoaccrpt106.CursorLocation = adUseClient
'            adoaccrpt106.Open strQuery, adoTaie, adOpenDynamic, adLockBatchOptimistic
'            Text2.Tag = "J"
'            GoTo NextComp
'        Else
'            '沒輸公司別 跑J公司清單
''            If Check4.Value = 0 Then
''                wksAnnuity.Range(Chr(GetValue("客戶案件案號") + 65) & ":" & Chr(GetValue("客戶案件案號") + 65)).Delete Shift:=xlToLeft
''            End If
'            wksAnnuity.Range(Chr(intField) & "2:" & Chr(UBound(strFieldN) + intField) & intCounter).Font.Size = 12
'            '工作表更名
'            wksAnnuity.Name = "智權"
'            '繳回清單
'            If Text1 = "" And m_strSalesList <> "" Then
'                intCounter = 1
'                PrintSalesListExcel
'            End If
'        End If
'    Else
'        '有輸公司別
'        If Check4.Value = 0 Then
'            wksAnnuity.Range(Chr(GetValue("客戶案件案號") + 65) & ":" & Chr(GetValue("客戶案件案號") + 65)).Delete Shift:=xlToLeft
'        End If
'        wksAnnuity.Range(Chr(intField) & "2:" & Chr(UBound(strFieldN) + intField) & intCounter).Font.Size = 12
'        '工作表更名
'        wksAnnuity.Name = IIf(Text2 = "2", "智權", "台一")
'        '繳回清單
'        If Text1 = "" And m_strSalesList <> "" Then
'            intCounter = 1
'            Call PrintSalesListExcel
'        End If
'    End If
'
'    'xlsAnnuity.Visible = True
'    'Mark by Amy 2015/04/16 關掉worksheet(非整個Excel) 會出現Excel Error
''    xlsAnnuity.WindowState = wdWindowStateMaximize
''    Set xlsAnnuity = Nothing
''    Set wksAnnuity = Nothing
'
'    'Add by Amy 2015/05/14
'    'Modify by Amy 2016/06/23 +判斷版本
'    If Val(xlsAnnuity.Version) < 12 Then
'        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
'    Else
'        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
'    End If
'    'end 2016/06/23
'    xlsAnnuity.Workbooks.Close
'    xlsAnnuity.Quit
'    If bolHasData = True Then MsgBox "Excel檔已產生!" & vbCrLf & vbCrLf & strFileName, vbInformation, Me.Caption
'    adoaccrpt106.Close
'    Set xlsAnnuity = Nothing
'    Set wksAnnuity = Nothing
'    Exit Sub
'
'ErrHnd:
'    xlsAnnuity.Visible = True
''    xlsAnnuity.WindowState = wdWindowStateMaximize
''    Set xlsAnnuity = Nothing
''    Set wksAnnuity = Nothing
'    adoaccrpt106.Close
'    'Modify by Amy 2016/06/23 +判斷版本
'    If Val(xlsAnnuity.Version) < 12 Then
'        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
'    Else
'        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
'    End If
'    'end 2016/06/23
'    xlsAnnuity.Workbooks.Close
'    xlsAnnuity.Quit
'    Set xlsAnnuity = Nothing
'    Set wksAnnuity = Nothing
'    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'*************************************************
' 產生Excel資料(收回)
'
'*************************************************
'Add by Amy 2020/04/27 +L 公司,程式調整
Public Sub PrintExcel2()
Dim strStart As Integer, CountData_S As Integer
Dim bolIsFirst As Boolean
Dim strQuery As String, strA As String, strCmp As String
Dim intMaxTitle As Integer 'Add by Amy 2022/07/08

On Error GoTo ErrHnd
    
    'Modify by Amy 2020/07/08 公司別改回輸入(不用下拉)
    strCmp = Trim(Text2) ' Trim(CboComp)
'    If strCmp <> MsgText(60) Then
'        If InStr(strCmp, "　") > 0 Then
'            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
'        End If
'    End If
    If strCmp <> MsgText(601) Then
'        If strCmp <> "J" And strCmp <> "L" Then
'            strA = " And r10603 Not In ('J','L') "
'        Else
            strA = " And r10603='" & strCmp & "' "
'        End If
    End If
    
    'strA = "Select Distinct Decode(r10603,'J','J','L','L','1') as Cmp From accrpt106 Where r10601 = '" & strUserNum & "' " & strA
    strA = "Select Distinct r10603 as Cmp From accrpt106 Where r10601 = '" & strUserNum & "' " & strA
    'Add by Amy 2022/07/04 法律所再產生一份以介紹人為主的資料
    If Trim(Text2) = MsgText(601) Or Trim(Text2) = "L" Then
        If ChkLaw = True Then
            strA = strA & " Union Select 'LZ' as Cmp From Dual Order by Cmp "
        End If
    End If
    'end 2020/07/08
    If rsA.State = adStateOpen Then rsA.Close
    rsA.CursorLocation = adUseClient
    rsA.Open strA, adoTaie, adOpenDynamic, adLockBatchOptimistic
    If rsA.RecordCount = 0 Then
        rsA.Close
        Exit Sub
    End If
    
    ReDim strFieldN(14)
    ReDim intWidth(14)
    'Modify by Amy 2020/05/28 +申請案號
    'Modify by Amy 2022/07/08 +介紹人
    strFieldN = Array("收據日期", "公司", "收款單號", "介紹人", "收據號碼", _
                      "收據抬頭", "客戶案件案號", "申請案號", "本所案號", "案件性質", _
                      "申請國家", "案件名稱", "收款金額", "溢收金額", "服務費", _
                      "規費", "扣繳")
    intWidth = Array(13, 4.5, 10, 8, 10, 10, 10, 10, 10, 10, _
                                10, 10, 10, 10, 14, 8, 8)
    'end 2022/07/08
    'end 2020/05/28
    bolIsFirst = True: intXlsSheet = 1: strText = "": CountData_S = 0: intField = 65: bolHasData = True
    intDelF = 0 'Add by Amy 2020/06/30
    
    rsA.MoveFirst
    For i = 0 To rsA.RecordCount - 1
        If i = 0 Then
            strFileName = strExcelPath & "智權人員帳款明細-收回" & Format(Now, "yyyymmddhhmmss") & MsgText(43)
            If Dir(strFileName) = MsgText(601) Then
               If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
                  MkDir strExcelPath
               End If
            Else
               Kill strFileName
            End If
            xlsAnnuity.SheetsInNewWorkbook = 3
            xlsAnnuity.Workbooks.add
            xlsAnnuity.Application.WindowState = xlMinimized
        End If
        If strWkName = MsgText(601) Then strWkName = Left(xlsAnnuity.Worksheets(1).Name, Len(xlsAnnuity.Worksheets(1).Name) - 1)
        'Add by Amy 2020/05/28 畫面公司別為空白=全部公司,若全公司都有資料需有6個Sheet,故需加工作表
        'Modify by Amy 2021/05/13 2010需增加工作表會錯,目前財務只有2010及2013版本 拿掉 And intXlsSheet <> 1 And Val(xlsAnnuity.Version) = 15
        If intXlsSheet > 3 Then
            'Modify by Amy 2022/07/04 +After:=wksAnnuity 加於最後
            xlsAnnuity.Worksheets.add After:=wksAnnuity  '插入sheet
        End If
        Set wksAnnuity = xlsAnnuity.Worksheets(strWkName & intXlsSheet)
        wksAnnuity.Activate
        '*** 抓取各公司別資料 ***
        'Modify by Amy 2020/07/28 每一個公司一個sheet顯示
'        If "" & rsA.Fields("Cmp") = "1" Then
'            strQuery = " And r10603<>'J' And r10603<>'L' "
'        Else
            strQuery = " And r10603='" & rsA.Fields("Cmp") & "' "
'        End If
        'end 2020/07/28
        'Modify by Amy 2022/07/08 法律所再產生一份以介紹人為主的資料
        If "" & rsA.Fields("Cmp") = "LZ" Then
            '介紹人以目前區為主-瑞婷
            strQuery = "Select r10603,r10606,r10607,r10608,r10609,r10610,r10611,r10612,r10613,r10614," & _
                                "r10615,r10616,r10617,r10624,r10625,r10626,r10630," & _
                                "r10631 as r10605,st15 as r10604,'' as st02,'' as r10631 From accrpt106,Staff " & _
                                "Where r10601 = '" & strUserNum & "' And r10631=st01(+) And r10631 Is Not Null "
            '勾選「客戶案件案號」
            If Check4.Value = 1 Then
                strQuery = strQuery & " Order by st15 asc,r10631,r10624 asc,r10602 asc"
            Else
                strQuery = strQuery & " Order by st15 asc,r10631,r10602 asc"
            End If
        Else
            'Modify by Amy 2022/07/08 +介紹人
            strQuery = "Select accrpt106.*,st02 From accrpt106,Staff Where r10601 = '" & strUserNum & "' And r10631=st01(+) " & strQuery
            '勾選「客戶案件案號」
            If Check4.Value = 1 Then
                strQuery = strQuery & " Order by r10624 asc,r10601 asc,r10602 asc"
            Else
                strQuery = strQuery & " Order by r10601 asc,r10602 asc"
            End If
        End If
        'end 2022/07/08
        If adoaccrpt106.State = adStateOpen Then adoaccrpt106.Close
        adoaccrpt106.CursorLocation = adUseClient
        adoaccrpt106.Open strQuery, adoTaie, adOpenDynamic, adLockBatchOptimistic
        With wksAnnuity
            If adoaccrpt106.RecordCount > 0 Then
                strSalesNo = "": strSameName = "":  m_strSalesList = "": strOldArea = "": CountData_S = 0
                intCounter = 1
                adoaccrpt106.MoveFirst
                Do While adoaccrpt106.EOF = False
                    If strSalesNo <> "" & adoaccrpt106.Fields("r10605").Value Then
                        m_strSalesList = m_strSalesList & "" & adoaccrpt106.Fields("r10605").Value & vbCrLf
                        '顯示員編
                        If strSalesNo = "" Then
                            Call SetExcel(1, bolIsFirst, A0902Query("" & adoaccrpt106.Fields("r10604").Value) & StaffQuery("" & adoaccrpt106.Fields("r10605").Value) & " " & adoaccrpt106.Fields("r10605").Value, "" & rsA.Fields("Cmp"))
                            strStart = intCounter '合計公式
                        End If
                    End If
                    '智權人員/業務區不同時換頁(ex:86052 104年及105年不同區的資料)
                    'Modify by Amy 2022/07/08 改變數 intCounter > 6/CountData_S = 27->改顯示資料25列
                    intMaxTitle = intTitleR_Fix + intTitleR + 1
                    If (intCounter > intMaxTitle + 1 And strSalesNo <> "" & adoaccrpt106.Fields("r10605").Value) Or (intCounter > intMaxTitle + 1 And strOldArea <> "" & adoaccrpt106.Fields("r10604").Value) _
                      Or CountData_S = 32 - (intMaxTitle + 1) Then
                        If intCounter > intMaxTitle + 1 And (strSalesNo <> "" & adoaccrpt106.Fields("r10605").Value Or strOldArea <> "" & adoaccrpt106.Fields("r10604").Value) Then
                    'end 2022/07/08
                            Call PrintSum2(strStart)
                            intCounter = intCounter + 1
                            strStart = intCounter
                        End If
                        '換頁
                        .Range("A" & intCounter).Select
                        .HPageBreaks.add Before:=.Application.ActiveCell
                        Call SetExcel(1, , A0902Query("" & adoaccrpt106.Fields("r10604").Value) & StaffQuery("" & adoaccrpt106.Fields("r10605").Value) & " " & adoaccrpt106.Fields("r10605").Value, "" & rsA.Fields("Cmp"))
                        CountData_S = 0
                    End If
                    
                    If IsNull(adoaccrpt106.Fields("r10608").Value) = False Then
                       .Range(Chr(GetValue("收款日期") + 65) & intCounter).Value = CFDate(adoaccrpt106.Fields("r10608").Value)
                    End If
                    If IsNull(adoaccrpt106.Fields("r10603").Value) = False Then
                        .Range(Chr(GetValue("公司") + 65) & intCounter).Value = "" & adoaccrpt106.Fields("r10603").Value
                        .Range(Chr(GetValue("公司") + 65) & intCounter).HorizontalAlignment = xlCenter
                    End If
                    If IsNull(adoaccrpt106.Fields("r10609").Value) = False Then
                        .Range(Chr(GetValue("收款單號") + 65) & intCounter).Value = "" & adoaccrpt106.Fields("r10609").Value
                    End If
                    'Add by Amy 2022/07/08 公司別為L顯示,介紹人
                    If "" & rsA.Fields("Cmp") = "L" Then
                        If IsNull(adoaccrpt106.Fields("r10631").Value) = False Then
                            .Range(Chr(GetValue("介紹人") + 65) & intCounter).Value = "" & adoaccrpt106.Fields("st02").Value
                        End If
                    End If
                    'end 2022/07/08
                    If IsNull(adoaccrpt106.Fields("r10610").Value) = False Then
                        .Range(Chr(GetValue("收據號碼") + 65) & intCounter).Value = "" & adoaccrpt106.Fields("r10610").Value
                    End If
                    strSameName = ("" & adoaccrpt106.Fields("r10607").Value)
                    If IsNull(adoaccrpt106.Fields("r10607").Value) = False Then
                       .Range(Chr(GetValue("收據抬頭") + 65) & intCounter).Value = Left("" & adoaccrpt106.Fields("r10607").Value, 6)
                    End If
                    If IsNull(adoaccrpt106.Fields("r10624").Value) = False Then
                       .Range(Chr(GetValue("客戶案件案號") + 65) & intCounter).Value = adoaccrpt106.Fields("r10624").Value
                    End If
                    'Add by Amy 2020/05/28 +申請案號
                    If IsNull(adoaccrpt106.Fields("r10630").Value) = False Then
                       .Range(Chr(GetValue("申請案號") + 65) & intCounter).NumberFormatLocal = "@"
                       .Range(Chr(GetValue("申請案號") + 65) & intCounter).Value = adoaccrpt106.Fields("r10630").Value
                    End If
                    If IsNull(adoaccrpt106.Fields("r10611").Value) = False Then
                       .Range(Chr(GetValue("本所案號") + 65) & intCounter).Value = adoaccrpt106.Fields("r10611").Value
                    End If
                    If IsNull(adoaccrpt106.Fields("r10613").Value) = False Then
                       .Range(Chr(GetValue("申請國家") + 65) & intCounter).Value = Left(adoaccrpt106.Fields("r10613").Value, 4)
                    End If
                    .Range(Chr(GetValue("案件名稱") + 65) & intCounter).Value = StrToStr(GetCaseName("" & adoaccrpt106.Fields("r10611").Value), 20)
    
                    If IsNull(adoaccrpt106.Fields("r10612").Value) = False Then
                       .Range(Chr(GetValue("案件性質") + 65) & intCounter).Value = adoaccrpt106.Fields("r10612").Value 'Modify by Amy 2024/12/24 原只取4個字
                    End If
    
                    'Excel資料的數字為0時以 0顯示之(DDollar => DDollar2)
                    If IsNull(adoaccrpt106.Fields("r10614").Value) = False Then
                       .Range(Chr(GetValue("收款金額") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10614").Value, DDollar2)
                    Else
                       .Range(Chr(GetValue("收款金額") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
                    End If
                    If IsNull(adoaccrpt106.Fields("r10615").Value) = False Then
                       .Range(Chr(GetValue("溢收金額") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10615").Value, DDollar2)
                    Else
                       .Range(Chr(GetValue("溢收金額") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
                    End If
                    If IsNull(adoaccrpt106.Fields("r10625").Value) = False Then
                       .Range(Chr(GetValue("服務費") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10625").Value, DDollar2)
                    Else
                       .Range(Chr(GetValue("服務費") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
                    End If
                    If IsNull(adoaccrpt106.Fields("r10617").Value) = False Then
                       .Range(Chr(GetValue("規費") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10617").Value, DDollar2)
                    Else
                       .Range(Chr(GetValue("規費") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
                    End If
                    If IsNull(adoaccrpt106.Fields("r10626").Value) = False Then
                       .Range(Chr(GetValue("扣繳") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10626").Value, DDollar2)
                    Else
                       .Range(Chr(GetValue("扣繳") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
                    End If
                    intCounter = intCounter + 1
                    CountData_S = CountData_S + 1 '計算智權人員明細筆數 for 換頁
                    strSalesNo = adoaccrpt106.Fields("r10605").Value
                    strOldArea = "" & adoaccrpt106.Fields("r10604").Value '業務區
                 
                    adoaccrpt106.MoveNext
                Loop
                '合計公式
                Call PrintSum2(strStart)
            End If
            'Add by Amy 2022/07/08 公司別非L公司,刪除「介紹人」欄
            If "" & rsA.Fields("Cmp") <> "L" Then
                .Range(Chr(GetValue("介紹人") + 65 - intDelF) & ":" & Chr(GetValue("介紹人") + 65 - intDelF)).Delete Shift:=xlToLeft
                intDelF = intDelF + 1
            End If
            'Modify by Amy 2020/06/30 加- intDelF 因欄位全顯示,若未勾「客戶案件案號」及「申請案號」欄,刪申請案號欄時會刪錯
            If Check4.Value = 0 Then
                wksAnnuity.Range(Chr(GetValue("客戶案件案號") + 65 - intDelF) & ":" & Chr(GetValue("客戶案件案號") + 65 - intDelF)).Delete Shift:=xlToLeft
                intDelF = intDelF + 1
            End If
            'Add by Amy 2020/05/28 +申請案號
            If Check5.Value = 0 Then
                wksAnnuity.Range(Chr(GetValue("申請案號") + 65 - intDelF) & ":" & Chr(GetValue("申請案號") + 65 - intDelF)).Delete Shift:=xlToLeft
                intDelF = intDelF + 1
            End If
            'end 2020/06/30
        End With
        'Modify by Amy 2022/07/08 固定抬頭改抓變數 /+ rsA.Fields("Cmp") = "LZ" 判斷/SetExcel(3)
        Call SetExcel(3)
        wksAnnuity.Range(Chr(intField) & intTitleR_Fix + 1 & ":" & Chr(UBound(strFieldN) + intField) & intCounter).Font.Size = 12
        If "" & rsA.Fields("Cmp") = "LZ" Then
            wksAnnuity.Name = CompNameQuery("L", 4) & "-介紹人"
        Else
            wksAnnuity.Name = CompNameQuery("" & rsA.Fields("Cmp"), 4) 'Modify by Amy 2020/07/08 要可輸2公司 原:A0802Query("" & rsA.Fields("Cmp"), True)
        End If
        'end 2022/07/08
        intXlsSheet = intXlsSheet + 1
        intDelF = 0 'Add by Amy 2020/07/28 每印完一個公司欄位刪除欄位要設0
        bolIsFirst = False
        '繳回清單
        If Text1 = "" And m_strSalesList <> "" Then
            intCounter = 1
            Call PrintSalesListExcel("" & rsA.Fields("Cmp"))
        End If
        
        rsA.MoveNext
    Next i
    '判斷版本
    If Val(xlsAnnuity.Version) < 12 Then
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
    Else
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
    End If

    xlsAnnuity.Workbooks.Close
    xlsAnnuity.Quit
    'Modify by Amy 2021/06/22 路徑改中文字顯示
    If bolHasData = True Then MsgBox "Excel檔已產生!" & vbCrLf & vbCrLf & strExcelPathN & Replace(strFileName, strExcelPath, ""), vbInformation, Me.Caption
    adoaccrpt106.Close
    Set xlsAnnuity = Nothing
    Set wksAnnuity = Nothing
    Exit Sub

ErrHnd:
    '判斷版本
    If Val(xlsAnnuity.Version) < 12 Then
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
    Else
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
    End If
    xlsAnnuity.Workbooks.Close
    xlsAnnuity.Quit
    adoaccrpt106.Close
    Set xlsAnnuity = Nothing
    Set wksAnnuity = Nothing
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Public Sub PrintExcel2_Old2()
''Modify by Amy 2016/09/19 合計改公式
''Dim strTot14 As String, strTot15 As String, strTot16 As String, strTot17 As String
'Dim strStart As Integer
''end 2016/09/19
'Dim strQuery As String, bolIsFirst As Boolean, CountData_S As Integer
'
'On Error GoTo ErrHnd
'
'    bolIsFirst = True: intXlsSheet = 1: strText = "": CountData_S = 0
'
'    If Text2 = "" Then
'        '未輸公司別 台一及智權分開顯示,先產生台一的資料
'        Text2.Tag = "1"
'    Else
'        Text2.Tag = IIf(Text2 = "2", "J", "1")
'    End If
'
'    strQuery = "Select * From accrpt106 Where r10601 = '" & strUserNum & "' And r10603" & IIf(Text2.Tag = "J", "='J'", "<>'J'")
'    If Check4.Value = 1 Then
'        strQuery = strQuery & " Order by r10624 asc,r10601 asc,r10602 asc"
'    Else
'        strQuery = strQuery & " Order by r10601 asc,r10602 asc"
'    End If
'
'    adoaccrpt106.CursorLocation = adUseClient
'    adoaccrpt106.Open strQuery, adoTaie, adOpenDynamic, adLockBatchOptimistic
'
'    If adoaccrpt106.RecordCount = 0 Then
'        adoaccrpt106.Close
'        MsgBox "無資料產生！" 'Add by Amy 2017/06/08
'        Exit Sub
'    End If
'
'    'Modify by Amy 2016/09/19 +公司/扣繳
'    'Modify by Amy 2017/10/03 +案件名稱
'    ReDim strFieldN(14)
'    ReDim intWidth(14)
'    strFieldN = Array("收據日期", "公司", "收款單號", "收據號碼", "收據抬頭", "客戶案件案號", _
'                      "本所案號", "案件性質", "申請國家", "案件名稱", "收款金額", "溢收金額", _
'                      "服務費", "規費", "扣繳")
'    intWidth = Array(13, 4.5, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, _
'                                14, 8, 8)
'    'end 2016/09/19
'    intField = 65: bolHasData = True
'
'    strFileName = strExcelPath & "智權人員帳款明細-收回" & Format(Now, "yyyymmddhhmmss") & MsgText(43)
'   If Dir(strFileName) = MsgText(601) Then
'      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
'         MkDir strExcelPath
'      End If
'   Else
'      Kill strFileName
'   End If
'
'NextComp:
'
'    strSalesNo = "": strSameName = "":  m_strSalesList = ""
'    strOldArea = "": CountData_S = 0 'Add by Amy 2016/09/19
'    intCounter = 1
'    'strTot14 = "": strTot15 = "": strTot16 = "": strTot17 = "" 'Modify by Amy 2019/09/19 合計改公式
'
'    If bolIsFirst = True Then
'        xlsAnnuity.SheetsInNewWorkbook = 3 'Added by Lydia 2019/04/08 預設工作表數量
'        xlsAnnuity.Workbooks.add
'        xlsAnnuity.Application.WindowState = xlMinimized
'    End If
'
'    'Modify by Amy 2017/09/25 for 工作表名稱改為中文
'    If strWkName = MsgText(601) Then strWkName = Left(xlsAnnuity.Worksheets(1).Name, Len(xlsAnnuity.Worksheets(1).Name) - 1)
'    Set wksAnnuity = xlsAnnuity.Worksheets(strWkName & intXlsSheet)
'    'end 2017/09/25
'    wksAnnuity.Activate
'
'    With wksAnnuity
'        If adoaccrpt106.RecordCount > 0 Then
'            '逐筆填值
'            Do While adoaccrpt106.EOF = False
'                If strSalesNo <> "" & adoaccrpt106.Fields("r10605").Value Then
'                    m_strSalesList = m_strSalesList & "" & adoaccrpt106.Fields("r10605").Value & vbCrLf
'                    If strSalesNo = "" Then
'                        Call SetExcel(1, bolIsFirst, A0902Query("" & adoaccrpt106.Fields("r10604").Value) & StaffQuery("" & adoaccrpt106.Fields("r10605").Value) & " " & adoaccrpt106.Fields("r10605").Value)
'                        strStart = intCounter 'Add by Amy 2016/09/19
'                    End If
'                End If
'                '智權人員不同時換頁
'                'Modfiy by Amy 2016/04/27 +業務區不同也換頁(ex:86052 104年及105年不同區的資料)
'                If (intCounter > 6 And strSalesNo <> "" & adoaccrpt106.Fields("r10605").Value) Or (intCounter > 6 And strOldArea <> "" & adoaccrpt106.Fields("r10604").Value) Or CountData_S = 27 Then
'                    If intCounter > 6 And (strSalesNo <> "" & adoaccrpt106.Fields("r10605").Value Or strOldArea <> "" & adoaccrpt106.Fields("r10604").Value) Then
'                        '合計
'                        'Modify by Amy 2016/09/19 合計改公式
'                        'Call PrintSum2(strTot14, strTot15, strTot16, strTot17)
'                        'strTot14 = "": strTot15 = "": strTot16 = "": strTot17 = ""
'                        Call PrintSum2(strStart)
'                        intCounter = intCounter + 1
'                        strStart = intCounter  'Add by Amy 2016/09/19
'                    End If
'
'                    '換頁
'                    .Range("A" & intCounter).Select
'                    .HPageBreaks.add Before:=.Application.ActiveCell
'                    Call SetExcel(1, , A0902Query("" & adoaccrpt106.Fields("r10604").Value) & StaffQuery("" & adoaccrpt106.Fields("r10605").Value) & " " & adoaccrpt106.Fields("r10605").Value)
'                    CountData_S = 0
'                End If
'
'               'Modify by Amy 2015/05/14 調整欄位顯示錯誤
'                If IsNull(adoaccrpt106.Fields("r10608").Value) = False Then
'                   .Range(Chr(GetValue("收款日期") + 65) & intCounter).Value = CFDate(adoaccrpt106.Fields("r10608").Value)
'                End If
'                'Add by Amy 2016/09/19 +公司
'                If IsNull(adoaccrpt106.Fields("r10603").Value) = False Then
'                    .Range(Chr(GetValue("公司") + 65) & intCounter).Value = "" & adoaccrpt106.Fields("r10603").Value
'                    .Range(Chr(GetValue("公司") + 65) & intCounter).HorizontalAlignment = xlCenter
'                End If
'                If IsNull(adoaccrpt106.Fields("r10609").Value) = False Then
'                    .Range(Chr(GetValue("收款單號") + 65) & intCounter).Value = "" & adoaccrpt106.Fields("r10609").Value
'                End If
'                If IsNull(adoaccrpt106.Fields("r10610").Value) = False Then
'                    .Range(Chr(GetValue("收據號碼") + 65) & intCounter).Value = "" & adoaccrpt106.Fields("r10610").Value
'                End If
'                strSameName = ("" & adoaccrpt106.Fields("r10607").Value)
'                If IsNull(adoaccrpt106.Fields("r10607").Value) = False Then
'                   .Range(Chr(GetValue("收據抬頭") + 65) & intCounter).Value = Left("" & adoaccrpt106.Fields("r10607").Value, 6)
'                End If
'                If IsNull(adoaccrpt106.Fields("r10624").Value) = False Then
'                   .Range(Chr(GetValue("客戶案件案號") + 65) & intCounter).Value = adoaccrpt106.Fields("r10624").Value
'                End If
'                If IsNull(adoaccrpt106.Fields("r10611").Value) = False Then
'                   .Range(Chr(GetValue("本所案號") + 65) & intCounter).Value = adoaccrpt106.Fields("r10611").Value
'                End If
'                If IsNull(adoaccrpt106.Fields("r10613").Value) = False Then
'                   .Range(Chr(GetValue("申請國家") + 65) & intCounter).Value = Left(adoaccrpt106.Fields("r10613").Value, 4)
'                End If
'                'Add by Amy 2017/10/03
'                .Range(Chr(GetValue("案件名稱") + 65) & intCounter).Value = StrToStr(GetCaseName("" & adoaccrpt106.Fields("r10611").Value), 20)
'
'                If IsNull(adoaccrpt106.Fields("r10612").Value) = False Then
'                   .Range(Chr(GetValue("案件性質") + 65) & intCounter).Value = Left(adoaccrpt106.Fields("r10612").Value, 4)
'                End If
'
'                'Excel資料的數字為0時以 0顯示之(DDollar => DDollar2)
'                If IsNull(adoaccrpt106.Fields("r10614").Value) = False Then
'                   .Range(Chr(GetValue("收款金額") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10614").Value, DDollar2)
'                   'strTot14 = Val(strTot14) + Val(adoaccrpt106.Fields("r10614").Value)
'                Else
'                   .Range(Chr(GetValue("收款金額") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
'                End If
'                If IsNull(adoaccrpt106.Fields("r10615").Value) = False Then
'                   .Range(Chr(GetValue("溢收金額") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10615").Value, DDollar2)
'                   'strTot15 = Val(strTot15) + Val(adoaccrpt106.Fields("r10615").Value)
'                Else
'                   .Range(Chr(GetValue("溢收金額") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
'                End If
'                If IsNull(adoaccrpt106.Fields("r10625").Value) = False Then
'                   .Range(Chr(GetValue("服務費") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10625").Value, DDollar2)
'                   'strTot16 = Val(strTot16) + Val(adoaccrpt106.Fields("r10625").Value)
'                Else
'                   .Range(Chr(GetValue("服務費") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
'                End If
'                If IsNull(adoaccrpt106.Fields("r10617").Value) = False Then
'                   .Range(Chr(GetValue("規費") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10617").Value, DDollar2)
'                   'strTot17 = Val(strTot17) + Val(adoaccrpt106.Fields("r10617").Value)
'                Else
'                   .Range(Chr(GetValue("規費") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
'                End If
'                'Add by Amy 2016/09/19 +扣繳
'                If IsNull(adoaccrpt106.Fields("r10626").Value) = False Then
'                   .Range(Chr(GetValue("扣繳") + 65) & intCounter).Value = Format(adoaccrpt106.Fields("r10626").Value, DDollar2)
'                Else
'                   .Range(Chr(GetValue("扣繳") + 65) & intCounter).Value = PUB_ChkExcelZero(1)
'                End If
'
'                intCounter = intCounter + 1
'                CountData_S = CountData_S + 1 '計算智權人員明細筆數 for 換頁
'                strSalesNo = adoaccrpt106.Fields("r10605").Value
'                strOldArea = "" & adoaccrpt106.Fields("r10604").Value 'Add by Amy 2016/04/27 +業務區
'                adoaccrpt106.MoveNext
'            Loop
'            '合計
'            Call PrintSum2(strStart) 'Modify by Amy 2019/09/19 合計改公式
'        Else
'            '沒資料只印表頭
'            Call SetExcel(1, True, "")
'        End If
'    End With
'
'    intXlsSheet = intXlsSheet + 1
'    '公司別為空先跑台一,再跑智權
'    If Trim(Text2) = "" Then
'        If bolIsFirst = True Then
'            If Check4.Value = 0 Then
'                wksAnnuity.Range(Chr(GetValue("客戶案件案號") + 65) & ":" & Chr(GetValue("客戶案件案號") + 65)).Delete Shift:=xlToLeft
'            End If
'            wksAnnuity.Range(Chr(intField) & "2:" & Chr(UBound(strFieldN) + intField) & intCounter).Font.Size = 12
'            wksAnnuity.Name = "台一"
'            bolIsFirst = False
'            '繳回清單
'            If Text1 = "" And m_strSalesList <> "" Then
'                intCounter = 1
'                Call PrintSalesListExcel
'            End If
'
'            '智權公司
'            strQuery = "Select * From accrpt106 Where r10601 = '" & strUserNum & "' And r10603='J' "
'            If Check4.Value = 1 Then
'                strQuery = strQuery & "Order by r10624 asc,r10601 asc,r10602 asc"
'            Else
'                strQuery = strQuery & "Order by r10601 asc,r10602 asc"
'            End If
'            If adoaccrpt106.State <> adStateClosed Then adoaccrpt106.Close
'            adoaccrpt106.CursorLocation = adUseClient
'            adoaccrpt106.Open strQuery, adoTaie, adOpenDynamic, adLockBatchOptimistic
'            Text2.Tag = "J"
'            GoTo NextComp
'        Else
'            '沒輸公司別 跑J公司清單
'            If Check4.Value = 0 Then
'                wksAnnuity.Range(Chr(GetValue("客戶案件案號") + 65) & ":" & Chr(GetValue("客戶案件案號") + 65)).Delete Shift:=xlToLeft
'            End If
'            wksAnnuity.Range(Chr(intField) & "2:" & Chr(UBound(strFieldN) + intField) & intCounter).Font.Size = 12
'            '工作表更名
'            wksAnnuity.Name = "智權"
'            '繳回清單
'            If Text1 = "" And m_strSalesList <> "" Then
'                intCounter = 1
'                PrintSalesListExcel
'            End If
'        End If
'    Else
'        '有輸公司別
'        If Check4.Value = 0 Then
'            wksAnnuity.Range(Chr(GetValue("客戶案件案號") + 65) & ":" & Chr(GetValue("客戶案件案號") + 65)).Delete Shift:=xlToLeft
'        End If
'        wksAnnuity.Range(Chr(intField) & "2:" & Chr(UBound(strFieldN) + intField) & intCounter).Font.Size = 12
'        '工作表更名
'        wksAnnuity.Name = IIf(Text2 = "2", "智權", "台一")
'        '繳回清單
'        If Text1 = "" And m_strSalesList <> "" Then
'            intCounter = 1
'            Call PrintSalesListExcel
'        End If
'    End If
'
'    'xlsAnnuity.Visible = True
'    'Add by Amy 2015/05/14
'    'Modify by Amy 2016/06/23 +判斷版本
'    If Val(xlsAnnuity.Version) < 12 Then
'        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
'    Else
'        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
'    End If
'    'end 2016/06/23
'    xlsAnnuity.Workbooks.Close
'    xlsAnnuity.Quit
'    If bolHasData = True Then MsgBox "Excel檔已產生!" & vbCrLf & vbCrLf & strFileName, vbInformation, Me.Caption
'    adoaccrpt106.Close
'    Set xlsAnnuity = Nothing
'    Set wksAnnuity = Nothing
'    Exit Sub
'
'ErrHnd:
'    'xlsAnnuity.Visible = True
'     'Modify by Amy 2016/06/23 +判斷版本
'    If Val(xlsAnnuity.Version) < 12 Then
'        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
'    Else
'        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
'    End If
'    'end 2016/06/23
'    xlsAnnuity.Workbooks.Close
'    xlsAnnuity.Quit
'    adoaccrpt106.Close
'    Set xlsAnnuity = Nothing
'    Set wksAnnuity = Nothing
'    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'Add By Sindy 2010/5/20
'*************************************************
' 產生Excel資料(未收)
'
'*************************************************
'Public Sub Old_PrintExcel1()
'Dim dblTot14 As Double, dblTot15 As Double, dblTot16 As Double, dblTot18 As Double, dblTot20 As Double
'Dim strTemp As String, dblSkipPageRow As Double
'
'On Error GoTo ErrHnd
'
'   strSalesNo = ""
'   strSameName = ""
'   intPage = 0
'   adoaccrpt106.CursorLocation = adUseClient
'   'Add By Sindy 2013/12/5
'   If Check4.Value = 1 Then
'      adoaccrpt106.Open "select * from accrpt106 where r10601 = '" & strUserNum & "' order by r10624 asc,r10601 asc,r10602 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   Else
'   '2013/12/5 END
'      adoaccrpt106.Open "select * from accrpt106 where r10601 = '" & strUserNum & "' order by r10601 asc,r10602 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   End If
'   'Modify By Sindy 2010/11/30
'   'If adoaccrpt106.RecordCount <= 2 Then
'   If adoaccrpt106.RecordCount = 0 Then
'   '2010/11/30 End
'      MsgBox MsgText(28), , MsgText(5)
'      adoaccrpt106.Close
'      Exit Sub
'   End If
'   m_strSalesList = ""
'   intCounter = 0: dblTot14 = 0: dblTot15 = 0: dblTot16 = 0: dblTot18 = 0: dblTot20 = 0: dblSkipPageRow = 0
'   Set xlsAnnuity = New Excel.Application
'   xlsAnnuity.Workbooks.Add
'   Set wksAnnuity = xlsAnnuity.Worksheets(1)
'   With wksAnnuity
'      .PageSetup.Orientation = xlLandscape '橫印
'      .PageSetup.LeftMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
'      .PageSetup.RightMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
'      .PageSetup.TopMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
'      .PageSetup.BottomMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
'      .PageSetup.HeaderMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
'      .PageSetup.FooterMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
'      '設定各欄位長度
'      .Columns("A:A").ColumnWidth = 13
'      .Columns("B:B").ColumnWidth = 10
'      'Add By Sindy 2013/6/14
'      .Columns("C:C").ColumnWidth = 10 '客戶代號
'      .Columns("D:D").ColumnWidth = 5  '印
'      '2013/6/14 END
'      .Columns("E:E").ColumnWidth = 10
'      'Add By Sindy 2013/12/6
'      .Columns("F:F").ColumnWidth = 10 '客戶案件案號
'      '2013/12/6 END
'      .Columns("G:G").ColumnWidth = 10 '本所案號
'      .Columns("H:H").ColumnWidth = 10
'      .Columns("I:I").ColumnWidth = 10
'      .Columns("J:J").ColumnWidth = 10
'      .Columns("K:K").ColumnWidth = 10
'      .Columns("L:L").ColumnWidth = 10
'      .Columns("M:M").ColumnWidth = 14
'      .Columns("N:N").ColumnWidth = 8
'      .Columns("O:O").ColumnWidth = 8
'      .Columns("P:P").ColumnWidth = 10
'      .Columns("Q:Q").ColumnWidth = 10
'      '逐筆填值
'      Do While adoaccrpt106.EOF = False
'         intCounter = intCounter + 1
'         If strSalesNo <> "" & adoaccrpt106.Fields("r10605").Value Then
'            m_strSalesList = m_strSalesList & "" & adoaccrpt106.Fields("r10605").Value & vbCrLf
'         End If
'         '若智權人員不同或資料已填滿一頁時
'         If strSalesNo <> (adoaccrpt106.Fields("r10605").Value) Or _
'            dblSkipPageRow = 26 Then
'            If intCounter <> 1 And _
'               strSalesNo <> (adoaccrpt106.Fields("r10605").Value) Then
'               '合計
'               .Range("H" & intCounter).Value = "合計"
'               '2012/5/2 modify by sonia
'               'strTemp = "A" & (intCounter - 1) & ":M" & (intCounter - 1)
'               strTemp = "A" & (intCounter - 1) & ":Q" & (intCounter - 1)
'               .Range(strTemp).Select
'               With .Application.Selection.Borders(xlEdgeBottom)
'                  .LineStyle = xlContinuous
'                  .Weight = xlThin
'                  .ColorIndex = xlAutomatic
'               End With
'               'Modified by Lydia 2014/12/18 excel資料的數字為0時以 0顯示之(DDollar => DDollar2)
'               .Range("J" & intCounter).Value = Format(dblTot14, DDollar2)
'               .Range("K" & intCounter).Value = Format(dblTot15, DDollar2)
'               .Range("L" & intCounter).Value = Format(dblTot16, DDollar2)
'               .Range("M" & intCounter).Value = Format(dblTot18, DDollar2)
'               .Range("O" & intCounter).Value = dblTot20
'               dblTot14 = 0: dblTot15 = 0: dblTot16 = 0: dblTot18 = 0: dblTot20 = 0
'               intCounter = intCounter + 1
'               '換頁
'               .Range("A" & intCounter).Select
'               .HPageBreaks.Add Before:=.Application.ActiveCell
'            End If
'            Call PrintExcelTitle1(A0902Query(adoaccrpt106.Fields("r10604").Value), adoaccrpt106.Fields("r10605").Value)
'            strSalesNo = adoaccrpt106.Fields("r10605").Value
'            dblSkipPageRow = 0
'         End If
'         '若收據抬頭不同時
'         'MODIFY BY SONIA 2014/4/30 瑞婷每一筆都要帶收據抬頭
'         'If strSameName <> (adoaccrpt106.Fields("r10607").Value) Then
'            If IsNull(adoaccrpt106.Fields("r10607").Value) = False Then
'               .Range("A" & intCounter).Value = Left("" & adoaccrpt106.Fields("r10607").Value, 6)
'            End If
'            strSameName = (adoaccrpt106.Fields("r10607").Value)
'         'End If
'         If IsNull(adoaccrpt106.Fields("r10608").Value) = False Then
'            .Range("B" & intCounter).Value = CFDate(adoaccrpt106.Fields("r10608").Value)
'         End If
'         'Add By Sindy 2013/6/14
'         If IsNull(adoaccrpt106.Fields("r10606").Value) = False Then
'            .Range("C" & intCounter).Value = adoaccrpt106.Fields("r10606").Value
'         End If
'         If IsNull(adoaccrpt106.Fields("r10623").Value) = False Then
'            .Range("D" & intCounter).Value = "N"
'         End If
'         '2013/6/14 END
'         If IsNull(adoaccrpt106.Fields("r10610").Value) = False Then
'            .Range("E" & intCounter).Value = adoaccrpt106.Fields("r10610").Value
'         End If
'         'Add By Sindy 2013/12/6 客戶案件案號
'         If IsNull(adoaccrpt106.Fields("r10624").Value) = False Then
'            .Range("F" & intCounter).Value = adoaccrpt106.Fields("r10624").Value
'         End If
'         '2013/12/6 END
'         '本所案號
'         If IsNull(adoaccrpt106.Fields("r10611").Value) = False Then
'            .Range("G" & intCounter).Value = adoaccrpt106.Fields("r10611").Value
'         End If
'         If IsNull(adoaccrpt106.Fields("r10612").Value) = False Then
'            .Range("H" & intCounter).Value = Left(adoaccrpt106.Fields("r10612").Value, 4)
'         End If
'         If IsNull(adoaccrpt106.Fields("r10613").Value) = False Then
'            .Range("I" & intCounter).Value = Left(adoaccrpt106.Fields("r10613").Value, 4)
'         End If
'         'Modified by Lydia 2014/12/18 excel資料的數字為0時以 0顯示之(DDollar => DDollar2)
'         If IsNull(adoaccrpt106.Fields("r10614").Value) = False Then
'            .Range("J" & intCounter).Value = Format(adoaccrpt106.Fields("r10614").Value, DDollar2)
'            dblTot14 = dblTot14 + Val(adoaccrpt106.Fields("r10614").Value)
'         Else '新增else 'Modified by Lydia
'            .Range("J" & intCounter).Value = PUB_ChkExcelZero(1)
'         End If
'         If IsNull(adoaccrpt106.Fields("r10615").Value) = False Then
'            .Range("K" & intCounter).Value = Format(adoaccrpt106.Fields("r10615").Value, DDollar2)
'            dblTot15 = dblTot15 + Val(adoaccrpt106.Fields("r10615").Value)
'         Else '新增else
'            .Range("K" & intCounter).Value = PUB_ChkExcelZero(1)
'         End If
'         If IsNull(adoaccrpt106.Fields("r10616").Value) = False Then
'            .Range("L" & intCounter).Value = Format(adoaccrpt106.Fields("r10616").Value, DDollar2)
'            dblTot16 = dblTot16 + Val(adoaccrpt106.Fields("r10616").Value)
'         Else '新增else
'            .Range("L" & intCounter).Value = PUB_ChkExcelZero(1)
'         End If
'         If IsNull(adoaccrpt106.Fields("r10618").Value) = False Then
'            .Range("M" & intCounter).Value = Format(adoaccrpt106.Fields("r10618").Value, DDollar2)
'            dblTot18 = dblTot18 + Val(adoaccrpt106.Fields("r10618").Value)
'         Else '新增else
'            .Range("M" & intCounter).Value = PUB_ChkExcelZero(1)
'         End If
'
''            .Range("J" & intCounter).Value = Format(PUB_ChkExcelZero(1, adoaccrpt106.Fields("r10614").Value), DDollar2)
''            dblTot14 = dblTot14 + Val(adoaccrpt106.Fields("r10614").Value)
''            .Range("K" & intCounter).Value = Format(PUB_ChkExcelZero(1, adoaccrpt106.Fields("r10615").Value), DDollar2)
''            dblTot15 = dblTot15 + Val(adoaccrpt106.Fields("r10615").Value)
''            .Range("L" & intCounter).Value = Format(PUB_ChkExcelZero(1, adoaccrpt106.Fields("r10616").Value), DDollar2)
''            dblTot16 = dblTot16 + Val(adoaccrpt106.Fields("r10616").Value)
''            .Range("M" & intCounter).Value = Format(PUB_ChkExcelZero(1, adoaccrpt106.Fields("r10618").Value), DDollar2)
''            dblTot18 = dblTot18 + Val(adoaccrpt106.Fields("r10618").Value)
'
'         If IsNull(adoaccrpt106.Fields("r10619").Value) = False Then
'            .Range("N" & intCounter).Value = adoaccrpt106.Fields("r10619").Value
'         End If
'         If IsNull(adoaccrpt106.Fields("r10620").Value) = False Then
'            .Range("O" & intCounter).Value = Val(adoaccrpt106.Fields("r10620").Value)
'            dblTot20 = dblTot20 + Val(adoaccrpt106.Fields("r10620").Value)
'         End If
'         If IsNull(adoaccrpt106.Fields("r10621").Value) = False Then
'            .Range("P" & intCounter).Value = CFDate(Val(adoaccrpt106.Fields("r10621").Value) - 19110000)
'         End If
'         '2012/5/2 add by sonia
'         If IsNull(adoaccrpt106.Fields("r10622").Value) = False Then
'            .Range("Q" & intCounter).Value = CFDate(Val(adoaccrpt106.Fields("r10622").Value) - 19110000)
'         End If
'         '2012/5/2 end
'         dblSkipPageRow = dblSkipPageRow + 1
'         adoaccrpt106.MoveNext
'      Loop
'      intCounter = intCounter + 1
'      '合計
'      .Range("H" & intCounter).Value = "合計"
'      '2012/5/2 modify by sonia
'      'strTemp = "A" & (intCounter - 1) & ":M" & (intCounter - 1)
'      strTemp = "A" & (intCounter - 1) & ":Q" & (intCounter - 1)
'      .Range(strTemp).Select
'      With .Application.Selection.Borders(xlEdgeBottom)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
'      'Modified by Lydia 2014/12/18 excel資料的數字為0時以 0顯示之(DDollar => DDollar2)
'      .Range("J" & intCounter).Value = Format(dblTot14, DDollar2)
'      .Range("K" & intCounter).Value = Format(dblTot15, DDollar2)
'      .Range("L" & intCounter).Value = Format(dblTot16, DDollar2)
'      .Range("M" & intCounter).Value = Format(dblTot18, DDollar2)
'      .Range("O" & intCounter).Value = dblTot20
'      intCounter = intCounter + 1
'      'Add By Sindy 2013/12/6
'      If Check4.Value = 0 Then
'         .Range("F:F").Delete Shift:=xlToLeft
'      End If
'      '2013/12/6 END
'   End With
'   If Text1 = "" And m_strSalesList <> "" Then
'      Call PrintSalesListExcel
'   End If
'   xlsAnnuity.Visible = True
'   xlsAnnuity.WindowState = wdWindowStateMaximize
'   Set xlsAnnuity = Nothing
'   Set wksAnnuity = Nothing
'   adoaccrpt106.Close
'   Exit Sub
'ErrHnd:
'   xlsAnnuity.Visible = True
'   xlsAnnuity.WindowState = wdWindowStateMaximize
'   Set xlsAnnuity = Nothing
'   Set wksAnnuity = Nothing
'   adoaccrpt106.Close
'   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
'End Sub

'Add By Sindy 2010/5/20
'Public Sub Old_PrintExcelTitle2(strDept As String, strSales)
'Dim i As Integer, strTemp As String
'Dim strText As String
'   intPage = intPage + 1
'   With wksAnnuity
'      For i = 1 To 2
'         If i = 1 Then
'            .Range("F" & intCounter).Value = "智權人員帳款明細表 (收回)"
'         ElseIf i = 2 Then
'            intCounter = intCounter + 1
'            strText = ""
'            If MaskEdBox1.Text <> MsgText(29) Then
'               If strText <> "" Then strText = strText & "　"
'               strText = strText & "帳款日期：" & MaskEdBox1.Text & "~" & MaskEdBox2.Text
'            End If
'            If Text4 <> "" Then
'               If strText <> "" Then strText = strText & "　"
'               strText = strText & "收據號碼：" & Text4.Text & "~" & Text5.Text
'            End If
'            If Text7 <> "" Then
'               If strText <> "" Then strText = strText & "　"
'               strText = strText & "客戶代號：" & Text7.Text & "~" & Text8.Text
'            End If
'            If MaskEdBox3.Text <> MsgText(29) Then
'               If strText <> "" Then strText = strText & "　"
'               strText = strText & "收文日：" & MaskEdBox3.Text & "~" & MaskEdBox4.Text
'            End If
'            If MaskEdBox5.Text <> MsgText(29) Then
'               If strText <> "" Then strText = strText & "　"
'               strText = strText & "發文日：" & MaskEdBox5.Text & "~" & MaskEdBox6.Text
'            End If
'            .Range("F" & intCounter).Value = strText
'         End If
'         strTemp = "A" & intCounter & ":L" & intCounter
'         .Range(strTemp).Select
'         With .Application.Selection
'             .HorizontalAlignment = xlCenter
'             '.MergeCells = True
'         End With
'         If i = 1 Then
'           With .Application.Selection
'             .Font.Size = 18
'            End With
'         End If
'      Next i
'      intCounter = intCounter + 1
'      '.Range("J" & intCounter).Value = "頁　次：" & intPage
'      .Range("A" & intCounter).Value = "頁　次：" & intPage
'      intCounter = intCounter + 1
'      '.Range("A" & intCounter).Value = "部門別：" & strDept
'      '.Range("J" & intCounter).Value = "智權人員：" & strSales
'      .Range("A" & intCounter).Value = "智權人員：" & strDept & " " & strSales
'      intCounter = intCounter + 1
'      .Range("A" & intCounter).Value = "收款日期"
'      .Range("B" & intCounter).Value = "收款單號"
'      .Range("C" & intCounter).Value = "收據號碼"
'      .Range("D" & intCounter).Value = "收據抬頭"
'      'Add By Sindy 2013/12/5
'      .Range("E" & intCounter).Value = "客戶案件案號"
'      '2013/12/5 END
'      .Range("F" & intCounter).Value = "本所案號"
'      .Range("G" & intCounter).Value = "案件性質"
'      .Range("H" & intCounter).Value = "申請國家"
'      .Range("I" & intCounter).Value = "收款金額"
'      .Range("J" & intCounter).Value = "溢收金額"
'      .Range("K" & intCounter).Value = "服務費"
'      .Range("L" & intCounter).Value = "規費"
'      strTemp = "A" & intCounter & ":L" & intCounter
'      .Range(strTemp).Select
'      With .Application.Selection.Borders(xlEdgeBottom)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'       End With
'       intCounter = intCounter + 1
'   End With
'End Sub
'
''Add By Sindy 2010/5/20
''*************************************************
'' 產生Excel資料(收回)
''
''*************************************************
Public Sub PrintExcel2_Old()
'Dim dblTot14 As Double, dblTot15 As Double, dblTot16 As Double, dblTot17 As Double
'Dim strTemp As String, dblSkipPageRow As Double
'
'On Error GoTo ErrHnd
'
'   strSalesNo = ""
'   strSameName = ""
'   intPage = 0
'   adoaccrpt106.CursorLocation = adUseClient
'   'Add By Sindy 2013/12/5
'   If Check4.Value = 1 Then
'      adoaccrpt106.Open "select * from accrpt106 where r10601 = '" & strUserNum & "' order by r10624 asc,r10601 asc,r10602 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   Else
'   '2013/12/5 END
'      adoaccrpt106.Open "select * from accrpt106 where r10601 = '" & strUserNum & "' order by r10601 asc,r10602 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   End If
'   'Modify By Sindy 2010/11/30
'   'If adoaccrpt106.RecordCount <= 2 Then
'   If adoaccrpt106.RecordCount = 0 Then
'   '2010/11/30 End
'      MsgBox MsgText(28), , MsgText(5)
'      adoaccrpt106.Close
'      Exit Sub
'   End If
'   m_strSalesList = ""
'   intCounter = 0: dblTot14 = 0: dblTot15 = 0: dblTot16 = 0: dblTot17 = 0: dblSkipPageRow = 0
'   Set xlsAnnuity = New Excel.Application
'   xlsAnnuity.Workbooks.Add
'   Set wksAnnuity = xlsAnnuity.Worksheets(1)
'   With wksAnnuity
'      .PageSetup.Orientation = xlLandscape '橫印
'      .PageSetup.LeftMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
'      .PageSetup.RightMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
'      .PageSetup.TopMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
'      .PageSetup.BottomMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
'      .PageSetup.HeaderMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
'      .PageSetup.FooterMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
'      '設定各欄位長度
'      .Columns("A:A").ColumnWidth = 13
'      .Columns("B:B").ColumnWidth = 10
'      .Columns("C:C").ColumnWidth = 10
'      .Columns("D:D").ColumnWidth = 10
'      'Add By Sindy 2013/12/5
'      .Columns("E:E").ColumnWidth = 10 '客戶案件案號
'      '2013/12/5 END
'      .Columns("F:F").ColumnWidth = 10 '本所案號
'      .Columns("G:G").ColumnWidth = 10
'      .Columns("H:H").ColumnWidth = 10
'      .Columns("I:I").ColumnWidth = 10
'      .Columns("J:J").ColumnWidth = 10
'      .Columns("K:K").ColumnWidth = 14
'      .Columns("L:L").ColumnWidth = 8
'      '逐筆填值
'      Do While adoaccrpt106.EOF = False
'         intCounter = intCounter + 1
'         If strSalesNo <> "" & adoaccrpt106.Fields("r10605").Value Then
'            m_strSalesList = m_strSalesList & "" & adoaccrpt106.Fields("r10605").Value & vbCrLf
'         End If
'         '若智權人員不同或資料已填滿一頁時
'         If strSalesNo <> (adoaccrpt106.Fields("r10605").Value) Or _
'            dblSkipPageRow = 26 Then
'            If intCounter <> 1 And _
'               strSalesNo <> (adoaccrpt106.Fields("r10605").Value) Then
'               '合計
'               .Range("F" & intCounter).Value = "合計"
'               strTemp = "A" & (intCounter - 1) & ":L" & (intCounter - 1)
'               .Range(strTemp).Select
'               With .Application.Selection.Borders(xlEdgeBottom)
'                  .LineStyle = xlContinuous
'                  .Weight = xlThin
'                  .ColorIndex = xlAutomatic
'               End With
'               'Modified by Lydia 2014/12/18 excel資料的數字為0時以 0顯示之(DDollar => DDollar2)
'               .Range("I" & intCounter).Value = Format(dblTot14, DDollar2)
'               .Range("J" & intCounter).Value = Format(dblTot15, DDollar2)
'               .Range("K" & intCounter).Value = Format(dblTot16, DDollar2)
'               .Range("L" & intCounter).Value = Format(dblTot17, DDollar2)
'               dblTot14 = 0: dblTot15 = 0: dblTot16 = 0: dblTot17 = 0
'               intCounter = intCounter + 1
'               '換頁
'               .Range("A" & intCounter).Select
'               .HPageBreaks.Add Before:=.Application.ActiveCell
'            End If
'            Call PrintExcelTitle2(A0902Query(adoaccrpt106.Fields("r10604").Value), adoaccrpt106.Fields("r10605").Value)
'            strSalesNo = adoaccrpt106.Fields("r10605").Value
'            dblSkipPageRow = 0
'         End If
'         If IsNull(adoaccrpt106.Fields("r10608").Value) = False Then
'            .Range("A" & intCounter).Value = CFDate(adoaccrpt106.Fields("r10608").Value)
'         End If
'         If IsNull(adoaccrpt106.Fields("r10609").Value) = False Then
'            .Range("B" & intCounter).Value = adoaccrpt106.Fields("r10609").Value
'         End If
'         If IsNull(adoaccrpt106.Fields("r10610").Value) = False Then
'            .Range("C" & intCounter).Value = adoaccrpt106.Fields("r10610").Value
'         End If
'         If IsNull(adoaccrpt106.Fields("r10607").Value) = False Then
'            .Range("D" & intCounter).Value = Left("" & adoaccrpt106.Fields("r10607").Value, 6)
'         End If
'         'Add By Sindy 2013/12/5 客戶案件案號
'         If IsNull(adoaccrpt106.Fields("r10624").Value) = False Then
'            .Range("E" & intCounter).Value = adoaccrpt106.Fields("r10624").Value
'         End If
'         '2013/12/5 END
'         '本所案號
'         If IsNull(adoaccrpt106.Fields("r10611").Value) = False Then
'            .Range("F" & intCounter).Value = adoaccrpt106.Fields("r10611").Value
'         End If
'         If IsNull(adoaccrpt106.Fields("r10612").Value) = False Then
'            .Range("G" & intCounter).Value = Left(adoaccrpt106.Fields("r10612").Value, 4)
'         End If
'         If IsNull(adoaccrpt106.Fields("r10613").Value) = False Then
'            .Range("H" & intCounter).Value = Left(adoaccrpt106.Fields("r10613").Value, 4)
'         End If
'         'Modified by Lydia 2014/12/18 excel資料的數字為0時以 0顯示之(DDollar => DDollar2)
'         If IsNull(adoaccrpt106.Fields("r10614").Value) = False Then
'            .Range("I" & intCounter).Value = Format(adoaccrpt106.Fields("r10614").Value, DDollar2)
'            dblTot14 = dblTot14 + Val(adoaccrpt106.Fields("r10614").Value)
'         Else '新增else
'            .Range("I" & intCounter).Value = PUB_ChkExcelZero(1)
'         End If
'         If IsNull(adoaccrpt106.Fields("r10615").Value) = False Then
'            .Range("J" & intCounter).Value = Format(adoaccrpt106.Fields("r10615").Value, DDollar2)
'            dblTot15 = dblTot15 + Val(adoaccrpt106.Fields("r10615").Value)
'         Else '新增else
'            .Range("J" & intCounter).Value = PUB_ChkExcelZero(1)
'         End If
'         If IsNull(adoaccrpt106.Fields("r10616").Value) = False Then
'            .Range("K" & intCounter).Value = Format(adoaccrpt106.Fields("r10616").Value, DDollar2)
'            dblTot16 = dblTot16 + Val(adoaccrpt106.Fields("r10616").Value)
'         Else '新增else
'            .Range("K" & intCounter).Value = PUB_ChkExcelZero(1)
'         End If
'         If IsNull(adoaccrpt106.Fields("r10617").Value) = False Then
'            .Range("L" & intCounter).Value = Format(adoaccrpt106.Fields("r10617").Value, DDollar2)
'            dblTot17 = dblTot17 + Val(adoaccrpt106.Fields("r10617").Value)
'         Else '新增else
'            .Range("L" & intCounter).Value = PUB_ChkExcelZero(1)
'         End If
''            .Range("I" & intCounter).Value = Format(PUB_ChkExcelZero(1, adoaccrpt106.Fields("r10614").Value), DDollar2)
''            dblTot14 = dblTot14 + Val(adoaccrpt106.Fields("r10614").Value)
''            .Range("J" & intCounter).Value = Format(PUB_ChkExcelZero(1, adoaccrpt106.Fields("r10615").Value), DDollar2)
''            dblTot15 = dblTot15 + Val(adoaccrpt106.Fields("r10615").Value)
''            .Range("K" & intCounter).Value = Format(PUB_ChkExcelZero(1, adoaccrpt106.Fields("r10616").Value), DDollar2)
''            dblTot16 = dblTot16 + Val(adoaccrpt106.Fields("r10616").Value)
''            .Range("L" & intCounter).Value = Format(PUB_ChkExcelZero(1, adoaccrpt106.Fields("r10617").Value), DDollar2)
''            dblTot17 = dblTot17 + Val(adoaccrpt106.Fields("r10617").Value)
'         dblSkipPageRow = dblSkipPageRow + 1
'         adoaccrpt106.MoveNext
'      Loop
'      intCounter = intCounter + 1
'      '合計
'      .Range("F" & intCounter).Value = "合計"
'      strTemp = "A" & (intCounter - 1) & ":L" & (intCounter - 1)
'      .Range(strTemp).Select
'      With .Application.Selection.Borders(xlEdgeBottom)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
'      'Modified by Lydia 2014/12/18 excel資料的數字為0時以 0顯示之(DDollar => DDollar2)
'      .Range("I" & intCounter).Value = Format(dblTot14, DDollar2)
'      .Range("J" & intCounter).Value = Format(dblTot15, DDollar2)
'      .Range("K" & intCounter).Value = Format(dblTot16, DDollar2)
'      .Range("L" & intCounter).Value = Format(dblTot17, DDollar2)
'      intCounter = intCounter + 1
'      'Add By Sindy 2013/12/5
'      If Check4.Value = 0 Then
'         .Range("E:E").Delete Shift:=xlToLeft
'      End If
'      '2013/12/5 END
'   End With
'   If Text1 = "" And m_strSalesList <> "" Then
'      Call PrintSalesListExcel
'   End If
'   xlsAnnuity.Visible = True
'   xlsAnnuity.WindowState = wdWindowStateMaximize
'   Set xlsAnnuity = Nothing
'   Set wksAnnuity = Nothing
'   adoaccrpt106.Close
'   Exit Sub
'ErrHnd:
'   xlsAnnuity.Visible = True
'   xlsAnnuity.WindowState = wdWindowStateMaximize
'   Set xlsAnnuity = Nothing
'   Set wksAnnuity = Nothing
'   adoaccrpt106.Close
'   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

''*************************************************
'' 列印資料(未收)
''
''*************************************************
'Public Sub PrintReport1()
'on error GoTo ErrHnd
'
'   strSalesNo = ""
'   strSameName = ""
'   lngCounter = 0
'   lngAmount = 0
'   intPage = 0
'   adoaccrpt106.CursorLocation = adUseClient
'   adoaccrpt106.Open "select * from accrpt106 where r10601 = '" & strUserNum & "' order by r10601 asc, r10602 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   If adoaccrpt106.RecordCount <= 2 Then
'      MsgBox MsgText(28), , MsgText(5)
'      adoaccrpt106.Close
'      Exit Sub
'   End If
'   m_strSalesList = ""
'   Do While adoaccrpt106.EOF = False
'      'Modify By Sindy 2010/5/20
'      'If Mid(strSalesNo, 2) <> "" & adoaccrpt106.Fields("r10605").Value Then
'      If strSalesNo <> "" & adoaccrpt106.Fields("r10605").Value Then
'         m_strSalesList = m_strSalesList & "" & adoaccrpt106.Fields("r10605").Value & vbCrLf
'      End If
'      'Modify By Sindy 2010/5/20
''      '若公司別+智權人員不同時
''      If strSalesNo <> (adoaccrpt106.Fields("r10603").Value & adoaccrpt106.Fields("r10605").Value) Then
'      '若智權人員不同時
'      If strSalesNo <> (adoaccrpt106.Fields("r10605").Value) Then
'         If intPage <> 0 Then
'            Printer.NewPage
'         End If
'         intCounter = 0
'         intPage = intPage + 1
'         PrintHead1
'         'Modify By Sindy 2010/5/20
'         'strSalesNo = adoaccrpt106.Fields("r10603").Value & adoaccrpt106.Fields("r10605").Value
'         strSalesNo = adoaccrpt106.Fields("r10605").Value
'      End If
'      If intCounter > 35 Then
'         intCounter = 0
'         intPage = intPage + 1
'         Printer.NewPage
'         PrintHead1
'      End If
'      '若客戶編號+收據抬頭不同時
'      If strSameName <> (adoaccrpt106.Fields("r10606").Value & adoaccrpt106.Fields("r10607").Value) Then
'         Printer.CurrentX = 500 - m_dblLeftDiff
'         Printer.CurrentY = 3800 + intCounter * 300
'         If adoaccrpt106.Fields("r10611").Value = ReportSum(24) Or adoaccrpt106.Fields("r10611").Value = ReportSum(25) Or adoaccrpt106.Fields("r10614").Value = ReportSum(4) Or adoaccrpt106.Fields("r10614").Value = ReportSum(8) Then
'         Else
'            If IsNull(adoaccrpt106.Fields("r10606").Value) = False Then
'                '若有收據編號時才印客戶編號
'                If "" & adoaccrpt106("R10610").Value <> "" Then
'                    Printer.Print adoaccrpt106.Fields("r10606").Value
'                End If
'            Else
'               Printer.Print ""
'            End If
'         End If
'         Printer.CurrentX = 1700 - m_dblLeftDiff
'         Printer.CurrentY = 3800 + intCounter * 300
'         If IsNull(adoaccrpt106.Fields("r10607").Value) = False Then
'            Printer.Print Left("" & adoaccrpt106.Fields("r10607").Value, 6)
'         Else
'            Printer.Print ""
'         End If
'         strSameName = (adoaccrpt106.Fields("r10606").Value & adoaccrpt106.Fields("r10607").Value)
'      End If
'      Printer.CurrentX = 3500 - m_dblLeftDiff - 400
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt106.Fields("r10608").Value) = False Then
'         Printer.Print CFDate(adoaccrpt106.Fields("r10608").Value)
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 4500 - m_dblLeftDiff - 400
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt106.Fields("r10610").Value) = False Then
'         Printer.Print adoaccrpt106.Fields("r10610").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 5700 - m_dblLeftDiff - 400
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt106.Fields("r10611").Value) = False Then
'         Printer.Print adoaccrpt106.Fields("r10611").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 7100 - m_dblLeftDiff - 400
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt106.Fields("r10612").Value) = False Then
'         Printer.Print Left(adoaccrpt106.Fields("r10612").Value, 4)
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 8100 - m_dblLeftDiff - 400
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt106.Fields("r10613").Value) = False Then
'         Printer.Print Left(adoaccrpt106.Fields("r10613").Value, 4)
'      Else
'         Printer.Print ""
'      End If
'      If IsNull(adoaccrpt106.Fields("r10614").Value) = False Then
'         strAmount = Format(adoaccrpt106.Fields("r10614").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 9100 + 1100 - intLength - m_dblLeftDiff - 400
'         Printer.CurrentY = 3800 + intCounter * 300
'         Printer.Print strAmount
'      Else
'         Printer.Print ""
'      End If
'      If IsNull(adoaccrpt106.Fields("r10615").Value) = False Then
'         strAmount = Format(adoaccrpt106.Fields("r10615").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 10200 + 1100 - intLength - m_dblLeftDiff - 400
'         Printer.CurrentY = 3800 + intCounter * 300
'         Printer.Print strAmount
'      Else
'         Printer.Print ""
'      End If
'      If IsNull(adoaccrpt106.Fields("r10616").Value) = False Then
'         strAmount = Format(adoaccrpt106.Fields("r10616").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 11300 + 1100 - intLength - m_dblLeftDiff - 400
'         Printer.CurrentY = 3800 + intCounter * 300
'         Printer.Print strAmount
'      Else
'         Printer.Print ""
'      End If
'      intCounter = intCounter + 1
'      adoaccrpt106.MoveNext
'   Loop
'   adoaccrpt106.Close
'   Printer.EndDoc
'ErrHnd:
'   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
'End Sub
'
''*************************************************
''  抬頭列印
''
''*************************************************
'Private Sub PrintHead2()
'   Printer.FontSize = 16
'   Printer.CurrentX = 4000 - m_dblLeftDiff
'   Printer.CurrentY = 1000
'   Printer.Print ReportTitle(106)
'   Printer.CurrentX = 8000 - m_dblLeftDiff
'   Printer.CurrentY = 1000
'   Printer.Print "(收回)"
'   Printer.FontSize = 10
'   Printer.CurrentX = 4500 - m_dblLeftDiff
'   Printer.CurrentY = 1800
'   Printer.Print "帳款日期: "
'   Printer.CurrentX = 5600 - m_dblLeftDiff
'   Printer.CurrentY = 1800
'   Printer.Print MaskEdBox1.Text & " ~ " & MaskEdBox2.Text
'   Printer.CurrentX = 500 - m_dblLeftDiff
'   Printer.CurrentY = 2100
'   Printer.Print "列印人員: "
'   Printer.CurrentX = 1800 - m_dblLeftDiff
'   Printer.CurrentY = 2100
'   Printer.Print StaffQuery(strUserNum)
'   Printer.CurrentX = 9700 - m_dblLeftDiff
'   Printer.CurrentY = 2100
'   Printer.Print "列印日期: "
'   Printer.CurrentX = 10700 - m_dblLeftDiff
'   Printer.CurrentY = 2100
'   Printer.Print CFDate(ACDate(ServerDate))
'   Printer.CurrentX = 9700 - m_dblLeftDiff
'   Printer.CurrentY = 2400
'   Printer.Print "頁　　次: "
'   Printer.CurrentX = 10700 - m_dblLeftDiff
'   Printer.CurrentY = 2400
'   Printer.Print intPage
'   Printer.CurrentX = 500 - m_dblLeftDiff
'   Printer.CurrentY = 2700
'   Printer.Print "公司別: "
'   Printer.CurrentX = 1300 - m_dblLeftDiff
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt106.Fields("r10603").Value) = False Then
'      Printer.Print adoaccrpt106.Fields("r10603").Value
'      Printer.CurrentX = 1500 - m_dblLeftDiff
'      Printer.CurrentY = 2700
'      Printer.Print A0802Query(adoaccrpt106.Fields("r10603").Value)
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 4500 - m_dblLeftDiff
'   Printer.CurrentY = 2700
'   Printer.Print "部門別: "
'   Printer.CurrentX = 5600 - m_dblLeftDiff
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt106.Fields("r10604").Value) = False Then
'      Printer.Print adoaccrpt106.Fields("r10604").Value
'      Printer.CurrentX = 6300 - m_dblLeftDiff
'      Printer.CurrentY = 2700
'      Printer.Print A0902Query(adoaccrpt106.Fields("r10604").Value)
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 9700 - m_dblLeftDiff
'   Printer.CurrentY = 2700
'   Printer.Print "智權人員: "
'   Printer.CurrentX = 10700 - m_dblLeftDiff
'   Printer.CurrentY = 2700
'   If IsNull(adoaccrpt106.Fields("r10605").Value) = False Then
'      Printer.Print adoaccrpt106.Fields("r10605").Value
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 500 - m_dblLeftDiff
'   Printer.CurrentY = 3300
'   Printer.Print "收款日期"
'   Printer.CurrentX = 1500 - m_dblLeftDiff
'   Printer.CurrentY = 3300
'   Printer.Print "收款單號"
'   Printer.CurrentX = 2700 - m_dblLeftDiff
'   Printer.CurrentY = 3300
'   Printer.Print "收據號碼"
'   Printer.CurrentX = 3900 - m_dblLeftDiff
'   Printer.CurrentY = 3300
'   Printer.Print "收據抬頭"
'   Printer.CurrentX = 5300 - m_dblLeftDiff
'   Printer.CurrentY = 3300
'   Printer.Print "本所案號"
'   Printer.CurrentX = 6700 - m_dblLeftDiff
'   Printer.CurrentY = 3300
'   Printer.Print "案件性質"
'   Printer.CurrentX = 7700 - m_dblLeftDiff
'   Printer.CurrentY = 3300
'   Printer.Print "申請國家"
'   Printer.CurrentX = 8700 - m_dblLeftDiff
'   Printer.CurrentY = 3300
'   Printer.Print "收款金額"
'   Printer.CurrentX = 9700 - m_dblLeftDiff
'   Printer.CurrentY = 3300
'   Printer.Print "溢收金額"
'   Printer.CurrentX = 10700 - m_dblLeftDiff
'   Printer.CurrentY = 3300
'   Printer.Print "服務費"
'   Printer.CurrentX = 11500 - m_dblLeftDiff
'   Printer.CurrentY = 3300
'   Printer.Print "規費"
'   Printer.Line (500 - m_dblLeftDiff, 3700)-(12300 - m_dblLeftDiff, 3700)
'End Sub
'
''*************************************************
'' 列印資料(收回)
''
''*************************************************
'Public Sub PrintReport2()
'   strSalesNo = ""
'   lngCounter = 0
'   lngAmount = 0
'   intPage = 0
'   adoaccrpt106.CursorLocation = adUseClient
'   adoaccrpt106.Open "select * from accrpt106 where r10601 = '" & strUserNum & "' order by r10601 asc, r10602 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   If adoaccrpt106.RecordCount <= 2 Then
'      MsgBox MsgText(28), , MsgText(5)
'      adoaccrpt106.Close
'      Exit Sub
'   End If
'   m_strSalesList = ""
'   Do While adoaccrpt106.EOF = False
'      'Modify By Sindy 2010/5/20
''      If Mid(strSalesNo, 2) <> "" & adoaccrpt106.Fields("r10605").Value Then
'      If strSalesNo <> "" & adoaccrpt106.Fields("r10605").Value Then
'         m_strSalesList = m_strSalesList & "" & adoaccrpt106.Fields("r10605").Value & vbCrLf
'      End If
'      'Modify By Sindy 2010/5/20
''      '若公司別+智權人員不同時
''      If strSalesNo <> (adoaccrpt106.Fields("r10603").Value & adoaccrpt106.Fields("r10605").Value) Then
'      '若智權人員不同時
'      If strSalesNo <> (adoaccrpt106.Fields("r10605").Value) Then
'         If intPage <> 0 Then
'            Printer.NewPage
'         End If
'         intCounter = 0
'         intPage = intPage + 1
'         PrintHead2
'         'Modify By Sindy 2010/5/20
''         strSalesNo = (adoaccrpt106.Fields("r10603").Value & adoaccrpt106.Fields("r10605").Value)
'         strSalesNo = (adoaccrpt106.Fields("r10605").Value)
'      End If
'      If intCounter > 35 Then
'         intCounter = 0
'         intPage = intPage + 1
'         Printer.NewPage
'         PrintHead2
'      End If
'      Printer.CurrentX = 500 - m_dblLeftDiff
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt106.Fields("r10608").Value) = False Then
'         Printer.Print CFDate(adoaccrpt106.Fields("r10608").Value)
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 1500 - m_dblLeftDiff
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt106.Fields("r10609").Value) = False Then
'         Printer.Print adoaccrpt106.Fields("r10609").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 2700 - m_dblLeftDiff
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt106.Fields("r10610").Value) = False Then
'         Printer.Print adoaccrpt106.Fields("r10610").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 3900 - m_dblLeftDiff
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt106.Fields("r10607").Value) = False Then
'         Printer.Print Left("" & adoaccrpt106.Fields("r10607").Value, 6)
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 5300 - m_dblLeftDiff
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt106.Fields("r10611").Value) = False Then
'         Printer.Print adoaccrpt106.Fields("r10611").Value
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 6700 - m_dblLeftDiff
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt106.Fields("r10612").Value) = False Then
'         Printer.Print Left(adoaccrpt106.Fields("r10612").Value, 4)
'      Else
'         Printer.Print ""
'      End If
'      Printer.CurrentX = 7700 - m_dblLeftDiff
'      Printer.CurrentY = 3800 + intCounter * 300
'      If IsNull(adoaccrpt106.Fields("r10613").Value) = False Then
'         Printer.Print Left(adoaccrpt106.Fields("r10613").Value, 4)
'      Else
'         Printer.Print ""
'      End If
'      If IsNull(adoaccrpt106.Fields("r10614").Value) = False Then
'         strAmount = Format(adoaccrpt106.Fields("r10614").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 8700 + 1000 - intLength - m_dblLeftDiff
'         Printer.CurrentY = 3800 + intCounter * 300
'         Printer.Print strAmount
'      Else
'         Printer.Print ""
'      End If
'      If IsNull(adoaccrpt106.Fields("r10615").Value) = False Then
'         strAmount = Format(adoaccrpt106.Fields("r10615").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 9700 + 1000 - intLength - m_dblLeftDiff
'         Printer.CurrentY = 3800 + intCounter * 300
'         Printer.Print strAmount
'      Else
'         Printer.Print ""
'      End If
'      If IsNull(adoaccrpt106.Fields("r10616").Value) = False Then
'         strAmount = Format(adoaccrpt106.Fields("r10616").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 10700 + 800 - intLength - m_dblLeftDiff
'         Printer.CurrentY = 3800 + intCounter * 300
'         Printer.Print strAmount
'      Else
'         Printer.Print ""
'      End If
'      If IsNull(adoaccrpt106.Fields("r10617").Value) = False Then
'         strAmount = Format(adoaccrpt106.Fields("r10617").Value, DDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 11500 + 800 - intLength - m_dblLeftDiff
'         Printer.CurrentY = 3800 + intCounter * 300
'         Printer.Print strAmount
'      Else
'         Printer.Print ""
'      End If
'      intCounter = intCounter + 1
'      adoaccrpt106.MoveNext
'   Loop
'   adoaccrpt106.Close
'   Printer.EndDoc
'End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck(ByRef stMsg As String) As Boolean
   'Add by Amy 2022/07/04 列印類別判斷改至此
   If Trim(Text3) = MsgText(601) Then
      stMsg = Label3 & "不可為空！"
      Text3.SetFocus
      Exit Function
   End If
   
   'Add by Amy 2022/07/20 避免抓案源資料,沒下日期資料會很多
   If (MaskEdBox1.Text = MsgText(29) Or MaskEdBox1.Text = MsgText(601)) _
    And (MaskEdBox2.Text = MsgText(29) Or MaskEdBox2.Text = MsgText(601)) Then
        stMsg = Label2 & "起或迄至少輸一個！"
        Exit Function
   End If
   'end 2022/07/20
   
   If MaskEdBox1.Text <> MsgText(29) Then
      If MaskEdBox1.Text <> MsgText(601) Then
         FormCheck = True
         Exit Function
      End If
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      If MaskEdBox2.Text <> MsgText(601) Then
         FormCheck = True
         Exit Function
      End If
   End If
   
   If Text1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   'Modify by Morgan 2007/10/1 智權人員範圍改成一個
   'If Text2 <> MsgText(601) Then
   '   FormCheck = True
   '   Exit Function
   'End If
   'end 2007/10/1
   
   'Add By Sindy 2016/6/13
   If Text4 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text5 <> MsgText(601) Then
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
   If MaskEdBox3.Text <> MsgText(29) Then
      If MaskEdBox3.Text <> MsgText(601) Then
         FormCheck = True
         Exit Function
      End If
   End If
   If MaskEdBox4.Text <> MsgText(29) Then
      If MaskEdBox4.Text <> MsgText(601) Then
         FormCheck = True
         Exit Function
      End If
   End If
   If MaskEdBox5.Text <> MsgText(29) Then
      If MaskEdBox5.Text <> MsgText(601) Then
         FormCheck = True
         Exit Function
      End If
   End If
   If MaskEdBox6.Text <> MsgText(29) Then
      If MaskEdBox6.Text <> MsgText(601) Then
         FormCheck = True
         Exit Function
      End If
   End If
   If Text9 <> MsgText(601) Then
      If Text10 <> MsgText(601) Then
         FormCheck = True
         Exit Function
      End If
   End If
   '2016/6/13 END
   FormCheck = False
End Function

Private Sub Text3_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2003/05/12
    KeyAscii = UpperCase(KeyAscii)
    'Modify by Amy 2020/04/27 +3.往來
    If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 Then
        KeyAscii = 0
    End If
End Sub

'Add by Amy 2016/09/19
Private Sub Text3_Validate(Cancel As Boolean)
    'Mark by Amy 2020/03/12 不使用 原:輸1(未收)可勾選服務費、規費欄分開
'    Check5.Enabled = False: Check5.Value = 0
'    If Text3 = "1" Then Check5.Enabled = True
    'Add by Amy 2020/04/27輸3(往來)自動勾選客戶案件案號,帳款日期改名稱
    Check4.Value = 0
    Label2.Caption = "帳款日期": Label2.ForeColor = &H80000012
    If Text3 = "2" Then
        Label2.Caption = "收款日期": Label2.ForeColor = &HFF&
    Else
        If Text3 = "3" Then Check4.Value = 1
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
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2010/9/13
'客戶代號(起)
Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Text7_Validate(Cancel As Boolean)
   If Len(Text7) = 6 Then
      Text7 = AfterZero(Text7)
   End If
   '帶出關係企業(尾3碼改999)
   'Modify By Sindy 2014/8/11 999=>ZZZ
   'If Len(Text7.Text) >= 6 Then Text8.Text = Left(Text7.Text, 6) & "999"
   If Len(Text7.Text) >= 6 Then Text8.Text = Left(Text7.Text, 6) & "ZZZ"
End Sub
'客戶代號(迄)
Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Text8_Validate(Cancel As Boolean)
   If Len(Text8) = 6 Then
      Text8 = AfterZero(Text8)
   End If
End Sub
'2010/9/13 End

'Add by Amy 2015/03/03
Private Function GetValue(pFieldN As String) As Integer
   Dim jj As Integer
 
    For jj = 1 To UBound(strFieldN)
       If UCase(strFieldN(jj)) = UCase(pFieldN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function

'intChoose:1.明細 / 2.繳回清單 / 3.明細抬頭
'Modify by Amy 2020/04/27 +公司別參數
Private Sub SetExcel(ByVal intChoose As Integer, Optional IsFirst As Boolean = False, Optional strSales As String, Optional ByVal stCmp As String)
    Dim ii As Integer, strTitleField As String
    Dim strChoose As String, strCol As String, strCol2 As String 'Add by Amy 2020/04/27
    Dim strTpField As String 'Add by Amy 2022/07/04
    
    If IsFirst = True And strText = MsgText(601) Then
        If MaskEdBox1.Text <> MsgText(29) Then
            If strText <> "" Then strText = strText & "　"
            strText = strText & "帳款日期：" & MaskEdBox1.Text & "~" & MaskEdBox2.Text
        End If
        If Text4 <> "" Then
            If strText <> "" Then strText = strText & "　"
            strText = strText & "收據號碼：" & Text4.Text & "~" & Text5.Text
        End If
        If Text7 <> "" Then
            If strText <> "" Then strText = strText & "　"
            strText = strText & "客戶代號：" & Text7.Text & "~" & Text8.Text
        End If
        If MaskEdBox3.Text <> MsgText(29) Then
            If strText <> "" Then strText = strText & "　"
            strText = strText & "收文日：" & MaskEdBox3.Text & "~" & MaskEdBox4.Text
        End If
        If MaskEdBox5.Text <> MsgText(29) Then
            If strText <> "" Then strText = strText & "　"
            strText = strText & "發文日：" & MaskEdBox5.Text & "~" & MaskEdBox6.Text
        End If
    Else
        If intChoose = 2 Then
             'Modify by Amy 2022/07/04 +After:=wksAnnuity 加於最後
            If intXlsSheet > 3 Then xlsAnnuity.Worksheets.add After:=wksAnnuity '插入sheet
            Set wksAnnuity = xlsAnnuity.Worksheets(strWkName & intXlsSheet)
            wksAnnuity.Activate
        End If
    End If
    
    'Add by Amy 2020/04/27 從下面搬上來 +往來
    'Modify by Amy 2022/07/04 +if
    If intChoose <> 3 Then
        Select Case Trim(Text3)
            Case "1"
                strChoose = "未收"
            Case "2"
                strChoose = "收回"
            Case "3"
                strChoose = "往來"
        End Select
    End If
    
    'Add by Amy 2022/07/04
    strTpField = "智權人員"
    If stCmp = "LZ" Then strTpField = "介紹人": stCmp = "L"
    'end 2022/07/04
        
    With wksAnnuity
        'Modify by Amy 2022/07/04 調整抬頭顯示
        If intCounter = 1 Then
            '智權人員繳回清單
            If intChoose = 2 Then
                .Range(Chr(intField) & intCounter).Value = "智權人員帳款明細表 (" & strChoose & ")"
                .Range(Chr(intField) & intCounter).Font.Size = 12
                .Range(Chr(intField) & intCounter & ":" & Chr(intField + 11) & intCounter).MergeCells = True
                .Range(Chr(intField) & intCounter & ":" & Chr(intField + 11) & intCounter).HorizontalAlignment = xlCenter
                intCounter = intCounter + 1
                
                .Range(Chr(intField) & intCounter).Value = "智權人員繳回清單"
                .Range(Chr(intField) & intCounter).Font.Size = 10
                .Range(Chr(intField) & intCounter & ":" & Chr(intField + 11) & intCounter).MergeCells = True
                .Range(Chr(intField) & intCounter & ":" & Chr(intField + 11) & intCounter).HorizontalAlignment = xlCenter
                intCounter = intCounter + 1
                
                
                .Range(Chr(intField) & intCounter).Value = CompNameQuery(stCmp, 4)
                .Range(Chr(intField) & intCounter).Font.Color = RGB(255, 0, 0)
                .Range(Chr(intField) & intCounter).Font.Size = 10
                .Range(Chr(intField) & intCounter & ":" & Chr(intField + 11) & intCounter).MergeCells = True
                .Range(Chr(intField) & intCounter & ":" & Chr(intField + 11) & intCounter).HorizontalAlignment = xlCenter
                intCounter = intCounter + 1
          
                .Range(Chr(intField) & intCounter).Value = strText
                .Range(Chr(intField) & intCounter).Font.Size = 10
                .Range(Chr(intField) & intCounter & ":" & Chr(intField + 11) & intCounter).MergeCells = True
                .Range(Chr(intField) & intCounter & ":" & Chr(intField + 11) & intCounter).HorizontalAlignment = xlCenter
                intCounter = intCounter + 1
            '明細
            'Memo by Amy 2022/07/04 避免刪欄後抬頭不見,先寫於A欄再計算置中欄位
            ElseIf intChoose = 1 Then
                strTitleField = Chr(intField)
                .Range(strTitleField & intCounter).Value = "智權人員帳款明細表 (" & strChoose & ")"
                intCounter = intCounter + 1
                    
                'Modify by Amy 2020/07/08 改回輸入,要可輸2公司 原:cboComp Modify by Amy 2020/04/27 原:Text2
                'Modify by Amy 2022/07/04 原未輸公司別,改都顯示
                'If stCmp = MsgText(601) Then
                    .Range(strTitleField & intCounter).Value = CompNameQuery(stCmp, 4) 'A0802Query(stCmp, True) '原:IIf(Text2.Tag = "J", "智權", "台一")
                    intCounter = intCounter + 1
                'End If
                .Range(strTitleField & intCounter).Value = strText
                intCounter = intCounter + 1
            End If 'intChoose = 2
             intTitleR_Fix = intCounter - 1
        '明細抬頭-最後
        ElseIf intChoose = 3 Then
            strTitleField = Chr(Fix((UBound(strFieldN) - intDelF) / 2 + 65))
            For ii = 1 To intTitleR_Fix
                .Range(strTitleField & ii).Value = .Range(Chr(intField) & ii).Value
                .Range(strTitleField & ii).Font.Size = 10
                If ii <= 2 Then
                    .Range(strTitleField & ii).Font.Bold = True
                    If ii = 1 Then
                        .Range(strTitleField & ii).Font.Size = 12
                    Else
                        .Range(strTitleField & ii).Font.Color = RGB(255, 0, 0)
                    End If
                End If
                .Range(strTitleField & ii).HorizontalAlignment = xlCenter
                .Range(Chr(intField) & ii).Value = ""
            Next ii
        End If 'intCounter = 1
        'end 2022/07/04
        
        '明細
        If intChoose = 1 Then
            .Range(Chr(intField) & intCounter).Value = strTpField & "：" & strSales 'Modify by Amy 2022/07/04 原:智權人員
            intCounter = intCounter + 1
            .Range(Chr(intField) & intCounter).Value = "頁　次：" & .HPageBreaks.Count + 1
            intCounter = intCounter + 1
            intTitleR = 2 'Add by Amy 2022/07/04 非設定PrintTitleRows的列數
            
            'Add by Amy 2020/04/27
            If Trim(Text3) = "3" Then
                'Add by Amy 2022/07/11 +未收框線
                strCol = GetFieldStr(GetValue("未收金額"), 65) '超過Z欄轉換
                strCol2 = GetFieldStr(GetValue("可扣稅額"), 65)
                .Range(strCol & intCounter).Value = "未收"
                .Range(strCol & intCounter & ":" & strCol2 & intCounter).MergeCells = True
                .Range(strCol & intCounter & ":" & strCol2 & intCounter).HorizontalAlignment = xlCenter
                'Modify by Amy 2022/07/11 應收應收->收據金額
                strCol = GetFieldStr(GetValue("收據金額"), 65) '超過Z欄轉換
                strCol2 = GetFieldStr(GetValue("應收扣繳"), 65)
                .Range(strCol & intCounter).Value = "收據" 'Modify by Amy 2024/12/24 原:收據 A
                .Range(strCol & intCounter & ":" & strCol2 & intCounter).MergeCells = True
                .Range(strCol & intCounter & ":" & strCol2 & intCounter).HorizontalAlignment = xlCenter
                
                 strCol = GetFieldStr(GetValue("已收服務費"), 65) '超過Z欄轉換
                strCol2 = GetFieldStr(GetValue("已收扣繳"), 65)
                .Range(strCol & intCounter).Value = "已收" 'Modify by Amy 2024/12/24 原:已收 B
                .Range(strCol & intCounter & ":" & strCol2 & intCounter).MergeCells = True
                .Range(strCol & intCounter & ":" & strCol2 & intCounter).HorizontalAlignment = xlCenter
             
                strCol = GetFieldStr(GetValue("銷帳服務費"), 65) '超過Z欄轉換
                strCol2 = GetFieldStr(GetValue("銷退規費"), 65)
                .Range(strCol & intCounter).Value = "銷帳/退" 'Modify by Amy 2024/12/24 原:銷帳/退 C
                .Range(strCol & intCounter & ":" & strCol2 & intCounter).MergeCells = True
                .Range(strCol & intCounter & ":" & strCol2 & intCounter).HorizontalAlignment = xlCenter
                               
                intCounter = intCounter + 1
            End If
            
            'Moidfy by Amy 2020/04/27 加超過Z欄轉換
            For ii = 0 To UBound(strFieldN)
                strCol = Chr(intField + ii)
                If Trim(Text3) = "3" Then strCol = GetFieldStr(ii, 65) '超過Z欄轉換
                
                .Columns(strCol & ":" & strCol).ColumnWidth = intWidth(ii)
                .Range(strCol & intCounter).Value = strFieldN(ii)
                .Range(strCol & intCounter).HorizontalAlignment = xlCenter
            Next ii
            strCol = Chr(UBound(strFieldN) + 65)
            If Trim(Text3) = "3" Then strCol = GetFieldStr(UBound(strFieldN), 65) '超過Z欄轉換
            .Range(Chr(intField) & intCounter & ":" & strCol & intCounter).Select
            'end 2020/04/27
        '繳回清單
        ElseIf intChoose = 2 Then
            .Range(Chr(intField) & intCounter).Value = "列印日期：" & CFDate(ACDate(ServerDate))
            intCounter = intCounter + 1
            .Range(Chr(intField) & intCounter).Value = "列印人員：" & strUserName
            intCounter = intCounter + 1
            .Range(Chr(intField) & intCounter).Value = "頁　次：" & .HPageBreaks.Count + 1
            intCounter = intCounter + 1
            intTitleR = 3 'Add by Amy 2022/07/04 非設定PrintTitleRows的列數
            
            For ii = 0 To 11 Step 2
                'Modify by Amy 2022/07/04 原:"智權人員"
                 .Range(Chr(intField + ii) & intCounter).Value = strTpField
                 .Range(Chr(intField + ii) & intCounter).HorizontalAlignment = xlCenter
                 .Range(Chr(intField + 1 + ii) & intCounter).Value = "繳回日"
                 .Range(Chr(intField + 1 + ii) & intCounter).HorizontalAlignment = xlCenter
            Next ii
            .Range(Chr(intField) & intCounter & ":" & Chr(intField + ii - 1) & intCounter).Select
        End If
        intCounter = intCounter + 1
        
        'Modify by Amy 2022/07/04 +if PrintTitleRows改抓變數
        If intChoose <> 3 Then
    '        If intChoose = 1 Then
    '           .PageSetup.PrintTitleRows = "$1:$2"
    '        ElseIf intChoose = 2 Then
    '             .PageSetup.PrintTitleRows = "$1:$5"
    '        'Add by Amy 2020/04/27 加往來
    '        Else
    '             .PageSetup.PrintTitleRows = "$1:$6"
    '        End If
            .PageSetup.PrintTitleRows = "$1:$" & intTitleR_Fix
            'end 2022/07/04
            .PageSetup.Orientation = xlLandscape '橫印
            .PageSetup.LeftMargin = 28.34
            .PageSetup.RightMargin = 28.34
            .PageSetup.TopMargin = 42.51
            .PageSetup.BottomMargin = 42.51
            .PageSetup.HeaderMargin = 28.34
            .PageSetup.FooterMargin = 28.34
            
            With .Application.Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
        End If
    End With
End Sub

'Private Sub PrintSum1(strTot14 As String, strTot15 As String, strTot16 As String, strTot18 As String, strTot20 As String)
Private Sub PrintSum1(strStart As Integer)
    Dim strTemp As String
    
    With wksAnnuity
        wksAnnuity.Range(Chr(GetValue("案件性質") + 65) & intCounter).Value = "合計"
        strTemp = Chr(intField) & intCounter - 1 & ":" & Chr(UBound(strFieldN) + 65) & intCounter - 1
        .Range(strTemp).Select
        With .Application.Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
         'Modify by Amy 2016/09/19 合計改公式
        .Range(Chr(GetValue("應收金額") + 65) & intCounter).Formula = "=Sum(" & Chr(GetValue("應收金額") + 65) & strStart & ":" & Chr(GetValue("應收金額") + 65) & intCounter - 1 & ")" '.Value = Format(Val(strTot14), DDollar2)
        .Range(Chr(GetValue("應收金額") + 65) & intCounter).NumberFormatLocal = "#,##0"
        .Range(Chr(GetValue("應收服務費") + 65) & intCounter).Formula = "=Sum(" & Chr(GetValue("應收服務費") + 65) & strStart & ":" & Chr(GetValue("應收服務費") + 65) & intCounter - 1 & ")"
        .Range(Chr(GetValue("應收服務費") + 65) & intCounter).NumberFormatLocal = "#,##0"
        .Range(Chr(GetValue("應收規費") + 65) & intCounter).Formula = "=Sum(" & Chr(GetValue("應收規費") + 65) & strStart & ":" & Chr(GetValue("應收規費") + 65) & intCounter - 1 & ")"
        .Range(Chr(GetValue("應收規費") + 65) & intCounter).NumberFormatLocal = "#,##0"
        .Range(Chr(GetValue("已收金額") + 65) & intCounter).Formula = "=Sum(" & Chr(GetValue("已收金額") + 65) & strStart & ":" & Chr(GetValue("已收金額") + 65) & intCounter - 1 & ")"
        .Range(Chr(GetValue("已收金額") + 65) & intCounter).NumberFormatLocal = "#,##0"
        .Range(Chr(GetValue("未收金額") + 65) & intCounter).Formula = "=Sum(" & Chr(GetValue("未收金額") + 65) & strStart & ":" & Chr(GetValue("未收金額") + 65) & intCounter - 1 & ")"
        .Range(Chr(GetValue("未收金額") + 65) & intCounter).NumberFormatLocal = "#,##0"
        .Range(Chr(GetValue("案件規費餘額") + 65) & intCounter).Formula = "=Sum(" & Chr(GetValue("案件規費餘額") + 65) & strStart & ":" & Chr(GetValue("案件規費餘額") + 65) & intCounter - 1 & ")"
        .Range(Chr(GetValue("案件規費餘額") + 65) & intCounter).NumberFormatLocal = "#,##0"
        .Range(Chr(GetValue("點數") + 65) & intCounter).Formula = "=Sum(" & Chr(GetValue("點數") + 65) & strStart & ":" & Chr(GetValue("點數") + 65) & intCounter - 1 & ")"
        .Range(Chr(GetValue("點數") + 65) & intCounter).NumberFormatLocal = "#,##0"
        .Range(Chr(GetValue("扣繳") + 65) & intCounter).Formula = "=Sum(" & Chr(GetValue("扣繳") + 65) & strStart & ":" & Chr(GetValue("扣繳") + 65) & intCounter - 1 & ")"
        .Range(Chr(GetValue("扣繳") + 65) & intCounter).NumberFormatLocal = "#,##0"
    End With
End Sub

'Private Sub PrintSum2(strTot14 As String, strTot15 As String, strTot16 As String, strTot17 As String)
Private Sub PrintSum2(strStart As Integer)
    Dim strTemp As String
    
    With wksAnnuity
        wksAnnuity.Range(Chr(GetValue("案件性質") + 65) & intCounter).Value = "合計"
        strTemp = Chr(intField) & intCounter - 1 & ":" & Chr(UBound(strFieldN) + 65) & intCounter - 1
        .Range(strTemp).Select
        With .Application.Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        'Modify by Amy 2016/09/19 合計改公式
        .Range(Chr(GetValue("收款金額") + 65) & intCounter).Formula = "=Sum(" & Chr(GetValue("收款金額") + 65) & strStart & ":" & Chr(GetValue("收款金額") + 65) & intCounter - 1 & ")"   'Format(Val(strTot14), DDollar2)
        .Range(Chr(GetValue("收款金額") + 65) & intCounter).NumberFormatLocal = "#,##0"
        .Range(Chr(GetValue("溢收金額") + 65) & intCounter).Formula = "=Sum(" & Chr(GetValue("溢收金額") + 65) & strStart & ":" & Chr(GetValue("溢收金額") + 65) & intCounter - 1 & ")"
        .Range(Chr(GetValue("溢收金額") + 65) & intCounter).NumberFormatLocal = "#,##0"
        .Range(Chr(GetValue("服務費") + 65) & intCounter).Formula = "=Sum(" & Chr(GetValue("服務費") + 65) & strStart & ":" & Chr(GetValue("服務費") + 65) & intCounter - 1 & ")"
        .Range(Chr(GetValue("服務費") + 65) & intCounter).NumberFormatLocal = "#,##0"
        .Range(Chr(GetValue("規費") + 65) & intCounter).Formula = "=Sum(" & Chr(GetValue("規費") + 65) & strStart & ":" & Chr(GetValue("規費") + 65) & intCounter - 1 & ")"
        .Range(Chr(GetValue("規費") + 65) & intCounter).NumberFormatLocal = "#,##0"
        .Range(Chr(GetValue("扣繳") + 65) & intCounter).Formula = "=Sum(" & Chr(GetValue("扣繳") + 65) & strStart & ":" & Chr(GetValue("扣繳") + 65) & intCounter - 1 & ")"
        .Range(Chr(GetValue("扣繳") + 65) & intCounter).NumberFormatLocal = "#,##0"
    End With
End Sub

'取得案件名稱
Private Function GetCaseName(strCP0104 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

StrSQLa = "Select Decode(PA05,null,Decode(PA06,null,Nvl(PA07,''),PA06),PA05) as CName From Patent Where " & ChgPatent(strCP0104)
StrSQLa = StrSQLa & " union Select Decode(TM05,null,Decode(TM06,null,Nvl(TM07,''),TM06),TM05) as CName From Trademark Where " & ChgTradeMark(strCP0104)
StrSQLa = StrSQLa & " union Select Decode(LC05,null,Decode(LC06,null,Nvl(LC07,''),LC06),LC05) as CName From Lawcase Where " & ChgLawcase(strCP0104)
StrSQLa = StrSQLa & " union Select Decode(SP05,null,Decode(SP06,null,Nvl(SP07,''),SP06),SP05) as CName From ServicePractice Where " & ChgService(strCP0104)
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
     GetCaseName = "" & rsA.Fields("CName")
Else
    GetCaseName = ""
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'智權人員繳回清單
'Modify by Amy 2020/04/27 +公司別參數
Private Sub PrintSalesListExcel(stCmp As String)
Dim ii As Integer, iIdx As Integer, iPos As Integer
Dim stNames() As String
Dim stWkName As String  'Add by Amy 2022/07/04
   
    'Add by Amy 2022/07/04 LZ為L公司-以介紹人列資料
    stTpCmp = stCmp
    If stTpCmp = "LZ" Then stCmp = "L"
    'end 2022/07/04
    
    stNames = Split(m_strSalesList, vbCrLf)
    If UBound(stNames) <= 0 Then Exit Sub
    
    '設定表頭
    'Modify by Amy 2022/07/04 原:stCmp
    Call SetExcel(2, , , stTpCmp)  'Modify by Amy 2020/04/27 +公司別
    With wksAnnuity
        '一頁印162個智權人員
        For ii = 0 To UBound(stNames)
            If ii > 161 And ii Mod 161 = 0 Then
                '換頁
                .Range("A" & intCounter).Select
                .HPageBreaks.add Before:=.Application.ActiveCell
                'Modify by Amy 2022/07/04 原:stCmp
                Call SetExcel(2, , , stTpCmp) 'Modify by Amy 2020/04/27 +公司別
            End If
            .Range(Chr(intField + (ii Mod 6) * 2) & intCounter).Value = StaffQuery(stNames(ii))
            If ii >= 5 And ii Mod 6 = 5 Then intCounter = intCounter + 1
        Next ii
    End With
    'Modify by Amy 2020/04/27 公司別改抓變數
'    If Text2.Tag = "J" Then
'        wksAnnuity.Name = "智權繳回清單"
'    Else
'        wksAnnuity.Name = "台一繳回清單"
'    End If
    'Modify by Amy 2022/07/04
    stWkName = CompNameQuery(stCmp, 4)  'Modify by Amy 2020/07/08  要可輸2公司 原:A0802Query(stCmp, True) & "繳回清單"
    If stTpCmp = "LZ" Then stWkName = stWkName & "-介紹人"
    stWkName = stWkName & "繳回清單"
    wksAnnuity.Name = stWkName
    'end 2022/07/04
    'end 2020/04/27
    intXlsSheet = intXlsSheet + 1
End Sub
'end 2015/02/10

'Add By Sindy 2016/6/8
Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub
Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 = MsgText(601) Then
      Exit Sub
   End If
   Text11 = "0"
   Text12 = "00"
End Sub
Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub
'2016/6/8 END

' Add by Amy 2020/04/27 產生Excel資料(往來)
Private Sub PrintExcel3()
    Dim strStart As Integer, CountData_S As Integer, j As Integer
    Dim bolIsFirst As Boolean
    Dim strCol As String, strQuery As String, strA As String, strCmp As String, strTmp As String, strSheetN As String
    Dim strHA As String '對齊方式 1.置中/2.靠右
    Dim strVal As String, intMaxTitle As Integer 'Add by Amy 2022/07/11
   
On Error GoTo ErrHnd
    
    'Modify by Amy 2020/07/08 公司別改回輸入(不用下拉)
    strCmp = Trim(Text2) 'Trim(CboComp)
'    If strCmp <> MsgText(60) Then
'        If InStr(strCmp, "　") > 0 Then
'            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
'        End If
'    End If
    If strCmp <> MsgText(601) Then
'        If strCmp <> "J" And strCmp <> "L" Then
'            strA = " And r10603 Not In ('J','L') "
'        Else
            strA = " And r10603='" & strCmp & "' "
'        End If
    End If
    
    'strA = "Select Distinct Decode(r10603,'J','J','L','L','1') as Cmp From accrpt106_1,accrpt106 " & _
                "Where ID = '" & strUserNum & "' And ID=R10601(+) And R001=R10610(+) And R002=R10627(+) " & strA
    strA = "Select Distinct r10603 as Cmp From accrpt106_1,accrpt106 " & _
                "Where ID = '" & strUserNum & "' And ID=R10601(+) And R001=R10610(+) And R002=R10627(+) " & strA
    'Add by Amy 2022/07/11 法律所再產生一份以介紹人為主的資料
    If Trim(Text2) = MsgText(601) Or Trim(Text2) = "L" Then
        If ChkLaw = True Then
            strA = strA & " Union Select 'LZ' as Cmp From Dual Order by Cmp "
        End If
    End If
    'end 2020/07/08
    If rsA.State = adStateOpen Then rsA.Close
    rsA.CursorLocation = adUseClient
    rsA.Open strA, adoTaie, adOpenDynamic, adLockBatchOptimistic
    If rsA.RecordCount = 0 Then
        rsA.Close
        MsgBox MsgText(28), , MsgText(5)
        Exit Sub
    End If
    
    ReDim strFieldN(25)
    ReDim intWidth(25)
    'Modify by Amy 2020/05/28 +申請案號
    'Modify by Amy 2022/07/11 應收帳款->未收金額/應收應收->收據金額;+介紹人/未收服務費/未收規費
    strFieldN = Array("收據抬頭", "公司", "收據日期", "客戶編號", "印", _
                                "介紹人", "收據號碼", "本所案號", "客戶案件案號", "申請案號", _
                                "案件性質", "申請國家", "案件名稱", "未收金額", "未收服務費", _
                                "未收規費", "可扣稅額", "收據金額", "應收服務費", "應收規費", _
                                 "應收扣繳", "已收服務費", "已收規費", "已收扣繳", "案件規費餘額", _
                                "備註", "點數", "銷帳服務費", "銷帳規費", "銷退服務費", _
                                "銷退規費", "發文日", "付款週期")

    intWidth = Array(13, 4.5, 10, 10, 2.5, 8, 10, 13, 14, 10, _
                               10, 10, 10, 10, 10, 10, 10, 10, 10, 10, _
                               10, 10, 10, 10, 14, 8, 11, 10, 10, 10, _
                               10, 8, 10)
    'end 2022/07/11
    'end 2020/05/28
                                
    bolIsFirst = True: intXlsSheet = 1: strText = "": CountData_S = 0: intField = 65: bolHasData = True
    intDelF = 0 'Add by Amy 2020/06/30
    
    rsA.MoveFirst
    For i = 0 To rsA.RecordCount - 1
        '*** 設定Excel檔名 ***
        If i = 0 Then
            strFileName = strExcelPath & "智權人員帳款明細-往來" & Format(Now, "yyyymmddhhmmss") & MsgText(43)
            If Dir(strFileName) = MsgText(601) Then
               If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
                  MkDir strExcelPath
               End If
            Else
               Kill strFileName
            End If
            xlsAnnuity.SheetsInNewWorkbook = 3 'Add by Amy 2020/11/12 2010未設若一進入只有一個sheet 會Error
            xlsAnnuity.Workbooks.add
            xlsAnnuity.Application.WindowState = xlMinimized
        End If
        If strWkName = MsgText(601) Then strWkName = Left(xlsAnnuity.Worksheets(1).Name, Len(xlsAnnuity.Worksheets(1).Name) - 1)
        'Add by Amy 2020/05/28 畫面公司別為空白=全部公司,若全公司都有資料需有6個Sheet,故需加工作表
        'Modify by Amy 2021/05/13 2010需增加工作表會錯,目前財務只有2010及2013版本 拿掉 And intXlsSheet <> 1 And Val(xlsAnnuity.Version) = 15
        If intXlsSheet > 3 Then
            'Modify by Amy 2022/07/04 +After:=wksAnnuity 加於最後
            xlsAnnuity.Worksheets.add After:=wksAnnuity '插入sheet
        End If
        Set wksAnnuity = xlsAnnuity.Worksheets(strWkName & intXlsSheet)
        wksAnnuity.Activate
        
        '*** 抓取各公司別資料 ***
        strQuery = GetSql("" & rsA.Fields("Cmp"))
        '勾選「客戶案件案號」
        If Check4.Value = 1 Then
            strQuery = strQuery & " Order by r10605,r10624 asc,r10608 asc,r10610 asc"
        Else
            strQuery = strQuery & " Order by r10605,r10608 asc,r10610 asc"
        End If
      
        If adoaccrpt106.State <> adStateClosed Then adoaccrpt106.Close
        adoaccrpt106.CursorLocation = adUseClient
        adoaccrpt106.Open strQuery, adoTaie, adOpenDynamic, adLockBatchOptimistic
        With wksAnnuity
            If adoaccrpt106.RecordCount > 0 Then
                strSalesNo = "": strSameName = "":  m_strSalesList = "": strOldArea = "": CountData_S = 0
                intCounter = 1
                adoaccrpt106.MoveFirst
                Do While adoaccrpt106.EOF = False
                    If strSalesNo <> "" & adoaccrpt106.Fields("r10605").Value Then
                        m_strSalesList = m_strSalesList & "" & adoaccrpt106.Fields("r10605").Value & vbCrLf
                        If strSalesNo = "" Then
                            Call SetExcel(1, bolIsFirst, A0902Query("" & adoaccrpt106.Fields("r10604").Value) & StaffQuery("" & adoaccrpt106.Fields("r10605").Value) & " " & adoaccrpt106.Fields("r10605").Value, "" & rsA.Fields("Cmp"))
                            strStart = intCounter
                        End If
                    End If
                    '智權人員不同時換頁/業務區不同也換頁(ex:86052 104年及105年不同區的資料)
                    'Modify by Amy 2022/07/04 改變數 intCounter > 7/CountData_S = 27 ->改顯示資料25列
                    intMaxTitle = intTitleR_Fix + intTitleR + 2
                    If (intCounter > intMaxTitle + 1 And strSalesNo <> "" & adoaccrpt106.Fields("r10605").Value) Or (intCounter > intMaxTitle + 1 And strOldArea <> "" & adoaccrpt106.Fields("r10604").Value) _
                      Or CountData_S = 32 - (intMaxTitle + 1) Then
                        If intCounter > intMaxTitle + 1 And (strSalesNo <> "" & adoaccrpt106.Fields("r10605").Value Or strOldArea <> "" & adoaccrpt106.Fields("r10604").Value) Then
                            Call PrintSum3(strStart)
                            intCounter = intCounter + 1
                            strStart = intCounter + intTitleR + 2 'Modify by Amy 2022/07/11 原:intCounter
                        End If
    
                        '換頁
                        .Range("A" & intCounter).Select
                        .HPageBreaks.add Before:=.Application.ActiveCell
                        Call SetExcel(1, bolIsFirst, A0902Query("" & adoaccrpt106.Fields("r10604").Value) & StaffQuery("" & adoaccrpt106.Fields("r10605").Value) & " " & adoaccrpt106.Fields("r10605").Value, "" & rsA.Fields("Cmp"))
                        CountData_S = 0
                    End If
                    For j = LBound(strFieldN) To UBound(strFieldN)
                        strTmp = "": strHA = ""
                        strTmp = "" & adoaccrpt106.Fields(j)
                        Select Case j
                             Case GetValue("收據抬頭")
                                strSameName = strTmp
                                strTmp = Left(strTmp, 6)
                            Case GetValue("印")
                                If IsNull(adoaccrpt106.Fields("r10623").Value) = False Then
                                    strTmp = "N"
                                End If
                            'Add by Amy 2022/07/04 公司別為L顯示,介紹人
                            Case GetValue("介紹人")
                                If "" & rsA.Fields("Cmp") = "L" Then
                                    If IsNull(adoaccrpt106.Fields("r10631").Value) = False Then
                                        strTmp = "" & adoaccrpt106.Fields("st02").Value
                                    End If
                                End If
                            Case GetValue("申請國家") 'GetValue("案件性質")'Modify by Amy 2024/12/24 全顯示不取字-瑞婷
                               strTmp = Left(strTmp, 4)
                            Case GetValue("案件名稱")
                                If strTmp = MsgText(601) Then
                                    strTmp = StrToStr(GetCaseName("" & adoaccrpt106.Fields("r10611").Value), 20)
                                End If
                            'Modify by Amy 2022/07/11 應收帳款->未收金額
                            Case GetValue("未收金額"), GetValue("可扣稅額")
                                If strTmp <> MsgText(601) Then
                                    strTmp = Format(strTmp, DDollar2)
                                Else
                                    strTmp = PUB_ChkExcelZero(1)
                                End If
                                strHA = "2"
                            'Add by Amy 2022/07/11 +未收服務費/未收規費
                            Case GetValue("未收服務費")
                                strTmp = "=" & GetFieldStr(GetValue("應收服務費"), intField) & intCounter & "-" & _
                                                         GetFieldStr(GetValue("已收服務費"), intField) & intCounter
                                strHA = "2"
                            Case GetValue("未收規費")
                                strTmp = "=" & GetFieldStr(GetValue("應收規費"), intField) & intCounter & "-" & _
                                                         GetFieldStr(GetValue("已收規費"), intField) & intCounter
                                strHA = "2"
                            'end 2022/07/11
                            'Modify by Amy 2022/07/11 應收應收->收據金額
                            Case GetValue("收據金額"), GetValue("應收服務費"), GetValue("應收規費"), GetValue("應收扣繳")
                                If strTmp <> MsgText(601) Then
                                    strTmp = Format(strTmp, DDollar2)
                                Else
                                    strTmp = PUB_ChkExcelZero(1)
                                End If
                                strHA = "2"
                            Case GetValue("已收服務費"), GetValue("已收規費"), GetValue("已收扣繳")
                                If strTmp <> MsgText(601) Then
                                    strTmp = Format(strTmp, DDollar2)
                                Else
                                    strTmp = PUB_ChkExcelZero(1)
                                End If
                                strHA = "2"
                            'Modify by Amy 2022/07/11應收帳款->未收金額
                            Case GetValue("未收金額"), GetValue("案件規費餘額"), GetValue("點數"), GetValue("付款週期")
                                If strTmp <> MsgText(601) Then
                                    strTmp = Format(strTmp, DDollar2)
                                Else
                                    strTmp = PUB_ChkExcelZero(1)
                                End If
                                strHA = "2"
                            Case GetValue("銷帳服務費"), GetValue("銷帳規費"), GetValue("銷退服務費"), GetValue("銷退規費")
                                If strTmp <> MsgText(601) Then
                                    strTmp = Format(strTmp, DDollar2)
                                Else
                                    strTmp = PUB_ChkExcelZero(1)
                                End If
                                strHA = "2"
                        End Select
                        
                        strCol = GetFieldStr(j, 65) '超過Z欄轉換
                        'Add by Amy 2020/05/28 +申請案號
                        If j = GetValue("申請案號") Then
                            .Range(strCol & intCounter).NumberFormatLocal = "@"
                        End If
                        'end 2020/05/28
                        .Range(strCol & intCounter).Value = strTmp
                        If strHA = "1" Then
                            .Range(strCol & intCounter).HorizontalAlignment = xlCenter
                        ElseIf strHA = "2" Then
                            .Range(strCol & intCounter).HorizontalAlignment = xlRight
                        End If
                        'Add by Amy 2022/07/11 不顯示公式(財務會依不同狀況刪欄位),並確認未收金額是否正確
                        If j = UBound(strFieldN) Then
                            strTmp = GetFieldStr(GetValue("未收金額"), intField) & intCounter
                            strVal = Val(GetFieldStr(GetValue("未收服務費"), intField) & intCounter) + Val(GetFieldStr(GetValue("未收規費"), intField) & intCounter)
                            If Val(strTmp) <> Val(strVal) Then
                                '公式加總<>未收金額,顯示紅色
                                .Range(GetFieldStr(GetValue("未收金額"), intField) & intCounter).Font.Color = vbRed
                            End If
                            '不顯示公式
                            .Range(GetFieldStr(GetValue("未收服務費"), intField) & intCounter).Copy
                            .Range(GetFieldStr(GetValue("未收服務費"), intField) & intCounter).PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
                           
                            .Range(GetFieldStr(GetValue("未收規費"), intField) & intCounter).Copy
                            .Range(GetFieldStr(GetValue("未收規費"), intField) & intCounter).PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
                        End If
                    Next j
                    intCounter = intCounter + 1
                    CountData_S = CountData_S + 1 '計算智權人員明細筆數 for 換頁
                    strSalesNo = adoaccrpt106.Fields("r10605").Value
                    strOldArea = adoaccrpt106.Fields("r10604").Value '業務區
                    adoaccrpt106.MoveNext
                Loop
                '合計
                Call PrintSum3(strStart) '合計公式
            End If
            'Add by Amy 2022/07/11 公司別非L公司,刪除「介紹人」欄
            If "" & rsA.Fields("Cmp") <> "L" Then
                strCol = GetFieldStr(GetValue("介紹人"), 65 - intDelF)
                .Range(strCol & ":" & strCol).Delete Shift:=xlToLeft
                intDelF = intDelF + 1
            End If
            'Modify by Amy 2020/06/30 加- intDelF 因欄位全顯示,若未勾「客戶案件案號」及「申請案號」欄,刪申請案號欄時會刪錯
            If Check4.Value = 0 Then
                strCol = GetFieldStr(GetValue("客戶案件案號"), 65 - intDelF)
                .Range(strCol & ":" & strCol).Delete Shift:=xlToLeft
                intDelF = intDelF + 1
            End If
            'Add by Amy 2020/05/28 +申請案號
            If Check5.Value = 0 Then
                strCol = GetFieldStr(GetValue("申請案號"), 65 - intDelF)
                .Range(strCol & ":" & strCol).Delete Shift:=xlToLeft
                intDelF = intDelF + 1
            End If
            'end 2020/06/30
        End With
        strCol = GetFieldStr(UBound(strFieldN), 65)
        'Modify by Amy 2022/07/11 固定抬頭改抓變數 /+ rsA.Fields("Cmp") = "LZ" 判斷/SetExcel(3)
        Call SetExcel(3)
        wksAnnuity.Range(Chr(intField) & intTitleR_Fix + 1 & ":" & strCol & intCounter).Font.Size = 12
        If "" & rsA.Fields("Cmp") = "LZ" Then
            wksAnnuity.Name = CompNameQuery("L", 4) & "-介紹人"
        Else
            wksAnnuity.Name = CompNameQuery("" & rsA.Fields("Cmp"), 4) 'Modify by Amy 2020/07/08 要可輸2公司 原:A0802Query("" & rsA.Fields("Cmp"), True)
        End If
        'end 2022/07/11
        intXlsSheet = intXlsSheet + 1
        intDelF = 0 'Add by Amy 2020/07/28 每印完一個公司欄位刪除欄位要設0
        bolIsFirst = False
       
        rsA.MoveNext
    Next i
    
    '判斷版本
    If Val(xlsAnnuity.Version) < 12 Then
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
    Else
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
    End If
    xlsAnnuity.Workbooks.Close
    xlsAnnuity.Quit
    'Modify by Amy 2021/06/22 路徑改中文字顯示
    If bolHasData = True Then MsgBox "Excel檔已產生!" & vbCrLf & vbCrLf & strExcelPathN & Replace(strFileName, strExcelPath, ""), vbInformation, Me.Caption
    adoaccrpt106.Close
    Set xlsAnnuity = Nothing
    Set wksAnnuity = Nothing
    Exit Sub
   
ErrHnd:
    xlsAnnuity.Visible = True
    If adoaccrpt106.State <> adStateClosed Then adoaccrpt106.Close
    If Val(xlsAnnuity.Version) < 12 Then
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
    Else
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
    End If
    xlsAnnuity.Workbooks.Close
    xlsAnnuity.Quit
    Set xlsAnnuity = Nothing
    Set wksAnnuity = Nothing
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'傳入公司別及收款單號回傳 傳票編號-for 往來
Private Function GetA1P22(ByVal stA1p01 As String, ByVal stA1P04 As String) As String
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    strQ = "Select Distinct a1p22 From Acc1P0 " & _
                "Where a1p01='" & stA1p01 & "' And a1p04='" & stA1P04 & "' "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        GetA1P22 = "" & RsQ.Fields("a1p22")
    End If
    RsQ.Close
    
End Function

'Add by Amy 2020/04/27
'*************************************************
'  列印類別-選擇往來
'  (參考 Frmacc1220,若此有修改需看Frmacc1220是否要改)
'*************************************************
Private Sub Select3()
    Dim RsQ As New ADODB.Recordset, rsA As New ADODB.Recordset
    Dim intQ As Integer, intA As Integer, strCP(4) As String
    Dim strQ As String, strA As String, strUpd As String, strF1 As String, strF2 As String, strF3 As String, strF4 As String
    Dim strWhere1 As String, strWhere2 As String, strWhere3 As String
    Dim strCmp As String, strCustCaseNo As String, strTmp As String
    Dim strReceAmt As String, strWriteOff As String, strNotPay As String '應收金額/銷帳金額/應收帳款->未收金額(未收)
    Dim strWhere_S(1 To 3) As String 'Add by Amy 2022/07/20
    
    'Modify by Amy 2022/07/20 原畫面條件改至function
    Call GetWhere3(False, strWhere1, strWhere2, strWhere3)
    '有下智權人員,抓案源資料
    If Trim(Text1) <> MsgText(601) Then
        Call GetWhere3(True, strWhere_S(1), strWhere_S(2), strWhere_S(3))
    End If
    'end 2022/07/20
    
    '固定欄位
    strF1 = " r10601,r10603,r10604,r10605,r10606,"
    strF2 = "'" & strUserNum & "',a0k11,a0k22,  a0k20,  a0k03,"
    strF1 = strF1 & "r10607,r10610,r10611,r10612, r10613,"
    strF2 = strF2 & "a0k04,a0k01,a0j02,GetCp10Desc(cp01,cp10,a0j04) cp10N,na03,"
    strF1 = strF1 & "r10619,r10620,r10621,r10622,r10623,"
    strF2 = strF2 & "a0k08,Nvl(Round(a0j09/1000,1),0) cp18,cp27,Decode(substr(a0k03,1,1),'X',nvl(cu175,2),null) cu175,a0k32,"
    strF1 = strF1 & "r10627,r10628,r10629,"
    strF2 = strF2 & "a0j01,a0j07,a0k05,"
    
    strQ = "Insert Into Accrpt106 (" & strF1 & "r10608,r10609) "
'*** Where 同 Frmacc1220  AdodcRefresh ***

    '收據(舊收據無a0j01)
    strQ = strQ & "Select " & strF2 & "a0k02,a0k01 From Acc0k0,Acc0j0,CaseProgress,Customer,Staff,Nation " & _
                        "Where (a0k09 is null or a0k09 = 0) And a0k01=a0j13(+) And a0j01=cp09(+) " & _
                        "And na01(+)=a0j04 And a0k20=st01(+) And CU01(+)=SubStr(a0k03,1,8) And SubStr(a0k03,9,1)=CU02(+) " & strWhere1
    'Add by Amy 2022/07/20 有下智權人員,抓案源資料(Memo by Amy 2022/07/26 以介紹人的所別為主-與Morgan討論)
    If Trim(Text1) <> MsgText(601) Then
        strQ = strQ & " Union " & _
                        "Select " & strF2 & "a0k02,a0k01 From Customer,Staff,Nation,(" & GetCaseSource("3.1", strWhere_S(1)) & ") " & _
                        "Where na01(+)=a0j04 And los04=st01(+) And CU01(+)=SubStr(a0k03,1,8) And SubStr(a0k03,9,1)=CU02(+) "
        If pub_strUserOffice <> "1" Then
            strQ = strQ & " And ''||ST06='" & pub_strUserOffice & "' "
        End If
    End If
    '收款
    strQ = strQ & " Union " & _
                    "Select " & strF2 & "a0k02,a0L01 From Acc0l0,Acc1u0,Acc0k0,Acc0j0,CaseProgress,Customer,Staff,Nation " & _
                    "Where (a0k09 is null or a0k09 = 0) And Nvl(a0k17,0)+Nvl(a0k18,0)>0 And a1u01=a0l01(+) " & _
                    "And a0k01=a0j13(+) And a0j01=cp09(+) And a0j13=a1u02(+) And a0j01=a1u03(+) And SubStr(a1u01,1,1)='F' " & _
                    "And na01(+)=a0j04 And a0k20=st01(+) And SubStr(a0k03,1,8)=CU01(+) And SubStr(a0k03,9,1)=CU02(+) " & strWhere2
    'Add by Amy 2022/07/20 有下智權人員,抓案源資料
    If Trim(Text1) <> MsgText(601) Then
        strQ = strQ & " Union " & _
                    "Select " & strF2 & "a0k02,a0k01 From Customer,Staff,Nation,(" & GetCaseSource("3.2", strWhere_S(2)) & ") " & _
                    "Where na01(+)=a0j04 And los04=st01(+) And SubStr(a0k03,1,8)=CU01(+) And SubStr(a0k03,9,1)=CU02(+) "
        If pub_strUserOffice <> "1" Then
            strQ = strQ & " And ''||ST06='" & pub_strUserOffice & "' "
        End If
    End If
    '銷退
     strQ = strQ & " Union " & _
                    "Select " & strF2 & "a0k02,a0S01 From Acc0s0,Acc1u0,Acc0k0,Acc0j0,CaseProgress,Customer,Staff,Nation " & _
                    "Where (a0k09 is null or a0k09 = 0) And a0k10 is not null And (Nvl(cp77, 0) <> 0 or Nvl(cp78, 0) <> 0) And a1u01=a0s01(+) And a1u02=a0s02(+) " & _
                    "And a0k01=a0j13(+) And a0j01=cp09(+) And a0j13=a1u02(+) And a0j01=a1u03(+) And SubStr(a1u01,1,1)='I' " & _
                    "And na01(+)=a0j04 And a0k20=st01(+) And SubStr(a0k03,1,8)=CU01(+) And SubStr(a0k03,9,1)=CU02(+) " & strWhere3
    'Add by Amy 2022/07/20 有下智權人員,抓案源資料
    If Trim(Text1) <> MsgText(601) Then
        strWhere3 = Replace(strWhere3, "a0k20", "Los04")
        strQ = strQ & " Union " & _
                    "Select " & strF2 & "a0k02,a0S01 From Customer,Staff,Nation,(" & GetCaseSource("3.3", strWhere_S(3)) & ") " & _
                    "Where na01(+)=a0j04 And los04=st01(+) And SubStr(a0k03,1,8)=CU01(+) And SubStr(a0k03,9,1)=CU02(+) "
        If pub_strUserOffice <> "1" Then
            strQ = strQ & " And ''||ST06='" & pub_strUserOffice & "' "
        End If
    End If
    adoTaie.Execute strQ
'*** End Where同 Frmacc1220  AdodcRefresh ***

    'Add by Amy 2022/07/11 更新「介紹人 」
    strQ = "Update Accrpt106 a Set r10631=(Select Los04 From Accrpt106,LawOfficeSource Where r10601='" & strUserNum & "' And r10627 Is Not Null " & _
                                                                    "And r10627=Los06 And a.r10603=r10603 And a.r10609=r10609 And a.r10610=r10610 And a.r10605=r10605 And a.r10627=r10627) " & _
                     "Where r10601='" & strUserNum & "' And r10603='L' And r10627 Is Not Null "
    adoTaie.Execute strQ
    
    '新增資料至暫存檔
    'ex:X23402030 1090304~0304 E10905267 於Frmacc1220 I編號仍會列一筆資料,瑞婷說此不需
    strQ = "Insert Into Accrpt106_1 (ID,R001,R002) " & _
                "Select Distinct '" & strUserNum & "',r10610,r10627 From Accrpt106 Where r10601='" & strUserNum & "' "
    adoTaie.Execute strQ
    
    'Memo X23402030 收據日:1090304 E105267 / 收款日:1030916 F10902325 / 銷退日:1090304 I10902325
    'Frmacc1220 收據日期區間會顯示所有欄位金額,但只下1030916 只會有已收金額；若區間只有銷帳/退資料只顯示銷帳/退金額(ex:X55361 1030114 E10221507)
    '此支與Frmacc1220 不同處:先抓收據日、收款日、銷帳/退日於畫面區間資料,將所有金額全部帶出-瑞婷
    strQ = ""
    '應收 服務費/規費/金額
    strQ = strQ & ",(R003,R004,R005)=(Select a0j09,a0j10,Nvl(a0j09, 0)+Nvl(a0j10, 0) From Acc0j0 Where a0j13=R001 and a0j01=R002)"
    '已收 服務費/規費/扣繳
    strQ = strQ & ",(R007,R008,R009)=(Select Nvl(Sum(a1u04),0),Nvl(Sum(a1u05),0),Nvl(Sum(a1u06),0) " & _
                                                            "From Acc1u0 Where  a1u02=R001 and a1u03=R002 And SubStr(a1u01,1,1)='F' " & _
                                                            "Group by R001,R002)"
    '銷帳/退 服務費/規費
    strQ = strQ & ",(R011,R012,R013,R014)=(Select Nvl(Sum(a1u07),0),Nvl(Sum(a1u09),0),Nvl(Sum(a1u08),0),Nvl(Sum(a1u10),0) " & _
                                                                      "From Acc1u0 Where a1u02=R001 and a1u03=R002 And SubStr(a1u01,1,1)='I' " & _
                                                                      "Group by R001,R002) "
                                                                      
    '更新金額欄位
    strQ = "Update Accrpt106_1 Set " & Mid(strQ, 2) & " Where ID='" & strUserNum & "' "
    adoTaie.Execute strQ

    '更新相關資料
    'Modify by Amy 2020/07/13 拿掉SubStr(r10609,1,1)='E',否則下1090101-0710 X42376, E10827049 不會更新到資料
    strQ = "Select r10603 as Cmp,r10628 as A0J07,r10611 as CaseNo,r10627 as CP09,r10629 as A0K05,R001,R003,R004,R005,R007,R008,R011,R012,R013,R014 " & _
                "From Accrpt106_1,Accrpt106 Where ID='" & strUserNum & "' And ID=r10601(+) And R001=r10610(+) And R002=r10627(+) "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        RsQ.MoveFirst
        Do While RsQ.EOF = False
            strUpd = "": strCustCaseNo = ""
            'Modify by Amy 2020/05/28
            strAppNo = "申請案號": strUpd = ""
            '*** 客戶案件案號 /申請案號 ***
            Call CaseNameQuery("" & RsQ.Fields("CP09"), 1, strCustCaseNo, strAppNo)
            If strCustCaseNo <> MsgText(601) Then
                strUpd = strUpd & ",r10624='" & strCustCaseNo & "'"
            End If
            If strAppNo <> MsgText(601) Then
                strUpd = strUpd & ",r10630='" & strAppNo & "'"
            End If
            If strUpd <> MsgText(601) Then
                strUpd = "Update Accrpt106 Set " & Mid(strUpd, 2) & " Where R10601='" & strUserNum & "' And R10611='" & RsQ.Fields("CaseNo") & "' "
                adoTaie.Execute strUpd
            End If
            'end 2020/05/28
            '*** End 客戶案件案號 /申請案號 ***
            
            strUpd = ""
             '*** 案件規費餘額 ***
            strCP(0) = "" & RsQ.Fields("CaseNo")
            strCP(1) = Mid(strCP(0), 1, Len(strCP(0)) - 9)
            strCP(2) = Mid(strCP(0), Len(strCP(0)) - 8, 6)
            strCP(3) = Mid(strCP(0), Len(strCP(0)) - 2, 1)
            strCP(4) = Mid(strCP(0), Len(strCP(0)) - 1, Len(strCP(0)))
            strA = GetCaseFeesBalance(strCP(1), strCP(2), strCP(3), strCP(4))
            intA = 1
            Set rsA = ClsLawReadRstMsg(intA, strA)
            If intA = 1 Then
                rsA.MoveFirst
                Do While rsA.EOF = False
                    If IsNull(rsA.Fields(0).Value) = False Then
                        If Val(rsA.Fields(0).Value) < 0 Then
                            '此案號此區間的第一筆寫入規費餘額(sum(ax207)-sum(ax206))即可E09921672(P096224)
                             strUpd = "Update Accrpt106 Set R10618=" & rsA.Fields(0) & " Where R10601='" & strUserNum & "' And R10611='" & strCP(0) & "' " & _
                                            "And r10627=(Select Max(r10627) From Accrpt106 Where R10601='" & strUserNum & "' And R10611='" & strCP(0) & "' )"
                            adoTaie.Execute strUpd
                       End If
                    End If
                    rsA.MoveNext
                Loop
            End If
        '*** End 案件規費餘額 ***
        
            strUpd = ""
        '*** 應收服務費/規費 ***
            '選擇「服務費、規費合併」
            If Option1(0).Value = True And "" & RsQ.Fields("A0J07") = "Y" Then
                'P119680 a0k30不等於a0j07,先抓a0j07否則非台灣案扣繳金額會錯
                '合併時,應收規費0 金額列於應收服務費
                '應收服務費
                strUpd = strUpd & ",R003=" & Val("" & RsQ.Fields("R003")) + Val("" & RsQ.Fields("R004"))
                '應收規費
                strUpd = strUpd & ",R004=0"
                'Modify by Amy 2020/05/28 已收之服務費規費同上處理
                '已收服務費
                strUpd = strUpd & ",R007=" & Val("" & RsQ.Fields("R007")) + Val("" & RsQ.Fields("R008"))
                '已收規費
                strUpd = strUpd & ",R008=0"
            End If
        '*** End 應收服務費/規費 ***
            
        '*** 應收帳款->未收金額(未收) ***
            'Modify by Amy 2020/05/20 拿if 判斷, 台灣案應收金額會錯 ex:69005 10701/04 E10702272
            'If "" & RsQ.Fields("A0J07") = "Y" Then
                '應收金額
                strReceAmt = "" & RsQ.Fields("R005")
        
                '銷帳金額
                strWriteOff = Val("" & RsQ.Fields("R011")) + Val("" & RsQ.Fields("R012"))
                
                '應收金額-已收(已收服務費+已收規費)-銷帳(銷帳服務費+銷帳規費)+銷退(銷退服務費+銷退規費)
                strTmp = Val("" & RsQ.Fields("r007")) + Val("" & RsQ.Fields("R008")) '已收
                strNotPay = Val(strReceAmt) - Val(strTmp) - Val(strWriteOff) + Val("" & RsQ.Fields("R013")) + Val("" & RsQ.Fields("R014"))
'            Else
'                '應收金額
'                strReceAmt = "" & RsQ.Fields("R003")
'
'                '銷帳金額
'                strWriteOff = Val("" & RsQ.Fields("R011"))
'
'                '應收服務費-已收服務費-銷帳服務費+銷退服務費
'                strTmp = Val("" & RsQ.Fields("r007"))
'                strNotPay = Val(strReceAmt) - Val(strTmp) - Val(strWriteOff) + Val("" & RsQ.Fields("R013"))
'            End If
            strUpd = strUpd & ",R010=" & strNotPay
        '*** End 應收帳款->未收金額 ***
        
        '*** 可扣稅額 / 應收扣繳(有修改需看 Select1 是否需修改) ***
            '公司別為 J or a0k05為 個人,扣繳固定為 0
            If "" & RsQ.Fields("Cmp") = "J" Or "" & RsQ.Fields("A0K05") = "1" Then
                strUpd = strUpd & ",R006=0,R015=0"
            'a0k30不等於a0j07(是否合併),先抓a0j07否則非台灣案扣繳金額會錯
            ElseIf "" & RsQ.Fields("A0J07") = "Y" Then
                If Val(strReceAmt) <= Val(strWriteOff) Then
                    strUpd = strUpd & ",R006=0,R015=0"
                Else
                    '應收扣繳=應收金額/10
                    strUpd = strUpd & ",R006=" & Val(strReceAmt) / 10
                    '可扣稅額=應收帳款->未收金額/10
                    strUpd = strUpd & ",R015=" & Val(strNotPay) / 10
                End If
            Else
                If Val("" & RsQ.Fields("R003")) = 0 Then
                    strUpd = strUpd & ",R006=0,R015=0"
                Else
                    '應收扣繳=應收服務費/10
                    strUpd = strUpd & ",R006=" & Val(RsQ.Fields("R003").Value) / 10
                    'Modify by Amy 2020/12/25 +if  strNotPay > 0
                    If strNotPay > 0 Then
                        '未付-可扣稅額=應收服務費/10
                        strUpd = strUpd & ",R015=" & Val(RsQ.Fields("R003").Value) / 10
                    Else
                        '已付-可扣稅額=應收帳款->未收金額/10
                        strUpd = strUpd & ",R015=" & Val(strNotPay) / 10
                    End If
                End If
            End If
        '*** End 扣繳 ***
        
            If strUpd <> MsgText(601) Then
                strUpd = "Update Accrpt106_1 Set " & Mid(strUpd, 2) & " Where ID='" & strUserNum & "' " & _
                               "And R001='" & RsQ.Fields("R001") & "' And R002='" & RsQ.Fields("CP09") & "' "
                adoTaie.Execute strUpd
            End If
            RsQ.MoveNext
        Loop
    End If
End Sub

Private Function GetSql(ByVal stCmp As String) As String
    Dim stField As String
    
    'Modify by Amy 2020/07/28 每一個公司一個sheet顯示(原:1090101-0728 X811630 會有商標工作表,顯示1及2公司;智慧所工作表顯示2公司-重覆bug)-瑞婷
'    If stCmp = "1" Then
'        GetSql = " And r10603<>'J' And r10603<>'L' "
'    Else
        GetSql = " And r10603='" & stCmp & "' "
'    End If
    'end 2020/07/28
    
    '依欄位順序寫語法
    'Modify by Amy 2020/05/28 +申請案號 r10630
    'Modify by Amy 2022/07/11 +介紹人r10631/0 as NotPayS(未收服務費)/0 as NotPayP(未收規費)
    stField = "R10607,R10603,R10608,R10606,R10623,R10631,R10610,R10611,R10624,R10630,R10612,R10613,'' as CaseName," & _
                  "R010,0 as NotPayS,0 as NotPayP,R015,R005,R003,R004,R006,R007,R008,R009,R10618,R10619,R10620,R011,R012,R013,R014,sqlDateT(R10621),R10622,st02,"
    If stCmp = "LZ" Then
        '介紹人以目前區為主-瑞婷
        stField = stField & "r10631 as r10605,st15 as r10604"
        GetSql = " And r10603='L'  And r10631 Is Not Null "
    Else
        stField = stField & "R10605,R10604"
    End If
    GetSql = "Select Distinct " & stField & " " & _
                    "From Accrpt106_1,Accrpt106,Staff " & _
                    "Where ID='" & strUserNum & "' And ID=R10601(+) And R001=R10610(+) And R002=R10627(+) And r10631=st01(+) " & GetSql
    
End Function

Private Sub PrintSum3(strStart As Integer)
    Dim strColS As String, strColE As String
    Dim strCol As String, strTemp As String
    
    With wksAnnuity
        strCol = GetFieldStr(GetValue("案件性質"), 65)  '超過Z欄轉換
        strTemp = GetFieldStr(UBound(strFieldN), 65)
        wksAnnuity.Range(strCol & intCounter).Value = "合計"
        strTemp = Chr(intField) & intCounter - 1 & ":" & strTemp & intCounter - 1
        .Range(strTemp).Select
        With .Application.Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
         '合計公式
         'Modify by Amy 2022/07/11 應收帳款->未收金額;+未收服務費/未收規費
         strCol = GetFieldStr(GetValue("未收金額"), 65)
         strColS = strCol
        .Range(strCol & intCounter).Formula = "=Sum(" & strCol & strStart & ":" & strCol & intCounter - 1 & ")"
        .Range(strCol & intCounter).NumberFormatLocal = "#,##0"
         strCol = GetFieldStr(GetValue("未收服務費"), 65)
        .Range(strCol & intCounter).Formula = "=Sum(" & strCol & strStart & ":" & strCol & intCounter - 1 & ")"
        .Range(strCol & intCounter).NumberFormatLocal = "#,##0"
         strCol = GetFieldStr(GetValue("未收規費"), 65)
        .Range(strCol & intCounter).Formula = "=Sum(" & strCol & strStart & ":" & strCol & intCounter - 1 & ")"
        .Range(strCol & intCounter).NumberFormatLocal = "#,##0"
        'end 2022/07/11
        strCol = GetFieldStr(GetValue("可扣稅額"), 65)
        strColE = strCol
        .Range(strCol & intCounter).Formula = "=Sum(" & strCol & strStart & ":" & strCol & intCounter - 1 & ")"
        .Range(strCol & intCounter).NumberFormatLocal = "#,##0"
        'Mark by Amy 2024/12/24 不畫框-瑞婷
        'Call SetFrame(strColS, strColE, Val(strStart) - 2, intCounter) 'Add by Amy 2022/07/11 框線
        
        'Modify by Amy 2022/07/11 應收應收->收據金額
        strCol = GetFieldStr(GetValue("收據金額"), 65)
        strColS = strCol
        .Range(strCol & intCounter).Formula = "=Sum(" & strCol & strStart & ":" & strCol & intCounter - 1 & ")"
        .Range(strCol & intCounter).NumberFormatLocal = "#,##0"
        strCol = GetFieldStr(GetValue("應收服務費"), 65)
        .Range(strCol & intCounter).Formula = "=Sum(" & strCol & strStart & ":" & strCol & intCounter - 1 & ")"
        .Range(strCol & intCounter).NumberFormatLocal = "#,##0"
        strCol = GetFieldStr(GetValue("應收規費"), 65)
        .Range(strCol & intCounter).Formula = "=Sum(" & strCol & strStart & ":" & strCol & intCounter - 1 & ")"
        .Range(strCol & intCounter).NumberFormatLocal = "#,##0"
        strCol = GetFieldStr(GetValue("應收扣繳"), 65)
        strColE = strCol
        .Range(strCol & intCounter).Formula = "=Sum(" & strCol & strStart & ":" & strCol & intCounter - 1 & ")"
        .Range(strCol & intCounter).NumberFormatLocal = "#,##0"
        'Mark by Amy 2024/12/24 不畫框-瑞婷
        'Call SetFrame(strColS, strColE, Val(strStart) - 2, intCounter) '框線
        
        strCol = GetFieldStr(GetValue("已收服務費"), 65)
        strColS = strCol
        .Range(strCol & intCounter).Formula = "=Sum(" & strCol & strStart & ":" & strCol & intCounter - 1 & ")"
        .Range(strCol & intCounter).NumberFormatLocal = "#,##0"
        strCol = GetFieldStr(GetValue("已收規費"), 65)
        .Range(strCol & intCounter).Formula = "=Sum(" & strCol & strStart & ":" & strCol & intCounter - 1 & ")"
        .Range(strCol & intCounter).NumberFormatLocal = "#,##0"
        strCol = GetFieldStr(GetValue("已收扣繳"), 65)
        strColE = strCol
        .Range(strCol & intCounter).Formula = "=Sum(" & strCol & strStart & ":" & strCol & intCounter - 1 & ")"
        .Range(strCol & intCounter).NumberFormatLocal = "#,##0"
        'Mark by Amy 2024/12/24 不畫框-瑞婷
        'Call SetFrame(strColS, strColE, Val(strStart) - 2, intCounter) '框線
         
        strCol = GetFieldStr(GetValue("案件規費餘額"), 65)
        .Range(strCol & intCounter).Formula = "=Sum(" & strCol & strStart & ":" & strCol & intCounter - 1 & ")"
        .Range(strCol & intCounter).NumberFormatLocal = "#,##0"
        strCol = GetFieldStr(GetValue("點數"), 65)
        .Range(strCol & intCounter).Formula = "=Sum(" & strCol & strStart & ":" & strCol & intCounter - 1 & ")"
        .Range(strCol & intCounter).NumberFormatLocal = "#,##0"
        
        strCol = GetFieldStr(GetValue("銷帳服務費"), 65)
        strColS = strCol
        .Range(strCol & intCounter).Formula = "=Sum(" & strCol & strStart & ":" & strCol & intCounter - 1 & ")"
        .Range(strCol & intCounter).NumberFormatLocal = "#,##0"
        strCol = GetFieldStr(GetValue("銷帳規費"), 65)
        .Range(strCol & intCounter).Formula = "=Sum(" & strCol & strStart & ":" & strCol & intCounter - 1 & ")"
        .Range(strCol & intCounter).NumberFormatLocal = "#,##0"
        
        strCol = GetFieldStr(GetValue("銷退服務費"), 65)
        .Range(strCol & intCounter).Formula = "=Sum(" & strCol & strStart & ":" & strCol & intCounter - 1 & ")"
        .Range(strCol & intCounter).NumberFormatLocal = "#,##0"
        strCol = GetFieldStr(GetValue("銷退規費"), 65)
        strColE = strCol
        .Range(strCol & intCounter).Formula = "=Sum(" & strCol & strStart & ":" & strCol & intCounter - 1 & ")"
        .Range(strCol & intCounter).NumberFormatLocal = "#,##0"
        'Mark by Amy 2024/12/24 不畫框-瑞婷
        'Call SetFrame(strColS, strColE, Val(strStart) - 2, intCounter) '框線
    End With
End Sub

'取得案件規費餘額語法(從Select1搬來 )
Private Function GetCaseFeesBalance(ByVal stCP01 As String, ByVal stCP02 As String, ByVal stCP03 As String, ByVal stCP04 As String) As String
   If stCP01 = "TF" Then
        GetCaseFeesBalance = "ax214>='" & stCP01 & stCP02 & "000' AND ax214<='" & stCP01 & stCP02 & "ZZZ' "
    ElseIf stCP01 = "CFP" Then
        GetCaseFeesBalance = "ax214>='" & stCP01 & stCP02 & stCP03 & "00' AND ax214<='" & stCP01 & stCP02 & stCP03 & "99' "
    Else
        GetCaseFeesBalance = "ax214='" & stCP01 & stCP02 & stCP03 & stCP04 & "'"
    End If
    GetCaseFeesBalance = "Select sum(ax207)-sum(ax206) From acc021 Where " & GetCaseFeesBalance & " and SubStr(ax205,1,4) = '2201'"
End Function

Private Sub SetFrame(strColStart, strColEnd, intStartRow, intEndRow)
    '框線
    wksAnnuity.Range(strColStart & intStartRow & ":" & strColEnd & intEndRow).Select
    xlsAnnuity.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    xlsAnnuity.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    xlsAnnuity.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlsAnnuity.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
   
End Sub
'end 2020/04/27

'Add by Amy 2022/07/04 取得案源介紹人
Private Function GetLos04(ByVal stA0j01 As String) As String
    Dim RsQ As New ADODB.Recordset, intQ As Integer
    Dim strQ As String

    GetLos04 = ""
    strQ = "Select Los04 From LawOfficeSource Where Los06='" & stA0j01 & "' "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        GetLos04 = "" & RsQ.Fields("Los04")
    End If
    Set RsQ = Nothing
End Function

'是否有介紹人資料
Private Function ChkLaw() As Boolean
    Dim RsQ As New ADODB.Recordset, intQ As Integer
    Dim strQ As String
    
    ChkLaw = False
    strQ = "Select * From accrpt106 Where r10601 = '" & strUserNum & "' And r10631 Is Not Null"
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        ChkLaw = True
    End If
    Set RsQ = Nothing
End Function
'end 2022/07/04

'Add by Amy 2022/07/20 畫面條件合併處理
Private Function GetWhere(IsCaseSource As Boolean) As String
    Dim strCmp As String 'Add by Amy 2020/04/27
    
    'Modify by Morgan 2007/10/1 「智權人員」範圍改成一個
    'If Text1 <> MsgText(601) Then
    '   strSQL = " and a0k20 >= '" & Text1 & "'"
    'End If
    'If Text2 <> MsgText(601) Then
    '   strSQL = strSQL & " and a0k20 <= '" & Text2 & "'"
    'End If
    'Modify by Amy 2022/07/20 抓「案源」判斷Los04
    If Text1 <> MsgText(601) Then
        If IsCaseSource = False Then
            GetWhere = GetWhere & " and a0k20= '" & Text1 & "'"
        Else
            GetWhere = GetWhere & " and InStr(los04,'" & Text1 & "')>0"
        End If
    End If
    'end 2007/10/1
   
    'Add by Amy 2020/04/27 之前改「公司別」語法未加入公司別條件
    'Modify by Amy 2020/07/08 改回用輸入不用下拉
    strCmp = Trim(Text2) 'Trim(cboComp)
    If strCmp <> MsgText(601) Then
'       If InStr(strCmp, "　") > 0 Then
'             strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
'       End If
        GetWhere = GetWhere & " and a0k11= '" & strCmp & "'"
    End If
    'end 2020/07/08
    'end 2020/04/27
    
    'Add by Sindy 2010/8/17 「收據號碼」
    If Text4 <> MsgText(601) Then
        GetWhere = GetWhere & " and a0k01>= '" & Text4 & "'"
    End If
    If Text5 <> MsgText(601) Then
        GetWhere = GetWhere & " and a0k01<= '" & Text5 & "'"
    End If
    '2010/8/17 End
   
    'Add By Sindy 2016/6/13
    '收據抬頭
    If cboTitle.Text <> MsgText(601) Then
        GetWhere = GetWhere & " And Instr(UPPER(a0k04), UPPER('" & cboTitle.Text & "')) > 0"
    End If
    '2016/6/13 END
   
    'Add By Sindy 2010/9/13
    '客戶代號
    If Text7 <> MsgText(601) Then
        GetWhere = GetWhere & " and a0k03>= '" & Text7 & "'"
    End If
    If Text8 <> MsgText(601) Then
        GetWhere = GetWhere & " and a0k03<= '" & Text8 & "'"
    End If
    
    '收文日
    If MaskEdBox3.Text <> MsgText(601) And MaskEdBox3.Text <> MsgText(29) Then
        GetWhere = GetWhere & " and cp05 >= " & ChangeTStringToWString(Val(FCDate(MaskEdBox3.Text))) & ""
    End If
    If MaskEdBox4.Text <> MsgText(601) And MaskEdBox4.Text <> MsgText(29) Then
        GetWhere = GetWhere & " and cp05 <= " & ChangeTStringToWString(Val(FCDate(MaskEdBox4.Text))) & ""
    End If
        
    '發文日
    If MaskEdBox5.Text <> MsgText(601) And MaskEdBox5.Text <> MsgText(29) Then
        GetWhere = GetWhere & " and cp27 >= " & ChangeTStringToWString(Val(FCDate(MaskEdBox5.Text))) & ""
    End If
    If MaskEdBox6.Text <> MsgText(601) And MaskEdBox6.Text <> MsgText(29) Then
        GetWhere = GetWhere & " and cp27 <= " & ChangeTStringToWString(Val(FCDate(MaskEdBox6.Text))) & ""
    End If
    '2010/9/13 End
   
    '若非北所員工, 只能列印該所資料
    'Modify by Amy 2022/07/20 不是抓「案源」資料才判斷「所別」條件
    If IsCaseSource = False Then
        If pub_strUserOffice <> "1" Then
            GetWhere = GetWhere & " And ''||ST06='" & pub_strUserOffice & "' "
        End If
    End If
   
    'Add By Sindy 2010/5/20
    If Check1.Value = 1 Then
       GetWhere = GetWhere & " And (a0k08 is not null And a0k08<>' ') "
    End If
   
    'Add by Sindy 2010/8/17 僅列印已送件者
    If Check2.Value = 1 Then
        GetWhere = GetWhere & " And (cp27 is not null And cp27>0) "
    End If
    
   'Add By Sindy 2016/8/22
   '本所案號
   If Text9 <> "" And Text10 <> "" Then
      GetWhere = GetWhere & " And a0j02='" & Text9 & Text10 & Text11 & Text12 & "'"
   End If
   '2016/8/22 END
   
    If Text3 = "1" Then
        '帳款日期
        If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
            GetWhere = GetWhere & " And a0k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
        End If
        If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
            GetWhere = GetWhere & " And a0k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
        End If
    Else
        '帳款日期
        If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
            GetWhere = GetWhere & " And a0l02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
        End If
        If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
            GetWhere = GetWhere & " And a0l02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
        End If
    End If
End Function

Private Sub GetWhere3(IsCaseSource As Boolean, ByRef strWhere1 As String, ByRef strWhere2 As String, ByRef strWhere3 As String)
    Dim strCmp As String
    
    '智權人員
    'Modify by Amy 2022/07/20 抓「案源」判斷Los04
    If Text1 <> MsgText(601) Then
        If IsCaseSource = False Then
            strWhere1 = strWhere1 & " And a0k20= '" & Text1 & "'"
            strWhere2 = strWhere2 & " And a0k20= '" & Text1 & "'"
            strWhere3 = strWhere3 & " And a0k20= '" & Text1 & "'"
        Else
            strWhere1 = strWhere1 & " And InStr(los04,'" & Text1 & "')>0"
            strWhere2 = strWhere2 & " And InStr(los04,'" & Text1 & "')>0"
            strWhere3 = strWhere3 & " And InStr(los04,'" & Text1 & "')>0"
        End If
    End If
    
    '公司別
    'Modify by Amy 2020/07/08 改回輸入(不用下拉)
    strCmp = Trim(Text2) 'Trim(cboComp)
    If strCmp <> MsgText(601) Then
'        If InStr(strCmp, "　") > 0 Then
'            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
'        End If
        strWhere1 = strWhere1 & " And a0k11= '" & strCmp & "'"
        strWhere2 = strWhere2 & " And a0k11= '" & strCmp & "'"
        strWhere3 = strWhere3 & " And a0k11= '" & strCmp & "'"
    End If
    
    '帳款日期
    If MaskEdBox1 <> MsgText(601) And MaskEdBox1 <> MsgText(29) Then
        strWhere1 = strWhere1 & " And a0k02 >= " & Val(FCDate(MaskEdBox1)) & ""
        strWhere2 = strWhere2 & " And a0L02 >= " & Val(FCDate(MaskEdBox1)) & ""
        strWhere3 = strWhere3 & " And a0s03 >= " & Val(FCDate(MaskEdBox1)) & ""
    End If
    If MaskEdBox2 <> MsgText(601) And MaskEdBox2 <> MsgText(29) Then
        strWhere1 = strWhere1 & " And a0k02 <= " & Val(FCDate(MaskEdBox2)) & ""
        strWhere2 = strWhere2 & " And a0L02 <= " & Val(FCDate(MaskEdBox2)) & ""
        strWhere3 = strWhere3 & " And a0s03 <= " & Val(FCDate(MaskEdBox2)) & ""
    End If
    
    '收據抬頭
    If cboTitle <> MsgText(601) Then
        strWhere1 = strWhere1 & " And InStr(UPPER(a0k04), UPPER('" & cboTitle & "')) > 0"
        strWhere2 = strWhere2 & " And InStr(UPPER(a0k04), UPPER('" & cboTitle & "')) > 0"
        strWhere3 = strWhere3 & " And InStr(UPPER(a0k04), UPPER('" & cboTitle & "')) > 0"
    End If
    
    '客戶代號
    If Text7 <> MsgText(601) Then
        strWhere1 = strWhere1 & " and a0k03>= '" & Text7 & "'"
        strWhere2 = strWhere2 & " and a0k03>= '" & Text7 & "'"
        strWhere3 = strWhere3 & " and a0k03>= '" & Text7 & "'"
    End If
    If Text8 <> MsgText(601) Then
        strWhere1 = strWhere1 & " and a0k03<= '" & Text8 & "'"
        strWhere2 = strWhere2 & " and a0k03<= '" & Text8 & "'"
        strWhere3 = strWhere3 & " and a0k03<= '" & Text8 & "'"
    End If
    
    '收文日
    If MaskEdBox3 <> MsgText(601) And MaskEdBox3 <> MsgText(29) Then
        strWhere1 = strWhere1 & " and cp05 >= " & ChangeTStringToWString(Val(FCDate(MaskEdBox3))) & ""
        strWhere2 = strWhere2 & " and cp05 >= " & ChangeTStringToWString(Val(FCDate(MaskEdBox3))) & ""
        strWhere3 = strWhere3 & " and cp05 >= " & ChangeTStringToWString(Val(FCDate(MaskEdBox3))) & ""
    End If
    If MaskEdBox4 <> MsgText(601) And MaskEdBox4 <> MsgText(29) Then
        strWhere1 = strWhere1 & " and cp05 <= " & ChangeTStringToWString(Val(FCDate(MaskEdBox4))) & ""
        strWhere2 = strWhere2 & " and cp05 <= " & ChangeTStringToWString(Val(FCDate(MaskEdBox4))) & ""
        strWhere3 = strWhere3 & " and cp05 <= " & ChangeTStringToWString(Val(FCDate(MaskEdBox4))) & ""
    End If
        
    '發文日
    If MaskEdBox5 <> MsgText(601) And MaskEdBox5 <> MsgText(29) Then
        strWhere1 = strWhere1 & " and cp27 >= " & ChangeTStringToWString(Val(FCDate(MaskEdBox5))) & ""
        strWhere2 = strWhere2 & " and cp27 >= " & ChangeTStringToWString(Val(FCDate(MaskEdBox5))) & ""
        strWhere3 = strWhere3 & " and cp27 >= " & ChangeTStringToWString(Val(FCDate(MaskEdBox5))) & ""
    End If
    If MaskEdBox6 <> MsgText(601) And MaskEdBox6 <> MsgText(29) Then
        strWhere1 = strWhere1 & " and cp27 <= " & ChangeTStringToWString(Val(FCDate(MaskEdBox6))) & ""
        strWhere2 = strWhere2 & " and cp27 <= " & ChangeTStringToWString(Val(FCDate(MaskEdBox6))) & ""
        strWhere3 = strWhere3 & " and cp27 <= " & ChangeTStringToWString(Val(FCDate(MaskEdBox6))) & ""
    End If
    
    '若非北所員工, 只能列印該所資料
    'Modify by Amy 2022/07/20 不是抓「案源」資料才判斷發文日條件
    If IsCaseSource = False Then
        If pub_strUserOffice <> "1" Then
            strWhere1 = strWhere1 & " And ''||ST06='" & pub_strUserOffice & "' "
            strWhere2 = strWhere2 & " And ''||ST06='" & pub_strUserOffice & "' "
            strWhere3 = strWhere3 & " And ''||ST06='" & pub_strUserOffice & "' "
        End If
    End If
   
    '只列印有備註
    If Check1.Value = 1 Then
        strWhere1 = strWhere1 & " And (a0k08 is not null And a0k08<>' ') "
        strWhere2 = strWhere2 & " And (a0k08 is not null And a0k08<>' ') "
        strWhere3 = strWhere3 & " And (a0k08 is not null And a0k08<>' ') "
    End If
    
    '只列印已送件
    If Check2.Value = 1 Then
        strWhere1 = strWhere1 & " And (cp27 is not null And cp27>0) "
        strWhere2 = strWhere2 & " And (cp27 is not null And cp27>0) "
        strWhere3 = strWhere3 & " And (cp27 is not null And cp27>0) "
    End If
   
    '本所案號
    If Text9 <> "" And Text10 <> "" Then
        strWhere1 = strWhere1 & " And a0j02='" & Text9 & Text10 & Text11 & Text12 & "'"
        strWhere2 = strWhere2 & " And a0j02='" & Text9 & Text10 & Text11 & Text12 & "'"
        strWhere3 = strWhere3 & " And a0j02='" & Text9 & Text10 & Text11 & Text12 & "'"
    End If
End Sub

'案源資料
Private Function GetCaseSource(ByVal strChoose As String, ByVal strWhere As String) As String
    Dim strQ As String, strTB As String
    
    Select Case strChoose
        Case "1", "3.1" '未收
            strTB = ",Acc0j0"
            strQ = " And a0k01= a0j13(+) And a0j01= los06(+) And los06=cp09(+) "
            If strChoose = "1" Then strQ = strQ & " And a0k37 is null And (a0k06+a0k07) > (Nvl(a0k17, 0)+Nvl(a0k18, 0)) "
        Case "2" '收回
            strTB = ",Acc1u0,Acc0l0"
            strQ = " And a1u02=a0k01(+) And a0l01=a1u01(+) And a1u03= los06(+) And los06=cp09(+) "
        Case "3.2" '往來-收款
            strTB = ",Acc1u0,Acc0l0,Acc0j0"
            strQ = " And SubStr(a1u01,1,1)='F' And a1u01=a0l01(+) And a1u03= los06(+) And a1u02=a0j13(+) And a1u03=a0j01(+) " & _
                       "And los06=cp09(+) And a0j13=a0k01(+) And Nvl(a0k17,0)+Nvl(a0k18,0)>0 "
        Case "3.3" '往來-銷退
            strTB = ",Acc1u0,Acc0s0,Acc0j0"
            strQ = " And SubStr(a1u01,1,1)='I' And a0k01=a0j13(+) And a1u03= los06(+) And a0j13=a1u02(+) And a0j01=a1u03(+) " & _
                       "And los06=cp09(+) And (Nvl(cp77, 0) <> 0 or Nvl(cp78, 0) <> 0) And a1u01=a0s01(+) And a1u02=a0s02(+) And a0k10 is not null "
    End Select
    
    GetCaseSource = "Select * From Acc0k0,LawOfficeSource,CaseProgress" & strTB & _
                               " Where Los01 is not Null And (a0k09 is null Or a0k09 = 0) " & strWhere & strQ
End Function
