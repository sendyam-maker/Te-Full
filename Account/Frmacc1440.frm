VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc1440 
   AutoRedraw      =   -1  'True
   Caption         =   "客戶對帳單"
   ClientHeight    =   3900
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5508
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3900
   ScaleWidth      =   5508
   Begin VB.TextBox txtNote 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   585
      Left            =   30
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "產生資料中，暫時不要使用Excel..."
      Top             =   1560
      Visible         =   0   'False
      Width           =   5460
   End
   Begin VB.CommandButton CmdQuery 
      BackColor       =   &H00C0FFC0&
      Caption         =   "不催款查詢"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2670
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   2550
      Width           =   2265
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   990
      Style           =   2  '單純下拉式
      TabIndex        =   11
      Top             =   3300
      Width           =   3945
   End
   Begin VB.TextBox txtType 
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
      Left            =   1395
      MaxLength       =   1
      TabIndex        =   8
      Top             =   2010
      Width           =   612
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
      Left            =   330
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   2550
      Width           =   2265
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
      Height          =   315
      Left            =   1395
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1620
      Width           =   612
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
      Left            =   1395
      MaxLength       =   5
      TabIndex        =   6
      Top             =   1230
      Width           =   1080
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
      Left            =   3330
      MaxLength       =   9
      TabIndex        =   1
      Top             =   150
      Width           =   1572
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
      Left            =   1395
      MaxLength       =   9
      TabIndex        =   0
      Top             =   150
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1395
      TabIndex        =   2
      Top             =   510
      Width           =   1575
      _ExtentX        =   2794
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
      Left            =   3330
      TabIndex        =   3
      Top             =   510
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
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   300
      Left            =   1395
      TabIndex        =   4
      Top             =   870
      Width           =   1575
      _ExtentX        =   2794
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
      Left            =   3330
      TabIndex        =   5
      Top             =   870
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
   Begin VB.Image tmpImg_L1 
      Height          =   525
      Left            =   480
      Top             =   4740
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image tmpImg_L2 
      Height          =   585
      Left            =   1290
      Top             =   4680
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
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
      Left            =   180
      TabIndex        =   24
      Top             =   3360
      Width           =   885
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "PS：未列印收據不會印出來！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   216
      Left            =   300
      TabIndex        =   23
      Top             =   3036
      Width           =   2976
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
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
      Index           =   1
      Left            =   210
      TabIndex        =   22
      Top             =   870
      Width           =   975
   End
   Begin VB.Label Label8 
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
      Left            =   3090
      TabIndex        =   21
      Top             =   870
      Width           =   255
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "列印別                 (1.單一 2.複合)"
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
      Left            =   510
      TabIndex        =   20
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Label Label5 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "報表類別"
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
      Left            =   180
      TabIndex        =   19
      Top             =   1650
      Width           =   1005
   End
   Begin VB.Label lblSalesName 
      BackStyle       =   0  '透明
      Caption         =   "智權人員名稱"
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
      Left            =   2610
      TabIndex        =   18
      Top             =   1260
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "(1.往來帳 2.應收帳)"
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
      Left            =   2160
      TabIndex        =   17
      Top             =   1650
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
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
      Left            =   -30
      TabIndex        =   16
      Top             =   1260
      Width           =   1215
   End
   Begin VB.Label Label3 
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
      Left            =   3090
      TabIndex        =   15
      Top             =   510
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "收據日期"
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
      Index           =   0
      Left            =   30
      TabIndex        =   14
      Top             =   510
      Width           =   1155
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
      Left            =   3090
      TabIndex        =   13
      Top             =   150
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
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
      Index           =   1
      Left            =   210
      TabIndex        =   12
      Top             =   150
      Width           =   975
   End
   Begin VB.Image tmpImg_2 
      Height          =   555
      Left            =   4200
      Top             =   4080
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Image tmpImg_1 
      Height          =   555
      Left            =   3030
      Top             =   4050
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image tmpImg_J2 
      Height          =   585
      Left            =   2010
      Top             =   4020
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image tmpImg_J1 
      Height          =   525
      Left            =   1200
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "Frmacc1440"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/4/11 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
Option Explicit

Public adoacc0k0 As New ADODB.Recordset
Public adocaseprogress As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Dim strSql As String, strNo As String, strNoB As String
Dim strTitle As String 'Add By Sindy 2010/11/25
Dim strTitleB As String
Dim strCompany As String, strCompanyB As String 'Modify By Sindy 2014/4/14
Dim lngAmount1 As Long
Dim strAmount1 As String
Dim strAmount2 As String
Dim strAmount3 As String
Dim intLength As Integer
Dim intCounter As Integer, intCounterB As Integer
Dim intPage As Integer
Dim lngAmount3 As Long
Dim strSalesMan As String
Private Const intY As Integer = 40
'Add By Cheng 2003/07/21
'本所案號
Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String
Dim strName As String
Dim strPrinter As String 'Add By Sindy 2014/4/14
Dim bolHaveA42Data As Boolean, strA4203 As String 'Add By Sindy 2016/11/22
'Add By Sindy 2022/3/25
Dim m_curCompany As String
Dim xlsAnnuity As New Excel.Application
Dim wksAnnuity As New Worksheet
Const intPageRow As Integer = 27
Const intPageTotRow As Integer = 38
Dim intCounterD As Integer '一頁已放多少列數
'2022/3/25 END


'Add By Sindy 2014/9/4 不催款查詢
Private Sub cmdQuery_Click()
   Frmacc1441.Show
   Me.Hide
End Sub

Private Sub Command1_Click()
   If FormCheck = False Then
      'MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   'MsgBox MsgText(100), , MsgText(5)
   
   m_curCompany = "" 'Add By Sindy 2022/3/25
   Screen.MousePointer = vbHourglass
   PUB_SetOsDefaultPrinter Combo1 'Add by Sindy 2022/3/24 切換Word/Excel印表機
   PUB_RestorePrinter Combo1 'Add By Sindy 2014/4/14
   'Add By Sindy 2014/11/13
   If txtType = "2" Then '2.複合
      MsgBox "複合資料，請確認傳真及地址是否正確。", vbInformation
   End If
   '2014/11/13 END
   PrintDetail_Excel 'Modify By Sindy 2022/4/1
'   PrintDetail
   PUB_SetOsDefaultPrinter strPrinter 'Add by Sindy 2022/3/24
   PUB_RestorePrinter strPrinter 'Add By Sindy 2014/4/14
   Screen.MousePointer = vbDefault
   
   'Add By Sindy 2012/10/26 為往來帳款時才記錄
   If MaskEdBox4 <> MsgText(29) Then
      If Trim(Text4) = "1" Then
         'Modify By Sindy 2013/5/22
'         SaveSetting "TAIE", "ACCOUNT", "DATE1", MaskEdBox3
'         SaveSetting "TAIE", "ACCOUNT", "DATE2", MaskEdBox4
         PUB_SaveLastDate Me.Name, "MaskEdBox3", ChangeTDateStringToTString(MaskEdBox3)
         PUB_SaveLastDate Me.Name, "MaskEdBox4", ChangeTDateStringToTString(MaskEdBox4)
         '2013/5/22 End
      End If
   End If
   '2012/10/26 End
   
   FormClear
   
'   'Add By Sindy 2012/11/12
'   'Modify By Sindy 2013/5/22
'   'MaskEdBox3.Text = GetSetting("TAIE", "ACCOUNT", "DATE2", "")
'   If PUB_GetLastDate(Me.Name, "MaskEdBox4") <> "" Then
'      MaskEdBox3.Text = ChangeTStringToTDateString(PUB_GetLastDate(Me.Name, "MaskEdBox4"))
'   End If
'   '2013/5/22 End
'   '2012/11/12 End
   
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5280
   Me.Height = 4365 'Modify by Amy 2023/10/11 原4245
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'modify by sonia 2014/11/12
   'Text1 = "X"
   'Text2 = "X"
   Text1 = "X00001000"
   Text2 = "X99999ZZZ"
   'end 2014/11/12
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   lblSalesName = ""
   'Add By Sindy 2012/11/12
   'Modify By Sindy 2013/5/22
   'MaskEdBox3.Text = GetSetting("TAIE", "ACCOUNT", "DATE2", "")
   If PUB_GetLastDate(Me.Name, "MaskEdBox4") <> "" Then
      MaskEdBox3.Text = ChangeTStringToTDateString(PUB_GetLastDate(Me.Name, "MaskEdBox4"))
   End If
   '2013/5/22 End
   '2012/11/12 End
   'Add By Sindy 2012/11/12
   MaskEdBox3.Mask = DFormat
   MaskEdBox4.Mask = DFormat
   '2012/11/12 End
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   
   'Add By Sindy 2014/4/14
   PUB_SetPrinter Me.Name, Combo1, strPrinter
   If Dir(App.path & "\J-Letterhead.jpg") <> "" Then
      Kill App.path & "\J-Letterhead.jpg"
   End If
   If Dir(App.path & "\J-LastLetter.jpg") <> "" Then
      Kill App.path & "\J-LastLetter.jpg"
   End If
   If Dir(App.path & "\Taie-Letterhead.jpg") <> "" Then
      Kill App.path & "\Taie-Letterhead.jpg"
   End If
   If Dir(App.path & "\Taie-LastLetter.jpg") <> "" Then
      Kill App.path & "\Taie-LastLetter.jpg"
   End If
   Call PUB_GetSampleFile("J-Letterhead.jpg", "M51-000028-0-00")
   Call PUB_GetSampleFile("J-LastLetter.jpg", "M51-000031-0-00")
   Call PUB_GetSampleFile("Taie-Letterhead.jpg", "M51-000079-0-00") '12
   Call PUB_GetSampleFile("Taie-LastLetter.jpg", "M51-000082-0-00") '11
   Set tmpImg_J1.Picture = LoadPicture()
   Set tmpImg_J2.Picture = LoadPicture()
   Set tmpImg_1.Picture = LoadPicture()
   Set tmpImg_2.Picture = LoadPicture()
   tmpImg_J1 = LoadPicture(Trim(App.path & "\J-Letterhead.jpg"))
   tmpImg_J2 = LoadPicture(Trim(App.path & "\J-LastLetter.jpg"))
   tmpImg_1 = LoadPicture(Trim(App.path & "\Taie-Letterhead.jpg"))
   tmpImg_2 = LoadPicture(Trim(App.path & "\Taie-LastLetter.jpg"))
   '2014/4/14 END
   'Add By Sindy 2020/4/20
   If Dir(App.path & "\L-Letterhead.jpg") <> "" Then
      Kill App.path & "\L-Letterhead.jpg"
   End If
   If Dir(App.path & "\L-LastLetter.jpg") <> "" Then
      Kill App.path & "\L-LastLetter.jpg"
   End If
   Call PUB_GetSampleFile("L-Letterhead.jpg", "M51-000080-0-00")
   Call PUB_GetSampleFile("L-LastLetter.jpg", "M51-000078-0-00")
   Set tmpImg_L1.Picture = LoadPicture()
   Set tmpImg_L2.Picture = LoadPicture()
   tmpImg_L1 = LoadPicture(Trim(App.path & "\L-Letterhead.jpg"))
   tmpImg_L2 = LoadPicture(Trim(App.path & "\L-LastLetter.jpg"))
   '2020/4/20 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2014/4/14
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   '2014/4/14 END
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc1440 = Nothing
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(29) Then
      Exit Sub
   End If
   MaskEdBox2.Mask = ""
   '92.10.16 CANCEL BY SONIA
   'MaskEdBox2.Text = MaskEdBox1.Text
   MaskEdBox2.Mask = DFormat
End Sub

'Add By Sindy 2012/11/12
Private Sub MaskEdBox3_Validate(Cancel As Boolean)
   If MaskEdBox3.Text = MsgText(29) Then
      Exit Sub
   End If
   MaskEdBox4.Mask = ""
   MaskEdBox4.Mask = DFormat
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   'Modify by Amy 2014/11/13 +判斷X00001000 因2014/11/12 Form_Load Text2 預帶 "X99999ZZZ" -瑞婷
   If Text1 = MsgText(601) Or Trim(Text1) = "X00001000" Then
      Exit Sub
   End If
   Select Case Len(Text1)
      Case 6
         Text1 = Text1 & "000"
      Case 8
         Text1 = Text1 & "0"
   End Select
   Text2 = Text1
End Sub

Private Sub Text2_GotFocus()
   If Text1.Text <> "" Then
      'Modify By Sindy 2014/8/11 999=>ZZZ
      'Text2.Text = Left(Text1.Text, 6) & "999"
      Text2.Text = Left(Text1.Text, 6) & "ZZZ"
   End If
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

''Add By Sindy 2016/11/22
'Private Function GetA42Data(ByVal strA4201 As String, ByRef strA4203 As String) As Boolean
'Dim rsA As New ADODB.Recordset
'
'   GetA42Data = False
'   strA4203 = ""
'   strSql = "select *" & _
'            " from acc420" & _
'            " where a4201=" & CNULL(ChgSQL(strA4201))
'   intI = 1
'   Set rsA = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      GetA42Data = True
'      strA4203 = rsA.Fields("A4203")
'   End If
'   Set rsA = Nothing
'End Function

'*************************************************
'  抬頭列印
'  Mark by Amy 2023/05/22 不使用
'*************************************************
Private Sub PrintHead(strCompany As String)
'Dim strLocation As String
'
'   m_curCompany = strCompany
'   'Add By Sindy 2014/4/14
'   If strCompany = "J" Then
'      Printer.PaintPicture tmpImg_J1, 0, 300, 11100, 1800
'      Printer.PaintPicture tmpImg_J2, 0, 14700, 11100, 1300
'   'Add By Sindy 2020/4/20
'   ElseIf strCompany = "L" Then
'      Printer.PaintPicture tmpImg_L1, 0, 150, 11100, 2000
'      Printer.PaintPicture tmpImg_L2, 0, 14700, 11100, 750
'   '2020/4/20 END
'   Else
'      'Printer.PaintPicture tmpImg_1, 0, 300, 11100, 1800
'      Printer.PaintPicture tmpImg_1, 0, 150, 11100, 2000
'      Printer.PaintPicture tmpImg_2, 0, 14700, 11100, 1300
'   End If
'   '2014/4/14 END
'   Printer.CurrentX = 9800
'   Printer.CurrentY = 2500 - intY - 100
'   Printer.Print CFDate(ACDate(ServerDate))
'   Printer.CurrentX = 1200
'   Printer.CurrentY = 2500 - intY - 100
'   If IsNull(adoacc0k0.Fields("a0k03").Value) = False Then
'      Printer.Print adoacc0k0.Fields("a0k03").Value
'   Else
'      Printer.Print ""
'   End If
'
'   'Modify By Sindy 2020/4/20 Mark,頁數取消
''   Printer.CurrentX = 9800
''   Printer.CurrentY = 2970 - intY - 100
''   Printer.Print intPage
'
'   Printer.CurrentX = 1200
'   Printer.CurrentY = 2970 - intY - 100
'   If IsNull(adoacc0k0.Fields("a0k04").Value) = False Then
'      Printer.Print adoacc0k0.Fields("a0k04").Value
'   Else
'      Printer.Print ""
'   End If
'   'Add By Sindy 2016/11/22 取得收據抬頭的地址
'   'Modify By Sindy 2017/11/20
'   Dim m_CU30 As String, m_CU31 As String
'   'bolHaveA42Data = GetA42Data(adoacc0k0.Fields("a0k04").Value, strA4203)
'   'Modify By Sindy + adoacc0k0.Fields("a0k03").Value
'   bolHaveA42Data = GetTitleCustData(adoacc0k0.Fields("a0k04").Value, adoacc0k0.Fields("a0k03").Value, "", "", "", _
'                            "", "", "", "", "", "", "", _
'                            "", "", "", "", "", m_CU30, m_CU31)
'   'If bolHaveA42Data = True And strA4203 <> "" Then
'   If bolHaveA42Data = True And m_CU31 <> "" Then '聯絡地址
'   '2016/11/22 END
'      'If IsNull(adoacc0k0.Fields("cu30").Value) Then
'      If IsNull(m_CU30) Then
'         strLocation = ""
'      Else
'         'strLocation = adoacc0k0.Fields("cu30").Value
'         strLocation = m_CU30 '聯絡地址郵遞區號
'      End If
'      Printer.CurrentX = 1200
'      Printer.CurrentY = 3440 - intY - 100
'      If IsNull(m_CU31) = False Then
'         If LenB(strLocation & m_CU31) <= 30 Then
'            Printer.Print strLocation & m_CU31
'         Else
'            Printer.Print MidB(strLocation & m_CU31, 1, 30)
'            Printer.CurrentX = 1200
'            Printer.CurrentY = 3780 - intY - 100
'            Printer.Print RightB(strLocation & m_CU31, LenB(strLocation & m_CU31) - 30)
'         End If
'      Else
'         Printer.Print strLocation
'      End If
'   End If
'   '2017/11/20 END
'
'   Printer.FontSize = 16
'   Printer.CurrentX = 4200
'   Printer.CurrentY = 4400 ' - intY - 100
''   If Text4 = "1" Then
''      Printer.Print ReportTitle(1041)
''   Else
'      Printer.Print ReportTitle(1042) '***  客戶應收帳款對帳單  ***
''   End If
'   Printer.FontSize = 11
''   Printer.CurrentX = 3000
''   Printer.CurrentY = 2700
''   Printer.Print "帳款日期: "
'   Printer.CurrentX = 5300
'   Printer.CurrentY = 4850 ' - intY - 100
'   If MaskEdBox1.Text <> MsgText(29) Then
'      Printer.Print MaskEdBox1.Text
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 6200
'   Printer.CurrentY = 4850 ' - intY - 100
'   Printer.Print " ~ "
'   Printer.CurrentX = 6400
'   Printer.CurrentY = 4850 ' - intY - 100
'   If MaskEdBox2.Text <> MsgText(29) Then
'      Printer.Print MaskEdBox2.Text
'   Else
'      Printer.Print ""
'   End If
'
'   'Add By Sindy 2014/4/14
'   Printer.CurrentX = 200
'   Printer.CurrentY = 5300
'   Printer.Print "收據日期"
'   Printer.CurrentX = 1250
'   Printer.CurrentY = 5300
'   Printer.Print "收據號碼"
'
'   'Modify By Sindy 2020/4/27 + 將(本所案號)取消  含 商標 /智慧所 或 法律所
''   Printer.CurrentX = 2340
''   Printer.CurrentY = 5300
''   Printer.Print "本所案號"
'
'   Printer.CurrentX = 2350 '3650
'   Printer.CurrentY = 5300
'   Printer.Print "國　　別"
'   Printer.CurrentX = 3800 '4880
'   Printer.CurrentY = 5300
'   Printer.Print "案件性質"
'   Printer.CurrentX = 5200 '6000
'   Printer.CurrentY = 5300
'   Printer.Print "案　件　名　稱"
'   Printer.CurrentX = 9800 - Printer.TextWidth("應收金額")
'   Printer.CurrentY = 5300
'   Printer.Print "應收金額"
'   Printer.CurrentX = 11100 - Printer.TextWidth("已收金額")
'   Printer.CurrentY = 5300
'   Printer.Print "已收金額"
'   '2014/4/14 END
End Sub

'*************************************************
'  抬頭列印
'
'*************************************************
'Add By Sindy 2022/3/25
'strType=1.客戶應收帳款對帳單
'strType=2.客戶收款對帳單
Private Sub PrintHead_Excel(strCompany As String, stra0k03 As String, strA0K04 As String, _
   strType As String)
Dim strLocation As String
Dim strTemp As String
   
   If m_curCompany = "" Then
      '啟動Excel Object
      '預設A4紙張/橫式/比例 80%/水平置中/邊界左右都改0
      Set xlsAnnuity = New Excel.Application
      'xlsAnnuity.Visible = True
      xlsAnnuity.SheetsInNewWorkbook = 1 '預設工作表數量
   End If
   '檢查頁首頁尾是否要換圖檔
   If m_curCompany <> strCompany Then
      If m_curCompany <> "" Then
         Call PrintExcel '列印,Close
      End If
      '開WorkSheet
      xlsAnnuity.Workbooks.add
      xlsAnnuity.Visible = False 'True
      Set wksAnnuity = xlsAnnuity.Worksheets(1)
      xlsAnnuity.ActiveWindow.Zoom = 75 '畫面比例100%太大了,調整為75%
      '把Excel的警告訊息關掉
      xlsAnnuity.DisplayAlerts = False
      
      wksAnnuity.PageSetup.PaperSize = 9 'A4
      'wksAnnuity.PageSetup.Orientation = xlLandscape '橫印
      wksAnnuity.PageSetup.Orientation = wdOrientLandscape '直印
      wksAnnuity.PageSetup.LeftMargin = xlsAnnuity.InchesToPoints(0.373700787401575) '邊界 0.2
      wksAnnuity.PageSetup.RightMargin = xlsAnnuity.InchesToPoints(0.373700787401575) '0.2
      wksAnnuity.PageSetup.TopMargin = xlsAnnuity.InchesToPoints(1.81102362204724) '1.614
      wksAnnuity.PageSetup.BottomMargin = xlsAnnuity.InchesToPoints(0.984251968503937) '0.984
      wksAnnuity.PageSetup.HeaderMargin = xlsAnnuity.InchesToPoints(0.118110236220472) '0.1
      If strCompany = "J" Then
         wksAnnuity.PageSetup.FooterMargin = xlsAnnuity.InchesToPoints(0)
      Else
         wksAnnuity.PageSetup.FooterMargin = xlsAnnuity.InchesToPoints(0.196850393700787) '0.2
      End If
      wksAnnuity.PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
      '插入頁首頁尾圖檔
      wksAnnuity.PageSetup.CenterHeader = "&G" 'Enable the image to show up
      wksAnnuity.PageSetup.CenterFooter = "&G" 'Enable the image to show up
      '智權公司
      If strCompany = "J" Then
         wksAnnuity.PageSetup.CenterHeaderPicture.FileName = App.path & "\J-Letterhead.jpg"
         wksAnnuity.PageSetup.CenterFooterPicture.FileName = App.path & "\J-LastLetter.jpg"
' .Brightness = 0.36
' .ColorType = msoPictureGrayscale
' .Contrast = 0.39
' .CropBottom = -14.4
' .CropLeft = -28.8
' .CropRight = -14.4
' .CropTop = 21.6
      '法律所
      ElseIf strCompany = "L" Then
         wksAnnuity.PageSetup.CenterHeaderPicture.FileName = App.path & "\L-Letterhead.jpg"
         wksAnnuity.PageSetup.CenterFooterPicture.FileName = App.path & "\L-LastLetter.jpg"
      '智慧所
      Else
         wksAnnuity.PageSetup.CenterHeaderPicture.FileName = App.path & "\Taie-Letterhead.jpg"
         wksAnnuity.PageSetup.CenterFooterPicture.FileName = App.path & "\Taie-LastLetter.jpg"
      End If
      wksAnnuity.PageSetup.CenterHeaderPicture.Height = 120 '127.5
      wksAnnuity.PageSetup.CenterHeaderPicture.Width = 650 '680.25
      wksAnnuity.PageSetup.CenterFooterPicture.Height = 52 '54.75
      wksAnnuity.PageSetup.CenterFooterPicture.Width = 650 '680.25
    '   xlsAnnuity.SheetsInNewWorkbook = 1 '預設工作表數量
   '   xlsAnnuity.Workbooks.add
   '   Set wksAnnuity = xlsAnnuity.Worksheets(1)
      wksAnnuity.Activate
      
      '設定各欄位長度
      wksAnnuity.Columns("A").ColumnWidth = 10
      wksAnnuity.Columns("B").ColumnWidth = 10
      wksAnnuity.Columns("C").ColumnWidth = 10
      wksAnnuity.Columns("D").ColumnWidth = 14
      wksAnnuity.Columns("E").ColumnWidth = 25
      wksAnnuity.Columns("F").ColumnWidth = 10
      wksAnnuity.Columns("G").ColumnWidth = 10
      intCounter = 0: intCounterD = 0
   Else
      '換頁
      wksAnnuity.Rows(intCounter + 1 & ":" & intCounter + 1).Select
      wksAnnuity.HPageBreaks.add Before:=wksAnnuity.Application.ActiveCell
      'Rows("38:38").Select
      'ActiveWindow.SelectedSheets.HPageBreaks.add Before:=ActiveCell
      intCounterD = 0
   End If
   m_curCompany = strCompany
   
   '標題
   intCounter = intCounter + 1: intCounterD = intCounterD + 1
   xlsAnnuity.Range("G" & intCounter).Value = CFDate(ACDate(ServerDate))
   If IsNull(stra0k03) = False Then
      xlsAnnuity.Range("A" & intCounter).Value = "　　" & stra0k03
   Else
      xlsAnnuity.Range("A" & intCounter).Value = ""
   End If
   intCounter = intCounter + 1: intCounterD = intCounterD + 1
   If IsNull(strA0K04) = False Then
      xlsAnnuity.Range("A" & intCounter).Value = "　　" & strA0K04
   Else
      xlsAnnuity.Range("A" & intCounter).Value = ""
   End If
   xlsAnnuity.Rows(intCounter - 1 & ":" & intCounter).Select
   xlsAnnuity.Selection.RowHeight = 30 '列高
   'Add By Sindy 2016/11/22 取得收據抬頭的地址
   'Modify By Sindy 2017/11/20
   Dim m_CU30 As String, m_CU31 As String
   'bolHaveA42Data = GetA42Data(strA0k04, strA4203)
   'Modify By Sindy + strA0k03
   'Modify by Amy 2023/05/22 避免多筆資料抓錯,先抓抬頭+客戶編號,抓不到再抓抬頭資料的第一筆(同frmacc11t0)
   '條件:X69365070/收據日期1100301-1120228/收款日1100315/報表類別2-應收/列印別1-單一 ->地址[不]應該出現X6936504的聯絡地址
   bolHaveA42Data = GetTitleCustData(strA0K04, stra0k03, "", "", "", _
                            "", "", "", "", "", "", "", _
                            "", "", "", "", "", m_CU30, m_CU31, , , , False, , , , , , , , , , , , Me.Name)
   If bolHaveA42Data = False Then
        bolHaveA42Data = GetTitleCustData(strA0K04, stra0k03, "", "", "", _
                                 "", "", "", "", "", "", "", _
                                 "", "", "", "", "", m_CU30, m_CU31)
   End If
   'end 2023/05/22
   intCounter = intCounter + 1: intCounterD = intCounterD + 1
   xlsAnnuity.Range("A" & intCounter).NumberFormatLocal = "@" '文字
   'If bolHaveA42Data = True And strA4203 <> "" Then
   If bolHaveA42Data = True And m_CU31 <> "" Then '聯絡地址
   '2016/11/22 END
      'If IsNull(adoacc0k0.Fields("cu30").Value) Then
      If IsNull(m_CU30) Then
         strLocation = ""
      Else
         'strLocation = adoacc0k0.Fields("cu30").Value
         strLocation = m_CU30 '聯絡地址郵遞區號
      End If
      If IsNull(m_CU31) = False Then
         If LenB(strLocation & m_CU31) <= 30 Then
            xlsAnnuity.Range("A" & intCounter).Value = "　　" & strLocation & m_CU31
         Else
            strExc(10) = MidB(strLocation & m_CU31, 1, 30)
            strExc(10) = strExc(10) & vbCrLf & "　　" & RightB(strLocation & m_CU31, LenB(strLocation & m_CU31) - 30)
            xlsAnnuity.Range("A" & intCounter).Value = "　　" & strExc(10)
         End If
      Else
         xlsAnnuity.Range("A" & intCounter).Value = "　　" & strLocation
      End If
   End If
   '2017/11/20 END
   xlsAnnuity.Range("A" & intCounter & ":" & "D" & intCounter).Select
   With xlsAnnuity.Selection
      .WrapText = True '換行
      .MergeCells = True
   End With
   xlsAnnuity.Rows(intCounter & ":" & intCounter).Select
   xlsAnnuity.Selection.RowHeight = 34 '列高
   
   intCounter = intCounter + 2: intCounterD = intCounterD + 2
   If strType = "1" Then
      xlsAnnuity.Range("A" & intCounter).Value = ReportTitle(1042) '***  客戶應收帳款對帳單  ***
   Else
      xlsAnnuity.Range("A" & intCounter).Value = "***  客戶收款對帳單  ***"
   End If
   xlsAnnuity.Range("A" & intCounter & ":" & "G" & intCounter).Select
   With xlsAnnuity.Selection
      .HorizontalAlignment = xlCenter '置中
      .VerticalAlignment = xlCenter
      .WrapText = False
      .Orientation = 0
      .AddIndent = False
      .IndentLevel = 0
      .ShrinkToFit = False
      .ReadingOrder = xlContext
      .MergeCells = True
   End With
   With xlsAnnuity.Selection.Font
      '.Bold = True '粗體
      .Name = "新細明體"
      .Size = 16
   End With
   
   intCounter = intCounter + 1: intCounterD = intCounterD + 1
   strTemp = ""
   If strType = "1" Then
      '客戶應收帳款對帳單
      If MaskEdBox1.Text <> MsgText(29) Then
         strTemp = MaskEdBox1.Text
      End If
      strTemp = strTemp & " ~ "
      If MaskEdBox2.Text <> MsgText(29) Then
         strTemp = strTemp & MaskEdBox2.Text
      End If
   Else
      '客戶收款對帳單
      If MaskEdBox3.Text <> MsgText(29) Then
         strTemp = MaskEdBox3.Text
      End If
      strTemp = strTemp & " ~ "
      If MaskEdBox4.Text <> MsgText(29) Then
         strTemp = strTemp & MaskEdBox4.Text
      End If
   End If
   xlsAnnuity.Range("A" & intCounter).Value = strTemp
   xlsAnnuity.Range("A" & intCounter & ":" & "G" & intCounter).Select
   With xlsAnnuity.Selection
      .HorizontalAlignment = xlCenter '置中
      .VerticalAlignment = xlCenter
      .WrapText = False
      .Orientation = 0
      .AddIndent = False
      .IndentLevel = 0
      .ShrinkToFit = False
      .ReadingOrder = xlContext
      .MergeCells = True
   End With
   With xlsAnnuity.Selection.Font
      '.Bold = True '粗體
      .Name = "新細明體"
      .Size = 11
   End With
   
   intCounter = intCounter + 1: intCounterD = intCounterD + 1
   xlsAnnuity.Range("A" & intCounter).Value = "收據日期"
   xlsAnnuity.Range("B" & intCounter).Value = "收據號碼"
   xlsAnnuity.Range("C" & intCounter).Value = "國　　別"
   xlsAnnuity.Range("D" & intCounter).Value = "案件性質"
   xlsAnnuity.Range("E" & intCounter).Value = "案件名稱"
   xlsAnnuity.Range("F" & intCounter).Value = "應收金額"
   xlsAnnuity.Range("G" & intCounter).Value = "已收金額"
End Sub

'Add By Sindy 2012/11/16 客戶收款對帳單
Private Sub PrintHead2(strCompany As String)
Dim strLocation As String
   
   m_curCompany = strCompany
   'Add By Sindy 2014/4/14
   If strCompany = "J" Then
      Printer.PaintPicture tmpImg_J1, 0, 300, 11100, 1800
      Printer.PaintPicture tmpImg_J2, 0, 14700, 11100, 1300
   'Add By Sindy 2020/4/20
   ElseIf strCompany = "L" Then
      Printer.PaintPicture tmpImg_L1, 0, 300, 11100, 1800
      Printer.PaintPicture tmpImg_L2, 0, 14700, 11100, 1300
   '2020/4/20 END
   Else
      Printer.PaintPicture tmpImg_1, 0, 300, 11100, 1800
      Printer.PaintPicture tmpImg_2, 0, 14700, 11100, 1300
   End If
   '2014/4/14 END
   Printer.CurrentX = 9800 '11000
   Printer.CurrentY = 2500 - intY - 100
   Printer.Print CFDate(ACDate(ServerDate))
   Printer.CurrentX = 1200
   Printer.CurrentY = 2500 - intY - 100
   If IsNull(adocaseprogress.Fields("a0k03").Value) = False Then
      Printer.Print adocaseprogress.Fields("a0k03").Value
   Else
      Printer.Print ""
   End If
   Printer.CurrentX = 9800 '11000
   Printer.CurrentY = 2970 - intY - 100
   Printer.Print intPage
   Printer.CurrentX = 1200
   Printer.CurrentY = 2970 - intY - 100
   If IsNull(adocaseprogress.Fields("a0k04").Value) = False Then
      Printer.Print adocaseprogress.Fields("a0k04").Value
   Else
      Printer.Print ""
   End If
   
   'Add By Sindy 2016/11/22 取得收據抬頭的地址
   'Modify By Sindy 2017/11/20
   Dim m_CU30 As String, m_CU31 As String
   'bolHaveA42Data = GetA42Data(adoacc0k0.Fields("a0k04").Value, strA4203)
   'Modify By Sindy + adocaseprogress.Fields("a0k03").Value
   'Modify by Amy 2023/05/22 避免多筆資料抓錯,先抓抬頭+客戶編號,抓不到再抓抬頭資料的第一筆
   '條件:X69365070/收據日期1100301-1120228/收款日1100315/報表類別2-應收/列印別1-單一 ->地址[不]應該出現X6936504的聯絡地址
   bolHaveA42Data = GetTitleCustData(adocaseprogress.Fields("a0k04").Value, adocaseprogress.Fields("a0k03").Value, "", "", "", _
                            "", "", "", "", "", "", "", _
                            "", "", "", "", "", m_CU30, m_CU31, , , , False, , , , , , , , , , , , Me.Name)
   If bolHaveA42Data = False Then
        bolHaveA42Data = GetTitleCustData(adocaseprogress.Fields("a0k04").Value, adocaseprogress.Fields("a0k03").Value, "", "", "", _
                                 "", "", "", "", "", "", "", _
                                 "", "", "", "", "", m_CU30, m_CU31)
   End If
   'end 2023/05/22
   'If bolHaveA42Data = True And strA4203 <> "" Then
   If bolHaveA42Data = True And m_CU31 <> "" Then '聯絡地址
      'If IsNull(adocaseprogress.Fields("cu30").Value) Then
      If IsNull(m_CU30) Then
         strLocation = ""
      Else
         'strLocation = adocaseprogress.Fields("cu30").Value
         strLocation = m_CU30 '聯絡地址郵遞區號
      End If
      Printer.CurrentX = 1200
      Printer.CurrentY = 3440 - intY - 100
      If IsNull(m_CU31) = False Then
         If LenB(strLocation & m_CU31) <= 30 Then
            Printer.Print strLocation & m_CU31
         Else
            Printer.Print MidB(strLocation & m_CU31, 1, 30)
            Printer.CurrentX = 1200
            Printer.CurrentY = 3780 - intY - 100
            Printer.Print RightB(strLocation & m_CU31, LenB(strLocation & m_CU31) - 30)
         End If
      Else
         Printer.Print strLocation
      End If
   End If
   '2017/11/20 END
   
   Printer.FontSize = 16
   Printer.CurrentX = 4200
   Printer.CurrentY = 4400 ' - intY - 100
   Printer.Print "***  客戶收款對帳單  ***"
   Printer.FontSize = 11
'   Printer.CurrentX = 3000
'   Printer.CurrentY = 2700
'   Printer.Print "帳款日期: "
   Printer.CurrentX = 5300
   Printer.CurrentY = 4850 ' - intY - 100
   If MaskEdBox3.Text <> MsgText(29) Then
      Printer.Print MaskEdBox3.Text
   Else
      Printer.Print ""
   End If
   Printer.CurrentX = 6200
   Printer.CurrentY = 4850 ' - intY - 100
   Printer.Print " ~ "
   Printer.CurrentX = 6400
   Printer.CurrentY = 4850 ' - intY - 100
   If MaskEdBox4.Text <> MsgText(29) Then
      Printer.Print MaskEdBox4.Text
   Else
      Printer.Print ""
   End If
   
   'Add By Sindy 2014/4/14
   Printer.CurrentX = 200
   Printer.CurrentY = 5300
   Printer.Print "收據日期"
   Printer.CurrentX = 1250
   Printer.CurrentY = 5300
   Printer.Print "收據號碼"
   
   'Modify By Sindy 2020/4/27 + 將(本所案號)取消  含 商標 /智慧所 或 法律所
'   Printer.CurrentX = 2340
'   Printer.CurrentY = 5300
'   Printer.Print "本所案號"

   Printer.CurrentX = 2350 '3650
   Printer.CurrentY = 5300
   Printer.Print "國　　別"
   Printer.CurrentX = 3800 '4880
   Printer.CurrentY = 5300
   Printer.Print "案件性質"
   Printer.CurrentX = 5200 '6000
   Printer.CurrentY = 5300
   Printer.Print "案　件　名　稱"
   Printer.CurrentX = 9800 - Printer.TextWidth("應收金額")
   Printer.CurrentY = 5300
   Printer.Print "應收金額"
   Printer.CurrentX = 11100 - Printer.TextWidth("已收金額")
   Printer.CurrentY = 5300
   Printer.Print "已收金額"
   '2014/4/14 END
End Sub

'*************************************************
' 合計位置
' Mark by Amy 2023/05/22 不使用
'*************************************************
'Modify By Sindy 2012/11/16 +bolShow
Private Sub PrintSum(bolShow As Boolean, strCompany As String)
'Dim lngRow As Long, intLine As Integer 'Add By Sindy 2014/4/14
'
'   strAmount1 = Format(lngAmount1, DDollar)
'   If strAmount1 = "" Then
'      strAmount1 = "0"
'   End If
'   intLength = Printer.TextWidth(strAmount1)
'   'Add By Sindy 2014/4/14
'   Printer.CurrentX = 6000
'   Printer.CurrentY = 11500 ' - 100
'   Printer.Print "總　　　　　　　　　計："
'   '2014/4/14 END
'   Printer.CurrentX = 9800 - intLength
'   Printer.CurrentY = 11500 ' - 100
'   If bolShow = True Then
'      Printer.Print 0
'   Else
'      Printer.Print strAmount1
'   End If
'   strAmount3 = Format(lngAmount3, DDollar)
'   If strAmount3 = "" Then
'      strAmount3 = "0"
'   End If
'   intLength = Printer.TextWidth(strAmount3)
'   Printer.CurrentX = 11100 - intLength
'   Printer.CurrentY = 11500 ' - 100
'   Printer.Print strAmount3
'   strAmount3 = Format(lngAmount1 - lngAmount3, DDollar)
'   If strAmount3 = "" Then
'      strAmount3 = "0"
'   End If
'   intLength = Printer.TextWidth(Trim(strAmount3))
'   'Add By Sindy 2014/4/14
'   Printer.CurrentX = 6000
'   Printer.CurrentY = 11800
'   Printer.Print "應　收　帳　款　餘　額："
'   '2014/4/14 END
'   Printer.CurrentX = 9800 - intLength
'   Printer.CurrentY = 11800
'   Printer.Print strAmount3
'
'   'Modify By Sindy 2014/4/14
'   lngRow = 12000
'   intLine = 1
'   Printer.FontSize = 12
'   Printer.CurrentX = 200
'   Printer.CurrentY = lngRow + (intLine * 300)
'   Printer.Print "衷心感謝您的愛護與支持，本所為加強服務謹將您的帳款往來資料以不定期郵寄方式與您核對，倘本資料"
'   intLine = intLine + 1
'   Printer.CurrentX = 200
'   Printer.CurrentY = lngRow + (intLine * 300)
'   Printer.Print "內容與您記載不符時，為確保您的權益，請儘速與本所財務處聯繫，謝謝合作。"
'   intLine = intLine + 2
'   Printer.CurrentX = 500
'   Printer.CurrentY = lngRow + (intLine * 300)
'   Printer.Print "　　　　　　　　順　頌"
'   intLine = intLine + 1
'   Printer.CurrentX = 500
'   Printer.CurrentY = lngRow + (intLine * 300)
'   Printer.Print "　　　　　商　祺"
'
'   'Modify By Sindy 2020/4/20 法律所,不印智權人員
'   If strCompany <> "L" Then
'   '2020/4/20 END
'      intLine = intLine + 1
'      Printer.CurrentX = 8000
'      Printer.CurrentY = lngRow + (intLine * 300)
'      If adoquery.State <> adStateClosed Then adoquery.Close
'      adoquery.CursorLocation = adUseClient
'   '   adoquery.Open "select * from staff where st01 = '" & strSalesMan & "'", adoTaie, adOpenStatic, adLockReadOnly
'      adoquery.Open "select * from staff where st01 = '" & PUB_GetAKindSalesNo(m_CP01, m_CP02, m_CP03, m_CP04) & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adoquery.RecordCount <> 0 Then
'         If IsNull(adoquery.Fields("st02").Value) Then
'            Printer.Print "智權人員："
'         Else
'            Printer.Print "智權人員：" & adoquery.Fields("st02").Value
'         End If
'      Else
'         Printer.Print "智權人員："
'      End If
'      adoquery.Close
'   End If
'
'   intLine = intLine + 1
'   Printer.CurrentX = 7500
'   Printer.CurrentY = lngRow + (intLine * 300)
'   'Modify By Sindy 2020/3/30
'   If strCompany = "J" Or strCompany = "L" Then
'      Printer.Print A0802Query(strCompany) & "　敬上"
'      'Printer.Print "台一智權股份有限公司　敬上"
'   Else
'      Printer.Print A0802Query("2") & "　敬上"
'      'Printer.Print "台一國際專利商標事務所　敬上"
'   End If
'   '2020/3/30 END
''2013/10/18 CANCEL BY SONIA 瑞婷說取消
''   Printer.FontSize = 11    '2013/5/21 原為16,改為11
''   Printer.CurrentX = 0
''   Printer.CurrentY = 14250
''   Printer.Print MsgText(99)
''   '2013/5/21 add by sonia
''   Printer.CurrentX = 0
''   Printer.CurrentY = 14550
''   Printer.Print "* 為能更有效率的核校資料，擬以電子郵件方式寄發對帳單，"
''   Printer.CurrentX = 0
''   Printer.CurrentY = 14850
''   Printer.Print "  請提供您的公司名稱及台一的客戶代號，e-MAIL至71006@taie.com.tw，謝謝合作！"
''   '2013/5/21 end
''   Printer.FontSize = 11
''2013/10/18 END
'
'   'Add By Sindy 2012/11/16
'   lngAmount1 = 0
'   lngAmount3 = 0
'   '2012/11/16 End
End Sub

'*************************************************
' 合計位置
'
'*************************************************
'Modify By Sindy 2012/11/16 +bolShow
Private Sub PrintSum_Excel(bolShow As Boolean, strCompany As String)
   
   If intCounterD <= intPageRow Then
      If intCounterD Mod intPageRow = 0 Then
         intCounter = intCounter + 1
      Else
         intCounter = intCounter + 1: intCounterD = intCounterD + 1
         Do While intCounterD Mod intPageRow <> 0
            intCounter = intCounter + 1: intCounterD = intCounterD + 1
         Loop
         intCounter = intCounter + 1
      End If
   Else
      '換頁
      wksAnnuity.Rows(intCounter + 1 & ":" & intCounter + 1).Select
      wksAnnuity.HPageBreaks.add Before:=wksAnnuity.Application.ActiveCell
      intCounter = intCounter + intPageRow + 10 '13
   End If
   strAmount1 = Format(lngAmount1, DDollar)
   If strAmount1 = "" Then
      strAmount1 = "0"
   End If
   xlsAnnuity.Range("E" & intCounter).Value = "總　　　　計："
   If bolShow = True Then
      xlsAnnuity.Range("F" & intCounter).Value = 0
   Else
      xlsAnnuity.Range("F" & intCounter).Value = strAmount1
   End If
   strAmount3 = Format(lngAmount3, DDollar)
   If strAmount3 = "" Then
      strAmount3 = "0"
   End If
   xlsAnnuity.Range("G" & intCounter).Value = strAmount3
   
   strAmount3 = Format(lngAmount1 - lngAmount3, DDollar)
   If strAmount3 = "" Then
      strAmount3 = "0"
   End If
   
   intCounter = intCounter + 1
   xlsAnnuity.Range("E" & intCounter).Value = "應收帳款餘額："
   xlsAnnuity.Range("F" & intCounter).Value = strAmount3
   
   intCounter = intCounter + 2
   xlsAnnuity.Range("A" & intCounter).Value = "衷心感謝您的愛護與支持，本所為加強服務謹將您的帳款往來資料以不定期郵寄方式與您核對，"
   intCounter = intCounter + 1
   xlsAnnuity.Range("A" & intCounter).Value = "倘本資料內容與您記載不符時，為確保您的權益，請儘速與本所財務處聯繫，謝謝合作。"
   
   intCounter = intCounter + 2
   xlsAnnuity.Range("A" & intCounter).Value = "　　　　　　　　順　頌"
   intCounter = intCounter + 1
   xlsAnnuity.Range("A" & intCounter).Value = "　　　　　商　祺"
   
   'Modify By Sindy 2020/4/20 法律所,不印智權人員
   intCounter = intCounter + 1
   If strCompany <> "L" Then
   '2020/4/20 END
      If adoquery.State <> adStateClosed Then adoquery.Close
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select * from staff where st01 = '" & PUB_GetAKindSalesNo(m_CP01, m_CP02, m_CP03, m_CP04) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         If IsNull(adoquery.Fields("st02").Value) Then
            xlsAnnuity.Range("E" & intCounter).Value = "　　智權人員："
         Else
            xlsAnnuity.Range("E" & intCounter).Value = "　　智權人員：" & adoquery.Fields("st02").Value
         End If
      Else
         xlsAnnuity.Range("E" & intCounter).Value = "　　智權人員："
      End If
      adoquery.Close
   End If
   intCounter = intCounter + 1
   If strCompany = "J" Or strCompany = "L" Then
      xlsAnnuity.Range("E" & intCounter).Value = "　　" & A0802Query(strCompany) & "　敬上"
   Else
      xlsAnnuity.Range("E" & intCounter).Value = "　　" & A0802Query("2") & "　敬上"
   End If
   
   lngAmount1 = 0
   lngAmount3 = 0
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1 = "X"
   Text2 = "X"
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   'Add By Sindy 2012/11/12
   MaskEdBox3.Mask = ""
   MaskEdBox3.Text = ""
   MaskEdBox3.Mask = DFormat
   MaskEdBox4.Mask = ""
   MaskEdBox4.Text = ""
   MaskEdBox4.Mask = DFormat
   '2012/11/12 End
   Text3 = ""
   Text4 = ""
   lblSalesName = ""
   Text1.SetFocus
   txtType = "" 'Add By Sindy 2010/12/1
End Sub

'*************************************************
' 列印明細
' Mark by Amy 2023/05/22 不使用
'*************************************************
Public Sub PrintDetail()
'Dim douIamount As Double
'Dim douCAmount As Double
'Dim douPAmount As Double
'Dim lngAmount As Double
'Dim stVTB As String, stVTB2 As String, stVTB3 As String 'Add By Sindy 2010/12/1
'Dim bolPrintItem As Boolean, bolPrintItemAmt As Boolean, stLstItem As String, stLstRecNo As String 'Added by Morgan 2011/12/19
'Dim dblItemAmt1 As Double, dblItemAmt3 As Double 'Added by Morgan 2011/12/19
''Add By Sindy 2012/10/31
'Dim strSQL2 As String, strIsCustGoTo As String, strA0j13A0j01 As String, strA1u02A1u03 As String
'Dim dblAmt1 As Double, dblAmt2 As Double, dblAmtI As Double, dblAmtC As Double, dblAmtP As Double
'Dim strIsCustChked As String, dblCount As Double
''2012/10/31 End
'Dim bolHaveData As Boolean 'Add By Sindy 2013/11/28
'
'On Error GoTo ErrHnd
'
'   m_curCompany = ""
'   intCounter = 0
'   intCounterB = 0
'   lngAmount = 0
'   intLength = 0
'   intPage = 0
'   lngAmount1 = 0
'   lngAmount3 = 0
'   strSql = ""
'   strNo = "": strTitle = "" 'Modify By Sindy 2010/11/25
'   strNoB = "": strTitleB = "" 'Modify By Sindy 2013/11/28
'   strCompany = "": strCompanyB = "" 'Modify By Sindy 2014/4/14
'   m_CP01 = "": m_CP02 = "": m_CP03 = "": m_CP04 = "'"
'   If Text1 <> MsgText(601) Then
'      strSql = strSql & " and a0k03 >= '" & Text1 & "'"
'   End If
'   If Text2 <> MsgText(601) Then
'      strSql = strSql & " and a0k03 <= '" & Text2 & "'"
'   End If
'   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'      strSql = strSql & " and a0k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'   End If
'   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'      strSql = strSql & " and a0k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'   End If
'   If Text3 <> MsgText(601) Then
'      strSql = strSql & " and a0k20 = '" & Text3 & "'"
'   End If
'
'   'Modify By Sindy 2012/11/13 Mark 因均要顯示應收帳資料,再依該客戶及收款日期去讀取收款資料
'   'Modify By Sindy 2013/11/26 還是要判斷報表類別 X37597
'   If Text4 = "2" Then
'      strSql = strSql & " and (a0k06+a0k07-decode(a0s05, null, 0, a0s05)) > (nvl(a0k17, 0)+nvl(a0k18, 0))"
'   End If
'
'   '若非北所員工, 只能列印該所資料
'   If pub_strUserOffice <> "1" Then
'       strSql = strSql & " And ''||ST06='" & pub_strUserOffice & "' "
'   End If
'
'   'Add By Sindy 2010/11/5 未列印收據的資料,不要發
'   '2011/6/7 cancel by sonia 改在下面檢查a0k32及是否發文
'   'strSql = strSql & " And a0k19>0 "
'   '2010/11/5 End
'
'   If adoacc0k0.State <> adStateClosed Then adoacc0k0.Close
'   adoacc0k0.CursorLocation = adUseClient
'   'Modify By Sindy 2010/12/1
'   If txtType = "1" Then '1.單一
'      stVTB2 = " HAVING COUNT(DISTINCT a0k04)<=1"
'   Else '2.複合
'      stVTB2 = " HAVING COUNT(DISTINCT a0k04)>1"
'   End If
'
'   'Modify By Sindy 2012/11/13 Mark 因均要顯示應收帳資料,再依該客戶及收款日期去讀取收款資料
'   'Modify By Sindy 2013/11/26 還是要判斷報表類別 X37597
'   If Text4 = "2" Then '2.應收帳
'      stVTB3 = " and nvl(cp79, 0) > 0"
'   Else '1.往來帳
'      stVTB3 = " and (nvl(cp79, 0) > 0 or nvl(cp75,0)>nvl(cp78,0))"
'   End If
'
'   '2011/6/7 add by sonia 未列印收據者不印A0K32,但已發文者仍要印
'   'stVTB3 = stVTB3 & " and (a0k32 is null or (a0k32 is not null and cp27 is not null)) "
'   'Modify By Sindy 2012/11/9辜來電通知,恢復 控制a0k32 is null才印
'   stVTB3 = stVTB3 & " and (a0k32 is null) "
'
'   '2011/6/7 end
'   'Modified by Morgan 2011/11/1 考慮拆收據情形
'   'stVTB = "SELECT substr(A0K03,1,6) from acc0k0, customer, nation, (select a0s02, sum(nvl(a0s05,0)+nvl(a0s06, 0)+nvl(a0s07, 0)) as a0s05 from acc0s0 group by a0s02) acc0s0, Staff, caseprogress where cp60 = a0k01  and substr(a0k03, 1, 8) = cu01 (+) and substr(a0k03, 9, 1) = cu02 (+) and a0k23 = na01 (+) and a0k01 = a0s02 (+) and (a0k09 is null or a0k09 = 0) and (a0k06+a0k07) > decode(a0s05, null, 0, a0s05) And A0K20=ST01(+) " & strSql & stVTB3 & " GROUP BY substr(A0K03,1,6)" & stVTB2
'   'Modify By Sindy 2014/4/11 + and (cu140<>'N' or cu140 is null) 不催款者,不產生對帳單
'   'Modify by Amy 2020/09/18 原:and (cu140<>'N' or cu140 is null),因不寄催款單改輸1-3
'   stVTB = "SELECT substr(A0K03,1,6) from acc0k0,acc0j0,customer" & _
'           ",(select a0s02, sum(nvl(a0s05,0)+nvl(a0s06, 0)+nvl(a0s07, 0)) as a0s05 from acc0s0 group by a0s02) acc0s0" & _
'           ",Staff,caseprogress" & _
'           " where a0j13(+)=a0k01 and cp09(+)=a0j01" & _
'           " and substr(a0k03, 1, 8) = cu01(+) and substr(a0k03, 9, 1) = cu02(+)" & _
'           " and cu140 is null and a0k01 = a0s02(+) and (a0k09 is null or a0k09 = 0)" & _
'           " and (a0k06+a0k07) > decode(a0s05, null, 0, a0s05)" & _
'           " and A0K20=ST01(+) " & strSql & stVTB3 & " GROUP BY substr(A0K03,1,6)" & stVTB2
'   '2011/6/7 modify by sonia 未列印收據者不印A0K32,但已發文者仍要印,且為免同一收據二個收文號資料會重覆故加CP09
'   'strSql = "select A0K01,A0K02,A0K03,A0K04,CU30,CU31 from acc0k0, customer, nation, (select a0s02, sum(nvl(a0s05,0)+nvl(a0s06, 0)+nvl(a0s07, 0)) as a0s05 from acc0s0 group by a0s02) acc0s0, Staff where substr(a0k03, 1, 8) = cu01 (+) and substr(a0k03, 9, 1) = cu02 (+) and a0k23 = na01 (+) and a0k01 = a0s02 (+) and (a0k09 is null or a0k09 = 0) and (a0k06+a0k07) > decode(a0s05, null, 0, a0s05) And A0K20=ST01(+) " & strSql & " and substr(a0k03,1,6) in (" & stVTB & ") order by a0k03 asc, A0K04 asc, a0k01 asc"
'   'Modified by Morgan 2011/11/1 考慮拆收據情形
'   'strSql = "select A0K01,A0K02,A0K03,A0K04,CU30,CU31,CP09 from acc0k0, customer, nation, caseprogress, (select a0s02, sum(nvl(a0s05,0)+nvl(a0s06, 0)+nvl(a0s07, 0)) as a0s05 from acc0s0 group by a0s02) acc0s0, Staff where a0k01=cp60(+) and substr(a0k03, 1, 8) = cu01 (+) and substr(a0k03, 9, 1) = cu02 (+) and a0k23 = na01 (+) and a0k01 = a0s02 (+) and (a0k09 is null or a0k09 = 0) and (a0k06+a0k07) > decode(a0s05, null, 0, a0s05) And A0K20=ST01(+) " & strSql & stVTB3 & " and substr(a0k03,1,6) in (" & stVTB & ") order by a0k03 asc, A0K04 asc, a0k01 asc"
'   'Modify By Sindy 2014/4/11 + and (cu140<>'N' or cu140 is null) 不催款者,不產生對帳單
'   '                          + ,decode(a0k11,'J','J','1') as A0K11
'   '                          order by + ,a0k11 asc
'   'Modify By Sindy 2020/4/20 decode(a0k11,'J','J','1') => decode(a0k11,'J','J','L','L','1')
'   'Modify by Amy 2020/09/18 原:and (cu140<>'N' or cu140 is null),因不寄催款單改輸1-3
'   strSQL2 = "select A0K01,A0K02,A0K03,A0K04,CU30,CU31,CP09,A0K33,A0J22,decode(a0k11,'J','J','L','L','1') as A0K11 from acc0k0,acc0j0, customer, caseprogress, (select a0s02, sum(nvl(a0s05,0)+nvl(a0s06, 0)+nvl(a0s07, 0)) as a0s05 from acc0s0 group by a0s02) acc0s0, Staff where a0j13(+)=a0k01 and cp09(+)=a0j01 and substr(a0k03, 1, 8) = cu01(+) and substr(a0k03, 9, 1) = cu02(+) and cu140 is null and a0k01 = a0s02 (+) and (a0k09 is null or a0k09 = 0) and (a0k06+a0k07) > decode(a0s05, null, 0, a0s05) And A0K20=ST01(+) " & strSql & stVTB3 & " and substr(a0k03,1,6) in (" & stVTB & ") order by a0k03 asc,A0K04 asc,a0k11 asc,a0k01 asc,a0j25 asc"
'   '2010/12/1 End
'   adoacc0k0.Open strSQL2, adoTaie, adOpenStatic, adLockReadOnly
'   If adoacc0k0.RecordCount = 0 Then
'      adoacc0k0.Close
'      MsgBox MsgText(28), , MsgText(5)
'      Exit Sub
'   End If
'
'   Printer.FontName = "新細明體"
'   Printer.FontSize = 11
'   'Modify By Sindy 2014/4/11
'   'Printer.PaperSize = PUB_GetPaperSize(5)
'   Printer.PaperSize = 9
'   '2014/4/11 END
'   'Added by Morgan 2011/12/19 若收據有變更帳款類別則相同的依照列印順序合併
'   bolPrintItem = True
'   dblItemAmt1 = 0
'   dblItemAmt3 = 0
'   'end 2011/12/19
'
'   Do While adoacc0k0.EOF = False
'      'Modify by Morgan 2008/6/2
'      'adocaseprogress.Open "select * from caseprogress, acc0j0, casepropertymap where cp09 = a0j01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp60 = '" & adoacc0k0.Fields("a0k01").Value & "' and nvl(cp16, 0) > nvl(cp77, 0)", adoTaie, adOpenStatic, adLockReadOnly
'      'Add By Sindy 2012/11/13 若選往來時,再依該客戶及收款日期去讀取收款資料
'      If Text4 = "1" Then
'         If (strNoB <> "" And strNoB <> adoacc0k0.Fields("a0k03").Value) Or _
'            (strTitleB <> "" And strTitleB <> adoacc0k0.Fields("a0k04").Value) Or _
'            (strCompanyB <> "" And strCompanyB <> adoacc0k0.Fields("a0k11").Value) Then 'Modify By Sindy 2014/4/14 +strCompanyB
'            If lngAmount1 <> 0 Then
'               Call PrintSum(False, strCompanyB)
'            End If
'            Call PrintType1Data(strNoB, strTitleB, strCompanyB)
'         End If
'      End If
'      '2012/11/13 End
'
'      'Modify By Sindy 2012/11/13 Mark 因均要顯示應收帳資料,再依該客戶及收款日期去讀取收款資料
'      '2011/6/7 modify by sonia 加cp09條件
'      'adocaseprogress.Open "select * from caseprogress, acc0j0, casepropertymap where cp09 = a0j01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp60 = '" & adoacc0k0.Fields("a0k01").Value & "' and nvl(cp79, 0) > 0", adoTaie, adOpenStatic, adLockReadOnly
'      'Modified by Morgan 2011/11/1 考慮拆收據情形
'      'adocaseprogress.Open "select * from caseprogress, acc0j0, casepropertymap where cp09 = a0j01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp60 = '" & adoacc0k0.Fields("a0k01").Value & "' and cp09 = '" & adoacc0k0.Fields("CP09").Value & "' and nvl(cp79, 0) > 0", adoTaie, adOpenStatic, adLockReadOnly
'      'Modified by Morgan 2011/12/30 取消 a0j21
'      If adocaseprogress.State <> adStateClosed Then adocaseprogress.Close
'      adocaseprogress.CursorLocation = adUseClient
'      '尚有未收金額
'      adocaseprogress.Open "select a.*,b.*,c.*,na03 from acc0j0 a, caseprogress b, casepropertymap c,nation" & _
'                           " where cp09(+) = a0j01 and cp01 = cpm01 (+) and cp10 = cpm02 (+)" & _
'                           " and a0j13 = '" & adoacc0k0.Fields("a0k01").Value & "'" & _
'                           " and a0j01 = '" & adoacc0k0.Fields("CP09").Value & "'" & _
'                           " and nvl(cp79, 0) > 0 and na01(+)=a0j04 ", adoTaie, adOpenStatic, adLockReadOnly
'      'Modify by Morgan 2008/5/22 判斷有資料才印,否則會重疊
'      If adocaseprogress.RecordCount > 0 Then
'         'Modify By Sindy 2010/11/25 改為不同客戶編號+收據抬頭做跳頁
'         'If strNo <> adoacc0k0.Fields("a0k03").Value Then
'         'Modify By Sindy 2014/4/14 改為不同客戶編號+收據抬頭+公司別做跳頁
'         'If (strNo <> adoacc0k0.Fields("a0k03").Value) Or (strTitle <> adoacc0k0.Fields("a0k04").Value) Then
'         If (strNo <> adoacc0k0.Fields("a0k03").Value) Or (strTitle <> adoacc0k0.Fields("a0k04").Value) Or _
'            (strCompany <> adoacc0k0.Fields("a0k11").Value) Then
'            If lngAmount1 <> 0 Then
'               Call PrintSum(False, strCompany)
'            End If
'            intCounter = 0
'            intPage = 1
'            If m_curCompany <> "" Then
'               Printer.NewPage '*****
'            End If
'            Call PrintHead(adoacc0k0.Fields("a0k11").Value)
'            strNo = adoacc0k0.Fields("a0k03").Value
'            strTitle = adoacc0k0.Fields("a0k04").Value 'Add By Sindy 2010/11/25
'            strNoB = adoacc0k0.Fields("a0k03").Value 'Add By Sindy 2013/11/28
'            strTitleB = adoacc0k0.Fields("a0k04").Value 'Add By Sindy 2013/11/28
'            'Add By Sindy 2014/4/14
'            strCompany = adoacc0k0.Fields("a0k11").Value
'            strCompanyB = adoacc0k0.Fields("a0k11").Value
'            '2014/4/14 END
'            m_CP01 = "": m_CP02 = "": m_CP03 = "": m_CP04 = "'"
'         End If
'         If intCounter > 19 Then
'            intCounter = 0
'            Printer.NewPage
'            intPage = intPage + 1
'            Call PrintHead(adoacc0k0.Fields("a0k11").Value)
'         End If
'
'         If bolPrintItem = True Then 'Added by Morgan 2011/12/19
'            Printer.CurrentX = 200
'            Printer.CurrentY = 5600 + intCounter * 300 - intY
'            If IsNull(adoacc0k0.Fields("a0k02").Value) = False Then
'               Printer.Print CFDate(adoacc0k0.Fields("a0k02").Value)
'            Else
'               Printer.Print ""
'            End If
'            Printer.CurrentX = 1250
'            Printer.CurrentY = 5600 + intCounter * 300 - intY
'            Printer.Print adoacc0k0.Fields("a0k01").Value
'         End If 'Added by Morgan 2011/12/19
'
'         Do While adocaseprogress.EOF = False
'           'Add By Cheng 2003/07/21
'           '記錄本所案號
'           If m_CP01 = "" Then
'               m_CP01 = "" & adocaseprogress("CP01").Value
'               m_CP02 = "" & adocaseprogress("CP02").Value
'               m_CP03 = "" & adocaseprogress("CP03").Value
'               m_CP04 = "" & adocaseprogress("CP04").Value
'           End If
'            If intCounter > 19 Then
'               intCounter = 0
'               Printer.NewPage
'               intPage = intPage + 1
'               Call PrintHead(strCompany)
'            End If
'
'            If bolPrintItem = True Then 'Added by Morgan 2011/12/19
'               'Modify By Sindy 2020/4/27 + 將(本所案號)取消  含 商標 /智慧所 或 法律所
''               Printer.CurrentX = 2340
''               Printer.CurrentY = 5600 + intCounter * 300 - intY
''               If adocaseprogress.Fields("cp03").Value = "0" And adocaseprogress.Fields("cp04").Value = "00" Then
''                  Printer.Print adocaseprogress.Fields("cp01").Value & adocaseprogress.Fields("cp02").Value
''               Else
''                  Printer.Print adocaseprogress.Fields("cp01").Value & adocaseprogress.Fields("cp02").Value & adocaseprogress.Fields("cp03").Value & adocaseprogress.Fields("cp04").Value
''               End If
'
'               Printer.CurrentX = 2350 '3650
'               Printer.CurrentY = 5600 + intCounter * 300 - intY
'               'Modified by Morgan 2011/12/30 取消 a0j21
'               If IsNull(adocaseprogress.Fields("na03").Value) = False Then
'                  Printer.Print MidB(adocaseprogress.Fields("na03").Value, 1, 10)
'               Else
'                  Printer.Print ""
'               End If
'               Printer.CurrentX = 3800 '4880
'               Printer.CurrentY = 5600 + intCounter * 300 - intY
'
'               'Added by Morgan 2011/12/19
'               If adoacc0k0.Fields("a0k33") = "Y" Then
'                  Printer.Print MidB(adocaseprogress.Fields("a0j22").Value, 1, 10)
'               Else
'               'end 2011/12/19
'                  If adocaseprogress.Fields("a0j04").Value = "000" Then
'                     Printer.Print MidB(adocaseprogress.Fields("cpm03").Value, 1, 10)
'                  Else
'                     Printer.Print MidB(adocaseprogress.Fields("cpm04").Value, 1, 10)
'                  End If
'               End If
'
'               Printer.CurrentX = 5200 '6000
'               Printer.CurrentY = 5600 + intCounter * 300 - intY
'               strName = MidB(CaseNameShow(adocaseprogress.Fields("cp01").Value, adocaseprogress.Fields("cp02").Value, adocaseprogress.Fields("cp03").Value, adocaseprogress.Fields("cp04").Value, 1), 1, 24)
'               If strName = "" Then
'                  strName = MidB(CaseNameShow(adocaseprogress.Fields("cp01").Value, adocaseprogress.Fields("cp02").Value, adocaseprogress.Fields("cp03").Value, adocaseprogress.Fields("cp04").Value, 2), 1, 24)
'               End If
'               If strName = "" Then
'                  strName = MidB(CaseNameShow(adocaseprogress.Fields("cp01").Value, adocaseprogress.Fields("cp02").Value, adocaseprogress.Fields("cp03").Value, adocaseprogress.Fields("cp04").Value, 3), 1, 24)
'               End If
'               If m_CP01 = "LA" Then strName = "" 'Add By Sindy 2020/4/27 法律所的顧問案件,案件名稱欄放空白
'               Printer.Print strName
'            End If
'
'            If adoquery.State <> adStateClosed Then adoquery.Close
'            adoquery.CursorLocation = adUseClient
'            'Modified by Morgan 2011/11/1 考慮拆收據情形改抓0j0
'            'adoquery.Open "select sum(a1u04+a1u05) as Iamount, sum(a1u07+a1u09) as Camount, sum(a1u08+a1u10) as Pamount from acc1u0 where a1u03 = '" & adocaseprogress.Fields("cp09").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'            adoquery.Open "select sum(a1u04+a1u05) as Iamount, sum(a1u07+a1u09) as Camount, sum(a1u08+a1u10) as Pamount from acc1u0 where a1u02='" & adocaseprogress.Fields("a0j13").Value & "' and a1u03 = '" & adocaseprogress.Fields("a0j01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'            If adoquery.RecordCount <> 0 Then
'               If IsNull(adoquery.Fields("Iamount").Value) Then
'                  douIamount = 0
'               Else
'                  douIamount = adoquery.Fields("Iamount").Value
'               End If
'               If IsNull(adoquery.Fields("Camount").Value) Then
'                  douCAmount = 0
'               Else
'                  douCAmount = adoquery.Fields("Camount").Value
'               End If
'               If IsNull(adoquery.Fields("Pamount").Value) Then
'                  douPAmount = 0
'               Else
'                  douPAmount = adoquery.Fields("Pamount").Value
'               End If
'            Else
'               douIamount = 0
'               douCAmount = 0
'               douPAmount = 0
'            End If
'            adoquery.Close
'
'            If Val("" & adocaseprogress("a0j09")) + Val("" & adocaseprogress("a0j10")) > 0 Then
'               dblItemAmt1 = dblItemAmt1 + Val("" & adocaseprogress("a0j09")) + Val("" & adocaseprogress("a0j10")) - douCAmount
'            End If
'
'            If douIamount > 0 Then
'               dblItemAmt3 = dblItemAmt3 + douIamount - douPAmount
'            End If
'
'            adocaseprogress.MoveNext
'            bolHaveData = True 'Add By Sindy 2013/11/28
'         Loop
'      'Add By Sindy 2013/11/27
'      Else
'         strNoB = adoacc0k0.Fields("a0k03").Value
'         strTitleB = adoacc0k0.Fields("a0k04").Value
'         strCompanyB = adoacc0k0.Fields("a0k11").Value 'Add By Sindy 2014/4/14
'         bolHaveData = False 'Add By Sindy 2013/11/28
'      '2013/11/27 END
'      End If
'      adocaseprogress.Close
'
'      'Added by Morgan 2011/12/19
'      stLstItem = "" & adoacc0k0.Fields("a0j22")
'      stLstRecNo = "" & adoacc0k0.Fields("a0k01")
'      'end 2011/12/19
'
'      adoacc0k0.MoveNext
'
'      'Add By Sindy 2013/11/28
'      If bolHaveData = False Then
'         bolPrintItemAmt = False
'      Else
'      '2013/11/28 END
'         'Added by Morgan 2011/12/19
'         If adoacc0k0.EOF Then
'            bolPrintItemAmt = True
'         Else
'            '變更帳款類別=Y / 帳款類別 / 收據編號
'            If adoacc0k0.Fields("a0k33") = "Y" And stLstItem = adoacc0k0.Fields("a0j22") And stLstRecNo = adoacc0k0.Fields("a0k01") Then
'               bolPrintItem = False
'               bolPrintItemAmt = False
'            Else
'               bolPrintItem = True
'               bolPrintItemAmt = True
'            End If
'         End If
'      End If
'      If bolPrintItemAmt = True Then
'         strAmount1 = Format(dblItemAmt1, DDollar2)
'         intLength = Printer.TextWidth(strAmount1)
'         Printer.CurrentX = 9800 - intLength
'         Printer.CurrentY = 5600 + intCounter * 300 - intY
'         Printer.Print strAmount1
'
'         strAmount3 = Format(dblItemAmt3, DDollar2)
'         intLength = Printer.TextWidth(strAmount3)
'         Printer.CurrentX = 11100 - intLength
'         Printer.CurrentY = 5600 + intCounter * 300 - intY
'         Printer.Print strAmount3
'         intCounter = intCounter + 1
'
'         lngAmount1 = lngAmount1 + dblItemAmt1 '應收金額
'         lngAmount3 = lngAmount3 + dblItemAmt3 '已收金額
'         dblItemAmt1 = 0
'         dblItemAmt3 = 0
'      End If
'      'end 2011/12/19
''ReadNext: 'Add By Sindy 2012/10/31
'   Loop
''   'Add By Sindy 2012/10/31
''   If Trim(Text4) = "1" And dblCount = 0 Then
''      adoacc0k0.Close
''      MsgBox MsgText(28), , MsgText(5)
''      Exit Sub
''   End If
''   '2012/10/31 End
'
'   'Add By Sindy 2012/11/13 若選往來時,再依該客戶及收款日期去讀取收款資料
'   If Text4 = "1" Then
'      If lngAmount1 <> 0 Then
'         Call PrintSum(False, strCompanyB)
'      End If
'      Call PrintType1Data(strNoB, strTitleB, strCompanyB)
'   End If
'   '2012/11/13 End
'
'   If lngAmount1 <> 0 Then
'      Call PrintSum(False, strCompany)
'   End If
'   adoacc0k0.Close
'   Printer.EndDoc
'
'ErrHnd:
'   'Resume
'   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'*************************************************
' 列印明細
'
'*************************************************
Public Sub PrintDetail_Excel()
Dim douIamount As Double
Dim douCAmount As Double
Dim douPAmount As Double
Dim lngAmount As Double
Dim stVTB As String, stVTB2 As String, stVTB3 As String 'Add By Sindy 2010/12/1
Dim bolPrintItem As Boolean, bolPrintItemAmt As Boolean, stLstItem As String, stLstRecNo As String 'Added by Morgan 2011/12/19
Dim dblItemAmt1 As Double, dblItemAmt3 As Double 'Added by Morgan 2011/12/19
'Add By Sindy 2012/10/31
Dim strSQL2 As String, strIsCustGoTo As String, strA0j13A0j01 As String, strA1u02A1u03 As String
Dim dblAmt1 As Double, dblAmt2 As Double, dblAmtI As Double, dblAmtC As Double, dblAmtP As Double
Dim strIsCustChked As String, dblCount As Double
'2012/10/31 End
Dim bolHaveData As Boolean 'Add By Sindy 2013/11/28
   
On Error GoTo ErrHnd
   
   m_curCompany = ""
   intCounter = 0: intCounterD = 0
   lngAmount = 0
   intLength = 0
   intPage = 0
   lngAmount1 = 0
   lngAmount3 = 0
   strSql = ""
   strNo = "": strTitle = "" 'Modify By Sindy 2010/11/25
   strNoB = "": strTitleB = "" 'Modify By Sindy 2013/11/28
   strCompany = "": strCompanyB = "" 'Modify By Sindy 2014/4/14
   m_CP01 = "": m_CP02 = "": m_CP03 = "": m_CP04 = "'"
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and a0k03 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a0k03 <= '" & Text2 & "'"
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If Text3 <> MsgText(601) Then
      strSql = strSql & " and a0k20 = '" & Text3 & "'"
   End If
   
   'Modify By Sindy 2012/11/13 Mark 因均要顯示應收帳資料,再依該客戶及收款日期去讀取收款資料
   'Modify By Sindy 2013/11/26 還是要判斷報表類別 X37597
   If Text4 = "2" Then
      strSql = strSql & " and (a0k06+a0k07-decode(a0s05, null, 0, a0s05)) > (nvl(a0k17, 0)+nvl(a0k18, 0))"
   End If
   
   '若非北所員工, 只能列印該所資料
   If pub_strUserOffice <> "1" Then
       strSql = strSql & " And ''||ST06='" & pub_strUserOffice & "' "
   End If
   
   'Add By Sindy 2010/11/5 未列印收據的資料,不要發
   '2011/6/7 cancel by sonia 改在下面檢查a0k32及是否發文
   'strSql = strSql & " And a0k19>0 "
   '2010/11/5 End
   
   If adoacc0k0.State <> adStateClosed Then adoacc0k0.Close
   adoacc0k0.CursorLocation = adUseClient
   'Modify By Sindy 2010/12/1
   If txtType = "1" Then '1.單一
      stVTB2 = " HAVING COUNT(DISTINCT a0k04)<=1"
   Else '2.複合
      stVTB2 = " HAVING COUNT(DISTINCT a0k04)>1"
   End If
   
   'Modify By Sindy 2012/11/13 Mark 因均要顯示應收帳資料,再依該客戶及收款日期去讀取收款資料
   'Modify By Sindy 2013/11/26 還是要判斷報表類別 X37597
   If Text4 = "2" Then '2.應收帳
      stVTB3 = " and nvl(cp79, 0) > 0"
   Else '1.往來帳
      stVTB3 = " and (nvl(cp79, 0) > 0 or nvl(cp75,0)>nvl(cp78,0))"
   End If
   
   '2011/6/7 add by sonia 未列印收據者不印A0K32,但已發文者仍要印
   'stVTB3 = stVTB3 & " and (a0k32 is null or (a0k32 is not null and cp27 is not null)) "
   'Modify By Sindy 2012/11/9辜來電通知,恢復 控制a0k32 is null才印
   'Modified by Lydia 2025/06/10 (a0k32 is null) 改用函數判斷：geta0k32type(a0k01)='1'
   stVTB3 = stVTB3 & " and geta0k32type(a0k01)='1' "
   
   '2011/6/7 end
   'Modified by Morgan 2011/11/1 考慮拆收據情形
   'stVTB = "SELECT substr(A0K03,1,6) from acc0k0, customer, nation, (select a0s02, sum(nvl(a0s05,0)+nvl(a0s06, 0)+nvl(a0s07, 0)) as a0s05 from acc0s0 group by a0s02) acc0s0, Staff, caseprogress where cp60 = a0k01  and substr(a0k03, 1, 8) = cu01 (+) and substr(a0k03, 9, 1) = cu02 (+) and a0k23 = na01 (+) and a0k01 = a0s02 (+) and (a0k09 is null or a0k09 = 0) and (a0k06+a0k07) > decode(a0s05, null, 0, a0s05) And A0K20=ST01(+) " & strSql & stVTB3 & " GROUP BY substr(A0K03,1,6)" & stVTB2
   'Modify By Sindy 2014/4/11 + and (cu140<>'N' or cu140 is null) 不催款者,不產生對帳單
   'Modify by Amy 2020/09/18 原:and (cu140<>'N' or cu140 is null),因不寄催款單改輸1-3
   'Modify by Amy 2023/02/15 剔除ACS案-杜協理
   stVTB = "SELECT substr(A0K03,1,6) from acc0k0,acc0j0,customer" & _
           ",(select a0s02, sum(nvl(a0s05,0)+nvl(a0s06, 0)+nvl(a0s07, 0)) as a0s05 from acc0s0 group by a0s02) acc0s0" & _
           ",Staff,caseprogress" & _
           " where a0j13(+)=a0k01 and cp09(+)=a0j01" & _
           " and substr(a0k03, 1, 8) = cu01(+) and substr(a0k03, 9, 1) = cu02(+)" & _
           " and cu140 is null and a0k01 = a0s02(+) and (a0k09 is null or a0k09 = 0)" & _
           " and (a0k06+a0k07) > decode(a0s05, null, 0, a0s05) And SubStr(a0j02, 1, length(a0j02) - 9)<>'ACS' " & _
           " and A0K20=ST01(+) " & strSql & stVTB3 & " GROUP BY substr(A0K03,1,6)" & stVTB2
   '2011/6/7 modify by sonia 未列印收據者不印A0K32,但已發文者仍要印,且為免同一收據二個收文號資料會重覆故加CP09
   'strSql = "select A0K01,A0K02,A0K03,A0K04,CU30,CU31 from acc0k0, customer, nation, (select a0s02, sum(nvl(a0s05,0)+nvl(a0s06, 0)+nvl(a0s07, 0)) as a0s05 from acc0s0 group by a0s02) acc0s0, Staff where substr(a0k03, 1, 8) = cu01 (+) and substr(a0k03, 9, 1) = cu02 (+) and a0k23 = na01 (+) and a0k01 = a0s02 (+) and (a0k09 is null or a0k09 = 0) and (a0k06+a0k07) > decode(a0s05, null, 0, a0s05) And A0K20=ST01(+) " & strSql & " and substr(a0k03,1,6) in (" & stVTB & ") order by a0k03 asc, A0K04 asc, a0k01 asc"
   'Modified by Morgan 2011/11/1 考慮拆收據情形
   'strSql = "select A0K01,A0K02,A0K03,A0K04,CU30,CU31,CP09 from acc0k0, customer, nation, caseprogress, (select a0s02, sum(nvl(a0s05,0)+nvl(a0s06, 0)+nvl(a0s07, 0)) as a0s05 from acc0s0 group by a0s02) acc0s0, Staff where a0k01=cp60(+) and substr(a0k03, 1, 8) = cu01 (+) and substr(a0k03, 9, 1) = cu02 (+) and a0k23 = na01 (+) and a0k01 = a0s02 (+) and (a0k09 is null or a0k09 = 0) and (a0k06+a0k07) > decode(a0s05, null, 0, a0s05) And A0K20=ST01(+) " & strSql & stVTB3 & " and substr(a0k03,1,6) in (" & stVTB & ") order by a0k03 asc, A0K04 asc, a0k01 asc"
   'Modify By Sindy 2014/4/11 + and (cu140<>'N' or cu140 is null) 不催款者,不產生對帳單
   '                          + ,decode(a0k11,'J','J','1') as A0K11
   '                          order by + ,a0k11 asc
   'Modify By Sindy 2020/4/20 decode(a0k11,'J','J','1') => decode(a0k11,'J','J','L','L','1')
   'Modify by Amy 2020/09/18 原:and (cu140<>'N' or cu140 is null),因不寄催款單改輸1-3
   strSQL2 = "select A0K01,A0K02,A0K03,A0K04,CU30,CU31,CP09,A0K33,A0J22,decode(a0k11,'J','J','L','L','1') as A0K11 " & _
                    "from acc0k0,acc0j0, customer, caseprogress, (select a0s02, sum(nvl(a0s05,0)+nvl(a0s06, 0)+nvl(a0s07, 0)) as a0s05 from acc0s0 group by a0s02) acc0s0, Staff " & _
                    "where a0j13(+)=a0k01 and cp09(+)=a0j01 and substr(a0k03, 1, 8) = cu01(+) and substr(a0k03, 9, 1) = cu02(+) and cu140 is null and a0k01 = a0s02 (+) " & _
                    "and (a0k09 is null or a0k09 = 0) and (a0k06+a0k07) > decode(a0s05, null, 0, a0s05) And A0K20=ST01(+) And SubStr(a0j02, 1, length(a0j02) - 9)<>'ACS' " & _
                    strSql & stVTB3 & " and substr(a0k03,1,6) in (" & stVTB & ") order by a0k03 asc,A0K04 asc,a0k11 asc,a0k01 asc,a0j25 asc"
   '2010/12/1 End
   'end 2023/02/15
   adoacc0k0.Open strSQL2, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0k0.RecordCount = 0 Then
      adoacc0k0.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   
   txtNote.Visible = True
   'Added by Morgan 2011/12/19 若收據有變更帳款類別則相同的依照列印順序合併
   bolPrintItem = True
   dblItemAmt1 = 0
   dblItemAmt3 = 0
   'end 2011/12/19
   
   Do While adoacc0k0.EOF = False
      'Modify by Morgan 2008/6/2
      'adocaseprogress.Open "select * from caseprogress, acc0j0, casepropertymap where cp09 = a0j01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp60 = '" & adoacc0k0.Fields("a0k01").Value & "' and nvl(cp16, 0) > nvl(cp77, 0)", adoTaie, adOpenStatic, adLockReadOnly
      'Add By Sindy 2012/11/13 若選往來時,再依該客戶及收款日期去讀取收款資料
      If Text4 = "1" Then
         If (strNoB <> "" And strNoB <> adoacc0k0.Fields("a0k03").Value) Or _
            (strTitleB <> "" And strTitleB <> adoacc0k0.Fields("a0k04").Value) Or _
            (strCompanyB <> "" And strCompanyB <> adoacc0k0.Fields("a0k11").Value) Then 'Modify By Sindy 2014/4/14 +strCompanyB
            If lngAmount1 <> 0 Then
               Call PrintSum_Excel(False, strCompanyB)
            End If
            Call PrintType1Data_Excel(strNoB, strTitleB, strCompanyB)
         End If
      End If
      '2012/11/13 End
      
      'Modify By Sindy 2012/11/13 Mark 因均要顯示應收帳資料,再依該客戶及收款日期去讀取收款資料
      '2011/6/7 modify by sonia 加cp09條件
      'adocaseprogress.Open "select * from caseprogress, acc0j0, casepropertymap where cp09 = a0j01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp60 = '" & adoacc0k0.Fields("a0k01").Value & "' and nvl(cp79, 0) > 0", adoTaie, adOpenStatic, adLockReadOnly
      'Modified by Morgan 2011/11/1 考慮拆收據情形
      'adocaseprogress.Open "select * from caseprogress, acc0j0, casepropertymap where cp09 = a0j01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp60 = '" & adoacc0k0.Fields("a0k01").Value & "' and cp09 = '" & adoacc0k0.Fields("CP09").Value & "' and nvl(cp79, 0) > 0", adoTaie, adOpenStatic, adLockReadOnly
      'Modified by Morgan 2011/12/30 取消 a0j21
      If adocaseprogress.State <> adStateClosed Then adocaseprogress.Close
      adocaseprogress.CursorLocation = adUseClient
      '尚有未收金額
      adocaseprogress.Open "select a.*,b.*,c.*,na03 from acc0j0 a, caseprogress b, casepropertymap c,nation" & _
                           " where cp09(+) = a0j01 and cp01 = cpm01 (+) and cp10 = cpm02 (+)" & _
                           " and a0j13 = '" & adoacc0k0.Fields("a0k01").Value & "'" & _
                           " and a0j01 = '" & adoacc0k0.Fields("CP09").Value & "'" & _
                           " and nvl(cp79, 0) > 0 and na01(+)=a0j04 ", adoTaie, adOpenStatic, adLockReadOnly
      'Modify by Morgan 2008/5/22 判斷有資料才印,否則會重疊
      If adocaseprogress.RecordCount > 0 Then
         'Modify By Sindy 2010/11/25 改為不同客戶編號+收據抬頭做跳頁
         'If strNo <> adoacc0k0.Fields("a0k03").Value Then
         'Modify By Sindy 2014/4/14 改為不同客戶編號+收據抬頭+公司別做跳頁
         'If (strNo <> adoacc0k0.Fields("a0k03").Value) Or (strTitle <> adoacc0k0.Fields("a0k04").Value) Then
         If (strNo <> adoacc0k0.Fields("a0k03").Value) Or (strTitle <> adoacc0k0.Fields("a0k04").Value) Or _
            (strCompany <> adoacc0k0.Fields("a0k11").Value) Then
            If lngAmount1 <> 0 Then
               Call PrintSum_Excel(False, strCompany)
            End If
'            If m_curCompany <> "" Then
'               Printer.NewPage '*****
'            End If
            intPage = 1
            Call PrintHead_Excel(adoacc0k0.Fields("a0k11").Value, _
                  adoacc0k0.Fields("a0k03").Value, adoacc0k0.Fields("a0k04").Value, "1")
            strNo = adoacc0k0.Fields("a0k03").Value
            strTitle = adoacc0k0.Fields("a0k04").Value 'Add By Sindy 2010/11/25
            strNoB = adoacc0k0.Fields("a0k03").Value 'Add By Sindy 2013/11/28
            strTitleB = adoacc0k0.Fields("a0k04").Value 'Add By Sindy 2013/11/28
            'Add By Sindy 2014/4/14
            strCompany = adoacc0k0.Fields("a0k11").Value
            strCompanyB = adoacc0k0.Fields("a0k11").Value
            '2014/4/14 END
            m_CP01 = "": m_CP02 = "": m_CP03 = "": m_CP04 = "'"
         End If
         
         '檢查明細筆數是否已滿,滿了即換頁
         If intCounterD Mod intPageRow = 0 Then
            intPage = intPage + 1
            Call PrintHead_Excel(adoacc0k0.Fields("a0k11").Value, _
                  adoacc0k0.Fields("a0k03").Value, adoacc0k0.Fields("a0k04").Value, "1")
         End If
         If bolPrintItem = True Then 'Added by Morgan 2011/12/19
            intCounter = intCounter + 1: intCounterD = intCounterD + 1
            '收據日期
            If IsNull(adoacc0k0.Fields("a0k02").Value) = False Then
               xlsAnnuity.Range("A" & intCounter).Value = CFDate(adoacc0k0.Fields("a0k02").Value)
            Else
               xlsAnnuity.Range("A" & intCounter).Value = ""
            End If
            '收據號碼
            xlsAnnuity.Range("B" & intCounter).Value = adoacc0k0.Fields("a0k01").Value
         End If 'Added by Morgan 2011/12/19
      
         Do While adocaseprogress.EOF = False
           'Add By Cheng 2003/07/21
           '記錄本所案號
           If m_CP01 = "" Then
               m_CP01 = "" & adocaseprogress("CP01").Value
               m_CP02 = "" & adocaseprogress("CP02").Value
               m_CP03 = "" & adocaseprogress("CP03").Value
               m_CP04 = "" & adocaseprogress("CP04").Value
           End If
'            If intCounter Mod intPageTotRow = 0 Then
'               intPage = intPage + 1
'               Call PrintHead_Excel(strCompany, strNo, strTitle, "1")
'            End If
            
            If bolPrintItem = True Then 'Added by Morgan 2011/12/19
               '國別
               If IsNull(adocaseprogress.Fields("na03").Value) = False Then
                  xlsAnnuity.Range("C" & intCounter).Value = MidB(adocaseprogress.Fields("na03").Value, 1, 10)
               Else
                  xlsAnnuity.Range("C" & intCounter).Value = ""
               End If
               '案件性質
               'Added by Morgan 2011/12/19
               If adoacc0k0.Fields("a0k33") = "Y" Then
                  xlsAnnuity.Range("D" & intCounter).Value = MidB(adocaseprogress.Fields("a0j22").Value, 1, 12)
               Else
               'end 2011/12/19
                  If adocaseprogress.Fields("a0j04").Value = "000" Then
                     xlsAnnuity.Range("D" & intCounter).Value = MidB(adocaseprogress.Fields("cpm03").Value, 1, 12)
                  Else
                     xlsAnnuity.Range("D" & intCounter).Value = MidB(adocaseprogress.Fields("cpm04").Value, 1, 12)
                  End If
               End If
               '案件名稱
               strName = MidB(CaseNameShow(adocaseprogress.Fields("cp01").Value, adocaseprogress.Fields("cp02").Value, adocaseprogress.Fields("cp03").Value, adocaseprogress.Fields("cp04").Value, 1), 1, 24)
               If strName = "" Then
                  strName = MidB(CaseNameShow(adocaseprogress.Fields("cp01").Value, adocaseprogress.Fields("cp02").Value, adocaseprogress.Fields("cp03").Value, adocaseprogress.Fields("cp04").Value, 2), 1, 24)
               End If
               If strName = "" Then
                  strName = MidB(CaseNameShow(adocaseprogress.Fields("cp01").Value, adocaseprogress.Fields("cp02").Value, adocaseprogress.Fields("cp03").Value, adocaseprogress.Fields("cp04").Value, 3), 1, 24)
               End If
               If m_CP01 = "LA" Then strName = "" 'Add By Sindy 2020/4/27 法律所的顧問案件,案件名稱欄放空白
               xlsAnnuity.Range("E" & intCounter).Value = strName
            End If
            
            If adoquery.State <> adStateClosed Then adoquery.Close
            adoquery.CursorLocation = adUseClient
            'Modified by Morgan 2011/11/1 考慮拆收據情形改抓0j0
            'adoquery.Open "select sum(a1u04+a1u05) as Iamount, sum(a1u07+a1u09) as Camount, sum(a1u08+a1u10) as Pamount from acc1u0 where a1u03 = '" & adocaseprogress.Fields("cp09").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
            adoquery.Open "select sum(a1u04+a1u05) as Iamount, sum(a1u07+a1u09) as Camount, sum(a1u08+a1u10) as Pamount from acc1u0 where a1u02='" & adocaseprogress.Fields("a0j13").Value & "' and a1u03 = '" & adocaseprogress.Fields("a0j01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
            If adoquery.RecordCount <> 0 Then
               If IsNull(adoquery.Fields("Iamount").Value) Then
                  douIamount = 0
               Else
                  douIamount = adoquery.Fields("Iamount").Value
               End If
               If IsNull(adoquery.Fields("Camount").Value) Then
                  douCAmount = 0
               Else
                  douCAmount = adoquery.Fields("Camount").Value
               End If
               If IsNull(adoquery.Fields("Pamount").Value) Then
                  douPAmount = 0
               Else
                  douPAmount = adoquery.Fields("Pamount").Value
               End If
            Else
               douIamount = 0
               douCAmount = 0
               douPAmount = 0
            End If
            adoquery.Close
            
            If Val("" & adocaseprogress("a0j09")) + Val("" & adocaseprogress("a0j10")) > 0 Then
               dblItemAmt1 = dblItemAmt1 + Val("" & adocaseprogress("a0j09")) + Val("" & adocaseprogress("a0j10")) - douCAmount
            End If
            
            If douIamount > 0 Then
               dblItemAmt3 = dblItemAmt3 + douIamount - douPAmount
            End If
            
            adocaseprogress.MoveNext
            bolHaveData = True 'Add By Sindy 2013/11/28
         Loop
      'Add By Sindy 2013/11/27
      Else
         strNoB = adoacc0k0.Fields("a0k03").Value
         strTitleB = adoacc0k0.Fields("a0k04").Value
         strCompanyB = adoacc0k0.Fields("a0k11").Value 'Add By Sindy 2014/4/14
         bolHaveData = False 'Add By Sindy 2013/11/28
      '2013/11/27 END
      End If
      adocaseprogress.Close
      
      'Added by Morgan 2011/12/19
      stLstItem = "" & adoacc0k0.Fields("a0j22")
      stLstRecNo = "" & adoacc0k0.Fields("a0k01")
      'end 2011/12/19
      
      adoacc0k0.MoveNext
      
      'Add By Sindy 2013/11/28
      If bolHaveData = False Then
         bolPrintItemAmt = False
      Else
      '2013/11/28 END
         'intCounter = intCounter + 1
         'Added by Morgan 2011/12/19
         If adoacc0k0.EOF Then
            bolPrintItemAmt = True
         Else
            '變更帳款類別=Y / 帳款類別 / 收據編號
            If adoacc0k0.Fields("a0k33") = "Y" And stLstItem = adoacc0k0.Fields("a0j22") And stLstRecNo = adoacc0k0.Fields("a0k01") Then
               bolPrintItem = False
               bolPrintItemAmt = False
            Else
               bolPrintItem = True
               bolPrintItemAmt = True
            End If
         End If
      End If
      If bolPrintItemAmt = True Then
         '應收金額
         strAmount1 = Format(dblItemAmt1, DDollar2)
         xlsAnnuity.Range("F" & intCounter).Value = strAmount1
         '已收金額
         strAmount3 = Format(dblItemAmt3, DDollar2)
         xlsAnnuity.Range("G" & intCounter).Value = strAmount3
         
         lngAmount1 = lngAmount1 + dblItemAmt1 '應收金額
         lngAmount3 = lngAmount3 + dblItemAmt3 '已收金額
         dblItemAmt1 = 0
         dblItemAmt3 = 0
      End If
      'end 2011/12/19
'ReadNext: 'Add By Sindy 2012/10/31
   Loop
   
   'Add By Sindy 2012/11/13 若選往來時,再依該客戶及收款日期去讀取收款資料
   If Text4 = "1" Then
      If lngAmount1 <> 0 Then
         Call PrintSum_Excel(False, strCompanyB)
      End If
      Call PrintType1Data_Excel(strNoB, strTitleB, strCompanyB)
   End If
   '2012/11/13 End
   
   If lngAmount1 <> 0 Then
      Call PrintSum_Excel(False, strCompany)
   End If
   adoacc0k0.Close
   
   Call PrintExcel '列印,Close
   Set xlsAnnuity = Nothing
   
   txtNote.Visible = False
   
ErrHnd:
   'Resume
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'Add By Sindy 2022/4/8
Private Sub PrintExcel()
   '列印,Close Sheet
   With xlsAnnuity.ActiveSheet.PageSetup
      '.Zoom = False
      '.FitToPagesTall = 1000 '縮放成一頁高; 預設為1,筆數多時會縮小
      '.FitToPagesWide = 1 '縮放成一頁寬
   End With
   'strTempFile = App.path & "\" & strUserNum & "\$$demo"
   xlsAnnuity.Workbooks(1).PrintOut
   xlsAnnuity.Workbooks.Close 'SaveChanges:=False
   xlsAnnuity.Quit
End Sub

'Add By Sindy 2012/11/14 收款資料
'Mark by Amy 2023/05/22 不使用
Private Sub PrintType1Data(stra0k03 As String, strA0K04 As String, strCompany As String)
'Dim strSubNo As String, strSubTitle As String, strSubCompany As String
'Dim bolPrintItem As Boolean, bolPrintItemAmt As Boolean
'Dim dblItemAmt1 As Double, dblItemAmt3 As Double
'Dim stLstItem As String, stLstRecNo As String
'Dim strCon As String
'
'   bolPrintItem = True
'
'   'Add By Sindy 2014/4/14
'   If strCompany = "J" Then
'      strCon = " and a0k11='J'"
'   ElseIf strCompany <> "" Then
'      strCon = " and a0k11<>'J'"
'   End If
'   '2014/4/14 END
'
'   If adocaseprogress.State <> adStateClosed Then adocaseprogress.Close
'   adocaseprogress.CursorLocation = adUseClient
'   'Modify By Sindy 2014/4/11 + and (cu140<>'N' or cu140 is null) 不催款者,不產生對帳單
'   '                          + ,decode(a0k11,'J','J','1') as a0k11
'   '                          order by + , a0k11 asc
'   'Modify By Sindy 2020/4/20 decode(a0k11,'J','J','1') => decode(a0k11,'J','J','L','L','1')
'   'Modify by Amy 2020/09/18 原:and (cu140<>'N' or cu140 is null),因不寄催款單改輸1-3
'   strSql = "select * from (" & _
'            "select a1u01,a0L02 as a0L02,a0j01,a0k01,a0k02,a0k03,a0k04,a0k33,a0j04,a0j22,cp01,cp02,cp03,cp04,cp10,na03,cpm03,cpm04,cu30,cu31,a1u04+a1u05 as Amt,decode(a0k11,'J','J','L','L','1') as a0k11 " & _
'            "From acc0L0, acc1u0, acc0j0, acc0k0, caseprogress, nation, casepropertymap, customer " & _
'            "Where a0L02>=" & Val(FCDate(MaskEdBox3.Text)) & " And a0L02<=" & Val(FCDate(MaskEdBox4.Text)) & " " & _
'            "and a0L01=a1u01(+) and a1u02=a0k01(+) and a1u03=cp09(+) and a1u02=a0j13(+) and a1u03=a0j01(+) " & _
'            "and a0j04=na01(+) and substr(a0k03,1,8)=cu01(+) and substr(a0k03,9,1)=cu02(+) and cu140 is null and cp01=cpm01(+) and cp10=cpm02(+) " & _
'            "and a0k03='" & stra0k03 & "' and a0k04='" & strA0K04 & "'" & strCon & _
'            " Union " & _
'            "select a1u01,a0S03 as a0L02,a0j01,a0k01,a0k02,a0k03,a0k04,a0k33,a0j04,a0j22,cp01,cp02,cp03,cp04,cp10,na03,cpm03,cpm04,cu30,cu31,((a1u08+a1u10) * -1) as Amt,decode(a0k11,'J','J','L','L','1') as a0k11 " & _
'            "From acc0S0, acc1u0, acc0j0, acc0k0, caseprogress, nation, casepropertymap, customer " & _
'            "Where a0S03>=" & Val(FCDate(MaskEdBox3.Text)) & " " & _
'            "and a0S01=a1u01(+) and a1u02=a0k01(+) and a1u03=cp09(+) and a1u02=a0j13(+) and a1u03=a0j01(+) " & _
'            "and a0j04=na01(+) and substr(a0k03,1,8)=cu01(+) and substr(a0k03,9,1)=cu02(+) and cu140 is null and cp01=cpm01(+) and cp10=cpm02(+) " & _
'            "and a0k03='" & stra0k03 & "' and a0k04='" & strA0K04 & "'" & strCon & ") " & _
'            "where a0k01||a0j01 not in(" & _
'            "select u1.a1u02||u1.a1u03 " & _
'            "from acc1u0 u1,acc1u0 u2 where u1.a1u02=a0k01 and u1.a1u03=a0j01 and substr(u1.a1u01,1,1)='F' " & _
'            "and u2.a1u02=a0k01 and u2.a1u03=a0j01 and substr(u2.a1u01,1,1)='I' " & _
'            "and u1.a1u02=u2.a1u02(+) and u1.a1u03=u2.a1u03(+) " & _
'            "having Sum(u1.a1u04 + u1.a1u05) = Sum(u2.a1u08 + u2.a1u10) " & _
'            "group by u1.a1u01,u1.a1u02,u1.a1u03) " & _
'            "order by a0k03 asc, A0K04 asc, a0k11 asc, a0L02 asc, a0k01 asc "
'   adocaseprogress.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adocaseprogress.RecordCount > 0 Then
'      adocaseprogress.MoveFirst
'
'      '客戶編號+收據抬頭做跳頁
'      'Modify By Sindy 2014/4/14 再加公司別做跳頁
'      'If (strSubNo <> adocaseprogress.Fields("a0k03").Value) Or (strSubTitle <> adocaseprogress.Fields("a0k04").Value) Then
'      If (strSubNo <> adocaseprogress.Fields("a0k03").Value) Or (strSubTitle <> adocaseprogress.Fields("a0k04").Value) Or _
'         (strSubCompany <> adocaseprogress.Fields("a0k11").Value) Then
'         If lngAmount3 <> 0 Then
'            lngAmount1 = lngAmount3 '為了讓往來結餘為0
'            Call PrintSum(True, strSubCompany)
'         End If
'         intPage = 1
'         If m_curCompany <> "" Then
'            Printer.NewPage '*****
'         End If
'         intCounterB = 0
'         Call PrintHead2(adocaseprogress.Fields("a0k11").Value)
'         strSubNo = adocaseprogress.Fields("a0k03").Value
'         strSubTitle = adocaseprogress.Fields("a0k04").Value
'         strSubCompany = adocaseprogress.Fields("a0k11").Value 'Add By Sindy 2014/4/14
'         m_CP01 = "": m_CP02 = "": m_CP03 = "": m_CP04 = "'"
'      End If
'      If intCounterB > 19 Then
'         Printer.NewPage
'         intPage = intPage + 1
'         intCounterB = 0
'         Call PrintHead2(adocaseprogress.Fields("a0k11").Value)
'      End If
'
'      Do While adocaseprogress.EOF = False
'         '記錄本所案號
'         If m_CP01 = "" Then
'             m_CP01 = "" & adocaseprogress("CP01").Value
'             m_CP02 = "" & adocaseprogress("CP02").Value
'             m_CP03 = "" & adocaseprogress("CP03").Value
'             m_CP04 = "" & adocaseprogress("CP04").Value
'         End If
'         If intCounterB > 19 Then
'            Printer.NewPage
'            intPage = intPage + 1
'            intCounterB = 0
'            Call PrintHead2(strSubCompany)
'         End If
'
'         If bolPrintItem = True Then
'            Printer.CurrentX = 200
'            Printer.CurrentY = 5600 + intCounterB * 300 - intY
'            'Modify By Sindy 2013/11/27 帶收款日
''            If IsNull(adocaseprogress.Fields("a0k02").Value) = False Then
''               Printer.Print CFDate(adocaseprogress.Fields("a0k02").Value)
''            Else
''               Printer.Print ""
''            End If
'            If IsNull(adocaseprogress.Fields("a0L02").Value) = False Then
'               Printer.Print CFDate(adocaseprogress.Fields("a0L02").Value)
'            Else
'               Printer.Print ""
'            End If
'            '2013/11/27 END
'            Printer.CurrentX = 1250
'            Printer.CurrentY = 5600 + intCounterB * 300 - intY
'            Printer.Print adocaseprogress.Fields("a0k01").Value
'
'            'Modify By Sindy 2020/4/27 + 將(本所案號)取消  含 商標 /智慧所 或 法律所
''            Printer.CurrentX = 2340
''            Printer.CurrentY = 5600 + intCounterB * 300 - intY
''            If adocaseprogress.Fields("cp03").Value = "0" And adocaseprogress.Fields("cp04").Value = "00" Then
''               Printer.Print adocaseprogress.Fields("cp01").Value & adocaseprogress.Fields("cp02").Value
''            Else
''               Printer.Print adocaseprogress.Fields("cp01").Value & adocaseprogress.Fields("cp02").Value & adocaseprogress.Fields("cp03").Value & adocaseprogress.Fields("cp04").Value
''            End If
'
'            Printer.CurrentX = 2350 '3650
'            Printer.CurrentY = 5600 + intCounterB * 300 - intY
'            If IsNull(adocaseprogress.Fields("na03").Value) = False Then
'               Printer.Print MidB(adocaseprogress.Fields("na03").Value, 1, 10)
'            Else
'               Printer.Print ""
'            End If
'            Printer.CurrentX = 3800 '4880
'            Printer.CurrentY = 5600 + intCounterB * 300 - intY
'            If adocaseprogress.Fields("a0k33") = "Y" Then
'               Printer.Print MidB(adocaseprogress.Fields("a0j22").Value, 1, 10)
'            Else
'               If adocaseprogress.Fields("a0j04").Value = "000" Then
'                  Printer.Print MidB(adocaseprogress.Fields("cpm03").Value, 1, 10)
'               Else
'                  Printer.Print MidB(adocaseprogress.Fields("cpm04").Value, 1, 10)
'               End If
'            End If
'            Printer.CurrentX = 5200 '6000
'            Printer.CurrentY = 5600 + intCounterB * 300 - intY
'            strName = MidB(CaseNameShow(adocaseprogress.Fields("cp01").Value, adocaseprogress.Fields("cp02").Value, adocaseprogress.Fields("cp03").Value, adocaseprogress.Fields("cp04").Value, 1), 1, 24)
'            If strName = "" Then
'               strName = MidB(CaseNameShow(adocaseprogress.Fields("cp01").Value, adocaseprogress.Fields("cp02").Value, adocaseprogress.Fields("cp03").Value, adocaseprogress.Fields("cp04").Value, 2), 1, 24)
'            End If
'            If strName = "" Then
'               strName = MidB(CaseNameShow(adocaseprogress.Fields("cp01").Value, adocaseprogress.Fields("cp02").Value, adocaseprogress.Fields("cp03").Value, adocaseprogress.Fields("cp04").Value, 3), 1, 24)
'            End If
'            If m_CP01 = "LA" Then strName = "" 'Add By Sindy 2020/4/27 法律所的顧問案件,案件名稱欄放空白
'            Printer.Print strName
'         End If
'
'         dblItemAmt1 = 0
'         'If Val("" & adocaseprogress("Amt")) > 0 Then
'            dblItemAmt3 = dblItemAmt3 + Val("" & adocaseprogress("Amt"))
'         'End If
'
'         stLstItem = "" & adocaseprogress.Fields("a0j22")
'         stLstRecNo = "" & adocaseprogress.Fields("a0k01")
'
'         adocaseprogress.MoveNext
'
'         If adocaseprogress.EOF Then
'            bolPrintItemAmt = True
'         Else
'            If adocaseprogress.Fields("a0k33") = "Y" And stLstItem = adocaseprogress.Fields("a0j22") And stLstRecNo = adocaseprogress.Fields("a0k01") Then
'               bolPrintItem = False
'               bolPrintItemAmt = False
'            Else
'               bolPrintItem = True
'               bolPrintItemAmt = True
'            End If
'         End If
'         If bolPrintItemAmt = True Then
'            strAmount1 = Format(dblItemAmt1, DDollar2)
'            intLength = Printer.TextWidth(strAmount1)
'            Printer.CurrentX = 9800 - intLength
'            Printer.CurrentY = 5600 + intCounterB * 300 - intY
'            Printer.Print strAmount1
'
'            strAmount3 = Format(dblItemAmt3, DDollar2)
'            intLength = Printer.TextWidth(strAmount3)
'            Printer.CurrentX = 11100 - intLength
'            Printer.CurrentY = 5600 + intCounterB * 300 - intY
'            Printer.Print strAmount3
'            intCounterB = intCounterB + 1
'
'            lngAmount1 = lngAmount1 + dblItemAmt1
'            lngAmount3 = lngAmount3 + dblItemAmt3
'            dblItemAmt1 = 0
'            dblItemAmt3 = 0
'         End If
'      Loop
'   End If
'   If lngAmount3 <> 0 Then
'      lngAmount1 = lngAmount3 '為了讓往來結餘為0
'      Call PrintSum(True, strSubCompany)
'   End If
'   adocaseprogress.Close
End Sub

'Add By Sindy 2012/11/14 收款資料
Private Sub PrintType1Data_Excel(stra0k03 As String, strA0K04 As String, strCompany As String)
Dim strSubNo As String, strSubTitle As String, strSubCompany As String
Dim bolPrintItem As Boolean, bolPrintItemAmt As Boolean
Dim dblItemAmt1 As Double, dblItemAmt3 As Double
Dim stLstItem As String, stLstRecNo As String
Dim strCon As String
   
   bolPrintItem = True
   
   'Add By Sindy 2014/4/14
   If strCompany = "J" Then
      strCon = " and a0k11='J'"
   ElseIf strCompany <> "" Then
      strCon = " and a0k11<>'J'"
   End If
   '2014/4/14 END
   
   If adocaseprogress.State <> adStateClosed Then adocaseprogress.Close
   adocaseprogress.CursorLocation = adUseClient
   'Modify By Sindy 2014/4/11 + and (cu140<>'N' or cu140 is null) 不催款者,不產生對帳單
   '                          + ,decode(a0k11,'J','J','1') as a0k11
   '                          order by + , a0k11 asc
   'Modify By Sindy 2020/4/20 decode(a0k11,'J','J','1') => decode(a0k11,'J','J','L','L','1')
   'Modify by Amy 2020/09/18 原:and (cu140<>'N' or cu140 is null),因不寄催款單改輸1-3
   strSql = "select * from (" & _
            "select a1u01,a0L02 as a0L02,a0j01,a0k01,a0k02,a0k03,a0k04,a0k33,a0j04,a0j22,cp01,cp02,cp03,cp04,cp10,na03,cpm03,cpm04,cu30,cu31,a1u04+a1u05 as Amt,decode(a0k11,'J','J','L','L','1') as a0k11 " & _
            "From acc0L0, acc1u0, acc0j0, acc0k0, caseprogress, nation, casepropertymap, customer " & _
            "Where a0L02>=" & Val(FCDate(MaskEdBox3.Text)) & " And a0L02<=" & Val(FCDate(MaskEdBox4.Text)) & " " & _
            "and a0L01=a1u01(+) and a1u02=a0k01(+) and a1u03=cp09(+) and a1u02=a0j13(+) and a1u03=a0j01(+) " & _
            "and a0j04=na01(+) and substr(a0k03,1,8)=cu01(+) and substr(a0k03,9,1)=cu02(+) and cu140 is null and cp01=cpm01(+) and cp10=cpm02(+) " & _
            "and a0k03='" & stra0k03 & "' and a0k04='" & strA0K04 & "'" & strCon & _
            " Union " & _
            "select a1u01,a0S03 as a0L02,a0j01,a0k01,a0k02,a0k03,a0k04,a0k33,a0j04,a0j22,cp01,cp02,cp03,cp04,cp10,na03,cpm03,cpm04,cu30,cu31,((a1u08+a1u10) * -1) as Amt,decode(a0k11,'J','J','L','L','1') as a0k11 " & _
            "From acc0S0, acc1u0, acc0j0, acc0k0, caseprogress, nation, casepropertymap, customer " & _
            "Where a0S03>=" & Val(FCDate(MaskEdBox3.Text)) & " " & _
            "and a0S01=a1u01(+) and a1u02=a0k01(+) and a1u03=cp09(+) and a1u02=a0j13(+) and a1u03=a0j01(+) " & _
            "and a0j04=na01(+) and substr(a0k03,1,8)=cu01(+) and substr(a0k03,9,1)=cu02(+) and cu140 is null and cp01=cpm01(+) and cp10=cpm02(+) " & _
            "and a0k03='" & stra0k03 & "' and a0k04='" & strA0K04 & "'" & strCon & ") " & _
            "where a0k01||a0j01 not in(" & _
            "select u1.a1u02||u1.a1u03 " & _
            "from acc1u0 u1,acc1u0 u2 where u1.a1u02=a0k01 and u1.a1u03=a0j01 and substr(u1.a1u01,1,1)='F' " & _
            "and u2.a1u02=a0k01 and u2.a1u03=a0j01 and substr(u2.a1u01,1,1)='I' " & _
            "and u1.a1u02=u2.a1u02(+) and u1.a1u03=u2.a1u03(+) " & _
            "having Sum(u1.a1u04 + u1.a1u05) = Sum(u2.a1u08 + u2.a1u10) " & _
            "group by u1.a1u01,u1.a1u02,u1.a1u03) " & _
            "order by a0k03 asc, A0K04 asc, a0k11 asc, a0L02 asc, a0k01 asc "
   adocaseprogress.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adocaseprogress.RecordCount > 0 Then
      adocaseprogress.MoveFirst
      
      '客戶編號+收據抬頭做跳頁
      'Modify By Sindy 2014/4/14 再加公司別做跳頁
      'If (strSubNo <> adocaseprogress.Fields("a0k03").Value) Or (strSubTitle <> adocaseprogress.Fields("a0k04").Value) Then
      If (strSubNo <> adocaseprogress.Fields("a0k03").Value) Or (strSubTitle <> adocaseprogress.Fields("a0k04").Value) Or _
         (strSubCompany <> adocaseprogress.Fields("a0k11").Value) Then
         If lngAmount3 <> 0 Then
            lngAmount1 = lngAmount3 '為了讓往來結餘為0
            Call PrintSum_Excel(True, strSubCompany)
         End If
         intPage = 1
         Call PrintHead_Excel(adocaseprogress.Fields("a0k11").Value, _
               adocaseprogress.Fields("a0k03").Value, adocaseprogress.Fields("a0k04").Value, "2")
         strSubNo = adocaseprogress.Fields("a0k03").Value
         strSubTitle = adocaseprogress.Fields("a0k04").Value
         strSubCompany = adocaseprogress.Fields("a0k11").Value 'Add By Sindy 2014/4/14
         m_CP01 = "": m_CP02 = "": m_CP03 = "": m_CP04 = "'"
      End If
      
      '檢查明細筆數是否已滿,滿了即換頁
      If intCounterD Mod intPageRow = 0 Then
         intPage = intPage + 1
         Call PrintHead_Excel(adocaseprogress.Fields("a0k11").Value, _
               adocaseprogress.Fields("a0k03").Value, adocaseprogress.Fields("a0k04").Value, "2")
      End If
      Do While adocaseprogress.EOF = False
         '記錄本所案號
         If m_CP01 = "" Then
             m_CP01 = "" & adocaseprogress("CP01").Value
             m_CP02 = "" & adocaseprogress("CP02").Value
             m_CP03 = "" & adocaseprogress("CP03").Value
             m_CP04 = "" & adocaseprogress("CP04").Value
         End If
'         If intCounter Mod intPageTotRow = 0 Then
'            intPage = intPage + 1
'            Call PrintHead_Excel(strSubCompany, strSubNo, strSubTitle, "2")
'         End If
         
         If bolPrintItem = True Then
            intCounter = intCounter + 1: intCounterD = intCounterD + 1
            
            '收據日期
            If IsNull(adocaseprogress.Fields("a0L02").Value) = False Then
               xlsAnnuity.Range("A" & intCounter).Value = CFDate(adocaseprogress.Fields("a0L02").Value)
            Else
               xlsAnnuity.Range("A" & intCounter).Value = ""
            End If
            '收據號碼
            xlsAnnuity.Range("B" & intCounter).Value = adocaseprogress.Fields("a0k01").Value
            '國別
            If IsNull(adocaseprogress.Fields("na03").Value) = False Then
               xlsAnnuity.Range("C" & intCounter).Value = MidB(adocaseprogress.Fields("na03").Value, 1, 10)
            Else
               xlsAnnuity.Range("C" & intCounter).Value = ""
            End If
            '案件性質
            If adocaseprogress.Fields("a0k33") = "Y" Then
               xlsAnnuity.Range("D" & intCounter).Value = MidB(adocaseprogress.Fields("a0j22").Value, 1, 12)
            Else
               If adocaseprogress.Fields("a0j04").Value = "000" Then
                  xlsAnnuity.Range("D" & intCounter).Value = MidB(adocaseprogress.Fields("cpm03").Value, 1, 12)
               Else
                  xlsAnnuity.Range("D" & intCounter).Value = MidB(adocaseprogress.Fields("cpm04").Value, 1, 12)
               End If
            End If
            '案件名稱
            strName = MidB(CaseNameShow(adocaseprogress.Fields("cp01").Value, adocaseprogress.Fields("cp02").Value, adocaseprogress.Fields("cp03").Value, adocaseprogress.Fields("cp04").Value, 1), 1, 24)
            If strName = "" Then
               strName = MidB(CaseNameShow(adocaseprogress.Fields("cp01").Value, adocaseprogress.Fields("cp02").Value, adocaseprogress.Fields("cp03").Value, adocaseprogress.Fields("cp04").Value, 2), 1, 24)
            End If
            If strName = "" Then
               strName = MidB(CaseNameShow(adocaseprogress.Fields("cp01").Value, adocaseprogress.Fields("cp02").Value, adocaseprogress.Fields("cp03").Value, adocaseprogress.Fields("cp04").Value, 3), 1, 24)
            End If
            If m_CP01 = "LA" Then strName = "" 'Add By Sindy 2020/4/27 法律所的顧問案件,案件名稱欄放空白
            xlsAnnuity.Range("E" & intCounter).Value = strName
         End If
         
         dblItemAmt1 = 0
         'If Val("" & adocaseprogress("Amt")) > 0 Then
            dblItemAmt3 = dblItemAmt3 + Val("" & adocaseprogress("Amt"))
         'End If
         
         stLstItem = "" & adocaseprogress.Fields("a0j22")
         stLstRecNo = "" & adocaseprogress.Fields("a0k01")
         
         adocaseprogress.MoveNext
         
         If adocaseprogress.EOF Then
            bolPrintItemAmt = True
         Else
            If adocaseprogress.Fields("a0k33") = "Y" And stLstItem = adocaseprogress.Fields("a0j22") And stLstRecNo = adocaseprogress.Fields("a0k01") Then
               bolPrintItem = False
               bolPrintItemAmt = False
            Else
               bolPrintItem = True
               bolPrintItemAmt = True
            End If
         End If
         If bolPrintItemAmt = True Then
            '應收金額
            strAmount1 = Format(dblItemAmt1, DDollar2)
            xlsAnnuity.Range("F" & intCounter).Value = strAmount1
            
            '已收金額
            strAmount3 = Format(dblItemAmt3, DDollar2)
            xlsAnnuity.Range("G" & intCounter).Value = strAmount3
            
            lngAmount1 = lngAmount1 + dblItemAmt1
            lngAmount3 = lngAmount3 + dblItemAmt3
            dblItemAmt1 = 0
            dblItemAmt3 = 0
         End If
      Loop
   End If
   If lngAmount3 <> 0 Then
      lngAmount1 = lngAmount3 '為了讓往來結餘為0
      Call PrintSum_Excel(True, strSubCompany)
   End If
   adocaseprogress.Close
End Sub

Private Sub Text3_Change()
   If Len(Text3) = 5 Then
      lblSalesName = StaffQuery(Text3)
   Else
      lblSalesName = MsgText(601)
   End If
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
   CloseIme 'Add by Morgan 2007/10/2
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   FormCheck = False
   'Add By Sindy 2010/12/1
   If txtType = "" Then
      MsgBox "請輸入列印別！", , MsgText(5)
      txtType.SetFocus
      Exit Function
   End If
   '2010/12/1 End
   If Text1 = MsgText(601) Or Text1 = "X" Then
      MsgBox "請輸入起始客戶編號！", , MsgText(5)
      Text1.SetFocus
      Exit Function
   End If
   If Text2 = MsgText(601) Or Text2 = "X" Then
      MsgBox "請輸入迄止客戶編號！", , MsgText(5)
      Text2.SetFocus
      Exit Function
   End If
   If MaskEdBox1.Text = MsgText(29) Then
      'Modified by Morgan 2014/5/16 欄位意義更明確--瑞婷
      'MsgBox "請輸入起始應收帳日期！", , MsgText(5)
      MsgBox "請輸入起始收據日期！", , MsgText(5)
      MaskEdBox1.SetFocus
      Exit Function
   End If
   If MaskEdBox2.Text = MsgText(29) Then
      'Modified by Morgan 2014/5/16 欄位意義更明確--瑞婷
      'MsgBox "請輸入迄止應收帳日期！", , MsgText(5)
      MsgBox "請輸入迄止收據日期！", , MsgText(5)
      MaskEdBox2.SetFocus
      Exit Function
   End If
   'Add By Sindy 2012/11/12
   If Text4 = "1" Then
      If MaskEdBox3.Text = MsgText(29) Then
         MsgBox "請輸入起始收款日期！", , MsgText(5)
         MaskEdBox3.SetFocus
         Exit Function
      End If
      If MaskEdBox4.Text = MsgText(29) Then
         MsgBox "請輸入迄止收款日期！", , MsgText(5)
         MaskEdBox4.SetFocus
         Exit Function
      End If
   End If
   '2012/11/12 End
'   If Text3 = MsgText(601) Then
'      MsgBox "請輸智權人員！", , MsgText(5)
'      Text3.SetFocus
'      Exit Function
'   End If
   If Text4 = MsgText(601) Then
      MsgBox "請輸入報表類別！", , MsgText(5)
      Text4.SetFocus
      Exit Function
   End If
   FormCheck = True
End Function

'Add By Sindy 2010/12/1
Private Sub txtType_GotFocus()
   TextInverse txtType
   CloseIme
End Sub

'Add By Sindy 2010/12/1
Private Sub txtType_KeyPress(KeyAscii As Integer)
   If KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub
