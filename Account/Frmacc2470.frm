VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc2470 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "FC催款單"
   ClientHeight    =   5676
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   5256
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5676
   ScaleWidth      =   5256
   Begin VB.TextBox txtSend 
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
      Height          =   324
      Left            =   1395
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   5232
      Width           =   1300
   End
   Begin VB.CheckBox Check2 
      Caption         =   "每月"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3672
      TabIndex        =   11
      Top             =   3180
      Width           =   900
   End
   Begin VB.CheckBox Check1 
      Caption         =   "CSV"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4250
      TabIndex        =   39
      Top             =   120
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.TextBox Text10 
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
      Height          =   324
      Left            =   1335
      MaxLength       =   9
      TabIndex        =   2
      Top             =   430
      Width           =   1300
   End
   Begin VB.TextBox Text11 
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
      Height          =   324
      Left            =   2892
      MaxLength       =   9
      TabIndex        =   3
      Top             =   430
      Width           =   1300
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
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
      Left            =   150
      Style           =   1  '圖片外觀
      TabIndex        =   34
      Top             =   2700
      Width           =   3300
   End
   Begin VB.CommandButton Cmd_Dizhang 
      BackColor       =   &H00C0FFC0&
      Caption         =   "帳款處理中查詢"
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
      Left            =   1230
      Style           =   1  '圖片外觀
      TabIndex        =   33
      Top             =   2700
      Visible         =   0   'False
      Width           =   2230
   End
   Begin VB.CommandButton Cmd_Excel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生Excel"
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
      Left            =   3510
      Style           =   1  '圖片外觀
      TabIndex        =   32
      Top             =   2700
      Visible         =   0   'False
      Width           =   1500
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
      Height          =   315
      Left            =   1815
      MaxLength       =   1
      TabIndex        =   12
      Top             =   3510
      Width           =   450
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
      Height          =   300
      Left            =   1815
      MaxLength       =   1
      TabIndex        =   10
      Top             =   3150
      Width           =   450
   End
   Begin VB.TextBox Text7 
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
      Height          =   315
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   9
      Top             =   1930
      Width           =   450
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   4584
      Top             =   696
      _ExtentX        =   974
      _ExtentY        =   974
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   4632
      Top             =   1896
      _ExtentX        =   974
      _ExtentY        =   974
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin VB.TextBox txtReceiver 
      Height          =   285
      Left            =   1395
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4620
      Width           =   3480
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   495
      Left            =   3420
      ScaleHeight     =   444
      ScaleWidth      =   864
      TabIndex        =   23
      Top             =   4296
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox Text6 
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
      Height          =   315
      Left            =   1815
      MaxLength       =   1
      TabIndex        =   13
      Top             =   3915
      Width           =   450
   End
   Begin VB.TextBox Text3 
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
      Height          =   315
      Left            =   1335
      TabIndex        =   6
      Top             =   1145
      Width           =   852
   End
   Begin VB.TextBox Text4 
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
      Height          =   315
      Left            =   2490
      TabIndex        =   7
      Top             =   1145
      Width           =   852
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1335
      TabIndex        =   8
      Top             =   1535
      Width           =   3495
   End
   Begin VB.TextBox Text2 
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
      Height          =   324
      Left            =   2892
      MaxLength       =   9
      TabIndex        =   1
      Top             =   60
      Width           =   1300
   End
   Begin VB.TextBox Text1 
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
      Height          =   324
      Left            =   1335
      MaxLength       =   9
      TabIndex        =   0
      Top             =   60
      Width           =   1300
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1335
      TabIndex        =   4
      Top             =   805
      Width           =   1300
      _ExtentX        =   2307
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
      Left            =   2892
      TabIndex        =   5
      Top             =   804
      Width           =   1300
      _ExtentX        =   2286
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
   Begin VB.Label lblSend2 
      BackStyle       =   0  '透明
      Caption         =   "今日Email筆數："
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
      Left            =   2808
      TabIndex        =   42
      Top             =   5280
      Width           =   2376
   End
   Begin VB.Label lblSend 
      BackStyle       =   0  '透明
      Caption         =   "寄送對象："
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
      Left            =   240
      TabIndex        =   40
      Top             =   5280
      Width           =   1152
   End
   Begin VB.Label Label19 
      BackStyle       =   0  '透明
      Caption         =   "EMail附件為PDF檔"
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
      Height          =   345
      Left            =   3210
      TabIndex        =   38
      Top             =   3570
      Width           =   2115
   End
   Begin VB.Label Lbl_Inf 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H008080FF&
      BorderStyle     =   1  '單線固定
      Caption         =   "？"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   4248
      TabIndex        =   37
      Top             =   432
      Width           =   252
   End
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "客戶編號："
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
      Left            =   195
      TabIndex        =   36
      Top             =   430
      Width           =   1155
   End
   Begin VB.Label Label17 
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
      Left            =   2700
      TabIndex        =   35
      Top             =   432
      Width           =   252
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "國外部人員操作時將會產生催款單電子檔且自動彈郵件視窗供編輯並加該電子檔為副件。"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   180
      TabIndex        =   31
      Top             =   2280
      Width           =   4815
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "輸 1 或 2 時將不剔除有設定不寄催款單者"
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
      Height          =   435
      Left            =   240
      TabIndex        =   30
      Top             =   4320
      Width           =   3975
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "是否不發Mail：         ( Y:是 )"
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
      Left            =   180
      TabIndex        =   29
      Top             =   3570
      Width           =   3000
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "是否重印紙本：         ( Y:是 )"
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
      Left            =   180
      TabIndex        =   28
      Top             =   3180
      Width           =   3000
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "注意 : 若不存電子檔時會發Mail"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   240
      TabIndex        =   27
      Top             =   4980
      Width           =   3330
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "是否產生請款單電子檔：         ( Y:是 )"
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
      Left            =   195
      TabIndex        =   26
      Top             =   1930
      Width           =   3900
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "收件人："
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
      Left            =   240
      TabIndex        =   25
      Top             =   4620
      Width           =   1125
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "是否存電子檔：         ( Y:是 )"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   180
      TabIndex        =   24
      Top             =   3960
      Width           =   3012
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "大陸一定要輸！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   3390
      TabIndex        =   22
      Top             =   1195
      Width           =   1785
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "國籍："
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
      Left            =   195
      TabIndex        =   21
      Top             =   1145
      Width           =   975
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
      Left            =   2250
      TabIndex        =   20
      Top             =   1145
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
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
      Left            =   195
      TabIndex        =   19
      Top             =   1550
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2700
      TabIndex        =   18
      Top             =   804
      Width           =   252
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "請款日期："
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
      Left            =   195
      TabIndex        =   17
      Top             =   805
      Width           =   1125
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
      Height          =   252
      Left            =   2700
      TabIndex        =   16
      Top             =   60
      Width           =   252
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "請款對象："
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
      Left            =   195
      TabIndex        =   15
      Top             =   60
      Width           =   1155
   End
End
Attribute VB_Name = "Frmacc2470"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/2 日期欄已修改
'Modify by Morgan 2009/2/20 +A4格式及e化功能
Option Explicit

Public adoacc1k0 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Dim strSql As String, strNo As String, strAmount As String, strCurr As String
Dim intLength As Integer, intCounter As Integer, intDays As Integer
Dim douOverDue1 As Double, douOverDue2 As Double, douOverDue3 As Double, douAmount As Double
'Dim prnPrint As Printer 'Remove by Lydia 2020/10/15 改用模組PUB_RestorePrinter
Dim strPrinter As String  '原本預設印表機
'Add by Morgan 2006/11/28
Dim bolChina As Boolean '是否大陸
Dim iRowH As Integer '列高
Dim iXo As Integer, iYo As Integer 'X, Y 起始座標(不可列印區)
Dim lngX As Long, lngY As Long
Dim iPageRows As Integer '每頁筆數
Dim douTAmount As Double '應收金額
Dim douRAmount As Double '已收金額
'Add by Morgan 2009/3/27

Dim bolExcel As Boolean  'Add by Amy 2013/06/10 是否產生Excel
Dim bolA4 As Boolean '是否A4格式
Dim bol2Jpg As Boolean '是否產生電子檔 'Memo by Lydia 2016/09/08 存JPG檔  'Memo by Lydia 2020/04/23 從bol2File=>bol2Jpg
Dim bol2Pdf As Boolean 'Added by Lydia 2016/09/08 電子檔存PDF檔  'Memo by Lydia 2020/04/23 從bol2File2=>bol2Pdf
Dim bol2Printer As Boolean '是否列印
Dim bolJpg2Pdf As Boolean 'Added by Morgan 2024/9/6

Dim bolEmail As Boolean '是否發EMail
Dim strEMailBox As String 'EMail Box
Dim strPicLetter As String '暫存圖檔路徑 'Memo by Lydia 2020/04/23 信頭/尾的暫存圖: 從strPicFileName=>strPicLetter
Dim strPicFileNames As String '暫存圖檔路徑組(*號分隔)
Dim iPageNo As Integer '頁數
Dim strSavePath As String '電子檔存放路徑
Dim bolChinese As Boolean '本文是否印中文
Dim douExtRate As Double '字型位置縮放比
Dim Px(8) As Long 'X 座標
Dim Py(5) As Long 'Y 座標
Dim lngXi As Long, lngYi As Long '圖檔列印位置
Const LeftMG As Integer = 80 '字左邊空白
Const TopMG As Integer = 45 '字上邊空白
Const TwPerCm As Integer = 567 '每一公分印表機點數
Dim strData As String
Dim strMailFailList() As String 'Mail 失敗清單
Dim iXfix As Integer, iYfix As Integer '起始位置修正值
Dim strAttention As String
Dim strReceiver As String '實際收件信箱
Dim strEmailCC As String 'Added by Lydia 2024/09/18 財務副本信箱
Dim bolPromoter As Boolean '是否為國外部承辦人員
Dim m_DNRate As Double     '2009/6/4 add by sonia
Dim m_DNCurr As String     '請款單請款幣別 2009/6/4 add by sonia

'Add by Morgan 2009/7/7
Dim m_iDocCount As Integer '催款單份數統計
Dim m_iMailCount As Integer '郵寄份數統計
Dim m_iPrintCount As Integer '列印份數統計
Dim m_FNo As String '代理人編號

'Added by Lydia 2015/10/19
Dim m_PrevForm  As Form '前一畫面
Dim bolCallMail As Boolean '外部呼叫,發outlook mail
Public currAmount As String '帶出收款幣別及金額
Public strLDate As String '收款日期
Dim strDBno As String

'Added by Morgan 2016/2/15
Dim bolIsBatchInvoice As Boolean '是否為整批列印的請款單
Dim strBatchInvoiceNo As String '整批列印請款單號
Dim strBatchInvoiceStartNo As String '整批列印請款單起始號
Dim strBatchNoList As String '整批列印請款單清單
Dim strBatchNoRecList As String '整批列印請款單收款清單
'end 2016/2/15

Dim strA1k01List As String 'Added by Lydia 2016/09/08 請款單列印單號
Dim strA1k01Dir As String 'Added by Lydia 2017/03/02 請款單存放的資料夾路徑
Dim strDefDir As String 'Added by Lydia 2016/09/10 預設存檔目的地來源,可能有不同代理人
Public strCallCase As String  'Added by Lydia 2016/12/22 傳入本所案號，指定催款單範圍(T收款寄證1728)
Dim strNowDoc As String  'Added by Lydia 2020/02/15 記錄現在列印的.pdf

'Add by Amy 2017/02/02
Dim bolShowCus As Boolean, i As Integer
Dim strF, intWidth 'Modify by Amy 2020/11/19 原:strF()

Public m_SavePath As String   'Added by Lydia 2017/02/18 指定存檔目的地來源
Dim intTitleRow As Integer 'Add by Amy 2017/08/07
Dim strOldCus As String, strNowCus As String 'Add by Amy 2017/08/15
Dim m_strErr2480 As String 'Added by Lydia 2020/09/10 請款單：判斷PDF檔案是否存在
Dim strTmp(7) As String, intNo As Integer  'Add by Amy 2022/06/14 從ExcelSaveNew搬過來
Const CSV特殊格式代理人 = "Y53715;" 'Add by Amy 2024/08/02 產生CSV 相關
'Added by Lydia 2024/12/31 Excel列印
Dim xRows As Integer, xRowE As Integer 'Excel列印表格的起始,終止位置
Dim xCols As Integer, xColE As Integer 'Excel列印表格的起始/終止欄位Ascii值
Dim nRow As Integer '目前資料列位置
Dim maxRows As Integer '頁面最大列數
Dim bolColTitle As Boolean '是否列印欄位抬頭
Dim m_iNo As Integer, m_iNo2 As Integer '圖檔編號
Dim strPrtPath As String, strPrtFile As String '列印Excel檔案路徑,名稱
Dim xlsRpt As New Excel.Application
Dim WksRpt1 As New Worksheet
Dim oShape
Dim oShape2
Dim m_FrmName As String 'Added by Lydia 2025/03/11 為了區隔記錄，以檔案名稱+系統時間為主鍵

Private Sub SetPx()
   If bolChina Then
      Px(0) = 0.4 * TwPerCm
      Px(1) = Px(0) + 2 * TwPerCm '3
      Px(2) = Px(1) + 2.5 * TwPerCm '3
      Px(3) = Px(2) + 2.5 * TwPerCm '3.3
      Px(4) = Px(3) + 3 * TwPerCm '3.3
      Px(5) = Px(4) + 2.6 * TwPerCm '2.8
      Px(6) = Px(5) + 3 * TwPerCm '2.8
      Px(7) = Px(6) + 2.5 * TwPerCm '2.8
      Px(8) = Px(7) 'Add By Sindy 2013/1/7
   Else
      Px(0) = 0.4 * TwPerCm
      Px(1) = Px(0) + 1.8 * TwPerCm
      'Modified by Morgan2016/2/6
      'Px(2) = Px(1) + 2.4 * TwPerCm
      Px(2) = Px(1) + 2.8 * TwPerCm
      Px(3) = Px(2) + 3.2 * TwPerCm
      'Modified by Morgan2016/2/6
      'Px(4) = Px(3) + 5 * TwPerCm
      Px(4) = Px(3) + 4.6 * TwPerCm
      Px(5) = Px(4) + 2.2 * TwPerCm
      Px(6) = Px(5) + 2.2 * TwPerCm
      Px(7) = Px(6) + 2.2 * TwPerCm
   End If
End Sub

Private Sub SetPy()
   If bolChina Then
      Py(0) = 9.7 * TwPerCm
      Py(1) = Py(0) + 1 * TwPerCm
      Py(2) = Py(1) + 9 * TwPerCm
      Py(3) = Py(2) + 0.7 * TwPerCm
      Py(4) = Py(3) + 0.7 * TwPerCm
      Py(5) = Py(4) + 0.7 * TwPerCm 'Add By Sindy 2013/1/7
   Else
      Py(0) = 9.7 * TwPerCm
      Py(1) = Py(0) + TwPerCm
      Py(2) = Py(1) + 11 * TwPerCm
      Py(3) = Py(2) + 1.8 * TwPerCm
      Py(4) = Py(3) + 3.5 * TwPerCm
   End If
End Sub

'Add by Amy 2024/08/02
Private Sub Check1_Click()
   Cmd_Excel.Caption = "產生Excel"
   If Check1.Value = 1 Then
      Cmd_Excel.Caption = "產生CSV"
   End If
End Sub

Private Sub Cmd_Dizhang_Click()
    Frmacc2471.Show
    Me.Hide
End Sub

'Add by Amy 2013/06/10 產生Excel
Private Sub Cmd_Excel_Click()
   'Modify by Amy 2024/08/02 產生CSV共用此按鈕(只有電腦中心及財務 可用)
   If Check1.Value = True Then
      If InStr(CSV特殊格式代理人, Left(Text1, 6)) = 0 Then
         MsgBox "此代理人無CSV格式", vbExclamation
         Exit Sub
      End If
   Else
      bolExcel = True
   End If
   
   If FormCheck = False Then
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   PrintData
   Screen.MousePointer = vbDefault
   bolExcel = False
End Sub

'Modified by Lydia 2015/10/19
'Private Sub Command2_Click()
Public Sub Command2_Click()
   Command2.Enabled = False
   If FormCheck = False Then
      Command2.Enabled = True
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   PUB_SetOsDefaultPrinter Combo1 'Added by Lydia 2025/03/07 切換Word/Excel印表機
   'Modified by Lydia 2020/10/15  改用模組
   'For Each prnPrint In Printers
   '   If prnPrint.DeviceName = Combo1 Then
   '      Set Printer = prnPrint
   '   End If
   'Next
   PUB_RestorePrinter Combo1
   'end 2020/10/15
   
   'Modify by Morgan 2010/10/14 全部改用新格式
   bolA4 = True
   
   strA1k01List = "" 'Added by Lydia 2016/09/08
   strA1k01Dir = "" 'Added by Lydia 2017/03/02
   m_strErr2480 = "" 'Added by Lydia 2020/09/10
     
   'Added by Lydia 2017/12/08 變數清空
   bol2Jpg = False
   bol2Pdf = False
   bol2Printer = False
   bolEmail = False
   bolJpg2Pdf = False 'Added by Morgan 2024/9/6
   'end 2017/12/08
   'Modified by Lydia 2024/12/31 改用EXCEL
   'PrintData
   PrintExcelMain
   If strCon10 <> MsgText(602) Then
      'FormClear
   End If
   
   PUB_SetOsDefaultPrinter strPrinter  'Added by Lydia 2025/03/07 切換Word/Excel印表機
   'Modified by Lydia 2020/10/15 改用模組
   'For Each prnPrint In Printers
   '   If prnPrint.DeviceName = strPrinter Then
   '      Set Printer = prnPrint
   '   End If
   'Next
   PUB_RestorePrinter strPrinter
   'end 2020/10/15
   
   Screen.MousePointer = vbDefault
   Command2.Enabled = True
   StatusView MsgText(100)
   'Added by Lydia 2025/03/11
   If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M31" Then
      txtSend.Visible = False: lblSend.Visible = False
      GetEmailCount
   End If
End Sub

Private Sub Form_Activate()
   '93.3.16 ADD BY SONIA
   'Modify by Morgan 2006/3/27
   'If IsObject(mdiMain) Then
   '   mdiMain.ToolShow
   If Forms(0).Name = "mdiMain" Then
      Forms(0).ToolShow
   End If
   '93.3.16 END
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      StatusView MsgText(100)
   End If
End Sub

Private Sub Form_Load()
   
   'Added by Lydia 2025/03/11
   txtSend.Visible = False: lblSend.Visible = False

   '表單初始化
   'Add by Amy 2017/02/02 操作人員部門st03為財務處M31或電腦中心M51時增加下客戶編號條件-婉莘
   'Mark by Amy 2018/05/23 開放所有人下客戶編號條件-婉莘
'   Text10.Enabled = False
'   Text11.Enabled = False
'   Label18.Enabled = False
   'Add by Amy 2013/06/10 操作人員部門st03為財務處M31或電腦中心M51時增加產生Excel 婧瑄
   If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M31" Then
      'Modify by Amy 2013/11/01 +帳款處理中查詢 按鈕 (顯示字體 原14)
      Command2.Width = 1000 '原 2230
      Command2.Left = 120
      Cmd_Dizhang.Visible = True
      'strFormName = Name
      'end 2013/10/31
      'Add by Amy 2017/02/02
      Text10.Enabled = True
      Text11.Enabled = True
      Label18.Enabled = True
      'end 2017/02/02
      Cmd_Excel.Visible = True
      Cmd_Excel.Enabled = False
      bolExcel = False '預設不產生Excel
      'Add by Amy 2024/08/02 +勾選CSV (特定代理人)
      Check1.Visible = True
      Check1.Enabled = False
      'end 2024/08/02
      Check2.Visible = True 'Added by Lydia 2024/10/28 每月催款
      PUB_InitForm Me, Me.Width, Me.Height, strBackPicPath4
      GetEmailCount 'Added by Lydia 2025/03/11
   'Add by Amy 2018/05/23  其他部門只能看到「列印」鈕上方-秀玲
   Else
      'Added by Morgan 2020/10/20 配合 Y54570 Amkor Technology需求，財務處同意開放Excel功能--婧瑄,David
      Cmd_Excel.Visible = True
      Cmd_Excel.Enabled = False
      'end 2020/10/20
      Check2.Visible = False 'Added by Lydia 2024/10/28 每月催款
      PUB_InitForm Me, Me.Width, 3510, strBackPicPath4
   End If
   
   '國外部承辦組只能產生電子檔
   'Added by Lydia 2015/10/19 + bolCallMail
   If Left(Pub_StrUserSt03, 1) = "F" Or bolCallMail Then
      'Mark by Amy 2018/05/23 往上搬
      'PUB_InitForm Me, Me.Width, 3600, strBackPicPath4 'Modify by Amy 2017/02/02 原:3400(增加客戶編號條件)
      bolPromoter = True
      'Modified by Lydia 2016/09/08 改可選擇JPG 或PDF
      'Text6 = "Y"
      '只有frmacc2110 才預設jpg-秀玲
      'Remove by Lydia 2018/09/03 預設為PDF
      'If bolCallMail Then
      '  Text6 = "1"
      'Else
        'Modified by Lydia 2024/12/31 只存PDF檔
        'Text6 = "2"
        Text6 = "Y"
      'End If
      'end 2018/09/03
      Text6.Enabled = False
      txtReceiver.Enabled = False
   Else
      'PUB_InitForm Me, Me.Width, Me.Height, strBackPicPath4  'Mark by Amy 2018/05/23 往上搬
      'Modify by Amy 2018/06/04 財務處不預設,非財務處不可寄信
      If Not (Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M31") Then
        'Modified by Lydia 2024/12/31 只存PDF檔
        'Text6 = "2" 'Add by Amy 2018/05/23 預設PDF-秀玲
        Text6 = "Y"
        Text6.Enabled = False
        txtReceiver.Enabled = False
      End If
      'end 2018/06/04
      bolPromoter = False
   End If
   
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim ii As Integer

   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat

   'Modified by Morgan 2021/2/5 +只顯示有效的印表機設 True
   'Modified by Lydia 2022/09/05 排除特定印表機; 111/9/2整批寄發FC催款單遇到Y54047為紙本並且預設印表機為PDFCreator，因為在印紙本要輸入檔名未輸入，但是同時持續轉PDF作業，所以造成Y54047的列印轉成Y54067的附件。
   'PUB_SetPrinter Me.Name, Combo1, strPrinter, , , , , True 'Modified by Morgan 2017/11/8 設定印表機改呼叫公用函數,原程式移除
   PUB_SetPrinter Me.Name, Combo1, strPrinter, , , , , True, "PDFCreator"
   
   StatusView MsgText(100)
   
   'Added by Lydia 2019/08/12下載郵件範本
   If bolCallMail = True Then
       Call PUB_GetSampleFile("$$TOT-000M31-0-02.oft", "TOT-000M31-0-02") '英文
       Call PUB_GetSampleFile("$$TOT-000M31-0-03.oft", "TOT-000M31-0-03") '中文
   End If
   
   'Added by Lydia 2024/12/31
   strPrtPath = App.path & "\" & strUserNum
   Call Pub_ChkExcelPath(strPrtPath)
   Call PUB_KillTempFile(strUserNum & "\$*.*")
   'end 2024/12/31
End Sub
'Add by Amy 2013/11/01
Private Sub Form_Resize()
    tool3_enabled
   strFormName = Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
   '刪除舊的暫存圖檔
   strExc(1) = App.path & "\X*.jpg"
   If Dir(strExc(1)) <> "" Then Kill strExc(1)
   strExc(1) = App.path & "\Y*.jpg"
   If Dir(strExc(1)) <> "" Then Kill strExc(1)
   'Added by Lydia 2019/06/10 刪除舊的PDF
   strExc(1) = App.path & "\X*.pdf"
   If Dir(strExc(1)) <> "" Then Kill strExc(1)
   strExc(1) = App.path & "\Y*.pdf"
   If Dir(strExc(1)) <> "" Then Kill strExc(1)
   
   'Printer.DrawWidth = 1  '2010/12/1 add by sonia 'Removed by Morgan 2020/4/9
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   
   Set Frmacc2470 = Nothing
End Sub

'Add by Amy 2017/09/12 說明-莘
Private Sub Lbl_Inf_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Lbl_Inf.ToolTipText = "查詢條件說明：" & _
                        "所有申請人各自的催款單: X00001~X99999／" & _
                        "無申請人的催款單: 空白~XZZZZZ"
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus()
   'Modify by Amy 2013/11/01 +Pub_StrUserSt03 判斷及text1 有值
   'Modified by Morgan 2020/10/20 不再限制部門
   'If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M31" Then
      'Add by Amy 2013/06/10 請款對象條件起迄有值且前6碼相同 顯示產生Excel鈕
      'Modify by Amy 2024/08/02  電腦中心及財務 +勾選CSV
      Check1.Enabled = False
      Check1.Value = 0
      Cmd_Excel.Caption = "產生Excel"
      If Len(Trim(Text1)) > 0 And Left(Text2, 6) = Left(Text1, 6) Then
         Cmd_Excel.Enabled = True
         If (Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M31") Then
            If InStr(CSV特殊格式代理人, Left(Text1, 6)) > 0 Then
               Check1.Value = 1
               Check1.Enabled = True
            End If
         End If
      'end 2024/08/02
      Else
         Cmd_Excel.Enabled = False
      End If
   'End If
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Len(Text1) = 6 Then
      Text1 = AfterZero(Text1)
   End If
   '2009/6/2 ADD BY SONIA 預設尾碼999
   If Text1.Text <> "" Then
      'Modify by Morgan 2011/1/12 承辦人預設相同,因為有限制編號起迄不可不同
      'Text2.Text = Left(Me.Text1.Text, 6) & "999"
      If bolPromoter = True Then
         Text2.Text = Me.Text1.Text
         'Modify by Amy 2018/05/23 改為6碼+ZZZ -秀玲
         'Text2.Text = Left(Me.Text2.Text, 8) & "Z"    'add by sonia 2017/4/19 第9碼改為Z,因為有更名前的資料Y51562002
          Text2.Text = Left(Me.Text1.Text, 6) & "ZZZ"
      Else
         'Modify By Sindy 2014/8/11 999=>ZZZ
         'Text2.Text = Left(Me.Text1.Text, 6) & "999"
         Text2.Text = Left(Me.Text1.Text, 6) & "ZZZ"
      End If
   End If
End Sub

Private Sub Text10_GotFocus()
    TextInverse Text10
    CloseIme
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
    If Len(Text10) = 6 Then
        Text10 = AfterZero(Text10)
    End If
    If Text10 <> "" Then
        Text11 = Left(Me.Text10, 6) & "ZZZ"
    End If
End Sub

Private Sub Text11_GotFocus()
    TextInverse Text11
    CloseIme
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_GotFocus()
'Add by Amy 2013/06/10 請款對象條件起迄有值且前6碼相同 顯示產生Excel鈕
   If Text2 = "" And Text1 <> "" Then
      Text2 = Text1
   End If
   TextInverse Text2
   CloseIme
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_LostFocus()
     'Modify by Amy 2013/11/01 +Pub_StrUserSt03 判斷及text1 有值
     If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M31" Then
        If Len(Trim(Text1)) > 0 And Left(Text2, 6) = Left(Text1, 6) Then
            Cmd_Excel.Enabled = True
        Else
            Cmd_Excel.Enabled = False
        End If
    End If
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Len(Text2) = 6 Then
      Text2 = AfterZero(Text2)
      Text2.Text = Left(Text2, 8) & "Z"    'add by sonia 2017/4/19 第9碼改為Z,因為有更名前的資料Y51562002
   End If
   'add by sonia 2017/4/19 第9碼改為Z,因為有更名前的資料Y51562002
   If Mid(Text2, 9, 1) <> "Z" Then
      Text2.Text = Left(Text2, 8) & "Z"
      MsgBox ("可能會有變更名稱前的資料,故請款對象迄號的第9碼自動改為 Z！")
   End If
   'end 2017/4/19
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1 = ""
   Text2 = ""
   Text3 = ""
   Text4 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text1.SetFocus
End Sub

Private Function ChgDateFormat(pDate As String) As String
   Dim iYear As Integer, iMonth As Integer, iDay As Integer
   
   pDate = Replace(Replace(pDate, "_", " "), "/", "")
   iYear = Val(Left(pDate, 3))
   iMonth = Val(Mid(pDate, 4, 2))
   iDay = Val(Right(pDate, 2))
   ChgDateFormat = 10000# * iYear + 100# * iMonth + iDay
End Function

'*************************************************
' 列印明細資料
'
'*************************************************
Private Sub PrintData()
Dim strDocNo As String
Dim StrSQLa As String
'2005/8/2 ADD BY SONIA
Dim strSQL1 As String
Dim strSQL2 As String
Dim StrSQL3 As String, StrSQL4 As String 'Add by Amy 2013/11/01
Dim strField As String 'Add by Amy 2017/02/02
Dim hLocalFile As Long, strXLSFile As String 'Added by Morgan 2020/12/11

   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/22 清除查詢印表記錄檔欄位
   bolShowCus = False 'Add by Amy 2017/02/02 是否下客戶編號條件
   strSQL1 = MsgText(601)
   strSQL2 = MsgText(601)
   StrSQL3 = MsgText(601) 'Add by Amy 2013/11/01
   StrSQL4 = MsgText(601) 'Add by Amy 2013/11/01
   
   If Text1 <> MsgText(601) Then
      strSQL1 = strSQL1 & " and a1k28 >= '" & Text1 & "'"
      strSQL2 = strSQL2 & " and a1k28 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) Then
      strSQL1 = strSQL1 & " and a1k28 <= '" & Text2 & "'"
      strSQL2 = strSQL2 & " and a1k28 <= '" & Text2 & "'"
   End If
   If Text1 <> MsgText(601) Or Text2 <> MsgText(601) Then
      pub_QL05 = pub_QL05 & ";" & Label1 & Text1 & "-" & Text2 'Add By Sindy 2010/12/22
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSQL1 = strSQL1 & " and a1k02 >= " & ChgDateFormat(MaskEdBox1.Text) & ""
      strSQL2 = strSQL2 & " and a1k02 >= " & ChgDateFormat(MaskEdBox1.Text) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSQL1 = strSQL1 & " and a1k02 <= " & ChgDateFormat(MaskEdBox2.Text) & ""
      strSQL2 = strSQL2 & " and a1k02 <= " & ChgDateFormat(MaskEdBox2.Text) & ""
   End If
   If (MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29)) Or _
      (MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29)) Then
      pub_QL05 = pub_QL05 & ";" & Label4 & MaskEdBox1 & "-" & MaskEdBox2 'Add By Sindy 2010/12/22
   End If
   
   'Add by Amy 2017/02/02 +客戶編號(Y27766 Excel為特殊所以不加)
   If (Text10 <> MsgText(601) Or Text11 <> MsgText(601)) And Left(Text1, 6) <> "Y27766" Then
        bolShowCus = True
        strField = "'" & strUserNum & "',"
        pub_QL05 = pub_QL05 & ";" & Label8 & Text10 & "-" & Text11
   End If

   'Add by Amy 2013/11/01
   '+if 請款對象條件為空(整批)，且有設帳款處理者，不列印該筆請款單
   If Text1 = "" And Text2 = "" Then
        StrSQL3 = " and fa103 is null"
        StrSQL4 = " and cu142 is null"
   End If
   
   ' +有輸請款對象(單筆)且帳款處理情形有值顯示訊息
    If Text1 <> "" And Text2 <> "" Then
        strExc(0) = GetDizhang(Text1, Text2, True)
    End If
    'end 2013/10/31
    
   'Modified by Lydia 2016/09/08 改可選擇JPG 或PDF
   'If Text6 <> "Y" Then 'Added by Morgan 2011/12/7 是否存電子檔輸 Y 時不剔除有設定不寄催款單者--婧瑄
   'Modified by Lydia 2024/12/31 只存PDF檔
   'If Text6 <> "1" And Text6 <> "2" Then
   If Text6 <> "Y" Then
      'Added by Lydia 2024/10/28 增加不寄催款單:1.每月催款
      If Check2.Visible = True And Check2.Value = 1 Then
           strSQL1 = strSQL1 & " and nvl(fa101,'0')='1' "
           strSQL2 = strSQL2 & " and nvl(cu140,'0')='1' "
      Else
         'Add by Amy 2013/06/26 不存電子檔且不是產生excel,抓不寄催款單 為空值者
         If bolExcel <> True Then
           strSQL1 = strSQL1 & " and fa101 is null "
           strSQL2 = strSQL2 & " and cu140 is null "
         End If
         'end 2013/06/26
      End If 'Added by Lydia 2024/10/28
      
      'Add by Morgan 2011/10/13 若請款對象起訖前6碼相同時先檢查是否有設定不寄催款單 --秀玲
      If Text1 <> "" And Left(Text1, 6) = Left(Text2, 6) Then
         strExc(0) = "select fa01||fa02 FNo from fagent where fa01 between '" & Left(Text1, 8) & "' and '" & Left(Text2, 8) & "' and fa101 is not null" & _
             " union select cu01||cu02 FNo from customer where cu01 between '" & Left(Text1, 8) & "' and '" & Left(Text2, 8) & "' and CU140 is not null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(1) = RsTemp.GetString(, , , vbCrLf)
            If MsgBox("下列請款對象有設定不寄催款單程式將自動略過，是否仍要繼續??" & vbCrLf & vbCrLf & strExc(1), vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Sub
            End If
         End If
      End If
      'end 2011/10/13
   End If
   
   strCon10 = ""
   adoacc1k0.CursorLocation = adUseClient
   '93.6.1 MODIFY BY SONIA 扣除折讓金額
   'strSQLA = "select a1k02, a1k28 as FagentNo, a1k02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, 0 as OAmount, a1k08 as FAmount, a1k04 as Yno, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa17, fa43, a1k18 as Curr, a1k10 as Rate from acc1k0, fagent where substr(a1k28, 1, 8) = fa01 and substr(a1k28, 9, 1) = fa02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k06, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0)" & strSQL & " union " & _
   '               "select a1k02, a1k28 as FagentNo, a1k02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, 0 as OAmount, a1k08 as FAmount, a1k04 as Yno, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu23 as fa17, cu76 as fa43, a1k18 as Curr, a1k10 as Rate from acc1k0, customer where substr(a1k28, 1, 8) = cu01 and substr(a1k28, 9, 1) = cu02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k06, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0)" & strSQL & " union " & _
   '               "select a1k02, a1k28 as FagentNo, a0y02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a0z04 as OAmount, 0 as FAmount, a1k04 as Yno, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa17, fa43, a0z03 as Curr, a0y04 as Rate from acc0z0, acc0y0, acc1k0, fagent where a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = fa01 and substr(a1k28, 9, 1) = fa02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k06, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0)" & strSQL & " union " & _
   '               "select a1k02, a1k28 as FagentNo, a0y02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a0z04 as OAmount, 0 as FAmount, a1k04 as Yno, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu23 as fa17, cu76 as fa43, a0z03 as Curr, a0y04 as Rate from acc0z0, acc0y0, acc1k0, customer where a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = cu01 and substr(a1k28, 9, 1) = cu02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k06, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0)" & strSQL & " union " & _
   '               "select a1k02, a1k28 as FagentNo, a1202 as DocDate, a1201 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a1207 as OAmount, 0 as FAmount, a1k04 as Yno, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa17, fa43, a1204 as Curr, a1205 as Rate from acc120, acc0z0, acc1k0, fagent where a1k01 = a0z02 and a0z01 = a1210 and substr(a1k28, 1, 8) = fa01 and substr(a1k28, 9, 1) = fa02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k06, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0)" & strSQL & " union " & _
   '               "select a1k02, a1k28 as FagentNo, a1202 as DocDate, a1201 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a1207 as OAmount, 0 as FAmount, a1k04 as Yno, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu23 as fa17, cu76 as fa43, a1204 as Curr, a1205 as Rate from acc120, acc0z0, acc1k0, customer where a1k01 = a0z02 and a0z01 = a1210 and substr(a1k28, 1, 8) = cu01 and substr(a1k28, 9, 1) = cu02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k06, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0)" & strSQL & _
   '               " order by FagentNo asc, DocDate asc, DocNo asc"
   '2005/8/2 MODIFY BY SONIA 加 FA10 OR CU10 as NATION
   'Modify by Morgan 2009/3/27 +nvl(fa79,fa16) fa16
   'Modify by Sindy 2010/7/12 過濾FA101 is null, CU140 is null
   'Modify by Sindy 2011/3/8 +FA108
   'Modified by Morgan 2011/12/7 FA101, CU140 條件移到上面
   'Modify By Sindy 2012/10/23 原程式抓A0Z03改抓A0Y03
   '2012/12/17 modify by sonia 國外暫收款ACC120要再檢查是否已沖收款N10100186於M10103124沖收款,故2012/6/22的X10102195不該出現N10100186,故加入ACC1P0的A1P23檢查
   '2013/10/3 mdofi by sonia (a1k08 - nvl(a1k06, 0)) as FAmount改為(a1k08 - nvl(a1k31, 0)) as FAmount,nvl(a1k08, 0) <> nvl(a1k06, 0))條件改為nvl(a1k08, 0) <> nvl(a1k31, 0))
   'Modify by Amy 2013/11/01 +整批(編號空白)時有設帳款處理者不印
   'Mofified by Morgan 2016/2/5 +整批列印ACC1T0相關欄位
   'StrSQLa = "select a1k02, a1k28 as FagentNo, a1k02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, 0 as OAmount, (a1k08 - nvl(a1k31, 0)) as FAmount, a1k04 as Yno, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa17, fa43, a1k18 as Curr, a1k10 as Rate, FA10 as NATION,nvl(fa79,fa16) as EBox,FA108,fa70 from acc1k0, fagent where substr(a1k28, 1, 8) = fa01 and substr(a1k28, 9, 1) = fa02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) " & strSQL1 & StrSQL3 & " union " & _
             "select a1k02, a1k28 as FagentNo, a1k02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, 0 as OAmount, (a1k08 - nvl(a1k31, 0)) as FAmount, a1k04 as Yno, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu23 as fa17, cu76 as fa43, a1k18 as Curr, a1k10 as Rate, CU10 as NATION,nvl(cu115,cu20) as EBox,cu148 as FA108,cu102 fa70 from acc1k0, customer where substr(a1k28, 1, 8) = cu01 and substr(a1k28, 9, 1) = cu02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) " & strSQL2 & StrSQL4 & " union " & _
             "select a1k02, a1k28 as FagentNo, a0y02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a0z04 as OAmount, 0 as FAmount, a1k04 as Yno, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa17, fa43, A0Y03 as Curr, a0y04 as Rate, FA10 as NATION,nvl(fa79,fa16) as EBox,FA108,fa70 from acc0z0, acc0y0, acc1k0, fagent where a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = fa01 and substr(a1k28, 9, 1) = fa02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) " & strSQL1 & StrSQL3 & " union " & _
             "select a1k02, a1k28 as FagentNo, a0y02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a0z04 as OAmount, 0 as FAmount, a1k04 as Yno, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu23 as fa17, cu76 as fa43, A0Y03 as Curr, a0y04 as Rate, CU10 as NATION,nvl(cu115,cu20) as EBox,cu148 as FA108,cu102 as fa70 from acc0z0, acc0y0, acc1k0, customer where a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = cu01 and substr(a1k28, 9, 1) = cu02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) " & strSQL2 & StrSQL4 & " union " & _
             "select a1k02, a1k28 as FagentNo, a1202 as DocDate, a1201 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a1207 as OAmount, 0 as FAmount, a1k04 as Yno, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa17, fa43, a1204 as Curr, a1205 as Rate, FA10 as NATION,nvl(fa79,fa16) as EBox,FA108,fa70 from acc120, acc0z0, acc1k0, fagent, acc1p0 where a1k01 = a0z02 and a0z01 = a1210 and substr(a1k28, 1, 8) = fa01 and substr(a1k28, 9, 1) = fa02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) and a1201=a1p23(+) and a1p04 is null " & strSQL1 & StrSQL3 & " union " & _
             "select a1k02, a1k28 as FagentNo, a1202 as DocDate, a1201 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a1207 as OAmount, 0 as FAmount, a1k04 as Yno, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu23 as fa17, cu76 as fa43, a1204 as Curr, a1205 as Rate, CU10 as NATION,nvl(cu115,cu20) as EBox,cu148 as FA108,cu102 as fa70 from acc120, acc0z0, acc1k0, customer, acc1p0 where a1k01 = a0z02 and a0z01 = a1210 and substr(a1k28, 1, 8) = cu01 and substr(a1k28, 9, 1) = cu02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) and a1201=a1p23(+) and a1p04 is null " & strSQL2 & StrSQL4
   
   'Modfiy by Amy 2017/02/02 +產生Excel寫入暫TB
   'Modify by Amy 2020/11/19 +a1k33
   'Modified by Morgan 2020/12/14 + and d.a1k25 is null(排除已銷帳請款單) Ex:X10805261-X10911362
   'Modify by Amy 2021/02/02 +a1k38
   'Modified by Lydia 2022/03/04  只抓整批請款單尚未結清的請款單號 exists(select * => and a1k01 in (select d.a1k01
   'Modified by Lydia 2024/09/18  +財務副本信箱strEmailCC; 比照frmacc2450財務信箱CF, 優先抓FA105->FA79
   'StrSQLa = "select " & strField & "a1k02, a1k28 as FagentNo, a1k02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, 0 as OAmount, (a1k08 - nvl(a1k31, 0)) as FAmount, a1k04 as Yno, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa17, fa43, a1k18 as Curr, a1k10 as Rate, FA10 as NATION,nvl(fa79,fa16) as EBox,FA108,fa70, a1k32, a1t01, a1t02,a1k33,a1k38 From acc1k0, fagent, acc1t0 where a1t01(+)=a1k01 and substr(a1k28, 1, 8) = fa01 and substr(a1k28, 9, 1) = fa02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) " & strSQL1 & StrSQL3 & " union " & _
             "select " & strField & "a1k02, a1k28 as FagentNo, a1k02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, 0 as OAmount, (a1k08 - nvl(a1k31, 0)) as FAmount, a1k04 as Yno, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu23 as fa17, cu76 as fa43, a1k18 as Curr, a1k10 as Rate, CU10 as NATION,nvl(cu115,cu20) as EBox,cu148 as FA108,cu102 fa70, a1k32, a1t01, a1t02,a1k33,a1k38 From acc1k0, customer , acc1t0 where a1t01(+)=a1k01 and  substr(a1k28, 1, 8) = cu01 and substr(a1k28, 9, 1) = cu02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) " & strSQL2 & StrSQL4 & " union " & _
             "select " & strField & "a1k02, a1k28 as FagentNo, a0y02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a0z04 as OAmount, 0 as FAmount, a1k04 as Yno, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa17, fa43, A0Y03 as Curr, a0y04 as Rate, FA10 as NATION,nvl(fa79,fa16) as EBox,FA108,fa70, a1k32, a1t01, a1t02,a1k33,a1k38 From acc0z0, acc0y0, acc1k0, fagent , acc1t0 where a1t01(+)=a1k01 and nvl(a1k32,'N')<>'C' and  a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = fa01 and substr(a1k28, 9, 1) = fa02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) " & strSQL1 & StrSQL3 & " union " & _
             "select " & strField & "a1k02, a1k28 as FagentNo, a0y02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a0z04 as OAmount, 0 as FAmount, a1k04 as Yno, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu23 as fa17, cu76 as fa43, A0Y03 as Curr, a0y04 as Rate, CU10 as NATION,nvl(cu115,cu20) as EBox,cu148 as FA108,cu102 as fa70, a1k32, a1t01, a1t02,a1k33,a1k38 From acc0z0, acc0y0, acc1k0, customer , acc1t0 where a1t01(+)=a1k01 and nvl(a1k32,'N')<>'C' and  a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = cu01 and substr(a1k28, 9, 1) = cu02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) " & strSQL2 & StrSQL4 & " union " & _
             "select " & strField & "a1k02, a1k28 as FagentNo, a0y02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a0z04 as OAmount, 0 as FAmount, a1k04 as Yno, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa17, fa43, A0Y03 as Curr, a0y04 as Rate, FA10 as NATION,nvl(fa79,fa16) as EBox,FA108,fa70, a1k32, a1t01, a1t02,a1k33,a1k38 From acc0z0, acc0y0, acc1k0 a, fagent , acc1t0 b where a1t01(+)=a1k01 and nvl(a1k32,'N')='C' and  a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = fa01 and substr(a1k28, 9, 1) = fa02 and a1k01 in (select d.a1k01 from acc1t0 c,acc1k0 d where c.a1t02=b.a1t02 and d.a1k01=c.a1t01 and nvl(d.a1k29,'N')<>'Y' and nvl(d.a1k12,0)=0 and d.a1k25 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0) and a1k25 is null and (a1k12 is null or a1k12 = 0) " & strSQL1 & StrSQL3 & " union " & _
             "select " & strField & "a1k02, a1k28 as FagentNo, a0y02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a0z04 as OAmount, 0 as FAmount, a1k04 as Yno, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu23 as fa17, cu76 as fa43, A0Y03 as Curr, a0y04 as Rate, CU10 as NATION,nvl(cu115,cu20) as EBox,cu148 as FA108,cu102 as fa70, a1k32, a1t01, a1t02,a1k33,a1k38 From acc0z0, acc0y0, acc1k0 a, customer , acc1t0 b where a1t01(+)=a1k01 and nvl(a1k32,'N')='C' and a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = cu01 and substr(a1k28, 9, 1) = cu02 and a1k01 in (select d.a1k01 from acc1t0 c,acc1k0 d where c.a1t02=b.a1t02 and d.a1k01=c.a1t01 and nvl(d.a1k29,'N')<>'Y' and nvl(d.a1k12,0)=0 and d.a1k25 is null)" & _
             " and nvl(a1k08, 0) <> nvl(a1k31, 0) and a1k25 is null and (a1k12 is null or a1k12 = 0) " & strSQL2 & StrSQL4 & " union " & _
             "select " & strField & "a1k02, a1k28 as FagentNo, a1202 as DocDate, a1201 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a1207 as OAmount, 0 as FAmount, a1k04 as Yno, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa17, fa43, a1204 as Curr, a1205 as Rate, FA10 as NATION,nvl(fa79,fa16) as EBox,FA108,fa70, a1k32, '', '',a1k33,a1k38 From acc120, acc0z0, acc1k0, fagent, acc1p0 where a1k01 = a0z02 and a0z01 = a1210 and substr(a1k28, 1, 8) = fa01 and substr(a1k28, 9, 1) = fa02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) and a1201=a1p23(+) and a1p04 is null " & strSQL1 & StrSQL3 & " union " & _
             "select " & strField & "a1k02, a1k28 as FagentNo, a1202 as DocDate, a1201 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a1207 as OAmount, 0 as FAmount, a1k04 as Yno, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu23 as fa17, cu76 as fa43, a1204 as Curr, a1205 as Rate, CU10 as NATION,nvl(cu115,cu20) as EBox,cu148 as FA108,cu102 as fa70, a1k32, '', '',a1k33,a1k38 From acc120, acc0z0, acc1k0, customer, acc1p0 where a1k01 = a0z02 and a0z01 = a1210 and substr(a1k28, 1, 8) = cu01 and substr(a1k28, 9, 1) = cu02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) and a1201=a1p23(+) and a1p04 is null " & strSQL2 & StrSQL4
   'end 2020/11/19
   '93.6.1 END
   StrSQLa = "select " & strField & "a1k02, a1k28 as FagentNo, a1k02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, 0 as OAmount, (a1k08 - nvl(a1k31, 0)) as FAmount, a1k04 as Yno, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa17, fa43, a1k18 as Curr, a1k10 as Rate, FA10 as NATION,nvl(fa79,fa16) as EBox,FA108,fa70, a1k32, a1t01, a1t02,a1k33,a1k38,decode(fa105||fa79,null,'',fa134) as emailcc From acc1k0, fagent, acc1t0 where a1t01(+)=a1k01 and substr(a1k28, 1, 8) = fa01 and substr(a1k28, 9, 1) = fa02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) " & strSQL1 & StrSQL3 & " union " & _
             "select " & strField & "a1k02, a1k28 as FagentNo, a1k02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, 0 as OAmount, (a1k08 - nvl(a1k31, 0)) as FAmount, a1k04 as Yno, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu23 as fa17, cu76 as fa43, a1k18 as Curr, a1k10 as Rate, CU10 as NATION,nvl(cu115,cu20) as EBox,cu148 as FA108,cu102 fa70, a1k32, a1t01, a1t02,a1k33,a1k38,decode(cu115,null,'',cu200) as emailcc From acc1k0, customer , acc1t0 where a1t01(+)=a1k01 and  substr(a1k28, 1, 8) = cu01 and substr(a1k28, 9, 1) = cu02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) " & strSQL2 & StrSQL4 & " union " & _
             "select " & strField & "a1k02, a1k28 as FagentNo, a0y02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a0z04 as OAmount, 0 as FAmount, a1k04 as Yno, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa17, fa43, A0Y03 as Curr, a0y04 as Rate, FA10 as NATION,nvl(fa79,fa16) as EBox,FA108,fa70, a1k32, a1t01, a1t02,a1k33,a1k38,decode(fa105||fa79,null,'',fa134) as emailcc From acc0z0, acc0y0, acc1k0, fagent , acc1t0 where a1t01(+)=a1k01 and nvl(a1k32,'N')<>'C' and  a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = fa01 and substr(a1k28, 9, 1) = fa02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) " & strSQL1 & StrSQL3 & " union " & _
             "select " & strField & "a1k02, a1k28 as FagentNo, a0y02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a0z04 as OAmount, 0 as FAmount, a1k04 as Yno, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu23 as fa17, cu76 as fa43, A0Y03 as Curr, a0y04 as Rate, CU10 as NATION,nvl(cu115,cu20) as EBox,cu148 as FA108,cu102 as fa70, a1k32, a1t01, a1t02,a1k33,a1k38,decode(cu115,null,'',cu200) as emailcc From acc0z0, acc0y0, acc1k0, customer , acc1t0 where a1t01(+)=a1k01 and nvl(a1k32,'N')<>'C' and  a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = cu01 and substr(a1k28, 9, 1) = cu02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) " & strSQL2 & StrSQL4 & " union " & _
             "select " & strField & "a1k02, a1k28 as FagentNo, a0y02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a0z04 as OAmount, 0 as FAmount, a1k04 as Yno, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa17, fa43, A0Y03 as Curr, a0y04 as Rate, FA10 as NATION,nvl(fa79,fa16) as EBox,FA108,fa70, a1k32, a1t01, a1t02,a1k33,a1k38,decode(fa105||fa79,null,'',fa134) as emailcc From acc0z0, acc0y0, acc1k0 a, fagent , acc1t0 b where a1t01(+)=a1k01 and nvl(a1k32,'N')='C' and  a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = fa01 and substr(a1k28, 9, 1) = fa02 and a1k01 in (select d.a1k01 from acc1t0 c,acc1k0 d where c.a1t02=b.a1t02 and d.a1k01=c.a1t01 and nvl(d.a1k29,'N')<>'Y' and nvl(d.a1k12,0)=0 and d.a1k25 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0) and a1k25 is null and (a1k12 is null or a1k12 = 0) " & strSQL1 & StrSQL3 & " union " & _
             "select " & strField & "a1k02, a1k28 as FagentNo, a0y02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a0z04 as OAmount, 0 as FAmount, a1k04 as Yno, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu23 as fa17, cu76 as fa43, A0Y03 as Curr, a0y04 as Rate, CU10 as NATION,nvl(cu115,cu20) as EBox,cu148 as FA108,cu102 as fa70, a1k32, a1t01, a1t02,a1k33,a1k38,decode(cu115,null,'',cu200) as emailcc From acc0z0, acc0y0, acc1k0 a, customer , acc1t0 b where a1t01(+)=a1k01 and nvl(a1k32,'N')='C' and a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = cu01 and substr(a1k28, 9, 1) = cu02 and a1k01 in (select d.a1k01 from acc1t0 c,acc1k0 d where c.a1t02=b.a1t02 and d.a1k01=c.a1t01 and nvl(d.a1k29,'N')<>'Y' and nvl(d.a1k12,0)=0 and d.a1k25 is null)" & _
             " and nvl(a1k08, 0) <> nvl(a1k31, 0) and a1k25 is null and (a1k12 is null or a1k12 = 0) " & strSQL2 & StrSQL4 & " union " & _
             "select " & strField & "a1k02, a1k28 as FagentNo, a1202 as DocDate, a1201 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a1207 as OAmount, 0 as FAmount, a1k04 as Yno, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa17, fa43, a1204 as Curr, a1205 as Rate, FA10 as NATION,nvl(fa79,fa16) as EBox,FA108,fa70, a1k32, '', '',a1k33,a1k38,decode(fa105||fa79,null,'',fa134) as emailcc From acc120, acc0z0, acc1k0, fagent, acc1p0 where a1k01 = a0z02 and a0z01 = a1210 and substr(a1k28, 1, 8) = fa01 and substr(a1k28, 9, 1) = fa02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) and a1201=a1p23(+) and a1p04 is null " & strSQL1 & StrSQL3 & " union " & _
             "select " & strField & "a1k02, a1k28 as FagentNo, a1202 as DocDate, a1201 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a1207 as OAmount, 0 as FAmount, a1k04 as Yno, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu23 as fa17, cu76 as fa43, a1204 as Curr, a1205 as Rate, CU10 as NATION,nvl(cu115,cu20) as EBox,cu148 as FA108,cu102 as fa70, a1k32, '', '',a1k33,a1k38,decode(cu115,null,'',cu200) as emailcc From acc120, acc0z0, acc1k0, customer, acc1p0 where a1k01 = a0z02 and a0z01 = a1210 and substr(a1k28, 1, 8) = cu01 and substr(a1k28, 9, 1) = cu02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) and a1201=a1p23(+) and a1p04 is null " & strSQL2 & StrSQL4
             
   'Add by Morgan 2006/11/28 加控制, NATION欄位要加''過濾,否則會很慢
   If Text3 = "020" Then
      bolChina = True
      StrSQLa = "select * from (" & StrSQLa & ") X where NATION||''='020'"
      'Added by Lydia 2024/10/24 代理人Y55822來信，對帳單須以全英文顯示；帳單將改建在Y55822020
      If Left(Text1, 6) = "Y55822" And Left(Text2, 6) = "Y55822" Then     'Y55822000的FA101=3不會產生催款單
         bolChina = False
      End If
      'end 2024/10/24
   Else
      bolChina = False
      StrSQLa = "select * from (" & StrSQLa & ") X where NATION||''<>'020' "
   End If
   
   'Added by Lydia 2016/12/22 傳入本所案號，指定催款單範圍(T收款寄證1728)
   If strCallCase <> "" Then
      StrSQLa = StrSQLa & " and CaseNo1||CaseNo2||CaseNo3||CaseNo4='" & strCallCase & "'"
   End If
   
   If Text3 <> MsgText(601) Or Text4 <> MsgText(601) Then
      pub_QL05 = pub_QL05 & ";" & Label6 & Text3 & "-" & Text4 'Add By Sindy 2010/12/22
   End If
   If Text7 = "Y" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label12, 11) & Text7 'Add By Sindy 2010/12/22
   End If
   If Text8 = "Y" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label14, 7) & Text8 'Add By Sindy 2010/12/22
   End If
   If Text9 = "Y" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label15, 9) & Text9 'Add By Sindy 2010/12/22
   End If
   'Modified by Lydia 2016/09/08 改可選擇JPG 或PDF
   'If Text6 = "Y" Then
   '   pub_QL05 = pub_QL05 & ";" & Left(Label9, 7) & Text6 'Add By Sindy 2010/12/22
   'End If
   'Modified by Lydia 2024/12/31 只存PDF檔
   'If Text6 = "1" Then
   '   pub_QL05 = pub_QL05 & ";" & Left(Label9, 7) & "JPG檔"
   'End If
   'If Text6 = "2" Then
   If Text6 = "Y" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label9, 7) & "PDF檔"
   End If
   'end 2016/09/08
   
   If txtReceiver <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label11 & txtReceiver 'Add By Sindy 2010/12/22
   End If
   
   '2009/6/3 MODIFY BY SONIA 請款幣別不同也要跳頁
   'StrSQLa = StrSQLa & " order by FagentNo asc, DocDate asc, DocNo asc"
   'Modify by Amy 2017/02/02 +if 客戶編號
   'Modify by Amy 2017/08/15 紙本列印有下客戶編號且非Y27766則檔名改為 Y編號+X編號+幣別,故寫入暫存檔
   If bolShowCus = True And Left(Text1, 6) <> "Y27766" Then
        cnnConnection.Execute "Delete Accrpt2470 Where ID='" & strUserNum & "' "
        'Modify by Amy 2020/11/19 +R041 列印幣別格式(a1k33)
        'Modify by Amy 2021/02/02 +R042 美金請款金額(a1k38)
        'Modified by Lydia 2024/09/18 +R043財務副本信箱(emailcc)
        StrSQLa = "Insert Into Accrpt2470 (ID,R001,R002,R003,R004,R005,R006,R007,R008,R009,R010" & _
                        ",R011,R012,R013,R014,R015,R016,R017,R018,R019,R020,R021,R022,R023,R024,R025,R026,R027,R028,R029,R030" & _
                        ",R031,R032,R033,R034,R035,R036,R037,R038,R039,R041,R042,R043) " & StrSQLa
        cnnConnection.Execute StrSQLa
        Call UpdCusData
        'Modify by Amy 2017/07/31 若只下客戶編號迄號為XZZZZZ 則抓此代理人無客戶編號資料 ex:Y51817 1060630 TS-001452
        If Text10 = MsgText(601) And Left(Text11, 6) = "XZZZZZ" Then
            cnnConnection.Execute "Delete Accrpt2470 Where ID='" & strUserNum & "' And R040 is not null"
        Else
            cnnConnection.Execute "Delete Accrpt2470 Where ID='" & strUserNum & "' And R040 is null"
        End If
        'Modify by Amy 2017/08/07 中文格式:中->英->日 /英文格式:英->日->中
        If bolChina = True Then
            StrSQLa = ",NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) as CusName "
        Else
            StrSQLa = ",NVL(CU05||CU88||CU89||CU90,Nvl(CU06,CU04)) as CusName "
        End If
        'Modify by Amy 2020/11/19 +R041 列印幣別格式(a1k33)
        'Modify by Amy 2021/02/02 +R042 美金請款金額(a1k38)
        'Modified by Lydia 2024/12/04 +R043財務副本信箱(emailcc)
        StrSQLa = "Select R001 as a1k02, R002 as FagentNo, R003 as DocDate, R004 as DocNo, R005 as CaseNo1, R006 as CaseNo2, R007 as CaseNo3, R008 as CaseNo4, " & _
                        "R009 as OAmount, R010 as FAmount, R011 as Yno, R012 as  fa05, R013 as fa63, R014 as fa64, R015 as fa65, R016 as fa32, R017 as fa18, R018 as fa33,R019 as fa19, R020 as fa34, " & _
                        "R021 as fa20, R022 as fa35, R023 as fa21, R024 as fa36, R025 as fa22, R026 as fa06, R027 as fa23, R028 as fa04, R029 as fa17, R030 as fa43, " & _
                        "R031 as Curr, R032 as Rate, R033 as Nation,R034 as EBox,R035 as FA108,R036 as fa70, R037 as a1k32, R038 as a1t01, R039 as a1t02, " & _
                        "R040 as CusNo,R041 as a1k33,R042 as a1k38, R043 as emailcc " & StrSQLa & _
                        "From Accrpt2470,Customer Where ID='" & strUserNum & "' And Substr(R040,4,8)=cu01(+) And Substr(R040,12,1)=cu02(+) " & _
                        "Order by R002 asc,CusNo asc, R031 asc, R003 asc, R004 asc"
        'end 2017/08/07
   Else
        StrSQLa = StrSQLa & " order by FagentNo asc, Curr asc, DocDate asc, DocNo asc"
   End If
   
   adoacc1k0.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc1k0.RecordCount = 0 Then
      InsertQueryLog (0) 'Add By Sindy 2010/12/22
      strCon10 = MsgText(602)
      adoacc1k0.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   
   'Add by Amy 2024/08/02 +產生csv
   If Check1.Value = 1 Then
      If CSVSave(strXLSFile) = True Then
         If MsgBox("CSV檔案已產生！" & vbCrLf & vbCrLf & strXLSFile & vbCrLf & vbCrLf & "是否開啟資料夾？", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
            ShellExecute hLocalFile, "explore", strExcelPath, vbNullString, vbNullString, 1
         End If
      End If
      If adoacc1k0.State <> adStateClosed Then adoacc1k0.Close
      Exit Sub
   'Add by Amy 2013/06/10 +產生Excel
   ElseIf bolExcel = True Then
      'Modified by Morgan 2020/12/11 增加顯示檔名(已開放外專承辦也能用，但會不知道XLS存放位置)
      'ExcelSaveNew
      'MsgBox ("EXCEL檔案已產生！")
      ExcelSaveNew strXLSFile
      If MsgBox("EXCEL檔案已產生！" & vbCrLf & vbCrLf & strXLSFile & vbCrLf & vbCrLf & "是否開啟資料夾？", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
         ShellExecute hLocalFile, "explore", strExcelPath, vbNullString, vbNullString, 1
      End If
      'end 2020/12/11
      adoacc1k0.Close
      Exit Sub
   End If
   
'Modified by Morgan 2014/3/6 已固定用A4,以下程式取消
   PrintDataA4
   adoacc1k0.Close
'
'   'Modify by Morgan 2009/3/27 +A4格式
'   If bolA4 Then
'      PrintDataA4
'      adoacc1k0.Close
'      Exit Sub
'   End If
'   'end 2009/3/27
'
'   intLength = 0
'   douAmount = 0
'   douTAmount = 0: douRAmount = 0 'Add by Morgan 2006/11/30
'   douOverDue1 = 0
'   douOverDue2 = 0
'   douOverDue3 = 0
'   strNo = ""
'   m_DNCurr = "" '2009/6/4 add by sonia
'   m_FNo = "" 'Add by Morgan 2009/7/7
'
'   'Modify by Morgan 2006/11/29 目前XP自定紙張需手動設定並將印表機預設為該紙張
'   'Printer.Width = 13000
'   'Printer.Height = 13000
'   If bolChina Then
'      Printer.PaperSize = PUB_GetPaperSize(5)
'      iRowH = 280
'   Else
'      Printer.PaperSize = PUB_GetPaperSize(4)
'      iRowH = 200
'   End If
'   iXo = 0
'   iYo = -1 * (Printer.Height - Printer.ScaleHeight) / 2
'   iPageRows = 15
'   'END 2006/11/29
'
'   Do While adoacc1k0.EOF = False
'      '2005/8/2 ADD BY SONIA
'      If Text3 <> MsgText(601) Then
'         If adoacc1k0.Fields("NATION").Value < Text3 Then
'            GoTo NextSkip
'         End If
'      End If
'      If Text4 <> MsgText(601) Then
'         If adoacc1k0.Fields("NATION").Value > Text4 & "z" Then
'            GoTo NextSkip
'         End If
'      End If
'      '2005/8/2 END
'      If Len(adoacc1k0.Fields("DocNo").Value) = 10 And strDocNo = Mid(adoacc1k0.Fields("DocNo").Value, 1, 8) And adoacc1k0.Fields("OAmount").Value = 0 Then
'         GoTo NextSkip
'      Else
'         If Len(adoacc1k0.Fields("DocNo").Value) <> 10 Then
'            strDocNo = adoacc1k0.Fields("DocNo").Value
'         Else
'            strDocNo = Mid(adoacc1k0.Fields("DocNo").Value, 1, 8)
'         End If
'      End If
'      '2009/6/3 MODIFY BY SONIA 請款幣別不同也要跳頁
'      'If strNo <> adoacc1k0.Fields("FagentNo").Value Then
'      If strNo <> adoacc1k0.Fields("FagentNo").Value & adoacc1k0.Fields("Curr").Value Then
'         If douAmount <> 0 Then
'            PrintSum
'            douOverDue1 = 0
'            douOverDue2 = 0
'            douOverDue3 = 0
'            douAmount = 0
'            douTAmount = 0: douRAmount = 0 'Add by Morgan 2006/11/30
'            Printer.NewPage
'            iPageNo = 0
'         End If
'         intCounter = 0
'         PrintHead
'         '2009/6/4 MODIFY BY SONIA 請款幣別不同也要跳頁
'         'strNo = adoacc1k0.Fields("FagentNo").Value
'         strNo = adoacc1k0.Fields("FagentNo").Value & adoacc1k0.Fields("Curr").Value
'         m_DNCurr = adoacc1k0.Fields("Curr").Value
'         '2009/6/4 end
'         m_FNo = adoacc1k0.Fields("FagentNo").Value 'Add by Morgan 2009/7/7
'      Else
'         intCounter = intCounter + 1
'         If intCounter > iPageRows Then
'            intCounter = 0
'            Printer.NewPage
'            PrintHead
'         End If
'      End If
'      PrintRow
'
'NextSkip:
'      adoacc1k0.MoveNext
'   Loop
'   PrintSum
'   adoacc1k0.Close
'   Printer.EndDoc
End Sub

'Mark by Amy 2022/06/14 因資料不正確,而將 2021/10/26  及2022/03/07 改的還原至以前計算方式-莘
'2021/10/26 小數位與請款單有差異,非 整批 服務費 改抓a1k39/ 規費 改抓a1k40,改後會因整批使用原公式及非整批抓值,導致Y27766資料(非整批直接放值)不正確(婧瑄2022/5/26 (週四) 下午 02:06 mail)
'*************************************************
'  轉成Excel檔案(動態產生欄位)
'  下客戶編號條件增加客戶編號/名稱欄位並依客戶編號產生不同工作表,代理編號不同產生不同檔案
'  Add by Amy 2017/02/02
'*************************************************
'Modified by Morgan 2020/12/11 +pFileName
Private Sub ExcelSaveNew_Old2(Optional ByRef pFileName As String)
'    Dim xlsAgentPoint As New Excel.Application
'    Dim wksrpt As New Worksheet
'    Dim xlsFileName As String, strSql As String, strSQL2 As String
'    Dim strTmp(8) As String, strField As String, strValue As String, strFieldN As String 'Modify by Amy 2021/10/26 strTmp 原:7
'    Dim strOldAgNo As String, strOldAppNo As String
'    Dim intField As Integer, intXlsSheet As Integer
'    Dim bolFormula As Boolean
'    Dim strOldCurr As String, intStartR As String, intNo As Integer 'Add by Amy 2017/03/03
'    Dim bolFirst As Boolean 'Add by Amy 2017/08/07
'    Dim strWkName As String 'Add by Amy 2017/09/25 for 2010 工作表名稱為中文
'    Dim strAllField As String, strAllWidth As String 'Add by Amy 2020/11/19
'On Error GoTo ErrHand
'
'    'Modify by Amy 2020/11/19 非Y27766 需增加USD欄
''    ReDim strF(17)
''    ReDim intwith(17)
'    '中文Excel抬頭(國籍為020)
'    If bolChina = True Then
'        strAllField = "編號,本事務所名稱,帳單日期<mm/dd/yyy>,帳單編號,本所案號,貴所案號,Application No.(Only for new case),Filing Date(Only for new case),商標/專利 申請號,Filing Date,Category<select>" & _
'                            ",客戶編號,客戶名稱,幣別,服務費,規費,雜費,帳單金額"
'        strAllWidth = "5.5, 12, 10, 10, 9, 13, 13, 13, 13, 13 " & _
'                            ", 6, 10, 10, 7, 8, 8, 9.6, 10"
'    Else
'        strAllField = "NO.,Law Firm's Name<select>,Date<mm/dd/yyy>,Invoice No.,Our Ref,Your Ref,Application No.(Only for new case),Filing Date(Only for new case),TradeMark/Patent Application No.,Filing Date,Category<select>" & _
'                           ",Application No.,Application,Currency<select>,Attorney Fee,Official Fee,Disburesment Fee,Total Fee"
'        strAllWidth = "5.5, 12, 10, 10, 9, 13, 13, 13, 13, 13" & _
'                            ", 6, 10, 10, 7, 8, 8, 9.6, 10"
'    End If
'    'Y27766(特殊格式) 維持不變
'    If Left(Text1, 6) <> "Y27766" Then
'        strAllField = strAllField & ",USD"
'        strAllWidth = strAllWidth & ",5"
'    End If
'    'Add by Amy 2022/03/07 因有舊資料未收款(服務費、規費不會另外列),而非整批「請款金額」已改為公式,怕有公式會與原始資料不合,故顯示＊
'    strAllField = strAllField & ",＊金額</br>需再確認,整批主號"
'    strAllWidth = strAllWidth & ",6.5,10"
'    'end 2021/3/07
'    strF = Split(strAllField, ",")
'    intWidth = Split(strAllWidth, ",")
'    'end 2020/11/19
'NextAg:
'    intField = 65: intXlsSheet = 1: intCounter = 1: strOldAppNo = "": intNo = 1: bolFirst = True
'
'    'Modify by Amy 2017/03/16 有下客戶編號條件顯示客戶編號前六碼-莘
'    If strOldAgNo = MsgText(601) Then
'        xlsFileName = Text1 & "催款單" & IIf(bolShowCus = True, "(vs " & Left(Text10, 6) & ")", "") & ServerDate & MsgText(43)
'    Else
'         xlsFileName = strOldAgNo & "催款單" & IIf(bolShowCus = True, "(vs " & Left(Text10, 6) & ")", "") & ServerDate & MsgText(43)
'    End If
'    If Dir(strExcelPath & xlsFileName) = MsgText(601) Then
'       If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
'            MkDir strExcelPath
'       End If
'    Else
'         Kill strExcelPath & xlsFileName
'    End If
'
'    pFileName = strExcelPath & xlsFileName 'Added by Morgan 2020/12/11
'
'    xlsAgentPoint.SheetsInNewWorkbook = 3 'Added by Lydia 2019/03/13 預設工作表數量
'    xlsAgentPoint.Workbooks.add
'    xlsAgentPoint.Application.WindowState = xlMinimized
'
'NextCus:
'    If intXlsSheet > 3 Then
'        'Modified by Morgan 2020/10/20 應加在最後(目前的後面)
'        'xlsAgentPoint.Worksheets.add
'        xlsAgentPoint.Worksheets.add After:=wksrpt
'        'end 2020/10/20
'    End If
'
'    'Modify by Amy 2017/09/25 for 工作表名稱改為中文
'    If strWkName = MsgText(601) Then strWkName = Left(xlsAgentPoint.Worksheets(1).Name, Len(xlsAgentPoint.Worksheets(1).Name) - 1)
'    Set wksrpt = xlsAgentPoint.Worksheets(strWkName & intXlsSheet)
'    'end 2017/09/25
'    wksrpt.Activate
'
'    With adoacc1k0
'        Do While .EOF = False
'            'Add by Amy 2017/08/07 從上搬下來 SetField 傳入代理人編號
'            If bolFirst = True Then
'                Call SetField(xlsAgentPoint.Version, wksrpt, intField, IIf(strOldAgNo = MsgText(601), .Fields("FagentNo"), strOldAgNo), bolShowCus)
'                intTitleRow = intCounter
'                intCounter = intCounter + 1: intStartR = intCounter
'                bolFirst = False
'            End If
'            'end 2017/08/07
'
'            If bolShowCus = True Then
'                If (strOldAppNo <> MsgText(601) And strOldAppNo <> Mid(.Fields("CusNo"), 4)) Or (strOldAgNo <> MsgText(601) And strOldAgNo <> .Fields("FagentNo")) Then
'                    '合計
'                    If intCounter > intTitleRow + 1 Then Call SetLastSet(wksrpt, intCounter, intField, intStartR)
'                    '改工作表名稱
'                    If strOldAppNo <> Mid(.Fields("CusNo"), 4) Then
'                        wksrpt.Name = strOldAppNo
'                         intXlsSheet = intXlsSheet + 1
'                         strOldAppNo = ""
'                    End If
'                    intCounter = 1
'                    If strOldAgNo <> MsgText(601) And strOldAgNo <> .Fields("FagentNo") Then
'                        strOldAgNo = .Fields("FagentNo")
'                        If Val(xlsAgentPoint.Version) < 12 Then
'                           xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
'                        Else
'                           xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
'                        End If
'                        xlsAgentPoint.Workbooks.Close
'                        bolFirst = True 'Add by Amy 2017/08/07
'                        GoTo NextAg
'                    Else
'                        bolFirst = True 'Add by Amy 2017/08/07
'                        GoTo NextCus
'                    End If
'                End If
'            End If
'            'Add by Amy 2017/03/03 +幣別合計(Y27766除外)
'            If Left(Text1, 6) <> "Y27766" And strOldCurr <> MsgText(601) And strOldCurr <> .Fields("Curr") And intCounter > intTitleRow + 1 Then
'                Call SetLastSet(wksrpt, intCounter, intField, intStartR, True)
'                intCounter = intCounter + 2
'                intStartR = intCounter
'            End If
'            'Modify by Amy 2021/10/26 小數位與請款單有差異,應財務要求改為一致,＋a1k32判斷
'            Call GetOtherData(strTmp(), "" & .Fields("a1k32"))
'            'Modify by Amy 2017/08/07 中文Excel抬頭(國籍為020)
'            For i = 0 To UBound(strF)
'                bolFormula = False: strValue = ""
'                Select Case i
'                    Case GetValue("NO."), GetValue("編號") '序號
'                        strValue = intNo
'                    Case GetValue("Law Firm's Name<select>"), GetValue("本事務所名稱")
'                        strValue = "Tai E"
'                    Case GetValue("Date<mm/dd/yyy>"), GetValue("帳單日期<mm/dd/yyy>")
'                        strValue = ChangeTStringToWDateString(.Fields("A1K02"))
'                    Case GetValue("Invoice No."), GetValue("帳單編號") 'A1K01
'                        strValue = .Fields("DocNo")
'                    Case GetValue("Our Ref"), GetValue("本所案號") '本所案號
'                        strValue = .Fields("CaseNo1") & "-" & .Fields("CaseNo2")
'                    Case GetValue("Your Ref"), GetValue("貴所案號") '彼所案號
'                        strValue = strTmp(0)
'                    Case GetValue("Application No.(Only for new case)") '該案號為第一張請款單之申請號
'                        strValue = strTmp(1)
'                    Case GetValue("Filing Date(Only for new case)") '該案號為第一張請款單
'                        strValue = strTmp(2)
'                    Case GetValue("TradeMark/Patent Application No."), GetValue("商標/專利 申請號") '申請號
'                        strValue = strTmp(3)
'                    Case GetValue("Filing Date")
'                        strValue = strTmp(4)
'                    Case GetValue("Category<select>")
'                        strValue = ""
'                    Case GetValue("Application No."), GetValue("客戶編號") '客戶編號
'                        If bolShowCus = True Then strValue = "" & .Fields("CusNo")
'                    Case GetValue("Application"), GetValue("客戶名稱") '客戶名稱
'                        If bolShowCus = True Then strValue = "" & .Fields("CusName")
'                    Case GetValue("Currency<select>"), GetValue("幣別") 'A1K18
'                        strValue = .Fields("Curr")
'                    Case GetValue("Attorney Fee"), GetValue("服務費")
'                        bolFormula = True
'                        'Modify by Amy 2021/10/26 +if 小數位與請款單有差異,應財務要求改為一致,服務費改抓 A1k39/雜費a1k40,a1k32=C(整批)照原寫法
'                        If "" & .Fields("a1k32") = "C" Then
'                            If bolChina = False Then
'                                strValue = "=" & Chr(GetValue("Total Fee") + intField) & intCounter & "-" & _
'                                        Chr(GetValue("Official Fee") + intField) & intCounter & "-" & Chr(GetValue("Disburesment Fee") + intField) & intCounter
'                            Else
'                                strValue = "=" & Chr(GetValue("帳單金額") + intField) & intCounter & "-" & _
'                                        Chr(GetValue("規費") + intField) & intCounter & "-" & Chr(GetValue("雜費") + intField) & intCounter
'                            End If
'                        Else
'                            strValue = strTmp(8)
'                        End If
'                    Case GetValue("Official Fee"), GetValue("規費")
'                        strValue = strTmp(5)
'                    Case GetValue("Disburesment Fee"), GetValue("雜費")
'                        strValue = strTmp(6)
'                    Case GetValue("Total Fee"), GetValue("帳單金額") 'A1K08-A1K31
'                        'Modify by Amy 2021/10/26 小數位與請款單有差異,應財務要求改為一致,整批(a1k32='C') 照舊,其他改公式計算-莘
'                        If "" & .Fields("a1k32") = "C" Then
'                            strValue = Val(.Fields("FAmount"))
'                        Else
'                            If bolChina = False Then
'                                strValue = "=" & Chr(GetValue("Attorney Fee") + intField) & intCounter & "+" & Chr(GetValue("Official Fee") + intField) & intCounter
'                            Else
'                                strValue = "=" & Chr(GetValue("服務費") + intField) & intCounter & "+" & Chr(GetValue("規費") + intField) & intCounter
'                            End If
'                        End If
'                        'Mark by Amy 2021/02/02 改抓a1k38
'                        'strTmp(7) = Val(strValue) - Val(strTmp(5)) - Val(strTmp(6)) 'Add by Amy 2020/11/19 算 服務費 值,算USD用
'                    'Add by Amy 2020/11/19 +USD
'                    Case GetValue("USD")
'                        If "" & .Fields("a1k33") = "4" Then
'                            'Modify by Amy 2021/02/02 統一改抓a1k38(美金請款金額)
''                            strValue = PUB_GetDNRate("" & .Fields("a1k02"), "" & .Fields("Curr"))
''                            '規費/雜費/服務費 各自*當時匯率後去小數相加,為盡量與 frmacc2480一致-莘
''                            strValue = Trunc(Val(strTmp(5)) * Val(strValue)) + Trunc(Val(strTmp(6)) * Val(strValue)) _
''                                            + Trunc(Val(strTmp(7)) * Val(strValue))
'                            strValue = Val("" & .Fields("a1k38"))
'                        End If
'                    'Add by Amy 2021/03/07
'                    '最後一欄公式判斷計算之總金額是否和原始值一致-婉莘
'                    Case GetValue("＊金額</br>需再確認")
'                        If bolChina = False Then
'                            strExc(0) = wksrpt.Range(Chr(GetValue("Total Fee") + intField) & intCounter).Value
'                        Else
'                            strExc(0) = wksrpt.Range(Chr(GetValue("帳單金額") + intField) & intCounter).Value
'                        End If
'                        If Val(strExc(0)) <> Val(.Fields("FAmount")) Then
'                            strValue = "＊"
'                        End If
'                    '整批未結清需顯示整批主號-秀玲
'                    Case GetValue("整批主號")
'                        If "" & .Fields("a1k32") = "C" Then
'                            strValue = "" & .Fields("a1t02")
'                        End If
'                    'end 2022/03/07
'                End Select
'                If bolFormula = True Then
'                    wksrpt.Range(Chr(i + intField) & intCounter).Formula = strValue
'                    If Left(Text1, 6) = "Y27766" And i = GetValue("Total Fee") Then
'                        'Attorney Fee 不顯示公式,改顯示值
'                        wksrpt.Range(Chr(GetValue("Attorney Fee") + intField) & intCounter).Copy
'                        wksrpt.Range(Chr(GetValue("Attorney Fee") + intField) & intCounter).PasteSpecial xlPasteValues, xlPasteSpecialOperationNone, False, False
'                        'Total Fee 改顯示公式
'                         strValue = "=IF(" & Chr(GetValue("Attorney Fee") + intField) & intCounter & "+" & Chr(GetValue("Official Fee") + intField) & intCounter & _
'                                    "+" & Chr(GetValue("Disburesment Fee") + intField) & intCounter & "=0,"""", " & Chr(GetValue("Attorney Fee") + intField) & intCounter & "+" & Chr(GetValue("Official Fee") + intField) & intCounter & _
'                                    "+" & Chr(GetValue("Disburesment Fee") + intField) & intCounter & ")"
'                        wksrpt.Range(Chr(i + intField) & intCounter).Formula = strValue
'                    End If
'                Else
'                    If i = GetValue("Application No.(Only for new case)") Or i = GetValue("TradeMark/Patent Application No.") Or i = GetValue("商標/專利 申請號") Then
'                        wksrpt.Range(Chr(i + intField) & intCounter).NumberFormatLocal = "@"
'                    End If
'                    wksrpt.Range(Chr(i + intField) & intCounter).Value = strValue
'                End If
'                'Add by Amy 2020/11/19 不是Y27766 若列印幣別格式為4,且特殊請款單=C(整批)顯示顏色
'                'Modify by Amy 2021/10/26 小數位與請款單有差異,應財務要求改為一致,故拿掉"" & .Fields("a1k33") = "4"(列印幣別格式),讓財務注意-莘
'                If Left(Text1, 6) <> "Y27766" Then
'                    If i = GetValue("USD") And "" & .Fields("a1k32") = "C" Then
'                        wksrpt.Range(Chr(i + intField) & intCounter).Interior.ColorIndex = 40   '膚色
'                        wksrpt.Range(Chr(i + intField) & intCounter).Interior.tintandshade = 0.2 '設深淺
'                    End If
'                End If
'            Next i
'            'end 2017/08/07
'            intCounter = intCounter + 1: intNo = intNo + 1
'            If bolShowCus = True Then
'                strOldAgNo = "" & .Fields("FagentNo")
'                strOldAppNo = Mid("" & .Fields("CusNo"), 4)
'            End If
'            strOldCurr = "" & .Fields("Curr") 'Add by Amy 2017/03/03
'            .MoveNext
'        Loop
'        '最後一個加總
'        'Add by Amy 2017/03/03 +幣別合計(Y27766除外)
'        If Left(Text1, 6) = "Y27766" Then
'            Call SetLastSet(wksrpt, intCounter, intField)
'        Else
'            Call SetLastSet(wksrpt, intCounter, intField, intStartR)
'            If bolChina = True And strOldAppNo <> MsgText(601) Then
'                wksrpt.Name = strOldAppNo
'
'            'Added by Morgan 2020/10/20
'            ElseIf bolShowCus = True Then
'               wksrpt.Name = strOldAppNo
'            'end 2020/10/20
'            End If
'        End If
'    End With
'
'    If Val(xlsAgentPoint.Version) < 12 Then
'       xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
'    Else
'       xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
'    End If
'    xlsAgentPoint.Workbooks.Close
'    xlsAgentPoint.Quit
'    StatusClear
'    Exit Sub
'
'ErrHand:
'    MsgBox Err.Description, , MsgText(5)
'    If Val(xlsAgentPoint.Version) < 12 Then
'        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
'    Else
'        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
'    End If
'    xlsAgentPoint.Workbooks.Close
'    xlsAgentPoint.Quit
'    Set xlsAgentPoint = Nothing
'    Set wksrpt = Nothing
End Sub

Private Sub SetLastSet(ByRef Wks As Worksheet, ByVal intCount As Integer, ByVal intField As Integer, Optional ByVal intStartR As Integer = 0, Optional ByVal bolOnlySum As Boolean = False)
    Dim strValue As String, strStyle As String
    Dim bolShow As Boolean, bolFormula As Boolean
    Dim intNowR As Integer 'Add by Amy 2017/03/03
    Dim strFieldN As String 'Add by Amy 2017/08/07
    
    'Add by Amy 2017/03/03
    intNowR = 0
    'Modify by Amy 2017/08/07 +中文Excel抬頭(國籍為020),避免某些欄位無法Del故最後顯示欄位名稱改至此做
    If intStartR = 0 Then intStartR = intTitleRow + 1: intNowR = intStartR
   
    For i = 0 To UBound(strF)
        bolShow = True: bolFormula = True: strStyle = ""
        Select Case i
            Case GetValue("NO."), GetValue("編號")
                If Left(Text1, 6) <> "Y27766" Then bolShow = False 'Add by Amy2017/03/03
                bolFormula = False
                strValue = "TOTAL"
            Case GetValue("Date<mm/dd/yyy>"), GetValue("帳單日期<mm/dd/yyy>")
                bolShow = False
                strStyle = "m/d/yyyy"
            Case GetValue("Invoice No."), GetValue("帳單編號")
                'Add by Amy2017/03/03
                If Left(Text1, 6) <> "Y27766" Then
                    bolFormula = False
                    If bolChina = True Then
                        strValue = "合 計"
                    Else
                        strValue = "Total"
                    End If
                Else
                    bolShow = False
                End If
            Case GetValue("Filing Date(Only for new case)")
                bolShow = False
                strStyle = "m/d/yyyy"
            Case GetValue("Filing Date")
                bolShow = False
                strStyle = "m/d/yyyy"
            Case GetValue("Currency<select>"), GetValue("幣別")
                bolFormula = False
                strValue = "=n" & intStartR 'Modify by Amy 2017/03/03
            Case GetValue("Attorney Fee"), GetValue("服務費")
                If bolChina = True Then
                    strValue = "=Sum(" & Chr(GetValue("服務費") + intField) & intStartR & ":" & Chr(GetValue("服務費") + intField) & intCount - 1 & ")"
                Else
                    strValue = "=Sum(" & Chr(GetValue("Attorney Fee") + intField) & intStartR & ":" & Chr(GetValue("Attorney Fee") + intField) & intCount - 1 & ")"
                End If
                strStyle = "0.00"
            Case GetValue("Official Fee"), GetValue("規費")
                If bolChina = True Then
                    strValue = "=Sum(" & Chr(GetValue("規費") + intField) & intStartR & ":" & Chr(GetValue("規費") + intField) & intCount - 1 & ")"
                Else
                    strValue = "=Sum(" & Chr(GetValue("Official Fee") + intField) & intStartR & ":" & Chr(GetValue("Official Fee") + intField) & intCount - 1 & ")"
                End If
                strStyle = "0.00"
            Case GetValue("Disburesment Fee"), GetValue("雜費")
                If bolChina = True Then
                    strValue = "=Sum(" & Chr(GetValue("雜費") + intField) & intStartR & ":" & Chr(GetValue("雜費") + intField) & intCount - 1 & ")"
                Else
                    strValue = "=Sum(" & Chr(GetValue("Disburesment Fee") + intField) & intStartR & ":" & Chr(GetValue("Disburesment Fee") + intField) & intCount - 1 & ")"
                End If
                strStyle = "0.00"
            Case GetValue("Total Fee"), GetValue("帳單金額")
                If bolChina = True Then
                    strValue = "=Sum(" & Chr(GetValue("帳單金額") + intField) & intStartR & ":" & Chr(GetValue("帳單金額") + intField) & intCount - 1 & ")"
                Else
                    strValue = "=Sum(" & Chr(GetValue("Total Fee") + intField) & intStartR & ":" & Chr(GetValue("Total Fee") + intField) & intCount - 1 & ")"
                End If
                strStyle = "0.00"
            Case Else
                bolShow = False
        End Select
        If bolShow = True Then
            If bolFormula = True Then
                Wks.Range(Chr(i + intField) & intCount + intNowR).Formula = strValue
            Else
                Wks.Range(Chr(i + intField) & intCount + intNowR).Value = strValue
            End If
        End If
        If strStyle <> MsgText(601) Then
            Wks.Range(Chr(i + intField) & intStartR & ":" & Chr(i + intField) & intCount + 2).NumberFormatLocal = strStyle
        End If
    Next i
    
    If bolOnlySum = False Then
        '設定內文置中
        If bolChina = True Then
            Wks.Range(Chr(GetValue("本事務所名稱") + intField) & intTitleRow & ":" & Chr(GetValue("幣別") + intField) & intCount + 2).HorizontalAlignment = xlCenter
            Wks.Range(Chr(GetValue("貴所案號") + intField) & intTitleRow & ":" & Chr(GetValue("貴所案號") + intField) & intCount + 2).HorizontalAlignment = xlLeft
        Else
            Wks.Range(Chr(GetValue("Law Firm's Name<select>") + intField) & intTitleRow & ":" & Chr(GetValue("Currency<select>") + intField) & intCount + 2).HorizontalAlignment = xlCenter
            Wks.Range(Chr(GetValue("Your Ref") + intField) & intTitleRow & ":" & Chr(GetValue("Your Ref") + intField) & intCount + 2).HorizontalAlignment = xlLeft
        End If
        'Add by Amy 2022/03/07 ＊金額需再確認欄置中
        Wks.Range(Chr(GetValue("＊金額</br>需再確認") + intField) & intTitleRow & ":" & Chr(GetValue("＊金額</br>需再確認") + intField) & intCount + 2).HorizontalAlignment = xlCenter
        'Add by Amy 2017/03/03 申請人名稱/Your Ref靠左-莘
        If bolShow = True And Left(Text1, 6) <> "Y27766" Then
            If bolChina = True Then
                Wks.Range(Chr(GetValue("客戶編號") + intField) & intTitleRow & ":" & Chr(GetValue("客戶編號") + intField) & intCount + 2).HorizontalAlignment = xlLeft
            Else
                'Modify by Amy 2022/07/20 申請人英文顯示錯誤 原:Application-莘
                Wks.Range(Chr(GetValue("Applicant") + intField) & intTitleRow & ":" & Chr(GetValue("Applicant") + intField) & intCount + 2).HorizontalAlignment = xlLeft
            End If
        End If
        'end 2017/03/03
        
        With Wks.Range(Chr(intField) & intTitleRow & ":" & Chr(UBound(strF) + intField) & intCount + 2)
            .Font.Name = "新細明體"
            .Font.Size = 10
            
            '框線
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        End With
        If bolChina = True Then
            Wks.Range(Chr(GetValue("規費") + intField) & intTitleRow & ":" & Chr(GetValue("規費") + intField) & intCount + 2).Borders(xlEdgeLeft).LineStyle = xlDouble
        Else
            Wks.Range(Chr(GetValue("Official Fee") + intField) & intTitleRow & ":" & Chr(GetValue("Official Fee") + intField) & intCount + 2).Borders(xlEdgeLeft).LineStyle = xlDouble
        End If
        
        '更改欄位名稱
        For i = 0 To UBound(strF)
            If (Left(Text1, 6) <> "Y27766" And (strF(i) = "Application No.(Only for new case)" Or strF(i) = "Filing Date(Only for new case)" Or strF(i) = "Category<select>")) _
             Or (Left(Text1, 6) = "Y27766" And strF(i) = "TradeMark/Patent Application No.") Then
                strFieldN = strF(i)
            'Add by Amy 2022/03/07 +取代</br>為換行
            ElseIf InStr(strF(i), "</br>") > 0 Then
                strFieldN = Replace(strF(i), "</br>", vbCrLf)
            ElseIf InStr(strF(i), "<") > 0 Then
                If strF(i) = "Date<mm/dd/yyy>" And Left(Text1, 6) <> "Y27766" Then
                    strFieldN = "Invoice " & Left(strF(i), Val(InStr(strF(i), "<")) - 1) & vbCrLf & Mid(strF(i), InStr(strF(i), "<"))
                Else
                    strFieldN = Left(strF(i), Val(InStr(strF(i), "<")) - 1) & vbCrLf & Mid(strF(i), InStr(strF(i), "<"))
                End If
            ElseIf InStr(strF(i), "(") > 0 Then
                strFieldN = Left(strF(i), Val(InStr(strF(i), "(")) - 1) & vbCrLf & Mid(strF(i), InStr(strF(i), "("))
            ElseIf strF(i) = "Disburesment Fee" Or strF(i) = "TradeMark/Patent Application No." Then
                strFieldN = Left(strF(i), Val(InStr(strF(i), " ")) - 1) & vbCrLf & Trim(Mid(strF(i), InStr(strF(i), " ")))
            Else
                If Left(Text1, 6) = "Y27766" And i = GetValue("Our Ref") Then
                    strFieldN = "Your Ref"
                ElseIf Left(Text1, 6) = "Y27766" And i = GetValue("Your Ref") Then
                    strFieldN = "Murata's Ref"
                Else
                    strFieldN = strF(i)
                End If
            End If
            Wks.Range(Chr(i + intField) & intTitleRow).Value = strFieldN
            If strF(i) = "Category<select>" Then
                '備註
                Wks.Range(Chr(i + intField) & intTitleRow).AddComment
                Wks.Range(Chr(i + intField) & intTitleRow).Comment.Visible = False
                Wks.Range(Chr(i + intField) & intTitleRow).Comment.Text Text:="1 : New application" & Chr(10) & "2 : Office Action (filing remarks and/or claim amendment)" & _
                                                                                                    Chr(10) & "3 : Office Action (others: notice of publication, search report, etc.)" & Chr(10) & "4 : Issue" & Chr(10) & _
                                                                                                    "5 : maintenance" & Chr(10) & "6 : Litigation" & Chr(10) & "Z : Other", Start:=200
                Wks.Range(Chr(i + intField) & intTitleRow).Comment.Shape.ScaleWidth 2.5, 0, 0
                Wks.Range(Chr(i + intField) & intTitleRow).Comment.Shape.ScaleHeight 2, 0, 0
            End If
        Next i
        
        '刪除不顯示欄
        If Left(Text1, 6) = "Y27766" Then
            Wks.Range(Chr(GetValue("TradeMark/Patent Application No.") + intField) & ":" & Chr(GetValue("Filing Date") + intField)).Delete Shift:=xlToLeft
            '客戶編號/名稱不顯示-莘
            'Modify by Amy 2022/07/20 申請人英文顯示錯誤 原:Application
            Wks.Range(Chr(GetValue("Applicant No.") + intField - 2) & ":" & Chr(GetValue("Applicant") + intField - 2)).Delete Shift:=xlToLeft
        Else
            Wks.Range(Chr(GetValue("Application No.(Only for new case)") + intField) & ":" & Chr(GetValue("Filing Date(Only for new case)") + intField)).Delete Shift:=xlToLeft
            Wks.Range(Chr(GetValue("Filing Date") + intField - 2) & ":" & Chr(GetValue("Category<select>") + intField - 2)).Delete Shift:=xlToLeft
            '沒下客戶編號條件 客戶編號 / 名稱欄位顯示 - 莘
            If bolShowCus = False Then
                If bolChina = True Then
                    Wks.Range(Chr(GetValue("客戶編號") + intField - 4) & ":" & Chr(GetValue("客戶名稱") + intField - 4)).Delete Shift:=xlToLeft
                Else
                    'Modify by Amy 2022/07/20 申請人英文顯示錯誤 原:Application
                    Wks.Range(Chr(GetValue("Applicant No.") + intField - 4) & ":" & Chr(GetValue("Applicant") + intField - 4)).Delete Shift:=xlToLeft
                End If
            End If
        End If
    End If
    'end 2017/08/07
End Sub

'Modify by Amy 2017/08/02 +Excel版本
Private Sub SetField(ByVal stVersion As String, ByRef Wks As Worksheet, ByVal intField As Integer, ByVal stFaNo As String, ByVal bolShowCu As Boolean)
    'Modify by Amy 2017/08/07 +表頭(Y27766特殊無表頭,故維持舊格式,固定寫於C欄-婉莘),欄位名稱最後顯示改至SetLastSet做
    If Left(Text1, 6) <> "Y27766" Then
        If bolChina = True Then
            'Modified by Morgan 2020/3/30
            'Wks.Range("C" & intCounter).Value = "台一國際專利法律事務所"
            Wks.Range("C" & intCounter).Value = CompNameQuery("2")
            'end 2020/3/30
            intCounter = intCounter + 2
            Wks.Range("C" & intCounter).Value = "應收帳款對帳單"
            'Add by Amy 2022/06/14 +提醒文字
            Wks.Range(Chr(intField + GetValue("商標/專利 申請號")) & intCounter).Font.Color = vbRed
            Wks.Range(Chr(intField + GetValue("商標/專利 申請號")) & intCounter).Font.Bold = True
            Wks.Range(Chr(intField + GetValue("商標/專利 申請號")) & intCounter).Value = "雜費計算可能會有小數點誤差 , 如需顯示給代理人看, 需再檢查確認 ！"
            intCounter = intCounter + 2
            Wks.Range("C" & intCounter).Value = "代理人代號: " & Text1 & "~" & Text2
            intCounter = intCounter + 1
            Wks.Range("C" & intCounter).Value = "代理人名稱: " & GetPrjName1(stFaNo)
            intCounter = intCounter + 1
            Wks.Range("C" & intCounter).Value = "Date: " & IIf(Trim(Val(FCDate(MaskEdBox1.Text))) <> MsgText(601), ChangeTStringToWDateString(Val(FCDate(MaskEdBox1.Text))), "") & "~" & _
                                                                        IIf(Trim(Val(FCDate(MaskEdBox2.Text))) <> MsgText(601), ChangeTStringToWDateString(Val(FCDate(MaskEdBox2.Text))), "")
            
            'Mark by Amy 2022/06/14 還原至 2021/10/26 前改的
'            'Modify by Amy 2021/10/26 +說明
'            intCounter = intCounter + 1
'            Wks.Range(Chr(intField + GetValue("服務費")) & intCounter).Value = "A" & vbCrLf & "(含雜費)"
'            Wks.Range(Chr(intField + GetValue("規費")) & intCounter).Value = "B" & vbCrLf
'            Wks.Range(Chr(intField + GetValue("雜費")) & intCounter).Value = vbCrLf & "(參考用)"
'            Wks.Range(Chr(intField + GetValue("帳單金額")) & intCounter).Value = "A+B" & vbCrLf
'            Wks.Range(Chr(intField + GetValue("服務費")) & intCounter & ":" & Chr(intField + GetValue("帳單金額")) & intCounter).HorizontalAlignment = xlCenter
'            Wks.Range(Chr(intField + GetValue("服務費")) & intCounter & ":" & Chr(intField + GetValue("帳單金額")) & intCounter).VerticalAlignment = xlCenter

        Else
            Wks.Range("C" & intCounter).Value = "Tai E International Patent and Law Office"
            intCounter = intCounter + 2
            Wks.Range("C" & intCounter).Value = "Statement of account"
            'Add by Amy 2022/06/14 +提醒文字
            Wks.Range(Chr(intField + GetValue("TradeMark/Patent Application No.")) & intCounter).Font.Color = vbRed
            Wks.Range(Chr(intField + GetValue("TradeMark/Patent Application No.")) & intCounter).Font.Bold = True
            Wks.Range(Chr(intField + GetValue("TradeMark/Patent Application No.")) & intCounter).Value = "雜費計算可能會有小數點誤差 , 如需顯示給代理人看, 需再檢查確認 ！"
            intCounter = intCounter + 2
            'Modify by Amy 2022/07/20 申請人英文顯示錯誤 原:Application
            Wks.Range("C" & intCounter).Value = "Applicant no.: " & Text1 & "~" & Text2
            intCounter = intCounter + 1
            Wks.Range("C" & intCounter).Value = "Name of Applicant: " & GetFAgentName(stFaNo)
            'end 2022/07/20
            intCounter = intCounter + 1
            Wks.Range("C" & intCounter).Value = "Date: " & IIf(Trim(Val(FCDate(MaskEdBox1.Text))) <> MsgText(601), ChangeTStringToWDateString(Val(FCDate(MaskEdBox1.Text))), "") & "~" & _
                                                                        IIf(Trim(Val(FCDate(MaskEdBox2.Text))) <> MsgText(601), ChangeTStringToWDateString(Val(FCDate(MaskEdBox2.Text))), "")
            'Mark by Amy 2022/06/14 還原至 2021/10/26 前改的
'            'Modify by Amy 2021/10/26 +說明
'            intCounter = intCounter + 1
'            Wks.Range(Chr(intField + GetValue("Attorney Fee")) & intCounter).Value = "A" & vbCrLf & "(含雜費)"
'            Wks.Range(Chr(intField + GetValue("Official Fee")) & intCounter).Value = "B" & vbCrLf
'            Wks.Range(Chr(intField + GetValue("Disburesment Fee")) & intCounter).Value = vbCrLf & "(參考用)"
'            Wks.Range(Chr(intField + GetValue("Total Fee")) & intCounter).Value = "A+B" & vbCrLf
'            Wks.Range(Chr(intField + GetValue("Attorney Fee")) & intCounter & ":" & Chr(intField + GetValue("Total Fee")) & intCounter).HorizontalAlignment = xlCenter
'            Wks.Range(Chr(intField + GetValue("Attorney Fee")) & intCounter & ":" & Chr(intField + GetValue("Total Fee")) & intCounter).VerticalAlignment = xlCenter
        End If
        intCounter = intCounter + 1
    End If
    
    With Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strF) + intField) & intCounter)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
        If Val(stVersion) < 12 Then
            If Left(Text1, 6) = "Y27766" Then
                .Interior.ColorIndex = 19 '設置儲存格填充色(黃)
            Else
                .Interior.Color = RGB(255, 165, 79) '設置儲存格填充色
            End If
        Else
            If Left(Text1, 6) = "Y27766" Then
                .Interior.ColorIndex = 19 '設置儲存格填充色(黃)
            Else
                .Interior.ColorIndex = 26 '設置儲存格填充色(粉),新版會出現相容性msg
            End If
        End If
    End With
        
    If Left(Text1, 6) = "Y27766" Then
        Wks.Range(Chr(GetValue("Application No.(Only for new case)") + intField) & intCounter & ":" & Chr(GetValue("Filing Date(Only for new case)") + intField) & intCounter).Interior.ColorIndex = 38 '設置儲存格填充色(粉)
    End If
    
    For i = 0 To UBound(strF)
        Wks.Range(Chr(i + intField) & intCounter).Value = strF(i)
        Wks.Columns(Chr(i + intField) & ":" & Chr(i + intField)).ColumnWidth = intWidth(i)
    Next i
    'end 2017/08/07
End Sub

Private Function GetValue(pFieldN As String) As Integer
   Dim jj As Integer
 
    For jj = 1 To UBound(strF)
       If UCase(strF(jj)) = UCase(pFieldN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function

'Modify by Amy 2022/06/14 a1k32改成Optional(改回2021/10/26前抓法)
'Modify by Amy 2021/10/26 +a1k32 整批照原抓法,非整批改抓a1k40/a1k39
Private Sub GetOtherData(ByRef stTmp() As String, Optional ByVal stA1k32 As String = "")
    Dim Rs As New ADODB.Recordset, Rs2 As New ADODB.Recordset
    Dim j As Integer, intR As Integer, intR2 As Integer, intI As Integer
    Dim strSql As String, strSQL2 As String, strField As String
    
    For j = LBound(stTmp) To UBound(stTmp)
        stTmp(j) = ""
    Next j
    
    'Modify by Amy 2021/10/26 非整批(a1k32<>'C')規費 改抓A1K40並加A1K39(服務費),因與請款單小數位有差異,應財務要求改為一致-莘
    If stA1k32 = "C" Then
        strSql = "Select NVL(round(sA1L05/A1K10,2),0.0) Offical,NVL(Round(sA2/A1K10,2),0.0) Disbur From " & _
                        "(Select Sum(A1L05) sA2 From Acc1L0,Acc1J0 Where A1L03=A1J01 and A1L04=A1J02 And A1j03='雜費' And A1L01='" & adoacc1k0.Fields("DocNo") & "'), " & _
                        "(Select sum(A1L05) as sA1L05 From Acc1L0 Where A1L01='" & adoacc1k0.Fields("DocNo") & "' And Substr(A1L04,-2)='99'), (Select A1K10 From Acc1K0 Where A1K01='" & adoacc1k0.Fields("DocNo") & "')"
    Else
        strSql = "Select Nvl(A1K39,0) as Service,Nvl(A1K40,0) as Offical,Nvl(Round(sA2/A1K10,2),0.0) as Disbur From " & _
                        "(Select Sum(A1L05) sA2 From Acc1L0,Acc1J0 Where A1L03=A1J01 and A1L04=A1J02 And A1j03='雜費' And A1L01='" & adoacc1k0.Fields("DocNo") & "'), " & _
                        "(Select A1K10,A1K39,A1K40 From Acc1K0 Where A1K01='" & adoacc1k0.Fields("DocNo") & "' ) "
    End If
    'end 2021/10/26
    
    Select Case adoacc1k0.Fields("CaseNo1")
        Case "CFP", "FCP", "P"   '專利
            If Left(Text1, 6) = "Y52269" Then
                strField = "Nvl(PA77,PA47)"
            Else
                strField = "PA77"
            End If
            'Modify by Amy 2021/10/26 原:NVL(round(sA1L21/A1K10,2),0.0) Offical,NVL(round(sA2/A1K10,2),0.0) Disbur,小數位與請款單有差異,應財務要求改為一致-莘
            strSql = "Select * From (Select " & strField & " YourR From Patent Where PA01 = '" & adoacc1k0.Fields("CaseNo1") & "' and PA02 = '" & adoacc1k0.Fields("CaseNo2") & "' and PA03 = '" & adoacc1k0.Fields("CaseNo3") & "' and PA04 = '" & adoacc1k0.Fields("CaseNo4") & "')," & _
                        "(" & strSql & ")"
              
            strSQL2 = "Select pa01,pa02,pa03,pa04,Min(A1K01||A1K02) M,pa11 AppNo,pa10 FilingD From Acc1K0,Patent Where PA01='" & adoacc1k0.Fields("CaseNo1") & "' and PA02='" & adoacc1k0.Fields("CaseNo2") & "'  and PA03='" & adoacc1k0.Fields("CaseNo3") & "'  and PA04='" & adoacc1k0.Fields("CaseNo4") & "' " & _
                            "and PA10 is not null And A1K13=PA01 And A1K14=PA02 And A1K15=PA03 And A1K16=PA04 Group by pa01,pa02,pa03,pa04,pa11,pa10"

        Case "CFT", "FCT", "T", "TF"   '商標
            If Left(Text1, 6) = "Y52269" Then
                strField = "Nvl(TM45,TM34)"
            Else
                strField = "TM45"
            End If
            'Modify by Amy 2021/10/26 原:NVL(round(sA1L21/A1K10,2),0.0) Offical,NVL(round(sA2/A1K10,2),0.0) Disbur,小數位與請款單有差異,應財務要求改為一致-莘
            strSql = "Select * From (Select " & strField & " YourR From TradeMark Where TM01 = '" & adoacc1k0.Fields("CaseNo1") & "' and TM02 = '" & adoacc1k0.Fields("CaseNo2") & "' and TM03 = '" & adoacc1k0.Fields("CaseNo3") & "' and TM04 = '" & adoacc1k0.Fields("CaseNo4") & "')," & _
                        "(" & strSql & ")"
              
            strSQL2 = "Select tm01,tm02,tm03,tm04,Min(A1K01||A1K02) M,TM12 AppNo,TM11 FilingD From Acc1K0,TradeMark Where  TM01='" & adoacc1k0.Fields("CaseNo1") & "' and TM02='" & adoacc1k0.Fields("CaseNo2") & "'  and TM03='" & adoacc1k0.Fields("CaseNo3") & "'  and TM04='" & adoacc1k0.Fields("CaseNo4") & "' " & _
                            "and TM11 is not null And A1K13=TM01 And A1K14=TM02 And A1K15=TM03 And A1K16=TM04 Group by tm01,tm02,tm03,tm04,tm12,tm11"

        Case "CFL", "FCL", "L", "LIN"
            If Left(Text1, 6) = "Y52269" Then
                strField = "Nvl(LC23,LC16)"
            Else
                strField = "LC23"
            End If
            'Modify by Amy 2021/10/26 原:NVL(round(sA1L21/A1K10,2),0.0) Offical,NVL(round(sA2/A1K10,2),0.0) Disbur,小數位與請款單有差異,應財務要求改為一致-莘
            strSql = "Select * From (Select " & strField & " YourR From LawCase Where LC01 = '" & adoacc1k0.Fields("CaseNo1") & "' and LC02 = '" & adoacc1k0.Fields("CaseNo2") & "' and LC03 = '" & adoacc1k0.Fields("CaseNo3") & "' and LC04 = '" & adoacc1k0.Fields("CaseNo4") & "')," & _
                        "(" & strSql & ")"
              
            strSQL2 = "Select lc01,lc02,lc03,lc04,Min(A1K01||A1K02) M,'' AppNo,'' FilingD From Acc1K0,LawCase Where LC01='" & adoacc1k0.Fields("CaseNo1") & "' and LC02='" & adoacc1k0.Fields("CaseNo2") & "'  and LC03='" & adoacc1k0.Fields("CaseNo3") & "'  and LC04='" & adoacc1k0.Fields("CaseNo4") & "' " & _
                            "And A1K13=LC01 And A1K14=LC02 And A1K15=LC03 And A1K16=LC04 Group by LC01,LC02,LC03,LC04"
        Case Else  '服務
            If Left(Text1, 6) = "Y52269" Then
                strField = "Nvl(SP27,SP28)"
            Else
                strField = "SP27"
            End If
            'Modify by Amy 2021/10/26 原:NVL(round(sA1L21/A1K10,2),0.0) Offical,NVL(round(sA2/A1K10,2),0.0) Disbur,小數位與請款單有差異,應財務要求改為一致-莘
            strSql = "Select * From (Select " & strField & " YourR From ServicePractice Where SP01 = '" & adoacc1k0.Fields("CaseNo1") & "' and SP02 = '" & adoacc1k0.Fields("CaseNo2") & "' and SP03 = '" & adoacc1k0.Fields("CaseNo3") & "' and SP04 = '" & adoacc1k0.Fields("CaseNo4") & "')," & _
                        "(" & strSql & ")"
              
            strSQL2 = "Select sp01,sp02,sp03,sp04,Min(A1K01||A1K02) M,SP11 AppNo,SP10 FilingD From Acc1K0,ServicePractice Where SP01='" & adoacc1k0.Fields("CaseNo1") & "' and SP02='" & adoacc1k0.Fields("CaseNo2") & "'  and SP03='" & adoacc1k0.Fields("CaseNo3") & "'  and SP04='" & adoacc1k0.Fields("CaseNo4") & "' " & _
                             "and SP10 is not null And A1K13=SP01 And A1K14=SP02 And A1K15=SP03 And A1K16=SP04 Group by sp01,sp02,sp03,sp04,sp11,sp10"
    End Select
    intR = 1
    Set Rs = ClsLawReadRstMsg(intR, strSql)
    If intR = 1 Then
        stTmp(0) = "" & Rs.Fields("YourR") '彼所案號
        '比照請款單抓法
        If adoacc1k0.Fields("CaseNo1") = "FCP" Then
            strExc(0) = "Select PA106 From CaseProgress,Patent Where PA01(+)=CP01 And PA02(+)=CP02 And PA03(+)=CP03 And PA04(+)=CP04 And CP60='" & adoacc1k0.Fields("DocNo").Value & "' And CP10='605' And CP01='FCP' and pa76 is not null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         '檢查商標延展
        ElseIf InStr("T,FCT,CFT,TF", adoacc1k0.Fields("CaseNo1")) > 0 Then
            strExc(0) = "Select TM65 From CaseProgress,Trademark Where TM01(+)=CP01 And TM02(+)=CP02 And TM03(+)=CP03 And TM04(+)=CP04 And CP60='" & adoacc1k0.Fields("DocNo").Value & "' And CP10='102' And CP01 in('T','FCT','CFT','TF') and TM33 is not null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        Else
            intI = 0
        End If
        '年費彼所案號
        If intI = 1 Then
            If PUB_GetFCCaseNo(adoacc1k0.Fields("DocNo").Value, strExc(2), True) = True Then
               stTmp(0) = "" & strExc(2)
            Else
               stTmp(0) = "" & RsTemp(0)
            End If
        ElseIf PUB_GetFCCaseNo(adoacc1k0("DocNo"), strExc(2)) = True Then
            stTmp(0) = "" & strExc(2)
        End If
        stTmp(5) = Val("" & Rs.Fields("Offical")) '規費
        stTmp(6) = Val("" & Rs.Fields("Disbur")) '雜費
        'Mark by Amy 2022/06/14 還原回 2021/10/26 前
        'If stA1k32 <> "C" Then stTmp(8) = "" & rs.Fields("Service") 'Add by Amy 2021/10/26 服務費
    End If
      
    '該案號為第一張請款單且申請日有值,則Application No及Filing Date需填入
    intR2 = 1
    Set Rs2 = ClsLawReadRstMsg(intR2, strSQL2)
    If intR2 = 1 Then
        If adoacc1k0.Fields("DocNo") = Left(Rs2.Fields("M"), 9) And Val(adoacc1k0.Fields("A1K02")) = Val(Right(Rs2.Fields("M"), Len(Rs2.Fields("M")) - 9)) Then
            stTmp(1) = "" & Rs2.Fields("AppNo")
            If IsNull(Rs2.Fields("FilingD").Value) Then
                stTmp(2) = ""
            Else
                stTmp(2) = ChangeWStringToWDateString(Rs2.Fields("FilingD"))
            End If
        End If
        '綜合使用，申請號全帶出-婧瑄
        stTmp(3) = "" & Rs2.Fields("AppNo")
        If IsNull(Rs2.Fields("FilingD").Value) Then
            stTmp(4) = ""
        Else
            stTmp(4) = ChangeWStringToWDateString(Rs2.Fields("FilingD"))
        End If
    End If
    Rs.Close
    Rs2.Close
End Sub
'end 2017/02/02

'*************************************************
'  轉成Excel檔案
'  Add by Amy 2013/06/10
'*************************************************
Private Sub ExcelSaveNew_Old()
'Dim xlsAgentPoint As New Excel.Application
'Dim wksrpt As New Worksheet
'Dim rs As New ADODB.Recordset, Rs2 As New ADODB.Recordset
'Dim xlsFileName As String, strSql As String, strSQL2 As String
'Dim i As Integer, intR As Integer, intR2 As Integer, TotalRow As Integer
'Dim strField As String 'Add by Amy 2016/03/31
'
' If Dir(strExcelPath & Left(Text1, 6) & "催款單" & ServerDate & MsgText(43)) = MsgText(601) Then
'    If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
'         MkDir strExcelPath
'    End If
' Else
'      Kill strExcelPath & Left(Text1, 6) & "催款單" & ServerDate & MsgText(43)
' End If
'
' xlsAgentPoint.Workbooks.add
' Set wksrpt = xlsAgentPoint.Worksheets(1)
'  wksrpt.PageSetup.Orientation = xlLandscape '橫印
'
'  With wksrpt.Range("a1:r1")
'       .HorizontalAlignment = xlCenter
'       .VerticalAlignment = xlCenter
'       .WrapText = False
'       .Orientation = 0
'       .AddIndent = False
'       .ShrinkToFit = False
'       .MergeCells = False
'       'Add by Amy 2013/08/01
'        If Left(Text1, 6) = "Y27766" Then
'             .Interior.ColorIndex = 19 '設置儲存格填充色(黃)
'        Else
'            .Interior.Color = RGB(255, 165, 79) '設置儲存格填充色
'        End If
'
'   End With
'
'   '設置儲存格填充色(粉)
'   If Left(Text1, 6) = "Y27766" Then
'        wksrpt.Range("g1:h1").Interior.ColorIndex = 38
'   End If
'   'end 2013/08/01
'
' wksrpt.Columns("a:a").ColumnWidth = 5.5
' wksrpt.Columns("b:b").ColumnWidth = 12
' wksrpt.Columns("c:c").ColumnWidth = 10
' wksrpt.Columns("d:d").ColumnWidth = 10
' wksrpt.Columns("e:e").ColumnWidth = 9
' wksrpt.Columns("f:f").ColumnWidth = 13
' wksrpt.Columns("g:g").ColumnWidth = 13
' wksrpt.Columns("h:h").ColumnWidth = 13
' 'Add by Amy 2013/06/27 Excel多加二欄
' wksrpt.Columns("i:i").ColumnWidth = 13
' wksrpt.Columns("j:j").ColumnWidth = 13
' 'end 2013/06/27
' wksrpt.Columns("k:k").ColumnWidth = 6
' 'Add by Amy 2017/02/02 +客戶編號/名稱
' wksrpt.Columns("l:l").ColumnWidth = 10
' wksrpt.Columns("m:m").ColumnWidth = 10
' 'end 2016/12/27
' wksrpt.Columns("n:n").ColumnWidth = 6
' wksrpt.Columns("o:o").ColumnWidth = 8
' wksrpt.Columns("p:p").ColumnWidth = 8
' wksrpt.Columns("q:q").ColumnWidth = 9
' wksrpt.Columns("r:r").ColumnWidth = 10
'
' wksrpt.Range("a1").Value = "NO."
' wksrpt.Range("a1").HorizontalAlignment = xlCenter
' wksrpt.Range("b1").Value = "Law Firm's Name" & Chr(10) & "<select>"
' wksrpt.Range("b1").HorizontalAlignment = xlCenter
'
' 'Add by Amy 2013/08/01 +特殊代理人顯示
' If Left(Text1, 6) = "Y27766" Then
'    wksrpt.Range("c1").Value = "Date" & Chr(10) & "<mm/dd/yyy>"
' Else
'    wksrpt.Range("c1").Value = "Invoice Date" & Chr(10) & "<mm/dd/yyy>"
' End If
' 'end 2013/08/01
' wksrpt.Range("c1").HorizontalAlignment = xlCenter
' wksrpt.Range("d1").Value = "Invoice Num"
' wksrpt.Range("d1").HorizontalAlignment = xlCenter
' 'Add by Amy 2013/08/01 +特殊代理人顯示
' If Left(Text1, 6) = "Y27766" Then
'    wksrpt.Range("e1").Value = "Your Ref"
'    wksrpt.Range("e1").HorizontalAlignment = xlCenter
'    wksrpt.Range("f1").Value = "Murata's Ref"
'    wksrpt.Range("f1").HorizontalAlignment = xlCenter
' Else
'    wksrpt.Range("e1").Value = "Our Ref"
'    wksrpt.Range("e1").HorizontalAlignment = xlCenter
'    wksrpt.Range("f1").Value = "Your Ref"
'    wksrpt.Range("f1").HorizontalAlignment = xlCenter
' End If
' 'end 2013/08/01
' wksrpt.Range("g1").Value = "Application No." & Chr(10) & "(Only for new case)"
' wksrpt.Range("g1").HorizontalAlignment = xlCenter
' wksrpt.Range("h1").Value = "Filing Date" & Chr(10) & "(Only for new case)"
' wksrpt.Range("h1").HorizontalAlignment = xlCenter
' 'Add by Amy 2013/06/27 Excel多加二欄 為綜合使用，申請號全帶出-婧瑄
' wksrpt.Range("i1").Value = "Application No."
' wksrpt.Range("i1").HorizontalAlignment = xlCenter
' wksrpt.Range("j1").Value = "Filing Date"
' wksrpt.Range("j1").HorizontalAlignment = xlCenter
' 'end 2013/06/27
' wksrpt.Range("k1").Value = "Category" & Chr(10) & "<select>"
' wksrpt.Range("k1").HorizontalAlignment = xlCenter
' '2013/07/22 +備註
'  wksrpt.Range("k1").AddComment
'  wksrpt.Range("k1").Comment.Visible = False
'  wksrpt.Range("k1").Comment.Text Text:="1 : New application" & Chr(10) & "2 : Office Action (filing remarks and/or claim amendment)" & _
'                                                                            Chr(10) & "3 : Office Action (others: notice of publication, search report, etc.)" & Chr(10) & "4 : Issue" & Chr(10) & _
'                                                                            "5 : maintenance" & Chr(10) & "6 : Litigation" & Chr(10) & "Z : Other", Start:=200
' wksrpt.Range("k1").Comment.Shape.ScaleWidth 2.5, 0, 0
' wksrpt.Range("k1").Comment.Shape.ScaleHeight 2, 0, 0
' 'end 2013/07/22
' 'Add by Amy 2017/02/02 +客戶編號/名稱
' wksrpt.Range("l1").Value = "Application No."
' wksrpt.Range("l1").HorizontalAlignment = xlCenter
' wksrpt.Range("m1").Value = "Application Name"
' wksrpt.Range("m1").HorizontalAlignment = xlCenter
' 'end 2016/12/27
' wksrpt.Range("n1").Value = "Currency" & Chr(10) & "<select>"
' wksrpt.Range("n1").HorizontalAlignment = xlCenter
' wksrpt.Range("o1").Value = "Attorney Fee"
' wksrpt.Range("o1").HorizontalAlignment = xlCenter
' wksrpt.Range("p1").Value = "Official Fee"
' wksrpt.Range("p1").HorizontalAlignment = xlCenter
' wksrpt.Range("q1").Value = "Disburesment" & Chr(10) & "Fee"
' wksrpt.Range("q1").HorizontalAlignment = xlCenter
' wksrpt.Range("r1").Value = "Total Fee"
' wksrpt.Range("r1").HorizontalAlignment = xlCenter
'
' 'Modify by Amy 2017/02/02 +客戶編號/名稱
' With adoacc1k0
'    For i = 1 To .RecordCount
'      wksrpt.Range("a" & i + 1).Value = i
'      wksrpt.Range("b" & i + 1).Value = "Tai E"
'      wksrpt.Range("c" & i + 1).Value = ChangeTStringToWDateString(.Fields("A1K02"))
'      wksrpt.Range("d" & i + 1).Value = .Fields("DocNo") 'A1K01
'      wksrpt.Range("e" & i + 1).Value = .Fields("CaseNo1") & "-" & .Fields("CaseNo2")
'      wksrpt.Range("n" & i + 1).Value = .Fields("Curr") 'A1K18
'      wksrpt.Range("r" & i + 1).Formula = Val(.Fields("FAmount"))  'A1K08-A1K31
'      wksrpt.Range("l" & i + 1).Value = "" & .Fields("CusNo") '客戶編號
'      wksrpt.Range("m" & i + 1).Value = "" & .Fields("CusName") '客戶名稱
'
'      '抓取其他資料
'      'Modify By Amy 2013/09/04 簡化所有strSQL2的寫法
'      'Modify by Amy 2016/03/31 +巨京沒彼所案號抓分所案號 (有改要確認GetYourRefNo1是否也改)
'      Select Case .Fields("CaseNo1")
'           Case "CFP", "FCP", "P"   '專利
'              If Left(Text1, 6) = "Y52269" Then
'                    strField = "Nvl(PA77,PA47)"
'              Else
'                    strField = "PA77"
'              End If
'              strSql = "Select * From (Select " & strField & " YourR From Patent Where PA01 = '" & .Fields("CaseNo1") & "' and PA02 = '" & .Fields("CaseNo2") & "' and PA03 = '" & .Fields("CaseNo3") & "' and PA04 = '" & .Fields("CaseNo4") & "')," & _
'                          "(Select NVL(round(sA1L05/A1K10,2),0.0) Offical,NVL(round(sA2/A1K10,2),0.0) Disbur From (Select Sum(A1L05) sA2 From Acc1L0,Acc1J0 Where A1L03=A1J01 and A1L04=A1J02 And A1j03='雜費' And A1L01='" & .Fields("DocNo") & "'), " & _
'                          "(Select sum(A1L05) as sA1L05 From Acc1L0 Where A1L01='" & .Fields("DocNo") & "' And Substr(A1L04,-2)='99'), (Select A1K10 From Acc1K0 Where A1K01='" & .Fields("DocNo") & "'))"
'              'strSQL2 = "Select PA11 AppNo,PA10 FilingD,M From Patent,(Select A1K13,A1K14,A1K15,A1K16,M From Acc1K0,(Select Min(A1K01||A1K02) M From Acc1K0,Patent Where PA01='" & .Fields("CaseNo1") & "' and PA02='" & .Fields("CaseNo2") & "'  and PA03='" & .Fields("CaseNo3") & "'  and PA04='" & .Fields("CaseNo4") & "' and PA10 is not null And A1K13=PA01 And A1K14=PA02 And A1K15=PA03 And A1K16=PA04) Where A1K01=SubStr(M,1,9) And A1K02=SubStr(M,10,7)) Where PA01=A1K13 And PA02=A1K14 And PA03=A1K15 And PA04=A1K16"
'              strSQL2 = "Select pa01,pa02,pa03,pa04,Min(A1K01||A1K02) M,pa11 AppNo,pa10 FilingD From Acc1K0,Patent Where PA01='" & .Fields("CaseNo1") & "' and PA02='" & .Fields("CaseNo2") & "'  and PA03='" & .Fields("CaseNo3") & "'  and PA04='" & .Fields("CaseNo4") & "' " & _
'                              "and PA10 is not null And A1K13=PA01 And A1K14=PA02 And A1K15=PA03 And A1K16=PA04 Group by pa01,pa02,pa03,pa04,pa11,pa10"
'
'           Case "CFT", "FCT", "T", "TF"   '商標
'              If Left(Text1, 6) = "Y52269" Then
'                    strField = "Nvl(TM45,TM34)"
'              Else
'                    strField = "TM45"
'              End If
'              strSql = "Select * From (Select " & strField & " YourR From TradeMark Where TM01 = '" & .Fields("CaseNo1") & "' and TM02 = '" & .Fields("CaseNo2") & "' and TM03 = '" & .Fields("CaseNo3") & "' and TM04 = '" & .Fields("CaseNo4") & "')," & _
'                          "(Select NVL(round(sA1L05/A1K10,2),0.0) Offical,NVL(round(sA2/A1K10,2),0.0) Disbur From (Select Sum(A1L05) sA2 From Acc1L0,Acc1J0 Where A1L03=A1J01 and A1L04=A1J02 And A1j03='雜費' And A1L01='" & .Fields("DocNo") & "'), " & _
'                          "(Select sum(A1L05) as sA1L05 From Acc1L0 Where A1L01='" & .Fields("DocNo") & "' And Substr(A1L04,-2)='99'), (Select A1K10 From Acc1K0 Where A1K01='" & .Fields("DocNo") & "'))"
'              'strSQL2 = "Select TM12 AppNo,TM11 FilingD,M From TradeMark,(Select A1K13,A1K14,A1K15,A1K16,M From Acc1K0,(Select Min(A1K01||A1K02) M From Acc1K0,TradeMark Where  TM01='" & .Fields("CaseNo1") & "' and TM02='" & .Fields("CaseNo2") & "'  and TM03='" & .Fields("CaseNo3") & "'  and TM04='" & .Fields("CaseNo4") & "' and TM11 is not null And A1K13=TM01 And A1K14=TM02 And A1K15=TM03 And A1K16=TM04) Where A1K01=SubStr(M,1,9) And A1K02=SubStr(M,10,7)) Where TM01=A1K13 And TM02=A1K14 And TM03=A1K15 And TM04=A1K16"
'              strSQL2 = "Select tm01,tm02,tm03,tm04,Min(A1K01||A1K02) M,TM12 AppNo,TM11 FilingD From Acc1K0,TradeMark Where  TM01='" & .Fields("CaseNo1") & "' and TM02='" & .Fields("CaseNo2") & "'  and TM03='" & .Fields("CaseNo3") & "'  and TM04='" & .Fields("CaseNo4") & "' " & _
'                              "and TM11 is not null And A1K13=TM01 And A1K14=TM02 And A1K15=TM03 And A1K16=TM04 Group by tm01,tm02,tm03,tm04,tm12,tm11"
'
'           Case "CFL", "FCL", "L", "LIN"  '法務 Modify by Amy 2013/09/27
'              If Left(Text1, 6) = "Y52269" Then
'                    strField = "Nvl(LC23,LC16)"
'              Else
'                    strField = "LC23"
'              End If
'              strSql = "Select * From (Select " & strField & " YourR From LawCase Where LC01 = '" & .Fields("CaseNo1") & "' and LC02 = '" & .Fields("CaseNo2") & "' and LC03 = '" & .Fields("CaseNo3") & "' and LC04 = '" & .Fields("CaseNo4") & "')," & _
'                          "(Select NVL(round(sA1L05/A1K10,2),0.0) Offical,NVL(round(sA2/A1K10,2),0.0) Disbur From (Select Sum(A1L05) sA2 From Acc1L0,Acc1J0 Where A1L03=A1J01 and A1L04=A1J02 And A1j03='雜費' And A1L01='" & .Fields("DocNo") & "'), " & _
'                          "(Select sum(A1L05) as sA1L05 From Acc1L0 Where A1L01='" & .Fields("DocNo") & "' And Substr(A1L04,-2)='99'), (Select A1K10 From Acc1K0 Where A1K01='" & .Fields("DocNo") & "'))"
'              strSQL2 = "Select lc01,lc02,lc03,lc04,Min(A1K01||A1K02) M,'' AppNo,'' FilingD From Acc1K0,LawCase Where LC01='" & .Fields("CaseNo1") & "' and LC02='" & .Fields("CaseNo2") & "'  and LC03='" & .Fields("CaseNo3") & "'  and LC04='" & .Fields("CaseNo4") & "' " & _
'                             "And A1K13=LC01 And A1K14=LC02 And A1K15=LC03 And A1K16=LC04 Group by LC01,LC02,LC03,LC04"
'           Case Else                  '服務
'              If Left(Text1, 6) = "Y52269" Then
'                    strField = "Nvl(SP27,SP28)"
'              Else
'                    strField = "SP27"
'              End If
'              strSql = "Select * From (Select " & strField & " YourR From ServicePractice Where SP01 = '" & .Fields("CaseNo1") & "' and SP02 = '" & .Fields("CaseNo2") & "' and SP03 = '" & .Fields("CaseNo3") & "' and SP04 = '" & .Fields("CaseNo4") & "')," & _
'                          "(Select NVL(round(sA1L05/A1K10,2),0.0) Offical,NVL(round(sA2/A1K10,2),0.0) Disbur From (Select Sum(A1L05) sA2 From Acc1L0,Acc1J0 Where A1L03=A1J01 and A1L04=A1J02 And A1j03='雜費' And A1L01='" & .Fields("DocNo") & "'), " & _
'                          "(Select sum(A1L05) as sA1L05 From Acc1L0 Where A1L01='" & .Fields("DocNo") & "' And Substr(A1L04,-2)='99'), (Select A1K10 From Acc1K0 Where A1K01='" & .Fields("DocNo") & "'))"
'              'strSQL2 = "Select SP11 AppNo,SP10 FilingD,M From ServicePractice,(Select A1K13,A1K14,A1K15,A1K16,M From Acc1K0,(Select Min(A1K01||A1K02) M From Acc1K0,ServicePractice Where SP01='" & .Fields("CaseNo1") & "' and SP02='" & .Fields("CaseNo2") & "'  and SP03='" & .Fields("CaseNo3") & "'  and SP04='" & .Fields("CaseNo4") & "' and SP10 is not null And A1K13=SP01 And A1K14=SP02 And A1K15=SP03 And A1K16=SP04) Where A1K01=SubStr(M,1,9) And A1K02=SubStr(M,10,7)) Where SP01=A1K13 And SP02=A1K14 And SP03=A1K15 And SP04=A1K16"
'              strSQL2 = "Select sp01,sp02,sp03,sp04,Min(A1K01||A1K02) M,SP11 AppNo,SP10 FilingD From Acc1K0,ServicePractice Where SP01='" & .Fields("CaseNo1") & "' and SP02='" & .Fields("CaseNo2") & "'  and SP03='" & .Fields("CaseNo3") & "'  and SP04='" & .Fields("CaseNo4") & "' " & _
'                             "and SP10 is not null And A1K13=SP01 And A1K14=SP02 And A1K15=SP03 And A1K16=SP04 Group by sp01,sp02,sp03,sp04,sp11,sp10"
'      End Select
'      'end 2016/03/31
'
'      intR = 1
'      Set rs = ClsLawReadRstMsg(intR, strSql)
'      If intR = 1 Then
'        wksrpt.Range("f" & i + 1).Value = rs.Fields("YourR")
'         'Added by Morgan 2014/2/17 改比照請款單抓法
'         If adoacc1k0.Fields("CaseNo1") = "FCP" Then
'            strExc(0) = "Select PA106 From CaseProgress,Patent Where PA01(+)=CP01 And PA02(+)=CP02 And PA03(+)=CP03 And PA04(+)=CP04 And CP60='" & adoacc1k0.Fields("DocNo").Value & "' And CP10='605' And CP01='FCP' and pa76 is not null"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         'Add By Sindy 2016/7/14 +檢查商標延展
'         ElseIf InStr("T,FCT,CFT,TF", adoacc1k0.Fields("CaseNo1")) > 0 Then
'            strExc(0) = "Select TM65 From CaseProgress,Trademark Where TM01(+)=CP01 And TM02(+)=CP02 And TM03(+)=CP03 And TM04(+)=CP04 And CP60='" & adoacc1k0.Fields("DocNo").Value & "' And CP10='102' And CP01 in('T','FCT','CFT','TF') and TM33 is not null"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         Else
'            intI = 0
'         '2016/7/14 END
'         End If
'         '年費彼所案號
'         If intI = 1 Then
'            If PUB_GetFCCaseNo(adoacc1k0.Fields("DocNo").Value, strExc(2), True) = True Then
'               wksrpt.Range("f" & i + 1).Value = strExc(2)
'            Else
'               wksrpt.Range("f" & i + 1).Value = "" & RsTemp(0)
'            End If
'         ElseIf PUB_GetFCCaseNo(adoacc1k0("DocNo"), strExc(2)) = True Then
'            wksrpt.Range("f" & i + 1).Value = strExc(2)
'         End If
'         'end 2014/2/17
'
'        wksrpt.Range("p" & i + 1).Formula = Val(rs.Fields("Offical"))
'        wksrpt.Range("q" & i + 1).Formula = Val(rs.Fields("Disbur"))
'        wksrpt.Range("o" & i + 1).Formula = "=r" & i + 1 & "-p" & i + 1 & "-q" & i + 1
'
'        'Add by Amy 2013/08/01 特殊代理人 TotalFee欄位以公式顯示
'        If Left(Text1, 6) = "Y27766" Then
'            wksrpt.Range("o" & i + 1).Copy
'            wksrpt.Range("o" & i + 1).PasteSpecial xlPasteValues, xlPasteSpecialOperationNone, False, False
'            wksrpt.Range("r" & i + 1).Formula = "=IF((o" & i + 1 & "+p" & i + 1 & "+q" & i + 1 & ")=0,"""",(o" & i + 1 & "+p" & i + 1 & "+q" & i + 1 & "))"  '公式
'        End If
'      End If
'
'      '該案號為第一張請款單且申請日有值,則Application No及Filing Date需填入
'      intR2 = 1
'      Set Rs2 = ClsLawReadRstMsg(intR2, strSQL2)
'      If intR2 = 1 Then
'        If .Fields("DocNo") = Left(Rs2.Fields("M"), 9) And Val(.Fields("A1K02")) = Val(Right(Rs2.Fields("M"), Len(Rs2.Fields("M")) - 9)) Then
'           wksrpt.Range("g" & i + 1).Value = Rs2.Fields("AppNo")
'           'Modify by Amy 2013/09/27 +IsNull
'           If IsNull(Rs2.Fields("FilingD").Value) Then
'                wksrpt.Range("h" & i + 1).Value = ""
'           Else
'                wksrpt.Range("h" & i + 1).Value = ChangeWStringToWDateString(Rs2.Fields("FilingD"))
'           End If
'        End If
'        'Add by Amy 2013/06/27 Excel多加二欄 為綜合使用，申請號全帶出-婧瑄
'        wksrpt.Range("i" & i + 1).Value = Rs2.Fields("AppNo")
'        If IsNull(Rs2.Fields("FilingD").Value) Then
'            wksrpt.Range("j" & i + 1).Value = ""
'        Else
'            wksrpt.Range("j" & i + 1).Value = ChangeWStringToWDateString(Rs2.Fields("FilingD"))
'        End If
'        'end 2013/06/27
'      End If
'      If Not .EOF Then
'        .MoveNext
'      End If
'    Next i
'
'    TotalRow = i + 2
'    '合計
'    wksrpt.Range("a" & TotalRow).Value = "TOTAL"
'    wksrpt.Range("n" & TotalRow).Value = "=n2"
'    'Modify by Amy 2017/02/02 +客戶編號/名稱
'    wksrpt.Range("o" & TotalRow).Formula = "=sum(o2:o" & TotalRow - 1 & ")"
'    wksrpt.Range("p" & TotalRow).Formula = "=sum(p2:p" & TotalRow - 1 & ")"
'    wksrpt.Range("q" & TotalRow).Formula = "=sum(q2:q" & TotalRow - 1 & ")"
'    wksrpt.Range("r" & TotalRow).Formula = "=sum(r2:r" & TotalRow - 1 & ")"
'    'end 2016/12/27
' End With
'
' 'Excel格式設定
' With wksrpt.Range("a1:r" & TotalRow)
'    .Font.Name = "新細明體"
'    .Font.Size = 10
'
'    '框線
'    .Borders(xlEdgeLeft).LineStyle = xlContinuous
'    .Borders(xlEdgeTop).LineStyle = xlContinuous
'    .Borders(xlEdgeBottom).LineStyle = xlContinuous
'    .Borders(xlEdgeRight).LineStyle = xlContinuous
'    .Borders(xlInsideVertical).LineStyle = xlContinuous
'    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
' End With
' wksrpt.Range("p1:p" & TotalRow).Borders(xlEdgeLeft).LineStyle = xlDouble
' wksrpt.Range("b1:l" & TotalRow).HorizontalAlignment = xlCenter
' wksrpt.Range("c1:c" & TotalRow).NumberFormatLocal = "m/d/yyyy"
' wksrpt.Range("h1:h" & TotalRow).NumberFormatLocal = "m/d/yyyy" '2013/07/22 改日期格式 原:"m/d/yy"
' wksrpt.Range("j1:j" & TotalRow).NumberFormatLocal = "m/d/yyyy"   '2013/07/22 改日期格式 原:"m/d/yy"
' wksrpt.Range("o1:r" & TotalRow).NumberFormatLocal = "0.00"
' 'end 2016/12/27
'
' 'Add by Amy 2013/08/01 特殊代理人 不需 i (Application NO)j(Filing Date) 欄位,其他不需ghj 欄位-婧瑄
' If Left(Text1, 6) = "Y27766" Then
'    wksrpt.Range("i:j").Delete Shift:=xlToLeft
'    wksrpt.Range("j:k").Delete Shift:=xlToLeft 'Add by Amy 2017/02/02 客戶編號/名稱不顯示-莘
' Else
'    wksrpt.Range("g:h").Delete Shift:=xlToLeft
'    'Modify by Amy 2013/08/02 不需Category欄(原本k欄)-婧瑄
'    wksrpt.Range("h:i").Delete Shift:=xlToLeft
'    'wksrpt.Columns("h").Delete Shift:=xlToLeft '原本的j欄
'    'end 2013/08/02
'    'Add by Amy 2017/02/02 沒下客戶編號條件 客戶編號/名稱欄位顯示-莘
'    If bolShowCus = False Then
'        wksrpt.Range("h:i").Delete Shift:=xlToLeft
'    End If
' End If
'
' rs.Close
' 'Modify by Amy 2016/06/23 +判斷版本
' If Val(xlsAgentPoint.Version) < 12 Then
'    xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & Left(Text1, 6) & "催款單" & ServerDate & MsgText(43), FileFormat:=-4143
' Else
'    xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & Left(Text1, 6) & "催款單" & ServerDate & MsgText(43), FileFormat:=56
' End If
' xlsAgentPoint.Workbooks.Close
' xlsAgentPoint.Quit
' StatusClear
End Sub

'Add by Morgan 2009/4/7
'畫表格
Private Sub printTable()
   Dim ii As Integer
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.DrawWidth = 5
      
      'Added by Morgan 2012/3/15 與電子檔一致
      If bolChina = True Then
         Printer.Line (iXfix + Px(0), iYfix + Py(0))-(iXfix + Px(7), iYfix + Py(4)), , B
         '縱線
         For ii = 1 To 3
            Printer.Line (iXfix + Px(ii), iYfix + Py(0))-(iXfix + Px(ii), iYfix + Py(2))
         Next
         For ii = 4 To 4 '5
            Printer.Line (iXfix + Px(ii), iYfix + Py(0))-(iXfix + Px(ii), iYfix + Py(3))
         Next
         For ii = 5 To 6
            Printer.Line (iXfix + Px(ii), iYfix + Py(0))-(iXfix + Px(ii), iYfix + Py(2))
         Next
         '橫線
         For ii = 1 To 3
            Printer.Line (iXfix + Px(0), iYfix + Py(ii))-(iXfix + Px(7), iYfix + Py(ii))
         Next
         
      Else
      'end 2012/3/15
         '框
         Printer.Line (iXfix + Px(0), iYfix + Py(0))-(iXfix + Px(7), iYfix + Py(2)), , B
         '縱線
         For ii = 1 To 6
            Printer.Line (iXfix + Px(ii), iYfix + Py(0))-(iXfix + Px(ii), iYfix + Py(2))
         Next
         '橫線
         Printer.Line (iXfix + Px(0), iYfix + Py(1))-(iXfix + Px(7), iYfix + Py(1))
      End If
   End If
   
   If bol2Jpg Then
      '2009/6/18 ADD BY SONIA 區分大陸
      If bolChina = True Then
         Picture1.DrawWidth = 2
         '框
         Picture1.Line (Px(0) * douExtRate, Py(0) * douExtRate)-(Px(7) * douExtRate, Py(4) * douExtRate), , B
         '縱線
         For ii = 1 To 3
            Picture1.Line (Px(ii) * douExtRate, Py(0) * douExtRate)-(Px(ii) * douExtRate, Py(2) * douExtRate)
         Next
         '縱線
         For ii = 4 To 4 '5
            Picture1.Line (Px(ii) * douExtRate, Py(0) * douExtRate)-(Px(ii) * douExtRate, Py(3) * douExtRate)
         Next
         For ii = 5 To 6
            Picture1.Line (Px(ii) * douExtRate, Py(0) * douExtRate)-(Px(ii) * douExtRate, Py(2) * douExtRate)
         Next
         '橫線
         For ii = 1 To 3
           Picture1.Line (Px(0) * douExtRate, Py(ii) * douExtRate)-(Px(7) * douExtRate, Py(ii) * douExtRate)
         Next
      Else
      '2009/6/18 END
         Picture1.DrawWidth = 2
         '框
         Picture1.Line (Px(0) * douExtRate, Py(0) * douExtRate)-(Px(7) * douExtRate, Py(2) * douExtRate), , B
         '縱線
         For ii = 1 To 6
            Picture1.Line (Px(ii) * douExtRate, Py(0) * douExtRate)-(Px(ii) * douExtRate, Py(2) * douExtRate)
         Next
         '橫線
         Picture1.Line (Px(0) * douExtRate, Py(1) * douExtRate)-(Px(7) * douExtRate, Py(1) * douExtRate)
      End If
   End If
End Sub

'畫表格
Private Sub printTable2()
   Dim ii As Integer
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.DrawWidth = 5
      '框
      Printer.Line (iXfix + Px(0), iYfix + Py(3))-(iXfix + Px(7), iYfix + Py(4)), , B
   End If
   If bol2Jpg Then
      Picture1.DrawWidth = 2
      '框
      Picture1.Line (Px(0) * douExtRate, Py(3) * douExtRate)-(Px(7) * douExtRate, Py(4) * douExtRate), , B
   End If
End Sub
'表尾
Private Sub PrintFooter()
   
   lngY = Py(2) + 450
   lngX = Px(1)
   strData = "We would appreciate receiving your remittance as soon as possible."
   MyPrint strData, lngX, lngY
   
   lngY = lngY + 275
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      lngX = Px(1) + Printer.TextWidth("We would appreciate receiving your remittance ")
      Printer.Line (iXfix + lngX, iYfix + lngY)-(iXfix + lngX + 1800, iYfix + lngY)
   End If
   If bol2Jpg Then
      lngX = Px(1) * douExtRate + Picture1.TextWidth("We would appreciate receiving your remittance ")
      Picture1.Line (lngX, lngY * douExtRate)-(lngX + Picture1.TextWidth("as soon as possible"), lngY * douExtRate)
   End If
   
   printTable2
   
   strData = "REMITTANCE ADVICE"
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.FontSize = 16
      Printer.FontBold = True
   End If
   If bol2Jpg Then
      Picture1.FontSize = 16 * douExtRate
      Picture1.FontBold = True
   End If
   
   lngY = Py(3) + 2 * TopMG
   lngX = Px(0) + (Px(7) - Px(0) - 3340) / 2
   MyPrint strData, lngX, lngY
   
   lngY = Py(3) + 500
   lngX = Px(0) + LeftMG
   
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.FontSize = 12
   End If
   If bol2Jpg Then
      Picture1.FontSize = 12 * douExtRate
   End If
   strData = "Checks to:"
   MyPrint strData, lngX, lngY
      
   lngY = lngY + 250
   lngX = Px(0) + LeftMG
   strData = "Tai E International Patent & Law Office"
   MyPrint strData, lngX, lngY
      
   lngY = lngY + 250
   lngX = Px(0) + LeftMG
   strData = "P. O. Box 46-478, Taipei 104, Taiwan"
   MyPrint strData, lngX, lngY
   
   lngY = Py(3) + 500
   lngX = Px(0) + 8.5 * TwPerCm
   MyPrint "Bankers:", lngX, lngY
   
   lngY = lngY + 250
   lngX = Px(0) + 8.5 * TwPerCm
   
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.FontBold = False
      Printer.FontSize = 9
   End If
   If bol2Jpg Then
      Picture1.FontBold = False
      Picture1.FontSize = 9 * douExtRate
   End If
      
   MyPrint "Name of Account", lngX, lngY
      
   lngX = Px(0) + 11 * TwPerCm
   strData = ": TAI E INTERNATIONAL PATENT & LAW OFFICE"
   MyPrint strData, lngX, lngY
      
   lngY = lngY + 205
   lngX = Px(0) + 8.5 * TwPerCm
   strData = "Account No."
   MyPrint strData, lngX, lngY
   
   lngX = Px(0) + 11 * TwPerCm
   strData = ": 003007052646"
   MyPrint strData, lngX, lngY
   
   lngY = lngY + 205
   lngX = Px(0) + 8.5 * TwPerCm
   strData = "Name of Bank"
   MyPrint strData, lngX, lngY
      
   lngX = Px(0) + 11 * TwPerCm
   strData = ": BANK OF TAIWAN"
   MyPrint strData, lngX, lngY
   
   lngY = lngY + 205
   lngX = Px(0) + 8.5 * TwPerCm
   strData = "Address"
   MyPrint strData, lngX, lngY
      
   lngX = Px(0) + 11 * TwPerCm
   strData = ": 120, Sec. 1. Chungkings S. Rd. Taipei Taiwan R.O.C."
   MyPrint strData, lngX, lngY
   
   lngY = lngY + 205
   lngX = Px(0) + 8.5 * TwPerCm
   strData = "Swift Code"
   MyPrint strData, lngX, lngY
   
   lngX = Px(0) + 11 * TwPerCm
   strData = ": BKTW TWTP"
   MyPrint strData, lngX, lngY
End Sub
'2009/6/18 ADD BY SONIA
'中文表尾
Private Sub PrintFooter1()
   
   lngY = Py(4) + 450
   lngX = Px(0)
   'Modified by Moran 2018/12/17
   'strData = "Name of Bank : Bank of Taiwan, Depetment of Business(一)"
   strData = ReportSum(71001)
   'end 2018/12/17
   MyPrint strData, lngX, lngY
   
   lngY = lngY + 275
   lngX = Px(0)
   'Modified by Morgan 2018/12/17
   'strData = "Address : 120, Sec. 1, Chungking S. Rd., Taipei, Taiwan, R.O.C."
   strData = ReportSum(72)
   'end 2018/12/17
   MyPrint strData, lngX, lngY
   
   lngY = lngY + 275
   lngX = Px(0)
   'Modified by Morgan 2018/12/17
   'strData = "SWIFT Code : BKTW TWTP"
   strData = ReportSum(73001)
   'end 2018/12/17
   MyPrint strData, lngX, lngY
   
   lngY = lngY + 275
   lngX = Px(0)
   'Modified by Morgan 2018/12/17
   'strData = "Account Name : Tai E International Patent & Law Office"
   strData = ReportSum(85)
   'end 2018/12/17
   MyPrint strData, lngX, lngY
   
   lngY = lngY + 275
   lngX = Px(0)
   'Modified by Moran 2018/12/17
   'strData = "Account No. : 003007052646 (for US currency)"
   strData = ReportSum(74)
   'end 2018/12/17
   MyPrint strData, lngX, lngY
   
   Picture1.FontBold = True
   
   'Modify By Sindy 2021/4/14 Mark,不顯示了
'   lngY = lngY + 500
'   lngX = Px(0)
'   '2009/8/4改銀行帳戶,原為中國工商銀行上海東安路支行,林晉章,1001239101213902786*
'   strData = "銀行：招商銀行北京分行金融街支行"
'   MyPrint strData, lngX, lngY
'
'   lngY = lngY + 300
'   lngX = Px(0)
'   strData = "賬戶名稱：林晉章（人民幣個人賬戶）"
'   MyPrint strData, lngX, lngY
'
'   lngY = lngY + 300
'   lngX = Px(0)
'   'modify by sonia 2018/7/17活存800100603817111改金卡6226 0901 0488 1723
'   'strData = "賬號：800100603817111"
'   strData = "賬號：6226 0901 0488 1723"
'   MyPrint strData, lngX, lngY
   
   lngY = lngY + 300
   lngX = Px(0)
   'strData = "＊　貴公司可將款項匯至本所北京或台灣之銀行賬戶，惟于匯款后請務必將匯款憑證傳真至"
   strData = "※　貴公司于匯款后請務必將匯款憑證傳真至台北所，否則本所無法知悉　貴公司已匯款。"
   MyPrint strData, lngX, lngY
   
   lngY = lngY + 300
   lngX = Px(0)
   'strData = "　　台北所，以利查帳，謝謝合作。(傳真號碼：886 2 25011666)"
   strData = "　　(傳真號碼：886 2 25011666)"
   MyPrint strData, lngX, lngY
   
   Picture1.FontBold = False
End Sub

'Add by Morgan 2006/11/29 中文抬頭
Private Sub PrintHead1()
   Printer.FontSize = 12
   Printer.Font = "細明體"
   
   lngX = iXo + 2090
   '代理人代號
   lngY = iYo + 3470
   Printer.CurrentX = lngX
   Printer.CurrentY = lngY
   Printer.Print "" & adoacc1k0.Fields("FagentNo").Value
   
   '代理人名稱
   lngY = lngY + iRowH
   Printer.CurrentX = lngX
   Printer.CurrentY = lngY
   Printer.Print "" & adoacc1k0.Fields("fa04").Value
            
   '代理人地址
   lngY = lngY + iRowH
   Printer.CurrentX = lngX
   Printer.CurrentY = lngY
   Pub_SmartPrint "" & adoacc1k0("fa17"), lngX, lngY, 75, CLng(iRowH)
         
   lngX = iXo + 5275
   lngY = iYo + 6130
   '帳款日期
   Printer.CurrentX = lngX
   Printer.CurrentY = lngY
   Printer.Print IIf(Me.MaskEdBox1.Text = "___/__/__", "", Format(ChangeTStringToWDateString(Replace(Me.MaskEdBox1.Text, "/", "")), "MM/DD/YY")) & _
                        IIf(Me.MaskEdBox1.Text = "___/__/__" And Me.MaskEdBox2.Text = "___/__/__", "", " ～ ") & _
                        IIf(Me.MaskEdBox2.Text = "___/__/__", "", Format(ChangeTStringToWDateString(Replace(Me.MaskEdBox2.Text, "/", "")), "MM/DD/YY"))
   
   lngX = iXo + 9500
   lngY = iYo + 6130
   Printer.CurrentX = lngX
   Printer.CurrentY = lngY
   '2009/6/3 MODIFY BY SONIA
   'If strCurr <> "N" Then
   '   Printer.Print "幣別：美金"
   'End If
   Select Case strCurr
      'Modify By Sindy 2013/1/18
      Case "NTD"
      Case "USD"
         Printer.Print "幣別：美金"
      Case Else
         Select Case adoacc1k0.Fields("Curr")
            Case "USD"
               Printer.Print "幣別：美金"
            Case "RMB"
               Printer.Print "幣別：人民幣"
         End Select
   End Select
   '2009/6/3 END
End Sub
'Add by Morgan 2009/4/7 中文抬頭
Private Sub PrintHead1A4()
   Dim strFA17 As String 'Add by Amy 2018/11/01
   
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.FontSize = 12
      Printer.Font = "細明體"
   End If
   If bol2Jpg Then
      Picture1.FontSize = 12 * douExtRate
      Picture1.Font = "細明體"
      Picture1.AutoRedraw = True
   End If
   
   'Add by Amy 2018/11/01 +地址有「竹曆退件」字樣不顯示地址
   strFA17 = "" & adoacc1k0.Fields("fa17")
   If InStr(strFA17, "竹曆退件") > 0 Then strFA17 = ""
   'end 2018/11/01
   
   '代理人代號
   lngX = iXo + 950
   lngY = iYo + 2470
   strData = "代理人代號: " & adoacc1k0.Fields("FagentNo").Value
   MyPrint strData, lngX, lngY
      
   '代理人名稱
   lngY = lngY + iRowH
   'Modify by Amy 2014/10/13 原抓中文改中->英->日
   strData = "代理人名稱: " '& adoacc1k0.Fields("fa04").Value
   MyPrint strData, lngX, lngY
   If IsNull(adoacc1k0.Fields("fa04").Value) = False Then
        MyPrint adoacc1k0.Fields("fa04").Value, lngX + 1500, lngY
        lngY = lngY + 300
   ElseIf "" & adoacc1k0.Fields("fa05").Value <> "" Then
        MyPrint adoacc1k0.Fields("fa05").Value, lngX + 1500, lngY
        lngY = lngY + 250
        If IsNull(adoacc1k0.Fields("fa63").Value) = False Then
            MyPrint adoacc1k0.Fields("fa63").Value, lngX + 1500, lngY
            lngY = lngY + 250
        End If
            
        If IsNull(adoacc1k0.Fields("fa64").Value) = False Then
            MyPrint adoacc1k0.Fields("fa64").Value, lngX + 1500, lngY
            lngY = lngY + 250
        End If
            
        If IsNull(adoacc1k0.Fields("fa65").Value) = False Then
            MyPrint adoacc1k0.Fields("fa65").Value, lngX + 1500, lngY
            lngY = lngY + 250
        End If
            
   ElseIf "" & adoacc1k0.Fields("fa06").Value <> "" Then
        MyPrint adoacc1k0.Fields("fa06").Value, lngX + 1500, lngY
        lngY = lngY + 300
   End If
   'end 2014/10/13
                  
   '代理人地址
   lngY = lngY + iRowH
   SmartPrint "代理人地址: " & strFA17, lngX, lngY, 75, CLng(iRowH)  'Modify by Amy 2018/11/01 原:adoacc1k0("fa17")
            
   lngX = iXo + 4120
   lngY = iYo + 4550
   
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.FontSize = 18
   End If
   If bol2Jpg Then
      Picture1.FontSize = 18 * douExtRate
   End If
   
   MyPrint "應收帳款對帳單", lngX, lngY
      
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.FontSize = 12
   End If
   If bol2Jpg Then
      Picture1.FontSize = 12 * douExtRate
   End If
      
   '帳款日期
   lngX = iXo + 3830
   lngY = iYo + 5130
   strData = IIf(Me.MaskEdBox1.Text = "___/__/__", "", Format(ChangeTStringToWDateString(Replace(Me.MaskEdBox1.Text, "/", "")), "MM/DD/YY")) & _
            IIf(Me.MaskEdBox1.Text = "___/__/__" And Me.MaskEdBox2.Text = "___/__/__", "", " ～ ") & _
            IIf(Me.MaskEdBox2.Text = "___/__/__", "", Format(ChangeTStringToWDateString(Replace(Me.MaskEdBox2.Text, "/", "")), "MM/DD/YY"))
   MyPrint "帳款日期：" & strData, lngX, lngY

   lngX = iXo + 9000
   lngY = iYo + 5130
   '2009/6/3 MODIFY BY SONIA
   'If strCurr <> "N" Then
   '   MyPrint "幣別：美金", lngX, lngY
   'End If
   Select Case strCurr
      'Modify By Sindy 2013/1/18
      Case "NTD"
         MyPrint "幣別：台幣", lngX, lngY 'Add By Sindy 2015/6/12
      Case "USD"
         MyPrint "幣別：美金", lngX, lngY
      Case Else
         Select Case adoacc1k0.Fields("Curr")
            Case "NTD"
               MyPrint "幣別：台幣", lngX, lngY 'Add By Sindy 2015/6/12
            Case "USD"
               MyPrint "幣別：美金", lngX, lngY
            Case "RMB"
               MyPrint "幣別：人民幣", lngX, lngY
         End Select
   End Select
   '2009/6/3 END
   
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.FontSize = 10
   End If
   If bol2Jpg Then
      Picture1.FontSize = 10 * douExtRate
   End If
      
   '表頭
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.FontBold = True
   End If
   If bol2Jpg Then
      Picture1.FontBold = True
   End If
   
   lngY = Py(0) + (Py(1) - Py(0) - 240) / 2
   lngX = Px(0) + LeftMG
   MyPrint "帳款日期", lngX, lngY
   
   lngX = Px(1) + LeftMG
   MyPrint "帳單編號", lngX, lngY
   
   lngX = Px(2) + LeftMG
   MyPrint "我方文號", lngX, lngY
   
   lngX = Px(3) + LeftMG
   MyPrint "貴方文號", lngX, lngY
   
   lngX = Px(4) + LeftMG
   MyPrint "應收金額", lngX, lngY
   
   'Modify By Sindy 2013/1/7
'   lngX = Px(5) + LeftMG
'   MyPrint "已收金額", lngX, lngY
   lngX = Px(5) + LeftMG
   MyPrint "案件名稱", lngX, lngY
   lngX = Px(6) + LeftMG
   MyPrint "申請人", lngX, lngY
   '2013/1/7 End
   
   lngY = Py(0) + TopMG / 2
   'Modify By Sindy 2013/1/7
   'lngX = Px(6) + LeftMG
   lngX = Px(7) + LeftMG
   'Sindy 2013/1/7 End
'   If bol2Printer Then
'      Printer.FontSize = 9
'   End If
'   If bol2Jpg Then
'      Picture1.FontSize = 9 * douExtRate
'   End If

   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.FontSize = 12
      Printer.FontBold = False
   End If
   If bol2Jpg Then
      Picture1.FontSize = 12 * douExtRate
      Picture1.FontBold = False
   End If
   printTable
   
End Sub
'Removed by Morgan 2014/2/6 不再使用
''*************************************************
''  抬頭列印
''
''*************************************************
''Modify by Morgan 2006/3/27 改可印混合列高,英文250,其他300
'Private Sub PrintHead()
'
'   Dim lngXo As Integer
'   Dim lngYo As Integer
'
'   iPageNo = iPageNo + 1
'
'   'Modify By Sindy 2011/3/8
'   If CheckSys("" & adoacc1k0.Fields("CaseNo1")) = "2" Or _
'      CheckSys("" & adoacc1k0.Fields("CaseNo1")) = "6" Then
'      If IsNull(adoacc1k0.Fields("fa108").Value) = False Then
'         strCurr = adoacc1k0.Fields("fa108").Value
'      Else
'         strCurr = MsgText(601)
'      End If
'   '2011/3/8 End
'   Else
'      If IsNull(adoacc1k0.Fields("fa43").Value) = False Then
'         strCurr = adoacc1k0.Fields("fa43").Value
'      Else
'         strCurr = MsgText(601)
'      End If
'   End If
'
'   'Add by Morgan 2006/11/30 大陸格式
'   If bolChina Then
'      PrintHead1
'      Exit Sub
'   End If
'
'   Dim lngYPos As Long
'   Dim strCustName As String
'   Dim strSystemName As String
'   Dim strCaseName As String
'   Dim strAppNo As String
'   Dim strCustNo As String
'   Dim strLanguage As String
'
'   Printer.FontSize = 12
'   Printer.Font = "Times New Roman"
'
'   lngXo = 1 * TwPerCm - (Printer.Width - Printer.ScaleWidth) / 2
'   lngYo = (Printer.Height - Printer.ScaleHeight) / 2
'   lngYPos = 1500
'
'
'   Printer.CurrentX = lngXo
'   Printer.CurrentY = lngYPos
'   Printer.Print "Account No:"
'   Printer.CurrentX = lngXo + 1500
'   Printer.CurrentY = lngYPos
'   Printer.Print adoacc1k0.Fields("FagentNo").Value
'
'   lngYPos = lngYPos + 250
'   Printer.CurrentX = lngXo
'   Printer.CurrentY = lngYPos
'   Printer.Print "Date:"
'   Printer.CurrentX = lngXo + 1500
'   Printer.CurrentY = lngYPos
'   Printer.Print Format(AFDate(ServerDate), "mmm. d, yyyy")
'
'   lngYPos = lngYPos + 400
'   Printer.CurrentX = lngXo
'   Printer.CurrentY = lngYPos
'   Printer.Print "Attention:"
'
'   adoquery.CursorLocation = adUseClient
'   adoquery.Open "select fa31 as Lang from fagent where fa01 = '" & Mid(adoacc1k0.Fields("FagentNo").Value, 1, 8) & "' and fa02 = '" & Mid(adoacc1k0.Fields("FagentNo").Value, 9, 1) & "' union " & _
'                 "select cu64 as Lang from customer where cu01 = '" & Mid(adoacc1k0.Fields("FagentNo").Value, 1, 8) & "' and cu02 = '" & Mid(adoacc1k0.Fields("FagentNo").Value, 9, 1) & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoquery.RecordCount <> 0 Then
'      If IsNull(adoquery.Fields("Lang").Value) = False Then
'         strLanguage = adoquery.Fields("Lang").Value
'      Else
'         strLanguage = "2"
'      End If
'   Else
'      strLanguage = "2"
'      strCustName = ""
'   End If
'   adoquery.Close
'
'   Select Case strLanguage
'      Case "1" '中文(中-->英-->日)
'         If IsNull(adoacc1k0.Fields("fa04").Value) = False Then
'            Printer.CurrentX = lngXo + 1500
'            Printer.CurrentY = lngYPos
'            Printer.Print adoacc1k0.Fields("fa04").Value
'            lngYPos = lngYPos + 300
'        ElseIf "" & adoacc1k0.Fields("fa05").Value <> "" Then
'            Printer.CurrentX = lngXo + 1500
'            Printer.CurrentY = lngYPos
'            Printer.Print adoacc1k0.Fields("fa05").Value
'            lngYPos = lngYPos + 250
'            If IsNull(adoacc1k0.Fields("fa63").Value) = False Then
'               Printer.CurrentX = lngXo + 1500
'               Printer.CurrentY = lngYPos
'               Printer.Print adoacc1k0.Fields("fa63").Value
'               lngYPos = lngYPos + 250
'            End If
'            If IsNull(adoacc1k0.Fields("fa64").Value) = False Then
'               Printer.CurrentX = lngXo + 1500
'               Printer.CurrentY = lngYPos
'               Printer.Print adoacc1k0.Fields("fa64").Value
'               lngYPos = lngYPos + 250
'            End If
'            If IsNull(adoacc1k0.Fields("fa65").Value) = False Then
'               Printer.CurrentX = lngXo + 1500
'               Printer.CurrentY = lngYPos
'               Printer.Print adoacc1k0.Fields("fa65").Value
'               lngYPos = lngYPos + 250
'            End If
'        ElseIf "" & adoacc1k0.Fields("fa06").Value <> "" Then
'            Printer.CurrentX = lngXo + 1500
'            Printer.CurrentY = lngYPos
'            Printer.Print adoacc1k0.Fields("fa06").Value
'            lngYPos = lngYPos + 300
'         End If
'
'         If IsNull(adoacc1k0.Fields("fa17").Value) = False Then
'            Printer.CurrentX = lngXo + 1500
'            Printer.CurrentY = lngYPos
'            Pub_SmartPrint adoacc1k0("fa17"), lngXo + 1500, lngYPos
'            lngYPos = lngYPos + 300
'        ElseIf "" & adoacc1k0.Fields("fa32").Value <> "" Or "" & adoacc1k0.Fields("fa18").Value <> "" Then
'            If IsNull(adoacc1k0.Fields("fa32").Value) Then
'               If IsNull(adoacc1k0.Fields("fa18").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print adoacc1k0.Fields("fa18").Value
'                  lngYPos = lngYPos + 250
'               End If
'               If IsNull(adoacc1k0.Fields("fa19").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print adoacc1k0.Fields("fa19").Value
'                  lngYPos = lngYPos + 250
'               End If
'               If IsNull(adoacc1k0.Fields("fa20").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print adoacc1k0.Fields("fa20").Value
'                  lngYPos = lngYPos + 250
'               End If
'               If IsNull(adoacc1k0.Fields("fa21").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print adoacc1k0.Fields("fa21").Value
'                  lngYPos = lngYPos + 250
'               End If
'               If IsNull(adoacc1k0.Fields("fa22").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print adoacc1k0.Fields("fa22").Value
'                  lngYPos = lngYPos + 250
'               End If
'              'Add by Morgan 2011/5/25
'              '英文地址6
'               If IsNull(adoacc1k0.Fields("fa70").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print adoacc1k0.Fields("fa70").Value
'                  lngYPos = lngYPos + 250
'               End If
'            Else
'               Printer.CurrentX = lngXo + 1500
'               Printer.CurrentY = lngYPos
'               Printer.Print adoacc1k0.Fields("fa32").Value
'               lngYPos = lngYPos + 250
'               If IsNull(adoacc1k0.Fields("fa33").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print "" & adoacc1k0.Fields("fa33").Value
'                  lngYPos = lngYPos + 250
'               End If
'               If IsNull(adoacc1k0.Fields("fa34").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print "" & adoacc1k0.Fields("fa34").Value
'                  lngYPos = lngYPos + 250
'               End If
'               If IsNull(adoacc1k0.Fields("fa35").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print "" & adoacc1k0.Fields("fa35").Value
'                  lngYPos = lngYPos + 250
'               End If
'               If IsNull(adoacc1k0.Fields("fa36").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print "" & adoacc1k0.Fields("fa36").Value
'                  lngYPos = lngYPos + 250
'               End If
'            End If
'        ElseIf "" & adoacc1k0.Fields("fa23").Value <> "" Then
'            Printer.CurrentX = lngXo + 1500
'            Printer.CurrentY = lngYPos
'            Pub_SmartPrint adoacc1k0("fa23"), lngXo + 1500, lngYPos
'            lngYPos = lngYPos + 300
'         End If
'
'      Case "2" '英文(英-->中-->日)
'         If IsNull(adoacc1k0.Fields("fa05").Value) = False Then
'            Printer.CurrentX = lngXo + 1500
'            Printer.CurrentY = lngYPos
'            Printer.Print adoacc1k0.Fields("fa05").Value
'            lngYPos = lngYPos + 250
'            If IsNull(adoacc1k0.Fields("fa63").Value) = False Then
'               Printer.CurrentX = lngXo + 1500
'               Printer.CurrentY = lngYPos
'               Printer.Print adoacc1k0.Fields("fa63").Value
'               lngYPos = lngYPos + 250
'            End If
'            If IsNull(adoacc1k0.Fields("fa64").Value) = False Then
'               Printer.CurrentX = lngXo + 1500
'               Printer.CurrentY = lngYPos
'               Printer.Print adoacc1k0.Fields("fa64").Value
'               lngYPos = lngYPos + 250
'            End If
'            If IsNull(adoacc1k0.Fields("fa65").Value) = False Then
'               Printer.CurrentX = lngXo + 1500
'               Printer.CurrentY = lngYPos
'               Printer.Print adoacc1k0.Fields("fa65").Value
'               lngYPos = lngYPos + 250
'            End If
'        ElseIf "" & adoacc1k0.Fields("fa04").Value <> "" Then
'            Printer.CurrentX = lngXo + 1500
'            Printer.CurrentY = lngYPos
'            Printer.Print adoacc1k0.Fields("fa04").Value
'            lngYPos = lngYPos + 300
'        ElseIf "" & adoacc1k0.Fields("fa06").Value <> "" Then
'            Printer.CurrentX = lngXo + 1500
'            Printer.CurrentY = lngYPos
'            Printer.Print adoacc1k0.Fields("fa06").Value
'            lngYPos = lngYPos + 300
'        End If
'        'POB,英文地址
'        If "" & adoacc1k0.Fields("fa32").Value <> "" Or "" & adoacc1k0.Fields("fa18").Value <> "" Then
'            '英文地址
'            If IsNull(adoacc1k0.Fields("fa32").Value) Then
'               If IsNull(adoacc1k0.Fields("fa18").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print adoacc1k0.Fields("fa18").Value
'                  lngYPos = lngYPos + 250
'               End If
'               If IsNull(adoacc1k0.Fields("fa19").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print adoacc1k0.Fields("fa19").Value
'                  lngYPos = lngYPos + 250
'               End If
'               If IsNull(adoacc1k0.Fields("fa20").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print adoacc1k0.Fields("fa20").Value
'                  lngYPos = lngYPos + 250
'               End If
'               If IsNull(adoacc1k0.Fields("fa21").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print adoacc1k0.Fields("fa21").Value
'                  lngYPos = lngYPos + 250
'               End If
'               If IsNull(adoacc1k0.Fields("fa22").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print adoacc1k0.Fields("fa22").Value
'                  lngYPos = lngYPos + 250
'               End If
'              'Add by Morgan 2011/5/25
'              '英文地址6
'               If IsNull(adoacc1k0.Fields("fa70").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print adoacc1k0.Fields("fa70").Value
'                  lngYPos = lngYPos + 250
'               End If
'            'POB
'            Else
'               Printer.CurrentX = lngXo + 1500
'               Printer.CurrentY = lngYPos
'               Printer.Print adoacc1k0.Fields("fa32").Value
'               lngYPos = lngYPos + 250
'               If IsNull(adoacc1k0.Fields("fa33").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print "" & adoacc1k0.Fields("fa33").Value
'                  lngYPos = lngYPos + 250
'               End If
'               If IsNull(adoacc1k0.Fields("fa34").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print "" & adoacc1k0.Fields("fa34").Value
'                  lngYPos = lngYPos + 250
'               End If
'               If IsNull(adoacc1k0.Fields("fa35").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print "" & adoacc1k0.Fields("fa35").Value
'                  lngYPos = lngYPos + 250
'               End If
'               If IsNull(adoacc1k0.Fields("fa36").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print "" & adoacc1k0.Fields("fa36").Value
'                  lngYPos = lngYPos + 250
'               End If
'            End If
'        ElseIf "" & adoacc1k0.Fields("fa17").Value <> "" Then
'            Printer.CurrentX = lngXo + 1500
'            Printer.CurrentY = lngYPos
'            Pub_SmartPrint adoacc1k0("fa17"), lngXo + 1500, lngYPos
'            lngYPos = lngYPos + 300
'        ElseIf "" & adoacc1k0.Fields("fa23").Value <> "" Then
'            Printer.CurrentX = lngXo + 1500
'            Printer.CurrentY = lngYPos
'            Pub_SmartPrint adoacc1k0("fa23"), lngXo + 1500, lngYPos
'            lngYPos = lngYPos + 300
'        End If
'
'      Case "3" '日文(日-->中-->英)
'         If IsNull(adoacc1k0.Fields("fa06").Value) = False Then
'            Printer.CurrentX = lngXo + 1500
'            Printer.CurrentY = lngYPos
'            Printer.Print adoacc1k0.Fields("fa06").Value
'            lngYPos = lngYPos + 300
'        ElseIf "" & adoacc1k0.Fields("fa04").Value <> "" Then
'            Printer.CurrentX = lngXo + 1500
'            Printer.CurrentY = lngYPos
'            Printer.Print adoacc1k0.Fields("fa04").Value
'            lngYPos = lngYPos + 300
'        Else
'             If IsNull(adoacc1k0.Fields("fa05").Value) = False Then
'                Printer.CurrentX = lngXo + 1500
'                Printer.CurrentY = lngYPos
'                Printer.Print adoacc1k0.Fields("fa05").Value
'                lngYPos = lngYPos + 250
'                If IsNull(adoacc1k0.Fields("fa63").Value) = False Then
'                   Printer.CurrentX = lngXo + 1500
'                   Printer.CurrentY = lngYPos
'                   Printer.Print adoacc1k0.Fields("fa63").Value
'                   lngYPos = lngYPos + 250
'                End If
'                If IsNull(adoacc1k0.Fields("fa64").Value) = False Then
'                   Printer.CurrentX = lngXo + 1500
'                   Printer.CurrentY = lngYPos
'                   Printer.Print adoacc1k0.Fields("fa64").Value
'                   lngYPos = lngYPos + 250
'                End If
'                If IsNull(adoacc1k0.Fields("fa65").Value) = False Then
'                   Printer.CurrentX = lngXo + 1500
'                   Printer.CurrentY = lngYPos
'                   Printer.Print adoacc1k0.Fields("fa65").Value
'                   lngYPos = lngYPos + 250
'                End If
'            End If
'         End If
'
'         If IsNull(adoacc1k0.Fields("fa23").Value) = False Then
'            Printer.CurrentX = lngXo + 1500
'            Printer.CurrentY = lngYPos
'            Pub_SmartPrint adoacc1k0("fa23"), lngXo + 1500, lngYPos
'            lngYPos = lngYPos + 300
'        ElseIf "" & adoacc1k0.Fields("fa17").Value <> "" Then
'            Printer.CurrentX = lngXo + 1500
'            Printer.CurrentY = lngYPos
'            Pub_SmartPrint adoacc1k0("fa17"), lngXo + 1500, lngYPos
'            lngYPos = lngYPos + 300
'        Else
'            If IsNull(adoacc1k0.Fields("fa32").Value) Then
'               If IsNull(adoacc1k0.Fields("fa18").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print adoacc1k0.Fields("fa18").Value
'                  lngYPos = lngYPos + 250
'               End If
'               If IsNull(adoacc1k0.Fields("fa19").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print adoacc1k0.Fields("fa19").Value
'                  lngYPos = lngYPos + 250
'               End If
'               If IsNull(adoacc1k0.Fields("fa20").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print adoacc1k0.Fields("fa20").Value
'                  lngYPos = lngYPos + 250
'               End If
'               If IsNull(adoacc1k0.Fields("fa21").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print adoacc1k0.Fields("fa21").Value
'                  lngYPos = lngYPos + 250
'               End If
'               If IsNull(adoacc1k0.Fields("fa22").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print adoacc1k0.Fields("fa22").Value
'                  lngYPos = lngYPos + 250
'               End If
'               'Add by Morgan 2011/5/25
'               '英文地址6
'               If IsNull(adoacc1k0.Fields("fa70").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print adoacc1k0.Fields("fa70").Value
'                  lngYPos = lngYPos + 250
'               End If
'            Else
'               Printer.CurrentX = lngXo + 1500
'               Printer.CurrentY = lngYPos
'               Printer.Print adoacc1k0.Fields("fa32").Value
'               lngYPos = lngYPos + 250
'               If IsNull(adoacc1k0.Fields("fa33").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print "" & adoacc1k0.Fields("fa33").Value
'                  lngYPos = lngYPos + 250
'               End If
'               If IsNull(adoacc1k0.Fields("fa34").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print "" & adoacc1k0.Fields("fa34").Value
'                  lngYPos = lngYPos + 250
'               End If
'               If IsNull(adoacc1k0.Fields("fa35").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print "" & adoacc1k0.Fields("fa35").Value
'                  lngYPos = lngYPos + 250
'               End If
'               If IsNull(adoacc1k0.Fields("fa36").Value) = False Then
'                  Printer.CurrentX = lngXo + 1500
'                  Printer.CurrentY = lngYPos
'                  Printer.Print "" & adoacc1k0.Fields("fa36").Value
'                  lngYPos = lngYPos + 250
'               End If
'            End If
'         End If
'   End Select
'
'   strExc(1) = IIf(Me.MaskEdBox1.Text = "___/__/__", "", Format(ChangeTStringToWDateString(Replace(Me.MaskEdBox1.Text, "/", "")), "MM/DD/YY")) & _
'               IIf(Me.MaskEdBox1.Text = "___/__/__" And Me.MaskEdBox2.Text = "___/__/__", "", " ～ ") & _
'               IIf(Me.MaskEdBox2.Text = "___/__/__", "", Format(ChangeTStringToWDateString(Replace(Me.MaskEdBox2.Text, "/", "")), "MM/DD/YY"))
'
'   '列印日期條件
'   Printer.CurrentX = 8500
'   Printer.CurrentY = 4200
'   Printer.Print strExc(1)
'End Sub

'Add by Morgan 2009/4/7
Private Sub PrintHeadA4()
   Dim lngXo As Long
   Dim lngYo As Long
   
   iPageNo = iPageNo + 1
   
   strAttention = ""
   
   'Modify By Sindy 2015/6/12 不需這樣判斷幣別
'   'Modify By Sindy 2011/3/8
'   If CheckSys("" & adoacc1k0.Fields("CaseNo1")) = "2" Or _
'      CheckSys("" & adoacc1k0.Fields("CaseNo1")) = "6" Then
'      If IsNull(adoacc1k0.Fields("fa108").Value) = False Then
'         strCurr = adoacc1k0.Fields("fa108").Value
'      Else
'         strCurr = MsgText(601)
'      End If
'   '2011/3/8 End
'   Else
'      If IsNull(adoacc1k0.Fields("fa43").Value) = False Then
'         strCurr = adoacc1k0.Fields("fa43").Value
'      Else
'         strCurr = MsgText(601)
'      End If
'   End If
   
   '大陸格式
   If bolChina Then
      PrintHead1A4
      Exit Sub
   End If
   
   Dim lngYPos As Long
   Dim strCustName As String
   Dim strSystemName As String
   Dim strCaseName As String
   Dim strAppNo As String
   Dim strCustNo As String
   Dim strLanguage As String
   'Add by Amy 2018/10/31
   Dim strFA17 As String, strFA18 As String, strFA19 As String, strFA20 As String, strFA21 As String, strFA22 As String, strFA70 As String, strFA23 As String
   Dim strFA32 As String, strFA33 As String, strFA34 As String, strFA35 As String, strFA36 As String

   
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.FontSize = 12
      Printer.Font = "Times New Roman"
   End If
   
   If bol2Jpg Then
      Picture1.FontSize = 12 * douExtRate
      Picture1.Font = "Times New Roman"
      Picture1.AutoRedraw = True
   End If
   
   lngXo = Px(0)
   lngYo = (Printer.Height - Printer.ScaleHeight) / 2
   lngYPos = 4.2 * TwPerCm - lngYo
      
   MyPrint "Account No:", lngXo, lngYPos
   MyPrint "" & adoacc1k0.Fields("FagentNo").Value, lngXo + 1500, lngYPos
   lngYPos = lngYPos + 250
   
   MyPrint "Date:", lngXo, lngYPos
   MyPrint Format(AFDate(ServerDate), "mmm. d, yyyy"), lngXo + 1500, lngYPos
   lngYPos = lngYPos + 400

   MyPrint "Attention:", lngXo, lngYPos
   
   strAttention = "Attention:  Accounts Department" & vbCrLf
   
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select fa31 as Lang from fagent where fa01 = '" & Mid(adoacc1k0.Fields("FagentNo").Value, 1, 8) & "' and fa02 = '" & Mid(adoacc1k0.Fields("FagentNo").Value, 9, 1) & "' union " & _
                 "select cu64 as Lang from customer where cu01 = '" & Mid(adoacc1k0.Fields("FagentNo").Value, 1, 8) & "' and cu02 = '" & Mid(adoacc1k0.Fields("FagentNo").Value, 9, 1) & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields("Lang").Value) = False Then
         strLanguage = adoquery.Fields("Lang").Value
      Else
         strLanguage = "2"
      End If
   Else
      strLanguage = "2"
      strCustName = ""
   End If
   adoquery.Close
   'Modify by Amy 2018/11/01 +地址有「竹曆退件」字樣不顯示地址,地址改為變數判斷
   strFA17 = "" & adoacc1k0.Fields("fa17").Value
   strFA18 = "" & adoacc1k0.Fields("fa18").Value: strFA19 = "" & adoacc1k0.Fields("fa19").Value: strFA20 = "" & adoacc1k0.Fields("fa20").Value
   strFA21 = "" & adoacc1k0.Fields("fa21").Value: strFA22 = "" & adoacc1k0.Fields("fa22").Value: strFA70 = "" & adoacc1k0.Fields("fa70").Value
   strFA23 = "" & adoacc1k0.Fields("fa23").Value
   strFA32 = "" & adoacc1k0.Fields("fa32").Value: strFA33 = "" & adoacc1k0.Fields("fa33").Value: strFA34 = "" & adoacc1k0.Fields("fa34").Value
   strFA35 = "" & adoacc1k0.Fields("fa35").Value: strFA36 = "" & adoacc1k0.Fields("fa36").Value
   If InStr(strFA17, "竹曆退件") > 0 Then strFA17 = ""
   If InStr(strFA18 & strFA19 & strFA20 & strFA21 & strFA22 & strFA70, "竹曆退件") > 0 Then
        strFA18 = "": strFA19 = "": strFA20 = "": strFA21 = "": strFA22 = "": strFA70 = ""
   End If
   If InStr(strFA23, "竹曆退件") > 0 Then strFA23 = ""
   If InStr(strFA32 & strFA33 & strFA34 & strFA35 & strFA36, "竹曆退件") > 0 Then
        strFA32 = "": strFA33 = "": strFA34 = "": strFA35 = "": strFA36 = ""
   End If
   'end 2018/11/01
   Select Case strLanguage
      Case "1" '中文(中-->英-->日)
         If IsNull(adoacc1k0.Fields("fa04").Value) = False Then
            MyPrint adoacc1k0.Fields("fa04").Value, lngXo + 1500, lngYPos
            strAttention = strAttention & Space(12) & adoacc1k0.Fields("fa04").Value & vbCrLf
            lngYPos = lngYPos + 300
        ElseIf "" & adoacc1k0.Fields("fa05").Value <> "" Then
            MyPrint adoacc1k0.Fields("fa05").Value, lngXo + 1500, lngYPos
            strAttention = strAttention & Space(12) & adoacc1k0.Fields("fa05").Value & vbCrLf
            lngYPos = lngYPos + 250
            If IsNull(adoacc1k0.Fields("fa63").Value) = False Then
               MyPrint adoacc1k0.Fields("fa63").Value, lngXo + 1500, lngYPos
               strAttention = strAttention & Space(12) & adoacc1k0.Fields("fa63").Value & vbCrLf
               lngYPos = lngYPos + 250
            End If
            
            If IsNull(adoacc1k0.Fields("fa64").Value) = False Then
               MyPrint adoacc1k0.Fields("fa64").Value, lngXo + 1500, lngYPos
               strAttention = strAttention & Space(12) & adoacc1k0.Fields("fa64").Value & vbCrLf
               lngYPos = lngYPos + 250
            End If
            
            If IsNull(adoacc1k0.Fields("fa65").Value) = False Then
               MyPrint adoacc1k0.Fields("fa65").Value, lngXo + 1500, lngYPos
               strAttention = strAttention & Space(12) & adoacc1k0.Fields("fa65").Value & vbCrLf
               lngYPos = lngYPos + 250
            End If
            
        ElseIf "" & adoacc1k0.Fields("fa06").Value <> "" Then
            MyPrint adoacc1k0.Fields("fa06").Value, lngXo + 1500, lngYPos
            strAttention = strAttention & Space(12) & adoacc1k0.Fields("fa06").Value & vbCrLf
            lngYPos = lngYPos + 300
         End If
         '地址
         If strFA17 <> MsgText(601) Then
            MyPrint strFA17, lngXo + 1500, lngYPos
            strAttention = strAttention & Space(12) & strFA17 & vbCrLf
            lngYPos = lngYPos + 300
            
         ElseIf strFA32 <> "" Or "" & strFA18 <> "" Then
            'If IsNull(adoacc1k0.Fields("fa32").Value) Then
            If strFA32 = MsgText(601) Then
               If strFA18 <> MsgText(601) Then
                  MyPrint strFA18, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA18 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               
               If strFA19 <> MsgText(601) Then
                  MyPrint strFA19, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA19 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               
               If strFA20 <> MsgText(601) Then
                  MyPrint strFA20, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA20 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               
               If strFA21 <> MsgText(601) Then
                  MyPrint strFA21, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA21 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               
               If strFA22 <> MsgText(601) Then
                  MyPrint strFA22, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA22 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               'Add by Morgan 2011/5/25
               '英文地址6
               If strFA70 <> MsgText(601) Then
                  MyPrint strFA70, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA70 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               
            Else
               MyPrint strFA32, lngXo + 1500, lngYPos
               strAttention = strAttention & Space(12) & strFA32 & vbCrLf
               lngYPos = lngYPos + 250
               
               If strFA33 <> MsgText(601) Then
                  MyPrint strFA33, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA33 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               
               If strFA34 <> MsgText(601) Then
                  MyPrint strFA34, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA34 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               
               If strFA35 <> MsgText(601) Then
                  MyPrint strFA35, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA35 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               
               If strFA36 <> MsgText(601) Then
                  MyPrint adoacc1k0.Fields("fa05").Value, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & adoacc1k0.Fields("fa05").Value & vbCrLf
                  lngYPos = lngYPos + 250
               End If
            End If
            
         ElseIf strFA23 <> MsgText(601) Then
            SmartPrint strFA23, lngXo + 1500, lngYPos
            strAttention = strAttention & Space(12) & strFA23 & vbCrLf
            lngYPos = lngYPos + 300
         End If
         
      Case "2" '英文(英-->中-->日)
         If IsNull(adoacc1k0.Fields("fa05").Value) = False Then
            MyPrint adoacc1k0.Fields("fa05").Value, lngXo + 1500, lngYPos
            strAttention = strAttention & Space(12) & adoacc1k0.Fields("fa05").Value & vbCrLf
            lngYPos = lngYPos + 250
            
            If IsNull(adoacc1k0.Fields("fa63").Value) = False Then
               MyPrint adoacc1k0.Fields("fa63").Value, lngXo + 1500, lngYPos
               strAttention = strAttention & Space(12) & adoacc1k0.Fields("fa63").Value & vbCrLf
               lngYPos = lngYPos + 250
            End If
            
            If IsNull(adoacc1k0.Fields("fa64").Value) = False Then
               MyPrint adoacc1k0.Fields("fa64").Value, lngXo + 1500, lngYPos
               strAttention = strAttention & Space(12) & adoacc1k0.Fields("fa64").Value & vbCrLf
               lngYPos = lngYPos + 250
            End If
            
            If IsNull(adoacc1k0.Fields("fa65").Value) = False Then
               MyPrint adoacc1k0.Fields("fa65").Value, lngXo + 1500, lngYPos
               strAttention = strAttention & Space(12) & adoacc1k0.Fields("fa65").Value & vbCrLf
               lngYPos = lngYPos + 250
            End If
            
         ElseIf "" & adoacc1k0.Fields("fa04").Value <> "" Then
            MyPrint adoacc1k0.Fields("fa04").Value, lngXo + 1500, lngYPos
            strAttention = strAttention & Space(12) & adoacc1k0.Fields("fa04").Value & vbCrLf
            lngYPos = lngYPos + 300
            
         ElseIf "" & adoacc1k0.Fields("fa06").Value <> "" Then
            MyPrint adoacc1k0.Fields("fa06").Value, lngXo + 1500, lngYPos
            strAttention = strAttention & Space(12) & adoacc1k0.Fields("fa06").Value & vbCrLf
            lngYPos = lngYPos + 300
        End If
        'POB,英文地址
        If strFA32 <> "" Or "" & strFA18 <> "" Then
            '英文地址
            'If IsNull(adoacc1k0.Fields("fa32").Value) Then
            If strFA32 = MsgText(601) Then
               If strFA18 <> MsgText(601) Then
                  MyPrint strFA18, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA18 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               If strFA19 <> MsgText(601) Then
                  MyPrint strFA19, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA19 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               If strFA20 <> MsgText(601) Then
                  MyPrint strFA20, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA20 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               If strFA21 <> MsgText(601) Then
                  MyPrint strFA21, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA21 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               If strFA22 <> MsgText(601) Then
                  MyPrint strFA22, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA22 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
              'Add by Morgan 2011/5/25
              '英文地址6
               If strFA70 <> MsgText(601) Then
                  MyPrint strFA70, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA70 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               
            'POB
            Else
               MyPrint strFA32, lngXo + 1500, lngYPos
               strAttention = strAttention & Space(12) & strFA32 & vbCrLf
               lngYPos = lngYPos + 250
               
               If strFA33 <> MsgText(601) Then
                  MyPrint strFA33, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA33 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               If strFA34 <> MsgText(601) Then
                  MyPrint strFA34, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA34 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               
               If strFA35 <> MsgText(601) Then
                  MyPrint strFA35, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA35 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               
               If strFA36 <> MsgText(601) Then
                  MyPrint strFA36, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA36 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
            End If
         ElseIf strFA17 <> "" Then
            SmartPrint strFA17, lngXo + 1500, lngYPos
            strAttention = strAttention & Space(12) & strFA17 & vbCrLf
            lngYPos = lngYPos + 300
         ElseIf strFA23 <> "" Then
            SmartPrint adoacc1k0("fa23"), lngXo + 1500, lngYPos
            strAttention = strAttention & Space(12) & strFA23 & vbCrLf
            lngYPos = lngYPos + 300
         End If
        
      Case "3" '日文(日-->中-->英)
         If IsNull(adoacc1k0.Fields("fa06").Value) = False Then
            MyPrint adoacc1k0.Fields("fa06").Value, lngXo + 1500, lngYPos
            strAttention = strAttention & Space(12) & adoacc1k0.Fields("fa06").Value & vbCrLf
            lngYPos = lngYPos + 300
         ElseIf "" & adoacc1k0.Fields("fa04").Value <> "" Then
            MyPrint adoacc1k0.Fields("fa04").Value, lngXo + 1500, lngYPos
            strAttention = strAttention & Space(12) & adoacc1k0.Fields("fa04").Value & vbCrLf
            lngYPos = lngYPos + 300
         Else
            If IsNull(adoacc1k0.Fields("fa05").Value) = False Then
               MyPrint adoacc1k0.Fields("fa05").Value, lngXo + 1500, lngYPos
               strAttention = strAttention & Space(12) & adoacc1k0.Fields("fa05").Value & vbCrLf
               lngYPos = lngYPos + 250
               If IsNull(adoacc1k0.Fields("fa63").Value) = False Then
                  MyPrint adoacc1k0.Fields("fa63").Value, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & adoacc1k0.Fields("fa63").Value & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               If IsNull(adoacc1k0.Fields("fa64").Value) = False Then
                  MyPrint adoacc1k0.Fields("fa64").Value, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & adoacc1k0.Fields("fa64").Value & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               If IsNull(adoacc1k0.Fields("fa65").Value) = False Then
                  MyPrint adoacc1k0.Fields("fa65").Value, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & adoacc1k0.Fields("fa65").Value & vbCrLf
                  lngYPos = lngYPos + 250
               End If
            End If
         End If
         '地址
         If strFA23 <> MsgText(601) Then
            SmartPrint strFA23, lngXo + 1500, lngYPos
            strAttention = strAttention & Space(12) & strFA23 & vbCrLf
            lngYPos = lngYPos + 300
         ElseIf strFA17 <> "" Then
            SmartPrint strFA17, lngXo + 1500, lngYPos
            strAttention = strAttention & Space(12) & strFA17 & vbCrLf
            lngYPos = lngYPos + 300
         Else
            'If IsNull(adoacc1k0.Fields("fa32").Value) Then
            If strFA32 = MsgText(601) Then
               If strFA18 <> MsgText(601) Then
                  MyPrint strFA18, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA18 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               If strFA19 <> MsgText(601) Then
                  MyPrint strFA19, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA19 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               If strFA20 <> MsgText(601) Then
                  MyPrint strFA20, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA20 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               If strFA21 <> MsgText(601) Then
                  MyPrint strFA21, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA21 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               If strFA22 <> MsgText(601) Then
                  MyPrint strFA22, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA22 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               'Add by Morgan 2011/5/25
               '英文地址6
               If strFA70 <> MsgText(601) Then
                  MyPrint strFA70, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA70 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
            Else
               MyPrint strFA32, lngXo + 1500, lngYPos
               strAttention = strAttention & Space(12) & strFA32 & vbCrLf
               lngYPos = lngYPos + 250
               If strFA33 <> MsgText(601) Then
                  MyPrint strFA33, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA33 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               If strFA34 <> MsgText(601) Then
                  MyPrint strFA34, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA34 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               If strFA35 <> MsgText(601) Then
                  MyPrint strFA35, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA35 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
               If strFA36 <> MsgText(601) Then
                  MyPrint strFA36, lngXo + 1500, lngYPos
                  strAttention = strAttention & Space(12) & strFA36 & vbCrLf
                  lngYPos = lngYPos + 250
               End If
            End If
         End If
   End Select
   'end 2018/11/01
   
   'Added by Morgan 2020/2/19
   '下列請款對象若有財務編號也要印(專利->商標)
   'Modified by Morgan 2022/3/17 +Y55666--Ryan
   If InStr("Y25061000,Y25061010,Y25061030,Y55363000,Y25061020,Y25061050,Y55666000", adoacc1k0.Fields("FagentNo").Value) > 0 Then
      strData = PUB_GetACCNO(adoacc1k0.Fields("FagentNo").Value)
      If strData <> "" Then
         lngYPos = lngYPos + 300
         MyPrint strData, lngXo, lngYPos
         lngYPos = lngYPos + 300
      End If
   End If
   'end 2020/2/19
   
   strData = "STATEMENT OF ACCOUNT"
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.FontSize = 16
      Printer.FontBold = True
   End If
   If bol2Jpg Then
      Picture1.FontSize = 16 * douExtRate
      Picture1.FontBold = True
   End If
   
   lngY = Py(0) - 0.2 * TwPerCm - 320
   lngX = Px(0) + (Px(7) - Px(0) - 3780) / 2
   
   MyPrint strData, lngX, lngY
   
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.FontSize = 12
      Printer.FontBold = False
   End If
   If bol2Jpg Then
      Picture1.FontSize = 12 * douExtRate
      Picture1.FontBold = False
   End If
   
   strData = IIf(Me.MaskEdBox1.Text = "___/__/__", "", Format(ChangeTStringToWDateString(Replace(Me.MaskEdBox1.Text, "/", "")), "MM/DD/YY")) & _
               IIf(Me.MaskEdBox1.Text = "___/__/__" And Me.MaskEdBox2.Text = "___/__/__", "", " ～ ") & _
               IIf(Me.MaskEdBox2.Text = "___/__/__", "", Format(ChangeTStringToWDateString(Replace(Me.MaskEdBox2.Text, "/", "")), "MM/DD/YY"))
   
   lngY = Py(0) - 0.2 * TwPerCm - 240
   lngX = Px(5)
   MyPrint strData, lngX, lngY
      
   '表頭
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.FontBold = True
   End If
   If bol2Jpg Then
      Picture1.FontBold = True
   End If
   
   lngY = Py(0) + (Py(1) - Py(0) - 240) / 2
   lngX = Px(0) + LeftMG
   MyPrint "DATE", lngX, lngY
   
   lngX = Px(1) + LeftMG
   MyPrint "DEBIT NO.", lngX, lngY
   
   lngX = Px(2) + LeftMG
   MyPrint "O/REF. NO.", lngX, lngY
   
   lngX = Px(3) + LeftMG
   MyPrint "Y/REF. NO.", lngX, lngY
   
   lngX = Px(4) + LeftMG
   MyPrint "CREDIT", lngX, lngY
   
   lngX = Px(5) + LeftMG
   MyPrint "AMOUNT", lngX, lngY
   
   lngY = Py(0) + TopMG / 2
   lngX = Px(6) + LeftMG
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.FontSize = 9
   End If
   If bol2Jpg Then
      Picture1.FontSize = 9 * douExtRate
   End If
   MyPrint "DAYS SINCE", lngX, lngY
   
   lngY = Py(0) + TopMG / 2 + Printer.TextHeight("a")
   lngX = Px(6) + LeftMG
   MyPrint "DN. SENT", lngX, lngY
   
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.FontSize = 12
      Printer.FontBold = False
   End If
   If bol2Jpg Then
      Picture1.FontSize = 12 * douExtRate
      Picture1.FontBold = False
   End If
   printTable
End Sub

'Removed by Morgan 2014/2/6 不再使用
''Add by Morgan 2006/11/29
'' 列印大陸格式合計
'Private Sub PrintSum1()
'   '應收總計
'   lngX = iXo + 9500
'   lngY = iYo + 11900
'   strAmount = Format(douTAmount, FDollar)
'   intLength = Printer.TextWidth(strAmount)
'   Printer.CurrentX = lngX - intLength
'   Printer.CurrentY = lngY
'   Printer.Print strAmount
'   '已收總計
'   If douRAmount > 0 Then
'      lngX = iXo + 11000
'      lngY = iYo + 11900
'      strAmount = Format(douRAmount, FDollar)
'      intLength = Printer.TextWidth(strAmount)
'      Printer.CurrentX = lngX - intLength
'      Printer.CurrentY = lngY
'      Printer.Print strAmount
'   End If
'
'   '應收餘額
'   lngX = iXo + 8550
'   lngY = iYo + 12450
'   Printer.CurrentX = lngX
'   Printer.CurrentY = lngY
'   strAmount = Format(douAmount, FDollar)
'   Select Case strCurr
'      'Modify By Sindy 2013/1/18
'      Case "NTD"
'         Printer.Print "NTD " & strAmount
'      '2009/6/3 MODIFY BY SONIA
'      'Case Else
'         'Printer.Print "USD " & strAmount
'      Case "USD"
'         Printer.Print "USD " & strAmount
'      Case Else
'         Printer.Print m_DNCurr & " " & strAmount
'      '2009/6/3 END
'   End Select
'End Sub

'Add by Morgan 2009/4/7
' 列印大陸格式合計
Private Sub PrintSum1A4()
   '總計
   lngY = iYo + 11250
   lngX = 4500
   MyPrint "總　　　　計", lngX, lngY, True
   
   '應收
   lngY = iYo + 11250
   lngX = Px(5) - LeftMG
   strAmount = Format(douTAmount, FDollar)
   MyPrint strAmount, lngX, lngY, True
       
   'Modify By Sindy 2013/1/7 Mark
'   '已收
'   If douRAmount > 0 Then
'      lngX = Px(6) - LeftMG
'      lngY = iYo + 11250
'      strAmount = Format(douRAmount, FDollar)
'      MyPrint strAmount, lngX, lngY, True
'   End If
      
   '應收餘額
   lngY = iYo + 11670
   lngX = 4500
   MyPrint "應收帳款餘額", lngX, lngY, True
   
   lngX = Px(5) - LeftMG
   lngY = iYo + 11670
   strAmount = Format(douAmount, FDollar)
   Select Case strCurr
      'Modify By Sindy 2013/1/18
      Case "NTD"
         MyPrint "NTD " & strAmount, lngX, lngY, True
      '2009/6/3 MODIFY BY SONIA
      'Case Else
      '   MyPrint "USD " & strAmount, lngX, lngY, True
      Case "USD"
         MyPrint "USD " & strAmount, lngX, lngY, True
      Case Else
         MyPrint m_DNCurr & " " & strAmount, lngX, lngY, True
      '2009/6/3 END
   End Select
   
   PrintFooter1
End Sub

'Removed by Morgan 2014/3/6 不再使用
'*************************************************
' 合計位置
'
'*************************************************
'Private Sub PrintSum()
'   If bolChina = True Then
'      PrintSum1
'      Exit Sub
'   End If
'
'   Printer.CurrentX = 8600
'   Printer.CurrentY = 10300
'   Select Case strCurr
'      'Modify By Sindy 2013/1/18
'      Case "NTD"
'         Printer.Print "NTD"
'      '2009/6/4 MODIFY BY SONIA
'      'Case Else
'      '   Printer.Print "USD"
'      Case "USD"
'         Printer.Print "USD"
'      Case Else
'         Printer.Print m_DNCurr
'      '2009/6/3 END
'   End Select
'   strAmount = Format(douAmount, FDollar)
'   intLength = Printer.TextWidth(strAmount)
'   Printer.CurrentX = 11000 - intLength
'   Printer.CurrentY = 10300
'   Printer.Print strAmount
'   If douOverDue3 = 0 Then
'      strAmount = ""
'   Else
'      strAmount = Format(douOverDue3, FDollar)
'   End If
'   intLength = Printer.TextWidth(strAmount)
'   Printer.CurrentX = 2500 - intLength
'   'Modify by Morgan 2008/4/22
'   'Printer.CurrentY = 12000
'   Printer.CurrentY = 12000 - 250
'   Printer.Print strAmount
'   If douOverDue2 = 0 Then
'      strAmount = ""
'   Else
'      strAmount = Format(douOverDue2, FDollar)
'   End If
'   intLength = Printer.TextWidth(strAmount)
'   Printer.CurrentX = 4500 - intLength
'   'Modify by Morgan 2008/4/22
'   'Printer.CurrentY = 12000
'   Printer.CurrentY = 12000 - 250
'   Printer.Print strAmount
'   If douOverDue1 = 0 Then
'      strAmount = ""
'   Else
'      strAmount = Format(douOverDue1, FDollar)
'   End If
'   intLength = Printer.TextWidth(strAmount)
'   Printer.CurrentX = 6000 - intLength
'   'Modify by Morgan 2008/4/22
'   'Printer.CurrentY = 12000
'   Printer.CurrentY = 12000 - 250
'   Printer.Print strAmount
'End Sub

'Add by Morgan 2009/4/7
Private Sub PrintSumA4()
   If bolChina = True Then
      PrintSum1A4
      Exit Sub
   End If
   Select Case strCurr
      'Modify By Sindy 2013/1/18
      Case "NTD"
         strExc(0) = "NTD"
      '2009/6/3 MODIFY BY SONIA
      'Case Else
      '   strExc(0) = "USD"
      Case "USD"
         strExc(0) = "USD"
      Case Else
         strExc(0) = m_DNCurr
      '2009/6/3 END
   End Select
   
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.FontBold = True
   End If
   If bol2Jpg Then
      Picture1.FontBold = True
   End If
   
   strData = "TOTAL AMOUNT DUE:      " & strExc(0) & "   " & Format(douAmount, FDollar)
   lngY = Py(2) + 100
   lngX = Px(6) - LeftMG
   MyPrint strData, lngX, lngY, True
      
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.FontBold = False
   End If
   If bol2Jpg Then
      Picture1.FontBold = False
   End If
   PrintFooter
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   'Add by Morgan 2009/4/16
   If MaskEdBox2.Text = MsgText(29) Then
      MsgBox "請輸入請款日期迄日!!", vbExclamation
      MaskEdBox2.SetFocus
      Exit Function
   ElseIf IsDate(ChangeTStringToWDateString(Replace(Me.MaskEdBox2.Text, "/", ""))) = False Then
      MsgBox "日期輸入錯誤!!", vbExclamation
      MaskEdBox2.SetFocus
      Exit Function
   End If
   If bolPromoter = True Then
      If Trim(Text1) = "" Then
         MsgBox "代理人編號不可空白!"
         Text1.SetFocus
         Exit Function
      End If
      'Modify by Amy 2018/05/23 原判斷前8碼
      If Left(Text1, 6) <> Left(Text2, 6) Then
         MsgBox "代理人編號起迄不可不同!"
         Text2.SetFocus
         Exit Function
      End If
   End If
   'end 2009/4/16
   
   
   'Add by Morgan 2009/5/5
   If Text7 = "Y" And Text6 = "" And Text6.Enabled = True Then
      'Modified by Lydia 2016/09/08 改可選擇JPG 或PDF
      'MsgBox "要產生請款單電子檔時，是否存電子檔必須輸入Y！"
      'Modified by Lydia 2024/12/31 只存PDF檔
      'MsgBox "要產生請款單電子檔時，是否存電子檔必須輸入1或2！"
      MsgBox "要產生請款單電子檔時，是否存電子檔必須輸入Y！"
      Text6.SetFocus
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
   FormCheck = False
End Function

'Removed by Morgan 2014/3/6 不再使用
''Add by Morgan 2006/11/30
'Private Sub PrintRow()
'   Dim douDollar As Double
'   Dim douODollar As Double
'   Dim lngYc As Long '大陸格式調整
'
'   If bolChina Then lngYc = 1300
'
'   'Date
'   If IsNull(adoacc1k0.Fields("DocDate").Value) Then
'      strExc(1) = ""
'   Else
'      strExc(1) = Format(AFDate(CADate(adoacc1k0.Fields("DocDate").Value)), "MM/DD/YY")
'   End If
'   If strExc(1) <> "" Then
'      Printer.CurrentX = 1200
'      Printer.CurrentY = 5100 + intCounter * 300 + lngYc
'      Printer.Print strExc(1)
'   End If
'
'   'Debit No.
'   If IsNull(adoacc1k0.Fields("DocNo").Value) Then
'      strExc(1) = ""
'   Else
'      If Len(adoacc1k0.Fields("DocNo").Value) = 10 Then
'         strExc(1) = Mid(adoacc1k0.Fields("DocNo").Value, 3, 6)
'      Else
'         strExc(1) = adoacc1k0.Fields("DocNo").Value
'      End If
'   End If
'   If strExc(1) <> "" Then
'      Printer.CurrentX = 2850
'      Printer.CurrentY = 5100 + intCounter * 300 + lngYc
'      Printer.Print strExc(1)
'   End If
'
'   If bolChina = True Then Printer.FontSize = 10 'Add  by Morgan 2006/11/30
'
'   'O/Ref. No.
'   If IsNull(adoacc1k0.Fields("CaseNo1").Value) Then
'      strExc(1) = ""
'   Else
'      If adoacc1k0.Fields("CaseNo3").Value = "0" And adoacc1k0.Fields("CaseNo4").Value = "00" Then
'         strExc(1) = adoacc1k0.Fields("CaseNo1").Value & "-" & adoacc1k0.Fields("CaseNo2").Value
'      Else
'         strExc(1) = adoacc1k0.Fields("CaseNo1").Value & "-" & adoacc1k0.Fields("CaseNo2").Value & "-" & adoacc1k0.Fields("CaseNo3").Value
'      End If
'   End If
'   If strExc(1) <> "" Then
'      Printer.CurrentX = 4400
'      Printer.CurrentY = 5100 + intCounter * 300 + lngYc
'      Printer.Print strExc(1)
'   End If
'
'   'Y/Ref. No.
'   adoquery.CursorLocation = adUseClient
'   adoquery.Open "select pa77 as Yno from patent where pa01 = '" & adoacc1k0.Fields("CaseNo1").Value & "' and pa02 = '" & adoacc1k0.Fields("CaseNo2").Value & "' and pa03 = '" & adoacc1k0.Fields("CaseNo3").Value & "' and pa04 = '" & adoacc1k0.Fields("CaseNo4").Value & "' union " & _
'                 "select tm45 as Yno from trademark where tm01 = '" & adoacc1k0.Fields("CaseNo1").Value & "' and tm02 = '" & adoacc1k0.Fields("CaseNo2").Value & "' and tm03 = '" & adoacc1k0.Fields("CaseNo3").Value & "' and tm04 = '" & adoacc1k0.Fields("CaseNo4").Value & "' union " & _
'                 "select lc23 as Yno from lawcase where lc01 = '" & adoacc1k0.Fields("CaseNo1").Value & "' and lc02 = '" & adoacc1k0.Fields("CaseNo2").Value & "' and lc03 = '" & adoacc1k0.Fields("CaseNo3").Value & "' and lc04 = '" & adoacc1k0.Fields("CaseNo4").Value & "' union " & _
'                 "select sp27 as Yno from servicepractice where sp01 = '" & adoacc1k0.Fields("CaseNo1").Value & "' and sp02 = '" & adoacc1k0.Fields("CaseNo2").Value & "' and sp03 = '" & adoacc1k0.Fields("CaseNo3").Value & "' and sp04 = '" & adoacc1k0.Fields("CaseNo4").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoquery.RecordCount <> 0 Then
'      If IsNull(adoquery.Fields("Yno").Value) = False Then
'         strExc(1) = adoquery.Fields("Yno").Value
'      End If
'   End If
'   adoquery.Close
'
'   'Added by Morgan 2014/2/17 改比照請款單抓法
'   If adoacc1k0.Fields("CaseNo1") = "FCP" Then
'      strExc(0) = "Select PA106 From CaseProgress,Patent Where PA01(+)=CP01 And PA02(+)=CP02 And PA03(+)=CP03 And PA04(+)=CP04 And CP60='" & adoacc1k0.Fields("DocNo").Value & "' And CP10='605' And CP01='FCP' and pa76 is not null"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      '年費彼所案號
'      If intI = 1 Then
'         If PUB_GetFCCaseNo(adoacc1k0.Fields("DocNo").Value, strExc(2), True) = True Then
'            strExc(1) = strExc(2)
'         Else
'           strExc(1) = "" & RsTemp(0)
'         End If
'      ElseIf PUB_GetFCCaseNo(adoacc1k0("DocNo"), strExc(2)) = True Then
'         strExc(1) = strExc(2)
'      End If
'   ElseIf PUB_GetFCCaseNo(adoacc1k0("DocNo"), strExc(2)) = True Then
'      strExc(1) = strExc(2)
'   End If
'   'end 2014/2/17
'
'   If strExc(1) <> "" Then
'      Printer.CurrentX = 6270
'      Printer.CurrentY = 5100 + intCounter * 300 + lngYc
'      Printer.Print strExc(1)
'   End If
'
'   Printer.FontSize = 12
'
'   'Credit
'   strExc(1) = ""
'   If IsNull(adoacc1k0.Fields("OAmount").Value) = False And adoacc1k0.Fields("OAmount").Value <> 0 Then
'      Select Case strCurr
'         'Modify By Sindy 2013/1/18
'         Case "NTD"
'            douODollar = Format(adoacc1k0.Fields("OAmount").Value * adoacc1k0.Fields("Rate").Value, FDollar)
'         '2009/6/4 MODIFY BY SONIA
'         '2013/10/29 modify by sonia 還原 X10211363EUR請款
'         Case Else
'            douODollar = Format(adoacc1k0.Fields("OAmount").Value, FDollar)
'         'Case "USD"
'         '   douODollar = Format(adoacc1k0.Fields("OAmount").Value, FDollar)
'         'Case Else
'         '   Select Case adoacc1k0.Fields("Curr")
'         '      Case "USD"
'         '         douODollar = Format(adoacc1k0.Fields("OAmount").Value, FDollar)
'         '      Case Else
'         '         m_DNRate = PUB_GetUSXRate_1("" & adoacc1k0.Fields("a1k02").Value, adoacc1k0.Fields("Curr"))
'         '         douODollar = (Format(adoacc1k0.Fields("OAmount").Value * adoacc1k0.Fields("Rate").Value, FDollar)) / m_DNRate
'         '   End Select
'         '2009/6/4 END
'      End Select
'      strAmount = Format(douODollar, FDollar)
'      intLength = Printer.TextWidth(strAmount)
'      strExc(1) = strAmount
'      douAmount = douAmount - Val(douODollar)
'      douRAmount = douRAmount + Val(douODollar) 'Add by Morgan 2006/11/30
'      '92.4.15 MODIFY BY SONIA 若部分收款時, CREDIT資料也會印出應收金額, 故將下列改無CREDIT時才做
'      douDollar = 0
'      '92.4.15 END
''         intDays = CalculateDays(CADate(adoacc1k0.Fields("DocDate").Value), ServerDate)
''         If intDays < 31 Then
''            douOverDue1 = douOverDue1 - Val(adoacc1k0.Fields("OAmount").Value)
''         Else
''            If intDays < 61 Then
''              douOverDue2 = douOverDue2 - Val(adoacc1k0.Fields("OAmount").Value)
''            Else
''              douOverDue3 = douOverDue3 - Val(adoacc1k0.Fields("OAmount").Value)
''            End If
''         End If
'   Else
'      douODollar = 0
'   End If
'
'   If strExc(1) <> "" Then
'      'Modify by Morgan 2006/11/30
'      'Printer.CurrentX = 9500 - intLength
'      If bolChina Then
'         Printer.CurrentX = 11000 - intLength
'      Else
'         Printer.CurrentX = 9500 - intLength
'      End If
'      'end 2006/11/30
'      Printer.CurrentY = 5100 + intCounter * 300 + lngYc
'      Printer.Print strExc(1)
'   End If
'
'   adoquery.CursorLocation = adUseClient
'   '92.4.15 MODIFY BY SONIA 若部分收款時, CREDIT資料也會印出應收金額, 故將下列改無CREDIT時才做
'   If Not (IsNull(adoacc1k0.Fields("OAmount").Value) = False And adoacc1k0.Fields("OAmount").Value <> 0) Then
'   '92.4.15 END
'      'adoquery.Open "select sum(a1k08) from acc1k0 where substr(a1k01, 1, 8) = '" & Mid(adoacc1k0.Fields("DocNo").Value, 1, 8) & "' and length(a1k01) = 10 union " & _
'      '              "select sum(a1k08) from acc1k0 where a1k01 = '" & adoacc1k0.Fields("DocNo").Value & "' and length(a1k01) = 9", adoTaie, adOpenStatic, adLockReadOnly
'      '92.4.14 MODIFY BY SONIA(a1k08 - nvl(a1k06, 0))
'      '93.6.1 MODIFY BY SONIA 扣除折讓金額
'      'adoquery.Open "select sum(a1k08) from acc1k0 where substr(a1k01, 1, 8) = '" & Mid(adoacc1k0.Fields("DocNo").Value, 1, 8) & "' and A1K02 < 920201 and a1k29 is null union " & _
'      '              "select sum(a1k08) from acc1k0 where a1k01 = '" & adoacc1k0.Fields("DocNo").Value & "' and A1K02 > 920201 and a1k29 is null", adoTaie, adOpenStatic, adLockReadOnly
'      '93.7.15 MODIFY BY SONIA
'      'adoquery.Open "select sum((a1k08 - nvl(a1k06, 0))) from acc1k0 where substr(a1k01, 1, 8) = '" & Mid(adoacc1k0.Fields("DocNo").Value, 1, 8) & "' and A1K02 < 920201 and a1k29 is null union " & _
'      '              "select sum((a1k08 - nvl(a1k06, 0))) from acc1k0 where a1k01 = '" & adoacc1k0.Fields("DocNo").Value & "' and A1K02 > 920201 and a1k29 is null", adoTaie, adOpenStatic, adLockReadOnly
'      '2007/11/7 MODIFY BY SONIA 舊系統之作廢或銷帳都不計算 X003589500 因 X003589501未銷
'      'adoquery.Open "select sum((a1k08 - nvl(a1k06, 0))),sum(a1k11),sum(nvl(a1k06, 0)) from acc1k0 where substr(a1k01, 1, 8) = '" & Mid(adoacc1k0.Fields("DocNo").Value, 1, 8) & "' and A1K02 < 920201 and a1k29 is null union " & _
'      '              "select sum((a1k08 - nvl(a1k06, 0))),sum(a1k11),sum(nvl(a1k06, 0)) from acc1k0 where a1k01 = '" & adoacc1k0.Fields("DocNo").Value & "' and A1K02 > 920201 and a1k29 is null", adoTaie, adOpenStatic, adLockReadOnly
'      '2009/6/4 MODIFY BY SONIA 加A1K31折讓外幣金額
'      'adoquery.Open "select sum((a1k08 - nvl(a1k06, 0))),sum(a1k11),sum(nvl(a1k06, 0)) from acc1k0 where substr(a1k01, 1, 8) = '" & Mid(adoacc1k0.Fields("DocNo").Value, 1, 8) & "' and A1K02 < 920201 and a1k29 is null and a1k12 is null and a1k25 is null union " & _
'      '              "select sum((a1k08 - nvl(a1k06, 0))),sum(a1k11),sum(nvl(a1k06, 0)) from acc1k0 where a1k01 = '" & adoacc1k0.Fields("DocNo").Value & "' and A1K02 > 920201 and a1k29 is null", adoTaie, adOpenStatic, adLockReadOnly
'      '2013/10/3 mdofi by sonia (a1k08 - nvl(a1k06, 0)) as FAmount改為(a1k08 - nvl(a1k31, 0)) as FAmount
'      adoquery.Open "select sum((a1k08 - nvl(a1k31, 0))),sum(a1k11),sum(nvl(a1k06, 0)),sum(nvl(a1k31, 0)) from acc1k0 where substr(a1k01, 1, 8) = '" & Mid(adoacc1k0.Fields("DocNo").Value, 1, 8) & "' and A1K02 < 920201 and a1k29 is null and a1k12 is null and a1k25 is null union " & _
'                    "select sum((a1k08 - nvl(a1k31, 0))),sum(a1k11),sum(nvl(a1k06, 0)),sum(nvl(a1k31, 0)) from acc1k0 where a1k01 = '" & adoacc1k0.Fields("DocNo").Value & "' and A1K02 > 920201 and a1k29 is null", adoTaie, adOpenStatic, adLockReadOnly
'      '2009/6/4 END
'      '2007/11/7 END
'      '93.7.15 END
'      '93.6.1 END
'      '92.4.14 END
'      If adoquery.RecordCount <> 0 Then
'         If IsNull(adoquery.Fields(0).Value) Then
'            douDollar = 0
'         Else
'            Select Case strCurr
'                 'Modify By Sindy 2013/1/18
'                 Case "NTD"
'                  '93.7.15 MODIFY BY SONIA
'                  'douDollar = Val(adoquery.Fields(0).Value) * Val(adoacc1k0.Fields("Rate").Value)
'                  douDollar = Val(adoquery.Fields(1).Value) - Val(adoquery.Fields(2).Value) * Val(adoacc1k0.Fields("Rate").Value)
'                  '93.7.15 END
'               '2009/6/4 MODIFY BY SONIA
'               'Case Else
'               '   douDollar = Val(adoquery.Fields(0).Value)
'               '2013/7/29 MODIFY BY SONIA 再改回
''               Case "USD"
''                  douDollar = Val(adoquery.Fields(0).Value)
''               Case Else
''                  Select Case adoacc1k0.Fields("Curr")
''                     Case "USD"
''                        douDollar = Val(adoquery.Fields(0).Value)
''                     Case Else
''                        m_DNRate = PUB_GetUSXRate_1("" & adoacc1k0.Fields("a1k02").Value, adoacc1k0.Fields("Curr"))
''                        douDollar = (Val(adoquery.Fields(1).Value) - Val(adoquery.Fields(3).Value) * m_DNRate) / m_DNRate
''                  End Select
'               Case Else
'                  douDollar = Val(adoquery.Fields(0).Value)
'               '2009/6/4 END
'            End Select
'         End If
'      Else
'         douDollar = 0
'      End If
'      adoquery.Close
'   '92.4.15 MODIFY BY SONIA
'   End If
'   '92.4.15 END
'      'douAmount = douAmount + Val(douDollar)
'
'   'Amount
'   strExc(1) = ""
'   If douDollar <> 0 Then
'      strAmount = Format(douDollar, FDollar)
'      intLength = Printer.TextWidth(strAmount)
'      strExc(1) = strAmount
'      douAmount = douAmount + douDollar
'      douTAmount = douTAmount + douDollar 'Add by Morgan 2006/11/30
'   End If
'
'   If strExc(1) <> "" Then
'      'Modify by Morgan 2006/11/30
'      'Printer.CurrentX = 11000 - intLength
'      If bolChina Then
'         Printer.CurrentX = 9500 - intLength
'      Else
'         Printer.CurrentX = 11000 - intLength
'      End If
'      'end 2006/11/30
'      Printer.CurrentY = 5100 + intCounter * 300 + lngYc
'      Printer.Print strExc(1)
'   End If
'
'   intDays = CalculateDays(CADate(adoacc1k0.Fields("a1k02").Value), ServerDate)
'   '92.4.15 MODIFY BY SONIA
'   'If intDays < 61 Then
'   '   douOverDue1 = douOverDue1 + douDollar - douODollar
'   'Else
'   '   If intDays < 91 Then
'   '     douOverDue2 = douOverDue2 + douDollar - douODollar
'   '   Else
'   '     douOverDue3 = douOverDue3 + douDollar - douODollar
'   '   End If
'   'End If
'   '天期
'   If intDays > 90 Then
'      douOverDue3 = douOverDue3 + douDollar - douODollar
'   ElseIf intDays > 60 Then
'      douOverDue2 = douOverDue2 + douDollar - douODollar
'   ElseIf intDays > 30 Then
'      douOverDue1 = douOverDue1 + douDollar - douODollar
'   End If
'   '92.4.15 END
'End Sub

'Add by Morgan 2009/4/7
Private Sub PrintRowA4()
   Dim douDollar As Double, douODollar As Double
   Dim bolTQM As Boolean 'Add by Amy 2024/06/14
   
   '2009/6/10 add by sonia 大陸格式
   If bolChina Then
      PrintRow1A4
      Exit Sub
   End If
   
   bolTQM = False 'Add by Amy 2024/06/14
   lngY = Py(1) + TopMG + intCounter * (260 + TopMG)
   
   'Date
   strData = ""
   If Not IsNull(adoacc1k0.Fields("DocDate").Value) Then
      strData = Format(AFDate(CADate(adoacc1k0.Fields("DocDate").Value)), "MM/DD/YY")
   End If
   If strData <> "" Then
      lngX = Px(0) + LeftMG
      MyPrint strData, lngX, lngY
   End If
      
   'Debit No.
   strData = ""
   If Not IsNull(adoacc1k0.Fields("DocNo").Value) Then
      If Len(adoacc1k0.Fields("DocNo").Value) = 10 Then
         strData = Mid(adoacc1k0.Fields("DocNo").Value, 3, 6)
      Else
         strData = adoacc1k0.Fields("DocNo").Value
      End If
   End If
   
   If bolIsBatchInvoice Then strData = strBatchInvoiceNo 'Added by Morgan 2016/2/15 整批列印請款單

   'Modified by Lydia 2019/11/04 "," =>", "
   strDBno = strDBno & strData & ", " 'Added by Lydia 2015/10/19
   
   If strData <> "" Then
      lngX = Px(1) + LeftMG
      MyPrint strData, lngX, lngY
   End If
   
   If bolChina = True Then
      'Modified by Lydia 2016/09/08 +存PDF檔
      If bol2Printer Or bol2Pdf Then
         Printer.FontSize = 10
      End If
      If bol2Jpg Then
         Picture1.FontSize = 10 * douExtRate
      End If
   End If
   
   'O/Ref. No.
   strData = ""
   If Not IsNull(adoacc1k0.Fields("CaseNo1").Value) Then
      If adoacc1k0.Fields("CaseNo3").Value = "0" And adoacc1k0.Fields("CaseNo4").Value = "00" Then
         strData = adoacc1k0.Fields("CaseNo1").Value & "-" & adoacc1k0.Fields("CaseNo2").Value
      Else
         strData = adoacc1k0.Fields("CaseNo1").Value & "-" & adoacc1k0.Fields("CaseNo2").Value & "-" & adoacc1k0.Fields("CaseNo3").Value
      End If
      'Add by Amy 2024/06/18 商標查名請款單由TS或S 案轉入者,在O/REF.NO.之本所案號後再加入"(查名本所案號)" ex:FCT-051898(S-008316)
      If InStr("" & adoacc1k0.Fields("CaseNo1"), "T") > 0 And "" & adoacc1k0.Fields("CaseNo1") <> "TS" Then
         strData = strData & GetTMQCaseNo(False, "" & adoacc1k0.Fields("DocNo"), bolTQM)
      End If
   End If
   
   'Added by Morgan 2016/2/16 整批列印請款單:本所抓第一個請款單號的本所號+",..."
   If bolIsBatchInvoice Then
      strExc(0) = "select a1k13||'-'||a1k14||decode(a1k15||a1k16,'000','','-'||a1k15||'-'||a1k16) from acc1k0 where a1k01='" & adoacc1k0.Fields("a1t02") & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strData = RsTemp(0) & ",..."
      End If
   End If
   'end 2016/2/16
   
   If strData <> "" Then
      lngX = Px(2) + LeftMG
      'Modify by Amy 2024/06/14 商標查名請款單由TS或S 案轉入者,在O/REF.NO.之本所案號後再加入"(查名本所案號)"
      If bolTQM = True Then
         MyPrint strData, lngX, lngY, , 0
      Else
         MyPrint strData, lngX, lngY
      End If
   End If

   'Y/Ref. No.
   strData = ""
   strExc(1) = adoacc1k0.Fields("DocNo").Value
   adoquery.CursorLocation = adUseClient
   'Modified by Morgan 2016/2/16 配合整批列印請款單改語法
   'adoquery.Open "select pa77 as Yno from patent where pa01 = '" & adoacc1k0.Fields("CaseNo1").Value & "' and pa02 = '" & adoacc1k0.Fields("CaseNo2").Value & "' and pa03 = '" & adoacc1k0.Fields("CaseNo3").Value & "' and pa04 = '" & adoacc1k0.Fields("CaseNo4").Value & "' union " & _
                 "select tm45 as Yno from trademark where tm01 = '" & adoacc1k0.Fields("CaseNo1").Value & "' and tm02 = '" & adoacc1k0.Fields("CaseNo2").Value & "' and tm03 = '" & adoacc1k0.Fields("CaseNo3").Value & "' and tm04 = '" & adoacc1k0.Fields("CaseNo4").Value & "' union " & _
                 "select lc23 as Yno from lawcase where lc01 = '" & adoacc1k0.Fields("CaseNo1").Value & "' and lc02 = '" & adoacc1k0.Fields("CaseNo2").Value & "' and lc03 = '" & adoacc1k0.Fields("CaseNo3").Value & "' and lc04 = '" & adoacc1k0.Fields("CaseNo4").Value & "' union " & _
                 "select sp27 as Yno from servicepractice where sp01 = '" & adoacc1k0.Fields("CaseNo1").Value & "' and sp02 = '" & adoacc1k0.Fields("CaseNo2").Value & "' and sp03 = '" & adoacc1k0.Fields("CaseNo3").Value & "' and sp04 = '" & adoacc1k0.Fields("CaseNo4").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
   If bolIsBatchInvoice Then strExc(1) = adoacc1k0.Fields("a1t02").Value
   adoquery.Open "select pa77 as Yno from patent where (pa01,pa02,pa03,pa04) in (select a1k13,a1k14,a1k15,a1k16 from acc1k0 where a1k01='" & strExc(1) & "') union " & _
                 "select tm45 as Yno from trademark where (tm01,tm02,tm03,tm04) in (select a1k13,a1k14,a1k15,a1k16 from acc1k0 where a1k01='" & strExc(1) & "') union " & _
                 "select lc23 as Yno from lawcase where (lc01,lc02,lc03,lc04) in (select a1k13,a1k14,a1k15,a1k16 from acc1k0 where a1k01='" & strExc(1) & "') union " & _
                 "select sp27 as Yno from servicepractice where (sp01,sp02,sp03,sp04) in (select a1k13,a1k14,a1k15,a1k16 from acc1k0 where a1k01='" & strExc(1) & "')", adoTaie, adOpenStatic, adLockReadOnly
   'end 2016/2/16
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields("Yno").Value) = False Then
         strData = adoquery.Fields("Yno").Value
      End If
   End If
   adoquery.Close
   'Add by Amy 2024/06/14 商標查名請款單由TS或S 案轉入者,Y/REF.NO.有值不秀,避免看不出案號 ex:T-245686(TS-001979) -秀玲
   If bolTQM = True Then strData = ""
      
   'Added by Morgan 2014/2/17 改比照請款單抓法
   If adoacc1k0.Fields("CaseNo1") = "FCP" Then
      strExc(0) = "Select PA106 From CaseProgress,Patent Where PA01(+)=CP01 And PA02(+)=CP02 And PA03(+)=CP03 And PA04(+)=CP04 And CP60='" & strExc(1) & "' And CP10='605' And CP01='FCP' and pa76 is not null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'Add By Sindy 2016/7/14 +檢查商標延展
   ElseIf InStr("T,FCT,CFT,TF", adoacc1k0.Fields("CaseNo1")) > 0 Then
      strExc(0) = "Select TM65 From CaseProgress,Trademark Where TM01(+)=CP01 And TM02(+)=CP02 And TM03(+)=CP03 And TM04(+)=CP04 And CP60='" & strExc(1) & "' And CP10='102' And CP01 in('T','FCT','CFT','TF') and TM33 is not null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   Else
      intI = 0
   '2016/7/14 END
   End If
   '有年費/延展彼所案號
   If intI = 1 Then
      If PUB_GetFCCaseNo(strExc(1), strExc(2), True) = True Then
         strData = strExc(2)
      Else
         strData = "" & RsTemp(0)
      End If
   ElseIf PUB_GetFCCaseNo(strExc(1), strExc(2)) = True Then
      strData = strExc(2)
   End If
   'end 2014/2/17
   
   If strData <> "" Then
      If bolIsBatchInvoice Then strData = strData & ",..." 'Added by Morgan 2016/2/15 整批列印請款單
      lngX = Px(3) + LeftMG
      MyPrint strData, lngX, lngY, , 3
   End If
   
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.FontSize = 12
   End If
   If bol2Jpg Then
      Picture1.FontSize = 12 * douExtRate
   End If
      
   'Credit
   strData = ""
   If IsNull(adoacc1k0.Fields("OAmount").Value) = False And adoacc1k0.Fields("OAmount").Value <> 0 Then
      Select Case strCurr
         'Modify By Sindy 2013/1/18
         Case "NTD"
            douODollar = Format(adoacc1k0.Fields("OAmount").Value * adoacc1k0.Fields("Rate").Value, FDollar)
         '2009/6/4 MODIFY BY SONIA
         '2013/10/29 modify by sonia 還原 X10211363EUR請款
         Case Else
            douODollar = Format(adoacc1k0.Fields("OAmount").Value, FDollar)
         'Case "USD"
         '   douODollar = Format(adoacc1k0.Fields("OAmount").Value, FDollar)
         'Case Else
         '   Select Case adoacc1k0.Fields("Curr")
         '      Case "USD"
         '         douODollar = Format(adoacc1k0.Fields("OAmount").Value, FDollar)
         '      Case Else
         '         m_DNRate = PUB_GetUSXRate_1("" & adoacc1k0.Fields("a1k02").Value, adoacc1k0.Fields("Curr"))
         '         douODollar = (Format(adoacc1k0.Fields("OAmount").Value * adoacc1k0.Fields("Rate").Value, FDollar)) / m_DNRate
         '   End Select
         '2009/6/4 END
      End Select
      
      'Added by Morgan 2016/2/16
      If bolIsBatchInvoice Then
         strExc(0) = "select sum(a0z04) from acc1t0,acc1k0,acc0z0,acc0y0 where a1t02='" & adoacc1k0.Fields("a1t02") & "' and a1k01(+)=a1t01 and a0z02(+)= a1k01 and a0y01(+)=a0z01 and a0y02=" & adoacc1k0.Fields("DocDate")
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Select Case strCurr
            Case "NTD"
               douODollar = Format(RsTemp(0).Value * adoacc1k0.Fields("Rate").Value, FDollar)
            Case Else
               douODollar = Format(RsTemp(0).Value, FDollar)
            End Select
         End If
      End If
      'end 2016/2/16
      strAmount = Format(douODollar, FDollar)
      intLength = Printer.TextWidth(strAmount)
      strData = strAmount
      douAmount = douAmount - Val(douODollar)
      douRAmount = douRAmount + Val(douODollar)
      douDollar = 0
   Else
      douODollar = 0
   End If
   
   If strData <> "" Then
      lngX = Px(5) - LeftMG
      MyPrint strData, lngX, lngY, True
   End If
   
   adoquery.CursorLocation = adUseClient
   If Not (IsNull(adoacc1k0.Fields("OAmount").Value) = False And adoacc1k0.Fields("OAmount").Value <> 0) Then
      '2009/6/4 MODIFY BY SONIA 加A1K31折讓外幣金額
      'adoquery.Open "select sum((a1k08 - nvl(a1k06, 0))),sum(a1k11),sum(nvl(a1k06, 0)) from acc1k0 where substr(a1k01, 1, 8) = '" & Mid(adoacc1k0.Fields("DocNo").Value, 1, 8) & "' and A1K02 < 920201 and a1k29 is null and a1k12 is null and a1k25 is null union " & _
      '              "select sum((a1k08 - nvl(a1k06, 0))),sum(a1k11),sum(nvl(a1k06, 0)) from acc1k0 where a1k01 = '" & adoacc1k0.Fields("DocNo").Value & "' and A1K02 > 920201 and a1k29 is null", adoTaie, adOpenStatic, adLockReadOnly
      '2013/10/3 mdofi by sonia (a1k08 - nvl(a1k06, 0)) as FAmount改為(a1k08 - nvl(a1k31, 0)) as FAmount
      'Added by Morgan 2016/2/15 整批列印請款單
      If bolIsBatchInvoice = True Then
         '若部分單號結清時就該整批請款單而言仍為未結清故應收款不排除已結清者
         adoquery.Open "select sum((a1k08 - nvl(a1k31, 0))),sum(a1k11),sum(nvl(a1k06, 0)),sum(nvl(a1k31, 0)) from acc1t0, acc1k0 where a1t02= '" & adoacc1k0.Fields("a1t02") & "' and a1k01(+)=a1t01", adoTaie, adOpenStatic, adLockReadOnly
      Else
      'end 2016/2/15
         adoquery.Open "select sum((a1k08 - nvl(a1k31, 0))),sum(a1k11),sum(nvl(a1k06, 0)),sum(nvl(a1k31, 0)) from acc1k0 where substr(a1k01, 1, 8) = '" & Mid(adoacc1k0.Fields("DocNo").Value, 1, 8) & "' and A1K02 < 920201 and a1k29 is null and a1k12 is null and a1k25 is null union " & _
                       "select sum((a1k08 - nvl(a1k31, 0))),sum(a1k11),sum(nvl(a1k06, 0)),sum(nvl(a1k31, 0)) from acc1k0 where a1k01 = '" & adoacc1k0.Fields("DocNo").Value & "' and A1K02 > 920201 and a1k29 is null", adoTaie, adOpenStatic, adLockReadOnly
      End If 'Added by Morgan 2016/2/15
      '2009/6/4 END
      If adoquery.RecordCount <> 0 Then
         If IsNull(adoquery.Fields(0).Value) Then
            douDollar = 0
         Else
            Select Case strCurr
               'Modify By Sindy 2013/1/18
               Case "NTD"
                  douDollar = Val(adoquery.Fields(1).Value) - Val(adoquery.Fields(2).Value) * Val(adoacc1k0.Fields("Rate").Value)
               '2009/6/4 MODIFY BY SONIA
               '2013/10/29 modify by sonia 還原 X10211363EUR請款
               Case Else
                  douDollar = Val(adoquery.Fields(0).Value)
               'Case "USD"
               '   douDollar = Val(adoquery.Fields(0).Value)
               'Case Else
               '   Select Case adoacc1k0.Fields("Curr")
               '      Case "USD"
               '         douDollar = Val(adoquery.Fields(0).Value)
               '      Case Else
               '         m_DNRate = PUB_GetUSXRate_1("" & adoacc1k0.Fields("a1k02").Value, adoacc1k0.Fields("Curr"))
               '         douDollar = (Val(adoquery.Fields(1).Value) - Val(adoquery.Fields(3).Value) * m_DNRate) / m_DNRate
               '   End Select
               '2009/6/4 END
            End Select
         End If
      Else
         douDollar = 0
      End If
      adoquery.Close
   End If
   
   'Amount
   strData = ""
   If douDollar <> 0 Then
      strAmount = Format(douDollar, FDollar)
      intLength = Printer.TextWidth(strAmount)
      strData = strAmount
      douAmount = douAmount + douDollar
      douTAmount = douTAmount + douDollar
   End If
   
   If strData <> "" Then
      lngX = Px(6) - LeftMG
      MyPrint strData, lngX, lngY, True
   End If
   
   intDays = CalculateDays(CADate(adoacc1k0.Fields("a1k02").Value), ServerDate)
   '天期
   strData = ""
   If intDays > 90 Then
      strData = "+90"
   ElseIf intDays > 60 Then
      strData = "+60"
   ElseIf intDays > 30 Then
      strData = "+30"
   End If
   
   If strData <> "" Then
      lngX = Px(7) - LeftMG
      MyPrint strData, lngX, lngY, True
   End If
End Sub

'2009/6/10 add by sonia '中文格式
Private Sub PrintRow1A4()
   Dim douDollar As Double, douODollar As Double
   Dim bolTQM As Boolean 'Add by Amy 2024/06/14
   
   bolTQM = False 'Add by Amy 2024/06/14
   lngY = Py(1) + TopMG + intCounter * (260 + TopMG)
   
   '帳款日期
   strData = ""
   If Not IsNull(adoacc1k0.Fields("DocDate").Value) Then
      strData = Format(AFDate(CADate(adoacc1k0.Fields("DocDate").Value)), "MM/DD/YY")
   End If
   If strData <> "" Then
      lngX = Px(0) + LeftMG
      MyPrint strData, lngX, lngY
   End If
   
   '帳單編號
   strData = ""
   If Not IsNull(adoacc1k0.Fields("DocNo").Value) Then
      If Len(adoacc1k0.Fields("DocNo").Value) = 10 Then
         strData = Mid(adoacc1k0.Fields("DocNo").Value, 3, 6)
      Else
         strData = adoacc1k0.Fields("DocNo").Value
      End If
   End If
   
   If bolIsBatchInvoice Then strData = strBatchInvoiceNo 'Added by Morgan 2016/2/15 整批列印請款單
   
   'Modified by Lydia 2019/11/04 "," =>", "
   strDBno = strDBno & strData & ", " 'Added by Lydia 2015/10/19
   
   If strData <> "" Then
      lngX = Px(1) + LeftMG
      MyPrint strData, lngX, lngY
   End If
   
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.FontSize = 10
   End If
   If bol2Jpg Then
      Picture1.FontSize = 10 * douExtRate
   End If
   
   '我方文號
   strData = ""
   If Not IsNull(adoacc1k0.Fields("CaseNo1").Value) Then
      If adoacc1k0.Fields("CaseNo3").Value = "0" And adoacc1k0.Fields("CaseNo4").Value = "00" Then
         strData = adoacc1k0.Fields("CaseNo1").Value & "-" & adoacc1k0.Fields("CaseNo2").Value
      Else
         strData = adoacc1k0.Fields("CaseNo1").Value & "-" & adoacc1k0.Fields("CaseNo2").Value & "-" & adoacc1k0.Fields("CaseNo3").Value
      End If
      'Add by Amy 2024/06/18 商標查名請款單由TS或S 案轉入者,在O/REF.NO.之本所案號後再加入"(查名本所案號)" ex:FCT-051898(S-008316)
      If InStr("" & adoacc1k0.Fields("CaseNo1"), "T") > 0 And "" & adoacc1k0.Fields("CaseNo1") <> "TS" Then
         strData = strData & GetTMQCaseNo(False, "" & adoacc1k0.Fields("DocNo"), bolTQM)
      End If
   End If
   If strData <> "" Then
      If bolIsBatchInvoice Then strData = strData & ",..." 'Added by Morgan 2016/2/15 整批列印請款單
      lngX = Px(2) + LeftMG
      'Modify by Amy 2024/06/14 商標查名請款單由TS或S 案轉入者,在O/REF.NO.之本所案號後再加入"(查名本所案號)" ex:FCT-051898(S-008316)
      If bolTQM = True Then
         MyPrint strData, lngX, lngY, , 0
      Else
         MyPrint strData, lngX, lngY
      End If
   End If

   '貴方文號
'   strData = ""
   'Modify by Amy 2016/03/31 +巨京沒彼所案號抓分所案號
'   adoquery.CursorLocation = adUseClient
'   adoquery.Open "select pa77 as Yno from patent where pa01 = '" & adoacc1k0.Fields("CaseNo1").Value & "' and pa02 = '" & adoacc1k0.Fields("CaseNo2").Value & "' and pa03 = '" & adoacc1k0.Fields("CaseNo3").Value & "' and pa04 = '" & adoacc1k0.Fields("CaseNo4").Value & "' union " & _
'                 "select tm45 as Yno from trademark where tm01 = '" & adoacc1k0.Fields("CaseNo1").Value & "' and tm02 = '" & adoacc1k0.Fields("CaseNo2").Value & "' and tm03 = '" & adoacc1k0.Fields("CaseNo3").Value & "' and tm04 = '" & adoacc1k0.Fields("CaseNo4").Value & "' union " & _
'                 "select lc23 as Yno from lawcase where lc01 = '" & adoacc1k0.Fields("CaseNo1").Value & "' and lc02 = '" & adoacc1k0.Fields("CaseNo2").Value & "' and lc03 = '" & adoacc1k0.Fields("CaseNo3").Value & "' and lc04 = '" & adoacc1k0.Fields("CaseNo4").Value & "' union " & _
'                 "select sp27 as Yno from servicepractice where sp01 = '" & adoacc1k0.Fields("CaseNo1").Value & "' and sp02 = '" & adoacc1k0.Fields("CaseNo2").Value & "' and sp03 = '" & adoacc1k0.Fields("CaseNo3").Value & "' and sp04 = '" & adoacc1k0.Fields("CaseNo4").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoquery.RecordCount <> 0 Then
'      If IsNull(adoquery.Fields("Yno").Value) = False Then
'         strData = adoquery.Fields("Yno").Value
'      End If
'   End If
'   adoquery.Close
   'GetYourRefNo1有改要確認ExcelSaveNew  2016/03/31 +巨京沒彼所案號抓分所案號 是否也要改(未寫抓function因有抓其他資料)
   strData = GetYourRefNo1(adoacc1k0.Fields("CaseNo1"), adoacc1k0.Fields("CaseNo2"), adoacc1k0.Fields("CaseNo3"), adoacc1k0.Fields("CaseNo4"), IIf(Left(adoacc1k0.Fields("FagentNo"), 6) = "Y52269", True, False))
   'Add by Amy 2024/06/14 商標查名請款單由TS或S 案轉入者,Y/REF.NO.有值不秀,避免看不出案號 ex:T-245686(TS-001979) -秀玲
   If bolTQM = True Then strData = ""
   
   'Added by Morgan 2014/2/17 改比照請款單抓法
   If adoacc1k0.Fields("CaseNo1") = "FCP" Then
      strExc(0) = "Select PA106 From CaseProgress,Patent Where PA01(+)=CP01 And PA02(+)=CP02 And PA03(+)=CP03 And PA04(+)=CP04 And CP60='" & adoacc1k0.Fields("DocNo").Value & "' And CP10='605' And CP01='FCP' and pa76 is not null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'Add By Sindy 2016/7/14 +檢查商標延展
   ElseIf InStr("T,FCT,CFT,TF", adoacc1k0.Fields("CaseNo1")) > 0 Then
      strExc(0) = "Select TM65 From CaseProgress,Trademark Where TM01(+)=CP01 And TM02(+)=CP02 And TM03(+)=CP03 And TM04(+)=CP04 And CP60='" & adoacc1k0.Fields("DocNo").Value & "' And CP10='102' And CP01 in('T','FCT','CFT','TF') and TM33 is not null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   Else
      intI = 0
   '2016/7/14 END
   End If
   '年費彼所案號
   If intI = 1 Then
      If PUB_GetFCCaseNo(adoacc1k0.Fields("DocNo").Value, strExc(2), True) = True Then
         strData = strExc(2)
      Else
        strData = "" & RsTemp(0)
      End If
   ElseIf PUB_GetFCCaseNo(adoacc1k0("DocNo"), strExc(2)) = True Then
      strData = strExc(2)
   End If
   'end 2014/2/17
   
   If strData <> "" Then
      If bolIsBatchInvoice Then strData = strData & ",..." 'Added by Morgan 2016/2/15 整批列印請款單
      lngX = Px(3) + LeftMG
      'Modify By Sindy 2013/1/7
      'MyPrint strData, lngX, lngY
      MyPrint Left(strData, 15), lngX, lngY
      '2013/1/7 End
   End If
   
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.FontSize = 12
   End If
   If bol2Jpg Then
      Picture1.FontSize = 12 * douExtRate
   End If
      
   '應收金額
   adoquery.CursorLocation = adUseClient
   If Not (IsNull(adoacc1k0.Fields("OAmount").Value) = False And adoacc1k0.Fields("OAmount").Value <> 0) Then
      '2009/6/4 MODIFY BY SONIA 加A1K31折讓外幣金額
      'adoquery.Open "select sum((a1k08 - nvl(a1k06, 0))),sum(a1k11),sum(nvl(a1k06, 0)) from acc1k0 where substr(a1k01, 1, 8) = '" & Mid(adoacc1k0.Fields("DocNo").Value, 1, 8) & "' and A1K02 < 920201 and a1k29 is null and a1k12 is null and a1k25 is null union " & _
      '              "select sum((a1k08 - nvl(a1k06, 0))),sum(a1k11),sum(nvl(a1k06, 0)) from acc1k0 where a1k01 = '" & adoacc1k0.Fields("DocNo").Value & "' and A1K02 > 920201 and a1k29 is null", adoTaie, adOpenStatic, adLockReadOnly
      '2013/10/3 mdofi by sonia (a1k08 - nvl(a1k06, 0)) as FAmount改為(a1k08 - nvl(a1k31, 0)) as FAmount
      'Added by Morgan 2016/2/15 整批列印請款單
      If bolIsBatchInvoice = True Then
         'Modified by Morgan 2020/12/14 已銷帳也要排除 Ex:X10805261-X10911362
         'Modified by Lydia 2022/03/04 只抓整批請款單尚未結清的金額做小計
         'adoquery.Open "select sum((a1k08 - nvl(a1k31, 0))),sum(a1k11),sum(nvl(a1k06, 0)),sum(nvl(a1k31, 0)) from acc1t0, acc1k0 where a1t02= '" & adoacc1k0.Fields("DocNo").Value & "' and a1k01(+)=a1t01 and a1k29 is null and a1k25 is null", adoTaie, adOpenStatic, adLockReadOnly
         adoquery.Open "select sum((a1k08 - nvl(a1k31, 0))),sum(a1k11),sum(nvl(a1k06, 0)),sum(nvl(a1k31, 0)) from acc1t0, acc1k0 where a1t02 in (select b.a1t01 from acc1t0 a, acc1t0 b where a.a1t01='" & adoacc1k0.Fields("DocNo").Value & "' and a.a1t02=b.a1t02(+)) and a1k01(+)=a1t01 and a1k29 is null and a1k25 is null", adoTaie, adOpenStatic, adLockReadOnly
      Else
      'end 2016/2/15
         adoquery.Open "select sum((a1k08 - nvl(a1k31, 0))),sum(a1k11),sum(nvl(a1k06, 0)),sum(nvl(a1k31, 0)) from acc1k0 where substr(a1k01, 1, 8) = '" & Mid(adoacc1k0.Fields("DocNo").Value, 1, 8) & "' and A1K02 < 920201 and a1k29 is null and a1k12 is null and a1k25 is null union " & _
                    "select sum((a1k08 - nvl(a1k31, 0))),sum(a1k11),sum(nvl(a1k06, 0)),sum(nvl(a1k31, 0)) from acc1k0 where a1k01 = '" & adoacc1k0.Fields("DocNo").Value & "' and A1K02 > 920201 and a1k29 is null", adoTaie, adOpenStatic, adLockReadOnly
      End If 'Added by Morgan 2016/2/15
      '2009/6/4 END
      If adoquery.RecordCount <> 0 Then
         If IsNull(adoquery.Fields(0).Value) Then
            douDollar = 0
         Else
            Select Case strCurr
               'Modify By Sindy 2013/1/18
               Case "NTD"
                  douDollar = Val(adoquery.Fields(1).Value) - Val(adoquery.Fields(2).Value) * Val(adoacc1k0.Fields("Rate").Value)
               '2009/6/4 MODIFY BY SONIA
               'Case Else
               '   douDollar = Val(adoquery.Fields(0).Value)
               '2013/7/29 MODIFY BY SONIA 再改回
'               Case "USD"
'                  douDollar = Val(adoquery.Fields(0).Value)
'               Case Else
'                  Select Case adoacc1k0.Fields("Curr")
'                     Case "USD"
'                        douDollar = Val(adoquery.Fields(0).Value)
'                     Case Else
'                        m_DNRate = PUB_GetUSXRate_1("" & adoacc1k0.Fields("a1k02").Value, adoacc1k0.Fields("Curr"))
'                        douDollar = (Val(adoquery.Fields(1).Value) - Val(adoquery.Fields(3).Value) * m_DNRate) / m_DNRate
'                  End Select
               Case Else
                  douDollar = Val(adoquery.Fields(0).Value)
               '2009/6/4 END
            End Select
         End If
      Else
         douDollar = 0
      End If
      adoquery.Close
   End If
   
   strData = ""
   If douDollar <> 0 Then
      strAmount = Format(douDollar, FDollar)
      intLength = Printer.TextWidth(strAmount)
      strData = strAmount
      douAmount = douAmount + douDollar
      douTAmount = douTAmount + douDollar
   End If
   If strData <> "" Then
      lngX = Px(5) - LeftMG
      MyPrint strData, lngX, lngY, True
   End If
   
   '已收金額
   strData = ""
   If IsNull(adoacc1k0.Fields("OAmount").Value) = False And adoacc1k0.Fields("OAmount").Value <> 0 Then
      Select Case strCurr
         'Modify By Sindy 2013/1/18
         Case "NTD"
            douODollar = Format(adoacc1k0.Fields("OAmount").Value * adoacc1k0.Fields("Rate").Value, FDollar)
         '2009/6/4 MODIFY BY SONIA
         '2013/10/29 modify by sonia 還原 X10211363EUR請款
         Case Else
            douODollar = Format(adoacc1k0.Fields("OAmount").Value, FDollar)
         'Case "USD"
         '   douODollar = Format(adoacc1k0.Fields("OAmount").Value, FDollar)
         'Case Else
         '   Select Case adoacc1k0.Fields("Curr")
         '      Case "USD"
         '         douODollar = Format(adoacc1k0.Fields("OAmount").Value, FDollar)
         '      Case Else
         '         m_DNRate = PUB_GetUSXRate_1("" & adoacc1k0.Fields("a1k02").Value, adoacc1k0.Fields("Curr"))
         '         douODollar = (Format(adoacc1k0.Fields("OAmount").Value * adoacc1k0.Fields("Rate").Value, FDollar)) / m_DNRate
         '   End Select
         '2009/6/4 END
      End Select
      strAmount = Format(douODollar, FDollar)
      intLength = Printer.TextWidth(strAmount)
      strData = strAmount
      douAmount = douAmount - Val(douODollar)
      douRAmount = douRAmount + Val(douODollar)
      douDollar = 0
   Else
      douODollar = 0
   End If
   'Modify By Sindy 2013/1/7 Mark
'   If strData <> "" Then
'      lngX = Px(6) - LeftMG
'      MyPrint strData, lngX, lngY, True
'   End If
   
'2009/6/17 CANCEL BY SONIA 中文無此
'   intDays = CalculateDays(CADate(adoacc1k0.Fields("a1k02").Value), ServerDate)
'   '天期
'   strData = ""
'   If intDays > 90 Then
'      strData = "+90"
'   ElseIf intDays > 60 Then
'      strData = "+60"
'   ElseIf intDays > 30 Then
'      strData = "+30"
'   End If
'
'   If strData <> "" Then
'      lngX = Px(7) - LeftMG
'      MyPrint strData, lngX, lngY, True
'   End If
'2009/6/17 END
   
   'Add By Sindy 2013/1/7
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.FontSize = 10
   End If
   If bol2Jpg Then
      Picture1.FontSize = 10 * douExtRate
   End If
   '案件名稱
   'Modify By Sindy 2015/7/1 +讀取TM131
   'strData = GetPrjName(adoacc1k0.Fields("CaseNo1").Value & "-" & adoacc1k0.Fields("CaseNo2").Value & "-" & adoacc1k0.Fields("CaseNo3").Value & "-" & adoacc1k0.Fields("CaseNo4").Value)
   strData = GetPrjName(adoacc1k0.Fields("CaseNo1").Value & "-" & adoacc1k0.Fields("CaseNo2").Value & "-" & adoacc1k0.Fields("CaseNo3").Value & "-" & adoacc1k0.Fields("CaseNo4").Value, True)
   '2015/7/1 END
   If strData <> "" Then
      lngX = Px(5) - LeftMG + 100
      MyPrint Left(strData, 8), lngX, lngY
   End If
   '申請人
   strData = GetPrjPeopleNum1(adoacc1k0.Fields("CaseNo1").Value & "-" & adoacc1k0.Fields("CaseNo2").Value & "-" & adoacc1k0.Fields("CaseNo3").Value & "-" & adoacc1k0.Fields("CaseNo4").Value)
   If strData <> "" Then
      strData = GetPrjPeople1(strData)
      lngX = Px(6) - LeftMG + 100
      MyPrint Left(strData, 6), lngX, lngY
   End If
   '還原字體大小
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.FontSize = 12
   End If
   If bol2Jpg Then
      Picture1.FontSize = 12 * douExtRate
   End If
   '2013/1/7 End
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
   CloseIme
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Modified by Lydia 2016/09/08 改可選擇JPG 或PDF
   'If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
   'Modified by Lydia 2024/12/31 只存PDF檔
   'If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Add by Morgan 2009/4/7 改A4紙張並可存電子檔
Private Sub PrintDataA4()
   Dim strDocNo As String
   Dim bolReport As Boolean
   Dim strKind As String, strKey1 As String, StrKey2 As String, strKey3 As String
   Dim ff As Integer, strFileList(3) As String, strFileName As String
   Dim strSubject As String, strDate As String
   Dim strContent As String
   Dim dblUsAmount As Double, dblFeeAmount As Double, dblPt As Double
   Dim dblTotUsAmount As Double, dblTotFeeAmount As Double, dblTotPt As Double
   Dim bolSum As Boolean '是否印依代理人加總欄位
   Dim bolNote As Boolean  '是否產生請款單
   Dim strSource As String, strDestination As String
   Dim arrFile
   Dim strMsg As String 'Add by Morgan 2009/7/7
   Dim iPicNo As Integer 'Added by Morgan 2020/3/30
   
   intLength = 0
   douAmount = 0
   douTAmount = 0
   douRAmount = 0
   douOverDue1 = 0
   douOverDue2 = 0
   douOverDue3 = 0
   strNo = ""
   m_DNCurr = "" '2009/6/4 add by sonia
   'Add by Morgan 2009/7/7
   m_iDocCount = 0
   m_iPrintCount = 0
   m_iMailCount = 0
   m_FNo = ""
   'end 2009/7/7
   
   '非單筆時另外印催款明細(大陸或有指定收件者時也不印)
   '2009/6/10 modify by sonia
   'If Text3 <> "020" And (Text1 <> Text2 Or Text1 & Text2 = "") And txtReceiver = "" Then
   If Text3 <> "020" And (Mid(Text1, 1, 6) <> Mid(Text2, 1, 6) Or Text1 & Text2 = "") And txtReceiver = "" Then
   '2009/6/10 end
      bolReport = True
      adoTaie.Execute "Delete From ACCRPT207 Where R20701='" & strUserNum & "'"
   Else
      bolReport = False
   End If
   'Added by Lydia 2024/10/24 代理人Y55822來信，對帳單須以全英文顯示；帳單將改建在Y55822020
   If Text3 = "020" Then
      If Left("" & adoacc1k0.Fields("FAGENTNO"), 6) = "Y55822" Then
         bolChina = False
      Else
         bolChina = True
      End If
   End If
   'end 2024/10/24
   'douExtRate = Screen.TwipsPerPixelX / 15 'Remove by Morgan 2011/10/5
   
   Printer.PaperSize = 9
   Printer.ScaleMode = 1
   iXfix = 400 - 1 * (Printer.Width - Printer.ScaleWidth) / 2
   'Modified by Morgan 2012/3/20 新信紙頁尾有字紙本要上移兩行
   'iYfix = 400 - 1 * (Printer.Height - Printer.ScaleHeight) / 2
   iYfix = -400
   If bolChina Then
      iRowH = 280
   Else
      iRowH = 200
   End If
   SetPx
   SetPy
   '2009/6/22 modify by sonia
   'iPageRows = 20
   If bolChina = True Then
      iPageRows = 16
   Else
      iPageRows = 20
   End If
   '2009/6/22 end
   
   '刪除舊的暫存圖檔
   strExc(1) = App.path & "\$*.jpg"
   If Dir(strExc(1)) <> "" Then Kill strExc(1)
   'Added by Lydia 2019/06/10
   strExc(1) = App.path & "\$*.pdf"
   If Dir(strExc(1)) <> "" Then Kill strExc(1)
   
   Erase strMailFailList
      
   ReDim strMailFailList(0)
   strPicLetter = ""
   strPicFileNames = ""
   iPageNo = 0
   
   strBatchNoList = "" 'Added by Morgan 2016/2/16
   strBatchNoRecList = "" 'Added by Morgan 2016/2/16
   With adoacc1k0
   .MoveFirst
   Do While Not .EOF
      'Added by Lydia 2024/10/24 代理人Y55822來信，對帳單須以全英文顯示；帳單將改建在Y55822020
      If Text3 = "020" Then
         If Left("" & .Fields("FAGENTNO"), 6) = "Y55822" Then
            bolChina = False
         Else
            bolChina = True
         End If
      End If
      'end 2024/10/24
      '若有下國籍條件
      If Text3 <> MsgText(601) Then
         If .Fields("NATION").Value < Text3 Then
            GoTo NextSkip
         End If
      End If
      If Text4 <> MsgText(601) Then
         If .Fields("NATION").Value > Text4 & "z" Then
            GoTo NextSkip
         End If
      End If
      '若是舊系統資料
      If Len(.Fields("DocNo").Value) = 10 And strDocNo = Mid(.Fields("DocNo").Value, 1, 8) And .Fields("OAmount").Value = 0 Then
         GoTo NextSkip
      ElseIf Len(.Fields("DocNo").Value) = 10 Then
         strDocNo = Mid(.Fields("DocNo").Value, 1, 8)
      Else
         strDocNo = .Fields("DocNo").Value
      End If
      
      'Added by Morgan 2016/2/15 整批列印請款單
      bolIsBatchInvoice = False
      If Not IsNull(.Fields("a1t01")) Then
         If .Fields("OAmount").Value > 0 Then
            If InStr(strBatchNoRecList, .Fields("DocDate") & .Fields("a1t02")) = 0 Then
               bolIsBatchInvoice = True
               strBatchNoRecList = strBatchNoRecList & "," & .Fields("DocDate") & .Fields("a1t02")
            End If
         Else
            If InStr(strBatchNoList, .Fields("a1t02")) = 0 Then
               bolIsBatchInvoice = True
               strBatchNoList = strBatchNoList & "," & .Fields("a1t02")
            End If
         End If
         
         If bolIsBatchInvoice Then
            strExc(0) = "select min(a1t02)||'-'||substr(max(a1t01),-3) from acc1t0 where a1t02='" & .Fields("a1t02") & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strBatchInvoiceNo = "" & RsTemp(0)
            End If
         Else
            GoTo NextSkip
         End If
      End If
      'end 2016/2/15
      
      '代理人不同
      '2009/6/3 MODIFY BY SONIA 請款幣別不同也要跳頁
      'If strNo <> .Fields("FagentNo").Value Then
      'Modified by Lydia 2017/02/18 T收款寄證是以個案命名
      'If strNo <> .Fields("FagentNo").Value & .Fields("Curr").Value Then
      'Modify by Amy 2017/08/15 紙本有下客戶編號改檔名-婉莘
      'Memo by Lydia 2019/06/10 從收款作業Frmacc2110來的催款作業,如果遇到同一代理人不同幣別會產生2封催款Email,只有附件不同 (ex.M10802625, M10802626)
      If bolShowCus = True Then strNowCus = Mid(.Fields("CusNo"), 4)
      If (strCallCase = "" And ((bolShowCus = False And strNo <> .Fields("FagentNo").Value & .Fields("Curr").Value) Or (bolShowCus = True And strNo <> .Fields("FagentNo").Value & "(vs " & strNowCus & ")" & .Fields("Curr").Value))) _
        Or (strCallCase <> "" And bolShowCus = False And strNo <> strCallCase & .Fields("Curr").Value) Then
      
         'If douAmount <> 0 Then 'Removed by Morgan 2017/12/6 不論合計是否為0 都要清空否則會殘留到後面,另Email時加控制有欠款才寄 Ex.106/12/4 Y54511,Y54518
            PrintSumA4
            MyNewPage True
         
            douOverDue1 = 0
            douOverDue2 = 0
            douOverDue3 = 0
            douAmount = 0
            douTAmount = 0
            douRAmount = 0
            iPageNo = 0
            
         'End If 'Removed by Morgan 2017/12/6
         
         '電子信箱
         strEMailBox = "" & adoacc1k0("EBox")
         strEmailCC = "" & adoacc1k0("emailcc") 'Added by Lydia 2024/09/18 財務副本信箱
         '產生電子檔
         'Modified by Lydia 2016/09/08 改可選擇JPG 或PDF
         'If Text6 = "Y" Then
         'Modified by Lydia 2024/12/31 只存PDF檔
         'If Text6 = "1" Or Text6 = "2" Then
         If Text6 = "Y" Then
            bol2Jpg = True
            bol2Printer = False
            'Added by Lydia 2016/09/08 選2存PDF檔
            'Modified by Lydia 2024/12/31 只存PDF檔
            'If Text6 = "2" Then
            If Text6 = "Y" Then
               'Modified by Morgan 2024/9/6 因PDFCreator偶會出現無法預期的錯誤,改用Word將JPG轉PDF
               'bol2Jpg = False
               'bol2Pdf = True
               'iYfix = 0 '移動起始Y軸
               bol2Jpg = True
               bol2Pdf = False
               bolJpg2Pdf = True
               'end 2024/9/6
            Else
               bol2Pdf = False
            End If
            'end 2016/09/08
            
            bolEmail = False
            '國外部承辦組
            Select Case Left(Pub_StrUserSt03, 2)
               Case "F1" 'FCT
                  strSavePath = PUB_GetEFilePath("FCT") & "\Account"
               Case "F2" 'FCP
                  strSavePath = PUB_GetEFilePath("FCP") & "\Account"
               Case "F3"
                  strSavePath = PUB_GetEFilePath("FCL") & "\Account"
               Case Else
                  strSavePath = PUB_Getdesktop
            End Select
            
            'Added by Lydia 2016/12/22 傳入本所案號，指定催款單範圍(T收款寄證1728)
            'Modified by Lydia 2017/02/18 改傳變數
            'If Me.Tag <> "" And strCallCase <> "" Then
            '   strSavePath = Me.Tag
            If m_SavePath <> "" Then
                strSavePath = m_SavePath
            End If
            'end 2016/12/22
            
            strDefDir = strSavePath 'Added by Lydia 2016/09/10
            
            'Added by Lydia 2015/11/05 催款通知不放在桌面
            If bolCallMail = False Then
                'Modified by Lydia 2016/12/22 指定催款單範圍(T收款寄證1728)，不建子目錄
                'strSavePath = strSavePath & "\" & .Fields("FagentNo")
                If strCallCase = "" Then strSavePath = strSavePath & "\" & .Fields("FagentNo")
                
                If Dir(strSavePath, vbDirectory) = "" Then
                   MkDir strSavePath
                End If
            Else
                strSavePath = App.path
            End If
            'end 2015/11/05
         Else
            strSavePath = App.path
            'Modify by Morgan 2010/2/8 +可不發Mail
            'If strEMailBox <> "" And UCase(strEMailBox) <> "NO" Then
            If strEMailBox <> "" And UCase(strEMailBox) <> "NO" And Text9 = "" Then
               'Modified by Lydia 2019/06/10 婉莘反應從2018/9/??修改後,依舊以JPG寄送; 與秀玲討論: 直接跑FC催款Email作業,統一預設為PDF檔並且在畫面加註
               'bol2Jpg = True 'Memo by Lydia 2016/09/08 預設為JPG檔
               'bol2Pdf = False 'Added by Lydia 2016/09/08
               'Modified by Morgan 2024/9/6 因PDFCreator偶會出現無法預期的錯誤,改用Word將JPG轉PDF
               'bol2Jpg = False
               'bol2Pdf = True
               bol2Jpg = True
               bol2Pdf = False
               bolJpg2Pdf = True
               'end 2024/9/6
               'end 2019/06/10
               bol2Printer = False
               bolEmail = True
            Else
               bol2Jpg = False
               bol2Pdf = False 'Added by Lydia 2016/09/08
               bol2Printer = True
               bolEmail = False
            End If
         End If
         
         '國外部承辦或有指定收件者
         If bolPromoter Or txtReceiver <> "" Then
            strReceiver = txtReceiver
            strEmailCC = "" 'Added by Lydia 2024/09/18
            'Modified by Lydia 2016/09/06
            'bol2Jpg = True
            'Modified by Lydia 2019/08/02 輸入收件人發送的email附件預設為PDF檔
            'If Text6 = "2" Then
            If Text6 <> "1" Then '除非指定為JPG檔
               'Modified by Morgan 2024/9/6 因PDFCreator偶會出現無法預期的錯誤,改用Word將JPG轉PDF
               'bol2Jpg = False
               'bol2Pdf = True
               'iYfix = 0 '移動起始Y軸
               bol2Jpg = True
               bol2Pdf = False
               bolJpg2Pdf = True
               'end 2024/9/6
            Else
               bol2Jpg = True
               bol2Pdf = False
            End If
            'end 2016/09/02
            bol2Printer = False
            bolEmail = True
         Else
            strReceiver = strEMailBox
         End If
         
         'Add by Morgan 2009/7/6
         '重印紙本控制
         If Text8 = "Y" Then
            bol2Jpg = False
            bolEmail = False
            bolReport = False
            bol2Pdf = False 'Added by Morgan 2020/12/7
         End If
         'Added by Lydia 2020/12/07 debug: 整批催款單有紙本和Email-PDF，因為沒有切換為預設印表機，造成紙本的資料跑到下一筆的Email
         PUB_RestorePrinter Combo1
         
         '是否存電子檔
         'Modified by Lydia 2016/09/08 +存PDF檔
         If bol2Jpg Or bol2Pdf Then
            If strPicLetter = "" Then
               strPicLetter = App.path & "\$Tmp1.jpg"
               'Added by Morgan 2020/3/31
               If strSrvDate(1) >= 智慧所更名日 Then
                  PUB_GetLetterPicID "2", , iPicNo, , , , , True
               Else
               'end 2020/3/31
                  iPicNo = 6
               End If 'Added by Morgan 2020/3/31
               
               If PUB_ReadDB2File(strPicLetter, iPicNo) = True Then
                  Set Picture1.Picture = LoadPicture(strPicLetter)
                  Picture1.AutoSize = True
                  douExtRate = Picture1.Height / 16836 'Add by Morgan 2011/8/10
                  'Added by Lydia 2016/09/08 PDF檔和JPG檔一樣有信頭 'Memo by Lydia 2020/10/15 因為Picture的高寬比和印表機的高寬比不同,所以再次縮放
                  If bol2Pdf Then
                     'Added by Lydia 2020/10/15 因為不同印表機的高寬不同,所以先預設為PDF印表機
                     PUB_RestorePrinter "PDFCreator"
                     'end 2020/10/15
                     If Printer.ScaleHeight < Picture1.Height Then strExc(2) = Format(Printer.ScaleHeight / Picture1.Height, "0.0000")
                     If Printer.ScaleWidth < Picture1.Width Then strExc(3) = Format(Printer.ScaleWidth / Picture1.Width, "0.0000")
                     If Val(strExc(2)) > 0 And Val(strExc(3)) > 0 Then
                        If Val(strExc(2)) <= Val(strExc(3)) Then
                           douExtRate = Val(strExc(2))
                        Else
                           douExtRate = Val(strExc(3))
                        End If
                     Else
                        If Val(strExc(2)) > 0 Then
                           douExtRate = Val(strExc(2))
                        ElseIf Val(strExc(3)) > 0 Then
                           douExtRate = Val(strExc(3))
                        End If
                     End If
                     douExtRate = Printer.ScaleHeight / Picture1.Height '信頭調整為A4滿版
                     If douExtRate > 0 Then
                        Picture1.Height = Picture1.Height * douExtRate
                        Picture1.Width = Picture1.Width * douExtRate
                     End If
            'Added by Lydia 2019/06/12 只在第一次下載信頭
                  End If
               End If
            End If
            'end 2019/06/12
               
               'Added by Morgan 2024/9/6
               If strCallCase <> "" Then
                  strNowDoc = strCallCase & .Fields("Curr").Value & ".PDF"
                  
               ElseIf bolShowCus = True Then
                  strNowDoc = .Fields("FagentNo").Value & "(vs " & strOldCus & ")" & .Fields("Curr").Value & ".PDF"
                  
               Else
                  strNowDoc = .Fields("FagentNo").Value & .Fields("Curr").Value & ".PDF"
                  
               End If
               If bol2Pdf Then
               'end 2024/9/6
               
                     frmPDF.Show
                     'Modified by Lydia 2016/12/22 指定催款單範圍(T收款寄證1728),以案號為檔名
                     'frmPDF.StartProcess strSavePath, .Fields("FagentNo").Value & .Fields("Curr").Value & ".PDF"
                     If strCallCase <> "" Then
                         frmPDF.StartProcess strSavePath, strCallCase & .Fields("Curr").Value & ".PDF"
                         strNowDoc = strCallCase & .Fields("Curr").Value & ".PDF"  'Added by Lydia 2020/02/15 記錄現在列印的.pdf
                     'Add by Amy 2017/08/15 有下客戶編號改檔名
                     ElseIf bolShowCus = True Then
                        strOldCus = Mid("" & .Fields("CusNo"), 4)
                        If strOldCus = MsgText(601) Then strOldCus = "其他"
                        frmPDF.StartProcess strSavePath, .Fields("FagentNo").Value & "(vs " & strOldCus & ")" & .Fields("Curr").Value & ".PDF"
                        strNowDoc = .Fields("FagentNo").Value & "(vs " & strOldCus & ")" & .Fields("Curr").Value & ".PDF" 'Added by Lydia 2020/02/15 記錄現在列印的.pdf
                     Else
                        frmPDF.StartProcess strSavePath, .Fields("FagentNo").Value & .Fields("Curr").Value & ".PDF"
                        strNowDoc = .Fields("FagentNo").Value & .Fields("Curr").Value & ".PDF" 'Added by Lydia 2020/02/15 記錄現在列印的.pdf
                     End If
                     'end 2016/12/22
                  
                     Printer.PaintPicture Picture1, 0, 0, Picture1.Width, Picture1.Height
               
               End If 'Added by Morgan 2024/9/6
                  
            'Remove by Lydia 2019/06/12 只在第一次下載信頭
            '      End If
                  'end 2016/09/08
            '   End If
            'End If
            'end 2019/06/12
         End If
         
         strCurr = adoacc1k0.Fields("Curr").Value 'Modify By Sindy 2015/6/12
         PrintHeadA4
         '2009/6/4 MODIFY BY SONIA 請款幣別不同也要跳頁
         'strNo = .Fields("FagentNo").Value
         'Modified by Lydia 2016/12/22 指定催款單範圍(T收款寄證1728),以案號為檔名
         'strNo = .Fields("FagentNo").Value & .Fields("Curr").Value
         If strCallCase <> "" Then
             strNo = strCallCase & .Fields("Curr").Value
         'Add by Amy 2017/08/15 有下客戶編號改檔名
         ElseIf bolShowCus = True Then
            strOldCus = Mid("" & .Fields("CusNo"), 4)
            If strOldCus = MsgText(601) Then strOldCus = "其他"
            strNo = .Fields("FagentNo").Value & "(vs " & strOldCus & ")" & .Fields("Curr").Value
         Else
             strNo = .Fields("FagentNo").Value & .Fields("Curr").Value
         End If
         'end 2016/12/22
         
         m_DNCurr = adoacc1k0.Fields("Curr").Value
         '2009/6/4 end
         m_FNo = .Fields("FagentNo").Value 'Add by Morgan 2009/7/7
         intCounter = 0
      Else
         intCounter = intCounter + 1
         If intCounter >= iPageRows Then
            MyNewPage
            PrintHeadA4
            intCounter = 0
         End If
      End If
      PrintRowA4
      '產生請款單的電子檔
      If Text7 = "Y" Then
         '舊系統的由人工掃描到原來請款單的存放路徑
         If Len(.Fields("DocNo").Value) = 10 Then
            strFileName = "DN" & Mid(adoacc1k0.Fields("DocNo").Value, 3, 6) & ".pdf"
            strSource = PUB_GetEFilePath(.Fields("CaseNo1")) & "\" & .Fields("CaseNo1") & .Fields("CaseNo2") & IIf(.Fields("CaseNo3") & .Fields("CaseNo4") <> "000", .Fields("CaseNo3") & .Fields("CaseNo4"), "") & "\" & strFileName
            strDestination = strSavePath & "\" & strFileName
            If Dir(strSource) <> "" Then
               FileCopy strSource, strDestination
            End If
         Else
            'Modified by Lydia 2016/09/08 改到催款單列印完後
            'Load Frmacc2480
            'With Frmacc2480
            '   .Text1.Text = strDocNo
            '   .Text2.Text = strDocNo
            '   .txtOutMode = "2"
            '   .m_bBeCalled = True
            '   .m_bEMail = True
            '   .m_SavePath = strSavePath
            '   .Command2_Click
            'End With
            'Unload Frmacc2480
            'strFormName = Me.Name
            'tool3_enabled
            strA1k01List = strA1k01List & strDocNo & ","
            'Added by Lydia 2017/03/02 請款單存放的資料夾路徑
            strA1k01Dir = strA1k01Dir & strDefDir & "\" & .Fields("FagentNo").Value & ","
         End If
      End If
      
NextSkip:
      
      If bolReport = True Then
         '請款日期小於請款日期止日-1年的才要印
         If Val(.Fields("DocDate")) < Val(Replace(Me.MaskEdBox2.Text, "/", "")) - 10000 Then
         
            If InStr("'FCP','FG'", "'" & .Fields("CaseNo1") & "'") > 0 Then
               strExc(0) = "FCP"
            ElseIf InStr("'FCT','CFT','CFC','S','L'", "'" & .Fields("CaseNo1") & "'") > 0 Then
               strExc(0) = "FCT"
            ElseIf InStr("'FCL','LIN'", "'" & .Fields("CaseNo1") & "'") > 0 Then
               strExc(0) = "FCL"
            ElseIf InStr("'P','PS','CFP','CPS'", "'" & .Fields("CaseNo1") & "'") > 0 Then
               strExc(0) = "P"
            ElseIf Left("" & .Fields("CaseNo1"), 1) = "T" Then
               strExc(0) = "T"
            Else
               strExc(0) = ""
            End If
            If .Fields("FAmount").Value > 0 Then
               adoTaie.Execute "insert into ACCRPT207(R20701,R20702,R20703,R20704,R20705,R20706,R20707)" & _
                  " values ('" & strUserNum & "','" & .Fields("DocNo").Value & "','" & .Fields("CaseNo1") & "','" & .Fields("CaseNo2") & "','" & .Fields("CaseNo3") & "','" & .Fields("CaseNo4") & "','" & strExc(0) & "')"
            End If
         End If
      End If
      .MoveNext
   Loop
   End With
   
   PrintSumA4
   
   MyNewPage , True
   'Modified by Lydia 2016/09/08 改可選擇JPG 或PDF
   'If Text6 = "Y" Then
   'Modified by Lydia 2024/12/31 只存PDF檔
   'If Text6 = "1" Or Text6 = "2" Then
   If Text6 = "Y" Then
      'Modified by Lydia 2016/12/22 指定催款單範圍(T收款寄證1728)
      'If Left(Pub_StrUserSt03, 1) = "F" Then
      If strCallCase <> "" Then
            strMsg = strMsg & "電子檔已存於" & strSavePath & " ！" & vbCrLf
            m_iDocCount = 0 '不詢問是否列印
      ElseIf Left(Pub_StrUserSt03, 1) = "F" Then
      'end 2016/12/22
         If strSavePath <> "" Then
            strMsg = strMsg & "電子檔已存於" & strSavePath & " ！" & vbCrLf
         Else
            strMsg = strMsg & "電子檔已存於相關路徑！" & vbCrLf
         End If
      Else
         strMsg = strMsg & "電子檔已存桌面！" & vbCrLf
      End If
   End If
   
   strExc(1) = App.path & "\$*.jpg"
   If Dir(strExc(1)) <> "" Then Kill strExc(1)
   
   'Add by Morgan 2009/7/7
   If m_iDocCount > 0 Then
      InsertQueryLog (m_iDocCount) 'Add By Sindy 2010/12/22
      strExc(0) = ""
      strExc(0) = strExc(0) & "本次共產生 " & m_iDocCount & " 份催款資料" & vbCrLf
      strExc(0) = strExc(0) & "紙本 " & m_iPrintCount & " 份" & vbCrLf
      strExc(0) = strExc(0) & "Email " & m_iMailCount & " 份" & vbCrLf
      
      If strMailFailList(0) <> "" Then
         strExc(0) = strExc(0) & "Email失敗 " & UBound(strMailFailList) + 1 & " 份，清單如下：" & vbCrLf & vbCrLf
         For intI = 0 To UBound(strMailFailList)
            strExc(0) = strExc(0) & strMailFailList(intI) & vbCrLf
         Next
      End If
      'Added by Lydia 2020/09/10 請款單：判斷PDF檔案是否存在
      If m_strErr2480 <> "" Then
         strExc(0) = strExc(0) & vbCrLf & "請款單電子檔產生失敗：" & vbCrLf & Replace(m_strErr2480, "＆", vbCrLf) & vbCrLf
      End If
      'end 2020/09/10
      
      'If MsgBox(strExc(0) & vbCrLf & "是否要列印？" & vbCrLf, vbYesNo + vbDefaultButton1) = vbYes Then 'Remove by Lydia 2020/09/10
         strExc(1) = "催款日期：" & strSrvDate(1) & vbCrLf & vbCrLf
         strExc(1) = strExc(1) & "請款日期：" & MaskEdBox1 & " ～ " & MaskEdBox2 & vbCrLf
         If Text3 & Text4 <> "" Then
            strExc(1) = strExc(1) & "國籍：" & Text3 & " ～ " & Text4 & vbCrLf
         End If
         strExc(0) = strExc(1) & vbCrLf & strExc(0)
      'Modified by Lydia 2020/09/10 改成Email失敗或檔案未能產生要彈訊息，不論成功或有部份失敗都要發email
      '   Printer.Print strExc(0)
      '   Printer.DrawWidth = 1 'Added by Morgan 2020/4/9
      '   Printer.EndDoc
      'End If
            
      If strMailFailList(0) <> "" Or m_strErr2480 <> "" Or (Left(Text1, 6) <> Left(Text2, 6)) Or (Left(Text10, 6) <> Left(Text11, 6)) Then 'Added by Lydia 2020/10/08 單筆(代理人前6碼)的催款單不用寄MAIL給操作人員, 除非有錯誤 !
           PUB_SendMail strUserNum, strUserNum, "", Me.Caption & "-" & "完成作業" & IIf(InStr(strExc(0), "失敗") > 0, "，有發生失敗請參考內文！", ""), strExc(0)
      End If 'Added by Lydia 2020/10/05
      If strMailFailList(0) <> "" Or m_strErr2480 <> "" Then
          MsgBox strExc(0), vbInformation, "發生失敗"
      End If
      'end 2020/09/10
   End If
   
   If bolReport = True Then
      '更新智權人員
      'FCP,P--以請款對象國籍抓國家檔的FCP承辦智權人員
      adoTaie.Execute "update accrpt207 set R20708=(select max(na51) from acc1k0,fagent,customer,nation" & _
         " where a1k01=R20702 and fa01(+)=substr(a1k28,1,8) and fa02(+)=substr(a1k28,9)" & _
         " and cu01(+)=substr(a1k28,1,8) and cu02(+)=substr(a1k28,9) and na01=decode(fa01,null,cu10,fa10))" & _
         " where R20701='" & strUserNum & "' and R20707 in ('FCP','P')"
      'FCT,T--抓該案號最後接洽單之智權人員(若已離職則改放虛建智權人員)
      adoTaie.Execute "update accrpt207 set R20708=(select substr(max(cp09||cp13),10) from caseprogress" & _
         " where cp01=R20703 and cp02=R20704 and cp03=R20705 and cp04=R20706 and cp09<'B')" & _
         " where R20701='" & strUserNum & "' and R20707 in ('FCT','T')"
         
      adoTaie.Execute "update accrpt207 set R20708='F4103'" & _
         " where R20701='" & strUserNum & "' and R20707 in ('FCT','T')" & _
         " and not exists(select * from staff where st01=R20708 and st04='1')"
      
      'FCL--抓請款單之智權人員(最後收文智權人員)
      adoTaie.Execute "update accrpt207 set R20708=(select substr(max(cp05||cp13),9) from caseprogress" & _
         " where cp60=R20702)" & _
         " where R20701='" & strUserNum & "' and R20707='FCL'"
      
      'Modify by Amy 2021/12/09 +R20701='" & strUserNum & "'",因會抓到其他人的資料
      strExc(0) = "select R20707 類別" & _
         ",substrb(ST02,1,6) 智權人員" & _
         ",substrb(na03,1,8) 國籍" & _
         ",substrb(a1k28||' '||rtrim(fa05||' '||fa63||' '||fa64||' '||fa65)||rtrim(cu05||' '||cu88||' '||cu89||' '||cu90),1,30) 請款對象" & _
         ",a1k13||'-'||a1k14||decode(a1k15||a1k16,'000','','-'||a1k15||'-'||a1k16) 本所案號" & _
         ",a1k01 請款單號" & _
         ",substrb(sqldatew(a1k02+19110000),1,10) 請款日期" & _
         ",decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k31, 0)),decode(nvl(a1k10,0),0,0,(a1k11-a1k30)/a1k10)) 應收外幣" & _
         ",decode(nvl(a1k30,0),0,nvl(a1k09,0),sign(nvl(a1k09,0)-a1k30),1,nvl(a1k09,0)-a1k30,0) 應收規費" & _
         ",(nvl(a1k11,0)-nvl(a1k09,0))/1000 點數" & _
         " From (select a.*,b.*,c.*,d.*,e.*,decode(fa01,null,cu10,fa10) X1 from accrpt207 a, staff b, acc1k0 c, fagent d,customer e" & _
         " where st01(+)=R20708 and a1k01(+)=R20702 And R20701='" & strUserNum & "' " & _
         " and fa01(+)=substr(a1k28,1,8) and fa02(+)=substr(a1k28,9)" & _
         " and cu01(+)=substr(a1k28,1,8) and cu02(+)=substr(a1k28,9)" & _
         ") X, Nation where na01(+)=X1" & _
         " order by 1,2,3,4,5,7,6"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         .MoveFirst
         strKind = ""
         strSubject = Format(strSrvDate(1), "####/##/##") & "催款明細"
         Erase strFileList
         strDate = IIf(Me.MaskEdBox1.Text = "___/__/__", "", Format(ChangeTStringToWDateString(Replace(Me.MaskEdBox1.Text, "/", "")), "YYYY/MM/DD")) & _
               IIf(Me.MaskEdBox1.Text = "___/__/__" And Me.MaskEdBox2.Text = "___/__/__", "", " ～ ") & _
               IIf(Me.MaskEdBox2.Text = "___/__/__", "", Format(ChangeTStringToWDateString(Replace(Me.MaskEdBox2.Text, "/", "")), "YYYY/MM/DD"))
         Do While Not .EOF
            If .Fields(0) <> strKind Then
               If strKind <> "" And ff > 0 Then
                  '合計
                  strContent = Space(90) & convForm(Format(dblTotUsAmount, "#,##0"), 12, , True) '代理人應收總額
                  strContent = strContent & " " & convForm(Format(dblTotFeeAmount, "#,##0"), 10, , True) '應收總規費
                  strContent = strContent & " " & convForm(Format(dblTotPt, "#,##0.000"), 8, , True)  '總點數
                  
                  Print #ff, "------ -------- -------------------------- --------------- ---------- ---------- -------- ------------ ---------- --------"
                  Print #ff, strContent
               End If
               
               If ff > 0 Then Close #ff
               ff = FreeFile
               
               strKind = "" & .Fields(0)
               strKey1 = ""
               StrKey2 = ""
               strKey3 = ""
               dblTotUsAmount = 0
               dblTotFeeAmount = 0
               dblTotPt = 0
               
               strFileName = App.path & "\催款明細(" & strKind & ").txt"
               
               Open strFileName For Output As ff
                              
               strExc(1) = Format(DBDATE(MaskEdBox2.Text) - 10000, "####/##/##")
               Print #ff, ""
               Print #ff, "※本報表只列出請款日期早於 " & strExc(1) & " 的資料。"
               Print #ff, ""
               Print #ff, "                                           " & strSubject & "(" & strKind & ")"
               Print #ff, ""
               Print #ff, "                                           請款日期：" & strDate
               Print #ff, "                                                                                          代理人(美金)     (台幣)   (台幣)"
               Print #ff, "智權人 國籍     請款對象                   本所案號        請款單號   請款日期   應收美金 應收總額     應收總規費   總點數"
               Print #ff, "------ -------- -------------------------- --------------- ---------- ---------- -------- ------------ ---------- --------"
               
               
               strFileList(0) = strFileList(0) & strFileName & ";"
               Select Case strKind
                  Case "FCP", "P"
                     strFileList(1) = strFileList(1) & strFileName & ";"
                  Case "FCT", "T"
                     strFileList(2) = strFileList(2) & strFileName & ";"
                 Case "FCL"
                     strFileList(3) = strFileList(3) & strFileName & ";"
               End Select
                  
            End If
            strContent = Empty
            '智權人員
            If strKey1 = "" & .Fields(1) Then
               strContent = strContent & Space(6)
            Else
               strContent = strContent & convForm("" & .Fields(1), 6) '智權人員
            End If
            '國籍
            If strKey1 & StrKey2 = "" & .Fields(1) & .Fields(2) Then
               strContent = strContent & Space(9)
            Else
               strContent = strContent & " " & convForm("" & .Fields(2), 8) '國籍
            End If
            '請款對象
            If strKey3 = "" & .Fields(3) Then
               strContent = strContent & Space(27)
            Else
               strContent = strContent & " " & convForm("" & .Fields(3), 26) '請款對象
            End If
            strContent = strContent & " " & convForm("" & .Fields(4), 15) '本所案號
            strContent = strContent & " " & convForm("" & .Fields(5), 10) '請款單號
            strContent = strContent & " " & convForm("" & .Fields(6), 10) '請款日期
            strContent = strContent & " " & convForm(Format("" & .Fields(7), "#,##0"), 8, , True) '應收美金
            dblUsAmount = dblUsAmount + Val("" & .Fields(7))
            dblTotUsAmount = dblTotUsAmount + Val("" & .Fields(7))
            dblFeeAmount = dblFeeAmount + Val("" & .Fields(8))
            dblTotFeeAmount = dblTotFeeAmount + Val("" & .Fields(8))
            dblPt = dblPt + Val("" & .Fields(9))
            dblTotPt = dblTotPt + Val("" & .Fields(9))
            strKey1 = "" & .Fields(1)
            StrKey2 = "" & .Fields(2)
            strKey3 = "" & .Fields(3)
            .MoveNext
            bolSum = False
            If .EOF Then
               bolSum = True
            '代理人編號不同
            ElseIf strKind & strKey1 & StrKey2 & strKey3 <> "" & .Fields(0) & .Fields(1) & .Fields(2) & .Fields(3) Then
               bolSum = True
            End If
            If bolSum = True Then
               strContent = strContent & " " & convForm(Format(dblUsAmount, "#,##0"), 12, , True) '代理人應收總額
               strContent = strContent & " " & convForm(Format(dblFeeAmount, "#,##0"), 10, , True) '應收總規費
               strContent = strContent & " " & convForm(Format(dblPt, "#,##0.000"), 8, , True)  '總點數
               dblUsAmount = 0
               dblFeeAmount = 0
               dblPt = 0
            End If
            Print #ff, strContent
         Loop
         '合計
         strContent = Space(90) & convForm(Format(dblTotUsAmount, "#,##0"), 12, , True) '代理人應收總額
         strContent = strContent & " " & convForm(Format(dblTotFeeAmount, "#,##0"), 10, , True) '應收總規費
         strContent = strContent & " " & convForm(Format(dblTotPt, "#,##0.000"), 8, , True)  '總點數
         
         Print #ff, "------ -------- -------------------------- --------------- ---------- ---------- -------- ------------ ---------- --------"
         Print #ff, strContent
         If ff > 0 Then Close #ff
         'FCP,P
         If strFileList(1) <> "" Then
            'Modify by Morgan 2011/2/25 85030 留職停薪半年,暫改 88003
            'Modify by Amy 2022/11/24 +if 88003(王文安協理)退休暫改77015(顏裕洋副理)99037(簡偉倫經理)
            If strSrvDate(1) >= 20221130 Then
                SendReport "77015;99037", strSubject, strFileList(1)
            Else
                SendReport "88003", strSubject, strFileList(1)
            End If
         End If
         'FCT,T
         If strFileList(2) <> "" Then
            'modify by sonia 2021/10/20 68005離職改78011葉易雲,80030洪琬姿
            SendReport "78011;80030", strSubject, strFileList(2)
         End If
         'FCL
         If strFileList(3) <> "" Then
            'Modified by Morgan 2021/6/9 73029 離職改 99015
            SendReport "99015", strSubject, strFileList(3)
         End If
         'ALL
         If strFileList(0) <> "" Then
            'modify by sonia 2015/6/30 改68009為81040
            SendReport "81040", strSubject, strFileList(0)
            arrFile = Split(strFileList(0), ";")
            For intI = LBound(arrFile) To UBound(arrFile)
               If arrFile(intI) <> "" Then
                  Kill arrFile(intI)
               End If
            Next
         End If
         End With
      End If
   End If
   
   strMsg = strMsg & "作業結束！"
   'Modified by Lydia 2015/10/19
   'MsgBox strMsg, vbInformation
   If bolCallMail = False Then MsgBox strMsg, vbInformation
End Sub

Private Sub MyNewPage(Optional bolEndPage As Boolean, Optional bolEndDoc As Boolean)
'Added by Lydia 2016/09/08
Dim tmpArr As Variant
Dim inX As Integer
Dim strTempPath As String
Dim strPdfA1k01 As String 'Added by Lydia 2017/02/18 記錄-請款單pdf檔路徑
Dim tmpArr2 As Variant 'Added by Lydia 2017/03/02
Dim tmpErr As String 'Added by Lydia 2020/09/10

   'Add by Morgan 2009/7/7
   If bolPromoter = False Then
      If (bolEndPage Or bolEndDoc) Then
         If bolEndDoc = False Then m_iDocCount = m_iDocCount + 1 'Modify by Amy 2018/06/08 +if  因結束多加1造成m_iDocCount錯
         If bol2Printer Then
            m_iPrintCount = m_iPrintCount + 1
         'Removed by Morgan 2017/12/6 移到下面
         'ElseIf bolEmail Then
         '   m_iMailCount = m_iMailCount + 1
         End If
      End If
   End If
   'end 2009/7/7
   
   '頁碼
   PrintPageNo
   
   If bolEndDoc Then
      Printer.DrawWidth = 1 'Added by Morgan 2020/4/9
      Printer.EndDoc
      'Added by Lydia 2016/09/08 輸出PDF檔
      If bol2Pdf Then
          frmPDF.EndtProcess
          'Added by Lydia 2020/02/15 判斷檔案是否存在, 超過時間就繼續
          'Modified by Lydia 2020/09/10 超過時間，改不發email直接出發信失敗清單
          'If PUB_ChkFileStatus(strSavePath & "\" & strNowDoc) = False Then
          tmpErr = ""
          If PUB_ChkFileStatus(strSavePath & "\" & strNowDoc, False, tmpErr) = False Then
            If tmpErr <> "" Then
                If strMailFailList(0) <> "" Then
                   ReDim Preserve strMailFailList(UBound(strMailFailList) + 1)
                End If
                strMailFailList(UBound(strMailFailList)) = strNo & " : " & Mid(tmpErr, 2) & IIf(bolEmail = True, "請重新Email", "")
            End If
          'end 2020/09/10
          End If
          'end 2020/02/15
          Unload frmPDF
      End If

      'Added by Lydia 2016/09/08 改到催款單列印完後
      If strA1k01List <> "" Then
         tmpArr = Split(strA1k01List, ",")
         tmpArr2 = Split(strA1k01Dir, ",") 'Added by Lydia 2017/03/02
         For inX = 0 To UBound(tmpArr)
            If Trim(tmpArr(inX)) <> "" Then
                Load Frmacc2480
                With Frmacc2480
                   .Text1.Text = Trim(tmpArr(inX))
                   .Text2.Text = Trim(tmpArr(inX))
                   .txtOutMode = "2"
                   .m_bBeCalled = True
                   .m_CallPrevForm = Me.Name 'Added by Lydia 2020/01/06 呼叫請款單的程式名稱
                   .m_bEMail = True
                   'Modified by Lydia 2017/03/02 一般催款單的資料夾用代理人區分
                   '.m_SavePath = strSavePath
                   If bolCallMail = False And strCallCase = "" Then  '排除催款通知,T收款寄證1728
                       .m_SavePath = IIf(Trim(tmpArr2(inX)) <> "", Trim(tmpArr2(inX)), strSavePath)
                   Else
                       .m_SavePath = strSavePath
                   End If
                   'end 2017/03/02
                   .Command2_Click
                End With
                'Added by Lydia 2020/09/10 請款單：判斷PDF檔案是否存在
                If Frmacc2480.m_strOutErr <> "" Then
                     m_strErr2480 = m_strErr2480 & Frmacc2480.m_strOutErr
                Else
                'end 2020/09/10
                     strPdfA1k01 = strPdfA1k01 & "*" & Frmacc2480.Tag  'Added by Lydia 2017/02/18 記錄-請款單pdf檔路徑
                End If 'Added by Lydia 2020/09/10
                Unload Frmacc2480
                strFormName = Me.Name
                tool3_enabled
            End If
         Next
      End If
      'end 2016/09/08
      strA1k01List = "" 'Added by Lydia 2017/02/18 跑完後清空
      strA1k01Dir = "" 'Added by Lydia 2017/03/02
      
   ElseIf bol2Printer Then
      Printer.NewPage

   'Added by Lydia 2016/09/08 存PDF檔
   ElseIf bol2Pdf Then
      If bolEndPage Then '列印最後=>存檔
          Printer.DrawWidth = 1 'Added by Morgan 2020/4/9
          Printer.EndDoc
          frmPDF.EndtProcess
          'Added by Lydia 2020/02/15 判斷檔案是否存在, 超過時間就繼續
          'Modified by Lydia 2020/09/10 超過時間，改不發email直接出發信失敗清單
          'If PUB_ChkFileStatus(strSavePath & "\" & strNowDoc) = False Then
          tmpErr = ""
          If PUB_ChkFileStatus(strSavePath & "\" & strNowDoc, False, tmpErr) = False Then
            If tmpErr <> "" Then
                If strMailFailList(0) <> "" Then
                   ReDim Preserve strMailFailList(UBound(strMailFailList) + 1)
                End If
                strMailFailList(UBound(strMailFailList)) = strNo & " : " & Mid(tmpErr, 2) & IIf(bolEmail = True, "請重新Email", "")
            End If
          'end 2020/09/10
          End If
          'end 2020/02/15
          Unload frmPDF
          'Modified by Lydia 2019/06/12 因為現在email作業預設PDF檔(參考「Morgan 2017/12/6 不論合計是否為0 都要清空否則會殘留到後面,」)
                                  '所以不同代理人都會執行本段,造成重複frmPDF.StartProcess無法正常產生附件
          'If bolCallMail = False Then
          If bolCallMail = False And strDefDir <> "" Then
              strTempPath = strDefDir & "\" & adoacc1k0.Fields("FagentNo") 'Memo by Lydia 2017/02/18 可能有不同代理人的資料
              If Dir(strTempPath, vbDirectory) = "" Then
                 MkDir strTempPath
                 strSavePath = strDefDir & "\" & adoacc1k0.Fields("FagentNo")
              End If
          'Remove by Lydia 2019/06/12
          'Else
          '    strTempPath = App.path
          End If
          'Remove by Lydia 2019/06/12 回到PrintDataA4執行
          'frmPDF.Show
          ''Add by Amy 2017/08/15 有下客戶編號改檔名
          'If bolShowCus = True Then
          '  strOldCus = Mid("" & adoacc1k0.Fields("CusNo"), 4)
          '  If strOldCus = MsgText(601) Then strOldCus = "其他"
          '  frmPDF.StartProcess strTempPath, adoacc1k0.Fields("FagentNo").Value & "(vs " & strOldCus & ")" & adoacc1k0.Fields("Curr").Value & ".PDF"
          'Else
          '  frmPDF.StartProcess strTempPath, adoacc1k0.Fields("FagentNo").Value & adoacc1k0.Fields("Curr").Value & ".PDF"
          'End If
          'end 2019/06/12
      Else   '列印下一頁
          Printer.NewPage
          'Added by Lydia 2020/04/23 新頁面要印信頭/尾
          If strPicLetter <> "" And douExtRate > 0 Then
               Printer.PaintPicture Picture1, 0, 0, Picture1.Width, Picture1.Height
          End If
          'end 2020/04/23
      End If
      'Printer.PaintPicture Picture1, 0, 0, Picture1.Width, Picture1.Height 'Remove by Lydia 2019/06/12 回到PrintDataA4執行
   End If
   
   'Modified by Lydia 2016/09/08 +存PDF檔 bol2Pdf
   If bol2Jpg Or bol2Pdf Then
      If bol2Jpg Then 'Added by Lydia 2016/09/08 JPG檔一頁一個檔案
        If (bolEndPage Or bolEndDoc) And iPageNo = 1 Then
           strExc(1) = strSavePath & "\" & strNo & ".jpg"
        Else
           strExc(1) = strSavePath & "\" & strNo & IIf(iPageNo > 0, "_" & Format(iPageNo, "00"), "") & ".jpg"
        End If
        PUB_SavePic Picture1, strExc(1)
        Set Picture1.Picture = LoadPicture(strPicLetter)
        Picture1.AutoSize = True
        douExtRate = Picture1.Height / 16836 'Add by Morgan 2011/8/10
        
        strPicFileNames = strPicFileNames & strExc(1) & "*"
      Else 'PDF合併為一個檔案
          If InStr(strPicFileNames, strNo & ".pdf") = 0 Then
             strPicFileNames = strPicFileNames & "*" & strSavePath & "\" & strNo & ".pdf"
          End If
      End If
      
      '最後一頁
      If bolEndPage Or bolEndDoc Then
         
         'Added by Morgan 2024/9/6
         If bol2Jpg And bolJpg2Pdf Then
            If PUB_JPG2PDF(strPicFileNames, strSavePath & "\" & strNowDoc) = True Then
               strPicFileNames = strSavePath & "\" & strNowDoc
            End If
         End If
         'end 2024/9/6
      
         'Modified by Morgan 2017/12/6
         'If bolEmail Then
         If bolEmail And douAmount <> 0 Then
            If tmpErr = "" Then 'Added by Lydia 2020/09/10 增加判斷
               m_iMailCount = m_iMailCount + 1
         'end 2017/12/6
               If bolPromoter Then
                  ShowOutLookMail
               Else
                  bolMailFailNoAlert = True
                  bolMailSendOk = False
                  'Memo by Lydia 2016/09/09 若是以PDF檔寄mail,SendMail會將郵件格式改成非html (JPG檔會以html寄出)
                  '所內同仁
                  If InStr(strReceiver, "@") = 0 Then
                     'Modified by Morgan 2014/3/4 主旨加代理人編號
                     'PUB_SendMail strUserNum, strReceiver, "", "Statement of Accounts", GetMailContent, , strPicFileNames, True
                     'Modified by Lydia 2024/09/18 +財務副本信箱+ , , , strEmailCC
                     PUB_SendMail strUserNum, strReceiver, "", "Statement of Accounts (" & m_FNo & ")", GetMailContent, , strPicFileNames, True, , , strEmailCC
                  Else
                     'Modify by Morgan 2011/4/22 改以ipdept@taie.com.tw 寄但回覆還是給寄件人(70004)
                     'PUB_SendMail strUserNum, strReceiver, "", "Statement of Accounts", GetMailContent, , strPicFileNames, True, True, True
                     'Modified by Morgan 2011/10/12 改用 account@taie.com.tw 寄
                     'PUB_SendMail strUserNum, strReceiver, "", "Statement of Accounts", GetMailContent, , strPicFileNames, True, True, True, , "ipdept@taie.com.tw", "TAI E INTERNATIONAL PATENT & LAW OFFICE", strUserNum
                     'Modified by Morgan 2014/3/4 主旨加代理人編號
                     'PUB_SendMail strUserNum, strReceiver, "", "Statement of Accounts", GetMailContent, , strPicFileNames, True, True, True, , strAccMailBox, "TAI E INTERNATIONAL PATENT & LAW OFFICE", strUserNum
                     'Modified by Morgan 2014/8/27 改回覆到財務信箱 -- 婧瑄
                     'PUB_SendMail strUserNum, strReceiver, "", "Statement of Accounts (" & m_FNo & ")", GetMailContent, , strPicFileNames, True, True, True, , strAccMailBox, "TAI E INTERNATIONAL PATENT & LAW OFFICE", strUserNum
                     'Modified by Lydia 2024/09/18 +財務副本信箱strEmailCC
                     PUB_SendMail strUserNum, strReceiver, "", "Statement of Accounts (" & m_FNo & ")", GetMailContent, , strPicFileNames, True, True, True, strEmailCC, strAccMailBox, "TAI E INTERNATIONAL PATENT & LAW OFFICE", strAccMailBox
                  End If
                  bolMailFailNoAlert = False
                  If bolMailSendOk = False Then
                     If strMailFailList(0) <> "" Then
                        ReDim Preserve strMailFailList(UBound(strMailFailList) + 1)
                     End If
                     strMailFailList(UBound(strMailFailList)) = m_FNo & " : " & strEMailBox
                  End If
               End If
            End If 'Added by Lydia 2020/09/10
         End If
         
         'Added by Lydia 2016/12/22　指定催款單範圍(T收款寄證1728),回傳附件路徑
         If strCallCase <> "" Then
            'Modified by Lydia 2017/02/18 +請款單PDF
            'Me.Tag = strPicFileNames
            Me.Tag = strPicFileNames & strPdfA1k01
         Else
            Me.Tag = ""
         End If
         'end 2016/12/22
         
         strPicFileNames = ""
      End If
   End If
End Sub
'Add by Morgan 2009/4/20
'呼叫 OutLook 撰寫郵件視窗
Private Sub ShowOutLookMail()
Dim objOutLook As Object
Dim objMail As Object
Dim strTemplatePath As String
Dim arrAttachment
Dim ii As Integer
Dim strContent  As String 'Added by Lydia 2015/10/19
Dim strNation As String 'Added by Lydia 2019/08/12

On Error GoTo ErrHnd

'Added by Lydia 2015/10/19 + bolCallmail
If bolCallMail = False Then
   '郵件範本檔
   Select Case Left(Pub_StrUserSt03, 2)
      Case "F1"
         strTemplatePath = App.path & "\FCT.oft"
      Case "F2"
         strTemplatePath = App.path & "\FCP.oft"
      Case "F3"
         strTemplatePath = App.path & "\FCL.oft"
      Case Else
         strTemplatePath = App.path & "\FCP.oft"
   End Select
   
   Set objOutLook = CreateObject("Outlook.Application")
   If Dir(strTemplatePath) <> "" Then
      Set objMail = objOutLook.CreateItemFromTemplate(strTemplatePath)
   Else
      Set objMail = objOutLook.CreateItem(0)
   End If
   
   objMail.Subject = "Statement of Accounts"
   objMail.To = strEMailBox
   'Added by Lydia 2024/09/18 財務副本信箱
   If strEMailBox <> "" And strEmailCC <> "" Then
      objMail.cc = strEmailCC
   End If
   'end 2024/09/18
   
   objMail.Body = GetMailContent & vbCrLf & objMail.Body & vbCrLf & vbCrLf
   If strPicFileNames <> "" Then
      arrAttachment = Split(strPicFileNames, "*")
      For ii = LBound(arrAttachment) To UBound(arrAttachment)
         If arrAttachment(ii) <> "" Then
            objMail.Attachments.add (arrAttachment(ii))
         End If
      Next
   End If
   '較早帳款未付即時催款
Else
      'Added by Lydia 2019/08/12 下載信尾的範本
      strNation = GetPrjNationNumber(Text1)
      If strNation <= "010" Or InStr("020,013,044", Left(strNation, 3)) > 0 Then '臺灣,大陸地區,香港,澳門
          '中文
            If Dir(App.path & "\$$TOT-000M31-0-03.oft") = "" Then
                Call PUB_GetSampleFile("$$TOT-000M31-0-03.oft", "TOT-000M31-0-03")
            End If
            strTemplatePath = App.path & "\$$TOT-000M31-0-03.oft"
      Else
          '英文
            If Dir(App.path & "\$$TOT-000M31-0-02.oft") = "" Then
                Call PUB_GetSampleFile("$$TOT-000M31-0-02.oft", "TOT-000M31-0-02")
            End If
            strTemplatePath = App.path & "\$$TOT-000M31-0-02.oft"
      End If
      '呼叫新郵件：
      Set objOutLook = CreateObject("Outlook.Application")
      'Added by Lydia 2019/08/12
      If strTemplatePath <> "" Then
          Set objMail = objOutLook.CreateItemFromTemplate(strTemplatePath)
      Else
      'end 2019/08/12
          Set objMail = objOutLook.CreateItem(0)
      End If
      objMail.Subject = "Payment Issue " & ChangeCustomerS(Text1.Text)
      strContent = GetMailContent2
      '轉HTML格式
      strContent = Replace(strContent, "新細明體", "Times New Roman")
      strContent = Replace(strContent, vbCrLf, "<BR>")
      strContent = Replace(strContent, "  ", "&nbsp;&nbsp;")
      'Modified by Lydia 2018/09/03 預設字體
      'objMail.HTMLBody = "<FONT FACE=""Times New Roman"">" & strContent & "<BR>" & Replace(objMail.HTMLBody, "&lt;LetterMemo&gt;", IIf(ExceptFieldData("公用備註/英") <> "", "<BR>Message:<BR>" & Replace(ChgHTMLFormat(ExceptFieldData("公用備註/英")), vbCrLf, "<BR>") & "<BR><BR>", "&nbsp;")) & "</FONT>"
      'Added by Lydia 2019/08/12 中文
      If strNation <= "010" Or InStr("020,013,044", Left(strNation, 3)) > 0 Then '臺灣,大陸地區,香港,澳門
           objMail.HTMLBody = "<body style=font-size:12pt;font-family:細明體;serif&quot;>" & strContent & "</body><BR>" & Replace(objMail.HTMLBody, "&lt;LetterMemo&gt;", IIf(ExceptFieldData("公用備註/英") <> "", "<BR>Message:<BR>" & Replace(ChgHTMLFormat(ExceptFieldData("公用備註/英")), vbCrLf, "<BR>") & "<BR><BR>", "&nbsp;"))
      Else
      'end 2019/08/12
           objMail.HTMLBody = "<body style=font-size:11pt;font-family:&quot;Times&nbsp;&nbsp;New&nbsp;&nbsp;Roman&quot;,&quot;serif&quot;>" & strContent & "</body><BR>" & Replace(objMail.HTMLBody, "&lt;LetterMemo&gt;", IIf(ExceptFieldData("公用備註/英") <> "", "<BR>Message:<BR>" & Replace(ChgHTMLFormat(ExceptFieldData("公用備註/英")), vbCrLf, "<BR>") & "<BR><BR>", "&nbsp;"))
      End If
      If strPicFileNames <> "" Then
         arrAttachment = Split(strPicFileNames, "*")
         For ii = LBound(arrAttachment) To UBound(arrAttachment)
            If arrAttachment(ii) <> "" Then
               objMail.Attachments.add (arrAttachment(ii))
            End If
         Next
      End If
      
End If
'end 2015/10/19
   objMail.Display

ErrHnd:
   If Err.Number <> 0 Then
      MsgBox "開啟撰寫郵件視窗失敗，請人工作業！"
   End If
   
   Set objMail = Nothing
   Set objOutLook = Nothing
   'Added by Lydia 2015/11/05 刪檔
   If bolCallMail And strPicFileNames <> "" Then
     ' If strPicFileNames <> "" Then
         arrAttachment = Split(strPicFileNames, "*")
         For ii = LBound(arrAttachment) To UBound(arrAttachment)
            If arrAttachment(ii) <> "" Then
               Kill arrAttachment(ii)
            End If
         Next
     ' End If
   End If
   
End Sub

Private Sub PrintPageNo()
   Dim stPageNo As String
   stPageNo = Format(iPageNo, "#")
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      Printer.FontSize = 12
      Printer.FontBold = True
      Printer.CurrentX = iXfix + Px(0) + (Px(7) - Px(0) - Printer.TextWidth(stPageNo)) / 2
      Printer.CurrentY = iYfix + Py(4) + 30
      Printer.Print stPageNo
   End If
   
   If bol2Jpg Then
      Picture1.FontSize = 12 * douExtRate
      Picture1.FontBold = True
      Picture1.CurrentX = (Px(0) + (Px(7) - Px(0) - Picture1.TextWidth(stPageNo)) / 2) * douExtRate
      Picture1.CurrentY = (Py(4) + 30) * douExtRate
      Picture1.Print stPageNo
   End If
End Sub

Private Function GetMailContent() As String
   Dim stDate As String
   Dim StrMailContent As String
   
   If bolChinese = False Then
      If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
         stDate = Val(FCDate(MaskEdBox2.Text)) + 19110000
      Else
         stDate = strSrvDate(1)
      End If
      StrMailContent = ""
      
      '非國外部承辦且有指定收件人時
      If Not bolPromoter And Me.txtReceiver <> "" Then
         StrMailContent = StrMailContent & "To: " & strEMailBox & vbCrLf & vbCrLf
      End If
      
      StrMailContent = StrMailContent & vbTab & Space(50) & ChgEngDate(strSrvDate(1)) & vbCrLf
      StrMailContent = StrMailContent & "Account No: " & m_FNo & vbCrLf
      StrMailContent = StrMailContent & strAttention & vbCrLf
      StrMailContent = StrMailContent & "Re: Statement of Accounts" & vbCrLf & vbCrLf
      'Modified by Morgan 2024/4/10 對外統一用 Dear Colleagues --林總
      'StrMailContent = StrMailContent & "Dear Sirs," & vbCrLf & vbCrLf
      StrMailContent = StrMailContent & "Dear Colleagues," & vbCrLf & vbCrLf
      'end 2024/4/10
      If Not bolPromoter Then
         StrMailContent = StrMailContent & "As identified below, please find the statement of accounts summarizing all unpaid invoices up to" & vbCrLf
      Else
         StrMailContent = StrMailContent & "Please find enclosed the statement of accounts summarizing all unpaid invoices up to" & vbCrLf
      End If
      StrMailContent = StrMailContent & "the period ending " & ChgEngDate(stDate) & "." & vbCrLf & vbCrLf
      StrMailContent = StrMailContent & "Should your records show full or partial payment of the enclosed statement of accounts," & vbCrLf
      StrMailContent = StrMailContent & "please inform us accordingly and, if possible, enclose copies of remittances or other" & vbCrLf
      StrMailContent = StrMailContent & "such proof so we may confirm the same." & vbCrLf
      
      '非國外部承辦
      If Not bolPromoter Then
         StrMailContent = StrMailContent & vbCrLf & "Please be advised that this e-mail address be reserved for account matters only."
         StrMailContent = StrMailContent & vbCrLf & "Please direct all case matter to ipdept@taie.com.tw to ensure the quickest reply." & vbCrLf
         StrMailContent = StrMailContent & vbCrLf & "Thank you for your services and do not hesitate to contact us regarding accounts matters." & vbCrLf
         StrMailContent = StrMailContent & vbCrLf & "TAI E INTERNATIONAL PATENT & LAW OFFICE"
         StrMailContent = StrMailContent & vbCrLf & "Accounting department"
         StrMailContent = StrMailContent & vbCrLf & vbCrLf & vbCrLf & vbCrLf
      End If
      
      GetMailContent = StrMailContent
   End If
End Function
'Add by Morgan 2009/4/7
Private Sub SmartPrint(ByVal p_Data As String, ByRef p_lngX As Long, ByRef p_lngY As Long, Optional p_lngWidth As Long = 80, Optional p_lngLineH As Long = 300)
   Dim strData As String, strCache As String, i As Integer
   
   strData = Trim(p_Data) '去除空白以免多跳行
   For i = 1 To Len(p_Data)
      If Printer.TextWidth(strCache & Mid(strData, i, 1)) > p_lngWidth * 56.7 Then
         'Modified by Lydia 2016/09/08 +存PDF檔
         If bol2Printer Or bol2Pdf Then
            Printer.CurrentX = iXfix + p_lngX
            Printer.CurrentY = iYfix + p_lngY
            'Modified by Lydia 2022/05/19 逐字檢查Unicode文字改以圖片方式列印
            'Printer.Print strCache
            PUB_PrintUnicodeText strCache, Printer.CurrentX, Printer.CurrentY, 0
         End If
         If bol2Jpg Then
            Picture1.CurrentX = p_lngX * douExtRate
            Picture1.CurrentY = p_lngY * douExtRate
            Picture1.Print strCache
         End If
         strCache = Mid(strData, i, 1)
         p_lngY = p_lngY + p_lngLineH
      Else
         strCache = strCache & Mid(strData, i, 1)
      End If
   Next
   If strCache <> "" Then
      'Modified by Lydia 2016/09/08 +存PDF檔
      If bol2Printer Or bol2Pdf Then
         Printer.CurrentX = iXfix + p_lngX
         Printer.CurrentY = iYfix + p_lngY
         'Modified by Lydia 2022/05/19 逐字檢查Unicode文字改以圖片方式列印
         'Printer.Print strCache
         PUB_PrintUnicodeText strCache, Printer.CurrentX, Printer.CurrentY, 0
      End If
      If bol2Jpg Then
         Picture1.CurrentX = p_lngX * douExtRate
         Picture1.CurrentY = p_lngY * douExtRate
         Picture1.Print strCache
      End If
   End If
End Sub
'Add by Morgan 2009/4/7
Private Sub MyPrint(ByVal p_Data As String, ByRef p_lngX As Long, ByRef p_lngY As Long, Optional ByVal bolReverse As Boolean, Optional ByVal iLimitPos As Integer = 0)
   'Modified by Lydia 2016/09/08 +存PDF檔
   If bol2Printer Or bol2Pdf Then
      'Added by Morgan 2014/8/26
      '控制長度不可超過下一欄位
      If iLimitPos > 0 Then
         Do While (Printer.TextWidth(p_Data) > (Px(iLimitPos + 1) - p_lngX - 15))
            p_Data = Left(p_Data, Len(p_Data) - 1)
         Loop
      End If
      'end 2014/8/26
      If bolReverse = True Then
         Printer.CurrentX = iXfix + p_lngX - Printer.TextWidth(p_Data)
      Else
         Printer.CurrentX = iXfix + p_lngX
      End If
      Printer.CurrentY = iYfix + p_lngY
      'Modified by Lydia 2022/05/19 逐字檢查Unicode文字改以圖片方式列印
      'Printer.Print p_Data
      PUB_PrintUnicodeText p_Data, Printer.CurrentX, Printer.CurrentY, 0
   End If
   If bol2Jpg Then
      'Added by Morgan 2014/8/26
      '控制長度不可超過下一欄位
      If iLimitPos > 0 Then
         Do While (Picture1.TextWidth(p_Data) > douExtRate * (Px(iLimitPos + 1) - p_lngX - 15))
            p_Data = Left(p_Data, Len(p_Data) - 1)
         Loop
      End If
      'end 2014/8/26
      
      If bolReverse = True Then
         Picture1.CurrentX = p_lngX * douExtRate - Picture1.TextWidth(p_Data)
      Else
         Picture1.CurrentX = p_lngX * douExtRate
      End If
      Picture1.CurrentY = p_lngY * douExtRate
      Picture1.Print p_Data
   End If
End Sub

Private Sub SendReport(MailTo As String, MailSubject As String, Attachment As String)

   Dim strText As String
   Dim ii As Integer, jj As Integer
   Dim arrAtt
   
   '測試用
   If Pub_StrUserSt03 = "M51" Then
      MailTo = strUserNum
   End If
   
   strText = vbCrLf & vbCrLf & vbCrLf & String(2, "　") & "列印注意事項：" & vbCrLf & _
         vbCrLf & String(4, "　") & "1.利用筆記本開啟附件" & _
         vbCrLf & String(4, "　") & "2.將視窗展開到最大" & _
         vbCrLf & String(4, "　") & "3.取消<自動換行>設定" & _
         vbCrLf & String(4, "　") & "4.<字型>設定為<細明體 標準 11>" & _
         vbCrLf & String(4, "　") & "5.左右邊界分別設<10mm 0mm>" & _
         vbCrLf & String(4, "　") & "6.選擇<橫印>"
            
   'Added by Morgan 2021/6/9 轉HTML格式,改用 SMTP 寄信並 CC 給自己
   strText = Replace(strText & " ", " ", "&nbsp;")
   strText = "<DIV style=""FONT: 12pt 細明體"">" & strText & "</DIV>"
   PUB_SendMail strUserNum, MailTo, "", MailSubject, strText, , Replace(Attachment, ";", "*"), , , , strUserNum
   'end 2021/6/9
   
'Removed by Morgan 2021/6/9 取消 Outlook 寄信,都改用 SMTP
'   If Attachment <> Empty Then
'      arrAtt = Split(Attachment, ";")
'      strText = Space(UBound(arrAtt) + 1) & vbCrLf & strText
'   End If
'
'   DoEvents
'   MAPISession1.LogonUI = False
'   MAPISession1.UserName = strUserNum
'   Err.Clear
'On Error Resume Next
'   MAPISession1.SignOn
'   If Err.Number <> 0 Then
'      MsgBox "EMail發送失敗!!請啟動 OutLook 後重試!!"
'      Screen.MousePointer = vbDefault
'      Exit Sub
'   End If
'   MAPIMessages1.SessionID = MAPISession1.SessionID
'   MAPIMessages1.MsgIndex = -1
'   MAPIMessages1.Compose
'   'Modify By Sindy 2014/1/16
'   'MAPIMessages1.MsgSubject = "◎系統代發◎" & MailSubject
'   MAPIMessages1.MsgSubject = "◎" & IIf(Pub_StrUserSt03 = "M51" And PUB_GetST05(strUserNum) <> "" And UCase(PUB_GetDbTerminal) = "(M51-1)", PUB_GetDbTerminal, "") & MailSubject
'   '2014/1/16 END
'   MAPIMessages1.MsgNoteText = strText
'   If Attachment <> Empty Then
'      jj = 0
'      For ii = 0 To UBound(arrAtt)
'         If arrAtt(ii) <> "" Then
'            MAPIMessages1.AttachmentIndex = jj
'            MAPIMessages1.AttachmentPosition = jj
'            MAPIMessages1.AttachmentPathName = arrAtt(ii)
'            jj = jj + 1
'         End If
'      Next
'   End If
'   MAPIMessages1.RecipIndex = 0
'   MAPIMessages1.RecipDisplayName = MailTo
'   MAPIMessages1.ResolveName
'   MAPIMessages1.Send
'   MAPISession1.SignOff
'end 2021/6/9
End Sub
'限定字串長度
'Remove by Lydia 2018/08/24 與basQuery重複
'Private Function convForm(ByVal p_InStr As String, ByVal p_Num As Integer, Optional ByVal p_Char As String = " ", Optional ByVal p_IsNumber As Boolean) As String
'   If p_IsNumber Then
'      convForm = StrConv(RightB(StrConv(String(p_Num, p_Char) & p_InStr, vbFromUnicode), p_Num), vbUnicode)
'   Else
'      convForm = StrConv(LeftB(StrConv(p_InStr & String(p_Num, p_Char), vbFromUnicode), p_Num), vbUnicode)
'   End If
'End Function

Private Sub Text7_GotFocus()
   TextInverse Text7
   CloseIme
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then '
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
   CloseIme
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then '
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
   CloseIme
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then '
      KeyAscii = 0
      Beep
   End If
End Sub

'Added by Lydia 2015/10/19
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
   If m_PrevForm.Name = "Frmacc2110" Then
      bolCallMail = True
   End If
End Sub

'Added by Lydia 2015/10/19 較早帳款未付即時催款(內文)
Private Function GetMailContent2() As String
   Dim strTxt As String, iR As Integer
   Dim StrMailContent As String
   Dim rs1 As New ADODB.Recordset
   Dim strDizhang As String 'Add by Amy 2022/08/16 帳款處理情形
    
    'Modified by Lydia 2019/08/12 +NA01
    'Modify by Amy +帳款處理情形 fa103/cu142
    strTxt = "SELECT FA79,FA16,FA10 as NA01,FA103 as Dizhang FROM FAGENT WHERE FA01=" & CNULL(Mid(ChangeCustomerL(Text1), 1, 8)) & " AND FA02=" & CNULL(Mid(ChangeCustomerL(Text1), 9, 1))
    'Modified by Lydia 2016/02/03 +X編號(客戶)
    strTxt = strTxt & "Union SELECT CU115,CU20,CU10 as NA01,CU142 as Dizhang FROM CUSTOMER WHERE CU01=" & CNULL(Mid(ChangeCustomerL(Text1), 1, 8)) & " AND CU02=" & CNULL(Mid(ChangeCustomerL(Text1), 9, 1))
    iR = 1
    Set RsTemp = ClsLawReadRstMsg(iR, strTxt)
    If iR = 1 Then
         'Add by Amy 2022/08/16
         strDizhang = "" & RsTemp.Fields("Dizhang")
         If strDizhang <> MsgText(601) And strDizhang <> "D" Then
            StrMailContent = StrMailContent & "<DIV><FONT color=""red""><STRONG><B>不能催款,再確認!</B></STRONG></DIV>" & vbCrLf
         End If
         'Added by Lydia 2016/07/12 如果代理人(Yor X編號)有暫收款, 請在內文上方加註文字
         strTxt = "select * from acc120 where a1203='" & ChangeCustomerL(Text1.Text) & "' AND a1201 not in (select a1303 from acc130) and a1201 not in (select a1p30 from acc1p0 where a1p02 in ('F','K', 'I') and a1p05 = '2401' and a1p07 <> 0 and a1p30 is not null) "
         iR = 1
         Set rs1 = ClsLawReadRstMsg(iR, strTxt)
         If iR = 1 Then
            With rs1
                .MoveFirst
                strTxt = ""
                Do While Not .EOF
                   strTxt = strTxt & "" & .Fields("a1201") & " / " & ChangeTStringToTDateString(.Fields("a1202")) & " / " & .Fields("a1204") & " / " & .Fields("a1207") & " / " & IIf("" & .Fields("a1209") = "2", "收款轉暫收", "" & .Fields("a1211"))
                   strTxt = strTxt & vbCrLf
                   .MoveNext
                Loop
            End With
            StrMailContent = StrMailContent & vbCrLf & "<DIV><FONT color=""red""><STRONG>" & strTxt & "</STRONG></DIV>"
            StrMailContent = StrMailContent & vbCrLf & "<DIV><FONT color=""black"">"
            strTxt = ""
         End If
         'end 2016/07/12

         If "" & RsTemp(0) <> "" Then StrMailContent = StrMailContent & vbCrLf & RsTemp(0) & vbCrLf
         If "" & RsTemp(1) <> "" Then StrMailContent = StrMailContent & vbCrLf & RsTemp(1) & vbCrLf
         
         'Added by Lydia 2019/08/12 中文版
         If "" & RsTemp("NA01") <= "010" Or InStr("020,013,044", Left("" & RsTemp.Fields("NA01"), 3)) > 0 Then  '臺灣,大陸地區,香港,澳門
             StrMailContent = StrMailContent & vbCrLf
             StrMailContent = StrMailContent & "To: 財務部門" & vbCrLf & vbCrLf
             StrMailContent = StrMailContent & "您好," & vbCrLf & vbCrLf
             StrMailContent = StrMailContent & "感謝 貴公司在" & ChangeTStringToTDateString(DBDATE(strLDate)) & "的匯款" & currAmount & ", 本所已依據您的指示沖帳." & vbCrLf & vbCrLf
             'Modified by Lydia 2019/11/04 -1 改為- 2
             StrMailContent = StrMailContent & "經查本所資料, 帳單 " & Mid(strDBno, 1, Len(strDBno) - 2) & " 並不包含在此次付款中. 請協助確認該帳單是否已付款. 若尚未付款, 請盡速安排匯款. 隨信附上對帳單及帳單供您核閱. " & vbCrLf & vbCrLf
             StrMailContent = StrMailContent & "祝好," & vbCrLf & vbCrLf
             StrMailContent = StrMailContent & strUserName & " 財務處"
         Else
         'end 2019/08/12
             StrMailContent = StrMailContent & vbCrLf
             StrMailContent = StrMailContent & "Attn: Accounting Dept." & vbCrLf & vbCrLf
             'Modified by Morgan 2024/4/10 對外統一用 Dear Colleagues --林總
             'StrMailContent = StrMailContent & "Dear Sirs," & vbCrLf & vbCrLf
             StrMailContent = StrMailContent & "Dear Colleagues," & vbCrLf & vbCrLf
             'end 2024/4/10
             'Modified by Lydia 2019/11/12 內文變更
             'StrMailContent = StrMailContent & "Thank you for your payment of " & currAmount & " on " & ChgEngDate((strLDate)) & ", we have credited your account accordingly." & vbCrLf & vbCrLf
             StrMailContent = StrMailContent & "Thank you for your payment of " & currAmount & " on " & ChgEngDate((strLDate)) & ". We have credited your account accordingly." & vbCrLf & vbCrLf
             'Modified by Lydia 2015/11/05 內文變更
            ' StrMailContent = StrMailContent & "After viewing our records, we noticed that there is an older debit note " & Mid(strDBno, 1, Len(strDBno) - 1) & " skipped from this payment. Please check your records and confirm whether this debit note has been paid. Attached herewith is our current statement of account for your kind reference. " & vbCrLf & vbCrLf
             'Modified by Lydia 2018/09/03 內文變更
             'StrMailContent = StrMailContent & "After viewing our records, we noticed that several older debit notes " & Mid(strDBno, 1, Len(strDBno) - 1) & " were skipped from this payment. Please check your records and confirm whether these debit notes have been paid. Attached herewith is our current statement of account for your kind reference. " & vbCrLf & vbCrLf
             'Modified by Lydia 2019/11/04 內文變更
             'StrMailContent = StrMailContent & "After viewing our records, debit notes " & Mid(strDBno, 1, Len(strDBno) - 1) & " were skipped from this payment. Please check your records and confirm whether these debit notes have been paid. Attached herewith is our current statement of account for your kind reference. " & vbCrLf & vbCrLf
             strExc(1) = Mid(strDBno, 1, Len(strDBno) - 2)
             If InStr(strExc(1), ",") > 0 Then strExc(1) = Mid(strExc(1), 1, InStrRev(strExc(1), ",") - 1) & " and" & Mid(strExc(1), InStrRev(strExc(1), ",") + 1)
             StrMailContent = StrMailContent & "After reviewing our records, we noticed that debit notes " & strExc(1) & " (attached for your reference) were omitted from this payment. Please check your records and confirm whether these debit notes have been paid. Our current statement of account is also attached for your kind reference. " & vbCrLf & vbCrLf
             StrMailContent = StrMailContent & "We look forward to your reply." & vbCrLf & vbCrLf
             'end 2019/11/04
             StrMailContent = StrMailContent & "Best regards," & vbCrLf & vbCrLf
             
             '抓英文名稱
             strTxt = "SELECT NVL(ST12,' ') FROM STAFF WHERE ST01='" & strUserNum & "' "
             iR = 1
             Set rs1 = ClsLawReadRstMsg(iR, strTxt)
             StrMailContent = StrMailContent & rs1(0) & Space(60 - (Len(rs1(0)) * 2) + 2) & "Fred C. T. Yen" & vbCrLf
             'Modified by Lydia 2019/10/23 Accounting   Department=> Accounting Dept.
             'StrMailContent = StrMailContent & "Accounting   Department" & Space(60 - (Len("Accounting   Department") * 2) + 8) & "Patent Attorney" & vbCrLf
             StrMailContent = StrMailContent & "Accounting Dept." & Space(60 - (Len("Accounting Dept.") * 2) + 6) & "Patent Attorney" & vbCrLf
             'Modified by Lydia 2019/10/23
             'StrMailContent = StrMailContent & Space(60) & "Managing  Partner" & vbCrLf & vbCrLf & vbCrLf
             StrMailContent = StrMailContent & Space(60) & "Managing  Partner" & vbCrLf
             StrMailContent = StrMailContent & "</DIV>" 'Added by Lydia 2016/07/12
             'Remove by Lydia 2019/08/12 採用範本的信尾
             'StrMailContent = StrMailContent & "<DIV><FONT face=""Times New Roman""><STRONG>" & "Tai E International Patent & Law Office" & "</STRONG>" & vbCrLf
'             StrMailContent = StrMailContent & "9Fl., No. 112, Sec. 2, Chang-An E. Rd." & vbCrLf
'             StrMailContent = StrMailContent & "Taipei 104, Taiwan, R.O.C." & vbCrLf
'             StrMailContent = StrMailContent & "P.O. Box: 46-478, Taipei 104, Taiwan, R.O.C." & vbCrLf
'             StrMailContent = StrMailContent & "Tel: 886-2-25061023, 25081531" & vbCrLf
'             StrMailContent = StrMailContent & "Fax: 886-2-25068147, 25064319, 25076571, 25090804" & vbCrLf
'             StrMailContent = StrMailContent & "URL: https://www.taie.com.tw" & vbCrLf
         End If
    End If
    GetMailContent2 = StrMailContent
End Function

'Add by Amy 2017/02/02 更新客戶編號
Private Sub UpdCusData()
    Dim RsQ As New ADODB.Recordset, rsA As New ADODB.Recordset
    Dim strQ As String, strSql As String, strUpd As String
    Dim intQ As Integer, intA As Integer
    Dim strNo(1 To 4) As String
    Dim strWhere(1 To 5) As String 'Add by Amy 2017/07/31
    
    '無法一次語法更新,多申請人前六碼相同,抓申請人第一個是此客戶編號  ex:Y20804 X21419 多申請人(前六碼相同)
    strQ = "Select Distinct R005 as CaseNo1,R006 as CaseNo2,R007 as CaseNo3,R008 as CaseNo4 From Accrpt2470 " & _
                "Where Id='" & strUserNum & "' Order by R005||R006||R007||R008"
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        With RsQ
            .MoveFirst
            Do While Not .EOF
                strNo(1) = .Fields("CaseNo1")
                strNo(2) = .Fields("CaseNo2")
                strNo(3) = .Fields("CaseNo3")
                strNo(4) = .Fields("CaseNo4")
                'Modify by Amy 2017/07/31 若只下客戶編號迄號為XZZZZZ 則抓此代理人無客戶編號資料 ex:Y51817 1060630 TS-001452
                For i = LBound(strWhere) To UBound(strWhere)
                    strWhere(i) = ""
                Next i
                Select Case strNo(1)
                    Case "CFP", "FCP", "P" '專利
                        If Text10 = MsgText(601) And Left(Text11, 6) = "XZZZZZ" Then
                            strWhere(1) = " And pa26 is not null "
                            strWhere(2) = " And pa27 is not null "
                            strWhere(3) = " And pa28 is not null "
                            strWhere(4) = " And pa29 is not null "
                            strWhere(5) = " And pa30 is not null "
                        Else
                            strWhere(1) = " And pa26>= '" & Text10 & "' And pa26 <='" & Text11 & "' And pa26 is not null "
                            strWhere(2) = " And pa27>= '" & Text10 & "' And pa27 <='" & Text11 & "' And pa27 is not null "
                            strWhere(3) = " And pa28>= '" & Text10 & "' And pa28 <='" & Text11 & "' And pa28 is not null "
                            strWhere(4) = " And pa29>= '" & Text10 & "' And pa29 <='" & Text11 & "' And pa29 is not null "
                            strWhere(5) = " And pa30>= '" & Text10 & "' And pa30 <='" & Text11 & "' And pa30 is not null "
                        End If
                        strSql = "Select CusNo From (" & _
                                    "Select '(1)'||pa26 as CusNo From Patent Where pa01='" & strNo(1) & "' And pa02='" & strNo(2) & "' And pa03='" & strNo(3) & "' and pa04='" & strNo(4) & "' " & strWhere(1) & _
                         "Union Select '(2)'||pa27 as CusNo From Patent Where pa01='" & strNo(1) & "' And pa02='" & strNo(2) & "' And pa03='" & strNo(3) & "' and pa04='" & strNo(4) & "' " & strWhere(2) & _
                         "Union Select '(3)'||pa28 as CusNo From Patent Where pa01='" & strNo(1) & "' And pa02='" & strNo(2) & "' And pa03='" & strNo(3) & "' and pa04='" & strNo(4) & "' " & strWhere(3) & _
                         "Union Select '(4)'||pa29 as CusNo From Patent Where pa01='" & strNo(1) & "' And pa02='" & strNo(2) & "' And pa03='" & strNo(3) & "' and pa04='" & strNo(4) & "' " & strWhere(4) & _
                         "Union Select '(5)'||pa30 as CusNo From Patent Where pa01='" & strNo(1) & "' And pa02='" & strNo(2) & "' And pa03='" & strNo(3) & "' and pa04='" & strNo(4) & "' " & strWhere(5) & _
                                ") Order by CusNo "
                    Case "CFT", "FCT", "T", "TF" '商標
                        If Text10 = MsgText(601) And Left(Text11, 6) = "XZZZZZ" Then
                            strWhere(1) = " And tm23 is not null "
                            strWhere(2) = " And tm78 is not null "
                            strWhere(3) = " And tm79 is not null "
                            strWhere(4) = " And tm80 is not null "
                            strWhere(5) = " And tm81 is not null "
                        Else
                            strWhere(1) = " And tm23>= '" & Text10 & "' And tm23 <='" & Text11 & "' And tm23 is not null "
                            strWhere(2) = " And tm78>= '" & Text10 & "' And tm78 <='" & Text11 & "' And tm78 is not null "
                            strWhere(3) = " And tm79>= '" & Text10 & "' And tm79 <='" & Text11 & "' And tm79 is not null "
                            strWhere(4) = " And tm80>= '" & Text10 & "' And tm80 <='" & Text11 & "' And tm80 is not null "
                            strWhere(5) = " And tm81>= '" & Text10 & "' And tm81 <='" & Text11 & "' And tm81 is not null "
                        End If
                        strSql = "Select CusNo From (" & _
                                    "Select '(1)'||tm23 as CusNo From Trademark Where tm01='" & strNo(1) & "' And tm02='" & strNo(2) & "' And tm03='" & strNo(3) & "' and tm04='" & strNo(4) & "' " & strWhere(1) & _
                         "Union Select '(2)'||tm78 as CusNo From Trademark Where tm01='" & strNo(1) & "' And tm02='" & strNo(2) & "' And tm03='" & strNo(3) & "' and tm04='" & strNo(4) & "' " & strWhere(2) & _
                         "Union Select '(3)'||tm79 as CusNo From Trademark Where tm01='" & strNo(1) & "' And tm02='" & strNo(2) & "' And tm03='" & strNo(3) & "' and tm04='" & strNo(4) & "' " & strWhere(3) & _
                         "Union Select '(4)'||tm80 as CusNo From Trademark Where tm01='" & strNo(1) & "' And tm02='" & strNo(2) & "' And tm03='" & strNo(3) & "' and tm04='" & strNo(4) & "' " & strWhere(4) & _
                         "Union Select '(5)'||tm81 as CusNo From Trademark Where tm01='" & strNo(1) & "' And tm02='" & strNo(2) & "' And tm03='" & strNo(3) & "' and tm04='" & strNo(4) & "' " & strWhere(5) & _
                                ") Order by CusNo "
                    Case "CFL", "FCL", "L", "LIN" '法務
                         If Text10 = MsgText(601) And Left(Text11, 6) = "XZZZZZ" Then
                            strWhere(1) = " And lc11 is not null "
                            strWhere(2) = " And lc43 is not null "
                            strWhere(3) = " And lc44 is not null "
                            strWhere(4) = " And lc45 is not null "
                            strWhere(5) = " And lc46 is not null "
                        Else
                            strWhere(1) = " And lc11>= '" & Text10 & "' And lc11 <='" & Text11 & "' And lc11 is not null "
                            strWhere(2) = " And lc43>= '" & Text10 & "' And lc43 <='" & Text11 & "' And lc43 is not null "
                            strWhere(3) = " And lc44>= '" & Text10 & "' And lc44 <='" & Text11 & "' And lc44 is not null "
                            strWhere(4) = " And lc45>= '" & Text10 & "' And lc45 <='" & Text11 & "' And lc45 is not null "
                            strWhere(5) = " And lc46>= '" & Text10 & "' And lc46 <='" & Text11 & "' And lc46 is not null "
                        End If
                        strSql = "Select CusNo From (" & _
                                    "Select '(1)'||lc11 as CusNo From Lawcase Where lc01='" & strNo(1) & "' And lc02='" & strNo(2) & "' And lc03='" & strNo(3) & "' And lc04='" & strNo(4) & "' " & strWhere(1) & _
                         "Union Select '(2)'||lc43 as CusNo From Lawcase Where lc01='" & strNo(1) & "' And lc02='" & strNo(2) & "' And lc03='" & strNo(3) & "' And lc04='" & strNo(4) & "' " & strWhere(2) & _
                         "Union Select '(3)'||lc44 as CusNo From Lawcase Where lc01='" & strNo(1) & "' And lc02='" & strNo(2) & "' And lc03='" & strNo(3) & "' And lc04='" & strNo(4) & "' " & strWhere(3) & _
                         "Union Select '(4)'||lc45 as CusNo From Lawcase Where lc01='" & strNo(1) & "' And lc02='" & strNo(2) & "' And lc03='" & strNo(3) & "' And lc04='" & strNo(4) & "' " & strWhere(4) & _
                         "Union Select '(5)'||lc46 as CusNo From Lawcase Where lc01='" & strNo(1) & "' And lc02='" & strNo(2) & "' And lc03='" & strNo(3) & "' And lc04='" & strNo(4) & "' " & strWhere(5) & _
                                ") Order by CusNo "
                    Case Else '服務
                        If Text10 = MsgText(601) And Left(Text11, 6) = "XZZZZZ" Then
                            strWhere(1) = " And sp08 is not null "
                            strWhere(2) = " And sp58 is not null "
                            strWhere(3) = " And sp59 is not null "
                            strWhere(4) = " And sp65 is not null "
                            strWhere(5) = " And sp66 is not null "
                        Else
                            strWhere(1) = " And sp08>= '" & Text10 & "' And sp08 <='" & Text11 & "' And sp08 is not null "
                            strWhere(2) = " And sp58>= '" & Text10 & "' And sp58 <='" & Text11 & "' And sp58 is not null "
                            strWhere(3) = " And sp59>= '" & Text10 & "' And sp59 <='" & Text11 & "' And sp59 is not null "
                            strWhere(4) = " And sp65>= '" & Text10 & "' And sp65 <='" & Text11 & "' And sp65 is not null "
                            strWhere(5) = " And sp66>= '" & Text10 & "' And sp66 <='" & Text11 & "' And sp66 is not null "
                        End If
                        strSql = "Select CusNo From (" & _
                                    "Select '(1)'||sp08 as CusNo From Servicepractice Where sp01='" & strNo(1) & "' And sp02='" & strNo(2) & "' And sp03='" & strNo(3) & "' And sp04='" & strNo(4) & "' " & strWhere(1) & _
                         "Union Select '(2)'||sp58 as CusNo From Servicepractice Where sp01='" & strNo(1) & "' And sp02='" & strNo(2) & "' And sp03='" & strNo(3) & "' And sp04='" & strNo(4) & "' " & strWhere(2) & _
                         "Union Select '(3)'||sp59 as CusNo From Servicepractice Where sp01='" & strNo(1) & "' And sp02='" & strNo(2) & "' And sp03='" & strNo(3) & "' And sp04='" & strNo(4) & "' " & strWhere(3) & _
                         "Union Select '(4)'||sp65 as CusNo From Servicepractice Where sp01='" & strNo(1) & "' And sp02='" & strNo(2) & "' And sp03='" & strNo(3) & "' And sp04='" & strNo(4) & "' " & strWhere(4) & _
                         "Union Select '(5)'||sp66 as CusNo From Servicepractice Where sp01='" & strNo(1) & "' And sp02='" & strNo(2) & "' And sp03='" & strNo(3) & "' And sp04='" & strNo(4) & "' " & strWhere(5) & _
                                ") Order by CusNo "
                End Select
                'end 2017/07/31
                
                 intA = 1
                Set rsA = ClsLawReadRstMsg(intA, strSql)
                If intA = 1 Then
                    strUpd = "Update Accrpt2470 Set R040='" & rsA.Fields("CusNo") & "' Where Id='" & strUserNum & "' " & _
                                    "And R005='" & strNo(1) & "' And R006='" & strNo(2) & "' " & _
                                    "And R007='" & strNo(3) & "' And R008='" & strNo(4) & "' "
                    cnnConnection.Execute strUpd
                End If
                rsA.Close
                .MoveNext
            Loop
        End With
    End If
    RsQ.Close
    Set RsQ = Nothing
End Sub

'Add by Amy 2017/08/07 取得代理人名稱
Private Function GetFAgentName(ByVal stCode As String) As String
    Dim RsQ  As New ADODB.Recordset
    Dim strQ As String

    stCode = Left(stCode & "0000000", 9)
    '中->英->日
    strQ = "NVL(FA04,Nvl(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)) as AgName"
    If bolChina = False Then
        '英->日>中
        strQ = "Nvl(FA05||' '||FA63||' '||FA64||' '||FA65,Nvl(FA06,FA04)) as AgName"
    End If
    strQ = "Select " & strQ & " From Fagent Where FA01='" & Left(stCode, 8) & "' and FA02='" & Right(stCode, 1) & "'"
    If RsQ.State <> adStateClosed Then RsQ.Close
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
        GetFAgentName = "" & RsQ.Fields("AgName")
    End If
    RsQ.Close
    Set RsQ = Nothing
End Function

'Add by Amy 2022/06/14 還原整理,將Y2776格式拆開(格式不同,避免改其他時改到)
Private Sub ExcelSaveNew(Optional ByRef pFileName As String)
    Dim xlsAgentPoint As New Excel.Application, wksrpt As New Worksheet, strWkName As String
    Dim xlsFileName As String, strSql As String, strSQL2 As String, strField As String, strFieldN As String
    Dim strOldAgNo As String, strOldAppNo As String, strOldCurr As String, intStartR As String
    Dim intXlsSheet As Integer, intField As Integer
    Dim bolFirst As Boolean
    Dim strAllField As String, strAllWidth As String
On Error GoTo ErrHand
    '中文Excel抬頭(國籍為020)
    If bolChina = True Then
        strAllField = "編號,本事務所名稱,帳單日期<mm/dd/yyy>,帳單編號,本所案號,貴所案號,Application No.(Only for new case),Filing Date(Only for new case),商標/專利 申請號,Filing Date,Category<select>" & _
                            ",客戶編號,客戶名稱,幣別,服務費,規費,雜費,帳單金額"
        strAllWidth = "5.5, 12, 10, 10, 9.5, 13, 13, 13, 13, 13 " & _
                            ", 6, 10, 10, 7, 8, 8, 9.6, 10"
    Else
        'Modify by Amy 2022/07/20 申請人英文顯示錯誤 原:Application
        strAllField = "NO.,Law Firm's Name<select>,Date<mm/dd/yyy>,Invoice No.,Our Ref,Your Ref,Application No.(Only for new case),Filing Date(Only for new case),TradeMark/Patent Application No.,Filing Date,Category<select>" & _
                           ",Applicant No.,Applicant,Currency<select>,Attorney Fee,Official Fee,Disburesment Fee,Total Fee"
        strAllWidth = "5.5, 12, 10, 10, 9.5, 13, 13, 13, 13, 13" & _
                            ", 6, 10, 10, 7, 8, 8, 9.6, 10"
    End If
    '不是Y27766(特殊格式) +USD欄
    If Left(Text1, 6) <> "Y27766" Then
        strAllField = strAllField & ",USD"
        strAllWidth = strAllWidth & ",5"
    End If
    strF = Split(strAllField, ",")
    intWidth = Split(strAllWidth, ",")
    
NextAg:
    intField = 65: intXlsSheet = 1: intCounter = 1: strOldAppNo = "": intNo = 1: bolFirst = True
    
    '有下客戶編號條件顯示客戶編號前六碼-莘
    If strOldAgNo = MsgText(601) Then
        xlsFileName = Text1 & "催款單" & IIf(bolShowCus = True, "(vs " & Left(Text10, 6) & ")", "") & ServerDate & MsgText(43)
    Else
         xlsFileName = strOldAgNo & "催款單" & IIf(bolShowCus = True, "(vs " & Left(Text10, 6) & ")", "") & ServerDate & MsgText(43)
    End If
    If Dir(strExcelPath & xlsFileName) = MsgText(601) Then
       If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
       End If
    Else
         Kill strExcelPath & xlsFileName
    End If
    pFileName = strExcelPath & xlsFileName
    
    xlsAgentPoint.SheetsInNewWorkbook = 3 '預設工作表數量
    xlsAgentPoint.Workbooks.add
    xlsAgentPoint.Application.WindowState = xlMinimized
    
NextCus:
    If intXlsSheet > 3 Then
        xlsAgentPoint.Worksheets.add After:=wksrpt '增加應加在最後(目前的後面)
    End If
    If strWkName = MsgText(601) Then strWkName = Left(xlsAgentPoint.Worksheets(1).Name, Len(xlsAgentPoint.Worksheets(1).Name) - 1)
    Set wksrpt = xlsAgentPoint.Worksheets(strWkName & intXlsSheet)
    wksrpt.Activate
    
    With adoacc1k0
        Do While .EOF = False
            If bolFirst = True Then
                Call SetField(xlsAgentPoint.Version, wksrpt, intField, IIf(strOldAgNo = MsgText(601), .Fields("FagentNo"), strOldAgNo), bolShowCus)
                intTitleRow = intCounter
                intCounter = intCounter + 1: intStartR = intCounter
                bolFirst = False
            End If
            '顯示客戶
            If bolShowCus = True Then
                If (strOldAppNo <> MsgText(601) And strOldAppNo <> Mid(.Fields("CusNo"), 4)) Or (strOldAgNo <> MsgText(601) And strOldAgNo <> .Fields("FagentNo")) Then
                    '合計
                    If intCounter > intTitleRow + 1 Then Call SetLastSet(wksrpt, intCounter, intField, intStartR)
                    '改工作表名稱
                    If strOldAppNo <> Mid(.Fields("CusNo"), 4) Then
                        wksrpt.Name = strOldAppNo
                         intXlsSheet = intXlsSheet + 1
                         strOldAppNo = ""
                    End If
                    intCounter = 1
                    If strOldAgNo <> MsgText(601) And strOldAgNo <> .Fields("FagentNo") Then
                        strOldAgNo = .Fields("FagentNo")
                        If Val(xlsAgentPoint.Version) < 12 Then
                           xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
                        Else
                           xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
                        End If
                        xlsAgentPoint.Workbooks.Close
                        bolFirst = True
                        GoTo NextAg
                    Else
                        bolFirst = True
                        GoTo NextCus
                    End If
                End If
            End If
            '幣別合計 (Y27766除外)
            If Left(Text1, 6) <> "Y27766" And strOldCurr <> MsgText(601) And strOldCurr <> .Fields("Curr") And intCounter > intTitleRow + 1 Then
                Call SetLastSet(wksrpt, intCounter, intField, intStartR, True)
                intCounter = intCounter + 2
                intStartR = intCounter
            End If
            Call GetOtherData(strTmp(), "C") '2021/10/26 前以 stA1k32="C" 計算,故直接設C
            'Memo by Amy 因分不同格式,若任一種格式修改時都需詢問財務另一個是否也需調整
            'Y27766特殊格式
            If Left(Text1, 6) = "Y27766" Then
                Call ShowSpec(wksrpt, intCounter, intField)
            Else
                Call ShowOrg(wksrpt, intCounter, intField)
            End If
            
            intCounter = intCounter + 1: intNo = intNo + 1
            If bolShowCus = True Then
                strOldAgNo = "" & .Fields("FagentNo")
                strOldAppNo = Mid("" & .Fields("CusNo"), 4)
            End If
            strOldCurr = "" & .Fields("Curr")
            .MoveNext
        Loop
        '合計
        If Left(Text1, 6) = "Y27766" Then
            Call SetLastSet(wksrpt, intCounter, intField)
        Else
            Call SetLastSet(wksrpt, intCounter, intField, intStartR) '依幣別合計
            If bolChina = True And strOldAppNo <> MsgText(601) Then
                wksrpt.Name = strOldAppNo
            ElseIf bolShowCus = True Then
               wksrpt.Name = strOldAppNo
            End If
        End If
    End With
    
    If Val(xlsAgentPoint.Version) < 12 Then
       xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
    Else
       xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
    End If
    xlsAgentPoint.Workbooks.Close
    xlsAgentPoint.Quit
    StatusClear
    Exit Sub
 
ErrHand:
    MsgBox Err.Description, , MsgText(5)
    If Val(xlsAgentPoint.Version) < 12 Then
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
    Else
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
    End If
    xlsAgentPoint.Workbooks.Close
    xlsAgentPoint.Quit
    Set wksrpt = Nothing
    Set xlsAgentPoint = Nothing
End Sub

'Y27766欄位格式
Private Sub ShowSpec(ByRef wksrpt As Worksheet, ByVal intCount As Integer, ByVal intField As Integer)
    Dim bolFormula As Boolean, strValue As String
    
    With adoacc1k0
        For i = 0 To UBound(strF)
            bolFormula = False: strValue = ""
            Select Case i
                Case GetValue("NO.") '序號
                    strValue = intNo
                Case GetValue("Law Firm's Name<select>")
                    strValue = "Tai E"
                Case GetValue("Date<mm/dd/yyy>")
                    strValue = ChangeTStringToWDateString(.Fields("A1K02"))
                Case GetValue("Invoice No.") 'A1K01
                    strValue = .Fields("DocNo")
                Case GetValue("Our Ref") '本所案號
                    strValue = .Fields("CaseNo1") & "-" & .Fields("CaseNo2")
                    'Modify by Amy 2024/06/18 商標查名請款單由TS或S 案轉入者,在O/REF.NO.之本所案號後再加入"(查名本所案號)" ex:FCT-051898(S-008316)
                    If InStr("" & adoacc1k0.Fields("CaseNo1"), "T") > 0 And "" & adoacc1k0.Fields("CaseNo1") <> "TS" Then
                        strValue = strValue & GetTMQCaseNo(True, "" & .Fields("DocNo"))
                    End If
                Case GetValue("Your Ref") '彼所案號
                    strValue = strTmp(0)
                Case GetValue("Application No.(Only for new case)") '該案號為第一張請款單之申請號
                    strValue = strTmp(1)
                Case GetValue("Filing Date(Only for new case)") '該案號為第一張請款單
                    strValue = strTmp(2)
                Case GetValue("TradeMark/Patent Application No.") '申請號
                    strValue = strTmp(3)
                Case GetValue("Filing Date")
                    strValue = strTmp(4)
                Case GetValue("Category<select>")
                    strValue = ""
                'Modify by Amy 2022/07/20 申請人英文顯示錯誤 原:Application
                Case GetValue("Applicant No.") '客戶編號
                    If bolShowCus = True Then strValue = "" & .Fields("CusNo")
                Case GetValue("Applicant") '客戶名稱
                'end 2022/07/20
                    If bolShowCus = True Then strValue = "" & .Fields("CusName")
                Case GetValue("Currency<select>") 'A1K18
                    strValue = .Fields("Curr")
                Case GetValue("Attorney Fee")
                    bolFormula = True
                    strValue = "=" & Chr(GetValue("Total Fee") + intField) & intCounter & "-" & _
                                Chr(GetValue("Official Fee") + intField) & intCounter & "-" & Chr(GetValue("Disburesment Fee") + intField) & intCounter
                Case GetValue("Official Fee")
                    strValue = strTmp(5)
                Case GetValue("Disburesment Fee")
                    strValue = strTmp(6)
                Case GetValue("Total Fee") 'A1K08-A1K31
                    strValue = Val(.Fields("FAmount"))
            End Select
             
            If bolFormula = True Then
                wksrpt.Range(Chr(i + intField) & intCounter).Formula = strValue
            ElseIf i = GetValue("Total Fee") Then
                '先把帳單金額寫入,後面再調整為代理人要的公式
                wksrpt.Range(Chr(i + intField) & intCounter).Value = strValue
                'Attorney Fee 不顯示公式,改顯示值
                wksrpt.Range(Chr(GetValue("Attorney Fee") + intField) & intCounter).Copy
                wksrpt.Range(Chr(GetValue("Attorney Fee") + intField) & intCounter).PasteSpecial xlPasteValues, xlPasteSpecialOperationNone, False, False
                'Total Fee 改顯示公式 (2013/07/31 15:48 婧瑄 mail-同代理人給的格式)
                strValue = "=IF(" & Chr(GetValue("Attorney Fee") + intField) & intCounter & "+" & Chr(GetValue("Official Fee") + intField) & intCounter & _
                                   "+" & Chr(GetValue("Disburesment Fee") + intField) & intCounter & "=0,"""", " & Chr(GetValue("Attorney Fee") + intField) & intCounter & "+" & Chr(GetValue("Official Fee") + intField) & intCounter & _
                                    "+" & Chr(GetValue("Disburesment Fee") + intField) & intCounter & ")"
                wksrpt.Range(Chr(i + intField) & intCounter).Value = strValue
            Else
                If i = GetValue("Application No.(Only for new case)") Or i = GetValue("TradeMark/Patent Application No.") Then
                    wksrpt.Range(Chr(i + intField) & intCounter).NumberFormatLocal = "@"
                End If
                wksrpt.Range(Chr(i + intField) & intCounter).Value = strValue
            End If
            
        Next i
    End With
End Sub

'非 Y27766欄位格式
Private Sub ShowOrg(ByRef wksrpt As Worksheet, ByVal intCount As Integer, ByVal intField As Integer)
    Dim bolFormula As Boolean, strValue As String
    
     With adoacc1k0
        For i = 0 To UBound(strF)
            bolFormula = False: strValue = ""
            Select Case i
                Case GetValue("NO."), GetValue("編號") '序號
                    strValue = intNo
                Case GetValue("Law Firm's Name<select>"), GetValue("本事務所名稱")
                    strValue = "Tai E"
                Case GetValue("Date<mm/dd/yyy>"), GetValue("帳單日期<mm/dd/yyy>")
                    strValue = ChangeTStringToWDateString(.Fields("A1K02"))
                Case GetValue("Invoice No."), GetValue("帳單編號") 'A1K01
                    strValue = .Fields("DocNo")
                Case GetValue("Our Ref"), GetValue("本所案號") '本所案號
                    strValue = .Fields("CaseNo1") & "-" & .Fields("CaseNo2")
                    'Modify by Amy 2024/06/18 商標查名請款單由TS或S 案轉入者,在O/REF.NO.之本所案號後再加入"(查名本所案號)" ex:FCT-051898(S-008316)
                    If InStr("" & adoacc1k0.Fields("CaseNo1"), "T") > 0 And "" & adoacc1k0.Fields("CaseNo1") <> "TS" Then
                        strValue = strValue & GetTMQCaseNo(True, "" & .Fields("DocNo"))
                    End If
                Case GetValue("Your Ref"), GetValue("貴所案號") '彼所案號
                    strValue = strTmp(0)
                Case GetValue("Application No.(Only for new case)") '該案號為第一張請款單之申請號
                    strValue = strTmp(1)
                Case GetValue("Filing Date(Only for new case)") '該案號為第一張請款單
                    strValue = strTmp(2)
                Case GetValue("TradeMark/Patent Application No."), GetValue("商標/專利 申請號") '申請號
                    strValue = strTmp(3)
                Case GetValue("Filing Date")
                    strValue = strTmp(4)
                Case GetValue("Category<select>")
                    strValue = ""
                'Modify by Amy 2022/07/20 申請人英文顯示錯誤 原:Application
                Case GetValue("Applicant No."), GetValue("客戶編號") '客戶編號
                    If bolShowCus = True Then strValue = "" & .Fields("CusNo")
                Case GetValue("Applicant"), GetValue("客戶名稱") '客戶名稱
                'end 2022/07/20
                    If bolShowCus = True Then strValue = "" & .Fields("CusName")
                Case GetValue("Currency<select>"), GetValue("幣別") 'A1K18
                    strValue = .Fields("Curr")
                Case GetValue("Attorney Fee"), GetValue("服務費")
                    bolFormula = True
                    If bolChina = False Then
                        strValue = "=" & Chr(GetValue("Total Fee") + intField) & intCounter & "-" & _
                                Chr(GetValue("Official Fee") + intField) & intCounter & "-" & Chr(GetValue("Disburesment Fee") + intField) & intCounter
                    Else
                        strValue = "=" & Chr(GetValue("帳單金額") + intField) & intCounter & "-" & _
                                Chr(GetValue("規費") + intField) & intCounter & "-" & Chr(GetValue("雜費") + intField) & intCounter
                    End If
                Case GetValue("Official Fee"), GetValue("規費")
                    strValue = strTmp(5)
                Case GetValue("Disburesment Fee"), GetValue("雜費")
                    strValue = strTmp(6)
                Case GetValue("Total Fee"), GetValue("帳單金額") 'A1K08-A1K31
                    strValue = Val(.Fields("FAmount"))
                Case GetValue("USD")
                    '列印幣別格式=4.外幣+美金,顯示a1k38(美金請款金額)
                    If "" & .Fields("a1k33") = "4" Then
                        strValue = Val("" & .Fields("a1k38"))
                    End If
            End Select
            
            If bolFormula = True Then
                wksrpt.Range(Chr(i + intField) & intCounter).Formula = strValue
            Else
                If i = GetValue("Application No.(Only for new case)") Or i = GetValue("TradeMark/Patent Application No.") Or i = GetValue("商標/專利 申請號") Then
                    wksrpt.Range(Chr(i + intField) & intCounter).NumberFormatLocal = "@"
                End If
                wksrpt.Range(Chr(i + intField) & intCounter).Value = strValue
            End If
            '若列印幣別格式為4,且特殊請款單=C(整批)顯示顏色
            If i = GetValue("USD") And "" & .Fields("a1k32") = "C" And "" & .Fields("a1k33") = "4" Then
                wksrpt.Range(Chr(i + intField) & intCounter).Interior.ColorIndex = 40   '膚色
                wksrpt.Range(Chr(i + intField) & intCounter).Interior.tintandshade = 0.2 '設深淺
            End If
        Next i
    End With
End Sub

'Modify by Amy 2024/06/18 取得轉案前的 查名本所案號
Private Function GetTMQCaseNo(ByVal IsExcel As Boolean, ByVal stCP60 As String, Optional ByRef bolTMQ As Boolean = False) As String
   Dim RsQ As New ADODB.Recordset, strQ As String, intQ As Integer
   Dim stTP(1) As String, stCaseNo(1 To 4) As String
   
   GetTMQCaseNo = ""
   strQ = "Select cp10,cp64 From CaseProgress Where cp60='" & stCP60 & "' Order by cp10"
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      RsQ.MoveFirst
      'FCT-050288 [查名]與[申請]請款單號都為 X11202384,屬於轉號後才有請款單號,[不需]顯示原查名本所案號-秀玲
      If RsQ.RecordCount = 1 And "" & RsQ.Fields("cp10") = "001" And InStr("" & RsQ.Fields("cp64"), "原查名本所案號") > 0 Then
         bolTMQ = True
         stTP(0) = "" & RsQ.Fields("cp64")
         stTP(0) = Mid(Mid(stTP(0), InStr(stTP(0), "原查名本所案號：")), 9)
         stCaseNo(1) = SystemNumber(stTP(0), 1)
         stCaseNo(2) = SystemNumber(stTP(0), 2)
         stCaseNo(3) = SystemNumber(stTP(0), 3)
         stCaseNo(4) = Left(SystemNumber(stTP(0), 4), 2)
         GetTMQCaseNo = stCaseNo(1) & "-" & stCaseNo(2)
         If Not (stCaseNo(3) = "0" And stCaseNo(4) = "00") Then
             GetTMQCaseNo = GetTMQCaseNo & stCaseNo(3) & "-" & stCaseNo(4)
         End If
         GetTMQCaseNo = "(" & GetTMQCaseNo & ")"
         If IsExcel = True Then GetTMQCaseNo = vbCrLf & GetTMQCaseNo
      End If
   End If
   
   Set RsQ = Nothing
End Function

'Add by Amy 2024/08/02 CSV特殊格式代理,產生 CSV格式-斯閔
Private Function CSVSave(Optional ByRef pFileName As String) As Boolean
   Dim Xls As New Excel.Application, Wk As New Worksheet, intField As Integer, intRow As Integer, intXlsSheet As Integer
   Dim IsOpen As Boolean, strAllField As String, strAllWidth As String, strWkName As String, XlsFileN As String, strTp As String
   
On Error GoTo ErrH
   CSVSave = False
   
   Select Case Left(Text1, 6)
      Case "Y53715"
         strAllField = "Our Ref,Your Ref,Invoice Date,Invoice No,Total Fee"
         strAllWidth = "13, 13, 13, 13, 13"
         XlsFileN = Text1 & " SoA " & ServerDate & ".csv" '檔名-斯閔
      Case Else
   End Select
   strF = Split(strAllField, ",")
   intWidth = Split(strAllWidth, ",")
    
   If Dir(strExcelPath & XlsFileN) = MsgText(601) Then
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
         MkDir strExcelPath
      End If
   Else
      Kill strExcelPath & XlsFileN
   End If
   pFileName = strExcelPath & XlsFileN
    
   intXlsSheet = 1: intField = 65: intRow = 1
    
   Xls.Workbooks.add
   IsOpen = True
   Xls.Application.WindowState = xlMinimized
   If strWkName = MsgText(601) Then strWkName = Left(Xls.Worksheets(1).Name, Len(Xls.Worksheets(1).Name) - 1)
   Set Wk = Xls.Worksheets(strWkName & intXlsSheet)
   Wk.Activate
   
   With adoacc1k0
      .MoveFirst
      Do While .EOF = False
         For i = LBound(strF) To UBound(strF)
            strTp = ""
            Call GetOtherData(strTmp(), "C")
            Select Case strF(i)
               Case "Our Ref"
                  strTp = .Fields("CaseNo1") & "-" & .Fields("CaseNo2")
               Case "Your Ref"
                  strTp = strTmp(0)
               Case "Invoice Date"
                  strTp = ChangeTStringToWDateString(.Fields("A1K02"))
               Case "Invoice No"
                  strTp = .Fields("DocNo")
               Case "Total Fee"
                  strTp = Val(.Fields("FAmount"))
            End Select
            Wk.Range(Chr(intField + i) & intRow).Value = strTp
         Next i
         intRow = intRow + 1
         .MoveNext
      Loop
   End With
  
    '存CSV (xlCSV/xlCSVUTF8)
    Xls.Workbooks(1).SaveAs FileName:=strExcelPath & XlsFileN, FileFormat:=xlCSV
    CSVSave = True
    Xls.Workbooks.Close
    Xls.Quit
    StatusClear
    Exit Function
 
ErrH:
    MsgBox Err.Description, , MsgText(5)
    If IsOpen = True Then
      Xls.Workbooks(1).SaveAs FileName:=strExcelPath & XlsFileN, FileFormat:=xlCSV
      Xls.Workbooks.Close
      Xls.Quit
      Set Wk = Nothing
      Set Xls = Nothing
    End If
End Function

'Added by Lydia 2024/12/31 Excel列印-催款單信頭、信尾
Private Function PrintExcel_BFile(ByVal bolOpenFile As Boolean, ByVal iPicNo1 As Integer, Optional ByVal iPicNo2 As Integer) As Boolean
Dim strPic01 As String, strPic02 As String '下載檔案路徑:信頭Pic01、信尾Pic02

   If bolOpenFile = True Then
      strPrtFile = strPrtPath & "\$" & Me.Caption & MsgText(43)
      If Dir(strPrtFile) <> "" Then
         Kill strPrtFile
      End If
      xlsRpt.SheetsInNewWorkbook = 1
      xlsRpt.Workbooks.add
      Set WksRpt1 = xlsRpt.Worksheets(1)
      WksRpt1.Activate
      If Val(xlsRpt.Version) < 12 Then
         xlsRpt.Workbooks(1).SaveAs FileName:=strPrtFile, FileFormat:=-4143
      Else
         xlsRpt.Workbooks(1).SaveAs FileName:=strPrtFile, FileFormat:=56
      End If
      WksRpt1.PageSetup.Orientation = xlPortrait '直印
      WksRpt1.PageSetup.Zoom = 100 '縮放比例為100%
      WksRpt1.PageSetup.HeaderMargin = Excel.Application.InchesToPoints(0.1) '頁首
      WksRpt1.PageSetup.FooterMargin = Excel.Application.InchesToPoints(0.1) '頁尾
      WksRpt1.PageSetup.TopMargin = xlsRpt.InchesToPoints(0.1) '上
      WksRpt1.PageSetup.BottomMargin = xlsRpt.InchesToPoints(0.1) '下
      WksRpt1.PageSetup.LeftMargin = xlsRpt.InchesToPoints(0.1) '左邊界
      WksRpt1.PageSetup.RightMargin = xlsRpt.InchesToPoints(0.1) '右邊界
      xlsRpt.Visible = False
   Else
      If iPageNo = 0 Then  '刪除前一張收據的內容
         WksRpt1.Shapes.SelectAll
         xlsRpt.Selection.Delete  '刪除所有圖片
         WksRpt1.Range(Chr(xCols - 1) & ":" & Chr(xColE + 1)).Select
         xlsRpt.Selection.Delete  '刪除文字
         WksRpt1.Range("A1").Select
      Else
         '跨頁不清除
      End If
   End If
'-------------------欄寬和列高-----------------------------
   xCols = 66  'B欄
   
   If bolChina = True Then  '大陸格式
'---------中文
      maxRows = 45
      xColE = 65 + 7
      If iPageNo = 0 Then
         For intI = 0 To 8
            WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).Font.Name = "Times New Roman"
            WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).Font.Size = 12
            WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).Font.Bold = False
            Select Case intI
               Case 0, 8  'A,I
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 3.75
               Case 1 'B
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 9
               Case 1, 2 'B,C
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 11
               Case 3 'D
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 14
               Case 4 'E
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 15
               Case 5 'F
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 12
               Case 6, 7  'G,H
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 13.5
            End Select
         Next intI
      End If
   Else    '國外代理人
'---------英文
      maxRows = 45
      xColE = 65 + 7
      If iPageNo = 0 Then
         For intI = 0 To 8
            WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).Font.Name = "Times New Roman"
            WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).Font.Size = 12
            WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).Font.Bold = False
            Select Case intI
               Case 0, 8  'A,I
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 4
               Case 1  'B
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 9
               Case 2  'C
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 11
               Case 3 'D
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 15
               Case 4 'E
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 17
               Case 5, 6  'F,G,H
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 12
               Case 7   'H
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 11.5
            End Select
         Next intI
      End If
   End If 'If bolChina = True Then
   For intI = 1 To maxRows
      If intI = 1 Then  '信頭
         WksRpt1.Range(intI + iPageNo * maxRows & ":" & intI + iPageNo * maxRows).RowHeight = 110 '列高=3.87CM
      ElseIf intI = maxRows Then '信尾
         WksRpt1.Range(intI + iPageNo * maxRows & ":" & intI + iPageNo * maxRows).RowHeight = 35  '列高=1.23CM
      ElseIf bolChina = False And intI >= maxRows - 7 And intI <= maxRows - 2 Then  '英文：匯款銀行資料
         WksRpt1.Range(intI + iPageNo * maxRows & ":" & intI + iPageNo * maxRows).RowHeight = 15
      ElseIf bolChina = False And intI = maxRows - 8 Then        '英文：匯款銀行資料
         WksRpt1.Range(intI + iPageNo * maxRows & ":" & intI + iPageNo * maxRows).RowHeight = 20
      Else
         WksRpt1.Range(intI + iPageNo * maxRows & ":" & intI + iPageNo * maxRows).RowHeight = 16
      End If
      WksRpt1.Range(intI + iPageNo * maxRows & ":" & intI + iPageNo * maxRows).NumberFormat = "@" '預設文字格式
      WksRpt1.Range(intI + iPageNo * maxRows & ":" & intI + iPageNo * maxRows).HorizontalAlignment = xlLeft
   Next intI

   xRows = (iPageNo * maxRows) + 2
   If bolChina = True Then
      xRowE = ((iPageNo + 1) * maxRows) - 15 '列印表格的終止位置
   Else
      xRowE = ((iPageNo + 1) * maxRows) - 12 '列印表格的終止位置
   End If
   nRow = xRows '目前
'-------------------欄寬和列高-----------------------------
   If PrintExcel_Head(WksRpt1, iPageNo) = False Then
      MsgBox "代理人名稱、地址列印錯誤:" & vbCrLf & Err.Description
      GoTo EXITSUB
   End If
   
   '信頭
   If iPicNo1 > 0 Then  '信頭
      strPic01 = strPrtPath & "\$Tmp01.jpg"
      If iPageNo = 0 Then
         If PUB_ReadDB2File(strPic01, iPicNo1) = True Then
         End If
      Else
         strExc(0) = Dir(strPic01)
         If strExc(0) = "" Then
            If PUB_ReadDB2File(strPic01, iPicNo1) = True Then
            End If
         End If
      End If
      Set oShape = WksRpt1.Shapes.AddPicture(strPic01, True, True, 0, WksRpt1.Cells((iPageNo * maxRows) + 1, "A").Top, xlsRpt.CentimetersToPoints(19.5), xlsRpt.CentimetersToPoints(3.66))
   End If
   '信尾
   If iPicNo2 > 0 Then  '信尾
      strPic02 = strPrtPath & "\$Tmp02.jpg"
      If iPageNo = 0 Then
         If PUB_ReadDB2File(strPic02, iPicNo2) = True Then
         End If
      Else
         strExc(0) = Dir(strPic02)
         If strExc(0) = "" Then
            If PUB_ReadDB2File(strPic02, iPicNo2) = True Then
            End If
         End If
      End If
      Set oShape2 = WksRpt1.Shapes.AddPicture(strPic02, True, True, 0, WksRpt1.Cells(((iPageNo + 1) * maxRows), "A").Top + 2, xlsRpt.CentimetersToPoints(19.5), xlsRpt.CentimetersToPoints(0.91))
   End If
   iPageNo = iPageNo + 1
   
   PrintExcel_BFile = True
   Exit Function
   
EXITSUB:
   
End Function

'Added by Lydia 2024/12/31 Excel列印-頁首、表格
Private Function PrintExcel_Head(ByVal pNowWks, ByVal pPageNo As Integer) As Boolean
Dim strLanguage As String
Dim strFA17 As String, strFA18 As String, strFA19 As String, strFA20 As String, strFA21 As String, strFA22 As String, strFA70 As String, strFA23 As String
Dim strFA32 As String, strFA33 As String, strFA34 As String, strFA35 As String, strFA36 As String
Dim intS As Integer
   
On Error GoTo ErrHandle

   'Add by Amy 2018/10/31 +地址有「竹曆退件」字樣不顯示地址
   strFA17 = "" & adoacc1k0.Fields("fa17").Value
   strFA18 = "" & adoacc1k0.Fields("fa18").Value: strFA19 = "" & adoacc1k0.Fields("fa19").Value: strFA20 = "" & adoacc1k0.Fields("fa20").Value
   strFA21 = "" & adoacc1k0.Fields("fa21").Value: strFA22 = "" & adoacc1k0.Fields("fa22").Value: strFA70 = "" & adoacc1k0.Fields("fa70").Value
   strFA23 = "" & adoacc1k0.Fields("fa23").Value
   strFA32 = "" & adoacc1k0.Fields("fa32").Value: strFA33 = "" & adoacc1k0.Fields("fa33").Value: strFA34 = "" & adoacc1k0.Fields("fa34").Value
   strFA35 = "" & adoacc1k0.Fields("fa35").Value: strFA36 = "" & adoacc1k0.Fields("fa36").Value
   
   If InStr(strFA17, "竹曆退件") > 0 Then strFA17 = ""
   If InStr(strFA18 & strFA19 & strFA20 & strFA21 & strFA22 & strFA70, "竹曆退件") > 0 Then
     strFA18 = "": strFA19 = "": strFA20 = "": strFA21 = "": strFA22 = "": strFA70 = ""
   End If
   If InStr(strFA23, "竹曆退件") > 0 Then strFA23 = ""
   If InStr(strFA32 & strFA33 & strFA34 & strFA35 & strFA36, "竹曆退件") > 0 Then
     strFA32 = "": strFA33 = "": strFA34 = "": strFA35 = "": strFA36 = ""
   End If
            
   If bolChina = True Then '大陸格式: PrintHead1A4
'----------中文
      '代理人代號
      pNowWks.Range(Chr(xCols) & nRow).Value = "代理人代號: " & adoacc1k0.Fields("FagentNo")
      nRow = nRow + 1
      intS = 22
      '代理人名稱: 中文->英文->日文
      strData = "代理人名稱: "
      If "" & adoacc1k0.Fields("fa04").Value <> "" Then
         pNowWks.Range(Chr(xCols) & nRow).Value = strData & adoacc1k0.Fields("fa04").Value
         nRow = nRow + 1
      ElseIf "" & adoacc1k0.Fields("fa05").Value <> "" Then
         pNowWks.Range(Chr(xCols) & nRow).Value = strData & adoacc1k0.Fields("fa05").Value
         nRow = nRow + 1
         If "" & adoacc1k0.Fields("fa63").Value <> "" Then
            pNowWks.Range(Chr(xCols) & nRow).Value = Space(intS) & adoacc1k0.Fields("fa63").Value
            nRow = nRow + 1
         End If
         If "" & adoacc1k0.Fields("fa64").Value <> "" Then
            pNowWks.Range(Chr(xCols) & nRow).Value = Space(intS) & adoacc1k0.Fields("fa64").Value
            nRow = nRow + 1
         End If
         If "" & adoacc1k0.Fields("fa65").Value <> "" Then
            pNowWks.Range(Chr(xCols) & nRow).Value = Space(intS) & adoacc1k0.Fields("fa65").Value
            nRow = nRow + 1
         End If
      ElseIf "" & adoacc1k0.Fields("fa06").Value <> "" Then
         pNowWks.Range(Chr(xCols) & nRow).Value = strData & adoacc1k0.Fields("fa06").Value
         nRow = nRow + 1
      End If
      nRow = nRow + 1
      '代理人地址
      '地址,順序:中文->POB->英文->日文 --- Memo by Lydia 2025/09/25
      strData = "代理人地址: "
      If strFA17 <> "" Then
         intI = 0
         '中文地址
JumpToCAddr1:
         intI = intI + 1
         If LenB(strFA17) > 44 Then
            strExc(0) = PUB_StrToStr(strFA17, 42)
            strFA17 = Mid(strFA17, Len(strExc(0)) + 1)
            WksRpt1.Range(Chr(xCols) & nRow).Value = IIf(intI = 1, strData, Space(intS)) & strExc(0)
         Else
            WksRpt1.Range(Chr(xCols) & nRow).Value = IIf(intI = 1, strData, Space(intS)) & strFA17
            strFA17 = ""
         End If
         If strFA17 <> "" Then
            nRow = nRow + 1
            GoTo JumpToCAddr1
         End If
      'POB1~POB5
      ElseIf Trim(strFA32 & strFA33 & strFA34 & strFA35 & strFA36) <> "" Then
         If strFA32 <> "" Then
            WksRpt1.Range(Chr(xCols) & nRow).Value = strData & strFA32
            nRow = nRow + 1
         End If
         If strFA33 <> "" Then
            WksRpt1.Range(Chr(xCols) & nRow).Value = Space(intS) & strFA33
            nRow = nRow + 1
         End If
         If strFA34 <> "" Then
            WksRpt1.Range(Chr(xCols) & nRow).Value = Space(intS) & strFA34
            nRow = nRow + 1
         End If
         If strFA35 <> "" Then
            WksRpt1.Range(Chr(xCols) & nRow).Value = Space(intS) & strFA35
            nRow = nRow + 1
         End If
         If strFA36 <> "" Then
            WksRpt1.Range(Chr(xCols) & nRow).Value = Space(intS) & strFA36
            nRow = nRow + 1
         End If
      '英文地址1~6
      ElseIf Trim(strFA18 & strFA19 & strFA20 & strFA21 & strFA22 & strFA70) <> "" Then
         If strFA18 <> "" Then
            WksRpt1.Range(Chr(xCols) & nRow).Value = strData & strFA18
            nRow = nRow + 1
         End If
         If strFA19 <> "" Then
            WksRpt1.Range(Chr(xCols) & nRow).Value = Space(intS) & strFA19
            nRow = nRow + 1
         End If
         If strFA20 <> "" Then
            WksRpt1.Range(Chr(xCols) & nRow).Value = Space(intS) & strFA20
            nRow = nRow + 1
         End If
         If strFA21 <> "" Then
            WksRpt1.Range(Chr(xCols) & nRow).Value = Space(intS) & strFA21
            nRow = nRow + 1
         End If
         If strFA22 <> "" Then
            WksRpt1.Range(Chr(xCols) & nRow).Value = Space(intS) & strFA22
            nRow = nRow + 1
         End If
         If strFA70 <> "" Then
            WksRpt1.Range(Chr(xCols) & nRow).Value = Space(intS) & strFA70
            nRow = nRow + 1
         End If
      ElseIf strFA23 <> "" Then
         '日文地址
         WksRpt1.Range(Chr(xCols) & nRow).Value = strData & strFA23
         nRow = nRow + 1
      End If
      
      nRow = nRow + 2
      WksRpt1.Range(nRow & ":" & nRow).RowHeight = 25
      WksRpt1.Range(Chr(xCols + 3) & nRow).Value = "應收帳款對帳單"
      WksRpt1.Range(Chr(xCols + 3) & nRow).Font.Size = 18
      nRow = nRow + 1
      strData = IIf(Me.MaskEdBox1.Text = "___/__/__", "", Format(ChangeTStringToWDateString(Replace(Me.MaskEdBox1.Text, "/", "")), "MM/DD/YY")) & _
               IIf(Me.MaskEdBox1.Text = "___/__/__" And Me.MaskEdBox2.Text = "___/__/__", "", " ～ ") & _
               IIf(Me.MaskEdBox2.Text = "___/__/__", "", Format(ChangeTStringToWDateString(Replace(Me.MaskEdBox2.Text, "/", "")), "MM/DD/YY"))
      WksRpt1.Range(Chr(xCols + 2) & nRow).Value = "帳款日期：" & strData
      strExc(0) = IIf(strCurr <> "", strCurr, "" & adoacc1k0.Fields("Curr"))
      
      Select Case strExc(0)
         Case "NTD": strData = "台幣"
         Case "USD": strData = "美金"
         Case "RMB": strData = "人民幣"
      End Select
      WksRpt1.Range(Chr(xCols + 5) & nRow).Value = "幣別：" & strData
      nRow = nRow + 1
      xRows = nRow '列印表格的起始位置
      '畫表格 printTable
       WksRpt1.Range(xRows & ":" & xRows).RowHeight = 30
      WksRpt1.Range(xRows & ":" & xRows).VerticalAlignment = xlCenter
      WksRpt1.Range(Chr(xCols) & xRows).Value = "帳款日期"
      WksRpt1.Range(Chr(xCols + 1) & xRows).Value = "帳單編號"
      WksRpt1.Range(Chr(xCols + 2) & xRows).Value = "我方文號"
      WksRpt1.Range(Chr(xCols + 3) & xRows).Value = "貴方文號"
      WksRpt1.Range(Chr(xCols + 4) & xRows).Value = "應收金額"
      WksRpt1.Range(Chr(xCols + 5) & xRows).Value = "案件名稱"
      WksRpt1.Range(Chr(xCols + 6) & xRows).Value = "申請人"
      nRow = nRow + 1
      '表頭框線
      WksRpt1.Range(Chr(xCols) & xRows & ":" & Chr(xColE) & xRows).Borders(xlEdgeTop).LineStyle = xlContinuous
      WksRpt1.Range(Chr(xCols) & xRows & ":" & Chr(xColE) & xRows).Borders(xlEdgeBottom).LineStyle = xlContinuous
      '欄位框線
      WksRpt1.Range(Chr(xCols) & xRows & ":" & Chr(xCols) & xRowE + 2).Borders(xlEdgeLeft).LineStyle = xlContinuous
      WksRpt1.Range(Chr(xCols + 1) & xRows & ":" & Chr(xCols + 1) & xRowE).Borders(xlEdgeLeft).LineStyle = xlContinuous
      WksRpt1.Range(Chr(xCols + 2) & xRows & ":" & Chr(xCols + 2) & xRowE).Borders(xlEdgeLeft).LineStyle = xlContinuous
      WksRpt1.Range(Chr(xCols + 3) & xRows & ":" & Chr(xCols + 3) & xRowE).Borders(xlEdgeLeft).LineStyle = xlContinuous
      WksRpt1.Range(Chr(xCols + 4) & xRows & ":" & Chr(xCols + 4) & xRowE + 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
      WksRpt1.Range(Chr(xCols + 5) & xRows & ":" & Chr(xCols + 5) & xRowE).Borders(xlEdgeLeft).LineStyle = xlContinuous
      WksRpt1.Range(Chr(xCols + 6) & xRows & ":" & Chr(xCols + 6) & xRowE).Borders(xlEdgeLeft).LineStyle = xlContinuous
      WksRpt1.Range(Chr(xCols + 6) & xRows & ":" & Chr(xCols + 6) & xRowE + 2).Borders(xlEdgeRight).LineStyle = xlContinuous
      '表尾框線
      WksRpt1.Range(Chr(xCols) & xRowE + 1 & ":" & Chr(xColE) & xRowE + 1).Borders(xlEdgeTop).LineStyle = xlContinuous
      WksRpt1.Range(Chr(xCols) & xRowE + 1 & ":" & Chr(xColE) & xRowE + 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
      WksRpt1.Range(Chr(xCols) & xRowE + 2 & ":" & Chr(xColE) & xRowE + 2).Borders(xlEdgeBottom).LineStyle = xlContinuous
      '------------------------
      '頁碼
      WksRpt1.Range(Chr(xCols + 3) & xRowE + 3).Value = pPageNo + 1
      WksRpt1.Range(Chr(xCols + 3) & xRowE + 3).HorizontalAlignment = xlRight
   Else
'----------英文 PrintHeadA4
      strExc(0) = "select fa31 as Lang from fagent where fa01 = '" & Mid(adoacc1k0.Fields("FagentNo").Value, 1, 8) & "' and fa02 = '" & Mid(adoacc1k0.Fields("FagentNo").Value, 9, 1) & "' union " & _
                  "select cu64 as Lang from customer where cu01 = '" & Mid(adoacc1k0.Fields("FagentNo").Value, 1, 8) & "' and cu02 = '" & Mid(adoacc1k0.Fields("FagentNo").Value, 9, 1) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strLanguage = "" & RsTemp.Fields("Lang")
      End If
      If strLanguage = "" Then strLanguage = "2"
      intS = 3
      '代理人代號
      WksRpt1.Range(Chr(xCols) & nRow).Value = "Account No: " & adoacc1k0.Fields("FagentNo")
      nRow = nRow + 1
      '列印日期
      WksRpt1.Range(Chr(xCols) & nRow).Value = "Date:"
      WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & Format(AFDate(ServerDate), "mmm. d, yyyy")
      nRow = nRow + 1
      '代理人名稱
      WksRpt1.Range(Chr(xCols) & nRow).Value = "Attention:"
      'WksRpt1.Range(Chr(xCols + 1) & nRow).Value = '顯示代理人名稱，最多5行
''''''''''''''''''''''''''''''''''
      Select Case strLanguage
         Case "1" '中文(中-->英-->日)
            If "" & adoacc1k0.Fields("fa04").Value <> "" Then
               pNowWks.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & adoacc1k0.Fields("fa04").Value
               nRow = nRow + 1
            ElseIf "" & adoacc1k0.Fields("fa05").Value <> "" Then
               pNowWks.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & adoacc1k0.Fields("fa05").Value
               nRow = nRow + 1
               If "" & adoacc1k0.Fields("fa63").Value <> "" Then
                  pNowWks.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & adoacc1k0.Fields("fa63").Value
                  nRow = nRow + 1
               End If
               If "" & adoacc1k0.Fields("fa64").Value <> "" Then
                  pNowWks.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & adoacc1k0.Fields("fa64").Value
                  nRow = nRow + 1
               End If
               If "" & adoacc1k0.Fields("fa65").Value <> "" Then
                  pNowWks.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & adoacc1k0.Fields("fa65").Value
                  nRow = nRow + 1
               End If
            ElseIf "" & adoacc1k0.Fields("fa06").Value <> "" Then
               pNowWks.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & adoacc1k0.Fields("fa06").Value
               nRow = nRow + 1
            End If
            '地址
            '地址,順序:中文->POB->英文->日文 --- Memo by Lydia 2025/09/25
            If strFA17 <> "" Then
JumpToCAddr21:
               intI = intI + 1
               If LenB(strFA17) > 44 Then
                  strExc(0) = PUB_StrToStr(strFA17, 42)
                  strFA17 = Mid(strFA17, Len(strExc(0)) + 1)
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strExc(0)
               Else
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA17
                  strFA17 = ""
               End If
               If strFA17 <> "" Then
                  nRow = nRow + 1
                  GoTo JumpToCAddr21
               End If
            'POB1~POB5
            ElseIf Trim(strFA32 & strFA33 & strFA34 & strFA35 & strFA36) <> "" Then
               If strFA32 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA32
                  nRow = nRow + 1
               End If
               If strFA33 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA33
                  nRow = nRow + 1
               End If
               If strFA34 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA34
                  nRow = nRow + 1
               End If
               If strFA35 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA35
                  nRow = nRow + 1
               End If
               If strFA36 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA36
                  nRow = nRow + 1
               End If
            '英文地址1~6
            ElseIf Trim(strFA18 & strFA19 & strFA20 & strFA21 & strFA22 & strFA70) <> "" Then
               If strFA18 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA18
                  nRow = nRow + 1
               End If
               If strFA19 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA19
                  nRow = nRow + 1
               End If
               If strFA20 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA20
                  nRow = nRow + 1
               End If
               If strFA21 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA21
                  nRow = nRow + 1
               End If
               If strFA22 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA22
                  nRow = nRow + 1
               End If
               If strFA70 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA70
                  nRow = nRow + 1
               End If
            ElseIf strFA23 <> "" Then
               '日文地址
               WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA23
               nRow = nRow + 1
            End If
               
         Case "2" '英文(英-->中-->日)
            If "" & adoacc1k0.Fields("fa05").Value <> "" Then
               pNowWks.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & adoacc1k0.Fields("fa05").Value
               nRow = nRow + 1
               If "" & adoacc1k0.Fields("fa63").Value <> "" Then
                  pNowWks.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & adoacc1k0.Fields("fa63").Value
                  nRow = nRow + 1
               End If
               If "" & adoacc1k0.Fields("fa64").Value <> "" Then
                  pNowWks.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & adoacc1k0.Fields("fa64").Value
                  nRow = nRow + 1
               End If
               If "" & adoacc1k0.Fields("fa65").Value <> "" Then
                  pNowWks.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & adoacc1k0.Fields("fa65").Value
                  nRow = nRow + 1
               End If
            ElseIf "" & adoacc1k0.Fields("fa04").Value <> "" Then
               pNowWks.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & adoacc1k0.Fields("fa04").Value
               nRow = nRow + 1
            ElseIf "" & adoacc1k0.Fields("fa06").Value <> "" Then
               pNowWks.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & adoacc1k0.Fields("fa06").Value
               nRow = nRow + 1
            End If
            '地址
            '地址,順序:POB->英文->中文->日文 --- Memo by Lydia 2025/09/25
            'POB1~POB5
            If Trim(strFA32 & strFA33 & strFA34 & strFA35 & strFA36) <> "" Then
               If strFA32 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA32
                  nRow = nRow + 1
               End If
               If strFA33 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA33
                  nRow = nRow + 1
               End If
               If strFA34 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA34
                  nRow = nRow + 1
               End If
               If strFA35 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA35
                  nRow = nRow + 1
               End If
               If strFA36 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA36
                  nRow = nRow + 1
               End If
            '英文地址1~6
            ElseIf Trim(strFA18 & strFA19 & strFA20 & strFA21 & strFA22 & strFA70) <> "" Then
               If strFA18 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA18
                  nRow = nRow + 1
               End If
               If strFA19 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA19
                  nRow = nRow + 1
               End If
               If strFA20 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA20
                  nRow = nRow + 1
               End If
               If strFA21 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA21
                  nRow = nRow + 1
               End If
               If strFA22 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA22
                  nRow = nRow + 1
               End If
               If strFA70 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA70
                  nRow = nRow + 1
               End If
            ElseIf strFA17 <> "" Then
JumpToCAddr22:
               intI = intI + 1
               If LenB(strFA17) > 44 Then
                  strExc(0) = PUB_StrToStr(strFA17, 42)
                  strFA17 = Mid(strFA17, Len(strExc(0)) + 1)
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strExc(0)
               Else
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA17
                  strFA17 = ""
               End If
               If strFA17 <> "" Then
                  nRow = nRow + 1
                  GoTo JumpToCAddr22
               End If
            ElseIf strFA23 <> "" Then
               '日文地址
               WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA23
               nRow = nRow + 1
            End If
           
         Case "3" '日文(日-->中-->英)
            If "" & adoacc1k0.Fields("fa06").Value <> "" Then
               pNowWks.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & adoacc1k0.Fields("fa06").Value
               nRow = nRow + 1
            ElseIf "" & adoacc1k0.Fields("fa04").Value <> "" Then
               pNowWks.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & adoacc1k0.Fields("fa04").Value
               nRow = nRow + 1
            ElseIf "" & adoacc1k0.Fields("fa05").Value <> "" Then
               pNowWks.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & adoacc1k0.Fields("fa05").Value
               nRow = nRow + 1
               If "" & adoacc1k0.Fields("fa63").Value <> "" Then
                  pNowWks.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & adoacc1k0.Fields("fa63").Value
                  nRow = nRow + 1
               End If
               If "" & adoacc1k0.Fields("fa64").Value <> "" Then
                  pNowWks.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & adoacc1k0.Fields("fa64").Value
                  nRow = nRow + 1
               End If
               If "" & adoacc1k0.Fields("fa65").Value <> "" Then
                  pNowWks.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & adoacc1k0.Fields("fa65").Value
                  nRow = nRow + 1
               End If
            End If
            '地址
            '地址,順序:日文->中文->POB->英文 --- Memo by Lydia 2025/09/25
            If strFA23 <> "" Then
               '日文地址
               WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA23
               nRow = nRow + 1
            ElseIf strFA17 <> "" Then
JumpToCAddr23:
               intI = intI + 1
               If LenB(strFA17) > 44 Then
                  strExc(0) = PUB_StrToStr(strFA17, 42)
                  strFA17 = Mid(strFA17, Len(strExc(0)) + 1)
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strExc(0)
               Else
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA17
                  strFA17 = ""
               End If
               If strFA17 <> "" Then
                  nRow = nRow + 1
                  GoTo JumpToCAddr23
               End If
            'POB1~POB5
            ElseIf Trim(strFA32 & strFA33 & strFA34 & strFA35 & strFA36) <> "" Then
               If strFA32 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA32
                  nRow = nRow + 1
               End If
               If strFA33 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA33
                  nRow = nRow + 1
               End If
               If strFA34 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA34
                  nRow = nRow + 1
               End If
               If strFA35 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA35
                  nRow = nRow + 1
               End If
               If strFA36 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA36
                  nRow = nRow + 1
               End If
            '英文地址1~6
            ElseIf Trim(strFA18 & strFA19 & strFA20 & strFA21 & strFA22 & strFA70) <> "" Then
               If strFA18 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA18
                  nRow = nRow + 1
               End If
               If strFA19 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA19
                  nRow = nRow + 1
               End If
               If strFA20 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA20
                  nRow = nRow + 1
               End If
               If strFA21 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA21
                  nRow = nRow + 1
               End If
               If strFA22 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA22
                  nRow = nRow + 1
               End If
               If strFA70 <> "" Then
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = Space(intS) & strFA70
                  nRow = nRow + 1
               End If
            End If
      End Select
      
      'Added by Morgan 2020/2/19
      '下列請款對象若有財務編號也要印(專利->商標)
      'Modified by Morgan 2022/3/17 +Y55666--Ryan
      strData = ""
      If InStr("Y25061000,Y25061010,Y25061030,Y55363000,Y25061020,Y25061050,Y55666000", adoacc1k0.Fields("FagentNo").Value) > 0 Then
         strData = PUB_GetACCNO(adoacc1k0.Fields("FagentNo").Value)
      End If
      If strData <> "" Then
          WksRpt1.Range(Chr(xCols) & nRow + 1).Value = strData
      End If
      'end 2020/2/19
      nRow = nRow + 2
      WksRpt1.Range(nRow & ":" & nRow).RowHeight = 25
      WksRpt1.Range(Chr(xCols + 1) & nRow).Value = String(10, " ") & "STATEMENT OF ACCOUNT"
      WksRpt1.Range(Chr(xCols + 1) & nRow).Font.Size = 18
      strData = IIf(Me.MaskEdBox1.Text = "___/__/__", "", Format(ChangeTStringToWDateString(Replace(Me.MaskEdBox1.Text, "/", "")), "MM/DD/YY")) & _
                  IIf(Me.MaskEdBox1.Text = "___/__/__" And Me.MaskEdBox2.Text = "___/__/__", "", " ～ ") & _
                  IIf(Me.MaskEdBox2.Text = "___/__/__", "", Format(ChangeTStringToWDateString(Replace(Me.MaskEdBox2.Text, "/", "")), "MM/DD/YY"))
      WksRpt1.Range(Chr(xCols + 5) & nRow).Value = strData
      nRow = nRow + 1
      xRows = nRow '列印表格的起始位置
      '畫表格 printTable
      WksRpt1.Range(xRows & ":" & xRows).RowHeight = 30
      WksRpt1.Range(xRows & ":" & xRows).VerticalAlignment = xlCenter
      WksRpt1.Range(Chr(xCols) & xRows).Value = "DATE"
      WksRpt1.Range(Chr(xCols + 1) & xRows).Value = "DEBIT NO."
      WksRpt1.Range(Chr(xCols + 2) & xRows).Value = "O/REF. NO."
      WksRpt1.Range(Chr(xCols + 3) & xRows).Value = "Y/REF. NO."
      WksRpt1.Range(Chr(xCols + 4) & xRows).Value = "CREDIT"
      WksRpt1.Range(Chr(xCols + 5) & xRows).Value = "AMOUNT"
      WksRpt1.Range(Chr(xCols + 6) & xRows).Value = "DAYS SINCE DN.SENT"
      WksRpt1.Range(Chr(xCols + 6) & xRows).Font.Size = 10
      WksRpt1.Range(Chr(xCols + 6) & xRows).WrapText = True   '自動換列
      nRow = nRow + 1
      '表頭框線
      WksRpt1.Range(Chr(xCols) & xRows & ":" & Chr(xColE) & xRows).Borders(xlEdgeTop).LineStyle = xlContinuous
      WksRpt1.Range(Chr(xCols) & xRows & ":" & Chr(xColE) & xRows).Borders(xlEdgeBottom).LineStyle = xlContinuous
      '欄位框線
      WksRpt1.Range(Chr(xCols) & xRows & ":" & Chr(xCols) & xRowE).Borders(xlEdgeLeft).LineStyle = xlContinuous
      WksRpt1.Range(Chr(xCols + 1) & xRows & ":" & Chr(xCols + 1) & xRowE).Borders(xlEdgeLeft).LineStyle = xlContinuous
      WksRpt1.Range(Chr(xCols + 2) & xRows & ":" & Chr(xCols + 2) & xRowE).Borders(xlEdgeLeft).LineStyle = xlContinuous
      WksRpt1.Range(Chr(xCols + 3) & xRows & ":" & Chr(xCols + 3) & xRowE).Borders(xlEdgeLeft).LineStyle = xlContinuous
      WksRpt1.Range(Chr(xCols + 4) & xRows & ":" & Chr(xCols + 4) & xRowE).Borders(xlEdgeLeft).LineStyle = xlContinuous
      WksRpt1.Range(Chr(xCols + 5) & xRows & ":" & Chr(xCols + 5) & xRowE).Borders(xlEdgeLeft).LineStyle = xlContinuous
      WksRpt1.Range(Chr(xCols + 6) & xRows & ":" & Chr(xCols + 6) & xRowE).Borders(xlEdgeLeft).LineStyle = xlContinuous
      WksRpt1.Range(Chr(xCols + 6) & xRows & ":" & Chr(xCols + 6) & xRowE).Borders(xlEdgeRight).LineStyle = xlContinuous
      '表尾框線
      WksRpt1.Range(Chr(xCols) & xRowE & ":" & Chr(xColE) & xRowE).Borders(xlEdgeBottom).LineStyle = xlContinuous
      '頁碼
      WksRpt1.Range(Chr(xCols + 3) & xRowE + 11).Value = iPageNo + 1
      WksRpt1.Range(Chr(xCols + 3) & xRowE + 11).HorizontalAlignment = xlCenter
   End If   'If bolChina = True Then
   
   'Added by Lydia 2025/02/11 記錄份數
   If pPageNo = 0 Then
      m_iDocCount = m_iDocCount + 1
      If bolEmail = True Then
         m_iMailCount = m_iMailCount + 1
      End If
      If bol2Printer = True Then
         m_iPrintCount = m_iPrintCount + 1
      End If
   End If
   'end 2025/02/11
   
   PrintExcel_Head = True
   Exit Function
   
ErrHandle:

End Function

'Added by Lydia 2024/12/31 Excel列印-總計金額、匯款銀行資訊
Private Function PrintExcel_Footer() As Boolean
Dim tmpArr As Variant, tmpArr2 As Variant, inX As Integer
Dim strPdfA1k01 As String '記錄-請款單pdf檔路徑

On Error GoTo EXITSUB

   If bolChina = True Then '大陸格式
'---------中文
      '總計 PrintSum1A4
      nRow = xRowE + 1
      WksRpt1.Range(Chr(xCols + 2) & nRow).Value = "總　　　　計"
      WksRpt1.Range(Chr(xCols + 4) & nRow).NumberFormatLocal = FDollar
      WksRpt1.Range(Chr(xCols + 4) & nRow).HorizontalAlignment = xlRight
      WksRpt1.Range(Chr(xCols + 4) & nRow).Value = Format(douTAmount, FDollar)
      '應收帳款餘額
      nRow = nRow + 1
      WksRpt1.Range(Chr(xCols + 2) & nRow).Value = "應收帳款餘額"
      WksRpt1.Range(Chr(xCols + 3) & nRow).Value = strCurr
      WksRpt1.Range(Chr(xCols + 3) & nRow).HorizontalAlignment = xlRight
      WksRpt1.Range(Chr(xCols + 4) & nRow).NumberFormatLocal = FDollar
      WksRpt1.Range(Chr(xCols + 4) & nRow).HorizontalAlignment = xlRight
      WksRpt1.Range(Chr(xCols + 4) & nRow).Value = Format(douAmount, FDollar)
      nRow = nRow + 2 '跳過頁碼
      '匯款銀行資訊 PrintFooter1
      WksRpt1.Range(Chr(xCols) & nRow).Value = ReportSum(71001) 'Name of Bank:
      nRow = nRow + 1
      WksRpt1.Range(Chr(xCols) & nRow).Value = ReportSum(72) 'Address:
      nRow = nRow + 1
      WksRpt1.Range(Chr(xCols) & nRow).Value = ReportSum(73001) 'SWIFT Code:
      nRow = nRow + 1
      WksRpt1.Range(Chr(xCols) & nRow).Value = ReportSum(85) 'Account Name:
      nRow = nRow + 1
      WksRpt1.Range(Chr(xCols) & nRow).Value = ReportSum(74) 'Account No.:
      nRow = nRow + 1
      WksRpt1.Range(Chr(xCols) & nRow).Value = "※　貴公司于匯款后請務必將匯款憑證傳真至台北所，否則本所無法知悉　貴公司已匯款。"
      WksRpt1.Range(Chr(xCols) & nRow).Font.Bold = True
      nRow = nRow + 1
      WksRpt1.Range(Chr(xCols) & nRow).Value = "　　(傳真號碼：886 2 25011666)"
      WksRpt1.Range(Chr(xCols) & nRow).Font.Bold = True
   Else
'---------英文
      '總計 PrintSumA4
      nRow = xRowE + 1
      WksRpt1.Range(Chr(xCols + 2) & nRow).Value = String(18, " ") & "TOTAL AMOUNT DUE:"
      WksRpt1.Range(Chr(xCols + 2) & nRow).Font.Bold = True
      WksRpt1.Range(Chr(xCols + 4) & nRow).Value = strCurr
      WksRpt1.Range(Chr(xCols + 4) & nRow).Font.Bold = True
      WksRpt1.Range(Chr(xCols + 4) & nRow).HorizontalAlignment = xlRight
      WksRpt1.Range(Chr(xCols + 5) & nRow).NumberFormatLocal = FDollar
      WksRpt1.Range(Chr(xCols + 5) & nRow).HorizontalAlignment = xlRight
      WksRpt1.Range(Chr(xCols + 5) & nRow).Value = Format(douAmount, FDollar)
      WksRpt1.Range(Chr(xCols + 5) & nRow).Font.Bold = True
      nRow = nRow + 1
      '匯款銀行資訊 PrintFooter
      WksRpt1.Range(Chr(xCols + 3) & nRow).Value = "We would appreciate receiving your remittance"
      WksRpt1.Range(Chr(xCols + 3) & nRow).HorizontalAlignment = xlRight
      WksRpt1.Range(Chr(xCols + 4) & nRow).Value = "as soon as possible."
      WksRpt1.Range(Chr(xCols + 4) & nRow).Font.Underline = True
      nRow = nRow + 2
      WksRpt1.Range(nRow & ":" & nRow).RowHeight = 20
      WksRpt1.Range(Chr(xCols + 2) & nRow).Value = String(12, " ") & "REMITTANCE ADVICE"
      WksRpt1.Range(Chr(xCols + 2) & nRow).Font.Size = 16
      WksRpt1.Range(Chr(xCols + 2) & nRow).Font.Bold = True
      For intI = 1 To 6
         WksRpt1.Range(nRow + intI & ":" & nRow + intI).RowHeight = 15
         WksRpt1.Range(nRow + intI & ":" & nRow + intI).Font.Size = 10
         Select Case intI
            Case 1
               WksRpt1.Range(Chr(xCols) & nRow + intI).Value = "Checks to:"
               WksRpt1.Range(Chr(xCols) & nRow + intI).Font.Bold = True
               WksRpt1.Range(Chr(xCols + 3) & nRow + intI).Value = "Bankers:"
               WksRpt1.Range(Chr(xCols + 3) & nRow + intI).Font.Bold = True
            Case 2
               WksRpt1.Range(Chr(xCols) & nRow + intI).Value = ReportSum(116) 'Tai E International Patent & Law Office
               WksRpt1.Range(Chr(xCols) & nRow + intI).Font.Bold = True
               WksRpt1.Range(Chr(xCols + 3) & nRow + intI).Value = ReportSum(85) 'Account Name:
            Case 3
               WksRpt1.Range(Chr(xCols) & nRow + intI).Value = "P.O. Box 46-478, Taipei 104, Taiwan"
               WksRpt1.Range(Chr(xCols) & nRow + intI).Font.Bold = True
               WksRpt1.Range(Chr(xCols + 3) & nRow + intI).Value = ReportSum(74)  'Account No.:
            Case 4
               WksRpt1.Range(Chr(xCols + 3) & nRow + intI).Value = ReportSum(71001) 'Name of Bank:
            Case 5
               WksRpt1.Range(Chr(xCols + 3) & nRow + intI).Value = ReportSum(72) 'Address:
            Case 6
               WksRpt1.Range(Chr(xCols + 3) & nRow + intI).Value = ReportSum(73001) 'SWIFT Code:
         End Select
      Next intI
      '加外框
      WksRpt1.Range(Chr(xCols) & nRow & ":" & Chr(xColE) & nRow).Borders(xlEdgeTop).LineStyle = xlContinuous
      WksRpt1.Range(Chr(xCols) & nRow + 6 & ":" & Chr(xColE) & nRow + 6).Borders(xlEdgeBottom).LineStyle = xlContinuous
      WksRpt1.Range(Chr(xCols) & nRow & ":" & Chr(xCols) & nRow + 6).Borders(xlEdgeLeft).LineStyle = xlContinuous
      WksRpt1.Range(Chr(xColE) & nRow & ":" & Chr(xColE) & nRow + 6).Borders(xlEdgeRight).LineStyle = xlContinuous
   End If   'If bolChina = True Then
   
   'Added by Lydia 2016/09/08 改到催款單列印完後
   If strA1k01List <> "" Then
      tmpArr = Split(strA1k01List, ",")
      tmpArr2 = Split(strA1k01Dir, ",")
      For inX = 0 To UBound(tmpArr)
         If Trim(tmpArr(inX)) <> "" Then
             Load Frmacc2480
             With Frmacc2480
                .Text1.Text = Trim(tmpArr(inX))
                .Text2.Text = Trim(tmpArr(inX))
                .txtOutMode = "2"
                .m_bBeCalled = True
                .m_CallPrevForm = Me.Name 'Added by Lydia 2020/01/06 呼叫請款單的程式名稱
                .m_bEMail = True
                'Modified by Lydia 2017/03/02 一般催款單的資料夾用代理人區分
                If bolCallMail = False And strCallCase = "" Then  '排除催款通知,T收款寄證1728
                    .m_SavePath = IIf(Trim(tmpArr2(inX)) <> "", Trim(tmpArr2(inX)), strSavePath)
                Else
                    .m_SavePath = strSavePath
                End If
                .Command2_Click
             End With
             '請款單：判斷PDF檔案是否存在
             If Frmacc2480.m_strOutErr <> "" Then
                  m_strErr2480 = m_strErr2480 & Frmacc2480.m_strOutErr
             Else
             'end 2020/09/10
                  strPdfA1k01 = strPdfA1k01 & "*" & Frmacc2480.Tag  'Added by Lydia 2017/02/18 記錄-請款單pdf檔路徑
             End If
             Unload Frmacc2480
             strFormName = Me.Name
             tool3_enabled
         End If
      Next
   End If
   strA1k01List = ""
   strA1k01Dir = ""
   'Added by Lydia 2016/12/22　指定催款單範圍(T收款寄證1728),回傳附件路徑
   If strCallCase <> "" Then
      Me.Tag = strPicFileNames & strPdfA1k01
   Else
      Me.Tag = ""
   End If
   strPicFileNames = ""
   
   PrintExcel_Footer = True
   Exit Function
   
EXITSUB:
   If Err.Number <> 0 Then
      MsgBox "匯款銀行資料列印錯誤:" & vbCrLf & Err.Description
   End If
End Function

'Added by Lydia 2024/12/31 改用EXCEL：列印FC催款單
Private Sub PrintExcelMain()
Dim strDocNo As String
Dim StrSQLa As String, strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String
Dim strField As String
Dim hLocalFile As Long, strXLSFile As String
Dim bolReport As Boolean
Dim strKind As String, strKey1 As String, StrKey2 As String, strKey3 As String
Dim ff As Integer, strFileList(3) As String, strFileName As String
Dim strSubject As String, strDate As String
Dim strContent As String
Dim dblUsAmount As Double, dblFeeAmount As Double, dblPt As Double
Dim dblTotUsAmount As Double, dblTotFeeAmount As Double, dblTotPt As Double
Dim bolSum As Boolean '是否印依代理人加總欄位
Dim strSource As String, strDestination As String
Dim arrFile
Dim strMsg As String
Dim bolOpenXls As Boolean

   ClearQueryLog (Me.Name)
   bolShowCus = False
   strSQL1 = ""
   strSQL2 = ""
   
   If Text1 <> "" Then
      strSQL1 = strSQL1 & " and a1k28 >= '" & Text1 & "'"
      strSQL2 = strSQL2 & " and a1k28 >= '" & Text1 & "'"
   End If
   If Text2 <> "" Then
      strSQL1 = strSQL1 & " and a1k28 <= '" & Text2 & "'"
      strSQL2 = strSQL2 & " and a1k28 <= '" & Text2 & "'"
   End If
   If Text1 <> "" Or Text2 <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1 & Text1 & "-" & Text2
   End If
   If MaskEdBox1.Text <> "" And MaskEdBox1.Text <> MsgText(29) Then
      strSQL1 = strSQL1 & " and a1k02 >= " & ChgDateFormat(MaskEdBox1.Text) & ""
      strSQL2 = strSQL2 & " and a1k02 >= " & ChgDateFormat(MaskEdBox1.Text) & ""
   End If
   If MaskEdBox2.Text <> "" And MaskEdBox2.Text <> MsgText(29) Then
      strSQL1 = strSQL1 & " and a1k02 <= " & ChgDateFormat(MaskEdBox2.Text) & ""
      strSQL2 = strSQL2 & " and a1k02 <= " & ChgDateFormat(MaskEdBox2.Text) & ""
   End If
   If (MaskEdBox1.Text <> "" And MaskEdBox1.Text <> MsgText(29)) Or _
      (MaskEdBox2.Text <> "" And MaskEdBox2.Text <> MsgText(29)) Then
      pub_QL05 = pub_QL05 & ";" & Label4 & MaskEdBox1 & "-" & MaskEdBox2
   End If
   
   'Add by Amy 2017/02/02 +客戶編號(Y27766 Excel為特殊所以不加)
   If (Text10 <> "" Or Text11 <> "") And Left(Text1, 6) <> "Y27766" Then
        bolShowCus = True
        strField = "'" & strUserNum & "',"
        pub_QL05 = pub_QL05 & ";" & Label8 & Text10 & "-" & Text11
   End If

   'Add by Amy 2013/11/01 請款對象條件為空(整批)，且有設帳款處理者，不列印該筆請款單
   If Text1 = "" And Text2 = "" Then
        StrSQL3 = " and fa103 is null"
        StrSQL4 = " and cu142 is null"
   End If
   
   ' +有輸請款對象(單筆)且帳款處理情形有值顯示訊息
    If Text1 <> "" And Text2 <> "" Then
        strExc(0) = GetDizhang(Text1, Text2, True)
    End If
    'end 2013/10/31
    
   If Text6 <> "Y" Then
      'Added by Lydia 2024/10/28 增加不寄催款單:1.每月催款
      If Check2.Visible = True And Check2.Value = 1 Then
           strSQL1 = strSQL1 & " and nvl(fa101,'0')='1' "
           strSQL2 = strSQL2 & " and nvl(cu140,'0')='1' "
      Else
         'Add by Amy 2013/06/26 不存電子檔且不是產生excel,抓不寄催款單 為空值者
         If bolExcel <> True Then
           strSQL1 = strSQL1 & " and fa101 is null "
           strSQL2 = strSQL2 & " and cu140 is null "
         End If
      End If
      
      'Add by Morgan 2011/10/13 若請款對象起訖前6碼相同時先檢查是否有設定不寄催款單 --秀玲
      If Text1 <> "" And Left(Text1, 6) = Left(Text2, 6) Then
         strExc(0) = "select fa01||fa02 FNo from fagent where fa01 between '" & Left(Text1, 8) & "' and '" & Left(Text2, 8) & "' and fa101 is not null" & _
             " union select cu01||cu02 FNo from customer where cu01 between '" & Left(Text1, 8) & "' and '" & Left(Text2, 8) & "' and CU140 is not null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(1) = RsTemp.GetString(, , , vbCrLf)
            If MsgBox("下列請款對象有設定不寄催款單程式將自動略過，是否仍要繼續??" & vbCrLf & vbCrLf & strExc(1), vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Sub
            End If
         End If
      End If
      'end 2011/10/13
   End If
   
   strCon10 = ""
   adoacc1k0.CursorLocation = adUseClient
   'Modify by Amy 2013/11/01 +整批(編號空白)時有設帳款處理者不印
   'Mofified by Morgan 2016/2/5 +整批列印ACC1T0相關欄位
   'Modified by Morgan 2020/12/14 + and d.a1k25 is null(排除已銷帳請款單) Ex:X10805261-X10911362
   'Modified by Lydia 2022/03/04  只抓整批請款單尚未結清的請款單號 exists(select * => and a1k01 in (select d.a1k01
   'Modified by Lydia 2024/09/18  +財務副本信箱strEmailCC; 比照frmacc2450財務信箱CF, 優先抓FA105->FA79
   StrSQLa = "select " & strField & "a1k02, a1k28 as FagentNo, a1k02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, 0 as OAmount, (a1k08 - nvl(a1k31, 0)) as FAmount, a1k04 as Yno, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa17, fa43, a1k18 as Curr, a1k10 as Rate, FA10 as NATION,nvl(fa79,fa16) as EBox,FA108,fa70, a1k32, a1t01, a1t02,a1k33,a1k38,decode(fa105||fa79,null,'',fa134) as emailcc From acc1k0, fagent, acc1t0 where a1t01(+)=a1k01 and substr(a1k28, 1, 8) = fa01 and substr(a1k28, 9, 1) = fa02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) " & strSQL1 & StrSQL3 & " union " & _
             "select " & strField & "a1k02, a1k28 as FagentNo, a1k02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, 0 as OAmount, (a1k08 - nvl(a1k31, 0)) as FAmount, a1k04 as Yno, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu23 as fa17, cu76 as fa43, a1k18 as Curr, a1k10 as Rate, CU10 as NATION,nvl(cu115,cu20) as EBox,cu148 as FA108,cu102 fa70, a1k32, a1t01, a1t02,a1k33,a1k38,decode(cu115,null,'',cu200) as emailcc From acc1k0, customer , acc1t0 where a1t01(+)=a1k01 and  substr(a1k28, 1, 8) = cu01 and substr(a1k28, 9, 1) = cu02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) " & strSQL2 & StrSQL4 & " union " & _
             "select " & strField & "a1k02, a1k28 as FagentNo, a0y02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a0z04 as OAmount, 0 as FAmount, a1k04 as Yno, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa17, fa43, A0Y03 as Curr, a0y04 as Rate, FA10 as NATION,nvl(fa79,fa16) as EBox,FA108,fa70, a1k32, a1t01, a1t02,a1k33,a1k38,decode(fa105||fa79,null,'',fa134) as emailcc From acc0z0, acc0y0, acc1k0, fagent , acc1t0 where a1t01(+)=a1k01 and nvl(a1k32,'N')<>'C' and  a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = fa01 and substr(a1k28, 9, 1) = fa02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) " & strSQL1 & StrSQL3 & " union " & _
             "select " & strField & "a1k02, a1k28 as FagentNo, a0y02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a0z04 as OAmount, 0 as FAmount, a1k04 as Yno, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu23 as fa17, cu76 as fa43, A0Y03 as Curr, a0y04 as Rate, CU10 as NATION,nvl(cu115,cu20) as EBox,cu148 as FA108,cu102 as fa70, a1k32, a1t01, a1t02,a1k33,a1k38,decode(cu115,null,'',cu200) as emailcc From acc0z0, acc0y0, acc1k0, customer , acc1t0 where a1t01(+)=a1k01 and nvl(a1k32,'N')<>'C' and  a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = cu01 and substr(a1k28, 9, 1) = cu02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) " & strSQL2 & StrSQL4 & " union " & _
             "select " & strField & "a1k02, a1k28 as FagentNo, a0y02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a0z04 as OAmount, 0 as FAmount, a1k04 as Yno, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa17, fa43, A0Y03 as Curr, a0y04 as Rate, FA10 as NATION,nvl(fa79,fa16) as EBox,FA108,fa70, a1k32, a1t01, a1t02,a1k33,a1k38,decode(fa105||fa79,null,'',fa134) as emailcc From acc0z0, acc0y0, acc1k0 a, fagent , acc1t0 b where a1t01(+)=a1k01 and nvl(a1k32,'N')='C' and  a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = fa01 and substr(a1k28, 9, 1) = fa02 and a1k01 in (select d.a1k01 from acc1t0 c,acc1k0 d where c.a1t02=b.a1t02 and d.a1k01=c.a1t01 and nvl(d.a1k29,'N')<>'Y' and nvl(d.a1k12,0)=0 and d.a1k25 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0) and a1k25 is null and (a1k12 is null or a1k12 = 0) " & strSQL1 & StrSQL3 & " union " & _
             "select " & strField & "a1k02, a1k28 as FagentNo, a0y02 as DocDate, a1k01 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a0z04 as OAmount, 0 as FAmount, a1k04 as Yno, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu23 as fa17, cu76 as fa43, A0Y03 as Curr, a0y04 as Rate, CU10 as NATION,nvl(cu115,cu20) as EBox,cu148 as FA108,cu102 as fa70, a1k32, a1t01, a1t02,a1k33,a1k38,decode(cu115,null,'',cu200) as emailcc From acc0z0, acc0y0, acc1k0 a, customer , acc1t0 b where a1t01(+)=a1k01 and nvl(a1k32,'N')='C' and a1k01 = a0z02 and a0z01 = a0y01 and substr(a1k28, 1, 8) = cu01 and substr(a1k28, 9, 1) = cu02 and a1k01 in (select d.a1k01 from acc1t0 c,acc1k0 d where c.a1t02=b.a1t02 and d.a1k01=c.a1t01 and nvl(d.a1k29,'N')<>'Y' and nvl(d.a1k12,0)=0 and d.a1k25 is null)" & _
             " and nvl(a1k08, 0) <> nvl(a1k31, 0) and a1k25 is null and (a1k12 is null or a1k12 = 0) " & strSQL2 & StrSQL4 & " union " & _
             "select " & strField & "a1k02, a1k28 as FagentNo, a1202 as DocDate, a1201 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a1207 as OAmount, 0 as FAmount, a1k04 as Yno, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa17, fa43, a1204 as Curr, a1205 as Rate, FA10 as NATION,nvl(fa79,fa16) as EBox,FA108,fa70, a1k32, '', '',a1k33,a1k38,decode(fa105||fa79,null,'',fa134) as emailcc From acc120, acc0z0, acc1k0, fagent, acc1p0 where a1k01 = a0z02 and a0z01 = a1210 and substr(a1k28, 1, 8) = fa01 and substr(a1k28, 9, 1) = fa02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) and a1201=a1p23(+) and a1p04 is null " & strSQL1 & StrSQL3 & " union " & _
             "select " & strField & "a1k02, a1k28 as FagentNo, a1202 as DocDate, a1201 as DocNo, a1k13 as CaseNo1, a1k14 as CaseNo2, a1k15 as CaseNo3, a1k16 as CaseNo4, a1207 as OAmount, 0 as FAmount, a1k04 as Yno, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu23 as fa17, cu76 as fa43, a1204 as Curr, a1205 as Rate, CU10 as NATION,nvl(cu115,cu20) as EBox,cu148 as FA108,cu102 as fa70, a1k32, '', '',a1k33,a1k38,decode(cu115,null,'',cu200) as emailcc From acc120, acc0z0, acc1k0, customer, acc1p0 where a1k01 = a0z02 and a0z01 = a1210 and substr(a1k28, 1, 8) = cu01 and substr(a1k28, 9, 1) = cu02 and ((a1k29 <> 'Y' or a1k29 is null) and nvl(a1k08, 0) <> nvl(a1k31, 0)) and a1k25 is null and (a1k12 is null or a1k12 = 0) and a1201=a1p23(+) and a1p04 is null " & strSQL2 & StrSQL4
             
   'Add by Morgan 2006/11/28 加控制, NATION欄位要加''過濾,否則會很慢
   If Text3 = "020" Then
      bolChina = True
      StrSQLa = "select * from (" & StrSQLa & ") X where NATION||''='020'"
      'Added by Lydia 2024/10/24 代理人Y55822來信，對帳單須以全英文顯示；帳單將改建在Y55822020
      If Left(Text1, 6) = "Y55822" And Left(Text2, 6) = "Y55822" Then     'Y55822000的FA101=3不會產生催款單
         bolChina = False
      End If
      'end 2024/10/24
   Else
      bolChina = False
      StrSQLa = "select * from (" & StrSQLa & ") X where NATION||''<>'020' "
   End If
   
   'Added by Lydia 2016/12/22 傳入本所案號，指定催款單範圍(T收款寄證1728)
   If strCallCase <> "" Then
      StrSQLa = StrSQLa & " and CaseNo1||CaseNo2||CaseNo3||CaseNo4='" & strCallCase & "'"
   End If
   
   If Text3 <> "" Or Text4 <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label6 & Text3 & "-" & Text4
   End If
   If Text7 = "Y" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label12, 11) & Text7
   End If
   If Text8 = "Y" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label14, 7) & Text8
   End If
   If Text9 = "Y" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label15, 9) & Text9
   End If
   If Text6 = "Y" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label9, 7) & "PDF檔"
   End If
   
   If txtReceiver <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label11 & txtReceiver
   End If
   
   '2009/6/3 MODIFY BY SONIA 請款幣別不同也要跳頁
   'Modify by Amy 2017/08/15 紙本列印有下客戶編號且非Y27766則檔名改為 Y編號+X編號+幣別,故寫入暫存檔
   If bolShowCus = True And Left(Text1, 6) <> "Y27766" Then
        cnnConnection.Execute "Delete Accrpt2470 Where ID='" & strUserNum & "' "
        StrSQLa = "Insert Into Accrpt2470 (ID,R001,R002,R003,R004,R005,R006,R007,R008,R009,R010" & _
                        ",R011,R012,R013,R014,R015,R016,R017,R018,R019,R020,R021,R022,R023,R024,R025,R026,R027,R028,R029,R030" & _
                        ",R031,R032,R033,R034,R035,R036,R037,R038,R039,R041,R042,R043) " & StrSQLa
        cnnConnection.Execute StrSQLa
        Call UpdCusData
        'Modify by Amy 2017/07/31 若只下客戶編號迄號為XZZZZZ 則抓此代理人無客戶編號資料 ex:Y51817 1060630 TS-001452
        If Text10 = "" And Left(Text11, 6) = "XZZZZZ" Then
            cnnConnection.Execute "Delete Accrpt2470 Where ID='" & strUserNum & "' And R040 is not null"
        Else
            cnnConnection.Execute "Delete Accrpt2470 Where ID='" & strUserNum & "' And R040 is null"
        End If
        If bolChina = True Then
            StrSQLa = ",NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) as CusName "
        Else
            StrSQLa = ",NVL(CU05||CU88||CU89||CU90,Nvl(CU06,CU04)) as CusName "
        End If
        StrSQLa = "Select R001 as a1k02, R002 as FagentNo, R003 as DocDate, R004 as DocNo, R005 as CaseNo1, R006 as CaseNo2, R007 as CaseNo3, R008 as CaseNo4, " & _
                        "R009 as OAmount, R010 as FAmount, R011 as Yno, R012 as  fa05, R013 as fa63, R014 as fa64, R015 as fa65, R016 as fa32, R017 as fa18, R018 as fa33,R019 as fa19, R020 as fa34, " & _
                        "R021 as fa20, R022 as fa35, R023 as fa21, R024 as fa36, R025 as fa22, R026 as fa06, R027 as fa23, R028 as fa04, R029 as fa17, R030 as fa43, " & _
                        "R031 as Curr, R032 as Rate, R033 as Nation,R034 as EBox,R035 as FA108,R036 as fa70, R037 as a1k32, R038 as a1t01, R039 as a1t02, " & _
                        "R040 as CusNo,R041 as a1k33,R042 as a1k38, R043 as emailcc " & StrSQLa & _
                        "From Accrpt2470,Customer Where ID='" & strUserNum & "' And Substr(R040,4,8)=cu01(+) And Substr(R040,12,1)=cu02(+) " & _
                        "Order by R002 asc,CusNo asc, R031 asc, R003 asc, R004 asc"
   Else
        StrSQLa = StrSQLa & " order by FagentNo asc, Curr asc, DocDate asc, DocNo asc"
   End If
   
   adoacc1k0.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc1k0.RecordCount = 0 Then
      InsertQueryLog (0)
      strCon10 = MsgText(602)
      adoacc1k0.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   
   '產生csv
   If Check1.Value = 1 Then
      If CSVSave(strXLSFile) = True Then
         If MsgBox("CSV檔案已產生！" & vbCrLf & vbCrLf & strXLSFile & vbCrLf & vbCrLf & "是否開啟資料夾？", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
            ShellExecute hLocalFile, "explore", strPrtPath, vbNullString, vbNullString, 1
         End If
      End If
      If adoacc1k0.State <> adStateClosed Then adoacc1k0.Close
      Exit Sub
   '產生Excel
   ElseIf bolExcel = True Then
      ExcelSaveNew strXLSFile
      If MsgBox("EXCEL檔案已產生！" & vbCrLf & vbCrLf & strXLSFile & vbCrLf & vbCrLf & "是否開啟資料夾？", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
         ShellExecute hLocalFile, "explore", strPrtPath, vbNullString, vbNullString, 1
      End If
      adoacc1k0.Close
      Exit Sub
   End If
  
   intLength = 0
   douAmount = 0
   douTAmount = 0
   douRAmount = 0
   douOverDue1 = 0
   douOverDue2 = 0
   douOverDue3 = 0
   strNo = ""
   m_DNCurr = ""
   m_iDocCount = 0
   m_iPrintCount = 0
   m_iMailCount = 0
   m_FNo = ""
   
   '非單筆時另外印催款明細(大陸或有指定收件者時也不印)
   If Text3 <> "020" And (Mid(Text1, 1, 6) <> Mid(Text2, 1, 6) Or Text1 & Text2 = "") And txtReceiver = "" Then
      bolReport = True
      adoTaie.Execute "Delete From ACCRPT207 Where R20701='" & strUserNum & "'"
   Else
      bolReport = False
   End If
   'Added by Lydia 2024/10/24 代理人Y55822來信，對帳單須以全英文顯示；帳單將改建在Y55822020
   If Text3 = "020" Then
      If Left("" & adoacc1k0.Fields("FAGENTNO"), 6) = "Y55822" Then
         bolChina = False
      Else
         bolChina = True
      End If
   End If
   'end 2024/10/24
   
   'Added by Lydia 2025/03/11
   If Text9 = "" Then
      If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M31" Then
         txtSend.Visible = True: lblSend.Visible = True
      End If
      txtSend.Text = ""
   End If
   'end 2025/03/11
   m_FrmName = Me.Name & "-" & Format(ServerTime, "000000") 'Move by Lydia 2025/05/06 從txtSend.Text = ""下方移過來
   
   '產生電子檔---路徑
   If Text6 = "Y" Then
        '國外部承辦組
      Select Case Left(Pub_StrUserSt03, 2)
         Case "F1" 'FCT
            strSavePath = PUB_GetEFilePath("FCT") & "\Account"
         Case "F2" 'FCP
            strSavePath = PUB_GetEFilePath("FCP") & "\Account"
         Case "F3"
            strSavePath = PUB_GetEFilePath("FCL") & "\Account"
         Case Else
            strSavePath = PUB_Getdesktop
      End Select
      
      '傳入本所案號，指定催款單範圍(T收款寄證1728)
      If m_SavePath <> "" Then
          strSavePath = m_SavePath
      End If

      If bolCallMail = False Then
          If Dir(strSavePath, vbDirectory) = "" Then
             MkDir strSavePath
          End If
      Else
          strSavePath = strPrtPath
      End If
   Else
      strSavePath = strPrtPath
   End If
   strDefDir = strSavePath
   
   Erase strMailFailList
   ReDim strMailFailList(0)
   strPicLetter = ""
   strPicFileNames = ""
   iPageNo = 0
   
   strBatchNoList = ""
   strBatchNoRecList = ""
   With adoacc1k0
      .MoveFirst
      Do While Not .EOF
         'Added by Lydia 2024/10/24 代理人Y55822來信，對帳單須以全英文顯示；帳單將改建在Y55822020
         If Text3 = "020" Then
            If Left("" & .Fields("FAGENTNO"), 6) = "Y55822" Then
               bolChina = False
            Else
               bolChina = True
            End If
         End If
         'end 2024/10/24
         '若有下國籍條件
         If Text3 <> "" Then
            If .Fields("NATION").Value < Text3 Then
               GoTo NextSkip
            End If
         End If
         If Text4 <> "" Then
            If .Fields("NATION").Value > Text4 & "z" Then
               GoTo NextSkip
            End If
         End If
         '若是舊系統資料
         If Len(.Fields("DocNo").Value) = 10 And strDocNo = Mid(.Fields("DocNo").Value, 1, 8) And .Fields("OAmount").Value = 0 Then
            GoTo NextSkip
         ElseIf Len(.Fields("DocNo").Value) = 10 Then
            strDocNo = Mid(.Fields("DocNo").Value, 1, 8)
         Else
            strDocNo = .Fields("DocNo").Value
         End If
         
         '整批列印請款單
         bolIsBatchInvoice = False
         If Not IsNull(.Fields("a1t01")) Then
            If .Fields("OAmount").Value > 0 Then
               If InStr(strBatchNoRecList, .Fields("DocDate") & .Fields("a1t02")) = 0 Then
                  bolIsBatchInvoice = True
                  strBatchNoRecList = strBatchNoRecList & "," & .Fields("DocDate") & .Fields("a1t02")
               End If
            Else
               If InStr(strBatchNoList, .Fields("a1t02")) = 0 Then
                  bolIsBatchInvoice = True
                  strBatchNoList = strBatchNoList & "," & .Fields("a1t02")
               End If
            End If
            
            If bolIsBatchInvoice Then
               strExc(0) = "select min(a1t02)||'-'||substr(max(a1t01),-3) from acc1t0 where a1t02='" & .Fields("a1t02") & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strBatchInvoiceNo = "" & RsTemp(0)
               End If
            Else
               GoTo NextSkip
            End If
         End If
         
         '代理人不同
         '2009/6/3 MODIFY BY SONIA 請款幣別不同也要跳頁
         'Modified by Lydia 2017/02/18 T收款寄證是以個案命名
         'Modify by Amy 2017/08/15 紙本有下客戶編號改檔名-婉莘
         'Memo by Lydia 2019/06/10 從收款作業Frmacc2110來的催款作業,如果遇到同一代理人不同幣別會產生2封催款Email,只有附件不同 (ex.M10802625, M10802626)
         If bolShowCus = True Then strNowCus = Mid(.Fields("CusNo"), 4)
         
         If (strCallCase = "" And ((bolShowCus = False And strNo <> .Fields("FagentNo").Value & .Fields("Curr").Value) Or (bolShowCus = True And strNo <> .Fields("FagentNo").Value & "(vs " & strNowCus & ")" & .Fields("Curr").Value))) _
              Or (strCallCase <> "" And bolShowCus = False And strNo <> strCallCase & .Fields("Curr").Value) Then
               '不論合計是否為0 都要清空否則會殘留到後面,另Email時加控制有欠款才寄 Ex.106/12/4 Y54511,Y54518
               If bolOpenXls = True Then
                  If PrintExcel_Footer = True Then
                  End If
                  If bol2Pdf = True Or bolEmail = True Then
                     '先存PDF檔(另存新檔)放在桌面，不關EXCEL後面再處理信頭、信尾>>PrintExcel_BFile
                     If PUB_PrintExcel2File(xlsRpt, strSavePath, strNowDoc, strExc(1), False) = True Then
                        strPicFileNames = strSavePath & "\" & strExc(1)
                        If bolEmail And douAmount <> 0 Then
                           Call ProcSendMail(strPicFileNames)
                        End If
                     End If
                  Else
                     If bol2Printer = True Then
                        WksRpt1.PrintOut Copies:=1, Collate:=True '列印
                        'Added by Lydia 2025/03/11 SEQ=2>>紙本
                        strSql = "insert into rdatafactory (formname,id,seqno,rowseq,r001) values ('" & m_FrmName & "','" & strUserNum & "','2'," & m_iPrintCount & ",'" & m_FNo & "') "
                        cnnConnection.Execute strSql
                        'end 2025/03/11
                     End If
                  End If
                  If bolEmail And douAmount <> 0 Then
                  End If
               End If
                     
               douOverDue1 = 0
               douOverDue2 = 0
               douOverDue3 = 0
               douAmount = 0
               douTAmount = 0
               douRAmount = 0
               iPageNo = 0
               '電子信箱
               strEMailBox = "" & adoacc1k0("EBox")
               strEmailCC = "" & adoacc1k0("emailcc") '財務副本信箱
               strCurr = adoacc1k0.Fields("Curr").Value 'Move by Lydia 2025/01/16 催款幣別;從下方移上來
               
               '產生電子檔
               If Text6 = "Y" Then
                  bol2Printer = False
                  bol2Pdf = True
                  bolEmail = False
                  If bolCallMail = False Then
                      '指定催款單範圍(T收款寄證1728)，不建子目錄
                      If strCallCase = "" Then strSavePath = strDefDir & "\" & .Fields("FagentNo")
                      
                      If Dir(strSavePath, vbDirectory) = "" Then
                         MkDir strSavePath
                      End If
                  End If
               Else
                   If strEMailBox <> "" And UCase(strEMailBox) <> "NO" And Text9 = "" Then
                     bol2Pdf = True
                     bol2Printer = False
                     bolEmail = True
                  Else
                     bol2Pdf = False
                     bol2Printer = True
                     bolEmail = False
                  End If
               End If
            
               '國外部承辦或有指定收件者
               If bolPromoter Or txtReceiver <> "" Then
                  strReceiver = txtReceiver
                  strEmailCC = ""
                  bol2Pdf = True
                  bol2Printer = False
                  bolEmail = True
               Else
                  strReceiver = strEMailBox
               End If
            
               'Add by Morgan 2009/7/6
               '重印紙本控制
               If Text8 = "Y" Then
                  bolEmail = False
                  bolReport = False
                  bol2Pdf = False
               End If

               '是否存電子檔
               If bol2Pdf = True Then
                  m_iNo = 0: m_iNo2 = 0
                  If strSrvDate(1) >= 智慧所更名日 Then
                     'M改用EXCEL：信頭、信尾分開來
                     PUB_GetLetterPicID "2", , m_iNo, m_iNo2, , , "HALF"
                  Else
                     m_iNo = 6
                  End If
               End If
               If bol2Printer = True Or bol2Pdf = True Then
                  If PrintExcel_BFile(IIf(bolOpenXls = False, True, False), m_iNo, m_iNo2) = True Then
                     bolOpenXls = True
                  End If
               End If
                  
               'Added by Morgan 2024/9/6
               If strCallCase <> "" Then '指定催款單範圍(T收款寄證1728),以案號為檔名
                  strNowDoc = strCallCase & .Fields("Curr").Value & ".PDF"
                  
               ElseIf bolShowCus = True Then  '有下客戶編號改檔名
                  strNowDoc = .Fields("FagentNo").Value & "(vs " & strOldCus & ")" & .Fields("Curr").Value & ".PDF"
               Else     '預設Y編號+幣別
                  strNowDoc = .Fields("FagentNo").Value & .Fields("Curr").Value & ".PDF"
               End If
            
               'strCurr = adoacc1k0.Fields("Curr").Value 'Mark by Lydia 2025/01/16 移到上方
      
               '2009/6/4 MODIFY BY SONIA 請款幣別不同也要跳頁
               'Modified by Lydia 2016/12/22 指定催款單範圍(T收款寄證1728),以案號為檔名
               If strCallCase <> "" Then
                   strNo = strCallCase & .Fields("Curr").Value
               'Add by Amy 2017/08/15 有下客戶編號改檔名
               ElseIf bolShowCus = True Then
                  strOldCus = Mid("" & .Fields("CusNo"), 4)
                  If strOldCus = "" Then strOldCus = "其他"
                  strNo = .Fields("FagentNo").Value & "(vs " & strOldCus & ")" & .Fields("Curr").Value
               Else
                   strNo = .Fields("FagentNo").Value & .Fields("Curr").Value
               End If
               m_DNCurr = adoacc1k0.Fields("Curr").Value
               m_FNo = .Fields("FagentNo").Value
               intCounter = 0
         End If  'If (strCallCase = "" And ((bolShowCus = False And strNo <> .Fields("FagentNo").Value & .Fields("Curr").Value) Or
         
         PrintExcel_Row
         
         '產生請款單的電子檔
         If Text7 = "Y" Then
            '舊系統的由人工掃描到原來請款單的存放路徑
            If Len(.Fields("DocNo").Value) = 10 Then
               strFileName = "DN" & Mid(adoacc1k0.Fields("DocNo").Value, 3, 6) & ".pdf"
               strSource = PUB_GetEFilePath(.Fields("CaseNo1")) & "\" & .Fields("CaseNo1") & .Fields("CaseNo2") & IIf(.Fields("CaseNo3") & .Fields("CaseNo4") <> "000", .Fields("CaseNo3") & .Fields("CaseNo4"), "") & "\" & strFileName
               strDestination = strSavePath & "\" & strFileName
               If Dir(strSource) <> "" Then
                  FileCopy strSource, strDestination
               End If
            Else
               'Modified by Lydia 2016/09/08 改到催款單列印完後
               strA1k01List = strA1k01List & strDocNo & ","
               'Added by Lydia 2017/03/02 請款單存放的資料夾路徑
               strA1k01Dir = strA1k01Dir & strDefDir & "\" & .Fields("FagentNo").Value & ","
            End If
         End If
         
NextSkip:
         
         If bolReport = True Then
            '請款日期小於請款日期止日-1年的才要印
            If Val(.Fields("DocDate")) < Val(Replace(Me.MaskEdBox2.Text, "/", "")) - 10000 Then
            
               If InStr("'FCP','FG'", "'" & .Fields("CaseNo1") & "'") > 0 Then
                  strExc(0) = "FCP"
               ElseIf InStr("'FCT','CFT','CFC','S','L'", "'" & .Fields("CaseNo1") & "'") > 0 Then
                  strExc(0) = "FCT"
               ElseIf InStr("'FCL','LIN'", "'" & .Fields("CaseNo1") & "'") > 0 Then
                  strExc(0) = "FCL"
               ElseIf InStr("'P','PS','CFP','CPS'", "'" & .Fields("CaseNo1") & "'") > 0 Then
                  strExc(0) = "P"
               ElseIf Left("" & .Fields("CaseNo1"), 1) = "T" Then
                  strExc(0) = "T"
               Else
                  strExc(0) = ""
               End If
               If .Fields("FAmount").Value > 0 Then
                  adoTaie.Execute "insert into ACCRPT207(R20701,R20702,R20703,R20704,R20705,R20706,R20707)" & _
                     " values ('" & strUserNum & "','" & .Fields("DocNo").Value & "','" & .Fields("CaseNo1") & "','" & .Fields("CaseNo2") & "','" & .Fields("CaseNo3") & "','" & .Fields("CaseNo4") & "','" & strExc(0) & "')"
               End If
            End If
         End If
         .MoveNext
      Loop
   End With  'With adoacc1k0
   
   If bolOpenXls = True Then
      If PrintExcel_Footer = True Then
      End If
      If bol2Pdf = True Or bolEmail = True Then
         '先存PDF檔(另存新檔)放在桌面，不關EXCEL後面再處理信頭、信尾>>PrintExcel_BFile
         If PUB_PrintExcel2File(xlsRpt, strSavePath, strNowDoc, strExc(1), False) = True Then
            strPicFileNames = strSavePath & "\" & strExc(1)
            If bolEmail And douAmount <> 0 Then
               Call ProcSendMail(strPicFileNames)
            End If
         End If
      End If
      
      xlsRpt.Workbooks(1).Save
      If bol2Printer = True Then
         WksRpt1.PrintOut Copies:=1, Collate:=True '列印
         'Added by Lydia 2025/03/11 SEQ=2>>紙本
         strSql = "insert into rdatafactory (formname,id,seqno,rowseq,r001) values ('" & m_FrmName & "','" & strUserNum & "','2'," & m_iPrintCount & ",'" & m_FNo & "') "
         cnnConnection.Execute strSql
         'end 2025/03/11
      End If
      xlsRpt.Workbooks.Close
      xlsRpt.Quit
      Set xlsRpt = Nothing
      Set WksRpt1 = Nothing
   End If

   If Text6 = "Y" Then
      'Modified by Lydia 2016/12/22 指定催款單範圍(T收款寄證1728)
      If strCallCase <> "" Then
            strMsg = strMsg & "電子檔已存於" & strDefDir & " ！" & vbCrLf
            m_iDocCount = 0 '不詢問是否列印
      ElseIf Left(Pub_StrUserSt03, 1) = "F" Then
         If strDefDir <> "" Then
            strMsg = strMsg & "電子檔已存於" & strDefDir & " ！" & vbCrLf
         Else
            strMsg = strMsg & "電子檔已存於相關路徑！" & vbCrLf
         End If
      Else
         strMsg = strMsg & "電子檔已存桌面！" & vbCrLf
      End If
   End If

   '刪除舊的暫存檔
   Call PUB_KillTempFile(strUserNum & "\$*.*")
   adoacc1k0.Close
   
   '列印清單
   If m_iDocCount > 0 Then
      InsertQueryLog (m_iDocCount)
      strExc(0) = ""
      strExc(0) = strExc(0) & "本次共產生 " & m_iDocCount & " 份催款資料" & vbCrLf
      strExc(0) = strExc(0) & "紙本 " & m_iPrintCount & " 份" & vbCrLf
      strExc(0) = strExc(0) & "Email " & m_iMailCount & " 份" & vbCrLf
      
      If strMailFailList(0) <> "" Then
         strExc(0) = strExc(0) & "Email失敗 " & UBound(strMailFailList) + 1 & " 份，清單如下：" & vbCrLf & vbCrLf
         For intI = 0 To UBound(strMailFailList)
            strExc(0) = strExc(0) & strMailFailList(intI) & vbCrLf
         Next
      End If
      'Added by Lydia 2020/09/10 請款單：判斷PDF檔案是否存在
      If m_strErr2480 <> "" Then
         strExc(0) = strExc(0) & vbCrLf & "請款單電子檔產生失敗：" & vbCrLf & Replace(m_strErr2480, "＆", vbCrLf) & vbCrLf
      End If

         strExc(1) = "催款日期：" & strSrvDate(1) & vbCrLf & vbCrLf
         strExc(1) = strExc(1) & "請款日期：" & MaskEdBox1 & " ～ " & MaskEdBox2 & vbCrLf
         If Text3 & Text4 <> "" Then
            strExc(1) = strExc(1) & "國籍：" & Text3 & " ～ " & Text4 & vbCrLf
         End If
         strExc(0) = strExc(1) & vbCrLf & strExc(0)
      'Added by Lydia 2020/10/08 單筆(代理人前6碼)的催款單不用寄MAIL給操作人員, 除非有錯誤 !
      'Modofied by Morgan 2025/7/15 非單筆時都要MAIL--斯閔
      'If strMailFailList(0) <> "" Or m_strErr2480 <> "" Or (Left(Text1, 6) <> Left(Text2, 6)) Or (Left(Text10, 6) <> Left(Text11, 6)) Then
      If strMailFailList(0) <> "" Or m_strErr2480 <> "" Or (Left(Text1, 6) <> Left(Text2, 6)) Or (Left(Text10, 6) <> Left(Text11, 6)) Or (Text2 = "" And Text11 = "") Then
           PUB_SendMail strUserNum, strUserNum, "", Me.Caption & "-" & "完成作業" & IIf(InStr(strExc(0), "失敗") > 0, "，有發生失敗請參考內文！", ""), strExc(0)
      End If
      If strMailFailList(0) <> "" Or m_strErr2480 <> "" Then
          MsgBox strExc(0), vbInformation, "發生失敗"
      End If

   End If
   
   If bolReport = True Then
      '更新智權人員
      'FCP,P--以請款對象國籍抓國家檔的FCP承辦智權人員
      adoTaie.Execute "update accrpt207 set R20708=(select max(na51) from acc1k0,fagent,customer,nation" & _
         " where a1k01=R20702 and fa01(+)=substr(a1k28,1,8) and fa02(+)=substr(a1k28,9)" & _
         " and cu01(+)=substr(a1k28,1,8) and cu02(+)=substr(a1k28,9) and na01=decode(fa01,null,cu10,fa10))" & _
         " where R20701='" & strUserNum & "' and R20707 in ('FCP','P')"
      'FCT,T--抓該案號最後接洽單之智權人員(若已離職則改放虛建智權人員)
      adoTaie.Execute "update accrpt207 set R20708=(select substr(max(cp09||cp13),10) from caseprogress" & _
         " where cp01=R20703 and cp02=R20704 and cp03=R20705 and cp04=R20706 and cp09<'B')" & _
         " where R20701='" & strUserNum & "' and R20707 in ('FCT','T')"
         
      adoTaie.Execute "update accrpt207 set R20708='F4103'" & _
         " where R20701='" & strUserNum & "' and R20707 in ('FCT','T')" & _
         " and not exists(select * from staff where st01=R20708 and st04='1')"
      
      'FCL--抓請款單之智權人員(最後收文智權人員)
      adoTaie.Execute "update accrpt207 set R20708=(select substr(max(cp05||cp13),9) from caseprogress" & _
         " where cp60=R20702)" & _
         " where R20701='" & strUserNum & "' and R20707='FCL'"
 
      strExc(0) = "select R20707 類別" & _
         ",substrb(ST02,1,6) 智權人員" & _
         ",substrb(na03,1,8) 國籍" & _
         ",substrb(a1k28||' '||rtrim(fa05||' '||fa63||' '||fa64||' '||fa65)||rtrim(cu05||' '||cu88||' '||cu89||' '||cu90),1,30) 請款對象" & _
         ",a1k13||'-'||a1k14||decode(a1k15||a1k16,'000','','-'||a1k15||'-'||a1k16) 本所案號" & _
         ",a1k01 請款單號" & _
         ",substrb(sqldatew(a1k02+19110000),1,10) 請款日期" & _
         ",decode(nvl(a1k30,0),0,(a1k08 - nvl(a1k31, 0)),decode(nvl(a1k10,0),0,0,(a1k11-a1k30)/a1k10)) 應收外幣" & _
         ",decode(nvl(a1k30,0),0,nvl(a1k09,0),sign(nvl(a1k09,0)-a1k30),1,nvl(a1k09,0)-a1k30,0) 應收規費" & _
         ",(nvl(a1k11,0)-nvl(a1k09,0))/1000 點數" & _
         " From (select a.*,b.*,c.*,d.*,e.*,decode(fa01,null,cu10,fa10) X1 from accrpt207 a, staff b, acc1k0 c, fagent d,customer e" & _
         " where st01(+)=R20708 and a1k01(+)=R20702 And R20701='" & strUserNum & "' " & _
         " and fa01(+)=substr(a1k28,1,8) and fa02(+)=substr(a1k28,9)" & _
         " and cu01(+)=substr(a1k28,1,8) and cu02(+)=substr(a1k28,9)" & _
         ") X, Nation where na01(+)=X1" & _
         " order by 1,2,3,4,5,7,6"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         .MoveFirst
         strKind = ""
         strSubject = Format(strSrvDate(1), "####/##/##") & "催款明細"
         Erase strFileList
         strDate = IIf(Me.MaskEdBox1.Text = "___/__/__", "", Format(ChangeTStringToWDateString(Replace(Me.MaskEdBox1.Text, "/", "")), "YYYY/MM/DD")) & _
               IIf(Me.MaskEdBox1.Text = "___/__/__" And Me.MaskEdBox2.Text = "___/__/__", "", " ～ ") & _
               IIf(Me.MaskEdBox2.Text = "___/__/__", "", Format(ChangeTStringToWDateString(Replace(Me.MaskEdBox2.Text, "/", "")), "YYYY/MM/DD"))
         Do While Not .EOF
            If .Fields(0) <> strKind Then
               If strKind <> "" And ff > 0 Then
                  '合計
                  strContent = Space(90) & convForm(Format(dblTotUsAmount, "#,##0"), 12, , True) '代理人應收總額
                  strContent = strContent & " " & convForm(Format(dblTotFeeAmount, "#,##0"), 10, , True) '應收總規費
                  strContent = strContent & " " & convForm(Format(dblTotPt, "#,##0.000"), 8, , True)  '總點數
                  
                  Print #ff, "------ -------- -------------------------- --------------- ---------- ---------- -------- ------------ ---------- --------"
                  Print #ff, strContent
               End If
               
               If ff > 0 Then Close #ff
               ff = FreeFile
               
               strKind = "" & .Fields(0)
               strKey1 = ""
               StrKey2 = ""
               strKey3 = ""
               dblTotUsAmount = 0
               dblTotFeeAmount = 0
               dblTotPt = 0
               
               strFileName = App.path & "\催款明細(" & strKind & ").txt"
               
               Open strFileName For Output As ff
                              
               strExc(1) = Format(DBDATE(MaskEdBox2.Text) - 10000, "####/##/##")
               Print #ff, ""
               Print #ff, "※本報表只列出請款日期早於 " & strExc(1) & " 的資料。"
               Print #ff, ""
               Print #ff, "                                           " & strSubject & "(" & strKind & ")"
               Print #ff, ""
               Print #ff, "                                           請款日期：" & strDate
               Print #ff, "                                                                                          代理人(美金)     (台幣)   (台幣)"
               Print #ff, "智權人 國籍     請款對象                   本所案號        請款單號   請款日期   應收美金 應收總額     應收總規費   總點數"
               Print #ff, "------ -------- -------------------------- --------------- ---------- ---------- -------- ------------ ---------- --------"
               
               
               strFileList(0) = strFileList(0) & strFileName & ";"
               Select Case strKind
                  Case "FCP", "P"
                     strFileList(1) = strFileList(1) & strFileName & ";"
                  Case "FCT", "T"
                     strFileList(2) = strFileList(2) & strFileName & ";"
                 Case "FCL"
                     strFileList(3) = strFileList(3) & strFileName & ";"
               End Select
                  
            End If
            strContent = Empty
            '智權人員
            If strKey1 = "" & .Fields(1) Then
               strContent = strContent & Space(6)
            Else
               strContent = strContent & convForm("" & .Fields(1), 6) '智權人員
            End If
            '國籍
            If strKey1 & StrKey2 = "" & .Fields(1) & .Fields(2) Then
               strContent = strContent & Space(9)
            Else
               strContent = strContent & " " & convForm("" & .Fields(2), 8) '國籍
            End If
            '請款對象
            If strKey3 = "" & .Fields(3) Then
               strContent = strContent & Space(27)
            Else
               strContent = strContent & " " & convForm("" & .Fields(3), 26) '請款對象
            End If
            strContent = strContent & " " & convForm("" & .Fields(4), 15) '本所案號
            strContent = strContent & " " & convForm("" & .Fields(5), 10) '請款單號
            strContent = strContent & " " & convForm("" & .Fields(6), 10) '請款日期
            strContent = strContent & " " & convForm(Format("" & .Fields(7), "#,##0"), 8, , True) '應收美金
            dblUsAmount = dblUsAmount + Val("" & .Fields(7))
            dblTotUsAmount = dblTotUsAmount + Val("" & .Fields(7))
            dblFeeAmount = dblFeeAmount + Val("" & .Fields(8))
            dblTotFeeAmount = dblTotFeeAmount + Val("" & .Fields(8))
            dblPt = dblPt + Val("" & .Fields(9))
            dblTotPt = dblTotPt + Val("" & .Fields(9))
            strKey1 = "" & .Fields(1)
            StrKey2 = "" & .Fields(2)
            strKey3 = "" & .Fields(3)
            .MoveNext
            bolSum = False
            If .EOF Then
               bolSum = True
            '代理人編號不同
            ElseIf strKind & strKey1 & StrKey2 & strKey3 <> "" & .Fields(0) & .Fields(1) & .Fields(2) & .Fields(3) Then
               bolSum = True
            End If
            If bolSum = True Then
               strContent = strContent & " " & convForm(Format(dblUsAmount, "#,##0"), 12, , True) '代理人應收總額
               strContent = strContent & " " & convForm(Format(dblFeeAmount, "#,##0"), 10, , True) '應收總規費
               strContent = strContent & " " & convForm(Format(dblPt, "#,##0.000"), 8, , True)  '總點數
               dblUsAmount = 0
               dblFeeAmount = 0
               dblPt = 0
            End If
            Print #ff, strContent
         Loop
         '合計
         strContent = Space(90) & convForm(Format(dblTotUsAmount, "#,##0"), 12, , True) '代理人應收總額
         strContent = strContent & " " & convForm(Format(dblTotFeeAmount, "#,##0"), 10, , True) '應收總規費
         strContent = strContent & " " & convForm(Format(dblTotPt, "#,##0.000"), 8, , True)  '總點數
         
         Print #ff, "------ -------- -------------------------- --------------- ---------- ---------- -------- ------------ ---------- --------"
         Print #ff, strContent
         If ff > 0 Then Close #ff
         'FCP,P
         If strFileList(1) <> "" Then
            'Modify by Morgan 2011/2/25 85030 留職停薪半年,暫改 88003
            'Modify by Amy 2022/11/24 +if 88003(王文安協理)退休暫改77015(顏裕洋副理)99037(簡偉倫經理)
            If strSrvDate(1) >= 20221130 Then
                SendReport "77015;99037", strSubject, strFileList(1)
            Else
                SendReport "88003", strSubject, strFileList(1)
            End If
         End If
         'FCT,T
         If strFileList(2) <> "" Then
            SendReport "78011;80030", strSubject, strFileList(2)
         End If
         'FCL
         If strFileList(3) <> "" Then
            SendReport "99015", strSubject, strFileList(3)
         End If
         'ALL
         If strFileList(0) <> "" Then
            SendReport "81040", strSubject, strFileList(0)
            arrFile = Split(strFileList(0), ";")
            For intI = LBound(arrFile) To UBound(arrFile)
               If arrFile(intI) <> "" Then
                  Kill arrFile(intI)
               End If
            Next
         End If
         End With
      End If
   End If
   
   strMsg = strMsg & "作業結束！"
   If bolCallMail = False Then MsgBox strMsg, vbInformation
End Sub

'Adeed by Lydia 2024/12/31 改用EXCEL：催款單換行
Private Sub PrintExcel_BPage(Optional ByVal pAddLine As Integer = 1)
   nRow = nRow + pAddLine
   If nRow > xRowE Then
      Call PrintExcel_BFile(False, m_iNo, m_iNo2)
   End If
End Sub

'Added by Lydia 2024/12/31 改用EXCEL：催款單明細
Private Sub PrintExcel_Row()
Dim douDollar As Double, douODollar As Double
Dim bolTQM As Boolean
   
   '大陸格式PrintRow1A4, 國外代理人PrintRowA4
   bolTQM = False
   
   intCounter = 0 '列印欄位Column
   '帳款日期：Date
   strData = ""
   If Not IsNull(adoacc1k0.Fields("DocDate").Value) Then
      strData = Format(AFDate(CADate(adoacc1k0.Fields("DocDate").Value)), "MM/DD/YY")
   End If
   If strData <> "" Then
      WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = strData
   End If
   intCounter = intCounter + 1
   
   '帳單編號：Debit No.
   strData = ""
   If Not IsNull(adoacc1k0.Fields("DocNo").Value) Then
      If Len(adoacc1k0.Fields("DocNo").Value) = 10 Then
         strData = Mid(adoacc1k0.Fields("DocNo").Value, 3, 6)
      Else
         strData = adoacc1k0.Fields("DocNo").Value
      End If
   End If
   If bolIsBatchInvoice Then strData = strBatchInvoiceNo 'Added by Morgan 2016/2/15 整批列印請款單
   strDBno = strDBno & strData & ", "
   If strData <> "" Then
      WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = strData
   End If
   intCounter = intCounter + 1
   
   '我方文號：O/Ref. No.
   strData = ""
   If Not IsNull(adoacc1k0.Fields("CaseNo1").Value) Then
      If adoacc1k0.Fields("CaseNo3").Value = "0" And adoacc1k0.Fields("CaseNo4").Value = "00" Then
         strData = adoacc1k0.Fields("CaseNo1").Value & "-" & adoacc1k0.Fields("CaseNo2").Value
      Else
         strData = adoacc1k0.Fields("CaseNo1").Value & "-" & adoacc1k0.Fields("CaseNo2").Value & "-" & adoacc1k0.Fields("CaseNo3").Value
      End If
      'Add by Amy 2024/06/18 商標查名請款單由TS或S 案轉入者,在O/REF.NO.之本所案號後再加入"(查名本所案號)" ex:FCT-051898(S-008316)
      If InStr("" & adoacc1k0.Fields("CaseNo1"), "T") > 0 And "" & adoacc1k0.Fields("CaseNo1") <> "TS" Then
         strData = strData & GetTMQCaseNo(False, "" & adoacc1k0.Fields("DocNo"), bolTQM)
      End If
   End If
   'Added by Morgan 2016/2/16 整批列印請款單:本所抓第一個請款單號的本所號+",..."
   If bolIsBatchInvoice Then
      strExc(0) = "select a1k13||'-'||a1k14||decode(a1k15||a1k16,'000','','-'||a1k15||'-'||a1k16) from acc1k0 where a1k01='" & adoacc1k0.Fields("a1t02") & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strData = RsTemp(0) & ",..."
      End If
   End If
   'end 2016/2/16
   If strData <> "" Then
      strData = Trim(Left(strData, 15))
      WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = strData
      'If Len(strData) > 12 Then
         WksRpt1.Range(Chr(xCols + intCounter) & nRow).Font.Size = 10
      'End If
   End If
   intCounter = intCounter + 1
   
   '貴方文號：Y/Ref. No.
   strData = ""
   strExc(1) = adoacc1k0.Fields("DocNo").Value
   If bolChina = True Then '大陸格式
      strData = GetYourRefNo1(adoacc1k0.Fields("CaseNo1"), adoacc1k0.Fields("CaseNo2"), adoacc1k0.Fields("CaseNo3"), adoacc1k0.Fields("CaseNo4"), IIf(Left(adoacc1k0.Fields("FagentNo"), 6) = "Y52269", True, False))
   Else
      '國外代理人-英文
      adoquery.CursorLocation = adUseClient
      'Modified by Morgan 2016/2/16 配合整批列印請款單改語法
      If bolIsBatchInvoice Then strExc(1) = adoacc1k0.Fields("a1t02").Value
      adoquery.Open "select pa77 as Yno from patent where (pa01,pa02,pa03,pa04) in (select a1k13,a1k14,a1k15,a1k16 from acc1k0 where a1k01='" & strExc(1) & "') union " & _
                    "select tm45 as Yno from trademark where (tm01,tm02,tm03,tm04) in (select a1k13,a1k14,a1k15,a1k16 from acc1k0 where a1k01='" & strExc(1) & "') union " & _
                    "select lc23 as Yno from lawcase where (lc01,lc02,lc03,lc04) in (select a1k13,a1k14,a1k15,a1k16 from acc1k0 where a1k01='" & strExc(1) & "') union " & _
                    "select sp27 as Yno from servicepractice where (sp01,sp02,sp03,sp04) in (select a1k13,a1k14,a1k15,a1k16 from acc1k0 where a1k01='" & strExc(1) & "')", adoTaie, adOpenStatic, adLockReadOnly
      'end 2016/2/16
      If adoquery.RecordCount <> 0 Then
         If IsNull(adoquery.Fields("Yno").Value) = False Then
            strData = adoquery.Fields("Yno").Value
         End If
      End If
      adoquery.Close
   End If
   'Add by Amy 2024/06/14 商標查名請款單由TS或S 案轉入者,Y/REF.NO.有值不秀,避免看不出案號 ex:T-245686(TS-001979) -秀玲
   If bolTQM = True Then strData = ""
      
   'Added by Morgan 2014/2/17 改比照請款單抓法
   If adoacc1k0.Fields("CaseNo1") = "FCP" Then
      strExc(0) = "Select PA106 From CaseProgress,Patent Where PA01(+)=CP01 And PA02(+)=CP02 And PA03(+)=CP03 And PA04(+)=CP04 And CP60='" & strExc(1) & "' And CP10='605' And CP01='FCP' and pa76 is not null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'Add By Sindy 2016/7/14 +檢查商標延展
   ElseIf InStr("T,FCT,CFT,TF", adoacc1k0.Fields("CaseNo1")) > 0 Then
      strExc(0) = "Select TM65 From CaseProgress,Trademark Where TM01(+)=CP01 And TM02(+)=CP02 And TM03(+)=CP03 And TM04(+)=CP04 And CP60='" & strExc(1) & "' And CP10='102' And CP01 in('T','FCT','CFT','TF') and TM33 is not null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   Else
      intI = 0
   End If
   '有年費/延展彼所案號
   If intI = 1 Then
      If PUB_GetFCCaseNo(strExc(1), strExc(2), True) = True Then
         strData = strExc(2)
      Else
         strData = "" & RsTemp(0)
      End If
   ElseIf PUB_GetFCCaseNo(strExc(1), strExc(2)) = True Then
      strData = strExc(2)
   End If
   If strData <> "" Then
      If bolIsBatchInvoice Then strData = strData & ",..." 'Added by Morgan 2016/2/15 整批列印請款單
      strData = Trim(Left(strData, 15))
      WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = strData
      'If Len(strData) > 12 Then
         WksRpt1.Range(Chr(xCols + intCounter) & nRow).Font.Size = 10
      'End If
   End If
   intCounter = intCounter + 1
     
   '應收金額：Credit
   strData = ""
   If bolChina = True Then  '大陸格式>>應收金額
      adoquery.CursorLocation = adUseClient
      If Not (IsNull(adoacc1k0.Fields("OAmount").Value) = False And adoacc1k0.Fields("OAmount").Value <> 0) Then
         If bolIsBatchInvoice = True Then
            adoquery.Open "select sum((a1k08 - nvl(a1k31, 0))),sum(a1k11),sum(nvl(a1k06, 0)),sum(nvl(a1k31, 0)) from acc1t0, acc1k0 where a1t02 in (select b.a1t01 from acc1t0 a, acc1t0 b where a.a1t01='" & adoacc1k0.Fields("DocNo").Value & "' and a.a1t02=b.a1t02(+)) and a1k01(+)=a1t01 and a1k29 is null and a1k25 is null", adoTaie, adOpenStatic, adLockReadOnly
         Else
            adoquery.Open "select sum((a1k08 - nvl(a1k31, 0))),sum(a1k11),sum(nvl(a1k06, 0)),sum(nvl(a1k31, 0)) from acc1k0 where substr(a1k01, 1, 8) = '" & Mid(adoacc1k0.Fields("DocNo").Value, 1, 8) & "' and A1K02 < 920201 and a1k29 is null and a1k12 is null and a1k25 is null union " & _
                       "select sum((a1k08 - nvl(a1k31, 0))),sum(a1k11),sum(nvl(a1k06, 0)),sum(nvl(a1k31, 0)) from acc1k0 where a1k01 = '" & adoacc1k0.Fields("DocNo").Value & "' and A1K02 > 920201 and a1k29 is null", adoTaie, adOpenStatic, adLockReadOnly
         End If
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields(0).Value) Then
               douDollar = 0
            Else
               Select Case strCurr
                  Case "NTD"
                     douDollar = Val(adoquery.Fields(1).Value) - Val(adoquery.Fields(2).Value) * Val(adoacc1k0.Fields("Rate").Value)
                  Case Else
                     douDollar = Val(adoquery.Fields(0).Value)
               End Select
            End If
         Else
            douDollar = 0
         End If
         adoquery.Close
      End If
      
      strData = ""
      If douDollar <> 0 Then
         strAmount = Format(douDollar, FDollar)
         strData = strAmount
         douAmount = douAmount + douDollar
         douTAmount = douTAmount + douDollar
      End If
      If strData <> "" Then
         WksRpt1.Range(Chr(xCols + intCounter) & nRow).NumberFormatLocal = FDollar
         WksRpt1.Range(Chr(xCols + intCounter) & nRow).HorizontalAlignment = xlRight
         WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = strData
      End If
      '已收金額
      strData = ""
      If IsNull(adoacc1k0.Fields("OAmount").Value) = False And adoacc1k0.Fields("OAmount").Value <> 0 Then
         Select Case strCurr
            Case "NTD"
               douODollar = Format(adoacc1k0.Fields("OAmount").Value * adoacc1k0.Fields("Rate").Value, FDollar)
            Case Else
               douODollar = Format(adoacc1k0.Fields("OAmount").Value, FDollar)
         End Select
         strAmount = Format(douODollar, FDollar)
         strData = strAmount
         douAmount = douAmount - Val(douODollar)
         douRAmount = douRAmount + Val(douODollar)
         douDollar = 0
      Else
         douODollar = 0
      End If
   Else
      '國外代理人-英文>>Credit
      If IsNull(adoacc1k0.Fields("OAmount").Value) = False And adoacc1k0.Fields("OAmount").Value <> 0 Then
         Select Case strCurr
            Case "NTD"
               douODollar = Format(adoacc1k0.Fields("OAmount").Value * adoacc1k0.Fields("Rate").Value, FDollar)
            Case Else
               douODollar = Format(adoacc1k0.Fields("OAmount").Value, FDollar)
         End Select
   
         If bolIsBatchInvoice Then
            strExc(0) = "select sum(a0z04) from acc1t0,acc1k0,acc0z0,acc0y0 where a1t02='" & adoacc1k0.Fields("a1t02") & "' and a1k01(+)=a1t01 and a0z02(+)= a1k01 and a0y01(+)=a0z01 and a0y02=" & adoacc1k0.Fields("DocDate")
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               Select Case strCurr
               Case "NTD"
                  douODollar = Format(RsTemp(0).Value * adoacc1k0.Fields("Rate").Value, FDollar)
               Case Else
                  douODollar = Format(RsTemp(0).Value, FDollar)
               End Select
            End If
         End If
   
         strAmount = Format(douODollar, FDollar)
         strData = strAmount
         douAmount = douAmount - Val(douODollar)
         douRAmount = douRAmount + Val(douODollar)
         douDollar = 0
      Else
         douODollar = 0
      End If
      
      If strData <> "" Then
         WksRpt1.Range(Chr(xCols + intCounter) & nRow).NumberFormatLocal = FDollar
         WksRpt1.Range(Chr(xCols + intCounter) & nRow).HorizontalAlignment = xlRight
         WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = strData
      End If
      '已收金額
      adoquery.CursorLocation = adUseClient
      If Not (IsNull(adoacc1k0.Fields("OAmount").Value) = False And adoacc1k0.Fields("OAmount").Value <> 0) Then
         If bolIsBatchInvoice = True Then
            '若部分單號結清時就該整批請款單而言仍為未結清故應收款不排除已結清者
            adoquery.Open "select sum((a1k08 - nvl(a1k31, 0))),sum(a1k11),sum(nvl(a1k06, 0)),sum(nvl(a1k31, 0)) from acc1t0, acc1k0 where a1t02= '" & adoacc1k0.Fields("a1t02") & "' and a1k01(+)=a1t01", adoTaie, adOpenStatic, adLockReadOnly
         Else

            adoquery.Open "select sum((a1k08 - nvl(a1k31, 0))),sum(a1k11),sum(nvl(a1k06, 0)),sum(nvl(a1k31, 0)) from acc1k0 where substr(a1k01, 1, 8) = '" & Mid(adoacc1k0.Fields("DocNo").Value, 1, 8) & "' and A1K02 < 920201 and a1k29 is null and a1k12 is null and a1k25 is null union " & _
                          "select sum((a1k08 - nvl(a1k31, 0))),sum(a1k11),sum(nvl(a1k06, 0)),sum(nvl(a1k31, 0)) from acc1k0 where a1k01 = '" & adoacc1k0.Fields("DocNo").Value & "' and A1K02 > 920201 and a1k29 is null", adoTaie, adOpenStatic, adLockReadOnly
         End If 'Added by Morgan 2016/2/15
         '2009/6/4 END
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields(0).Value) Then
               douDollar = 0
            Else
               Select Case strCurr
                  'Modify By Sindy 2013/1/18
                  Case "NTD"
                     douDollar = Val(adoquery.Fields(1).Value) - Val(adoquery.Fields(2).Value) * Val(adoacc1k0.Fields("Rate").Value)
                  Case Else
                     douDollar = Val(adoquery.Fields(0).Value)
               End Select
            End If
         Else
            douDollar = 0
         End If
         adoquery.Close
      End If
            
   End If
   intCounter = intCounter + 1
   
   '案件名稱：Amount
   strData = ""
   If bolChina = True Then '大陸格式>>案件名稱
      strData = GetPrjName(adoacc1k0.Fields("CaseNo1").Value & "-" & adoacc1k0.Fields("CaseNo2").Value & "-" & adoacc1k0.Fields("CaseNo3").Value & "-" & adoacc1k0.Fields("CaseNo4").Value, True)
      WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = convForm(CheckStr(strData), 12)
   Else
      '國外代理人-英文>>Amount
      If douDollar <> 0 Then
         strAmount = Format(douDollar, FDollar)
         strData = strAmount
         douAmount = douAmount + douDollar
         douTAmount = douTAmount + douDollar
      End If
      If strData <> "" Then
         WksRpt1.Range(Chr(xCols + intCounter) & nRow).NumberFormatLocal = FDollar
         WksRpt1.Range(Chr(xCols + intCounter) & nRow).HorizontalAlignment = xlRight
         WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = strData
      End If
   End If
   intCounter = intCounter + 1
   
   '申請人：天期(DAYS SINCE DN. SENT)
   strData = ""
   If bolChina = True Then '大陸格式>>申請人
      strData = GetPrjPeopleNum1(adoacc1k0.Fields("CaseNo1").Value & "-" & adoacc1k0.Fields("CaseNo2").Value & "-" & adoacc1k0.Fields("CaseNo3").Value & "-" & adoacc1k0.Fields("CaseNo4").Value)
      strData = GetPrjPeople1(strData)
   Else
      '國外代理人-英文>>天期
      intDays = CalculateDays(CADate(adoacc1k0.Fields("a1k02").Value), ServerDate)
      If intDays > 90 Then
         strData = "+90"
      ElseIf intDays > 60 Then
         strData = "+60"
      ElseIf intDays > 30 Then
         strData = "+30"
      End If
   End If
   If strData <> "" Then
      WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = convForm(CheckStr(strData), 12)
   End If
   
   Call PrintExcel_BPage
End Sub

'Added by Lydia 2024/12/31
Private Sub ProcSendMail(ByVal pFilePath As String)

   If bolPromoter Then
      ShowOutLookMail
   Else
      bolMailFailNoAlert = True
      bolMailSendOk = False
      txtSend = m_FNo 'Added by Lydia 2025/03/11
      '所內同仁
      If InStr(strReceiver, "@") = 0 Then
         PUB_SendMail strUserNum, strReceiver, "", "Statement of Accounts (" & m_FNo & ")", GetMailContent, , pFilePath, True, , , strEmailCC
      Else
         PUB_SendMail strUserNum, strReceiver, "", "Statement of Accounts (" & m_FNo & ")", GetMailContent, , pFilePath, True, True, True, strEmailCC, strAccMailBox, "TAI E INTERNATIONAL PATENT & LAW OFFICE", strAccMailBox
      End If
      bolMailFailNoAlert = False
      If bolMailSendOk = False Then
         If strMailFailList(0) <> "" Then
            ReDim Preserve strMailFailList(UBound(strMailFailList) + 1)
         End If
         strMailFailList(UBound(strMailFailList)) = m_FNo & " : " & strEMailBox
      'Added by Lydia 2025/03/11
      Else
         'SEQ=1>>EMAIL
         strSql = "insert into rdatafactory (formname,id,seqno,rowseq,r001) values ('" & m_FrmName & "','" & strUserNum & "','1'," & m_iMailCount & ",'" & m_FNo & "') "
         cnnConnection.Execute strSql
      'end 2025/03/11
      End If
   End If
End Sub

'Added by Lydia 2025/03/11
Private Sub GetEmailCount()
   strExc(0) = "select count(*) cnt from rdatafactory where upper(formname) like '" & UCase(Me.Name) & "%' and id ='" & strUserNum & "' and seqno='1' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      lblSend2.Caption = "今日Email筆數：" & RsTemp.Fields("cnt")
   Else
      lblSend2.Caption = ""
   End If
End Sub
