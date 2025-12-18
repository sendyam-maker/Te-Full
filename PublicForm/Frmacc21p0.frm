VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc21p0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "相同案件性質整批請款作業"
   ClientHeight    =   4704
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11484
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4704
   ScaleWidth      =   11484
   Begin VB.CheckBox Check1 
      Caption         =   "使用預留單號"
      Enabled         =   0   'False
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
      Left            =   3912
      TabIndex        =   47
      Top             =   144
      Width           =   1560
   End
   Begin VB.CommandButton Command6 
      Caption         =   "預留單號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   144
      TabIndex        =   46
      Top             =   48
      Width           =   1020
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   1176
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   84
      Width           =   1215
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   2616
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   84
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "查詢結果"
      Enabled         =   0   'False
      Height          =   4212
      Left            =   5736
      TabIndex        =   37
      Top             =   432
      Width           =   5676
      Begin VB.CheckBox chkAutoSave 
         Caption         =   "不要修改請款明細"
         Height          =   252
         Left            =   3672
         TabIndex        =   43
         Top             =   3480
         Width           =   1836
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "請款"
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
         Index           =   1
         Left            =   768
         Style           =   1  '圖片外觀
         TabIndex        =   40
         Top             =   3792
         Width           =   2064
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "取消"
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
         Index           =   0
         Left            =   3168
         Style           =   1  '圖片外觀
         TabIndex        =   39
         Top             =   3768
         Width           =   1680
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
         Height          =   3228
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   5436
         _ExtentX        =   9589
         _ExtentY        =   5694
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblCount 
         Caption         =   "0"
         Height          =   180
         Left            =   1524
         TabIndex        =   42
         Top             =   3480
         Width           =   492
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "已勾選案件數："
         Height          =   180
         Index           =   1
         Left            =   168
         TabIndex        =   41
         Top             =   3504
         Width           =   1260
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "查詢"
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
      Left            =   120
      Style           =   1  '圖片外觀
      TabIndex        =   36
      Top             =   3936
      Width           =   1680
   End
   Begin VB.ComboBox cboAddrPrinter 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1620
      TabIndex        =   30
      Top             =   4344
      Width           =   4020
   End
   Begin VB.ComboBox Combo1 
      Height          =   276
      Left            =   3960
      TabIndex        =   8
      Top             =   2880
      Width           =   1560
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
      Height          =   300
      Left            =   1164
      TabIndex        =   0
      Top             =   504
      Width           =   4344
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "取消"
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
      Left            =   3960
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   3936
      Width           =   1680
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "確定"
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
      Left            =   1848
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   3936
      Width           =   2064
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
      Height          =   300
      Left            =   2412
      MaxLength       =   1
      TabIndex        =   11
      Top             =   3564
      Width           =   528
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
      Height          =   300
      Left            =   3948
      MaxLength       =   9
      TabIndex        =   10
      Top             =   3240
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
      Height          =   315
      Left            =   1164
      MaxLength       =   9
      TabIndex        =   9
      Top             =   3204
      Width           =   1572
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
      Left            =   1164
      MaxLength       =   9
      TabIndex        =   2
      Top             =   1200
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
      Height          =   315
      Left            =   1164
      TabIndex        =   7
      Top             =   2844
      Width           =   1344
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
      Left            =   1164
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1560
      Width           =   1572
   End
   Begin VB.TextBox Text1 
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
      Left            =   2736
      TabIndex        =   16
      Top             =   1560
      Width           =   2772
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
      Height          =   315
      Left            =   1164
      MaxLength       =   9
      TabIndex        =   1
      Top             =   840
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1176
      TabIndex        =   4
      Top             =   2124
      Width           =   1572
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   2916
      TabIndex        =   5
      Top             =   2124
      Width           =   1572
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
      Left            =   1176
      TabIndex        =   6
      Top             =   2484
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   16777215
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
   Begin VB.Frame Frame1 
      Caption         =   "電腦中心專用"
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   30
      TabIndex        =   33
      Top             =   4740
      Visible         =   0   'False
      Width           =   5625
      Begin VB.TextBox Text8 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1755
         TabIndex        =   35
         Top             =   150
         Width           =   3795
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "特殊條件(SQL)："
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
         Left            =   60
         TabIndex        =   34
         Top             =   210
         Width           =   1680
      End
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   2625
      Y1              =   228
      Y2              =   228
   End
   Begin MSForms.TextBox Text6 
      Height          =   336
      Left            =   2736
      TabIndex        =   22
      Top             =   1200
      Width           =   2772
      VariousPropertyBits=   671107097
      BackColor       =   14737632
      Size            =   "4895;582"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text3 
      Height          =   336
      Left            =   2736
      TabIndex        =   14
      Top             =   840
      Width           =   2772
      VariousPropertyBits=   671107097
      BackColor       =   14737632
      Size            =   "4895;582"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "FCT的101(申請)會自動將108(主張優先權 )併入"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   192
      Left            =   1176
      TabIndex        =   32
      Top             =   1884
      Width           =   3888
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "地址條印表機"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   132
      TabIndex        =   31
      Top             =   4380
      Width           =   1440
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "系統別"
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
      Left            =   216
      TabIndex        =   29
      Top             =   504
      Width           =   972
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   12
      Top             =   4020
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "是否列印申請人(Y/N)"
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
      Left            =   216
      TabIndex        =   28
      Top             =   3564
      Width           =   2208
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "請款對象"
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
      Left            =   2976
      TabIndex        =   27
      Top             =   3240
      Width           =   972
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "列印對象"
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
      Left            =   216
      TabIndex        =   26
      Top             =   3204
      Width           =   972
   End
   Begin VB.Label Label9 
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
      Height          =   252
      Left            =   2976
      TabIndex        =   25
      Top             =   2880
      Width           =   972
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "請款日期"
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
      Left            =   216
      TabIndex        =   24
      Top             =   2520
      Width           =   972
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "申請人"
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
      Left            =   216
      TabIndex        =   23
      Top             =   1200
      Width           =   972
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "%"
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
      Left            =   2556
      TabIndex        =   21
      Top             =   2856
      Width           =   312
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "折扣"
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
      Left            =   216
      TabIndex        =   20
      Top             =   2844
      Width           =   972
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "發文日期"
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
      Left            =   216
      TabIndex        =   19
      Top             =   2160
      Width           =   972
   End
   Begin VB.Label Label2 
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
      Left            =   2772
      TabIndex        =   18
      Top             =   2124
      Width           =   252
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "案件性質"
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
      Index           =   0
      Left            =   216
      TabIndex        =   17
      Top             =   1560
      Width           =   972
   End
   Begin VB.Label Label3 
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
      Height          =   252
      Left            =   216
      TabIndex        =   15
      Top             =   864
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc21p0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/12 改成Form2.0 (Text3,Text6)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo By Sindy 2010/8/12 日期欄已修改
Option Explicit

Public adoquery As New ADODB.Recordset
Dim strTemp1 As Variant
Dim strTemp2 As Variant
Dim i As Integer
Dim j As Integer
Dim s As Integer
Dim strSystemKind As String
Dim strSql As String
Dim strCaseName As String
Dim strName As String
Dim strDebitNo As String
Dim strYes As String
Dim douExchange As Double
Dim douFAmount As Double
Dim douDAmount As Double
Dim strDNo As String
Dim strDNoArray As Variant
Dim strCP09 As String 'Added by Morgan 2014/8/15
'Added by Morgan 2023/8/8
Dim bolQuery As Boolean
Dim strSelList As String

Private Sub Combo1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
   If Combo1 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc1y0", "a1y01", Combo1, Label9) = False Then
      Cancel = True
      Combo1.SetFocus
   End If
End Sub

Private Sub Command1_Click()
   FormClear
End Sub

Private Sub Command2_Click()
   If TxtValidate = False Then Exit Sub
    '代理人
   If Text2 <> MsgText(601) Then
      strCon1 = Text2
   Else
      strCon1 = MsgText(601)
   End If
    '幣別
   If Combo1 <> MsgText(601) Then
      strCon2 = Combo1
   Else
      strCon2 = MsgText(601)
   End If
    '列印對象
   If Text9 <> MsgText(601) Then
      strCon3 = Text9
   Else
      strCon3 = MsgText(601)
   End If
    '請款對象
   If Text10 <> MsgText(601) Then
      strCon4 = Text10
   Else
      strCon4 = MsgText(601)
   End If
    '請款日期
   If MaskEdBox3.Text <> MsgText(601) And MaskEdBox3.Text <> MsgText(29) Then
      strCon5 = Val(FCDate(MaskEdBox3.Text))
   Else
      strCon5 = MsgText(601)
   End If
    '是否列印申請人
   If Text11 <> MsgText(601) Then
      strCon6 = Text11
   Else
      strCon6 = MsgText(601)
   End If
   Screen.MousePointer = vbHourglass
   If ProcessData = False Then
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   '請款單號起迄
   If strDebitNo <> MsgText(601) Then
      strCon7 = strDebitNo
      strCon8 = strDNoArray(0) 'Add by Morgan 2007/1/17
      strItemNo = strDNoArray(0) 'Add by Morgan 2014/8/15
   Else
      strCon7 = MsgText(601)
   End If
   Screen.MousePointer = vbDefault
   
   tool3_enabled
   Screen.MousePointer = vbHourglass
   '用全域變數傳本所號到下畫面
   strExc(1) = Mid(strName, 1, Len(strName) - 9)
   strExc(2) = Mid(strName, (Len(strName) - 8), 6)
   strExc(3) = Mid(strName, (Len(strName) - 2), 1)
   strExc(4) = Mid(strName, (Len(strName) - 1), 2)
 
'Modified by Morgan 2016/11/2 不同案件性質可合併請款--陳金蓮
''Modified by Morgan 2014/8/15
''   With Frmacc21p1
''      .Text19.Text = douExchange 'Added by Morgan 2012/12/20
'   With Frmacc21h1
'      dblRate = douExchange
'      strCon9 = strCP09
'      strFormLink = Name
'      .m_Discount = Text5.Text
'      Set .m_FromForm = Me
''end 2014/8/15
'      .Show
'      .strDNoArray = strDNoArray
'   End With
   goNextForm
'end 2016/11/2
   Screen.MousePointer = vbDefault
End Sub

'Added by Morgan 2016/11/1
Private Function goNextForm() As Boolean
   Dim stSQL As String, intQ As Integer, strDnNo As String
   Dim strCP09 As String, strDebitNoStart As String, strCP10List As String
   Dim rsQuery As ADODB.Recordset
   Dim intDnCount As Integer, intCP10Count As Integer, bolShowMsg As Boolean
   
   If strDebitNo <> "" Then
      stSQL = "select cp60,cp10,cp09,cp16,cp17 from caseprogress where cp60>='" & strDNoArray(0) & "' and cp60<='" & strDebitNo & "'" & _
         " and not exists(select * from acc1l0 where a1l01=cp60)" & _
         " order by 1,3"
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         With rsQuery
         strDebitNoStart = .Fields("cp60") '目前請款單號
         strCP09 = .Fields("cp09") '最小收文號
         strCP10List = .Fields("cp10") '本次請款案件性質(可以多個)
         strDnNo = .Fields("cp60")
         intDnCount = 1
         intCP10Count = 1
         .MoveNext
         Do While Not .EOF
            If .Fields("cp60") = strDebitNoStart Then
               intCP10Count = intCP10Count + 1
               strCP10List = strCP10List & "," & .Fields("cp10")
            ElseIf strDnNo <> .Fields("cp60") Then
               intDnCount = intDnCount + 1
               strDnNo = .Fields("cp60")
            End If
            .MoveNext
         Loop
         End With
         
         '檢查是否還有明細要輸入(要先檢查否則載入21h0後第1張請款單的明細就會自動存檔)
         '檢查條件:有相同案性質且沒有不同案件性質的待輸明細請款單數量小於所有待輸明細的請款單數量
         bolShowMsg = False
         stSQL = "select count(*) from (select cp60 a1 from caseprogress where cp60>='" & strDNoArray(0) & "' and cp60<='" & strDebitNo & "'" & _
            " and not exists(select * from acc1l0 where a1l01=cp60) and cp10 in ('" & Replace(strCP10List, ",", "','") & "')" & _
            " group by cp60 having count(*)=" & intCP10Count & ") a where not exists(select * from caseprogress where cp60=a1 and cp10 not in ('" & Replace(strCP10List, ",", "','") & "'))"
         intQ = 1
         Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
         If intQ = 1 Then
            If rsQuery(0) < intDnCount Then
               bolShowMsg = True
            End If
         End If
         
         strCon7 = strDebitNo '最大請款單號
         strCon8 = strDebitNoStart
         strItemNo = strDebitNoStart
         With Frmacc21h1
            dblRate = douExchange
            strCon9 = strCP09
            strFormLink = Name
            .m_Discount = Text5.Text
            .m_CP10List = strCP10List
            .m_AutoSave = (chkAutoSave.Value = vbChecked)
            Set .m_FromForm = Me
            .Show
            .strDNoArray = strDNoArray
         End With
         Me.Hide
         
         If bolShowMsg Then
            MsgBox "請留意！本次明細輸完後還有其他種案件性質/費用組合的明細要輸入...", vbExclamation
         End If
         goNextForm = True
         
      ElseIf intQ = 0 Then
         'Added by Morgan 2023/8/9
         If strDebitNo <> "" Then
            MsgBox "請款完成！" & vbCrLf & vbCrLf & "請款單號：" & strDNoArray(0) & " - " & strDebitNo, vbInformation
         End If
         'end 2023/8/9
         strDebitNo = ""
      End If
   End If
   Set rsQuery = Nothing
End Function

Private Sub Command3_Click()
   bolQuery = True
   chkAutoSave.Value = vbUnchecked
   Command2.Value = True
   bolQuery = False
End Sub

Private Sub Command4_Click(Index As Integer)
   Dim ii As Integer
   
   strSelList = ""
   If Index = 1 Then
      With grdDataList
      For ii = 1 To .Rows - 1
         If .TextMatrix(ii, 0) = "V" Then
            strSelList = strSelList & "," & .TextMatrix(ii, 5)
         End If
      Next ii
      End With
      If strSelList = "" Then
         MsgBox "請勾選要請款的案件！", vbExclamation
         Exit Sub
      End If
   End If
   
   If strSelList <> "" Then
      strSelList = Mid(strSelList, 2)
      Command2.Value = True
      strSelList = ""
   End If
   
   Frame2.Enabled = False
   Frame2.Visible = False
End Sub

'Added by Morgan 2023/8/9
Private Sub Command6_Click()
      
   If GetRsvDN() = True Then
      MsgBox "您尚有預留單號未使用，已自動載入！", vbInformation
   Else
   
      strExc(0) = InputBox("請輸入要預留單號的數量：", Me.Caption & "-" & Command6.Caption)
      If Val(strExc(0)) > 0 Then
         If Val(strExc(0)) > 10 Then
            If MsgBox("系統將預留 " & Val(strExc(0)) & " 個單號，是否確定要繼續？", vbYesNo) = vbNo Then
               Exit Sub
            End If
         End If
         If PUB_RsvDN(CInt(strExc(0)), strExc(1), strExc(2)) = True Then
            Text14 = strExc(1)
            Text13 = strExc(2)
            Check1.Enabled = True
            Check1.Value = 1
         End If
      End If
   End If
End Sub

'Added by Morgan 2023/8/9
Private Function GetRsvDN() As Boolean
   If PUB_GetRsvDN(strExc(1), strExc(2)) = True Then
      Text14 = strExc(1)
      Text13 = strExc(2)
      Check1.Enabled = True
      Check1.Value = 1
      GetRsvDN = True
   Else
      Text14 = ""
      Text13 = ""
      Check1.Enabled = False
      Check1.Value = 0
   End If
End Function

Private Sub Form_Activate()
   
   '93.3.16 ADD BY SONIA
   If IsObject(mdiMain) Then
      ToolShow
   End If
   '93.3.16 END
   strFormName = Name
   strCon1 = MsgText(601)
   strCon2 = MsgText(601)
   strCon3 = MsgText(601)
   strCon4 = MsgText(601)
   strCon5 = MsgText(601)
   strCon6 = MsgText(601)
   
   goNextForm 'Added by Morgan 2016/11/2
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
   
'   Dim intX As Integer
'   Dim intY As Integer
'   Dim sglWidth As Single
'   Dim sglHeight As Single
'
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   Me.Width = 5800
'   Me.Height = 4300
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath1)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   
   PUB_InitForm Me, Me.Width, Me.Height
   
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   MaskEdBox3.Mask = DFormat
   MaskEdBox3.Text = CFDate(strSrvDate(2)) 'Added by Morgan 2017/3/8 預設系統日--陳金蓮(有發生輸錯年度情形,通常都是請當天)
   
   Combo1 = "USD"
   Text12 = GetSystemKindByNick
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from acc1y0 order by a1y01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoquery.EOF = False
      Combo1.AddItem adoquery.Fields("a1y01").Value
      adoquery.MoveNext
   Loop
   adoquery.Close
   
   'Removed by Morgan 2015/1/23 開放其他幣別
   'Add by Morgan 2009/9/17 暫時取消幣別選項,只能請美金
   'Label9.Visible = False
   'Combo1.Visible = False
   'end 2015/1/23
   
   'Added by Morgan 2014/8/19
   '地址條印表機
   PUB_SetPrinter Me.Name, Me.cboAddrPrinter
   'end 2014/8/19
   
   'Added by Morgan 2021/8/23
   If Pub_StrUserSt03 = "M51" Then
      Frame1.Visible = True
      Me.Height = 5808
   Else
      Me.Height = 5124
   End If
   'end 2021/8/23
   
   Me.Width = 5772 'Added by Morgan 2023/8/8
   
   'Added by Morgan 2023/8/9
   If GetRsvDN() = True Then
      MsgBox "您尚有預留單號未使用，已自動載入！", vbInformation
   End If
   'end 2023/8/9
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   strItemNo = "" 'Added by Morgan 2014/8/18
   
   If Me.cboAddrPrinter.Text <> Me.cboAddrPrinter.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cboAddrPrinter.Name, "0", "0", Me.cboAddrPrinter.Text
   End If
   
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc21p0 = Nothing
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   'Add by Morgan 2011/4/15
   If MaskEdBox1.Tag <> MaskEdBox1 Then
       SetDefault
   End If
   MaskEdBox1.Tag = MaskEdBox1
End Sub

Private Sub MaskEdBox2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
   'Add by Morgan 2011/4/15
   If MaskEdBox2.Tag <> MaskEdBox2 Then
       SetDefault
   End If
   MaskEdBox2.Tag = MaskEdBox2
End Sub

Private Sub MaskEdBox3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub MaskEdBox3_LostFocus()
'Modify by Morgan 2011/4/15 需考慮各案的設定改比照單筆請款模式
'    '折扣
'    Me.Text5.Text = GetA1L07Disc("" & strTemp2(0), Me.Text4.Text, IIf(Me.Text2.Text <> "", Me.Text2.Text, Me.Text7.Text), Replace(Me.MaskEdBox3.Text, "/", ""))
'    '列印對象
'    Me.Text9.Text = GetA1K27("" & strTemp2(0), Me.Text4.Text, IIf(Me.Text2.Text <> "", Me.Text2.Text, Me.Text7.Text))
'    If Me.Text9.Text = "" Then Me.Text9.Text = IIf(Me.Text2.Text <> "", Me.Text2.Text, Me.Text7.Text)
'    '請款對象
'    Me.Text10.Text = GetA1K28("" & strTemp2(0), Me.Text4.Text, IIf(Me.Text2.Text <> "", Me.Text2.Text, Me.Text7.Text))
'    If Me.Text10.Text = "" Then Me.Text10.Text = IIf(Me.Text2.Text <> "", Me.Text2.Text, Me.Text7.Text)
'    '是否列印申請人
'    'Modify By Sindy 2011/3/8 增加系統別
'    'Me.Text11.Text = GetA1K04(IIf(Me.Text2.Text <> "", Me.Text2.Text, Me.Text7.Text))
'    Me.Text11.Text = GetA1K04("" & strTemp2(0), IIf(Me.Text2.Text <> "", Me.Text2.Text, Me.Text7.Text))

End Sub

Private Sub MaskEdBox3_Validate(Cancel As Boolean)
   'Add by Morgan 2011/4/15
   If MaskEdBox3.Tag <> MaskEdBox3 Then
       SetDefault
   End If
   MaskEdBox3.Tag = MaskEdBox3
End Sub

'Added by Morgan 2015/1/23
Private Sub Text10_Change()
   If Len(Text10) = 9 Then
      SetCurr
   End If
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 = MsgText(601) Then
      Exit Sub
   End If
   Select Case Len(Text10)
      Case 6
         Text10 = Text10 & "000"
      Case 8
         Text10 = Text10 & "0"
   End Select
   If ExistCheck("customer", "cu01", Mid(Text10, 1, 8), Label11, False) = False Then
      If ExistCheck("fagent", "fa01", Mid(Text10, 1, 8), Label11) = False Then
         Cancel = True
         Text10.SetFocus
      End If
   End If
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

Private Sub Text12_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
     strTemp1 = Split(UCase(GetSystemKindByNick), ",")
     strTemp2 = Split(UCase(Text12), ",")
     For i = 0 To UBound(strTemp2)
        s = 0
        For j = 0 To UBound(strTemp1)
            If strTemp1(j) = strTemp2(i) Then
                s = 1
                Exit For
            '2011/4/22 ADD BY SONIA FCP程序可輸入P及CFP請款單
            'Modified by Morgan 2014/8/19 部門改抓變數否則系統別多的時候會重複抓資料庫多次
            'ElseIf GetStaffDepartment(strUserNum) = "F22" And (strTemp2(i) = "P" Or strTemp2(i) = "PS" Or strTemp2(i) = "CFP" Or strTemp2(i) = "CPS") Then
            ElseIf Pub_StrUserSt03 = "F22" And (strTemp2(i) = "P" Or strTemp2(i) = "PS" Or strTemp2(i) = "CFP" Or strTemp2(i) = "CPS") Then
            'end 2014/8/19
                s = 1
                Exit For
            '2011/4/22 END
            End If
        Next j
        If s = 0 Then
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限 ", , "權限問題")
            Cancel = True
            Text12.SetFocus
            Exit Sub
        End If
     Next i
     
     'Add by Morgan 2011/4/15
     If Text12.Tag <> Text12 Then
         SetDefault
     End If
     Text12.Tag = Text12
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 = MsgText(601) Then
        Me.Text3.Text = ""
   Else
      Select Case Len(Text2)
         Case 6
            Text2 = Text2 & "000"
         Case 8
            Text2 = Text2 & "0"
      End Select
       If Me.Text2.Text <> "" Then Me.Text2.Text = Left(Me.Text2.Text & "000000000", 9)
      If ExistCheck("customer", "cu01", Mid(Text2, 1, 8), Label3, False) = False Then
         If ExistCheck("fagent", "fa01", Mid(Text2, 1, 8), Label3) = False Then
            Cancel = True
            Text2.SetFocus
         End If
      End If
      Text3 = FagentQuery(Text2, 2)
      If Text3 = MsgText(601) Then
         Text3 = FagentQuery(Text2, 1)
      End If
   End If

   'Add by Morgan 2011/4/15
   If Text2.Tag <> Text2 Then
       SetDefault
   End If
   Text2.Tag = Text2
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
   If Text4 <> MsgText(601) Then
      If adoquery.State = adStateOpen Then
         adoquery.Close
      End If
      adoquery.CursorLocation = adUseClient
      'Modified by Morgan 2017/2/15 案件性質改可輸入多個
      'Modified by Morgan 2017/2/20 改回只能輸入一個
      'adoquery.Open "select decode(substr(cu10, 1, 3), '020', cpm04, cpm03),cpm02 from casepropertymap, customer where cpm01 = '" & strTemp2(0) & "' and cpm02 in ( '" & Replace(Replace(Text4, " ", ""), ",", "','") & "' ) and cu01 = '" & Mid(Text7, 1, 8) & "' and cu02 = '" & Mid(Text7, 9, 1) & "' union " & _
                    "select decode(substr(fa10, 1, 3), '020', cpm04, cpm03),cpm02 from casepropertymap, fagent where cpm01 = '" & strTemp2(0) & "' and cpm02 in ( '" & Replace(Replace(Text4, " ", ""), ",", "','") & "' ) and fa01 = '" & Mid(Text2, 1, 8) & "' and fa02 = '" & Mid(Text2, 9, 1) & "' order by 2", adoTaie, adOpenStatic, adLockReadOnly
      adoquery.Open "select decode(substr(cu10, 1, 3), '020', cpm04, cpm03),cpm02 from casepropertymap, customer where cpm01 = '" & strTemp2(0) & "' and cpm02='" & Text4 & "' and cu01 = '" & Mid(Text7, 1, 8) & "' and cu02 = '" & Mid(Text7, 9, 1) & "' union " & _
                    "select decode(substr(fa10, 1, 3), '020', cpm04, cpm03),cpm02 from casepropertymap, fagent where cpm01 = '" & strTemp2(0) & "' and cpm02='" & Text4 & "' and fa01 = '" & Mid(Text2, 1, 8) & "' and fa02 = '" & Mid(Text2, 9, 1) & "' order by 2", adoTaie, adOpenStatic, adLockReadOnly
      'end 2017/2/20
      If adoquery.RecordCount <> 0 Then
         If IsNull(adoquery.Fields(0).Value) Then
            Text1 = MsgText(601)
         Else
            Text1 = adoquery.Fields(0).Value
            'Removed by Morgan 2017/2/20
            ''Added by Morgan 2017/2/15
            'Text4 = adoquery.Fields(1).Value
            'If adoquery.RecordCount > 1 Then
            '   adoquery.MoveNext
            '   Do While Not adoquery.EOF
            '      Text1 = Text1 & "," & adoquery.Fields(0).Value
            '      Text4 = Text4 & "," & adoquery.Fields(1).Value
            '      adoquery.MoveNext
            '   Loop
            'End If
            ''end 2017/2/15
            'end 2017/2/20
         End If
      Else
         Text1 = MsgText(601)
      End If
      adoquery.Close
   End If
   
   'Add by Morgan 2011/4/15
   If Text4.Tag <> Text4 Then
       SetDefault
   End If
   Text4.Tag = Text4
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
   If Text7 = MsgText(601) Then
      Me.Text6.Text = ""
   Else
      Select Case Len(Text7)
         Case 6
            Text7 = Text7 & "000"
         Case 8
            Text7 = Text7 & "0"
      End Select
       If Me.Text7.Text <> "" Then Me.Text7.Text = Left(Me.Text7.Text & "000000000", 9)
      If ExistCheck("customer", "cu01", Mid(Text7, 1, 8), Label7, False) = False Then
         If ExistCheck("fagent", "fa01", Mid(Text7, 1, 8), Label7) = False Then
            Cancel = True
            Text7.SetFocus
         End If
      End If
      Text6 = CustomerQuery(Text7, 2)
      If Text6 = MsgText(601) Then
         Text6 = CustomerQuery(Text7, 1)
      End If
   End If

   'Add by Morgan 2011/4/15
   If Text7.Tag <> Text7 Then
       SetDefault
   End If
   Text7.Tag = Text7
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   If Text9 = MsgText(601) Then
      Exit Sub
   End If
   Select Case Len(Text9)
      Case 6
         Text9 = Text9 & "000"
      Case 8
         Text9 = Text9 & "0"
   End Select
   If ExistCheck("customer", "cu01", Mid(Text9, 1, 8), Label10, False) = False Then
      If ExistCheck("fagent", "fa01", Mid(Text9, 1, 8), Label10) = False Then
         Cancel = True
         Text9.SetFocus
      End If
   End If
End Sub

'*************************************************
'  產生請款單主檔資料
'
'*************************************************
Public Function ProcessData() As Boolean
   Dim strCase(5) As String
   Dim strFaNo As String
   Dim strA1K32 As String 'Add by Morgan 2010/7/22
   Dim iCount As Integer 'Added by Morgan 2023/8/9
   
On Error GoTo Checking
      
'Modify by Morgan 2009/9/17
'   Select Case Combo1
'      Case "USD"
'         If adoquery.State = adStateOpen Then
'            adoquery.Close
'         End If
'         adoquery.CursorLocation = adUseClient
'         adoquery.Open "select usxr02 from usxrate order by usxr01 desc"
'         If adoquery.RecordCount <> 0 Then
'            If IsNull(adoquery.Fields("usxr02").Value) Then
'               douExchange = 35
'            Else
'               douExchange = Val(adoquery.Fields("usxr02").Value)
'            End If
'         Else
'            douExchange = 35
'         End If
'         adoquery.Close
'      Case Else
'         douExchange = 1
'   End Select
   'Modified by Morgan 2024/3/12
   'douExchange = PUB_GetUSXRate_1(Replace(MaskEdBox3.Text, "/", ""), Combo1)
   douExchange = PUB_GetRate(Replace(MaskEdBox3.Text, "/", ""), Combo1, Text10, Text12, Text7)
   'end 2024/3/12
'end 2009/9/17

   dblRate = douExchange
   'adoTaie.BeginTrans 'Removed by Morgan 2023/7/20 移到下面
   strSql = ""
   strName = ""
   strSystemKind = ""
   strDNo = ""
   For i = 0 To UBound(strTemp2)
      strSystemKind = strSystemKind & "'" & strTemp2(i) & "',"
   Next i
   If strSystemKind <> "" Then
      strSystemKind = Mid(strSystemKind, 1, Len(strSystemKind) - 1)
   End If
   'If Text2 <> MsgText(601) Then
   '   strSQL = strSQL & " and cp44 = '" & Text2 & "'"
   'End If
   If Text4 <> MsgText(601) Then
      'Modified by Morgan 2017/2/15 案件性質改可輸入多個
      'Modified by Morgan 2017/2/20 改回只能輸入一個,另加判斷FCT101要一併抓108
      'strSql = strSql & " and cp10 in ( '" & Replace(Replace(Text4, " ", ""), ",", "','") & "' ) "
      If InStr(strSystemKind, "'FCT'") > 0 And Text4 = "101" Then
         strSql = strSql & " and (cp10 = '101' or cp01||cp10='FCT108')"
      Else
         strSql = strSql & " and cp10 = '" & Text4 & "'"
      End If
      'end 2017/2/20
      'end 2017/2/15
   End If
   
   'Added by Morgan 2021/8/23
   If Text8 <> "" Then
      strSql = strSql & Text8
   End If
   'end 2021/8/23
   
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and cp27 >= " & Val(CADate(FCDate(MaskEdBox1.Text))) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and cp27 <= " & Val(CADate(FCDate(MaskEdBox2.Text))) & ""
   End If
   If strSystemKind <> "" Then
      strSql = strSql & " and cp01 in (" & strSystemKind & ")"
   End If
   '代理人
   If Text2 <> MsgText(601) Then
      strCase(0) = strCase(0) & " and pa75 = '" & Text2 & "'"
      strCase(1) = strCase(1) & " and tm44 = '" & Text2 & "'"
      strCase(2) = strCase(2) & " and lc22 = '" & Text2 & "'"
      strCase(4) = strCase(4) & " and sp26 = '" & Text2 & "'"
   End If
   '申請人
   If Text7 <> MsgText(601) Then
      strCase(0) = strCase(0) & " and pa26 = '" & Text7 & "'"
      strCase(1) = strCase(1) & " and tm23 = '" & Text7 & "'"
      strCase(2) = strCase(2) & " and lc11 = '" & Text7 & "'"
      strCase(4) = strCase(4) & " and sp08 = '" & Text7 & "'"
   End If
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   'Modify by Morgan 2009/12/25 +代理人欄位(只下申請人條件時用)
   'Modify by Morgan 2010/7/22 +CP12
   'Modified by Morgan 2023/4/27 排除不請款
   'Modified by Morgan 2023/8/8 +cp10,pa09
   strExc(0) = "select cp01, cp02, cp03, cp04, cp09, cp16, nvl(cp17, 0) as cp17,pa75,cp12,cp10,pa09 from caseprogress, patent where cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and cp60 is null and cp20 is null and (cp16 is not null and cp16 <> 0)" & strCase(0) & strSql & " union " & _
      "select cp01, cp02, cp03, cp04, cp09, cp16, nvl(cp17, 0) as cp17,tm44,cp12,cp10,tm10 from caseprogress, trademark where cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and cp60 is null and cp20 is null and (cp16 is not null and cp16 <> 0)" & strCase(1) & strSql & " union " & _
      "select cp01, cp02, cp03, cp04, cp09, cp16, nvl(cp17, 0) as cp17,lc22,cp12,cp10,lc15 from caseprogress, lawcase where cp01 = lc01 and cp02 = lc02 and cp03 = lc03 and cp04 = lc04 and cp60 is null and cp20 is null and (cp16 is not null and cp16 <> 0)" & strCase(2) & strSql & " union " & _
      "select cp01, cp02, cp03, cp04, cp09, cp16, nvl(cp17, 0) as cp17,sp26,cp12,cp10,sp09 from caseprogress, servicepractice where cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and cp60 is null and cp20 is null and (cp16 is not null and cp16 <> 0)" & strCase(4) & strSql
   
   'Added by Morgan 2023/8/8
   If bolQuery Then
      strExc(0) = "select 'V',cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CaseNo" & _
         ",decode(pa09,'020',cpm04,cpm03) Pty,cp16,cp17,cp09" & _
         " from (" & strExc(0) & ") X,casepropertymap where cpm01(+)=cp01 and cpm02(+)=cp10"
   ElseIf strSelList <> "" Then
      strExc(0) = "select * from (" & strExc(0) & ") X where instr('" & strSelList & "',cp09)>0"
   End If
   'end 2023/8/8
   
   strExc(0) = strExc(0) & " order by cp01 asc, cp02 asc, cp03 asc, cp04 asc"
   adoquery.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount = 0 Then
      adoquery.Close
      ProcessData = False
      'adoTaie.RollbackTrans 'Removed by Morgan 2023/7/20 BeginTrans已移到下面
      MsgBox MsgText(28), , MsgText(5)
      Exit Function
   End If
   
   'Added by Morgan 2021/8/23
   If Text8 <> "" Then
      If MsgBox("共有 " & adoquery.RecordCount & " 筆，是否要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
         adoquery.Close
         ProcessData = False
         'adoTaie.RollbackTrans 'Removed by Morgan 2023/7/20 BeginTrans已移到下面
         MsgBox "已取消！", vbOKOnly + vbInformation
         Exit Function
      End If
   End If
   'end 2021/8/23
   
   'Added by Morgan 2023/7/20
   If bolQuery Then
      lblCount = ""
      grdDataList.FixedCols = 0
      Set grdDataList.Recordset = adoquery
      grdDataList.FixedCols = 2
      SetGridHead
      lblCount = adoquery.RecordCount
      Frame2.Top = Label13.Top
      Frame2.Left = 0
      Frame2.Visible = True
      Frame2.Enabled = True
      
      ProcessData = False
      Exit Function
   End If
   'end 2023/7/20
   
   adoTaie.BeginTrans 'Added by Morgan 2023/7/20 從上面移下來
   
   'Added by Morgan 2023/8/9
   '檢查預留單號是否夠用
   If Check1.Value = vbChecked Then
      iCount = 0
      With adoquery
      .MoveFirst
      strName = ""
      Do While Not .EOF
         strCaseName = .Fields("cp01").Value & .Fields("cp02").Value & .Fields("cp03").Value & .Fields("cp04").Value
         If strName <> strCaseName Then
            iCount = iCount + 1
            strName = strCaseName
         End If
         .MoveNext
      Loop
      .MoveFirst
      
      End With
      If PUB_LockRsvDN() = True Then
         GetRsvDN
         intI = Val(Mid(Text13, 2)) - Val(Mid(Text14, 2)) + 1
         If iCount > intI Then
            adoTaie.RollbackTrans
            MsgBox "預留單號不足！", vbCritical
            Exit Function
         End If
      Else
         adoTaie.RollbackTrans
         MsgBox "預留單號鎖定失敗！", vbCritical
         Exit Function
      End If
   End If
   strName = ""
   'end 2023/8/9
   
   Do While adoquery.EOF = False
      strFaNo = "" & adoquery.Fields("pa75")
      strCaseName = adoquery.Fields("cp01").Value & adoquery.Fields("cp02").Value & adoquery.Fields("cp03").Value & adoquery.Fields("cp04").Value
      If IsNull(adoquery.Fields("cp16").Value) Then
         douFAmount = 0
      Else
         If Text5 <> MsgText(601) Then
            douFAmount = Format(Val(adoquery.Fields("cp16").Value) * (100 - Val(Text5)) / 100 / douExchange, FAmount)
            douDAmount = Format(Val(adoquery.Fields("cp16").Value) * (100 - Val(Text5)) / 100, DAmount)
         Else
            douFAmount = Format(Val(adoquery.Fields("cp16").Value) / douExchange, FAmount)
            douDAmount = Val(adoquery.Fields("cp16").Value)
         End If
      End If
      If strName <> strCaseName Then
         'Add by Morgan 2010/7/22
         If (adoquery("cp01") = "FCL" Or adoquery("cp01") = "CFL" Or adoquery("cp01") = "LIN") Or ((adoquery("cp01") = "P" Or adoquery("cp01") = "T") And Left(adoquery("cp12"), 1) = "F") Then
         'Modified by Lydia 2015/04/15 為了區別整批請款單,a1k32=C
           ' strA1K32 = "Y"
            strA1K32 = "C"
         Else
            strA1K32 = ""
         End If
         'end 2010/7/22
         
         'Added by Morgan 2023/8/9
         If Check1.Value = vbChecked Then
            strDebitNo = Text14.Text
            PUB_UpdRsvDN strDebitNo
            GetRsvDN
         Else
         'end 2023/8/9
         
            strSql = "update acc1r0 set a1r04 = a1r04 where a1r01 = 'X'"
            adoTaie.Execute strSql, intI
            strDebitNo = AccAutoNo(MsgText(815), 5)
            strYes = AccSaveAutoNo(MsgText(815), Right(strDebitNo, 5))
         End If
         
        'Modify By Cheng 2004/04/23
        '美金欄位取至整數位(無條件捨去)
'         adoTaie.Execute "insert into acc1k0 (a1k01, a1k02, a1k03, a1k04, a1k08, a1k09, a1k10, a1k11, a1k13, a1k14, a1k15, a1k16, a1k18, a1k27, a1k28, a1k19, a1k20, a1k21, a1k30) values " & _
'                         "('" & strDebitNo & "', " & Val(FCDate(MaskEdBox3.Text)) & ", '" & Text2 & "', '" & Text11 & "', " & douFAmount & ", " & Val(adoquery.Fields("cp17").Value) & ", " & douExchange & ", " & douDAmount & ", '" & adoquery.Fields("cp01").Value & "', '" & adoquery.Fields("cp02").Value & "', '" & adoquery.Fields("cp03").Value & "', '" & adoquery.Fields("cp04").Value & "'" & _
'                         ", '" & IIf(Combo1 <> "USD", "TWD", Combo1) & "', '" & Text9 & "', '" & Text10 & "', " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", '" & strUserNum & "', 0)"
         'Modify by Morgan 2009/12/25 代理人欄位抓查詢值
         'Modify by Morgan 2010/7/22 +a1k32
         'Modified by Morgan 2015/1/23 幣別直接用畫面的
         strSql = "insert into acc1k0 (a1k01, a1k02, a1k03, a1k04, a1k08, a1k09, a1k10, a1k11, a1k13, a1k14, a1k15, a1k16, a1k18, a1k27, a1k28, a1k19, a1k20, a1k21, a1k30,a1k32) values " & _
                         "('" & strDebitNo & "', " & Val(FCDate(MaskEdBox3.Text)) & ", '" & strFaNo & "', '" & Text11 & "', " & Fix(Val("" & douFAmount)) & ", " & Val(adoquery.Fields("cp17").Value) & ", " & douExchange & ", " & douDAmount & ", '" & adoquery.Fields("cp01").Value & "', '" & adoquery.Fields("cp02").Value & "', '" & adoquery.Fields("cp03").Value & "', '" & adoquery.Fields("cp04").Value & "'" & _
                         ", '" & Combo1 & "', '" & Text9 & "', '" & Text10 & "', " & strSrvDate(2) & ", to_char(sysdate,'hh24miss'), '" & strUserNum & "', 0,'" & strA1K32 & "')"
         adoTaie.Execute strSql, intI
         strName = strCaseName
         strDNo = strDNo & strDebitNo & ","
      Else
            'Modify By Cheng 2004/04/23
            '美金欄位取至整數位(無條件捨去)
'         adoTaie.Execute "update acc1k0 set a1k11 = a1k11 + " & douDAmount & ", a1k09 = a1k09 + " & Val(adoquery.Fields("cp17").Value) & ", a1k08 = a1k08 + " & douFAmount & " where a1k01 = '" & strDebitNo & "'"
         adoTaie.Execute "update acc1k0 set a1k11 = a1k11 + " & douDAmount & ", a1k09 = a1k09 + " & Val(adoquery.Fields("cp17").Value) & ", a1k08 = a1k08 + " & Fix(Val("" & douFAmount)) & " where a1k01 = '" & strDebitNo & "'"
            'End
      End If
      strSql = "update caseprogress set cp60 = '" & strDebitNo & "' where cp09 = '" & adoquery.Fields("cp09").Value & "'"
      adoTaie.Execute strSql, intI
      
      strSql = "insert into acc1w0 (a1w01,a1w02) values ('" & strDebitNo & "','" & adoquery.Fields("cp09").Value & "')" 'Added by Morgan 2016/7/20
      adoTaie.Execute strSql, intI
      
      strCP09 = adoquery.Fields("cp09").Value 'Added by Morgan 2014/8/15
      adoquery.MoveNext
   Loop
   adoquery.Close
   adoTaie.CommitTrans
   strDNoArray = Split(UCase(strDNo), ",")
   ProcessData = True
   Exit Function
Checking:
   ProcessData = False
   adoTaie.RollbackTrans
   If Err.Number <> 0 Then
      MsgBox MsgText(185), , MsgText(5)
   End If
End Function

'*************************************************
'  清除畫面
'
'*************************************************
Public Sub FormClear()
   Text2 = ""
   Text3 = ""
   Text7 = ""
   Text6 = ""
   Text4 = ""
   Text1 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text5 = ""
   MaskEdBox3.Mask = ""
   MaskEdBox3.Text = ""
   MaskEdBox3.Mask = DFormat
   Combo1 = ""
   Text9 = ""
   Text10 = ""
   Text11 = ""
   Text2.SetFocus
End Sub

'Add By Cheng 2003/09/25
'列印對象
Private Function GetA1K27(strCP01 As String, strCP10 As String, strFACUCode As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

GetA1K27 = ""
'系統類別
Select Case strCP01
Case "P", "CFP", "FCP"
    '若案件性質為年費
    If strCP10 = "605" Then
        'Modify by Morgan 2007/4/25 取消代理人檔的年費請款對象&年費代理人欄位
        StrSQLa = " SELECT CU105, CU106, CU96 FROM CUSTOMER WHERE CU01=SUBSTR('" & strFACUCode & "',1,8) AND CU02=SUBSTR('" & strFACUCode & "',9,1) "
        StrSQLa = StrSQLa & " Union Select FA71, '', '' From FAGENT WHERE SUBSTR('" & strFACUCode & "',1,8)=FA01 AND SUBSTR('" & strFACUCode & "',9,1)=FA02 "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            If "" & rsA.Fields(1).Value <> "" Then
                GetA1K27 = rsA.Fields(1).Value
            ElseIf "" & rsA.Fields(2).Value <> "" Then
                GetA1K27 = rsA.Fields(2).Value
            ElseIf "" & rsA.Fields(0).Value <> "" Then
                GetA1K27 = rsA.Fields(0).Value
            End If
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        Exit Function
    End If
Case "T", "FCT", "CFT", ""
    '若案件性質為延展
    If strCP10 = "102" Then
        '取得國外代理人檔的"列印對象"
        'Modify By Sindy 2011/3/8 CU106改成CU152; FA72改成FA112; CU105改成CU151; FA71改成FA111
        StrSQLa = "Select FA111, FA112, FA66 From FAGENT WHERE SUBSTR('" & strFACUCode & "',1,8)=FA01 AND SUBSTR('" & strFACUCode & "',9,1)=FA02 "
        StrSQLa = StrSQLa & " Union SELECT CU151, CU152, CU98 FROM CUSTOMER WHERE CU01=SUBSTR('" & strFACUCode & "',1,8) AND CU02=SUBSTR('" & strFACUCode & "',9,1) "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            If "" & rsA.Fields(1).Value <> "" Then
                GetA1K27 = rsA.Fields(1).Value
            ElseIf "" & rsA.Fields(2).Value <> "" Then
                GetA1K27 = rsA.Fields(2).Value
            ElseIf "" & rsA.Fields(0).Value <> "" Then
                GetA1K27 = rsA.Fields(0).Value
            End If
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        Exit Function
    End If
End Select

'取得"固定列印對象"
'Modify By Sindy 2011/3/8 CU105改成CU151; FA71改成FA111
If CheckSys(strCP01) = "2" Or CheckSys(strCP01) = "6" Then
   StrSQLa = "Select FA111 From FAGENT WHERE SUBSTR('" & strFACUCode & "',1,8)=FA01 AND SUBSTR('" & strFACUCode & "',9,1)=FA02 AND FA111 IS NOT NULL "
   StrSQLa = StrSQLa & " Union SELECT CU151 FROM CUSTOMER WHERE CU01=SUBSTR('" & strFACUCode & "',1,8) AND CU02=SUBSTR('" & strFACUCode & "',9,1) AND CU151 IS NOT NULL "
Else
   StrSQLa = "Select FA71 From FAGENT WHERE SUBSTR('" & strFACUCode & "',1,8)=FA01 AND SUBSTR('" & strFACUCode & "',9,1)=FA02 AND FA71 IS NOT NULL "
   StrSQLa = StrSQLa & " Union SELECT CU105 FROM CUSTOMER WHERE CU01=SUBSTR('" & strFACUCode & "',1,8) AND CU02=SUBSTR('" & strFACUCode & "',9,1) AND CU105 IS NOT NULL "
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetA1K27 = rsA.Fields(0).Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'Add By Cheng 2003/09/25
'請款對象
Private Function GetA1K28(strCP01 As String, strCP10 As String, strFACUCode As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

GetA1K28 = ""
'系統類別
Select Case strCP01
Case "P", "CFP", "FCP"
    '若案件性質為年費
    If strCP10 = "605" Then
        'Modify by Morgan 2007/4/25 取消代理人檔的年費請款對象&年費代理人欄位
        StrSQLa = " SELECT CU57, CU97, CU96 FROM CUSTOMER WHERE CU01=SUBSTR('" & strFACUCode & "',1,8) AND CU02=SUBSTR('" & strFACUCode & "',9,1) "
        StrSQLa = StrSQLa & " Union Select FA30, '', '' From FAGENT WHERE SUBSTR('" & strFACUCode & "',1,8)=FA01 AND SUBSTR('" & strFACUCode & "',9,1)=FA02 "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            If "" & rsA.Fields(1).Value <> "" Then
                GetA1K28 = rsA.Fields(1).Value
            ElseIf "" & rsA.Fields(2).Value <> "" Then
                GetA1K28 = rsA.Fields(2).Value
            ElseIf "" & rsA.Fields(0).Value <> "" Then
                GetA1K28 = rsA.Fields(0).Value
            End If
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        Exit Function
    End If
Case "T", "FCT", "CFT", ""
    '若案件性質為延展
    If strCP10 = "102" Then
        'Modify By Sindy 2011/3/8 CU57改成CU147; FA30改成FA107
        StrSQLa = "Select FA107, FA67, FA66 From FAGENT WHERE SUBSTR('" & strFACUCode & "',1,8)=FA01 AND SUBSTR('" & strFACUCode & "',9,1)=FA02 "
        StrSQLa = StrSQLa & " Union SELECT CU147, CU99, CU98 FROM CUSTOMER WHERE CU01=SUBSTR('" & strFACUCode & "',1,8) AND CU02=SUBSTR('" & strFACUCode & "',9,1) "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            If "" & rsA.Fields(1).Value <> "" Then
                GetA1K28 = rsA.Fields(1).Value
            ElseIf "" & rsA.Fields(2).Value <> "" Then
                GetA1K28 = rsA.Fields(2).Value
            ElseIf "" & rsA.Fields(0).Value <> "" Then
                GetA1K28 = rsA.Fields(0).Value
            End If
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        Exit Function
    End If
End Select

'取得"固定請款對象"
'Modify By Sindy 2011/3/8 CU57改成CU147; FA30改成FA107
If CheckSys(strCP01) = "2" Or CheckSys(strCP01) = "6" Then
   StrSQLa = "Select FA107 From FAGENT WHERE SUBSTR('" & strFACUCode & "',1,8)=FA01 AND SUBSTR('" & strFACUCode & "',9,1)=FA02 AND FA107 IS NOT NULL "
   StrSQLa = StrSQLa & " Union SELECT CU147 FROM CUSTOMER WHERE CU01=SUBSTR('" & strFACUCode & "',1,8) AND CU02=SUBSTR('" & strFACUCode & "',9,1) AND CU147 IS NOT NULL "
'2011/3/8 End
Else
   StrSQLa = "Select FA30 From FAGENT WHERE SUBSTR('" & strFACUCode & "',1,8)=FA01 AND SUBSTR('" & strFACUCode & "',9,1)=FA02 AND FA30 IS NOT NULL "
   StrSQLa = StrSQLa & " Union SELECT CU57 FROM CUSTOMER WHERE CU01=SUBSTR('" & strFACUCode & "',1,8) AND CU02=SUBSTR('" & strFACUCode & "',9,1) AND CU57 IS NOT NULL "
End If
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetA1K28 = rsA.Fields(0).Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

''Add By Cheng 2003/09/25
''取得D/N折扣
'Private Function GetA1L07Disc(strCP01 As String, strCP10 As String, strFACUCode As String, strDNDate As String) As String
'Dim StrSQLa As String
'Dim rsA As New ADODB.Recordset
'
'GetA1L07Disc = ""
''2006/7/4 MODIFY BY SONIA 加商標折扣欄位
''StrSQLa = "Select FA25, FA26, FA27 From Fagent Where FA01='" & Mid(strFACUCode, 1, 8) & "' And FA02='" & Mid(strFACUCode, 9, 1) & "' "
''StrSQLa = StrSQLa & " Union Select CU36, CU37, CU38 From Customer Where CU01='" & Mid(strFACUCode, 1, 8) & "' And CU02='" & Mid(strFACUCode, 9, 1) & "' "
'StrSQLa = "Select FA25, FA26, FA27, FA73, FA74, FA75 From Fagent Where FA01='" & Mid(strFACUCode, 1, 8) & "' And FA02='" & Mid(strFACUCode, 9, 1) & "' "
'StrSQLa = StrSQLa & " Union Select CU36, CU37, CU38, CU107, CU108, CU109 From Customer Where CU01='" & Mid(strFACUCode, 1, 8) & "' And CU02='" & Mid(strFACUCode, 9, 1) & "' "
''2006/7/4 END
'rsA.CursorLocation = adUseClient
'rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'If rsA.RecordCount > 0 Then
'    '系統類別
'    Select Case strCP01
'    Case "P", "CFP", "FCP"
'        Select Case strCP10
'        Case "101", "102", "103", "104", "105", "201" '申請或翻譯
'            If "" & rsA.Fields(1).Value <> "" Then GetA1L07Disc = "" & rsA.Fields(1).Value
'        Case Else '其他
'            '若無全部折扣起始日
'            If "" & rsA.Fields(2).Value <> "" Then
'                If DBDATE(rsA.Fields(2).Value) < DBDATE(strDate) Then
'                    If "" & rsA.Fields(0).Value <> "" Then GetA1L07Disc = "" & rsA.Fields(0).Value
'                End If
'            '若有全部折扣起始日
'            Else
'                If "" & rsA.Fields(0).Value <> "" Then GetA1L07Disc = "" & rsA.Fields(0).Value
'            End If
'        End Select
'    'Modify by Morgan 2008/7/1 +S,CFC--湘斕
'    Case "T", "CFT", "FCT", "TF", "S", "CFC"
'        Select Case strCP10
'        Case "101" '申請
'            If "" & rsA.Fields(4).Value <> "" Then GetA1L07Disc = "" & rsA.Fields(4).Value
'        Case Else '其他
'            '若無全部折扣起始日
'            If "" & rsA.Fields(5).Value <> "" Then
'                If DBDATE(rsA.Fields(5).Value) < DBDATE(strDate) Then
'                    If "" & rsA.Fields(4).Value <> "" Then GetA1L07Disc = "" & rsA.Fields(4).Value
'                End If
'            '若有全部折扣起始日
'            Else
'                If "" & rsA.Fields(4).Value <> "" Then GetA1L07Disc = "" & rsA.Fields(4).Value
'            End If
'        End Select
'    Case Else
'        '若無全部折扣起始日
'        If "" & rsA.Fields(2).Value <> "" Then
'            If DBDATE(rsA.Fields(2).Value) < DBDATE(strDate) Then
'                If "" & rsA.Fields(0).Value <> "" Then GetA1L07Disc = "" & rsA.Fields(0).Value
'            End If
'        '若有全部折扣起始日
'        Else
'            If "" & rsA.Fields(0).Value <> "" Then GetA1L07Disc = "" & rsA.Fields(0).Value
'        End If
'    End Select
'End If
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'
'End Function

'Add By Cheng 2003/09/25
'是否列印申請人
Private Function GetA1K04(strCP01 As String, strFACUCode As String) As String
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

GetA1K04 = ""
'Modify By Sindy 2011/3/8 CU77改成CU149; FA44改成FA109
If CheckSys(strCP01) = "2" Or CheckSys(strCP01) = "6" Then
   '取得國外代理人檔的"D/N是否列印申請人"
   StrSQLa = "SELECT FA109 FROM FAGENT WHERE FA01='" & Mid(strFACUCode, 1, 8) & "' AND FA02='" & Mid(strFACUCode, 9, 1) & "' AND FA109 IS NOT NULL "
   '取得客戶基本檔的"D/N是否列印申請人"
   StrSQLa = StrSQLa & " Union SELECT CU149 FROM CUSTOMER WHERE CU01='" & Mid(strFACUCode, 1, 8) & "' AND CU02='" & Mid(strFACUCode, 9, 1) & "' AND CU149 IS NOT NULL "
'2011/3/8 End
Else
   '取得國外代理人檔的"D/N是否列印申請人"
   StrSQLa = "SELECT FA44 FROM FAGENT WHERE FA01='" & Mid(strFACUCode, 1, 8) & "' AND FA02='" & Mid(strFACUCode, 9, 1) & "' AND FA44 IS NOT NULL "
   '取得客戶基本檔的"D/N是否列印申請人"
   StrSQLa = StrSQLa & " Union SELECT CU77 FROM CUSTOMER WHERE CU01='" & Mid(strFACUCode, 1, 8) & "' AND CU02='" & Mid(strFACUCode, 9, 1) & "' AND CU77 IS NOT NULL "
End If
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   GetA1K04 = rsA.Fields(0).Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

Private Function TxtValidate() As Boolean
   If Me.Text12.Text = "" Then
       MsgBox "請輸入系統類別!!!", vbExclamation + vbOKOnly
       Me.Text12.SetFocus
       Exit Function
   End If
   If Me.Text2.Text = "" And Me.Text7.Text = "" Then
       MsgBox "請輸入代理人或申請人!!!", vbExclamation + vbOKOnly
       Me.Text2.SetFocus
       Exit Function
   End If
   
'Removed by Morgan 2016/11/1 不同案件性質可合併請款--陳金蓮
'Modified by Morgan 2017/2/15 改回要輸案件性質
   If Me.Text4.Text = "" Then
       MsgBox "請輸入案件性質!!!", vbExclamation + vbOKOnly
       Me.Text4.SetFocus
       Exit Function
   End If
'end 2017/2/15
'end 2016/11/1
   
   If Text12 <> "FCT" Then 'Added by Morgan 2025/8/22 開放FCT案可不需有發文日即可請款--湘嫻
      If Me.MaskEdBox1.Text = "___/__/__" Or Me.MaskEdBox1.Text = "" Then
          MsgBox "請輸入發文起日!!!", vbExclamation + vbOKOnly
          Me.MaskEdBox1.SetFocus
          Exit Function
      End If
      If Me.MaskEdBox2.Text = "___/__/__" Or Me.MaskEdBox2.Text = "" Then
          MsgBox "請輸入發文止日!!!", vbExclamation + vbOKOnly
          Me.MaskEdBox2.SetFocus
          Exit Function
      End If
   End If
   
   If Me.MaskEdBox3.Text = "___/__/__" Or Me.MaskEdBox3.Text = "" Then
       MsgBox "請輸入請款日期!!!", vbExclamation + vbOKOnly
       Me.MaskEdBox3.SetFocus
       Exit Function
   End If
   
   If Me.Combo1.Text = "" Then
       MsgBox "請輸入幣別!!!", vbExclamation + vbOKOnly
       Me.Combo1.SetFocus
       Exit Function
   End If
   
   
   If Me.Text9.Text = "" Then
       MsgBox "請輸入請款對象!!!", vbExclamation + vbOKOnly
       Me.Text9.SetFocus
       Exit Function
   End If
   If Me.Text10.Text = "" Then
       MsgBox "請輸入列印對象!!!", vbExclamation + vbOKOnly
       Me.Text10.SetFocus
       Exit Function
   End If
     
   'Added by Morgan 2013/3/29
   If DBDATE(FCDate(MaskEdBox3.Text)) > strSrvDate(1) Then
      If MsgBox("請款日期大於系統日，是否確定要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
         Me.MaskEdBox3.SetFocus
         Exit Function
      End If
   End If
   'end 2013/3/29
TxtValidate = True
End Function
'Add by Morgan 2011/4/15
'設定請款相關欄位預設值
Private Sub SetDefault()

   If Me.Text12.Text = "" Then
       Exit Sub
   End If
   If Me.Text2.Text = "" And Me.Text7.Text = "" Then
       Exit Sub
   End If
   
'Removed by Morgan 2016/11/1 不同案件性質可合併請款--陳金蓮
'Modified by Morgan 2017/2/15 改回要輸案件性質
   If Me.Text4.Text = "" Then
       Exit Sub
   End If
'end 2017/2/15
'end 2016/11/1

   If Me.MaskEdBox1.Text = "___/__/__" Or Me.MaskEdBox1.Text = "" Then
       Exit Sub
   End If
   If Me.MaskEdBox2.Text = "___/__/__" Or Me.MaskEdBox2.Text = "" Then
       Exit Sub
   End If
   If Me.MaskEdBox3.Text = "___/__/__" Or Me.MaskEdBox3.Text = "" Then
       Exit Sub
   End If
   
   Dim stCon As String, adoRst As ADODB.Recordset
   Dim stA1L07 As String, stLstA1L07 As String
   Dim stA1K27 As String, stLstA1K27 As String
   Dim stA1K28 As String, stLstA1K28 As String
   Dim stA1K04 As String, stLstA1K04 As String
   Dim strDNDate As String
   Dim bCheckErr1 As Boolean, bCheckErr2 As Boolean, bCheckErr3 As Boolean, bCheckErr4 As Boolean
   
   stCon = ""
   If Text2 <> "" Then
      stCon = stCon & " and nvl(pa75,nvl(tm44,nvl(lc22,sp26)))='" & Text2 & "'"
   End If
   
   If Text7 <> "" Then
      stCon = stCon & " and nvl(pa26,nvl(tm23,nvl(lc11,sp08)))='" & Text7 & "'"
   End If
   
   If Text4 <> "" Then
      'Modified by Morgan 2017/2/15
      'Modified by Morgan 2017/2/20 改回只能輸入一個
      'stCon = stCon & " and cp10 in ( '" & Replace(Replace(Text4, " ", ""), ",", "','") & "' )"
      stCon = stCon & " and cp10='" & Text4 & "'"
      'end 2017/2/20
      'end 2017/2/15
   End If
    
   strDNDate = Val(FCDate(MaskEdBox3.Text))
   
   strExc(0) = "select cp01,cp02,cp03,cp04,cp09,cp10,pa75,pa26 from caseprogress,patent,trademark,servicepractice,lawcase" & _
      " where cp27>=" & Val(CADate(FCDate(MaskEdBox1.Text))) & _
      " and cp27<=" & Val(CADate(FCDate(MaskEdBox2.Text))) & _
      " and cp60 is null and cp01 in ('" & Join(Split(Text12, ","), "','") & "')" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
      " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04" & _
      " and lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04" & _
      "" & stCon
      
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With adoRst
      Do While Not .EOF
         If bCheckErr1 = False Then
            stA1L07 = PUB_GetA1L07Disc(.Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04"), .Fields("cp10"), strDNDate)
            'Modified by Morgan 2015/8/5 +申請人
            stA1L07 = 100 * PUB_GetDiscX("" & .Fields("pa75"), .Fields("cp01"), .Fields("cp10"), IIf(Val(stA1L07) = 0, 1, Val(stA1L07) / 100), "" & .Fields("pa26")) 'Added by Morgan 2013/5/2
         End If
         
         If bCheckErr2 = False Then
            stA1K27 = PUB_GetA1K27(.Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04"), .Fields("cp10"))
         End If
         
         If bCheckErr3 = False Then
            stA1K28 = PUB_GetA1K28(.Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04"), .Fields("cp10"))
         End If
         
         If bCheckErr4 = False Then
            'Modified by Morgan 2020/3/17 參數傳錯
            'stA1K04 = PUB_GetA1K04(.Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04"), .Fields("cp10"))
            stA1K04 = PUB_GetA1K04(.Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04"), stA1K28, .Fields("cp10"))
         End If
         
         If .AbsolutePosition > 1 Then
            If bCheckErr1 = False Then
               If stLstA1L07 <> stA1L07 Then
                  MsgBox "折扣個案有不同，請自行設定"
                  bCheckErr1 = True
               End If
            End If
            
            If bCheckErr2 = False Then
               If stLstA1K27 <> stA1K27 Then
                  MsgBox "列印對象個案有不同，請自行設定"
                  bCheckErr2 = True
               End If
            End If
            
            If bCheckErr3 = False Then
               If stLstA1K28 <> stA1K28 Then
                  MsgBox "請款對象個案有不同，請自行設定"
                  bCheckErr3 = True
               End If
            End If
            
            If bCheckErr4 = False Then
               If stLstA1K04 <> stA1K04 Then
                  MsgBox "是否列印申請人個案有不同，請自行設定"
                  bCheckErr4 = True
               End If
            End If
         End If
         stLstA1L07 = stA1L07
         stLstA1K27 = stA1K27
         stLstA1K28 = stA1K28
         stLstA1K04 = stA1K04
         .MoveNext
      Loop
      If bCheckErr1 = False Then
         Text5 = stLstA1L07
      End If
      If bCheckErr2 = False Then
         Text9 = stLstA1K27
      End If
      If bCheckErr3 = False Then
         Text10 = stLstA1K28
      End If
      If bCheckErr4 = False Then
         Text11 = stLstA1K04
      End If
      End With
   'Added by Morgan 2012/1/4 無資料先提醒--陳金蓮
   Else
      MsgBox "查無資料，請檢查輸入條件是否錯誤！"
   End If
   Set adoRst = Nothing
End Sub

'Added by Morgan 2015/1/23
Private Sub SetCurr()
   Dim arrSys() As String
   Dim strA1K18 As String
   
   '預設幣別
   If Text12 <> "" And Text10 <> "" Then
      arrSys = Split(Text12, ",")
      'Modified by Morgan 2018/4/27
      'Call PUB_GetDefaultCurrPrintType(arrSys(0), Text10, "", strA1K18)
      Call PUB_GetDefaultCurrPrintType(arrSys(0), Text10, "", strA1K18, , , , Text9)
      'end 2018/4/27
      Combo1.Text = strA1K18
   End If
   
End Sub

'Added by Morgan 2023/7/20
Private Sub SetGridHead()
   Dim ii As Integer
   FixGrid grdDataList
   With grdDataList
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(.col) = 250: .Text = "V"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignmentFixed(.col) = flexAlignCenterCenter
      .col = 1: .ColWidth(.col) = 1400: .Text = "本所案號"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(.col) = 800: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(.col) = 850: .Text = "費用"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(.col) = flexAlignRightCenter
      .col = 4: .ColWidth(.col) = 800: .Text = "規費"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(.col) = flexAlignRightCenter
      .col = 5: .ColWidth(.col) = 950: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      For ii = .col + 1 To .Cols - 1
         .ColWidth(ii) = 0
      Next
      .BackColor = &HFFC0C0
      .Visible = True
   End With
End Sub

Private Sub grdDataList_SelChange()
   Dim iRow As Integer, iCol As Integer, ii As Integer
   Dim stV As String
   
   With grdDataList
   iRow = .MouseRow
   iCol = .MouseCol
   If iRow <> 0 Then
      .row = iRow
      ClickGrid
   
   ElseIf iRow = 0 And iCol = 0 Then
      If .TextMatrix(1, 0) = "V" Then
         stV = ""
      Else
         stV = "V"
      End If
      For ii = 1 To .Rows - 1
         If .TextMatrix(ii, 0) <> stV Then
            .row = ii
            ClickGrid
         End If
      Next
   End If
   End With
End Sub

Private Sub ClickGrid()
Dim i As Integer
grdDataList.Visible = False
grdDataList.col = 0
If grdDataList.Text = "V" Then
   lblCount = Val(lblCount) - 1
     grdDataList.Text = ""
     For i = 2 To grdDataList.Cols - 1
          grdDataList.col = i
          grdDataList.CellBackColor = QBColor(15)
    Next i
Else
   lblCount = Val(lblCount) + 1
     grdDataList.Text = "V"
     For i = 2 To grdDataList.Cols - 1
         grdDataList.col = i
         grdDataList.CellBackColor = &HFFC0C0
     Next i
End If
grdDataList.Visible = True
End Sub

'Added by Morgan 2023/8/9
'檢查是否不同規費的收文需重新計算金額
Private Function ChkDifFee() As Boolean

End Function
