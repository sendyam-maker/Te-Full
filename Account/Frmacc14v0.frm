VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc14v0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "國外請款單產生國內收據"
   ClientHeight    =   5610
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   9030
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   9030
   Begin VB.CheckBox Check1 
      Caption         =   "舊收據(點陣)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   7380
      TabIndex        =   37
      Top             =   2250
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3810
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1110
      Width           =   405
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4230
      TabIndex        =   5
      Top             =   1110
      Width           =   4335
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0FFC0&
      Caption         =   "修改存檔"
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
      Left            =   7440
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   120
      Width           =   1300
   End
   Begin VB.TextBox txtA1V09 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1350
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1143
      Width           =   800
   End
   Begin VB.TextBox txtA1K01 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1350
      MaxLength       =   15
      TabIndex        =   0
      Top             =   90
      Width           =   1440
   End
   Begin VB.CommandButton cmdWord 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Word(&W)"
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
      Left            =   6210
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   120
      Width           =   1155
   End
   Begin VB.CommandButton Command3 
      Height          =   330
      Left            =   2820
      Picture         =   "Frmacc14v0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   90
      Width           =   350
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  '平面
      Height          =   375
      Left            =   7290
      TabIndex        =   6
      Text            =   "txtInput"
      Top             =   1680
      Width           =   1635
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd1 
      Height          =   2145
      Left            =   390
      TabIndex        =   7
      Top             =   2100
      Width           =   6945
      _ExtentX        =   12241
      _ExtentY        =   3792
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
      FormatString    =   "請款項目|服務費|代收規費"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.TextBox txtA1K35 
      Height          =   330
      Left            =   1350
      TabIndex        =   2
      Top             =   750
      Width           =   7470
      VariousPropertyBits=   679493659
      BackColor       =   16777215
      MaxLength       =   100
      Size            =   "13176;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblA1K28 
      Height          =   255
      Left            =   1350
      TabIndex        =   39
      Top             =   480
      Width           =   7440
      BackColor       =   -2147483637
      VariousPropertyBits=   19
      Size            =   "13123;450"
      BorderStyle     =   1
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblCP13 
      Height          =   255
      Left            =   1350
      TabIndex        =   38
      Top             =   1770
      Width           =   1950
      BackColor       =   -2147483637
      VariousPropertyBits=   19
      Size            =   "3440;450"
      BorderStyle     =   1
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收據公司別："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   2550
      TabIndex        =   36
      Top             =   1140
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "扣繳　年度："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   120
      TabIndex        =   35
      Top             =   1140
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "折讓外幣金額："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   4860
      TabIndex        =   34
      Top             =   5070
      Width           =   1470
   End
   Begin VB.Label Lbldiscount 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   6570
      TabIndex        =   33
      Top             =   5070
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "折讓台幣金額："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   4860
      TabIndex        =   32
      Top             =   4830
      Width           =   1470
   End
   Begin VB.Label LblNTdiscount 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   6570
      TabIndex        =   31
      Top             =   4830
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "外幣金額："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   3570
      TabIndex        =   30
      Top             =   1800
      Width           =   1050
   End
   Begin VB.Label LblMoney 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4650
      TabIndex        =   29
      Top             =   1785
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "小計規費："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   3510
      TabIndex        =   28
      Top             =   4290
      Width           =   1050
   End
   Begin VB.Label LblSub_2 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4590
      TabIndex        =   27
      Top             =   4290
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "請款　對象："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   120
      TabIndex        =   26
      Top             =   480
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "請款單抬頭："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   25
      Top             =   810
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "智權　人員："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   53
      Left            =   120
      TabIndex        =   24
      Top             =   1800
      Width           =   1230
   End
   Begin VB.Label LblCaseName 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4650
      TabIndex        =   23
      Top             =   1470
      Width           =   4260
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "請款單編號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   22
      Top             =   150
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "案件名稱："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   3570
      TabIndex        =   21
      Top             =   1470
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "本所　案號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   20
      Top             =   1470
      Width           =   1260
   End
   Begin VB.Label LblCaseNo 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1350
      TabIndex        =   19
      Top             =   1524
      Width           =   1980
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "注意：在產生國內收據時，不要使用Word！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   210
      TabIndex        =   18
      Top             =   5310
      Width           =   4935
   End
   Begin VB.Label LblA1K02 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4650
      TabIndex        =   17
      Top             =   150
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "請款單日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   3360
      TabIndex        =   16
      Top             =   150
      Width           =   1260
   End
   Begin VB.Label LblSub_1 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1740
      TabIndex        =   15
      Top             =   4320
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "小計服務費："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   450
      TabIndex        =   14
      Top             =   4320
      Width           =   1260
   End
   Begin VB.Label LblTot 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1740
      TabIndex        =   13
      Top             =   4590
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "總　　　計："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   450
      TabIndex        =   12
      Top             =   4590
      Width           =   1260
   End
   Begin VB.Label LblNTAmt 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   6570
      TabIndex        =   11
      Top             =   4590
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "請款單台幣金額："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   4860
      TabIndex        =   10
      Top             =   4590
      Width           =   1680
   End
End
Attribute VB_Name = "Frmacc14v0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/09 改成Form2.0 ; Grd1改字型=新細明體-ExtB、LblCP13、LblA1K28、txtA1K35
'Create By Sindy 2015/4/13
Option Explicit

Public adoacc0k0 As New ADODB.Recordset
Dim m_dftColor As Long '預設顏色
Dim m_dftColor2 As Long '預設顏色2
Dim m_dftColor3 As Long '點選顏色
Dim iRow As Integer, iCol As Integer
Dim m_PA01 As String, m_PA02 As String, m_PA03 As String, m_PA04 As String
Dim m_PADate As String, m_PANo As String, m_PAKind As String
Dim m_FileName As String
Dim m_FileName2 As String 'Add By Sindy 2020/8/25
Dim m_FileName3 As String 'Add By Sindy 2020/12/16
Dim m_Nation As String, m_A1K02 As String
Dim m_CP05 As String '接洽日期
Dim m_AttachPath As String 'Add By Sindy 2025/9/18


Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
   
   TxtValidate = False
   
   If Trim(txtA1K35) = "" Then
      MsgBox "請輸入請款單抬頭！"
      txtA1K35.SetFocus
      Exit Function
   End If
   
   'Add By Sindy 2020/5/11
   If Trim(Text1) = "" Then
      MsgBox "請輸入收據公司別！"
      Text1.SetFocus
      Exit Function
   Else
      Call Text1_Validate(False) 'Add By Sindy 2020/12/16
   End If
   '2020/5/11 END
   
   If LblTot <> LblNTAmt Then
      MsgBox "請款項目的「總計」要等於「請款單台幣金額」！"
      Exit Function
   End If
   
   Cancel = False
   Call txtA1K35_Validate(Cancel)
   If Cancel = True Then
      Exit Function
   End If
   
   'Add By Sindy 2017/5/15 檢查收據抬頭是否存在
   If txtA1K35.Tag <> txtA1K35.Text Then
      Call PUB_ChkTitleNmExist(txtA1K35.Text)
   End If
   '2017/5/15 END
   
   'add by sonia 2019/5/29
   If Val(txtA1V09.Tag) > 0 And Val(txtA1V09.Text) = 0 Then
      MsgBox "扣繳年度不可刪除！將會還原原扣繳年度！"
      txtA1V09.Text = Val(txtA1V09.Tag)
      Exit Function
   End If
   'end 2019/5/29
   
    'Added by Lydia 2021/12/09 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
         Exit Function
    End If
    'end 2021/12/09
    
   TxtValidate = True
End Function

Private Sub JCallWordPrint()
Dim strNo As String, strFileName As String, strA0k02 As String
Dim i As Integer, jj As Integer
Dim strName As String
Dim strText As String
Dim dblSerFee As Double, dblFee As Double
Dim dblTotSerFee As Double, dblTotFee As Double
Dim intFeeItem As Integer
Dim m_CU11 As String, m_CU173 As String
Dim strLOS02 As String, strLOScp12 As String, strLOScp13 As String
Dim StrSQLa As String, intR As Integer
   
On Error GoTo ErrHand
   
   'Add By Sindy 2017/5/15
   Call GetTitleCustData(txtA1K35, "", "", , , _
                            , , , , , , , _
                            , , , , , , , , , , , , , , , m_CU11, , , m_CU173)
   '2017/5/15 END
   
   '判斷word是否已開啟
   If g_WordAp Is Nothing Then
RestarWord:
      Set g_WordAp = New Word.Application
      g_WordAp.Visible = False
   End If
   
   'Modify By Sindy 2025/9/18
   'strFileName = "$$" & txtA1K01 & "-" & strSrvDate(1) & "-" & ServerTime & ".doc"
   strFileName = "$$" & strSrvDate(2) & ServerTime & "_J_" & txtA1K01 & ".doc"
   '2025/9/18 END
   If Dir(m_AttachPath & strFileName) <> "" Then
      Kill m_AttachPath & strFileName
   End If
   
   'Add By Sindy 2020/12/16
   If Text1 = "L" Then
      g_WordAp.Documents.Open m_AttachPath & m_FileName3
      
      'Add By Sindy 2020/12/16 抓案源資料:
      '以收據號抓ACC0J0之收文號，再抓法律所案源檔LAWOFFICESOURCE之法律所總收文號LOS06欄，
      '再抓案件進度檔等相關法務資料
      '例:E10910184
      StrSQLa = "SELECT cp01,cp02,cp03,cp04,cp12,cp13,LOS02,LC11,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as CUname,cp66,cp67" & _
                " From LawOfficeSource, CaseProgress, acc1w0, lawcase, customer" & _
                " WHERE a1w01='" & txtA1K01 & "' AND LOS06(+)=a1w02" & _
                " AND LOS06=cp09(+)" & _
                " AND cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND LC01 is not null" & _
                " AND substr(LC11,1,8)=cu01(+) AND substr(LC11,9,1)=cu02(+)" & _
                " union all " & _
                "SELECT cp01,cp02,cp03,cp04,cp12,cp13,LOS02,hc05,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as CUname,cp66,cp67" & _
                " From LawOfficeSource, CaseProgress, acc1w0, hirecase, customer" & _
                " WHERE a1w01='" & txtA1K01 & "' AND LOS06(+)=a1w02" & _
                " AND LOS06=cp09(+)" & _
                " AND cp01=HC01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND cp04=HC04(+) AND HC01 is not null" & _
                " AND substr(hc05,1,8)=cu01(+) AND substr(hc05,9,1)=cu02(+)" & _
                " order by cp66 asc,cp67 asc"
      intR = 1
      Set RsTemp = ClsLawReadRstMsg(intR, StrSQLa)
       If RsTemp.RecordCount > 0 Then
         strLOS02 = "" & RsTemp.Fields("LOS02").Value
         'strLOSCName = "" & RsTemp.Fields("CUname").Value
         '以收據號抓ACC0J0之收文號，再抓法律所案源檔LAWOFFICESOURCE之法律所總收文號LOS06欄，
         '抓出TT總收文號(收據)LOS10，再抓案件進度檔的CP12及CP13。
         '例:E10909604
         StrSQLa = "SELECT cp12,cp13,cp66,cp67" & _
                   " From LawOfficeSource, CaseProgress, acc1w0" & _
                   " WHERE a1w01='" & txtA1K01 & "' AND LOS06(+)=a1w02" & _
                   " AND LOS10=cp09(+) AND LOS10 is not null" & _
                   " order by cp66 asc,cp67 asc"
         intR = 1
         Set RsTemp = ClsLawReadRstMsg(intR, StrSQLa)
         If RsTemp.RecordCount > 0 Then
            strLOScp12 = "" & RsTemp.Fields("cp12").Value
            strLOScp13 = "" & RsTemp.Fields("cp13").Value
         End If
      End If
      '2020/4/23 END
   Else
   '2020/12/16 END
      If Check1.Value = 1 Then
         g_WordAp.Documents.Open m_AttachPath & m_FileName
      'Modify By Sindy 2020/8/25
      Else
         '新表格
         g_WordAp.Documents.Open m_AttachPath & m_FileName2
      End If
      '2020/8/25 END
   End If
   g_WordAp.ActiveDocument.SaveAs m_AttachPath & strFileName
   g_WordAp.ActiveDocument.Close
   g_WordAp.Documents.Open m_AttachPath & strFileName
   
   'Modify By Sindy 2020/12/16 + 法律所收據
   With g_WordAp
      .Selection.WholeStory
      .Selection.Copy
      For i = 1 To 16 '14 '13 '10
         strName = ""
         strText = ""
         If i = 1 Then
            strName = "收據日期"
            strText = Left(m_A1K02, 4) - 1911 & "年" & Mid(m_A1K02, 5, 2) & "月" & Right(m_A1K02, 2) & "日"
            If Check1.Value = 1 Then  '舊收據(點陣)
               strText = Left(m_A1K02, 4) - 1911 & "     " & Mid(m_A1K02, 5, 2) & "     " & Right(m_A1K02, 2)
'               'Modify By Sindy 2020/8/25
'               Else
'                  '新表格
'                  strText = Left(m_A1K02, 4) - 1911 & "年" & Mid(m_A1K02, 5, 2) & "月" & Right(m_A1K02, 2) & "日"
            End If
            '2020/8/25 END
         ElseIf i = 2 Then
            strName = "收據號碼"
            strText = txtA1K01
         ElseIf i = 3 Then
            If Text1 = "L" Then
               strName = "客戶資料"
            Else
               strName = "請款單抬頭"
            End If
            strText = txtA1K35 & IIf(m_CU173 = "Y", m_CU11, "")
         ElseIf i = 4 And Text1 <> "L" Then
            strName = "地址"
            strText = ""
         ElseIf i = 5 Then
            strName = "智權人員"
            '業務區+智權人員代號
            strText = LblCP13.Tag & "　" & Left(LblCP13, 5)
            
            'Add By Sindy 2020/4/23 加印介紹人之部門及員工編號
            '例:E10909604　L02 82021 / S22 A3023
            If strLOS02 <> "" Then
               'Modify By Sindy 2020/4/30 例:E10910184 L02 88028
               '不存在法律所案源檔或案源案件類型LOS02為C類時則不加印，也不要斜線
               If strLOS02 <> "C" Then
                  strText = strText & " / " & strLOScp12 & " " & strLOScp13
               End If
            End If
            '2020/4/23 END
         ElseIf i = 6 And Text1 <> "L" Then
            If Check1.Value = 1 Then
               strName = "本所案號"
               strText = Replace(lblCaseNo, "-0-00", "")
            'Modify By Sindy 2020/8/25
            Else
               '新表格
               strName = "備註1"
               strText = "本所案號：" & Replace(lblCaseNo, "-0-00", "") & "    案件名稱：" & LblCaseName
            End If
            '2020/8/25 END
         'Modify By Sindy 2020/8/25
         'ElseIf i = 7 Then
         ElseIf i = 7 Then
            If Text1 = "L" Then
               strName = "補開"
               strText = ""
            ElseIf Check1.Value = 1 Then
               strName = "案件名稱"
               strText = LblCaseName
            End If
         ElseIf i = 8 Then
            If Text1 = "L" Then
               strName = "服務小計"
            Else
               strName = "小計服務費"
            End If
            strText = "NTD " & Format(LblSub_1, DDollar)
         ElseIf i = 9 Then
            If Text1 = "L" Then
               strName = "規費小計"
            Else
               strName = "小計規費"
            End If
            strText = "NTD " & IIf(LblSub_2 = 0, "0", Format(LblSub_2, DDollar))
         ElseIf i = 10 Then
            If Text1 = "L" Then
               strName = "總計"
            Else
               strName = "總金額"
            End If
            strText = "NTD " & Format(LblTot, DDollar)
         ElseIf i = 11 Then
            If Text1 = "L" Then
               strName = "客戶案號"
            Else
               strName = "客戶案件案號"
            End If
            strText = ""
         ElseIf i = 12 Then
            If Text1 = "L" Then
               strName = "戶名"
               strText = A0802Query("L")
            ElseIf Check1.Value = 0 Then
               strName = "備註2"
               strText = ""
            End If
         ElseIf i = 13 Then
            If Text1 = "L" Then
               strName = "支票抬頭"
               strText = A0802Query("L")
            Else
               strName = "客戶代號"
               strText = ""
            End If
         ElseIf i = 14 And Text1 = "L" Then
            strName = "總計新台幣"
            strText = ChangeNumber(LblTot)
         'Add By Sindy 2022/7/20
         ElseIf i = 15 And Text1 = "2" Then
            strName = "補開"
            strText = ""
         ElseIf i = 16 And Text1 = "2" Then
            strName = "列印次數"
            strText = ""
         '2022/7/20 END
         End If
         If Trim(strName) <> "" Then
            .Selection.Find.ClearFormatting
            .Selection.Find.Text = "|#" & strName & "#|"
            .Selection.Find.Replacement.Text = ""
            .Selection.Find.Forward = True
            .Selection.Find.Wrap = wdFindContinue
            .Selection.Find.Format = False
            .Selection.Find.MatchCase = False
            .Selection.Find.MatchWholeWord = False
            .Selection.Find.MatchWildcards = False
            .Selection.Find.MatchSoundsLike = False
            .Selection.Find.MatchAllWordForms = False
            .Selection.Find.MatchByte = True
            .Selection.Find.Execute
            .Selection.Delete
            .Selection.Font.ColorIndex = wdBlack
            .Selection.TypeText strText
         End If
      Next i
      For i = 1 To 6 '4 '3
         'Modify By Sindy 2020/12/16 + 法律所收據
         If Text1 = "L" Then
            intFeeItem = 4
         Else
         '2020/12/16 END
            intFeeItem = 3
         End If
         For jj = 1 To intFeeItem '3
            strName = ""
            strText = ""
'            If jj = 1 Then
'               strName = "日期" & IIf(i > 1, i, "")
''               strText = ChangeWStringToTDateString(m_CP05)
'            Else
            If jj = 1 Then
               strName = IIf(Text1 = "L", "委辦事項" & i, "帳款類別" & IIf(i > 1, i, ""))
               If GRD1.Rows - 1 < i Then
                  strText = ""
               Else
                  strText = GRD1.TextMatrix(i, 1)
               End If
            ElseIf jj = 2 Then
               strName = IIf(Text1 = "L", "服務費" & i, "服務費" & IIf(i > 1, i, ""))
               If GRD1.Rows - 1 < i Then
                  strText = ""
               Else
                  strText = IIf(GRD1.TextMatrix(i, 2) = 0, 0, Format(GRD1.TextMatrix(i, 2), DDollar))
               End If
            ElseIf jj = 3 Then
               strName = IIf(Text1 = "L", "規費" & i, "規費" & IIf(i > 1, i, ""))
               If GRD1.Rows - 1 < i Then
                  strText = ""
               Else
                  strText = IIf(GRD1.TextMatrix(i, 3) = 0, 0, Format(GRD1.TextMatrix(i, 3), DDollar))
               End If
            ElseIf jj = 4 Then '法律所收據使用
               strName = "本所案號" & i
               If i = 1 Then
                  strText = Replace(lblCaseNo, "-0-00", "")
               End If
            End If
            If Trim(strName) <> "" Then
               .Selection.Find.ClearFormatting
               .Selection.Find.Text = "|#" & strName & "#|"
               .Selection.Find.Replacement.Text = ""
               .Selection.Find.Forward = True
               .Selection.Find.Wrap = wdFindContinue
               .Selection.Find.Format = False
               .Selection.Find.MatchCase = False
               .Selection.Find.MatchWholeWord = False
               .Selection.Find.MatchWildcards = False
               .Selection.Find.MatchSoundsLike = False
               .Selection.Find.MatchAllWordForms = False
               .Selection.Find.MatchByte = True
               .Selection.Find.Execute
               .Selection.Delete
               .Selection.Font.ColorIndex = wdBlack
               .Selection.TypeText strText
            End If
         Next jj
      Next i
      .ActiveDocument.Save 'Add By Sindy 2025/9/18
   End With
   
   g_WordAp.Visible = True
   Set g_WordAp = Nothing
   MsgBox "已產生完畢！"
   Exit Sub
   
ErrHand:
   If Err.Number = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   ElseIf Err.Number <> 0 Then
      MsgBox (Err.Description)
      If Not g_WordAp Is Nothing Then
         g_WordAp.Quit
         Set g_WordAp = Nothing
      End If
   End If
End Sub

'列印(&P)
'Word(&W)
Private Sub cmdWord_Click()
'Add By Sindy 2015/11/4
Dim strA1k29 As String, strA0y02 As String
Dim strCaseProperty As String, strNation As String, strCP09 As String
Dim dblA1V04 As Double, dblA1V06 As Double, dblA1V07 As Double
'2015/11/4 END
Dim strSqlMain As String
Dim strNo As String
   
   If TxtValidate Then
      Screen.MousePointer = vbHourglass
      
On Error GoTo ErrHand
      
      txtA1K35 = Trim(txtA1K35) 'Add By Sindy 2016/3/16 因瑞婷在輸資料時,出現下列狀況抬頭多了空白,導至在Run扣繳憑單維護或查詢時,資料有少
      '|'||A1K35||'|'                                                                  LENGTH(A1K A1K01           A1K02      A1K03     A A1K05                                                                            A1K06      A1K07      A1K08      A1K09      A1K10      A1K11      A1K12      A1K A1K14  A A1 A1K17           A1K1 A1K25           A1K26           A1K27     A1K28     A1K21  A1K19      A1K20      A1K24  A1K22      A1K23      A A1K30      A1K31      A A A1K34                                                                            A1K35
      '-------------------------------------------------------------------------------- ---------- --------------- ---------- --------- - -------------------------------------------------------------------------------- ---------- ---------- ---------- ---------- ---------- ---------- ---------- --- ------ - -- --------------- ---- --------------- --------------- --------- --------- ------ ---------- ---------- ------ ---------- ---------- - ---------- ---------- - - -------------------------------------------------------------------------------- --------------------------------------------------------------------------------
      '|康銀流通股份有限公司|                                                                   10 X10309490          1030624 Y34073000                                                                                                                1250          0         28      35000            FCP 039919 0 00                 USD                                  Y34073000 Y34073000 82045     1030624     151127 82045     1030624     151149 Y      35000            Y 2 105/03/16由財務處以國內收據格式產生Word檔;                                       康銀流通股份有限公司
      '|康銀流通股份有限公司|                                                                   10 X10406050          1040428 Y34073000                                                                                                                2217          0         29      64300            FCP 039919 0 00                 USD                                  Y34073000 Y34073000 82045     1040428     164628 82045     1040428     164710 Y      64300            Y 2 105/03/16由財務處以國內收據格式產生Word檔;                                       康銀流通股份有限公司
      '|康銀流通股份有限公司  |                                                                 12 X10408539          1040604 Y34073000                                                                                                                 149       1700         28       4200            FCP 039919 0 00                 USD                                  Y34073000 Y34073000 82045     1040604     163839 82045     1040604     164319 Y       4200            Y 2 105/03/16由財務處以國內收據格式產生Word檔;                                       康銀流通股份有限公司
      '|康銀流通股份有限公司|                                                                   10 X10410947          1040722 Y34073000                                                                                                                 359          0       28.5      10250            FCP 039919 0 00                 USD                                  Y34073000 Y34073000 85033     1040722     152605 85033     1040722     152641 Y      10250            Y 2 105/03/16由財務處以國內收據格式產生Word檔;                                       康銀流通股份有限公司
      '|康銀流通股份有限公司 |                                                                  11 X10410948          1040722 Y34073000                                                                                                                 359          0       28.5      10250            FCP 039919 0 00                 USD                                  Y34073000 Y34073000 85033     1040722     152757 85033     1040722     152820 Y      10250            Y 2 105/03/16由財務處以國內收據格式產生Word檔;                                       康銀流通股份有限公司
      
      '******
      ' 存檔
      '******
      'Modify By Sindy 2020/5/11 + ,a1k37='" & Text1 & "'
      strSql = "update acc1k0 set a1k32='Y',a1k35='" & txtA1K35 & "',a1k37='" & Text1 & "'" & _
               " where a1k01='" & txtA1K01 & "'"
      cnnConnection.Execute strSql
      strExc(0) = "select a1k01 from acc1k0" & _
                  " where a1k01='" & txtA1K01 & "'" & _
                  " and instr(a1k34,'" & ChangeWStringToTDateString(strSrvDate(1)) & "由財務處以國內收據格式產生Word檔" & "')>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 0 Then
         strSql = "update acc1k0 set a1k34='" & ChangeWStringToTDateString(strSrvDate(1)) & "由財務處以國內收據格式產生Word檔;'||a1k34" & _
                  " where a1k01='" & txtA1K01 & "'"
         cnnConnection.Execute strSql
      End If
      
      'Add By Sindy 2015/11/4 補資料至ACC1V0
      strExc(0) = "select a1v02 from acc1v0 where a1v02='" & txtA1K01 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 0 Then
         '已收款(要全結清的) : A1k30-規費(A1k09)
'         strExc(0) = "select sum(a0z04 * a0y04) as Namount,a1k29" & _
'                     " From acc0z0, acc0y0, acc1k0" & _
'                     " where a0z02='" & txtA1K01 & "' and a0z01=a0y01(+) and a0z01=a1k01(+)" & _
'                     " group by a1k29"
         'Modify By Sindy 2016/1/20 先檢查a0z12扣繳金額是否有金額,若有,以a0z12金額為主寫入acc1v0
         '                          否則才讀acc1k0資料
         strExc(0) = "select sum(a0z12) a0z12" & _
                     " From acc0z0" & _
                     " where a0z02='" & txtA1K01 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Val("" & RsTemp.Fields("a0z12")) > 0 Then
               dblA1V04 = RsTemp.Fields("a0z12")
               dblA1V06 = RsTemp.Fields("a0z12")
               dblA1V07 = 0
            End If
         End If
         '2016/1/20 END
         strExc(0) = "select nvl(A1k30,0)-nvl(A1k09,0) as Namount,a1k29" & _
                     " From acc1k0" & _
                     " where a1k01='" & txtA1K01 & "' and a1k29='Y'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            'Modify By Sindy 2016/1/20
            'dblAmt = Val("" & RsTemp.Fields("Namount"))
            If dblA1V04 = 0 Then
               dblA1V04 = Val("" & RsTemp.Fields("Namount")) / 10
               dblA1V06 = 0
               dblA1V07 = Val("" & RsTemp.Fields("Namount")) / 10
            End If
            '2016/1/20 END
            
            strA1k29 = "" & RsTemp.Fields("a1k29")
            'Modifie by Lydia 2016/03/25
            'strExc(0) = "select max(a0y02) From acc0z0, acc0y0 where a0z02='" & txtA1K01 & "' and a0z01=a0y01(+)"
            'intI = 1
            'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            'If intI = 1 Then
            '   strA0y02 = RsTemp.Fields(0)
            'End If
            strA0y02 = txtA1V09.Text
            
            '設定分錄欄位預設值
            strExc(0) = "select CP09, DECODE(PA09,'000',CPM03,CPM04) as Property, (cu01||cu02) as CustNo, pa09 as nation from caseprogress, salesno, staff, casepropertyMap, patent, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and substr(pa26, 1, 8) = cu01 (+) and substr(pa26, 9, 1) = cu02 (+) and cp60 = '" & txtA1K01 & "' union " & _
               "select CP09, DECODE(TM10,'000',CPM03,CPM04) as Property, (cu01||cu02) as CustNo, tm10 as nation from caseprogress, salesno, staff, casepropertyMap, trademark, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and substr(tm23, 1, 8) = cu01 (+) and substr(tm23, 9, 1) = cu02 (+) and cp60 = '" & txtA1K01 & "' union " & _
               "select CP09, DECODE(LC15,'000',CPM03,CPM04) as Property, (cu01||cu02) as CustNo, lc15 as nation from caseprogress, salesno, staff, casepropertyMap, lawcase, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = lc01 and cp02 = lc02 and cp03 = lc03 and cp04 = lc04 and substr(lc11, 1, 8) = cu01 (+) and substr(lc11, 9, 1) = cu02 (+) and cp60 = '" & txtA1K01 & "' union " & _
               "select CP09, nvl(cpm03, cpm04) as Property, (cu01||cu02) as CustNo, null as nation from caseprogress, salesno, staff, casepropertyMap, hirecase, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = hc01 and cp02 = hc02 and cp03 = hc03 and cp04 = hc04 and substr(hc05, 1, 8) = cu01 (+) and substr(hc05, 9, 1) = cu02 (+) and cp60 = '" & txtA1K01 & "' union " & _
               "select CP09, DECODE(SP09,'000',CPM03,CPM04) as Property, (cu01||cu02) as CustNo, sp09 as nation from caseprogress, salesno, staff, casepropertyMap, servicepractice, customer where cp13 = sn02 (+) and cp13 = st01 (+) and cp01 = cpm01 (+) and cp10 = cpm02 (+) and cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and substr(sp08, 1, 8) = cu01 (+) and substr(sp08, 9, 1) = cu02 (+) and cp60 = '" & txtA1K01 & "' order by cp09 asc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strCustNo = "" & RsTemp.Fields("Custno").Value
               strCaseProperty = "" & RsTemp.Fields("Property").Value
               strNation = "" & RsTemp.Fields("nation").Value
               strCP09 = "" & RsTemp.Fields("CP09").Value
            End If
            'Modified by Lydia 2016/03/25 畫面輸入扣繳年度
            'strSql = "insert into ACC1V0 (a1v01,a1v02,a1v03,a1v04,a1v05,a1v06,a1v07,a1v09,a1v18,a1v12,a1v13)" & _
                     " values('" & strCP09 & "','" & txtA1K01 & "',GetA0k11('" & strCP09 & "')" & _
                     "," & dblA1V04 & ",'" & IIf(strA1k29 = "", "Y", "N") & "'," & dblA1V06 & "," & dblA1V07 & _
                     "," & Left(strA0y02, Len(strA0y02) - 4) & ",'1','" & strCaseProperty & "','" & IIf(strNation = "", "臺灣", GetPrjNationName(strNation)) & "')"
            'Modify By Sindy 2020/5/11 GetA0k11('" & strCP09 & "') => '" & Text1 & "'
            strSql = "insert into ACC1V0 (a1v01,a1v02,a1v03,a1v04,a1v05,a1v06,a1v07,a1v09,a1v18,a1v12,a1v13)" & _
                     " values('" & strCP09 & "','" & txtA1K01 & "','" & Text1 & "'" & _
                     "," & dblA1V04 & ",'" & IIf(strA1k29 = "", "Y", "N") & "'," & dblA1V06 & "," & dblA1V07 & _
                     "," & CNULL(txtA1V09.Text, True) & ",'1','" & strCaseProperty & "','" & IIf(strNation = "", "臺灣", GetPrjNationName(strNation)) & "')"
                    
            cnnConnection.Execute strSql
         End If
      'Added by Lydia 2016/03/25 畫面輸入扣繳年度
      Else
         'Modify By Sindy 2022/1/21 + Or Text1.Tag <> Text1.Text
         'If txtA1V09.Tag <> txtA1V09.Text Or Text1.Tag <> Text1.Text Then
            'Modify By Sindy 2020/5/11 + ,a1v03='" & Text1 & "'
            strSql = " update acc1v0 set a1v09=" & CNULL(txtA1V09.Text, True) & ",a1v03='" & Text1 & "' where a1v02='" & txtA1K01 & "' "
            cnnConnection.Execute strSql
         'End If
      'end 2016/03/25
      End If
      '2015/11/3 END
      
      '產生Word檔
      Call JCallWordPrint
      
'      'Modify By Sindy 2020/6/2 改同收據的套印
'      strSqlMain = "select * from ( " & _
'                   "select * from acc0k0, customer " & _
'                   "where substr(a0k03, 1, 8) = cu01 (+) and substr(a0k03, 9, 1) = cu02 (+) " & _
'                   "and to_number(substr(a0k01, 5, 5)) > 2000 " & _
'                   "and a0k19 = 0 and (a0k09 is null or a0k09 = 0) and (a0k37 is null or a0k37<>'N') " & _
'                   "and a0k32 IS NULL and a0k11<>'J' " & strSql & _
'                   " Union All " & _
'                   "select * from acc0k0, customer " & _
'                   "where substr(a0k03, 1, 8) = cu01 (+) and substr(a0k03, 9, 1) = cu02 (+) " & _
'                   "and to_number(substr(a0k01, 5, 5)) > 2000 " & _
'                   "and a0k19 = 0 and (a0k09 is null or a0k09 = 0) and (a0k37 is null or a0k37<>'N') " & _
'                   "and (a0k32 ='Y') and a0k11 = '" & Text1 & "'" & _
'                   ") order by a0k01 asc "
'
'      adoacc0k0.CursorLocation = adUseClient
'      '收據餘額檔(一張收據一筆資料)
'      adoacc0k0.Open strSqlMain, adoTaie, adOpenStatic, adLockReadOnly
'      If adoacc0k0.RecordCount = 0 Then
'         adoacc0k0.Close
'         Screen.MousePointer = vbDefault
'         MsgBox MsgText(28), , MsgText(5)
'         Exit Sub
'      End If
'      '初始化記錄收據號碼的變數
'      strNo = ""
'      Printer.EndDoc
'      Printer.PaperSize = PUB_GetPaperSize(1)
'      Forms(0).StatusBar1.Panels(2).Text = Printer.PaperSize
'      Printer.FontSize = 12
'      Printer.Font = "標楷體"
'      Do While adoacc0k0.EOF = False
'         '改呼叫共用(與補印收據統一)
'         If strNo <> adoacc0k0.Fields("a0k01").Value Then
'            If strNo <> "" Then Printer.NewPage
'            PUB_PrintCaseReceipt adoacc0k0, 0, 0
'            strNo = adoacc0k0.Fields("a0k01").Value
'         End If
'         adoacc0k0.MoveNext
'      Loop
'      '改呼叫共用(與補印收據統一)
'      adoacc0k0.Close
'      Printer.EndDoc
'      Printer.Font = "新細明體"
      
      Screen.MousePointer = vbDefault
   End If
   Exit Sub
   
ErrHand:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox "更新資料失敗！" & vbCrLf & Err.Description
   End If
End Sub

Private Sub Command3_Click()
   Call doQuery
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Public Sub KeyDefine(KeyCode As Integer)
On Error GoTo Checking
   
   Select Case KeyCode
      Case vbKeyF12
         If FormCheck Then
            Screen.MousePointer = vbHourglass
            doQuery
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         
      Case Else
         KeyEnter KeyCode
   End Select
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
   Exit Sub
   
Checking:
   Screen.MousePointer = vbDefault
   MsgBox Err.Description, , MsgBox(5)
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   FormCheck = False
   
   If Trim(txtA1K01) = "" Then
      MsgBox "請款單編號不可空白！", , MsgText(5)
      txtA1K01.SetFocus
      Exit Function
   End If
   
   FormCheck = True
End Function

Private Sub doQuery()
Dim AdoRs As New ADODB.Recordset
Dim strCustomerName As String
Dim ii As Integer
Dim strCon As String 'Add By Sindy 2020/4/28
   
   GRD1.Clear
   LblA1K02.Caption = ""
   LblA1K28.Caption = ""
   txtA1K35.Text = ""
   lblCaseNo.Caption = ""
   LblCaseName.Caption = ""
   LblCP13.Caption = ""
   LblCP13.Tag = "" 'Add By Sindy 2020/11/15
   txtA1V09.Text = "" 'Added by Lydia 2016/03/25
   
   GridHead
   cmdWord.Enabled = False
   cmdSave.Enabled = False 'Added by Lydia 2016/03/25
   
   Screen.MousePointer = vbHourglass
   
   '抓請款單資料:必須未作廢未銷帳
   'Modified by Lydia 2016/03/25 + a1v09
   'Modify by Sindy 2020/12/9 + ,a1k37
   'strExc(0) = "select a1k01,sqldatet(a1k02) a1k02,a1k06,a1k08,a1k28,a1k35,a1k13,a1k14,a1k15,a1k16,a1k18,a1k11,a1k31 from acc1k0 where a1k01='" & txtA1K01 & "' and a1k12 is null and a1k25 is null"
   strExc(0) = "select a1k01,sqldatet(a1k02) a1k02,a1k06,a1k08,a1k28,a1k35,a1k37,a1k13,a1k14,a1k15,a1k16,a1k18,a1k11,a1k31,a1v09 from acc1k0,acc1v0 where a1k01='" & txtA1K01 & "' and a1k12 is null and a1k25 is null and a1k01=a1v02(+)"
   intI = 1
   Set AdoRs = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      LblA1K02.Caption = AdoRs.Fields("a1k02")
      m_A1K02 = DBDATE(AdoRs.Fields("a1k02"))
      lblCaseNo.Caption = AdoRs.Fields("a1k13") & "-" & AdoRs.Fields("a1k14") & "-" & AdoRs.Fields("a1k15") & "-" & AdoRs.Fields("a1k16")
      m_PA01 = AdoRs.Fields("a1k13")
      m_PA02 = AdoRs.Fields("a1k14")
      m_PA03 = AdoRs.Fields("a1k15")
      m_PA04 = AdoRs.Fields("a1k16")
      '請款對象
      'Modify By Sindy 2017/12/15
      'Call GetAgentAndState("" & adoRs.Fields("a1k28"), strCustomerName, , , True)
      Call GetAgentAndState("" & AdoRs.Fields("a1k28"), strCustomerName, , False, True)
      If strCustomerName = "" And Left("" & AdoRs.Fields("a1k28"), 1) = 客戶編號 Then
         strExc(0) = "select cu01||cu02,nvl(cu04,Decode(cu05,null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90))" & _
                     " from customer where " & ChgCustomer("" & AdoRs.Fields("a1k28")) & " order by cu01,cu02"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strCustomerName = "" & RsTemp.Fields(1)
         End If
      End If
      '2017/12/15 END
      LblA1K28 = strCustomerName
      '請款單抬頭
      If "" & AdoRs.Fields("a1k35") = "" Then
         txtA1K35.Text = Trim(strCustomerName)
      Else
         txtA1K35.Text = Trim("" & AdoRs.Fields("a1k35"))
      End If
      'Add By Sindy 2020/12/9
      If "" & AdoRs.Fields("a1k37") <> "" Then
         Text1.Text = Trim("" & AdoRs.Fields("a1k37"))
         Call Text1_Validate(False) 'Add By Sindy 2020/12/16
      End If
      '2020/12/9 END
      
      'Modify By Sindy 2020/4/28 查出來當下是代理人,所以不需要檢查
'      'Add By Sindy 2017/5/15 檢查收據抬頭是否存在
      txtA1K35.Tag = txtA1K35.Text
'      Call PUB_ChkTitleNmExist(txtA1K35.Text)
'      '2017/5/15 END

      '案件名稱
      LblCaseName = GetPrjName(lblCaseNo.Caption)
      '申請國家
      m_Nation = GetPrjNation1(lblCaseNo.Caption)
      '智權人員
      strExc(0) = "select cp13||' '||st02,cp05,st15 from caseprogress,staff where cp09=(" & _
                  " select min(cp09) from caseprogress where cp60='" & txtA1K01 & "'" & _
                  " and cp05=(select min(cp05) from caseprogress where cp60='" & txtA1K01 & "'))" & _
                  " and cp13=st01(+)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         LblCP13 = RsTemp.Fields(0)
         LblCP13.Tag = "" & RsTemp.Fields("st15") 'Add By Sindy 2020/11/15
         m_CP05 = RsTemp.Fields("cp05")
      End If
      '請款單台幣金額
      LblNTAmt = Val("" & AdoRs.Fields("a1k11")) - Val("" & AdoRs.Fields("a1k06"))
      '外幣金額
      LblMoney = "" & AdoRs.Fields("a1k18") & CStr(Val("" & AdoRs.Fields("a1k08")))
      '折讓台幣金額
      LblNTdiscount = Val("" & AdoRs.Fields("a1k06"))
      '折讓外幣金額
      Lbldiscount = Val("" & AdoRs.Fields("a1k31"))
      
      'Added by Lydia 2016/03/25 + 扣繳年度
      If IsNull(AdoRs.Fields("a1v09")) Then
          strExc(0) = "select max(a0y02) From acc0z0, acc0y0 where a0z02='" & txtA1K01 & "' and a0z01=a0y01(+)"
          intI = 1
          Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
          If intI = 1 Then
             'Modified by Lydia 2016/03/28 + 非null
             If Not IsNull(RsTemp.Fields(0)) Then txtA1V09 = Left(RsTemp.Fields(0), Len(RsTemp.Fields(0)) - 4)
          End If
      Else
          txtA1V09.Text = AdoRs.Fields("a1v09")
      End If
   Else
      Screen.MousePointer = vbDefault
      MsgBox "無此請款單資料！", , MsgText(5)
      Exit Sub
   End If
   
   txtA1V09.Tag = txtA1V09.Text 'Added by Lydia 2016/03/25 扣繳年度
'   Text1.Tag = Text1.Text 'Added by Sindy 2022/1/21 收據公司別
   
   '抓國外請款交易資料
   strExc(0) = "select X.item,a1J03,'','' from" & _
               " (select substr(a1L04,1,length(a1L04)-2) item,a1L03 from acc1L0 where a1L01='" & txtA1K01 & "' and substr(a1L04,-2) in('98','99')" & _
               " union select a1L04,a1L03 item from acc1L0 where a1L01='" & txtA1K01 & "' and substr(a1L04,-2) not in('98','99')) X,acc1J0" & _
               " Where x.Item = a1J02 And x.a1L03 = a1J01" & _
               " order by X.item asc"
   intI = 1
   Set GRD1.Recordset = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      cmdWord.Enabled = True
      cmdSave.Enabled = True 'Added by Lydia 2016/03/25
      For ii = 1 To GRD1.Rows - 1
         'Add By Sindy 2020/4/28 X10905271 FCT-045652 有A01.35類 要歸入101.商申
         If InStr(m_PA01, "T") > 0 And GRD1.TextMatrix(ii, 0) = "101" Then
            strCon = " IN ('101','A01')"
         Else
            strCon = " ='" & GRD1.TextMatrix(ii, 0) & "'"
         End If
         '2020/4/28 END
         '申請國家為台灣
         If m_Nation = "000" Then
            'A:服務費A=A1L04尾數非99之A1L05-A1L07
'            strExc(0) = "select sum(nvl(a1L05,0))-sum(nvl(a1L07,0))" & _
'                        " from acc1L0 where a1L01='" & txtA1K01 & "'" & _
'                        " and ((substr(a1L04,-2) in('98') and substr(a1L04,1,length(a1L04)-2)='" & Grd1.TextMatrix(ii, 0) & "')" & _
'                             " or a1L04='" & Grd1.TextMatrix(ii, 0) & "')"
            strExc(0) = "select sum(nvl(a1L05,0))-sum(nvl(a1L07,0))" & _
                        " from acc1L0 where a1L01='" & txtA1K01 & "'" & _
                        " and ((substr(a1L04,-2) in('98') and substr(a1L04,1,length(a1L04)-2)" & strCon & ")" & _
                             " or a1L04" & strCon & ")"
            intI = 1
            Set AdoRs = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               GRD1.TextMatrix(ii, 2) = Val("" & AdoRs.Fields(0))
            End If
            'B:代收規費B=A1L04尾數為99之A1L05-A1L07
'            strExc(0) = "select sum(nvl(a1L05,0))-sum(nvl(a1L07,0))" & _
'                        " from acc1L0 where a1L01='" & txtA1K01 & "'" & _
'                        " and (substr(a1L04,-2) in('99') and substr(a1L04,1,length(a1L04)-2)='" & Grd1.TextMatrix(ii, 0) & "')"
            strExc(0) = "select sum(nvl(a1L05,0))-sum(nvl(a1L07,0))" & _
                        " from acc1L0 where a1L01='" & txtA1K01 & "'" & _
                        " and (substr(a1L04,-2) in('99') and substr(a1L04,1,length(a1L04)-2)" & strCon & ")"
            intI = 1
            Set AdoRs = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               GRD1.TextMatrix(ii, 3) = Val("" & AdoRs.Fields(0))
            End If
         '申請國家非台灣
         Else
            '服務費=A+B
'            strExc(0) = "select sum(nvl(a1L05,0))-sum(nvl(a1L07,0))" & _
'                        " from acc1L0 where a1L01='" & txtA1K01 & "'" & _
'                        " and ((substr(a1L04,-2) in('98') and substr(a1L04,1,length(a1L04)-2)='" & Grd1.TextMatrix(ii, 0) & "')" & _
'                             " or (substr(a1L04,-2) in('99') and substr(a1L04,1,length(a1L04)-2)='" & Grd1.TextMatrix(ii, 0) & "')" & _
'                             " or a1L04='" & Grd1.TextMatrix(ii, 0) & "')"
            strExc(0) = "select sum(nvl(a1L05,0))-sum(nvl(a1L07,0))" & _
                        " from acc1L0 where a1L01='" & txtA1K01 & "'" & _
                        " and ((substr(a1L04,-2) in('98') and substr(a1L04,1,length(a1L04)-2)" & strCon & ")" & _
                             " or (substr(a1L04,-2) in('99') and substr(a1L04,1,length(a1L04)-2)" & strCon & ")" & _
                             " or a1L04" & strCon & ")"
            intI = 1
            Set AdoRs = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               GRD1.TextMatrix(ii, 2) = Val("" & AdoRs.Fields(0))
            End If
            '代收規費=0
            GRD1.TextMatrix(ii, 3) = 0
         End If
      Next ii
      Call CountTot
   End If
   GridHead
   
   Screen.MousePointer = vbDefault
End Sub

Private Sub CountTot()
Dim ii As Integer
Dim dblTot1 As Double, dblTot2 As Double
   
   dblTot1 = 0: dblTot2 = 0
   For ii = 1 To GRD1.Rows - 1
      dblTot1 = dblTot1 + GRD1.TextMatrix(ii, 2)
      dblTot2 = dblTot2 + GRD1.TextMatrix(ii, 3)
   Next ii
   LblSub_1 = dblTot1
   LblSub_2 = dblTot2
   LblTot = dblTot1 + dblTot2
End Sub

Private Sub GridHead()
   With GRD1
      .Visible = False
      .Cols = 4
      .row = 0
      .col = 0: .ColWidth(0) = 0: .Text = "請款項目代碼"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(0) = flexAlignCenterCenter
      
      .col = 1: .ColWidth(1) = 3000: .Text = "請款項目"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(1) = flexAlignLeftCenter
      
      .col = 2: .ColWidth(2) = 1500: .Text = "服務費"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(2) = flexAlignRightCenter
      
      .col = 3: .ColWidth(3) = 1500: .Text = "代收規費"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(3) = flexAlignRightCenter
      
      For intI = 4 To .Cols - 1
         .ColWidth(intI) = 0
      Next
      .Visible = True
   End With
End Sub

Private Sub UpdateCol()
   Dim ii As Integer
   
   If txtInput <> txtInput.Tag Then
      With GRD1
      If iCol = 2 Or iCol = 3 Then
         .TextMatrix(iRow, iCol) = Val(txtInput.Text)
         Call CountTot
'      Else
'         For ii = 1 To .Rows - 1
'            If .TextMatrix(ii, 0) = .TextMatrix(iRow, 0) Then
'               .TextMatrix(ii, iCol) = txtInput.Text
'            End If
'         Next
      End If
      End With
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   '底色
   m_dftColor = &HFFFFFF
   '底色2
   m_dftColor2 = RGB(&HFF, &HFA, &HCD)
   '底色3
   m_dftColor3 = &HFFC0C0
   
   'Modified by Lydia 2021/12/09 改成模組
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   Me.Width = 9120
'   'Modified by Lydia 2016/03/23
'   'Me.Height = 5700
'   Me.Height = 5985
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath4)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   strFormName = Name
   PUB_InitForm Me, 9120, 6015, strBackPicPath4
   'end 2021/12/09
   
   txtInput.Visible = False
   GRD1.Clear
   GridHead
   cmdWord.Enabled = False
   cmdSave.Enabled = False 'Added by Lydia 2016/03/25
   
   LblA1K02.BorderStyle = 0
   LblA1K28.BorderStyle = 0
   lblCaseNo.BorderStyle = 0
   LblCaseName.BorderStyle = 0
   LblCP13.BorderStyle = 0
   LblSub_1.BorderStyle = 0
   LblSub_2.BorderStyle = 0
   LblTot.BorderStyle = 0
   LblNTAmt.BorderStyle = 0
   LblMoney.BorderStyle = 0 'Add By Sindy 2015/10/20
   LblNTdiscount.BorderStyle = 0 'Add By Sindy 2015/10/20
   Lbldiscount.BorderStyle = 0 'Add By Sindy 2015/10/20
   
   'Add By Sindy 2025/9/18
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   m_AttachPath = m_AttachPath & "\收據"
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   m_AttachPath = m_AttachPath & "\"
   '清除暫存檔
   Call PUB_KillTempFile("$$" & Left(TransDate(DBDATE(DateAdd("m", -1, ChangeWStringToWDateString(strSrvDate(1)))), 1), 6) & "*.*", m_AttachPath)
   '2025/9/18 END
   
   'Add By Sindy 2020/8/25 M31-000010-0-00 智慧所收據.doc
   m_FileName2 = "$$智慧所收據.doc"
   'Modify By Sindy 2025/9/18 增加收據資料夾存放
   If Dir(m_AttachPath & m_FileName2) <> "" Then
      Kill m_AttachPath & m_FileName2
   End If
   Call PUB_GetSampleFile(m_FileName2, "M31-000010-0-00", , m_AttachPath)
   '2020/8/25 END
   
   'Add By Sindy 2020/12/16
   m_FileName3 = "$$法律所收據.doc"
   'Modify By Sindy 2025/9/18 增加收據資料夾存放
   If Dir(m_AttachPath & m_FileName3) <> "" Then
      Kill m_AttachPath & m_FileName3
   End If
   Call PUB_GetSampleFile(m_FileName3, "M31-000007-0-00", , m_AttachPath)
   '2020/12/16 END
   
   m_FileName = "$$國外請款單產生國內收據.doc"
   'Modify By Sindy 2025/9/18 增加收據資料夾存放
   If Dir(m_AttachPath & m_FileName) <> "" Then
      Kill m_AttachPath & m_FileName
   End If
   Call PUB_GetSampleFile(m_FileName, "M31-000005-0-00", , m_AttachPath)
   
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHand

   If Not g_WordAp Is Nothing Then
      g_WordAp.Visible = True
      g_WordAp.Quit
CloseWord:
      Set g_WordAp = Nothing
   End If
   
   StatusClear
   strFormName = MsgText(601)
   MenuEnabled
   Set Frmacc14v0 = Nothing
   
   Exit Sub
   
ErrHand:
   If Err.Number = 462 Then '遠端伺服器不存在或無法使用
      GoTo CloseWord
   ElseIf Err.Number <> 0 Then
      MsgBox (Err.Description)
   End If
End Sub

Private Sub SetColor(pRow As Integer, pColor As Long)
   With GRD1
   .row = pRow
   For intI = 0 To .Cols - 1
      .col = intI
      .CellBackColor = pColor
   Next
   End With
End Sub

Private Sub Grd1_Click()
   Dim iCurCol As Integer, iCurRow As Integer
   
   With GRD1
   If .MouseRow > 0 And .MouseRow < .Rows And .MouseCol < 18 Then
      iCurRow = .MouseRow
      iCurCol = .MouseCol
      .Visible = False
      
      .row = iCurRow
      .col = 0
      If Trim(.TextMatrix(.row, .col)) = "" Then
         '.TextMatrix(.row, .col) = "V"
         SetColor iCurRow, m_dftColor3
      Else
         '.TextMatrix(.row, .col) = ""
         SetColor iCurRow, m_dftColor
      End If
      
      .col = iCurCol
      iRow = .row: iCol = .col
      If .col = 2 Then SetBox
      If .col = 3 Then SetBox
           
      .Visible = True
   End If
   End With
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   txtInput.Visible = False
End Sub

Private Sub Grd1_Scroll()
   If txtInput.Visible = True Then
      SetBox False
   End If
End Sub

'Add By Sindy 2020/5/11
Private Sub Text1_Change()
   If Text1 = MsgText(601) Then
      Exit Sub
   Else
      Text2 = A0802Query(Text1, True)
   End If
End Sub
Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If Not (KeyAscii = 50 Or KeyAscii = 76) Then
      KeyAscii = 0
      Beep
      Exit Sub
   End If
End Sub
'2020/5/11 END

'Add By Sindy 2020/12/16
Private Sub Text1_Validate(Cancel As Boolean)
   Check1.Visible = True
   If Text1.Text = "L" Then
      Check1.Visible = False
      Check1.Value = 0
   End If
End Sub

'Add By Sindy 2015/10/20
Private Sub txtA1K01_GotFocus()
   CloseIme
   InverseTextBox txtA1K01
End Sub
Private Sub txtA1K01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'2015/10/20 END

Private Sub txtA1K35_GotFocus()
   OpenIme
   InverseTextBox txtA1K35
End Sub

Private Sub txtA1K35_Validate(Cancel As Boolean)
   '剔除跳行符號
   txtA1K35.Text = PUB_StringFilter(txtA1K35.Text)
   
   If txtA1K35.Enabled = False Then Exit Sub
   If txtA1K35.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtA1K35, txtA1K35.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub txtInput_GotFocus()
   InverseTextBox txtInput
   CloseIme
End Sub

Private Sub txtInput_LostFocus()
   If txtInput.Locked = False Then UpdateCol
   txtInput.Visible = False
End Sub

Private Sub SetBox(Optional pbolSetValue As Boolean = True)
   Dim lngLeft As Long, lngTop As Long
   Dim ii As Integer
   
   With GRD1
   If .LeftCol > .col Or .TopRow > .row Then
      txtInput.Visible = False
   Else
      txtInput.FontName = .CellFontName
      txtInput.FontSize = .CellFontSize
      If .CellAlignment < 3 Then
         txtInput.Alignment = 0 '靠左
      ElseIf .CellAlignment < 6 Then
         txtInput.Alignment = 2 '置中
      ElseIf .CellAlignment < 9 Then
         txtInput.Alignment = 1 '靠右
      Else
         txtInput.Alignment = 0 '靠左
      End If
      If pbolSetValue = True Then
         txtInput.Text = .TextMatrix(.row, .col)
      End If
      txtInput.Tag = txtInput.Text
      txtInput.Width = .ColWidth(.col) + 10
      txtInput.Height = .RowHeight(.row) - 5
      lngLeft = .Left + 20
      lngTop = .Top + .RowHeight(0) + 20
      For ii = .LeftCol To .col - 1
         lngLeft = lngLeft + .ColWidth(ii)
      Next
      For ii = .TopRow To .row - 1
         lngTop = lngTop + .RowHeight(ii)
      Next
      txtInput.Left = lngLeft: txtInput.Top = lngTop
      If txtInput.Left + txtInput.Width < .Left + .Width Then
         txtInput.Visible = True
         txtInput.SetFocus
         TextInverse txtInput
         iRow = .row: iCol = .col
      Else
         txtInput.Visible = False
      End If
   End If
   End With
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
   If KeyAscii = 8 Then Exit Sub
   
   If KeyAscii = vbKeyReturn Then
      UpdateCol
      GoNext
   ElseIf KeyAscii = vbKeyEscape Then
      txtInput = txtInput.Tag
      TextInverse txtInput
   '輸入欄位
   ElseIf iCol = 2 Or iCol = 3 Then
      If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9"))) And KeyAscii <> 46 Then
         KeyAscii = 0
         Beep
         Exit Sub
      End If
   End If
End Sub

Private Sub GoNext()
   With GRD1
      .col = 2
      If .row < .Rows - 1 Then
         .row = .row + 1
      Else
         .row = 1
      End If
      SetBox
   End With
End Sub

'Added by Lydia 2016/03/25
Private Sub txtA1V09_GotFocus()
   TextInverse txtA1V09
End Sub
Private Sub txtA1V09_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

'修改存檔
Private Sub CmdSave_Click()
   
   If TxtValidate Then
      Screen.MousePointer = vbHourglass
      
On Error GoTo ErrHand
      
      txtA1K35 = Trim(txtA1K35)
      
      'Modify By Sindy 2020/5/11 + ,a1k37='" & Text1 & "'
      strSql = "update acc1k0 set a1k35='" & txtA1K35 & "',a1k37='" & Text1 & "' where a1k01='" & txtA1K01 & "'"
      cnnConnection.Execute strSql
      'Modify By Sindy 2022/1/21 + Or Text1.Tag <> Text1.Text
      'If txtA1V09.Tag <> txtA1V09.Text Or Text1.Tag <> Text1.Text Then
         'Modify By Sindy 2020/5/11 + ,a1v03='" & Text1 & "'
         strSql = " update acc1v0 set a1v09=" & CNULL(txtA1V09.Text, True) & ",a1v03='" & Text1 & "' where a1v02='" & txtA1K01 & "' "
         cnnConnection.Execute strSql
      'End If
      Screen.MousePointer = vbDefault
   End If
   Exit Sub
   
ErrHand:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox "更新資料失敗！" & vbCrLf & Err.Description
   End If
End Sub
'end 2016/03/25
