VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc1420 
   AutoRedraw      =   -1  'True
   Caption         =   "補開收據列印"
   ClientHeight    =   6840
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   5930
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6840
   ScaleWidth      =   5930
   Begin VB.CheckBox Check2 
      BackColor       =   &H0000FFFF&
      Caption         =   "測試DOC列印"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   4050
      TabIndex        =   47
      Top             =   5160
      Value           =   1  '核取
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   270
      TabIndex        =   41
      Top             =   5550
      Width           =   4755
      Begin VB.CheckBox ChkST06 
         BackColor       =   &H00FFC0C0&
         Caption         =   "北所"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   930
         TabIndex        =   45
         Top             =   120
         Width           =   705
      End
      Begin VB.CheckBox ChkST06 
         BackColor       =   &H00FFC0C0&
         Caption         =   "中所"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   1820
         TabIndex        =   44
         Top             =   120
         Width           =   705
      End
      Begin VB.CheckBox ChkST06 
         BackColor       =   &H00FFC0C0&
         Caption         =   "南所"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   2710
         TabIndex        =   43
         Top             =   120
         Width           =   705
      End
      Begin VB.CheckBox ChkST06 
         BackColor       =   &H00FFC0C0&
         Caption         =   "高所"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   3600
         TabIndex        =   42
         Top             =   120
         Width           =   705
      End
      Begin VB.Label Label15 
         BackStyle       =   0  '透明
         Caption         =   "所別："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.5
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   46
         Top             =   180
         Width           =   675
      End
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1860
      TabIndex        =   14
      Top             =   3150
      Width           =   3735
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1860
      TabIndex        =   15
      Top             =   3480
      Width           =   3735
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "舊收據(點陣)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   4050
      TabIndex        =   22
      Top             =   4980
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.TextBox TextX 
      Height          =   285
      Left            =   1950
      TabIndex        =   20
      Top             =   4890
      Width           =   705
   End
   Begin VB.TextBox TextY 
      Height          =   285
      Left            =   1950
      TabIndex        =   21
      Top             =   5190
      Width           =   705
   End
   Begin VB.Frame Frame1 
      Caption         =   "智慧所收據設定"
      Height          =   660
      Left            =   270
      TabIndex        =   35
      Top             =   4200
      Width           =   5415
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   17
         Top             =   240
         Width           =   3240
      End
      Begin VB.OptionButton Option2 
         Caption         =   "列表機"
         Height          =   180
         Index           =   0
         Left            =   4170
         TabIndex        =   18
         Top             =   150
         Width           =   915
      End
      Begin VB.OptionButton Option2 
         Caption         =   "影印機"
         Height          =   180
         Index           =   1
         Left            =   4170
         TabIndex        =   19
         Top             =   390
         Width           =   915
      End
      Begin VB.Label Label12 
         Caption         =   "印表機"
         Height          =   315
         Left            =   105
         TabIndex        =   36
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   2130
      Style           =   2  '單純下拉式
      TabIndex        =   16
      Top             =   3870
      Width           =   3450
   End
   Begin VB.CheckBox chk 
      Caption         =   "列印客戶案件案號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   2640
      TabIndex        =   11
      Top             =   2070
      Value           =   1  '核取
      Width           =   2565
   End
   Begin VB.CheckBox chk 
      Caption         =   "不列印專利年費年度"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   270
      TabIndex        =   10
      Top             =   2070
      Width           =   2355
   End
   Begin VB.CommandButton Command2 
      Caption         =   "收據抬頭修改"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3780
      TabIndex        =   8
      Top             =   1395
      Width           =   1785
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2790
      TabIndex        =   7
      Top             =   1380
      Width           =   732
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1860
      TabIndex        =   13
      Top             =   2820
      Width           =   3735
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1860
      TabIndex        =   12
      Top             =   2490
      Width           =   3735
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3990
      TabIndex        =   6
      Top             =   1050
      Width           =   732
   End
   Begin VB.CheckBox chk 
      Caption         =   "列印統一編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   270
      TabIndex        =   9
      Top             =   1710
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1980
      TabIndex        =   5
      Top             =   1050
      Width           =   732
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1230
      TabIndex        =   2
      Top             =   390
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   480
      Style           =   1  '圖片外觀
      TabIndex        =   23
      Top             =   6180
      Width           =   4935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1230
      TabIndex        =   3
      Top             =   720
      Width           =   732
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1230
      MaxLength       =   10
      TabIndex        =   0
      Top             =   60
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   3990
      TabIndex        =   1
      Top             =   60
      Width           =   1575
      _ExtentX        =   2787
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   14737632
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11
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
      Left            =   3990
      TabIndex        =   4
      Top             =   720
      Width           =   1575
      _ExtentX        =   2787
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "案件性質名稱3"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   40
      Top             =   3180
      Width           =   1605
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "案件性質名稱4"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   39
      Top             =   3510
      Width           =   1605
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "橫軸偏移值(X)：　　　　　　(單位公分)"
      Height          =   180
      Index           =   2
      Left            =   465
      TabIndex        =   38
      Top             =   4950
      Width           =   3240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "縱軸偏移值(Y)：　　　　　　(單位公分)"
      Height          =   180
      Index           =   1
      Left            =   465
      TabIndex        =   37
      Top             =   5250
      Width           =   3240
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "法律所收據印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   34
      Top             =   3900
      Width           =   2805
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "是否結清(Y:全收/N:全銷)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   33
      Top             =   1410
      Width           =   2655
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "案件性質名稱2"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   32
      Top             =   2850
      Width           =   1605
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "案件性質名稱1"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   31
      Top             =   2520
      Width           =   1605
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3030
      TabIndex        =   30
      Top             =   1050
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "是否補開(Y/N)?"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   29
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   28
      Top             =   390
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "收據日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3030
      TabIndex        =   27
      Top             =   60
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "列印次數"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   270
      TabIndex        =   26
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "列印日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3030
      TabIndex        =   25
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收據號碼"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   24
      Top             =   60
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc1420"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/1/26 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoacc0k0 As New ADODB.Recordset
Public adocaseprogress As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Dim lngAmount1 As Long
Dim lngAmount2 As Long
Dim strAmount1 As String
Dim strAmount2 As String
Dim intLength As Integer
Dim intCounter As Integer
Dim douSAmount As Double
Dim douTAmount As Double
Dim intTimes As Integer
'Add by Morgan 2004/8/12
Dim strTemp As String
Dim strCustCaseNo As String '客戶案件案號
Dim m_FixNo As Integer   '2010/2/12 add by sonia 修法次數
Dim strPrinter As String, strPrinter2 As String, m_FileName As String 'Add By Sindy 2020/3/23
Dim SeekPrintL As Integer, SeekPrint As Integer 'Add By Sindy 2020/7/14
Public ProState As String '權限: 1.全所 2.該所
Dim m_sqlST06 As String 'Add By Sindy 2021/5/21
Dim m_FileName2 As String 'Add By Sindy 2022/1/26
Dim m_AttachPath As String 'Add By Sindy 2025/9/18


Private Sub Command1_Click()
Dim strName As String
Dim intList As Integer
Dim douAmount As Double
Dim douFee As Double
Dim strProduct As String
'Add By Cheng 2002/01/17
'Modify by Morgan 2004/8/12   '改成全域變數
'Dim strCustCaseNo As String '客戶案件案號
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim blnCombPrint As Boolean '收據項目合併列印
Dim intRecPos As Integer '目前指向第幾筆資料
Dim ii As Integer
'add by nickc 2007/02/08
Dim lngAmount As Double
'Add By Sindy 2009/07/14
Dim strPA08 As String, strPA09 As String
Dim strFeeType As String, strYF15 As String
Dim strKey(5) As String
Dim strCaseFee(1 To 2) As String
Dim bFind As Boolean
Dim varRef As Variant
'2009/07/14 End
Dim bolChk As Boolean
Dim strA0K11 As String, strVal As String
   
   If FormCheck = False Then
      'MsgBox MsgText(181), , MsgText(5)  'CANCEL BY SONIA 2015/9/17
      Exit Sub
   End If
   'Add By Sindy 2021/5/21
   bolChk = False
   m_sqlST06 = ""
   For ii = 0 To 3
      If ChkST06(ii).Value = 1 Then
         bolChk = True
         m_sqlST06 = m_sqlST06 & "'" & ii + 1 & "'"
      End If
   Next ii
   If bolChk = False Then
      MsgBox "請勾選所別！", vbExclamation
      ChkST06(0).SetFocus
      m_sqlST06 = ""
      Exit Sub
   End If
   m_sqlST06 = Replace(m_sqlST06, "''", "','")
   '2021/5/21 END
   'Add By Sindy 2021/9/10 智權人員收文純法務案件所沿生之佣金收據(智慧所向法律所請款,案號T999999), 收據列印的所別全部放在北所
   'TT-999999的收據固定北所列印。
   If ChkST06(0).Value = 1 Then
      m_sqlST06 = " and (GetAcc0j0(A0k01,'TT9')='TT999999000' or ST06 in(" & m_sqlST06 & "))"
   Else
      m_sqlST06 = " and (GetAcc0j0(A0k01,'TT9')<>'TT999999000' or GetAcc0j0(A0k01,'TT9') is null) and ST06 in(" & m_sqlST06 & ")"
   End If
   '2021/9/10 END
   
   Screen.MousePointer = vbHourglass
   lngAmount = 0
   intLength = 0
   intCounter = 0
   
   'Modify By Sindy 2021/5/28
   strA0K11 = ""
   strExc(0) = "select a0k11 from acc0k0 where a0k01 = '" & Text1 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strA0K11 = "" & RsTemp.Fields("a0k11")
   End If
   'Modify By Sindy 2017/6/29 + and (a0k37 is null or a0k37<>'N') 剔除全銷
   'Modified by Lydia 2023/11/13 開立INVOICE，不列印收據=> + and nvl(a0k32,'Y') <> 'Z'
   strVal = "select * from acc0k0, customer, staff where substr(a0k03, 1, 8) = cu01(+) and substr(a0k03, 9, 1) = cu02(+) and a0k11<>'J' and a0k01 = '" & Text1 & "' and ((to_number(substr(a0k01, 5, 5)) > 2000) or to_number(substr(a0k01, 5, 5)) <= 2000 and a0k02 >= 920101)" & _
               " and (a0k09 is null or a0k09 = 0) and (a0k37 is null or a0k37<>'N') and nvl(a0k32,'Y') <> 'Z'" & _
               " and a0k01 not in (select a0m02 from acc0m0 where a0m02 = '" & Text1 & "' and a0m03 is not null)" & _
               " and a0k20=st01(+)"
   If strA0K11 = "L" Then
      strSql = strVal & " and exists (" & Replace(strLOSSalesDuty, "\#ST06SQL#\", m_sqlST06) & ")" & _
               " union " & _
               strVal & m_sqlST06 & " and not exists (" & Replace(strLOSSalesDuty, "\#ST06SQL#\", "") & ")"
   Else
   '2021/5/28 END
      strSql = strVal & m_sqlST06
   End If
   If adoacc0k0.State = adStateOpen Then
      adoacc0k0.Close
   End If
   adoacc0k0.CursorLocation = adUseClient
   adoacc0k0.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0k0.RecordCount = 0 Then
      adoacc0k0.Close
      'Modify By Sindy 2021/5/28
      strVal = "select * from acc0k0, customer, staff where substr(a0k03, 1, 8) = cu01(+) and substr(a0k03, 9, 1) = cu02(+) and a0k11<>'J' and a0k01 = '" & Text1 & "' and ((to_number(substr(a0k01, 5, 5)) > 2000) or to_number(substr(a0k01, 5, 5)) <= 2000 and a0k02 >= 920101)" & _
                     " and (a0k09 is null or a0k09 = 0) and (a0k37 is null or a0k37<>'N')" & _
                     " and a0k20=st01(+)"
      If strA0K11 = "L" Then
         strSql = strVal & " and exists (" & Replace(strLOSSalesDuty, "\#ST06SQL#\", m_sqlST06) & ")" & _
                  " union " & _
                  strVal & m_sqlST06 & " and not exists (" & Replace(strLOSSalesDuty, "\#ST06SQL#\", "") & ")"
      Else
      '2021/5/28 END
         strSql = strVal & m_sqlST06
      End If
      adoacc0k0.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoacc0k0.RecordCount = 0 Then
         adoacc0k0.Close
         Screen.MousePointer = vbDefault
         MsgBox MsgText(28), , MsgText(5)
         Exit Sub
      Else
         adoacc0k0.Close
         Screen.MousePointer = vbDefault
         MsgBox MsgText(214), , MsgText(5)
         Exit Sub
      End If
   'Add by Morgan 2011/8/22 +控制已列印才可補開 and a0k19>0
   ElseIf Val("" & adoacc0k0("a0k19")) = 0 Then
      adoacc0k0.Close
      Screen.MousePointer = vbDefault
      MsgBox "收據尚未列印不可補開", , MsgText(5)
      Exit Sub
   End If
   
   If IsNull(adoacc0k0.Fields("a0k11").Value) = False Then
      'Modify by Morgan 2004/11/17
      'MsgBox MsgText(163) & adoacc0k0.Fields("a0k11").Value & MsgText(164), , MsgText(5)
      If MsgBox(MsgText(163) & adoacc0k0.Fields("a0k11").Value & MsgText(164) & "，按確定開始列印...", vbOKCancel + vbDefaultButton2, MsgText(5)) = vbCancel Then
         adoacc0k0.Close
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
   End If
         
   Do While adoacc0k0.EOF = False
      
      'Added by Morgan 2020/10/14
      '案件性質有改時更新帳款類別
      If Text6.Enabled = True Then
         If UpdateItemDesc(adoacc0k0.Fields("a0k01").Value, "" & adoacc0k0.Fields("a0k33").Value) = False Then
            adoacc0k0.Close
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         
         If IsNull(adoacc0k0.Fields("a0k33").Value) Then
            adoacc0k0.Requery
         End If
      End If
      'end 2020/10/14
   
      'Add By Sindy 2020/3/23 空白A4直接列印收據
      If adoacc0k0.Fields("a0k11").Value = "L" Then
         '切換印表機
         PUB_SetOsDefaultPrinter Combo1
         PUB_RestorePrinter Combo1
         PUB_PrintCaseReceipt_L m_AttachPath & m_FileName, adoacc0k0, Me.chk(0), Me.chk(1), Text4, Text6, Text7, FCDate(MaskEdBox2.Text)
         '還原印表機
         PUB_SetOsDefaultPrinter strPrinter
         PUB_RestorePrinter strPrinter
      Else
      '2020/3/23 END
         
         'Add By Sindy 2022/1/26 直接用Word範本列印收據
         If Check2.Value = 1 Then
            '切換印表機
            PUB_SetOsDefaultPrinter Combo2
            PUB_RestorePrinter Combo2
            PUB_PrintCaseReceipt_Doc m_AttachPath & m_FileName2, adoacc0k0, Me.chk(0), Me.chk(1), Me.chk(2), Text4, Text6, Text7, FCDate(MaskEdBox2.Text), , True
            '還原印表機
            PUB_SetOsDefaultPrinter strPrinter2
            PUB_RestorePrinter strPrinter2
         Else
         '2022/1/26 END
         
'            '舊收據(點陣)
'            If Check1.Value = 1 Then
'               'Modify by Morgan 2008/3/25 XP自定紙張需手動設定並將印表機預設為該紙張
'               '9x
'      '         If pub_OS = "1" Then
'      '            Printer.Height = 8750
'      '            Printer.Width = 13000
'      '         Else
'                  Printer.EndDoc 'Add By Sindy 2020/4/1 紙張會跑到,因此再下一次 .EndDoc
'                  Printer.PaperSize = PUB_GetPaperSize(1)
'                  Printer.Orientation = 1 '1.直印 2.橫印
'                  'Forms(0).StatusBar1.Panels(2).Text = Printer.PaperSize
'      '         End If
'               'end 2008/3/25
'               Printer.FontSize = 12
'               Printer.Font = "標楷體"
'               'Modify by Morgan 2011/9/20 改呼叫共用(與補印收據統一)
'               PUB_PrintCaseReceipt adoacc0k0, Me.chk(0), Me.chk(1), Me.chk(2), Text4, Text6, Text7, FCDate(MaskEdBox2.Text)
'               'PrintSum 'Remove by Morgan 2011/9/20 改呼叫共用(與補印收據統一)
'               'Remove by Morgan 2011/9/20 改呼叫共用(與補印收據統一)
'               'If intCounter = 0 Then
'               '   MsgBox MsgText(191), , MsgText(5)
'               '   Printer.KillDoc
'               'Else
'               '   Printer.EndDoc
'               'End If
'               Printer.EndDoc
'               'end 2011/9/20
'
'            Else
'               'Add By Sindy 2020/7/15 列印A4收據
'               PUB_RestorePrinter Combo2 'Add By Sindy 2020/7/14 切換印表機
'
'               Call PUB_PrintCaseReceiptTableMain(TextX, TextY, IIf(Option2(0).Value = True, 0, 1), 1, True)
'               PUB_PrintCaseReceipt adoacc0k0, Me.chk(0), Me.chk(1), Me.chk(2), Text4, Text6, Text7, FCDate(MaskEdBox2.Text), , True
'               Printer.EndDoc
'
'               PUB_RestorePrinter strPrinter2 'Add By Sindy 2020/7/14 還原印表機
'               '2020/7/15 END
'            End If
'            Printer.Font = "新細明體"
         End If
      End If
      
      adoacc0k0.MoveNext
   Loop
   adoacc0k0.Close
   
   EndOfficeAp 'Added by Morgan 2025/9/10 印完要清除物件，否則印表機不會變
   
   Screen.MousePointer = vbDefault
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   
'Remove by Morgan 2011/9/20
'
'      intCounter = 0
'      intList = 0
'      lngAmount1 = 0
'      lngAmount2 = 0
'      If IsNull(adoacc0k0.Fields("a0k19").Value) Then
'         intTimes = 1
'      Else
'         intTimes = Val(adoacc0k0.Fields("a0k19").Value) + 1
''      intTimes = Val(Text2)
'      End If
'      'Add by Morgan 2004/10/14
'      'Modify by Morgan 2006/2/6 判斷列印日期不同於收據日期時要回寫
'      'adoTaie.Execute "update acc0k0 set a0k19 = " & intTimes & " where a0k01 = '" & adoacc0k0.Fields("a0k01").Value & "'"
'      If MaskEdBox2.Text = MaskEdBox1.Text Then
'         '2010/4/23 MODIFY BY SONIA 同時更新A0K32
'         adoTaie.Execute "update acc0k0 set a0k19 = " & intTimes & " where a0k01 = '" & adoacc0k0.Fields("a0k01").Value & "'"
'         'adoTaie.Execute "update acc0k0 set a0k19 = " & intTimes & ",a0k32=null where a0k01 = '" & adoacc0k0.Fields("a0k01").Value & "'"
'      Else
'         '2010/4/23 MODIFY BY SONIA 同時更新A0K32
'         adoTaie.Execute "update acc0k0 set a0k02=" & FCDate(MaskEdBox2.Text) & ",a0k19 = " & intTimes & " where a0k01 = '" & adoacc0k0.Fields("a0k01").Value & "'"
'         'adoTaie.Execute "update acc0k0 set a0k02=" & FCDate(MaskEdBox2.Text) & ",a0k19 = " & intTimes & ",a0k32=null where a0k01 = '" & adoacc0k0.Fields("a0k01").Value & "'"
'      End If
'      '2006/2/6 end
'      '2004/10/14 end
'
'      PrintHead
'        'Add By Cheng 2003/08/14
'        blnCombPrint = False
'        'CFP 美國 領證(601)及公開費(217)
'        'Modify by Morgan 2007/5/8
'        'StrSQLa = "select Count(*) from caseprogress, casepropertymap, acc0j0 where cp01 = cpm01 and cp10 = cpm02 and cp09 = a0j01 (+) and cp60 = '" & adoacc0k0.Fields("a0k01").Value & "' and (cp79 <> 0 or (cp79 = 0 and cp75 <> 0)) And CP01='CFP' And A0j04='101' And A0j03 In ('601','217') "
'        StrSQLa = "select Count(*) from caseprogress, acc0j0 where CP01='CFP' and cp60 = '" & adoacc0k0.Fields("a0k01").Value & "' and (cp79 <> 0 or (cp79 = 0 and cp75 <> 0)) And a0j01(+)=cp09 And A0j04='101' And A0j03 In ('601','217') "
'        rsA.CursorLocation = adUseClient
'        rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'        If rsA.RecordCount > 0 Then
'            If Val("" & rsA.Fields(0).Value) = 2 Then blnCombPrint = True
'        End If
'        If rsA.State <> adStateClosed Then rsA.Close
'        Set rsA = Nothing
'      adocaseprogress.CursorLocation = adUseClient
'      '2005/6/16 MODIFY BY SONIA 加收文號排序
'      'adocaseprogress.Open "select * from caseprogress, casepropertymap, acc0j0 where cp01 = cpm01 and cp10 = cpm02 and cp09 = a0j01 (+) and cp60 = '" & adoacc0k0.Fields("a0k01").Value & "' and (cp79 <> 0 or (cp79 = 0 and cp75 <> 0))", adoTaie, adOpenStatic, adLockReadOnly
'      'Modify by Morgan 2007/5/8 案件性質名稱直接抓0j0的以免資料不一致卻沒發現
'      'adocaseprogress.Open "select * from caseprogress, casepropertymap, acc0j0 where cp01 = cpm01 and cp10 = cpm02 and cp09 = a0j01 (+) and cp60 = '" & adoacc0k0.Fields("a0k01").Value & "' and (cp79 <> 0 or (cp79 = 0 and cp75 <> 0)) ORDER BY CP09 ", adoTaie, adOpenStatic, adLockReadOnly
'      adocaseprogress.Open "select * from caseprogress, acc0j0 where cp60 = '" & adoacc0k0.Fields("a0k01").Value & "' and (cp79 <> 0 or (cp79 = 0 and cp75 <> 0)) and a0j01(+)=cp09 ORDER BY CP09 ", adoTaie, adOpenStatic, adLockReadOnly
'      'end 2007/5/8
'      '2005/6/16 END
'        intRecPos = 0
'        ii = 0
'      Do While adocaseprogress.EOF = False
'         intRecPos = intRecPos + 1
'
'         'Modify by Morgan 2004/10/14 移到Loop前面只做一次就好
'         'adoTaie.Execute "update acc0k0 set a0k19 = " & intTimes & " where a0k01 = '" & adoacc0k0.Fields("a0k01").Value & "'"
'         '2004/10/14 end
'
'         adoquery.CursorLocation = adUseClient
'         adoquery.Open "select sum(a1u07), sum(a1u08), sum(a1u10), sum(a1u09) from acc1u0 where a1u03 = '" & adocaseprogress.Fields("cp09").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'         If adoquery.RecordCount <> 0 Then
'            If IsNull(adoquery.Fields(0).Value) = False Then
'               'Modify by Morgan 2004/10/14 銷帳後費用為0的不印
'               'If Val(adoquery.Fields(0).Value) = Val(adocaseprogress.Fields("cp16").Value) Then
'               If Val("" & adoquery.Fields(0).Value) + Val("" & adoquery.Fields(3).Value) = Val("" & adocaseprogress.Fields("cp16").Value) Then
'                    If blnCombPrint = True And intRecPos = 2 Then
'                    Else
'                        adoquery.Close
'                        GoTo NextRecord
'                    End If
'               End If
'               If (Val(adoquery.Fields(1).Value) + Val(adoquery.Fields(2).Value)) = Val(adocaseprogress.Fields("cp16").Value) Then
'                    If blnCombPrint = True And intRecPos = 2 Then
'                    Else
'                        adoquery.Close
'                        GoTo NextRecord
'                    End If
'               End If
'
'               If IsNull(adoquery.Fields(0).Value) = False Then
'                  douAmount = adoquery.Fields(0).Value
'               Else
'                  douAmount = 0
'               End If
'               If IsNull(adoquery.Fields(1).Value) = False Then
'                  douSAmount = adoquery.Fields(1).Value
'               Else
'                  douSAmount = 0
'               End If
'               If IsNull(adoquery.Fields(2).Value) = False Then
'                  douTAmount = adoquery.Fields(2).Value
'               Else
'                  douTAmount = 0
'               End If
'               If IsNull(adoquery.Fields(3).Value) = False Then
'                  douFee = adoquery.Fields(3).Value
'               Else
'                  douFee = 0
'               End If
'            Else
'               douAmount = 0
'               douSAmount = 0
'               douTAmount = 0
'               douFee = 0
'            End If
'         Else
'            douAmount = 0
'            douSAmount = 0
'            douTAmount = 0
'            douFee = 0
'         End If
'         adoquery.Close
'
'        If blnCombPrint = False Then
'            Printer.CurrentX = 350
'            Printer.CurrentY = 3600 + intCounter * 300
'            If IsNull(adocaseprogress.Fields("cp05").Value) = False Then
'               Printer.Print CFDate(ACDate(adocaseprogress.Fields("cp05").Value))
'            Else
'               Printer.Print ""
'            End If
'        Else
'            If adocaseprogress.Fields("A0j03") = "601" Then
'                Printer.CurrentX = 350
'                Printer.CurrentY = 3600 + intCounter * 300
'                If IsNull(adocaseprogress.Fields("cp05").Value) = False Then
'                   Printer.Print CFDate(ACDate(adocaseprogress.Fields("cp05").Value))
'                Else
'                   Printer.Print ""
'                End If
'            End If
'        End If
'
'      'Modify by Morgan 2007/5/8 案件性質名稱直接抓0j0的以免資料不一致卻沒發現(程式有修剪)
'      If blnCombPrint = False Or adocaseprogress.Fields("A0j03") = "601" Then
'         'Add By Sindy 2009/07/14
'         '加印繳費年度說明
'         If ((adocaseprogress.Fields("cp01").Value = "P" And adocaseprogress.Fields("cp10").Value = "601") Or _
'               ((adocaseprogress.Fields("cp01").Value = "P" Or adocaseprogress.Fields("cp01").Value = "CFP") And _
'               (adocaseprogress.Fields("cp10").Value = "605" Or adocaseprogress.Fields("cp10").Value = "606" Or adocaseprogress.Fields("cp10").Value = "607"))) And _
'            adocaseprogress.Fields("cp53").Value <> "" And _
'            adocaseprogress.Fields("cp54").Value <> "" Then
'            strSql = "SELECT * FROM patent WHERE pa01='" & adocaseprogress.Fields("cp01").Value & "' and pa02='" & adocaseprogress.Fields("cp02").Value & "' and pa03='" & adocaseprogress.Fields("cp03").Value & "' and pa04='" & adocaseprogress.Fields("cp04").Value & "' "
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'            If intI = 1 Then
'               strPA08 = RsTemp("PA08")
'               strPA09 = RsTemp("PA09")
'            End If
'            strKey(0) = ""
'            strKey(1) = adocaseprogress.Fields("cp01").Value
'            strKey(2) = adocaseprogress.Fields("cp02").Value
'            strKey(3) = adocaseprogress.Fields("cp03").Value
'            strKey(4) = adocaseprogress.Fields("cp04").Value
'            '2010/2/12 modify by sonia 抓修法次數
'            'bFind = GetMoneyDate(strPA08, strPA09, strKey, strCaseFee(1), strCaseFee(2))
'            bFind = GetMoneyDate(strPA08, strPA09, strKey, strCaseFee(1), strCaseFee(2), , , m_FixNo)
'            If bFind Then
'               If IsEmptyText(strCaseFee(2)) = False Then
'                  varRef = Split(strCaseFee(2), ",")
'                  strFeeType = PUB_GetNa20Na22Na24(strPA09, strPA08)
'                  If adocaseprogress.Fields("cp10").Value = "605" Or _
'                     adocaseprogress.Fields("cp10").Value = "601" Then
'                     strYF15 = PUB_GetYF15(strPA09, strPA08, "Y000000" & m_FixNo, strFeeType, CDbl(Val(adocaseprogress.Fields("cp53").Value)))
'                  Else
'                     strYF15 = PUB_GetYF15(strPA09, strPA08, "Y000000" & m_FixNo, strFeeType, Val(varRef(Val(adocaseprogress.Fields("cp53").Value) - 1)))
'                  End If
'                  If chk(2).Value <> 1 Then 'Modify By Sindy 2010/3/30
'                     Printer.FontSize = 11
'                     Printer.Font = "標楷體"
'                     Printer.CurrentX = 1550
'                     Printer.CurrentY = 3600 + intCounter * 300
'                     If Val(adocaseprogress.Fields("cp53").Value) = Val(adocaseprogress.Fields("cp54").Value) Then
'                        Printer.Print adocaseprogress.Fields("a0j20").Value & " " & strYF15
'                     Else
'                        If adocaseprogress.Fields("cp10").Value = "605" Or _
'                           adocaseprogress.Fields("cp10").Value = "601" Then
'                           Printer.Print adocaseprogress.Fields("a0j20").Value & " (" & strYF15 & "至" & PUB_GetYF15(strPA09, strPA08, "Y000000" & m_FixNo, strFeeType, CDbl(Val(adocaseprogress.Fields("cp54").Value))) & ")"
'                        Else
'                           Printer.Print adocaseprogress.Fields("a0j20").Value & " (" & strYF15 & "至" & PUB_GetYF15(strPA09, strPA08, "Y000000" & m_FixNo, strFeeType, Val(varRef(Val(adocaseprogress.Fields("cp54").Value) - 1))) & ")"
'                        End If
'                     End If
'                  End If '2010/3/30
'               End If
'            End If
'         '2009/07/14 End
'         Else
'            ii = ii + 1
'            Printer.CurrentX = 1550
'            Printer.CurrentY = 3600 + intCounter * 300
'
'            If ii = 1 And Me.Text6.Text <> "" Then
'               Printer.Print Me.Text6.Text
'            ElseIf ii = 2 And Me.Text7.Text <> "" Then
'               Printer.Print Me.Text7.Text
'            Else
'               Printer.Print adocaseprogress.Fields("a0j20").Value
'            End If
'         End If
'      End If
'      Printer.FontSize = 12
'      Printer.Font = "標楷體"
'      'end 2007/5/8
'
'      If IsNull(adocaseprogress.Fields("cp16").Value) = False Then
'         If IsNull(adocaseprogress.Fields("cp17").Value) = False And IsNull(adoacc0k0.Fields("a0k30").Value) Then
'            If (Val(adocaseprogress.Fields("cp16").Value) - Val(adocaseprogress.Fields("cp17").Value) - douAmount) > 0 Then
'               strAmount1 = Format((Val(adocaseprogress.Fields("cp16").Value) - Val(adocaseprogress.Fields("cp17").Value) - douAmount), DDollar)
'            Else
'               strAmount1 = "0"
'            End If
'             If blnCombPrint = False Then
'                 intLength = Printer.TextWidth(strAmount1)
'                 Printer.CurrentX = 5900 - intLength
'                 Printer.CurrentY = 3600 + intCounter * 300
'                 Printer.Print strAmount1
'             End If
'            If (Val(adocaseprogress.Fields("cp16").Value) - Val(adocaseprogress.Fields("cp17").Value) - douAmount) > 0 Then
'               lngAmount1 = lngAmount1 + Val(adocaseprogress.Fields("cp16").Value) - Val(adocaseprogress.Fields("cp17").Value) - douAmount
'            Else
'               'lngAmount1 = 0
'            End If
'            If Val(adocaseprogress.Fields("cp17").Value) = 0 Then
'               strAmount2 = "0"
'            Else
'               If (Val(adocaseprogress.Fields("cp17").Value) - douFee) = 0 Then
'                  strAmount2 = "0"
'               Else
'                  strAmount2 = Format(Val(adocaseprogress.Fields("cp17").Value) - douFee, DDollar)
'               End If
'            End If
'             If blnCombPrint = False Then
'                 intLength = Printer.TextWidth(strAmount2)
'                 Printer.CurrentX = 7600 - intLength
'                 Printer.CurrentY = 3600 + intCounter * 300
'                 Printer.Print strAmount2
'             End If
'            If (Val(adocaseprogress.Fields("cp17").Value) - douFee) = 0 Then
'               lngAmount2 = lngAmount2
'            Else
'               lngAmount2 = lngAmount2 + Val(adocaseprogress.Fields("cp17").Value) - douFee
'            End If
'         Else
'            strAmount1 = Format(Val(adocaseprogress.Fields("cp16").Value) - douAmount - douFee, DDollar)
'             If blnCombPrint = False Then
'                 intLength = Printer.TextWidth(strAmount1)
'                 Printer.CurrentX = 5900 - intLength
'                 Printer.CurrentY = 3600 + intCounter * 300
'                 If strAmount1 = "" Then
'                    Printer.Print "0"
'                 Else
'                    Printer.Print strAmount1
'                 End If
'             End If
'            lngAmount1 = lngAmount1 + Val(adocaseprogress.Fields("cp16").Value) - douAmount - douFee
'            strAmount2 = "0"
'             If blnCombPrint = False Then
'                 intLength = Printer.TextWidth(strAmount2)
'                 Printer.CurrentX = 7600 - intLength
'                 Printer.CurrentY = 3600 + intCounter * 300
'                 Printer.Print strAmount2
'                 lngAmount2 = lngAmount2 + 0
'             End If
'         End If
'      End If
'
'      'Add By Cheng 2003/08/14
'      If blnCombPrint = True And intRecPos = 2 Then
'          '列印費用
'          If lngAmount1 = 0 Then
'              strAmount1 = "0"
'          Else
'              strAmount1 = Format(lngAmount1, DDollar)
'          End If
'          intCounter = intCounter - 1
'          intLength = Printer.TextWidth(strAmount1)
'          Printer.CurrentX = 5900 - intLength
'          Printer.CurrentY = 3600 + intCounter * 300
'          Printer.Print strAmount1
'          '列印規費
'          If lngAmount2 = 0 Then
'              strAmount2 = "0"
'          Else
'              strAmount2 = Format(lngAmount2, DDollar)
'          End If
'          intLength = Printer.TextWidth(strAmount2)
'          Printer.CurrentX = 7600 - intLength
'          Printer.CurrentY = 3600 + intCounter * 300
'          Printer.Print strAmount2
'      End If
'
'         If intList < 2 Then
'            'Modify By Cheng 2003/08/14
'            If blnCombPrint = False Or (blnCombPrint = True And intRecPos = 1) Then
'                Printer.CurrentX = 1500
'                Printer.CurrentY = 7200 + intList * 550
'                If (adocaseprogress.Fields("cp03").Value & adocaseprogress.Fields("cp04").Value) = "000" Then
'                   Printer.Print adocaseprogress.Fields("cp01").Value & adocaseprogress.Fields("cp02").Value
'                Else
'                   Printer.Print adocaseprogress.Fields("cp01").Value & adocaseprogress.Fields("cp02").Value & adocaseprogress.Fields("cp03").Value & adocaseprogress.Fields("cp04").Value
'                End If
'                Printer.CurrentX = 3000
'                Printer.CurrentY = 7200 + intList * 550
'                If IsNull(adocaseprogress.Fields("a0j21").Value) Then
'                   Printer.Print ""
'                Else
'                   Printer.Print adocaseprogress.Fields("a0j21").Value
'                End If
'                If adoquery.State = adStateOpen Then
'                   adoquery.Close
'                End If
'                adoquery.CursorLocation = adUseClient
'                adoquery.Open "select tm09 from trademark where tm01 = '" & adocaseprogress.Fields("cp01").Value & "' and tm02 = '" & adocaseprogress.Fields("cp02").Value & "' and tm03 = '" & adocaseprogress.Fields("cp03").Value & "' and tm04 = '" & adocaseprogress.Fields("cp04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'                If adoquery.RecordCount <> 0 Then
'                   If IsNull(adoquery.Fields("tm09").Value) Then
'                      strProduct = MsgText(601)
'                   Else
'                      '2010/8/24 modify by sonia 加 '類'
'                      strProduct = "第" & Trim(adoquery.Fields("tm09").Value) & " 類"
'                   End If
'                Else
'                   strProduct = MsgText(601)
'                End If
'                adoquery.Close
'                '2005/7/1 ADD BY SONIA 查名案加印商品類別
'                If adocaseprogress.Fields("cp01").Value = "S" Or adocaseprogress.Fields("cp01").Value = "TS" Then
'                  adoquery.CursorLocation = adUseClient
'                  adoquery.Open "select SP18 from SERVICEPRACTICE where SP01 = '" & adocaseprogress.Fields("cp01").Value & "' and SP02 = '" & adocaseprogress.Fields("cp02").Value & "' and SP03 = '" & adocaseprogress.Fields("cp03").Value & "' and SP04 = '" & adocaseprogress.Fields("cp04").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'                  If adoquery.RecordCount <> 0 Then
'                     If IsNull(adoquery.Fields("SP18").Value) Then
'                        strProduct = MsgText(601)
'                     Else
'                        '2010/8/24 modify by sonia 加 '類'
'                        strProduct = "第" & adoquery.Fields("SP18").Value & " 類"
'                     End If
'                  Else
'                     strProduct = MsgText(601)
'                  End If
'                  adoquery.Close
'                End If
'                '2005/7/1 END
'                Printer.CurrentX = 5500
'                Printer.CurrentY = 7200 + intList * 550
'                '列印案件名稱(以總收文號抓案件名稱)及客
'                'Modify by Morgan 2004/8/12
''                If CaseNameQuery(adocaseprogress.Fields("cp09").Value, 1, strCustCaseNo) <> "" Then
''                   'Modify By Cheng 2002/01/17
''    '               Printer.Print CaseNameQuery(adocaseprogress.Fields("cp09").Value, 1)
''                    'Modify By Cheng 2003/06/02
''    '               Printer.Print Trim(CaseNameQuery(adocaseprogress.Fields("cp09").Value, 1, strCustCaseNo)) & _
''    '                              IIf(Me.chk(1).Value = vbChecked, " " & strCustCaseNo, "") & " " & strProduct
''                   Printer.Print ReOrgPrintTemp(Trim(CaseNameQuery(adocaseprogress.Fields("cp09").Value, 1, strCustCaseNo)) & _
''                                  IIf(Me.chk(1).Value = vbChecked, " " & strCustCaseNo, "") & " " & strProduct, 22)
''                Else
''    '               Printer.Print CaseNameQuery(adocaseprogress.Fields("cp09").Value, 2)
''                    'Modify By Cheng 2003/06/02
''    '               Printer.Print Trim(CaseNameQuery(adocaseprogress.Fields("cp09").Value, 2, strCustCaseNo)) & _
''    '                              IIf(Me.chk(1).Value = vbChecked, " " & strCustCaseNo, "") & " " & strProduct
''                   Printer.Print ReOrgPrintTemp(Trim(CaseNameQuery(adocaseprogress.Fields("cp09").Value, 2, strCustCaseNo)) & _
''                                  IIf(Me.chk(1).Value = vbChecked, " " & strCustCaseNo, "") & " " & strProduct, 22)
''                End If
'                strTemp = CaseNameQuery(adocaseprogress.Fields("cp09").Value, 1, strCustCaseNo)
'                If strTemp <> "" Then
'                     Printer.Print ReOrgPrintTemp(Trim(strTemp) & " " & strProduct, 22)
'                Else
'                     Printer.Print ReOrgPrintTemp(Trim(CaseNameQuery(adocaseprogress.Fields("cp09").Value, 2, strCustCaseNo)) & " " & strProduct, 22)
'                End If
'                'Modify end
'
'                intList = intList + 1
'                strName = adocaseprogress.Fields("cp01").Value & adocaseprogress.Fields("cp02").Value & adocaseprogress.Fields("cp03").Value & adocaseprogress.Fields("cp04").Value
'            End If
'         End If
'         intCounter = intCounter + 1
'
'         'Add by Morgan 2005/9/15 顧問案的顧問聘任(0)時
'         If adocaseprogress("CP01") = "LA" And adocaseprogress("CP10") = "0" Then
'            Printer.CurrentX = 1550
'            Printer.CurrentY = 3600 + intCounter * 300
'            Printer.Print "顧問期間："
'            intCounter = intCounter + 1
'            Printer.CurrentX = 1550
'            Printer.CurrentY = 3600 + intCounter * 300
'            Printer.Print Format(TransDate("" & adocaseprogress("CP53"), 1), "###/##/##") & " - " & Format(TransDate("" & adocaseprogress("CP54"), 1), "###/##/##")
'            intCounter = intCounter + 1
'         End If
'
'NextRecord:
'         adocaseprogress.MoveNext
'      Loop
'      adocaseprogress.Close
'      'Add by Morgan 2004/8/12
'      If chk(1).Value = 1 And strCustCaseNo <> "" Then
'         Printer.CurrentX = 8200 - Printer.TextWidth("客戶案件案號：")
'         Printer.CurrentY = 2700
'         Printer.Print "客戶案件案號：" & strCustCaseNo
'      End If
'      'Add end
End Sub

'Add By Sindy 2021/5/21
Private Sub Form_Activate()
Dim intST06 As Integer, ii As Integer
   
   For ii = 0 To 3
      ChkST06(ii).Value = 0
   Next ii
   
   intST06 = Val(PUB_GetST06(strUserNum)) - 1
   ChkST06(intST06).Value = 1
   '權限: 1.全所 2.該所
   If ProState = "2" Then
      For ii = 0 To 3
         ChkST06(ii).Enabled = False
      Next ii
   Else
      For ii = 0 To 3
         ChkST06(ii).Enabled = True
      Next ii
   End If
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
   Me.Width = 6045
   'Modify By Cheng 2002/01/17
'   Me.Height = 2600
   'Modified by Lydia 2016/04/13
   'Me.Height = 4080
   Me.Height = 7300 'Modify by Amy 2023/10/11 原:7095
   'Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   MoveFormToCenter Me
   
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   Text1 = MsgText(802)
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Text = CFDate(ACDate(ServerDate))
   MaskEdBox2.Mask = DFormat
   Text4 = MsgText(602)
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   
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
   
   'Add By Sindy 2020/3/23
   PUB_SetPrinter Me.Name, Combo1, strPrinter
   m_FileName = "$$法律所收據.doc"
   'Modify By Sindy 2025/9/18 增加收據資料夾存放
   If Dir(m_AttachPath & m_FileName) <> "" Then
      Kill m_AttachPath & m_FileName
   End If
   Call PUB_GetSampleFile(m_FileName, "M31-000007-0-00", , m_AttachPath)
   '2020/3/23 END
   'Add By Sindy 2022/1/26
   m_FileName2 = "$$智慧所收據.doc"
   'Modify By Sindy 2025/9/18 增加收據資料夾存放
   If Dir(m_AttachPath & m_FileName2) <> "" Then
      Kill m_AttachPath & m_FileName2
   End If
   Call PUB_GetSampleFile(m_FileName2, "M31-000010-0-00", , m_AttachPath)
   '2022/1/26 END
   
   'Add By Sindy 2020/7/14 設定印表機改呼叫公用函數,原程式移除
   SeekPrintL = Printer.Orientation
   PUB_SetPrinter Me.Name, Combo2, strPrinter2, , SeekPrint, TextX, TextY
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2020/3/23
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   '2020/3/23 END
   
   'Add By Sindy 2020/7/14
   '若印表機或偏移值有變動, 則更新列印設定
   If Me.Combo2.Text <> Me.Combo2.Tag Or Me.TextX.Text <> Me.TextX.Tag Or Me.TextY.Text <> Me.TextY.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, Me.TextX.Text, Me.TextY.Text, Me.Combo2.Text
   End If
   Set Printer = Printers(SeekPrint)
   Printer.Orientation = SeekPrintL
   
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc1420 = Nothing
End Sub

Private Sub Combo2_Click()
   Option2(0).Value = True
   If InStr(Combo2.Text, "影印機") > 0 Then
      Option2(1).Value = True
   End If
End Sub

Private Sub Text1_Change()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim blnShowCPName As Boolean  '是否顯示案件性質
   
   If Len(Text1) < 9 Then
      Exit Sub
   End If
   Text6.Text = "": Text6.Enabled = False: Text6.BackColor = Text3.BackColor
   Text7.Text = "": Text7.Enabled = False: Text7.BackColor = Text3.BackColor
   Text9.Text = "": Text9.Enabled = False: Text9.BackColor = Text3.BackColor 'Added by Morgan 2020/10/14
   Text10.Text = "": Text10.Enabled = False: Text10.BackColor = Text3.BackColor 'Added by Morgan 2020/10/14
   
   If adoacc0k0.State = adStateOpen Then
      adoacc0k0.Close
   End If
   adoacc0k0.CursorLocation = adUseClient
   adoacc0k0.Open "select * from acc0k0 where a0k01 = '" & Text1 & "' and a0k11<>'J'", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0k0.RecordCount <> 0 Then
      MaskEdBox1.Mask = ""
      If IsNull(adoacc0k0.Fields("a0k02").Value) = False Then
         MaskEdBox1.Text = CFDate(adoacc0k0.Fields("a0k02").Value)
      Else
         MaskEdBox1.Text = ""
      End If
      MaskEdBox1.Mask = DFormat
      If IsNull(adoacc0k0.Fields("a0k19").Value) = False Then
         Text2 = adoacc0k0.Fields("a0k19").Value
      Else
         Text2 = ""
      End If
      If IsNull(adoacc0k0.Fields("a0k04").Value) = False Then
         Text3 = adoacc0k0.Fields("a0k04").Value
      Else
         Text3 = ""
      End If
      '92.6.27 add by sonia
      If IsNull(adoacc0k0.Fields("a0k11").Value) = False Then
         Text5 = adoacc0k0.Fields("a0k11").Value
      Else
         Text5 = ""
      End If
      'Added by Lydia 2016/04/13 顯示已收款
      If IsNull(adoacc0k0.Fields("a0k37").Value) = False Then
         Text8 = adoacc0k0.Fields("a0k37").Value
      Else
         Text8 = ""
      End If
      'end 2016/04/13
      
      'Add By Cheng 2003/10/31
      blnShowCPName = False
      'Modified by Morgan 2015/4/28 +考慮拆收據
      'StrSQLa = "select * from caseprogress, casepropertymap, acc0j0 where cp01 = cpm01 and cp10 = cpm02 and cp09 = a0j01 (+) and cp60 = '" & Me.Text1.Text & "' and (cp79 <> 0 or (cp79 = 0 and cp75 <> 0)) And (CP77 Is Not Null And CP77 > 0 ) "
      StrSQLa = "select * from acc0j0, caseprogress, casepropertymap where a0j13 = '" & Me.Text1.Text & "' and cp09(+)=a0j01 and (cp79 <> 0 or (cp79 = 0 and cp75 <> 0)) And (CP77 Is Not Null And CP77 > 0 ) and cpm01(+)=cp01 and cpm02(+)=cp10"
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
          blnShowCPName = True
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      If blnShowCPName = True Then
         
         'Modified by Morgan 2019/4/25 控制未變更帳款類別才可改案件性質(原來改了也不會印),只能由電腦中心人工修改--瑞婷
         'Modified by Morgan 2020/10/14  改有變更帳款類別也可改案件性質
         'If IsNull(adoacc0k0.Fields("a0k33").Value) Then
         If Not IsNull(adoacc0k0.Fields("a0k33").Value) Then
            If rsA.State <> adStateClosed Then rsA.Close
            StrSQLa = "select a0j22,min(a0j25) a0j25 from acc0j0 where a0j13 = '" & Me.Text1.Text & "' group by a0j22 order by a0j25 asc"
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Me.Text6.Text = "" & rsA.Fields("a0j22").Value
               Me.Text6.Enabled = True: Text6.BackColor = Text1.BackColor
               Me.Text6.Tag = Me.Text6.Text
               rsA.MoveNext
               If rsA.EOF = False Then
                  Me.Text7.Text = "" & rsA.Fields("a0j22").Value
                  Me.Text7.Enabled = True: Text7.BackColor = Text1.BackColor
                  Me.Text7.Tag = Me.Text7.Text
                  rsA.MoveNext
                  If rsA.EOF = False Then
                     Me.Text9.Text = "" & rsA.Fields("a0j22").Value
                     Me.Text9.Enabled = True: Text9.BackColor = Text1.BackColor
                     Me.Text9.Tag = Me.Text9.Text
                     rsA.MoveNext
                     If rsA.EOF = False Then
                        Me.Text10.Text = "" & rsA.Fields("a0j22").Value
                        Me.Text10.Enabled = True: Text10.BackColor = Text1.BackColor
                        Me.Text10.Tag = Me.Text10.Text
                     End If
                  End If
               End If
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
         Else
         'end 2020/10/14
         
             'Modified by Morgan 2014/5/13 +和收據列印一樣用收文號排序
             'Modified by Morgan 2015/4/28 +考慮拆收據
             'StrSQLa = "select * from caseprogress, casepropertymap, acc0j0 where cp01 = cpm01 and cp10 = cpm02 and cp09 = a0j01 (+) and cp60 = '" & Me.Text1.Text & "' and (cp79 <> 0 or (cp79 = 0 and cp75 <> 0)) order by cp09"
             StrSQLa = "select * from acc0j0, caseprogress, casepropertymap where a0j13 = '" & Me.Text1.Text & "' and cp09(+)=a0j01 and (cp79 <> 0 or (cp79 = 0 and cp75 <> 0)) and cpm01(+)=cp01 and cpm02(+)=cp10  order by cp09"
             rsA.CursorLocation = adUseClient
             rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
             If rsA.RecordCount > 0 Then
                 'Modify By Sindy 2020/7/31
                 'Me.Text6.Text = IIf("" & rsA.Fields("a0j04").Value = "020", "" & rsA("CPM04").Value, "" & rsA("CPM03").Value)
                 Me.Text6.Text = IIf("" & rsA.Fields("a0j04").Value = "000", "" & rsA("CPM03").Value, "" & rsA("CPM04").Value)
                 '2020/7/31 END
                 Me.Text6.Enabled = True: Text6.BackColor = Text1.BackColor
                 Me.Text6.Tag = Me.Text6.Text 'Added by Morgan 2020/10/14
                 rsA.MoveNext
                 If rsA.EOF = False Then
                     'Modify By Sindy 2020/7/31
                     'Me.Text7.Text = IIf("" & rsA.Fields("a0j04").Value = "020", "" & rsA("CPM04").Value, "" & rsA("CPM03").Value)
                     Me.Text7.Text = IIf("" & rsA.Fields("a0j04").Value = "000", "" & rsA("CPM03").Value, "" & rsA("CPM04").Value)
                     '2020/7/31 END
                     Me.Text7.Enabled = True: Text7.BackColor = Text1.BackColor
                     Me.Text7.Tag = Me.Text7.Text 'Added by Morgan 2020/10/14
                     
                     'Added by Morgan 2020/10/14 +第3,4案件性質
                     rsA.MoveNext
                     If rsA.EOF = False Then
                        Me.Text9.Text = IIf("" & rsA.Fields("a0j04").Value = "000", "" & rsA("CPM03").Value, "" & rsA("CPM04").Value)
                        Me.Text9.Enabled = True: Text9.BackColor = Text1.BackColor
                        Me.Text9.Tag = Me.Text9.Text
                        rsA.MoveNext
                        If rsA.EOF = False Then
                            Me.Text10.Text = IIf("" & rsA.Fields("a0j04").Value = "000", "" & rsA("CPM03").Value, "" & rsA("CPM04").Value)
                            Me.Text10.Enabled = True: Text10.BackColor = Text1.BackColor
                            Me.Text10.Tag = Me.Text10.Text
                        End If
                     End If
                     'end 2020/10/14
                 End If
             End If
             If rsA.State <> adStateClosed Then rsA.Close
             Set rsA = Nothing
         End If
         'End 2019/4/25
      End If
      'End
   Else
      MaskEdBox1.Mask = ""
      MaskEdBox1.Text = ""
      MaskEdBox1.Mask = DFormat
      Text2 = ""
      Text3 = ""
      Text5 = ""    '92.6.27 add by sonia
        Me.Text6.Text = ""
        Me.Text7.Text = ""
   End If
   adoacc0k0.Close
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

''*************************************************
''  抬頭列印
''
''*************************************************
'Private Sub PrintHead()
'   Select Case adoacc0k0.Fields("a0k11").Value
'      Case "1", "2"
'      Case Else
'         adoquery.CursorLocation = adUseClient
'         adoquery.Open "select * from acc080 where a0801 = '" & adoacc0k0.Fields("a0k11").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'         If adoquery.RecordCount <> 0 Then
'            Printer.FontSize = 18
'            Printer.CurrentX = 3500
'            Printer.CurrentY = 300
'            If IsNull(adoquery.Fields("a0802").Value) Then
'               Printer.Print ""
'            Else
'               Printer.Print adoquery.Fields("a0802").Value
'            End If
'            Printer.FontSize = 12
'            Printer.CurrentX = 3000
'            Printer.CurrentY = 700
'            If IsNull(adoquery.Fields("a0804").Value) Then
'               Printer.Print ReportSum(102)
'            Else
'               Printer.Print ReportSum(102) & adoquery.Fields("a0804").Value
'            End If
'            Printer.CurrentX = 3000
'            Printer.CurrentY = 1000
'            If IsNull(adoquery.Fields("a0813").Value) Then
'               Printer.Print ReportSum(103)
'            Else
'               Printer.Print ReportSum(103) & adoquery.Fields("a0813").Value
'            End If
'         End If
'         adoquery.Close
'   End Select
'   Printer.CurrentX = 8200
'   Printer.CurrentY = 1950
'   Printer.Print Mid(MaskEdBox2.Text, 1, 3)
'   Printer.CurrentX = 9200
'   Printer.CurrentY = 1950
'   Printer.Print Mid(MaskEdBox2.Text, 5, 2)
'   Printer.CurrentX = 10200
'   Printer.CurrentY = 1950
'   Printer.Print Mid(MaskEdBox2.Text, 8, 2)
'   Printer.CurrentX = 1250
'   Printer.CurrentY = 2200
'   '列印客戶編號
'   If IsNull(adoacc0k0.Fields("a0k03").Value) = False Then
'      Printer.Print adoacc0k0.Fields("a0k03").Value
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 8200
'   Printer.CurrentY = 2200
'   '列印收據號碼
'   Printer.Print adoacc0k0.Fields("a0k01").Value
'   Printer.CurrentX = 1250
'   Printer.CurrentY = 2450
'   '列印收據抬頭及統一編號(Optional)
'   If IsNull(adoacc0k0.Fields("a0k04").Value) = False Then
'      'Modify By Cheng 2002/01/17
''      Printer.Print adoacc0k0.Fields("a0k04").Value
'      Printer.Print Trim(adoacc0k0.Fields("a0k04").Value) & _
'                     IIf(Me.chk(0).Value, _
'                     IIf(IsNull(Me.adoacc0k0("CU11").Value), "", " (" & Me.adoacc0k0("CU11").Value & ")"), _
'                     "")
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 8200
'   Printer.CurrentY = 2450
'   '列印智權人員
'   If IsNull(adoacc0k0.Fields("a0k20").Value) = False Then
'      If adoquery.State = adStateOpen Then
'         adoquery.Close
'      End If
'      adoquery.Open "select cp12 from caseprogress where cp60 = '" & adoacc0k0.Fields("a0k01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adoquery.RecordCount <> 0 Then
'         If IsNull(adoquery.Fields("cp12").Value) Then
'            Printer.Print ""
'         Else
'            Printer.Print adoquery.Fields("cp12").Value
'         End If
'      Else
'         Printer.Print ""
'      End If
'      adoquery.Close
''      Printer.Print StaffDeptQuery(adoacc0k0.Fields("a0k20").Value)
'      Printer.CurrentX = 8650
'      Printer.CurrentY = 2450
'      Printer.Print adoacc0k0.Fields("a0k20").Value
'   Else
'      Printer.Print ""
'   End If
'   Printer.CurrentX = 1250
'   Printer.CurrentY = 2700
'   '列印地址
'   'Modify By Cheng 2002/01/17
'   '取消列印地址
''   If IsNull(adoacc0k0.Fields("cu31").Value) = False Then
''      Printer.Print adoacc0k0.Fields("cu31").Value
''   Else
''      Printer.Print ""
''   End If
'   If Text4 = MsgText(602) Then
'      Printer.CurrentX = 8200
'      Printer.CurrentY = 3600
'      Printer.Print "***  補開  ***"
'   End If
'   'Modify by Morgan 2005/6/24
'   'Printer.CurrentX = 10200
'   Printer.CurrentX = 10200 + 350
'
'   Printer.CurrentY = 6800
'   Printer.Print intTimes
'End Sub

''*************************************************
'' 合計位置
''
''*************************************************
'Private Sub PrintSum()
'Dim lngTotal As Long
'Dim intLength As Integer
''Add by Morgan 2005/1/25
'Dim bolNewReceipt As Boolean
''Modify by Morgan 2005/6/22 只剩2公司用舊收據
''If Val(Text5) > 2 Then bolNewReceipt = True
'bolNewReceipt = True
''2005/8/15 CANCEL BY SONIA 全部都用新收據
''If Val(Text5) = 2 Then bolNewReceipt = False
''2005/8/15 END
''2005/6/22 end
''2005/1/25 end
'
'   'Modify by Morgan 2004/8/12   0 也要印
'   'strAmount1 = Format(lngAmount1, DDollar)
'   If lngAmount1 = 0 Then
'      strAmount1 = "0"
'   Else
'      strAmount1 = Format(lngAmount1, DDollar)
'   End If
'   'Modify end
'
'   intLength = Printer.TextWidth(strAmount1)
'   Printer.CurrentX = 5900 - intLength
'   'Modify by Morgan 2005/1/26 改用新收據
'   'Printer.CurrentY = 5500
'   If bolNewReceipt = True Then
'      Printer.CurrentY = 5000
'   Else
'      Printer.CurrentY = 5500
'   End If
'   '2005/1/26 end
'
'   Printer.Print strAmount1
'   If lngAmount2 = 0 Then
'      strAmount2 = "0"
'   Else
'      strAmount2 = Format(lngAmount2, DDollar)
'   End If
'   intLength = Printer.TextWidth(strAmount2)
'   Printer.CurrentX = 7600 - intLength
'   'Modify by Morgan 2005/1/25 改用新收據
'   'Printer.CurrentY = 5500
'   If bolNewReceipt = True Then
'      Printer.CurrentY = 5000
'   Else
'      Printer.CurrentY = 5500
'   End If
'   '2005/1/25 end
'
'   Printer.Print strAmount2
'   lngTotal = lngAmount1 + lngAmount2
'   'Add by Morgan 2005/1/25 改用新收據
'   If bolNewReceipt = True Then
'         Printer.FontSize = 18
'         intLength = Printer.TextWidth("NT$" & Format(lngTotal, DDollar))
'         Printer.CurrentX = 7400 - intLength
'         Printer.CurrentY = 5650
'         Printer.Print "NT$" & Format(lngTotal, DDollar)
'         Printer.FontSize = 12
'   Else
'   '2005/1/25 end
'      For intLength = 1 To 7
'         Printer.CurrentX = 7900 - intLength * 820
'         Printer.CurrentY = 6000
'         If intLength > Len(CStr(lngTotal)) Then
'            Printer.Print ShowNumberWord(0)
'         Else
'            Printer.Print ShowNumberWord(Mid(CStr(lngTotal), Len(CStr(lngTotal)) - intLength + 1, 1))
'         End If
'      Next intLength
'   End If
'End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1 = "E"
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = CFDate(ACDate(ServerDate))
   MaskEdBox2.Mask = DFormat
   Text3 = ""
   Text2 = ""
   Text5 = ""
    Me.Text6.Text = "": Me.Text6.Enabled = False
    Me.Text7.Text = "": Me.Text7.Enabled = False
    Me.Text9.Text = "": Me.Text9.Enabled = False 'Added by Morgan 2020/10/14
    Me.Text10.Text = "": Me.Text10.Enabled = False 'Added by Morgan 2020/10/14
    
   Text1.SetFocus
   'Add By Cheng 2002/01/17
   Me.chk(0).Value = vbUnchecked
   'Modify by Morgan 2004/8/12
   'Me.chk(1).Value = vbUnchecked
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
   CloseIme
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   'Add by Morgan 2006/2/6 列印日期檢查
   If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
      MsgBox "列印日期格式錯誤！", vbExclamation
      MaskEdBox2.SetFocus
      Exit Function
   End If
   'ADD BY SONIA 2015/9/17 瑞婷說也要加檢查工作日
   If ChkWorkDay(FCDate(MaskEdBox2.Text) + 19110000) = False Then
      MsgBox Label2 & "請輸入工作日！", vbExclamation, "日期錯誤！"
      MaskEdBox2.SetFocus
      Exit Function
   End If
   'END 2015/9/17
   
   'Added by Morgan 2020/10/14
   If Text6.Enabled = True Then
      If Trim(Text6) = "" Then
         MsgBox "案件性質名稱1不可空白！", vbExclamation
         Text6.SetFocus
         Exit Function
      End If
   End If
   If Text7.Enabled = True Then
      If Trim(Text7) = "" Then
         MsgBox "案件性質名稱2不可空白！", vbExclamation
         Text7.SetFocus
         Exit Function
      End If
   End If
   If Text9.Enabled = True Then
      If Trim(Text9) = "" Then
         MsgBox "案件性質名稱3不可空白！", vbExclamation
         Text9.SetFocus
         Exit Function
      End If
   End If
   If Text10.Enabled = True Then
      If Trim(Text10) = "" Then
         MsgBox "案件性質名稱4不可空白！", vbExclamation
         Text10.SetFocus
         Exit Function
      End If
   End If
   'end 2020/10/14
   
   If Text1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text4 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

'*************************************************
'  以總收文號查詢案件名稱
'
'Modify By Cheng 2002/01/17
'多加客戶案件案號欄位
'92.3.21 add by sonia
'列印收據為商標案件時 中英文案件名稱同時列印 , 此function原為共用置於acc_fun,
'因只有收據改, 故改寫在此
'*************************************************
Public Function CaseNameQuery(InputNo As String, InputSelect As Integer, Optional strCustCaseNo As String) As String
Dim adocaseprogress As New ADODB.Recordset

   'Add By Cheng 2002/01/17
   strCustCaseNo = ""
   
   If adocaseprogress.State = adStateOpen Then
      adocaseprogress.Close
   End If
   adocaseprogress.CursorLocation = adUseClient
   '92.3.21 MODIFY BY SONIA 商標案件時 中英文案件名稱同時列印, 此處加 cp01 供下方判斷
   '2010/5/21 MODIFY BY SONIA 杜副總說顧問案件不印案件名稱
   adocaseprogress.Open "select pa05, pa06, pa07, pa48, cp01 from caseprogress, patent where cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and cp09 = '" & InputNo & "' " & _
                        "union select tm05, tm06, tm07, tm35, cp01 from caseprogress, trademark where cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and cp09 = '" & InputNo & "' " & _
                        "union select lc05, lc06, lc07, lc17, cp01 from caseprogress, lawcase where cp01 = lc01 and cp02 = lc02 and cp03 = lc03 and cp04 = lc04 and cp09 = '" & InputNo & "' " & _
                        "union select '顧問', '', '', '', cp01 from caseprogress, hirecase where cp01 = hc01 and cp02 = hc02 and cp03 = hc03 and cp04 = hc04 and cp09 = '" & InputNo & "' " & _
                        "union select sp05, sp06, sp07, sp29, cp01 from caseprogress, servicepractice where cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and cp09 = '" & InputNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adocaseprogress.RecordCount <> 0 Then
      Select Case InputSelect
         Case 1
            If IsNull(adocaseprogress.Fields(0).Value) Then
               CaseNameQuery = MsgText(601)
            Else
               CaseNameQuery = adocaseprogress.Fields(0).Value
               '92.3.21 add by sonia
               If adocaseprogress.Fields(4).Value = "T" Or adocaseprogress.Fields(4).Value = "TF" Or adocaseprogress.Fields(4).Value = "FCT" Or adocaseprogress.Fields(4).Value = "CFT" Then
                  If adocaseprogress.Fields(1).Value <> "" Then
                     CaseNameQuery = CaseNameQuery & adocaseprogress.Fields(1).Value
                  End If
               End If
               '92.3.21 end
            End If
         Case 2
            If IsNull(adocaseprogress.Fields(1).Value) Then
               CaseNameQuery = MsgText(601)
            Else
               CaseNameQuery = adocaseprogress.Fields(1).Value
            End If
         Case 3
            If IsNull(adocaseprogress.Fields(2).Value) Then
               CaseNameQuery = MsgText(601)
            Else
               CaseNameQuery = adocaseprogress.Fields(2).Value
            End If
      End Select
      'Add By Cheng 2002/01/17
      strCustCaseNo = "" & adocaseprogress.Fields(3).Value
   Else
      CaseNameQuery = MsgText(601)
      'Add By Cheng 2002/01/17
      strCustCaseNo = ""
   End If
   adocaseprogress.Close
End Function

'Add By Cheng 2003/06/02
'設定列印字數
Private Function ReOrgPrintTemp(strPrintTemp As String, intCnt As Integer) As String
Dim ii As Integer

ReOrgPrintTemp = ""
For ii = 1 To Len(strPrintTemp)
    ReOrgPrintTemp = ReOrgPrintTemp & Mid(strPrintTemp, ii, 1)
    If Printer.TextWidth(ReOrgPrintTemp) > Printer.TextWidth(String(intCnt, "　")) Then
        ReOrgPrintTemp = Left(ReOrgPrintTemp, Len(ReOrgPrintTemp) - 1)
        Exit For
    End If
Next ii
End Function

Private Sub Text6_GotFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
OpenIme
End Sub

Private Sub Text6_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
CloseIme
End Sub

Private Sub Text7_GotFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
OpenIme
End Sub

Private Sub Text7_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
CloseIme
End Sub

'Added by Lydia 2016/04/13 點選呼叫"收據抬頭修改"
Private Sub Command2_Click()
   If Text1.Text <> "" And MaskEdBox1.Text <> "" And Text3.Text <> "" Then
      strItemNo = Trim(Text1.Text)
      strTitle = Me.Name
      If Mid(strItemNo, 1, 1) = "E" Then
         Set Frmacc1140.TmpFrmacc1420 = Me
         tool14_enabled
         Frmacc1140.Show
         Me.Enabled = False
      Else
         MsgBox "請輸入收據號碼..."
         strItemNo = ""
         strTitle = ""
      End If
   Else
      MsgBox "請先輸入收據號碼..."
   End If
End Sub
'end 2016/04/13

'Added by Morgan 2020/10/14
'更新帳款類別
Private Function UpdateItemDesc(pA0K01 As String, pA0K33 As String) As Boolean
   Dim bolInTrans As Boolean
      
On Error GoTo ErrHnd
   
   cnnConnection.BeginTrans
   bolInTrans = True
   
   If pA0K33 = "Y" Then
      If Text6.Text <> Text6.Tag Then
         strSql = "update acc0j0 set a0j22='" & ChgSQL(Text6.Text) & "' where a0j13='" & pA0K01 & "' and a0j22='" & ChgSQL(Text6.Tag) & "'"
         cnnConnection.Execute strSql, intI
      End If
      If Text7.Enabled = True And Text7.Text <> Text7.Tag Then
         strSql = "update acc0j0 set a0j22='" & ChgSQL(Text7.Text) & "' where a0j13='" & pA0K01 & "' and a0j22='" & ChgSQL(Text7.Tag) & "'"
         cnnConnection.Execute strSql, intI
      End If
      If Text9.Enabled = True And Text9.Text <> Text9.Tag Then
         strSql = "update acc0j0 set a0j22='" & ChgSQL(Text9.Text) & "' where a0j13='" & pA0K01 & "' and a0j22='" & ChgSQL(Text9.Tag) & "'"
         cnnConnection.Execute strSql, intI
      End If
      If Text10.Enabled = True And Text10.Text <> Text10.Tag Then
         strSql = "update acc0j0 set a0j22='" & ChgSQL(Text10.Text) & "' where a0j13='" & pA0K01 & "' and a0j22='" & ChgSQL(Text10.Tag) & "'"
         cnnConnection.Execute strSql, intI
      End If
   Else
      strSql = "update acc0k0 set a0k33='Y' where a0k01='" & pA0K01 & "' and a0k33 is null"
      cnnConnection.Execute strSql, intI
      
      strSql = "update acc0j0 set a0j22='" & ChgSQL(Text6.Text) & "',a0j25=1 where a0j13='" & pA0K01 & "' and a0j01=(select min(a0j01) from acc0j0 where a0j13='" & pA0K01 & "' and a0j22 is null)"
      cnnConnection.Execute strSql, intI
      
      If Text7.Enabled = True Then
         strSql = "update acc0j0 set a0j22='" & ChgSQL(Text7.Text) & "',a0j25=2 where a0j13='" & pA0K01 & "' and a0j01=(select min(a0j01) from acc0j0 where a0j13='" & pA0K01 & "' and a0j22 is null)"
         cnnConnection.Execute strSql, intI
      End If
      
      If Text9.Enabled = True Then
         strSql = "update acc0j0 set a0j22='" & ChgSQL(Text9.Text) & "',a0j25=3 where a0j13='" & pA0K01 & "' and a0j01=(select min(a0j01) from acc0j0 where a0j13='" & pA0K01 & "' and a0j22 is null)"
         cnnConnection.Execute strSql, intI
      End If
      
      If Text10.Enabled = True Then
         strSql = "update acc0j0 set a0j22='" & ChgSQL(Text9.Text) & "',a0j25=4 where a0j13='" & pA0K01 & "' and a0j01=(select min(a0j01) from acc0j0 where a0j13='" & pA0K01 & "' and a0j22 is null)"
         cnnConnection.Execute strSql, intI
      End If
   End If
   cnnConnection.CommitTrans
   UpdateItemDesc = True
   Exit Function
   
ErrHnd:
   If bolInTrans Then cnnConnection.RollbackTrans
   MsgBox Err.Description, vbExclamation, "帳款類別更新失敗!!!"
   
End Function
