VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc1410 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "收據列印"
   ClientHeight    =   6840
   ClientLeft      =   40
   ClientTop       =   310
   ClientWidth     =   5780
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   5780
   Begin VB.CheckBox Check3 
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
      Left            =   3900
      TabIndex        =   45
      Top             =   5040
      Value           =   1  '核取
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.TextBox txtNote 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   585
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   44
      Text            =   "產生收據中，暫時不要使用Word..."
      Top             =   2790
      Visible         =   0   'False
      Width           =   5610
   End
   Begin VB.CheckBox Check2 
      Caption         =   "單張列印"
      Height          =   195
      Left            =   390
      TabIndex        =   43
      Top             =   5430
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   120
      TabIndex        =   41
      Top             =   5670
      Width           =   4755
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   120
         Width           =   705
      End
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
         TabIndex        =   21
         Top             =   120
         Width           =   705
      End
      Begin VB.Label Label5 
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
         TabIndex        =   42
         Top             =   180
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "智慧所收據設定"
      Height          =   660
      Left            =   120
      TabIndex        =   37
      Top             =   4110
      Width           =   5415
      Begin VB.OptionButton Option2 
         Caption         =   "影印機"
         Height          =   180
         Index           =   1
         Left            =   4170
         TabIndex        =   16
         Top             =   390
         Width           =   915
      End
      Begin VB.OptionButton Option2 
         Caption         =   "列表機"
         Height          =   180
         Index           =   0
         Left            =   4170
         TabIndex        =   15
         Top             =   150
         Width           =   915
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   14
         Top             =   240
         Width           =   3240
      End
      Begin VB.Label Label4 
         Caption         =   "印表機"
         Height          =   315
         Left            =   105
         TabIndex        =   38
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.TextBox TextY 
      Height          =   285
      Left            =   1860
      TabIndex        =   18
      Top             =   5100
      Width           =   705
   End
   Begin VB.TextBox TextX 
      Height          =   285
      Left            =   1860
      TabIndex        =   17
      Top             =   4800
      Width           =   705
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "測試收據表格"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3930
      Style           =   1  '圖片外觀
      TabIndex        =   19
      Top             =   4710
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   2010
      Style           =   2  '單純下拉式
      TabIndex        =   1
      Top             =   570
      Width           =   3450
   End
   Begin VB.TextBox Text2 
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
      Index           =   2
      Left            =   1575
      MaxLength       =   9
      TabIndex        =   8
      Top             =   2430
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "重新列印"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   390
      TabIndex        =   7
      Top             =   2160
      Width           =   1125
   End
   Begin VB.OptionButton Option1 
      Caption         =   "一般列印"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   390
      TabIndex        =   2
      Top             =   1140
      Value           =   -1  'True
      Width           =   1125
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
      Left            =   2790
      TabIndex        =   13
      Top             =   3690
      Value           =   1  '核取
      Width           =   2355
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
      Left            =   375
      TabIndex        =   12
      Top             =   3690
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.TextBox Text2 
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
      Index           =   1
      Left            =   3495
      MaxLength       =   9
      TabIndex        =   6
      Top             =   1770
      Width           =   1575
   End
   Begin VB.TextBox Text2 
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
      Index           =   0
      Left            =   1575
      MaxLength       =   9
      TabIndex        =   5
      Top             =   1770
      Width           =   1575
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
      Height          =   315
      Left            =   900
      MaxLength       =   1
      TabIndex        =   0
      Top             =   180
      Width           =   405
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
      TabIndex        =   25
      Top             =   6300
      Width           =   4692
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   315
      Left            =   1575
      TabIndex        =   3
      Top             =   1410
      Width           =   1575
      _ExtentX        =   2769
      _ExtentY        =   564
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3495
      TabIndex        =   4
      Top             =   1410
      Width           =   1575
      _ExtentX        =   2769
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
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   315
      Left            =   1575
      TabIndex        =   9
      Top             =   2790
      Width           =   1575
      _ExtentX        =   2769
      _ExtentY        =   564
      _Version        =   393216
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
   Begin MSMask.MaskEdBox MaskEdBox4 
      Height          =   315
      Left            =   1575
      TabIndex        =   10
      Top             =   3150
      Width           =   1575
      _ExtentX        =   2769
      _ExtentY        =   564
      _Version        =   393216
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
   Begin MSMask.MaskEdBox MaskEdBox5 
      Height          =   315
      Left            =   3495
      TabIndex        =   11
      Top             =   3150
      Width           =   1575
      _ExtentX        =   2769
      _ExtentY        =   564
      _Version        =   393216
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
      Height          =   405
      Left            =   3930
      TabIndex        =   20
      Top             =   5130
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "縱軸偏移值(Y)：　　　　　　(單位公分)"
      Height          =   180
      Index           =   1
      Left            =   375
      TabIndex        =   40
      Top             =   5160
      Width           =   3240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "橫軸偏移值(X)：　　　　　　(單位公分)"
      Height          =   180
      Index           =   0
      Left            =   375
      TabIndex        =   39
      Top             =   4860
      Width           =   3240
   End
   Begin VB.Label LblCmpName 
      BackStyle       =   0  '透明
      Caption         =   "公司名稱"
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
      Left            =   1350
      TabIndex        =   36
      Top             =   210
      Width           =   3795
   End
   Begin VB.Label Label10 
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
      Left            =   150
      TabIndex        =   35
      Top             =   630
      Width           =   2805
   End
   Begin VB.Label Label1 
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
      Index           =   4
      Left            =   615
      TabIndex        =   34
      Top             =   2820
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3225
      TabIndex        =   33
      Top             =   3180
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "列印時間"
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
      Index           =   3
      Left            =   615
      TabIndex        =   32
      Top             =   3180
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收據號碼                       以後(含)"
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
      Index           =   2
      Left            =   615
      TabIndex        =   31
      Top             =   2460
      Width           =   4440
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2565
      Left            =   240
      Top             =   1050
      Width           =   5010
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
      Index           =   1
      Left            =   615
      TabIndex        =   30
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3255
      TabIndex        =   29
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
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
      Left            =   150
      TabIndex        =   28
      Top             =   210
      Width           =   705
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   135
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3255
      TabIndex        =   27
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label1 
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
      Index           =   0
      Left            =   615
      TabIndex        =   26
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc1410"
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
Dim strSql As String
Dim strNo As String
Dim lngAmount1 As Long
Dim lngAmount2 As Long
Dim strAmount1 As String
Dim strAmount2 As String
Dim intLength As Integer
Dim intCounter As Integer
Dim intTimes As Integer
'Add by Morgan 2004/8/12
Dim strTemp As String
Dim strCustCaseNo As String '客戶案件案號
Dim m_FixNo As Integer   '2010/2/12 add by sonia 修法次數
Dim strPrinter As String, strPrinter2 As String, m_FileName As String 'Add By Sindy 2020/3/23
Dim SeekPrintL As Integer, SeekPrint As Integer 'Add By Sindy 2020/7/14
Public ProState As String '權限: 1.全所 2.該所 Add By Sindy 2021/5/21
Dim m_sqlST06 As String 'Add By Sindy 2021/5/21
'Add By Sindy 2022/1/26
Dim m_FileName2 As String, m_FileName3 As String
Dim strWordFileName As String
Dim strPType As String
'2022/1/26 END
Dim bolRunWord As Boolean 'Added by Lydia 2023/11/13
Dim bolRunReceipt As Boolean 'Added by Lydia 2023/11/20
Dim m_AttachPath As String 'Add By Sindy 2025/9/18


Private Sub Command1_Click()
Dim ii As Integer, bolChk As Boolean
  
   Screen.MousePointer = vbHourglass
   
   bolRunWord = False 'Added by Lydia 2023/11/13
   bolRunReceipt = False 'Added by Lydia 2023/11/20
   
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
   
   If Option1(0).Value = True Then '一般列印
      PrintDoc
   Else
      PrintDocX '重新列印
   End If

   'Added by Lydia 2023/11/13 限財務人員和電腦中心人員執行
   If bolRunWord = True And ChkST06(0).Value = 1 And (Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M31") Then
      PrintDebitList
   End If
   'end 2023/11/13
   'Added by Lydia 2023/11/20 因為會清空畫面條件，所以移到外層
   If bolRunReceipt = True Then
      FormClear
   End If
   'end 2023/11/20
   EndOfficeAp 'Added by Morgan 2025/9/10 印完要清除物件，否則印表機不會變
   Screen.MousePointer = vbDefault

End Sub

Private Sub Command2_Click()
   PUB_RestorePrinter Combo2 'Add By Sindy 2020/7/14 切換印表機
   
   Call PUB_PrintCaseReceiptTableMain(TextX, TextY, IIf(Option2(0).Value = True, 0, 1), 1, True)
   Call PUB_PrintCaseReceiptTableMain(TextX, TextY, IIf(Option2(0).Value = True, 0, 1), 2, False)
   Printer.EndDoc
   
   PUB_RestorePrinter strPrinter2 'Add By Sindy 2020/7/14 還原印表機
   'MsgBox "列印結束 !"
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
   
'Modified by Lydia 2023/11/13 表單初始化
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   'Removed by Morgan 2013/4/30 改單線固定(調整大小不用再設定)
'   'Me.Width = 5280
'   'Me.Height = 3435
'   'end 2013/4/30
'
'   'Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   MoveFormToCenter Me
'
'   Image1 = LoadPicture(strBackPicPath4)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   PUB_InitForm Me, 5850, 7260, strBackPicPath4, lngWidth, lngHeight
'end 2023/11/13

'   Text2(0) = MsgText(802)
'   Text2(1) = MsgText(802)
   'MODIFY BY SONIA 2014/4/9 預設上一個工作日至當天
   'MaskEdBox1.Text = CFDate(strSrvDate(2))
   MaskEdBox1.Text = CFDate(TransDate(PUB_GetWorkDay1(strSrvDate(1) - 1, 1), 1))
   '2014/4/9 END
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Text = CFDate(strSrvDate(2))
   MaskEdBox2.Mask = DFormat
   
   'Added by Morgan 2013/4/30
   MaskEdBox3.Mask = DFormat
   MaskEdBox4.Mask = Tformat
   MaskEdBox5.Mask = Tformat
   'end 2013/4/30
   
   LblCmpName.Caption = "" 'Add By Sindy 2020/3/20
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
'   ChkUnPrintData 'Add by Morgan 2011/9/20

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
   m_FileName3 = "$$智慧所收據-雙.doc"
   'Modify By Sindy 2025/9/18 增加收據資料夾存放
   If Dir(m_AttachPath & m_FileName3) <> "" Then
      Kill m_AttachPath & m_FileName3
   End If
   Call PUB_GetSampleFile(m_FileName3, "M31-000010-1-00", , m_AttachPath)
   '2022/1/26 END
   
   'Add By Sindy 2020/7/14 設定印表機改呼叫公用函數,原程式移除
   SeekPrintL = Printer.Orientation
   PUB_SetPrinter Me.Name, Combo2, strPrinter2, , SeekPrint, TextX, TextY
   
'   If Pub_StrUserSt03 = "M51" Then
'      Command2.Visible = True
'   End If
End Sub

''Modify by Morgan 2011/9/20 改寫成函數方便公用
'Private Sub ChkUnPrintData()
'Dim s_printmsg As String   '2013/9/17 add by sonia
'
'   'Add By Sindy 2010/4/29 檢查待列印收據張數
'   Text3 = ""
'   strSql = "select a0k11,count(*) from acc0k0 " & _
'               "where (a0k32='Y') " & _
'               "and (a0k19=0 or a0k19 is null) group by a0k11 order by a0k11 "
'   intI = 1
'   Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      adoRecordset.MoveFirst
'      Do While Not adoRecordset.EOF
'         If Text3 = "" Then Text3 = "尚有待列印收據："
'         Text3 = Text3 & adoRecordset.Fields(0) & "公司: " & adoRecordset.Fields(1) & " 張；"
'         adoRecordset.MoveNext
'      Loop
'   End If
'   '2010/4/29 End
'
'   '2013/9/17 ADD BY SONIA 檢查是否有應列印但未列印的收據日期
'   '2013/10/3 modify by sonia 加顯示公司別a0k11
'   s_printmsg = ""
'   strSql = "select sqldatet(a0k02),a0k11,count(*) from acc0k0 " & _
'               "where a0k02>=920201 and a0k02<" & strSrvDate(2) & " and a0k32 is null and a0k10 is null " & _
'               "and nvl(a0k09,0)=0 and nvl(a0k19,0)=0 group by a0k02,a0k11 order by a0k02,a0k11 "
'   intI = 1
'   Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      adoRecordset.MoveFirst
'      Do While Not adoRecordset.EOF
'         If s_printmsg = "" Then s_printmsg = "尚有過去未列印收據："
'         '2013/10/3 modify by sonia 加顯示公司別a0k11
'         s_printmsg = s_printmsg & adoRecordset.Fields(0) & " (" & adoRecordset.Fields(1) & "公司) : " & adoRecordset.Fields(2) & " 張；"
'         adoRecordset.MoveNext
'      Loop
'   End If
'   If s_printmsg <> "" Then
'      MsgBox s_printmsg
'   End If
'   '2013/9/17 END
'End Sub

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
   Set Frmacc1410 = Nothing
End Sub

Private Sub Combo2_Click()
   Option2(0).Value = True
   If InStr(Combo2.Text, "影印機") > 0 Then
      Option2(1).Value = True
   End If
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
'            '列印公司名稱
'            If IsNull(adoquery.Fields("a0802").Value) Then
'               Printer.Print ""
'            Else
'               Printer.Print adoquery.Fields("a0802").Value
'            End If
'            Printer.FontSize = 12
'            Printer.CurrentX = 3000
'            Printer.CurrentY = 700
'            '列印發票地址
'            If IsNull(adoquery.Fields("a0804").Value) Then
'               Printer.Print ReportSum(102)
'            Else
'               Printer.Print ReportSum(102) & adoquery.Fields("a0804").Value
'            End If
'            Printer.CurrentX = 3000
'            Printer.CurrentY = 1000
'            '列印電話
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
'   '列印收據日期-年份
'   Printer.Print Mid(CFDate(adoacc0k0.Fields("a0k02").Value), 1, 3)
'   Printer.CurrentX = 9200
'   Printer.CurrentY = 1950
'   '列印收據日期-月份
'   Printer.Print Mid(CFDate(adoacc0k0.Fields("a0k02").Value), 5, 2)
'   Printer.CurrentX = 10200
'   Printer.CurrentY = 1950
'   '列印收據日期-日
'   Printer.Print Mid(CFDate(adoacc0k0.Fields("a0k02").Value), 8, 2)
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
'   '列印收文號
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
'   '列印智權人員代號
'   If IsNull(adoacc0k0.Fields("a0k20").Value) = False Then
'      If adoquery.State = adStateOpen Then
'         adoquery.Close
'      End If
'      'Modified by Morgan 2011/10/31
'      'adoquery.Open "select cp12 from caseprogress where cp60 = '" & adoacc0k0.Fields("a0k01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'      adoquery.Open "select cp12 from acc0j0,caseprogress where a0j13 = '" & adoacc0k0.Fields("a0k01").Value & "' and cp09(+)=a0j01 order by cp05,cp09", adoTaie, adOpenStatic, adLockReadOnly
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
'   '列印聯絡地址
'   'Modify By Cheng 2002/01/16
'   '取消列印
''   If IsNull(adoacc0k0.Fields("cu31").Value) = False Then
''      Printer.Print adoacc0k0.Fields("cu31").Value
''   Else
''      Printer.Print ""
''   End If
'   'Modify by Morgan 2005/6/24
'   'Printer.CurrentX = 10200
'   Printer.CurrentX = 10200 + 350
'
'   Printer.CurrentY = 6800
'   Printer.Print intTimes
'End Sub
'
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
''If Val(Text1) > 2 Then bolNewReceipt = True
'bolNewReceipt = True
''2005/8/15 CANCEL BY SONIA 全部都用新收據
''If Val(Text1) = 2 Then bolNewReceipt = False
''2005/8/15 END
''2005/6/22 end
''2005/1/25 end
'
'   If lngAmount1 = 0 Then
'      strAmount1 = "0"
'   Else
'      strAmount1 = Format(lngAmount1, DDollar)
'   End If
'   intLength = Printer.TextWidth(strAmount1)
'   Printer.CurrentX = 5900 - intLength
'   'Modify by Morgan 2005/1/26 非1,2公司改用新收據
'   'Printer.CurrentY = 5500
'   If bolNewReceipt = True Then
'      Printer.CurrentY = 5000
'   Else
'      Printer.CurrentY = 5500
'   End If
'   '2005/1/26 end
'   Printer.Print strAmount1
'   If lngAmount2 = 0 Then
'      strAmount2 = "0"
'   Else
'      strAmount2 = Format(lngAmount2, DDollar)
'   End If
'   intLength = Printer.TextWidth(strAmount2)
'   Printer.CurrentX = 7600 - intLength
'   'Modify by Morgan 2005/1/26 改用新收據
'   'Printer.CurrentY = 5500
'   If bolNewReceipt = True Then
'      Printer.CurrentY = 5000
'   Else
'      Printer.CurrentY = 5500
'   End If
'   '2005/1/26 end
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
   Text1 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text1.SetFocus
   'Add By Cheng 2002/01/17
   Me.Text2(0).Text = Empty
   Me.Text2(1).Text = Empty
   Me.chk(0).Value = vbUnchecked
   'Modify by Morgan 2004/8/12
   'Me.chk(1).Value = vbUnchecked
   
   'Added by Morgan 2013/5/1
   Text2(2).Text = Empty
   MaskEdBox3.Mask = ""
   MaskEdBox3.Text = ""
   MaskEdBox3.Mask = DFormat
   MaskEdBox4.Mask = ""
   MaskEdBox4.Text = ""
   MaskEdBox4.Mask = Tformat
   MaskEdBox5.Mask = ""
   MaskEdBox5.Text = ""
   MaskEdBox5.Mask = Tformat
   Option1(0).Value = True
   'end 2013/5/1
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   MaskEdBox2.Mask = MsgText(601)
   MaskEdBox2.Text = MaskEdBox1.Text
   MaskEdBox2.Mask = DFormat
End Sub

Private Sub Option1_Click(Index As Integer)
   If Index = 0 Then
      MaskEdBox1.Enabled = True
      MaskEdBox2.Enabled = True
      Text2(0).Enabled = True
      Text2(1).Enabled = True
      
      Text2(2).Enabled = False
      MaskEdBox3.Enabled = False
      MaskEdBox4.Enabled = False
      MaskEdBox5.Enabled = False
   Else
      MaskEdBox1.Enabled = False
      MaskEdBox2.Enabled = False
      Text2(0).Enabled = False
      Text2(1).Enabled = False
      
      Text2(2).Enabled = True
      MaskEdBox3.Enabled = True
      MaskEdBox4.Enabled = True
      MaskEdBox5.Enabled = True
      
      Text2(2).Text = ""
      MaskEdBox3.Mask = ""
      MaskEdBox3.Text = ""
      MaskEdBox3.Mask = DFormat
      MaskEdBox4.Mask = ""
      MaskEdBox4.Text = ""
      MaskEdBox4.Mask = Tformat
      MaskEdBox5.Mask = ""
      MaskEdBox5.Text = ""
      MaskEdBox5.Mask = Tformat
      Text2(2).SetFocus
   End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   LblCmpName = "" 'Add By Sindy 2020/3/20
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from acc080 where a0801 = '" & Text1 & "' and a0801<>'J'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount = 0 Then
      MsgBox MsgText(188) & Label2, , MsgText(5)
      adoquery.Close
      Cancel = True
      Text1.SetFocus
      TextInverse Text1
      Exit Sub
   'Modify By Sindy 2020/3/20
   Else
      LblCmpName = A0802Query(Text1)
   '2020/3/20 END
   End If
   adoquery.Close
End Sub

Private Sub Text2_Change(Index As Integer)
   If Index = 2 Then
      If Len(Text2(Index)) = 9 Then
         strExc(0) = "select to_char(a0k14,'yyyymmdd'),to_char(a0k14,'hh24miss'),to_char(sysdate,'hh24miss'),a0k11 from acc0k0 where a0k01='" & Text2(2) & "' and a0k14 is not null and a0k11<>'J'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            MaskEdBox3.Mask = ""
            MaskEdBox3.Text = Format(TransDate(RsTemp(0), 1), DFormat)
            If Len(MaskEdBox3.Text) = 8 Then MaskEdBox3.Text = 0 & MaskEdBox3.Text 'Add By Sindy 2024/9/4 不補足,後面Run Mask,欄位值又會清空
            MaskEdBox3.Mask = DFormat
            MaskEdBox4.Mask = ""
            MaskEdBox4.Text = Format(RsTemp(1), Tformat)
            If Len(MaskEdBox4.Text) = 7 Then MaskEdBox4.Text = 0 & MaskEdBox4.Text 'Add By Sindy 2024/9/4 不補足,後面Run Mask,欄位值又會清空
            MaskEdBox4.Mask = Tformat
            If RsTemp(0) < strSrvDate(1) Then
               MaskEdBox5.Mask = ""
               MaskEdBox5.Text = Format("235959", Tformat)
               MaskEdBox5.Mask = Tformat
            Else
               MaskEdBox5.Mask = ""
               MaskEdBox5.Text = Format(RsTemp(2), Tformat)
               If Len(MaskEdBox5.Text) = 7 Then MaskEdBox5.Text = 0 & MaskEdBox5.Text 'Add By Sindy 2024/9/4 不補足,後面Run Mask,欄位值又會清空
               MaskEdBox5.Mask = Tformat
            End If
            Text1.Text = RsTemp(3)
         End If
      End If
   End If
End Sub

Private Sub Text2_GotFocus(Index As Integer)
   TextInverse Text2(Index)
   CloseIme
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   Me.Text2(Index).Text = UCase(Me.Text2(Index).Text)
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean

   'Removed by Morgan 2013/8/14 改為必要條件
   'If Text1 <> MsgText(601) Then
   '   FormCheck = True
   '   Exit Function
   'End If

   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If Text2(0) <> MsgText(601) And Text2(0) <> MsgText(802) Then
      FormCheck = True
      Exit Function
   End If
   If Text2(1) <> MsgText(601) And Text2(1) <> MsgText(802) Then
      FormCheck = True
      Exit Function
   End If
   
   FormCheck = False
End Function

'Added by Morgan 2013/4/30
Private Function FormCheck1() As Boolean

   If Text1 = "" Then
      MsgBox "請輸入公司別", vbExclamation
      Text1.SetFocus
      Exit Function
   End If
   
   If MaskEdBox3.Text = MsgText(29) Then
      MsgBox "請輸入列印日期！", vbExclamation
      MaskEdBox3.SetFocus
      Exit Function
   ElseIf MaskEdBox3.Text <> CFDate(strSrvDate(2)) Then
      If MsgBox("列印日期不是當天，是否確定要繼續？", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
         MaskEdBox3.SetFocus
         Exit Function
      End If
   End If
   If MaskEdBox4.Text = "__:__:__" Then
      MsgBox "請輸入列印時間(起)！", vbExclamation
      MaskEdBox4.SetFocus
      Exit Function
   End If
   If MaskEdBox5.Text = "__:__:__" Then
      MsgBox "請輸入列印時間(迄)！", vbExclamation
      MaskEdBox5.SetFocus
      Exit Function
   End If
   
   FormCheck1 = True
End Function

''*************************************************
''  以總收文號查詢案件名稱
''
''Modify By Cheng 2002/01/17
''多加客戶案件案號欄位
''92.3.21 add by sonia
''列印收據為商標案件時 中英文案件名稱同時列印 , 此function原為共用置於acc_fun,
''因只有收據改, 故改寫在此
''*************************************************
'Public Function CaseNameQuery(InputNo As String, InputSelect As Integer, Optional strCustCaseNo As String) As String
'Dim adocaseprogress As New ADODB.Recordset
'
'   'Add By Cheng 2002/01/17
'   strCustCaseNo = ""
'
'   adocaseprogress.CursorLocation = adUseClient
'   '92.3.21 MODIFY BY SONIA 商標案件時 中英文案件名稱同時列印, 此處加 cp01 供下方判斷
'   '2010/5/21 MODIFY BY SONIA 杜副總說顧問案件不印案件名稱
'   adocaseprogress.Open "select pa05, pa06, pa07, pa48, cp01 from caseprogress, patent where cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and cp09 = '" & InputNo & "' " & _
'                        "union select tm05, tm06, tm07, tm35, cp01 from caseprogress, trademark where cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and cp09 = '" & InputNo & "' " & _
'                        "union select lc05, lc06, lc07, lc17, cp01 from caseprogress, lawcase where cp01 = lc01 and cp02 = lc02 and cp03 = lc03 and cp04 = lc04 and cp09 = '" & InputNo & "' " & _
'                        "union select '顧問', '', '', '', cp01 from caseprogress, hirecase where cp01 = hc01 and cp02 = hc02 and cp03 = hc03 and cp04 = hc04 and cp09 = '" & InputNo & "' " & _
'                        "union select sp05, sp06, sp07, sp29, cp01 from caseprogress, servicepractice where cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and cp09 = '" & InputNo & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adocaseprogress.RecordCount <> 0 Then
'      Select Case InputSelect
'         Case 1
'            If IsNull(adocaseprogress.Fields(0).Value) Then
'               CaseNameQuery = MsgText(601)
'            Else
'               CaseNameQuery = adocaseprogress.Fields(0).Value
'               '92.3.21 add by sonia
'               If adocaseprogress.Fields(4).Value = "T" Or adocaseprogress.Fields(4).Value = "TF" Or adocaseprogress.Fields(4).Value = "FCT" Or adocaseprogress.Fields(4).Value = "CFT" Then
'                  If adocaseprogress.Fields(1).Value <> "" Then
'                     CaseNameQuery = CaseNameQuery & adocaseprogress.Fields(1).Value
'                  End If
'               End If
'               '92.3.21 end
'            End If
'         Case 2
'            If IsNull(adocaseprogress.Fields(1).Value) Then
'               CaseNameQuery = MsgText(601)
'            Else
'               CaseNameQuery = adocaseprogress.Fields(1).Value
'            End If
'         Case 3
'            If IsNull(adocaseprogress.Fields(2).Value) Then
'               CaseNameQuery = MsgText(601)
'            Else
'               CaseNameQuery = adocaseprogress.Fields(2).Value
'            End If
'      End Select
'      'Add By Cheng 2002/01/17
'      strCustCaseNo = "" & adocaseprogress.Fields(3).Value
'   Else
'      CaseNameQuery = MsgText(601)
'      'Add By Cheng 2002/01/17
'      strCustCaseNo = ""
'   End If
'   adocaseprogress.Close
'End Function
'
''Add By Cheng 2003/06/02
''設定列印字數
'Private Function ReOrgPrintTemp(strPrintTemp As String, intCnt As Integer) As String
'Dim ii As Integer
'
'ReOrgPrintTemp = ""
'For ii = 1 To Len(strPrintTemp)
'    ReOrgPrintTemp = ReOrgPrintTemp & Mid(strPrintTemp, ii, 1)
'    If Printer.TextWidth(ReOrgPrintTemp) > Printer.TextWidth(String(intCnt, "　")) Then
'        ReOrgPrintTemp = Left(ReOrgPrintTemp, Len(ReOrgPrintTemp) - 1)
'        Exit For
'    End If
'Next ii
'End Function

'一般列印
Private Sub PrintDoc()

Dim strName As String
Dim intList As Integer
Dim strProduct As String
Dim douService As Double
Dim douFee As Double
'Add By Cheng 2002/01/17
'Modify by Morgan 2004/8/12   '改成全域變數
'Dim strCustCaseNo As String '客戶案件案號
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim blnCombPrint As Boolean '收據項目合併列印
Dim intRecPos As Integer '目前指向第幾筆資料
'Add By Sindy 2009/07/14
Dim strPA08 As String, strPA09 As String
Dim strFeeType As String, strYF15 As String
Dim strKey(5) As String
Dim strCaseFee(1 To 2) As String
Dim bFind As Boolean
Dim varRef As Variant
'2009/07/14 End
Dim strSqlMain As String 'Add By Sindy 2010/5/13
Dim intRow As String 'Add By Sindy 2020/7/17
   
On Error GoTo Checking:

   'Added by Morgan 2013/8/14--瑞婷
   If Text1 = "" Then
      MsgBox "請輸入公司別", vbExclamation
      Text1.SetFocus
      Exit Sub
   End If
   'end 2013/8/14
   
   If FormCheck = False Then
      'Modified by Morgan 2013/4/30
      'MsgBox MsgText(181), , MsgText(5)
      MsgBox "請輸入列印條件！", vbExclamation
      Exit Sub
   End If
   
   strSql = ""
   
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If Me.Text2(0).Text <> MsgText(601) And Me.Text2(0).Text <> MsgText(29) Then
      strSql = strSql & " and a0k01 >= '" & Me.Text2(0).Text & "' "
   End If
   If Me.Text2(1).Text <> MsgText(601) And Me.Text2(1).Text <> MsgText(29) Then
      strSql = strSql & " and a0k01 <= '" & Me.Text2(1).Text & "' "
   End If
   
   'Add by Morgan 2007/2/6
   If strSql = "" Then
      MsgBox "【收據日期】及【收據號碼】條件不可同時空白！", vbExclamation
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   'end 2007/2/6
   
   strSql = strSql & " and a0k02>=920201" 'Add By Sindy 2024/8/20 以防人為誤下到舊系統資料
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and a0k11 = '" & Text1 & "'"
   End If
   
   bolRunWord = True 'Added by Lydia 2023/11/13
   
   lngAmount1 = 0
   lngAmount2 = 0
   intLength = 0
   
   'Add By Sindy 2010/5/13
   '2010/5/17 MODIFY BY SONIA A0K32='Y'者還是要考慮公司別條件
   'Modify By Sindy 2017/6/29 + and (a0k37 is null or a0k37<>'N') 剔除全銷
   'Modify By Sindy 2021/5/28
'Modify By Sindy 2024/5/16:
'在一般列印時會同時將已達可列印條件收據會一併列印出來 (不管收據日期), 請修改:
'1. 選擇收據日期條件時才將收據號碼範圍列印時，才將已達可列印條件收據會一併列印出來(不管收據日期)；
'2. 選擇收據號碼條件時，只列印符合收據號碼且可列印的收據印出，不在此號碼區間之可列印收據也不列印。
   If Text1.Text = "L" Then
      strSqlMain = "select * from ( " & _
                           "select * from acc0k0, customer, staff " & _
                           "where substr(a0k03, 1, 8) = cu01 (+) and substr(a0k03, 9, 1) = cu02 (+) " & _
                           "and to_number(substr(a0k01, 5, 5)) > 2000 " & _
                           "and a0k19 = 0 and (a0k09 is null or a0k09 = 0) and (a0k37 is null or a0k37<>'N') " & _
                           "and a0k32 IS NULL and a0k11<>'J' and a0k20=st01(+) " & strSql
      'Modify By Sindy 2024/5/16
      If Me.Text2(0).Text = MsgText(601) And Me.Text2(1).Text = MsgText(601) Then
      'Sindy 2024/5/16 END
         strSqlMain = strSqlMain & _
                           " Union All " & _
                           "select * from acc0k0, customer, staff " & _
                           "where substr(a0k03, 1, 8) = cu01 (+) and substr(a0k03, 9, 1) = cu02 (+) " & _
                           "and to_number(substr(a0k01, 5, 5)) > 2000 " & _
                           "and a0k19 = 0 and (a0k09 is null or a0k09 = 0) and (a0k37 is null or a0k37<>'N') " & _
                           "and (a0k32 ='Y') and a0k11 = '" & Text1 & "' and a0k20=st01(+) "
      End If
         strSqlMain = strSqlMain & _
                   ") where exists (" & Replace(strLOSSalesDuty, "\#ST06SQL#\", m_sqlST06) & ")"
      strSqlMain = strSqlMain & " union " & _
                   "select * from ( " & _
                           "select * from acc0k0, customer, staff " & _
                           "where substr(a0k03, 1, 8) = cu01 (+) and substr(a0k03, 9, 1) = cu02 (+) " & _
                           "and to_number(substr(a0k01, 5, 5)) > 2000 " & _
                           "and a0k19 = 0 and (a0k09 is null or a0k09 = 0) and (a0k37 is null or a0k37<>'N') " & _
                           "and a0k32 IS NULL and a0k11<>'J' and a0k20=st01(+) " & strSql & m_sqlST06
      'Modify By Sindy 2024/5/16
      If Me.Text2(0).Text = MsgText(601) And Me.Text2(1).Text = MsgText(601) Then
      'Sindy 2024/5/16 END
         strSqlMain = strSqlMain & _
                           " Union All " & _
                           "select * from acc0k0, customer, staff " & _
                           "where substr(a0k03, 1, 8) = cu01 (+) and substr(a0k03, 9, 1) = cu02 (+) " & _
                           "and to_number(substr(a0k01, 5, 5)) > 2000 " & _
                           "and a0k19 = 0 and (a0k09 is null or a0k09 = 0) and (a0k37 is null or a0k37<>'N') " & _
                           "and (a0k32 ='Y') and a0k11 = '" & Text1 & "' and a0k20=st01(+) " & m_sqlST06
      End If
         strSqlMain = strSqlMain & _
                   ") where not exists (" & Replace(strLOSSalesDuty, "\#ST06SQL#\", "") & ")"
      strSqlMain = "select * from ( " & strSqlMain & " ) order by a0k01 asc"
   Else
   '2021/5/28 END
      strSqlMain = "select * from ( " & _
                           "select * from acc0k0, customer, staff " & _
                           "where substr(a0k03, 1, 8) = cu01 (+) and substr(a0k03, 9, 1) = cu02 (+) " & _
                           "and to_number(substr(a0k01, 5, 5)) > 2000 " & _
                           "and a0k19 = 0 and (a0k09 is null or a0k09 = 0) and (a0k37 is null or a0k37<>'N') " & _
                           "and a0k32 IS NULL and a0k11<>'J' and a0k20=st01(+) " & strSql & m_sqlST06
      'Modify By Sindy 2024/5/16
      If Me.Text2(0).Text = MsgText(601) And Me.Text2(1).Text = MsgText(601) Then
      'Sindy 2024/5/16 END
         strSqlMain = strSqlMain & _
                           " Union All " & _
                           "select * from acc0k0, customer, staff " & _
                           "where substr(a0k03, 1, 8) = cu01 (+) and substr(a0k03, 9, 1) = cu02 (+) " & _
                           "and to_number(substr(a0k01, 5, 5)) > 2000 " & _
                           "and a0k19 = 0 and (a0k09 is null or a0k09 = 0) and (a0k37 is null or a0k37<>'N') " & _
                           "and (a0k32 ='Y') and a0k11 = '" & Text1 & "' and a0k20=st01(+) " & m_sqlST06
      End If
         strSqlMain = strSqlMain & _
                           ") order by a0k01 asc "
   End If
         
   If adoacc0k0.State = adStateOpen Then
      adoacc0k0.Close
   End If
   adoacc0k0.CursorLocation = adUseClient
   '收據餘額檔(一張收據一筆資料)
   'Modify By Sindy 2010/4/19
   'adoacc0k0.Open "select * from acc0k0, customer where substr(a0k03, 1, 8) = cu01 (+) and substr(a0k03, 9, 1) = cu02 (+) and to_number(substr(a0k01, 5, 5)) > 2000 and a0k19 = 0 and (a0k09 is null or a0k09 = 0)" & strSQL & " order by a0k01 asc", adoTaie, adOpenStatic, adLockReadOnly
   'adoacc0k0.Open "select * from acc0k0, customer where substr(a0k03, 1, 8) = cu01 (+) and substr(a0k03, 9, 1) = cu02 (+) and to_number(substr(a0k01, 5, 5)) > 2000 and a0k19 = 0 and (a0k09 is null or a0k09 = 0) and (a0k32 IS NULL OR A0K32<>'N') " & strSql & " order by a0k01 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoacc0k0.Open strSqlMain, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0k0.RecordCount = 0 Then
      adoacc0k0.Close
      Screen.MousePointer = vbDefault
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   If Check3.Value = 1 Then txtNote.Visible = True 'Add By Sindy 2022/1/27
   
   '初始化記錄收據號碼的變數
   strNo = ""
   adoacc0k0.MoveFirst
   'Add By Sindy 2020/3/23 空白A4直接列印收據
   If Text1.Text = "L" Then
      '切換印表機
      PUB_SetOsDefaultPrinter Combo1
      PUB_RestorePrinter Combo1
      Do While adoacc0k0.EOF = False
         '呼叫共用(與補印收據統一)
         If strNo <> adoacc0k0.Fields("a0k01").Value Then
            PUB_PrintCaseReceipt_L m_AttachPath & m_FileName, adoacc0k0, Me.chk(0), Me.chk(1)
            strNo = adoacc0k0.Fields("a0k01").Value
         End If
         adoacc0k0.MoveNext
      Loop
      adoacc0k0.Close
      '還原印表機
      PUB_SetOsDefaultPrinter strPrinter
      PUB_RestorePrinter strPrinter
   Else
   '2020/3/23 END
      intRow = 0
      
      'Add By Sindy 2022/1/26 直接用Word範本列印收據
      If Check3.Value = 1 Then
         Dim strStarTime As String
         strStarTime = Format(ServerTime, "##:##:##") & vbCrLf & vbCrLf & "收據日期= " & MaskEdBox1.Text & " ~ " & MaskEdBox2.Text & vbCrLf & "收據號碼= " & Text2(0).Text & " ~ " & Text2(1).Text & vbCrLf
         '切換印表機
         PUB_SetOsDefaultPrinter Combo2
         PUB_RestorePrinter Combo2
         
         Do While adoacc0k0.EOF = False
            'Modify by Morgan 2011/9/20 改呼叫共用(與補印收據統一)
            If strNo <> adoacc0k0.Fields("a0k01").Value Then
               intRow = intRow + 1
               
               'Modified by Morgan 2021/8/17 +單張列印選項
               'If intRow Mod 2 = 0 Then '雙數
               If (intRow Mod 2 = 0) And Check2.Value = 0 Then '雙數
               'end 2021/8/17
                  strWordFileName = ""
                  strPType = "雙張收據列印"
               Else '單數
                  If intRow < adoacc0k0.RecordCount And Check2.Value = 0 Then
                     strWordFileName = m_AttachPath & m_FileName3 '雙張收據範本檔
                     strPType = "雙張收據不印"
                  Else
                     strWordFileName = m_AttachPath & m_FileName2 '單張收據範本檔
                     strPType = ""
                  End If
               End If
               PUB_PrintCaseReceipt_Doc strWordFileName, adoacc0k0, Me.chk(0), Me.chk(1), , , , , , , True, , strPType
               
               strNo = adoacc0k0.Fields("a0k01").Value
            End If
            adoacc0k0.MoveNext
         Loop
         
         '還原印表機
         PUB_SetOsDefaultPrinter strPrinter2
         PUB_RestorePrinter strPrinter2
      Else
      '2022/1/26 END
      
'         '舊收據(點陣)
'         If Check1.Value = 1 Then
'      '      'Modify by Morgan 2008/3/25 控制 9x 才自訂
'      '      If pub_OS = "1" Then
'      '         Printer.Height = 8750
'      '         Printer.Width = 13000
'      '      Else
'               Printer.EndDoc
'               Printer.PaperSize = PUB_GetPaperSize(1)
'               Forms(0).StatusBar1.Panels(2).Text = Printer.PaperSize
'      '      End If
'            'end 2008/3/25
'            Printer.FontSize = 12
'            Printer.Font = "標楷體"
'            'Add By Cheng 2003/02/13
'            Do While adoacc0k0.EOF = False
'               'Modify by Morgan 2011/9/20 改呼叫共用(與補印收據統一)
'               If strNo <> adoacc0k0.Fields("a0k01").Value Then
'                  If strNo <> "" Then Printer.NewPage
'                  PUB_PrintCaseReceipt adoacc0k0, Me.chk(0), Me.chk(1)
'                  strNo = adoacc0k0.Fields("a0k01").Value
'               End If
'               adoacc0k0.MoveNext
'            Loop
'            'PrintSum 'Remove by Morgan 2011/9/19 改呼叫共用(與補印收據統一)
'            Printer.EndDoc
'
'         Else
'            'Add By Sindy 2020/7/15 列印A4收據
'            PUB_RestorePrinter Combo2 'Add By Sindy 2020/7/14 切換印表機
'
'            Printer.EndDoc
'            Printer.PaperSize = 9
'            Printer.Orientation = 1
'            Printer.Font = "新細明體"
'            Do While adoacc0k0.EOF = False
'               'Modify by Morgan 2011/9/20 改呼叫共用(與補印收據統一)
'               If strNo <> adoacc0k0.Fields("a0k01").Value Then
'                  intRow = intRow + 1
'                  'Modified by Morgan 2021/8/17 +單張列印選項
'                  'If intRow Mod 2 = 0 Then '雙數
'                  If (intRow Mod 2 = 0) And Check2.Value = 0 Then '雙數
'                  'end 2021/8/17
'                     Call PUB_PrintCaseReceiptTableMain(TextX, TextY, IIf(Option2(0).Value = True, 0, 1), 2)
'                  Else '單數
'                     If strNo <> "" Then Printer.NewPage
'                     Call PUB_PrintCaseReceiptTableMain(TextX, TextY, IIf(Option2(0).Value = True, 0, 1), 1)
'                  End If
'                  PUB_PrintCaseReceipt adoacc0k0, Me.chk(0), Me.chk(1), , , , , , , True
'                  strNo = adoacc0k0.Fields("a0k01").Value
'               End If
'               adoacc0k0.MoveNext
'            Loop
'            Printer.EndDoc
'            PUB_RestorePrinter strPrinter2 'Add By Sindy 2020/7/14 還原印表機
'            '2020/7/15 END
'         End If
'         Printer.Font = "新細明體"
      End If
      
      adoacc0k0.Close
   End If
   
   Screen.MousePointer = vbDefault
   txtNote.Visible = False 'Add By Sindy 2022/1/27
     
   'Modified by Lydia 2023/11/20
   'FormClear
   bolRunReceipt = True
   
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
'   ChkUnPrintData 'Add by Morgan 2011/9/20
   Exit Sub
   
Checking:
   Screen.MousePointer = vbDefault
   txtNote.Visible = False 'Add By Sindy 2022/1/27
   
   MsgBox Err.Description
   Printer.EndDoc 'Add By Sindy 2015/11/17
   
'Remove by Morgan 2011/9/20
'
'      If strNo <> adoacc0k0.Fields("a0k01").Value Then
'         If lngAmount1 <> 0 Or lngAmount2 <> 0 Then
'            PrintSum
'            lngAmount1 = 0
'            lngAmount2 = 0
'            Printer.NewPage
'         End If
'         strName = ""
'         intCounter = 0
'         intList = 0
'         If IsNull(adoacc0k0.Fields("a0k19").Value) Then
'            intTimes = 1
'         Else
'            intTimes = Val(adoacc0k0.Fields("a0k19").Value) + 1
'         End If
'         'Modify By Sindy 2010/4/23
'         adoTaie.Execute "update acc0k0 set a0k19 = " & intTimes & " where a0k01 = '" & adoacc0k0.Fields("a0k01").Value & "'"
'         'adoTaie.Execute "update acc0k0 set a0k19 = " & intTimes & ",a0k32=null where a0k01 = '" & adoacc0k0.Fields("a0k01").Value & "'"
'         '2010/4/23 End
'         PrintHead
'         strNo = adoacc0k0.Fields("a0k01").Value
'      End If
'
'
'        'Add By Cheng 2003/08/14
'        blnCombPrint = False
'        'CFP 美國 領證(601)及公開費(217)
'        ''Modify by Morgan 2007/5/8
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
'      '收據明細資料
'      '2005/6/16 MODIFY BY SONIA 加收文號排序
'      'adocaseprogress.Open "select * from caseprogress, casepropertymap, acc0j0 where cp01 = cpm01 and cp10 = cpm02 and cp09 = a0j01 (+) and cp60 = '" & adoacc0k0.Fields("a0k01").Value & "' and (cp79 <> 0 or (cp79 = 0 and cp75 <> 0))", adoTaie, adOpenStatic, adLockReadOnly
'      'Modify by Morgan 2007/5/8 案件性質名稱直接抓0j0的以免資料不一致卻沒發現
'      'adocaseprogress.Open "select * from caseprogress, casepropertymap, acc0j0 where cp01 = cpm01 and cp10 = cpm02 and cp09 = a0j01 (+) and cp60 = '" & adoacc0k0.Fields("a0k01").Value & "' and (cp79 <> 0 or (cp79 = 0 and cp75 <> 0)) ORDER BY CP09", adoTaie, adOpenStatic, adLockReadOnly
'      adocaseprogress.Open "select * from caseprogress, acc0j0 where cp60 = '" & adoacc0k0.Fields("a0k01").Value & "' and (cp79 <> 0 or (cp79 = 0 and cp75 <> 0)) and a0j01(+)=cp09 ORDER BY CP09 ", adoTaie, adOpenStatic, adLockReadOnly
'      'end 2007/5/8
'
'      '2005/6/16 END
'      intRecPos = 0
'      Do While adocaseprogress.EOF = False
'            intRecPos = intRecPos + 1
'
'         'Modify by Morgan 2004/10/14 程式碼上移，改先抓收款資料以便判斷是否已銷退
'         If adoquery.State = adStateOpen Then
'            adoquery.Close
'         End If
'         adoquery.CursorLocation = adUseClient
'         adoquery.Open "select sum(a1u07), sum(a1u09) from acc1u0 where a1u03 = '" & adocaseprogress.Fields("cp09").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'         If adoquery.RecordCount <> 0 Then
'            'Add by Morgan 2004/10/14 銷帳後費用為0的不印
'            If Val("" & adoquery.Fields(0).Value) + Val("" & adoquery.Fields(1).Value) = Val("" & adocaseprogress.Fields("cp16").Value) Then
'               adoquery.Close
'               GoTo NextRecord
'            End If
'
'            If IsNull(adoquery.Fields(0).Value) Then
'               douService = 0
'            Else
'               douService = Val(adoquery.Fields(0).Value)
'            End If
'            If IsNull(adoquery.Fields(1).Value) Then
'               douFee = 0
'            Else
'               douFee = Val(adoquery.Fields(1).Value)
'            End If
'         Else
'            douService = 0
'            douFee = 0
'         End If
'         adoquery.Close
'         '2004/10/14 end 程式碼上移
'
'         '列印收文日
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
'      '列印案件性質
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
'                  Printer.FontSize = 11
'                  Printer.Font = "標楷體"
'                  Printer.CurrentX = 1550
'                  Printer.CurrentY = 3600 + intCounter * 300
'                  If Val(adocaseprogress.Fields("cp53").Value) = Val(adocaseprogress.Fields("cp54").Value) Then
'                     Printer.Print adocaseprogress.Fields("a0j20").Value & " " & strYF15
'                  Else
'                     If adocaseprogress.Fields("cp10").Value = "605" Or _
'                        adocaseprogress.Fields("cp10").Value = "601" Then
'                        Printer.Print adocaseprogress.Fields("a0j20").Value & " (" & strYF15 & "至" & PUB_GetYF15(strPA09, strPA08, "Y000000" & m_FixNo, strFeeType, CDbl(Val(adocaseprogress.Fields("cp54").Value))) & ")"
'                     Else
'                        Printer.Print adocaseprogress.Fields("a0j20").Value & " (" & strYF15 & "至" & PUB_GetYF15(strPA09, strPA08, "Y000000" & m_FixNo, strFeeType, Val(varRef(Val(adocaseprogress.Fields("cp54").Value) - 1))) & ")"
'                     End If
'                  End If
'               End If
'            End If
'         '2009/07/14 End
'         Else
'            Printer.CurrentX = 1550
'            Printer.CurrentY = 3600 + intCounter * 300
'            Printer.Print adocaseprogress.Fields("a0j20").Value
'         End If
'      End If
'      Printer.FontSize = 12
'      Printer.Font = "標楷體"
'      'end 2007/5/8
'
'        '2004/10/14 程式碼上移前位置
'         If IsNull(adocaseprogress.Fields("cp16").Value) = False Then
'            If IsNull(adocaseprogress.Fields("cp17").Value) = False And IsNull(adoacc0k0.Fields("a0k30").Value) Then
'               If adocaseprogress.Fields("cp16").Value - adocaseprogress.Fields("cp17").Value = 0 Then
'                  strAmount1 = "0"
'               Else
'                  strAmount1 = Format(Val(adocaseprogress.Fields("cp16").Value) - Val(adocaseprogress.Fields("cp17").Value) - douService, DDollar)
'               End If
'                If blnCombPrint = False Then
'                    intLength = Printer.TextWidth(strAmount1)
'                    Printer.CurrentX = 5900 - intLength
'                    Printer.CurrentY = 3600 + intCounter * 300
'                    '列印費用
'                    Printer.Print strAmount1
'                End If
'               lngAmount1 = lngAmount1 + Val(adocaseprogress.Fields("cp16").Value) - Val(adocaseprogress.Fields("cp17").Value) - douService
'               If Val(adocaseprogress.Fields("cp17").Value) = 0 Then
'                  strAmount2 = "0"
'               Else
'                  strAmount2 = Format(Val(adocaseprogress.Fields("cp17").Value) - douFee, DDollar)
'               End If
'                If blnCombPrint = False Then
'                    intLength = Printer.TextWidth(strAmount2)
'                    Printer.CurrentX = 7600 - intLength
'                    Printer.CurrentY = 3600 + intCounter * 300
'                    '列印規費
'                    Printer.Print strAmount2
'                End If
'               lngAmount2 = lngAmount2 + Val(adocaseprogress.Fields("cp17").Value) - douFee
'            Else
'               strAmount1 = Format(Val(adocaseprogress.Fields("cp16").Value) - douService - douFee, DDollar)
'                'edit by nickc 2007/02/08
'                'If blnCombPrint = fasle Then
'                If blnCombPrint = False Then
'                    intLength = Printer.TextWidth(strAmount1)
'                    Printer.CurrentX = 5900 - intLength
'                    Printer.CurrentY = 3600 + intCounter * 300
'                    '列印費用
'                    If strAmount1 = "" Then
'                       Printer.Print "0"
'                    Else
'                       Printer.Print strAmount1
'                    End If
'                End If
'               lngAmount1 = lngAmount1 + Val(adocaseprogress.Fields("cp16").Value) - douService - douFee
'               strAmount2 = "0"
'                If blnCombPrint = False Then
'                    intLength = Printer.TextWidth(strAmount2)
'                    Printer.CurrentX = 7600 - intLength
'                    Printer.CurrentY = 3600 + intCounter * 300
'                    Printer.Print strAmount2
'                End If
'               lngAmount2 = lngAmount2 + 0
'            End If
'         End If
'            'Add By Cheng 2003/08/14
'            If blnCombPrint = True And intRecPos = 2 Then
'                '列印費用
'                intCounter = intCounter - 1
'                If lngAmount1 = 0 Then
'                    strAmount1 = "0"
'                Else
'                    strAmount1 = Format(lngAmount1, DDollar)
'                End If
'                intLength = Printer.TextWidth(strAmount1)
'                Printer.CurrentX = 5900 - intLength
'                Printer.CurrentY = 3600 + intCounter * 300
'                Printer.Print strAmount1
'                '列印規費
'                If lngAmount2 = 0 Then
'                    strAmount2 = "0"
'                Else
'                    strAmount2 = Format(lngAmount2, DDollar)
'                End If
'                intLength = Printer.TextWidth(strAmount2)
'                Printer.CurrentX = 7600 - intLength
'                Printer.CurrentY = 3600 + intCounter * 300
'                Printer.Print strAmount2
'            End If
''         If strName <> (adocaseprogress.Fields("cp01").Value & adocaseprogress.Fields("cp02").Value & adocaseprogress.Fields("cp03").Value & adocaseprogress.Fields("cp04").Value) Then
'         If intList < 2 Then
'            'Modify By Cheng 2003/08/14
'            If blnCombPrint = False Or (blnCombPrint = True And intRecPos = 1) Then
'                Printer.CurrentX = 1500
'                Printer.CurrentY = 7200 + intList * 550
'                '列印本所案號
'                If (adocaseprogress.Fields("cp03").Value & adocaseprogress.Fields("cp04").Value) = "000" Then
'                   Printer.Print adocaseprogress.Fields("cp01").Value & adocaseprogress.Fields("cp02").Value
'                Else
'                   Printer.Print adocaseprogress.Fields("cp01").Value & adocaseprogress.Fields("cp02").Value & adocaseprogress.Fields("cp03").Value & adocaseprogress.Fields("cp04").Value
'                End If
'                Printer.CurrentX = 3000
'                Printer.CurrentY = 7200 + intList * 550
'                '列印申請國家名稱
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
'                      strProduct = "第" & adoquery.Fields("tm09").Value & " 類"
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
'                        strProduct = "第" & adoquery.Fields("SP18").Value & "類"
'                     End If
'                  Else
'                     strProduct = MsgText(601)
'                  End If
'                  adoquery.Close
'                End If
'                '2005/7/1 END
'                Printer.CurrentX = 5500
'                Printer.CurrentY = 7200 + intList * 550
'                '列印案件名稱(以總收文號查詢)及客戶案件名稱(Optional)
'                'Modify by Morgan 2004/8/12
''                If CaseNameQuery(adocaseprogress.Fields("cp09").Value, 1, strCustCaseNo) <> "" Then
''                   'Modify By Cheng 2002/01/17
''    '               Printer.Print CaseNameQuery(adocaseprogress.Fields("cp09").Value, 1)
''                    'Modify By Cheng 2003/06/02
''    '               Printer.Print Trim(CaseNameQuery(adocaseprogress.Fields("cp09").Value, 1, strCustCaseNo)) & _
''    '                              IIf(Me.chk(1).Value = vbChecked, " " & strCustCaseNo, "") & " " & strProduct
''                    Printer.Print ReOrgPrintTemp(Trim(CaseNameQuery(adocaseprogress.Fields("cp09").Value, 1, strCustCaseNo)) & _
''                                  IIf(Me.chk(1).Value = vbChecked, " " & strCustCaseNo, "") & " " & strProduct, 22)
'
'                   'Modify By Cheng 2002/01/17
'    '               Printer.Print CaseNameQuery(adocaseprogress.Fields("cp09").Value, 2)
'                    'Modify By Cheng 2003/06/02
'    '               Printer.Print Trim(CaseNameQuery(adocaseprogress.Fields("cp09").Value, 2, strCustCaseNo)) & _
'    '                              IIf(Me.chk(1).Value = vbChecked, " " & strCustCaseNo, "") & " " & strProduct
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
'         'Add by Morgan 2005/9/15 顧問案的顧問聘任(0)時要印顧問期間
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

'Added by Morgan 2013/4/30
'重新列印
Private Sub PrintDocX()
   Dim stCon As String
   Dim intRow As String 'Add By Sindy 2020/7/17
   Dim strVal As String
   
   If FormCheck1() = False Then
      Exit Sub
   End If
   
   bolRunWord = True 'Added by Lydia 2023/11/13
   
On Error GoTo ErrHnd
   
   stCon = " and a0k02>=920201" 'Add By Sindy 2024/8/20 以防人為誤下到舊系統資料
   stCon = stCon & " and a0k14>=to_date('" & DBDATE(MaskEdBox3.Text) & Replace(MaskEdBox4.Text, ":", "") & "','YYYYMMDDHH24MISS')"
   stCon = stCon & " and a0k14<=to_date('" & DBDATE(MaskEdBox3.Text) & Replace(MaskEdBox5.Text, ":", "") & "','YYYYMMDDHH24MISS')"
   
   'Modify By Sindy 2015/11/17 + customer
   'strSql = "select * from acc0k0 where a0k19=1 and a0k11='" & Text1 & "'" & stCon
   'Modify By Sindy 2022/3/7 + and (a0k09 is null or a0k09 = 0)
   strVal = "select * from acc0k0,customer,staff where a0k19=1 and a0k11='" & Text1 & "'" & _
            " and substr(a0k03,1,8)=cu01(+) and substr(a0k03,9,1)=cu02(+) and a0k20=st01(+) " & stCon & " and (a0k09 is null or a0k09 = 0)"
   '2015/11/17 END
   'Modify By Sindy 2021/5/28
   If Text1.Text = "L" Then
      strSql = strVal & " and exists (" & Replace(strLOSSalesDuty, "\#ST06SQL#\", m_sqlST06) & ")" & _
               " union " & _
               strVal & m_sqlST06 & " and not exists (" & Replace(strLOSSalesDuty, "\#ST06SQL#\", "") & ")"
   Else
   '2021/5/28 END
      strSql = strVal & m_sqlST06
      strSql = strSql & " order by a0k01" 'Add By Sindy 2025/9/17
   End If
   intI = 1
   Set adoacc0k0 = ClsLawReadRstMsg(intI, strSql)
   If intI = 0 Then
      adoacc0k0.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   If Check3.Value = 1 Then txtNote.Visible = True 'Add By Sindy 2022/1/27
   
   '初始化記錄收據號碼的變數
   strNo = ""
   adoacc0k0.MoveFirst
   'Add By Sindy 2020/3/23 空白A4直接列印收據
   If Text1.Text = "L" Then
      '切換印表機
      PUB_SetOsDefaultPrinter Combo1
      PUB_RestorePrinter Combo1
      Do While adoacc0k0.EOF = False
         '呼叫共用(與補印收據統一)
         If strNo <> adoacc0k0.Fields("a0k01").Value Then
            PUB_PrintCaseReceipt_L m_AttachPath & m_FileName, adoacc0k0, Me.chk(0), Me.chk(1), , , , , False
            strNo = adoacc0k0.Fields("a0k01").Value
         End If
         adoacc0k0.MoveNext
      Loop
      adoacc0k0.Close
      '還原印表機
      PUB_SetOsDefaultPrinter strPrinter
      PUB_RestorePrinter strPrinter
   Else
   '2020/3/23 END
      intRow = 0
      
      'Add By Sindy 2022/1/26 直接用Word範本列印收據
      If Check3.Value = 1 Then
         Dim strStarTime As String
         strStarTime = Format(ServerTime, "##:##:##") & vbCrLf & vbCrLf & "收據號碼= " & Text2(2) & " 以後(含)" & vbCrLf & stCon & vbCrLf
         '切換印表機
         PUB_SetOsDefaultPrinter Combo2
         PUB_RestorePrinter Combo2
         
         Do While adoacc0k0.EOF = False
            'Modify by Morgan 2011/9/20 改呼叫共用(與補印收據統一)
            If strNo <> adoacc0k0.Fields("a0k01").Value Then
               intRow = intRow + 1
               
               'Modified by Morgan 2021/8/17 +單張列印選項
               'If intRow Mod 2 = 0 Then '雙數
               If (intRow Mod 2 = 0) And Check2.Value = 0 Then '雙數
               'end 2021/8/17
                  strWordFileName = ""
                  strPType = "雙張收據列印"
               Else '單數
                  If intRow < adoacc0k0.RecordCount And Check2.Value = 0 Then
                     strWordFileName = m_AttachPath & m_FileName3 '雙張收據範本檔
                     strPType = "雙張收據不印"
                  Else
                     strWordFileName = m_AttachPath & m_FileName2 '單張收據範本檔
                     strPType = ""
                  End If
               End If
               PUB_PrintCaseReceipt_Doc strWordFileName, adoacc0k0, Me.chk(0), Me.chk(1), , , , , , False, True, , strPType
               
               strNo = adoacc0k0.Fields("a0k01").Value
            End If
            adoacc0k0.MoveNext
         Loop
         
         '還原印表機
         PUB_SetOsDefaultPrinter strPrinter2
         PUB_RestorePrinter strPrinter2
      Else
      '2022/1/26 END
      
'         '舊收據(點陣)
'         If Check1.Value = 1 Then
'      '      If pub_OS = "1" Then
'      '         Printer.Height = 8750
'      '         Printer.Width = 13000
'      '      Else
'               Printer.EndDoc
'               Printer.PaperSize = PUB_GetPaperSize(1)
'               Forms(0).StatusBar1.Panels(2).Text = Printer.PaperSize
'      '      End If
'            Printer.FontSize = 12
'            Printer.Font = "標楷體"
'            Do While adoacc0k0.EOF = False
'               If strNo <> adoacc0k0.Fields("a0k01").Value Then
'                  If strNo <> "" Then Printer.NewPage
'                  PUB_PrintCaseReceipt adoacc0k0, Me.chk(0), Me.chk(1), , , , , , False
'                  strNo = adoacc0k0.Fields("a0k01").Value
'               End If
'               adoacc0k0.MoveNext
'            Loop
'            Printer.EndDoc
'
'         Else
'            'Add By Sindy 2020/7/15 列印A4收據
'            PUB_RestorePrinter Combo2 'Add By Sindy 2020/7/14 切換印表機
'
'            Printer.EndDoc
'            Printer.PaperSize = 9
'            Printer.Orientation = 1
'            Printer.Font = "新細明體"
'            Do While adoacc0k0.EOF = False
'               'Modify by Morgan 2011/9/20 改呼叫共用(與補印收據統一)
'               If strNo <> adoacc0k0.Fields("a0k01").Value Then
'                  intRow = intRow + 1
'                  If intRow Mod 2 = 0 Then '雙數
'                     Call PUB_PrintCaseReceiptTableMain(TextX, TextY, IIf(Option2(0).Value = True, 0, 1), 2)
'                  Else '單數
'                     If strNo <> "" Then Printer.NewPage
'                     Call PUB_PrintCaseReceiptTableMain(TextX, TextY, IIf(Option2(0).Value = True, 0, 1), 1)
'                  End If
'                  PUB_PrintCaseReceipt adoacc0k0, Me.chk(0), Me.chk(1), , , , , , False, True
'                  strNo = adoacc0k0.Fields("a0k01").Value
'               End If
'               adoacc0k0.MoveNext
'            Loop
'            Printer.EndDoc
'            PUB_RestorePrinter strPrinter2 'Add By Sindy 2020/7/14 還原印表機
'            '2020/7/15 END
'         End If
'         Printer.Font = "新細明體"
      End If
      
      adoacc0k0.Close
   End If
   
   Screen.MousePointer = vbDefault
   txtNote.Visible = False 'Add By Sindy 2022/1/27
   
   'Modified by Lydia 2023/11/20
   'FormClear
   bolRunReceipt = True
   
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
'   ChkUnPrintData
   Exit Sub
   
ErrHnd:
   Screen.MousePointer = vbDefault
   txtNote.Visible = False 'Add By Sindy 2022/1/27
   
   MsgBox Err.Description
   Printer.EndDoc 'Add By Sindy 2015/11/17
End Sub

''Add By Sindy 2020/7/8
''畫表格
'Private Sub PrintTable1(Px() As Integer, Py() As Integer)
'Dim yi As Long
'
'   Px(1) = Px(0) + 4 * TwPerCm
'   Px(2) = Px(1) + 4 * TwPerCm
'   Px(3) = Px(2) + 4 * TwPerCm
'   Px(4) = Px(3) + 4 * TwPerCm
'   Px(5) = Px(4) + 4 * TwPerCm + 150
'
'   For intI = 1 To 12
'      Py(intI) = Py(intI - 1) + 1.5 * iLineHeight
'   Next
'
'   Printer.DrawWidth = 7
'   '框
'   Printer.Line (Px(0), Py(0))-(Px(5), Py(10)), , B
'
'   Printer.DrawWidth = 1
'   '縱線
'   Printer.Line (Px(2) - 500, Py(0))-(Px(2) - 500, Py(8))
'   Printer.Line (Px(3) - 500, Py(0))-(Px(3) - 500, Py(7))
'   Printer.Line (Px(4) - 500, Py(0))-(Px(4) - 500, Py(8))
'
'   '橫線
'   Printer.Line (Px(0), Py(1))-(Px(5), Py(1)) '1
'   For intI = 6 To 7
'      Printer.Line (Px(0), Py(intI))-(Px(4) - 500, Py(intI))
'   Next
'   Printer.Line (Px(0), Py(8))-(Px(5), Py(8)) '8
'
'   Printer.FontSize = 12
'   Printer.CurrentX = Px(0) + 400
'   Printer.CurrentY = Py(0) + 100
'   HPrint "帳　　款　　類　　別"
'   Printer.CurrentX = Px(2) - 200
'   Printer.CurrentY = Py(0) + 100
'   HPrint "本所服務費"
'   Printer.CurrentX = Px(3) - 300
'   Printer.CurrentY = Py(0) + 100
'   HPrint "代收政府規費"
'   Printer.CurrentX = Px(4) - 100
'   Printer.CurrentY = Py(0) + 100
'   HPrint "收 據 專 用 章"
'
'   Printer.CurrentX = Px(0) + 400
'   Printer.CurrentY = Py(6) + 100
'   HPrint "小　　　　　　　　計"
'   Printer.CurrentX = Px(0) + 400
'   Printer.CurrentY = Py(7) + 100
'   HPrint "總　　　　　　　　計"
'   Printer.CurrentX = Px(3) + 1000
'   Printer.CurrentY = Py(7) + 100
'   Printer.Print "元整"
'
'   Printer.FontSize = 9
'   Printer.CurrentX = Px(3) - 150
'   Printer.CurrentY = Py(5) + 200
'   Printer.Print "(此項規費請勿扣繳)"
'   Printer.CurrentX = Px(0) + 200
'   yi = Py(8) + 100
'   Printer.CurrentY = yi
'   Printer.Print "附註：1.依所得稅法第88條規定，貴公司因報帳之需要，服務費金額在二萬元以上，請扣除執行業務所得稅10%。扣除之稅款"
'   Printer.CurrentX = Px(0) + 200
'   yi = yi + 1 * iLineHeight - 50
'   Printer.CurrentY = yi
'   Printer.Print "　　　   請務必於次月十日前向國庫繳清。請依(收據資料)填具扣繳憑單寄本所財務處。"
'   Printer.CurrentX = Px(0) + 200
'   yi = yi + 1 * iLineHeight - 50
'   Printer.CurrentY = yi
'   Printer.Print "　　　2.申請人為個人者，請勿扣繳。"
'
'   'Printer.Line (Px(0), Py(12))-(Px(12), Py(12))
'   Printer.FontSize = 12
'   Printer.CurrentX = Px(0)
'   yi = Py(10) + 50
'   Printer.CurrentY = yi
'   Printer.Print "本所案號：" & "FCT-052232" & "　　案件名稱：淨水器"
'   Printer.CurrentX = Px(0)
'   yi = yi + 1 * iLineHeight
'   Printer.CurrentY = yi
'   Printer.Print "　　　　　台灣"
'End Sub

'Added by Lydia 2023/11/13 國內接洽單：DEBIT NOTE請款選項->1.立即開立DEBIT NOTE清單
Private Sub PrintDebitList()
Dim strQuery As String, intQ As Integer, iStart As Integer
Dim rsQuery As New ADODB.Recordset
Dim strFilePath As String, intCounter As Integer, intB As Integer
Dim stTmp, intWidth
Dim xlsAgentPoint As New Excel.Application
Dim wksrpt As New Worksheet
Dim strTitle As String
Dim tmpArray As Variant, MaxCols As Integer


On Error GoTo ErrHandle

   strSql = ""
   
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If InStr(UCase(strSql), "A0K02") > 0 Then
      strTitle = strTitle & "收據日期：" & Val(FCDate(MaskEdBox1.Text)) & "~" & Val(FCDate(MaskEdBox2.Text)) & String(10, " ")
   End If
   
   If Me.Text2(0).Text <> MsgText(601) And Me.Text2(0).Text <> MsgText(29) Then
      strSql = strSql & " and a0k01 >= '" & Me.Text2(0).Text & "' "
   End If
   If Me.Text2(1).Text <> MsgText(601) And Me.Text2(1).Text <> MsgText(29) Then
      strSql = strSql & " and a0k01 <= '" & Me.Text2(1).Text & "' "
   End If
   If InStr(UCase(strSql), "A0K01") > 0 Then
      strTitle = strTitle & "收據號碼：" & Me.Text2(0).Text & "~" & Me.Text2(1).Text & String(10, " ")
   End If
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and a0k11 = '" & Text1 & "'"
   End If
   If InStr(UCase(strSql), "A0K11") > 0 Then
      strTitle = strTitle & "公司別：" & Me.Text1.Text & LblCmpName & String(10, " ")
   End If
   strQuery = "select a0k01,a0k03,nvl(cu04,nvl(cu05,cu06)) as a0k03name,a0k06,a0k07,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) caseno,cp13,st02 " & _
              "from acc0k0,caseprogress,ConsultRecordList,customer,staff " & _
              "where a0k32='Z' " & strSql & " and a0k01=cp60(+) and cp140=crl01(+) and CRL153='1' " & _
              "and substr(a0k03,1,8)=cu01(+) and substr(a0k03,9,1)=cu02(+) and cp13=st01(+) " & _
              "group by a0k01,a0k03,nvl(cu04,nvl(cu05,cu06)),a0k06,a0k07,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04),cp13,st02 "
   strQuery = strQuery & " order by a0k03, a0k01"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strQuery)
   If intQ = 1 Then
      strFilePath = App.path & "\" & strUserNum
      Call Pub_ChkExcelPath(strFilePath)
      Call PUB_KillTempFile(strUserNum & "\立即開立DEBIT NOTE清單*.*")
      strFilePath = strFilePath & "\立即開立DEBIT NOTE清單_" & strSrvDate(1) & ".xls"
      stTmp = Array("客戶代號", "客戶名稱", "智權人員", "收據編號", "本所案號", "金額")
      intWidth = Array(12, 20, 15, 12, 12, 12)
      MaxCols = UBound(stTmp) + 1
      ReDim tmpArray(1 To MaxCols)
      intCounter = 1
      xlsAgentPoint.SheetsInNewWorkbook = 1 '預設工作表數目
      xlsAgentPoint.Workbooks.add
      xlsAgentPoint.Application.Visible = False
      Set wksrpt = xlsAgentPoint.Worksheets(1)
      xlsAgentPoint.Sheets(1).Select '選擇工作表
      xlsAgentPoint.ActiveWindow.FreezePanes = False
      xlsAgentPoint.ActiveWindow.SplitColumn = 0
      xlsAgentPoint.ActiveWindow.SplitRow = IIf(strTitle <> "", 4, 3)
      xlsAgentPoint.ActiveWindow.FreezePanes = True '凍結窗格(要在有資料前設定
      
      wksrpt.Range("C" & intCounter).Value = "立即開立DEBIT NOTE清單"
      wksrpt.Range("C" & intCounter).Font.Size = 16
      wksrpt.Range(intCounter & ":" & intCounter).RowHeight = 22
      If strTitle <> "" Then
         intCounter = intCounter + 1
         wksrpt.Range("A" & intCounter).Value = Trim(strTitle)
      End If
      intCounter = intCounter + 2
      
      '設定欄位名稱及欄寬
      For intB = LBound(stTmp) To UBound(stTmp)
          wksrpt.Range(Pub_NumberToSystem26(intB + 1) & intCounter).Value = stTmp(intB)
          wksrpt.Columns(Pub_NumberToSystem26(intB + 1) & ":" & Pub_NumberToSystem26(intB + 1)).ColumnWidth = intWidth(intB)
          wksrpt.Range(Pub_NumberToSystem26(intB + 1) & intCounter).HorizontalAlignment = xlCenter
      Next intB
      With wksrpt.Range("A" & intCounter & ":" & Pub_NumberToSystem26(MaxCols) & intCounter)
          .Borders(xlEdgeBottom).LineStyle = xlContinuous
          .Borders(xlEdgeBottom).Weight = xlThin
      End With
      
      intCounter = intCounter + 1
      iStart = intCounter
      intQ = rsQuery.RecordCount
      rsQuery.MoveFirst
      Do While Not rsQuery.EOF
         '因為逐筆輸入過慢,改成陣列輸入
         tmpArray(1) = "" & rsQuery.Fields("a0k03")
         tmpArray(2) = "" & rsQuery.Fields("a0k03name")
         tmpArray(3) = "" & rsQuery.Fields("cp13") & rsQuery.Fields("st02")
         tmpArray(4) = "" & rsQuery.Fields("a0k01")
         tmpArray(5) = "" & rsQuery.Fields("caseno")
         tmpArray(6) = Val("" & rsQuery.Fields("a0k06")) + Val("" & rsQuery.Fields("a0k07"))
         wksrpt.Range("A" & intCounter & ":" & Pub_NumberToSystem26(MaxCols) & intCounter).Value = tmpArray
         wksrpt.Range("F" & intCounter).NumberFormatLocal = "##,##0"
         intCounter = intCounter + 1
         rsQuery.MoveNext
      Loop
      With wksrpt.Range("A" & intCounter & ":" & Pub_NumberToSystem26(MaxCols) & intCounter)
          .Borders(xlEdgeTop).LineStyle = xlContinuous
          .Borders(xlEdgeTop).Weight = xlThin
      End With
      wksrpt.Range("C" & intCounter).Value = "共" & intQ & "筆"
      wksrpt.Range(Pub_NumberToSystem26(MaxCols) & intCounter).Formula = "=SUM(" & Pub_NumberToSystem26(MaxCols) & iStart & ":" & Pub_NumberToSystem26(MaxCols) & intCounter - 1 & ")"
      wksrpt.Range(Pub_NumberToSystem26(MaxCols) & intCounter).NumberFormatLocal = "##,##0"

      wksrpt.Range("A1").Select
      
      If Val(xlsAgentPoint.Version) < 12 Then
         xlsAgentPoint.Workbooks(1).SaveAs FileName:=strFilePath, FileFormat:=-4143
      Else
         xlsAgentPoint.Workbooks(1).SaveAs FileName:=strFilePath, FileFormat:=56
      End If
      xlsAgentPoint.Workbooks.Close
      xlsAgentPoint.Quit
      Set xlsAgentPoint = Nothing
      Set wksrpt = Nothing
      '寄給操作者
      Sleep 100
      PUB_SendMail strUserNum, strUserNum, "", "立即開立DEBIT NOTE清單_" & strSrvDate(1), vbCrLf & "請參考附件。", , strFilePath
      
   End If
   Set rsQuery = Nothing
   
   Exit Sub
   
ErrHandle:
    If Err.Number <> 0 Then
        MsgBox "產生Excel失敗：" & vbCrLf & Err.Description, vbCritical
    End If
End Sub

