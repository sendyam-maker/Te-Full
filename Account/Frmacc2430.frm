VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc2430 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "國外結匯水單列印"
   ClientHeight    =   5484
   ClientLeft      =   7476
   ClientTop       =   6852
   ClientWidth     =   5832
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5484
   ScaleWidth      =   5832
   Begin VB.TextBox txtWordTB 
      Height          =   315
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   36
      Top             =   3900
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txtXY 
      Height          =   285
      Index           =   1
      Left            =   7470
      TabIndex        =   30
      Top             =   1560
      Width           =   705
   End
   Begin VB.TextBox txtXY 
      Height          =   285
      Index           =   0
      Left            =   7470
      TabIndex        =   27
      Top             =   1200
      Width           =   705
   End
   Begin VB.TextBox txtKind 
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
      Height          =   315
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   7
      Text            =   "Y"
      Top             =   2520
      Width           =   400
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1620
      Style           =   2  '單純下拉式
      TabIndex        =   9
      Top             =   3510
      Width           =   4050
   End
   Begin VB.TextBox Text6 
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
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1800
      Width           =   400
   End
   Begin VB.TextBox Text5 
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
      Left            =   6960
      MaxLength       =   1
      TabIndex        =   29
      Top             =   2280
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.OptionButton Option1 
      Caption         =   "只印合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   1
      Left            =   3240
      TabIndex        =   17
      Top             =   210
      Width           =   1360
   End
   Begin VB.OptionButton Option1 
      Caption         =   "一般列印"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   0
      Left            =   720
      TabIndex        =   16
      Top             =   210
      Value           =   -1  'True
      Width           =   1360
   End
   Begin VB.TextBox Text4 
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
      Left            =   1680
      TabIndex        =   2
      Top             =   990
      Width           =   1572
   End
   Begin VB.TextBox Text3 
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
      Left            =   3600
      TabIndex        =   3
      Top             =   990
      Width           =   1572
   End
   Begin VB.TextBox Text2 
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
      Height          =   315
      Left            =   3600
      TabIndex        =   5
      Top             =   1395
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
      Left            =   1680
      TabIndex        =   4
      Top             =   1395
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1680
      TabIndex        =   0
      Top             =   600
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command2 
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
      Left            =   570
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   3060
      Width           =   4725
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3600
      TabIndex        =   1
      Top             =   600
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label20 
      BackStyle       =   0  '透明
      Caption         =   "華銀水單：只列印第1面"
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
      Height          =   255
      Left            =   3210
      TabIndex        =   38
      Top             =   2730
      Width           =   2565
   End
   Begin VB.Label Label19 
      BackStyle       =   0  '透明
      Caption         =   "臺銀水單：雙面列印"
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
      Height          =   255
      Left            =   3210
      TabIndex        =   37
      Top             =   2400
      Width           =   2115
   End
   Begin VB.Label Label18 
      Caption         =   "台銀Word列印："
      Height          =   225
      Left            =   2400
      TabIndex        =   35
      Top             =   3930
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "結匯明細匯總表產生在C:\個人桌面\XLS"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   720
      TabIndex        =   34
      Top             =   5040
      Width           =   4845
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "白俄羅斯3.代客戶結匯者4.票匯5.手續費為71:BEN "
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   720
      TabIndex        =   33
      Top             =   4766
      Width           =   4815
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "必須紙本列印之條件:1.匯款銀行為中文字2.匯款地為"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   720
      TabIndex        =   32
      Top             =   4493
      Width           =   4695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "縱軸偏移值(Y)：　　　　　　(單位公分)"
      Height          =   180
      Index           =   3
      Left            =   6030
      TabIndex        =   31
      Top             =   1605
      Width           =   3240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "橫軸偏移值(X)：　　　　　　(單位公分)"
      Height          =   180
      Index           =   1
      Left            =   6030
      TabIndex        =   28
      Top             =   1245
      Width           =   3240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "注意：此為調整畫面上印表機的X及Y偏移值。"
      ForeColor       =   &H00000080&
      Height          =   180
      Index           =   2
      Left            =   6000
      TabIndex        =   26
      Top             =   1920
      Width           =   3660
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "必須紙本結匯　　　(Y)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   25
      Top             =   2550
      Width           =   3495
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "注意：在產生水單時，不要使用Word！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   4220
      Width           =   4455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "（白紙、雷射印表機）"
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
      Left            =   60
      TabIndex        =   23
      Top             =   3840
      Width           =   2355
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "（空白：不含J公司）"
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
      Left            =   7410
      TabIndex        =   22
      Top             =   2340
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "（空白：不含J公司）"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2010
      TabIndex        =   21
      Top             =   1845
      Width           =   2250
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "匯款單印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   3600
      Width           =   1515
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   600
      TabIndex        =   19
      Top             =   1845
      Width           =   720
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   240
      X2              =   240
      Y1              =   60
      Y2              =   2300
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   5550
      X2              =   5550
      Y1              =   60
      Y2              =   2280
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   240
      X2              =   5550
      Y1              =   2310
      Y2              =   2310
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   210
      X2              =   5550
      Y1              =   60
      Y2              =   60
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
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
      Height          =   210
      Left            =   6000
      TabIndex        =   18
      Top             =   2340
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "付款單號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   1020
      Width           =   975
   End
   Begin VB.Label Label4 
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
      Left            =   3360
      TabIndex        =   14
      Top             =   1020
      Width           =   255
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
      Index           =   0
      Left            =   3360
      TabIndex        =   13
      Top             =   600
      Width           =   255
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
      Left            =   3360
      TabIndex        =   12
      Top             =   1425
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "代理人"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   1425
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   30
      Top             =   780
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "付款日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   623
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc2430"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2025/07/31 台銀水單改版J20-01B(114.04)：受款銀行增加「城市、國別、中間銀行代號(Swift Code)」，受款人增加增加「城市、國別」
'Memo by Lydia 2023/04/13 台銀水單改成Word套印-版本J20-01B(2020.12)； 啟用日控制; 等印表機色帶印完再換版本--溫斯閔
      '請作Email:109/11/17 Emal: 台銀水單因點陣印表機取消改以影印機列印雙面如附件
      '109/11/26 婉莘說:台銀要求兩面都印, 所以出9F影印機不彩印
      '109/12/22 婉莘：確認是正反兩面都印並且一件二份，目前優先收複寫紙版
'end 2023/04/13
'Memo by Lydia 2018/11/12 台銀水單格式為Letter Extra(寬9.5英吋 x 12英吋),若指定印表機可設定的紙張無該項,必須先在列印伺服器內容新增紙張格式-水單(寬9.5英吋 x 12英吋),然後在Account\系統管理\報表紙張格式設定將指定印表機開始設定PUB_GetPaperSize
'Memo by Lydia 2018/11/09 台銀新版(107.08) ;LEDOMARS 7800II 與財務室IBM-5577印表機在列印受款行資料有所不同,財務室比較適當,LEDOMARS比較靠左
'Memo by Lydia 2015/04/17 台銀新版(104.04); 原名"水單列印",更名為"非台銀媒體水單列印"
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2011/6/28 modify by sonia 原有option2可選單張紙或連續紙,婧瑄說很久沒用單張紙,因列印位置不同,判斷應為舊版故刪除此選項
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

'Modified by Lydia 2023/04/13 Public改成Dim
Dim adoacc190 As New ADODB.Recordset
Dim adoacc190_1 As New ADODB.Recordset
Dim adoquery As New ADODB.Recordset
'end 2023/04/13
'Added by Lydia 2015/04/16 調整座標
'Private Const intY As Integer = 200
'Private Const intX As Integer = 300
Dim intY As Integer, intX As Integer
'Add By Cheng 2003/05/30
Dim m_strSQLA As String
'add by nickc 2005/05/17
Dim IsEnd As Boolean
Dim m_Data As Boolean   '2011/11/25 ADD BY SONIA 有抓到資料才清畫面
'Add By Sindy 2014/3/19
Dim m_FileName As String, m_TempFileName As String '智權華南
Dim m_A1803 As String, m_A1811 As String, m_A1810 As String
Dim m_NationName As String '受款地區國別
Dim m_BeneBankName1 As String, m_BeneBankName2 As String, m_BeneBankName3 As String '受款銀行名稱
Dim m_AccountNum As String '受款人帳號
Dim m_Payee_01  As String, m_Payee_02 As String, m_Payee_03 As String, m_Payee_04 As String '受款人名稱
Dim m_MiddleBank_01 As String, m_MiddleBank_02 As String, m_MiddleBank_03 As String '中間銀行
Dim m_FAAddr As String '代理人地址
Dim m_Amount As String '金額
'2014/3/19 END
Dim strPrinter As String 'Add By Sindy 2014/3/21
Dim excelSql As String 'Added by Lydia 2015/03/20
Dim strA2222 As String, tmpX As Double, tmpY As Double 'Added by Lydia 2015/03/30 媒體備註
Dim tmpCol As Integer, tmpNum As Integer 'Added by Lydia 2015/04/16 台銀新版位置 ,備註欄位數,最大備註數
Dim m_A2219 As String 'Added by Lydia 2015/04/30 改用 A2219(手續費方式)
Dim bolAddr As Boolean 'Added by Lydia 2015/05/12
Dim bolPayee71OUR As Boolean 'Modified by Lydia 2015/06/18 是否顯示足額到行(J公司)
Private Const NoteTitle As String = "INV."  'Added by Lydia 2017/07/20 DB note 最開頭
'Dim strDBMax As Integer 'Added by Lydia 2016/01/21 'Remove by Lydia 2016/02/15
'Added by Lydia 2017/09/14
Dim strPrtOrt As Integer '系統預設印表機的紙張方向
Dim m_A0k11Chi As String, m_A0k11Eng As String '公司名稱(中文,英文)
Dim m_A0k11Id As String, m_A0k11Tel As String  '公司統編,電話
Dim m_A0k11AddrC As String, m_A0k11AddrE As String '公司地址(中文,英文)
'Added by Lydia 2023/04/13 台銀水單改成Word套印
Dim m_WordLeft As Long, m_WordTop As Long 'Word開啟位置
Dim bVisible As Boolean
'Modified by Lydia 2025/07/31 台銀水單改版J20-01B(114.04)：19=>26
Private Const cntTB As Integer = 26
Dim strTempTB(1 To cntTB) As String
Dim tb_FileName As String, tb_TempFileName As String '臺銀華南
Dim m_AccNA01, m_AccNA02 As String  '受款地區國別,區別
Dim J_代辦人員 As String, J_代辦人員分機 As String, J_代辦人員ID As String
Dim mDefDir As String '預設資料夾
'Modified by Lydia 2023/04/13 從各模組改為全表單
Dim strCompanyName As String
Dim strCompanyNo As String
Dim strAddress As String
Dim strPhone As String
Dim strCompAddr1 As String, strCompAddr2 As String  'Added by Lydia 2017/09/06 公司別-英文地址
'Modified by Lydia 2024/04/23 開始啟用
Private Const 台銀Word水單啟用日 = 20240423 '= 20990413 'Added by Lydia 2023/04/13 啟用日控制; 等印表機色帶印完再換版本--溫斯閔
'Added by Lydia 2025/07/31
Dim m_BankCity As String, m_BankNA As String, m_BankMidCode As String '受款銀行-城市、國家代號、中間銀行Swift Code
Dim m_RecCity As String, m_RecNA As String '受款人-城市、國家代號

'Add By Sindy 2014/3/14
Private Sub JCallWordPrint(ByVal strDBNote As String)
Dim i As Integer
Dim strName As String
Dim strText As String
'Modified by Lydia 2015/06/18 移到最上方
'Dim bolPayee71OUR As Boolean 'Add By Sindy 2014/9/9
Dim iErrNo As Integer
Dim intA As Integer  'Added by Lydia 2017/09/14

On Error GoTo ErrHand
   
   '判斷word是否已開啟
   'Modified by Lydia 2023/04/13 改用模組
'   If g_WordAp Is Nothing Then
'RestarWord:
'      Set g_WordAp = New Word.Application
'      g_WordAp.Visible = False
'   End If
   If Pub_NewWordDoc(g_WordAp, bVisible, m_WordLeft, m_WordTop) = False Then Exit Sub
   'end 2023/04/13
   
   m_TempFileName = "$$智權華南匯出匯款申請書_" & strSrvDate(1) & ServerTime & ".doc"
   'Modified by Lydia 2023/04/13 App.path => mDefDir
   If Dir(mDefDir & "\" & m_TempFileName) <> "" Then
      Kill mDefDir & "\" & m_TempFileName
   End If
   g_WordAp.Documents.Open mDefDir & "\" & m_FileName
   g_WordAp.ActiveDocument.SaveAs mDefDir & "\" & m_TempFileName
   g_WordAp.ActiveDocument.Close
   g_WordAp.Documents.Open mDefDir & "\" & m_TempFileName
   'end 2023/04/13
   
   'Mark by Lydia 2018/05/28 找出特定TextBox名稱;保留程式，以後改版可以用
'   For intI = 1 To g_WordAp.ActiveDocument.Shapes.Count
'         If InStr(UCase(g_WordAp.ActiveDocument.Shapes(intI).Name), "TEXT") > 0 Then
'            strExc(1) = strExc(1) & "Name: " & g_WordAp.ActiveDocument.Shapes(intI).Name & vbCrLf & _
'                                  "     Text:" & g_WordAp.ActiveDocument.Shapes(intI).TextFrame.TextRange.Text
'         End If
'   Next intI
'   Debug.Print strExc(1)
   'end 2018/05/28
   
   With g_WordAp
      .Selection.WholeStory
      .Selection.Copy
      'Modified by Lydia 2018/05/28 增加其他資訊
      'For intA = 0 To 30
      For intA = 0 To 34
         strName = ""
         strText = ""
         If intA = 0 Then
            strName = "Y" '民國年
            strText = " " & String(3, " ") & " "
         ElseIf intA = 1 Then
            strName = "M" '月
            strText = " " & String(2, " ") & " "
         ElseIf intA = 2 Then
            strName = "D" '日
            strText = " " & String(2, " ") & " "
         ElseIf intA = 3 Then
            strName = "中文公司名"
            strText = m_A0k11Chi
         ElseIf intA = 4 Then
            strName = "公司電話"
            strText = m_A0k11Tel
         ElseIf intA = 5 Then
            strName = "英文公司名"
            strText = m_A0k11Eng
         ElseIf intA = 6 Then
            strName = "中文地址"
            strText = m_A0k11AddrC
         ElseIf intA = 7 Then
            strName = "英文地址"
            strText = m_A0k11AddrE
         ElseIf intA >= 8 And intA <= 17 Then '公司統編(1字1格)
            strName = "N" & Format(intA - 7, "00")
            If Len(m_A0k11Id) >= intA - 7 Then
               strText = Mid(m_A0k11Id, intA - 7, 1)
            Else
               strText = ""
            End If
         ElseIf intA = 18 Then
            strName = "受款人名稱1"
            strText = m_Payee_01
         ElseIf intA = 19 Then
            strName = "受款人名稱2"
            strText = m_Payee_02
         ElseIf intA = 20 Then
            strName = "受款人名稱3"
            strText = m_Payee_03
         ElseIf intA = 21 Then
            strName = "受款人國別"
            strText = m_NationName
         ElseIf intA = 22 Then
            strName = "地址"
            strText = m_FAAddr
         ElseIf intA = 23 Then
            'A1811.匯款方式:1.票匯 2.電匯
            strName = "電匯"
            strText = m_A1811
         ElseIf intA = 24 Then
            strName = "票匯"
            strText = m_A1811
         ElseIf intA = 25 Then
            strName = "存款帳號"
            strText = m_AccountNum
         ElseIf intA = 26 Then
            strName = "受款銀行"
            '拆二行, 受款銀行代號種類 ex:Swift code
            strText = Trim(m_BeneBankName1) & " " & Trim(m_BeneBankName2)
            If m_BeneBankName3 <> "" Then
               strText = strText & vbCrLf & Trim(Replace(m_BeneBankName3, "SWIFT CODE", "swift"))
            End If
            'CNAPS改位置到受款銀行代號後面
            If Trim(m_Payee_04) <> "" Then
               strText = strText & " CNAPS:" & Trim(m_Payee_04)
            End If
            If m_MiddleBank_01 <> "" Or m_MiddleBank_02 <> "" Or m_MiddleBank_03 <> "" Then '中間銀行
               strText = strText & vbCrLf & "correspondent bank：" & Trim(m_MiddleBank_01)
               If m_MiddleBank_02 <> "" Then
                  strText = strText & vbCrLf & "                      " & Trim(m_MiddleBank_02)
               End If
               If m_MiddleBank_03 <> "" Then
                  strText = strText & vbCrLf & "                      " & Trim(m_MiddleBank_03)
               End If
            End If
         ElseIf intA = 27 Then
            strName = "A1706" '付款明細
            strText = strDBNote
         ElseIf intA = 28 Then
            strName = "金額"
            strText = m_Amount & IIf(bolPayee71OUR = True, " 足額到行", "")
         ElseIf intA = 29 Then
            strName = "匯款性質"
            strText = Pub_DBtype
         'Added by Lydia 2018/05/28
         ElseIf intA = 30 Then
            strName = "受款人國別2"
            strText = m_NationName
         ElseIf intA = 31 Then
            strName = "關係"
            strText = "供應商"
         ElseIf intA = 32 Then
            strName = "代辦人"
            'Modified by Lydia 2023/04/13 改負責人
            'strText = "吳婉莘"
            strText = GetStaffName(J_代辦人員, True)
         ElseIf intA = 33 Then
            strName = "代辦人身份證號"
            'Modified by Lydia 2023/04/13 改負責人
            'strText = "Q222914536"
            strText = J_代辦人員ID
         'end 2018/05/28
         Else
            'A1803.代理人
            'Modified by Lydia 2018/05/28 若版面的文字欄位有變更,請使用前面"找出特定TextBox名稱"來確定欄位名稱
            '.ActiveDocument.Shapes("Text Box 43").Select
            'Modified by Lydia 2019/06/26 華銀新版面(108.7)
            '.ActiveDocument.Shapes("Text Box 69").Select
            .ActiveDocument.Shapes("Text Box 70").Select
            .Selection.ShapeRange.TextFrame.TextRange.Select
            .Selection.TypeText Text:=m_A1803
            .Selection.ShapeRange.Select
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
            If intA = 23 Then '電匯
               If strText = "2" Then
                  .Selection.InsertSymbol Font:="Wingdings", CharacterNumber:=-3843, Unicode:=True '已核取
               Else
                  .Selection.InsertSymbol Font:="Wingdings", CharacterNumber:=-3985, Unicode:=True '未核取
               End If
            ElseIf intA = 24 Then '票匯
               .Selection.TypeText " "
               If strText = "1" Then
                  .Selection.InsertSymbol Font:="Wingdings", CharacterNumber:=-3843, Unicode:=True '已核取
               Else
                  .Selection.InsertSymbol Font:="Wingdings", CharacterNumber:=-3985, Unicode:=True '未核取
               End If
            Else
                If intA >= 8 And intA <= 17 Then
                   .Selection.Font.Size = 12 '公司統編(恢復原本字型大小)
                End If
               .Selection.TypeText strText
            End If
         End If

      Next intA
'----------------------------
'end 2017/09/14

      iErrNo = 1
      '選項/列印的設定
      With .Options
          'Removed by Morgan 2014/11/26 2007 此參數無效,改調整範本的物件位置
          '.PrintDrawingObjects = True '不附帶印出繪圖物件
          'end 2014/11/26
      End With
      iErrNo = 2
      With .ActiveDocument
          'Removed by Morgan 2014/11/26 2007 此參數無效
          '.PrintPostScriptOverText = False
          '.PrintFormsData = False
          'end 2014/11/26
      End With
      iErrNo = 3
   End With
   '只列印第1頁,一式2份
   'Modified by Morgan 2014/11/26
   'g_WordAp.PrintOut FileName:="", Range:=wdPrintRangeOfPages, Item:=wdPrintDocumentContent, _
                     Copies:=2, Pages:="1", _
                     ManualDuplexPrint:=False, Collate:=True, Background:=False, PrintToFile:=False
   ''g_WordAp.ActiveDocument.Close
   'g_WordAp.ActiveDocument.Close wdDoNotSaveChanges
   g_WordAp.PrintOut Background:=False, Range:=4, Item:=0, Copies:=2, Pages:="1", Collate:=True
   'Modified by Lydia 2023/04/13
   'g_WordAp.ActiveDocument.Close wdSaveChanges
   Pub_RePosWord g_WordAp, bVisible, m_WordLeft, m_WordTop  '模組還原Word位置
   g_WordAp.Quit wdDoNotSaveChanges
   Set g_WordAp = Nothing '避免快速開啟Word,程式出錯
   'end 2014/11/26
   'end 2023/04/13
   
'   g_WordAp.Quit
'   Set g_WordAp = Nothing
   Exit Sub
   
ErrHand:
   'Modified by Lydia 2023/04/13 改用模組
   'If Err.Number = 462 Then '遠端伺服器不存在或無法使用
   '   GoTo RestarWord
   'ElseIf Err.Number <> 0 Then
   If Err.Number <> 0 Then
   'end 2023/04/13
      'Modified by Morgan 2014/11/26
      'MsgBox Err.Description
      MsgBox Err.Number & ":" & Err.Description, , "錯誤 " & iErrNo
   End If
End Sub

Private Sub Command2_Click()
'add by nick 2004/10/14
Dim strSql As String
Dim m_dbl_LeftMargin  As Double 'Added by Lydia 2015/04/16 橫軸偏移值
Dim m_dbl_TopMargin  As Double  'Added by Lydia 2015/04/16 縱軸偏移值

   If FormCheck = False Then
      'edit by nickc 2005/05/30
      'MsgBox MsgText(181), , MsgText(5)
      MsgBox "請輸入相關必要條件！", , MsgText(5)
      Exit Sub
   End If

   m_Data = False  '2011/11/25 ADD BY SONIA
   'add by nick 2004/10/14
   
   'Add By Sindy 2014/3/21
   PUB_SetOsDefaultPrinter Combo1
   PUB_RestorePrinter Combo1
   '2014/3/21 END
    
   'Added by Lydia 2023/04/13
   strExc(1) = PUB_GetOsDefaultPrinter
   If txtWordTB = "Y" And Text6 = "" Then
      '檢查是否為雙面列印的印表機
      If InStr(strExc(1), "雙面") = 0 And InStr(UCase(strExc(1)), UCase("PDFCreator")) = 0 And InStr(UCase(strExc(1)), UCase("PDF reDirect")) = 0 Then
          If MsgBox("選擇印表機:" & strExc(1) & vbCrLf & "臺銀水單需要雙面列印的印表機，請問是否要繼續列印？" & vbCrLf & "選""否""會重新選擇印表機！", vbYesNo + vbDefaultButton2 + vbInformation, "臺銀水單Word列印") = vbNo Then
              Exit Sub
          End If
      End If
   ElseIf Text6 = "J" And strSrvDate(1) >= 台銀Word水單啟用日 Then
      If InStr(strExc(1), "雙面") > 0 Then
          If MsgBox("選擇印表機:" & strExc(1) & vbCrLf & "華銀水單為單面列印，請問是否要繼續列印？" & vbCrLf & "選""否""會重新選擇印表機！", vbYesNo + vbDefaultButton2 + vbInformation, "華銀水單列印") = vbNo Then
              Exit Sub
          End If
      End If
   Else
   'end 2023/04/13
      'Added by Lydia 2015/04/16 調整座標
      '注意在原本的intX,intY 改成使用者自己調整(*-1)
      m_dbl_LeftMargin = CDbl(Me.txtXY(0).Text) * 576: m_dbl_TopMargin = CDbl(Me.txtXY(1).Text) * 576
      'Modified by Lydia 2017/09/14 欄位隱藏
      'If (m_dbl_LeftMargin = 0 And m_dbl_TopMargin = 0) Or (Option1(0).Value = True And Text6.Text = "J") Or (Option1(1).Value = True And Text5.Text = "J") Then
      If (m_dbl_LeftMargin = 0 And m_dbl_TopMargin = 0) Or Text6.Text = "J" Then
          'Modified by Lydia 2015/04/29 台銀新版(104.04)
         'intY = 200: intX = 300
         intY = 370: intX = 450
      Else
         If m_dbl_LeftMargin <> 0 Then
            intX = m_dbl_LeftMargin * -1
         Else
            'Modified by Lydia 2015/04/29 台銀新版(104.04)
            'intX = 300
            intX = 450
         End If
         If m_dbl_TopMargin <> 0 Then
            intY = m_dbl_TopMargin * -1
         Else
            'Modified by Lydia 2015/04/29 台銀新版(104.04)
            'intY = 200
            intY = 370
         End If
      End If
      'end 2015/04/16
   End If 'Added by Lydia 2023/04/13
   
   Screen.MousePointer = vbHourglass 'Move by Lydia 2023/04/13 從m_Data上方移下來
   If Option1(0).Value = True Then '一般列印
      PrintDataNewPaper
   Else
      PrintSumOnlyNewPaper
   End If
   'Modified by Lydia 2015/03/12 水單結匯清單EXCEL
   'PUB_ExcelSave
   ' Call PUB_ExcelSave2(Me.Name, m_strSQLA, Text6.Text) '2015/03/20
   'Memo by Lydia 2017/09/14 雖然1,2公司可以用4.華銀電匯紙本,因為是極特例,所以結匯excel仍是1,2公司
   Call PUB_ExcelSave2(Me.Name, excelSql, Text6.Text)
   
   'Add By Sindy 2014/3/21
   PUB_SetOsDefaultPrinter strPrinter
   'Modified by Lydia 2017/09/14 +strPrtOrt
   'Modifed by Lydia 2019/05/09 debug
   'PUB_RestorePrinter strPrint, strPrtOrt
   PUB_RestorePrinter strPrinter, strPrtOrt
   '2014/3/21 END
   
   Screen.MousePointer = vbDefault
   '2011/11/25 MODIFY BY SONIA
   'FormClear
   If m_Data = True Then FormClear
   '2011/11/25 END
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   End If
End Sub

Private Sub Form_Load()
'Modified by Lydia 2017/09/14 改成模組
'Memoed by Morgan 2017/11/8 舊程式已刪除方便檢查程式碼
strPrtOrt = Printer.Orientation
PUB_SetPrinter Me.Name, Combo1, strPrinter, , , Me.txtXY(0), Me.txtXY(1)
'end 2017/09/14

   'add by nickc 2005/05/17
   IsEnd = True
   strFormName = Name
   Me.Icon = LoadPicture(strIcoPath)
   'Modified by Lydia 2017/09/14 表單初始化
'   Me.Width = 5880
'   'Added by Lydia 2015/04/16
'   'Me.Height = 5055
'   'Modified by Lydia 2015/04/29 隱藏調整欄位
'   'Me.Height = 5925
'   Me.Height = 4995
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath4)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   PUB_InitForm Me, 5880, 5900, strBackPicPath4
   'end 2017/09/14
   
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   'Added by Lydia 2023/04/13
   mDefDir = App.path & "\" & strUserNum
   Pub_ChkExcelPath mDefDir
   'end 2023/04/13
   
   'Add By Sindy 2014/3/19
   m_FileName = "$$智權華南匯出匯款申請書.doc"
   'Modified by Lydia 2023/04/13 App.path => mDefDir
   If Dir(mDefDir & "\" & m_FileName) <> "" Then
      Kill mDefDir & "\" & m_FileName
   End If
   'end 2023/04/13
   'Modified by Lydia 2023/04/13 +, , mDefDir
   Call PUB_GetSampleFile(m_FileName, "M31-000001-0-00", , mDefDir)
   'PUB_SetPrinter Me.Name, Combo1, strPrinter 'Mark by Lydia 2017/09/14
   '2014/3/19 END
   
   'Added by Lydia 2023/04/13
   J_代辦人員 = "B1007"
   strSql = "Select st26, ed01 From staff, ExtensionData Where st01='" & J_代辦人員 & "' and st01=ed02(+) "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
       J_代辦人員ID = "" & RsTemp.Fields("st26")
       J_代辦人員分機 = "" & RsTemp.Fields("ed01")
   End If
   
   If strSrvDate(1) < 台銀Word水單啟用日 Then
      Label19.Visible = False
      Label20.Visible = False
      txtWordTB = ""
      Label18.Visible = False: txtWordTB.Visible = False
   Else
      If Pub_StrUserSt03 = "M51" Then
         'Label18.Visible = True: txtWordTB.Visible = True 'Mark by Lydia 2025/07/31 屬性已設定不顯示
         '上傳檔案
         'Modified by Lydia 2024/07/22 改用變數
         'intI = SaveImgByteFile("\\" & Pub_GetSpecMan("FTP_VOL_IP_LINUX") & "\PolyCOM\TaieNew\RptSample\M31-000015-0-00 臺銀-匯出匯款申請書.doc", "M31", "000015", "0", "00", "4", "1")
      Else
         Label18.Visible = False: txtWordTB.Visible = False
      End If
      txtWordTB = "Y"
      tb_FileName = "$$臺銀匯出匯款申請書.doc"
      If Dir(mDefDir & "\" & tb_FileName) <> "" Then
         Kill mDefDir & "\" & tb_FileName
      End If
      Call PUB_GetSampleFile(tb_FileName, "M31-000015-0-00", , mDefDir)
      'end 2023/04/13
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo ErrHand

   'Add By Sindy 2014/3/20
   If Not g_WordAp Is Nothing Then
      'g_WordAp.Visible = True
      g_WordAp.Quit
CloseWord:
      Set g_WordAp = Nothing
   End If
   '2014/3/20 END
   
   'Add By Sindy 2014/3/21
'   '印表機設回預設印表機
'   For Each prnPrint In Printers
'      If prnPrint.DeviceName = strPrinter Then
'         Set Printer = prnPrint
'      End If
'   Next
   'Added by Lydia 2015/04/16 調整座標
'   If Me.Combo1.Text <> Me.Combo1.Tag Then
'      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
'   End If
   '2014/3/21 END
    
    'Added by Lydia 2023/04/13 還原印表機設定
    If PUB_GetOsDefaultPrinter <> strPrinter Then
       PUB_SetOsDefaultPrinter strPrinter
       PUB_RestorePrinter strPrinter, strPrtOrt
    End If
    'end 2023/04/13
    
    '若有變動印表機或偏移值, 則更新列印設定
    If Me.Combo1.Text <> Me.Combo1.Tag Or Me.txtXY(0).Text <> Me.txtXY(0).Tag Or Me.txtXY(1).Text <> Me.txtXY(1).Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, Me.txtXY(0).Text, Me.txtXY(1).Text, Me.Combo1.Text
    End If
    Unload Me
      
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc2430 = Nothing
   '2011/6/28 ADD BY SONIA 回復原字體
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   '2011/6/28 end
   
   Exit Sub
   
ErrHand:
   If Err.Number = 462 Then '遠端伺服器不存在或無法使用
      GoTo CloseWord
   ElseIf Err.Number <> 0 Then
      MsgBox (Err.Description)
   End If
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text1 = ""
   Text2 = ""
   Text3 = ""
   Text4 = ""
   'add by nick 2004/10/14
   Text5 = ""
   Text6 = ""  '2011/11/25 ADD BY SONIA
   Option1(0).Value = True
   MaskEdBox1.SetFocus
   'Modified by Lydia 2015/04/17 配合單張,不預設Y
   'txtKind = "Y" 'Modified by Lydia 2015/03/06
   'Added by Lydia 2017/09/14 因為華銀也有媒體結匯,所以預設"必須紙本結匯"=Y
   txtKind = "Y"
End Sub

'2013/7/26 add by sonia
Private Sub MaskEdBox1_LostFocus()
   If MaskEdBox2.Text = MsgText(29) Then
      MaskEdBox2.Text = MaskEdBox1.Text
   End If
End Sub
'2013/7/26 end

'add by nick 2004/10/14
Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
        MaskEdBox1.Enabled = True
        MaskEdBox2.Enabled = True
        Text4.Enabled = True
        Text3.Enabled = True
        Text6.Enabled = True    '2011/11/25 ADD BY SONIA
        Text5.Enabled = False
        MaskEdBox1.SetFocus
      Case 1
        MaskEdBox1.Enabled = False
        MaskEdBox2.Enabled = False
        Text4.Enabled = False
        Text3.Enabled = False
         'Modified by Lydia 2018/06/04 公司別改成Text5
        'Text6.Enabled = False  '2011/11/25 ADD BY SONIA
        'Text5.Enabled = True
        'Text5.SetFocus
        Text6.SetFocus
      Case Else
   End Select
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme 'Added by Lydia 2017/09/29
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus()
   '2009/6/2 MODIFY BY SONIA 預設尾碼999
   'If Me.Text1.Text <> "" Then
   '    Me.Text2.Text = Me.Text1.Text
   'End If
   'Modify By Sindy 2014/8/11 999=>ZZZ
   'If Text1.Text <> "" Then Text2.Text = Left(Me.Text1.Text, 6) & "999"
   If Text1.Text <> "" Then
      Text1.Text = Left(Trim(Me.Text1.Text) & String(9, "0"), 9) 'Added by Lydia 2017/09/27
      Text2.Text = Left(Me.Text1.Text, 6) & "ZZZ"
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   'add by nick 2004/10/14
   If Option1(0).Value = True Then
      If MaskEdBox1.Text <> MsgText(29) Then
         FormCheck = True
         Exit Function
      End If
      If MaskEdBox2.Text <> MsgText(29) Then
         FormCheck = True
         Exit Function
      End If
      If Text3 <> MsgText(601) Then
         FormCheck = True
         Exit Function
      End If
      If Text4 <> MsgText(601) Then
         FormCheck = True
         Exit Function
      End If
   Else
      'Modified by Lydia 2018/06/04
      'If Text5 = MsgText(601) Then
      If Text6 = MsgText(601) Then
         FormCheck = False
         Exit Function
      'Add by Morgan 2005/11/14
      '因為跳過下面檢查，所以此處要設true
      Else
         FormCheck = True
         Exit Function
      End If
      
   End If
   'edit by nickc 2005/10/20 婧瑄說可以不必強制輸入
   '    If Text1 <> MsgText(601) Then
   '      FormCheck = True
   '      Exit Function
   '   End If
   '   If Text2 <> MsgText(601) Then
   '      FormCheck = True
   '      Exit Function
   '   End If

   FormCheck = False
End Function

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
   CloseIme 'Added by Lydia 2017/09/29
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub PrintDropLine(strText As String, intX As Integer, intY As Integer, intHeight As Integer)
Dim arrPrintText
Dim strPrintText As String
Dim ii As Integer
Dim jj As Integer
    
    If Trim(strText) <> "" Then
        jj = 0
        strPrintText = ""
        arrPrintText = Split(Trim(strText))
        For ii = LBound(arrPrintText) To UBound(arrPrintText)
            If arrPrintText(ii) <> "" Then
                If Len(strPrintText & arrPrintText(ii)) > 30 Then
                    Printer.CurrentX = intX
                    Printer.CurrentY = intY + (intHeight * jj)
                    Printer.Print strPrintText
                    jj = jj + 1
                    strPrintText = arrPrintText(ii)
                Else
                    strPrintText = Trim(strPrintText & " " & arrPrintText(ii))
                End If
            End If
        Next ii
        If strPrintText <> "" Then
            Printer.CurrentX = intX
            Printer.CurrentY = intY + (intHeight * jj)
            Printer.Print strPrintText
        End If
    End If
End Sub

'add by nick  2004/10/14
Private Sub Text5_GotFocus()
    TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   Cancel = False
   If Option1(1).Value = True Then
       If Trim(Text5.Text) <> "" Then
           Select Case Trim(Text5.Text)
           'Modify By Sindy 2014/3/13 +9,J
           Case "1", "2", "9", "J"
           Case Else
                   MsgBox "公司別只能 1 跟 2 跟 9 跟 J ！", , "錯誤！"
                   Cancel = True
           End Select
       End If
   End If
End Sub

'2011/11/25 ADD BY SONIA
Private Sub Text6_GotFocus()
    TextInverse Text6
    CloseIme 'Added by Lydia 2017/09/29
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   Cancel = False
   If Option1(1).Value = True Then
       If Trim(Text6.Text) <> "" Then
           Select Case Trim(Text6.Text)
           'Modify By Sindy 2014/3/13 +9,J
           Case "1", "2", "9", "J"
           Case Else
                   MsgBox "公司別只能 1 跟 2 跟 9 跟 J ！", , "錯誤！"
                   Cancel = True
           End Select
       End If
   End If
End Sub
'2011/11/25 END

'*************************************************
' 列印資料
'*************************************************
Private Sub PrintDataNewPaper()
Dim strNo As String
Dim intLength As Integer
'Mark by Lydia 2023/04/13
'Dim strAmount As String
'Dim strCompanyName As String
'Dim strCompanyNo As String
'Dim strAddress As String
'Dim strPhone As String
'Dim strCompAddr1 As String, strCompAddr2 As String  'Added by Lydia 2017/09/06 公司別-英文地址
'end 2023/04/13
Dim strSql As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim strA0K11 As String
Dim strSQLc As String
Dim strCompany As String '公司別
Dim blnNewPage As Boolean '是否印新頁
Dim ii As Integer
Dim strDBNote As String '代理人帳單號碼
'Dim m_Na01 As String 'Added by Lydia 2016/08/02 'Mark by Lydia 2023/04/13 改成m_AccNA01
'Dim m_NA02 As String 'Added by Lydia 2017/07/19 受款地區 'Mark by Lydia 2023/04/13 改成m_AccNA02

   StrSqlB = "Delete From ACCRPT427 Where r42701='" & strUserNum & "' "
   adoTaie.Execute StrSqlB
   '初始化公司別變數
   strCompany = ""
   strSql = ""
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a1802 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a1802 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and a1803 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a1803 <= '" & Text2 & "'"
   End If
   If Text4 <> MsgText(601) Then
      strSql = strSql & " and a1801 >= '" & Text4 & "'"
   End If
   If Text3 <> MsgText(601) Then
      strSql = strSql & " and a1801 <= '" & Text3 & "'"
   End If
   '2011/11/25 ADD BY SONIA
   
   If Text6 <> MsgText(601) Then
      strSql = strSql & " and a1917 = '" & Text6 & "'"
   'Add By Sindy 2014/3/13 若未輸入公司別，要加控制不可為J公司的條件
   Else
      strSql = strSql & " and a1917<>'J'"
   End If
   '2011/11/25 END
     
   strSql = strSql & " and a1811<>'6' " 'Added by Lydia 2024/09/03 排除匯款方式6-抵帳
   
    excelSql = strSql  'Added by Lydia 2015/03/20 結匯匯總表包含票匯+電匯(傳SQL條件)
    
   'Modified by Lydia 2015/03/06 +只印票匯資料
   'Modified by Lydia 2015/05/05 +台銀電匯紙本
   If txtKind = "Y" Then
      'Added by Lydia 2017/09/14 1,2公司有可能用4.華銀電匯紙本
      If Text6 = "J" Then
         strSql = Replace(strSql, " and a1917 = '" & Text6 & "'", " and ((a1917='J' and a1811=1) or a1811=4)")
      Else
      'end 2017/09/14
         'Modified by Lydia 2017/09/22 +台銀匯款方式 5-台銀合併結匯
         strSql = strSql & " and (a1811=1 or a1811=3 or a1811=5)"
      End If 'end 2017/09/14
   End If
      
    'Add By Cheng 2003/05/30
    'Modified by Lydia 2017/09/22 台銀匯款方式 5-台銀合併結匯,只印單張水單,不印合計水單
    'm_strSQLA = strSql
    m_strSQLA = Replace(strSql, "or a1811=5", "")
   
   'Added by Lydia 2018/09/14 台銀合併結匯,只印單張水單,不印合計水單
   If txtKind = "" Then
        m_strSQLA = m_strSQLA & " and a1811<> 5"
   End If
   
' 明細
    adoacc190.CursorLocation = adUseClient
    '代理人名稱(英-->中-->日)
'edit by nickc 2005/05/17 依照定稿語文
'edit by nickc 2005/06/06 改回原設定
    '2006/3/7 MODIFY BY SONIA
    'strSQLc = "select a1901, a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05) As FA05, Decode(FA05, Null, Null, fa63) As FA63, Decode(FA05, Null, Null, fa64) As FA64, Decode(FA05, Null, Null, fa65) As FA65, fa10, na03, a1903, a1810, a1811, sum(a1904) as Amount, a1917 As a0k11 from acc190, acc180, fagent, nation where a1901 = a1801 and substr(a1803, 1, 8) = fa01 and substr(a1803, 9, 1) = fa02 and fa10 = na01(+) and a1908 is null and a1810 is null" & strSQL & " group by a1803, a1901, Decode(fa05, Null, Nvl(FA04, FA06), FA05), Decode(FA05, Null, Null, fa63), Decode(FA05, Null, Null, fa64), Decode(FA05, Null, Null, fa65), fa10, na03, a1903, a1810, a1811, a1917 "
    'strSQLc = strSQLc & " Union select a1901, a1803, substr(a1810, 1, 30) as fa05, substr(a1810, 31, 30) as fa63, substr(a1810, 61, 30) as fa64, substr(a1810, 91, 30) as fa65, '' as fa10, na03, a1903, a1810, a1811, sum(a1904) as Amount, a1917 As a0k11 from acc190, acc180, fagent, nation where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01(+) and a1908 is null and a1810 is not null" & strSQL & " group by a1803, a1901, na03, a1903, a1810, a1811, a1917 "
    'strSQLc = strSQLc & " Order By a0k11, a1901 "
    'Modified by Morgan 2014/1/27 +代理人地址(澳洲電匯用)
    'modify by sonia 2014/7/10 +a1812,獨立水單以台幣結匯且合計水單不算此筆
    'Modified by Lydia 2015/10/06 +A1718 代為結匯之客戶編號(申請人)
'    strSQLc = "select a1901, a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05) As FA05, Decode(FA05, Null, Null, fa63) As FA63, Decode(FA05, Null, Null, fa64) As FA64, Decode(FA05, Null, Null, fa65) As FA65, fa10, na03, a1903, a1810, a1811, sum(a1904) as Amount, a1917 As a0k11,rtrim(max(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70)) Addr,a1812 from acc190, acc180, fagent, nation where a1901 = a1801 and A1803>'Y' AND substr(a1803, 1, 8) = fa01 and substr(a1803, 9, 1) = fa02 and fa10 = na01(+) and a1908 is null and a1810 is null" & strSql & " group by a1803, a1901, Decode(fa05, Null, Nvl(FA04, FA06), FA05), Decode(FA05, Null, Null, fa63), Decode(FA05, Null, Null, fa64), Decode(FA05, Null, Null, fa65), fa10, na03, a1903, a1810, a1811, a1917, a1812 "
'    strSQLc = strSQLc & " Union select a1901, a1803, substr(a1810, 1, 30) as fa05, substr(a1810, 31, 30) as fa63, substr(a1810, 61, 30) as fa64, substr(a1810, 91, 30) as fa65, '' as fa10, na03, a1903, a1810, a1811, sum(a1904) as Amount, a1917 As a0k11,rtrim(max(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70)) Addr,a1812 from acc190, acc180, fagent, nation where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and A1803>'Y' AND substr(a1803, 9, 1) = fa02 (+) and fa10 = na01(+) and a1908 is null and a1810 is not null" & strSql & " group by a1803, a1901, na03, a1903, a1810, a1811, a1917, a1812 "
'    strSQLc = strSQLc & " Union select a1901, a1803, Decode(CU05, Null, Nvl(CU04, CU06), CU05) As FA05, Decode(CU05, Null, Null, CU88) As FA63, Decode(CU05, Null, Null, CU89) As FA64, Decode(CU05, Null, Null, CU90) As FA65, CU10 AS FA10, na03, a1903, a1810, a1811, sum(a1904) as Amount, a1917 As a0k11,rtrim(max(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102))) Addr,a1812 from acc190, acc180, CUSTOMER, nation where a1901 = a1801 and A1803<'Y' AND substr(a1803, 1, 8) = CU01 and substr(a1803, 9, 1) = CU02 and CU10 = na01(+) and a1908 is null and a1810 is null" & strSql & " group by a1803, a1901, Decode(CU05, Null, Nvl(CU04, CU06), CU05), Decode(CU05, Null, Null, CU88), Decode(CU05, Null, Null, CU89), Decode(CU05, Null, Null, CU90), CU10, na03, a1903, a1810, a1811, a1917, a1812 "
'    strSQLc = strSQLc & " Union select a1901, a1803, substr(a1810, 1, 30) as fa05, substr(a1810, 31, 30) as fa63, substr(a1810, 61, 30) as fa64, substr(a1810, 91, 30) as fa65, '' as fa10, na03, a1903, a1810, a1811, sum(a1904) as Amount, a1917 As a0k11,rtrim(max(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102))) Addr,a1812 from acc190, acc180, CUSTOMER, nation where a1901 = a1801 and A1803<'Y' AND substr(a1803, 1, 8) = CU01 (+) and substr(a1803, 9, 1) = CU02 (+) and CU10 = na01(+) and a1908 is null and a1810 is not null" & strSql & " group by a1803, a1901, na03, a1903, a1810, a1811, a1917, a1812 "
'    strSQLc = strSQLc & " Order By a0k11,a1901"
    'Modified by Lydia 2016/08/02 抓na01
    'Modified by Lydia 2017/07/19 +na02
    'Modified by Lydia 2017/08/14 + 判斷是否為暫收款退費 -> sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt
    'Modified by Lydia 2017/09/05 debug 客戶地址抓錯欄位 cu04||' '||cu05||' '||cu06||' '||cu07||' '||cu08||' '||cu102 -> cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102
    'Modified by Lydia 2017/09/14 +acc080
    'strSQLc = "select a1901, a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05) As FA05, Decode(FA05, Null, Null, fa63) As FA63, Decode(FA05, Null, Null, fa64) As FA64, Decode(FA05, Null, Null, fa65) As FA65, fa10, na03, a1903, a1810, a1811, sum(a1904) as Amount, a1917 As a0k11,rtrim(max(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70)) Addr,a1812,a1718, na01, na02,sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt from acc190, acc180, fagent, nation, acc170 where a1901 = a1801 and A1803>'Y' AND substr(a1803, 1, 8) = fa01 and substr(a1803, 9, 1) = fa02 and fa10 = na01(+) and a1908 is null and a1810 is null and a1902=a1702(+)" & strSql & " group by a1803, a1901, Decode(fa05, Null, Nvl(FA04, FA06), FA05), Decode(FA05, Null, Null, fa63), Decode(FA05, Null, Null, fa64), Decode(FA05, Null, Null, fa65), fa10, na03, a1903, a1810, a1811, a1917, a1812, a1718, na01 , na02 "
    'strSQLc = strSQLc & " Union select a1901, a1803, substr(a1810, 1, 30) as fa05, substr(a1810, 31, 30) as fa63, substr(a1810, 61, 30) as fa64, substr(a1810, 91, 30) as fa65, '' as fa10, na03, a1903, a1810, a1811, sum(a1904) as Amount, a1917 As a0k11,rtrim(max(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70)) Addr,a1812,a1718, na01, na02,sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt from acc190, acc180, fagent, nation, acc170 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and A1803>'Y' AND substr(a1803, 9, 1) = fa02 (+) and fa10 = na01(+) and a1908 is null and a1810 is not null and a1902=a1702(+)" & strSql & " group by a1803, a1901, na03, a1903, a1810, a1811, a1917, a1812, a1718, na01 , na02 "
    'strSQLc = strSQLc & " Union select a1901, a1803, Decode(CU05, Null, Nvl(CU04, CU06), CU05) As FA05, Decode(CU05, Null, Null, CU88) As FA63, Decode(CU05, Null, Null, CU89) As FA64, Decode(CU05, Null, Null, CU90) As FA65, CU10 AS FA10, na03, a1903, a1810, a1811, sum(a1904) as Amount, a1917 As a0k11,rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)) Addr,a1812,a1718, na01, na02,sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt from acc190, acc180, CUSTOMER, nation ,acc170 where a1901 = a1801 and A1803<'Y' AND substr(a1803, 1, 8) = CU01 and substr(a1803, 9, 1) = CU02 and CU10 = na01(+) and a1908 is null and a1810 is null and a1902=a1702(+)" & strSql & " group by a1803, a1901, Decode(CU05, Null, Nvl(CU04, CU06), CU05), Decode(CU05, Null, Null, CU88), Decode(CU05, Null, Null, CU89), Decode(CU05, Null, Null, CU90), CU10, na03, a1903, a1810, a1811, a1917, a1812, a1718, na01 , na02 "
    'strSQLc = strSQLc & " Union select a1901, a1803, substr(a1810, 1, 30) as fa05, substr(a1810, 31, 30) as fa63, substr(a1810, 61, 30) as fa64, substr(a1810, 91, 30) as fa65, '' as fa10, na03, a1903, a1810, a1811, sum(a1904) as Amount, a1917 As a0k11,rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)) Addr,a1812,a1718, na01, na02,sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt from acc190, acc180, CUSTOMER, nation ,acc170 where a1901 = a1801 and A1803<'Y' AND substr(a1803, 1, 8) = CU01 (+) and substr(a1803, 9, 1) = CU02 (+) and CU10 = na01(+) and a1908 is null and a1810 is not null and a1902=a1702(+)" & strSql & " group by a1803, a1901, na03, a1903, a1810, a1811, a1917, a1812, a1718, na01 , na02 "
    'Added by Lydia 2017/09/30 華銀整批媒體RMB改CNY,紙本也配合
    If Text6 = "J" Then
        strExc(3) = "DECODE(A1903,'RMB','" & J_RMB & "',A1903) A1903"
        strExc(4) = "DECODE(A1903,'RMB','" & J_RMB & "',A1903)"
    Else
        strExc(3) = "A1903": strExc(4) = strExc(3)
    End If
    'end 2017/09/30
    'Modified by Lydia 2017/09/30 華銀整批媒體RMB改CNY,紙本也配合 a1903=> strExc(3),strexc(4)
    strSQLc = "select a1901, a1803, decode(a1810,null,Decode(fa05, Null, Nvl(FA04, FA06), FA05),substr(a1810,1,30)) As FA05, " & _
              "decode(a1810,null,Decode(FA05, Null, Null, fa63),substr(a1810,31,30)) As FA63, decode(a1810,null,Decode(FA05, Null, Null, fa64),substr(a1810,61,30)) As FA64, " & _
              "decode(a1810,null,Decode(FA05, Null, Null, fa65),substr(a1810,91,30)) As FA65, fa10, na03, " & strExc(3) & ", a1810, a1811, sum(a1904) as Amount," & _
              "a1917 As a0k11,rtrim(max(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70)) Addr,a1812,a1718, na01, na02,sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt " & _
              ",a0807,a0802,a0803,a0804,a0813,a0822,a0823 from acc190, acc180, fagent, nation, acc170,acc080 " & _
              "where a1901 = a1801 and A1803>'Y' AND substr(a1803, 1, 8) = fa01 and substr(a1803, 9, 1) = fa02 and fa10 = na01(+) and a1908 is null and a1902=a1702(+) and a0801=a1917" & strSql & _
              " group by a1803, a1901, decode(a1810,null,Decode(fa05, Null, Nvl(FA04, FA06), FA05),substr(a1810,1,30))," & _
              " decode(a1810,null,Decode(FA05, Null, Null, fa63),substr(a1810,31,30)), decode(a1810,null,Decode(FA05, Null, Null, fa64),substr(a1810,61,30))," & _
              " decode(a1810,null,Decode(FA05, Null, Null, fa65),substr(a1810,91,30)), fa10, na03, " & strExc(4) & ", a1810, a1811, a1917, a1812, a1718, na01,na02,a0807,a0802,a0803,a0804,a0813,a0822,a0823 "
    'Modified by Lydia 2017/09/30 華銀整批媒體RMB改CNY,紙本也配合 a1903=> strExc(3),strexc(4)
    'Modified by Lydia 2025/07/08 debug=>+ AND a0801=a1917
    strSQLc = strSQLc & " Union select a1901, a1803,decode(a1810,null,Decode(CU05, Null, Nvl(CU04, CU06), CU05),substr(a1810,1,30)) As FA05, " & _
              "decode(a1810,null,Decode(CU05, Null, Null, CU88),substr(a1810,31,30)) As FA63, decode(a1810,null,Decode(CU05, Null, Null, CU89),substr(a1810,61,30)) As FA64, " & _
              "decode(a1810,null,Decode(CU05, Null, Null, CU90),substr(a1810,91,30)) As FA65, CU10 AS FA10, na03, " & strExc(3) & ", a1810, a1811, sum(a1904) as Amount," & _
              "a1917 As a0k11,rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)) Addr,a1812,a1718, na01, na02,sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt " & _
              ",a0807,a0802,a0803,a0804,a0813,a0822,a0823 from acc190, acc180, CUSTOMER, nation,acc170,acc080 " & _
              "where a1901 = a1801 and A1803<'Y' AND substr(a1803, 1, 8) = CU01 and substr(a1803, 9, 1) = CU02 and CU10 = na01(+) and a1908 is null and a1902=a1702(+) AND a0801=a1917" & strSql & _
              " group by a1803, a1901, decode(a1810,null,Decode(CU05, Null, Nvl(CU04, CU06), CU05),substr(a1810,1,30))," & _
              " decode(a1810,null,Decode(CU05, Null, Null, CU88),substr(a1810,31,30)), decode(a1810,null,Decode(CU05, Null, Null, CU89),substr(a1810,61,30))," & _
              " decode(a1810,null,Decode(CU05, Null, Null, CU90),substr(a1810,91,30)), CU10, na03, " & strExc(4) & ", a1810, a1811, a1917, a1812, a1718, na01,na02,a0807,a0802,a0803,a0804,a0813,a0822,a0823 "
    'end 2017/09/14
    strSQLc = strSQLc & " Order By a0k11,a1901"
   adoacc190.Open strSQLc, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc190.RecordCount = 0 Then
      adoacc190.Close
      'Modified by Lydia 2018/02/22 區別訊息
      'MsgBox MsgText(28), , MsgText(5)
      MsgBox "查無資料可供列印!! ", , MsgText(5)
      
      'Added by Lydia 2015/03/17 避免2次查無資料訊息
       'm_strSQLA = MsgText(28) '2015/03/20
        excelSql = MsgText(28)
      Exit Sub
   End If
   
   m_Data = True   '2011/11/25 ADD BY SONIA
   'Add By Cheng 2003/05/16
   'edit by nickc 2005/05/17
   'MsgBox "準備列印公司別為< " & IIf("" & adoacc190.Fields("a0k11").Value = "", "無", "" & adoacc190.Fields("a0k11").Value) & " >的水單資料，請更換紙張!!!", vbExclamation + vbOKOnly
   'papersize = 204 是2000 的   15' * 12' 的紙 IBM5577 KC2
   '                  291         98
    
   'Modify by Morgan 2008/4/25
   'If Printer.PaperSize <> 291 Then
   '     Printer.PaperSize = 291
   '   End If
   'Modify by Morgan 2008/4/25
   'Printer.PaperSize = 291
   'Modified by Lydia 2023/04/13 +台銀水單改成Word套印
   'If Text6 = "J" Then
   If Text6 = "J" Or (txtWordTB = "Y" And Text6 = "") Then
      Printer.PaperSize = 9
   Else
      Printer.PaperSize = PUB_GetPaperSize(10)
   End If
   'end 2008/4/25
   'edit by nick 2004/07/21 修改字體
   '2011/6/28 MODIFY BY SONIA 台灣銀行要求改字體,全部改用Printer.Font.Bold = True
   '2011/7/13 modify by sonia 不要用粗體Printer.Font.Bold = True且全部用10號字
   'Printer.Font.Name = "細明體"
   '2011/8/12 modify by sonia 金額及帳號欄改為Times New Roman,12號字,其他仍為Arial,10號字
   'Printer.Font.Name = "Arial"
   'Printer.Font.Size = 10
   Printer.Font.Name = "Times New Roman"
   Printer.Font.Size = 12
   'Printer.Font.Bold = True
   
   strCompany = "" & adoacc190.Fields("a0k11").Value '公司別
   Do While adoacc190.EOF = False
      blnNewPage = True
      '若公司別不同時
      If strCompany <> "" & adoacc190.Fields("a0k11").Value Then
         'Modify By Sindy 2014/3/20
         'Modified by Lydia 2023/04/13 +台銀水單改成Word套印
         'If Text6 = "J" Then
         If Text6 = "J" Or (txtWordTB = "Y" And Text6 = "") Then
            If blnNewPage = False Then
               Printer.NewPage
            End If
         Else
         '2014/3/20 END
            Printer.NewPage
         End If
         blnNewPage = False
         IsEnd = False
         '列印合計
         PrintDataNewSumNewPaper strCompany
         IsEnd = True
         'MsgBox "準備列印公司別為< " & IIf("" & adoacc190.Fields("a0k11").Value = "", "無", "" & adoacc190.Fields("a0k11").Value) & " >的水單明細資料，請更換紙張!!!", vbExclamation + vbOKOnly
         strCompany = "" & adoacc190.Fields("a0k11").Value
      End If
      If strNo <> (adoacc190.Fields("a1901").Value & adoacc190.Fields("a1903").Value) Then '付款單號＋幣別
         If strNo <> "" Then
            If blnNewPage = True Then
               'Modify By Sindy 2014/3/20
               'Printer.NewPage
               'Modified by Lydia 2023/04/13 +台銀水單改成Word套印
               'If Text6 <> "J" Then
               If Text6 <> "J" And txtWordTB = "" Then
                  Printer.NewPage
               End If
               '2014/3/20 END
            End If
         End If
        '付款單號&幣別
         strNo = (adoacc190.Fields("a1901").Value & adoacc190.Fields("a1903").Value)
      End If
      
'*************************************************************************************
      'Modify By Sindy 2014/3/19 先抓資料,再依公司別跑不同的水單
'*************************************************************************************
      
      'Modified by Lydia 2023/04/13 改成模組
'      '代理人
'      If IsNull(adoacc190.Fields("a1803").Value) = False Then '代理人
'         m_A1803 = adoacc190.Fields("a1803").Value
'      Else
'         m_A1803 = ""
'      End If
'      '受款地區國別
'      If IsNull(adoacc190.Fields("na03").Value) = False Then
'         m_NationName = adoacc190.Fields("na03").Value
'      Else
'         m_NationName = ""
'      End If
'      'Added by Lydia 2016/08/02 受款地區國別代號
'      'Modified by Lydia 2023/04/13
'      'm_Na01 = "" & adoacc190.Fields("na01")
'      'If Len(m_Na01) > 3 Then m_Na01 = Mid(m_Na01, 1, 3)
'      m_AccNA01 = "" & adoacc190.Fields("na01")
'      If Len(m_AccNA01) > 3 Then m_AccNA01 = Mid(m_AccNA01, 1, 3)
'
'      'Added by Lydia 2017/07/19 受款地區
'      'Modified by Lydia 2023/04/13
'      'm_NA02 = "" & adoacc190.Fields("na02")
'      m_AccNA02 = "" & adoacc190.Fields("na02")
'
'      '匯款方式
'      'Added by Lydia 2015/04/17 +匯款方式3:台銀電匯紙本
'      'Modified by Lydia 2017/09/14 + 4.華銀電匯紙本
'      'Modified by Lydia 2017/10/02 + 5.台銀合併結匯
'      'If adoacc190.Fields("a1811").Value = "2" Or adoacc190.Fields("a1811").Value = "3" Or adoacc190.Fields("a1811").Value = "4" Then
'      If InStr("2,3,4,5", "" & adoacc190.Fields("a1811").Value) > 0 Then
'         m_A1811 = "2" '電匯
'      Else
'         m_A1811 = "1" '票匯
'      End If
'      '金額
'      If IsNull(adoacc190.Fields("Amount").Value) Then
'         '金額前不印幣別
'         m_Amount = adoacc190.Fields("a1903").Value & "0.00"
'      Else
'         '金額前印幣別
'         m_Amount = adoacc190.Fields("a1903").Value & Format(adoacc190.Fields("Amount").Value, FDollar)
'      End If
'      '中間銀行
'      m_MiddleBank_01 = ""
'      m_MiddleBank_02 = ""
'      m_MiddleBank_03 = ""
'      strA2222 = "" 'Added by Lydia 2015/03/30 媒體備註
'      m_A2219 = "" 'Added by Lydia 2015/04/30
'      'Modified by Lydia 2016/11/22 抓受款銀行國別m_NationName
'      'StrSqlB = "Select * From ACC220 Where a2201='" & adoacc190.Fields("a1803").Value & "' And a2202='" & adoacc190.Fields("a1903").Value & "' "
'      'Modified by Lydia 2017/07/19 + na02
'      'Modified by Lydia 2017/10/03 華銀整批媒體RMB改CNY
'      'StrSqlB = "Select a.*,na03,na02 From ACC220 a ,nation Where a2201='" & adoacc190.Fields("a1803").Value & "' And a2202='" & adoacc190.Fields("a1903").Value & "' and a2217=na01(+) "
'      StrSqlB = "Select a.*,na03,na02 From ACC220 a ,nation Where a2201='" & adoacc190.Fields("a1803").Value & "' And a2202='" & IIf(adoacc190.Fields("a1903").Value = J_RMB, "RMB", adoacc190.Fields("a1903").Value) & "' and a2217=na01(+) "
'
'      rsB.CursorLocation = adUseClient
'      rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsB.RecordCount > 0 Then
'         'Remove by Lydia 2015/08/11
'         'If "" & rsB("A2214").Value <> "" Or rsB("A2215").Value <> "" Or rsB("A2216").Value <> "" Then
'            'modify by sonia 2014/8/14 取消中間銀行前的1,2,3
'            If "" & rsB("A2214").Value <> "" Then
'               m_MiddleBank_01 = rsB("A2214").Value
'            End If
'            If "" & rsB("A2215").Value <> "" Then
'               m_MiddleBank_02 = rsB("A2215").Value
'            End If
'            If "" & rsB("A2216").Value <> "" Then
'               m_MiddleBank_03 = rsB("A2216").Value
'            End If
'            'Added by Lydia 2015/03/30 媒體備註
'            strA2222 = "" & UCase(Trim(rsB.Fields("a2222")))
'
'            'Added by Lydia 2015/04/30 改用 A2219(手續費方式)
'            If "" & rsB("A2219").Value <> "" Then
'               m_A2219 = rsB("A2219").Value
'            End If
'            'Added by Lydia 2019/10/03 匯款日幣, 以OUR方式結匯的, OUR要改成全額到行
'            If adoacc190.Fields("a1903").Value = "JPY" And UCase(m_A2219) = "71:OUR" Then
'                m_A2219 = "全額到行" '僅供台銀承辦人員查看,不改變媒體
'            End If
'            'end 2019/10/03
'
'            'Added by Lydia 2016/11/22 華南抓受款銀行國別
'            'Modified by Lydia 2017/03/23 統一抓受款銀行國別
'            'If Text6.Text = "J" And "" & rsB("na03").Value <> "" Then
'            If "" & rsB("na03").Value <> "" Then
'               m_NationName = "" & rsB("na03").Value
'            End If
'
'            'Added by Lydia 2017/07/19 +國家地區
'            If "" & rsB("na02").Value <> "" Then
'               'Modified by Lydia 2023/04/13
'               'm_NA02 = "" & rsB("na02").Value
'               m_AccNA02 = "" & rsB.Fields("na02")
'            End If
'            'end 2017/07/19
'
'        ' End If Remove by Lydia 2015/08/11
'      End If
'      If rsB.State <> adStateClosed Then rsB.Close
'      Set rsB = Nothing
'      '受款人:
'      m_BeneBankName1 = ""
'      m_BeneBankName2 = ""
'      m_BeneBankName3 = ""
'      m_AccountNum = ""
'      m_Payee_01 = ""
'      m_Payee_02 = ""
'      m_Payee_03 = ""
'      m_Payee_04 = ""
'      m_FAAddr = ""
'      m_A1810 = ""
'      Erase strTempTB 'Added by Lydia 2023/04/13 台銀水單改成Word套印
'      bolPayee71OUR = False  'added by Lydia 2015/06/18
'      'Added by Lydia 2017/09/14 智權-公司資料
'      m_A0k11Chi = Trim("" & adoacc190.Fields("a0802"))
'      m_A0k11Eng = Trim("" & adoacc190.Fields("a0803"))
'      m_A0k11Id = Trim("" & adoacc190.Fields("a0807"))
'      m_A0k11Tel = Trim("" & adoacc190.Fields("a0813")) & "#546" '電話#分機
'      m_A0k11AddrC = Trim(Replace("" & adoacc190.Fields("a0804"), "朱園里7鄰", ""))
'      m_A0k11AddrE = Trim("" & adoacc190.Fields("a0822")) & " " & Trim("" & adoacc190.Fields("a0823"))
'      'end 2017/09/14
'      'Added by Lydia 2021/08/27 寰華結匯水單本所地址不能出現ROC
'      If InStr("Y53374,", Left(Trim(adoacc190.Fields("a1803")), 6)) > 0 Then
'          m_A0k11AddrE = Replace(m_A0k11AddrE, ", R.O.C.", "")
'      End If
'      'end 2021/08/27
'
'      '若不為電匯
'      '2006/1/18 MODIFY BY SONIA 婧瑄說其他結匯不管是否電匯都抓A1810
'      'If "" & adoacc190.Fields("a1811").Value <> "2" Then
'      'Added by Lydia 2015/04/17 +匯款方式3:台銀電匯紙本
'      'If "" & adoacc190.Fields("a1811").Value <> "2" Or Len(adoacc190.Fields("a1803").Value) = 5 Then
'      'Modified by Lydia 2017/09/14  +匯款方式4:華銀電匯紙本
'      'Modified by Lydia 2017/10/02  +匯款方式5:台銀合併結匯
'      'If ("" & adoacc190.Fields("a1811").Value <> "2" And "" & adoacc190.Fields("a1811").Value <> "3" And "" & adoacc190.Fields("a1811").Value <> "4") Or Len(adoacc190.Fields("a1803").Value) = 5 Then
'      If InStr("2,3,4,5", "" & adoacc190.Fields("a1811").Value) = 0 Or Len(adoacc190.Fields("a1803").Value) = 5 Then
'      '2006/1/18 END
'         '抓受款人相關資料
'         'Modified by Lydia 2017/10/03 華銀整批媒體RMB改CNY
'         'StrSqlB = "Select * From ACC220 Where a2201='" & adoacc190.Fields("a1803").Value & "' And a2202='" & adoacc190.Fields("a1903").Value & "' "
'         StrSqlB = "Select * From ACC220 Where a2201='" & adoacc190.Fields("a1803").Value & "' And a2202='" & IIf(adoacc190.Fields("a1903").Value = J_RMB, "RMB", adoacc190.Fields("a1903").Value) & "' "
'
'         rsB.CursorLocation = adUseClient
'         rsB.Open StrSqlB, adoTaie, adOpenStatic, adLockReadOnly
'         If rsB.RecordCount > 0 Then
'            '受款人名稱
'            If IsNull(rsB.Fields("a2203").Value) = False Then
'               m_Payee_01 = UCase(Trim(rsB.Fields("a2203").Value))
'            End If
'            If IsNull(rsB.Fields("a2204").Value) = False Then
'               m_Payee_02 = UCase(Trim(rsB.Fields("a2204").Value))
'            End If
'            If IsNull(rsB.Fields("a2205").Value) = False Then
'               m_Payee_03 = UCase(Trim(rsB.Fields("a2205").Value))
'            End If
'            'Modifed by Lydia 2015/03/30 名稱改3欄,第4欄改為CNAPS
''            If IsNull(rsB.Fields("a2206").Value) = False Then
''               m_Payee_04 = UCase(Trim(rsB.Fields("a2206").Value))
''            End If
'            If IsNull(rsB.Fields("a2220").Value) = False Then
'               m_Payee_04 = UCase(Trim(rsB.Fields("a2220").Value))
'            End If
''            'Added by Lydia 2015/03/30 媒體備註
''            strA2222 = "" & UCase(Trim(rsB.Fields("a2222")))
'            'Added by Lydia 2015/06/18 +華南(J公司)71:our判斷
'            If UCase(Trim("" & rsB.Fields("a2219"))) = UCase("71:our") Then
'               bolPayee71OUR = True
'            End If
'            'Added by Lydia 2016/08/02 +大陸中文_ 大陸地區的匯款
'            'Modified by Lydia 2023/04/13 m_Na01=>m_AccNA01
'            If "" & rsB.Fields("a2217") = "020" Or ("" & rsB.Fields("a2217") = "" And m_AccNA01 = "020") Then '受款銀行國籍A2217優先判斷
'                If PUB_CheckStrNEC("" & rsB.Fields("a2203")) = True Then '受款人名稱有中文
'                   bolPayee71OUR = True
'                End If
'            End If
'         Else
'            '受款人名稱
'            If "" & adoacc190.Fields("A1810").Value <> "" Then
'               m_A1810 = UCase(Trim(adoacc190.Fields("A1810").Value))
'               m_Payee_01 = UCase(Trim(adoacc190.Fields("A1810").Value))
'            Else
'               If IsNull(adoacc190.Fields("fa05").Value) = False Then
'                  m_Payee_01 = UCase(Trim(adoacc190.Fields("fa05").Value))
'               End If
'               If IsNull(adoacc190.Fields("fa63").Value) = False Then
'                  m_Payee_02 = UCase(Trim(adoacc190.Fields("fa63").Value))
'               End If
'               If IsNull(adoacc190.Fields("fa64").Value) = False Then
'                  m_Payee_03 = UCase(Trim(adoacc190.Fields("fa64").Value))
'               End If
'                'Modifed by Lydia 2015/03/30 名稱改3欄,第4欄改為CNAPS(沒匯款銀行資料)
'               'If IsNull(adoacc190.Fields("fa65").Value) = False Then
'               '   m_Payee_04 = UCase(Trim(adoacc190.Fields("fa65").Value))
'               'End If
'            End If
'         End If
'         If rsB.State <> adStateClosed Then rsB.Close
'         Set rsB = Nothing
'      '若為電匯
'      Else
'         '抓受款人相關資料
'         'Modified by Lydia 2017/04/21 統一抓受款銀行國別
'         'StrSqlB = "Select * From ACC220 Where a2201='" & adoacc190.Fields("a1803").Value & "' And a2202='" & adoacc190.Fields("a1903").Value & "' "
'         'Modified by Lydia 2017/07/19 +na02
'         'Modified by Lydia 2017/10/03 華銀整批媒體RMB改CNY
'         'StrSqlB = "Select a.*,na03,na02 From ACC220 a ,nation Where a2201='" & adoacc190.Fields("a1803").Value & "' And a2202='" & adoacc190.Fields("a1903").Value & "' and a2217=na01(+) "
'         StrSqlB = "Select a.*,na03,na02 From ACC220 a ,nation Where a2201='" & adoacc190.Fields("a1803").Value & "' And a2202='" & IIf(adoacc190.Fields("a1903").Value = J_RMB, "RMB", adoacc190.Fields("a1903").Value) & "' and a2217=na01(+) "
'
'         rsB.CursorLocation = adUseClient
'         rsB.Open StrSqlB, adoTaie, adOpenStatic, adLockReadOnly
'         If rsB.RecordCount > 0 Then
'            '受款銀行名稱
'            If IsNull(rsB.Fields("a2208").Value) = False Then
'               m_BeneBankName1 = UCase(Trim(rsB.Fields("a2208").Value))
'            End If
'            If IsNull(rsB.Fields("a2209").Value) = False Then
'               m_BeneBankName2 = UCase(Trim(rsB.Fields("a2209").Value))
'            End If
'            '受款銀行帳號
'            If IsNull(rsB.Fields("a2210").Value) = False Then
'               m_BeneBankName3 = UCase(Trim(rsB.Fields("a2210").Value) & " " & Trim(rsB.Fields("a2211").Value))
'            End If
'            '受款人帳號
'            If IsNull(rsB.Fields("a2207").Value) = False Then
'               m_AccountNum = rsB.Fields("a2207").Value
'            End If
'            'Added by Lydia 2017/04/21 統一抓受款銀行國別
'            If "" & rsB("na03").Value <> "" Then
'               m_NationName = "" & rsB("na03").Value
'            End If
'
'            'Added by Lydia 2017/07/19 抓國家地區
'            If "" & rsB("na02").Value <> "" Then
'               'Modified by Lydia 2023/04/13
'               'm_NA02 = "" & rsB("na02").Value
'               m_AccNA02 = "" & rsB.Fields("na02")
'            End If
'
''            'Added by Lydia 2015/03/30 媒體備註
''            strA2222 = "" & UCase(Trim(rsB.Fields("a2222")))
'
'            'Added by Lydia 2015/06/18 +華南(J公司)71:our判斷
'            If UCase(Trim("" & rsB.Fields("a2219"))) = UCase("71:our") Then
'               bolPayee71OUR = True
'            End If
'            'Added by Lydia 2016/08/02 +大陸中文_ 大陸地區的匯款
'            'Modified by Lydia 2023/04/13 m_Na01=> m_AccNA01
'            If "" & rsB.Fields("a2217") = "020" Or ("" & rsB.Fields("a2217") = "" And m_AccNA01 = "020") Then '受款銀行國籍A2217優先判斷
'                If PUB_CheckStrNEC("" & rsB.Fields("a2203")) = True Then '受款人名稱有中文
'                   bolPayee71OUR = True
'                End If
'            End If
'            'Added by Morgan 2014/1/27
'            '澳洲代理人的水單要印地址,名稱上移
'            'modify by sonia 2014/9/3 加南非301
'            'Modified by Lydia 2015/03/25 + 加拿大
'            'Modified by Lydia 2015/05/12 改成依匯款行國別
'            'If adoacc190.Fields("fa10") = "015" Or adoacc190.Fields("fa10") = "301" Or adoacc190.Fields("fa10") = "102" Then
'            bolAddr = False
'            'Modified by Lydia 2015/08/05 若受款行資料未設定國別,以代理人資料判斷
'            'If InStr("015,301,102", rsB.Fields("a2217")) > 0 Then
'            'Modified by Lydia 2017/07/19 台銀要求於南非, 加拿大, 澳洲和歐洲地區(m_na02)要印地址,若有短地址則優先列印
'            'Modified by Lydia 2017/08/01 台銀要求全部都要印地址
'            'If InStr("015,301,102", rsB.Fields("a2217")) > 0 Or InStr("015,301,102", adoacc190.Fields("fa10")) > 0 Or m_NA02 = "C20" Then
'               bolAddr = True
'            'end 2015/05/12
'               'Modified by Lydia 2017/08/22 +第3行名稱
'               'Modified by Lydia 2017/09/14 華銀版面調整,受款人名稱分3行
'               If Text6.Text = "J" Then
'                  m_Payee_01 = Trim("" & rsB.Fields("a2203"))
'                  m_Payee_02 = Trim("" & rsB.Fields("a2204"))
'                  m_Payee_03 = Trim("" & rsB.Fields("a2205"))
'               Else
'               'end 2017/09/14
'                   m_Payee_01 = Trim("" & rsB.Fields("a2203") & " " & rsB.Fields("a2204")) & IIf("" & rsB.Fields("a2205") <> "", " " & rsB.Fields("a2205"), "")
'               End If 'end 2017/09/14
'               m_FAAddr = Trim("" & adoacc190.Fields("addr"))
'               'Added by Lydia 2017/07/19 短地址優先列印
'               If Trim("" & rsB.Fields("a2218")) <> "" Then m_FAAddr = Trim("" & rsB.Fields("a2218"))
'               m_FAAddr = Replace(m_FAAddr, "#", "") 'Added by Lydia 2017/09/18 配合華銀不接受#,預設拿掉#
'               m_Payee_04 = UCase(Trim("" & rsB.Fields("a2220").Value))  'Added by Lydia 2017/10/16
'            'Else
'            ''end 2014/1/27
'            '   '受款人名稱
'            'If IsNull(rsB.Fields("a2203").Value) = False Then
'            '   m_Payee_01 = UCase(Trim(rsB.Fields("a2203").Value))
'            'End If
'            'If IsNull(rsB.Fields("a2204").Value) = False Then
'            '   m_Payee_02 = UCase(Trim(rsB.Fields("a2204").Value))
'            'End If
'            'If IsNull(rsB.Fields("a2205").Value) = False Then
'            '   m_Payee_03 = UCase(Trim(rsB.Fields("a2205").Value))
'            'End If
'            '   'Modifed by Lydia 2015/03/30 名稱改3欄,第4欄(a2206)改為CNAPS(a2220)
''           '    If IsNull(rsB.Fields("a2206").Value) = False Then
''           '       m_Payee_04 = UCase(Trim(rsB.Fields("a2206").Value))
''           '    End If
'            '   If IsNull(rsB.Fields("a2220").Value) = False Then
'            '      m_Payee_04 = UCase(Trim(rsB.Fields("a2220").Value))
'            '   End If
'            'End If
'            'end 2017/07/31
'
'         End If
'         If rsB.State <> adStateClosed Then rsB.Close
'         Set rsB = Nothing
'      End If
'
'      '收據公司別
'      strA0K11 = "" & adoacc190.Fields("a0k11").Value '公司別
'      adoquery.CursorLocation = adUseClient
'      'Modified by Lydia 2015/10/06 有代為結匯之客戶編號,列印的申請人資料為該客戶的資料
'      'adoquery.Open "select * from acc080 where a0801 = '" & stra0k11 & "' ", adoTaie, adOpenStatic, adLockReadOnly
'      If "" & adoacc190.Fields("a1718").Value <> "" Then
'         ' 以\分2行列印
'         'Modified by Lydia 2015/12/04 有短地址就用短地址
'         'strSql = "select CU05||' '||CU88||' '||CU89||' '||CU90 as a0803,CU11 as a0807,NVL(CU16,CU17) as a0813,CU24||' '||CU25||'\'||CU26||' '||CU27||' '||CU28 as custaddr from customer where cu01||cu02=" & CNULL(adoacc190.Fields("a1718") & "0")
'         'Modified by Lydia 2017/10/03 華銀整批媒體RMB改CNY
'         'strSql = "select CU05||' '||CU88||' '||CU89||' '||CU90 as a0803,CU11 as a0807,NVL(CU16,CU17) as a0813,nvl(a2218,CU24||' '||CU25||'\'||CU26||' '||CU27||' '||CU28) as custaddr " & _
'                  "from customer,acc220 where cu01||cu02=" & CNULL(adoacc190.Fields("a1718") & "0") & " and cu01||cu02=a2201(+) "
'         strSql = "select CU05||' '||CU88||' '||CU89||' '||CU90 as a0803,CU11 as a0807,NVL(CU16,CU17) as a0813,nvl(a2218,CU24||' '||CU25||'\'||CU26||' '||CU27||' '||CU28) as custaddr " & _
'                  "from customer,acc220 where cu01||cu02=" & CNULL(adoacc190.Fields("a1718") & "0") & " and cu01||cu02=a2201(+) and a2202='" & IIf(adoacc190.Fields("a1903") = J_RMB, "RMB", adoacc190.Fields("a1903")) & "' "
'      Else
'         'Modified by Lydia 2020/09/03 收據公司別1,2,L都歸2公司
'         'strSql = "select * from acc080 where a0801 = '" & strA0K11 & "' "
'         strSql = "select * from acc080 where a0801 = '" & IIf(strA0K11 = "J", "J", "2") & "' "
'      End If
'      adoquery.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'      'end 2015/10/06
'      If adoquery.RecordCount <> 0 Then
'         strCompanyName = "" & adoquery.Fields("a0803").Value
'         If IsNull(adoquery.Fields("a0807").Value) Then
'            strCompanyNo = ""
'         Else
'            strCompanyNo = adoquery.Fields("a0807").Value
'         End If
'           ' strAddress = ReportSum(104)
'            'Modified by Lydia 2015/10/06
'            'strAddress = ReportSum(104)
'            If "" & adoacc190.Fields("a1718").Value <> "" Then
'               strAddress = "" & adoquery.Fields("custaddr")
'            Else
'               'Modified by Lydia 2017/09/14 公司別-中文,英文地址
'               'strAddress = ReportSum(104)
'               strAddress = Replace("" & adoquery.Fields("a0804").Value, "朱園里7鄰", "")
'               strCompAddr1 = "" & adoquery.Fields("a0822").Value
'               strCompAddr2 = "" & adoquery.Fields("a0823").Value
'            End If
'            'end 2015/10/06
'         If IsNull(adoquery.Fields("a0813").Value) Then
'            strPhone = ""
'         Else
'            strPhone = adoquery.Fields("a0813").Value
'         End If
'      Else
'         strCompanyName = ""
'         strCompanyNo = ""
'         strAddress = ""
'         strPhone = ""
'         'Added by Lydia 2017/09/06 公司別-英文地址
'         strCompAddr1 = ""
'         strCompAddr2 = ""
'      End If
'      'Added by Lydia 2017/09/06 公司別-英文地址(預設)
'      If strCompAddr1 & strCompAddr2 = "" Then
'         strCompAddr1 = "9F, No. 112, Sec. 2, Chang-An E. Rd.,"
'         strCompAddr2 = "Taipei 104, Taiwan, R.O.C."
'      End If
'      'end 2017/09/07
'      'Added by Lydia 2021/08/27 寰華結匯水單本所地址不能出現ROC
'      If InStr("Y53374,", Left(Trim(adoacc190.Fields("a1803")), 6)) > 0 Then
'          strCompAddr2 = Replace(strCompAddr2, ", R.O.C.", "")
'      End If
'      'end 2021/08/27
'      adoquery.Close
      Call GetAccData(adoacc190, strA0K11)
'end 2023/04/13

      'Add By Sindy 2014/3/20 新增華銀的水單
      If Text6 = "J" Then
         '抓代理人D/BNo
         '婧瑄說Y37580都不要印代理人D/BNo
         'Modified by Lydia 2018/04/23 寰華Y53374備註欄都不要再填寫(by 郭)
         'If Left(Trim(adoacc190.Fields("a1803").Value), 6) <> "Y37580" Then
         'Modified by Lydia 2020/04/22 +建毅Y51566,唯源Y52404
         If InStr("Y37580,Y53374,Y51566,Y52404", Left(Trim(adoacc190.Fields("a1803").Value), 6)) = 0 Then
            StrSqlB = "Select a1706 From acc190, acc170 Where a1902=a1702 And a1901='" & adoacc190.Fields("a1901").Value & "' Group By a1706 Order By 1 "
         Else
            StrSqlB = "select null from dual"
         End If
         rsB.CursorLocation = adUseClient
         rsB.Open StrSqlB, adoTaie, adOpenStatic, adLockReadOnly
         ii = 0
'         'Added by Lydia 2015/03/30 +媒體備註
'         If Len(Trim(strA2222)) > 0 Then
'            strDBNote = Trim(strA2222)
'         Else
'         'end 2015/03/30
            'Modified by Lydia 2017/07/20 DB note 加
            'Modified by Lydia 2017/08/01 票匯不加NoteTitle
            'strDBNote = "" & NoteTitle
            'Modified by Lydia 2017/08/14 暫收款退費不加INV. (a1902=a1702為O單號)
            'strDBNote = "" & IIf(m_A1811 = "1", "", NoteTitle)
            'Modified by Lydia 2018/04/23 寰華Y53374備註欄都不要再填寫(by 郭)
            'If m_A1811 <> "1" And Val("" & adoacc190.Fields("Ocnt")) = 0 Then
            'Modified by Lydia 2020/04/22 +天津三元Y37580; +建毅Y51566,唯源Y52404
            'If InStr("Y53374", Left(Trim(adoacc190.Fields("a1803").Value), 6)) = 0 And m_A1811 <> "1" And Val("" & adoacc190.Fields("Ocnt")) = 0 Then
            If InStr("Y37580,Y53374,Y51566,Y52404,", Left(Trim(adoacc190.Fields("a1803").Value), 6)) = 0 And m_A1811 <> "1" And Val("" & adoacc190.Fields("Ocnt")) = 0 Then
               strDBNote = NoteTitle
            'Added by Lydia 2017/11/1 預設清空
            Else
               strDBNote = ""
            End If
            'end 2017/08/14
            
'         End If
         Do While Not rsB.EOF
            ii = ii + 1
            If ii > 12 Then Exit Do
            'Modified by Lydia 2017/07/20 DB note 加NoteTitle
            'strDBNote = strDBNote & "," & rsB.Fields(0).Value
            'Modified by Lydia 2017/08/01 +strDBNote <> "" And
            strDBNote = strDBNote & IIf(strDBNote <> "" And strDBNote <> NoteTitle, ",", "") & rsB.Fields(0).Value
            strDBNote = strDBNote & IIf(ii = 12 And rsB.RecordCount > 12, " etc.", "")
            'Modified by Lydia 2016/01/20 備註輸入單引號
            'adoTaie.Execute "Insert Into ACCRPT427 Values('" & strUserNum & "','" & adoacc190.Fields("a0k11").Value & "','" & adoacc190.Fields("a1803").Value & "','" & rsB.Fields(0).Value & "' )"
            'Modified by Lydia 2018/04/23
            'adoTaie.Execute "Insert Into ACCRPT427 Values('" & strUserNum & "','" & adoacc190.Fields("a0k11").Value & "','" & adoacc190.Fields("a1803").Value & "','" & ChgSQL(rsB.Fields(0).Value) & "' )"
            adoTaie.Execute "Insert Into ACCRPT427 Values('" & strUserNum & "','" & adoacc190.Fields("a0k11").Value & "','" & adoacc190.Fields("a1803").Value & "','" & ChgSQL(strDBNote) & "' )"
            rsB.MoveNext
         Loop
         
         'If strDBNote <> "" Then strDBNote = Mid(strDBNote, 2) 'Remove by Lydia 2017/07/20
         If rsB.State <> adStateClosed Then rsB.Close
         Set rsB = Nothing
         
         Call JCallWordPrint(strDBNote)
      'Added by Lydia 2023/04/13 台銀水單改成Word套印
      ElseIf txtWordTB = "Y" Then
         Call TBCallWordPrint(adoacc190, "1")
      'end 2023/04/13
      Else
      '2014/3/20 END
      
         '代理人編號
         '2013/3/6 add by sonia 婧瑄說有時會不清楚,找不出原因故再寫一次
         Printer.Font.Name = "Times New Roman"
         Printer.Font.Size = 12
         '2013/3/6 end
         Printer.CurrentX = 7850 - intX
         'Modified by Lydia 2015/04/29 台銀新版(104.04)
         'Printer.CurrentY = 750 - intY
         Printer.CurrentY = 1255 - intY
         Printer.Font.Bold = True   '2013/2/19 ADD BY SONIA
         Printer.Print m_A1803
         Printer.Font.Bold = False  '2013/2/19 ADD BY SONIA
         '受款地區國別
         If m_NationName <> "" Then
            'Modified by Lydia 2015/04/29 台銀新版(104.04)
            'Printer.CurrentX = 1734 - intX
            'Printer.CurrentY = 1965 - intY
            'Modified by Lydia 2018/11/09 台銀新版(107.08)
            'Printer.CurrentX = 2504 - intX
            'Printer.CurrentY = 2135 - intY
            Printer.CurrentX = 2560 - intX
            Printer.CurrentY = 2240 - intY
            
            'Modified by Lydia 2017/04/17
            'Printer.Print adoacc190.Fields("na03").Value
            Printer.Print m_NationName
         End If
         '國外受款人身分別(民間)
         'Modified by Lydia 2015/04/29 台銀新版(104.04)
         'Printer.CurrentX = 7300 - intX
         'Printer.CurrentY = 1865 - intY
         'Modified by Lydia 2018/11/09 台銀新版(107.08) ; 移到申請人地址同一行
         'Printer.CurrentX = 7920 - intX
         'Printer.CurrentY = 2180 - intY
         Printer.CurrentX = 10930 - intX
         Printer.CurrentY = 4160 - intY
         Printer.Print "V"
         '匯款方式
         'Modified by Lydia 2015/04/29 台銀新版(104.04)
        ' Printer.CurrentY = 1865 - intY
         'Modified by Lydia 2018/11/09 台銀新版(107.08) ; 移到受款地區國別同一行
         'Printer.CurrentY = 2615 - intY
         Printer.CurrentY = 2240 - intY
         If m_A1811 = "2" Then
            '電匯
            'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
            'Printer.CurrentX = 9450 - intX
            'Modified by Lydia 2015/04/29 台銀新版(104.04)
            'Printer.CurrentX = 9310 - intX
            'Modified by Lydia 2018/11/09 台銀新版(107.08)
            'Printer.CurrentX = 8440 - intX
            Printer.CurrentX = 8880 - intX
         Else
            '票匯
            'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
            'Printer.CurrentX = 10625 - intX
            'Modified by Lydia 2015/04/29 台銀新版(104.04)
            'Printer.CurrentX = 10870 - intX
            'Modified by Lydia 2018/11/09 台銀新版(107.08)
            'Printer.CurrentX = 10430 - intX
            'Modified by Lydia 2018/12/28 調左
            'Printer.CurrentX = 10900 - intX
            Printer.CurrentX = 10840 - intX
         End If
         Printer.Print "V"
                  
         '申請人名稱:匯款人名稱(收據公司別)
         'Printer.Font.Name = "Arial"
         Printer.Font.Size = 10
         Printer.CurrentX = 1934 - intX
         'Modified by Lydia 2015/04/29 台銀新版(104.04)
         'Printer.CurrentY = 2395 - intY
         'Modified by Lydia 2018/11/09 台銀新版(107.08)
         'Printer.CurrentY = 2725 - intY
         Printer.CurrentY = 2820 - intY
         Printer.Print strCompanyName
         
         '外匯去處-匯往國外
         'Remove by Lydia 2018/11/09 台銀新版(107.08)
'         Printer.Font.Name = "Times New Roman"
'         Printer.Font.Size = 12
'         'Modified by Lydia 2015/04/29 台銀新版(104.04)
''         Printer.CurrentX = 7490 - intX
''         Printer.CurrentY = 2295 - intY
'         Printer.CurrentX = 7600 - intX
'         Printer.CurrentY = 3050 - intY
'         Printer.Print "V"
         'end -- Remove 2018/11/09
         
         '金額
         Printer.Font.Size = 12 'Added by Lydia 2018/11/09 台銀新版(107.08)
         intLength = Printer.TextWidth(m_Amount)
         'Modified by Lydia 2015/04/29 台銀新版(104.04)
         'Printer.CurrentY = 3175 - intY
         'Printer.CurrentX = 10050 - intLength - intX
         'Modified by Lydia 2018/11/09 台銀新版(107.08) ; 移到受款地區國別同一行
         'Printer.CurrentY = 3640 - intY
         'Printer.CurrentX = 10970 - intLength - intX
         Printer.CurrentY = 2240 - intY
         Printer.CurrentX = 6360 - intLength - intX
         Printer.Print m_Amount
         
         '統一編號
         'Modified by Lydia 2015/04/29 台銀新版(104.04)
        ' Printer.CurrentX = 3174 - intX
        ' Printer.CurrentY = 2720 - intY
         Printer.CurrentX = 3264 - intX
         Printer.CurrentY = 3320 - intY
         Printer.Print strCompanyNo
         '地址
         'Printer.Font.Name = "Arial"
         Printer.Font.Size = 10
         Printer.CurrentX = 1300 - intX
         'Modified by Lydia 2015/04/29 台銀新版(104.04)
         'Printer.CurrentY = 3170 - intY
         'Modified by Lydia 2018/11/09 台銀新版(107.08)
         'Printer.CurrentY = 3440 - intY
         Printer.CurrentY = 4250 - intY
         'Modified by Morgan 2011/11/23 +9 公司地址--婧瑄
         'Printer.Print "9F1., No. 112, Sec. 2, Chang-An E. Rd.,"
         'Modified by Lydia 2015/10/06 有代為結匯之客戶編號,列印的申請人資料為該客戶的資料
         'If stra0k11 = "9" Then
         If "" & adoacc190.Fields("a1718").Value <> "" Then
           '以\分2行列印
            'Modified by Lydia 2015/12/04
            'Printer.Print Mid(strAddress, 1, InStr(strAddress, "\") - 1)
            If InStr(strAddress, "\") = 0 Then
               Printer.Print strAddress
            Else
               Printer.Print Mid(strAddress, 1, InStr(strAddress, "\") - 1)
            End If
         'Modified by Lydia 2017/09/07 改成公司別-英文地址1
         'ElseIf stra0k11 = "9" Then
         ''end 2015/10/06
         '   Printer.Print "7F-1, No. 112, Sec. 2, Chang-An E. Rd.,"
         'ElseIf stra0k11 = "1" Then
         '   Printer.Print "10F, No. 112, Sec. 2, Chang-An E. Rd.,"
         ''2014/1/29 add by sonia
         'ElseIf stra0k11 = "J" Then
         '   Printer.Print "4F, No. 110, Sec. 2, Chang-An E. Rd.,"
         ''2014/1/29 end
         'Else
         '   Printer.Print "9F, No. 112, Sec. 2, Chang-An E. Rd.,"
         Else
             Printer.Print strCompAddr1
         'end 2017/09/06
         End If

         Printer.CurrentX = 1300 - intX
         'Modified by Lydia 2015/04/29 台銀新版(104.04)
         'Printer.CurrentY = 3400 - intY
         'Modified by Lydia 2018/11/09 台銀新版(107.08)
         'Printer.CurrentY = 3670 - intY
         Printer.CurrentY = 4520 - intY
        'Modified by Lydia 2015/10/06 有代為結匯之客戶編號,列印的申請人資料為該客戶的資料
        'Printer.Print "      Taipei 104, Taiwan, R.O.C."
         If "" & adoacc190.Fields("a1718").Value <> "" Then
            '以\分2行列印
            'Added by Lydia 2015/12/04
            If InStr(strAddress, "\") > 0 Then
               Printer.Print "      " & Mid(strAddress, InStr(strAddress, "\") + 1)
            End If
         Else
            'Modified by Lydia 2017/09/07 改成公司別-英文地址2
            'Printer.Print "      Taipei 104, Taiwan, R.O.C."
            Printer.Print "      " & strCompAddr2
         End If
        
         '電話
         Printer.Font.Name = "Times New Roman"
         Printer.Font.Size = 12
         'Modified by Lydia 2018/11/09 台銀新版(107.08)
         'Printer.CurrentX = 4800 - intX
         'Modified by Lydia 2018/12/28 調右
         'Printer.CurrentX = 1800 - intX
         Printer.CurrentX = 2200 - intX
         'Modified by Lydia 2015/04/29 台銀新版(104.04)
         'Printer.CurrentY = 3400 - intY
         'Added by Lydia 2015/10/06 有代為結匯之客戶編號,列印的申請人資料為該客戶的資料
         If "" & adoacc190.Fields("a1718").Value <> "" Then
             'Modified by Lydia 2018/11/09 台銀新版(107.08)
             'Printer.CurrentX = 2200 - intX
             'Printer.CurrentY = 3900 - intY
             'Modified by Lydia 2018/12/28 調右
             'Printer.CurrentX = 1800 - intX
             Printer.CurrentX = 2200 - intX
             Printer.CurrentY = 4850 - intY
             Printer.Print strPhone
         Else
             'Modified by Lydia 2018/11/09 台銀新版(107.08)
             'Printer.CurrentY = 3640 - intY
             'Printer.Print "Tel:" & strPhone
             Printer.CurrentY = 4850 - intY
             Printer.Print strPhone
         End If
         'end 2015/10/06
         
         '繳款方式
         'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
         'Printer.CurrentX = 7580 - intX
         'Modified by Lydia 2015/04/29 台銀新版(104.04)
         'Printer.CurrentX = 7460 - intX
         'Modified by Lydia 2018/11/09 台銀新版(107.08)
         'Printer.CurrentX = 7760 - intX
         Printer.CurrentX = 7560 - intX
         'add by sonia 2014/7/10 獨立水單以台幣結匯且合計水單不算此筆
         If adoacc190.Fields("a1812").Value = "Y" Then
            '以新台幣結購
            'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
            'Printer.CurrentY = 3567 - intY
            'Modified by Lydia 2015/04/29 台銀新版(104.04)
            'Printer.CurrentY = 3640 - intY
            'Modified by Lydia 2018/11/09 台銀新版(107.08)
            'Printer.CurrentY = 4080 - intY
            'Modified by Lydia 2019/07/08 調整位置
            'Printer.CurrentY = 2540 - intY
            Printer.CurrentY = 2730 - intY
         Else
            Select Case adoacc190.Fields("a1903").Value
               'Modify by Morgan 2010/6/15 取消歐元
               'Case "USD", "EUR"
               Case "USD"
                  '以外匯存款提出
                  'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                  'Printer.CurrentY = 3867 - intY
                  'Modified by Lydia 2015/04/29 台銀新版(104.04)
                 ' Printer.CurrentY = 3940 - intY
                  'Modified by Lydia 2018/11/09 台銀新版(107.08)
                  'Printer.CurrentY = 4460 - intY
                  Printer.CurrentY = 3050 - intY
               Case Else
                  '以新台幣結購
                  'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                  'Printer.CurrentY = 3567 - intY
                  'Modified by Lydia 2015/04/29 台銀新版(104.04)
                  'Printer.CurrentY = 3640 - intY
                  'Modified by Lydia 2018/11/09 台銀新版(107.08)
                  'Printer.CurrentY = 4080 - intY
                  Printer.CurrentY = 2730 - intY
            End Select
         End If  'add by sonia 2014/7/10 獨立水單以台幣結匯
        'Modified by Lydia 2015/10/06 有代為結匯之客戶編號,USD以新台幣結購
        If "" & adoacc190.Fields("a1718").Value <> "" Then
                'Modified by Lydia 2018/11/09 台銀新版(107.08)
                'Printer.CurrentY = 4080 - intY
                Printer.CurrentY = 2730 - intY
        End If
        Printer.Print "V"
        
         '金額
         'Modified by Lydia 2015/10/06  有代為結匯之客戶編號,USD以新台幣結購
         'If IsNull(adoacc190.Fields("a1812").Value) Then       '2014/7/28 ADD BY SONIA 非獨立水單且為美金才印金額
         If IsNull(adoacc190.Fields("a1812").Value) And IsNull(adoacc190.Fields("a1718").Value) Then
            Select Case adoacc190.Fields("a1903").Value
               'Modify by Morgan 2010/6/15 取消歐元
               'Case "USD", "EUR"
               Case "USD"
               'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                  'Printer.CurrentY = 3827 - intY
                  'Modified by Lydia 2015/04/29 台銀新版(104.04)
                  'Printer.CurrentY = 3900 - intY
                  'Modified by Lydia 2018/11/09 台銀新版(107.08)
                  'Printer.CurrentY = 4570 - intY
                  Printer.CurrentY = 3200 - intY
                  
                  If IsNull(adoacc190.Fields("Amount").Value) Then
                     '金額前不印幣別
                     m_Amount = adoacc190.Fields("a1903").Value & "0.00"
                     intLength = Printer.TextWidth(m_Amount)
                     'Modified by Lydia 2015/04/29 台銀新版(104.04)
                     'Printer.CurrentX = 11220 - intLength - intX
                  Else
                     '金額前印幣別
                     m_Amount = adoacc190.Fields("a1903").Value & Format(adoacc190.Fields("Amount").Value, FDollar)
                     intLength = Printer.TextWidth(m_Amount)
                     'Modified by Lydia 2015/04/29 台銀新版(104.04)
                     'Printer.CurrentX = 11220 - intLength - intX
                  End If
                  'Modified by Lydia 2018/11/09 台銀新版(107.08)
                  'Printer.CurrentX = 11520 - intLength - intX
                  Printer.CurrentX = 11300 - intLength - intX
                  Printer.Print m_Amount
               Case Else
                  'Remove by Lydia 2018/11/09 台銀新版(107.08) ; 只印美金幣別
                  'Printer.CurrentY = 3907 - intY
            End Select
         End If                                                '2014/7/28 ADD BY SONIA
         '中間銀行
         'Printer.Font.Name = "Arial"
         'Modified by Lydia 2015/04/29 台銀新版(104.04)
         'Modified by Lydia 2017/ 中間銀行Y座標
         'Modified by Lydia 2018/11/09 台銀新版(107.08)
         'tmpY = 5435: tmpX = 220  'tmpY暫存起始高度,tmpX暫存每行高度
         tmpY = 6230: tmpX = 220
         tmpY = tmpY + IIf(m_Payee_04 <> "", tmpX, 0) 'Added by Lydia 2017/10/16 中間銀行Y座標
         
         If m_MiddleBank_01 <> "" Or m_MiddleBank_02 <> "" Or m_MiddleBank_03 <> "" Then
            Printer.CurrentX = 1300 - intX
            'Modified by Lydia 2015/04/29 台銀新版(104.04)
            'Printer.CurrentY = 5035 - intY
            Printer.CurrentY = tmpY - intY: tmpY = tmpY + tmpX
            Printer.Print "correspondent bank:"
            If m_MiddleBank_01 <> "" Then
               '2011/6/28 改位置
               'Printer.CurrentX = 300 - intX + 2367
               'Printer.CurrentY = 7450 - intY - 2175
               Printer.CurrentX = 1300 - intX
               'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
              ' Printer.CurrentY = 5235 - intY
              'Modified by Lydia 2015/04/29 台銀新版(104.04)
               'Printer.CurrentY = 5335 - intY
               Printer.CurrentY = tmpY - intY: tmpY = tmpY + tmpX
               Printer.Print m_MiddleBank_01
            End If
            If m_MiddleBank_02 <> "" Then
               '2011/6/28 改位置
               'Printer.CurrentX = 300 - intX + 2367
               'Printer.CurrentY = 7650 - intY - 2175
               Printer.CurrentX = 1300 - intX
               'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
              ' Printer.CurrentY = 5455 - intY
              'Modified by Lydia 2015/04/29 台銀新版(104.04)
               'Printer.CurrentY = 5555 - intY
               Printer.CurrentY = tmpY - intY: tmpY = tmpY + tmpX
               Printer.Print m_MiddleBank_02
            End If
            If m_MiddleBank_03 <> "" Then
               '2011/6/28 改位置
               'Printer.CurrentX = 300 - intX + 2367
               'Printer.CurrentY = 7850 - intY - 2175
               Printer.CurrentX = 1300 - intX
               'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
              ' Printer.CurrentY = 5675 - intY
              'Modified by Lydia 2015/04/29 台銀新版(104.04)
               'Printer.CurrentY = 5775 - intY
               Printer.CurrentY = tmpY - intY: tmpY = tmpY + tmpX
               Printer.Print m_MiddleBank_03
            End If
         End If
         'Added by Lydia 2015/04/30
         If m_A2219 <> "" Then
            Printer.CurrentX = 1300 - intX
            Printer.CurrentY = tmpY - intY: tmpY = tmpY + tmpX
            Printer.Print m_A2219
         End If
         
         Printer.Font.Name = "Times New Roman"
         Printer.Font.Size = 12
         Printer.CurrentX = 7740 - intX
         'Modified by Lydia 2015/04/29 台銀新版(104.04)
         'Printer.CurrentY = 5725 - intY
         'Modified by Lydia 2018/11/09 台銀新版(107.08)
         'Printer.CurrentY = 6075 - intY
         Printer.CurrentY = 5350 - intY
         '2012/5/4 MODIFY BY SONIA 婧瑄說Y37580都不要印
         'Printer.Print "192 代理費支出"
         'Modified by Lydia 2015/10/06
         If "" & adoacc190.Fields("a1718").Value <> "" Then
            Printer.Print "19F"
         'Modified by Lydia 2018/04/23 寰華Y53374備註欄都不要再填寫(by 郭)
         'ElseIf Left(Trim(adoacc190.Fields("a1803").Value), 6) <> "Y37580" Then
         'Modified by Lydia 2018/06/11 寰華Y53374要印匯款類別
         'ElseIf InStr("Y37580,Y53374", Left(Trim(adoacc190.Fields("a1803").Value), 6)) = 0 Then
         ElseIf InStr("Y37580", Left(Trim(adoacc190.Fields("a1803").Value), 6)) = 0 Then
         'end 2015/10/06
            'Modified by Lydia 2017/04/07 匯款類別改為192 代理費支出
            'Printer.Print "19D 專業技術收入"    '2015/1/15 MODIFY BY SONIA 原為'192 代理費支出'
            'Modified by Lydia 2017/04/21 改回19D
            'Printer.Print "192 代理費支出"
            'Modified by Lydia 2017/09/14 改成常數Pub_DBtype
            'Printer.Print "19D 專業技術支出"
            Printer.Print Pub_DBtype
         Else
            Printer.Print ""
         End If
         '2012/5/4 end
         
         '受款人:
         '若不為電匯
         'Printer.Font.Name = "Arial"
         Printer.Font.Size = 10
         '2006/1/18 MODIFY BY SONIA 婧瑄說其他結匯不管是否電匯都抓A1810
         If m_A1811 <> "2" Or Len(m_A1803) = 5 Then
         '2006/1/18 END
            If Trim(m_Payee_01 & m_Payee_02 & m_Payee_03 & m_Payee_04) <> "" Then
               '受款人名稱
               If m_Payee_01 <> "" Then
                  '2011/6/28 改位置
                  'Printer.CurrentX = 300 - intX + 3550
                  'Printer.CurrentY = 5150 - intY + 1184
                  Printer.CurrentX = 1300 - intX
                  'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                  'Printer.CurrentY = 6634 - intY
                  'Modified by Lydia 2015/04/29 台銀新版(104.04)
                 ' Printer.CurrentY = 7734 - intY
                  'Modified by Lydia 2018/11/09 台銀新版(107.08)
                  'Printer.CurrentY = 8034 - intY
                  'Modified by Lydia 2018/12/28 往下調
                  'Printer.CurrentY = 8394 - intY
                  Printer.CurrentY = 8494 - intY
                  Printer.Print m_Payee_01
               End If
               If m_Payee_02 <> "" Then
                  'Modify by Morgan 2009/12/24 要和第一列對齊--婧瑄
                  '2011/6/28 改位置
                  'Printer.CurrentX = 300 - intX + 3550
                  'Printer.CurrentY = 5350 - intY + 1184
                  Printer.CurrentX = 1300 - intX
                  'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                  'Printer.CurrentY = 6854 - intY
                  'Modified by Lydia 2015/04/29 台銀新版(104.04)
                  'Printer.CurrentY = 7954 - intY
                  'Modified by Lydia 2018/11/09 台銀新版(107.08)
                  'Printer.CurrentY = 8254 - intY
                  'Modified by Lydia 2018/12/28 往下調
                  'Printer.CurrentY = 8614 - intY
                  Printer.CurrentY = 8714 - intY
                  Printer.Print m_Payee_02
               End If
               If m_Payee_03 <> "" Then
                  'Modify by Morgan 2009/12/24 要和第一列對齊--婧瑄
                  '2011/6/28 改位置
                  'Printer.CurrentX = 300 - intX + 3550
                  'Printer.CurrentY = 5550 - intY + 1184
                  Printer.CurrentX = 1300 - intX
                  'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                  'Printer.CurrentY = 7074 - intY
                  'Modified by Lydia 2015/04/29 台銀新版(104.04)
                  'Printer.CurrentY = 8174 - intY
                  'Modified by Lydia 2018/11/09 台銀新版(107.08)
                  'Printer.CurrentY = 8474 - intY
                  'Modified by Lydia 2018/12/28 往下調
                  'Printer.CurrentY = 8834 - intY
                  Printer.CurrentY = 8934 - intY
                  Printer.Print m_Payee_03
               End If
               If m_Payee_04 <> "" Then
                  'Modify by Morgan 2009/12/24 要和第一列對齊--婧瑄
                  '2011/6/28 改位置
                  'Printer.CurrentX = 300 - intX + 3550
                  'Printer.CurrentY = 5750 - intY + 1184
                  Printer.CurrentX = 1300 - intX
                  'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                  'Printer.CurrentY = 7294 - intY
                  'Modified by Lydia 2015/04/29 台銀新版(104.04)
                  'Printer.CurrentY = 8394 - intY
                  'Modified by Lydia 2017/10/16 改在受款銀行名稱(和中間銀行)下方
                  'Printer.CurrentY = 8694 - intY
                  Printer.CurrentY = tmpY - intY: tmpY = tmpY + tmpX
                  'Modifed by Lydia 2015/04/20 +抬頭CNAPS
                  Printer.Print "CNAPS:" & m_Payee_04
               End If
            Else
               '受款人名稱
               If m_A1810 <> "" Then
                  'Modify by Morgan 2006/8/16 全印大寫
                  'Modified by Lydia 2018/11/09 台銀新版(107.08)
                  'PrintDropLine m_Payee_01, 300 - intX + 3550, 5150 - intY + 1184, 200
                  PrintDropLine m_Payee_01, 300 - intX + 3550, 6150 - intY + 1184, 200
               Else
                  If m_Payee_01 <> "" Then
                     '2011/6/28 改位置
                     'Printer.CurrentX = 300 - intX + 3550
                     'Printer.CurrentY = 5150 - intY + 1184
                     Printer.CurrentX = 1300 - intX
                     'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                     'Printer.CurrentY = 6634 - intY
                     'Modified by Lydia 2015/04/29 台銀新版(104.04)
                     'Printer.CurrentY = 7734 - intY
                     'Modified by Lydia 2018/11/09 台銀新版(107.08)
                     'Printer.CurrentY = 8034 - intY
                     'Modified by Lydia 2018/12/28 往下調
                     'Printer.CurrentY = 8394 - intY
                     Printer.CurrentY = 8494 - intY
                     Printer.Print m_Payee_01
                  End If
                  If m_Payee_02 <> "" Then
                     '2011/6/28 改位置
                     'Printer.CurrentX = 300 - intX + 2300
                     'Printer.CurrentY = 5350 - intY + 1184
                     Printer.CurrentX = 1300 - intX
                     'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                     'Printer.CurrentY = 6854 - intY
                     'Modified by Lydia 2015/04/29 台銀新版(104.04)
                     'Printer.CurrentY = 7954 - intY
                     'Modified by Lydia 2018/11/09 台銀新版(107.08)
                     'Printer.CurrentY = 8254 - intY
                     'Modified by Lydia 2018/12/28 往下調
                     'Printer.CurrentY = 8614 - intY
                     Printer.CurrentY = 8714 - intY
                     Printer.Print m_Payee_02
                  End If
                  If m_Payee_03 <> "" Then
                     '2011/6/28 改位置
                     'Printer.CurrentX = 300 - intX + 1000
                     'Printer.CurrentY = 5550 - intY + 1184
                     Printer.CurrentX = 1300 - intX
                     'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                     'Printer.CurrentY = 7074 - intY
                     'Modified by Lydia 2015/04/29 台銀新版(104.04)
                     'Printer.CurrentY = 8174 - intY
                     'Modified by Lydia 2018/11/09 台銀新版(107.08)
                     'Printer.CurrentY = 8474 - intY
                     'Modified by Lydia 2018/12/28 往下調
                     'Printer.CurrentY = 8834 - intY
                     Printer.CurrentY = 8934 - intY
                     Printer.Print m_Payee_03
                  End If
                  If m_Payee_04 <> "" Then
                     '2011/6/28 改位置
                     'Printer.CurrentX = 300 - intX + 1000
                     'Printer.CurrentY = 5750 - intY + 1184
                     Printer.CurrentX = 1300 - intX
                     'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                     'Printer.CurrentY = 7294 - intY
                     'Modified by Lydia 2015/04/29 台銀新版(104.04)
                     'Printer.CurrentY = 8394 - intY
                     'Modified by Lydia 2017/10/16 改在受款銀行名稱(和中間銀行)下方
                     'Printer.CurrentY = 8694 - intY
                     Printer.CurrentY = tmpY - intY: tmpY = tmpY + tmpX
                     'Modifed by Lydia 2015/04/20 +抬頭CNAPS
                     Printer.Print "CNAPS:" & m_Payee_04
                  End If
               End If
            End If
            'Added by Lydia 2017/07/20 無條件列印
            'Remove by Lydia 2017/08/01 票匯不加
            'Printer.Font.Size = 12
            'Printer.CurrentX = 7210 - intX
            'Printer.CurrentY = 6680 - intY
            'Printer.Print IIf(m_A1811 = "1", "", NoteTitle)
            'end 2017/07/20
            
         '若為電匯
         Else
            'Printer.Font.Name = "Arial"
            Printer.Font.Size = 10
            '受款銀行名稱
            If m_BeneBankName1 <> "" Then
               '2011/6/28 改位置
               'Printer.CurrentX = 300 - intX + 1367
               Printer.CurrentX = 1300 - intX
               'Modified by Lydia 2015/04/29 台銀新版(104.04)
               'Printer.CurrentY = 4375 - intY
               'Modified by Lydia 2018/11/09 台銀新版(107.08)
               'Printer.CurrentY = 4775 - intY
               Printer.CurrentY = 5560 - intY
               Printer.Print m_BeneBankName1
            End If
            If m_BeneBankName2 <> "" Then
               '2011/6/28 改位置
               'Printer.CurrentX = 300 - intX + 1367
               Printer.CurrentX = 1300 - intX
               'Modified by Lydia 2015/04/29 台銀新版(104.04)
               'Printer.CurrentY = 4595 - intY
               'Modified by Lydia 2018/11/09 台銀新版(107.08)
               'Printer.CurrentY = 4995 - intY
               Printer.CurrentY = 5800 - intY
               Printer.Print m_BeneBankName2
            End If
            
            Printer.Font.Name = "Times New Roman"
            Printer.Font.Size = 12
            '受款銀行帳號
            If m_BeneBankName3 <> "" Then
               '2011/6/28 改位置
               'Printer.CurrentX = 300 - intX + 1367
               Printer.CurrentX = 1300 - intX
               'Modified by Lydia 2015/04/29 台銀新版(104.04)
               'Printer.CurrentY = 4815 - intY
               'Modified by Lydia 2018/11/09 台銀新版(107.08)
               'Printer.CurrentY = 5215 - intY
               Printer.CurrentY = 6010 - intY
               'Modified by Lydia 2017/09/14 SWIFT CODE 改為swift ,方便區隔後面的銀行代號
               'Printer.Print m_BeneBankName3
               Printer.Print Replace(m_BeneBankName3, "SWIFT CODE", "swift")
            End If
            
            'Added by Lydia 2017/10/16 CNAPS都顯示,改在受款銀行名稱(和中間銀行)下方
            If m_Payee_04 <> "" Then
               Printer.CurrentX = 1300 - intX
               Printer.CurrentY = tmpY - intY: tmpY = tmpY + tmpX
               Printer.Print "CNAPS:" & m_Payee_04
            End If
            'end 2017/10/16
            
            '受款人帳號
            If m_AccountNum <> "" Then
               'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
'               Printer.CurrentX = 4050 - intX
'               Printer.CurrentY = 5929 - intY
               Printer.CurrentX = 1300 - intX
               'Modified by Lydia 2015/04/29 台銀新版(104.04)
               'Printer.CurrentY = 7029 - intY
               'Modified by Lydia 2018/11/09 台銀新版(107.08)
               'Printer.CurrentY = 7279 - intY
               Printer.CurrentY = 7740 - intY
               Printer.Print m_AccountNum
            End If
            
            'Added by Morgan 2014/1/27
            '澳洲代理人的水單要印地址,名稱上移
            'modify by sonia 2014/9/3 加南非301
            'Modified by Lydia 2015/03/25 + 加拿大
            'Modified by Lydia 2015/05/12 改成依匯款行國別
            'If adoacc190.Fields("fa10") = "015" Or adoacc190.Fields("fa10") = "301" Or adoacc190.Fields("fa10") = "102" Then
            If bolAddr = True Then
               Printer.Font.Size = 10
               '受款人名稱
               'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
               'XPrint m_Payee_01, 3750, 6164, 3100
               'Modified by Lydia 2015/04/29 台銀新版(104.04)
               'XPrint m_Payee_01, 1300, 7264, 3100
               'Modified by Lydia 2018/11/09 台銀新版(107.08)
               'XPrint m_Payee_01, 3800, 7564, 3100
               XPrint m_Payee_01, 3800, 8080, 3100
               '代理人地址
               'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
               'XPrint m_FAAddr, 1300, 6634, 5670
               'Modified by Lydia 2015/04/29 台銀新版(104.04)
               'XPrint m_FAAddr, 1300, 7734, 5670
               'Modified by Lydia 2017/08/22 名稱多1行
               'XPrint m_FAAddr, 1300, 8034, 5670
               'Modified by Lydia 2018/11/09 台銀新版(107.08)
               'XPrint m_FAAddr, 1300, 8434, 5670
               XPrint m_FAAddr, 1300, 8600, 5670
               Printer.Font.Size = 12
            Else
            'end 2014/1/27
                'Printer.Font.Name = "Arial"
                Printer.Font.Size = 10
                '受款人名稱
                If m_Payee_01 <> "" Then
                   '2011/6/28 改位置
                   'Printer.CurrentX = 300 - intX + 3550
                   'Printer.CurrentY = 5150 - intY + 1184
                   Printer.CurrentX = 1300 - intX
                   'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                   'Printer.CurrentY = 6634 - intY
                   'Modified by Lydia 2015/04/29 台銀新版(104.04)
                  ' Printer.CurrentY = 7734 - intY
                   'Modified by Lydia 2018/11/09 台銀新版(107.08)
                   'Printer.CurrentY = 8034 - intY
                   Printer.CurrentY = 8600 - intY
                   Printer.Print m_Payee_01
                End If
                If m_Payee_02 <> "" Then
                   'Modify by Morgan 2009/12/24 要和第一列對齊--婧瑄
                   '2011/6/28 改位置
                   'Printer.CurrentX = 300 - intX + 3550
                   'Printer.CurrentY = 5350 - intY + 1184
                   Printer.CurrentX = 1300 - intX
                   'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                   'Printer.CurrentY = 6854 - intY
                    'Modified by Lydia 2015/04/29 台銀新版(104.04)
                    'Printer.CurrentY = 7954 - intY
                   'Modified by Lydia 2018/11/09 台銀新版(107.08)
                   'Printer.CurrentY = 8254 - intY
                   Printer.CurrentY = 8820 - intY
                   Printer.Print m_Payee_02
                End If
                If m_Payee_03 <> "" Then
                   'Modify by Morgan 2009/12/24 要和第一列對齊--婧瑄
                   '2011/6/28 改位置
                   'Printer.CurrentX = 300 - intX + 3550
                   'Printer.CurrentY = 5550 - intY + 1184
                   Printer.CurrentX = 1300 - intX
                   'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                   'Printer.CurrentY = 7074 - intY
                   'Modified by Lydia 2015/04/29 台銀新版(104.04)
                   'Printer.CurrentY = 8174 - intY
                   'Modified by Lydia 2018/11/09 台銀新版(107.08)
                   'Printer.CurrentY = 8474 - intY
                   Printer.CurrentY = 9120 - intY
                   Printer.Print m_Payee_03
                End If
                If m_Payee_04 <> "" Then
                   'Modify by Morgan 2009/12/24 要和第一列對齊--婧瑄
                   '2011/6/28 改位置
                   'Printer.CurrentX = 300 - intX + 3550
                   'Printer.CurrentY = 5750 - intY + 1184
                   Printer.CurrentX = 1300 - intX
                   'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                   'Printer.CurrentY = 7294 - intY
                   'Modified by Lydia 2015/04/29 台銀新版(104.04)
                   'Printer.CurrentY = 8394 - intY
                   'Modified by Lydia 2017/10/16 改在受款銀行名稱(和中間銀行)下方
                   'Printer.CurrentY = 8694 - intY
                   Printer.CurrentY = tmpY - intY: tmpY = tmpY + tmpX
                   'Modifed by Lydia 2015/04/20 +抬頭CNAPS
                   Printer.Print "CNAPS:" & m_Payee_04
                End If
            End If 'Added by Morgan 2014/1/27
            'Remove by Lydia 2015/04/30 改用 A2219(手續費方式)
'            'Add by Morgan 2006/7/14
'            If adoacc190.Fields("a1803").Value = "Y45589000" Then
'               Printer.CurrentX = 8517 - intX
'               'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
''               Printer.CurrentY = 6224 - intY
'               Printer.CurrentY = 7324 - intY
'               Printer.Print "ours"
'            End If
'            'end 2006/7/14
            
            'Modify by Lydia 2015/04/16 台銀新版(手寫紙),設備註欄數
            If rsB.State <> adStateClosed Then rsB.Close
            Set rsB = Nothing
            'Modified by Lydia 2018/04/23 寰華Y53374備註欄都不要再填寫(by 郭)
            'If Left(Trim(adoacc190.Fields("a1803").Value), 6) <> "Y37580" Then
            'Modified by Lydia 2020/04/22  +建毅Y51566,唯源Y52404
            If InStr("Y37580,Y53374,Y51566,Y52404,", Left(Trim(adoacc190.Fields("a1803").Value), 6)) = 0 Then
                StrSqlB = "Select max(length(a1706))  From acc190, acc170 Where a1902=a1702 And a1901='" & adoacc190.Fields("a1901").Value & "' Group By a1706 Order By 1 "
            Else
                StrSqlB = "select null from dual"
            End If
            rsB.CursorLocation = adUseClient
            rsB.Open StrSqlB, adoTaie, adOpenStatic, adLockReadOnly
            tmpCol = 4: tmpNum = 20
            If rsB.Fields(0) >= 9 Then
               tmpCol = 3: tmpNum = 15
            End If
            'Remove by Lydia 2016/02/15
            'strDBMax = rsB.Fields(0) 'Added by Lydia 2016/01/22
            If rsB.State <> adStateClosed Then rsB.Close
            Set rsB = Nothing
            'end 2015/04/16 台銀新版位置,設備註欄數
            
            Printer.Font.Name = "Times New Roman"
            Printer.Font.Size = 12
            '抓代理人D/BNo
            '2012/5/4 MODIFY BY SONIA 婧瑄說Y37580都不要印代理人D/BNo
            'StrSqlB = "Select a1504 From acc190, acc150 Where a1902=a1501 And a1901='" & adoacc190.Fields("a1901").Value & "' Group By a1504 Order By 1 "
            'Modified by Lydia 2018/04/23 寰華Y53374備註欄都不要再填寫(by 郭)
            'If Left(Trim(adoacc190.Fields("a1803").Value), 6) <> "Y37580" Then
            'Modified by Lydia 2020/04/22 +建毅Y51566,唯源Y52404
            If InStr("Y37580,Y53374,Y51566,Y52404,", Left(Trim(adoacc190.Fields("a1803").Value), 6)) = 0 Then
               'modify by sonia 2013/11/15 改抓a1706
               'StrSqlB = "Select a1504 From acc190, acc150 Where a1902=a1501 And a1901='" & adoacc190.Fields("a1901").Value & "' Group By a1504 Order By 1 "
               'Modified by Lydia 2015/10/06 +A1718,A1716
               StrSqlB = "Select a1706,a1718,a1716 From acc190, acc170 Where a1902=a1702 And a1901='" & adoacc190.Fields("a1901").Value & "' Group By a1706,a1718,a1716 Order By 1 "
            Else
               'Modified by Lydia 2018/01/15 避免抓不到值
               'StrSqlB = "select null from dual"
               StrSqlB = "select null as a1706, null as a1718, null as a1716 from dual"
            End If
            '2012/5/4 END
            rsB.CursorLocation = adUseClient
            rsB.Open StrSqlB, adoTaie, adOpenStatic, adLockReadOnly
            ii = 0
            'Added by Lydia 2015/04/01 先印媒體備註
            'tmpX = 7317: tmpY = 6430
            'tmpX = 7150: tmpY = 6400
            'Modified by Lydia 2015/04/29 台銀新版(104.04)
            'Modified by Lydia 2015/07/22 備註往下移
            'tmpX = 7210: tmpY = 6580
            'Modified by Lydia 2018/11/09 台銀新版(107.08)
            'tmpX = 7210: tmpY = 6680
            tmpX = 7210: tmpY = 6780
            If Len(Trim(strA2222)) > 0 Then
               Printer.CurrentX = tmpX - intX
               Printer.CurrentY = tmpY - intY
               'Modified by Lydia 2018/11/09 台銀新版(107.08) :長度從20改30
               If GetTextLength_1(strA2222) > 30 Then
                  strExc(9) = PUB_StrToStr(strA2222, 30)
                  strExc(10) = Trim(MidB(strA2222, LenB(strExc(9)) + 1))
                  Printer.Print strExc(9)
                  tmpY = tmpY + Printer.TextHeight(strExc(9)) + 30
               Else
                  strExc(10) = Trim(strA2222)
               End If
               Printer.CurrentX = tmpX - intX
               Printer.CurrentY = tmpY - intY
               Printer.Print PUB_StrToStr(strExc(10), 30)
               tmpY = tmpY + Printer.TextHeight(strExc(10)) + 30
               'end 2018/11/09
            End If
            'end 2015/04/01
            
            'Modified by Lydia 2017/07/20 DB note 加NoteTitle
            'Modified by Lydia 2017/08/01 票匯不加NoteTitle
            'strDBNote = "" & NoteTitle
            'Modified by Lydia 2017/08/14 暫收款退費不加INV. (a1902=a1702為O單號)
            'strDBNote = "" & IIf(m_A1811 = "1", "", NoteTitle)
            'Modified by Lydia 2018/04/23 寰華Y53374備註欄都不要再填寫(by 郭)
            'If m_A1811 <> "1" And Val("" & adoacc190.Fields("Ocnt")) = 0 Then
            'Modified by Lydia 2020/04/22 +天津三元Y37580; +建毅Y51566,唯源Y52404
            'If InStr("Y53374", Left(Trim(adoacc190.Fields("a1803").Value), 6)) = 0 And m_A1811 <> "1" And Val("" & adoacc190.Fields("Ocnt")) = 0 Then
            If InStr("Y37580,Y53374,Y51566,Y52404", Left(Trim(adoacc190.Fields("a1803").Value), 6)) = 0 And m_A1811 <> "1" And Val("" & adoacc190.Fields("Ocnt")) = 0 Then
               strDBNote = NoteTitle
            End If
            'end 2017/08/14
            
            'Added by Lydia 2017/07/20 無條件列印
            If rsB.EOF = True Then
               Printer.Font.Size = 12
               Printer.CurrentX = tmpX - intX
               Printer.CurrentY = tmpY - intY
               Printer.Print strDBNote
            End If
            'end 2017/07/20
            
            'Modify by Lydia 2015/04/16 台銀新版(手寫紙),設備註欄數
            '總欄位數12=>tmpNum ,一列2欄=>tmpCol
            Do While Not rsB.EOF
               ii = ii + 1
               'Modify by Lydia 2015/04/16
'               If ii > 12 Then Exit Do
'               If ii Mod 2 = 1 Then
               If ii > tmpNum Then Exit Do
               If ii Mod tmpCol > 0 Then
                  'Added by Lydia 2015/03/30 +媒體備註
                 ' strDBNote = strDBNote & PUB_RepToOneSpace(PUB_StringFilter("" & rsB.Fields(0).Value))
                 'Modified by Lydia 2017/07/20 DB note 加NoteTitle
                  'strDBNote = IIf(Len(strDBNote) > 0, strDBNote & ",", "") & PUB_RepToOneSpace(PUB_StringFilter("" & rsB.Fields(0).Value))
                  strDBNote = strDBNote & IIf(Len(strDBNote) > 0 And strDBNote <> NoteTitle, ",", "") & PUB_RepToOneSpace(PUB_StringFilter("" & rsB.Fields(0).Value))
                  'Modified by Lydia 2015/10/06
                  'adoTaie.Execute "Insert Into ACCRPT427 Values('" & strUserNum & "','" & adoacc190.Fields("a0k11").Value & "','" & adoacc190.Fields("a1803").Value & "','" & PUB_RepToOneSpace(PUB_StringFilter("" & rsB.Fields(0).Value)) & "' )"
               Else
                  strDBNote = strDBNote & "," & PUB_RepToOneSpace(PUB_StringFilter("" & rsB.Fields(0).Value))
                   'Modified by Lydia 2015/04/01
'                  Printer.CurrentX = 7317 - intX
'                  Printer.CurrentY = 6430 - intY + (ii / 2 - 1) * 200
                  Printer.CurrentX = tmpX - intX
                  'Modify by Lydia 2015/04/16
                  'Printer.CurrentY = tmpY - intY + (ii / 2 - 1) * Printer.TextHeight("W")
                  Printer.CurrentY = tmpY - intY + (ii / tmpCol - 1) * Printer.TextHeight("W")
                  'Modify By Sindy 2010/12/15
                  If Left(Trim(adoacc190.Fields("a1803").Value), 6) = "Y52401" Then
                     Printer.Print "Honorarium"
                  '2010/12/15 End
                  Else
                   'Modify by Lydia 2015/04/16
                     'Printer.Print strDBNote & IIf(ii = 12 And rsB.RecordCount > 12, " etc.", "")
                     Printer.Print strDBNote & IIf(ii = tmpNum And rsB.RecordCount > tmpNum, " etc.", "")
                  End If
                  'Modified by Lydia 2015/10/06
                  'adoTaie.Execute "Insert Into ACCRPT427 Values('" & strUserNum & "','" & adoacc190.Fields("a0k11").Value & "','" & adoacc190.Fields("a1803").Value & "','" & PUB_RepToOneSpace(PUB_StringFilter("" & rsB.Fields(0).Value)) & "' )"
                  strDBNote = ""
               End If
                  'Added by Lydia 2015/10/06 +A1716
                  If "" & rsB.Fields("a1718") <> "" Then
                     'Modified by Lydia 2016/01/20 備註輸入單引號
                     'strSql = "Insert Into ACCRPT427 Values('" & strUserNum & "','" & adoacc190.Fields("a0k11").Value & "','" & adoacc190.Fields("a1803").Value & "','" & rsB.Fields("a1706").Value & rsB.Fields("a1716").Value & "' )"
                     strSql = "Insert Into ACCRPT427 Values('" & strUserNum & "','" & adoacc190.Fields("a0k11").Value & "','" & adoacc190.Fields("a1803").Value & "','" & ChgSQL("" & rsB.Fields("a1706").Value & rsB.Fields("a1716").Value) & "' )"
                  Else
                     'Modified by Lydia 2016/01/20 備註輸入單引號
                     'strSql = "Insert Into ACCRPT427 Values('" & strUserNum & "','" & adoacc190.Fields("a0k11").Value & "','" & adoacc190.Fields("a1803").Value & "','" & PUB_RepToOneSpace(PUB_StringFilter("" & rsB.Fields(0).Value)) & "' )"
                     strSql = "Insert Into ACCRPT427 Values('" & strUserNum & "','" & adoacc190.Fields("a0k11").Value & "','" & adoacc190.Fields("a1803").Value & "','" & ChgSQL("" & rsB.Fields(0).Value) & "' )"
                  End If
                  adoTaie.Execute strSql
                  
               rsB.MoveNext
            Loop
            
            'If ii <= 10 And ii Mod 2 = 1 Then
            'Modify by Lydia 2015/04/16
            'If ii <= 12 And ii Mod 2 = 1 Then
            If ii <= tmpNum And ii Mod tmpCol > 0 Then
                 'Modified by Lydia 2015/04/01
'               Printer.CurrentX = 7317 - intX
'               Printer.CurrentY = 6430 - intY + ((ii + 1) / 2 - 1) * 200
               Printer.CurrentX = tmpX - intX
               'Modify by Lydia 2015/04/16
'               Printer.CurrentY = tmpY - intY + IIf(ii > 1, (ii \ 2), 0) * Printer.TextHeight("W")
               Printer.CurrentY = tmpY - intY + IIf(ii > 1, (ii \ tmpCol), 0) * Printer.TextHeight("W")
               'Modify By Sindy 2010/12/15
               If Left(Trim(adoacc190.Fields("a1803").Value), 6) = "Y52401" Then
                  Printer.Print "Honorarium"
               '2010/12/15 End
               Else
                  Printer.Print strDBNote
               End If
               'Modify by Lydia 2015/04/16
               'adoTaie.Execute "Insert Into ACCRPT427 Values('" & strUserNum & "','" & adoacc190.Fields("a0k11").Value & "','" & adoacc190.Fields("a1803").Value & "','" & strDBNote & "' )"
               strDBNote = ""
               tmpY = tmpY + IIf(ii > 1, (ii \ tmpCol), 0) * Printer.TextHeight("W") + Printer.TextHeight("W") + 20 'Added by Lydia 2016/01/22
            End If
            'Added by Lydia 2015/10/06 +A1716(A1718的備註)
            rsB.MoveFirst
            strExc(7) = ""
            Do While Not rsB.EOF
               If "" & rsB.Fields("a1718") <> "" Then
                  strExc(7) = strExc(7) & "," & rsB.Fields("a1716")
               End If
               rsB.MoveNext
            Loop
            If strExc(7) <> "" Then
                'Modified by Lydia 2015/12/04 備註折行
                'strExc(7) = Mid(strExc(7), 2)
                'Printer.CurrentX = tmpX - intX
                'Printer.CurrentY = tmpY - intY + Printer.TextHeight("W")
                'Printer.Print strExc(7)
                'Modified by Lydia 2016/01/26 備註折行(依a1716的換行符號)
'                strExc(7) = Mid(PUB_RepToOneSpace(PUB_StringFilter(strExc(7))), 2)
'                Do While Len(strExc(7)) > 0
'                    'Modified by Lydia 2016/01/20 大寫字太長
'                    'strExc(8) = Mid(strExc(7), 1, 40)
'                    strExc(8) = Mid(strExc(7), 1, 36)
'                    Printer.CurrentX = tmpX - intX
'                    Printer.CurrentY = tmpY - intY + Printer.TextHeight("W")
'                    Printer.Print strExc(8)
'                    'Modified by Lydia 2016/01/20 大寫字太長
'                    'strExc(7) = Mid(strExc(7), 41)
'                    strExc(7) = Mid(strExc(7), 37)
'                    tmpY = tmpY + Printer.TextHeight("W") + 20
'                Loop
         '-----------------------
                strExc(7) = Mid(strExc(7), 2)
                Do While strExc(7) <> ""
                   intI = InStr(strExc(7), vbCrLf)
                   If intI = 0 Then
                      Printer.CurrentX = tmpX - intX
                      Printer.CurrentY = tmpY - intY
                      Printer.Print strExc(7)
                      strExc(7) = ""
                   Else
                      strExc(8) = Left(strExc(7), intI - 1)
                      strExc(7) = Mid(strExc(7), intI + 2)
                      Printer.CurrentX = tmpX - intX
                      Printer.CurrentY = tmpY - intY
                      If tmpY >= 7500 Then
                          Printer.Print strExc(8) & " etc."
                          strExc(7) = ""
                      Else
                          Printer.Print strExc(8)
                      End If
                   End If
                   tmpY = tmpY + Printer.TextHeight("W") + 20
                Loop
            End If
            'end 2015/10/06
            
            If rsB.State <> adStateClosed Then rsB.Close
            Set rsB = Nothing
            Printer.Font.Size = 10
         End If
         '列印End
      End If
      adoacc190.MoveNext
   Loop
   'add by nick 2004/07/28 回復原字型
   'Printer.Font.Name = "細明體"   '2011/6/29 CANCEL
   'edit by nickc 2005/05/17
   'Printer.EndDoc
   'Modify By Sindy 2014/3/20
   'Modified by Lydia 2023/04/13 +台銀水單改成Word套印
   'If Text6 = "J" Then
   If Text6 = "J" Or (txtWordTB = "Y" And Text6 = "") Then
      'Add By Sindy 2014/8/26
      Clipboard.Clear
      If Not g_WordAp Is Nothing Then
         'g_WordAp.Visible = True
         g_WordAp.Quit
         Set g_WordAp = Nothing
      End If
      '2014/8/26 END
      If blnNewPage = False Then
         Printer.NewPage
      End If
   Else
   '2014/3/20 END
      Printer.NewPage
   End If
   adoacc190.Close
   'Add By Cheng 2003/05/30
   '列印合計
   PrintDataNewSumNewPaper strCompany
End Sub

'Create by nickc 2005/03/29 只印合計
Private Sub PrintSumOnlyNewPaper()
Dim strNo As String
Dim intLength As Integer
'Mark by Lydia 2023/04/13
'Dim strAmount As String
'Dim strCompanyName As String
'Dim strCompanyNo As String
'Dim strAddress As String
'Dim strPhone As String
'Dim strCompAddr1 As String, strCompAddr2 As String  '2017/09/06 公司別-英文地址
'end 2023/04/13
Dim strSql As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim strA0K11 As String
Dim strSQLc As String
Dim strCompany As String '公司別
Dim blnNewPage As Boolean '是否印新頁
Dim ii As Integer
Dim strDBNote As String '代理人帳單號碼

   StrSqlB = "Delete From ACCRPT427 Where r42701='" & strUserNum & "' "
   adoTaie.Execute StrSqlB
   '初始化公司別變數
   strCompany = ""
   strSql = ""
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and a1803 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a1803 <= '" & Text2 & "'"
   End If
   'Modified by Lydia 2017/09/14 欄位隱藏
   'If Text5 <> MsgText(601) Then
   '   strSql = strSql & " and a1917 = '" & Text5 & "'"
   If Text6 <> MsgText(601) Then
      strSql = strSql & " and a1917 = '" & Text6 & "'"
   'end 2017/09/14
   'Add By Sindy 2014/3/13 若未輸入公司別，要加控制不可為J公司的條件
   Else
      strSql = strSql & " and a1917<>'J'"
   End If
   
   strSql = strSql & " and a1811<>'6' " 'Added by Lydia 2024/09/03 排除匯款方式6-抵帳
   'Added by Lydia 2015/03/20 結匯匯總表包含票匯+電匯
    excelSql = strSql
    
   'Modified by Lydia 2015/03/06 +只印票匯資料
   'Modified by Lydia 2015/05/05 +台銀電匯紙本
   If txtKind = "Y" Then
      'Added by Lydia 2017/09/14 1,2公司有可能用4.華銀電匯紙本
      If Text6 = "J" Then
         strSql = Replace(strSql, " and a1917 = '" & Text6 & "'", " and ((a1917='J' and a1811=1) or a1811=4)")
      Else
      'end 2017/09/14
         strSql = strSql & " and (a1811=1 or a1811=3)"
      End If 'end 2017/09/14
   'Added by Lydia 2018/09/14 台銀合併結匯,只印單張水單,不印合計水單
   Else
         strSql = strSql & " and a1811<> 5"
   'end 2018/09/14
   End If
   
    m_strSQLA = strSql
    
' 明細
   adoacc190.CursorLocation = adUseClient
'edit by nickc 2005/05/17 依照定稿語文
'edit by nickc 2005/06/06 改回原設定
   '2006/3/7 MODIFY BY SONIA
   'strSQLc = "select a1901, a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05) As FA05, Decode(FA05, Null, Null, fa63) As FA63, Decode(FA05, Null, Null, fa64) As FA64, Decode(FA05, Null, Null, fa65) As FA65, fa10, na03, a1903, a1810, a1811, sum(a1904) as Amount, a1917 As a0k11 from acc190, acc180, fagent, nation where a1901 = a1801 and substr(a1803, 1, 8) = fa01 and substr(a1803, 9, 1) = fa02 and fa10 = na01(+) and a1908 is null and a1810 is null" & strSQL & " group by a1803, a1901, Decode(fa05, Null, Nvl(FA04, FA06), FA05), Decode(FA05, Null, Null, fa63), Decode(FA05, Null, Null, fa64), Decode(FA05, Null, Null, fa65), fa10, na03, a1903, a1810, a1811, a1917 "
   'strSQLc = strSQLc & " Union select a1901, a1803, substr(a1810, 1, 30) as fa05, substr(a1810, 31, 30) as fa63, substr(a1810, 61, 30) as fa64, substr(a1810, 91, 30) as fa65, '' as fa10, na03, a1903, a1810, a1811, sum(a1904) as Amount, a1917 As a0k11 from acc190, acc180, fagent, nation where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01(+) and a1908 is null and a1810 is not null" & strSQL & " group by a1803, a1901, na03, a1903, a1810, a1811, a1917 "
   'strSQLc = strSQLc & " Order By a0k11, a1901 "
   'Modified by Lydia 2023/04/13 為了共用模組+'' as A1812 ----參考modify by sonia 2014/7/10 +a1812,獨立水單以台幣結匯且合計水單不算此筆
   strSQLc = "select a1901, a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05) As FA05, Decode(FA05, Null, Null, fa63) As FA63, Decode(FA05, Null, Null, fa64) As FA64, Decode(FA05, Null, Null, fa65) As FA65, fa10, na03, a1903, a1810, a1811, sum(a1904) as Amount, a1917 As a0k11 from acc190, acc180, fagent, nation where a1901 = a1801 and A1803>'Y' AND substr(a1803, 1, 8) = fa01 and substr(a1803, 9, 1) = fa02 and fa10 = na01(+) and a1908 is null and a1810 is null" & strSql & " group by a1803, a1901, Decode(fa05, Null, Nvl(FA04, FA06), FA05), Decode(FA05, Null, Null, fa63), Decode(FA05, Null, Null, fa64), Decode(FA05, Null, Null, fa65), fa10, na03, a1903, a1810, a1811, a1917,'' as a1812 "
   strSQLc = strSQLc & " Union select a1901, a1803, substr(a1810, 1, 30) as fa05, substr(a1810, 31, 30) as fa63, substr(a1810, 61, 30) as fa64, substr(a1810, 91, 30) as fa65, '' as fa10, na03, a1903, a1810, a1811, sum(a1904) as Amount, a1917 As a0k11 from acc190, acc180, fagent, nation where a1901 = a1801 and A1803>'Y' AND substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01(+) and a1908 is null and a1810 is not null" & strSql & " group by a1803, a1901, na03, a1903, a1810, a1811, a1917,'' as a1812 "
   strSQLc = strSQLc & " Union select a1901, a1803, Decode(CU05, Null, Nvl(CU04, CU06), CU05) As FA05, Decode(CU05, Null, Null, CU88) As FA63, Decode(CU05, Null, Null, CU89) As FA64, Decode(CU05, Null, Null, CU90) As FA65, CU10, na03, a1903, a1810, a1811, sum(a1904) as Amount, a1917 As a0k11,'' as a1812 from acc190, acc180, CUSTOMER, nation where a1901 = a1801 and A1803<'Y' AND substr(a1803, 1, 8) = CU01 and substr(a1803, 9, 1) = CU02 and CU10 = na01(+) and a1908 is null and a1810 is null" & strSql & " group by a1803, a1901, Decode(CU05, Null, Nvl(CU04, CU06), CU05), Decode(CU05, Null, Null, CU88), Decode(CU05, Null, Null, CU89), Decode(CU05, Null, Null, CU90), CU10, na03, a1903, a1810, a1811, a1917 "
   strSQLc = strSQLc & " Union select a1901, a1803, substr(a1810, 1, 30) as fa05, substr(a1810, 31, 30) as fa63, substr(a1810, 61, 30) as fa64, substr(a1810, 91, 30) as fa65, '' as fa10, na03, a1903, a1810, a1811, sum(a1904) as Amount, a1917 As a0k11,'' as a1812 from acc190, acc180, CUSTOMER, nation where a1901 = a1801 and A1803<'Y' AND substr(a1803, 1, 8) = CU01 (+) and substr(a1803, 9, 1) = CU02 (+) and CU10 = na01(+) and a1908 is null and a1810 is not null" & strSql & " group by a1803, a1901, na03, a1903, a1810, a1811, a1917 "
   strSQLc = strSQLc & " Order By a0k11, a1901 "
   
   adoacc190.Open strSQLc, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc190.RecordCount = 0 Then
      adoacc190.Close
      'Modified by Lydia 2018/02/22 區別訊息
      'MsgBox MsgText(28), , MsgText(5)
      MsgBox "查無資料可供列印!! ", , MsgText(5)
      
      'Added by Lydia 2015/03/17 避免2次查無資料訊息
       'm_strSQLA = MsgText(28) '2015/03/20
        excelSql = MsgText(28)
        
      Exit Sub
   End If
   
   m_Data = True   '2011/11/25 ADD BY SONIA
   'Modify by Morgan 2008/4/25
'   If Printer.PaperSize <> 291 Then
'     Printer.PaperSize = 291
'   End If
   'Modified by Lydia 2023/04/13 +台銀水單改成Word套印
   'If Text6 = "J" Then
   If Text6 = "J" Or (txtWordTB = "Y" And Text6 = "") Then
      Printer.PaperSize = 9
   Else
      Printer.PaperSize = PUB_GetPaperSize(10)
   End If
   'end 0208/4/25
    strCompany = "" & adoacc190.Fields("a0k11").Value
   Do While adoacc190.EOF = False
      blnNewPage = True
      '若公司別不同時
      If strCompany <> "" & adoacc190.Fields("a0k11").Value Then
          '列印合計
          PrintDataNewSumNewPaper strCompany
          strCompany = "" & adoacc190.Fields("a0k11").Value
          blnNewPage = False
      End If
      If strNo <> (adoacc190.Fields("a1901").Value & adoacc190.Fields("a1903").Value) Then
        '付款單號&幣別
         strNo = (adoacc190.Fields("a1901").Value & adoacc190.Fields("a1903").Value)
      End If
      'Mark by Lydia 2023/04/13 改成模組GetAccData
'      strA0K11 = "" & adoacc190.Fields("a0k11").Value
'      adoquery.CursorLocation = adUseClient
'      'Modified by Lydia 2020/09/03 收據公司別1,2,L都歸2公司
'      'adoquery.Open "select * from acc080 where a0801 = '" & strA0K11 & "' ", adoTaie, adOpenStatic, adLockReadOnly
'      adoquery.Open "select * from acc080 where a0801 = '" & IIf(strA0K11 = "J", "J", "2") & "' ", adoTaie, adOpenStatic, adLockReadOnly
'      If adoquery.RecordCount <> 0 Then
'         strCompanyName = "" & adoquery.Fields("a0803").Value
'         If IsNull(adoquery.Fields("a0807").Value) Then
'            strCompanyNo = ""
'         Else
'            strCompanyNo = adoquery.Fields("a0807").Value
'         End If
'         'Modified by Lydia 2017/09/14 公司別-中文,英文地址
'         'strAddress = ReportSum(104)
'         strAddress = Replace("" & adoquery.Fields("a0804").Value, "朱園里7鄰", "")
'         strCompAddr1 = "" & adoquery.Fields("a0822").Value
'         strCompAddr2 = "" & adoquery.Fields("a0823").Value
'         'end 2017/09/14
'
'         If IsNull(adoquery.Fields("a0813").Value) Then
'            strPhone = ""
'         Else
'            strPhone = adoquery.Fields("a0813").Value
'         End If
'      Else
'         strCompanyName = ""
'         strCompanyNo = ""
'         strAddress = ""
'         strPhone = ""
'         'Added by Lydia 2017/09/06 公司別-英文地址
'         strCompAddr1 = ""
'         strCompAddr2 = ""
'      End If
'      'Added by Lydia 2017/09/06 公司別-英文地址(預設)
'      If strCompAddr1 & strCompAddr2 = "" Then
'         strCompAddr1 = "9F, No. 112, Sec. 2, Chang-An E. Rd.,"
'         strCompAddr2 = "Taipei 104, Taiwan, R.O.C."
'      End If
'      'end 2017/09/07
'      'Added by Lydia 2021/08/27 寰華結匯水單本所地址不能出現ROC
'      If InStr("Y53374,", Left(Trim(adoacc190.Fields("a1803")), 6)) > 0 Then
'          strCompAddr2 = Replace(strCompAddr2, ", R.O.C.", "")
'      End If
'      'end 2021/08/27
'      adoquery.Close
'
'      If IsNull(adoacc190.Fields("Amount").Value) Then
'            '金額前不印幣別
'         'Modified by Lydia 2023/04/13 strAmount=> m_Amount
'         m_Amount = adoacc190.Fields("a1903").Value & "0.00"
'      Else
'            '金額前不印幣別
'         'Modified by Lydia 2023/04/13 strAmount=> m_Amount
'         m_Amount = adoacc190.Fields("a1903").Value & Format(adoacc190.Fields("Amount").Value, FDollar)
'      End If
'      If rsB.State <> adStateClosed Then rsB.Close
'      Set rsB = Nothing
'      Select Case adoacc190.Fields("a1903").Value
'         Case "USD", "EUR"
'            If IsNull(adoacc190.Fields("Amount").Value) Then
'                  '金額前不印幣別
'               'Modified by Lydia 2023/04/13 strAmount=> m_Amount
'               m_Amount = adoacc190.Fields("a1903").Value & "0.00"
'            Else
'                  '金額前不印幣別
'               'Modified by Lydia 2023/04/13 strAmount=> m_Amount
'               m_Amount = adoacc190.Fields("a1903").Value & Format(adoacc190.Fields("Amount").Value, FDollar)
'            End If
'         Case Else
'      End Select
'      '若不為電匯
'      'Added by Lydia 2015/04/17 +匯款方式3:台銀電匯紙本
'      If "" & adoacc190.Fields("a1811").Value <> "2" And "" & adoacc190.Fields("a1811").Value <> "3" Then
'      '若為電匯
'      Else
'          If rsB.State <> adStateClosed Then rsB.Close
'          Set rsB = Nothing
'          '抓代理人D/BNo
'          '2012/5/4 MODIFY BY SONIA 婧瑄說Y37580都不要印代理人D/BNo
'          'StrSqlB = "Select a1504 From acc190, acc150 Where a1902=a1501 And a1901='" & adoacc190.Fields("a1901").Value & "' Group By a1504 Order By 1 "
'          'Modified by Lydia 2018/04/23 寰華Y53374備註欄都不要再填寫(by 郭)
'          'If Left(Trim(adoacc190.Fields("a1803").Value), 6) <> "Y37580" Then
'          'Modified by Lydia 2020/04/22 +建毅Y51566,唯源Y52404
'          If InStr("Y37580,Y53374,Y51566,Y52404,", Left(Trim(adoacc190.Fields("a1803").Value), 6)) = 0 Then
'             'modify by sonia 2013/11/15 改抓a1706
'             'StrSqlB = "Select a1504 From acc190, acc150 Where a1902=a1501 And a1901='" & adoacc190.Fields("a1901").Value & "' Group By a1504 Order By 1 "
'             StrSqlB = "Select a1706 From acc190, acc170 Where a1902=a1702 And a1901='" & adoacc190.Fields("a1901").Value & "' Group By a1706 Order By 1 "
'          'Added by Lydia 2018/06/11 補查詢,避免抓不到值
'          Else
'             StrSqlB = "select null from dual"
'          'end 2018/06/11
'          End If
'          '2012/5/4 END
'          rsB.CursorLocation = adUseClient
'          rsB.Open StrSqlB, adoTaie, adOpenStatic, adLockReadOnly
'          ii = 0
'            'Added by Lydia 2015/03/30 +媒體備註
'            If Len(Trim(strA2222)) > 0 Then
'               'Modified by Lydia 2017/07/20 DB note 加NoteTitle
'               'Modified by Lydia 2017/08/14 暫收款退費不加INV. (a1902=a1702為O單號)
'               'strDBNote = Trim(strA2222) & " " & NoteTitle
'               strDBNote = Trim(strA2222)
'               If Val("" & adoacc190.Fields("a1901").Value) = 0 Then
'                  strDBNote = strDBNote & " " & NoteTitle
'               End If
'               'end 2017/08/14
'            Else
'            'end 2015/03/30
'               'Modified by Lydia 2017/07/20 DB note 加NoteTitle
'               'Modified by Lydia 2017/08/14 暫收款退費不加INV. (a1902=a1702為O單號)
'               'strDBNote = "" & NoteTitle
'               If Val("" & adoacc190.Fields("a1901").Value) = 0 Then
'                  strDBNote = "" & NoteTitle
'               End If
'            End If
'
'            'Added by Lydia 2018/04/23  寰華Y53374備註欄都不要再填寫(by 郭)
'            'Modified by Lydia 2020/04/22 +天津三元Y37580; +建毅Y51566,唯源Y52404
'            If InStr("Y37580,Y53374,Y51566,Y52404", Left(Trim(adoacc190.Fields("a1803").Value), 6)) > 0 Then
'                 strDBNote = ""
'            End If
'            'end 2018/04/23
'
'          Do While Not rsB.EOF
'              ii = ii + 1
'              If ii > 12 Then Exit Do
'              If ii Mod 2 = 1 Then
'                  'Added by Lydia 2015/03/30 +媒體備註
'                 ' strDBNote = strDBNote & rsB.Fields(0).Value
'                  'Modified by Lydia 2017/07/20 DB note 加NoteTitle
'                  'strDBNote = IIf(Len(strDBNote) > 0, strDBNote & ",", "") & rsB.Fields(0).Value
'                  strDBNote = strDBNote & IIf(Len(strDBNote) > 0 And strDBNote <> NoteTitle, ",", "") & rsB.Fields(0).Value
'                  'Modified by Lydia 2016/01/20 備註輸入單引號
'                  'adoTaie.Execute "Insert Into ACCRPT427 Values('" & strUserNum & "','" & adoacc190.Fields("a0k11").Value & "','" & adoacc190.Fields("a1803").Value & "','" & rsB.Fields(0).Value & "' )"
'                  adoTaie.Execute "Insert Into ACCRPT427 Values('" & strUserNum & "','" & adoacc190.Fields("a0k11").Value & "','" & adoacc190.Fields("a1803").Value & "','" & ChgSQL("" & rsB.Fields(0).Value) & "' )"
'              Else
'                  strDBNote = strDBNote & "," & rsB.Fields(0).Value
'                  'Modified by Lydia 2016/01/20 備註輸入單引號
'                  'adoTaie.Execute "Insert Into ACCRPT427 Values('" & strUserNum & "','" & adoacc190.Fields("a0k11").Value & "','" & adoacc190.Fields("a1803").Value & "','" & rsB.Fields(0).Value & "' )"
'                  adoTaie.Execute "Insert Into ACCRPT427 Values('" & strUserNum & "','" & adoacc190.Fields("a0k11").Value & "','" & adoacc190.Fields("a1803").Value & "','" & ChgSQL("" & rsB.Fields(0).Value) & "' )"
'                  strDBNote = ""
'              End If
'              rsB.MoveNext
'          Loop
''            If ii <= 10 And ii Mod 2 = 1 Then
'          If ii <= 12 And ii Mod 2 = 1 Then
'              'Modified by Lydia 2016/01/20 備註輸入單引號
'              'adoTaie.Execute "Insert Into ACCRPT427 Values('" & strUserNum & "','" & adoacc190.Fields("a0k11").Value & "','" & adoacc190.Fields("a1803").Value & "','" & strDBNote & "' )"
'              adoTaie.Execute "Insert Into ACCRPT427 Values('" & strUserNum & "','" & adoacc190.Fields("a0k11").Value & "','" & adoacc190.Fields("a1803").Value & "','" & ChgSQL(strDBNote) & "' )"
'              strDBNote = ""
'          End If
'          If rsB.State <> adStateClosed Then rsB.Close
'          Set rsB = Nothing
'      End If
      Call GetAccData(adoacc190, strA0K11)
      'end 2023/04/13
      adoacc190.MoveNext
   Loop
   adoacc190.Close
   '列印合計
   PrintDataNewSumNewPaper strCompany
End Sub

'列印合計
Private Sub PrintDataNewSumNewPaper(strA0K11 As String)
Dim strNo As String
Dim intLength As Integer
'Mark by Lydia 2023/04/13
'Dim strAmount As String
'Dim strCompanyName As String
'Dim strCompanyNo As String
'Dim strAddress As String
'Dim strPhone As String
'Dim strCompAddr1 As String, strCompAddr2 As String  '2017/09/06 公司別-英文地址
'end 2023/04/13
Dim StrSQLa As String
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim strDBNote As String
Dim ii As Integer
'Dim m_NA02 As String 'Added by Lydia 2017/07/19 受款地區 'Mark by Lydia 2023/04/13 改成m_AccNA02

    strNo = ""
   'Modify by Morgan 2008/4/25
'   If Printer.PaperSize <> 291 Then
'     Printer.PaperSize = 291
'   End If
   'Modified by Lydia 2023/04/13 +台銀水單改成Word套印
   'If Text6 = "J" Then
   If Text6 = "J" Or (txtWordTB = "Y" And Text6 = "") Then
      Printer.PaperSize = 9
   Else
      Printer.PaperSize = PUB_GetPaperSize(10)
   End If
   'end 0208/4/25
    'Printer.Font.Size = 12
    'edit by nick 2004/07/21 修改字體
    '2011/6/28 MODIFY BY SONIA 台灣銀行要求改字體,全部改用Printer.Font.Bold = True
    'Printer.Font.Name = "細明體"
    '2011/8/12 modify by sonia 金額及帳號欄改為Times New Roman,12號字,其他仍為Arial,10號字
    'Printer.Font.Name = "Arial"
    '2012/2/16 MODIFY BY SONIA 婧瑄說改全部都是Times New Roman
     Printer.Font.Name = "Times New Roman"
     Printer.Font.Size = 12
    'Printer.Font.Bold = True
    adoacc190_1.CursorLocation = adUseClient

'edit by nickc 2005/05/17 依照定稿語文
'edit by nickc 2005/06/06 改回原設定
   '2006/3/7 MODIFY BY SONIA
   'StrSQLa = "select a1803, fa05, fa63, fa64, fa65, fa10, na03, a1903, a1811, sum(a1904) as Amount, count(distinct a1801) as TCounter, a1917 as a0k11, A1810 from acc190, acc180, fagent, nation where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01(+) and a1908 is null and a1810 is null And a1917='" & stra0k11 & "' " & m_strSQLA & " group by a1803, fa05, fa63, fa64, fa65, fa10, na03, a1903, a1811, a1917, A1810 having count(distinct a1801) > 1 union " & _
   '               "select a1803, fa05, fa63, fa64, fa65, fa10, na03, a1903, a1811, sum(a1904) as Amount, count(distinct a1801) as TCounter, a1917 as a0k11, A1810 from acc190, acc180, fagent, nation where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01(+) and a1908 is null and a1810 is not null And a1917='" & stra0k11 & "' " & m_strSQLA & " group by a1803, fa05, fa63, fa64, fa65, fa10, na03, a1903, a1811, a1917, A1810 having count(distinct a1801) > 1 " & _
   '               " Order By a0k11 asc, a1803 asc "
'edit by nickc 2006/10/27 抓代理人部份改跟明細相同
'   StrSQLa = "select a1803, fa05, fa63, fa64, fa65, fa10, na03, a1903, a1811, sum(a1904) as Amount, count(distinct a1801) as TCounter, a1917 as a0k11, A1810 from acc190, acc180, fagent, nation where a1901 = a1801 and A1803>'Y' AND substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01(+) and a1908 is null and a1810 is null And a1917='" & stra0k11 & "' " & m_strSQLA & " group by a1803, fa05, fa63, fa64, fa65, fa10, na03, a1903, a1811, a1917, A1810 having count(distinct a1801) > 1 union " & _
             "select a1803, fa05, fa63, fa64, fa65, fa10, na03, a1903, a1811, sum(a1904) as Amount, count(distinct a1801) as TCounter, a1917 as a0k11, A1810 from acc190, acc180, fagent, nation where a1901 = a1801 and A1803>'Y' AND substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01(+) and a1908 is null and a1810 is not null And a1917='" & stra0k11 & "' " & m_strSQLA & " group by a1803, fa05, fa63, fa64, fa65, fa10, na03, a1903, a1811, a1917, A1810 having count(distinct a1801) > 1 union " & _
             "select a1803, CU05, CU88, CU89, CU90, CU10, na03, a1903, a1811, sum(a1904) as Amount, count(distinct a1801) as TCounter, a1917 as a0k11, A1810 from acc190, acc180, CUSTOMER, nation where a1901 = a1801 and A1803<'Y' AND substr(a1803, 1, 8) = CU01 (+) and substr(a1803, 9, 1) = CU02 (+) and CU10 = na01(+) and a1908 is null and a1810 is null And a1917='" & stra0k11 & "' " & m_strSQLA & " group by a1803, CU05, CU88, CU89, CU90, CU10, na03, a1903, a1811, a1917, A1810 having count(distinct a1801) > 1 union " & _
             "select a1803, CU05, CU88, CU89, CU90, CU10, na03, a1903, a1811, sum(a1904) as Amount, count(distinct a1801) as TCounter, a1917 as a0k11, A1810 from acc190, acc180, CUSTOMER, nation where a1901 = a1801 and A1803<'Y' AND substr(a1803, 1, 8) = CU01 (+) and substr(a1803, 9, 1) = CU02 (+) and CU10 = na01(+) and a1908 is null and a1810 is not null And a1917='" & stra0k11 & "' " & m_strSQLA & " group by a1803, CU05, CU88, CU89, CU90, CU10, na03, a1903, a1811, a1917, A1810 having count(distinct a1801) > 1 " & _
             " Order By a0k11 asc, a1803 asc "
   'modify by sonia 2014/7/28 合計不含獨立水單以台幣結匯的資料,故加a1812 is null
   'Modified by Lydia 2015/03/10 + Addr
   'Modified by Lydia 2015/10/06 +A1718 代為結匯之客戶編號(申請人)
'   StrSQLa = "select a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05) As FA05, Decode(FA05, Null, Null, fa63) As FA63, Decode(FA05, Null, Null, fa64) As FA64, Decode(FA05, Null, Null, fa65) As FA65, fa10, na03, a1903, a1811, sum(a1904) as Amount, count(distinct a1801) as TCounter, a1917 as a0k11, A1810,rtrim(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70) Addr from acc190, acc180, fagent, nation where a1901 = a1801 and a1812 is null and A1803>'Y' AND substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01(+) and a1908 is null and a1810 is null And a1917='" & stra0k11 & "' " & m_strSQLA & " group by a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05), Decode(FA05, Null, Null, fa63), Decode(FA05, Null, Null, fa64), Decode(FA05, Null, Null, fa65), fa10, na03, a1903, a1811, a1917, A1810,rtrim(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70) having count(distinct a1801) > 1 union " & _
'             "select a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05) As FA05, Decode(FA05, Null, Null, fa63) As FA63, Decode(FA05, Null, Null, fa64) As FA64, Decode(FA05, Null, Null, fa65) As FA65, fa10, na03, a1903, a1811, sum(a1904) as Amount, count(distinct a1801) as TCounter, a1917 as a0k11, A1810,rtrim(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70) Addr from acc190, acc180, fagent, nation where a1901 = a1801 and a1812 is null and A1803>'Y' AND substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01(+) and a1908 is null and a1810 is not null And a1917='" & stra0k11 & "' " & m_strSQLA & " group by a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05), Decode(FA05, Null, Null, fa63), Decode(FA05, Null, Null, fa64), Decode(FA05, Null, Null, fa65), fa10, na03, a1903, a1811, a1917, A1810,rtrim(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70) having count(distinct a1801) > 1 union " & _
'             "select a1803, Decode(CU05, Null, Nvl(CU04, CU06), CU05) As FA05, Decode(CU05, Null, Null, CU88) As FA63, Decode(CU05, Null, Null, CU89) As FA64, Decode(CU05, Null, Null, CU90) As FA65, CU10, na03, a1903, a1811, sum(a1904) as Amount, count(distinct a1801) as TCounter, a1917 as a0k11, A1810,rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)) Addr from acc190, acc180, CUSTOMER, nation where a1901 = a1801 and a1812 is null and A1803<'Y' AND substr(a1803, 1, 8) = CU01 (+) and substr(a1803, 9, 1) = CU02 (+) and CU10 = na01(+) and a1908 is null and a1810 is null And a1917='" & stra0k11 & "' " & m_strSQLA & " group by a1803, Decode(CU05, Null, Nvl(CU04, CU06), CU05), Decode(CU05, Null, Null, CU88), Decode(CU05, Null, Null, CU89), Decode(CU05, Null, Null, CU90), CU10, na03, a1903, a1811, a1917, A1810,rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)) having count(distinct a1801) > 1 union " & _
'             "select a1803, Decode(CU05, Null, Nvl(CU04, CU06), CU05) As FA05, Decode(CU05, Null, Null, CU88) As FA63, Decode(CU05, Null, Null, CU89) As FA64, Decode(CU05, Null, Null, CU90) As FA65, CU10, na03, a1903, a1811, sum(a1904) as Amount, count(distinct a1801) as TCounter, a1917 as a0k11, A1810,rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)) Addr from acc190, acc180, CUSTOMER, nation where a1901 = a1801 and a1812 is null and A1803<'Y' AND substr(a1803, 1, 8) = CU01 (+) and substr(a1803, 9, 1) = CU02 (+) and CU10 = na01(+) and a1908 is null and a1810 is not null And a1917='" & stra0k11 & "' " & m_strSQLA & " group by a1803, Decode(CU05, Null, Nvl(CU04, CU06), CU05), Decode(CU05, Null, Null, CU88), Decode(CU05, Null, Null, CU89), Decode(CU05, Null, Null, CU90), CU10, na03, a1903, a1811, a1917, A1810,rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)) having count(distinct a1801) > 1 " & _
'             " Order By a0k11 asc, a1803 asc "
   'Modified by Lydia 2017/03/23 加抓受款行國別
   'StrSQLa = "select a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05) As FA05, Decode(FA05, Null, Null, fa63) As FA63, Decode(FA05, Null, Null, fa64) As FA64, Decode(FA05, Null, Null, fa65) As FA65, fa10, na03, a1903, a1811, sum(a1904) as Amount, count(distinct a1801) as TCounter, a1917 as a0k11, A1810,rtrim(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70) Addr,a1718 from acc190, acc180, fagent, nation, acc170 where a1901 = a1801 and a1812 is null and A1803>'Y' AND substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01(+) and a1908 is null and a1810 is null And a1902=a1702(+) And a1917='" & stra0k11 & "' " & m_strSQLA & " group by a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05), Decode(FA05, Null, Null, fa63), Decode(FA05, Null, Null, fa64), Decode(FA05, Null, Null, fa65), fa10, na03, a1903, a1811, a1917, A1810,rtrim(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70),a1718 having count(distinct a1801) > 1 union " & _
             "select a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05) As FA05, Decode(FA05, Null, Null, fa63) As FA63, Decode(FA05, Null, Null, fa64) As FA64, Decode(FA05, Null, Null, fa65) As FA65, fa10, na03, a1903, a1811, sum(a1904) as Amount, count(distinct a1801) as TCounter, a1917 as a0k11, A1810,rtrim(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70) Addr,a1718 from acc190, acc180, fagent, nation, acc170 where a1901 = a1801 and a1812 is null and A1803>'Y' AND substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01(+) and a1908 is null and a1810 is not null And a1902=a1702(+) And a1917='" & stra0k11 & "' " & m_strSQLA & " group by a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05), Decode(FA05, Null, Null, fa63), Decode(FA05, Null, Null, fa64), Decode(FA05, Null, Null, fa65), fa10, na03, a1903, a1811, a1917, A1810,rtrim(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70),a1718 having count(distinct a1801) > 1 union " & _
             "select a1803, Decode(CU05, Null, Nvl(CU04, CU06), CU05) As FA05, Decode(CU05, Null, Null, CU88) As FA63, Decode(CU05, Null, Null, CU89) As FA64, Decode(CU05, Null, Null, CU90) As FA65, CU10, na03, a1903, a1811, sum(a1904) as Amount, count(distinct a1801) as TCounter, a1917 as a0k11, A1810,rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)) Addr,a1718 from acc190, acc180, CUSTOMER, nation, acc170 where a1901 = a1801 and a1812 is null and A1803<'Y' AND substr(a1803, 1, 8) = CU01 (+) and substr(a1803, 9, 1) = CU02 (+) and CU10 = na01(+) and a1908 is null and a1810 is null And a1902=a1702(+) And a1917='" & stra0k11 & "' " & m_strSQLA & " group by a1803, Decode(CU05, Null, Nvl(CU04, CU06), CU05), Decode(CU05, Null, Null, CU88), Decode(CU05, Null, Null, CU89), Decode(CU05, Null, Null, CU90), CU10, na03, a1903, a1811, a1917, A1810,rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)),a1718 having count(distinct a1801) > 1 union " & _
             "select a1803, Decode(CU05, Null, Nvl(CU04, CU06), CU05) As FA05, Decode(CU05, Null, Null, CU88) As FA63, Decode(CU05, Null, Null, CU89) As FA64, Decode(CU05, Null, Null, CU90) As FA65, CU10, na03, a1903, a1811, sum(a1904) as Amount, count(distinct a1801) as TCounter, a1917 as a0k11, A1810,rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)) Addr,a1718 from acc190, acc180, CUSTOMER, nation, acc170 where a1901 = a1801 and a1812 is null and A1803<'Y' AND substr(a1803, 1, 8) = CU01 (+) and substr(a1803, 9, 1) = CU02 (+) and CU10 = na01(+) and a1908 is null and a1810 is not null And a1902=a1702(+) And a1917='" & stra0k11 & "' " & m_strSQLA & " group by a1803, Decode(CU05, Null, Nvl(CU04, CU06), CU05), Decode(CU05, Null, Null, CU88), Decode(CU05, Null, Null, CU89), Decode(CU05, Null, Null, CU90), CU10, na03, a1903, a1811, a1917, A1810,rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)),a1718 having count(distinct a1801) > 1 " & _
             " Order By a0k11 asc, a1803 asc "
    'Modified by Lydia 2017/07/19 客戶國家na01 地區na02
    'Modified by Lydia 2017/08/14 + 判斷是否為暫收款退費 -> sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt
    'StrSQLa = "select a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05) As FA05, Decode(FA05, Null, Null, fa63) As FA63, Decode(FA05, Null, Null, fa64) As FA64, Decode(FA05, Null, Null, fa65) As FA65, fa10, na03 , na02 , a1903, a1811, sum(a1904) as Amount, count(distinct a1801) as TCounter, a1917 as a0k11, A1810,rtrim(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70) Addr,a1718 from acc190, acc180, fagent, nation, acc170 where a1901 = a1801 and a1812 is null and A1803>'Y' AND substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01(+) and a1908 is null and a1810 is null And a1902=a1702(+) And a1917='" & stra0k11 & "' " & m_strSQLA & " group by a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05), Decode(FA05, Null, Null, fa63), Decode(FA05, Null, Null, fa64), Decode(FA05, Null, Null, fa65), fa10, na03 , na02 , a1903, a1811, a1917, A1810,rtrim(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70),a1718 having count(distinct a1801) > 1 union " & _
             "select a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05) As FA05, Decode(FA05, Null, Null, fa63) As FA63, Decode(FA05, Null, Null, fa64) As FA64, Decode(FA05, Null, Null, fa65) As FA65, fa10, na03 , na02 , a1903, a1811, sum(a1904) as Amount, count(distinct a1801) as TCounter, a1917 as a0k11, A1810,rtrim(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70) Addr,a1718 from acc190, acc180, fagent, nation, acc170 where a1901 = a1801 and a1812 is null and A1803>'Y' AND substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01(+) and a1908 is null and a1810 is not null And a1902=a1702(+) And a1917='" & stra0k11 & "' " & m_strSQLA & " group by a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05), Decode(FA05, Null, Null, fa63), Decode(FA05, Null, Null, fa64), Decode(FA05, Null, Null, fa65), fa10, na03 , na02 , a1903, a1811, a1917, A1810,rtrim(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70),a1718 having count(distinct a1801) > 1 union " & _
             "select a1803, Decode(CU05, Null, Nvl(CU04, CU06), CU05) As FA05, Decode(CU05, Null, Null, CU88) As FA63, Decode(CU05, Null, Null, CU89) As FA64, Decode(CU05, Null, Null, CU90) As FA65, CU10, na03 , na02 , a1903, a1811, sum(a1904) as Amount, count(distinct a1801) as TCounter, a1917 as a0k11, A1810,rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)) Addr,a1718 from acc190, acc180, CUSTOMER, nation, acc170 where a1901 = a1801 and a1812 is null and A1803<'Y' AND substr(a1803, 1, 8) = CU01 (+) and substr(a1803, 9, 1) = CU02 (+) and CU10 = na01(+) and a1908 is null and a1810 is null And a1902=a1702(+) And a1917='" & stra0k11 & "' " & m_strSQLA & " group by a1803, Decode(CU05, Null, Nvl(CU04, CU06), CU05), Decode(CU05, Null, Null, CU88), Decode(CU05, Null, Null, CU89), Decode(CU05, Null, Null, CU90), CU10, na03 , na02 , a1903, a1811, a1917, A1810,rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)),a1718 having count(distinct a1801) > 1 union " & _
             "select a1803, Decode(CU05, Null, Nvl(CU04, CU06), CU05) As FA05, Decode(CU05, Null, Null, CU88) As FA63, Decode(CU05, Null, Null, CU89) As FA64, Decode(CU05, Null, Null, CU90) As FA65, CU10, na03 , na02 , a1903, a1811, sum(a1904) as Amount, count(distinct a1801) as TCounter, a1917 as a0k11, A1810,rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)) Addr,a1718 from acc190, acc180, CUSTOMER, nation, acc170 where a1901 = a1801 and a1812 is null and A1803<'Y' AND substr(a1803, 1, 8) = CU01 (+) and substr(a1803, 9, 1) = CU02 (+) and CU10 = na01(+) and a1908 is null and a1810 is not null And a1902=a1702(+) And a1917='" & stra0k11 & "' " & m_strSQLA & " group by a1803, Decode(CU05, Null, Nvl(CU04, CU06), CU05), Decode(CU05, Null, Null, CU88), Decode(CU05, Null, Null, CU89), Decode(CU05, Null, Null, CU90), CU10, na03 , na02 , a1903, a1811, a1917, A1810,rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)),a1718 having count(distinct a1801) > 1 "
    'Modified by Lydia 2023/04/13 為了共用模組+'' as A1812 ----參考modify by sonia 2014/7/10 +a1812,獨立水單以台幣結匯且合計水單不算此筆
    StrSQLa = "select a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05) As FA05, Decode(FA05, Null, Null, fa63) As FA63, Decode(FA05, Null, Null, fa64) As FA64, Decode(FA05, Null, Null, fa65) As FA65, fa10, na03 , na02, na01 , a1903, a1811, sum(a1904) as Amount, count(distinct a1801) as TCounter, a1917 as a0k11, A1810," & _
              " rtrim(max(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70)) Addr,a1718,sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt,'' as a1812 " & _
              " from acc190, acc180, fagent, nation, acc170 where a1901 = a1801 and a1812 is null and A1803>'Y' AND substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01(+) and a1908 is null and a1810 is null And a1902=a1702(+) And a1917='" & strA0K11 & "' " & m_strSQLA & _
              " group by a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05), Decode(FA05, Null, Null, fa63), Decode(FA05, Null, Null, fa64), Decode(FA05, Null, Null, fa65), fa10, na03 , na02, na01 , a1903, a1811, a1917, A1810,a1718 having count(distinct a1801) > 1"
    StrSQLa = StrSQLa & " union select a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05) As FA05, Decode(FA05, Null, Null, fa63) As FA63, Decode(FA05, Null, Null, fa64) As FA64, Decode(FA05, Null, Null, fa65) As FA65, fa10, na03 , na02, na01 , a1903, a1811, sum(a1904) as Amount, count(distinct a1801) as TCounter, a1917 as a0k11, A1810," & _
              " rtrim(max(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70)) Addr,a1718,sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt,'' as a1812 " & _
              " from acc190, acc180, fagent, nation, acc170 where a1901 = a1801 and a1812 is null and A1803>'Y' AND substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01(+) and a1908 is null and a1810 is not null And a1902=a1702(+) And a1917='" & strA0K11 & "' " & m_strSQLA & _
              " group by a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05), Decode(FA05, Null, Null, fa63), Decode(FA05, Null, Null, fa64), Decode(FA05, Null, Null, fa65), fa10, na03 , na02, na01 , a1903, a1811, a1917, A1810,a1718 having count(distinct a1801) > 1 "
    StrSQLa = StrSQLa & " union select a1803, Decode(CU05, Null, Nvl(CU04, CU06), CU05) As FA05, Decode(CU05, Null, Null, CU88) As FA63, Decode(CU05, Null, Null, CU89) As FA64, Decode(CU05, Null, Null, CU90) As FA65, CU10, na03 , na02, na01 , a1903, a1811, sum(a1904) as Amount, count(distinct a1801) as TCounter, a1917 as a0k11, A1810," & _
              " rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)) Addr,a1718,sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt,'' as a1812 " & _
              " from acc190, acc180, CUSTOMER, nation, acc170 where a1901 = a1801 and a1812 is null and A1803<'Y' AND substr(a1803, 1, 8) = CU01 (+) and substr(a1803, 9, 1) = CU02 (+) and CU10 = na01(+) and a1908 is null and a1810 is null And a1902=a1702(+) And a1917='" & strA0K11 & "' " & m_strSQLA & _
              " group by a1803, Decode(CU05, Null, Nvl(CU04, CU06), CU05), Decode(CU05, Null, Null, CU88), Decode(CU05, Null, Null, CU89), Decode(CU05, Null, Null, CU90), CU10, na03 , na02, na01 , a1903, a1811, a1917, A1810,a1718 having count(distinct a1801) > 1 "
    StrSQLa = StrSQLa & " union select a1803, Decode(CU05, Null, Nvl(CU04, CU06), CU05) As FA05, Decode(CU05, Null, Null, CU88) As FA63, Decode(CU05, Null, Null, CU89) As FA64, Decode(CU05, Null, Null, CU90) As FA65, CU10, na03 , na02, na01 , a1903, a1811, sum(a1904) as Amount, count(distinct a1801) as TCounter, a1917 as a0k11, A1810," & _
              " rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)) Addr,a1718,sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt,'' as a1812 " & _
              " from acc190, acc180, CUSTOMER, nation, acc170 where a1901 = a1801 and a1812 is null and A1803<'Y' AND substr(a1803, 1, 8) = CU01 (+) and substr(a1803, 9, 1) = CU02 (+) and CU10 = na01(+) and a1908 is null and a1810 is not null And a1902=a1702(+) And a1917='" & strA0K11 & "' " & m_strSQLA & _
              " group by a1803, Decode(CU05, Null, Nvl(CU04, CU06), CU05), Decode(CU05, Null, Null, CU88), Decode(CU05, Null, Null, CU89), Decode(CU05, Null, Null, CU90), CU10, na03 , na02, na01 , a1903, a1811, a1917, A1810,a1718 having count(distinct a1801) > 1 "
    'end 2017/08/14
    
    'Modified by Lydia 2017/07/19 受款行國家地區+Accna02
    'Modified by Lydia 2017/10/03 華銀整批媒體RMB改CNY
    'StrSQLa = "select a.*,n2.na03 as Accna03,n2.na02 as Accna02 from (" & StrSQLa & ") a, acc220,nation n2 where a1803=a2201(+) and a1903=a2202(+) and a2217=n2.na01(+) "
    StrSQLa = "select a.*,n2.na03 as Accna03,n2.na02 as Accna02 from (" & StrSQLa & ") a, acc220,nation n2 where a1803=a2201(+) and decode(a1903,'" & J_RMB & "','RMB',a1903)=a2202(+) and a2217=n2.na01(+) "
               
    StrSQLa = StrSQLa & "Order By a0k11 asc, a1803 asc "
    'end 2017/03/23
    adoacc190_1.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
    'edit by nickc 2005/11/21
    'If adoacc190_1.RecordCount <= 0 Then adoacc190_1.Close: Printer.EndDoc: Exit Sub
    If adoacc190_1.RecordCount <= 0 Then
      adoacc190_1.Close
      'Modified by Lydia 2023/04/13 +台銀水單改成Word套印
      'If IsEnd = True Then
      If IsEnd = True And txtWordTB <> "Y" Then
         Printer.EndDoc
      Else
         'Modified by Lydia 2015/03/18 付款日104/03/12 非J公司只印票匯,2張票匯(不同公司別),中間有空白頁
        ' Printer.NewPage
      End If
      Exit Sub
    End If
    
    strNo = "" & adoacc190_1.Fields("a1803").Value & adoacc190_1.Fields("a1903").Value & adoacc190_1.Fields("fa05").Value
    Do While adoacc190_1.EOF = False
        If strNo <> (adoacc190_1.Fields("a1803").Value & adoacc190_1.Fields("a1903").Value & adoacc190_1.Fields("fa05").Value) Then
            'Modified by Lydia 2023/04/13 +台銀水單改成Word套印
            'If strNo <> "" Then
            If strNo <> "" And txtWordTB <> "Y" Then
                Printer.NewPage
            End If
            strNo = (adoacc190_1.Fields("a1803").Value & adoacc190_1.Fields("a1903").Value & adoacc190_1.Fields("fa05").Value)
        End If
        
        Call GetAccData(adoacc190_1, strA0K11) 'Added by Lydia 2023/04/13 改成模組
        'Added by Lydia 2023/04/13 台銀水單改成Word套印
        If txtWordTB = "Y" Then
            Call TBCallWordPrint(adoacc190_1, "2")
        Else
        'end 2023/04/13
           '2013/3/6 add by sonia 婧瑄說有時會不清楚,找不出原因故再寫一次
           Printer.Font.Name = "Times New Roman"
           Printer.Font.Size = 12
           '2013/3/6 end
           '代理人
           If IsNull(adoacc190_1.Fields("a1803").Value) = False Then
               Printer.CurrentX = 7850 - intX
               'Modified by Lydia 2015/04/29 台銀新版(104.04)
               'Printer.CurrentY = 750 - intY
               Printer.CurrentY = 1255 - intY
               Printer.Print adoacc190_1.Fields("a1803").Value & " " & "合計"
           Else
               Printer.CurrentX = 7850 - intX
               'Modified by Lydia 2015/04/29 台銀新版(104.04)
               'Printer.CurrentY = 750 - intY
               Printer.CurrentY = 1255 - intY
               Printer.Print "合計"
           End If
           '受款地址國別
           If IsNull(adoacc190_1.Fields("na03").Value) = False Then
               'Printer.CurrentX = 300 - intX + 2666   '2011/6/28改位置
               'Modified by Lydia 2015/04/29 台銀新版(104.04)
               'Printer.CurrentX = 1734 - intX
               'Printer.CurrentY = 1965 - intY
               'Modified by Lydia 2018/11/09 台銀新版(107.08)
               'Printer.CurrentX = 2504 - intX
               'Printer.CurrentY = 2135 - intY
               Printer.CurrentX = 2560 - intX
               Printer.CurrentY = 2240 - intY
               'Modify by Morgan 2009/8/4
               'If adoacc190_1.Fields("na03").Value = "大陸" Then
               '2010/3/9 CANCEL BY SONIA
               'If InStr(adoacc190_1.Fields("na03").Value, "大陸") > 0 Then
               '   Printer.Print "香港"
               'Else
                  'Modified by Lydia 2017/03/23 改抓受款行國別
                  'Modifiec by Lydia 2023/04/13
                  'If Trim("" & adoacc190_1.Fields("accna03").Value) <> "" Then
                  '   Printer.Print adoacc190_1.Fields("accna03").Value
                  'Else
                  '   Printer.Print adoacc190_1.Fields("na03").Value
                  'End If
                  ''end 2017/03/23
                  Printer.Print m_NationName
                  
                  'Modified by Lydia 2017/07/19 抓國家地區
                  'Mark by Lydia 2023/04/13 改用模組取得
                  'If Trim("" & adoacc190_1.Fields("accna02").Value) <> "" Then
                  '   m_NA02 = "" & adoacc190_1.Fields("accna02").Value
                  'Else
                  '   m_NA02 = "" & adoacc190_1.Fields("na02").Value
                  'End If
                  ''end 2017/07/19
                  'end 2023/04/13
               'End If
               '2010/3/9 END
           End If
           '國外受款人身分別
           'Modified by Lydia 2015/04/29 台銀新版(104.04)
           'Printer.CurrentX = 7300 - intX
           'Printer.CurrentY = 1865 - intY
           'Modified by Lydia 2018/11/09 台銀新版(107.08) ; 移到申請人地址同一行
           'Printer.CurrentX = 7920 - intX
           'Printer.CurrentY = 2180 - intY
           Printer.CurrentX = 10930 - intX
           Printer.CurrentY = 4160 - intY
           Printer.Print "V"
           '幣別
           '匯款方式
           'Modified by Lydia 2015/04/29 台銀新版(104.04)
           'Printer.CurrentY = 1865 - intY
           'Modified by Lydia 2018/11/09 台銀新版(107.08) ; 移到受款地區國別同一行
           'Printer.CurrentY = 2615 - intY
           Printer.CurrentY = 2240 - intY
           'Added by Lydia 2015/04/17 +匯款方式3:台銀電匯紙本
           If adoacc190_1.Fields("a1811").Value = "2" Or adoacc190_1.Fields("a1811").Value = "3" Then
               '電匯
               'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
               'Printer.CurrentX = 9450 - intX
               'Modified by Lydia 2015/04/29 台銀新版(104.04)
               'Printer.CurrentX = 9310 - intX
               'Modified by Lydia 2018/11/09 台銀新版(107.08)
               'Printer.CurrentX = 8440 - intX
               Printer.CurrentX = 8880 - intX
           Else
               '票匯
               'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
               'Printer.CurrentX = 10625 - intX
               'Modified by Lydia 2015/04/29 台銀新版(104.04)
               'Printer.CurrentX = 10870 - intX
               'Modified by Lydia 2018/11/09 台銀新版(107.08)
               'Printer.CurrentX = 10430 - intX
               'Modified by Lydia 2018/12/28 調左
               'Printer.CurrentX = 10900 - intX
               Printer.CurrentX = 10840 - intX
           End If
           Printer.Print "V"
           
           'Mark by Lydia 2023/04/13 改成模組GetAccData
'           '抓公司別基本資料
'           strA0K11 = "" & adoacc190_1.Fields("a0k11").Value
'           adoquery.CursorLocation = adUseClient
'           'Modified by Lydia 2015/10/06 有代為結匯之客戶編號,列印的申請人資料為該客戶的資料
'           'adoquery.Open "select * from acc080 where a0801 = '" & stra0k11 & "' ", adoTaie, adOpenStatic, adLockReadOnly
'           If "" & adoacc190_1.Fields("a1718").Value <> "" Then
'              ' 以\分2行列印
'              'Modified by Lydia 2015/12/04 有短地址就用短地址
'              'Modified by Lydia 2017/10/03 華銀整批媒體RMB改CNY
'              'strSql = "select CU05||' '||CU88||' '||CU89||' '||CU90 as a0803,CU11 as a0807,NVL(CU16,CU17) as a0813,nvl(a2218,CU24||' '||CU25||'\'||CU26||' '||CU27||' '||CU28) as custaddr " & _
'                      "from customer,acc220 where cu01||cu02=" & CNULL(adoacc190_1.Fields("a1718") & "0") & " and cu01||cu02=a2201(+) "
'              strSql = "select CU05||' '||CU88||' '||CU89||' '||CU90 as a0803,CU11 as a0807,NVL(CU16,CU17) as a0813,nvl(a2218,CU24||' '||CU25||'\'||CU26||' '||CU27||' '||CU28) as custaddr " & _
'                      "from customer,acc220 where cu01||cu02=" & CNULL(adoacc190_1.Fields("a1718") & "0") & " and cu01||cu02=a2201(+) and a2202='" & IIf(adoacc190_1.Fields("a1903") = J_RMB, "RMB", adoacc190_1.Fields("a1903")) & "' "
'           Else
'              'Modified by Lydia 2020/09/03 收據公司別1,2,L都歸2公司
'              'strSql = "select * from acc080 where a0801 = '" & strA0K11 & "' "
'              strSql = "select * from acc080 where a0801 = '" & IIf(strA0K11 = "J", "J", "2") & "' "
'           End If
'           adoquery.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'           'end 2015/10/06
'           If adoquery.RecordCount <> 0 Then
'               strCompanyName = "" & adoquery.Fields("a0803").Value
'
'               If IsNull(adoquery.Fields("a0807").Value) Then
'                   strCompanyNo = ""
'               Else
'                   strCompanyNo = adoquery.Fields("a0807").Value
'               End If
'               'Modified by Lydia 2015/10/06
'               'strAddress = ReportSum(104)
'               If "" & adoacc190_1.Fields("a1718").Value <> "" Then
'                  strAddress = "" & adoquery.Fields("custaddr")
'               Else
'                  'Modified by Lydia 2017/09/14 公司別-中文,英文地址
'                  'strAddress = ReportSum(104)
'                  strAddress = Replace("" & adoquery.Fields("a0804").Value, "朱園里7鄰", "")
'                  strCompAddr1 = "" & adoquery.Fields("a0822").Value
'                  strCompAddr2 = "" & adoquery.Fields("a0823").Value
'               End If
'               'end 2015/10/06
'               If IsNull(adoquery.Fields("a0813").Value) Then
'                   strPhone = ""
'               Else
'                  strPhone = adoquery.Fields("a0813").Value
'               End If
'           Else
'               strCompanyName = ""
'               strCompanyNo = ""
'               strAddress = ""
'               strPhone = ""
'               'Added by Lydia 2017/09/06 公司別-英文地址
'               strCompAddr1 = ""
'               strCompAddr2 = ""
'           End If
'           'Added by Lydia 2017/09/06 公司別-英文地址(預設)
'           If strCompAddr1 & strCompAddr2 = "" Then
'              strCompAddr1 = "9F, No. 112, Sec. 2, Chang-An E. Rd.,"
'              strCompAddr2 = "Taipei 104, Taiwan, R.O.C."
'           End If
'           'end 2017/09/07
'           'Added by Lydia 2021/08/27 寰華結匯水單本所地址不能出現ROC
'           If InStr("Y53374,", Left(Trim("" & adoacc190_1.Fields("a1803")), 6)) > 0 Then
'               strCompAddr2 = Replace(strCompAddr2, ", R.O.C.", "")
'           End If
'           'end 2021/08/27
'           adoquery.Close
           'end 2023/04/13 改成模組GetAccData
           
           '申請人名稱
           '匯款人名稱
           'Printer.Font.Name = "Arial"
           Printer.Font.Size = 10
           Printer.CurrentX = 1934 - intX
           'Modified by Lydia 2015/04/29 台銀新版(104.04)
           'Printer.CurrentY = 2395 - intY
           'Modified by Lydia 2018/11/09 台銀新版(107.08)
           'Printer.CurrentY = 2725 - intY
           Printer.CurrentY = 2820 - intY
           Printer.Print strCompanyName
           
           '外匯去處--匯往國外
           'Remove by Lydia 2018/11/09 台銀新版(107.08)
   '        Printer.Font.Name = "Times New Roman"
   '        Printer.Font.Size = 12
   '        'Modified by Lydia 2015/04/29 台銀新版(104.04)
   ''        Printer.CurrentX = 7490 - intX
   ''        Printer.CurrentY = 2295 - intY
   '        Printer.CurrentX = 7600 - intX
   '        Printer.CurrentY = 3050 - intY
   '        Printer.Print "V"
           'end -- Remove 2018/11/09
           
           '金額
           Printer.Font.Size = 12 'Added by Lydia 2018/11/09 台銀新版(107.08)
           'Modified by Lydia 2015/04/29 台銀新版(104.04)
   '        Printer.CurrentY = 3175 - intY
           'Modified by Lydia 2018/11/09 台銀新版(107.08) ; 移到受款地區國別同一行
           'Printer.CurrentY = 3640 - intY
           'Printer.CurrentX = 10970 - intLength - intX
           Printer.CurrentY = 2240 - intY
           'Mark by Lydia 2023/04/13
'           If IsNull(adoacc190_1.Fields("Amount").Value) Then
'               '金額前不加幣別
'               strAmount = adoacc190_1.Fields("a1903").Value & "0.00"
'   '            strAmount = "0.00"
'               intLength = Printer.TextWidth(m_Amount)
'               'Modified by Lydia 2015/04/29 台銀新版(104.04)
'   '            Printer.CurrentX = 10050 - intLength - intX
'           Else
'               '金額前不加幣別
'               strAmount = adoacc190_1.Fields("a1903").Value & Format(adoacc190_1.Fields("Amount").Value, FDollar)
'               intLength = Printer.TextWidth(m_Amount)
'               'Modified by Lydia 2015/04/29 台銀新版(104.04)
'               'Printer.CurrentX = 10050 - intLength - intX
'           End If
           intLength = Printer.TextWidth(m_Amount)
           'end 2023/04/13
           Printer.CurrentX = 6360 - intLength - intX 'Memo by Lydia 2018/11/12 與單張不同
           'Modified by Lydia 2023/04/13 strAmount=>m_Amount
           Printer.Print m_Amount
           
           '統一編號
           'Modified by Lydia 2015/04/29 台銀新版(104.04)
           'Printer.CurrentX = 3174 - intX
           'Printer.CurrentY = 2720 - intY
           Printer.CurrentX = 3264 - intX
           Printer.CurrentY = 2990 - intY
           Printer.Print strCompanyNo
           '地址
           'Printer.Font.Name = "Arial"
           Printer.Font.Size = 10
           Printer.CurrentX = 1300 - intX
           'Modified by Lydia 2015/04/29 台銀新版(104.04)
           'Printer.CurrentY = 3170 - intY
           'Modified by Lydia 2018/11/09 台銀新版(107.08)
           'Printer.CurrentY = 3440 - intY
           Printer.CurrentY = 4250 - intY
           'Modified by Lydia 2015/10/06 有代為結匯之客戶編號,列印的申請人資料為該客戶的資料
           'Printer.Print "9F1., No. 112, Sec. 2, Chang-An E. Rd.,"
           If "" & adoacc190_1.Fields("a1718").Value <> "" Then
              '以\分2行列印
               'Modified by Lydia 2015/12/04
               'Printer.Print Mid(strAddress, 1, InStr(strAddress, "\") - 1)
               If InStr(strAddress, "\") = 0 Then
                  Printer.Print strAddress
               Else
                  Printer.Print Mid(strAddress, 1, InStr(strAddress, "\") - 1)
               End If
           Else
              'Modified by Lydia 2017/09/07 改成公司別-英文地址1
              'Printer.Print "9F1., No. 112, Sec. 2, Chang-An E. Rd.,"
              Printer.Print strCompAddr1
           End If
           
           Printer.CurrentX = 1300 - intX
           'Modified by Lydia 2015/04/29 台銀新版(104.04)
           'Printer.CurrentY = 3400 - intY
           'Modified by Lydia 2018/11/09 台銀新版(107.08)
           'Printer.CurrentY = 3670 - intY
           Printer.CurrentY = 4520 - intY
           'Modified by Lydia 2015/10/06 有代為結匯之客戶編號,列印的申請人資料為該客戶的資料
           'Printer.Print "      Taipei 104, Taiwan, R.O.C."
           If "" & adoacc190_1.Fields("a1718").Value <> "" Then
              '以\分2行列印
              'Added by Lydia 2015/12/04
              If InStr(strAddress, "\") > 0 Then
                  Printer.Print "      " & Mid(strAddress, InStr(strAddress, "\") + 1)
              End If
           Else
              'Modified by Lydia 2017/09/07 改成公司別-英文地址2
              'Printer.Print "      Taipei 104, Taiwan, R.O.C."
              Printer.Print "      " & strCompAddr2
           End If
           
           '電話
           Printer.Font.Name = "Times New Roman"
           Printer.Font.Size = 12
           'Modified by Lydia 2018/11/09 台銀新版(107.08)
           'Printer.CurrentX = 4800 - intX
           'Modified by Lydia 2018/12/28 靠右
           'Printer.CurrentX = 1800 - intX
           Printer.CurrentX = 2200 - intX
           'Modified by Lydia 2015/04/29 台銀新版(104.04)
           'Printer.CurrentY = 3400 - intY
           Printer.CurrentY = 3640 - intY
            'Added by Lydia 2015/10/06 有代為結匯之客戶編號,列印的申請人資料為該客戶的資料
            If "" & adoacc190_1.Fields("a1718").Value <> "" Then
                'Modified by Lydia 2018/11/09 台銀新版(107.08)
                'Printer.CurrentX = 2200 - intX
                'Printer.CurrentY = 3900 - intY
                'Modified by Lydia 2018/12/28 靠右
                'Printer.CurrentX = 1800 - intX
                Printer.CurrentX = 2200 - intX
                Printer.CurrentY = 4850 - intY
                Printer.Print strPhone
            Else
                'Modified by Lydia 2018/11/09 台銀新版(107.08)
                'Printer.CurrentY = 3640 - intY
                'Printer.Print "Tel:" & strPhone
                Printer.CurrentY = 4850 - intY
                Printer.Print strPhone
            End If
            'end 2015/10/06
           
           '繳款方式
           'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
           'Printer.CurrentX = 7580 - intX
           'Modified by Lydia 2015/04/29 台銀新版(104.04)
           'Printer.CurrentX = 7460 - intX
           'Modified by Lydia 2018/11/09 台銀新版(107.08)
           'Printer.CurrentX = 7760 - intX
           Printer.CurrentX = 7560 - intX
           Select Case adoacc190_1.Fields("a1903").Value
               'Modify by Morgan 2010/6/15 取消歐元
               'Case "USD", "EUR"
               Case "USD"
                   '以外匯存款提出
                   'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                   'Printer.CurrentY = 3867 - intY
                   'Modified by Lydia 2015/04/29 台銀新版(104.04)
                   'Printer.CurrentY = 3940 - intY
                   'Modified by Lydia 2018/11/09 台銀新版(107.08)
                   'Printer.CurrentY = 4460 - intY
                   Printer.CurrentY = 3050 - intY
               Case Else
                   '以新台幣結購
                   'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                   'Printer.CurrentY = 3567 - intY
                   'Modified by Lydia 2015/04/29 台銀新版(104.04)
                   'Printer.CurrentY = 3640 - intY
                   'Modified by Lydia 2018/11/09 台銀新版(107.08)
                   'Printer.CurrentY = 4080 - intY
                   Printer.CurrentY = 2730 - intY
           End Select
           'Modified by Lydia 2015/10/06 有代為結匯之客戶編號,USD以新台幣結購
           If "" & adoacc190_1.Fields("a1718").Value <> "" Then
                   'Modified by Lydia 2018/11/09 台銀新版(107.08)
                   'Printer.CurrentY = 4080 - intY
                   Printer.CurrentY = 2730 - intY
           End If
           
           Printer.Print "V"
           
           '金額
           'Added by Lydia 2015/10/06  有代為結匯之客戶編號,USD以新台幣結購
           If IsNull(adoacc190_1.Fields("a1718").Value) Then
               Select Case adoacc190_1.Fields("a1903").Value
                   'Modify by Morgan 2010/6/15 取消歐元
                   'Case "USD", "EUR"
                   Case "USD"
                   'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                      'Printer.CurrentY = 3827 - intY
                      'Modified by Lydia 2015/04/29 台銀新版(104.04)
                      'Printer.CurrentY = 3900 - intY
                      'Modified by Lydia 2018/11/09 台銀新版(107.08)
                      'Printer.CurrentY = 4570 - intY
                      Printer.CurrentY = 3200 - intY
                      
                      If IsNull(adoacc190_1.Fields("Amount").Value) Then
                           '金額前不加幣別
                         'Modified by Lydia 2023/04/13 strAmount=>m_Amount
                         m_Amount = adoacc190_1.Fields("a1903").Value & "0.00"
                         intLength = Printer.TextWidth(m_Amount)
                         'Modified by Lydia 2015/04/29 台銀新版(104.04)
                         'Printer.CurrentX = 11220 - intLength - intX
                      Else
                           '金額前不加幣別
                         'Modified by Lydia 2023/04/13 strAmount=>m_Amount
                         m_Amount = adoacc190_1.Fields("a1903").Value & Format(adoacc190_1.Fields("Amount").Value, FDollar)
                         intLength = Printer.TextWidth(m_Amount)
                         'Modified by Lydia 2015/04/29 台銀新版(104.04)
                         'Printer.CurrentX = 11220 - intLength - intX
                      End If
                      'Modified by Lydia 2018/11/09 台銀新版(107.08)
                      'Printer.CurrentX = 11520 - intLength - intX
                      Printer.CurrentX = 11300 - intLength - intX
                      'Modified by Lydia 2023/04/13 strAmount=>m_Amount
                      Printer.Print m_Amount
                   Case Else
                     'Remove by Lydia 2018/11/09 台銀新版(107.08) ; 只印美金幣別
                     'Printer.CurrentY = 3907 - intY
               End Select
           End If
           'end 2015/10/06
           
         '列印correspondent bank
         'Printer.Font.Name = "Arial"
         'Added by Lydia 2015/03/30 媒體備註
         'strA2222 = "" 'Mark by Lydia 2023/04/13
         Printer.Font.Size = 10
           'Modified by Lydia 2015/04/29 台銀新版(104.04)
           'Modified by Lydia 2018/11/09 台銀新版(107.08)
           'tmpY = 5435: tmpX = 220  'tmpY暫存起始高度,tmpX暫存每行高度
           tmpY = 6230: tmpX = 220
           
           'Modified by Lydia 2017/10/03 華銀整批媒體RMB改CNY
           'StrSqlB = "Select * From ACC220 Where a2201='" & adoacc190_1.Fields("a1803").Value & "' And a2202='" & adoacc190_1.Fields("a1903").Value & "' "
           'Mark by Lydia 2023/04/13 改成模組GetAccData
           'StrSqlB = "Select * From ACC220 Where a2201='" & adoacc190_1.Fields("a1803").Value & "' And a2202='" & IIf(adoacc190_1.Fields("a1903").Value = J_RMB, "RMB", adoacc190_1.Fields("a1903").Value) & "' "
           'rsB.CursorLocation = adUseClient
           'rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
           'If rsB.RecordCount > 0 Then
           '    'Added by Lydia 2015/03/30 媒體備註
           '    strA2222 = "" & UCase(Trim(rsB.Fields("a2222")))
           '    tmpY = tmpY + IIf("" & rsB.Fields("A2220") <> "", tmpX, 0)  'Modified by Lydia 2017/10/16 中間銀行Y座標
           '    If "" & rsB("A2214").Value <> "" Or rsB("A2215").Value <> "" Or rsB("A2216").Value <> "" Then
           tmpY = tmpY + IIf(m_Payee_04 <> "", tmpX, 0)  'Modified by Lydia 2017/10/16 中間銀行Y座標
               If m_MiddleBank_01 <> "" Or m_MiddleBank_02 <> "" Or m_MiddleBank_03 <> "" Then
           'end 2023/04/13
                   '2011/6/28 改位置
                   'Printer.CurrentX = 300 - intX + 2367
                   'Printer.CurrentY = 7250 - intY - 2175
                   Printer.CurrentX = 1300 - intX
                   'Modified by Lydia 2015/04/29 台銀新版(104.04)
                   'Printer.CurrentY = 5035 - intY
                   'Printer.CurrentY = 5335 - intY
                   Printer.CurrentY = tmpY - intY: tmpY = tmpY + tmpX
                   Printer.Print "correspondent bank:"
                   'Printer.Font.Size = 12
                   'Modified by Lydia 2023/04/13
                   'If "" & rsB("A2214").Value <> "" Then
                   If m_MiddleBank_01 <> "" Then
                       '2011/6/28 改位置
                       'Printer.CurrentX = 300 - intX + 2367
                       'Printer.CurrentY = 7450 - intY - 2175
                       Printer.CurrentX = 1300 - intX
                        'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                      ' Printer.CurrentY = 5235 - intY
                      'Modified by Lydia 2015/04/29 台銀新版(104.04)
                       'Printer.CurrentY = 5335 - intY
                       'Printer.CurrentY = 5635 - intY
                       Printer.CurrentY = tmpY - intY: tmpY = tmpY + tmpX
                       'Modified by Lydia 2023/04/13 比照modify by sonia 2014/8/14 取消中間銀行前的1,2,3
                       'Printer.Print "1 " & rsB("A2214").Value
                       Printer.Print m_MiddleBank_01
                   End If
                   'Modified by Lydia 2023/04/13
                   'If "" & rsB("A2215").Value <> "" Then
                   If m_MiddleBank_02 <> "" Then
                       '2011/6/28 改位置
                       'Printer.CurrentX = 300 - intX + 2367
                       'Printer.CurrentY = 7650 - intY - 2175
                       Printer.CurrentX = 1300 - intX
                       'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                       'Printer.CurrentY = 5455 - intY
                       'Modified by Lydia 2015/04/29 台銀新版(104.04)
                       'Printer.CurrentY = 5555 - intY
                       'Printer.CurrentY = 5855 - intY
                       Printer.CurrentY = tmpY - intY: tmpY = tmpY + tmpX
                       'Modified by Lydia 2023/04/13 比照modify by sonia 2014/8/14 取消中間銀行前的1,2,3
                       'Printer.Print "2 " & rsB("A2215").Value
                       Printer.Print m_MiddleBank_02
                   End If
                   'Modified by Lydia 2023/04/13
                   'If "" & rsB("A2216").Value <> "" Then
                   If m_MiddleBank_03 <> "" Then
                       '2011/6/28 改位置
                       'Printer.CurrentX = 300 - intX + 2367
                       'Printer.CurrentY = 7850 - intY - 2175
                       Printer.CurrentX = 1300 - intX
                      'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                       'Printer.CurrentY = 5675 - intY
                       'Modified by Lydia 2015/04/29 台銀新版(104.04)
   '                    Printer.CurrentY = 5775 - intY
                       'Printer.CurrentY = 6075 - intY
                       Printer.CurrentY = tmpY - intY: tmpY = tmpY + tmpX
                       'Modified by Lydia 2023/04/13 比照modify by sonia 2014/8/14 取消中間銀行前的1,2,3
                       'Printer.Print "3 " & rsB("A2216").Value
                       Printer.Print m_MiddleBank_03
                   End If
               End If
               'Added by Lydia 2015/04/30
               'Modified by Lydia 2023/04/13
               'If "" & rsB("A2219").Value <> "" Then
               If m_A2219 <> "" Then
                  Printer.CurrentX = 1300 - intX
                  'Printer.CurrentY = 6295 - intY
                  Printer.CurrentY = tmpY - intY: tmpY = tmpY + tmpX
                  'Modified by Lydia 2023/04/13
                  'Printer.Print rsB("A2219").Value
                  Printer.Print m_A2219
               End If
           'End If 'Mark by Lydia 2023/04/13
           If rsB.State <> adStateClosed Then rsB.Close
           Set rsB = Nothing
           
           Printer.Font.Name = "Times New Roman"
           Printer.Font.Size = 12
           Printer.CurrentX = 7740 - intX
           'Modified by Lydia 2015/04/29 台銀新版(104.04)
           'Printer.CurrentY = 5775 - intY
           'Modified by Lydia 2018/11/09 台銀新版(107.08)
           'Printer.CurrentY = 6075 - intY
           Printer.CurrentY = 5350 - intY
           '2012/5/4 MODIFY BY SONIA 婧瑄說Y37580都不要印
           'Printer.Print "192 代理費支出"
            'Modified by Lydia 2015/10/06
           If "" & adoacc190_1.Fields("a1718").Value <> "" Then
              Printer.Print "19F"
           'Modified by Lydia 2018/04/23 寰華Y53374備註欄都不要再填寫(by 郭)
           'ElseIf Left(Trim(adoacc190_1.Fields("a1803").Value), 6) <> "Y37580" Then
           'Modified by Lydia 2018/06/11 寰華Y53374要印匯款類別
           'ElseIf InStr("Y37580,Y53374", Left(Trim(adoacc190_1.Fields("a1803").Value), 6)) = 0 Then
           ElseIf InStr("Y37580", Left(Trim(adoacc190_1.Fields("a1803").Value), 6)) = 0 Then
           'end 2015/10/06
              'Modified by Lydia 2017/04/07 匯款類別改為192 代理費支出
              'Printer.Print "19D 專業技術收入"    '2015/1/15 MODIFY BY SONIA 原為'192 代理費支出'
              'Modified by Lydia 2017/04/21 改回19D
              'Printer.Print "192 代理費支出"
              'Modified by Lydia 2017/09/14 改成常數
              'Printer.Print "19D 專業技術支出"
              Printer.Print Pub_DBtype
           Else
              Printer.Print ""
           End If
           '2012/5/4 end
           
           '受款人名稱
           '若不為電匯
           'Printer.Font.Name = "Arial"
           Printer.Font.Size = 10
           'Added by Lydia 2015/04/17 +匯款方式3:台銀電匯紙本
           If "" & adoacc190_1.Fields("a1811").Value <> "2" And "" & adoacc190_1.Fields("a1811").Value <> "3" Then
               '抓受款人相關資料
               'Modified by Lydia 2017/10/03 華銀整批媒體RMB改CNY
               'StrSqlB = "Select * From ACC220 Where a2201='" & adoacc190_1.Fields("a1803").Value & "' And a2202='" & adoacc190_1.Fields("a1903").Value & "' "
               'Mark by Lydia 2023/04/13 改成模組GetAccData
               ''受款人名稱
               'StrSqlB = "Select * From ACC220 Where a2201='" & adoacc190_1.Fields("a1803").Value & "' And a2202='" & IIf(adoacc190_1.Fields("a1903").Value = J_RMB, "RMB", adoacc190_1.Fields("a1903").Value) & "' "
               'rsB.CursorLocation = adUseClient
               'rsB.Open StrSqlB, adoTaie, adOpenStatic, adLockReadOnly
               'If rsB.RecordCount > 0 Then
               '    '受款人名稱
               '    If IsNull(rsB.Fields("a2203").Value) = False Then
               If m_A1810 <> "" Then
                     PrintDropLine UCase(m_A1810), 300 - intX + 3550, 6150 - intY + 1184, 200
               Else
                   If m_Payee_01 <> "" Then
               'end 2023/04/13
                      '2011/6/28 改位置
                      'Printer.CurrentX = 300 - intX + 3550
                      'Printer.CurrentY = 5150 - intY + 1184
                      Printer.CurrentX = 1300 - intX
                      'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                      'Printer.CurrentY = 6634 - intY
                      'Modified by Lydia 2015/04/29 台銀新版(104.04)
                     ' Printer.CurrentY = 7734 - intY
                      'Modified by Lydia 2018/11/09 台銀新版(107.08)
                      'Printer.CurrentY = 8034 - intY
                      'Modified by Lydia 2018/12/28 往下調
                      'Printer.CurrentY = 8394 - intY
                      Printer.CurrentY = 8494 - intY
                      '2011/6/29 modify by sonia 全印大寫
                      'Modified by Lydia 2023/04/13
                      'Printer.Print UCase(rsB.Fields("a2203").Value)
                      Printer.Print UCase(m_Payee_01)
                   End If
                   'Modified by Lydia 2023/04/13
                   'If IsNull(rsB.Fields("a2204").Value) = False Then
                   If m_Payee_02 <> "" Then
                      'Modify by Morgan 2009/12/24 要和第一列對齊--婧瑄
                      '2011/6/28 改位置
                      'Printer.CurrentX = 300 - intX + 3550
                      'Printer.CurrentY = 5350 - intY + 1184
                      Printer.CurrentX = 1300 - intX
                      'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                      'Printer.CurrentY = 6854 - intY
                      'Modified by Lydia 2015/04/29 台銀新版(104.04)
                      'Printer.CurrentY = 7954 - intY
                      'Modified by Lydia 2018/11/09 台銀新版(107.08)
                      'Printer.CurrentY = 8254 - intY
                      'Modified by Lydia 2018/12/28 往下調
                      'Printer.CurrentY = 8614 - intY
                      Printer.CurrentY = 8714 - intY
                      '2011/6/29 modify by sonia 全印大寫
                      'Modified by Lydia 2023/04/13
                      'Printer.Print UCase(rsB.Fields("a2204").Value)
                      Printer.Print UCase(m_Payee_02)
                   End If
                   'Modified by Lydia 2023/04/13
                   'If IsNull(rsB.Fields("a2205").Value) = False Then
                   If m_Payee_03 <> "" Then
                      'Modify by Morgan 2009/12/24 要和第一列對齊--婧瑄
                      '2011/6/28 改位置
                      'Printer.CurrentX = 300 - intX + 3550
                      'Printer.CurrentY = 5550 - intY + 1184
                      Printer.CurrentX = 1300 - intX
                      'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                      'Printer.CurrentY = 7074 - intY
                      'Modified by Lydia 2015/04/29 台銀新版(104.04)
                      'Printer.CurrentY = 8174 - intY
                      'Modified by Lydia 2018/11/09 台銀新版(107.08)
                      'Printer.CurrentY = 8474 - intY
                      'Modified by Lydia 2018/12/28 往下調
                      'Printer.CurrentY = 8834 - intY
                      Printer.CurrentY = 8934 - intY
                      '2011/6/29 modify by sonia 全印大寫
                      'Modified by Lydia 2023/04/13
                      'Printer.Print UCase(rsB.Fields("a2205").Value)
                      Printer.Print UCase(m_Payee_03)
                   End If
                   'Modifed by Lydia 2015/03/30 名稱改3欄,第4欄(a2206)改為CNAPS(a2220)
                   'Modified by Lydia 2023/04/13
                   'If IsNull(rsB.Fields("a2220").Value) = False Then
                   If m_Payee_04 <> "" Then
                      'Modify by Morgan 2009/12/24 要和第一列對齊--婧瑄
                      '2011/6/28 改位置
                      'Printer.CurrentX = 300 - intX + 3550
                      'Printer.CurrentY = 5750 - intY + 1184
                      Printer.CurrentX = 1300 - intX
                      'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                      'Printer.CurrentY = 7294 - intY
                      'Modified by Lydia 2015/04/29 台銀新版(104.04)
                      'Printer.CurrentY = 8394 - intY
                      'Modified by Lydia 2017/10/16 改在受款銀行名稱(和中間銀行)下方
                      'Printer.CurrentY = 8694 - intY
                      Printer.CurrentY = tmpY - intY: tmpY = tmpY + tmpX
                      '2011/6/29 modify by sonia 全印大寫
                      'Modifed by Lydia 2015/04/20 +抬頭CNAPS
                      'Modified by Lydia 2023/04/13
                      'Printer.Print "CNAPS:" & UCase(rsB.Fields("a2220").Value)
                      Printer.Print "CNAPS:" & UCase(m_Payee_04)
                   End If
               End If 'Added by Lydia 2023/04/13
               
               'Mark by Lydia 2023/04/13
               'Else
               '    '受款人名稱
               '    If "" & adoacc190_1.Fields("A1810").Value <> "" Then
               '        '2011/6/29 modify by sonia 全印大寫
               '        'Modified by Lydia 2018/11/09 台銀新版(107.08)
               '        'PrintDropLine UCase(adoacc190_1.Fields("A1810").Value), 300 - intX + 3550, 5150 - intY + 1184, 200
               '        PrintDropLine UCase(adoacc190_1.Fields("A1810").Value), 300 - intX + 3550, 6150 - intY + 1184, 200
               '    Else
               '        If IsNull(adoacc190_1.Fields("fa05").Value) = False Then
               '            '2011/6/28 改位置
               '            'Printer.CurrentX = 300 - intX + 3550
               '            'Printer.CurrentY = 5150 - intY + 1184
               '            Printer.CurrentX = 1300 - intX
               '            'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
               '            'Printer.CurrentY = 6634 - intY
               '            'Modified by Lydia 2015/04/29 台銀新版(104.04)
               '            'Printer.CurrentY = 7734 - intY
               '            'Modified by Lydia 2018/11/09 台銀新版(107.08)
               '            'Printer.CurrentY = 8034 - intY
               '            'Modified by Lydia 2018/12/28 往下調
               '            'Printer.CurrentY = 8394 - intY
               '            Printer.CurrentY = 8494 - intY
               '            '2011/6/29 modify by sonia 全印大寫
               '            Printer.Print UCase(adoacc190_1.Fields("fa05").Value)
               '        End If
               '        If IsNull(adoacc190_1.Fields("fa63").Value) = False Then
               '            '2011/6/28 改位置
               '            'Printer.CurrentX = 300 - intX + 2300
               '            'Printer.CurrentY = 5350 - intY + 1184
               '            Printer.CurrentX = 1300 - intX
               '            'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
               '            'Printer.CurrentY = 6854 - intY
               '            'Modified by Lydia 2015/04/29 台銀新版(104.04)
               '            'Printer.CurrentY = 7954 - intY
               '            'Modified by Lydia 2018/11/09 台銀新版(107.08)
               '            'Printer.CurrentY = 8254 - intY
               '            'Modified by Lydia 2018/12/28 往下調
               '            'Printer.CurrentY = 8614 - intY
               '            Printer.CurrentY = 8714 - intY
               '            '2011/6/29 modify by sonia 全印大寫
               '            Printer.Print UCase(adoacc190_1.Fields("fa63").Value)
               '        End If
               '        If IsNull(adoacc190_1.Fields("fa64").Value) = False Then
               '            '2011/6/28 改位置
               '            'Printer.CurrentX = 300 - intX + 1000
               '            'Printer.CurrentY = 5550 - intY + 1184
               '            Printer.CurrentX = 1300 - intX
               '            'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
               '            'Printer.CurrentY = 7074 - intY
               '            'Modified by Lydia 2015/04/29 台銀新版(104.04)
               '            'Printer.CurrentY = 8174 - intY
               '            'Modified by Lydia 2018/11/09 台銀新版(107.08)
               '            'Printer.CurrentY = 8474 - intY
               '            'Modified by Lydia 2018/12/28 往下調
               '            'Printer.CurrentY = 8834 - intY
               '            Printer.CurrentY = 8934 - intY
               '            '2011/6/29 modify by sonia 全印大寫
               '            Printer.Print UCase(adoacc190_1.Fields("fa64").Value)
               '        End If
               '         'Modifed by Lydia 2015/03/30 名稱改3欄,第4欄(a2206)改為CNAPS(a2220)
   '            '        If IsNull(adoacc190_1.Fields("fa65").Value) = False Then
   '            '            '2011/6/28 改位置
   '            '            'Printer.CurrentX = 300 - intX + 1000
   '            '            'Printer.CurrentY = 5750 - intY + 1184
   '            '            Printer.CurrentX = 1300 - intX
   '            '            Printer.CurrentY = 7294 - intY
   '            '            '2011/6/29 modify by sonia 全印大寫
   '            '            Printer.Print UCase(adoacc190_1.Fields("fa65").Value)
   '            '        End If
   
                '   End If
               'End If
               'If rsB.State <> adStateClosed Then rsB.Close
               'Set rsB = Nothing
               'end --- Mark by Lydia 2023/04/13
               
               'Added by Lydia 2017/07/20 無條件列印
               'Remove by Lydia 2017/08/01 票匯不加
               'Printer.Font.Size = 12
               'Printer.CurrentX = 7210 - intX
               'Printer.CurrentY = 6680 - intY
               'Printer.Print NoteTitle
               'end 2017/07/20
           '若為電匯
           Else
               'Printer.Font.Name = "Arial"
               Printer.Font.Size = 10
               '抓受款人相關資料
               'Modified by Lydia 2017/10/03 華銀整批媒體RMB改CNY
               'StrSqlB = "Select * From ACC220 Where a2201='" & adoacc190_1.Fields("a1803").Value & "' And a2202='" & adoacc190_1.Fields("a1903").Value & "' "
               'Mark by Lydia 2023/04/13 改成模組GetAccData
                'StrSqlB = "Select * From ACC220 Where a2201='" & adoacc190_1.Fields("a1803").Value & "' And a2202='" & IIf(adoacc190_1.Fields("a1903").Value = J_RMB, "RMB", adoacc190_1.Fields("a1903").Value) & "' "
               'rsB.CursorLocation = adUseClient
               'rsB.Open StrSqlB, adoTaie, adOpenStatic, adLockReadOnly
               'If rsB.RecordCount > 0 Then
               '    '受款銀行名稱
               '    If IsNull(rsB.Fields("a2208").Value) = False Then
                   If m_BeneBankName1 <> "" Then
               'end 2023/04/13
                      '2011/6/28 改位置
                      'Printer.CurrentX = 300 - intX + 1367
                      Printer.CurrentX = 1300 - intX
                      'Modified by Lydia 2018/11/09 台銀新版(107.08)
                      'Printer.CurrentY = 4775 - intY
                      Printer.CurrentY = 5560 - intY
                      '2011/6/29 modify by sonia 全印大寫
                      'Modified by Lydia 2023/04/13
                      'Printer.Print UCase(rsB.Fields("a2208").Value)
                      Printer.Print UCase(m_BeneBankName1)
                   End If
                   'Modified by Lydia 2023/04/13
                   'If IsNull(rsB.Fields("a2209").Value) = False Then
                   If m_BeneBankName2 <> "" Then
                      '2011/6/28 改位置
                      'Printer.CurrentX = 300 - intX + 1367
                      Printer.CurrentX = 1300 - intX
                      'Modified by Lydia 2015/04/29 台銀新版(104.04)
                      'Printer.CurrentY = 4595 - intY
                      'Modified by Lydia 2018/11/09 台銀新版(107.08)
                      'Printer.CurrentY = 4995 - intY
                      Printer.CurrentY = 5800 - intY
                      '2011/6/29 modify by sonia 全印大寫
                      'Modified by Lydia 2023/04/13
                      'Printer.Print UCase(rsB.Fields("a2209").Value)
                      Printer.Print UCase(m_BeneBankName2)
                   End If
                   Printer.Font.Name = "Times New Roman"
                   Printer.Font.Size = 12
                   '受款銀行帳號
                   'Modified by Lydia 2023/04/13
                   'If IsNull(rsB.Fields("a2210").Value) = False Then
                   If m_BeneBankName3 <> "" Then
                      '2011/6/28 改位置
                      'Printer.CurrentX = 300 - intX + 1367
                      Printer.CurrentX = 1300 - intX
                      'Modified by Lydia 2015/04/29 台銀新版(104.04)
                      'Printer.CurrentY = 4815 - intY
                      'Modified by Lydia 2018/11/09 台銀新版(107.08)
                      'Printer.CurrentY = 5215 - intY
                      Printer.CurrentY = 6010 - intY
                      '2011/6/29 modify by sonia 全印大寫
                      'Modified by Lydia 2017/09/14 SWIFT CODE 改為swift ,方便區隔後面的銀行代號
                      'Printer.Print UCase(rsB.Fields("a2210").Value & " " & rsB.Fields("a2211").Value)
                      'Modified by Lydia 2023/04/13
                      'Printer.Print Replace(UCase(rsB.Fields("a2210").Value & " " & rsB.Fields("a2211").Value), "SWIFT CODE", "swift")
                      Printer.Print m_BeneBankName3
                   End If
                   
                   'Added by Lydia 2017/10/16 CNAPS都顯示,改在受款銀行名稱(和中間銀行)下方
                   'Modified by Lydia 2023/04/13
                   'If "" & rsB.Fields("a2220").Value <> "" Then
                   If m_Payee_04 <> "" Then
                      Printer.CurrentX = 1300 - intX
                      Printer.CurrentY = tmpY - intY: tmpY = tmpY + tmpX
                      'Modified by Lydia 2023/04/13
                      'Printer.Print "CNAPS:" & UCase(Trim(rsB.Fields("a2220")))
                      Printer.Print "CNAPS:" & m_Payee_04
                   End If
                   'end 2017/10/16
               
                    '受款人帳號
                   'Modified by Lydia 2023/04/13
                   'If IsNull(rsB.Fields("a2207").Value) = False Then
                   If m_AccountNum <> "" Then
                       'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
   '                   Printer.CurrentX = 4050 - intX
   '                   Printer.CurrentY = 5929 - intY
                       Printer.CurrentX = 1300 - intX
                       'Modified by Lydia 2015/04/29 台銀新版(104.04)
                       'Printer.CurrentY = 7029 - intY
                       'Modified by Lydia 2018/11/09 台銀新版(107.08)
                       'Printer.CurrentY = 7279 - intY
                       Printer.CurrentY = 7740 - intY
                       'Modified by Lydia 2023/04/13
                       'Printer.Print rsB.Fields("a2207").Value
                       Printer.Print m_AccountNum
                   End If
   
                    'Modified by Lydia 2015/03/10  澳洲(015),南非(301)代理人的水單要印地址,名稱上移
                    'Modified by Lydia 2015/03/25 + 加拿大
                    'Modified by Lydia 2015/05/12 改成依匯款行國別
                   'If adoacc190_1.Fields("fa10") = "015" Or adoacc190_1.Fields("fa10") = "301" Or adoacc190_1.Fields("fa10") = "102" Then
                   'bolAddr = False 'Mark by Lydia 2023/04/13 改成模組GetAccData
                   'Modified by Lydia 2015/08/05 若受款行資料未設定國別,以代理人資料判斷
                   'If InStr("015,301,102", rsB.Fields("a2217")) > 0 Then
                   'Modified by Lydia 2017/07/19 台銀要求於南非, 加拿大, 澳洲和歐洲地區(m_na02)要印地址,若有短地址則優先列印
                   'Modified by Lydia 2017/08/01 台銀要求全部都要印地址
                   'If InStr("015,301,102", rsB.Fields("a2217")) > 0 Or InStr("015,301,102", adoacc190_1.Fields("fa10")) > 0 Or m_NA02 = "C20" Then
                       'bolAddr = True 'Mark by Lydia 2023/04/13 改成模組GetAccData
                    'end 2015/05/12
                       'Modified by Lydia 2017/08/22 +第3行名稱
                       'm_Payee_01 = Trim("" & rsB.Fields("a2203") & " " & rsB.Fields("a2204"))
                       'Mark by Lydia 2023/04/13 改成模組GetAccData
                       'm_Payee_01 = Trim("" & rsB.Fields("a2203") & " " & rsB.Fields("a2204")) & IIf("" & rsB.Fields("a2205") <> "", " " & rsB.Fields("a2205"), "")
                       'm_FAAddr = Trim("" & adoacc190_1.Fields("addr"))
                       ''Added by Lydia 2017/07/19 短地址優先列印
                       'If Trim("" & rsB.Fields("a2218")) <> "" Then m_FAAddr = Trim("" & rsB.Fields("a2218"))
                       'end -- Mark by Lydia 2023/04/13
                       
                       Printer.Font.Size = 10
                       '受款人名稱
                       'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                       'XPrint m_Payee_01, 3750, 6164, 3100
                       'Modified by Lydia 2015/04/29 台銀新版(104.04)
                       'XPrint m_Payee_01, 1300, 7264, 3100
                       'Modified by Lydia 2018/11/09 台銀新版(107.08)
                       'XPrint m_Payee_01, 3800, 7564, 3100
                       XPrint m_Payee_01, 3800, 8080, 3100
                       '代理人地址
                       'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                       'XPrint m_FAAddr, 1300, 6634, 5670
                       'Modified by Lydia 2015/04/29 台銀新版(104.04)
                       'XPrint m_FAAddr, 1300, 7734, 5670
                       'Modified by Lydia 2017/08/22 名稱多1行
                       'XPrint m_FAAddr, 1300, 8034, 5670
                       'Modified by Lydia 2018/11/09 台銀新版(107.08)
                       'XPrint m_FAAddr, 1300, 8434, 5670
                       XPrint m_FAAddr, 1300, 8600, 5670
                       Printer.Font.Size = 12
                   'Else
                  ''Printer.Font.Name = "Arial"
                   '    Printer.Font.Size = 10
                   '    '受款人名稱
                   '     If IsNull(rsB.Fields("a2203").Value) = False Then
                   '        '2011/6/28 改位置
                   '        'Printer.CurrentX = 300 - intX + 3550
                   '        'Printer.CurrentY = 5150 - intY + 1184
                   '        Printer.CurrentX = 1300 - intX
                   '        'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                   '        'Printer.CurrentY = 6634 - intY
                   '        'Modified by Lydia 2015/04/29 台銀新版(104.04)
   '               '         Printer.CurrentY = 7734 - intY
                   '        Printer.CurrentY = 8034 - intY
                   '        '2011/6/29 modify by sonia 全印大寫
                   '        Printer.Print UCase(rsB.Fields("a2203").Value)
                   '     End If
                   '     If IsNull(rsB.Fields("a2204").Value) = False Then
                   '        'Modify by Morgan 2009/12/24 要和第一列對齊--婧瑄
                   '        '2011/6/28 改位置
                   '        'Printer.CurrentX = 300 - intX + 3550
                   '        'Printer.CurrentY = 5350 - intY + 1184
                   '        Printer.CurrentX = 1300 - intX
                   '        'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                   '        'Printer.CurrentY = 6854 - intY
                   '        'Modified by Lydia 2015/04/29 台銀新版(104.04)
                   '        'Printer.CurrentY = 7954 - intY
                   '        Printer.CurrentY = 8254 - intY
                   '        '2011/6/29 modify by sonia 全印大寫
                   '        Printer.Print UCase(rsB.Fields("a2204").Value)
                   '     End If
                   '     If IsNull(rsB.Fields("a2205").Value) = False Then
                   '        'Modify by Morgan 2009/12/24 要和第一列對齊--婧瑄
                   '        '2011/6/28 改位置
                   '        'Printer.CurrentX = 300 - intX + 3550
                   '        'Printer.CurrentY = 5550 - intY + 1184
                   '        Printer.CurrentX = 1300 - intX
                   '        'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                   '        'Printer.CurrentY = 7074 - intY
                   '        'Modified by Lydia 2015/04/29 台銀新版(104.04)
                   '        'Printer.CurrentY = 8174 - intY
                   '        Printer.CurrentY = 8474 - intY
                   '        '2011/6/29 modify by sonia 全印大寫
                   '        Printer.Print UCase(rsB.Fields("a2205").Value)
                   '     End If
                   '      'Modifed by Lydia 2015/03/30 名稱改3欄,第4欄(a2206)改為CNAPS(a2220)
                   '     If IsNull(rsB.Fields("a2220").Value) = False Then
                   '        'Modify by Morgan 2009/12/24 要和第一列對齊--婧瑄
                   '        '2011/6/28 改位置
                   '        'Printer.CurrentX = 300 - intX + 3550
                   '        'Printer.CurrentY = 5750 - intY + 1184
                   '        Printer.CurrentX = 1300 - intX
                   '        'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
                   '        'Printer.CurrentY = 7294 - intY
                   '        'Modified by Lydia 2015/04/29 台銀新版(104.04)
                   '        'Printer.CurrentY = 8394 - intY
                   '        Printer.CurrentY = 8694 - intY
                   '        '2011/6/29 modify by sonia 全印大寫
                   '        'Modifed by Lydia 2015/04/20 +抬頭CNAPS
                   '        Printer.Print "CNAPS:" & UCase(rsB.Fields("a2220").Value)
                   '     End If
                   'End If
                   ''end 2015/03/10
                   'end 2017/08/01
               'Mark by Lydia 2023/04/13 改成模組GetAccData
               'End If
               'If rsB.State <> adStateClosed Then rsB.Close
               'Set rsB = Nothing
               'end ---- Mark by Lydia 2023/04/13 改成模組GetAccData
               
               'Remove by Lydia 2015/04/30 改用 A2219(手續費方式)
   '            'Add by Morgan 2006/7/14
   '            If adoacc190_1.Fields("a1803").Value = "Y45589000" Then
   '               Printer.CurrentX = 8517 - intX
   '               'Modify by Lydia 2015/04/16 台銀新版(手寫紙)
   '               'Printer.CurrentY = 6224 - intY
   '               Printer.CurrentY = 7324 - intY
   '               Printer.Print "ours"
   '            End If
   '            'end 2006/7/14
   
               'Modify by Lydia 2015/04/16 台銀新版(手寫紙),設備註欄數
               If rsB.State <> adStateClosed Then rsB.Close
               Set rsB = Nothing
               StrSqlB = "Select max(length(r42704)) From ACCRPT427 Where R42701='" & strUserNum & "' And R42702='" & strA0K11 & "' And R42703='" & adoacc190_1.Fields("a1803").Value & "' Order By 1 "
               rsB.CursorLocation = adUseClient
               rsB.Open StrSqlB, adoTaie, adOpenStatic, adLockReadOnly
               tmpCol = 4: tmpNum = 20
               If rsB.Fields(0) >= 9 Then
                  tmpCol = 3: tmpNum = 15
               End If
               'Remove by Lydia 2016/02/15
               'strDBMax = rsB.Fields(0) 'Added by Lydia 2016/01/22
               If rsB.State <> adStateClosed Then rsB.Close
               Set rsB = Nothing
               'end 2015/04/16 台銀新版位置,設備註欄數
   
               
               '抓代理人D/BNo
               StrSqlB = "Select Distinct R42704 From ACCRPT427 Where R42701='" & strUserNum & "' And R42702='" & strA0K11 & "' And R42703='" & adoacc190_1.Fields("a1803").Value & "' Order By 1 "
               rsB.CursorLocation = adUseClient
               rsB.Open StrSqlB, adoTaie, adOpenStatic, adLockReadOnly
               ii = 0
               'Added by Lydia 2015/04/01 先印媒體備註
               'tmpX = 7317: tmpY = 6430
               'Modified by Lydia 2015/04/29 台銀新版(104.04)
               'tmpX = 7150: tmpY = 6400
               'Modified by Lydia 2015/07/22 備註往下移
               'tmpX = 7210: tmpY = 6580
               'Modified by Lydia 2018/11/09 台銀新版(107.08)
               'tmpX = 7210: tmpY = 6680
               tmpX = 7210: tmpY = 6780
               If Len(Trim(strA2222)) > 0 Then
                  Printer.CurrentX = tmpX - intX
                  Printer.CurrentY = tmpY - intY
                  'Modified by Lydia 2018/11/09 台銀新版(107.08) :長度從20改30
                  If GetTextLength_1(strA2222) > 30 Then
                     strExc(9) = PUB_StrToStr(strA2222, 30)
                     strExc(10) = Trim(MidB(strA2222, LenB(strExc(9)) + 1))
                     Printer.Print strExc(9)
                     tmpY = tmpY + Printer.TextHeight(strExc(9)) + 30
                  Else
                     strExc(10) = Trim(strA2222)
                  End If
                  Printer.CurrentX = tmpX - intX
                  Printer.CurrentY = tmpY - intY
                  Printer.Print PUB_StrToStr(strExc(10), 30)
                  tmpY = tmpY + Printer.TextHeight(strExc(10)) + 30
                  'end 2018/11/09
               End If
               'end 2015/04/01
               
               'Modified by Lydia 2017/07/20 DB note 加 INV.
               'Modified by Lydia 2017/08/14 暫收款退費不加INV. (a1902=a1702為O單號)
               'strDBNote = "" & NoteTitle
               'Modified by Lydia 2018/04/23 寰華Y53374備註欄都不要再填寫(by 郭)
               'If Val(adoacc190_1.Fields("Ocnt")) = 0 Then
               'Modified by Lydia 2020/04/22 +天津三元Y37580; +建毅Y51566,唯源Y52404
               If InStr("Y37580,Y53374,Y51566,Y52404,", Left(Trim(adoacc190_1.Fields("a1803").Value), 6)) = 0 And Val(adoacc190_1.Fields("Ocnt")) = 0 Then
                  strDBNote = "" & NoteTitle
               Else
                  strDBNote = ""
               End If
               'end 2017/08/14
               
               'Added by Lydia 2017/07/20 無條件列印
               If rsB.EOF = True Then
                  Printer.Font.Size = 12
                  Printer.CurrentX = tmpX - intX
                  Printer.CurrentY = tmpY - intY
                  Printer.Print strDBNote
               End If
               'end 2017/07/20
               
               'Modify by Lydia 2015/04/16 台銀新版(手寫紙),設備註欄數
               '總欄位數12=>tmpNum ,一列2欄=>tmpCol
               Do While Not rsB.EOF
                   ii = ii + 1
                   'Modify by Lydia 2015/04/16
   '                If ii > 12 Then Exit Do
   '                If ii Mod 2 = 1 Then
                   If ii > tmpNum Then Exit Do
                   If ii Mod tmpCol > 0 Then
                      'Added by Lydia 2015/03/30 +媒體備註
                     ' strDBNote = strDBNote & rsB.Fields(0).Value
                      'Modified by Lydia 2017/07/20 DB note 加 INV.
                      'strDBNote = IIf(Len(strDBNote) > 0, strDBNote & ",", "") & rsB.Fields(0).Value
                      strDBNote = strDBNote & IIf(Len(strDBNote) > 0 And strDBNote <> NoteTitle, ",", "") & rsB.Fields(0).Value
                   Else
                       strDBNote = strDBNote & "," & rsB.Fields(0).Value
   '                    Printer.CurrentX = 7317 - intX
   '                    Printer.CurrentY = 6430 - intY + (ii / 2 - 1) * 200
                       Printer.CurrentX = tmpX - intX
                       'Modify by Lydia 2015/04/16
                       'Printer.CurrentY = tmpY - intY + (ii / 2 - 1) * Printer.TextHeight("W")
                       Printer.CurrentY = tmpY - intY + (ii / tmpCol - 1) * Printer.TextHeight("W")
                       'Modify By Sindy 2010/12/15
                       If Left(Trim(adoacc190_1.Fields("a1803").Value), 6) = "Y52401" Then
                           Printer.Print "Honorarium"
                       '2010/12/15 End
                       Else
                          'Modify by Lydia 2015/04/16
                           'Printer.Print strDBNote & IIf(ii = 12 And rsB.RecordCount > 12, " etc.", "")
                           Printer.Print strDBNote & IIf(ii = tmpNum And rsB.RecordCount > tmpNum, " etc.", "")
                       End If
                       strDBNote = ""
                   End If
                   rsB.MoveNext
               Loop
               'Modify by Lydia 2015/04/16
               'If ii <= 12 And ii Mod 2 = 1 Then
               If ii <= tmpNum And ii Mod tmpCol > 0 Then
                   'Modified by Lydia 2015/04/01
   '                Printer.CurrentX = 7317 - intX
   '                Printer.CurrentY = 6430 - intY + ((ii + 1) / 2 - 1) * 200
                   Printer.CurrentX = tmpX - intX
                   'Modify by Lydia 2015/04/16
                   'Printer.CurrentY = tmpY - intY + IIf(ii > 1, (ii \ 2), 0) * Printer.TextHeight("W")
                   'Modified by Lydia 2016/01/21
                   'Printer.CurrentY = tmpY - intY + IIf(ii > 1, (ii \ tmpCol), 0) * Printer.TextHeight("W")
                   tmpY = tmpY + IIf(ii > 1, (ii \ tmpCol), 0) * Printer.TextHeight("W")
                   Printer.CurrentY = tmpY - intY
                   'Modify By Sindy 2010/12/15
                   If Left(Trim(adoacc190_1.Fields("a1803").Value), 6) = "Y52401" Then
                       Printer.Print "Honorarium"
                   '2010/12/15 End
                   Else
                       'Modified by Lydia 2016/01/22 備註折行
                       'Printer.Print strDBNote
                       'Modified by Lydia 2016/01/26 備註折行(依a1716的換行符號)
   '                     If strDBMax <= 36 Then
   '                        Printer.Print strDBNote
   '                     Else
   '                        '除了DB note,尚有A1716(A1718的備註)
   '                        strExc(7) = PUB_RepToOneSpace(PUB_StringFilter(strDBNote))
   '                        Do While Len(strExc(7)) > 0
   '                            strExc(8) = Mid(strExc(7), 1, 36)
   '                            Printer.CurrentX = tmpX - intX
   '                            Printer.CurrentY = tmpY - intY
   '                            Printer.Print strExc(8)
   '                            strExc(7) = Mid(strExc(7), 37)
   '                            tmpY = tmpY + Printer.TextHeight("W") + 20
   '                        Loop
   '                     End If
            '-----------------------
                        strExc(7) = strDBNote
                        Do While strExc(7) <> ""
                           intI = InStr(strExc(7), vbCrLf)
                           If intI = 0 Then
                              Printer.CurrentX = tmpX - intX
                              Printer.CurrentY = tmpY - intY
                              Printer.Print strExc(7)
                              strExc(7) = ""
                           Else
                              strExc(8) = Left(strExc(7), intI - 1)
                              strExc(7) = Mid(strExc(7), intI + 2)
                              Printer.CurrentX = tmpX - intX
                              Printer.CurrentY = tmpY - intY
                              If tmpY >= 7500 Then
                                  Printer.Print strExc(8) & " etc."
                                  strExc(7) = ""
                              Else
                                  Printer.Print strExc(8)
                              End If
                           End If
                           tmpY = tmpY + Printer.TextHeight("W") + 20
                        Loop
                   End If
                   strDBNote = ""
               End If
   
               If rsB.State <> adStateClosed Then rsB.Close
               Set rsB = Nothing
           End If  'Memo by Lydia 2023/04/13  --end If "" & adoacc190_1.Fields("a1811").Value <> "2" And "" & adoacc190_1.Fields("a1811").Value <> "3" Then
        End If 'Added by Lydia 2023/04/13
        adoacc190_1.MoveNext
   Loop
   adoacc190_1.Close
   If txtWordTB <> "Y" Then 'Added by Lydia 2023/04/13
      'add by nickc   2005/05/17
      If IsEnd = True Then
         Printer.EndDoc
      Else
         Printer.NewPage
      End If
   End If 'Added by Lydia 2023/04/13
End Sub

Private Sub XPrint(ByVal pContent As String, Px As Long, Py As Long, pWidth As Long)
   Dim iPos As Integer, iRow As Integer
   Dim strTmp As String, strAdd As String
   
   pContent = Trim(pContent)
   If pContent = "" Then Exit Sub
   
   For iRow = 0 To 3
      strTmp = ""
      iPos = InStr(pContent, " ")
      If iPos = 0 Then
         strTmp = pContent
         pContent = ""
      Else
         strTmp = Left(pContent, iPos - 1)
         pContent = LTrim(Mid(pContent, iPos + 1))
      End If
      
      Do While pContent <> ""
         iPos = InStr(pContent, " ")
         If iPos = 0 Then
            strAdd = pContent
         Else
            strAdd = Left(pContent, iPos - 1)
         End If
         If Printer.TextWidth(strTmp & " " & strAdd) > pWidth Then
            Exit Do
         Else
            strTmp = strTmp & " " & strAdd
            If iPos = 0 Then
               pContent = ""
            Else
               pContent = LTrim(Mid(pContent, iPos + 1))
            End If
         End If
      Loop
      
      Printer.CurrentX = Px - intX
      'Modified by Lydia 2017/08/22 加行間高度
      'Printer.CurrentY = Py - intY + 200 * iRow
      Printer.CurrentY = Py - intY + 228 * iRow
      Printer.Print strTmp
      strTmp = ""
      If pContent = "" Then Exit For
   Next
End Sub

Private Sub txtkind_GotFocus()
   TextInverse txtKind
End Sub

Private Sub txtKind_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
    Select Case KeyAscii
    Case 8, 89
        '無動作
    Case Else
        KeyAscii = 0
    End Select
End Sub

'Added by Lydia 2017/09/14
Private Sub Text4_Validate(Cancel As Boolean)
  If Text4 <> "" Then '預設代入
     Text3 = Text4
  End If
End Sub

'Added by Lydia 2023/04/13
Private Sub txtWordTB_GotFocus()
   TextInverse txtWordTB
End Sub

'Added by Lydia 2023/04/13
Private Sub txtWordTB_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
    Select Case KeyAscii
    Case 8, 89
        '無動作
    Case Else
        KeyAscii = 0
    End Select
End Sub

'Added by Lydia 2023/04/13 台銀新版(手寫紙),設備註欄數 =>改模組
Private Function GetDBnote(ByVal mA0k11 As String, ByVal mA1901 As String, ByVal mA1803 As String, ByVal mOcnt As String) As String
Dim strQ1 As String, intQ As Integer, intB As Integer
Dim rsQuery As New ADODB.Recordset
Dim strMidNote As String
Dim tmpCol  As Integer, tmpNum As Integer
Dim tmpArr As Variant
Dim lenMax As Single, lenHigh As Integer
   
   lenMax = 40: lenHigh = 4  'Word版備註最多40個字元,高度4行
   
   'Modify by Lydia 2015/04/16 台銀新版(手寫紙),設備註欄數
   'Modified by Lydia 2018/04/23 寰華Y53374備註欄都不要再填寫(by 郭)
   'Modified by Lydia 2020/04/22  +建毅Y51566,唯源Y52404
   If InStr("Y37580,Y53374,Y51566,Y52404,", Left(Trim(mA1803), 6)) = 0 Then
       strQ1 = "Select max(length(a1706))  From acc190, acc170 Where a1902=a1702 And a1901='" & mA1901 & "' Group By a1706 Order By 1 "
   Else
       strQ1 = "select null from dual"
   End If
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strQ1)
   If txtWordTB = "Y" Then
       tmpCol = 4: tmpNum = 4 * lenHigh
   Else
       tmpCol = 4: tmpNum = 20
   End If
   If intQ = 1 Then
      If rsQuery.Fields(0) >= 9 Then
          If txtWordTB = "Y" Then 'Word版備註
              If (Len(NoteTitle) + (rsQuery.Fields(0) * 2)) > lenMax Then
                 tmpCol = 1: tmpNum = 1 * lenHigh
              ElseIf (Len(NoteTitle) + (rsQuery.Fields(0) * 3)) > lenMax Then
                 tmpCol = 2: tmpNum = 2 * lenHigh
              Else
                 tmpCol = 3: tmpNum = 3 * lenHigh
              End If
          Else
              tmpCol = 3: tmpNum = 15
          End If
      End If
         'Modified by Lydia 2017/07/20 DB note 加NoteTitle
         'Modified by Lydia 2017/08/01 票匯不加NoteTitle
         'Modified by Lydia 2017/08/14 暫收款退費不加INV. (a1902=a1702為O單號)
         'Modified by Lydia 2018/04/23 寰華Y53374備註欄都不要再填寫(by 郭)
         'Modified by Lydia 2020/04/22 +天津三元Y37580; +建毅Y51566,唯源Y52404
         If InStr("Y37580,Y53374,Y51566,Y52404", Left(Trim(mA1803), 6)) = 0 And m_A1811 <> "1" And Val("" & mOcnt) = 0 Then
            strMidNote = NoteTitle
         End If
         '抓代理人D/BNo
         '2012/5/4 MODIFY BY SONIA 婧瑄說Y37580都不要印代理人D/BNo
         'Modified by Lydia 2018/04/23 寰華Y53374備註欄都不要再填寫(by 郭)
         'Modified by Lydia 2020/04/22 +建毅Y51566,唯源Y52404
         If InStr("Y37580,Y53374,Y51566,Y52404,", Left(Trim(mA1803), 6)) = 0 Then
            'Modified by Lydia 2015/10/06 +A1718,A1716
            strQ1 = "Select a1706,a1718,a1716 From acc190, acc170 Where a1902=a1702 And a1901='" & mA1901 & "' Group By a1706,a1718,a1716 Order By 1 "
         Else
            strQ1 = "select null as a1706, null as a1718, null as a1716 from dual"
         End If
         intQ = 1
         Set rsQuery = ClsLawReadRstMsg(intQ, strQ1)
         If intQ = 1 Then
             rsQuery.MoveFirst
             Do While Not rsQuery.EOF
                  intB = intB + 1
                  If intB > tmpNum Then Exit Do
                  If intB Mod tmpCol > 0 Then
                     'Added by Lydia 2015/03/30 +媒體備註
                     'Modified by Lydia 2017/07/20 DB note 加NoteTitle
                     strMidNote = strMidNote & IIf(Len(strMidNote) > 0 And strMidNote <> NoteTitle, ",", "") & PUB_RepToOneSpace(PUB_StringFilter("" & rsQuery.Fields(0).Value))
                  Else
                     strMidNote = strMidNote & "," & PUB_RepToOneSpace(PUB_StringFilter("" & rsQuery.Fields(0).Value))
                     If Left(Trim(mA1803), 6) = "Y52401" Then
                        strMidNote = "Honorarium"
                     Else
                        strMidNote = strMidNote & IIf(intB = tmpNum And rsQuery.RecordCount > tmpNum, " etc.", "")
                     End If
                     strMidNote = strMidNote & vbCrLf  '人工換行
                  End If
                  If "" & rsQuery.Fields("a1718") <> "" Then
                     strSql = "Insert Into ACCRPT427 Values('" & strUserNum & "','" & mA0k11 & "','" & mA1803 & "','" & ChgSQL("" & rsQuery.Fields("a1706").Value & rsQuery.Fields("a1716").Value) & "' )"
                  Else
                     strSql = "Insert Into ACCRPT427 Values('" & strUserNum & "','" & mA0k11 & "','" & mA1803 & "','" & ChgSQL("" & rsQuery.Fields(0).Value) & "' )"
                  End If
                  adoTaie.Execute strSql
                  rsQuery.MoveNext
             Loop
             If intB <= tmpNum And intB Mod tmpCol > 0 Then
                If Left(Trim(mA1803), 6) = "Y52401" Then
                   strMidNote = "Honorarium"
                End If
             End If
             'Added by Lydia 2015/10/06 +A1716(A1718的備註) =>不存ACCRPT427
             rsQuery.MoveFirst
             strExc(7) = ""
             Do While Not rsQuery.EOF
                If "" & rsQuery.Fields("a1718") <> "" Then
                   strExc(7) = strExc(7) & "," & rsQuery.Fields("a1716")
                End If
                rsQuery.MoveNext
             Loop
             If strExc(7) <> "" Then
                If GetTextLength(strMidNote) < 240 Then
                   strMidNote = PUB_StrToStr(strMidNote & vbCrLf & strExc(7), 240)
                End If
             End If
             'end 2015/10/06
         End If
         'Added by Lydia 2015/04/01 先印媒體備註=>不存ACCRPT427
         If Len(Trim(strA2222)) > 0 Then
            strMidNote = strA2222 & vbCrLf & strMidNote
            If GetTextLength(strMidNote) > 140 Then
               strMidNote = PUB_StrToStr(strMidNote, 240)
            End If
         End If
         'Word版備註: 人工換行
         If txtWordTB = "Y" Then
            strQ1 = ""
            tmpArr = Split(strMidNote, vbCrLf)
            intQ = 0
            For intB = 0 To UBound(tmpArr)
               If Trim(tmpArr(intB)) <> "" And intQ < lenHigh + 1 Then
                  intQ = intQ + 1
                  If intQ > lenHigh Then
                     strQ1 = strQ1 & " .etc"
                  Else
                     If GetTextLength("" & tmpArr(intB)) > lenMax Then
                        strExc(1) = tmpArr(intB)
                        Do While strExc(1) <> ""
                           If GetTextLength(strExc(1)) > lenMax Then
                               strExc(2) = Trim(PUB_StrToStr(strExc(1), lenMax))
                               strQ1 = strQ1 & IIf(strQ1 <> "", vbCrLf, "") & strExc(2)
                               strExc(1) = Trim(Replace(strExc(1), strExc(2), ""))
                               If strExc(1) <> "" Then intQ = intQ + 1
                           Else
                               strQ1 = strQ1 & IIf(strQ1 <> "", vbCrLf, "") & strExc(1)
                               strExc(1) = ""
                           End If
                        Loop
                     Else
                        strQ1 = strQ1 & IIf(strQ1 <> "", vbCrLf, "") & tmpArr(intB)
                     End If
                  End If
               End If
            Next intB
            strMidNote = strQ1
         End If
   End If
   Set rsQuery = Nothing
   GetDBnote = strMidNote
End Function

'Added by Lydia 2023/04/13 台銀水單改成Word套印
Private Sub TBCallWordPrint(ByRef rsPD As ADODB.Recordset, ByVal pKind As String)
Dim strName As String
Dim strText As String
Dim intA As Integer
Dim intFont As Integer

On Error GoTo ErrHand
   
   Erase strTempTB
   '代理人編號
   strTempTB(cntTB) = m_A1803 & IIf(pKind = "2", " 合計", "")
   '受款地區國別
   strTempTB(1) = m_NationName
   '匯款幣別金額
   strTempTB(2) = m_Amount
   '國外受款人身分別(民間)：Word預設
   '匯款方式
   If m_A1811 = "2" Then '電匯
       strTempTB(3) = "Y"
       strTempTB(4) = ""
   Else   '票匯
       strTempTB(3) = ""
       strTempTB(4) = "Y"
   End If
   '繳款方式
       '以新台幣結購: 獨立水單以台幣結匯, 有代為結匯之客戶編號, 非美金(參考Modify by Morgan 2010/6/15 取消歐元)
   If "" & rsPD.Fields("a1812") = "Y" Or "" & rsPD.Fields("a1718") <> "" Or "" & rsPD.Fields("a1903") <> "USD" Then
      strTempTB(5) = "Y"
      strTempTB(6) = ""  '不印金額
   Else
      strTempTB(5) = ""
      strTempTB(6) = ""
   End If
       '以外匯存款提出：目前只有美金
   If Not ("" & rsPD.Fields("a1812") = "Y" Or "" & rsPD.Fields("a1718") <> "") And "" & rsPD.Fields("a1903").Value = "USD" Then
      strTempTB(7) = "Y"
      strTempTB(8) = rsPD.Fields("a1903").Value & Format(rsPD.Fields("Amount").Value, FDollar) '幣別金額
   Else
      strTempTB(7) = ""
      strTempTB(8) = ""
   End If
   '申請人名稱:匯款人名稱(收據公司別)
   strTempTB(9) = strCompanyName
   '申請人統一編號
   strTempTB(10) = strCompanyNo
   '申請人地址
   If "" & rsPD.Fields("a1718").Value <> "" Then '有代為結匯之客戶編號,列印的申請人資料為該客戶的資料
       strTempTB(11) = Replace(strAddress, "\", vbCrLf)
   Else
       strTempTB(11) = strCompAddr1 & vbCrLf & strCompAddr2
   End If
   '申請人電話
   If "" & rsPD.Fields("a1718").Value <> "" Then '有代為結匯之客戶編號,列印的申請人資料為該客戶的資料
       strTempTB(12) = strPhone
   Else
       strTempTB(12) = strPhone
   End If

   '受款銀行名稱、地址、CNAPS ：電匯的受款銀行+中間銀行資料，固定寫CNAPS=m_A2219、m_Payee_04
   'Mark by Lydia 2025/07/31 中間銀行只顯示SWIFT
   'If m_MiddleBank_01 <> "" Or m_MiddleBank_02 <> "" Or m_MiddleBank_03 <> "" Then
   '   strTempTB(13) = "correspondent bank:"
   '   If m_MiddleBank_01 <> "" Then strTempTB(13) = strTempTB(13) & vbCrLf & m_MiddleBank_01
   '   If m_MiddleBank_02 <> "" Then strTempTB(13) = strTempTB(13) & vbCrLf & m_MiddleBank_02
   '   If m_MiddleBank_03 <> "" Then strTempTB(13) = strTempTB(13) & vbCrLf & m_MiddleBank_03
   'Else
      strTempTB(13) = ""
   'End If 'Mark by Lydia 2025/07/31 中間銀行只顯示SWIFT
   
   '受款人若不為電匯: '2006/1/18 MODIFY BY SONIA 婧瑄說其他結匯不管是否電匯都抓A1810
   If m_A1811 <> "2" Or Len(m_A1803) = 5 Then
       strTempTB(14) = ""
       strTempTB(15) = ""  '受款人名稱
       strTempTB(16) = ""  '受款人地址
       If Trim(m_Payee_01 & m_Payee_02 & m_Payee_03 & m_Payee_04) <> "" Then
           If m_Payee_01 <> "" Then strTempTB(16) = strTempTB(16) & IIf(strTempTB(16) <> "", vbCrLf, "") & m_Payee_01
           If m_Payee_02 <> "" Then strTempTB(16) = strTempTB(16) & IIf(strTempTB(16) <> "", vbCrLf, "") & m_Payee_02
           If m_Payee_03 <> "" Then strTempTB(16) = strTempTB(16) & IIf(strTempTB(16) <> "", vbCrLf, "") & m_Payee_03
           'Modified by Lydia 2017/10/16 改在受款銀行名稱(和中間銀行)下方
           If m_Payee_04 <> "" Then
              strTempTB(13) = strTempTB(13) & IIf(strTempTB(13) <> "", vbCrLf, vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf) & "CNAPS:" & m_Payee_04
           End If
       End If
   '若為電匯
   Else
       strExc(1) = strTempTB(13)  '先取得中間銀行資料
       '受款銀行名稱、地址、CNAPS ：電匯的受款銀行+中間銀行資料，固定寫CNAPS=m_A2219、m_Payee_04
       strTempTB(13) = ""
       If m_BeneBankName1 <> "" Then strTempTB(13) = strTempTB(13) & m_BeneBankName1 & vbCrLf
       If m_BeneBankName2 <> "" Then strTempTB(13) = strTempTB(13) & m_BeneBankName2 & vbCrLf
       'Modified by Lydia 2017/09/14 SWIFT CODE 改為swift ,方便區隔後面的銀行代號
       If m_BeneBankName3 <> "" Then strTempTB(13) = strTempTB(13) & Replace(m_BeneBankName3, "SWIFT CODE", "swift") & vbCrLf
       If strExc(1) <> "" Then
          strTempTB(13) = strTempTB(13) & strExc(1) & vbCrLf
       End If
       If m_A2219 <> "" Then strTempTB(13) = strTempTB(13) & m_A2219 & vbCrLf
       If m_Payee_04 <> "" Then strTempTB(13) = strTempTB(13) & "CNAPS:" & m_Payee_04 & vbCrLf
       '受款人帳號
       strTempTB(14) = m_AccountNum
       strTempTB(15) = ""  '受款人名稱
       strTempTB(16) = ""  '受款人地址
       '受款人名稱：澳洲代理人的水單要印地址,名稱上移；
       'modify by sonia 2014/9/3 加南非301
       'Modified by Lydia 2015/03/25 + 加拿大
       'Modified by Lydia 2015/05/12 改成依匯款行國別
       If bolAddr = True Then
          strTempTB(15) = m_Payee_01
          strTempTB(16) = m_FAAddr
       Else
         If Trim(m_Payee_01 & m_Payee_02 & m_Payee_03 & m_Payee_04) <> "" Then
             If m_Payee_01 <> "" Then strTempTB(16) = strTempTB(16) & IIf(strTempTB(16) <> "", vbCrLf, "") & m_Payee_01
             If m_Payee_02 <> "" Then strTempTB(16) = strTempTB(16) & IIf(strTempTB(16) <> "", vbCrLf, "") & m_Payee_02
             If m_Payee_03 <> "" Then strTempTB(16) = strTempTB(16) & IIf(strTempTB(16) <> "", vbCrLf, "") & m_Payee_03
             'Modified by Lydia 2017/10/16 改在受款銀行名稱(和中間銀行)下方
             If m_Payee_04 <> "" Then
                strTempTB(13) = strTempTB(13) & IIf(strTempTB(13) <> "", vbCrLf, vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf) & "CNAPS:" & m_Payee_04
             End If
         End If
       End If
   End If
   '匯款分類名稱及編號
   If "" & rsPD.Fields("a1718").Value <> "" Then
       strTempTB(17) = "19F"
   ElseIf InStr("Y37580", Left(Trim(rsPD.Fields("a1803").Value), 6)) = 0 Then
       strTempTB(17) = Pub_DBtype
   Else
       strTempTB(17) = ""
   End If
   '台銀新版(手寫紙),設備註欄數 =>改模組
   If m_A1811 <> "2" Or Len(m_A1803) = 5 Then '若不為電匯: 不印備註
      strTempTB(18) = ""
   Else
      strTempTB(18) = GetDBnote("" & rsPD.Fields("a0k11"), "" & "" & rsPD.Fields("a1901"), "" & rsPD.Fields("a1803").Value, "" & rsPD.Fields("Ocnt"))
   End If
   
   'Added by Lydia 2025/07/31
   '申請人-城市、國別
   strTempTB(19) = "TAIPEI"
   strTempTB(20) = "TW"
   '受款銀行-城市、國別、中間銀行SwiftCode
   strTempTB(21) = m_BankCity
   strTempTB(22) = m_BankNA
   strTempTB(23) = m_BankMidCode
   '受款人-城市、國別
   strTempTB(24) = m_RecCity
   strTempTB(25) = m_RecNA
   'end 2025/07/31
'-------------------------------------------
   '判斷word是否已開啟
   If Pub_NewWordDoc(g_WordAp, bVisible, m_WordLeft, m_WordTop) = False Then Exit Sub
   
   tb_TempFileName = "$$臺銀匯出匯款申請書_" & strSrvDate(1) & ServerTime & ".doc"
   If Dir(mDefDir & "\" & tb_TempFileName) <> "" Then
      Kill mDefDir & "\" & tb_TempFileName
   End If
   g_WordAp.Documents.Open mDefDir & "\" & tb_FileName
   g_WordAp.ActiveDocument.SaveAs mDefDir & "\" & tb_TempFileName
   g_WordAp.ActiveDocument.Close
   g_WordAp.Documents.Open mDefDir & "\" & tb_TempFileName
   
   '找出特定TextBox名稱;保留程式，以後改版可以用
'   For intI = 1 To g_WordAp.ActiveDocument.Shapes.Count
'         If InStr(UCase(g_WordAp.ActiveDocument.Shapes(intI).Name), "TEXT") > 0 Then
'            strExc(1) = strExc(1) & "Name: " & g_WordAp.ActiveDocument.Shapes(intI).Name & vbCrLf & _
'                                  "     Text:" & g_WordAp.ActiveDocument.Shapes(intI).TextFrame.TextRange.Text
'         End If
'   Next intI
'   Debug.Print strExc(1)
   
   With g_WordAp
      .Selection.WholeStory
      .Selection.Copy
      For intA = 1 To cntTB
         strName = ""
         strText = ""
         
         If intA = cntTB Then
            '代理人編號：固定在最後替換
            'Modified by Lydia 2025/07/31
            '.ActiveDocument.Shapes("Text Box 7").Select
            .ActiveDocument.Shapes("Text Box 8").Select
            .Selection.ShapeRange.TextFrame.TextRange.Select
            .Selection.TypeText Text:=strTempTB(cntTB)
            .Selection.HomeKey Unit:=wdStory
         Else
             strName = "T" & Format(intA, "000")
             strText = "" & strTempTB(intA)
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
            If InStr("001,002,003,004,005,007,017,", Format(intA, "000")) > 0 Then
               .Selection.Font.Size = 12
            ElseIf InStr("006,008,010,012,014,018", Format(intA, "000")) > 0 Then
               .Selection.Font.Size = 11
            ElseIf intA = 9 Then
               .Selection.Font.Size = 9
            Else
               .Selection.Font.Size = 10
            End If
            If InStr("003,004,005,007", Format(intA, "000")) > 0 Then '核取項目：電匯/票匯、以新臺幣結購/外匯存款
               If strText = "Y" Then
                  .Selection.InsertSymbol Font:="Wingdings", CharacterNumber:=-3843, Unicode:=True '已核取
               Else
                  .Selection.InsertSymbol Font:="Wingdings", CharacterNumber:=-3985, Unicode:=True '未核取
               End If
            Else
               .Selection.TypeText strText
            End If
         End If
      Next intA
   End With
   
   '正反兩面都印並且一件二份
   g_WordAp.PrintOut Background:=False, Range:=4, Item:=0, Copies:=2, Pages:="1-2", Collate:=True
   Pub_RePosWord g_WordAp, bVisible, m_WordLeft, m_WordTop  '模組還原Word位置
   g_WordAp.Quit wdDoNotSaveChanges
   Set g_WordAp = Nothing '避免快速開啟Word,程式出錯
   
   Exit Sub
   
ErrHand:

   If Err.Number <> 0 Then
      MsgBox Err.Number & ":" & Err.Description, , "錯誤 "
   End If
End Sub

'Added by Lydia 2023/04/13 取得相關資料; 台銀和華銀共用
Private Sub GetAccData(ByRef rsPD As ADODB.Recordset, ByRef pA0K11 As String)
Dim intQ As Integer, strQ1 As String
Dim rsQuery As New ADODB.Recordset
   
   pA0K11 = "" & rsPD.Fields("a0k11")
   '代理人
   m_A1803 = "" & rsPD.Fields("a1803")
   '受款地區國別
   m_NationName = rsPD.Fields("na03")
   '受款地區國別代號
   m_AccNA01 = "" & rsPD.Fields("na01")
   If Len(m_AccNA01) > 3 Then m_AccNA01 = Mid(m_AccNA01, 1, 3)
   '受款地區
   m_AccNA02 = "" & rsPD.Fields("na02")
   '匯款方式
   'Added by Lydia 2015/04/17 +匯款方式3:台銀電匯紙本
   'Modified by Lydia 2017/09/14 + 4.華銀電匯紙本
   'Modified by Lydia 2017/10/02 + 5.台銀合併結匯
   If InStr("2,3,4,5", "" & rsPD.Fields("a1811")) > 0 Then
      m_A1811 = "2" '電匯
   Else
      m_A1811 = "1" '票匯
   End If
   '金額前印幣別
   m_Amount = "" & rsPD.Fields("a1903") & Format(Val("" & rsPD.Fields("Amount")), FDollar)
   '中間銀行
   m_MiddleBank_01 = ""
   m_MiddleBank_02 = ""
   m_MiddleBank_03 = ""
   strA2222 = "" 'Added by Lydia 2015/03/30 媒體備註
   m_A2219 = "" 'Added by Lydia 2015/04/30
   'Added by Lydia 2025/07/31
   m_BankCity = ""
   m_BankNA = ""
   m_BankMidCode = ""
   m_RecCity = ""
   m_RecNA = ""
   'end 2025/07/31
   
   'Modified by Lydia 2016/11/22 抓受款銀行國別m_NationName
   'Modified by Lydia 2017/10/03 華銀整批媒體RMB改CNY
   'Modified by Lydia 2025/07/31 增加抓受款人地址城市、受款人地址國別
   'strQ1 = "Select a.*,na03,na02 From ACC220 a ,nation Where a2201='" & rsPD.Fields("a1803") & "' And a2202='" & IIf(rsPD.Fields("a1903") = J_RMB, "RMB", rsPD.Fields("a1903")) & "' and a2217=na01(+) "
   strQ1 = "Select a.*,n1.na03,n1.na02,n1.na60,n2.na60 as recna60 From ACC220 a ,nation n1, nation n2 " & _
           "Where a2201='" & rsPD.Fields("a1803") & "' And a2202='" & IIf(rsPD.Fields("a1903") = J_RMB, "RMB", rsPD.Fields("a1903")) & "' and a2217=n1.na01(+) and a2225=n2.na01(+) "
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strQ1)
   If intQ = 1 Then
      'modify by sonia 2014/8/14 取消中間銀行前的1,2,3
      If "" & rsQuery("A2214").Value <> "" Then
         m_MiddleBank_01 = rsQuery("A2214").Value
      End If
      If "" & rsQuery("A2215").Value <> "" Then
         m_MiddleBank_02 = rsQuery("A2215").Value
      End If
      If "" & rsQuery("A2216").Value <> "" Then
         m_MiddleBank_03 = rsQuery("A2216").Value
      End If
      'Added by Lydia 2015/03/30 媒體備註
      strA2222 = "" & UCase(Trim(rsQuery.Fields("a2222")))
      'Added by Lydia 2015/04/30 改用 A2219(手續費方式)
      If "" & rsQuery("A2219").Value <> "" Then
         m_A2219 = rsQuery("A2219").Value
      End If
      'Added by Lydia 2019/10/03 匯款日幣, 以OUR方式結匯的, OUR要改成全額到行
      If "" & rsPD.Fields("a1903") = "JPY" And UCase(m_A2219) = "71:OUR" Then
          m_A2219 = "全額到行" '僅供台銀承辦人員查看,不改變媒體
      End If
      'end 2019/10/03
      
      'Added by Lydia 2016/11/22 華南抓受款銀行國別
      'Modified by Lydia 2017/03/23 統一抓受款銀行國別
      If "" & rsQuery("na03").Value <> "" Then
         m_NationName = "" & rsQuery("na03").Value
      End If
      'Added by Lydia 2017/07/19 +國家地區
      If "" & rsQuery("na02").Value <> "" Then
         m_AccNA02 = "" & rsQuery.Fields("na02")
      End If
      'end 2017/07/19
      'Added by Lydia 2025/07/31
      m_BankNA = "" & rsQuery.Fields("na60")
      m_RecCity = "" & rsQuery.Fields("a2224")
      m_RecNA = "" & rsQuery.Fields("recna60")
      strExc(1) = m_MiddleBank_01
      strQ1 = UCase(m_MiddleBank_01)
      If InStr(strQ1, "SWIFT") > 0 Then
         m_BankMidCode = Trim(Replace(Replace(Replace(strQ1, "SWIFT", ""), "SWIFT CODE", ""), ":", ""))
      End If
      strQ1 = UCase(m_MiddleBank_02)
      If InStr(strQ1, "SWIFT") > 0 Then
         m_BankMidCode = Trim(Replace(Replace(Replace(strQ1, "SWIFT", ""), "SWIFT CODE", ""), ":", ""))
      End If
      strQ1 = UCase(m_MiddleBank_03)
      If InStr(strQ1, "SWIFT") > 0 Then
         m_BankMidCode = Trim(Replace(Replace(Replace(strQ1, "SWIFT", ""), "SWIFT CODE", ""), ":", ""))
      End If
      'end 2025/07/31
   End If

   '受款人:
   m_BeneBankName1 = ""
   m_BeneBankName2 = ""
   m_BeneBankName3 = ""
   m_AccountNum = ""
   m_Payee_01 = ""
   m_Payee_02 = ""
   m_Payee_03 = ""
   m_Payee_04 = ""
   m_FAAddr = ""
   m_A1810 = ""
   bolPayee71OUR = False  'added by Lydia 2015/06/18
   'Added by Lydia 2017/09/14 智權-公司資料
   'Modified by Lydia 2023/07/26 含華銀電匯紙本a1811=4 ;ex.W11201531的收據公司別=2,但是要用華銀電匯紙本
   If pA0K11 = "J" Or "" & rsPD.Fields("a1811") = "4" Then
      m_A0k11Chi = Trim("" & rsPD.Fields("a0802"))
      m_A0k11Eng = Trim("" & rsPD.Fields("a0803"))
      m_A0k11Id = Trim("" & rsPD.Fields("a0807"))
      m_A0k11Tel = Trim("" & rsPD.Fields("a0813")) & "#" & J_代辦人員分機  '電話#分機
      m_A0k11AddrC = Trim(Replace("" & Trim("" & rsPD.Fields("a0804")), "朱園里7鄰", ""))
      m_A0k11AddrE = Trim("" & rsPD.Fields("a0822")) & " " & Trim("" & rsPD.Fields("a0823"))
      'end 2017/09/14
      'Added by Lydia 2021/08/27 寰華結匯水單本所地址不能出現ROC
      If InStr("Y53374,", Left(Trim("" & rsPD.Fields("a0803")), 6)) > 0 Then
          m_A0k11AddrE = Replace(m_A0k11AddrE, ", R.O.C.", "")
      End If
      'end 2021/08/27
   End If
   
   '若不為電匯
   '2006/1/18 MODIFY BY SONIA 婧瑄說其他結匯不管是否電匯都抓A1810
   'Added by Lydia 2015/04/17 +匯款方式3:台銀電匯紙本
   'Modified by Lydia 2017/09/14  +匯款方式4:華銀電匯紙本
   'Modified by Lydia 2017/10/02  +匯款方式5:台銀合併結匯
   If InStr("2,3,4,5", "" & rsPD.Fields("a1811")) = 0 Or Len(rsPD.Fields("a1803")) = 5 Then
   '2006/1/18 END
      '抓受款人相關資料
      'Modified by Lydia 2017/10/03 華銀整批媒體RMB改CNY
      strQ1 = "Select * From ACC220 Where a2201='" & rsPD.Fields("a1803") & "' And a2202='" & IIf(rsPD.Fields("a1903") = J_RMB, "RMB", rsPD.Fields("a1903")) & "' "
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, strQ1)
      If intQ = 1 Then
         '受款人名稱
         m_Payee_01 = UCase(Trim("" & rsQuery.Fields("a2203").Value))
         m_Payee_02 = UCase(Trim("" & rsQuery.Fields("a2204").Value))
         m_Payee_03 = UCase(Trim("" & rsQuery.Fields("a2205").Value))
         'Modifed by Lydia 2015/03/30 名稱改3欄,第4欄改為CNAPS
         m_Payee_04 = UCase(Trim("" & rsQuery.Fields("a2220").Value))
         'Added by Lydia 2015/03/30 媒體備註
'            strA2222 = "" & UCase(Trim(rsB.Fields("a2222")))
         'Added by Lydia 2015/06/18 +華南(J公司)71:our判斷
         If UCase(Trim("" & rsQuery.Fields("a2219"))) = UCase("71:our") Then
            bolPayee71OUR = True
         End If
         'Added by Lydia 2016/08/02 +大陸中文_ 大陸地區的匯款
         If "" & rsQuery.Fields("a2217") = "020" Or ("" & rsQuery.Fields("a2217") = "" And m_AccNA01 = "020") Then '受款銀行國籍A2217優先判斷
             If PUB_CheckStrNEC("" & rsQuery.Fields("a2203")) = True Then '受款人名稱有中文
                bolPayee71OUR = True
             End If
         End If
      Else
         '受款人名稱
         If "" & rsPD.Fields("A1810").Value <> "" Then
            m_A1810 = UCase(Trim(rsPD.Fields("A1810").Value))
            m_Payee_01 = UCase(Trim(rsPD.Fields("A1810").Value))
         Else
            m_Payee_01 = UCase(Trim("" & rsPD.Fields("fa05").Value))
            m_Payee_02 = UCase(Trim("" & rsPD.Fields("fa63").Value))
            m_Payee_03 = UCase(Trim("" & rsPD.Fields("fa64").Value))
            'Memo 2015/03/30 名稱改3欄,第4欄改為CNAPS(沒匯款銀行資料)
            m_Payee_04 = ""
         End If
      End If
   '若為電匯
   Else
      '抓受款人相關資料
      'Modified by Lydia 2017/04/21 統一抓受款銀行國別
      'Modified by Lydia 2017/07/19 +na02
      'Modified by Lydia 2017/10/03 華銀整批媒體RMB改CNY
      strQ1 = "Select a.*,na03,na02 From ACC220 a ,nation Where a2201='" & rsPD.Fields("a1803") & "' And a2202='" & IIf(rsPD.Fields("a1903") = J_RMB, "RMB", rsPD.Fields("a1903")) & "' and a2217=na01(+) "
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, strQ1)
      If intQ = 1 Then
         '受款銀行名稱
         m_BeneBankName1 = UCase(Trim("" & rsQuery.Fields("a2208").Value))
         m_BeneBankName2 = UCase(Trim("" & rsQuery.Fields("a2209").Value))
         m_BeneBankName3 = UCase(Trim("" & rsQuery.Fields("a2210").Value) & " " & Trim("" & rsQuery.Fields("a2211").Value))
         '受款人帳號
         m_AccountNum = "" & rsQuery.Fields("a2207")
         'Added by Lydia 2017/04/21 統一抓受款銀行國別
         If "" & rsQuery("na03").Value <> "" Then
            m_NationName = "" & rsQuery("na03").Value
         End If
         'Added by Lydia 2017/07/19 抓國家地區
         If "" & rsQuery("na02").Value <> "" Then
            m_AccNA02 = "" & rsQuery.Fields("na02")
         End If
'            'Added by Lydia 2015/03/30 媒體備註
'            strA2222 = "" & UCase(Trim(rsquery.Fields("a2222")))
         'Added by Lydia 2015/06/18 +華南(J公司)71:our判斷
         If UCase(Trim("" & rsQuery.Fields("a2219"))) = UCase("71:our") Then
            bolPayee71OUR = True
         End If
         'Added by Lydia 2016/08/02 +大陸中文_ 大陸地區的匯款
         If "" & rsQuery.Fields("a2217") = "020" Or ("" & rsQuery.Fields("a2217") = "" And m_AccNA01 = "020") Then '受款銀行國籍A2217優先判斷
             If PUB_CheckStrNEC("" & rsQuery.Fields("a2203")) = True Then '受款人名稱有中文
                bolPayee71OUR = True
             End If
         End If
         'Added by Morgan 2014/1/27
         '澳洲代理人的水單要印地址,名稱上移
         'modify by sonia 2014/9/3 加南非301
         'Modified by Lydia 2015/03/25 + 加拿大
         'Modified by Lydia 2015/05/12 改成依匯款行國別
         'If rspd.Fields("fa10") = "015" Or rspd.Fields("fa10") = "301" Or rspd.Fields("fa10") = "102" Then
         bolAddr = False
         'Modified by Lydia 2015/08/05 若受款行資料未設定國別,以代理人資料判斷
         'If InStr("015,301,102", rsB.Fields("a2217")) > 0 Then
         'Modified by Lydia 2017/07/19 台銀要求於南非, 加拿大, 澳洲和歐洲地區(m_na02)要印地址,若有短地址則優先列印
         'Modified by Lydia 2017/08/01 台銀要求全部都要印地址
         'If InStr("015,301,102", rsB.Fields("a2217")) > 0 Or InStr("015,301,102", rspd.Fields("fa10")) > 0 Or m_NA02 = "C20" Then
            bolAddr = True
         'end 2015/05/12
            'Modified by Lydia 2017/08/22 +第3行名稱
            'Modified by Lydia 2017/09/14 華銀版面調整,受款人名稱分3行
            If Text6.Text = "J" Then
               m_Payee_01 = Trim("" & rsQuery.Fields("a2203"))
               m_Payee_02 = Trim("" & rsQuery.Fields("a2204"))
               m_Payee_03 = Trim("" & rsQuery.Fields("a2205"))
            Else
            'end 2017/09/14
                m_Payee_01 = Trim("" & rsQuery.Fields("a2203") & " " & rsQuery.Fields("a2204")) & IIf("" & rsQuery.Fields("a2205") <> "", " " & rsQuery.Fields("a2205"), "")
            End If 'end 2017/09/14
            m_FAAddr = Trim("" & rsPD.Fields("addr"))
            'Added by Lydia 2017/07/19 短地址優先列印
            If Trim("" & rsQuery.Fields("a2218")) <> "" Then m_FAAddr = Trim("" & rsQuery.Fields("a2218"))
            m_FAAddr = Replace(m_FAAddr, "#", "") 'Added by Lydia 2017/09/18 配合華銀不接受#,預設拿掉#
            m_Payee_04 = UCase(Trim("" & rsQuery.Fields("a2220").Value))  'Added by Lydia 2017/10/16
         'Else
         ''end 2014/1/27
         '   '受款人名稱
         'If IsNull(rsB.Fields("a2203").Value) = False Then
         '   m_Payee_01 = UCase(Trim(rsB.Fields("a2203").Value))
         'End If
         'If IsNull(rsB.Fields("a2204").Value) = False Then
         '   m_Payee_02 = UCase(Trim(rsB.Fields("a2204").Value))
         'End If
         'If IsNull(rsB.Fields("a2205").Value) = False Then
         '   m_Payee_03 = UCase(Trim(rsB.Fields("a2205").Value))
         'End If
         '   'Modifed by Lydia 2015/03/30 名稱改3欄,第4欄(a2206)改為CNAPS(a2220)
'           '    If IsNull(rsB.Fields("a2206").Value) = False Then
'           '       m_Payee_04 = UCase(Trim(rsB.Fields("a2206").Value))
'           '    End If
         '   If IsNull(rsB.Fields("a2220").Value) = False Then
         '      m_Payee_04 = UCase(Trim(rsB.Fields("a2220").Value))
         '   End If
         'End If
         'end 2017/07/31
      End If
   End If   'end ---- 若為電匯

   '收據公司別
   pA0K11 = "" & rsPD.Fields("a0k11").Value '公司別
   'Modified by Lydia 2015/10/06 有代為結匯之客戶編號,列印的申請人資料為該客戶的資料
   If "" & rsPD.Fields("a1718").Value <> "" Then
      ' 以\分2行列印
      'Modified by Lydia 2015/12/04 有短地址就用短地址
      'Modified by Lydia 2017/10/03 華銀整批媒體RMB改CNY
      strQ1 = "select CU05||' '||CU88||' '||CU89||' '||CU90 as a0803,CU11 as a0807,NVL(CU16,CU17) as a0813,nvl(a2218,CU24||' '||CU25||'\'||CU26||' '||CU27||' '||CU28) as custaddr " & _
               "from customer,acc220 where cu01||cu02=" & CNULL(rsPD.Fields("a1718") & "0") & " and cu01||cu02=a2201(+) and a2202='" & IIf(rsPD.Fields("a1903") = J_RMB, "RMB", rsPD.Fields("a1903")) & "' "
   Else
      'Modified by Lydia 2020/09/03 收據公司別1,2,L都歸2公司
      strQ1 = "select * from acc080 where a0801 = '" & IIf(pA0K11 = "J", "J", "2") & "' "
   End If
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strQ1)
   If intQ = 1 Then
      strCompanyName = "" & rsQuery.Fields("a0803").Value
      If IsNull(rsQuery.Fields("a0807").Value) Then
         strCompanyNo = ""
      Else
         strCompanyNo = rsQuery.Fields("a0807").Value
      End If
        ' strAddress = ReportSum(104)
         'Modified by Lydia 2015/10/06
         'strAddress = ReportSum(104)
         If "" & rsPD.Fields("a1718").Value <> "" Then
            strAddress = "" & rsQuery.Fields("custaddr")
         Else
            'Modified by Lydia 2017/09/14 公司別-中文,英文地址
            'strAddress = ReportSum(104)
            strAddress = Replace("" & rsQuery.Fields("a0804").Value, "朱園里7鄰", "")
            strCompAddr1 = "" & rsQuery.Fields("a0822").Value
            strCompAddr2 = "" & rsQuery.Fields("a0823").Value
         End If
         'end 2015/10/06
         strPhone = "" & rsQuery.Fields("a0813").Value
   Else
      strCompanyName = ""
      strCompanyNo = ""
      strAddress = ""
      strPhone = ""
      'Added by Lydia 2017/09/06 公司別-英文地址
      strCompAddr1 = ""
      strCompAddr2 = ""
   End If
   'Added by Lydia 2017/09/06 公司別-英文地址(預設)
   If strCompAddr1 & strCompAddr2 = "" Then
      strCompAddr1 = "9F, No. 112, Sec. 2, Chang-An E. Rd.,"
      strCompAddr2 = "Taipei 104, Taiwan, R.O.C."
   End If
   'end 2017/09/07
   'Added by Lydia 2021/08/27 寰華結匯水單本所地址不能出現ROC
   If InStr("Y53374,", Left(Trim(rsPD.Fields("a1803")), 6)) > 0 Then
       strCompAddr2 = Replace(strCompAddr2, ", R.O.C.", "")
   End If
   'end 2021/08/27
   Set rsQuery = Nothing
End Sub


