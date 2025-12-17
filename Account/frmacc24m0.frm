VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc24m0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "國外結匯媒體檔產生作業"
   ClientHeight    =   4464
   ClientLeft      =   7476
   ClientTop       =   6852
   ClientWidth     =   6120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4464
   ScaleWidth      =   6120
   Begin VB.TextBox txtKind 
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
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   8
      Top             =   2130
      Width           =   375
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
      Left            =   1920
      TabIndex        =   9
      Top             =   2520
      Width           =   3495
   End
   Begin VB.TextBox TxtList 
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
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   7
      Text            =   "Y"
      Top             =   1695
      Width           =   375
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
      Left            =   1920
      MaxLength       =   9
      TabIndex        =   3
      Top             =   945
      Width           =   1572
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
      Left            =   3825
      MaxLength       =   9
      TabIndex        =   4
      Top             =   945
      Width           =   1572
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
      Left            =   3825
      MaxLength       =   9
      TabIndex        =   6
      Top             =   1320
      Width           =   1572
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
      Height          =   315
      Left            =   1920
      MaxLength       =   9
      TabIndex        =   5
      Top             =   1320
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1920
      TabIndex        =   1
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生TXT(&P)"
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
      Left            =   720
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   3000
      Width           =   4725
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3825
      TabIndex        =   2
      Top             =   585
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
      Left            =   1920
      TabIndex        =   0
      Top             =   180
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
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "媒體格式　　  (空白:全部 1.台銀 2.華銀)"
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
      Left            =   960
      TabIndex        =   24
      Top             =   2160
      Width           =   5175
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "媒體檔不含""必須紙本列印""資料"
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
      Left            =   480
      TabIndex        =   23
      Top             =   3840
      Width           =   3015
   End
   Begin VB.Label Label12 
      Caption         =   "若有電匯要改成電匯紙本,請修改frmacc2170或aacc_fun"
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   3720
      TabIndex        =   22
      Top             =   240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "結匯明細匯總表產生在C:\XLS"
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
      Left            =   480
      TabIndex        =   21
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label Label10 
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
      Left            =   960
      TabIndex        =   20
      Top             =   2550
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "是否列印清單         (Y)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   255
      Left            =   960
      TabIndex        =   19
      Top             =   1725
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "結匯日期"
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
      Height          =   255
      Left            =   960
      TabIndex        =   18
      Top             =   210
      Width           =   975
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "PS: 檔案產生在桌面的\ X銀結匯水單 (檔名:幣別+結匯日.TXT)"
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
      TabIndex        =   17
      Top             =   3600
      Width           =   6000
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
      Left            =   960
      TabIndex        =   16
      Top             =   975
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
      Left            =   3600
      TabIndex        =   15
      Top             =   930
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
      Left            =   3600
      TabIndex        =   14
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
      Left            =   3600
      TabIndex        =   13
      Top             =   1320
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
      Left            =   960
      TabIndex        =   12
      Top             =   1350
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
      Left            =   960
      TabIndex        =   11
      Top             =   630
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc24m0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2025/05/27 變更模組名稱；原本104年上線的台銀模組名稱後面_T104, 106年上線的華銀模組名稱後面_H106---114/6/4通知提前上線; 114/6/9 回復TXT格式無誤; 114/6/11以後直接上線
'Memo By Lydia 2022/02/25 Form2.0已檢查 (無需修改的物件)
             '因為銀行本身就要求TXT不可使用中文,並且列印的清單也是依據TXT的內容來產生。
'Memo by Lydia 2017/09/07 更名為:國外結匯媒體檔產生作業
'Add by Lydia 2015/02/16 台銀結匯水單媒體產生作業
Option Explicit
Dim ADO24m0 As New ADODB.Recordset
'Modified by Lydia 2015/4/17 +OR16台一編號(Y編號代理人)
'Dim mOR(1 To 15) As String '台銀資料變數
'Private Const mOR_Num = 15 '台銀資料欄位數
Dim mOR(1 To 16) As String '台銀資料變數
Private Const mOR_Num As Integer = 16 '台銀資料欄位數
'Added by Lydia 2025/05/27 台銀電匯格式114年8月上線
Dim m_T4(1 To 27) As String
Private Const m_T4Num As Integer = 27
'----------------------------------------------------
Dim strDBdate As String  '結匯日期
Dim idx As Integer  '產出檔案數
Dim excelSql As String 'Added by Lydia 2015/4/17 傳結匯清單SQL條件
'----------------------------------------
'列印用
Dim prnPrint As Printer
Dim strPrint As String
Dim strPrtOrt As Integer 'Added by Lydia 2017/09/06 系統預設印表機的紙張方向
'Modified by Lydia 2015/4/17 +OR16台一編號(Y編號代理人)
'Dim PLeft(0 To 12) As Integer
'Dim tStr(0 To 10) As String
'Dim pStr(0 To 10) As String
'Modified by Lydia 2017/09/13 改成變數
'Private Const pLMax As Integer = 13  '台銀-列印定位最大值
'Private Const pStrMax As Integer = 11  '台銀-列印欄位最大值
Dim pLmax As Integer, pStrMax As Integer
'end 2017/09/13
Dim PLeft(0 To 13) As Integer
Dim tStr(0 To 11) As String
Dim pStr(0 To 11) As String

Dim iPrint As Integer, iPage As Integer
Private Const ciTitleFontSize = 14, ciFontSize = 9
Private Const ciStartX = 500, ciStartY = 400, ciColGap = 100
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim ColH_1 As Integer, ColH_2 As Integer
Dim mCount As Integer, mTotal As Double
Dim startL2 As Integer  '第2行 起始欄
Dim pStrL3 As String '第3行中間銀行資料
'Added by Lydia 2017/09/07
Dim mHX(0 To 42) As String
Private Const mHX_Num As Integer = 42   '華銀資料欄位數
'Added by Lydia 2025/06/26 華銀114年電匯格式(2025.01.14 更新)
Dim m_H4(0 To 65) As String
Private Const m_H4Num As Integer = 65
'----------------------------------------------------
Dim PLt2(0 To 12) As Integer
Dim tlStr(0 To 10) As String
Dim plStr(0 To 10) As String
Dim m_FileName As String '整批匯出匯款申請書-範本
Dim mAppDesc As String '華銀-整批匯出匯款申請書的匯款筆數/匯款幣別及匯款金額,以;區隔
Dim tmpArr As Variant
Dim mCompany As String, mAcctNo As String  '華銀清單-匯款人,帳號

Private Sub GetPleft_T104() '列印位置
Printer.Font.Name = "新細明體"
Printer.Font.Size = ciFontSize
Printer.Font.Bold = False
Printer.Font.Underline = False

Erase PLeft

   PLeft(0) = ciStartX 'OR01編號
   PLeft(1) = PLeft(0) + Printer.TextWidth(String(2, "　")) + ciColGap
'Modified by Lydia 2015/03/25 不印:匯款申請人,匯款分類 ; 印+手續費
'   PLeft(2) = PLeft(1) + Printer.TextWidth(String(22, "　")) + ciColGap 'OR03+OR04匯款申請人名稱地址
'   PLeft(3) = PLeft(2) + Printer.TextWidth(String(5, "　")) + ciColGap 'OR05匯款人證號
'   PLeft(4) = PLeft(3) + Printer.TextWidth(String(2, "　")) + ciColGap 'OR06國別
'   PLeft(5) = PLeft(4) + Printer.TextWidth(String(2, "　")) + ciColGap 'OR07匯款分類
'   PLeft(6) = PLeft(5) + Printer.TextWidth(String(1, "　")) + ciColGap 'OR08 Swift code
'
'   PLeft(7) = PLeft(6) + Printer.TextWidth(String(17, "　")) + ciColGap 'OR09
'   PLeft(8) = PLeft(7) + Printer.TextWidth(String(14, "　")) + ciColGap 'OR11+OR12收款人名稱地址
'   PLeft(9) = PLeft(8) + Printer.TextWidth(String(7, "　")) + ciColGap  'OR13收款人帳號
'   PLeft(10) = PLeft(9) + Printer.TextWidth(String(6, "　")) + ciColGap 'OR14匯款金額
'   PLeft(11) = PLeft(10) + Printer.TextWidth(String(4, "　")) + ciColGap '簽名確認

   PLeft(2) = PLeft(1) + Printer.TextWidth(String(5, "　")) + ciColGap 'OR05匯款人證號
   PLeft(3) = PLeft(2) + Printer.TextWidth(String(2, "　")) + ciColGap 'OR06國別
   PLeft(4) = PLeft(3) + Printer.TextWidth(String(1, "　")) + ciColGap 'OR08 Swift code
   
   PLeft(5) = PLeft(4) + Printer.TextWidth(String(40, "　")) + ciColGap 'OR09 swift code
  ' PLeft(6) = PLeft(5) + Printer.TextWidth(String(14, "　")) + ciColGap 'OR11+OR12收款人名稱地址
   PLeft(6) = PLeft(5) + Printer.TextWidth(String(18, "　")) + ciColGap  'OR13收款人帳號
   PLeft(7) = PLeft(6) + Printer.TextWidth(String(7, "　")) + ciColGap 'OR14匯款金額
   PLeft(8) = PLeft(7) + Printer.TextWidth(String(4, "　")) + ciColGap 'A2219手續費
   PLeft(9) = PLeft(8) + Printer.TextWidth(String(4, "　")) + ciColGap '簽名確認
   '第2行
   'Modified by Lydia 2015/4/17 +OR16台一編號(Y編號代理人)
'   PLeft(10) = PLeft(4) 'OR11 收款人名稱
'   PLeft(11) = PLeft(10) + Printer.TextWidth(String(40, "　")) + ciColGap 'OR12 收款人地址
'   PLeft(12) = lngPageWidth - ciStartX   '列印右邊界
   PLeft(10) = PLeft(1) 'OR16 台一編號
   PLeft(11) = PLeft(4) 'OR11 收款人名稱
   PLeft(12) = PLeft(10) + Printer.TextWidth(String(40, "　")) + ciColGap 'OR12 收款人地址
   PLeft(13) = lngPageWidth - ciStartX   '列印右邊界
   'Added by Lydia 2017/09/13
   pLmax = 13  '台銀-列印定位最大值
   pStrMax = 11  '台銀-列印欄位最大值
End Sub

Private Sub SetColumnName_T104() '列印位置
   tStr(0) = "編號"
   'Modified by Lydia 2015/03/25 不印:匯款申請人,匯款分類 ; 印+手續費
'   tStr(1) = "匯款申請人名稱地址"
'   tStr(2) = "匯款人證號"
'   tStr(3) = "國別"
'   tStr(4) = "匯款分類"
'   tStr(5) = "　" 'OR08 Swift code
'   tStr(6) = "收款行SWIFT CODE或名稱"
'   tStr(7) = "收款人名稱地址"
'   tStr(8) = "收款人帳號"
'   tStr(9) = "匯款金額"
'   tStr(10) = "簽名確認"
  
   tStr(1) = "匯款人證號"
   tStr(2) = "國別"
   tStr(3) = "　" 'OR08 Swift code
   tStr(4) = "收款行SWIFT CODE或名稱"
   tStr(5) = "收款人帳號"
   tStr(6) = "匯款金額"
   tStr(7) = "手續費"
   'Modified by Lydia 2017/08/07 拿掉
   'tStr(8) = "簽名確認"
   tStr(8) = ""
   '第2行
   'Modified by Lydia 2015/4/17 +OR16台一編號(Y編號代理人)
    startL2 = 9
'   tStr(9) = "收款人名稱"
'   tStr(10) = "收款人地址"
   tStr(9) = "台一編號"
   'Modified by Lydia 2017/08/07 合併列印
   'tStr(10) = "收款人名稱"
   'tStr(11) = "收款人地址"
   tStr(10) = "收款人名稱 ／ 地址"
End Sub

'Modified by Lydia 2017/09/13 +類別cType
Private Sub PrintLine(Optional ByVal cType As String = "1")
   'Modified by Lydia 2015/4/17
  'Printer.Line (PLeft(0) - 50, iPrint)-(PLeft(12), iPrint)
  'Added by Lydia 2017/09/13 判斷
  If cType = "2" Then
     Printer.Line (PLt2(0) - 50, iPrint)-(PLt2(pLmax), iPrint)
  Else
  'end 2017/09/13
     Printer.Line (PLeft(0) - 50, iPrint)-(PLeft(pLmax), iPrint)
  End If 'end 2017/09/13
End Sub
Private Sub PrintColH(ByVal aX As Integer)
  Printer.Line (aX, ColH_1)-(aX, ColH_2)
End Sub
'蓋公司章
Private Sub PrintSign()
Dim x1 As Integer
  
  Printer.Font.Size = ciFontSize
  Printer.CurrentX = 13200 'PLeft(9)
  Printer.CurrentY = lngPageHeight - 3 * lngLineHeight
  Printer.Print "匯款申請人"
  
  x1 = Printer.TextWidth("匯款申請人")
  Printer.Font.Size = ciFontSize - 2
  Printer.CurrentX = 13200 + x1
  Printer.CurrentY = lngPageHeight - 3 * lngLineHeight + 20
  Printer.Print "(請蓋公司戳印)"
End Sub

'合計
Private Sub PrintSubTotal_T104()
Dim strX As String
  
  strX = "總計筆數：" & PUB_StrToStr(CheckStr(mCount), 3, True, True) & " 筆　　總計金額：共 " & Format(mTotal, "#,##0.00") & " 元"
  
  PrintLine
  PrintNewLine_T104 (0.5)
  'iPrint = iPrint + lngLineHeight / 2
  
  Printer.Font.Size = ciTitleFontSize - 2
  Printer.CurrentX = PLeft(7) - Printer.TextWidth(strX)
  Printer.CurrentY = iPrint
  Printer.Print strX
End Sub

Private Sub PrintNewLine_T104(Optional ByVal mRate As Single = 1, Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 4)
   iPrint = iPrint + lngLineHeight * mRate
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint

      iPage = iPage + 1
      Printer.NewPage
      PrintHeader_T104
   End If
End Sub

Private Sub PrintHeader_T104()
Dim PriTitle1 As String, PriTitle2 As String
Dim x1 As Integer, x2 As Integer, x3 As Integer
Dim aP1 As Integer
PriTitle1 = "臺灣銀行外匯整批匯款明細表"
PriTitle2 = "(代匯出匯款申請書)"

iPrint = ciStartY
ColH_1 = ciStartY + lngLineHeight * 2.5 '欄線
ColH_2 = ColH_1 + lngLineHeight * 2

'title line=1
Printer.Font.Size = ciFontSize
Printer.CurrentX = PLeft(0) + 500
Printer.CurrentY = iPrint: x3 = iPrint
Printer.Print ChangeTStringToWDateString(strDBdate)

x1 = 5800
Printer.Font.Size = ciTitleFontSize
Printer.Font.Bold = True
Printer.CurrentX = x1
Printer.CurrentY = iPrint
Printer.Print PriTitle1

x2 = Printer.TextWidth(PriTitle1)
Printer.Font.Size = ciTitleFontSize - 4
Printer.Font.Bold = False
Printer.CurrentX = x1 + x2 + 50
Printer.CurrentY = iPrint + 50
Printer.Print PriTitle2

Printer.Font.Size = ciFontSize
Printer.CurrentX = 14300
Printer.CurrentY = x3
Printer.Print "頁　　次：" & iPage

'title line = 2
PrintNewLine_T104
Printer.Font.Size = ciTitleFontSize - 4
Printer.CurrentX = 6500
Printer.CurrentY = iPrint + 100
Printer.Print "幣　　別：" & pStr(0)

Printer.Font.Size = ciFontSize
PrintNewLine_T104 (1.5)
'iPrint = iPrint + lngLineHeight / 2
PrintLine
'title
PrintNewLine_T104 (0.5)
'iPrint = iPrint + lngLineHeight / 2
'Modified by Lydia 2015/4/17 +OR16台一編號(Y編號代理人)
'For aP1 = 0 To 10
For aP1 = 0 To pStrMax
  If aP1 = 0 Then '畫欄線
    PrintColH (PLeft(aP1) - 60)
  Else
    If aP1 < startL2 Then PrintColH (PLeft(aP1) - 30)
  End If
   x1 = iPrint - 40
  If aP1 >= startL2 Then     '第2行
    x1 = x1 + Printer.TextHeight("匯款") + 20
    Printer.CurrentX = PLeft(aP1 + 1) '跳過簽名確認
    Printer.CurrentY = x1
    Printer.Print tStr(aP1)
  Else
    Printer.CurrentX = PLeft(aP1)
    Printer.CurrentY = x1
    Printer.Print tStr(aP1)
  End If
Next aP1
'Modified by Lydia 2015/4/17
'PrintColH (PLeft(12))
PrintColH (PLeft(pLmax))
PrintNewLine_T104 (1.5)
'iPrint = iPrint + lngLineHeight * 1.5
PrintLine
PrintNewLine_T104 (0.5)
'iPrint = iPrint + lngLineHeight / 2
End Sub

Private Sub PrintDetail_T104()
Dim aP1 As Integer, pB As String

'Modified by Lydia 2015/4/17 +OR16台一編號(Y編號代理人)
'For aP1 = 1 To 10
For aP1 = 1 To pStrMax
    If aP1 = 7 Then '金額類-置右
        Printer.CurrentX = PLeft(aP1) - Printer.TextWidth(pStr(aP1)) - ciColGap
        Printer.CurrentY = iPrint
    Else
        If aP1 >= startL2 Then '第2行
            If aP1 = startL2 Then PrintNewLine_T104
            'Modified by Lydia 2017/08/07 合併列印
            Printer.CurrentX = PLeft(aP1 + 1) '跳過簽名確認
            Printer.CurrentY = iPrint
        Else
            Printer.CurrentX = PLeft(aP1 - 1)
            Printer.CurrentY = iPrint
        End If
    End If
    
    'Added by Lydia 2017/09/14 判斷名稱+地址
    If aP1 = startL2 + 1 Then
       '列印寬度超過可印範圍，分2行
       If PLeft(aP1 + 1) + Printer.TextWidth(pStr(aP1)) > PLeft(pLmax) Then
          Printer.Print Mid(pStr(aP1), 1, InStr(pStr(aP1), "／"))
          PrintNewLine_T104
          Printer.CurrentX = PLeft(aP1 + 1)
          Printer.CurrentY = iPrint
          Printer.Print Mid(pStr(aP1), InStr(pStr(aP1), "／") + 1)
       Else
          Printer.Print pStr(aP1)
       End If
    Else
    'end 2017/09/14
       Printer.Print pStr(aP1)
    End If 'end 2017/09/14

Next aP1

PrintNewLine_T104
'第3行
If Len(pStrL3) > 0 Then
    Printer.CurrentX = PLeft(1)
    Printer.CurrentY = iPrint
    Printer.Print pStrL3
    PrintNewLine_T104
End If
PrintNewLine_T104 '空一行
    
End Sub

Private Sub Command2_Click()

   If FormCheck = False Then
      MsgBox "請輸入相關必要條件！", , MsgText(5)
      Exit Sub
   End If
   If Len(Replace(ChangeTStringToWString(FCDate(MaskEdBox3.Text)), "/", "")) <> 8 Then
      MsgBox "請輸入結匯日期！", , MsgText(5)
      Exit Sub
   End If
   If MaskEdBox1.Text > MaskEdBox2.Text Then
      MsgBox "付款日期起不可大於付款日期止！", , MsgText(5)
      Exit Sub
   End If
   If Text3.Text <> "" And Text4.Text <> "" And Text4.Text > Text3.Text Then
      MsgBox "付款單號起不可大於付款單號止！", , MsgText(5)
      Exit Sub
   End If
   If Text1.Text <> "" And Text2.Text <> "" And Text1.Text > Text2.Text Then
      MsgBox "代理人範圍起不可大於代理人範圍止！", , MsgText(5)
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   
   PUB_RestorePrinter Combo1
   
   strExc(10) = "N" 'Added by Lydia 2025/06/25
   If txtKind = "" Or txtKind = "1" Then 'Added by Lydia 2017/09/06 判斷
      If strSrvDate(1) <= "20250613" Then 'Added by Lydia 2025/05/27 114/6/4通知提前上線; 114/6/9 回復TXT格式無誤; 114/6/11以後直接上線
         strExc(10) = "Y" 'Added by Lydia 2025/06/25
         PStoreData_T104  '明細資料存暫存檔
         PStoreRead_T104  '轉出txt媒體檔和列印
      'Added by Lydia 2025/05/27
      Else
         'Added by Lydia 2025/06/25 台銀為要產生紙本收據，需要舊版沒有OR16的TXT
         PStoreData_T104
         PStoreRead_T104 "1"
         'end 2025/06/25
         PStoreData_T114
         PStoreRead_T114  '轉出txt媒體檔和列印
      End If
      'end 2025/05/27
      
      Set ADO24m0 = Nothing 'Added by Lydia 2017/09/07
      'Added by Lydia 2015/4/17 傳結匯清單SQL條件
      Call PUB_ExcelSave2(Me.Name, excelSql) '結匯明細彙總表
   End If 'end 2017/09/06
   
   'Added by Lydia 2017/09/06 新增華銀媒體
   If txtKind = "" Or txtKind = "2" Then
      'Added by Lydia 2025/06/26 華銀格式也從MT格式(不需城市+國家)改成MX格式;6/27 經過與華銀溝通，8/1再上線
      If strSrvDate(1) >= "20250701" Then    'Memo by Lydia 2025/07/29 誤寫為7/1，溝通後7/29可以先上線
         PStoreData_H114
         PStoreRead_H114
      Else
      'end 2025/06/26
         PStoreData_H106
         PStoreRead_H106
      End If
      
      Set ADO24m0 = Nothing 'Added by Lydia 2017/09/07
      Call PUB_ExcelSave2(Me.Name, excelSql, "J") '結匯明細彙總表
   End If
   'end 2017/09/06
   
   'Modified by Lydia 2017/09/06 改用模組
   'For Each prnPrint In Printers
   '   If prnPrint.DeviceName = strPrint Then
   '      Set Printer = prnPrint
   '   End If
   'Next
   PUB_RestorePrinter strPrint, strPrtOrt
   
   Screen.MousePointer = vbDefault

   FormClear
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
'Memo by Lydia 2022/10/05 若 PStoreData_T104, PStoreData_H106 結匯規則有變更，請加註文件：\\LINUX\PolyCOM\TaieNew\電腦中心日常工作\結匯-預設匯款方式(a1811和媒體檔).doc

   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   
   'Modified by Lydia 2017/09/05 表單初始化
   PUB_InitForm Me, 6200, 4875, strBackPicPath4
   'end 2017/09/07
   
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   MaskEdBox3.Mask = DFormat
   'Added by Lydia 2017/03/20 預設系統日
   'Modified by Lydia 2017/09/08 預設為最近的周二
   'MaskEdBox3.Text = Format(strSrvDate(2), DFormat)
   'Modified by Morgan 2024/3/5 改預設星期三--思閔
    'strExc(1) = CompWorkDay(1, strSrvDate(1), , "星期二")
    strExc(1) = CompWorkDay(1, strSrvDate(1), , "星期三")
    MaskEdBox3.Text = Format(TransDate(strExc(1), 1), DFormat)
   'end 2017/09/07
   
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   
   strPrtOrt = Printer.Orientation 'Added by Lydia 2017/09/06
   PUB_SetPrinter Me.Name, Combo1, strPrint
    
   'Added by Lydia 2017/09/08
   txtKind.Text = "1" '預設台銀
   '下載華銀範本
   m_FileName = "$$華銀整批匯出匯款申請書.doc"
   If Dir(App.path & "\" & m_FileName) <> "" Then
      Kill App.path & "\" & m_FileName
   End If
   Call PUB_GetSampleFile(m_FileName, "M31-000006-0-00")
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo ErrHand

   If Not g_WordAp Is Nothing Then
      g_WordAp.Quit
CloseWord:
      Set g_WordAp = Nothing
   End If
    
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   
   Set Frmacc24m0 = Nothing
   Set ADO24m0 = Nothing
   
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   
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
   MaskEdBox3.Mask = ""
   MaskEdBox3.Text = ""
   MaskEdBox3.Mask = DFormat
   'Added by Lydia 2017/03/20 預設系統日
   'Modified by Lydia 2017/09/08 預設為最近的周二
   'MaskEdBox3.Text = Format(strSrvDate(2), DFormat)
   strExc(1) = CompWorkDay(1, strSrvDate(1), , "星期二")
   MaskEdBox3.Text = Format(TransDate(strExc(1), 1), DFormat)
   'end 2017/09/08
   
   Text1 = ""
   Text2 = ""
   Text3 = ""
   Text4 = ""
   TxtList = "Y"
End Sub

Private Sub MaskEdBox1_LostFocus()
   If MaskEdBox2.Text = MsgText(29) Then
      MaskEdBox2.Text = MaskEdBox1.Text
   End If
End Sub
Private Sub MaskEdBox3_LostFocus()
Dim tempStr As String
tempStr = Replace(ChangeTStringToWString(FCDate(MaskEdBox3.Text)), "/", "")
If tempStr < strSrvDate(1) Then
    MsgBox "結匯日期不可小於系統日!!", vbCritical
    MaskEdBox3.SetFocus
Else
    If Len(tempStr) = 8 Then
        If ChkWorkDay(tempStr) = False Then
           MsgBox "請輸入日期為工作日!!", vbCritical
           MaskEdBox3.SetFocus
        End If
    End If
End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme 'Added by Lydia 2017/09/29
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus()
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

If MaskEdBox1.Text <> MsgText(29) Or MaskEdBox2.Text <> MsgText(29) _
   Or MaskEdBox3.Text <> MsgText(29) Then
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

Private Sub TxtList_GotFocus()
   TextInverse TxtList
End Sub

Private Sub TxtList_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
    Select Case KeyAscii
    Case 8, 89
        '無動作
    Case Else
        KeyAscii = 0
    End Select
End Sub

'*************************************************
' 產生暫存檔---台銀104年上線
'*************************************************
Private Sub PStoreData_T104()
'Memo by Lydia 2022/10/05 若結匯規則有變更，請加註文件：\\LINUX\PolyCOM\TaieNew\電腦中心日常工作\結匯-預設匯款方式(a1811和媒體檔).doc
Dim strSql As String, StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim strSQLc As String, strA0K11 As String
Dim ii As Integer
Dim bolUnit As Boolean '獨立水單資料不合併計算
Dim strNA60 As String 'Added by Lydia 2015/03/25 受款地國別(台銀代號)
Dim strNA01 As String 'Added by Lydia 2017/06/13 國家代號(受款行>代理人或客戶)
Dim strNA02 As String 'Added by Lydia 2017/07/19 國家地區(受款行>代理人或客戶)
Dim strOR1112 As String 'Added by Lydia 2017/08/07 記錄列印用的名稱+地址
Dim strA2220 As String 'Added by Lydia 2017/09/20 記錄CNAPS
Dim strA1901List As String 'Added by Lydia 2017/09/22 記錄付款單號
Dim strA2222 As String 'Added by Lydia 2020/01/06 記錄媒體備註

On Error GoTo ErrHandle 'Added by Lydia 2017/09/22

   adoTaie.Execute " delete from accrpt24m0 where UNO='" & strUserNum & "' "

   '結匯日期
   'Modified by Lydia 2017/09/08 台銀直接用系統日
   'strDBdate = Replace(FCDate(MaskEdBox3.Text), "/", "")
   strDBdate = strSrvDate(2)
   
   '付款日期
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a1802 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a1802 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   '代理人
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and a1803 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a1803 <= '" & Text2 & "'"
   End If
   '付款單號
   If Text4 <> MsgText(601) Then
      strSql = strSql & " and a1801 >= '" & Text4 & "'"
   End If
   If Text3 <> MsgText(601) Then
      strSql = strSql & " and a1801 <= '" & Text3 & "'"
   End If
   '公司別
   strSql = strSql & " and a1917<>'J'" '不含Ｊ公司
   
   strSql = strSql & " and a1811<>'6' " 'Added by Lydia 2024/09/03 排除匯款方式6-抵帳
   
   excelSql = strSql 'Added by Lydia 2015/4/17 傳結匯清單SQL條件 (全電匯+票匯)
   
   '限電匯
   'Modified by Lydia 2017/09/22 + 5-台銀合併結匯
   'strSql = strSql & " and a1811=2"
   strSql = strSql & " and (a1811=2 or a1811=5) "
   
   '台銀只接受英數字資料,排除020大陸
   'Mark by Lydia 2017/09/06 轉國外付款單已經將中文名稱改為3.台銀電匯紙本
   'strSql = strSql & " and na01<>'020'"
   
' 明細
    '與excelsav2 可能有差異,例如:104/03/05~104/03/06 accrpt218 有W10400499,無W10400503
    ADO24m0.CursorLocation = adUseClient
    'Modified by Lydia 2017/03/20 增加A2219(Acc220),凡手續費為71:OUR集中在幣別群組的前方
    'Modified by Lydia 2017/07/19 + na02
    'Modified by Lydia 2017/08/14 + 判斷是否為暫收款退費 -> sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt
    'Modified by Lydia 2017/09/05 debug 客戶地址抓錯欄位 cu04||' '||cu05||' '||cu06||' '||cu07||' '||cu08||' '||cu102 -> cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102
    'Modified by Lydia 2017/09/06 +A0822,A0823(公司-英文地址)
    'Modified by Lydia 2020/09/03 收據公司別1,2,L都歸2公司
'    strSQLc = "select a1901, a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05) As FA05, Decode(FA05, Null, Null, fa63) As FA63, Decode(FA05, Null, Null, fa64) As FA64, Decode(FA05, Null, Null, fa65) As FA65, fa10, na03, a1903, a1810, a1811, sum(a1904) as Amount, a1917 As a0k11,rtrim(max(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70)) Addr,a0803,a0807,a0813,na60, na02,A0822,A0823  " & _
'              ",decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1, a1812,sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt from acc190, acc180, fagent, nation,acc080 where a1801=a1901 and a1908 is null and A1803>'Y' AND substr(a1803,1,8)=fa01 and substr(a1803,9,1)=fa02 and fa10=na01(+) and a1810 is null and a0801=a1917" & strSql & " group by a1803, a1901, Decode(fa05, Null, Nvl(FA04, FA06), FA05), Decode(FA05, Null, Null, fa63)," & _
'              "Decode(FA05, Null, Null, fa64), Decode(FA05, Null, Null, fa65), fa10, na03, a1903, a1810, a1811, a1917,a0803,a0807,a0813,na60, na02,A0822,A0823 ,decode(a1812,'Y','1',decode(a1903,'USD','2','1')),a1812"
'    strSQLc = strSQLc & " Union select a1901, a1803, substr(a1810, 1, 30) as fa05, substr(a1810, 31, 30) as fa63, substr(a1810, 61, 30) as fa64, substr(a1810, 91, 30) as fa65, '' as fa10, na03, a1903, a1810, a1811, sum(a1904) as Amount, a1917 As a0k11,rtrim(max(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70)) Addr,a0803,a0807,a0813,na60, na02,A0822,A0823  " & _
'              ",decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1, a1812,sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt from acc190, acc180, fagent, nation,acc080 where a1801=a1901 and a1908 is null and substr(a1803,1,8)=fa01(+) and A1803>'Y' AND substr(a1803,9,1)=fa02(+) and fa10=na01(+) and a1810 is not null and a0801=a1917" & strSql & " group by a1803,a1901,na03,a1903,a1810,a1811,a1917,a0803,a0807,a0813,na60, na02,A0822,A0823 " & _
'              ",decode(a1812,'Y','1',decode(a1903,'USD','2','1')),a1812"
'    strSQLc = strSQLc & " Union select a1901, a1803, Decode(CU05, Null, Nvl(CU04, CU06), CU05) As FA05, Decode(CU05, Null, Null, CU88) As FA63, Decode(CU05, Null, Null, CU89) As FA64, Decode(CU05, Null, Null, CU90) As FA65, CU10 AS FA10, na03, a1903, a1810, a1811, sum(a1904) as Amount, a1917 As a0k11,rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)) Addr,a0803,a0807,a0813,na60, na02,A0822,A0823  " & _
'              ",decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1, a1812,sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt from acc190, acc180, CUSTOMER, nation,acc080 where a1801=a1901 and a1908 is null and A1803<'Y' AND substr(a1803,1,8)=CU01 and substr(a1803,9,1)=CU02 and CU10=na01(+) and a1810 is null and a0801=a1917" & strSql & " group by a1803, a1901, Decode(CU05, Null, Nvl(CU04, CU06), CU05), Decode(CU05, Null, Null, CU88), Decode(CU05, Null, Null, CU89)," & _
'              "Decode(CU05, Null, Null, CU90), CU10, na03, a1903, a1810, a1811, a1917,a0803,a0807,a0813,na60, na02,A0822,A0823 ,decode(a1812,'Y','1',decode(a1903,'USD','2','1')),a1812"
'    strSQLc = strSQLc & " Union select a1901, a1803, substr(a1810, 1, 30) as fa05, substr(a1810, 31, 30) as fa63, substr(a1810, 61, 30) as fa64, substr(a1810, 91, 30) as fa65, '' as fa10, na03, a1903, a1810, a1811, sum(a1904) as Amount, a1917 As a0k11,rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)) Addr,a0803,a0807,a0813,na60, na02,A0822,A0823 " & _
'              ",decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1, a1812,sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt from acc190, acc180, CUSTOMER, nation,acc080 where a1801=a1901 and a1908 is null and A1803<'Y' AND substr(a1803,1,8)=CU01(+) and substr(a1803,9,1)=CU02(+) and CU10=na01(+) and a1810 is not null and a0801=a1917" & strSql & " group by a1803,a1901,na03,a1903,a1810,a1811,a1917,a0803,a0807,a0813,na60, na02,A0822,A0823 " & _
'              ",decode(a1812,'Y','1',decode(a1903,'USD','2','1')),a1812"
    strSQLc = "select a1901, a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05) As FA05, Decode(FA05, Null, Null, fa63) As FA63, Decode(FA05, Null, Null, fa64) As FA64, Decode(FA05, Null, Null, fa65) As FA65, fa10, na03, a1903, a1810, a1811, sum(a1904) as Amount, decode(a1917,'J','J','2') As a0k11,rtrim(max(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70)) Addr,a0803,a0807,a0813,na60, na02,A0822,A0823  " & _
              ",decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1, a1812,sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt from acc190, acc180, fagent, nation,acc080 where a1801=a1901 and a1908 is null and A1803>'Y' AND substr(a1803,1,8)=fa01 and substr(a1803,9,1)=fa02 and fa10=na01(+) and a1810 is null and a0801=decode(a1917,'J','J','2') " & strSql & " group by a1803, a1901, Decode(fa05, Null, Nvl(FA04, FA06), FA05), Decode(FA05, Null, Null, fa63)," & _
              "Decode(FA05, Null, Null, fa64), Decode(FA05, Null, Null, fa65), fa10, na03, a1903, a1810, a1811, decode(a1917,'J','J','2') ,a0803,a0807,a0813,na60, na02,A0822,A0823 ,decode(a1812,'Y','1',decode(a1903,'USD','2','1')),a1812"
    strSQLc = strSQLc & " Union select a1901, a1803, substr(a1810, 1, 30) as fa05, substr(a1810, 31, 30) as fa63, substr(a1810, 61, 30) as fa64, substr(a1810, 91, 30) as fa65, '' as fa10, na03, a1903, a1810, a1811, sum(a1904) as Amount, decode(a1917,'J','J','2') As a0k11,rtrim(max(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70)) Addr,a0803,a0807,a0813,na60, na02,A0822,A0823  " & _
              ",decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1, a1812,sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt from acc190, acc180, fagent, nation,acc080 where a1801=a1901 and a1908 is null and substr(a1803,1,8)=fa01(+) and A1803>'Y' AND substr(a1803,9,1)=fa02(+) and fa10=na01(+) and a1810 is not null and a0801=decode(a1917,'J','J','2') " & strSql & " group by a1803,a1901,na03,a1903,a1810,a1811,decode(a1917,'J','J','2'),a0803,a0807,a0813,na60, na02,A0822,A0823 " & _
              ",decode(a1812,'Y','1',decode(a1903,'USD','2','1')),a1812"
    strSQLc = strSQLc & " Union select a1901, a1803, Decode(CU05, Null, Nvl(CU04, CU06), CU05) As FA05, Decode(CU05, Null, Null, CU88) As FA63, Decode(CU05, Null, Null, CU89) As FA64, Decode(CU05, Null, Null, CU90) As FA65, CU10 AS FA10, na03, a1903, a1810, a1811, sum(a1904) as Amount, decode(a1917,'J','J','2') As a0k11,rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)) Addr,a0803,a0807,a0813,na60, na02,A0822,A0823  " & _
              ",decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1, a1812,sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt from acc190, acc180, CUSTOMER, nation,acc080 where a1801=a1901 and a1908 is null and A1803<'Y' AND substr(a1803,1,8)=CU01 and substr(a1803,9,1)=CU02 and CU10=na01(+) and a1810 is null and a0801=decode(a1917,'J','J','2')" & strSql & " group by a1803, a1901, Decode(CU05, Null, Nvl(CU04, CU06), CU05), Decode(CU05, Null, Null, CU88), Decode(CU05, Null, Null, CU89)," & _
              "Decode(CU05, Null, Null, CU90), CU10, na03, a1903, a1810, a1811, decode(a1917,'J','J','2'),a0803,a0807,a0813,na60, na02,A0822,A0823 ,decode(a1812,'Y','1',decode(a1903,'USD','2','1')),a1812"
    strSQLc = strSQLc & " Union select a1901, a1803, substr(a1810, 1, 30) as fa05, substr(a1810, 31, 30) as fa63, substr(a1810, 61, 30) as fa64, substr(a1810, 91, 30) as fa65, '' as fa10, na03, a1903, a1810, a1811, sum(a1904) as Amount, decode(a1917,'J','J','2') As a0k11,rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)) Addr,a0803,a0807,a0813,na60, na02,A0822,A0823 " & _
              ",decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1, a1812,sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt from acc190, acc180, CUSTOMER, nation,acc080 where a1801=a1901 and a1908 is null and A1803<'Y' AND substr(a1803,1,8)=CU01(+) and substr(a1803,9,1)=CU02(+) and CU10=na01(+) and a1810 is not null and a0801=decode(a1917,'J','J','2')" & strSql & " group by a1803,a1901,na03,a1903,a1810,a1811,decode(a1917,'J','J','2'),a0803,a0807,a0813,na60, na02,A0822,A0823 " & _
              ",decode(a1812,'Y','1',decode(a1903,'USD','2','1')),a1812"
    'end 2020/09/03
    'Modified by Lydia 2017/03/20 增加A2219(Acc220),凡手續費為71:OUR集中在幣別群組的前方
   'strSQLc = strSQLc & " Order By a1903,flag1,a0k11,a1803,a1901 "
    strSQLc = "select X.*,DECODE(A2219,'71:OUR',0,1) PKIND from (" & strSQLc & ") X,acc220 where a1803=a2201(+) and a1903=a2202(+) "
    'Modified by Lydia 2017/05/08 USD有新台幣結構(flag1)
    'strSQLc = strSQLc & " Order By a1903,pkind,flag1,a0k11,a1803,a1901 "
    '2017.9.15排序:幣別、獨立水單以台幣結匯(排最後面)、匯款方式(71:OUR排最前面)、是否為台銀電匯(程式沒這段判斷??)、收據公司別、代理人
    'Modified by Lydia 2019/02/22 改為 幣別、獨立水單以台幣結匯(排後面)、匯款方式(71:OUR排最前面)、代理人、收據公司別
    'strSQLc = strSQLc & " Order By a1903,flag1 desc,pkind,a0k11,a1803,a1901 "
    strSQLc = strSQLc & " Order By a1903,flag1 desc,pkind,a1803,a0k11,a1901 "
    'end 2017/03/20
   ADO24m0.Open strSQLc, adoTaie, adOpenStatic, adLockReadOnly
   If ADO24m0.RecordCount = 0 Then
      ADO24m0.Close
      If strExc(10) = "Y" Then 'Added by Lydia 2025/06/25
         MsgBox MsgText(28), , MsgText(5)
      End If
      Exit Sub
   End If
   ii = 1
  
   If rsB.State <> adStateClosed Then rsB.Close
   Set rsB = Nothing
   Do While ADO24m0.EOF = False
      For intI = 1 To mOR_Num
          mOR(intI) = ""
      Next intI
      bolUnit = False
      'Modified by Lydia 2015/03/25 受款地國別 =預設為代理人或客戶
      strNA60 = "" & ADO24m0.Fields("na60")
      'Added by Lydia 2017/06/13 國家代號(受款行>代理人或客戶)
      strNA01 = "" & ADO24m0.Fields("fa10")
      'Added by Lydia 2017/07/19 國家地區(受款行>代理人或客戶)
      strNA02 = "" & ADO24m0.Fields("na02")
      'Added by Lydia 2017/09/20
      strA2220 = ""
      strA2222 = "" 'Added by Lydia 2020/01/06
      
       '收據公司別
        strA0K11 = "" & ADO24m0.Fields("a0k11") '公司別
        '公司地址
         'Modified by Lydia 2017/09/06 改放在Acc080的英文地址1,2(A0822,A0823)
         'If stra0k11 = "9" Then
         '   mOR(4) = "7F-1, No. 112, Sec. 2, Chang-An E. Rd.,"
         'ElseIf stra0k11 = "1" Then
         '   mOR(4) = "10F, No. 112, Sec. 2, Chang-An E. Rd.,"
         'ElseIf stra0k11 = "J" Then
         '   mOR(4) = "4F, No. 110, Sec. 2, Chang-An E. Rd.,"
         'Else
         '   mOR(4) = "9F, No. 112, Sec. 2, Chang-An E. Rd.,"
         'End If
         'mOR(4) = mOR(4) & " Taipei 104, Taiwan, R.O.C." & "      " & ADO24m0.Fields("a0813") '合併電話
         If "" & ADO24m0.Fields("A0822") & ADO24m0.Fields("A0823") <> "" Then
            mOR(4) = ADO24m0.Fields("A0822") & " " & ADO24m0.Fields("A0823")
         Else
            mOR(4) = "9F, No. 112, Sec. 2, Chang-An E. Rd., Taipei 104, Taiwan, R.O.C."
         End If
         'Added by Lydia 2021/08/27 寰華結匯水單本所地址不能出現ROC
         If InStr("Y53374,", Left(Trim(ADO24m0.Fields("a1803")), 6)) > 0 Then
             mOR(4) = Replace(mOR(4), ", R.O.C.", String(8, " "))
         End If
         'end 2021/08/27
         mOR(4) = mOR(4) & "      " & ADO24m0.Fields("a0813")  '合併電話
         'end 2017/09/06
          
          '受款人:
             strExc(2) = "":    strExc(3) = "": strExc(4) = ""
             strExc(5) = "":    strExc(6) = "": strExc(7) = "": strExc(8) = ""
          'Modified by Lydia 2017/09/22 已限制匯款方式為2-電匯和5-合計電匯
          'If "" & ADO24m0.Fields("a1811") <> "2" Or Len(ADO24m0.Fields("a1803")) = 5 Then
          If InStr("2,5", ADO24m0.Fields("a1811")) = 0 Or Len(ADO24m0.Fields("a1803")) = 5 Then
             '抓受款人相關資料+A2217受款人(行)國別
             'Modified by Lydia 2017/07/19 + na02
             StrSqlB = "Select a.*,b.NA60,b.NA02 From ACC220 a,NATION b Where a.A2217=b.NA01(+) and a2201='" & ADO24m0.Fields("a1803") & "' And a2202='" & ADO24m0.Fields("a1903") & "' "
             intI = 1
             Set rsB = ClsLawReadRstMsg(intI, StrSqlB)
             If intI = 1 Then
                'Modified by Lydia 2015/03/25 受款地國別 =A2217受款人(行)國別
                If Not IsNull(rsB.Fields("NA60")) Then
                   strNA60 = "" & rsB.Fields("NA60")
                End If
                'Added by Lydia 2017/06/13 國家代號(受款行>代理人或客戶)
                If "" & rsB.Fields("a2217") <> "" Then
                   strNA01 = "" & rsB.Fields("a2217")
                End If
                'end 2017/06/13

                'Added by Lydia 2017/07/19 國家地區(受款行>代理人或客戶)
                If "" & rsB.Fields("NA02") <> "" Then
                   strNA02 = "" & rsB.Fields("NA02")
                End If
                'end 2017/07/19

                strA2220 = "" & rsB.Fields("A2220") 'Added by Lydia 2017/09/20
                strA2222 = "" & rsB.Fields("A2222") 'Added by Lydia 2020/01/06
                
                '受款人名稱
                If IsNull(rsB.Fields("a2203")) = False Then
                   strExc(5) = RepEnter2zero(UCase(Trim(rsB.Fields("a2203"))))
                End If
                If IsNull(rsB.Fields("a2204")) = False Then
                   strExc(6) = RepEnter2zero(UCase(Trim(rsB.Fields("a2204"))))
                End If
                If IsNull(rsB.Fields("a2205")) = False Then
                   strExc(7) = RepEnter2zero(UCase(Trim(rsB.Fields("a2205"))))
                End If
                If IsNull(rsB.Fields("a2206")) = False Then
                   strExc(8) = RepEnter2zero(UCase(Trim(rsB.Fields("a2206"))))
                End If
                'OR11
                'Modified by Lydia 2017/08/31 與電匯一致
                'mOR(11) = strExc(5) & " " & strExc(6) & " " & strExc(7) & " " & strExc(8)
                mOR(11) = strExc(5) & " " & strExc(6) & IIf(strExc(7) <> "", " " & strExc(7), "")
             Else
                '受款人名稱
                If "" & ADO24m0.Fields("A1810") <> "" Then
                   mOR(11) = RepEnter2zero(UCase(Trim(ADO24m0.Fields("A1810"))))
                Else
                   If IsNull(ADO24m0.Fields("fa05")) = False Then
                      strExc(5) = RepEnter2zero(UCase(Trim(ADO24m0.Fields("fa05"))))
                   End If
                   If IsNull(ADO24m0.Fields("fa63")) = False Then
                      strExc(6) = RepEnter2zero(UCase(Trim(ADO24m0.Fields("fa63"))))
                   End If
                   If IsNull(ADO24m0.Fields("fa64")) = False Then
                      strExc(7) = RepEnter2zero(UCase(Trim(ADO24m0.Fields("fa64"))))
                   End If
                   If IsNull(ADO24m0.Fields("fa65")) = False Then
                      strExc(8) = RepEnter2zero(UCase(Trim(ADO24m0.Fields("fa65"))))
                   End If
                   'OR11
                   'Modified by Lydia 2017/08/31 與電匯一致
                    'mOR(11) = strExc(5) & " " & strExc(6) & " " & strExc(7) & " " & strExc(8)
                    mOR(11) = strExc(5) & " " & strExc(6) & IIf(strExc(7) <> "", " " & strExc(7), "") & IIf(strExc(8) <> "", " " & strExc(8), "")
                End If
             End If

          '若為電匯 'Memo by Lydia 2017/09/22 含5-台銀合併結匯
          Else
             '抓受款人相關資料+A2217受款人(行)國別
             'Modified by Lydia 2017/07/19 + na02
             StrSqlB = "Select a.*,b.NA60,b.NA02 From ACC220 a,NATION b Where a.A2217=b.NA01(+) and a2201='" & ADO24m0.Fields("a1803") & "' And a2202='" & ADO24m0.Fields("a1903") & "' "
             intI = 1
             Set rsB = ClsLawReadRstMsg(intI, StrSqlB)
             If intI = 1 Then
                'Modified by Lydia 2015/03/25 受款地國別 =A2217受款人(行)國別
                If Not IsNull(rsB.Fields("NA60")) Then
                   strNA60 = "" & rsB.Fields("NA60")
                End If
                'Added by Lydia 2017/06/13 國家代號(受款行>代理人或客戶)
                If "" & rsB.Fields("a2217") <> "" Then
                   strNA01 = "" & rsB.Fields("a2217")
                End If
                'end 2017/06/13
                'Added by Lydia 2017/07/19 國家地區(受款行>代理人或客戶)
                If "" & rsB.Fields("NA02") <> "" Then
                   strNA02 = "" & rsB.Fields("NA02")
                End If
                'end 2017/07/19
                
                strA2220 = "" & rsB.Fields("A2220") 'Added by Lydia 2017/09/20
                strA2222 = "" & rsB.Fields("A2222") 'Added by Lydia 2020/01/06
                
                '受款銀行名稱
                If IsNull(rsB.Fields("a2208")) = False Then
                   strExc(2) = RepEnter2zero(UCase(Trim(rsB.Fields("a2208"))))
                End If
                If IsNull(rsB.Fields("a2209")) = False Then
                   strExc(3) = RepEnter2zero(UCase(Trim(rsB.Fields("a2209"))))
                End If
                '受款銀行帳號
                If IsNull(rsB.Fields("a2210")) = False Then
                   strExc(4) = RepEnter2zero(UCase(Trim(rsB.Fields("a2210")) & " " & Trim(rsB.Fields("a2211"))))
                End If
                
                mOR(9) = strExc(2) & " " & strExc(3) & " " & strExc(4) 'OR09
                mOR(10) = "" & UCase(Trim(rsB.Fields("a2211"))): mOR(10) = RepEnter2zero(mOR(10))
                
                '受款人帳號 OR13
                If IsNull(rsB.Fields("a2207")) = False Then
                   mOR(13) = RepEnter2zero(rsB.Fields("a2207"))
                End If
                  
                '澳洲,南非-代理人的水單要印地址,名稱上移
                'Modified by Lydia 2015/03/25 + 加拿大
                'Modified by Lydia 2017/06/13 用受款行國籍
                'If InStr("015,301,102", ADO24m0.Fields("fa10")) > 0 Then
                'Modified by Lydia 2017/07/19 台銀要求於南非, 加拿大, 澳洲和歐洲地區(m_na02)要印地址,若有短地址則優先列印
                'Modified by Lydia 2017/08/01 台銀要求全部都要印地址
                'If InStr("015,301,102", strNA01) > 0 Or strNA02 = "C20" Then
                   'Modified by Lydia 2017/08/31 +名稱3 => & IIf("" & rsB.Fields("a2205") <> "", " " & rsB.Fields("a2205"), "")
                   mOR(11) = RepEnter2zero(Trim("" & rsB.Fields("a2203") & " " & rsB.Fields("a2204"))) & IIf("" & rsB.Fields("a2205") <> "", " " & rsB.Fields("a2205"), "")
                    'Modified by Lydia 2015/03/25 客戶短地址
                    If Not IsNull(rsB.Fields("a2218")) Then
                        mOR(12) = RepEnter2zero(Trim("" & rsB.Fields("a2218"))) 'OR12
                    Else
                        '原地址
                        mOR(12) = RepEnter2zero(Trim("" & ADO24m0.Fields("addr"))) 'OR12
                    End If
                   mOR(12) = Replace(mOR(12), "#", "") 'Added by Lydia 2017/09/18 配合華銀不接受#,預設拿掉#
                'Else
                '   If IsNull(rsB.Fields("a2203")) = False Then
                '      strExc(5) = RepEnter2zero(UCase(Trim(rsB.Fields("a2203"))))
                '   End If
                '   If IsNull(rsB.Fields("a2204")) = False Then
                '      strExc(6) = RepEnter2zero(UCase(Trim(rsB.Fields("a2204"))))
                '   End If
                '   If IsNull(rsB.Fields("a2205")) = False Then
                '      strExc(7) = RepEnter2zero(UCase(Trim(rsB.Fields("a2205"))))
                '   End If
                '   If IsNull(rsB.Fields("a2206")) = False Then
                '      strExc(8) = RepEnter2zero(UCase(Trim(rsB.Fields("a2206"))))
                '   End If
                '   'OR11
                '    mOR(11) = strExc(5) & " " & strExc(6) & " " & strExc(7) & " " & strExc(8)
                'End If
                'end 2017/08/01
             End If
          End If
         'OR08
          If rsB.RecordCount = 0 Then
              mOR(8) = "D"
          Else
                If UCase(Trim(rsB.Fields("a2210"))) = "SWIFT CODE" Then
                    mOR(8) = "A"
                   'OR08=A ,OR09存A2211
                    mOR(9) = "" & UCase(Trim(rsB.Fields("a2211"))): mOR(9) = RepEnter2zero(mOR(9))
                    'Modified by Lydia 2015/03/25 OR08為A, OR10此欄空白
                    mOR(10) = ""
                Else
                    mOR(8) = "D"
                    'Added by Lydia 2018/10/12 台銀表示, 用ABA匯款的話, 需要在"OR10"(詳附件)加上FW
                    If UCase("" & rsB.Fields("a2210")) = "ABA NO." Then
                        mOR(10) = "FW" & mOR(10)
                    End If
                    'end 2018/10/12
                End If
          End If

        '金額 OR14 ~~ 靠右對齊,12字
          mOR(14) = PUB_StrToStr(Format(ADO24m0.Fields("Amount"), "###0.00"), 12, True, True)
          
         'Modified by Lydia 2018/04/23 寰華Y53374備註欄都不要再填寫(by 郭)
         'If Left(Trim(ADO24m0.Fields("a1803")), 6) <> "Y37580" Then
         'Modified by Lydia 2020/04/22 +建毅Y51566,唯源Y52404
         If InStr("Y37580,Y53374,Y51566,Y52404,", Left(Trim(ADO24m0.Fields("a1803")), 6)) = 0 Then
            StrSqlB = "Select a1706 From acc190, acc170 Where a1902=a1702 And a1901='" & ADO24m0.Fields("a1901") & "' Group By a1706 Order By 1 "
         Else
            StrSqlB = "select null from dual"
         End If
         intI = 1
         Set rsB = ClsLawReadRstMsg(intI, StrSqlB)
         If intI = 1 Then
            intI = 0
            Do While Not rsB.EOF
               intI = intI + 1
               If intI > 12 Then Exit Do
               mOR(15) = mOR(15) & "," & rsB.Fields(0).Value
               mOR(15) = mOR(15) & IIf(ii = 12 And rsB.RecordCount > 12, " etc.", "")
               rsB.MoveNext
            Loop
         End If
         If mOR(15) <> "" Then mOR(15) = Mid(mOR(15), 2)
         'Modified by Lydia 2017/08/01 票匯不加INV.
         'mOR(15) = "INV." & mOR(15) 'Added by Lydia 2017/07/20 DB note 加 INV.
         'Modified by Lydia 2017/08/14 暫收款退費不加INV. (a1902=a1702為O單號)
         'If "" & ADO24m0.Fields("a1811") <> "1" Then mOR(15) = "INV." & mOR(15) 'Remove by Lydia 2018/06/26 備註重覆INV.
         'Modified by Lydia 2020/04/22 +天津三元Y37580; +建毅Y51566,唯源Y52404
         If InStr("Y37580,Y53374,Y51566,Y52404,", Left(Trim(ADO24m0.Fields("a1803")), 6)) = 0 Then  'Added by Lydia 2018/04/23 寰華Y53374備註欄都不要再填寫(by 郭)
            If "" & ADO24m0.Fields("a1811") <> "1" And Val("" & ADO24m0.Fields("Ocnt")) = 0 Then 'Memo by Lydia 2018/06/26 非票匯和暫收款退費
              mOR(15) = "INV." & mOR(15)
            End If
         End If
         'Added by Lydia 2020/01/06 媒體備註+備註
         If strA2222 <> "" Then mOR(15) = PUB_GetSimpleName(strA2222, True, True) & " " & mOR(15)
         
         'Added by Lydia 2017/09/20 大陸RMB結匯必須輸入CNAPS
         If strA2220 <> "" And strNA01 = "020" And "" & ADO24m0.Fields("a1903") = "RMB" Then
            mOR(10) = "CNAPS:" & strA2220
         End If
         'end 2017/09/20
         
         mOR(1) = Format(ii, "0000")
         mOR(2) = " "
         mOR(3) = convForm(RepEnter2zero(CheckStr("" & ADO24m0.Fields("a0803"))), 70)
         mOR(4) = convForm(mOR(4), 70)
         mOR(5) = convForm(RepEnter2zero(CheckStr("" & ADO24m0.Fields("a0807"))), 10)
         'Modified by Lydia 2015/03/25 受款地國別
         'mOR(6) = IIf(Len(ADO24m0.Fields("na60")) = 0, "US", ADO24m0.Fields("na60"))
         mOR(6) = IIf(Len(Trim(strNA60)) = 0, "US", strNA60)
         
         'Modified by Lydia 2017/04/07 匯款類別改為192 代理費支出
         'mOR(7) = "19D" 'OR07固定為19D
         'Modified by Lydia 2017/04/21 改回19D
         'mOR(7) = "192"
         'Modified by Lydia 2017/09/14 改成常數Pub_DBtype
         'mOR(7) = "19D"
         mOR(7) = Mid(Pub_DBtype, 1, 3)
         
         '注意雙字元切字會有奇數字元,無法insert
         mOR(9) = convForm(PUB_StrToStr(mOR(9), 105), 105)
         mOR(10) = convForm(PUB_StrToStr(mOR(10), 30), 30)
         'Added by Lydia 2017/09/06 收款人名稱去掉中文(婉莘說:銀行資料的地址都已經設成英文)
         mOR(11) = PUB_GetSimpleName(mOR(11), True)
         mOR(12) = PUB_GetSimpleName(mOR(12), True) '地址(去掉中文)
         'end 2017/09/06
         
         'Added by Lydia 2017/08/07 記錄列印用的名稱+地址
         strOR1112 = Trim(mOR(11)) & " ／ " & Trim(mOR(12))
         'Added by Lydia 2017/09/22 加寬欄位提示
         If GetTextLength(strOR1112) > 200 Then
            MsgBox "受款人代號：" & ADO24m0.Fields("a1803") & " 幣別：" & ADO24m0.Fields("a1903") & vbCrLf & "名稱 ／ 地址的長度超過200，清單會去掉尾部字元!"
            strOR1112 = convForm(strOR1112, 400)
         End If
         'end 2017/09/22
         
         'Modified by Lydia 2015/03/25 若OR11長度>70, 可連續放到OR12,唯不可超過140
         'Modified by Lydia 2015/03/31 增加受款人短地址A2218,只限定受款人名稱+地址長度140字
         'Memo by Lydia 2016/05/10 因為名稱和地址各自只有70字元,所以那個超過70字元就佔另一個欄位空間,然後中間補空白到70字,可是空白所佔寬度又比一般大寫英文字小,所以在列印時可能離很遠。
         'Memo by Lydia 2017/09/08 台銀:名稱+地址 總共140個字元,TXT檔名稱及地址要分開列示, 且不可以越界(地址必須從OR12開始, OR12總共70字元), 多出來的地址尾端就切掉.
                                      '印出的紙本清單, 名稱後面斜線, 接著地址. 印出的紙本結匯申請書 , 名稱地址欄位都夠大, 所以毋須調整
         'Mark by Lydia 2017/09/08 名稱和地址都限70字元
         'If Len(mOR(11)) > 70 Then
         '   '名稱超過70字
         '   mOR(11) = PUB_StrToStr(mOR(11), 140)
         '   mOR(12) = Mid(mOR(11), 71, Len(mOR(11)) - 70) & " " & mOR(12)
         '   mOR(11) = Mid(mOR(11), 1, 70)
         '   mOR(12) = convForm(PUB_StrToStr(mOR(12), 70), 70)
         'ElseIf Len(mOR(12)) > 70 Then
         '   '地址超過70字
         '   mOR(12) = PUB_StrToStr(mOR(12), 140)
         '   strExc(10) = " " & Mid(mOR(12), 1, Len(mOR(12)) - 70)
         '   mOR(11) = convForm(PUB_StrToStr(mOR(11), 70 - Len(strExc(10)), True), 70 - Len(strExc(10))) & strExc(10)
         '   mOR(12) = Mid(mOR(12), Len(strExc(10)), 70)
         'Else
            mOR(11) = convForm(PUB_StrToStr(mOR(11), 70), 70)
            mOR(12) = convForm(PUB_StrToStr(mOR(12), 70), 70)
         'End If 'end Mark by Lydia 2017/09/08
         
         'Added by Lydia 2022/04/13 代理Y25061430名稱or11=C / Zurbano 76, 7 ° Madrid 28010 SPAIN,  經過模組只有69個字元
         If Len(mOR(11)) < 70 Then
            mOR(11) = Mid(mOR(11) & String(70, " "), 1, 70)
         End If
         If Len(mOR(12)) < 70 Then
            mOR(12) = Mid(mOR(12) & String(70, " "), 1, 70)
         End If
         'end 2022/04/13
         
         mOR(13) = convForm(PUB_StrToStr(mOR(13), 34), 34)
         mOR(15) = convForm(PUB_StrToStr(RepEnter2zero(CheckStr(mOR(15))), 140), 140)
         If Len(Trim(mOR(15))) > 135 Then
            If InStr(Mid(mOR(15), 134, 6), "etc.") > 0 Then
              mOR(15) = convForm(PUB_StrToStr(mOR(15), 140), 140)
            Else
              mOR(15) = convForm(PUB_StrToStr(mOR(15), 135) & " etc.", 140)
            End If
         End If
         
         'Modified by Lydia 2015/4/17 +OR16台一編號(Y編號代理人)
         mOR(16) = convForm(PUB_StrToStr(ADO24m0.Fields("A1803"), 20), 20)
         
         '獨立水單不合併明細
         If ADO24m0.Fields("a1812") = "Y" Then bolUnit = True
         
         '合併明細資料 'decode(substr(or04,1,2),'7F','9','10','1','4F','J','9F','2') 公司別, or mor(5) 匯款人證號
         'Modified by Lydia 2021/04/12 +判斷AR05獨立水單
         StrSqlB = "Select * From accrpt24m0  where UNO='" & strUserNum & "' and ar00='" & ADO24m0.Fields("a1803") & "' " & _
                   "And ar01='" & ADO24m0.Fields("a1903") & "' And ar03='" & ADO24m0.Fields("flag1") & "' and ar04='" & strA0K11 & "' and ar05 is null "
         strExc(1) = "":          intI = 1
         Set rsB = ClsLawReadRstMsg(intI, StrSqlB)
         If intI = 1 And bolUnit = False Then
            '串D/B note
            If Trim(rsB.Fields("OR15")) <> Trim(mOR(15)) Then
               strExc(2) = convForm(Trim(rsB.Fields("OR15")) & "," & Trim(mOR(15)), 140)
               If Len(Trim(strExc(2))) > 135 Then
                  If InStr(Mid(strExc(2), 134, 6), "etc.") > 0 Then
                  strExc(2) = convForm(PUB_StrToStr(strExc(2), 140), 140)
                Else
                  strExc(2) = convForm(PUB_StrToStr(strExc(2), 135) & " etc.", 140)
                  End If
               End If
               strExc(1) = ",OR15=" & CNULL(strExc(2))
            End If
            'Modified by Lydia 2022/05/13 debug=> and ar05 is null;
            strExc(1) = "update accrpt24m0 set OR14=OR14+" & Val(mOR(14)) & strExc(1)
            strExc(1) = strExc(1) & " where UNO='" & strUserNum & "' and ar00='" & ADO24m0.Fields("a1803") & "' And ar01='" & ADO24m0.Fields("a1903") & "' " & _
                       "And ar03='" & ADO24m0.Fields("flag1") & "' and ar04='" & strA0K11 & "' and ar05 is null "
         Else
            'PK ~ UNO(使用者),AR00(代理人編號)+AR01(幣別)+AR02(代理人名稱)+AR03(flag1)+AR04(A1917公司別)
            'Modified by Lydia 2021/04/12 +AR05獨立水單=a1812
            strExc(1) = "INSERT INTO ACCRPT24M0 values(" & CNULL(strUserNum) & "," & CNULL(ChgSQL("" & ADO24m0.Fields("a1803"))) & "," & CNULL(ChgSQL("" & ADO24m0.Fields("a1903"))) & _
                        "," & CNULL(convForm(PUB_StrToStr(ChgSQL("" & ADO24m0.Fields("fa05")), 30), 30)) & "," & CNULL(ADO24m0.Fields("flag1")) & "," & CNULL(strA0K11) & "," & CNULL("" & ADO24m0.Fields("a1812"))
            'OR01~OR13
           'Modified by Lydia 2015/4/17
           ' For intI = 1 To mOR_Num - 2
            For intI = 1 To 13
                strExc(1) = strExc(1) & "," & CNULL(ChgSQL(mOR(intI)))
            Next intI
            'OR14 金額 ,OR15
            'Modified by Lydia 2015/4/17 +OR16台一編號(Y編號代理人)
            'Modified by Lydia 2017/08/07 +OR1112名稱+地址合併列印
            strExc(1) = strExc(1) & "," & Val(mOR(14)) & "," & CNULL(ChgSQL(mOR(15))) & "," & CNULL(ChgSQL(mOR(16))) & "," & CNULL(ChgSQL(strOR1112))
            strExc(1) = strExc(1) & ")"
            ii = ii + 1
         End If
         
         adoTaie.Execute strExc(1)
         strA1901List = strA1901List & ADO24m0.Fields("a1901") & "," 'Added by Lydia 2017/09/22 記錄付款單號
         
      ADO24m0.MoveNext
   Loop
   
   ADO24m0.Close
   If rsB.State <> adStateClosed Then rsB.Close
   Set rsB = Nothing
   
'Added by Lydia 2017/09/22 更新付款單的匯款方式=>5.台銀合併結匯
   If strA1901List <> "" Then
      PUB_UpdateA1811toType strA1901List
   End If
   Exit Sub
   
ErrHandle:
   If Err.Number <> 0 Then MsgBox Err.Description
End Sub

'清除字串中的enter和連續空白
Private Function RepEnter2zero(ByVal dStr As String) As String
    dStr = PUB_StringFilter(dStr) '清除字串中的enter
    dStr = PUB_RepToOneSpace(dStr) '清除連續空白
    dStr = Replace(dStr, "&", "and") '台銀不可有＆
    dStr = Replace(dStr, "＆", "and")
    RepEnter2zero = dStr
End Function

Private Function GetFlagTitle(ByVal strA As String, ByVal Fs1 As String) As String
Dim b1 As String
    If Fs1 = "1" Then
       b1 = "(以新台幣結購)"
    Else
       b1 = "(以外匯存款提出)"
    End If
  
    If Len(Trim(strA)) > 0 Then b1 = strA & b1
    
    GetFlagTitle = b1
End Function

Private Sub PStoreRead_T104(Optional ByVal aKind As String)   '讀取明細暫存---台銀104年上線
Dim ff As Integer, strPath As String
Dim strR As String
Dim rsR As New ADODB.Recordset
Dim strFileNo As String, strFileName As String '檔名
Dim Id As Integer
Dim bolPrint As Boolean '開始列印
Dim iCall As Integer 'Added by Lydia 2018/12/11 結匯轉出的TXT檔案要產出2個檔案

    If rsR.State <> adStateClosed Then rsR.Close
    Set rsR = Nothing
    'Modified by Lydia 2015/03/25 抓手續費
'    strR = "Select * From accrpt24m0 where UNO='" & strUserNum & "' order by OR01 "
    'Modified by Lydia 2017/08/07 抓收款人名稱
    strR = "Select a.*,b.a2219,(b.a2214||' '||b.a2215||' '||b.a2216) midbk From accrpt24m0 a, acc220 b " & _
           "where ar00=a2201(+) and ar01=a2202(+) and UNO='" & strUserNum & "' order by OR01 "
    intI = 1
    Set rsR = ClsLawReadRstMsg(intI, strR)
    If intI = 1 Then
        strPath = PUB_Getdesktop
        strPath = strPath & "\台銀結匯水單"
        If Dir(strPath, vbDirectory) = "" Then
           MkDir strPath
        End If
        
StartPrt:
        strFileNo = "": pStr(0) = ""
        If bolPrint Then
           rsR.MoveFirst
           '設定印表機
            Printer.EndDoc
            Printer.Orientation = 2 '1.直印 2.橫印
            Printer.PaperSize = 9  'A4
               
            lngPageHeight = Printer.ScaleHeight
            lngPageWidth = Printer.ScaleWidth
            lngLineHeight = 270
            GetPleft_T104 '設定邊界
            SetColumnName_T104
            iPage = 0
        Else
           idx = 0
        End If
        
        'Added by Lydia 2018/12/11 台銀要求媒體TXT檔案拿掉最後面的OR16欄位(台一編號)；
        '所以改成結匯轉出的TXT檔案要產出2個檔案，一個有OR16欄位，一個沒有OR16欄位的檔案名稱在尾端多一個"X"。
        'Modified by Lydia 2025/06/25 台銀為要產生紙本收據，需要舊版沒有OR16的TXT
        'For iCall = 1 To IIf(bolPrint = False, 2, 1)
        For iCall = 1 To IIf(aKind = "1", 1, IIf(bolPrint = False, 2, 1))
            strFileNo = "": pStr(0) = ""
            rsR.MoveFirst
        'end 2018/12/11
            Do While rsR.EOF = False
                '幣別+結匯日期+結購類型
                If strFileNo = "" Or strFileNo <> rsR.Fields("ar01") & strDBdate & GetFlagTitle("", rsR.Fields("ar03")) Then
                    pStr(0) = GetFlagTitle(rsR.Fields("ar01"), rsR.Fields("ar03"))
                    strFileNo = GetFlagTitle(rsR.Fields("ar01") & strDBdate, rsR.Fields("ar03"))
                    Id = 1
                    If bolPrint Then
                         iPage = iPage + 1
                        If iPage > 1 Then '換頁
                           PrintSubTotal_T104
                           PrintSign
                           Printer.NewPage
                        End If
                        iPage = 1: mCount = 0: mTotal = 0  '以各分類為準
                        PrintHeader_T104
                    Else
                        idx = idx + 1
                        If ff > 0 Then Close #ff
                        ff = FreeFile
                        'Added by Lydia 2025/06/25 台銀為要產生紙本收據，需要舊版沒有OR16的TXT
                        If aKind = "1" Then
                           strFileName = strPath & "\" & strFileNo & "_O.txt"
                        Else
                        'end 2025/06/25
                           If iCall = 2 Then  'Added by Lydia 2018/12/11 去掉OR16
                               strFileName = strPath & "\" & strFileNo & ".txt"
                           'Added by Lydia 2018/12/11
                           Else
                               strFileName = strPath & "\" & strFileNo & "X.txt" '保留OR16
                           End If
                           'end 2018/12/11
                        End If
                        Open strFileName For Output As ff
                        
                        'Modified by Lydia 2015/4/17 +OR16台一編號(Y編號代理人)
                        'Modified by Lydia 2018/12/11 判斷是否去掉OR16
                        'strExc(0) = "OR01" & "2OR03" & convForm(CheckStr(""), 66) & "OR04" & convForm(CheckStr(""), 66) & "OR05" & convForm(CheckStr(""), 6) & "0607 8OR09" & convForm(CheckStr(""), 101) & _
                                   "OR10" & convForm(CheckStr(""), 26) & "OR11" & convForm(CheckStr(""), 66) & "OR12" & convForm(CheckStr(""), 66) & "OR13" & convForm(CheckStr(""), 30) & "OR14" & convForm(CheckStr(""), 8) & "OR15" & convForm(CheckStr(""), 136) & _
                                   "OR16" & convForm(CheckStr(""), 16)
                        strExc(0) = "OR01" & "2OR03" & convForm(CheckStr(""), 66) & "OR04" & convForm(CheckStr(""), 66) & "OR05" & convForm(CheckStr(""), 6) & "0607 8OR09" & convForm(CheckStr(""), 101) & _
                                   "OR10" & convForm(CheckStr(""), 26) & "OR11" & convForm(CheckStr(""), 66) & "OR12" & convForm(CheckStr(""), 66) & "OR13" & convForm(CheckStr(""), 30) & "OR14" & convForm(CheckStr(""), 8) & "OR15" & convForm(CheckStr(""), 136) & _
                                   IIf(iCall = 2, "", "OR16" & convForm(CheckStr(""), 16))
                        Print #ff, strExc(0)
                    End If
                End If
             
                mOR(1) = Format(Id, "0000")
                strExc(1) = mOR(1)
                For intI = 2 To mOR_Num
                   If intI = 14 Then
                      mOR(intI) = PUB_StrToStr(Format(rsR.Fields("OR14"), "###0.00"), 12, True, True)
                   Else
                      'Modified by Lydia 2021/04/12  +AR05獨立水單=a1812
                      'mOR(intI) = "" & rsR.Fields(intI + 5) '0~5屬PK
                      mOR(intI) = "" & rsR.Fields(intI + 6) '0~6屬PK
                   End If
                   'Modified by Lydia 2018/12/11 判斷是否去掉OR16
                   'strExc(1) = strExc(1) & mOR(intI)
                   If iCall <> 2 Or intI <> 16 Then
                       strExc(1) = strExc(1) & mOR(intI)
                   End If
                   'end 2018/12/11
                Next intI
    
                If bolPrint Then
                    pStr(1) = mOR(1)
                    'Modified by Lydia 2015/03/25 不印:匯款申請人,匯款分類 ; 印+手續費
    '                pStr(2) = convForm(mOR(3) & mOR(4), 40)
    '                pStr(3) = mOR(5)
    '                pStr(4) = mOR(6)
    '                pStr(5) = mOR(7)
    '                pStr(6) = mOR(8)
    '                pStr(7) = convForm(mOR(9), 30)
    '                 '注意雙字元切字會有奇數字元,無法insert
    '                pStr(8) = convForm(PUB_StrToStr(mOR(11) & mOR(12), 24), 24)
    '                pStr(9) = convForm(mOR(13), 14)
    '                pStr(10) = PUB_StrToStr(Format(Val(mOR(14)), "##,##0.00"), 12, True, True) '列印金額加千分號
    
                    pStr(2) = mOR(5)
                    pStr(3) = mOR(6)
                    pStr(4) = mOR(8)
                    pStr(5) = convForm(PUB_StrToStr(mOR(9), 70), 70)
                    '匯款人帳號及匯款金額要全印
                    pStr(6) = convForm(mOR(13), 34)
                    pStr(7) = PUB_StrToStr(Format(Val(mOR(14)), "##,##0.00"), 14, True, True)
                    '+手續費
                    'Added by Lydia 2019/10/03 匯款日幣, 以OUR方式結匯的, OUR要改成全額到行
                    If rsR.Fields("ar01") = "JPY" And UCase("" & rsR.Fields("a2219")) = "71:OUR" Then
                        pStr(8) = "全額到行" '僅供台銀承辦人員查看,不改變媒體
                    Else
                    'end 2019/10/03
                        pStr(8) = convForm("" & rsR.Fields("a2219"), 7)
                    End If
    
                    'Modified by Lydia 2015/4/17 +OR16台一編號(Y編號代理人)
                    '注意雙字元切字會有奇數字元,無法insert
    '                pStr(9) = convForm(mOR(11), 70)
    '                pStr(10) = convForm(mOR(12), 70)
                    pStr(9) = convForm(mOR(16), 20)
                    'Modified by Lydia 2017/08/07 名稱+地址合併列印
                    'pStr(10) = convForm(mOR(11), 70)
                    'pStr(11) = convForm(mOR(12), 70)
                    pStr(10) = Trim("" & rsR.Fields("OR1112"))
                    pStr(11) = ""
                    
                    '+第3行 中間銀行資料
                    If Len(Trim(rsR.Fields("midbk"))) > 0 Then
                       pStrL3 = "中間銀行：" & rsR.Fields("midbk")
                    Else
                       pStrL3 = ""
                    End If
                    mCount = mCount + 1
                    mTotal = mTotal + Val(mOR(14))
                    PrintDetail_T104
                Else
                    Print #ff, strExc(1)
                    
                End If
               Id = Id + 1
               rsR.MoveNext
            Loop
               
            If bolPrint Then GoTo EndPrt
            
            Close #ff
        Next iCall 'End 2018/12/11
        
        If TxtList <> "Y" Then
           rsR.Close
           If aKind <> "1" Then 'Added by Lydia 2025/06/25
               MsgBox "已在桌面的台銀結匯水單資料夾產生 " & idx & " 個檔案!!", vbInformation
           End If
        End If
        
        'Modified by Lydia 2025/07/21 +And aKind <> "1"
        If TxtList = "Y" And bolPrint = False And aKind <> "1" Then
           bolPrint = True
           GoTo StartPrt
        End If
    End If
'------------------------
EndPrt:
   If bolPrint Then
        PrintSubTotal_T104
        PrintSign
        Printer.EndDoc
        MsgBox "已在桌面的台銀結匯水單資料夾產生 " & idx & " 個檔案, 並列印清單完成!!", vbInformation
        rsR.Close
   End If
End Sub

'Added by Lydia 2017/09/06
Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
  If Text4 <> "" Then '預設代入
     Text3 = Text4
  End If
End Sub

'Added by Lydia 2017/09/06
Private Sub txtkind_GotFocus()
   TextInverse txtKind
End Sub

Private Sub txtKind_Validate(Cancel As Boolean)
   If txtKind <> "" And txtKind <> "1" And txtKind <> "2" Then
      MsgBox "請輸入1或2 !!"
      txtKind.SetFocus
      Cancel = True
      Exit Sub
   End If
End Sub

' Added by Lydia 2017/09/06 產生華銀媒體暫存檔
Private Sub PStoreData_H106()
'Memo by Lydia 2022/10/05 若結匯規則有變更，請加註文件：\\LINUX\PolyCOM\TaieNew\電腦中心日常工作\結匯-預設匯款方式(a1811和媒體檔).doc
Dim strSql As String, StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim strSQLc As String, strA0K11 As String
Dim ii As Integer
Dim inR As Integer '序號
Dim atJ As Integer '下一欄位序號
Dim bolUnit As Boolean '獨立水單資料不合併計算
Dim strHX1518 As String '記錄列印用的名稱+地址
Dim strA2220 As String 'CNAPS(大陸匯款)
Dim strA2219 As String '手續費方式
Dim strCurr  As String '目前幣別
Dim strA1a12  As String '目前幣別的匯率議價編號
Dim tmpStr(1 To 4) As String   '合併-付款明細
Dim strA2222 As String 'Added by Lydia 2020/05/28 記錄媒體備註
Dim bolAdd As Boolean 'Added by Lydia 2023/08/28 是否增加序號

   adoTaie.Execute " delete from accrpt24m0_2 where UNO='" & strUserNum & "' "

   '結匯日期
   strDBdate = Replace(FCDate(MaskEdBox3.Text), "/", "")
   '付款日期
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a1802 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a1802 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   '代理人
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and a1803 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a1803 <= '" & Text2 & "'"
   End If
   '付款單號
   If Text4 <> MsgText(601) Then
      strSql = strSql & " and a1801 >= '" & Text4 & "'"
   End If
   If Text3 <> MsgText(601) Then
      strSql = strSql & " and a1801 <= '" & Text3 & "'"
   End If
   '公司別
   strSql = strSql & " and a1917='J'" '限Ｊ公司
   
   strSql = strSql & " and a1811<>'6' " 'Added by Lydia 2024/09/03 排除匯款方式6-抵帳
   
   excelSql = strSql '傳結匯清單SQL條件 (全電匯+票匯)
   
   '限電匯
   strSql = strSql & " and a1811=2"
   
' 明細
    '與excelsav2 可能有差異
    ADO24m0.CursorLocation = adUseClient
    '排序:幣別、代理人、匯款方式、收據公司別
    '代理人(Y 編號)
    'Modified by Lydia 2017/09/30 華銀整批媒體RMB改CNY a1903=> DECODE(A1903,'RMB','" & J_RMB & "',A1903) A1903
    strSQLc = "select a1901, a1803, DECODE(A1903,'RMB','" & J_RMB & "',A1903) A1903, a1917 As a0k11,decode(a1810,null,Decode(fa05, Null, Nvl(FA04, FA06), FA05),substr(a1810,1,30)) As FA05, " & _
              "decode(a1810,null,Decode(FA05, Null, Null, fa63),substr(a1810,31,30)) As FA63, decode(a1810,null,Decode(FA05, Null, Null, fa64),substr(a1810,61,30)) As FA64, " & _
              "decode(a1810,null,Decode(FA05, Null, Null, fa65),substr(a1810,91,30)) As FA65, fa10, na03, na02,nvl(na80,na04) na04, a1810, a1811,a1812, sum(a1904) as Amount, " & _
              "sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt,rtrim(max(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70)) Addr,a0803,a0807,a0813,na60,A0822,A0823 " & _
              ",decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1 from acc190, acc180, fagent, nation,acc080 where a1801=a1901 and a1908 is null and A1803>'Y' " & _
              "AND substr(a1803,1,8)=fa01 and substr(a1803,9,1)=fa02 and fa10=na01(+) and a0801=a1917" & strSql & _
              " group by a1901,a1803, decode(a1903,'RMB','" & J_RMB & "',a1903), a1917, decode(a1810,null,Decode(fa05, Null, Nvl(FA04, FA06), FA05),substr(a1810,1,30))," & _
              " decode(a1810,null,Decode(FA05, Null, Null, fa63),substr(a1810,31,30)), decode(a1810,null,Decode(FA05, Null, Null, fa64),substr(a1810,61,30))," & _
              " decode(a1810,null,Decode(FA05, Null, Null, fa65),substr(a1810,91,30)), fa10, na03,na02,nvl(na80,na04), a1810, a1811,a1812,a0803,a0807,a0813,na60,A0822,A0823," & _
              " decode(a1812,'Y','1',decode(a1903,'USD','2','1'))"
    '客戶(X 編號)
    'Modified by Lydia 2017/09/30 華銀整批媒體RMB改CNY a1903=> DECODE(A1903,'RMB','" & J_RMB & "',A1903) A1903
    strSQLc = strSQLc & " Union select a1901, a1803, DECODE(A1903,'RMB','" & J_RMB & "',A1903) A1903, a1917 As a0k11,decode(a1810,null,Decode(CU05, Null, Nvl(CU04, CU06), CU05),substr(a1810,1,30)) As FA05, " & _
              "decode(a1810,null,Decode(CU05, Null, Null, CU88),substr(a1810,31,30)) As FA63, decode(a1810,null,Decode(CU05, Null, Null, CU89),substr(a1810,61,30)) As FA64, " & _
              "decode(a1810,null,Decode(CU05, Null, Null, CU90),substr(a1810,91,30)) As FA65, CU10 AS FA10, na03,na02,nvl(na80,na04) na04, a1810, a1811,a1812, sum(a1904) as Amount, " & _
              "sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt,rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)) Addr,a0803,a0807,a0813,na60,A0822,A0823 " & _
              ",decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1 from acc190, acc180, customer, nation,acc080 where a1801=a1901 and a1908 is null and A1803<'Y' " & _
              "AND substr(a1803,1,8)=CU01 and substr(a1803,9,1)=CU02 and CU10=na01(+) and a0801=a1917" & strSql & _
              " group by a1901,a1803, decode(a1903,'RMB','" & J_RMB & "',a1903), a1917,  decode(a1810,null,Decode(CU05, Null, Nvl(CU04, CU06), CU05),substr(a1810,1,30))," & _
              " decode(a1810,null,Decode(CU05, Null, Null, CU88),substr(a1810,31,30)), decode(a1810,null,Decode(CU05, Null, Null, CU89),substr(a1810,61,30))," & _
              " decode(a1810,null,Decode(CU05, Null, Null, CU90),substr(a1810,91,30)), CU10, na03,na02,nvl(na80,na04), a1810, a1811,a1812,a0803,a0807,a0813,na60,A0822,A0823," & _
              " decode(a1812,'Y','1',decode(a1903,'USD','2','1'))"
    '排序:幣別、獨立水單、代理人、匯款方式、收據公司別
    strSQLc = strSQLc & " order by a1903,flag1 desc,a1803,a0k11 "

   ADO24m0.Open strSQLc, adoTaie, adOpenStatic, adLockReadOnly
   If ADO24m0.RecordCount = 0 Then
      ADO24m0.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   
   ii = 1
   inR = 1
   strCurr = "" '幣別:預設空白
   
   If rsB.State <> adStateClosed Then rsB.Close
   Set rsB = Nothing
   Do While ADO24m0.EOF = False
      For intI = 0 To mHX_Num
          mHX(intI) = ""
      Next intI
      bolUnit = False
      strA0K11 = "" & ADO24m0.Fields("a0k11")  '收據公司別
      strHX1518 = ""
      strA2220 = ""
      strA2219 = "SHA" '手續費空白,預設SHA
      
      '序號(用在清單列印)
      mHX(0) = Format(inR, "0000")
      '(申請人)統一編號
      mHX(1) = convForm("" & ADO24m0.Fields("a0807"), 10)
      '(申請人)名稱及地址1 ~ (申請人)名稱及地址4 mHX(2),mHX(3),mHX(4),mHX(5)
      strExc(1) = "" & ADO24m0.Fields("a0803")
      strExc(2) = ADO24m0.Fields("a0822") & " " & ADO24m0.Fields("a0823")
      If Trim(strExc(2)) = "" Then
         strExc(2) = "9F, No. 112, Sec. 2, Chang-An E. Rd., Taipei 104, Taiwan, R.O.C." '預設地址
      End If
      'Added by Lydia 2021/08/27 寰華結匯水單本所地址不能出現ROC
      If InStr("Y53374,", Left(Trim(ADO24m0.Fields("a1803")), 6)) > 0 Then
          strExc(2) = Replace(strExc(2), ", R.O.C.", String(8, " "))
      End If
      'end 2021/08/27
      
      strExc(1) = UCase(strExc(1)): strExc(2) = UCase(strExc(2)) 'Swift文數字規定:大寫英文
      strExc(1) = CheckFstSpec(PUB_GetSimpleName(strExc(1), True)) 'Added by Lydia 2017/09/30 遇到特殊字元處理
      If GetTextLength(strExc(1)) <= 35 Then
         mHX(2) = convForm(strExc(1), 35)
         atJ = 3
      Else
         mHX(2) = Mid(strExc(1), 1, 35)
         'Modified by Lydia 2017/09/30 取消特殊字元處理
         'mHX(3) = convForm(CheckFstSpec(Mid(strExc(1), 36)), 35)
         mHX(3) = convForm(Mid(strExc(1), 36), 35)
         atJ = 4
      End If
      If GetTextLength(strExc(2)) <= 35 Then
         mHX(atJ) = convForm(strExc(2), 35)
         atJ = atJ + 1
      Else
         mHX(atJ) = Mid(strExc(2), 1, 35)
         'Modified by Lydia 2017/09/30 取消特殊字元處理
         'mHX(atJ + 1) = convForm(CheckFstSpec(Mid(strExc(2), 36)), 35)
         mHX(atJ + 1) = convForm(Mid(strExc(2), 36), 35)
         atJ = atJ + 2
      End If
      Do While atJ < 6
         mHX(atJ) = String(35, " ")
         atJ = atJ + 1
      Loop
      
      '匯款人身分別
      mHX(6) = " " '空白
      '匯款幣別
      mHX(7) = convForm("" & ADO24m0.Fields("a1903"), 3)
      '匯款金額
      mHX(8) = PUB_StrToStr(Format(ADO24m0.Fields("Amount"), "###0.00"), 15, True, True)
      '交易國別代號
      mHX(9) = convForm("" & ADO24m0.Fields("na60"), 2)
      '交易國別名稱
      mHX(10) = convForm("" & ADO24m0.Fields("na04"), 20)
      '匯款性質編號--預設19D
      mHX(11) = Mid(Pub_DBtype, 1, 3)
      '匯款性質編號其他補充說明 --設空白
      mHX(12) = String(35, " ")
      '受款人身分別--預設3.民間
      mHX(13) = "3"
      '(受益人)帳號
      mHX(14) = String(35, " ")
      'Added by Lydia 2017/10/03 (受益人)銀行Swift帳號
      mHX(19) = String(12, " ")
      
      '中間銀行Swift--預設空白
      mHX(24) = String(12, " ")

      '受益人和設帳銀行的名稱和地址 mHX(15) ~ mHX(23)
      strExc(5) = "":   strExc(6) = ""
      If IsNull(ADO24m0.Fields("fa05")) = False Then
         strExc(5) = RepEnter2zero(UCase(Trim(ADO24m0.Fields("fa05"))))
      End If
      If IsNull(ADO24m0.Fields("fa63")) = False Then
         strExc(5) = strExc(5) & " " & RepEnter2zero(UCase(Trim(ADO24m0.Fields("fa63"))))
      End If
      If IsNull(ADO24m0.Fields("fa64")) = False Then
         strExc(5) = strExc(5) & " " & RepEnter2zero(UCase(Trim(ADO24m0.Fields("fa64"))))
      End If
      If IsNull(ADO24m0.Fields("fa65")) = False Then
         strExc(5) = strExc(5) & " " & RepEnter2zero(UCase(Trim(ADO24m0.Fields("fa65"))))
      End If
      If IsNull(ADO24m0.Fields("Addr")) = False Then
         strExc(6) = RepEnter2zero(UCase(Trim(ADO24m0.Fields("Addr"))))
      End If
      '抓受款人相關資料
      strExc(7) = "": strExc(8) = ""
      'Modified by Lydia 2017/10/03 華銀整批媒體RMB改CNY
      'StrSqlB = "Select a.*,b.NA60,nvl(b.na04,b.na80) na04 From ACC220 a,NATION b Where a.A2217=b.NA01(+) and a2201='" & ADO24m0.Fields("a1803") & "' And A2202='" & ADO24m0.Fields("a1903") & "' "
      StrSqlB = "Select a.*,b.NA60,nvl(b.na04,b.na80) na04 From ACC220 a,NATION b Where a.A2217=b.NA01(+) and a2201='" & ADO24m0.Fields("a1803") & "' And A2202='" & IIf(ADO24m0.Fields("a1903") = J_RMB, "RMB", ADO24m0.Fields("a1903")) & "' "

      intI = 1
      Set rsB = ClsLawReadRstMsg(intI, StrSqlB)
      If intI = 1 Then
        '交易國別代號 =A2217受款人(行)國別
        If "" & rsB.Fields("NA60") <> "" Then
           mHX(9) = convForm("" & rsB.Fields("NA60"), 2)
        End If
        '國家名稱(受款行>代理人或客戶)
        If "" & rsB.Fields("NA04") <> "" Then
           mHX(10) = convForm("" & rsB.Fields("NA04"), 20)
        End If
        '(受益人)帳號--第一個須為/
        mHX(14) = convForm("/" & rsB.Fields("A2207"), 35)
        '(受益人)名稱和地址
        If "" & rsB.Fields("A2203") <> "" Then
           strExc(5) = RepEnter2zero(UCase(Trim(rsB.Fields("A2203"))))
        End If
        If "" & rsB.Fields("A2204") <> "" Then
           strExc(5) = strExc(5) & " " & RepEnter2zero(UCase(Trim(rsB.Fields("A2204"))))
        End If
        If "" & rsB.Fields("A2205") <> "" Then
           strExc(5) = strExc(5) & " " & RepEnter2zero(UCase(Trim(rsB.Fields("A2205"))))
        End If
        '短地址(優先)
        If "" & rsB.Fields("A2218") <> "" Then
           strExc(6) = RepEnter2zero(UCase(Trim(rsB.Fields("A2218"))))
        End If
        '(設帳銀行)swift代號 =A2211 受款銀行代號
        mHX(19) = convForm("" & rsB.Fields("A2211"), 12)
        '(設帳銀行)名稱
        If "" & rsB.Fields("A2208") <> "" Then
           strExc(7) = RepEnter2zero(UCase(Trim(rsB.Fields("A2208"))))
        End If
        If "" & rsB.Fields("A2209") <> "" Then
           strExc(7) = strExc(7) & " " & RepEnter2zero(UCase(Trim(rsB.Fields("A2209"))))
        End If
        '匯往大陸地區之人民幣匯款,請在設帳銀行名稱及地址第一行輸入"/CN"+12位CNAPS 或 "//CN"+12~14位CNAPS
        'Modified by Lydia 2017/09/30 改成變數
        'If mHX(9) = "CN" And mHX(7) = "RMB" Then
        If mHX(9) = "CN" And mHX(7) = J_RMB Then
           'Modified by Lydia 2023/12/29 銀行來電通知：匯款至大陸的代理人，TXT檔中需刪除CNAPS編碼的 "//CN" 資訊，僅需顯示Swift Code
           'strA2220 = "//CN" & RepEnter2zero(UCase(Trim("" & rsB.Fields("A2220"))))
           strA2220 = RepEnter2zero(UCase(Trim("" & rsB.Fields("A2220"))))
        End If
        '(中間銀行)
        If "" & rsB.Fields("A2214") <> "" Then
           If InStr(UCase(rsB.Fields("A2214")), "SWIFT") = 0 Then
              strExc(8) = strExc(8) & IIf(strExc(8) <> "", " ", "") & RepEnter2zero(UCase(Trim(rsB.Fields("A2214"))))
           Else  '抓Swift Code
              strExc(0) = Replace(UCase(rsB.Fields("A2214")), "：", ":")
              mHX(24) = convForm(RepEnter2zero(Trim(Mid(strExc(0), InStr(strExc(0), ":") + 1))), 12)
           End If
        End If
        If "" & rsB.Fields("A2215") <> "" Then
           If InStr(UCase(rsB.Fields("A2215")), "SWIFT") = 0 Then
              strExc(8) = strExc(8) & IIf(strExc(8) <> "", " ", "") & RepEnter2zero(UCase(Trim(rsB.Fields("A2215"))))
           Else
              strExc(0) = Replace(UCase(rsB.Fields("A2215")), "：", ":")
              mHX(24) = convForm(RepEnter2zero(Trim(Mid(strExc(0), InStr(strExc(0), ":") + 1))), 12)
           End If
        End If
        If "" & rsB.Fields("A2216") <> "" Then
           If InStr(UCase(rsB.Fields("A2216")), "SWIFT") = 0 Then
              strExc(8) = strExc(8) & IIf(strExc(8) <> "", " ", "") & RepEnter2zero(UCase(Trim(rsB.Fields("A2216"))))
           Else
              strExc(0) = Replace(UCase(rsB.Fields("A2216")), "：", ":")
              mHX(24) = convForm(RepEnter2zero(Trim(Mid(strExc(0), InStr(strExc(0), ":") + 1))), 12)
           End If
        End If
        '手續費
        If "" & rsB.Fields("A2219") <> "" Then
           strA2219 = UCase(rsB.Fields("A2219"))
        End If
        strA2222 = "" & rsB.Fields("A2222") 'Added by Lydia 2020/05/28 媒體備註
      End If
      
      '收款人名稱和地址去掉中文(婉莘說:銀行資料的地址都已經設成英文)
      strExc(5) = PUB_GetSimpleName(strExc(5), True) '暫存收款人名稱
      strExc(6) = PUB_GetSimpleName(strExc(6), True) '暫存收款人地址
      strExc(6) = Replace(strExc(6), "#", "") 'Added by Lydia 2017/09/18 配合華銀不接受#,預設拿掉#
      strExc(5) = UCase(strExc(5)): strExc(6) = UCase(strExc(6)) 'Swift文數字規定:大寫英文
       
      '記錄列印用的名稱+地址
      strHX1518 = strExc(5) & " ／ " & strExc(6)
      
      '1.處理受益人名稱和地址 mHX(15),mHX(16),mHX(17),mHX(18)
      strExc(5) = CheckFstSpec(PUB_GetSimpleName(strExc(5), True)) 'Added by Lydia 2017/09/30 遇到特殊字元處理
      If GetTextLength(strExc(5)) <= 35 Then
         mHX(15) = convForm(strExc(5), 35)
         atJ = 16
      Else
         'Modified by Lydia 2023/08/25
         'mHX(15) = Mid(strExc(5), 1, 35)
         mHX(15) = convForm(strExc(5), 35)
         'Modified by Lydia 2017/09/30 取消特殊字元處理
         'mHX(16) = convForm(CheckFstSpec(Mid(strExc(5), 36)), 35)
         'Modified by Lydia 2023/08/25
         'mHX(16) = convForm(Mid(strExc(5), 36), 35)
         mHX(16) = convForm(Replace(strExc(5), mHX(15), ""), 35)
         atJ = 17
      End If
      strExc(6) = CheckFstSpec(PUB_GetSimpleName(strExc(6), True)) 'Added by Lydia 2017/09/30 遇到特殊字元處理
      If GetTextLength(strExc(6)) <= 35 Then
         mHX(atJ) = convForm(strExc(6), 35)
         atJ = atJ + 1
      Else
         'Modified by Lydia 2023/08/25
         'mHX(atJ) = Mid(strExc(6), 1, 35)
         mHX(atJ) = convForm(strExc(6), 35)
         'Modified by Lydia 2017/09/30 取消特殊字元處理
         'mHX(atJ + 1) = convForm(CheckFstSpec(Mid(strExc(6), 36)), 35)
         'Modified by Lydia 2023/08/25
         'mHX(atJ + 1) = convForm(Mid(strExc(6), 36), 35)
         mHX(atJ + 1) = convForm(Replace(strExc(6), mHX(atJ), ""), 35)
         atJ = atJ + 2
      End If
      Do While atJ < 19
         mHX(atJ) = String(35, " ")
         atJ = atJ + 1
      Loop
      
      '2.處理設帳銀行(受款銀行)名稱,地址預設為空白 mHX(20),mHX(21),mHX(22),mHX(23)
        '匯往大陸地區之人民幣匯款,請在設帳銀行名稱及地址第一行輸入"/CN"+12位CNAPS 或 "//CN"+12~14位CNAPS
        'Memo by Lydia 2023/12/29 銀行來電通知：匯款至大陸的代理人，TXT檔中需刪除CNAPS編碼的 "//CN" 資訊，僅需顯示Swift Code
      strExc(7) = UCase(strExc(7)): strExc(8) = UCase(strExc(8)) 'Swift文數字規定:大寫英文
      atJ = 20
      'Modified by Lydia 2017/09/30 改成變數
      'If mHX(9) = "CN" And mHX(7) = "RMB" Then
      If mHX(9) = "CN" And mHX(7) = J_RMB Then
         mHX(20) = convForm(strA2220, 35)
         atJ = 21
      End If
      strExc(0) = CheckFstSpec(PUB_GetSimpleName(strExc(7), True))
      If GetTextLength(strExc(0)) <= 35 Then
         mHX(atJ) = convForm(strExc(0), 35)
         atJ = atJ + 1
      Else
         mHX(atJ) = Mid(strExc(0), 1, 35)
         'Modified by Lydia 2017/09/30 取消特殊字元處理
         'mHX(atJ + 1) = convForm(CheckFstSpec(Mid(strExc(0), 36)), 35)
         mHX(atJ + 1) = convForm(Mid(strExc(0), 36), 35)
         atJ = atJ + 2
      End If
      Do While atJ < 24
         mHX(atJ) = String(35, " ")
         atJ = atJ + 1
      Loop
      
      '(中間銀行)名稱及地址 mHX(25),mHX(26),mHX(27),mHX(28)
        atJ = 25
        strExc(0) = CheckFstSpec(PUB_GetSimpleName(strExc(8), True))
        Do While atJ < 29
           If GetTextLength(strExc(0)) = 0 Then
              mHX(atJ) = String(35, " ")
           ElseIf GetTextLength(strExc(0)) < 35 Then
              mHX(atJ) = convForm(strExc(0), 35)
              strExc(0) = ""
           Else
              mHX(atJ) = Mid(strExc(0), 1, 35)
              strExc(0) = CheckFstSpec(Mid(strExc(0), 36))
           End If
           atJ = atJ + 1
        Loop
      
      '費用明細
      mHX(29) = "SHA" '預設
      
      '付款明細1~4  mHX(30),mHX(31),mHX(32),mHX(33)
        '抓Debit Note
        strExc(9) = ""
        'Modified by Lydia 2018/04/23 寰華Y53374備註欄都不要再填寫(by 郭)
        'If Left(Trim(ADO24m0.Fields("a1803")), 6) <> "Y37580" Then '婧瑄說Y37580都不要印代理人D/BNo
        'Modified by Lydia 2020/04/22 +建毅Y51566,唯源Y52404
        If InStr("Y37580,Y53374,Y51566,Y52404,", Left(Trim(ADO24m0.Fields("a1803")), 6)) = 0 Then
           StrSqlB = "Select a1706 From acc190, acc170 Where a1902=a1702 And a1901='" & ADO24m0.Fields("a1901") & "' Group By a1706 Order By 1 "
        Else
           StrSqlB = "select null from dual"
        End If
        intI = 1
        Set rsB = ClsLawReadRstMsg(intI, StrSqlB)
        If intI = 1 Then
           intI = 0
           Do While Not rsB.EOF
              intI = intI + 1
              If intI > 12 Then Exit Do
              strExc(9) = strExc(9) & "," & rsB.Fields(0).Value
              strExc(9) = strExc(9) & IIf(ii = 12 And rsB.RecordCount > 12, " etc.", "")
              rsB.MoveNext
           Loop
        End If
        If strExc(9) <> "" Then strExc(9) = Mid(strExc(9), 2)

        '票匯和暫收款退費不加INV. (a1902=a1702為O單號)
        'Modified by Lydia 2020/04/22 +天津三元Y37580; +建毅Y51566,唯源Y52404
        'If InStr("Y53374", Left(Trim(ADO24m0.Fields("a1803")), 6)) = 0 Then  'Added by Lydia 2018/04/23 寰華Y53374備註欄都不要再填寫(by 郭)
        If InStr("Y37580,Y53374,Y51566,Y52404,", Left(Trim(ADO24m0.Fields("a1803")), 6)) = 0 Then
            If "" & ADO24m0.Fields("a1811") <> "1" And Val("" & ADO24m0.Fields("Ocnt")) = 0 Then  'Memo by Lydia 2018/06/26 非票匯和暫收款退費
              strExc(9) = "INV." & strExc(9)
            End If
        End If
        
        If strA2222 <> "" Then strExc(9) = PUB_GetSimpleName(strA2222, True, True) & " " & strExc(9) 'Added by Lydia 2020/05/28 備註前面+媒體備註
        
        strExc(9) = convForm(PUB_StrToStr(RepEnter2zero(CheckStr(strExc(9))), 140), 140)
        If Len(Trim(strExc(9))) > 135 Then
           If InStr(Mid(strExc(9), 134, 6), "etc.") > 0 Then
             strExc(9) = convForm(PUB_StrToStr(strExc(9), 140), 140)
           Else
             strExc(9) = convForm(PUB_StrToStr(strExc(9), 135) & " etc.", 140)
           End If
        End If
      '處理付款明細1~4 mHX(30),mHX(31),mHX(32),mHX(33)
      atJ = 30
      strExc(9) = UCase(strExc(9)) 'Swift文數字規定:大寫英文
      strExc(0) = CheckFstSpec(PUB_GetSimpleName(strExc(9), True))
      Do While atJ < 34
         If GetTextLength(strExc(0)) = 0 Then
            mHX(atJ) = String(35, " ")
         ElseIf GetTextLength(strExc(0)) < 35 Then
            mHX(atJ) = convForm(strExc(0), 35)
            strExc(0) = ""
         Else
            mHX(atJ) = Mid(strExc(0), 1, 35)
            strExc(0) = CheckFstSpec(Mid(strExc(0), 36))
         End If
         atJ = atJ + 1
      Loop
      
      '授權扣帳帳號1
      mHX(34) = "145100236819"
      '授權扣帳幣別1
      mHX(35) = "TWD"
      '授權扣帳帳號2
      mHX(36) = String(12, " ")
      '授權扣帳幣別2
      mHX(37) = String(3, " ")
      '匯率議價編號
      If strCurr <> mHX(7) Then
         '抓最新日期的議價編號(有值)
         'Modified by Lydia 2017/09/30 改為抓取結匯日期or前一工作天的資料
         'StrSqlB = "select * from acc1a0 where a1a03='" & mHX(7) & "' and a1a01=(select max(a1a01) from acc1a0 where a1a03='" & mHX(7) & "' and nvl(a1a12,'N') <> 'N')"
         StrSqlB = "select * from acc1a0 where a1a03='" & mHX(7) & "' and a1a01=(select max(a1a01) from acc1a0 where a1a03='" & mHX(7) & _
                   "' and a1a01>=" & TransDate(CompWorkDay(2, DBDATE(strDBdate), 1), 1) & " and a1a01<=" & strDBdate & " and nvl(a1a12,'N') <> 'N')"
         intI = 1
         Set rsB = ClsLawReadRstMsg(intI, StrSqlB)
         If intI = 1 Then
            strA1a12 = "" & rsB.Fields("a1a12")
         Else
            strA1a12 = ""
         End If
      End If
      strCurr = mHX(7)
      mHX(38) = convForm(strA1a12, 12)
      
      '匯款繳款方式
      mHX(39) = "1" '新台幣
      '費用繳款方式
      mHX(40) = "1" '新台幣
      '聯絡電話(申請人)
      mHX(41) = convForm("" & ADO24m0.Fields("a0813"), 20)
      '拍發兩通電文
      '當受款銀行手續費為OUR,採2通電文; 當交易國別為US(美國)時,不可選擇拍發兩通電文
      If InStr(strA2219, "OUR") > 0 And mHX(9) <> "US" Then
         mHX(42) = "Y"
      Else
         mHX(42) = "N"
      End If
      'Added by Lydia 2017/09/30 華銀結匯之手續費是以OUR方式, 且匯款地在美國者, 因為美國的OUR無法拍2通電文, 所以需要改成29欄要改OUR
      'Modified by Lydia 2018/12/17 by Lydia 2018/12/17 華銀表示, 匯款幣別是美金的話, 需改成OUR+Y (過去規則不符現況)
      'If InStr(strA2219, "OUR") > 0 And mHX(9) = "US" Then
      '   mHX(29) = "OUR"
      'End If
      ''end 2017/09/30
      'Modified by Lydia 2019/01/28 美金的OUR才要改Y
      'If UCase(mHX(7)) = "USD" Then
      '   mHX(29) = "OUR"
      'Modified by Lydia 2022/09/14 Y55766(German Patent and Trade Mark德國專利局)代理人手續費=71:OUR，須呈現"OUR"+2通電文
      'If InStr(strA2219, "OUR") > 0 And UCase(mHX(7)) = "USD" Then
      If InStr(strA2219, "OUR") > 0 And (UCase(mHX(7)) = "USD" Or "" & ADO24m0.Fields("a1803") = "Y55766000") Then
      'end 2019/01/28
         mHX(29) = "OUR" 'Added by Lydia 2019/03/08 只有美金並且手續費為OUR時, 手續費才能從預設的SHA改為OUR
         mHX(42) = "Y"
      End If
      'end 2018/12/17

'-------------------------------------------------
      bolAdd = True 'Added by Lydia 2023/08/28 預設增加序號
      '獨立水單不合併明細
      If ADO24m0.Fields("a1812") = "Y" Then bolUnit = True
         '合併明細資料
         'Modified by Lydia 2021/04/12 +判斷AR05獨立水單
         StrSqlB = "Select * From accrpt24m0_2 where UNO='" & strUserNum & "' and ar00='" & ADO24m0.Fields("a1803") & "' " & _
                   "And ar01='" & ADO24m0.Fields("a1903") & "' And ar03='" & ADO24m0.Fields("flag1") & "' and ar04='" & strA0K11 & "' and ar05 is null "
         strExc(1) = "":          intI = 1
         Erase tmpStr
         Set rsB = ClsLawReadRstMsg(intI, StrSqlB)
         If intI = 1 And bolUnit = False Then
            '串D/B note
            If Trim(rsB.Fields("HX30") & rsB.Fields("HX31") & rsB.Fields("HX32") & rsB.Fields("HX33")) <> Trim(mHX(30) & mHX(31) & mHX(32) & mHX(33)) Then
               strExc(2) = convForm(Trim(rsB.Fields("HX30") & rsB.Fields("HX31") & rsB.Fields("HX32") & rsB.Fields("HX33")) & "," & Trim(mHX(30) & mHX(31) & mHX(32) & mHX(33)), 140)
               If Len(Trim(strExc(2))) > 135 Then
                  If InStr(Mid(strExc(2), 134, 6), "etc.") > 0 Then
                  strExc(2) = convForm(PUB_StrToStr(strExc(2), 140), 140)
                Else
                  strExc(2) = convForm(PUB_StrToStr(strExc(2), 135) & " etc.", 140)
                  End If
               End If
               strExc(0) = CheckFstSpec(UCase(PUB_GetSimpleName(strExc(2), True))) 'Swift文數字規定:大寫英文
               atJ = 1
               Do While atJ < 5
                  If GetTextLength(strExc(0)) = 0 Then
                      tmpStr(atJ) = String(35, " ")
                  ElseIf GetTextLength(strExc(0)) < 35 Then
                      tmpStr(atJ) = convForm(strExc(0), 35)
                      strExc(0) = ""
                  Else
                      tmpStr(atJ) = Mid(strExc(0), 1, 35)
                      strExc(0) = CheckFstSpec(Mid(strExc(0), 36))
                  End If
                  atJ = atJ + 1
               Loop
               strExc(1) = ",HX30=" & CNULL(tmpStr(1)) & ",HX31=" & CNULL(tmpStr(2)) & ",HX32=" & CNULL(tmpStr(3)) & ",HX33=" & CNULL(tmpStr(4))
            End If
            
            bolAdd = False 'Added by Lydia 2023/08/28
            strExc(1) = "update accrpt24m0_2 set HX08=HX08+" & Val(mHX(8)) & strExc(1)
            strExc(1) = strExc(1) & " where UNO='" & strUserNum & "' and ar00='" & ADO24m0.Fields("a1803") & "' And ar01='" & ADO24m0.Fields("a1903") & "' " & _
                       "And ar03='" & ADO24m0.Fields("flag1") & "' and ar04='" & strA0K11 & "' "
         Else
            'PK ~ UNO(使用者),AR00(代理人編號)+AR01(幣別)+AR02(代理人名稱)+AR03(flag1)+AR04(A1917公司別)
            'Modified by Lydia 2021/04/12 +AR05獨立水單=a1812
            strExc(1) = "INSERT INTO accrpt24m0_2 values(" & CNULL(strUserNum) & "," & CNULL(ChgSQL("" & ADO24m0.Fields("a1803"))) & "," & CNULL(ChgSQL("" & ADO24m0.Fields("a1903"))) & _
                        "," & CNULL(convForm(PUB_StrToStr(ChgSQL("" & ADO24m0.Fields("fa05")), 30), 30)) & "," & CNULL(ADO24m0.Fields("flag1")) & "," & CNULL(strA0K11) & "," & CNULL("" & ADO24m0.Fields("a1812"))
            
            For intI = 0 To mHX_Num
                'HX08 金額
                If intI = 8 Then
                   strExc(1) = strExc(1) & "," & Val(mHX(intI))
                Else
                   strExc(1) = strExc(1) & "," & CNULL(ChgSQL(mHX(intI)))
                End If
            Next intI
            'HX1518名稱+地址合併列印
            strExc(1) = strExc(1) & "," & CNULL(PUB_StrToStr(ChgSQL(strHX1518), 200))
            strExc(1) = strExc(1) & ")"
            ii = ii + 1
         End If
         
         adoTaie.Execute strExc(1)
         
      'Modified by Lydia 2023/08/28 判斷合併水單不跳號
      'inR = inR + 1
      If bolAdd = True Then
        inR = inR + 1
      End If
      'end 2023/08/28
      ADO24m0.MoveNext
   Loop
   
   ADO24m0.Close
   If rsB.State <> adStateClosed Then rsB.Close
   Set rsB = Nothing
End Sub

'Added by Lydia 2017/09/06 讀取明細暫存
Private Sub PStoreRead_H106(Optional ByVal aKind As String)
Dim ff As Integer, strPath As String
Dim strR As String
Dim rsR As New ADODB.Recordset
Dim strFileNo As String, strFileName As String '檔名
Dim bolPrint As Boolean '開始列印
Dim tmpGrp As String, tmpSubTot As Double, TmpCnt As Integer, tmpI As Integer
Dim m2FileName As String  '整批匯出匯款申請書(檔名)
Dim strName As String, strText As String
Dim intJ As Integer
Dim strTel As String

    If rsR.State <> adStateClosed Then rsR.Close
    Set rsR = Nothing
    strR = "Select a.*,a0802,a0813 From accrpt24m0_2 a ,acc080 where UNO='" & strUserNum & "' and ar04=a0801(+) order by HX00 "
    
    mCompany = "": mAcctNo = ""
    intI = 1
    Set rsR = ClsLawReadRstMsg(intI, strR)
    If intI = 1 Then
        strPath = PUB_Getdesktop
        strPath = strPath & "\華銀結匯水單"
        If Dir(strPath, vbDirectory) = "" Then
           MkDir strPath
        End If
        
StartPrt:
        strFileNo = "": plStr(0) = ""
        mAppDesc = ""
        If bolPrint Then
            rsR.MoveFirst
           '設定印表機
            Printer.EndDoc
            Printer.Orientation = 2 '1.直印 2.橫印
            Printer.PaperSize = 9  'A4
               
            lngPageHeight = Printer.ScaleHeight
            lngPageWidth = Printer.ScaleWidth
            lngLineHeight = 270
            GetPleft_H106 '設定邊界
            SetColumnName_H106
            iPage = 0
        Else
           idx = 0
        End If
    
        Do While rsR.EOF = False
            '華銀不分檔
            If strFileNo = "" Then
                plStr(0) = rsR.Fields("HX00")
                strFileNo = strDBdate
                If bolPrint Then
                    iPage = iPage + 1
                    mCount = 0: mTotal = 0
                    PrintHeader_H106
                Else
                    idx = idx + 1
                    If ff > 0 Then Close #ff
                    ff = FreeFile
                    strFileName = strPath & "\" & strFileNo & ".txt"
                    
                    Open strFileName For Output As ff
                End If
            End If
            
            '匯款人
            mCompany = Trim("" & rsR.Fields("a0802"))
            mAcctNo = Trim("" & rsR.Fields("HX34"))
            strTel = Trim("" & rsR.Fields("a0813"))
            
            mHX(0) = Trim("" & rsR.Fields("HX00")) '列印編號,不輸出
            strExc(1) = ""
            For intI = 1 To mHX_Num
               If intI = 8 Then
                  '匯款金額15字元,轉媒體檔若長度<15,在左方補0
                  If bolPrint = False Then
                     '小數位2位預設補0,去掉小數點
                     mHX(intI) = Right(String(15, "0") & Replace(Format(rsR.Fields("HX08"), "###0.00"), ".", ""), 15)
                  Else
                     mHX(intI) = PUB_StrToStr(Format(rsR.Fields("HX08"), "###0.00"), 15, True, True)
                  End If
               Else
                  'Modified by Lydia 2021/04/12  +AR05獨立水單=a1812
                  'mHX(intI) = "" & rsR.Fields(intI + 6) '0~5屬PK,6=HX00編號
                  mHX(intI) = "" & rsR.Fields(intI + 7) '0~6屬PK,7=HX00編號
               End If
               strExc(1) = strExc(1) & mHX(intI)
            Next intI

            '計算:華銀-整批匯出匯款申請書的匯款筆數/匯款幣別及匯款金額
            'Mark by Lydia 2017/09/28 各項幣別合計放在報表後面
            'If bolPrint = False Then
                If tmpGrp <> "" & rsR.Fields("ar01") Then
                   If tmpGrp <> "" Then
                      mAppDesc = mAppDesc & tmpI & "/" & tmpGrp & Format(tmpSubTot, "##,##0.00") & ";"
                   End If
                   tmpGrp = rsR.Fields("ar01")
                   tmpI = 1
                   tmpSubTot = Val(rsR.Fields("HX08"))
                Else
                   tmpI = tmpI + 1
                   tmpSubTot = tmpSubTot + Val(rsR.Fields("HX08"))
                End If
                TmpCnt = TmpCnt + 1
            'End If
            
            If bolPrint Then
                '編號
                plStr(0) = mHX(0)
                '匯款人證號(統編)
                plStr(1) = mHX(1)
                '交易國別
                plStr(2) = mHX(9)
                '收款行Swift 或名稱
                plStr(3) = PUB_StrToStr(IIf(Trim(mHX(19)) <> "", mHX(19), Trim(mHX(20) & " " & mHX(21) & " " & mHX(22) & " " & mHX(23))), 80)
                '收款人帳號
                plStr(4) = mHX(14)
                '幣別
                plStr(5) = mHX(7)
                '匯款金額
                plStr(6) = PUB_StrToStr(Format(Val(mHX(8)), "##,##0.00"), 15, True, True)
                '手續費
                plStr(7) = mHX(29)
                '匯款性質
                plStr(8) = mHX(11)
                '台一編號
                plStr(9) = "" & rsR.Fields("ar00")
                plStr(10) = Trim("" & rsR.Fields("HX1518"))
                
                '+第3行 中間銀行資料
                pStrL3 = ""
                If Len(Trim(mHX(24) & mHX(25) & mHX(26) & mHX(27) & mHX(28))) > 0 Then
                   pStrL3 = "中間銀行："
                   If Trim(mHX(24)) <> "" Then
                      pStrL3 = pStrL3 & "swift " & mHX(24)
                   End If
                   pStrL3 = pStrL3 & " " & Trim(mHX(25) & " " & mHX(26) & " " & mHX(27) & " " & mHX(28))
                End If
                mCount = mCount + 1
                mTotal = mTotal + Val(mHX(8))
                PrintDetail_H106
            Else
                Print #ff, strExc(1)
            End If
            
           rsR.MoveNext
        Loop
           
           
        If bolPrint Then
           'Added by Lydia 2017/09/28
           mAppDesc = mAppDesc & tmpI & "/" & tmpGrp & Format(tmpSubTot, "##,##0.00") & ";"
           tmpGrp = ""
           'end 2017/09/28
           GoTo EndPrt
        End If
        
        Close #ff
        
        '華銀-整批匯出匯款申請書的匯款筆數/匯款幣別及匯款金額
        '---------------------------
        mAppDesc = mAppDesc & tmpI & "/" & tmpGrp & Format(tmpSubTot, "##,##0.00") & ";"
        m2FileName = strPath & "\" & strDBdate & "_" & "整批匯出匯款申請書.doc"
        If Dir(m2FileName) <> "" Then
           Kill m2FileName
        End If
        tmpGrp = "" 'Added by Lydia 2017/09/28
        
        '判斷word是否已開啟
        If g_WordAp Is Nothing Then
           Set g_WordAp = New Word.Application
           g_WordAp.Visible = False
        End If
        
        tmpArr = Empty
        tmpArr = Split(mAppDesc, ";")
        If UBound(tmpArr) >= 10 Then
           If Trim(tmpArr(10)) <> "" Then
              MsgBox "幣別項目超過9項,請通知電腦中心!!"
           End If
        End If
        
        intJ = 0
        idx = idx + 1
        g_WordAp.Documents.Open App.path & "\" & m_FileName
        g_WordAp.ActiveDocument.SaveAs m2FileName
        g_WordAp.ActiveDocument.Close
        g_WordAp.Documents.Open m2FileName
        With g_WordAp
           .Selection.WholeStory
           .Selection.Copy
           For intI = 0 To 26
              strName = ""
              strText = ""
              If intI = 0 Then
                 strName = "YY"
                 strText = Mid(strDBdate, 1, 3)
              ElseIf intI = 1 Then
                 strName = "MM"
                 strText = Mid(strDBdate, 4, 2)
              ElseIf intI = 2 Then
                 strName = "DD"
                 strText = Mid(strDBdate, 6, 2)
              ElseIf intI = 3 Then
                 strName = "中文公司名"
                 strText = mCompany
              ElseIf intI = 4 Then
                 strName = "英文公司名"
                 strText = Trim(mHX(2) & mHX(3))
              ElseIf intI = 5 Then
                 strName = "英文地址"
                 strText = Trim(mHX(4)) & " " & Trim(mHX(5))
              ElseIf intI = 6 Then
                 strName = "公司電話"
                 strText = strTel
              ElseIf intI = 7 Then
                 strName = "公司統編"
                 strText = Trim(mHX(1))
             '---匯款幣別和匯款金額 ------------
              ElseIf intI >= 8 And intI <= 25 Then
                 If intI Mod 2 = 0 Then
                    Select Case intI
                        Case 8:  strName = "D10"
                        Case 10: strName = "D20"
                        Case 12: strName = "D30"
                        Case 14: strName = "D40"
                        Case 16: strName = "D50"
                        Case 18: strName = "D60"
                        Case 20: strName = "D70"
                        Case 22: strName = "D80"
                        Case 24: strName = "D90"
                    End Select
                    If intJ <= UBound(tmpArr) Then
                       If Trim(tmpArr(intJ)) <> "" Then
                          strText = Mid(Trim(tmpArr(intJ)), 1, InStr(Trim(tmpArr(intJ)), "/") - 1)
                       Else
                          strText = ""
                       End If
                    Else
                       strText = ""
                    End If
                 Else
                    Select Case intI
                        Case 9:  strName = "D11"
                        Case 11: strName = "D21"
                        Case 13: strName = "D31"
                        Case 15: strName = "D41"
                        Case 17: strName = "D51"
                        Case 19: strName = "D61"
                        Case 21: strName = "D71"
                        Case 23: strName = "D81"
                        Case 25: strName = "D91"
                    End Select
                    If intJ <= UBound(tmpArr) Then
                       If Trim(tmpArr(intJ)) <> "" Then
                          strText = Mid(Trim(tmpArr(intJ)), InStr(Trim(tmpArr(intJ)), "/") + 1)
                          intJ = intJ + 1
                       Else
                          strText = ""
                       End If
                    Else
                       strText = ""
                    End If
                 End If
              ElseIf intI = 26 Then
                 strName = "TOT1"
                 strText = Trim(TmpCnt)
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
                 .Selection.TypeText strText
              End If
           Next intI
        End With
          
        ''Modified by Lydia 2024/01/31 改存成PDF
        'g_WordAp.ActiveDocument.Save
        'g_WordAp.ActiveDocument.Close
        If PUB_PrintWord2File(g_WordAp, strPath, strDBdate & "_" & "整批匯出匯款申請書") = False Then
           Exit Sub
        Else
           PUB_DelPCOrgFile m2FileName '刪除原本Word檔
        End If
        'end 2024/01/31
        
        Clipboard.Clear '清除剪貼簿動作
        '---------------------------
        
        If TxtList <> "Y" Then
           rsR.Close
           MsgBox "已在桌面的華銀結匯水單資料夾產生 " & idx & " 個檔案!!", vbInformation
        End If
        
        If TxtList = "Y" And bolPrint = False Then
           bolPrint = True
           GoTo StartPrt
        End If
    End If
'------------------------
EndPrt:
   If bolPrint Then
        PrintSubTotal_H106
        PrintSign
        Printer.EndDoc
        MsgBox "已在桌面的華銀結匯水單資料夾產生 " & idx & " 個檔案, 並列印清單完成!!", vbInformation
        rsR.Close
   End If
End Sub

'Added by Lydia 2017/09/12
Private Sub GetPleft_H106() '列印位置
Printer.Font.Name = "新細明體"
Printer.Font.Size = ciFontSize
Printer.Font.Bold = False
Printer.Font.Underline = False

Erase PLt2

   PLt2(0) = ciStartX 'OR01編號
   PLt2(1) = PLt2(0) + Printer.TextWidth(String(2, "　")) + ciColGap    '+序號
   PLt2(2) = PLt2(1) + Printer.TextWidth(String(5, "　")) + ciColGap    '+匯款人證號=HX01申請人統一編號
   PLt2(3) = PLt2(2) + Printer.TextWidth(String(2, "　")) + ciColGap    '+國別(HX09)
   PLt2(4) = PLt2(3) + Printer.TextWidth(String(40, "　")) + ciColGap   '+收款行Swift code或名稱(HX19~HX23)
   PLt2(5) = PLt2(4) + Printer.TextWidth(String(18, "　")) + ciColGap   '+收款人帳號(HX14)
   PLt2(6) = PLt2(5) + Printer.TextWidth(String(2, "　")) + ciColGap    '+幣別(HX07)
   PLt2(7) = PLt2(6) + Printer.TextWidth(String(7, "　")) + ciColGap    '+匯款金額(HX08)
   PLt2(8) = PLt2(7) + Printer.TextWidth(String(3, "　")) + ciColGap    '+手續費(HX29)
   '第2行
   PLt2(9) = PLt2(1)   '台一編號=>客戶/代理人編號
   PLt2(10) = PLt2(3)  '收款人名稱 ／ 地址
   PLt2(11) = lngPageWidth - ciStartX   '列印右邊界
   
   pLmax = 11  '華銀-列印定位最大值
   pStrMax = 10  '華銀-列印欄位最大值
   
End Sub

Private Sub SetColumnName_H106() '欄位抬頭
   tlStr(0) = "編號"
   tlStr(1) = "匯款人證號"
   tlStr(2) = "國別"
   tlStr(3) = "收款行SWIFT CODE或名稱"
   tlStr(4) = "收款人帳號"
   tlStr(5) = "幣別"
   tlStr(6) = "匯款金額"
   tlStr(7) = "手續費"
   tlStr(8) = "匯款性質"
   '第2行
   startL2 = 9
   tlStr(9) = "台一編號"
   tlStr(10) = "收款人名稱 ／ 地址"
End Sub

'合計
Private Sub PrintSubTotal_H106()
Dim strX As String
Dim intP As Integer 'Added by Lydia 2017/09/28

  'strX = "總計筆數：" & PUB_StrToStr(CheckStr(mCount), 3, True, True) & " 筆　　總計金額：共 " & Format(mTotal, "#,##0.00") & " 元" 'Remove by Lydia 2017/09/28
  PrintLine "2"
  PrintNewLine_H106 (0.5)
  
  Printer.Font.Size = ciTitleFontSize - 2
  'Modified by Lydia 2017/09/28 各項幣別合計放在報表後面
  'Printer.CurrentX = PLt2(5) - Printer.TextWidth(strX)
  'Printer.CurrentY = iPrint
  'Printer.Print strX
  tmpArr = Empty
  tmpArr = Split(mAppDesc, ";")
  For intI = 0 To UBound(tmpArr)
     strExc(1) = Trim(tmpArr(intI))
     If strExc(1) <> "" Then
        strExc(2) = Mid(strExc(1), 1, InStr(strExc(1), "/") - 1)      '筆數
        strExc(3) = Mid(strExc(1), InStr(strExc(1), "/") + 1, 3)    '幣別
        strExc(4) = Mid(strExc(1), InStr(strExc(1), strExc(3)) + 3) '金額
        intP = PLt2(4) - 2000
        Printer.CurrentX = intP
        Printer.CurrentY = iPrint
        Printer.Print strExc(3)
        Printer.CurrentX = intP + 650
        Printer.CurrentY = iPrint
        Printer.Print "筆數：" & Right(String(3, " ") & strExc(2), 3) & " 筆"
        Printer.CurrentX = intP + 3000
        Printer.CurrentY = iPrint
        Printer.Print "金額：共 " & strExc(4) & " 元"
        PrintNewLine_H106
     End If
  Next intI
  'end 2017/09/28
    
End Sub

Private Sub PrintHeader_H106()
Dim PriTitle1 As String, PriTitle2 As String
Dim x1 As Integer, x2 As Integer, x3 As Integer
Dim aP1 As Integer
PriTitle1 = "華南銀行外匯整批匯款明細表"
PriTitle2 = "(代匯出匯款申請書)"

iPrint = ciStartY
ColH_1 = ciStartY + lngLineHeight * 3 '欄線
ColH_2 = ColH_1 + lngLineHeight * 2

x1 = 5800
Printer.Font.Size = ciTitleFontSize
Printer.Font.Bold = True
Printer.CurrentX = x1
Printer.CurrentY = iPrint
Printer.Print PriTitle1

x2 = Printer.TextWidth(PriTitle1)
Printer.Font.Size = ciTitleFontSize - 4
Printer.Font.Bold = False
Printer.CurrentX = x1 + x2 + 50
Printer.CurrentY = iPrint + 50
Printer.Print PriTitle2

Printer.Font.Size = ciFontSize
Printer.CurrentX = 14300
Printer.CurrentY = iPrint + 50
Printer.Print "頁　　次：" & iPage

PrintNewLine_H106 (1.5)
Printer.Font.Size = ciTitleFontSize - 4
Printer.CurrentX = ciStartX
Printer.CurrentY = iPrint
Printer.Print "匯款人：" & mCompany
Printer.CurrentX = 5800
Printer.CurrentY = iPrint
Printer.Print "扣款帳號：" & mAcctNo
Printer.CurrentX = 14300
Printer.CurrentY = iPrint
Printer.Print "交易生效日：" & ChangeTStringToTDateString(strDBdate)

Printer.Font.Size = ciFontSize
PrintNewLine_H106 (1.5)

PrintLine "2"

PrintNewLine_H106 (0.5)

For aP1 = 0 To pStrMax
  If aP1 = 0 Then '畫欄線
    PrintColH (PLt2(aP1) - 60)
  Else
    If aP1 < startL2 Then PrintColH (PLt2(aP1) - 30)
  End If
  x1 = iPrint - 40
  If aP1 >= startL2 Then     '第2行
    x1 = x1 + Printer.TextHeight("匯款") + 20
  End If
  Printer.CurrentX = PLt2(aP1)
  Printer.CurrentY = x1
  Printer.Print tlStr(aP1)
Next aP1

PrintColH (PLt2(pLmax))
PrintNewLine_H106 (1.5)
PrintLine "2"
PrintNewLine_H106 (0.5)
End Sub

Private Sub PrintDetail_H106()
Dim aP1 As Integer, pB As String

For aP1 = 0 To pStrMax
    If aP1 = 6 Then '金額類-置右
        Printer.CurrentX = PLt2(aP1 + 1) - Printer.TextWidth(plStr(aP1)) - ciColGap
        Printer.CurrentY = iPrint
    Else
        If aP1 >= startL2 Then '第2行
            If aP1 = startL2 Then PrintNewLine_H106
        End If
            Printer.CurrentX = PLt2(aP1)
            Printer.CurrentY = iPrint
    End If
    '判斷名稱+地址
    If aP1 = startL2 + 1 Then
       '列印寬度超過可印範圍，分2行
       If PLt2(aP1) + Printer.TextWidth(plStr(aP1)) > PLt2(pLmax) Then
          Printer.Print Mid(plStr(aP1), 1, InStr(plStr(aP1), "／"))
          PrintNewLine_H106
          Printer.CurrentX = PLt2(aP1)
          Printer.CurrentY = iPrint
          Printer.Print Mid(plStr(aP1), InStr(plStr(aP1), "／") + 1)
       Else
          Printer.Print plStr(aP1)
       End If
    Else
       Printer.Print plStr(aP1)
    End If
Next aP1

PrintNewLine_H106
'第3行
If Len(pStrL3) > 0 Then
    Printer.CurrentX = PLt2(1)
    Printer.CurrentY = iPrint
    Printer.Print pStrL3
    PrintNewLine_H106
End If
PrintNewLine_H106 '空一行
    
End Sub

Private Sub PrintNewLine_H106(Optional ByVal mRate As Single = 1, Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 4)
   iPrint = iPrint + lngLineHeight * mRate
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint

      iPage = iPage + 1
      Printer.NewPage
      PrintHeader_H106
   End If
End Sub

'Added by Lydia 2025/05/27
'*************************************************
' 產生暫存檔---台銀114年8月上線---114/6/4通知提前上線; 114/6/9 回復TXT格式無誤; 114/6/11以後直接上線
'*************************************************
Private Sub PStoreData_T114()
'若結匯規則有變更，請加註文件：\\LINUX\PolyCOM\TaieNew\電腦中心日常工作\結匯-預設匯款方式(a1811和媒體檔).doc
Dim strSql As String, StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim strSQLc As String, strA0K11 As String
Dim ii As Integer
Dim bolUnit As Boolean '獨立水單資料不合併計算
Dim strNA60 As String '受款地國別(台銀代號)
Dim strNA01 As String '國家代號(受款行>代理人或客戶)
Dim strNA02 As String '國家地區(受款行>代理人或客戶)
Dim strAR07 As String '記錄列印用的名稱+地址
Dim strAR06 As String
Dim strA2220 As String '記錄CNAPS
Dim strA1901List As String '記錄付款單號
Dim strA2222 As String '記錄媒體備註
Dim strA2210 As String, strA2226 As String  'Added by Lydia 2025/10/17

On Error GoTo ErrHandle 'Added by Lydia 2017/09/22

   adoTaie.Execute " delete from accrpt24m0_T114 where UNO='" & strUserNum & "' "

   '結匯日期=>台銀直接用系統日
   strDBdate = strSrvDate(2)
   
   '付款日期
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a1802 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a1802 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   '代理人
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and a1803 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a1803 <= '" & Text2 & "'"
   End If
   '付款單號
   If Text4 <> MsgText(601) Then
      strSql = strSql & " and a1801 >= '" & Text4 & "'"
   End If
   If Text3 <> MsgText(601) Then
      strSql = strSql & " and a1801 <= '" & Text3 & "'"
   End If
   '公司別
   strSql = strSql & " and a1917<>'J'" '不含Ｊ公司
   
   strSql = strSql & " and a1811<>'6' " '排除匯款方式6-抵帳
   
   excelSql = strSql '傳結匯清單SQL條件 (全電匯+票匯)
   
   '限電匯=> 2.電匯+5.台銀合併結匯
   strSql = strSql & " and (a1811=2 or a1811=5) "
   
   '明細：與excelsav2 可能有差異,例如:104/03/05~104/03/06 accrpt218 有W10400499,無W10400503
    ADO24m0.CursorLocation = adUseClient
    strSQLc = "select a1901, a1803, Decode(fa05, Null, Nvl(FA04, FA06), FA05) As FA05, Decode(FA05, Null, Null, fa63) As FA63, Decode(FA05, Null, Null, fa64) As FA64, Decode(FA05, Null, Null, fa65) As FA65, fa10, na03, a1903, a1810, a1811, sum(a1904) as Amount, decode(a1917,'J','J','2') As a0k11,rtrim(max(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70)) Addr,a0803,a0807,a0813,na60, na02,A0822,A0823  " & _
              ",decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1, a1812,sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt from acc190, acc180, fagent, nation,acc080 where a1801=a1901 and a1908 is null and A1803>'Y' AND substr(a1803,1,8)=fa01 and substr(a1803,9,1)=fa02 and fa10=na01(+) and a1810 is null and a0801=decode(a1917,'J','J','2') " & strSql & " group by a1803, a1901, Decode(fa05, Null, Nvl(FA04, FA06), FA05), Decode(FA05, Null, Null, fa63)," & _
              "Decode(FA05, Null, Null, fa64), Decode(FA05, Null, Null, fa65), fa10, na03, a1903, a1810, a1811, decode(a1917,'J','J','2') ,a0803,a0807,a0813,na60, na02,A0822,A0823 ,decode(a1812,'Y','1',decode(a1903,'USD','2','1')),a1812"
    strSQLc = strSQLc & " Union select a1901, a1803, substr(a1810, 1, 30) as fa05, substr(a1810, 31, 30) as fa63, substr(a1810, 61, 30) as fa64, substr(a1810, 91, 30) as fa65, '' as fa10, na03, a1903, a1810, a1811, sum(a1904) as Amount, decode(a1917,'J','J','2') As a0k11,rtrim(max(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70)) Addr,a0803,a0807,a0813,na60, na02,A0822,A0823  " & _
              ",decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1, a1812,sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt from acc190, acc180, fagent, nation,acc080 where a1801=a1901 and a1908 is null and substr(a1803,1,8)=fa01(+) and A1803>'Y' AND substr(a1803,9,1)=fa02(+) and fa10=na01(+) and a1810 is not null and a0801=decode(a1917,'J','J','2') " & strSql & " group by a1803,a1901,na03,a1903,a1810,a1811,decode(a1917,'J','J','2'),a0803,a0807,a0813,na60, na02,A0822,A0823 " & _
              ",decode(a1812,'Y','1',decode(a1903,'USD','2','1')),a1812"
    strSQLc = strSQLc & " Union select a1901, a1803, Decode(CU05, Null, Nvl(CU04, CU06), CU05) As FA05, Decode(CU05, Null, Null, CU88) As FA63, Decode(CU05, Null, Null, CU89) As FA64, Decode(CU05, Null, Null, CU90) As FA65, CU10 AS FA10, na03, a1903, a1810, a1811, sum(a1904) as Amount, decode(a1917,'J','J','2') As a0k11,rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)) Addr,a0803,a0807,a0813,na60, na02,A0822,A0823  " & _
              ",decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1, a1812,sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt from acc190, acc180, CUSTOMER, nation,acc080 where a1801=a1901 and a1908 is null and A1803<'Y' AND substr(a1803,1,8)=CU01 and substr(a1803,9,1)=CU02 and CU10=na01(+) and a1810 is null and a0801=decode(a1917,'J','J','2')" & strSql & " group by a1803, a1901, Decode(CU05, Null, Nvl(CU04, CU06), CU05), Decode(CU05, Null, Null, CU88), Decode(CU05, Null, Null, CU89)," & _
              "Decode(CU05, Null, Null, CU90), CU10, na03, a1903, a1810, a1811, decode(a1917,'J','J','2'),a0803,a0807,a0813,na60, na02,A0822,A0823 ,decode(a1812,'Y','1',decode(a1903,'USD','2','1')),a1812"
    strSQLc = strSQLc & " Union select a1901, a1803, substr(a1810, 1, 30) as fa05, substr(a1810, 31, 30) as fa63, substr(a1810, 61, 30) as fa64, substr(a1810, 91, 30) as fa65, '' as fa10, na03, a1903, a1810, a1811, sum(a1904) as Amount, decode(a1917,'J','J','2') As a0k11,rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)) Addr,a0803,a0807,a0813,na60, na02,A0822,A0823 " & _
              ",decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1, a1812,sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt from acc190, acc180, CUSTOMER, nation,acc080 where a1801=a1901 and a1908 is null and A1803<'Y' AND substr(a1803,1,8)=CU01(+) and substr(a1803,9,1)=CU02(+) and CU10=na01(+) and a1810 is not null and a0801=decode(a1917,'J','J','2')" & strSql & " group by a1803,a1901,na03,a1903,a1810,a1811,decode(a1917,'J','J','2'),a0803,a0807,a0813,na60, na02,A0822,A0823 " & _
              ",decode(a1812,'Y','1',decode(a1903,'USD','2','1')),a1812"

    '凡手續費為71:OUR集中在幣別群組的前方
    strSQLc = "select X.*,DECODE(A2219,'71:OUR',0,1) PKIND from (" & strSQLc & ") X,acc220 where a1803=a2201(+) and a1903=a2202(+) "
   '排序: 幣別、獨立水單以台幣結匯(排後面)、匯款方式(71:OUR排最前面)、代理人、收據公司別
   strSQLc = strSQLc & " Order By a1903,flag1 desc,pkind,a1803,a0k11,a1901 "
   
   ADO24m0.Open strSQLc, adoTaie, adOpenStatic, adLockReadOnly
   If ADO24m0.RecordCount = 0 Then
      ADO24m0.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   ii = 1
  
   If rsB.State <> adStateClosed Then rsB.Close
   Set rsB = Nothing
   Do While ADO24m0.EOF = False
      For intI = 1 To m_T4Num
          m_T4(intI) = ""
      Next intI
      bolUnit = False
      '受款地國別 =預設為代理人或客戶
      strNA60 = "" & ADO24m0.Fields("na60")
      '國家代號(受款行>代理人或客戶)
      strNA01 = "" & ADO24m0.Fields("fa10")
      '國家地區(受款行>代理人或客戶)
      strNA02 = "" & ADO24m0.Fields("na02")
      strA2220 = ""
      strA2222 = ""
      'Added by Lydia 2025/10/17
      strA2210 = ""
      strA2226 = ""
      
      '收據公司別
      strA0K11 = "" & ADO24m0.Fields("a0k11") '公司別
      '公司地址
      If "" & ADO24m0.Fields("A0822") & ADO24m0.Fields("A0823") <> "" Then
         m_T4(5) = ADO24m0.Fields("A0822") & " " & ADO24m0.Fields("A0823")
      Else
         m_T4(5) = "9F, No. 112, Sec. 2, Chang-An E. Rd., Taipei 104, Taiwan, R.O.C."
      End If
      '寰華結匯水單本所地址不能出現ROC
      If InStr("Y53374,", Left(Trim(ADO24m0.Fields("a1803")), 6)) > 0 Then
          m_T4(5) = Replace(m_T4(5), ", R.O.C.", String(8, " "))
      End If
      m_T4(5) = m_T4(5) & "      " & ADO24m0.Fields("a0813")  '合併電話

      '受款人:
         strExc(2) = "":    strExc(3) = "": strExc(4) = ""
         strExc(5) = "":    strExc(6) = "": strExc(7) = "": strExc(8) = ""
      '已限制匯款方式為2-電匯和5-合計電匯
         '抓受款人相關資料+A2217受款人(行)國別
         'Modified by Lydia 2025/07/21 抓受款人地址城市、受款人地址國別
         'StrSqlB = "Select a.*,b.NA60,b.NA02 From ACC220 a,NATION b Where a.A2217=b.NA01(+) and a2201='" & ADO24m0.Fields("a1803") & "' And a2202='" & ADO24m0.Fields("a1903") & "' "
         StrSqlB = "Select a.*,b.NA60,b.NA02,c.NA60 as na60x From ACC220 a,NATION b,Nation C Where a.A2217=b.NA01(+) and a.A2225=c.NA01(+) and a2201='" & ADO24m0.Fields("a1803") & "' And a2202='" & ADO24m0.Fields("a1903") & "' "
         intI = 1
         Set rsB = ClsLawReadRstMsg(intI, StrSqlB)
         If intI = 1 Then
            '受款地國別 =A2217受款人(行)國別
            If Not IsNull(rsB.Fields("NA60")) Then
               strNA60 = "" & rsB.Fields("NA60")
            End If
            '國家代號(受款行>代理人或客戶)
            If "" & rsB.Fields("a2217") <> "" Then
               strNA01 = "" & rsB.Fields("a2217")
            End If
            '國家地區(受款行>代理人或客戶)
            If "" & rsB.Fields("NA02") <> "" Then
               strNA02 = "" & rsB.Fields("NA02")
            End If
            
            strA2220 = "" & rsB.Fields("A2220")
            strA2222 = "" & rsB.Fields("A2222")
            'Added by Lydia 2025/10/17
            strA2210 = "" & rsB.Fields("A2210")
            strA2226 = "" & rsB.Fields("A2226")
            
            '受款銀行名稱：T414，原OR09
            If IsNull(rsB.Fields("a2208")) = False Then
               strExc(2) = RepEnter2zero(UCase(Trim(rsB.Fields("a2208"))))
            End If
            If IsNull(rsB.Fields("a2209")) = False Then
               strExc(3) = RepEnter2zero(UCase(Trim(rsB.Fields("a2209"))))
            End If
            '受款銀行帳號：T424，原OR13
            If IsNull(rsB.Fields("a2210")) = False Then
               strExc(4) = RepEnter2zero(UCase(Trim(rsB.Fields("a2210")) & " " & Trim(rsB.Fields("a2211"))))
            End If
            '受款銀行名稱：T414，原OR09有串a2210+a2211=strExc(4)
            m_T4(14) = strExc(2) & " " & strExc(3)
            '受款銀行 SWIFT CODE：T411，原OR10”受款行銀行編號”
            m_T4(11) = "" & UCase(Trim(rsB.Fields("a2211"))): m_T4(11) = RepEnter2zero(m_T4(11))
            
            '受款人帳號：T424，原OR13
            If IsNull(rsB.Fields("a2207")) = False Then
               m_T4(24) = RepEnter2zero(rsB.Fields("a2207"))
            End If

            '受款人名稱一：T418，原OR11
            m_T4(18) = RepEnter2zero(Trim("" & rsB.Fields("a2203") & " " & rsB.Fields("a2204"))) & IIf("" & rsB.Fields("a2205") <> "", " " & rsB.Fields("a2205"), "")
            
            '受款人地址：T420，原OR12
            'Modified by Lydia 2017/08/01 台銀要求全部都要印地址
             '客戶短地址
             If Not IsNull(rsB.Fields("a2218")) Then
                 m_T4(20) = RepEnter2zero(Trim("" & rsB.Fields("a2218")))
             Else
                 '原地址
                 m_T4(20) = RepEnter2zero(Trim("" & ADO24m0.Fields("addr")))
             End If
            m_T4(20) = Replace(m_T4(20), "#", "") 'Added by Lydia 2017/09/18 配合華銀不接受#,預設拿掉#
            'Added by Lydia 2025/07/21
            m_T4(21) = "" & rsB.Fields("A2224")  '受款人地址城市
            m_T4(22) = "" & rsB.Fields("NA60X")  '受款人地址國別
            If Trim(m_T4(22)) = "" Then m_T4(22) = strNA60
            'end 2025/07/21
         End If

      '受款行SWIFT格式：T410，原OR08
      If rsB.RecordCount = 0 Then
          m_T4(10) = "D"
      Else
         If UCase(Trim(rsB.Fields("a2210"))) = "SWIFT CODE" Then
             m_T4(10) = "A"
            'T410=A ,T411存A2211
             m_T4(11) = "" & UCase(Trim(rsB.Fields("a2211"))): m_T4(11) = RepEnter2zero(m_T4(11))
         Else
             m_T4(10) = "D"
             'Added by Lydia 2018/10/12 台銀表示, 用ABA匯款的話, 需要在"受款行銀行編號"(詳附件)加上FW
             If UCase("" & rsB.Fields("a2210")) = "ABA NO." Then
                 m_T4(11) = "FW" & m_T4(11)
             End If
         End If
      End If

      '金額：T425，原OR14 ~~ 靠右對齊,12字
      m_T4(25) = PUB_StrToStr(Format(ADO24m0.Fields("Amount"), "###0.00"), 12, True, True)
       
      '備註：T426，原OR15
      'Modified by Lydia 2018/04/23 寰華Y53374備註欄都不要再填寫(by 郭) ; 'Modified by Lydia 2020/04/22  +天津三元Y37580; +建毅Y51566,唯源Y52404
      If InStr("Y37580,Y53374,Y51566,Y52404,", Left(Trim(ADO24m0.Fields("a1803")), 6)) = 0 Then
         StrSqlB = "Select a1706 From acc190, acc170 Where a1902=a1702 And a1901='" & ADO24m0.Fields("a1901") & "' Group By a1706 Order By 1 "
      Else
         StrSqlB = "select null from dual"
      End If
      intI = 1
      Set rsB = ClsLawReadRstMsg(intI, StrSqlB)
      If intI = 1 Then
         intI = 0
         Do While Not rsB.EOF
            intI = intI + 1
            If intI > 12 Then Exit Do
            m_T4(26) = m_T4(26) & "," & rsB.Fields(0).Value
            m_T4(26) = m_T4(26) & IIf(ii = 12 And rsB.RecordCount > 12, " etc.", "")
            rsB.MoveNext
         Loop
      End If
      If m_T4(26) <> "" Then m_T4(26) = Mid(m_T4(26), 2)
      
      'Added by Lydia 2018/04/23 寰華Y53374備註欄都不要再填寫(by 郭) ; 'Modified by Lydia 2020/04/22  +天津三元Y37580; +建毅Y51566,唯源Y52404
      If InStr("Y37580,Y53374,Y51566,Y52404,", Left(Trim(ADO24m0.Fields("a1803")), 6)) = 0 Then
         If "" & ADO24m0.Fields("a1811") <> "1" And Val("" & ADO24m0.Fields("Ocnt")) = 0 Then 'Memo by Lydia 2018/06/26 非票匯和暫收款退費
           m_T4(26) = "INV." & m_T4(26)
         End If
      End If
      '媒體備註+備註
      If strA2222 <> "" Then m_T4(26) = PUB_GetSimpleName(strA2222, True, True) & " " & m_T4(26)
      
      'Memo by Lydia 2025/05/26 現匯款大陸不須CNAPS碼了---斯閔(已確認)

      m_T4(1) = Format(ii, "0000")
      m_T4(2) = String(1, " ")
      '匯款申請人證號：T403，原OR05
      m_T4(3) = convForm(RepEnter2zero(CheckStr("" & ADO24m0.Fields("a0807"))), 10)
      '匯款申請人名稱：T404，原OR03
      m_T4(4) = convForm(RepEnter2zero(CheckStr("" & ADO24m0.Fields("a0803"))), 70)
      '匯款申請人地址：T405，原OR04
      m_T4(5) = convForm(m_T4(5), 70)
      '匯款申請人地址/城市T：406
      m_T4(6) = convForm("TAIPEI", 28)
      '匯款申請人地址/國別：T407
      m_T4(7) = "TW"
      '受款地國別：T408，原OR06
      m_T4(8) = IIf(Len(Trim(strNA60)) = 0, "US", strNA60)
      '匯款類別：T409=19D
      m_T4(9) = Mid(Pub_DBtype, 1, 3)
      '注意雙字元切字會有奇數字元,無法insert; convForm+ PUB_StrToStr
      '受款行SWIFT格式：T410，原OR08=>前面已設定
      '受款銀行 SWIFT CODE=原”受款行銀行編號”=T411，原OR10
      m_T4(11) = convForm(PUB_StrToStr(m_T4(11), 11), 11)
      '受款銀行清算系統代碼：T412 此欄位於 114 年 11 月後啟用(待台銀通知)
      'Added by Lydia 2025/10/17 於11/17後上線新格式
      If Replace(MaskEdBox3.Text, "/", "") >= "1141117" Then
         '「受款銀行清算系統代號」和「清算編號」是針對 特定國家匯款編碼"ABA(美國)/BSB(澳洲)/CPA(加拿大)"，系統中維護於 "銀行資料維護中 受款銀行代號種類 ABA/Routing no."。Eg. 若系統 受款銀行代號 欄位為 "ABA"，填寫方式如下:
         ' 受款銀行清算 系統代碼>>USABA ；　受款銀行清算 編號>> 系統受款人帳號欄位資料 (目前暫無ABA以外的代理人) 。
         m_T4(12) = String(5, " ")
         '請填受款銀行所在地區之清算系統代碼 五碼 )EX ：美國的USABA 、澳洲的 AUBSB 、加拿大的 CACPA 等
         If InStr(UCase(Trim(strA2210)), "ABA CODE") > 0 Then
            m_T4(12) = "USABA" '美國
         ElseIf InStr(UCase(Trim(strA2210)), "CC CODE") > 0 Then
            m_T4(12) = "CACPA" '加拿大
         ElseIf InStr(UCase(Trim(strA2210)), "BSB CODE") > 0 Then
            m_T4(12) = "AUBSB" '澳洲
         End If
      Else
      'end 2025/10/17
          m_T4(12) = String(5, " ")
      End If
      '受款銀行清算編號：T413 此欄位於 114 年 11 月後啟用(待台銀通知)
      'Memo by Lydia 2025/10/17 清算編號與受款人帳號一致
      m_T4(13) = String(28, " ")

      '受款銀行名稱：T414，原OR09
      m_T4(14) = convForm(PUB_StrToStr(m_T4(14), 70), 70)
      '受款銀行地址：T415，原OR09；Swift Code可以不用輸入地址，另外有地址同時要輸入城市和國定，所以T415~T417全設定空白
      m_T4(15) = String(70, " ")
      m_T4(16) = String(28, " ")
      m_T4(17) = String(2, " ")
      
      '受款人名稱一：T418 --- 受款人名稱去掉中文(婉莘說:銀行資料的地址都已經設成英文)
      m_T4(18) = PUB_GetSimpleName(m_T4(18), True)
      '受款人地址：T420
      m_T4(20) = PUB_GetSimpleName(m_T4(20), True) '地址(去掉中文)

      '記錄列印用的名稱+地址
      strAR07 = Trim(m_T4(18)) & " ／ " & Trim(m_T4(20))
      '寬欄位提示
      If GetTextLength(strAR07) > 200 Then
         MsgBox "受款人代號：" & ADO24m0.Fields("a1803") & " 幣別：" & ADO24m0.Fields("a1903") & vbCrLf & "名稱 ／ 地址的長度超過200，清單會去掉尾部字元!"
         strAR07 = convForm(strAR07, 400)
      End If

      '受款人名稱一：T418，原OR11；名稱一和地址都限70字元
      m_T4(18) = convForm(PUB_StrToStr(m_T4(18), 70), 70)
      '受款人名稱二
      m_T4(19) = String(35, " ")
      '受款人地址
      'Added by Lydia 2025/07/21 沒地址不要帶入城市、國家
      'Mark by Lydia 2025/07/22 台銀:在8月前會先幫我們補資料
      'If Trim(m_T4(20)) = "" Or Trim(m_T4(21)) = "" Or Trim(m_T4(22)) = "" Then
      '   m_T4(20) = " "
      '   m_T4(21) = " "
      '   m_T4(22) = " "
      'End If
      ''end 2025/07/21
      'end 2025/07/22
      'Modified by Lydia 2025/09/10 +PUB_GetSimpleName
      m_T4(20) = convForm(PUB_StrToStr(PUB_GetSimpleName(m_T4(20), , True, True), 70), 70)
      '受款人地址城市，國別 --- Modified by Lydia 2025/07/21 從預設空白改成有資料 ---Modified by Lydia 2025/09/10 +PUB_GetSimpleName
      m_T4(21) = convForm(PUB_StrToStr(PUB_GetSimpleName(m_T4(21), , True), 28), 28)
      m_T4(22) = convForm(PUB_StrToStr(PUB_GetSimpleName(m_T4(22), , True), 2), 2)
      'end 2025/07/21
      
      '代理Y25061430名稱T418=C / Zurbano 76, 7 ° Madrid 28010 SPAIN,  經過模組只有69個字元
      If Len(m_T4(18)) < 70 Then
         m_T4(18) = Mid(m_T4(18) & String(70, " "), 1, 70)
      End If
      If Len(m_T4(20)) < 70 Then
         m_T4(20) = Mid(m_T4(20) & String(70, " "), 1, 70)
      End If
      
      '受款人帳號為IBAN：T423 此欄位於 114 年 11 月後啟用(待台銀通知)
      'Added by Lydia 2025/10/17 (於11/17後上線新格式) 是IBAN才要顯示Y
      If Replace(MaskEdBox3.Text, "/", "") >= "1141117" And strA2226 = "Y" Then
          m_T4(23) = "Y"
      Else
      'end 2025/10/17
          m_T4(23) = String(1, " ")
      End If
      '受款人帳號：T424，原OR13
      m_T4(24) = convForm(PUB_StrToStr(m_T4(24), 34), 34)
      '金額：T425，原OR14 ~~ 靠右對齊,12字=>前面已設定
      '備註：T426，原OR15 ----Modified by Lydia 2025/09/10 +PUB_GetSimpleName
      m_T4(26) = convForm(PUB_StrToStr(RepEnter2zero(PUB_GetSimpleName(CheckStr(m_T4(26)), , True, True)), 140), 140)
      If Len(Trim(m_T4(26))) > 135 Then
         If InStr(Mid(m_T4(26), 134, 6), "etc.") > 0 Then
           m_T4(26) = convForm(PUB_StrToStr(m_T4(26), 140), 140)
         Else
           m_T4(26) = convForm(PUB_StrToStr(m_T4(26), 135) & " etc.", 140)
         End If
      End If
      '保留欄位：T427
      m_T4(27) = String(200, " ")
      
      '台一編號(Y編號代理人) =原OR16
      strAR06 = convForm(PUB_StrToStr(ADO24m0.Fields("A1803"), 20), 20)
      
      '獨立水單不合併明細
      If ADO24m0.Fields("a1812") = "Y" Then bolUnit = True
      
      '合併明細資料 'decode(substr(or04,1,2),'7F','9','10','1','4F','J','9F','2') 公司別, or m_T4(3) 匯款人證號
      '判斷AR05獨立水單
      StrSqlB = "Select * From accrpt24m0_T114  where UNO='" & strUserNum & "' and ar00='" & ADO24m0.Fields("a1803") & "' " & _
                "And ar01='" & ADO24m0.Fields("a1903") & "' And ar03='" & ADO24m0.Fields("flag1") & "' and ar04='" & strA0K11 & "' and ar05 is null "
      strExc(1) = "":          intI = 1
      Set rsB = ClsLawReadRstMsg(intI, StrSqlB)
      If intI = 1 And bolUnit = False Then
         '串D/B note
         If Trim(rsB.Fields("T426")) <> Trim(m_T4(26)) Then
            strExc(2) = convForm(Trim(rsB.Fields("T426")) & "," & Trim(m_T4(26)), 140)
            If Len(Trim(strExc(2))) > 135 Then
               If InStr(Mid(strExc(2), 134, 6), "etc.") > 0 Then
               strExc(2) = convForm(PUB_StrToStr(strExc(2), 140), 140)
             Else
               strExc(2) = convForm(PUB_StrToStr(strExc(2), 135) & " etc.", 140)
               End If
            End If
            strExc(1) = ",T426=" & CNULL(strExc(2))
         End If
         strExc(1) = "update accrpt24m0_T114 set T425=T425+" & Val(m_T4(25)) & strExc(1)
         strExc(1) = strExc(1) & " where UNO='" & strUserNum & "' and ar00='" & ADO24m0.Fields("a1803") & "' And ar01='" & ADO24m0.Fields("a1903") & "' " & _
                    "And ar03='" & ADO24m0.Fields("flag1") & "' and ar04='" & strA0K11 & "' and ar05 is null "
      Else
         'PK ~ UNO(使用者),AR00(代理人編號)+AR01(幣別)+AR02(代理人名稱)+AR03(flag1)+AR04(A1917公司別) , AR05獨立水單=a1812, AR06=台一編號(Y編號代理人), AR07名稱+地址合併列印
         strExc(1) = "INSERT INTO ACCRPT24m0_T114 values(" & CNULL(strUserNum) & "," & CNULL(ChgSQL("" & ADO24m0.Fields("a1803"))) & "," & CNULL(ChgSQL("" & ADO24m0.Fields("a1903"))) & _
                     "," & CNULL(convForm(PUB_StrToStr(ChgSQL("" & ADO24m0.Fields("fa05")), 30), 30)) & "," & CNULL(ADO24m0.Fields("flag1")) & "," & CNULL(strA0K11) & "," & CNULL("" & ADO24m0.Fields("a1812")) & _
                     "," & CNULL(ChgSQL(strAR06)) & "," & CNULL(ChgSQL(strAR07))
         '台銀資料欄位
         For intI = 1 To m_T4Num
             strExc(1) = strExc(1) & "," & CNULL(ChgSQL(m_T4(intI)))
         Next intI
         strExc(1) = strExc(1) & ")"
         ii = ii + 1
      End If
      
      adoTaie.Execute strExc(1)
      strA1901List = strA1901List & ADO24m0.Fields("a1901") & "," '記錄付款單號
         
      ADO24m0.MoveNext
   Loop
   
   ADO24m0.Close
   If rsB.State <> adStateClosed Then rsB.Close
   Set rsB = Nothing
   
'更新付款單的匯款方式=>5.台銀合併結匯
   If strA1901List <> "" Then
      PUB_UpdateA1811toType strA1901List
   End If
   Exit Sub
   
ErrHandle:
   If Err.Number <> 0 Then MsgBox Err.Description
End Sub

'Added by Lydia 2025/05/27
Private Sub PStoreRead_T114(Optional ByVal aKind As String)   '讀取明細暫存---台銀114年8月上線
Dim ff As Integer, strPath As String
Dim strR As String
Dim rsR As New ADODB.Recordset
Dim strFileNo As String, strFileName As String '檔名
Dim Id As Integer
Dim bolPrint As Boolean '開始列印
Dim iCall As Integer '結匯轉出的TXT檔案要產出2個檔案

    If rsR.State <> adStateClosed Then rsR.Close
    Set rsR = Nothing

    'Added by Morgan 2025/7/16
    'Mark by Lydia 2025/07/21 當時請假; 已在PStoreData_T114處理
    'strSql = "update accrpt24m0_T114 set (T421,T422)=(select rpad(a2224,28,' ') T421,rpad(na60,2,' ') T422 from acc220,nation where a2201=ar00 and a2202=ar01 and na01(+)=a2225) where UNO='" & strUserNum & "'"
    'adoTaie.Execute strSql, intI
    ''end 2025/7/16
    'end 2025/07/21
    
    strR = "Select a.*,b.a2219,(b.a2214||' '||b.a2215||' '||b.a2216) midbk From accrpt24m0_T114 a, acc220 b " & _
           "where ar00=a2201(+) and ar01=a2202(+) and UNO='" & strUserNum & "' order by T401 "
    intI = 1
    Set rsR = ClsLawReadRstMsg(intI, strR)
    If intI = 1 Then
        strPath = PUB_Getdesktop
        strPath = strPath & "\台銀結匯水單"
        If Dir(strPath, vbDirectory) = "" Then
           MkDir strPath
        End If
        
StartPrt:
        strFileNo = "": pStr(0) = ""
        If bolPrint Then
           rsR.MoveFirst
           '設定印表機
            Printer.EndDoc
            Printer.Orientation = 2 '1.直印 2.橫印
            Printer.PaperSize = 9  'A4
               
            lngPageHeight = Printer.ScaleHeight
            lngPageWidth = Printer.ScaleWidth
            lngLineHeight = 270
            GetPleft_T104 '設定邊界
            SetColumnName_T104
            iPage = 0
        Else
           idx = 0
        End If
        
        '台銀要求媒體TXT檔案拿掉最後面的欄位(台一編號)；所以改成結匯轉出的TXT檔案要產出2個檔案，一個有(台一編號)欄位，一個沒有(台一編號)欄位的檔案名稱在尾端多一個"X"。
        For iCall = 1 To IIf(bolPrint = False, 2, 1)
            strFileNo = "": pStr(0) = ""
            rsR.MoveFirst
            Do While rsR.EOF = False
                '幣別+結匯日期+結購類型
                If strFileNo = "" Or strFileNo <> rsR.Fields("ar01") & strDBdate & GetFlagTitle("", rsR.Fields("ar03")) Then
                    pStr(0) = GetFlagTitle(rsR.Fields("ar01"), rsR.Fields("ar03"))
                    strFileNo = GetFlagTitle(rsR.Fields("ar01") & strDBdate, rsR.Fields("ar03"))
                    Id = 1
                    If bolPrint Then
                         iPage = iPage + 1
                        If iPage > 1 Then '換頁
                           PrintSubTotal_T104
                           PrintSign
                           Printer.NewPage
                        End If
                        iPage = 1: mCount = 0: mTotal = 0  '以各分類為準
                        PrintHeader_T104
                    Else
                        idx = idx + 1
                        If ff > 0 Then Close #ff
                        ff = FreeFile
                        
                        If iCall = 2 Then '去掉(台一編號)
                            strFileName = strPath & "\" & strFileNo & ".txt"
                        Else
                            strFileName = strPath & "\" & strFileNo & "_X.txt" '保留(台一編號)
                        End If

                        Open strFileName For Output As ff

                        'TXT：首筆資料為檔案欄位名稱說明，實際匯款明細資料請自第二筆（含）以後開始寫檔
                        'T401~T410
                        strExc(0) = "編號" & String(1, " ") & convForm("匯款人號", 10) & convForm("匯款申請人名稱", 70) & convForm("匯款申請人地址", 70) & convForm("城市", 28) & "國" & "國" & String(4, " ")
                        'T411~T420
                        strExc(0) = strExc(0) & convForm("受款行BIC", 11) & convForm("清算碼 編號", 33) & convForm("受款銀行名稱", 70) & convForm("受款銀行地址", 70) & convForm("城市", 28) & "國" & convForm("受款人名稱一", 70) & convForm("受款人名稱二", 35) & convForm("受款人地址", 70)
                        'T421~T427
                        strExc(0) = strExc(0) & convForm("城市", 28) & "國" & String(1, " ") & convForm("受款人帳號", 34) & convForm("匯款金額", 12) & convForm("備註", 140) & convForm("保留欄位", 200)
                        If iCall = 1 Then strExc(0) = strExc(0) & convForm("台一編號", 20)
                        
                        Print #ff, strExc(0)
                    End If
                End If
             
                m_T4(1) = Format(Id, "0000")
                strExc(1) = m_T4(1)
                For intI = 2 To m_T4Num
                   If intI = 25 Then
                      m_T4(intI) = PUB_StrToStr(Format("" & rsR.Fields("T4" & Format(intI, "00")), "###0.00"), 12, True, True)
                   Else
                      m_T4(intI) = "" & rsR.Fields("T4" & Format(intI, "00"))
                   End If
                   '只列印不匯出TXT：受款銀行地址+城市+國別、受款人地址+城市+國別; 2025/07/03 增加受款行名稱14(已經提供swift code，就無需再提供受款銀行名稱)
                   'Modified by Morgan 2025/7/14 受款人地址+城市+國別(20,21,22) 改也要匯出--斯閔; Memo by Lydia 2025/07/21 當時請假
                   If InStr("15,16,17,14", Format(intI, "00")) > 0 Then
                       strExc(1) = strExc(1) & convForm(" ", Len(m_T4(intI)))
                   Else
                       strExc(1) = strExc(1) & m_T4(intI)
                   End If
                Next intI
                '台一編號
                If iCall = 1 Then
                   strExc(1) = strExc(1) & rsR.Fields("AR06")
                End If
                
                If bolPrint Then
                    pStr(1) = m_T4(1)
                    '匯款人證號：T403，原OR05
                    pStr(2) = m_T4(3)
                    '受款地國別：T408，原OR06
                    pStr(3) = m_T4(8)
                    '受款行SWIFT格式：T410，原OR08
                    pStr(4) = m_T4(10)
                    '受款銀行名稱：T414，原OR09(收款行名稱及地址)，若為SWIFT格式改放受款銀行 SWIFT CODE
                    If m_T4(10) = "A" Then
                       pStr(5) = m_T4(11)
                    Else
                       pStr(5) = convForm(PUB_StrToStr(m_T4(14), 70), 70)
                    End If
                    '匯款人帳號及匯款金額要全印
                    '匯款人帳號：T424，原OR13
                    pStr(6) = convForm(m_T4(24), 34)
                    '匯款金額：T425，原OR14
                    pStr(7) = PUB_StrToStr(Format(Val(m_T4(25)), "##,##0.00"), 14, True, True)
                    '手續費
                    'Added by Lydia 2019/10/03 匯款日幣, 以OUR方式結匯的, OUR要改成全額到行
                    If rsR.Fields("ar01") = "JPY" And UCase("" & rsR.Fields("a2219")) = "71:OUR" Then
                        pStr(8) = "全額到行" '僅供台銀承辦人員查看,不改變媒體
                    Else
                        pStr(8) = convForm("" & rsR.Fields("a2219"), 7)
                    End If
    
                    '台一編號(Y編號代理人)
                     pStr(9) = convForm("" & rsR.Fields("AR06"), 20)
                    '名稱+地址合併列印
                    pStr(10) = Trim("" & rsR.Fields("AR07"))
                    pStr(11) = ""
                    
                    '第3行 中間銀行資料
                    If Len(Trim(rsR.Fields("midbk"))) > 0 Then
                       pStrL3 = "中間銀行：" & rsR.Fields("midbk")
                    Else
                       pStrL3 = ""
                    End If
                    mCount = mCount + 1
                    mTotal = mTotal + Val(m_T4(25))
                    PrintDetail_T104
                Else
                    Print #ff, strExc(1)
                End If
               Id = Id + 1
               rsR.MoveNext
            Loop
               
            If bolPrint Then GoTo EndPrt
            
            Close #ff
        Next iCall
        
        If TxtList <> "Y" Then
           rsR.Close
           'Modified by Lydia 2025/06/25 idx + (idx/2) (舊TXT)
           MsgBox "已在桌面的台銀結匯水單資料夾產生 " & idx + (idx / 2) & " 個檔案!!", vbInformation
        End If
        
        If TxtList = "Y" And bolPrint = False Then
           bolPrint = True
           GoTo StartPrt
        End If
    End If
'------------------------
EndPrt:
   If bolPrint Then
        PrintSubTotal_T104
        PrintSign
        Printer.EndDoc
        'Modified by Lydia 2025/06/25 拿掉「 " & idx & " 個」
        MsgBox "已在桌面的台銀結匯水單資料夾產生檔案, 並列印清單完成!!", vbInformation
        rsR.Close
   End If
End Sub

'Added by Lydia 2025/06/26 產生華銀媒體暫存檔：華銀格式也從MT格式(不需城市+國家)改成MX格式
Private Sub PStoreData_H114()
'若結匯規則有變更，請加註文件：\\LINUX\PolyCOM\TaieNew\電腦中心日常工作\結匯-預設匯款方式(a1811和媒體檔).doc
Dim strSql As String, StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim strSQLc As String, strA0K11 As String
Dim ii As Integer
Dim inR As Integer '序號
Dim atJ As Integer '下一欄位序號
Dim bolUnit As Boolean '獨立水單資料不合併計算
Dim strPrtNameAddr As String '記錄列印用的名稱+地址
Dim strA2220 As String 'CNAPS(大陸匯款)
Dim strA2219 As String '手續費方式
Dim strCurr  As String '目前幣別
Dim strA1a12  As String '目前幣別的匯率議價編號
Dim tmpStr(1 To 4) As String   '合併-付款明細
Dim strA2222 As String '記錄媒體備註
Dim bolAdd As Boolean '是否增加序號

   adoTaie.Execute " delete from accrpt24m0_H114 where UNO='" & strUserNum & "' "

   '結匯日期
   strDBdate = Replace(FCDate(MaskEdBox3.Text), "/", "")
   '付款日期
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a1802 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a1802 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   '代理人
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and a1803 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a1803 <= '" & Text2 & "'"
   End If
   '付款單號
   If Text4 <> MsgText(601) Then
      strSql = strSql & " and a1801 >= '" & Text4 & "'"
   End If
   If Text3 <> MsgText(601) Then
      strSql = strSql & " and a1801 <= '" & Text3 & "'"
   End If
   '公司別
   strSql = strSql & " and a1917='J'" '限Ｊ公司
   
   strSql = strSql & " and a1811<>'6' " '排除匯款方式6-抵帳
   
   excelSql = strSql '傳結匯清單SQL條件 (全電匯+票匯)
   
   '限電匯
   strSql = strSql & " and a1811=2"
   
' 明細
    '與excelsav2 可能有差異
    ADO24m0.CursorLocation = adUseClient
    '排序:幣別、代理人、匯款方式、收據公司別
    '代理人(Y 編號)
    'Modified by Lydia 2017/09/30 華銀整批媒體RMB改CNY a1903=> DECODE(A1903,'RMB','" & J_RMB & "',A1903) A1903
    strSQLc = "select a1901, a1803, DECODE(A1903,'RMB','" & J_RMB & "',A1903) A1903, a1917 As a0k11,decode(a1810,null,Decode(fa05, Null, Nvl(FA04, FA06), FA05),substr(a1810,1,30)) As FA05, " & _
              "decode(a1810,null,Decode(FA05, Null, Null, fa63),substr(a1810,31,30)) As FA63, decode(a1810,null,Decode(FA05, Null, Null, fa64),substr(a1810,61,30)) As FA64, " & _
              "decode(a1810,null,Decode(FA05, Null, Null, fa65),substr(a1810,91,30)) As FA65, fa10, na03, na02,nvl(na80,na04) na04, a1810, a1811,a1812, sum(a1904) as Amount, " & _
              "sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt,rtrim(max(fa18||' '||fa19||' '||fa20||' '||fa21||' '||fa22||' '||fa70)) Addr,a0803,a0807,a0813,na60,A0822,A0823 " & _
              ",decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1 from acc190, acc180, fagent, nation,acc080 where a1801=a1901 and a1908 is null and A1803>'Y' " & _
              "AND substr(a1803,1,8)=fa01 and substr(a1803,9,1)=fa02 and fa10=na01(+) and a0801=a1917" & strSql & _
              " group by a1901,a1803, decode(a1903,'RMB','" & J_RMB & "',a1903), a1917, decode(a1810,null,Decode(fa05, Null, Nvl(FA04, FA06), FA05),substr(a1810,1,30))," & _
              " decode(a1810,null,Decode(FA05, Null, Null, fa63),substr(a1810,31,30)), decode(a1810,null,Decode(FA05, Null, Null, fa64),substr(a1810,61,30))," & _
              " decode(a1810,null,Decode(FA05, Null, Null, fa65),substr(a1810,91,30)), fa10, na03,na02,nvl(na80,na04), a1810, a1811,a1812,a0803,a0807,a0813,na60,A0822,A0823," & _
              " decode(a1812,'Y','1',decode(a1903,'USD','2','1'))"
    '客戶(X 編號)
    'Modified by Lydia 2017/09/30 華銀整批媒體RMB改CNY a1903=> DECODE(A1903,'RMB','" & J_RMB & "',A1903) A1903
    strSQLc = strSQLc & " Union select a1901, a1803, DECODE(A1903,'RMB','" & J_RMB & "',A1903) A1903, a1917 As a0k11,decode(a1810,null,Decode(CU05, Null, Nvl(CU04, CU06), CU05),substr(a1810,1,30)) As FA05, " & _
              "decode(a1810,null,Decode(CU05, Null, Null, CU88),substr(a1810,31,30)) As FA63, decode(a1810,null,Decode(CU05, Null, Null, CU89),substr(a1810,61,30)) As FA64, " & _
              "decode(a1810,null,Decode(CU05, Null, Null, CU90),substr(a1810,91,30)) As FA65, CU10 AS FA10, na03,na02,nvl(na80,na04) na04, a1810, a1811,a1812, sum(a1904) as Amount, " & _
              "sum(decode(substr(a1902,1,1),'O',1,0)) Ocnt,rtrim(max(cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102)) Addr,a0803,a0807,a0813,na60,A0822,A0823 " & _
              ",decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1 from acc190, acc180, customer, nation,acc080 where a1801=a1901 and a1908 is null and A1803<'Y' " & _
              "AND substr(a1803,1,8)=CU01 and substr(a1803,9,1)=CU02 and CU10=na01(+) and a0801=a1917" & strSql & _
              " group by a1901,a1803, decode(a1903,'RMB','" & J_RMB & "',a1903), a1917,  decode(a1810,null,Decode(CU05, Null, Nvl(CU04, CU06), CU05),substr(a1810,1,30))," & _
              " decode(a1810,null,Decode(CU05, Null, Null, CU88),substr(a1810,31,30)), decode(a1810,null,Decode(CU05, Null, Null, CU89),substr(a1810,61,30))," & _
              " decode(a1810,null,Decode(CU05, Null, Null, CU90),substr(a1810,91,30)), CU10, na03,na02,nvl(na80,na04), a1810, a1811,a1812,a0803,a0807,a0813,na60,A0822,A0823," & _
              " decode(a1812,'Y','1',decode(a1903,'USD','2','1'))"
    '排序:幣別、獨立水單、代理人、匯款方式、收據公司別
    strSQLc = strSQLc & " order by a1903,flag1 desc,a1803,a0k11 "

   ADO24m0.Open strSQLc, adoTaie, adOpenStatic, adLockReadOnly
   If ADO24m0.RecordCount = 0 Then
      ADO24m0.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   
   ii = 1
   inR = 1
   strCurr = "" '幣別:預設空白
   
   If rsB.State <> adStateClosed Then rsB.Close
   Set rsB = Nothing
   Do While ADO24m0.EOF = False
      For intI = 0 To m_H4Num
          m_H4(intI) = ""
      Next intI
      bolUnit = False
      strA0K11 = "" & ADO24m0.Fields("a0k11")  '收據公司別
      strPrtNameAddr = ""
      strA2220 = ""
      strA2219 = "SHA" '手續費空白,預設SHA
      
      '序號(用在清單列印)
      m_H4(0) = Format(inR, "0000")
      '(申請人)統一編號：與106年X(10)不同
      m_H4(1) = convForm("" & ADO24m0.Fields("a0807"), 12)
      '(申請人)名稱1~2
      strExc(1) = convForm(CheckFstSpec(PUB_GetSimpleName("" & ADO24m0.Fields("a0803"), True)), 70)
      m_H4(2) = Mid(strExc(1), 1, 35)
      m_H4(3) = Mid(strExc(1), 36, 35)
      '(申請人)地址1~2
      strExc(2) = ADO24m0.Fields("a0822") & " " & ADO24m0.Fields("a0823")
      If Trim(strExc(2)) = "" Then
         strExc(2) = "9F, No. 112, Sec. 2, Chang-An E. Rd., Taipei 104, Taiwan, R.O.C." '預設地址
      End If
      '寰華結匯水單本所地址不能出現ROC
      If InStr("Y53374,", Left(Trim(ADO24m0.Fields("a1803")), 6)) > 0 Then
          strExc(2) = Replace(strExc(2), ", R.O.C.", String(8, " "))
      End If
      strExc(2) = convForm(CheckFstSpec(PUB_GetSimpleName(strExc(2), True)), 70)
      m_H4(4) = Mid(strExc(2), 1, 35)
      m_H4(5) = Mid(strExc(2), 36, 35)
      '申請人地址/城市
      m_H4(6) = convForm("TAIPEI", 35)
      m_H4(7) = "TW"
      '匯款人身分別
      m_H4(8) = " " '空白
      '匯款模式類別 => P：付款Pay R：匯款 Remit(預設值)
      m_H4(9) = "R"
      '匯款幣別
      m_H4(10) = convForm("" & ADO24m0.Fields("a1903"), 3)
      '匯款金額
      m_H4(11) = PUB_StrToStr(Format(ADO24m0.Fields("Amount"), "###0.00"), 15, True, True)
      '受款地區國別代號
      m_H4(12) = convForm("" & ADO24m0.Fields("na60"), 2)
      '受款地區國別名稱
      m_H4(13) = convForm("" & ADO24m0.Fields("na04"), 20)
      '匯款性質編號--預設19D
      m_H4(14) = Mid(Pub_DBtype, 1, 3)
      '匯款性質編號其他補充說明 --設空白
      m_H4(15) = String(35, " ")
      '受款人身分別--預設3.民間
      m_H4(16) = "3"
      
      '受益人和設帳銀行的名稱的名稱和地址
      strExc(5) = "":   strExc(6) = ""
      If IsNull(ADO24m0.Fields("fa05")) = False Then
         strExc(5) = RepEnter2zero(UCase(Trim(ADO24m0.Fields("fa05"))))
      End If
      If IsNull(ADO24m0.Fields("fa63")) = False Then
         strExc(5) = strExc(5) & " " & RepEnter2zero(UCase(Trim(ADO24m0.Fields("fa63"))))
      End If
      If IsNull(ADO24m0.Fields("fa64")) = False Then
         strExc(5) = strExc(5) & " " & RepEnter2zero(UCase(Trim(ADO24m0.Fields("fa64"))))
      End If
      If IsNull(ADO24m0.Fields("fa65")) = False Then
         strExc(5) = strExc(5) & " " & RepEnter2zero(UCase(Trim(ADO24m0.Fields("fa65"))))
      End If
      If IsNull(ADO24m0.Fields("Addr")) = False Then
         strExc(6) = RepEnter2zero(UCase(Trim(ADO24m0.Fields("Addr"))))
      End If
      '抓受款人相關資料
      strExc(7) = "": strExc(8) = ""
      strExc(1) = "": strExc(2) = ""
      'Modified by Lydia 2017/10/03 華銀整批媒體RMB改CNY
      'Modified by Lydia 2025/07/21 抓受款人地址城市、受款人地址國別
      'StrSqlB = "Select a.*,b.NA60,nvl(b.na04,b.na80) na04 From ACC220 a,NATION b Where a.A2217=b.NA01(+) and a2201='" & ADO24m0.Fields("a1803") & "' And A2202='" & IIf(ADO24m0.Fields("a1903") = J_RMB, "RMB", ADO24m0.Fields("a1903")) & "' "
      StrSqlB = "Select a.*,b.NA60,nvl(b.na04,b.na80) na04,c.NA60 as na60x From ACC220 a,NATION b, Nation C Where a.A2217=b.NA01(+) and a.A2225=c.NA01(+) and a2201='" & ADO24m0.Fields("a1803") & "' And A2202='" & IIf(ADO24m0.Fields("a1903") = J_RMB, "RMB", ADO24m0.Fields("a1903")) & "' "
      intI = 1
      Set rsB = ClsLawReadRstMsg(intI, StrSqlB)
      If intI = 1 Then
        '交易國別代號 =A2217受款人(行)國別
        If "" & rsB.Fields("NA60") <> "" Then
           m_H4(12) = convForm("" & rsB.Fields("NA60"), 2)
        End If
        '國家名稱(受款行>代理人或客戶)
        If "" & rsB.Fields("NA04") <> "" Then
           m_H4(13) = convForm("" & rsB.Fields("NA04"), 20)
        End If
        '(受益人)帳號
        m_H4(17) = convForm(CheckFstSpec(PUB_GetSimpleName("" & rsB.Fields("A2207"), True)), 35)
        '(受益人)名稱
        If "" & rsB.Fields("A2203") <> "" Then
           strExc(5) = RepEnter2zero(UCase(Trim(rsB.Fields("A2203"))))
        End If
        If "" & rsB.Fields("A2204") <> "" Then
           strExc(5) = strExc(5) & " " & RepEnter2zero(UCase(Trim(rsB.Fields("A2204"))))
        End If
        If "" & rsB.Fields("A2205") <> "" Then
           strExc(5) = strExc(5) & " " & RepEnter2zero(UCase(Trim(rsB.Fields("A2205"))))
        End If
        '短地址(優先)
        If "" & rsB.Fields("A2218") <> "" Then
           strExc(6) = RepEnter2zero(UCase(Trim(rsB.Fields("A2218"))))
        End If
        '(設帳銀行)swift代號 =A2211 受款銀行代號
        m_H4(27) = convForm("" & rsB.Fields("A2211"), 12)
        '(設帳銀行)名稱
        If "" & rsB.Fields("A2208") <> "" Then
           strExc(7) = RepEnter2zero(UCase(Trim(rsB.Fields("A2208"))))
        End If
        If "" & rsB.Fields("A2209") <> "" Then
           strExc(7) = strExc(7) & " " & RepEnter2zero(UCase(Trim(rsB.Fields("A2209"))))
        End If

        '(中間銀行)
        If "" & rsB.Fields("A2214") <> "" Then
           If InStr(UCase(rsB.Fields("A2214")), "SWIFT") = 0 Then
              strExc(8) = strExc(8) & IIf(strExc(8) <> "", " ", "") & RepEnter2zero(UCase(Trim(rsB.Fields("A2214"))))
           Else  '抓Swift Code
              strExc(0) = Replace(UCase(rsB.Fields("A2214")), "：", ":")
              m_H4(35) = convForm(RepEnter2zero(Trim(Mid(strExc(0), InStr(strExc(0), ":") + 1))), 12)
           End If
        End If
        If "" & rsB.Fields("A2215") <> "" Then
           If InStr(UCase(rsB.Fields("A2215")), "SWIFT") = 0 Then
              strExc(8) = strExc(8) & IIf(strExc(8) <> "", " ", "") & RepEnter2zero(UCase(Trim(rsB.Fields("A2215"))))
           Else
              strExc(0) = Replace(UCase(rsB.Fields("A2215")), "：", ":")
              m_H4(35) = convForm(RepEnter2zero(Trim(Mid(strExc(0), InStr(strExc(0), ":") + 1))), 12)
           End If
        End If
        If "" & rsB.Fields("A2216") <> "" Then
           If InStr(UCase(rsB.Fields("A2216")), "SWIFT") = 0 Then
              strExc(8) = strExc(8) & IIf(strExc(8) <> "", " ", "") & RepEnter2zero(UCase(Trim(rsB.Fields("A2216"))))
           Else
              strExc(0) = Replace(UCase(rsB.Fields("A2216")), "：", ":")
              m_H4(35) = convForm(RepEnter2zero(Trim(Mid(strExc(0), InStr(strExc(0), ":") + 1))), 12)
           End If
        End If
        '手續費
        If "" & rsB.Fields("A2219") <> "" Then
           strA2219 = UCase(rsB.Fields("A2219"))
        End If
        strA2222 = "" & rsB.Fields("A2222") '媒體備註
        'Added by Lydia 2025/07/21
        strExc(1) = "" & rsB.Fields("A2224") '受款人地址城市
        strExc(2) = "" & rsB.Fields("NA60X") '受款人地址國別
        If Trim(strExc(2)) = "" Then strExc(2) = m_H4(12) '改受款銀行地址國別
        'end 2025/07/21
      End If
      
      '收款人名稱和地址去掉中文(婉莘說:銀行資料的地址都已經設成英文)
      strExc(5) = PUB_GetSimpleName(strExc(5), True) '暫存收款人名稱
      strExc(6) = PUB_GetSimpleName(strExc(6), True) '暫存收款人地址
      strExc(5) = UCase(strExc(5)): strExc(6) = UCase(strExc(6)) 'Swift文數字規定:大寫英文
       
      '記錄列印用的名稱+地址
      strPrtNameAddr = strExc(5) & " ／ " & strExc(6)
      
      '受益人幣別
      m_H4(18) = m_H4(10)
      '受益人統編
      m_H4(19) = convForm(" ", 12)
      '1.處理受益人名稱和地址
      strExc(5) = convForm(CheckFstSpec(PUB_GetSimpleName(strExc(5), True)), 70)
      m_H4(20) = Mid(strExc(5), 1, 35)
      m_H4(21) = Mid(strExc(5), 36, 35)
      'Added by Lydia 2025/07/21 沒地址不要帶入城市、國家 'Memo by Lydia 2025/07/29 與台銀不同，華銀若沒有城市就都不要帶
      If Trim(strExc(6)) = "" Or Trim(strExc(1)) = "" Or Trim(strExc(2)) = "" Then
         strExc(6) = " "
         strExc(1) = " "
         strExc(2) = " "
      End If
      'end 2025/07/21
      'Modified by Lydia 2025/09/10
      'strExc(6) = convForm(CheckFstSpec(PUB_GetSimpleName(strExc(6), True)), 70)
      strExc(6) = convForm(CheckFstSpec(PUB_GetSimpleName(strExc(6), , True, True)), 70)
      m_H4(22) = Mid(strExc(6), 1, 35)
      m_H4(23) = Mid(strExc(6), 36, 35)
      'Modified by Lydia 2025/07/21 從預設空白改成有資料
      '(受益人)城市
      'Modified by Lydia 2025/09/10
      'm_H4(24) = convForm(CheckFstSpec(PUB_GetSimpleName(strExc(1), True)), 35)
      m_H4(24) = convForm(CheckFstSpec(PUB_GetSimpleName(strExc(1), , True, True)), 35)
      '(受益人)國家
      m_H4(25) = convForm(CheckFstSpec(PUB_GetSimpleName(strExc(2), True)), 2)
      'end 2025/07/21
      '(設帳銀行)swift/fisc
      m_H4(26) = "S"
      '(設帳銀行)SWIFT代號：m_H4(27)前面已設定
      
      '2.處理設帳銀行(受款銀行)名稱,地址預設為空白
      strExc(0) = convForm(CheckFstSpec(PUB_GetSimpleName(strExc(7), True)), 70)
      m_H4(28) = Mid(strExc(0), 1, 35)
      m_H4(29) = Mid(strExc(0), 36, 35)
      m_H4(30) = convForm(" ", 35)
      m_H4(31) = convForm(" ", 35)
      '(設帳銀行)城市 => 空白
      m_H4(32) = convForm(" ", 35)
      '(設帳銀行)國家 => 空白
      m_H4(33) = convForm(" ", 2)
      '清算系統號碼(國別代號+號碼) => 空白
      m_H4(34) = convForm(" ", 35)
      '(中間銀行)SWIFT CODE
      m_H4(35) = CheckFstSpec(m_H4(35))
      If Len(m_H4(35)) <> 12 Then
         m_H4(35) = convForm(" ", 12)
      End If
      '(中間銀行)名稱及地址 => 空白
      m_H4(36) = convForm(" ", 35)
      m_H4(37) = convForm(" ", 35)
      m_H4(38) = convForm(" ", 35)
      m_H4(39) = convForm(" ", 35)
      m_H4(40) = convForm(" ", 35)
      m_H4(41) = convForm(" ", 2)
      
      '費用明細
      m_H4(42) = "SHA" '預設
      
      '付款明細1~4：抓Debit Note
        strExc(9) = ""
        'Modified by Lydia 2018/04/23 寰華Y53374備註欄都不要再填寫(by 郭) ; 'Modified by Lydia 2020/04/22  +天津三元Y37580; +建毅Y51566,唯源Y52404
        If InStr("Y37580,Y53374,Y51566,Y52404,", Left(Trim(ADO24m0.Fields("a1803")), 6)) = 0 Then
           StrSqlB = "Select a1706 From acc190, acc170 Where a1902=a1702 And a1901='" & ADO24m0.Fields("a1901") & "' Group By a1706 Order By 1 "
        Else
           StrSqlB = "select null from dual"
        End If
        intI = 1
        Set rsB = ClsLawReadRstMsg(intI, StrSqlB)
        If intI = 1 Then
           intI = 0
           Do While Not rsB.EOF
              intI = intI + 1
              If intI > 12 Then Exit Do
              strExc(9) = strExc(9) & "," & rsB.Fields(0).Value
              strExc(9) = strExc(9) & IIf(ii = 12 And rsB.RecordCount > 12, " etc.", "")
              rsB.MoveNext
           Loop
        End If
        If strExc(9) <> "" Then strExc(9) = Mid(strExc(9), 2)

        '票匯和暫收款退費不加INV. (a1902=a1702為O單號)
        'Modified by Lydia 2018/04/23 寰華Y53374備註欄都不要再填寫(by 郭) ; 'Modified by Lydia 2020/04/22  +天津三元Y37580; +建毅Y51566,唯源Y52404
        If InStr("Y37580,Y53374,Y51566,Y52404,", Left(Trim(ADO24m0.Fields("a1803")), 6)) = 0 Then
            If "" & ADO24m0.Fields("a1811") <> "1" And Val("" & ADO24m0.Fields("Ocnt")) = 0 Then  '非票匯和暫收款退費
              strExc(9) = "INV." & strExc(9)
            End If
        End If
        
        If strA2222 <> "" Then strExc(9) = PUB_GetSimpleName(strA2222, True, True) & " " & strExc(9) '備註前面+媒體備註
        
        strExc(9) = convForm(PUB_StrToStr(CheckFstSpec(RepEnter2zero(CheckStr(strExc(9)))), 105), 105)
        If Len(Trim(strExc(9))) > 100 Then
           If InStr(Mid(strExc(9), 99, 6), "etc.") > 0 Then
             strExc(9) = convForm(PUB_StrToStr(strExc(9), 105), 105)
           Else
             strExc(9) = convForm(PUB_StrToStr(strExc(9), 100) & " etc.", 105)
           End If
        End If
      '處理付款明細1~4
      m_H4(43) = Mid(strExc(9), 1, 35)
      m_H4(44) = Mid(strExc(9), 36, 35)
      m_H4(45) = Mid(strExc(9), 71, 35)
      m_H4(46) = convForm(" ", 31) '全球轉帳格式，可放「入帳議價編號」
      '(AML) 目前無資料 => 空白
      m_H4(47) = convForm(" ", 2) '(AML)與受款人關係
      m_H4(48) = convForm(" ", 2) '(AML)受款人國籍/受款公司註冊地
      '本金扣帳帳號 => 授權扣帳帳號1
      m_H4(49) = convForm("145100236819", 17)
      '本金扣帳幣別 => 授權扣帳幣別1
      m_H4(50) = "TWD"
      '手續費扣帳帳號
      m_H4(51) = m_H4(49)
      '手續費扣帳幣別
      m_H4(52) = m_H4(50)
      '手續費內含/外加 => 1:內含 2:外加 (H4(9)匯款R預設值為外加)
      m_H4(53) = "2"
      
      '匯率議價編號
      If strCurr <> m_H4(10) Then
         '抓取結匯日期or前一工作天的資料
         StrSqlB = "select * from acc1a0 where a1a03='" & m_H4(10) & "' and a1a01=(select max(a1a01) from acc1a0 where a1a03='" & m_H4(10) & _
                   "' and a1a01>=" & TransDate(CompWorkDay(2, DBDATE(strDBdate), 1), 1) & " and a1a01<=" & strDBdate & " and nvl(a1a12,'N') <> 'N')"
         intI = 1
         Set rsB = ClsLawReadRstMsg(intI, StrSqlB)
         If intI = 1 Then
            strA1a12 = "" & rsB.Fields("a1a12")
         Else
            strA1a12 = ""
         End If
      End If
      strCurr = m_H4(10)
      'Modified by Lydia 2025/09/23 去掉換行符號+PUB_GetSimpleName
      m_H4(54) = convForm(PUB_GetSimpleName(strA1a12), 16)
      '聯絡電話(申請人)
      m_H4(55) = convForm("" & ADO24m0.Fields("a0813"), 20)
      '(Email通知)是否通知受款人 => Y:通知 N不通知(預設值為 N)
      m_H4(56) = "N"
      '(Email通知)受款人郵件信箱
      m_H4(57) = convForm(" ", 40)
      '其他備註1~4 無資料 => 空白
      m_H4(58) = convForm(" ", 35)
      m_H4(59) = convForm(" ", 35)
      m_H4(60) = convForm(" ", 35)
      m_H4(61) = convForm(" ", 35)
      '轉帳日
      m_H4(62) = convForm(DBDATE(Replace(MaskEdBox3.Text, "/", "")), 8)
      '付款人銷帳參考資料 => 空白
      m_H4(63) = convForm(" ", 34)
      '是否為大陸進口：匯款性質不為710、711，則為空白。
      m_H4(64) = " "
      'H465  VARCHAR2(1) 拍發兩通電文
      '拍發兩通電文
      '當受款銀行手續費為OUR,採2通電文; 當交易國別為US(美國)時,不可選擇拍發兩通電文
      If InStr(strA2219, "OUR") > 0 And m_H4(12) <> "US" Then
         m_H4(65) = "Y"
      Else
         m_H4(65) = "N"
      End If

      'Modified by Lydia 2019/01/28 美金的OUR才要改Y
      'Modified by Lydia 2022/09/14 Y55766(German Patent and Trade Mark德國專利局)代理人手續費=71:OUR，須呈現"OUR"+2通電文
      If InStr(strA2219, "OUR") > 0 And (UCase(m_H4(10)) = "USD" Or "" & ADO24m0.Fields("a1803") = "Y55766000") Then
         m_H4(42) = "OUR" 'Added by Lydia 2019/03/08 只有美金並且手續費為OUR時, 手續費才能從預設的SHA改為OUR
         m_H4(65) = "Y"
      End If
      '改變手續費代碼，原本為SHA, BEN, OUR
      Select Case m_H4(42)
         Case "SHA" 'SHAR: 各自負擔 (SHA)
            m_H4(42) = "SHAR"
         Case "BEN" 'CRED: 由受益人負擔 (BEN)
            m_H4(42) = "CRED"
         Case "OUR" 'DEBT: 由匯款人負擔 (OUR)
            m_H4(42) = "DEBT"
      End Select

      bolAdd = True '預設增加序號
      '獨立水單不合併明細
      If ADO24m0.Fields("a1812") = "Y" Then bolUnit = True
         '合併明細資料
         StrSqlB = "Select * From accrpt24m0_H114 where UNO='" & strUserNum & "' and ar00='" & ADO24m0.Fields("a1803") & "' " & _
                   "And ar01='" & ADO24m0.Fields("a1903") & "' And ar03='" & ADO24m0.Fields("flag1") & "' and ar04='" & strA0K11 & "' and ar05 is null "
         strExc(1) = "":          intI = 1
         Erase tmpStr
         Set rsB = ClsLawReadRstMsg(intI, StrSqlB)
         If intI = 1 And bolUnit = False Then
            '串D/B note
            If Trim(rsB.Fields("H443") & rsB.Fields("H444") & rsB.Fields("H445")) <> Trim(m_H4(43) & m_H4(44) & m_H4(45)) Then
               strExc(2) = convForm(Trim(rsB.Fields("H443") & rsB.Fields("H444") & rsB.Fields("H445")) & "," & Trim(m_H4(43) & m_H4(44) & m_H4(45)), 105)
               If Len(Trim(strExc(2))) > 105 Then
                  If InStr(Mid(strExc(2), 99, 6), "etc.") > 0 Then
                  strExc(2) = convForm(PUB_StrToStr(strExc(2), 105), 105)
                Else
                  strExc(2) = convForm(PUB_StrToStr(strExc(2), 100) & " etc.", 105)
                  End If
               End If
               strExc(0) = CheckFstSpec(UCase(PUB_GetSimpleName(strExc(2), True))) 'Swift文數字規定:大寫英文
               atJ = 1
               Do While atJ < 4
                  If GetTextLength(strExc(0)) = 0 Then
                      tmpStr(atJ) = String(35, " ")
                  ElseIf GetTextLength(strExc(0)) < 35 Then
                      tmpStr(atJ) = convForm(strExc(0), 35)
                      strExc(0) = ""
                  Else
                      tmpStr(atJ) = Mid(strExc(0), 1, 35)
                      strExc(0) = CheckFstSpec(Mid(strExc(0), 36))
                  End If
                  atJ = atJ + 1
               Loop
               strExc(1) = ",H443=" & CNULL(tmpStr(1)) & ",H444=" & CNULL(tmpStr(2)) & ",H445=" & CNULL(tmpStr(3))
            End If
            
            bolAdd = False
            strExc(1) = "update accrpt24m0_H114 set H408=H408+" & Val(m_H4(8)) & strExc(1)
            strExc(1) = strExc(1) & " where UNO='" & strUserNum & "' and ar00='" & ADO24m0.Fields("a1803") & "' And ar01='" & ADO24m0.Fields("a1903") & "' " & _
                       "And ar03='" & ADO24m0.Fields("flag1") & "' and ar04='" & strA0K11 & "' "
         Else
            'PK ~ UNO(使用者),AR00(代理人編號)+AR01(幣別)+AR02(代理人名稱)+AR03(flag1)+AR04(A1917公司別)+AR05獨立水單=a1812+AR06=名稱+地址合併列印
            strExc(1) = "INSERT INTO accrpt24m0_H114 values(" & CNULL(strUserNum) & "," & CNULL(ChgSQL("" & ADO24m0.Fields("a1803"))) & "," & CNULL(ChgSQL("" & ADO24m0.Fields("a1903"))) & _
                        "," & CNULL(convForm(PUB_StrToStr(ChgSQL("" & ADO24m0.Fields("fa05")), 30), 30)) & "," & CNULL(ADO24m0.Fields("flag1")) & "," & CNULL(strA0K11) & "," & CNULL("" & ADO24m0.Fields("a1812")) & "," & CNULL(PUB_StrToStr(ChgSQL(strPrtNameAddr), 200))
            
            For intI = 0 To m_H4Num
                'H411 金額
                If intI = 11 Then
                   strExc(1) = strExc(1) & "," & Val(m_H4(intI))
                Else
                   strExc(1) = strExc(1) & "," & CNULL(ChgSQL(m_H4(intI)))
                End If
            Next intI
            strExc(1) = strExc(1) & ")"
            ii = ii + 1
         End If
         
         adoTaie.Execute strExc(1)
         
      '判斷合併水單不跳號
         If bolAdd = True Then
        inR = inR + 1
      End If
      ADO24m0.MoveNext
   Loop
   
   ADO24m0.Close
   If rsB.State <> adStateClosed Then rsB.Close
   Set rsB = Nothing
End Sub

'Added by Lydia 2025/06/26 讀取明細暫存
Private Sub PStoreRead_H114(Optional ByVal aKind As String)
Dim ff As Integer, strPath As String
Dim strR As String
Dim rsR As New ADODB.Recordset
Dim strFileNo As String, strFileName As String '檔名
Dim bolPrint As Boolean '開始列印
Dim tmpGrp As String, tmpSubTot As Double, TmpCnt As Integer, tmpI As Integer
Dim m2FileName As String  '整批匯出匯款申請書(檔名)
Dim strName As String, strText As String
Dim intJ As Integer
Dim strTel As String

    If rsR.State <> adStateClosed Then rsR.Close
    Set rsR = Nothing
    strR = "Select a.*,a0802,a0813 From accrpt24m0_H114 a ,acc080 where UNO='" & strUserNum & "' and ar04=a0801(+) order by H400 "
    
    mCompany = "": mAcctNo = ""
    intI = 1
    Set rsR = ClsLawReadRstMsg(intI, strR)
    If intI = 1 Then
        strPath = PUB_Getdesktop
        strPath = strPath & "\華銀結匯水單"
        If Dir(strPath, vbDirectory) = "" Then
           MkDir strPath
        End If
        
StartPrt:
        strFileNo = "": plStr(0) = ""
        mAppDesc = ""
        If bolPrint Then
            rsR.MoveFirst
           '設定印表機
            Printer.EndDoc
            Printer.Orientation = 2 '1.直印 2.橫印
            Printer.PaperSize = 9  'A4
               
            lngPageHeight = Printer.ScaleHeight
            lngPageWidth = Printer.ScaleWidth
            lngLineHeight = 270
            GetPleft_H106 '設定邊界
            SetColumnName_H106
            iPage = 0
        Else
           idx = 0
        End If
    
        Do While rsR.EOF = False
            '華銀不分檔
            If strFileNo = "" Then
                plStr(0) = rsR.Fields("H400")
                strFileNo = strDBdate
                If bolPrint Then
                    iPage = iPage + 1
                    mCount = 0: mTotal = 0
                    PrintHeader_H106
                Else
                    idx = idx + 1
                    If ff > 0 Then Close #ff
                    ff = FreeFile
                    strFileName = strPath & "\" & strFileNo & ".txt"
                    
                    Open strFileName For Output As ff
                End If
            End If
            
            '匯款人
            mCompany = Trim("" & rsR.Fields("a0802"))
            mAcctNo = Trim("" & rsR.Fields("H449"))
            strTel = Trim("" & rsR.Fields("a0813"))
            
            m_H4(0) = Trim("" & rsR.Fields("H400")) '列印編號,不輸出
            strExc(1) = ""
            For intI = 1 To m_H4Num
               If intI = 11 Then
                  '匯款金額15字元,轉媒體檔若長度<15,在左方補0
                  If bolPrint = False Then
                     '小數位2位預設補0,去掉小數點
                     m_H4(intI) = Right(String(15, "0") & Replace(Format(rsR.Fields("H411"), "###0.00"), ".", ""), 15)
                  Else
                     m_H4(intI) = PUB_StrToStr(Format(rsR.Fields("H411"), "###0.00"), 15, True, True)
                  End If
               Else
                  m_H4(intI) = "" & rsR.Fields(intI + 8) 'UserID+0~6屬PK,8=H400編號
               End If
               '只列印不匯出TXT：受款銀行地址、中間銀行地址；6/27 華銀需要「受款人地址+城市+國別」，拿掉22,23,
               If InStr("30,31,38,39", Format(intI, "00")) > 0 Then
                   strExc(1) = strExc(1) & convForm(" ", Len(m_H4(intI)))
               Else
                   strExc(1) = strExc(1) & m_H4(intI)
               End If
            Next intI

            '計算:華銀-整批匯出匯款申請書的匯款筆數/匯款幣別及匯款金額
            If tmpGrp <> "" & rsR.Fields("ar01") Then
               If tmpGrp <> "" Then
                  mAppDesc = mAppDesc & tmpI & "/" & tmpGrp & Format(tmpSubTot, "##,##0.00") & ";"
               End If
               tmpGrp = rsR.Fields("ar01")
               tmpI = 1
               tmpSubTot = Val(rsR.Fields("H411"))
            Else
               tmpI = tmpI + 1
               tmpSubTot = tmpSubTot + Val(rsR.Fields("H411"))
            End If
            TmpCnt = TmpCnt + 1
            
            If bolPrint Then
                '編號
                plStr(0) = m_H4(0)
                '匯款人證號(統編)
                plStr(1) = m_H4(1)
                '交易國別
                plStr(2) = m_H4(12)
                '收款行Swift 或名稱
                plStr(3) = PUB_StrToStr(IIf(Trim(m_H4(27)) <> "", m_H4(27), Trim(m_H4(28) & " " & m_H4(29))), 80)
                '收款人帳號
                plStr(4) = m_H4(17)
                '幣別
                plStr(5) = m_H4(10)
                '匯款金額
                plStr(6) = PUB_StrToStr(Format(Val(m_H4(11)), "##,##0.00"), 15, True, True)
                '手續費
                plStr(7) = m_H4(42)
                '匯款性質
                plStr(8) = m_H4(14)
                '台一編號
                plStr(9) = "" & rsR.Fields("ar00")
                plStr(10) = Trim("" & rsR.Fields("ar06"))
                
                '+第3行 中間銀行資料
                pStrL3 = ""
                If Len(Trim(m_H4(35) & m_H4(36) & m_H4(37) & m_H4(38) & m_H4(39))) > 0 Then
                   pStrL3 = "中間銀行："
                   If Trim(m_H4(35)) <> "" Then
                      pStrL3 = pStrL3 & "swift " & m_H4(35)
                   End If
                   pStrL3 = pStrL3 & " " & Trim(m_H4(36) & " " & m_H4(37) & " " & m_H4(38) & " " & m_H4(39))
                End If
                mCount = mCount + 1
                mTotal = mTotal + Val(m_H4(11))
                PrintDetail_H106
            Else
                Print #ff, strExc(1)
            End If
            
           rsR.MoveNext
        Loop
           
           
        If bolPrint Then
           mAppDesc = mAppDesc & tmpI & "/" & tmpGrp & Format(tmpSubTot, "##,##0.00") & ";"
           tmpGrp = ""
           GoTo EndPrt
        End If
        
        Close #ff
        
        '華銀-整批匯出匯款申請書的匯款筆數/匯款幣別及匯款金額
        '---------------------------
        mAppDesc = mAppDesc & tmpI & "/" & tmpGrp & Format(tmpSubTot, "##,##0.00") & ";"
        m2FileName = strPath & "\" & strDBdate & "_" & "整批匯出匯款申請書.doc"
        If Dir(m2FileName) <> "" Then
           Kill m2FileName
        End If
        tmpGrp = ""
        
        '判斷word是否已開啟
        If g_WordAp Is Nothing Then
           Set g_WordAp = New Word.Application
           g_WordAp.Visible = False
        End If
        
        tmpArr = Empty
        tmpArr = Split(mAppDesc, ";")
        If UBound(tmpArr) >= 10 Then
           If Trim(tmpArr(10)) <> "" Then
              MsgBox "幣別項目超過9項,請通知電腦中心!!"
           End If
        End If
        
        intJ = 0
        idx = idx + 1
        g_WordAp.Documents.Open App.path & "\" & m_FileName
        g_WordAp.ActiveDocument.SaveAs m2FileName
        g_WordAp.ActiveDocument.Close
        g_WordAp.Documents.Open m2FileName
        With g_WordAp
           .Selection.WholeStory
           .Selection.Copy
           For intI = 0 To 26
              strName = ""
              strText = ""
              If intI = 0 Then
                 strName = "YY"
                 strText = Mid(strDBdate, 1, 3)
              ElseIf intI = 1 Then
                 strName = "MM"
                 strText = Mid(strDBdate, 4, 2)
              ElseIf intI = 2 Then
                 strName = "DD"
                 strText = Mid(strDBdate, 6, 2)
              ElseIf intI = 3 Then
                 strName = "中文公司名"
                 strText = mCompany
              ElseIf intI = 4 Then
                 strName = "英文公司名"
                 strText = Trim(m_H4(2) & m_H4(3))
              ElseIf intI = 5 Then
                 strName = "英文地址"
                 strText = Trim(m_H4(4)) & " " & Trim(m_H4(5))
              ElseIf intI = 6 Then
                 strName = "公司電話"
                 strText = strTel
              ElseIf intI = 7 Then
                 strName = "公司統編"
                 strText = Trim(m_H4(1))
             '---匯款幣別和匯款金額 ------------
              ElseIf intI >= 8 And intI <= 25 Then
                 If intI Mod 2 = 0 Then
                    Select Case intI
                        Case 8:  strName = "D10"
                        Case 10: strName = "D20"
                        Case 12: strName = "D30"
                        Case 14: strName = "D40"
                        Case 16: strName = "D50"
                        Case 18: strName = "D60"
                        Case 20: strName = "D70"
                        Case 22: strName = "D80"
                        Case 24: strName = "D90"
                    End Select
                    If intJ <= UBound(tmpArr) Then
                       If Trim(tmpArr(intJ)) <> "" Then
                          strText = Mid(Trim(tmpArr(intJ)), 1, InStr(Trim(tmpArr(intJ)), "/") - 1)
                       Else
                          strText = ""
                       End If
                    Else
                       strText = ""
                    End If
                 Else
                    Select Case intI
                        Case 9:  strName = "D11"
                        Case 11: strName = "D21"
                        Case 13: strName = "D31"
                        Case 15: strName = "D41"
                        Case 17: strName = "D51"
                        Case 19: strName = "D61"
                        Case 21: strName = "D71"
                        Case 23: strName = "D81"
                        Case 25: strName = "D91"
                    End Select
                    If intJ <= UBound(tmpArr) Then
                       If Trim(tmpArr(intJ)) <> "" Then
                          strText = Mid(Trim(tmpArr(intJ)), InStr(Trim(tmpArr(intJ)), "/") + 1)
                          intJ = intJ + 1
                       Else
                          strText = ""
                       End If
                    Else
                       strText = ""
                    End If
                 End If
              ElseIf intI = 26 Then
                 strName = "TOT1"
                 strText = Trim(TmpCnt)
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
                 .Selection.TypeText strText
              End If
           Next intI
        End With
          
        ''Modified by Lydia 2024/01/31 改存成PDF
        'g_WordAp.ActiveDocument.Save
        'g_WordAp.ActiveDocument.Close
        If PUB_PrintWord2File(g_WordAp, strPath, strDBdate & "_" & "整批匯出匯款申請書") = False Then
           Exit Sub
        Else
           PUB_DelPCOrgFile m2FileName '刪除原本Word檔
        End If
        'end 2024/01/31
        
        Clipboard.Clear '清除剪貼簿動作
        '---------------------------
        
        If TxtList <> "Y" Then
           rsR.Close
           MsgBox "已在桌面的華銀結匯水單資料夾產生 " & idx & " 個檔案!!", vbInformation
        End If
        
        If TxtList = "Y" And bolPrint = False Then
           bolPrint = True
           GoTo StartPrt
        End If
    End If
'------------------------
EndPrt:
   If bolPrint Then
        PrintSubTotal_H106
        PrintSign
        Printer.EndDoc
        MsgBox "已在桌面的華銀結匯水單資料夾產生 " & idx & " 個檔案, 並列印清單完成!!", vbInformation
        rsR.Close
   End If
End Sub





